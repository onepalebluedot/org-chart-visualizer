import { Position, type Edge, type Node, type XYPosition } from "@xyflow/react";
import type { LayoutNode, OrgData, PersonRecord, RoleType, ViewMode } from "../types";

interface SavedOrgFile {
  version: 1;
  savedAt: string;
  data: OrgData;
}

export interface SpreadsheetImportPreview {
  sheetName: string;
  rowCount: number;
  importedCount: number;
  rootName: string;
  missingColumns: string[];
  duplicateNames: string[];
  unresolvedManagers: string[];
  warnings: string[];
}

export interface SpreadsheetImportResult {
  data: OrgData;
  preview: SpreadsheetImportPreview;
}

interface ImportedRosterPerson extends PersonRecord {
  sourceManagerName: string;
}

export interface OrgNodeData extends Record<string, unknown> {
  person: PersonRecord;
  selected: boolean;
  collapsed: boolean;
  isDropTarget: boolean;
  isInvalidTarget: boolean;
  isFocusGroup: boolean;
  directReportCount: number;
  viewMode: ViewMode;
  lightMode: boolean;
}

export type AppNode = Node<OrgNodeData>;

const FOCUSED_STANDARD_LAYOUT = {
  nodeWidth: 260,
  nodeHeight: 104,
  columnGap: 72,
  rowGap: 72,
  leafGap: 16,
  leafOffset: 14,
  branchColumnGap: 58,
  branchVerticalGap: 58,
  marginX: 118,
  marginY: 28
};

const FOCUSED_LIGHT_LAYOUT = {
  nodeWidth: 168,
  nodeHeight: 42,
  columnGap: 14,
  rowGap: 30,
  leafGap: 10,
  leafOffset: 8,
  branchColumnGap: 28,
  branchVerticalGap: 28,
  marginX: 92,
  marginY: 20
};

const LOCATION_STANDARD_LAYOUT = {
  columnWidth: 278,
  headerY: 24,
  memberStartY: 118,
  memberGap: 96
};

const LOCATION_LIGHT_LAYOUT = {
  columnWidth: 176,
  headerY: 20,
  memberStartY: 82,
  memberGap: 52
};

const REQUIRED_ROSTER_COLUMNS = [
  "Name",
  "Role",
  "Manager Or IC",
  "Full Time or Contractor",
  "Title",
  "Manager",
  "Level",
  "Location"
] as const;

const normalizeHeader = (value: string): string => value.trim().toLowerCase().replace(/[^a-z0-9]+/g, "");

const rosterHeaderAliases: Record<(typeof REQUIRED_ROSTER_COLUMNS)[number], string[]> = {
  Name: ["name"],
  Role: ["role"],
  "Manager Or IC": ["manageroric", "manageric", "managerorindividualcontributor"],
  "Full Time or Contractor": ["fulltimeorcontractor", "fulltimecontractor", "workertype", "employmenttype"],
  Title: ["title"],
  Manager: ["manager"],
  Level: ["level"],
  Location: ["location"]
};

const slugify = (value: string): string =>
  value
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "") || "person";

const coerceSpreadsheetValue = (value: unknown): string => {
  if (value === null || value === undefined) return "";
  return String(value).trim();
};

const resolveRosterValue = (
  row: Record<string, unknown>,
  keyMap: Record<(typeof REQUIRED_ROSTER_COLUMNS)[number], string | null>,
  column: (typeof REQUIRED_ROSTER_COLUMNS)[number]
): string => {
  const key = keyMap[column];
  if (!key) return "";
  return coerceSpreadsheetValue(row[key]);
};

export const buildOrgDataFromRosterRows = (
  rows: Record<string, unknown>[],
  sheetName = "Sheet1"
): SpreadsheetImportResult => {
  if (rows.length === 0) {
    throw new Error("The spreadsheet is empty.");
  }

  const firstRow = rows.find((row) => Object.keys(row).length > 0);
  if (!firstRow) {
    throw new Error("The spreadsheet does not contain any readable rows.");
  }

  const rowKeys = Object.keys(firstRow);
  const normalizedKeyLookup = new Map(rowKeys.map((key) => [normalizeHeader(key), key]));
  const keyMap = Object.fromEntries(
    REQUIRED_ROSTER_COLUMNS.map((column) => [
      column,
      rosterHeaderAliases[column].map((alias) => normalizedKeyLookup.get(alias) ?? null).find(Boolean) ?? null
    ])
  ) as Record<(typeof REQUIRED_ROSTER_COLUMNS)[number], string | null>;

  const missingColumns = REQUIRED_ROSTER_COLUMNS.filter((column) => !keyMap[column]);
  if (missingColumns.length > 0) {
    throw new Error(`Missing required columns: ${missingColumns.join(", ")}`);
  }

  const warnings: string[] = [];
  let skippedRows = 0;
  const idCounts = new Map<string, number>();
  const nameCounts = new Map<string, number>();
  const importedRows = rows.flatMap<ImportedRosterPerson>((row, index) => {
    const name = resolveRosterValue(row, keyMap, "Name");
    const role = resolveRosterValue(row, keyMap, "Role");
    const managerOrIcRaw = resolveRosterValue(row, keyMap, "Manager Or IC");
    const workerType = resolveRosterValue(row, keyMap, "Full Time or Contractor");
    const title = resolveRosterValue(row, keyMap, "Title");
    const managerName = resolveRosterValue(row, keyMap, "Manager");
    const levelValue = resolveRosterValue(row, keyMap, "Level");
    const location = resolveRosterValue(row, keyMap, "Location");

    if (![name, role, managerOrIcRaw, workerType, title, managerName, levelValue, location].some(Boolean)) {
      return [];
    }

    if (!name) {
      skippedRows += 1;
      return [];
    }

    const slug = slugify(name);
    const count = (idCounts.get(slug) ?? 0) + 1;
    idCounts.set(slug, count);
    nameCounts.set(name, (nameCounts.get(name) ?? 0) + 1);

    const managerOrIc: PersonRecord["managerOrIc"] = managerOrIcRaw.toLowerCase().startsWith("manager")
      ? "Manager"
      : "IC";
    const roleType =
      name.toLowerCase() === "open role" ? "open-role" : managerOrIc === "Manager" ? "manager" : "ic";
    const parsedLevel = Number(levelValue);

    return [
      {
        id: count === 1 ? slug : `${slug}-${count}`,
        parentId: null as string | null,
        sortOrder: index,
        name,
        role,
        managerOrIc,
        workerType: workerType || "Full Time",
        title: title || (managerOrIc === "Manager" ? "Manager" : "Individual contributor"),
        managerName,
        level: Number.isFinite(parsedLevel) ? parsedLevel : 0,
        location: location || "Unassigned",
        roleType: roleType as PersonRecord["roleType"],
        sourceManagerName: managerName
      }
    ];
  });

  if (importedRows.length === 0) {
    throw new Error("No usable roster rows were found in the spreadsheet.");
  }

  const duplicateNames = [...nameCounts.entries()]
    .filter(([, count]) => count > 1)
    .map(([name, count]) => `${name} (${count})`);
  const nameToId = new Map<string, string>();
  importedRows.forEach((person) => {
    if (!nameToId.has(person.name)) {
      nameToId.set(person.name, person.id);
    }
  });

  const unresolvedManagers = [
    ...new Set(
      importedRows
        .map((person) => person.sourceManagerName)
        .filter((managerName) => managerName && !nameToId.has(managerName))
    )
  ].sort((a, b) => a.localeCompare(b));

  const rootCandidates = importedRows.filter(
    (person) =>
      !person.sourceManagerName || !nameToId.has(person.sourceManagerName) || person.sourceManagerName === person.name
  );
  const rankedRootCandidates = [...rootCandidates].sort((a, b) => {
    if (b.level !== a.level) return b.level - a.level;
    if (a.managerOrIc !== b.managerOrIc) return a.managerOrIc === "Manager" ? -1 : 1;
    return (a.sortOrder ?? 0) - (b.sortOrder ?? 0);
  });
  const root = rankedRootCandidates[0] ?? importedRows[0];

  if (rootCandidates.length > 1) {
    warnings.push(`Multiple top-level candidates found. Using ${root.name} as the root.`);
  }
  if (unresolvedManagers.length > 0) {
    warnings.push(`Some manager names were not found and were attached to the root.`);
  }
  if (duplicateNames.length > 0) {
    warnings.push("Duplicate names were detected. Unique IDs were generated automatically.");
  }
  if (skippedRows > 0) {
    warnings.push(`${skippedRows} blank or unnamed row${skippedRows === 1 ? "" : "s"} were skipped.`);
  }

  const people: PersonRecord[] = importedRows.map((person) => {
    const parentId =
      person.id === root.id
        ? null
        : person.sourceManagerName && nameToId.has(person.sourceManagerName)
          ? nameToId.get(person.sourceManagerName)!
          : root.id;

    return {
      id: person.id,
      parentId,
      sortOrder: person.sortOrder,
      name: person.name,
      role: person.role || (person.managerOrIc === "Manager" ? "Leadership" : "General"),
      managerOrIc: person.id === root.id ? "Manager" : person.managerOrIc,
      workerType: person.workerType,
      title: person.title,
      managerName: person.id === root.id ? "" : person.managerName,
      level: person.level,
      location: person.location,
      roleType: person.id === root.id ? "executive" : person.roleType
    };
  });

  const data = normalizeOrgData({
    rootId: root.id,
    people
  });

  return {
    data,
    preview: {
      sheetName,
      rowCount: rows.length,
      importedCount: data.people.length,
      rootName: data.people.find((person) => person.id === data.rootId)?.name ?? root.name,
      missingColumns,
      duplicateNames,
      unresolvedManagers,
      warnings
    }
  };
};

export const createEmptyPerson = (parentId: string | null, roleType: RoleType): PersonRecord => {
  return {
    id: `person-${Math.random().toString(36).slice(2, 10)}`,
    parentId,
    sortOrder: Date.now(),
    name: roleType === "open-role" ? "Open Role" : "New Person",
    role: roleType === "manager" ? "Leadership" : "Automation",
    managerOrIc: roleType === "manager" || roleType === "executive" ? "Manager" : "IC",
    workerType: "Full Time",
    title: roleType === "manager" ? "Manager" : "Deployment Engineer",
    managerName: "",
    level: roleType === "manager" ? 8 : 6,
    location: "Remote",
    roleType
  };
};

export const peopleById = (people: PersonRecord[]): Record<string, PersonRecord> =>
  Object.fromEntries(people.map((person) => [person.id, person]));

export const childrenByParent = (people: PersonRecord[]): Record<string, PersonRecord[]> => {
  const map: Record<string, PersonRecord[]> = {};

  for (const person of people) {
    const key = person.parentId ?? "__root__";
    map[key] ??= [];
    map[key].push(person);
  }

  Object.values(map).forEach((siblings) =>
    siblings.sort((a, b) => {
      if ((a.sortOrder ?? 0) !== (b.sortOrder ?? 0)) return (a.sortOrder ?? 0) - (b.sortOrder ?? 0);
      if (a.roleType === "open-role" && b.roleType !== "open-role") return 1;
      if (a.roleType !== "open-role" && b.roleType === "open-role") return -1;
      return a.name.localeCompare(b.name);
    })
  );

  return map;
};

const sortReports = (reports: PersonRecord[]): PersonRecord[] =>
  [...reports].sort((a, b) => {
    const rank = (person: PersonRecord) => {
      if (person.roleType === "manager") return 0;
      if (person.roleType === "ic") return 1;
      if (person.roleType === "open-role") return 2;
      return 3;
    };

    return rank(a) - rank(b) || (a.sortOrder ?? 0) - (b.sortOrder ?? 0) || a.name.localeCompare(b.name);
  });

const getFocusParts = (data: OrgData, focusId: string | null) => {
  const byId = peopleById(data.people);
  const childrenMap = childrenByParent(data.people);
  const root = byId[data.rootId] ?? data.people[0];
  const focus = (focusId ? byId[focusId] : null) ?? root;
  const ancestors: PersonRecord[] = [];
  const visited = new Set<string>([focus.id]);
  let current = focus;

  while (current.parentId) {
    const parent = byId[current.parentId];
    if (!parent || visited.has(parent.id)) break;
    ancestors.unshift(parent);
    visited.add(parent.id);
    current = parent;
  }

  const peers =
    focus.parentId && childrenMap[focus.parentId]?.length
      ? sortReports(childrenMap[focus.parentId])
      : [focus];

  return {
    childrenMap,
    focus,
    ancestors,
    peers
  };
};

export const getFocusedExpandableIds = (data: OrgData, focusId: string | null): Set<string> => {
  const { childrenMap, ancestors, peers } = getFocusParts(data, focusId);
  const expandableIds = new Set<string>();
  const visited = new Set<string>();

  const collect = (people: PersonRecord[]) => {
    for (const person of people) {
      if (visited.has(person.id)) continue;
      visited.add(person.id);

      const reports = sortReports(childrenMap[person.id] ?? []);
      if (reports.length === 0) continue;

      expandableIds.add(person.id);
      collect(reports);
    }
  };

  collect([...ancestors, ...peers]);
  return expandableIds;
};

export const getDefaultFocusedExpandedIds = (data: OrgData, focusId: string | null): Set<string> => {
  const { childrenMap, focus } = getFocusParts(data, focusId);
  return (childrenMap[focus.id] ?? []).length > 0 ? new Set([focus.id]) : new Set();
};

const buildFocusedLaneOrgLayout = (data: OrgData, expandedIds: Set<string>, focusId: string | null, lightMode: boolean): LayoutNode[] => {
  const layout = lightMode ? FOCUSED_LIGHT_LAYOUT : FOCUSED_STANDARD_LAYOUT;
  const byId = peopleById(data.people);
  const { childrenMap, focus, ancestors, peers } = getFocusParts(data, focusId);
  const visibleIds = new Set<string>();
  const positions = new Map<string, LayoutNode>();
  const leafStep = layout.nodeHeight + layout.leafGap;
  const blockGap = layout.columnGap;

  const markVisible = (person: PersonRecord, rowIndex: number) => {
    if (!byId[person.id]) return;
    visibleIds.add(person.id);
  };

  const peerRowFor = (person: PersonRecord): PersonRecord[] =>
    person.parentId && (childrenMap[person.parentId] ?? []).length > 0
      ? sortReports(childrenMap[person.parentId])
      : [person];

  const usesCompactBranch = (reports: PersonRecord[]): boolean =>
    reports.length > 0 && !reports.some((report) => expandedIds.has(report.id));

  const contextRows = [...ancestors.map(peerRowFor), peers.length > 0 ? peers : [focus]];
  contextRows.forEach((row, rowIndex) => row.forEach((person) => markVisible(person, rowIndex)));

  const blockWidth = (person: PersonRecord, visited = new Set<string>()): number => {
    if (visited.has(person.id) || !expandedIds.has(person.id)) return layout.nodeWidth;
    visited.add(person.id);

    const reports = sortReports(childrenMap[person.id] ?? []);
    if (reports.length === 0) return layout.nodeWidth;

    if (usesCompactBranch(reports)) return Math.max(layout.nodeWidth, layout.nodeWidth * 2 + layout.branchColumnGap);

    const childrenWidth =
      reports.reduce((total, report) => total + blockWidth(report, new Set(visited)), 0) +
      Math.max(0, reports.length - 1) * blockGap;
    return Math.max(layout.nodeWidth, childrenWidth);
  };

  const contextRowWidth = (row: PersonRecord[]): number =>
    row.reduce((total, person) => total + blockWidth(person), 0) + Math.max(0, row.length - 1) * blockGap;

  const maxRowWidth = Math.max(0, ...contextRows.map(contextRowWidth));
  const centerX = layout.marginX + maxRowWidth / 2;
  const rowStep = layout.nodeHeight + layout.rowGap;

  const placeSubtree = (person: PersonRecord, x: number, rowIndex: number, visited = new Set<string>()) => {
    if (visited.has(person.id)) return;
    visited.add(person.id);
    markVisible(person, rowIndex);

    const width = blockWidth(person);
    positions.set(person.id, {
      id: person.id,
      x: x + (width - layout.nodeWidth) / 2,
      y: layout.marginY + rowIndex * rowStep
    });

    if (!expandedIds.has(person.id)) return;

    const reports = sortReports(childrenMap[person.id] ?? []);
    if (reports.length === 0) return;

    const parentPosition = positions.get(person.id);
    if (usesCompactBranch(reports) && parentPosition) {
      const parentCenterX = parentPosition.x + layout.nodeWidth / 2;

      reports.forEach((report, index) => {
        const branchIndex = index % 2;
        const leafRowIndex = Math.floor(index / 2);
        const childX =
          branchIndex === 0
            ? parentCenterX - layout.branchColumnGap / 2 - layout.nodeWidth
            : parentCenterX + layout.branchColumnGap / 2;

        markVisible(report, rowIndex + 1 + leafRowIndex);
        positions.set(report.id, {
          id: report.id,
          x: childX,
          y: parentPosition.y + layout.nodeHeight + layout.branchVerticalGap + leafRowIndex * leafStep
        });
      });
      return;
    }

    const childrenWidth =
      reports.reduce((total, report) => total + blockWidth(report, new Set(visited)), 0) +
      Math.max(0, reports.length - 1) * blockGap;
    let childX = x + (width - childrenWidth) / 2;

    reports.forEach((report) => {
      const childWidth = blockWidth(report, new Set(visited));
      placeSubtree(report, childX, rowIndex + 1, new Set(visited));
      childX += childWidth + blockGap;
    });
  };

  contextRows.forEach((row, rowIndex) => {
    let cursorX = centerX - contextRowWidth(row) / 2;
    row.forEach((person) => {
      const width = blockWidth(person);
      placeSubtree(person, cursorX, rowIndex);
      cursorX += width + blockGap;
    });
  });

  return [...visibleIds].flatMap((personId) => {
    const position = positions.get(personId);
    return position ? [position] : [];
  });
};

export const buildOrgLayout = (data: OrgData, expandedIds: Set<string>, lightMode: boolean, focusId: string | null): LayoutNode[] =>
  buildFocusedLaneOrgLayout(data, expandedIds, focusId, lightMode);

export const buildLocationLayout = (people: PersonRecord[], lightMode: boolean): LayoutNode[] => {
  const layout = lightMode ? LOCATION_LIGHT_LAYOUT : LOCATION_STANDARD_LAYOUT;
  const groups = new Map<string, PersonRecord[]>();

  for (const person of people) {
    const location = person.location || "Unassigned";
    groups.set(location, [...(groups.get(location) ?? []), person]);
  }

  const entries = [...groups.entries()].sort(([a], [b]) => a.localeCompare(b));
  const positions: LayoutNode[] = [];

  entries.forEach(([location, members], columnIndex) => {
    const headerId = `location:${location}`;
    positions.push({
      id: headerId,
      x: columnIndex * layout.columnWidth + 120,
      y: layout.headerY
    });

    members
      .sort((a, b) => a.name.localeCompare(b.name))
      .forEach((member, rowIndex) => {
        const compoundId = `${headerId}:${member.id}`;
        positions.push({
          id: compoundId,
          x: columnIndex * layout.columnWidth + 120,
          y: layout.memberStartY + rowIndex * layout.memberGap
        });
      });
  });

  return positions;
};

export const isDescendant = (people: PersonRecord[], ancestorId: string, possibleDescendantId: string): boolean => {
  const byId = peopleById(people);
  let current = byId[possibleDescendantId];

  while (current?.parentId) {
    if (current.parentId === ancestorId) {
      return true;
    }
    current = byId[current.parentId];
  }

  return false;
};

export const reparentPerson = (people: PersonRecord[], personId: string, newParentId: string): PersonRecord[] =>
  people.map((person) => (person.id === personId ? { ...person, parentId: newParentId } : person));

export const reorderSiblings = (
  people: PersonRecord[],
  movedId: string,
  targetId: string,
  placement: "before" | "after"
): PersonRecord[] => {
  const byId = peopleById(people);
  const moved = byId[movedId];
  const target = byId[targetId];

  if (!moved || !target || moved.parentId !== target.parentId || moved.id === target.id) {
    return people;
  }

  const siblings = people
    .filter((person) => person.parentId === moved.parentId)
    .sort((a, b) => (a.sortOrder ?? 0) - (b.sortOrder ?? 0) || a.name.localeCompare(b.name));
  const remaining = siblings.filter((person) => person.id !== movedId);
  const targetIndex = remaining.findIndex((person) => person.id === targetId);

  if (targetIndex === -1) {
    return people;
  }

  const insertIndex = placement === "before" ? targetIndex : targetIndex + 1;
  remaining.splice(insertIndex, 0, moved);
  const orderMap = new Map(remaining.map((person, index) => [person.id, index]));

  return people.map((person) =>
    orderMap.has(person.id)
      ? {
          ...person,
          sortOrder: orderMap.get(person.id)!
        }
      : person
  );
};

export const normalizeOrgData = (data: OrgData): OrgData => {
  const byId = peopleById(data.people);
  const root = byId[data.rootId] ?? data.people[0];

  if (!root) {
    return data;
  }

  const normalizedPeople: PersonRecord[] = data.people.map((person, index) => {
    if (person.id === root.id) {
      return {
        ...person,
        parentId: null,
        sortOrder: person.sortOrder ?? index,
        managerName: person.managerName || "",
        managerOrIc: "Manager" as const,
        roleType: "executive" as const
      };
    }

    let parentId = person.parentId;
    const parent = parentId ? byId[parentId] : null;

    if (!parent || parent.id === person.id) {
      parentId = root.id;
    }

    const isManagerTrack = person.managerOrIc === "Manager";
    const normalizedRoleType =
      person.name.trim().toLowerCase() === "open role"
        ? "open-role"
        : isManagerTrack
          ? "manager"
          : "ic";

    return {
      ...person,
      parentId,
      sortOrder: person.sortOrder ?? index,
      managerName: parentId ? (byId[parentId]?.name ?? person.managerName) : person.managerName,
      managerOrIc: (isManagerTrack ? "Manager" : "IC") as "Manager" | "IC",
      roleType: normalizedRoleType
    };
  });

  return {
    rootId: root.id,
    people: normalizedPeople
  };
};

export const serializeOrgData = (data: OrgData): string =>
  JSON.stringify(
    {
      version: 1,
      savedAt: new Date().toISOString(),
      data
    } satisfies SavedOrgFile,
    null,
    2
  );

export const parseOrgData = (text: string): OrgData => {
  const parsed = JSON.parse(text) as OrgData | SavedOrgFile;

  if ("data" in parsed && parsed.data && "rootId" in parsed.data && Array.isArray(parsed.data.people)) {
    return parsed.data;
  }

  if (!("rootId" in parsed) || !("people" in parsed) || !parsed.rootId || !Array.isArray(parsed.people)) {
    throw new Error("Invalid org data format.");
  }
  return parsed as OrgData;
};

export const filterPeople = (people: PersonRecord[], query: string): Set<string> => {
  if (!query.trim()) {
    return new Set(people.map((person) => person.id));
  }

  const normalized = query.trim().toLowerCase();
  return new Set(
    people
      .filter((person) => {
        const haystack = [
          person.name,
          person.role,
          person.managerOrIc,
          person.workerType,
          person.title,
          person.managerName,
          String(person.level),
          person.location,
        ]
          .join(" ")
          .toLowerCase();
        return haystack.includes(normalized);
      })
      .map((person) => person.id)
  );
};

export const buildOrgFlow = ({
  data,
  expandedIds,
  selectedId,
  focusId,
  dropTargetId,
  invalidTargetId,
  previewPositions,
  lightMode
}: {
  data: OrgData;
  expandedIds: Set<string>;
  selectedId: string | null;
  focusId: string | null;
  dropTargetId: string | null;
  invalidTargetId: string | null;
  previewPositions: Record<string, XYPosition>;
  lightMode: boolean;
}): { nodes: AppNode[]; edges: Edge[] } => {
  const byId = peopleById(data.people);
  const childrenMap = childrenByParent(data.people);
  const layout = buildOrgLayout(data, expandedIds, lightMode, focusId);
  const layoutById = new Map(layout.map((item) => [item.id, item]));
  const visibleIds = new Set(layout.map((item) => item.id));
  const nodes: AppNode[] = layout.map((layoutNode) => {
    const person = byId[layoutNode.id];
    const directReportCount = (childrenMap[person.id] ?? []).length;
    const hasReports = directReportCount > 0;

    return {
      id: person.id,
      type: "person",
      position: previewPositions[person.id] ?? { x: layoutNode.x, y: layoutNode.y },
      data: {
        person,
        selected: person.id === selectedId,
        collapsed: hasReports && !expandedIds.has(person.id),
        isDropTarget: person.id === dropTargetId,
        isInvalidTarget: person.id === invalidTargetId,
        isFocusGroup: person.id === focusId && expandedIds.has(person.id),
        directReportCount,
        viewMode: "org",
        lightMode
      },
      sourcePosition: Position.Bottom,
      targetPosition: Position.Top,
      draggable: person.id !== data.rootId
    };
  });

  const edges: Edge[] = data.people
    .filter((person) => person.parentId && visibleIds.has(person.id) && visibleIds.has(person.parentId))
    .map((person) => {
      const parentPosition = layoutById.get(person.parentId!);
      const childPosition = layoutById.get(person.id);
      const parentCenterX = parentPosition ? parentPosition.x + (lightMode ? FOCUSED_LIGHT_LAYOUT.nodeWidth : FOCUSED_STANDARD_LAYOUT.nodeWidth) / 2 : 0;
      const childCenterX = childPosition ? childPosition.x + (lightMode ? FOCUSED_LIGHT_LAYOUT.nodeWidth : FOCUSED_STANDARD_LAYOUT.nodeWidth) / 2 : 0;
      const horizontalDelta = childCenterX - parentCenterX;
      const parentReports = childrenMap[person.parentId!] ?? [];
      const isCompactBranch =
        expandedIds.has(person.parentId!) &&
        parentReports.length > 0 &&
        !parentReports.some((report) => expandedIds.has(report.id));
      const targetHandle =
        isCompactBranch && Math.abs(horizontalDelta) >= 40
          ? horizontalDelta < 0
            ? "report-target-right"
            : "report-target-left"
          : "report-target-top";

      return {
        id: `${person.parentId}-${person.id}`,
        source: person.parentId!,
        target: person.id,
        sourceHandle: "report-source-bottom",
        targetHandle,
        type: "reporting",
        animated: false,
        style: {
          stroke: person.id === focusId || person.parentId === focusId ? "#3042f5" : "#aeb3cb",
          strokeWidth: person.id === focusId || person.parentId === focusId ? 2.8 : 1.8
        }
      };
    });

  return { nodes, edges };
};

export const buildLocationFlow = ({
  data,
  selectedId,
  previewPositions,
  lightMode
}: {
  data: OrgData;
  selectedId: string | null;
  previewPositions: Record<string, XYPosition>;
  lightMode: boolean;
}): { nodes: AppNode[]; edges: Edge[] } => {
  const layout = lightMode ? LOCATION_LIGHT_LAYOUT : LOCATION_STANDARD_LAYOUT;
  const groups = new Map<string, PersonRecord[]>();
  const directReportCounts = childrenByParent(data.people);

  for (const person of data.people) {
    const location = person.location || "Unassigned";
    groups.set(location, [...(groups.get(location) ?? []), person]);
  }

  const nodes: AppNode[] = [];
  const edges: Edge[] = [];
  const locations = [...groups.entries()].sort(([a], [b]) => a.localeCompare(b));

  locations.forEach(([location, members], columnIndex) => {
    const headerId = `location:${location}`;
    nodes.push({
      id: headerId,
      type: "project",
      position: { x: columnIndex * layout.columnWidth + 120, y: layout.headerY },
      data: {
        person: {
          id: headerId,
          parentId: null,
          name: location,
          role: "Location",
          managerOrIc: "Manager",
          workerType: `${members.length} assigned`,
          title: `${members.length} assigned`,
          managerName: "",
          level: 0,
          location,
          roleType: "manager",
        },
        selected: false,
        collapsed: false,
        isDropTarget: false,
        isInvalidTarget: false,
        isFocusGroup: false,
        directReportCount: members.length,
        viewMode: "location",
        lightMode
      },
      draggable: false
    });

    members
      .sort((a, b) => a.name.localeCompare(b.name))
      .forEach((member, rowIndex) => {
        const compoundId = `${headerId}:${member.id}`;
        nodes.push({
          id: compoundId,
          type: "person",
          position:
            previewPositions[compoundId] ?? {
              x: columnIndex * layout.columnWidth + 120,
              y: layout.memberStartY + rowIndex * layout.memberGap
            },
          data: {
            person: member,
            selected: member.id === selectedId,
            collapsed: false,
            isDropTarget: false,
            isInvalidTarget: false,
            isFocusGroup: false,
            directReportCount: (directReportCounts[member.id] ?? []).length,
            viewMode: "location",
            lightMode
          },
          draggable: true
        });

        edges.push({
          id: `${headerId}-${compoundId}`,
          source: headerId,
          target: compoundId,
          type: "smoothstep"
        });
      });
  });

  return { nodes, edges };
};
