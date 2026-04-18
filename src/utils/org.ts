import type { Edge, Node, XYPosition } from "@xyflow/react";
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
  viewMode: ViewMode;
  lightMode: boolean;
}

export type AppNode = Node<OrgNodeData>;

const STANDARD_GRID = {
  columnWidth: 278,
  rowHeight: 96,
  indent: 10,
  rootY: 24,
  columnStartX: 110,
  columnStartY: 118
};

const LIGHT_GRID = {
  columnWidth: 176,
  rowHeight: 52,
  indent: 6,
  rootY: 18,
  columnStartX: 92,
  columnStartY: 84
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

export const buildOrgLayout = (data: OrgData, collapsedIds: Set<string>, lightMode: boolean): LayoutNode[] => {
  const grid = lightMode ? LIGHT_GRID : STANDARD_GRID;
  const childrenMap = childrenByParent(data.people);
  const positions: LayoutNode[] = [];
  const topLevelReports = sortReports(childrenMap[data.rootId] ?? []);
  const rootCenterX =
    topLevelReports.length > 0
      ? grid.columnStartX + ((topLevelReports.length - 1) * grid.columnWidth) / 2
      : grid.columnStartX;

  positions.push({
    id: data.rootId,
    x: rootCenterX,
    y: grid.rootY
  });

  const nextYByColumn = topLevelReports.map(() => grid.columnStartY);

  const placeInColumn = (personId: string, columnIndex: number, depth: number) => {
    positions.push({
      id: personId,
      x: grid.columnStartX + columnIndex * grid.columnWidth + depth * grid.indent,
      y: nextYByColumn[columnIndex]
    });

    nextYByColumn[columnIndex] += grid.rowHeight;

    if (collapsedIds.has(personId)) {
      return;
    }

    const reports = sortReports(childrenMap[personId] ?? []);
    for (const child of reports) {
      placeInColumn(child.id, columnIndex, depth + 1);
    }
  };

  topLevelReports.forEach((report, columnIndex) => {
    placeInColumn(report.id, columnIndex, 0);
  });

  return positions;
};

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

  const managerIds = data.people.filter((person) => person.roleType === "manager").map((person) => person.id);
  const fallbackManagerId = managerIds[0] ?? root.id;
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
      parentId = person.roleType === "manager" ? root.id : fallbackManagerId;
    }

    const resolvedParent = parentId ? byId[parentId] ?? root : root;
    const isManagerTrack = person.managerOrIc === "Manager";
    const normalizedRoleType =
      person.name.trim().toLowerCase() === "open role"
        ? "open-role"
        : isManagerTrack
          ? "manager"
          : "ic";

    if ((normalizedRoleType === "ic" || normalizedRoleType === "open-role") && resolvedParent.roleType !== "manager") {
      parentId = fallbackManagerId;
    }

    if ((resolvedParent.roleType === "ic" || resolvedParent.roleType === "open-role") && resolvedParent.parentId) {
      parentId = resolvedParent.parentId;
    }

    return {
      ...person,
      parentId,
      sortOrder: person.sortOrder ?? index,
      managerName: parentId ? (byId[parentId]?.name ?? person.managerName) : person.managerName,
      managerOrIc: (isManagerTrack ? "Manager" : "IC") as "Manager" | "IC",
      roleType: normalizedRoleType
    };
  });

  const normalizedChildren = childrenByParent(normalizedPeople);

  return {
    rootId: root.id,
    people: normalizedPeople.map((person) => {
      if ((person.roleType === "ic" || person.roleType === "open-role") && (normalizedChildren[person.id] ?? []).length > 0) {
        return {
          ...person,
          managerOrIc: "Manager" as const,
          roleType: "manager" as const
        };
      }

      return person;
    })
  };
};

export const toggleCollapseForAll = (data: OrgData, nextCollapsed: boolean): Set<string> => {
  if (!nextCollapsed) {
    return new Set<string>();
  }

  return new Set(
    data.people
      .filter((person) => person.id !== data.rootId && person.roleType !== "ic" && person.roleType !== "open-role")
      .map((person) => person.id)
  );
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
  collapsedIds,
  selectedId,
  dropTargetId,
  invalidTargetId,
  previewPositions,
  lightMode
}: {
  data: OrgData;
  collapsedIds: Set<string>;
  selectedId: string | null;
  dropTargetId: string | null;
  invalidTargetId: string | null;
  previewPositions: Record<string, XYPosition>;
  lightMode: boolean;
}): { nodes: AppNode[]; edges: Edge[] } => {
  const byId = peopleById(data.people);
  const layout = buildOrgLayout(data, collapsedIds, lightMode);
  const visibleIds = new Set(layout.map((item) => item.id));
  const nodes: AppNode[] = layout.map((layoutNode) => {
    const person = byId[layoutNode.id];
    return {
      id: person.id,
      type: "person",
      position: previewPositions[person.id] ?? { x: layoutNode.x, y: layoutNode.y },
      data: {
        person,
        selected: person.id === selectedId,
        collapsed: collapsedIds.has(person.id),
        isDropTarget: person.id === dropTargetId,
        isInvalidTarget: person.id === invalidTargetId,
        viewMode: "org",
        lightMode
      },
      draggable: person.id !== data.rootId
    };
  });

  const edges: Edge[] = data.people
    .filter((person) => person.parentId && visibleIds.has(person.id) && visibleIds.has(person.parentId))
    .map((person) => ({
      id: `${person.parentId}-${person.id}`,
      source: person.parentId!,
      target: person.id,
      type: "default",
      animated: false
    }));

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
