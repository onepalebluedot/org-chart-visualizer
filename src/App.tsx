import {
  Background,
  BackgroundVariant,
  BaseEdge,
  Handle,
  Panel,
  Position,
  ReactFlow,
  getViewportForBounds,
  getSmoothStepPath,
  type EdgeProps,
  type OnNodeDrag,
  type ReactFlowInstance,
  type XYPosition
} from "@xyflow/react";
import "@xyflow/react/dist/style.css";
import { useEffect, useLayoutEffect, useMemo, useRef, useState } from "react";
import mockOrgData from "./data/mockOrg";
import type { OrgData, PersonFormState, PersonRecord, RoleType, ViewMode } from "./types";
import {
  buildOrgDataFromRosterRows,
  buildOrgFlow,
  buildLocationFlow,
  childrenByParent,
  createEmptyPerson,
  filterPeople,
  getDefaultFocusedExpandedIds,
  getFocusedExpandableIds,
  isDescendant,
  normalizeOrgData,
  parseOrgData,
  reparentPerson,
  reorderSiblings,
  serializeOrgData
} from "./utils/org";
import type { AppNode, OrgNodeData, SpreadsheetImportPreview } from "./utils/org";

const roleLabels: Record<RoleType, string> = {
  executive: "Executive",
  manager: "Manager",
  ic: "Individual contributor",
  "open-role": "Open role"
};

const getInitials = (name: string): string => {
  if (name.toLowerCase() === "open role") return "+";
  const parts = name.trim().split(/\s+/).filter(Boolean);
  return (parts[0]?.[0] ?? "").concat(parts[1]?.[0] ?? "").toUpperCase() || "?";
};

const initialForm = (person: PersonRecord | null): PersonFormState => ({
  name: person?.name ?? "",
  role: person?.role ?? "",
  managerOrIc: person?.managerOrIc ?? "IC",
  workerType: person?.workerType ?? "Full Time",
  title: person?.title ?? "",
  managerName: person?.managerName ?? "",
  level: person?.level ? String(person.level) : "",
  location: person?.location ?? "",
});

const nodeTypes = {
  person: PersonNode,
  project: LocationNode
};

const edgeTypes = {
  reporting: ReportingEdge
};

type InteractionMode = "view" | "drag";
type CardDensity = "standard" | "light";
type ImportMessageTone = "error" | "success";

interface PendingSpreadsheetImport {
  fileName: string;
  data: OrgData;
  preview: SpreadsheetImportPreview;
}

interface ViewportAnchor {
  nodeId: string;
  screenPosition: XYPosition;
}

const STANDARD_PERSON_CARD_WIDTH = 260;
const STANDARD_PERSON_CARD_HEIGHT = 104;
const LIGHT_PERSON_CARD_WIDTH = 168;
const LIGHT_PERSON_CARD_HEIGHT = 42;
const GROUP_CARD_HEIGHT = 76;
const LOCATION_COLUMN_START_X = 120;
const GRID_SIZE = 20;
const EXPORT_WIDTH = 1600;
const EXPORT_HEIGHT = 900;

const getCanvasBounds = (
  nodes: AppNode[],
  personCardWidth: number,
  personCardHeight: number,
  xPadding = 0,
  yPadding = xPadding
): { x: number; y: number; width: number; height: number } => {
  if (nodes.length === 0) {
    return {
      x: -400,
      y: -200,
      width: 2000,
      height: 1400
    };
  }

  const bounds = nodes.reduce(
    (acc, node) => {
      const width =
        typeof node.style?.width === "number"
          ? node.style.width
          : node.type === "project"
            ? STANDARD_PERSON_CARD_WIDTH
            : personCardWidth;
      const height =
        typeof node.style?.height === "number"
          ? node.style.height
          : node.type === "project"
            ? GROUP_CARD_HEIGHT
            : personCardHeight;

      return {
        minX: Math.min(acc.minX, node.position.x),
        minY: Math.min(acc.minY, node.position.y),
        maxX: Math.max(acc.maxX, node.position.x + width),
        maxY: Math.max(acc.maxY, node.position.y + height)
      };
    },
    {
      minX: Number.POSITIVE_INFINITY,
      minY: Number.POSITIVE_INFINITY,
      maxX: Number.NEGATIVE_INFINITY,
      maxY: Number.NEGATIVE_INFINITY
    }
  );

  return {
    x: bounds.minX - xPadding,
    y: bounds.minY - yPadding,
    width: bounds.maxX - bounds.minX + xPadding * 2,
    height: bounds.maxY - bounds.minY + yPadding + xPadding
  };
};

export default function App() {
  const [orgData, setOrgData] = useState<OrgData>(normalizeOrgData(mockOrgData));
  const [viewMode, setViewMode] = useState<ViewMode>("org");
  const [interactionMode, setInteractionMode] = useState<InteractionMode>("view");
  const [cardDensity, setCardDensity] = useState<CardDensity>("standard");
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false);
  const [isMobileLayout, setIsMobileLayout] = useState(false);
  const [mobilePanelOpen, setMobilePanelOpen] = useState(false);
  const [selectedId, setSelectedId] = useState<string | null>(orgData.rootId);
  const [orgFocusId, setOrgFocusId] = useState<string | null>(orgData.rootId);
  const [canvasLocked, setCanvasLocked] = useState(false);
  const [expandedIds, setExpandedIds] = useState<Set<string>>(new Set([orgData.rootId]));
  const [search, setSearch] = useState("");
  const [orgPreviewPositions, setOrgPreviewPositions] = useState<Record<string, XYPosition>>({});
  const [locationPreviewPositions, setLocationPreviewPositions] = useState<Record<string, XYPosition>>({});
  const [dropTargetId, setDropTargetId] = useState<string | null>(null);
  const [invalidTargetId, setInvalidTargetId] = useState<string | null>(null);
  const [reorderTarget, setReorderTarget] = useState<{ id: string; placement: "before" | "after" } | null>(null);
  const [draggedNodeId, setDraggedNodeId] = useState<string | null>(null);
  const [pendingSpreadsheetImport, setPendingSpreadsheetImport] = useState<PendingSpreadsheetImport | null>(null);
  const [importMessage, setImportMessage] = useState<{ tone: ImportMessageTone; text: string } | null>(null);
  const [printMessage, setPrintMessage] = useState<{ tone: ImportMessageTone; text: string } | null>(null);
  const [isPrinting, setIsPrinting] = useState(false);
  const [formState, setFormState] = useState<PersonFormState>(initialForm(mockOrgData.people[0]));
  const jsonFileInputRef = useRef<HTMLInputElement | null>(null);
  const spreadsheetFileInputRef = useRef<HTMLInputElement | null>(null);
  const canvasFrameRef = useRef<HTMLDivElement | null>(null);
  const reactFlowInstanceRef = useRef<ReactFlowInstance<AppNode> | null>(null);
  const pendingViewportAnchorRef = useRef<ViewportAnchor | null>(null);
  const undoStackRef = useRef<OrgData[]>([]);
  const redoStackRef = useRef<OrgData[]>([]);
  const isLightMode = cardDensity === "light";
  const personCardWidth = isLightMode ? LIGHT_PERSON_CARD_WIDTH : STANDARD_PERSON_CARD_WIDTH;
  const personCardHeight = isLightMode ? LIGHT_PERSON_CARD_HEIGHT : STANDARD_PERSON_CARD_HEIGHT;
  const locationColumnWidth = isLightMode ? 176 : 278;

  const selectedPerson =
    orgData.people.find((person) => person.id === selectedId) ??
    orgData.people.find((person) => person.id === orgData.rootId) ??
    null;

  const filteredIds = useMemo(() => filterPeople(orgData.people, search), [orgData.people, search]);
  const orgFlow = useMemo(
    () =>
      buildOrgFlow({
        data: orgData,
        expandedIds,
        selectedId,
        focusId: orgFocusId,
        dropTargetId: reorderTarget ? null : dropTargetId,
        invalidTargetId,
        previewPositions: orgPreviewPositions,
        lightMode: isLightMode
      }),
    [dropTargetId, expandedIds, invalidTargetId, isLightMode, orgData, orgFocusId, orgPreviewPositions, reorderTarget, selectedId]
  );
  const locationFlow = useMemo(
    () =>
      buildLocationFlow({
        data: orgData,
        selectedId,
        previewPositions: locationPreviewPositions,
        lightMode: isLightMode
      }),
    [isLightMode, locationPreviewPositions, orgData, selectedId]
  );

  const visibleOrgNodes = useMemo(
    () => orgFlow.nodes.filter((node) => filteredIds.has(node.id) || node.id === orgData.rootId),
    [filteredIds, orgData.rootId, orgFlow.nodes]
  );

  const visibleOrgNodeIds = useMemo(() => new Set(visibleOrgNodes.map((node) => node.id)), [visibleOrgNodes]);
  const visibleOrgEdges = useMemo(
    () => orgFlow.edges.filter((edge) => visibleOrgNodeIds.has(edge.source) && visibleOrgNodeIds.has(edge.target)),
    [orgFlow.edges, visibleOrgNodeIds]
  );

  const visibleLocationNodes = useMemo(
    () =>
      locationFlow.nodes.filter((node) =>
        node.type === "project" ? true : filteredIds.has((node.data as OrgNodeData).person.id)
      ),
    [filteredIds, locationFlow.nodes]
  );

  const visibleLocationNodeIds = useMemo(() => new Set(visibleLocationNodes.map((node) => node.id)), [visibleLocationNodes]);
  const visibleLocationEdges = useMemo(
    () => locationFlow.edges.filter((edge) => visibleLocationNodeIds.has(edge.source) && visibleLocationNodeIds.has(edge.target)),
    [locationFlow.edges, visibleLocationNodeIds]
  );
  const roleSummary = useMemo(() => {
    const managers = orgData.people.filter((person) => person.managerOrIc === "Manager" && person.roleType !== "open-role").length;
    const openRoles = orgData.people.filter((person) => person.roleType === "open-role").length;
    const individualContributors = orgData.people.length - managers - openRoles;

    return {
      managers,
      individualContributors,
      openRoles
    };
  }, [orgData.people]);
  const orgChildrenMap = useMemo(() => childrenByParent(orgData.people), [orgData.people]);

  const captureViewportAnchor = () => {
    if (viewMode !== "org") return;
    const instance = reactFlowInstanceRef.current;
    const canvas = canvasFrameRef.current;
    if (!instance || !canvas || visibleOrgNodes.length === 0) return;

    const canvasRect = canvas.getBoundingClientRect();
    const candidates = visibleOrgNodes
      .filter((node) => node.type === "person")
      .map((node) => {
        const screenPosition = instance.flowToScreenPosition(node.position);
        const width = personCardWidth * instance.getViewport().zoom;
        const height = personCardHeight * instance.getViewport().zoom;
        const intersectsViewport =
          screenPosition.x + width >= canvasRect.left &&
          screenPosition.x <= canvasRect.right &&
          screenPosition.y + height >= canvasRect.top &&
          screenPosition.y <= canvasRect.bottom;

        return {
          node,
          screenPosition,
          intersectsViewport
        };
      })
      .filter((candidate) => candidate.intersectsViewport);

    const anchor =
      candidates.sort((a, b) => a.screenPosition.y - b.screenPosition.y || a.screenPosition.x - b.screenPosition.x)[0] ??
      visibleOrgNodes
        .filter((node) => node.type === "person")
        .map((node) => ({
          node,
          screenPosition: instance.flowToScreenPosition(node.position)
        }))
        .sort((a, b) => a.screenPosition.y - b.screenPosition.y || a.screenPosition.x - b.screenPosition.x)[0];

    if (!anchor) return;

    pendingViewportAnchorRef.current = {
      nodeId: anchor.node.id,
      screenPosition: anchor.screenPosition
    };
  };

  useLayoutEffect(() => {
    const anchor = pendingViewportAnchorRef.current;
    const instance = reactFlowInstanceRef.current;
    if (!anchor || !instance || viewMode !== "org") return;

    const nextAnchorNode = visibleOrgNodes.find((node) => node.id === anchor.nodeId);
    if (!nextAnchorNode) {
      pendingViewportAnchorRef.current = null;
      return;
    }

    const currentScreenPosition = instance.flowToScreenPosition(nextAnchorNode.position);
    const viewport = instance.getViewport();
    pendingViewportAnchorRef.current = null;

    void instance.setViewport(
      {
        x: viewport.x + anchor.screenPosition.x - currentScreenPosition.x,
        y: viewport.y + anchor.screenPosition.y - currentScreenPosition.y,
        zoom: viewport.zoom
      },
      { duration: 0 }
    );
  }, [viewMode, visibleOrgNodes]);

  useEffect(() => {
    setFormState(initialForm(selectedPerson));
  }, [selectedPerson]);

  useEffect(() => {
    if (typeof window === "undefined") return;

    const mediaQuery = window.matchMedia("(max-width: 980px)");
    const syncLayout = () => setIsMobileLayout(mediaQuery.matches);
    syncLayout();

    mediaQuery.addEventListener("change", syncLayout);
    return () => mediaQuery.removeEventListener("change", syncLayout);
  }, []);

  useEffect(() => {
    if (!isMobileLayout) {
      setMobilePanelOpen(false);
    }
  }, [isMobileLayout]);

  useEffect(() => {
    if (interactionMode === "drag") return;
    setOrgPreviewPositions({});
    setLocationPreviewPositions({});
    setDropTargetId(null);
    setInvalidTargetId(null);
    setReorderTarget(null);
    setDraggedNodeId(null);
  }, [interactionMode]);

  const commitOrgChange = (updater: (current: OrgData) => OrgData) => {
    setOrgData((current) => {
      undoStackRef.current.push(current);
      if (undoStackRef.current.length > 40) {
        undoStackRef.current.shift();
      }
      redoStackRef.current = [];
      return normalizeOrgData(updater(current));
    });
  };

  const updatePerson = (personId: string, patch: Partial<PersonRecord>) => {
    commitOrgChange((current) => ({
      ...current,
      people: current.people.map((person) => (person.id === personId ? { ...person, ...patch } : person))
    }));
  };

  const applyForm = (nextForm: PersonFormState) => {
    if (!selectedPerson) return;
    setFormState(nextForm);
    updatePerson(selectedPerson.id, {
      name: nextForm.name,
      role: nextForm.role,
      managerOrIc: nextForm.managerOrIc,
      workerType: nextForm.workerType,
      title: nextForm.title,
      managerName: nextForm.managerName,
      location: nextForm.location,
    });
  };

  const handleLevelInputChange = (value: string) => {
    setFormState((current) => ({
      ...current,
      level: value
    }));
  };

  const commitLevelInput = () => {
    if (!selectedPerson) return;

    const trimmed = formState.level.trim();

    if (!trimmed) {
      setFormState((current) => ({
        ...current,
        level: String(selectedPerson.level)
      }));
      return;
    }

    const nextLevel = Number(trimmed);
    if (!Number.isFinite(nextLevel)) {
      setFormState((current) => ({
        ...current,
        level: String(selectedPerson.level)
      }));
      return;
    }

    if (nextLevel !== selectedPerson.level) {
      updatePerson(selectedPerson.id, { level: nextLevel });
    } else {
      setFormState((current) => ({
        ...current,
        level: String(nextLevel)
      }));
    }
  };

  const addPerson = (roleType: RoleType) => {
    const parentId = selectedPerson?.id ?? orgData.rootId;
    const parentName = selectedPerson?.name ?? "";
    const nextPerson = createEmptyPerson(parentId, roleType);
    nextPerson.managerName = parentName;

    captureViewportAnchor();
    commitOrgChange((current) => ({
      ...current,
      people: [...current.people, nextPerson]
    }));
    setSelectedId(parentId);
    if (!canvasLocked) {
      setOrgFocusId(parentId ?? nextPerson.id);
      setExpandedIds((current) => new Set(current).add(parentId));
    }
  };

  const addLeader = () => {
    const currentRoot = orgData.people.find((person) => person.id === orgData.rootId) ?? orgData.people[0];
    if (!currentRoot) return;

    const nextLeader: PersonRecord = {
      ...createEmptyPerson(null, "executive"),
      name: "New Leader",
      role: "Leadership",
      managerOrIc: "Manager",
      workerType: currentRoot.workerType || "Full Time",
      title: "Leader",
      managerName: "",
      level: Number.isFinite(currentRoot.level) ? currentRoot.level + 1 : 1,
      location: currentRoot.location || "Remote",
      roleType: "executive"
    };

    captureViewportAnchor();
    commitOrgChange((current) => ({
      rootId: nextLeader.id,
      people: [
        nextLeader,
        ...current.people.map((person) =>
          person.id === current.rootId
            ? {
                ...person,
                parentId: nextLeader.id,
                managerName: nextLeader.name
              }
            : person
        )
      ]
    }));
    setSelectedId(nextLeader.id);
    if (!canvasLocked) {
      setOrgFocusId(nextLeader.id);
      setExpandedIds((current) => new Set([...current, nextLeader.id, currentRoot.id]));
    }
  };

  const deleteSelected = () => {
    if (!selectedPerson || selectedPerson.id === orgData.rootId) return;
    const childMap = childrenByParent(orgData.people);
    const descendants = new Set<string>();
    const collect = (id: string) => {
      descendants.add(id);
      for (const child of childMap[id] ?? []) {
        collect(child.id);
      }
    };
    collect(selectedPerson.id);

    captureViewportAnchor();
    commitOrgChange((current) => ({
      ...current,
      people: current.people.filter((person) => !descendants.has(person.id))
    }));
    setSelectedId(selectedPerson.parentId ?? orgData.rootId);
  };

  const downloadOrgData = (data: OrgData, fileBaseName = "org-chart") => {
    const blob = new Blob([serializeOrgData(data)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = `${fileBaseName}-${new Date().toISOString().slice(0, 10)}.json`;
    anchor.click();
    URL.revokeObjectURL(url);
  };

  const exportJson = () => {
    downloadOrgData(orgData);
  };

  const waitForNextPaint = () =>
    new Promise<void>((resolve) => {
      requestAnimationFrame(() => requestAnimationFrame(() => resolve()));
    });

  const printCurrentView = async () => {
    const instance = reactFlowInstanceRef.current;
    const canvas = canvasFrameRef.current;

    if (!instance || !canvas || renderedNodes.length === 0) {
      setPrintMessage({
        tone: "error",
        text: "The current view could not be exported."
      });
      return;
    }

    setIsPrinting(true);
    setPrintMessage(null);

    const previousViewport = instance.getViewport();

    try {
      const { toPng } = await import("html-to-image");
      await waitForNextPaint();

      const exportBounds = getCanvasBounds(
        renderedNodes,
        personCardWidth,
        personCardHeight,
        140,
        120
      );
      const exportViewport = getViewportForBounds(
        exportBounds,
        EXPORT_WIDTH,
        EXPORT_HEIGHT,
        0.25,
        1.4,
        0.02
      );

      await instance.setViewport(exportViewport, { duration: 0 });
      await waitForNextPaint();

      const dataUrl = await toPng(canvas, {
        cacheBust: true,
        pixelRatio: 2,
        width: EXPORT_WIDTH,
        height: EXPORT_HEIGHT,
        style: {
          width: `${EXPORT_WIDTH}px`,
          height: `${EXPORT_HEIGHT}px`,
          borderRadius: "0"
        }
      });

      const downloadLink = document.createElement("a");
      const densityLabel = cardDensity === "light" ? "light" : "full";
      downloadLink.href = dataUrl;
      downloadLink.download = `org-chart-${viewMode}-${densityLabel}-${new Date().toISOString().slice(0, 10)}.png`;
      downloadLink.click();

      setPrintMessage({
        tone: "success",
        text: "Landscape image exported."
      });
    } catch (error) {
      setPrintMessage({
        tone: "error",
        text: error instanceof Error ? error.message : "Image export failed."
      });
    } finally {
      await instance.setViewport(previousViewport, { duration: 0 });
      await waitForNextPaint();
      setIsPrinting(false);
    }
  };

  const loadOrgData = (nextData: OrgData) => {
    undoStackRef.current = [];
    redoStackRef.current = [];
    setOrgData(nextData);
    setSelectedId(nextData.rootId);
    setOrgFocusId(nextData.rootId);
    setExpandedIds(new Set([nextData.rootId]));
    setOrgPreviewPositions({});
    setLocationPreviewPositions({});
    setDropTargetId(null);
    setInvalidTargetId(null);
    setReorderTarget(null);
    setDraggedNodeId(null);
  };

  const importJson = async (file: File) => {
    const text = await file.text();
    const parsed = normalizeOrgData(parseOrgData(text));
    setPendingSpreadsheetImport(null);
    setImportMessage(null);
    loadOrgData(parsed);
  };

  const importSpreadsheet = async (file: File) => {
    try {
      const XLSX = await import("xlsx");
      const workbook = XLSX.read(await file.arrayBuffer(), { type: "array" });
      const sheetName = workbook.SheetNames[0];
      if (!sheetName) {
        throw new Error("The spreadsheet does not contain a readable sheet.");
      }

      const worksheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(worksheet, {
        defval: "",
        raw: false
      });
      const imported = buildOrgDataFromRosterRows(rows, sheetName);

      setPendingSpreadsheetImport({
        fileName: file.name,
        data: imported.data,
        preview: imported.preview
      });
      setImportMessage({
        tone: "success",
        text: `Parsed ${imported.preview.importedCount} roles from ${file.name}.`
      });
    } catch (error) {
      setPendingSpreadsheetImport(null);
      setImportMessage({
        tone: "error",
        text: error instanceof Error ? error.message : "The spreadsheet could not be imported."
      });
    }
  };

  const handleNodeSelection = (node: AppNode) => {
    if (node.type === "project") return;
    const personId = node.data.person.id;
    const parentId = node.data.person.parentId;
    const hasDirectReports = (orgChildrenMap[personId] ?? []).length > 0;
    if (!canvasLocked && viewMode === "org") {
      captureViewportAnchor();
      if (!hasDirectReports && parentId) {
        setOrgFocusId(parentId);
        setExpandedIds((current) => (current.has(parentId) ? current : new Set(current).add(parentId)));
      } else {
        setOrgFocusId(personId);
        setExpandedIds((current) => {
          if (!hasDirectReports || current.has(personId)) return current;
          const next = new Set(current);
          next.add(personId);
          return next;
        });
      }
    }
    setSelectedId(personId);
    if (isMobileLayout) {
      setMobilePanelOpen(true);
    }
  };

  const detectOrgDropTarget = (
    node: AppNode
  ): { valid: string | null; invalid: string | null; reorder: { id: string; placement: "before" | "after" } | null } => {
    const sourceId = node.id;
    const nodeRect = {
      left: node.position.x,
      top: node.position.y,
      right: node.position.x + personCardWidth,
      bottom: node.position.y + personCardHeight
    };

    for (const candidate of orgFlow.nodes) {
      if (candidate.id === sourceId) continue;

      const overlaps =
        nodeRect.right > candidate.position.x &&
        nodeRect.left < candidate.position.x + personCardWidth &&
        nodeRect.bottom > candidate.position.y &&
        nodeRect.top < candidate.position.y + personCardHeight;

      if (!overlaps) continue;

      const movingPerson = orgData.people.find((person) => person.id === sourceId);
      const candidatePerson = orgData.people.find((person) => person.id === candidate.id);
      const overlapWidth = Math.min(nodeRect.right, candidate.position.x + personCardWidth) - Math.max(nodeRect.left, candidate.position.x);
      const overlapRatio = overlapWidth / personCardWidth;

      const isPeerManagerReorder =
        movingPerson &&
        candidatePerson &&
        movingPerson.roleType === "manager" &&
        candidatePerson.roleType === "manager" &&
        movingPerson.parentId === candidatePerson.parentId &&
        overlapRatio < 0.42;

      if (isPeerManagerReorder) {
        return {
          valid: null,
          invalid: null,
          reorder: {
            id: candidate.id,
            placement: node.position.x < candidate.position.x ? "before" : "after"
          }
        };
      }

      const invalid =
        candidate.id === sourceId ||
        !movingPerson ||
        !candidatePerson ||
        isDescendant(orgData.people, sourceId, candidate.id) ||
        orgData.rootId === sourceId;

      if (invalid) {
        return { valid: null, invalid: candidate.id, reorder: null };
      }

      return { valid: candidate.id, invalid: null, reorder: null };
    }

    return { valid: null, invalid: null, reorder: null };
  };

  const onOrgNodeDrag: OnNodeDrag<AppNode> = (_, node) => {
    if (interactionMode !== "drag") return;
    setDraggedNodeId(node.id);
    const target = detectOrgDropTarget(node);
    setDropTargetId(target.valid);
    setInvalidTargetId(target.invalid);
    setReorderTarget(target.reorder);
    setOrgPreviewPositions({
      [node.id]: {
        x: Math.round(node.position.x / GRID_SIZE) * GRID_SIZE,
        y: Math.round(node.position.y / GRID_SIZE) * GRID_SIZE
      }
    });
  };

  const onOrgNodeDragStop: OnNodeDrag<AppNode> = (_, node) => {
    if (interactionMode !== "drag") return;
    const target = detectOrgDropTarget(node);
    if (target.reorder) {
      captureViewportAnchor();
      commitOrgChange((current) => ({
        ...current,
        people: reorderSiblings(current.people, node.id, target.reorder!.id, target.reorder!.placement)
      }));
    } else if (target.valid) {
      captureViewportAnchor();
      commitOrgChange((current) => ({
        ...current,
        people: reparentPerson(current.people, node.id, target.valid!)
      }));
    }
    setOrgPreviewPositions({});
    setDropTargetId(null);
    setInvalidTargetId(null);
    setReorderTarget(null);
    setDraggedNodeId(null);
  };

  const detectLocationDropTarget = (node: AppNode): string | null => {
    const locationHeaders = locationFlow.nodes.filter((candidate) => candidate.type === "project");
    if (locationHeaders.length === 0) return null;

    const centerX = node.position.x + personCardWidth / 2;
    const closest = locationHeaders.reduce<{ id: string; distance: number } | null>((best, header) => {
      const distance = Math.abs(centerX - (header.position.x + STANDARD_PERSON_CARD_WIDTH / 2));
      if (!best || distance < best.distance) {
        return { id: header.id, distance };
      }
      return best;
    }, null);

    return closest && closest.distance <= locationColumnWidth / 2 ? closest.id : null;
  };

  const onLocationNodeDrag: OnNodeDrag<AppNode> = (_, node) => {
    if (interactionMode !== "drag") return;
    if (node.type === "project") return;

    setDraggedNodeId(node.id);
    const targetId = detectLocationDropTarget(node);
    setDropTargetId(targetId);
    setInvalidTargetId(null);

    if (targetId) {
      const targetNode = locationFlow.nodes.find((candidate) => candidate.id === targetId);
      if (targetNode) {
        setLocationPreviewPositions({
          [node.id]: {
            x: targetNode.position.x,
            y: Math.round(node.position.y / GRID_SIZE) * GRID_SIZE
          }
        });
        return;
      }
    }

    const snappedColumnIndex = Math.max(
      0,
      Math.round((node.position.x - LOCATION_COLUMN_START_X) / locationColumnWidth)
    );

    setLocationPreviewPositions({
      [node.id]: {
        x: LOCATION_COLUMN_START_X + snappedColumnIndex * locationColumnWidth,
        y: Math.round(node.position.y / GRID_SIZE) * GRID_SIZE
      }
    });
  };

  const onLocationNodeDragStop: OnNodeDrag<AppNode> = (_, node) => {
    if (interactionMode !== "drag") return;
    if (node.type === "project") return;

    const targetId = detectLocationDropTarget(node);
    if (targetId) {
      updatePerson(node.data.person.id, {
        location: targetId.replace("location:", "")
      });
    }

    setLocationPreviewPositions({});
    setDropTargetId(null);
    setInvalidTargetId(null);
    setReorderTarget(null);
    setDraggedNodeId(null);
  };

  const getCollapseTargetId = (): string | null => {
    if (canvasLocked) return orgFocusId;
    if (selectedPerson && (orgChildrenMap[selectedPerson.id] ?? []).length > 0) {
      return selectedPerson.id;
    }
    return orgFocusId;
  };

  const toggleCollapse = () => {
    const targetId = getCollapseTargetId();
    if (!targetId) return;
    if ((orgChildrenMap[targetId] ?? []).length === 0) return;

    captureViewportAnchor();
    setExpandedIds((current) => {
      const next = new Set(current);
      if (next.has(targetId)) {
        next.delete(targetId);
      } else {
        next.add(targetId);
      }
      return next;
    });
  };

  const collapseAll = () => {
    captureViewportAnchor();
    setExpandedIds(new Set());
  };

  const expandAll = () => {
    captureViewportAnchor();
    setExpandedIds(getFocusedExpandableIds(orgData, orgFocusId ?? orgData.rootId));
  };

  const locationSummary = useMemo(() => {
    const counts = new Map<string, number>();
    for (const person of orgData.people) {
      counts.set(person.location, (counts.get(person.location) ?? 0) + 1);
    }
    return [...counts.entries()].sort(([a], [b]) => a.localeCompare(b));
  }, [orgData.people]);

  const isDragEditing = interactionMode === "drag";
  const activeNodes = viewMode === "org" ? visibleOrgNodes : visibleLocationNodes;
  const activeEdges = viewMode === "org" ? visibleOrgEdges : visibleLocationEdges;
  const renderedNodes = useMemo(
    () =>
      activeNodes.map((node) => ({
        ...node,
        draggable:
          isDragEditing &&
          node.draggable !== false &&
          !(node.type === "person" && "person" in node.data && node.data.person.id === orgData.rootId)
      })),
    [activeNodes, isDragEditing, orgData.rootId]
  );
  const collapseTargetId = getCollapseTargetId();
  const canCollapseSelection = !!collapseTargetId && (orgChildrenMap[collapseTargetId] ?? []).length > 0;
  const collapseLabel =
    collapseTargetId && canCollapseSelection && expandedIds.has(collapseTargetId) ? "Collapse selected" : "Expand selected";
  const canUndo = undoStackRef.current.length > 0;
  const canRedo = redoStackRef.current.length > 0;
  const undoChange = () =>
    setOrgData((current) => {
      const previous = undoStackRef.current.pop();
      if (!previous) return current;
      redoStackRef.current.push(current);
      return normalizeOrgData(previous);
    });
  const redoChange = () =>
    setOrgData((current) => {
      const next = redoStackRef.current.pop();
      if (!next) return current;
      undoStackRef.current.push(current);
      return normalizeOrgData(next);
    });

  return (
    <div
      className={`app-shell ${sidebarCollapsed && !isMobileLayout ? "sidebar-collapsed" : ""} ${
        isMobileLayout ? "mobile-layout" : ""
      }`}
    >
      {isMobileLayout && mobilePanelOpen ? <div className="mobile-backdrop" onClick={() => setMobilePanelOpen(false)} /> : null}
      <aside
        className={`sidebar ${sidebarCollapsed && !isMobileLayout ? "is-collapsed" : ""} ${isMobileLayout ? "is-mobile" : ""} ${
          isMobileLayout && mobilePanelOpen ? "is-mobile-open" : ""
        }`}
      >
        <div className="sidebar-toggle-row">
          <button
            className="sidebar-toggle"
            onClick={() => {
              if (isMobileLayout) {
                setMobilePanelOpen((current) => !current);
                return;
              }
              setSidebarCollapsed((current) => !current);
            }}
          >
            {isMobileLayout ? (mobilePanelOpen ? "Close panel" : "Open panel") : sidebarCollapsed ? "Show panel" : "Hide panel"}
          </button>
        </div>

        {sidebarCollapsed && !isMobileLayout ? (
          <div className="sidebar-collapsed-content">
            <div className="collapsed-brand">
              <span>OC</span>
            </div>
            <div className="collapsed-meta">
              <strong>{viewMode === "org" ? "Org" : "Location"}</strong>
              <span>{orgData.people.length} roles</span>
            </div>
          </div>
        ) : (
          <>
        <div className="sidebar-section">
          <p className="eyebrow">Workspace</p>
          <h1>Org Chart Visualizer</h1>
          <p className="lede">
            Local planning tool for org design, staffing placeholders, and coverage planning across hierarchy and
            location views.
          </p>
        </div>

        <div className="sidebar-section selected-card-section">
          <div className="section-title-row">
            <h2>Selected card</h2>
            {selectedPerson && selectedPerson.id !== orgData.rootId ? (
              <button className="danger-link" onClick={deleteSelected}>
                Remove
              </button>
            ) : null}
          </div>
          {selectedPerson ? (
            <div className="editor-form">
              <label>
                <span>Name</span>
                <input
                  value={formState.name}
                  onChange={(event) => applyForm({ ...formState, name: event.target.value })}
                />
              </label>
              <label>
                <span>Role</span>
                <input
                  value={formState.role}
                  onChange={(event) => applyForm({ ...formState, role: event.target.value })}
                />
              </label>
              <label>
                <span>Manager or IC</span>
                <select
                  value={formState.managerOrIc}
                  onChange={(event) =>
                    applyForm({ ...formState, managerOrIc: event.target.value as PersonFormState["managerOrIc"] })
                  }
                >
                  <option value="Manager">Manager</option>
                  <option value="IC">IC</option>
                </select>
              </label>
              <label>
                <span>Full Time or Contractor</span>
                <input
                  value={formState.workerType}
                  onChange={(event) => applyForm({ ...formState, workerType: event.target.value })}
                />
              </label>
              <label>
                <span>Title</span>
                <input
                  value={formState.title}
                  onChange={(event) => applyForm({ ...formState, title: event.target.value })}
                />
              </label>
              <label>
                <span>Manager</span>
                <input value={formState.managerName} readOnly />
              </label>
              <label>
                <span>Level</span>
                <input
                  value={formState.level}
                  inputMode="numeric"
                  onChange={(event) => handleLevelInputChange(event.target.value)}
                  onBlur={commitLevelInput}
                />
              </label>
              <label>
                <span>Location</span>
                <input
                  value={formState.location}
                  onChange={(event) => applyForm({ ...formState, location: event.target.value })}
                />
              </label>
            </div>
          ) : (
            <p className="muted">Select a person card to edit details.</p>
          )}
        </div>

        <div className="sidebar-section">
          <p className="section-kicker">Canvas</p>
          <div className="toolbar-grid">
            <button
              onClick={() => setInteractionMode("view")}
              className={interactionMode === "view" ? "selected-action" : ""}
            >
              View mode
            </button>
            <button
              onClick={() => setInteractionMode("drag")}
              className={interactionMode === "drag" ? "selected-action" : ""}
            >
              Drag edit
            </button>
            <button onClick={() => setViewMode("org")} className={viewMode === "org" ? "selected-action" : ""}>
              Org view
            </button>
            <button onClick={() => setViewMode("location")} className={viewMode === "location" ? "selected-action" : ""}>
              Location view
            </button>
            <button onClick={() => setCardDensity("standard")} className={cardDensity === "standard" ? "selected-action" : ""}>
              Full cards
            </button>
            <button onClick={() => setCardDensity("light")} className={cardDensity === "light" ? "selected-action" : ""}>
              Light view
            </button>
          </div>
        </div>

        <div className="sidebar-section">
          <p className="section-kicker">Files</p>
          <div className="toolbar-grid">
            <button onClick={exportJson}>Save</button>
            <button onClick={() => jsonFileInputRef.current?.click()}>Load JSON</button>
            <button onClick={() => spreadsheetFileInputRef.current?.click()}>Import spreadsheet</button>
            <button onClick={() => void printCurrentView()} disabled={isPrinting}>
              {isPrinting ? "Preparing image" : "Print image"}
            </button>
          </div>
          <p className="muted import-note">
            Expected columns: Name, Role, Manager Or IC, Full Time or Contractor, Title, Manager, Level, Location.
          </p>
          {printMessage ? (
            <p className={`import-status ${printMessage.tone === "error" ? "is-error" : "is-success"}`}>
              {printMessage.text}
            </p>
          ) : null}
          {importMessage ? (
            <p className={`import-status ${importMessage.tone === "error" ? "is-error" : "is-success"}`}>
              {importMessage.text}
            </p>
          ) : null}
          {pendingSpreadsheetImport ? (
            <div className="import-preview">
              <div className="import-preview-grid">
                <span>File</span>
                <strong>{pendingSpreadsheetImport.fileName}</strong>
                <span>Sheet</span>
                <strong>{pendingSpreadsheetImport.preview.sheetName}</strong>
                <span>Rows</span>
                <strong>{pendingSpreadsheetImport.preview.rowCount}</strong>
                <span>Imported</span>
                <strong>{pendingSpreadsheetImport.preview.importedCount}</strong>
                <span>Root</span>
                <strong>{pendingSpreadsheetImport.preview.rootName}</strong>
              </div>
              {pendingSpreadsheetImport.preview.warnings.length > 0 ? (
                <div className="import-preview-block">
                  <span className="field-label">Warnings</span>
                  <ul className="import-list">
                    {pendingSpreadsheetImport.preview.warnings.map((warning) => (
                      <li key={warning}>{warning}</li>
                    ))}
                  </ul>
                </div>
              ) : null}
              {pendingSpreadsheetImport.preview.duplicateNames.length > 0 ? (
                <div className="import-preview-block">
                  <span className="field-label">Duplicate names</span>
                  <p className="muted">{pendingSpreadsheetImport.preview.duplicateNames.join(", ")}</p>
                </div>
              ) : null}
              {pendingSpreadsheetImport.preview.unresolvedManagers.length > 0 ? (
                <div className="import-preview-block">
                  <span className="field-label">Managers not found</span>
                  <p className="muted">{pendingSpreadsheetImport.preview.unresolvedManagers.join(", ")}</p>
                </div>
              ) : null}
              <div className="toolbar-grid compact-toolbar">
                <button
                  onClick={() => {
                    setImportMessage(null);
                    loadOrgData(pendingSpreadsheetImport.data);
                    setPendingSpreadsheetImport(null);
                  }}
                  className="selected-action"
                >
                  Load imported data
                </button>
                <button
                  onClick={() =>
                    downloadOrgData(
                      pendingSpreadsheetImport.data,
                      pendingSpreadsheetImport.fileName.replace(/\.[^.]+$/, "") || "converted-org"
                    )
                  }
                >
                  Download JSON
                </button>
              </div>
            </div>
          ) : null}
          {pendingSpreadsheetImport ? (
            <button className="clear-button" onClick={() => setPendingSpreadsheetImport(null)}>
              Clear import preview
            </button>
          ) : null}
        </div>

        <div className="sidebar-section">
          <p className="section-kicker">Search</p>
          <label className="field-label" htmlFor="search">
            Search
          </label>
          <input
            id="search"
            value={search}
            onChange={(event) => setSearch(event.target.value)}
            placeholder="Name, role, title, manager, location"
          />
        </div>

        <div className="sidebar-section">
          <h2>Locations</h2>
          <div className="project-list">
            {locationSummary.map(([location, count]) => (
              <div key={location} className="project-pill">
                <span>{location}</span>
                <strong>{count}</strong>
              </div>
            ))}
          </div>
        </div>

        <input
          ref={jsonFileInputRef}
          className="hidden-input"
          type="file"
          accept="application/json"
          onChange={(event) => {
            const file = event.target.files?.[0];
            if (file) {
              void importJson(file);
            }
            event.currentTarget.value = "";
          }}
        />
        <input
          ref={spreadsheetFileInputRef}
          className="hidden-input"
          type="file"
          accept=".csv,.xlsx,.xls,text/csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"
          onChange={(event) => {
            const file = event.target.files?.[0];
            if (file) {
              void importSpreadsheet(file);
            }
            event.currentTarget.value = "";
          }}
        />
          </>
        )}
      </aside>

      <main className="canvas-shell">
        <div className="canvas-header">
          <div className="canvas-title-block">
            <p className="eyebrow">View</p>
            <h2>{viewMode === "org" ? "Reporting hierarchy" : "Location grouping"}</h2>
            {isMobileLayout ? (
              <div className="mobile-quick-actions">
                <button onClick={() => setMobilePanelOpen(true)}>Panel</button>
                <button onClick={() => setViewMode((current) => (current === "org" ? "location" : "org"))}>
                  {viewMode === "org" ? "Location view" : "Org view"}
                </button>
                <button onClick={() => setInteractionMode((current) => (current === "view" ? "drag" : "view"))}>
                  {isDragEditing ? "View mode" : "Drag edit"}
                </button>
                <button onClick={() => setCardDensity((current) => (current === "standard" ? "light" : "standard"))}>
                  {isLightMode ? "Full cards" : "Light view"}
                </button>
              </div>
            ) : null}
          </div>
          {!isMobileLayout ? (
            <div className="top-edit-bar" aria-label="Edit controls">
              <div className="top-control-group add-control-group">
                <span>Add</span>
                <div className="top-control-buttons">
                  <button className="add-person-button" onClick={() => addPerson("ic")}>
                    Person
                  </button>
                  <button className="add-role-button" onClick={() => addPerson("open-role")}>
                    Role
                  </button>
                  <button className="add-leader-button" onClick={addLeader}>
                    Leader
                  </button>
                </div>
              </div>
              <div className="top-control-group structure-control-group">
                <span>Structure</span>
                <div className="top-control-buttons">
                  <button onClick={toggleCollapse} disabled={!canCollapseSelection}>
                    {collapseLabel}
                  </button>
                  <button onClick={collapseAll}>Collapse all</button>
                  <button onClick={expandAll}>Expand all</button>
                </div>
              </div>
            </div>
          ) : null}
          <div className="canvas-meta">
            <span>
              {roleSummary.managers} managers · {roleSummary.individualContributors} IC · {roleSummary.openRoles} open
            </span>
            <span>{search ? `${filteredIds.size} matching` : "All visible"}</span>
            <span>{isDragEditing ? "Drag editing enabled" : "Pan and inspect mode"}</span>
            {draggedNodeId ? (
              <span className="drop-hint">
                {reorderTarget
                  ? `Release to move team ${reorderTarget.placement === "before" ? "left" : "right"}`
                  : dropTargetId
                  ? viewMode === "org"
                    ? "Release to reassign reporting line"
                    : "Release to update location"
                  : invalidTargetId
                    ? "Invalid drop target"
                    : "Release to return to the grid"}
              </span>
            ) : null}
          </div>
        </div>

        <div
          ref={canvasFrameRef}
          className={`canvas-frame ${isDragEditing ? "is-drag-editing" : "is-viewing"} ${
            viewMode === "org" ? "is-org-view" : "is-location-view"
          } ${isPrinting ? "is-exporting" : ""}`}
        >
          <ReactFlow<AppNode>
            nodes={renderedNodes}
            edges={activeEdges}
            nodeTypes={nodeTypes}
            edgeTypes={edgeTypes}
            onNodeClick={(_, node) => handleNodeSelection(node)}
            onNodeDrag={isDragEditing ? (viewMode === "org" ? onOrgNodeDrag : onLocationNodeDrag) : undefined}
            onNodeDragStop={isDragEditing ? (viewMode === "org" ? onOrgNodeDragStop : onLocationNodeDragStop) : undefined}
            minZoom={0.25}
            maxZoom={1.5}
            fitView
            fitViewOptions={{ padding: isMobileLayout ? 0.18 : 0.12, minZoom: 0.25, maxZoom: 1 }}
            snapToGrid
            snapGrid={[GRID_SIZE, GRID_SIZE]}
            nodesDraggable={isDragEditing}
            nodesConnectable={false}
            elementsSelectable
            zoomOnScroll
            zoomOnPinch
            panOnScroll={false}
            panOnDrag={isDragEditing ? (isMobileLayout ? true : [1]) : true}
            autoPanOnNodeFocus={false}
            selectionOnDrag={false}
            selectNodesOnDrag={false}
            nodeDragThreshold={1}
            proOptions={{ hideAttribution: true }}
            onInit={(instance) => {
              reactFlowInstanceRef.current = instance;
            }}
          >
            <Panel position="top-right">
              <div className="canvas-chip-row">
                <button
                  className={`canvas-lock-button ${canvasLocked ? "is-locked" : ""}`}
                  onClick={() => setCanvasLocked((current) => !current)}
                  type="button"
                >
                  {canvasLocked ? "Unlock focus" : "Lock focus"}
                </button>
                <div className="view-chip">
                  {viewMode === "org" ? "Hierarchy view" : "Location view"} · {isDragEditing ? "Drag edit" : "View"}
                  {viewMode === "org" && canvasLocked ? " · Locked" : ""}
                </div>
              </div>
            </Panel>
            <Panel position="bottom-left">
              <div className="canvas-history-controls" aria-label="History controls">
                <button onClick={undoChange} disabled={!canUndo} type="button">
                  Undo
                </button>
                <button onClick={redoChange} disabled={!canRedo} type="button">
                  Redo
                </button>
              </div>
            </Panel>
            <Background
              color={viewMode === "org" ? "rgba(49, 66, 245, 0.08)" : "#d8dde6"}
              variant={BackgroundVariant.Dots}
              gap={viewMode === "org" ? 28 : 18}
              size={1}
            />
          </ReactFlow>
        </div>
      </main>
    </div>
  );
}

function ReportingEdge({
  id,
  sourceX,
  sourceY,
  targetX,
  targetY,
  sourcePosition,
  targetPosition,
  style,
  markerEnd
}: EdgeProps) {
  const [edgePath] = getSmoothStepPath({
    sourceX,
    sourceY,
    sourcePosition,
    targetX,
    targetY,
    targetPosition,
    borderRadius: 12
  });

  return (
    <>
      <BaseEdge id={`${id}-halo`} path={edgePath} className="reporting-edge-halo" />
      <BaseEdge id={id} path={edgePath} markerEnd={markerEnd} style={style} className="reporting-edge" />
    </>
  );
}

function PersonNode({ data }: { data: OrgNodeData }) {
  const {
    person,
    selected,
    collapsed,
    isDropTarget,
    isInvalidTarget,
    isFocusGroup,
    directReportCount,
    viewMode,
    lightMode
  } = data;
  const roleToneClass = `tone-${person.role.toLowerCase().replace(/[^a-z0-9]+/g, "-")}`;
  const cardClasses = [
    "person-card",
    `role-${person.roleType}`,
    selected ? "is-selected" : "",
    collapsed ? "is-collapsed" : "",
    isFocusGroup ? "is-focus-open" : "",
    isDropTarget ? "is-drop-target" : "",
    isInvalidTarget ? "is-invalid-target" : "",
    viewMode === "location" ? "is-project-card" : "",
    lightMode ? "is-light" : ""
  ]
    .filter(Boolean)
    .join(" ");

  return (
    <div className={cardClasses}>
      {viewMode === "org" ? (
        <>
          <Handle id="report-target-top" type="target" position={Position.Top} className="reporting-handle" />
          <Handle id="report-target-left" type="target" position={Position.Left} className="reporting-handle" />
          <Handle id="report-target-right" type="target" position={Position.Right} className="reporting-handle" />
          <Handle id="report-source-bottom" type="source" position={Position.Bottom} className="reporting-handle" />
        </>
      ) : null}
      {!lightMode && viewMode === "org" ? <div className="card-avatar" aria-hidden="true">{getInitials(person.name)}</div> : null}
      {!lightMode ? (
        <>
          <div className="card-topline">
            <span className={`role-badge ${roleToneClass}`}>{person.role}</span>
            <span className={`status-tag track-${person.managerOrIc.toLowerCase()}`}>{person.managerOrIc}</span>
          </div>
          <div className="card-main">
            <strong>{person.name}</strong>
            <p>
              {person.title} <span className="card-divider">|</span> L{person.level}{" "}
              <span className="card-divider">|</span> {person.location}
            </p>
          </div>
        </>
      ) : (
        <div className="card-main">
          <strong>{person.name}</strong>
        </div>
      )}
      {directReportCount > 0 ? <span className="report-count-badge">{directReportCount}</span> : null}
    </div>
  );
}

function LocationNode({ data }: { data: OrgNodeData }) {
  return (
    <div className="project-header-card">
      <p>{data.person.name}</p>
      <span>{data.person.workerType}</span>
    </div>
  );
}
