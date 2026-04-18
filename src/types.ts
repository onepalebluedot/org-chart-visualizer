export type RoleType = "executive" | "manager" | "ic" | "open-role";
export type ViewMode = "org" | "location";

export interface PersonRecord {
  id: string;
  parentId: string | null;
  sortOrder?: number;
  name: string;
  role: string;
  managerOrIc: "Manager" | "IC";
  workerType: string;
  title: string;
  managerName: string;
  level: number;
  location: string;
  roleType: RoleType;
}

export interface OrgData {
  rootId: string;
  people: PersonRecord[];
}

export interface LayoutNode {
  id: string;
  x: number;
  y: number;
}

export interface PersonFormState {
  name: string;
  role: string;
  managerOrIc: "Manager" | "IC";
  workerType: string;
  title: string;
  managerName: string;
  level: string;
  location: string;
}
