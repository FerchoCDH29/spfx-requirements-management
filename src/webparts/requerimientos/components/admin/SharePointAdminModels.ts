export interface ISharePointListInfo {
  Id: string;
  Title: string;
  ItemCount: number;
  Hidden: boolean;
  BaseTemplate: number;
}

export interface ISharePointFieldInfo {
  Title: string;
  InternalName: string;
  TypeAsString: string;
  Hidden: boolean;
  ReadOnlyField: boolean;
  Required: boolean;
  Choices?: string[];
  LookupList?: string;
  LookupField?: string;
}

export interface ISharePointItem {
  Id: number;
  [key: string]: unknown;
}

export interface ISharePointUser {
  Id: number;
  Title: string;
  Email: string;
  LoginName: string;
}

export interface ILookupOption {
  Id: number;
  Label: string;
}

export type RestMethod = 'GET' | 'POST' | 'MERGE' | 'DELETE';
