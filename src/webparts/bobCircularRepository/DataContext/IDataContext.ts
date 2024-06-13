import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IDataContext {
    webPartContext: WebPartContext;
    loggedInUser?:string;
  }