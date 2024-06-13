import { ResponsiveMode } from "@fluentui/react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IServices } from "../services/IServices";

export interface IBobCircularRepositoryProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context?: WebPartContext;
  responsiveMode?: ResponsiveMode;
  serverRelativeUrl?: any;
  services?: IServices;
  isUserMaker?: boolean;
  isUserAdmin?: boolean;
  isUserChecker?: boolean;
  isUserCompliance?: boolean;
  sizeLimit?: string;
  publishingDays?: Number;
  circularListID?: string;
  updateItem?: (itemID: any) => void;
}
