import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ICircularListItem } from "../../Models/IModel";

export interface IFileViewerProps {
    listItem: ICircularListItem;
    stateKey?: string;
    context?: WebPartContext
    onClose?: () => void;
    onUpdate?: (itemID: any) => void;

}