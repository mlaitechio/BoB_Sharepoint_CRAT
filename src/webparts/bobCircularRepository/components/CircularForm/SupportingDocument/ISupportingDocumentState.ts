import { ICircularListItem } from "../../../Models/IModel";

export interface ISupportingDocumentState {
    isSupportingPanelOpen?: boolean;
    circularCollection: any[];
    checkBoxCollection: Map<string, any[]>;
    filterPanelCheckBoxCollection:Map<string,any[]>
}