import { ICircularListItem } from "../../../Models/IModel";

export interface ISupportingDocumentState {
    isSupportingPanelOpen?: boolean;
    circularCollection: any[];
    allItems:any[];
    checkBoxCollection: Map<string, any[]>;
    optionsYear: any[];
    selectedYear: string;
    filterPanelCheckBoxCollection: Map<string, any[]>
}