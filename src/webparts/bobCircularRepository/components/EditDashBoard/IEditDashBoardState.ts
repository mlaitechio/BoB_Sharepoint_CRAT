import { ICircularListItem } from "../../Models/IModel";

export interface IEditDashBoardState {
    listItems: any[];
    accordionFields?: any;
    currentSelectedItemId?: any;
    currentSelectedItem?: ICircularListItem;
    isLoading?: boolean;
    openSupportingDoc?: boolean;
    supportingDocItem?: ICircularListItem;
    isItemEdited?: boolean;
}