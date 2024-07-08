import { ICircularListItem } from "../../Models/IModel";

export interface IEditDashBoardState {
    listItems: any[];
    accordionFields?: any;
    currentSelectedItemId?: any;
    currentSelectedItem?: ICircularListItem;
    editFormItem?: any;
    isItemEdited?: boolean;
    isLoading?: boolean;
    openSupportingDoc?: boolean;
    supportingDocItem?: ICircularListItem;
    
}