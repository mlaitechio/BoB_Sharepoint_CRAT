import { ICircularListItem } from "../../Models/IModel";

export interface IEditDashBoardState {
    listItems: any[];
    accordionFields?: any;
    items?: ICircularListItem[];
    filteredItems?: ICircularListItem[];
    currentPage?: number;
    itemsPerPage?: number;
    showDeleteDialog?: boolean;
    currentSelectedItemId?: any;
    currentSelectedItem?: ICircularListItem;
    editFormItem?: any;
    loadDashBoard?: boolean;
    isLoading?: boolean;
    currentPageName?: string;
    openSupportingDoc?: boolean;
    supportingDocItem?: ICircularListItem;
    loadEditForm?: boolean;
    loadViewForm?: boolean;

}