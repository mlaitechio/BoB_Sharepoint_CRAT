import { INavLinkGroup, ITag } from "@fluentui/react";
import { ICircularListItem, IListItem } from "../../Models/IModel";

export interface ICircularSearchState {
    navLinkGroups?: INavLinkGroup[];
    selectedKey?: string;
    searchText?: string;
    items?: IListItem[];
    filteredItems?: IListItem[];
    columns?: any[];
    currentPage?: number;
    sortingOptions?: any[];
    itemsPerPage?: number;
    renderEmpty?: boolean;
    isLoading?: boolean;
    previewItems?: ICircularListItem;
    currentSelectedItemId?: any;
    currentSelectedItem?: ICircularListItem;
    sortingFields?: string;
    selectedSortFields?: string;
    sortDirection?: string;
    isVertical?: boolean;
    isFinancialYear?: boolean;
    filterValue?: string;
    searchItems?: IListItem[];
    departments?: any[];
    selectedDepartment?: string[];
    circularNumber?: string;
    checkCircularRefiner?: string;
    circularRefinerOperator?: string;
    switchSearchText?: string;
    isNormalSearch?: boolean;
    isSearchNavOpen?: boolean;
    currentSelectedFile?: any[];
    filePreviewItem?: any;
    isAccordionSelected?: any;
    accordionFields?: any;
    openFileViewer?: boolean;

    // selectedFinancialYear: string;
    // selectedVertical: string;
    // selectedFinanicalPicker: ITag[];
    // selectedVerticalPicker: ITag[];
    showHideFilters?: boolean;
    // documentCreator: any[];
    // documentCategoryDD: any[];
    // selectedDocCategory: ITag[];
    publishedStartDate?: Date;
    publishedEndDate?: Date | null;
    // modifiedStartDate?: Date;
    // modifiedEndDate?: Date | null;
    createdBy?: string;

}