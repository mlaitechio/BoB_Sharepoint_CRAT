import { INavLinkGroup, ITag } from "@fluentui/react";
import { ICheckBoxCollection, ICircularListItem, IListItem } from "../../Models/IModel";

export interface ICircularSearchState {
    navLinkGroups?: INavLinkGroup[];
    selectedKey?: string;
    searchText?: string;
    openSearchFilters?: boolean;
    supportingDocItem?: ICircularListItem;
    openSupportingDoc?: boolean;
    items?: ICircularListItem[];
    filteredItems?: ICircularListItem[];
    columns?: any[];
    currentPage?: number;
    sortingOptions?: any[];
    itemsPerPage?: number;
    renderEmpty?: boolean;
    isLoading?: boolean;
    previewItems?: ICircularListItem;
    currentSelectedItemId?: any;
    currentSelectedItem?: ICircularListItem;
    relevanceDepartment?: any[];
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
    checkBoxCollection?: Map<string, ICheckBoxCollection[]>;
    isFilterPanel?: boolean;
    publishedYear?: any[];
    filterLabelName?: string;
    filterAccordion?: any;
    openPanelCheckedValues?: ICheckBoxCollection[];
    filterPanelCheckBoxCollection?: Map<string, ICheckBoxCollection[]>;

    showHideFilters?: boolean;
    publishedStartDate?: Date;
    publishedEndDate?: Date | null;
    createdBy?: string;

}