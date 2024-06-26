export interface IServices {
    getPagedListItems: (serverRelativeUrl: string, listName: string, selectColumns: string, filterString: string, expandColumns: string, orderByColum: string, asc?: boolean) => Promise<any>;
    getLargeListItems: (serverRelativeUrl: string, listName: string, selectedColumn, expandColumns) => Promise<any[]>;
    addListItemAttachments(serverRelativeUrl: string, listName: string, itemID: number, fileMetadata: Map<string, any>): Promise<any>
    recycleListItemAttachments: (serverRelativeUrl: string, listName: string, itemID: number, files: Map<string, any>) => Promise<any>;
    readListItemAttachment: (serverRelativeUrl: string, listName: string, itemId: number) => Promise<any[]>
    updateItem: (serverRelativeUrl: string, listName: string, itemID: number, metadataValues: any, etag: any) => Promise<any>;
    createItem: (serverRelativeUrl: string, listName: string, metadataValues: any) => Promise<boolean | any>;
    deleteListItem: (serverRelativeUrl: string, listName: string, itemId: number) => Promise<any>;
    checkIfUserBelongToGroup: (groupName: any, userEmail: string) => Promise<any>;
    readFieldValues: (serverRelativeUrl: string, listName: string, fieldName: string) => Promise<any>;
    getListDataAsStream: (serverRelativeUrl: string, listName: string, itemID) => Promise<any>;
    breakItemRoleInheritance: (serverRelativeUrl: string, listName: string, itemID) => Promise<any>;
    getSearchResults: (queryText: string, selectedProperties: any[], queryTemplate?: string, refinementFilters?: string, sortList?: any[]) => Promise<any>;
    renderListDataStream: (serverRelativeUrl: string, listName: string, viewXML: string) => Promise<any>;
    getListInfo: (serverRelativeUrl: string, listName: string) => Promise<any>;
    getFileById: (fileArray: any[]) => Promise<any>;
    getCurrentUserInformation: (userEmail?:string) => Promise<any[]>;
    getAllListItemAttachments: (serverRelativeUrl: string, listName: string, itemID: number) => Promise<Map<string, any>>;
}