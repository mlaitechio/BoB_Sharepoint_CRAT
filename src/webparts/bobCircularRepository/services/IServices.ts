export interface IServices {
    getPagedListItems: (serverRelativeUrl: string, listName: string, selectColumns: string, filterString: string, expandColumns: string, orderByColum: string, asc?: boolean) => Promise<any>;
    getLargeListItems: (serverRelativeUrl: string, listName: string, selectedColumn, expandColumns) => Promise<any[]>;
    addListItemAttachments(serverRelativeUrl: string, listName: string, itemID: number, fileMetadata: any[]): Promise<any>
    recycleListItemAttachments: (serverRelativeUrl: string, listName: string, itemID: number, files: Map<string, any>) => Promise<any>;
    readListItemAttachment: (serverRelativeUrl: string, listName: string, itemId: number) => Promise<any[]>
    updateItemBatch: (serverRelativeUrl, listName, itemIDs: any[], items: any[], departmentMapping: any) => Promise<any[]>
    updateItem: (serverRelativeUrl: string, listName: string, itemID: number, metadataValues: any, etag?: any) => Promise<any>;
    createItem: (serverRelativeUrl: string, listName: string, metadataValues: any) => Promise<boolean | any>;
    deleteListItem: (serverRelativeUrl: string, listName: string, itemId: number) => Promise<any>;
    checkIfUserBelongToGroup: (groupName: any, userEmail: string) => Promise<any>;
    readFieldValues: (serverRelativeUrl: string, listName: string, fieldName: string) => Promise<any>;
    getListDataAsStream: (serverRelativeUrl: string, listName: string, itemID) => Promise<any>;
    breakItemRoleInheritance: (serverRelativeUrl: string, listName: string, itemID) => Promise<any>;
    getSearchResults: (queryText: string, selectedProperties: any[], queryTemplate?: string, refinementFilters?: string, sortList?: any[], startRow?: any, rowLimit?: any) => Promise<any>;
    renderListDataStream: (serverRelativeUrl: string, listName: string, listItemIDs: any[]) => Promise<any>;
    filterLargeListItem: (serverRelativeUrl: string, listName: string, filterString: string) => Promise<any>;
    getListInfo: (serverRelativeUrl: string, listName: string) => Promise<any>;
    getFileById: (fileArray: any[]) => Promise<any>;
    getCurrentUserInformation: (userEmail: string, selectedColumns: string) => Promise<any[]>;
    getListItemById: (serverRelativeUrl: string, listName: string, itemID: number) => Promise<any>;
    addListItemAttachmentAsBuffer: (listName: string, serverRelativeUrl: string, itemID: number, fileName: string, buffer: any) => Promise<any>;
    getAllUsersInformationFromGroup: (groupName: string) => Promise<any>;
    getAllFiles: (folderServerRelativeUrl: string) => Promise<any[]>;
    getFileContent: (fileServerRelativeUrl: string) => Promise<any>;
    getLatestItemId: (serverRelativeUrl: string, listName: string) => Promise<any>
    getAllListItemAttachments: (serverRelativeUrl: string, listName: string, itemID: number) => Promise<Map<string, any>>;
    deleteListItemAttachment: (serverRelativeUrl: string, listName: string, itemID: number, fileName: string) => Promise<any>;
    addFileToListItem: (serverRelativeUrl: string, listName: string, itemID: number, fileArray: any[]) => Promise<any>;
    convertDocxToPDF: (serverRelativeUrl: string, listName: string, itemID: number, fileName: string) => Promise<any>;
    updateMultipleListItem: (serverRelativeUrl: string, listName: string, itemID: any[], metadata: any) => Promise<any[]>;
    getSupportingDocuments: (queryText: string, selectedProperties: any[], queryTemplate?: string, refinementFilters?: string, sortList?: any[]) => Promise<any>;
    sendEmail: (emailAddress: string[], ccEmailAddress: string[], subject: string, body: any) => Promise<any>;
}