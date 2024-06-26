import { IDocumentCardPreviewImage } from "@fluentui/react";

export interface IUserInformation {
    Id: number;
    EMail: string;
    Title: string;
}

export interface Category {
    text: string;
    id: string;
    key: string;
    name: string;
}

export interface IServiceCard {
    Title: string;
    Description: string;
    IsActive: string;
    IconName: string;
    Services: string;
    SortOrder: any;
    Route: string;
    IsHyperLink: any;
}

export interface ILeftNavigation {
    Title: string;
    ID: any;
    IsHeader: string;
    IsActive: string;
    IsHyperLink: string;
    Link: IHyperLink;
    IconName: string;
    Route: string;
    IconColor: string;
    FontSize: string;
    OrderNumber: any;
    Parent: IParent;
}

export interface IHyperLink {
    Url: string;
    Description: string;
}

export interface IParent {
    Title: string;
    Id: any;
}

export interface IListItem {
    Title: string;
    ID: any;
    Subject: string;
    PublishedDate: string;
    PublisherEmailID: string;
    ModifiedDate: string;
}

export interface ICheckBoxCollection {
    checked: string | boolean,
    value: string,
    refinableString: string
}

export interface ICircularListItem {
    ID?: string;
    Id?: string;
    CreatedBy?: string;
    Created?: string;
    Category?: string;
    CircularContent?: string;
    CircularCreationDate?: string;
    CircularFAQ?: string;
    CircularNumber?: string;
    CircularSOP?: string;
    CircularStatus?: string;
    CircularType?: string;
    Classification?: string;
    CommentsChecker?: string;
    CommentsCompliance?: string;
    CommentsMaker?: string;
    Compliance?: string;
    DeleteComments?: string;
    DeleteDate?: string;
    DeleteRemarks?: string;
    Department?: string;
    Expiry?: string | Date;
    Gist?: string;
    IssuedFor?: string;
    Keywords?: string;
    IsMigrated?: string;
    Attachments?: IAttachments;
    MasterCircularMapping?: string;
    Author?: IUserInformation;
    MigratedDepartment?: string;
    MigratedDocPath?: string;
    MigratedIssuedFor?: string;
    MigratedOriginator?: string;
    MigratedRefNumber?: string;
    MigratedSubFileNo?: string;
    Modified?: string;
    PublishedDate?: string;
    SubFileCode?: string;
    Subject?: string;
    SubmittedDate?: string;
}

export interface IAttachments {
    Attachments: IAttachmentsInfo[];
    UrlPrefix: string;
}

export interface IAttachmentsInfo {
    FileName: string;
    AttachmentId: string;
    FileTypeProgId: string;
    RedirectUrl: string;
}

export interface ILookUp {
    Id: any;
    Title: string;
}

export class IAttachmentFile {
    FileName: string;
    ServerRelativeUrl: string;
    name?: string;
}

export interface IPreviewImageCollection {
    [fileName: string]: IDocumentCardPreviewImage;
}