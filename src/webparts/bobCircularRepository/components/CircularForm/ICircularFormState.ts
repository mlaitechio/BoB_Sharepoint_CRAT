import { IAttachments, ICircularConfig, ICircularListItem, ICommentsAuditLogs } from "../../Models/IModel";

export interface ICircularFormState {
    circularListItem?: ICircularListItem;
    currentCircularListItemValue?: ICircularListItem;
    auditListItem?: ICommentsAuditLogs;
    isRequesterMaker?: boolean;
    configuration?: ICircularConfig[];
    selectedCommentSection?: any;
    expiryDate?: any;
    showSubmitDialog?: boolean;
    submittedStatus?: string;
    isDuplicateCircular?: any;
    isExpiryDateDisabled?: boolean;
    openSupportingDocument?: boolean;
    selectedSupportingCirculars?: any[];
    openSupportingCircularFile?: boolean;
    supportingDocLinkItem?: any
    isFormInValid?: boolean;
    isFileSizeAlert?: boolean;
    isFileTypeAlert?: boolean;
    alertTitle?: string;
    alertMessage?: string;
    isDeleteCircularFile?: boolean;
    /**
    |--------------------------------------------------
    | Supporting Documents & SOP Documents State
    |--------------------------------------------------
    */
    supportingDocAttachmentColl?: any[];
    supportingDocUploads?: Map<string, any>;
    sopUploads?: Map<string, any>;
    sopAttachmentColl?: any[];


    documentPreviewURL?: string;
    attachedFile?: any;
    selectedFileName?: string;
    currentItemID?: any;
    lblCircularType?: string;
    lblCompliance?: string;
    isLimited?: boolean;
    issuedFor?: any[];
    classification?: any[];
    category?: any[];
    templates?: any[];
    selectedTemplate?: string;
    templateFiles?: any[];
    emailTemplates?: any[];
    isBack?: boolean;
    isDelete?: boolean;
    isMaker?: boolean;
    isChecker?: boolean;
    isCompliance?: boolean;
    isLoading?: boolean;
    isSuccess?: boolean;
    isNewForm?: boolean;
    isEditForm?: boolean;
    currentPage?: string;
    comments?: Map<string, any[]>


}