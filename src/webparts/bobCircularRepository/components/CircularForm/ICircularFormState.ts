import { IAttachments, ICircularListItem } from "../../Models/IModel";

export interface ICircularFormState {
    circularListItem?: ICircularListItem;
    currentCircularListItemValue?: ICircularListItem;
    expiryDate?: any;
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
    sopUploads?: Map<string, any>;
    sopAttachmentColl?: any[];
    documentPreviewURL?: string;
    attachedFile?: any;
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
    isBack?: boolean;
    isDelete?: boolean;
    isMaker?: boolean;
    isChecker?: boolean;
    isCompliance?: boolean;
    isLoading?: boolean;
    isSuccess?: boolean;
    isNewForm?: boolean;
    isEditForm?: boolean;


}