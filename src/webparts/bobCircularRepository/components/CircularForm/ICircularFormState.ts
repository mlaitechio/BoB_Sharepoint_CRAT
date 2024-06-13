import { IAttachments, ICircularListItem } from "../../Models/IModel";

export interface ICircularFormState {
    circularListItem?: ICircularListItem;
    currentCircularListItemValue?: ICircularListItem;
    expiryDate?: any;
    lblCircularType?: string;
    lblCompliance?: string;
    isLimited?: boolean;
    issuedFor?: any[];
    classification?: any[];
    category?: any[];
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