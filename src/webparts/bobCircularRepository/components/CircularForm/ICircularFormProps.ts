import { ICircularListItem } from "../../Models/IModel";

export interface ICircularFormProps {
    onGoBack?: (currentPage?: string) => void;
    displayMode?: string;
    editFormItem?: ICircularListItem;
    
}