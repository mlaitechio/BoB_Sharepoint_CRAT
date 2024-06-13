export interface ICircularSearchProps{
    isSites?: boolean;
    goBack?: (message?: string) => void;
    onAddNewItem?: () => void;
    onSubjectLinkClick?: (id?: any) => void;
    onUpdateItem?: (itemID: any) => void;
}