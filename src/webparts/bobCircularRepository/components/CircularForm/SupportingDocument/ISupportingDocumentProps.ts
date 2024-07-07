import { IBobCircularRepositoryProps } from "../../IBobCircularRepositoryProps";

export interface ISupportingDocumentProps {
    department?: string;
    selectedSupportingCirculars?: any[];
    providerValue?: IBobCircularRepositoryProps;
    onDismiss?: (supportingCirculars?: any[]) => void;
    completeLoading?: () => void;
}
