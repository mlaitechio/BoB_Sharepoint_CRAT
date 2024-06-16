import { IChoiceGroupOption } from "@fluentui/react";

export interface IFileViewerState {
    isPanelOpen: boolean;
    initialPreviewFileUrl: string;
    allFiles: any[];
    fileContent: any;
    choiceGroup: IChoiceGroupOption[];
    selectedFile: string;
    isAllowedToUpdate: boolean;
    showLoading: boolean;
}