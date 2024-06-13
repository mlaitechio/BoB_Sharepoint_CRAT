import { IChoiceGroupOption } from "@fluentui/react";

export interface IFileViewerState {
    isPanelOpen: boolean;
    initialPreviewFileUrl: string;
    allFiles: any[];
    choiceGroup: IChoiceGroupOption[];
    selectedFile: string;
    isAllowedToUpdate: boolean;
    showLoading: boolean;
}