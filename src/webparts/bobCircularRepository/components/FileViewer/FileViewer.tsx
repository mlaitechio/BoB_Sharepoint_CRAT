import {
    ChoiceGroup,
    DirectionalHint, DocumentCard, DocumentCardActions,
    DocumentCardPreview, Icon, IDocumentCardPreviewImage,
    ImageFit, Label, Panel, PanelType, TooltipHost, ActionButton, IconButton, Image, Link, DefaultButton, DialogType, SpinnerSize, mergeStyleSets, FontSizes, FontWeights, getTheme, IImageProps, IChoiceGroup,
    DialogContent
} from '@fluentui/react';

import { Dialog, DialogBody, DialogSurface, Spinner } from "@fluentui/react-components";
import { Attach16Filled } from '@fluentui/react-icons'
import * as React from 'react'
import { Constants } from '../../Constants/Constants'
import { IFileViewerProps } from './IFileViewerProps'
import { IFileViewerState } from './IFileViewerState';
import styles from '../BobCircularRepository.module.scss';
import { IAttachmentFile } from '../../Models/IModel';
import utilities from '../../utilities/utilities';
import { IChoiceGroupOption } from '@fluentui/react';
import { PrimaryButton } from 'office-ui-fabric-react';
import { DataContext } from '../../DataContext/DataContext';
import { IBobCircularRepositoryProps } from '../IBobCircularRepositoryProps';
import JSZip from "jszip";
import { saveAs } from 'file-saver';
import MyPdfViewer from '../PDFViewer/PDFViewer';




const fileInfo = {
    FileURL: ``,
    FileName: ``,
    fileType: ``,
    ID: null,
    UniqueID: ``,

}

const theme = getTheme();

export default class FileViewer extends React.Component<IFileViewerProps, IFileViewerState> {

    static contextType = DataContext;
    context!: React.ContextType<typeof DataContext>;

    private _utilities: utilities;
    private contentStyles = mergeStyleSets({
        container: {
            display: "flex",
            flexFlow: "column nowrap",
            alignItems: "stretch",
            width: "30%"
        },
        header: [
            // tslint:disable-next-line:deprecation
            theme.fonts.xLargePlus,
            {
                flex: "1 1 auto",
                borderTop: `4px solid ${theme.palette.themePrimary}`,
                color: theme.palette.neutralPrimary,
                display: "flex",
                fontSize: FontSizes.xLarge,
                alignItems: "center",
                fontWeight: FontWeights.semibold,
                padding: "12px 12px 14px 24px"
            }
        ],
        body: {
            flex: "4 4 auto",
            padding: "0 24px 24px 24px",
            overflowY: "hidden",
            selectors: {
                p: {
                    margin: "14px 0"
                },
                "p:first-child": {
                    marginTop: 0
                },
                "p:last-child": {
                    marginBottom: 0
                }
            }
        }
    });

    private imageProps: IImageProps = {
        src: 'http://placehold.it/800x300',
        imageFit: ImageFit.center,
        width: 350,
        height: 150,
        onLoad: ev => console.log('image loaded', ev)
    };

    constructor(props) {
        super(props)

        this.state = {
            isPanelOpen: false,
            initialPreviewFileUrl: ``,
            allFiles: [],
            choiceGroup: [],
            selectedFile: ``,
            isAllowedToUpdate: false,
            showLoading: true,
            fileContent: null
        }

        this._utilities = new utilities();
    }

    public async componentDidMount() {

        const { listItem } = this.props;

        let providerValue = this.context;
        const { services, context, serverRelativeUrl } = providerValue as IBobCircularRepositoryProps;
        let publisherID = context?.pageContext?.user?.email?.split('@')[0] ?? ``;
        let isUpdateAllowed = false;

        if (listItem?.Attachments != undefined
            && listItem.Attachments?.Attachments?.length > 0) {

            let file = listItem.Attachments.Attachments[0];
            let fileArray = listItem.Attachments.Attachments;
            let choiceGroup: any[] = [];

            /**
            |--------------------------------------------------
            | Custom Logic
            |--------------------------------------------------
            */

            let circularFileName = listItem.CircularNumber.replace(/:/g, "_") + `.pdf`;

            let isMigrated = listItem?.IsMigrated ?? `Yes`;
            if (isMigrated == Constants.lblNo) {
                fileArray = listItem.Attachments.Attachments.filter(val => val.FileName == circularFileName);
                file = fileArray.filter(val => val.FileName == circularFileName)[0];
            }


            await services.getAllListItemAttachments(serverRelativeUrl, Constants.circularList, parseInt(listItem.ID)).then((fileMetadata) => {
                let attachment: any[] = [];

                if (fileMetadata.size > 0) {
                    fileMetadata.forEach(async (value, key) => {
                        attachment.push({
                            "name": key,
                            "content": value
                        });
                    });

                    let allFiles = fileArray.map((value, index) => {

                        let fileObject = {
                            name: value.FileName,
                            FileName: value.FileName,
                            ServerRelativeUrl: ``,
                            FileType: value.FileName.substring(value.FileName.lastIndexOf(".") + 1, value.FileName.length),
                            UniqueID: value.AttachmentId,
                            ID: value.AttachmentId,
                            FileURL: `${listItem.Attachments.UrlPrefix}${value.FileName}`
                        }

                        choiceGroup.push({
                            key: value.FileName,
                            text: value.FileName,
                            //text: value.FileName.length > 45 ? `${value.FileName.substring(0, 40)}` : value.FileName,
                            imageSrc: `${this.loadPreviewAttachment(fileObject).previewImageSrc}`,
                            styles: {
                                labelWrapper: {
                                    height: 50,
                                    font: 11.5,
                                    fontWeight: 600,
                                    maxWidth: 120
                                }
                            },
                            selectedImageSrc: `${this.loadPreviewAttachment(fileObject).previewImageSrc}`,
                            previewURL: `${this.props.listItem.Attachments.UrlPrefix}${value.FileName}`,
                            index: index
                        }
                        )

                        return fileObject
                    });

                    this.setState({ fileContent: attachment[0].content }, () => {
                        this.setState({
                            isPanelOpen: true,
                            initialPreviewFileUrl: this.previewURL(allFiles[0]),
                            allFiles: allFiles,
                            choiceGroup: choiceGroup,
                            selectedFile: choiceGroup[0].key,
                            isAllowedToUpdate: isUpdateAllowed
                        })
                    })

                }
            }).catch((error) => {
                console.log(error)
            })

        }
        else {
            this.setState({
                isPanelOpen: true
            }, () => {
                this.props.documentLoaded()
            })
        }

    }

    public render() {
        const { listItem } = this.props;
        const { initialPreviewFileUrl, fileContent } = this.state;
        const isLoading = initialPreviewFileUrl == null && fileContent == null;
        return (
            <div>
                {/* {isLoading && this.workingOnIt()} */}
                {listItem != undefined && this.openInfoPanel()}
            </div>
        )
    }


    private openInfoPanel = (): JSX.Element => {

        const { isPanelOpen, initialPreviewFileUrl, allFiles, fileContent, choiceGroup, selectedFile, isAllowedToUpdate, showLoading } = this.state;
        const { listItem } = this.props;
        let providerValue = this.context;
        const { responsiveMode, context, userInformation } = providerValue as IBobCircularRepositoryProps;
        const waterMarkText = userInformation?.employeeId ?? context.pageContext.user.displayName;  //context.pageContext.user.displayName;
        let isMobileMode = responsiveMode == 0 || responsiveMode == 1 || responsiveMode == 2;
        let informationColumn = isMobileMode ? `${styles.column12}` : `${styles.column12}`;
        let filePreviewColumn = isMobileMode ? `${styles.column12}` : `${styles.column12}`
        let references = [];
        let hidePreviewColor = initialPreviewFileUrl.includes('.ppt') ? `#444444` : initialPreviewFileUrl.includes('.xls') ? `#217346` : `white`;

        console.log(listItem)

        let infoPanelJSX = <>
            {
                <Panel
                    isOpen={isPanelOpen}
                    isLightDismiss={true}
                    onDismiss={this.onDismissPanel}
                    type={PanelType.custom}

                    closeButtonAriaLabel="Close"
                    headerText={`${this.props.listItem.Subject}`}
                    styles={{
                        commands: { background: "white" },
                        headerText: {
                            fontSize: "1.3em", fontWeight: "600",
                            marginBlockStart: "0.83em", marginBlockEnd: "0.83em",
                            color: "black", fontFamily: 'Roboto'
                        },
                        main: { background: "white" },
                        content: { paddingBottom: 0 },
                        navigation: {
                            borderBottom: "1px solid #ccc",
                            selectors: {
                                ".ms-Button": { color: "black" },
                                ".ms-Button:hover": { color: "black" }
                            }
                        }
                    }}
                >
                    {this.props.listItem != undefined &&
                        <>
                            <div className={`${styles.row}`}>
                                <div className={`${informationColumn}`}>

                                    {/* {showLoading && this.workingOnIt()} */}

                                </div>
                                <div className={`${filePreviewColumn} `} style={{ minHeight: "100vh" }}>

                                    {choiceGroup && choiceGroup.length > 0 && <>
                                        <div className={`${styles.row}`}>
                                            <div className={`${!isMobileMode ? `${styles.column7} ${styles.fontSizeFileName}` : `${styles.column6} ${styles.mobileFontFileName}`}`}>
                                                {`${selectedFile}`}
                                            </div>
                                            <div className={`${!isMobileMode ? styles.column5 : styles.column6} ${styles.textAlignEnd}`}>

                                                {/* <IconButton iconProps={{
                                                    iconName: "Copy",
                                                    styles: { root: { fontSize: 24 } }
                                                }}
                                                    alt='Copy'
                                                    title='Copy File Path'
                                                    onClick={() => { this.copyToClipboard(selectedFile) }}
                                                ></IconButton> */}

                                            </div>
                                        </div>

                                        {initialPreviewFileUrl != "" &&
                                            <div className={`${styles.row}`}>


                                                <div className={`${!isMobileMode ? styles.column12 : styles.column12}`}
                                                    style={{ top: 3, background: hidePreviewColor, minHeight: 34, opacity: 1, width: "98.5%" }}>

                                                </div>
                                            </div>
                                        }
                                        {
                                            initialPreviewFileUrl != "" && fileContent != null &&
                                            <MyPdfViewer
                                                context={context}
                                                pdfFilePath={allFiles[0].FileURL}
                                                currentSelectedFileContent={fileContent}
                                                watermarkText={`${waterMarkText}`}
                                                documentLoaded={() => { this.props.documentLoaded() }}
                                            />
                                        }
                                        {/* {initialPreviewFileUrl != "" &&
                                            <iframe
                                                src={initialPreviewFileUrl ?? ``}
                                                style={{
                                                    minHeight: 800,
                                                    height: 180,
                                                    width: "100%",
                                                    border: 0
                                                }} role="presentation" tabIndex={-1}></iframe>
                                        } */}


                                    </>
                                    }
                                    {choiceGroup.length == 0 &&
                                        <div className={`${styles.row}`}>
                                            <div className={`${styles.column9} ${styles.headerTextAlignCenter}`}>
                                                {this.labelControl(``, `No Circular Content Found`, false)}

                                            </div>
                                            <div className={`${styles.column3}`}>
                                                {/* <IconButton iconProps={{ iconName: "PageLink" }} title='Copy Item Link' onClick={() => { this.copyItemLink(listItem.ID) }}>
                                                </IconButton> */}
                                            </div>
                                            <div className={`${styles.column2}`}>

                                            </div>
                                            <div className={`${styles.column12}`}>
                                                <Image src={require(`../../assets/emptyFile.gif`)} styles={{
                                                    root: {
                                                        display: "flex",
                                                        justifyContent: "center"
                                                    }
                                                }}></Image>
                                            </div>
                                            <div className={`${styles.column2}`}>

                                            </div>
                                        </div>
                                    }
                                </div>

                            </div>

                        </>
                    }
                </Panel>

            }
        </>
        return infoPanelJSX;

    }

    private onAttachmentClick = (index, file) => {
        const { allFiles } = this.state
        let fileObject = allFiles[index]
        this.setState({ selectedFile: file, initialPreviewFileUrl: this.previewURL(fileObject) })
    }

    private workingOnIt = (): JSX.Element => {

        let submitDialogJSX = <>

            <Dialog modalType="alert" defaultOpen={true}>
                <DialogSurface style={{ maxWidth: 250 }}>
                    <DialogBody style={{ display: "block" }}>
                        <DialogContent>
                            {<Spinner labelPosition="below" label={"Working on It..."}></Spinner>}
                        </DialogContent>
                    </DialogBody>
                </DialogSurface>
            </Dialog>

        </>;
        return submitDialogJSX;
    }

    private convertDate = (dateText) => {
        let formattedDate = `N/A`;
        if (dateText != "") {
            const date = new Date(dateText);
            const day = date.getDate().toString().padStart(2, '0');
            const month = (date.getMonth() + 1).toString().padStart(2, '0');
            const year = date.getFullYear();
            formattedDate = `${day}/${month}/${year}`;
        }
        return formattedDate;
    }

    private onUpdateclick = (itemID) => {
        this.props.onUpdate(itemID)
    }

    private onChoicePreviewChange = (ev?: React.FormEvent<HTMLElement | HTMLInputElement>, option?: any) => {
        const { allFiles } = this.state
        let fileObject = allFiles[option.index]
        this.setState({ selectedFile: option.key, initialPreviewFileUrl: this.previewURL(fileObject) })
    }

    private showDocumentMetadata = (): JSX.Element => {
        let documentMetadataJSX = <>

        </>

        return documentMetadataJSX;
    }

    private onDismissPanel = (ev?: React.SyntheticEvent<HTMLElement> | KeyboardEvent) => {
        this.setState({ isPanelOpen: false }, () => {
            this.props.onClose()
        })
    }

    private onDownLoadAllClick = async (attachments: any[]) => {
        let providerValue = this.context;
        const { serverRelativeUrl } = providerValue as IBobCircularRepositoryProps;


        attachments.map((currentFile) => {
            const fileURL = `${serverRelativeUrl}/_layouts/download.aspx?SourceUrl=` + currentFile?.previewURL;
            window.open(fileURL, "_blank")
        })



    }

    private onDownloadAllZip = async (itemID, subject) => {
        let providerValue = this.context;
        const { context, serverRelativeUrl, services } = providerValue as IBobCircularRepositoryProps;

        this.setState({ showLoading: true }, async () => {
            await services.getAllListItemAttachments(serverRelativeUrl, Constants.circularList, itemID).then((val) => {

                const zip = new JSZip();

                val.forEach((fileBuffer, fileName) => {
                    zip.file(fileName, fileBuffer)
                })
                this.setState({ showLoading: false }, () => {
                    zip.generateAsync({ type: "blob" })
                        .then(function (content) {
                            // see FileSaver.js
                            saveAs(content, `${subject}.zip`);
                        });
                })

            }).catch((error) => {
                this.setState({ showLoading: false });
            })
        })


    }

    private onViewDownloadClick = (fileInfo, isView) => {
        const { allFiles } = this.state;
        const currentFileInfo = allFiles.filter((value) => {
            return value.FileName == fileInfo
        });

        if (isView && currentFileInfo.length > 0) {
            /**
                * View File
                */
            const fileURL = this.previewURL(currentFileInfo[0])//currentFileInfo[0].FileURL + "?web=1"
            window.open(fileURL, "_blank")
        }
        else if (currentFileInfo.length > 0) {
            /**
                * Download File
                */
            let providerValue = this.context;
            const { serverRelativeUrl } = providerValue as IBobCircularRepositoryProps;

            const fileURL = `${serverRelativeUrl}/_layouts/download.aspx?SourceUrl=` + currentFileInfo[0].FileURL;
            window.open(fileURL, "_blank")
        }

    }


    private previewURL = (fileInfo) => {
        let finalUrl: string = "";
        const { context } = this.props
        if (Constants.imageTypes.filter((element, index, array) => { return element.toLocaleLowerCase() == fileInfo.FileType.toLocaleLowerCase(); }).length > 0)
            finalUrl = encodeURI(fileInfo.FileURL);
        //finalUrl=`${window.location.origin + context.pageContext.legacyPageContext.webServerRelativeUrl}/_layouts/15/getpreview.ashx.aspx?resolution=6&path=${fileInfo.FileURL}`
        else if (Constants.videoTypes.filter((element, index, array) => { return element.toLocaleLowerCase() == fileInfo.FileType.toLocaleLowerCase(); }).length > 0)
            finalUrl = `${window.location.origin + context.pageContext.legacyPageContext.webServerRelativeUrl}/_layouts/15/embed.aspx?UniqueId=${fileInfo.UniqueID}&client_id=FileViewerWebPart&embed={"af":false,"id":"${fileInfo.ID}","o":"${window.location.origin}","p":1,"z":"width"}`;
        else if (Constants.pdfFileType.filter((element, index, array) => { return element.toLocaleLowerCase() == fileInfo.FileType.toLocaleLowerCase(); }).length > 0)
            finalUrl = `${window.location.origin + context.pageContext.legacyPageContext.webServerRelativeUrl}/_layouts/15/WopiFrame.aspx?sourcedoc=${fileInfo.UniqueID}&action=interactivepreview`;
        else if (Constants.officeFileTypes.filter((element, index, array) => { return element.toLocaleLowerCase() == fileInfo.FileType.toLocaleLowerCase(); }).length > 0)
            finalUrl = `${window.location.origin}/:w:/r${context.pageContext.legacyPageContext.webServerRelativeUrl}/_layouts/15/Doc.aspx?sourcedoc=${fileInfo.UniqueID}&file=${encodeURI(fileInfo.FileName)}&action=interactivepreview&mobileredirect=true`;
        else if (Constants.otherFileTypes.filter((element, index, array) => { return element.toLocaleLowerCase() == fileInfo.FileType.toLocaleLowerCase(); }).length > 0)
            finalUrl = encodeURI(fileInfo.FileURL);

        return (finalUrl);
    }

    private createDocumentPreview = (attachmentCollection: any[], ItemAttachmentUrl: string): JSX.Element => {

        let previewJSXDocument =
            <>
                {
                    attachmentCollection.map((file) => {

                        let fileObject: any = {
                            name: file.FileName,
                            FileName: file.FileName,
                            ServerRelativeUrl: ``,
                            FileType: file.FileName.split('.')[1],
                            UniqueID: file.AttachmentId,
                            ID: file.AttachmentId,
                            FileURL: `${ItemAttachmentUrl}${file.FileName}`

                        }
                        let filePreview = this.loadPreviewAttachment(fileObject);
                        // let previewImages = { };
                        // previewImages[filePreview.name] = filePreview;

                        return <>
                            <div key={file.FileName} className={`${styles.documentCardFilePreview} ${styles.column3}`}>
                                <TooltipHost
                                    content={file.FileName}
                                    calloutProps={{ gapSpace: 0, isBeakVisible: true }}
                                    closeDelay={200}
                                    directionalHint={DirectionalHint.rightCenter}>

                                    <DocumentCard
                                        onClick={this.openDocument.bind(this, fileObject)}
                                        className={styles.documentCard}>
                                        <DocumentCardPreview previewImages={[filePreview]} styles={{ root: { cursor: "pointer" } }} />
                                        <Label className={styles.fileLabel}>{file.FileName}</Label>
                                    </DocumentCard>
                                </TooltipHost>
                            </div>
                        </>
                    })
                }

            </>

        return previewJSXDocument;
    }

    private loadPreviewAttachment = (file: IAttachmentFile): IDocumentCardPreviewImage => {
        let previewImageUrl = this._utilities.GetAttachmentImageUrl(file);
        return {
            name: file.name,
            previewImageSrc: previewImageUrl,
            iconSrc: '',
            imageFit: ImageFit.center,
            width: 187,
            height: 130
        };
    }

    private copyItemLink = (id: string) => {
        const { context } = this.props
        let itemURL = `${window.location.origin + context.pageContext.legacyPageContext.webServerRelativeUrl}?itemID=${id}`;

        navigator.clipboard.writeText(itemURL)
            .then(() => {
                alert('Document link copied.Please store it in your editor for further use in forms as references')
            })
            .catch((error) => {
                console.error('Error copying document to clipboard:', error);
            });
    }

    private copyToClipboard = (fileInfo) => {
        const { allFiles } = this.state;
        const currentFileInfo = allFiles.filter((value) => {
            return value.FileName == fileInfo
        });

        const fileURL = this.previewURL(currentFileInfo[0]);

        navigator.clipboard.writeText(fileURL)
            .then(() => {
                alert('File Link Copied .Please store it in your editor for further use in forms as references')
            })
            .catch((error) => {
                console.error('Error copying text to clipboard:', error);
            });
    }

    private openDocument = (file: any, ev?: React.SyntheticEvent<HTMLElement>): void => {
        this.setState({ initialPreviewFileUrl: this.previewURL(file) })
    }

    public labelControl = (labelClassName: string, labelName: string, isRequired: boolean, requiredColor?: any, toolTipContent?: string): JSX.Element => {
        return <>
            {/* <div className={`${styles.mediumLargeColumntest}`}> */}
            {
                isRequired ?
                    <Label
                        className={labelClassName}
                        required

                        styles={{
                            root: {
                                display: "flex",
                                fontFamily: 'Roboto',
                                fontSize: 12.5,
                                selectors: {
                                    ":after": {
                                        color: (requiredColor != undefined || requiredColor != "") ? requiredColor : `red`,
                                        fontSize: 13.5,
                                        content: " *  ",
                                        paddingRight: 12,
                                        paddingLeft: 4
                                    }
                                }
                            }
                        }}>
                        {labelName} {
                            toolTipContent != undefined && toolTipContent != "" &&
                            <TooltipHost content={toolTipContent} id={labelClassName} styles={{ root: { display: "flex" } }}
                                directionalHint={DirectionalHint.topLeftEdge}>
                                <Icon iconName='Info'
                                    styles={{
                                        root: {
                                            float: "right",
                                            position: "absolute",
                                            cursor: "pointer",
                                            padding: "0px 0px 3px 15px"
                                        }
                                    }}>

                                </Icon>
                            </TooltipHost>
                        }</Label>
                    : <Label className={labelClassName} styles={{
                        root: {
                            display: "flex",
                            fontFamily: 'Roboto',
                            fontSize: 12.5,
                            selectors: {
                                ":after": {
                                    color: (requiredColor != undefined || requiredColor != "") ? requiredColor : `red`,
                                    fontSize: 13.5,
                                    content: "   ",
                                    paddingRight: 12,
                                    paddingLeft: 4
                                }
                            }
                        }
                    }}>
                        {labelName}
                        {
                            toolTipContent != undefined && toolTipContent != "" &&
                            <TooltipHost content={toolTipContent} id={labelClassName} styles={{ root: { display: "flex" } }}
                                directionalHint={DirectionalHint.topLeftEdge}>
                                <Icon iconName='Info'
                                    styles={{
                                        root: {
                                            float: "right",
                                            position: "absolute",
                                            cursor: "pointer",
                                            padding: "0px 0px 3px 5px"
                                        }
                                    }}>

                                </Icon>
                            </TooltipHost>
                        }
                    </Label>
            }
            {/* </div> */}

        </>
    }



}