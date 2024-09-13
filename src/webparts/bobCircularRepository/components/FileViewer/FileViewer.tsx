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
import { Text } from '@microsoft/sp-core-library'
import { rgb } from 'pdf-lib';




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
            if (isMigrated == Constants.lblNo || isMigrated == "") {
                fileArray = listItem.Attachments.Attachments.filter(val => val.FileName == circularFileName);
                file = fileArray.filter(val => val.FileName == circularFileName)[0];
            }


            await services.getAllListItemAttachments(serverRelativeUrl, Constants.circularList, parseInt(listItem.ID)).then((fileMetadata) => {
                let attachment: any[] = [];

                if (fileMetadata.size > 0) {
                    fileMetadata.forEach(async (value, key) => {
                        if (key == circularFileName && (isMigrated == Constants.lblNo || isMigrated == "")) {
                            attachment.push({
                                "name": key,
                                "content": value
                            });
                        }
                        else if (isMigrated == Constants.lblYes) {
                            attachment.push({
                                "name": key,
                                "content": value
                            });
                        }
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

                    this.setState({ fileContent: attachment.length > 0 ? attachment[0].content : null }, () => {
                        this.setState({
                            isPanelOpen: true,
                            initialPreviewFileUrl: allFiles && allFiles.length > 0 ? this.previewURL(allFiles[0]) : "",
                            allFiles: allFiles,
                            choiceGroup: choiceGroup,
                            selectedFile: circularFileName,
                            isAllowedToUpdate: isUpdateAllowed
                        }, () => {
                            const { fileContent } = this.state;
                            if (fileContent == null) {
                                this.props.documentLoaded()
                            }
                        })
                    })

                }
            }).catch((error) => {
                console.log(error);
                this.props.documentLoaded()
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
        const { published, expired, archived, infoPDFText, warninglimitedPDFText, warningUnlimitedPDFText } = Constants;
        let providerValue = this.context;
        const { responsiveMode, context, userInformation } = providerValue as IBobCircularRepositoryProps;
        const waterMarkText = userInformation?.employeeId ? userInformation?.employeeId + ` \n \n \n` + context.pageContext.user.displayName : context.pageContext.user.displayName;  //context.pageContext.user.displayName;
        let isMobileMode = responsiveMode == 0 || responsiveMode == 1 || responsiveMode == 2;
        let informationColumn = isMobileMode ? `${styles.column12}` : `${styles.column12}`;
        let filePreviewColumn = isMobileMode ? `${styles.column12}` : `${styles.column12}`
        let references = [];
        let hidePreviewColor = initialPreviewFileUrl.includes('.ppt') ? `#444444` : initialPreviewFileUrl.includes('.xls') ? `#217346` : `white`;

        const circularStatus = listItem.CircularStatus;
        const limited = listItem.Classification == Constants.limited;
        const expiryDate = listItem.Expiry != null ? (listItem.Expiry as string).split('T')[0] : "";
        const archivalDate = listItem.ArchivalDate != null ? (listItem.ArchivalDate as string).split('T')[0] : ""
        const currentDate = circularStatus == archived && limited ? new Date(expiryDate) : circularStatus == archived ? new Date(archivalDate) : new Date();
        const month = (currentDate.getMonth() + 1 < 10 ? '0' : '') + (currentDate.getMonth() + 1);
        const day = (currentDate.getDate() < 10 ? '0' : '') + currentDate.getDate();
        const year = currentDate.getFullYear().toString();
        const formatDate = `${day}/${month}/${year}`;

        const footerText = circularStatus == published ? Text.format(infoPDFText, formatDate) :
            circularStatus == archived && limited ? Text.format(warninglimitedPDFText, formatDate) : Text.format(warningUnlimitedPDFText, formatDate);
        const footerTextColor = circularStatus == published ? rgb(0.02, 0.02, 0.02) : circularStatus == archived ? rgb(0.86, 0.09, 0.26) : rgb(0.86, 0.09, 0.26);

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
                    {listItem != undefined &&
                        <>
                            <div className={`${styles.row}`}>
                                <div className={`${informationColumn}`}>

                                    {/* {showLoading && this.workingOnIt()} */}

                                </div>
                                <div className={`${filePreviewColumn} `} style={{ minHeight: "100vh" }}>

                                    {choiceGroup && choiceGroup.length > 0 && <>
                                        <div className={`${styles.row}`} >
                                            <div className={`${!isMobileMode ? `${styles.column7} ${styles.fontSizeFileName}` : `${styles.column6} ${styles.mobileFontFileName}`}`}>

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
                                                <div className={`${!isMobileMode ? `${styles.column12} ${styles.fontSizeFileName}` : `${styles.column12} ${styles.mobileFontFileName}`}`}
                                                    style={{ top: 3, background: hidePreviewColor, minHeight: 40, opacity: 1, width: "99.5%" }}>
                                                    {`${selectedFile}`}
                                                </div>
                                            </div>
                                        }
                                        {/* && allFiles.length > 0 && fileContent != null */}
                                        {
                                            initialPreviewFileUrl != "" && fileContent != null &&
                                            <MyPdfViewer
                                                context={context}
                                                pdfFilePath={allFiles[0].FileURL}
                                                currentSelectedFileContent={fileContent}
                                                watermarkText={`${waterMarkText}`}
                                                footerText={footerText}
                                                footerTextColor={footerTextColor}
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

                                    {choiceGroup.length == 0 || fileContent == null &&
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

    private onDismissPanel = (ev?: React.SyntheticEvent<HTMLElement> | KeyboardEvent) => {
        this.setState({ isPanelOpen: false }, () => {
            this.props.onClose()
        })
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