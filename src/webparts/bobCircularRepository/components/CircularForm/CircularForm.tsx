import * as React from 'react';
import styles from '../BobCircularRepository.module.scss';
import { ICircularFormProps } from './ICircularFormProps';
import { ICircularFormState } from './ICircularFormState';
import {
    Dropdown, Field, Image,
    Tooltip,
    Input, Label, Persona, Option, SelectionEvents,
    OptionOnSelectData, Divider, Button, Switch,
    SwitchOnChangeData, Textarea, InputOnChangeData,
    TextareaOnChangeData, DialogSurface, DialogBody,
    DialogTitle, DialogActions, Spinner, DialogTrigger,
    Toast,
    Toaster, useToastController,
    ToastTitle,
    useId,
    MessageBar,
    MessageBarBody,
    MessageBarTitle,
    Link,
    MessageBarIntent,
} from '@fluentui/react-components';
import { Constants } from '../../Constants/Constants';
import { Text } from '@microsoft/sp-core-library';
import { DataContext } from '../../DataContext/DataContext';
import { DatePicker } from '@fluentui/react-datepicker-compat';
import { Add16Filled, ArrowCounterclockwiseRegular, ArrowLeftFilled, ArrowUpload16Regular, Attach16Filled, CalendarRegular, ChevronDownRegular, ChevronUpRegular, Delete16Regular, DeleteRegular, InfoRegular, OpenRegular } from '@fluentui/react-icons';
import { IBobCircularRepositoryProps } from '../IBobCircularRepositoryProps';
import { Dialog } from '@fluentui/react-components';
import { AnimationClassNames, DialogContent } from '@fluentui/react';
import { IADProperties, IAttachmentFile, IAttachmentsInfo, ICircularListItem } from '../../Models/IModel';
import { IFileInfo } from '@pnp/sp/files';
import SupportingDocument from './SupportingDocument/SupportingDocument';
import FileViewer from '../FileViewer/FileViewer';
import { error } from 'pdf-lib';
import { IAttachmentAddResult } from '@pnp/sp/attachments';




export default class CircularForm extends React.Component<ICircularFormProps, ICircularFormState> {

    static contextType = DataContext;
    context!: React.ContextType<typeof DataContext>;
    private sopFileAttachments: Map<string, any>;
    private editSOPFileAttachments: Map<string, any>;
    private deleteSOPFileAttachments: Map<string, any>;
    private sopFileInput;

    public constructor(props) {
        super(props)

        this.state = {
            circularListItem: {
                CircularNumber: ``,
                CircularStatus: `New`,
                CircularType: `${Constants.limited}`,
                CircularTemplate: ``,
                SubFileCode: ``,
                IssuedFor: ``,
                Category: ``,
                Classification: ``,
                Expiry: null,
                Subject: ``,
                Keywords: ``,
                Department: ``,
                Compliance: `No`,
                Gist: ``,
                CommentsMaker: ``,
                IsMigrated: `No`,
                CommentsChecker: ``,
                CommentsCompliance: ``,
                CircularContent: ``,
                CircularFAQ: ``,
                CircularSOP: ``,
                SubmittedDate: null,
                CircularCreationDate: null,
                SupportingDocuments: ``
            },
            selectedCommentSection: {
                isMakerSelected: false,
                isCheckerSelected: false,
                isComplianceSelected: false
            },
            auditListItem: {
                Title: ``,
                CommentsHistory: ``,
                CircularAuditStatus: ``,
                CircularMetadata: ``
            },
            currentCircularListItemValue: undefined,
            isRequesterMaker: true, // Initially when form loads this will be true for new form & it will change in edit/view form load 
            selectedSupportingCirculars: [],
            sopAttachmentColl: [],
            showSubmitDialog: false,
            submittedStatus: ``,
            alertTitle: `Validation Alert!`,
            alertMessage: `Please Input all fields marked as *`,
            lblCompliance: ``,
            lblCircularType: Constants.limited,
            issuedFor: [],
            category: [],
            classification: [],
            selectedTemplate: ``,
            templates: [],
            isDeleteCircularFile: false,
            isBack: false,
            isDelete: false,
            isLoading: false,
            isSuccess: false,
            isNewForm: true,
            isEditForm: false,
            expiryDate: null,
            attachedFile: null,
            isFormInValid: false,
            documentPreviewURL: ``,
            selectedFileName: ``,
            openSupportingDocument: false,
            currentItemID: -1,
            isExpiryDateDisabled: false,
            openSupportingCircularFile: false,
            supportingDocLinkItem: undefined,
            isFileSizeAlert: false,
            isFileTypeAlert: false,
            isDuplicateCircular: false,
            comments: new Map<string, any[]>(),
            configuration: []

        }


        this.sopFileInput = React.createRef();
        this.sopFileAttachments = new Map<string, any>();
        this.editSOPFileAttachments = new Map<string, any>();
        this.deleteSOPFileAttachments = new Map<string, any>();
    }


    public async componentDidMount() {

        let providerValue = this.context;
        const { services, serverRelativeUrl, context, userInformation } = providerValue as IBobCircularRepositoryProps;
        const { circularListItem } = this.state
        const { displayMode, editFormItem } = this.props

        //context.pageContext.user.email;
        this.setState({ isLoading: true }, async () => {

            await this.fieldValues(Constants.colIssuedFor).then((val) => {
                this.setState({ issuedFor: val?.Choices ?? [] })
            }).catch((error) => {
                console.log(error);
                this.setState({ isLoading: false })
            });

            await this.fieldValues(Constants.colCategory).then((val) => {
                this.setState({ category: val?.Choices ?? [] })
            }).catch((error) => {
                console.log(error);
                this.setState({ isLoading: false })
            });

            await this.fieldValues(Constants.colClassification).then((val) => {
                this.setState({ classification: val?.Choices ?? [], isLoading: false })
            }).catch((error) => {
                console.log(error);
                this.setState({ isLoading: false })
            });

            await services.getAllFiles(`${serverRelativeUrl}/${Constants.templateFolder}`).then((files: any[]) => {
                let templates = files.map((file) => {
                    return file.Name.split('.')[0]
                });

                let templateFiles = files.map((file) => {
                    file.templateName = file.Name.split('.')[0];
                    return file;
                })

                this.setState({ templates, templateFiles })

            }).catch((error) => {
                console.log(error);
                this.setState({ isLoading: false })
            });

            await services.getPagedListItems(serverRelativeUrl, Constants.configurationList, Constants.configSelectColumns, ``, ``, `ID`).then((configVal) => {
                this.setState({ configuration: configVal })
            }).catch((error) => {
                this.setState({ isLoading: false })
            })

            /**
            |--------------------------------------------------
            | When form is in new mode
            |--------------------------------------------------
            */
            if (displayMode == Constants.lblNew) {
                circularListItem.Department = userInformation?.department ?? ``;
                await services.getLatestItemId(serverRelativeUrl, Constants.circularList).then((itemID) => {
                    const { circularListItem } = this.state;
                    circularListItem.CircularNumber = parseInt(itemID + 1).toString();
                    this.setState({ circularListItem, isLoading: false })
                }).catch((error) => {
                    console.log(`Latest Item ID` + error);
                    this.setState({ isLoading: false })
                })

                this.setState({ circularListItem })
            }
            else if (displayMode == Constants.lblEditCircular) {
                await services.getListDataAsStream(serverRelativeUrl, Constants.circularList, parseInt(editFormItem.ID)).then((listItem) => {
                    listItem.ListData.ID = parseInt(editFormItem.ID);
                    this.onEditViewFormLoad(listItem?.ListData ?? [])
                }).catch((error) => {
                    console.log(error)
                })

            }
            else if (displayMode == Constants.lblViewCircular) {
                await services.getListDataAsStream(serverRelativeUrl, Constants.circularList, parseInt(editFormItem.ID)).then((listItem) => {
                    listItem.ListData.ID = parseInt(editFormItem.ID);
                    this.onEditViewFormLoad(listItem?.ListData ?? [])
                }).catch((error) => {
                    console.log(error);
                })
                //this.onEditViewFormLoad(editFormItem)
            }

        })


    }

    /**
    |--------------------------------------------------
    | When form gets edited 
    |--------------------------------------------------
    */

    private onEditViewFormLoad = (editFormItem) => {

        let providerValue = this.context;
        const { context, services } = providerValue as IBobCircularRepositoryProps;
        const { displayMode } = this.props

        let requesterMail = editFormItem?.Author?.split('#')[4].replace(',', '');
        let currentUserEmail = context.pageContext.user.email;
        let isUserRequesterMaker = currentUserEmail == requesterMail;

        let editCircularItem = {
            CircularCreationDate: editFormItem.CircularCreationDate && editFormItem.CircularCreationDate != "" ? new Date(editFormItem?.CircularCreationDate)?.toISOString() : null,
            Subject: editFormItem?.Subject ?? ``,
            CircularNumber: editFormItem?.CircularNumber ?? ``,
            IssuedFor: editFormItem?.IssuedFor ?? ``,
            Category: editFormItem?.Category ?? ``,
            CircularStatus: editFormItem?.CircularStatus ?? ``,
            Classification: editFormItem?.Classification ?? ``,
            SubFileCode: editFormItem?.SubFileCode ?? ``,
            Keywords: editFormItem?.Keywords ?? ``,
            Expiry: editFormItem?.Expiry && editFormItem?.Expiry != "" ? new Date(editFormItem?.Expiry)?.toISOString() : null,
            CircularType: editFormItem?.CircularType ?? ``,
            Department: editFormItem?.Department ?? ``,
            Compliance: editFormItem?.Compliance ?? ``,
            CircularTemplate: editFormItem?.CircularTemplate ?? ``,
            SupportingDocuments: editFormItem?.SupportingDocuments ?? ``,
            Gist: editFormItem?.Gist ?? ``,
            CircularFAQ: editFormItem?.CircularFAQ ?? ``,
            CommentsMaker: ``, //editFormItem?.CommentsMaker ?? 
            CommentsChecker: ``,//editFormItem?.CommentsChecker ??
            CommentsCompliance: ``,//editFormItem?.CommentsCompliance ??
            MakerCommentsHistory: editFormItem?.MakerCommentsHistory ?? ``,
            ComplianceCommentsHistory: editFormItem?.ComplianceCommentsHistory ?? ``,
            CheckerCommentsHistory: editFormItem?.CheckerCommentsHistory ?? ``

        } as ICircularListItem

        this.setState({
            circularListItem: editCircularItem,
            currentCircularListItemValue: editFormItem,
            currentItemID: parseInt(editFormItem.ID),
            lblCompliance: editCircularItem.Compliance == Constants.lblComplianceYes ? Constants.lblCompliance : ``,
            isNewForm: false,
            isRequesterMaker: isUserRequesterMaker
            //isEditForm: displayMode == Constants.lblEditCircular

        }, async () => {
            let supportingCirculars = [];
            if (editFormItem?.SupportingDocuments) {
                supportingCirculars = JSON.parse(editFormItem.SupportingDocuments);
            }

            let circularFileName = editFormItem.CircularNumber.split(':').join('_') + `.docx`;



            let allSopFiles = this.allSopFiles(editFormItem, circularFileName);
            let sopUploads = new Map<string, any[]>();
            if (allSopFiles.length > 0) {
                await services.getFileById(allSopFiles).then((fileInfo) => {
                    sopUploads = this.setEditSOPFiles(fileInfo)
                }).catch((error) => {
                    console.log(error)
                })
            }

            let circularFileContent = editFormItem?.Attachments?.Attachments?.filter((val) => {
                return val.FileName == circularFileName
            });

            this.setState({
                selectedSupportingCirculars: supportingCirculars,
                attachedFile: circularFileContent?.length > 0 ? circularFileContent[0] : null,
                sopUploads: sopUploads
            }, () => {

                const { attachedFile, sopUploads } = this.state;
                let attachmentColl = [];
                let i = 0;
                sopUploads.forEach((value, key) => {
                    value.index = i;
                    attachmentColl.push(value);
                    i++;
                });

                if (attachmentColl.length > 0) {
                    this.setState({
                        sopAttachmentColl: attachmentColl
                    });
                }
                else {
                    this.setState({
                        sopAttachmentColl: attachmentColl
                    })
                }

                editCircularItem.CircularTemplate = attachedFile == null ? `` : editFormItem?.CircularTemplate;
                let comments = [{
                    FieldName: Constants.lblCommentsMaker,
                    History: editFormItem.MakerCommentsHistory
                }, {
                    FieldName: Constants.lblCommentsCompliance,
                    History: editFormItem.ComplianceCommentsHistory
                }, {
                    FieldName: Constants.lblCommentsChecker,
                    History: editFormItem.CheckerCommentsHistory
                }];

                let commentHistory = this.commentsJSON(comments);

                this.setState({
                    documentPreviewURL: this.circularContentPreviewURL(context),
                    selectedTemplate: attachedFile != null ? editCircularItem.CircularTemplate : ``,
                    selectedFileName: attachedFile != null ? attachedFile.FileName : ``,
                    circularListItem: editCircularItem,
                    expiryDate: editFormItem?.Expiry != null && editFormItem.Expiry != "" ?
                        new Date(editFormItem.Expiry) : null,
                    comments: commentHistory,
                    isLoading: false,

                })
            })

        })
    }


    private allSopFiles = (editFormItem, circularFileName) => {
        let allSopFiles = [];

        if (editFormItem?.Attachments && editFormItem.Attachments.Attachments.length > 0) {
            let allFiles: IAttachmentsInfo[] = editFormItem.Attachments.Attachments as IAttachmentsInfo[];
            allSopFiles = allFiles.filter((val) => {
                return val.FileName != circularFileName
            })
        }

        return allSopFiles;
    }

    private setEditSOPFiles = (fileResults: any[]) => {
        let fileAttachments = new Map<string, any>();
        //let attachedSize = 0;
        if (fileResults.length > 0) {
            fileResults.map((file) => {
                fileAttachments.set(file.Name, {
                    name: file.Name,
                    FileName: file.Name,
                    ServerRelativeUrl: file.ServerRelativeUrl,
                    UniqueId: file?.UniqueId ?? ``,
                    size: parseInt(file.Length),
                    isFileEdit: true
                })
                //attachedSize += parseInt(file.Length)
            });
        }

        this.sopFileAttachments = fileAttachments;
        this.editSOPFileAttachments = fileAttachments;

        return fileAttachments;
    }

    private circularContentPreviewURL = (context) => {
        let documentPreviewURL = ``;
        const { displayMode } = this.props;
        const { attachedFile } = this.state
        let action = displayMode == Constants.lblEditCircular ? `edit` : `interactivepreview`;
        if (attachedFile != null && attachedFile.FileName.indexOf('.docx') > -1) {
            documentPreviewURL = `${window.location.origin}/:w:/r${context.pageContext.legacyPageContext.webServerRelativeUrl}/_layouts/15/Doc.aspx?sourcedoc=`;
            documentPreviewURL += `${attachedFile.AttachmentId}&file=${encodeURI(attachedFile.FileName)}&action=${action}&mobileredirect=true`;
        };

        return documentPreviewURL;
    }

    private deleteSOPUploadedFiles = (fileName) => {
        if (this.sopFileAttachments.has(fileName)) {

            if (this.editSOPFileAttachments.has(fileName)) {
                this.deleteSOPFileAttachments.set(fileName, this.editSOPFileAttachments.get(fileName))
            }

            this.sopFileAttachments.delete(fileName);

            this.setState({
                sopUploads: this.sopFileAttachments
            }, () => {

                let { sopUploads } = this.state;
                let attachmentColl = [];
                let i = 0;
                sopUploads.forEach((value, key) => {
                    value.index = i;
                    attachmentColl.push(value);
                    i++;
                });

                if (attachmentColl.length > 0) {
                    this.setState({
                        sopAttachmentColl: attachmentColl
                    });
                }
                else {
                    this.setState({
                        sopAttachmentColl: attachmentColl
                    })
                }

            })

        }
    }


    private createSOPPreviewURL = (file: IAttachmentFile) => {
        let documentPreviewURL = ``
        let providerValue = this.context;
        const { context } = providerValue as IBobCircularRepositoryProps;

        if (file.FileName.indexOf(`.pdf`) > -1 && file.UniqueId != ``) {
            documentPreviewURL = `${window.location.origin + context.pageContext.legacyPageContext.webServerRelativeUrl}/_layouts/15/WopiFrame.aspx?sourcedoc={${file.UniqueId}}&action=interactivepreview`;
        }
        else if (file.FileName.indexOf(`.docx`) > -1 && file.UniqueId != ``) {
            documentPreviewURL = `${window.location.origin}/:w:/r${context.pageContext.legacyPageContext.webServerRelativeUrl}/_layouts/15/Doc.aspx?sourcedoc=`;
            documentPreviewURL += `{${file.UniqueId}}&file=${encodeURI(file.FileName)}&action=interactivepreview&mobileredirect=true`;

        }

        this.setState({ documentPreviewURL: documentPreviewURL, selectedFileName: file.FileName })
    }

    private commentsJSON = (comments: any[]) => {
        let commentsMap = new Map<string, any[]>();

        comments.map((comment) => {
            if (comment.History) {
                commentsMap.set(comment.FieldName, JSON.parse(comment.History).reverse())
            }
            else {
                commentsMap.set(comment.FieldName, [])
            }

        })

        return commentsMap;
    }


    private fieldValues = async (colName: string): Promise<any> => {

        let providerValue = this.context;
        const { services, serverRelativeUrl } = providerValue as IBobCircularRepositoryProps;

        let fieldPromise = await services.readFieldValues(serverRelativeUrl, Constants.circularList, colName).
            then((val) => {
                return Promise.resolve(val);
            }).catch((error) => {
                return Promise.reject(error);
            });

        return fieldPromise;
    }

    public render() {

        const { displayMode } = this.props

        const { isBack, isDelete, isLoading, isSuccess,
            documentPreviewURL, attachedFile,
            isFormInValid, openSupportingDocument,
            isDeleteCircularFile, isFileSizeAlert, isFileTypeAlert, isDuplicateCircular,
            showSubmitDialog, submittedStatus, selectedFileName,
            openSupportingCircularFile } = this.state;

        let showAlert = (isDelete || isBack);

        let title = isFormInValid || isFileSizeAlert || isFileTypeAlert || isDuplicateCircular ?
            Constants.validationAlertTitle :
            isDeleteCircularFile ? `${Constants.deleteCircularTitle}` : ``;

        let message = isFormInValid ? Constants.validationAlertMessage :
            isDeleteCircularFile ? `${Constants.deleteCircularMessage}` :
                isFileSizeAlert ? Constants.validationAlertMessageFileSize :
                    isFileTypeAlert ? Constants.validationAlertMessageFileType :
                        isDuplicateCircular ? Constants.validationCircularNumber : ``;


        return (
            <>

                <div className={`${styles.row}`}>
                    <div className={`${styles.column1} ${styles.headerBackgroundColor} `}>
                        <Button icon={<ArrowLeftFilled />}
                            onClick={this.onBtnClick.bind(this, Constants.goBack)}
                            title="Go Back"
                            appearance="transparent"
                            className={`${styles.formHeader}`}></Button>
                    </div>
                    <div className={`${styles.column11} ${styles.headerBackgroundColor} ${styles['text-center']}`}>
                        <Label className={`${styles.formHeader}`}>
                            {Text.format(Constants.headerCircularUpload,
                                `${displayMode == Constants.lblNew ? Constants.lblNew :
                                    displayMode == Constants.lblEditCircular ? "Edit" : "View"}`)}
                        </Label>

                    </div>
                    {/* <div className={`${styles.column1} ${styles.headerBackgroundColor} `}>
                        <Button icon={<DeleteRegular />}
                            onClick={this.onBtnClick.bind(this, Constants.delete)}
                            style={{ float: "right" }}
                            title="Delete Circular" appearance="transparent"
                            className={`${styles.formHeader}`}></Button>
                    </div> */}

                </div>
                <div className={`${styles.row}`}>
                    <div className={`${styles.column6}`}>
                        <div className={`${styles.row}`} style={{ padding: 15, borderRight: "1px solid #80808036" }}>
                            {this.infoHeader()}
                            {this.formSection()}

                        </div >
                    </div>
                    <div className={`${styles.column6} `} style={{ minHeight: "100vh" }}>
                        <div className={`${styles.row}`} style={{ padding: 15 }}>
                            <div className={`${styles.column12}`}>
                                <Label className={`${styles.formLabel}`} >
                                    {selectedFileName != `` && `${selectedFileName}`}
                                </Label>
                            </div>
                            <div className={`${styles.column12}`} style={{ display: "flex", justifyContent: "center", alignItems: "center" }}>
                                {/* <Label className={`${styles.formLabel}`} >Attachment preview section</Label> */}
                                {documentPreviewURL != "" && <iframe
                                    src={documentPreviewURL ?? ``}
                                    style={{
                                        minHeight: 800,
                                        height: 1080,
                                        width: "100%",
                                        border: 0
                                    }} role="presentation" tabIndex={-1}></iframe>}
                            </div>
                        </div>
                        {documentPreviewURL != "" &&
                            displayMode != Constants.lblNew &&
                            displayMode != Constants.lblEditCircular &&
                            < div className={`${styles.row}`}>
                                <div className={`${styles.column10}`}></div>

                                <div className={`${styles.column2}`}
                                    style={{ marginTop: -32, background: "white", minHeight: 21, opacity: 1, marginLeft: -11 }}>

                                </div>
                            </div>
                        }
                    </div>
                </div >
                <div className={`${styles.row} ${styles.formFieldMarginTop} ${styles['text-center']}`} style={{ borderTop: "1px solid lightgoldenrodyellow" }}>
                    {this.saveCancelBtn()}
                </div>
                {/* <div className={`${styles.row} ${styles.formFieldMarginTop} ${styles['text-center']}`}>
                    {this.messageBarControl(`error`)}
                </div> */}

                {
                    showAlert &&
                    this.deleteBackDialogControl(showAlert)
                }

                {
                    (isFormInValid || isFileSizeAlert || isFileTypeAlert || isDuplicateCircular) &&
                    this.alertControl((isFormInValid || isFileSizeAlert || isFileTypeAlert || isDuplicateCircular),
                        title, undefined, message, this.alertButton())
                }
                {
                    openSupportingDocument && this.filterPanelSupportingDocument()
                }
                {
                    isDeleteCircularFile && this.alertControl(isDeleteCircularFile, title, undefined, message, this.alertButton())
                }
                {openSupportingCircularFile && this.supportingDocumentFileViewerPanel()}
                {
                    (isLoading || isSuccess) && this.workingOnIt()
                }
                {
                    showSubmitDialog && this.submitDialog(showSubmitDialog, submittedStatus)
                }

            </>
        )
    }

    private infoHeader = (): JSX.Element => {

        const { circularListItem, currentCircularListItemValue } = this.state;
        const { displayMode } = this.props
        let providerValue = this.context;
        const { context } = providerValue as IBobCircularRepositoryProps;
        let author = currentCircularListItemValue?.Author;
        let requester = displayMode == Constants.lblNew ? context.pageContext.user.displayName :
            currentCircularListItemValue?.Author?.split('#')[1].replace(',', '');
        let circularCreationDate = displayMode == Constants.lblNew ? this.onFormatDate(new Date()) :
            this.onFormatDate(new Date(currentCircularListItemValue?.Created));


        let infoSectionJSX = <>

            <div className={`${styles.column12}`}>
                <div className={`${styles.row} ${styles.formRequestInfo}`}>
                    <div className={`${styles.column5}`}>
                        <Label className={`${styles.formLabel}`}
                            title={requester}
                            style={{
                                maxWidth: 275,
                                textOverflow: "ellipsis",
                                overflow: "hidden",
                                display: "inline-block"
                            }}>
                            Requester : {requester}
                        </Label>
                    </div>
                    <div className={`${styles.column4}`}>
                        <Label className={`${styles.formLabel}`}
                            title={circularListItem.CircularStatus}
                            style={{
                                maxWidth: 210,
                                textOverflow: "ellipsis",
                                overflow: "hidden",
                                display: "inline-block"
                            }}>Status : {circularListItem.CircularStatus}</Label>
                    </div>
                    <div className={`${styles.column2}`}>
                        <Label className={`${styles.formLabel}`}>Creation Date : {circularCreationDate}</Label>
                    </div>
                </div>
            </div>


        </>
        return infoSectionJSX;

    }

    private formSection = (): JSX.Element => {

        const { circularListItem, expiryDate, lblCompliance,
            issuedFor, category, templates, selectedTemplate, attachedFile,
            selectedCommentSection,
            classification, configuration, selectedSupportingCirculars, isRequesterMaker } = this.state;
        let providerValue = this.context;
        const { context, isUserChecker, isUserMaker, isUserCompliance } = providerValue as IBobCircularRepositoryProps;
        const { displayMode, currentPage } = this.props
        const { keywordsToolTip, categoryToolTip } = Constants
        let circularStatus = circularListItem.CircularStatus;
        let showCommentHistory = circularStatus != Constants.lblNew && circularStatus != Constants.draft;
        const { configVal } = Constants;
        let maxSupportingDocument = configuration.filter(val => val.Title == configVal.SupportingDocuments)[0]?.Limit ?? 20;
        let maxGistWord = configuration.filter(val => val.Title == configVal.GistMaxWord)[0]?.Limit ?? 5;

        /**
        |--------------------------------------------------
        | Disable all html controls if form is in view mode
        |--------------------------------------------------
        */
        let disableControl = displayMode == Constants.lblViewCircular;

        /**
        |--------------------------------------------------
        | Switch button Check or Uncheck condition
        |--------------------------------------------------
        */
        let isTypeChecked = (circularListItem.CircularType == Constants.unlimited);
        let isTypeDisabled = circularListItem.Classification == Constants.lblMaster;
        let isComplianceChecked = circularListItem.Compliance == Constants.lblComplianceYes

        /**
        |--------------------------------------------------
        | show Maker Checker Compliance Comment Box
        |--------------------------------------------------
        */
        let showMakerCommentBox = (circularStatus == Constants.cmmtCompliance || circularStatus == Constants.cmmtChecker) && currentPage == Constants.makerGroup;
        let showComplianceCommentBox = circularStatus == Constants.sbmtCompliance && currentPage == Constants.complianceGroup;
        let showCheckerCommentBox = circularStatus == Constants.sbmtChecker && currentPage == Constants.checkerGroup;

        let disableCircularNumber = circularListItem.CircularNumber != "" && circularListItem.CircularStatus != Constants.lblNew;

        let formSectionJSX = <>
            <div className={`${styles.column12}`} >
                <div className={`${styles.row} ${styles.formFieldMarginTop}`}>
                    <div className={`${styles.column12}`}>
                        {this.textAreaControl(`${Constants.subject}`, true, `${circularListItem.Subject}`, disableControl, `Field cannot be empty`)}
                    </div>
                    {/* <div className={`${styles.column6}`}>
                        {this.avatarControl(`${Constants.circularInitator}`, `${context.pageContext.user.displayName}`)}
                    </div> */}


                </div>
                <Divider appearance="subtle" ></Divider>
                <div className={`${styles.row}  ${styles.formFieldMarginTop}`}>
                    <div className={`${styles.column6}`}>
                        {this.textFieldControl(`${Constants.circularNumber}`, true, `${circularListItem.CircularNumber}`, disableCircularNumber, `Field cannot be empty`)}

                    </div>
                    <div className={`${styles.column6}`}>
                        {this.dropDownControl(`${Constants.issuedFor}`, true, `${circularListItem.IssuedFor}`, issuedFor, disableControl, `Field cannot be empty`)}
                    </div>
                </div>
                <Divider appearance="subtle" ></Divider>
                <div className={`${styles.row}  ${styles.formFieldMarginTop}`}>
                    <div className={`${styles.column6}`}>
                        {this.dropDownControl(`${Constants.category}`, true, `${circularListItem.Category}`, category, disableControl, `Field cannot be empty`, true, categoryToolTip)}
                    </div>
                    <div className={`${styles.column6}`}>
                        {this.dropDownControl(`${Constants.classification}`, true, `${circularListItem.Classification}`, classification, disableControl, `Field cannot be empty`)}
                    </div>
                </div>
                <Divider appearance="subtle" ></Divider>

                {/* <Divider></Divider> */}
                <div className={`${styles.row} ${styles.formFieldMarginTop}`}>
                    <div className={`${styles.column6}`}>
                        {this.textFieldControl(`${Constants.subFileNo}`, false, `${circularListItem.SubFileCode}`, disableControl, ``)}
                    </div>
                    <div className={`${styles.column6}`}>
                        {this.textFieldControl(`${Constants.keyWords}`, true, `${circularListItem.Keywords}`, disableControl, `Field cannot be empty`, ``, true, keywordsToolTip)}
                    </div>
                </div>
                <Divider appearance="subtle" ></Divider>

                <div className={`${styles.row}  ${styles.formFieldMarginTop}`}>
                    <div className={`${styles.column6}`}>
                        {this.datePickerControl(`${Constants.expiry}`, expiryDate, !isTypeChecked, (isTypeChecked || disableControl))}
                    </div>
                    <div className={`${styles.column6}`}>
                        {this.switchControl(`${Constants.type}`, false, `${circularListItem?.CircularType ?? ``}`, "vertical", isTypeChecked, (isTypeDisabled || disableControl))}
                    </div>
                </div>
                <Divider appearance="subtle" ></Divider>
                <div className={`${styles.row} ${styles.formFieldMarginTop}`}>
                    <div className={`${styles.column6}`}>
                        {this.textFieldControl(`${Constants.department}`, false, `${circularListItem.Department}`, disableControl)}
                    </div>
                    <div className={`${styles.column6}`}>
                        {this.switchControl(`${Constants.compliance}`, false, `${lblCompliance}`, "vertical", isComplianceChecked, disableControl)}
                    </div>
                </div>
                <Divider appearance="subtle" ></Divider>
                <div className={`${styles.row} ${styles.formFieldMarginTop}`}>
                    <div className={`${styles.column6}`}>
                        {this.dropDownControl(`${Constants.lblTemplate}`, false, `${selectedTemplate}`, templates, disableControl, `Field cannot be empty`, true, ["Select Template File"])}

                    </div>
                    <div className={`${styles.column6}`}>
                        <Field label={<Label className={`${styles.formLabel} ${styles.fieldTitle}`}>{`Circular Content`}</Label>}>
                            {attachedFile && this.attachmentLink(attachedFile)}
                        </Field>
                    </div>
                </div>
                <Divider appearance="subtle" ></Divider>

                <div className={`${styles.row} ${styles.formFieldMarginTop}`}>
                    {/* {Check if Logged in User is Maker & Get All Circulars from Makers Department in search} */}
                    <div className={`${styles.column6}`}>
                        <Label className={`${styles.formLabel} ${styles.fieldTitle}`}>{`${Constants.supportingDocument}`}</Label>
                    </div>
                    <div className={`${styles.column6}`}>
                        <Tooltip content={`Maximum ${maxSupportingDocument} supporting documents can be attached`}
                            relationship="description" positioning={"after"} withArrow={true}>
                            <Button appearance="primary" icon={<Add16Filled />}
                                disabled={disableControl}
                                style={{ width: "100%", padding: 5, cursor: "pointer" }}
                                onClick={() => {
                                    this.setState({ openSupportingDocument: true, isLoading: true })
                                }}
                                iconPosition="before">Click here to add supporting documents</Button>
                        </Tooltip>
                    </div>
                </div>
                <div className={`${styles.row} ${styles.formFieldMarginTop}`}>
                    <div className={`${styles.column12}`}>
                        <div className={`${styles.row}`}>
                            {selectedSupportingCirculars && selectedSupportingCirculars.length > 0 &&
                                selectedSupportingCirculars.map((listItem) => {

                                    let supportingCircularNumber = listItem?.CircularNumber ?? ``;

                                    return <>
                                        {supportingCircularNumber.length > 15 &&
                                            <Tooltip content={listItem.CircularNumber ?? ``} relationship="label" positioning={'after'} withArrow={true}>
                                                <div className={`${styles.column2}`}>

                                                    <Link
                                                        className={`${styles.colorLabel}`}
                                                        // title={listItem.CircularNumber ?? ``}
                                                        onClick={() => {
                                                            this.openSupportingCircularFile(listItem);
                                                        }}>{supportingCircularNumber}</Link>

                                                    {/* <Label>{listItem.CircularNumber ?? ``}</Label> */}

                                                </div>
                                            </Tooltip>
                                        }
                                        {supportingCircularNumber.length < 15 &&
                                            <div className={`${styles.column2}`}>
                                                <Link
                                                    className={`${styles.colorLabel}`}
                                                    // title={listItem.CircularNumber ?? ``}
                                                    onClick={() => {
                                                        this.openSupportingCircularFile(listItem);
                                                    }}>{supportingCircularNumber}</Link>

                                                {/* <Label>{listItem.CircularNumber ?? ``}</Label> */}

                                            </div>}
                                        <div className={`${styles.column1}`}>
                                            <Button
                                                disabled={disableControl}
                                                icon={<Delete16Regular />}
                                                appearance="transparent"
                                                onClick={() => { this.deleteSupportingCircular(listItem) }}></Button>
                                        </div>

                                    </>
                                })}
                        </div>
                    </div>
                    <div className={`${styles.column12}`}>
                        {/* <Label className={`${styles.formLabel}`}>{`(Max ${maxSupportingDocument} supporting documents can be attached)`}</Label> */}
                        {/* {this.messageBarControl("warning",`(Max ${maxSupportingDocument} supporting documents can be attached)`)} */}
                    </div>
                    <div>

                    </div>
                </div>
                <Divider appearance="subtle" ></Divider>

                <div className={`${styles.row} ${styles.formFieldMarginTop}`}>
                    <div className={`${styles.column12}`}>
                        {this.textAreaControl(`${Constants.gist}`, false, `${circularListItem.Gist}`, disableControl, ``,
                            `Maximum 2000 characters are allowed`)}
                    </div>
                </div>

                <Divider appearance="subtle" ></Divider>

                <div className={`${styles.row} ${styles.formFieldMarginTop}`}>
                    <div className={`${styles.column12}`}>
                        {this.textAreaControl(`${Constants.faqs}`, false, `${circularListItem.CircularFAQ}`, disableControl, ``, `Maximum 2000 characters are allowed`)}
                    </div>
                </div>

                <Divider appearance="subtle" ></Divider>
                {
                    <div className={`${styles.row} ${styles.formFieldMarginTop}`}>
                        <div className={`${styles.column12}`}>
                            {this.fileUploadControl(`${Constants.sop}`, this.sopFileInput, disableControl)}
                        </div>
                        <div className={`${styles.column6}`} style={{ padding: 10 }}>
                            {this.sopFilesControl(disableControl)}
                        </div>
                    </div>
                }

                {
                    showCommentHistory &&
                    this.commentsSection(`${Constants.lblCommentsMaker}`, selectedCommentSection.isMakerSelected)
                }

                {
                    circularListItem.Compliance == Constants.lblComplianceYes && showCommentHistory &&
                    this.commentsSection(`${Constants.lblCommentsCompliance}`, selectedCommentSection.isComplianceSelected)
                }

                {
                    showCommentHistory && this.commentsSection(`${Constants.lblCommentsChecker}`, selectedCommentSection.isCheckerSelected)
                }


                {isUserMaker && showMakerCommentBox && isRequesterMaker && <>
                    <div className={`${styles.row} ${styles.formFieldMarginTop}`}>
                        <div className={`${styles.column12}`}>
                            {this.textAreaControl(`${Constants.lblCommentsMaker}`, true, `${circularListItem.CommentsMaker}`)}
                        </div>
                    </div>
                    <Divider appearance="subtle" ></Divider>
                </>
                }

                {isUserCompliance && showComplianceCommentBox && !isRequesterMaker && <>
                    <div className={`${styles.row} ${styles.formFieldMarginTop}`}>
                        <div className={`${styles.column12}`}>
                            {this.textAreaControl(`${Constants.lblCommentsCompliance}`, true, `${circularListItem.CommentsCompliance}`)}
                        </div>
                    </div>
                    <Divider appearance="subtle" ></Divider>
                </>
                }

                {isUserChecker && showCheckerCommentBox && !isRequesterMaker && <>
                    <div className={`${styles.row} ${styles.formFieldMarginTop}`}>
                        <div className={`${styles.column12}`}>
                            {this.textAreaControl(`${Constants.lblCommentsChecker}`, true, `${circularListItem.CommentsChecker}`)}
                        </div>
                    </div>
                </>}

            </div>
        </>
        return formSectionJSX;
    }

    private openSupportingCircularFile = (listItem) => {
        let providerValue = this.context;
        const { services, serverRelativeUrl } = providerValue as IBobCircularRepositoryProps;

        this.setState({
            isLoading: true
        }, async () => {
            await services.getListDataAsStream(serverRelativeUrl, Constants.circularList, listItem.ID).then((result) => {
                result.ListData.ID = listItem.Id;
                this.setState({
                    openSupportingCircularFile: true,
                    supportingDocLinkItem: result.ListData,
                    isLoading: false
                })
            }).catch((error) => {
                this.setState({
                    isLoading: false,
                    supportingDocLinkItem: undefined,
                    openSupportingCircularFile: false
                })
                console.log(error)
            })
        })
    }

    private supportingDocumentFileViewerPanel = (): JSX.Element => {
        const { supportingDocLinkItem } = this.state
        let providerValue = this.context;
        const { context } = providerValue as IBobCircularRepositoryProps;
        let supportingFileViewer = <>
            <FileViewer listItem={supportingDocLinkItem}
                context={context}
                documentLoaded={() => { this.setState({ isLoading: false }) }}
                onClose={() => { this.setState({ openSupportingCircularFile: false }) }}
            />
        </>;

        return supportingFileViewer;
    }

    private saveCancelBtn = (): JSX.Element => {
        let providerValue = this.context;
        const { isUserMaker, isUserCompliance, isUserChecker } = providerValue as IBobCircularRepositoryProps;
        const { displayMode, currentPage } = this.props
        const { circularListItem, isRequesterMaker } = this.state
        let displayButton = (displayMode == Constants.lblNew || displayMode == Constants.lblEditCircular);
        let circularStatus = circularListItem.CircularStatus;




        /**
        |--------------------------------------------------
        | Hide or show Button based on form status
         1.User is Maker & Status is (New & Draft) -> (show Draft & Submit Button & Delete button )
         2.User is Compliance & status is (Submitted to Compliance & Submitted to checker)-> (Show Return to Maker Button )
         3.User is Checker & status is  (Submitted to Checker) ->(Show Return to Maker & Publish Button)
        |--------------------------------------------------
        */
        let showDraftClearSubmitBtn = isUserMaker &&
            (circularStatus == Constants.lblNew || circularStatus == Constants.draft ||
                circularStatus == Constants.cmmtCompliance || circularStatus == Constants.cmmtChecker);

        let showReturnToMakerBtn = circularStatus == Constants.sbmtCompliance || circularStatus == Constants.sbmtChecker;
        let showSbmtCheckerBtn = circularStatus == Constants.sbmtCompliance;
        let showPublishRejectButton = circularStatus == Constants.sbmtChecker;


        /**
        |--------------------------------------------------
        | Set Form status based on Submit Click
        |--------------------------------------------------
        */

        let submtStatus = isUserMaker && circularListItem.Compliance == Constants.lblYes ? Constants.sbmtCompliance : Constants.sbmtChecker;
        let returnStatus = circularListItem.CircularStatus == Constants.sbmtCompliance ? Constants.cmmtCompliance : Constants.cmmtChecker;
        let status = ""


        let saveCancelBtnJSX = <>
            {/* {showDraftClearSubmitBtn &&
                <Button appearance="primary" className={`${styles.formBtn}`}
                    onClick={this.clearAllFormFields}>Clear
                </Button>
            } */}

            {showDraftClearSubmitBtn && displayButton && isRequesterMaker &&
                <Button appearance="primary"
                    className={`${styles.formBtn}`}
                    onClick={this.saveForm.bind(this, Constants.draft)}>
                    Save as Draft
                </Button>
            }
            {showDraftClearSubmitBtn && displayButton && isRequesterMaker &&

                <Button appearance="primary"
                    className={`${styles.formBtn}`}
                    disabled={showReturnToMakerBtn || showSbmtCheckerBtn}
                    onClick={() => {
                        status = submtStatus;
                        this.setState({ submittedStatus: status, showSubmitDialog: true })
                    }}>
                    Submit
                </Button>
            }
            {(isUserCompliance || isUserChecker) && showReturnToMakerBtn && !isRequesterMaker
                && (currentPage == Constants.complianceGroup || currentPage == Constants.checkerGroup) &&
                <Button
                    appearance="primary"

                    onClick={() => {
                        status = returnStatus;
                        this.setState({ submittedStatus: status, showSubmitDialog: true })
                    }}
                    className={`${styles.formBtn}`}>
                    Return to maker
                </Button>
            }
            {
                isUserCompliance && showSbmtCheckerBtn && currentPage == Constants.complianceGroup && !isRequesterMaker
                && <Button appearance="primary"
                    onClick={() => {
                        status = Constants.sbmtChecker;
                        this.setState({ submittedStatus: status, showSubmitDialog: true })
                    }}
                    className={`${styles.formBtn}`}>
                    Submit to Checker
                </Button>

            }
            {isUserChecker && showPublishRejectButton && currentPage == Constants.checkerGroup && !isRequesterMaker
                && <Button appearance="primary"
                    onClick={() => {
                        status = Constants.published;
                        this.setState({ submittedStatus: status, showSubmitDialog: true })
                    }}
                    className={`${styles.formBtn}`}>
                    Publish
                </Button>
            }
            {isUserChecker && showPublishRejectButton && currentPage == Constants.checkerGroup && !isRequesterMaker
                && <Button appearance="primary"

                    onClick={() => {
                        status = Constants.archived;
                        this.setState({ submittedStatus: status, showSubmitDialog: true })
                    }}
                    className={`${styles.formBtn}`}>
                    Reject
                </Button>
            }

            {
                <Button
                    className={`${styles.formBtn}`}
                    onClick={this.onBtnClick.bind(this, Constants.goBack)}
                    appearance="secondary">
                    Cancel
                </Button>
            }


        </>;

        return saveCancelBtnJSX;
    }

    private submitDialog = (showDialog, currentStatus): JSX.Element => {
        let submitDialogJSX = <>
            <>
                <Dialog modalType="alert" defaultOpen={(showDialog)} >
                    <DialogSurface style={{ maxWidth: 480, padding: 14 }}>
                        <DialogBody style={{ display: "block" }}>
                            <DialogTitle style={{ fontFamily: "Roboto", marginBottom: 10 }}>{`${`Save Circular` ?? ``}`}</DialogTitle>
                            <DialogContent styles={{
                                header: { display: "none" },
                                inner: { padding: 0 },

                                innerContent: { fontFamily: "Roboto", marginBottom: 15, textAlign: "center" }
                            }}>
                                {`${`Are you sure you want to change status to ${currentStatus}?`}`}
                            </DialogContent>
                            <DialogActions style={{ justifyContent: "center" }}>
                                <div className={`${styles.row}`}>
                                    <div className={`${styles.column6}`}>
                                        <Button appearance="primary"
                                            onClick={() => {
                                                this.setState({ showSubmitDialog: false }, () => {
                                                    this.saveForm(currentStatus);
                                                })

                                            }}>Yes</Button>
                                    </div>
                                    <div className={`${styles.column6}`}>
                                        <Button appearance="secondary"
                                            onClick={() => {
                                                this.setState({ showSubmitDialog: false })
                                            }}>No</Button>
                                    </div>
                                </div>
                            </DialogActions>
                        </DialogBody>
                    </DialogSurface>
                </Dialog>
            </>
        </>;

        return submitDialogJSX
    }

    private onBtnClick = (labelName: string) => {
        switch (labelName) {
            case Constants.goBack: this.setState({ isBack: true, isDelete: false })
                break;
            case Constants.delete: this.setState({ isBack: false, isDelete: true })
                break;
            default:
                break;
        }
    }

    private textAreaControl = (labelName: string, isRequired: boolean, value: string, isDisabled?: boolean, errorMessage?: string, hintMessage?: string): JSX.Element => {
        let textAreaJSX = <>
            <Field label={<Label className={`${styles.formLabel} ${styles.fieldTitle}`}>{labelName}</Label>}
                required={isRequired}
                hint={hintMessage ?? ``}
                validationState={isRequired && value == "" ? "error" : "none"}
                validationMessage={isRequired && value == "" ? errorMessage : ``}  >
                <Textarea value={value} appearance="outline"
                    disabled={isDisabled}
                    root={{ className: `${styles.formLabel}` }}
                    resize="vertical" onChange={this.onTextAreaChange.bind(this, labelName)}></Textarea>
            </Field>
        </>;

        return textAreaJSX;
    }

    private onTextAreaChange = (labelName: string, ev: React.ChangeEvent<HTMLTextAreaElement>, data: TextareaOnChangeData) => {

        const { circularListItem, configuration, auditListItem } = this.state;
        const { configVal } = Constants;
        let wordLength = this.getWordsLength(data.value?.trim());
        let textValue = data.value?.replace(/[^a-zA-Z0-9.&,() ]/g, '')
        switch (labelName) {
            case Constants.subject:
                let subjectMaxWord = configuration.filter(val => val.Title == configVal.SubjectMaxWord)[0]?.Limit ?? 500;
                if (textValue.length <= subjectMaxWord && textValue.trim() != "") {
                    circularListItem.Subject = textValue;//data.value?.replace(/[^a-zA-Z0-9.&,() ]/g, '');
                    this.setState({ circularListItem });
                }
                else {
                    if (textValue == "") {
                        circularListItem.Subject = ``;
                        this.setState({ circularListItem })
                    }

                }
                break;
            case Constants.gist:
                let gistMaxWord = configuration.filter(val => val.Title == configVal.GistMaxWord)[0]?.Limit ?? 500;
                if (textValue.length <= gistMaxWord && textValue.trim() != "") {
                    circularListItem.Gist = textValue;//data.value?.replace(/[^a-zA-Z0-9.&,() ]/g, '');
                    this.setState({ circularListItem })
                }
                else {
                    if (textValue == "") {
                        circularListItem.Gist = ``;
                        this.setState({ circularListItem })
                    }
                    // circularListItem.Gist = data.value?.replace(/[^a-zA-Z0-9.&,() ]/g, '').trim();
                    // this.setState({ circularListItem })
                }
                break;
            case Constants.faqs:
                let faqMaxWord = configuration.filter(val => val.Title == configVal.FAQMaxWord)[0]?.Limit ?? 500;
                if (textValue.length <= faqMaxWord && textValue.trim() != "") {
                    circularListItem.CircularFAQ = textValue; //data.value?.replace(/[^a-zA-Z0-9.&,() ]/g, '');
                    this.setState({ circularListItem });
                }
                else {
                    if (textValue == "") {
                        circularListItem.CircularFAQ = ``;
                        this.setState({ circularListItem })
                    }
                    // circularListItem.CircularFAQ = data.value?.replace(/[^a-zA-Z0-9.&,() ]/g, '').trim();
                    // this.setState({ circularListItem });
                }
                break;
            case Constants.lblCommentsMaker:
                let commentsMakerMaxWord = configuration.filter(val => val.Title == configVal.MakerCommentsMaxWord)[0]?.Limit ?? 50;
                if (textValue.length <= commentsMakerMaxWord && textValue.trim() != "") {

                    circularListItem.CommentsMaker = textValue;//data.value.replace(/[^a-zA-Z0-9.&,() ]/g, '');
                    auditListItem.CommentsHistory = textValue;
                    this.setState({ circularListItem, auditListItem })
                }
                else {
                    if (textValue == "") {
                        circularListItem.CommentsMaker = ``;
                        this.setState({ circularListItem })
                    }
                    // circularListItem.CommentsMaker = data.value.replace(/[^a-zA-Z0-9.&,() ]/g, '').trim();
                    // this.setState({ circularListItem })
                }
                break;
            case Constants.lblCommentsChecker:
                let checkerMaxWord = configuration.filter(val => val.Title == configVal.CheckerCommentsMaxWord)[0]?.Limit ?? 50;
                if (textValue.length <= checkerMaxWord && textValue.trim() != "") {
                    circularListItem.CommentsChecker = textValue;//data.value.replace(/[^a-zA-Z0-9.&,() ]/g, '');
                    auditListItem.CommentsHistory = textValue;
                    this.setState({ circularListItem, auditListItem })
                }
                else {

                    if (textValue == "") {
                        circularListItem.CommentsChecker = ``;
                        this.setState({ circularListItem })
                    }
                    // circularListItem.CommentsChecker = data.value.replace(/[^a-zA-Z0-9.&,() ]/g, '').trim();
                    // this.setState({ circularListItem })
                }
                break;
            case Constants.lblCommentsCompliance:
                let complianceMaxWord = configuration.filter(val => val.Title == configVal.ComplianceCommentsMaxWord)[0]?.Limit ?? 50;
                if (textValue.length <= complianceMaxWord && textValue.trim() != "") {
                    circularListItem.CommentsCompliance = textValue; //data.value.replace(/[^a-zA-Z0-9.&,() ]/g, '');
                    auditListItem.CommentsHistory = textValue;
                    this.setState({ circularListItem, auditListItem })
                }
                else {
                    if (textValue == "") {
                        circularListItem.CommentsCompliance = ``;
                        this.setState({ circularListItem })
                    }
                    // circularListItem.CommentsCompliance = data.value.replace(/[^a-zA-Z0-9.&,() ]/g, '').trim();
                    // this.setState({ circularListItem })
                }
                break;

            default:
                break;
        }
    }



    private getWordsLength = (text: string): number => {
        text.replace(/(<([^>]+)>)/ig, "");
        text.replace(/(^\s*)|(\s*$)/gi, "");
        text.replace(/[ ]{2,}/gi, " ");
        text.replace(/\n /, "\n");
        return text.split(' ').length;
    }

    private getWords = (text: string) => {
        text.replace(/(<([^>]+)>)/ig, "");
        text.replace(/(^\s*)|(\s*$)/gi, "");
        text.replace(/[ ]{2,}/gi, " ");
        text.replace(/\n /, "\n");
        return text.split(' ');
    }

    private textFieldControl = (labelName: string, isRequired: boolean, value: string, isDisabled?: boolean, errorMessage?: string, placeholder?: string, isInfoLabel?: boolean, infoText?: any[]): JSX.Element => {
        const { displayMode } = this.props
        let columnClassLabel = labelName == Constants.circularNumber ? `${styles.column3}` : ``;
        let columnClassInput = labelName == Constants.circularNumber && displayMode == Constants.lblNew ?
            `${styles.column9}` : `${styles.column12}`

        let textFieldJSX = <>
            <Field label={{
                children: <>
                    <Label className={`${styles.formLabel} ${styles.fieldTitle}`}>{labelName}</Label>
                    {isRequired &&
                        <Label style={{ color: "var(--colorPaletteRedForeground3)", paddingLeft: "var(--spacingHorizontalXS)" }}>*</Label>
                    }
                    {isInfoLabel &&
                        <Tooltip content={{
                            children: <div className={`${styles.row}`}>
                                {infoText && infoText.length > 0 && infoText.map((val, index) => {
                                    return <>
                                        <div className={`${styles.column12} ${styles.marginBottom}`} style={{ marginTop: index == 0 ? 10 : 0 }}>
                                            <Label className={`${styles.toolTipLabel}`}>{val}</Label>
                                        </div>
                                        <Divider appearance="subtle"></Divider>
                                    </>
                                })
                                }
                            </div>

                        }}
                            relationship="description"
                            positioning={'after'}
                            withArrow={true}>
                            <InfoRegular style={{ paddingLeft: "var(--spacingHorizontalXS)", cursor: "pointer" }} ></InfoRegular>
                        </Tooltip>
                    }

                </>
            }}
                // required={isRequired}
                validationState={isRequired && value == "" ? "error" : "none"}
                validationMessage={isRequired && value == "" ? errorMessage : ``} >
                <div className={`${styles.row}`}>
                    {labelName == Constants.circularNumber && displayMode == Constants.lblNew &&
                        <div className={`${columnClassLabel}`} style={{ marginTop: 5 }}>
                            <Label className={`${styles.formLabel} ${styles.fieldTitle}`} style={{ fontWeight: 400 }}>
                                {this.getCircularNumber()}
                            </Label>
                        </div>
                    }
                    <div className={`${columnClassInput}`}>
                        <Input value={value} maxLength={255}
                            disabled={isDisabled}
                            type={labelName == Constants.circularNumber && displayMode == Constants.lblNew ? "number" : "text"}
                            style={{ width: "100%" }}
                            input={{ style: { width: "100%" } }}
                            className={`${styles.formInput}`}
                            placeholder={placeholder ?? ``}
                            onChange={this.onInputChange.bind(this, labelName)}></Input>
                    </div>
                </div>
            </Field>
        </>;

        return textFieldJSX;
    }

    private onInputChange = (labelName: string, ev: React.ChangeEvent<HTMLInputElement>, data: InputOnChangeData) => {
        const { circularListItem } = this.state
        switch (labelName) {
            case Constants.circularNumber:
                circularListItem.CircularNumber = data?.value?.replace(/[^a-zA-Z0-9 ]/g, '');
                this.setState({ circularListItem })
                break;
            case Constants.subFileNo: circularListItem.SubFileCode = data.value?.replace(/[^a-zA-Z0-9,& ]/g, '');
                this.setState({ circularListItem });
                break;
            case Constants.keyWords: circularListItem.Keywords = data.value?.replace(/[^a-zA-Z0-9,& ]/g, '');
                this.setState({ circularListItem });
                break;
        }
    }

    private getCircularNumber = (circularNo?: string): string => {
        const { displayMode } = this.props
        let currentDate = new Date();
        let circularNumber = Text.format(Constants.circularNo, (currentDate.getFullYear() - 1908))
        // if(displayMode!=Constants.lblNew){
        //     circularNumber=circularNumber.split()
        // }

        return circularNumber;
    }

    private avatarControl = (labelName: string, value: string): JSX.Element => {
        let avatarControlJSX = <>
            <Field label={<Label className={`${styles.formLabel} ${styles.fieldTitle}`}>{labelName}</Label>} >
                <Persona
                    name={value}
                    size="medium"
                    primaryText={{ className: `${styles.formLabel}`, style: { margin: 5 } }}
                />
            </Field>
        </>;

        return avatarControlJSX;
    }

    private dropDownControl = (labelName: string, isRequired: boolean, value: string, options: any[], isDisabled?: boolean, errorMessage?: string, isInfoLabel?: boolean, infoText?: any[]): JSX.Element => {
        let dropDownControlJSX = <>
            <Field
                label={{
                    children: <>
                        <Label className={`${styles.formLabel} ${styles.fieldTitle}`}>{labelName}</Label>
                        {isRequired &&
                            <Label style={{ color: "var(--colorPaletteRedForeground3)", paddingLeft: "var(--spacingHorizontalXS)" }}>*</Label>
                        }
                        {isInfoLabel &&
                            <Tooltip content={{
                                children: <div className={`${styles.row}`}>
                                    {infoText && infoText.length > 0 && infoText.map((val, index) => {
                                        return <>
                                            <div className={`${styles.column12} ${styles.marginBottom}`} style={{ marginTop: index == 0 ? 10 : 0 }}>
                                                <Label className={`${styles.toolTipLabel}`}>{val}</Label>
                                            </div>

                                        </>
                                    })
                                    }
                                </div>,
                                style: { maxWidth: 400 }
                            }}
                                relationship="description" positioning={'after'} withArrow={true}>
                                <InfoRegular style={{ paddingLeft: "var(--spacingHorizontalXS)", cursor: "pointer" }} ></InfoRegular>
                            </Tooltip>
                        }

                    </>
                }}
                //required={isRequired}
                validationState={isRequired && value == "" ? "error" : "none"}
                validationMessage={isRequired && value == "" ? errorMessage : ``}
            >
                <Dropdown mountNode={{}} placeholder={`Select ${labelName}`} value={value}
                    selectedOptions={[value]}
                    disabled={isDisabled}
                    onOptionSelect={this.onDropDownChange.bind(this, `${labelName}`)}>
                    {options && options.length > 0 && options.map((val) => {
                        return <><Option key={`${val}`} className={`${styles.formLabel}`}>{val}</Option></>
                    })}
                </Dropdown>
            </Field>
        </>;
        return dropDownControlJSX;
    }

    private onDropDownChange = (labelName: string, event: SelectionEvents, data: OptionOnSelectData) => {
        const { circularListItem, attachedFile } = this.state;


        switch (labelName) {

            case Constants.issuedFor: circularListItem.IssuedFor = data.optionValue;
                this.setState({ circularListItem })

                break;
            case Constants.category: circularListItem.Category = data.optionValue;
                this.setState({ circularListItem });
                break;

            case Constants.classification:

                if (data.optionValue == Constants.lblMaster) {
                    circularListItem.Classification = data.optionValue;
                    circularListItem.CircularType = Constants.unlimited;
                    circularListItem.Expiry = null;
                    this.setState({ isLimited: false, expiryDate: null, isExpiryDateDisabled: true, ...circularListItem })
                } else if (data.optionValue == Constants.lblCircular) {
                    circularListItem.Classification = data.optionValue;
                    circularListItem.CircularType = Constants.limited;
                    this.setState({ isExpiryDateDisabled: false, isLimited: true, ...circularListItem });
                }

                break;

            case Constants.lblTemplate:
                let isFormValid = this.validateAllRequiredFields();
                if (isFormValid) {
                    const { attachedFile, selectedTemplate } = this.state
                    if (attachedFile == null) {
                        circularListItem.CircularTemplate = data.optionValue
                        this.setState({ selectedTemplate: data.optionValue, circularListItem }, async () => {
                            this.createUpdateCircularFile();
                        })
                    }
                    else if (attachedFile != null && isFormValid) {
                        let selectedTemplateVal = attachedFile != null ? selectedTemplate : ``;
                        circularListItem.CircularTemplate = selectedTemplateVal
                        this.setState({ selectedTemplate: selectedTemplateVal, circularListItem })
                    }

                }
                else {
                    this.setState({ isFormInValid: true })
                }
            default:
                break;
        }
    }

    private commentsSection = (labelName, isSelected) => {

        const { comments } = this.state;
        let history = comments.get(labelName);

        let commentSectionJSX = <>
            <div className={`${styles.row}`} >
                <div className={`${styles.column12}`} style={{ paddingLeft: 0 }}>
                    <Button
                        appearance="transparent"
                        iconPosition="before"
                        onClick={() => {
                            this.onCommentHistoryClick(labelName)
                        }}
                        icon={isSelected ? <ChevronDownRegular /> : <ChevronUpRegular />}>
                        {labelName}
                    </Button>
                </div>
            </div>

            {
                isSelected && history.length > 0 && history?.map((val) => {
                    return <>
                        <Divider appearance="subtle"></Divider>
                        <div className={`${styles.row} ${AnimationClassNames.slideDownIn20}`} style={{ paddingTop: 5 }}>
                            {/* <div className={`${styles.column2}`} style={{ paddingLeft: 20 }}>
                                {this.onFormatDate(new Date(val.commentDate))}
                            </div>
                            <div className={`${styles.column10}`} style={{ borderLeft: "1px solid lightgrey" }}>
                                {val.comment}
                            </div>
                            <div className={`${styles.column2}`}>
                            </div>
                            <div className={`${styles.column10}`} style={{ borderLeft: "1px solid lightgrey" }}>
                                <Label size="small">  {val?.user?.split('|')[0]}</Label>
                            </div>*/}
                            <div className={`${styles.column2}`} style={{ textAlign: "end" }}>
                                {this.onFormatCommentsDate(val.commentDate)}
                            </div>
                            <div className={`${styles.column7}`} style={{
                                borderLeft: "1px solid lightgrey",
                                borderRight: "1px solid lightgrey"
                            }}>
                                {val.comment}
                            </div>
                            <div className={`${styles.column3}`} style={{
                                paddingLeft: 20,

                            }}>
                                <Persona primaryText={{ style: { fontFamily: "Roboto" } }} size="small" name={val?.user?.split('|')[0]}></Persona>
                            </div>
                        </div>
                        <div className={`${styles.row} ${AnimationClassNames.slideDownIn20}`} style={{ paddingBottom: 10 }}>


                        </div>

                        <Divider appearance="subtle"></Divider>

                    </>
                })
            }
        </>;

        return commentSectionJSX;
    }

    private onCommentHistoryClick = (labelName) => {
        const { selectedCommentSection } = this.state;
        switch (labelName) {
            case Constants.lblCommentsMaker: selectedCommentSection.isMakerSelected = !selectedCommentSection.isMakerSelected;
                break;
            case Constants.lblCommentsCompliance: selectedCommentSection.isComplianceSelected = !selectedCommentSection.isComplianceSelected;
                break;
            case Constants.lblCommentsChecker: selectedCommentSection.isCheckerSelected = !selectedCommentSection.isCheckerSelected;
                break;
        }

        this.setState({ selectedCommentSection })

    }


    /**
    |--------------------------------------------------
    | This attachment Link is Circular Content File
    |--------------------------------------------------
    */
    private attachmentLink = (selectedFile): JSX.Element => {
        const { displayMode } = this.props;
        let disableControl = displayMode == Constants.lblViewCircular;
        let attachedLinkJSX = <div className={`${styles.row}`}>
            <div className={`${styles.column12}`}>
                <Attach16Filled></Attach16Filled>
                <Link
                    onClick={() => {
                        let providerValue = this.context;
                        const { context, services } = providerValue as IBobCircularRepositoryProps;
                        const { attachedFile } = this.state
                        this.setState({
                            selectedFileName: attachedFile.FileName,
                            documentPreviewURL: this.circularContentPreviewURL(context)
                        }, () => {

                        })

                    }}
                    style={{
                        wordBreak: "break-all",
                        padding: 5
                    }}
                >{`${selectedFile.FileName}`}</Link>
                <Button icon={<Delete16Regular></Delete16Regular>} style={{ marginLeft: 5 }}
                    disabled={disableControl}
                    onClick={() => {
                        this.setState({ isDeleteCircularFile: true, isFormInValid: false })
                    }}></Button>

            </div>
        </div>

        return attachedLinkJSX;

    }

    private datePickerControl = (labelName: string, value: any, isRequired?: boolean, isDisabled?: boolean): JSX.Element => {
        let datePickerJSX = <>
            <Field
                label={<Label className={`${styles.formLabel} ${styles.fieldTitle}`}>{labelName}</Label>}
                validationState={isRequired && value == null ? "error" : "none"}
                validationMessage={isRequired && value == null ? `Field cannot be empty` : ``}
                required={isRequired}>
                {/* <Input input={{ readOnly: true, type: "date" }} root={{ style: { fontFamily: "Roboto" } }}></Input> */}

                <DatePicker mountNode={{}}
                    formatDate={this.onFormatDate}
                    value={value}
                    disabled={isDisabled}
                    minDate={new Date()}
                    contentAfter={
                        <>
                            <Button icon={<ArrowCounterclockwiseRegular />}
                                disabled={isDisabled}
                                appearance="transparent"
                                title="Reset"
                                onClick={this.onResetDateClick.bind(this, `${labelName}`)}>
                            </Button>
                            <Button disabled={isDisabled} icon={<CalendarRegular />} appearance="transparent"></Button>
                        </>}
                    onSelectDate={this.onSelectDate.bind(this, `${labelName}`)}
                    input={{ style: { fontFamily: "Roboto", background: isDisabled ? `#7676761c` : `inherit` } }} />


            </Field>
        </>

        return datePickerJSX;
    }

    private onFormatDate = (date?: Date): string => {

        // if (date != null && date?.toString() != "Invalid Date") {
        //     date.setHours(0, 0, 0)
        //     return (
        //         (date.getDate() < 10 ? `0${date.getDate()}` : date.getDate()) + "/" + ((date.getMonth() + 1) < 10 ? `0` + (date.getMonth() + 1) : date.getMonth() + 1) + "/" + date.getFullYear()
        //     );
        // }
        // else {
        //     return "";
        // }

        return !date
            ? ""
            : (date.getDate() < 9 ? (`0` + date.getDate()) : date.getDate()) +
            "/" +
            ((date.getMonth() + 1 < 9 ? (`0${date.getMonth() + 1}`) : date.getMonth() + 1)) +
            "/" +
            (date.getFullYear());
    };


    private onFormatCommentsDate = (dateString?: Date | string): string => {

        // if (date != null && date?.toString() != "Invalid Date") {
        //     date.setHours(0, 0, 0)
        //     return (
        //         (date.getDate() < 10 ? `0${date.getDate()}` : date.getDate()) + "/" + ((date.getMonth() + 1) < 10 ? `0` + (date.getMonth() + 1) : date.getMonth() + 1) + "/" + date.getFullYear()
        //     );
        // }
        // else {
        //     return "";
        // }

        const date = new Date(dateString);

        // Extract local date components
        const day = String(date.getDate()).padStart(2, '0');
        const month = String(date.getMonth() + 1).padStart(2, '0'); // Months are zero-indexed
        const year = date.getFullYear();

        let hours = date.getHours();
        const minutes = String(date.getMinutes()).padStart(2, '0');
        const seconds = String(date.getSeconds()).padStart(2, '0');

        // Determine AM or PM
        const ampm = hours >= 12 ? 'PM' : 'AM';

        // Convert 24-hour time to 12-hour time
        hours = hours % 12 || 12; // Convert 0 to 12 for midnight



        // Format hours with leading zero if necessary
        const formattedHours = String(hours).padStart(2, '0');

        // Format the date as dd/MM/YYYY HH:MM:SS AM/PM
        const formattedDate = `${day}/${month}/${year} ${formattedHours}:${minutes} ${ampm}`;

        return formattedDate;

        // return !date
        //     ? ""
        //     : (date.getDate() < 9 ? (`0` + date.getDate()) : date.getDate()) +
        //     "/" +
        //     ((date.getMonth() + 1 < 9 ? (`0${date.getMonth() + 1}`) : date.getMonth() + 1)) +
        //     "/" +
        //     (date.getFullYear());
    };


    private onSelectDate = (labelName: string, date: Date | null) => {
        const { circularListItem } = this.state;
        const dateString = date.toISOString();
        let dateFormat = date.getFullYear() + `-` + (date.getMonth() + 1 < 10 ? `0` + (date.getMonth() + 1) : (date.getMonth() + 1));
        dateFormat += `-` + (date.getDate() < 10 ? `0` + (date.getDate()) : date.getDate());
        switch (labelName) {
            case Constants.expiry:
                circularListItem.Expiry = dateFormat + "T00:00:00Z";
                this.setState({ circularListItem, expiryDate: date });
                break;

        }

    }

    private onResetDateClick = (labelName: string) => {
        const { circularListItem } = this.state
        switch (labelName) {
            case Constants.expiry: circularListItem.Expiry = null;
                this.setState({ circularListItem, expiryDate: null });
                break;

        }
    }

    private switchControl = (labelName, isRequired, switchLabel, orientation: any = "vertical", isChecked?: boolean, isDisabled?: boolean): JSX.Element => {
        let switchControlJSX = <>
            <Field
                label={{ children: <Label className={`${styles.fieldLabel} ${styles.fieldTitle}`}>{labelName}</Label>, style: { width: 0 } }}
                orientation={orientation}
                style={{ width: 0 }}
                required={isRequired}>
                <Switch required={isRequired}
                    checked={isChecked}
                    disabled={isDisabled}
                    onChange={this.onSwitchChange.bind(this, labelName)}
                    label={<Label className={`${styles.formLabel}`}>{switchLabel}</Label>} />
            </Field>
        </>;
        return switchControlJSX;
    }

    private onSwitchChange = (labelName: string, ev: React.ChangeEvent<HTMLInputElement>, data: SwitchOnChangeData) => {

        const { circularListItem } = this.state;

        switch (labelName) {
            case Constants.type:

                if (data.checked) {
                    circularListItem.CircularType = Constants.unlimited;
                    circularListItem.Expiry = null;
                    this.setState({ isLimited: false, expiryDate: null, ...circularListItem });
                }
                else {
                    circularListItem.CircularType = Constants.limited;
                    this.setState({ isLimited: true, ...circularListItem })
                }

                break;

            case Constants.compliance:
                if (data.checked) {
                    circularListItem.Compliance = Constants.lblYes;
                    this.setState({ circularListItem, lblCompliance: Constants.lblCompliance });
                }
                else {
                    circularListItem.Compliance = Constants.lblNo;
                    this.setState({ circularListItem, lblCompliance: `` })
                }


            default:
                break;
        }

    }

    /**
    |--------------------------------------------------
    | SOP File Upload Controls
    |--------------------------------------------------
    */

    private fileUploadControl = (labelName: string, filePickerRef: any, isDisabled?: boolean): JSX.Element => {
        const { configuration } = this.state;
        const { configVal } = Constants;
        let maxFileSizeMB = configuration.filter(val => val.Title == configVal.SOPFileMaxSizeinMB)[0]?.Limit ?? 5;
        let sopFileLimit = configuration.filter(val => val.Title == configVal.SOPFileUpload)[0]?.Limit ?? 5;
        let fileUploadJSX = <>
            <input
                id={`file-picker_${labelName}`}
                style={{ display: "none" }}
                type="file"
                onChange={(e) => { this.onFileUploadChange(labelName, e) }}
                ref={filePickerRef}
                multiple
            />

            <Button icon={<ArrowUpload16Regular />}
                onClick={this.onUploadClick.bind(this, labelName)}
                disabled={isDisabled}
                iconPosition="before"
            > Upload SOP File
            </Button>
            <Field label={
                <Label className={`${styles.formLabel} `}>
                    {` (Maximum ${maxFileSizeMB}MB .pdf & .docx file allowed. Up to ${sopFileLimit} files.)`}
                </Label>
            }
                required={false}>
            </Field>

        </>;

        return fileUploadJSX;
    }

    private onFileUploadChange = (labelName: string, e: React.ChangeEvent<HTMLInputElement>) => {
        const { sopAttachmentColl, circularListItem, configuration } = this.state
        const { configVal } = Constants;
        const files = e.target.files;

        let providerValue = this.context;
        const { context, services, serverRelativeUrl } = providerValue as IBobCircularRepositoryProps;



        let invalidFileSize = [];
        let inValidFileType = [];
        let invalidFileLimit = [];

        let fileCount = files.length + sopAttachmentColl?.length;
        let isFormValid = this.validateAllRequiredFields();
        let maxFileSizeMB = configuration.filter(val => val.Title == configVal.SOPFileMaxSizeinMB)[0]?.Limit ?? 5120;
        let uploadLimit = configuration.filter(val => val.Title == configVal.SOPFileUpload)[0]?.Limit ?? 5;

        if (files && isFormValid) {

            for (let i = 0; i < files.length; i++) {

                let fileExtension = files[i].name.split('.');

                let circularNumberText = circularListItem.CircularNumber;
                let circularNumberIndexOf = circularNumberText.indexOf(`${this.getCircularNumber()}`);
                let circularFileName = ``;
                let isCircularFile = false;
                // if BOB:BR:116: not present then circular Number will be this
                if (circularNumberIndexOf == -1) {
                    circularFileName = `${this.getCircularNumber()}` + `${circularNumberText}` + `.docx`;
                    isCircularFile = circularFileName == files[i].name
                }
                else {
                    circularFileName = circularNumberText.split(':').join('_') + `.docx`;
                    isCircularFile = circularFileName == files[i].name
                }

                if (fileExtension.length < 3 && !isCircularFile && (files[i].name.indexOf('.docx') > -1 || files[i].name.indexOf('.pdf') > -1)) {
                    let sizeInMB = Math.round((files[i].size) / 1024);
                    if (sizeInMB <= (maxFileSizeMB * 1024)) {
                        if ((this.sopFileAttachments.has(files[i].name)) || fileCount <= 5) {
                            this.sopFileAttachments.delete(files[i].name);
                            this.sopFileAttachments.set(files[i].name, files[i]);
                        }
                        else if (fileCount <= uploadLimit) {
                            this.sopFileAttachments.set(files[i].name, files[i]);
                        }
                    }
                    else {
                        invalidFileSize.push(sizeInMB);
                    }

                }
                else {
                    inValidFileType.push(files[i].type)
                }

            }

            if (invalidFileSize.length > 0) {
                this.setState({ isFileSizeAlert: true, isFileTypeAlert: false })
            }
            else if (inValidFileType.length > 0) {
                this.setState({ isFileSizeAlert: false, isFileTypeAlert: true })
            }

            this.setState({ sopUploads: this.sopFileAttachments }, async () => {
                const { sopUploads } = this.state
                let attachmentColl = [];
                let i = 0;

                // await services.addFileToListItem(serverRelativeUrl, Constants.circularList, 3152, files[0].name, files[0]).then((val) => {
                //     console.log(val);
                // }).catch((error) => {
                //     console.log(error);
                // })

                sopUploads.forEach(async (value, key) => {
                    value.index = i;
                    value.FileName = key;
                    value.ServerRelativeUrl = value?.ServerRelativeUrl ?? ``;
                    value.UniqueId = value?.UniqueId ?? ``;
                    if (value.hasOwnProperty(`isFileEdit`)) {
                        value.isFileEdit = value?.isFileEdit
                    }

                    attachmentColl.push(value);
                    i++;
                });

                this.setState({ sopAttachmentColl: attachmentColl })
            })
        }

        else {
            this.setState({ isFormInValid: true })
        }

    }

    private onUploadClick = (labelName: string) => {
        switch (labelName) {
            case Constants.sop:
                this.sopFileInput.current.value = "";
                this.sopFileInput.current.click();

                break;
            default:
                break;
        }
    }

    private sopFilesControl = (isDisabled?: boolean): JSX.Element => {

        const { sopAttachmentColl } = this.state;

        let sopFileUploadJSX = <>
            {
                sopAttachmentColl && sopAttachmentColl.length > 0 &&

                sopAttachmentColl.map((file) => {
                    const fileName = file.FileName;

                    return <div className={`${styles.column12}`} style={{ marginBottom: 5 }}>
                        <div className={`${styles.row}`}>
                            <div className={`${styles.column1}`}> <Attach16Filled></Attach16Filled></div>
                            <div className={`${styles.column10}`}>
                                <Link
                                    onClick={() => {
                                        this.setState({ selectedFileName: fileName }, () => {
                                            this.createSOPPreviewURL(file)
                                        })
                                    }}
                                    style={{
                                        wordBreak: "break-all",
                                        padding: 5
                                    }}
                                >{fileName}</Link>
                                <Button disabled={isDisabled}
                                    icon={<Delete16Regular></Delete16Regular>} style={{ marginLeft: 5 }}
                                    onClick={() => { this.deleteSOPUploadedFiles(fileName) }}></Button>
                            </div>
                            <div className={`${styles.column1}`}>

                            </div>
                        </div>
                    </div>
                })


            }
        </>

        return sopFileUploadJSX;
    }




    private alertControl = (showAlert, headerTitle?: string, headerColor?: string, validationMessage?: string, buttonJSX?: any): JSX.Element => {
        let alertJSX = <>
            <Dialog modalType="alert" defaultOpen={(showAlert)} >
                <DialogSurface style={{ maxWidth: 300 }}>
                    <DialogBody style={{ display: "block" }}>
                        <DialogTitle style={{ fontFamily: "Roboto", marginBottom: 15, color: headerColor ?? "#B10E1C" }}>{`${headerTitle ?? ``}`}</DialogTitle>
                        <DialogContent styles={{
                            header: { display: "none" },
                            inner: { padding: 0 },
                            innerContent: { fontFamily: "Roboto", marginBottom: 15 }
                        }}>
                            {`${validationMessage ?? `Please input all the fields marked as *`}`}
                        </DialogContent>
                        <DialogActions style={{ justifyContent: "center" }}>

                            {buttonJSX}

                        </DialogActions>
                    </DialogBody>
                </DialogSurface>
            </Dialog>
        </>

        return alertJSX
    }

    private alertButton = (): JSX.Element => {

        const { isFormInValid, isDeleteCircularFile, isFileSizeAlert, isFileTypeAlert, isDuplicateCircular } = this.state
        let alertButtonJSX;

        if (isFormInValid || isFileSizeAlert || isFileTypeAlert || isDuplicateCircular) {
            alertButtonJSX = <div className={`${styles.row}`}>
                <div className={`${styles.column12}`}>
                    <Button appearance="secondary"
                        onClick={() => {
                            this.setState({
                                isFormInValid: false,
                                isDeleteCircularFile: false,
                                isFileSizeAlert: false,
                                isFileTypeAlert: false,
                                isDuplicateCircular: false
                            })
                        }}>Close</Button>
                </div>
            </div>

        } else if (isDeleteCircularFile) {
            alertButtonJSX = <div className={`${styles.row}`}>
                <div className={`${styles.column6}`}>
                    <Button appearance="primary"
                        onClick={() => {
                            this.deleteAttachment();
                        }}>Delete</Button>
                </div>
                <div className={`${styles.column6}`}>
                    <Button appearance="secondary"
                        onClick={() => {
                            this.setState({ isFormInValid: false, isDeleteCircularFile: false })
                        }}>Close</Button>
                </div>
            </div>
        }


        return alertButtonJSX;

    }


    private deleteAttachment = () => {
        let providerValue = this.context;
        const { services, serverRelativeUrl } = providerValue as IBobCircularRepositoryProps;
        const { currentCircularListItemValue, attachedFile } = this.state

        this.setState({ isFormInValid: false, isDeleteCircularFile: false, isLoading: true }, async () => {
            await services.deleteListItemAttachment(serverRelativeUrl, Constants.circularList,
                parseInt(currentCircularListItemValue.ID), attachedFile.FileName).then((error) => {
                    this.setState({ isLoading: false, attachedFile: null, selectedTemplate: ``, documentPreviewURL: `` })
                }).catch((error) => {
                    console.log(error);
                    this.setState({ isLoading: false })
                })
        })

    }

    private deleteBackDialogControl = (showAlert): JSX.Element => {
        const { isDelete, isBack } = this.state
        let dialogControlJSX = <>
            <Dialog modalType="alert" defaultOpen={(showAlert)}>
                <DialogSurface>
                    <DialogBody style={{ gridTemplateColumns: "1fr 0fr auto" }}>
                        <DialogTitle style={{ fontFamily: "Roboto" }}>{isDelete ? `Delete Circular` : `Back to Home`}</DialogTitle>
                        <DialogContent styles={{ header: { display: "none" }, inner: { padding: 0 }, innerContent: { fontFamily: "Roboto" } }}>
                            {isDelete ? `Are you sure you want to delete the circular?` : `Are you sure you want to leave this page?`}
                        </DialogContent>
                        <DialogActions>
                            <DialogTrigger>
                                <Button appearance="primary" onClick={() => {
                                    this.setState({ isBack: false, isDelete: false }, () => {
                                        this.props.onGoBack()
                                    })
                                }}>{isDelete ? Constants.delete : Constants.goBack}</Button>
                            </DialogTrigger>
                            <Button appearance="secondary"
                                onClick={() => {
                                    this.setState({ isBack: false, isDelete: false })
                                }}>Close</Button>

                        </DialogActions>
                    </DialogBody>
                </DialogSurface>
            </Dialog>
        </>
        return dialogControlJSX;
    }

    private workingOnIt = (): JSX.Element => {
        const { isLoading, isSuccess } = this.state
        let workingJSX = <>
            <Dialog modalType="alert" defaultOpen={true}>
                <DialogSurface style={{ maxWidth: 300 }}>
                    <DialogBody style={{ display: "block" }}>
                        <DialogContent styles={{ header: { display: "none" }, inner: { padding: 0 }, innerContent: { fontFamily: "Roboto" } }}>
                            {isLoading && <Spinner size="large" labelPosition="below" label={"Please Wait..."}></Spinner>}
                            {isSuccess && <>
                                <Image style={{ width: 160, paddingLeft: 100 }} src={require(`../../assets/success.gif`)} shape="circular" fit="contain"></Image>
                                <Label className={`${styles.formLabel}`} style={{
                                    display: "block",
                                    width: "100%",
                                    textAlign: "center",
                                    paddingTop: 5, paddingBottom: 5
                                }}>
                                    Item saved successfully
                                </Label>

                            </>}
                        </DialogContent>
                        {isSuccess && <DialogActions>
                            <DialogTrigger>
                                <Button style={{ width: "100%", marginTop: 4 }} appearance="primary"
                                    onClick={() => {
                                        this.setState({ isSuccess: false }, () => {
                                            const { circularListItem } = this.state;
                                            if (circularListItem.CircularStatus != Constants.draft) {
                                                this.props.onGoBack()
                                            }

                                        })
                                    }} >OK</Button>
                            </DialogTrigger>
                        </DialogActions>
                        }
                    </DialogBody>
                </DialogSurface>
            </Dialog>
        </>;
        return workingJSX;
    }

    private messageBarControl = (intent, message): JSX.Element => {
        let messageBarJSX = <>
            <MessageBar key={intent} intent={intent}>
                <MessageBarBody>
                    <Label className={`${styles.formLabel} ${styles.fieldTitle}`}>{message}</Label>
                </MessageBarBody>
            </MessageBar>
        </>;

        return messageBarJSX;
    }

    private validateAllRequiredFields = (): boolean => {
        const { circularListItem } = this.state;
        let providerValue = this.context;
        const { isUserChecker, isUserMaker, isUserCompliance } = providerValue as IBobCircularRepositoryProps;
        const { lblNew, draft, cmmtCompliance, cmmtChecker, limited } = Constants;
        const { Subject, CircularNumber, IssuedFor, Category, Classification, CircularType, Expiry, Keywords } = circularListItem;

        let circularStatus = circularListItem.CircularStatus;
        let isValid = true;
        if (circularStatus == lblNew || circularStatus == draft || circularStatus == cmmtCompliance || circularStatus == cmmtChecker) {

            if (Subject?.trim() == "" || CircularNumber?.trim() == "" || IssuedFor == "" ||
                Category == "" || Classification == "" || Keywords == "") {
                isValid = false
            }
            else if (CircularType == limited) {
                isValid = !(Expiry == null)
            }
        }
        else {
            if (circularStatus == Constants.sbmtCompliance && isUserCompliance) {
                if (circularListItem.CommentsCompliance?.trim() == "") {
                    isValid = false
                }
            }

            if (circularStatus == Constants.sbmtChecker && isUserChecker) {
                if (circularListItem.CommentsChecker?.trim() == "") {
                    isValid = false
                }
            }
        }

        return isValid;
    }


    private filterPanelSupportingDocument = (): JSX.Element => {

        const { selectedSupportingCirculars, circularListItem, configuration } = this.state
        let providerValue = this.context as IBobCircularRepositoryProps;
        const { configVal } = Constants;
        let maxSupportingDocument = configuration.filter(val => val.Title == configVal.SupportingDocuments)[0].Limit;

        let panelSupportingDocumentsJSX = <>
            <SupportingDocument department={`${circularListItem.Department}`}
                providerValue={providerValue}
                selectedSupportingCirculars={selectedSupportingCirculars}
                onDismiss={(supportingCirculars) => {
                    this.setState({
                        openSupportingDocument: false,
                        selectedSupportingCirculars: supportingCirculars.slice(0, maxSupportingDocument)
                    }, () => {

                        const { circularListItem, selectedSupportingCirculars } = this.state;
                        console.log(selectedSupportingCirculars.length);
                        console.log(maxSupportingDocument)

                        if (supportingCirculars.length > 0) {

                            let supportingDoc = supportingCirculars.map((val: ICircularListItem) => {
                                return {
                                    ID: val.ID,
                                    Id: val.Id,
                                    CircularNumber: val.CircularNumber
                                }
                            });

                            circularListItem.SupportingDocuments = JSON.stringify(supportingDoc);
                            this.setState({ circularListItem })
                        }
                    })
                }}
                completeLoading={() => { this.setState({ isLoading: false }) }}
            />
        </>
        return panelSupportingDocumentsJSX;
    }

    private deleteSupportingCircular = (supportingCircular) => {
        const { selectedSupportingCirculars, circularListItem } = this.state;
        let index = selectedSupportingCirculars.indexOf(supportingCircular);
        if (index > -1) {
            selectedSupportingCirculars.splice(index, 1);
            //delete selectedSupportingCirculars[index];
            this.setState({ selectedSupportingCirculars }, () => {
                const { selectedSupportingCirculars } = this.state;

                let supportingDoc = selectedSupportingCirculars?.map((val: ICircularListItem) => {
                    return {
                        ID: val.ID,
                        Id: val.Id,
                        CircularNumber: val.CircularNumber
                    }
                });

                if (supportingDoc.length > 0) {
                    circularListItem.SupportingDocuments = JSON.stringify(supportingDoc);
                }
                else {
                    circularListItem.SupportingDocuments = ``;
                }

                this.setState({ circularListItem })

            })
        }
    }


    private checkCircularNumberExist = async (circularNumber): Promise<boolean> => {

        let providerValue = this.context;
        const { services, serverRelativeUrl } = providerValue as IBobCircularRepositoryProps;

        let circularFilterString = Text.format(Constants.filterCircularNumber, circularNumber);

        let validCircularNumberPromise = await services.filterLargeListItem(serverRelativeUrl, Constants.circularList, circularFilterString).then((item) => {
            if (item.length > 0) {
                return Promise.resolve(true)
            }
            else {
                return Promise.resolve(false)
            }
        }).catch((error) => {
            console.log(error)
            return Promise.reject(error)
        })

        return validCircularNumberPromise;

    }

    private updateCommentsHistory = () => {
        const { circularListItem, currentItemID } = this.state;
        let providerValue = this.context;
        const { context } = providerValue as IBobCircularRepositoryProps;
        let listItemID = circularListItem.ID;

        if (circularListItem.CommentsMaker != "") {

            let makerCommentsHistory: any[] = [];
            if (circularListItem?.MakerCommentsHistory) {
                makerCommentsHistory = makerCommentsHistory.concat(JSON.parse(circularListItem?.MakerCommentsHistory));
            }

            let makerComment = [{
                commentDate: new Date().toISOString(),
                user: `${context.pageContext.user.displayName}|${context.pageContext.user.email}|${currentItemID}`,
                comment: circularListItem.CommentsMaker
            }]

            makerCommentsHistory = makerCommentsHistory.concat(makerComment);
            circularListItem.MakerCommentsHistory = JSON.stringify(makerCommentsHistory);
            this.setState({ circularListItem });

        }

        if (circularListItem.CommentsCompliance != "") {

            let complianceCommentsHistory: any[] = []
            if (circularListItem?.ComplianceCommentsHistory) {
                complianceCommentsHistory = complianceCommentsHistory.concat(JSON.parse(circularListItem?.ComplianceCommentsHistory))
            }

            let complianceComment = [{
                commentDate: new Date().toISOString(),
                user: `${context.pageContext.user.displayName}|${context.pageContext.user.email}|${currentItemID}`,
                comment: circularListItem.CommentsCompliance
            }];

            complianceCommentsHistory = complianceCommentsHistory.concat(complianceComment);
            circularListItem.ComplianceCommentsHistory = JSON.stringify(complianceCommentsHistory);
            this.setState({ circularListItem })
        }

        if (circularListItem.CommentsChecker != "") {

            let checkerCommentsHistory: any[] = [];

            if (circularListItem?.CheckerCommentsHistory) {
                checkerCommentsHistory = checkerCommentsHistory.concat(JSON.parse(circularListItem?.CheckerCommentsHistory))
            }

            let checkerComment = [{
                commentDate: new Date().toISOString(),
                user: `${context.pageContext.user.displayName}|${context.pageContext.user.email}|${currentItemID}`,
                comment: circularListItem.CommentsChecker
            }];

            checkerCommentsHistory = checkerCommentsHistory.concat(checkerComment);
            circularListItem.CheckerCommentsHistory = JSON.stringify(checkerCommentsHistory);
            this.setState({ circularListItem });
        }
    }


    private saveForm = (status?: string) => {
        const { circularListItem, currentCircularListItemValue, sopAttachmentColl, sopUploads, attachedFile, auditListItem } = this.state;
        let isFormValid = this.validateAllRequiredFields();
        const { displayMode } = this.props;

        if (isFormValid) {

            let providerValue = this.context;
            const { services, serverRelativeUrl } = providerValue as IBobCircularRepositoryProps;

            let circularNumberText = circularListItem.CircularNumber;
            let circularNumberIndexOf = circularListItem.CircularNumber.indexOf(`${this.getCircularNumber()}`);
            // if BOB:BR:116: not present then circular Number will be this
            if (circularNumberIndexOf == -1) {
                circularListItem.CircularNumber = `${this.getCircularNumber()}` + `${circularNumberText}`;
            }




            //circularListItem.CircularCreationDate = new Date().toISOString();

            this.setState({ isLoading: true }, async () => {

                // console.log(circularListItem)

                let isCircularNumberExist = await this.checkCircularNumberExist(circularListItem.CircularNumber).
                    then((val) => {
                        return val
                    }).catch((error) => {
                        return false;
                    })


                if (currentCircularListItemValue == undefined && !isCircularNumberExist) {

                    circularListItem.CircularStatus = status;
                    circularListItem.CircularCreationDate = new Date().toISOString().split('T')[0] + `T00:00:00Z`;


                    await services.createItem(serverRelativeUrl, Constants.circularList, circularListItem).then(async (value) => {
                        circularListItem.CircularNumber = circularNumberText.replace(`${this.getCircularNumber()}`, ``);
                        circularListItem.CircularCreationDate = value?.Created;
                        let itemID = parseInt(value.ID);
                        if (this.sopFileAttachments.size > 0) {

                            let newSOPUpload: any[] = []
                            this.sopFileAttachments.forEach((val, key) => {
                                return newSOPUpload.push(val);
                            })

                            let newSOPUploadPromise = await services.addListItemAttachments(serverRelativeUrl, Constants.circularList, itemID, newSOPUpload).
                                then((result) => {
                                    return result;
                                }).catch((error) => {
                                    return error;
                                })

                            console.log(newSOPUploadPromise);
                        }
                        this.setState({ isSuccess: true, isLoading: false, circularListItem, currentCircularListItemValue: value })
                    }).catch((error) => {
                        this.setState({ isLoading: false })
                    });
                }
                else {

                    let ID = parseInt(currentCircularListItemValue.ID);
                    /**
                    |--------------------------------------------------
                    | Update Comments History & Store in respective comments History Columns
                    |--------------------------------------------------
                    */
                    console.log(`Updating Comments History`)
                    this.updateCommentsHistory();
                    console.log(`Updated Comments History Object`)

                    /**
                    |--------------------------------------------------
                    | 1. Current Circular Item Status is draft then keep it as draft 
                      2. if status(coming as parameter) from button call & is other than draft and 
                      when save as Draft is clicked then Current item status should stay as item Status 
                    |--------------------------------------------------
                    */
                    if (circularListItem.CircularStatus == Constants.draft) {
                        circularListItem.CircularStatus = status;
                        auditListItem.CircularAuditStatus = status;
                    }
                    else if (status != Constants.draft) {
                        circularListItem.CircularStatus = status;
                        auditListItem.CircularAuditStatus = status;
                    }

                    if (status == Constants.published) {
                        let circularFileName = attachedFile ? attachedFile.FileName : ``;
                        if (circularFileName != "") {
                            const { circularList } = Constants;
                            /**
                            |--------------------------------------------------
                            | Convert File to PDF on publish
                            |--------------------------------------------------
                            */

                           
                        }

                        circularListItem.PublishedDate = new Date().toISOString().split('T')[0] + `T00:00:00Z`;
                    }


                    if (circularListItem.CircularStatus != Constants.lblNew) {

                        await services.updateItem(serverRelativeUrl, Constants.circularList, ID, circularListItem).then(async (value) => {
                            circularListItem.CircularNumber = displayMode == Constants.lblNew ? circularNumberText.replace(`${this.getCircularNumber()}`, ``) : circularListItem.CircularNumber;
                            circularListItem.CircularCreationDate = value?.Created;
                            value.Author = currentCircularListItemValue.Author;


                            if (this.deleteSOPFileAttachments && this.deleteSOPFileAttachments.size > 0) {

                                await services.recycleListItemAttachments(serverRelativeUrl, Constants.circularList,
                                    ID, this.deleteSOPFileAttachments).then((deleteResult) => {
                                        console.log(deleteResult)
                                    }).catch((error) => {
                                        console.log(error)
                                    })

                                // let deleteFileRelativeUrl: string[] = [];

                                // this.deleteEditFileAttachments.forEach((value, key) => {
                                //     deleteFileRelativeUrl.push(value.ServerRelativeUrl);
                                // });
                            }


                            if (sopAttachmentColl.length > 0) {

                                let updateNewAttachment = [];
                                sopAttachmentColl.map((val) => {
                                    if (!val.hasOwnProperty(`isFileEdit`)) {
                                        let file = sopUploads.get(val.FileName);
                                        updateNewAttachment.push(file)
                                    }
                                })

                                if (updateNewAttachment.length > 0) {

                                    await services.addListItemAttachments(serverRelativeUrl, Constants.circularList, ID, updateNewAttachment).then((results) => {
                                        return results
                                    }).catch((error) => {
                                        return error;
                                    })
                                }

                            }

                            /**
                                |--------------------------------------------------
                                | Logging in to Comments Audit Log List
                                |--------------------------------------------------
                                */
                            if (circularListItem.CircularStatus != Constants.draft) {

                                auditListItem.Title = circularListItem.CircularNumber;

                                await services.createItem(serverRelativeUrl, Constants.commentsAuditLogs, auditListItem).then((auditItem) => {
                                    console.log(`Item Added to comments Log List`, auditItem);
                                }).catch((error) => {
                                    console.log(`Error Adding to comments Audit logs`);
                                    console.log(error)
                                })
                            }

                            this.setState({
                                isSuccess: true,
                                circularListItem,
                                currentCircularListItemValue: value,
                                isLoading: false
                            })
                        }).catch((error) => {
                            console.log(error);
                            this.setState({ isLoading: false })
                        });
                    }
                    else {
                        circularListItem.CircularNumber = circularNumberText.replace(`${this.getCircularNumber()}`, ``);
                        if (circularListItem.CircularStatus == Constants.lblNew) {
                            this.setState({ isDuplicateCircular: true, isLoading: false, circularListItem })
                        }
                    }
                }

            })
        }

        else {
            this.setState({ isFormInValid: true })
        }

    }

    private createUpdateCircularFile = async () => {
        const { circularListItem } = this.state;
        let circularNumberText = circularListItem.CircularNumber;
        let circularNumberIndexOf = circularListItem.CircularNumber.indexOf(`${this.getCircularNumber()}`);

        // if BOB:BR:116: not present then circular Number will be this
        if (circularNumberIndexOf == -1) {
            circularListItem.CircularNumber = `${this.getCircularNumber()}` + `${circularNumberText}`
        }

        let circularNumberExist = await this.checkCircularNumberExist(circularListItem.CircularNumber).
            then((val) => {
                return val
            }).catch((error) => {
                console.log(error)
                return false;
            })

        if (circularListItem.CircularStatus == Constants.lblNew) {
            if (!circularNumberExist) {
                this.addCircularItemAndFile();
            }
            else {
                circularListItem.CircularNumber = circularNumberText.replace(`${this.getCircularNumber()}`, ``);
                this.setState({ isDuplicateCircular: true, circularListItem, selectedTemplate: `` });
            }
        }
        else if (circularListItem.CircularStatus != Constants.lblNew) {
            this.updateCircularItemAndFile()
        }
    }

    private addCircularItemAndFile = () => {

        let providerValue = this.context;
        const { services, serverRelativeUrl, context } = providerValue as IBobCircularRepositoryProps;
        const { templateFiles, selectedTemplate, currentCircularListItemValue, attachedFile, circularListItem } = this.state;
        let circularStatus = circularListItem.CircularStatus;
        let selectedTemplateFile = templateFiles.filter((val) => {
            return val.templateName == selectedTemplate;
        })

        if (selectedTemplateFile.length > 0) {

            if (attachedFile == null && currentCircularListItemValue == undefined) {
                this.setState({ isLoading: true }, async () => {

                    if (circularListItem.CircularStatus == Constants.lblNew) {

                        await services.getFileContent(selectedTemplateFile[0].ServerRelativeUrl).then(async (fileContent) => {

                            /**
                            |--------------------------------------------------
                            | 
                            |--------------------------------------------------
                            */
                            //if (circularStatus == Constants.draft) {
                            circularListItem.CircularStatus = Constants.draft;
                            circularListItem.CircularCreationDate = new Date().toISOString().split('T')[0] + `T00:00:00Z`;

                            //}

                            await services.createItem(serverRelativeUrl, Constants.circularList, circularListItem).then(async (listItem) => {

                                this.addAttachmentAsBuffer(listItem, fileContent)

                            }).catch((error) => {
                                console.log(error);
                                this.setState({ isLoading: false })
                            })
                        }).catch((error) => {
                            console.log(error);
                            this.setState({ isLoading: false })
                        })
                    }

                })

            }

        }

    }

    private updateCircularItemAndFile = () => {
        let providerValue = this.context;
        const { services, serverRelativeUrl, context } = providerValue as IBobCircularRepositoryProps;
        const { attachedFile, circularListItem, currentCircularListItemValue, templateFiles, selectedTemplate } = this.state;
        let selectedTemplateFile = templateFiles.filter((val) => {
            return val.templateName == selectedTemplate;
        })

        if (currentCircularListItemValue && attachedFile == null) {

            this.setState({ isLoading: true }, async () => {
                await services.getFileContent(selectedTemplateFile[0].ServerRelativeUrl).then(async (fileContent) => {

                    let ID = parseInt(currentCircularListItemValue.ID);
                    await services.updateItem(serverRelativeUrl, Constants.circularList, ID, circularListItem).then(async (listItem) => {
                        listItem.Author = currentCircularListItemValue.Author;
                        circularListItem.CircularCreationDate = listItem?.Created;
                        this.addAttachmentAsBuffer(listItem, fileContent)

                    }).catch((error) => {
                        console.log(error);
                        this.setState({ isLoading: false })
                    })
                }).catch((error) => {
                    console.log(error);
                    this.setState({ isLoading: false })
                })
            })
        }
    }

    private addAttachmentAsBuffer = async (listItem, fileContent) => {
        const { circularListItem } = this.state
        let providerValue = this.context;
        const { services, serverRelativeUrl, context } = providerValue as IBobCircularRepositoryProps;
        let circularNumberText = circularListItem.CircularNumber;
        const { displayMode } = this.props
        let fileName = circularListItem.CircularNumber.split(':').join('_') + `.docx`; //this.getCircularNumber().split(':').join('_') + `_` + circularNumberText + `.docx`;

        await services.addListItemAttachmentAsBuffer(Constants.circularList, serverRelativeUrl, listItem.ID, fileName, fileContent).
            then(async () => {
                await services.getListDataAsStream(serverRelativeUrl, Constants.circularList, listItem.ID).then((val) => {

                    circularListItem.CircularNumber = displayMode == Constants.lblNew ? circularNumberText.replace(`${this.getCircularNumber()}`, ``) : circularListItem.CircularNumber;

                    let circularFileContent = val?.ListData?.Attachments?.Attachments.filter((val) => {
                        return val.FileName == fileName;
                    })

                    this.setState({
                        attachedFile: circularFileContent?.length > 0 ? circularFileContent[0] : null,
                        selectedFileName: fileName,
                        currentCircularListItemValue: listItem,
                        ...circularListItem
                    }, () => {
                        const { attachedFile } = this.state;
                        let documentPreviewURL = ``;
                        //interactivepreview
                        if (attachedFile != null) {
                            documentPreviewURL = `${window.location.origin}/:w:/r${context.pageContext.legacyPageContext.webServerRelativeUrl}/_layouts/15/Doc.aspx?sourcedoc=`;
                            documentPreviewURL += `${attachedFile.AttachmentId}&file=${encodeURI(attachedFile.FileName)}&action=edit&mobileredirect=true`;
                        }
                        this.setState({ documentPreviewURL, isLoading: false })
                    })
                }).catch((error) => {
                    console.log(error);
                    this.setState({ isLoading: false })
                })
            }).catch((error) => {
                console.log(error);
                this.setState({ isLoading: false })
            })
    }


    private clearAllFormFields = () => {
        this.setState({
            circularListItem: {
                CircularNumber: ``,
                CircularStatus: `New`,
                CircularType: `${Constants.limited}`,
                SubFileCode: ``,
                IssuedFor: ``,
                Category: ``,
                Classification: ``,
                Expiry: null,
                Subject: ``,
                Keywords: ``,
                Department: ``,
                Compliance: ``,
                Gist: ``,
                CommentsMaker: ``,
                CommentsChecker: ``,
                CommentsCompliance: ``,
                CircularContent: ``,
                CircularFAQ: ``,
                CircularSOP: ``,
                SupportingDocuments: ``
            },
            lblCompliance: ``,
            lblCircularType: Constants.limited,
            isBack: false,
            isDelete: false,
            isSuccess: false,
            isLoading: false,
            expiryDate: null

        });
    }

}



