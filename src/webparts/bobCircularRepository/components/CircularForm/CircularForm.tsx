import * as React from 'react';
import styles from '../BobCircularRepository.module.scss';
import { ICircularFormProps } from './ICircularFormProps';
import { ICircularFormState } from './ICircularFormState';
import {
    Dropdown, Field, Image,
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
import { Add16Filled, ArrowCounterclockwiseRegular, ArrowLeftFilled, ArrowUpload16Regular, Attach16Filled, CalendarRegular, Delete16Regular, DeleteRegular, OpenRegular } from '@fluentui/react-icons';
import { IBobCircularRepositoryProps } from '../IBobCircularRepositoryProps';
import { Dialog } from '@fluentui/react-components';
import { DialogContent } from '@fluentui/react';
import { IADProperties } from '../../Models/IModel';
import { IFileInfo } from '@pnp/sp/files';
import SupportingDocument from './SupportingDocument/SupportingDocument';
import FileViewer from '../FileViewer/FileViewer';
import { error } from 'pdf-lib';




export default class CircularForm extends React.Component<ICircularFormProps, ICircularFormState> {

    static contextType = DataContext;
    context!: React.ContextType<typeof DataContext>;
    private sopFileAttachments;
    private sopFileInput;

    public constructor(props) {
        super(props)

        this.state = {
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
                CircularSOP: ``
            },
            currentCircularListItemValue: undefined,
            selectedSupportingCirculars: [],
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
            openSupportingDocument: false,
            currentItemID: -1,
            isExpiryDateDisabled: false,
            openSupportingCircularFile: false,
            supportingDocLinkItem: undefined,
            isFileSizeAlert: false,
            isFileTypeAlert: false

        }


        this.sopFileInput = React.createRef();
        this.sopFileAttachments = new Map<string, any>();
    }


    public async componentDidMount() {

        let providerValue = this.context;
        const { services, serverRelativeUrl, context } = providerValue as IBobCircularRepositoryProps;

        //context.pageContext.user.email;
        this.setState({ isLoading: true }, async () => {

            await services.filterLargeListItem(serverRelativeUrl, Constants.circularList, `CircularNumber eq 'BCC : BR : 95 :89'`).then((listItem) => {
                console.log(listItem)
            }).catch((error) => {
                console.log(error)
            })

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
            })

            //context.pageContext.user.email
            await services.getCurrentUserInformation(`Aditya.Pal@bankofbaroda.com`, Constants.adSelectedColumns).then((val: IADProperties[]) => {
                console.log(val)
            }).catch((error) => {
                console.log(error);
                this.setState({ isLoading: false })
            })

            await services.getLatestItemId(serverRelativeUrl, Constants.circularList).then((itemID) => {
                const { circularListItem } = this.state;
                circularListItem.CircularNumber = itemID.toString();
                this.setState({ circularListItem })
            }).catch((error) => {
                console.log(`Latest Item ID` + error);
                this.setState({ isLoading: false })
            })

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
        })


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
        const { isBack, isDelete, isLoading, isSuccess,
            documentPreviewURL, attachedFile,
            isFormInValid, openSupportingDocument, isDeleteCircularFile, isFileSizeAlert, isFileTypeAlert,
            openSupportingCircularFile } = this.state;
        let showAlert = (isDelete || isBack);
        let title = isFormInValid || isFileSizeAlert || isFileTypeAlert ?
            Constants.validationAlertTitle :
            isDeleteCircularFile ? `${Constants.deleteCircularTitle}` : ``;
        let message = isFormInValid ? Constants.validationAlertMessage :
            isDeleteCircularFile ? `${Constants.deleteCircularMessage}` : isFileSizeAlert ? Constants.validationAlertMessageFileSize :
                isFileTypeAlert ? Constants.validationAlertMessageFileType : ``;

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
                    <div className={`${styles.column10} ${styles.headerBackgroundColor} ${styles['text-center']}`}>
                        <Label className={`${styles.formHeader}`}>
                            {Text.format(Constants.headerCircularUpload, "New")}
                        </Label>

                    </div>
                    <div className={`${styles.column1} ${styles.headerBackgroundColor} `}>
                        <Button icon={<DeleteRegular />}
                            onClick={this.onBtnClick.bind(this, Constants.delete)}
                            style={{ float: "right" }}
                            title="Delete Circular" appearance="transparent"
                            className={`${styles.formHeader}`}></Button>
                    </div>

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
                                    {attachedFile != null && `${attachedFile.FileName}`}
                                </Label>
                            </div>
                            <div className={`${styles.column12}`} style={{ display: "flex", justifyContent: "center", alignItems: "center" }}>
                                {/* <Label className={`${styles.formLabel}`} >Attachment preview section</Label> */}
                                {documentPreviewURL != "" && <iframe
                                    src={documentPreviewURL ?? ``}
                                    style={{
                                        minHeight: 800,
                                        height: 180,
                                        width: "100%",
                                        border: 0
                                    }} role="presentation" tabIndex={-1}></iframe>}
                            </div>
                        </div>
                    </div>
                </div>
                <div className={`${styles.row} ${styles.formFieldMarginTop} ${styles['text-center']}`} style={{ borderTop: "1px solid lightgoldenrodyellow" }}>
                    {this.saveCancelBtn()}
                </div>
                {/* <div className={`${styles.row} ${styles.formFieldMarginTop} ${styles['text-center']}`}>
                    {this.messageBarControl(`error`)}
                </div> */}

                {showAlert &&
                    this.deleteBackDialogControl(showAlert)
                }
                {
                    (isFormInValid || isFileSizeAlert || isFileTypeAlert) &&
                    this.alertControl((isFormInValid || isFileSizeAlert || isFileTypeAlert), title, undefined, message, this.alertButton())
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

            </>
        )
    }

    private infoHeader = (): JSX.Element => {

        const { circularListItem } = this.state
        let providerValue = this.context;
        const { context } = providerValue as IBobCircularRepositoryProps;
        let requester = context.pageContext.user.displayName;
        let circularCreationDate = this.onFormatDate(new Date());
        let infoSectionJSX = <>

            <div className={`${styles.column12}`}>
                <div className={`${styles.row} ${styles.formRequestInfo}`}>
                    <div className={`${styles.column4}`}>
                        <Label className={`${styles.formLabel}`}>Requester : {requester}</Label>
                    </div>
                    <div className={`${styles.column4}`}>
                        <Label className={`${styles.formLabel}`}>Status : {circularListItem.CircularStatus}</Label>
                    </div>
                    <div className={`${styles.column4}`}>
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
            classification, isNewForm, isEditForm, selectedSupportingCirculars } = this.state;
        let providerValue = this.context;
        const { context, isUserChecker, isUserMaker, isUserCompliance } = providerValue as IBobCircularRepositoryProps;
        let isTypeChecked = circularListItem.CircularType == Constants.unlimited;
        let isTypeDisabled = circularListItem.Classification == Constants.lblMaster;

        let formSectionJSX = <>
            <div className={`${styles.column12}`} >
                <div className={`${styles.row} ${styles.formFieldMarginTop}`}>
                    <div className={`${styles.column12}`}>
                        {this.textAreaControl(`${Constants.subject}`, true, `${circularListItem.Subject}`, false, `Field cannot be empty`)}
                    </div>
                    {/* <div className={`${styles.column6}`}>
                        {this.avatarControl(`${Constants.circularInitator}`, `${context.pageContext.user.displayName}`)}
                    </div> */}


                </div>
                <Divider appearance="subtle" ></Divider>
                <div className={`${styles.row}  ${styles.formFieldMarginTop}`}>
                    <div className={`${styles.column6}`}>

                        {this.textFieldControl(`${Constants.circularNumber}`, true, `${circularListItem.CircularNumber}`, false, `Field cannot be empty`)}
                    </div>
                    <div className={`${styles.column6}`}>
                        {this.dropDownControl(`${Constants.issuedFor}`, true, `${circularListItem.IssuedFor}`, issuedFor, false, `Field cannot be empty`)}
                    </div>
                </div>
                <Divider appearance="subtle" ></Divider>
                <div className={`${styles.row}  ${styles.formFieldMarginTop}`}>
                    <div className={`${styles.column6}`}>
                        {this.dropDownControl(`${Constants.category}`, true, `${circularListItem.Category}`, category, false, `Field cannot be empty`)}
                    </div>
                    <div className={`${styles.column6}`}>
                        {this.dropDownControl(`${Constants.classification}`, true, `${circularListItem.Classification}`, classification, false, `Field cannot be empty`)}
                    </div>
                </div>
                <Divider appearance="subtle" ></Divider>

                {/* <Divider></Divider> */}
                <div className={`${styles.row} ${styles.formFieldMarginTop}`}>
                    <div className={`${styles.column6}`}>
                        {this.textFieldControl(`${Constants.subFileNo}`, false, `${circularListItem.SubFileCode}`, false, ``)}
                    </div>
                    <div className={`${styles.column6}`}>
                        {this.textFieldControl(`${Constants.keyWords}`, false, `${circularListItem.Keywords}`, false, ``)}
                    </div>
                </div>
                <Divider appearance="subtle" ></Divider>

                <div className={`${styles.row}  ${styles.formFieldMarginTop}`}>
                    <div className={`${styles.column6}`}>
                        {this.datePickerControl(`${Constants.expiry}`, expiryDate, !isTypeChecked, isTypeChecked)}
                    </div>
                    <div className={`${styles.column6}`}>
                        {this.switchControl(`${Constants.type}`, false, `${circularListItem?.CircularType ?? ``}`, "vertical", isTypeChecked, isTypeDisabled)}
                    </div>
                </div>
                <Divider appearance="subtle" ></Divider>
                <div className={`${styles.row} ${styles.formFieldMarginTop}`}>
                    <div className={`${styles.column6}`}>
                        {this.textFieldControl(`${Constants.department}`, false, `${circularListItem.Department}`)}
                    </div>
                    <div className={`${styles.column6}`}>
                        {this.switchControl(`${Constants.compliance}`, false, `${lblCompliance}`)}
                    </div>
                </div>
                <Divider appearance="subtle" ></Divider>
                <div className={`${styles.row} ${styles.formFieldMarginTop}`}>
                    <div className={`${styles.column6}`}>
                        {this.dropDownControl(`${Constants.lblTemplate}`, false, `${selectedTemplate}`, templates, false, `Field cannot be empty`)}

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

                        <Button appearance="primary" icon={<Add16Filled />}
                            style={{ width: "100%", padding: 5, cursor: "pointer" }}
                            onClick={() => {
                                this.setState({ openSupportingDocument: true, isLoading: true })
                            }}
                            iconPosition="before">Click here to add supporting documents</Button>
                    </div>
                </div>
                <div className={`${styles.row} ${styles.formFieldMarginTop}`}>
                    <div className={`${styles.column12}`}>
                        <div className={`${styles.row}`}>
                            {selectedSupportingCirculars && selectedSupportingCirculars.length > 0 &&
                                selectedSupportingCirculars.map((listItem) => {

                                    return <>
                                        <div className={`${styles.column3}`}>
                                            <Link
                                                className={`${styles.formLabel}`}
                                                onClick={() => {
                                                    this.openSupportingCircularFile(listItem);
                                                }}>{listItem.CircularNumber ?? ``}</Link>
                                            {/* <Label>{listItem.CircularNumber ?? ``}</Label> */}
                                            <Button
                                                icon={<Delete16Regular />}
                                                appearance="transparent"
                                                onClick={() => { this.deleteSupportingCircular(listItem) }}></Button>
                                        </div>

                                    </>
                                })}
                        </div>
                    </div>
                </div>
                <Divider appearance="subtle" ></Divider>

                <div className={`${styles.row} ${styles.formFieldMarginTop}`}>
                    <div className={`${styles.column12}`}>
                        {this.textAreaControl(`${Constants.gist}`, false, `${circularListItem.Gist}`, false, ``, `Maximum 500 words are allowed`)}
                    </div>
                </div>

                <Divider appearance="subtle" ></Divider>

                <div className={`${styles.row} ${styles.formFieldMarginTop}`}>
                    <div className={`${styles.column12}`}>
                        {this.textAreaControl(`${Constants.faqs}`, false, `${circularListItem.CircularFAQ}`)}
                    </div>
                </div>

                <Divider appearance="subtle" ></Divider>
                {
                    <div className={`${styles.row} ${styles.formFieldMarginTop}`}>
                        <div className={`${styles.column6}`}>
                            {this.fileUploadControl(`${Constants.sop}`, this.sopFileInput)}
                        </div>
                        <div className={`${styles.column6}`} style={{ padding: 10 }}>
                            {this.sopFilesControl()}
                        </div>
                    </div>
                }



                {isUserMaker && isEditForm && <>
                    <div className={`${styles.row} ${styles.formFieldMarginTop}`}>
                        <div className={`${styles.column12}`}>
                            {this.textAreaControl(`${Constants.commentsMaker}`, true, `${circularListItem.CommentsMaker}`)}
                        </div>
                    </div>
                    <Divider appearance="subtle" ></Divider>
                </>
                }

                {isUserCompliance && isEditForm && <>
                    <div className={`${styles.row} ${styles.formFieldMarginTop}`}>
                        <div className={`${styles.column12}`}>
                            {this.textAreaControl(`${Constants.commentsCompliance}`, true, `${circularListItem.CommentsCompliance}`)}
                        </div>
                    </div>
                    <Divider appearance="subtle" ></Divider>
                </>
                }

                {isUserChecker && isEditForm && <>
                    <div className={`${styles.row} ${styles.formFieldMarginTop}`}>
                        <div className={`${styles.column12}`}>
                            {this.textAreaControl(`${Constants.commentsChecker}`, true, `${circularListItem.CommentsChecker}`)}
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
        const { isUserChecker, isUserMaker, isUserCompliance } = providerValue as IBobCircularRepositoryProps;
        const { circularListItem, isEditForm, isNewForm } = this.state
        let submtStatus = isUserMaker && circularListItem.Compliance == Constants.lblYes ? Constants.sbmtCompliance : Constants.sbmtChecker;
        let showDraftClearSubmitBtn = isNewForm && (circularListItem.CircularStatus == "New" || circularListItem.CircularStatus == Constants.draft);
        let showReturnToMakerBtn = circularListItem.CircularStatus == Constants.sbmtChecker || circularListItem.CircularStatus == Constants.sbmtCompliance;
        let showSbmtCheckerBtn = circularListItem.CircularStatus == Constants.sbmtChecker;
        let returnStatus = showReturnToMakerBtn && isUserCompliance ? Constants.cmmtCompliance : Constants.commentsChecker;

        let saveCancelBtnJSX = <>
            {/* {showDraftClearSubmitBtn &&
                <Button appearance="primary" className={`${styles.formBtn}`}
                    onClick={this.clearAllFormFields}>Clear
                </Button>
            } */}
            {showDraftClearSubmitBtn &&
                <Button appearance="primary"
                    className={`${styles.formBtn}`}
                    onClick={this.saveForm.bind(this, Constants.draft)}>
                    Save as Draft
                </Button>
            }
            {showDraftClearSubmitBtn &&
                <Button appearance="primary"
                    className={`${styles.formBtn}`}
                    onClick={this.saveForm.bind(this, submtStatus)}>
                    Submit
                </Button>
            }
            {(isUserCompliance || isUserChecker) && showReturnToMakerBtn &&
                <Button
                    appearance="primary"
                    onClick={this.saveForm.bind(this, returnStatus)}
                    className={`${styles.formBtn}`}>
                    Return to maker
                </Button>
            }
            {
                isUserCompliance && showSbmtCheckerBtn &&
                <Button appearance="primary"
                    onClick={this.saveForm.bind(this, Constants.sbmtChecker)}
                    className={`${styles.formBtn}`}>
                    Submit to Checker
                </Button>

            }
            {isUserChecker && showSbmtCheckerBtn &&
                <Button appearance="primary"
                    className={`${styles.formBtn}`}>
                    Publish
                </Button>
            }
        </>;

        return saveCancelBtnJSX;
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
                    root={{ className: `${styles.formLabel}` }}
                    resize="vertical" onChange={this.onTextAreaChange.bind(this, labelName)}></Textarea>
            </Field>
        </>;

        return textAreaJSX;
    }

    private onTextAreaChange = (labelName: string, ev: React.ChangeEvent<HTMLTextAreaElement>, data: TextareaOnChangeData) => {
        const { circularListItem } = this.state
        switch (labelName) {
            case Constants.subject: circularListItem.Subject = data.value;
                this.setState({ circularListItem });

                break;
            case Constants.gist: circularListItem.Gist = data.value;
                this.setState({ circularListItem })
                break;
            case Constants.faqs: circularListItem.CircularFAQ = data.value;
                this.setState({ circularListItem });
                break;
            case Constants.commentsMaker: circularListItem.CommentsMaker = data.value;
                this.setState({ circularListItem })
                break;
            case Constants.commentsChecker: circularListItem.CommentsChecker = data.value;
                this.setState({ circularListItem })
                break;
            case Constants.commentsCompliance: circularListItem.CommentsCompliance = data.value;
                this.setState({ circularListItem })
                break;

            default:
                break;
        }
    }

    private textFieldControl = (labelName: string, isRequired: boolean, value: string, isDisabled?: boolean, errorMessage?: string, placeholder?: string): JSX.Element => {
        let columnClassLabel = labelName == Constants.circularNumber ? `${styles.column3}` : ``;
        let columnClassInput = labelName == Constants.circularNumber ? `${styles.column9}` : `${styles.column12}`

        let textFieldJSX = <>
            <Field label={<Label className={`${styles.formLabel} ${styles.fieldTitle}`}>{labelName}</Label>}
                required={isRequired}
                validationState={isRequired && value == "" ? "error" : "none"}
                validationMessage={isRequired && value == "" ? errorMessage : ``} >
                <div className={`${styles.row}`}>
                    {labelName == Constants.circularNumber &&
                        <div className={`${columnClassLabel}`} style={{ marginTop: 5 }}>
                            <Label className={`${styles.formLabel} ${styles.fieldTitle}`} style={{ fontWeight: 400 }}>
                                {this.getCircularNumber()}
                            </Label>
                        </div>
                    }
                    <div className={`${columnClassInput}`}>
                        <Input value={value} maxLength={255}
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
            case Constants.circularNumber: circularListItem.CircularNumber = data.value;
                this.setState({ circularListItem })
                break;
            case Constants.subFileNo: circularListItem.SubFileCode = data.value;
                this.setState({ circularListItem });
                break;
            case Constants.keyWords: circularListItem.Keywords = data.value;
                this.setState({ circularListItem });
                break;
        }
    }

    private getCircularNumber = (): string => {
        let currentDate = new Date();

        let circularNumber = Text.format(Constants.circularNo, (currentDate.getFullYear() - 1908))
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

    private dropDownControl = (labelName: string, isRequired: boolean, value: string, options: any[], isDisabled?: boolean, errorMessage?: string): JSX.Element => {
        let dropDownControlJSX = <>
            <Field
                label={<Label className={`${styles.formLabel} ${styles.fieldTitle}`}>{labelName}</Label>}
                required={isRequired}
                validationState={isRequired && value == "" ? "error" : "none"}
                validationMessage={isRequired && value == "" ? errorMessage : ``}
            >
                <Dropdown mountNode={{}} placeholder={`Select ${labelName}`} value={value}
                    selectedOptions={[value]}
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
                        this.setState({ selectedTemplate: data.optionValue }, async () => {
                            this.createUpdateCircularFile();
                        })
                    }
                    else if (attachedFile != null && isFormValid) {
                        let selectedTemplateVal = attachedFile != null ? selectedTemplate : ``;
                        this.setState({ selectedTemplate: selectedTemplateVal })
                    }

                }
                else {
                    this.setState({ isFormInValid: true })
                }
            default:
                break;
        }
    }


    private createUpdateCircularFile = () => {
        const { circularListItem } = this.state;
        let circularNumberText = circularListItem.CircularNumber;
        let circularNumberIndexOf = circularListItem.CircularNumber.indexOf(`${this.getCircularNumber()}`);

        // if BOB:BR:116: not present then circular Number will be this
        if (circularNumberIndexOf == -1) {
            circularListItem.CircularNumber = `${this.getCircularNumber()}` + `${circularNumberText}`
        }
        if (circularListItem.CircularStatus == Constants.lblNew) {
            this.addCircularItemAndFile();
        }
        else {
            this.updateCircularItemAndFile()
        }
    }

    private addCircularItemAndFile = () => {

        let providerValue = this.context;
        const { services, serverRelativeUrl, context } = providerValue as IBobCircularRepositoryProps;
        const { templateFiles, selectedTemplate, currentCircularListItemValue, attachedFile, circularListItem } = this.state;
        let selectedTemplateFile = templateFiles.filter((val) => {
            return val.templateName == selectedTemplate;
        })

        if (selectedTemplateFile.length > 0) {

            if (attachedFile == null && currentCircularListItemValue == undefined) {
                this.setState({ isLoading: true }, async () => {

                    if (circularListItem.CircularStatus == Constants.lblNew) {

                        await services.getFileContent(selectedTemplateFile[0].ServerRelativeUrl).then(async (fileContent) => {
                            //Set Circular Status as Draft
                            circularListItem.CircularStatus = Constants.draft;

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

        if (currentCircularListItemValue != undefined && attachedFile == null) {

            this.setState({ isLoading: true }, async () => {
                await services.getFileContent(selectedTemplateFile[0].ServerRelativeUrl).then(async (fileContent) => {

                    let ID = parseInt(currentCircularListItemValue.ID);

                    await services.updateItem(serverRelativeUrl, Constants.circularList, ID, circularListItem).then(async (listItem) => {

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

        let fileName = circularListItem.CircularNumber.split(':').join('_') + `.docx`; //this.getCircularNumber().split(':').join('_') + `_` + circularNumberText + `.docx`;

        await services.addListItemAttachmentAsBuffer(Constants.circularList, serverRelativeUrl, listItem.ID, fileName, fileContent).
            then(async () => {
                await services.getListDataAsStream(serverRelativeUrl, Constants.circularList, listItem.ID).then((val) => {
                    //circularListItem.CircularNumber = circularNumberText;
                    circularListItem.CircularNumber = circularNumberText.replace(`${this.getCircularNumber()}`, ``);
                    this.setState({
                        attachedFile: val.ListData.Attachments.Attachments[0],
                        currentCircularListItemValue: listItem,
                        ...circularListItem
                    }, () => {
                        const { attachedFile } = this.state;
                        //interactivepreview
                        let documentPreviewURL = `${window.location.origin}/:w:/r${context.pageContext.legacyPageContext.webServerRelativeUrl}/_layouts/15/Doc.aspx?sourcedoc=`;
                        documentPreviewURL += `${attachedFile.AttachmentId}&file=${encodeURI(attachedFile.FileName)}&action=edit&mobileredirect=true`;

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
    /**
    |--------------------------------------------------
    | This attachment Link is Circular Content File
    |--------------------------------------------------
    */
    private attachmentLink = (selectedFile): JSX.Element => {

        let attachedLinkJSX = <div className={`${styles.row}`}>
            <div className={`${styles.column12}`}>
                <Attach16Filled></Attach16Filled>
                <Link
                    //onClick={this.openDocument.bind(this, file.ServerRelativeUrl)}
                    style={{
                        wordBreak: "break-all",
                        padding: 5
                    }}
                >{`${selectedFile.FileName}`}</Link>
                <Button icon={<Delete16Regular></Delete16Regular>} style={{ marginLeft: 5 }}
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
                label={<Label className={`${styles.formLabel} ${styles.fieldTitle}`}>{labelName}</Label>}
                orientation={orientation}
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

    private fileUploadControl = (labelName: string, filePickerRef: any): JSX.Element => {
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
                iconPosition="before"
            > Upload SOP File
            </Button>
            <Field label={
                <Label className={`${styles.formLabel} `}>
                    {` (Maximum 5MB .pdf & .docx file allowed.)`}
                </Label>
            }
                required={false}>
            </Field>

        </>;

        return fileUploadJSX;
    }

    private onFileUploadChange = (labelName: string, e: React.ChangeEvent<HTMLInputElement>) => {
        const files = e.target.files;
        let invalidFileSize = [];
        let inValidFileType = [];

        if (files) {

            for (let i = 0; i < files.length; i++) {

                if (files[i].name.indexOf('.docx') > -1 || files[i].name.indexOf('.pdf') > -1) {

                    let sizeInMB = Math.round((files[i].size) / 1024);
                    if (sizeInMB <= 5120) {

                        if ((this.sopFileAttachments.has(files[i].name))) {
                            this.sopFileAttachments.delete(files[i].name);
                            this.sopFileAttachments.set(files[i].name, files[i]);
                        }
                        else {
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

            this.setState({ sopUploads: this.sopFileAttachments }, () => {
                const { sopUploads } = this.state
                let attachmentColl = [];
                let i = 0;

                sopUploads.forEach(async (value, key) => {
                    value.index = i;
                    attachmentColl.push(value);
                    i++;
                });
                this.setState({ sopAttachmentColl: attachmentColl })
            })
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

    private sopFilesControl = (): JSX.Element => {

        const { sopAttachmentColl } = this.state;

        let sopFileUploadJSX = <>
            {
                sopAttachmentColl && sopAttachmentColl.length > 0 &&

                sopAttachmentColl.map((file) => {
                    const fileName = file.name;

                    return <div className={`${styles.column12}`} style={{ marginBottom: 5 }}>
                        <Attach16Filled></Attach16Filled>
                        <Link
                            //onClick={this.openDocument.bind(this, file.ServerRelativeUrl)}
                            style={{
                                wordBreak: "break-all",
                                padding: 5
                            }}
                        >{fileName}</Link>
                        <Button icon={<Delete16Regular></Delete16Regular>} style={{ marginLeft: 5 }}
                            onClick={() => { this.deleteSOPUploadedFiles(fileName) }}></Button>
                    </div>
                })


            }
        </>

        return sopFileUploadJSX;
    }

    private deleteSOPUploadedFiles = (fileName) => {
        if (this.sopFileAttachments.has(fileName)) {
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

        const { isFormInValid, isDeleteCircularFile, isFileSizeAlert, isFileTypeAlert } = this.state
        let alertButtonJSX;

        if (isFormInValid || isFileSizeAlert || isFileTypeAlert) {
            alertButtonJSX = <div className={`${styles.row}`}>
                <div className={`${styles.column12}`}>
                    <Button appearance="secondary"
                        onClick={() => {
                            this.setState({
                                isFormInValid: false,
                                isDeleteCircularFile: false,
                                isFileSizeAlert: false,
                                isFileTypeAlert: false
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
        const { currentCircularListItemValue, attachedFile } = this.state
        const { services, serverRelativeUrl } = providerValue as IBobCircularRepositoryProps;

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
                    <DialogBody >
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
                                    onClick={() => { this.setState({ isSuccess: false }) }} >OK</Button>
                            </DialogTrigger>
                        </DialogActions>
                        }
                    </DialogBody>
                </DialogSurface>
            </Dialog>
        </>;
        return workingJSX;
    }

    private validateAllRequiredFields = (): boolean => {
        const { circularListItem } = this.state;
        let isValid = true;
        if (circularListItem.Subject == "" || circularListItem.CircularNumber == "" || circularListItem.IssuedFor == "" ||
            circularListItem.Category == "" || circularListItem.Classification == "") {

            isValid = false
        }
        else if (circularListItem.CircularType == Constants.limited) {
            isValid = !(circularListItem.Expiry == null)
        }

        return isValid
    }


    private filterPanelSupportingDocument = (): JSX.Element => {

        const { selectedSupportingCirculars } = this.state
        let providerValue = this.context as IBobCircularRepositoryProps;

        let panelSupportingDocumentsJSX = <>
            <SupportingDocument department={``}
                providerValue={providerValue}
                selectedSupportingCirculars={selectedSupportingCirculars}
                onDismiss={(supportingCirculars) => {
                    this.setState({ openSupportingDocument: false, selectedSupportingCirculars: supportingCirculars })
                }}
                completeLoading={() => { this.setState({ isLoading: false }) }}
            />
        </>
        return panelSupportingDocumentsJSX;
    }

    private deleteSupportingCircular = (listItem) => {
        const { selectedSupportingCirculars } = this.state;
        let index = selectedSupportingCirculars.indexOf(listItem);
        if (index > -1) {
            //selectedSupportingCirculars.splice(index, 1);
            delete selectedSupportingCirculars[index];
            this.setState({ selectedSupportingCirculars })
        }
    }

    private messageBarControl = (intent): JSX.Element => {
        let messageBarJSX = <>
            <MessageBar key={intent} intent={intent}>
                <MessageBarBody>
                    <Label className={`${styles.formLabel} ${styles.fieldTitle}`}>Please input all fields mark as <b>*</b></Label>
                </MessageBarBody>
            </MessageBar>
        </>;

        return messageBarJSX;
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
                CircularSOP: ``
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

    private saveForm = (status?: string) => {
        const { circularListItem, currentCircularListItemValue, isNewForm } = this.state;
        let isFormValid = this.validateAllRequiredFields();
        if (isFormValid) {

            let providerValue = this.context;
            const { services, serverRelativeUrl } = providerValue as IBobCircularRepositoryProps;

            let circularNumberText = circularListItem.CircularNumber;
            let circularNumberIndexOf = circularListItem.CircularNumber.indexOf(`${this.getCircularNumber()}`);
            // if BOB:BR:116: not present then circular Number will be this
            if (circularNumberIndexOf == -1) {
                circularListItem.CircularNumber = `${this.getCircularNumber()}` + `${circularNumberText}`;
            }

            /**
            |--------------------------------------------------
            | If form is new Mode
            |--------------------------------------------------
            */
            if (isNewForm) {
                circularListItem.CircularStatus = Constants.draft;
            }
            this.setState({ isLoading: true }, async () => {

                console.log(circularListItem)
                //this.setState({ isLoading: false });
                if (currentCircularListItemValue == undefined) {


                    // circularListItem.CircularNumber = this.getCircularNumber() + ":" + circularNumberText;

                    await services.createItem(serverRelativeUrl, Constants.circularList, circularListItem).then(async (value) => {

                        circularListItem.CircularNumber = circularNumberText.replace(`${this.getCircularNumber()}`, ``);
                        //circularListItem.CircularNumber = circularNumberText;

                        console.log(value)

                        this.setState({ isSuccess: true, isLoading: false, circularListItem, currentCircularListItemValue: value })
                    }).catch((error) => {
                        this.setState({ isLoading: false })
                    });
                }
                else {
                    let ID = parseInt(currentCircularListItemValue.ID);
                    // circularListItem.CircularNumber = this.getCircularNumber() + ":" + circularListItem.CircularNumber;
                    //let eTag = currentCircularListItemValue["odata.etag"];
                    await services.updateItem(serverRelativeUrl, Constants.circularList, ID, circularListItem).then((value) => {
                        // circularListItem.CircularNumber = circularNumberText;
                        circularListItem.CircularNumber = circularNumberText.replace(`${this.getCircularNumber()}`, ``);
                        console.log(value)
                        this.setState({ isSuccess: true, isLoading: false, circularListItem, currentCircularListItemValue: value })
                    }).catch((error) => {
                        console.log(error);
                        this.setState({ isLoading: false })
                    });
                }
            })
        }

        else {
            this.setState({ isFormInValid: true })
        }

    }

}



