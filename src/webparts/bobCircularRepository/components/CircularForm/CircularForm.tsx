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
import { ArrowCounterclockwiseRegular, ArrowLeftFilled, ArrowUpload16Regular, CalendarRegular, DeleteRegular } from '@fluentui/react-icons';
import { IBobCircularRepositoryProps } from '../IBobCircularRepositoryProps';
import { Dialog } from '@fluentui/react-components';
import { DialogContent } from '@fluentui/react';




export default class CircularForm extends React.Component<ICircularFormProps, ICircularFormState> {

    static contextType = DataContext;
    context!: React.ContextType<typeof DataContext>;

    private faqFileInput;
    private sopFileInput;
    private supportingFileInput;

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
            lblCompliance: ``,
            lblCircularType: Constants.limited,
            issuedFor: [],
            category: [],
            classification: [],
            isBack: false,
            isDelete: false,
            isLoading: false,
            isSuccess: false,
            isNewForm: true,
            isEditForm: false,
            expiryDate: null

        }

        this.faqFileInput = React.createRef();
        this.sopFileInput = React.createRef();
        this.supportingFileInput = React.createRef()
    }


    public async componentDidMount() {

        await this.fieldValues(Constants.colIssuedFor).then((val) => {
            this.setState({ issuedFor: val?.Choices ?? [] })
        });

        await this.fieldValues(Constants.colCategory).then((val) => {
            this.setState({ category: val?.Choices ?? [] })
        });

        await this.fieldValues(Constants.colClassification).then((val) => {
            this.setState({ classification: val?.Choices ?? [] })
        });

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
        const { isBack, isDelete, isLoading, isSuccess } = this.state;
        let showAlert = (isDelete || isBack);
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
                        <div className={`${styles.row}`} style={{ padding: 10 }}>
                            {this.infoHeader()}
                            {this.formSection()}

                        </div >
                    </div>
                    <div className={`${styles.column6} `} style={{ minHeight: "100vh" }}>
                        <div className={`${styles.row}`} style={{ padding: 10 }}>
                            <div className={`${styles.column12}`}>
                                <Label className={`${styles.formLabel}`} >File Name</Label>
                            </div>
                            <div className={`${styles.column12}`} style={{ display: "flex", justifyContent: "center", alignItems: "center" }}>
                                <Label className={`${styles.formLabel}`} >Attachment preview section</Label>
                            </div>
                        </div>
                    </div>
                </div>
                <div className={`${styles.row} ${styles.formFieldMarginTop} ${styles['text-center']}`}>
                    {this.saveCancelBtn()}
                </div>
                {/* <div className={`${styles.row} ${styles.formFieldMarginTop} ${styles['text-center']}`}>
                    {this.messageBarControl(`error`)}
                </div> */}

                {showAlert &&
                    this.deleteBackDialogControl(showAlert)
                }
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

        const { circularListItem, expiryDate, lblCompliance, issuedFor, category, classification, isNewForm, isEditForm } = this.state;
        let providerValue = this.context;
        const { context, isUserChecker, isUserMaker, isUserCompliance } = providerValue as IBobCircularRepositoryProps;

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
                {/* <Divider></Divider> */}
                <div className={`${styles.row}  ${styles.formFieldMarginTop}`}>
                    <div className={`${styles.column6}`}>
                        {this.textFieldControl(`${Constants.circularNumber}`, true, `${circularListItem.CircularNumber}`, false, `Field cannot be empty`, this.getCircularNumber())}
                    </div>
                    <div className={`${styles.column6}`}>
                        {this.dropDownControl(`${Constants.issuedFor}`, true, `${circularListItem.IssuedFor}`, issuedFor, false, `Field cannot be empty`)}
                    </div>
                </div>
                {/* <Divider></Divider> */}
                <div className={`${styles.row}  ${styles.formFieldMarginTop}`}>
                    <div className={`${styles.column6}`}>
                        {this.dropDownControl(`${Constants.category}`, true, `${circularListItem.Category}`, category, false, `Field cannot be empty`)}
                    </div>
                    <div className={`${styles.column6}`}>
                        {this.dropDownControl(`${Constants.classification}`, true, `${circularListItem.Classification}`, classification, false, `Field cannot be empty`)}
                    </div>
                </div>
                {/* <Divider></Divider> */}

                {/* <Divider></Divider> */}
                <div className={`${styles.row} ${styles.formFieldMarginTop}`}>
                    <div className={`${styles.column6}`}>
                        {this.textFieldControl(`${Constants.subFileNo}`, false, `${circularListItem.SubFileCode}`, false, ``)}
                    </div>
                    <div className={`${styles.column6}`}>
                        {this.textFieldControl(`${Constants.keyWords}`, false, `${circularListItem.Keywords}`, false, ``)}
                    </div>
                </div>
                {/* <Divider></Divider> */}
                <div className={`${styles.row} ${styles.formFieldMarginTop}`}>
                    <div className={`${styles.column6}`}>
                        {this.datePickerControl(`${Constants.expiry}`, expiryDate, true)}
                    </div>
                    <div className={`${styles.column6}`}>
                        {this.textFieldControl(`${Constants.department}`, false, `${circularListItem.Department}`)}
                    </div>

                </div>
                <div className={`${styles.row}  ${styles.formFieldMarginTop}`}>
                    <div className={`${styles.column6}`}>
                        {this.switchControl(`${Constants.type}`, false, `${circularListItem?.CircularType ?? ``}`)}
                    </div>
                    <div className={`${styles.column6}`}>
                        {this.switchControl(`${Constants.compliance}`, false, `${lblCompliance}`)}
                    </div>

                </div>
                <div className={`${styles.row} ${styles.formFieldMarginTop}`}>
                    <div className={`${styles.column12}`}>
                        {this.textAreaControl(`${Constants.gist}`, false, `${circularListItem.Gist}`)}
                    </div>
                </div>


                {/* <div className={`${styles.row} ${styles.formFieldMarginTop}`}>
                    {this.fileUploadControl(`${Constants.faqs}`, this.faqFileInput)}
                </div>
                <div className={`${styles.row} ${styles.formFieldMarginTop}`}>
                    {this.fileUploadControl(`${Constants.sop}`, this.sopFileInput)}
                </div>
                <div className={`${styles.row} ${styles.formFieldMarginTop}`}>
                    {this.fileUploadControl(`${Constants.supportingDocument}`, this.supportingFileInput)}
                </div> */}

                {isUserMaker && isEditForm && <div className={`${styles.row} ${styles.formFieldMarginTop}`}>
                    <div className={`${styles.column12}`}>
                        {this.textAreaControl(`${Constants.commentsMaker}`, true, `${circularListItem.CommentsMaker}`)}
                    </div>
                </div>}

                {isUserCompliance && isEditForm && <div className={`${styles.row} ${styles.formFieldMarginTop}`}>
                    <div className={`${styles.column12}`}>
                        {this.textAreaControl(`${Constants.commentsCompliance}`, true, `${circularListItem.CommentsCompliance}`)}
                    </div>
                </div>}

                {isUserChecker && isEditForm && <div className={`${styles.row} ${styles.formFieldMarginTop}`}>
                    <div className={`${styles.column12}`}>
                        {this.textAreaControl(`${Constants.commentsChecker}`, true, `${circularListItem.CommentsChecker}`)}
                    </div>
                </div>}


            </div>
        </>
        return formSectionJSX;
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
            {showDraftClearSubmitBtn &&
                <Button appearance="primary" className={`${styles.formBtn}`}
                    onClick={this.clearAllFormFields}>Clear
                </Button>
            }
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

    private textAreaControl = (labelName: string, isRequired: boolean, value: string, isDisabled?: boolean, errorMessage?: string): JSX.Element => {
        let textAreaJSX = <>
            <Field label={<Label className={`${styles.formLabel} ${styles.fieldTitle}`}>{labelName}</Label>}
                required={isRequired}
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
        let textFieldJSX = <>
            <Field label={<Label className={`${styles.formLabel} ${styles.fieldTitle}`}>{labelName}</Label>}
                required={isRequired}
                validationState={isRequired && value == "" ? "error" : "none"}
                validationMessage={isRequired && value == "" ? errorMessage : ``} >
                <Input value={value} maxLength={255} className={`${styles.formInput}`}
                    placeholder={placeholder ?? ``}
                    onChange={this.onInputChange.bind(this, labelName)}></Input>
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
        const { circularListItem } = this.state
        switch (labelName) {

            case Constants.issuedFor: circularListItem.IssuedFor = data.optionValue;
                this.setState({ circularListItem })

                break;
            case Constants.category: circularListItem.Category = data.optionValue;
                this.setState({ circularListItem });
                break;

            case Constants.classification: circularListItem.Classification = data.optionValue;
                this.setState({ circularListItem });
                break;

            default:
                break;
        }
    }

    private datePickerControl = (labelName: string, value: any, isRequired?: boolean, isDisabled?: boolean): JSX.Element => {
        let datePickerJSX = <>
            <Field
                label={<Label className={`${styles.formLabel} ${styles.fieldTitle}`}>{labelName}</Label>}
                required={isRequired}>
                {/* <Input input={{ readOnly: true, type: "date" }} root={{ style: { fontFamily: "Roboto" } }}></Input> */}

                <DatePicker mountNode={{}}
                    formatDate={this.onFormatDate}
                    value={value}
                    contentAfter={
                        <>
                            <Button icon={<ArrowCounterclockwiseRegular />}
                                appearance="transparent"
                                title="Reset"
                                onClick={this.onResetDateClick.bind(this, `${labelName}`)}>
                            </Button>
                            <Button icon={<CalendarRegular />} appearance="transparent"></Button>
                        </>}
                    onSelectDate={this.onSelectDate.bind(this, `${labelName}`)}
                    input={{ style: { fontFamily: "Roboto" } }} />


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
        dateFormat += `-` + date.getDate();
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

    private switchControl = (labelName, isRequired, switchLabel, orientation: any = "vertical", isDisabled?: boolean): JSX.Element => {
        let switchControlJSX = <>
            <Field
                label={<Label className={`${styles.formLabel} ${styles.fieldTitle}`}>{labelName}</Label>}
                orientation={orientation}
                required={isRequired}
            >
                <Switch required={isRequired}
                    onChange={this.onSwitchChange.bind(this, labelName)}
                    label={<Label className={`${styles.formLabel}`}>{switchLabel}</Label>}></Switch>
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
                    this.setState({ isLimited: false, ...circularListItem });
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

    private fileUploadControl = (labelName: string, filePickerRef: any): JSX.Element => {
        let fileUploadJSX = <>
            <div className={`${styles.column4} ${styles.formFieldMarginTop}`}>
                <input
                    id={`file-picker_${labelName}`}
                    style={{ display: "none" }}
                    type="file"
                    onChange={e => this.onFileUploadChange.bind(this, labelName)}
                    ref={filePickerRef}
                    multiple
                />
                <Field label={<Label className={`${styles.formLabel} ${styles.fieldTitle}`}>{labelName}</Label>} required={false}>

                    <Button icon={<ArrowUpload16Regular />}
                        onClick={this.onUploadClick.bind(this, labelName)}

                        iconPosition="before"
                    >
                        File Upload
                    </Button>

                </Field>
            </div>
        </>;

        return fileUploadJSX;
    }

    private onFileUploadChange = (labelName: string, e: React.ChangeEvent<HTMLInputElement>) => {

        switch (labelName) {
            case Constants.faqs: break;
            case Constants.sop: break;
            case Constants.supportingDocument: break;
            default: break;
        }

    }

    private onUploadClick = (labelName: string) => {
        switch (labelName) {
            case Constants.faqs:
                this.faqFileInput.current.value = "";
                this.faqFileInput.current.click();

                break;
            case Constants.sop:
                this.sopFileInput.current.value = "";
                this.sopFileInput.current.click();

                break;
            case Constants.supportingDocument:
                this.supportingFileInput.current.value = "";
                this.supportingFileInput.current.click();

                break;
            default:
                break;
        }
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
                                    Item Submitted successfully
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
        const { circularListItem, currentCircularListItemValue } = this.state;
        let providerValue = this.context;
        const { services, serverRelativeUrl } = providerValue as IBobCircularRepositoryProps;

        let circularNumberText = circularListItem.CircularNumber;
        circularListItem.CircularNumber = this.getCircularNumber() + ":" + circularNumberText;

        this.setState({ isLoading: true }, async () => {

            console.log(circularListItem)
            //this.setState({ isLoading: false });
            if (currentCircularListItemValue == undefined) {

                await services.createItem(serverRelativeUrl, Constants.circularList, circularListItem).then((value) => {

                    circularListItem.CircularNumber = circularNumberText;

                    this.setState({ isSuccess: true, isLoading: false, circularListItem, currentCircularListItemValue: value })
                }).catch((error) => {
                    this.setState({ isLoading: false })
                });
            }
            else {
                let ID = parseInt(currentCircularListItemValue.ID);
                let eTag = currentCircularListItemValue["odata.etag"];
                await services.updateItem(serverRelativeUrl, Constants.circularList, ID, circularListItem, eTag).then((value) => {
                    circularListItem.CircularNumber = circularNumberText;
                    this.setState({ isSuccess: true, isLoading: false, circularListItem, currentCircularListItemValue: value })
                }).catch((error) => {
                    this.setState({ isLoading: false })
                });
            }
        })
    }

    private validateAllFields = () => {

    }


}



