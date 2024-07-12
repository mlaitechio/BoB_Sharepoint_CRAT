import * as React from 'react'
import { IEditDashBoardProps } from './IEditDashBoardProps'
import { IEditDashBoardState } from './IEditDashBoardState'
import { Constants } from '../../Constants/Constants'
import {
    Button, Dialog, DialogActions, DialogBody, DialogContent,
    DialogSurface, DialogTitle, Divider, Label, Link, Spinner,
    Table, TableBody, TableCell, TableCellLayout, TableHeader,
    TableHeaderCell, TableRow
} from '@fluentui/react-components';
import styles1 from '../BobCircularRepository.module.scss';
import { ICircularListItem } from '../../Models/IModel';
import { ChevronDownRegular, ChevronUpRegular, Delete12Regular, Delete16Regular, DeleteRegular, Edit12Regular, Edit16Regular, EditRegular, EyeRegular, OpenRegular } from '@fluentui/react-icons';
import { AnimationClassNames } from '@fluentui/react';
import { IBobCircularRepositoryProps } from '../IBobCircularRepositoryProps';
import { DataContext } from '../../DataContext/DataContext';
import { error } from 'pdf-lib';
import FileViewer from '../FileViewer/FileViewer';
import { Text } from '@microsoft/sp-core-library';
import CircularForm from '../CircularForm/CircularForm';


export default class EditDashBoard extends React.Component<IEditDashBoardProps, IEditDashBoardState> {

    static contextType = DataContext;
    context!: React.ContextType<typeof DataContext>;

    constructor(props) {
        super(props)

        this.state = {
            listItems: [], // All List Filter Items
            accordionFields: {
                isSummarySelected: false,
                isTypeSelected: false,
                isCategorySelected: false,
                isSupportingDocuments: false
            },
            isLoading: false,
            openSupportingDoc: false,
            loadDashBoard: false,
            loadEditForm: false,
            loadViewForm: false,
            showDeleteDialog: false,
            supportingDocItem: null // Current Selected Item Supporting Doc Item object not array, Object will be (result.ListData)
        }
    }

    public async componentDidMount() {
        this.onEditDashBoardLoad();
    }

    public componentDidUpdate(prevProps: Readonly<IEditDashBoardProps>, prevState: Readonly<IEditDashBoardState>, snapshot?: any): void {
        if (prevProps.stateKey != this.props.stateKey) {
            this.onEditDashBoardLoad()
        }
    }




    render() {
        const { isLoading,
            openSupportingDoc,
            supportingDocItem, loadDashBoard,
            currentSelectedItem,
            loadEditForm, loadViewForm, editFormItem, showDeleteDialog } = this.state;
        let providerValue = this.context;
        const { context } = providerValue as IBobCircularRepositoryProps;
        return (
            <>
                {
                    isLoading && this.workingOnIt()
                }
                {
                    loadDashBoard && this.circularResults()
                }

                {
                    openSupportingDoc && <FileViewer
                        listItem={supportingDocItem}
                        documentLoaded={() => { this.setState({ isLoading: false }) }}
                        onClose={() => { this.setState({ openSupportingDoc: false }) }}
                        context={context}
                    />
                }
                {
                    loadEditForm &&
                    <CircularForm
                        editFormItem={editFormItem}
                        displayMode={Constants.lblEditCircular}
                        onGoBack={() => {
                            this.setState({ loadEditForm: false, loadViewForm: false }, () => {
                                this.onEditDashBoardLoad()
                            })
                        }} />
                }
                {
                    loadViewForm &&
                    <CircularForm
                        editFormItem={editFormItem}
                        displayMode={Constants.lblViewCircular}
                        onGoBack={() => {
                            this.setState({ loadEditForm: false, loadViewForm: false }, () => {
                                this.onEditDashBoardLoad()
                            })
                        }} />
                }

                {showDeleteDialog && this.deleteDialog(showDeleteDialog, currentSelectedItem)}

            </>
        )
    }

    private onEditDashBoardLoad = () => {
        let providerValue = this.context;
        const { services, serverRelativeUrl } = providerValue as IBobCircularRepositoryProps;
        const { filterString } = this.props

        this.setState({ isLoading: true }, async () => {
            await services.filterLargeListItem(serverRelativeUrl, Constants.circularList, `${filterString}`).then(async (itemIDColl: any[]) => {
                let allListItems = await Promise.all(itemIDColl?.map(async (item) => {
                    return await services.getListDataAsStream(serverRelativeUrl, Constants.circularList, item.ID).then((listItem) => {
                        listItem.ListData.ID = item.ID;
                        return listItem?.ListData ?? []
                    }).catch((error) => {
                        console.log("Error:" + error);
                        return []
                    })
                }))
                this.setState({
                    listItems: allListItems.sort((a, b) => a.ID > b.ID ? -1 : 1),
                    loadDashBoard: true,
                    loadEditForm: false,
                    loadViewForm: false,
                    isLoading: false,

                })
            }).catch((error) => {
                console.log(error);
                this.setState({ isLoading: false })
            })
        })
    }

    private circularResults = () => {

        let providerValue = this.context;
        const { isUserChecker, isUserCompliance, isUserMaker } = providerValue as IBobCircularRepositoryProps;

        const { listItems, accordionFields, currentSelectedItem, currentSelectedItemId } = this.state;
        const columns = [
            { columnKey: "Title", label: "Document Title" },
            { coluumnKey: "ID", label: "ID" },
            { columnKey: "Date", label: "Created Date" },
            { columnKey: "Status", label: "Circular Status" },
            { columnKey: "Edit", label: "" },
        ];


        let circularResultJSX = <>
            <div className={`${styles1.row}`} >
                <div className={`${styles1.column12} ${styles1.headerBackgroundColor}`} style={{ textAlign: "center" }} >
                    {listItems && listItems.length > 0 &&
                        <Label style={{
                            fontFamily: "Roboto",
                            padding: 10,
                            cursor: "pointer",
                            fontSize: "var(--fontSizeBase500)",
                            fontWeight: "var(--fontWeightSemibold)",
                            lineHeight: "var(--lineHeightBase500)",
                            color: "white",

                        }}> {`CIRCULAR DASHBOARD`}
                        </Label>}

                </div>
                {/* <div className={`${styles1.row}`}>
                    <div className={`${styles1.column12}`} style={{ paddingLeft: 20, paddingRight: 20 }}>
                        {listItems && listItems.length > 0 && <Label style={{
                            fontFamily: "Roboto",
                            padding: 10,
                            cursor: "pointer",
                            fontSize: "var(--fontSizeBase500)",
                            fontWeight: "var(--fontWeightSemibold)",
                            lineHeight: "var(--lineHeightBase500)",
                        }}> {`Circulars (${listItems.length})`}
                        </Label>}
                    </div>
                </div> */}
                <Divider appearance="subtle"></Divider>
                <div className={`${styles1.column12}`} style={{ paddingLeft: 20, paddingRight: 20 }}>
                    <Table arial-label="Default table">
                        <TableHeader>
                            <TableRow >

                                {columns.map((column, index) => (
                                    <TableHeaderCell
                                        key={column.columnKey} colSpan={index == 0 ? 6 : 1}
                                        style={index == 0 ? { paddingLeft: 15 } : {}}
                                        className={`${styles1.fontWeightBold}`}>
                                        {column.label}
                                    </TableHeaderCell>
                                ))}
                            </TableRow>
                        </TableHeader>
                        <TableBody>
                            {listItems && listItems.length > 0 && listItems.map((val: ICircularListItem, index) => {

                                let isFieldSelected = (accordionFields.isSummarySelected || accordionFields.isTypeSelected || accordionFields.isCategorySelected || accordionFields.isSupportingDocuments);
                                let isCurrentItem = currentSelectedItemId == val.ID;
                                let tableRowClass = isFieldSelected && isCurrentItem ? `${styles1.tableRow}` : ``;
                                let isEditButtonVisible = val.CircularStatus == Constants.draft ||
                                    val.CircularStatus == Constants.cmmtChecker
                                    || val.CircularStatus == Constants.cmmtCompliance;

                                return <>
                                    <TableRow className={`${styles1.tableRow}`}>
                                        <TableCell colSpan={6} >
                                            <TableCellLayout className={`${styles1.verticalSpacing}`} style={{ padding: 5 }}>
                                                <div
                                                    className={`${styles1.colorLabel}`}
                                                    style={{
                                                        color: val.Classification == "Master" ? "#f26522" : "#162B75"
                                                    }}>{val.CircularNumber}</div>
                                                <div className={`${styles1.verticalSpacing}`}>
                                                    <Button
                                                        style={{
                                                            padding: 0, fontWeight: 400,
                                                            justifyContent: "flex-start",
                                                            alignItems: "flex-start"
                                                        }}
                                                        appearance="transparent"
                                                        onClick={this.onDetailItemClick.bind(this, val, Constants.colSubject)}>
                                                        <div style={{
                                                            textAlign: "left",
                                                            marginTop: 5,
                                                            color: val.Classification == "Master" ? "#f26522" : "#162B75"
                                                        }}>{val.Subject} </div>
                                                        {/* <OpenRegular /> */}
                                                    </Button>
                                                </div>
                                            </TableCellLayout>
                                        </TableCell>
                                        <TableCell>
                                            <TableCellLayout>
                                                {val.ID != "" ? val.ID : ``}
                                            </TableCellLayout>
                                        </TableCell>
                                        <TableCell>
                                            <TableCellLayout>
                                                {val.Created != "" ? this.formatDate(val.Created) : ``}
                                            </TableCellLayout>
                                        </TableCell>
                                        <TableCell>
                                            <TableCellLayout>
                                                {val.CircularStatus ? val.CircularStatus : ``}
                                            </TableCellLayout>
                                        </TableCell>
                                        <TableCell colSpan={1}>
                                            <TableCellLayout className={`${styles1.verticalSpacing}`}>
                                                {!isEditButtonVisible &&
                                                    <Button onClick={() => { this.viewCircular(val) }}
                                                        icon={<EyeRegular />}
                                                        style={{ marginRight: 5 }} />
                                                }

                                                {isUserMaker && isEditButtonVisible && <>
                                                    <Button onClick={() => { this.editCircular(val) }}
                                                        icon={<EditRegular />}
                                                        style={{ marginRight: 5 }} />

                                                    {/* Delete icon to be visible only for draft status | val.CircularStatus == Constants.draft && |*/}
                                                    {val.CircularStatus == Constants.draft &&
                                                        < Button icon={<DeleteRegular />}
                                                            onClick={() => {
                                                                this.setState({ showDeleteDialog: true, currentSelectedItem: val })
                                                            }}
                                                        />
                                                    }
                                                </>}

                                            </TableCellLayout>
                                        </TableCell>

                                    </TableRow >
                                    <TableRow className={`${tableRowClass}`}>

                                        <TableCell colSpan={6}>
                                            <div className={`${styles1.row}`}>
                                                <div className={`${styles1.column2}`} style={{ paddingLeft: "0px" }}>
                                                    <Button icon={accordionFields.isSummarySelected && isCurrentItem ? <ChevronUpRegular /> : <ChevronDownRegular />}
                                                        iconPosition="after"
                                                        className={accordionFields.isSummarySelected && isCurrentItem ? styles1.colorLabel : ``}
                                                        appearance={accordionFields.isSummarySelected && isCurrentItem ? "outline" : "transparent"}
                                                        onClick={this.onDetailItemClick.bind(this, val, Constants.colSummary)}>Summary</Button>
                                                </div>
                                                <div className={`${styles1.column2}`}>
                                                    <Button icon={accordionFields.isTypeSelected && isCurrentItem ? <ChevronUpRegular /> : <ChevronDownRegular />}
                                                        iconPosition="after"
                                                        className={accordionFields.isTypeSelected && isCurrentItem ? styles1.colorLabel : ``}
                                                        appearance={accordionFields.isTypeSelected && isCurrentItem ? "outline" : "transparent"}
                                                        onClick={this.onDetailItemClick.bind(this, val, Constants.colType)}>Type</Button>
                                                </div>
                                                <div className={`${styles1.column2}`}>
                                                    <Button
                                                        icon={accordionFields.isCategorySelected && isCurrentItem ? <ChevronUpRegular /> : <ChevronDownRegular />}
                                                        iconPosition="after"
                                                        className={accordionFields.isCategorySelected && isCurrentItem ? styles1.colorLabel : ``}
                                                        appearance={accordionFields.isCategorySelected && isCurrentItem ? "outline" : "transparent"}
                                                        onClick={this.onDetailItemClick.bind(this, val, Constants.colCategory)}>Category</Button>
                                                </div>
                                                <div className={`${styles1.column4}`}>
                                                    <Button
                                                        icon={accordionFields.isSupportingDocuments && isCurrentItem ? <ChevronUpRegular /> : <ChevronDownRegular />}
                                                        iconPosition="after"
                                                        className={accordionFields.isSupportingDocuments && isCurrentItem ? styles1.colorLabel : ``}
                                                        appearance={accordionFields.isSupportingDocuments && isCurrentItem ? "outline" : "transparent"}
                                                        onClick={this.onDetailItemClick.bind(this, val, Constants.colSupportingDoc)}>Supporting Documents</Button>
                                                </div>

                                            </div>
                                        </TableCell>
                                    </TableRow>
                                    {
                                        isFieldSelected && currentSelectedItemId == val.ID &&
                                        <TableRow >
                                            <TableCell colSpan={6}>
                                                <div className={`${styles1.row}`}>
                                                    <div className={`${styles1.column12} ${AnimationClassNames.slideDownIn20}`} style={{ padding: 10 }}>
                                                        {accordionFields.isSummarySelected &&
                                                            <>{`${currentSelectedItem?.Gist ?? ``}`}</>
                                                        }
                                                        {accordionFields.isTypeSelected &&
                                                            <>{currentSelectedItem?.CircularType ?? ``}</>}
                                                        {accordionFields.isCategorySelected &&
                                                            <>{currentSelectedItem?.Category ?? ``}</>}
                                                        {accordionFields.isSupportingDocuments && <>
                                                            {currentSelectedItem?.SupportingDocuments && currentSelectedItem?.SupportingDocuments != ""
                                                                ? this.supportingDocument(currentSelectedItem.SupportingDocuments) : ``}
                                                        </>}
                                                    </div>
                                                </div>
                                            </TableCell>
                                        </TableRow>
                                    }

                                </>
                            })}

                        </TableBody>
                    </Table>
                </div>
                <div className={`${styles1.column12}`}>
                    {
                        listItems && listItems.length == 0 && this.noItemFound()
                    }
                </div>
            </div>
        </>;

        return circularResultJSX;
    }


    private editCircular = (selectedItem) => {
        this.setState({
            loadDashBoard: false,
            editFormItem: selectedItem, loadEditForm: true, loadViewForm: false
        });
    }

    private viewCircular = (selectedItem) => {
        this.setState({
            loadDashBoard: false,
            editFormItem: selectedItem, loadEditForm: false, loadViewForm: true
        })
    }

    private supportingDocument = (supportingCirculars): JSX.Element => {
        let jsonSupportingCirculars: any[] = JSON.parse(supportingCirculars);

        let supportingDOCJSX = <div className={styles1.row}>
            {
                jsonSupportingCirculars && jsonSupportingCirculars.length > 0 && jsonSupportingCirculars.map((circular) => {
                    return <div className={`${styles1.column12}`} style={{ padding: 5 }}>
                        <Link onClick={() => { this.onSupportingDocClick(circular) }}>{circular.CircularNumber} </Link>
                    </div>
                })
            }
        </div>

        return supportingDOCJSX;
    }

    private noItemFound = (): JSX.Element => {
        let noItemFoundJSX = <>

            <div className={`${styles1.OneUpError} `}>
                <div className={`${styles1.odError} `}>
                    <div className={`${styles1.odErrorTitle} `}>No Circulars Found</div>
                </div>

            </div>
        </>

        return noItemFoundJSX;
    }

    private onDetailItemClick = (item: ICircularListItem, fieldName: string) => {

        const { currentSelectedItemId } = this.state;
        const { accordionFields } = this.state;
        let isCurrentItem = currentSelectedItemId == item.ID;

        switch (fieldName) {
            case Constants.colSummary:

                this.setState({
                    accordionFields: {
                        isSummarySelected: isCurrentItem ? !accordionFields.isSummarySelected : true,
                        isTypeSelected: false,
                        isCategorySelected: false,
                        isSupportingDocuments: false
                    },
                    currentSelectedItem: item,
                    currentSelectedItemId: item.ID
                }, () => {
                    // this.readItemsAsStream(item)
                })

                break;

            case Constants.colType:
                this.setState({
                    accordionFields: {
                        isSummarySelected: false,
                        isTypeSelected: isCurrentItem ? !accordionFields.isTypeSelected : true,
                        isCategorySelected: false,
                        isSupportingDocuments: false
                    },
                    currentSelectedItem: item,
                    currentSelectedItemId: item.ID
                }, () => {
                    //this.readItemsAsStream(item)
                })
                break;
            case Constants.colCategory:
                this.setState({
                    accordionFields: {
                        isSummarySelected: false,
                        isTypeSelected: false,
                        isCategorySelected: isCurrentItem ? !accordionFields.isCategorySelected : true,
                        isSupportingDocuments: false
                    },
                    currentSelectedItem: item,
                    currentSelectedItemId: item.ID
                }, () => {
                    //this.readItemsAsStream(item)
                })
                break;

            case Constants.colSupportingDoc: this.setState({
                accordionFields: {
                    isSummarySelected: false,
                    isTypeSelected: false,
                    isCategorySelected: false,
                    isSupportingDocuments: isCurrentItem ? !accordionFields.isSupportingDocuments : true
                },
                currentSelectedItem: item,
                currentSelectedItemId: item.ID
            }, () => {
                //this.readItemsAsStream(item)
            })
                break;
        }

    }

    private onSupportingDocClick = (supportingCircular) => {
        this.setState({ isLoading: true }, async () => {
            const { currentSelectedItem } = this.state
            let providerValue = this.context;
            const { services, serverRelativeUrl } = providerValue as IBobCircularRepositoryProps;

            await services.getListDataAsStream(serverRelativeUrl, Constants.circularList, parseInt(supportingCircular.ID)).
                then((linkItem) => {
                    linkItem.ListData.ID = supportingCircular.ID;
                    this.setState({ supportingDocItem: linkItem.ListData, openSupportingDoc: true })
                }).catch((error) => {
                    console.log(error);
                    this.setState({ isLoading: false })
                })

        })
    }


    private deleteCircular = (selectedItem) => {
        let providerValue = this.context;
        const { services, serverRelativeUrl } = providerValue as IBobCircularRepositoryProps;



        this.setState({ isLoading: true, showDeleteDialog: false }, async () => {
            await services.deleteListItem(serverRelativeUrl, Constants.circularList, selectedItem.ID).then((val) => {

                this.setState({ isLoading: false }, () => {
                    this.onEditDashBoardLoad()
                })

            }).catch((error) => {
                console.log(error);
                this.setState({ isLoading: false })
            })
        })
    }

    private workingOnIt = (): JSX.Element => {

        let submitDialogJSX = <>

            <Dialog modalType="alert" defaultOpen={true}>
                <DialogSurface style={{ maxWidth: 250 }}>
                    <DialogBody style={{ display: "block" }}>
                        <DialogContent>
                            {<Spinner labelPosition="below" label={"please wait..."}></Spinner>}
                        </DialogContent>
                    </DialogBody>
                </DialogSurface>
            </Dialog>

        </>;
        return submitDialogJSX;
    }

    private deleteDialog = (showDialog, selectedItem): JSX.Element => {
        let submitDialogJSX = <>
            <>
                <Dialog modalType="alert" defaultOpen={(showDialog)} >
                    <DialogSurface style={{ maxWidth: 330 }}>
                        <DialogBody style={{ display: "block" }}>
                            <DialogTitle style={{ fontFamily: "Roboto", marginBottom: 10, textAlign: "center" }}>{`${`Delete Circular` ?? ``}`}</DialogTitle>
                            <DialogContent style={{ fontFamily: "Roboto", minHeight: 45 }}>
                                {`${`Are you sure you want to delete the circular?`}`}
                            </DialogContent>
                            <DialogActions style={{ justifyContent: "center" }}>
                                <div className={`${styles1.row}`}>
                                    <div className={`${styles1.column6}`}>
                                        <Button appearance="primary"
                                            onClick={() => {
                                                this.setState({ showDeleteDialog: false }, () => {
                                                    this.deleteCircular(selectedItem)
                                                })

                                            }}>Yes</Button>
                                    </div>
                                    <div className={`${styles1.column6}`}>
                                        <Button appearance="secondary"
                                            onClick={() => {
                                                this.setState({ showDeleteDialog: false })
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

    private formatDate(dateStr: string): string {
        const date = new Date(dateStr);
        const month = (date.getMonth() + 1 < 10 ? '0' : '') + (date.getMonth() + 1);
        const day = (date.getDate() < 10 ? '0' : '') + date.getDate();
        const year = date.getFullYear().toString();
        let hour = date.getHours();
        const ampm = hour >= 12 ? 'pm' : 'am';
        hour = hour % 12;
        hour = hour ? hour : 12;
        const minute = (date.getMinutes() < 10 ? '0' : '') + date.getMinutes();

        const dateString = `${day} -${month} -${year} `;
        const dateParts: any[] = dateString.split("-");

        // create a new Date object with the year, month, and day
        const dateObject = new Date(dateParts[2], dateParts[1] - 1, dateParts[0]);

        // format the date as a string using the desired format
        const formattedDate = dateObject.toLocaleDateString("en-UK", { day: "numeric", month: "short", year: "numeric" });

        return `${formattedDate} `;
    }
}
