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
import styles from '../Search/CircularSearch.module.scss';
import { ICircularListItem } from '../../Models/IModel';
import { ArrowUpRegular, ChevronDownRegular, ChevronUpRegular, Delete12Regular, Delete16Regular, DeleteRegular, Edit12Regular, Edit16Regular, EditRegular, EyeRegular, OpenRegular } from '@fluentui/react-icons';
import { AnimationClassNames, Icon } from '@fluentui/react';
import { IBobCircularRepositoryProps } from '../IBobCircularRepositoryProps';
import { DataContext } from '../../DataContext/DataContext';
import { error } from 'pdf-lib';
import FileViewer from '../FileViewer/FileViewer';
import { Text } from '@microsoft/sp-core-library';
import CircularForm from '../CircularForm/CircularForm';
import { SortControlled } from './SortTable';
import Pagination from 'react-js-pagination';
import { IRenderListDataAsStreamResult } from '@pnp/sp/lists';

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
            currentPage: 1,
            itemsPerPage: 11,
            items: [],
            filteredItems: [],
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

    private onEditDashBoardLoad = () => {
        let providerValue = this.context;
        const { services, serverRelativeUrl } = providerValue as IBobCircularRepositoryProps;
        const { filterString } = this.props

        this.setState({ isLoading: true }, async () => {
            await services.filterLargeListItem(serverRelativeUrl, Constants.circularList, `${filterString}`).
                then(async (itemIDColl: any[]) => {
                    let allListItems: any[] = [];


                    await services.renderListDataStream(serverRelativeUrl, Constants.circularList, itemIDColl).then((listItems: IRenderListDataAsStreamResult) => {
                        allListItems = listItems?.Row ?? []

                    }).catch((error) => {
                        console.log(error)
                    })
                    // await Promise.all(itemIDColl?.map(async (item) => {
                    //     return await services.getListDataAsStream(serverRelativeUrl, Constants.circularList, item.ID).then((listItem) => {
                    //         listItem.ListData.ID = item.ID;
                    //         return listItem?.ListData ?? []
                    //     }).catch((error) => {
                    //         console.log("Error:" + error);
                    //         return []
                    //     })
                    // }))
                    this.setState({
                        listItems: allListItems.sort((a, b) => parseInt(a.ID) > parseInt(b.ID) ? -1 : 1)
                    }, () => {
                        const { listItems } = this.state;
                        this.setState({
                            filteredItems: listItems,
                            items: listItems,
                            loadDashBoard: true,
                            loadEditForm: false,
                            loadViewForm: false,
                            isLoading: false,
                        })
                    })
                }).catch((error) => {
                    console.log(error);
                    this.setState({ isLoading: false })
                })
        })
    }



    render() {
        const { isLoading,
            openSupportingDoc,
            supportingDocItem, loadDashBoard,
            currentSelectedItem,
            loadEditForm, loadViewForm, editFormItem, showDeleteDialog } = this.state;
        let providerValue = this.context;
        const { context } = providerValue as IBobCircularRepositoryProps;
        const { currentPage } = this.props;

        return (
            <>
                {
                    isLoading && this.workingOnIt()
                }
                {
                    loadDashBoard && <>
                        {this.circularResults()}
                        {this.createPagination()}
                    </>
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
                        currentPage={currentPage}
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
                        currentPage={currentPage}
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



    private circularResults = () => {

        let providerValue = this.context;
        const { isUserChecker, isUserCompliance, isUserMaker, context } = providerValue as IBobCircularRepositoryProps;
        let currentUserEmail = context.pageContext.user.email;
        const { currentPage } = this.props
        const { listItems, accordionFields, currentSelectedItem, currentSelectedItemId, filteredItems } = this.state;
        let filteredPageItems = this.paginateFn(filteredItems);
        const columns = [
            { columnKey: "Subject", label: "Document Title", columnType: "Text" },
            { columnKey: "ID", label: "ID", columnType: "Number" },
            { columnKey: "Created", label: "Created Date", columnType: "Date" },
            { columnKey: "CircularStatus", label: "Circular Status", columnType: "Text" },
            { columnKey: "Requester", label: "Requester", columnType: "Text" },
            { columnKey: "Edit", label: "", columnType: "" },
        ];


        let circularResultJSX = <>
            <div className={`${styles1.row}`} >
                <div className={`${styles1.column12} ${styles1.headerBackgroundColor}`} style={{ textAlign: "center" }} >
                    {
                        <Label style={{
                            fontFamily: "Roboto",
                            padding: 10,
                            cursor: "pointer",
                            fontSize: "var(--fontSizeBase500)",
                            fontWeight: "var(--fontWeightSemibold)",
                            lineHeight: "var(--lineHeightBase500)",
                            color: "white",

                        }}> {`${currentPage} Circular Dashboard`}
                        </Label>}

                </div>

                {/* {columns && columns.length > 0 && listItems && listItems.length > 0 &&
                    <SortControlled
                        tableColumns={columns}
                        sortColumn={columns[0].columnKey}
                        accordionFields={accordionFields}
                        listItems={listItems} />
                } */}

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
                                        key={column.columnKey}
                                        colSpan={index == 0 ? 3 : 1}
                                        style={index == 0 ? { paddingLeft: 15 } : {}}
                                        className={`${styles1.fontWeightBold}`}>
                                        {column.label}
                                    </TableHeaderCell>
                                ))}
                            </TableRow>
                        </TableHeader>
                        <TableBody>
                            {filteredPageItems && filteredPageItems.length > 0 &&
                                filteredPageItems.map((val: ICircularListItem, index) => {

                                    let isFieldSelected = (accordionFields.isSummarySelected || accordionFields.isTypeSelected || accordionFields.isCategorySelected || accordionFields.isSupportingDocuments);
                                    let isCurrentItem = currentSelectedItemId == val.ID;
                                    let tableRowClass = isFieldSelected && isCurrentItem ? `${styles1.tableRow}` : ``;
                                    let createdBy = val.Author
                                    let requesterMail = createdBy[0].email ?? ``; //val?.Author?.split('#')[4].replace(',', '');
                                    let requesterName = val?.Author[0].title ?? `` //val?.Author?.split('#')[1].replace(',', '');
                                    let isEditButtonVisible = (val.CircularStatus == Constants.draft ||
                                        val.CircularStatus == Constants.cmmtChecker
                                        || val.CircularStatus == Constants.cmmtCompliance) && (requesterMail == currentUserEmail);



                                    return <>
                                        <TableRow className={`${styles1.tableRow}`} >
                                            <TableCell colSpan={3}>
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
                                            <TableCell as="div">
                                                <TableCellLayout>
                                                    {val.ID != "" ? val.ID : ``}
                                                </TableCellLayout>
                                            </TableCell>
                                            <TableCell as="div">
                                                <TableCellLayout>
                                                    {val.Created != "" ? this.formatDate(val.Created) : ``}
                                                </TableCellLayout>
                                            </TableCell>
                                            <TableCell as="div">
                                                <TableCellLayout>
                                                    {val.CircularStatus ? val.CircularStatus : ``}
                                                </TableCellLayout>
                                            </TableCell>
                                            <TableCell as="div">
                                                <TableCellLayout main={{
                                                    style: {
                                                        display: "block",
                                                        maxWidth: 150,
                                                        textOverflow: "ellipsis",
                                                        overflow: "hidden"
                                                    },
                                                    title: requesterMail ?? ``
                                                }}>
                                                    {requesterName ?? ``}
                                                </TableCellLayout>
                                            </TableCell>
                                            <TableCell as="div">
                                                <TableCellLayout className={`${styles1.verticalSpacing}`} as="div">
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

                                            <TableCell colSpan={4}>
                                                <div className={`${styles1.row}`}>
                                                    <div className={`${styles1.column1}`} style={{ paddingLeft: "0px", marginRight: 25 }}>
                                                        <Button icon={accordionFields.isSummarySelected && isCurrentItem ? <ChevronUpRegular /> : <ChevronDownRegular />}
                                                            iconPosition="after"
                                                            className={accordionFields.isSummarySelected && isCurrentItem ? styles1.colorLabel : ``}
                                                            appearance={accordionFields.isSummarySelected && isCurrentItem ? "outline" : "transparent"}
                                                            onClick={this.onDetailItemClick.bind(this, val, Constants.colSummary)}>Summary</Button>
                                                    </div>
                                                    <div className={`${styles1.column1}`} style={{ marginRight: 20 }}>
                                                        <Button icon={accordionFields.isTypeSelected && isCurrentItem ? <ChevronUpRegular /> : <ChevronDownRegular />}
                                                            iconPosition="after"
                                                            className={accordionFields.isTypeSelected && isCurrentItem ? styles1.colorLabel : ``}
                                                            appearance={accordionFields.isTypeSelected && isCurrentItem ? "outline" : "transparent"}
                                                            onClick={this.onDetailItemClick.bind(this, val, Constants.colType)}>Type</Button>
                                                    </div>
                                                    <div className={`${styles1.column1}`} style={{ marginRight: 32 }}>
                                                        <Button
                                                            icon={accordionFields.isCategorySelected && isCurrentItem ? <ChevronUpRegular /> : <ChevronDownRegular />}
                                                            iconPosition="after"
                                                            className={accordionFields.isCategorySelected && isCurrentItem ? styles1.colorLabel : ``}
                                                            appearance={accordionFields.isCategorySelected && isCurrentItem ? "outline" : "transparent"}
                                                            onClick={this.onDetailItemClick.bind(this, val, Constants.colCategory)}>Category</Button>
                                                    </div>
                                                    <div className={`${styles1.column4}`} >
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
                                                <TableCell colSpan={4}>
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


    private paginateFn = (filterItem: any[]) => {
        let { itemsPerPage, currentPage } = this.state;
        return (itemsPerPage > 0
            ? filterItem ? filterItem.slice((currentPage - 1) * itemsPerPage, (currentPage - 1) * itemsPerPage + itemsPerPage) : filterItem
            : filterItem
        );
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

    private createPagination = (): JSX.Element => {
        const { items, currentPage, itemsPerPage, filteredItems } = this.state;
        let providerContext = this.context;
        const { responsiveMode } = providerContext as IBobCircularRepositoryProps;
        const totalItems = filteredItems.length;
        const _themeWindow: any = window;
        const _theme = _themeWindow.__themeState__.theme;
        let isMobileMode = responsiveMode == 0 || responsiveMode == 1 || responsiveMode == 2;
        let lastItemCount = ((itemsPerPage * (currentPage - 1)) + itemsPerPage) > filteredItems.length ? filteredItems.length : ((itemsPerPage * (currentPage - 1)) + itemsPerPage)
        let pagination: any =
            <>
                <div className={`${styles.paginationContainer} ${styles1.row} `}>

                    <div className={`${styles1.column4} `} style={{ padding: isMobileMode ? 0 : `inherit` }}>
                        {/* {<Label>{JSON.stringify(theme.palette)}</Label>} */}
                        {/* {<Label>{JSON.stringify(_theme)}</Label>} */}
                        {<Label style={{
                            paddingTop: 20,
                            textAlign: "left",
                            fontSize: isMobileMode ? 11 : 14,
                            display: "block",
                            fontWeight: 700,
                            paddingLeft: 20,
                            fontFamily: 'Roboto'
                        }}>
                            {filteredItems.length > 0
                                &&
                                `Showing ${(itemsPerPage * (currentPage - 1) + 1)} to ${lastItemCount} of ${totalItems} `
                            }
                        </Label>}
                    </div>
                    <div className={`${styles.searchWp__paginationContainer__pagination} ${styles1.column8} `} style={{ padding: isMobileMode ? 0 : `inherit` }}>
                        {filteredItems.length > 0 &&
                            <Pagination
                                activePage={currentPage}
                                firstPageText={<Icon iconName="DoubleChevronLeftMed"
                                    styles={{ root: { color: _theme.themePrimary, fontWeight: 600 } }}
                                ></Icon>}
                                lastPageText={<Icon iconName="DoubleChevronRight" styles={{ root: { color: _theme.themePrimary, fontWeight: 600 } }}></Icon>}
                                prevPageText={<Icon iconName="ChevronLeft" styles={{ root: { color: _theme.themePrimary, fontWeight: 600 } }}></Icon>}
                                nextPageText={<Icon iconName="ChevronRight" styles={{ root: { color: _theme.themePrimary, fontWeight: 600 } }} ></Icon>}
                                activeLinkClass={`${styles.active} `}
                                itemsCountPerPage={itemsPerPage}
                                totalItemsCount={totalItems}
                                pageRangeDisplayed={5}
                                onChange={this.handlePageChange.bind(this)} />
                        }
                    </div>
                </div>
            </>;
        return pagination
    }

    private handlePageChange(pageNo) {
        this.setState({ currentPage: pageNo });
    }
}
