import * as React from 'react'
import { IEditDashBoardProps } from './IEditDashBoardProps'
import { IEditDashBoardState } from './IEditDashBoardState'
import { Constants } from '../../Constants/Constants'
import {
    Button, Card, CardHeader, CardPreview, Dialog, DialogActions, DialogBody, DialogContent,
    DialogSurface, DialogTitle, Divider, Label, Link, Menu, MenuButton, MenuItem, MenuList, MenuPopover, MenuTrigger, Spinner,
    Table, TableBody, TableCell, TableCellLayout, TableHeader,
    TableHeaderCell, TableRow
} from '@fluentui/react-components';

import { Label as Label1 } from "@fluentui/react";


import styles1 from '../BobCircularRepository.module.scss';
import styles from '../Search/CircularSearch.module.scss';
import { ICircularListItem } from '../../Models/IModel';
import { ArrowUpRegular, ChevronDownRegular, ChevronUpRegular, Delete12Regular, Delete16Regular, Delete20Regular, DeleteRegular, Edit12Regular, Edit16Regular, Edit20Regular, EditRegular, Eye20Regular, EyeRegular, MoreHorizontal20Regular, MoreHorizontalRegular, OpenRegular } from '@fluentui/react-icons';
import { AnimationClassNames, Icon, IconButton } from '@fluentui/react';
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
                isFaqSelected: false,
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

        let loadDashBoard = localStorage.getItem("loadDashBoard") == "true";

        if (prevProps.stateKey != this.props.stateKey && loadDashBoard) {
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

                        localStorage.setItem("loadDashBoard", "true");

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
        const { context, responsiveMode } = providerValue as IBobCircularRepositoryProps;
        const { currentPage } = this.props;

        console.log(responsiveMode)
        let isMobileMode = responsiveMode == 0;
        let isMobileDesktopMode = responsiveMode == 1;
        let isTabletMode = responsiveMode == 2;
        let isDesktopMode = responsiveMode == 3 || responsiveMode == 4 || responsiveMode == 5;
        let isLoadDashboard = localStorage.getItem("loadDashBoard") == "true" ? true : false;

        return (
            <>
                {
                    isLoading && this.workingOnIt()
                }
                {
                    loadDashBoard && (isDesktopMode) && <>
                        {this.circularResults()}
                        {this.createPagination()}
                    </>
                }

                {
                    loadDashBoard && (isMobileDesktopMode || isMobileMode || isTabletMode) && isLoadDashboard &&
                    <>
                        {this.mobileDetailListView()}
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
                    loadEditForm && !isLoadDashboard &&
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
                    loadViewForm && !isLoadDashboard &&
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

                {showDeleteDialog && this.deleteDialog(showDeleteDialog, currentSelectedItem, (isMobileDesktopMode || isMobileMode))}

            </>
        )
    }


    private mobileDetailListView = (): JSX.Element => {

        let providerValue = this.context;
        const { context, isUserMaker } = providerValue as IBobCircularRepositoryProps;
        let currentUserEmail = context.pageContext.user.email;
        const { currentPage } = this.props
        const { accordionFields, currentSelectedItem, currentSelectedItemId, filteredItems } = this.state;
        let filteredPageItems = this.paginateFn(filteredItems);


        let mobileListViewJSX = <>

            <div className={`${styles1.mobileColumn12} ${styles1.headerBackgroundColor} ${styles1.marginBottom}`}
                style={{ textAlign: "center" }} >
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

            {
                filteredPageItems.length > 0 && filteredPageItems.map((value: ICircularListItem, index) => {

                    let isCurrentItem = currentSelectedItemId == value.ID;
                    let isFieldSelected = (accordionFields.isSummarySelected || accordionFields.isFaqSelected || accordionFields.isSupportingDocuments);
                    let createdBy = value.Author
                    let requesterMail = createdBy[0].email ?? ``; //val?.Author?.split('#')[4].replace(',', '');
                    let requesterName = value?.Author[0].title ?? `` //val?.Author?.split('#')[1].replace(',', '');
                    let isEditButtonVisible = (value.CircularStatus == Constants.draft ||
                        value.CircularStatus == Constants.cmmtChecker
                        || value.CircularStatus == Constants.cmmtCompliance) && (requesterMail == currentUserEmail);
                    return <>
                        <Card className={`${styles1.marginBottom} ${styles1.mobileCard}`} size={`small`} appearance='outline'>
                            <CardHeader //description={`${value.Subject}`}

                                description={this.createHyper(value)}
                                className={`${styles1.mobileHeaderFont} ${styles1.borderBottom}`}

                                action={
                                    <Menu>
                                        <MenuTrigger disableButtonEnhancement>
                                            <MenuButton menuIcon={<MoreHorizontal20Regular />} appearance="transparent"></MenuButton>
                                        </MenuTrigger>

                                        <MenuPopover>
                                            <MenuList>
                                                {!isEditButtonVisible &&
                                                    <MenuItem className={`${styles1.fontRoboto}`} icon={<Eye20Regular />} onClick={() => { this.viewCircular(value) }}>View</MenuItem>
                                                }
                                                {isUserMaker && isEditButtonVisible && <>
                                                    <MenuItem className={`${styles1.fontRoboto}`} icon={<Edit20Regular />} onClick={() => { this.editCircular(value) }}>Edit</MenuItem>
                                                    {value.CircularStatus == Constants.draft &&
                                                        <MenuItem className={`${styles1.fontRoboto}`} icon={<Delete20Regular />} onClick={() => {
                                                            this.setState({ showDeleteDialog: true, currentSelectedItem: value })
                                                        }}>Delete</MenuItem>
                                                    }
                                                </>
                                                }
                                            </MenuList>
                                        </MenuPopover>
                                    </Menu>

                                    // <IconButton
                                    //     // iconProps={{ iconName: "More" }}
                                    //     menuProps={{
                                    //         items: [{
                                    //             key: "View",
                                    //             text: "View",
                                    //             // onClick: this.readItemsAsStream.bind(this, value),
                                    //             iconProps: { iconName: "View" }
                                    //         },
                                    //         {
                                    //             key: "Edit",
                                    //             text: "Edit",
                                    //             // onClick: this.readItemsAsStream.bind(this, value),
                                    //             iconProps: { iconName: "Edit" }
                                    //         },
                                    //         {
                                    //             key: "Delete",
                                    //             text: "Delete",
                                    //             // onClick: this.readItemsAsStream.bind(this, value),
                                    //             iconProps: { iconName: "Delete" }
                                    //         }],
                                    //     }} />
                                }
                            >

                            </CardHeader>
                            <CardPreview >

                                <div className={`${styles1.row}`}>
                                    <div className={`${styles1.mobileColumn4} ${styles1.mobileFontFamily} ${styles1.mobileFont} ${styles1.paddingMobileCard}`}>Circular Number</div>
                                    <div className={`${styles1.mobileColumn8} ${styles1.mobileFontFamily} ${styles1.mobileFont} ${styles1.paddingMobileCard}`}>{value.CircularNumber}</div>
                                </div>
                                <div className={`${styles1.row}`}>
                                    <div className={`${styles1.mobileColumn4} ${styles1.mobileFontFamily} ${styles1.mobileFont} ${styles1.paddingMobileCard}`}>Created Date</div>
                                    <div className={`${styles1.mobileColumn8} ${styles1.mobileFontFamily} ${styles1.mobileFont} ${styles1.paddingMobileCard}`}>{this.formatDate(value.Created)}</div>
                                </div>
                                <div className={`${styles1.row}`}>
                                    <div className={`${styles1.mobileColumn4} ${styles1.mobileFontFamily} ${styles1.mobileFont} ${styles1.paddingMobileCard}`}>Department</div>
                                    <div className={`${styles1.mobileColumn8} ${styles1.mobileFontFamily} ${styles1.mobileFont} ${styles1.paddingMobileCard}`}>{value.Department}</div>
                                </div>
                                <div className={`${styles1.row}`}>
                                    <div className={`${styles1.mobileColumn4} ${styles1.mobileFontFamily} ${styles1.mobileFont} ${styles1.paddingMobileCard}`}>Circular Status</div>
                                    <div className={`${styles1.mobileColumn8} ${styles1.mobileFontFamily} ${styles1.mobileFont} ${styles1.paddingMobileCard}`}>{value.CircularStatus}</div>
                                </div>

                                <div className={`${styles1.row}`}>
                                    <div className={`${styles1.mobileColumn4} ${styles1.mobileFontFamily} ${styles1.mobileFont} ${styles1.paddingMobileCard}`}>Requester</div>
                                    <div className={`${styles1.mobileColumn8} ${styles1.mobileFontFamily} ${styles1.mobileFont} ${styles1.paddingMobileCard}`}>{requesterName}</div>
                                </div>

                                <div className={`${styles1.row} ${styles1.marginTop}`}>
                                    <div className={`${styles1.mobileColumn3} ${styles1.paddingLeftZero}`}>
                                        <Button icon={accordionFields.isSummarySelected && isCurrentItem ? <ChevronUpRegular /> : <ChevronDownRegular />}
                                            iconPosition="after"
                                            className={accordionFields.isSummarySelected && isCurrentItem ? styles1.colorLabelMobile : ``}
                                            appearance={accordionFields.isSummarySelected && isCurrentItem ? "transparent" : "transparent"}
                                            onClick={this.onDetailItemClick.bind(this, value, Constants.colSummary)}>Summary</Button>
                                    </div>
                                    <div className={`${styles1.mobileColumn2} ${styles1.paddingLeftZero}`}>
                                        <Button icon={accordionFields.isFaqSelected && isCurrentItem ? <ChevronUpRegular /> : <ChevronDownRegular />}
                                            iconPosition="after"
                                            className={accordionFields.isFaqSelected && isCurrentItem ? styles1.colorLabelMobile : ``}
                                            appearance={accordionFields.isFaqSelected && isCurrentItem ? "transparent" : "transparent"}
                                            onClick={this.onDetailItemClick.bind(this, value, Constants.faqs)}>FAQs</Button>
                                    </div>
                                    <div className={`${styles1.mobileColumn7}`} >
                                        <Button
                                            icon={accordionFields.isSupportingDocuments && isCurrentItem ? <ChevronUpRegular /> : <ChevronDownRegular />}
                                            iconPosition="after"
                                            className={accordionFields.isSupportingDocuments && isCurrentItem ? styles1.colorLabelMobile : ``}
                                            appearance={accordionFields.isSupportingDocuments && isCurrentItem ? "transparent" : "transparent"}
                                            onClick={this.onDetailItemClick.bind(this, value, Constants.colSupportingDoc)}>Supporting Documents</Button>
                                    </div>

                                </div>

                                {isFieldSelected && currentSelectedItemId == value.ID &&
                                    <div className={`${styles1.row}`}>
                                        <div className={`${styles1.mobileColumn12} ${AnimationClassNames.slideDownIn20}`} style={{ paddingLeft: 12 }}>
                                            {accordionFields.isSummarySelected &&
                                                <>{`${currentSelectedItem?.Gist != "" ? currentSelectedItem?.Gist ?? `No summary available` : `No summary available`}`}</>
                                            }
                                            {accordionFields.isFaqSelected &&
                                                <>{`${currentSelectedItem?.CircularFAQ != "" ? currentSelectedItem?.CircularFAQ ?? `No Faqs available` : `No Faqs available`}`}</>
                                            }

                                            {accordionFields.isSupportingDocuments && <>
                                                {currentSelectedItem?.SupportingDocuments && currentSelectedItem?.SupportingDocuments != ""
                                                    ? this.supportingDocument(currentSelectedItem.SupportingDocuments) : `No Supporting Documents Available`}
                                            </>}

                                        </div>
                                    </div>
                                }

                                {/* <div className={`${styles1.row}`}>
                  <div className={`${styles1.column5} ${styles1.mobileFont} ${styles1.paddingMobileCard}`}>Status</div>
                  <div className={`${styles1.column7} ${styles1.mobileFont} ${styles1.paddingMobileCard}`}>{value.DocumentCategory.Title}</div>
                </div>
                <div className={`${styles1.row}`}>
                  <div className={`${styles1.column5} ${styles1.mobileFont} ${styles1.paddingMobileCard}`}>Published By</div>
                  <div className={`${styles1.column7} ${styles1.mobileFont} ${styles1.paddingMobileCard}`}>{value.PublisherEmailID}</div>
                </div> */}
                            </CardPreview>
                        </Card>
                    </>
                })
            }
        </>
        {/* {this.createPagination()} */ }

        return mobileListViewJSX;
    }

    private createHyper(item: any): JSX.Element {
        const name = item?.Subject;

        return (
            <>
                <div className={`${styles.viewList} ${styles1.mobileFont} ${styles1.mobileFontFamily}`} style={{ color: item.Classification == "Master" ? "#f26522" : "#162B75" }}>
                    {name}

                </div>
            </>
        )
    }

    private circularResults = () => {

        let providerValue = this.context;
        const { responsiveMode, isUserMaker, context } = providerValue as IBobCircularRepositoryProps;
        let currentUserEmail = context.pageContext.user.email;
        const { currentPage } = this.props
        const { listItems, accordionFields, currentSelectedItem, currentSelectedItemId, filteredItems } = this.state;
        let filteredPageItems = this.paginateFn(filteredItems);
        let colorLabelClass = responsiveMode == 4 ? styles1.colorLabelDesktop : responsiveMode == 3 ? styles1.colorLabelTablet : responsiveMode == 2 ? styles1.colorLabelTablet1 : styles1.colorLabel;
        const columns = [
            { columnKey: "CircularNo", label: "Circular No", columnType: "Text" },
            { columnKey: "Subject", label: "Subject", columnType: "Text" },
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
                                        colSpan={index == 1 ? 2 : 1}
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

                                    let isFieldSelected = (accordionFields.isSummarySelected || accordionFields.isFaqSelected || accordionFields.isCategorySelected || accordionFields.isSupportingDocuments);
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
                                            <TableCell>
                                                <div
                                                    className={`${colorLabelClass}`}
                                                    style={{
                                                        color: val.Classification == "Master" ? "#f26522" : "#162B75"
                                                    }}>{val.CircularNumber}</div>
                                            </TableCell>
                                            <TableCell colSpan={2}>
                                                <TableCellLayout>
                                                    <div style={{ paddingLeft: 5 }}>
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
                                                                // marginTop: 5,
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

                                            <TableCell colSpan={5}>
                                                <div className={`${styles1.row}`}>
                                                    <div className={`${styles1.column1}`} style={{ paddingLeft: "0px", marginRight: 25 }}>
                                                        <Button icon={accordionFields.isSummarySelected && isCurrentItem ? <ChevronUpRegular /> : <ChevronDownRegular />}
                                                            iconPosition="after"
                                                            style={{ width: `110px` }}
                                                            className={accordionFields.isSummarySelected && isCurrentItem ? colorLabelClass : ``}
                                                            appearance={accordionFields.isSummarySelected && isCurrentItem ? "transparent" : "transparent"}
                                                            onClick={this.onDetailItemClick.bind(this, val, Constants.colSummary)}>Summary</Button>
                                                    </div>
                                                    <div className={`${styles1.column1}`} style={{ marginRight: 20 }}>
                                                        <Button icon={accordionFields.isFaqSelected && isCurrentItem ? <ChevronUpRegular /> : <ChevronDownRegular />}
                                                            iconPosition="after"
                                                            style={{ width: `90px` }}
                                                            className={accordionFields.isFaqSelected && isCurrentItem ? colorLabelClass : ``}
                                                            appearance={accordionFields.isFaqSelected && isCurrentItem ? "transparent" : "transparent"}
                                                            onClick={this.onDetailItemClick.bind(this, val, Constants.faqs)}>FAQ</Button>
                                                    </div>
                                                    <div className={`${styles1.column1}`} style={{ marginRight: 32 }}>
                                                        <Button
                                                            icon={accordionFields.isCategorySelected && isCurrentItem ? <ChevronUpRegular /> : <ChevronDownRegular />}
                                                            iconPosition="after"
                                                            style={{ width: `110px` }}
                                                            className={accordionFields.isCategorySelected && isCurrentItem ? colorLabelClass : ``}
                                                            appearance={accordionFields.isCategorySelected && isCurrentItem ? "transparent" : "transparent"}
                                                            onClick={this.onDetailItemClick.bind(this, val, Constants.colCategory)}>Category</Button>
                                                    </div>
                                                    <div className={`${styles1.column4}`} >
                                                        <Button
                                                            icon={accordionFields.isSupportingDocuments && isCurrentItem ? <ChevronUpRegular /> : <ChevronDownRegular />}
                                                            iconPosition="after"
                                                            style={{ width: `210px` }}
                                                            className={accordionFields.isSupportingDocuments && isCurrentItem ? colorLabelClass : ``}
                                                            appearance={accordionFields.isSupportingDocuments && isCurrentItem ? "transparent" : "transparent"}
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
                                                                <>{`${currentSelectedItem?.Gist != "" ? currentSelectedItem?.Gist ?? `No summary available` : `No summary available`}`}</>
                                                            }
                                                            {accordionFields.isCategorySelected &&
                                                                <>{currentSelectedItem?.Category != "" ? currentSelectedItem?.Category ?? `No category available` : `No category available`}</>}
                                                            {accordionFields.isFaqSelected &&
                                                                <>{`${currentSelectedItem?.CircularFAQ != "" ? currentSelectedItem?.CircularFAQ ?? `No Faqs available` : `No Faqs available`}`}</>
                                                            }

                                                            {accordionFields.isSupportingDocuments && <>
                                                                {currentSelectedItem?.SupportingDocuments && currentSelectedItem?.SupportingDocuments != ""
                                                                    ? this.supportingDocument(currentSelectedItem.SupportingDocuments) : `No Supporting Documents Available`}
                                                            </>}
                                                            {/* {accordionFields.isSummarySelected &&
                                                                <>{`${currentSelectedItem?.Gist ?? ``}`}</>
                                                            } */}
                                                            {/* {accordionFields.isFaqSelected &&
                                                                <>{currentSelectedItem?.CircularFAQ ?? ``}</>} */}

                                                            {/* {accordionFields.isSupportingDocuments && <>
                                                                {currentSelectedItem?.SupportingDocuments && currentSelectedItem?.SupportingDocuments != ""
                                                                    ? this.supportingDocument(currentSelectedItem.SupportingDocuments) : `No Supporting Documents Available`}
                                                            </>} */}
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

        localStorage.setItem("loadDashBoard", "false");

        this.setState({
            loadDashBoard: false,
            editFormItem: selectedItem, loadEditForm: true, loadViewForm: false
        });
    }

    private viewCircular = (selectedItem) => {

        localStorage.setItem("loadDashBoard", "false");

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
                        isFaqSelected: false,
                        isCategorySelected: false,
                        isSupportingDocuments: false
                    },
                    currentSelectedItem: item,
                    currentSelectedItemId: item.ID
                }, () => {
                    // this.readItemsAsStream(item)
                })

                break;

            case Constants.faqs:
                this.setState({
                    accordionFields: {
                        isSummarySelected: false,
                        isFaqSelected: isCurrentItem ? !accordionFields.isFaqSelected : true,
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
                        isFaqSelected: false,
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
                    isFaqSelected: false,
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

    private deleteDialog = (showDialog, selectedItem, mode?: any): JSX.Element => {
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
                                    <div className={`${mode ? `${styles1.mobileColumn6} ${styles1.textAlignEnd}` : styles1.column6}`}>
                                        <Button appearance="primary"
                                            onClick={() => {
                                                this.setState({ showDeleteDialog: false }, () => {
                                                    this.deleteCircular(selectedItem)
                                                })

                                            }}>Yes</Button>
                                    </div>
                                    <div className={`${mode ? `${styles1.mobileColumn6}` : styles1.column6}`}>
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
        let isMobileDesktopMode = responsiveMode == 1;
        let isMobileMode = responsiveMode == 0;
        let lastItemCount = ((itemsPerPage * (currentPage - 1)) + itemsPerPage) > filteredItems.length ? filteredItems.length : ((itemsPerPage * (currentPage - 1)) + itemsPerPage)
        let pagination: any =
            <>
                <div className={`${styles.paginationContainer} ${styles1.row} `}>

                    <div className={`${(isMobileDesktopMode || isMobileMode) ? styles1.mobileColumn5 : styles1.column4} `} style={{ padding: (isMobileMode || isMobileDesktopMode) ? 0 : `inherit` }}>
                        {/* {<Label>{JSON.stringify(theme.palette)}</Label>} */}
                        {/* {<Label>{JSON.stringify(_theme)}</Label>} */}
                        {<Label1 styles={{
                            root: {
                                paddingTop: 20,
                                textAlign: "left",
                                fontSize: isMobileDesktopMode ? 14 : isMobileMode ? 12 : 14,
                                paddingLeft: 15,
                                fontFamily: 'Roboto'
                            }
                        }}>
                            {filteredItems.length > 0
                                &&
                                `Showing ${(itemsPerPage * (currentPage - 1) + 1)} to ${lastItemCount} of ${totalItems} `
                            }
                        </Label1>}
                    </div>
                    <div className={`${styles.searchWp__paginationContainer__pagination} ${(isMobileMode || isMobileDesktopMode) ? styles1.mobileColumn7 : styles1.column8} `} style={{ padding: isMobileMode ? 0 : `inherit` }}>
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
