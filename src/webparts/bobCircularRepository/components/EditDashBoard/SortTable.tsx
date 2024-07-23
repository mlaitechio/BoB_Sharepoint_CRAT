import * as React from "react";
import {
    FolderRegular,
    EditRegular,
    OpenRegular,
    DocumentRegular,
    PeopleRegular,
    DocumentPdfRegular,
    VideoRegular,
    EyeRegular,
    DeleteRegular,
} from "@fluentui/react-icons";
import {
    PresenceBadgeStatus,
    Avatar,
    TableBody,
    TableCell,
    TableRow,
    Table,
    TableHeader,
    TableHeaderCell,
    useTableFeatures,
    TableColumnDefinition,
    TableColumnId,
    useTableSort,
    TableCellLayout,
    createTableColumn,
    Button,
} from "@fluentui/react-components";

import styles1 from '../BobCircularRepository.module.scss';
import { ICircularListItem, ISortTable } from "../../Models/IModel";
import { Constants } from "../../Constants/Constants";

type FileCell = {
    label: string;
    icon: JSX.Element;
};

type LastUpdatedCell = {
    label: string;
    timestamp: number;
};

type LastUpdateCell = {
    label: string;
    icon: JSX.Element;
};

type AuthorCell = {
    label: string;
    status: PresenceBadgeStatus;
};

type Item = {
    file: FileCell;
    author: AuthorCell;
    lastUpdated: LastUpdatedCell;
    lastUpdate: LastUpdateCell;
};

const items: Item[] = [
    {
        file: { label: "Meeting notes", icon: <DocumentRegular /> },
        author: { label: "Max Mustermann", status: "available" },
        lastUpdated: { label: "7h ago", timestamp: 3 },
        lastUpdate: {
            label: "You edited this",
            icon: <EditRegular />,
        },
    },
    {
        file: { label: "Thursday presentation", icon: <FolderRegular /> },
        author: { label: "Erika Mustermann", status: "busy" },
        lastUpdated: { label: "Yesterday at 1:45 PM", timestamp: 2 },
        lastUpdate: {
            label: "You recently opened this",
            icon: <OpenRegular />,
        },
    },
    {
        file: { label: "Training recording", icon: <VideoRegular /> },
        author: { label: "John Doe", status: "away" },
        lastUpdated: { label: "Yesterday at 1:45 PM", timestamp: 2 },
        lastUpdate: {
            label: "You recently opened this",
            icon: <OpenRegular />,
        },
    },
    {
        file: { label: "Purchase order", icon: <DocumentPdfRegular /> },
        author: { label: "Jane Doe", status: "offline" },
        lastUpdated: { label: "Tue at 9:30 AM", timestamp: 1 },
        lastUpdate: {
            label: "You shared this in a Teams chat",
            icon: <PeopleRegular />,
        },
    },
];


export const SortControlled = (props: ISortTable) => {

    const { listItems, sortColumn, accordionFields, tableColumns } = props

    const [sortState, setSortState] = React.useState<{
        sortDirection: "ascending" | "descending";
        sortColumn: TableColumnId | undefined;
    }>({
        sortDirection: "ascending" as const,
        sortColumn: sortColumn,
    });

    const columns: TableColumnDefinition<Item>[] = tableColumns.map((val) => {
        return createTableColumn<Item>({
            columnId: val.columnKey,
            compare: (a, b) => {
                let sortingItem: any = null;
                let columnKey = val.columnKey;
                switch (val.columnType) {

                    case `Text`: sortingItem = a[columnKey].localeCompare(b[columnKey]);
                        break;
                    case `Date`: sortingItem = a[columnKey] - b[columnKey];
                        break;
                    case `Number`: sortingItem = a[columnKey] - b[columnKey];
                        break;
                    default: sortingItem = null;
                        break;
                }

                return sortingItem
            },
        })
    })

    const items = listItems as any[];

    const {
        getRows,
        sort: { getSortDirection, toggleColumnSort, sort },
    } = useTableFeatures(
        {
            columns,
            items,
        },
        [
            useTableSort({
                sortState,
                onSortChange: (e, nextSortState) => setSortState(nextSortState),
            }),
        ]
    );

    const headerSortProps = (columnId: TableColumnId) => ({
        onClick: (e: React.MouseEvent) => toggleColumnSort(e, columnId),
        sortDirection: getSortDirection(columnId),
    });

    const rows = sort(getRows());

    return (
        <Table
            sortable
            aria-label="Table with controlled sort"
            style={{ minWidth: "500px" }}
        >
            <TableHeader>
                <TableRow>
                    {/* <TableHeaderCell {...headerSortProps("file")}>File</TableHeaderCell>
                    <TableHeaderCell {...headerSortProps("author")}>
                        Author
                    </TableHeaderCell>
                    <TableHeaderCell {...headerSortProps("lastUpdated")}>
                        Last updated
                    </TableHeaderCell>
                    <TableHeaderCell {...headerSortProps("lastUpdate")}>
                        Last update
                    </TableHeaderCell> */}
                    {tableColumns.map((column, index) => (
                        <TableHeaderCell
                            {...headerSortProps(`${column.columnKey}`)}
                            key={column.columnId}
                            sortable={column.columnType != ""}
                            colSpan={index == 0 ? 3 : 1}
                            style={index == 0 ? { paddingLeft: 15 } : {}}
                            className={`${styles1.fontWeightBold}`}>
                            {column.label}
                        </TableHeaderCell>
                    ))}
                </TableRow>
            </TableHeader>
            {/* <TableBody>
        {rows.map(({ item }) => (
          <TableRow key={item.file.label}>
            <TableCell>
              <TableCellLayout media={item.file.icon}>
                {item.file.label}
              </TableCellLayout>
            </TableCell>
            <TableCell>
              <TableCellLayout
                media={
                  <Avatar
                    aria-label={item.author.label}
                    name={item.author.label}
                    badge={{
                      status: item.author.status as PresenceBadgeStatus,
                    }}
                  />
                }
              >
                {item.author.label}
              </TableCellLayout>
            </TableCell>
            <TableCell>{item.lastUpdated.label}</TableCell>
            <TableCell>
              <TableCellLayout media={item.lastUpdate.icon}>
                {item.lastUpdate.label}
              </TableCellLayout>
            </TableCell>
          </TableRow>
        ))}
      </TableBody> */}

            <TableBody>
                {rows && rows.length > 0 && rows.map(({ item }) => {

                    let isFieldSelected = (accordionFields.isSummarySelected ||
                        accordionFields.isTypeSelected ||
                        accordionFields.isCategorySelected ||
                        accordionFields.isSupportingDocuments);
                    // let isCurrentItem = currentSelectedItemId == item.ID;
                    //let tableRowClass = isFieldSelected && isCurrentItem ? `${styles1.tableRow}` : ``;
                    let isEditButtonVisible = item.CircularStatus == Constants.draft ||
                        item.CircularStatus == Constants.cmmtChecker
                        || item.CircularStatus == Constants.cmmtCompliance;

                    return <>
                        <TableRow className={`${styles1.tableRow}`} >
                            <TableCell colSpan={3}>
                                <TableCellLayout className={`${styles1.verticalSpacing}`} style={{ padding: 5 }}>
                                    <div
                                        className={`${styles1.colorLabel}`}
                                        style={{
                                            color: item.Classification == "Master" ? "#f26522" : "#162B75"
                                        }}>{item.CircularNumber}</div>
                                    <div className={`${styles1.verticalSpacing}`}>
                                        <Button
                                            style={{
                                                padding: 0, fontWeight: 400,
                                                justifyContent: "flex-start",
                                                alignItems: "flex-start"
                                            }}
                                            appearance="transparent"
                                        //onClick={this.onDetailItemClick.bind(this, val, Constants.colSubject)}
                                        >
                                            <div style={{
                                                textAlign: "left",
                                                marginTop: 5,
                                                color: item.Classification == "Master" ? "#f26522" : "#162B75"
                                            }}>{item.Subject} </div>
                                            {/* <OpenRegular /> */}
                                        </Button>
                                    </div>
                                </TableCellLayout>
                            </TableCell>
                            <TableCell>
                                <TableCellLayout>
                                    {item.ID != "" ? item.ID : ``}
                                </TableCellLayout>
                            </TableCell>
                            <TableCell>
                                <TableCellLayout>
                                    {item.Created != "" ? `` : ``} //this.formatDate(item.Created)
                                </TableCellLayout>
                            </TableCell>
                            <TableCell>
                                <TableCellLayout>
                                    {item.CircularStatus ? item.CircularStatus : ``}
                                </TableCellLayout>
                            </TableCell>
                            <TableCell colSpan={1}>
                                <TableCellLayout className={`${styles1.verticalSpacing}`}>
                                    {!isEditButtonVisible &&
                                        <Button onClick={() => { }} //this.viewCircular(val)
                                            icon={<EyeRegular />}
                                            style={{ marginRight: 5 }} />
                                    }

                                    {/* {
                                    isUserMaker && isEditButtonVisible && <>
                                        <Button onClick={() => {  }} //this.editCircular(val)
                                            icon={<EditRegular />}
                                            style={{ marginRight: 5 }} />

                                        
                                        {item.CircularStatus == Constants.draft &&
                                            < Button icon={<DeleteRegular />}
                                                onClick={() => {
                                                    //this.setState({ showDeleteDialog: true, currentSelectedItem: val })
                                                }}
                                            />
                                        }
                                    </>} */}

                                </TableCellLayout>
                            </TableCell>

                        </TableRow >
                        {/* <TableRow className={`${tableRowClass}`}>

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
                        } */}

                    </>
                })}

            </TableBody>
        </Table>
    );
};