import * as React from 'react'
import styles from './CircularSearch.module.scss';
import styles1 from '../BobCircularRepository.module.scss';
import {
  AnimationClassNames,
  DefaultButton, DetailsList, DetailsListLayoutMode, DetailsRow, DialogContent, getResponsiveMode,
  IColumn, Icon, IDetailsListProps, IDetailsRowStyles,
  Image, Label, PrimaryButton, SearchBox, SelectionMode, Stack
} from '@fluentui/react';
import {
  Checkbox,
  CheckboxOnChangeData,
  Dropdown,
  Field, Label as FluentLabel,
  Input,
  OptionOnSelectData,
  Option, Button as FluentUIBtn,
  SelectionEvents,
  Switch,
  SwitchOnChangeData,
  Tag,
  Avatar,
  InputOnChangeData,
  Drawer,
  DrawerHeader,
  DrawerHeaderTitle,
  Button,
  DrawerBody,
  Accordion,
  AccordionItem,
  AccordionPanel,
  AccordionHeader,
  FluentProvider,
  Table,
  TableHeader,
  TableRow,
  TableHeaderCell,
  TableBody,
  TableCell,
  TableCellLayout,
} from "@fluentui/react-components";
import { DatePicker } from "@fluentui/react-datepicker-compat";
import { ICircularSearchProps } from './ICircularSearchProps';
import { ICircularSearchState } from './ICircularSearchState';
import { Constants } from '../../Constants/Constants';
import { Badge, Dialog, DialogBody, DialogSurface, Spinner } from '@fluentui/react-components';
import {
  TagPicker,
  TagPickerList,
  TagPickerInput,
  TagPickerControl,
  TagPickerProps,
  TagPickerOption,
  TagPickerGroup,
} from "@fluentui/react-components";
import { IBobCircularRepositoryProps } from '../IBobCircularRepositoryProps';
import Pagination from 'react-js-pagination';
import { DataContext } from '../../DataContext/DataContext';
import FileViewer from '../FileViewer/FileViewer';
import { AddCircleRegular, ArrowClockwise24Regular, ArrowClockwiseRegular, ArrowCounterclockwiseRegular, ArrowDownloadRegular, ArrowDownRegular, ArrowUpRegular, Attach12Filled, CalendarRegular, ChevronDownRegular, ChevronUpRegular, Dismiss24Regular, DismissRegular, EyeRegular, FilterRegular, OpenRegular, Search24Regular, ShareAndroidRegular, TextAlignJustifyRegular } from '@fluentui/react-icons';
import { ICircularListItem } from '../../Models/IModel';
import { PDFDocument, StandardFonts, degrees, rgb } from 'pdf-lib';
import download from 'downloadjs'

export default class CircularSearch extends React.Component<ICircularSearchProps, ICircularSearchState> {

  static contextType = DataContext;
  context!: React.ContextType<typeof DataContext>;

  constructor(props) {
    super(props)

    const columns: IColumn[] = [{
      key: 'Subject',
      name: 'Subject',
      fieldName: 'Subject',
      minWidth: 200,
      maxWidth: 450,
      isMultiline: true,
      isRowHeader: true,
      isResizable: true,
      data: 'string',
      styles: { cellName: { width: "100%" } },
      // isPadded: true,
      isSorted: false,
      isSortedDescending: true,
      sortAscendingAriaLabel: 'Sorted A to Z',
      sortDescendingAriaLabel: 'Sorted Z to A',
      onColumnClick: this.handleSorting(`${Constants.colSubject}`),
      headerClassName: styles.header,
      onRender: this.createHyper.bind(this)
    },
    {
      key: 'PublishedDate',
      name: 'Published Date',
      fieldName: 'PublishedDate',
      minWidth: 150,
      maxWidth: 200,
      // isCollapsible: true,
      isResizable: true,
      data: 'string',
      // isPadded: true,
      headerClassName: styles.header,
      styles: { cellName: { width: "100%", textAlign: "center" } },
      isSorted: false,
      isSortedDescending: true,
      sortAscendingAriaLabel: 'Sorted A to Z',
      sortDescendingAriaLabel: 'Sorted Z to A',
      onColumnClick: this.handleSorting(`${Constants.colPublishedDate}`),
      onRender: this.renderDate.bind(this)

    },
    {
      key: 'Department',
      name: 'Department',
      fieldName: 'Department',
      minWidth: 200,
      maxWidth: 400,
      isResizable: true,
      data: 'string',
      isSorted: false,
      isSortedDescending: false,
      sortAscendingAriaLabel: 'Sorted A to Z',
      sortDescendingAriaLabel: 'Sorted Z to A',
      styles: { cellName: { width: "100%" } },
      headerClassName: styles.header,
      onColumnClick: this.handleSorting(`${Constants.colCircularNumber}`)
      // isPadded: true,
      //onRender: this.renderCategory.bind(this)
    },
    {
      key: 'CircularNumber',
      name: 'Circular Number',
      fieldName: 'CircularNumber',
      minWidth: 100,
      maxWidth: 150,
      isResizable: true,
      data: 'string',
      isSorted: false,
      isSortedDescending: false,
      sortAscendingAriaLabel: 'Sorted A to Z',
      sortDescendingAriaLabel: 'Sorted Z to A',
      styles: { cellName: { width: "100%" } },
      headerClassName: styles.header,
      onColumnClick: this.handleSorting(`${Constants.colCircularNumber}`)
      // isPadded: true,
      //onRender: this.renderCategory.bind(this)
    },
    {
      key: 'Classification',
      name: 'Classification',
      fieldName: 'Classification',
      minWidth: 100,
      maxWidth: 150,
      isResizable: true,
      data: 'string',
      isSorted: false,
      isSortedDescending: false,
      sortAscendingAriaLabel: 'Sorted A to Z',
      sortDescendingAriaLabel: 'Sorted Z to A',
      styles: { cellName: { width: "100%" } },
      headerClassName: styles.header,
      onColumnClick: this.handleSorting(`${Constants.colClassification}`),
      // isPadded: true,
      //onRender: this.renderTextField.bind(this)
    }]

    this.state = {
      searchText: "",
      items: [],
      filteredItems: [],
      columns,
      currentPage: 1,
      itemsPerPage: 8,
      isLoading: false,
      departments: [],
      selectedDepartment: [],
      circularNumber: ``,
      checkCircularRefiner: `Contains`,
      circularRefinerOperator: ``,
      switchSearchText: `Normal Search`,
      isNormalSearch: false,
      sortDirection: 'asc',
      sortingFields: ``,
      publishedStartDate: null,
      publishedEndDate: null,
      previewItems: undefined,
      filePreviewItem: undefined,
      isSearchNavOpen: true,
      currentSelectedItemId: -1,
      sortingOptions: ["Date", "Subject"],
      currentSelectedFile: undefined,
      isAccordionSelected: false,
      accordionFields: {
        isSummarySelected: false,
        isTypeSelected: false,
        isCategorySelected: false
      }
    }


  }

  public componentDidMount() {
    let providerValue = this.context;
    const { services, serverRelativeUrl } = providerValue as IBobCircularRepositoryProps;

    this.setState({ isLoading: true }, async () => {
      await services.getPagedListItems(serverRelativeUrl,
        Constants.circularList, Constants.colCircularRepository, `${Constants.filterString}`,
        Constants.expandColCircularRepository, 'PublishedDate', false).then((value) => {

          const uniqueDepartment: any[] = [...new Set(value.map((item) => {
            return item.Department;
          }))];

          this.setState({
            items: value, filteredItems: value, departments: uniqueDepartment.filter((option) => {
              return option != undefined
            })
          }, () => {
            this.setState({ isLoading: false });
          })
        }).catch((error) => {
          console.log(error);
          this.setState({ isLoading: false })
        });



    })
  }

  public render() {

    let providerValue = this.context;
    const { context, services, serverRelativeUrl, circularListID } = providerValue as IBobCircularRepositoryProps;
    const { onAddNewItem } = this.props;
    const { openFileViewer,
      isLoading, isSearchNavOpen, filePreviewItem } = this.state;
    const responsiveMode = getResponsiveMode(window);
    let mobileMode = responsiveMode == 0 || responsiveMode == 1 || responsiveMode == 2;
    let searchBoxColumn = mobileMode ? `${styles1.column12}` : `${styles1.column10}`;
    let searchClearColumn = mobileMode ? `${styles1.column12} ${styles1.textAlignEnd}` : `${styles1.column2}`;
    let searchResultsColumn = isSearchNavOpen ? `${styles1.column9}` : `${styles1.column11}`


    let detailListClass = mobileMode ? `${styles1.column12}` : `${styles1.column12}`;

    return (<>

      {
        isLoading && this.workingOnIt()
      }

      <div className={`${styles1.row}`}>
        {

          !isSearchNavOpen &&
          <Drawer type={"inline"} open={!isSearchNavOpen} separator className={`${styles1.column1} ${styles1.marginTop}`}>
            <Button
              style={{ maxWidth: "100%", minWidth: "100%" }}
              className={`${styles1.fontRoboto}`}
              icon={<TextAlignJustifyRegular />} appearance="transparent" iconPosition={"after"}
              onClick={() => { this.setState({ isSearchNavOpen: true }) }}></Button>
          </Drawer>
        }
        {isSearchNavOpen &&
          <Drawer type="inline" style={{ maxHeight: "200vh" }} separator open={isSearchNavOpen} className={`${styles1.column3}`}>
            <DrawerHeader>
              <DrawerHeaderTitle
                heading={{ className: `${styles1.fontRoboto}` }}
                className={`${styles1.formLabel}`}
                action={
                  <>
                    {/* <Button
                    appearance="subtle"
                    aria-label="Close"
                    icon={<Dismiss24Regular />}
                    onClick={() => { this.setState({ isSearchNavOpen: false }) }}
                  /> */}
                  </>
                }>
                <FilterRegular />Refine Search
              </DrawerHeaderTitle>
            </DrawerHeader>
            <DrawerBody>
              {this.searchFilters()}
            </DrawerBody>
          </Drawer>
        }
        <div className={`${searchResultsColumn}`}>
          {this.searchFilterResults()}

        </div>
        {
          filePreviewItem && openFileViewer && <FileViewer
            listItem={filePreviewItem}
            context={context}
            documentLoaded={() => { this.setState({ isLoading: false }) }}
            onClose={this.onPanelClose} />
        }
      </div>


    </>
    )
  }

  private searchFilters = (): JSX.Element => {
    const { circularNumber, publishedEndDate, publishedStartDate } = this.state
    let searchFiltersJSX = <>
      <div className={`${styles1.row}`}>
        <div className={`${styles1.column12}`}>
          <div className={`${styles1.row}`}>

            <div className={`${styles1.column12} ${styles1.marginTop} `}>
              {this.pickerControl()}
            </div>
            <div className={`${styles1.column12} ${styles1.marginTop} `}>
              <Field label={<FluentLabel weight="semibold" style={{ fontFamily: "Roboto" }}>{`Circular Number`}</FluentLabel>} ></Field>
            </div>
            <div className={`${styles1.column4}`}>
              {this.checkBoxControl(`Contains`)}
            </div>
            <div className={`${styles1.column8}`} style={{ padding: 0 }}>

              <Input placeholder="Input at least 2 characters"
                input={{ className: `${styles.input}` }}
                className={`${styles.input}`}
                value={circularNumber}
                onChange={this.onInputChange} />
            </div>
            <div className={`${styles1.column12}`}>
              {this.checkBoxControl(`Starts With`)}
            </div>

            <div className={`${styles1.column12}`}>
              {this.checkBoxControl(`Ends With`)}
            </div>
          </div>

          <div className={`${styles1.row} ${styles1.marginTop}`}>
            <div className={`${styles1.column12}`}>
              <Field label={<FluentLabel weight="semibold" style={{ fontFamily: "Roboto" }}>{`Published From Date`}</FluentLabel>} >
                {/* <Input input={{ readOnly: true, type: "date" }} root={{ style: { fontFamily: "Roboto" } }}></Input> */}

                <DatePicker mountNode={{}}
                  formatDate={this.onFormatDate}
                  value={publishedStartDate}
                  contentAfter={
                    <>
                      <FluentUIBtn icon={<ArrowCounterclockwiseRegular />}
                        appearance="transparent"
                        title="Reset"
                        onClick={this.onResetClick.bind(this, `FromDate`)}>
                      </FluentUIBtn>
                      <FluentUIBtn icon={<CalendarRegular />} appearance="transparent"></FluentUIBtn>
                    </>}
                  onSelectDate={this.onSelectDate.bind(this, `FromDate`)}
                  input={{ style: { fontFamily: "Roboto" } }} />


              </Field>
              <Field label={<FluentLabel weight="semibold" style={{ fontFamily: "Roboto" }}>{`Published To Date`}</FluentLabel>}>
                <DatePicker mountNode={{}}
                  formatDate={this.onFormatDate}
                  value={publishedEndDate}
                  contentAfter={
                    <>
                      <FluentUIBtn
                        icon={<ArrowCounterclockwiseRegular />}
                        appearance="transparent" title="Reset"
                        onClick={this.onResetClick.bind(this, `ToDate`)}>
                      </FluentUIBtn>
                      <FluentUIBtn icon={<CalendarRegular />} appearance="transparent"></FluentUIBtn>

                    </>}
                  onSelectDate={this.onSelectDate.bind(this, `ToDate`)}
                  input={{ style: { fontFamily: "Roboto" } }} />
              </Field>
            </div>
          </div>
          <div className={`${styles1.row} ${styles1.marginTop}`}>
            <div className={`${styles1.column12} ${styles1.marginTop} `}>
              <Field label={<FluentLabel weight="semibold" style={{ fontFamily: "Roboto" }}>{`Classification`}</FluentLabel>} ></Field>
            </div>
            <div className={`${styles1.column12}`}>
              {this.checkBoxControl(`Master`)}
            </div>
            <div className={`${styles1.column12}`}>
              {this.checkBoxControl(`Circular`)}
            </div>
          </div>

          <div className={`${styles1.row} ${styles1.marginTop}`}>
            <div className={`${styles1.column12} ${styles1.marginTop} `}>
              <Field label={<FluentLabel weight="semibold" style={{ fontFamily: "Roboto" }}>{`Issued For`}</FluentLabel>} ></Field>
            </div>
            <div className={`${styles1.column12}`}>
              {this.checkBoxControl(`India`)}
            </div>
            <div className={`${styles1.column12}`}>
              {this.checkBoxControl(`Global`)}
            </div>
          </div>
          <div className={`${styles1.row} ${styles1.marginTop}`}>
            <div className={`${styles1.column12} ${styles1.marginTop} `}>
              <Field label={<FluentLabel weight="semibold" style={{ fontFamily: "Roboto" }}>{`Regulatory`}</FluentLabel>} ></Field>
            </div>
            <div className={`${styles1.column12}`}>
              {this.checkBoxControl(`Yes`)}
            </div>
            <div className={`${styles1.column12}`}>
              {this.checkBoxControl(`No`)}
            </div>
          </div>
          <div className={`${styles1.row} ${styles1.marginTop}`}>
            <div className={`${styles1.column12} ${styles1.marginTop} `}>
              <Field label={<FluentLabel weight="semibold" style={{ fontFamily: "Roboto" }}>{`Category`}</FluentLabel>} ></Field>
            </div>
            <div className={`${styles1.column12}`}>
              {this.checkBoxControl(`Intimation`)}
            </div>
            <div className={`${styles1.column12}`}>
              {this.checkBoxControl(`Information`)}
            </div>
            <div className={`${styles1.column12}`}>
              {this.checkBoxControl(`Action`)}
            </div>
          </div>
          <div className={`${styles1.row}`}>
            <div className={`${styles1.column12} ${styles1.marginTop} `}>
              {this.searchClearButtons()}
            </div>
          </div>

        </div>
      </div>
    </>;

    return searchFiltersJSX;

  }

  private searchFilterResults = (): JSX.Element => {
    const { filteredItems, isLoading, currentSelectedItemId,
      previewItems, sortingOptions, selectedSortFields, sortDirection, isAccordionSelected } = this.state
    let filteredPageItems = this.paginateFn(filteredItems);


    let searchFilterResultsJSX = <>
      <div className={`${styles1.row} ${styles1.marginTop}`}>
        <div className={`${styles1.column10}`}>
          {this.searchBox()}
        </div>

        <Dropdown className={`${styles1.column1}`}
          style={{ maxWidth: 95, minWidth: 95 }}
          mountNode={{}} placeholder={`Sorting`} value={selectedSortFields ?? ``}
          selectedOptions={[selectedSortFields ?? ""]}
          onOptionSelect={this.onDropDownChange.bind(this, `${Constants.sorting}`)}>
          {sortingOptions && sortingOptions.length > 0 && sortingOptions.map((val) => {
            return <><Option key={`${val}`} className={`${styles1.formLabel}`}>{val}</Option></>
          })}
        </Dropdown>

        <div className={`${styles1.column1}`}>
          <Button icon={sortDirection == "asc" ? <ArrowUpRegular /> : <ArrowDownRegular />} appearance="transparent"
            onClick={() => { this.onSorting() }} />
        </div>

      </div>
      <div className={`${styles1.row} ${styles1.marginTop}`}>

        {this.createSearchResultsTable()}
      </div>


      <div className={`${styles1.row} `}>
        {!isLoading && filteredItems.length == 0 && this.noItemFound()}
      </div>

      {this.createPagination()}

    </>;

    return searchFilterResultsJSX;
  }

  private createSearchResultsTable = (): JSX.Element => {
    const { filteredItems, previewItems, currentSelectedItemId, accordionFields } = this.state
    let filteredPageItems = this.paginateFn(filteredItems);
    const columns = [
      { columnKey: "Title", label: "Document Title" },
      { columnKey: "Date", label: "Date" },
      { columnKey: "Classification", label: "Classification" },
      { columnKey: "Department", label: "Department" },
      { columnKey: "IssuedFor", label: "Issued For" }
    ];

    let tableJSX = <>
      <Table arial-label="Default table">
        <TableHeader>
          <TableRow >
            {columns.map((column, index) => (
              <TableHeaderCell key={column.columnKey} colSpan={index == 0 ? 5 : 1} className={`${styles1.fontWeightBold}`}>
                {column.label}
              </TableHeaderCell>
            ))}
          </TableRow>
        </TableHeader>
        <TableBody>
          {filteredPageItems && filteredPageItems.length > 0 && filteredPageItems.map((val: ICircularListItem, index) => {

            let isFieldSelected = (accordionFields.isSummarySelected || accordionFields.isTypeSelected || accordionFields.isCategorySelected);
            let isCurrentItem = currentSelectedItemId == val.ID;
            let tableRowClass = isFieldSelected && isCurrentItem ? `${styles1.tableRow}` : ``;

            return <>
              <TableRow className={`${styles1.tableRow}`}>
                <TableCell colSpan={5} >
                  <TableCellLayout className={`${styles1.verticalSpacing}`}>
                    <div className={`${styles1.colorLabel}`} style={{ padding: 0 }}>{val.CircularNumber}</div>
                    <div className={`${styles1.fontWeightBold} ${styles1.verticalSpacing}`}>
                      <Button
                        style={{ padding: 0 }}
                        appearance="transparent"
                        onClick={this.onDetailItemClick.bind(this, val, Constants.colSubject)}>
                        <div style={{ textAlign: "left", marginTop: 5 }}>{val.Subject} <OpenRegular /></div>
                      </Button>
                    </div>
                  </TableCellLayout>
                </TableCell>
                <TableCell>
                  <TableCellLayout>
                    {this.formatDate(val.PublishedDate)}
                  </TableCellLayout>
                </TableCell>
                <TableCell>
                  <TableCellLayout content={{ style: { width: "100%" } }}
                    className={val.Classification == "Master" ? `${styles1.master}` : `${styles1.circular}`}>
                    {val.Classification}
                  </TableCellLayout>
                </TableCell>
                <TableCell>
                  <TableCellLayout className={`${styles1.verticalSpacing}`}>
                    {val.Department}
                  </TableCellLayout>
                </TableCell>
                <TableCell>
                  <TableCellLayout>
                    {val.IssuedFor}
                  </TableCellLayout>
                </TableCell>
              </TableRow>
              <TableRow className={`${tableRowClass}`}>
                <TableCell colSpan={5}>
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
                      <Button icon={accordionFields.isCategorySelected && isCurrentItem ? <ChevronUpRegular /> : <ChevronDownRegular />}
                        iconPosition="after"
                        className={accordionFields.isCategorySelected && isCurrentItem ? styles1.colorLabel : ``}
                        appearance={accordionFields.isCategorySelected && isCurrentItem ? "outline" : "transparent"}
                        onClick={this.onDetailItemClick.bind(this, val, Constants.colCategory)}>Category</Button>
                    </div>
                    <div className={`${styles1.column4}`}>
                      <Button
                        icon={<OpenRegular />}
                        iconPosition="after"
                        appearance="transparent"
                        onClick={this.onDetailItemClick.bind(this, val, Constants.colSupportingDoc)}>Supporting Documents</Button>
                    </div>

                  </div>
                </TableCell>
              </TableRow>
              {isFieldSelected && currentSelectedItemId == val.ID &&
                <TableRow >
                  <TableCell colSpan={7}>
                    <div className={`${styles1.row}`}>
                      <div className={`${styles1.column12}`}>
                        {accordionFields.isSummarySelected &&
                          <>{`${Constants.loreumIPSUM}`}</>
                        }
                        {accordionFields.isTypeSelected &&
                          <>{previewItems?.CircularType ?? ``}</>}
                        {accordionFields.isCategorySelected &&
                          <>{previewItems?.Category ?? ``}</>}
                      </div>
                    </div>
                  </TableCell>
                </TableRow>
              }

            </>
          })}
        </TableBody>
      </Table>
    </>;

    return tableJSX;
  }

  private onDetailItemClick = (item: ICircularListItem, fieldName: string) => {
    const { currentSelectedItemId } = this.state;
    const { accordionFields } = this.state;
    let isCurrentItem = currentSelectedItemId == item.ID;

    switch (fieldName) {
      case Constants.colSummary:

        this.setState({
          accordionFields: {
            "isSummarySelected": isCurrentItem ? !accordionFields.isSummarySelected : true,
            "isTypeSelected": false,
            isCategorySelected: false
          },
          currentSelectedItem: item,
          currentSelectedItemId: item.ID
        }, () => {
          this.readItemsAsStream(item)
        })

        break;

      case Constants.colType:
        this.setState({
          accordionFields: {
            "isSummarySelected": false,
            "isTypeSelected": isCurrentItem ? !accordionFields.isTypeSelected : true,
            isCategorySelected: false
          },
          currentSelectedItem: item,
          currentSelectedItemId: item.ID
        }, () => {
          this.readItemsAsStream(item)
        })
        break;
      case Constants.colCategory:
        this.setState({
          accordionFields: {
            "isSummarySelected": false,
            "isTypeSelected": false,
            isCategorySelected: isCurrentItem ? !accordionFields.isCategorySelected : true
          },
          currentSelectedItem: item,
          currentSelectedItemId: item.ID
        }, () => {
          this.readItemsAsStream(item)
        })
        break;

      case Constants.colSubject: this.setState({
        currentSelectedItem: item,
        currentSelectedItemId: item.ID
      }, () => {
        this.readItemsAsStream(item, true)
      })
        break;
    }
  }


  private downloadCircularContent = (listItem: ICircularListItem) => {
    let providerValue = this.context;
    const { context, services, serverRelativeUrl, userDisplayName } = providerValue as IBobCircularRepositoryProps;
    let listItemID = parseInt(listItem.ID);

    this.setState({ isLoading: true }, async () => {
      await services.getAllListItemAttachments(serverRelativeUrl, Constants.circularList, listItemID).then((fileMetadata) => {
        let fileArray: any[] = [];

        if (fileMetadata.size > 0) {
          fileMetadata.forEach(async (value, key) => {
            fileArray.push({
              "name": key,
              "content": value
            });


          });

          this.setState({ currentSelectedFile: fileArray }, () => {
            const { currentSelectedFile } = this.state
            const currentLoggedInUser = context.pageContext.user.displayName;
            this.downloadWaterMarkPDF(currentSelectedFile[0], currentLoggedInUser).then((val) => {
              this.setState({ isLoading: false });
            })
          })
        }
        else {
          this.setState({ isLoading: false }, () => {
            alert("No Circular Content Found")
          });
        }

      }).catch((error) => {
        console.log(error)
      });
    })

  }

  private handleToggle = (item, event, data) => {
    this.readItemsAsStream(item);

  }

  private checkBoxControl = (labelName): JSX.Element => {
    const { checkCircularRefiner } = this.state
    let checkBoxJSX = <>
      <Checkbox
        checked={checkCircularRefiner == labelName}
        label={<FluentLabel weight="semibold" style={{ fontFamily: "Roboto" }}>{labelName}</FluentLabel>}
        shape="circular" size="medium" onChange={this.onCheckBoxChange.bind(this, labelName)} />
    </>

    return checkBoxJSX;
  }


  private onResetClick = (labelName: string) => {
    switch (labelName) {
      case `FromDate`: this.setState({ publishedStartDate: null });
        break;
      case `ToDate`: this.setState({ publishedEndDate: null });
        break;

    }
  }



  private onCheckBoxChange = (labelName: string, ev: React.ChangeEvent<HTMLInputElement>, data: CheckboxOnChangeData) => {
    switch (labelName) {
      case `Contains`: this.setState({ checkCircularRefiner: labelName, circularRefinerOperator: `` });
        break;
      case `Starts With`: this.setState({ checkCircularRefiner: labelName, circularRefinerOperator: `starts-with` });
        break;
      case `Ends With`: this.setState({ checkCircularRefiner: labelName, circularRefinerOperator: `ends-with` });
        break;
    }
  }

  private onSwitchChange = (ev: React.ChangeEvent<HTMLInputElement>, data: SwitchOnChangeData) => {
    if (data.checked) {
      this.setState({ isNormalSearch: false, switchSearchText: `Advanced Search` })
    }
    else {
      this.setState({ isNormalSearch: true, switchSearchText: `Normal Search` })
    }
  }

  private pickerControl = (): JSX.Element => {
    const { departments, selectedDepartment } = this.state
    var pickerJSX = <>
      <Field label={<FluentLabel weight="semibold"
        style={{ fontFamily: "Roboto" }}>{`Select Department`}</FluentLabel>} >
        <TagPicker

          onOptionSelect={this.onOptionSelect}
          selectedOptions={selectedDepartment}
        >
          <TagPickerControl >
            <TagPickerGroup>
              {selectedDepartment.map((option) => (
                <Tag
                  primaryText={{ style: { fontFamily: "Roboto", fontWeight: 600 } }}
                  key={option}
                  shape="rounded"
                  size="small"

                  // media={<Avatar aria-hidden name={option} color="colorful" />}
                  value={option}
                >
                  {option}
                </Tag>
              ))}
            </TagPickerGroup>
            <TagPickerInput aria-label="Select Employees" />
          </TagPickerControl>
          <TagPickerList >
            {departments.length > 0
              ? departments.filter((option) => {
                if (selectedDepartment[0] !== option) {
                  return option;
                }

              }).map((option) => (
                <TagPickerOption
                  style={{ fontFamily: "Roboto", fontWeight: 600 }}
                  value={option}
                  key={option}>
                  {option}
                </TagPickerOption>
              ))
              : "No options available"}
          </TagPickerList>
        </TagPicker>
      </Field>
    </>

    return pickerJSX;
  }

  private onOptionSelect: TagPickerProps["onOptionSelect"] = (e, data) => {

    if (data.selectedOptions.length > 0) {
      this.setState({ selectedDepartment: [data.value] })
    }
    else {
      this.setState({ selectedDepartment: [] })
    }


  }

  private onInputChange = (ev: React.ChangeEvent<HTMLInputElement>, data: InputOnChangeData) => {
    this.setState({ circularNumber: data.value });
  }

  private onSelectDate = (labelName: string, date: Date | null) => {
    switch (labelName) {
      case `FromDate`: this.setState({ publishedStartDate: date });
        break;
      case `ToDate`: this.setState({ publishedEndDate: date });
        break;
    }

  }

  private onDropDownChange = (textLabel: string, event: SelectionEvents, data: OptionOnSelectData) => {

    switch (textLabel) {
      case Constants.sorting: this.setState({
        selectedSortFields: data.optionValue,
        sortingFields: data.optionValue == "Date" ? Constants.colPublishedDate : Constants.colSubject
      }, () => {
        const { filteredItems, sortingFields, sortDirection } = this.state;
        this.setState({ filteredItems: this.sortListItems(filteredItems, sortingFields, sortDirection) })
      });
        break;
    }
  }

  private onSorting = () => {
    const { sortDirection } = this.state;
    this.setState({ sortDirection: sortDirection == "asc" ? "desc" : "asc" }, () => {
      const { filteredItems, sortingFields, sortDirection } = this.state;
      this.setState({ filteredItems: this.sortListItems(filteredItems, sortingFields, sortDirection) })
    })
  }



  private onFormatDate = (date?: Date): string => {
    return !date
      ? ""
      : (date.getDate() < 9 ? (`0` + date.getDate()) : date.getDate()) +
      "/" +
      ((date.getMonth() + 1 < 9 ? (`0${date.getMonth() + 1}`) : date.getMonth() + 1)) +
      "/" +
      (date.getFullYear());
  };

  private updateItem = (itemID: any) => {
    this.props.onUpdateItem(itemID)
  }

  private onPanelClose = () => {
    this.setState({ filePreviewItem: null, openFileViewer: false })
  }


  private searchClearButtons = (): JSX.Element => {
    let searchClearJSX = <>
      <FluentUIBtn appearance="primary" style={{ marginRight: 5 }} icon={<FilterRegular />} onClick={() => { this.searchResults() }}>
        Search
      </FluentUIBtn>
      <FluentUIBtn appearance="secondary" icon={<DismissRegular />} onClick={() => { this.clearFilters() }}>
        Clear
      </FluentUIBtn>
    </>;

    return searchClearJSX;

  }

  private searchBox = (): JSX.Element => {
    const { searchText } = this.state;
    let searchBoxJSX =
      <Stack tokens={{ childrenGap: 20 }}>
        <SearchBox
          placeholder={`${Constants.searchText} `}
          onChange={this.onSearchBoxChange}
          onSearch={this.handleSearch}
          //onClear={(ev?: any) => { this.onClear() }}
          onClear={(ev?: any) => { this.onSearchClear() }}
          defaultValue={searchText}
          value={searchText}
          styles={{
            root: {
              border: "1px solid #bac6f7",
              fontFamily: "Roboto",
              borderRadius: 5,
              ":hover": {
                borderColor: "1px solid #bac6f7"
              },
              ".is-active": {
                border: "1px solid #bac6f7"
              }
            }
          }}
        />
      </Stack>;

    return searchBoxJSX;
  }

  private onSearchBoxChange = (event?: React.ChangeEvent<HTMLInputElement>, newValue?: string) => {
    this.setState({ searchText: newValue })
  }

  private handleSearch = (newValue?: string) => {

    this.searchResults(newValue)

  }

  private onSearchClear = () => {
    this.setState({ searchText: `` })
  }

  private searchResults = (newValue?: string) => {
    let providerValue = this.context;
    const { context, services, serverRelativeUrl, circularListID } = providerValue as IBobCircularRepositoryProps;

    let siteID = context.pageContext.site.id;
    let webID = context.pageContext.web.id;
    let siteURL = context.pageContext.site.absoluteUrl;

    this.setState({ isLoading: true }, async () => {

      let listItemData: any[] = [];
      let searchProperties = Constants.selectedSearchProperties;
      const { searchText, sortingFields, sortDirection } = this.state
      let queryTemplate = `{searchTerms} (siteId:{${siteID}} OR siteId:${siteID}) (webId:{${webID}} OR webId:${webID}) (NormListID:${circularListID}) `;
      queryTemplate += `(path:"${siteURL}/Lists/${Constants.circularList}" OR ParentLink:"${siteURL}/Lists/${Constants.circularList}*") ContentTypeId:0x0* `;

      let refinableFilterQuery = this.refinableQuery();
      let advancedSearchTextAndFilterQuery = this.searchQueryAndFilterQuery();
      let sortListProperty = [{
        Property: Constants.managePropPublishedDate,
        Direction: 1 //0 for asc & 1 for descending
      }]

      /**
      |--------------------------------------------------
      | This is to search text inside List Metadata & attachments with below refinments
      | It will search using Search Text in QueryText Search Properties & then use below
        refinableFilterQuery -> Department,CircularNumber,PublishedDate filters 
      |--------------------------------------------------
      */
      await services.
        getSearchResults(searchText.trim() == '' ? `` : searchText, searchProperties, queryTemplate, refinableFilterQuery, sortListProperty).
        then(async (searchResults: any[]) => {
          searchResults.map((val) => {

            listItemData.push({
              ID: parseInt(val.ListItemID),
              Id: parseInt(val.ListItemID),
              Created: val?.Created,
              CircularNumber: val.RefinableString00,
              Subject: val.RefinableString01,
              MigratedDepartment: val.RefinableString02,
              Department: val.RefinableString03,
              Category: val.RefinableString04,
              IsMigrated: val.RefinableString05,
              Classification: val.RefinableString06,
              PublishedDate: val.RefinableDate00,
              IssuedFor: val.RefinableString08
            })

          })
        }).catch((error) => {
          console.log(error);
          this.setState({ isLoading: false })
        });

      /**
      |--------------------------------------------------
      | Query Text= * , 
      | Will do the search text from search box inside Subject using Refinment Filters & Other Refinable Filters
      | This for combination of Search Text + Refinment Filters . 
      | It searches Subject using search text & then adds Refinment Filters.
      | It will do * Search in Query Text Property & add RefinmentFilters= (Subject:or(SearchText) + Refinment Filters)
      | 
      |--------------------------------------------------
      */
      if (advancedSearchTextAndFilterQuery != "") {

        await services.getSearchResults('', searchProperties, queryTemplate, advancedSearchTextAndFilterQuery, sortListProperty).
          then(async (searchResults: any[]) => {

            searchResults.map((val) => {

              listItemData.push({
                ID: parseInt(val.ListItemID),
                Id: parseInt(val.ListItemID),
                Created: val?.Created,
                CircularNumber: val.RefinableString00,
                Subject: val.RefinableString01,
                MigratedDepartment: val.RefinableString02,
                Department: val.RefinableString03,
                Category: val.RefinableString04,
                IsMigrated: val.RefinableString05,
                Classification: val.RefinableString06,
                PublishedDate: val.RefinableDate00,
                IssuedFor: val.RefinableString08
              })

            })

          }).catch((error) => {
            console.log(error);
            this.setState({ isLoading: false })
          });
      }


      let uniqueResults = advancedSearchTextAndFilterQuery != "" ? [...new Map(listItemData.map(item =>
        [item["Id"], item])).values()] : listItemData;

      let searchFilterItems = uniqueResults.filter((val) => { return val !== undefined });

      this.setState({
        filteredItems: this.sortListItems(searchFilterItems, sortingFields, sortDirection),
        searchItems: searchFilterItems, currentPage: 1, isLoading: false
      })

    });

  }

  private normalSearchQuery = (searchText): string => {

    let subject = Constants.managePropSubject;
    let queryTextFilters = [];
    let normalSearchString = "";

    if (searchText && searchText.trim() != "") {
      queryTextFilters = searchText.trim().split(' ');
      if (queryTextFilters.length > 1) {

        let queryText = [];
        queryTextFilters.map((word) => {
          queryText.push(`"${word}*"`)
        })
        normalSearchString = `${subject}:or(` + queryText.join(',') + `)`;

        // refinmentString += `, ${department}: or(`
        // queryText = "";

        // queryTextFilters.map((word) => {
        //   queryText += `"${word}*", `
        // })

        //  refinmentString += queryText.substring(0, queryText.length - 1) + `)`;

        // refinmentString += `, ${ circularNumber }: or(`;
        // queryText = "";
        // queryTextFilters.map((word) => {
        //   queryText += `"${word}*", `
        // })
        // refinmentString += queryText.substring(0, queryText.length - 1) + `)`;

      }

      else {
        normalSearchString += `${subject}:"${queryTextFilters[0]}*"`; //,${documentNo}:"${queryTextFilters[0]}*",${keywords}:"${queryTextFilters[0]}*"
      }
    }

    return normalSearchString;

  }

  private searchQueryAndFilterQuery = (): string => {

    const { selectedDepartment, searchText, isNormalSearch, circularNumber, circularRefinerOperator,
      publishedStartDate, publishedEndDate } = this.state;

    /**
    |--------------------------------------------------
    |  | RefinableString00 -> CircularNumber
        RefinableString01 -> Subject
        RefinableString02 -> Migrated Department
        RefinableString03 -> Department
        RefinableString04 -> Category
        RefinableString05 -> IsMigrated 
        RefinableString06 -> Classification
        RefinableDate00 -> PublishedDate 
        RefinableString07 -> CircularStatus
  
    |--------------------------------------------------
    */
    let departmentVal = selectedDepartment[0] ?? ``;//RefinableString03
    let circularVal = circularNumber != "" ? circularNumber : ``;
    let publishedStartVal = publishedStartDate?.toISOString() ?? ``;//RefinableDate00
    let publishedEndVal = publishedEndDate?.toISOString() ?? ``;


    let advanceFilterString = "";

    let filterProperties = Constants.filterSearchProperties;

    /**
    |--------------------------------------------------
    | Just to check if Search box has some text
    |--------------------------------------------------
    */
    let searchTextRefinment = this.normalSearchQuery(searchText);

    /**
  |--------------------------------------------------
  | Default filter will be Circular Status equal to published
  |--------------------------------------------------
  */
    let filterArray = [];

    // Default Search will always be Circular Status as Published
    filterArray.push(`${filterProperties[5]}:equals("${Constants.published}")`);

    if (!isNormalSearch) {
      `${departmentVal != "" ? filterArray.push(`${filterProperties[3]}:equals("${departmentVal}")`) : ``} `;
      `${circularVal != "" ? filterArray.push(`${filterProperties[0]}:${circularRefinerOperator}("${circularVal}*")`) : ``} `;
      if (publishedStartVal != "" && publishedEndVal != "") {
        filterArray.push(`${filterProperties[4]}: range(${publishedStartVal.split('T')[0]}T23:59:59Z, ${publishedEndVal.split('T')[0]}T23:59:59Z)`)
      }
    }

    if (filterArray.length > 1) {
      if (searchTextRefinment != "") {
        //  advanceFilterString += `and(${filterArray.join(',')})`;// ${searchTextRefinment}//,or(${searchTextRefinment}))
        advanceFilterString += `and(${filterArray.join(',')},${searchTextRefinment})`;
      }
      else {
        advanceFilterString += `and(${filterArray.join(',')})`
      }
    }
    else if (filterArray.length == 1) {
      if (searchTextRefinment != "") {
        advanceFilterString += `and(${filterArray.join(',')},${searchTextRefinment})`;
        //advanceFilterString += `and(${filterArray.join(',')})`;
        //advanceFilterString += filterArray[0];
      }
      else {
        advanceFilterString += filterArray[0];
      }

    }
    else {
      if (searchTextRefinment != "") {
        advanceFilterString += `${searchTextRefinment}`;
      }
      else {
        advanceFilterString += ``;
      }

    }

    console.log(advanceFilterString)

    return advanceFilterString

  }


  /**
  |--------------------------------------------------
  | This function is for Department, Circular Number & Published Date Filters
  |--------------------------------------------------
  */

  private refinableQuery = () => {
    const { selectedDepartment, searchText, isNormalSearch, circularNumber, circularRefinerOperator,
      publishedStartDate, publishedEndDate } = this.state;

    /**
    |--------------------------------------------------
    |  | RefinableString00 -> CircularNumber
        RefinableString01 -> Subject
        RefinableString02 -> Migrated Department
        RefinableString03 -> Department
        RefinableString04 -> Category
        RefinableString05 -> IsMigrated 
        RefinableString06 -> Classification
        RefinableDate00 -> PublishedDate 
        RefinableString07 -> CircularStatus
    |--------------------------------------------------
    */
    let departmentVal = selectedDepartment[0] ?? ``;//RefinableString03
    let circularVal = circularNumber != "" ? circularNumber : ``;
    let publishedStartVal = publishedStartDate?.toISOString() ?? ``;//RefinableDate00
    let publishedEndVal = publishedEndDate?.toISOString() ?? ``;

    let filterArray = [];


    let advanceFilterString = "";

    let filterProperties = Constants.filterSearchProperties;

    let searchTextRefinment = this.normalSearchQuery(searchText);

    // Default Search will always be Circular Status as Published
    //filterArray.push(`${filterProperties[5]}:equals("${Constants.published}")`);


    if (!isNormalSearch) {
      `${departmentVal != "" ? filterArray.push(`${filterProperties[3]}:equals("${departmentVal}")`) : ``} `;
      `${circularVal != "" ? filterArray.push(`${filterProperties[0]}:${circularRefinerOperator}("${circularVal}*")`) : ``} `;
      if (publishedStartVal != "" && publishedEndVal != "") {
        filterArray.push(`${filterProperties[4]}: range(${publishedStartVal.split('T')[0]}T23:59:59Z, ${publishedEndVal.split('T')[0]}T23:59:59Z)`)
      }
    }

    if (filterArray.length > 1) {
      if (searchTextRefinment != "") {
        //advanceFilterString += `and(${filterArray.join(',')},${searchTextRefinment})`;
        advanceFilterString += `and(${filterArray.join(',')})`;//, ${searchTextRefinment}
      }
      else {
        advanceFilterString += `and(${filterArray.join(',')})`;
      }
    }
    else if (filterArray.length == 1) {
      if (searchTextRefinment != "") {
        //advanceFilterString += `and(${filterArray.join(',')},${searchTextRefinment})`;
        //advanceFilterString += `and((${filterArray[0]}),or(${searchTextRefinment}))`
        advanceFilterString += `${filterArray[0]}`;
        //advanceFilterString += `and(${filterArray.join(',')})`;
      }
      else {
        advanceFilterString += filterArray[0];
      }

    }
    else {
      advanceFilterString += ``
    }

    console.log(advanceFilterString)

    return advanceFilterString


  }


  private sortListItems(listItems: any[], sortingFields, sortDirection) {

    const isDesc = sortDirection === 'desc' ? 1 : -1;
    // let sortFieldDetails = this.props.fields.filter(f => f.key === sortingFields)[0];

    switch (sortingFields) {

      case Constants.colPublishedDate: let sortFn: (a, b) => number;
        sortFn = (a, b) => ((new Date(a[sortingFields]).getTime() < new Date(b[sortingFields]).getTime()) ? 1 : -1) * isDesc;
        listItems.sort(sortFn);
        break;
      case Constants.colSubject: sortFn = (a, b) => ((a[sortingFields] > b[sortingFields]) ? 1 : -1) * isDesc;
        listItems.sort(sortFn);
        break;
      case Constants.colMigratedDepartment: sortFn = (a, b) => ((a[sortingFields] > b[sortingFields]) ? 1 : -1) * isDesc;
        listItems.sort(sortFn);
        break;
      case Constants.colCircularNumber: sortFn = (a, b) => ((a[sortingFields] > b[sortingFields]) ? 1 : -1) * isDesc;
        listItems.sort(sortFn);
        break;
      case Constants.colClassification: sortFn = (a, b) => ((a[sortingFields] > b[sortingFields]) ? 1 : -1) * isDesc;
        listItems.sort(sortFn);
        break;

    }

    return listItems;
  }

  private handleSorting = (property: string) => (event: React.MouseEvent<unknown>, column: IColumn) => {
    property = column.key;

    this.setState({ sortingFields: column.key }, () => {
      let { sortingFields, sortDirection, filteredItems } = this.state;
      //const isDesc = sortingFields && sortingFields === property && sortDirection === 'desc';
      const isDesc = property && sortingFields === property && sortDirection === 'desc';
      let updateColumns = this.state.columns.map(c => {
        //isSortedDescending: (isAsc ? false : true)
        //return c.key === property ? {...c, isSorted: true, isSortedDescending: (isDesc ? false : true) } : {...c};
        if (c.key == Constants.colPublishedDate) {
          return c.key === property ? { ...c, isSorted: true, isSortedDescending: !isDesc } : { ...c, isSorted: false, isSortedDescending: !c.isSortedDescending };
        }
        else {
          return c.key === property ? { ...c, isSorted: true, isSortedDescending: !c.isSortedDescending } : { ...c, isSorted: false, isSortedDescending: !c.isSortedDescending };
        }

      });

      this.setState({
        sortDirection: (isDesc ? 'asc' : 'desc'),
        sortingFields: property,
        columns: updateColumns,
      }, () => {
        const { sortDirection, sortingFields } = this.state;
        this.setState({ filteredItems: this.sortListItems(filteredItems, sortingFields, sortDirection) })
      });
    })

  }


  private circularSearchResultsTable = (): JSX.Element => {

    const { filteredItems } = this.state
    let filteredPageItems = this.paginateFn(filteredItems);
    const { columns } = this.state
    let searchResultsJSX = <>
      {this.detailListView(filteredPageItems, columns)}
    </>

    return searchResultsJSX;

  }


  private detailListView = (filteredPageItems, columns): JSX.Element => {
    let detailListViewJSX =
      <>
        <DetailsList
          className={` ${styles1.detailsListBorderRadius} `}
          styles={{
            root: {
              ".ms-DetailsHeader-cell": {

                ".ms-DetailsHeader-cellTitle": {
                  color: "white",
                  ".ms-Icon": {
                    color: "white",
                    fontWeight: 600,
                    left: -30
                  }
                }
              },
              ".ms-DetailsHeader-cell:hover": {

                background: "#f26522",
                color: "white",
                cursor: "pointer"
              }
            },
            focusZone: {
              ".ms-List": {
                ".ms-List-surface": {
                  ".ms-List-page": {
                    ".ms-List-cell": {
                      ".ms-DetailsRow": {

                        borderBottom: "1px solid #ccc",

                        ".ms-DetailsRow-fields": {
                          ".ms-DetailsRow-cell": {
                            fontWeight: 400,
                            fontSize: "13.5px",
                            fontFamily: 'Roboto',
                            color: "black"
                          }
                        }
                      },
                      ".ms-DetailsRow:hover": {
                        borderBottom: "1px solid #ccc",
                        background: "#f265221a"
                      }
                    }
                  }
                }
              }
            },
            headerWrapper: {
              ".ms-DetailsHeader": {
                color: "white",//"#003171",
                paddingTop: 0,
                backgroundColor: "#f26522"//"#495057" //"rgb(225 234 244)"//"#EEEFF0" //"#5581F6"//"rgb(3, 120, 124)"
              }
            }
          }}
          items={filteredPageItems}
          columns={columns}
          compact={true}
          selectionMode={SelectionMode.none}
          getKey={this._getKey}
          setKey="none"
          layoutMode={DetailsListLayoutMode.fixedColumns}
          isHeaderVisible={true}
          onItemInvoked={this._onItemInvoked}
          onRenderRow={this._onRenderRow}

        // onRenderDetailsHeader={(props, defaultRender) =>
        //   defaultRender({ ...props, styles: { root: { width: 200 } } })
        // }
        // onRenderDetailsFooter={this.createPagination.bind(this)}
        />
        {this.createPagination()}

      </>

    return detailListViewJSX;
  }

  private paginateFn = (filterItem: any[]) => {
    let { itemsPerPage, currentPage } = this.state;
    return (itemsPerPage > 0
      ? filterItem ? filterItem.slice((currentPage - 1) * itemsPerPage, (currentPage - 1) * itemsPerPage + itemsPerPage) : filterItem
      : filterItem
    );
  }

  private noItemFound = (): JSX.Element => {
    let noItemFoundJSX = <>

      <div className={`${styles1.OneUpError} `}>
        <div className={`${styles1.odError} `}>
          {/* <div className={`${ styles1.odErrorImage } `}>
            <Image src={require('../../assets/error2.svg')}
              styles={{
                root: {
                  display: 'inline-flex',
                  height: 230
                }
              }}></Image>
          </div> */}
          <div className={`${styles1.odErrorTitle} `}>No Circulars Found. Try to search circulars with relevant keywords</div>
        </div>

      </div>
    </>

    return noItemFoundJSX;
  }

  private _getKey(item: any, index?: number): string {
    return item?.key;
  }

  private _onItemInvoked(item: any): void {
  }

  private _onRenderRow: IDetailsListProps['onRenderRow'] = props => {
    const customStyles: Partial<IDetailsRowStyles> = {};
    if (props) {
      if (props.itemIndex % 2 === 0) {
        customStyles.root = { backgroundColor: '#fff' };
        customStyles.fields = { lineHeight: 25 }
      }
      else {
        customStyles.root = { background: "#e9ecef7a" };
        customStyles.fields = { lineHeight: 25 }
      }
      return <DetailsRow {...props} styles={customStyles} />;
    }
    return null;
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
            {<Label styles={{
              root: {
                paddingTop: 20,
                textAlign: "left",
                fontSize: isMobileMode ? 11 : 14,
                paddingLeft: 15,
                fontFamily: 'Roboto'
              }
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
    const formattedDate = dateObject.toLocaleDateString("en-UK", { day: "2-digit", month: "short", year: "numeric" });

    return `${formattedDate} `;
  }

  private createHyper(item: any): JSX.Element {
    const name = item?.Subject;

    return (
      <>
        <div className={styles.viewList}>
          <a onClick={this.readItemsAsStream.bind(this, item)}>{name}
            <Icon iconName="OpenInNewTab"></Icon>
          </a>
        </div>
      </>
    )
  }

  private renderTextField(item: any): JSX.Element {
    const category = item?.DocumentCategory?.Title;
    // const {documentCategoryDD} = this.state
    // const itemColor = documentCategoryDD.filter((val) => {
    //   return val.name == category
    // })
    // console.log(item.DocumentCategory?.FontColor + ":" + item.DocumentCategory?.BackgroundColor)
    return (
      <>
        <span style={{
          whiteSpace: "break-spaces",
          textAlign: "center", display: "block",
          //color: itemColor[0]?.color,//item.DocumentCategory?.FontColor ?? `black`
          //background: itemColor[0]?.backgroundColor,// item.DocumentCategory?.BackgroundColor ?? `white`,
          borderRadius: 15,
          fontWeight: 600
        }}>{category}</span>
      </>
    )
  }

  private renderDate(item: any): JSX.Element {
    const dateVal = this.formatDate(item.PublishedDate)
    return (
      <>
        <span style={{ textAlign: 'center', display: "inherit" }}>{item.PublishedDate != null ? dateVal : ``}</span>
      </>
    )
  }

  private readItemsAsStream(item: ICircularListItem, openFileViewer: boolean = false) {
    let providerContext = this.context;
    const { currentSelectedItem } = this.state
    const { services, serverRelativeUrl } = providerContext as IBobCircularRepositoryProps;

    this.setState({ isLoading: true }, async () => {
      if (currentSelectedItem.ID == item.ID) {
        await services.getListDataAsStream(serverRelativeUrl, Constants.circularList, item.Id).then((result) => {
          console.log(result.ListData);
          result.ListData.ID = item.Id;

          this.setState({
            previewItems: result.ListData,
            currentSelectedItemId: item.Id,
            filePreviewItem: result.ListData ?? null,
            isLoading: openFileViewer,
            openFileViewer: openFileViewer

          })

        }).catch((error) => {
          console.log(error);
          this.setState({ isLoading: false })
        })
      }

    })
  }

  public renderOwners(item: any): JSX.Element {
    const firstOwn = item?.PublisherEmailID;// item.Author.Title;
    return (<>
      {/* <Persona
        text={firstOwn}
        size={PersonaSize.size24}>
      </Persona> */}
      <span style={{ whiteSpace: "break-spaces" }}>{firstOwn}</span>
    </>)
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


  private handlePageChange(pageNo) {
    this.setState({ currentPage: pageNo });
  }

  private clearFilters = () => {

    this.setState({
      searchText: ``,
      selectedDepartment: [],
      circularNumber: ``,
      publishedStartDate: null,
      publishedEndDate: null,

    })

  }


  private downloadWaterMarkPDF = async (file, watermarkText) => {

    const pdfDoc = await PDFDocument.load(file.content);
    const totalPages = pdfDoc.getPageCount();

    for (let pageNum = 0; pageNum < totalPages; pageNum++) {
      const page = pdfDoc.getPage(pageNum);
      const { width, height } = page.getSize();
      const textFont = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
      const fontSize = 50;

      page.drawText(watermarkText, {
        x: width / 6,
        y: (1.6 * height) / 6,
        size: fontSize,
        font: textFont,
        opacity: 0.4,
        color: rgb(0.8392156862745098, 0.807843137254902, 0.792156862745098),
        rotate: degrees(30)
      });
    }

    let pdfBytes = await pdfDoc.save();

    let base64File = this.bufferToBase64(pdfBytes).then((val) => {
      console.log(val)
    }).catch((error) => {
      console.log(error)
    });

    download(pdfBytes, file.name, "application/pdf");

    return pdfDoc;
  }

  private bufferToBase64 = async (buffer): Promise<any> => {
    // use a FileReader to generate a base64 data URI:
    const base64url = await new Promise(r => {
      const reader = new FileReader()
      reader.onload = () => r(reader.result)
      reader.readAsDataURL(new Blob([buffer]))
    });

    // remove the `data:...;base64,` part from the start
    return Promise.resolve(base64url);
  }

  private pdfArray = async (file) => {
    // if (typeof file === Uint8Array) {
    //   return file;
    // }
    const fileURL = URL.createObjectURL(file);
    const data = await fetch(fileURL);
    const arrayBuffer = await data.arrayBuffer();
    return new Uint8Array(arrayBuffer);
    // or
    // return arrayBuffer;
  }

}

