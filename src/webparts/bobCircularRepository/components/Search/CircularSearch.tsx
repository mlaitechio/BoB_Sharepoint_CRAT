import * as React from 'react'
import styles from './CircularSearch.module.scss';
import styles1 from '../BobCircularRepository.module.scss';
import {
  AnimationClassNames,
  DefaultButton, DetailsList, DetailsListLayoutMode, DetailsRow, DialogContent, getResponsiveMode,
  IBasePickerSuggestionsProps,
  IColumn, Icon, IDetailsListProps, IDetailsRowStyles,
  IInputProps, TagPicker as Picker,
  Image, Label, Panel, PanelType, PrimaryButton, SelectionMode, Stack,
  ValidationState,
  ITag
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
  Divider,
  SearchBox,
  SearchBoxChangeEvent,
  webLightTheme,
  Overflow,
  OverflowItem
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
import { AddCircleRegular, ArrowClockwise24Regular, ArrowClockwiseRegular, ArrowCounterclockwiseRegular, ArrowDownloadRegular, ArrowDownRegular, ArrowResetRegular, ArrowUpRegular, Attach12Filled, CalendarRegular, ChevronDownRegular, ChevronUpRegular, Dismiss24Regular, DismissRegular, EyeRegular, Filter12Regular, Filter16Regular, FilterRegular, OpenRegular, Search24Regular, ShareAndroidRegular, TextAlignJustifyRegular } from '@fluentui/react-icons';
import { ICheckBoxCollection, ICircularListItem } from '../../Models/IModel';
import { PDFDocument, StandardFonts, degrees, rgb } from 'pdf-lib';
import download from 'downloadjs'

export default class CircularSearch extends React.Component<ICircularSearchProps, ICircularSearchState> {

  static contextType = DataContext;
  context!: React.ContextType<typeof DataContext>;

  private tagPickerRef: any = React.createRef();
  private tagPickerRefYear: any = React.createRef();

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
      itemsPerPage: 9,
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
      openPanelCheckedValues: [],
      checkBoxCollection: new Map<string, ICheckBoxCollection[]>(),
      filterPanelCheckBoxCollection: new Map<string, ICheckBoxCollection[]>(),
      accordionFields: {
        isSummarySelected: false,
        isTypeSelected: false,
        isCategorySelected: false,
        isSupportingDocuments: false
      },
      isFilterPanel: false,
      filterLabelName: ``,
      filterAccordion: {
        isDepartmentSelected: false,
        isCircularNumberSelected: false,
        isPublishedYearSelected: false,
        isClassificationSelected: false,
        isIssuedForSelected: false,
        isComplianceSelected: false,
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
        Constants.expandColCircularRepository, 'PublishedDate', false).then(async (value) => {

          const uniqueDepartment: any[] = [...new Set(value.map((item) => {
            return item.Department;
          }))].sort((a, b) => a < b ? -1 : 1);

          const uniquePublishedYear: any[] = [...new Set(value.map((item) => {
            return new Date(item.PublishedDate).getFullYear().toString();
          }))];

          this.setState({
            items: value,
            filteredItems: value,
            departments: uniqueDepartment.filter((option) => {
              return option != undefined
            }),
            publishedYear: uniquePublishedYear
          }, () => {
            let checkBoxCollection = this.initializeCheckBoxFilter();
            this.setState({ checkBoxCollection: checkBoxCollection, isLoading: false }, () => {
              this.setState({ filterPanelCheckBoxCollection: checkBoxCollection })
            });
          })
        }).catch((error) => {
          console.log(error);
          this.setState({ isLoading: false })
        });



    })
  }


  private initializeCheckBoxFilter = (): Map<string, any[]> => {

    const { departments, publishedYear } = this.state;

    let checkBoxCollection = new Map<string, ICheckBoxCollection[]>();


    checkBoxCollection.set(`${Constants.colPublishedDate}`, publishedYear.map((val) => {
      return {
        checked: false,
        value: val,
        refinableString: "RefinableDate00"
      }
    }))

    checkBoxCollection.set(`${Constants.department}`, departments.map((val) => {
      return {
        checked: false,
        value: val,
        refinableString: "RefinableString03"
      } as ICheckBoxCollection
    }));



    checkBoxCollection.set(`${Constants.circularNumber}`,
      [{
        checked: true,
        value: `${Constants.lblContains}`,
        refinableString: "RefinableString00"
      },
      {
        checked: false,
        value: `${Constants.lblStartsWith}`,
        refinableString: "RefinableString00"
      },
      {
        checked: false,
        value: `${Constants.lblEndsWith}`,
        refinableString: "RefinableString00"
      }
      ]);

    checkBoxCollection.set(`${Constants.classification}`, [
      {
        checked: false,
        value: `${Constants.lblMaster}`,
        refinableString: "RefinableString06"
      },
      {
        checked: false,
        value: `${Constants.lblCircular}`,
        refinableString: "RefinableString06"
      }
    ]);

    checkBoxCollection.set(`${Constants.issuedFor}`, [
      {
        checked: false,
        refinableString: "RefinableString08",
        value: "India"
      },
      {
        checked: false,
        value: "Global",
        refinableString: "RefinableString08"
      }
    ]);

    checkBoxCollection.set(`${Constants.compliance}`, [
      {
        checked: false,
        value: "Yes",
        refinableString: "RefinableString09"
      },
      {
        checked: false,
        value: "No",
        refinableString: "RefinableString09"
      }
    ]);

    checkBoxCollection.set(`${Constants.category}`, [
      {
        checked: false,
        value: "Intimation",
        refinableString: "RefinableString04"
      },
      {
        checked: false,
        value: "Information",
        refinableString: "RefinableString04"
      },
      {
        checked: false,
        value: "Action",
        refinableString: "RefinableString04"
      }
    ]);

    return checkBoxCollection;

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
    let searchResultsColumn = isSearchNavOpen ? `${styles1.column10}` : `${styles1.column11}`


    let detailListClass = mobileMode ? `${styles1.column12}` : `${styles1.column12}`;

    return (<>

      {
        isLoading && this.workingOnIt()
      }

      <div className={`${styles1.row} ${styles1.marginFilterTop}`}>
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
          <Drawer type="inline" style={{ maxHeight: "200vh" }} separator open={isSearchNavOpen} className={`${styles1.column2}`}>
            <DrawerHeader style={{ padding: 16 }}>
              <DrawerHeaderTitle
                heading={{ className: `${styles1.fontRoboto}` }}
                className={`${styles1.formLabel}`}
                action={
                  <>
                    <Button
                      appearance="subtle"
                      aria-label="Reset"
                      icon={<ArrowCounterclockwiseRegular />}
                      onClick={() => { this.clearFilters() }}
                    >Reset Filters</Button>
                  </>
                }>
                <Filter16Regular />Refine
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
    const { circularNumber, checkBoxCollection, filterLabelName, filterAccordion } = this.state;
    let circularBox = checkBoxCollection.get(`${Constants.circularNumber}`);
    let departmentBox = checkBoxCollection.get(`${Constants.department}`);
    let publishedYearBox = checkBoxCollection.get(`${Constants.colPublishedDate}`);
    let categoryBox = checkBoxCollection.get(`${Constants.category}`);
    let regulatoryBox = checkBoxCollection.get(`${Constants.compliance}`);
    let issuedForBox = checkBoxCollection.get(`${Constants.issuedFor}`);
    let classificationBox = checkBoxCollection.get(`${Constants.classification}`);
    let searchFiltersJSX = <>
      {this.createFilterPanel(filterLabelName)}
      <div className={`${styles1.row}`}>
        <div className={`${styles1.column12}`} style={{ paddingLeft: 0 }}>

          <div className={`${styles1.row}`}>

            {/* {Department} */}
            <div className={`${styles1.column10} ${styles1.marginFilterTop} `}>
              <Field label={<FluentLabel weight="semibold" style={{ fontFamily: "Roboto" }}>{`${Constants.department}`}</FluentLabel>} ></Field>
            </div>
            <div className={`${styles1.column2} ${styles1.marginFilterTop} `}>
              <Button appearance="transparent"
                icon={filterAccordion.isDepartmentSelected ? <ChevronUpRegular /> : <ChevronDownRegular />}
                onClick={() => { this.onFilterAccordionClick(Constants.department) }}></Button>
            </div>

            <Divider appearance="subtle"></Divider>

            {
              filterAccordion.isDepartmentSelected && <>
                <div className={`${styles1.column12} ${styles1.marginFilterTop} ${AnimationClassNames.slideDownIn20}`}>
                  {/* {this.pickerControl()} */}
                  {checkBoxCollection.size > 0 && departmentBox.length > 0 && departmentBox.slice(0, 5).map((val, index) => {
                    return <div className={`${styles1.row}`}>
                      <div className={`${styles1.column12}`} >
                        {this.checkBoxControl(`${Constants.department}`, `${val.value}`, val.checked, index)}
                      </div>
                    </div>
                  })}
                </div>
                <div className={`${styles1.column12} ${styles1.marginFilterTop} ${AnimationClassNames.slideDownIn20}`} style={{ paddingLeft: 0 }}>
                  <Button icon={<OpenRegular />}
                    style={{ textDecoration: "underline", color: "var(--colorBrandBackground)" }}
                    iconPosition="after"
                    appearance="transparent" onClick={() => {
                      const { checkBoxCollection } = this.state
                      this.setState({
                        isFilterPanel: true,
                        filterPanelCheckBoxCollection: checkBoxCollection,
                        filterLabelName: `${Constants.department}`
                      })
                    }}>See All </Button>
                </div>
              </>
            }

            <Divider appearance="subtle"></Divider>

            {/* {Circular Number} */}
            <div className={`${styles1.column10} ${styles1.marginFilterTop} `}>
              <Field label={<FluentLabel weight="semibold" style={{ fontFamily: "Roboto" }}>{`${Constants.circularNumber}`}</FluentLabel>} ></Field>
            </div>
            <div className={`${styles1.column2} ${styles1.marginFilterTop} `}>
              <Button appearance="transparent"
                onClick={() => { this.onFilterAccordionClick(Constants.circularNumber) }}
                icon={filterAccordion.isCircularNumberSelected ? <ChevronUpRegular /> : <ChevronDownRegular />}></Button>
            </div>

            <Divider appearance="subtle"></Divider>

            {
              filterAccordion.isCircularNumberSelected && <>
                <div className={`${styles1.column12} ${styles1.marginFilterTop} ${AnimationClassNames.slideDownIn20}`} style={{ padding: 0 }}>
                  <Input placeholder="Input at least 2 characters"
                    input={{ className: `${styles.input}` }}
                    className={`${styles.input}`}
                    value={circularNumber}
                    onChange={this.onInputChange} />
                </div>
                <div className={`${styles1.column12} ${AnimationClassNames.slideDownIn20}`}>
                  {checkBoxCollection.size > 0 &&
                    this.checkBoxControl(`${Constants.circularNumber}`, `${Constants.lblContains}`, circularBox[0].checked, 0)}
                </div>
                <div className={`${styles1.column12} ${AnimationClassNames.slideDownIn20}`}>
                  {checkBoxCollection.size > 0 &&
                    this.checkBoxControl(`${Constants.circularNumber}`, `${Constants.lblStartsWith}`, circularBox[1].checked, 1)
                  }
                </div>

                <div className={`${styles1.column12} ${AnimationClassNames.slideDownIn20}`}>
                  {checkBoxCollection.size > 0 &&
                    this.checkBoxControl(`${Constants.circularNumber}`, `${Constants.lblEndsWith}`, circularBox[2].checked, 2)
                  }

                </div>
              </>
            }
          </div>

          {/* {Published Year} */}
          <div className={`${styles1.row} `}>
            <div className={`${styles1.column10} ${styles1.marginFilterTop}`}>
              <Field label={<FluentLabel weight="semibold" style={{ fontFamily: "Roboto" }}>{`Published Year`}</FluentLabel>} >
              </Field>
            </div>
            <div className={`${styles1.column2} ${styles1.marginFilterTop} `}>
              <Button appearance="transparent"
                onClick={() => { this.onFilterAccordionClick(Constants.publishedYear) }}
                icon={filterAccordion.isPublishedYearSelected ? <ChevronUpRegular /> : <ChevronDownRegular />}></Button>
            </div>
            <Divider appearance="subtle"></Divider>

            {filterAccordion.isPublishedYearSelected && <>
              <div className={`${styles1.column12} ${styles1.marginFilterTop} ${AnimationClassNames.slideDownIn20}`}>
                {/* {this.pickerControl()} */}
                {checkBoxCollection.size > 0 && publishedYearBox.length > 0 && publishedYearBox.slice(0, 5).map((val, index) => {
                  return <div className={`${styles1.row}`}>
                    <div className={`${styles1.column12}`} >
                      {this.checkBoxControl(`${Constants.colPublishedDate}`, `${val.value}`, val.checked, index)}
                    </div>
                  </div>
                })}
              </div>
              <div className={`${styles1.column12} ${styles1.marginFilterTop} ${AnimationClassNames.slideDownIn20}`} style={{ paddingLeft: 0 }}>
                <Button icon={<OpenRegular />}
                  style={{ textDecoration: "underline", color: "var(--colorBrandBackground)" }}
                  iconPosition="after"
                  appearance="transparent" onClick={() => {
                    const { checkBoxCollection } = this.state
                    this.setState({
                      isFilterPanel: true,
                      filterPanelCheckBoxCollection: checkBoxCollection,
                      filterLabelName: `${Constants.colPublishedDate}`
                    })
                  }}>See All</Button>
              </div>
            </>
            }
          </div>

          {/* {Classification} */}
          <div className={`${styles1.row} ${styles1.marginFilterTop}`}>
            <div className={`${styles1.column10} ${styles1.marginFilterTop} `}>
              <Field label={<FluentLabel weight="semibold" style={{ fontFamily: "Roboto" }}>{`${Constants.classification}`}</FluentLabel>} ></Field>
            </div>
            <div className={`${styles1.column2} ${styles1.marginFilterTop} `}>
              <Button
                appearance="transparent"
                onClick={() => { this.onFilterAccordionClick(Constants.classification) }}
                icon={filterAccordion.isClassificationSelected ? <ChevronUpRegular /> : <ChevronDownRegular />}></Button>
            </div>
            <Divider appearance="subtle"></Divider>
            {
              checkBoxCollection.size > 0 && filterAccordion.isClassificationSelected &&
              classificationBox.length > 0 && classificationBox.map((val, index) => {
                return <div className={`${styles1.column12} ${AnimationClassNames.slideDownIn20}`} >
                  {this.checkBoxControl(`${Constants.classification}`, `${val.value}`, val.checked, index)}
                </div>
              })
            }
          </div>

          {/* {Issued For} */}
          <div className={`${styles1.row} ${styles1.marginFilterTop}`}>
            <div className={`${styles1.column10} ${styles1.marginFilterTop} `}>
              <Field label={<FluentLabel weight="semibold" style={{ fontFamily: "Roboto" }}>{`${Constants.issuedFor}`}</FluentLabel>} >
              </Field>
            </div>
            <div className={`${styles1.column2} ${styles1.marginFilterTop} `}>
              <Button
                onClick={() => { this.onFilterAccordionClick(Constants.issuedFor) }}
                appearance="transparent"
                icon={filterAccordion.isIssuedForSelected ? <ChevronUpRegular /> : <ChevronDownRegular />}></Button>
            </div>
            <Divider appearance="subtle"></Divider>
            {
              checkBoxCollection.size > 0 && filterAccordion.isIssuedForSelected && issuedForBox.length > 0 && issuedForBox.map((val, index) => {
                return <div className={`${styles1.column12} ${AnimationClassNames.slideDownIn20}`} >
                  {this.checkBoxControl(`${Constants.issuedFor}`, `${val.value}`, val.checked, index)}
                </div>
              })
            }

          </div>

          {/* {Compliance} */}
          <div className={`${styles1.row} ${styles1.marginFilterTop}`}>
            <div className={`${styles1.column10} ${styles1.marginFilterTop} `}>
              <Field label={<FluentLabel weight="semibold" style={{ fontFamily: "Roboto" }}>{`${Constants.compliance}`}</FluentLabel>} ></Field>
            </div>
            <div className={`${styles1.column2} ${styles1.marginFilterTop} `}>
              <Button
                onClick={() => { this.onFilterAccordionClick(Constants.compliance) }}
                appearance="transparent"
                icon={filterAccordion.isComplianceSelected ? <ChevronUpRegular /> : <ChevronDownRegular />}></Button>
            </div>
            <Divider appearance="subtle"></Divider>
            {
              checkBoxCollection.size > 0 && filterAccordion.isComplianceSelected &&
              regulatoryBox.length > 0 && regulatoryBox.map((val, index) => {
                return <div className={`${styles1.column12} ${AnimationClassNames.slideDownIn20}`} >
                  {this.checkBoxControl(`${Constants.compliance}`, `${val.value}`, val.checked, index)}
                </div>
              })
            }

          </div>

          {/* {Category} */}
          <div className={`${styles1.row} ${styles1.marginFilterTop}`}>
            <div className={`${styles1.column10} ${styles1.marginFilterTop} `}>
              <Field label={<FluentLabel weight="semibold" style={{ fontFamily: "Roboto" }}>{`${Constants.category}`}</FluentLabel>} ></Field>
            </div>
            <div className={`${styles1.column2} ${styles1.marginFilterTop} `}>
              <Button
                onClick={() => { this.onFilterAccordionClick(Constants.category) }}
                appearance="transparent"
                icon={filterAccordion.isCategorySelected ? <ChevronUpRegular /> : <ChevronDownRegular />}></Button>
            </div>
            <Divider appearance="subtle"></Divider>
            {
              checkBoxCollection.size > 0 && filterAccordion.isCategorySelected && categoryBox.length > 0 && categoryBox.map((val, index) => {
                return <div className={`${styles1.column12} ${AnimationClassNames.slideDownIn20}`}>
                  {this.checkBoxControl(`${Constants.category}`, `${val.value}`, val.checked, index)}
                </div>
              })
            }
          </div>
          <div className={`${styles1.row}`}>
            <div className={`${styles1.column12} ${styles1.marginFilterTop} `}>
              {this.searchClearButtons()}
            </div>
          </div>

        </div>
      </div >
    </>;

    return searchFiltersJSX;

  }

  private searchFilterResults = (): JSX.Element => {
    const { filteredItems, isLoading, currentSelectedItemId, checkBoxCollection,
      previewItems, sortingOptions, selectedSortFields, sortDirection, isAccordionSelected } = this.state
    let filteredPageItems = this.paginateFn(filteredItems);


    let searchFilterResultsJSX = <>

      <div className={`${styles1.row} `}>

        {checkBoxCollection && checkBoxCollection.size > 0 &&

          <>
            <div className={`${styles1.row}`}>
              {this.selectedFilters()}
            </div>

            <Divider appearance="subtle"></Divider>
          </>
        }
      </div>

      <div className={`${styles1.row} `}>
        <div className={`${styles1.column9} ${styles1.marginTop}`}>
          {this.searchBox()}
        </div>
        <div className={`${styles1.column1} ${styles1.marginTop}`}>
          {this.searchClearButtons()}
        </div>
        <div className={`${styles1.column2} ${styles1.marginTop}`}>
          <Dropdown
            style={{ maxWidth: 95, minWidth: 95 }}
            mountNode={{}} placeholder={`Sorting`} value={selectedSortFields ?? ``}
            selectedOptions={[selectedSortFields ?? ""]}
            onOptionSelect={this.onDropDownChange.bind(this, `${Constants.sorting}`)}>
            {sortingOptions && sortingOptions.length > 0 && sortingOptions.map((val) => {
              return <><Option key={`${val}`} className={`${styles1.formLabel}`}>{val}</Option></>
            })}
          </Dropdown>
          <Button icon={sortDirection == "asc" ? <ArrowUpRegular /> : <ArrowDownRegular />} appearance="transparent"
            onClick={() => { this.onSorting() }} />
        </div>

        {/* <div className={`${styles1.column1} ${styles1.marginTop}`}>
          <Button icon={sortDirection == "asc" ? <ArrowUpRegular /> : <ArrowDownRegular />} appearance="transparent"
            onClick={() => { this.onSorting() }} />
        </div> */}

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

  private selectedFilters = (): JSX.Element => {
    const { checkBoxCollection } = this.state;
    let circularBox = checkBoxCollection.get(`${Constants.circularNumber}`)?.filter((val) => { return val.checked == true });
    let departmentBox = checkBoxCollection.get(`${Constants.department}`)?.filter((val) => { return val.checked == true });
    let publishedYearBox = checkBoxCollection.get(`${Constants.colPublishedDate}`)?.filter((val) => { return val.checked == true });
    let categoryBox = checkBoxCollection.get(`${Constants.category}`)?.filter((val) => { return val.checked == true });
    let regulatoryBox = checkBoxCollection.get(`${Constants.compliance}`)?.filter((val) => { return val.checked == true });
    let issuedForBox = checkBoxCollection.get(`${Constants.issuedFor}`)?.filter((val) => { return val.checked == true });
    let classificationBox = checkBoxCollection.get(`${Constants.classification}`)?.filter((val) => { return val.checked == true });
    let selectedFiltersJSX = <>

      {<div className={`${styles1.column12} ${styles1.marginFilterTop}`}>
        {this.badgeControl(departmentBox, `${Constants.department}`)}
        {this.badgeControl(publishedYearBox, `${Constants.colPublishedDate}`)}
        {this.badgeControl(classificationBox, `${Constants.classification}`)}
        {this.badgeControl(issuedForBox, `${Constants.issuedFor}`)}
        {this.badgeControl(regulatoryBox, `${Constants.compliance}`)}
        {this.badgeControl(categoryBox, `${Constants.category}`)}
      </div>
      }
    </>;

    return selectedFiltersJSX;
  }


  private badgeControl = (selectedFilters: ICheckBoxCollection[], labelName: string): JSX.Element => {
    let badgeJSX = <>{
      selectedFilters && selectedFilters.length > 0 && selectedFilters.map((val, index) => {
        return <Tag
          appearance="outline"
          as="button"
          dismissible={true}
          shape="circular"
          dismissIcon={<DismissRegular onClick={() => { this.onTagDismiss(labelName, val) }} />}
          className={`${styles1.tagClass}`}
          size="small"
          primaryText={{ style: { textOverflow: "ellipsis", fontFamily: 'Roboto' } }}
          title={`${val.value}`}>{val.value}</Tag>
      })
    }
    </>
    return badgeJSX;
  }

  private onTagDismiss = (labelName, selectedVal) => {
    const { checkBoxCollection } = this.state;
    let index = checkBoxCollection.get(labelName).indexOf(selectedVal);
    checkBoxCollection.get(labelName)[index].checked = false;
    this.setState({ checkBoxCollection });
  }

  private onFilterAccordionClick = (labelName) => {
    const { filterAccordion } = this.state;

    switch (labelName) {
      case `${Constants.department}`: filterAccordion.isDepartmentSelected = !filterAccordion.isDepartmentSelected;
        this.setState({ filterAccordion });
        break;

      case `${Constants.circularNumber}`: filterAccordion.isCircularNumberSelected = !filterAccordion.isCircularNumberSelected;
        this.setState({ filterAccordion });
        break;

      case `${Constants.publishedYear}`: filterAccordion.isPublishedYearSelected = !filterAccordion.isPublishedYearSelected;
        this.setState({ filterAccordion });
        break;
      case `${Constants.classification}`: filterAccordion.isClassificationSelected = !filterAccordion.isClassificationSelected;
        this.setState({ filterAccordion });
        break;
      case `${Constants.issuedFor}`: filterAccordion.isIssuedForSelected = !filterAccordion.isIssuedForSelected;
        this.setState({ filterAccordion });
        break;
      case `${Constants.compliance}`: filterAccordion.isComplianceSelected = !filterAccordion.isComplianceSelected;
        this.setState({ filterAccordion });
        break;
      case `${Constants.category}`: filterAccordion.isCategorySelected = !filterAccordion.isCategorySelected;
        this.setState({ filterAccordion });
        break;


    }
  }

  private createSearchResultsTable = (): JSX.Element => {
    const { filteredItems, previewItems, currentSelectedItemId, accordionFields } = this.state
    let filteredPageItems = this.paginateFn(filteredItems);
    const columns = [
      { columnKey: "Title", label: "Document Title" },
      { columnKey: "Date", label: "Date" },
      // { columnKey: "Classification", label: "Classification" },
      { columnKey: "Department", label: "Department" },
      { columnKey: "IssuedFor", label: "Issued For" }
    ];

    let tableJSX = <>
      <Table arial-label="Default table">
        <TableHeader>
          <TableRow >
            {columns.map((column, index) => (
              <TableHeaderCell key={column.columnKey} colSpan={index == 0 ? 6 : index == 2 ? 2 : 1} className={`${styles1.fontWeightBold}`}>
                {column.label}
              </TableHeaderCell>
            ))}
          </TableRow>
        </TableHeader>
        <TableBody>
          {filteredPageItems && filteredPageItems.length > 0 && filteredPageItems.map((val: ICircularListItem, index) => {

            let isFieldSelected = (accordionFields.isSummarySelected || accordionFields.isTypeSelected || accordionFields.isCategorySelected || accordionFields.isSupportingDocuments);
            let isCurrentItem = currentSelectedItemId == val.ID;
            let tableRowClass = isFieldSelected && isCurrentItem ? `${styles1.tableRow}` : ``;

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
                        style={{ padding: 0, fontWeight: 400 }}
                        appearance="transparent"
                        onClick={this.onDetailItemClick.bind(this, val, Constants.colSubject)}>
                        <div style={{
                          textAlign: "left",
                          marginTop: 5,
                          color: val.Classification == "Master" ? "#f26522" : "#162B75"
                        }}>{val.Subject} <OpenRegular /></div>
                      </Button>
                    </div>
                  </TableCellLayout>
                </TableCell>
                <TableCell>
                  <TableCellLayout>
                    {this.formatDate(val.PublishedDate)}
                  </TableCellLayout>
                </TableCell>
                {/* <TableCell>
                  <TableCellLayout content={{ style: { width: "100%" } }}
                    className={val.Classification == "Master" ? `${styles1.master}` : `${styles1.circular}`}>
                    {val.Classification}
                  </TableCellLayout>
                </TableCell> */}
                <TableCell colSpan={2}>
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
              {isFieldSelected && currentSelectedItemId == val.ID &&
                <TableRow >
                  <TableCell colSpan={6}>
                    <div className={`${styles1.row}`}>
                      <div className={`${styles1.column12} ${AnimationClassNames.slideDownIn20}`}>
                        {accordionFields.isSummarySelected &&
                          <>{`${previewItems?.Gist ?? ``}`}</>
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
            isSummarySelected: isCurrentItem ? !accordionFields.isSummarySelected : true,
            isTypeSelected: false,
            isCategorySelected: false,
            isSupportingDocuments: false
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
            isSummarySelected: false,
            isTypeSelected: isCurrentItem ? !accordionFields.isTypeSelected : true,
            isCategorySelected: false,
            isSupportingDocuments: false
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
            isSummarySelected: false,
            isTypeSelected: false,
            isCategorySelected: isCurrentItem ? !accordionFields.isCategorySelected : true,
            isSupportingDocuments: false
          },
          currentSelectedItem: item,
          currentSelectedItemId: item.ID
        }, () => {
          this.readItemsAsStream(item)
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
        this.readItemsAsStream(item)
      })
        break;

      case Constants.colSubject:

        this.setState({
          currentSelectedItem: item,
          currentSelectedItemId: item.ID
        }, () => {
          this.readItemsAsStream(item, true);
          if (currentSelectedItemId != item.ID) {
            this.setState({
              accordionFields: {
                isSummarySelected: false,
                isTypeSelected: false,
                isCategorySelected: false,
                isSupportingDocuments: false
              }
            })
          }

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



  private checkBoxControl = (labelName, checkBoxVal, isChecked, index): JSX.Element => {

    const { checkBoxCollection } = this.state
    let currentCheck = checkBoxCollection.get(`${labelName}`)[index].checked;

    let checkBoxJSX = <>
      <Checkbox
        checked={isChecked}
        //label={`${checkBoxVal}`}
        //style={{ fontFamily: "Roboto" }}
        label={
          <FluentLabel weight="regular"
            onClick={this.onCheckBoxLabelClick.bind(this, labelName, index, !currentCheck)}
            style={{ fontFamily: "Roboto", cursor: "pointer" }}>{checkBoxVal}</FluentLabel>
        }
        shape="square" size="medium" onChange={this.onCheckBoxChange.bind(this, labelName, index)} />
    </>

    return checkBoxJSX;
  }

  private onCheckBoxChange = (labelName: string, index, ev: React.ChangeEvent<HTMLInputElement>, data: CheckboxOnChangeData) => {

    this.onCheckBoxLabelClick(labelName, index, data.checked);

  }

  private onCheckBoxLabelClick = (labelName, index, isChecked) => {
    const { checkBoxCollection, isFilterPanel } = this.state;
    // const allBoxCollection = checkBoxCollection;
    // const boxColl = new Map<string, ICheckBoxCollection[]>();
    // allBoxCollection.forEach((val, key) => {
    //   boxColl.set(key, val);
    // });
    // const allCheckedCollection: Map<string, ICheckBoxCollection[]> = boxColl;
    let circularBox = checkBoxCollection.get(`${Constants.circularNumber}`);

    switch (labelName) {
      case `${Constants.circularNumber}`: checkBoxCollection.set(`${Constants.circularNumber}`,
        [{
          checked: index == 0 ? !circularBox[0].checked : false,
          value: `${Constants.lblContains}`,
          refinableString: "RefinableString00"
        },
        {
          checked: index == 1 ? !circularBox[1].checked : false,
          value: `${Constants.lblStartsWith}`,
          refinableString: "RefinableString00"
        },
        {
          checked: index == 2 ? !circularBox[2].checked : false,
          value: `${Constants.lblEndsWith}`,
          refinableString: "RefinableString00"
        }
        ]);
        this.setState({
          checkCircularRefiner: labelName,
          circularRefinerOperator: index == 0 ? `` : index == 1 ? `starts-with` : `ends-with`
        });
        break;

      case `${Constants.department}`:
        checkBoxCollection.get(`${labelName}`)[index].checked = isChecked;
        this.setState({ checkBoxCollection });
        break;
      case `${Constants.colPublishedDate}`:
        checkBoxCollection.get(`${labelName}`)[index].checked = isChecked;
        this.setState({ checkBoxCollection })
        break;

      case `${Constants.classification}`: checkBoxCollection.get(`${labelName}`)[index].checked = isChecked;
        this.setState({ checkBoxCollection })
        break;
      case `${Constants.issuedFor}`: checkBoxCollection.get(`${labelName}`)[index].checked = isChecked;
        this.setState({ checkBoxCollection })
        break;
      case `${Constants.compliance}`: checkBoxCollection.get(`${labelName}`)[index].checked = isChecked;
        this.setState({ checkBoxCollection })
        break;
      case `${Constants.category}`: checkBoxCollection.get(`${labelName}`)[index].checked = isChecked;
        this.setState({ checkBoxCollection })
        break;

    }
  }

  private createFilterPanel = (labelName): JSX.Element => {
    const { isFilterPanel, filterPanelCheckBoxCollection, filterLabelName } = this.state;
    let currentFilterCheckBox = filterPanelCheckBoxCollection.get(`${labelName}`);
    let checkedBoxes = [];
    if (currentFilterCheckBox) {
      checkedBoxes = currentFilterCheckBox.filter((val) => {
        return val.checked == true;
      });
    }

    let filterPanelJSX = <>
      <Panel isOpen={isFilterPanel}
        isLightDismiss={true}
        onDismiss={() => {
          const { checkBoxCollection } = this.state;
          this.setState({ checkBoxCollection, isFilterPanel: false })
        }}
        type={PanelType.smallFixedFar}
        onRenderFooterContent={() => <>
          <FluentProvider theme={webLightTheme}>
            {/* <Button appearance="primary"
              style={{ marginRight: 5 }}
              onClick={() => { this.applyFilters(labelName) }}
              disabled={checkedBoxes.length > 0 ? false : true}>Apply</Button> */}
            <Button onClick={() => { this.clearAll(filterLabelName) }} >Clear all</Button>
          </FluentProvider>
        </>}
        closeButtonAriaLabel="Close"
        headerText={`Filter ${filterLabelName} (${checkedBoxes.length})`}
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
        }} >
        <div className={`${styles1.row} ${styles1.marginFilterTop}`}>
          <div className={`${styles1.column12}`} style={{ paddingLeft: 0 }}>
            {filterLabelName == Constants.department && this.tagPicker({ placeholder: `Search ${Constants.department}` }, this.tagPickerRef, [])}
            {filterLabelName == Constants.colPublishedDate && this.tagPicker({ placeholder: `Search ${Constants.publishedYear}` }, this.tagPickerRefYear, [])}

          </div>
        </div>
        {filterPanelCheckBoxCollection.size > 0 && currentFilterCheckBox?.length > 0 && currentFilterCheckBox?.map((val, index) => {
          return <div className={`${styles1.row}`}>

            <div className={`${styles1.column12}`} style={{ paddingLeft: 0, paddingRight: 0 }}>
              <FluentProvider theme={webLightTheme}>
                {this.checkBoxControl(`${labelName}`, val.value, val.checked, index)}
              </FluentProvider>
            </div>
          </div>

        })}
      </Panel>
    </>
    return filterPanelJSX;
  }

  private applyFilters = (labelName) => {
    const { isFilterPanel, checkBoxCollection, openPanelCheckedValues } = this.state;
    if (isFilterPanel && openPanelCheckedValues && openPanelCheckedValues.length > 0) {
      openPanelCheckedValues.map((val) => {
        let checkedIndex = checkBoxCollection.get(`${labelName}`).indexOf(val);
        if (checkedIndex > -1) {
          checkBoxCollection.get(`${labelName}`)[checkedIndex].checked = true;
        }
      });
      this.setState({ checkBoxCollection, filterPanelCheckBoxCollection: checkBoxCollection, isFilterPanel: false })
    }
  }

  private clearAll = (labelName?: string) => {
    const { checkBoxCollection } = this.state;
    switch (labelName) {
      case `${Constants.department}`: checkBoxCollection.get(`${Constants.department}`).map((val) => {
        val.checked = false
      });

        this.setState({ checkBoxCollection, filterPanelCheckBoxCollection: checkBoxCollection })

        break;

      case `${Constants.colPublishedDate}`: checkBoxCollection.get(`${Constants.colPublishedDate}`).map((val) => {
        val.checked = false
      });

        this.setState({ checkBoxCollection, filterPanelCheckBoxCollection: checkBoxCollection })
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

  private tagPicker = (inputProps: IInputProps, tagPickerRef: any, selectedItem?: any[]): JSX.Element => {

    const pickerSuggestionsProps: IBasePickerSuggestionsProps = {
      suggestionsHeaderText: `${inputProps.placeholder}`,
      noResultsFoundText: 'No results found!',
    };

    let tagPickerJSX = <>

      <Picker
        componentRef={tagPickerRef}
        onResolveSuggestions={this.onResolveSuggestions.bind(this, inputProps.placeholder)}
        itemLimit={1}
        getTextFromItem={this.getTextFromItem}
        pickerSuggestionsProps={pickerSuggestionsProps}
        inputProps={inputProps}
        onEmptyResolveSuggestions={(tagList) => this._onEmptyInputFocus(inputProps.placeholder, '', tagList)}
        onChange={this.onPickerChange.bind(this, inputProps.placeholder)}
        styles={{ text: { fontSize: 13.5, fontWeight: 600 } }}
        selectedItems={[]}
        onValidateInput={this.onValidateInput}
        //onInputChange={this.onInputChange}
        onBlur={this.onTagPickerBlur.bind(this, inputProps.placeholder)}
        //onItemSelected={this.onFilterItemSelection.bind(this, isFinancial)}
        defaultSelectedItems={[]} //isFinancial ? [{name:"FY 2023 - 2024",key:"FY 2023 - 2024|0"}] : []
      />
    </>;

    return tagPickerJSX;
  }

  private onTagPickerBlur = (selectedInput: string) => {
    switch (selectedInput) {
      case `Search ${Constants.department}`: this.tagPickerRef?.current?.input?.current?._updateValue("");
        break;
      case `Search ${Constants.publishedYear}`: this.tagPickerRefYear?.current?.input?.current?._updateValue("");
        break
    }

  }

  private _onEmptyInputFocus = (selectedInput, filterText, tagList) => {
    const { departments, publishedYear } = this.state;
    let filters = [];
    switch (selectedInput) {
      case `Search ${Constants.department}`: filters = departments;
        break;
      case `Search ${Constants.publishedYear}`: filters = publishedYear;
        break
    }

    return []

    // return filters
    //   .map((value: any, index): any => {
    //     return {
    //       name: value,
    //       key: value
    //     };
    //   });
  }

  private onValidateInput = (input: string): ValidationState => {
    return input ? ValidationState.valid : ValidationState.invalid;
  }

  private onResolveSuggestions = async (selectedInput: string, filter: string,
    selectedItems: any[] | undefined): Promise<ITag[]> => {

    const { departments, publishedYear } = this.state
    if (filter) {

      let filters = [];
      switch (selectedInput) {
        case `Search ${Constants.department}`: filters = departments;
          break;
        case `Search ${Constants.publishedYear}`: filters = publishedYear;
          break
      }

      return filters
        .filter((value: any) => value.toLowerCase().indexOf(filter.toLowerCase()) > -1)
        .map((value: any, index): any => {
          // console.log(category.id)
          return {
            name: value,
            key: value
          };
        });
    }

    return []
  }

  private getTextFromItem = (item: ITag) => {
    return item.name;
  }

  private onPickerChange = (selectedInput, items?: ITag[] | undefined) => {

    const { checkBoxCollection, departments, publishedYear } = this.state;

    let filters = [];
    switch (selectedInput) {
      case `Search ${Constants.department}`: let departmentBoxIndex = departments.indexOf(items[0].name);

        if (checkBoxCollection && checkBoxCollection.size > 0 && items.length > 0) {
          checkBoxCollection.get(`${Constants.department}`)[departmentBoxIndex].checked = true;
          this.setState({ filterPanelCheckBoxCollection: checkBoxCollection })
        };
        break;
      case `Search ${Constants.publishedYear}`: let publishedYearIndex = publishedYear.indexOf(items[0].name);

        if (checkBoxCollection && checkBoxCollection.size > 0 && items.length > 0) {
          checkBoxCollection.get(`${Constants.colPublishedDate}`)[publishedYearIndex].checked = true;
          this.setState({ filterPanelCheckBoxCollection: checkBoxCollection })
        }
        break
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

  private onResetClick = (labelName: string) => {
    switch (labelName) {
      case `FromDate`: this.setState({ publishedStartDate: null });
        break;
      case `ToDate`: this.setState({ publishedEndDate: null });
        break;

    }
  }

  private updateItem = (itemID: any) => {
    this.props.onUpdateItem(itemID)
  }

  private onPanelClose = () => {
    this.setState({ filePreviewItem: null, openFileViewer: false })
  }


  private searchClearButtons = (): JSX.Element => {
    let searchClearJSX = <>
      <FluentUIBtn appearance="primary" style={{ marginRight: 2 }} icon={<FilterRegular />} onClick={() => { this.searchResults() }}>
        Search
      </FluentUIBtn>
      {/* <FluentUIBtn appearance="secondary" icon={<DismissRegular />} onClick={() => { this.clearFilters() }}>
        Clear
      </FluentUIBtn> */}
    </>;

    return searchClearJSX;

  }

  private searchBox = (): JSX.Element => {
    const { searchText } = this.state;
    let searchBoxJSX =
      <>
        <SearchBox appearance="underline"
          onChange={this.onSearchTextChange}
          placeholder="Type here to search"
          onKeyUp={this.handleSearchEvent}
          input={{ className: `${styles1.fontRoboto}` }}
          style={{ width: "100%", maxWidth: "100%" }} />
      </>
    // <Stack tokens={{ childrenGap: 20 }}>
    //   <SearchBox
    //     placeholder={`${Constants.searchText} `}
    //     onChange={this.onSearchBoxChange}
    //     onSearch={this.handleSearch}
    //     //onClear={(ev?: any) => { this.onClear() }}
    //     onClear={(ev?: any) => { this.onSearchClear() }}
    //     defaultValue={searchText}
    //     value={searchText}
    //     styles={{
    //       root: {
    //         border: "1px solid #bac6f7",
    //         fontFamily: "Roboto",
    //         borderRadius: 5,
    //         ":hover": {
    //           borderColor: "1px solid #bac6f7"
    //         },
    //         ".is-active": {
    //           border: "1px solid #bac6f7"
    //         }
    //       }
    //     }}
    //   />
    // </Stack>;



    return searchBoxJSX;
  }

  private onSearchBoxChange = (event?: React.ChangeEvent<HTMLInputElement>, newValue?: string) => {
    this.setState({ searchText: newValue })
  }

  private onSearchTextChange = (event: SearchBoxChangeEvent, data: InputOnChangeData) => {
    this.setState({ searchText: data.value })
  }

  private handleSearchEvent = (event?: any) => {
    if (event.keyCode === 13) {
      const { searchText } = this.state;
      this.handleSearch(searchText)
    }
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
      | This for combination of Search Text + Refinment Filters. 
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
      publishedStartDate, publishedEndDate, checkBoxCollection } = this.state;

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
        RefinableString08 -> IssuedFor
        RefinableString09 -> Compliance
   
    |--------------------------------------------------
    */
    let departmentVal = selectedDepartment[0] ?? ``;//RefinableString03
    let circularVal = circularNumber != "" ? circularNumber : ``;
    let publishedStartVal = publishedStartDate?.toISOString() ?? ``;//RefinableDate00
    let publishedEndVal = publishedEndDate?.toISOString() ?? ``;


    let advanceFilterString = "";
    let checkBoxFilterString = "";

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
    let checkBoxFilter = [];

    let publishedYears = checkBoxCollection.get(`${Constants.colPublishedDate}`).filter((val) => val.checked == true).map((val) => {
      return parseInt(val.value);
    });

    let publishedStartYear = null;
    let publishedEndYear = null;
    if (publishedYears.length > 0) {
      publishedStartYear = Math.min(...publishedYears) + `-01-01`;
      publishedEndYear = Math.max(...publishedYears) + `-12-31`;
    }


    // Default Search will always be Circular Status as Published
    filterArray.push(`${filterProperties[5]}:equals("${Constants.published}")`);

    if (!isNormalSearch) {
      `${departmentVal != "" ? filterArray.push(`${filterProperties[3]}:equals("${departmentVal}")`) : ``} `;
      `${circularVal != "" ? filterArray.push(`${filterProperties[0]}:${circularRefinerOperator}("${circularVal}*")`) : ``} `;
      // if (publishedStartVal != "" && publishedEndVal != "") {
      //   filterArray.push(`${filterProperties[4]}: range(${publishedStartVal.split('T')[0]}T23:59:59Z, ${publishedEndVal.split('T')[0]}T23:59:59Z)`)
      // }

      if (publishedStartYear != null && publishedEndYear != null) {
        filterArray.push(`${filterProperties[4]}: range(${publishedStartYear}T23:59:59Z, ${publishedEndYear}T23:59:59Z)`)
      }




      if (checkBoxCollection.size > 0) {
        checkBoxCollection.forEach((checkMap) => {
          let checkMapColl = checkMap.filter((val) => {
            return val.checked == true && val.refinableString != "RefinableString00" && val.refinableString != "RefinableDate00";
          })
          if (checkMapColl.length > 1) {
            checkMapColl.map((val) => {
              checkBoxFilter.push(`${val.refinableString}:equals("${val.value}")`);
            })
            checkBoxFilterString += `or(${checkBoxFilter.join(',')}),`;
          }
          else if (checkMapColl.length == 1) {
            checkMapColl.map((val) => {
              checkBoxFilterString += `${val.refinableString}:equals("${val.value}"),`;
            });
          }
          else {
            checkBoxFilterString += "";
          }
        })
      }
    }

    if (checkBoxFilterString != "") {
      checkBoxFilterString = checkBoxFilterString.substring(0, checkBoxFilterString.length - 1)
    }


    if (filterArray.length > 1 || checkBoxFilter.length > 1) {
      if (searchTextRefinment != "") {
        //  advanceFilterString += `and(${filterArray.join(',')})`;// ${searchTextRefinment}//,or(${searchTextRefinment}))

        advanceFilterString += `and(${filterArray.join(',')},${searchTextRefinment}${checkBoxFilterString != "" ? `,${checkBoxFilterString}` : ``})`;
      }
      else {
        advanceFilterString += `and(${filterArray.join(',')}${checkBoxFilterString != "" ? `,${checkBoxFilterString}` : ``})`
      }
    }
    else if (filterArray.length == 1 || checkBoxFilterString != "") {
      if (searchTextRefinment != "") {
        advanceFilterString += `and(${filterArray.join(',')},${searchTextRefinment}${checkBoxFilterString != "" ? `,${checkBoxFilterString}` : ``})`;
        //advanceFilterString += `and(${filterArray.join(',')})`;
        //advanceFilterString += filterArray[0];
      }
      else {
        if (filterArray.length == 1 && checkBoxFilterString != "") {
          advanceFilterString += `and(` + filterArray[0] + `${checkBoxFilterString != "" ? `,${checkBoxFilterString}` : ``})`;
        }
        else if (filterArray.length == 1) {
          advanceFilterString += `${filterArray.join(',')}`
        }
        else if (checkBoxFilterString != "") {
          advanceFilterString += `${checkBoxFilterString}`;
        }
        else {
          advanceFilterString += ``;
        }

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

    console.log("Search Query & Filter")

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
      publishedStartDate, publishedEndDate, checkBoxCollection } = this.state;

    /**
    |--------------------------------------------------
    |  |RefinableString00 -> CircularNumber
        RefinableString01 -> Subject
        RefinableString02 -> Migrated Department
        RefinableString03 -> Department
        RefinableString04 -> Category
        RefinableString05 -> IsMigrated 
        RefinableString06 -> Classification
        RefinableDate00 -> PublishedDate 
        RefinableString07 -> CircularStatus
        RefinableString08 -> IssuedFor
        RefinableString09 -> Compliance
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

    let checkBoxFilterString = "";

    let checkBoxFilter = [];

    // Default Search will always be Circular Status as Published
    filterArray.push(`${filterProperties[5]}:equals("${Constants.published}")`);

    let publishedYears = checkBoxCollection.get(`${Constants.colPublishedDate}`).filter((val) =>
      val.checked == true
    ).map((val) => {
      return parseInt(val.value);
    });

    let publishedStartYear = null;
    let publishedEndYear = null;
    if (publishedYears.length > 0) {
      publishedStartYear = Math.min(...publishedYears) + `-01-01`;
      publishedEndYear = Math.max(...publishedYears) + `-12-31`;
    }

    if (!isNormalSearch) {
      `${departmentVal != "" ? filterArray.push(`${filterProperties[3]}:equals("${departmentVal}")`) : ``} `;
      `${circularVal != "" ? filterArray.push(`${filterProperties[0]}:${circularRefinerOperator}("${circularVal}*")`) : ``} `;
      // if (publishedStartVal != "" && publishedEndVal != "") {
      //   filterArray.push(`${filterProperties[4]}: range(${publishedStartVal.split('T')[0]}T23:59:59Z, ${publishedEndVal.split('T')[0]}T23:59:59Z)`)
      // }

      if (publishedStartYear != null && publishedEndYear != null) {
        filterArray.push(`${filterProperties[4]}: range(${publishedStartYear}T23:59:59Z, ${publishedEndYear}T23:59:59Z)`)
      }


      if (checkBoxCollection.size > 0) {
        checkBoxCollection.forEach((checkMap) => {
          let checkMapColl = checkMap.filter((val) => {
            return val.checked == true && val.refinableString != "RefinableString00" && val.refinableString != "RefinableDate00";
          })
          if (checkMapColl.length > 1) {
            checkMapColl.map((val) => {
              checkBoxFilter.push(`${val.refinableString}:equals("${val.value}")`);
            })
            checkBoxFilterString += `or(${checkBoxFilter.join(',')}),`;
          }
          else if (checkMapColl.length == 1) {
            checkMapColl.map((val) => {
              checkBoxFilterString += `${val.refinableString}:equals("${val.value}"),`;
            });
          }
          else {
            checkBoxFilterString += "";
          }
        })
      }
    }

    if (checkBoxFilterString != "") {
      checkBoxFilterString = checkBoxFilterString.substring(0, checkBoxFilterString.length - 1)
    }

    if (filterArray.length > 1 || checkBoxFilter.length > 1) {
      advanceFilterString += `and(${filterArray.join(',')}${checkBoxFilterString != "" ? `,${checkBoxFilterString}` : ``})`;
    }
    else if (filterArray.length == 1 || checkBoxFilterString != "") {
      if (filterArray.length == 1 && checkBoxFilterString != "") {
        advanceFilterString += `and(${filterArray.join(',')}${checkBoxFilterString != "" ? `,${checkBoxFilterString}` : ``})`;
      }
      else if (filterArray.length == 1) {
        advanceFilterString += `${filterArray.join(',')}`
      }
      else if (checkBoxFilterString != "") {
        advanceFilterString += `${checkBoxFilterString}`;
      }
      else {
        advanceFilterString += ``;
      }
    }


    // if (filterArray.length > 1 && checkBoxFilter.length > 1) {
    //   if (searchTextRefinment != "") {

    //     advanceFilterString += filterArray.length > 1 ? `and(` : ``;
    //     advanceFilterString += filterArray.length > 1 ? `${filterArray.join(',')}` : ``
    //     advanceFilterString += checkBoxFilterString != "" ? `,${checkBoxFilterString}` : ``;
    //     advanceFilterString += filterArray.length > 1 ? `)` : ``;
    //   }
    //   else {
    //     advanceFilterString += filterArray.length > 1 ? `and(` : ``;
    //     advanceFilterString += filterArray.length > 1 ? `${filterArray.join(',')}` : ``
    //     advanceFilterString += checkBoxFilterString != "" ? `,${checkBoxFilterString}` : ``;
    //     advanceFilterString += filterArray.length > 1 ? `)` : ``;
    //   }
    // }
    // else if ((filterArray.length > 1 || filterArray.length == 1) && checkBoxFilterString == "") {
    //   advanceFilterString += filterArray.length > 1 ? `and(` : ``;
    //   advanceFilterString += (filterArray.length > 1 || filterArray.length == 1) ? `${filterArray.join(',')}` : ``;
    //   advanceFilterString += filterArray.length > 1 ? `)` : ``;
    // }
    // else if (filterArray.length == 1 && checkBoxFilterString != "") {
    //   advanceFilterString += filterArray.length == 1 ? `and(` : ``;
    //   advanceFilterString += filterArray.length == 1 ? `${filterArray.join(',')}` : ``;
    //   advanceFilterString += checkBoxFilterString != "" ? `,${checkBoxFilterString}` : ``;
    //   advanceFilterString += filterArray.length == 1 ? `)` : ``;
    // }
    // else if (filterArray.length == 1 || checkBoxFilterString != "") {
    //   if (searchTextRefinment != "") {
    //     advanceFilterString += filterArray.length == 1 && checkBoxFilterString != "" ? `and(` : ``;
    //     advanceFilterString += filterArray.length == 1 ? `${filterArray[0]}` : ``;
    //     advanceFilterString += checkBoxFilterString != "" ? `,${checkBoxFilterString}` : ``;
    //     advanceFilterString += filterArray.length == 1 && checkBoxFilterString != "" ? `)` : ``;

    //   }
    //   else {
    //     advanceFilterString += filterArray.length == 1 && checkBoxFilterString != "" ? `and(` : ``;
    //     advanceFilterString += filterArray.length == 1 ? `${filterArray[0]}` : ``;
    //     advanceFilterString += checkBoxFilterString != "" ? `,${checkBoxFilterString}` : ``;
    //     advanceFilterString += filterArray.length == 1 && checkBoxFilterString != "" ? `)` : ``;
    //   }
    // }
    // else {
    //   advanceFilterString += ``
    // }

    console.log("Refinable Filter")

    console.log(advanceFilterString);

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
    const formattedDate = dateObject.toLocaleDateString("en-UK", { day: "numeric", month: "short", year: "numeric" });

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
      checkBoxCollection: this.initializeCheckBoxFilter()
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

