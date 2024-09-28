import * as React from 'react';
import { ISupportingDocumentProps } from './ISupportingDocumentProps';
import { ISupportingDocumentState } from './ISupportingDocumentState';
import { IBasePickerSuggestionsProps, IInputProps, ITag, Panel, PanelType, TagPicker as Picker, ValidationState } from '@fluentui/react';
import { Button, Checkbox, Divider, Dropdown, Option, FluentProvider, Label, OptionOnSelectData, SelectionEvents, Theme, webLightTheme, Tooltip } from '@fluentui/react-components';
import { Constants } from '../../../Constants/Constants';
import styles from '../../BobCircularRepository.module.scss';
import { Info20Regular, Info24Regular, InfoRegular } from '@fluentui/react-icons';
import { Text } from '@microsoft/sp-core-library';

export const customLightTheme: Theme = {
    ...webLightTheme,
    colorBrandBackground: "#162B75",
    colorBrandBackgroundHover: "#162B75",
    colorBrandBackgroundSelected: "#162B75",
    colorBrandForegroundOnLightPressed: "#162B75",
    colorNeutralForeground2BrandHover: "#162B75",
    colorSubtleBackgroundHover: "#ffff",
    colorSubtleBackgroundPressed: "#ffff"
}

export default class SupportingDocument extends React.Component<ISupportingDocumentProps, ISupportingDocumentState> {

    private tagPickerRef: any = React.createRef();


    public constructor(props) {
        super(props)

        this.state = {
            isSupportingPanelOpen: true,
            circularCollection: [],
            allItems: [],
            checkBoxCollection: new Map<string, any[]>(),
            filterPanelCheckBoxCollection: new Map<string, any[]>(),
            optionsYear: ["All Years", "Previous Year"],
            selectedYear: `All Years`
        }
    }

    public componentDidMount() {
        this.setState({ isSupportingPanelOpen: true }, () => {
            this.searchResults()
        })
    }

    private searchResults = async () => {

        const { providerValue, department, selectedSupportingCirculars } = this.props
        const { context, circularListID, services } = providerValue;
        let siteID = context.pageContext.site.id;
        let webID = context.pageContext.web.id;
        let siteURL = context.pageContext.site.absoluteUrl;

        let listItemData: any[] = [];
        let searchProperties = Constants.selectedSearchProperties;

        let queryTemplate = `{searchTerms} (siteId:{${siteID}} OR siteId:${siteID}) (webId:{${webID}} OR webId:${webID}) (NormListID:${circularListID}) `;
        queryTemplate += `(path:"${siteURL}/Lists/${Constants.circularList}" OR ParentLink:"${siteURL}/Lists/${Constants.circularList}*") ContentTypeId:0x0* `;

        let filterProperties = Constants.filterSearchProperties;
        let filterArray = [];

        const { managePropClassification, managePropCircularType, published, lblMaster, lblCircular, unlimited, limited } = Constants;

        /**
        |--------------------------------------------------
        | Default Search will always be "Circular Status as Published" & 
        "Current Users Department" & 
        "Circular Type as unlimited" & 
        "last year start date will be 1st October"
        |--------------------------------------------------
        */

        filterArray.push(`${filterProperties[5]}:equals("${published}")`);
        filterArray.push(`${filterProperties[3]}:equals("${department}")`);//${department} Information Technology HR ADMINISTRATION MSME BANKING DEPARTMENT IT PROJECTS AND CRM
        filterArray.push(`${managePropClassification}:equals("${lblCircular}")`);


        let filterString = `and(${filterArray.join(',')})`
        let sortListProperty = [{
            Property: Constants.managePropPublishedDate,
            Direction: 1 //0 for asc & 1 for descending
        }]

        await services.
            getSupportingDocuments('', searchProperties, queryTemplate, filterString, sortListProperty).
            then(async (searchResults: any[]) => {
                searchResults?.map((val) => {

                    /**
                    |--------------------------------------------------
                    | RefinableString102 is Circular Type
                      CircularType ne null && CircularType -eq Unlimited

                    |--------------------------------------------------
                    */

                    if (val.RefinableString102 != null && (val.RefinableString102 == unlimited || val.RefinableString102 == limited)) {

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
                            IssuedFor: val.RefinableString08,
                        });
                    }
                })

                this.setState({ circularCollection: listItemData.map((val) => val.CircularNumber), allItems: listItemData }, () => {

                    this.allYears()

                })
            }).catch((error) => {
                console.log(error);
                this.props.completeLoading();
            });

    }


    private allYears = () => {
        const { selectedSupportingCirculars } = this.props
        let checkBoxCollection = new Map<string, any[]>();
        const { allItems } = this.state


        checkBoxCollection.set(`${Constants.searchSupportingCirculars}`, allItems.map((val) => {
            let isValueChecked = selectedSupportingCirculars?.filter((circular) => {
                return circular.CircularNumber == val.CircularNumber
            })
            return {
                checked: isValueChecked?.length > 0 ?? false,
                value: val.CircularNumber,
                ...val
            }
        }));

        this.setState({ checkBoxCollection }, () => {
            this.props.completeLoading();
        })
    }


    public render() {
        const { isSupportingPanelOpen, checkBoxCollection } = this.state
        let checkedCollectionLength = checkBoxCollection.get(Constants.searchSupportingCirculars)?.filter((val) => {
            return val.checked == true
        }).length;
        return (
            checkBoxCollection.size > 0 && <FluentProvider theme={customLightTheme}>
                <Panel isOpen={isSupportingPanelOpen}
                    isLightDismiss={true}

                    onDismiss={() => {
                        this.applyAll()
                    }}
                    type={PanelType.medium}
                    onRenderFooterContent={() => <>
                        <FluentProvider theme={customLightTheme}>
                            <div className={`${styles.row}`}>
                                <div className={`${styles.column12}`}>
                                    <Button appearance="primary"
                                        style={{ marginRight: 5 }}
                                        onClick={() => {
                                            this.applyAll()
                                        }}> Apply</Button>
                                    <Button
                                        onClick={() => {
                                            this.clearAll(Constants.searchSupportingCirculars)

                                        }} >Clear all</Button>
                                </div>
                            </div>
                        </FluentProvider>
                    </>}
                    closeButtonAriaLabel="Close"
                    // ${filterLabelName} (${checkedBoxes.length}
                    headerText={`Supporting Documents (${checkedCollectionLength})`}

                    styles={{
                        commands: { background: "white" },
                        footerInner: {
                            background: `white`,
                            borderTop: `1px solid lightgrey`
                        },
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
                    {this.tagPicker({ placeholder: `Search Circulars` }, this.tagPickerRef, [])}

                    {this.checkBoxControl()}
                </Panel>
            </FluentProvider>
        );
    }





    private checkBoxControl = (): JSX.Element => {
        const { circularCollection, checkBoxCollection } = this.state
        let allCirculars = checkBoxCollection.get(`${Constants.searchSupportingCirculars}`);

        let checkBoxControlJSX = <>
            <FluentProvider theme={webLightTheme}>

                <div className={`${styles.row} `} style={{ marginLeft: 1, marginTop: 5, marginBottom: 5 }}>
                    <div className={`${styles.column4}`}>
                        <Label className={`${styles.headerLabel}`}>Circular Number</Label>
                    </div>

                    <div className={`${styles.column1}`}>
                        <Label className={`${styles.headerLabel}`}>Date</Label>
                    </div>

                </div>

                <Divider appearance="subtle"></Divider>

                {
                    allCirculars && allCirculars.length > 0 && allCirculars.map((circular, index) => {
                        return <>
                            <div className={`${styles.row}`}>
                                <div className={`${styles.column4}`}>
                                    <Checkbox
                                        checked={circular.checked} //isChecked
                                        label={
                                            <Label weight="regular"
                                                onClick={() => { this.onCheckBoxLabelClick(index, Constants.searchSupportingCirculars) }}
                                                //this.onCheckBoxLabelClick.bind(this, labelName, index, !currentCheck)
                                                style={{
                                                    fontFamily: "Roboto", cursor: "pointer", textTransform: "capitalize", width: "120px",
                                                    display: "block",
                                                    textOverflow: "ellipsis",
                                                    overflow: "hidden"
                                                }}>
                                                {circular.value}
                                            </Label>
                                        }
                                        shape="square" size="medium"
                                        //this.onCheckBoxChange.bind(this, labelName, index)
                                        onChange={() => { this.onCheckBoxLabelClick(index, Constants.searchSupportingCirculars) }} />

                                </div >
                                {/* <div className={`${styles.column7}`}>
                                <Label style={{ fontFamily: "Roboto", cursor: "pointer", textTransform: "capitalize" }}>
                                    {circular.Subject}
                                </Label>
                            </div> */}
                                <div className={`${styles.column1}`} style={{ marginTop: 5 }}>
                                    <Label style={{ fontFamily: "Roboto", textTransform: "capitalize" }}>
                                        {new Date(circular.PublishedDate).toLocaleDateString("en-UK", { year: "numeric", month: "2-digit", day: "2-digit" })}
                                    </Label>
                                </div>
                            </div>
                            <Divider appearance="subtle"></Divider>
                        </>
                    })
                }

            </FluentProvider>
        </>;

        return checkBoxControlJSX;
    }

    private tagPicker = (inputProps: IInputProps, tagPickerRef: any, selectedItem?: any[]): JSX.Element => {

        const { optionsYear, selectedYear } = this.state
        const { configuration } = this.props;
        const { configVal } = Constants

        let allYearsToolTipLimit = configuration.filter(val => val.Title == configVal.AllYearsToolTipText)[0].Limit ?? 1000;
        let allYearsToolTipText = configuration.filter(val => val.Title == configVal.AllYearsToolTipText)[0].ToolTip ?? ``;
        let previousYearsToolTipLimit = configuration.filter(val => val.Title == configVal.PreviousYearToolTipText)[0].Limit ?? 1000;
        let previousYearsToolTipText = configuration.filter(val => val.Title == configVal.PreviousYearToolTipText)[0].ToolTip ?? ``;

        const pickerSuggestionsProps: IBasePickerSuggestionsProps = {
            suggestionsHeaderText: `${inputProps.placeholder}`,
            noResultsFoundText: 'No results found!',
        };

        let tagPickerJSX = <>
            <div className={`${styles.row} `}>
                <div className={`${styles.column4} ${styles.marginFilterTop}`}>
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
                        onBlur={this.onTagPickerBlur.bind(this, inputProps.placeholder)}
                        defaultSelectedItems={[]}
                    />
                </div>
                <div className={`${styles.column4} ${styles.marginFilterTop}`}>
                    <FluentProvider theme={webLightTheme}>
                        <Dropdown
                            style={{ maxWidth: 180, minWidth: 180 }}
                            defaultSelectedOptions={[optionsYear[0]]}
                            mountNode={{}} placeholder={`Year`}
                            value={selectedYear ?? ``}
                            selectedOptions={[selectedYear ?? ""]}
                            button={{ className: `${styles.formLabel}` }}
                            onOptionSelect={this.onDropDownChange.bind(this, `${Constants.sorting}`)}>
                            {optionsYear && optionsYear.length > 0 && optionsYear.map((val) => {
                                return <>
                                    <Option key={`${val}`} className={`${styles.formLabel}`}>
                                        {val}
                                    </Option>
                                </>
                            })}
                        </Dropdown>

                    </FluentProvider>
                </div >
                <div className={`${styles.column4} ${styles.marginFilterTop}`}>
                    <FluentProvider theme={webLightTheme}>
                        <Tooltip content={
                            <>
                                <Label className={`${styles.normalLabel}`}>
                                    {`${allYearsToolTipText.substring(0, allYearsToolTipLimit)}`}
                                </Label>
                                <Divider appearance="subtle"></Divider>
                                <Label className={`${styles.normalLabel}`}>
                                    {`${Text.format(previousYearsToolTipText, (new Date().getFullYear() - 1)).substring(0, previousYearsToolTipLimit)}`}
                                </Label>
                            </>
                        }
                            withArrow={true}
                            relationship="label">
                            <Info20Regular />
                        </Tooltip>
                    </FluentProvider>
                </div>
            </div>
        </>;

        return tagPickerJSX;
    }

    private onDropDownChange = (textLabel: string, event: SelectionEvents, data: OptionOnSelectData) => {

        const { allItems } = this.state;
        const { selectedSupportingCirculars } = this.props
        switch (textLabel) {
            case Constants.sorting: this.setState({
                selectedYear: data.optionValue,
                // sortingFields: data.optionValue == "Date" ? Constants.colPublishedDate : Constants.colSubject
            }, () => {

                const { selectedYear, } = this.state
                // const { filteredItems, sortingFields, sortDirection } = this.state;
                // this.setState({ filteredItems: this.sortListItems(filteredItems, sortingFields, sortDirection) })

                if (selectedYear == "Previous Year") {

                    let currentDate = new Date();
                    let lastYearStartDate = `${currentDate.getFullYear() - 1}` + `-10-01`;//${currentDate.getFullYear() - 1}
                    let currentMonth = currentDate.getMonth() + 1 < 10 ? `0` + (currentDate.getMonth() + 1) : (currentDate.getMonth() + 1);
                    let currentDay = currentDate.getDate() < 10 ? `0` + (currentDate.getDate()) : currentDate.getDate()
                    let currentEndDate = `${currentDate.getFullYear()}` + `-` + currentMonth;
                    currentEndDate += `-` + (currentDay); //+ `T23:59:59:59Z`;

                    let fromDate = new Date(lastYearStartDate);
                    let toDate = new Date(currentEndDate);

                    let previousYearCirculars = allItems.filter((val) => {
                        let publishedDate = new Date(val.PublishedDate);
                        return publishedDate >= fromDate && publishedDate <= toDate
                    });

                    let checkBoxCollection = new Map<string, any[]>();

                    checkBoxCollection.set(`${Constants.searchSupportingCirculars}`, previousYearCirculars.map((val) => {
                        let isValueChecked = selectedSupportingCirculars?.filter((circular) => {
                            return circular.CircularNumber == val.CircularNumber
                        })
                        return {
                            checked: isValueChecked?.length > 0 ?? false,
                            value: val.CircularNumber,
                            ...val
                        }
                    }));

                    this.setState({ circularCollection: previousYearCirculars.map((val) => val.CircularNumber), checkBoxCollection })

                }
                else {
                    this.allYears()
                }

            });
                break;
        }
    }

    private onTagPickerBlur = (selectedInput: string) => {
        switch (selectedInput) {
            case `${Constants.searchSupportingCirculars}`: this.tagPickerRef?.current?.input?.current?._updateValue("");
                break;
        }
    }

    private _onEmptyInputFocus = (selectedInput, filterText, tagList) => {
        const { circularCollection } = this.state;
        let filters = [];
        // switch (selectedInput) {
        //     case `Search ${Constants.department}`: filters = circularCollection;
        //         break;

        // }

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

        const { circularCollection } = this.state
        if (filter) {

            let filters = [];
            switch (selectedInput) {
                case `${Constants.searchSupportingCirculars}`: filters = circularCollection;
                    break;

            }

            return filters
                ?.filter((value: any) => value.toLowerCase().indexOf(filter.toLowerCase()) > -1)
                .map((value: any, index): any => {
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

        const { circularCollection, checkBoxCollection } = this.state;
        switch (selectedInput) {
            case `${Constants.searchSupportingCirculars}`:
                let circularIndex = circularCollection.indexOf(items[0].name);
                if (checkBoxCollection && checkBoxCollection.size > 0 && items.length > 0) {
                    checkBoxCollection.get(`${Constants.searchSupportingCirculars}`)[circularIndex].checked = true;
                    this.setState({ checkBoxCollection })
                };
                break;
        }
    }

    private onCheckBoxLabelClick = (index, labelName) => {
        const { checkBoxCollection } = this.state
        let currentCheck = checkBoxCollection.get(`${Constants.searchSupportingCirculars}`)[index].checked;
        switch (labelName) {
            case `${Constants.searchSupportingCirculars}`:
                checkBoxCollection.get(`${labelName}`)[index].checked = !currentCheck;
                this.setState({ checkBoxCollection });
                break;
        }

    }

    private applyAll = () => {
        const { checkBoxCollection } = this.state;
        let supportingCirculars = checkBoxCollection.get(`${Constants.searchSupportingCirculars}`).filter((val) => { return val.checked == true });
        this.setState({ isSupportingPanelOpen: false }, () => {
            this.props.onDismiss(supportingCirculars)
        })

    }

    private clearAll = (labelName) => {
        const { checkBoxCollection } = this.state;
        switch (labelName) {
            case `${Constants.searchSupportingCirculars}`: checkBoxCollection.get(`${Constants.searchSupportingCirculars}`).map((val) => {
                val.checked = false
            });
                this.setState({ checkBoxCollection }, () => {
                    //this.props.onDismiss([])
                });
                break;
        }
    }

}
