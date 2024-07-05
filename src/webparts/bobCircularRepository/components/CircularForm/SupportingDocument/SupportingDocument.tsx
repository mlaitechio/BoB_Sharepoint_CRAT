import * as React from 'react';
import { ISupportingDocumentProps } from './ISupportingDocumentProps';
import { ISupportingDocumentState } from './ISupportingDocumentState';
import { IBasePickerSuggestionsProps, IInputProps, ITag, Panel, PanelType, TagPicker as Picker, ValidationState } from '@fluentui/react';
import { Button, Checkbox, FluentProvider, Label, webLightTheme } from '@fluentui/react-components';
import { Constants } from '../../../Constants/Constants';
import styles from '../../BobCircularRepository.module.scss';

export default class SupportingDocument extends React.Component<ISupportingDocumentProps, ISupportingDocumentState> {

    private tagPickerRef: any = React.createRef();


    public constructor(props) {
        super(props)

        this.state = {
            isSupportingPanelOpen: true,
            circularCollection: [],
            checkBoxCollection: new Map<string, any[]>(),
            filterPanelCheckBoxCollection: new Map<string, any[]>()
        }
    }

    public componentDidMount() {
        const { providerValue, department } = this.props;
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
        // Default Search will always be Circular Status as Published

        filterArray.push(`${filterProperties[5]}:equals("${Constants.published}")`);
        filterArray.push(`${filterProperties[3]}:equals("${department}")`);//${department} MSME BANKING DEPARTMENT IT PROJECTS AND CRM

        let filterString = `and(${filterArray.join(',')})`
        let sortListProperty = [{
            Property: Constants.managePropPublishedDate,
            Direction: 1 //0 for asc & 1 for descending
        }]

        await services.
            getSearchResults('', searchProperties, queryTemplate, filterString, sortListProperty).
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
                        IssuedFor: val.RefinableString08,

                    });
                })

                this.setState({ circularCollection: listItemData.map((val) => val.CircularNumber) }, () => {
                    let checkBoxCollection = new Map<string, any[]>();

                    checkBoxCollection.set(`${Constants.searchSupportingCirculars}`, listItemData.map((val) => {
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
                })
            }).catch((error) => {
                console.log(error);
                this.props.completeLoading();
            });

    }


    public render() {
        const { isSupportingPanelOpen, checkBoxCollection } = this.state
        let checkedCollectionLength = checkBoxCollection.get(Constants.searchSupportingCirculars)?.filter((val) => {
            return val.checked == true
        }).length;
        return (
            checkBoxCollection.size > 0 && <FluentProvider theme={webLightTheme}>
                <Panel isOpen={isSupportingPanelOpen}
                    isLightDismiss={true}
                    onDismiss={() => {
                        this.applyAll()
                        // const { checkBoxCollection } = this.state;
                        // this.setState({ isSupportingPanelOpen: false }, () => {
                        //     const { checkBoxCollection } = this.state
                        //     this.props.onDismiss()
                        // })
                    }}
                    type={PanelType.smallFixedFar}
                    onRenderFooterContent={() => <>
                        <FluentProvider theme={webLightTheme}>
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
                <div className={`${styles.row}`}>
                    {
                        allCirculars && allCirculars.length > 0 && allCirculars.map((circular, index) => {
                            return <div className={`${styles.column12}`}>
                                <Checkbox
                                    checked={circular.checked} //isChecked
                                    label={
                                        <Label weight="regular"
                                            onClick={() => { this.onCheckBoxLabelClick(index, Constants.searchSupportingCirculars) }} //this.onCheckBoxLabelClick.bind(this, labelName, index, !currentCheck)
                                            style={{ fontFamily: "Roboto", cursor: "pointer", textTransform: "capitalize" }}>
                                            {circular.value}
                                        </Label>
                                    }
                                    shape="square" size="medium"
                                    //this.onCheckBoxChange.bind(this, labelName, index)
                                    onChange={() => { this.onCheckBoxLabelClick(index, Constants.searchSupportingCirculars) }} />
                            </div>
                        })
                    }
                </div>
            </FluentProvider>
        </>;

        return checkBoxControlJSX;
    }

    private tagPicker = (inputProps: IInputProps, tagPickerRef: any, selectedItem?: any[]): JSX.Element => {

        const pickerSuggestionsProps: IBasePickerSuggestionsProps = {
            suggestionsHeaderText: `${inputProps.placeholder}`,
            noResultsFoundText: 'No results found!',
        };

        let tagPickerJSX = <>
            <div className={`${styles.row} `}>
                <div className={`${styles.column12} ${styles.marginFilterTop}`}>
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
            </div>
        </>;

        return tagPickerJSX;
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
        switch (selectedInput) {
            case `Search ${Constants.department}`: filters = circularCollection;
                break;

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
