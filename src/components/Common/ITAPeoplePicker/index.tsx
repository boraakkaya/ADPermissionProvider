import * as React from 'react';
import {
    BaseComponent,
    assign,
    autobind
} from 'office-ui-fabric-react/lib/Utilities';
import {
    BaseClientSideWebPart,
    IPropertyPaneConfiguration,
    IWebPartContext,
    PropertyPaneTextField,
    WebPartContext
} from '@microsoft/sp-webpart-base';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { IPersonaProps, Persona } from 'office-ui-fabric-react/lib/Persona';
import {
    CompactPeoplePicker,
    IBasePickerSuggestionsProps,
    IBasePicker,
    ListPeoplePicker,
    NormalPeoplePicker,
    ValidationState
} from 'office-ui-fabric-react/lib/Pickers';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { IPersonaWithMenu } from 'office-ui-fabric-react/lib/components/pickers/PeoplePicker/PeoplePickerItems/PeoplePickerItem.Props';
import { people } from './mockData';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { connect } from 'react-redux';
import {
    Environment,
    EnvironmentType
} from '@microsoft/sp-core-library';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
const suggestionProps: IBasePickerSuggestionsProps = {
    suggestionsHeaderText: 'Suggested People',
    mostRecentlyUsedHeaderText: 'Suggested Contacts',
    noResultsFoundText: 'No results found',
    loadingText: 'Loading',
    showRemoveButtons: true,
    suggestionsAvailableAlertText: 'People Picker Suggestions available',
    suggestionsContainerAriaLabel: 'Suggested contacts'
};
//extended Persona to hold userID
export interface ITAPersonaProps extends IPersonaWithMenu {
    userID: number
}
export interface ITAPeoplePickerProps {
    /**
     * Number of items alllowed in this picker
     */
    itemLimit?: number;
    spContext: IWebPartContext | WebPartContext;
    /**
     * Callback when Compact Picker items change.
     * Returns selected items as first parameter as IPersonaProps array.
     */
    onChange: (a: ITAPersonaProps[]) => any;
    defaultItems?: ITAPersonaProps[];
}
export interface ITAPeoplePickerState {
    peopleList: ITAPersonaProps[];
}

/**
 * Extends compact people picker. Uses site users as data source instead of search since we will also need the ID of the user for that specific site collection.
 */
export class ITAPeoplePicker extends React.Component<ITAPeoplePickerProps, ITAPeoplePickerState> {
    public pickerElement = null;
    constructor(props) {
        super(props);
        let _peopleList: ITAPersonaProps[] = [];
        people.map((item, index) => {
            _peopleList.push(item);
        });
        if(this.props.defaultItems != undefined)
        {
            _peopleList = this.props.defaultItems;            
        }
        this.state = {
            peopleList: _peopleList,
        };
    }
    public render(): JSX.Element {
        return (
            <CompactPeoplePicker
                onResolveSuggestions={this._onFilterChanged}
                onChange={(a: ITAPersonaProps[]) => { this.props.onChange(a); }}
                //onEmptyInputFocus={ this._returnMostRecentlyUsed }
                getTextFromItem={this._getTextFromItem}
                pickerSuggestionsProps={suggestionProps}
                className={'ms-PeoplePicker'}
                //onRemoveSuggestion={ this._onRemoveSuggestion }
                //onValidateInput={ this._validateInput }
                itemLimit={this.props.itemLimit ? this.props.itemLimit : 20}
                defaultSelectedItems = {this.props.defaultItems != undefined ? this.props.defaultItems : []}
                />
        );
    }
    private _getTextFromItem(persona: ITAPersonaProps): string {
        return persona.primaryText as string;
    }
    @autobind
    private async _onFilterChanged(filterText: string, currentPersonas: ITAPersonaProps[], limitResults?: number) {
        if (filterText) {
            if (Environment.type == EnvironmentType.Local) {
                let filteredPersonas: ITAPersonaProps[] = this.state.peopleList.filter((value) => {
                    return value.primaryText.toLowerCase().indexOf(filterText.toLowerCase()) === 0;
                });
                return filteredPersonas;
            }
            if (Environment.type == EnvironmentType.SharePoint) {
                const searchResults: IPersonaProps[] = await this.getProductionUsers(filterText);
                return searchResults;
            }
        }
        else {
            return [];
        }
    }
    public getProductionUsers = async (seacrhTerm: string) => {
        seacrhTerm = seacrhTerm.charAt(0).toUpperCase() + seacrhTerm.slice(1).toLowerCase();       
        var resultSet = [];
        const results = await this.props.spContext.spHttpClient.get(`${this.props.spContext.pageContext.web.absoluteUrl}/_api/web/siteusers?$filter=startswith(Title,'${seacrhTerm}')`, SPHttpClient.configurations.v1, {
                headers: {
                    'Accept': 'application/json;odata=nometadata',
                    'odata-version': ''
                }
            }).then(async (response) => {
                await response.json().then(async (queryResults: any) => {                    
                    var foundUsersRows: Array<any> = queryResults.value;
                    var _people: ITAPersonaProps[] = [];
                    await foundUsersRows.map((item, index) => {
                        _people.push({
                            primaryText: item.Title, imageUrl: `https://boraakkaya.sharepoint.com/_layouts/15/userphoto.aspx?size=M&username=${item.Email}`,
                            secondaryText: item.Email,
                            tertiaryText: "",
                            optionalText: "",
                            userID: item.Id
                        });
                    });
                    this.setState({ peopleList: _people });
                    resultSet = _people;
                });
            });
        return resultSet;
    }
}
