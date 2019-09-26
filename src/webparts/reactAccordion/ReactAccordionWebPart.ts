import * as React from "react";
import * as ReactDom from "react-dom";
import { Version, DisplayMode, Log } from "@microsoft/sp-core-library";
import * as microsoftTeams from '@microsoft/teams-js';
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { 
  IPropertyPaneConfiguration, 
  PropertyPaneButtonType, 
  PropertyPaneButton, 
  PropertyPaneCheckbox, 
  IPropertyPaneDropdownOption, 
  PropertyPaneDropdown, 
  PropertyPaneSlider, 
  PropertyPaneTextField } from "@microsoft/sp-property-pane";
import {
  PropertyFieldColorPicker,
  PropertyFieldColorPickerStyle
} from "@pnp/spfx-property-controls/lib/PropertyFieldColorPicker";
import { update, get } from '@microsoft/sp-lodash-subset';
import { IPersonaProps, ITag } from 'office-ui-fabric-react';
import { PropertyPaneQueryFilterPanel } from '../../controls/PropertyPaneQueryFilterPanel/PropertyPaneQueryFilterPanel';
import { IQueryFilter } from "../../controls/PropertyPaneQueryFilterPanel/components/QueryFilter/IQueryFilter";
import { IQueryFilterField } from "../../controls/PropertyPaneQueryFilterPanel/components/QueryFilter/IQueryFilterField";
import { IContentQueryService } from "../common/services/IContentQueryService";
import * as strings from "ReactAccordionWebPartStrings";
import ReactAccordion from "./components/ReactAccordion";
import { IReactAccordionProps } from "./components/IReactAccordionProps";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { ISPLists } from "./models/ISPList";
import { ContentQueryService } from '../common/services/ContentQueryService';

export interface IReactAccordionWebPartProps {
  headerBackgroundColor: string;
  headerTextColor: string;
  questionBackgroundColor: string;
  questionTextColor: string;
  answerBackgroundColor: string;
  answerTextColor: string;
  listName: string;
  choice: string;
  title: string;
  displayMode: DisplayMode;
  maxItemsPerPage: number;
  maxItemsToFetchFromTheList: number;
  isSearchAble: boolean;
  filters: IQueryFilter[];
}

export default class ReactAccordionWebPart extends BaseClientSideWebPart<IReactAccordionWebPartProps> {
  private readonly logSource = "ReactAccordionWebPart.ts";
  private lists: IPropertyPaneDropdownOption[];
  private _teamsContext: microsoftTeams.Context;
  private filtersPanel: PropertyPaneQueryFilterPanel; // custom toolpart property pane
  private currentListSource: string = '';

  /***************************************************************************
   * Service used to perform REST calls 
   ***************************************************************************/
  private ContentQueryService: IContentQueryService;

  protected onInit(): Promise<any> {
    let retVal: Promise<any> = Promise.resolve();
    if (this.context.microsoftTeams) {
      retVal = new Promise((resolve, reject) => {
        this.context.microsoftTeams.getContext(context => {
          this._teamsContext = context;
          resolve();
        });
      });
    }
    else {
      retVal = new Promise<void>((resolve, reject) => {
        this.ContentQueryService = new ContentQueryService(this.context, this.context.spHttpClient);
        resolve();
      });
    }
    return retVal;
  }

  public render(): void {
    const element: React.ReactElement<IReactAccordionProps> = React.createElement(ReactAccordion, {
      headerBackgroundColor: this.properties.headerBackgroundColor,
      headerTextColor: this.properties.headerTextColor,
      questionBackgroundColor: this.properties.questionBackgroundColor,
      questionTextColor: this.properties.questionTextColor,
      answerBackgroundColor: this.properties.answerBackgroundColor,
      answerTextColor: this.properties.answerTextColor,
      listName: this.properties.listName,
      spHttpClient: this.context.spHttpClient,
      siteUrl: this.context.pageContext.web.absoluteUrl,
      title: this.properties.title,
      displayMode: this.displayMode,
      maxItemsPerPage: this.properties.maxItemsPerPage,
      maxItemsToFetchFromTheList: this.properties.maxItemsToFetchFromTheList,
      isSearchAble: this.properties.isSearchAble,
      filters: this.properties.filters,
      updateListName: () => {
        this.render();
      }
    });
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  /***************************************************
   * Fetches the lists that are on the sharepoint site
   * @returns Promise<ISPLists>
   **************************************************/
  private _getListData(): Promise<ISPLists> {
    let restAPI = this.context.pageContext.web.absoluteUrl + `/_api/web/lists?$filter=Hidden eq false`;
    return this.context.spHttpClient
      .get(restAPI, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  /***************************************************
   * Returns propertypane dropdown options in key and value pairs to be used for
   * property pane dropdown
   * @returns Promise<IPropertyPaneDropdownOption[]>
   **************************************************/
  private _loadSPLists(): Promise<IPropertyPaneDropdownOption[]> {
    return new Promise<IPropertyPaneDropdownOption[]>(
      (
        resolve: (options: IPropertyPaneDropdownOption[]) => void,
        reject: (error: any) => void
      ) => {
        this._getListData().then(data => {
          var list = [];
          data.value.map((item, i) => {
            list.push({ key: item.Title, text: item.Title });
          });
          resolve(list);
        });
      }
    );
  }

  protected onPropertyPaneConfigurationStart(): void {
    if (this.lists) {
      return;
    }
    this.context.statusRenderer.displayLoadingIndicator(this.domElement, "lists");
    this._loadSPLists().then(
      (listOptions: IPropertyPaneDropdownOption[]): void => {
        this.lists = listOptions;
        this.context.propertyPane.refresh();
        this.context.statusRenderer.clearLoadingIndicator(this.domElement);
        this.render();
      }
    );
  }

  /***************************************************
   * This method is called after reset to default button is clicked
   * in properties pane
   **************************************************/
  protected onResetHeaderColorProperty = (): void => {
    this.properties.headerBackgroundColor = "#000047";
    this.properties.headerTextColor = "#ffffff";
  }

  /***************************************************
   * This method is called after reset to default button is clicked
   * in properties pane
   **************************************************/
  protected onResetQuestionColorProperty = (): void => {
    this.properties.questionBackgroundColor = "#ffffff";
    this.properties.questionTextColor = "#000000";
  }

  /***************************************************
   * This method is called after reset to default button is clicked
   * in properties pane
   **************************************************/
  protected onResetAnswerColorProperty = (): void => {
    this.properties.answerBackgroundColor = "#ffffff";
    this.properties.answerTextColor = "#000000";
  }

  /***************************************************************************
   * Loads the dropdown options for the listTitle property
   ***************************************************************************/
  private loadFilterFields():Promise<IQueryFilterField[]> {
    if (this.currentListSource != this.properties.listName) {
      this.ContentQueryService.clearCachedFilterFields();
      this.currentListSource = this.properties.listName;
    }
    return this.ContentQueryService.getFilterFields(this.context.pageContext.web.absoluteUrl, this.properties.listName);
  }

  /***************************************************************************
   * Returns the user suggestions based on the user entered picker input
   * @param filterText : The filter specified by the user in the people picker
   * @param currentPersonas : The IPersonaProps already selected in the people picker
   * @param limitResults : The results limit if any
   ***************************************************************************/
  private loadPeoplePickerSuggestions(filterText: string, currentPersonas: IPersonaProps[], limitResults?: number):Promise<IPersonaProps[]> {
    return this.ContentQueryService.getPeoplePickerSuggestions(this.context.pageContext.web.absoluteUrl, filterText, currentPersonas, limitResults);
  }

  /***************************************************************************
   * Returns the taxonomy suggestions based on the user entered picker input
   * @param field : The taxonomy field from which to load the terms from
   * @param filterText : The filter specified by the user in the people picker
   * @param currentPersonas : The IPersonaProps already selected in the people picker
   * @param limitResults : The results limit if any
   ***************************************************************************/
  private loadTaxonomyPickerSuggestions(field: IQueryFilterField, filterText: string, currentTerms: ITag[]):Promise<ITag[]> {
    return this.ContentQueryService.getTaxonomyPickerSuggestions(this.context.pageContext.web.absoluteUrl, this.properties.listName, field, filterText, currentTerms);
  }
  
  /***************************************************************************
   * When a custom property pane updates
   ***************************************************************************/
  private onCustomPropertyPaneChange(propertyPath: string, newValue: any): void {
    Log.verbose(this.logSource, "WebPart property '" + propertyPath + "' has changed, refreshing WebPart...", this.context.serviceScope);
    
    const oldValue = get(this.properties, propertyPath);
    
    // Stores the new value in web part properties
    update(this.properties, propertyPath, (): any => { return newValue; });
    this.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    if (!this.disableReactivePropertyChanges)
      this.render();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    let filterDisabled: boolean = !this.properties.listName;

    // Creates a custom PropertyPaneQueryFilterPanel for the filters property
    this.filtersPanel = new PropertyPaneQueryFilterPanel("filters", {
      filters: this.properties.filters,
      loadFields: this.loadFilterFields.bind(this),
      onLoadTaxonomyPickerSuggestions: this.loadTaxonomyPickerSuggestions.bind(this),
      onLoadPeoplePickerSuggestions: this.loadPeoplePickerSuggestions.bind(this),
      onPropertyChange: this.onCustomPropertyPaneChange.bind(this),
      trimEmptyFiltersOnChange: true,
      disabled: filterDisabled,
      strings: strings.queryFilterPanelStrings
    });

    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneGeneralDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("title", {
                  label: strings.TitleLabel
                }),
                PropertyPaneDropdown("listName", {
                  label: strings.ListNameLabel,
                  options: this.lists
                }),
                PropertyPaneSlider("maxItemsToFetchFromTheList", {
                  label: strings.MaxItemsToFetchFromTheListLabel,
                  ariaLabel: strings.MaxItemsToFetchFromTheListLabel,
                  min: 3,
                  max: 30,
                  value: 5,
                  showValue: true,
                  step: 1
                }),
                PropertyPaneSlider("maxItemsPerPage", {
                  label: strings.MaxItemsPerPageLabel,
                  ariaLabel: strings.MaxItemsPerPageLabel,
                  min: 3,
                  max: 30,
                  value: 5,
                  showValue: true,
                  step: 1
                }),
                PropertyPaneCheckbox("isSearchAble", {
                  text: strings.isSearchAbleText,
                  checked: true
                }),
                this.filtersPanel
              ]
            }
          ]
        },
        {
          header: {
            description: strings.PropertyPaneHeaderStylesDescription
          },
          groups: [
            {
              groupName: strings.HeaderGroupName,
              groupFields: [
                PropertyFieldColorPicker("headerBackgroundColor", {
                  label: strings.HeaderBackgroundColorLabel,
                  selectedColor: this.properties.headerBackgroundColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  iconName: "Precipitation",
                  key: "headerBackgroundColor"
                }),
                PropertyFieldColorPicker("headerTextColor", {
                  label: strings.HeaderTextColorLabel,
                  selectedColor: this.properties.headerTextColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  iconName: "Precipitation",
                  key: "headerTextColor"
                }),
                PropertyPaneButton("resetBtn", {
                  onClick: this.onResetHeaderColorProperty,
                  text: strings.ResetStyleButtonText,
                  buttonType: PropertyPaneButtonType.Normal
                })
              ]
            }
          ]
        },
        {
          header: {
            description: strings.PropertyPaneQuestionStylesDescription
          },
          groups: [
            {
              groupName: strings.QuestionGroupName,
              groupFields: [
                PropertyFieldColorPicker("questionBackgroundColor", {
                  label: strings.QuestionBackgroundColorLabel,
                  selectedColor: this.properties.questionBackgroundColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  iconName: "Precipitation",
                  key: "questionBackgroundColor"
                }),
                PropertyFieldColorPicker("questionTextColor", {
                  label: strings.QuestionTextColorLabel,
                  selectedColor: this.properties.questionTextColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  iconName: "Precipitation",
                  key: "questionTextColor"
                }),
                PropertyPaneButton("resetBtn", {
                  onClick: this.onResetQuestionColorProperty,
                  text: strings.ResetStyleButtonText,
                  buttonType: PropertyPaneButtonType.Normal
                })
              ]
            }
          ]
        },
        {
          header: {
            description: strings.PropertyPaneAnswerStylesDescription
          },
          groups: [
            {
              groupName: strings.AnswerGroupName,
              groupFields: [
                PropertyFieldColorPicker("answerBackgroundColor", {
                  label: strings.AnswerBackgroundColorLabel,
                  selectedColor: this.properties.answerBackgroundColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  iconName: "Precipitation",
                  key: "answerBackgroundColor"
                }),
                PropertyFieldColorPicker("answerTextColor", {
                  label: strings.AnswerTextColorLabel,
                  selectedColor: this.properties.answerTextColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  iconName: "Precipitation",
                  key: "answerTextColor"
                }),
                PropertyPaneButton("resetBtn", {
                  onClick: this.onResetAnswerColorProperty,
                  text: strings.ResetStyleButtonText,
                  buttonType: PropertyPaneButtonType.Normal
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
