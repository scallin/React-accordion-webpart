import * as React from "react";
import styles from "./ReactAccordion.module.scss";
import { IReactAccordionProps } from "./IReactAccordionProps";
import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions
} from "@microsoft/sp-http";
import {
  PrimaryButton,
  IButtonStyles
} from "office-ui-fabric-react/lib/Button";
import { SearchBox } from "office-ui-fabric-react/lib/SearchBox";
import { Spinner, SpinnerSize } from "office-ui-fabric-react/lib/Spinner";
import { Accordion, AccordionItem } from "react-accessible-accordion";
import { Text } from '@microsoft/sp-core-library';
import "react-accessible-accordion/dist/react-accessible-accordion.css";
import { IReactAccordionState } from "./IReactAccordionState";
import IAccordionListItem from "../models/IAccordionListItem";
import { AccordionWrapper } from "./AccordionWrapper";
import "./accordion.css";
import IAccordionStyles from "../models/IAccordionStyles";
import { IQueryFilter } from "../../../controls/PropertyPaneQueryFilterPanel/components/QueryFilter/IQueryFilter";
import { QueryFilterOperator } from '../../../controls/PropertyPaneQueryFilterPanel/components/QueryFilter/QueryFilterOperator';
import { QueryFilterJoin } from '../../../controls/PropertyPaneQueryFilterPanel/components/QueryFilter/QueryFilterJoin';

export default class ReactAccordion extends React.Component<IReactAccordionProps, IReactAccordionState> {

  constructor(props: IReactAccordionProps, state: IReactAccordionState) {
    super(props);
    this.state = {
      status: this.listNotConfigured(this.props) ? "Please configure list in Web Part properties" : "Ready",
      items: [],
      listItems: [],
      isLoading: false,
      loaderMessage: "",
      listName: this.props.listName,
      filters: this.props.filters,
      activeButtonIndex: 0
    };

    this.readItems = this.readItems.bind(this);

    if (!this.listNotConfigured(this.props)) {
      this.readItems();
    }

    this.searchTextChange = this.searchTextChange.bind(this);
  }

  /***************************************************
   * Using this life cycle method to check if the slider value for max items to fetch is changed
   * And then calling the readItems method to update the state of the component
   * @param nextProps IReactAccordionProps
   */
  public componentWillReceiveProps(nextProps: IReactAccordionProps): void {
    if (this.props.maxItemsToFetchFromTheList !== nextProps.maxItemsToFetchFromTheList || this.props.maxItemsPerPage !== nextProps.maxItemsPerPage) {
      this.readItems(nextProps.maxItemsToFetchFromTheList);
    }
  }

  /***************************************************
   * Checks if the list is not configured
   * @param props IReactAccordionProps
   * @returns boolean
   **************************************************/
  private listNotConfigured(props: IReactAccordionProps): boolean {
    return (
      props.listName === undefined ||
      props.listName === null ||
      props.listName.length === 0
    );
  }

  /***************************************************
   * Triggered for onkeypress event of searchbox
   * @param event
   **************************************************/
  private searchTextChange(event) {
    if (event === undefined || event === null || event === "") {
      let listItemsCollection = [...this.state.listItems];
      this.setState({
        items: listItemsCollection.splice(0, this.props.maxItemsPerPage)
      });
    } else {
      var updatedList = [...this.state.listItems];
      updatedList = updatedList.filter(item => {
        return (
          item.Title.toLowerCase().search(event.toLowerCase()) !== -1 ||
          item.Description.toLowerCase().search(event.toLowerCase()) !== -1
        );
      });
      this.setState({ items: updatedList });
    }
  }

  private static generateFilters(filters:IQueryFilter[]): string {
    return '';
  }

  /************************************************************
   * Fetches the list content from the site contents
   * @param nextLimit : The number used to limit number of items to fetch.
   ************************************************************/
  private readItems(nextLimit?: number): void {
    // Limits the api request to fetch only a specific number of records
    const limit: number = nextLimit === undefined ? this.props.maxItemsToFetchFromTheList : nextLimit;

    // create filter
    let filterRestFormat: string = "({0} {1} '{2}'){3}";  // ex: (Status ne 'Completed')and
    let filterRestFormatContains: string = "(substringof('{0}',{1})){2}"; // ex: (substringof('External',Title))and
    let filterRestFormatStartsWith: string = "(startswith('{0}',{1})){2}";  // ex: (startswith(ContentTypeId,'0x0101'))and
    let filterJoin: string = '';
    let filterRest: string = '';
    if (this.props.filters && this.props.filters.length > 0) {
      let counter : number = 1;
      filterRest = "&$filter=";
      for(let filter of this.props.filters) {
        // filterJoin is whether or not to include the and/or at the end of the filter
        filterJoin = (counter == this.props.filters.length) ? '' : QueryFilterJoin[filter.join].toLowerCase();
        if (filter.operator < 7)
          filterRest += Text.format(filterRestFormat, filter.field.internalName, QueryFilterOperator[filter.operator], filter.value, filterJoin);
        if (filter.operator == 7)
          filterRest += Text.format(filterRestFormatContains, filter.value, filter.field.internalName, filterJoin);
        if (filter.operator == 8)
          filterRest += Text.format(filterRestFormatStartsWith, filter.field.internalName, filter.value, filterJoin);
        counter += 1;
      }
    }
    this.setState({ isLoading: true });
    let restAPI = this.props.siteUrl + `/_api/web/Lists/GetByTitle('${this.props.listName}')/items?$select=Title,Description,SortOrder&$orderby=SortOrder&$top=${limit}${filterRest}`;
    console.log("rest api: " + restAPI);
    this.props.spHttpClient
      .get(restAPI, SPHttpClient.configurations.v1, {
        headers: {
          Accept: "application/json;odata=nometadata", "odata-version": ""
        }
      })
      .then(
        (
          response: SPHttpClientResponse
        ): Promise<{ value: IAccordionListItem[] }> => {
          if (response.status === 200) return response.json();
          else {
            return Promise.reject(
              new Error(`Bad Request - List ${this.props.listName} does not have required columns (Title, Description, SortOrder)`
              )
            );
          }
        }
      )
      .then(
        (response: { value: IAccordionListItem[] }): void => {
          let listItemsCollection = [...response.value];
          this.setState({
            status: "",
            items: listItemsCollection.splice(0, this.props.maxItemsPerPage),
            listItems: response.value,
            isLoading: false,
            loaderMessage: ""
          });
        },
        (error: any): void => {
          this.setState({
            status: "Loading all items failed with error: " + error,
            items: [],
            listItems: [],
            isLoading: false,
            loaderMessage: ""
          });
        }
      );
  }

  /***************************************************
   * Render method for the component
   * @returns React.ReactElement<IReactAccordionProps>
   **************************************************/
  public render(): React.ReactElement<IReactAccordionProps> {
    if (this.props.listName !== this.state.listName) {
      let _listName = this.props.listName;
      this.props.updateListName();
      this.setState({ listName: _listName });
      this.readItems();
    }

    let displayLoader;
    let faqTitle;
    let { listItems } = this.state;
    let pageCountDivisor: number = this.props.maxItemsPerPage;
    let pageCount: number;
    let pageButtons = [];

    let _pagedButtonClick = (pageNumber: number, listData: any) => {
      let btnIndex = pageNumber - 1;

      let startIndex: number = (pageNumber - 1) * pageCountDivisor;
      let listItemsCollection = [...listData];
      this.setState({
        items: listItemsCollection.splice(startIndex, pageCountDivisor),
        activeButtonIndex: btnIndex
      });
    };

    const {
      questionBackgroundColor,
      questionTextColor,
      answerBackgroundColor,
      answerTextColor
    } = this.props;

    const accordionStyles: IAccordionStyles = {
      questionBGColor: questionBackgroundColor,
      questionTextColor: questionTextColor,
      answerBGColor: answerBackgroundColor,
      answerTextColor: answerTextColor
    };

    const items: JSX.Element[] = this.state.items.map(
      (item: IAccordionListItem, i: number): JSX.Element => {
        return (
          <AccordionItem>
            <AccordionWrapper styles={accordionStyles} id={i} item={item} />
          </AccordionItem>
        );
      }
    );

    if (this.state.isLoading) {
      displayLoader = (
        <div className={`ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}`}>
          <div className="ms-Grid-col ms-u-lg12">
            <Spinner size={SpinnerSize.large} label={this.state.loaderMessage} />
          </div>
        </div>
      );
    } else {
      displayLoader = null;
    }

    if (this.state.listItems.length > 0) {
      pageCount = Math.ceil(this.state.listItems.length / pageCountDivisor);
    }

    // dynamic styles for the pagination buttons
    const customBtnStyle: IButtonStyles = {
      root: {
        backgroundColor: this.props.headerBackgroundColor,
        color: this.props.headerTextColor,
        marginRight: "2px"
      }
    };

    const customBtnStyle_active: IButtonStyles = {
      root: {
        backgroundColor: "silver",
        color: "black",
        marginRight: "2px"
      }
    };
    let handleActiveButtonChanges = (index: number): IButtonStyles => {
      if (index !== this.state.activeButtonIndex) {
        return customBtnStyle;
      }
      return customBtnStyle_active;
    };

    for (let i = 0; i < pageCount; i++) {
      if (pageCount > 1)
        pageButtons.push(
          <PrimaryButton styles={handleActiveButtonChanges(i)} onClick={() => {_pagedButtonClick(i + 1, listItems);}}>
            {" "}
            {i + 1}{" "}
          </PrimaryButton>
        );
    }

    const titleStyleNoSearchBox = {
      backgroundColor: this.props.headerBackgroundColor,
      color: this.props.headerTextColor,
      marginBottom: "0px"
    };

    const titleStyle = {
      backgroundColor: this.props.headerBackgroundColor,
      color: this.props.headerTextColor
    };

    let handleTitleStyles = () => {
      return this.props.isSearchAble ? titleStyle : titleStyleNoSearchBox;
    };

    return (
      <div className={styles.reactAccordion}>
        <div className={styles.container}>
          {faqTitle}
          {displayLoader}
          <div
            role="heading"
            className={styles.webpartTitle}
            style={handleTitleStyles()}
          >
            {this.props.title}
          </div>
          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-u-lg12">
              {this.props.isSearchAble ? (
                <SearchBox onChange={this.searchTextChange} />
              ) : null}
            </div>
          </div>
          <div className={`ms-Grid-row`}>
            <div className="ms-Grid-col ms-u-lg12">
              {this.state.status}
              <div >
                <Accordion accordion={false}>{items}</Accordion>
              </div>
            </div>
          </div>
          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-u-lg12">{pageButtons}</div>
          </div>
        </div>
      </div>
    );
  }
}
