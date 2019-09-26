import { SPHttpClient } from "@microsoft/sp-http";
import { DisplayMode } from "@microsoft/sp-core-library";
import { IQueryFilter } from "../../../controls/PropertyPaneQueryFilterPanel/components/QueryFilter/IQueryFilter";

export interface IReactAccordionProps {
  headerBackgroundColor: string;
  headerTextColor: string;
  questionBackgroundColor: string;
  questionTextColor: string;
  answerBackgroundColor: string;
  answerTextColor: string;
  listName: string;
  spHttpClient: SPHttpClient;
  siteUrl: string;
  title: string;
  displayMode: DisplayMode;
  maxItemsPerPage: number;
  maxItemsToFetchFromTheList: number;
  isSearchAble: boolean;
  filters: IQueryFilter[];
  updateListName: () => void;
}
