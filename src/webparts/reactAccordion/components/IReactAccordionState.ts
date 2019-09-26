import IAccordionListItem from "../models/IAccordionListItem";
import { PrimaryButton } from "office-ui-fabric-react/lib/Button";
import { IQueryFilter } from "../../../controls/PropertyPaneQueryFilterPanel/components/QueryFilter/IQueryFilter";

export interface IReactAccordionState {
  status: string;
  items: IAccordionListItem[];
  listItems: IAccordionListItem[];
  isLoading: boolean;
  loaderMessage: string;
  listName: string;
  filters: IQueryFilter[];
  activeButtonIndex: number;
}
