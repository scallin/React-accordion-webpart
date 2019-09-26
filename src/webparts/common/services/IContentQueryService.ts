import { IPersonaProps, ITag } from 'office-ui-fabric-react';
import { IQueryFilterField } from '../../../controls/PropertyPaneQueryFilterPanel/components/QueryFilter/IQueryFilterField';

export interface IContentQueryService {
    getFilterFields: (webUrl: string, listId: string) => Promise<IQueryFilterField[]>;
    getPeoplePickerSuggestions: (webUrl: string, filterText: string, currentPersonas: IPersonaProps[], limitResults?: number) => Promise<IPersonaProps[]>;
    getTaxonomyPickerSuggestions: (webUrl: string, listId: string, field: IQueryFilterField, filterText: string, currentTerms: ITag[]) => Promise<ITag[]>;
    clearCachedFilterFields: () => void;
}