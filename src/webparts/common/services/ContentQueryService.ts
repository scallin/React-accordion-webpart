import { IPersonaProps, ITag } from 'office-ui-fabric-react';
import { SPHttpClient } from '@microsoft/sp-http';
import { isEmpty } from '@microsoft/sp-lodash-subset';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Text, Log } from '@microsoft/sp-core-library';
import { IContentQueryService } from './IContentQueryService';
import { IQueryFilterField } from '../../../controls/PropertyPaneQueryFilterPanel/components/QueryFilter/IQueryFilterField';
import { QueryFilterFieldType } from '../../../controls/PropertyPaneQueryFilterPanel/components/QueryFilter/QueryFilterFieldType';
import { ListService } from './ListService';
import { PeoplePickerService } from './PeoplePickerService';
import { TaxonomyService } from './TaxonomyService';

export class ContentQueryService implements IContentQueryService {

    private readonly logSource = "ContentQueryService.ts";

    /**************************************************************************************************
     * The page context and http clients used for performing REST calls
     **************************************************************************************************/
    private context: WebPartContext;
    private spHttpClient: SPHttpClient;

    /**************************************************************************************************
     * The different services used to perform REST calls
     **************************************************************************************************/
    private listService: ListService;
    private peoplePickerService: PeoplePickerService;
    private taxonomyService: TaxonomyService; 

    /**************************************************************************************************
     * Stores the first async calls locally to avoid useless redundant calls
     **************************************************************************************************/
    private filterFields: IQueryFilterField[];

    /**************************************************************************************************
     * Constructor
     * @param context : A IWebPartContext for logging and page context
     * @param spHttpClient : A SPHttpClient for performing SharePoint specific requests
     **************************************************************************************************/
    constructor(context: WebPartContext, spHttpClient: SPHttpClient) {
        Log.verbose(this.logSource, "Initializing a new IContentQueryService instance...", context.serviceScope);

        this.context = context;
        this.spHttpClient = spHttpClient;
        this.listService = new ListService(this.spHttpClient);
        this.peoplePickerService = new PeoplePickerService(this.spHttpClient);
        this.taxonomyService = new TaxonomyService(this.spHttpClient);
    }

    /**************************************************************************************************
     * Gets the available fields out of the specified web/list
     * @param webUrl : The url of the web from which the list comes from
     * @param listName : The id of the list from which the field must be loaded from
     **************************************************************************************************/
    public getFilterFields(webUrl: string, listName: string):Promise<IQueryFilterField[]> {
        Log.verbose(this.logSource, "Loading dropdown options for toolpart property 'Filters'...", this.context.serviceScope);

        // Resolves an empty array if no web or no list has been selected
        if (isEmpty(webUrl) || isEmpty(listName)) {
            return Promise.resolve(new Array<IQueryFilterField>());
        }

        // Resolves the already loaded data if available
        if(this.filterFields) {
            return Promise.resolve(this.filterFields);
        }

        // Otherwise gets the options asynchronously
        return new Promise<IQueryFilterField[]>((resolve, reject) => {
            this.listService.getListFields(webUrl, listName, ['InternalName', 'Title', 'TypeAsString'], 'Title').then((data:any) => {
                let fields:any[] = data.value;
                let options:IQueryFilterField[] = fields.map((field) => { return { 
                    internalName: field.InternalName, 
                    displayName: field.Title,
                    type: this.getFieldTypeFromString(field.TypeAsString)
                }; });
                this.filterFields = options;
                resolve(options);
            })
            .catch((error) => {
                reject(this.getErrorMessage(webUrl, error));
            });
        });
    }

    /**************************************************************************************************
     * Returns the user suggestions based on the user entered picker input
     * @param webUrl : The web url on which to query for users
     * @param filterText : The filter specified by the user in the people picker
     * @param currentPersonas : The IPersonaProps already selected in the people picker
     * @param limitResults : The results limit if any
     **************************************************************************************************/
    public getPeoplePickerSuggestions(webUrl: string, filterText: string, currentPersonas: IPersonaProps[], limitResults?: number):Promise<IPersonaProps[]> {
        Log.verbose(this.logSource, "Getting people picker suggestions for toolpart property 'Filters'...", this.context.serviceScope);

        return new Promise<IPersonaProps[]>((resolve, reject) => {
            this.peoplePickerService.getUserSuggestions(webUrl, filterText, 1, 15, limitResults).then((data) => {
                let users: any[] = JSON.parse(data.value);
                let userSuggestions:IPersonaProps[] = users.map((user) => { return { 
                    primaryText: user.DisplayText,
                    optionalText: user.EntityData.SPUserID || user.EntityData.SPGroupID
                }; });
                resolve(this.removeUserSuggestionsDuplicates(userSuggestions, currentPersonas));
            })
            .catch((error) => {
                reject(error);
            });
        });
    }

    /**************************************************************************************************
     * Returns the taxonomy suggestions based on the user entered picker input
     * @param webUrl : The web url on which to look for the list
     * @param listName : The id of the list on which to look for the taxonomy field
     * @param field : The IQueryFilterField which contains the selected taxonomy field 
     * @param filterText : The filter text entered by the user
     * @param currentTerms : The current terms
     **************************************************************************************************/
    public getTaxonomyPickerSuggestions(webUrl: string, listName: string, field: IQueryFilterField, filterText: string, currentTerms: ITag[]):Promise<ITag[]> {
        Log.verbose(this.logSource, "Getting taxonomy picker suggestions for toolpart property 'Filters'...", this.context.serviceScope);

        return new Promise<ITag[]>((resolve, reject) => {
            this.taxonomyService.getSiteTaxonomyTermsByTermSet(webUrl, listName, field.internalName, this.context.pageContext.web.language).then((data:any) => {
                let termField = Text.format('Term{0}', this.context.pageContext.web.language);
                let terms: any[] = data.value;
                let termSuggestions: ITag[] = terms.map((term:any) => { return { key: term.Id, name: term[termField] }; });
                resolve(this.removeTermSuggestionsDuplicates(termSuggestions, currentTerms));
            })
            .catch((error) => {
                reject(error);
            });
        });
    }

    /**************************************************************************************************
     * Resets the stored filter fields
     **************************************************************************************************/
    public clearCachedFilterFields() {
        Log.verbose(this.logSource, "Clearing cached dropdown options for toolpart property 'Filter'...", this.context.serviceScope);
        this.filterFields = null;
    }

    /**************************************************************************************************
     * Returns an error message based on the specified error object
     * @param error : An error string/object
     **************************************************************************************************/
    private getErrorMessage(webUrl: string, error: any): string {
        let errorMessage:string = error.statusText ? error.statusText : error;
        return errorMessage;
    }
    
    /**************************************************************************************************
     * Returns a field type enum value based on the provided string type
     * @param fieldTypeStr : The field type as a string
     **************************************************************************************************/
    private getFieldTypeFromString(fieldTypeStr: string): QueryFilterFieldType {
        let fieldType:QueryFilterFieldType;

        switch(fieldTypeStr.toLowerCase().trim()) {
            case 'user':
                fieldType = QueryFilterFieldType.User;
                break;
            case 'usermulti':
                fieldType = QueryFilterFieldType.User;
                break;
            case 'datetime':
                fieldType= QueryFilterFieldType.Datetime;
                break;
            case 'lookup':
                fieldType = QueryFilterFieldType.Lookup;
                break;
            case 'url':
                fieldType = QueryFilterFieldType.Url;
                break;
            case 'number':
                fieldType = QueryFilterFieldType.Number;
                break;
            case 'taxonomyfieldtype':
                fieldType = QueryFilterFieldType.Taxonomy;
                break;
            case 'taxonomyfieldtypemulti':
                fieldType = QueryFilterFieldType.Taxonomy;
                break;
            default:
                fieldType = QueryFilterFieldType.Text;
                break;
        }
        return fieldType;
    }

    /**************************************************************************************************
     * Returns the specified users with possible duplicates removed
     * @param users : The user suggestions from which duplicates must be removed
     * @param currentUsers : The current user suggestions that could be duplicates
     **************************************************************************************************/
    private removeUserSuggestionsDuplicates(users: IPersonaProps[], currentUsers: IPersonaProps[]): IPersonaProps[] {
        Log.verbose(this.logSource, "Removing user suggestions duplicates for toolpart property 'Filters'...", this.context.serviceScope);
        let trimmedUsers: IPersonaProps[] = [];

        for(let user of users) {
            let isDuplicate = currentUsers.filter((u) => { return u.optionalText === user.optionalText; }).length > 0;

            if(!isDuplicate) {
                trimmedUsers.push(user);
            }
        }
        return trimmedUsers;
    }


    /**************************************************************************************************
     * Returns the specified users with possible duplicates removed
     * @param users : The user suggestions from which duplicates must be removed
     * @param currentUsers : The current user suggestions that could be duplicates
     **************************************************************************************************/
    private removeTermSuggestionsDuplicates(terms: ITag[], currentTerms: ITag[]): ITag[] {
        Log.verbose(this.logSource, "Removing term suggestions duplicates for toolpart property 'Filters'...", this.context.serviceScope);
        let trimmedTerms: ITag[] = [];

        for(let term of terms) {
            let isDuplicate = currentTerms.filter((t) => { return t.key === term.key; }).length > 0;

            if(!isDuplicate) {
                trimmedTerms.push(term);
            }
        }
        return trimmedTerms;
    }
}