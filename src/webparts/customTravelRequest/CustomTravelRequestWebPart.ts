import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  PropertyPaneLabel,
  PropertyPaneToggle,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import CustomTravelRequest from './components/CustomTravelRequest';
import { ICustomTravelRequestWebPartProps } from './components/ICustomTravelRequestProps';
import { PropertyFieldPeoplePicker, PrincipalType } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';


export default class CustomTravelRequestWebPart extends BaseClientSideWebPart<ICustomTravelRequestWebPartProps> {
  private _lists: IPropertyPaneDropdownOption[] = [];
  
  // Mapped Parent Columns (Categorized)
  private _parentTextFields: IPropertyPaneDropdownOption[] = [];
  private _parentUserFields: IPropertyPaneDropdownOption[] = [];
  private _parentChoiceFields: IPropertyPaneDropdownOption[] = [];
  private _parentNumberFields: IPropertyPaneDropdownOption[] = [];
  private _parentDateFields: IPropertyPaneDropdownOption[] = [];

  // Mapped Child Columns (Categorized)
  private _childTextFields: IPropertyPaneDropdownOption[] = [];
  private _childLookupFields: IPropertyPaneDropdownOption[] = [];
  private _childChoiceFields: IPropertyPaneDropdownOption[] = [];
  private _childNumberFields: IPropertyPaneDropdownOption[] = [];
  private _childDateFields: IPropertyPaneDropdownOption[] = [];

  // Choices cache for default value selections
  private _parentFieldChoices: { [fieldName: string]: string[] } = {};

  private _loadingLists: boolean = false;
  private _loadingParentFields: boolean = false;
  private _loadingChildFields: boolean = false;

  protected onInit(): Promise<void> {
    return super.onInit().then(() => {
      // Safely ensure fallbackApprovers is an array (handling string or empty value upgrade paths)
      if (typeof this.properties.fallbackApprovers === 'string') {
        const fallbackStr = this.properties.fallbackApprovers as string;
        this.properties.fallbackApprovers = fallbackStr.split(',')
          .map(email => email.trim())
          .filter(email => email.length > 0)
          .map(email => ({
            email: email,
            login: email,
            fullName: email
          }));
      } else if (!this.properties.fallbackApprovers || !Array.isArray(this.properties.fallbackApprovers)) {
        this.properties.fallbackApprovers = [];
      }

      // Initialize list and field options on startup if properties are already configured
      if (this.properties.parentListId) {
        this._fetchFields(this.properties.parentListId, true);
      }
      if (this.properties.childListId) {
        this._fetchFields(this.properties.childListId, false);
      }
    });
  }

  public render(): void {
    const element: React.ReactElement<any> = React.createElement(
      CustomTravelRequest,
      {
        spHttpClient: this.context.spHttpClient,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        userEmail: this.context.pageContext.user.email,
        userName: this.context.pageContext.user.displayName,
        properties: this.properties
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneConfigurationStart(): void {
    this._fetchLists();
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    if (propertyPath === 'parentListId') {
      // Reset dependent column configurations
      this.properties.parentTitleColumn = 'Title';
      this.properties.parentTravelIdColumn = '';
      this.properties.parentTravelerNameColumn = '';
      this.properties.parentTravelerEmailColumn = '';
      this.properties.parentTravelTypeColumn = '';
      this.properties.parentTravelStatusColumn = '';
      this.properties.parentEstimatedCostColumn = '';
      this.properties.parentFromDateColumn = '';
      this.properties.parentToDateColumn = '';
      this.properties.defaultTravelStatus = '';
      this.properties.defaultTravelType = '';
      
      this._parentTextFields = [];
      this._parentUserFields = [];
      this._parentChoiceFields = [];
      this._parentNumberFields = [];
      this._parentDateFields = [];
      this._parentFieldChoices = {};

      this._fetchFields(newValue, true);
    }

    if (propertyPath === 'parentTravelStatusColumn') {
      this.properties.defaultTravelStatus = '';
    }

    if (propertyPath === 'parentTravelTypeColumn') {
      this.properties.defaultTravelType = '';
    }

    if (propertyPath === 'childListId') {
      // Reset dependent column configurations
      this.properties.childTitleColumn = 'Title';
      this.properties.childParentIdColumn = '';
      this.properties.childTravelIdColumn = '';
      this.properties.childExpenseDateColumn = '';
      this.properties.childCategoryColumn = '';
      this.properties.childDescriptionColumn = '';
      this.properties.childAmountColumn = '';

      this._childTextFields = [];
      this._childLookupFields = [];
      this._childChoiceFields = [];
      this._childNumberFields = [];
      this._childDateFields = [];

      this._fetchFields(newValue, false);
    }
  }

  // Fetch Lists in the Current Site
  private _fetchLists(): void {
    if (this._lists.length > 0 || this._loadingLists) {
      return;
    }

    this._loadingLists = true;
    const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$filter=Hidden eq false and BaseType eq 0`;

    this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => response.json())
      .then((data: { value: any[] }) => {
        this._lists = data.value.map(list => ({
          key: list.Id,
          text: list.Title
        }));
        this._loadingLists = false;
        this.context.propertyPane.refresh();
      })
      .catch(error => {
        console.error("Error fetching SharePoint lists:", error);
        this._loadingLists = false;
      });
  }

  // Fetch and Classify Fields of Selected List
  private _fetchFields(listId: string, isParent: boolean): void {
    if (!listId) return;

    if (isParent) {
      this._loadingParentFields = true;
    } else {
      this._loadingChildFields = true;
    }
    this.context.propertyPane.refresh();

    const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${listId}')/fields?$filter=Hidden eq false and ReadOnlyField eq false`;

    this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => response.json())
      .then((data: { value: any[] }) => {
        const textOptions: IPropertyPaneDropdownOption[] = [];
        const userOptions: IPropertyPaneDropdownOption[] = [];
        const choiceOptions: IPropertyPaneDropdownOption[] = [];
        const numberOptions: IPropertyPaneDropdownOption[] = [];
        const dateOptions: IPropertyPaneDropdownOption[] = [];
        const lookupOptions: IPropertyPaneDropdownOption[] = [];

        // Include Title column by default if not read-only
        data.value.forEach(field => {
          const option = { key: field.InternalName, text: `${field.Title} (${field.InternalName})` };
          
          if (field.TypeAsString === 'Text' || field.TypeAsString === 'Note') {
            textOptions.push(option);
          } else if (field.TypeAsString === 'User') {
            userOptions.push(option);
          } else if (field.TypeAsString === 'Choice') {
            choiceOptions.push(option);
            if (isParent) {
              this._parentFieldChoices[field.InternalName] = field.Choices || [];
            }
          } else if (field.TypeAsString === 'Number' || field.TypeAsString === 'Currency') {
            numberOptions.push(option);
          } else if (field.TypeAsString === 'DateTime') {
            dateOptions.push(option);
          } else if (field.TypeAsString === 'Lookup') {
            lookupOptions.push(option);
          }
        });

        // Always ensure Title is in textOptions (sometimes it is set to FromBaseType=TRUE and handled differently)
        if (!textOptions.some(o => o.key === 'Title')) {
          textOptions.unshift({ key: 'Title', text: 'Title (Title)' });
        }

        if (isParent) {
          this._parentTextFields = textOptions;
          this._parentUserFields = userOptions;
          this._parentChoiceFields = choiceOptions;
          this._parentNumberFields = numberOptions;
          this._parentDateFields = dateOptions;
          this._loadingParentFields = false;
        } else {
          this._childTextFields = textOptions;
          this._childLookupFields = lookupOptions;
          this._childChoiceFields = choiceOptions;
          this._childNumberFields = numberOptions;
          this._childDateFields = dateOptions;
          this._loadingChildFields = false;
        }

        this.context.propertyPane.refresh();
      })
      .catch(error => {
        console.error(`Error fetching fields for list ${listId}:`, error);
        if (isParent) {
          this._loadingParentFields = false;
        } else {
          this._loadingChildFields = false;
        }
        this.context.propertyPane.refresh();
      });
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "Map your Travel Requests Parent List and Details Child List to construct a dynamic, single-form submission layout."
          },
          groups: [
            {
              groupName: "1. Parent Travel Requests List Mapping",
              groupFields: [
                PropertyPaneDropdown('parentListId', {
                  label: "Select Travel Requests List (Parent)",
                  options: this._lists,
                  disabled: this._loadingLists
                }),
                this._loadingParentFields ? PropertyPaneLabel('parentListId', { text: "Loading columns..." }) : null,
                
                // Show mappings if list is selected
                this.properties.parentListId ? PropertyPaneDropdown('parentTitleColumn', {
                  label: "Title (Purpose of Travel) Column",
                  options: this._parentTextFields
                }) : null,
                this.properties.parentListId ? PropertyPaneDropdown('parentTravelIdColumn', {
                  label: "Travel ID Column",
                  options: this._parentTextFields
                }) : null,
                this.properties.parentListId ? PropertyPaneDropdown('parentTravelerNameColumn', {
                  label: "Traveler Column (Person)",
                  options: this._parentUserFields
                }) : null,
                this.properties.parentListId ? PropertyPaneDropdown('parentTravelerEmailColumn', {
                  label: "Traveler Email Column (Text)",
                  options: this._parentTextFields
                }) : null,
                this.properties.parentListId ? PropertyPaneDropdown('parentTravelTypeColumn', {
                  label: "Travel Type Column (Choice)",
                  options: this._parentChoiceFields
                }) : null,
                this.properties.parentListId ? PropertyPaneDropdown('parentTravelStatusColumn', {
                  label: "Travel Status Column (Choice)",
                  options: this._parentChoiceFields
                }) : null,
                this.properties.parentListId ? PropertyPaneDropdown('parentEstimatedCostColumn', {
                  label: "Estimated Cost Column (Currency/Number)",
                  options: this._parentNumberFields
                }) : null,
                this.properties.parentListId ? PropertyPaneDropdown('parentFromDateColumn', {
                  label: "From Date Column (DateTime)",
                  options: this._parentDateFields
                }) : null,
                this.properties.parentListId ? PropertyPaneDropdown('parentToDateColumn', {
                  label: "To Date Column (DateTime)",
                  options: this._parentDateFields
                }) : null,
                this.properties.parentListId ? PropertyFieldPeoplePicker('fallbackApprovers', {
                  label: 'Fallback Approvers',
                  initialData: this.properties.fallbackApprovers,
                  allowDuplicate: false,
                  principalType: [PrincipalType.Users],
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  context: this.context as any,
                  properties: this.properties,
                  onGetErrorMessage: undefined,
                  deferredValidationTime: 0,
                  searchTextLimit: 3,
                  key: 'fallbackApproversFieldId'
                }) : null,
                this.properties.parentListId ? PropertyPaneToggle('showParentTitle', {
                  label: "Show Purpose Field on Submission Form",
                  checked: this.properties.showParentTitle
                }) : null,

                this.properties.parentListId && this.properties.parentTravelTypeColumn ? PropertyPaneDropdown('defaultTravelType', {
                  label: "Default Travel Type Value",
                  options: (this._parentFieldChoices[this.properties.parentTravelTypeColumn] || []).map(c => ({ key: c, text: c }))
                }) : null,
                this.properties.parentListId && this.properties.parentTravelStatusColumn ? PropertyPaneDropdown('defaultTravelStatus', {
                  label: "Default Travel Status Value",
                  options: (this._parentFieldChoices[this.properties.parentTravelStatusColumn] || []).map(c => ({ key: c, text: c }))
                }) : null
              ].filter(field => field !== null)
            },
            {
              groupName: "2. Child Travel Details Mapping",
              groupFields: [
                PropertyPaneDropdown('childListId', {
                  label: "Select Details List (Child)",
                  options: this._lists,
                  disabled: this._loadingLists
                }),
                this._loadingChildFields ? PropertyPaneLabel('childListId', { text: "Loading columns..." }) : null,
                
                // Show mappings if list is selected
                this.properties.childListId ? PropertyPaneDropdown('childTitleColumn', {
                  label: "Title Column",
                  options: this._childTextFields
                }) : null,
                this.properties.childListId ? PropertyPaneDropdown('childParentIdColumn', {
                  label: "Parent ID Lookup Column (Lookup)",
                  options: this._childLookupFields
                }) : null,
                this.properties.childListId ? PropertyPaneDropdown('childTravelIdColumn', {
                  label: "Travel ID Column",
                  options: this._childTextFields
                }) : null,
                this.properties.childListId ? PropertyPaneDropdown('childExpenseDateColumn', {
                  label: "Expense Date Column (DateTime)",
                  options: this._childDateFields
                }) : null,
                this.properties.childListId ? PropertyPaneDropdown('childCategoryColumn', {
                  label: "Category Column (Choice)",
                  options: this._childChoiceFields
                }) : null,
                this.properties.childListId ? PropertyPaneDropdown('childDescriptionColumn', {
                  label: "Description Column (Note/Text)",
                  options: this._childTextFields
                }) : null,
                this.properties.childListId ? PropertyPaneDropdown('childAmountColumn', {
                  label: "Amount Column (Currency/Number)",
                  options: this._childNumberFields
                }) : null,
                this.properties.childListId ? PropertyPaneToggle('showChildDescription', {
                  label: "Show Description Field on Submission Grid",
                  checked: this.properties.showChildDescription
                }) : null,
                this.properties.childListId ? PropertyPaneToggle('showChildCategory', {
                  label: "Show Category Field on Submission Grid",
                  checked: this.properties.showChildCategory
                }) : null,
                this.properties.childListId ? PropertyPaneToggle('showChildExpenseDate', {
                  label: "Show Expense Date Field on Submission Grid",
                  checked: this.properties.showChildExpenseDate
                }) : null
              ].filter(field => field !== null)
            }
          ]
        }
      ]
    };
  }
}
