import { SPHttpClient } from '@microsoft/sp-http';
import { IPropertyFieldGroupOrPerson } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';


export interface ICustomTravelRequestWebPartProps {
  parentListId: string;
  childListId: string;
  
  parentTitleColumn: string;
  parentTravelIdColumn: string;
  parentTravelerNameColumn: string;
  parentTravelerEmailColumn: string;
  parentTravelTypeColumn: string;
  parentTravelStatusColumn: string;
  parentEstimatedCostColumn: string;
  parentFromDateColumn: string;
  parentToDateColumn: string;

  // Visibility Toggles
  showParentTitle: boolean;


  // Defaults configured in Web Part
  defaultTravelStatus: string;
  defaultTravelType: string;
  fallbackApprovers: IPropertyFieldGroupOrPerson[]; // Configured via People Picker


  childTitleColumn: string;
  childParentIdColumn: string;
  childTravelIdColumn: string;
  childExpenseDateColumn: string;
  childCategoryColumn: string;
  childDescriptionColumn: string;
  childAmountColumn: string;

  showChildDescription: boolean;
  showChildCategory: boolean;
  showChildExpenseDate: boolean;
}

export interface ICustomTravelRequestProps {
  spHttpClient: SPHttpClient;
  siteUrl: string;
  userEmail: string;
  userName: string;
  properties: ICustomTravelRequestWebPartProps;
}
