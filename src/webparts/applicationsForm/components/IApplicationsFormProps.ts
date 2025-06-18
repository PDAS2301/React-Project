import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IApplicationsListItem } from './IApplicationsListItem';

export interface IApplicationsFormProps {
  description: string;
  context: WebPartContext;
}

export interface IApplicationsFormState {
  items: IApplicationsListItem[];
  applicationName: string;
  applicationType: string;
  contact: string;
  errors: {
    applicationName: string;
    contact: string;
  };
}
// export interface IApplicationsFormProps {
//   description: string;
//   isDarkTheme: boolean;
//   environmentMessage: string;
//   hasTeamsContext: boolean;
//   userDisplayName: string;
// }
