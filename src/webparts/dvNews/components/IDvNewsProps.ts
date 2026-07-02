import {
  IPropertyFieldGroupOrPerson,
  IPropertyFieldSite,
} from "@pnp/spfx-property-controls";
import { IServiceProvider } from "../../../service/IServiceProvider";

export interface IDvNewsProps {
  title: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  userEmail: string;
  searchSites: IPropertyFieldSite[];
  provider: IServiceProvider;
  adminUsers: IPropertyFieldGroupOrPerson[];
  context: any;
  layout: string;
  sitetitle: string;
}
