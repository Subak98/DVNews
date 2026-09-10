import { IPropertyFieldSite } from "@pnp/spfx-property-controls";
import { IList } from "./IObject";

export interface IServiceProvider {
  getNewsPost(
    sites: IPropertyFieldSite[],
    top?: number,
    skip?: number,
  ): Promise<IList[]>;
  getNewsPostWithNoDepartment(
    sites: IPropertyFieldSite[],
    top?: number,
    skip?: number,
  ): Promise<IList[]>;
  getDepartmentFieldOptions(sites: IPropertyFieldSite[]): Promise<any[]>;
  createNewsPost(
    data: any,
    sites: IPropertyFieldSite[],
    sitetitle: string,
  ): Promise<any>;
  ensureDepartmentFieldExists(sites: IPropertyFieldSite[]): Promise<boolean>;
  ensureExpiryDateExists(sites: IPropertyFieldSite[]): Promise<boolean>;
}
