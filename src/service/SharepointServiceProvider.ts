import { IServiceProvider } from "./IServiceProvider";

import { BaseWebPartContext, WebPartContext } from "@microsoft/sp-webpart-base";

import { spfi, SPFI, SPFx } from "@pnp/sp";
import { graphfi, SPFx as graphSPFx } from "@pnp/graph";
import "@pnp/graph/groups";
import "@pnp/graph/members";

import { getSP } from "./pnpjsConfig";
import { IFeaturedNewsList, IList } from "./IObject";
import { SPHttpClient } from "@microsoft/sp-http";
import moment from "moment";
import { PromotedState } from "@pnp/sp/clientside-pages";
import { IPropertyFieldSite } from "@pnp/spfx-property-controls";
import {
  CalendarType,
  ChoiceFieldFormatType,
  DateTimeFieldFormatType,
  DateTimeFieldFriendlyFormatType,
} from "@pnp/sp/fields";

export default class SharepointServiceProvider implements IServiceProvider {
  _webPartContext: BaseWebPartContext;
  _webAbsoluteUrl: string;
  public sp: SPFI;
  private _cachedGraphGroups: any[] | null = null;
  //   private _webPartContext: BaseWebPartContext;
  //   private _webAbsoluteUrl: string;

  constructor(_context: BaseWebPartContext) {
    this._webPartContext = _context;
    this._webAbsoluteUrl = _context.pageContext.web.absoluteUrl;
    this.sp = getSP(_context as WebPartContext);
  }
  public async getDepartmentFieldOptions(
    sites: IPropertyFieldSite[],
  ): Promise<any[]> {
    const siteUrl = `${sites[0].url}`;
    const subweb = spfi(siteUrl).using(SPFx(this._webPartContext));
    const field = await subweb.web.lists
      .getByTitle("Site Pages")
      .fields.getByTitle("Department")();

    const choices = field.Choices || [];
    return [...choices];
  }

  public async ensureDepartmentFieldExists(
    sites: IPropertyFieldSite[],
  ): Promise<boolean> {
    try {
      const siteUrl = `${sites?.[0].url}`;
      const subweb = spfi(siteUrl).using(SPFx(this._webPartContext));
      const list = subweb.web.lists.getByTitle("Site Pages");

      // Try to get the Department field
      try {
        const field = await list.fields.getByTitle("Department")();
        // Field exists
        return true;
      } catch (error: any) {
        console.log("not exist");

        await list.fields.addMultiChoice("Department", {
          Choices: ["All"],
          FillInChoice: false,
          Group: "My Group",
        });
        await list.defaultView.fields.add("Department");
        return true;
      }
    } catch (error) {
      console.error("Error ensuring Department field exists:", error);
      return false;
    }
  }
  public async ensureExpiryDateExists(
    sites: IPropertyFieldSite[],
  ): Promise<boolean> {
    try {
      const siteUrl = `${sites?.[0].url}`;
      const subweb = spfi(siteUrl).using(SPFx(this._webPartContext));
      const list = subweb.web.lists.getByTitle("Site Pages");

      // Try to get the ExpiryDate field
      try {
        const field = await list.fields.getByTitle("ExpiryDate")();
        // Field exists
        return true;
      } catch (error: any) {
        console.log("not exist");

        await list.fields.addDateTime("ExpiryDate", {
          DisplayFormat: DateTimeFieldFormatType.DateOnly,
          DateTimeCalendarType: CalendarType.Gregorian,
          FriendlyDisplayFormat: DateTimeFieldFriendlyFormatType.Disabled,
          Group: "My Group",
        });
        await list.fields.addChoice(`Status`, {
          Choices: ["Active", "Expired"],
          EditFormat: ChoiceFieldFormatType.Dropdown,
          FillInChoice: false,
          Group: "My Group",
        });
        await list.defaultView.fields.add("ExpiryDate");
        await list.defaultView.fields.add("Status");
        return true;
      }
    } catch (error) {
      console.error("Error ensuring ExpiryDate field exists:", error);
      return false;
    }
  }

  // public async getNewsPost(
  //   sites: IPropertyFieldSite[],
  //   top: number = 200,
  //   skip: number = 0,
  // ): Promise<IFeaturedNewsList[]> {
  //   // console.log(sites);
  //   const siteUrl = `${sites[0].url}`;
  //   const subweb = spfi(siteUrl).using(SPFx(this._webPartContext));

  //   const filterNews =
  //     "PromotedState eq 2 and (Department ne null and Department ne '')";
  //   const list = subweb.web.lists.getByTitle("Site Pages");
  //   const getitems = await list.items
  //     .filter(filterNews)
  //     .select(
  //       "*",
  //       "Title",
  //       "Department",
  //       "ShowNewsinHome",
  //       "Modified",
  //       "BannerImageUrl",
  //       "Description",
  //       "FileRef",
  //       "FileDirRef",
  //       "Author/ID",
  //       "Author/Title",
  //       "Author/EMail",
  //     )
  //     .expand("Author")
  //     .orderBy("Modified", false)
  //     .skip(skip)
  //     .top(top)();
  //   console.log("getitems", getitems);
  //   // Sort by Modified date (newest first)
  //   const orderedNews = getitems.sort((a: any, b: any) => {
  //     return new Date(b.Modified).getTime() - new Date(a.Modified).getTime();
  //   });

  //   const bannerUrl = `https://media.akamai.odsp.cdn.office.net/southindia1-mediap.svc.ms/transform/thumbnail?provider=url&inputFormat=png&docid=https://media.akamai.odsp.cdn.office.net/${window.location.host}/_layouts/15/images/sitepagethumbnail.png&w=200`;

  //   const graph = graphfi().using(graphSPFx(this._webPartContext as any));
  //   const currentUserEmail =
  //     this._webPartContext.pageContext.user.email?.toLowerCase();

  //   const groups: any =
  //     this._cachedGraphGroups ||
  //     (await graph.groups.select("displayName", "id")());
  //   this._cachedGraphGroups = groups;

  //   const _items: IFeaturedNewsList[] = [];
  //   const filteredItems: IFeaturedNewsList[] = [];
  //   const groupAccessCache = new Map<string, boolean>();

  //   for (let i = 0; i < orderedNews.length; i++) {
  //     const item = orderedNews[i];
  //     const siteUrl = item._originSiteUrl;
  //     const departmentArray = Array.isArray(item.Department)
  //       ? item.Department
  //       : item.Department && typeof item.Department === "string"
  //         ? (item.Department as string)
  //             .split(",")
  //             .map((dept: string) => dept.trim())
  //             .filter(Boolean)
  //         : [];

  //     const uniqueDepartments = Array.from(new Set(departmentArray));
  //     const isDepartmentEmpty = uniqueDepartments.length === 0;
  //     let userHasAccess = isDepartmentEmpty;

  //     if (!isDepartmentEmpty) {
  //       for (const departmentName of uniqueDepartments) {
  //         const targetGroup: any = groups.find(
  //           (g: any) => g.displayName === departmentName,
  //         );

  //         if (!targetGroup) {
  //           continue;
  //         }

  //         if (groupAccessCache.has(targetGroup.id)) {
  //           if (groupAccessCache.get(targetGroup.id)) {
  //             userHasAccess = true;
  //             break;
  //           }
  //           continue;
  //         }

  //         const [groupMembers, groupOwners] = await Promise.all([
  //           graph.groups.getById(targetGroup.id).members(),
  //           graph.groups.getById(targetGroup.id).owners(),
  //         ]);

  //         const isMember = groupMembers.some(
  //           (member: any) =>
  //             member.mail?.toLowerCase() === currentUserEmail ||
  //             member.userPrincipalName?.toLowerCase() === currentUserEmail,
  //         );
  //         const isOwner = groupOwners.some(
  //           (owner: any) =>
  //             owner.mail?.toLowerCase() === currentUserEmail ||
  //             owner.userPrincipalName?.toLowerCase() === currentUserEmail,
  //         );

  //         const hasAccess = isMember || isOwner;
  //         groupAccessCache.set(targetGroup.id, hasAccess);

  //         if (hasAccess) {
  //           userHasAccess = true;
  //           break;
  //         }
  //       }
  //     }

  //     const linkUrl = item.FileRef;

  //     const imageJSON = item.Image && JSON.parse(item.Image);
  //     const relativeImage = imageJSON?.serverRelativeUrl;

  //     const customUrl =
  //       imageJSON &&
  //       siteUrl +
  //         "/Lists/SiteAssets/Attachments/" +
  //         item.Id +
  //         "/" +
  //         imageJSON?.fileName;

  //     const customPhoto = decodeURI(relativeImage) ?? customUrl;

  //     const lst: IFeaturedNewsList = {
  //       Title: item.Title || item.FileRef?.split("/SitePages/")[1],
  //       date: moment(item.Created).format("MMMM DD, yyyy"),
  //       Created: moment(item.Created).format(),
  //       Author: item.Author,
  //       Department: item.Department || "",
  //       ShowNewsinHome: item.ShowNewsinHome,
  //       AuthorByLine: item.Author ? item.Author : "",
  //       Image:
  //         customPhoto &&
  //         customPhoto !== undefined &&
  //         customPhoto !== "undefined"
  //           ? customPhoto
  //           : (item.BannerImageUrl?.Url ?? bannerUrl),
  //       Description: item.Description,
  //       Link: linkUrl || "",
  //     };

  //     _items.push(lst);
  //     if (userHasAccess) {
  //       filteredItems.push(lst);
  //     }
  //   }
  //   console.log(filteredItems);

  //   return filteredItems.length > 0 ? filteredItems : _items;
  // }
  public async getNewsPost(
    sites: IPropertyFieldSite[],
    top: number = 200,
    skip: number = 0,
  ): Promise<IFeaturedNewsList[]> {
    const siteUrl = `${sites[0].url}`;
    const subweb = spfi(siteUrl).using(SPFx(this._webPartContext));

    const today = moment().format("YYYY-MM-DD");
    const filterNews = `PromotedState eq 2 and (ExpiryDate ge datetime'${today}T23:59:59Z' or ExpiryDate eq null)`;
    const list = subweb.web.lists.getByTitle("Site Pages");
    const siteTitle = await this.sp.web.select("Title")();
    const getitems = await list.items
      .filter(filterNews)
      .select(
        "*",
        "Title",
        "Department",

        "Modified",
        "BannerImageUrl",
        "Description",
        "FileRef",
        "FileDirRef",
        "Author/ID",
        "Author/Title",
        "Author/EMail",
      )
      .expand("Author")
      .orderBy("Modified", false)
      .skip(skip)
      .top(top)();
    console.log(getitems);

    const normalizedSiteTitle = (siteTitle?.Title || "").trim().toLowerCase();
    console.log(normalizedSiteTitle);

    const filteredItems = getitems.filter((item: any) => {
      const departmentValues = Array.isArray(item.Department)
        ? item.Department
        : typeof item.Department === "string"
          ? item.Department.split(",").map((dept: string) => dept.trim())
          : [];

      const normalizedDepartments = departmentValues
        .filter(Boolean)
        .map((dept: string) => dept.toLowerCase());

      return (
        normalizedDepartments.includes("all") ||
        normalizedDepartments.includes(normalizedSiteTitle)
      );
    });

    console.log(filteredItems);

    return filteredItems.map((item: any) => {
      const itemSiteUrl = item._originSiteUrl;
      const imageJSON = item.Image && JSON.parse(item.Image);
      const relativeImage = imageJSON?.serverRelativeUrl;
      const customUrl =
        imageJSON &&
        itemSiteUrl +
          "/Lists/SiteAssets/Attachments/" +
          item.Id +
          "/" +
          imageJSON?.fileName;
      const customPhoto = decodeURI(relativeImage) ?? customUrl;
      const bannerUrl = `https://media.akamai.odsp.cdn.office.net/southindia1-mediap.svc.ms/transform/thumbnail?provider=url&inputFormat=png&docid=https://media.akamai.odsp.cdn.office.net/${window.location.host}/_layouts/15/images/sitepagethumbnail.png&w=200`;

      const lst: IFeaturedNewsList = {
        Title: item.Title || item.FileRef?.split("/SitePages/")[1],
        date: moment(item.Created).format("MMMM DD, yyyy"),
        Created: moment(item.Created).format(),
        Author: item.Author,
        Department: item.Department || "",
        // ShowNewsinHome: item.ShowNewsinHome,
        AuthorByLine: item.Author ? item.Author : "",
        Image:
          customPhoto &&
          customPhoto !== undefined &&
          customPhoto !== "undefined"
            ? customPhoto
            : (item.BannerImageUrl?.Url ?? bannerUrl),
        Description: item.Description,
        Link: item.FileRef || "",
      };

      return lst;
    });
  }

  public async getNewsPostWithNoDepartment(
    sites: IPropertyFieldSite[],
    top: number = 200,
    skip: number = 0,
  ): Promise<IFeaturedNewsList[]> {
    const siteUrl = `${sites[0].url}`;
    const subweb = spfi(siteUrl).using(SPFx(this._webPartContext));
    const today = moment().format("YYYY-MM-DD");
    const filterNews = `PromotedState eq 2 and (Department eq 'All') and (ExpiryDate ge datetime'${today}T23:59:59Z' or ExpiryDate eq null)`;
    const list = subweb.web.lists.getByTitle("Site Pages");
    const getitems = await list.items
      .filter(filterNews)
      .select(
        "*",
        "Title",
        "Department",

        "Modified",
        "BannerImageUrl",
        "Description",
        "FileRef",
        "FileDirRef",
        "Author/ID",
        "Author/Title",
        "Author/EMail",
      )
      .expand("Author")
      .orderBy("Modified", false)
      .skip(skip)
      .top(top)();

    return getitems.map((item: any) => {
      const itemSiteUrl = item._originSiteUrl;
      const imageJSON = item.Image && JSON.parse(item.Image);
      const relativeImage = imageJSON?.serverRelativeUrl;
      const customUrl =
        imageJSON &&
        itemSiteUrl +
          "/Lists/SiteAssets/Attachments/" +
          item.Id +
          "/" +
          imageJSON?.fileName;
      const customPhoto = decodeURI(relativeImage) ?? customUrl;
      const bannerUrl = `https://media.akamai.odsp.cdn.office.net/southindia1-mediap.svc.ms/transform/thumbnail?provider=url&inputFormat=png&docid=https://media.akamai.odsp.cdn.office.net/${window.location.host}/_layouts/15/images/sitepagethumbnail.png&w=200`;

      const lst: IFeaturedNewsList = {
        Title: item.Title || item.FileRef?.split("/SitePages/")[1],
        date: moment(item.Created).format("MMMM DD, yyyy"),
        Created: moment(item.Created).format(),
        Author: item.Author,
        Department: item.Department || "",
        // ShowNewsinHome: item.ShowNewsinHome,
        AuthorByLine: item.Author ? item.Author : "",
        Image:
          customPhoto &&
          customPhoto !== undefined &&
          customPhoto !== "undefined"
            ? customPhoto
            : (item.BannerImageUrl?.Url ?? bannerUrl),
        Description: item.Description,
        Link: item.FileRef || "",
      };

      return lst;
    });
  }

  public async createNewsPost(
    data: any,
    sites: IPropertyFieldSite[],
    sitetitle: string,
  ): Promise<any> {
    console.log(data);
    const siteUrl = `${sites[0].url}`;
    const subweb = spfi(siteUrl).using(SPFx(this._webPartContext));
    const list = subweb.web.lists.getByTitle("Site Pages");

    const page3 = await subweb.web.addClientsidePage(
      data.title,
      data.title,
      "Article",
      PromotedState.PromoteOnPublish,
    );
    const graph = graphfi().using(graphSPFx(this._webPartContext));
    const groups = await graph.groups.top(999).select("displayName", "id")();

    // you must publish the new page, after which the page will immediately be promoted to a news article
    // await page3.save();
    const items: any[] = await subweb.web.lists
      .getByTitle("Site Pages")
      .items.orderBy("ID", false)
      .select("Id", "Title", "FileRef")
      .top(1)();
    console.log(items);

    // const item = await subweb.web
    //   .getFileByServerRelativePath(
    //     `/sites/${sitetitle}/SitePages/${data.title.replace(/\s/g, "-")}.aspx`,
    //   )
    //   .getItem();
    // const gitem = await item.select("Id", "Title")();

    const updateitem = await list.items.getById(items[0].Id).update({
      Department: data.departments,
      ExpiryDate: data.expiryDate,
      Status: "Active",
      // ShowNewsinHome: data.showInHome,
    });
    console.log(data.departments);
    if (
      data.departments.length > 0 &&
      data.departments.includes("All") === false
    ) {
      try {
        const context = await this._webPartContext.spHttpClient.post(
          `${siteUrl}/_api/web/lists/getByTitle('Site Pages')/Items/getById(${items[0].Id})/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept: "application/json;odata=nometadata",
              "Content-Type": "application/json;odata=nometadata",
              "odata-version": "",
            },
          },
        );

        if (!context.ok) {
          console.error(
            "Failed to break role inheritance:",
            await context.text(),
          );
        } else {
          // console.log("Role inheritance broken successfully");
        }
      } catch (error) {
        console.error("Error breaking role inheritance:", error);
      }
      const sitePagelist = subweb.web.lists.getByTitle("Site Pages");

      // we can use this 'list' variable to run more queries on the list:
      const sitePagesListID = await sitePagelist.select("Id")();
      console.log(groups);

      console.log(data.departments);

      for (const departmentName of data.departments) {
        const groupID: any = groups.find(
          (g) => g.displayName === departmentName,
        );

        const peoplePickerInput = JSON.stringify([
          {
            Key: `c:0o.c|federateddirectoryclaimprovider|${groupID.id}`,
            DisplayText: `${departmentName} Members`,
            IsResolved: true,
            Description: departmentName,
            EntityType: "SecGroup",
            EntityData: {},
            MultipleMatches: [],
            ProviderName: `FederatedDirectoryClaimProvider`,
            ProviderDisplayName: `Federated Directory`,
          },
        ]);
        const sharePayload = {
          peoplePickerInput: peoplePickerInput,
          roleValue: "role:1073741827",
          sendEmail: false,
          emailBody: null,
          includeAnonymousLinkInEmail: false,
          propagateAcl: true,
          useSimplifiedRoles: true,
        };
        const shareResponse = await this._webPartContext.spHttpClient.post(
          `${siteUrl}/_api/web/Lists(@a1)/GetItemById(@a2)/ShareObject?@a1='${sitePagesListID.Id}'&@a2='${items[0].Id}'`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept: "application/json;odata=nometadata",
              "Content-Type": "application/json;odata=nometadata",
              "odata-version": "",
            },
            body: JSON.stringify(sharePayload),
          },
        );

        if (!shareResponse.ok) {
          console.error(
            `Failed to share with user ${departmentName}:`,
            await shareResponse.text(),
          );
        } else {
          // console.log(`Successfully shared with user ${user.text}`);
        }
      }
    }

    return `${items[0].FileRef}?Mode=Edit`;

    // return `/sites/${sitetitle}/SitePages/${sanitizedTitle}.aspx?Mode=Edit`;
  }
}
