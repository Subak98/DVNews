export interface IList {
  Title?: string;
  Link: string;
}
export interface IFeaturedNewsList {
  Title: string;
  date: string;
  Created: string;
  Description: string;
  Image: string;
  Link: string;
  Department: string;
  ShowNewsinHome: boolean;

  Author: IAuthor;
  AuthorByLine: IAuthor;
}
interface IAuthor {
  ID: number;
  Title: string;
  EMail: string;
}
