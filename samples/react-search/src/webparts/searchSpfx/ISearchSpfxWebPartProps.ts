export interface ISearchSpfxWebPartProps {
  title: string;
  query?: string;
  sorting?: string;
  filtering?: string;
  template?: string;
  maxResults?: number;
  external?: boolean;
  externalUrl?: string;
}
