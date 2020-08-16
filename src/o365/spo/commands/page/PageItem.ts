export interface PageItem {
  ListItemAllFields: {
    CommentsDisabled: boolean,
    FileSystemObjectType: number,
    Id: number,
    ServerRedirectedEmbedUri: any,
    ServerRedirectedEmbedUrl: string,
    ContentTypeId: string,
    ComplianceAssetId: any,
    WikiField: any,
    Title: string,
    ClientSideApplicationId: string,
    PageLayoutType: string,
    CanvasContent1: string,
    BannerImageUrl: {
      Description: string,
      Url: string
    },
    Description: string,
    PromotedState: number,
    FirstPublishedDate: Date,
    LayoutWebpartsContent: string,
    OData__AuthorBylineId: any,
    _AuthorBylineStringId: any,
    OData__TopicHeader: any,
    OData__SPSitePageFlags: any,
    OData__OriginalSourceUrl: any,
    OData__OriginalSourceSiteId: any,
    OData__OriginalSourceWebId: any,
    OData__OriginalSourceListId: any,
    OData__OriginalSourceItemId: any,
    ID: 20,
    Created: Date,
    AuthorId: number,
    Modified: Date,
    EditorId: number,
    OData__CopySource: any,
    CheckoutUserId: any,
    OData__UIVersionString: string,
    GUID: any
  },
  CheckInComment: string,
  CheckOutType: number,
  ContentTag: string,
  CustomizedPageStatus: number,
  ETag: string,
  Exists: boolean,
  IrmEnabled: boolean,
  Length: number,
  Level: number,
  LinkingUri: string,
  LinkingUrl: string,
  MajorVersion: number,
  MinorVersion: number,
  Name: string,
  ServerRelativeUrl: string,
  TimeCreated: Date,
  TimeLastModified: Date,
  Title: string,
  UIVersion: number,
  UIVersionLabel: string,
  UniqueId: string,
  layoutType: string,
  commentsDisabled?: boolean,
  numSections?: number,
  numControls?: number,
  title?: string
}