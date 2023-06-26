import { AdaptiveCardExtensionContext } from '@microsoft/sp-adaptive-card-extension-base';
import { SPHttpClient } from '@microsoft/sp-http'
// import { parameters } from './ConstantParameters';

// interface IExternalLink {

//     Description:any;
//     Url: any;
// }

// interface IPGCarouselImage {

//     Description:any;
//     Url: any;
// }

export interface IListItem {
  id: number;
  pageid: number;
  siteName: string;
  siteabsoluteUrl:string;
  pageName: string;
  pageURL: any;
  userName: string;
  userEmail: string;
  addedBy: string;
  index: number;
  IsActive:any;
  
}


// fetch the list item where,  isactive =1 and (current email id shoud match or addedby "admin")
export const fetchListItems = async (spContext: AdaptiveCardExtensionContext): Promise<IListItem[]> => {
  if (!spContext) { return Promise.reject('No spContext.'); }

  const response = await (await spContext.spHttpClient.get(
    `https://pgonedev.sharepoint.com/_api/Web/Lists/GetByTitle('Bookmark')/Items?$select=ID,Title,SiteAbsoluteUrl,PageName,PageURL,AddedBy,UserName,UserEmail,PageID,IsActive&$filter=((IsActive eq 1) and ((UserEmail eq `+"'"+spContext.pageContext.user.email+"'"+`) or (AddedBy eq 'Admin')))&$orderby=Modified desc`,
    SPHttpClient.configurations.v1
  )).json();

  if (response.value?.length > 0) {
    return Promise.resolve(response.value.map((listItem: any, index: number) => {
        return <IListItem>{
          id: listItem.ID,
          pageid: listItem.PageID != (null || undefined)? listItem.PageID : 0,
          siteName: listItem.Title,
          siteabsoluteUrl: listItem.SiteAbsoluteUrl != (null || undefined)? listItem.SiteAbsoluteUrl : "Global Bookmark",
          pageName: listItem.PageName,
          pageURL: listItem.PageURL,
          userName: listItem.UserName,
          userEmail: listItem.UserEmail,
          addedBy: listItem.AddedBy,
          index: index,
          IsActive: listItem.IsActive

        };

      }
    ));
  } else {
    return Promise.resolve([]);
  }
}


export const deleteRecord=async(context : AdaptiveCardExtensionContext, id:any)=>{

  await (await context.spHttpClient.post(`${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Bookmark')/items(${id})`,
    SPHttpClient.configurations.v1, {
    headers: {
      'Accept': 'application/json;odata=nometadata',
      'Content-type': 'application/json;odata=nometadata',
      'odata-version': '',
      'IF-MATCH': '*',
      'X-HTTP-Method': 'DELETE'
    }

  })).json();
}

