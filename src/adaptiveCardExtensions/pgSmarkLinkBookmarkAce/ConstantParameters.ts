import { AdaptiveCardExtensionContext } from '@microsoft/sp-adaptive-card-extension-base';
let spContext : AdaptiveCardExtensionContext;
export let parameters = {
 getBookmarkItems : `https://pgonedev.sharepoint.com/_api/Web/Lists/GetByTitle('Bookmark')/Items?$select=ID,Title,SiteAbsoluteUrl,PageName,PageURL,AddedBy,UserName,UserEmail,PageID,IsActive&$filter=((IsActive eq 1) and ((UserEmail eq `+"'"+spContext.pageContext.user.email+"'"+`) or (AddedBy eq 'Admin')))&$orderby=Modified desc`,

}