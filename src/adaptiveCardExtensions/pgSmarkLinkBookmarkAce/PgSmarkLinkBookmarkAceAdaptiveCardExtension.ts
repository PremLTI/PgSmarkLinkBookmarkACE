import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { PgSmarkLinkBookmarkAcePropertyPane } from './PgSmarkLinkBookmarkAcePropertyPane';
import { SPHttpClient } from '@microsoft/sp-http'

import {
  IListItem,
  fetchListItems
} from './pgSmarkLink.service';

export interface IPgSmarkLinkBookmarkAceAdaptiveCardExtensionProps {
  title: string;

}

export interface IPgSmarkLinkBookmarkAceAdaptiveCardExtensionState {
  bookmarkByAdmin: IListItem[];
  bookmarkByUser: IListItem[];
}

const CARD_VIEW_REGISTRY_ID: string = 'PgSmarkLinkBookmarkAce_CARD_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'PgSmarkLinkBookmarkAce_QUICK_VIEW';

export default class PgSmarkLinkBookmarkAceAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  IPgSmarkLinkBookmarkAceAdaptiveCardExtensionProps,
  IPgSmarkLinkBookmarkAceAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: PgSmarkLinkBookmarkAcePropertyPane | undefined;

  public AddedByAdmin: any[];
  public AddedByUser: any[];

  public async onInit(): Promise<void> {
    this.state = {
      bookmarkByAdmin: [],
      bookmarkByUser: []
    };

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());

    if (this.properties.title != null) {
     
      let listitems = await fetchListItems(this.context);
    await this.checkifitemexist(listitems);

    }

    return Promise.resolve();
  }

  // checking the item is deleted in soruce or not. 
  protected checkifitemexist = async (items: any) => {

    this.AddedByAdmin = [];
    this.AddedByUser = [];
// no condition if its admin. 
    items.map(async (ele: any, index: number) => {

      if (ele.addedBy == "Admin") {
        // let pushItemByAdmin = {
        //   id: ele.id,
        //   pageid: ele.pageid,
        //   siteName: ele.siteName,
        //   siteabsoluteUrl: ele.siteabsoluteUrl,
        //   pageName: ele.pageName,
        //   pageURL: ele.pageURL,
        //   userName: ele.userName,
        //   userEmail: ele.userEmail,
        //   addedBy: ele.addedBy,
        //   index: index,
        //   IsActive: ele.IsActive

        // };
        this.AddedByAdmin.push(ele);
      }

      // condition to check page id is present or not in the site absolute url.
      if (ele.addedBy == "User") {
        const response = await (await this.context.spHttpClient.get(
          ele.siteabsoluteUrl + `/_api/web/lists/getbyTitle('Site Pages')/GetItemById(` + ele.pageid + `)`,
          SPHttpClient.configurations.v1
        )).json();

        if (response.Id != undefined) {
          // let pushItemByUser = {
          //   id: ele.id,
          //   pageid: ele.pageid,
          //   siteName: ele.siteName,
          //   siteabsoluteUrl: ele.siteabsoluteUrl,
          //   pageName: ele.pageName,
          //   pageURL: ele.pageURL,
          //   userName: ele.userName,
          //   userEmail: ele.userEmail,
          //   addedBy: ele.addedBy,
          //   index: index,
          //   IsActive: ele.IsActive

          // };
          this.AddedByUser.push(ele);
        }
        // if the site id not exist, then setting IsActive = 0 and on the next time refresh of the page, based on conditon the isactive = 0 items will not appear in ACE.
        else {
          const body: string = JSON.stringify({
            'IsActive': 0
          });
          await (await this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Bookmark')/items(${ele.id})`,
            SPHttpClient.configurations.v1, {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'Content-type': 'application/json;odata=nometadata',
              'odata-version': '',
              'IF-MATCH': '*',
              'X-HTTP-Method': 'MERGE'
            },
            body: body

          })).json();


        }

      }
      // setting up state for both admin and User variables.
      this.setState({ bookmarkByAdmin: this.AddedByAdmin, bookmarkByUser: this.AddedByUser })
    })

  }


  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'PgSmarkLinkBookmarkAce-property-pane'*/
      './PgSmarkLinkBookmarkAcePropertyPane'
    )
      .then(
        (component) => {
          this._deferredPropertyPane = new component.PgSmarkLinkBookmarkAcePropertyPane();
        }
      );
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane?.getPropertyPaneConfiguration();
  }

  // protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any) {

  // }
}
