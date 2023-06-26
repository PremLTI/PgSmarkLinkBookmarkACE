import {
  BasePrimaryTextCardView,
  IPrimaryTextCardParameters,
  IExternalLinkCardAction,
  IQuickViewCardAction,
  ICardButton
} from '@microsoft/sp-adaptive-card-extension-base';
// import * as strings from 'PgSmarkLinkBookmarkAceAdaptiveCardExtensionStrings';
import { IPgSmarkLinkBookmarkAceAdaptiveCardExtensionProps, IPgSmarkLinkBookmarkAceAdaptiveCardExtensionState, QUICK_VIEW_REGISTRY_ID } from '../PgSmarkLinkBookmarkAceAdaptiveCardExtension';

export class CardView extends BasePrimaryTextCardView<IPgSmarkLinkBookmarkAceAdaptiveCardExtensionProps, IPgSmarkLinkBookmarkAceAdaptiveCardExtensionState> {
  public get cardButtons(): [ICardButton] | [ICardButton, ICardButton] | undefined {
    return [
      {
        title: 'View All',
        action: {
          type: 'QuickView',
          parameters: {
            view: QUICK_VIEW_REGISTRY_ID
          }
        }
      }
    ];
  }

  public get data(): IPrimaryTextCardParameters {
    return {
      primaryText: (this.state.bookmarkByUser && this.state.bookmarkByAdmin)  ? `My Bookmarks` : `Bookmark Links Not Found for your profile`,
      description: 'Your personlized Bookmarks are here',
      title: this.properties.title
    };
  }

  public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
    return {
      type: 'QuickView',
      parameters: {
        view: QUICK_VIEW_REGISTRY_ID
      }
    };
  }
}
