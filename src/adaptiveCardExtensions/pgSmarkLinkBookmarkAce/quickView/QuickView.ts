import { ISPFxAdaptiveCard, BaseAdaptiveCardView , IActionArguments} from '@microsoft/sp-adaptive-card-extension-base';
// import * as strings from 'PgSmarkLinkBookmarkAceAdaptiveCardExtensionStrings';
import { IPgSmarkLinkBookmarkAceAdaptiveCardExtensionProps, IPgSmarkLinkBookmarkAceAdaptiveCardExtensionState } from '../PgSmarkLinkBookmarkAceAdaptiveCardExtension';
// import {
//   deleteRecord
// } from '../pgSmarkLink.service'

export interface IQuickViewData {
  subTitle: string;
  title: string;
  bookmarkByAdmin: any[];
  bookmarkByUser: any[];

}

export class QuickView extends BaseAdaptiveCardView<
  IPgSmarkLinkBookmarkAceAdaptiveCardExtensionProps,
  IPgSmarkLinkBookmarkAceAdaptiveCardExtensionState,
  IQuickViewData
> {
  public get data(): IQuickViewData {
    return {
      subTitle: "VIVA SubTitle",
      title: "Viva Title",
      bookmarkByAdmin: this.state.bookmarkByAdmin,
      bookmarkByUser: this.state.bookmarkByUser
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return require('./template/QuickViewTemplate.json');
  }

  public async onAction(action: IActionArguments): Promise<void> {
    if (action.type !== 'Submit') { return ;}
   
  }
}