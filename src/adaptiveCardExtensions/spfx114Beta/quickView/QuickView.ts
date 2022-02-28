import { ISPFxAdaptiveCard, BaseAdaptiveCardView } from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'Spfx114BetaAdaptiveCardExtensionStrings';
import { ISpfx114BetaAdaptiveCardExtensionProps, ISpfx114BetaAdaptiveCardExtensionState } from '../Spfx114BetaAdaptiveCardExtension';

export interface IQuickViewData {
  subTitle: string;
  title: string;
}

export class QuickView extends BaseAdaptiveCardView<
  ISpfx114BetaAdaptiveCardExtensionProps,
  ISpfx114BetaAdaptiveCardExtensionState,
  IQuickViewData
> {
  public get data(): IQuickViewData {
    return {
      subTitle: strings.SubTitle,
      title: strings.Title,
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return require('./template/QuickViewTemplate.json');
  }
}