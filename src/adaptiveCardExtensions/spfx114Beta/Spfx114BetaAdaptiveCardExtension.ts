import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension, ICachedLoadParameters, ICacheSettings } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { Spfx114BetaPropertyPane } from './Spfx114BetaPropertyPane';

export interface ISpfx114BetaAdaptiveCardExtensionProps {
  title: string;
}

export interface ISpfx114BetaAdaptiveCardExtensionState {
  title: string;
}

const CARD_VIEW_REGISTRY_ID: string = 'Spfx114Beta_CARD_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'Spfx114Beta_QUICK_VIEW';

export default class Spfx114BetaAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  ISpfx114BetaAdaptiveCardExtensionProps,
  ISpfx114BetaAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: Spfx114BetaPropertyPane | undefined;

  public onInit(cachedLoadParameters?: ICachedLoadParameters): Promise<void> {

    //expecting cachedLoadParameters to have "Setting from onInit" on next page load
    console.log("cache - onInit - %o", cachedLoadParameters);

    this.state = { 
      title: 'Setting from onInit'
    };

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());

    return Promise.resolve();
  }

  protected getCacheSettings(): Partial<ICacheSettings> {
    return {
      isEnabled: true, // can be set to false to disable caching
      expiryTimeInSeconds: 86400, // controls how long until the cached card and state are stale
      cachedCardView: () => new CardView() // function that returns the custom Card view that will be used to generate the cached card
    };
  }

  protected getCachedState(state: ISpfx114BetaAdaptiveCardExtensionState): Partial<ISpfx114BetaAdaptiveCardExtensionState> {
    console.log("cache - getCachedState - %o", state);
    return {
      title: state.title
    }
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'Spfx114Beta-property-pane'*/
      './Spfx114BetaPropertyPane'
    )
      .then(
        (component) => {
          this._deferredPropertyPane = new component.Spfx114BetaPropertyPane();
        }
      );
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane!.getPropertyPaneConfiguration();
  }
}
