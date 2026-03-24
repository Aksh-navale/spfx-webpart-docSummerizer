import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { NotificationSummaryPropertyPane } from './NotificationSummaryPropertyPane';
import { QuickView } from './quickView/QuickView';

export interface INotificationSummaryAdaptiveCardExtensionProps {
  title: string;
}

export interface INotificationSummaryAdaptiveCardExtensionState {
}

const CARD_VIEW_REGISTRY_ID: string = 'NotificationSummary_CARD_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'NotificationSummary_QUICK_VIEW';


export default class NotificationSummaryAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  INotificationSummaryAdaptiveCardExtensionProps,
  INotificationSummaryAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: NotificationSummaryPropertyPane;

  public onInit(): Promise<void> {
    this.state = { };

    // registers the card view to be shown in a dashboard
    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    // registers the quick view to open via QuickView action
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());

    return Promise.resolve();
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'NotificationSummary-property-pane'*/
      './NotificationSummaryPropertyPane'
    )
      .then(
        (component) => {
          this._deferredPropertyPane = new component.NotificationSummaryPropertyPane();
        }
      );
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane?.getPropertyPaneConfiguration();
  }
}
