import {
  BasePrimaryTextCardView,
  IPrimaryTextCardParameters,
  ICardButton
} from '@microsoft/sp-adaptive-card-extension-base';

import { INotificationSummaryAdaptiveCardExtensionProps, INotificationSummaryAdaptiveCardExtensionState, QUICK_VIEW_REGISTRY_ID } from '../NotificationSummaryAdaptiveCardExtension';

export class CardView extends BasePrimaryTextCardView<
INotificationSummaryAdaptiveCardExtensionProps,
INotificationSummaryAdaptiveCardExtensionState
> {

  public get data(): IPrimaryTextCardParameters {
  return {
    primaryText: "👋 Good Morning",
    description: "AI Notification Summary",
    title: "🤖 Copilot Assistant"
  };
}

  public get cardButtons(): [ICardButton] {
  return [
    {
      title: "View Details",
      action: {
        type: 'QuickView',
        parameters: {
          view: QUICK_VIEW_REGISTRY_ID
        }
      }
    }
  ];
}
}