import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView } from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'NotificationSummaryAdaptiveCardExtensionStrings';
import {
  INotificationSummaryAdaptiveCardExtensionProps,
  INotificationSummaryAdaptiveCardExtensionState
} from '../NotificationSummaryAdaptiveCardExtension';

export interface IQuickViewData {
  subTitle: string;
  title: string;
}

export class QuickView extends BaseAdaptiveCardQuickView<
  INotificationSummaryAdaptiveCardExtensionProps,
  INotificationSummaryAdaptiveCardExtensionState,
  IQuickViewData
> {
  public get data(): IQuickViewData {
    return {
      subTitle: strings.SubTitle,
      title: strings.Title
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return require('./template/QuickViewTemplate.json');
  }
}
