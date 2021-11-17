import { ISPFxAdaptiveCard, BaseAdaptiveCardView } from '@microsoft/sp-adaptive-card-extension-base';
import { IEvent } from '../../models/IEvent';
import * as strings from 'AgendaAdaptiveCardExtensionStrings';
import { IAgendaAdaptiveCardExtensionProps, IAgendaAdaptiveCardExtensionState } from '../AgendaAdaptiveCardExtension';

export interface IQuickViewData {
  events: IEvent[];
}

export class QuickView extends BaseAdaptiveCardView<IAgendaAdaptiveCardExtensionProps, IAgendaAdaptiveCardExtensionState, IQuickViewData> {
  public get data(): IQuickViewData {
    return {
      events: this.state.events
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return require('./template/QuickViewTemplate.json');
  }
}