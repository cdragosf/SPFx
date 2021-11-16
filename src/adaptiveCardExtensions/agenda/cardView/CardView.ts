import {
  BaseBasicCardView,
  IBasicCardParameters,
  IExternalLinkCardAction,
  IQuickViewCardAction,
  ICardButton
} from '@microsoft/sp-adaptive-card-extension-base';

import * as strings from 'AgendaAdaptiveCardExtensionStrings';
import { IAgendaAdaptiveCardExtensionProps, IAgendaAdaptiveCardExtensionState, QUICK_VIEW_REGISTRY_ID } from '../AgendaAdaptiveCardExtension';

export class CardView extends BaseBasicCardView<IAgendaAdaptiveCardExtensionProps, IAgendaAdaptiveCardExtensionState> {
  public get cardButtons(): [ICardButton] | [ICardButton, ICardButton] | undefined {
    let buttons = undefined;

    if (!this.state.loading && this.state.eventsCount != 0) {
      buttons = [{
        title: strings.QuickViewButton,
        action: {
          type: 'QuickView',
          parameters: {
            view: QUICK_VIEW_REGISTRY_ID
          }
        }
      }];
    }

    return buttons;
  }

  public get data(): IBasicCardParameters {
    let textToShow: string = "";

    if (!this.state.loading) {
      if (this.state.eventsCount === 0) {
        textToShow = strings.NoNextEvents;
      } else {
        textToShow = this.state.eventsCount.toString() + strings.NextEvents;
      }
    }
    else {
      textToShow = strings.Loading;
    }
    return {
      primaryText: textToShow,
    };
  }

  public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
    return {
      type: 'ExternalLink',
      parameters: {
        target: 'https://outlook.office.com/calendar/view/day'
      }
    };
  }
}
