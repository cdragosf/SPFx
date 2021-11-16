import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { AgendaPropertyPane } from './AgendaPropertyPane';
import { IEvent } from '../models/IEvent';
import { graph } from '@pnp/graph';
import { GraphCalendar } from '../services/Graph';

export interface IAgendaAdaptiveCardExtensionProps {
  title: string;
  iconProperty: string;
}

export interface IAgendaAdaptiveCardExtensionState {
  eventsCount: number;
  loading: boolean;
  events: IEvent[];
}

const CARD_VIEW_REGISTRY_ID: string = 'Agenda_CARD_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'Agenda_QUICK_VIEW';

export default class AgendaAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  IAgendaAdaptiveCardExtensionProps,
  IAgendaAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: AgendaPropertyPane | undefined;
  private GraphCalendar: GraphCalendar = new GraphCalendar();


  public onInit(): Promise<void> {
    graph.setup({
      spfxContext: this.context
    });

    this.state = {
      loading: true,
      eventsCount: null,
      events: null
    };

    this.GraphCalendar.getTodayEvents().then(response => {
      this.setState({ events: response, loading: false, eventsCount: response.length });
    });

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());

    return Promise.resolve();
  }

  public get title(): string {
    return this.properties.title;
  }

  protected get iconProperty(): string {
    return this.properties.iconProperty || require('./assets/SharePointLogo.svg');
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'Agenda-property-pane'*/
      './AgendaPropertyPane'
    )
      .then(
        (component) => {
          this._deferredPropertyPane = new component.AgendaPropertyPane();
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
