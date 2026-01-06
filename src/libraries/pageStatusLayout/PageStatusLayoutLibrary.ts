import {
  IAdaptiveCardAction,
  IComponentDefinition,
  IDataSourceDefinition,
  IExtensibilityLibrary,
  ILayoutDefinition,
  IQueryModifierDefinition,
  ISuggestionProviderDefinition
} from "@pnp/modern-search-extensibility";
import { ActionWebComponent } from "../ActionComponent";
import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { SPHttpClient } from "@microsoft/sp-http";
import { PageContext } from "@microsoft/sp-page-context";


export class PageStatusLayoutLibrary implements IExtensibilityLibrary{
public static readonly serviceKey: ServiceKey<PageStatusLayoutLibrary> =
  ServiceKey.create<PageStatusLayoutLibrary>('SPFx:PageStatusLayoutLibrary', PageStatusLayoutLibrary);

  private _spHttpClient: SPHttpClient;
public _pageContext: PageContext;
private _currentWebUrl: string;

constructor(serviceScope: ServiceScope) {
  serviceScope.whenFinished(() => {
    this._spHttpClient = serviceScope.consume(SPHttpClient.serviceKey);

    this._pageContext = serviceScope.consume(PageContext.serviceKey);
    this._currentWebUrl = this._pageContext.web.absoluteUrl;
  });
}

  getCustomLayouts(): ILayoutDefinition[] {
    return [];
  }
  public getCustomWebComponents(): IComponentDefinition<any>[] {
    return [
      {
        componentName: 'action-component',
        componentClass: ActionWebComponent
      }
    ];
  }
  getCustomSuggestionProviders(): ISuggestionProviderDefinition[] {
    return [];
  }
  registerHandlebarsCustomizations?(handlebarsNamespace: typeof Handlebars): void {

  }
  invokeCardAction(action: IAdaptiveCardAction): void {

  }
  getCustomQueryModifiers?(): IQueryModifierDefinition[] {
    return [];
  }
  getCustomDataSources?(): IDataSourceDefinition[] {
    return [];
  }


  public name(): string {
    return 'PageStatusLayoutLibrary';
  }
}
