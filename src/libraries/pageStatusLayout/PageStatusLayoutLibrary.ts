import {
  IAdaptiveCardAction,
  IComponentDefinition,
  IDataSourceDefinition,
  IExtensibilityLibrary,
  ILayoutDefinition,
  IQueryModifierDefinition,
  ISuggestionProviderDefinition,
  LayoutType,
  LayoutRenderType,
  BaseLayout
} from "@pnp/modern-search-extensibility";
import { ActionWebComponent } from "../ActionComponent";
import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";


export interface IPageStatusLayoutProperties {
}

export class PageStatusLayout extends BaseLayout<IPageStatusLayoutProperties> {
}


export class PageStatusLayoutLibrary implements IExtensibilityLibrary{
public static readonly serviceKey: ServiceKey<PageStatusLayoutLibrary> =
  ServiceKey.create<PageStatusLayoutLibrary>('SPFx:PageStatusLayoutLibrary', PageStatusLayoutLibrary);

constructor(serviceScope: ServiceScope) {
  // Constructor intentionally left minimal
}

  getCustomLayouts(): ILayoutDefinition[] {
    return [
      {
        name: 'Manage pages',
        iconName: 'DocumentManagement',
        key: 'ManagePagesLayout',
        type: LayoutType.Results,
        renderType: LayoutRenderType.Handlebars,
        templateContent: require('./manage-pages-layout.html').default.toString(),
        serviceKey: ServiceKey.create<PageStatusLayout>('PageStatusLayout', PageStatusLayout)
      }
    ];
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
