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
import { PropertyPaneCheckbox } from "@microsoft/sp-property-pane";
import { ActionWebComponent } from "../ActionComponent";
import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";


export interface IPageStatusLayoutProperties {
  showDeleteButton: boolean;
}

export class PageStatusLayout extends BaseLayout<IPageStatusLayoutProperties> {

  public onInit(): void {
    // Default to showing the delete button when property is not yet set
    if (this.properties.showDeleteButton === undefined) {
      this.properties.showDeleteButton = true;
    }
  }

  public getPropertyPaneFieldsConfiguration(): import("@microsoft/sp-property-pane").IPropertyPaneField<any>[] {
    return [
      PropertyPaneCheckbox('layoutProperties.showDeleteButton', {
        text: 'Show delete button',
        checked: true
      })
    ];
  }
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
