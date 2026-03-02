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
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
import { ActionWebComponent } from "../ActionComponent";
import { BadgeWebComponent } from "../BadgeComponent";
import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";


export interface IPageStatusLayoutProperties {
  showDeleteButton: boolean;
  showPromoteButton: boolean;
  newsBadgeColor: string;
  pageBadgeColor: string;
}

export class PageStatusLayout extends BaseLayout<IPageStatusLayoutProperties> {

  public onInit(): void {
    if (this.properties.showDeleteButton === undefined) {
      this.properties.showDeleteButton = true;
    }
    if (this.properties.showPromoteButton === undefined) {
      this.properties.showPromoteButton = true;
    }
    if (!this.properties.newsBadgeColor) {
      this.properties.newsBadgeColor = '#107c10';
    }
    if (!this.properties.pageBadgeColor) {
      this.properties.pageBadgeColor = '#605e5c';
    }
  }

  public getPropertyPaneFieldsConfiguration(): import("@microsoft/sp-property-pane").IPropertyPaneField<any>[] {
    return [
      PropertyPaneCheckbox('layoutProperties.showDeleteButton', {
        text: 'Show delete button',
        checked: true
      }),
      PropertyPaneCheckbox('layoutProperties.showPromoteButton', {
        text: 'Show promote button',
        checked: true
      }),
      PropertyFieldColorPicker('layoutProperties.newsBadgeColor', {
        label: 'News badge color',
        selectedColor: this.properties.newsBadgeColor,
        onPropertyChange: this.onPropertyUpdate.bind(this),
        properties: { layoutProperties: this.properties },
        style: PropertyFieldColorPickerStyle.Inline,
        alphaSliderHidden: true,
        key: 'newsBadgeColorField'
      }),
      PropertyFieldColorPicker('layoutProperties.pageBadgeColor', {
        label: 'Page badge color',
        selectedColor: this.properties.pageBadgeColor,
        onPropertyChange: this.onPropertyUpdate.bind(this),
        properties: { layoutProperties: this.properties },
        style: PropertyFieldColorPickerStyle.Inline,
        alphaSliderHidden: true,
        key: 'pageBadgeColorField'
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
      },
      {
        componentName: 'badge-component',
        componentClass: BadgeWebComponent
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
