import * as React from 'react';
import { BaseWebComponent } from '@pnp/modern-search-extensibility';
import * as ReactDOM from 'react-dom';
import { ServiceScope } from '@microsoft/sp-core-library';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { SPFx, spfi } from "@pnp/sp";
import { IClientsidePage } from "@pnp/sp/clientside-pages";
import { IIconProps } from '@fluentui/react/lib/Icon';
import { IconButton } from '@fluentui/react/lib/Button';
import "@pnp/sp/clientside-pages";
import "@pnp/sp/comments/clientside-page";
import "@pnp/sp/webs";
import "@pnp/sp/items"
import "@pnp/sp/lists"
import { ILikedByInformation } from '@pnp/sp/comments/types';
import { Web } from "@pnp/sp/webs";


export class ActionWebComponent extends BaseWebComponent {
        private _spHttpClient: SPHttpClient;
    private _pageContext: PageContext;
    private _currentWebUrl: string;
 public async connectedCallback() {

        let props = this.resolveAttributes();
        let serviceScope: ServiceScope = this._serviceScope;
        let _spHttpClient: SPHttpClient;
        let _pageContext: PageContext;
        serviceScope.whenFinished(() => {
            this._spHttpClient = serviceScope.consume(SPHttpClient.serviceKey);
            this._pageContext = serviceScope.consume(PageContext.serviceKey);
            this._currentWebUrl = this._pageContext.web.absoluteUrl;
        });
        const customComponent = <CustomComponent context={this._pageContext} {...props} />;
        ReactDOM.render(customComponent, this);
    }


    protected onDispose(): void {
        ReactDOM.unmountComponentAtNode(this);
    }
}

export interface ICustomComponentProps {
    context: PageContext;
}

export interface ICustomComponenState {

}

export class CustomComponent extends React.Component<ICustomComponentProps, ICustomComponenState> {


    private sp: ReturnType<typeof spfi>;
    public constructor(props: ICustomComponentProps) {
        super(props);

        this.state = {
        };
        this.sp = spfi().using(SPFx({ pageContext: this.props.context }));
    }
}