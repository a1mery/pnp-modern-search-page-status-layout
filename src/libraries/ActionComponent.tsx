import * as React from 'react';
import { BaseWebComponent } from '@pnp/modern-search-extensibility';
import * as ReactDOM from 'react-dom';
import { ServiceScope } from '@microsoft/sp-core-library';
import { PageContext } from '@microsoft/sp-page-context';
import { SPFx, spfi } from "@pnp/sp";
import { IconButton, PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Stack } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import "@pnp/sp/clientside-pages";
import "@pnp/sp/comments/clientside-page";
import "@pnp/sp/webs";
import "@pnp/sp/items";
import "@pnp/sp/lists";
import { ILikedByInformation } from '@pnp/sp/comments/types';
import { Web } from "@pnp/sp/webs";


export class ActionWebComponent extends BaseWebComponent {
    private _pageContext: PageContext;

    public async connectedCallback() {
        let props = this.resolveAttributes();
        let serviceScope: ServiceScope = this._serviceScope;
        serviceScope.whenFinished(() => {
            this._pageContext = serviceScope.consume(PageContext.serviceKey);
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
    itemId?: string;
    pageUrl?: string;
    siteUrl?: string;
    webUrl?: string;
    title?: string;
    promotedState?: string;
    isNews?: string;
    showDeleteButton?: string;
}

export interface IPageDetails {
    viewCount: number;
    likeCount: number;
    commentCount: number;
    lastModifiedDate: string;
}

export interface ICustomComponentState {
    showDeleteDialog: boolean;
    showDetailsDialog: boolean;
    showPromoteDialog: boolean;
    isDeleting: boolean;
    isLoadingDetails: boolean;
    isPromoting: boolean;
    pageDetails: IPageDetails | null;
    error: string | null;
}

export class CustomComponent extends React.Component<ICustomComponentProps, ICustomComponentState> {
    private sp: ReturnType<typeof spfi>;

    public constructor(props: ICustomComponentProps) {
        super(props);

        this.state = {
            showDeleteDialog: false,
            showDetailsDialog: false,
            showPromoteDialog: false,
            isDeleting: false,
            isLoadingDetails: false,
            isPromoting: false,
            pageDetails: null,
            error: null
        };
        this.sp = spfi().using(SPFx({ pageContext: this.props.context }));
    }

    private handlePromoteClick = (): void => {
        this.setState({ showPromoteDialog: true });
    }

    private handleDeleteClick = (): void => {
        this.setState({ showDeleteDialog: true });
    }

    private handleDetailsClick = async (): Promise<void> => {
        this.setState({ 
            showDetailsDialog: true, 
            isLoadingDetails: true,
            error: null 
        });

        try {
            await this.loadPageDetails();
        } catch (error) {
            this.setState({ 
                error: 'Failed to load page details.',
                isLoadingDetails: false 
            });
        }
    }

    private loadPageDetails = async (): Promise<void> => {
        try {
            const { webUrl, itemId } = this.props;
            
            if (!webUrl || !itemId) {
                throw new Error('Missing required parameters');
            }

            const web = Web([this.sp.web, webUrl]);
            
            // Get the Site Pages library
            const list = web.lists.getByTitle("Site Pages");
            
            // Get the page item
            const item = await list.items.getById(parseInt(itemId))
                .select("Modified")();

            // Get like information
            let likeCount = 0;
            try {
                const likeInfo: ILikedByInformation = await list.items
                    .getById(parseInt(itemId))
                    .getLikedByInformation();
                likeCount = likeInfo.likeCount || 0;
            } catch (error) {
                // Likes might not be enabled, continue with 0
                console.warn('Could not retrieve like information:', error);
            }

            // Get view count
            const host = new URL(webUrl).hostname;
            
            // Get site and web IDs from the target web
            const siteResponse = await fetch(`${webUrl}/_api/site?$select=Id`, {
                headers: { accept: "application/json;odata=nometadata" },
                credentials: "same-origin"
            });
            const siteData = await siteResponse.json();
            const siteId = siteData.Id;
            
            const webResponse = await fetch(`${webUrl}/_api/web?$select=Id`, {
                headers: { accept: "application/json;odata=nometadata" },
                credentials: "same-origin"
            });
            const webData = await webResponse.json();
            const webId = webData.Id;
            
            const listId = await list.select("Id")().then(l => l.Id);
            
            const url = `https://${host}/_api/v2.1/sites/${host},${siteId},${webId}/lists/${listId}/items/${itemId}/driveItem?$select=id,analytics&$expand=analytics($expand=allTime)`;

            const res = await fetch(url, {
                method: "GET",
                headers: { accept: "application/json;odata=nometadata" },
                credentials: "same-origin"
            });

            if (!res.ok) throw new Error('message');
            const json = await res.json();
            const views = Number(json?.analytics?.allTime?.access?.actionCount ?? 0);


            // Get comment count
            let commentCount = 0;
            try {
                const comments = await list.items
                    .getById(parseInt(itemId))
                    .comments();
                commentCount = comments.length;
            } catch (error) {
                // Comments might not be enabled, continue with 0
                console.warn('Could not retrieve comments:', error);
            }

            const pageDetails: IPageDetails = {
                viewCount: views,
                likeCount: likeCount,
                commentCount: commentCount,
                lastModifiedDate: item.Modified
            };

            this.setState({ 
                pageDetails,
                isLoadingDetails: false 
            });
        } catch (error) {
            console.error('Error loading page details:', error);
            throw error;
        }
    }

    private confirmDelete = async (): Promise<void> => {
        this.setState({ isDeleting: true, error: null });

        try {
            const { webUrl, itemId } = this.props;
            
            if (!webUrl || !itemId) {
                throw new Error('Missing required parameters');
            }

            const web = Web([this.sp.web, webUrl]);
            const list = web.lists.getByTitle("Site Pages");
            
            await list.items.getById(parseInt(itemId)).recycle();

            // Close dialog and refresh page
            this.setState({ 
                showDeleteDialog: false,
                isDeleting: false 
            });
            
            // Reload the page to reflect changes
            window.location.reload();
        } catch (error) {
            console.error('Error deleting page:', error);
            this.setState({ 
                error: 'Failed to delete page.',
                isDeleting: false 
            });
        }
    }

    private confirmPromote = async (): Promise<void> => {
        this.setState({ isPromoting: true, error: null });

        try {
            const { webUrl, pageUrl } = this.props;
            
            if (!webUrl || !pageUrl) {
                throw new Error('Missing required parameters');
            }

            const web = Web([this.sp.web, webUrl]);
            
            // Get the page file path from the URL
            const url = new URL(pageUrl);
            const pagePath = url.pathname;
            
            // Load the page and promote to news
            const page = await web.loadClientsidePage(pagePath);
            await page.promoteToNews();

            // Close dialog and refresh page
            this.setState({ 
                showPromoteDialog: false,
                isPromoting: false 
            });
            
            // Reload the page to reflect changes
            window.location.reload();
        } catch (error) {
            console.error('Error promoting page:', error);
            this.setState({ 
                error: 'Failed to promote page to news.',
                isPromoting: false 
            });
        }
    }

    private closeDeleteDialog = (): void => {
        if (!this.state.isDeleting) {
            this.setState({ showDeleteDialog: false, error: null });
        }
    }

    private closePromoteDialog = (): void => {
        if (!this.state.isPromoting) {
            this.setState({ showPromoteDialog: false, error: null });
        }
    }

    private closeDetailsDialog = (): void => {
        this.setState({ 
            showDetailsDialog: false, 
            pageDetails: null,
            error: null 
        });
    }

    private formatDate = (dateString: string): string => {
        try {
            const date = new Date(dateString);
            return date.toLocaleDateString(undefined, { 
                year: 'numeric', 
                month: 'long', 
                day: 'numeric',
                hour: '2-digit',
                minute: '2-digit'
            });
        } catch (error) {
            return dateString;
        }
    }

    public render(): React.ReactElement<ICustomComponentProps> {
        const { 
            showDeleteDialog, 
            showDetailsDialog, 
            showPromoteDialog,
            isDeleting, 
            isLoadingDetails,
            isPromoting,
            pageDetails,
            error 
        } = this.state;

        const isNews = this.props.promotedState?.toString() === '2';
        const showDelete = this.props.showDeleteButton?.toString().toLowerCase() !== 'false';

        return (
            <div>
                <Stack horizontal tokens={{ childrenGap: 4 }}>
                    {!isNews && (
                        <IconButton
                            iconProps={{ iconName: 'Megaphone' }}
                            title="Promote to News"
                            ariaLabel="Promote page to news"
                            onClick={this.handlePromoteClick}
                        />
                    )}
                    <IconButton
                        iconProps={{ iconName: 'Info' }}
                        title="Details"
                        ariaLabel="View page details"
                        onClick={this.handleDetailsClick}
                    />
                    {showDelete && (
                        <IconButton
                            iconProps={{ iconName: 'Delete' }}
                            title="Delete"
                            ariaLabel="Delete page"
                            onClick={this.handleDeleteClick}
                        />
                    )}
                </Stack>

                {/* Delete Confirmation Dialog */}
                <Dialog
                    hidden={!showDeleteDialog}
                    onDismiss={this.closeDeleteDialog}
                    dialogContentProps={{
                        type: DialogType.normal,
                        title: 'Delete Page',
                        subText: `Are you sure you want to delete "${this.props.title || 'this page'}"? This will move the page to the recycle bin.`
                    }}
                    modalProps={{
                        isBlocking: isDeleting
                    }}
                >
                    {error && (
                        <MessageBar messageBarType={MessageBarType.error}>
                            {error}
                        </MessageBar>
                    )}
                    <DialogFooter>
                        <PrimaryButton 
                            onClick={this.confirmDelete} 
                            text="Delete" 
                            disabled={isDeleting}
                        />
                        <DefaultButton 
                            onClick={this.closeDeleteDialog} 
                            text="Cancel" 
                            disabled={isDeleting}
                        />
                    </DialogFooter>
                    {isDeleting && <Spinner label="Deleting..." size={SpinnerSize.large} />}
                </Dialog>

                {/* Promote to News Confirmation Dialog */}
                <Dialog
                    hidden={!showPromoteDialog}
                    onDismiss={this.closePromoteDialog}
                    dialogContentProps={{
                        type: DialogType.normal,
                        title: 'Promote to News',
                        subText: `Are you sure you want to promote "${this.props.title || 'this page'}" to news? This will feature the page as a news article.`
                    }}
                    modalProps={{
                        isBlocking: isPromoting
                    }}
                >
                    {error && (
                        <MessageBar messageBarType={MessageBarType.error}>
                            {error}
                        </MessageBar>
                    )}
                    <DialogFooter>
                        <PrimaryButton 
                            onClick={this.confirmPromote} 
                            text="Promote" 
                            disabled={isPromoting}
                        />
                        <DefaultButton 
                            onClick={this.closePromoteDialog} 
                            text="Cancel" 
                            disabled={isPromoting}
                        />
                    </DialogFooter>
                    {isPromoting && <Spinner label="Promoting..." size={SpinnerSize.large} />}
                </Dialog>

                {/* Details Dialog */}
                <Dialog
                    hidden={!showDetailsDialog}
                    onDismiss={this.closeDetailsDialog}
                    dialogContentProps={{
                        type: DialogType.normal,
                        title: `Page Details - ${this.props.title || 'Page'}`
                    }}
                    minWidth={400}
                >
                    {isLoadingDetails ? (
                        <Stack horizontalAlign="center" tokens={{ childrenGap: 10, padding: 20 }}>
                            <Spinner label="Loading page details..." size={SpinnerSize.large} />
                        </Stack>
                    ) : error ? (
                        <MessageBar messageBarType={MessageBarType.error}>
                            {error}
                        </MessageBar>
                    ) : pageDetails ? (
                        <Stack tokens={{ childrenGap: 15, padding: '10px 0' }}>
                            <Stack horizontal tokens={{ childrenGap: 10 }}>
                                <Text variant="medium" style={{ fontWeight: 600, minWidth: 150 }}>
                                    Views:
                                </Text>
                                <Text variant="medium">{pageDetails.viewCount}</Text>
                            </Stack>
                            <Stack horizontal tokens={{ childrenGap: 10 }}>
                                <Text variant="medium" style={{ fontWeight: 600, minWidth: 150 }}>
                                    Likes:
                                </Text>
                                <Text variant="medium">{pageDetails.likeCount}</Text>
                            </Stack>
                            <Stack horizontal tokens={{ childrenGap: 10 }}>
                                <Text variant="medium" style={{ fontWeight: 600, minWidth: 150 }}>
                                    Comments:
                                </Text>
                                <Text variant="medium">{pageDetails.commentCount}</Text>
                            </Stack>
                            <Stack horizontal tokens={{ childrenGap: 10 }}>
                                <Text variant="medium" style={{ fontWeight: 600, minWidth: 150 }}>
                                    Last Modified:
                                </Text>
                                <Text variant="medium">{this.formatDate(pageDetails.lastModifiedDate)}</Text>
                            </Stack>
                        </Stack>
                    ) : null}
                    <DialogFooter>
                        <DefaultButton onClick={this.closeDetailsDialog} text="Close" />
                    </DialogFooter>
                </Dialog>
            </div>
        );
    }
}