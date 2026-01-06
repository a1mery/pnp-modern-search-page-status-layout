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
    isDeleting: boolean;
    isLoadingDetails: boolean;
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
            isDeleting: false,
            isLoadingDetails: false,
            pageDetails: null,
            error: null
        };
        this.sp = spfi().using(SPFx({ pageContext: this.props.context }));
    }

    private handleEditClick = (): void => {
        try {
            if (this.props.pageUrl) {
                // Open page in edit mode
                window.open(`${this.props.pageUrl}?Mode=Edit`, '_blank');
            }
        } catch (error) {
            this.setState({ error: 'Failed to open page in edit mode.' });
        }
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
                .select("ViewsLifeTime", "ViewsRecent", "Modified")();

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
                viewCount: item.ViewsLifeTime || 0,
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

    private closeDeleteDialog = (): void => {
        if (!this.state.isDeleting) {
            this.setState({ showDeleteDialog: false, error: null });
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
            isDeleting, 
            isLoadingDetails,
            pageDetails,
            error 
        } = this.state;

        return (
            <div>
                <Stack horizontal tokens={{ childrenGap: 4 }}>
                    <IconButton
                        iconProps={{ iconName: 'Edit' }}
                        title="Edit"
                        ariaLabel="Edit page"
                        onClick={this.handleEditClick}
                    />
                    <IconButton
                        iconProps={{ iconName: 'Info' }}
                        title="Details"
                        ariaLabel="View page details"
                        onClick={this.handleDetailsClick}
                    />
                    <IconButton
                        iconProps={{ iconName: 'Delete' }}
                        title="Delete"
                        ariaLabel="Delete page"
                        onClick={this.handleDeleteClick}
                    />
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