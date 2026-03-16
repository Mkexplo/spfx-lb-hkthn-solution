import type { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { LbViewDbCardPropertyPane } from './LbViewDbCardPropertyPane';

export interface ILbViewDbCardAdaptiveCardExtensionProps {
  title: string;
}

export interface ILbViewDbCardAdaptiveCardExtensionState {
  currentUserIndex?: number;
  users?: any[];
  currentUserAvatarUrl?: string;
}

const CARD_VIEW_REGISTRY_ID: string = 'LbViewDbCard_CARD_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'LbViewDbCard_QUICK_VIEW';

export default class LbViewDbCardAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  ILbViewDbCardAdaptiveCardExtensionProps,
  ILbViewDbCardAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: LbViewDbCardPropertyPane;

  public onInit(): Promise<void> {
    this.state = {
      currentUserIndex: 0,
      users: [],
      currentUserAvatarUrl: ''
    };

    // registers the card view to be shown in a dashboard
    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    // registers the quick view to open via QuickView action
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());

    // Fetch users from RegisteredUsers list
    this.fetchUsersWithAvatars();

    return Promise.resolve();
  }

  private async fetchUsersWithAvatars(): Promise<void> {
    try {
      // Fetch all users from RegisteredUsers list
      const usersUrl = `https://ygc8n.sharepoint.com/sites/OneIntranet/_api/web/lists/getbytitle('RegisteredUsers')/items?$select=ID,Email,ScreenName,AreaOfInterest,Hobbies,About,SMEFor,NewJoiner,avatarID/ID&$expand=avatarID`;
      
      const usersResponse = await fetch(usersUrl, {
        method: 'GET',
        headers: {
          'Accept': 'application/json'
        }
      });

      if (!usersResponse.ok) {
        console.error('Error fetching users:', usersResponse.statusText);
        return;
      }

      const usersData = await usersResponse.json();
      const users = usersData.value || [];

      console.log('Fetched users:', users);

      // Fetch avatars for each user
      const usersWithAvatars = await Promise.all(
        users.map(async (user: any) => {
          let avatarUrl = '';
          try {
            if (user.avatarID) {
              const avatarId = user.avatarID?.ID || user.avatarID;
              avatarUrl = await this.getAvatarImageByAvatarId(avatarId);
            }
          } catch (err) {
            console.error(`Error fetching avatar for user ${user.ScreenName}:`, err);
          }
          return {
            ...user,
            avatarUrl: avatarUrl || require('./assets/MicrosoftLogo.png')
          };
        })
      );

      console.log('Users with avatars:', usersWithAvatars);

      // Update state with users and set the first user's avatar
      if (usersWithAvatars.length > 0) {
        this.setState({
          users: usersWithAvatars,
          currentUserIndex: 0,
          currentUserAvatarUrl: usersWithAvatars[0].avatarUrl
        } as ILbViewDbCardAdaptiveCardExtensionState);
      }
    } catch (error) {
      console.error('Error fetching users with avatars:', error);
    }
  }

  private async getAvatarImageByAvatarId(avatarId: number): Promise<string> {
    try {
      const url = `https://ygc8n.sharepoint.com/sites/OneIntranet/_api/web/lists/getbytitle('AvatarLibrary')/items(${avatarId})?$select=File/ServerRelativeUrl&$expand=File`;
      
      const response = await fetch(url, {
        method: 'GET',
        headers: {
          'Accept': 'application/json'
        }
      });

      if (response.ok) {
        const data = await response.json();
        if (data.File?.ServerRelativeUrl) {
          return 'https://ygc8n.sharepoint.com' + data.File.ServerRelativeUrl;
        }
      }
    } catch (error) {
      console.error('Error fetching avatar image:', error);
    }
    
    return '';
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'LbViewDbCard-property-pane'*/
      './LbViewDbCardPropertyPane'
    )
      .then(
        (component) => {
          this._deferredPropertyPane = new component.LbViewDbCardPropertyPane();
        }
      );
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane?.getPropertyPaneConfiguration();
  }
}
