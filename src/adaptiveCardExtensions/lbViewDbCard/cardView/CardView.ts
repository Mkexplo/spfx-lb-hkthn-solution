import {
  BaseImageCardView,
  IImageCardParameters,
  IExternalLinkCardAction,
  IQuickViewCardAction,
  ICardButton
} from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'LbViewDbCardAdaptiveCardExtensionStrings';
import {
  ILbViewDbCardAdaptiveCardExtensionProps,
  ILbViewDbCardAdaptiveCardExtensionState,
  QUICK_VIEW_REGISTRY_ID
} from '../LbViewDbCardAdaptiveCardExtension';

export class CardView extends BaseImageCardView<
  ILbViewDbCardAdaptiveCardExtensionProps,
  ILbViewDbCardAdaptiveCardExtensionState
> {
  /**
   * Buttons will not be visible if card size is 'Medium' with Image Card View.
   * It will support up to two buttons for 'Large' card size.
   */
  public get cardButtons(): [ICardButton] | [ICardButton, ICardButton] | undefined {
    return [
      {
        title: 'Next',
        action: {
          type: 'Submit',
          parameters: {
            id: 'nextUser'
          }
        }
      },
      {
        title: strings.QuickViewButton,
        action: {
          type: 'QuickView',
          parameters: {
            view: QUICK_VIEW_REGISTRY_ID
          }
        }
      }
    ];
  }

  public onAction(action: any): void {
    console.log('onAction triggered:', action);
    
    // Handle card click - navigate to LunchConnect page
    if (action.type === 'QuickView' && action.data?.view === QUICK_VIEW_REGISTRY_ID) {
      console.log('Card clicked - navigating to LunchConnect page');
      window.location.href = 'https://ygc8n.sharepoint.com/sites/OneIntranet/SitePages/LunchConnect.aspx';
      return;
    }
    
    if (action.type === 'Submit' && action.data.id === 'nextUser') {
      console.log('Next button clicked - navigating to next user');
      
      // Navigate to next user
      const currentIndex = (this.state as any)?.currentUserIndex || 0;
      const users = (this.state as any)?.users || [];
      
      console.log('Current index:', currentIndex);
      console.log('Total users:', users.length);
      
      if (users.length === 0) {
        console.warn('No users available');
        return;
      }

      // Calculate next index (circular)
      const nextIndex = (currentIndex + 1) % users.length;
      
      console.log('Moving to next index:', nextIndex);
      
      // Update state with next user index
      this.setState({
        currentUserIndex: nextIndex
      } as ILbViewDbCardAdaptiveCardExtensionState);
    }
  }

  public get data(): IImageCardParameters {
    // Get current user index from state, default to 0
    const currentIndex = (this.state as any)?.currentUserIndex || 0;
    let users = (this.state as any)?.users || [];
    
    // If users is empty, try to fetch from state or use empty array
    if (users.length === 0) {
      console.warn('Users array is empty in state');
      users = [];
    }
    
    // Get current user
    const currentUser = users.length > 0 ? users[currentIndex] : null;
    
    // Get avatar URL from user object (already fetched)
    const avatarUrl = currentUser?.avatarUrl || require('../assets/MicrosoftLogo.png');
    const screenName = currentUser?.ScreenName || 'Employee Profile';

    console.log('Current user:', currentUser);
    console.log('Avatar URL:', avatarUrl);
    console.log('Total users:', users.length);

    return {
      primaryText: screenName,
      imageUrl: avatarUrl,
      title: this.properties.title
    };
  }

  public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
    return undefined;
  }
}
