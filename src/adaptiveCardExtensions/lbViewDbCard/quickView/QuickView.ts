import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView } from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'LbViewDbCardAdaptiveCardExtensionStrings';
import {
  ILbViewDbCardAdaptiveCardExtensionProps,
  ILbViewDbCardAdaptiveCardExtensionState
} from '../LbViewDbCardAdaptiveCardExtension';
import { createUserInteraction, checkUserInteractionExists, getMatchScoreFromUserInteractions } from '../spservice';

export interface IQuickViewData {
  screenName: string;
  areaOfInterest: string;
  hobbies: string;
  about: string;
  matchScore: number | null;
  joinButtonText: string;
  joinButtonEnabled: boolean;
}

export class QuickView extends BaseAdaptiveCardQuickView<
  ILbViewDbCardAdaptiveCardExtensionProps,
  ILbViewDbCardAdaptiveCardExtensionState,
  IQuickViewData
> {
  private matchScoreFetched: boolean = false;
  private matchScoreCache: number | null = null;

  public get data(): IQuickViewData {
    // Get current user from state
    const currentIndex = (this.state as any)?.currentUserIndex || 0;
    const users = (this.state as any)?.users || [];
    const currentUser = users.length > 0 ? users[currentIndex] : null;

    // Get first line of About
    const aboutText = currentUser?.About || '';
    const firstLineAbout = aboutText.split('\n')[0] || '';

    // Fetch match score if not already fetched
    if (!this.matchScoreFetched && currentUser) {
      this.matchScoreFetched = true;
      this.fetchAndCacheMatchScore(currentUser);
    }

    return {
      screenName: currentUser?.ScreenName || 'N/A',
      areaOfInterest: currentUser?.AreaOfInterest || 'N/A',
      hobbies: currentUser?.Hobbies || 'N/A',
      about: firstLineAbout,
      matchScore: this.matchScoreCache,
      joinButtonText: 'Join',
      joinButtonEnabled: true
    };
  }

  private async fetchAndCacheMatchScore(currentUser: any): Promise<void> {
    try {
      // Get logged-in user email from context
      const loggedInUserEmail = (this.context as any)?.pageContext?.user?.loginName || 
                               (this.context as any)?.pageContext?.user?.email;

      if (!loggedInUserEmail || !currentUser.Email) {
        console.log('Unable to determine user emails for match score fetch');
        return;
      }

      console.log('Fetching match score for:', loggedInUserEmail, '->', currentUser.Email);

      // Fetch match score from UserInteractions list
      const matchScore = await getMatchScoreFromUserInteractions(
        this.context,
        loggedInUserEmail,
        currentUser.Email
      );

      if (matchScore !== null) {
        console.log('Match score fetched:', matchScore);
        this.matchScoreCache = matchScore;
      } else {
        console.log('No existing match score found');
        this.matchScoreCache = null;
      }
    } catch (error) {
      console.error('Error fetching match score:', error);
      this.matchScoreCache = null;
    }
  }

  public get template(): ISPFxAdaptiveCard {
    return require('./template/QuickViewTemplate.json');
  }

  public async onAction(action: any): Promise<void> {
    if (action.type === 'Submit' && action.id === 'joinUser') {
      console.log('Join button clicked');
      
      // Get current user from state
      const currentIndex = (this.state as any)?.currentUserIndex || 0;
      const users = (this.state as any)?.users || [];
      const currentUser = users.length > 0 ? users[currentIndex] : null;

      if (!currentUser) {
        console.error('No current user found');
        alert('Error: User information not available');
        return;
      }

      console.log('Current user email:', currentUser.Email);
      
      // Get logged-in user email from context
      const loggedInUserEmail = (this.context as any)?.pageContext?.user?.loginName || 
                               (this.context as any)?.pageContext?.user?.email ||
                               'unknown';

      console.log('Logged-in user email:', loggedInUserEmail);

      if (!loggedInUserEmail || loggedInUserEmail === 'unknown') {
        console.error('Unable to determine logged-in user email');
        alert('Error: Unable to determine your email');
        return;
      }

      try {
        // Check if interaction already exists
        const interactionExists = await checkUserInteractionExists(
          this.context,
          loggedInUserEmail,
          currentUser.Email,
          'requested'
        );

        if (interactionExists) {
          console.log('Join request already sent');
          alert('You have already sent a join request to this user');
        } else {
          // Create user interaction
          const success = await createUserInteraction(
            this.context,
            loggedInUserEmail,
            currentUser.Email,
            'requested'
          );

          if (success) {
            alert('Join request sent successfully');
            console.log('User interaction created:', loggedInUserEmail, '->', currentUser.Email);
          } else {
            alert('Failed to send join request');
          }
        }
      } catch (error) {
        console.error('Error handling join request:', error);
        alert('Error sending join request');
      }
    }
  }
}
