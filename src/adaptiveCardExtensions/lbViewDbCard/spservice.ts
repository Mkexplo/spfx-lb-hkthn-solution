import { WebPartContext } from "@microsoft/sp-webpart-base";

export const createUserInteraction = async (
  context: any,
  requesterEmail: string,
  recipientEmail: string,
  status: string = 'requested'
): Promise<boolean> => {
  try {
    const url = `https://ygc8n.sharepoint.com/sites/OneIntranet/_api/web/lists/getbytitle('UserInteractions')/items`;
    
    const body = JSON.stringify({
      Requester: requesterEmail,
      Recipient: recipientEmail,
      Status: status
    });

    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Accept': 'application/json'
      },
      body: body
    });

    if (response.ok) {
      console.log('User interaction created successfully');
      return true;
    } else {
      console.error('Failed to create user interaction:', response.status);
      return false;
    }
  } catch (error) {
    console.error('Error creating user interaction:', error);
    return false;
  }
};

export const checkUserInteractionExists = async (
  context: any,
  requesterEmail: string,
  recipientEmail: string,
  status: string = 'requested'
): Promise<boolean> => {
  try {
    const filterQuery = `Requester eq '${requesterEmail}' and Recipient eq '${recipientEmail}' and Status eq '${status}'`;
    const encodedFilter = encodeURIComponent(filterQuery);
    const url = `https://ygc8n.sharepoint.com/sites/OneIntranet/_api/web/lists/getbytitle('UserInteractions')/items?$filter=${encodedFilter}`;

    const response = await fetch(url, {
      method: 'GET',
      headers: {
        'Accept': 'application/json'
      }
    });

    if (response.ok) {
      const data = await response.json();
      if (data.value && data.value.length > 0) {
        console.log('User interaction found:', data.value[0]);
        return true;
      }
    } else {
      console.error('Error checking user interaction:', response.status);
    }
    
    return false;
  } catch (error) {
    console.error('Error checking user interaction:', error);
    return false;
  }
};

export const getMatchScoreFromUserInteractions = async (
  context: any,
  loggedInUserEmail: string,
  clickedUserEmail: string
): Promise<number | null> => {
  try {
    const filterQuery = `Requester eq '${loggedInUserEmail.trim()}' and Recipient eq '${clickedUserEmail.trim()}'`;
    const encodedFilter = encodeURIComponent(filterQuery);
    const url = `https://ygc8n.sharepoint.com/sites/OneIntranet/_api/web/lists/getbytitle('UserInteractions')/items?$filter=${encodedFilter}&$select=ID,MatchScore`;

    const response = await fetch(url, {
      method: 'GET',
      headers: {
        'Accept': 'application/json'
      }
    });

    if (response.ok) {
      const data = await response.json();
      if (data.value && data.value.length > 0 && data.value[0].MatchScore !== null && data.value[0].MatchScore !== undefined) {
        console.log('Match score retrieved:', data.value[0].MatchScore);
        return data.value[0].MatchScore;
      }
    } else {
      console.error('Error retrieving match score:', response.status);
    }
    
    console.log('No match score found for this interaction');
    return null;
  } catch (error) {
    console.error('Error retrieving match score from UserInteractions:', error);
    return null;
  }
};
