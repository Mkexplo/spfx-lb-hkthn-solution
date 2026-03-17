import { SPHttpClient } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IAvatarImage {
  avatarID: number;
  url: string;
}

export const getUserImages = async (context: WebPartContext): Promise<IAvatarImage[]> => {
  // Hardcoded test images - replace with actual API call once verified
  /*
  const hardcodedImages: IAvatarImage[] = [
    {
      avatarID: 1,
      url: 'https://via.placeholder.com/150x100?text=Avatar1'
    },
    {
      avatarID: 2,
      url: 'https://via.placeholder.com/150x100?text=Avatar2'
    },
    {
      avatarID: 3,
      url: 'https://via.placeholder.com/150x100?text=Avatar3'
    }
  ];
  console.log('Using hardcoded images:', hardcodedImages);
  return hardcodedImages;
*/
  //Original API call - uncomment when list is ready
  const url = `https://ygc8n.sharepoint.com/sites/OneIntranet/_api/web/lists/getbytitle('AvatarLibrary')/items?$select=ID,File/ServerRelativeUrl&$expand=File`;
  const response = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
  const data = await response.json();
  console.log('Avatar API Response:', data);
  if (data.value && data.value.length > 0) {
    return data.value.map((item: any) => {
      const imageUrl = item.File?.ServerRelativeUrl ? 'https://ygc8n.sharepoint.com' + item.File.ServerRelativeUrl : '';
      return {
        avatarID: item.ID,
        url: imageUrl
      };
    });
  }
  return [];
  
};

export const getUserRecord = async (context: WebPartContext, userEmail: string): Promise<any> => {
  const url = `https://ygc8n.sharepoint.com/sites/OneIntranet/_api/web/lists/getbytitle('RegisteredUsers')/items?$filter=Email eq '${userEmail}'`;
  const response = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
  const data = await response.json();
  if (data.value && data.value.length > 0) {
    return data.value[0];
  }
  return null;
};

export const updateUserAvatarId = async (context: WebPartContext, userEmail: string, avatarId: number): Promise<boolean> => {
  try {
    const userRecord = await getUserRecord(context, userEmail);
    if (!userRecord) {
      console.error('User record not found');
      return false;
    }
    
    const updateUrl = `https://ygc8n.sharepoint.com/sites/OneIntranet/_api/web/lists/getbytitle('RegisteredUsers')/items(${userRecord.ID})`;
    const body = JSON.stringify({
      avatarIDId: avatarId
    });
    
    console.log('Updating avatar with URL:', updateUrl);
    console.log('Body:', body);
    
    const response = await context.spHttpClient.post(updateUrl, SPHttpClient.configurations.v1, {
      headers: {
        'Content-Type': 'application/json',
        'X-HTTP-Method': 'MERGE',
        'IF-MATCH': '*'
      },
      body: body
    });
    
    if (response.ok) {
      console.log('Avatar ID updated successfully');
      return true;
    } else {
      const errorText = await response.text();
      console.error('Failed to update avatar ID:', response.status, errorText);
      return false;
    }
  } catch (error) {
    console.error('Error updating avatar ID:', error);
    return false;
  }
};

export const updateUserProfile = async (context: WebPartContext, userEmail: string, profileData: any): Promise<boolean> => {
  try {
    const userRecord = await getUserRecord(context, userEmail);
    if (!userRecord) {
      console.error('User record not found');
      return false;
    }

    // Check for profanity in ScreenName and About before updating
    if (profileData.ScreenName) {
      const hasScreenNameProfanity = await checkForProfanity(profileData.ScreenName, context);
      if (hasScreenNameProfanity) {
        alert('Profanity detected in Screen Name');
        return false;
      }
    }

    if (profileData.About) {
      const hasAboutProfanity = await checkForProfanity(profileData.About, context);
      if (hasAboutProfanity) {
        alert('Profanity detected in About section');
        return false;
      }
    }
    
    const updateUrl = `https://ygc8n.sharepoint.com/sites/OneIntranet/_api/web/lists/getbytitle('RegisteredUsers')/items(${userRecord.ID})`;
    const body = JSON.stringify({
      ScreenName: profileData.ScreenName,
      About: profileData.About,
      AreaOfInterest: profileData.AreaOfInterest,
      Hobbies: profileData.Hobbies,
      Communities: profileData.Communities
    });
    
    console.log('Updating user profile with body:', body);
    
    const response = await context.spHttpClient.post(updateUrl, SPHttpClient.configurations.v1, {
      headers: {
        'Content-Type': 'application/json',
        'X-HTTP-Method': 'MERGE',
        'IF-MATCH': '*'
      },
      body: body
    });
    
    if (response.ok) {
      console.log('User profile updated successfully');
      return true;
    } else {
      console.error('Failed to update user profile:', response.status);
      return false;
    }
  } catch (error) {
    console.error('Error updating user profile:', error);
    return false;
  }
};export const getAreaOfInterestList = async (context: WebPartContext): Promise<any[]> => {
  try {
    const url = `https://ygc8n.sharepoint.com/sites/OneIntranet/_api/web/lists/getbytitle('ListAreaOfInterest')/items?$select=ID,Title`;
    const response = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    const data = await response.json();
    console.log('Area of Interest List Response:', data);
    if (data.value && data.value.length > 0) {
      return data.value.map((item: any) => ({
        key: item.ID,
        text: item.Title
      }));
    }
    return [];
  } catch (error) {
    console.error('Error fetching area of interest list:', error);
    return [];
  }
};

export const getHobbiesList = async (context: WebPartContext): Promise<any[]> => {
  try {
    const url = `https://ygc8n.sharepoint.com/sites/OneIntranet/_api/web/lists/getbytitle('ListHobbies')/items?$select=ID,Title`;
    const response = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    const data = await response.json();
    console.log('Hobbies List Response:', data);
    if (data.value && data.value.length > 0) {
      return data.value.map((item: any) => ({
        key: item.ID,
        text: item.Title
      }));
    }
    return [];
  } catch (error) {
    console.error('Error fetching hobbies list:', error);
    return [];
  }
};

export const getUsersByAreaOfInterest = async (context: WebPartContext, areaOfInterest: string, loggedInUserEmail?: string): Promise<any[]> => {
  try {
    const url = `https://ygc8n.sharepoint.com/sites/OneIntranet/_api/web/lists/getbytitle('RegisteredUsers')/items?$select=ID,Email,ScreenName,AreaOfInterest,Hobbies,About,SMEFor,NewJoiner,avatarID/ID&$expand=avatarID`;
    const response = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    const data = await response.json();
    console.log('Users list response:', data);
    
    if (data.value && data.value.length > 0) {
      // Filter users whose AreaOfInterest contains the selected area
      const filteredUsers = data.value.filter((user: any) => {
        // Exclude logged-in user if email is provided
        if (loggedInUserEmail && user.Email === loggedInUserEmail) {
          return false;
        }
        if (!user.AreaOfInterest) return false;
        const interests = user.AreaOfInterest.split('|').map((item: string) => item.trim());
        return interests.includes(areaOfInterest.trim());
      });
      return filteredUsers;
    }
    return [];
  } catch (error) {
    console.error('Error fetching users by area of interest:', error);
    return [];
  }
};

export const getAvatarImageByAvatarId = async (context: WebPartContext, avatarId: number): Promise<string> => {
  try {
    const url = `https://ygc8n.sharepoint.com/sites/OneIntranet/_api/web/lists/getbytitle('AvatarLibrary')/items(${avatarId})?$select=File/ServerRelativeUrl&$expand=File`;
    const response = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    const data = await response.json();
    
    if (data.File?.ServerRelativeUrl) {
      return 'https://ygc8n.sharepoint.com' + data.File.ServerRelativeUrl;
    }
    return '';
  } catch (error) {
    console.error('Error fetching avatar image:', error);
    return '';
  }
};

export const getUsersByHobbies = async (context: WebPartContext, hobby: string, loggedInUserEmail?: string): Promise<any[]> => {
  try {
    const url = `https://ygc8n.sharepoint.com/sites/OneIntranet/_api/web/lists/getbytitle('RegisteredUsers')/items?$select=ID,Email,ScreenName,AreaOfInterest,Hobbies,About,SMEFor,NewJoiner,avatarID/ID&$expand=avatarID`;
    const response = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    const data = await response.json();
    console.log('Users list response:', data);
    
    if (data.value && data.value.length > 0) {
      // Filter users whose Hobbies contains the selected hobby
      const filteredUsers = data.value.filter((user: any) => {
        // Exclude logged-in user if email is provided
        if (loggedInUserEmail && user.Email === loggedInUserEmail) {
          return false;
        }
        if (!user.Hobbies) return false;
        const hobbies = user.Hobbies.split('|').map((item: string) => item.trim());
        return hobbies.includes(hobby.trim());
      });
      return filteredUsers;
    }
    return [];
  } catch (error) {
    console.error('Error fetching users by hobbies:', error);
    return [];
  }
};

export const createUserInteraction = async (
  context: WebPartContext,
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

    const response = await context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        'Content-Type': 'application/json'
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
  context: WebPartContext,
  requesterEmail: string,
  recipientEmail: string,
  status: string = 'requested'
): Promise<boolean> => {
  try {
    const filterQuery = `Requester eq '${requesterEmail}' and Recipient eq '${recipientEmail}' and Status eq '${status}'`;
    const encodedFilter = encodeURIComponent(filterQuery);
    const url = `https://ygc8n.sharepoint.com/sites/OneIntranet/_api/web/lists/getbytitle('UserInteractions')/items?$filter=${encodedFilter}`;

    const response = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    const data = await response.json();

    if (data.value && data.value.length > 0) {
      console.log('User interaction found:', data.value[0]);
      return true;
    }
    return false;
  } catch (error) {
    console.error('Error checking user interaction:', error);
    return false;
  }
};

export const checkUserInteractionExistsAnyStatus = async (
  context: WebPartContext,
  requesterEmail: string,
  recipientEmail: string
): Promise<number> => {
  try {
    const filterQuery = `Requester eq '${requesterEmail}' and Recipient eq '${recipientEmail}'`;
    const encodedFilter = encodeURIComponent(filterQuery);
    const url = `https://ygc8n.sharepoint.com/sites/OneIntranet/_api/web/lists/getbytitle('UserInteractions')/items?$filter=${encodedFilter}`;

    const response = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    const data = await response.json();

    if (data.value && data.value.length > 0) {
      console.log('User interaction found with any status:', data.value[0]);
      return data.value[0].ID;
    }
    return 0;
  } catch (error) {
    console.error('Error checking user interaction:', error);
    return 0;
  }
};

export const getMatchScoreFromUserInteractions = async (
  context: WebPartContext,
  loggedInUserEmail: string,
  clickedUserEmail: string
): Promise<number | null> => {
  try {
    const filterQuery = `Requester eq '${loggedInUserEmail.trim()}' and Recipient eq '${clickedUserEmail.trim()}'`;
    const encodedFilter = encodeURIComponent(filterQuery);
    const url = `https://ygc8n.sharepoint.com/sites/OneIntranet/_api/web/lists/getbytitle('UserInteractions')/items?$filter=${encodedFilter}&$select=ID,MatchScore`;

    const response = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    const data = await response.json();

    if (data.value && data.value.length > 0 && data.value[0].MatchScore !== null && data.value[0].MatchScore !== undefined) {
      console.log('Match score retrieved:', data.value[0].MatchScore);
      return data.value[0].MatchScore;
    }
    
    console.log('No match score found for this interaction');
    return null;
  } catch (error) {
    console.error('Error retrieving match score from UserInteractions:', error);
    return null;
  }
};

export const updateUserInteractionStatus = async (
  context: WebPartContext,
  requesterEmail: string,
  recipientEmail: string,
  newStatus: string
): Promise<boolean> => {
  try {
    // First, get the item ID
    const filterQuery = `Requester eq '${requesterEmail}' and Recipient eq '${recipientEmail}' and Status eq 'requested'`;
    const encodedFilter = encodeURIComponent(filterQuery);
    const getUrl = `https://ygc8n.sharepoint.com/sites/OneIntranet/_api/web/lists/getbytitle('UserInteractions')/items?$filter=${encodedFilter}`;

    const getResponse = await context.spHttpClient.get(getUrl, SPHttpClient.configurations.v1);
    const getData = await getResponse.json();

    if (!getData.value || getData.value.length === 0) {
      console.error('No interaction record found to update');
      return false;
    }

    const itemId = getData.value[0].ID;

    // Update the item
    const updateUrl = `https://ygc8n.sharepoint.com/sites/OneIntranet/_api/web/lists/getbytitle('UserInteractions')/items(${itemId})`;
    const body = JSON.stringify({
      Status: newStatus
    });

    const updateResponse = await context.spHttpClient.post(updateUrl, SPHttpClient.configurations.v1, {
      headers: {
        'Content-Type': 'application/json',
        'X-HTTP-Method': 'MERGE',
        'If-Match': '*'
      },
      body: body
    });

    if (updateResponse.ok) {
      console.log('User interaction status updated successfully');
      
      // If status is being set to 'Matched', update both users' status in RegisteredUsers list to 'Closed'
      if (newStatus.toLowerCase() === 'matched') {
        try {
          // Update requester's status to 'Closed'
          const requesterRecord = await getUserRecord(context, requesterEmail);
          if (requesterRecord) {
            const requesterUpdateUrl = `https://ygc8n.sharepoint.com/sites/OneIntranet/_api/web/lists/getbytitle('RegisteredUsers')/items(${requesterRecord.ID})`;
            const requesterBody = JSON.stringify({
              Status: 'Closed'
            });
            
            await context.spHttpClient.post(requesterUpdateUrl, SPHttpClient.configurations.v1, {
              headers: {
                'Content-Type': 'application/json',
                'X-HTTP-Method': 'MERGE',
                'If-Match': '*'
              },
              body: requesterBody
            });
            console.log('Requester status updated to Closed');
          }

          // Update recipient's status to 'Closed'
          const recipientRecord = await getUserRecord(context, recipientEmail);
          if (recipientRecord) {
            const recipientUpdateUrl = `https://ygc8n.sharepoint.com/sites/OneIntranet/_api/web/lists/getbytitle('RegisteredUsers')/items(${recipientRecord.ID})`;
            const recipientBody = JSON.stringify({
              Status: 'Closed'
            });
            
            await context.spHttpClient.post(recipientUpdateUrl, SPHttpClient.configurations.v1, {
              headers: {
                'Content-Type': 'application/json',
                'X-HTTP-Method': 'MERGE',
                'If-Match': '*'
              },
              body: recipientBody
            });
            console.log('Recipient status updated to Closed');
          }
        } catch (error) {
          console.error('Error updating user status in RegisteredUsers list:', error);
        }
      }
      
      return true;
    } else {
      console.error('Failed to update user interaction:', updateResponse.status);
      return false;
    }
  } catch (error) {
    console.error('Error updating user interaction:', error);
    return false;
  }
};

export const checkReceivedInteractionExists = async (
  context: WebPartContext,
  requesterEmail: string,
  recipientEmail: string,
  status: string = 'requested'
): Promise<boolean> => {
  try {
    const filterQuery = `Requester eq '${requesterEmail}' and Recipient eq '${recipientEmail}' and Status eq '${status}'`;
    const encodedFilter = encodeURIComponent(filterQuery);
    const url = `https://ygc8n.sharepoint.com/sites/OneIntranet/_api/web/lists/getbytitle('UserInteractions')/items?$filter=${encodedFilter}`;

    const response = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    const data = await response.json();

    if (data.value && data.value.length > 0) {
      console.log('Received user interaction found:', data.value[0]);
      return true;
    }
    return false;
  } catch (error) {
    console.error('Error checking received user interaction:', error);
    return false;
  }
};

export const checkMatchedInteractionExists = async (
  context: WebPartContext,
  userEmail1: string,
  userEmail2: string
): Promise<boolean> => {
  try {
    // Check if there's a matched interaction between the two users (in either direction)
    const filterQuery = `((Requester eq '${userEmail1}' and Recipient eq '${userEmail2}') or (Requester eq '${userEmail2}' and Recipient eq '${userEmail1}')) and Status eq 'Matched'`;
    const encodedFilter = encodeURIComponent(filterQuery);
    const url = `https://ygc8n.sharepoint.com/sites/OneIntranet/_api/web/lists/getbytitle('UserInteractions')/items?$filter=${encodedFilter}`;

    const response = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    const data = await response.json();

    if (data.value && data.value.length > 0) {
      console.log('Matched user interaction found:', data.value[0]);
      return true;
    }
    return false;
  } catch (error) {
    console.error('Error checking matched user interaction:', error);
    return false;
  }
};

export const getCommunities = async (context: WebPartContext, userEmail?: string): Promise<any[]> => {
  try {
    // Get the AAD token provider
    const tokenProvider = await context.aadTokenProviderFactory.getTokenProvider();
    
    // Get the token for Yammer API
    const token = await tokenProvider.getToken('https://api.yammer.com');
    
    console.log('Token obtained for Yammer API:', !!token);
    
    if (!token) {
      console.error('Unable to get AAD token for Yammer API');
      return [];
    }

    let userId = null;
    
    // If userEmail is provided, get the user ID from Yammer
    if (userEmail) {
      console.log('Fetching user ID for email:', userEmail);
      
      const userLookupUrl = `https://api.yammer.com/api/v1/users/by_email.json?email=${encodeURIComponent(userEmail)}`;
      
      const userResponse = await fetch(userLookupUrl, {
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${token}`,
          'Accept': 'application/json'
        }
      });

      console.log('User lookup response status:', userResponse.status);

      if (userResponse.ok) {
        const userData = await userResponse.json();
        console.log('User data from Yammer:', userData);
        
        // Handle both array and single object responses
        if (Array.isArray(userData) && userData.length > 0) {
          userId = userData[0].id;
        } else if (userData && userData.id) {
          userId = userData.id;
        }
        
        console.log('User ID extracted:', userId);
      } else {
        const errorText = await userResponse.text();
        console.error('Error fetching user from Yammer:', userResponse.statusText, errorText);
        return [];
      }
    }

    if (!userId) {
      console.error('Unable to determine user ID');
      return [];
    }

    // Fetch groups for this user
    console.log('Fetching groups for user ID:', userId);
    const groupsUrl = `https://api.yammer.com/api/v1/users/${userId}/groups.json`;
    
    const groupsResponse = await fetch(groupsUrl, {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Accept': 'application/json'
      }
    });

    console.log('Groups response status:', groupsResponse.status);

    if (!groupsResponse.ok) {
      const errorText = await groupsResponse.text();
      console.error('Error fetching groups from Yammer:', groupsResponse.statusText, errorText);
      return [];
    }

    const groupsData = await groupsResponse.json();
    console.log('Groups data from Yammer:', groupsData);
    console.log('Groups data type:', typeof groupsData);
    console.log('Is array?', Array.isArray(groupsData));

    let groupIds: any[] = [];
    
    if (Array.isArray(groupsData)) {
      groupIds = groupsData;
    } else if (groupsData && Array.isArray(groupsData.groups)) {
      groupIds = groupsData.groups;
    } else {
      console.warn('Unexpected groups response structure:', groupsData);
      return [];
    }

    console.log('Group IDs found:', groupIds);

    if (!groupIds || groupIds.length === 0) {
      console.log('User is not a member of any groups');
      return [];
    }

    // Fetch full group details to get names
    const communities: any[] = [];
    
    for (const groupId of groupIds) {
      console.log('Fetching details for group ID:', groupId);
      
      const groupDetailUrl = `https://api.yammer.com/api/v1/groups/${groupId}.json`;
      
      const groupDetailResponse = await fetch(groupDetailUrl, {
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${token}`,
          'Accept': 'application/json'
        }
      });

      if (groupDetailResponse.ok) {
        const groupDetail = await groupDetailResponse.json();
        console.log('Group detail:', groupDetail);
        
        if (groupDetail && groupDetail.id && groupDetail.name) {
          communities.push({
            key: groupDetail.id.toString(),
            text: groupDetail.name
          });
        }
      } else {
        console.warn(`Failed to fetch group details for ID ${groupId}`);
      }
    }

    console.log('Final communities:', communities);
    return communities;

  } catch (error) {
    console.error('Error fetching communities:', error);
    return [];
  }
};

export const checkForProfanity = async (text: string, context: WebPartContext): Promise<boolean> => {
  try {
    if (!text || text.trim() === '') {
      return false; // No profanity if text is empty
    }

    const config = await getAzureOpenAiConfig(context);

    if (!config || !config.endpoint || !config.apiKey) {
      console.warn('Azure OpenAI config not available for profanity check, allowing content');
      return false; // Allow content if config is not available
    }

    let endpoint = config.endpoint.trim();
    const apiKey = config.apiKey.trim();
    const apiVersion = (config.apiVersion || "2025-01-01-preview").trim();
    const deployment = (config.deployment || "o4-mini").trim();

    if (endpoint.endsWith('/')) {
      endpoint = endpoint.slice(0, -1);
    }

    const profanityCheckPrompt = `Check if the following text contains any profane, offensive, or inappropriate language. Respond with ONLY "YES" if it contains profanity or "NO" if it doesn't. Text: "${text}"`;

    const url = `${endpoint}/openai/deployments/${deployment}/chat/completions?api-version=${apiVersion}`;
    
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'api-key': apiKey
      },
      body: JSON.stringify({
        model: deployment,
        messages: [
          { role: "user", content: profanityCheckPrompt }
        ],
        max_completion_tokens: 4000
      })
    });

    if (!response.ok) {
      console.warn('Profanity check API error, allowing content');
      return false;
    }

    const result = await response.json();
    
    if (result.choices && result.choices.length > 0 && result.choices[0].message) {
      const hasProfanity = result.choices[0].message.content.trim().toUpperCase().includes('YES');
      console.log('Profanity check result:', hasProfanity ? 'PROFANITY DETECTED' : 'OK');
      return hasProfanity;
    }
    
    return false;
  } catch (error) {
    console.error('Error checking for profanity:', error);
    return false; // Allow content if there's an error
  }
};

export const getAzureOpenAiConfig = async (context: WebPartContext): Promise<any> => {
  try {
    const url = `https://ygc8n.sharepoint.com/sites/OneIntranet/_api/web/lists/getbytitle('AppConfigList')/items?$select=Key,Value`;
    const response = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    const data = await response.json();
    
    const config: any = {
      endpoint: '',
      apiKey: '',
      apiVersion: '2025-01-01-preview',
      deployment: 'o4-mini'
    };

    if (data.value && Array.isArray(data.value)) {
      data.value.forEach((item: any) => {
        if (item.Key === 'AZURE_OPENAI_ENDPOINT') {
          config.endpoint = item.Value;
        } else if (item.Key === 'AZURE_OPENAI_API_KEY') {
          config.apiKey = item.Value;
        } else if (item.Key === 'AZURE_OPENAI_API_VERSION') {
          config.apiVersion = item.Value;
        } else if (item.Key === 'AZURE_OPENAI_DEPLOYMENT') {
          config.deployment = item.Value;
        }
      });
    }

    return config;
  } catch (error) {
    console.error('Error fetching Azure OpenAI config from AppConfigList:', error);
    return null;
  }
};

export const getAIO4miniResponse = async (userMessage: string, context: WebPartContext): Promise<string> => {
  try {
    // Fetch configuration from AppConfigList
    const config = await getAzureOpenAiConfig(context);

    // Validate that required config is set
    if (!config || !config.endpoint || !config.apiKey) {
      console.error('Azure OpenAI credentials are not configured in AppConfigList.');
      return 'Configuration error: Azure OpenAI credentials not found.';
    }

    let endpoint = config.endpoint.trim();
    const apiKey = config.apiKey.trim();
    const apiVersion = (config.apiVersion || "2025-01-01-preview").trim();
    const deployment = (config.deployment || "o4-mini").trim();

    // Ensure endpoint doesn't have trailing slash
    if (endpoint.endsWith('/')) {
      endpoint = endpoint.slice(0, -1);
    }

    // Use native fetch to call Azure OpenAI REST API
    const url = `${endpoint}/openai/deployments/${deployment}/chat/completions?api-version=${apiVersion}`;
    
    console.log('Azure OpenAI Request:', { url, endpoint, deployment, apiVersion });
    
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'api-key': apiKey
      },
      body: JSON.stringify({
        model: deployment,
        messages: [
          { role: "user", content: userMessage }
        ],
        max_completion_tokens: 8000
      })
    });

    if (!response.ok) {
      const errorData = await response.text();
      console.error(`Azure OpenAI API error: ${response.status}`, { url, status: response.status, statusText: response.statusText, errorData });
      return `Error: API returned ${response.status} - ${response.statusText}`;
    }

    const result = await response.json();
    console.log('AI Response received successfully:', JSON.stringify(result));
    
    if (result.choices && result.choices.length > 0) {
      const choice = result.choices[0];
      const message = choice.message;
      console.log('Finish reason:', choice.finish_reason);
      console.log('Message object:', JSON.stringify(message));
      
      if (message && message.content) {
        const content = message.content.trim();
        console.log('Extracted content:', content);
        
        if (!content) {
          console.warn('Warning: Message content is empty after trim');
          return 'Error: Received empty response from AI';
        }
        
        return content;
      } else {
        console.error('No message.content found in response', { message });
        return 'Error: No content in AI response';
      }
    } else {
      console.error('No choices found in response', { result });
      return 'Error: No choices in AI response';
    }
  } catch (error) {
    console.error('Error getting AI response:', error);
    return 'Error: Unable to get AI response. Please check your configuration.';
  }
};

export const getScreenNameAIResponse = async (userIntro: string, context: WebPartContext): Promise<string> => {
  try {
    // Fetch configuration from AppConfigList
    const config = await getAzureOpenAiConfig(context);

    if (!config || !config.endpoint || !config.apiKey) {
      console.error('Azure OpenAI credentials are not configured in AppConfigList');
      return '';
    }

    let endpoint = config.endpoint.trim();
    const apiKey = config.apiKey.trim();
    const apiVersion = (config.apiVersion || "2025-01-01-preview").trim();
    const deployment = (config.deployment || "o4-mini").trim();

    // Ensure endpoint doesn't have trailing slash
    if (endpoint.endsWith('/')) {
      endpoint = endpoint.slice(0, -1);
    }

    const aiPrompt = "Generate a creative and professional screen name (nickname) for a business professional. The name should be:\n1. Professional yet approachable\n2. Memorable and unique\n3. Between 25 letters and no space\n4. No special characters except hyphens or underscores\n5. Create name based on the introduction text of user - " + userIntro + " \n6. Return only single best screen name, nothing else";

    // Use native fetch to call Azure OpenAI REST API
    const url = `${endpoint}/openai/deployments/${deployment}/chat/completions?api-version=${apiVersion}`;
    
    console.log('Azure OpenAI Request for Screen Name:', { url, endpoint, deployment, apiVersion });
    
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'api-key': apiKey
      },
      body: JSON.stringify({
        model: deployment,
        messages: [
          { role: "system", content: "You are a creative naming assistant that generates professional screen names." },
          { role: "user", content: aiPrompt }
        ],
        max_completion_tokens: 8000
      })
    });

    if (!response.ok) {
      const errorData = await response.text();
      console.error(`Azure OpenAI API error: ${response.status}`, { url, status: response.status, statusText: response.statusText, errorData });
      return '';
    }

    const result = await response.json();
    console.log('Screen Name AI Response received successfully:', JSON.stringify(result));

    if (result.choices && result.choices.length > 0) {
      const choice = result.choices[0];
      const message = choice.message;
      console.log('Screen Name finish reason:', choice.finish_reason);
      console.log('Screen Name Message object:', JSON.stringify(message));
      
      if (message && message.content) {
        const screenName = message.content.trim().replace(/^["']|["']$/g, '');
        console.log('Generated screen name:', screenName);
        
        if (!screenName) {
          console.warn('Warning: Generated screen name is empty after trim');
          return '';
        }
        
        return screenName;
      } else {
        console.error('No message.content found in screen name response', { message });
        return '';
      }
    } else {
      console.error('No choices found in screen name response', { result });
      return '';
    }
  } catch (error) {
    console.error('Error generating screen name:', error);
    return '';
  }
};

export const calculateMatchScore = async (loggedInUser: any, clickedUser: any, context: WebPartContext): Promise<number> => {
  try {
    console.log('=== Starting calculateMatchScore ===');
    console.log('Logged-in user:', loggedInUser);
    console.log('Clicked user:', clickedUser);
    
    // Fetch Yammer communities for both users
    console.log('Fetching communities for logged-in user...');
    const loggedInUserCommunities = await getCommunities(context, loggedInUser.Email);
    
    console.log('Fetching communities for clicked user...');
    const clickedUserCommunities = await getCommunities(context, clickedUser.Email);
    
    console.log('Logged-in user communities count:', loggedInUserCommunities.length);
    console.log('Logged-in user communities:', loggedInUserCommunities);
    console.log('Clicked user communities count:', clickedUserCommunities.length);
    console.log('Clicked user communities:', clickedUserCommunities);

    // Find shared communities
    const sharedCommunities = loggedInUserCommunities.filter((community: any) => 
      clickedUserCommunities.some((c: any) => c.key === community.key)
    );

    console.log('Shared communities count:', sharedCommunities.length);
    console.log('Shared communities:', sharedCommunities);

    // Get configuration from AppConfigList
    const configUrl = `https://ygc8n.sharepoint.com/sites/OneIntranet/_api/web/lists/getbytitle('AppConfigList')/items?$select=Key,Value`;
    const configResponse = await context.spHttpClient.get(configUrl, SPHttpClient.configurations.v1);
    const configData = await configResponse.json();

    let endpoint = '';
    let apiKey = '';
    let apiVersion = '2025-01-01-preview';
    let deployment = 'o4-mini';

    if (configData.value && Array.isArray(configData.value)) {
      configData.value.forEach((item: any) => {
        switch (item.Key) {
          case 'AZURE_OPENAI_ENDPOINT':
            endpoint = item.Value;
            break;
          case 'AZURE_OPENAI_API_KEY':
            apiKey = item.Value;
            break;
          case 'AZURE_OPENAI_API_VERSION':
            apiVersion = item.Value;
            break;
          case 'AZURE_OPENAI_DEPLOYMENT':
            deployment = item.Value;
            break;
        }
      });
    }

    if (!endpoint || !apiKey) {
      console.error('Azure OpenAI configuration not found');
      return 0;
    }

    // Clean up endpoint URL
    if (endpoint.charAt(endpoint.length - 1) === '/') {
      endpoint = endpoint.slice(0, -1);
    }

    // Format communities as comma-separated names for the AI prompt
    const loggedInCommunitiesText = loggedInUserCommunities.map((c: any) => c.text).join(', ') || 'None';
    const clickedCommunitiesText = clickedUserCommunities.map((c: any) => c.text).join(', ') || 'None';
    const sharedCommunitiesText = sharedCommunities.map((c: any) => c.text).join(', ') || 'None';

    console.log('Formatted logged-in user communities:', loggedInCommunitiesText);
    console.log('Formatted clicked user communities:', clickedCommunitiesText);
    console.log('Formatted shared communities:', sharedCommunitiesText);

    const matchPrompt = `Compare two professionals and calculate a match score (0-100) based on:
1. Years of experience (similarity in career stage)
2. About (professional background and interests)
3. Hobbies (common interests and activities)
4. Area of interest (shared professional interests)
5. Yammer communities (shared community memberships) - Pay special attention to shared communities
6. New joiners status (both new or both experienced)

Professional 1:
- Years of experience: ${loggedInUser.YearsOfExperience || 'Not specified'}
- About: ${loggedInUser.About || 'Not specified'}
- Hobbies: ${loggedInUser.Hobbies || 'Not specified'}
- Area of interest: ${loggedInUser.AreaOfInterest || 'Not specified'}
- Yammer communities: ${loggedInCommunitiesText}
- New joiner: ${loggedInUser.NewJoiner === true || loggedInUser.NewJoiner === 'Yes' || loggedInUser.NewJoiner === 1 ? 'Yes' : 'No'}

Professional 2:
- Years of experience: ${clickedUser.YearsOfExperience || 'Not specified'}
- About: ${clickedUser.About || 'Not specified'}
- Hobbies: ${clickedUser.Hobbies || 'Not specified'}
- Area of interest: ${clickedUser.AreaOfInterest || 'Not specified'}
- Yammer communities: ${clickedCommunitiesText}
- New joiner: ${clickedUser.NewJoiner === true || clickedUser.NewJoiner === 'Yes' || clickedUser.NewJoiner === 1 ? 'Yes' : 'No'}

Shared communities between both professionals: ${sharedCommunitiesText}

Respond with ONLY a single number between 0-100 representing the match score.`;

    const url = `${endpoint}/openai/deployments/${deployment}/chat/completions?api-version=${apiVersion}`;
    
    console.log('Sending to Azure OpenAI:', { url, prompt: matchPrompt.substring(0, 100) + '...' });
    
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'api-key': apiKey
      },
      body: JSON.stringify({
        model: deployment,
        messages: [
          { role: "user", content: matchPrompt }
        ],
        max_completion_tokens: 5000
      })
    });

    if (!response.ok) {
      console.warn('Match score API error, returning 0');
      return 0;
    }

    const result = await response.json();
    
    if (result.choices && result.choices.length > 0 && result.choices[0].message) {
      const scoreText = result.choices[0].message.content.trim();
      const score = parseInt(scoreText, 10);
      
      if (!isNaN(score) && score >= 0 && score <= 100) {
        console.log('Match score calculated:', score);
        
        // Update or create UserInteractions record with the match score
        try {
          // Check if interaction already exists (regardless of status) - returns ID or 0
          const itemId = await checkUserInteractionExistsAnyStatus(
            context,
            loggedInUser.Email,
            clickedUser.Email
          );

          if (itemId > 0) {
            // Update existing record
            console.log('Updating existing interaction record with match score:', score);
            
            const updateUrl = `https://ygc8n.sharepoint.com/sites/OneIntranet/_api/web/lists/getbytitle('UserInteractions')/items(${itemId})`;
            
            const body = JSON.stringify({
              MatchScore: score
            });

            const updateResponse = await context.spHttpClient.post(updateUrl, SPHttpClient.configurations.v1, {
              headers: {
                'Content-Type': 'application/json',
                'X-HTTP-Method': 'MERGE',
                'If-Match': '*'
              },
              body: body
            });

            if (updateResponse.ok) {
              console.log('UserInteractions record updated with match score:', score);
            } else {
              console.warn('Failed to update UserInteractions record:', updateResponse.status);
            }
          } else {
            // Create new record with match score
            console.log('Creating new interaction record with match score:', score);
            
            const createUrl = `https://ygc8n.sharepoint.com/sites/OneIntranet/_api/web/lists/getbytitle('UserInteractions')/items`;
            
            const body = JSON.stringify({
              Requester: loggedInUser.Email,
              Recipient: clickedUser.Email,
              Status: "ScoreChecked",
              MatchScore: score
            });

            const createResponse = await context.spHttpClient.post(createUrl, SPHttpClient.configurations.v1, {
              headers: {
                'Content-Type': 'application/json'
              },
              body: body
            });

            if (createResponse.ok) {
              console.log('New UserInteractions record created with match score:', score);
            } else {
              console.warn('Failed to create UserInteractions record:', createResponse.status);
            }
          }
        } catch (error) {
          console.error('Error updating/creating UserInteractions record:', error);
        }
        
        return score;
      } else {
        console.warn('Invalid match score received:', scoreText);
        return 0;
      }
    }
    
    return 0;
  } catch (error) {
    console.error('Error calculating match score:', error);
    return 0;
  }
};