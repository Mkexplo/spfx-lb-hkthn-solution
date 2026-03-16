import * as React from 'react';
import { useState, useEffect, useCallback } from 'react';
import styles from './LbEmpProfileWp.module.scss';
import type { ILbEmpProfileWpProps } from './ILbEmpProfileWpProps';
import { getUserImages, IAvatarImage, getUserRecord, updateUserAvatarId, updateUserProfile as updateUserProfileAPI, getAreaOfInterestList, getHobbiesList, getUsersByAreaOfInterest, getAvatarImageByAvatarId, getUsersByHobbies, createUserInteraction, checkUserInteractionExists, updateUserInteractionStatus, checkReceivedInteractionExists, checkMatchedInteractionExists, getCommunities, getScreenNameAIResponse, calculateMatchScore, getMatchScoreFromUserInteractions } from '../spservice';
import { IconButton, PrimaryButton, TextField, ComboBox } from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import AvatarCard from './AvatarCard';
import { set } from '@microsoft/sp-lodash-subset';
//import { escape } from '@microsoft/sp-lodash-subset';

const LbEmpProfileWp: React.FC<ILbEmpProfileWpProps> = (props) => {
  const [images, setImages] = useState<IAvatarImage[]>([]);
  const [userRecord, setUserRecord] = useState<any>(null);
  const [currentIndex, setCurrentIndex] = useState<number>(0);
  const [isUpdating, setIsUpdating] = useState<boolean>(false);
  const [userProfileLoaded, setUserProfileLoaded] = useState<boolean>(false);
  const [userProfileUpdated, setUserProfileUpdated] = useState<boolean>(false);
  const [editableRecord, setEditableRecord] = useState<any>(null);
  const [areaOfInterestInput, setAreaOfInterestInput] = useState<string>('');
  const [hobbiesInput, setHobbiesInput] = useState<string>('');
  const [areaOfInterestList, setAreaOfInterestList] = useState<any[]>([]);
  const [hobbiesList, setHobbiesList] = useState<any[]>([]);
  const [matchedUsers, setMatchedUsers] = useState<any[]>([]);
  const [userProfileDetailView, setUserProfileDetailView] = useState<boolean>(false);
  const [selectedUserDetail, setSelectedUserDetail] = useState<any>(null);
  const [selectedUserAvatarUrl, setSelectedUserAvatarUrl] = useState<string>('');
  const [isJoining, setIsJoining] = useState<boolean>(false);
  const [requestSent, setRequestSent] = useState<boolean>(false);
  const [requestReceived, setRequestReceived] = useState<boolean>(false);
  const [isMatched, setIsMatched] = useState<boolean>(false);
  const [isRecipientInMatch, setIsRecipientInMatch] = useState<boolean>(false);
  const [hobbiesSearchInput, setHobbiesSearchInput] = useState<string>('');
  const [SearchHobbiesInputLoading, setSearchHobbiesInputLoading] = useState<boolean>(false);
  const [SearchInterestInputLoading, setSearchInterestInputLoading] = useState<boolean>(false);
  const [matchedUsersByHobbies, setMatchedUsersByHobbies] = useState<any[]>([]);
  const [filterBySME, setFilterBySME] = useState<boolean>(false);
  const [filterByRecentJoiners, setFilterByRecentJoiners] = useState<boolean>(false);
  const [matchedInteractionUser, setMatchedInteractionUser] = useState<any>(null);
  const [communitiesList, setCommunitiesList] = useState<any[]>([]);
  const [communitiesInput, setCommunitiesInput] = useState<string>('');
  const [isGeneratingScreenName, setIsGeneratingScreenName] = useState<boolean>(false);
  const [matchScore, setMatchScore] = useState<number>(0);
  const [isCalculatingScore, setIsCalculatingScore] = useState<boolean>(false);

  useEffect(() => {
    getUserImages(props.context).then(imgs => {
      console.log('Loaded images:', imgs);
      let userIndex = -1;
      // Handle both object and direct ID format for avatarID
      const userAvatarId = userRecord?.avatarID?.ID || userRecord?.avatarID;
      for (let i = 0; i < imgs.length; i++) {
        if (imgs[i].avatarID === userAvatarId) {
          userIndex = i;
          break;
        }
      }
      if (userIndex >= 0) {
        const userImg = imgs.splice(userIndex, 1)[0];
        imgs.unshift(userImg);
      }
      setImages(imgs);
      setCurrentIndex(0);
    }).catch(err => {
      console.log('Error fetching images', err);
    });

    getUserRecord(props.context, props.userEmail).then(record => {
      setUserRecord(record);
    }).catch(err => {
      console.log('Error fetching user record', err);
    });

    getAreaOfInterestList(props.context).then(list => {
      setAreaOfInterestList(list);
    }).catch(err => {
      console.log('Error fetching area of interest list', err);
    });

    getHobbiesList(props.context).then(list => {
      setHobbiesList(list);
    }).catch(err => {
      console.log('Error fetching hobbies list', err);
    });

    // Fetch communities from Yammer API
    getCommunities(props.context,props.userEmail).then(communities => {
      console.log('Fetched communities:', communities);
      // Filter out 'allcompany' community if it exists
      const filteredCommunities = communities.filter((community: any) => 
        community.text && community.text.toLowerCase() !== 'allcompany'
      );
      setCommunitiesList(filteredCommunities);
    }).catch(err => {
      console.log('Error fetching communities', err);
      setCommunitiesList([]);
    });
  }, [props.context, props.userEmail]);

  // Reposition user avatar when userRecord is loaded
  useEffect(() => {
    if (userRecord && images.length > 0) {
      const userAvatarId = userRecord?.avatarIDId ;
      console.log('Repositioning user avatar. User Avatar ID:', userAvatarId, 'Available images:', images);
      let userIndex = -1;
      for (let i = 0; i < images.length; i++) {
        if (images[i].avatarID === userAvatarId) {
          userIndex = i;
          break;
        }
      }
      console.log('Found user avatar at index:', userIndex);
      if (userIndex > 0) {
        // Only reorder if user avatar is not already first
        const reorderedImages = [...images];
        const userImg = reorderedImages.splice(userIndex, 1)[0];
        reorderedImages.unshift(userImg);
        setImages(reorderedImages);
        setCurrentIndex(0);
        console.log('Reordered images, moving user avatar to first position');
      }
    }
  }, [userRecord, images.length]);

  // Check for matched interactions on component load
  useEffect(() => {
    const checkForMatchedInteractions = async () => {
      try {
        // Search for any matched interactions where logged-in user is requester or recipient
        const url = `https://ygc8n.sharepoint.com/sites/OneIntranet/_api/web/lists/getbytitle('UserInteractions')/items?$filter=Status eq 'Matched' and ((Requester eq '${props.userEmail}') or (Recipient eq '${props.userEmail}'))&$top=1`;
        const response = await props.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
        const data = await response.json();
        
        if (data.value && data.value.length > 0) {
          const interaction = data.value[0];
          // Determine the other person (if logged-in user is requester, get recipient and vice versa)
          const otherEmail = interaction.Requester === props.userEmail ? interaction.Recipient : interaction.Requester;
          
          // Fetch the other person's record
          const userUrl = `https://ygc8n.sharepoint.com/sites/OneIntranet/_api/web/lists/getbytitle('RegisteredUsers')/items?$filter=Email eq '${otherEmail}'`;
          const userResponse = await props.context.spHttpClient.get(userUrl, SPHttpClient.configurations.v1);
          const userData = await userResponse.json();
          
          if (userData.value && userData.value.length > 0) {
            setMatchedInteractionUser(userData.value[0]);
          }
        }
      } catch (error) {
        console.log('Error checking matched interactions:', error);
      }
    };

    checkForMatchedInteractions();
  }, [props.context, props.userEmail]);

  // Load avatar for selected user in detail view
  useEffect(() => {
    if (selectedUserDetail && userProfileDetailView) {
      const avatarId = selectedUserDetail.avatarID?.ID || selectedUserDetail.avatarID;
      if (avatarId) {
        getAvatarImageByAvatarId(props.context, avatarId).then(url => {
          setSelectedUserAvatarUrl(url);
        }).catch(err => {
          console.log('Error fetching selected user avatar:', err);
          setSelectedUserAvatarUrl('');
        });
      }

      // Check if user has already sent a request to this person
      checkUserInteractionExists(props.context, props.userEmail, selectedUserDetail.Email, 'requested')
        .then(exists => {
          console.log('User sent interaction exists:', exists);
          setRequestSent(exists);
        })
        .catch(err => {
          console.log('Error checking user interaction:', err);
          setRequestSent(false);
        });

      // Check if user has received a request from this person
      checkReceivedInteractionExists(props.context, selectedUserDetail.Email, props.userEmail, 'requested')
        .then(exists => {
          console.log('User received interaction exists:', exists);
          setRequestReceived(exists);
        })
        .catch(err => {
          console.log('Error checking received interaction:', err);
          setRequestReceived(false);
        });

      // Check if there's a matched interaction where user is requester
      checkMatchedInteractionExists(props.context, props.userEmail, selectedUserDetail.Email)
        .then(exists => {
          console.log('Matched interaction exists (user as requester):', exists);
          if (exists) {
            setIsMatched(true);
            setIsRecipientInMatch(false);
          } else {
            // Check if user is recipient in a matched interaction
            checkReceivedInteractionExists(props.context, selectedUserDetail.Email, props.userEmail, 'Matched')
              .then(recipientExists => {
                console.log('User is recipient in matched interaction:', recipientExists);
                if (recipientExists) {
                  setIsMatched(true);
                  setIsRecipientInMatch(true);
                } else {
                  setIsMatched(false);
                  setIsRecipientInMatch(false);
                }
              })
              .catch(err => {
                console.log('Error checking received matched interaction:', err);
                setIsMatched(false);
                setIsRecipientInMatch(false);
              });
          }
        })
        .catch(err => {
          console.log('Error checking matched interaction:', err);
          setIsMatched(false);
          setIsRecipientInMatch(false);
        });
    }
  }, [selectedUserDetail, userProfileDetailView, props.context, props.userEmail]);

  const handleNext = () => {
    if (images.length > 0) {
      setCurrentIndex((images.length + currentIndex + 1) % images.length);
    }
  };

  const handlePrev = () => {
    if (images.length > 0) {
      setCurrentIndex((images.length + currentIndex - 1) % images.length);
    }
  };

  const handleChange = () => {
    handleNext();
  };

  const handleUpdate = async () => {
    if (!currentImage || !userRecord) {
      alert('User record or image not found');
      return;
    }

    setIsUpdating(true);
    try {
      const success = await updateUserAvatarId(props.context, props.userEmail, currentImage.avatarID);
      if (success) {
        alert(`Avatar updated successfully!`);
      } else {
        alert('Failed to update avatar');
      }
    } catch (error) {
      console.error('Error:', error);
      alert('Error updating avatar');
    } finally {
      setIsUpdating(false);
    }
  };

  const loadUserProfile = async () => {
    if (!currentImage || !userRecord) {
      alert('User record not found');
      return;
    }
    setUserProfileLoaded(true);
  }


  const updateUserProfile = async () => {
    if (!currentImage || !userRecord) {
      alert('User record not found');
      return;
    }
    console.log('Loading user record for editing:', userRecord);
    setEditableRecord({ ...userRecord });
    setAreaOfInterestInput('');
    setHobbiesInput('');
    setCommunitiesInput('');
    setUserProfileUpdated(true);
  }

  const handleSaveProfile = async () => {
    if (!editableRecord) {
      alert('No editable record found');
      return;
    }

    console.log('Saving profile with data:', editableRecord);
    setIsUpdating(true);
    try {
      const success = await updateUserProfileAPI(props.context, props.userEmail, editableRecord);
      if (success) {
        setUserRecord(editableRecord);
        setUserProfileUpdated(false);
        alert('Profile updated successfully!');
      } else {
        alert('Failed to update profile');
      }
    } catch (error) {
      console.error('Error:', error);
      alert('Error updating profile');
    } finally {
      setIsUpdating(false);
    }
  }

  const parseMultiChoiceString = (value: string | null): string[] => {
    if (!value) return [];
    if (typeof value === 'string') {
      return value.split('|').map(item => item.trim()).filter(item => item.length > 0);
    }
    return [];
  };

  const getFilteredMatchedUsers = () => {
    let filtered = matchedUsers;
    
    if (filterBySME) {
      // SMEfor is a lookup field, so compare with the text value or ID
      filtered = filtered.filter((user: any) => {
        if (user.SMEFor?.Title) {
          return user.SMEFor.Title === areaOfInterestInput;
        } else if (typeof user.SMEFor === 'string') {
          return user.SMEFor === areaOfInterestInput;
        }
        return false;
      });
    }
    
    if (filterByRecentJoiners) {
      filtered = filtered.filter((user: any) => user.NewJoiner === true || user.NewJoiner === 'Yes' || user.NewJoiner === 1);
    }
    
    return filtered;
  };

  const getFilteredMatchedUsersByHobbies = () => {
    let filtered = matchedUsersByHobbies;
    
    if (filterByRecentJoiners) {
      filtered = filtered.filter((user: any) => user.NewJoiner === true || user.NewJoiner === 'Yes' || user.NewJoiner === 1);
    }
    
    return filtered;
  };

  const addMultiChoice = (fieldName: 'AreaOfInterest' | 'Hobbies' | 'Communities', newValue: string) => {
    if (!newValue.trim()) return;
    const currentValues = parseMultiChoiceString(editableRecord?.[fieldName]);
    if (currentValues.indexOf(newValue.trim()) === -1) {
      const updatedValues = [...currentValues, newValue.trim()];
      setEditableRecord({ ...editableRecord, [fieldName]: updatedValues.join(' | ') });
    }
  };

  const removeMultiChoice = (fieldName: 'AreaOfInterest' | 'Hobbies' | 'Communities', valueToRemove: string) => {
    const currentValues = parseMultiChoiceString(editableRecord?.[fieldName]);
    const updatedValues = currentValues.filter(item => item !== valueToRemove);
    setEditableRecord({ ...editableRecord, [fieldName]: updatedValues.join(' | ') });
  };

  // Callback to handle matched user card click
  const handleMatchedUserClick = useCallback(async (clickedUser: any) => {
    console.log('handleMatchedUserClick called with user:', clickedUser);
    console.log('Before setState - userProfileDetailView:', userProfileDetailView, 'selectedUserDetail:', selectedUserDetail, 'searchInput:', SearchHobbiesInputLoading, 'searchInterestInput:', SearchInterestInputLoading);
    setSearchHobbiesInputLoading(false);
    setSearchInterestInputLoading(false);
    setSelectedUserDetail(clickedUser);
    
    // Try to get existing match score from UserInteractions list
    try {
      const existingScore = await getMatchScoreFromUserInteractions(
        props.context,
        userRecord.Email,
        clickedUser.Email
      );
      setMatchScore(existingScore || 0);
    } catch (err) {
      console.error('Error fetching match score:', err);
      setMatchScore(0);
    }

    setUserProfileDetailView(true);
    console.log('After setState - state updates queued');
  }, [userRecord, props.context]);

  // Handler for calculate match score button
  const handleCalculateMatchScore = useCallback(async () => {
    if (!userRecord || !selectedUserDetail) {
      console.error('User record or selected user detail not available');
      return;
    }

    setIsCalculatingScore(true);
    try {
      const score = await calculateMatchScore(userRecord, selectedUserDetail, props.context);
      console.log('Calculated match score:', score);
      setMatchScore(score);
    } catch (err) {
      console.error('Error calculating match score:', err);
      setMatchScore(0);
    } finally {
      setIsCalculatingScore(false);
    }
  }, [userRecord, selectedUserDetail, props.context]);  // Handler for Join/Unjoin/Accept/Send Message button
  const handleJoinClick = async () => {
    if (!selectedUserDetail) {
      alert('No user selected');
      return;
    }

    setIsJoining(true);
    try {
      if (isMatched) {
        // Send Message - open Teams direct message
        // Always message the selectedUserDetail (the other person)
        const teamsUrl = `https://teams.microsoft.com/l/chat/0/0?users=${encodeURIComponent(selectedUserDetail.Email)}`;
        window.open(teamsUrl, '_blank');
      } 
      else if (requestReceived) {
        // Accept - update status to Matched
        const success = await updateUserInteractionStatus(
          props.context,
          selectedUserDetail.Email,
          props.userEmail,
          'Matched'
        );

        if (success) {
          alert(`Request accepted for ${selectedUserDetail.ScreenName}!`);
          setRequestReceived(false);
          setIsMatched(true);
          setIsRecipientInMatch(true);
        } else {
          alert('Failed to accept request. Please try again.');
        }
      } else if (requestSent) {
        // Unjoin - update status to Cancelled
        const success = await updateUserInteractionStatus(
          props.context,
          props.userEmail,
          selectedUserDetail.Email,
          'Cancelled'
        );

        if (success) {
          alert(`Request cancelled for ${selectedUserDetail.ScreenName}!`);
          setRequestSent(false);
        } else {
          alert('Failed to cancel request. Please try again.');
        }
      } else {
        // Join - create new interaction
        const success = await createUserInteraction(
          props.context,
          props.userEmail,
          selectedUserDetail.Email,
          'requested'
        );

        if (success) {
          alert(`Request sent to ${selectedUserDetail.ScreenName}!`);
          setRequestSent(true);
        } else {
          alert('Failed to send request. Please try again.');
        }
      }
    } catch (error) {
      console.error('Error with join/unjoin/accept request:', error);
      alert('Error processing request. Please try again.');
    } finally {
      setIsJoining(false);
    }
  };

  // Handler for Ask AI button to generate screen name
  const handleAskAI = async () => {
    setIsGeneratingScreenName(true);
    try {
      const generatedName = await getScreenNameAIResponse(userRecord.About, props.context);
      if (generatedName) {
        setEditableRecord({ ...editableRecord, ScreenName: generatedName });
       
      } else {
        alert('Failed to generate screen name.');
      }
    } catch (error) {
      console.error('Error generating screen name:', error);
      alert('Error generating screen name. Please try again.');
    } finally {
      setIsGeneratingScreenName(false);
    }
  };

  // Helper component to render matched user card with avatar
  const MatchedUserCardWrapper: React.FC<{ user: any; context: any; onClick: (user: any) => void }> = ({ user, context, onClick }) => {
    const [avatarUrl, setAvatarUrl] = React.useState<string>('');

    React.useEffect(() => {
      // Extract the AvatarID from the expanded lookup field
      const avatarId = user.avatarID?.ID || user.avatarID;
      if (avatarId) {
        getAvatarImageByAvatarId(context, avatarId).then(url => {
          setAvatarUrl(url);
        }).catch(err => {
          console.log('Error fetching avatar for user:', err);
        });
      }
    }, [user.AvatarID, context]);

    return (
      <div 
        onClick={() => {
          console.log('Card clicked for user:', user.ScreenName);
          onClick(user);
        }}
        style={{ cursor: 'pointer', transition: 'transform 0.2s', borderRadius: '8px', padding: '10px' }}
        onMouseEnter={(e) => (e.currentTarget.style.transform = 'scale(1.05)', e.currentTarget.style.backgroundColor = '#f0f0f0')}
        onMouseLeave={(e) => (e.currentTarget.style.transform = 'scale(1)', e.currentTarget.style.backgroundColor = 'transparent')}
      >
        <div style={{ pointerEvents: 'none' }}>
          <AvatarCard 
            imageUrl={avatarUrl}
            screenName={user.ScreenName || 'N/A'}
            imageWidth="150px"
            imageHeight="150px"
          />
        </div>
      </div>
    );
  };

  const currentImage = images.length > 0 ? images[currentIndex] : null;

  return (
    <div className={styles.lbEmpProfileWp} style={{ display: 'flex', gap: '20px', padding: '20px' }}>
      <div className={styles.userAvatarContainer}>
        <div className={styles.carousel}>
          {currentImage ? (
            <>
              <img 
                src={currentImage.url} 
                alt="Avatar" 
                style={{ width: '100%', height: '100%', objectFit: 'cover' }} 
              />
              
            </>
          ) : (
            <div style={{ width: '100%', height: '100%', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
              <p>Loading avatar...</p>
            </div>
          )}
        </div>
        <div style={{ marginTop: '10px', display: 'flex', gap: '10px',marginLeft:'25px',marginRight:'25px' }}>
          <PrimaryButton text="Change" onClick={handleChange} className={styles.pButton}/>
          <PrimaryButton text="Update" onClick={handleUpdate} disabled={isUpdating} className={styles.pButton}/>
        </div>
        <div style={{ marginTop: '10px', display: 'flex', gap: '10px', width: '80%' ,marginLeft:'25px',marginRight:'25px' }}>
          <PrimaryButton text="Update my profile" onClick={updateUserProfile} className={styles.pButton}/>
        </div>
        {/* Show Launch Connect container if matched interaction exists */}
        {matchedInteractionUser ? (
          <div style={{ 
            marginTop: '50px', 
            marginBottom: '20px', 
            marginLeft: '15px', 
            marginRight: '15px',
            padding: '20px',
            border: '2px solid #9D4EDD',
            borderRadius: '8px',
            backgroundColor: '#f3ecff',
            background: 'linear-gradient(to bottom, #d6a6ff 0%, #b87cff 40%, #9d4edd 100%)',
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            justifyContent: 'center',
            gap: '15px',
            minHeight: '120px'
          }}>
            <label style={{ 
              fontSize: '18px', 
              fontWeight: '600',
              color: '#333',
              textAlign: 'center'
            }}>
              Connect with {matchedInteractionUser.ScreenName} over lunch today!
            </label>
            <PrimaryButton 
              text="Send message" 
              onClick={() => {
                const teamsUrl = `https://teams.microsoft.com/l/chat/0/0?users=${encodeURIComponent(matchedInteractionUser.Email)}`;
                window.open(teamsUrl, '_blank');
              }}
              className={styles.pButton}
            />
          </div>
        ) : (
          <div style={{ 
            marginTop: '50px', 
            marginBottom: '20px', 
            marginLeft: '15px', 
            marginRight: '15px',
            padding: '20px',
            border: '2px solid #9D4EDD',
            borderRadius: '8px',
            backgroundColor: '#f3ecff',
            background: 'linear-gradient(to bottom, #d6a6ff 0%, #b87cff 40%, #9d4edd 100%)'
          }}>
          {/* Search by Area of Interest */}
          <div style={{ marginBottom: '20px' }}>
            <label className={styles.propLabelLight} style={{ marginBottom: '10px', display: 'block', fontSize: '18px' }}>Search by Area of interest</label>
            <ComboBox
              placeholder="Select area of interest"
              options={areaOfInterestList}
              selectedKey={null}
              text={areaOfInterestInput}
              onChange={(event: any, option: any, index: number | undefined, value: string | undefined) => {
                const selectedValue = option?.text || value || '';
                setAreaOfInterestInput(selectedValue);
                 setSearchInterestInputLoading(true);
                 setSearchHobbiesInputLoading(false);
                 setHobbiesSearchInput('');
                if (selectedValue) {
                  getUsersByAreaOfInterest(props.context, selectedValue, props.userEmail).then(users => {
                    console.log('Matched users:', users);
                    setMatchedUsers(users);
                  }).catch(err => {
                    console.log('Error fetching matched users:', err);
                    setMatchedUsers([]);
                  });
                } else {
                  setMatchedUsers([]);
                }
              }}
              autoComplete="on"
              allowFreeform={true}
              styles={{ root: { width: '100%' } }}
            />
          </div>

          {/* Search by Hobbies Section */}
          <div style={{ marginTop: '20px', marginBottom: '0px' }}>
            <label className={styles.propLabelLight} style={{ display: 'block', marginBottom: '10px',fontSize: '18px'  }}>
              Search by Hobbies
            </label>
            <ComboBox
              placeholder="Select hobbies"
              options={hobbiesList}
              selectedKey={null}
              text={hobbiesSearchInput}
              onChange={(event: any, option: any, index: number | undefined, value: string | undefined) => {
                const selectedValue = option?.text || value || '';
                setHobbiesSearchInput(selectedValue);
                setSearchHobbiesInputLoading(true);
                setSearchInterestInputLoading(false);
                setAreaOfInterestInput('');
                if (selectedValue) {
                  getUsersByHobbies(props.context, selectedValue, props.userEmail).then(users => {
                    console.log('Matched users by hobbies:', users);
                    setMatchedUsersByHobbies(users);
                  }).catch(err => {
                    console.log('Error fetching matched users by hobbies:', err);
                    setMatchedUsersByHobbies([]);
                  });
                } else {
                  setMatchedUsersByHobbies([]);
                }
              }}
              autoComplete="on"
              allowFreeform={true}
              styles={{ root: { width: '100%' } }}
            />
          </div>
        </div>
        )}
      </div>
      <div className={styles.userProfileDiv}>
        {userProfileUpdated ? (
          <div style={{ marginTop: '20px' }}>
            <div style={{ marginBottom: '15px' }}>
              <label className={styles.propLabel} style={{ marginTop: '10px',color:'rgb(157, 78, 221)',fontSize:'24px' }}>My Profile</label>
            </div>
            <div style={{ marginBottom: '15px' }}>
              <label className={styles.propLabel}>Screen Name:</label>
              <div style={{ marginTop: '10px', display: 'flex', gap: '5px', alignItems: 'flex-end' }}>
                <TextField 
                  value={editableRecord?.ScreenName || ''} 
                  onChange={(e, newValue) => setEditableRecord({ ...editableRecord, ScreenName: newValue })}
                  className={styles.textFieldNoBorder}
                  borderless
                  style={{ marginLeft: '10px', flex: 1 }}
                />
                <PrimaryButton 
                  text="Ask AI" 
                  onClick={handleAskAI}
                  disabled={isGeneratingScreenName}
                  className={styles.pButton}
                  style={{ width: '100px', marginLeft: '10px' }}
                />
              </div>
            </div>
            
            <div style={{ marginBottom: '15px' }}>
              <label className={styles.propLabel}>Area of Interest:</label>
              <div style={{ marginTop: '10px' }}>
                <div style={{ display: 'flex', gap: '5px' }}>
                  <ComboBox
                    placeholder="Search and select area of interest"
                    options={areaOfInterestList}
                    selectedKey={null}
                    onChange={(event: any, option: any, index: number | undefined, value: string | undefined) => {
                      if (option) {
                        setAreaOfInterestInput(option.text);
                      } else if (value) {
                        setAreaOfInterestInput(value);
                      }
                    }}
                    autoComplete="on"
                    allowFreeform={true}
                    styles={{ root: { width: '100%', flex: 1} }}
                  />
                  <PrimaryButton 
                    text="Add" 
                    onClick={() => {
                      addMultiChoice('AreaOfInterest', areaOfInterestInput);
                      setAreaOfInterestInput('');
                    }}
                    className={styles.pButton}
                    style={{ width: '60px', marginLeft: '5px' }}
                  />
                </div>
                <div style={{ marginTop: '10px', display: 'flex', flexWrap: 'wrap', gap: '8px' }}>
                  {parseMultiChoiceString(editableRecord?.AreaOfInterest).map((item: string, index: number) => (
                    <div 
                      key={index}
                      style={{
                        backgroundColor: '#9D4EDD',
                        color: 'white',
                        padding: '8px 16px',
                        borderRadius: '20px',
                        display: 'flex',
                        alignItems: 'center',
                        gap: '8px',
                        fontSize: '14px',
                        fontWeight: '500'
                      }}
                    >
                      {item}
                      <IconButton 
                        iconProps={{ iconName: 'Cancel' }}
                        onClick={() => removeMultiChoice('AreaOfInterest', item)}
                        styles={{ root: { color: 'white', fontSize: '12px', padding: '0px', minWidth: '20px', minHeight: '20px' } }}
                      />
                    </div>
                  ))}
                </div>
              </div>
            </div>
            <div style={{ marginBottom: '15px' }}>
              <label className={styles.propLabel}>Hobbies:</label>
              <div style={{ marginTop: '10px' }}>
                <div style={{ display: 'flex', gap: '5px' }}>
                  <ComboBox
                    placeholder="Search and select hobby"
                    options={hobbiesList}
                    selectedKey={null}
                    onChange={(event: any, option: any, index: number | undefined, value: string | undefined) => {
                      if (option) {
                        setHobbiesInput(option.text);
                      } else if (value) {
                        setHobbiesInput(value);
                      }
                    }}
                    autoComplete="on"
                    allowFreeform={true}
                    styles={{ root: { width: '100%', flex: 1 } }}
                  />
                  <PrimaryButton 
                    text="Add" 
                    onClick={() => {
                      addMultiChoice('Hobbies', hobbiesInput);
                      setHobbiesInput('');
                    }}
                    className={styles.pButton}
                    style={{ width: '60px', marginLeft: '5px' }}
                  />
                </div>
                <div style={{ marginTop: '10px', display: 'flex', flexWrap: 'wrap', gap: '8px' }}>
                  {parseMultiChoiceString(editableRecord?.Hobbies).map((item: string, index: number) => (
                    <div 
                      key={index}
                      style={{
                        backgroundColor: '#9D4EDD',
                        color: 'white',
                        padding: '8px 16px',
                        borderRadius: '20px',
                        display: 'flex',
                        alignItems: 'center',
                        gap: '8px',
                        fontSize: '14px',
                        fontWeight: '500'
                      }}
                    >
                      {item}
                      <IconButton 
                        iconProps={{ iconName: 'Cancel' }}
                        onClick={() => removeMultiChoice('Hobbies', item)}
                        styles={{ root: { color: 'white', fontSize: '12px', padding: '0px', minWidth: '20px', minHeight: '20px' } }}
                      />
                    </div>
                  ))}
                </div>
              </div>
            </div>
           
            <div style={{ marginBottom: '15px' }}>
              <label className={styles.propLabel}>About:</label>
              <div style={{ marginTop: '10px' }}>
                <TextField 
                  multiline
                  rows={9}
                  value={editableRecord?.About || ''} 
                  onChange={(e, newValue) => setEditableRecord({ ...editableRecord, About: newValue })}
                  className={styles.textFieldNoBorder}
                  borderless
                  style={{ marginLeft: '10px' }}
                />
              </div>
            </div>
            <div style={{ marginTop: '20px', display: 'flex', gap: '10px' }}>
              <PrimaryButton text="Save" onClick={handleSaveProfile} disabled={isUpdating} className={styles.pSubmitButton} />
              <PrimaryButton text="Cancel" onClick={() => setUserProfileUpdated(false)} className={styles.pSubmitButton} />
            </div>
          </div>):SearchInterestInputLoading?(
            <div style={{ marginTop: '20px', overflowY: 'auto', flex: 1 }}>
              <div style={{ marginBottom: '15px' }}>
                <label className={styles.propLabel} style={{ marginTop: '10px',color:'rgb(157, 78, 221)',fontSize:'24px' }}>Matches based on Area of Interest</label>
              </div>
              <div style={{ marginBottom: '20px', display: 'flex', gap: '20px', justifyContent: 'flex-end' }}>
                <label style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'pointer' }}>
                  <input
                    type="checkbox"
                    checked={filterBySME}
                    onChange={(e) => setFilterBySME(e.target.checked)}
                    style={{ cursor: 'pointer', width: '18px', height: '18px' }}
                  />
                  <span style={{ fontSize: '14px', color: '#333' }}>Show SMEs</span>
                </label>
                <label style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'pointer' }}>
                  <input
                    type="checkbox"
                    checked={filterByRecentJoiners}
                    onChange={(e) => setFilterByRecentJoiners(e.target.checked)}
                    style={{ cursor: 'pointer', width: '18px', height: '18px' }}
                  />
                  <span style={{ fontSize: '14px', color: '#333' }}>Recent joiners</span>
                </label>
              </div>
              {matchedUsers.length > 0 ? (
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: '20px', padding: '10px 0' }}>
                  {getFilteredMatchedUsers().map((user: any, index: number) => (
                    <MatchedUserCardWrapper 
                      key={index} 
                      user={user} 
                      context={props.context}
                      onClick={handleMatchedUserClick}
                    />
                  ))}
                </div>
              ) : (
                <p style={{ color: '#666', fontSize: '14px' }}>No matching employees found for this area of interest.</p>
              )}
            </div>
              
          ):SearchHobbiesInputLoading?(
            <div style={{ marginTop: '20px', overflowY: 'auto', flex: 1 }}>
              <div style={{ marginBottom: '15px' }}>
                <label className={styles.propLabel} style={{ marginTop: '10px',color:'rgb(157, 78, 221)',fontSize:'24px' }}>Matches based on Hobbies</label>
              </div>
              <div style={{ marginBottom: '20px', display: 'flex', gap: '20px', justifyContent: 'flex-end' }}>
                <label style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'pointer' }}>
                  <input
                    type="checkbox"
                    checked={filterByRecentJoiners}
                    onChange={(e) => setFilterByRecentJoiners(e.target.checked)}
                    style={{ cursor: 'pointer', width: '18px', height: '18px' }}
                  />
                  <span style={{ fontSize: '14px', color: '#333' }}>Recent joiners</span>
                </label>
              </div>
              {matchedUsersByHobbies.length > 0 ? (
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: '20px', padding: '10px 0' }}>
                  {getFilteredMatchedUsersByHobbies().map((user: any, index: number) => (
                    <MatchedUserCardWrapper 
                      key={index} 
                      user={user} 
                      context={props.context}
                      onClick={handleMatchedUserClick}
                    />
                  ))}
                </div>
              ) : (
                <p style={{ color: '#666', fontSize: '14px' }}>No matching employees found for this hobby.</p>
              )}
            </div>
              
          ):userProfileDetailView?(
        <div style={{ marginTop: '20px', overflowY: 'auto', flex: 1, display: 'flex', flexDirection: 'column' }}>
            {/* Header with Avatar and Screen Name */}
            <div style={{ marginBottom: '30px', display: 'flex', gap: '15px', alignItems: 'center', paddingBottom: '20px', borderBottom: '1px solid #e0e0e0' }}>
              <img 
                src={selectedUserAvatarUrl} 
                alt={selectedUserDetail?.ScreenName} 
                style={{ width: '80px', height: '80px', borderRadius: '8px', objectFit: 'cover' }} 
              />
              <div style={{ flex: 1 }}>
                <label className={styles.propLabel} style={{ marginTop: '0px', color:'rgb(157, 78, 221)', fontSize:'20px' }}>
                  {selectedUserDetail?.ScreenName || 'N/A'}
                </label>
                <div style={{ marginTop: '10px', display: 'flex', alignItems: 'center', gap: '10px' }}>
                  <PrimaryButton 
                    text={isCalculatingScore ? 'Calculating...' : 'View Match Score'}
                    onClick={handleCalculateMatchScore}
                    disabled={isCalculatingScore}
                    style={{ 
                      padding: '8px 16px', 
                      fontSize: '14px', 
                      height: '32px',
                      background: 'linear-gradient(to bottom, #e6ccff, #9b4ded)',
                      color: 'white',
                      border: 'none',
                      borderRadius: '10px',
                      boxShadow: '0 3px 8px rgba(106, 13, 173, 0.4)',
                      cursor: 'pointer'
                    }}
                  />
                  {matchScore > 0 && (
                    <div style={{
                      backgroundColor: matchScore >= 75 ? '#4CAF50' : matchScore >= 50 ? '#FFC107' : '#FF9800',
                      color: 'white',
                      padding: '8px 16px',
                      borderRadius: '20px',
                      fontWeight: '600',
                      fontSize: '14px',
                      
                    }}>
                      {`${matchScore}%`}
                    </div>
                  )}
                </div>
              </div>
            </div>
                     
            <div style={{ marginBottom: '15px', flex: 1 }}>
              <label className={styles.propLabel}>Area of Interest:</label>
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: '8px', marginTop: '10px' }}>
                {parseMultiChoiceString(selectedUserDetail?.AreaOfInterest).length > 0 ? (
                  parseMultiChoiceString(selectedUserDetail?.AreaOfInterest).map((item: string, index: number) => (
                    <div 
                      key={index}
                      style={{
                        backgroundColor: '#9D4EDD',
                        color: 'white',
                        padding: '8px 16px',
                        borderRadius: '20px',
                        fontSize: '14px',
                        fontWeight: '500'
                      }}
                    >
                      {item}
                    </div>
                  ))
                ) : (
                  <p className={styles.propLabel} >N/A</p>
                )}
              </div>
            </div>
            <div style={{ marginBottom: '15px' }}>
              <label className={styles.propLabel}>Hobbies:</label>
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: '8px', marginTop: '10px' }}>
                {parseMultiChoiceString(selectedUserDetail?.Hobbies).length > 0 ? (
                  parseMultiChoiceString(selectedUserDetail?.Hobbies).map((item: string, index: number) => (
                    <div 
                      key={index}
                      style={{
                        backgroundColor: '#9D4EDD',
                        color: 'white',
                        padding: '8px 16px',
                        borderRadius: '20px',
                        fontSize: '14px',
                        fontWeight: '500'
                      }}
                    >
                      {item}
                    </div>
                  ))
                ) : (
                  <p className={styles.propLabel} >N/A</p>
                )}
              </div>
            </div>
            <div style={{ marginBottom: '15px' }}>
              <label className={styles.propLabel}>About:</label>
              <div style={{ marginTop: '10px' }}>
                <p className={styles.propLabel} >{selectedUserDetail?.About || 'N/A'}</p>
              </div>
            </div>

            {/* Back Button at Bottom */}
            <div style={{ marginTop: '30px', paddingTop: '20px', borderTop: '1px solid #e0e0e0', display: 'flex', gap: '10px' }}>
              <PrimaryButton 
                text="Back" 
                onClick={() => {
                  setUserProfileDetailView(false);
                  setSelectedUserDetail(null);
                  setSelectedUserAvatarUrl('');
                  setRequestSent(false);
                  setRequestReceived(false);
                  setIsMatched(false);
                  setIsRecipientInMatch(false);
                  if(areaOfInterestInput!==''){
                    setSearchInterestInputLoading(true);
                  }else if(hobbiesSearchInput!==''){
                    setSearchHobbiesInputLoading(true);
                  }
                }}
                className={styles.pSubmitButton}
              />
              <PrimaryButton 
                text={
                  isMatched ? "Send message" : 
                  requestReceived ? "Accept" : 
                  requestSent ? "Unjoin" : 
                  "Join"
                } 
                onClick={handleJoinClick}
                disabled={isJoining}
                className={styles.pSubmitButton}
              />
            </div>
          </div>
          ):(
          <div style={{ marginTop: '20px' }}>
            <div style={{ marginBottom: '15px' }}>
              <label className={styles.propLabel} style={{ marginTop: '10px',color:'rgb(157, 78, 221)',fontSize:'24px' }}>My Profile</label>
            </div>
            <div style={{ marginBottom: '15px' }}>
              <label className={styles.propLabel}>Screen Name:</label>
              <label className={styles.propLabel}>{userRecord?.ScreenName || 'N/A'}</label>
            </div>
            
            <div style={{ marginBottom: '15px' }}>
              <label className={styles.propLabel}>Area of Interest:</label>
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: '8px', marginTop: '10px' }}>
                {parseMultiChoiceString(userRecord?.AreaOfInterest).length > 0 ? (
                  parseMultiChoiceString(userRecord?.AreaOfInterest).map((item: string, index: number) => (
                    <div 
                      key={index}
                      style={{
                        backgroundColor: '#9D4EDD',
                        color: 'white',
                        padding: '8px 16px',
                        borderRadius: '20px',
                        fontSize: '14px',
                        fontWeight: '500'
                      }}
                    >
                      {item}
                    </div>
                  ))
                ) : (
                  <p className={styles.propLabel} >N/A</p>
                )}
              </div>
            </div>
            <div style={{ marginBottom: '15px' }}>
              <label className={styles.propLabel}>Hobbies:</label>
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: '8px', marginTop: '10px' }}>
                {parseMultiChoiceString(userRecord?.Hobbies).length > 0 ? (
                  parseMultiChoiceString(userRecord?.Hobbies).map((item: string, index: number) => (
                    <div 
                      key={index}
                      style={{
                        backgroundColor: '#9D4EDD',
                        color: 'white',
                        padding: '8px 16px',
                        borderRadius: '20px',
                        fontSize: '14px',
                        fontWeight: '500'
                      }}
                    >
                      {item}
                    </div>
                  ))
                ) : (
                  <p className={styles.propLabel} >N/A</p>
                )}
              </div>
            </div>
            <div style={{ marginBottom: '15px' }}>
              <label className={styles.propLabel}>My Communities:</label>
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: '8px', marginTop: '10px' }}>
                {communitiesList.length > 0 ? (
                  communitiesList.map((community: any, index: number) => (
                    <div 
                      key={index}
                      style={{
                        backgroundColor: '#9D4EDD',
                        color: 'white',
                        padding: '8px 16px',
                        borderRadius: '20px',
                        fontSize: '14px',
                        fontWeight: '500'
                      }}
                    >
                      {community.text}
                    </div>
                  ))
                ) : (
                  <p className={styles.propLabel} >N/A</p>
                )}
              </div>
            </div>
            <div style={{ marginBottom: '15px' }}>
              <label className={styles.propLabel}>About:</label>
              <div style={{ marginTop: '10px' }}>
                <p className={styles.propLabel} >{userRecord?.About || 'N/A'}</p>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

export default LbEmpProfileWp;
