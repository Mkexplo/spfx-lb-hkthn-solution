import * as React from 'react';
import styles from './AvatarCard.module.scss';

export interface IAvatarCardProps {
  imageUrl: string;
  screenName: string;
  imageWidth?: string | number;
  imageHeight?: string | number;
}

const AvatarCard: React.FC<IAvatarCardProps> = (props) => {
  const {
    imageUrl,
    screenName,
    imageWidth = '200px',
    imageHeight = '200px'
  } = props;

  return (
    <div className={styles.avatarCard}>
      <div 
        className={styles.avatarImage}
        style={{
          width: imageWidth,
          height: imageHeight
        }}
      >
        {imageUrl ? (
          <img 
            src={imageUrl} 
            alt={screenName}
            style={{
              width: '100%',
              height: '100%',
              objectFit: 'cover',
              borderRadius: '8px'
            }}
          />
        ) : (
          <div style={{
            width: '100%',
            height: '100%',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            backgroundColor: '#f0f0f0',
            borderRadius: '8px'
          }}>
            <p>No Image</p>
          </div>
        )}
      </div>
      <div className={styles.screenNameContainer}>
        <p className={styles.screenName}>{screenName || 'N/A'}</p>
      </div>
    </div>
  );
};

export default AvatarCard;
