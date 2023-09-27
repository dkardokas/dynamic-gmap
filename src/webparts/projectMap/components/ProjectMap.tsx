import * as React from 'react';
import styles from './ProjectMap.module.scss';
import type { IProjectMapProps } from './IProjectMapProps';
import { escape } from '@microsoft/sp-lodash-subset';
import MapLoader from './MapLoader';
import { Wrapper } from '@googlemaps/react-wrapper';

export default class ProjectMap extends React.Component<IProjectMapProps, {}> {

  public render(): React.ReactElement<IProjectMapProps> {
    const {
      description,      
      hasTeamsContext,
      mapApiKey,
      listItems,
      startLat,
      startLon
    } = this.props;

    return (
      <section className={`${styles.projectMap} ${hasTeamsContext ? styles.teams : ''}`}>
        <Wrapper apiKey={mapApiKey}>
          <div className={styles.welcome}>
            <h2>{escape(description)}</h2>
          </div>
          <MapLoader spListItems={listItems} startLat={startLat} startLon={startLon} />
        </Wrapper>


        {/* <div className={styles.welcome}>
          <div>{environmentMessage}</div>
        </div> */}

      </section>
    );
  }
}
