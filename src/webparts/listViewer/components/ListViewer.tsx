import * as React from 'react';
import styles from './ListViewer.module.scss';
import { IListViewerProps } from './IListViewerProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as strings from 'ListViewerWebPartStrings';

export default class ListViewer extends React.Component < IListViewerProps, {} > {
  public render(): React.ReactElement<IListViewerProps> {
    return(
      <div className = { styles.listViewer } >
  <div className={styles.container}>
    <div className={styles.row}>
      <div className={styles.column}>
        <span className={styles.title}>{strings.ListViewerWelcomeTitle}</span>
        <p className={styles.subTitle}>{strings.ListViewerWelcomeSubtitle}</p>
        <p className={styles.description}>{escape(this.props.webPartTitle)}</p>
        <a href={strings.ListViewerLearnMoreUrl} className={styles.button}>
          <span className={styles.label}>{strings.ListViewerLearnMoreLabel}</span>
        </a>
      </div>
    </div>
  </div>
      </div >
    );
  }
}
