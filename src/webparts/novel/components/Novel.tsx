import * as React from 'react';
import styles from './Novel.module.scss';
import { INovelProps } from './INovelProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class Novel extends React.Component<INovelProps, {}> {
  public render(): React.ReactElement<INovelProps> {
    const {
      hasTeamsContext,
      datalist
    } = this.props;

    return (
      <section className={`${styles.novel} ${hasTeamsContext ? styles.teams : ''}`}>
        <div id="data-list">
          <ul>
            {
              this.props.datalist.map(item => <li key={item.applyId}>{item.applyId} {item.workspace} {item.applyNum} {item.applyName} {item.link}</li>)
            }
            </ul>
          </div>
      </section>
    );
  }
}
