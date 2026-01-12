import * as React from 'react';
import { IWebPartTitleProps } from './IWebPartTitleProps';
import styles from './WebPartTitle.module.scss';

// tslint:disable-next-line:variable-name
const WebPartTitle: React.SFC<IWebPartTitleProps> = (props) => {
    const navigateToLink: () => void = () => {
        window.location.href = props.link;
    };

    return (
        <div className={styles.webPartHeader}>
            <div className={styles.webPartTitle}>
                {props.link ?
                    <span
                        className={styles.link}
                        role='heading'
                        aria-level='2'
                        title={props.hoverText}
                        onDoubleClick={navigateToLink}
                    >
                        {props.text}
                    </span>
                    :
                    <span role='heading' aria-level='2'>{props.text}</span>
                }
            </div>
        </div>);
};

export default WebPartTitle;
