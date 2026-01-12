import * as React from 'react';
import { IListItemProps } from './IListItemProps';
import styles from './ListItem.module.scss';
import { IFieldInfo } from '../../service/IFieldInfo';
import { IconButton } from 'office-ui-fabric-react/lib/Button';
import * as strings from 'ListViewerWebPartStrings';

// tslint:disable-next-line:variable-name
const ListItem: React.StatelessComponent<IListItemProps> = (props: IListItemProps) => {
/*
    function twoDigits(n: number): string {
        const r: string = n.toString();
        return r.length < 2 ? `0${r}` : r;
    }

    function formatDate(dt: string): string {
        const d: Date = new Date(dt);
        return `${twoDigits(d.getDate())}.${twoDigits(d.getMonth() + 1)}.${d.getFullYear()}`;
    }
*/
    // tslint:disable-next-line:no-any
    function format(field: IFieldInfo, itemData: any): string {
        const encodedInternalName: string = field.InternalName.replace(/_/g, '_x005f_');
        const value: string = itemData.FieldValuesAsText[encodedInternalName];
        if (value) {
            return value;
        } else {
            return itemData[field.InternalName];
        }
    }

    const { item, fields } = props;

    return <tr
        className={styles.befehl}
        onDoubleClick={() => props.showDetailsDialog(item.ID)}
        title={strings.DoubleClickForDetails}
    >
        {fields.map(
            field => <td>{format(field, item)}</td>
        )}
        <td>
            <IconButton
                iconProps={{ iconName: 'More' }}
                title={strings.DetailsColumnHeader}
                ariaLabel={strings.DetailsColumnHeader}
                onClick={() => props.showDetailsDialog(item.ID)}
            />
        </td>
    </tr>;
};

export default ListItem;