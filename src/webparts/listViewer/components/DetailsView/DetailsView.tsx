// tslint:disable:no-any max-line-length
import * as React from 'react';
import { Dialog, DialogFooter, DialogType, IDialogContentProps } from 'office-ui-fabric-react/lib/Dialog';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';

import { IDetailsViewProps } from './IDetailsViewProps';
import { IFieldInfo } from '../../service/IFieldInfo';
import { IViewDefinition } from '../../service/IViewDefinition';
import * as strings from 'ListViewerWebPartStrings';
import styles from './DetailsView.module.scss';
import { IModalProps } from 'office-ui-fabric-react/lib/Modal';
import { ILink } from '../../service/ILink';

export interface IDetailsViewRow {
    internalName: string;
    title: string;
    type: string;
    valueHtml: string;
    valueText: string;
    attachmentFiles?: ILink[];
}

export interface IDetailsViewState {
    isLoading: boolean;
    error?: string;
    rows: IDetailsViewRow[];
}

export default class DetailsView extends React.Component<IDetailsViewProps, IDetailsViewState> {
    public constructor(props: IDetailsViewProps) {
        super(props);

        this.state = {
            isLoading: false,
            rows: []
        };
    }

    public componentDidMount(): void {
        if (this.props.isOpen) {
            this.load();
        }
    }

    public componentDidUpdate(prevProps: IDetailsViewProps): void {
        const opened: boolean = !prevProps.isOpen && this.props.isOpen;
        const itemChanged: boolean = prevProps.itemId !== this.props.itemId;
        const viewChanged: boolean = prevProps.viewId !== this.props.viewId;

        if (this.props.isOpen && (opened || itemChanged || viewChanged)) {
            this.load();
        }
    }

    public render(): React.ReactElement<IDetailsViewProps> {
        // tslint:disable-next-line:no-any
        const dialogContentProps: IDialogContentProps = {
            type: DialogType.normal,
            onDismiss: this.props.onDismiss,
            title: strings.DetailsDialogTitle,
            showCloseButton: true,
            closeButtonAriaLabel: strings.DetailsDialogCloseButtonAriaLabel,
            className: styles.DetailsDialogContent
        };

        const modalProps: IModalProps = {
            onDismiss: this.props.onDismiss,
            isBlocking: false,
            closeButtonAriaLabel: strings.DetailsDialogCloseButtonAriaLabel,
            containerClassName: styles.DetailsDialogModal
        };

        return (
            <Dialog
                hidden={!this.props.isOpen}
                dialogContentProps={dialogContentProps}
                modalProps={modalProps}
            >
                <div className={styles.DetailsDialogInner}>
                    {this.state.isLoading && <div>{strings.SpinnerLoadingLabel}</div>}
                    {!this.state.isLoading && this.state.error && <div>{this.state.error}</div>}

                    {!this.state.isLoading && !this.state.error && (
                        <div>
                            {this.state.rows.map(r => this.renderField(r))}
                        </div>
                    )}
                </div>
                <DialogFooter>
                    <DefaultButton onClick={this.props.onDismiss} text={strings.DetailsDialogClose} />
                </DialogFooter>
            </Dialog>
        );
    }

    private renderField(row: IDetailsViewRow): JSX.Element {
        switch (row.type) {
            case 'Note': {
                return <div className={styles.Row}>
                    {this.props.showBodyCaption && <div>
                        <label>{row.title}</label>
                        <br />
                    </div>}
                    <div className={styles.BodyText} dangerouslySetInnerHTML={{ __html: row.valueHtml }}></div>
                </div>;
            }
            case 'URL': {
                return <div className={styles.Row}>
                    <label>{row.title}</label><span dangerouslySetInnerHTML={{ __html: row.valueHtml }}></span>
                </div>;
            }
            case 'Attachments': {
                if (row.attachmentFiles && row.attachmentFiles.length) {
                    return <div className={styles.Row}>
                        <label>Anlagen</label>
                        <ul className={styles.Attachmentslist}>
                            {row.attachmentFiles.map(
                                (anlage: ILink, index: number) => {
                                    return <li key={index}>
                                        <a href={anlage.Url} target='_blank'>{anlage.Text}</a>
                                    </li>;
                                }
                            )}
                        </ul>
                    </div>;
                } else {
                    return undefined;
                }
            }
            default: {
                return <div className={styles.Row}>
                    <label>{row.title}</label><span>{row.valueText}</span>
                </div>;
            }
        }
    }

    private load = async (): Promise<void> => {
        if (!this.props.viewId || !this.props.itemId) {
            this.setState({ rows: [], error: undefined, isLoading: false });
            return;
        }

        this.setState({ isLoading: true, error: undefined });

        try {
            const viewFields: IFieldInfo[] = await this.props.service.GetViewFields(this.props.viewId);
            const listFields: IFieldInfo[] = await this.props.service.GetListFields();

            const fieldInternalNames: string[] = viewFields.map(f => f.InternalName);

            const viewForItem: IViewDefinition = {
                ViewQuery:
                    `<Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>${this.props.itemId}</Value></Eq></Where>`,
                ViewFields: { Items: fieldInternalNames },
                ServerRelativeUrl: ''
            };

            const items: any[] = await this.props.service.GetListItemForDetailView(viewForItem);
            const item: any = items && items.length > 0 ? items[0] : undefined;

            const htmlValues: any = item && item.FieldValuesAsHtml ? item.FieldValuesAsHtml : {};
            const textValues: any = item && item.FieldValuesAsText ? item.FieldValuesAsText : {};
            const linksToAttachments: ILink[] | undefined = item && item.AttachmentFiles ? item.AttachmentFiles.map((af) => {
                return {
                    Text: af.FileName,
                    Url: af.ServerRelativeUrl
                };
            }) : undefined;

            const rows: IDetailsViewRow[] = fieldInternalNames.map(internalName => {
                const encodedInternalName: string = internalName.replace(/_/g, '_x005f_');
                const fieldInfo: IFieldInfo = listFields.filter(f => f.InternalName === internalName)[0];
                const title: string = fieldInfo ? fieldInfo.Title : internalName;
                const type: string = fieldInfo ? fieldInfo.Type : '';
                const valueHtml: string = htmlValues && htmlValues[encodedInternalName] ? htmlValues[encodedInternalName] : '';
                const valueText: string = textValues && textValues[encodedInternalName] ? textValues[encodedInternalName] : '';
                const attachmentFiles: ILink[] | undefined = internalName === 'Attachments' ? linksToAttachments : undefined;

                return {
                    internalName,
                    title,
                    type,
                    valueHtml,
                    valueText,
                    attachmentFiles
                };
            });

            this.setState({ rows, isLoading: false, error: undefined });
        } catch (e) {
            const message: string = e && e.message ? e.message : strings.DetailsLoadingError;
            this.setState({ rows: [], isLoading: false, error: message });
        }
    }
}
