import * as React from 'react';
import styles from './ListViewer.module.scss';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { IListViewerProps } from './IListViewerProps';
import { IListViewerState } from './IListViewerState';
import ListItem from '../ListItem/ListItem';
import { IViewDefinition } from '../../service/IViewDefinition';
import WebPartTitle from '../WebPartTitle/WebPartTitle';
import { DetailsView } from '../DetailsView';
import * as strings from 'ListViewerWebPartStrings';

export default class ListViewer extends React.Component<IListViewerProps, IListViewerState> {

  constructor(props: IListViewerProps) {
    super(props);
    this.state = {
      items: [],
      viewFields: [],
      loading: true,
      configMissing: !props.viewId || !props.detailsViewId,
      idForDialog: 0
    };
  }

  public componentDidMount(): void {
    this.loadDataAndSetState();
  }

  // tslint:disable-next-line:no-any max-line-length
  public componentDidUpdate(prevProps: Readonly<IListViewerProps>, prevState: Readonly<IListViewerState>, prevContext: any): void {
    if (this.props.viewId !== prevProps.viewId || this.props.detailsViewId !== prevProps.detailsViewId) {
      this.setState({
        configMissing: !this.props.viewId || !this.props.detailsViewId
      }, () => this.loadDataAndSetState());
    }
  }

  public render(): React.ReactElement<IListViewerProps> {
    return <div className={styles.listViewer}>
      {this.props.webPartTitle &&
        <WebPartTitle
          text={this.props.webPartTitle}
          link={this.state.viewUrl}
          hoverText={strings.ViewHoverText}
        />
      }
      {this.state.loading &&
        <Spinner size={SpinnerSize.large} label={strings.SpinnerLoadingLabel}></Spinner>
      }
      {this.state.configMissing &&
        <div className={styles.hinweis}>{strings.ConfigMissingMessage}</div>
      }
      {!this.state.loading && !!this.state.items.length &&
        <div className={styles.itemsContainer}>
          <table className={styles.items}>
            <thead>
              <tr>
                {this.state.viewFields.map(field => <th>{field.Title}</th>)}
                <th>{strings.DetailsColumnHeader}</th>
              </tr>
            </thead>
            <tbody>
              {this.state.items.map(item =>
                <ListItem
                  key={item.ID}
                  item={item}
                  fields={this.state.viewFields}
                  showDetailsDialog={() => this.showDialog(item.ID)}
                />
              )}
            </tbody>
          </table>
        </div>
      }
      {
        !this.state.loading && !this.state.items.length && !!this.props.noEntriesText &&
        <span>{this.props.noEntriesText}</span>
      }
      {!!this.state.idForDialog &&
        <DetailsView
          isOpen={true}
          itemId={this.state.idForDialog}
          service={this.props.service}
          viewId={this.props.detailsViewId}
          onDismiss={this.dialogClosed}
          showBodyCaption={this.props.showBodyCaptionInDetails}
        />
      }
    </div>;
  }

  public async loadDataAndSetState(): Promise<void> {
    const { viewId, service } = this.props;
    if (viewId) {
      const view: IViewDefinition = await service.GetViewDefinition(viewId);
      // tslint:disable-next-line:no-any typedef
      const items = await service.GetListItems(view);
      // tslint:disable-next-line:typedef
      const fields = await this.props.service.GetListFields();
      // tslint:disable-next-line:typedef
      const usedFields = view.ViewFields.Items.map(fieldName => fields.filter(f => f.InternalName === fieldName)[0]);
      this.setState({
        items: items,
        loading: false,
        viewFields: usedFields,
        viewUrl: view.ServerRelativeUrl
      });
    } else {
      this.setState({
        items: [],
        loading: false,
        viewFields: [],
        viewUrl: undefined
      });
    }
  }

  private dialogClosed = (): void => {
    this.setState({
      idForDialog: 0
    });
  }

  private showDialog(id: number): void {
    this.setState({ idForDialog: id });
  }
}