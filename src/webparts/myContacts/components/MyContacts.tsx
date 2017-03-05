import * as React from 'react';
import { IWebPartContext } from '@microsoft/sp-webpart-base';
import { Button, ButtonType } from 'office-ui-fabric-react/lib/Button';
import styles from '../MyContacts.module.scss';
import { IMyContactsWebPartProps } from '../IMyContactsWebPartProps';
import * as Managers from '../managers/Managers';
import * as Model from '../model/Model';
import ContactCardList from './ContactCardList';

export interface IMyContactsProps extends IMyContactsWebPartProps {
  dataManager: Managers.IDataManager;
  spContext: IWebPartContext;
}

export interface IMyContactsState {
  pageNumber: number;
  nextResultsAvailable: boolean;
  prevResultsAvailable: boolean;
  contacts: Model.IContact[];
  isLoading: boolean;
}

export default class MyContacts extends React.Component<IMyContactsProps, IMyContactsState> {

  constructor(props) {
    super(props);
    this.state = { pageNumber: 0, nextResultsAvailable: false, prevResultsAvailable: false, contacts: [], isLoading: true };
  };

  public componentDidMount() : void {
    if(this.props.listId !== undefined && this.props.listId !== '') {
      this.getContacts(this.props.pageSize, this.state.pageNumber);
    }
  };

  public componentWillReceiveProps(props: IMyContactsProps) {
    this.getContacts(props.pageSize, this.state.pageNumber);
  };

  private getContacts(pageSize: number, pageNumber: number) : void {
    this.props.dataManager.GetContacts(pageSize, pageNumber).then((result) => {
      this.setState((previousState: IMyContactsState, currentProps: IMyContactsProps): IMyContactsState => {
        previousState.contacts = result.Results;
        previousState.nextResultsAvailable = ((this.state.pageNumber + 1) * this.props.pageSize) < result.ItemCount;
        previousState.prevResultsAvailable = (this.state.pageNumber > 0);
        previousState.isLoading = false;
        return previousState;
      });
    });
  };

  private nextContacts() : void {
    this.setState((previousState: IMyContactsState, currentProps: IMyContactsProps): IMyContactsState => {
      previousState.pageNumber = previousState.pageNumber + 1;
      previousState.isLoading = true;
      return previousState;
    }, () => {
      this.getContacts(this.props.pageSize, this.state.pageNumber);
    });
  };

  private prevContacts() : void {
    this.setState((previousState: IMyContactsState, currentProps: IMyContactsProps): IMyContactsState => {
      previousState.pageNumber = previousState.pageNumber - 1;
      previousState.isLoading = true;
      return previousState;
    }, () => {
      this.getContacts(this.props.pageSize, this.state.pageNumber);
    });
  };

  public render(): JSX.Element {
    if(this.props.listId === undefined || this.props.listId === '') {
      return <div>
              <h5>Please select a contact list from property pane.</h5>
             </div>;
    }
    return (
      <div className={styles.myContacts}>
        <div className={styles.container}>
          <h1>My contacts from "{this.props.spContext.pageContext.web.title}"</h1>
          <div className="ms-Grid">
            <div className={styles.contactList + " ms-Grid-row"}>
              <div className="ms-Grid-col ms-u-lg12">
                <ContactCardList Contacts={this.state.contacts} IsLoading={this.state.isLoading} PictureSize={this.props.pictureSize} ShowPhone={this.props.showPhone} />
              </div>
            </div>
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-u-lg4" />
              <div className="ms-Grid-col ms-u-lg4">
                <Button buttonType={ButtonType.primary} onClick={() => this.prevContacts() } disabled={!this.state.prevResultsAvailable}>Previous</Button>
                <Button buttonType={ButtonType.primary} onClick={() => this.nextContacts() } disabled={!this.state.nextResultsAvailable}>Next</Button>
              </div>
              <div className="ms-Grid-col ms-u-lg4" />
            </div>
          </div>
        </div>
      </div>
    );
  }
}
