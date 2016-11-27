import * as React from 'react';
import ContactCard from './ContactCard';
import styles from '../MyContacts.module.scss';
import { Spinner, SpinnerType } from 'office-ui-fabric-react/lib/Spinner';

export default class ContactCardList extends React.Component<any, {}>{
  public render():JSX.Element {
    var contactsList = this.props.Contacts.map((contact) => <ContactCard Contact={contact} PictureSize={this.props.PictureSize} ShowPhone={this.props.ShowPhone} key={contact.Id} />);
    return (this.props.IsLoading)
            ? <Spinner type={ SpinnerType.large } label='Loading contacts...' className={styles.spinner}/>
            : <div>
                {contactsList}
                <div className={styles.clearBoth}></div>
              </div>;
  };
}