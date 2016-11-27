import * as React from 'react';
import { Persona } from 'office-ui-fabric-react/lib/Persona';
import styles from '../MyContacts.module.scss';

export default class ContactCard extends React.Component<any, {}>{
  public render():JSX.Element {
    var phone = (this.props.ShowPhone) ? this.props.Contact.Phone : "";
    return <div className={styles.contact}>
              <Persona primaryText={this.props.Contact.DisplayName}
                    secondaryText={this.props.Contact.Email}
                    tertiaryText={phone}
                    size={this.props.PictureSize}
                    imageUrl={this.props.Contact.Image} />
           </div>;
  };
}