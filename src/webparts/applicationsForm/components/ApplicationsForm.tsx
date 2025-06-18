import * as React from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IApplicationsFormProps, IApplicationsFormState } from './IApplicationsFormProps';
//import { IApplicationsListItem } from './IApplicationsListItem';
import styles from './ApplicationsForm.module.scss';

export default class ApplicationsForm extends React.Component<IApplicationsFormProps, IApplicationsFormState> {
  private listName: string = "Applications";

  constructor(props: IApplicationsFormProps) {
    super(props);

    this.state = {
      items: [],
      applicationName: '',
      applicationType: 'COTS',
      contact: '',
      errors: {
        applicationName: '',
        contact: ''
      }
    };

    this.handleInputChange = this.handleInputChange.bind(this);
    this.handleSubmit = this.handleSubmit.bind(this);
    this.handleCancel = this.handleCancel.bind(this);
    this.validateForm = this.validateForm.bind(this);
  }

  public componentDidMount(): void {
    this.loadItems();
  }

  private loadItems(): void {
    const { context } = this.props;

    context.spHttpClient.get(
      `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.listName}')/items?$select=Id,Title,ApplicationType,Contact`,
      SPHttpClient.configurations.v1
    )
    .then((response: SPHttpClientResponse) => response.json())
    .then((response) => {
      this.setState({
        items: response.value
      });
    })
    .catch((error) => {
      console.error('Error loading items:', error);
    });
  }

  private handleInputChange(event: React.FormEvent<HTMLInputElement | HTMLSelectElement>): void {
    const { name, value } = event.currentTarget;

    this.setState({
      ...this.state,
      [name]: value
    });
  }

  private validateForm(): boolean {
    let isValid = true;
    const errors = {
      applicationName: '',
      contact: ''
    };

    // Validate Application Name
    if (!this.state.applicationName || this.state.applicationName.length < 3 || this.state.applicationName.length > 50) {
      errors.applicationName = 'Application Name must be between 3 and 50 characters';
      isValid = false;
    } else if (/[^a-zA-Z0-9 ]/.test(this.state.applicationName)) {
      errors.applicationName = 'Application Name cannot contain special characters';
      isValid = false;
    }

    // Validate Email
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!this.state.contact || !emailRegex.test(this.state.contact)) {
      errors.contact = 'Please enter a valid email address';
      isValid = false;
    }

    this.setState({
      errors: errors
    });

    return isValid;
  }

  private handleSubmit(event: React.FormEvent<HTMLFormElement>): void {
    event.preventDefault();

    if (!this.validateForm()) {
      return;
    }

    const { context } = this.props;
    const { applicationName, applicationType, contact } = this.state;

    const body: string = JSON.stringify({
      Title: applicationName,
      ApplicationType: applicationType,
      Contact: contact
    });

    context.spHttpClient.post(
      `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.listName}')/items`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata',
          'odata-version': ''
        },
        body: body
      }
    )
    .then((response: SPHttpClientResponse) => response.json())
    .then(() => {
      this.setState({
        applicationName: '',
        applicationType: 'COTS',
        contact: ''
      });
      this.loadItems();
    })
    .catch((error) => {
      console.error('Error saving item:', error);
    });
  }

  private handleCancel(): void {
    this.setState({
      applicationName: '',
      applicationType: 'COTS',
      contact: '',
      errors: {
        applicationName: '',
        contact: ''
      }
    });
  }

  public render(): React.ReactElement<IApplicationsFormProps> {
    const { items, applicationName, applicationType, contact, errors } = this.state;

    return (
      <div className={styles.applicationsForm}>
        <div className={styles.container}>
          <div className={styles.formSection}>
            <h2>Application Form</h2>
            <form onSubmit={this.handleSubmit}>
              <div className={styles.formRow}>
                <div className={styles.formGroup}>
                  <label htmlFor="applicationName">Application Name:</label>
                  <input
                    type="text"
                    id="applicationName"
                    name="applicationName"
                    value={applicationName}
                    onChange={this.handleInputChange}
                    className={errors.applicationName ? styles.error : ''}
                  />
                  {errors.applicationName && <div className={styles.errorMessage}>{errors.applicationName}</div>}
                </div>
                
                <div className={styles.formGroup}>
                  <label htmlFor="applicationType">Application Type:</label>
                  <select
                    id="applicationType"
                    name="applicationType"
                    value={applicationType}
                    onChange={this.handleInputChange}
                  >
                    <option value="Select option">---Select---</option>
                    <option value="COTS">COTS</option>
                    <option value="Custom">Custom</option>
                  </select>
                </div>
                
                <div className={styles.formGroup}>
                  <label htmlFor="contact">Contact Email:</label>
                  <input
                    type="email"
                    id="contact"
                    name="contact"
                    value={contact}
                    onChange={this.handleInputChange}
                    className={errors.contact ? styles.error : ''}
                  />
                  {errors.contact && <div className={styles.errorMessage}>{errors.contact}</div>}
                </div>
              </div>
              
              <div className={styles.buttonGroup}>
                <button type="submit" className={styles.saveButton}>Save</button>
                <div>
                 Data Saved Successfully
                 </div>
                <button type="button" onClick={this.handleCancel} className={styles.cancelButton}>Reset</button>
              </div>
            </form>
          </div>
          
          <div className={styles.gridSection}>
            <h2>Applications List</h2>
            <div className={styles.gridContainer}>
              <table>
                <thead>
                  <tr>
                    <th>Application Name</th>
                    <th>Application Type</th>
                    <th>Contact Email</th>
                  </tr>
                </thead>
                <tbody>
                  {items.map((item) => (
                    <tr key={item.Id}>
                      <td>{item.Title}</td>
                      <td>{item.ApplicationType}</td>
                      <td>{item.Contact}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
// import * as React from 'react';
// import styles from './ApplicationsForm.module.scss';
// import type { IApplicationsFormProps } from './IApplicationsFormProps';
// import { escape } from '@microsoft/sp-lodash-subset';

// export default class ApplicationsForm extends React.Component<IApplicationsFormProps, {}> {
//   public render(): React.ReactElement<IApplicationsFormProps> {
//     const {
//       description,
//       isDarkTheme,
//       environmentMessage,
//       hasTeamsContext,
//       userDisplayName
//     } = this.props;

//     return (
//       <section className={`${styles.applicationsForm} ${hasTeamsContext ? styles.teams : ''}`}>
//         <div className={styles.welcome}>
//           <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
//           <h2>Well done, {escape(userDisplayName)}!</h2>
//           <div>{environmentMessage}</div>
//           <div>Web part property value: <strong>{escape(description)}</strong></div>
//         </div>
//         <div>
//           <h3>Welcome to SharePoint Framework!</h3>
//           <p>
//             The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It&#39;s the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
//           </p>
//           <h4>Learn more about SPFx development:</h4>
//           <ul className={styles.links}>
//             <li><a href="https://aka.ms/spfx" target="_blank" rel="noreferrer">SharePoint Framework Overview</a></li>
//             <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank" rel="noreferrer">Use Microsoft Graph in your solution</a></li>
//             <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank" rel="noreferrer">Build for Microsoft Teams using SharePoint Framework</a></li>
//             <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank" rel="noreferrer">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
//             <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank" rel="noreferrer">Publish SharePoint Framework applications to the marketplace</a></li>
//             <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank" rel="noreferrer">SharePoint Framework API reference</a></li>
//             <li><a href="https://aka.ms/m365pnp" target="_blank" rel="noreferrer">Microsoft 365 Developer Community</a></li>
//           </ul>
//         </div>
//       </section>
//     );
//   }
// }
