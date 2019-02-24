import * as React from 'react';
import styles from './TreeOrgChart.module.scss';
import { ITreeOrgChartProps } from './ITreeOrgChartProps';
import { ITreeOrgChartState } from './ITreeOrgChartState';
import { escape } from '@microsoft/sp-lodash-subset';
import SortableTree from 'react-sortable-tree';
import 'react-sortable-tree/style.css';
import { IPersonaSharedProps, Persona, PersonaSize, PersonaPresence } from 'office-ui-fabric-react/lib/Persona';
import { IconButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import spservice from '../../../services/spservices';
import { ITreeChildren } from './ITreeChildren';
import { ITreeData } from './ITreeData';
import { Spinner ,SpinnerSize} from 'office-ui-fabric-react/lib/components/Spinner'

import { ColorClassNames } from '@uifabric/styling/lib';

export default class TreeOrgChart extends React.Component<ITreeOrgChartProps, ITreeOrgChartState> {
  private treeData: ITreeData[];
  private treeChildren: ITreeChildren[];
  private SPService: spservice;

  constructor(props) {
    super(props);

    this.SPService = new spservice(this.props.context);
    this.onContactInfo = this.onContactInfo.bind(this);

    this.state = {
      treeData: [],
      isLoading: true
    };
  }
  //
  private handleTreeOnChange(treeData) {
    this.setState({ treeData });
  }

  //
  public async componentDidUpdate(prevProps: ITreeOrgChartProps, prevState: ITreeOrgChartState) {
    if (this.props.currentUserTeam !== prevProps.currentUserTeam || this.props.maxLevels !== prevProps.maxLevels) {
      await this.loadOrgchart();
    }
  }
  //
  public async componentDidMount() {
    await this.loadOrgchart();
  }

  // Load Organization Chart
  public async loadOrgchart() {
    this.setState({isLoading: true});
    const currentUser = `i:0#.f|membership|${this.props.context.pageContext.user.loginName}`;
    const currentUserProperties = await this.SPService.getUserProperties(currentUser);
    this.treeData = [];
    // Test if show only my Team or All Organization Chart
    if (!this.props.currentUserTeam) {
      const treeManagers = await this.buildOrganizationChart(currentUserProperties);
      this.treeData.push({ title: (treeManagers.person), expanded: true, children: treeManagers.treeChildren });
    } else {
      const treeManagers = await this.buildMyTeamOrganizationChart(currentUserProperties);
      this.treeData.push({ title: (treeManagers.person), expanded: true, children: treeManagers.treeChildren });
    }
    console.log(JSON.stringify(this.treeData));
    this.setState({ treeData: this.treeData , isLoading: false});
  }

  public async buildOrganizationChart(currentUserProperties: any) {
    // Get Managers
    let managers: any[] = currentUserProperties.ExtendedManagers;
    const treeManagers = await this.getManagers(managers);
    return treeManagers;
  }
  // Get Managersyyy
  private async getManagers(managers: any[]) {
    let treeChildren: ITreeChildren[] = [];
    let person: any;
    let spUser: IPersonaSharedProps = {};

    for (let index = 0; index < managers.length; index++) {
      const manager = managers[index];
      const userProperties = await this.SPService.getUserProperties(manager);
      const imageInitials: string[] = userProperties.DisplayName.split(' ');
      spUser.imageUrl = `/_layouts/15/userphoto.aspx?size=L&username=${userProperties.Email}`;
      spUser.imageInitials = `${imageInitials[0].substring(0, 1).toUpperCase()}${imageInitials[1].substring(0, 1).toUpperCase()}`;
      spUser.text = userProperties.DisplayName;
      spUser.tertiaryText = userProperties.Email;
      spUser.secondaryText = userProperties.Title;
      // Top Manager
      if (index === 0) {
        const topManager = spUser;
        person = <Persona {...topManager} hidePersonaDetails={false} size={PersonaSize.size40} />;
      }
      else {
        const person = <Persona {...spUser} hidePersonaDetails={false} size={PersonaSize.size40} />;
        if (userProperties.DirectReports && userProperties.DirectReports.length > 0) {
          const usersDirectReports: any[] = await this.getDirectReports(userProperties.DirectReports);
          treeChildren.push({ title: (person), children: usersDirectReports });
        } else {
          treeChildren.push({ title: (person) });
        }
      }
    }

    return { 'person': person, 'treeChildren': treeChildren };
  }
  // Get Managers
  private async getDirectReports(userDirectReports: any[]) {

    let treeChildren: ITreeChildren[] = [];
    let spUser: IPersonaSharedProps = {};

    for (const user of userDirectReports) {
      const userProperties = await this.SPService.getUserProperties(user);
      const imageInitials: string[] = userProperties.DisplayName.split(' ');

      spUser.imageUrl = `/_layouts/15/userphoto.aspx?size=L&username=${userProperties.Email}`;
      spUser.imageInitials = `${imageInitials[0].substring(0, 1).toUpperCase()}${imageInitials[1].substring(0, 1).toUpperCase()}`;
      spUser.text = userProperties.DisplayName;
      spUser.tertiaryText = userProperties.Email;
      spUser.secondaryText = userProperties.Title;
      const person = <Persona {...spUser} hidePersonaDetails={false} size={PersonaSize.size40} />;
      const usersDirectReports = await this.getDirectReports(userProperties.DirectReports);

      usersDirectReports ? treeChildren.push({ title: (person), children: usersDirectReports }) :
        treeChildren.push({ title: (person) });
    }
    return treeChildren;
  }

  private async buildMyTeamOrganizationChart(currentUserProperties: any) {

    let spUser: IPersonaSharedProps = {};
    let me: IPersonaSharedProps = {};
    let treeChildren: ITreeChildren[] = [];
    const myManager = await this.SPService.getUserProfileProperty(currentUserProperties.AccountName, 'Manager');
    const userProperties = await this.SPService.getUserProperties(myManager);
    const imageInitials: string[] = userProperties.DisplayName.split(' ');

    spUser.imageUrl = `/_layouts/15/userphoto.aspx?size=L&username=${userProperties.Email}`;
    spUser.imageInitials = `${imageInitials[0].substring(0, 1).toUpperCase()}${imageInitials[1].substring(0, 1).toUpperCase()}`;
    spUser.text = userProperties.DisplayName;
    spUser.tertiaryText = userProperties.Email;
    spUser.secondaryText = userProperties.Title;
    const managerCard = <Persona {...spUser} hidePersonaDetails={false} size={PersonaSize.size40} />;
    const meImageInitials: string[] = currentUserProperties.DisplayName.split(' ');

    me.imageUrl = `/_layouts/15/userphoto.aspx?size=L&username=${currentUserProperties.Email}`;
    me.imageInitials = `${imageInitials[0].substring(0, 1).toUpperCase()}${meImageInitials[1].substring(0, 1).toUpperCase()}`;
    me.text = currentUserProperties.DisplayName;
    me.tertiaryText = currentUserProperties.Email;
    me.secondaryText = currentUserProperties.Title;
    const meCard = <Persona {...me} hidePersonaDetails={false} size={PersonaSize.size40} />;
    const usersDirectReports: any[] = await this.getDirectReports(currentUserProperties.DirectReports);

    treeChildren.push({ title: (meCard), expanded: true, children: usersDirectReports });

    return { 'person': managerCard, 'treeChildren': treeChildren };
  }
  // Contacto Info
  private onContactInfo(): void {

    window.open(`https://eur.delve.office.com/?p=${this.props.context.pageContext.user.loginName}&v=work`);
  }
  public render(): React.ReactElement<ITreeOrgChartProps> {
    return (
      <div className={styles.treeOrgChart}>
        <WebPartTitle displayMode={this.props.displayMode}
          title={this.props.title}
          updateProperty={this.props.updateProperty} />
          {
            this.state.isLoading ? <Spinner size={SpinnerSize.large} label="Loading ..."></Spinner> : null
          }

        <div className={styles.treeContainer}>
          <SortableTree
            treeData={this.state.treeData}
            onChange={this.handleTreeOnChange.bind(this)}
            canDrag={false}
            canDrop={false}
            rowHeight={70}
            maxDepth={this.props.maxLevels}
            generateNodeProps={rowInfo => ({
              buttons: [
                <IconButton
                  disabled={false}
                  checked={false}
                  iconProps={{ iconName: 'ContactInfo' }}
                  title="Contact Info"
                  ariaLabel="Contact"
                  onClick={(ev) => {
                    window.open(`https://eur.delve.office.com/?p=${rowInfo.node.title.props.tertiaryText}&v=work`);
                  }}
                />
              ],
            })}
          />
        </div>
      </div>
    );
  }


}
