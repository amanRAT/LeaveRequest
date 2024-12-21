import * as React from 'react';
//import styles from './LeaveRequest.module.scss';
import type { ILeaveRequestProps } from './ILeaveRequestProps';
import type { ILeaveRequestState } from './ILeaveRequestState';
import { IStyleSet, Label, ILabelStyles, Pivot, PivotItem } from '@fluentui/react';
import { TextField } from '@fluentui/react/lib/TextField';
import { PrimaryButton } from '@fluentui/react/lib/Button';
import { DetailsList, DetailsListLayoutMode, IColumn } from '@fluentui/react/lib/DetailsList';
// import { IPersonaProps } from '@fluentui/react/lib/Persona';
// import { CompactPeoplePicker, IBasePickerSuggestionsProps, ValidationState } from '@fluentui/react/lib/Pickers';
import { IPersonaSharedProps, Persona, PersonaSize } from '@fluentui/react/lib/Persona';
import {
  DatePicker,
  defaultDatePickerStrings,
} from '@fluentui/react';
//import { escape } from '@microsoft/sp-lodash-subset';
import { DataHelperSP } from "../../../dataPoviders/SharePointHelper";
import styles from "./LeaveRequest.module.scss";
// import { Web} from "../../../../node_modules/sp-pnp-js";
// let dev_siteURL="https://contoso.com/teams/EPUniversityDev";
//let web: any ;
var ManagerList = "ManagerList";
var leaveList = "LeaveRequest";
//var curUser:string=null;
const labelStyles: Partial<IStyleSet<ILabelStyles>> = {
  root: { marginTop: 10 },
};

export default class LeaveRequest extends React.Component<ILeaveRequestProps, ILeaveRequestState, {}> {
  public dataHelper: DataHelperSP;

  constructor(props: ILeaveRequestProps) {
    super(props);
    let curUser = props.userEmail;
    this.dataHelper = new DataHelperSP();
    this.state = {
      managerEmail: "",
      managerName: "",
      startDate: "",
      endDate: "",
      reason: "",
      leaveList: [],
      managerLeaveList: []
    };
    this.dataHelper.getEmpManager(ManagerList, curUser).then(response => {
      console.log("Res", response);
      if (response.length > 0) {
        this.setState({
          managerEmail: response[0].Manager.EMail,
          managerName: response[0].Manager.Title
        })
      }
    });
    this._getLeaveList();
    this.getManagerList();
    this._changeReason = this._changeReason.bind(this);
    this._changeEndtDate = this._changeEndtDate.bind(this);
    this._changeStartDate = this._changeStartDate.bind(this);
    this.handleApproveClick = this.handleApproveClick.bind(this);
    this.handleRejectClick = this.handleRejectClick.bind(this);
    this._getLeaveList = this._getLeaveList.bind(this);
    this.getManagerList = this.getManagerList.bind(this);
  }

  private _getLeaveList() {
    this.dataHelper.getLeaveListData(leaveList, this.props.userEmail).then(res => {
      console.log("LeaveList", res);
      if (res.length > 0) {
        this.setState({
          leaveList: res
        })
      }
    });
  }
  private getManagerList() {
    this.dataHelper.getManagerLeaveListData(leaveList, this.props.userEmail).then(resManager => {
      console.log("ManagerLeaveList", resManager);
      if (resManager.length > 0) {
        this.setState({
          managerLeaveList: resManager
        })
      }
    });
  }
  private _submitClicked(): void {
    let sDate = this.state.startDate;
    let eDate = this.state.endDate;
    let reason = this.state.reason;
    let manager = this.state.managerEmail;
    let emp = this.props.userEmail;
    this.dataHelper.addLeaveRequest(leaveList, sDate, eDate, reason, manager, emp).then(() => {
      this._getLeaveList();
      this.getManagerList();
    });
  }
  private _changeStartDate(event: Date): void {
    console.log(event.toLocaleDateString());
    //let sDate=event.toLocaleDateString();
    this.setState({
      startDate: event.toLocaleDateString()
    });
  } private _changeEndtDate(event: Date): void {
    console.log(event);
    this.setState({
      endDate: event.toLocaleDateString()
    });
  } private _changeReason(event: string): void {
    console.log(event);
    this.setState({
      reason: event
    });
  }
  private handleApproveClick(ID: number): void {
    console.log(ID);
    this.dataHelper._approval(leaveList, "Approved", ID).then(() => {
      this._getLeaveList();
      this.getManagerList();
    });
  }
  private handleRejectClick(ID: number): void {
    console.log(ID);
    this.dataHelper._approval(leaveList, "Rejected", ID).then(() => {
      this._getLeaveList();
      this.getManagerList();
    });
  }

  public render(): React.ReactElement<ILeaveRequestProps> {
    let _columns: IColumn[] = [
      { key: 'Start Date', name: 'Start Date', fieldName: 'sDate', minWidth: 100, maxWidth: 150, isResizable: true },
      { key: 'End Date', name: 'End Date', fieldName: 'eDate', minWidth: 100, maxWidth: 150, isResizable: true },
      { key: 'Reason', name: 'Reason', fieldName: 'Reason', minWidth: 200, maxWidth: 250, isResizable: true },
      { key: 'Status', name: 'Status', fieldName: 'Status', minWidth: 100, maxWidth: 150, isResizable: true },
    ];
    let examplePerson: IPersonaSharedProps;
    const examplePersona: IPersonaSharedProps = {
      //imageUrl: TestImages.personaFemale,
      //imageInitials: 'AL',
      text: this.state.managerName,
      secondaryText: 'Software Engineer',
      tertiaryText: 'In a meeting',
      optionalText: 'Available at 4:00pm',
    };
    return (
      <div>
        <Pivot aria-label="Basic Pivot Example">
          <PivotItem
            headerText="Employee View"
            headerButtonProps={{
              'data-order': 1,
              'data-title': 'Employee View Title',
            }}
          >
            <Label className={styles.labelsTag} styles={labelStyles}>Leave Details</Label>
            <DetailsList
              items={this.state.leaveList}
              columns={_columns}
              //setKey="set"
              layoutMode={DetailsListLayoutMode.justified}
              isSelectedOnFocus={false}
              disableSelectionZone
              //selection={this._selection}
              //selectionPreservedOnEmptyClick={true}
              ariaLabelForSelectionColumn="Toggle selection"
              ariaLabelForSelectAllCheckbox="Toggle selection for all items"
            //checkButtonAriaLabel="select row"
            />
            <br></br>
            <Label className={styles.labelsTag} styles={labelStyles}>Apply Leave</Label><br></br>
            <Label styles={labelStyles}>Start Date</Label>
            <DatePicker
              //firstDayOfWeek={firstDayOfWeek}
              placeholder="Start date"
              ariaLabel="Start Date"
              //onChange={e=>this._changeStartDate(e)}
              onSelectDate={e => this._changeStartDate(e as Date)}
              // DatePicker uses English strings by default. For localized apps, you must override this prop.
              strings={defaultDatePickerStrings}
            />

            <Label styles={labelStyles}>End Date</Label>
            <DatePicker
              //firstDayOfWeek={firstDayOfWeek}
              placeholder="End date"
              ariaLabel="End Date"
              onSelectDate={e => this._changeEndtDate(e as Date)}
              minDate={new Date(this.state.startDate)}
              // DatePicker uses English strings by default. For localized apps, you must override this prop.
              strings={defaultDatePickerStrings}
            />
            <Label styles={labelStyles}>Manager</Label>
            <Persona
              {...examplePersona}
              size={PersonaSize.size24}
              //hidePersonaDetails={!renderDetails}
              imageAlt="Annie Lindqvist, no presence detected"
            />
            <TextField label="Reason" multiline rows={3} onChange={e => this._changeReason((e.target as HTMLInputElement).value)} /><br></br>
            <PrimaryButton text="Submit" onClick={this._submitClicked.bind(this)} allowDisabledFocus disabled={false} checked={true} />
          </PivotItem>
          <PivotItem headerText="Manager View">
            <div>
              {this.state.managerLeaveList.length > 0 ?
                (<table className={styles.tables}>
                  <thead className={styles.tableHead}>
                    <tr>
                      <th className={styles.tableHead}>Employee</th>
                      <th className={styles.tableHead}>Start Date</th>
                      <th className={styles.tableHead}>End Date</th>
                      <th className={styles.tableHead}>Reason</th>
                      <th className={styles.tableHead}>Status</th>
                      <th className={styles.tableHead}>Actions</th>
                    </tr>
                  </thead>
                  <tbody className={styles.tableBody}>
                    {this.state.managerLeaveList.map((item) => {
                      examplePerson = {
                        text: item.Employee,
                        secondaryText: 'Software Engineer',
                        tertiaryText: 'In a meeting',
                        optionalText: 'Available at 4:00pm',
                      }
                      return (
                        <tr key={item.Id}>
                          {/* <td className={styles.tableBody}>{item.Employee}</td> */}
                          <td className={styles.tableBody}>
                            <Persona
                              {...examplePerson}
                              size={PersonaSize.size24}
                              //hidePersonaDetails={!renderDetails}
                              imageAlt="Annie Lindqvist, no presence detected"
                            /></td>
                          <td className={styles.tableBody}>{item.sDate}</td>
                          <td className={styles.tableBody}>{item.eDate}</td>
                          <td className={styles.tableBody}>{item.Reason}</td>
                          <td className={styles.tableBody}>{item.Status}</td>
                          <td className={styles.tableBody}>
                            <PrimaryButton
                              className={styles.buttonTag}
                              text="Approve"
                              onClick={() => this.handleApproveClick(item.Id as number)}
                            />
                            <PrimaryButton
                              className={styles.buttonTag}
                              text="Reject"
                              onClick={() => this.handleRejectClick(item.Id as number)}
                            />
                          </td>
                        </tr>
                      )
                    })}
                  </tbody>
                </table>
                ) : <div><Label styles={labelStyles}>No Data Available</Label></div>
              }

            </div>
          </PivotItem>
        </Pivot>
      </div>
    );
  }
}
