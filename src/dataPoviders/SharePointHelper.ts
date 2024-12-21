import * as pnp from "sp-pnp-js";
import * as jquery from "jquery";
import { LeaveRequest } from "../dataPoviders/LeaveRequest";
import { ItemAddResult, Web } from "../../node_modules/sp-pnp-js";
let dev_siteURL = "https://contoso.com/teams/EPUniversityDev";
let web: any;
web = new Web(dev_siteURL);
export class DataHelperSP {
  // private listNames: ListNames;
  // private appConfigColumns: AppConfgColumnNames;
  constructor() {
    web = new Web(dev_siteURL);
    //this.getUserProfile = this.getUserProfile.bind(this);
    this.getUserId = this.getUserId.bind(this);
  }

  public getListData(listName: string): Promise<any> {
    return pnp.sp.web.lists
      .getByTitle(listName)
      .items.get()
      .then(response => {
        console.log("Response getListData", response);
        return response;
      });
  }

  public async getListItems(ListName: string): Promise<void> {
    const get_iar: any = await web.lists.getByTitle(ListName).items.getAll();
    console.log("Response getListItems", get_iar);
    return get_iar;
  }

  private async getUserId(emailID: string) {
    const loginName = 'i:0#.f|membership|' + emailID + "'";
    const profile = await pnp.sp.profiles.getPropertiesFor(loginName);
    console.log(profile);
    // const user = await web.siteUsers.getByEmail(emailID);
    // console.log(user);
    // return user.UserId;
  }

  public async addLeaveRequest(ListName: string, sDate: string, eDate: string, reason: string, manager: string, emp: string) {
    //this.getUserId(emp)
    // let empId=this.getUserId(emp);
    // let managerId=this.getUserId(manager);
    const iar: ItemAddResult = await web.lists.getByTitle(ListName).items.add({
      StartDate: sDate,
      EndDate: eDate,
      Reason: reason,
      Employee: emp,
      Manager: manager
    });
    console.log(iar);
    alert("Data Submitted Succesfully");
    return true;
  }

  public async _approval(ListName: string, status: string, Id: number) {
    //this.getUserId(emp)
    // let empId=this.getUserId(emp);
    // let managerId=this.getUserId(manager);
    const iar: ItemAddResult = await web.lists.getByTitle(ListName).items.getById(Id).update({
      Status: status
    });
    console.log(iar);
    if (status == "Approved") {
      alert("Leave Approved Succesfully");
    } else {
      alert("Leave Rejected Succesfully");
    }
    return true;
  }

  public getEmpManager(listName: string, user: string): Promise<any[]> {
    return new Promise<any[]>(
      (
        resolve: (DenodoInfo: any[]) => void,
        reject: (error: any) => void
      ): void => {
        jquery.ajax({
          url:
            "https://contoso.com/teams/EPUniversityDev/" + "_api/web/lists/GetByTitle('" + listName + "')/items?$select=Employee/Title,Employee/EMail,Manager/Title," +
            "Manager/EMail&$expand=Employee,Manager&$filter=Employee/EMail eq '" + user + "'",
          type: "GET",
          headers: {
            accept: "application/json;odata=verbose"
          },
          success: function (data: any) {
            console.log("Responce getProjects", data);
            resolve(data.d.results);

          },
          error: function (err: any) {
            // console.log(err);
          }
        });
      }
    );
  }


  public getLeaveListData(listName: string, user: string): Promise<any[]> {
    let LeaveInfo: LeaveRequest[] = [];
    return new Promise<any[]>(
      (
        resolve: (LeaveInfo: any[]) => void,
        reject: (error: any) => void
      ): void => {
        jquery.ajax({
          url:
            "https://contoso.com/teams/EPUniversityDev/" + "_api/web/lists/GetByTitle('" + listName + "')/items?$select=Employee,Manager,StartDate,EndDate,Status,Reason" +
            "&$filter=Employee eq '" + user + "'",
          type: "GET",
          headers: {
            accept: "application/json;odata=verbose"
          },
          success: function (data: any) {
            console.log("Responce getProjects", data);
            let dataArr: LeaveRequest[] = data.d.results;
            if (dataArr.length > 0) {
              dataArr.forEach((value: any, index: any, array: any) => {
                LeaveInfo.push({
                  Manager: value.Manager,
                  Employee: value.Employee,
                  sDate: value.StartDate,
                  eDate: value.EndDate,
                  Status: value.Status,
                  Reason: value.Reason
                })
              })

              resolve(LeaveInfo);
            }
            else {
              resolve([]);
            }
          },
          error: function (err: any) {
            // console.log(err);
          }
        });
      }
    );
  }

  public getManagerLeaveListData(listName: string, user: string): Promise<any[]> {
    let LeaveInfo: LeaveRequest[] = [];
    return new Promise<any[]>(
      (
        resolve: (LeaveInfo: any[]) => void,
        reject: (error: any) => void
      ): void => {
        jquery.ajax({
          url:
            "https://contoso.com/teams/EPUniversityDev/" + "_api/web/lists/GetByTitle('" + listName + "')/items?$select=Id,Employee,Manager,StartDate,EndDate,Status,Reason" +
            "&$filter=Manager eq '" + user + "' and Status eq 'Pending'",
          type: "GET",
          headers: {
            accept: "application/json;odata=verbose"
          },
          success: function (data: any) {
            console.log("Responce getProjects", data);
            let dataArr: LeaveRequest[] = data.d.results;
            if (dataArr.length > 0) {
              dataArr.forEach((value: any, index: any, array: any) => {
                LeaveInfo.push({
                  Id: value.Id,
                  Manager: value.Manager,
                  Employee: value.Employee,
                  sDate: value.StartDate,
                  eDate: value.EndDate,
                  Status: value.Status,
                  Reason: value.Reason
                })
              })

              resolve(LeaveInfo);
            }
            else {
              resolve([]);
            }
          },
          error: function (err: any) {
            // console.log(err);
          }
        });
      }
    );
  }
}