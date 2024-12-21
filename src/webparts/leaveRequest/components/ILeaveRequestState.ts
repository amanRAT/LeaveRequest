import {LeaveRequest} from "../../../dataPoviders/LeaveRequest";

export interface ILeaveRequestState {
    managerEmail: string;
    managerName:string;
    startDate:string;
    endDate:string;
    reason:string;
    leaveList:LeaveRequest[];    
    managerLeaveList:LeaveRequest[];

}