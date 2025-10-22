import * as React from "react";
import { hideLoader, showLoader } from "../Shared/Loader";
import { Navigate } from "react-router-dom";
import { SPHttpClient } from "@microsoft/sp-http";
// import { SPFx, spfi } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";
import "@pnp/sp/folders";
// import "../CSS/CHECKLISTSTEP2Form.css";
import { ActionStatus, ControlType } from "../Constants/Contants";
import { showToast } from "../Shared/Toaster";
import { highlightCurrentNav } from "../Utilities/HighlightCurrentComponent";
import DatePickercontrol from "../Shared/DatePickerField";
import SearchableDropdown from "../Shared/Dropdown";
import { initCommonFunctions } from "../Utilities/CommonFunctions";
import DateUtilities from "../Utilities/DateUtilities";
import { addDays } from 'office-ui-fabric-react';

import { format } from "date-fns";
// import FileUpload from "../Shared/FileUpload";
// import InputCheckBox from "../Shared/InputCheckBox";
// import { format } from "date-fns";
import Formvalidator from "../Utilities/FormValidator";

export interface CHECKLISTSTEP2FormProps {
    match: any;
    spContext: any;
    spHttpClient: SPHttpClient;
    context: any;
    history: any;
    userDisplayName: string;
    siteURL: string;
    webAbsoluteURL: string;
    currPlantTitle: string;
    isSuperAdmin: boolean;
}

export interface CHECKLISTSTEP2FormState {
}

export default class CHECKLISTSTEP2Form extends React.Component<CHECKLISTSTEP2FormProps, CHECKLISTSTEP2FormState> {
    private rootSiteURL: string; private currentSiteURL: string; private MaycoURL: string;
    // private sp = spfi().using(SPFx(this.props.context)); 
    public state = {
        //Cascaded dropdowns
        Plants: [], Departments: [], Zones: [], Machines: [], Shifts: [], Auditors: [],
        PlantsOpt: [], DepartmentsOpt: [], ZonesOpt: [], MachinesOpt: [], ShiftsOpt: [], AuditorsOpt: [],
        formData: {
            Date: '',
            Plant: '',
            Department: '',
            Zone: '',
            Machine: '',
            Shift: '',
            Auditor: '',
        },
        CheckListTable: {
            'Safety Policies/Procedures': [
                { No: 0, Checks: 'Certification of Authorized employees in LOTO are complete.', Low: 1, Medium: 2, High: 3, AuditScore: '', Score: [1, 2, 3], Comments: '' },
                { No: 1, Checks: 'LOTO placards are visible on all machines/equipment.', Low: 1, Medium: 2, High: 3, AuditScore: '', Score: [1, 2, 3], Comments: '' },
                { No: 2, Checks: 'Machine Guarding Checklist - Implemented, employees trained on hazards.', Low: 1, Medium: 2, High: 3, AuditScore: '', Score: [1, 2, 3], Comments: '' },
                { No: 3, Checks: 'Legal Compliance Calendar is reviewed Quarterly and updated as needed.', Low: 0, Medium: 1, High: 2, AuditScore: '', Score: [0, 1, 2], Comments: '' },
                { No: 4, Checks: 'Safety Versatility Matrix is established and maintained to indicate training.', Low: 0, Medium: 1, High: 2, AuditScore: '', Score: [0, 1, 2], Comments: '' },
                { No: 5, Checks: 'Safety Incentive/Reward Program implemented. (Safety Game,Safety Incentive Program - Dept, etc.)', Low: 0, Medium: 0.5, High: 1, AuditScore: '', Score: [0, 0.5, 1], Comments: '' },
                { No: 6, Checks: 'Safety Team established with representation from various departments.', Low: 0, Medium: 1, High: 2, AuditScore: '', Score: [0, 1, 2], Comments: '' },
                { No: 7, Checks: 'Visitor/Contractor Policy and procedure exists and adhered to.', Low: 0, Medium: 0.5, High: 1, AuditScore: '', Score: [0, 0.5, 1], Comments: '' },
            ],
            'Management of Accidents': [
                { No: 8, Checks: 'Share Point is established for the location and being used to complete safety forms and identify issues.', Low: 0, Medium: 1, High: 2, AuditScore: '', Score: [0, 1, 2], Comments: '' },
                { No: 9, Checks: 'Accident events are attacked and analyzed using the SEWO form.Temporary solutions are implemented and Corrective actions implemented and tracked.', Low: 1, Medium: 2, High: 3, AuditScore: '', Score: [1, 2, 3], Comments: '' },
                { No: 10, Checks: 'The dynamics of the event are discussed with the persons involved (increased awareness of the worker, Safety Alert).', Low: 0, Medium: 1, High: 2, AuditScore: '', Score: [0, 1, 2], Comments: '' },
                { No: 11, Checks: "SA KPI's are reviewed regularly and communicated across location.(Heinrich, Incident Rate,etc)", Low: 1, Medium: 2, High: 3, AuditScore: '', Score: [1, 2, 3], Comments: '' },
            ],
            'Management of Near Misses/Unsafe Acts/ Unsafe Conditions': [
                { No: 12, Checks: 'SA Tag program is implemented to allow employees method to identify and track safety issues.', Low: 0, Medium: 1, High: 2, AuditScore: '', Score: [0, 1, 2], Comments: '' },
                { No: 13, Checks: 'UCAN form is utilized to report unsafe conditions and unsafe behaviors. Corrective actions are implemented and tracked.', Low: 0, Medium: 1, High: 2, AuditScore: '', Score: [0, 1, 2], Comments: '' },
                { No: 14, Checks: 'SMAT audits are implemented, management is trained and schedule is communicated.', Low: 0, Medium: 1, High: 2, AuditScore: '', Score: [0, 1, 2], Comments: '' },
                { No: 15, Checks: 'Safety Alerts are communicated regarding near misses/unsafe behaviors/acts and/or injuries.', Low: 0, Medium: 1, High: 2, AuditScore: '', Score: [0, 1, 2], Comments: '' },
            ],
            'Visual Management': [
                { No: 16, Checks: 'Visual control of accident data exists. PPE information, Safety Posters are visible throughout plant.', Low: 0, Medium: 1, High: 2, AuditScore: '', Score: [0, 1, 2], Comments: '' },
                { No: 17, Checks: 'Pedestrian Isleways are marked/painted. Stop signs/mirrors are visible throughout facility.', Low: 0, Medium: 1, High: 2, AuditScore: '', Score: [0, 1, 2], Comments: '' },
                { No: 18, Checks: 'PPE is easily available for employees and the use of PPE is visible throughout facility.', Low: 1, Medium: 2, High: 3, AuditScore: '', Score: [1, 2, 3], Comments: '' },
                { No: 19, Checks: 'Prescription eyewear program/employee footwear program in place.', Low: 0, Medium: 1, High: 2, AuditScore: '', Score: [0, 1, 2], Comments: '' },
                { No: 20, Checks: 'Dress code requirements are reviewed and visible throughout facility to Mayco Standard.	', Low: 0, Medium: 0.5, High: 1, AuditScore: '', Score: [0, 0.5, 1], Comments: '' },
                { No: 21, Checks: 'Work Instructions Reflect the job steps taken, and show required PPE', Low: 0, Medium: 1, High: 2, AuditScore: '', Score: [0, 1, 2], Comments: '' },
            ],
            'Assessment of risks (JSRA)': [
                { No: 22, Checks: 'JSRA Routine- In Process for all manufacturing positions(assembly, molding, paint, etc.), updated when changes are implemented.', Low: 0, Medium: 1, High: 2, AuditScore: '', Score: [0, 1, 2], Comments: '' },
                { No: 23, Checks: 'JSRA Non-Routine- In Process for non-routine positions (maint, diesetter, quality, etc.), updated when changes are implemented.', Low: 0, Medium: 1, High: 2, AuditScore: '', Score: [0, 1, 2], Comments: '' },
                { No: 24, Checks: 'Confined spaces are identified and posted. Confined space program exists.', Low: 0, Medium: 1, High: 2, AuditScore: '', Score: [0, 1, 2], Comments: '' },
                { No: 25, Checks: 'Fall protection program exists/employees are trained in hazards.', Low: 0, Medium: 1, High: 2, AuditScore: '', Score: [0, 1, 2], Comments: '' },
                { No: 26, Checks: 'Ladder inspection program implemented to inspect all fixed/portable ladders/platforms- identify and correct issues.', Low: 0, Medium: 1, High: 2, AuditScore: '', Score: [0, 1, 2], Comments: '' },
                { No: 27, Checks: 'New equipment is designed according to Mayco Standard and approved by Safety prior to production.', Low: 0, Medium: 1, High: 2, AuditScore: '', Score: [0, 1, 2], Comments: '' },
            ]
        },
        TotalAuditScore: '0',
        isEditForm: false,
        showSubmit: false,
        ItemId: 0,
        Redirect: false,
        RedirectTo: '',
    }

    constructor(props: CHECKLISTSTEP2FormProps) {
        super(props);
        this.rootSiteURL = this.props.spContext.siteAbsoluteUrl; //   sites/wcm
        this.currentSiteURL = this.props.spContext.webAbsoluteUrl;//  sites/wcm/mayco/merrill/sa
        this.MaycoURL = `${this.rootSiteURL}/mayco`;
    }

    public componentDidMount(): void {
        highlightCurrentNav("liCHECK-LISTSTEP2Form");
        document.title = "Mayco - Safety | CHECKLISTSTEP2";
        document.getElementById("divDate")?.getElementsByTagName('input')[0].focus();
        this.getOnLoadData();
    }

    private getOnLoadData = async () => {
        showLoader();
        let ItemId = this.props.match.params.id;
        let showSubmit = true;
        let { getListItems } = initCommonFunctions(this.props.context, this.rootSiteURL);
        let PlantList = 'Plant';
        let DepartmentList = 'Department', DepartmentSelQuery = 'Title,Plant/Title,Plant/Id,*', DepartmentExpFields = 'Plant';
        let ZoneList = 'Zones', ZoneSelQuery = 'Title,Plant/Title,Plant/Id,Department/Title,Department/Id,*', ZoneExpFields = 'Plant,Department';
        let MachineList = 'Machines', MachineSelQuery = 'Title,Plant/Title,Plant/Id,Department/Title,Department/Id,Zone/Title,Zone/Id,*', MachineExpFields = 'Plant,Department,Zone';
        let ShiftList = 'Shifts';
        let AuditorList = 'Auditors', AuditorSelQuery = 'Title,Plant/Title,Plant/Id,*', AuditorExpFields = 'Plant';

        try {
            let [Plants, Departments, Zones, Machines, Shifts, Auditors] = await Promise.all([
                getListItems(PlantList, this.MaycoURL),
                getListItems(DepartmentList, this.MaycoURL, DepartmentSelQuery, DepartmentExpFields),
                getListItems(ZoneList, this.MaycoURL, ZoneSelQuery, ZoneExpFields),
                getListItems(MachineList, this.MaycoURL, MachineSelQuery, MachineExpFields),
                getListItems(ShiftList, this.MaycoURL),
                getListItems(AuditorList, this.MaycoURL, AuditorSelQuery, AuditorExpFields),
            ])
            let PlantsOpt = this.getMapedOptions(Plants, 'Title', 'Title');
            let ShiftsOpt = this.getMapedOptions(Shifts, 'Title', 'Title');
            // for sorting
            PlantsOpt.sort((a, b) => a.label.localeCompare(b.label));
            ShiftsOpt.sort((a, b) => a.label.localeCompare(b.label));

            let formData: any = { ...this.state.formData };
            let currPlantTitle = PlantsOpt.find((plant: any) => plant.label.toLowerCase() == this.props.currPlantTitle);
            formData.Plant = currPlantTitle ? currPlantTitle.label : '';
            this.setState({ formData, Plants, PlantsOpt, ShiftsOpt, Departments, Zones, Machines, Auditors, showSubmit });
            this.onPlantChange(this.state, formData.Plant, true); // to get current plant departments,zones,machines,Auditors
            //Edit mode starts here
            let stateData: any = { ...this.state };
            if (ItemId > 0) {
                let CheckListRes: any = await getListItems('CheckListStep2', this.currentSiteURL, 'Author/Title,Author/Id,*', 'Author', `Id eq ${ItemId}`);


                if (!CheckListRes.isHttpRequestError) {
                    if (CheckListRes.length) {
                        let CheckListData = CheckListRes[0];
                        formData.Date = [null, undefined, '', 'None'].includes(CheckListData.Date) ? '' : CheckListData.Date;
                        formData.Plant = [null, undefined, '', 'None'].includes(CheckListData.Plant) ? '' : CheckListData.Plant;
                        formData.Department = [null, undefined, '', 'None'].includes(CheckListData.Department) ? '' : CheckListData.Department;
                        formData.Zone = [null, undefined, '', 'None'].includes(CheckListData.Zone) ? '' : CheckListData.Zone;
                        formData.Machine = [null, undefined, '', 'None'].includes(CheckListData.Machine) ? '' : CheckListData.Machine;
                        formData.Shift = [null, undefined, '', 'None'].includes(CheckListData.Shifts) ? '' : CheckListData.Shifts;
                        formData.Auditor = [null, undefined, '', 'None'].includes(CheckListData.Auditor) ? '' : CheckListData.Auditor;
                        let CheckListTable = { ...stateData.CheckListTable };
                        let AuditData = [null, undefined, ''].includes(CheckListData.Audit_x0020_Data) ? {} : JSON.parse(CheckListData.Audit_x0020_Data);
                        //Binding check list table score and comments :Start
                        const flatList: any[] = [];

                        // Flatten checklist data while keeping reference to the original objects
                        for (const topic in CheckListTable) {
                            CheckListTable[topic].forEach((item: any) => {
                                flatList.push(item);
                            });
                        }

                        // Loop through AuditData and update flat checklist items
                        AuditData.Data.forEach((auditItem: any, index: number) => {
                            const checklistItem = flatList[index];
                            if (!checklistItem) return; // Safety check

                            const auditScoreKey = Object.keys(auditItem).find(k => k.startsWith("AuditScore"));
                            const commentsKey = Object.keys(auditItem).find(k => k.startsWith("Comments"));

                            if (auditScoreKey) checklistItem.AuditScore = auditItem[auditScoreKey];
                            if (commentsKey) checklistItem.Comments = auditItem[commentsKey];
                        });
                        //Binding check list table score and comments :End


                        let isEditForm = true;
                        stateData.TotalAuditScore = CheckListData.Audit_x0020_Score;
                        stateData = this.onPlantChange(stateData, formData.Plant, false);
                        stateData = this.onDepartmentChange(stateData, formData.Plant, formData.Department, false);
                        stateData = this.onZoneChange(stateData, formData.Plant, formData.Department, formData.Zone, false);

                        this.setState({ formData: stateData.formData, TotalAuditScore: stateData.TotalAuditScore, CheckListTable, isEditForm, ItemId, showSubmit, DepartmentsOpt: stateData.DepartmentsOpt, ZonesOpt: stateData.ZonesOpt, MachinesOpt: stateData.MachinesOpt });
                    }
                    else {
                        showToast("error", "No Check List found");
                        this.setState({ Redirect: true, RedirectTo: 'Home' });
                    }
                }
            }
        } catch (e) {
            console.log(e);
            this.onError();
        }
        finally {
            hideLoader();
        }
    }
    //common functions
    private getMapedOptions = (ListItems: any[], labelText: string = '', valueText: string = '') => {
        return ListItems.map((item: any) => ({
            label: typeof (item) == "string" ? item : item[labelText],
            value: typeof (item) == "string" ? item : item[valueText],
        }));
    }
    //CRUD functions
    public handleSubmit = async () => {
        try {
            showLoader();
            var formData: any = { ...this.state.formData };
            var data = {
                Date: { val: formData.Date, required: true, Name: "Date", Type: ControlType.date, Focusid: "dtDate" },
                dateToday: { val: formData.Date, required: true, Name: "Date", Type: ControlType.lessthanTodayDate, Focusid: "dtDate" },
                Plant: { val: formData.Plant, required: true, Name: "Plant", Type: ControlType.reactSelect, Focusid: "ddlPlant" },
                Department: { val: formData.Department, required: true, Name: "Department", Type: ControlType.reactSelect, Focusid: "ddlDepartment" },
                Zone: { val: formData.Zone, required: true, Name: "Zone", Type: ControlType.reactSelect, Focusid: "ddlZone" },
                Machine: { val: formData.Machine, required: true, Name: "Machine", Type: ControlType.reactSelect, Focusid: "ddlMachine" },
                Shift: { val: formData.Shift, required: true, Name: "Shift", Type: ControlType.reactSelect, Focusid: "ddlShift" },
                Auditor: { val: formData.Auditor, required: true, Name: "Auditor", Type: ControlType.reactSelect, Focusid: 'ddlAuditor' },

            }
            let isValid = Formvalidator.FormValidation(data);
            isValid = isValid.status ? await this.CheckDuplicate() as any : isValid;
            if (isValid.status) {
                let CheckListPostObj = this.getCheckListPostData();
                this.InsertOrUpdateData(CheckListPostObj);
            }
            else {
                showToast("error", isValid.message);
                hideLoader();
            }
        }
        catch (e) {
            console.log(e);
            this.onError();
        }
    }
    private InsertOrUpdateData = async (CheckListPostObj: any) => {
        let { addListItem, updateListItem } = initCommonFunctions(this.props.context, this.rootSiteURL);
        let ItemId = this.props.match.params.id;
        if (ItemId > 0) {
            await updateListItem('CheckListStep2', this.currentSiteURL, CheckListPostObj, ItemId).then(async (res: any) => {
                if (!res.isHttpRequestError) {
                    let msg = "Check List Step - 2 updated successfully";
                    this.onSuccess(msg);
                }
            }, (error: any) => {
                console.log(error);
                this.onError();
            })
        }
        else {
            await addListItem('CheckListStep2', this.currentSiteURL, CheckListPostObj).then(async (res: any) => {
                //let adedItemId = res.Id;
                if (!res.isHttpRequestError) {
                    let msg = "Check List Step - 2 submitted successfully";
                    this.onSuccess(msg);
                }
            }, (error: any) => {
                console.log(error);
                this.onError();
            })
        }
    }
    private onSuccess = (successMessage: string) => {
        hideLoader();
        this.setState({ Redirect: true, RedirectTo: 'CHECK-LISTSTEP2View', ItemID: 0 });
        showToast("success", successMessage);
    }
    private getCheckListPostData = () => {
        let stateData: any = { ...this.state };
        let CheckListPostObj: any = {};
        let CheckListDate: any = DateUtilities.addBrowserwrtServer(new Date(DateUtilities.getDateMMDDYYYY(stateData.formData.Date)), this.props.spContext.webTimeZoneData).toISOString();
        let mmddyyyyCheckListDate = format(CheckListDate, "MM/dd/yyyy");
        // Audit Data formating
        let AuditData = { "Data": [] as any };
        const CheckListTable = { ...stateData.CheckListTable };
        for (const Topic in CheckListTable) {
            const topicKey = Topic as keyof typeof CheckListTable;
            // For each row in the topic section
            CheckListTable[topicKey].forEach((row: any, index: number) => {
                let ScoreKey = `AuditScore${row.No + 1}`
                let CommentsKey = `Comments${row.No + 1}`
                AuditData.Data.push({ [ScoreKey]: row.AuditScore, [CommentsKey]: row.Comments })
            });
        }
        CheckListPostObj = {
            Date: CheckListDate,
            Plant: stateData.formData.Plant,
            Department: stateData.formData.Department,
            Zone: stateData.formData.Zone,
            Machine: stateData.formData.Machine,
            Shifts: stateData.formData.Shift,
            Year: mmddyyyyCheckListDate.split("/")[2],
            Month: mmddyyyyCheckListDate.split("/")[0],
            Auditor: stateData.formData.Auditor,
            Audit_x0020_Score: stateData.TotalAuditScore,
            Audit_x0020_Data: JSON.stringify(AuditData)
        }
        return CheckListPostObj;
    }
    private CheckDuplicate = async () => {
        let startDate = DateUtilities.getDateMMDDYYYY(new Date(this.state.formData.Date));
        let endDate = DateUtilities.getDateMMDDYYYY(addDays(new Date(this.state.formData.Date), 1));
        let { getListItems } = initCommonFunctions(this.props.context, this.rootSiteURL);
        let isValid = { message: '', status: true };
        let CheckListName = 'CheckListStep2', CheckFilterQuery = `Plant eq '${this.state.formData.Plant}' and Department eq '${this.state.formData.Department}' and Zone eq '${this.state.formData.Zone}' and Machine eq '${this.state.formData.Machine}' and Shifts eq '${this.state.formData.Shift}' and (Date ge '${startDate}' and Date lt '${endDate}')`;
        if (this.state.ItemId > 0)
            CheckFilterQuery += ` and Id ne ${this.state.ItemId}`;

        let CheckLists = await getListItems(CheckListName, this.currentSiteURL, '', '', CheckFilterQuery);
        if (CheckLists.length > 0) {
            isValid = { message: 'Check List already submitted.', status: false }
        }

        return isValid;
    }

    // on change functions
    private handleDropdownChange = (event: any, actionMeta: any) => {
        let name, value: any;
        let formData = this.state.formData;
        if (actionMeta != undefined) {
            name = actionMeta.name;
            value = actionMeta.action == 'clear' ? '' : event.value;
        }
        formData = { ...formData, [name]: value };// setting form control value
        if (name == 'Plant') {
            this.onPlantChange(this.state, value, true);
        }
        else if (name == 'Department') {
            this.onDepartmentChange(this.state, formData.Plant, value, true);
        }
        else if (name == 'Zone') {
            this.onZoneChange(this.state, formData.Plant, formData.Department, value, true);
        }
        this.setState({ formData });
    }
    private handleDateChange = (dateValue: any, name: any, divId: any) => {
        const formData: any = { ...this.state.formData };

        if (!([null, undefined, ''].includes(divId))) {
            var ddlElement = document.getElementById(divId);
            if (!([null, undefined, ''].includes(dateValue))) {
                ddlElement?.classList.add("active");
            }
            else {
                ddlElement?.classList.remove("active");
            }
        }
        if (![null, undefined, ''].includes(dateValue))
            dateValue = format(DateUtilities.addBrowserwrtServer(new Date(DateUtilities.getDateMMDDYYYY(dateValue)), this.props.spContext.webTimeZoneData).toISOString(), "MM/dd/yyyy");
        else
            dateValue = '';
        formData[name] = dateValue;
        this.setState({ formData });
    }
    private onPlantChange = (state: any, PlantVal: any, isOnchange: boolean) => {
        let stateData: any = { ...state };
        let filteredDepts = stateData.Departments.filter((dept: any) => dept.Plant?.Title == PlantVal);
        let filteredAuditors = stateData.Auditors.filter((audit: any) => audit.Plant?.Title == PlantVal && audit.Is_x0020_Active);
        let DepartmentsOpt = filteredDepts.map((item: any) => ({
            label: item.Title,
            value: item.Title,
        }));
        let AuditorsOpt = filteredAuditors.map((item: any) => ({
            label: item.Title,
            value: item.Title,
        }));
        // update dependents
        DepartmentsOpt.sort((a: any, b: any) => a.label.localeCompare(b.label));
        AuditorsOpt.sort((a: any, b: any) => a.label.localeCompare(b.label));
        stateData.DepartmentsOpt = DepartmentsOpt;
        stateData.AuditorsOpt = AuditorsOpt;
        if (isOnchange) {
            stateData.ZonesOpt = [];
            stateData.MachinesOpt = [];
            stateData.formData.Department = '';
            stateData.formData.Zone = '';
            stateData.formData.Machine = '';
            stateData.formData.Auditor = '';
            this.setState(stateData);
        }
        return stateData;

    }
    private onDepartmentChange = (state: any, PlantVal: any, DepartmentVal: any, isOnchange: boolean) => {
        let stateData: any = { ...state };
        let filteredZones = stateData.Zones.filter((zone: any) => zone.Plant?.Title == PlantVal && zone.Department?.Title == DepartmentVal);
        let ZonesOpt = filteredZones.map((item: any) => ({
            label: item.Title,
            value: item.Title,
        }));
        // update dependents
        ZonesOpt.sort((a: any, b: any) => a.label.localeCompare(b.label));
        stateData.ZonesOpt = ZonesOpt;
        if (isOnchange) {
            stateData.MachinesOpt = [];
            stateData.formData.Zone = '';
            stateData.formData.Machine = '';
            this.setState(stateData);
        }
        return stateData;
    }
    private onZoneChange = (state: any, PlantVal: any, DepartmentVal: any, ZoneVal: any, isOnchange: boolean) => {
        let stateData: any = { ...state };
        let filteredMachines = stateData.Machines.filter((mch: any) => mch.Plant?.Title == PlantVal && mch.Department?.Title == DepartmentVal && mch.Zone?.Title == ZoneVal);
        let MachinesOpt = filteredMachines.map((item: any) => ({
            label: item.Title,
            value: item.Title,
        }));
        // update dependents
        MachinesOpt.sort((a: any, b: any) => a.label.localeCompare(b.label));
        stateData.MachinesOpt = MachinesOpt;
        if (isOnchange) {
            stateData.formData.Machine = '';
            this.setState(stateData);
        }
        return stateData;
    }
    private handleTableChange = (Topic: any, index: any, field: any, value: any) => {
        const CheckListTableData = { ...this.state.CheckListTable } as any;
        CheckListTableData[Topic][index] =
        {
            ...CheckListTableData[Topic][index],
            [field]: value
        };
        let [TotalAuditScore, TotalAuditScorePercent] = [0, this.state.TotalAuditScore];
        if (field == 'AuditScore') {
            for (let Topic in CheckListTableData) {
                const topicKey = Topic as keyof typeof CheckListTableData;
                CheckListTableData[topicKey].forEach((row: any, index: number) => {
                    if (![null, undefined, ''].includes(row.AuditScore)) {
                        TotalAuditScore += Number(row.AuditScore);
                    }
                })
            }
            TotalAuditScorePercent = ((TotalAuditScore / 59) * 100).toFixed(2);
        }

        this.setState({ CheckListTable: CheckListTableData, TotalAuditScore: TotalAuditScorePercent });
    }
    //Check List Table functions
    private bindCheckListTableBody = () => {
        // Define the body structure based on the table's layout.
        const bodyStructure = { ...this.state.CheckListTable };
        // Dynamically build the body rows
        let BodyRows: JSX.Element[] = [];
        for (const Topic in bodyStructure) {
            const topicKey = Topic as keyof typeof bodyStructure;

            // For each row in the topic section
            bodyStructure[topicKey].forEach((row, index) => {
                // Dynamically create table row
                BodyRows.push(
                    <tr key={`BodyRow${row.No + 1}`} className="BodyRows">
                        {index == 0 ? <td rowSpan={bodyStructure[topicKey].length} className="TopicCell text-center">{topicKey}</td> : ''}
                        <td className="text-center">{row.No + 1}</td>
                        <td >{row.Checks}</td>
                        <td className="text-center">{row.Low}</td>
                        <td className="text-center">{row.Medium}</td>
                        <td className="text-center">{row.High}</td>
                        <td className="ScoreCell">
                            <select value={row.AuditScore} onChange={(e) => this.handleTableChange(topicKey, index, 'AuditScore', e.target.value)}>
                                <option value=''>None</option>
                                {row.Score.map((score, scoreIndex) => (
                                    <option key={scoreIndex} value={score}>
                                        {score}
                                    </option>
                                ))}
                            </select>
                        </td>
                        <td>
                            <textarea className="form-control bs-textarea" rows={2} id={`${row.No}_Comments`} name="Comments" placeholder="Comments" value={row.Comments} onChange={(e) => this.handleTableChange(topicKey, index, 'Comments', e.target.value)} disabled={false} title={row.Comments}></textarea>

                        </td>
                    </tr>
                );
            });
        }

        return BodyRows;
    }
    private bindCheckListTableHeader = () => {
        // Define the header structure based on the table's layout.
        const headerStructure = [
            //1st heading
            [
                { text: 'Check-List', rowspan: 2, colspan: 3, classeNames: 'text-center Heading1' },
                { text: '' },
                { text: '' },
                { text: 'Audit Score %', rowspan: undefined, colspan: 3, classeNames: 'text-end Heading2' },
                { text: '' },
                { text: '' },
                { text: this.state.TotalAuditScore, rowspan: undefined, colspan: 2, classeNames: 'AuditScoreCell' },
                { text: '' },
            ],
            //2nd heading
            [{ text: '' },
            { text: '' },
            { text: '' },
            { text: 'Rating Standards', rowspan: undefined, colspan: 5, classeNames: 'text-center Heading1' },
            { text: '' },
            { text: '' },
            { text: '' },
            { text: '' },
            ],
            //3rd heading
            [{ text: 'Topic', rowspan: 2, colspan: undefined, classeNames: 'text-center Heading2 WPercent-15' },
            { text: 'No.', rowspan: 2, colspan: undefined, classeNames: 'text-center Heading2 WPercent-5' },
            { text: 'Checks', rowspan: 2, colspan: undefined, classeNames: 'text-center Heading2 WPercent-20' },
            { text: 'Rating Criteria', rowspan: undefined, colspan: 3, classeNames: 'text-center Heading2 WPercent-15' },
            { text: '' },
            { text: '' },
            { text: 'Score', rowspan: 2, colspan: undefined, classeNames: 'text-center Heading2 WPercent-10' },
            { text: 'Evidence, Comments', rowspan: 2, colspan: undefined, classeNames: 'text-center Heading2 WPercent-35' }],
            //4th heading
            [
                { text: '' },
                { text: '' },
                { text: '' },
                { text: 'Low', rowspan: undefined, colspan: undefined, classeNames: 'text-center Low' },
                { text: 'Medium', rowspan: undefined, colspan: undefined, classeNames: 'text-center Medium' },
                { text: 'High', rowspan: undefined, colspan: undefined, classeNames: 'text-center High' },
                { text: '' },
                { text: '' },

            ]
        ];
        // Dynamically build the header rows
        let HeaderRows: JSX.Element[] = [];

        headerStructure.forEach((row: any, index: number) => {
            HeaderRows.push(
                <tr key={`HeadRow${index}`} className="HeadRow">
                    {headerStructure[index].map((cell, index) => {
                        // Add appropriate colspan and rowspan to the cell
                        {
                            return cell.text != '' ?
                                <th key={index} rowSpan={cell.rowspan} colSpan={cell.colspan} className={cell.classeNames}>{cell.text}</th> :
                                ''
                        }
                    })}
                </tr>
            );
        })
        return HeaderRows;
    }
    private onError = () => {
        showToast("error", ActionStatus.Error);
        hideLoader();
    }
    private handlCancel = () => {
        this.setState({ Redirect: true, RedirectTo: 'CHECK-LISTSTEP2View', ItemId: 0 });
    }
    public render() {
        if (this.state.Redirect) {
            let url = `/${this.state.RedirectTo}`;
            return (<Navigate to={url} />)
        }
        else {
            return (
                <React.Fragment>
                    <div className="container-fluid">
                        <div className="light-box border-box-shadow">
                            <div className="div-form-title">
                                <div className="form-title">{" Mayco International - WCM Safety Pillar Step Audits - Level 2 Form " + (this.state.isEditForm ? (" - " + this.state.ItemId) : "")} </div>
                                <span className="span-mandatory-text"> <span className="text-danger">* </span> are mandatory fields</span>
                            </div>
                            <div className="">
                                <div className="greenborder">
                                    <div className="form-border-box p-2 mx-3 my-2">
                                        <div className="row py-2">
                                            <div className="col-md-3" id="divDate">
                                                <div className="light-text">
                                                    <label className=""> Date <span className="mandatoryhastrick">*</span></label>
                                                    <div className="custom-datepicker" id="divCompletedDate">
                                                        <DatePickercontrol placeholder="MM/DD/YYYY" selectedDate={this.state.formData.Date} title={this.state.formData.Date} id='dtDate' isDisabled={false} startDate={undefined} endDate={new Date()} name="Date" onDatechange={(dateProps: any) => this.handleDateChange(dateProps[0], dateProps[2], "divDate")} highlightDate={new Date()} showIcon />
                                                    </div>
                                                </div>
                                            </div>
                                            <div className="col-md-3 " title={this.state.formData.Plant}>
                                                <div className="custom-dropdown" id="divPlant">
                                                    <SearchableDropdown
                                                        label={"Plant"}
                                                        Title={"Plant"}
                                                        name={"Plant"}
                                                        id="ddlPlant"
                                                        placeholderText={""}
                                                        className={""}
                                                        selectedValue={this.state.formData.Plant}
                                                        OptionsList={this.state.PlantsOpt}
                                                        OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta) }}
                                                        isRequired={true}
                                                        disabled={true}
                                                        noOptionsMessage="No Plants available"
                                                    />
                                                </div>
                                            </div>
                                            <div className="col-md-3" title={this.state.formData.Department}>
                                                <div className="custom-dropdown" id="divDepartment">
                                                    <SearchableDropdown
                                                        label={"Department"}
                                                        Title={"Department"}
                                                        name={"Department"}
                                                        id="ddlDepartment"
                                                        placeholderText={""}
                                                        className={""}
                                                        selectedValue={this.state.formData.Department}
                                                        OptionsList={this.state.DepartmentsOpt}
                                                        OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta) }}
                                                        isRequired={true}
                                                        disabled={false}
                                                        noOptionsMessage="No Departments available"
                                                    />
                                                </div>
                                            </div>
                                            <div className="col-md-3" title={this.state.formData.Zone}>
                                                <div className="custom-dropdown" id="divZone">
                                                    <SearchableDropdown
                                                        label={"Zone"}
                                                        Title={"Zone"}
                                                        name={"Zone"}
                                                        id="ddlZone"
                                                        placeholderText={""}
                                                        className={""}
                                                        selectedValue={this.state.formData.Zone}
                                                        OptionsList={this.state.ZonesOpt}
                                                        OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta) }}
                                                        isRequired={true}
                                                        disabled={false}
                                                        noOptionsMessage="No Zones available"
                                                    />
                                                </div>
                                            </div>
                                        </div>
                                        <div className="row pb-3">
                                            <div className="col-md-3" title={this.state.formData.Machine}>
                                                <div className="custom-dropdown" id="divMachine">
                                                    <SearchableDropdown
                                                        label={"Machine"}
                                                        Title={"Machine"}
                                                        name={"Machine"}
                                                        id="ddlMachine"
                                                        placeholderText={""}
                                                        className={""}
                                                        selectedValue={this.state.formData.Machine}
                                                        OptionsList={this.state.MachinesOpt}
                                                        OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta) }}
                                                        isRequired={true}
                                                        disabled={false}
                                                        noOptionsMessage="No Machines available"
                                                    />
                                                </div>
                                            </div>
                                            <div className="col-md-3" title={this.state.formData.Shift}>
                                                <div className="custom-dropdown" id="divShift">
                                                    <SearchableDropdown
                                                        label={"Shift"}
                                                        Title={"Shift"}
                                                        name={"Shift"}
                                                        id="ddlShift"
                                                        placeholderText={""}
                                                        className={""}
                                                        selectedValue={this.state.formData.Shift}
                                                        OptionsList={this.state.ShiftsOpt}
                                                        OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta) }}
                                                        isRequired={true}
                                                        disabled={false}
                                                        noOptionsMessage="No Shifts available"
                                                    />
                                                </div>
                                            </div>
                                            <div className="col-md-3" title={this.state.formData.Auditor}>
                                                <div className="custom-dropdown" id="divAuditor">
                                                    <SearchableDropdown
                                                        label={"Auditor"}
                                                        Title={"Auditor"}
                                                        name={"Auditor"}
                                                        id="ddlAuditor"
                                                        placeholderText={""}
                                                        className={""}
                                                        selectedValue={this.state.formData.Auditor}
                                                        OptionsList={this.state.AuditorsOpt}
                                                        OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta) }}
                                                        isRequired={true}
                                                        disabled={false}
                                                        noOptionsMessage="No Auditors available"
                                                    />
                                                </div>
                                            </div>
                                        </div>
                                        {/* Check List Table */}
                                        <div className="col-12">
                                            <div className="divCheckListTable form-border-box px-0 my-2">
                                                <table className="CheckListTable col-md-12" id="tblCheckList">
                                                    <tbody className="tbodyHeadRows">
                                                        {this.bindCheckListTableHeader()}
                                                    </tbody>

                                                    <tbody>
                                                        {this.bindCheckListTableBody()}
                                                    </tbody>
                                                </table>
                                            </div>
                                        </div>
                                        {/* Buttons */}
                                        <div className="col-sm-12 text-center py-3" id="divButtons" >
                                            {this.state.showSubmit && <button type="button" id="btnSubmit" className="btn btn-primary mx-2" title={this.state.ItemId > 0 ? 'Update' : 'Submit'} onClick={this.handleSubmit} >{this.state.ItemId > 0 ? 'Update' : 'Submit'}</button>}
                                            <button type="button" id="btnCancel" className="btn btn-secondary" title="Cancel" onClick={this.handlCancel}>Cancel</button>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </React.Fragment>

            )
        }
    }

}