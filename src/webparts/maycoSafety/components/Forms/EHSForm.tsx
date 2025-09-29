import * as React from "react";
import { hideLoader, showLoader } from "../Shared/Loader";
import { Navigate } from "react-router-dom";
import { SPHttpClient } from "@microsoft/sp-http";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "../CSS/EHSForm.css";
import { ActionStatus, ControlType } from "../Constants/Contants";
import { showToast } from "../Shared/Toaster";
import { highlightCurrentNav } from "../Utilities/HighlightCurrentComponent";
import DatePickercontrol from "../Shared/DatePickerField";
import DateUtilities from "../Utilities/DateUtilities";
import { format } from "date-fns";
import SearchableDropdown from "../Shared/Dropdown";
import { initCommonFunctions } from "../Utilities/CommonFunctions";
import MultipleImageUploader from "../Shared/MutipleImageUploader";
import formValidation from "../Utilities/FormValidator";
import { createBatch } from "@pnp/sp/batching";

export interface EHSFormProps {
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

export interface EHSFormState {
}

export default class EHSForm extends React.Component<EHSFormProps, EHSFormState> {

    private sp = spfi().using(SPFx(this.props.context));
    private MaycoURL: string;
    private currPlantObj: any;
    private EHSList = "EHS";
    private EHSChildList = "EHSLine";

    private txtComments: any;
    public state = {
        formData: {
            Date: '',
            Plant: '',
            Department: '',
            Zone: '',
            Machine: '',
            Shifts: '',
            ToolNo: '',
            Comments: '',
            AuditorName: '',
            Supervisor: '',
            WorkCell: '',
            unsafeactCount: '',
            unsafeconditionCount: '',
            Year: '',
            YearMonth: ''
        },
        childFormData: {
            EHSId: 0,
            EHSCategory: '',
            EHSSubCategory: '',
            ImageAttachement: '',
            Status: ''
        },
        plantsData: [],
        departmentData: [],
        departmentOptions: [],
        zoneData: [],
        zoneOptions: [],
        machineData: [],
        machineOptions: [],
        shiftData: [],
        auditorNameData: [],
        workCellData: [],
        workCellOptions: [],
        ToolNosData: [],
        ToolNosOptions: [],
        supervisorsData: [],
        allMappingData: [
            {
                EHSID: '',
                EHSCategory: '',
                EHSSubCategory: '',
                ImageAttachement: '',
                Status: ''
            }
        ],
        activeAuditCategoriesData: [],
        auditCategoryStatusData: [],
        Redirect: false,
        RedirectTo: '',
        isEditForm: false,
        ItemId: 0,
        isInputDisabled: false,
        displayMessage: '',
        showSubmit: false
    }

    constructor(props: EHSFormProps) {
        super(props);
        this.MaycoURL = `${this.props.siteURL}/mayco`;

        this.txtComments = React.createRef();
    }

    public componentDidMount(): void {
        highlightCurrentNav("liEHSForm");
        document.title = "Mayco - Safety | EHS";
        this.getOnLoadData();
    }

    private getOnLoadData = async () => {
        try {
            showLoader();
            var formData = { ...this.state.formData };
            let itemId = this.props.match.params.id;
            let showSubmit = true;

            let { getListItems } = initCommonFunctions(this.props.context, this.props.siteURL);
            let PlantList = 'Plant', PlantSelQuery = 'Title,*', plantFiltQuery = '', PlantExpFields = '';
            let DepartmentList = 'Department', DepartmentSelQuery = 'Title,Plant/Title,Plant/Id,*', DepartmentFiltQuery = '', DepartmentExpFields = 'Plant';
            let ZoneList = 'Zones', ZoneSelQuery = 'Title,Plant/Title,Plant/Id,Department/Title,Department/Id,*', ZoneFiltQuery = '', ZoneExpFields = 'Plant,Department';
            let MachineList = 'Machines', MachineSelQuery = 'Title,Plant/Title,Plant/Id,Department/Title,Department/Id,Zone/Title,Zone/Id,*', MachineFiltQuery = '', MachineExpFields = 'Plant,Department,Zone';
            let ShiftsList = 'Shifts', ShiftsSelQuery = 'Title,*', ShiftsFiltQuery = '', ShiftsExpFields = '';
            let ToolNoList = 'Tool Numbers', ToolNoSelQuery = 'Title,Plant/Title,Plant/Id,Department/Title,Department/Id,Zone/Title,Zone/Id,*', ToolNoFiltQuery = '', ToolNoExpFields = 'Plant,Department,Zone';
            let supervisorList = 'Supervisor', supervisorSelQuery = 'Title,Plant/Title,Plant/Id,*', supervisorFiltQuery = 'Is_x0020_Active eq 1', supervisorExpFields = 'Plant';
            let auditorsList = 'Auditors', auditorsSelQuery = 'Title,Plant/Title,Plant/Id,*', auditorsFiltQuery = 'Is_x0020_Active eq 1', auditorsExpFields = 'Plant';
            let workCellList = 'WorkCells', workCellSelQuery = 'Title,*', workCellFiltQuery = '', workCellExpFields = '';
            let mappingList = 'WCC/EHS mapping screen', mappingSelQuery = 'Audit_categories/Id,Audit_categories/Title,*', mappingFiltQuery = "Form_x0020_Type eq 'EHS' and Is_x0020_Active eq 1", mappingExpFields = 'Audit_categories';
            let auditCategoriesList = 'Audit_Categories', auditCategoriesSelQuery = 'Title,*', auditCategoriesFiltQuery = 'Is_x0020_Active eq 1', auditCategoriesExpFields = '';
            let requirementsList = 'RequirementSelection', requirementsSelQuery = 'Title,*', requirementsFiltQuery = '', requirementsExpFields = '';
            let [Plants, departmentData, zoneData, machineData, shifts, ToolNosData, supervisorsData, auditorNameData, workCellData, mappingData, activeAuditCategoriesData, auditCategoryStatusData] = await Promise.all([
                getListItems(PlantList, this.MaycoURL, PlantSelQuery, PlantExpFields, plantFiltQuery),
                getListItems(DepartmentList, this.MaycoURL, DepartmentSelQuery, DepartmentExpFields, DepartmentFiltQuery),
                getListItems(ZoneList, this.MaycoURL, ZoneSelQuery, ZoneExpFields, ZoneFiltQuery),
                getListItems(MachineList, this.MaycoURL, MachineSelQuery, MachineExpFields, MachineFiltQuery),
                getListItems(ShiftsList, this.MaycoURL, ShiftsSelQuery, ShiftsExpFields, ShiftsFiltQuery),
                getListItems(ToolNoList, this.MaycoURL, ToolNoSelQuery, ToolNoExpFields, ToolNoFiltQuery),
                getListItems(supervisorList, this.MaycoURL, supervisorSelQuery, supervisorExpFields, supervisorFiltQuery),
                getListItems(auditorsList, this.MaycoURL, auditorsSelQuery, auditorsExpFields, auditorsFiltQuery),
                getListItems(workCellList, this.MaycoURL, workCellSelQuery, workCellExpFields, workCellFiltQuery),
                getListItems(mappingList, this.props.webAbsoluteURL, mappingSelQuery, mappingExpFields, mappingFiltQuery),
                getListItems(auditCategoriesList, this.props.webAbsoluteURL, auditCategoriesSelQuery, auditCategoriesExpFields, auditCategoriesFiltQuery),
                getListItems(requirementsList, this.props.webAbsoluteURL, requirementsSelQuery, requirementsExpFields, requirementsFiltQuery)
            ]);

            Plants.sort((a, b) => a.Title.localeCompare(b.Title));
            departmentData.sort((a, b) => a.Title.localeCompare(b.Title));
            zoneData.sort((a, b) => a.Title.localeCompare(b.Title));
            machineData.sort((a, b) => a.Title.localeCompare(b.Title));
            ToolNosData.sort((a, b) => a.Title.localeCompare(b.Title));
            supervisorsData.sort((a, b) => a.Title.localeCompare(b.Title));
            auditorNameData.sort((a, b) => a.Title.localeCompare(b.Title));
            workCellData.sort((a, b) => a.Title.localeCompare(b.Title));
            mappingData.sort((a, b) => a.Audit_categories.Title.localeCompare(b.Audit_categories.Title));
            auditCategoryStatusData.sort((a, b) => a.Title.localeCompare(b.Title));

            let plantsData = Plants.map((item: any) => ({ label: item.Title, value: item.Title }));
            this.currPlantObj = plantsData.find((plant: any) => plant.label.toLowerCase() == this.props.currPlantTitle);
            formData.Plant = this.currPlantObj.label;

            let departmentOptions = departmentData.filter((option: any) => option.Plant.Title == formData.Plant).map((item: any) => ({ label: item.Title, value: item.Title, id: item.Id }));
            supervisorsData = supervisorsData.filter((option: any) => option.Plant?.Title == formData.Plant).map((item: any) => ({ label: item.Title, value: item.Title, id: item.Id }));
            auditorNameData = auditorNameData.filter((option: any) => option.Plant.Title == formData.Plant).map((item: any) => ({ label: item.Title, value: item.Title, id: item.Id }));
            let zoneOptions: any = [];
            let machineOptions: any = [];
            let ToolNosOptions: any = [];
            let workCellOptions: any = [];
            let shiftData = shifts.map((item: any) => ({ label: item.Title, value: item.Title }));

            let allMappingData;

            if (itemId > 0) {
                let EHSRes: any, editEHSChildItems: any;
                [EHSRes, editEHSChildItems] = await Promise.all([
                    getListItems(this.EHSList, this.props.webAbsoluteURL, 'Author/Title,Author/Id,*', 'Author', `Id eq ${itemId}`),
                    this.sp.web.lists.getByTitle(this.EHSChildList).items.top(2000).filter("EHSID eq " + itemId + "")()
                ]);
                if (!EHSRes.isHttpRequestError) {
                    if (EHSRes.length) {
                        let editEHSItem = EHSRes[0];
                        formData.Date = editEHSItem.Date ?? '',
                            formData.Plant = editEHSItem.Plant ?? '',
                            formData.Department = editEHSItem.Department ?? '',
                            formData.Zone = editEHSItem.Zone ?? '',
                            formData.Machine = editEHSItem.Machine ?? '',
                            formData.Shifts = editEHSItem.Shifts ?? '',
                            formData.ToolNo = editEHSItem.ToolNo ?? '',
                            formData.Comments = editEHSItem.Comments ?? '',
                            formData.AuditorName = editEHSItem.AuditorName ?? '',
                            formData.Supervisor = editEHSItem.Supervisor ?? '',
                            formData.WorkCell = editEHSItem.WorkCell ?? '',
                            formData.unsafeactCount = editEHSItem.unsafeactCount ?? '',
                            formData.unsafeconditionCount = editEHSItem.unsafeconditionCount ?? '',
                            formData.Year = editEHSItem.Year ?? '',
                            formData.YearMonth = editEHSItem.YearMonth ?? ''

                        zoneOptions = zoneData.filter((option: any) => (option.Plant && option.Plant.Title == formData.Plant && option.Department && option.Department.Title == editEHSItem.Department)).map((item: any) => ({ label: item.Title, value: item.Title, id: item.Id }));

                        machineOptions = machineData.filter((option: any) => (option.Plant && option.Plant.Title == formData.Plant && option.Department && option.Department.Title == formData.Department && option.Zone && option.Zone.Title == editEHSItem.Zone)).map((item: any) => ({ label: item.Title, value: item.Title, id: item.Id }));

                        ToolNosOptions = ToolNosData.filter((option: any) => (option.Plant && option.Plant.Title == formData.Plant && option.Department && option.Department.Title == formData.Department && option.Zone && option.Zone.Title == editEHSItem.Zone)).map((item: any) => ({ label: item.Title, value: item.Title, id: item.Id }));

                        workCellOptions = workCellData.filter((option: any) => (option.h7kc == formData.Plant && option.Department == formData.Department && option.Zone == editEHSItem.Zone)).map((item: any) => ({ label: item.Title, value: item.Title, id: item.Id }));

                        if (editEHSChildItems.length > 0) {
                            allMappingData = editEHSChildItems.map((item: any) => {
                                return (
                                    {
                                        Id: item.Id,
                                        EHSID: item.EHSID ?? '',
                                        EHSCategory: item.EHSCategory ?? '',
                                        EHSSubCategory: item.EHSSubCategory ?? '',
                                        ImageAttachement: item.ImageAttachement ?? '',
                                        Status: item.Status ?? '',
                                    })
                            })
                        }
                        else {
                            allMappingData = mappingData.map((item: any) => {
                                if (item.Audit_categories && item.Audit_SubCategory && this.checkAuditCategory(activeAuditCategoriesData, item.Audit_categories.Title)) {
                                    return (
                                        {
                                            EHSID: '',
                                            EHSCategory: item.Audit_categories.Title,
                                            EHSSubCategory: item.Audit_SubCategory,
                                            ImageAttachement: '',
                                            Status: 'Satisfactory'
                                        })
                                }
                                else {
                                    return null;
                                }
                            }).filter(mapItem => mapItem != null);
                        }
                        showSubmit = (editEHSItem.Author == this.props.userDisplayName || this.props.isSuperAdmin) ? true : false;
                    }
                    else {
                        showToast("error", "No EHS found");
                        this.setState({ Redirect: true, RedirectTo: 'Home' });
                    }
                }
            }
            else {
                allMappingData = mappingData.map((item: any) => {
                    if (item.Audit_categories && item.Audit_SubCategory && this.checkAuditCategory(activeAuditCategoriesData, item.Audit_categories.Title)) {
                        return (
                            {
                                EHSID: '',
                                EHSCategory: item.Audit_categories.Title,
                                EHSSubCategory: item.Audit_SubCategory,
                                ImageAttachement: '',
                                Status: 'Satisfactory'
                            })
                    }
                    else {
                        return null;
                    }
                }).filter(mapItem => mapItem != null);
            }

            this.setState({ formData, plantsData, departmentData, departmentOptions, zoneData, zoneOptions, machineData, machineOptions, shiftData, ToolNosData, ToolNosOptions, supervisorsData, auditorNameData, workCellData, workCellOptions, allMappingData, activeAuditCategoriesData, auditCategoryStatusData, showSubmit, ItemId: itemId });
        } catch (e) {
            console.log(e);
            this.onError();
        }
        finally {
            hideLoader();
        }
    }

    private checkAuditCategory(activeAuditCategoriesData: any, category: any) {
        var isExist = false;
        // const activeAuditCategoriesData = [...this.state.activeAuditCategoriesData];
        for (var ac in activeAuditCategoriesData) {
            const categoryRecord: any = activeAuditCategoriesData[ac];
            var title = categoryRecord.Title;
            if (title == category) {
                isExist = true;
                break;
            }
        }
        return isExist;
    }

    private handleChange = (event: any) => {
        const formData: any = { ...this.state.formData };

        const name = event.target.name;
        let inputValue = (event.target.type == "text" || event.target.type == "textarea") ? event.target.value : (event.target.type == "checkbox" ? event.target.checked : event.target.value ? event.target.value : '');
        let classArr: any = event.target.className;

        if (classArr.includes("onlyNum")) {
            inputValue = inputValue.replace(/[^0-9]/g, '');
        }
        formData[name] = inputValue;
        this.setState({ formData });
    }

    private handleDropdownChange = (event: any, actionMeta: any, id: any) => {
        const formData: any = { ...this.state.formData };
        let departmentData = [...this.state.departmentData];
        let departmentOptions: any = [...this.state.departmentOptions];
        let zoneData = [...this.state.zoneData];
        let zoneOptions: any = [...this.state.zoneOptions];
        let machineData = [...this.state.machineData];
        let machineOptions: any = [...this.state.machineOptions];
        let ToolNosData = [...this.state.ToolNosData];
        let ToolNosOptions: any = [...this.state.ToolNosOptions];
        let workCellData = [...this.state.workCellData];
        let workCellOptions: any = [...this.state.workCellOptions];
        const name = actionMeta.name;
        const value = actionMeta.action == "clear" ? '' : event.value;
        formData[name] = value;

        if (name == "Plant") {
            formData.Department = "";
            formData.Zone = "";
            formData.Machine = "";
            departmentOptions = [];
            machineOptions = [];
            zoneOptions = [];

            if (actionMeta.action != "clear") {
                departmentOptions = departmentData.filter((option: any) => (option.Plant && option.Plant.Title == formData.Plant && option.Department)).map((item: any) => ({ label: item.Title, value: item.Title, id: item.Id }));
            }
        }
        else if (name == "Department") {
            formData.Zone = "";
            formData.Machine = "";
            formData.ToolNo = "";
            formData.WorkCell = "";

            if (actionMeta.action != "clear") {
                zoneOptions = [];
                zoneOptions = zoneData.filter((option: any) => (option.Plant && option.Plant.Title == formData.Plant && option.Department && option.Department.Id == event.id)).map((item: any) => ({ label: item.Title, value: item.Title, id: item.Id }));
            }
            else { zoneOptions = []; }
            machineOptions = [];
            ToolNosOptions = [];
            workCellOptions = [];
        }
        else if (name == "Zone") {
            formData.Machine = "";
            formData.ToolNo = "";
            formData.WorkCell = "";

            if (actionMeta.action != "clear") {
                machineOptions = [];
                ToolNosOptions = [];

                machineOptions = machineData.filter((option: any) => (option.Plant && option.Plant.Title == formData.Plant && option.Department && option.Department.Title == formData.Department && option.Zone.Id == event.id)).map((item: any) => ({ label: item.Title, value: item.Title, id: item.Id }));

                ToolNosOptions = ToolNosData.filter((option: any) => (option.Plant && option.Plant.Title == formData.Plant && option.Department && option.Department.Title == formData.Department && option.Zone.Title == event.value)).map((item: any) => ({ label: item.Title, value: item.Title, id: item.Id }));

                workCellOptions = workCellData.filter((option: any) => (option.h7kc == formData.Plant && option.Department == formData.Department && option.Zone == event.value)).map((item: any) => ({ label: item.Title, value: item.Title, id: item.Id }));
            }
            else {
                machineOptions = [];
                ToolNosOptions = [];
                workCellOptions = [];
            }
        }

        if (!([null, undefined, ''].includes(id))) {
            var ddlElement = document.getElementById(id);
            if (!([null, undefined, ''].includes(value))) {
                ddlElement?.classList.add("active");
            }
            else {
                ddlElement?.classList.remove("active");
            }
        }

        this.setState({ formData, departmentOptions, zoneOptions, machineOptions, ToolNosOptions, workCellOptions });
    }

    private handleDateChange = (dateValue: any, name: any, divId: any, dateProps: any) => {
        const formData: any = { ...this.state.formData };
        console.log(dateProps);

        if (!([null, undefined, ''].includes(divId))) {
            var ddlElement = document.getElementById(divId);
            if (!([null, undefined, ''].includes(dateValue))) {
                ddlElement?.classList.add("active");
            }
            else {
                ddlElement?.classList.remove("active");
            }
        }

        if (dateValue != null) {
            dateValue = format(DateUtilities.addBrowserwrtServer(new Date(DateUtilities.getDateMMDDYYYY(dateValue)), this.props.spContext.webTimeZoneData).toISOString(), "MM/dd/yyyy");
        }
        else {
            dateValue = "";
        }

        formData[name] = dateValue;

        this.setState({ formData });
    }

    private handleSubmit = async () => {
        try {
            showLoader();
            var formData: any = { ...this.state.formData };
            // let itemId = this.props.match.params.id;
            let data = {
                date: { val: formData.Date, required: true, Name: "Date", Type: ControlType.date, Focusid: "dtDate" },
                dateToday: { val: formData.Date, required: true, Name: "Date", Type: ControlType.lessthanTodayDate, Focusid: "dtDate" },
                shift: { val: formData.Shifts, required: true, Name: "Shifts", Type: ControlType.reactSelect, Focusid: "ddlShift" },
                auditorsName: { val: formData.AuditorName, required: true, Name: "Auditor's Name", Type: ControlType.reactSelect, Focusid: "ddlAuditorName" },
                plant: { val: formData.Plant, required: true, Name: "Plant", Type: ControlType.reactSelect, Focusid: "ddlPlant" },
                department: { val: formData.Department, required: true, Name: "Department", Type: ControlType.reactSelect, Focusid: "ddlDepartment" },
                zone: { val: formData.Zone, required: true, Name: "Zone", Type: ControlType.reactSelect, Focusid: "ddlZone" },
                machine: { val: formData.Machine, required: true, Name: "Machine", Type: ControlType.reactSelect, Focusid: "ddlMachine" },
                workCell: { val: formData.WorkCell, required: true, Name: "Work Cell", Type: ControlType.reactSelect, Focusid: "ddlWorkCell" },
                ToolNo: { val: formData.ToolNo, required: (formData.Department == "Molding"), Name: "Tool Number", Type: ControlType.reactSelect, Focusid: "ddlToolNo" },
                supervisor: { val: formData.Supervisor, required: true, Name: "Supervisor", Type: ControlType.reactSelect, Focusid: "ddlSupervisor" }
            }

            let isValid = formValidation.FormValidation(data);

            if (isValid.status) {
                // console.log("Valid Data");

                //Date formatting, Year and YearMonth
                formData.Date = DateUtilities.addBrowserwrtServer(new Date(DateUtilities.getDateMMDDYYYY(formData.Date)), this.props.spContext.webTimeZoneData).toISOString();
                let mmddyyyyDate = format(formData.Date, "MM/dd/yyyy");
                // console.log(mmddyyyyDate);
                formData.Year = mmddyyyyDate.split("/")[2];
                formData.YearMonth = mmddyyyyDate.split("/")[0];

                //Get Status Count
                let allMappingData = { ...this.state.allMappingData }
                let unsafeact = 0;
                let unsafeConddition = 0;
                for (let i = 0; i < Object.keys(allMappingData).length; i++) {
                    var subCategory = allMappingData[i];
                    var status = subCategory.Status;

                    if (status.toLowerCase() == "unsafe act") {
                        unsafeact += 1;
                    }
                    if (status.toLowerCase() == "unsafe condition") {
                        unsafeConddition += 1;
                    }
                    if (status.toLowerCase() == "both") {
                        unsafeact += 1;
                        unsafeConddition += 1;
                    }
                }

                formData.unsafeactCount = unsafeact.toString();
                formData.unsafeconditionCount = unsafeConddition.toString();

                console.clear();
                console.log(formData);
                console.log(this.state.allMappingData);

                await this.InsertOrUpdateData(formData);
            }
            else {
                showToast("error", isValid.message);
            }
            hideLoader();
        } catch (e) {
            console.log(e);
            this.onError();
        }
    }

    private InsertOrUpdateData = async (formData: any) => {
        try {
            let itemId = this.props.match.params.id;
            if (itemId > 0) {
                await this.sp.web.lists.getByTitle(this.EHSList).items.getById(itemId).update(formData).then((res: any) => {
                    this.UpdateLineItems();
                }, (error) => {
                    console.log(error);
                    this.onError();
                })
            }
            else {
                await this.sp.web.lists.getByTitle(this.EHSList).items.add(formData).then((res: any) => {
                    let childObjects = this.updateEHSId(res.Id.toString());
                    this.InsertLineItems(childObjects,res.Id);
                }, (error) => {
                    console.log(error);
                    this.onError();
                })
            }
        } catch (e) {
            console.log(e);
            this.onError();
        }
    }

    private InsertLineItems = async (childPostObjects: any, adedItemId: any) => {
        try {
            let { getGroupMemberEmails, sendEmail } = initCommonFunctions(this.props.context, this.props.siteURL);

            const [batchedPipe, execute] = createBatch(this.sp.web);
            for (const item of childPostObjects) {
                this.sp.web.lists.getByTitle(this.EHSChildList).items.using(batchedPipe).add(item);
            }
            await execute().then(async () => {
                let GroupName = this.getGroupName();
                if (GroupName != '') {
                    let GroupMemberEmails = await getGroupMemberEmails(GroupName, this.props.siteURL);
                    if (GroupMemberEmails.length) {
                        let link = this.props.webAbsoluteURL + '/SitePages/Home.aspx#/EHSForm/' + adedItemId
                        let body = "<p>Hi,</p>" + "<p>New 'EHS-" + adedItemId + "' has been submitted. Please <a href='" + link + "'><b>click here</b></a> to view the details.</p><p>Regards</p>";
                        await sendEmail(this.props.siteURL, GroupMemberEmails, "New 'EHS' Submitted", body);
                    }
                }
                let msg = "EHS Submitted Successfully";
                this.setState({ displayMessage: msg });
                this.onSuccess();

            }, (error: any) => {
                console.log(error);
                this.onError();
            })
        } catch (e) {
            console.log(e);
            this.onError();
        }
    }

    private UpdateLineItems = async () => {
        try {
            let childPostObjects = this.createUpdateObjects();
            const [batchedPipe, execute] = createBatch(this.sp.web);
            for (const item of childPostObjects) {
                this.sp.web.lists.getByTitle(this.EHSChildList).items.getById(item.id).using(batchedPipe).update(item.Obj);
            }
            await execute().then(async () => {
                let msg = "EHS Updated Successfully";
                this.setState({ displayMessage: msg });
                this.onSuccess();

            }, (error: any) => {
                console.log(error);
                this.onError();
            })
        } catch (e) {
            console.log(e);
            this.onError();
        }
    }

    private createUpdateObjects() {
        const allMappingData: any = [...this.state.allMappingData];
        let postObjects = [];

        for (let i = 0; i < allMappingData.length; i++) {
            let mappingItem = allMappingData[i];

            let postObj = {
                id: Number(mappingItem.Id),
                Obj: {
                    EHSID: mappingItem.EHSID,
                    EHSCategory: mappingItem.EHSCategory,
                    EHSSubCategory: mappingItem.EHSSubCategory,
                    ImageAttachement: mappingItem.ImageAttachement,
                    Status: mappingItem.Status
                }
            }
            postObjects.push(postObj);
        }
        return postObjects;
    }

    private updateEHSId = (ParentId: string) => {
        let allMappingData = [...this.state.allMappingData];

        let updatedMappingData = allMappingData.map(item => ({
            ...item,
            EHSID: ParentId
        }));

        return updatedMappingData;
    }

    private onSuccess = () => {
        hideLoader();
        this.setState({ Redirect: true, RedirectTo: 'EHSView', ItemID: 0 });
        showToast("success", this.state.displayMessage);
    }

    private onError = () => {
        showToast("error", ActionStatus.Error);
        hideLoader();
    }
private getGroupName()
{
	var selectedDept= this.state.formData.Department;
	var selectedZone= this.state.formData.Zone;	
    let group='';
	
	if(selectedDept =="IP Assembly")
		group="WCM Merrill Safety IP Assy";//group="WCM Safety IP Assy";
	else if(selectedDept =="Sequencing")
		group="WCM Merrill Safety Seq";//group="WCM Safety Seq";
	else if(selectedDept =="Thermoforming")
		group="WCM Merrill Safety Thermo";//group="WCM Safety IPM Thermo";
	else if(selectedDept =="Deco")
		group="WCM Merrill Safety Deco";//group="WCM Safety Deco";
	else if(selectedDept =="Molding"  && selectedZone=="Zone 1")
		group="WCM Safety Molding Zone 1";
	else if(selectedDept =="Molding" && selectedZone=="Zone 2")
		group="WCM Safety Molding Zone 2";
	else if(selectedDept =="Molding" && selectedZone=="Zone 3")
		group="WCM Safety Molding Zone 3"; 	
	else if(selectedDept =="Molding" && selectedZone=="Zone 4")
		group="WCM Safety Molding Zone 4";			
	return group; 
}
    private handleCancel = () => {
        this.setState({ Redirect: true, RedirectTo: 'EHSView', ItemID: 0 });
    }

    private bindDynamicTable = () => {
        const { allMappingData, auditCategoryStatusData } = this.state;

        // Ensure unique categories by creating a Set of EHSCategory
        const uniqueCategories = Array.from(new Set(allMappingData.map(item => item.EHSCategory)));

        // Initialize globalCategoryIndex for numbering
        let globalCategoryIndex = 1;

        const tBody = uniqueCategories.map((category, categoryIndex) => {
            // Filter subcategories that match the current category
            const categoryRows = allMappingData.filter(subCategory => subCategory.EHSCategory === category).sort((a: any, b: any) => a.EHSSubCategory.localeCompare(b.EHSSubCategory));

            return (
                <React.Fragment key={categoryIndex}>
                    {/* Category Row */}
                    <tr className="lightgreybg fs-6">
                        <td colSpan={3}><b>{category}</b></td>
                    </tr>

                    {/* Subcategory Rows */}
                    {categoryRows.map((subCategory, subIndex) => {

                        const rowIndex = allMappingData.indexOf(subCategory);

                        return (
                            <tr key={rowIndex}>
                                <td>
                                    <div className="pull-left">{globalCategoryIndex++}.{subCategory.EHSSubCategory}</div>
                                    <div className="pull-right">
                                        <MultipleImageUploader
                                            onImageUpload={this.onImageChange}
                                            onRemoveImage={this.onRemoveImage}
                                            initialImageSrc={subCategory.ImageAttachement}
                                            index={rowIndex}
                                        />
                                    </div>
                                </td>
                                <td>
                                    <div className="">
                                        <select
                                            className="form-select"
                                            id={"SubCategory" + rowIndex + "Select"}
                                            onChange={(event) => this.handleStatusChange(event, rowIndex)}
                                            value={subCategory.Status} // controlled input
                                        >
                                            {auditCategoryStatusData.map((status: any) => (
                                                <option key={status.Title} value={status.Title}>
                                                    {status.Title}
                                                </option>
                                            ))}
                                        </select>
                                    </div>
                                </td>
                            </tr>
                        );
                    })}
                </React.Fragment>
            );
        });

        return tBody;
    };

    private handleStatusChange = (event: any, CategoryIndex: number) => {
        const allMappingData: any = [...this.state.allMappingData];
        const value = event.target.value;
        allMappingData[CategoryIndex].Status = value;
        console.log(allMappingData);
        this.setState({ allMappingData });
    };

    private onImageChange = (base64: string, lineIndex: number) => {
        const allMappingData: any = [...this.state.allMappingData];
        allMappingData[lineIndex].ImageAttachement = base64;

        this.setState({ allMappingData });
    };

    private onRemoveImage = (lineIndex: number) => {
        const allMappingData: any = [...this.state.allMappingData];
        allMappingData[lineIndex].ImageAttachement = '';

        this.setState({ allMappingData });
    };

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
                            <div className="m-0 titlebg">
                                <h3 className="mb-0 pt-2 text-center">{" EHS " + (this.state.isEditForm ? (" - " + this.state.ItemId) : "")} </h3>
                                <label className="text-end px-1" style={{ width: "100%" }}> <span className="text-danger">* </span> are mandatory fields</label>
                            </div>

                            <div className="mainContent borderLine">
                                <div className="row py-2">
                                    {/* Date */}
                                    <div className="col-md-3">
                                        <div className="light-text">
                                            <label className="" htmlFor="dtDate"> Date <span className="text-danger">*</span></label>
                                            <div className="custom-datepicker" id="divDate">
                                                <DatePickercontrol placeholder="" selectedDate={this.state.formData.Date} id='dtDate' isDisabled={this.state.isInputDisabled} startDate={undefined} endDate={undefined} name="Date" onDatechange={(dateProps: any) => this.handleDateChange(dateProps[0], dateProps[2], "divDate", dateProps)} highlightDate={new Date()} showIcon />
                                            </div>
                                        </div>
                                    </div>
                                    {/* Shift */}
                                    <div className="col-md-3">
                                        <div className="custom-dropdown" id="divShift">
                                            <SearchableDropdown
                                                label={"Shift"}
                                                Title={"Shift"}
                                                name={"Shifts"}
                                                id="ddlShift"
                                                placeholderText={""}
                                                className={""}
                                                selectedValue={this.state.formData.Shifts}
                                                OptionsList={this.state.shiftData}
                                                OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta, "divShift") }}
                                                isRequired={true}
                                                disabled={this.state.isInputDisabled}
                                                noOptionsMessage="No Shifts"
                                            />
                                        </div>
                                    </div>
                                    {/* Auditor's Name */}
                                    <div className="col-md-3">
                                        <div className="custom-dropdown active" id="divAuditorName" title={this.state.formData.AuditorName}>
                                            <SearchableDropdown
                                                label={"Auditor's Name"}
                                                Title={"AuditorName"}
                                                name={"AuditorName"}
                                                id="ddlAuditorName"
                                                placeholderText={""}
                                                className={""}
                                                selectedValue={this.state.formData.AuditorName}
                                                OptionsList={this.state.auditorNameData}
                                                OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta, "divAuditorName") }}
                                                isRequired={true}
                                                disabled={this.state.isInputDisabled}
                                                noOptionsMessage="No Auditor's Names"
                                            />
                                        </div>
                                    </div>
                                    {/* Plant */}
                                    <div className="col-md-3">
                                        <div className="custom-dropdown active" id="divPlant" title={this.state.formData.Plant}>
                                            <SearchableDropdown
                                                label={"Plant"}
                                                Title={"Plant"}
                                                name={"Plant"}
                                                id="ddlPlant"
                                                placeholderText={""}
                                                className={""}
                                                selectedValue={this.state.formData.Plant}
                                                OptionsList={this.state.plantsData}
                                                OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta, "divPlant") }}
                                                isRequired={true}
                                                disabled={true}
                                                noOptionsMessage="No Plants available"
                                            />
                                        </div>
                                    </div>
                                </div>
                                <div className="row pb-2">
                                    {/* Department */}
                                    <div className="col-md-3">
                                        <div className="custom-dropdown" id="divDepartment">
                                            <SearchableDropdown
                                                label={"Department"}
                                                Title={"Department"}
                                                name={"Department"}
                                                id="ddlDepartment"
                                                placeholderText={""}
                                                className={""}
                                                selectedValue={this.state.formData.Department}
                                                OptionsList={this.state.departmentOptions}
                                                OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta, "divDepartment") }}
                                                isRequired={true}
                                                disabled={this.state.isInputDisabled}
                                                noOptionsMessage="No Departments"
                                            />
                                        </div>
                                    </div>
                                    {/* Zone */}
                                    <div className="col-md-3">
                                        <div className="custom-dropdown" id="divZone">
                                            <SearchableDropdown
                                                label={"Zone"}
                                                Title={"Zone"}
                                                name={"Zone"}
                                                id="ddlZone"
                                                placeholderText={""}
                                                className={""}
                                                selectedValue={this.state.formData.Zone}
                                                OptionsList={this.state.zoneOptions}
                                                OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta, "divZone") }}
                                                isRequired={true}
                                                disabled={this.state.isInputDisabled}
                                                noOptionsMessage="No Zones"
                                            />
                                        </div>
                                    </div>
                                    {/* Machine */}
                                    <div className="col-md-3">
                                        <div className="custom-dropdown" id="divMachine">
                                            <SearchableDropdown
                                                label={"Machine"}
                                                Title={"Machine"}
                                                name={"Machine"}
                                                id="ddlMachine"
                                                placeholderText={""}
                                                className={""}
                                                selectedValue={this.state.formData.Machine}
                                                OptionsList={this.state.machineOptions}
                                                OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta, "divMachine") }}
                                                isRequired={true}
                                                disabled={this.state.isInputDisabled}
                                                noOptionsMessage="No Machines"
                                            />
                                        </div>
                                    </div>
                                    {/* Work Cell */}
                                    <div className="col-md-3">
                                        <div className="custom-dropdown" id="divWorkCell" title={this.state.formData.WorkCell}>
                                            <SearchableDropdown
                                                label={"Work Cell"}
                                                Title={"WorkCell"}
                                                name={"WorkCell"}
                                                id="ddlWorkCell"
                                                placeholderText={""}
                                                className={""}
                                                selectedValue={this.state.formData.WorkCell}
                                                OptionsList={this.state.workCellOptions}
                                                OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta, "divWorkCell") }}
                                                isRequired={true}
                                                disabled={this.state.isInputDisabled}
                                                noOptionsMessage="No WorkCell"
                                            />
                                        </div>
                                    </div>
                                </div>
                                <div className="row pb-2">
                                    {/* Tool No. */}
                                    <div className="col-md-3">
                                        <div className="custom-dropdown" id="divToolNo" title={this.state.formData.ToolNo}>
                                            <SearchableDropdown
                                                label={"Tool No."}
                                                Title={"ToolNo"}
                                                name={"ToolNo"}
                                                id="ddlToolNo"
                                                placeholderText={""}
                                                className={""}
                                                selectedValue={this.state.formData.ToolNo}
                                                OptionsList={this.state.ToolNosOptions}
                                                OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta, "divToolNo") }}
                                                isRequired={this.state.formData.Department == "Molding"}
                                                disabled={this.state.isInputDisabled}
                                                noOptionsMessage="No ToolNo's"
                                            />
                                        </div>
                                    </div>
                                    {/* Supervisor */}
                                    <div className="col-md-3">
                                        <div className="custom-dropdown" id="divSupervisor" title={this.state.formData.Supervisor}>
                                            <SearchableDropdown
                                                label={"Supervisor"}
                                                Title={"Supervisor"}
                                                name={"Supervisor"}
                                                id="ddlSupervisor"
                                                placeholderText={""}
                                                className={""}
                                                selectedValue={this.state.formData.Supervisor}
                                                OptionsList={this.state.supervisorsData}
                                                OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta, "divSupervisor") }}
                                                isRequired={true}
                                                disabled={this.state.isInputDisabled}
                                                noOptionsMessage="No Supervisor's"
                                            />
                                        </div>
                                    </div>
                                </div>
                                <div className="divSection mx-3">
                                    <table>
                                        <thead className="darkgreybg">
                                            <tr className="fs-5">
                                                <th>Requirement</th>
                                                <th>Select</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            {this.bindDynamicTable()}
                                        </tbody>
                                    </table>
                                </div>
                                {/* Comments */}
                                    <div className="col-md-12">
                                        <div className="light-text" >
                                            <label className="" htmlFor="txtComments">Comments </label>
                                            <textarea className="form-control bs-textarea" rows={3} id="txtComments" name="Comments" ref={this.txtComments} placeholder="Comments" value={this.state.formData.Comments} onChange={this.handleChange} disabled={this.state.isInputDisabled} title="Comments"></textarea>
                                        </div>
                                    </div>

                                <div className="col-sm-12 text-center py-3" id="">
                                    {this.state.showSubmit && <button type="button" id="btnSubmit" className="btn btn-primary mx-2" onClick={this.handleSubmit} title={this.state.ItemId > 0 ? 'Update' : 'Submit'}>{this.state.ItemId > 0 ? 'Update' : 'Submit'}</button>}
                                    <button type="button" id="btnCancel" className="btn btn-secondary" onClick={this.handleCancel} >Cancel</button>
                                </div>
                            </div>
                        </div>
                    </div>
                </React.Fragment>
            )
        }
    }

}