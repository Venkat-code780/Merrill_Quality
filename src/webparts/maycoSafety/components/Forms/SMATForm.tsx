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
import { faHistory } from "@fortawesome/free-solid-svg-icons";
import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";
import ActionHistory from "../Shared/ActionHistory";
export interface SMATFormProps {
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

export interface SMATFormState {
}

export default class SMATForm extends React.Component<SMATFormProps, SMATFormState> {

    private sp = spfi().using(SPFx(this.props.context));
    private MaycoURL: string;
    private currPlantObj: any;
    private SMATList = "WCC";
    private SMATChildList = "WccLine";

    private txtComments: any;
    private txtActionCompleted: any;
    public state = {
        formData: {
            Plant: '',
            Department: '',
            Zone: '',
            Machine: '',
            ShiftType: '',
            ToolNumber: '',
            Comments: '',
            AuditorName: '',
            Supervisor: '',
            WorkCell: '',
            unsafeactCount: '',
            unsafeconditionCount: '',
            ActionCompleted: '',
            WCCDate: '',
            CompletedDate: '',
            Year: '',
            YearMonth: '',
            ActionHistory: [],
        },
        childFormData: {
            wccId: 0,
            WccCategory: '',
            WccSubCategory: '',
            Attachment: '',
            SubCategoryStatus: '',
            SubCategoryComments: ''
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
        toolNumbersData: [],
        toolNumbersOptions: [],
        supervisorsData: [],
        allMappingData: [
            {
                WCCID: '',
                WccCategory: '',
                WccSubCategory: '',
                Attachment: '',
                SubCategoryStatus: '',
                SubCategoryComments: '',
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

    constructor(props: SMATFormProps) {
        super(props);
        this.MaycoURL = `${this.props.siteURL}/mayco`;

        this.txtComments = React.createRef();
        this.txtActionCompleted = React.createRef();
    }

    public componentDidMount(): void {
        highlightCurrentNav("liSMATForm");
        document.title = "Mayco - Safety | SMAT";
        document.getElementById("divWCCDate")?.getElementsByTagName('input')[0].focus();
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
            let toolNumberList = 'Tool Numbers', toolNumberSelQuery = 'Title,Plant/Title,Plant/Id,Department/Title,Department/Id,Zone/Title,Zone/Id,*', toolNumberFiltQuery = '', toolNumberExpFields = 'Plant,Department,Zone';
            let supervisorList = 'Supervisor', supervisorSelQuery = 'Title,Plant/Title,Plant/Id,*', supervisorFiltQuery = 'Is_x0020_Active eq 1', supervisorExpFields = 'Plant';
            let auditorsList = 'Auditors', auditorsSelQuery = 'Title,Plant/Title,Plant/Id,*', auditorsFiltQuery = 'Is_x0020_Active eq 1', auditorsExpFields = 'Plant';
            let workCellList = 'WorkCells', workCellSelQuery = 'Title,*', workCellFiltQuery = '', workCellExpFields = '';
            let mappingList = 'WCC/EHS mapping screen', mappingSelQuery = 'Audit_categories/Id,Audit_categories/Title,*', mappingFiltQuery = "Form_x0020_Type eq 'WCC' and Is_x0020_Active eq 1", mappingExpFields = 'Audit_categories';
            let auditCategoriesList = 'Audit_Categories', auditCategoriesSelQuery = 'Title,*', auditCategoriesFiltQuery = 'Is_x0020_Active eq 1', auditCategoriesExpFields = '';
            let requirementsList = 'RequirementSelection', requirementsSelQuery = 'Title,*', requirementsFiltQuery = '', requirementsExpFields = '';
            let [Plants, departmentData, zoneData, machineData, shifts, toolNumbersData, supervisorsData, auditorNameData, workCellData, mappingData, activeAuditCategoriesData, auditCategoryStatusData] = await Promise.all([
                getListItems(PlantList, this.MaycoURL, PlantSelQuery, PlantExpFields, plantFiltQuery),
                getListItems(DepartmentList, this.MaycoURL, DepartmentSelQuery, DepartmentExpFields, DepartmentFiltQuery),
                getListItems(ZoneList, this.MaycoURL, ZoneSelQuery, ZoneExpFields, ZoneFiltQuery),
                getListItems(MachineList, this.MaycoURL, MachineSelQuery, MachineExpFields, MachineFiltQuery),
                getListItems(ShiftsList, this.MaycoURL, ShiftsSelQuery, ShiftsExpFields, ShiftsFiltQuery),
                getListItems(toolNumberList, this.MaycoURL, toolNumberSelQuery, toolNumberExpFields, toolNumberFiltQuery),
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
            toolNumbersData.sort((a, b) => a.Title.localeCompare(b.Title));
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
            let toolNumbersOptions: any = [];
            let workCellOptions: any = [];
            let shiftData = shifts.map((item: any) => ({ label: item.Title, value: item.Title }));
            let sortedData: any = [];

            let allMappingData;
            let isEditForm = false;
            if (itemId > 0) {
                let SMATRes: any, editSMATChildItems: any;
                [SMATRes, editSMATChildItems] = await Promise.all([
                    getListItems(this.SMATList, this.props.webAbsoluteURL, 'Author/Title,Author/Id,*', 'Author', `Id eq ${itemId}`),
                    this.sp.web.lists.getByTitle(this.SMATChildList).items.top(2000).filter("WCCID eq " + itemId + "")()]);

                if (!SMATRes.isHttpRequestError) {
                    if (SMATRes.length) {
                        let editSMATItem = SMATRes[0];
                        formData.Plant = editSMATItem.Plant ?? '',
                            formData.Department = editSMATItem.Department ?? '',
                            formData.Zone = editSMATItem.Zone ?? '',
                            formData.Machine = editSMATItem.Machine ?? '',
                            formData.ShiftType = editSMATItem.ShiftType ?? '',
                            formData.ToolNumber = editSMATItem.ToolNumber ?? '',
                            formData.Comments = editSMATItem.Comments ?? '',
                            formData.AuditorName = editSMATItem.AuditorName ?? '',
                            formData.Supervisor = editSMATItem.Supervisor ?? '',
                            formData.WorkCell = editSMATItem.WorkCell ?? '',
                            formData.unsafeactCount = editSMATItem.unsafeactCount ?? '',
                            formData.unsafeconditionCount = editSMATItem.unsafeconditionCount ?? '',
                            formData.ActionCompleted = editSMATItem.ActionCompleted ?? '',
                            formData.WCCDate = editSMATItem.WCCDate ?? '',
                            formData.CompletedDate = editSMATItem.CompletedDate ?? '',
                            formData.Year = editSMATItem.Year ?? '',
                            formData.YearMonth = editSMATItem.YearMonth ?? '',
                            formData.ActionHistory = editSMATItem.ActionHistory ? JSON.parse(editSMATItem.ActionHistory) : [];

                        zoneOptions = zoneData.filter((option: any) => (option.Plant && option.Plant.Title == formData.Plant && option.Department && option.Department.Title == editSMATItem.Department)).map((item: any) => ({ label: item.Title, value: item.Title, id: item.Id }));

                        machineOptions = machineData.filter((option: any) => (option.Plant && option.Plant.Title == formData.Plant && option.Department && option.Department.Title == formData.Department && option.Zone && option.Zone.Title == editSMATItem.Zone)).map((item: any) => ({ label: item.Title, value: item.Title, id: item.Id }));

                        toolNumbersOptions = toolNumbersData.filter((option: any) => (option.Plant && option.Plant.Title == formData.Plant && option.Department && option.Department.Title == formData.Department && option.Zone && option.Zone.Title == editSMATItem.Zone)).map((item: any) => ({ label: item.Title, value: item.Title, id: item.Id }));

                        workCellOptions = workCellData.filter((option: any) => (option.h7kc == formData.Plant && option.Department == formData.Department && option.Zone == editSMATItem.Zone)).map((item: any) => ({ label: item.Title, value: item.Title, id: item.Id }));

                        if (editSMATChildItems.length > 0) {
                            allMappingData = editSMATChildItems.map((item: any) => {
                                return (
                                    {
                                        Id: item.Id,
                                        WCCID: item.WCCID ?? '',
                                        WccCategory: item.WccCategory ?? '',
                                        WccSubCategory: item.WccSubCategory ?? '',
                                        Attachment: item.Attachment ?? '',
                                        SubCategoryStatus: item.SubCategoryStatus ?? '',
                                        SubCategoryComments: item.SubCategoryComments ?? ''
                                    })
                            })

                            sortedData = allMappingData.sort((a: any, b: any) => {
                                // First sort by WccCategory
                                if (a.WccCategory < b.WccCategory) return -1;
                                if (a.WccCategory > b.WccCategory) return 1;

                                // If categories are the same, then sort by WccSubCategory
                                if (a.WccSubCategory < b.WccSubCategory) return -1;
                                if (a.WccSubCategory > b.WccSubCategory) return 1;

                                return 0;
                            });
                        }
                        showSubmit = (editSMATItem.Author == this.props.userDisplayName || this.props.isSuperAdmin) ? true : false;
                        isEditForm = true;
                    }
                    else {
                        showToast("error", "No SMAT found");
                        this.setState({ Redirect: true, RedirectTo: 'Home' });
                    }
                }
            }
            else {
                allMappingData = mappingData.map((item: any) => {
                    if (item.Audit_categories && item.Audit_SubCategory && this.checkAuditCategory(activeAuditCategoriesData, item.Audit_categories.Title)) {
                        return (
                            {
                                WCCID: '',
                                WccCategory: item.Audit_categories.Title,
                                WccSubCategory: item.Audit_SubCategory,
                                Attachment: '',
                                SubCategoryStatus: 'Satisfactory',
                                SubCategoryComments: ''
                            })
                    }
                    else {
                        return null;
                    }
                }).filter(mapItem => mapItem != null);

                sortedData = allMappingData.sort((a: any, b: any) => {
                    // First sort by WccCategory
                    if (a.WccCategory < b.WccCategory) return -1;
                    if (a.WccCategory > b.WccCategory) return 1;

                    // If categories are the same, then sort by WccSubCategory
                    if (a.WccSubCategory < b.WccSubCategory) return -1;
                    if (a.WccSubCategory > b.WccSubCategory) return 1;

                    return 0;
                });
            }

            this.setState({ formData, plantsData, departmentData, departmentOptions, zoneData, zoneOptions, machineData, machineOptions, shiftData, toolNumbersData, toolNumbersOptions, supervisorsData, auditorNameData, workCellData, workCellOptions, allMappingData: sortedData, activeAuditCategoriesData, auditCategoryStatusData, showSubmit, ItemId: itemId, isEditForm });
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
        let toolNumbersData = [...this.state.toolNumbersData];
        let toolNumbersOptions: any = [...this.state.toolNumbersOptions];
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
            formData.ToolNumber = "";
            formData.WorkCell = "";

            if (actionMeta.action != "clear") {
                zoneOptions = [];
                zoneOptions = zoneData.filter((option: any) => (option.Plant && option.Plant.Title == formData.Plant && option.Department && option.Department.Id == event.id)).map((item: any) => ({ label: item.Title, value: item.Title, id: item.Id }));
            }
            else { zoneOptions = []; }
            machineOptions = [];
            toolNumbersOptions = [];
            workCellOptions = [];
        }
        else if (name == "Zone") {
            formData.Machine = "";
            formData.ToolNumber = "";
            formData.WorkCell = "";

            if (actionMeta.action != "clear") {
                machineOptions = [];
                toolNumbersOptions = [];

                machineOptions = machineData.filter((option: any) => (option.Plant && option.Plant.Title == formData.Plant && option.Department && option.Department.Title == formData.Department && option.Zone.Id == event.id)).map((item: any) => ({ label: item.Title, value: item.Title, id: item.Id }));

                toolNumbersOptions = toolNumbersData.filter((option: any) => (option.Plant && option.Plant.Title == formData.Plant && option.Department && option.Department.Title == formData.Department && option.Zone.Title == event.value)).map((item: any) => ({ label: item.Title, value: item.Title, id: item.Id }));

                workCellOptions = workCellData.filter((option: any) => (option.h7kc == formData.Plant && option.Department == formData.Department && option.Zone == event.value)).map((item: any) => ({ label: item.Title, value: item.Title, id: item.Id }));
            }
            else {
                machineOptions = [];
                toolNumbersOptions = [];
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

        this.setState({ formData, departmentOptions, zoneOptions, machineOptions, toolNumbersOptions, workCellOptions });
    }

    private handleDateChange = (dateValue: any, name: any, divId: any, dateProps: any) => {
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
                date: { val: formData.WCCDate, required: true, Name: "Date", Type: ControlType.date, Focusid: "dtWCCDate" },
                dateToday: { val: formData.WCCDate, required: true, Name: "Date", Type: ControlType.lessthanTodayDate, Focusid: "dtWCCDate" },
                shift: { val: formData.ShiftType, required: true, Name: "ShiftType", Type: ControlType.reactSelect, Focusid: "ddlShift" },
                auditorsName: { val: formData.AuditorName, required: true, Name: "Auditor's Name", Type: ControlType.reactSelect, Focusid: "ddlAuditorName" },
                plant: { val: formData.Plant, required: true, Name: "Plant", Type: ControlType.reactSelect, Focusid: "ddlPlant" },
                department: { val: formData.Department, required: true, Name: "Department", Type: ControlType.reactSelect, Focusid: "ddlDepartment" },
                zone: { val: formData.Zone, required: true, Name: "Zone", Type: ControlType.reactSelect, Focusid: "ddlZone" },
                machine: { val: formData.Machine, required: true, Name: "Machine", Type: ControlType.reactSelect, Focusid: "ddlMachine" },
                workCell: { val: formData.WorkCell, required: true, Name: "Work Cell", Type: ControlType.reactSelect, Focusid: "ddlWorkCell" },
                toolNumber: { val: formData.ToolNumber, required: (formData.Department == "Molding"), Name: "Tool Number", Type: ControlType.reactSelect, Focusid: "ddlToolNumber" },
                supervisor: { val: formData.Supervisor, required: true, Name: "Supervisor", Type: ControlType.reactSelect, Focusid: "ddlSupervisor" },
                completedDate: { startDate: formData.WCCDate, endDate: formData.CompletedDate, required: (formData.CompletedDate != ''), startDateName: "Date", endDateName: "Date Completed", Type: ControlType.compareDates, Focusid: "dtCompletedDate" }
            }

            let isValid = formValidation.FormValidation(data);

            if (isValid.status) {
                let isValidDynamicComments = this.validateDynamicSubCategoryComments();
                if (isValidDynamicComments) {

                    //Date formatting, Year and YearMonth
                    formData.WCCDate = DateUtilities.addBrowserwrtServer(new Date(DateUtilities.getDateMMDDYYYY(formData.WCCDate)), this.props.spContext.webTimeZoneData).toISOString();
                    let mmddyyyyDate = format(formData.WCCDate, "MM/dd/yyyy");
                    formData.Year = mmddyyyyDate.split("/")[2];
                    formData.YearMonth = mmddyyyyDate.split("/")[0];
                    if (formData.CompletedDate != "") {
                        formData.CompletedDate = DateUtilities.addBrowserwrtServer(new Date(DateUtilities.getDateMMDDYYYY(formData.CompletedDate)), this.props.spContext.webTimeZoneData).toISOString();
                    }
                    else {
                        delete formData.CompletedDate;
                    }

                    //Get Status Count
                    let allMappingData = { ...this.state.allMappingData }
                    let unsafeact = 0;
                    let unsafeConddition = 0;
                    for (let i = 0; i < Object.keys(allMappingData).length; i++) {
                        var subCategory = allMappingData[i];
                        var status = subCategory.SubCategoryStatus;

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
                    formData.ActionHistory.push({ ActionBy: this.props.userDisplayName, ActionDateTime: DateUtilities.addBrowserwrtServer(new Date(), this.props.spContext.webTimeZoneData) })
                    formData.ActionHistory = JSON.stringify(formData.ActionHistory);
                    await this.InsertOrUpdateData(formData);
                }
            }
            else {
                showToast("error", isValid.message);
                hideLoader();
            }
        } catch (e) {
            console.log(e);
            this.onError();
        }
    }

    private InsertOrUpdateData = async (formData: any) => {
        try {
            let itemId = this.props.match.params.id;
            if (itemId > 0) {
                await this.sp.web.lists.getByTitle(this.SMATList).items.getById(itemId).update(formData).then((res: any) => {
                    this.UpdateLineItems();
                }, (error) => {
                    console.log(error);
                    this.onError();
                })
            }
            else {
                await this.sp.web.lists.getByTitle(this.SMATList).items.add(formData).then((res: any) => {
                    let childObjects = this.updateWCCId(res.Id.toString());
                    this.InsertLineItems(childObjects, res.Id);
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
                this.sp.web.lists.getByTitle(this.SMATChildList).items.using(batchedPipe).add(item);
            }
            await execute().then(async () => {
                let GroupName = await this.getGroupName();
                if (GroupName != '') {
                    let GroupMemberEmails = await getGroupMemberEmails(GroupName, this.props.siteURL);
                    if (GroupMemberEmails.length) {
                        let link = this.props.webAbsoluteURL + '/SitePages/Home.aspx#/SMATForm/' + adedItemId
                        let body = "<p>Hi,</p>" + "<p>New 'SMAT-" + adedItemId + "' has been submitted. Please <a href='" + link + "'><b>click here</b></a> to view the details.</p><p>Regards<br>" + this.props.userDisplayName + "</p>";
                        await sendEmail(this.props.siteURL, GroupMemberEmails, "New 'SMAT' Submitted", body);
                    }
                }
                let msg = "SMAT Submitted Successfully";
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
        finally {
            hideLoader();
        }
    }

    private UpdateLineItems = async () => {
        try {
            let childPostObjects = this.createUpdateObjects();
            const [batchedPipe, execute] = createBatch(this.sp.web);
            for (const item of childPostObjects) {
                this.sp.web.lists.getByTitle(this.SMATChildList).items.getById(item.id).using(batchedPipe).update(item.Obj);
            }
            await execute().then(async () => {
                let msg = "SMAT Updated Successfully";
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
        finally {
            hideLoader();
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
                    WCCID: mappingItem.WCCID,
                    WccCategory: mappingItem.WccCategory,
                    WccSubCategory: mappingItem.WccSubCategory,
                    Attachment: mappingItem.Attachment,
                    SubCategoryStatus: mappingItem.SubCategoryStatus,
                    SubCategoryComments: mappingItem.SubCategoryStatus == "Satisfactory" ? '' : mappingItem.SubCategoryComments
                }
            }
            postObjects.push(postObj);
        }
        return postObjects;
    }

    private updateWCCId = (ParentId: string) => {
        let allMappingData = [...this.state.allMappingData];

        let updatedMappingData = allMappingData.map(item => ({
            ...item,
            WCCID: ParentId
        }));

        return updatedMappingData;
    }

    private validateDynamicSubCategoryComments = () => {
        let allMappingData: any = [...this.state.allMappingData];
        let isValid = true;
        for (var i = 0; i < allMappingData.length; i++) {
            if (allMappingData[i].SubCategoryStatus != "Satisfactory" && ([null, undefined, ''].includes(allMappingData[i].SubCategoryComments))) {
                isValid = false;
                let txtAreaId = "txtSubCategory" + i;
                let element = document.getElementById(txtAreaId);
                element?.classList.add("mandatory-FormContent-focus");
                element?.focus();
                showToast("error", "Comments cannot be blank for Status-" + allMappingData[i].SubCategoryStatus);
                hideLoader();
                break;
            }
        }
        return isValid;
    }

    private onSuccess = () => {
        hideLoader();
        this.setState({ Redirect: true, RedirectTo: 'SMATView', ItemID: 0 });
        showToast("success", this.state.displayMessage);
    }

    private onError = () => {
        showToast("error", ActionStatus.Error);
        hideLoader();
    }
    private getGroupName=async() =>{
        //var selectedDept = this.state.formData.Department;
        // let group = "OTHER [WCM Merrill Safety Mgt]"; // Default group
        let group = "WCM Merrill Safety Mgt"; // Default group
        let { getListItems } = initCommonFunctions(this.props.context, this.props.siteURL);
        let EmailConfigList = 'EmailsConfiguration', PlantSelQuery = 'Plant/Title,Department/Title,ToEmailGroup/Title,*', plantFiltQuery = `Plant/Title eq '${this.state.formData.Plant}' and Department/Title eq '${this.state.formData.Department}' and Form eq 'SMAT'`, PlantExpFields = 'Plant,Department,ToEmailGroup';
        try {
               let items=await getListItems(EmailConfigList, this.MaycoURL, PlantSelQuery, PlantExpFields, plantFiltQuery);
               if(items.length)
               {
                group=items[0].ToEmailGroup.Title;
               }
        }
        catch(e) {
            console.log(e);
            this.onError();
        }

        // switch (selectedDept) {
        //     case "Molding":
        //     case "Deco":
        //         group = "WCM Merrill Safety Deco";
        //         break;
        //     case "Foam":
        //     case "GM Assembly":
        //     case "IP Manufacturing":
        //     case "JL Assembly":
        //     case "Sequencing":
        //     case "Sports Bar":
        //     case "Thermoform":
        //     case "WD TP Assembly":
        //         group = "WCM Merrill Safety Seq";
        //         break;
        //     case "Maintenance":
        //         group = "WCM Safety Maintenance";
        //         break;
        //     case "Shipping  Receiving":
        //         group = "WCM Safety Shipping";
        //         break;
        //     case "Quality":
        //         group = "WCM SMAT Quality";
        //         break;
        // }
        return group;
    }

    private handleCancel = () => {
        this.setState({ Redirect: true, RedirectTo: 'SMATView', ItemID: 0 });
    }

    private bindDynamicTable = () => {
        const { allMappingData, auditCategoryStatusData, isInputDisabled } = this.state;

        // Ensure unique categories by creating a Set of WccCategory
        const uniqueCategories = Array.from(new Set(allMappingData.map(item => item.WccCategory)));

        // Initialize globalCategoryIndex for numbering
        let globalCategoryIndex = 1;

        const tBody = uniqueCategories.map((category, categoryIndex) => {
            // Filter subcategories that match the current category
            const categoryRows = allMappingData.filter(subCategory => subCategory.WccCategory === category).sort((a: any, b: any) => a.WccSubCategory.localeCompare(b.WccSubCategory));

            return (
                <React.Fragment key={categoryIndex}>
                    {/* Category Row */}
                    <tr className="lightgreybg fs-6">
                        <td colSpan={3}><b>{category}</b></td>
                    </tr>

                    {/* Subcategory Rows */}
                    {categoryRows.map((subCategory, subIndex) => {

                        // const rowIndex = allMappingData.findIndex( item =>  item.WccCategory == category && item.WccSubCategory == subCategory.WccSubCategory);
                        const rowIndex = allMappingData.indexOf(subCategory);
                        const isSatisfactory = subCategory.SubCategoryStatus === 'Satisfactory';

                        return (
                            <tr key={rowIndex}>
                                <td>
                                    <div className="pull-left">{globalCategoryIndex++}.{subCategory.WccSubCategory}</div>
                                    <div className="pull-right">
                                        <MultipleImageUploader
                                            onImageUpload={this.onImageChange}
                                            onRemoveImage={this.onRemoveImage}
                                            initialImageSrc={subCategory.Attachment}
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
                                            value={subCategory.SubCategoryStatus} // controlled input
                                        >
                                            {auditCategoryStatusData.map((status: any) => (
                                                <option key={status.Title} value={status.Title}>
                                                    {status.Title}
                                                </option>
                                            ))}
                                        </select>
                                    </div>
                                </td>
                                {allMappingData.some(i => i.SubCategoryStatus != 'Satisfactory') && <td>
                                    <div className={` ${isSatisfactory ? 'd-none' : ''}`} id={`divSubCategory${rowIndex}`}>
                                        <textarea
                                            className="form-control bs-textarea"
                                            rows={1}
                                            id={`txtSubCategory${rowIndex}`}
                                            value={subCategory.SubCategoryComments}
                                            name="Comments"
                                            placeholder="Comments"
                                            disabled={isInputDisabled}
                                            onChange={(event) => this.handleSubCategoryCommentsChange(event, rowIndex)}
                                            title="Comments"
                                        />
                                    </div>
                                </td>}
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
        allMappingData[CategoryIndex].SubCategoryStatus = value;
        this.setState({ allMappingData });
    };

    private handleSubCategoryCommentsChange = (event: any, CategoryIndex: number) => {
        var element = document.getElementsByClassName("mandatory-FormContent-focus");
        if (element.length > 0) {
            for (let i = 0; i < element.length; i++) {
                element[i].classList.remove("mandatory-FormContent-focus");
            }
        }
        const allMappingData: any = [...this.state.allMappingData];
        allMappingData[CategoryIndex].SubCategoryComments = event.target.value;
        this.setState({ allMappingData });
    };

    private onImageChange = (base64: string, lineIndex: number) => {
        const allMappingData: any = [...this.state.allMappingData];
        allMappingData[lineIndex].Attachment = base64;

        this.setState({ allMappingData });
    };

    private onRemoveImage = (lineIndex: number) => {
        const allMappingData: any = [...this.state.allMappingData];
        allMappingData[lineIndex].Attachment = '';

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
                            <div className="div-form-title">
                                <div className="form-title">{" SMAT Form " + (this.state.isEditForm ? (" - " + this.state.ItemId) : "")} </div>
                                <span className="span-mandatory-text"> <span className="text-danger">* </span> are mandatory fields</span>
                            </div>
                            <div className="">
                                <div className="greenborder">
                                    <div className="form-border-box p-2 mx-3 my-2">
                                        <div className="row">
                                            {/* Date */}
                                            <div className="col-md-3">
                                                <div className="light-text">
                                                    <label className="" htmlFor="dtWCCDate"> Date <span className="mandatoryhastrick">*</span></label>
                                                    <div className="custom-datepicker" id="divWCCDate">
                                                        <DatePickercontrol placeholder="MM/DD/YYYY" selectedDate={this.state.formData.WCCDate} title={this.state.formData.WCCDate} id='dtWCCDate' isDisabled={this.state.isInputDisabled} startDate={undefined} endDate={new Date()} name="WCCDate" onDatechange={(dateProps: any) => this.handleDateChange(dateProps[0], dateProps[2], "divWCCDate", dateProps)} highlightDate={new Date()} showIcon />
                                                    </div>
                                                </div>
                                            </div>
                                            {/* Shift */}
                                            <div className="col-md-3">
                                                <div className="custom-dropdown" id="divShift" title={this.state.formData.ShiftType}>
                                                    <SearchableDropdown
                                                        label={"Shift"}
                                                        Title={"Shift"}
                                                        name={"ShiftType"}
                                                        id="ddlShift"
                                                        placeholderText={""}
                                                        className={""}
                                                        selectedValue={this.state.formData.ShiftType}
                                                        OptionsList={this.state.shiftData}
                                                        OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta, "divShift") }}
                                                        isRequired={true}
                                                        disabled={this.state.isInputDisabled}
                                                        noOptionsMessage="No Shift Type available"
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
                                                        noOptionsMessage="No Auditor Names available"
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
                                                <div className="custom-dropdown" id="divDepartment" title={this.state.formData.Department}>
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
                                                        noOptionsMessage="No Departments available"
                                                    />
                                                </div>
                                            </div>
                                            {/* Zone */}
                                            <div className="col-md-3">
                                                <div className="custom-dropdown" id="divZone" title={this.state.formData.Zone}>
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
                                                        noOptionsMessage="No Zones available"
                                                    />
                                                </div>
                                            </div>
                                            {/* Machine */}
                                            <div className="col-md-3">
                                                <div className="custom-dropdown" id="divMachine" title={this.state.formData.Machine}>
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
                                                        noOptionsMessage="No Machines available"
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
                                                        noOptionsMessage="No Work Cell available"
                                                    />
                                                </div>
                                            </div>
                                        </div>
                                        <div className="row">
                                            {/* Tool No. */}
                                            <div className="col-md-3">
                                                <div className="custom-dropdown" id="divToolNumber" title={this.state.formData.ToolNumber}>
                                                    <SearchableDropdown
                                                        label={"Tool No."}
                                                        Title={"ToolNumber"}
                                                        name={"ToolNumber"}
                                                        id="ddlToolNumber"
                                                        placeholderText={""}
                                                        className={""}
                                                        selectedValue={this.state.formData.ToolNumber}
                                                        OptionsList={this.state.toolNumbersOptions}
                                                        OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta, "divToolNumber") }}
                                                        isRequired={this.state.formData.Department == "Molding"}
                                                        disabled={this.state.isInputDisabled}
                                                        noOptionsMessage="No Tool Numbers available"
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
                                                        noOptionsMessage="No Supervisors available"
                                                    />
                                                </div>
                                            </div>
                                        </div>
                                        <div className="row mt-2">
                                            <div className="col-md-12">
                                                <div className="form-border-box p-1 my-2">
                                                    <table className="col-md-12">
                                                        <thead >
                                                            <tr className="darkgraybg fs-5">
                                                                <th>Requirement</th>
                                                                <th>Status</th>
                                                                {this.state.allMappingData.some(i => i.SubCategoryStatus != 'Satisfactory') && <th>Comments<span className="mandatoryhastrick"> *</span></th>}
                                                            </tr>
                                                        </thead>
                                                        <tbody>
                                                            {this.bindDynamicTable()}
                                                        </tbody>
                                                    </table>
                                                </div>
                                            </div>
                                        </div>
                                        {/* Comments */}
                                        <div className="row pb-2">
                                            <div className="col-md-12">
                                                <div className="light-text" >
                                                    <label className=" col-form-label" htmlFor="txtComments">Comments </label>
                                                    <textarea className="form-control bs-textarea" rows={3} id="txtComments" name="Comments" ref={this.txtComments} placeholder="" value={this.state.formData.Comments} title={this.state.formData.Comments} onChange={this.handleChange} ></textarea>
                                                </div>
                                            </div>
                                            {/* Supervisor Action Completed */}
                                            <div className="col-md-9 mt-2">
                                                <div className="">
                                                    <div className="light-text">
                                                        <label className=" col-form-label" htmlFor="txtActionCompleted">Action Completed </label>
                                                        <textarea className="form-control bs-textarea" rows={3} id="txtActionCompleted" name="ActionCompleted" ref={this.txtActionCompleted} placeholder="" value={this.state.formData.ActionCompleted} title={this.state.formData.ActionCompleted} onChange={this.handleChange}></textarea>
                                                    </div>
                                                </div>
                                            </div>
                                            {/* Completed Date */}
                                            <div className="col-md-3 mt-2">
                                                <div className="light-text">
                                                    <label className="" htmlFor="dtCompletedDate"> Completed Date </label>
                                                    <div className="custom-datepicker" id="divCompletedDate">
                                                        <DatePickercontrol placeholder="MM/DD/YYYY" selectedDate={this.state.formData.CompletedDate} title={this.state.formData.CompletedDate} id='dtCompletedDate' startDate={undefined} endDate={undefined} name="CompletedDate" onDatechange={(dateProps: any) => this.handleDateChange(dateProps[0], dateProps[2], "divCompletedDate", dateProps)} highlightDate={new Date()} showIcon />
                                                    </div>
                                                </div>
                                            </div>
                                        </div>

                                        <div className="col-sm-12 text-center py-3" id="">
                                            {this.state.showSubmit && <button type="button" id="btnSubmit" className="btn btn-primary mx-2" onClick={this.handleSubmit} title={this.state.ItemId > 0 ? 'Update' : 'Submit'}>{this.state.ItemId > 0 ? 'Update' : 'Submit'}</button>}
                                            <button type="button" id="btnCancel" className="btn btn-secondary" onClick={this.handleCancel} title="Cancel">Cancel</button>
                                        </div>
                                        {this.state.formData.ActionHistory.length > 0 &&
                                            <div className="col-md-12">
                                                <div className="form-border-box p-2 mx-1">
                                                    <h6 className=""><FontAwesomeIcon icon={faHistory} /> Action History</h6>
                                                    <ActionHistory HeaderData={["Action By", "Date & Time"]} HistoryData={this.state.formData.ActionHistory} spContext={this.props.spContext} />
                                                </div>
                                            </div>
                                        }

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