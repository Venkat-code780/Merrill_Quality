import * as React from "react";
import { hideLoader, showLoader } from "../Shared/Loader";
import { Navigate } from "react-router-dom";
import { SPHttpClient } from "@microsoft/sp-http";
import { ActionStatus, ControlType } from "../Constants/Contants";
import { showToast } from "../Shared/Toaster";
import { highlightCurrentNav } from "../Utilities/HighlightCurrentComponent";
import SearchableDropdown from "../Shared/Dropdown";
import DatePickercontrol from "../Shared/DatePickerField";
import { format } from "date-fns";
import DateUtilities from "../Utilities/DateUtilities";
import { initCommonFunctions } from "../Utilities/CommonFunctions";
import formValidation from "../Utilities/FormValidator";
import "../CSS/form-input-style.css";

// import FileUpload from "../Shared/FileUpload";
// import InputCheckBox from "../Shared/InputCheckBox";
// import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";
// import { faChevronDown, faChevronUp } from "@fortawesome/free-solid-svg-icons";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";
import "@pnp/sp/folders";
// import SingleImageUpload from "../Shared/singleFileUpload";
import ImageUploader from "../Shared/ImageUploader";

export interface UCANFormProps {
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

export interface UCANFormState {
}

export default class UCANForm extends React.Component<UCANFormProps, UCANFormState> {

    private sp = spfi().using(SPFx(this.props.context));
    private currPlantObj: any;
    private txtReportedBy: any;
    private txtLocationPersons: any;
    private txtOriginialTagNo: any;
    private txtDate: any;
    private rdSafetyTag: any;
    private txtDateCompleted: any;
    private txtDescriptionOfIncident: any;
    private txtActionPlan: any;
    private txtActionCompleted: any;

    private MaycoURL: string;
    private uaTypesList = "UATypes";
    private subTypesList = "UAMicroTypes";
    private ucanList = "UCAN";
    public state = {
        formData: {
            UCAN_x0020_Type: '',
            Plant: '',
            Department: '',
            Zone: '',
            Machine: '',
            UATypeId: '',
            Sub_x002d_TypeId: '',
            Reported_x0020_By: '',
            Original_x0020_Tag_x0020_No_x002: '',
            Safety_x0020_Tag: false,
            Location_x002f_Persons: '',
            Shift: '',
            Date: '',
            Date_x0020_Completed: '',
            Description_x0020_of_x0020_Incid: '',
            Action_x0020_Plan: '',
            Action_x0020_Completed: '',
            Year: '',
            YearMonth: '',
            Attachment: '',
        },
        ucanTypeData: [],
        plantsData: [],
        departmentData: [],
        departmentOptions: [],
        zoneData: [],
        zoneOptions: [],
        machineData: [],
        machineOptions: [],
        shiftData: [],
        uaTypeData: [],
        subTypeData: [],
        subTypeOptions: [],
        isInputDisabled: false,
        isEditForm: false,
        ItemId: 0,
        Redirect: false,
        RedirectTo: '',
        displayMessage: '',
        image: [],
        imageBase64: '',
        showSubmit: '',
        isAdmin: false
    }

    constructor(props: UCANFormProps) {
        super(props);
        this.MaycoURL = `${this.props.siteURL}/mayco`;

        this.txtReportedBy = React.createRef();
        this.txtLocationPersons = React.createRef();
        this.txtOriginialTagNo = React.createRef();
        this.txtDate = React.createRef();
        this.rdSafetyTag = React.createRef();
        this.txtDateCompleted = React.createRef();
        this.txtDescriptionOfIncident = React.createRef();
        // this.txtActionPlan = React.createRef();
        this.txtActionCompleted = React.createRef();
    }

    public componentDidMount(): void {
        highlightCurrentNav("liUCANForm");
        document.title = "Mayco - Safety | UCAN";
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
            let uaTypesSelQuery = 'Title,*', uaTypesFiltQuery = '', uaTypesExpFields = '';
            let subTypesSelQuery = 'Title,UAType0/Id,UAType0/Title,*', subTypesFiltQuery = '', subTypesExpFields = 'UAType0';
            let [Plants, departmentData, zoneData, machineData, uaTypes, subTypeData, shifts] = await Promise.all([
                getListItems(PlantList, this.MaycoURL, PlantSelQuery, PlantExpFields, plantFiltQuery),
                getListItems(DepartmentList, this.MaycoURL, DepartmentSelQuery, DepartmentExpFields, DepartmentFiltQuery),
                getListItems(ZoneList, this.MaycoURL, ZoneSelQuery, ZoneExpFields, ZoneFiltQuery),
                getListItems(MachineList, this.MaycoURL, MachineSelQuery, MachineExpFields, MachineFiltQuery),
                getListItems(this.uaTypesList, this.props.webAbsoluteURL, uaTypesSelQuery, uaTypesExpFields, uaTypesFiltQuery),
                getListItems(this.subTypesList, this.props.webAbsoluteURL, subTypesSelQuery, subTypesExpFields, subTypesFiltQuery),
                getListItems(ShiftsList, this.MaycoURL, ShiftsSelQuery, ShiftsExpFields, ShiftsFiltQuery),
            ]);

            Plants.sort((a, b) => a.Title.localeCompare(b.Title));
            departmentData.sort((a, b) => a.Title.localeCompare(b.Title));
            zoneData.sort((a, b) => a.Title.localeCompare(b.Title));
            machineData.sort((a, b) => a.Title.localeCompare(b.Title));
            uaTypes.sort((a, b) => a.Title.localeCompare(b.Title));
            subTypeData.sort((a, b) => a.Title.localeCompare(b.Title));
            let plantsData = Plants.map((item: any) => ({ label: item.Title, value: item.Title }));
            this.currPlantObj = plantsData.find((plant: any) => plant.label.toLowerCase() == this.props.currPlantTitle);
            formData.Plant = this.currPlantObj.label;
            formData.Reported_x0020_By = this.props.userDisplayName;

            let departmentOptions = departmentData.filter((option: any) => option.Plant.Title == formData.Plant).map((item: any) => ({ label: item.Title, value: item.Title, id: item.Id }));
            let zoneOptions: any = [];
            let machineOptions: any = [];
            let uaTypeData = uaTypes.map((item: any) => ({ label: item.Title, value: item.Id }));
            let subTypeOptions: any = [];
            let shiftData = shifts.map((item: any) => ({ label: item.Title, value: item.Title }));

            if (itemId > 0) {
                let UCANRes: any = await getListItems(this.ucanList, this.props.webAbsoluteURL, 'Author/Title,Author/Id,*', 'Author', `Id eq ${itemId}`);
                if (!UCANRes.isHttpRequestError) {
                    if (UCANRes.length) {
                        let editUCANItem = UCANRes[0];
                        formData.UCAN_x0020_Type = [null, undefined, ""].includes(editUCANItem.UCAN_x0020_Type) ? "" : editUCANItem.UCAN_x0020_Type;
                        formData.Plant = [null, undefined, ""].includes(editUCANItem.Plant) ? "" : editUCANItem.Plant;
                        formData.Department = [null, undefined, ""].includes(editUCANItem.Department) ? "" : editUCANItem.Department;
                        formData.Zone = [null, undefined, ""].includes(editUCANItem.Zone) ? "" : editUCANItem.Zone;
                        formData.Machine = [null, undefined, ""].includes(editUCANItem.Machine) ? "" : editUCANItem.Machine;
                        formData.UATypeId = [null, undefined, ""].includes(editUCANItem.UATypeId) ? "" : editUCANItem.UATypeId;
                        formData.Sub_x002d_TypeId = [null, undefined, ""].includes(editUCANItem.Sub_x002d_TypeId) ? "" : editUCANItem.Sub_x002d_TypeId;
                        formData.Reported_x0020_By = [null, undefined, ""].includes(editUCANItem.Reported_x0020_By) ? "" : editUCANItem.Reported_x0020_By;
                        formData.Original_x0020_Tag_x0020_No_x002 = [null, undefined, ""].includes(editUCANItem.Original_x0020_Tag_x0020_No_x002) ? "" : editUCANItem.Original_x0020_Tag_x0020_No_x002;
                        formData.Safety_x0020_Tag = [null, undefined, ""].includes(editUCANItem.Safety_x0020_Tag) ? "" : editUCANItem.Safety_x0020_Tag;
                        formData.Location_x002f_Persons = [null, undefined, ""].includes(editUCANItem.Location_x002f_Persons) ? "" : editUCANItem.Location_x002f_Persons;
                        formData.Shift = [null, undefined, ""].includes(editUCANItem.Shift) ? "" : editUCANItem.Shift;
                        formData.Date = [null, undefined, ""].includes(editUCANItem.Date) ? "" : editUCANItem.Date;
                        formData.Date_x0020_Completed = [null, undefined, ""].includes(editUCANItem.Date_x0020_Completed) ? "" : editUCANItem.Date_x0020_Completed;
                        formData.Description_x0020_of_x0020_Incid = [null, undefined, ""].includes(editUCANItem.Description_x0020_of_x0020_Incid) ? "" : editUCANItem.Description_x0020_of_x0020_Incid;
                        formData.Action_x0020_Plan = [null, undefined, ""].includes(editUCANItem.Action_x0020_Plan) ? "" : editUCANItem.Action_x0020_Plan;
                        formData.Action_x0020_Completed = [null, undefined, ""].includes(editUCANItem.Action_x0020_Completed) ? "" : editUCANItem.Action_x0020_Completed;
                        formData.Year = [null, undefined, ""].includes(editUCANItem.Year) ? "" : editUCANItem.Year;
                        formData.YearMonth = [null, undefined, ""].includes(editUCANItem.YearMonth) ? "" : editUCANItem.YearMonth;
                        formData.Attachment = [null, undefined, ""].includes(editUCANItem.Attachment) ? "" : editUCANItem.Attachment;

                        zoneOptions = zoneData.filter((option: any) => (option.Plant && option.Plant.Title == formData.Plant && option.Department && option.Department.Title == editUCANItem.Department)).map((item: any) => ({ label: item.Title, value: item.Title, id: item.Id }));

                        machineOptions = machineData.filter((option: any) => (option.Plant && option.Plant.Title == formData.Plant && option.Department && option.Department.Title == formData.Department && option.Zone.Title == editUCANItem.Zone)).map((item: any) => ({ label: item.Title, value: item.Title, id: item.Id }));

                        subTypeOptions = subTypeData.filter((option: any) => option.UAType0.Id == editUCANItem.UATypeId).map((item: any) => ({ label: item.Title, value: item.Id, }));

                        //Super Admin
                        showSubmit = (editUCANItem.Author == this.props.userDisplayName || this.props.isSuperAdmin) ? true : false;
                    }
                    else {
                        showToast("error", "No UCAN found");
                        this.setState({ Redirect: true, RedirectTo: 'Home' });
                    }
                }
            }

            this.setState({ formData, plantsData, departmentData, departmentOptions, zoneData, zoneOptions, machineData, machineOptions, uaTypeData, subTypeData, subTypeOptions, shiftData, showSubmit, ItemId: itemId });
        } catch (e) {
            console.log(e);
            this.onError();
        }
        finally {
            hideLoader();
        }
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
        let subTypeData = [...this.state.subTypeData];
        let subTypeOptions: any = [...this.state.subTypeOptions];
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

            if (actionMeta.action != "clear") {
                zoneOptions = [];
                zoneOptions = zoneData.filter((option: any) => (option.Plant && option.Plant.Title == formData.Plant && option.Department && option.Department.Id == event.id)).map((item: any) => ({ label: item.Title, value: item.Title, id: item.Id }));
            }
            else { zoneOptions = []; }
            machineOptions = [];
        }
        else if (name == "Zone") {
            formData.Machine = "";

            if (actionMeta.action != "clear") {
                machineOptions = [];
                machineOptions = machineData.filter((option: any) => (option.Plant && option.Plant.Title == formData.Plant && option.Department && option.Department.Title == formData.Department && option.Zone && option.Zone.Id == event.id)).map((item: any) => ({ label: item.Title, value: item.Title, id: item.Id }));
            }
            else { machineOptions = []; }
        }
        else if (name == "UATypeId") {
            formData.Sub_x002d_TypeId = "";

            if (actionMeta.action != "clear") {
                subTypeOptions = [];
                subTypeOptions = subTypeData.filter((option: any) => option.UAType0.Id == value).map((item: any) => ({ label: item.Title, value: item.Id, }));
            }
            else { subTypeOptions = []; }
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

        this.setState({ formData, departmentOptions, zoneOptions, machineOptions, subTypeOptions });
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

    private onSuccess = () => {
        hideLoader();
        this.setState({ Redirect: true, RedirectTo: 'UCANView', ItemID: 0 });
        showToast("success", this.state.displayMessage);
    }

    private onError = () => {
        showToast("error", ActionStatus.Error);
        hideLoader();
    }

    public handleSubmit = () => {
        try {
            showLoader();
            var formData: any = { ...this.state.formData };
            var data = {
                ucanType: { val: formData.UCAN_x0020_Type, required: true, Name: "Near miss, Unsafe act, Unsafe condition", Type: ControlType.reactSelect, Focusid: "ddlUCANType" },
                plant: { val: formData.Plant, required: true, Name: "Plant", Type: ControlType.reactSelect, Focusid: "ddlPlant" },
                department: { val: formData.Department, required: true, Name: "Department", Type: ControlType.reactSelect, Focusid: "ddlDepartment" },
                zone: { val: formData.Zone, required: true, Name: "Zone", Type: ControlType.reactSelect, Focusid: "ddlZone" },
                machine: { val: formData.Machine, required: (formData.UCAN_x0020_Type != "Unsafe Act"), Name: "Machine", Type: ControlType.reactSelect, Focusid: "ddlMachine" },
                uaType: { val: formData.UATypeId, required: true, Name: "UA Type", Type: ControlType.reactSelect, Focusid: "ddlUAType" },
                subType: { val: formData.Sub_x002d_TypeId, required: true, Name: "Sub Type", Type: ControlType.reactSelect, Focusid: "ddlSubType" },
                shift: { val: formData.Shift, required: true, Name: "Shift", Type: ControlType.reactSelect, Focusid: "ddlShift" },
                date: { val: formData.Date, required: true, Name: "Date", Type: ControlType.date, Focusid: "dtDate" },
                dateToday: { val: formData.Date, required: true, Name: "Date", Type: ControlType.lessthanTodayDate, Focusid: "dtDate" },
                dateCompletedToday: { val: formData.Date_x0020_Completed, required: (formData.Date_x0020_Completed != ""), Name: "Date", Type: ControlType.lessthanTodayDate, Focusid: "dtDateCompleted" },
                dateComparision: { startDate: formData.Date, endDate: formData.Date_x0020_Completed, required: (formData.Date_x0020_Completed != ""), startDateName: "Date", endDateName: "Date Completed", Type: ControlType.compareDates, Focusid: "dtDateCompleted" },
                actionCompleted: { val: formData.Action_x0020_Completed, required: (formData.Date_x0020_Completed != ''), Name: "Action Completed", Type: ControlType.string, Focusid: this.txtActionCompleted }
            }

            let isValid = formValidation.FormValidation(data);

            if (isValid.status) {
                console.log("Valid Data");
                formData.Date = DateUtilities.addBrowserwrtServer(new Date(DateUtilities.getDateMMDDYYYY(formData.Date)), this.props.spContext.webTimeZoneData).toISOString();
                let mmddyyyyDate = format(formData.Date, "MM/dd/yyyy");
                console.log(mmddyyyyDate);
                formData.Year = mmddyyyyDate.split("/")[2];
                formData.YearMonth = mmddyyyyDate.split("/")[0];
                if (formData.Date_x0020_Completed != "") {
                    formData.Date_x0020_Completed = DateUtilities.addBrowserwrtServer(new Date(DateUtilities.getDateMMDDYYYY(formData.Date_x0020_Completed)), this.props.spContext.webTimeZoneData).toISOString();
                }
                else {
                    delete formData.Date_x0020_Completed;
                }

                this.InsertOrUpdateData(formData);
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

    private InsertOrUpdateData = async (formData: any) => {
        try {

            let itemId = this.props.match.params.id;
            if (itemId > 0) {
                await this.sp.web.lists.getByTitle(this.ucanList).items.getById(itemId).update(formData).then((res: any) => {
                    let msg = "UCAN updated successfully";
                    this.setState({ displayMessage: msg });
                    this.onSuccess();
                }, (error) => {
                    console.log(error);
                    this.onError();
                })
            }
            else {
                await this.sp.web.lists.getByTitle(this.ucanList).items.add(formData).then((res: any) => {
                    let msg = "UCAN submitted successfully";
                    this.setState({ displayMessage: msg });
                    this.onSuccess();
                }, (error) => {
                    console.log(error);
                    this.onError();
                })
            }
        }
        catch (e) {
            console.log(e);
            this.onError();
        }
    }

    private handlCancel = () => {
        this.setState({ Redirect: true, RedirectTo: 'UCANView', ItemID: 0 });
    }

    // private imageChanged = (selectedFiles: any) => {
    //     this.setState({ image: selectedFiles });
    // }

    private onImageChange = (base64: string) => {
        const formData = { ...this.state.formData };
        formData.Attachment = base64;
        console.log(base64);
        this.setState({ imageBase64: base64, formData })
    }

    private onRemoveImage = () => {
        const formData = { ...this.state.formData };
        formData.Attachment = '';
        this.setState({ imageBase64: '', formData })
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
                                <div className="form-title">{" Unsafe Conditions And Acts Form " + (this.state.isEditForm ? (" - " + this.state.ItemId) : "")} </div>
                                <span className="span-mandatory-text"> <span className="text-danger">* </span> are mandatory fields</span>
                            </div>
                            <div className="">
                                <div className="greenborder pt-2">
                                    <div className="row">
                                        <div className="col-lg-8 col-md-8 col-sm-8 col-xs-8 ps-0">
                                            <div className="row">
                                                {/* Near miss, Unsafe act, Unsafe condition */}
                                                <div className="col-md-6 ">
                                                    <div className="custom-dropdown" id="divUCANType">
                                                        <SearchableDropdown
                                                            label={"Near miss, Unsafe act, Unsafe condition"}
                                                            Title={"Near miss, Unsafe act, Unsafe condition"}
                                                            name={"UCAN_x0020_Type"}
                                                            id="ddlUCANType"
                                                            placeholderText={""}
                                                            className={""}
                                                            selectedValue={this.state.formData.UCAN_x0020_Type}
                                                            OptionsList={[{ label: "Near Miss", value: "Near Miss" }, { label: "Unsafe Act", value: "Unsafe Act" }, { label: "Unsafe Condition", value: "Unsafe Condition" }]}
                                                            OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta, "divUCANType") }}
                                                            isRequired={true}
                                                            disabled={this.state.isInputDisabled}
                                                            noOptionsMessage="No UCAN Types"
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
                                                            isRequired={this.state.formData.UCAN_x0020_Type != "Unsafe Act"}
                                                            disabled={this.state.isInputDisabled}
                                                            noOptionsMessage="No Machines"
                                                        />
                                                    </div>
                                                </div>
                                                {/*UA Type */}
                                                <div className="col-md-3">
                                                    <div className="custom-dropdown" id="divType">
                                                        <SearchableDropdown
                                                            label={"Type"}
                                                            Title={"Type"}
                                                            name={"UATypeId"}
                                                            id="ddlUAType"
                                                            placeholderText={""}
                                                            className={""}
                                                            selectedValue={this.state.formData.UATypeId}
                                                            OptionsList={this.state.uaTypeData}
                                                            OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta, "divType") }}
                                                            isRequired={true}
                                                            disabled={this.state.isInputDisabled}
                                                            noOptionsMessage="No Types"
                                                        />
                                                    </div>
                                                </div>
                                                {/* Sub-Type */}
                                                <div className="col-md-3">
                                                    <div className="custom-dropdown" id="divSubType">
                                                        <SearchableDropdown
                                                            label={"Sub-Type"}
                                                            Title={"Sub-Type"}
                                                            name={"Sub_x002d_TypeId"}
                                                            id="ddlSubType"
                                                            placeholderText={""}
                                                            className={""}
                                                            selectedValue={this.state.formData.Sub_x002d_TypeId}
                                                            OptionsList={this.state.subTypeOptions}
                                                            OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta, "divSubType") }}
                                                            isRequired={true}
                                                            disabled={this.state.isInputDisabled}
                                                            noOptionsMessage="No SubTypes"
                                                        />
                                                    </div>
                                                </div>
                                                <div className="col-lg-12 col-md-12 col-sm-12 col-xs-12 px-0">
                                                    <div className="row">
                                                        <div className="col-lg-3 col-md-3 col-sm-3 col-xs-3 pull-left px-0">
                                                            <div className="row">
                                                                {/* Reported By */}
                                                                <div className="col-lg-12 col-md-12 col-sm-12 col-xs-12 pb-1">
                                                                    <div className="light-text">
                                                                        <label className="" htmlFor="txtReportedBy">Reported By </label>
                                                                        <input className="form-control" placeholder="Reported By" name="Reported_x0020_By" type="text" id="txtReportedBy" ref={this.txtReportedBy} value={this.state.formData.Reported_x0020_By} onChange={this.handleChange} title="Reported By" />
                                                                    </div>
                                                                </div>
                                                                {/* Original Tag No. */}
                                                                <div className="col-lg-12 col-md-12 col-sm-12 col-xs-12 pb-1">
                                                                    <div className="light-text">
                                                                        <label className="" htmlFor="txtOriginialTagNo">Original Tag No. </label>
                                                                        <input className="form-control" placeholder="Original Tag No." name="Original_x0020_Tag_x0020_No_x002" type="text" id="txtOriginialTagNo" ref={this.txtOriginialTagNo} value={this.state.formData.Original_x0020_Tag_x0020_No_x002} onChange={this.handleChange} title="Original Tag No." />
                                                                    </div>
                                                                </div>
                                                                {/* Safety Tag */}
                                                                <div className="col-lg-12 col-md-12 col-sm-12 col-xs-12 ">
                                                                    <input className="" placeholder="Safety Tag" name="Safety_x0020_Tag" type="checkbox" id="rdSafetyTag" ref={this.rdSafetyTag} checked={this.state.formData.Safety_x0020_Tag} onChange={this.handleChange} title="Safety Tag" />
                                                                    <label className="pt-2 px-2" htmlFor="rdSafetyTag">Safety Tag </label>
                                                                </div>
                                                            </div>
                                                        </div>
                                                        <div className="col-lg-6 col-md-6 col-sm-6 col-xs-6 pull-left">
                                                            {/* Location/Persons */}
                                                            <div className="">
                                                                <div className="light-text mb-2">
                                                                    <label className="" htmlFor="txtLocationPersons">Location/Persons </label>
                                                                    <textarea className="form-control" rows={6} id="txtLocationPersons" name="Location_x002f_Persons" ref={this.txtLocationPersons} placeholder="Location/Persons" value={this.state.formData.Location_x002f_Persons} onChange={this.handleChange} title="Location/Persons" ></textarea>
                                                                </div>
                                                            </div>
                                                        </div>
                                                        <div className="col-lg-3 col-md-3 col-sm-3 col-xs-3  pull-left">
                                                            {/* Shift */}
                                                            <div className="pb-1">
                                                                <div className="custom-dropdown" id="divShift">
                                                                    <SearchableDropdown
                                                                        label={"Shift"}
                                                                        Title={"Shift"}
                                                                        name={"Shift"}
                                                                        id="ddlShift"
                                                                        placeholderText={""}
                                                                        className={""}
                                                                        selectedValue={this.state.formData.Shift}
                                                                        OptionsList={this.state.shiftData}
                                                                        OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta, "divShift") }}
                                                                        isRequired={true}
                                                                        disabled={this.state.isInputDisabled}
                                                                        noOptionsMessage="No Shifts"
                                                                    />
                                                                </div>
                                                            </div>
                                                            {/* Date */}
                                                            <div className="light-text pb-1">
                                                                <label className="z-in-9" htmlFor="dtDate"> Date <span className="mandatoryhastrick">*</span></label>
                                                                <div className="custom-datepicker" id="divDate">
                                                                    <DatePickercontrol placeholder="" selectedDate={this.state.formData.Date} id='dtDate' startDate={undefined} endDate={new Date()} name="Date" onDatechange={(dateProps: any) => this.handleDateChange(dateProps[0], dateProps[2], "divDate", dateProps)} ref={this.txtDate} highlightDate={new Date()} showIcon />
                                                                </div>
                                                            </div>
                                                            {/* Date Completed */}
                                                            <div className="light-text">
                                                                <label className="z-in-9" htmlFor="dtDateCompleted"> Date Completed </label>
                                                                <div className="custom-datepicker" id="divDateCompleted">
                                                                    <DatePickercontrol placeholder="" selectedDate={this.state.formData.Date_x0020_Completed} id='dtDateCompleted' startDate={undefined} endDate={new Date()} name="Date_x0020_Completed" onDatechange={(dateProps: any) => this.handleDateChange(dateProps[0], dateProps[2], "divDateCompleted", dateProps)} ref={this.txtDateCompleted} highlightDate={new Date()} showIcon />
                                                                </div>
                                                            </div>
                                                        </div>
                                                    </div>
                                                </div>
                                            </div>
                                        </div>
                                        <div className="col-lg-4 col-md-4 col-sm-4 col-xs-4">
                                            <div className="">
                                                <div className="light-text" >
                                                    <label className="" htmlFor="txtDescriptionOfIncident">Description of Incident </label>
                                                    <textarea className="form-control" rows={10} id="txtDescriptionOfIncident" name="Description_x0020_of_x0020_Incid" ref={this.txtDescriptionOfIncident} placeholder="Description of Incident" value={this.state.formData.Description_x0020_of_x0020_Incid} onChange={this.handleChange} title="Description of Incident" ></textarea>
                                                </div>
                                            </div>
                                        </div>
                                        <div className="col-lg-12 col-md-12 col-sm-12 col-xs-12">
                                            <div className="">
                                                <div className="light-text">
                                                    <label className="" htmlFor="txtActionPlan">Action Plan </label>
                                                    <textarea className="form-control" rows={3} id="txtActionPlan" name="Action_x0020_Plan" ref={this.txtActionPlan} placeholder="Action Plan" value={this.state.formData.Action_x0020_Plan} onChange={this.handleChange} title="Action Plan" ></textarea>
                                                </div>
                                            </div>
                                        </div>
                                        <div className="col-lg-12 col-md-12 col-sm-12 col-xs-12 mt-2">
                                            <div className="">
                                                <div className="light-text" >
                                                    <label className="" htmlFor="txtActionCompleted">Action Completed {this.state.formData.Date_x0020_Completed && <span className="mandatoryhastrick">*</span>}</label>
                                                    <textarea className="form-control" rows={3} id="txtActionCompleted" name="Action_x0020_Completed" ref={this.txtActionCompleted} placeholder="Action Completed" value={this.state.formData.Action_x0020_Completed} onChange={this.handleChange} title="Action Completed" ></textarea>
                                                </div>
                                            </div>
                                        </div>

                                        {/* <SingleImageUpload
                                    fileLabel="Upload Image"
                                    files={this.state.image}
                                    onChange={this.imageChanged}
                                    onRemove={this.imageChanged}
                                    isRequired={false}
                                    disabled={false}
                                /> */}
                                        <div className="col-md-12 mt-2">
                                            <ImageUploader
                                                onImageUpload={this.onImageChange}
                                                onRemoveImage={this.onRemoveImage}
                                                initialImageSrc={this.state.formData.Attachment}
                                            />
                                        </div>


                                        <div className="col-sm-12 text-center py-3 " id="">
                                            {this.state.showSubmit && <button type="button" id="btnSubmit" className="btn btn-primary mx-2" onClick={this.handleSubmit} title={this.state.ItemId > 0 ? 'Update' : 'Submit'}>{this.state.ItemId > 0 ? 'Update' : 'Submit'}</button>}
                                            <button type="button" id="btnCancel" className="btn btn-secondary" onClick={this.handlCancel} >Cancel</button>
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