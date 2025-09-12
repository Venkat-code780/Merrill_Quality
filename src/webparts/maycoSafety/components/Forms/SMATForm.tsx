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
import "../CSS/SMATForm.css";
import { ActionStatus } from "../Constants/Contants";
import { showToast } from "../Shared/Toaster";
import { highlightCurrentNav } from "../Utilities/HighlightCurrentComponent";
import DatePickercontrol from "../Shared/DatePickerField";
import DateUtilities from "../Utilities/DateUtilities";
import { format } from "date-fns";
import SearchableDropdown from "../Shared/Dropdown";
import { initCommonFunctions } from "../Utilities/CommonFunctions";

export interface SMATFormProps{
    match: any;
    spContext: any;
    spHttpClient: SPHttpClient;
    context: any;
    history: any;
    userDisplayName:string;
    siteURL:string;
    webAbsoluteURL:string;
    currPlantTitle:string;
}

export interface SMATFormState{
}

 export default class SMATForm extends React.Component<SMATFormProps,SMATFormState>{

    private sp = spfi().using(SPFx(this.props.context));
    private MaycoURL:string;
    private currPlantObj:any;
    private SMATList = "";

    private txtComments:any;
    private txtActionCompleted:any;
    public state = {
        formData:{
            Plant:'',
            Department:'',
            Zone:'',
            Machine:'',
            Shifts:'',
            ToolNumber:'',
            Comments:'',
            AuditorName:'',
            Supervisor:'',
            WorkCell:'',
            ShiftType:'',
            unsafeactCount:'',
            unsafeconditionCount:'',
            ActionCompleted:'',
            WCCDate:'',
            CompletedDate:'',
            Year:'',
            YearMonth:''
        },
        childFormData:{
            wccId:0,
            WccCategory:'',
            WccSubCategory:'',
            Attachment:'',
            SubCategoryStatus:'',
            SubCategoryComments:''
        },
        plantsData:[],
        departmentData:[],
        departmentOptions:[],
        zoneData:[],
        zoneOptions:[],
        machineData:[],
        machineOptions:[],
        shiftData:[],
        auditorNameData:[],
        workCellData:[],
        workCellOptions:[],
        toolNumbersData:[],
        toolNumbersOptions:[],
        supervisorsData:[],
        allCategoriesData:[],
        activeAuditCategoriesData:[],
        Homeredirect:false,
        isEditForm:false,
        ItemId:0,
        isInputDisabled: false,
        showSubmit: true
    }

    constructor(props: SMATFormProps){
        super(props);
        this.MaycoURL=`${this.props.siteURL}/mayco`;

        this.txtComments = React.createRef();
        this.txtActionCompleted = React.createRef();
    }
    
    public componentDidMount(): void {
        highlightCurrentNav("liSMATForm");
        document.title = "Mayco - Safety | SMAT";
        this.getOnLoadData();
    }

    private getOnLoadData = async () => {
        try {
            showLoader();
            var formData = {...this.state.formData};
            let itemId = this.props.match.params.id;
            let showSubmit = false;

            let  { getListItems } = initCommonFunctions(this.props.context,this.props.siteURL);
            let PlantList='Plant', PlantSelQuery='Title,*',plantFiltQuery='',PlantExpFields='';
            let DepartmentList='Department', DepartmentSelQuery='Title,Plant/Title,Plant/Id,*',DepartmentFiltQuery='',DepartmentExpFields='Plant';
            let ZoneList='Zones', ZoneSelQuery='Title,Plant/Title,Plant/Id,Department/Title,Department/Id,*',ZoneFiltQuery='',ZoneExpFields='Plant,Department';
            let MachineList='Machines', MachineSelQuery='Title,Plant/Title,Plant/Id,Department/Title,Department/Id,Zone/Title,Zone/Id,*',MachineFiltQuery='',MachineExpFields='Plant,Department,Zone';
            let ShiftsList='Shifts', ShiftsSelQuery='Title,*',ShiftsFiltQuery='',ShiftsExpFields='';
            let toolNumberList='Tool Numbers', toolNumberSelQuery='Title,Plant/Title,Plant/Id,Department/Title,Department/Id,Zone/Title,Zone/Id,*',toolNumberFiltQuery='',toolNumberExpFields='Plant,Department,Zone';
            let supervisorList='Supervisor', supervisorSelQuery='Title,Plant/Title,Plant/Id,*',supervisorFiltQuery='Is_x0020_Active eq 1',supervisorExpFields='Plant';
            let auditorsList='Auditors', auditorsSelQuery='Title,Plant/Title,Plant/Id,*',auditorsFiltQuery='Is_x0020_Active eq 1',auditorsExpFields='Plant';
            let workCellList='WorkCells', workCellSelQuery='Title,*',workCellFiltQuery='',workCellExpFields='';
            let mappingList='WCC/EHS mapping screen', mappingSelQuery='Audit_categories/Id,Audit_categories/Title,*',mappingFiltQuery="Form_x0020_Type eq 'WCC' and Is_x0020_Active eq 1",mappingExpFields='Audit_categories';
            let auditCategoriesList='Audit_Categories', auditCategoriesSelQuery='Title,*',auditCategoriesFiltQuery='Is_x0020_Active eq 1',auditCategoriesExpFields='';
            let [Plants,departmentData,zoneData,machineData, shifts, toolNumbersData, supervisorsData, auditorNameData, workCellData, mappingData, activeAuditCategoriesData ] = await Promise.all([
                getListItems(PlantList, this.MaycoURL, PlantSelQuery, PlantExpFields, plantFiltQuery ),
                getListItems(DepartmentList, this.MaycoURL, DepartmentSelQuery, DepartmentExpFields, DepartmentFiltQuery ),
                getListItems(ZoneList, this.MaycoURL, ZoneSelQuery, ZoneExpFields, ZoneFiltQuery ),
                getListItems(MachineList, this.MaycoURL, MachineSelQuery, MachineExpFields, MachineFiltQuery ),
                getListItems(ShiftsList, this.MaycoURL, ShiftsSelQuery, ShiftsExpFields, ShiftsFiltQuery ),
                getListItems(toolNumberList, this.MaycoURL, toolNumberSelQuery, toolNumberExpFields, toolNumberFiltQuery ),
                getListItems(supervisorList, this.MaycoURL, supervisorSelQuery, supervisorExpFields, supervisorFiltQuery ),
                getListItems(auditorsList, this.MaycoURL, auditorsSelQuery, auditorsExpFields, auditorsFiltQuery ),
                getListItems(workCellList, this.MaycoURL, workCellSelQuery, workCellExpFields, workCellFiltQuery ),
                getListItems(mappingList, this.props.webAbsoluteURL, mappingSelQuery, mappingExpFields, mappingFiltQuery ),
                getListItems(auditCategoriesList, this.props.webAbsoluteURL, auditCategoriesSelQuery, auditCategoriesExpFields, auditCategoriesFiltQuery )
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

            let plantsData = Plants.map((item: any) => ({ label: item.Title, value: item.Title }));
            this.currPlantObj = plantsData.find( (plant:any) => plant.label.toLowerCase() == this.props.currPlantTitle);
            formData.Plant = this.currPlantObj.label;
            
            let departmentOptions = departmentData.filter( (option:any) => option.Plant.Title == formData.Plant ).map((item: any) => ({ label: item.Title, value: item.Title, id:item.Id }));
            supervisorsData = supervisorsData.filter( (option:any) => option.Plant?.Title == formData.Plant ).map((item: any) => ({ label: item.Title, value: item.Title, id:item.Id }));
            auditorNameData = auditorNameData.filter( (option:any) => option.Plant.Title == formData.Plant ).map((item: any) => ({ label: item.Title, value: item.Title, id:item.Id }));
            let zoneOptions:any = [];
            let machineOptions:any = [];
            let toolNumbersOptions:any = [];
            let workCellOptions:any = [];
            let shiftData = shifts.map((item: any) => ({ label: item.Title, value: item.Title })); 

            let allCategoriesData = mappingData.map((item:any) => { 
                item.Audit_categories && item.Audit_SubCategory && this.checkAuditCaterory(item.Audit_categories.Title) ? { category: item.Audit_categories.Title, subCategory: item.Audit_SubCategory } : null
            }).filter( item => item!= null );

            if( itemId != undefined ){
                await this.sp.web.lists.getByTitle(this.SMATList).items.getById(itemId).expand("Author").select("*,Author/Title,Author/Id")().then( (editSMATItem:any) => {
                    if( editSMATItem != Error ){
                        formData.Plant = [null,undefined,""].includes(editSMATItem.Plant) ? "" : editSMATItem.Plant;
                        formData.Department = [null,undefined,""].includes(editSMATItem.Department) ? "" : editSMATItem.Department;
                        formData.Zone = [null,undefined,""].includes(editSMATItem.Zone) ? "" : editSMATItem.Zone;
                        formData.Machine = [null,undefined,""].includes(editSMATItem.Machine) ? "" : editSMATItem.Machine;
                        formData.Shifts = [null,undefined,""].includes(editSMATItem.Shifts) ? "" : editSMATItem.Shifts;
                        formData.ToolNumber = [null,undefined,""].includes(editSMATItem.ToolNumber) ? "" : editSMATItem.ToolNumber;
                        formData.Comments = [null,undefined,""].includes(editSMATItem.Comments) ? "" : editSMATItem.Comments;
                        formData.AuditorName = [null,undefined,""].includes(editSMATItem.AuditorName) ? "" : editSMATItem.AuditorName;
                        formData.Supervisor = [null,undefined,""].includes(editSMATItem.Supervisor) ? "" : editSMATItem.Supervisor;
                        formData.WorkCell = [null,undefined,""].includes(editSMATItem.WorkCell) ? "" : editSMATItem.WorkCell;
                        formData.ShiftType = [null,undefined,""].includes(editSMATItem.ShiftType) ? "" : editSMATItem.ShiftType;
                        formData.unsafeactCount = [null,undefined,""].includes(editSMATItem.unsafeactCount) ? "" : editSMATItem.unsafeactCount;
                        formData.unsafeconditionCount = [null,undefined,""].includes(editSMATItem.unsafeconditionCount) ? "" : editSMATItem.unsafeconditionCount;
                        formData.ActionCompleted = [null,undefined,""].includes(editSMATItem.ActionCompleted) ? "" : editSMATItem.ActionCompleted;
                        formData.WCCDate = [null,undefined,""].includes(editSMATItem.WCCDate) ? "" : editSMATItem.WCCDate;
                        formData.CompletedDate = [null,undefined,""].includes(editSMATItem.CompletedDate) ? "" : editSMATItem.CompletedDate;
                        formData.Year = [null,undefined,""].includes(editSMATItem.Year) ? "" : editSMATItem.Year;
                        formData.YearMonth = [null,undefined,""].includes(editSMATItem.YearMonth) ? "" : editSMATItem.YearMonth;

                        zoneOptions = zoneData.filter( (option:any) => ( option.Plant.Title == formData.Plant && option.Department.Title == editSMATItem.Department ) ).map((item: any) => ({ label: item.Title, value: item.Title, id:item.Id }));
                        machineOptions = machineData.filter( (option:any) => ( option.Plant.Title == formData.Plant && option.Department.Title == formData.Department && option.Zone.Title == editSMATItem.Zone ) ).map((item: any) => ({ label: item.Title, value: item.Title, id:item.Id }));
                        toolNumbersOptions = toolNumbersData.filter( (option:any) => ( option.Plant.Title == formData.Plant && option.Department.Title == formData.Department && option.Zone.Title == editSMATItem.Zone ) ).map((item: any) => ({ label: item.Title, value: item.Title, id:item.Id }));

                    }
                })
            }
            else{
                showSubmit = true;
            }

            this.setState({formData, plantsData, departmentData, departmentOptions, zoneData, zoneOptions, machineData, machineOptions, shiftData, toolNumbersData, toolNumbersOptions, supervisorsData, auditorNameData,workCellData, workCellOptions, showSubmit, allCategoriesData, activeAuditCategoriesData});

            hideLoader();
        } catch (e) {
            console.log(e);
            this.onError();
        }
    }

    private checkAuditCaterory(category:any){
        var isExist = false;
        const activeAuditCategoriesData = [...this.state.activeAuditCategoriesData];
        for (var ac in activeAuditCategoriesData) {
            const categoryRecord:any = activeAuditCategoriesData[ac];
            var title = categoryRecord.Title;
            if (title == category) {
                isExist = true;
                break;
            }
        }
        return isExist;
    }

    private handleChange= (event: any) => {
        const formData:any = {...this.state.formData};

        const name = event.target.name;
        let inputValue = (event.target.type == "text" || event.target.type == "textarea") ? event.target.value : (event.target.type == "checkbox" ? event.target.checked : event.target.value ? event.target.value : '' );
        let classArr: any = event.target.className;

        if (classArr.includes("onlyNum")) {
            inputValue = inputValue.replace(/[^0-9]/g, '');
        }
        formData[name] = inputValue;
        this.setState({formData });
    }

    private handleDropdownChange = ( event: any, actionMeta: any, id:any ) => {
        const formData:any = {...this.state.formData};
        // let departmentData = [...this.state.departmentData];
        let departmentOptions:any = [...this.state.departmentOptions];
        let zoneData = [...this.state.zoneData];
        let zoneOptions:any = [...this.state.zoneOptions];
        let machineData = [...this.state.machineData];
        let machineOptions:any = [...this.state.machineOptions];
        let toolNumbersData = [...this.state.toolNumbersData];
        let toolNumbersOptions:any = [...this.state.toolNumbersOptions];
        let workCellData = [...this.state.workCellData];
        let workCellOptions:any = [...this.state.workCellOptions];
        const name = actionMeta.name;
        const value = actionMeta.action == "clear" ? '' : event.value;
        formData[name] = value;

        if( name == "Department" ){
            formData.Zone = "";
            formData.Machine = "";

            if( actionMeta.action != "clear"){
                zoneOptions=[];
                zoneOptions = zoneData.filter( (option:any) => ( option.Plant.Title == formData.Plant && option.Department.Id == event.id ) ).map((item: any) => ({ label: item.Title, value: item.Title, id:item.Id }));
            }
            else{ zoneOptions = [];}
            machineOptions = [];
        }
        else if( name == "Zone" ){
            formData.Machine = "";
            formData.ToolNumber = "";
            formData.WorkCell = "";

            if( actionMeta.action != "clear"){
                machineOptions=[];
                toolNumbersOptions=[];

                machineOptions = machineData.filter( (option:any) => ( option.Plant.Title == formData.Plant && option.Department.Title == formData.Department && option.Zone.Id == event.id) ).map((item: any) => ({ label: item.Title, value: item.Title, id:item.Id }));
                toolNumbersOptions = toolNumbersData.filter( (option:any) => ( option.Plant.Title == formData.Plant && option.Department.Title == formData.Department && option.Zone.Id == event.id) ).map((item: any) => ({ label: item.Title, value: item.Title, id:item.Id }));
                workCellOptions = workCellData.filter( (option:any) => ( option.h7kc == formData.Plant && option.Department == formData.Department && option.Zone == event.value) ).map((item: any) => ({ label: item.Title, value: item.Title, id:item.Id }));
            }
            else{ machineOptions = [];}
        }

        if( !([null, undefined, ''].includes(id)) ){
            var ddlElement = document.getElementById(id);
            if( !([null, undefined, ''].includes(value)) ){
                ddlElement?.classList.add("active");
            }
            else{
                ddlElement?.classList.remove("active");
            }
        }

        this.setState({formData, departmentOptions, zoneOptions, machineOptions, toolNumbersOptions, workCellOptions });
    }

    private handleDateChange = (dateValue: any, name:any, divId:any, dateProps:any) => {
        const formData: any = { ...this.state.formData };
        console.log(dateProps);

        if( !([null, undefined, ''].includes(divId)) ){
            var ddlElement = document.getElementById(divId);
            if( !([null, undefined, ''].includes(dateValue)) ){
                ddlElement?.classList.add("active");
            }
            else{
                ddlElement?.classList.remove("active");
            }
        }

        if( dateValue != null ){
            dateValue = format( DateUtilities.addBrowserwrtServer( new Date(DateUtilities.getDateMMDDYYYY(dateValue)), this.props.spContext.webTimeZoneData).toISOString(), "MM/dd/yyyy");
        }
        else{
            dateValue = "";
        }

        formData[name] = dateValue;

        this.setState({ formData });
    }

    private handleSubmit = () => {

    }

    private onError = () => {
        showToast("error", ActionStatus.Error );
        hideLoader();
    }

    private handlefullClose = () => {
        this.setState({ Homeredirect: true, ItemID: 0 });
    }

    // private bindDynamicTable = () => {
    //     let tbody;

    //     return tbody;
    // }

    public render() {
        if (this.state.Homeredirect) {
            let url = "/Home";
            return (<Navigate to={url} />)
        }
        else{
            return(
                <React.Fragment>
                    <div className="container-fluid">
                        <div className="light-box border-box-shadow">
                            <div className="m-0 titlebg">
                                <h4 className="mb-0 pt-2 text-center">{" SMAT "+ (this.state.isEditForm ? (" - "+ this.state.ItemId): "")} </h4>
                                <label className="text-end px-1" style={{width:"100%"}}> <span className="text-danger">* </span> are mandatory fields</label>
                            </div>

                            <div className="mainContent row col-lg-12 col-md-12 col-sm-12 col-xs-12 borderLine">
                                {/* Date */}
                                <div className="col-md-3 greybg c-date-picker" id="divWCCDate">
                                    <label className="label-datePicker" htmlFor="dtWCCDate"> Date <span className="text-danger">*</span></label>
                                    <DatePickercontrol placeholder="" selectedDate={this.state.formData.WCCDate} id='dtWCCDate' isDisabled={this.state.isInputDisabled} startDate={undefined} endDate={undefined} name="WCCDate" onDatechange={(dateProps:any) => this.handleDateChange( dateProps[0], dateProps[2], "divWCCDate", dateProps)} highlightDate={new Date()} showIcon />
                                </div>
                                {/* Shift */}
                                <div className="col-md-3 greybg form-floating">
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
                                            OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta, "divShift" ) }}
                                            isRequired={true}
                                            disabled={this.state.isInputDisabled}
                                            noOptionsMessage="No Shifts"
                                        />
                                    </div>
                                </div>
                                {/* Auditor's Name */}
                                <div className="col-md-3 greybg form-floating">
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
                                            OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta, "divAuditorName" ) }}
                                            isRequired={true}
                                            disabled={this.state.isInputDisabled}
                                            noOptionsMessage="No Auditor's Names"
                                        />
                                    </div>
                                </div>
                                {/* Plant */}
                                <div className="col-md-3 greybg form-floating">
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
                                            OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta, "divPlant" ) }}
                                            isRequired={true}
                                            disabled={true}
                                            noOptionsMessage="No Plants available"
                                        />
                                    </div>
                                </div>
                                {/* Department */}
                                <div className="col-md-3 greybg form-floating">
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
                                            OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta, "divDepartment" ) }}
                                            isRequired={true}
                                            disabled={this.state.isInputDisabled}
                                            noOptionsMessage="No Departments"
                                        />
                                    </div>
                                </div>
                                {/* Zone */}
                                <div className="col-md-3 greybg form-floating">
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
                                            OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta, "divZone" ) }}
                                            isRequired={true}
                                            disabled={this.state.isInputDisabled}
                                            noOptionsMessage="No Zones"
                                        />
                                    </div>
                                </div>
                                {/* Machine */}
                                <div className="col-md-3 greybg form-floating">
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
                                                OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta, "divMachine" ) }}
                                                isRequired={ true }
                                                disabled={this.state.isInputDisabled}
                                                noOptionsMessage="No Machines"
                                            />
                                        </div>
                                    </div>
                                {/* Work Cell */}
                                <div className="col-md-3 greybg form-floating">
                                    <div className="custom-dropdown" id="divWorkCell" title={this.state.formData.WorkCell}>
                                        <SearchableDropdown
                                            label={"Work Cell"}
                                            Title={"WorkCell"}
                                            name={"WorkCell"}
                                            id="ddlWorkCell"
                                            placeholderText={""}
                                            className={""}
                                            selectedValue={this.state.formData.WorkCell}
                                            OptionsList={this.state.workCellData}
                                            OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta, "divWorkCell" ) }}
                                            isRequired={true}
                                            disabled={this.state.isInputDisabled}
                                            noOptionsMessage="No WorkCell"
                                        />
                                    </div>
                                </div>
                                {/* Tool No. */}
                                <div className="col-md-3 greybg form-floating">
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
                                            OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta, "divToolNumber" ) }}
                                            isRequired={true}
                                            disabled={this.state.isInputDisabled}
                                            noOptionsMessage="No ToolNumber's"
                                        />
                                    </div>
                                </div>
                                {/* Supervisor */}
                                <div className="col-md-3 greybg form-floating">
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
                                            OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta, "divSupervisor" ) }}
                                            isRequired={true}
                                            disabled={this.state.isInputDisabled}
                                            noOptionsMessage="No Supervisor's"
                                        />
                                    </div>
                                </div>

                                <table>
                                    <thead>
                                        <tr>
                                            <th>Requirement</th>
                                            <th>Select</th>
                                            <th>Comments</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        {/* {this.bindDynamicTable} */}
                                    </tbody>
                                </table>
                                {/* Comments */}
                                <div className="col-lg-12 col-md-12 col-sm-12 col-xs-12" style={{padding:"0px"}}>
                                    <div className="col-md-12 greybg">
                                        <div className={this.state.isInputDisabled? "textarea-disabled form-floating":"form-floating"} >
                                            <textarea className="form-control bs-textarea" rows={3} id="txtComments" name="Comments" ref={this.txtComments} placeholder="Comments" value={this.state.formData.Comments} onChange={this.handleChange} disabled={this.state.isInputDisabled} title="Comments" style={{height:"80px"}}></textarea>
                                            <span className="span-floating-textarea"></span>
                                            <label className=" col-form-label" htmlFor="txtComments">Comments </label>
                                        </div>
                                    </div>
                                </div>
                                {/* Supervisor Action Completed */}
                                <div className="col-lg-9 col-md-9 col-sm-9 col-xs-9" style={{padding:"0px"}}>
                                    <div className="col-md-12 greybg">
                                        <div className={this.state.isInputDisabled? "textarea-disabled form-floating":"form-floating"} >
                                            <textarea className="form-control bs-textarea" rows={3} id="txtActionCompleted" name="ActionCompleted" ref={this.txtActionCompleted} placeholder="ActionCompleted" value={this.state.formData.ActionCompleted} onChange={this.handleChange} disabled={this.state.isInputDisabled} title="Action Completed" style={{height:"80px"}}></textarea>
                                            <span className="span-floating-textarea"></span>
                                            <label className=" col-form-label" htmlFor="txtActionCompleted">Action Completed </label>
                                        </div>
                                    </div>
                                </div>
                                {/* Completed Date */}
                                <div className="col-md-3 greybg c-date-picker" id="divCompletedDate">
                                    <label className="label-datePicker" htmlFor="dtCompletedDate"> Completed Date <span className="text-danger">*</span></label>
                                    <DatePickercontrol placeholder="" selectedDate={this.state.formData.CompletedDate} id='dtCompletedDate' isDisabled={this.state.isInputDisabled} startDate={undefined} endDate={undefined} name="CompletedDate" onDatechange={(dateProps:any) => this.handleDateChange( dateProps[0], dateProps[2], "divCompletedDate", dateProps)} highlightDate={new Date()} showIcon />
                                </div>

                                <div className="col-sm-12 text-center py-3 greybg" id="">
                                    {this.state.showSubmit && <button type="button" id="btnSubmit" className="btn btn-primary mx-2" onClick={this.handleSubmit} >Submit</button>}
                                    <button type="button" id="btnCancel" className="btn btn-secondary" onClick={this.handlefullClose} >Cancel</button>
                                </div>
                            </div>
                        </div>
                    </div>
                </React.Fragment>
            )
        }
    }

}