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
import "../CSS/UCANForm.css";
import { ActionStatus, ControlType } from "../Constants/Contants";
import { showToast } from "../Shared/Toaster";
import { highlightCurrentNav } from "../Utilities/HighlightCurrentComponent";
import SearchableDropdown from "../Shared/Dropdown";
import DatePickercontrol from "../Shared/DatePickerField";
import { format } from "date-fns";
import DateUtilities from "../Utilities/DateUtilities";
import { initCommonFunctions } from "../Utilities/CommonFunctions";
import formValidation from "../Utilities/FormValidator";
// import FileUpload from "../Shared/FileUpload";
// import InputCheckBox from "../Shared/InputCheckBox";
// import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";
// import { faChevronDown, faChevronUp } from "@fortawesome/free-solid-svg-icons";

export interface UCANFormProps {
    match: any;
    spContext: any;
    spHttpClient: SPHttpClient;
    context: any;
    history: any;
    siteURL:string;
    webAbsoluteURL:string;
    currPlantTitle:string;
}

export interface UCANFormState {
}

export default class UCANForm extends React.Component<UCANFormProps, UCANFormState> {

    private currPlantObj:any;
    private txtReportedBy:any;
    private txtLocationPersons:any;
    private txtOriginialTagNo:any;
    private txtDate:any;
    private rdSafetyTag:any;
    private txtDateCompleted:any;
    private txtDescriptionOfIncident:any;
    private txtActionPlan:any;
    private txtActionCompleted:any;
    private sp = spfi().using(SPFx(this.props.context));

    private MaycoURL:string;
    private uaTypesList = "UATypes";
    private subTypesList = "UAMicroTypes";
    public state = {
        formData: {
            UCAN_x0020_Type:'',
            Plant:'',
            Department:'',
            Zone:'',
            Machine:'',
            Reported_x0020_By:'',
            Location_x002f_Persons:'',
            Safety_x0020_Tag:'',
            Completed:'',
            Date:'',
            Original_x0020_Tag_x0020_No_x002:'',
            UCAN_x0020_Number:'',
            Description_x0020_of_x0020_Incid:'',
            Action_x0020_Plan:'',
            Attachment:'',
            Date_x0020_Completed:'',
            Sub_x002d_Type:'',
            UAType:'',
            Action_x0020_Completed:'',
            Shift:'',
            Year:'',
            YearMonth:'',
            YearMonthCalc:''
        },
        ucanTypeData:[],
        plantsData:[],
        departmentData:[],
        departmentOptions:[],
        zoneData:[],
        zoneOptions:[],
        machineData:[],
        machineOptions:[],
        uaTypeData:[],
        subTypeData:[],
        subTypeOptions:[],
        shiftData:[],
        isInputDisabled: false,
        isEditForm: false,
        ItemId:0,
        Homeredirect: false,
    }

    constructor(props: UCANFormProps) {
        super(props);
        this.MaycoURL=`${this.props.siteURL}/mayco`;

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
            var formData = {...this.state.formData};

            let  { getListItems } = initCommonFunctions(this.props.context,this.props.siteURL);
            let PlantList='Plant', PlantSelQuery='Title,*',plantFiltQuery='',PlantExpFields='';
            let DepartmentList='Department', DepartmentSelQuery='Title,Plant/Title,Plant/Id,*',DepartmentFiltQuery='',DepartmentExpFields='Plant';
            let ZoneList='Zones', ZoneSelQuery='Title,Plant/Title,Plant/Id,Department/Title,Department/Id,*',ZoneFiltQuery='',ZoneExpFields='Plant,Department';
            let MachineList='Machines', MachineSelQuery='Title,Plant/Title,Plant/Id,Department/Title,Department/Id,Zone/Title,Zone/Id,*',MachineFiltQuery='',MachineExpFields='Plant,Department,Zone';
            let ShiftsList='Shifts', ShiftsSelQuery='Title,*',ShiftsFiltQuery='',ShiftsExpFields='';
            let [Plants,departmentData,zoneData,machineData, uaTypes, subTypeData, shifts] = await Promise.all([
                getListItems(PlantList,PlantSelQuery,plantFiltQuery,PlantExpFields,this.MaycoURL),
                getListItems(DepartmentList,DepartmentSelQuery,DepartmentFiltQuery,DepartmentExpFields,this.MaycoURL),
                getListItems(ZoneList,ZoneSelQuery,ZoneFiltQuery,ZoneExpFields,this.MaycoURL),
                getListItems(MachineList,MachineSelQuery,MachineFiltQuery,MachineExpFields,this.MaycoURL), 
                this.sp.web.lists.getByTitle(this.uaTypesList).items.orderBy("Title").top(2000)(), 
                this.sp.web.lists.getByTitle(this.subTypesList).items.orderBy("Title").top(2000)(), 
                getListItems(ShiftsList,ShiftsSelQuery,ShiftsFiltQuery,ShiftsExpFields,this.MaycoURL)
            ]);

            let plantsData = Plants.map((item: any) => ({ label: item.Title, value: item.Title }));
            this.currPlantObj = plantsData.find( (plant:any) => plant.label.toLowerCase() == this.props.currPlantTitle);
            formData.Plant = this.currPlantObj.label;
            
            let departmentOptions = departmentData.filter( (option:any) => option.Plant.Title == formData.Plant ).map((item: any) => ({ label: item.Title, value: item.Title, id:item.Id }))
            let zoneOptions:any = [];
            let machineOptions:any = [];
            let uaTypeData = uaTypes.map((item: any) => ({ label: item.Title, value: item.Id })); 
            let subTypeOptions:any = [];
            let shiftData = shifts.map((item: any) => ({ label: item.Title, value: item.Id })); 

            this.setState({formData, plantsData, departmentData, departmentOptions, zoneData, zoneOptions, machineData, machineOptions, uaTypeData, subTypeData, subTypeOptions, shiftData});

            hideLoader();
        } catch (e) {
            console.log(e);
            this.onError();
        }
    }

    private handleChange= (event: any) => {
        const formData:any = {...this.state.formData};

        const name = event.target.name;
        let inputValue = (event.target.type == "text" || event.target.type == "textarea") ? event.target.value : event.target.checked ? event.target.value : '';
        let classArr: any = event.target.className;

        if (classArr.includes("onlyNum")) {
            inputValue = inputValue.replace(/[^0-9]/g, '');
        }
        formData[name] = inputValue;
        this.setState({formData });
    }

    private handleDropdownChange = ( event: any, actionMeta: any, id:any ) => {
        const formData:any = {...this.state.formData};
        let departmentData = [...this.state.departmentData];
        let departmentOptions:any = [...this.state.departmentOptions];
        let zoneData = [...this.state.zoneData];
        let zoneOptions:any = [...this.state.zoneOptions];
        let machineData = [...this.state.machineData];
        let machineOptions:any = [...this.state.machineOptions];
        let subTypeData = [...this.state.subTypeData];
        let subTypeOptions:any = [...this.state.subTypeOptions];
        const name = actionMeta.name;
        const value = actionMeta.action == "clear" ? '' : event.value;
        formData[name] = value;

        if( name == "Department" ){
            formData.Zone = "";
            formData.Machine = "";

            if( actionMeta.action != "clear"){
                zoneOptions = zoneData.filter( (option:any) => option.Department.Id == event.id ).map((item: any) => ({ label: item.Title, value: item.Title, id:item.Id }));
            }
            else{ zoneOptions = [];}
            machineOptions = [];
        }
        else if( name == "Zone" ){
            formData.Machine = "";

            if( actionMeta.action != "clear"){
                machineOptions = machineData.filter( (option:any) => option.Zone.Id == event.id ).map((item: any) => ({ label: item.Title, value: item.Title, id:item.Id }));
            }
            else{ machineOptions = [];}
        }
        else if( name == "UAType" ){
            formData.Sub_x002d_Type = "";

            if( actionMeta.action != "clear"){
                subTypeOptions = subTypeData.filter( (option:any) => option.UAType.Id == event.id ).map((item: any) => ({ label: item.Title, value: item.Id, }));
            }
            else{ subTypeOptions = [];}
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

        this.setState({formData, departmentData, departmentOptions, zoneData, zoneOptions, machineData, machineOptions, subTypeData, subTypeOptions });
    }
    private handleDateChange = (dateValue: any, name:any, divId:any) => {
        const formData: any = { ...this.state.formData };

        if( !([null, undefined, ''].includes(divId)) ){
            var ddlElement = document.getElementById(divId);
            if( !([null, undefined, ''].includes(dateValue)) ){
                ddlElement?.classList.add("active");
            }
            else{
                ddlElement?.classList.remove("active");
            }
        }

       dateValue = format( DateUtilities.addBrowserwrtServer( new Date(DateUtilities.getDateMMDDYYYY(dateValue)), this.props.spContext.webTimeZoneData).toISOString(), "MM/dd/yyyy");

        formData[name] = dateValue;

        this.setState({ formData });
    }

    private onError = () => {
        showToast("error", ActionStatus.Error);
        hideLoader();
    }

    public handleSubmit = () => {
        var formData = {...this.state.formData};
        var data = {
            ucanType: {val: this.state.formData.UCAN_x0020_Type, required: true, Name: "Near miss, Unsafe act, Unsafe condition", Type: ControlType.reactSelect, Focusid: "ddlUCANType"},
            plant: {val: this.state.formData.Plant, required: true, Name: "Plant", Type: ControlType.reactSelect, Focusid: "ddlPlant"},
            department: {val: this.state.formData.Department, required: true, Name: "Department", Type: ControlType.reactSelect, Focusid: "ddlDepartment"},
            zone: {val: this.state.formData.Zone, required: true, Name: "Zone", Type: ControlType.reactSelect, Focusid: "ddlZone"},
            machine: {val: this.state.formData.Machine, required: (formData.UCAN_x0020_Type != "Unsafe Act"), Name: "Machine", Type: ControlType.reactSelect, Focusid: "ddlMachine"},
            uaType: {val: this.state.formData.UAType, required: true, Name: "UA Type", Type: ControlType.reactSelect, Focusid: "ddlUAType"},
            subType: {val: this.state.formData.Sub_x002d_Type, required: true, Name: "Sub Type", Type: ControlType.reactSelect, Focusid: "ddlSubType"},
            shift: {val: this.state.formData.Shift, required: true, Name: "Shift", Type: ControlType.reactSelect, Focusid: "ddlShift"},
            date: {val: this.state.formData.Date, required: true, Name: "Date", Type: ControlType.date, Focusid: "divDate"},
            dateToday: {val: this.state.formData.Date, required: true, Name: "Date", Type: ControlType.todayDate, Focusid: "divDate"},
            dateCompletedToday: {val: this.state.formData.Date_x0020_Completed, required: true, Name: "Date", Type: ControlType.todayDate, Focusid: "divDateCompleted"},
            dateComparision: {startDate: this.state.formData.Date_x0020_Completed, endDate: this.state.formData.Date_x0020_Completed, required: true, startDateName: "Date", endDateName:"Date Completed", Type: ControlType.todayDate, Focusid: "divDateCompleted"},
            actionCompleted: {val: this.state.formData.Action_x0020_Completed, required: (formData.Date_x0020_Completed != ''), Name: "Action Completed", Type: ControlType.string, Focusid: this.txtActionCompleted}
        }

        let isValid = formValidation.FormValidation(data);
        
        if( isValid.status ){
            console.log("Valid Data");
        }
        else{
            showToast("error", isValid.message);
            hideLoader();
        }
    }

    private handlefullClose = () => {
        this.setState({ Homeredirect: true, ItemID: 0 });
    }

    public render() {
        if (this.state.Homeredirect) {
            let url = "/Home";
            return (<Navigate to={url} />)
        }
        else {
            return (
                <React.Fragment>
                    <div className="container-fluid">
                        <div className="light-box border-box-shadow">
                            <div className="m-0 titlebg">
                                <h4 className="mb-0 pt-2 text-center">{" Unsafe Conditions And Acts Form "+ (this.state.isEditForm ? (" - "+ this.state.ItemId): "")} </h4>
                                <label className="text-end px-1" style={{width:"100%"}}> <span className="text-danger">* </span> are mandatory fields</label>
                            </div>

                            <div className="mainContent row col-lg-12 col-md-12 col-sm-12 col-xs-12 borderLine">
                                <div className="col-lg-8 col-md-8 col-sm-8 col-xs-8 pull-left row">
                                    {/* Near miss, Unsafe act, Unsafe condition */}
                                    <div className="col-md-6 greybg form-floating">
                                        <div className="custom-dropdown" id="divUCANType">
                                            <SearchableDropdown
                                                label={"Near miss, Unsafe act, Unsafe condition"}
                                                Title={"Near miss, Unsafe act, Unsafe condition"}
                                                name={"UCAN_x0020_Type"}
                                                id="ddlUCANType"
                                                placeholderText={""}
                                                className={""}
                                                selectedValue={this.state.formData.UCAN_x0020_Type}
                                                optionLabel={""}
                                                optionValue={""}
                                                OptionsList={[{label:"Near Miss", value:"Near Miss"}, {label:"Unsafe Act", value:"Unsafe Act"}, {label:"Unsafe Condition", value:"Unsafe Condition"}]}
                                                OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta, "divUCANType" ) }}
                                                isRequired={true}
                                                disabled={this.state.isInputDisabled}
                                                noOptionsMessage="No UCAN Types"
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
                                                optionLabel={""}
                                                optionValue={""}
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
                                                optionLabel={""}
                                                optionValue={""}
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
                                                optionLabel={""}
                                                optionValue={""}
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
                                                optionLabel={""}
                                                optionValue={""}
                                                OptionsList={this.state.machineOptions}
                                                OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta, "divMachine" ) }}
                                                isRequired={ this.state.formData.UCAN_x0020_Type != "Unsafe Act"}
                                                disabled={this.state.isInputDisabled}
                                                noOptionsMessage="No Machines"
                                            />
                                        </div>
                                    </div>
                                    {/*UA Type */}
                                    <div className="col-md-3 greybg form-floating">
                                        <div className="custom-dropdown" id="divType">
                                            <SearchableDropdown
                                                label={"Type"}
                                                Title={"Type"}
                                                name={"UAType"}
                                                id="ddlUAType"
                                                placeholderText={""}
                                                className={""}
                                                selectedValue={this.state.formData.UAType}
                                                optionLabel={""}
                                                optionValue={""}
                                                OptionsList={this.state.uaTypeData}
                                                OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta, "divType" ) }}
                                                isRequired={true}
                                                disabled={this.state.isInputDisabled}
                                                noOptionsMessage="No Types"
                                            />
                                        </div>
                                    </div>
                                    {/* Sub-Type */}
                                    <div className="col-md-3 greybg form-floating">
                                        <div className="custom-dropdown" id="divSubType">
                                            <SearchableDropdown
                                                label={"Sub-Type"}
                                                Title={"Sub-Type"}
                                                name={"Sub_x002d_Type"}
                                                id="ddlSubType"
                                                placeholderText={""}
                                                className={""}
                                                selectedValue={this.state.formData.Sub_x002d_Type}
                                                optionLabel={""}
                                                optionValue={""}
                                                OptionsList={this.state.subTypeOptions}
                                                OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta, "divSubType" ) }}
                                                isRequired={true}
                                                disabled={this.state.isInputDisabled}
                                                noOptionsMessage="No SubTypes"
                                            />
                                        </div>
                                    </div>
                                    <div className="col-lg-12 col-md-12 col-sm-12 col-xs-12" style={{padding:"0px"}}>
                                        <div className="col-lg-3 col-md-3 col-sm-3 col-xs-3 pull-left">
                                            {/* Reported By */}
                                            <div className="col-lg-12 col-md-12 col-sm-12 col-xs-12 greybg">
                                                <div className="form-floating">
                                                    <input className="form-control" placeholder="Reported By" name="Reported_x0020_By" type="text" id="txtReportedBy" ref={this.txtReportedBy} value={this.state.formData.Reported_x0020_By} onChange={this.handleChange} disabled={this.state.isInputDisabled} title="Reported By" />
                                                    <label className=" col-form-label" htmlFor="txtReportedBy">Reported By </label>
                                                </div>
                                            </div>
                                            {/* Original Tag No. */}
                                            <div className="col-lg-12 col-md-12 col-sm-12 col-xs-12 greybg">
                                                 <div className="form-floating">
                                                    <input className="form-control" placeholder="Original Tag No." name="Original_x0020_Tag_x0020_No_x002" type="text" id="txtOriginialTagNo" ref={this.txtOriginialTagNo} value={this.state.formData.Original_x0020_Tag_x0020_No_x002} onChange={this.handleChange} disabled={this.state.isInputDisabled} title="Original Tag No." />
                                                    <label className=" col-form-label" htmlFor="txtOriginialTagNo">Original Tag No. </label>
                                                </div>
                                            </div>
                                            {/* Safety Tag */}
                                            <div className="col-lg-12 col-md-12 col-sm-12 col-xs-12 greybg" style={{textAlign: "center", height: "64px"}}>
                                                <label className=" col-form-label" htmlFor="rdSafetyTag">Safety Tag </label>
                                                <input className="" placeholder="Safety Tag" name="Safety_x0020_Tag" type="checkbox" id="rdSafetyTag" ref={this.rdSafetyTag} value={this.state.formData.Safety_x0020_Tag} onChange={this.handleChange} disabled={this.state.isInputDisabled} title="Safety Tag" />
                                            </div>
                                        </div>
                                        <div className="col-lg-6 col-md-6 col-sm-6 col-xs-6 pull-left">
                                            {/* Location/Persons */}
                                            <div className="col-md-12 greybg">
                                                <div className={this.state.isInputDisabled? "textarea-disabled form-floating":"form-floating"} >
                                                    <textarea className="form-control bs-textarea" rows={3} id="txtLocationPersons" name="Location_x002f_Persons" ref={this.txtLocationPersons} placeholder="Location/Persons" value={this.state.formData.Location_x002f_Persons} onChange={this.handleChange} disabled={this.state.isInputDisabled} title="Location/Persons" style={{height:"186px"}}></textarea>
                                                    <span className="span-floating-textarea"></span>
                                                    <label className=" col-form-label" htmlFor="txtLocationPersons">Location/Persons </label>
                                                </div>
                                            </div>
                                        </div>
                                        <div className="col-lg-3 col-md-3 col-sm-3 col-xs-3 greybg pull-left">
                                            {/* Shift */}
                                            <div className="form-floating">
                                                <div className="custom-dropdown" id="divShift">
                                                    <SearchableDropdown
                                                        label={"Shift"}
                                                        Title={"Shift"}
                                                        name={"Shift"}
                                                        id="ddlShift"
                                                        placeholderText={""}
                                                        className={""}
                                                        selectedValue={this.state.formData.Shift}
                                                        optionLabel={""}
                                                        optionValue={""}
                                                        OptionsList={this.state.shiftData}
                                                        OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta, "divShift" ) }}
                                                        isRequired={true}
                                                        disabled={this.state.isInputDisabled}
                                                        noOptionsMessage="No Shifts"
                                                    />
                                                </div>
                                            </div>
                                            {/* Date */}
                                            <div className="c-date-picker" id="divDate" style={{margin:"8px 0"}}>
                                                <label className="label-datePicker" htmlFor="dtDate"> Date <span className="text-danger">*</span></label>
                                                <DatePickercontrol placeholder="" selectedDate={this.state.formData.Date} id='dtDate' isDisabled={this.state.isInputDisabled} startDate={undefined} endDate={new Date()} name="Date" onDatechange={(dateProps:any) => this.handleDateChange( dateProps[0], dateProps[2], "divDate")} ref={this.txtDate} highlightDate={new Date()} showIcon />
                                            </div>
                                            {/* Date Completed */}
                                            <div className="c-date-picker" id="divDateCompleted">
                                                <label className="label-datePicker" htmlFor="dtDateCompleted"> Date Completed </label>
                                                <DatePickercontrol placeholder="" selectedDate={this.state.formData.Date_x0020_Completed} id='dtDateCompleted' isDisabled={this.state.isInputDisabled} startDate={undefined} endDate={new Date()} name="Date_x0020_Completed" onDatechange={(dateProps:any) => this.handleDateChange( dateProps[0], dateProps[2], "divDateCompleted")} ref={this.txtDateCompleted} highlightDate={new Date()} showIcon />
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                <div className="col-lg-4 col-md-4 col-sm-4 col-xs-4 pull-right" style={{padding:"0px"}}>
                                    <div className="col-md-12 greybg">
                                        <div className={this.state.isInputDisabled? "textarea-disabled form-floating":"form-floating"} >
                                            <textarea className="form-control bs-textarea" rows={3} id="txtDescriptionOfIncident" name="Description_x0020_of_x0020_Incid" ref={this.txtDescriptionOfIncident} placeholder="Description of Incident" value={this.state.formData.Description_x0020_of_x0020_Incid} onChange={this.handleChange} disabled={this.state.isInputDisabled} title="Description of Incident" style={{height:"312px"}}></textarea>
                                            <span className="span-floating-textarea"></span>
                                            <label className=" col-form-label" htmlFor="txtDescriptionOfIncident">Description of Incident </label>
                                        </div>
                                    </div>
                                </div>
                                <div className="col-lg-12 col-md-12 col-sm-12 col-xs-12" style={{padding:"0px"}}>
                                    <div className="col-md-12 greybg">
                                        <div className={this.state.isInputDisabled? "textarea-disabled form-floating":"form-floating"} >
                                            <textarea className="form-control bs-textarea" rows={3} id="txtActionPlan" name="Action_x0020_Plan" ref={this.txtActionPlan} placeholder="Action Plan" value={this.state.formData.Action_x0020_Plan} onChange={this.handleChange} disabled={this.state.isInputDisabled} title="Action Plan" style={{height:"80px"}}></textarea>
                                            <span className="span-floating-textarea"></span>
                                            <label className=" col-form-label" htmlFor="txtActionPlan">Action Plan <span className="text-danger">*</span></label>
                                        </div>
                                    </div>
                                </div>
                                <div className="col-lg-12 col-md-12 col-sm-12 col-xs-12" style={{padding:"0px"}}>
                                    <div className="col-md-12 greybg">
                                        <div className={this.state.isInputDisabled? "textarea-disabled form-floating":"form-floating"} >
                                            <textarea className="form-control bs-textarea" rows={3} id="txtActionCompleted" name="Action_x0020_Completed" ref={this.txtActionCompleted} placeholder="Action Completed" value={this.state.formData.Action_x0020_Completed} onChange={this.handleChange} disabled={this.state.isInputDisabled} title="Action Completed" style={{height:"80px"}}></textarea>
                                            <span className="span-floating-textarea"></span>
                                            <label className=" col-form-label" htmlFor="txtActionCompleted">Action Completed <span className="text-danger">*</span></label>
                                        </div>
                                    </div>
                                </div>
                                <div className="col-sm-12 text-center py-3 greybg" id="">
                                        <button type="button" id="btnSubmit" className="btn btn-primary mx-2" onClick={this.handleSubmit} >Submit</button>
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