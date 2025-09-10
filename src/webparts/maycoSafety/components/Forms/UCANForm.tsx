import * as React from "react";
import { hideLoader, showLoader } from "../Shared/Loader";
import { Navigate } from "react-router-dom";
import { SPHttpClient } from "@microsoft/sp-http";
// import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "../CSS/UCANForm.css";
import { ActionStatus } from "../Constants/Contants";
import { showToast } from "../Shared/Toaster";
import { highlightCurrentNav } from "../Utilities/HighlightCurrentComponent";
import SearchableDropdown from "../Shared/Dropdown";
import DatePickercontrol from "../Shared/DatePickerField";
import { format } from "date-fns";
import DateUtilities from "../Utilities/DateUtilities";
// import SearchableDropdown from "../Shared/Dropdown";
// import FileUpload from "../Shared/FileUpload";
// import InputCheckBox from "../Shared/InputCheckBox";
// import { format } from "date-fns";
// import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";
// import { faChevronDown, faChevronUp } from "@fortawesome/free-solid-svg-icons";

export interface UCANFormProps {
    match: any;
    spContext: any;
    spHttpClient: SPHttpClient;
    context: any;
    history: any;
    isSupplierTeam: boolean;
    isDTETeam: boolean;
    isProcurementTeam: boolean;
}

export interface UCANFormState {
}

export default class UCANForm extends React.Component<UCANFormProps, UCANFormState> {

    private siteURL: string;
    private ddlUCANType:any;
    private ddlPlant:any;
    private ddlDepartment:any;
    private ddlZone:any;
    private ddlMachine:any;
    private ddlUAType:any;
    private ddlSubType:any;
    private txtReportedBy:any;
    private txtLocationPersons:any;
    private ddlShift:any;
    private txtOriginialTagNo:any;
    private txtDate:any;
    private rdSafetyTag:any;
    private txtDateCompleted:any;
    private txtDescriptionOfIncident:any;
    private txtActionPlan:any;
    private txtActionCompleted:any;
    // private sp = spfi().using(SPFx(this.props.context));
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
        zoneData:[],
        machineData:[],
        uaTypeData:[],
        subTypeData:[],
        isInputDisabled: false,
        isEditForm: false,
        ItemId:0,
        Homeredirect: false,
    }

    constructor(props: UCANFormProps) {
        super(props);
        this.siteURL = this.props.spContext.siteAbsoluteUrl;
        console.log(this.siteURL);

        this.ddlUCANType = React.createRef();
        this.ddlPlant = React.createRef();
        this.ddlDepartment = React.createRef();
        this.ddlZone = React.createRef();
        this.ddlMachine = React.createRef();
        this.ddlUAType = React.createRef();
        this.ddlSubType = React.createRef();
        this.txtReportedBy = React.createRef();
        this.txtLocationPersons = React.createRef();
        this.ddlShift = React.createRef();
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


            hideLoader();
        } catch (e) {
            console.log(e);
            this.onError();
        }
    }

    private handleChange= (event: any, actionMeta?: any, id?:any) => {

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

        if( name == "AssignmentsStartDate" ){
            formData.AssignmentStartDate = format( this.addBrowserwrtServer( new Date(DateUtilities.getDateMMDDYYYY(dateValue))).toISOString(), "MM/dd/yyyy");
        }
        else if( name == "AssignmentsEndDate" ){
            formData.AssignmentEndDate =  format( this.addBrowserwrtServer( new Date(DateUtilities.getDateMMDDYYYY(dateValue))).toISOString(), "MM/dd/yyyy");
        }

        formData[name] = dateValue;

        this.setState({ formData });
    }

    private addBrowserwrtServer(date:Date) {
        var utcOffsetMinutes = date.getTimezoneOffset();
        var newDate = new Date(date.getTime());
        newDate.setTime(newDate.getTime() + ((this.props.spContext.webTimeZoneData.Bias - utcOffsetMinutes + this.props.spContext.webTimeZoneData.DaylightBias) * 60 * 1000));
        return newDate;
    }

    private onError = () => {
        showToast("error", ActionStatus.Error);
        hideLoader();
    }

    public handleSubmit = () => {

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
                                                OptionsList={this.state.ucanTypeData}
                                                OnChange={(selectedOption: any, actionMeta: any) => { this.handleChange(selectedOption, actionMeta, "divUCANType" ) }}
                                                isRequired={true}
                                                disabled={this.state.isInputDisabled}
                                                refElement={this.ddlUCANType}
                                                noOptionsMessage="No UCAN Types"
                                            />
                                        </div>
                                    </div>
                                    {/* Plant */}
                                    <div className="col-md-3 greybg form-floating">
                                        <div className="custom-dropdown" id="divPlant">
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
                                                OnChange={(selectedOption: any, actionMeta: any) => { this.handleChange(selectedOption, actionMeta, "divPlant" ) }}
                                                isRequired={true}
                                                disabled={this.state.isInputDisabled}
                                                refElement={this.ddlPlant}
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
                                                OptionsList={this.state.departmentData}
                                                OnChange={(selectedOption: any, actionMeta: any) => { this.handleChange(selectedOption, actionMeta, "divDepartment" ) }}
                                                isRequired={true}
                                                disabled={this.state.isInputDisabled}
                                                refElement={this.ddlDepartment}
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
                                                OptionsList={this.state.zoneData}
                                                OnChange={(selectedOption: any, actionMeta: any) => { this.handleChange(selectedOption, actionMeta, "divZone" ) }}
                                                isRequired={true}
                                                disabled={this.state.isInputDisabled}
                                                refElement={this.ddlZone}
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
                                                OptionsList={this.state.machineData}
                                                OnChange={(selectedOption: any, actionMeta: any) => { this.handleChange(selectedOption, actionMeta, "divMachine" ) }}
                                                isRequired={true}
                                                disabled={this.state.isInputDisabled}
                                                refElement={this.ddlMachine}
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
                                                OnChange={(selectedOption: any, actionMeta: any) => { this.handleChange(selectedOption, actionMeta, "divType" ) }}
                                                isRequired={true}
                                                disabled={this.state.isInputDisabled}
                                                refElement={this.ddlUAType}
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
                                                OptionsList={this.state.subTypeData}
                                                OnChange={(selectedOption: any, actionMeta: any) => { this.handleChange(selectedOption, actionMeta, "divSubType" ) }}
                                                isRequired={true}
                                                disabled={this.state.isInputDisabled}
                                                refElement={this.ddlSubType}
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
                                                    <label className=" col-form-label" htmlFor="txtReportedBy">Reported By <span className="text-danger">*</span></label>
                                                </div>
                                            </div>
                                            {/* Original Tag No. */}
                                            <div className="col-lg-12 col-md-12 col-sm-12 col-xs-12 greybg">
                                                 <div className="form-floating">
                                                    <input className="form-control" placeholder="Original Tag No." name="Original_x0020_Tag_x0020_No_x002" type="text" id="txtOriginialTagNo" ref={this.txtOriginialTagNo} value={this.state.formData.Original_x0020_Tag_x0020_No_x002} onChange={this.handleChange} disabled={this.state.isInputDisabled} title="Original Tag No." />
                                                    <label className=" col-form-label" htmlFor="txtOriginialTagNo">Original Tag No. <span className="text-danger">*</span></label>
                                                </div>
                                            </div>
                                            {/* Safety Tag */}
                                            <div className="col-lg-12 col-md-12 col-sm-12 col-xs-12 greybg">
                                                 <div className="form-floating">
                                                    <label className=" col-form-label" htmlFor="rdSafetyTag">Safety Tag <span className="text-danger">*</span></label>
                                                    <input className="form-control" placeholder="Safety Tag" name="Safety_x0020_Tag" type="text" id="rdSafetyTag" ref={this.rdSafetyTag} value={this.state.formData.Safety_x0020_Tag} onChange={this.handleChange} disabled={this.state.isInputDisabled} title="Safety Tag" />
                                                </div>
                                            </div>
                                        </div>
                                        <div className="col-lg-6 col-md-6 col-sm-6 col-xs-6 pull-left">
                                            {/* Location/Persons */}
                                            <div className="col-md-12 greybg">
                                                <div className={this.state.isInputDisabled? "textarea-disabled form-floating":"form-floating"} >
                                                    <textarea className="form-control bs-textarea" rows={3} id="txtLocationPersons" name="Location_x002f_Persons" ref={this.txtLocationPersons} placeholder="Location/Persons" value={this.state.formData.Location_x002f_Persons} onChange={this.handleChange} disabled={this.state.isInputDisabled} title="Location/Persons" style={{height:"186px"}}></textarea>
                                                    <span className="span-floating-textarea"></span>
                                                    <label className=" col-form-label" htmlFor="txtLocationPersons">Location/Persons <span className="text-danger">*</span></label>
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
                                                        OptionsList={[1,2,3]}
                                                        OnChange={(selectedOption: any, actionMeta: any) => { this.handleChange(selectedOption, actionMeta, "divShift" ) }}
                                                        isRequired={true}
                                                        disabled={this.state.isInputDisabled}
                                                        refElement={this.ddlShift}
                                                        noOptionsMessage="No Shifts"
                                                    />
                                                </div>
                                            </div>
                                            {/* Date */}
                                            <div className="c-date-picker" id="divDate" style={{margin:"8px 0"}}>
                                                <label className="label-datePicker" htmlFor="dtDate"> Date <span className="text-danger">*</span></label>
                                                <DatePickercontrol placeholder="" selectedDate={this.state.formData.Date} id='dtDate' isDisabled={this.state.isInputDisabled} startDate={undefined} endDate={undefined} name="Date" onDatechange={(dateProps:any) => this.handleDateChange( dateProps[0], dateProps[2], "divDate")} ref={this.txtDate} highlightDate={new Date()} showIcon />
                                            </div>
                                            {/* Date Completed */}
                                            <div className="c-date-picker" id="divDateCompleted">
                                                <label className="label-datePicker" htmlFor="dtDateCompleted"> Date Completed <span className="text-danger">*</span></label>
                                                <DatePickercontrol placeholder="" selectedDate={this.state.formData.Date_x0020_Completed} id='dtDateCompleted' isDisabled={this.state.isInputDisabled} startDate={undefined} endDate={undefined} name="Date_x0020_Completed" onDatechange={(dateProps:any) => this.handleDateChange( dateProps[0], dateProps[2], "divDateCompleted")} ref={this.txtDateCompleted} highlightDate={new Date()} showIcon />
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                <div className="col-lg-4 col-md-4 col-sm-4 col-xs-4 pull-right" style={{padding:"0px"}}>
                                    <div className="col-md-12 greybg">
                                        <div className={this.state.isInputDisabled? "textarea-disabled form-floating":"form-floating"} >
                                            <textarea className="form-control bs-textarea" rows={3} id="txtDescriptionOfIncident" name="Description_x0020_of_x0020_Incid" ref={this.txtDescriptionOfIncident} placeholder="Description of Incident" value={this.state.formData.Description_x0020_of_x0020_Incid} onChange={this.handleChange} disabled={this.state.isInputDisabled} title="Description of Incident" style={{height:"312px"}}></textarea>
                                            <span className="span-floating-textarea"></span>
                                            <label className=" col-form-label" htmlFor="txtDescriptionOfIncident">Description of Incident <span className="text-danger">*</span></label>
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