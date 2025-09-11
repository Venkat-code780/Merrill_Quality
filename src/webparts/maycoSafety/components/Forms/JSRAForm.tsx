import * as React from "react";
import { hideLoader, showLoader } from "../Shared/Loader";
import { Navigate } from "react-router-dom";
import { SPHttpClient } from "@microsoft/sp-http";
// import { SPFI, spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "../CSS/JSRAForm.css";
import { ActionStatus } from "../Constants/Contants";
import { showToast } from "../Shared/Toaster";
import { highlightCurrentNav } from "../Utilities/HighlightCurrentComponent";
import DatePickercontrol from "../Shared/DatePickerField";
import SearchableDropdown from "../Shared/Dropdown";
import { initCommonFunctions } from "../Utilities/CommonFunctions";
// import FileUpload from "../Shared/FileUpload";
// import InputCheckBox from "../Shared/InputCheckBox";
// import { format } from "date-fns";
// import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";
// import { faChevronDown, faChevronUp } from "@fortawesome/free-solid-svg-icons";

export interface JSRAFormProps {
    match: any;
    spContext: any;
    spHttpClient: SPHttpClient;
    context: any;
    history: any;
}

export interface JSRAFormState {
}

export default class JSRAForm extends React.Component<JSRAFormProps, JSRAFormState> {

    private siteURL: string;private MaycoURL:string;
    // private sp = spfi().using(SPFx(this.props.context)); 
    public state = {
        //Cascaded dropdowns
        Plants:[],Departments:[], Zones:[], WorkCells:[], Machines:[],Shifts:[], Supervisors:[],
        PlantsOpt:[],DepartmentsOpt:[], ZonesOpt:[], WorkCellsOpt:[], MachinesOpt:[],ShiftsOpt:[], SupervisorsOpt:[],
        formData: {

        },
        Homeredirect: false,
    }

    constructor(props: JSRAFormProps) {
        super(props);
        this.siteURL = this.props.spContext.siteAbsoluteUrl;
        this.MaycoURL=`${this.siteURL}/mayco`;
        console.log(this.siteURL);      //sites/wcm
    }

    public componentDidMount(): void {
        highlightCurrentNav("liJSRAForm");
        document.title = "Mayco - Safety | JSRA";
        this.getOnLoadData();
    }

    private getOnLoadData = async () => {
        let  { getListItems } = initCommonFunctions(this.props.context,this.siteURL);
        let PlantList='Plant', PlantSelQuery='Title,*',plantFiltQuery='',PlantExpFields='';
        let DepartmentList='Department', DepartmentSelQuery='Title,Plant/Title,Plant/Id,*',DepartmentFiltQuery='',DepartmentExpFields='Plant';
        let ZoneList='Zones', ZoneSelQuery='Title,Plant/Title,Plant/Id,Department/Title,Department/Id,*',ZoneFiltQuery='',ZoneExpFields='Plant,Department';
        let MachineList='Machines', MachineSelQuery='Title,Plant/Title,Plant/Id,Department/Title,Department/Id,Zone/Title,Zone/Id,*',MachineFiltQuery='',MachineExpFields='Plant,Department,Zone';
        try {
            showLoader();
            let  [Plants,Departments,Zones,Machines]=await Promise.all([
            // this.MaycoSP.web.lists.getByTitle('Plant').items.top(2000).select('Title,*').expand('').orderBy("Title", true)(),
           getListItems(PlantList,PlantSelQuery,plantFiltQuery,PlantExpFields,this.MaycoURL),
           getListItems(DepartmentList,DepartmentSelQuery,DepartmentFiltQuery,DepartmentExpFields,this.MaycoURL),
           getListItems(ZoneList,ZoneSelQuery,ZoneFiltQuery,ZoneExpFields,this.MaycoURL),
           getListItems(MachineList,MachineSelQuery,MachineFiltQuery,MachineExpFields,this.MaycoURL)
           ])
           let PlantsOpt=Plants.map((item: any) => ({
        label: item.Title,
        value: item.Id,
             }));
             // for sorting
             PlantsOpt.sort((a,b)=> a.label.localeCompare(b.label));
         this.setState({Plants,PlantsOpt,Departments,Zones,Machines});
         console.log(Plants);
         console.log(Departments);
         console.log(Zones);
         console.log(Machines);
            hideLoader();
        } catch (e) {
            console.log(e);
            this.onError();
        }
    }

    private onError = () => {
        showToast("error", ActionStatus.Error);
        hideLoader();
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
                        <div className="light-box border-box-shadow mb-4 brd-f1f1f1">
                            <div className="m-0 px-2">
                                <div className="col-12"><h2 className="mb-0 mt-2 text-center"> JSRA</h2></div>
                            </div>

                            <div className="row">
                                <label className="text-end px-4"> <span className="text-danger">* </span> are mandatory fields</label>
                            </div>


                            <div className="divContactInformation divSections light-box border-box-shadow border" id="Form-0">
                                <div className="m-1">
                                    <div className="row px-2">
                                        <div className="col-md-3 bs-field">
                                            <div className="form-floating">
                                                {/* <DatePickercontrol placeholder="Date" selectedDate={''} id='JSRADate' isDisabled={false} startDate={undefined} endDate={undefined} name="Date" onDatechange={(dateProps:any) => this.handleDateChange( dateProps[0], dateProps[2], "divJSRADate")} ref={this.JSRADate} highlightDate={new Date()} showIcon /> */}
                                                <label className=" col-form-label" htmlFor="txtDate">Date<span className="text-danger">*</span></label>
                                                <DatePickercontrol placeholder="MM/DD/YYYY" selectedDate={''} id='JSRADate' isDisabled={false} startDate={undefined} endDate={undefined} name="Date" onDatechange={(dateProps: any) => null} highlightDate={new Date()} showIcon />
                                            </div>
                                        </div>
                                        <div className="col-md-3 bs-field form-floating" title={'Plant'}>
                                            <div className="custom-dropdown" id="divPlant">
                                                <SearchableDropdown
                                                    label={"Plant"}
                                                    Title={"Plant"}
                                                    name={"Plant"}
                                                    id="ddlPlant"
                                                    placeholderText={"Select Plant"}
                                                    className={""}
                                                    selectedValue={null}
                                                    optionLabel={""}
                                                    optionValue={""}
                                                    OptionsList={this.state.PlantsOpt}
                                                    // OnChange={(selectedOption: any, actionMeta: any) => { this.handleChange(selectedOption, actionMeta, "divPlant" ) }}
                                                    OnChange={(selectedOption: any, actionMeta: any) => { }}
                                                    isRequired={true}
                                                    disabled={false}
                                                    refElement={null}
                                                    noOptionsMessage="No Plant"
                                                />
                                            </div>
                                        </div>
                                        <div className="col-md-3 bs-field" title={'Department'}>
                                            <div className="custom-dropdown" id="divDepartment">
                                                <SearchableDropdown
                                                    label={"Department"}
                                                    Title={"Department"}
                                                    name={"Department"}
                                                    id="ddlDepartment"
                                                    placeholderText={"Department"}
                                                    className={""}
                                                    selectedValue={null}
                                                    optionLabel={""}
                                                    optionValue={""}
                                                    OptionsList={[]}
                                                    // OnChange={(selectedOption: any, actionMeta: any) => { this.handleChange(selectedOption, actionMeta, "divPlant" ) }}
                                                    OnChange={(selectedOption: any, actionMeta: any) => { }}
                                                    isRequired={true}
                                                    disabled={false}
                                                    refElement={null}
                                                    noOptionsMessage="No Department"
                                                />
                                            </div>
                                        </div>
                                        <div className="col-md-3 bs-field" title={'Zone'}>
                                            <div className="custom-dropdown" id="divZone">
                                                <SearchableDropdown
                                                    label={"Zone"}
                                                    Title={"Zone"}
                                                    name={"Zone"}
                                                    id="ddlZone"
                                                    placeholderText={"Zone"}
                                                    className={""}
                                                    selectedValue={null}
                                                    optionLabel={""}
                                                    optionValue={""}
                                                    OptionsList={[]}
                                                    // OnChange={(selectedOption: any, actionMeta: any) => { this.handleChange(selectedOption, actionMeta, "divPlant" ) }}
                                                    OnChange={(selectedOption: any, actionMeta: any) => { }}
                                                    isRequired={true}
                                                    disabled={false}
                                                    refElement={null}
                                                    noOptionsMessage="No Zone"
                                                />
                                            </div>
                                        </div>
                                        <div className="col-md-3 bs-field" title={'Work Cell'}>
                                            <div className="custom-dropdown" id="divWorkCell">
                                                <SearchableDropdown
                                                    label={"Work Cell"}
                                                    Title={"Work Cell"}
                                                    name={"WorkCell"}
                                                    id="ddlWorkCell"
                                                    placeholderText={"WorCell"}
                                                    className={""}
                                                    selectedValue={null}
                                                    optionLabel={""}
                                                    optionValue={""}
                                                    OptionsList={[]}
                                                    // OnChange={(selectedOption: any, actionMeta: any) => { this.handleChange(selectedOption, actionMeta, "divPlant" ) }}
                                                    OnChange={(selectedOption: any, actionMeta: any) => { }}
                                                    isRequired={true}
                                                    disabled={false}
                                                    refElement={null}
                                                    noOptionsMessage="No Work Cell"
                                                />
                                            </div>
                                        </div>
                                        <div className="col-md-3 bs-field" title={'Machine'}>
                                            <div className="custom-dropdown" id="divMachine">
                                                <SearchableDropdown
                                                    label={"Machine"}
                                                    Title={"Machine"}
                                                    name={"Machine"}
                                                    id="ddlMachine"
                                                    placeholderText={"Machine"}
                                                    className={""}
                                                    selectedValue={null}
                                                    optionLabel={""}
                                                    optionValue={""}
                                                    OptionsList={[]}
                                                    // OnChange={(selectedOption: any, actionMeta: any) => { this.handleChange(selectedOption, actionMeta, "divPlant" ) }}
                                                    OnChange={(selectedOption: any, actionMeta: any) => { }}
                                                    isRequired={true}
                                                    disabled={false}
                                                    refElement={null}
                                                    noOptionsMessage="No Machine"
                                                />
                                            </div>
                                        </div>
                                        <div className="col-md-3 bs-field" title={'Shift'}>
                                            <div className="custom-dropdown" id="divShift">
                                                <SearchableDropdown
                                                    label={"Shift"}
                                                    Title={"Shift"}
                                                    name={"Shift"}
                                                    id="ddlShift"
                                                    placeholderText={"Shift"}
                                                    className={""}
                                                    selectedValue={null}
                                                    optionLabel={""}
                                                    optionValue={""}
                                                    OptionsList={[]}
                                                    // OnChange={(selectedOption: any, actionMeta: any) => { this.handleChange(selectedOption, actionMeta, "divPlant" ) }}
                                                    OnChange={(selectedOption: any, actionMeta: any) => { }}
                                                    isRequired={true}
                                                    disabled={false}
                                                    refElement={null}
                                                    noOptionsMessage="No Shift"
                                                />
                                            </div>
                                        </div>
                                        <div className="col-md-3 bs-field" title={'Supervisor'}>
                                            <div className="custom-dropdown" id="divSupervisor">
                                                <SearchableDropdown
                                                    label={"Supervisor"}
                                                    Title={"Supervisor"}
                                                    name={"Supervisor"}
                                                    id="ddlSupervisor"
                                                    placeholderText={"Supervisor"}
                                                    className={""}
                                                    selectedValue={null}
                                                    optionLabel={""}
                                                    optionValue={""}
                                                    OptionsList={[]}
                                                    // OnChange={(selectedOption: any, actionMeta: any) => { this.handleChange(selectedOption, actionMeta, "divPlant" ) }}
                                                    OnChange={(selectedOption: any, actionMeta: any) => { }}
                                                    isRequired={true}
                                                    disabled={false}
                                                    refElement={null}
                                                    noOptionsMessage="No Supervisor"
                                                />
                                            </div>
                                        </div>
                                        <div className="col-md-3 bs-field" title={'Tool Number'}>
                                            <div className="custom-dropdown" id="divToolNumber">
                                                <SearchableDropdown
                                                    label={"Tool Number"}
                                                    Title={"Tool Number"}
                                                    name={"ToolNumber"}
                                                    id="ddlToolNumber"
                                                    placeholderText={"Tool Number"}
                                                    className={""}
                                                    selectedValue={null}
                                                    optionLabel={""}
                                                    optionValue={""}
                                                    OptionsList={[]}
                                                    // OnChange={(selectedOption: any, actionMeta: any) => { this.handleChange(selectedOption, actionMeta, "divPlant" ) }}
                                                    OnChange={(selectedOption: any, actionMeta: any) => { }}
                                                    isRequired={true}
                                                    disabled={false}
                                                    refElement={null}
                                                    noOptionsMessage="No Tool Number"
                                                />
                                            </div>
                                        </div>
                                    </div>
                                </div>

                                {/* Job Steps Table */}
                                <div style={{ backgroundColor: "#daeff7" }}>
                                    <table style={{ width: "100%" }} id="tableJSRALine">
                                        <thead>
                                            <tr className="theading darkgreybg border">
                                                <td colSpan={2} style={{ width: "22%" }}>Job Steps</td>
                                                <td style={{ width: "8%" }}>Required</td>
                                                <td style={{ width: "22%" }}>Risk Family</td>
                                                <td style={{ width: "22%" }}>Risk</td>
                                                <td colSpan={2} style={{ width: "30%" }}>Factors</td>
                                                <td style={{ width: "3%" }}><span style={{ display: "block", textAlign: "center", lineHeight: "15px" }}>Risk Level</span></td>
                                                <td style={{ width: "1%" }}><span style={{ display: "block", textAlign: "center", lineHeight: "15px" }}>Total Score</span></td>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            {/* Dynamic rows go here */}
                                        </tbody>
                                    </table>

                                    <input type="button" value="Add Job Step" className="addbutton" id="btnAddJobStep" />
                                    <b><label id="lblJosbStepMessage" style={{ color: "red" }}></label></b>
                                </div>

                                {/* Permits Section */}
                                <div className="darkgreybg"><h2>Permits</h2></div>

                                <table className="bluebg">
                                    <thead>
                                        <th>Type</th>
                                        <th>Required</th>
                                        <th>Acquired</th>
                                    </thead>
                                    <tbody>
                                        {[
                                    { type: "Hot Work", id: "HotWorkPermit" },
                                    { type: "Electrical", id: "ElectricalPermit" },
                                    { type: "Confined Space", id: "ConfinedSpacePermit" }
                                ].map(p =>(
                                        <tr><td>{`${p.type}`}</td>
                                        <td><input id={`${p.id}Required`} type="checkbox" className="checkinputfeild" /></td>
                                        <td><input id={`${p.id}Acquired`} type="checkbox" className="checkinputfeild" /></td>
                                        </tr>))}
                                        </tbody>
                                </table>

                                {/* PPE Requirements */}
                                <div className="darkgreybg"><h2>PPE Requirements</h2></div>

                                <table style={{ width: "100%", padding: "0 4px" }} id="tablePPE">
                                    <thead>
                                        <tr className="bluebg">
                                            <td style={{ padding: "4px 0 3px 5px" }}>PPE Type</td>
                                            <td></td>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        <tr>
                                            <td style={{ width: "20%" }}>
                                                  <SearchableDropdown
                                                    label={""}
                                                    Title={"PPEtype_0"}
                                                    name={"PPEtype_0"}
                                                    id="ddlPPEtype_0"
                                                    placeholderText={"Select PPE Type"}
                                                    className={""}
                                                    selectedValue={null}
                                                    optionLabel={""}
                                                    optionValue={""}
                                                    OptionsList={[]}
                                                    // OnChange={(selectedOption: any, actionMeta: any) => { this.handleChange(selectedOption, actionMeta, "divPlant" ) }}
                                                    OnChange={(selectedOption: any, actionMeta: any) => { }}
                                                    isRequired={true}
                                                    disabled={false}
                                                    refElement={null}
                                                    noOptionsMessage="No PPE Type"
                                                />
                                            </td>
                                            <td style={{ width: "80%" }}></td>
                                        </tr>
                                    </tbody>
                                </table>
                                <input type="button" value="Add PPE" className="addbutton" id="btnAddPPE" />
                                <b><label id="ppeStatus" style={{ color: "red" }}></label></b>

                                {/* Persons Involved */}
                                <div className="darkgreybg"><h2>Persons Involved</h2></div>

                                <table style={{ width: "100%", padding: "0 4px" }} id="tablePersons">
                                    <thead>
                                        <tr className="bluebg">
                                            <td style={{ width: "20%", border: "1px solid #99daf3", padding: "4px" }}>Name</td>
                                            <td style={{ width: "20%", border: "1px solid #99daf3", padding: "4px" }}>Date</td>
                                            <td colSpan={3}></td>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        <tr className="jsrapersonnel" data-text="0">
                                            <td style={{ border: "1px solid #99daf3" }}>
                                                    <input className="" placeholder="Name" name="PersonName_0" type="text" id="txtPersonName_0" ref={null} value={''}  disabled={false} title="Name" />
                                            </td>
                                            <td style={{ border: "1px solid #99daf3" }}>
                                                <DatePickercontrol placeholder="MM/DD/YYYY" selectedDate={''} id='PersonDate_0' isDisabled={false} startDate={undefined} endDate={undefined} name="Date" onDatechange={(dateProps: any) => null} highlightDate={new Date()} showIcon />

                                            </td>
                                            <td colSpan={3}></td>
                                        </tr>
                                    </tbody>
                                </table>

                                <input type="button" value="Add Person" className="addbutton" id="btnAddPerson" />
                                <b><label id="personnelStatus" style={{ color: "red" }}></label></b>

                                {/* Supervisor Info */}
                                <table style={{ width: "100%", padding: "0 4px" }}>
                                    <thead>
                                        <tr className="bluebg">
                                            <td className="t-l-padding border" style={{ width: "20%", fontSize: "14px", fontWeight: "bold", padding: "4px" }}>
                                                Supervisor Name
                                            </td>
                                            <td className="t-l-padding border" style={{ width: "20%", fontSize: "14px", fontWeight: "bold", padding: "4px" }}>
                                                Date
                                            </td>
                                            <td colSpan={3}></td>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        <tr>
                                            <td style={{ border: "1px solid #99daf3" }}>
                                                    <input className="" placeholder="Supervisor Name" name="SupervisorName" type="text" id="txtSupervisorName" ref={null} value={''}  disabled={false} title="Supervisor Name" />
                                            </td>
                                            <td style={{ border: "1px solid #99daf3" }}>
                                                <DatePickercontrol placeholder="MM/DD/YYYY" selectedDate={''} id='SupervisorDate' isDisabled={false} startDate={undefined} endDate={undefined} name="SupervisorDate" onDatechange={(dateProps: any) => null} highlightDate={new Date()} showIcon />

                                            </td>
                                            <td colSpan={3}></td>
                                        </tr>
                                    </tbody>
                                </table>

                                {/* Buttons */}
                                <div className="row my-3 px-3" id="divButtons" >
                                    <div className="col-sm-10 text-center">
                                    <button type="button" id="btnSubmit" className="btn SubmitButtons" title="Submit">Submit</button>
                                    <button type="button" id="btnCancel" className="btn btn-secondary"  title="Cancel">Cancel</button>
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