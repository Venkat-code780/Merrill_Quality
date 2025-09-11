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
    isSupplierTeam: boolean;
    isDTETeam: boolean;
    isProcurementTeam: boolean;
}

export interface JSRAFormState {
}

export default class JSRAForm extends React.Component<JSRAFormProps, JSRAFormState> {

    private rootSiteURL: string;private currentSiteURL: string; private MaycoURL: string;
    // private sp = spfi().using(SPFx(this.props.context)); 
    public state = {
        //Cascaded dropdowns
        Plants: [], Departments: [], Zones: [], WorkCells: [], Machines: [], Shifts: [], Supervisors: [], ToolNumbers: [],
        PlantsOpt: [], DepartmentsOpt: [], ZonesOpt: [], WorkCellsOpt: [], MachinesOpt: [], ShiftsOpt: [], SupervisorsOpt: [], ToolNumbersOpt: [],
        formData: {
            Date: null,
            Plant: '',
            Department: '',
            Zone: '',
            WorkCell: '',
            Machine: '',
            Shift: '',
            Supervisor: '',
            ToolNumber: '',
        },
       JSRASubCategories: [],Probability: [],Controls: [],Severity: [],PPETypes: [],
       JSRASubCategoriesOpt: [],ProbabilityOpt: [],ControlsOpt: [],SeverityOpt: [],PPETypesOpt: [],
       RiskFamily:{},Risk:{},
        jobSteps: [
            {
                id: 1,
                step: '',
                required: false,
                riskFamily: '',
                risk: '',
                probability: '',
                controls: '',
                severity: '',
                riskLevel: { probability: 0, controls: 0, severity: 0 },
                totalScore: 0,
            }
        ],
        Homeredirect: false,
        isToolNumberMandatory: false,
    }

    constructor(props: JSRAFormProps) {
        super(props);
        this.rootSiteURL = this.props.spContext.siteAbsoluteUrl;
        this.currentSiteURL = this.props.spContext.webAbsoluteUrl;
        this.MaycoURL = `${this.rootSiteURL}/mayco`;
        //console.log(this.rootSiteURL);      //sites/wcm
    }

    public componentDidMount(): void {
        highlightCurrentNav("liJSRAForm");
        document.title = "Mayco - Safety | JSRA";
        this.getOnLoadData();
    }

    private getOnLoadData = async () => {
        let { getListItems } = initCommonFunctions(this.props.context, this.rootSiteURL);
        let PlantList = 'Plant';
        let DepartmentList = 'Department', DepartmentSelQuery = 'Title,Plant/Title,Plant/Id,*', DepartmentExpFields = 'Plant';
        let ZoneList = 'Zones', ZoneSelQuery = 'Title,Plant/Title,Plant/Id,Department/Title,Department/Id,*', ZoneExpFields = 'Plant,Department';
        let WorkCellList = 'WorkCells';
        let MachineList = 'Machines', MachineSelQuery = 'Title,Plant/Title,Plant/Id,Department/Title,Department/Id,Zone/Title,Zone/Id,*', MachineExpFields = 'Plant,Department,Zone';
        let ShiftList = 'Shifts';
        let SupervisorList = 'Supervisor', SupervisorSelQuery = 'Title,Plant/Title,Plant/Id,*', SupervisorExpFields = 'Plant';
        let ToolNumberList = 'Tool Numbers', ToolNumberSelQuery = 'Title,Plant/Title,Plant/Id,*', ToolNumberExpFields = 'Plant';
        let JSRASubCategoriesList = 'JSRASubCategories', JSRASubCategoriesSelQuery = 'Title,Category/Title,Category/Id,*', JSRASubCategoriesExpFields = 'Category';
        let ProbabilityList = 'Probability';
        let ControlsList = 'Controls';
        let SeverityList = 'Severity';
        let PPETypesList = 'PPETypes';
        let filterIsActive='Is_x0020_Active eq 1';

        try {
            showLoader();
            let [Plants, Departments, Zones, WorkCells, Machines, Shifts, Supervisors, ToolNumbers,JSRASubCategories,Probability,Controls,Severity,PPETypes] = await Promise.all([
                getListItems(PlantList,this.MaycoURL),
                getListItems(DepartmentList,this.MaycoURL, DepartmentSelQuery,DepartmentExpFields),
                getListItems(ZoneList, this.MaycoURL, ZoneSelQuery, ZoneExpFields),
                getListItems(WorkCellList, this.MaycoURL),
                getListItems(MachineList, this.MaycoURL, MachineSelQuery, MachineExpFields),
                getListItems(ShiftList, this.MaycoURL),
                getListItems(SupervisorList, this.MaycoURL, SupervisorSelQuery, SupervisorExpFields),
                getListItems(ToolNumberList, this.MaycoURL, ToolNumberSelQuery, ToolNumberExpFields),
                getListItems(JSRASubCategoriesList, this.currentSiteURL, JSRASubCategoriesSelQuery, JSRASubCategoriesExpFields),
                getListItems(ProbabilityList, this.currentSiteURL,'','',filterIsActive),
                getListItems(ControlsList, this.currentSiteURL,'','',filterIsActive),
                getListItems(SeverityList, this.currentSiteURL,'','',filterIsActive),
                getListItems(PPETypesList, this.currentSiteURL),
            ])
            let PlantsOpt = this.getMapedOptions(Plants,'Title','Title');
            let ShiftsOpt = this.getMapedOptions(Shifts,'Title','Title');
            // RiskFamily and Risk depends on JSRASubCategories
            
            let JSRASubCategoriesOpt = this.getMapedOptions(JSRASubCategories,'Title','Title');
            let ProbabilityOpt = this.getMapedOptions(Probability,'Title','Title');
            let ControlsOpt = this.getMapedOptions(Controls,'Title','Title');
            let SeverityOpt = this.getMapedOptions(Severity,'Title','Title');
            let PPETypesOpt = this.getMapedOptions(PPETypes,'Title','Title');
            // for sorting
            PlantsOpt.sort((a, b) => a.label.localeCompare(b.label));
            ShiftsOpt.sort((a, b) => a.label.localeCompare(b.label));
            JSRASubCategoriesOpt.sort((a, b) => a.label.localeCompare(b.label));
            ProbabilityOpt.sort((a, b) => a.label.localeCompare(b.label));
            ControlsOpt.sort((a, b) => a.label.localeCompare(b.label));
            SeverityOpt.sort((a, b) => a.label.localeCompare(b.label));
            PPETypesOpt.sort((a, b) => a.label.localeCompare(b.label));
            this.setState({ Plants, PlantsOpt, ShiftsOpt,JSRASubCategoriesOpt,ProbabilityOpt,ControlsOpt,SeverityOpt,PPETypesOpt, Departments, Zones, WorkCells, Machines, Supervisors, ToolNumbers });
            hideLoader();
        } catch (e) {
            console.log(e);
            this.onError();
        }
    }
    //common functions
    private getMapedOptions=(ListItems: any[],labelText: string,valueText: string)=>
    {

       return ListItems.map((item: any) => ({
                label: item[labelText],
                value: item[valueText],
            }));
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
            let filteredDepts = this.state.Departments.filter((dept: any) => dept.Plant.Title == value && dept.IsActiveMRO);// Note:in ventureglobal IsActive column exists , in WCM  IsActiveMRO column exists
            let DepartmentsOpt = filteredDepts.map((item: any) => ({
                label: item.Title,
                value: item.Title,
            }));
            let filteredSupervisors = this.state.Supervisors.filter((sup: any) => sup.Plant?.Title == value);
            let SupervisorsOpt = filteredSupervisors.map((item: any) => ({
                label: item.Title,
                value: item.Title,
            }));
            let filteredToolNumbers = this.state.ToolNumbers.filter((TN: any) => TN.Plant?.Title == value);
            let ToolNumbersOpt = filteredToolNumbers.map((item: any) => ({
                label: item.Title,
                value: item.Title,
            }));
            // update dependents
            DepartmentsOpt.sort((a, b) => a.label.localeCompare(b.label));
            SupervisorsOpt.sort((a, b) => a.label.localeCompare(b.label));
            ToolNumbersOpt.sort((a, b) => a.label.localeCompare(b.label));
            this.setState({ formData, DepartmentsOpt, SupervisorsOpt, ToolNumbersOpt, ZonesOpt: [], WorkCellsOpt: [], MachinesOpt: [], Department: '', Zone: '', WorkCell: '', Machine: '' });
        }
        else if (name == 'Department') {
            let filteredZones = this.state.Zones.filter((dept: any) => dept.Plant.Title == formData.Plant && dept.Department.Title == value);
            let ZonesOpt = filteredZones.map((item: any) => ({
                label: item.Title,
                value: item.Title,
            }));
            //dynamic ToolNumber Mandatory
            let isToolNumberMandatory = false;
            if (value.toLowerCase() == 'molding') {
                isToolNumberMandatory = true;
            }
            // update dependents
            ZonesOpt.sort((a, b) => a.label.localeCompare(b.label));
            this.setState({ formData, ZonesOpt, WorkCellsOpt: [], MachinesOpt: [], Zone: '', WorkCell: '', Machine: '', isToolNumberMandatory });
        }
        else if (name == 'Zone') {
            let filteredMachines = this.state.Machines.filter((dept: any) => dept.Plant.Title == formData.Plant && dept.Department.Title == formData.Department && dept.Zone.Title == value);
            let MachinesOpt = filteredMachines.map((item: any) => ({
                label: item.Title,
                value: item.Title,
            }));
            let filteredWorkCells = this.state.WorkCells.filter((dept: any) => dept.h7kc == formData.Plant && dept.Department == formData.Department && dept.Zone == value);
            let WorkCellsOpt = filteredWorkCells.map((item: any) => ({
                label: item.Title,
                value: item.Title,
            }));
            // update dependents
            MachinesOpt.sort((a, b) => a.label.localeCompare(b.label));
            WorkCellsOpt.sort((a, b) => a.label.localeCompare(b.label));
            this.setState({ formData, MachinesOpt, WorkCellsOpt, WorkCell: '', Machine: '' });
        }
        this.setState({ formData });
    }
    // Job Steps Section

    private bindJobSteps = () => {
        let jobSteps = this.state.jobSteps;
        let DynamicHTML: JSX.Element[] = [];

        DynamicHTML = jobSteps.map((jobStep, index) => (
            <tr key={jobStep.id}>
                <td>{jobStep.id}</td>
                <td>
                 <input className="form-control" placeholder={`Job Step_${jobStep.id}`} name={`JobStep_${jobStep.id}`} type="text" id={`txtJobStep_${jobStep.id}`} value={jobStep.step} onChange={(e) => this.handleJobStepChange(index, 'step', e.target.value)} disabled={false} title={jobStep.step} />

                </td>
                <td>
                    <input
                        type="checkbox"
                        checked={jobStep.required}
                        onChange={(e) => this.handleJobStepChange(index, 'required', e.target.checked)} />
                </td>
                <td>
                    <div className="custom-dropdown" id={`divriskFamily_${jobStep.id}`}>
                                                <SearchableDropdown
                                                    label={""}
                                                    Title={""}
                                                    name={""}
                                                    id="ddlPlant"
                                                    placeholderText={""}
                                                    className={""}
                                                    selectedValue={jobStep.riskFamily}
                                                    OptionsList={this.state.JSRASubCategoriesOpt}
                                                    // OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta) }}
                                                    OnChange={(selectedOption: any, actionMeta: any) => this.handleJobStepChange(index, 'riskFamily',selectedOption.value)}
                                                    isRequired={false}
                                                    disabled={false}
                                                    refElement={null}
                                                    noOptionsMessage="No Risk Family"
                                                />
                    </div>
                </td>
                <td>
                    <select
                        value={jobStep.risk}
                        onChange={(e) => this.handleJobStepChange(index, 'risk', e.target.value)}>
                        <option value="None">None</option>
                        <option value="Low">Low</option>
                        <option value="Medium">Medium</option>
                        <option value="High">High</option>
                    </select>
                </td>
                <td>
                    <div><span>probability</span>
                        <select
                            value={jobStep.probability}
                            onChange={(e) => this.handleJobStepChange(index, 'probability', e.target.value)}
                            onBlur={() => this.calculateRiskLevelAndScore(index)} // Recalculate on blur
                        >
                            <option value="None">None</option>
                            <option value="Low">Low</option>
                            <option value="High">High</option>
                        </select>
                    </div>
                    <div><span>controls</span>
                        <select
                            value={jobStep.controls}
                            onChange={(e) => this.handleJobStepChange(index, 'controls', e.target.value)}>
                            <option value="None">None</option>
                            <option value="Basic">Basic</option>
                            <option value="Advanced">Advanced</option>
                        </select>
                    </div>
                    <div><span>severity</span>
                        <select
                            value={jobStep.severity}
                            onChange={(e) => this.handleJobStepChange(index, 'severity', e.target.value)}
                            onBlur={() => this.calculateRiskLevelAndScore(index)} // Recalculate on blur
                        >
                            <option value="None">None</option>
                            <option value="Low">Low</option>
                            <option value="Medium">Medium</option>
                            <option value="High">High</option>
                        </select>
                    </div>
                </td>

                <td>
                    <span>{jobStep.riskLevel.probability}</span>
                    <span>{jobStep.riskLevel.controls}</span>
                    <span>{jobStep.riskLevel.severity}</span>
                </td>
                <td>{jobStep.totalScore}</td>
            </tr>
        ))

        return DynamicHTML;
    };

    private handleJobStepChange = (index: any, field: any, value: any) => {
        const updatedJobSteps = [...this.state.jobSteps];
        updatedJobSteps[index] = {
            ...updatedJobSteps[index],
            [field]: value,
        };
        this.setState({ jobSteps: updatedJobSteps });
    };
    private addJobStep = () => {
        let prevjobSteps = this.state.jobSteps;
        let updatedjobSteps = [...prevjobSteps,
        {
            id: prevjobSteps.length + 2,
            step: '',
            required: false,
            riskFamily: '',
            risk: '',
            probability: '',
            controls: '',
            severity: '',
            riskLevel: { probability: 0, controls: 0, severity: 0 },
            totalScore: 0,
        },]
        this.setState({ jobSteps: updatedjobSteps });
    };
    private calculateRiskLevelAndScore = (index: any) => {
        const { jobSteps } = this.state;
        const jobStep = jobSteps[index];

        // Example: Risk level is based on probability , controls and severity
        const riskLevel = jobStep.riskLevel;
        const totalScore = 2;

        const updatedJobSteps = [...jobSteps];
        updatedJobSteps[index] = {
            ...updatedJobSteps[index],
            riskLevel,
            totalScore,
        };

        this.setState({ jobSteps: updatedJobSteps });
    };
    //CRUD functions,
    //Dynamic HTML Binding functions
    //Date functions

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
                                                    placeholderText={""}
                                                    className={""}
                                                    selectedValue={this.state.formData.Plant}
                                                    OptionsList={this.state.PlantsOpt}
                                                    OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta) }}
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
                                                    placeholderText={""}
                                                    className={""}
                                                    selectedValue={this.state.formData.Department}
                                                    OptionsList={this.state.DepartmentsOpt}
                                                    OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta) }}
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
                                                    placeholderText={""}
                                                    className={""}
                                                    selectedValue={this.state.formData.Zone}
                                                    OptionsList={this.state.ZonesOpt}
                                                    OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta) }}
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
                                                    placeholderText={""}
                                                    className={""}
                                                    selectedValue={this.state.formData.WorkCell}
                                                    OptionsList={this.state.WorkCellsOpt}
                                                    OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta) }}
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
                                                    placeholderText={""}
                                                    className={""}
                                                    selectedValue={this.state.formData.Machine}
                                                    OptionsList={this.state.MachinesOpt}
                                                    OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta) }}
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
                                                    placeholderText={""}
                                                    className={""}
                                                    selectedValue={this.state.formData.Shift}
                                                    OptionsList={this.state.ShiftsOpt}
                                                    OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta) }}
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
                                                    placeholderText={""}
                                                    className={""}
                                                    selectedValue={this.state.formData.Supervisor}
                                                    OptionsList={this.state.SupervisorsOpt}
                                                    OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta) }}
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
                                                    placeholderText={""}
                                                    className={""}
                                                    selectedValue={this.state.formData.ToolNumber}
                                                    OptionsList={this.state.ToolNumbersOpt}
                                                    OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta) }}
                                                    isRequired={this.state.isToolNumberMandatory}
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
                                    <table key="jobStepsTable">
                                        <thead>
                                            <tr>
                                                <th></th>
                                                <th>Job Step</th>
                                                <th>Required</th>
                                                <th>Risk Family</th>
                                                <th>Risk</th>
                                                <th>Factors</th>
                                                <th>Risk Level</th>
                                                <th>Total Score</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            {this.bindJobSteps()}
                                        </tbody>
                                    </table>
                                    <input type="button" value="Add Job Step" className="addbutton" onClick={this.addJobStep} id="btnAddJobStep" />
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
                                        ].map(p => (
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
                                                    placeholderText={""}
                                                    className={""}
                                                    selectedValue={null}
                                                    OptionsList={[]}
                                                    OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta) }}
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
                                                <input className="" placeholder="Name" name="PersonName_0" type="text" id="txtPersonName_0" ref={null} value={''} disabled={false} title="Name" />
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
                                                <input className="" placeholder="Supervisor Name" name="SupervisorName" type="text" id="txtSupervisorName" ref={null} value={''} disabled={false} title="Supervisor Name" />
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
                                        <button type="button" id="btnCancel" className="btn btn-secondary" title="Cancel">Cancel</button>
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