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
import DateUtilities from "../Utilities/DateUtilities";
import { format } from "date-fns";
// import FileUpload from "../Shared/FileUpload";
// import InputCheckBox from "../Shared/InputCheckBox";
// import { format } from "date-fns";
import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";
import { faTrash } from "@fortawesome/free-solid-svg-icons";

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

    private rootSiteURL: string; private currentSiteURL: string; private MaycoURL: string;
    private Date: any; private SupervisorDate: any; private SupervisorName: any;
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
            SupervisorName: '',
            SupervisorDate: null
        },
        Probability: [], Controls: [], Severity: [],
        ProbabilityOpt: [], ControlsOpt: [], SeverityOpt: [], PPETypesOpt: [],
        GroupedRiskFamilyandRiskObj: {}, RiskFamilyOpt: [],
        jobSteps: [
            {
                id: 1,
                Step: '',
                required: true,
                RiskFamily: '',
                Risk: '',
                Probability: '0',
                Controls: '0',
                Severity: '0',
                RiskLevel: { Probability: '0', Controls: '0', Severity: '0' },
                RiskLevelColorClasses: { Probability: 'RiskLevel-0', Controls: 'RiskLevel-0', Severity: 'RiskLevel-0', Total: 'RiskLevel-0' },
                TotalScore: '0',
                RiskOpt: [],
            }
        ],
        Permits:[
                    { type: "Hot Work", id: "0", Required: false, Acquired: false},
                    { type: "Electrical", id: "1", Required: false, Acquired: false },
                    { type: "Confined Space", id: "2", Required: false, Acquired: false }
                ],
        PPETypes: [
            {
                id: 1,
                PPEType: '',
            }
        ],
        Persons: [
            {
                id: 1,
                PersonName: '',
                PersonDate: null
            }
        ],
        isEditForm: false,
        ItemId: 0,
        Homeredirect: false,
        isToolNumberMandatory: false,
    }

    constructor(props: JSRAFormProps) {
        super(props);
        this.rootSiteURL = this.props.spContext.siteAbsoluteUrl; //   sites/wcm
        this.currentSiteURL = this.props.spContext.webAbsoluteUrl;//  sites/wcm/mayco/merrill/sa
        this.MaycoURL = `${this.rootSiteURL}/mayco`;
        this.Date = React.createRef();
        this.SupervisorDate = React.createRef();
        this.SupervisorName = React.createRef();
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
        let filterIsActive = 'Is_x0020_Active eq 1';

        try {
            showLoader();
            let [Plants, Departments, Zones, WorkCells, Machines, Shifts, Supervisors, ToolNumbers, JSRASubCategories, Probability, Controls, Severity, PPETypes] = await Promise.all([
                getListItems(PlantList, this.MaycoURL),
                getListItems(DepartmentList, this.MaycoURL, DepartmentSelQuery, DepartmentExpFields),
                getListItems(ZoneList, this.MaycoURL, ZoneSelQuery, ZoneExpFields),
                getListItems(WorkCellList, this.MaycoURL),
                getListItems(MachineList, this.MaycoURL, MachineSelQuery, MachineExpFields),
                getListItems(ShiftList, this.MaycoURL),
                getListItems(SupervisorList, this.MaycoURL, SupervisorSelQuery, SupervisorExpFields),
                getListItems(ToolNumberList, this.MaycoURL, ToolNumberSelQuery, ToolNumberExpFields),
                getListItems(JSRASubCategoriesList, this.currentSiteURL, JSRASubCategoriesSelQuery, JSRASubCategoriesExpFields),
                getListItems(ProbabilityList, this.currentSiteURL, '', '', filterIsActive),
                getListItems(ControlsList, this.currentSiteURL, '', '', filterIsActive),
                getListItems(SeverityList, this.currentSiteURL, '', '', filterIsActive),
                getListItems(PPETypesList, this.currentSiteURL),
            ])
            let PlantsOpt = this.getMapedOptions(Plants, 'Title', 'Title');
            let ShiftsOpt = this.getMapedOptions(Shifts, 'Title', 'Title');
            // RiskFamily=Category(lookup) and Risk(Title) depends on JSRASubCategories list 
            let GroupedRiskFamilyandRiskObj = JSRASubCategories.reduce((acc, item) => {
                const title = item.Category.Title;
                if (!acc[title]) {
                    acc[title] = [];
                }
                acc[title].push(item);
                return acc;
            }, {} as Record<string, { Title: string; Id: string }[]>);
            let RiskFamilyOpt = this.getMapedOptions(Object.keys(GroupedRiskFamilyandRiskObj));
            let ProbabilityOpt = this.getMapedOptions(Probability, 'Title', 'Value');
            let ControlsOpt = this.getMapedOptions(Controls, 'Title', 'Value');
            let SeverityOpt = this.getMapedOptions(Severity, 'Title', 'Value');
            let PPETypesOpt = this.getMapedOptions(PPETypes, 'Title', 'Title');
            // for sorting
            PlantsOpt.sort((a, b) => a.label.localeCompare(b.label));
            ShiftsOpt.sort((a, b) => a.label.localeCompare(b.label));
            RiskFamilyOpt.sort((a, b) => a.label.localeCompare(b.label));
            ProbabilityOpt.sort((a, b) => (a.value) - (b.value));
            ControlsOpt.sort((a, b) => (a.value) - (b.value));
            SeverityOpt.sort((a, b) => (a.value) - (b.value));
            PPETypesOpt.sort((a, b) => a.label.localeCompare(b.label));
            this.setState({ Plants, PlantsOpt, ShiftsOpt, RiskFamilyOpt, GroupedRiskFamilyandRiskObj, ProbabilityOpt, ControlsOpt, SeverityOpt, PPETypesOpt, Departments, Zones, WorkCells, Machines, Supervisors, ToolNumbers });
            hideLoader();
        } catch (e) {
            console.log(e);
            this.onError();
        }
    }
    //common functions
    private getMapedOptions = (ListItems: any[], labelText: string = '', valueText: string = '') => {
        return ListItems.map((item: any) => ({
            label: typeof (item) == "string" ? item : item[labelText],
            value: typeof (item) == "string" ? item : item[valueText],
        }));
    }

    // on change functions
    private handleChange = (event: any) => {
        const formData: any = { ...this.state.formData };

        const name = event.target.name;
        let inputValue = (event.target.type == "text" || event.target.type == "textarea") ? event.target.value : event.target.checked ? event.target.value : '';
        let classArr: any = event.target.className;

        if (classArr.includes("onlyNum")) {
            inputValue = inputValue.replace(/[^0-9]/g, '');
        }
        if(name.includes('chkPermits'))
        {
          let Permits=[...this.state.Permits];
          let index=name.split('_')[2],field=name.split('_')[1];
          Permits[index]={
            ...Permits[index],
            [field]:inputValue
          }

          this.setState({ Permits });
        }
        else{
            formData[name] = inputValue;
            this.setState({ formData });
        }
    }
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

        dateValue = format(DateUtilities.addBrowserwrtServer(new Date(DateUtilities.getDateMMDDYYYY(dateValue)), this.props.spContext.webTimeZoneData).toISOString(), "MM/dd/yyyy");
        if (divId.includes('divPersonDate')) {
            let updatedPersons = [...this.state.Persons];
            let index = name.split('_')[1];
            updatedPersons[index] = {
                ...updatedPersons[index],
                PersonDate: dateValue,
            };
            this.setState({ Persons: updatedPersons });
        }
        formData[name] = dateValue;

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
                    <input className="form-control" placeholder={``} name={`JobStep_${jobStep.id}`} type="text" id={`JobStep_${jobStep.id}`} value={jobStep.Step} onChange={(e) => this.handleJobStepChange(index, 'Step', e.target.value)} disabled={false} title={jobStep.Step} />

                </td>
                <td>
                    <input
                        type="checkbox"
                        checked={jobStep.required}
                        onChange={(e) => this.handleJobStepChange(index, 'required', e.target.checked)} />
                </td>
                <td>
                    <div className="custom-dropdown" id={`divRiskFamily_${jobStep.id}`}>
                        <SearchableDropdown
                            label={""}
                            Title={""}
                            name={""}
                            id={`RiskFamily_${jobStep.id}`}
                            placeholderText={""}
                            className={""}
                            selectedValue={jobStep.RiskFamily}
                            OptionsList={this.state.RiskFamilyOpt}
                            OnChange={(selectedOption: any, actionMeta: any) => this.handleJobStepChange(index, 'RiskFamily', selectedOption ? selectedOption.value : '')}
                            isRequired={false}
                            disabled={!jobStep.required}
                            noOptionsMessage="No Risk Family"
                        />
                    </div>
                </td>
                <td>
                    <div className="custom-dropdown" id={`divRisk_${jobStep.id}`}>
                        <SearchableDropdown
                            label={""}
                            Title={""}
                            name={""}
                            id={`Risk_${jobStep.id}`}
                            placeholderText={""}
                            className={""}
                            selectedValue={jobStep.Risk}
                            OptionsList={jobStep.RiskOpt}
                            OnChange={(selectedOption: any, actionMeta: any) => this.handleJobStepChange(index, 'Risk', selectedOption ? selectedOption.value : '')}
                            isRequired={false}
                            disabled={!jobStep.required}
                            noOptionsMessage="No Risk"
                        />
                    </div>
                </td>
                <td>
                    <div className="bs-field form-floating" title={'Probability'}>
                        <div className="custom-dropdown" id={`divProbability_${jobStep.id}`}>
                            <SearchableDropdown
                                label={"Probability"}
                                Title={"Probability"}
                                name={"Probability"}
                                id={`Probability_${jobStep.id}`}
                                placeholderText={""}
                                className={""}
                                selectedValue={jobStep.Probability}
                                OptionsList={this.state.ProbabilityOpt}
                                OnChange={(selectedOption: any, actionMeta: any) => this.handleJobStepChange(index, 'Probability', selectedOption ? selectedOption.value : '0')}
                                isRequired={jobStep.required}
                                disabled={!jobStep.required}
                                noOptionsMessage="No Probability"
                            />
                        </div>
                    </div>
                    <div className="bs-field form-floating" title={'Controls'}>
                        <div className="custom-dropdown" id={`divControls_${jobStep.id}`}>
                            <SearchableDropdown
                                label={"Controls"}
                                Title={"Controls"}
                                name={"Controls"}
                                id={`Controls_${jobStep.id}`}
                                placeholderText={""}
                                className={""}
                                selectedValue={jobStep.Controls}
                                OptionsList={this.state.ControlsOpt}
                                OnChange={(selectedOption: any, actionMeta: any) => this.handleJobStepChange(index, 'Controls', selectedOption ? selectedOption.value : '0')}
                                isRequired={jobStep.required}
                                disabled={!jobStep.required}
                                noOptionsMessage="No Controls"
                            />
                        </div>
                    </div>
                    <div className="bs-field form-floating" title={'Severity'}>
                        <div className="custom-dropdown" id={`divSeverity_${jobStep.id}`}>
                            <SearchableDropdown
                                label={"Severity"}
                                Title={"Severity"}
                                name={"Severity"}
                                id={`Severity_${jobStep.id}`}
                                placeholderText={""}
                                className={""}
                                selectedValue={jobStep.Severity}
                                OptionsList={this.state.SeverityOpt}
                                OnChange={(selectedOption: any, actionMeta: any) => this.handleJobStepChange(index, 'Severity', selectedOption ? selectedOption.value : '0')}
                                isRequired={jobStep.required}
                                disabled={!jobStep.required}
                                noOptionsMessage="No Severity"
                            />
                        </div>
                    </div>
                </td>

                <td>
                    <div className={`${jobStep.RiskLevelColorClasses.Probability}`}>{jobStep.RiskLevel.Probability}</div>
                    <div className={`${jobStep.RiskLevelColorClasses.Controls}`}>{jobStep.RiskLevel.Controls}</div>
                    <div className={`${jobStep.RiskLevelColorClasses.Severity}`}>{jobStep.RiskLevel.Severity}</div>
                </td>
                <td>
                    <div className={`${jobStep.RiskLevelColorClasses.Total}`}>{jobStep.TotalScore}</div>
                </td>
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
        if (field == 'RiskFamily') {
            let GroupedRiskFamilyandRiskObj: Record<string, []> = { ...this.state.GroupedRiskFamilyandRiskObj };
            let ItemRiskOpts = value == '' ? [] : GroupedRiskFamilyandRiskObj[value];
            let RiskOpt = this.getMapedOptions(ItemRiskOpts, 'Title', 'Title');
            console.log(RiskOpt);
            updatedJobSteps[index] = {
                ...updatedJobSteps[index],
                RiskOpt: RiskOpt as [],
            };
        }
        else if (['Probability', 'Controls', 'Severity'].includes(field)) // Calculate Total score
        {
            updatedJobSteps[index] =
            {
                ...updatedJobSteps[index],
                RiskLevel: { ...updatedJobSteps[index].RiskLevel, [field]: `${value}` },
                RiskLevelColorClasses: { ...updatedJobSteps[index].RiskLevelColorClasses, [field]: `RiskLevel-${value}` }
            };
            let TotalScore = parseInt(updatedJobSteps[index].Probability) + parseInt(updatedJobSteps[index].Controls) + parseInt(updatedJobSteps[index].Severity);
            let TotalScoreClass = TotalScore == 0 ? 0 : TotalScore <= 4 ? 1 : TotalScore <= 6 ? 2 : TotalScore >= 7 ? 3 : 0;
            updatedJobSteps[index] =
            {
                ...updatedJobSteps[index],
                TotalScore: `${TotalScore}`,
                RiskLevelColorClasses: { ...updatedJobSteps[index].RiskLevelColorClasses, Total: `RiskLevel-${TotalScoreClass}` }
            };


        }
        this.setState({ jobSteps: updatedJobSteps });
    };
    private addJobStep = () => {
        let isValid = this.validateJobSteps();
        if (isValid.status) {
            let prevjobSteps = this.state.jobSteps;
            let updatedjobSteps = [...prevjobSteps,
            {
                id: prevjobSteps.length + 1,
                Step: '',
                required: true,
                RiskFamily: '',
                Risk: '',
                Probability: '0',
                Controls: '0',
                Severity: '0',
                RiskLevel: { Probability: '0', Controls: '0', Severity: '0' },
                RiskLevelColorClasses: { Probability: 'RiskLevel-0', Controls: 'RiskLevel-0', Severity: 'RiskLevel-0', Total: 'RiskLevel-0' },
                TotalScore: '0',
                RiskOpt: [],
            },]
            this.setState({ jobSteps: updatedjobSteps });
        }
        else {
            showToast("error", isValid.message);
        }
    };
    private validateJobSteps = () => {
        const updatedJobSteps = [...this.state.jobSteps];
        let isValid = { message: '', status: true };

        for (let i = 0; i < updatedJobSteps.length; i++) {
            const Step = updatedJobSteps[i];

            // Validate 'Step' field
            if (!Step.Step.trim()) {
                isValid = { message: `'Job Step' cannot be blank.`, status: false };
                document.getElementById(`JobStep_${Step.id}`)?.classList.add('mandatory-FormContent-focus');
                document.getElementById(`JobStep_${Step.id}`)?.focus();
                break;
            }
            // If required is true, validate other fields
            if (Step.required) {
                const requiredFields = ['RiskFamily', 'Risk', 'Probability', 'Controls', 'Severity'];
                const displayFields = ['Risk Family', 'Risk', 'Probability', 'Controls', 'Severity'];
                for (const field of requiredFields) {
                    let key = (field as keyof typeof Step);
                    let Val = Step[key] as string;
                    Val = ['RiskFamily', 'Risk'].includes(field) ? Val.trim() : Val;
                    if (['', '0', 0].includes(Val)) {
                        isValid = { message: `'${displayFields[requiredFields.indexOf(field)]}' cannot be blank.`, status: false };
                        document.getElementById(`${field}_${Step.id}`)?.classList.add('searchMandatory');
                        document.getElementById(`${field}_${Step.id}`)?.getElementsByTagName("input")[0].focus();
                        break;
                    }
                }

                if (!isValid.status) break;
            }
        }

        return isValid;
    };
    // PPE Requirements Section
    private bindPPETypes = () => {
        let PPETypes = this.state.PPETypes;
        let DynamicHTML: JSX.Element[] = [];

        DynamicHTML = PPETypes.map((PPE, index) => (

            <tr key={PPE.id}>
                <td>
                    <div className="custom-dropdown" id={`divPPE_${PPE.id}`}>
                        <SearchableDropdown
                            label={""}
                            Title={""}
                            name={""}
                            id={`PPEType_${PPE.id}`}
                            placeholderText={""}
                            className={""}
                            selectedValue={PPE.PPEType}
                            OptionsList={this.state.PPETypesOpt}
                            OnChange={(selectedOption: any, actionMeta: any) => this.handlePPETypeChange(index, 'PPEType', selectedOption ? selectedOption.value : '')}
                            isRequired={false}
                            disabled={false}
                            noOptionsMessage="No PPE Type"
                        />
                    </div>
                </td>
                <td>{PPETypes.length>1?<button type='button' className="btn text-danger" onClick={() => this.deletePPEType(index)} title="Delete PPE Type"><FontAwesomeIcon icon={faTrash} /></button>:''}</td>
            </tr>

        ))

        return DynamicHTML;
    };
    private handlePPETypeChange = (index: any, field: any, value: any) => {
        const updatedPPETypes = [...this.state.PPETypes];
        updatedPPETypes[index] = {
            ...updatedPPETypes[index],
            [field]: value,
        };
        this.setState({ PPETypes: updatedPPETypes });
    };
    private addPPEType = () => {
        let isValid = this.validatePPEType();
        if (isValid.status) {
            let prevPPETypes = this.state.PPETypes;
            let updatedPPETypes = [...prevPPETypes,
            {
                id: prevPPETypes.length + 1,
                PPEType: '',
            },]
            this.setState({ PPETypes: updatedPPETypes });
        }
        else {
            showToast("error", isValid.message);
        }
    };
    private deletePPEType = (index: number) => {
        const prevPPETypes = [...this.state.PPETypes];
        const updatedPPETypes = prevPPETypes.filter((_, i) => i !== index);
        // Optionally reassign IDs if needed
        const reindexedPPETypes = updatedPPETypes.map((item, idx) => ({
            ...item,
            id: idx + 1
        }));
        this.setState({ PPETypes: reindexedPPETypes });
    };
    private validatePPEType = () => {
        const updatedPPETypes = [...this.state.PPETypes];
        let isValid = { message: '', status: true };

        for (let i = 0; i < updatedPPETypes.length; i++) {
            const PPE = updatedPPETypes[i];

            // Validate 'PPETypes' field
            if (!PPE.PPEType.trim()) {
                isValid = { message: `'PPE Type' cannot be blank.`, status: false };
                document.getElementById(`PPEType_${PPE.id}`)?.classList.add('searchMandatory');
                document.getElementById(`PPEType_${PPE.id}`)?.getElementsByTagName("input")[0].focus();
                break;
            }
            if (!isValid.status) break;
        }

        return isValid;
    };
    // Persons Section
    private bindPersons = () => {
        let Persons = this.state.Persons;
        let DynamicHTML: JSX.Element[] = [];

        DynamicHTML = Persons.map((Person, index) => (
            <tr key={Person.id}>
                <td>
                    <input className="form-control" placeholder={``} name={`PersonName_${Person.id}`} type="text" id={`PersonName_${Person.id}`} value={Person.PersonName} onChange={(e) => this.handlePersonsChange(index, 'PersonName', e.target.value)} disabled={false} title={Person.PersonName} />

                </td>
                <td>
                    <div className="c-date-picker">
                        <DatePickercontrol placeholder="" selectedDate={Person.PersonDate} isDisabled={false} id={`PersonDate_${Person.id}`} startDate={undefined} endDate={undefined} name={`PersonDate_${index}`} onDatechange={(dateProps: any) => this.handleDateChange(dateProps[0], dateProps[2], "divPersonDate")} highlightDate={new Date()} showIcon />
                    </div>
                </td>
                <td>{Persons.length>1?<button type='button' className="btn text-danger" onClick={() => this.deletePerson(index)} title="Delete Person"><FontAwesomeIcon icon={faTrash} /></button>:''}</td>
            </tr>

        ))

        return DynamicHTML;
    };
    private handlePersonsChange = (index: any, field: any, value: any) => {
        const updatedPersons = [...this.state.Persons];
        updatedPersons[index] = {
            ...updatedPersons[index],
            [field]: value,
        };
        this.setState({ Persons: updatedPersons });
    };
    private addPerson = () => {
        let isValid = this.validatePerson();
        if (isValid.status) {
            let prevPersons = this.state.Persons;
            let updatedPersons = [...prevPersons,
            {
                id: prevPersons.length + 1,
                PersonName: '',
                PersonDate: null
            },]
            this.setState({ Persons: updatedPersons });
        }
        else {
            showToast("error", isValid.message);
        }
    };
    private deletePerson = (index: number) => {
        const prevPersons = [...this.state.Persons];
        const updatedPersons = prevPersons.filter((_, i) => i !== index);
        // Optionally reassign IDs if needed
        const reindexedPersons = updatedPersons.map((person, idx) => ({
            ...person,
            id: idx + 1
        }));
        this.setState({ Persons: reindexedPersons });
    };
    private validatePerson = () => {
        const updatedPersons = [...this.state.Persons];
        let isValid = { message: '', status: true };
        for (let i = 0; i < updatedPersons.length; i++) {
            const Person = updatedPersons[i];

            // Validate 'Step' field
            if (!Person.PersonName.trim()) {
                isValid = { message: `'Person Name' cannot be blank.`, status: false };
                document.getElementById(`PersonName_${Person.id}`)?.classList.add('mandatory-FormContent-focus');
                document.getElementById(`PersonName_${Person.id}`)?.focus();
                break;
            }
            else if (!Person.PersonDate) {
                isValid = { message: `'Person Date' cannot be blank.`, status: false };
                setTimeout(() =>{document.getElementById(`PersonDate_${Person.id}`)?.classList.add('mandatory-FormContent-focus');},100);
                document.getElementById(`PersonDate_${Person.id}`)?.focus();
                break;
            }

            if (!isValid.status) break;
        }

        return isValid;
    };

    private onError = () => {
        showToast("error", ActionStatus.Error);
        hideLoader();
    }
    private handlCancel = () => {
        this.setState({ Homeredirect: true, ItemId: 0 });
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
                                <h4 className="mb-0 pt-2 text-center">{" JSRA " + (this.state.isEditForm ? (" - " + this.state.ItemId) : "")} </h4>
                                <label className="text-end px-1"> <span className="text-danger">* </span> are mandatory fields</label>
                            </div>

                            <div className="mainContent row px-4 borderLine">
                                    <div className="row py-3">
                                    <div className="col-md-3 c-date-picker form-floating" id="divDate">
                                        <label className="label-datePicker" htmlFor="dtDate"> Date <span className="text-danger">*</span></label>
                                        <DatePickercontrol placeholder="" selectedDate={this.state.formData.Date} id='dtSupervisorDate' isDisabled={false} startDate={undefined} endDate={undefined} name="Date" onDatechange={(dateProps: any) => this.handleDateChange(dateProps[0], dateProps[2], "divDate")} ref={this.Date} highlightDate={new Date()} showIcon />
                                    </div>
                                    <div className="col-md-3  form-floating" title={this.state.formData.Plant}>
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
                                    <div className="col-md-3 form-floating" title={this.state.formData.Department}>
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
                                    <div className="col-md-3 form-floating" title={this.state.formData.Zone}>
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
                                    </div>
                                    <div className="row py-3">
                                    <div className="col-md-3 form-floating" title={this.state.formData.WorkCell}>
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
                                    <div className="col-md-3 form-floating" title={this.state.formData.Machine}>
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
                                    <div className="col-md-3 form-floating" title={this.state.formData.Shift}>
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
                                    <div className="col-md-3 form-floating" title={this.state.formData.Supervisor}>
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
                                    </div>
                                    <div className="row py-3">
                                    <div className="col-md-3 form-floating" title={this.state.formData.ToolNumber}>
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

                                {/* Job Steps Table */}
                                <div className={'divSection'}>
                                <div className="SectionHeader">Job Steps</div>
                                    <table id="jobStepsTable">
                                        <thead>
                                            <tr className="bluebg">
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
                                <div className={'divSection'}>
                                <div className="SectionHeader">Permits</div>
                                    <table id="PermitsTable">
                                        <thead>
                                            <tr className="bluebg">
                                                <th>Type</th>
                                                <th>Required</th>
                                                <th>Acquired</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            {this.state.Permits.map(p => (
                                                <tr><td>{`${p.type}`}</td>
                                                    <td><input id={`Required_${p.id}`} type="checkbox" className="checkinputfeild" name={`chkPermits_Required_${p.id}`} onChange={this.handleChange} checked={p.Required} /></td>
                                                    <td><input id={`Acquired_${p.id}`} type="checkbox" className="checkinputfeild" name={`chkPermits_Acquired_${p.id}`} onChange={this.handleChange} checked={p.Acquired} /></td>
                                                </tr>))}
                                        </tbody>
                                    </table>
                                </div>

                                {/* PPE Requirements */}
                               
                                <div className={'divSection'}>
                                    <div className="SectionHeader">PPE Requirements</div>
                                    <table id="PPEREquirementsTable">
                                        <thead>
                                            <tr className="bluebg">
                                                <th >PPE Type</th>
                                                <th ></th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            {this.bindPPETypes()}
                                        </tbody>
                                    </table>
                                    <input type="button" value="Add PPE" className="addbutton" onClick={this.addPPEType} id="btnAddPPE" />
                                </div>

                                {/* Persons Involved */}
                                <div className={'divSection'}>
                                <div className="SectionHeader">Persons Involved</div>
                                    <table id="PersonsInvolvedTable">
                                        <thead>
                                            <tr className="bluebg">
                                                <th >Name</th>
                                                <th >Date</th>
                                                <td ></td>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            {this.bindPersons()}
                                        </tbody>
                                    </table>
                                    <input type="button" value="Add Person" className="addbutton" onClick={this.addPerson} id="btnAddPerson" />
                                </div>
                                {/* Supervisor Info */}
                                <div className={'row divSection'}>
                                <div className="SectionHeader">Supervisor</div>
                                    <div className="col-md-3">
                                        <div className="form-floating">
                                            <input className="form-control" placeholder="" name="SupervisorName" type="text" id="txtSupervisorName" ref={this.SupervisorName} value={this.state.formData.SupervisorName} onChange={this.handleChange} disabled={false} title={this.state.formData.SupervisorName} />
                                            <label className=" col-form-label">Supervisor Name </label>
                                        </div>
                                    </div>
                                    <div className="col-md-3 c-date-picker" id="divDate">
                                        <label className="label-datePicker" htmlFor="dtDate"> Date</label>
                                        <DatePickercontrol placeholder="" selectedDate={this.state.formData.SupervisorDate} id='dtSupervisorDate' isDisabled={false} startDate={undefined} endDate={undefined} name="SupervisorDate" onDatechange={(dateProps: any) => this.handleDateChange(dateProps[0], dateProps[2], "divSupervisorDate")} ref={this.SupervisorDate} highlightDate={new Date()} showIcon />
                                    </div>

                                </div>
                                {/* Buttons */}
                                <div className="row col-sm-12 text-center py-3" id="divButtons" >
                                    <div className="col-sm-10 text-center">
                                        <button type="button" id="btnSubmit" className="btn btn-primary" title="Submit" >Submit</button>
                                        <button type="button" id="btnCancel" className="btn btn-secondary" title="Cancel" onClick={this.handlCancel}>Cancel</button>
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