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
import "../CSS/JSRAForm.css";
import { ActionStatus, ControlType } from "../Constants/Contants";
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
import { faTrash, faAdd } from "@fortawesome/free-solid-svg-icons";
import Formvalidator from "../Utilities/FormValidator";

export interface JSRAFormProps {
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
            Date: '',
            Plant: '',
            Department: '',
            Zone: '',
            WorkCell: '',
            Machine: '',
            Shift: '',
            Supervisor: '',
            ToolNumber: '',
            SupervisorName: '',
            SupervisorDate: ''
        },
        Probability: [], Controls: [], Severity: [],
        ProbabilityOpt: [], ControlsOpt: [], SeverityOpt: [], PPETypesOpt: [],
        GroupedRiskFamilyandRiskObj: {}, RiskFamilyOpt: [],
        jobSteps: [
            {
                id: 1,
                Step: '',
                Required: true,
                RiskFamily: '',
                Risk: '',
                Probability: 0,
                Controls: 0,
                Severity: 0,
                RiskLevel: { Probability: '0', Controls: '0', Severity: '0' },
                RiskLevelColorClasses: { Probability: 'RiskLevel-0', Controls: 'RiskLevel-0', Severity: 'RiskLevel-0', Total: 'RiskLevel-0' },
                TotalScore: '0',
                RiskOpt: [],
            }
        ],
        Permits: [
            { type: "Hot Work", id: "0", Required: false, Acquired: false },
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
                PersonDate: ''
            }
        ],
        isEditForm: false,
        showSubmit: false,
        ItemId: 0,
        Redirect: false,
        RedirectTo: '',
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
        showLoader();
        let ItemId = this.props.match.params.id;
        let showSubmit = true;
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

            let formData: any = { ...this.state.formData };
            let currPlantTitle = PlantsOpt.find((plant: any) => plant.label.toLowerCase() == this.props.currPlantTitle);
            formData.Plant = currPlantTitle ? currPlantTitle.label : '';
            this.setState({ formData, Plants, PlantsOpt, ShiftsOpt, RiskFamilyOpt, GroupedRiskFamilyandRiskObj, ProbabilityOpt, ControlsOpt, SeverityOpt, PPETypesOpt, Departments, Zones, WorkCells, Machines, Supervisors, ToolNumbers, showSubmit });
            this.onPlantChange(this.state, formData.Plant, true); // to get current plant departments,zones,workcells,machines,supervisors,toolnumbers
            //Edit mode starts here
            let stateData: any = { ...this.state };
            if (ItemId > 0) {
                let jsraRes: any, JSRALineData: any;
                [jsraRes, JSRALineData] = await Promise.all([
                    getListItems('JSRA', this.currentSiteURL, 'Author/Title,Author/Id,*', 'Author', `Id eq ${ItemId}`),
                    getListItems('JSRA Line', this.currentSiteURL, '', '', `JSRA_x0020_ID eq '${ItemId}'`),
                ]);
                if (!jsraRes.isHttpRequestError) {
                    if (jsraRes.length) {
                        let jsraData = jsraRes[0];
                        formData.Date = [null, undefined, '', 'None'].includes(jsraData.Date) ? '' : jsraData.Date;
                        formData.Plant = [null, undefined, '', 'None'].includes(jsraData.Plant) ? '' : jsraData.Plant;
                        formData.Department = [null, undefined, '', 'None'].includes(jsraData.Department) ? '' : jsraData.Department;
                        formData.Zone = [null, undefined, '', 'None'].includes(jsraData.Zone) ? '' : jsraData.Zone;
                        formData.WorkCell = [null, undefined, '', 'None'].includes(jsraData.Work_x0020_Cell) ? '' : jsraData.Work_x0020_Cell;
                        formData.Machine = [null, undefined, '', 'None'].includes(jsraData.Machine) ? '' : jsraData.Machine;
                        formData.Shift = [null, undefined, '', 'None'].includes(jsraData.Shift) ? '' : jsraData.Shift;
                        formData.Supervisor = [null, undefined, '', 'None'].includes(jsraData.Supervisor) ? '' : jsraData.Supervisor;
                        formData.ToolNumber = [null, undefined, '', 'None'].includes(jsraData.Tool_x0020_Number) ? '' : jsraData.Tool_x0020_Number;
                        formData.SupervisorName = [null, undefined, '', 'None'].includes(jsraData.Supervisor_x0020_Name) ? '' : jsraData.Supervisor_x0020_Name;
                        formData.SupervisorDate = [null, undefined, '', 'None'].includes(jsraData.Supervisor_x0020_Date) ? '' : jsraData.Supervisor_x0020_Date;
                        let Permits = [...this.state.Permits];
                        Permits[0].Required = [null, undefined].includes(jsraData.Hot_x0020_Work_x0020_Permit_x002) ? false : jsraData.Hot_x0020_Work_x0020_Permit_x002;
                        Permits[0].Acquired = [null, undefined].includes(jsraData.Hot_x0020_Work_x0020_Permit_x0020) ? false : jsraData.Hot_x0020_Work_x0020_Permit_x0020;
                        Permits[1].Required = [null, undefined].includes(jsraData.Electrical_x0020_Permit_x0020_Re) ? false : jsraData.Electrical_x0020_Permit_x0020_Re;
                        Permits[1].Acquired = [null, undefined].includes(jsraData.Electrical_x0020_Permit_x0020_Ac) ? false : jsraData.Electrical_x0020_Permit_x0020_Ac;
                        Permits[2].Required = [null, undefined].includes(jsraData.Confined_x0020_Space_x0020_Permi) ? false : jsraData.Confined_x0020_Space_x0020_Permi;
                        Permits[2].Acquired = [null, undefined].includes(jsraData.Confined_x0020_Space_x0020_Permi0) ? false : jsraData.Confined_x0020_Space_x0020_Permi0;
                        let PPETypes = [...this.state.PPETypes];
                        if (![null, undefined, '', 'None'].includes(jsraData.PPE)) {
                            let ppeArr = jsraData.PPE.split(',').map((item: string) => item.trim());
                            PPETypes = ppeArr.map((item: string, index: number) => ({
                                id: index + 1,
                                PPEType: item,
                            }))
                        }
                        let Persons = [...this.state.Persons];
                        if (![null, undefined, '', 'None'].includes(jsraData.Personnel)) {
                            let personArr = jsraData.Personnel.split(',').map((item: string) => item.trim());
                            Persons = personArr.map((item: string, index: number) => ({
                                id: index + 1,
                                PersonName: item.split(';#')[0],
                                PersonDate: item.split(';#')[1] ?? ''
                            }))
                        }
                        let jobSteps = [...this.state.jobSteps];
                        if (JSRALineData.length > 0) {
                            jobSteps = JSRALineData.map((item: any, index: number) => ({
                                JSRALineitmeId: item.Id,
                                id: index + 1,
                                Step: [null, undefined, '', 'None'].includes(item.Details) ? '' : item.Details,
                                Required: [null, undefined, ''].includes(item.Required) ? true : item.Required,
                                RiskFamily: [null, undefined, '', 'None'].includes(item.Category) ? '' : item.Category,
                                Risk: [null, undefined, '', 'None'].includes(item.Sub_x0020_Category) ? '' : item.Sub_x0020_Category,
                                Probability: [null, undefined, '', 'None'].includes(item.Probability_x0020_Value) ? 0 : Number(item.Probability_x0020_Value),
                                Controls: [null, undefined, '', 'None'].includes(item.Controls_x0020_Value) ? 0 : Number(item.Controls_x0020_Value),
                                Severity: [null, undefined, '', 'None'].includes(item.Severity_x0020_Value) ? 0 : Number(item.Severity_x0020_Value),
                                RiskLevel: {
                                    Probability: [null, undefined, '', 'None'].includes(item.Probability_x0020_Value) ? 0 : item.Probability_x0020_Value,
                                    Controls: [null, undefined, '', 'None'].includes(item.Controls_x0020_Value) ? 0 : item.Controls_x0020_Value,
                                    Severity: [null, undefined, '', 'None'].includes(item.Severity_x0020_Value) ? 0 : item.Severity_x0020_Value
                                },
                                RiskLevelColorClasses: { Probability: `RiskLevel-${item.Probability_x0020_Value}`, Controls: `RiskLevel-${item.Controls_x0020_Value}`, Severity: `RiskLevel-${item.Severity_x0020_Value}`, Total: `RiskLevel-${item.Total_x0020_Score == 0 ? 0 : item.Total_x0020_Score <= 4 ? 1 : item.Total_x0020_Score <= 6 ? 2 : item.Total_x0020_Score >= 7 ? 3 : 0}` },
                                TotalScore: [null, undefined, '', 'None'].includes(item.Total_x0020_Score) ? '0' : item.Total_x0020_Score,
                                RiskOpt: item.RiskOpt,
                            }))
                            jobSteps.forEach((js, index) => {
                                let ItemRiskOpts = [null, undefined, '', 'None'].includes(js.RiskFamily) ? [] : GroupedRiskFamilyandRiskObj[js.RiskFamily];
                                let RiskOpt = this.getMapedOptions(ItemRiskOpts, 'Title', 'Title');
                                jobSteps[index].RiskOpt = RiskOpt as any;
                                js.RiskOpt.sort((a: any, b: any) => a.label.localeCompare(b.label));
                            }
                            );
                        }
                        let isEditForm = true;
                        stateData = this.onPlantChange(stateData, formData.Plant, false);
                        stateData = this.onDepartmentChange(stateData, formData.Plant, formData.Department, false);
                        stateData = this.onZoneChange(stateData, formData.Plant, formData.Department, formData.Zone, false);
                        showSubmit = this.props.isSuperAdmin || (jsraData.Author.Title == this.props.userDisplayName) ? true : false;

                        this.setState({ formData, isEditForm, ItemId, Permits, PPETypes, Persons, jobSteps, showSubmit, DepartmentsOpt: stateData.DepartmentsOpt, ZonesOpt: stateData.ZonesOpt, WorkCellsOpt: stateData.WorkCellsOpt, MachinesOpt: stateData.MachinesOpt, SupervisorsOpt: stateData.SupervisorsOpt, ToolNumbersOpt: stateData.ToolNumbersOpt });
                    }
                    else {
                        showToast("error", "No JSRA found");
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
    public handleSubmit = () => {
        try {
            showLoader();
            var formData: any = { ...this.state.formData };
            var data = {
                Date: { val: formData.Date, required: true, Name: "Date", Type: ControlType.date, Focusid: "dtDate" },
                dateToday: { val: formData.Date, required: true, Name: "Date", Type: ControlType.lessthanTodayDate, Focusid: "dtDate" },
                Plant: { val: formData.Plant, required: true, Name: "Plant", Type: ControlType.reactSelect, Focusid: "ddlPlant" },
                Department: { val: formData.Department, required: true, Name: "Department", Type: ControlType.reactSelect, Focusid: "ddlDepartment" },
                Zone: { val: formData.Zone, required: true, Name: "Zone", Type: ControlType.reactSelect, Focusid: "ddlZone" },
                WorkCell: { val: formData.WorkCell, required: true, Name: "Work Cell", Type: ControlType.reactSelect, Focusid: "ddlWorkCell" },
                Machine: { val: formData.Machine, required: true, Name: "Machine", Type: ControlType.reactSelect, Focusid: "ddlMachine" },
                ToolNumber: { val: formData.ToolNumber, required: true, Name: "Tool Number", Type: ControlType.reactSelect, Focusid: "ddlToolNumber" },
            }
            if (!this.state.isToolNumberMandatory) {
                delete (data as any).ToolNumber;
            }

            let isValid = Formvalidator.FormValidation(data);
            isValid = isValid.status ? this.validateJobSteps() : isValid;
            isValid = isValid.status && this.state.PPETypes.length > 1 ? this.validatePPEType() : isValid;
            isValid = isValid.status && this.state.Persons.length > 1 ? this.validatePerson() : isValid;

            if (isValid.status) {
                let postObjectJSRA: any = this.getJSRApostObject();
                this.InsertOrUpdateData(postObjectJSRA);
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
    private InsertOrUpdateData = async (postObjectJSRA: any) => {
        let { addListItem, updateListItem, batchAddItems, batchUpdateItems } = initCommonFunctions(this.props.context, this.rootSiteURL);
        let ItemId = this.props.match.params.id;
        if (ItemId > 0) {
            await updateListItem('JSRA', this.currentSiteURL, postObjectJSRA, ItemId).then(async (res: any) => {
                let postEditObjectJSRALineItems: any[] = this.getJSRALinepostObject(ItemId, this.state.jobSteps, this.state.isEditForm);
                await batchUpdateItems('JSRA Line', this.currentSiteURL, postEditObjectJSRALineItems).then((res: any) => {
                    let msg = "JSRA updated successfully";
                    this.onSuccess(msg);
                }, (error: any) => {
                    console.log(error);
                    this.onError();
                })

            }, (error: any) => {
                console.log(error);
                this.onError();
            })
        }
        else {
            await addListItem('JSRA', this.currentSiteURL, postObjectJSRA).then(async (res: any) => {
                let adedItemId = res.Id;
                // JSRA Line Items insertion
                let postObjectJSRALineItems: any[] = this.getJSRALinepostObject(adedItemId, this.state.jobSteps, this.state.isEditForm);
                await batchAddItems('JSRA Line', this.currentSiteURL, postObjectJSRALineItems).then((res: any) => {
                    let msg = "JSRA submitted successfully";
                    this.onSuccess(msg);
                }, (error: any) => {
                    console.log(error);
                    this.onError();
                })
            }, (error: any) => {
                console.log(error);
                this.onError();
            })
        }
    }
    private onSuccess = (successMessage: string) => {
        hideLoader();
        this.setState({ Redirect: true, RedirectTo: 'JSRAView', ItemID: 0 });
        showToast("success", successMessage);
    }
    private getJSRApostObject = () => {
        let stateData: any = { ...this.state };
        let JSRADate: any = DateUtilities.addBrowserwrtServer(new Date(DateUtilities.getDateMMDDYYYY(stateData.formData.Date)), this.props.spContext.webTimeZoneData).toISOString();
        let SupervisorDt: any = '';
        if (![null, undefined, ''].includes(stateData.formData.SupervisorDate))
            SupervisorDt = DateUtilities.addBrowserwrtServer(new Date(DateUtilities.getDateMMDDYYYY(stateData.formData.SupervisorDate)), this.props.spContext.webTimeZoneData).toISOString();
        let mmddyyyyJSRADate = format(JSRADate, "MM/dd/yyyy");
        //PPE Items getting
        let PPEItems = '';
        stateData.PPETypes.forEach((item: any, index: number) => {
            if (item.PPEType != '') {
                if (index == 0)
                    PPEItems = item.PPEType;
                else
                    PPEItems += ', ' + item.PPEType;
            }
        })
        //Personnel Items getting
        let PersonnelItems = '';
        stateData.Persons.forEach((item: any, index: number) => {
            let PersonDate: any = [null, undefined, ''].includes(item.PersonDate) ? '' : DateUtilities.getDateMMDDYYYY(item.PersonDate);
            if (item.PersonName != '') {
                if (index == 0)
                    PersonnelItems = item.PersonName + ';#' + PersonDate;
                else
                    PersonnelItems += ', ' + item.PersonName + ';#' + PersonDate;
            }
        })



        let postObjectJSRA = {
            Date: JSRADate,
            Year: mmddyyyyJSRADate.split("/")[2],
            YearMonth: mmddyyyyJSRADate.split("/")[0],
            Plant: stateData.formData.Plant,
            Department: stateData.formData.Department,
            Zone: stateData.formData.Zone,
            Work_x0020_Cell: stateData.formData.WorkCell,
            Machine: stateData.formData.Machine,
            Shift: stateData.formData.Shift,
            Supervisor: stateData.formData.Supervisor,
            Tool_x0020_Number: stateData.formData.ToolNumber,
            Supervisor_x0020_Name: stateData.formData.SupervisorName,
            Supervisor_x0020_Date: SupervisorDt,
            Hot_x0020_Work_x0020_Permit_x002: stateData.Permits[0].Required,
            Hot_x0020_Work_x0020_Permit_x0020: stateData.Permits[0].Acquired,
            Electrical_x0020_Permit_x0020_Re: stateData.Permits[1].Required,
            Electrical_x0020_Permit_x0020_Ac: stateData.Permits[1].Acquired,
            Confined_x0020_Space_x0020_Permi: stateData.Permits[2].Required,
            Confined_x0020_Space_x0020_Permi0: stateData.Permits[2].Acquired,
            PPE: PPEItems,
            Personnel: PersonnelItems,
        }
        return postObjectJSRA;
    }
    private getJSRALinepostObject = (JSRAItemId: any, jobSteps: any, isEditForm: boolean) => {
        let stateData: any = { ...this.state };
        let JSRADate: any = DateUtilities.addBrowserwrtServer(new Date(DateUtilities.getDateMMDDYYYY(stateData.formData.Date)), this.props.spContext.webTimeZoneData).toISOString();
        let postObjectJSRALineItems: any[] = [];
        let postEditObjectJSRALineItems: any[] = [];

        jobSteps.forEach((item: any, index: number) => {
            let postObjectJSRALine: any = {};
            postObjectJSRALine = {
                Date: JSRADate,
                Plant: stateData.formData.Plant,
                Department: stateData.formData.Department,
                Zone: stateData.formData.Zone,
                Machine: stateData.formData.Machine,
                Shift: stateData.formData.Shift,
                Tool_x0020_Number: stateData.formData.ToolNumber,
            }
            // labels and values are diffrent in Probability,Controls,Severity dropdowns, so below code is to get label text
            let Probability = document.getElementById(`Probability_${item.id}`) as HTMLSelectElement;
            let ProbabilityText: any = (Probability?.getElementsByClassName('css-1dimb5e-singleValue')[0] as HTMLElement)?.innerText;
            let Controls = document.getElementById(`Controls_${item.id}`) as HTMLSelectElement;
            let ControlsText: any = (Controls?.getElementsByClassName('css-1dimb5e-singleValue')[0] as HTMLElement)?.innerText;
            let Severity = document.getElementById(`Severity_${item.id}`) as HTMLSelectElement;
            let SeverityText: any = (Severity?.getElementsByClassName('css-1dimb5e-singleValue')[0] as HTMLElement)?.innerText;
            if (!isEditForm) {
                postObjectJSRALine['JSRA_x0020_ID'] = String(JSRAItemId);
                postObjectJSRALine['Details'] = item.Step;
                postObjectJSRALine['Required'] = item.Required;
                postObjectJSRALine['Category'] = item.RiskFamily;
                postObjectJSRALine['Sub_x0020_Category'] = item.Risk;
            }
            postObjectJSRALine['Probability'] = ProbabilityText;
            postObjectJSRALine['Controls'] = ControlsText;
            postObjectJSRALine['Severity'] = SeverityText;
            postObjectJSRALine['Probability_x0020_Value'] = String(item.Probability);
            postObjectJSRALine['Controls_x0020_Value'] = String(item.Controls);
            postObjectJSRALine['Severity_x0020_Value'] = String(item.Severity);
            postObjectJSRALine['Total_x0020_Score'] = Number(item.TotalScore);
            if (isEditForm) {
                postEditObjectJSRALineItems.push({ Id: item.JSRALineitmeId, data: { ...postObjectJSRALine } });
            }
            else {
                postObjectJSRALineItems.push(postObjectJSRALine);
            }
        });
        return isEditForm ? postEditObjectJSRALineItems : postObjectJSRALineItems;
    }




    // on change functions
    private handleChange = (event: any) => {
        const formData: any = { ...this.state.formData };

        const name = event.target.name;
        let inputValue = (event.target.type == "text" || event.target.type == "textarea") ? event.target.value : (event.target.type == "checkbox" || event.target.type == "radio") ? event.target.checked : event.target.value;
        let classArr: any = event.target.className;

        if (classArr.includes("onlyNum")) {
            inputValue = inputValue.replace(/[^0-9]/g, '');
        }
        if (name.includes('chkPermits')) {
            let Permits = [...this.state.Permits];
            let index = name.split('_')[2], field = name.split('_')[1];
            Permits[index] = {
                ...Permits[index],
                [field]: inputValue
            }

            this.setState({ Permits });
        }
        else {
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
            this.onPlantChange(this.state, value, true);
            // let filteredDepts = this.state.Departments.filter((dept: any) => dept.Plant.Title == value && dept.IsActiveMRO);// Note:in ventureglobal IsActive column exists , in WCM  IsActiveMRO column exists
            // let DepartmentsOpt = filteredDepts.map((item: any) => ({
            //     label: item.Title,
            //     value: item.Title,
            // }));
            // let filteredSupervisors = this.state.Supervisors.filter((sup: any) => sup.Plant?.Title == value);
            // let SupervisorsOpt = filteredSupervisors.map((item: any) => ({
            //     label: item.Title,
            //     value: item.Title,
            // }));
            // let filteredToolNumbers = this.state.ToolNumbers.filter((TN: any) => TN.Plant?.Title == value);
            // let ToolNumbersOpt = filteredToolNumbers.map((item: any) => ({
            //     label: item.Title,
            //     value: item.Title,
            // }));
            // // update dependents
            // DepartmentsOpt.sort((a, b) => a.label.localeCompare(b.label));
            // SupervisorsOpt.sort((a, b) => a.label.localeCompare(b.label));
            // ToolNumbersOpt.sort((a, b) => a.label.localeCompare(b.label));
            // this.setState({ formData, DepartmentsOpt, SupervisorsOpt, ToolNumbersOpt, ZonesOpt: [], WorkCellsOpt: [], MachinesOpt: [], Department: '', Zone: '', WorkCell: '', Machine: '' });
        }
        else if (name == 'Department') {
            this.onDepartmentChange(this.state, formData.Plant, value, true);
            // let filteredZones = this.state.Zones.filter((dept: any) => dept.Plant.Title == formData.Plant && dept.Department.Title == value);
            // let ZonesOpt = filteredZones.map((item: any) => ({
            //     label: item.Title,
            //     value: item.Title,
            // }));
            // //dynamic ToolNumber Mandatory
            // let isToolNumberMandatory = false;
            // if (value.toLowerCase() == 'molding') {
            //     isToolNumberMandatory = true;
            // }
            // // update dependents
            // ZonesOpt.sort((a, b) => a.label.localeCompare(b.label));
            //this.setState({ formData, ZonesOpt, WorkCellsOpt: [], MachinesOpt: [], Zone: '', WorkCell: '', Machine: '', isToolNumberMandatory });
        }
        else if (name == 'Zone') {
            this.onZoneChange(this.state, formData.Plant, formData.Department, value, true);
            // let filteredMachines = this.state.Machines.filter((dept: any) => dept.Plant.Title == formData.Plant && dept.Department.Title == formData.Department && dept.Zone.Title == value);
            // let MachinesOpt = filteredMachines.map((item: any) => ({
            //     label: item.Title,
            //     value: item.Title,
            // }));
            // let filteredWorkCells = this.state.WorkCells.filter((dept: any) => dept.h7kc == formData.Plant && dept.Department == formData.Department && dept.Zone == value);
            // let WorkCellsOpt = filteredWorkCells.map((item: any) => ({
            //     label: item.Title,
            //     value: item.Title,
            // }));
            // // update dependents
            // MachinesOpt.sort((a, b) => a.label.localeCompare(b.label));
            // WorkCellsOpt.sort((a, b) => a.label.localeCompare(b.label));
            // this.setState({ formData, MachinesOpt, WorkCellsOpt, WorkCell: '', Machine: '' });
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
    private onPlantChange = (state: any, PlantVal: any, isOnchange: boolean) => {
        let stateData: any = { ...state };
        let filteredDepts = stateData.Departments.filter((dept: any) => dept.Plant.Title == PlantVal && dept.IsActiveMRO);// Note:in ventureglobal IsActive column exists , in WCM  IsActiveMRO column exists
        let DepartmentsOpt = filteredDepts.map((item: any) => ({
            label: item.Title,
            value: item.Title,
        }));
        let filteredSupervisors = stateData.Supervisors.filter((sup: any) => sup.Plant?.Title == PlantVal);
        let SupervisorsOpt = filteredSupervisors.map((item: any) => ({
            label: item.Title,
            value: item.Title,
        }));
        let filteredToolNumbers = stateData.ToolNumbers.filter((TN: any) => TN.Plant?.Title == PlantVal);
        let ToolNumbersOpt = filteredToolNumbers.map((item: any) => ({
            label: item.Title,
            value: item.Title,
        }));
        // update dependents
        DepartmentsOpt.sort((a: any, b: any) => a.label.localeCompare(b.label));
        SupervisorsOpt.sort((a: any, b: any) => a.label.localeCompare(b.label));
        ToolNumbersOpt.sort((a: any, b: any) => a.label.localeCompare(b.label));
        stateData.DepartmentsOpt = DepartmentsOpt;
        stateData.SupervisorsOpt = SupervisorsOpt;
        stateData.ToolNumbersOpt = ToolNumbersOpt;
        if (isOnchange) {
            stateData.ZonesOpt = [];
            stateData.WorkCellsOpt = [];
            stateData.MachinesOpt = [];
            stateData.formData.Department = '';
            stateData.formData.Zone = '';
            stateData.formData.WorkCell = '';
            stateData.formData.Machine = '';
            this.setState(stateData);
        }
        return stateData;

    }
    private onDepartmentChange = (state: any, PlantVal: any, DepartmentVal: any, isOnchange: boolean) => {
        let stateData: any = { ...state };
        let filteredZones = stateData.Zones.filter((dept: any) => dept.Plant.Title == PlantVal && dept.Department.Title == DepartmentVal);
        let ZonesOpt = filteredZones.map((item: any) => ({
            label: item.Title,
            value: item.Title,
        }));
        //dynamic ToolNumber Mandatory
        let isToolNumberMandatory = false;
        if (DepartmentVal.toLowerCase() == 'molding') {
            isToolNumberMandatory = true;
        }
        // update dependents
        ZonesOpt.sort((a: any, b: any) => a.label.localeCompare(b.label));
        stateData.ZonesOpt = ZonesOpt;
        stateData.isToolNumberMandatory = isToolNumberMandatory;
        if (isOnchange) {
            stateData.WorkCellsOpt = [];
            stateData.MachinesOpt = [];
            stateData.formData.Zone = '';
            stateData.formData.WorkCell = '';
            stateData.formData.Machine = '';
            this.setState(stateData);
        }
        return stateData;
    }
    private onZoneChange = (state: any, PlantVal: any, DepartmentVal: any, ZoneVal: any, isOnchange: boolean) => {
        let stateData: any = { ...state };
        let filteredMachines = stateData.Machines.filter((dept: any) => dept.Plant.Title == PlantVal && dept.Department.Title == DepartmentVal && dept.Zone.Title == ZoneVal);
        let MachinesOpt = filteredMachines.map((item: any) => ({
            label: item.Title,
            value: item.Title,
        }));
        let filteredWorkCells = stateData.WorkCells.filter((dept: any) => dept.h7kc == PlantVal && dept.Department == DepartmentVal && dept.Zone == ZoneVal);
        let WorkCellsOpt = filteredWorkCells.map((item: any) => ({
            label: item.Title,
            value: item.Title,
        }));
        // update dependents
        MachinesOpt.sort((a: any, b: any) => a.label.localeCompare(b.label));
        WorkCellsOpt.sort((a: any, b: any) => a.label.localeCompare(b.label));
        stateData.MachinesOpt = MachinesOpt;
        stateData.WorkCellsOpt = WorkCellsOpt;
        if (isOnchange) {
            stateData.formData.WorkCell = '';
            stateData.formData.Machine = '';
            this.setState(stateData);
        }
        return stateData;
    }

    // Job Steps Section
    private bindJobSteps = () => {
        let jobSteps = this.state.jobSteps;
        let DynamicHTML: JSX.Element[] = [];

        DynamicHTML = jobSteps.map((jobStep, index) => (
            <tr key={jobStep.id}>
                <td className="text-center">{jobStep.id}</td>
                <td>
                    <input className="form-control" placeholder={``} name={`JobStep_${jobStep.id}`} type="text" id={`JobStep_${jobStep.id}`} value={jobStep.Step} title={jobStep.Step} onChange={(e) => this.handleJobStepChange(index, 'Step', e.target.value)} disabled={this.state.isEditForm} />

                </td>
                <td className="text-center">
                    <input
                        type="checkbox"
                        checked={jobStep.Required}
                        onChange={(e) => this.handleJobStepChange(index, 'Required', e.target.checked)} disabled={this.state.isEditForm} />
                </td>
                <td>
                    {/* <div className="custom-dropdown" id={`divRiskFamily_${jobStep.id}`} title={jobStep.RiskFamily}> */}
                        <SearchableDropdown
                            label={""}
                            Title={jobStep.RiskFamily}
                            name={""}
                            id={`RiskFamily_${jobStep.id}`}
                            placeholderText={""}
                            className={""}
                            selectedValue={jobStep.RiskFamily}
                            OptionsList={this.state.RiskFamilyOpt}
                            OnChange={(selectedOption: any, actionMeta: any) => this.handleJobStepChange(index, 'RiskFamily', selectedOption ? selectedOption.value : '')}
                            isRequired={false}
                            disabled={(!jobStep.Required) || this.state.isEditForm}
                            noOptionsMessage="No Risk Family"
                        />
                    {/* </div> */}
                </td>
                <td>
                    {/* <div className="custom-dropdown" id={`divRisk_${jobStep.id}`} title={jobStep.Risk}> */}
                        <SearchableDropdown
                            label={""}
                            Title={jobStep.Risk}
                            name={""}
                            id={`Risk_${jobStep.id}`}
                            placeholderText={""}
                            className={""}
                            selectedValue={jobStep.Risk}
                            OptionsList={jobStep.RiskOpt}
                            OnChange={(selectedOption: any, actionMeta: any) => this.handleJobStepChange(index, 'Risk', selectedOption ? selectedOption.value : '')}
                            isRequired={false}
                            disabled={(!jobStep.Required) || this.state.isEditForm}
                            noOptionsMessage="No Risk"
                        />
                    {/* </div> */}
                </td>
                <td>
                    <div className="bs-field form-floating" title={String(jobStep.Probability)}>
                        <div className="custom-dropdown" id={`divProbability_${jobStep.id}`}>
                            <SearchableDropdown
                                label={"Probability"}
                                Title={String(jobStep.Probability)}
                                name={"Probability"}
                                id={`Probability_${jobStep.id}`}
                                placeholderText={""}
                                className={""}
                                selectedValue={jobStep.Probability}
                                OptionsList={this.state.ProbabilityOpt}
                                OnChange={(selectedOption: any, actionMeta: any) => this.handleJobStepChange(index, 'Probability', selectedOption ? selectedOption.value : '0')}
                                isRequired={jobStep.Required}
                                disabled={!jobStep.Required}
                                noOptionsMessage="No Probability"
                            />
                        </div>
                    </div>
                    <div className="bs-field form-floating" title={String(jobStep.Controls)}>
                        <div className="custom-dropdown" id={`divControls_${jobStep.id}`}>
                            <SearchableDropdown
                                label={"Controls"}
                                Title={String(jobStep.Controls)}
                                name={"Controls"}
                                id={`Controls_${jobStep.id}`}
                                placeholderText={""}
                                className={""}
                                selectedValue={jobStep.Controls}
                                OptionsList={this.state.ControlsOpt}
                                OnChange={(selectedOption: any, actionMeta: any) => this.handleJobStepChange(index, 'Controls', selectedOption ? selectedOption.value : '0')}
                                isRequired={jobStep.Required}
                                disabled={!jobStep.Required}
                                noOptionsMessage="No Controls"
                            />
                        </div>
                    </div>
                    <div className="bs-field form-floating" title={String(jobStep.Severity)}>
                        <div className="custom-dropdown" id={`divSeverity_${jobStep.id}`}>
                            <SearchableDropdown
                                label={"Severity"}
                                Title={String(jobStep.Severity)}
                                name={"Severity"}
                                id={`Severity_${jobStep.id}`}
                                placeholderText={""}
                                className={""}
                                selectedValue={jobStep.Severity}
                                OptionsList={this.state.SeverityOpt}
                                OnChange={(selectedOption: any, actionMeta: any) => this.handleJobStepChange(index, 'Severity', selectedOption ? selectedOption.value : '0')}
                                isRequired={jobStep.Required}
                                disabled={!jobStep.Required}
                                noOptionsMessage="No Severity"
                            />
                        </div>
                    </div>
                </td>

                <td>
                    <div className={`${jobStep.RiskLevelColorClasses.Probability}`} title="Probability Risk Level">{jobStep.RiskLevel.Probability}</div>
                    <div className={`${jobStep.RiskLevelColorClasses.Controls}`} title="Controls Risk Level">{jobStep.RiskLevel.Controls}</div>
                    <div className={`${jobStep.RiskLevelColorClasses.Severity}`} title="Severity Risk Level">{jobStep.RiskLevel.Severity}</div>
                </td>
                <td>
                    <div className={`${jobStep.RiskLevelColorClasses.Total}`} title="Total Score">{jobStep.TotalScore}</div>
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
            let TotalScore = Number(updatedJobSteps[index].Probability) + Number(updatedJobSteps[index].Controls) + Number(updatedJobSteps[index].Severity);
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
                Required: true,
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
        let JobSteps: string[] = [];

        for (let i = 0; i < updatedJobSteps.length; i++) {
            const Step = updatedJobSteps[i];

            // Validate 'Step' field
            if (!Step.Step.trim()) {
                isValid = { message: `'Job Step' cannot be blank.`, status: false };
                document.getElementById(`JobStep_${Step.id}`)?.classList.add('mandatory-FormContent-focus');
                document.getElementById(`JobStep_${Step.id}`)?.focus();
                break;
            }
            if (JobSteps.includes(Step.Step.trim().toLowerCase())) {
                isValid = { message: `'Job Step' cannot be duplicate. (${Step.Step})`, status: false };
                document.getElementById(`JobStep_${Step.id}`)?.classList.add('mandatory-FormContent-focus');
                document.getElementById(`JobStep_${Step.id}`)?.focus();
                break;
            }
            else {
                JobSteps.push(Step.Step.trim().toLowerCase());
            }

            // If Required is true, validate other fields
            if (Step.Required) {
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
                    {/* <div className="custom-dropdown" id={`divPPE_${PPE.id}`} title={PPE.PPEType}> */}
                        <SearchableDropdown
                            label={""}
                            Title={PPE.PPEType}
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
                    {/* </div> */}
                </td>
                <td>{PPETypes.length > 1 ? <button type='button' className="btn text-danger" onClick={() => this.deletePPEType(index)} title="Delete PPE Type"><FontAwesomeIcon icon={faTrash} /></button> : ''}</td>
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
                    <input className="form-control" placeholder={``} name={`PersonName_${Person.id}`} type="text" id={`PersonName_${Person.id}`} value={Person.PersonName} title={Person.PersonName} onChange={(e) => this.handlePersonsChange(index, 'PersonName', e.target.value)} disabled={false} />

                </td>
                <td>
                    {/* <div className="c-date-picker"> */}
                        <DatePickercontrol placeholder="" selectedDate={Person.PersonDate} title={Person.PersonDate} isDisabled={false} id={`PersonDate_${Person.id}`} startDate={undefined} endDate={undefined} name={`PersonDate_${index}`} onDatechange={(dateProps: any) => this.handleDateChange(dateProps[0], dateProps[2], "divPersonDate")} highlightDate={new Date()} showIcon />
                    {/* </div> */}
                </td>
                <td>{Persons.length > 1 ? <button type='button' className="btn text-danger" onClick={() => this.deletePerson(index)} title="Delete Person"><FontAwesomeIcon icon={faTrash} /></button> : ''}</td>
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
                PersonDate: ''
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
                setTimeout(() => { document.getElementById(`PersonDate_${Person.id}`)?.classList.add('mandatory-FormContent-focus'); }, 100);
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
        this.setState({ Redirect: true, RedirectTo: 'Home', ItemId: 0 });
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
                            <div className="m-0 titlebg">
                                <h4 className="mb-0 pt-2 text-center">{" JSRA " + (this.state.isEditForm ? (" - " + this.state.ItemId) : "")} </h4>
                                <label className="text-end px-1" style={{ width: "100%" }}> <span className="text-danger">* </span> are mandatory fields</label>
                            </div>

                            <div className="mainContent px-4 borderLine">
                                <div className="row py-3">
                                    <div className="col-md-3 c-date-picker form-floating" id="divDate">
                                        <label className="label-datePicker"> Date <span className="text-danger">*</span></label>
                                        <DatePickercontrol placeholder="" selectedDate={this.state.formData.Date} id='dtDate' isDisabled={false} startDate={undefined} endDate={new Date()} name="Date" onDatechange={(dateProps: any) => this.handleDateChange(dateProps[0], dateProps[2], "divDate")} ref={this.Date} highlightDate={new Date()} showIcon />
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
                                                isRequired={false}
                                                disabled={false}
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
                                                isRequired={false}
                                                disabled={false}
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
                                                noOptionsMessage="No Tool Number"
                                            />
                                        </div>
                                    </div>
                                </div>

                                {/* Job Steps Table */}
                                <div className={'divSection'}>
                                    <div className="SectionHeader">Job Steps</div>
                                    <table id="jobStepsTable" className="TablejobSteps">
                                        <thead>
                                            <tr className="bluebg">
                                                <th></th>
                                                <th>Job Step</th>
                                                <th className="text-center">Required</th>
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
                                    {!this.state.isEditForm && <button type="button" value="Add Job Step" title="Add Job Step" className="addbutton" onClick={this.addJobStep} id="btnAddJobStep" ><FontAwesomeIcon icon={faAdd} />Add Job Step</button>}
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
                                <div className="row">
                                    {/* PPE Requirements */}

                                    <div className={'divSection'} style={{width:'30%'}}>
                                        <div className="SectionHeader">PPE Requirements</div>
                                        <table id="PPEREquirementsTable">
                                            <thead>
                                                <tr className="bluebg">
                                                    <th >PPE Type</th>
                                                    <th >{this.state.PPETypes.length>1?"Delete":""}</th>
                                                </tr>
                                            </thead>
                                            <tbody>
                                                {this.bindPPETypes()}
                                            </tbody>
                                        </table>
                                        <button type="button" value="Add PPE" title="Add PPE" className="addbutton" onClick={this.addPPEType} id="btnAddPPE" ><FontAwesomeIcon icon={faAdd} /> Add PPE</button>
                                    </div>

                                    {/* Persons Involved */}
                                    <div className={'divSection'} style={{width:'70%'}}>
                                        <div className="SectionHeader">Persons Involved</div>
                                        <table id="PersonsInvolvedTable">
                                            <thead>
                                                <tr className="bluebg">
                                                    <th >Name</th>
                                                    <th >Date</th>
                                                    <td >{this.state.Persons.length>1?"Delete":""}</td>
                                                </tr>
                                            </thead>
                                            <tbody>
                                                {this.bindPersons()}
                                            </tbody>
                                        </table>
                                        <button type="button" value="Add Person" title="Add Person" className="addbutton" onClick={this.addPerson} id="btnAddPerson"><FontAwesomeIcon icon={faAdd} /> Add Person</button>
                                    </div>
                                </div>
                                {/* Supervisor Info */}
                                <div className={'row divSection'}>
                                    <div className="SectionHeader">Supervisor</div>
                                    <div className="col-md-3 p-2">
                                        <div className="form-floating">
                                            <input className="form-control" placeholder="" name="SupervisorName" type="text" id="txtSupervisorName" ref={this.SupervisorName} value={this.state.formData.SupervisorName} title={this.state.formData.SupervisorName} onChange={this.handleChange} disabled={false} />
                                            <label className=" col-form-label">Supervisor Name </label>
                                        </div>
                                    </div>
                                    <div className="col-md-3 p-2 c-date-picker" id="divDate">
                                        <label className="label-datePicker" > Date</label>
                                        <DatePickercontrol placeholder="" selectedDate={this.state.formData.SupervisorDate} title={this.state.formData.SupervisorDate} id='dtSupervisorDate' isDisabled={false} startDate={undefined} endDate={undefined} name="SupervisorDate" onDatechange={(dateProps: any) => this.handleDateChange(dateProps[0], dateProps[2], "divSupervisorDate")} ref={this.SupervisorDate} highlightDate={new Date()} showIcon />
                                    </div>

                                </div>
                                {/* Buttons */}
                                <div className="col-sm-12 text-center py-3" id="divButtons" >
                                    {this.state.showSubmit && <button type="button" id="btnSubmit" className="btn btn-primary mx-2" title="Submit" onClick={this.handleSubmit} >Submit</button>}
                                    <button type="button" id="btnCancel" className="btn btn-secondary" title="Cancel" onClick={this.handlCancel}>Cancel</button>
                                </div>
                            </div>
                        </div>
                    </div>
                </React.Fragment>

            )
        }
    }

}