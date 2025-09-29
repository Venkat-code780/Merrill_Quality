import * as React from "react";
import { hideLoader, showLoader } from "../Shared/Loader";
import { Navigate } from "react-router-dom";
import { SPHttpClient } from "@microsoft/sp-http";
// import { SPFx, spfi } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "../CSS/TAGForm.css";
import { ActionStatus, ControlType } from "../Constants/Contants";
import { showToast } from "../Shared/Toaster";
import { highlightCurrentNav } from "../Utilities/HighlightCurrentComponent";
import DatePickercontrol from "../Shared/DatePickerField";
import SearchableDropdown from "../Shared/Dropdown";
import { initCommonFunctions } from "../Utilities/CommonFunctions";
import DateUtilities from "../Utilities/DateUtilities";
import { format } from "date-fns";
import { PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
// import FileUpload from "../Shared/FileUpload";
// import InputCheckBox from "../Shared/InputCheckBox";
// import { format } from "date-fns";
import Formvalidator from "../Utilities/FormValidator";

export interface TAGFormProps {
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

export interface TAGFormState {
}

export default class TAGForm extends React.Component<TAGFormProps, TAGFormState> {
    private rootSiteURL: string; private currentSiteURL: string; private MaycoURL: string;
    private Name: any; private TagType: any;
    private _peoplePickerContext = {
        absoluteUrl: this.props.context.pageContext.web.absoluteUrl,
        msGraphClientFactory: this.props.context.msGraphClientFactory,
        spHttpClient: this.props.context.spHttpClient
    };
    // private sp = spfi().using(SPFx(this.props.context)); 
    public state = {
        //Cascaded dropdowns
        Plants: [], Departments: [], Zones: [], Machines: [], Shifts: [],
        PlantsOpt: [], DepartmentsOpt: [], ZonesOpt: [], MachinesOpt: [], ShiftsOpt: [],
        formData: {
            Plant: '',
            Department: '',
            Zone: '',
            Machine: '',
            Shift: '',
            Name: '',
            Date: '',
        },
        SafetyformData: {
            TAG: '',
            NearMissSEWORequired: false,
            UnsafeConditionImmediateFixWorkOrder: false,
            UnsafeActCoachingCounselingTrainingOPLTWTTF: false,
            ProblemDetails: '',
            CounterMeasures: '',
            CompletedDate: '',
            CompletedByIds: [],
            CompletedByEmails: [],
            CompletedByTitles: [],
        },
        AMWOformData: {
            TAG: '',
            TagType: '',
            AMWOCheckList: [
                { id: '0', label: 'Oil leak', nameInList: 'Oil_x0020_Leak', isChecked: false },
                { id: '1', label: 'Wrong temperature', nameInList: 'Wrong_x0020_Temperature', isChecked: false },
                { id: '2', label: 'Water leak', nameInList: 'Water_x0020_Leak', isChecked: false },
                { id: '3', label: 'Wrong pressure', nameInList: 'Wrong_x0020_Pressure', isChecked: false },
                { id: '4', label: 'Grease / lube leak', nameInList: 'Grease_x0020__x002f__x0020_lube_', isChecked: false },
                { id: '5', label: 'Difficult to lube', nameInList: 'Difficult_x0020_To_x0020_Lube', isChecked: false },
                { id: '6', label: 'Air leak', nameInList: 'Air_x0020_Leak', isChecked: false },
                { id: '7', label: 'Difficult to clean', nameInList: 'Difficult_x0020_To_x0020_Clean', isChecked: false },
                { id: '8', label: 'Contamination', nameInList: 'Contamination', isChecked: false },
                { id: '9', label: 'Irregular function', nameInList: 'Irregular_x0020_Function', isChecked: false },
                { id: '10', label: 'Out of operating range', nameInList: 'Out_x0020_Of_x0020_Operating_x00', isChecked: false },
                { id: '11', label: 'Loose/missing fastner', nameInList: 'Loose_x002f_Missing_x0020_Fastne', isChecked: false },
                { id: '12', label: 'Vibration', nameInList: 'Vibration', isChecked: false },
                { id: '13', label: 'Other', nameInList: 'Other', isChecked: false },
            ],
            CompletedDate: '',
            CompletedByIds: [],
            CompletedByEmails: [],
            CompletedByTitles: [],
            Comments: '',
        },
        PMformData: {
            TAG: '',
            CompletedDate: '',
            CompletedByIds: [],
            CompletedByEmails: [],
            CompletedByTitles: [],
            Comments: '',
        },
        isEditForm: false,
        showSubmit: false,
        ItemId: 0,
        activeTag: 'Safety',
        Redirect: false,
        RedirectTo: '',
    }

    constructor(props: TAGFormProps) {
        super(props);
        this.rootSiteURL = this.props.spContext.siteAbsoluteUrl; //   sites/wcm
        this.currentSiteURL = this.props.spContext.webAbsoluteUrl;//  sites/wcm/mayco/merrill/sa
        this.MaycoURL = `${this.rootSiteURL}/mayco`;
        this.Name = React.createRef();
        this.TagType = React.createRef();
        //console.log(this.rootSiteURL);      //sites/wcm
    }

    public componentDidMount(): void {
        highlightCurrentNav("liTAGForm");
        document.title = "Mayco - Safety | TAG";
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
        let MachineList = 'Machines', MachineSelQuery = 'Title,Plant/Title,Plant/Id,Department/Title,Department/Id,Zone/Title,Zone/Id,*', MachineExpFields = 'Plant,Department,Zone';
        let ShiftList = 'Shifts';

        try {
            let [Plants, Departments, Zones, Machines, Shifts] = await Promise.all([
                getListItems(PlantList, this.MaycoURL),
                getListItems(DepartmentList, this.MaycoURL, DepartmentSelQuery, DepartmentExpFields),
                getListItems(ZoneList, this.MaycoURL, ZoneSelQuery, ZoneExpFields),
                getListItems(MachineList, this.MaycoURL, MachineSelQuery, MachineExpFields),
                getListItems(ShiftList, this.MaycoURL),
            ])
            let PlantsOpt = this.getMapedOptions(Plants, 'Title', 'Title');
            let ShiftsOpt = this.getMapedOptions(Shifts, 'Title', 'Title');
            // for sorting
            PlantsOpt.sort((a, b) => a.label.localeCompare(b.label));
            ShiftsOpt.sort((a, b) => a.label.localeCompare(b.label));

            let formData: any = { ...this.state.formData };
            let currPlantTitle = PlantsOpt.find((plant: any) => plant.label.toLowerCase() == this.props.currPlantTitle);
            formData.Plant = currPlantTitle ? currPlantTitle.label : '';
            this.setState({ formData, Plants, PlantsOpt, ShiftsOpt, Departments, Zones, Machines, showSubmit });
            this.onPlantChange(this.state, formData.Plant, true); // to get current plant departments,zones,machines
            //Edit mode starts here
            let stateData: any = { ...this.state };
            if (ItemId > 0) {
                stateData = await this.getTAGEditData(stateData, ItemId);
                let isEditForm = true;
                stateData = this.onPlantChange(stateData, formData.Plant, false);
                stateData = this.onDepartmentChange(stateData, formData.Plant, formData.Department, false);
                stateData = this.onZoneChange(stateData, formData.Plant, formData.Department, formData.Zone, false);
                //showSubmit = this.props.isSuperAdmin || (TAGData.Author.Title == this.props.userDisplayName) ? true : false;

                this.setState({ formData: stateData.formData, SafetyformData: stateData.SafetyformData, AMWOformData: stateData.AMWOformData, PMformData: stateData.PMformData, activeTag: stateData.activeTag, isEditForm, ItemId, showSubmit, DepartmentsOpt: stateData.DepartmentsOpt, ZonesOpt: stateData.ZonesOpt, MachinesOpt: stateData.MachinesOpt });
            }
        } catch (e) {
            console.log(e);
            this.onError();
        }
        finally {
            hideLoader();
        }
    }
    private getTAGEditData = async (stateData: any, ItemId: any) => {
        let { getListItems } = initCommonFunctions(this.props.context, this.rootSiteURL);
        let currentactiveTag = this.currentSiteURL.split('/')[this.currentSiteURL.split('/').length - 1]
        let selectedTagformData = `${stateData.activeTag}formData`;
        let updatedTagformData: any = {};
        let [TagListName, TagListURL] = ['', ''];
        if (currentactiveTag.toLocaleLowerCase() === 'sa') // Safety TAG
        {
            [TagListName, TagListURL] = ['Safety tag', `${this.rootSiteURL}/mayco/merrill/SA`];
            stateData.activeTag = 'Safety';
            selectedTagformData = `${stateData.activeTag}formData`;
            updatedTagformData = { ...(stateData as any)[selectedTagformData] };
        }
        else if (currentactiveTag.toLocaleLowerCase() === 'am')// AM TAG
        {
            [TagListName, TagListURL] = ['AM TAG', `${this.rootSiteURL}/mayco/merrill/AM`];
            stateData.activeTag = 'AMWO';
            selectedTagformData = `${stateData.activeTag}formData`;
            updatedTagformData = { ...(stateData as any)[selectedTagformData] };
        }
        else if (currentactiveTag.toLocaleLowerCase() === 'wo') // WO TAG
        {
            [TagListName, TagListURL] = ['WO Tag', `${this.rootSiteURL}/mayco/merrill/WO`];
            stateData.activeTag = 'AMWO';
            selectedTagformData = `${stateData.activeTag}formData`;
            updatedTagformData = { ...(stateData as any)[selectedTagformData] };
        }
        else if (currentactiveTag.toLocaleLowerCase() === 'pm') // PM TAG
        {
            [TagListName, TagListURL] = ['PM tag', `${this.rootSiteURL}/mayco/merrill/PM`];
            stateData.activeTag = 'PM';
            selectedTagformData = `${stateData.activeTag}formData`;
            updatedTagformData = { ...(stateData as any)[selectedTagformData] };
        }
        try {
            let TAGRes: any = await getListItems(TagListName, TagListURL, 'Author/Title,Author/Id,CompletedBy/Title,CompletedBy/Id,CompletedBy/EMail,*', 'Author,CompletedBy', `Id eq ${ItemId}`);

            if (!TAGRes.isHttpRequestError) {
                if (TAGRes.length) {
                    let TAGData = TAGRes[0];
                    stateData.formData.Plant = [null, undefined, '', 'None'].includes(TAGData.Plant) ? '' : TAGData.Plant;
                    stateData.formData.Department = [null, undefined, '', 'None'].includes(TAGData.Department) ? '' : TAGData.Department;
                    stateData.formData.Zone = [null, undefined, '', 'None'].includes(TAGData.Zone) ? '' : TAGData.Zone;
                    stateData.formData.Machine = [null, undefined, '', 'None'].includes(TAGData.Machine) ? '' : TAGData.Machine;
                    stateData.formData.Shift = [null, undefined, '', 'None'].includes(TAGData.Shifts) ? '' : TAGData.Shifts;
                    stateData.formData.Name = [null, undefined, '', 'None'].includes(TAGData.Name) ? '' : TAGData.Name;
                    stateData.formData.Date = [null, undefined, '', 'None'].includes(TAGData.Date) ? '' : TAGData.Date;
                    updatedTagformData.TAG = [null, undefined, '', 'None'].includes(TAGData.TAG_x0023_) ? '' : TAGData.TAG_x0023_;
                    updatedTagformData.CompletedDate = [null, undefined, '', 'None'].includes(TAGData.Completed_x0020_Date) ? '' : TAGData.Completed_x0020_Date;

                    if (TAGData.CompletedBy) {
                        let CompletedByIds = TAGData.CompletedBy.map((cb: any) => cb.Id);
                        let CompletedByEmails = TAGData.CompletedBy.map((cb: any) => cb.EMail);
                        let CompletedByTitles = TAGData.CompletedBy.map((cb: any) => cb.Title);
                        updatedTagformData.CompletedByIds = CompletedByIds;
                        updatedTagformData.CompletedByEmails = CompletedByEmails;
                        updatedTagformData.CompletedByTitles = CompletedByTitles;
                    }
                    if (stateData.activeTag === 'Safety') {
                        updatedTagformData.NearMissSEWORequired = [null, undefined].includes(TAGData.Near_x0020_Miss) ? false : TAGData.Near_x0020_Miss;
                        updatedTagformData.UnsafeConditionImmediateFixWorkOrder = [null, undefined].includes(TAGData.Unsafe_x0020_Condition) ? false : TAGData.Unsafe_x0020_Condition;
                        updatedTagformData.UnsafeActCoachingCounselingTrainingOPLTWTTF = [null, undefined].includes(TAGData.Unsafe_x0020_Act) ? false : TAGData.Unsafe_x0020_Act;
                        updatedTagformData.ProblemDetails = [null, undefined].includes(TAGData.Problem_x0020_Details) ? '' : TAGData.Problem_x0020_Details;
                        updatedTagformData.CounterMeasures = [null, undefined].includes(TAGData.Counter_x0020_Measures) ? '' : TAGData.Counter_x0020_Measures;
                        stateData.SafetyformData = updatedTagformData;

                    }
                    else if (stateData.activeTag === 'AMWO') {
                        updatedTagformData.TagType = [null, undefined].includes(TAGData.Tag_x0020_Type) ? '' : TAGData.Tag_x0020_Type;
                        updatedTagformData.Comments = [null, undefined,].includes(TAGData.Comments) ? '' : TAGData.Comments;
                        updatedTagformData.AMWOCheckList.forEach((cl: any, index: number) => // each activity checklist
                        {
                            updatedTagformData.AMWOCheckList[index].isChecked = [null, undefined].includes(TAGData[cl.nameInList]) ? false : TAGData[cl.nameInList];
                        }
                        )
                        stateData.AMWOformData = updatedTagformData;
                    }
                    else if (stateData.activeTag === 'PM') {
                        updatedTagformData.Comments = [null, undefined,].includes(TAGData.Comments) ? '' : TAGData.Comments;
                        stateData.PMformData = updatedTagformData;

                    }
                }

                else {
                    showToast("error", "No TAG found");
                    this.setState({ Redirect: true, RedirectTo: 'Home' });
                }
            }
        }
        catch (e) {
            console.log(e);
            this.onError();
        }

        return stateData;
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
                Plant: { val: formData.Plant, required: true, Name: "Plant", Type: ControlType.reactSelect, Focusid: "ddlPlant" },
                Department: { val: formData.Department, required: true, Name: "Department", Type: ControlType.reactSelect, Focusid: "ddlDepartment" },
                Zone: { val: formData.Zone, required: true, Name: "Zone", Type: ControlType.reactSelect, Focusid: "ddlZone" },
                Machine: { val: formData.Machine, required: true, Name: "Machine", Type: ControlType.reactSelect, Focusid: "ddlMachine" },
                Shift: { val: formData.Shift, required: true, Name: "Shift", Type: ControlType.reactSelect, Focusid: "ddlShift" },
                Name: { val: formData.Name, required: true, Name: "Name", Type: ControlType.string, Focusid: this.Name },
                Date: { val: formData.Date, required: true, Name: "Date", Type: ControlType.date, Focusid: "dtDate" },
                dateToday: { val: formData.Date, required: true, Name: "Date", Type: ControlType.lessthanTodayDate, Focusid: "dtDate" },
            }
            let isValid = Formvalidator.FormValidation(data);
            isValid = isValid.status ? this.validateTagForm() : isValid;
            if (isValid.status) {
                let TAGData: { TagListName: string, TagListURL: string, TagPostObj: {} } = this.getTAGPostData();
                this.InsertOrUpdateData(TAGData);
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
    private InsertOrUpdateData = async (TAGData: { TagListName: string, TagListURL: string, TagPostObj: {} }) => {
        let { addListItem, updateListItem } = initCommonFunctions(this.props.context, this.rootSiteURL);
        let ItemId = this.props.match.params.id;
        if (ItemId > 0) {
            await updateListItem(TAGData.TagListName, TAGData.TagListURL, TAGData.TagPostObj, ItemId).then(async (res: any) => {
                if (!res.isHttpRequestError) {
                    let msg = "TAG updated successfully";
                    this.onSuccess(msg);
                }
            }, (error: any) => {
                console.log(error);
                this.onError();
            })
        }
        else {
            await addListItem(TAGData.TagListName, TAGData.TagListURL, TAGData.TagPostObj).then(async (res: any) => {
                //let adedItemId = res.Id;
                console.log(res)
                if (!res.isHttpRequestError) {
                    let msg = "TAG submitted successfully";
                    this.onSuccess(msg);
                }
            }, (error: any) => {
                console.log(error);
                this.onError();
            })
        }
    }
    private onSuccess = (successMessage: string) => {
        hideLoader();
        let To = this.state.activeTag == 'Safety' ? 'TagView' : 'Home';
        this.setState({ Redirect: true, RedirectTo: To, ItemID: 0 });
        showToast("success", successMessage);
    }
    private getTAGPostData = () => {
        let stateData: any = { ...this.state };
        let selectedTagformData = `${stateData.activeTag}formData`;
        let updatedTagformData = { ...(this.state as any)[selectedTagformData] };
        let [TagListName, TagListURL] = ['', ''];
        let TagPostObj: any = {};
        //common post Data
        let TAGDate: any = DateUtilities.addBrowserwrtServer(new Date(DateUtilities.getDateMMDDYYYY(stateData.formData.Date)), this.props.spContext.webTimeZoneData).toISOString();
        let mmddyyyyTAGDate = format(TAGDate, "MM/dd/yyyy");
        let CompletedDate = ![null, undefined, ''].includes(updatedTagformData.CompletedDate) ? DateUtilities.addBrowserwrtServer(new Date(DateUtilities.getDateMMDDYYYY(updatedTagformData.CompletedDate)), this.props.spContext.webTimeZoneData).toISOString() : null;
        // let CompletedByIds = updatedTagformData.CompletedByIds.length > 0 ? { results : updatedTagformData.CompletedByIds } : null; this form not worked in this version
        let CompletedByIds = updatedTagformData.CompletedByIds.length > 0 ? updatedTagformData.CompletedByIds : [];
        let CompletedByTitles = updatedTagformData.CompletedByTitles.length > 0 ? updatedTagformData.CompletedByTitles.join(';') : '';
        TagPostObj = {
            Plant: stateData.formData.Plant,
            Department: stateData.formData.Department,
            Zone: stateData.formData.Zone,
            Machine: stateData.formData.Machine,
            Shifts: stateData.formData.Shift,
            Name: stateData.formData.Name,
            Date: TAGDate,
            Year: mmddyyyyTAGDate.split("/")[2],
            YearMonth: mmddyyyyTAGDate.split("/")[0],
            TAG_x0023_: updatedTagformData.TAG,
            Completed_x0020_Date: CompletedDate,
            CompletedById: CompletedByIds,
            Completed_x0020_By: CompletedByTitles,
        }
        if (stateData.activeTag === 'Safety') // Safety TAG
        {
            [TagListName, TagListURL] = ['Safety tag', `${this.rootSiteURL}/mayco/merrill/SA`];
            TagPostObj = this.getSATAGpostObject(TagPostObj);
        }
        else if (stateData.activeTag === 'AMWO') {
            if (updatedTagformData.TagType === 'AM') // AM TAG
            {
                [TagListName, TagListURL] = ['AM TAG', `${this.rootSiteURL}/mayco/merrill/AM`];
            }
            else if (updatedTagformData.TagType === 'WO') // WO TAG
            {
                [TagListName, TagListURL] = ['WO Tag', `${this.rootSiteURL}/mayco/merrill/WO`];
            }
            TagPostObj = this.getAMWOTAGpostObject(TagPostObj);
        }
        else if (stateData.activeTag === 'PM') // PM TAG
        {
            [TagListName, TagListURL] = ['PM tag', `${this.rootSiteURL}/mayco/merrill/PM`];
            // for  PM tag getpostObject method is not required , beacause of only one extra control  'Comments' is existed
            TagPostObj['Comments'] = updatedTagformData.Comments;
        }
        return { TagListName, TagListURL, TagPostObj };
    }
    private getSATAGpostObject = (TagPostObj: any) => {
        let stateData: any = { ...this.state };
        let SafetyformData = { ...stateData.SafetyformData };

        TagPostObj.Near_x0020_Miss = SafetyformData.NearMissSEWORequired;
        TagPostObj.Unsafe_x0020_Condition = SafetyformData.UnsafeConditionImmediateFixWorkOrder;
        TagPostObj.Unsafe_x0020_Act = SafetyformData.UnsafeActCoachingCounselingTrainingOPLTWTTF;
        TagPostObj.Problem_x0020_Details = SafetyformData.ProblemDetails;
        TagPostObj.Counter_x0020_Measures = SafetyformData.CounterMeasures;

        return TagPostObj;
    }
    private getAMWOTAGpostObject = (TagPostObj: any) => {
        let stateData: any = { ...this.state };
        let AMWOformData = { ...stateData.AMWOformData };

        TagPostObj.Tag_x0020_Type = AMWOformData.TagType;
        TagPostObj.Comments = AMWOformData.Comments;
        AMWOformData.AMWOCheckList.forEach((cl: any) => // each activity checklist
        {
            TagPostObj[cl.nameInList] = cl.isChecked;
        }
        )
        return TagPostObj;
    }


    // on change functions
    private handleChange = (event: any, activeTag: any = '') => {
        const formData: any = { ...this.state.formData };

        const name = event.target.name;
        let inputValue = (event.target.type == "text" || event.target.type == "textarea") ? event.target.value : (event.target.type == "checkbox" || event.target.type == "radio") ? event.target.checked : event.target.value;
        let classArr: any = event.target.className;

        if (classArr.includes("onlyNum")) {
            inputValue = inputValue.replace(/[^0-9]/g, '');
        }
        if (activeTag === 'Safety') // Safety form control change
        {
            let updatedSafetyformData = { ...this.state.SafetyformData };
            updatedSafetyformData = {
                ...updatedSafetyformData,
                [name]: inputValue,
            };
            this.setState({ SafetyformData: updatedSafetyformData });
        }
        else if (activeTag === 'AMWO') // AMWO form control change
        {
            let updatedAMWOformData = { ...this.state.AMWOformData };
            if (name.includes('AMWOCheckList')) // to handle AMWO checklist change
            {
                let chkIndex = name.split('_')[0];
                let updatedAMWOCheckList = [...updatedAMWOformData.AMWOCheckList];
                updatedAMWOCheckList[chkIndex].isChecked = inputValue;
                updatedAMWOformData = {
                    ...updatedAMWOformData,
                    AMWOCheckList: updatedAMWOCheckList,
                };
            }
            else {
                updatedAMWOformData = {
                    ...updatedAMWOformData,
                    [name]: inputValue,
                };
            }
            this.setState({ AMWOformData: updatedAMWOformData });
        }
        else if (activeTag === 'PM') // PM form control change
        {
            let updatedPMformData = { ...this.state.PMformData };
            updatedPMformData = {
                ...updatedPMformData,
                [name]: inputValue,
            };
            this.setState({ PMformData: updatedPMformData });
        }
        formData[name] = inputValue;
        this.setState({ formData });
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
        }
        else if (name == 'Department') {
            this.onDepartmentChange(this.state, formData.Plant, value, true);
        }
        else if (name == 'Zone') {
            this.onZoneChange(this.state, formData.Plant, formData.Department, value, true);
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
        if (divId.includes('divCompletedDate')) // to change CompletedDate in respective form Safety/AMWO/PM
        {
            let selectedTagformData = `${this.state.activeTag}formData`;
            let updatedTagformData = { ...(this.state as any)[selectedTagformData] };
            updatedTagformData = {
                ...updatedTagformData,
                CompletedDate: dateValue,
            };
            this.setState({ [selectedTagformData]: updatedTagformData });
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
        // update dependents
        DepartmentsOpt.sort((a: any, b: any) => a.label.localeCompare(b.label));
        stateData.DepartmentsOpt = DepartmentsOpt;
        if (isOnchange) {
            stateData.ZonesOpt = [];
            stateData.MachinesOpt = [];
            stateData.formData.Department = '';
            stateData.formData.Zone = '';
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
        // update dependents
        ZonesOpt.sort((a: any, b: any) => a.label.localeCompare(b.label));
        stateData.ZonesOpt = ZonesOpt;
        if (isOnchange) {
            stateData.MachinesOpt = [];
            stateData.formData.Zone = '';
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
        // update dependents
        MachinesOpt.sort((a: any, b: any) => a.label.localeCompare(b.label));
        stateData.MachinesOpt = MachinesOpt;
        if (isOnchange) {
            stateData.formData.Machine = '';
            this.setState(stateData);
        }
        return stateData;
    }
    private _getPeoplePickerItems(items: any, name: any) {
        let Ids: number[] = [];
        let Emails: string[] = [];
        let Titles: string[] = [];
        if (items.length > 0) {
            if (['CompletedBy'].includes(name)) {
                for (const user of items) {
                    Ids.push(user.id);
                    Emails.push(user.secondaryText);
                    Titles.push(user.text);
                }
            }
        }
        // to change CompletedBy in respective form Safety/AMWO/PM
        let selectedTagformData = `${this.state.activeTag}formData`;
        let updatedTagformData = { ...(this.state as any)[selectedTagformData] };
        updatedTagformData = {
            ...updatedTagformData,
            CompletedByIds: Ids,
            CompletedByTitles: Titles,
            CompletedByEmails: Emails,
        };
        this.setState({ [selectedTagformData]: updatedTagformData });
    }

    // Safety Tag functions
    private bindSafetyForm = () => {
        return <div className="tab-pane active" id="SafetyTab">

            <div className="row">
                <div className="col-md-3"><h1 className="text-left">Safety</h1></div>
                <div className="col-md-9">
                    <div className="light-text">
                        <label className="col-form-label">TAG# </label>
                        <input className="form-control onlyNum" placeholder="" name="TAG" type="text" id="txtTAG" value={this.state.SafetyformData.TAG} title={this.state.SafetyformData.TAG} onChange={(e) => this.handleChange(e, this.state.activeTag)} disabled={false} />
                    </div>
                </div>
            </div>
            <div className="col-md-12">
                <div className='mt-3'>
                    <input title={'Near Miss SEWO Required'} type='checkbox' checked={this.state.SafetyformData.NearMissSEWORequired} required={false} onChange={(e) => this.handleChange(e, this.state.activeTag)} name={'NearMissSEWORequired'} autoComplete="off" disabled={false} id={`chk_NearMissSEWORequired`} /> <label title={'Near Miss SEWO Required'} className="col-form-label chkLbl" htmlFor={`chk_NearMissSEWORequired`}>Near Miss SEWO Required</label>
                </div>
            </div>
            <div className="col-md-12">
                <div className='mt-3'>
                    <input title={'Unsafe Condition Immediate Fix/Work Order'} type='checkbox' checked={this.state.SafetyformData.UnsafeConditionImmediateFixWorkOrder} required={false} onChange={(e) => this.handleChange(e, this.state.activeTag)} name={'UnsafeConditionImmediateFixWorkOrder'} autoComplete="off" disabled={false} id={`chk_UnsafeConditionImmediateFix/WorkOrder`} /> <label title={'Unsafe Condition Immediate Fix/Work Order'} className="col-form-label chkLbl" htmlFor={`chk_UnsafeConditionImmediateFix/WorkOrder`}>Unsafe Condition Immediate Fix/Work Order</label>
                </div>
            </div>
            <div className="col-md-12">
                <div className='mt-3'>
                    <input title={'Unsafe Act Coaching,Counseling,Training,OPL TWTTF'} type='checkbox' checked={this.state.SafetyformData.UnsafeActCoachingCounselingTrainingOPLTWTTF} required={false} onChange={(e) => this.handleChange(e, this.state.activeTag)} name={'UnsafeActCoachingCounselingTrainingOPLTWTTF'} autoComplete="off" disabled={false} id={`chk_UnsafeActCoachingCounselingTrainingOPLTWTTF`} /> <label title={'Unsafe Act Coaching,Counseling,Training,OPL TWTTF'} className="col-form-label chkLbl" htmlFor={`chk_UnsafeActCoachingCounselingTrainingOPLTWTTF`}>Unsafe Act Coaching,Counseling,Training,OPL TWTTF</label>
                </div>
            </div>
            <div className="col-md-9 mb-3">
                <div className={"light-text"} >
                    <label className=" col-form-label" htmlFor="txtLocationPersons">Problem Details</label>
                    <textarea className="form-control bs-textarea" rows={3} id="txtProblemDetails" name="ProblemDetails" placeholder="Problem Details" value={this.state.SafetyformData.ProblemDetails} onChange={(e) => this.handleChange(e, this.state.activeTag)} disabled={false} title={this.state.SafetyformData.ProblemDetails}></textarea>
                </div>
            </div>
            <div className="col-md-9 mb-3">
                <div className={"light-text"} >
                    <label className=" col-form-label" htmlFor="txtCounterMeasures">Counter Measures</label>
                    <textarea className="form-control bs-textarea" rows={3} id="txtCounterMeasures" name="CounterMeasures" placeholder="Counter Measures" value={this.state.SafetyformData.CounterMeasures} onChange={(e) => this.handleChange(e, this.state.activeTag)} disabled={false} title={this.state.SafetyformData.CounterMeasures}></textarea>
                </div>
            </div>
            <div className="row">
                <div className="col-md-4" id="divDate">
                    <div className="light-text">
                        <label className="" >Completed Date {this.state.SafetyformData.CompletedByIds.length > 0 ? <span className="mandatoryhastrick">* </span> : ''}</label>
                        <div className="custom-datepicker" id="divCompletedDate">
                            <DatePickercontrol placeholder="" selectedDate={this.state.SafetyformData.CompletedDate} title={this.state.SafetyformData.CompletedDate} id='dtCompletedDate' isDisabled={false} startDate={undefined} endDate={new Date()} name="CompletedDate" onDatechange={(dateProps: any) => this.handleDateChange(dateProps[0], dateProps[2], "divCompletedDate")} highlightDate={new Date()} showIcon />
                        </div>
                    </div>
                </div>
                <div className="col-md-8">
                    <div className="light-text">
                        <label className='lblPeoplepicker'>Completed By{![null, undefined, ''].includes(this.state.SafetyformData.CompletedDate) ? <span className="mandatoryhastrick">*</span> : ''}</label>
                        <div className="custom-peoplepicker" id="divCompletedBy">
                            <PeoplePicker
                                context={this._peoplePickerContext}
                                titleText="Completed By"
                                personSelectionLimit={10}
                                showtooltip={false}
                                disabled={false}
                                defaultSelectedUsers={this.state.SafetyformData.CompletedByEmails}
                                onChange={(e) => this._getPeoplePickerItems(e, 'CompletedBy')}
                                ensureUser={true}
                                required={![null, undefined, ''].includes(this.state.SafetyformData.CompletedDate)}
                                principalTypes={[PrincipalType.User]}
                                placeholder=""
                                resolveDelay={1000}
                                peoplePickerCntrlclassName={"input-peoplePicker-custom"} />
                        </div>
                    </div>
                </div>
            </div>
        </div>
    }
    // Safety Tag functions
    private bindAMWOForm = () => {
        let AMWOCheckList = [...this.state.AMWOformData.AMWOCheckList];
        let CheckListElements = AMWOCheckList.map((item: any, index: number) => {
            return <div className="col-md-6" key={index}>
                <input title={item.label} type='checkbox' checked={item.isChecked} required={false} onChange={(e) => this.handleChange(e, this.state.activeTag)} name={`${index}_AMWOCheckList_${item.nameInList}`} autoComplete="off" disabled={false} id={`chk_AMWOCheckList_${item.nameInList}`} /> <label title={item.label} className="col-form-label chkLbl" htmlFor={`chk_AMWOCheckList_${item.nameInList}`}>{item.label}</label>
            </div>
        });

        return <div className="tab-pane active" id="AMWOTab">

            <div className="row mb-3">
                <div className="col-md-6"><h1 className="text-left">AM/WO</h1></div>
                <div className="col-md-6">
                    <div className="light-text">
                        <label className=" col-form-label">TAG# </label>
                        <input className="form-control onlyNum" placeholder="" name="TAG" type="text" id="txtTAG" value={this.state.AMWOformData.TAG} title={this.state.AMWOformData.TAG} onChange={(e) => this.handleChange(e, this.state.activeTag)} disabled={false} />
                    </div>
                </div>
            </div>
            <div className="row mb-3">
                <div className="col-md-9">
                    <div className="light-text">
                        <label className=" col-form-label">Tag Type<span className="mandatoryhastrick">* </span></label>
                        <select className="form-control" placeholder="" name="TagType" id="txtTagType" ref={this.TagType} value={this.state.AMWOformData.TagType} title={this.state.AMWOformData.TagType} onChange={(e) => this.handleChange(e, this.state.activeTag)} disabled={this.state.ItemId > 0}>
                            <option value="">None</option>
                            <option value="AM">AM</option>
                            <option value="WO">WO</option>
                        </select>
                    </div>
                </div>
            </div>
            <div className="row">
                <div className="col-md-12">
                    <h3>{'AUTONOMOUS ACTIVITIES TAG (WO/AM)'}</h3>
                </div>
            </div>
            <div className="row">
                {CheckListElements}
            </div>
            <div className="row mb-3">
                <div className="col-md-4" id="divDate">
                    <div className="light-text">
                        <label className="" >Completed Date{this.state.AMWOformData.CompletedByIds.length > 0 ? <span className="mandatoryhastrick">* </span> : ''}</label>
                        <div className="custom-datepicker" id="divCompletedDate">
                            <DatePickercontrol placeholder="" selectedDate={this.state.AMWOformData.CompletedDate} title={this.state.AMWOformData.CompletedDate} id='dtCompletedDate' isDisabled={false} startDate={undefined} endDate={new Date()} name="CompletedDate" onDatechange={(dateProps: any) => this.handleDateChange(dateProps[0], dateProps[2], "divCompletedDate")} highlightDate={new Date()} showIcon />
                        </div>
                    </div>
                </div>
                <div className="col-md-8">
                    <div className="light-text">
                        <label className='lblPeoplepicker'>Completed By{![null, undefined, ''].includes(this.state.AMWOformData.CompletedDate) ? <span className="mandatoryhastrick">*</span> : ''}</label>
                        <div className="custom-peoplepicker" id="divCompletedBy">
                            <PeoplePicker
                                context={this._peoplePickerContext}
                                titleText="Completed By"
                                personSelectionLimit={10}
                                showtooltip={false}
                                disabled={false}
                                defaultSelectedUsers={this.state.AMWOformData.CompletedByEmails}
                                onChange={(e) => this._getPeoplePickerItems(e, 'CompletedBy')}
                                ensureUser={true}
                                required={![null, undefined, ''].includes(this.state.AMWOformData.CompletedDate)}
                                principalTypes={[PrincipalType.User]}
                                placeholder=""
                                resolveDelay={1000}
                                peoplePickerCntrlclassName={"input-peoplePicker-custom"} />
                        </div>
                    </div>
                </div>
            </div>
            <div className="col-md-12">
                <div className={"light-text"} >
                    <textarea className="form-control bs-textarea" rows={3} id="txtComments" name="Comments" placeholder="Comments" value={this.state.AMWOformData.Comments} onChange={(e) => this.handleChange(e, this.state.activeTag)} disabled={false} title={this.state.AMWOformData.Comments}></textarea>
                    <label className=" col-form-label" htmlFor="txtComments">Comments</label>
                </div>
            </div>
        </div>
    }
    // PM Tag functions
    private bindPMForm = () => {
        return <div className="tab-pane active" id="PMTab">

            <div className="col-md-12"><h1 className="text-left">Professional Maintenance</h1></div>
            <div className="row mb-3">
                <div className="col-md-9 ms-auto">
                    <div className="light-text">
                        <label className=" col-form-label">TAG# </label>
                        <input className="form-control onlyNum" placeholder="" name="TAG" type="text" id="txtTAG" value={this.state.PMformData.TAG} title={this.state.PMformData.TAG} onChange={(e) => this.handleChange(e, this.state.activeTag)} disabled={false} />
                    </div>
                </div>
            </div>
            <div className="row mb-3">
                <div className="col-md-4" id="divDate">
                    <div className="light-text">
                        <label className="" >Completed Date{this.state.PMformData.CompletedByIds.length > 0 ? <span className="mandatoryhastrick">* </span> : ''}</label>
                        <div className="custom-datepicker" id="divCompletedDate">
                            <DatePickercontrol placeholder="" selectedDate={this.state.PMformData.CompletedDate} title={this.state.PMformData.CompletedDate} id='dtCompletedDate' isDisabled={false} startDate={undefined} endDate={new Date()} name="CompletedDate" onDatechange={(dateProps: any) => this.handleDateChange(dateProps[0], dateProps[2], "divCompletedDate")} highlightDate={new Date()} showIcon />
                        </div>
                    </div>
                </div>
                <div className="col-md-8">
                    <div className="light-text">
                        <label className='lblPeoplepicker'>Completed By{![null, undefined, ''].includes(this.state.PMformData.CompletedDate) ? <span className="mandatoryhastrick">*</span> : ''}</label>
                        <div className="custom-peoplepicker" id="divCompletedBy">
                            <PeoplePicker
                                context={this._peoplePickerContext}
                                titleText="Completed By"
                                personSelectionLimit={10}
                                showtooltip={false}
                                disabled={false}
                                defaultSelectedUsers={this.state.PMformData.CompletedByEmails}
                                onChange={(e) => this._getPeoplePickerItems(e, 'CompletedBy')}
                                ensureUser={true}
                                required={![null, undefined, ''].includes(this.state.PMformData.CompletedDate)}
                                principalTypes={[PrincipalType.User]}
                                placeholder=""
                                resolveDelay={1000}
                                peoplePickerCntrlclassName={"input-peoplePicker-custom"} />
                        </div>
                    </div>
                </div>
            </div>
            <div className="col-md-12">
                <div className={"light-text"} >
                    <label className=" col-form-label" htmlFor="txtComments">Comments</label>
                    <textarea className="form-control bs-textarea" rows={3} id="txtComments" name="Comments" placeholder="Comments" value={this.state.PMformData.Comments} onChange={(e) => this.handleChange(e, this.state.activeTag)} disabled={false} title={this.state.PMformData.Comments}></textarea>
                </div>
            </div>
        </div>
    }
    //common validation for all tags Safety/AMWO/PM
    private validateTagForm = () => {
        let selectedTagformData = `${this.state.activeTag}formData`;
        let updatedTagformData = { ...(this.state as any)[selectedTagformData] };
        //const updatedSafetyformData = {...this.state.SafetyformData};
        let isValid = { message: '', status: true };
        if (this.state.activeTag === 'AMWO' && [null, undefined, ''].includes(updatedTagformData.TagType)) {
            let data = {
                CompletedDate: { val: updatedTagformData.TagType, required: true, Name: 'Tag Type', Type: ControlType.string, Focusid: this.TagType },
            }
            isValid = Formvalidator.FormValidation(data);
            return isValid;
        }
        if (![null, undefined, ''].includes(updatedTagformData.CompletedDate)) {
            let data = {
                CompletedDate: { startDate: this.state.formData.Date, endDate: updatedTagformData.CompletedDate, startDateName: 'Date', endDateName: 'Completed Date', required: true, Name: "Date", Type: ControlType.compareDates, Focusid: "dtCompletedDate" },
                CompletedBy: { val: updatedTagformData.CompletedByIds, required: true, Name: 'Completed By', Type: ControlType.people, Focusid: 'divCompletedBy' },

            }

            // let pdata = {
            //     CompletedBy: { val: { results: updatedTagformData.CompletedByIds }, required: true, Name: 'Completed By', Type: ControlType.people, Focusid: 'divCompletedBy' },
            // }
            isValid = Formvalidator.FormValidation(data);
            // isValid = isValid.status ? Formvalidator.multiplePeoplePickerValidation(pdata) : isValid;
            return isValid;
        }
        else if (updatedTagformData.CompletedByIds.length > 0) {
            let data = {
                Date: { val: updatedTagformData.CompletedDate, required: true, Name: "Completed Date", Type: ControlType.date, Focusid: "dtCompletedDate" },
            }
            isValid = Formvalidator.FormValidation(data);
            return isValid;
        }
        return isValid;
    };
    private onError = () => {
        showToast("error", ActionStatus.Error);
        hideLoader();
    }
    private handlCancel = () => {
        this.setState({ Redirect: true, RedirectTo: 'TagView', ItemId: 0 });
    }
    private onTabClick = (TAGName: String, e?: React.MouseEvent<HTMLAnchorElement, MouseEvent>) => {
        showLoader();
        // Prevent tab click if ItemId exists and it's not the active tab
        if (this.state.ItemId > 0 && this.state.activeTag !== TAGName) {
            if (e) e.preventDefault();
            hideLoader();
            return;
        }
        this.setState({ activeTag: TAGName });
        hideLoader();
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
                                <h3 className="mb-0 pt-2 text-center">{" TAG " + (this.state.isEditForm ? (" - " + this.state.ItemId) : "")} </h3>
                                <label className="text-end px-1" style={{ width: "100%" }}> <span className="mandatoryhastrick">* </span> are mandatory fields</label>
                            </div>

                            <div className="mainContent row borderLine">
                                <div className="row py-2">
                                    <div className="col-md-3" title={this.state.formData.Plant}>
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
                                    <div className="col-md-3" title={this.state.formData.Department}>
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
                                    <div className="col-md-3" title={this.state.formData.Zone}>
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
                                    <div className="col-md-3" title={this.state.formData.Machine}>
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
                                </div>
                                <div className="row pb-2">

                                    <div className="col-md-3" title={this.state.formData.Shift}>
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
                                                noOptionsMessage="No Shift"
                                            />
                                        </div>
                                    </div>
                                    <div className="col-md-3">
                                        <div className="light-text">
                                            <input className="form-control" placeholder="" name="Name" type="text" id="txtName" ref={this.Name} value={this.state.formData.Name} title={this.state.formData.Name} onChange={this.handleChange} disabled={false} />
                                            <label className=" col-form-label">Name<span className="mandatoryhastrick">*</span> </label>
                                        </div>
                                    </div>
                                    <div className="col-md-3" id="divDate">
                                        <div className="light-text">
                                            <label className=""> Date <span className="mandatoryhastrick">*</span></label>
                                            <div className="custom-datepicker" id="divDate">
                                                <DatePickercontrol placeholder="" selectedDate={this.state.formData.Date} id='dtDate' isDisabled={false} startDate={undefined} endDate={new Date()} name="Date" onDatechange={(dateProps: any) => this.handleDateChange(dateProps[0], dateProps[2], "divDate")} highlightDate={new Date()} showIcon />
                                            </div>
                                        </div>
                                    </div>
                                </div>

                                {/* Multi TAGS Section */}
                                <div className="tab-container mb-4">
                                    <div className="tab-navigation">
                                        <ul className="nav nav-tabs nav-fill">
                                            <li className="nav-item" role="presentation">
                                                <a
                                                    className={`tab-link ${this.state.activeTag === 'Safety' ? 'active SafetyTab' : 'Safety'}`}
                                                    onClick={(e) => this.onTabClick('Safety', e)}
                                                >
                                                    Safety
                                                </a>
                                            </li>
                                            <li className="nav-item" role="presentation">
                                                <a
                                                    className={`tab-link ${this.state.activeTag === 'AMWO' ? 'active AMWOTab' : 'AMWO'}`}
                                                    onClick={(e) => this.onTabClick('AMWO', e)}
                                                >
                                                    AM/WO
                                                </a>
                                            </li>
                                            <li className="nav-item" role="presentation">
                                                <a
                                                    className={`tab-link ${this.state.activeTag === 'PM' ? 'active PMTab' : 'PM'}`}
                                                    onClick={(e) => this.onTabClick('PM', e)}
                                                >
                                                    PM
                                                </a>
                                            </li>
                                        </ul>
                                    </div>

                                    <div className={`tab-content ${this.state.activeTag === 'Safety' ? 'SafetyTabSection' : this.state.activeTag === 'AMWO' ? 'AMWOTabSection' : this.state.activeTag === 'PM' ? 'PMTabSection' : ''}`}>
                                        {this.state.activeTag === 'Safety' && (this.bindSafetyForm())}
                                        {this.state.activeTag === 'AMWO' && (this.bindAMWOForm())}
                                        {this.state.activeTag === 'PM' && (this.bindPMForm())}
                                    </div>

                                    {/* Buttons */}
                                    <div className="col-sm-12 text-center py-3" id="divButtons" >
                                        {this.state.showSubmit && <button type="button" id="btnSubmit" className="btn btn-primary mx-2" title={this.state.ItemId>0?'Update':'Submit'} onClick={this.handleSubmit} >{this.state.ItemId>0?'Update':'Submit'}</button>}
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