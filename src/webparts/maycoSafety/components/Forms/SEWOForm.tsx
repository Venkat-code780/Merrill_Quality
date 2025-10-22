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
import "../CSS/form-input-style.css";
import { ActionStatus, ControlType } from "../Constants/Contants";
import { showToast } from "../Shared/Toaster";
import { highlightCurrentNav } from "../Utilities/HighlightCurrentComponent";
import SearchableDropdown from "../Shared/Dropdown";
import DateUtilities from "../Utilities/DateUtilities";
import { format } from "date-fns";
import { initCommonFunctions } from "../Utilities/CommonFunctions";
import DatePickercontrol from "../Shared/DatePickerField";
import formValidation from "../Utilities/FormValidator";
import BodyPart from "../Utilities/BodyChart";
import Sketch, { SketchHandle } from "../Utilities/Sketch";
import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";
import { faChartArea, faChartLine, faCheck, faCheckCircle, faCheckDouble, faFileSignature, faPencil, faSearch, faUser, faUserInjured, faUserTie, faWarning } from "@fortawesome/free-solid-svg-icons";
import FileUpload from "../Shared/FileUpload";

export interface SEWOFormProps {
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
    currentUserGroups: any;
    isWCM: boolean;
}

export interface SEWOFormState {
}

export default class SEWOForm extends React.Component<SEWOFormProps, SEWOFormState> {

    private sp = spfi().using(SPFx(this.props.context));
    private MaycoURL: string;
    private currPlantObj: any;
    private SEWOList = "SEWO";
    private txtNameofInjured: any;
    private txtInjuredJob: any;
    private txtReportedBy: any;
    private rdIsHospitalRefused: any;
    private txtNameoftheHospital: any;
    private txtWhat: any;
    private txtWhen: any;
    private txtWhere: any;
    private txtWho: any;
    private txtWhich: any;
    private txtHow: any;
    private txtActionDescription: any;
    private txtFiveWhy1: any;
    private txtFiveWhy2: any;
    private txtFiveWhy3: any;
    private txtFiveWhy4: any;
    private txtFiveWhy5: any;
    private txtFiveWhyRootCause: any;
    private txtActionPlan: any;
    private txtResponsible: any;
    private txtNotes: any;
    private txtComments: any;
    private txtLocation: any;
    private txtAct: any;
    private txtDaysOff: any;
    private txtEmployeeName: any;
    private txtTeamLeadName: any;
    private txtSupervisorName: any;
    private txtDeptManagerName: any;
    private txtSafetyManagerName: any;
    private txtPlantManagerName: any;
    private txtEmployeeSignature: any;
    private txtTeamLeadSignature: any;
    private txtSupervisorSignature: any;
    private txtDeptManagerSignature: any;
    private txtSafetyManagerSignature: any;
    private txtPlantManagerSignature: any;
    private txtInjuredStatement: any;
    private txtWitnessStatement: any;
    private txtInjuredSignature: any;
    private txtWitnessSignature: any;
    private sketchRef = React.createRef<SketchHandle>();

    private dateFields = [
        'Reported_x0020_Date',
        'DueDate',
        'BackToWork',
        'InjuredDate',
        'WitnessDate',
        'EmployeeDate',
        'TeamLeadDate',
        "CloseDate",
        'SupervisorDate',
        'DeptManagerDate',
        'SafetyMgrDate',
        'PlantMgrDate'
    ];

    private optionalLookUpFields = [
        "AccidentCauseId",
        "ActionId",
        "ExpansionPlanId",
        "PPEInUseId",
        "PPESuppliedId",
        "StatusId",
        "UsualWorkId"
    ]

    public state = {
        formData: {
            AccidentCauseId: '',
            AccidentTypeId: '',
            Act: '',
            ActionId: '',
            ActionDescription: '',
            ActionPlan: '',
            BackToWork: '',
            BodyPartId: '',
            CloseDate: '',
            Comments: '',
            DaysOff: '',
            Department: '',
            DeptManagerDate: '',
            DeptManagerName: '',
            DeptManagerSignature: '',
            DueDate: '',
            EmployeeDate: '',
            EmployeeName: '',
            EmployeeSignature: '',
            ExpansionPlanId: '',
            FiveWhy1: '',
            FiveWhy2: '',
            FiveWhy3: '',
            FiveWhy4: '',
            FiveWhy5: '',
            FiveWhyRootCause: '',
            HHow: '',
            InjuredDate: '',
            InjuredJob: '',
            InjuredSignature: '',
            InjuredStatement: '',
            Injury_x0020_Date_x0020_Time: '',
            InjuryTime: '',
            InjuryTypeId: '',
            IsHospitalRefused: false,
            Location: '',
            Machine: '',
            MicroRootCauseId: '',
            InjuredName: '',
            NameoftheHospital: '',
            Notes: '',
            Plant: '',
            PlantMgrDate: '',
            PlantMgrName: '',
            PlantMgrSignature: '',
            PositionType: '',
            PPEInUseId: '',
            PPESuppliedId: '',
            ReportedBy: '',
            Reported_x0020_Date: '',
            Responsible: '',
            RootCauseId: '',
            SafetyMgrDate: '',
            SafetyMgrName: '',
            SafetyMgrSignature: '',
            SecondaryRootCauseId: '',
            Sex: '',
            Shift: '',
            Sketch: '',
            StatusId: '',
            SupervisorDate: '',
            SupervisorName: '',
            SuperVisorSignature: '',
            TeamLeadDate: '',
            TeamLeadName: '',
            TeamLeadSignature: '',
            Title: '',
            UsualWorkId: '',
            WitnessDate: '',
            WitnessSignature: '',
            WitnessStatement: '',
            WWhat: '',
            WWhen: '',
            WWhere: '',
            WWhich: '',
            WWho: '',
            Year: '',
            YearMonth: '',
            Zone: ''
        },
        plantsData: [],
        departmentData: [],
        departmentOptions: [],
        zoneData: [],
        zoneOptions: [],
        machineData: [],
        machineOptions: [],
        shiftData: [],
        accidentTypeData: [],
        accidentCauseData: [],
        injuryTypeData: [],
        bodyPartsData: [],
        rootCausesData: [],
        secondaryRootCausesData: [],
        secondaryRootCausesOptions: [],
        microRootCausesData: [],
        microRootCausesOptions: [],
        actionsData: [],
        actionsOptions: [],
        yesNoData: [],
        statusData: [],
        isInputDisabled: false,
        isEditForm: false,
        ItemId: 0,
        Redirect: false,
        RedirectTo: '',
        displayMessage: '',
        showSubmit: false,
        isAdmin: false,
        statusText: '',
        selBodyPart: [],
        fileArr: [],
        delfileArr: [],
        isFileCloseShow: true,
    }

    constructor(props: SEWOFormProps) {
        super(props);
        this.MaycoURL = `${this.props.siteURL}/mayco`;

        this.txtNameofInjured = React.createRef();
        this.txtInjuredJob = React.createRef();
        this.txtReportedBy = React.createRef();
        this.rdIsHospitalRefused = React.createRef();
        this.txtNameoftheHospital = React.createRef();
        this.txtWhat = React.createRef();
        this.txtWhen = React.createRef();
        this.txtWhere = React.createRef();
        this.txtWho = React.createRef();
        this.txtWhich = React.createRef();
        this.txtHow = React.createRef();
        this.txtActionDescription = React.createRef();
        this.txtFiveWhy1 = React.createRef();
        this.txtFiveWhy2 = React.createRef();
        this.txtFiveWhy3 = React.createRef();
        this.txtFiveWhy4 = React.createRef();
        this.txtFiveWhy5 = React.createRef();
        this.txtFiveWhyRootCause = React.createRef();
        this.txtResponsible = React.createRef();
        this.txtNotes = React.createRef();
        this.txtComments = React.createRef();
        this.txtLocation = React.createRef();
        this.txtAct = React.createRef();
        this.txtDaysOff = React.createRef();
        this.txtEmployeeName = React.createRef();
        this.txtTeamLeadName = React.createRef();
        this.txtSupervisorName = React.createRef();
        this.txtDeptManagerName = React.createRef();
        this.txtSafetyManagerName = React.createRef();
        this.txtPlantManagerName = React.createRef();
        this.txtEmployeeSignature = React.createRef();
        this.txtTeamLeadSignature = React.createRef();
        this.txtSupervisorSignature = React.createRef();
        this.txtDeptManagerSignature = React.createRef();
        this.txtSafetyManagerSignature = React.createRef();
        this.txtPlantManagerSignature = React.createRef();
        this.txtInjuredStatement = React.createRef();
        this.txtWitnessStatement = React.createRef();
        this.txtInjuredSignature = React.createRef();
        this.txtWitnessSignature = React.createRef();
    }

    public componentDidMount(): void {
        highlightCurrentNav("liSEWOForm");
        document.title = "Mayco - Safety | SEWO";
        document.getElementById("divDepartment")?.getElementsByTagName('input')[0].focus();
        this.getOnLoadData();
    }

    private getOnLoadData = async () => {
        try {
            showLoader();
            var formData: any = { ...this.state.formData };
            let itemId = this.props.match.params.id;
            let showSubmit = false;
            let selBodyPart: any = [];

            let { getListItems } = initCommonFunctions(this.props.context, this.props.siteURL);
            let PlantList = 'Plant', PlantSelQuery = 'Title,*', plantFiltQuery = '', PlantExpFields = '';
            let DepartmentList = 'Department', DepartmentSelQuery = 'Title,Plant/Title,Plant/Id,*', DepartmentFiltQuery = '', DepartmentExpFields = 'Plant';
            let ZoneList = 'Zones', ZoneSelQuery = 'Title,Plant/Title,Plant/Id,Department/Title,Department/Id,*', ZoneFiltQuery = '', ZoneExpFields = 'Plant,Department';
            let MachineList = 'Machines', MachineSelQuery = 'Title,Plant/Title,Plant/Id,Department/Title,Department/Id,Zone/Title,Zone/Id,*', MachineFiltQuery = '', MachineExpFields = 'Plant,Department,Zone';
            let ShiftsList = 'Shifts', ShiftsSelQuery = 'Title,*', ShiftsFiltQuery = '', ShiftsExpFields = '';
            let accidentTypesSelQuery = 'Title,*', accidentTypesFiltQuery = '', accidentTypesExpFields = '', accidentTypesList = "AccidentTypes";
            let accidentCauseSelQuery = 'Title,*', accidentCauseFiltQuery = '', accidentCauseExpFields = '', accidentCauseList = "AccidentCause";
            let injuryTypeSelQuery = 'Title,*', injuryTypeFiltQuery = '', injuryTypeExpFields = '', injuryTypeList = "InjuryTypes";
            let bodyPartsSelQuery = 'Title,*', bodyPartsFiltQuery = '', bodyPartsExpFields = '', bodyPartsList = "BodyParts";
            let rootCauseSelQuery = 'Title,*', rootCauseFiltQuery = '', rootCauseExpFields = '', rootCauseList = "RootCauses";
            let secondaryRootCauseSelQuery = 'RootCause/Title,RootCause/Id,Title,*', secondaryRootCauseFiltQuery = '', secondaryRootCauseExpFields = 'RootCause', secondaryRootCauseList = "SecondaryRootCauses";
            let microRootCauseSelQuery = 'RootCause/Title,RootCause/Id,SecondaryRootCause/Title,SecondaryRootCause/Id,Title,*', microRootCauseFiltQuery = '', microRootCauseExpFields = 'RootCause,SecondaryRootCause', microRootCauseList = "MicroRootCauses";
            let actionsSelQuery = 'RootCause/Title,RootCause/Id,SecondaryRootCause/Title,SecondaryRootCause/Id,Title,*', actionsFiltQuery = '', actionsExpFields = 'RootCause,SecondaryRootCause', actionsList = "Actions";
            let yesNoSelQuery = 'Title,*', yesNoFiltQuery = '', yesNoExpFields = '', yesNoList = "YesNo";
            let statusSelQuery = 'Title,*', statusFiltQuery = '', statusExpFields = '', statusList = "Status";
            let [Plants, departmentData, zoneData, machineData, shifts, accidentTypes, accidentCause, injuryTypes, bodyParts, rootCauses, secondaryRootCausesData, microRootCausesData, actionsData, yesNo, status] = await Promise.all([
                getListItems(PlantList, this.MaycoURL, PlantSelQuery, PlantExpFields, plantFiltQuery),
                getListItems(DepartmentList, this.MaycoURL, DepartmentSelQuery, DepartmentExpFields, DepartmentFiltQuery),
                getListItems(ZoneList, this.MaycoURL, ZoneSelQuery, ZoneExpFields, ZoneFiltQuery),
                getListItems(MachineList, this.MaycoURL, MachineSelQuery, MachineExpFields, MachineFiltQuery),
                getListItems(ShiftsList, this.MaycoURL, ShiftsSelQuery, ShiftsExpFields, ShiftsFiltQuery),
                getListItems(accidentTypesList, this.props.webAbsoluteURL, accidentTypesSelQuery, accidentTypesExpFields, accidentTypesFiltQuery),
                getListItems(accidentCauseList, this.props.webAbsoluteURL, accidentCauseSelQuery, accidentCauseExpFields, accidentCauseFiltQuery),
                getListItems(injuryTypeList, this.props.webAbsoluteURL, injuryTypeSelQuery, injuryTypeExpFields, injuryTypeFiltQuery),
                getListItems(bodyPartsList, this.props.webAbsoluteURL, bodyPartsSelQuery, bodyPartsExpFields, bodyPartsFiltQuery),
                getListItems(rootCauseList, this.props.webAbsoluteURL, rootCauseSelQuery, rootCauseExpFields, rootCauseFiltQuery),
                getListItems(secondaryRootCauseList, this.props.webAbsoluteURL, secondaryRootCauseSelQuery, secondaryRootCauseExpFields, secondaryRootCauseFiltQuery),
                getListItems(microRootCauseList, this.props.webAbsoluteURL, microRootCauseSelQuery, microRootCauseExpFields, microRootCauseFiltQuery),
                getListItems(actionsList, this.props.webAbsoluteURL, actionsSelQuery, actionsExpFields, actionsFiltQuery),
                getListItems(yesNoList, this.props.webAbsoluteURL, yesNoSelQuery, yesNoExpFields, yesNoFiltQuery),
                getListItems(statusList, this.props.webAbsoluteURL, statusSelQuery, statusExpFields, statusFiltQuery)
            ]);

            Plants.sort((a, b) => a.Title.localeCompare(b.Title));
            departmentData.sort((a, b) => a.Title.localeCompare(b.Title));
            zoneData.sort((a, b) => a.Title.localeCompare(b.Title));
            machineData.sort((a, b) => a.Title.localeCompare(b.Title));
            accidentTypes.sort((a, b) => a.Title.localeCompare(b.Title));
            accidentCause.sort((a, b) => a.Title.localeCompare(b.Title));
            injuryTypes.sort((a, b) => a.Title.localeCompare(b.Title));
            bodyParts.sort((a, b) => a.Title.localeCompare(b.Title));
            rootCauses.sort((a, b) => a.Title.localeCompare(b.Title));
            secondaryRootCausesData.sort((a, b) => a.Title.localeCompare(b.Title));
            microRootCausesData.sort((a, b) => a.Title.localeCompare(b.Title));
            actionsData.sort((a, b) => a.Title.localeCompare(b.Title));
            yesNo.sort((a, b) => a.Title.localeCompare(b.Title));
            status.sort((a, b) => a.Title.localeCompare(b.Title));
            let plantsData = Plants.map((item: any) => ({ label: item.Title, value: item.Title }));
            this.currPlantObj = plantsData.find((plant: any) => plant.label.toLowerCase() == this.props.currPlantTitle);
            formData.Plant = this.currPlantObj.label;

            let departmentOptions = departmentData.filter((option: any) => option.Plant.Title == formData.Plant).map((item: any) => ({ label: item.Title, value: item.Title, id: item.Id }));
            let zoneOptions: any = [];
            let machineOptions: any = [];
            let secondaryRootCausesOptions: any = [];
            let microRootCausesOptions: any = [];
            let actionsOptions: any = [];
            let shiftData = shifts.map((item: any) => ({ label: item.Title, value: item.Title }));
            let accidentTypeData = accidentTypes.map((item: any) => ({ label: item.Title, value: item.Id }));
            let accidentCauseData = accidentCause.map((item: any) => ({ label: item.Title, value: item.Id }));
            let injuryTypeData = injuryTypes.map((item: any) => ({ label: item.Title, value: item.Id }));
            let bodyPartsData = bodyParts.map((item: any) => ({ label: item.Title, value: item.Id }));
            let rootCausesData = rootCauses.map((item: any) => ({ label: item.Title, value: item.Id }));
            let yesNoData = yesNo.map((item: any) => ({ label: item.Title, value: item.Id }));
            let statusData = status.map((item: any) => ({ label: item.Title, value: item.Id }));
            let filesArry: any[] = [];
            let isEditForm = false;
            if (itemId != undefined) {
                let SEWORes: any, SEWOAttach: any;
                [SEWORes] = await Promise.all([getListItems(this.SEWOList, this.props.webAbsoluteURL, 'Author/Title,Author/Id,*', 'Author', `Id eq ${itemId}`)]);
                if (!SEWORes.isHttpRequestError) {
                    if (SEWORes.length) {
                       SEWOAttach= await this.sp.web.lists.getByTitle(this.SEWOList).items.getById(itemId).attachmentFiles();
                        SEWOAttach.map(async (selItem: any) => {
                            let name = selItem.FileName;
                            var fileUrl = selItem.ServerRelativeUrl;
                            let obj = { URL: fileUrl, IsDeleted: false, IsNew: false, name: name, FileID: selItem.Id };
                            filesArry.push(obj);
                        });

                        let editSEWOItem = SEWORes[0];
                        formData.AccidentCauseId = editSEWOItem.AccidentCauseId ?? '';
                        formData.AccidentTypeId = editSEWOItem.AccidentTypeId ?? '';
                        formData.Act = editSEWOItem.Act ?? '';
                        formData.ActionId = editSEWOItem.ActionId ?? '';
                        formData.ActionDescription = editSEWOItem.ActionDescription ?? '';
                        formData.ActionPlan = editSEWOItem.ActionPlan ?? '';
                        formData.BackToWork = editSEWOItem.BackToWork ?? '';
                        formData.BodyPartId = editSEWOItem.BodyPartId ?? '';
                        formData.CloseDate = editSEWOItem.CloseDate ?? '';
                        formData.Comments = editSEWOItem.Comments ?? '';
                        formData.DaysOff = editSEWOItem.DaysOff ?? '';
                        formData.Department = editSEWOItem.Department ?? '';
                        formData.DeptManagerDate = editSEWOItem.DeptManagerDate ?? '';
                        formData.DeptManagerName = editSEWOItem.DeptManagerName ?? '';
                        formData.DeptManagerSignature = editSEWOItem.DeptManagerSignature ?? '';
                        formData.DueDate = editSEWOItem.DueDate ?? '';
                        formData.EmployeeDate = editSEWOItem.EmployeeDate ?? '';
                        formData.EmployeeName = editSEWOItem.EmployeeName ?? '';
                        formData.EmployeeSignature = editSEWOItem.EmployeeSignature ?? '';
                        formData.ExpansionPlanId = editSEWOItem.ExpansionPlanId ?? '';
                        formData.FiveWhy1 = editSEWOItem.FiveWhy1 ?? '';
                        formData.FiveWhy2 = editSEWOItem.FiveWhy2 ?? '';
                        formData.FiveWhy3 = editSEWOItem.FiveWhy3 ?? '';
                        formData.FiveWhy4 = editSEWOItem.FiveWhy4 ?? '';
                        formData.FiveWhy5 = editSEWOItem.FiveWhy5 ?? '';
                        formData.FiveWhyRootCause = editSEWOItem.FiveWhyRootCause ?? '';
                        formData.HHow = editSEWOItem.HHow ?? '';
                        formData.InjuredDate = editSEWOItem.InjuredDate ?? '';
                        formData.InjuredJob = editSEWOItem.InjuredJob ?? '';
                        formData.InjuredSignature = editSEWOItem.InjuredSignature ?? '';
                        formData.InjuredStatement = editSEWOItem.InjuredStatement ?? '';
                        formData.Injury_x0020_Date_x0020_Time = DateUtilities.removeBrowserwrtServer(new Date(editSEWOItem.Injury_x0020_Date_x0020_Time), this.props.spContext.webTimeZoneData).toISOString();
                        formData.InjuryTime = editSEWOItem.InjuryTime ?? '';
                        formData.InjuryTypeId = editSEWOItem.InjuryTypeId ?? '';
                        formData.IsHospitalRefused = editSEWOItem.IsHospitalRefused ?? '';
                        formData.Location = editSEWOItem.Location ?? '';
                        formData.Machine = editSEWOItem.Machine ?? '';
                        formData.MicroRootCauseId = editSEWOItem.MicroRootCauseId ?? '';
                        formData.InjuredName = editSEWOItem.InjuredName ?? '';
                        formData.NameoftheHospital = editSEWOItem.NameoftheHospital ?? '';
                        formData.Notes = editSEWOItem.Notes ?? '';
                        formData.Plant = editSEWOItem.Plant ?? '';
                        formData.PlantMgrDate = editSEWOItem.PlantMgrDate ?? '';
                        formData.PlantMgrName = editSEWOItem.PlantMgrName ?? '';
                        formData.PlantMgrSignature = editSEWOItem.PlantMgrSignature ?? '';
                        formData.PositionType = editSEWOItem.PositionType ?? '';
                        formData.PPEInUseId = editSEWOItem.PPEInUseId ?? '';
                        formData.PPESuppliedId = editSEWOItem.PPESuppliedId ?? '';
                        formData.ReportedBy = editSEWOItem.ReportedBy ?? '';
                        formData.Reported_x0020_Date = editSEWOItem.Reported_x0020_Date ?? '';
                        formData.Responsible = editSEWOItem.Responsible ?? '';
                        formData.RootCauseId = editSEWOItem.RootCauseId ?? '';
                        formData.SafetyMgrDate = editSEWOItem.SafetyMgrDate ?? '';
                        formData.SafetyMgrName = editSEWOItem.SafetyMgrName ?? '';
                        formData.SafetyMgrSignature = editSEWOItem.SafetyMgrSignature ?? '';
                        formData.SecondaryRootCauseId = editSEWOItem.SecondaryRootCauseId ?? '';
                        formData.Sex = editSEWOItem.Sex ?? '';
                        formData.Sketch = editSEWOItem.Sketch ?? '';
                        formData.StatusId = editSEWOItem.StatusId ?? '';
                        formData.SupervisorDate = editSEWOItem.SupervisorDate ?? '';
                        formData.SupervisorName = editSEWOItem.SupervisorName ?? '';
                        formData.SuperVisorSignature = editSEWOItem.SuperVisorSignature ?? '';
                        formData.TeamLeadDate = editSEWOItem.TeamLeadDate ?? '';
                        formData.TeamLeadName = editSEWOItem.TeamLeadName ?? '';
                        formData.TeamLeadSignature = editSEWOItem.TeamLeadSignature ?? '';
                        formData.Title = editSEWOItem.Title ?? '';
                        formData.UsualWorkId = editSEWOItem.UsualWorkId ?? '';
                        formData.WitnessDate = editSEWOItem.WitnessDate ?? '';
                        formData.WitnessSignature = editSEWOItem.WitnessSignature ?? '';
                        formData.WitnessStatement = editSEWOItem.WitnessStatement ?? '';
                        formData.WWhat = editSEWOItem.WWhat ?? '';
                        formData.WWhen = editSEWOItem.WWhen ?? '';
                        formData.WWhere = editSEWOItem.WWhere ?? '';
                        formData.WWhich = editSEWOItem.WWhich ?? '';
                        formData.WWho = editSEWOItem.WWho ?? '';
                        formData.Year = editSEWOItem.Year ?? '';
                        formData.YearMonth = editSEWOItem.YearMonth ?? '';
                        formData.Zone = editSEWOItem.Zone ?? '';
                        //Shifts column if WCM
                        if (this.props.isWCM) {
                            formData.Shift = editSEWOItem.Shifts ?? '';
                        }
                        else {
                            formData.Shift = editSEWOItem.Shift ?? '';
                        }

                        zoneOptions = zoneData.filter((option: any) => (option.Plant && option.Plant.Title == formData.Plant && option.Department && option.Department.Title == editSEWOItem.Department)).map((item: any) => ({ label: item.Title, value: item.Title, id: item.Id }));

                        machineOptions = machineData.filter((option: any) => (option.Plant && option.Plant.Title == formData.Plant && option.Department && option.Department.Title == formData.Department && option.Zone.Title == editSEWOItem.Zone)).map((item: any) => ({ label: item.Title, value: item.Title, id: item.Id }));

                        secondaryRootCausesOptions = secondaryRootCausesData.filter((option: any) => (option.RootCauseId == formData.RootCauseId)).map((item: any) => ({ label: item.Title, value: item.Id }));

                        microRootCausesOptions = microRootCausesData.filter((option: any) => (option.RootCauseId == formData.RootCauseId && option.SecondaryRootCauseId == formData.SecondaryRootCauseId)).map((item: any) => ({ label: item.Title, value: item.Id }));

                        actionsOptions = actionsData.filter((option: any) => (option.RootCauseId == formData.RootCauseId && option.SecondaryRootCauseId == formData.SecondaryRootCauseId)).map((item: any) => ({ label: item.Title, value: item.Id }));

                        //Groups Check
                        let currentUserGroups = this.props.currentUserGroups;
                        if (currentUserGroups.includes("Venture Global Owners") || currentUserGroups.includes("WCM Safety Admin") || currentUserGroups.includes("WCM Merrill Safety Mgt")) {
                            showSubmit = true;
                        }
                        isEditForm = true;
                        let bodyPartValue = bodyPartsData.find((part: any) => (part.value == formData.BodyPartId))?.label;
                        selBodyPart = this.handleBodyPartChange(bodyPartValue);
                    }
                    else {
                        showToast("error", "No SEWO found");
                        this.setState({ Redirect: true, RedirectTo: 'Home' });
                    }
                }
            }
            else {
                showSubmit = true;
                selBodyPart = this.handleBodyPartChange("None");
            }

            this.setState({ formData, plantsData, departmentData, departmentOptions, zoneData, zoneOptions, machineData, machineOptions, shiftData, showSubmit, ItemId: itemId, accidentTypeData, accidentCauseData, injuryTypeData, bodyPartsData, rootCausesData, secondaryRootCausesData, microRootCausesData, actionsData, yesNoData, statusData, secondaryRootCausesOptions, microRootCausesOptions, actionsOptions, selBodyPart, fileArr: filesArry, isEditForm });
        }
        catch (e) {
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
        if (name == "InjuredName") {
            formData.EmployeeName = inputValue;
        }
        else if (name == "IsHospitalRefused") {
            formData.NameoftheHospital = '';
        }

        formData[name] = inputValue;
        this.setState({ formData, });
    }

    private handleDropdownChange = (event: any, actionMeta: any, id: any) => {
        const formData: any = { ...this.state.formData };
        let departmentData = [...this.state.departmentData];
        let departmentOptions: any = [...this.state.departmentOptions];
        let zoneData = [...this.state.zoneData];
        let zoneOptions: any = [...this.state.zoneOptions];
        let machineData = [...this.state.machineData];
        let machineOptions: any = [...this.state.machineOptions];
        let secondaryRootCausesData = [...this.state.secondaryRootCausesData];
        let secondaryRootCausesOptions: any = [...this.state.secondaryRootCausesOptions];
        let microRootCausesData = [...this.state.microRootCausesData];
        let microRootCausesOptions: any = [...this.state.microRootCausesOptions];
        let actionsData = [...this.state.actionsData];
        let actionsOptions: any = [...this.state.actionsOptions];
        const name = actionMeta.name;
        const value = actionMeta.action == "clear" ? '' : event.value;
        formData[name] = value;
        let statusText = this.state.statusText;
        var selBodyPart = [...this.state.selBodyPart];

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
            machineOptions = [];
            zoneOptions = [];

            if (actionMeta.action != "clear") {
                zoneOptions = zoneData.filter((option: any) => (option.Plant && option.Plant.Title == formData.Plant && option.Department && option.Department.Id == event.id)).map((item: any) => ({ label: item.Title, value: item.Title, id: item.Id }));
            }
        }
        else if (name == "Zone") {
            formData.Machine = "";
            machineOptions = [];

            if (actionMeta.action != "clear") {
                machineOptions = machineData.filter((option: any) => (option.Plant && option.Plant.Title == formData.Plant && option.Department && option.Department.Title == formData.Department && option.Zone && option.Zone.Id == event.id)).map((item: any) => ({ label: item.Title, value: item.Title, id: item.Id }));
            }
        }
        else if (name == "RootCauseId") {
            formData.SecondaryRootCauseId = '';
            formData.MicroRootCauseId = '';
            formData.ActionId = '';
            secondaryRootCausesOptions = [];
            microRootCausesOptions = [];
            actionsOptions = [];

            if (actionMeta.action != "clear") {
                secondaryRootCausesOptions = secondaryRootCausesData.filter((option: any) => (option.RootCauseId == formData.RootCauseId)).map((item: any) => ({ label: item.Title, value: item.Id }));
            }
        }
        else if (name == "SecondaryRootCauseId") {
            formData.MicroRootCauseId = '';
            formData.ActionId = '';
            microRootCausesOptions = [];
            actionsOptions = [];

            if (actionMeta.action != "clear") {
                microRootCausesOptions = microRootCausesData.filter((option: any) => (option.RootCauseId == formData.RootCauseId && option.SecondaryRootCauseId == formData.SecondaryRootCauseId)).map((item: any) => ({ label: item.Title, value: item.Id }));

                actionsOptions = actionsData.filter((option: any) => (option.RootCauseId == formData.RootCauseId && option.SecondaryRootCauseId == formData.SecondaryRootCauseId)).map((item: any) => ({ label: item.Title, value: item.Id }));
            }
        }
        else if (name == "StatusId") {
            if (actionMeta.action != "clear" && event.label == "Closed") {
                statusText = event.label;
            }
            else {
                statusText = '';
                formData.CloseDate = '';
            }
        }
        else if (name == "BodyPartId") {
            if (actionMeta.action != "clear") {
                selBodyPart = this.handleBodyPartChange(event.label)
            }
            else {
                selBodyPart = [];
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

        this.setState({ formData, departmentOptions, zoneOptions, machineOptions, secondaryRootCausesOptions, microRootCausesOptions, actionsOptions, statusText, selBodyPart });
    }

    private handleDateChange = (dateValue: any, name: any, divId: any, dateProps: any) => {
        const formData: any = { ...this.state.formData };
        let statusText = this.state.statusText;
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
            if (name == "Injury_x0020_Date_x0020_Time") {
                dateValue = format(dateValue, "MM/dd/yyyy HH:mm");
            }
            else {
                dateValue = format(DateUtilities.addBrowserwrtServer(new Date(DateUtilities.getDateMMDDYYYY(dateValue)), this.props.spContext.webTimeZoneData).toISOString(), "MM/dd/yyyy");
            }
        }
        else {
            dateValue = "";
        }

        if (name == "CloseDate") {

            if (!(["", null, undefined].includes(dateValue))) {
                let statusData = [...this.state.statusData];
                let closedStatusObj: any = statusData.find((status: any) => status.label == "Closed")
                formData.StatusId = closedStatusObj.value;
                statusText = 'Closed';
            }
            else {
                if (statusText == "Closed") {
                    statusText = "";
                    formData.StatusId = '';
                }
            }
        }

        formData[name] = dateValue;

        this.setState({ formData, statusText });
    }

    private handleSubmit = async () => {
        try {
            showLoader();
            var formData: any = { ...this.state.formData };
            var data = {
                plant: { val: formData.Plant, required: true, Name: "Plant", Type: ControlType.reactSelect, Focusid: "ddlPlant" },
                department: { val: formData.Department, required: true, Name: "Department", Type: ControlType.reactSelect, Focusid: "ddlDepartment" },
                zone: { val: formData.Zone, required: true, Name: "Zone", Type: ControlType.reactSelect, Focusid: "ddlZone" },
                machine: { val: formData.Machine, required: true, Name: "Machine", Type: ControlType.reactSelect, Focusid: "ddlMachine" },
                accidentType: { val: formData.AccidentTypeId, required: true, Name: "Accident Type", Type: ControlType.reactSelect, Focusid: "ddlAccidentType" },
                injuredName: { val: formData.InjuredName, required: true, Name: "Name of Injured", Type: ControlType.string, Focusid: this.txtNameofInjured },
                sex: { val: formData.Sex, required: true, Name: "Sex", Type: ControlType.reactSelect, Focusid: "ddlSex" },
                injuryType: { val: formData.InjuryTypeId, required: true, Name: "Injury Type", Type: ControlType.reactSelect, Focusid: "ddlInjuryType" },
                injuryDateTime: { val: formData.Injury_x0020_Date_x0020_Time, required: true, Name: "Injury Date Time", Type: ControlType.date, Focusid: "dtInjuryDateTime" },
                injuryDateTimeToday: { val: formData.Injury_x0020_Date_x0020_Time, required: true, Name: "Injury Date Time", Type: ControlType.lessthanTodayDate, Focusid: "dtInjuryDateTime" },
                reportedDate: { val: formData.Reported_x0020_Date, required: true, Name: "Reported Date", Type: ControlType.date, Focusid: "dtReportedDate" },
                reportedDateToday: { val: formData.Reported_x0020_Date, required: true, Name: "Reported Date", Type: ControlType.lessthanTodayDate, Focusid: "dtReportedDate" },
                nameofHospital: { val: formData.NameoftheHospital, required: (!formData.IsHospitalRefused), Name: "Name of the Clinic/Hospital", Type: ControlType.string, Focusid: this.txtNameoftheHospital },
                bodyPart: { val: formData.BodyPartId, required: true, Name: "Body Part", Type: ControlType.reactSelect, Focusid: "ddlBodyPart" },
                rootCause: { val: formData.RootCauseId, required: true, Name: "Root Cause", Type: ControlType.reactSelect, Focusid: "ddlRootCause" },
                secondaryRootCause: { val: formData.SecondaryRootCauseId, required: true, Name: "Secondary Root Cause", Type: ControlType.reactSelect, Focusid: "ddlSecondaryRootCause" },
                microRootCause: { val: formData.MicroRootCauseId, required: true, Name: "Micro Root Cause", Type: ControlType.reactSelect, Focusid: "ddlMicroRootCause" },
                closeDate: { val: formData.CloseDate, required: (this.state.statusText == "Closed"), Name: "Close Date", Type: ControlType.date, Focusid: "dtCloseDate" },
                closeDateToday: { val: formData.CloseDate, required: formData.CloseDate != "", Name: "Close Date", Type: ControlType.lessthanTodayDate, Focusid: "dtCloseDate" },
                closeDateCompare: { startDate: formData.Injury_x0020_Date_x0020_Time, endDate: formData.CloseDate, required: formData.CloseDate != "", startDateName: "Injured Date", endDateName: "Close Date", Type: ControlType.compareDates, Focusid: "dtCloseDate" },
                supervisorName: { val: formData.SupervisorName, required: true, Name: "Supervisor Name", Type: ControlType.string, Focusid: this.txtSupervisorName }
            }
            let isValid = formValidation.FormValidation(data);
            if (isValid.status) {

                let fileArr = this.state.fileArr;
                isValid = formValidation.FilesValidation(fileArr, false);

                if (isValid.status) {
                    let mmddyyyyDate = format(formData.Injury_x0020_Date_x0020_Time, "MM/dd/yyyy");
                    formData.Year = mmddyyyyDate.split("/")[2];
                    formData.YearMonth = mmddyyyyDate.split("/")[0];
                    formData.Injury_x0020_Date_x0020_Time = DateUtilities.addBrowserwrtServer(new Date(formData.Injury_x0020_Date_x0020_Time), this.props.spContext.webTimeZoneData).toISOString();
                    formData.DaysOff = Number(formData.DaysOff);
                    formData.Sketch = this.sketchRef.current?.getImageData();
                    //To handle optional Date fields
                    this.processOptionalDateFields(formData);
                    //To handle optional LookUp Fields
                    this.processOptionalLookUpFields(formData);
                    await this.InsertOrUpdateData(formData);
                }
                else {
                    showToast("error", isValid.message);
                    hideLoader();
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
            let newFileArry = this.state.fileArr.filter((file: any) => {
                return file.IsNew == true;
            });
            if (itemId > 0) {
                await this.sp.web.lists.getByTitle(this.SEWOList).items.getById(itemId).update(formData).then(async (res: any) => {
                    if (newFileArry.length > 0 || this.state.delfileArr.length > 0) {
                        await this.handleAttachmentUpload(itemId.toString(), newFileArry);
                    }
                    else {
                        let msg = "SEWO updated successfully";
                        this.onSuccess(msg);
                    }
                }, (error) => {
                    console.log(error);
                    this.onError();
                })
            }
            else {
                await this.sp.web.lists.getByTitle(this.SEWOList).items.add(formData).then(async (res: any) => {
                    if (newFileArry.length > 0 || this.state.delfileArr.length > 0) {
                        await this.handleAttachmentUpload(res.Id.toString(), newFileArry);
                    }
                    else {
                        let msg = "SEWO submitted successfully";
                        this.onSuccess(msg);
                    }
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

    private handleAttachmentUpload = async (SEWOId: any, fileArr: any) => {
        try {
            if (this.props.match.params.id > 0 && this.state.delfileArr.length > 0) {
                await this.handleAttachmentDelete(SEWOId, this.state.delfileArr);
                await new Promise(resolve => setTimeout(resolve, 2000));
            }
            for (const file of fileArr) {
                await this.sp.web.lists.getByTitle(this.SEWOList).items.getById(SEWOId).attachmentFiles.add(file.name, file);
            }
            let msg = this.state.ItemId > 0 ? "SEWO updated successfully" : "SEWO submitted successfully";
            this.onSuccess(msg);
        } catch (e) {
            console.log(e);
            this.onError();
        }
    }

    private handleAttachmentDelete = async (SEWOId: any, delFileArr: any) => {
        try {
            for (const file of delFileArr) {
                await this.sp.web.lists.getByTitle(this.SEWOList).items.getById(SEWOId).attachmentFiles.getByName(file.name).delete();
            }
        } catch (e) {
            console.log(e);
            this.onError();
        }
    }

    private processOptionalDateFields = (formData: any) => {
        for (const dateField of this.dateFields) {
            if (formData[dateField] !== "") {
                formData[dateField] = DateUtilities.addBrowserwrtServer(new Date(DateUtilities.getDateMMDDYYYY(formData[dateField])), this.props.spContext.webTimeZoneData).toISOString();
            } else {
                delete formData[dateField];
            }
        }
    };

    private processOptionalLookUpFields = (formData: any) => {
        for (const lookupField of this.optionalLookUpFields) {
            if (formData[lookupField] == "") {
                delete formData[lookupField];
            }
        }

        let selShift = formData.Shift;

        if (this.props.isWCM) {
            delete formData.Shift;
            formData["Shifts"] = selShift;
        }
    };

    private handleCancel = () => {
        this.setState({ Redirect: true, RedirectTo: 'SEWOView', ItemId: 0 });
    }
    private onSuccess = (displayMessage: string) => {
        hideLoader();
        this.setState({ Redirect: true, RedirectTo: 'SEWOView', ItemId: 0 });
        showToast("success", displayMessage);
    }

    private onError = () => {
        showToast("error", ActionStatus.Error);
        hideLoader();
    }

    private handleBodyPartChange = (bodyPart: string) => {
        const selValue = bodyPart;
        let selBodyPart: any = [];
        if (selValue === "None" || selValue === "NA" || selValue === "Other") { selBodyPart = [{ slug: undefined, intensity: 0 }]; return selBodyPart; }
        switch (selValue) {
            case "Back":
                selBodyPart = [
                    { slug: "upper-back", leftSideIntensity: 2, rightSideIntensity: 2 },
                    { slug: "lower-back", leftSideIntensity: 2, rightSideIntensity: 2 },
                    { slug: "trapezius", leftSideIntensity: 2, rightSideIntensity: 2 }
                ];
                break;

            case "Chest":
                selBodyPart = [
                    { slug: "chest", leftSideIntensity: 2, rightSideIntensity: 2 }
                ];
                break;

            case "Elbow/ Forearm":
                selBodyPart = [
                    { slug: "forearm", leftSideIntensity: 2, rightSideIntensity: 2 },
                    // { slug: "triceps", leftSideIntensity: 2, rightSideIntensity: 2 }
                ];
                break;

            case "Eye":
                selBodyPart = [
                    { slug: "head", intensity: 2 }
                ];
                break;

            case "Foot/ Ankle":
                selBodyPart = [
                    { slug: "feet", leftSideIntensity: 2, rightSideIntensity: 2 },
                    { slug: "ankles", leftSideIntensity: 2, rightSideIntensity: 2 }
                ];
                break;

            case "Hand/ Finger/ Wrist":
                selBodyPart = [
                    { slug: "hands", leftSideIntensity: 2, rightSideIntensity: 2 }
                ];
                break;

            case "Head/ Neck":
                selBodyPart = [
                    { slug: "head", intensity: 2 },
                    { slug: "neck", intensity: 2 },
                    { slug: "hair", intensity: 2 }
                ];
                break;

            case "Leg/ Knee":
                selBodyPart = [
                    { slug: "knees", leftSideIntensity: 2, rightSideIntensity: 2 },
                    { slug: "quadriceps", leftSideIntensity: 2, rightSideIntensity: 2 },
                    { slug: "hamstring", leftSideIntensity: 2, rightSideIntensity: 2 },
                    { slug: "calves", leftSideIntensity: 2, rightSideIntensity: 2 },
                    { slug: "tibialis", leftSideIntensity: 2, rightSideIntensity: 2 },
                    { slug: "adductors", leftSideIntensity: 2, rightSideIntensity: 2 },
                ];
                break;

            case "Neck/ Back Upper":
                selBodyPart = [
                    { slug: "neck", intensity: 2 },
                    { slug: "upper-back", leftSideIntensity: 2, rightSideIntensity: 2 },
                    { slug: "trapezius", leftSideIntensity: 2, rightSideIntensity: 2 }
                ];
                break;

            case "Shoulder":
                selBodyPart = [
                    { slug: "deltoids", leftSideIntensity: 2, rightSideIntensity: 2 },
                    // { slug: "trapezius", leftSideIntensity: 2, rightSideIntensity: 2 },
                ];
                break;

            case "Upper Arm":
                selBodyPart = [
                    { slug: "biceps", leftSideIntensity: 2, rightSideIntensity: 2 },
                    { slug: "triceps", leftSideIntensity: 2, rightSideIntensity: 2 }
                ];
                break;

            default:
                [{ slug: undefined, intensity: 0 }];

        }
        return selBodyPart;
    };

    private filesChanged = (selectedFiles: any) => {
        this.setState({ fileArr: selectedFiles[0], delfileArr: selectedFiles[1] });
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
                                <div className="form-title">{" SEWO Form " + (this.state.isEditForm ? (" - " + this.state.ItemId) : "")} </div>
                                <span className="span-mandatory-text"> <span className="text-danger">* </span> are mandatory fields</span>
                            </div>

                            {/* PLAN */}
                            <div className="form-border-box p-2 mx-3 my-2">
                                <h6 className="greenbg"><FontAwesomeIcon icon={faWarning} /> PLAN</h6>
                                <div className="row">
                                    {/* Plant */}
                                    <div className="col-md-3">
                                        <div className="custom-dropdown active " id="divPlant" title={this.state.formData.Plant}>
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
                                    {/* Accident Type */}
                                    <div className="col-md-3">
                                        <div className="custom-dropdown" id="divAccidentType" title={(this.state.accidentTypeData.find((i: { label: string; value: any }) => i.value == this.state.formData.AccidentTypeId) as { label: string; value: any } | undefined)?.label}>
                                            <SearchableDropdown
                                                label={"Accident Type"}
                                                Title={"Accident Type"}
                                                name={"AccidentTypeId"}
                                                id="ddlAccidentType"
                                                placeholderText={""}
                                                className={""}
                                                selectedValue={this.state.formData.AccidentTypeId}
                                                OptionsList={this.state.accidentTypeData}
                                                OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta, "divAccidentType") }}
                                                isRequired={true}
                                                disabled={this.state.isInputDisabled}
                                                noOptionsMessage="No Accident Types available"
                                            />
                                        </div>
                                    </div>
                                    {/* Accident Cause */}
                                    <div className="col-md-3">
                                        <div className="custom-dropdown" id="divAccidentCause" title={(this.state.accidentCauseData.find((i: { label: string; value: any }) => i.value == this.state.formData.AccidentCauseId) as { label: string; value: any } | undefined)?.label}>
                                            <SearchableDropdown
                                                label={"Accident Cause"}
                                                Title={"Accident Cause"}
                                                name={"AccidentCauseId"}
                                                id="ddlAccidentCause"
                                                placeholderText={""}
                                                className={""}
                                                selectedValue={this.state.formData.AccidentCauseId}
                                                OptionsList={this.state.accidentCauseData}
                                                OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta, "divAccidentCause") }}
                                                isRequired={false}
                                                disabled={this.state.isInputDisabled}
                                                noOptionsMessage="No Accident Causes available"
                                            />
                                        </div>
                                    </div>
                                    {/* Name of Injured */}
                                    <div className="col-md-3">
                                        <div className="light-text">
                                            <label className=" col-form-label" htmlFor="txtNameofInjured">Name of Injured  <span className="mandatoryhastrick"> *</span></label>
                                            <input className="form-control" placeholder="" name="InjuredName" type="text" id="txtNameofInjured" ref={this.txtNameofInjured} value={this.state.formData.InjuredName} title={this.state.formData.InjuredName} onChange={this.handleChange} disabled={this.state.isInputDisabled} />
                                        </div>
                                    </div>
                                    {/* Sex */}
                                    <div className="col-md-3">
                                        <div className="custom-dropdown" id="divSex" title={this.state.formData.Sex}>
                                            <SearchableDropdown
                                                label={"Sex"}
                                                Title={"Sex"}
                                                name={"Sex"}
                                                id="ddlSex"
                                                placeholderText={""}
                                                className={""}
                                                selectedValue={this.state.formData.Sex}
                                                OptionsList={[{ label: "Female", value: "Female" }, { label: "Male", value: "Male" }]}
                                                OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta, "divSex") }}
                                                isRequired={true}
                                                disabled={this.state.isInputDisabled}
                                                noOptionsMessage="No Sex types Available"
                                            />
                                        </div>
                                    </div>
                                    {/* Injured Job */}
                                    <div className="col-md-3">
                                        <div className="light-text">
                                            <label className=" col-form-label" htmlFor="txtInjuredJob">Injured Job </label>
                                            <input className="form-control" placeholder="" name="InjuredJob" type="text" id="txtInjuredJob" ref={this.txtInjuredJob} value={this.state.formData.InjuredJob} title={this.state.formData.InjuredJob} onChange={this.handleChange} disabled={this.state.isInputDisabled} />
                                        </div>
                                    </div>
                                    {/* Reported By */}
                                    <div className="col-md-3">
                                        <div className="light-text">
                                            <label className=" col-form-label" htmlFor="txtReportedBy">Reported By </label>
                                            <input className="form-control" placeholder="" name="ReportedBy" type="text" id="txtReportedBy" ref={this.txtReportedBy} value={this.state.formData.ReportedBy} title={this.state.formData.ReportedBy} onChange={this.handleChange} disabled={this.state.isInputDisabled} />
                                        </div>
                                    </div>
                                    {/* Injury Type */}
                                    <div className="col-md-3">
                                        <div className="custom-dropdown" id="divInjuryType" title={(this.state.injuryTypeData.find((i: { label: string; value: any }) => i.value == this.state.formData.InjuryTypeId) as { label: string; value: any } | undefined)?.label}>
                                            <SearchableDropdown
                                                label={"Injury Type"}
                                                Title={"Injury Type"}
                                                name={"InjuryTypeId"}
                                                id="ddlInjuryType"
                                                placeholderText={""}
                                                className={""}
                                                selectedValue={this.state.formData.InjuryTypeId}
                                                OptionsList={this.state.injuryTypeData}
                                                OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta, "divInjuryType") }}
                                                isRequired={true}
                                                disabled={this.state.isInputDisabled}
                                                noOptionsMessage="No Injury Types available"
                                            />
                                        </div>
                                    </div>
                                    {/* Injury Date Time*/}
                                    <div className="col-md-3">
                                        <div className="light-text">
                                            <label className="label-datePicker" htmlFor="dtInjuryDateTime"> Injury Date Time  <span className="mandatoryhastrick"> *</span></label>
                                            <div className="custom-datepicker" id="divInjuryDateTime">
                                                <DatePickercontrol placeholder="MM/DD/YYYY HH:MM" selectedDate={this.state.formData.Injury_x0020_Date_x0020_Time} title={this.state.formData.Injury_x0020_Date_x0020_Time} id='dtInjuryDateTime' isDisabled={this.state.isInputDisabled} startDate={undefined} endDate={new Date()} name="Injury_x0020_Date_x0020_Time" onDatechange={(dateProps: any) => this.handleDateChange(dateProps[0], dateProps[2], "divInjuryDateTime", dateProps)} highlightDate={new Date()} showIcon showTime={true} TimeFormat="MM/dd/yyyy hh:mm aa" />
                                            </div>
                                        </div>
                                    </div>
                                    {/* Usual Work & Shift */}
                                    <div className="col-md-3 row">
                                        {/* Usual Work */}
                                        <div className="col-md-6" id="divUsualWork">
                                            <div className="custom-dropdown" title={(this.state.yesNoData.find((i: { label: string; value: any }) => i.value == this.state.formData.UsualWorkId) as { label: string; value: any } | undefined)?.label}>
                                                <SearchableDropdown
                                                    label={"Usual Work"}
                                                    Title={"Usual Work"}
                                                    name={"UsualWorkId"}
                                                    id="ddlUsualWork"
                                                    placeholderText={""}
                                                    className={""}
                                                    selectedValue={this.state.formData.UsualWorkId}
                                                    OptionsList={this.state.yesNoData}
                                                    OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta, "divUsualWork") }}
                                                    isRequired={false}
                                                    disabled={this.state.isInputDisabled}
                                                    noOptionsMessage="No Usual Works available"
                                                />
                                            </div>
                                        </div>

                                        {/* Shift */}
                                        <div className="col-md-6" id="divShift">
                                            <div className="custom-dropdown" title={this.state.formData.Shift}>
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
                                                    isRequired={false}
                                                    disabled={this.state.isInputDisabled}
                                                    noOptionsMessage="No Shifts available"
                                                />
                                            </div>
                                        </div>
                                    </div>
                                    {/* Reported Date */}
                                    <div className="col-md-3">
                                        <div className="light-text">
                                            <label className="label-datePicker" htmlFor="dtReportedDate"> Reported Date  <span className="mandatoryhastrick"> *</span></label>
                                            <div className="custom-datepicker" id="divReportedDate">
                                                <DatePickercontrol placeholder="MM/DD/YYYY" selectedDate={this.state.formData.Reported_x0020_Date} title={this.state.formData.Reported_x0020_Date} id='dtReportedDate' isDisabled={this.state.isInputDisabled} startDate={undefined} endDate={new Date()} name="Reported_x0020_Date" onDatechange={(dateProps: any) => this.handleDateChange(dateProps[0], dateProps[2], "divReportedDate", dateProps)} highlightDate={new Date()} showIcon />
                                            </div>
                                        </div>
                                    </div>
                                    {/* Position Type */}
                                    <div className="col-md-3">
                                        <div className="custom-dropdown" id="divPositionType" title={this.state.formData.PositionType}>
                                            <SearchableDropdown
                                                label={"Position Type"}
                                                Title={"Position Type"}
                                                name={"PositionType"}
                                                id="ddlPositionType"
                                                placeholderText={""}
                                                className={""}
                                                selectedValue={this.state.formData.PositionType}
                                                OptionsList={[{ label: "Permanent", value: "Permanent" }, { label: "Temporory", value: "Temporory" }]}
                                                OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta, "divPositionType") }}
                                                isRequired={false}
                                                disabled={this.state.isInputDisabled}
                                                noOptionsMessage="No Position Types Available"
                                            />
                                        </div>
                                    </div>
                                    {/* PPE in Use */}
                                    <div className="col-md-3">
                                        <div className="custom-dropdown" id="divPPEinUse" title={(this.state.yesNoData.find((i: { label: string; value: any }) => i.value == this.state.formData.PPEInUseId) as { label: string; value: any } | undefined)?.label}>
                                            <SearchableDropdown
                                                label={"PPE in Use"}
                                                Title={"PPE in Use"}
                                                name={"PPEInUseId"}
                                                id="ddlPPEinUse"
                                                placeholderText={""}
                                                className={""}
                                                selectedValue={this.state.formData.PPEInUseId}
                                                OptionsList={this.state.yesNoData}
                                                OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta, "divPPEinUse") }}
                                                isRequired={false}
                                                disabled={this.state.isInputDisabled}
                                                noOptionsMessage="No PPE in Use Available"
                                            />
                                        </div>
                                    </div>
                                    {/* PPE Supplied */}
                                    <div className="col-md-3">
                                        <div className="custom-dropdown" id="divPPESupplied" title={(this.state.yesNoData.find((i: { label: string; value: any }) => i.value == this.state.formData.PPESuppliedId) as { label: string; value: any } | undefined)?.label}>
                                            <SearchableDropdown
                                                label={"PPE Supplied"}
                                                Title={"PPE Supplied"}
                                                name={"PPESuppliedId"}
                                                id="ddlPPESupplied"
                                                placeholderText={""}
                                                className={""}
                                                selectedValue={this.state.formData.PPESuppliedId}
                                                OptionsList={this.state.yesNoData}
                                                OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta, "divPPESupplied") }}
                                                isRequired={false}
                                                disabled={this.state.isInputDisabled}
                                                noOptionsMessage="No PPE Supplied Available"
                                            />
                                        </div>
                                    </div>
                                    {/* Is Hopital/Clinic Refused */}
                                    <div className="col-md-3" style={{ marginTop: "10px;" }} title="Is Hospital/Clinic Refused">
                                        <input className="" placeholder="Is Hospital/Clinic Refused" name="IsHospitalRefused" type="checkbox" id="rdIsHospitalRefused" ref={this.rdIsHospitalRefused} checked={this.state.formData.IsHospitalRefused} onChange={this.handleChange} disabled={this.state.isInputDisabled} />
                                        <label className="ps-1 col-form-label" htmlFor="rdIsHospitalRefused"> Is Hospital/Clinic Refused </label>
                                    </div>
                                    {/* Name of the Clinic/Hospital */}
                                    {!this.state.formData.IsHospitalRefused && <div className="col-md-6">
                                        <div className="light-text">
                                            <label className=" col-form-label" htmlFor="txtNameoftheHospital">Name of the Clinic/Hospital <span className="mandatoryhastrick"> *</span></label>
                                            <input className="form-control" placeholder="Name of the Clinic/Hospital" name="NameoftheHospital" type="text" id="txtNameoftheHospital" ref={this.txtNameoftheHospital} value={this.state.formData.NameoftheHospital} title={this.state.formData.NameoftheHospital} onChange={this.handleChange} disabled={this.state.isInputDisabled} />
                                        </div>
                                    </div>}
                                </div>
                            </div>

                            {/* 5W+1H Analysis*/}
                            <div className="form-border-box p-2 mx-3 mt-2">
                                <div className="row mt-2">
                                    {/* 5W+1H Analysis */}
                                    <h6 className="greenbg"><FontAwesomeIcon icon={faSearch} /> 5W+1H Analysis</h6>
                                    {/* What */}
                                    <div className="col-md-6">
                                        <div className="light-text">
                                            <label className=" col-form-label" htmlFor="txtWhat">What </label>
                                            <input className="form-control" placeholder="WWhat" name="WWhat" type="text" id="txtWhat" ref={this.txtWhat} value={this.state.formData.WWhat} title={this.state.formData.WWhat} onChange={this.handleChange} disabled={this.state.isInputDisabled} />
                                        </div>
                                    </div>
                                    {/* When */}
                                    <div className="col-md-6">
                                        <div className="light-text">
                                            <label className=" col-form-label" htmlFor="txtWhen">When </label>
                                            <input className="form-control" placeholder="WWhen" name="WWhen" type="text" id="txtWhen" ref={this.txtWhen} value={this.state.formData.WWhen} title={this.state.formData.WWhen} onChange={this.handleChange} disabled={this.state.isInputDisabled} />
                                        </div>
                                    </div>
                                    {/* Where */}
                                    <div className="col-md-6">
                                        <div className="light-text">
                                            <label className=" col-form-label" htmlFor="txtWhere">Where </label>
                                            <input className="form-control" placeholder="WWhere" name="WWhere" type="text" id="txtWhere" ref={this.txtWhere} value={this.state.formData.WWhere} title={this.state.formData.WWhere} onChange={this.handleChange} disabled={this.state.isInputDisabled} />
                                        </div>
                                    </div>
                                    {/* Who */}
                                    <div className="col-md-6">
                                        <div className="light-text">
                                            <label className=" col-form-label" htmlFor="txtWho">Who </label>
                                            <input className="form-control" placeholder="WWho" name="WWho" type="text" id="txtWho" ref={this.txtWho} value={this.state.formData.WWho} title={this.state.formData.WWho} onChange={this.handleChange} disabled={this.state.isInputDisabled} />
                                        </div>
                                    </div>
                                    {/* Which */}
                                    <div className="col-md-6">
                                        <div className="light-text">
                                            <label className=" col-form-label" htmlFor="txtWhich">Which </label>
                                            <input className="form-control" placeholder="WWhich" name="WWhich" type="text" id="txtWhich" ref={this.txtWhich} value={this.state.formData.WWhich} title={this.state.formData.WWhich} onChange={this.handleChange} disabled={this.state.isInputDisabled} />
                                        </div>
                                    </div>
                                    {/* How */}
                                    <div className="col-md-6">
                                        <div className="light-text">
                                            <label className=" col-form-label" htmlFor="txtHow">How </label>
                                            <input className="form-control" placeholder="HHow" name="HHow" type="text" id="txtHow" ref={this.txtHow} value={this.state.formData.HHow} title={this.state.formData.HHow} onChange={this.handleChange} disabled={this.state.isInputDisabled} />
                                        </div>
                                    </div>
                                </div>
                            </div>
                            {/* BODY CHART ,SKETCH*/}
                            <div className="row">
                                {/* BODY CHART */}
                                <div className="col-md-6">
                                    <div className="form-border-box p-2 mx-1 mt-2">
                                        <h6 className="greenbg"><FontAwesomeIcon icon={faUser} /> BODY CHART</h6>
                                        {/* Body Part */}
                                        <div className="col-md-12">
                                            <div className="custom-dropdown" id="divBodyPart" title={(this.state.bodyPartsData.find((i: { label: string; value: any }) => i.value == this.state.formData.BodyPartId) as { label: string; value: any } | undefined)?.label}>
                                                <SearchableDropdown
                                                    label={"Body Part"}
                                                    Title={"Body Part"}
                                                    name={"BodyPartId"}
                                                    id="ddlBodyPart"
                                                    placeholderText={""}
                                                    className={""}
                                                    selectedValue={this.state.formData.BodyPartId}
                                                    OptionsList={this.state.bodyPartsData}
                                                    OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta, "divBodyPart") }}
                                                    isRequired={true}
                                                    disabled={this.state.isInputDisabled}
                                                    noOptionsMessage="No Body Parts available"
                                                />
                                            </div>
                                            <BodyPart selectedBodyPart={this.state.selBodyPart} />
                                        </div>
                                    </div>
                                </div>
                                {/* SKETCH */}
                                <div className="col-md-6 ps-0">
                                    <div className="form-border-box p-2 mx-1 mt-2">
                                        <h6 className="greenbg"><FontAwesomeIcon icon={faPencil} /> SKETCH</h6>
                                        <Sketch ref={this.sketchRef} initialImage={this.state.formData.Sketch} />
                                        <FileUpload ismultiAllowed={true} onFileChanges={this.filesChanged} isFileCloseShow={!this.state.isFileCloseShow} files={[this.state.fileArr, this.state.delfileArr]} isRequired={false} disabled={!this.state.isFileCloseShow} />
                                    </div>
                                </div>
                            </div>
                            {/* CORRECTIVE ACTION */}
                            <div className="form-border-box p-2 mx-3 mt-2">
                                <h6 className="greenbg"><FontAwesomeIcon icon={faFileSignature} /> CORRECTIVE ACTION</h6>
                                {/* Action Description */}
                                {/* <div className="col-md-12"> */}
                                <div className="light-text" >
                                    {/* <label className=" col-form-label" htmlFor="txtActionDescription">Action Description </label> */}
                                    <textarea className="form-control" rows={3} id="txtActionDescription" name="ActionDescription" ref={this.txtActionDescription} placeholder="Action Description" value={this.state.formData.ActionDescription} onChange={this.handleChange} disabled={this.state.isInputDisabled} title="Action Description" ></textarea>
                                </div>
                            </div>
                            {/* Analysis Root Cause: write "5 Why's" for the most probable cause */}
                            <div className="form-border-box p-2 mx-3 mt-2">
                                <h6 className="greenbg"><FontAwesomeIcon icon={faChartLine} /> Analysis Root Cause: write "5 Why's" for the most probable cause</h6>
                                {/* FiveWhy1 */}
                                <div className="row g-0 insider-m-0">
                                    <div className="col-lg-1 col-md-1 col-sm-1 col-xs-1"><div className="div-root-cause">1</div></div>
                                    <div className="col-lg-11 col-md-11 col-sm-11 col-xs-11">
                                        <div className="col-md-12">
                                            <div className="light-text">
                                                <input className="form-control" placeholder="FiveWhy1" name="FiveWhy1" type="text" id="txtFiveWhy1" ref={this.txtFiveWhy1} value={this.state.formData.FiveWhy1} title={this.state.formData.FiveWhy1} onChange={this.handleChange} disabled={this.state.isInputDisabled} />
                                            </div>
                                        </div>
                                    </div>
                                    {/* FiveWhy2 */}
                                    <div className="col-lg-1 col-md-1 col-sm-1 col-xs-1"><div className="div-root-cause">..2</div></div>
                                    <div className="col-lg-11 col-md-11 col-sm-11 col-xs-11">
                                        <div className="col-md-12">
                                            <div className="light-text">
                                                <input className="form-control" placeholder="FiveWhy2" name="FiveWhy2" type="text" id="txtFiveWhy2" ref={this.txtFiveWhy2} value={this.state.formData.FiveWhy2} title={this.state.formData.FiveWhy2} onChange={this.handleChange} disabled={this.state.isInputDisabled} />
                                            </div>
                                        </div>
                                    </div>
                                    {/* FiveWhy3 */}
                                    <div className="col-lg-1 col-md-1 col-sm-1 col-xs-1"><div className="div-root-cause">....3</div></div>
                                    <div className="col-lg-11 col-md-11 col-sm-11 col-xs-11">
                                        <div className="col-md-12">
                                            <div className="light-text">
                                                <input className="form-control" placeholder="FiveWhy3" name="FiveWhy3" type="text" id="txtFiveWhy3" ref={this.txtFiveWhy3} value={this.state.formData.FiveWhy3} title={this.state.formData.FiveWhy3} onChange={this.handleChange} disabled={this.state.isInputDisabled} />
                                            </div>
                                        </div>
                                    </div>
                                    {/* FiveWhy4 */}
                                    <div className="col-lg-1 col-md-1 col-sm-1 col-xs-1"><div className="div-root-cause">......4</div></div>
                                    <div className="col-lg-11 col-md-11 col-sm-11 col-xs-11">
                                        <div className="col-md-12">
                                            <div className="light-text">
                                                <input className="form-control" placeholder="FiveWhy4" name="FiveWhy4" type="text" id="txtFiveWhy4" ref={this.txtFiveWhy4} value={this.state.formData.FiveWhy4} title={this.state.formData.FiveWhy4} onChange={this.handleChange} disabled={this.state.isInputDisabled} />
                                            </div>
                                        </div>
                                    </div>
                                    {/* FiveWhy5 */}
                                    <div className="col-lg-1 col-md-1 col-sm-1 col-xs-1"><div className="div-root-cause">........5</div></div>
                                    <div className="col-lg-11 col-md-11 col-sm-11 col-xs-11">
                                        <div className="col-md-12">
                                            <div className="light-text">
                                                <input className="form-control" placeholder="FiveWhy5" name="FiveWhy5" type="text" id="txtFiveWhy5" ref={this.txtFiveWhy5} value={this.state.formData.FiveWhy5} title={this.state.formData.FiveWhy5} onChange={this.handleChange} disabled={this.state.isInputDisabled} />
                                            </div>
                                        </div>
                                    </div>
                                    {/* RootCause */}
                                    <div className="col-lg-1 col-md-1 col-sm-1 col-xs-1"><div className="div-root-cause">Root Cause</div></div>
                                    <div className="col-lg-11 col-md-11 col-sm-11 col-xs-11">
                                        <div className="col-md-12">
                                            <div className="light-text">
                                                <input className="form-control" placeholder="FiveWhyRootCause" name="FiveWhyRootCause" type="text" id="FiveWhyRootCause" ref={this.txtFiveWhyRootCause} value={this.state.formData.FiveWhyRootCause} title={this.state.formData.FiveWhyRootCause} onChange={this.handleChange} disabled={this.state.isInputDisabled} />
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>
                            {/* Categorize Root Cause */}
                            <div className="form-border-box p-2 mx-3 mt-2">
                                <h6 className="greenbg"><FontAwesomeIcon icon={faChartArea} /> Categorize Root Cause</h6>
                                <div className="row">
                                    {/* Root Cause */}
                                    <div className="col-md-3">
                                        <div className="custom-dropdown" id="divRootCause" title={(this.state.rootCausesData.find((i: { label: string; value: any }) => i.value == this.state.formData.RootCauseId) as { label: string; value: any } | undefined)?.label}>
                                            <SearchableDropdown
                                                label={"Root Cause"}
                                                Title={"Root Cause"}
                                                name={"RootCauseId"}
                                                id="ddlRootCause"
                                                placeholderText={""}
                                                className={""}
                                                selectedValue={this.state.formData.RootCauseId}
                                                OptionsList={this.state.rootCausesData}
                                                OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta, "divRootCause") }}
                                                isRequired={true}
                                                disabled={this.state.isInputDisabled}
                                                noOptionsMessage="No Root Causes available"
                                            />
                                        </div>
                                    </div>
                                    {/* Secondary Root Cause */}
                                    <div className="col-md-3">
                                        <div className="custom-dropdown" id="divSecondaryRootCause" title={(this.state.secondaryRootCausesOptions.find((i: { label: string; value: any }) => i.value == this.state.formData.SecondaryRootCauseId) as { label: string; value: any } | undefined)?.label}>
                                            <SearchableDropdown
                                                label={"Secondary Root Cause"}
                                                Title={"Secondary Root Cause"}
                                                name={"SecondaryRootCauseId"}
                                                id="ddlSecondaryRootCause"
                                                placeholderText={""}
                                                className={""}
                                                selectedValue={this.state.formData.SecondaryRootCauseId}
                                                OptionsList={this.state.secondaryRootCausesOptions}
                                                OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta, "divSecondaryRootCause") }}
                                                isRequired={true}
                                                disabled={this.state.isInputDisabled}
                                                noOptionsMessage="No Secondary Root Causes available"
                                            />
                                        </div>
                                    </div>
                                    {/* Micro Root Cause */}
                                    <div className="col-md-3">
                                        <div className="custom-dropdown" id="divMicroRootCause" title={(this.state.microRootCausesOptions.find((i: { label: string; value: any }) => i.value == this.state.formData.MicroRootCauseId) as { label: string; value: any } | undefined)?.label}>
                                            <SearchableDropdown
                                                label={"Micro Root Cause"}
                                                Title={"Micro Root Cause"}
                                                name={"MicroRootCauseId"}
                                                id="ddlMicroRootCause"
                                                placeholderText={""}
                                                className={""}
                                                selectedValue={this.state.formData.MicroRootCauseId}
                                                OptionsList={this.state.microRootCausesOptions}
                                                OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta, "divMicroRootCause") }}
                                                isRequired={true}
                                                disabled={this.state.isInputDisabled}
                                                noOptionsMessage="No Micro Root Causes available"
                                            />
                                        </div>
                                    </div>
                                    {/* Actions */}
                                    <div className="col-md-3">
                                        <div className="custom-dropdown" id="divAction" title={(this.state.actionsOptions.find((i: { label: string; value: any }) => i.value == this.state.formData.ActionId) as { label: string; value: any } | undefined)?.label}>
                                            <SearchableDropdown
                                                label={"Action"}
                                                Title={"Action"}
                                                name={"ActionId"}
                                                id="ddlAction"
                                                placeholderText={""}
                                                className={""}
                                                selectedValue={this.state.formData.ActionId}
                                                OptionsList={this.state.actionsOptions}
                                                OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta, "divAction") }}
                                                isRequired={false}
                                                disabled={this.state.isInputDisabled}
                                                noOptionsMessage="No Actions available"
                                            />
                                        </div>
                                    </div>
                                </div>
                            </div>
                            {/* Do */}
                            <div className="form-border-box p-2 mx-3 mt-2">
                                <h6 className="greenbg"><FontAwesomeIcon icon={faCheckCircle} /> Do</h6>
                                {/* Action Plan */}
                                <div className="row">
                                    <div className="col-md-6">
                                        <div className="light-text" >
                                            <label className=" col-form-label" htmlFor="txtActionPlan">Action Plan </label>
                                            <textarea className="form-control bs-textarea" rows={3} id="txtActionPlan" name="ActionPlan" ref={this.txtActionPlan} placeholder="" value={this.state.formData.ActionPlan} title={this.state.formData.ActionPlan} onChange={this.handleChange} disabled={this.state.isInputDisabled}></textarea>
                                        </div>
                                        <div className="light-text" >
                                            <label className=" col-form-label" htmlFor="txtResponsible">Responsible </label>
                                            <textarea className="form-control bs-textarea" rows={3} id="txtResponsible" name="Responsible" ref={this.txtResponsible} placeholder="" value={this.state.formData.Responsible} title={this.state.formData.Responsible} onChange={this.handleChange} disabled={this.state.isInputDisabled} ></textarea>
                                        </div>
                                    </div>
                                    {/* Responsible, Due Date, Close Date */}
                                    <div className="col-md-6">
                                        {/* Responsible */}
                                        {/* Due Date */}
                                        <div className="">
                                            <div className="light-text">
                                                <label className="label-datePicker" htmlFor="dtDueDate"> Due Date </label>
                                                <div className="custom-datepicker" id="divDueDate">
                                                    <DatePickercontrol placeholder="MM/DD/YYYY" selectedDate={this.state.formData.DueDate} title={this.state.formData.DueDate} id='dtDueDate' isDisabled={this.state.isInputDisabled} startDate={undefined} endDate={undefined} name="DueDate" onDatechange={(dateProps: any) => this.handleDateChange(dateProps[0], dateProps[2], "divDueDate", dateProps)} highlightDate={new Date()} showIcon />
                                                </div>
                                            </div>
                                        </div>
                                        {/* Close Date */}
                                        <div className="" >
                                            <div className="light-text">
                                                <label className="label-datePicker" htmlFor="dtCloseDate"> Close Date {this.state.statusText == "Closed" && <span className="mandatoryhastrick"> *</span>}</label>
                                                <div className="custom-datepicker" id="divCloseDate">
                                                    <DatePickercontrol placeholder="MM/DD/YYYY" selectedDate={this.state.formData.CloseDate} title={this.state.formData.CloseDate} id='dtCloseDate' isDisabled={this.state.isInputDisabled} startDate={undefined} endDate={new Date()} name="CloseDate" onDatechange={(dateProps: any) => this.handleDateChange(dateProps[0], dateProps[2], "divCloseDate", dateProps)} highlightDate={new Date()} showIcon />
                                                </div>
                                            </div>
                                            <div className="light-text" >
                                                <label className="col-form-label" htmlFor="txtNotes">Notes </label>
                                                <textarea className="form-control bs-textarea" rows={3} id="txtNotes" name="Notes" ref={this.txtNotes} placeholder="" value={this.state.formData.Notes} title={this.state.formData.Notes} onChange={this.handleChange} disabled={this.state.isInputDisabled} ></textarea>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                {/* Notes */}
                            </div>
                            {/* Check, ACT/ */}
                            <div className="row g-0 mt-2">
                                {/* Check*/}
                                <div className="col-md-6">
                                    <div className="form-border-box p-2 mx-1">
                                        {/* Check */}
                                        <h6 className="redbg"><FontAwesomeIcon icon={faCheck} /> Check</h6>
                                        {/* Comments */}
                                        <div className="light-text" >
                                            <label className="col-form-label" htmlFor="txtComments">Comments </label>
                                            <textarea className="form-control bs-textarea" rows={3} id="txtComments" name="Comments" ref={this.txtComments} placeholder="" value={this.state.formData.Comments} title={this.state.formData.Comments} onChange={this.handleChange} disabled={this.state.isInputDisabled}></textarea>
                                        </div>
                                    </div>
                                </div>
                                {/* ACT */}
                                <div className="col-md-6">
                                    <div className="form-border-box p-2 mx-1">
                                        <h6 className="yellowbg"><FontAwesomeIcon icon={faCheckDouble} /> ACT</h6>
                                        {/* Expansion Plan */}
                                        <div className="row mt-3">
                                            <div className="col-md-4">
                                                <div className="custom-dropdown mt-0" id="divExpansionPlan" title={(this.state.yesNoData.find((i: { label: string; value: any }) => i.value == this.state.formData.ExpansionPlanId) as { label: string; value: any } | undefined)?.label}>
                                                    <SearchableDropdown
                                                        label={"Expansion Plan"}
                                                        Title={"Expansion Plan"}
                                                        name={"ExpansionPlanId"}
                                                        id="ddlExpansionPlan"
                                                        placeholderText={""}
                                                        className={""}
                                                        selectedValue={this.state.formData.ExpansionPlanId}
                                                        OptionsList={this.state.yesNoData}
                                                        OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta, "divExpansionPlan") }}
                                                        isRequired={false}
                                                        disabled={this.state.isInputDisabled}
                                                        noOptionsMessage="No Expansion Plan Available"
                                                    />
                                                </div>
                                            </div>
                                            {/* Location */}
                                            <div className="col-md-4 light-text" >
                                                <label className="col-form-label" htmlFor="txtLocation">Location </label>
                                                <textarea className="form-control bs-textarea" rows={3} id="txtLocation" name="Location" ref={this.txtLocation} placeholder="" value={this.state.formData.Location} title={this.state.formData.Location} onChange={this.handleChange} disabled={this.state.isInputDisabled} ></textarea>
                                            </div>
                                            {/* Act */}
                                            <div className="col-md-4 light-text">
                                                <label className="col-form-label" htmlFor="txtAct">Act </label>
                                                <textarea className="form-control bs-textarea" rows={3} id="txtAct" name="Act" ref={this.txtAct} placeholder="" value={this.state.formData.Act} title={this.state.formData.Act} onChange={this.handleChange} disabled={this.state.isInputDisabled}></textarea>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>
                            <div className="form-border-box p-2 mx-3 mt-2">
                                {/* Status */}
                                <div className="row">
                                    <div className="col-md-4">
                                        <div className="custom-dropdown" id="divStatus" title={(this.state.statusData.find((i: { label: string; value: any }) => i.value == this.state.formData.StatusId) as { label: string; value: any } | undefined)?.label}>
                                            <SearchableDropdown
                                                label={"Status"}
                                                Title={"Status"}
                                                name={"StatusId"}
                                                id="ddlStatus"
                                                placeholderText={""}
                                                className={""}
                                                selectedValue={this.state.formData.StatusId}
                                                OptionsList={this.state.statusData}
                                                OnChange={(selectedOption: any, actionMeta: any) => { this.handleDropdownChange(selectedOption, actionMeta, "divStatus") }}
                                                isRequired={false}
                                                disabled={this.state.isInputDisabled}
                                                noOptionsMessage="No Status Available"
                                            />
                                        </div>
                                    </div>
                                    {/* Days Off */}
                                    <div className="col-md-4">
                                        <div className="light-text">
                                            <label className="col-form-label" htmlFor="txtDaysOff">Days Off </label>
                                            <input className="form-control onlyNum" placeholder="" name="DaysOff" type="text" id="txtDaysOff" ref={this.txtDaysOff} value={this.state.formData.DaysOff} title={this.state.formData.DaysOff} onChange={this.handleChange} disabled={this.state.isInputDisabled} maxLength={3} />
                                        </div>
                                    </div>
                                    {/* Back To Work */}
                                    <div className="col-md-4">
                                        <div className="light-text">
                                            <label className="label-datePicker" htmlFor="dtBackToWork"> Back To Work </label>
                                            <div className="custom-datepicker" id="divBackToWork">
                                                <DatePickercontrol placeholder="MM/DD/YYYY" selectedDate={this.state.formData.BackToWork} title={this.state.formData.BackToWork} id='dtBackToWork' isDisabled={this.state.isInputDisabled} startDate={undefined} endDate={undefined} name="BackToWork" onDatechange={(dateProps: any) => this.handleDateChange(dateProps[0], dateProps[2], "divBackToWork", dateProps)} highlightDate={new Date()} showIcon />
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                <table className="table table-condensed table-bordered border col-xs-12 col-md-12 col-lg-12 col-sm-12 mt-2">
                                    <tr>
                                        <th scope="col">&nbsp;</th>
                                        <th scope="col">Employee</th>
                                        <th scope="col">Team Leader</th>
                                        <th scope="col">Supervisor <span className="mandatoryhastrick"> *</span></th>
                                        <th scope="col">Depart Mgr.</th>
                                        <th scope="col">Safety Mgr.</th>
                                        <th scope="col">Plant Mgr.</th>
                                    </tr>
                                    {/* Name */}
                                    <tr>
                                        <th scope="row">Name</th>
                                        <td>
                                            <div className="light-text">
                                                <input className="form-control" placeholder="Employee Name" name="EmployeeName" type="text" id="txtEmployeeName" ref={this.txtEmployeeName} value={this.state.formData.EmployeeName} title={this.state.formData.EmployeeName} onChange={this.handleChange} disabled={this.state.isInputDisabled} />
                                            </div>
                                        </td>
                                        <td>
                                            <div className="light-text">
                                                <input className="form-control" placeholder="Team Lead Name" name="TeamLeadName" type="text" id="txtTeamLeadName" ref={this.txtTeamLeadName} value={this.state.formData.TeamLeadName} title={this.state.formData.TeamLeadName} onChange={this.handleChange} disabled={this.state.isInputDisabled} />
                                            </div>
                                        </td>
                                        <td>
                                            <div className="light-text">
                                                <input className="form-control" placeholder="Supervisor Name" name="SupervisorName" type="text" id="txtSupervisorName" ref={this.txtSupervisorName} value={this.state.formData.SupervisorName} title={this.state.formData.SupervisorName} onChange={this.handleChange} disabled={this.state.isInputDisabled} />
                                            </div>
                                        </td>
                                        <td>
                                            <div className="light-text">
                                                <input className="form-control" placeholder="Dept Mgr Name" name="DeptManagerName" type="text" id="txtDeptManagerName" ref={this.txtDeptManagerName} value={this.state.formData.DeptManagerName} title={this.state.formData.DeptManagerName} onChange={this.handleChange} disabled={this.state.isInputDisabled} />
                                            </div>
                                        </td>
                                        <td>
                                            <div className="light-text">
                                                <input className="form-control" placeholder="Safety Mgr Name" name="SafetyMgrName" type="text" id="txtSafetyManagerName" ref={this.txtSafetyManagerName} value={this.state.formData.SafetyMgrName} title={this.state.formData.SafetyMgrName} onChange={this.handleChange} disabled={this.state.isInputDisabled} />
                                            </div>
                                        </td>
                                        <td>
                                            <div className="light-text">
                                                <input className="form-control" placeholder="Plant Mgr Name" name="PlantMgrName" type="text" id="txtPlantManagerName" ref={this.txtPlantManagerName} value={this.state.formData.PlantMgrName} title={this.state.formData.PlantMgrName} onChange={this.handleChange} disabled={this.state.isInputDisabled} />
                                            </div>
                                        </td>
                                    </tr>
                                    {/* Date */}
                                    <tr>
                                        <th scope="row">Date</th>
                                        <td>
                                            <div className="">
                                                <div className="light-text">
                                                    <div className="custom-datepicker" id="divEmployeeName">
                                                        <DatePickercontrol placeholder="MM/DD/YYYY" selectedDate={this.state.formData.EmployeeDate} title={this.state.formData.EmployeeDate} id='dtEmployeeName' isDisabled={this.state.isInputDisabled} startDate={undefined} endDate={undefined} name="EmployeeDate" onDatechange={(dateProps: any) => this.handleDateChange(dateProps[0], dateProps[2], "divEmployeeName", dateProps)} highlightDate={new Date()} showIcon />
                                                    </div>
                                                </div>
                                            </div>
                                        </td>
                                        <td>
                                            <div className="" >
                                                <div className="light-text">
                                                    <div className="custom-datepicker" id="divTeamLeadDate">                                                    <DatePickercontrol placeholder="MM/DD/YYYY" selectedDate={this.state.formData.TeamLeadDate} title={this.state.formData.TeamLeadDate} id='dtTeamLeadDate' isDisabled={this.state.isInputDisabled} startDate={undefined} endDate={undefined} name="TeamLeadDate" onDatechange={(dateProps: any) => this.handleDateChange(dateProps[0], dateProps[2], "divTeamLeadDate", dateProps)} highlightDate={new Date()} showIcon />
                                                    </div>
                                                </div>
                                            </div>
                                        </td>
                                        <td>
                                            <div className="" >
                                                <div className="light-text">
                                                    <div className="custom-datepicker" id="divSupervisorDate">                                                        <DatePickercontrol placeholder="MM/DD/YYYY" selectedDate={this.state.formData.SupervisorDate} title={this.state.formData.SupervisorDate} id='dtSupervisorDate' isDisabled={this.state.isInputDisabled} startDate={undefined} endDate={undefined} name="SupervisorDate" onDatechange={(dateProps: any) => this.handleDateChange(dateProps[0], dateProps[2], "divSupervisorDate", dateProps)} highlightDate={new Date()} showIcon />
                                                    </div>
                                                </div>
                                            </div>
                                        </td>
                                        <td>
                                            <div className="" id="divDeptMgrDate">
                                                <div className="light-text">
                                                    <div className="custom-datepicker" id="divDeptMgrDate">
                                                        <DatePickercontrol placeholder="MM/DD/YYYY" selectedDate={this.state.formData.DeptManagerDate} title={this.state.formData.DeptManagerDate} id='dtDeptMgrDate' isDisabled={this.state.isInputDisabled} startDate={undefined} endDate={undefined} name="DeptManagerDate" onDatechange={(dateProps: any) => this.handleDateChange(dateProps[0], dateProps[2], "divDeptMgrDate", dateProps)} highlightDate={new Date()} showIcon />
                                                    </div>
                                                </div>
                                            </div>
                                        </td>
                                        <td>
                                            <div className="" >
                                                <div className="light-text">
                                                    <div className="custom-datepicker" id="divSafetyMgrDate">
                                                        <DatePickercontrol placeholder="MM/DD/YYYY" selectedDate={this.state.formData.SafetyMgrDate} title={this.state.formData.SafetyMgrDate} id='dtSafetyMgrDate' isDisabled={this.state.isInputDisabled} startDate={undefined} endDate={undefined} name="SafetyMgrDate" onDatechange={(dateProps: any) => this.handleDateChange(dateProps[0], dateProps[2], "divSafetyMgrDate", dateProps)} highlightDate={new Date()} showIcon />
                                                    </div>
                                                </div>
                                            </div>
                                        </td>
                                        <td>
                                            <div className="" >
                                                <div className="light-text">
                                                    <div className="custom-datepicker" id="divPlantMgrDate">                                                           <DatePickercontrol placeholder="MM/DD/YYYY" selectedDate={this.state.formData.PlantMgrDate} title={this.state.formData.PlantMgrDate} id='dtPlantMgrDate' isDisabled={this.state.isInputDisabled} startDate={undefined} endDate={undefined} name="PlantMgrDate" onDatechange={(dateProps: any) => this.handleDateChange(dateProps[0], dateProps[2], "divPlantMgrDate", dateProps)} highlightDate={new Date()} showIcon />
                                                    </div>
                                                </div>
                                            </div>
                                        </td>
                                    </tr>
                                    {/* Signature */}
                                    <tr>
                                        <th scope="row">Signature</th>
                                        <td>
                                            <div className="light-text">
                                                <input className="form-control" placeholder="Employee Signature" name="EmployeeSignature" type="text" id="txtEmployeeSignature" ref={this.txtEmployeeSignature} value={this.state.formData.EmployeeSignature} onChange={this.handleChange} disabled={this.state.isInputDisabled} title="Employee Signature" />
                                            </div>
                                        </td>
                                        <td>
                                            <div className="light-text">
                                                <input className="form-control" placeholder="Team Lead Signature" name="TeamLeadSignature" type="text" id="txtTeamLeadSignature" ref={this.txtTeamLeadSignature} value={this.state.formData.TeamLeadSignature} onChange={this.handleChange} disabled={this.state.isInputDisabled} title="Team Lead Signature" />
                                            </div>
                                        </td>
                                        <td>
                                            <div className="light-text">
                                                <input className="form-control" placeholder="Supervisor Signature" name="SuperVisorSignature" type="text" id="txtSupervisorSignature" ref={this.txtSupervisorSignature} value={this.state.formData.SuperVisorSignature} onChange={this.handleChange} disabled={this.state.isInputDisabled} title="Supervisor Signature" />
                                            </div>
                                        </td>
                                        <td>
                                            <div className="light-text">
                                                <input className="form-control" placeholder="Dept Mgr Signature" name="DeptManagerSignature" type="text" id="txtDeptManagerSignature" ref={this.txtDeptManagerSignature} value={this.state.formData.DeptManagerSignature} onChange={this.handleChange} disabled={this.state.isInputDisabled} title="Dept Mgr Signature" />
                                            </div>
                                        </td>
                                        <td>
                                            <div className="light-text">
                                                <input className="form-control" placeholder="Safety Mgr Signature" name="SafetyMgrSignature" type="text" id="txtSafetyManagerSignature" ref={this.txtSafetyManagerSignature} value={this.state.formData.SafetyMgrSignature} onChange={this.handleChange} disabled={this.state.isInputDisabled} title="Safety Mgr Signature" />
                                            </div>
                                        </td>
                                        <td>
                                            <div className="light-text">
                                                <input className="form-control" placeholder="Plant Mgr Signature" name="PlantMgrSignature" type="text" id="txtPlantManagerSignature" ref={this.txtPlantManagerSignature} value={this.state.formData.PlantMgrSignature} onChange={this.handleChange} disabled={this.state.isInputDisabled} title="Plant Mgr Signature" />
                                            </div>
                                        </td>
                                    </tr>
                                </table>
                            </div>
                            {/* Injured,Whitness */}
                            <div className="row g-0 mt-2">
                                {/* Injured */}
                                <div className="col-md-6 mt-2">
                                    <div className="form-border-box p-2 mx-1">
                                        <h6 className="yellowbg"><FontAwesomeIcon icon={faUserInjured} />Injured Statement</h6>
                                        {/* Injured Statement */}
                                        <div className="row">
                                            <div className="col-md-12">
                                                <div className="light-text">
                                                    <label className="col-form-label" htmlFor="txtInjuredStatement">Statement </label>
                                                    <textarea className="form-control bs-textarea" rows={3} id="txtInjuredStatement" name="InjuredStatement" ref={this.txtInjuredStatement} placeholder="" value={this.state.formData.InjuredStatement} title={this.state.formData.InjuredStatement} onChange={this.handleChange} disabled={this.state.isInputDisabled} ></textarea>
                                                </div>
                                            </div>
                                            {/* Injured Signature */}
                                            <div className="col-md-6">
                                                <div className="light-text">
                                                    <label className=" col-form-label" htmlFor="txtInjuredSignature">Signature </label>
                                                    <input className="form-control" placeholder="" name="InjuredSignature" type="text" id="txtInjuredSignature" ref={this.txtInjuredSignature} value={this.state.formData.InjuredSignature} title={this.state.formData.InjuredSignature} onChange={this.handleChange} disabled={this.state.isInputDisabled} />
                                                </div>
                                            </div>
                                            {/* Injured Date */}
                                            <div className="col-md-6" >
                                                <div className="light-text">
                                                    <label className="label-datePicker" htmlFor="dtInjuredDate"> Date </label>
                                                    <div className="custom-datepicker" id="divInjuredDate">
                                                        <DatePickercontrol placeholder="MM/DD/YYYY" selectedDate={this.state.formData.InjuredDate} title={this.state.formData.InjuredDate} id='dtInjuredDate' isDisabled={this.state.isInputDisabled} startDate={undefined} endDate={undefined} name="InjuredDate" onDatechange={(dateProps: any) => this.handleDateChange(dateProps[0], dateProps[2], "divInjuredDate", dateProps)} highlightDate={new Date()} showIcon />
                                                    </div>
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                {/* Witness */}
                                <div className="col-md-6 mt-2">
                                    <div className="form-border-box p-2 mx-1">
                                        <h6 className="yellowbg"><FontAwesomeIcon icon={faUserTie} />Witness Statement</h6>
                                        {/* Witness Statement */}
                                        <div className="row">
                                            <div className="col-12">
                                                <div className="light-text" >
                                                    <label className="col-form-label" htmlFor="txtWitnessStatement"> Statement </label>
                                                    <textarea className="form-control bs-textarea" rows={3} id="txtWitnessStatement" name="WitnessStatement" ref={this.txtWitnessStatement} placeholder="" value={this.state.formData.WitnessStatement} title={this.state.formData.WitnessStatement} onChange={this.handleChange} disabled={this.state.isInputDisabled} ></textarea>
                                                </div>
                                            </div>
                                            {/* Witness Signature */}
                                            <div className="col-md-6">
                                                <div className="light-text">
                                                    <label className=" col-form-label" htmlFor="txtWitnessSignature">Signature </label>
                                                    <input className="form-control" placeholder="" name="WitnessSignature" type="text" id="txtWitnessSignature" ref={this.txtWitnessSignature} value={this.state.formData.WitnessSignature} title={this.state.formData.WitnessSignature} onChange={this.handleChange} disabled={this.state.isInputDisabled} />
                                                </div>
                                            </div>
                                            {/* Witness Date */}
                                            <div className="col-md-6">
                                                <div className="light-text">
                                                    <label className="label-datePicker" htmlFor="dtWitnessDate"> Date </label>
                                                    <div className="custom-datepicker" id="divWitnessDate"></div>
                                                    <DatePickercontrol placeholder="MM/DD/YYYY" selectedDate={this.state.formData.WitnessDate} title={this.state.formData.WitnessDate} id='dtWitnessDate' isDisabled={this.state.isInputDisabled} startDate={undefined} endDate={undefined} name="WitnessDate" onDatechange={(dateProps: any) => this.handleDateChange(dateProps[0], dateProps[2], "divWitnessDate", dateProps)} highlightDate={new Date()} showIcon />
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>
                            {/* Buttons */}
                            <div className="col-sm-12 text-center py-3" id="">
                                {this.state.showSubmit && <button type="button" id="btnSubmit" className="btn btn-primary mx-2" onClick={this.handleSubmit} title={this.state.ItemId > 0 ? 'Update' : 'Submit'}>{this.state.ItemId > 0 ? 'Update' : 'Submit'}</button>}
                                <button type="button" id="btnCancel" className="btn btn-secondary" onClick={this.handleCancel} title="Cancel">Cancel</button>
                            </div>
                        </div>
                    </div>
                </React.Fragment >
            )
        }
    }

}