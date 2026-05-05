import * as React from "react";
import { useState, useEffect } from "react";
import { SPHttpClient } from "@microsoft/sp-http";
import { highlightCurrentNav } from "../Utilities/HighlightCurrentComponent";
import { showToast } from "../Shared/Toaster";
import SearchableDropdown from "../Shared/Dropdown";
import { initCommonFunctions } from "../Utilities/CommonFunctions";
import DatePickercontrol from "../Shared/DatePickerField";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { hideLoader, showLoader } from "../Shared/Loader";
import { spfi, SPFx } from "@pnp/sp";
import formValidation from "../Utilities/FormValidator";
import { ControlType } from "../Constants/Contants";
import DateUtilities from "../Utilities/DateUtilities";
import { format } from "date-fns";
import { useParams, useNavigate } from "react-router-dom";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";



export interface LPAFormProps {
  spHttpClient: SPHttpClient;
  context: any;
  siteURL: string;
  spContext: any;
}

const QAMatrixForm: React.FC<LPAFormProps> = (props) => {
  const { id } = useParams<{ id: string }>();
  const sp = spfi().using(SPFx(props.context));
  const navigate = useNavigate();
  const dateRef = React.useRef<any>(null);
  const StartdateRef = React.useRef<any>(null);
  const EnddateRef = React.useRef<any>(null);
  const rootSiteURL = props.spContext.siteAbsoluteUrl;
  const JvisURL = `${rootSiteURL}/mayco`;
  const currentSiteURL = props.spContext.webAbsoluteUrl;
  const txtProblemDes = React.useRef<HTMLInputElement>(null);
  const txtFrequency = React.useRef<HTMLInputElement>(null);
  const txtMaterial = React.useRef<HTMLInputElement>(null);
  const txtSeriousness = React.useRef<HTMLInputElement>(null);



  const { getListItems } = initCommonFunctions(props.context, props.siteURL);
  const [Plants, setPlants] = useState<any[]>([]);
  const [Kpi, setKpi] = useState<any[]>([])
  const [Departments, setDepartments] = useState<any[]>([]);
  const [allZones, setAllZones] = useState<any[]>([]);
  const [filteredZones, setFilteredZones] = useState<any[]>([]);
  const [allMachines, setAllMachines] = useState<any[]>([]);
  const [filteredMachines, setFilteredMachines] = useState<any[]>([]);
  const [itemId, setItemId] = useState(0);
  const [severity, setSeverity] = useState({
    SupplierSort: false,
    PrimaryProductionStation: false,
    SecondaryStation: false,
    WarehouseSort: false,
    Repair: false,
    CustomerConcern: false,
    CPR: false,
    CPA: false,
    Yardhold: false,
    Warranty: false
  });
  const [formData, setFormData] = useState({
    Date: "",
    Plant: "",
    Department: "",
    Zone_x0009_: "",
    Machine: "",
    KPI: "",
    Problem_x0020_Description: "",
    Frequency: "",
    Frequency_x0020_Ratio: "",
    MaterialCost: "",
    MaterialCostRatio: "",
    Seriousness: "",
    SeriousnessRatio: "",
    Problem_x0020_Classification: "",
    Priority_x0020_Value: "",
    Type_x0020_of_x0020_Kaizen: "",
    Project_x0020_LeaderId: [] as number[],
    Project_x0020_LeaderEmail: "" as string,
    Date_x0020_Started: "",
    Date_x0020_Closed: "",
    Receiving_x0020_Inspection_x002f: 0,
    Primary_x0020_Production_x0020_S: 0,
    Secondary_x0020_Station: 0,
    QC_x002f_QE_x0020_Audit_x0020__x: 0,
    Plant_x0020_Rep_x0020__x002f__x0: 0,
    Customer_x0020_Concern: 0,
    _x0033_CPR: 0,
    CPA: 0,
    Yardpurge_x002f_Yardhold: 0,
    Warranty: 0,
    NOT_x0020_STARTED: true,
    PLAN: false,
    DO: false,
    CHECK: false,
    ACT: false,
    COMPLETE: false,
    Man: false,
    Material: false,
    Method: false,
    Year: "",
    OData__x0034_M_x0020_Machine: false,
    YearMonth: ""

  });
  const severityValues = {
    SupplierSort: 1,
    PrimaryProductionStation: 2,
    SecondaryStation: 3,
    WarehouseSort: 4,
    Repair: 5,
    CustomerConcern: 6,
    CPR: 7,
    CPA: 8,
    Yardhold: 9,
    Warranty: 10
  };
  const severityLabels = {
    SupplierSort: "Receiving Inspection / Supplier Sort",
    PrimaryProductionStation: "Primary Production Station",
    SecondaryStation: "Secondary Station",
    WarehouseSort: "Warehouse Sort",
    Repair: "Repair",
    CustomerConcern: "Customer Concern",
    CPR: "3 CPR",
    CPA: "CPA",
    Yardhold: "Yard Hold / Yard Purge",
    Warranty: "Warranty"
  };
  useEffect(() => {
    calculatePriorityValue(severity, formData);
  }, [
    severity,
    formData.Frequency_x0020_Ratio,
    formData.MaterialCostRatio,
    formData.SeriousnessRatio
  ]);
  useEffect(() => {
    highlightCurrentNav("liPPETypes");
    showLoader();
    loadData();
       setTimeout(() => {
      const input = dateRef.current?.querySelector("input");
      input?.focus();
    }, 0);
  }, []);

  // useEffect(() => {
  //   if (id && allZones.length > 0 && allMachines.length > 0) {
  //     const numericId = parseInt(id);
  //     setItemId(numericId);
  //     loadDataForEdit(numericId);
  //   }
  // }, [id, allZones, allMachines]);

  // const loadDataForEdit = async (editId: number) => {

  //   try {
  //     showLoader();
  //     // Fetch header item
  //     const item: any = await sp.web.lists.getByTitle("LPA").items.getById(editId)();
  //     // Populate form data



  //     // 🔥 Populate Zones based on Department
  //     const filteredZ = allZones.filter(
  //       (z: any) => z.Department?.Title === item.Department
  //     );

  //     setFilteredZones(filteredZ);

  //     // 🔥 Populate Machines based on Zone
  //     const filteredM = allMachines.filter(
  //       (m: any) => m.Zone?.Title === item.Zone_x0009_
  //     );

  //     setFilteredMachines(filteredM);
  //     // Fetch line items
  //   } catch (e) {
  //     console.error(e);
  //     showToast("error", "Error loading item for edit");
  //   } finally {
  //     hideLoader();
  //   }
  // };
  useEffect(() => {
    if (id && allZones.length > 0 && allMachines.length > 0) {
      const numericId = parseInt(id);
      setItemId(numericId);
      showLoader();
      loadDataForEdit(numericId);
    }
  }, [id, allZones, allMachines]);

  const loadDataForEdit = async (editId: number) => {
    try {
      

      const item: any = await sp.web.lists
        .getByTitle("QA Matrix")   // ✅ FIXED (was LPA ❌)
        .items.getById(editId)
        .select("*,Project_x0020_Leader/Id,Project_x0020_Leader/Title,Project_x0020_Leader/EMail")
        .expand("Project_x0020_Leader")();

      // ✅ Set form data
      setFormData(prev => ({
        ...prev,
        Date: item.Date ? format(new Date(item.Date), "MM/dd/yyyy") : "",
        Plant: item.Plant || "",
        Department: item.Department || "",
        Zone_x0009_: item.Zone || "",
        Machine: item.Machine || "",
        KPI: item.KPI || "",
        Problem_x0020_Description: item.Problem_x0020_Description || "",
        Frequency: item.Frequency || "",
        Frequency_x0020_Ratio: item.Frequency_x0020_Ratio || "",
        MaterialCost: item.MaterialCost || "",
        MaterialCostRatio: item.MaterialCostRatio || "",
        Seriousness: item.Seriousness || "",
        SeriousnessRatio: item.SeriousnessRatio || "",
        Priority_x0020_Value: item.Priority_x0020_Value || "",
        Problem_x0020_Classification: item.Problem_x0020_Classification || "",
        Type_x0020_of_x0020_Kaizen: item.Type_x0020_of_x0020_Kaizen || "",
        Project_x0020_LeaderId: item.Project_x0020_LeaderId
          ? [item.Project_x0020_LeaderId]
          : [],
        Project_x0020_LeaderEmail: item.Project_x0020_Leader?.EMail || "",
        Date_x0020_Started: item.Date_x0020_Started
          ? format(new Date(item.Date_x0020_Started), "MM/dd/yyyy")
          : "",
        Date_x0020_Closed: item.Date_x0020_Closed
          ? format(new Date(item.Date_x0020_Closed), "MM/dd/yyyy")
          : "",

        // Status
        NOT_x0020_STARTED: item.NOT_x0020_STARTED,
        PLAN: item.PLAN,
        DO: item.DO,
        CHECK: item.CHECK,
        ACT: item.ACT,
        COMPLETE: item.COMPLETE,

        // 4M
        OData__x0034_M_x0020_Machine: item.OData__x0034_M_x0020_Machine,
        Man: item.Man,
        Material: item.Material,
        Method: item.Method
      }));

      // ✅ Set Severity checkboxes
      setSeverity({
        SupplierSort: item.Receiving_x0020_Inspection_x002f === 1,
        PrimaryProductionStation: item.Primary_x0020_Production_x0020_S === 2,
        SecondaryStation: item.Secondary_x0020_Station === 3,
        WarehouseSort: item.QC_x002f_QE_x0020_Audit_x0020__x === 4,
        Repair: item.Plant_x0020_Rep_x0020__x002f__x0 === 5,
        CustomerConcern: item.Customer_x0020_Concern === 6,
        CPR: item._x0033_CPR === 7,
        CPA: item.CPA === 8,
        Yardhold: item.Yardpurge_x002f_Yardhold === 9,
        Warranty: item.Warranty === 10
      });

      // ✅ Populate dependent dropdowns
      const filteredZ = allZones.filter(
        (z: any) => z.Department?.Title === item.Department
      );
      setFilteredZones(filteredZ);

      const filteredM = allMachines.filter(
        (m: any) => m.Zone?.Title === item.Zone
      );
      setFilteredMachines(filteredM);

    } catch (e) {
      console.error(e);
      showToast("error", "Error loading item for edit");
    } finally {
      hideLoader();
    }
  };


  const loadData = async () => {
    try {
      const [plantData, deptData, ZoneData, MachineData, KPIData] = await Promise.all([
        getListItems("Plant", JvisURL, "*", "", "Title eq 'Merrill'"),
        getListItems(
          "Department",
          JvisURL,
          "Plant/Title,Title",
          "Plant",
          "Plant/Title eq 'Merrill'"
        ),
        getListItems("Zones", JvisURL, "Plant/Title,Department/Title,Title", "Plant,Department", "Plant/Title eq 'Merrill'"),
        getListItems("Machines", JvisURL, "Plant/Title,Department/Title,Zone/Title,Title", "Plant,Department,Zone", "Plant/Title eq 'Merrill'"),
        getListItems("KPI", currentSiteURL, "*", "", "")
      ]);
      setPlants(plantData);
      setDepartments(deptData);
      setAllZones(ZoneData);
      setFilteredZones([]);
      setAllMachines(MachineData);
      setFilteredMachines([]);
      setKpi(KPIData);


      if (plantData.length > 0) {
        setFormData(prev => ({
          ...prev,
          Plant: plantData[0].Title
        }));
      }
    }

    catch (e) {
      console.log(e);
      showToast("error", "Error loading data");
    }
    finally {
      hideLoader();
    }
  };






  const getStateControl = (formData: any) => {
    const { PLAN, DO, CHECK, ACT, COMPLETE } = formData;

    let disabled = {
      NOT_x0020_STARTED: false,
      PLAN: false,
      DO: false,
      CHECK: false,
      ACT: false,
      COMPLETE: false,
      show4M: false
    };

    if (!PLAN && !DO && !CHECK && !ACT && !COMPLETE) {
      // NOT STARTED
      disabled = {
        NOT_x0020_STARTED: false,
        PLAN: false,
        DO: true,
        CHECK: true,
        ACT: true,
        COMPLETE: true,
        show4M: false
      };
    }
    else if (COMPLETE) {
      disabled = {
        NOT_x0020_STARTED: true,
        PLAN: false,
        DO: false,
        CHECK: false,
        ACT: false,
        COMPLETE: false,
        show4M: true
      };
    }
    else if (ACT) {
      disabled = {
        NOT_x0020_STARTED: true,
        PLAN: false,
        DO: false,
        CHECK: false,
        ACT: false,
        COMPLETE: false,
        show4M: true
      };
    }
    else if (CHECK) {
      disabled = {
        NOT_x0020_STARTED: true,
        PLAN: false,
        DO: false,
        CHECK: false,
        ACT: false,
        COMPLETE: true,
        show4M: false
      };
    }
    else if (DO) {
      disabled = {
        NOT_x0020_STARTED: true,
        PLAN: false,
        DO: false,
        CHECK: false,
        ACT: true,
        COMPLETE: true,
        show4M: false
      };
    }
    else if (PLAN) {
      disabled = {
        NOT_x0020_STARTED: true,
        PLAN: false,
        DO: false,
        CHECK: true,
        ACT: true,
        COMPLETE: true,
        show4M: false
      };
    }

    return disabled;
  };
  const disabledState = getStateControl(formData);
  const isPlanSelected = formData.PLAN;
  const isCompleteSelected = formData.COMPLETE
  const isActSelected = formData.ACT
  const isAnyAnalysisSelected =
    formData.OData__x0034_M_x0020_Machine ||
    formData.Man ||
    formData.Material ||
    formData.Method;

  const handleFormCheckboxChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const { name, checked } = e.target;

    setFormData(prev => {
      const updated = { ...prev };

      switch (name) {
        case "PLAN":
          updated.PLAN = checked;
          if (!checked) {
            updated.DO = false;
            updated.CHECK = false;
            updated.ACT = false;
            updated.COMPLETE = false;
          }
          break;

        case "DO":
          updated.DO = checked;

          if (!checked) {
            // 🔥 force reset downstream
            updated.CHECK = false;
            updated.ACT = false;
            updated.COMPLETE = false;
          }

          return updated;

        case "CHECK":
          updated.CHECK = checked;
          if (!checked) {
            updated.ACT = false;
            updated.COMPLETE = false;
          }
          break;

        case "ACT":
          updated.ACT = checked;
          if (!checked) {
            updated.COMPLETE = false;
          }
          break;

        case "COMPLETE":
          if (prev.ACT) updated.COMPLETE = checked;
          break;
      }
      updated.NOT_x0020_STARTED =
        !updated.PLAN &&
        !updated.DO &&
        !updated.CHECK &&
        !updated.ACT &&
        !updated.COMPLETE;

      return updated;
    });
  };


  const handleChangeDropdown = (selectedOption: any, actionMeta: any) => {
    if (!actionMeta?.name) return;

    const name = actionMeta.name;

    const value =
      actionMeta.action === "clear"
        ? name === "AuditorId" ? 0 : ""   // ✅ fix for lookup
        : selectedOption?.value;

    /* ------------------ HARD DEPENDENCY BLOCK ------------------ */

    if (name === "Zone_x0009_" && !formData.Department) return;

    if (name === "Machine" && !formData.Zone_x0009_) return;

    /* ------------------ DEPARTMENT ------------------ */

    if (name === "Department") {

      const filteredZ = allZones.filter(
        (z: any) => z.Department?.Title === value
      );

      setFilteredZones(filteredZ);
      setFilteredMachines([]);
      setFormData(prev => ({
        ...prev,
        Department: value,
        Zone_x0009_: "",
        Machine: ""
      }));

      if (value) {

      }

      return;
    }

    /* ------------------ ZONE ------------------ */

    if (name === "Zone_x0009_") {

      const filteredM = allMachines.filter(
        (m: any) => m.Zone?.Title === value
      );

      setFilteredMachines(filteredM);

      setFormData(prev => ({
        ...prev,
        Zone_x0009_: value,
        Machine: ""
      }));

      return;
    }

    /* ------------------ MACHINE ------------------ */

    if (name === "Machine") {
      setFormData(prev => ({
        ...prev,
        Machine: value
      }));
      return;
    }
    /* ------------------ Kaizen ------------------ */
    if (name === "Type_x0020_of_x0020_Kaizen") {
      setFormData(prev => ({
        ...prev,
        Type_x0020_of_x0020_Kaizen: value
      }));
      return;
    }
    /* ------------------ KPI ------------------ */
    if (name === "KPI") {
      setFormData(prev => ({
        ...prev,
        KPI: value
      }));
      return;
    }



    /* ------------------ SUPERVISOR ------------------ */

    if (name === "Supervisor") {
      setFormData(prev => ({
        ...prev,
        Supervisor: value
      }));
      return;
    }

    /* ------------------ TOOL NUMBER ------------------ */

    if (name === "Tool_x0020_Number") {
      setFormData(prev => ({
        ...prev,
        Tool_x0020_Number: value
      }));
      return;
    }

    /* ------------------ DEFAULT ------------------ */

    setFormData(prev => ({
      ...prev,
      [name]: value
    }));
  };



  const handleSubmit = async (event: any) => {

    event.preventDefault();
    showLoader();
    const selectedSeverity = Object.keys(severity).filter(
      key => severity[key as keyof typeof severity]
    );
    const data = {
      Date: { val: formData.Date, required: true, Name: "Date", Type: ControlType.date, Focusid: 'dtDate' },
      Department: { val: formData.Department.trim(), required: true, Name: "Department", Type: ControlType.reactSelect, Focusid: 'txtDepartment' },
      Zone: { val: formData.Zone_x0009_.trim(), required: true, Name: "Zone", Type: ControlType.reactSelect, Focusid: 'txtZone' },
      Machine: { val: formData.Machine.trim(), required: true, Name: "Machine", Type: ControlType.reactSelect, Focusid: 'txtMachine' },
      KPI: { val: formData.KPI.trim(), required: true, Name: "KPI", Type: ControlType.reactSelect, Focusid: 'txtKPI' },
      ProblemDes: { val: formData.Problem_x0020_Description.trim(), required: true, Name: "Problem Description", Type: ControlType.string, Focusid: txtProblemDes },
      Frequency: { val: formData.Frequency, required: true, Name: "Frequency", Type: ControlType.string, Focusid: txtFrequency },
      Materialcost: { val: formData.MaterialCost, required: true, Name: "Material Cost", Type: ControlType.string, Focusid: txtMaterial },
      Sriousness: { val: formData.Seriousness, required: true, Name: "Seriousness", Type: ControlType.string, Focusid: txtSeriousness },
      Severity: {
        val: selectedSeverity,
        required: true,
        Name: "Severity",
        Type: ControlType.array,   // 👈 already supported in your validator
        Focusid: "severityDiv"     // 👈 give a div id in UI
      },
      ProjectLeader: {
        val: formData.Project_x0020_LeaderId,
        required: true,
        Name: "Project Leader",
        Type: ControlType.people,
        Focusid: "projectLeaderDiv"
      },
      DateStarted: { val: formData.Date_x0020_Started, required: formData.PLAN, Name: "Date Started", Type: ControlType.date, Focusid: 'dtstartDate' },
      DateCompare: {
        Type: ControlType.compareDates,
        startDate: formData.Date,
        endDate: formData.Date_x0020_Started,
        startDateName: "Date",
        endDateName: "Date Started",
        Focusid: "dtstartDate"
      },
      DateClosed: { val: formData.Date_x0020_Closed, required: formData.COMPLETE, Name: "Date Closed", Type: ControlType.date, Focusid: 'dtendDate' },
      EndDate: {
        Type: ControlType.compareDates,
        startDate: formData.Date_x0020_Started,
        endDate: formData.Date_x0020_Closed,
        startDateName: "Date Started",
        endDateName: "Date Closed",
        Focusid: "dtendDate"
      },
      Analysis: {
        val: [
          formData.OData__x0034_M_x0020_Machine,
          formData.Man,
          formData.Material,
          formData.Method
        ].filter(Boolean),
        required: formData.ACT,
        Name: "4M Analysis",
        Type: ControlType.array,
        Focusid: "analysisDiv"
      }


    };

    const isValid = formValidation.FormValidation(data);

    if (isValid.status) {
      insertOrUpdate();

      // const validDuplicate = await checkDuplicate();

      // if (validDuplicate) insertOrUpdate();

    } else {

      showToast("error", isValid.message);
      hideLoader();

    }

  };
  // const handleChange = (e: any) => {
  //   const { name, value } = e.target;

  //   setFormData(prev => ({
  //     ...prev,
  //     [name]: value
  //   }));
  // };
  const handleDateChange = (
    dateValue: any,
    name: string,
    divId: string,
    dateProps?: any
  ) => {

    setFormData((prev: any) => {

      const updatedFormData = { ...prev };

      // Handle floating label active class
      if (divId) {
        const ddlElement = document.getElementById(divId);

        if (dateValue) {
          ddlElement?.classList.add("active");
        } else {
          ddlElement?.classList.remove("active");
        }
      }

      // Date conversion
      let formattedDate = "";

      if (dateValue != null) {
        formattedDate = format(
          DateUtilities.addBrowserwrtServer(
            new Date(DateUtilities.getDateMMDDYYYY(dateValue)),
            props.spContext.webTimeZoneData
          ).toISOString(),
          "MM/dd/yyyy"
        );
      }

      updatedFormData[name] = formattedDate;

      return updatedFormData;
    });
  };

  // const handleCheckboxChange = (e: React.ChangeEvent<HTMLInputElement>) => {
  //   const name = e.target.name as keyof typeof severity;

  //   setSeverity(prev => ({
  //     ...prev,
  //     [name]: e.target.checked
  //   }));
  // };
  const calculatePriorityValue = (
    sevState: typeof severity,
    form = formData
  ) => {
    const frequency = Number(form.Frequency_x0020_Ratio) || 0;
    const material = Number(form.MaterialCostRatio) || 0;
    const seriousness = Number(form.SeriousnessRatio) || 0;

    let severityTotal = 0;

    (Object.keys(sevState) as Array<keyof typeof sevState>).forEach(key => {
      if (sevState[key]) {
        severityTotal += severityValues[key];
      }
    });

    const priority = frequency * (material + seriousness) * severityTotal;

    // ✅ prevent unnecessary re-render
    if (form.Priority_x0020_Value === priority.toString()) return;

    setFormData(prev => ({
      ...prev,
      Priority_x0020_Value: priority.toString(),

      Receiving_x0020_Inspection_x002f: sevState.SupplierSort ? 1 : 0,
      Primary_x0020_Production_x0020_S: sevState.PrimaryProductionStation ? 1 : 0,
      Secondary_x0020_Station: sevState.SecondaryStation ? 1 : 0,
      QC_x002f_QE_x0020_Audit_x0020__x: sevState.WarehouseSort ? 1 : 0,
      Plant_x0020_Rep_x0020__x002f__x0: sevState.Repair ? 1 : 0,
      Customer_x0020_Concern: sevState.CustomerConcern ? 1 : 0,
      _x0033_CPR: sevState.CPR ? 1 : 0,
      CPA: sevState.CPA ? 1 : 0,
      Yardpurge_x002f_Yardhold: sevState.Yardhold ? 1 : 0,
      Warranty: sevState.Warranty ? 1 : 0
    }));
  };
  const handleCheckboxChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const name = e.target.name as keyof typeof severity;

    const updatedSeverity = {
      ...severity,
      [name]: e.target.checked
    };

    setSeverity(updatedSeverity); // ✅ only this
  };

  const handleAnalysisCheckboxChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const { name, checked } = e.target;

    setFormData(prev => ({
      ...prev,
      // reset all first
      OData__x0034_M_x0020_Machine: false,
      Man: false,
      Material: false,
      Method: false,

      // set only selected one
      [name]: checked
    }));
  };
  //  const handleChange = (event: any) => {

  //     const { name, value } = event.target;
  //       if (name === "Frequency") {
  //     if (!/^[0-9]*$/.test(value)) return;   // only numbers
  //     if (value.length > 4) return;         // max 4 digits
  //   }
  //     if (name === "MaterialCost") {
  //     if (!/^[0-9]*$/.test(value)) return;
  //     if (value.length > 10) return;
  //   }
  //    if (name === "Seriousness") {
  //     if (!/^[1-5]?$/.test(value)) return; // allow only 1–5 (single digit)
  //   }
  //     setFormData({
  //       ...formData,
  //       [name]: value
  //     });

  //   };

  const handlePeoplePickerChange = (items: any[]) => {
    try {
      showLoader();

      const userIds = items?.map(item => Number(item.id)) || [];

      setFormData(prev => ({
        ...prev,
        Project_x0020_LeaderId: userIds   // ✅ always array
      }));

    } catch (e) {
      console.log("Error in peoplepicker", e);
    } finally {
      hideLoader();
    }
  };

  const handleChange = (event: any) => {
    const { name, value } = event.target;

    // ================= VALIDATIONS =================

    // Frequency → only numbers, max 4 digits
    if (name === "Frequency") {
      if (!/^[0-9]*$/.test(value)) return;
      if (value.length > 4) return;
    }

    // Material Cost → only numbers, max 10 digits
    if (name === "MaterialCost") {
      if (!/^[0-9]*$/.test(value)) return;
      if (value.length > 10) return;
    }

    // Seriousness → only 1 to 5
    if (name === "Seriousness") {
      if (!/^[1-5]?$/.test(value)) return;
    }

    // ================= CALCULATIONS =================

    let updatedData = {
      ...formData,
      [name]: value
    };

    // ✅ Frequency → Ratio
    if (name === "Frequency") {
      const val = Number(value);

      if (!value) {
        updatedData.Frequency_x0020_Ratio = "";
      } else if (val >= 0.1 && val <= 1) {
        updatedData.Frequency_x0020_Ratio = "1";
      } else if (val > 1 && val <= 3) {
        updatedData.Frequency_x0020_Ratio = "2";
      } else if (val > 3 && val <= 5) {
        updatedData.Frequency_x0020_Ratio = "3";
      } else if (val > 5 && val <= 10) {
        updatedData.Frequency_x0020_Ratio = "4";
      } else if (val > 10) {
        updatedData.Frequency_x0020_Ratio = "5";
      }
    }

    // ✅ Material Cost → Ratio
    if (name === "MaterialCost") {
      const val = Number(value);

      if (!value) {
        updatedData.MaterialCostRatio = "";
      } else if (val <= 500) {
        updatedData.MaterialCostRatio = "1";
      } else if (val <= 1000) {
        updatedData.MaterialCostRatio = "2";
      } else if (val <= 2500) {
        updatedData.MaterialCostRatio = "3";
      } else if (val <= 5000) {
        updatedData.MaterialCostRatio = "4";
      } else {
        updatedData.MaterialCostRatio = "5";
      }
    }

    // ✅ Seriousness → Ratio + Classification
    if (name === "Seriousness") {
      const val = Number(value);

      if (!value) {
        updatedData.SeriousnessRatio = "";
        updatedData.Problem_x0020_Classification = "";
      } else {
        updatedData.SeriousnessRatio = value;

        if (val === 1 || val === 2) {
          updatedData.Problem_x0020_Classification = "C";
        } else if (val === 3) {
          updatedData.Problem_x0020_Classification = "B";
        } else if (val === 4) {
          updatedData.Problem_x0020_Classification = "A";
        } else if (val === 5) {
          updatedData.Problem_x0020_Classification = "AA";
        }
      }
    }

    // ================= SET STATE =================
    setFormData(updatedData);
  };

  const insertOrUpdate = async () => {
    try {
      showLoader();

      /* ---------------- HEADER DATA ---------------- */

      const mainData: any = {
        Date: formData.Date ? new Date(formData.Date).toISOString() : null,
        Plant: formData.Plant,
        Department: formData.Department,
        Zone: formData.Zone_x0009_,
        Machine: formData.Machine,
        KPI: formData.KPI,
        Problem_x0020_Description: formData.Problem_x0020_Description,
        Frequency: formData.Frequency,
        Frequency_x0020_Ratio: formData.Frequency_x0020_Ratio,
        MaterialCost: formData.MaterialCost,
        MaterialCostRatio: formData.MaterialCostRatio,
        Seriousness: formData.Seriousness,
        SeriousnessRatio: formData.SeriousnessRatio,
        Receiving_x0020_Inspection_x002f: severity.SupplierSort ? 1 : 0,
        Primary_x0020_Production_x0020_S: severity.PrimaryProductionStation ? 2 : 0,
        Secondary_x0020_Station: severity.SecondaryStation ? 3 : 0,
        QC_x002f_QE_x0020_Audit_x0020__x: severity.WarehouseSort ? 4 : 0,
        Plant_x0020_Rep_x0020__x002f__x0: severity.Repair ? 5 : 0,
        Customer_x0020_Concern: severity.CustomerConcern ? 6 : 0,
        //  _x0033_CPR:severity.CPR ? 7:0,
        CPA: severity.CPA ? 8 : 0,
        Yardpurge_x002f_Yardhold: severity.Yardhold ? 9 : 0,
        Warranty: severity.Warranty ? 10 : 0,
        Priority_x0020_Value: formData.Priority_x0020_Value,
        Problem_x0020_Classification: formData.Problem_x0020_Classification,
        Type_x0020_of_x0020_Kaizen: formData.Type_x0020_of_x0020_Kaizen,
        Project_x0020_LeaderId: formData.Project_x0020_LeaderId[0] || null,
        Date_x0020_Started: formData.Date_x0020_Started ? new Date(formData.Date_x0020_Started).toISOString() : null,
        Date_x0020_Closed: formData.Date_x0020_Closed ? new Date(formData.Date_x0020_Closed).toISOString() : null,
        NOT_x0020_STARTED: formData.NOT_x0020_STARTED,
        PLAN: formData.PLAN,
        DO: formData.DO,
        CHECK: formData.CHECK,
        ACT: formData.ACT,
        COMPLETE: formData.COMPLETE,
        OData__x0034_M_x0020_Machine: formData.OData__x0034_M_x0020_Machine,
        Man: formData.Man,
        Material: formData.Material,
        Method: formData.Method,
        Year: new Date().getFullYear().toString(),
        YearMonth: (new Date().getMonth() + 1).toString()
      };



      /* ---------------- HEADER SAVE ---------------- */

      if (itemId > 0) {
        await sp.web.lists
          .getByTitle("QA Matrix")
          .items.getById(itemId)
          .update(mainData);

        showToast("success", "QA-Matrix updated successfully");
      } else {
        await sp.web.lists
          .getByTitle("QA Matrix")
          .items.add(mainData);

        showToast("success", "QA-Matrix created successfully");
      }

      /* ---------------- CHILD UPSERT ---------------- */
      closeForm();
      loadData();

    } catch (e) {
      console.log(e);
      showToast("error", "Error saving data");
    } finally {
      hideLoader();
    }
  };
  const closeForm = () => {

    setItemId(0);
    navigate('/QA-MatrixView');
  };

  return (
    <div className="container-fluid">
      <div className="light-box border-box-shadow">
        <div className="div-form-title">
          <div className="form-title">QA-Matrix</div>


          <span className="span-mandatory-text">
            <span className="text-danger">*</span> are mandatory fields
          </span>
        </div>

        <div className="p-2 mx-1 ViewTable">


          <div className="form-border-box p-2 mx-1 my-2">

            {/* ROW 1 */}
            <div className="row">
              <div className="col-md-3">
                <div className="light-text">
                  <div ref={dateRef}>
                  <label className="" htmlFor="dtDate"> Date <span className="text-danger">*</span></label>
                  <div className="custom-datepicker" id="divDatefield">
                    <DatePickercontrol placeholder="MM/DD/YYYY" selectedDate={formData.Date} ref={dateRef} title={formData.Date} id='dtDate' startDate={undefined} endDate={undefined} name="Date" onDatechange={(dateProps: any) => handleDateChange(dateProps[0], dateProps[2], "divDate", dateProps)} highlightDate={new Date()} showIcon />
                  </div>
                  </div>
                </div>
              </div>
              <div className="col-md-3">
                <div className="light-text">
                  <label>Plant <span className="text-danger">*</span></label>
                  <select className="form-control" value={formData.Plant} disabled>
                    {Plants.map((p: any) => (
                      <option key={p.Id} value={p.Title}>
                        {p.Title}
                      </option>
                    ))}
                  </select>
                </div>
              </div>


              <div className="col-md-3">
                <div className="light-text">
                  <SearchableDropdown
                    label="Department"
                    name="Department"
                    selectedValue={formData.Department}
                    OptionsList={Departments.map(d => ({
                      label: d.Title,
                      value: d.Title
                    }))}
                    OnChange={handleChangeDropdown}
                    isRequired={true} Title={"Department"} id={"txtDepartment"} className={""} disabled={false} />
                </div>

              </div>

              <div className="col-md-3">
                <div className="light-text">
                  <SearchableDropdown
                    label="Zone"
                    name="Zone_x0009_"
                    selectedValue={formData.Zone_x0009_}
                    OptionsList={filteredZones.map(d => ({
                      label: d.Title,
                      value: d.Title
                    }))}
                    OnChange={handleChangeDropdown}
                    isRequired={true} Title={"Zone"} id={"txtZone"} className={""} disabled={false} />
                </div>

              </div>
              <div className="col-md-3">
                <div className="light-text">
                  <SearchableDropdown
                    label="Machine"
                    name="Machine"
                    selectedValue={formData.Machine}
                    OptionsList={filteredMachines.map(d => ({
                      label: d.Title,
                      value: d.Title
                    }))}
                    OnChange={handleChangeDropdown}
                    isRequired={true} Title={"Machine"} id={"txtMachine"} className={""} disabled={false} />
                </div>

              </div>

              <div className="col-md-3">
                <div className="light-text">
                  <SearchableDropdown
                    label="KPI"
                    name="KPI"
                    selectedValue={formData.KPI}
                    OptionsList={Kpi.map(d => ({
                      label: d.Title,
                      value: d.Title
                    }))}
                    OnChange={handleChangeDropdown}
                    isRequired={true} Title={"KPI"} id={"txtKPI"} className={""} disabled={false} />
                </div>

              </div>
              <div className="col-md-6">

                <div className="light-text">

                  <input
                    className="form-control"
                    type="text"
                    name="Problem_x0020_Description"
                    value={formData.Problem_x0020_Description}
                    onChange={handleChange}
                    ref={txtProblemDes}
                  />

                  <label>
                    Problem Description <span className="mandatoryhastrick">*</span>
                  </label>

                </div>

              </div>
              <div className="col-md-4">

                {/* Frequency */}
                <table className="custom-table w-100 mt-2">
                  <thead>
                    <tr>
                      <th colSpan={2}>Frequency <span className="text-danger">*</span></th>
                    </tr>
                    <tr>
                      <th>%</th>
                      <th>R=1:5</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr>
                      <td>
                        <input
                          className="form-control"
                          type="text"
                          name="Frequency"
                          value={formData.Frequency}
                          onChange={handleChange}
                          maxLength={250}
                          ref={txtFrequency}
                        />
                      </td>
                      <td>
                        <input
                          className="form-control"
                          type="text"
                          name="Frequency_x0020_Ratio"
                          value={formData.Frequency_x0020_Ratio}
                          onChange={handleChange}
                          maxLength={250}
                          readOnly
                        />
                      </td>
                    </tr>
                  </tbody>
                </table>
                </div>
                

                {/* Material Cost */}
                <div className="col-md-4">
                <table className="custom-table w-100 mt-2">
                  <thead>
                    <tr>
                      <th colSpan={2}>Material Cost <span className="text-danger">*</span></th>
                    </tr>
                    <tr>
                      <th>$</th>
                      <th>R=1:5</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr>
                      <td>
                        <input
                          className="form-control"
                          type="text"
                          name="MaterialCost"
                          value={formData.MaterialCost}
                          onChange={handleChange}
                          maxLength={250}
                          ref={txtMaterial}
                        />
                      </td>
                      <td>
                        <input
                          className="form-control"
                          type="text"
                          name="MaterialCostRatio"
                          value={formData.MaterialCostRatio}
                          onChange={handleChange}
                          maxLength={250}
                          readOnly
                        />
                      </td>
                    </tr>
                  </tbody>
                </table>
                </div>

                {/* Seriousness */}
                  <div className="col-md-4">
                <table className="custom-table w-100 mt-2">
                  <thead>
                    <tr>
                      <th colSpan={2}>Seriousness <span className="text-danger">*</span></th>
                    </tr>
                    <tr>
                      <th>min</th>
                      <th>R=1:5</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr>
                      <td>
                        <input
                          className="form-control"
                          type="text"
                          name="Seriousness"
                          value={formData.Seriousness}
                          onChange={handleChange}
                          maxLength={250}
                          placeholder="only 1 to 5"
                          ref={txtSeriousness}
                        />
                      </td>
                      <td>
                        <input
                          className="form-control"
                          type="text"
                          name="SeriousnessRatio"
                          value={formData.SeriousnessRatio}
                          onChange={handleChange}
                          maxLength={250}
                          readOnly

                        />
                      </td>
                    </tr>
                  </tbody>
                </table>
                </div>

              
              <div className="row mt-3">

                {/* LEFT - Severity */}
                <div className="col-md-4">
                  <div className="custom-box info-box p-3 mb-2">
                    <div className="box-title fw-bold pb-2 mb-2 border-bottom">Severity <span className="text-danger">*</span></div>

                    {(Object.keys(severity) as Array<keyof typeof severity>).map((key) => (
                      <div key={key}>
                        <input
                          type="checkbox"
                          id={key}
                          name={key}
                          checked={severity[key]}
                          onChange={handleCheckboxChange}
                        />
                        <label htmlFor={key} className="ms-2">
                          {severityLabels[key]}
                        </label>
                      </div>
                    ))}

                  </div>
                </div>
                <div className="col-md-4">
                  <div className="custom-box info-box p-3 mb-2">
                    <div className="light-text">
                      <label>Priority Value</label>
                      <input type="text" className="form-control mb-2" value={formData.Priority_x0020_Value} readOnly />
                    </div>
                    <div className="light-text">
                      <label>Problem Classification</label>
                      <input type="text" className="form-control mb-2" value={formData.Problem_x0020_Classification} readOnly />
                    </div>
                    <div className="light-text">
                      <SearchableDropdown
                        label="Type of Kaizen"
                        name="Type_x0020_of_x0020_Kaizen"
                        selectedValue={formData.Type_x0020_of_x0020_Kaizen}
                        OptionsList={[
                          { label: "Quick", value: "Quick" },
                          { label: "Standard", value: "Standard" },
                          { label: "Major", value: "Major" },
                          { label: "Advanced", value: "Advanced" },
                        ]}
                        OnChange={handleChangeDropdown}
                        isRequired={false} Title={"Type_x0020_of_x0020_Kaizen"} id={"txtkaizen"} className={""} disabled={false} />
                    </div>

                    <div className="light-text" id="projectLeaderDiv">
                      <label>Project Leader <span className="text-danger">*</span></label>

                      <PeoplePicker
                        context={props.context}
                        webAbsoluteUrl={props.context.pageContext.web.absoluteUrl}
                        titleText=""
                        personSelectionLimit={1}
                        showtooltip={true}
                        required={true}
                        disabled={false}
                        ensureUser={true}
                        principalTypes={[PrincipalType.User]}
                        resolveDelay={500}
                        onChange={handlePeoplePickerChange}
                        defaultSelectedUsers={
                          formData.Project_x0020_LeaderEmail
                            ? [formData.Project_x0020_LeaderEmail]
                            : []
                        }
                      />
                    </div>

                    <div className="light-text">
                      <label className="" htmlFor="dtstartDate"> Date Started {isPlanSelected && <span className="text-danger">*</span>}</label>
                      <div className="custom-datepicker" id="divDatestarted">
                        <DatePickercontrol placeholder="MM/DD/YYYY" selectedDate={formData.Date_x0020_Started} ref={StartdateRef} title={formData.Date_x0020_Started} id='dtstartDate' startDate={undefined} endDate={undefined} name="Date" onDatechange={(dateProps: any) => handleDateChange(dateProps[0], "Date_x0020_Started", "divDatestarted", dateProps)} highlightDate={new Date()} showIcon />
                      </div>
                    </div>

                    <div className="light-text">
                      <label className="" htmlFor="dtendDate"> Date Closed {isCompleteSelected && <span className="text-danger">*</span>}</label>
                      <div className="custom-datepicker" id="divDateclosed">
                        <DatePickercontrol placeholder="MM/DD/YYYY" selectedDate={formData.Date_x0020_Closed} ref={EnddateRef} title={formData.Date_x0020_Closed} id='dtendDate' startDate={undefined} endDate={undefined} name="Date" onDatechange={(dateProps: any) => handleDateChange(dateProps[0], "Date_x0020_Closed", "dtendDate", dateProps)} highlightDate={new Date()} showIcon />
                      </div>
                    </div>

                  </div>
                </div>
                <div className="col-md-4">
                  <div className="custom-box info-box p-3 mb-2 ">

                    {/* Current Status */}
                    <div className="box-title fw-bold pb-2 mb-2 border-bottom">Current Status <span className="text-danger">*</span> </div>

                    <div>
                      <input
                        type="checkbox"
                        name="NOT_x0020_STARTED"
                        checked={formData.NOT_x0020_STARTED}
                        // checked={
                        //   !formData.PLAN &&
                        //   !formData.DO &&
                        //   !formData.CHECK &&
                        //   !formData.ACT &&
                        //   !formData.COMPLETE}
                        onChange={handleFormCheckboxChange}
                        disabled={disabledState.NOT_x0020_STARTED}
                      />
                      <label className="ms-2">NOT STARTED</label>
                    </div>

                    <div>
                      <input
                        type="checkbox"
                        name="PLAN"
                        checked={formData.PLAN}
                        onChange={handleFormCheckboxChange}
                        disabled={disabledState.PLAN}
                      />
                      <label className="ms-2">PLAN</label>
                    </div>

                    <div>
                      <input
                        type="checkbox"
                        name="DO"
                        checked={formData.DO}
                        onChange={handleFormCheckboxChange}
                        disabled={!formData.PLAN}
                      />
                      <label className="ms-2">DO</label>
                    </div>

                    <div>
                      <input
                        type="checkbox"
                        name="CHECK"
                        checked={formData.CHECK}
                        onChange={handleFormCheckboxChange}
                        disabled={!formData.DO}
                      />
                      <label className="ms-2">CHECK</label>
                    </div>

                    <div>
                      <input
                        type="checkbox"
                        name="ACT"
                        checked={formData.ACT}
                        onChange={handleFormCheckboxChange}
                        disabled={!formData.CHECK}
                      />
                      <label className="ms-2">ACT</label>
                    </div>

                    <div>
                      <input
                        type="checkbox"
                        name="COMPLETE"
                        checked={formData.COMPLETE}
                        onChange={handleFormCheckboxChange}
                        disabled={disabledState.COMPLETE}
                      />
                      <label className="ms-2">COMPLETE</label>
                    </div>

                    {/* 4M Analysis */}
                    <div className="box-title fw-bold pb-2 mt-3 mb-2 border-bottom">4M Analysis {isActSelected && <span className="text-danger">*</span>}</div>

                    <div>
                      <input
                        type="checkbox"
                        name="OData__x0034_M_x0020_Machine"
                        checked={formData.OData__x0034_M_x0020_Machine}
                        onChange={handleAnalysisCheckboxChange}
                        disabled={
                          isAnyAnalysisSelected && !formData.OData__x0034_M_x0020_Machine
                        }
                      />
                      <label className="ms-2">Machine</label>
                    </div>

                    <div>
                      <input
                        type="checkbox"
                        name="Man"
                        checked={formData.Man}
                        onChange={handleAnalysisCheckboxChange}
                        disabled={
                          isAnyAnalysisSelected && !formData.Man
                        }
                      />
                      <label className="ms-2">Man</label>
                    </div>

                    <div>
                      <input
                        type="checkbox"
                        name="Material"
                        checked={formData.Material}
                        onChange={handleAnalysisCheckboxChange}
                        disabled={
                          isAnyAnalysisSelected && !formData.Material
                        }
                      />
                      <label className="ms-2">Material</label>
                    </div>

                    <div>
                      <input
                        type="checkbox"
                        name="Method"
                        checked={formData.Method}
                        onChange={handleAnalysisCheckboxChange}
                        disabled={
                          isAnyAnalysisSelected && !formData.Method
                        }
                      />
                      <label className="ms-2">Method</label>
                    </div>

                  </div>
                </div>

              </div>
            </div>
          </div>

          {/* BUTTONS */}
          <div className="text-center mt-3">
            <button className="btn btn-primary mx-2" onClick={handleSubmit}>{itemId > 0 ? "Update" : "Submit"}</button>
            <button className="btn btn-secondary" onClick={closeForm}>
              Cancel
            </button>
          </div>
          <div className="row mt-4">

            {/* Frequency Box */}
            <div className="col-md-4">
              <div className="info-box">
                <div className="info-title">Frequency</div>
                <ul>
                  <li>=1 &gt; 0% and ≤ 1%</li>
                  <li>=2 &gt; 1% and ≤ 3%</li>
                  <li>=3 &gt; 3% and ≤ 5%</li>
                  <li>=4 &gt; 5% and ≤ 10%</li>
                  <li>=5 &gt; 10%</li>
                </ul>
              </div>
            </div>

            {/* Material Cost Box */}
            <div className="col-md-4">
              <div className="info-box">
                <div className="info-title">Material Cost</div>
                <ul>
                  <li>=1 ≤ 500 Dollars</li>
                  <li>=2 &gt; 500 and ≤ 1000 Dollars</li>
                  <li>=3 &gt; 1000 and ≤ 2500 Dollars</li>
                  <li>=4 &gt; 2500 and ≤ 5000 Dollars</li>
                  <li>=5 &gt; 5000 Dollars</li>
                </ul>
              </div>
            </div>

            {/* Seriousness Box */}
            <div className="col-md-4">
              <div className="info-box">
                <div className="info-title">Seriousness</div>
                <ul>
                  <li>=1 No perceived quality issue (C)</li>
                  <li>=2 Slight problem / disappointment (C)</li>
                  <li>=3 Clear problem to customer (B)</li>
                  <li>=4 Fit/function issue - Warranty (A)</li>
                  <li>=5 Safety / In-operation failure (AA)</li>
                </ul>
              </div>
            </div>

          </div>
        </div>

      </div>

    </div>

  );
};

export default QAMatrixForm;
