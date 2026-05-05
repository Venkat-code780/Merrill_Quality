import * as React from "react";
import { useState, useEffect, useRef } from "react";
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
// import { format } from "date-fns";
import { useParams,useNavigate  } from "react-router-dom";



export interface LPAFormProps {
  spHttpClient: SPHttpClient;
  context: any;
  siteURL: string;
  spContext: any;
}

const LPAForm: React.FC<LPAFormProps> = (props) => {
  const { id } = useParams<{ id: string }>();
  const sp = spfi().using(SPFx(props.context));
  const navigate=useNavigate();
  const txtPPEType = useRef<HTMLInputElement>(null);
  const dateRef = React.useRef<any>(null);
  const rootSiteURL = props.spContext.siteAbsoluteUrl;
  const JvisURL = `${rootSiteURL}/mayco`;
  const currentSiteURL = props.spContext.webAbsoluteUrl;
  // const txtsubcategory = React.useRef<HTMLInputElement>(null);
  const { getListItems } = initCommonFunctions(props.context, props.siteURL);
 const [lineItems, setLineItems] = useState<any[]>([]);
  const [Plants, setPlants] = useState<any[]>([]);
  const [Departments, setDepartments] = useState<any[]>([]);
 const [allZones, setAllZones] = useState<any[]>([]);
 const [filteredZones, setFilteredZones] = useState<any[]>([]);
  const [allMachines, setAllMachines] = useState<any[]>([]);
   const [filteredMachines, setFilteredMachines] = useState<any[]>([]);
   const [auditors,setAuditors]=useState<any[]>([]);
 const [supervisors,setsupervisors]=useState<any[]>([]);
  const [toolnumbers,settoolnumbers]=useState<any[]>([]);
   const [categories, setCategories] = useState<any>({});
  const [itemId, setItemId] = useState(0);
  const [formData, setFormData] = useState({
    Date:"",
    Title: "",
    Department: "",
    Zone_x0009_: "",
    Machine:"",
    AuditorId:0,
    Supervisor:"",
    Operators:"",
    Tool_x0020_Number:"",
    Comments:"",
    Year:new Date().getFullYear().toString(),
    YearMonth:(new Date().getMonth()+1).toString().padStart(2,'0')
  });

  // useEffect(() => {
  //   highlightCurrentNav("liPPETypes");
  //   loadData();
  //   if(id){
  //     const numericId = parseInt(id);
  //     setItemId(numericId);
  //     loadDataForEdit(numericId);
  //   }
  // }, []);

  useEffect(() => {
  highlightCurrentNav("liPPETypes");
  showLoader();
  loadData();
   setTimeout(() => {
      const input = dateRef.current?.querySelector("input");
      input?.focus();
    }, 0);

}, []);

useEffect(() => {
  if (id && allZones.length > 0 && allMachines.length > 0) {
    const numericId = parseInt(id);
    setItemId(numericId);
    loadDataForEdit(numericId);
  }
}, [id, allZones, allMachines]);

 const loadDataForEdit = async (editId: number) => {
    try {
      // showLoader();
      // Fetch header item
      const item: any = await sp.web.lists.getByTitle("LPA").items.getById(editId)();
      // Populate form data
      setFormData({
        Date: item.Date,
        Title: item.Title || "",
        Department: item.Department || "",
        Zone_x0009_: item.Zone_x0009_ || "",
        Machine: item.Machine || "",
        AuditorId: item.AuditorId || 0,
        Supervisor: item.Supervisor || "",
        Operators: item.Operators || "",
        Tool_x0020_Number: item.Tool_x0020_Number || "",
        Comments: item.Comments || "",
        Year: item.Year || new Date().getFullYear().toString(),
        YearMonth: item.YearMonth || (new Date().getMonth() + 1).toString(),
      });
        await getCategories(item.Title, item.Department);

      // 🔥 Populate Zones based on Department
      const filteredZ = allZones.filter(
        (z: any) => z.Department?.Title === item.Department
      );

      setFilteredZones(filteredZ);

      // 🔥 Populate Machines based on Zone
      const filteredM = allMachines.filter(
        (m: any) => m.Zone?.Title === item.Zone_x0009_
      );

setFilteredMachines(filteredM);
      // Fetch line items
      const lineItemsData: any[] = await sp.web.lists
        .getByTitle("LPALine")
        .items.filter(`Title eq '${editId}'`)();
      
      const initialLines = lineItemsData.map(line => ({
        Id: line.Id,
        Category: line.LPA_x0020_Category,
        SubCategory: line.LPA_x0020_Subcategory,
        Result: line.Status,
        Action: line.Remarks,
        // Date: line.Date ? line.Date : "",
        Date: line.ChildDate ? new Date(line.ChildDate) : null
      }));

      setLineItems(initialLines);
    } catch (e) {
      console.error(e);
      showToast("error", "Error loading item for edit");
    } finally {
      hideLoader();
    }
  };



  const loadData = async () => {
    try {
      const [plantData, deptData, ZoneData, MachineData, AuditorsData, supervisorsData,toolnumbersdata] = await Promise.all([
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
        getListItems("LPA Auditors", currentSiteURL, "Title,Id", "", "Is_x0020_Active eq 1"),
        getListItems("Supervisor", JvisURL, "Plant/Title,Title", "Plant", "Plant/Title eq 'Merrill'"),
        getListItems("Tool Numbers",JvisURL,"Plant/Title,Title","Plant","Plant/Title eq 'Merrill'")
      ]);
      setPlants(plantData);
      setDepartments(deptData);
      setAllZones(ZoneData);
      setFilteredZones([]);
      setAllMachines(MachineData);
      setFilteredMachines([]);
      setAuditors(AuditorsData);
      setsupervisors(supervisorsData);
      settoolnumbers(toolnumbersdata);



      if (plantData.length > 0) {
        setFormData(prev => ({
          ...prev,
          Title: plantData[0].Title
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

const groupByCategory = (data: any[]) => {
  const result: any = {};

  data.forEach(item => {
    const category = item["Title"]; // check exact name

    if (!result[category]) {
      result[category] = [];
    }

    result[category].push(item);
  });

  return result;
};
const getCategories = async (plant: string, department: string) => {
  try {
    showLoader();
  const filter = `Plant eq '${plant}' and Department eq '${department}' and Is_x0020_Active eq 1`;
    const items = await getListItems(
      "LPA Configuration", // 🔴 change this
      currentSiteURL,
      "*",
      "",
      filter
    );

    const grouped = groupByCategory(items);

    setCategories(grouped);
   
    const initialLines: any[] = [];

items.forEach((item: any) => {
  initialLines.push({
    Category: item.Title,
    SubCategory: item["Audit_x0020_SubCategories"],
    IsDateRequired: item["check"],
    Result: "",
    Date: "",
    Action: ""
  });
});

setLineItems(initialLines);
  } catch (e) {
    console.log(e);
    showToast("error", "Error loading categories");
  } finally {
    hideLoader();
  }
};
const handleLineChange = (index: number, field: string, value: any) => {
  const updated = [...lineItems];
  updated[index][field] = value;
  setLineItems(updated);
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

    setCategories({}); // clear old categories

    setFormData(prev => ({
      ...prev,
      Department: value,
      Zone_x0009_: "",
      Machine: ""
    }));

    if (value) {
      getCategories(formData.Title, value);
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

  /* ------------------ AUDITOR ------------------ */

  if (name === "AuditorId") {
    setFormData(prev => ({
      ...prev,
      AuditorId: value   // ✅ number (lookup id)
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
const validateLineItems = () => {

  // clear old highlights
  document.querySelectorAll(".mandatory-FormContent-focus").forEach(el => {
    el.classList.remove("mandatory-FormContent-focus");
  });

  for (let i = 0; i < lineItems.length; i++) {

    const line = lineItems[i];

    if (line.IsDateRequired && !line.Date) {

      showToast(
        "error",
        `'Date for ${line.SubCategory} under ${line.Category}' cannot be blank`
      );

      const parent = document.getElementById(`lineDate_${i}`);

      if (parent) {

        // 🔥 find actual visible element
        const input = parent.querySelector("input");
        const wrapper = parent.querySelector(".react-datepicker-wrapper");

        input?.focus();

        // apply class to correct element
        (wrapper || parent).classList.add("focus-Div");
      }

      return false;
    }

if (line.Date) {

  const selectedDate = new Date(line.Date);
  const today = formData.Date ? new Date(formData.Date) : new Date();

  selectedDate.setHours(0, 0, 0, 0);
  today.setHours(0, 0, 0, 0);

  if (selectedDate.getTime() > today.getTime()) {

    showToast(
      "error",
      `Date ${line.SubCategory} must be less than or equal to Date`
    );

    const parent = document.getElementById(`lineDate_${i}`);

    if (parent) {
      const input = parent.querySelector("input");
      input?.focus();
      parent.classList.add("focus-Div");
    }

    return false;
  }
}

  if (line.Result === "No" && !line.Action?.trim()) {

  showToast(
    "error",
    `'Action Taken for ${line.SubCategory} under ${line.Category}' cannot be blank`
  );

const input = document.querySelector(
  `input[data-index="${i}"]`
) as HTMLInputElement;

  if (input) {
    input.focus();
    input.classList.add("mandatory-FormContent-focus");
    input.scrollIntoView({ behavior: "smooth", block: "center" });
  }

  return false;
}
    
  }

  return true;
};


  const handleSubmit = async (event: any) => {

    event.preventDefault();

    showLoader();

    const data = {
           Date:  { val: formData.Date, required: true, Name: "Date", Type: ControlType.date, Focusid: 'dtDate' },
      Department: { val: formData.Department.trim(), required: true, Name: "Department", Type: ControlType.reactSelect, Focusid: 'txtDepartment' },
            Zone: { val: formData.Zone_x0009_.trim(), required: true, Name: "Zone", Type: ControlType.reactSelect, Focusid: 'txtZone' },
         Machine: { val: formData.Machine.trim(), required: true, Name: "Machine", Type: ControlType.reactSelect, Focusid: 'txtMachine' },
    AuditorsName: { val: formData.AuditorId, required: true, Name: "Auditor", Type: ControlType.reactSelect, Focusid: 'txtAuditorName' },
     Supervisor:  { val: formData.Supervisor.trim(), required: true, Name: "Supervisor", Type: ControlType.reactSelect, Focusid: 'txtSupervisorName' },
     operators:  { val: formData.Operators.trim(), required: true, Name: "Operators", Type: ControlType.string, Focusid: txtPPEType },
    //  ToolNumber:  { val: formData.Tool_x0020_Number.trim(), required: true, Name: "Tool Number", Type: ControlType.reactSelect, Focusid: 'txtToolNumber' },

    };

    const isValid = formValidation.FormValidation(data);

    if (isValid.status) {
      const isLineValid = validateLineItems(); 
       if (!isLineValid) {
    hideLoader();
    return;
  }
      insertOrUpdate();

      // const validDuplicate = await checkDuplicate();

      // if (validDuplicate) insertOrUpdate();

    } else {

      showToast("error", isValid.message);
      hideLoader();

    }

  };
const handleChange = (e: any) => {
  const { name, value } = e.target;

  setFormData(prev => ({
    ...prev,
    [name]: value
  }));
};
// const handleDateChange = (

//     dateValue: any,
//     name: string,
//     divId: string,
//     dateProps?: any
// ) => {

//     setFormData((prev: any) => {

//         const updatedFormData = { ...prev };

//         // Handle floating label active class
//         if (divId) {
//             const ddlElement = document.getElementById(divId);

//             if (dateValue) {
//                 ddlElement?.classList.add("active");
//             } else {
//                 ddlElement?.classList.remove("active");
//             }
//         }
  
//       // Date conversion
//         let formattedDate = "";

//         if (dateValue != null) {
//             formattedDate = format(
//                 DateUtilities.addBrowserwrtServer(
//                     new Date(DateUtilities.getDateMMDDYYYY(dateValue)),
//                     props.spContext.webTimeZoneData
//                 ).toISOString(),
//                 "MM/dd/yyyy"
//             );
//         }

//         updatedFormData[name] = formattedDate;

//         return updatedFormData;
//     });
// };


const handleDateChange = (dateValue: any, name: string) => {
  setFormData((prev) => ({
    ...prev,
    [name]: dateValue ? new Date(dateValue) : null
  }));
};
const toDateOnlyISO = (date: any) => {
    if (!date) return null;

  const d = new Date(date);

  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  // if (!date) return null;

  // const d = new Date(date);
  // return new Date(d.getFullYear(), d.getMonth(), d.getDate(), 12, 0, 0);
};


const handleLineDateChange = (index: number, dateValue: any) => {

  let isoDate = "";

  if (dateValue) {
    isoDate = DateUtilities.addBrowserwrtServer(
      new Date(dateValue),
      props.spContext.webTimeZoneData
    ).toISOString();   // ✅ STORE ISO DIRECTLY
  }

  const updated = [...lineItems];
  updated[index].Date = isoDate;

  setLineItems(updated);
};


const insertOrUpdate = async () => {
  try {
    showLoader();

    /* ---------------- HEADER DATA ---------------- */

    const mainData: any = {
      Title: formData.Title,
      Department: formData.Department,
      Zone_x0009_: formData.Zone_x0009_,
      Machine: formData.Machine,

      AuditorId: formData.AuditorId,

      Supervisor: formData.Supervisor,

      Tool_x0020_Number: formData.Tool_x0020_Number,
      Comments: formData.Comments,
      Date: toDateOnlyISO(formData.Date),
      Operators: formData.Operators,
      Year: formData.Year,
      YearMonth: formData.YearMonth
    };


    let headerId = itemId;

    /* ---------------- HEADER SAVE ---------------- */

    if (itemId > 0) {
      await sp.web.lists
        .getByTitle("LPA")
        .items.getById(itemId)
        .update(mainData);

      // showToast("success", "Item updated successfully");
    } else {
      const res = await sp.web.lists
        .getByTitle("LPA")
        .items.add(mainData);

      headerId = res.Id;

      // showToast("success", "Item created successfully");
    }

    /* ---------------- CHILD UPSERT ---------------- */

    // const [batchedSP, execute] = sp.batched();

    // lineItems.forEach((line: any) => {

    
    //   if (!line.Result || line.Result === "") return;

    //   const childData: any = {

  
    //     Title: headerId.toString(),

       
    //     Plant: formData.Title,
    //     Department: formData.Department,
    //     Zone: formData.Zone_x0009_,
    //     Machine: formData.Machine,

       
    //     AuditorId: formData.AuditorId,


    //     // 🔥 LINE DATA
  
    //     LPA_x0020_Category: line.Category,
    //     LPA_x0020_Subcategory: line.SubCategory,
    //     Status: line.Result,
    //     Remarks: line.Action,
    //      Date: line.Date ? line.Date : null
    //   };

    //   if (line.Id && line.Id > 0) {
    //     batchedSP.web.lists
    //       .getByTitle("LPALine")
    //       .items.getById(line.Id)
    //       .update(childData);
    //   }

    //   else {
    //     batchedSP.web.lists
    //       .getByTitle("LPALine")
    //       .items.add(childData);
    //   }

    // });

    // await execute();




    for (const line of lineItems) {
  const childData = {
    Title: headerId.toString(),
    Plant: formData.Title,
    Department: formData.Department,
    Zone: formData.Zone_x0009_,
    Machine: formData.Machine,
    AuditorId: formData.AuditorId,
    LPA_x0020_Category: line.Category,
    LPA_x0020_Subcategory: line.SubCategory,
    Status: line.Result,
    Remarks: line.Action || "",
     Date: formData.Date
    ? new Date(formData.Date).toISOString()
    : null,
    ChildDate: line.Date ? new Date(line.Date).toISOString() : null 
  };

  try {

    if (line.Id && line.Id > 0) {
      // ✅ UPDATE
      console.log("Updating line:", childData);

      await sp.web.lists
        .getByTitle("LPALine")
        .items.getById(line.Id)
        .update(childData);

    } else {
      // ✅ INSERT
      // console.log("Inserting line:", childData);

      await sp.web.lists
        .getByTitle("LPALine")
        .items.add(childData);
    }

  } catch (err) {
    console.error("❌ CHILD SAVE ERROR:", err);
  }
}
     showToast("success", itemId > 0 ? "LPA updated successfully" : "LPA created successfully");
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
    navigate('/LPAView');
  };

  return (
    <div className="container-fluid">
      <div className="light-box border-box-shadow">
        <div className="div-form-title">
          <div className="form-title">Layered Process Audit {itemId > 0 ? - itemId:""}</div>


            <span className="span-mandatory-text">
              <span className="text-danger">*</span> are mandatory fields
            </span>
        </div>

        <div className="p-2 mx-1 ViewTable">


              {/* ROW 1 */}
              <div className="row">
                      <div className="col-md-3">
                                                <div className="light-text">
                                                     <div ref={dateRef}>
                                                    <label className="" htmlFor="dtDate"> Date <span className="text-danger">*</span></label>
                                                    <div className="custom-datepicker" id="divDatefield">
                                                        <DatePickercontrol placeholder="MM/DD/YYYY" selectedDate={formData.Date} ref={dateRef} title={formData.Date} id='dtDate' startDate={undefined} endDate={new Date()} name="Date" onDatechange={(dateProps: any) => handleDateChange(dateProps[0],"Date")} highlightDate={new Date()} showIcon />
                                                    </div>
                                                    </div>
                                                </div>
                                            </div>
                <div className="col-md-3">
                   <div className="light-text">
                  <label>Plant <span className="text-danger">*</span></label>
                  <select className="form-control" value={formData.Title} disabled>
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
                    isRequired={true} Title={"Department"} id={"txtDepartment"} className={""} disabled={itemId > 0} />
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
                    isRequired={true} Title={"Zone"} id={"txtZone"} className={""} disabled={itemId > 0} />
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
                    isRequired={true} Title={"Machine"} id={"txtMachine"} className={""} disabled={itemId>0} />
                </div>

              </div>
                  <div className="col-md-3">
                   <div className="light-text">
                  <SearchableDropdown
                    label="Auditor's Name"
                    name="AuditorId"
                    selectedValue={formData.AuditorId}
                    OptionsList={auditors.map(d => ({
                      label: d.Title,
                      value: d.Id
                    }))}
                    OnChange={handleChangeDropdown}
                    isRequired={true} Title={"AuditorName"} id={"txtAuditorName"} className={""} disabled={false} />
                </div>

              </div>
               <div className="col-md-3">
                   <div className="light-text">
                  <SearchableDropdown
                    label="Supervisor's Name"
                    name="Supervisor"
                    selectedValue={formData.Supervisor}
                    OptionsList={supervisors.map(d => ({
                      label: d.Title,
                      value: d.Title
                    }))}
                    OnChange={handleChangeDropdown}
                    isRequired={true} Title={"SupervisorName"} id={"txtSupervisorName"} className={""} disabled={false} />
                </div>

              </div>
               <div className="col-md-3">

                  <div className="light-text">

                    <input
                      className="form-control"
                      type="text"
                      name="Operators"
                      value={formData.Operators}
                      onChange={handleChange}
                      ref={txtPPEType}
                      maxLength={250}
                    />

                    <label>
                      Operator(s) <span className="mandatoryhastrick">*</span>
                    </label>

                  </div>

                </div>
                <div className="col-md-3">
                   <div className="light-text">
                  <SearchableDropdown
                    label="Tool Number"
                    name="Tool_x0020_Number"
                    selectedValue={formData.Tool_x0020_Number}
                    OptionsList={toolnumbers.map(d => ({
                      label: d.Title,
                      value: d.Title
                    }))}
                    OnChange={handleChangeDropdown}
                    isRequired={false} Title={"ToolNumber"} id={"txtToolNumber"} className={""} disabled={false} />
                </div>

              </div>
                 
           
              </div>
         
           {Object.keys(categories).length > 0 && (
            <div className="col-12 select-appearance">
  <table className="table-bordered mt-3 w-100">
     
    <thead className="table-dark">
      <tr>
        <th className="py-2">Checks</th>
        <th className="py-2">Date</th>
        <th className="py-2"></th>
        <th className="py-2">Action Taken</th>
      </tr>
    </thead>

    <tbody>
      {Object.keys(categories).map((cat, index) => (
        <React.Fragment key={index}>

          {/* CATEGORY ROW */}
          <tr className="thead-bold">
            <td colSpan={4}>{cat}</td>
          </tr>

          {/* SUBCATEGORY ROWS */}
          {categories[cat].map((item: any, i: number) => {

            const globalIndex = lineItems.findIndex(
              (x) =>
                x.Category === cat &&
                x.SubCategory === item["Audit_x0020_SubCategories"]
            );

            return (
              <tr key={i}>
                <td>{item["Audit_x0020_SubCategories"]}</td>

                {/* DATE */}
                <td>
                  {item["check"] && (
                    <div className="light-text">
                      <label className="" htmlFor="dtDate"> Date <span className="text-danger">*</span></label>
                      <div className="custom-datepicker" id={`lineDate_${globalIndex}`}>
                        <DatePickercontrol placeholder="MM/DD/YYYY" selectedDate={
                          lineItems[globalIndex]?.Date
                            ? new Date(lineItems[globalIndex].Date).toISOString()
                            : undefined
                        } title={formData.Date} id='dtDate' startDate={undefined} endDate={new Date()} name="Date" onDatechange={(dateProps: any) => {
                          if (globalIndex === -1) return;
                          handleLineDateChange(globalIndex, dateProps[0]);
                        }} highlightDate={new Date()} showIcon />
                      </div>
                    </div>
                   
                  )}
                </td>

                {/* RESULT */}
                <td>
                  <select
                    className="form-control"
                    value={lineItems[globalIndex]?.Result || ""}
                    onChange={(e) =>
                      handleLineChange(globalIndex, "Result", e.target.value)
                    }
                  >
                    <option value="N/A">N/A</option>
                    <option value="Yes">Yes</option>
                    <option value="No">No</option>
                    
                  </select>
                </td>

                {/* ACTION */}
                <td>
                   <div className="light-text">

    <label>
      Action Taken
      {lineItems[globalIndex]?.Result === "No" && (
        <span className="text-danger"> *</span>
      )}
    </label>
                  <input
                    type="text"
                    id={`actionCell_${globalIndex}`}
                    className="form-control"
                    data-index={globalIndex}
                    placeholder="Action Taken"
                    value={lineItems[globalIndex]?.Action || ""}
                    onChange={(e) =>
                      handleLineChange(globalIndex, "Action", e.target.value)
                    }
                  />
                  </div>
                </td>
              </tr>
            );
          })}
        </React.Fragment>
      ))}
    </tbody>
  </table>
  </div>
)}
              {/* ROW 2 */}
      <div className="col-md-12">
  <div className="light-text">

        <textarea
      className="form-control"
      name="KPI"
      value={formData.Comments}
      onChange={handleChange}
      maxLength={250}
    />

    <label>
      Comments
    </label>

  </div>
</div>

              {/* BUTTONS */}
              <div className="text-center mt-3">
                <button className="btn btn-primary mx-2" onClick={handleSubmit}>{itemId >0 ? "Update":"Submit"}</button>
                <button className="btn btn-secondary" onClick={closeForm}>
                  Cancel
                </button>
              </div>
            </div>
  
        </div>
      </div>
   
  );
};

export default LPAForm;
