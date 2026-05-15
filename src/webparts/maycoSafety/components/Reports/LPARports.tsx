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
import DateUtilities from "../Utilities/DateUtilities";
import { format } from "date-fns";
import { useNavigate } from "react-router-dom";
// import { ControlType } from "../Constants/Contants";
import formValidation from "../Utilities/FormValidator";
import {
  Chart as ChartJS,
  CategoryScale,
  LinearScale,
  BarElement,
  Title,
  Tooltip,
  Legend
} from "chart.js";

import { Bar } from "react-chartjs-2";
import { ControlType } from "../Constants/Contants";

ChartJS.register(
  CategoryScale,
  LinearScale,
  BarElement,
  Title,
  Tooltip,
  Legend
);

export interface LPAReportProps {
  spHttpClient: SPHttpClient;
  context: any;
  siteURL: string;
  spContext: any;
}

const LPAReport: React.FC<LPAReportProps> = (props) => {
  //   const sp = spfi().using(SPFx(props.context));
  const navigate = useNavigate();
  const dateRef = React.useRef<any>(null);
  const rootSiteURL = props.spContext.siteAbsoluteUrl;
  const JvisURL = `${rootSiteURL}/mayco`;
  const currentSiteURL = props.spContext.webAbsoluteUrl;

  // const txtsubcategory = React.useRef<HTMLInputElement>(null);
  const { getListItems } = initCommonFunctions(props.context, props.siteURL);
  const [Plants, setPlants] = useState<any[]>([]);
  const [Departments, setDepartments] = useState<any[]>([]);
  const [Levels, setLevels] = useState<any[]>([]);
  const [allZones, setAllZones] = useState<any[]>([]);
  const [filteredZones, setFilteredZones] = useState<any[]>([]);
  const [weeklyData, setWeeklyData] = useState<any[]>([]);
  const [noData, setNoData] = useState(false);
  const [totalChartData, setTotalChartData] = useState<any[]>([]);
  const [lineChartData, setLineChartData] = useState<any[]>([]);
  const [issueDetails, setIssueDetails] = useState<any[]>([]);
  const [itemId, setItemId] = useState(0);
  const [formData, setFormData] = useState({
    Date: "",
    StartDate: "",
    Level: "All",
    EndDate: "",
    Title: "",
    Department: "",
    Zone_x0009_: "",
    Machine: "",
    ReportType: "LPAs by Auditor by Week",
    AuditorId: 0,
    Supervisor: "",
    Operators: "",
    Tool_x0020_Number: "",
    Comments: "",
    Year: new Date().getFullYear().toString(),
    YearMonth: (new Date().getMonth() + 1).toString()
  });


  useEffect(() => {
    highlightCurrentNav("liPPETypes");
    loadData();
  }, []);



  const chartData = {
    labels: totalChartData.slice(1).map((x: any[]) => x[0]),
    datasets: [
      {
        label: "Total LPAs",
        data: totalChartData.slice(1).map((x: any[]) => x[1]),
        backgroundColor: "#3366cc",
      }
    ]
  };
  const exportToCSV = (filename: string, rows: any[]) => {
    if (!rows || rows.length === 0) {
      showToast("error", "No data to export");
      return;
    }

    const separator = ",";
    const keys = Object.keys(rows[0]);

    const csvContent =
      keys.join(separator) +
      "\n" +
      rows
        .map(row =>
          keys
            .map(k => `"${row[k] ?? ""}"`)
            .join(separator)
        )
        .join("\n");

    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const link = document.createElement("a");

    link.href = URL.createObjectURL(blob);
    link.download = filename;
    link.click();
  };

  const loadData = async () => {
    try {
      const [plantData, deptData, ZoneData, levelData] = await Promise.all([
        getListItems("Plant", JvisURL, "*", "", "Title eq 'Merrill'"),
        getListItems(
          "Department",
          JvisURL,
          "Plant/Title,Title",
          "Plant",
          "Plant/Title eq 'Merrill' and IsActive eq 1"
        ),
        getListItems("Zones", JvisURL, "Plant/Title,Department/Title,Title", "Plant,Department", "Plant/Title eq 'Merrill'"),
        getListItems("LPA Auditor Level", currentSiteURL, "*", "", "")
      ]);
      setPlants(plantData);
      setDepartments(deptData);
      setAllZones(ZoneData);
      setFilteredZones([]);
      setLevels(levelData);





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
        (z: any) =>
              z.Department?.Title === value &&
              z.Plant?.Title === formData.Title
      );

      setFilteredZones(filteredZ);


      setFormData(prev => ({
        ...prev,
        Department: value,
        Zone_x0009_: "",
        Machine: ""
      }));
      setWeeklyData([]);
      setTotalChartData([]);
      setLineChartData([]);
      setIssueDetails([]);
      setNoData(false);
      return;
    }
    /* ------------------ REPORT TYPE ------------------ */

    if (name === "ReportType") {

      setFormData(prev => ({
        ...prev,
        ReportType: value,
        StartDate: "",
        EndDate: "",
        Date: "",
        Department: "",
        Zone_x0009_: "",

      }));
    }


    /* ------------------ ZONE ------------------ */

    if (name === "Zone_x0009_") {

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



  const handleSubmit = async (event: any) => {
    event.preventDefault();
    setWeeklyData([]);
    setTotalChartData([]);
    setLineChartData([]);
    setIssueDetails([]);
    setNoData(false);
    showLoader();

    const data = {

      ReportType: { val: formData.ReportType, required: true, Name: "Report Type", Type: ControlType.reactSelect, Focusid: 'txtReportType' },
      Date: { val: formData.Date, required: formData.ReportType == "LPAs by Auditor by Week", Name: "Date", Type: ControlType.date, Focusid: 'dtDate' },

      StartDate: { val: formData.StartDate, required: formData.ReportType !== "LPAs by Auditor by Week", Name: "Start Date", Type: ControlType.date, Focusid: 'dtStartDate' },
      EndDate: { val: formData.EndDate, required: formData.ReportType !== "LPAs by Auditor by Week", Name: "End Date", Type: ControlType.date, Focusid: 'dtEndDate' },
      DateCompare: {
        Type: ControlType.compareDates,
        startDate: formData.StartDate,
        endDate: formData.EndDate,
        startDateName: "Start Date",
        endDateName: "End Date",
        Focusid: "dtEndDate",
        required: formData.ReportType !== "LPAs by Auditor by Week"
      },
    };

    const isValid = formValidation.FormValidation(data);

    if (isValid.status) {

      if (formData.ReportType === "LPAs by Auditor by Week") {

        let start = formData.Date;
        let endDateObj = new Date(formData.Date);
        endDateObj.setDate(endDateObj.getDate() + 6);

        let end = format(endDateObj, "MM/dd/yyyy");

        await fetchLPAAuditWeekly(start, end);
      }
      else if (formData.ReportType === "LPAs by Auditor by Total") {

        const start = new Date(formData.StartDate);
        start.setHours(0, 0, 0, 0);

        const end = new Date(formData.EndDate);
        end.setHours(23, 59, 59, 999);

        await fetchLPAAuditTotal(start.toISOString(), end.toISOString());
      }

      else if (formData.ReportType === "LPAs Issues by Line Item") {

        fetchLPALineItemIssues();

      }
      else if (formData.ReportType === "LPAs Issues Details") {
        fetchLPAIssueDetails();
      }

    } else {

      showToast("error", isValid.message);
      hideLoader();

    }

  };
  const parseLocalDate = (dateInput: any) => {
    const date = new Date(dateInput);

    return new Date(Date.UTC(
      date.getFullYear(),
      date.getMonth(),
      date.getDate()
    ));
  };

  const fetchLPAAuditWeekly = async (startDate: string, endDate: string) => {
    try {
      showLoader();

      const start = parseLocalDate(startDate);
      start.setUTCHours(0, 0, 0, 0);

      const end = parseLocalDate(endDate);
      end.setUTCHours(23, 59, 59, 999);
      let filter = `Date ge datetime'${start.toISOString()}' and Date le datetime'${end.toISOString()}' and Title eq '${formData.Title}'`;

      if (formData.Department) {
        filter += ` and Department eq '${formData.Department}'`;
      }

      if (formData.Zone_x0009_) {
        filter += ` and Zone_x0009_ eq '${formData.Zone_x0009_}'`;
      }

      if (formData.Level && formData.Level.toLowerCase() !== "all") {
        filter += ` and Auditor/Level eq '${formData.Level}'`;
      }

      const url = `${currentSiteURL}/_api/web/lists/getbytitle('LPA')/items?$select=Date,ID,Auditor/Title,Auditor/Id,Auditor/Level&$expand=Auditor&$top=5000&$filter=${filter}`;

      const response = await props.spHttpClient.get(url, SPHttpClient.configurations.v1);
      const data = await response.json();

      if (data?.value?.length > 0) {
        const processed = processWeeklyData(data.value, startDate);
        setWeeklyData(processed);
        setNoData(processed.length === 0);
      } else {
        setWeeklyData([]);
        setNoData(true);
      }

    } catch (error) {
      console.error(error);
      showToast("error", "Error fetching report data");
    } finally {
      hideLoader();
    }
  };
  const exportWeeklyCSV = () => {
    const formatted = weeklyData.map((item: any) => ({
      "Auditor Name": item.name,
      "Day 1": item.days[0],
      "Day 2": item.days[1],
      "Day 3": item.days[2],
      "Day 4": item.days[3],
      "Day 5": item.days[4],
      "Day 6": item.days[5],
      "Day 7": item.days[6],
      Total: item.total
    }));

    exportToCSV("LPAs by Auditor by Week.csv", formatted);
  };
  const exportIssueDetailsCSV = () => {
    const formatted = issueDetails.map((item: any) => ({
      Date: item.date
        ? `'${DateUtilities.getDateMMDDYYYY(item.date.split("T")[0])}`
        : "",
      "Auditor Name": item.auditor,
      "Sub-Category": item.subCategory,
      Remarks: item.remarks
    }));

    exportToCSV("LPAs Issues Details.csv", formatted);
  };

  const processWeeklyData = (items: any[], startDate: string) => {

    const start = new Date(startDate);

    // Step 1: Initialize map per auditor
    const auditorMap = new Map<string, any>();

    items.forEach(item => {
      const date = item.Date?.split("T")[0];
      const auditor = item.Auditor?.Title;

      if (!date || !auditor) return;

      // Calculate day index (0–6)
      const current = new Date(date);
      const diffTime = current.getTime() - start.getTime();
      const dayIndex = Math.floor(diffTime / (1000 * 60 * 60 * 24));

      if (dayIndex < 0 || dayIndex > 6) return;

      if (!auditorMap.has(auditor)) {
        auditorMap.set(auditor, {
          name: auditor,
          days: [0, 0, 0, 0, 0, 0, 0],
          total: 0
        });
      }

      const auditorData = auditorMap.get(auditor);
      auditorData.days[dayIndex] += 1;
      auditorData.total += 1;
    });

    return Array.from(auditorMap.values());
  };


  const fetchLPAAuditTotal = async (startDate: string, endDate: string) => {
    try {
      showLoader();

      const start = new Date(startDate);
      start.setHours(0, 0, 0, 0);

      const end = new Date(endDate);
      end.setHours(23, 59, 59, 999);

      let filter = `
      Date ge datetime'${start.toISOString()}'
      and Date le datetime'${end.toISOString()}'
      and Title eq '${formData.Title}'
    `;

      if (formData.Department) {
        filter += ` and Department eq '${formData.Department}'`;
      }

      if (formData.Zone_x0009_) {
        filter += ` and Zone_x0009_ eq '${formData.Zone_x0009_}'`;
      }

      if (formData.Level && formData.Level.toLowerCase() !== "all") {
        filter += ` and Auditor/Level eq '${formData.Level}'`;
      }

      const url = `${currentSiteURL}/_api/web/lists/getbytitle('LPA')/items?$select=Auditor/Title&$expand=Auditor&$top=5000&$filter=${filter}`;

      const response = await props.spHttpClient.get(url, SPHttpClient.configurations.v1);
      const data = await response.json();

      if (data?.value && data?.value.length > 0) {

        const processed = processTotalData(data.value || []);
        const chartData = formatForChart(processed);
        setTotalChartData(chartData);
      }
      else {
        setNoData(true);
      }

    } catch (error) {
      console.error(error);
      showToast("error", "Error fetching report");
    } finally {
      hideLoader();
    }
  };

  const processTotalData = (items: any[]) => {

    const map = new Map<string, number>();

    items.forEach(item => {
      const auditor = item.Auditor?.Title;
      if (!auditor) return;

      map.set(auditor, (map.get(auditor) || 0) + 1);
    });

    return Array.from(map.entries()).map(([name, total]) => ({
      name,
      total
    }));
  };
  const formatForChart = (data: any[]) => {
    const chartData: any[] = [["Auditor", "Total"]];

    data.forEach(item => {
      chartData.push([item.name, item.total]);
    });

    return chartData;
  };


  const fetchLPALineItemIssues = async () => {
    try {
      showLoader();

      /* ---------------- STEP 1: GET LPA ITEMS ---------------- */

      const start = new Date(formData.StartDate);
      start.setHours(0, 0, 0, 0);

      const end = new Date(formData.EndDate);
      end.setHours(23, 59, 59, 999);

      let filter = `
      Date ge datetime'${start.toISOString()}'
      and Date le datetime'${end.toISOString()}'
      and Title eq '${formData.Title}'
    `;

      if (formData.Department) {
        filter += ` and Department eq '${formData.Department}'`;
      }

      if (formData.Zone_x0009_) {
        filter += ` and Zone_x0009_ eq '${formData.Zone_x0009_}'`;
      }

      const lpaUrl =
        `${currentSiteURL}/_api/web/lists/getbytitle('LPA')/items` +
        `?$select=Id` +
        `&$top=5000` +
        `&$filter=${filter}`;

      const lpaRes = await props.spHttpClient.get(lpaUrl, SPHttpClient.configurations.v1);
      const lpaData = await lpaRes.json();

      if (!lpaData.value || lpaData.value.length === 0) {
        setLineChartData([]);
        setNoData(true);
        return;
      }

      /* ---------------- STEP 2: COLLECT LPA IDs ---------------- */

      const lpaIds = lpaData.value.map((x: any) => x.Id);

      /* ---------------- STEP 3: GET LPALine ITEMS ---------------- */

      const lineUrl =
        `${currentSiteURL}/_api/web/lists/getbytitle('LPALine')/items` + `?$select=LPA_x0020_Subcategory,Title` + `&$top=5000`;

      const lineRes = await props.spHttpClient.get(lineUrl, SPHttpClient.configurations.v1);
      const lineData = await lineRes.json();

      if (!lineData.value || lineData.value.length === 0) {
        setLineChartData([]);
        setNoData(true);
        return;
      }

      /* ---------------- STEP 4: FILTER BY LPA ID ---------------- */

      const filteredLines = lineData.value.filter((item: any) =>
        lpaIds.includes(Number(item.Title))   // ⚠️ VERY IMPORTANT
      );

      /* ---------------- STEP 5: PROCESS DATA ---------------- */

      const processed = processLineItemData(filteredLines);

      setLineChartData(processed);
      setNoData(processed.length === 0);

    } catch (error) {
      console.error(error);
      showToast("error", "Error loading Line Item report");
    } finally {
      hideLoader();
    }
  };


  const processLineItemData = (items: any[]) => {
    const map = new Map<string, number>();

    items.forEach(item => {
      const key = item.LPA_x0020_Subcategory?.trim();

      if (!key) return;

      map.set(key, (map.get(key) || 0) + 1);
    });

    return Array.from(map.entries()).map(([name, count]) => ({
      name,
      count
    }));
  };


  const lineChartConfig = {
    labels: lineChartData.map((x: any) => x.name),
    datasets: [
      {
        label: "Issues",
        data: lineChartData.map((x: any) => x.count),
        backgroundColor: "#3366cc"
      }
    ]
  };
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

  const fetchLPAIssueDetails = async () => {
    try {
      showLoader();

      /* ---------------- STEP 1: GET LPA ITEMS ---------------- */

      const start = new Date(formData.StartDate);
      start.setHours(0, 0, 0, 0);

      const end = new Date(formData.EndDate);
      end.setHours(23, 59, 59, 999);

      let filter = `
      Date ge datetime'${start.toISOString()}'
      and Date le datetime'${end.toISOString()}'
      and Title eq '${formData.Title}'
    `;

      if (formData.Department) {
        filter += ` and Department eq '${formData.Department}'`;
      }

      if (formData.Zone_x0009_) {
        filter += ` and Zone_x0009_ eq '${formData.Zone_x0009_}'`;
      }

      const lpaUrl =
        `${currentSiteURL}/_api/web/lists/getbytitle('LPA')/items` +
        `?$select=Id` +
        `&$top=5000` +
        `&$filter=${filter}`;

      const lpaRes = await props.spHttpClient.get(lpaUrl, SPHttpClient.configurations.v1);
      const lpaData = await lpaRes.json();

      if (!lpaData.value || lpaData.value.length === 0) {
        setNoData(true);
        setIssueDetails([]);
        return;
      }

      const lpaIds = lpaData.value.map((x: any) => x.Id);

      /* ---------------- STEP 2: GET LPALine ITEMS ---------------- */

      const lineUrl =
        `${currentSiteURL}/_api/web/lists/getbytitle('LPALine')/items` +
        `?$select=Title,LPA_x0020_Subcategory,Remarks,Date,Status,Auditor/Title` +
        `&$expand=Auditor` +
        `&$top=5000`;

      const lineRes = await props.spHttpClient.get(lineUrl, SPHttpClient.configurations.v1);
      const lineData = await lineRes.json();

      if (!lineData.value || lineData.value.length === 0) {
        setNoData(true);
        setIssueDetails([]);
        return;
      }

      /* ---------------- STEP 3: FILTER CHILD ITEMS ---------------- */

      const filtered = lineData.value.filter((item: any) =>
        lpaIds.includes(Number(item.Title)) &&   // 🔥 Title = Parent Id
        item.Status === "No"                     // 🔥 Only issues
      );

      /* ---------------- STEP 4: FORMAT DATA ---------------- */

      const formatted = filtered.map((item: any) => ({
        subCategory: item.LPA_x0020_Subcategory,
        remarks: item.Remarks,
        auditor: item.Auditor?.Title || "",
        date: item.Date,
        status: item.Status
      }));

      setIssueDetails(formatted);
      setNoData(formatted.length === 0);

    } catch (error) {
      console.error(error);
      showToast("error", "Error loading Issue Details");
    } finally {
      hideLoader();
    }
  };

  const closeForm = () => {

    setItemId(0);
    navigate('/Home');
  };

  return (
    <div className="container-fluid">
      <div className="light-box border-box-shadow">
        <div className="div-form-title">
          <div className="form-title">LPA Reports</div>


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
                  <SearchableDropdown
                    label="Report Type"
                    name="ReportType"
                    selectedValue={formData.ReportType}
                    OptionsList={[
                      { label: "LPAs by Auditor by Week", value: "LPAs by Auditor by Week" },
                      { label: "LPAs by Auditor by Total", value: "LPAs by Auditor by Total" },
                      { label: "LPAs Issues by Line Item", value: "LPAs Issues by Line Item" },
                      { label: "LPAs Issues Details", value: "LPAs Issues Details" }
                    ]}
                    OnChange={handleChangeDropdown}
                    isRequired={true}
                    Title={"Report Type"}
                    id={"txtReportType"}
                    className={""}
                    disabled={false}
                  />
                </div>

              </div>


              <div className="col-md-3">
                <div className="light-text">
                  <label>Plant <span className="mandatoryhastrick">*</span></label>
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
                    isRequired={false} Title={"Department"} id={"txtDepartment"} className={""} disabled={false} />
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
                    isRequired={false} Title={"Zone"} id={"txtZone"} className={""} disabled={false} />
                </div>

              </div>
              {formData.ReportType == "LPAs by Auditor by Week" && (
                <div className="col-md-3">
                  <div className="light-text">
                    <label className="" htmlFor="dtDate"> Date <span className="mandatoryhastrick">*</span></label>
                    <div className="custom-datepicker" id="divDatefield">
                      <DatePickercontrol placeholder="MM/DD/YYYY" selectedDate={formData.Date} ref={dateRef} title={formData.Date} id='dtDate' startDate={undefined} endDate={new Date()} name="Date" onDatechange={(dateProps: any) => handleDateChange(dateProps[0], dateProps[2], "divDate", dateProps)} highlightDate={new Date()} showIcon />
                    </div>
                  </div>
                </div>
              )}
              {formData.ReportType !== "LPAs by Auditor by Week" && (
                <>
                  <div className="col-md-3">
                    <div className="light-text">
                      <label className="" htmlFor="dtStartDate"> Start Date <span className="mandatoryhastrick">*</span></label>
                      <div className="custom-datepicker" id="divStartDatefield">
                        <DatePickercontrol placeholder="MM/DD/YYYY" selectedDate={formData.StartDate} ref={dateRef} title={formData.Date} id='dtStartDate' startDate={undefined} endDate={new Date()} name="Date" onDatechange={(dateProps: any) => handleDateChange(dateProps[0], "StartDate", "divStartDate", dateProps)} highlightDate={new Date()} showIcon />
                      </div>
                    </div>
                  </div>
                  <div className="col-md-3">
                    <div className="light-text">
                      <label className="" htmlFor="dtEndDate">End Date <span className="mandatoryhastrick">*</span></label>
                      <div className="custom-datepicker" id="divEndDatefield">
                        <DatePickercontrol placeholder="MM/DD/YYYY" selectedDate={formData.EndDate} ref={dateRef} title={formData.Date} id='dtEndDate' startDate={undefined} endDate={new Date()} name="Date" onDatechange={(dateProps: any) => handleDateChange(dateProps[0], "EndDate", "divEndDate", dateProps)} highlightDate={new Date()} showIcon />
                      </div>
                    </div>
                  </div>
                </>
              )}
              {formData.ReportType !== "LPAs Issues by Line Item" && formData.ReportType !== "LPAs Issues Details" && (
                <div className="col-md-3">
                  <div className="light-text">
                    <SearchableDropdown
                      label="Level"
                      name="Level"
                      selectedValue={formData.Level}
                      OptionsList={[
                        { label: "All", value: "All" },
                        ...Levels.map(d => ({
                          label: d.Title,
                          value: d.Title
                        }))
                      ]}
                      OnChange={handleChangeDropdown}
                      isRequired={false} Title={"Level"} id={"txtLevel"} className={""} disabled={false} />
                  </div>

                </div>
              )}

            </div>
          </div>

          {/* ROW 2 */}

          {/* BUTTONS */}
          <div className="text-center mt-3">
            <button className="btn btn-primary mx-2" onClick={handleSubmit}>{itemId > 0 ? "Update" : "Submit"}</button>
            <button className="btn btn-secondary" onClick={closeForm}>
              Cancel
            </button>
          </div>
          {/* <div className="text-center mt-2">
  {formData.ReportType === "LPAs by Auditor by Week" && weeklyData.length > 0 && (
    <button className="btn btn-success btn-CSV mx-2" onClick={exportWeeklyCSV}>
      Export CSV
    </button>
  )}

  {formData.ReportType === "LPAs Issues Details" && issueDetails.length > 0 && (
    <button className="btn btn-success btn-CSV mx-2" onClick={exportIssueDetailsCSV}>
      Export CSV
    </button>
  )}
</div> */}
          {/* Auditor Days Table */}
          {/* NO DATA MESSAGE */}
          {noData && (
            <div className="text-center text-danger mt-3">
              No Data Found
            </div>
          )}
          {formData.ReportType === "LPAs by Auditor by Week" && weeklyData.length > 0 && (
            <div className="custom-box info-box p-3 my-2 position-relative">
              {formData.ReportType === "LPAs by Auditor by Week" && weeklyData.length > 0 && (
                <button type="button" className="btn btn-success btn-CSV mx-2" onClick={exportWeeklyCSV}>
                  Export CSV
                </button>
              )}
              <h5 className="fw-bold mb-4">Weekly Auditor Report</h5>

              <table className="table table-bordered mb-0">
                <thead>
                  <tr>
                    <th>Auditor Name</th>
                    <th>Day 1</th>
                    <th>Day 2</th>
                    <th>Day 3</th>
                    <th>Day 4</th>
                    <th>Day 5</th>
                    <th>Day 6</th>
                    <th>Day 7</th>
                    <th>Total</th>
                  </tr>
                </thead>

                <tbody>
                  {weeklyData.map((item, index) => (
                    <tr key={index}>
                      <td>{item.name}</td>

                      {item.days.map((d: number, i: number) => (
                        <td key={i}>{d}</td>
                      ))}

                      <td><b>{item.total}</b></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}

          {/* Auditor Total Table */}
          {formData.ReportType === "LPAs by Auditor by Total" && totalChartData.length > 1 && (
            <div className="outer-report mt-3">
              <h5>Auditor Total Report</h5>

              <Bar
                data={chartData}
                options={{
                  responsive: true,
                  datasets: {
                    bar: {
                      barThickness: 25,
                      maxBarThickness: 30
                    }
                  },
                  plugins: {
                    legend: {
                      display: true
                    }
                  },
                  scales: {
                    x: {
                      ticks: {
                        maxRotation: 45,
                        minRotation: 45
                      }
                    },
                    y: {
                      beginAtZero: true
                    }
                  }


                }}
              />
            </div>
          )}

          {formData.ReportType === "LPAs Issues by Line Item" && lineChartData.length > 1 && (
            <div className="outer-report mt-3">
              <h5>Line Item Issues Report</h5>

              <Bar
                data={lineChartConfig}
                options={{
                  responsive: true,
                  plugins: {
                    legend: {
                      display: true
                    }
                  },
                  scales: {
                    x: {
                      ticks: {
                        maxRotation: 45,
                        minRotation: 45
                      }
                    },
                    y: {
                      beginAtZero: true
                    }
                  }
                }}
              />
            </div>
          )}

          {formData.ReportType === "LPAs Issues Details" && issueDetails.length > 0 && (
            <div className="custom-box info-box p-3 my-2 position-relative">
              {formData.ReportType === "LPAs Issues Details" && issueDetails.length > 0 && (
                <button type="button" className="btn btn-success btn-CSV mx-2" onClick={exportIssueDetailsCSV}>
                  Export CSV
                </button>
              )}
              <h5 className="fw-bold mb-4">LPA Issue Details</h5>

              <table className="table table-bordered mb-0">
                <thead>
                  <tr>
                    <th>Date</th>
                    <th>Auditor Name</th>
                    <th>Sub-Category</th>
                    <th>Action Taken</th>


                    {/* <th>Status</th> */}
                  </tr>
                </thead>

                <tbody>
                  {issueDetails.map((item, i) => (
                    <tr key={i}>
                      <td>
                        {item.date
                          ? DateUtilities.getDateMMDDYYYY(item.date.split("T")[0])
                          : ""}
                      </td>
                      <td>{item.auditor}</td>
                      <td>{item.subCategory}</td>
                      <td>{item.remarks}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}

        </div>

      </div>
    </div>

  );
};

export default LPAReport;
