import * as React from "react";
import { useState, useEffect, useRef } from "react";
import { SPHttpClient } from "@microsoft/sp-http";
import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";
import { faEdit, faPlus } from "@fortawesome/free-solid-svg-icons";
import { highlightCurrentNav } from "../Utilities/HighlightCurrentComponent";
import { showToast } from "../Shared/Toaster";
import SearchableDropdown from "../Shared/Dropdown";
import { initCommonFunctions } from "../Utilities/CommonFunctions";
// import TableGenerator from "../Shared/TableGenerator";
import { NavLink } from "react-router-dom";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import InputCheckBox from "../Shared/InputCheckBox";
import { hideLoader, showLoader } from "../Shared/Loader";
import { spfi, SPFx } from "@pnp/sp";
import formValidation from "../Utilities/FormValidator";
import { ControlType } from "../Constants/Contants";
import AGGridDataTable from "../Shared/AGGridDataTable";

export interface LPAProps {
  spHttpClient: SPHttpClient;
  context: any;
  siteURL: string;
  spContext: any;
}

const LPA: React.FC<LPAProps> = (props) => {
  const sp = spfi().using(SPFx(props.context));
  const rootSiteURL = props.spContext.siteAbsoluteUrl;
  const JvisURL = `${rootSiteURL}/mayco`;
  const currentSiteURL = props.spContext.webAbsoluteUrl;

  const txtsubcategory = React.useRef<HTMLInputElement>(null);
  const auditCategoryRef = useRef<any>(null);


  // const [pageNumber, setPageNumber] = useState(1);
  const { getListItems } = initCommonFunctions(props.context, props.siteURL);
  const [categories, setCategories] = useState<any[]>([]);
  const [Plants, setPlants] = useState<any[]>([]);
  const [Departments, setDepartments] = useState<any[]>([]);
  const [isFormOpen, setIsFormOpen] = useState(false);
  const [LpaData, setLpaData] = useState<any>([]);
  const [itemId, setItemId] = useState(0);
  const [formData, setFormData] = useState({
    Plant: "",
    Department: "",
    Title: "",
    Audit_x0020_SubCategories: "",
    check: false,
    Is_x0020_Active: false
  });

  useEffect(() => {
    highlightCurrentNav("liPPETypes");
    showLoader();
    loadData();
  }, []);

  const loadData = async () => {
    try {
      const [plantData, deptData, LpaList] = await Promise.all([
        getListItems("Plant", JvisURL, "*", "", "Title eq 'Merrill'"),
        getListItems(
          "Department",
          JvisURL,
          "Plant/Title,Title",
          "Plant",
          "Plant/Title eq 'Merrill'"
        ),
        getListItems("LPA Configuration", currentSiteURL, "*")
      ]);

      setPlants(plantData);
      setDepartments(deptData);

      const normalizedLpaData = LpaList.map((item: any) => ({
        Id: item.Id,
        Plant: item.Plant || "",
        Department: item.Department || "",
        Title: item.Title,
        Audit_x0020_SubCategories: item.Audit_x0020_SubCategories,
        check: item.check,
        Is_x0020_Active: item.Is_x0020_Active
      }));

      setLpaData(normalizedLpaData);

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

  const handleChangeAuditcategory = (selectedOption: any, actionMeta: any) => {
    if (!actionMeta?.name) return;

    // const name = actionMeta.name;

    const value =
      actionMeta.action === "clear"
        ? ""
        : selectedOption?.value || "";   // ✅ use value, not label

    setFormData(prev => ({
      ...prev,
      Title: value
    }));
  };

  const handleChangeDropdown = (selectedOption: any, actionMeta: any) => {
    if (!actionMeta?.name) return;

    const name = actionMeta.name;
    const value =
      actionMeta.action === "clear" ? "" : selectedOption?.value || "";

    setFormData(prev => {
      let updated = {
        ...prev,
        [name]: value
      };

      // 🔥 KEY FIX
      if (name === "Department") {
        updated.Title = "";   // clear category
        updated.Audit_x0020_SubCategories = ""; // clear subcategory
        setCategories([]);
      }

      return updated;
    });

    if (name === "Department" && value) {
      fetchCategories(value);
    }
  };
  const fetchCategories = async (department: string) => {
    try {
      const data = await getListItems(
        "LPA Categories",
        currentSiteURL,
        "*",
        "",
        `Department eq '${department}'`
      );

      if (data.length > 0) {
        // Assuming categories field name is 'Categories'
        const raw = data[0].Categories || "";

        const parsed = raw.split("#;&@").filter((x: string) => x);

        const formatted = parsed.map((item: string) => ({
          label: item,
          value: item
        }));

        setCategories(formatted);
      } else {
        setCategories([]);
      }
    } catch (e) {
      console.log(e);
      showToast("error", "Error fetching categories");
    }
  };
  const addNew = () => {
    setIsFormOpen(true);
    setFormData({
      Plant: Plants.length > 0 ? Plants[0].Title : "",
      Department: "",
      Title: "",
      Audit_x0020_SubCategories: "",
      check: false,
      Is_x0020_Active: false
    });
    setTimeout(() => {
      const input = auditCategoryRef.current?.querySelector("input");
      input?.focus();
    }, 0);
  };

  const handleSubmit = async (event: any) => {

    event.preventDefault();

    showLoader();

    const data = {
      Department: { val: formData.Department.trim(), required: true, Name: "Department", Type: ControlType.reactSelect, Focusid: 'txtDepartment' },
      AuditCategoty: { val: formData.Title.trim(), required: true, Name: "Audit Category", Type: ControlType.reactSelect, Focusid: 'txtAuditCategory' },
      AuditSubCategoty: { val: formData.Audit_x0020_SubCategories.trim(), required: true, Name: "Plant", Type: ControlType.reactSelect, Focusid: txtsubcategory },

    };

    const isValid = formValidation.FormValidation(data);

    if (isValid.status) {

      const validDuplicate = await checkDuplicate();

      if (validDuplicate) insertOrUpdate();

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

  const checkDuplicate = async () => {
    try {
      showLoader();

      const plant = formData.Plant.trim();
      const department = formData.Department.trim();
      const category = formData.Title.trim();
      const subCategory = formData.Audit_x0020_SubCategories.trim();

      // Build filter query for multiple fields
      let filterQuery = `Plant eq '${plant}' and Department eq '${department}' and Title eq '${category}' and Audit_x0020_SubCategories eq '${subCategory}'`;

      // Exclude current item if editing
      if (itemId > 0) filterQuery += ` and Id ne ${itemId}`;

      // Query SharePoint
      const results = await sp.web.lists
        .getByTitle("LPA Configuration")
        .items.filter(filterQuery)();

      hideLoader();

      if (results.length > 0) {
        showToast("error", "Record with the same Plant, Department, Category, and SubCategory already exists");
        return false;
      }

      return true;

    } catch (e) {
      console.error(e);
      hideLoader();
      showToast("error", "Error checking duplicates");
      return false;
    }
  };

  const insertOrUpdate = async () => {
    try {
      const data = {
        Plant: formData.Plant,
        Department: formData.Department,
        Title: formData.Title,
        Audit_x0020_SubCategories: formData.Audit_x0020_SubCategories,
        check: formData.check,
        Is_x0020_Active: formData.Is_x0020_Active
      };

      if (itemId > 0) {
        await sp.web.lists
          .getByTitle("LPA Configuration")
          .items.getById(itemId)
          .update(data);

        showToast("success", "Item updated successfully");
      } else {
        await sp.web.lists
          .getByTitle("LPA Configuration")
          .items.add(data);

        showToast("success", "Item created successfully");
      }

      closeForm();
      loadData();

    } catch (e) {
      console.log(e);
      showToast("error", "Error saving data");
    } finally {
      hideLoader();
    }
  };


  const editItem = async (Id: number) => {
    setTimeout(() => {
  auditCategoryRef.current?.focus?.();
}, 100);
    try {

      showLoader();

      setIsFormOpen(true);
      setItemId(Id);
    
      const item = await sp.web.lists
        .getByTitle("LPA Configuration")
        .items.getById(Id)();

      setFormData({
        Plant: item.Plant || "",
        Department: item.Department || "",
        Title: item.Title || "",
        Audit_x0020_SubCategories: item.Audit_x0020_SubCategories || "",
        check: item.check || false,
        Is_x0020_Active: item.Is_x0020_Active || false
      });
      
      fetchCategories(item.Department);
          setTimeout(() => {
      const input = auditCategoryRef.current?.querySelector("input");
      input?.focus();
    }, 0);

    } catch (e) {
      console.log(e);
    } finally {
      hideLoader();
    }
  };


  const closeForm = () => {
    setIsFormOpen(false);
    setItemId(0);
  };

  // const columns = [
  //   {
  //     name: "Edit",
  //     selector: (row: any) => row.Id,
  //     cell: (record: any) => (
  //       <NavLink title="Edit" onClick={() => editItem(record.Id)} to="">
  //         <FontAwesomeIcon icon={faEdit} />
  //       </NavLink>
  //     ),
  //     width: "60px"
  //   },
  //   { name: "Plant", selector: (row: any) => row.Plant },
  //   { name: "Department", selector: (row: any) => row.Department },
  //   { name: "Audit Categories", selector: (row: any) => row.Title },
  //   {
  //     name: "Audit SubCategories",
  //     selector: (row: any) => row.Audit_x0020_SubCategories
  //   },
  //   {
  //     name: "Is Date Required",
  //     selector: (row: any) => (row.check ? "Yes" : "No")
  //   },
  //   {
  //     name: "Is Active",
  //     selector: (row: any) => (row.Is_x0020_Active ? "Yes" : "No")
  //   }
  // ];

  const columns = [
  {
    headerName: "Edit",
    field: "Id",
    width: 70,
    sortable: false,
    filter: false,
    cellRenderer: (params: any) => {
      const record = params.data;

      return (
        <NavLink
          title="Edit"
          className="csrLink ms-draggable"
          to=""
          onClick={(e) => {
            e.preventDefault();
            editItem(record.Id);
          }}
        >
          <FontAwesomeIcon icon={faEdit} />
        </NavLink>
      );
    },
  },

  {
    headerName: "Plant",
    field: "Plant",
    sortable: true,
    filter: "agTextColumnFilter",
    resizable: true,
    flex: 1,
  },

  {
    headerName: "Department",
    field: "Department",
    sortable: true,
    filter: "agTextColumnFilter",
    resizable: true,
    flex: 1,
  },

  {
    headerName: "Audit Categories",
    field: "Title",
    sortable: true,
    filter: "agTextColumnFilter",
    resizable: true,
    flex: 1,
  },

  {
    headerName: "Audit Sub-Categories",
    field: "Audit_x0020_SubCategories",
    sortable: true,
    filter: "agTextColumnFilter",
    resizable: true,
    flex: 1,
  },

  {
    headerName: "Is Date Required",
    field: "check",
    sortable: true,
    filter: "agTextColumnFilter",
    resizable: true,
    flex: 1,
    cellRenderer: (row: any) => (
        <div title={row.data.check ? "Yes" : "No"}>
        
          <span className="status-icon-text">{row.data.check ? "Yes" : "No"}
          </span>
        </div>
      ),
  },

  {
    headerName: "Is Active",
    field: "Is_x0020_Active",
    sortable: true,
    filter: "agTextColumnFilter",
    resizable: true,
    flex: 1,
      cellRenderer: (row: any) => (
        <div title={row.data.Is_x0020_Active ? "Yes" : "No"}>
        
          <span className="status-icon-text">{row.data.Is_x0020_Active ? "Yes" : "No"}
          </span>
        </div>
      ),
  },
];
  return (
    <div className="container-fluid">
      <div className="light-box border-box-shadow">
        <div className="div-form-title">
          <div className="form-title">​LPA ​Configuration​​​</div>

          {isFormOpen && (
            <span className="span-mandatory-text">
              <span className="text-danger">*</span> are mandatory fields
            </span>
          )}
        </div>

        <div className="p-2 mx-1 ViewTable">
          {!isFormOpen && (
            <div className="text-end my-2 mx-3">
              <button className="NewButton" onClick={addNew}>
                <FontAwesomeIcon icon={faPlus} /> New
              </button>
            </div>
          )}

          {isFormOpen && (
            <div className="form-border-box p-2 mx-1 my-2">

              {/* ROW 1 */}
              <div className="row">
                <div className="col-md-4">
                  <div className="light-text">
                    <label>Plant *</label>
                    <select className="form-control" value={formData.Plant} disabled>
                      {Plants.map((p: any) => (
                        <option key={p.Id} value={p.Title}>
                          {p.Title}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>

                <div className="col-md-4">
                  <div className="light-text">
                    <div ref={auditCategoryRef}>
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
                </div>

                <div className="col-md-4">
                  <div className="light-text">
                    <SearchableDropdown
                      label="Audit Category"
                      name="Title"
                      selectedValue={formData.Title}
                      OptionsList={categories}
                      OnChange={handleChangeAuditcategory}
                      isRequired={true} Title={"Audit Category"} id={"txtAuditCategory"} className={""} disabled={false} />
                  </div>
                </div>
              </div>

              {/* ROW 2 */}
              <div className="row mt-3">
                <div className="col-md-3">
                  <div className="light-text">
                    <label>Audit Sub-Category *</label>
                    <input
                      className="form-control"
                      type="text"
                      name="Audit_x0020_SubCategories"
                      id="txtSub"
                      ref={txtsubcategory}
                      value={formData.Audit_x0020_SubCategories}
                      onChange={handleChange}
                    />
                  </div>
                </div>

                <div className="col-md-3">
                  <InputCheckBox
                    label="Is Date Required"
                    name="check"
                    checked={formData.check}
                    onChange={(e: any) => setFormData(prev => ({
                      ...prev,
                      check: e.target.checked
                    }))} isdisable={false} isRequired={false} />
                </div>

                <div className="col-md-3">

                  <InputCheckBox
                    label="Is Active"
                    name="Is_x0020_Active"
                    checked={formData.Is_x0020_Active}
                    onChange={(e: any) => setFormData(prev => ({
                      ...prev,
                      Is_x0020_Active: e.target.checked
                    }))} isdisable={false} isRequired={false} />

                </div>
              </div>

              {/* BUTTONS */}
              <div className="text-center mt-3">
                <button className="btn btn-primary mx-2" onClick={handleSubmit}>Submit</button>
                <button className="btn btn-secondary" onClick={closeForm}>
                  Cancel
                </button>
              </div>
            </div>
          )}
              <div className="mx-2 table-head-1st-td right-search-table mb-3">
                    {/* <TableGenerator columns={columns} data={this.state.data} fileName={'Location2'} onRowClick={(row:any)=>this.onEditClickHandler(row.Id)} ></TableGenerator> */}
                    <AGGridDataTable
  data={LpaData}
  columns={columns}
  showExportExcel={false}
  showAddButton={false}
  customBtnClass="px-1 text-right"
  btnDivID=""
  btnSpanID=""
  btnTitle=""
  searchBoxLeft={true}
  onRowClicked={(event: any) => editItem(event.data.Id)}
  domLayout="normal"
  suppressColumnVirtualization={true}
  ensureDomOrder={true}
  suppressHorizontalScroll={false}
  suppressSizeToFit={true}
  suppressColumnHiding={true}
  suppressAutoSize={true}
  suppressColumnMoveAnimation={true}
  suppressMovableColumns={true}
/>
                  </div>
          {/* <TableGenerator
            columns={columns}
            data={LpaData}
            onChange={setPageNumber}
            prvPageNumber={pageNumber}
            onRowClick={(row: any) => editItem(row.Id)}
            prvDirection={false}
            fileName="KPI"
            className="sp-Datatable-hh"
            showPagination={true}
          /> */}
        </div>
      </div>
    </div>
  );
};

export default LPA;
