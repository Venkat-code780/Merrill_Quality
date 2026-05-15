import * as React from "react";
import { useState, useEffect, useRef } from "react";
import { SPHttpClient } from "@microsoft/sp-http";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
// import TableGenerator from "../Shared/TableGenerator";
import { hideLoader, showLoader } from "../Shared/Loader";
import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";
import { faEdit, faPlus } from "@fortawesome/free-solid-svg-icons";
import { highlightCurrentNav } from "../Utilities/HighlightCurrentComponent";
import { ActionStatus, ControlType } from "../Constants/Contants";
import formValidation from "../Utilities/FormValidator";
import { showToast } from "../Shared/Toaster";
import { NavLink } from "react-router-dom";
import SearchableDropdown from "../Shared/Dropdown";
import { initCommonFunctions } from "../Utilities/CommonFunctions";
import InputCheckBox from "../Shared/InputCheckBox";
import AGGridDataTable from "../Shared/AGGridDataTable";
export interface LPAAuditorsProps {
  match: any;
  spHttpClient: SPHttpClient;
  context: any;
  siteURL: string;
  spContext: any;
  webAbsoluteURL: string;
  currPlantTitle: string;

}

const LPAAuditors: React.FC<LPAAuditorsProps> = (props) => {
  const rootSiteURL = props.spContext.siteAbsoluteUrl;
  const currentSiteURL = props.spContext.webAbsoluteUrl;
  const JvisURL = `${rootSiteURL}/mayco`;
  const { getListItems } = initCommonFunctions(props.context, props.siteURL);


  const listName = "LPA Auditors";


  const sp = spfi().using(SPFx(props.context));
  // const txtPPEType = useRef<HTMLInputElement>(null);
  const auditCategoryRef = useRef<any>(null);

  const [actionsData, setActionsData] = useState<any[]>([]);
  // const [pageNumber, setPageNumber] = useState(1);
  const [isFormOpen, setIsFormOpen] = useState(false);
  const [itemId, setItemId] = useState(0);
  const [redirect, setRedirect] = useState(false);
  // const [displayMessage, setDisplayMessage] = useState("");
  const [Auditors, setAuditors] = useState<any[]>([]);
  const [AuditLevels, setAuditLevels] = useState<any[]>([]);
  const [formData, setFormData] = useState({
    Title: "",
    Level: "",

    Is_x0020_Active: false
  });

  // Load initial data
  useEffect(() => {
    highlightCurrentNav("liPPETypes");
    document.title = "JVIS - Quality | LPA Auditor Levels";
    loadListData();
    loadAuditorsLevels();
  }, []);

  useEffect(() => {
    if (redirect) loadListData();
  }, [redirect]);

  // Load Auditors filtered by Plant 'Shelby' and Audit Levels
  const loadAuditorsLevels = async () => {
    try {
      showLoader();

      // Create SP instance for the root site

      const [auditors, auditLevels] = await Promise.all([
        getListItems("Auditors", JvisURL, "Title,Plant/Title", "Plant", "Plant/Title eq 'Merrill'"),
        getListItems("LPA Auditor Level", currentSiteURL)
      ]);
      setAuditors(auditors);
      setAuditLevels(auditLevels);
    } catch (e) {
      console.error("Error loading auditors or levels:", e);
      showToast("error", "Failed to load auditors or levels");
    } finally {
      hideLoader();
    }
  };

  // Load main table data
  const loadListData = async () => {
    try {
      showLoader();
      setRedirect(false);
      const items = await sp.web.lists
        .getByTitle(listName)
        .items.select('*')
        .top(2000)
        .orderBy("Modified", false)();
      const tableData = items.map((item: any) => ({
        Id: item.Id,
        Title: item.Title,
        Level: item.Level,
        Is_x0020_Active: item.Is_x0020_Active ? "Yes" : "No"
      }));
      setActionsData(tableData);
    } catch (e) {
      console.log(e);
      onError();
    } finally {
      hideLoader();
    }
  };

  // Edit item
  const editItem = async (Id: number) => {
    try {
      showLoader();
      setIsFormOpen(true);
      setItemId(Id);
      const item = await sp.web.lists.getByTitle(listName).items.getById(Id)();
      setFormData({
        Title: item.Title || "",
        Level: item.Level || "",
        Is_x0020_Active: item.Is_x0020_Active || false
      });
      setTimeout(() => {
        const input = auditCategoryRef.current?.querySelector("input");
        input?.focus();
      }, 0);
    } catch (e) {
      console.log(e);
      onError();
    } finally {
      hideLoader();
    }
  };

  // Add new item
  const addNew = () => {
    setIsFormOpen(true);
    setItemId(0);
    setFormData({ Title: "", Level: "", Is_x0020_Active: false });
    setTimeout(() => {
      const input = auditCategoryRef.current?.querySelector("input");
      input?.focus();
    }, 0);
  };

  // Check duplicate
  const checkDuplicate = async () => {
    try {
      showLoader();
      const filterQuery = `Title eq '${formData.Title}' and Level eq '${formData.Level}'`;
      const results = await sp.web.lists.getByTitle(listName).items.filter(itemId > 0 ? `${filterQuery} and Id ne ${itemId}` : filterQuery)();
      hideLoader();
      if (results.length > 0) {
        showToast("error", "Record already exists");
        return false;
      }
      return true;
    } catch (e) {
      console.log(e);
      hideLoader();
      onError();
      return false;
    }
  };

  // Insert or update
  const insertOrUpdate = async () => {
    try {
      const data: any = {
        Title: formData.Title.trim(),
        Level: formData.Level,
        Is_x0020_Active: formData.Is_x0020_Active
      };
      if (itemId > 0) {
        await sp.web.lists.getByTitle(listName).items.getById(itemId).update(data);
        showToast("success", "Auditor Level updated successfully");
      } else {
        await sp.web.lists.getByTitle(listName).items.add(data);
        showToast("success", "Auditor Level submitted successfully");
      }
      setRedirect(true);
      closeForm();
      loadListData();
    } catch (e) {
      console.log(e);
      onError();
    }
    finally {
      hideLoader();
    }
  };
  const handleChangeDropdown = (selectedOption: any, actionMeta: any) => {
    if (!actionMeta?.name) return;

    const name = actionMeta.name;
    const value = actionMeta.action === 'clear' ? '' : selectedOption?.value || '';

    // update state
    setFormData(prev => ({
      ...prev,
      [name]: value // store the text (for Auditor or Level)
    }));

    // optional: reset Level when Auditor changes
    if (name === 'Title') {
      setFormData(prev => ({
        ...prev,
        Level: ''
      }));
    }
  };
  // Form submit
  const handleSubmit = async (event: any) => {
    event.preventDefault();
    showLoader();
    const data = {
      Auditors: { val: formData.Title.trim(), required: true, Name: "Auditor", Type: ControlType.reactSelect, Focusid: 'txtAuditor' },
      AuditorLevel: { val: formData.Level.trim(), required: true, Name: "Auditor Level", Type: ControlType.reactSelect, Focusid: 'txtLevel' },

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

  const closeForm = () => {
    setIsFormOpen(false);
    setFormData({ Title: "", Level: "", Is_x0020_Active: false });
  };

  // const onSuccess = () => {
  //   closeForm();
  //   showToast("success", "Auditor submitted successfully");
  //   hideLoader();
  // };

  const onError = () => {
    showToast("error", ActionStatus.Error);
    hideLoader();
  };

  // const handleRowClicked = (row: any) => editItem(row.Id);

  // const columns = [
  //   {
  //     name: "Edit",
  //     selector: (row: any) => row.Id,
  //     cell: (record: any) => (
  //       <NavLink title="Edit" onClick={() => editItem(record.Id)} to="">
  //         <FontAwesomeIcon icon={faEdit} />
  //       </NavLink>
  //     ),
  //     sortable: false,
  //     width: "60px"
  //   },
  //   {
  //     name: "Audit Level",
  //     selector: (row: any) => row.Title,
  //     sortable: true
  //   },
  //   {
  //     name: "Level",
  //     selector: (row: any) => row.Level,
  //     sortable: true
  //   },
  //   {
  //     name: "Is Active",
  //     selector: (row: any) => row.Is_x0020_Active,
  //     sortable: true
  //   }
  // ];

  const columns =React.useMemo(()=> [
    {
      headerName: "Edit",
      field: "Id",
      width: 80,
      minWidth: 80,
      maxWidth: 80,
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
      headerName: "Audit Level",
      field: "Title",
      sortable: true,
      filter: "agTextColumnFilter",
      resizable: true,
      flex: 1,
      getQuickFilterText: (params: any) => params.value || "",
    },
    {
      headerName: "Level",
      field: "Level",
      sortable: true,
      filter: "agTextColumnFilter",
      resizable: true,
      flex: 1,
    },
    {
      headerName: "Is Active",
      field: "Is_x0020_Active",
      sortable: true,
      filter: "agTextColumnFilter",
      resizable: true,
      flex: 1,
    },
  ],[]);

  return (
    <div className="container-fluid">
      <div className="light-box border-box-shadow">
        <div className="div-form-title">
          <div className="form-title">LPA Auditors</div>
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
              <div className="row">
                <div className="col-md-3">
                  <div className="light-text">
                    <div ref={auditCategoryRef}>
                      <SearchableDropdown
                        label="Auditor"
                        Title="Auditor"
                        name="Title"
                        id="txtAuditor"
                        className="txtAuditor"
                        selectedValue={formData.Title} // current value
                        OptionsList={Auditors.map(a => ({
                          label: a.Title,                  // use Title as key
                          value: a.Title                  // display Title
                        }))}
                        OnChange={async (selectedOption: any, actionMeta: any) => handleChangeDropdown(selectedOption, actionMeta)}

                        isRequired={true}
                        disabled={false}
                      />
                    </div>
                  </div>
                </div>

                <div className="col-md-3">
                  <div className="light-text">
                    <SearchableDropdown
                      label="Level"
                      Title="Level"
                      name="Level"
                      id="txtLevel"
                      className="txtLevel"
                      selectedValue={formData.Level} // current value
                      OptionsList={AuditLevels.map(l => ({ label: l.Title, value: l.Title }))}
                      OnChange={async (selectedOption: any, actionMeta: any) => handleChangeDropdown(selectedOption, actionMeta)}
                      isRequired={true}
                      disabled={false}
                    />
                  </div>
                </div>
                <div className="col-md-3">
                  <InputCheckBox label={"Is Active"} name={"Is_x0020_Active"} checked={formData.Is_x0020_Active} onChange={(e: any) => setFormData(prev => ({ ...prev, Is_x0020_Active: e.target.checked }))} isdisable={false} isRequired={false}></InputCheckBox>

                </div>

                <div className="col-md-3 py-2 text-center">
                  <button className="btn btn-primary mx-2" onClick={handleSubmit}>
                    {itemId ? "Update" : "Submit"}
                  </button>
                  <button className="btn btn-secondary" onClick={closeForm}>
                    Cancel
                  </button>
                </div>
              </div>
            </div>
          )}

          {/* <TableGenerator
            columns={columns}
            data={actionsData}
            onChange={setPageNumber}
            prvPageNumber={pageNumber}
            prvDirection={false}
            fileName="KPI"
            className="sp-Datatable-hh"
            onRowClick={handleRowClicked}
            showPagination={true}
          /> */}

          <div className="mx-2 table-head-1st-td right-search-table mb-3">
            {/* <TableGenerator columns={columns} data={this.state.data} fileName={'Location2'} onRowClick={(row:any)=>this.onEditClickHandler(row.Id)} ></TableGenerator> */}
            <AGGridDataTable
              data={actionsData}
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
        </div>
      </div>
    </div>
  );
};

export default LPAAuditors;