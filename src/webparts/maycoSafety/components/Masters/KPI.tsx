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
import AGGridDataTable from '../Shared/AGGridDataTable';

export interface KPIProps {
  spHttpClient: SPHttpClient;
  context: any;
}

const KPI: React.FC<KPIProps> = (props) => {

  const listName = "KPI";

  const sp = spfi().using(SPFx(props.context));

  const txtPPEType = useRef<HTMLInputElement>(null);

  const [actionsData, setActionsData] = useState<any[]>([]);
  // const [pageNumber, setPageNumber] = useState(1);
  const [isFormOpen, setIsFormOpen] = useState(false);
  const [itemId, setItemId] = useState(0);
  const [redirect, setRedirect] = useState(false);
  // const [displayMessage, setDisplayMessage] = useState("");

  const [formData, setFormData] = useState({
    Title: ""
  });

  useEffect(() => {
    highlightCurrentNav("liPPETypes");
    document.title = "Mayco - Safety | PPE Types";
    loadListData();
  }, []);

  useEffect(() => {
    if (redirect) loadListData();
  }, [redirect]);

  const loadListData = async () => {

    try {

      showLoader();

      setRedirect(false);

      const items = await sp.web.lists
        .getByTitle(listName)
        .items.select("Id", "Title")
        .top(2000)
        .orderBy("Modified", false)();

      const tableData = items.map((item: any) => ({
        Id: item.Id,
        Title: item.Title
      }));

      setActionsData(tableData);

    } catch (e) {

      onError();
      console.log(e);

    } finally {

      hideLoader();

    }

  };

  const editItem = async (Id: number) => {

    try {

      showLoader();

      setIsFormOpen(true);
      setItemId(Id);

      const item = await sp.web.lists.getByTitle(listName).items.getById(Id)();

      setFormData({
        Title: item.Title
      });

      txtPPEType.current?.focus();

      hideLoader();

    } catch {

      onError();

    }

  };

  const addNew = () => {

    setIsFormOpen(true);
    setItemId(0);
    setFormData({ Title: "" });
    setTimeout(() => {
      txtPPEType.current?.focus();
    }, 100);

  };

  const checkDuplicate = async () => {

    try {

      showLoader();

      const title = formData.Title.trim();

      let filterQuery = `Title eq '${title}'`;

      if (itemId > 0) filterQuery += ` and Id ne ${itemId}`;

      const results = await sp.web.lists
        .getByTitle(listName)
        .items.filter(filterQuery)();

      hideLoader();

      if (results.length > 0) {

        showToast("error", "Record already exists");
        return false;

      }

      return true;

    } catch {

      onError();
      hideLoader();
      return false;

    }

  };

  const insertOrUpdate = async () => {

    try {

      const data = {
        Title: formData.Title.trim()
      };

      if (itemId > 0) {

        await sp.web.lists.getByTitle(listName).items.getById(itemId).update(data);

        showToast("success", "KPI updated successfully");
      } else {

        await sp.web.lists.getByTitle(listName).items.add(data);

        showToast("success", "KPI submitted successfully");

      }

      setRedirect(true);
      closeForm();
      loadListData();

    } catch {

      onError();

    }
    finally {
      hideLoader();
    }

  };

  const handleSubmit = async (event: any) => {

    event.preventDefault();

    showLoader();

    const data = {
      Status: {
        val: formData.Title.trim(),
        required: true,
        Name: "KPI",
        Type: ControlType.string,
        Focusid: txtPPEType
      }
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

  const handleChange = (event: any) => {

    const { name, value } = event.target;

    setFormData({
      ...formData,
      [name]: value
    });

  };

  const closeForm = () => {

    setIsFormOpen(false);

    setFormData({
      Title: ""
    });

  };

  // const onSuccess = () => {

  //   closeForm();

  //   showToast("success", displayMessage);

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
  //     name: "AuditCategory",
  //     selector: (row: any) => row.Title,
  //     sortable: true
  //   }

  // ];

  const columns =React.useMemo(()=> [
    {
      field: "Id",
      headerName: "Edit",
      sortable: false,
      filter: false,
      width: 80,
      minWidth: 80,
      maxWidth: 80,
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
      field: "Title",
      headerName: "KPI",
      sortable: true,
      filter: "agTextColumnFilter",
      resizable: true,
      flex: 1,
      getQuickFilterText: (params: any) => params.value || "",
    },
  ],[]);
  return (

    <div className="container-fluid">

      <div className="light-box border-box-shadow">

        <div className="div-form-title">

          <div className="form-title">KPI</div>

          {isFormOpen &&
            <span className="span-mandatory-text">
              <span className="text-danger">*</span> are mandatory fields
            </span>
          }

        </div>

        <div className="p-2 mx-1 ViewTable">

          {!isFormOpen &&
            <div className="text-end my-2 mx-3">

              <button className="NewButton" onClick={addNew}>
                <FontAwesomeIcon icon={faPlus} /> New
              </button>

            </div>
          }

          {isFormOpen &&
            <div className="form-border-box p-2 mx-1 my-2">

              <div className="row">

                <div className="col-md-3">

                  <div className="light-text">

                    <input
                      className="form-control"
                      type="text"
                      name="Title"
                      value={formData.Title}
                      onChange={handleChange}
                      ref={txtPPEType}
                      maxLength={250}
                    />

                    <label>
                      KPI <span className="mandatoryhastrick">*</span>
                    </label>

                  </div>

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
          }

          {/* <TableGenerator
            columns={columns}
            data={actionsData}
            onChange={setPageNumber}
            prvPageNumber={pageNumber}
            prvDirection={false}
            fileName={"KPI"}
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

export default KPI;