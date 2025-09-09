// Functional component
import * as React from "react";
import { useState, useEffect, useRef } from "react";
import { SPHttpClient } from "@microsoft/sp-http";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import TableGenerator from "../Shared/TableGenerator";
import { hideLoader, showLoader } from "../Shared/Loader";
import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";
import { faEdit, faPlus} from "@fortawesome/free-solid-svg-icons";
import { highlightCurrentNav } from "../Utilities/HighlightCurrentComponent";
import { ActionStatus, ControlType } from "../Constants/Contants";
import formValidation from "../Utilities/FormValidator";
import { showToast } from "../Shared/Toaster";
import { Navigate } from "react-router-dom";

export interface ActionsProps {
  match: any;
  spContext: any;
  spHttpClient: SPHttpClient;
  context: any;
  history: any;
  currentUser: any;
}

const Actions: React.FC<ActionsProps> = (props) => {
  const ActionsList = "Actions";
  const sp = spfi().using(SPFx(props.context));
  const txtAction = useRef<HTMLInputElement>(null);

  // State
  const [actionsData, setActionsData] = useState<any[]>([]);
  const [pageNumber, setPageNumber] = useState<number>(1);
  const [sortBy, setSortBy] = useState<number>(1);
  const [sortOrder, setSortOrder] = useState<boolean>(false);
//   const [searchText, setSearchText] = useState<string>("");
  const [isFormOpen, setIsFormOpen] = useState<boolean>(false);
  const [itemId, setItemId] = useState<number>(0);
  const [formData, setFormData] = useState<any>({
    Title: "",
    RootCause: 0,
    SecondaryRootCause: 0,
  });
  const [redirect, setRedirect] = useState<boolean>(false);
  const [isUnauthorized, setIsUnauthorized] = useState<boolean>(false);
  const [rootCauses, setRootCauses] = useState<any[]>([]);
  const [secondaryRootCauses, setSecondaryRootCauses] = useState<any[]>([]);

  // ComponentDidMount equivalent
  useEffect(() => {
    highlightCurrentNav("liActions");
    document.title = "Mayco - Safety | Actions";
    loadListData();
  }, []);

  // ComponentDidUpdate equivalent
  useEffect(() => {
    if (redirect) {
      loadListData();
    }
  }, [redirect]);

  const loadListData = async () => {
    try {
      showLoader();
      setRedirect(false);
      const lsTableProps = { PageNumber: 1, sortOrder: false, sortBy: 1, searchText: null };
      localStorage.setItem("PrvData", JSON.stringify(lsTableProps));

      const [Actions, RootCauses, SecondaryRootCauses] = await Promise.all([
        sp.web.lists.getByTitle(ActionsList)
          .items.top(2000)
          .select('Title,RootCause/Title,RootCause/Id,SecondaryRootCause/Title,SecondaryRootCause/Id,*')
          .expand('RootCause,SecondaryRootCause')
          .orderBy("Modified", false)(),
        sp.web.lists.getByTitle("RootCauses").items.top(2000).orderBy("Title", true)(),
        sp.web.lists.getByTitle("SecondaryRootCauses").items.top(2000).orderBy("Title", true)(),
      ]);

      const tableData = Actions.map((Act: any) => ({
        Id: Act.Id,
        Title: Act.Title,
        RootCauseId: Act.RootCause.Id,
        SecondaryRootCauseId: Act.SecondaryRootCause.Id,
        RootCauseTitle: Act.RootCause.Title,
        SecondaryRootCauseTitle: Act.SecondaryRootCause.Title,
      }));

      setActionsData(tableData);
      setRootCauses(RootCauses);
      setSecondaryRootCauses(SecondaryRootCauses);
      setIsUnauthorized(false);
      console.log(rootCauses);
      console.log(secondaryRootCauses);
    } catch (e) {
      onError();
      console.log(e);
    } finally {
      hideLoader();
    }
  };

  const editItem = async (Id: number) => {
    try {
      const newFormData = { ...formData };
      newFormData.Title = "";
      showLoader();
      setIsFormOpen(true);
      setItemId(Id);
      await sp.web.lists.getByTitle(ActionsList).items.getById(Id)().then((item: any) => {
        if (item.Error) {
          hideLoader();
          console.log(item.Error);
        } else {
          newFormData.Title = item.Title;
          hideLoader();
          setFormData(newFormData);
        }
      });
    } catch (e) {
      onError();
      hideLoader();
      console.log(e);
    }
  };

  const addNew = () => {
    setIsFormOpen(true);
    setItemId(0);
  };

  const checkDuplicate = async () => {
    try {
      showLoader();
      let isValid = true;
      const escapedTitle = formData.Title.replace(/'/g, "''");
      let filterQuery = `Title eq '${escapedTitle}'`;

      if (itemId > 0) {
        filterQuery += ` and Id ne ${itemId}`;
      }

      await sp.web.lists.getByTitle(ActionsList).items.filter(filterQuery)().then((res: any) => {
        if (!res.Error && res.length > 0) {
          isValid = false;
          const message = "Action already exists";
          showToast("error", message);
          hideLoader();
        } else {
          hideLoader();
        }
      });

      return isValid;
    } catch (e) {
      onError();
      hideLoader();
      console.log(e);
    }
  };

  const handleSubmit = async (event: any) => {
    showLoader();
    try {
      event.preventDefault();
      const data = {
        Action: { val: formData.Title.trim(), required: true, Name: "'Action'", Type: ControlType.string, Focusid: txtAction }
      };

      const isValid = formValidation.FormValidation(data);

      if (isValid.status) {
        const validDuplicate = await checkDuplicate();
        if (validDuplicate) {
          insertOrUpdateData();
        }
      } else {
        showToast("error", isValid.message);
        hideLoader();
      }
    } catch (e) {
      onError();
      hideLoader();
      console.log(e);
    }
  };

  const insertOrUpdateData = () => {
    try {
      const formDataCopy = { ...formData };

      if (itemId > 0) {
        sp.web.lists.getByTitle(ActionsList).items.getById(itemId).update(formDataCopy).then(() => {
          const msg = "Action updated successfully";
          setRedirect(true);
          showToast("success", msg);
        }, (error) => {
          console.log(error);
          onError();
        });
      } else {
        sp.web.lists.getByTitle(ActionsList).items.add(formDataCopy).then(() => {
          const msg = "Action submitted successfully";
          setRedirect(true);
          showToast("success", msg);
        }, (error) => {
          console.log(error);
          onError();
        });
      }
    } catch (e) {
      onError();
      hideLoader();
      console.log(e);
    }
  };

  const onError = () => {
    showToast("error", ActionStatus.Error);
    hideLoader();
  };

  const closeForm = () => {
    const newFormData = { ...formData };
    newFormData.Title = "";
    setIsFormOpen(false);
    setFormData(newFormData);
  };

  const onPageChange = (pageIndex: any) => {
    setPageNumber(pageIndex);
  };

  const onSortOrder = (event: any, sortDirection: any) => {
    setSortBy(event.id);
    setSortOrder(sortDirection);
  };

  const handleChangeDynamic = (event: any) => {
    const newFormData = { ...formData };
    const name = event.target.name;
    const value = event.target.type === "checkbox" ? event.target.checked : event.target.value;
    newFormData[name] = value;
    setFormData(newFormData);
  };

  const handleRowClicked = (row: any, Id?: any) => {
    const ID = row.Id ? row.Id : Id;
    editItem(ID);
  };

  const columns = [
    {
      name: "Edit",
      selector: (row: { Id: any }) => row.Id,
      export: false,
      width: "100px",
      cell: (record: { Id: any }) => (
        <button type="button" id="btnEdit" className="btn" title="Edit" onClick={() => editItem(record.Id)}>
          <FontAwesomeIcon icon={faEdit} />
        </button>
      ),
      sortable: false,
    },
    {
      name: "Action",
      selector: (row: { Title: any }) => row.Title,
      sortable: true,
      cell: (record: { Title: any }) => record.Title,
    },
    {
      name: "Root Cause",
      selector: (row: { RootCauseTitle: any }) => row.RootCauseTitle,
      sortable: true,
      cell: (record: { RootCauseTitle: any }) => record.RootCauseTitle,
    },
    {
      name: "Secondary Root Cause",
      selector: (row: { SecondaryRootCauseTitle: any }) => row.SecondaryRootCauseTitle,
      sortable: true,
      cell: (record: { SecondaryRootCauseTitle: any }) => record.SecondaryRootCauseTitle,
    },
  ];

  if (isUnauthorized) {
    return <Navigate to="/UnAuthorized" />;
  }

  return (
    <div id="content" className="content p-2 pt-2">
      <div className="container-fluid">
        <div className="FormContent border-none">
          <div className="title">Actions</div>
          {!isFormOpen && (
            <div className="text-end">
              <button
                type="button"
                id="btnNew"
                className="SubmitButtons btn btn-new fw-bold"
                title="New"
                onClick={addNew}
              >
                <FontAwesomeIcon icon={faPlus} /> New
              </button>
            </div>
          )}
          {isFormOpen && (
            <div className="border-top mt-3 py-3">
              <div className="row">
                <div className="col-md-3">
                  <div className="form-floating">
                    <input
                      className="form-control"
                      required
                      placeholder="Action"
                      type="text"
                      name="Title"
                      title="Action"
                      value={formData.Title}
                      onChange={handleChangeDynamic}
                      id="txtAction"
                      autoComplete="off"
                      ref={txtAction}
                      maxLength={250}
                    />
                    <label>
                      Action <span className="mandatoryhastrick">*</span>
                    </label>
                  </div>
                </div>
                <div className="col-md-3 btnDiv">
                  <button type="button" id="btnSubmit" className="SubmitButtons btn" title="Submit" onClick={handleSubmit}>
                    Submit
                  </button>
                  <button
                    type="button"
                    id="btnCancel"
                    className="CancelButtons btn btn-secondary"
                    title="Cancel"
                    onClick={closeForm}
                  >
                    Cancel
                  </button>
                </div>
              </div>
              <span id="spanErrorMessage" style={{ display: "none", color: "red" }}></span>
            </div>
          )}
          <TableGenerator
            columns={columns}
            data={actionsData}
            onChange={onPageChange}
            onSortChange={onSortOrder}
            prvPageNumber={pageNumber}
            prvDirection={sortOrder}
            prvSort={sortBy}
            fileName={"Actions"}
            onRowClick={handleRowClicked}
            showPagination={true}
          />
        </div>
      </div>
    </div>
  );
};

export default Actions;

