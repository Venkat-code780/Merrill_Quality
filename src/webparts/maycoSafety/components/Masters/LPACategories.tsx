import * as React from "react";
import { useState, useEffect, useRef } from "react";
import { SPHttpClient } from "@microsoft/sp-http";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
// import TableGenerator from "../Shared/TableGenerator";
import { hideLoader, showLoader } from "../Shared/Loader";
// import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";
// import { faPlus } from "@fortawesome/free-solid-svg-icons";
import { highlightCurrentNav } from "../Utilities/HighlightCurrentComponent";
import { ActionStatus, ControlType } from "../Constants/Contants";
import formValidation from "../Utilities/FormValidator";
import { showToast } from "../Shared/Toaster";
// import { NavLink } from "react-router-dom";
import SearchableDropdown from "../Shared/Dropdown";
import { initCommonFunctions } from "../Utilities/CommonFunctions";
import { useNavigate } from "react-router-dom";

export interface LPACategoriesProps {
    match: any;
    spHttpClient: SPHttpClient;
    context: any;
    siteURL: string;
    spContext: any;
    webAbsoluteURL: string;
    currPlantTitle: string;

}

const LPACategories: React.FC<LPACategoriesProps> = (props) => {
    const rootSiteURL = props.spContext.siteAbsoluteUrl;
    const currentSiteURL = props.spContext.webAbsoluteUrl;
    const JvisURL = `${rootSiteURL}/mayco`;
    const { getListItems } = initCommonFunctions(props.context, props.siteURL);
    const navigate = useNavigate();

    const listName = "LPA Categories";
      const departmentRef = useRef<any>(null);
    

    const sp = spfi().using(SPFx(props.context));
    // const txtPPEType = useRef<HTMLInputElement>(null);

    //   const [actionsData, setActionsData] = useState<any[]>([]);
    //   const [pageNumber, setPageNumber] = useState(1);
    const [selectedPlant, setSelectedPlant] = useState("");
    const [categoryRows, setCategoryRows] = useState<any[]>([]);
    const [itemId, setItemId] = useState(0);

    // const [displayMessage, setDisplayMessage] = useState("");
    const [Plants, setPlants] = useState<any[]>([]);
    const [Departments, setDepartment] = useState<any[]>([]);
    const categoriesString =
        categoryRows.map(row => row.Category).join("#;&@") + "#;&@";
    const [formData, setFormData] = useState({
        Plant: "",
        Department: "",
        Categories: categoriesString
    });

    // Load initial data
    useEffect(() => {
        highlightCurrentNav("liPPETypes");
        document.title = "JVIS - Quality | LPA Auditor Levels";
        // loadListData();
        loadAuditorsLevels();
                setTimeout(() => {
      const input = departmentRef.current?.querySelector("input");
      input?.focus();
    }, 0);
        
    }, []);



    // Load Auditors filtered by Plant 'Shelby' and Audit Levels
    const loadAuditorsLevels = async () => {
        try {
            showLoader();

            // Create SP instance for the root site

            const [Plant, Department] = await Promise.all([
                getListItems("Plant", JvisURL, "*", "", "Title eq 'Merrill'"),
                getListItems("Department", JvisURL, "Plant/Title,Title", "Plant", "Plant/Title eq 'Merrill'")
            ]);
            setPlants(Plant);
            setDepartment(Department);

             if (Plant.length > 0) {
            setSelectedPlant(Plant[0].Title);

            setFormData(prev => ({
                ...prev,
                Plant: Plant[0].Title
            }));
        }
        } catch (e) {
            console.error("Error loading auditors or levels:", e);
            showToast("error", "Failed to load auditors or levels");
        } finally {
            hideLoader();
        }
    };
    const getCategoriesData = async (plant: string, department: string) => {
        try {
            showLoader();

            // 1️⃣ Audit Categories
            const auditItems = await getListItems(
                "Audit_Categories",
                currentSiteURL,
                "Title",
                "",
                ""
            );

            const allCategories = auditItems.map((i: any) => i.Title);

            // 2️⃣ LPA Categories
            const lpaItems = await getListItems(
                "LPA Categories",
                currentSiteURL,
                "Plant,Department,Categories,ID",
                "",
                `Plant eq '${plant}' and Department eq '${department}'`
            );

            let finalCategories: string[] = [];

            if (lpaItems.length > 0 && lpaItems[0].Categories) {
                setItemId(lpaItems[0].Id);

                const saved = lpaItems[0].Categories
                    ? lpaItems[0].Categories.split("#;&@").filter((c: string) => c)
                    : [];

                // ✅ DO NOT redeclare
                finalCategories = [...saved];

                allCategories.forEach((cat: string) => {
                    if (!finalCategories.includes(cat)) {
                        finalCategories.push(cat);
                    }
                });

            } else {
                finalCategories = allCategories;
                setItemId(0);
            }

            const rows = finalCategories.map((cat, index) => ({
                Category: cat,
                Order: index + 1,
                Options: Array.from({ length: finalCategories.length }, (_, i) => i + 1)
            }));

            setCategoryRows(rows);
            const categoriesString =
                rows.map(row => row.Category).join("#;&@") + "#;&@";

            setFormData(prev => ({
                ...prev,
                Categories: categoriesString
            }));

        } catch (e) {
            console.error(e);
        } finally {
            hideLoader();
        }
    };
  
    // Check duplicate
    const checkDuplicate = async () => {
        try {
            showLoader();
            const filterQuery = `Title eq '${formData.Plant}'`;
            const results = await sp.web.lists
                .getByTitle(listName)
                .items.filter(itemId > 0 ? `${filterQuery} and Id ne ${itemId}` : filterQuery)();
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
                Plant: formData.Plant.trim(),
                Department: formData.Department,
                Categories: formData.Categories
            };
            if (itemId > 0) {
                await sp.web.lists.getByTitle(listName).items.getById(itemId).update(data);
                // showToast("success", "Auditor Level updated successfully");
            } else {
                await sp.web.lists.getByTitle(listName).items.add(data);
                showToast("success", "Auditor Level submitted successfully");
            }
    
            onSuccess();
        } catch (e) {
            console.log(e);
            onError();
        }
    };
    //   const handleChangeDropdown = (selectedOption: any, actionMeta: any) => {
    //     if (!actionMeta?.name) return;

    //     const name = actionMeta.name;
    //     const value = actionMeta.action === 'clear' ? '' : selectedOption?.value || '';

    //     // update state
    //     setFormData(prev => ({
    //       ...prev,
    //       [name]: value // store the text (for Auditor or Level)
    //     }));

    //     // optional: reset Level when Auditor changes
    //     if (name === 'Plant') {
    //       setFormData(prev => ({
    //         ...prev,
    //         Level: ''
    //       }));
    //     }
    //   };

 const handleChangeDropdown = async (selectedOption: any, actionMeta: any) => {
    if (!actionMeta?.name) return;

    const name = actionMeta.name;
    const value = actionMeta.action === 'clear' ? '' : selectedOption?.value || '';

    // ✅ Update form state
    setFormData(prev => ({
        ...prev,
        [name]: value
    }));
    if (name === "Department") {
        if (value) {
            // Department selected → load categories
            await getCategoriesData(selectedPlant, value);
        } else {
            // Department cleared → hide table and reset categories
            setCategoryRows([]);
            setFormData(prev => ({
                ...prev,
                Categories: ""
            }));
        }
    }
    // 🔥 LOAD TABLE WHEN DEPARTMENT SELECTED
    // if (name === "Department" && value) {
    //     await getCategoriesData(selectedPlant, value);
    // }
};
    // const handleOrderChange = (changedIndex: number, newValue: number) => {
    //     const updated = [...categoryRows];
    //     const oldValue = updated[changedIndex].Order;

    //     const existingIndex = updated.findIndex(
    //         (row, i) => row.Order === newValue && i !== changedIndex
    //     );

    //     if (existingIndex !== -1) {
    //         const temp = updated[changedIndex].Category;
    //         updated[changedIndex].Category = updated[existingIndex].Category;
    //         updated[existingIndex].Category = temp;
    //     }

    //     updated[changedIndex].Order = oldValue;

    //     setCategoryRows(updated);
    // };
    const handleOrderChange = (changedIndex: number, newValue: number) => {
    const updated = [...categoryRows];
    const oldValue = updated[changedIndex].Order;

    const existingIndex = updated.findIndex(
        (row, i) => row.Order === newValue && i !== changedIndex
    );

    if (existingIndex !== -1) {
        const temp = updated[changedIndex].Category;
        updated[changedIndex].Category = updated[existingIndex].Category;
        updated[existingIndex].Category = temp;
    }

    updated[changedIndex].Order = oldValue;

    // ✅ SORT by Order before saving
    const sorted = [...updated].sort((a, b) => a.Order - b.Order);

    setCategoryRows(sorted);

    // ✅ UPDATE Categories string (🔥 MAIN FIX)
    const categoriesString =
        sorted.map(row => row.Category).join("#;&@") + "#;&@";

    setFormData(prev => ({
        ...prev,
        Categories: categoriesString
    }));
};
    // Form submit
    const handleSubmit = async (event: any) => {
        event.preventDefault();
        showLoader();
        const data = {
            Department: { val: formData.Department.trim(), required: true, Name: "Department", Type: ControlType.reactSelect, Focusid: 'txtDepartment' },

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
   
        setFormData({ Plant: "", Department: "", Categories: "" });
        navigate('/Home');
    };

    const onSuccess = () => {
        showToast("success", "Categories submitted successfully");
         setFormData({
        Plant: "Shelby",   // Keep Plant as Shelby
        Department: "",    // Clear Department
        Categories: ""     // Clear Categories
    });
        setCategoryRows([]);
        hideLoader();
    };

    const onError = () => {
        showToast("error", ActionStatus.Error);
        hideLoader();
    };

    return (
        <div className="container-fluid">
            <div className="light-box border-box-shadow">
                <div className="div-form-title">
                    <div className="form-title">LPA Categories</div>
                 
                        <span className="span-mandatory-text">
                            <span className="text-danger">*</span> are mandatory fields
                        </span>
            
                </div>

                <div className="p-2 mx-1 ViewTable">
                  
             
                        <div className="form-border-box p-2 mx-1 my-2">
                            <div className="row">
                                <div className="col-md-3">
                                    <div className="light-text">
                                    <label>
                                        Plant <span className="mandatoryhastrick">*</span>
                                    </label>
                                    <select
                                        className="form-control"
                                        value={formData.Plant || ""}
                                        disabled={true}
                                        onChange={(e) =>
                                            setFormData((prev) => ({
                                                ...prev,
                                                Plant: e.target.value,
                                                Department: "" // reset department when plant changes
                                            }))
                                        }
                                    >

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
                                        <div ref={departmentRef}>
                                    <SearchableDropdown
                                        label="Department"
                                        Title="Department"
                                        name="Department"
                                        id="txtDepartment"
                                        className="txtLevel"
                                        selectedValue={formData.Department} // current value
                                        OptionsList={Departments.map(l => ({ label: l.Title, value: l.Title }))}
                                        OnChange={async (selectedOption: any, actionMeta: any) => handleChangeDropdown(selectedOption, actionMeta)}
                                        isRequired={true}
                                        disabled={false}
                                    />
                                    </div>
                                    </div>
                                </div>

                             
                            </div>
                        </div>
        

                    {categoryRows.length > 0 && (
                        <div className="mt-3 select-appearance">
                            <table className="table-bordered w-100">
                                <thead className="table-dark">
                                    <tr>
                                        <th>Category</th>
                                        <th>Value</th>
                                    </tr>
                                </thead>

                                <tbody>
                                    {categoryRows.map((row, index) => (
                                        <tr key={index}>
                                            <td>{row.Category}</td>

                                            <td>
                                                <select
                                                    className="form-control"
                                                    value={row.Order}
                                                    onChange={(e) =>
                                                        handleOrderChange(index, Number(e.target.value))
                                                    }
                                                >
                                                    {row.Options.map((num: number) => (
                                                        <option key={num} value={num}>
                                                            {num}
                                                        </option>
                                                    ))}
                                                </select>
                                            </td>
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        </div>
                    )}
                       <div className="col-md-12 py-2 text-center">
                                    <button className="btn btn-primary mx-2" onClick={handleSubmit}>
                                        Submit
                                    </button>
                                    <button className="btn btn-secondary" onClick={closeForm}>
                                        Cancel
                                    </button>
                                </div>
                </div>
            </div>
        </div>
    );
};

export default LPACategories;