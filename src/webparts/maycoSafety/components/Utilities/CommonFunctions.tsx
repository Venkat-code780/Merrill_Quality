// import * as React from "react";
// import { SPHttpClient } from "@microsoft/sp-http";
// import { SPFI, spfi, SPFx } from "@pnp/sp";
// import { hideLoader } from "../Shared/Loader";
// import { showToast } from "../Shared/Toaster";
// import { ActionStatus } from "../Constants/Contants";


// export interface CommonFunctionsProps {
//     match: any;
//     spContext: any;
//     spHttpClient: SPHttpClient;
//     context: any;
//     history: any;
//     isSupplierTeam: boolean;
//     isDTETeam: boolean;
//     isProcurementTeam: boolean;
// }
// export interface JSRAFormState {
// }
// export default class CommonFunctions extends React.Component<CommonFunctionsProps> {
//     private siteURL: string; 
//     private MaycoSP:SPFI;
//     constructor(props: CommonFunctionsProps) {
//              super(props);
//             this.siteURL = this.props.spContext.siteAbsoluteUrl;
//             this.MaycoSP = spfi(`${this.siteURL}/mayco`).using(SPFx(this.props.context));
//             console.log(this.siteURL);      //sites/wcm
//         }
//         public state = {
//         //Cascaded dropdowns
//         Plants:[],Departments:[], Zones:[], WorkCells:[], Machines:[],Shifts:[], Supervisors:[],
//     }
// public componentDidMount() {
//                 this.getOnLoad();
//             }
        
// public  getOnLoad= async()=>
//     {
//         try
//         {
//       let  [Plants]=await Promise.all([
//             this.MaycoSP.web.lists.getByTitle('Plant').items.top(2000).select('Title,*').expand('').orderBy("Title", true)()
//            ])
//         }
//         catch(error){
//             console.log('Eroor while getting common data' ,error);
//           this.onError();
//         }
//     }
// private onError = () => {
//             showToast("error", ActionStatus.Error);
//             hideLoader();
// }

//  }


import { SPFI, spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/batching";
// import { SPHttpClient } from "@microsoft/sp-http";
import { hideLoader } from "../Shared/Loader";
import { showToast } from "../Shared/Toaster";
import { ActionStatus } from "../Constants/Contants";

// Initialize SPFI instance for mayco site
export const initCommonFunctions = (context: any,siteAbsoluteURL:string)=> {
    // get List Items
    const getListItems= async (ListName:string,URL:string,selectQuery:string='',expand:string='',filterQuery:string='') : Promise<any[]> => 
        {
        const SiteURL: SPFI = spfi(URL).using(SPFx(context)); 
            try{
            const items = await SiteURL.web.lists
                .getByTitle(ListName)
                .items.top(5000)
                .select(selectQuery)
                .filter(filterQuery)
                .expand(expand)()
                 return items;
        } catch (error) {
            console.error("Error fetching List:", error);
            showToast("error", ActionStatus.Error);
            hideLoader();
            return [];
        }
    }

    //  Add List Item
    const addListItem = async (ListName: string, URL: string, postObject: any): Promise<any> => {
        
        const SiteURL: SPFI = spfi(URL).using(SPFx(context)); 
        try {
            const item = await SiteURL.web.lists
                .getByTitle(ListName)
                .items.add(postObject);
            return item;
        } catch (error) {
            console.error("Error adding item:", error);
            showToast("error", ActionStatus.Error);
            hideLoader();
            return null;
        }
    };

    // Update List Item
    const updateListItem = async (ListName: string, URL: string, postObject: any, ItemId: number): Promise<any> => {

        const SiteURL: SPFI = spfi(URL).using(SPFx(context)); 
        try {
            const item = await SiteURL.web.lists
                .getByTitle(ListName)
                .items.getById(ItemId)
                .update(postObject);
               return item;
        } catch (error) {
            console.error("Error updating item:", error);
            showToast("error", ActionStatus.Error);
            hideLoader();
            return null;
        }
    };
    // 🔹 Batch Add Multiple Items
    const batchAddItems = async (ListName: string, URL: string, items: any[]): Promise<void> => {
        const SiteURL: SPFI = spfi(URL).using(SPFx(context)); 
        try {
            const [batchedSP, execute] = SiteURL.batched();
            const list = batchedSP.web.lists.getByTitle(ListName);
            for (const item of items) {
                list.items.add(item);
            }
            await execute();
            //showToast("success", "Batch add completed successfully");
        } catch (error) {
            console.error("Error in batch add:", error);
            showToast("error", ActionStatus.Error);
            hideLoader();
        }
    };

    // 🔹 Batch Update Multiple Items
    const batchUpdateItems = async (ListName: string, URL: string, items: { Id: any; data: any }[]): Promise<void> => {
        const SiteURL: SPFI = spfi(URL).using(SPFx(context)); 
        try {
            const [batchedSP, execute] = SiteURL.batched();
            const list = batchedSP.web.lists.getByTitle(ListName);

            for (const item of items) {
                list.items.getById(item.Id).update(item.data);
            }
            await execute().then(() => {
               console.log("Batch update completed successfully");
            }).catch((error) => {
                console.error("Error executing batch update:", error);
                showToast("error", ActionStatus.Error);
                hideLoader();
            });
            //showToast("success", "Batch update completed successfully");
        } catch (error) {
            console.error("Error in batch update:", error);
            showToast("error", ActionStatus.Error);
            hideLoader();
        }
    };

   return {
        getListItems,
        addListItem,
        updateListItem,
        batchAddItems,
        batchUpdateItems
    };

};

