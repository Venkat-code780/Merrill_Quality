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
// import { SPHttpClient } from "@microsoft/sp-http";
import { hideLoader } from "../Shared/Loader";
import { showToast } from "../Shared/Toaster";
import { ActionStatus } from "../Constants/Contants";

// Initialize SPFI instance for mayco site
export const initCommonFunctions = (context: any,siteAbsoluteURL:string)=> {
    const getListItems= async (ListName:string,selectQuery:string,filterQuery:string,expand:string,URL:string) : Promise<any[]> => 
        {
        const SiteURL: SPFI = spfi(URL).using(SPFx(context)); 
            try{
            const items = await SiteURL.web.lists
                .getByTitle(ListName)
                .items.top(5000)
                .select(selectQuery)
                .filter(filterQuery)
                .expand(expand)
                .orderBy("Title", true)()
                 return items;
        } catch (error) {
            console.error("Error fetching List:", error);
            showToast("error", ActionStatus.Error);
            hideLoader();
            return [];
        }
    }

   return {
        getListItems
    };

};

