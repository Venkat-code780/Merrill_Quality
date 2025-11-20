import { SPFI, spfi } from "@pnp/sp";
import "@pnp/sp/batching";
// import { SPHttpClient } from "@microsoft/sp-http";
import { hideLoader } from "../Shared/Loader";
import { showToast } from "../Shared/Toaster";
import { ActionStatus } from "../Constants/Contants";
import { SPFx } from "@pnp/sp/presets/all";
// import { ISPHttpClientOptions, SPHttpClient } from "@microsoft/sp-http";



// Initialize SPFI instance for mayco site
export const initCommonFunctions = (context: any, siteAbsoluteURL: string) => {
    // get List Items
    const getListItems = async (ListName: string, URL: string, selectQuery: string = '', expand: string = '', filterQuery: string = ''): Promise<any[]> => {
        const SiteURL: SPFI = spfi(URL).using(SPFx(context));
        try {
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
            return error;
        }
    }
    // Get List Item by Id
    const getListItemById = async (ListName: string, URL: string, ItemId: number, selectQuery: string = '', expand: string = ''): Promise<any> => {
        const SiteURL: SPFI = spfi(URL).using(SPFx(context));
        try {
            const item = await SiteURL.web.lists
                .getByTitle(ListName)
                .items.getById(ItemId)
                .select(selectQuery)
                .expand(expand)();
            return item;
        } catch (error) {
            console.error("Error fetching item by Id:", error);
            showToast("error", ActionStatus.Error);
            hideLoader();
            return error;
        }
    };

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
            return error;
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
            return error;
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
            return error;
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
            return error;
        }
    };
    // getting SP group meebers
    const getGroupMemberEmails = async (groupName: string, URL: string): Promise<any[]> => {
        const Site: SPFI = spfi(URL).using(SPFx(context));
        try {
            const group = await Site.web.siteGroups.getByName(groupName);
            const users = await group.users();
            const members = users.map(user => (
                user.Email
            ));
            return members;
        } catch (error) {
            console.error("Error fetching group members:", error);
            showToast("error", ActionStatus.Error);
            hideLoader();
            return [];
        }
    };

    // Send Email using SharePoint PNP API
    const sendEmail = async (URL: string, to: string[], subject: string, body: string, cc?: string[]): Promise<void> => {
        const sp: SPFI = spfi(URL).using(SPFx(context));

        try {
            let EmailRes = await sp.utility.sendEmail({
                To: to,
                Subject: subject,
                Body: body,
                CC: cc || []
            });
            return EmailRes;
        } catch (error) {
            console.error("Error sending email:", error);
            showToast("error", ActionStatus.Error);
            hideLoader();
            return error;
        }
    };

    // const sendEmail = async (URL:string,to: string[],subject: string,body: string,cc?: string[]): Promise<void> => {
    //   const emailProps = {
    //     properties: {
    //     //   To: { results: to },
    //     //   CC: { results: cc || [] },
    //       To:to,
    //       CC:cc || [],
    //       Subject: subject,
    //       Body: body
    //     }
    //   };

    //   const requestOptions: ISPHttpClientOptions = {
    //     body: JSON.stringify(emailProps),
    //     headers: {
    //       "Accept": "application/json;odata=verbose",
    //       "Content-Type": "application/json;odata=verbose"
    //     }
    //   };

    //   try {
    //     const response = await context.spHttpClient.post(
    //       `${URL}/_api/SP.Utilities.Utility.SendEmail`,
    //       SPHttpClient.configurations.v1,
    //       requestOptions
    //     );

    //     if (response.ok) {
    //       console.log("Email sent successfully");
    //     } else {
    //       const errorText = await response.text();
    //       console.error("Failed to send email:", errorText);
    //       throw new Error(errorText);
    //     }
    //   } catch (error) {
    //     console.error("Error sending email:", error);
    //     throw error;
    //   }
    // }

    return {
        //CRUD Functions
        getListItems,
        getListItemById,
        addListItem,
        updateListItem,
        batchAddItems,
        batchUpdateItems,
        //Email Functions
        getGroupMemberEmails,
        sendEmail
    };

};

