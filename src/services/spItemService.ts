import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/batching";
import { IItems } from "@pnp/sp/items";

export default class SPListService {

    public static async GetItems(listName: string, filter: string, spfi: SPFI): Promise<IItems>{
       
        const pnpList = await spfi.web.getList(listName);
        return await pnpList.items.filter(filter)
    }
}