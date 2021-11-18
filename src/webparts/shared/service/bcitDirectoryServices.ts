import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import { Web } from "@pnp/sp/webs";
import {ListColS} from "./ListColS"
import { MSGraphClient } from "@microsoft/sp-http";


//function to submit data into sharepoint list UserLog
 export class SPService  {
    
    private web;

    constructor(url: string) {
        
        this.web = Web(url);
    }
    public async createTask(listName: string, body: ListColS) {
        try {
           
            let createdItem = await this.web.lists
                .getByTitle('Users')
                .items
                .add(body);
            return createdItem;
        } catch (err) {
            Promise.reject(err);
        }
    }

}
        
        
 



 




