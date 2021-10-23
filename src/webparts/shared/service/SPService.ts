import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import { Web } from "@pnp/sp/webs";
import {ListCol} from "./ListCol"

 





 export class SPService {

    private web;

    constructor(url: string) {
        this.web = Web(url);
    }
    public async createTask(listName: string, body: ListCol) {
        try {
            let createdItem = await this.web.lists
                .getByTitle('UsersLog')
                .items
                .add(body);
            return createdItem;
        }
        catch (err) {
            Promise.reject(err);
        }
    }
}
