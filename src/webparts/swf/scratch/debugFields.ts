import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";

export async function debugToDoFields(context: WebPartContext) {
    const sp = spfi().using(SPFx(context));
    try {
        const fields = await sp.web.lists.getByTitle("ToDo").fields.select("Title", "InternalName", "TypeAsString")();
        console.log("--- ToDo List Fields ---");
        fields.forEach(f => {
            console.log(`${f.Title} -> ${f.InternalName} (${f.TypeAsString})`);
        });
        console.log("-----------------------");
    } catch (e) {
        console.error("Failed to fetch ToDo fields:", e);
    }
}
