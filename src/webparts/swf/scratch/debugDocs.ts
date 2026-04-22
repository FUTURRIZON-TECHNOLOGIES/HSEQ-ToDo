import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";

export async function debugLibraryDocs(context: WebPartContext, recordId: number) {
    const sp = spfi().using(SPFx(context));
    const libraryTitle = "Training & Inductions Documents";
    
    try {
        const list = sp.web.lists.getByTitle(libraryTitle);
        const fields = await list.fields.select("Title", "InternalName", "TypeAsString")();
        console.log("--- Library Fields ---");
        fields.forEach(f => {
            if (f.Title.includes("Record") || f.InternalName.includes("Record")) {
                console.log(`MATCH: ${f.Title} -> ${f.InternalName} (${f.TypeAsString})`);
            }
        });

        const items = await list.items.select("*", "RecordID/Id").expand("RecordID").top(10)();
        console.log("--- Sample Items ---");
        items.forEach(item => {
            console.log(`Item ID ${item.Id}: Title=${item.Title}, RecordID=${JSON.stringify(item.RecordID)}`);
        });

    } catch (e) {
        console.error("Debug failed:", e);
    }
}
