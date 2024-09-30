// import { Injectable } from "@angular/core";
// //PNP SP dependent modules
// import { SPFx, spfi } from "@pnp/sp";
// import "@pnp/sp/webs";
// import "@pnp/sp/lists";
// import "@pnp/sp/items";
// import "@pnp/sp/items/get-all";
// import { createBatch } from "@pnp/sp/batching";
// import "@pnp/sp/folders";
// import "@pnp/sp/files/folder";

// import { RestApiService } from "./rest-api.service";
// import { from, switchMap } from "rxjs";

// @Injectable({ providedIn: "root" })
// export class PnpSPService {
//   sp: any;
//   constructor(private restService: RestApiService) {
//     this.restService.getFormDigest().subscribe((data: any) => {
//       const contextVal = {
//         ...data.d.GetContextWebInformation,
//         pageContext: {
//           web: {
//             absoluteUrl: this.restService.siteURL,
//           },
//         },
//       };
//       this.sp = spfi(this.restService.siteURL).using(SPFx(contextVal));
//     });
//   }


//   //process batch post request
//   processBatch(postData: any[], listName: string, listCols: any) {
//     return from(
//       (async () => {
//         const [batchedSP, execute] = this.sp.batched();
//         const list = batchedSP.web.lists.getByTitle(listName);
//         let res: any = [];

//         postData?.forEach((item) => {
//           const listItem = {};
//           Object.keys(listCols).forEach((colKey) => {
//             listItem[colKey] = item[listCols[colKey]];
//           });
//           list.items.add(listItem).then((r) => res.push(r));
//         });

//         await execute();
//         return res;
//       })()
//     );
//   }

//   getBigListData(
//     listName: string,
//     select: string[] = [""],
//     expand: string[] = [""],
//     orderBy: string[] = ["Id", "asc"]
//   ) {
//     return this.restService.getFormDigest().pipe(
//       switchMap((data: any) => {
//         const sp = spfi(this.restService.siteURL).using(
//           SPFx(data.d.GetContextWebInformation)
//         );
//         const url = sp.web.lists
//           .getByTitle(listName)
//           .items.select(...select)
//           .expand(...expand)
//           .orderBy(orderBy[0], orderBy[1] === "asc")
//           .getAll();
//         return url;
//       })
//     );
//   }

//   getLibraryData() {
//     return this.restService.getFormDigest().pipe(
//       switchMap((data: any) => {
//         const sp = spfi(this.restService.siteURL).using(
//           SPFx(data.d.GetContextWebInformation)
//         );
//         // const url = sp.web.getFolderByServerRelativePath("Shared Documents").files();
//         const url = sp.web.getFolderByServerRelativePath("Shared Documents").folders()
//         return url;
//       })
//     );
//   }
// }