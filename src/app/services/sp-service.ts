import { HttpClient } from '@angular/common/http';
import { inject, Injectable } from '@angular/core';
import { Observable, forkJoin } from 'rxjs';
import { catchError, map, switchMap } from 'rxjs/operators';

export interface options {
  select: string;
  expand?: string;
  filter?: string;
  orderby?: string;
}

@Injectable({
  providedIn: 'root',
})
export class SpService {
  siteId: number | undefined;
  allTableDataObject: any;
  currentUser: any;
  // private http = inject(HttpClient);
  constructor(private http: HttpClient) {}

  getCurrUser() {
    return this.currentUser;
  }
  setCurrUser(user: any) {
    this.currentUser = user;
  }

  convertHTMLToStr(str: string) {
    let content: any = '';
    if (str && str.length > 0) {
      var element = document.createElement('div');
      element.innerHTML = str;
      content = element.textContent;
      element.remove();
    }
    return content;
  }

  get siteURL() {
    const prod = false;
    if(prod) {
      const index =
        window.location.href.toLowerCase().indexOf('sitepages') != -1
          ? window.location.href.toLowerCase().indexOf('sitepages')
          : window.location.href.toLowerCase().indexOf('siteassets');
      return window.location.href.substring(0, index);
    } else {
      return 'https://prakritimarine.sharepoint.com/sites/pmspvtltd/';
    }
  }

  getFormDigest(siteName: string = this.siteURL): Observable<any> {
    const headers = {
      Accept: 'application/json;odata=verbose',
      'Content-Type': 'application/json;odata=verbose',
    };

    return this.http.post(`${siteName}_api/contextinfo`, null, {
      headers,
    });
  }

  getListDataByQuery(listName: any, options?: string): Observable<any> {
    const url = `${
      this.siteURL
    }_api/lists/getbytitle('${listName}')/items?$top=5000${
      options ? '&' + options : ''
    }`;
    return this.http.get(url);
  }

  getListDataById(listname: string, id: number, options?: string) {
    const url = `${
      this.siteURL
    }_api/lists/getbytitle('${listname}')/items(${id})${
      options ? '?' + options : ''
    }`;
    return this.http.get(url);
  }

  getItemCount(listname: string, options?: string) {
    const url = `${this.siteURL}_api/lists/getbytitle('${listname}')/ItemCount${
      options ? '?' + options : ''
    }`;
    return this.http.get(url);
  }

  getuserdetails(
    listname: string,
    id: number,
    options?: string
  ): Observable<any> {
    const url = `${
      this.siteURL
    }_api/lists/getbytitle('${listname}')/items(${id})${
      options ? '?' + options : ''
    }`;
    return this.http.get(url).pipe(catchError(this.errorHandler));
  }

  getUserById(id: number): Observable<any> {
    const url = `${this.siteURL}_api/Web/GetUserById(${id})`;
    const headers = {
      Accept: 'application/json;odata=verbose',
      'Content-Type': 'application/json;odata=verbose',
    };
    return this.http.get(url, { headers }).pipe(catchError(this.errorHandler));
  }

  errorHandler(): any {
    return { error: true };
  }

  //Adding users to group
  addUserToGroup(groupId: number | undefined, loginName: string) {
    const item = {
      __metadata: {
        type: 'SP.User',
      },
      LoginName: loginName,
    };
    return this.getFormDigest().pipe(
      switchMap((data: any) => {
        const url = `${this.siteURL}_api/web/sitegroups('${groupId}')/users`;
        const headers = {
          'X-RequestDigest': data.d.GetContextWebInformation.FormDigestValue,
          Accept: 'application/json;odata=verbose',
          'content-type': 'application/json;odata=verbose',
        };
        return this.http.post(url, JSON.stringify(item), { headers });
      })
    );
  }

  removeUsersFromGroup(groupId: number | undefined, loginName: string) {
    const item = {
      loginName,
    };
    return this.getFormDigest().pipe(
      switchMap((data: any) => {
        const url = `${this.siteURL}_api/web/sitegroups('${groupId}')/users/removeByLoginName`;
        const headers = {
          'X-RequestDigest': data.d.GetContextWebInformation.FormDigestValue,
          Accept: 'application/json;odata=verbose',
          'content-type': 'application/json;odata=verbose',
        };
        return this.http.post(url, JSON.stringify(item), { headers });
      })
    );
  }

  getCurrentUserwithgroups(): Observable<any> {
    const url = `${this.siteURL}/_api/web/CurrentUser/?$expand=groups`;
    const headers = {
      Accept: 'application/json;odata=verbose',
      'Content-Type': 'application/json;odata=verbose',
    };
    return this.http.get(url, { headers });
  }

  getGroupUsers(groupName: string) {
    const url = `${this.siteURL}_api/Web/SiteGroups/GetByName('${groupName}')/users`;
    const headers = {
      Accept: 'application/json;odata=verbose',
      'Content-Type': 'application/json;odata=verbose',
    };
    return this.http.get(url, { headers });
  }
  getGroupInfo(groupName: string) {
    const url = `${this.siteURL}_api/Web/SiteGroups/GetByName('${groupName}')`;
    const headers = {
      Accept: 'application/json;odata=verbose',
      'Content-Type': 'application/json;odata=verbose',
    };
    return this.http.get(url, { headers });
  }

  getDocLibItem(listName: string, fileName: string) {
    const url = `ˀ`;
    return this.http.get(url);
  }

  postListItem(listname: string, item: {}): Observable<any> {
    return this.getFormDigest().pipe(
      switchMap((data: any) => {
        const url = `${this.siteURL}_api/lists/getbytitle('${listname}')/items`;
        const data1 = JSON.stringify(item);
        const headers = {
          'X-RequestDigest': data.d.GetContextWebInformation.FormDigestValue,
          Accept: 'application/json;odata=verbose',
          'content-type': 'application/json;odata=verbose',
        };
        return this.http.post(url, data1, { headers });
      })
    );
  }

  updateListItem(listname: string, id: number, item: {}): Observable<any> {
    return this.getFormDigest().pipe(
      switchMap((data: any) => {
        const url = `${this.siteURL}_api/lists/getbytitle('${listname}')/items(${id})`;
        const data1 = JSON.stringify(item);
        const headers = {
          'X-RequestDigest': data.d.GetContextWebInformation.FormDigestValue,
          Accept: 'application/json;odata=verbose',
          'content-type': 'application/json;odata=verbose',
          'X-HTTP-Method': 'MERGE',
          'IF-MATCH': '*',
        };
        return this.http.post(url, data1, { headers });
      })
    );
  }

  deleteListItem(listname: string, id: number): Observable<any> {
    return this.getFormDigest().pipe(
      switchMap((data: any) => {
        const url = `${this.siteURL}_api/lists/getbytitle('${listname}')/items(${id})`;
        const headers = {
          'X-RequestDigest': data.d.GetContextWebInformation.FormDigestValue,
          Accept: 'application/json;odata=verbose',
          'content-type': 'application/json;odata=verbose',
          'X-HTTP-Method': 'MERGE',
          'IF-MATCH': '*',
        };
        return this.http.delete(url, { headers });
      })
    );
  }

  getFieldsData(listName: string, ...fieldNames: string[]) {
    let url = `${this.siteURL}_api/lists/getbytitle('${listName}')/fields?$select=choices,Description&$filter=`;
    fieldNames.forEach((e) => {
      if (e !== fieldNames[0]) {
        url += `or`;
      }
      url += `(EntityPropertyName eq '${e}')`;
    });
    return this.http.get(url);
  }

  getCurrentUser(): Observable<any> {
    let url = `${this.siteURL}_api/web/CurrentUser`;
    const headers = {
      Accept: 'application/json;odata=verbose',
      'Content-Type': 'application/json;odata=verbose',
    };
    return this.http.get(url, { headers });
  }

  getCurrentUserInfo(): Observable<any> {
    return this.getFormDigest().pipe(
      switchMap((data: any) => {
        let url = `${this.siteURL}_api/web/CurrentUser`,
          urlGrp = `${this.siteURL}/_api/web/CurrentUser/?$expand=groups`;
        const headers = {
          Accept: 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose',
        };
        return forkJoin([
          this.http.get(url, { headers }),
          this.http.get(urlGrp, { headers }),
          // this.getListDataByQuery(
          //   'Configuration',
          //   `$select=Id,Title,Value&$filter=Title eq 'Owner' or Title eq 'Global'`
          // ),
        ]).pipe(
          map((res: any) => {
            const userInfo = res[0]?.d,
              grpInfo = res[1],
              configInfo = res[2]?.value || 'Owner',
              adminCongig = configInfo.find(
                (config: any) => config.Title === 'Owner'
              )?.Value,
              globalConfig = configInfo.find(
                (config: any) => config.Title === 'Global'
              )?.Value;
            userInfo['currentUserGrps'] = grpInfo?.d?.Groups?.results;
            if (userInfo['currentUserGrps'].length) {
              const adminGrp = userInfo['currentUserGrps'].find(
                (grps: any) => grps.LoginName === adminCongig
              );
              const globalGrp = userInfo['currentUserGrps'].find(
                (grps: any) => grps.LoginName === globalConfig
              );
              userInfo['isAdmin'] = adminGrp !== undefined;
              userInfo['isGlobalUser'] = globalGrp !== undefined;
            }
            return userInfo;
          })
        );
      })
    );
  }

  getAllGroups(group?: string | number, options?: string): Observable<any> {
    let url = `${this.siteURL}_api/web/sitegroups`;
    url += typeof group === 'string' ? `/getbyname('${group}')` : `(${group})`;
    url += `/users${options ? '?' + options : ''}`;
    const headers = {
      Accept: 'application/json;odata=verbose',
      'Content-Type': 'application/json;odata=verbose',
    };
    return this.http.get(url, { headers });
  }

  getAllUsers(queryString: string): Observable<any> | undefined {
    if (queryString.length > 2) {
      return this.getFormDigest().pipe(
        switchMap((data: any) => {
          const item = {
            queryParams: {
              __metadata: {
                type: 'SP.UI.ApplicationPages.ClientPeoplePickerQueryParameters',
              },
              AllowEmailAddresses: true,
              AllowMultipleEntities: false,
              AllUrlZones: false,
              MaximumEntitySuggestions: 50,
              PrincipalSource: 15,
              PrincipalType: 15,
              QueryString: queryString,
              Required: false,
              UrlZoneSpecified: false,
            },
          };
          let url = `${this.siteURL}_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.ClientPeoplePickerSearchUser`;
          const data1 = JSON.stringify(item);
          const headers = {
            Accept: 'application/json;odata=verbose',
            'X-RequestDigest': data.d.GetContextWebInformation.FormDigestValue,
            'Content-Type': 'application/json;odata=verbose',
          };

          return this.http.post(url, data1, { headers });
        })
      );
    }
    return;
  }

  ensureUser(loginName: any): Observable<any> {
    return this.getFormDigest().pipe(
      switchMap((data: any) => {
        const payload = { logonName: loginName };
        let url = `${this.siteURL}_api/web/ensureuser`;
        const data1 = JSON.stringify(payload);
        const headers = {
          'X-RequestDigest': data.d.GetContextWebInformation.FormDigestValue,
          accept: 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose',
        };
        return this.http.post(url, data1, { headers });
      })
    );
  }

  //Document Library Methods
  getFileByPath(folder: string, fileName: string) {
    const url = `${this.siteURL}_api/web/GetFolderByServerRelativeUrl('${folder}')/Files('${fileName}')/$value`;
    return this.http.get(url);
  }

  getFileByPath2(path: string) {
    const url = `${
      this.siteURL
    }_api/web/GetFileByServerRelativeUrl('${this.siteURL.substring(
      this.siteURL.indexOf('/sites/')
    )}${path}')/$value`;

    return this.http.get(url);
  }

  // async getDataFromFile(fileURI: string) {
  //   let response = await fetch(fileURI);
  //   if (response.ok) {
  //     let data = await response.blob();
  //     let file = new File([data], 'response.json');

  //     return new Promise((resolve, reject) => {
  //       const reader = new FileReader();
  //       reader.onload = (e) => {
  //         resolve((e.target as FileReader).result);
  //       };
  //       reader.onerror = (e) => {
  //         reject((e.target as FileReader).error);
  //       };
  //       reader.readAsText(file);
  //     });
  //   }
  //   return;
  // }

  // getAttachments(listName: string, itemId: Number): Observable<any> {
  //   let url = ${this.siteURL}_api/web/lists/getbytitle('${listName}')/items('${itemId}')/AttachmentFiles;
  //   const headers = {
  //     Accept: 'application/json;odata=verbose',
  //     'Content-Type': 'application/json;odata=verbose',
  //   };
  //   return this.http.get(url, { headers });
  // }

  // checkFileExists(
  //   listname: string,
  //   id: number,
  //   file: any,
  //   attachmentFile: any[]
  // ) {
  //   const attachment = attachmentFile.find((attachment: any) => {
  //     return file.name === attachment.FileName;
  //   });
  //   if (attachment) {
  //     return this.deleteAttachment(listname, id, attachment.FileName, file);
  //   } else {
  //     return this.postAttachment(listname, id, file);
  //   }
  // }

  // deleteAttachment(
  //   listname: string,
  //   Id: number,
  //   attachmentFile: any,
  //   file?: any
  // ) {
  //   return this.getFormDigest().pipe(
  //     switchMap((data: any): any => {
  //       const url = ${this.siteURL}_api/lists/getbytitle('${listname}')/GetItemById(${Id})/AttachmentFiles/getByFileName('${attachmentFile}');
  //       const headers = {
  //         'X-RequestDigest': data.d.GetContextWebInformation.FormDigestValue,
  //         Accept: 'application/json;odata=verbose',
  //         'content-type': 'application/json;odata=verbose',
  //         'X-HTTP-Method': 'DELETE',
  //         'IF-MATCH': '*',
  //       };
  //       return this.http.post(url, null, { headers });
  //       // if (file) {
  //       //   return this.http.post(url, null, { headers }).pipe(
  //       //     map((res) => {
  //       //       return this.postAttachment(listname, Id, file);
  //       //     })
  //       //   );
  //       // }
  //     })
  //   );
  // }

  // postAttachment(listname: string, id: number, file: any) {
  //   return this.getFormDigest().pipe(
  //     switchMap((data: any) => {
  //       return this.getDocFileBuffer(file).then((buffer: any) => {
  //         return $.ajax({
  //           url: ${this.siteURL}_api/lists/getbytitle('${listname}')/items(${id})/AttachmentFiles/add(FileName='${file.name}'),
  //           method: 'POST',
  //           data: buffer,
  //           processData: false,
  //           headers: {
  //             Accept: 'application/json; odata=verbose',
  //             'content-type': 'application/json; odata=verbose',
  //             'X-RequestDigest':
  //               data.d.GetContextWebInformation.FormDigestValue,
  //           },
  //         });
  //       });
  //     })
  //   );
  // }

  //Utility Methods
  // getDocFileBuffer(file: File) {
  //   return new Promise((resolve, reject) => {
  //     const reader = new FileReader();
  //     reader.onload = (e) => {
  //       resolve((e.target as FileReader).result);
  //     };
  //     reader.onerror = (e) => {
  //       reject((e.target as FileReader).error);
  //     };
  //     reader.readAsArrayBuffer(file);
  //   });
  // }

  // async getFileBuffer(uri: string) {
  //   let response = await fetch(uri);
  //   let data = await response.blob();
  //   let file = new File([data], 'response.json');
  //   const deferred = $.Deferred();
  //   const reader = new FileReader();
  //   reader.onload = (e) => {
  //     deferred.resolve((e.target as FileReader).result);
  //   };
  //   reader.onerror = (e) => {
  //     deferred.reject((e.target as FileReader).error);
  //   };
  //   reader.readAsText(file);
  //   return deferred.promise();
  // }

  formatDate(dateObj: Date) {
    const today = new Date(dateObj);
    const year = today.getFullYear().toString();
    const month = ('0' + (today.getMonth() + 1).toString()).slice(-2);
    const date = ('0' + today.getDate().toString()).slice(-2);
    return [month, date, year].join('/');
  }

  formatNumber(value: string) {
    if (value && value?.toString()?.trim().length) {
      const parts = value.toString().split('.');
      parts[0] = parts[0]
        ?.replace(/\D/g, '')
        .replace(/\B(?=(\d{3})+(?!\d))/g, ',');
      return parts.join('.') || '';
    }
    return '';
  }

  getChoices(listName: string, ...fieldNames: string[]) {
    let url = `${this.siteURL}_api/lists/getbytitle('${listName}')/fields?$select=choices&$filter=`;
    fieldNames.forEach((e) => {
      if (e !== fieldNames[0]) {
        url += 'or';
      }
      url += `(EntityPropertyName eq '${e}')`;
    });
    return this.http.get(url);
  }

  clear(table: any) {
    table.clear();
    if (table.filters) {
      if (table.filters.global && table.filters.global.value) {
        table.filters.global.value = '';
      }
      const data = Object.keys(table.filters);
      data.forEach((e) => {
        table.filters[e].value = '';
      });
    }
  }
}
