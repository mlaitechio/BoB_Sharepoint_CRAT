import { IServices } from "./IServices";
import {
    IAttachmentFileInfo, IAttachmentInfo, IItem, IItemAddResult, IItemUpdateResult,
    ISiteUserInfo, SPFI, spfi, SPFx as spSPFX, ControlMode, IFile, Web, IFileInfo,
    PagedItemCollection
} from '@pnp/sp/presets/all'
import { ISearchQuery, SearchResults, SearchQueryBuilder, QueryPropertyValueType } from "@pnp/sp/search";
import "@pnp/sp/batching";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/security/web";
import "@pnp/sp/site-users/web";
import { IList } from "@pnp/sp/lists";
import { Constants } from "../Constants/Constants";
import { WebPartContext } from "@microsoft/sp-webpart-base";

import { SPHttpClient, SPHttpClientResponse, MSGraphClientV3 } from '@microsoft/sp-http'
import { GraphBrowser, GraphFI, graphfi, SPFx as graphSPFx } from "@pnp/graph/presets/all";
import { error } from "pdf-lib";

let sp: SPFI;


export class Services implements IServices {

    private context: WebPartContext;
    private graph: GraphFI;

    public constructor(context: any) {
        sp = spfi().using(spSPFX(context));
        this.context = context;
        this.graph = graphfi().using(graphSPFx(context));
    }

    public async getListItemById(serverRelativeUrl: string, listName: string, itemID: number): Promise<any> {

        let listItem = await sp.web.getList(`${serverRelativeUrl}/Lists/${listName}`).items.getById(itemID)().then((val) => {
            return Promise.resolve(val);
        }).catch((error) => {
            return Promise.reject(error);
        })

        return listItem;

    }

    public async getPagedListItems(serverRelativeUrl: string, listName: string, selectColumns: string, filterString: string, expandColumns: string, orderByColum: string, asc: boolean = true): Promise<any> {
        try {
            let selectQuery: any[] = ['Id'];
            let expandQuery: any[] = [];
            let listItems = [];

            let items: PagedItemCollection<any[]> = undefined;
            do {
                if (!items) items = await sp.web.getList(`${serverRelativeUrl}/Lists/${listName}`).items.select(selectColumns).
                    orderBy(`${orderByColum}`, false).expand(expandColumns).top(4000).getPaged();
                else {
                    items = await items.getNext();
                }
                if (items.results.length > 0) {
                    listItems = listItems.concat(items.results);
                }
            } while (items.hasNext);

            return Promise.resolve(listItems);
        } catch (err) {
            return Promise.reject(err);
        }
    }

    public async updateItemBatch(serverRelativeUrl, listName, itemIDs: any[], items: any[], departmentMapping: any): Promise<any[]> {

        const [batchedSP, execute] = sp.batched();

        const list = batchedSP.web.getList(`${serverRelativeUrl}/Lists/${listName}`);

        const itemUpdated = [];

        for (let i = 0; i < itemIDs.length; i++) {

            let currentItem = items.filter((val) => {
                return itemIDs[i].ID == val.ID
            })

            let department = departmentMapping?.Department ?? ``;
            let migratedDepartment = departmentMapping?.Title ?? ``;

            //&& migratedDepartment != ""
            if (department != "") {
                let listUpdate = list.items.getById(itemIDs[i].ID).
                    update({
                        Department: department,
                        //MigratedDepartment: department //migratedDepartment
                    }, `*`).then(b => {
                        console.log(`Item Updated:`, itemIDs[i].ID);
                        itemUpdated.push(b)
                    }).catch((error) => {
                        console.log(error)
                    })
            }

        }

        await execute();

        return Promise.resolve(itemUpdated);


    }

    public async updateItem(serverRelativeUrl: string, listName: string, itemID: number, metadataValues: any, etag: any = "*"): Promise<any> {

        const updateItemResults = await sp.web.getList(`${serverRelativeUrl}/Lists/${listName}`).items.
            getById(itemID).update(metadataValues, etag).then(async (results: IItemUpdateResult) => {
                let item = await results.item().then((results) => {
                    return Promise.resolve(results)
                })
                return Promise.resolve(item);
            }).catch((error) => {
                return Promise.reject(error);

            });

        return updateItemResults;
    }

    public async filterLargeListItem(serverRelativeUrl: string, listName: string, filterString: string): Promise<any> {
        let moreThan5KPromise = await sp.web.getList(`${serverRelativeUrl}/Lists/${listName}`).
            items.select('ID').filter(filterString)().then((filterItem) => {
                return Promise.resolve(filterItem)
            }).catch((error) => {
                return Promise.reject(error);
            })

        return moreThan5KPromise;
    }

    public async getCurrentUserInformation(userEmail: string, selectedColumns: string): Promise<any[]> {

        // "2cf9fef8-c7cb-48b4-be0c-43c958d4f658"
        // getById("2cf9fef8-c7cb-48b4-be0c-43c958d4f658")()
        // let myProfile = await sp.profiles.getPropertiesFor(`i:0#.f|membership|Aditya.Pal@bankofbaroda.co.in`).then((val) => {
        //     console.log(val)
        // }).catch((error) => {
        //     console.log(error)
        // })
        //${userEmail}
        let users = await this.graph.users.filter(`mail eq '${userEmail}'`).
            select(`${selectedColumns}`)().then((value) => {
                return value
            }).catch((error) => {
                return error;
            });

        return Promise.resolve(users);
    }

    public async deleteListItem(serverRelativeUrl: string, listName: string, itemId: number) {
        const deleteItem = await sp.web.getList(`${serverRelativeUrl}/Lists/${listName}`).items.getById(itemId).delete().then((results) => {
            return Promise.resolve(true)
        }).catch((error) => {
            return Promise.reject(error);
        });

        return deleteItem;
    }


    public async getFileContent(fileServerRelativeUrl): Promise<any> {

        let fileContent = await sp.web.getFileByServerRelativePath(fileServerRelativeUrl).getBuffer().then((file) => {
            return Promise.resolve(file)
        }).catch((error) => {
            return Promise.reject(error)
        });

        return fileContent;
    }

    public async addListItemAttachmentAsBuffer(listName: string, serverRelativeUrl: string, itemID: number, fileName: string, buffer: any): Promise<any> {

        let attachmentPromise = await sp.web.getList(`${serverRelativeUrl}/Lists/${listName}`).items
            .getById(itemID)
            .attachmentFiles.add(fileName, buffer).then((result) => {
                return Promise.resolve(result)
            }).catch((error) => {
                return Promise.reject(error);
            });

        return attachmentPromise;
    }

    public async addListItemAttachments(serverRelativeUrl: string, listName: string, itemID: number,
        fileMetadata: Map<string, any>): Promise<any> {

        let isFileAdded = false;
        let fileArray = [];

        const [batchedSP, execute] = sp.web.batched();

        fileMetadata.forEach(async (value, key) => {
            fileArray.push({
                "name": key,
                "content": value
            });


        });

        /**
         * This recursive call is working
         */
        // const attachmentsPromise = await this.addFileAsAttachment(fileArray, serverRelativeUrl, listName, itemID).then((value) => {
        //     return Promise.resolve(value)
        // }).catch((error) => {
        //     return Promise.reject(error)
        // })

        let attachmentsPromise: any[] = [];

        for (let i = 0; i < fileArray.length; i++) {
            const file = fileArray[i];
            const fileName = file.name;

            attachmentsPromise.push(
                await sp.web.getList(`${serverRelativeUrl}/Lists/${listName}`).items
                    .getById(itemID)
                    .attachmentFiles.add(fileName, file.content).then((result) => {
                        return Promise.resolve(result)
                    }).catch((error) => {
                        Promise.reject(error);
                    })
            );


        }


        //v2 pnpjs
        // const fileAttachmentResults = await sp.web.getList(`${serverRelativeUrl}/Lists/${listName}`).items.
        //     getById(itemID).attachmentFiles.addMultiple(fileArray).then((value) => {
        //         return Promise.resolve(value);
        //     }).catch((error) => {
        //         return Promise.reject(error);
        //     });

        return Promise.resolve(attachmentsPromise);

        /* Working for array of promises */

        // return await Promise.all(
        //     fileArray.map(itemFile => 
        //     this.postFile(serverRelativeUrl, listName, itemID, itemFile.name, itemFile.content, itemMetadata)));

    }

    private addFilesAsMultipleAttachment = async (file, serverRelativeUrl, listName, itemID): Promise<any> => {

        const item = await sp.web.getList(`${serverRelativeUrl}/Lists/${listName}`).items.getById(itemID);
        return await item.attachmentFiles.add(file.name, file.content).then((r) => {
            return Promise.resolve(r)
        }).catch((error) => {
            return Promise.reject(error);
        })

    }


    private addFileAsAttachment = async (files: IAttachmentFileInfo[], serverRelativeUrl, listName, itemID, index: number = 0):
        Promise<any> => {
        if (files && index < files.length) {
            const file = files[index];
            const recordtoaddto = await sp.web.getList(`${serverRelativeUrl}/Lists/${listName}`).items.getById(itemID);
            return recordtoaddto.attachmentFiles.add(file.name, file.content).then(r => {
                index++;
                this.addFileAsAttachment(files, serverRelativeUrl, listName, itemID, index);
                return Promise.resolve(r);
            }).catch((error) => {
                return Promise.reject(error)
            });
        }


    }

    public async deleteListItemAttachment(serverRelativeUrl: string, listName: string, itemID: number, fileName: string): Promise<any> {

        let deleteAttachmentPromise = await sp.web.getFolderByServerRelativePath(`${serverRelativeUrl}/Lists/${listName}/Attachments/${itemID}`).files.
            getByUrl(`${fileName}`).deleteWithParams({
                BypassSharedLock: true
            }).then((val) => {
                return Promise.resolve(val)
            }).catch((error) => {
                return Promise.reject(error);
            });

        // let checkInFile = await sp.web.getFileByServerRelativePath(`${serverRelativeUrl}/Lists/${listName}/Attachments/${itemID}/${fileName}`).
        //     checkin('Deleting this file').then((val) => {
        //         console.log(`File Checked In`)
        //     }).catch((error) => {
        //         console.log(error)
        //     })

        // let deleteAttachmentPromise = await sp.web.getList(`${serverRelativeUrl}/Lists/${listName}`).items
        //     .getById(itemID).attachmentFiles.getByName(fileName).recycle().then((attachmentVal) => {
        //         return Promise.resolve(attachmentVal)
        //     }).catch((error) => {
        //         return Promise.reject(error);
        //     });

        return deleteAttachmentPromise;
    }

    public async recycleListItemAttachments(serverRelativeUrl: string, listName: string, itemID: number,
        files: Map<string, any>): Promise<any> {
        let fileNames: string[] = [];
        files.forEach(element => {
            fileNames.push(element.name);
        });


        let recycleAttachmentPromise: any[] = [];

        for (let i = 0; i < fileNames.length; i++) {

            await sp.web.getList(`${serverRelativeUrl}/Lists/${listName}`).items
                .getById(itemID).attachmentFiles.getByName(fileNames[i]).recycle().then((result) => {
                    recycleAttachmentPromise.push(result)
                }).catch((error) => {
                    Promise.reject(error);
                });
        }

        // const fileAttachmentResults = await sp.web.getList(`${serverRelativeUrl}/Lists/${listName}`).items.getById(itemID).
        // attachmentFiles.recycleMultiple(...fileNames)
        //     .then(r => { console.log(r, ' Deleted Successfully!'); })
        //     .catch(reject => console.error('Error deleting attachments ', reject));

        // let attachmentFiles = await sp.web.getList(`${serverRelativeUrl}/Lists/${listName}`).items.getById(itemID).attachmentFiles
        // const recycleAttachmentPromise = await Promise.all(fileNames.map(async (file) => {
        //     return attachmentFiles.getByName(file).recycle().then((value) => {
        //         return Promise.resolve(value)
        //     }).catch((error) => {
        //         return Promise.reject(error)
        //     })

        // }))


        return Promise.resolve(recycleAttachmentPromise);

    }

    public async getAllListItemAttachments(serverRelativeUrl: string, listName: string, itemID: number,): Promise<Map<string, any>> {

        let allFiles = new Map<string, any>();
        const attachmentsPromise = await sp.web.getList(`${serverRelativeUrl}/Lists/${listName}`).items.getById(itemID).
            attachmentFiles.select()().then(async (attachments) => {
                await Promise.all(attachments.map(async (val) => {
                    const fileBufferPromise = await sp.web.getFileByServerRelativePath(val.ServerRelativeUrl).getBuffer().then((bufferVal) => {
                        allFiles.set(val.FileName, bufferVal);
                    });
                }));

                return Promise.resolve(allFiles);
            }).catch((error) => {
                return Promise.reject(error);
            });

        return attachmentsPromise
    };

    private recycleItemAttachments = async (item: IItem, file, serverRelativeUrl, listName): Promise<any> => {

        return await item.attachmentFiles.getByName(file.name).recycle().then((r) => {
            return Promise.resolve(r)
        }).catch((error) => {
            return Promise.reject(error);
        })

    }

    public async readFieldValues(serverRelativeUrl: string, listName: string, fieldName: string): Promise<any> {
        const fieldValues = await sp.web.getList(`${serverRelativeUrl}/Lists/${listName}`).fields.getByInternalNameOrTitle(fieldName)().then((value) => {
            return Promise.resolve(value);
        }).catch((error) => {
            return Promise.reject(error);
        });

        return fieldValues;
    }

    public async readListItemAttachment(serverRelativeUrl: string, listName: string, itemId: number): Promise<any[]> {
        let listItemAttachment = await sp.web.getList(`${serverRelativeUrl}/Lists/${listName}`).
            items.getById(itemId).attachmentFiles().then((attachments: IAttachmentInfo[]) => {
                return Promise.resolve(attachments)
            }).catch((error) => {
                return Promise.resolve([]);
            });

        return listItemAttachment;
    }


    public async getFileById(fileArray: any[]): Promise<any> {
        const fileAttachmentPromise = await Promise.all(fileArray.map(async (file) => {
            return await sp.web.getFileById(file.AttachmentId.replace('{', '').replace('}', ''))().then((fileInfo) => {
                return Promise.resolve(fileInfo)
            }).catch((error) => {
                return Promise.reject(error)
            });
        }))

        return fileAttachmentPromise
    }

    public async getAllFiles(folderServerRelativeUrl): Promise<any[]> {

        let files = await sp.web.getFolderByServerRelativePath(folderServerRelativeUrl).files().then((fileInfo) => {
            return Promise.resolve(fileInfo)
        }).catch((error) => {
            return Promise.reject(error)
        })

        return files
    }


    public async createItem(serverRelativeUrl: string, listName: string, metadataValues: any): Promise<boolean | any> {
        const isItemCreated = await sp.web.getList(`${serverRelativeUrl}/Lists/${listName}`).items.add(metadataValues).
            then((result: IItemAddResult): Promise<boolean> => {

                const item = result.data;
                if (item != null) {
                    return Promise.resolve(item);
                }

            }).catch(error => {
                return Promise.reject(error);
            });

        return isItemCreated;
    }

    public async checkIfUserBelongToGroup(groupName: any, userEmail: string): Promise<any> {

        const isUserPresent = sp.web.siteGroups.getByName(groupName).users().then((allUsers: ISiteUserInfo[]) => {
            const isUserMember = allUsers.some((user) => user.Email === userEmail);
            return Promise.resolve(isUserMember);
        }).catch(error => {
            return Promise.reject(error);
        });

        return isUserPresent;
    }

    public async getListDataAsStream(serverRelativeUrl: string, listName: string, itemID): Promise<any> {
        const listDataAsStream = await sp.web.getList(`${serverRelativeUrl}/Lists/${listName}`).
            renderListFormData(itemID, 'editform', ControlMode.Edit).then((data) => {
                return Promise.resolve(data)
            }).catch((error) => {
                return Promise.reject(error);
            })

        return listDataAsStream
    }

    public async renderListDataStream(serverRelativeUrl: string, listName: string, viewXML: string, query?: any): Promise<any> {

        const listDataDataPromise = await sp.web.getList(`${serverRelativeUrl}/Lists/${listName}`).renderListDataAsStream({
            ViewXml: viewXML
        }, null, query)



    }

    public async breakItemRoleInheritance(serverRelativeUrl: string, listName: string, itemID: any): Promise<any> {
        const { Id: readRoleDefId } = await sp.web.roleDefinitions.getByName(
            Constants.ListReadPermission
        )();
        const { Id: conRoleDefId } = await sp.web.roleDefinitions.getByName(
            Constants.ListContriPermission
        )();
        const { Id: fullRoleDefId } = await sp.web.roleDefinitions.getByName(
            Constants.ListFullPermission
        )();
        const list = sp.web.getList(`${serverRelativeUrl}/Lists/${listName}`);

        const user = await sp.web.currentUser();

        let permission = list.items.getById(itemID);
        await permission.breakRoleInheritance(false);
        await permission.roleAssignments.add(user.Id, conRoleDefId);
        await permission.roleAssignments.add(Constants.OwnerGroupID, fullRoleDefId);
        await permission.roleAssignments
            .add(Constants.EveryoneID, readRoleDefId)
            .then((value) => {
                return Promise.resolve(true);
            })
            .catch((error) => {
                return Promise.reject(false);
            });
        return Promise.resolve(permission);
        // const breakRoleInheritancePromise = await sp.web.getList(`${serverRelativeUrl}/Lists/${listName}`).items.getById(itemID)
        //     .breakRoleInheritance(true).then((value) => {
        //         return Promise.resolve(true)
        //     }).catch((error) => {
        //         return Promise.reject(false)
        //     })
    };

    public async getSearchResults(queryText: string, selectedProperties: any[], queryTemplate?: string, refinementFilters?: string, sortList?: any[]): Promise<any> {

        // const queryBuilder =  SearchQueryBuilder();

        // queryBuilder.selectProperties(selectedProperties.join(','))
        // queryBuilder.rowLimit(500)
        // queryBuilder.rowsPerPage(10);
        // queryBuilder.refinementFilters(`RefinableString04:"FY 2019 - 2020"`)

        try {
            let searchItems: any[] = [];
            let textQuery = queryText.trim() != "" ? `(${queryText?.trim().split(' ').join(' OR ')}) XRANK(cb=100)` : `*`; //`${queryText?.trim().split(' ').join(' OR ')}` + `*` : `*`

            let _searchQuerySettings: ISearchQuery = {
                Querytext: `${textQuery}`,//`*`,//
                TrimDuplicates: false,
                QueryTemplate: queryTemplate,
                RowLimit: 500,
                RowsPerPage: 10,
                ClientType: 'ContentSearchRegular',
                EnableSorting: true,
                SortList: sortList,
                // BypassResultTypes: true,
                // ClientType: "sug_SPListInline",
                // SummaryLength: 100,
                EnableInterleaving: true,
                Properties: [{
                    Name: "EnableDynamicGroups",
                    Value: {
                        BoolVal: true,
                        QueryPropertyValueTypeIndex: QueryPropertyValueType.BooleanType
                    }
                }],
                //Culture:57,

                //[`RefinableString04:("FY 2023 - 2024*")`],RefinableDate00:range(2020-11-01T00:01:01Z,2023-12-31T00:01:01Z)`RefinableString04:equals("FY 2020 - 2021")`
                SelectProperties: selectedProperties,
                // SourceId: "264617d4-bb7d-463e-b494-bff7fded0f64" //List ID of Bulletin Board
            };

            if (refinementFilters.trim() != "") {
                _searchQuerySettings.RefinementFilters = [refinementFilters]
            }

            console.log(_searchQuerySettings);


            let searchResults = await sp.search(_searchQuerySettings);

            searchItems = searchItems.concat(searchResults.PrimarySearchResults);



            // Check if there are more items to retrieve
            // while (searchResults.TotalRowsIncludingDuplicates - 1 > searchItems.length) {
            while (searchItems.length < searchResults.TotalRowsIncludingDuplicates) {
                _searchQuerySettings.StartRow = searchItems.length
                searchResults = await sp.search(_searchQuerySettings);
                // Add the next batch of items to the array
                searchItems = searchItems.concat(searchResults.PrimarySearchResults);
            }

            return Promise.resolve(searchItems);
        }
        catch {
            return Promise.reject(`Error occured while performing search`)
        }

    }

    //queryText: string, batchSize: number, startRow: number
    private async getAllSearchResults(_searchQuerySettings: ISearchQuery, allResults: any[] = []): Promise<any[]> {
        const searchResults = await sp.search(_searchQuerySettings);

        // Add the current batch of items to the array
        allResults = allResults.concat(searchResults.PrimarySearchResults);

        // Check if there are more items to retrieve
        if (searchResults.TotalRows > allResults.length) {

            return this.getAllSearchResults(_searchQuerySettings, allResults);
        }

        return allResults;
    }

    public async getListInfo(serverRelativeUrl: string, listName: string): Promise<any> {

        const getListInfoPromise = await sp.web.getList(`${serverRelativeUrl}/Lists/${listName}`).select('Id,Title')().then((value) => {
            return Promise.resolve(value)
        }).catch((error) => {
            return Promise.reject(error)
        });

        return getListInfoPromise;

    }

    public async getLargeListItems(serverRelativeUrl: string, listName: string, selectedColumn, expandColumns): Promise<any[]> {
        var largeListItems: any[] = [];

        return new Promise<any[]>(async (resolve, reject) => {
            // Array to hold async calls  
            const asyncFunctions = [];

            let finalItems: any[] = [];
            let items: PagedItemCollection<any[]> = undefined;
            do {
                if (!items) items = await sp.web.getList(`${serverRelativeUrl}/Lists/${listName}`).items.select(selectedColumn)
                    .expand(expandColumns).top(2000).getPaged();
                else items = await items.getNext();
                if (items.results.length > 0) {
                    finalItems = finalItems.concat(items.results);
                }
            } while (items.hasNext);

            resolve(finalItems);

            // this.getLatestItemId(serverRelativeUrl, listName).then(async (itemCount: number) => {
            //     for (let i = 0; i < Math.ceil(itemCount / 5000); i++) {
            //         // Make multiple async calls  
            //         let resolvePagedListItems = () => {
            //             return new Promise(async (resolve) => {
            //                 let pagedItems: any[] = await this.getPageListItems(listName, serverRelativeUrl, i, selectedColumn, expandColumns);
            //                 resolve(pagedItems);
            //             })
            //         };
            //         asyncFunctions.push(resolvePagedListItems());
            //     }

            //     // Wait for all async calls to finish  
            //     const results: any = await Promise.all(asyncFunctions);
            //     for (let i = 0; i < results.length; i++) {
            //         largeListItems = largeListItems.concat(results[i]);
            //     }

            //     resolve(largeListItems);
            // });
        });
    }

    private getPageListItems(listName: string, serverRelativeUrl, index: number, selectedColumn: string, expandedColumn): Promise<any[]> {
        return new Promise<any[]>((resolve, reject): void => {

            let requestUrl = this.context.pageContext.web.absoluteUrl
                + `/_api/Web/GetList('${serverRelativeUrl}/Lists/${listName}')/items`
                + `?$skiptoken=Paged=TRUE%26p_ID=` + (index * 5000 + 1)
                + `&$top=` + 5000
                + `&$select=${selectedColumn}` + `&$expand=${expandedColumn}`;

            this.context.spHttpClient.get(requestUrl, SPHttpClient.configurations.v1)
                .then((response: SPHttpClientResponse) => {
                    response.json().then((responseJSON: any) => {
                        resolve(responseJSON.value);
                    });
                });
        });
    }

    public async getLatestItemId(serverRelativeUrl: string, listName: string): Promise<any> {

        return new Promise<number>((resolve: (itemId: number) => void, reject: (error: any) => void): void => {
            sp.web.getList(`${serverRelativeUrl}/Lists/${listName}`)
                .items.orderBy('Id', false).top(1).select('Id')()
                .then((items: { Id: number }[]): void => {
                    if (items.length === 0) {
                        resolve(-1);
                    }
                    else {
                        resolve(items[0].Id);
                    }
                }, (error: any): void => {
                    reject(error);
                });
        });
    }

}