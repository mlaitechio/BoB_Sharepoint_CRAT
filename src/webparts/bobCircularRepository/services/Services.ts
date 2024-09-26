import { IServices } from "./IServices";
import {
    IAttachmentFileInfo, IAttachmentInfo, IItem, IItemAddResult, IItemUpdateResult,
    ISiteUserInfo, SPFI, spfi, SPFx as spSPFX, ControlMode, IFile, Web, IFileInfo,
    PagedItemCollection,
    IEmailProperties
} from '@pnp/sp/presets/all'
import { ISearchQuery, SearchResults, SearchQueryBuilder, QueryPropertyValueType } from "@pnp/sp/search";
import "@pnp/sp/batching";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/security/web";
import "@pnp/sp/site-users/web";
import { IList, RenderListDataOptions } from "@pnp/sp/lists";
import { Constants } from "../Constants/Constants";
import { WebPartContext } from "@microsoft/sp-webpart-base";

import { SPHttpClient, SPHttpClientResponse, MSGraphClientV3 } from '@microsoft/sp-http'
import { GraphBrowser, GraphFI, graphfi, SPFx as graphSPFx } from "@pnp/graph/presets/all";
import { error } from "pdf-lib";
import { forEach } from "jszip";
import { HttpClient } from '@microsoft/sp-http';
import { SharePointFile } from "../Models/IModel";

let sp: SPFI;


export class Services implements IServices {

    private context: WebPartContext;
    private graph: GraphFI;

    public constructor(context: any) {
        sp = spfi().using(spSPFX(context));
        this.context = context;
        this.graph = graphfi().using(graphSPFx(context));
    }


    public async getAllUsersInformationFromGroup(groupName: string): Promise<any> {

        let groupUsersInfo = await sp.web.siteGroups.getByName(groupName).users().then(async (val) => {
            if (val.length > 0) {

                // let allUserDetails = await Promise.all(val.map(async (userInfo) => {
                //     return await sp.profiles.getPropertiesFor(userInfo.LoginName).then((userDetails) => {
                //         return userDetails;
                //     })

                // }));

                let allUserDetails = await Promise.all(val.map(async (userInfo) => {
                    return await await this.graph.users.filter(`mail eq '${userInfo.Email}'`).
                        select(`${Constants.adSelectedColumns}`)().then((user) => {
                            return user[0] ?? []
                        }).catch((error) => {
                            console.log(error);
                            return []
                        })
                }))

                return allUserDetails;

            }
            else {
                return [];
            }
        })

        return Promise.resolve(groupUsersInfo);
    };

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

            let items: PagedItemCollection<any[]> = null;
            do {
                if (!items) items = await sp.web.getList(`${serverRelativeUrl}/Lists/${listName}`).items.select(selectColumns).
                    orderBy(`${orderByColum}`, false).expand(expandColumns).top(4999).getPaged();
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

    public async updateMultipleListItem(serverRelativeUrl: string, listName: string, itemIDs: any[], metadata: any): Promise<any[]> {
        const [batchedSP, execute] = sp.batched();

        const list = batchedSP.web.getList(`${serverRelativeUrl}/Lists/${listName}`);

        const itemUpdated = [];

        for (let i = 0; i < itemIDs.length; i++) {

            let listUpdate = list.items.getById(itemIDs[i]).update(metadata, `*`).then((listItem) => {
                itemUpdated.push(listItem)
            }).catch((error) => {
                console.log(error)
            })
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
        fileArray: any[]): Promise<any> {

        let isFileAdded = false;

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
            const fileName = encodeURI(fileArray[i].name);

            attachmentsPromise.push(
                await sp.web.getList(`${serverRelativeUrl}/Lists/${listName}`).items
                    .getById(itemID)
                    .attachmentFiles.add(fileName, file).then((result) => {
                        return Promise.resolve(result)
                    }).catch((error) => {
                        Promise.reject(error);
                    })

                // await sp.web.getFolderByServerRelativePath(`${serverRelativeUrl}/Lists/${listName}/Attachments/${itemID}`).
                //     files.addUsingPath(fileName, file).then((val) => {
                //         return Promise.resolve(val)
                //     }).catch((error) => {
                //         return Promise.reject(error)
                //     })
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

    public async addFileToListItem(serverRelativeUrl: string, listName: string, itemID: number, fileArray: any[]): Promise<any> {

        const [batchedSP, execute] = sp.batched();

        const list = batchedSP.web.getList(`${serverRelativeUrl}/Lists/${listName}`);

        const listAttachmentUpdated = [];

        for (let index = 0; index < fileArray.length; index++) {
            const fileName = fileArray[index].name;
            const fileContent = fileArray[index];

            list.items.getById(itemID).attachmentFiles.
                add(fileName, fileContent).then((val) => {
                    listAttachmentUpdated.push(val)
                }).catch((error) => {
                    console.log(error)
                })

        }

        await execute();

        return Promise.resolve(listAttachmentUpdated);
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

    public async getAllListItemAttachments(serverRelativeUrl: string, listName: string, itemID: number): Promise<Map<string, any>> {

        let allFiles = new Map<string, any>();
        const attachmentsPromise = await sp.web.getList(`${serverRelativeUrl}/Lists/${listName}`).items.getById(itemID).
            attachmentFiles.select()().then(async (attachments) => {
                await Promise.all(attachments.map(async (val) => {
                    //getBuffer
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


        let options: RenderListDataOptions = RenderListDataOptions.EnableMediaTAUrls | RenderListDataOptions.ContextInfo | RenderListDataOptions.ListData | RenderListDataOptions.ListSchema;

        let values = `<Value Type='Counter'>${itemID}</Value>`;

        const viewXML: string = `
        <View Scope='RecursiveAll'>
            <Query>
                <Where>
                    <In>
                        <FieldRef Name='ID' />
                        <Values>
                            ${values}
                        </Values>
                    </In>
                </Where>
            </Query>
            <RowLimit>1</RowLimit>
        </View>`;

        const listDataAsStream = await sp.web.getList(`${serverRelativeUrl}/Lists/${listName}`).
            renderListFormData(itemID, 'editform', ControlMode.Edit).then((data) => {
                return Promise.resolve(data)
            }).catch((error) => {
                return Promise.reject(error);
            })

        // const listDataStream = await sp.web.getList(`${serverRelativeUrl}/Lists/${listName}`).
        //     renderListDataAsStream({ RenderOptions: options, ViewXml: viewXML }).then((val) => {
        //         console.log(val)
        //     }).catch((error) => {
        //         console.log(error)
        //     })


        return listDataAsStream
    }

    public async renderListDataStream(serverRelativeUrl: string, listName: string, listItemIDs: any[], query?: any): Promise<any> {

        //let options: RenderListDataOptions = RenderListDataOptions.EnableMediaTAUrls | RenderListDataOptions.ContextInfo | RenderListDataOptions.ListData | RenderListDataOptions.ListSchema;
        let values = listItemIDs.map(i => { return `<Value Type='Counter'>${i.ID}</Value>`; });
        const viewXML: string = `
        <View Scope='RecursiveAll'>
            <Query>
                <Where>
                    <In>
                        <FieldRef Name='ID' />
                        <Values>
                            ${values.join("")}
                        </Values>
                    </In>
                </Where>
            </Query>
            <RowLimit>${listItemIDs.length}</RowLimit>
        </View>`;





        // const getAllFiles = await sp.web.getFolderByServerRelativePath(`${serverRelativeUrl}/Lists/${listName}/Attachments/3153`).
        //     files().then((val) => {
        //         console.log(val)
        //     }).catch((error) => {
        //         console.log(error)
        //     })

        // const pdfURLs = await this.generatePdfUrls(["37"]).then((val) => {
        //     console.log(val)
        // }).catch((error) => {
        //     console.log(error)
        // })


        const listDataDataPromise = await sp.web.getList(`${serverRelativeUrl}/Lists/${listName}`).renderListDataAsStream({
            //RenderOptions: options,
            ViewXml: viewXML
        }, null, query).then((val) => {
            return Promise.resolve(val);
        }).catch((error) => {
            return Promise.reject(error);
        })

        return listDataDataPromise;


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


    public async sendEmail(emailAddress: string[], ccEmailAddress: string[], subject: string, body: any): Promise<any> {

        const emailProperties: IEmailProperties = {
            To: emailAddress,
            CC: ccEmailAddress,
            Subject: subject,
            Body: body,
            AdditionalHeaders: {
                "content-type": "text/html"
            }
        }

        let emailPromise = await sp.utility.sendEmail(emailProperties).then((emailVal) => {
            return Promise.resolve(emailVal)
        }).catch((error) => {
            return Promise.reject(error);
        });

        return emailPromise
    }

    public async getSearchResults(queryText: string, selectedProperties: any[], queryTemplate?: string, refinementFilters?: string, sortList?: any[], startRow?: any, rowLimit?: any): Promise<any> {

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
                StartRow: startRow,
                RowLimit: rowLimit, // maximum 500 can be row Limit
                RowsPerPage: rowLimit, // items per page
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


            let searchResults = await sp.search(_searchQuerySettings).then((val) => {
                return val;
            });

            // searchItems = searchItems.concat(searchResults);



            // Check if there are more items to retrieve
            // while (searchItems.length < searchResults.TotalRowsIncludingDuplicates) {
            //     _searchQuerySettings.StartRow = searchItems.length
            //     searchResults = await sp.search(_searchQuerySettings);
            //     // Add the next batch of items to the array
            //     searchItems = searchItems.concat(searchResults);
            // }

            return Promise.resolve(searchResults);
            //return Promise.resolve(searchItems);
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
            let items: PagedItemCollection<any[]> = null;
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

    public async convertDocxToPDF(serverRelativeUrl: string, listName: string, itemID: number, fileName: string): Promise<any> {

        let fileItemId = [];
        let capturePDFProcess = [];
        await sp.web.getFileByServerRelativePath(`${serverRelativeUrl}/Lists/${listName}/Attachments/${itemID}/${fileName}`).
            copyTo(`${serverRelativeUrl}/${Constants.sharedDocuments}/${fileName}`, true).then(async (metadata) => {

                capturePDFProcess.push(`Docx File ${fileName} copied successfully to Shared Documents Library`);

                await sp.web.getFileByServerRelativePath(`${serverRelativeUrl}/${Constants.sharedDocuments}/${fileName}`).getItem().then(async (item) => {
                    await item().then(async (item) => {

                        capturePDFProcess.push(`PDF File Item Found in SP Library`);

                        if (item) {
                            fileItemId.push(`${item.ID}`);
                            console.log(`Item ID of file`, `${item.ID}`);
                            capturePDFProcess.push(`Item ID of PDF File ${item.ID}`)
                        }

                        if (fileItemId.length > 0) {

                            await this.generatePdfUrls(fileItemId, serverRelativeUrl, listName).then(async (files) => {
                                console.log(`PDF URLs generated`);
                                capturePDFProcess.push(`Processing PDF File`);

                                await this.saveAsPDF(files).then(async (isOk) => {

                                    capturePDFProcess.push(`PDF File generated successfully`);

                                    let pdfFileName = fileName.replace(".docx", ".pdf");

                                    //add file as attachment
                                    if (files.length > 0) {
                                        await sp.web.getFileByServerRelativePath(`${serverRelativeUrl}/${Constants.sharedDocuments}/${pdfFileName}`).getBlob().
                                            then(async (blobVal) => {

                                                capturePDFProcess.push(`Getting PDF Blob from SP Library`);

                                                capturePDFProcess.push(`Adding PDF File as attachment back to list Item ${itemID}`)

                                                await sp.web.getList(`${serverRelativeUrl}/Lists/${listName}`).items.getById(itemID).
                                                    attachmentFiles.add(pdfFileName, blobVal).then(async (attachResult) => {

                                                        capturePDFProcess.push(`PDF file ${pdfFileName} attached successfully`);

                                                        await sp.web.getFileByServerRelativePath(`${serverRelativeUrl}/${Constants.sharedDocuments}/${fileName}`).
                                                            deleteWithParams({ BypassSharedLock: true }).then(async (val) => {

                                                                console.log(`Docx File Deleted`);
                                                                capturePDFProcess.push(`Docx File ${fileName} deleted successfully`);

                                                                await sp.web.getFileByServerRelativePath(`${serverRelativeUrl}/${Constants.sharedDocuments}/${pdfFileName}`).deleteWithParams({
                                                                    BypassSharedLock: true
                                                                }).then((val) => {
                                                                    console.log(`Deleted PDF File`);
                                                                    capturePDFProcess.push(`Deleted PDF File successfully`);

                                                                }).catch((error) => {
                                                                    console.log(`PDF File Not Deleted`);
                                                                    console.log(error);
                                                                    capturePDFProcess.push(`PDF File:${error}`);
                                                                    return Promise.reject(error)
                                                                })
                                                            }).catch((error) => {

                                                                console.log(`Docx file not deleted`);
                                                                console.log(error);
                                                                capturePDFProcess.push(`Docx File Delete Error:${error}`);
                                                                return Promise.reject(error)
                                                            })
                                                    }).catch((error) => {
                                                        return Promise.reject(error)
                                                    })
                                            }).catch((error) => {
                                                console.log(error);
                                                return Promise.reject(error)
                                            })
                                    }

                                }).catch((error) => {
                                    console.log(error)
                                    return Promise.reject(error)
                                })
                            }).catch((error) => {
                                console.log(error)
                                return Promise.reject(error)
                            });
                        }
                    })
                }).catch((error) => {
                    console.log(error);
                    return Promise.reject(error)
                });

            }).catch((error) => {
                console.log(error)
                return Promise.reject(error)
            })


        return Promise.resolve(capturePDFProcess)

    }

    private async saveAsPDF(files: SharePointFile[]): Promise<boolean> {
        let isOk = true;
        for (let i = 0; i < files.length; i++) {
            const file = files[i];
            let pdfUrl = file.serverRelativeUrl.replace("." + file.fileType, ".pdf");
            let exists = true;
            try {
                await sp.web.getFileByServerRelativePath(pdfUrl)();
                isOk = false;
            }
            catch (error) {
                exists = false;
            }

            if (!exists) {
                let response = await this.context.httpClient.get(file.pdfUrl, HttpClient.configurations.v1);

                if (response.ok) {
                    let blob = await response.blob();
                    await sp.web.getFileByServerRelativePath(file.serverRelativeUrl).copyTo(pdfUrl);
                    await sp.web.getFileByServerRelativePath(pdfUrl).setContentChunked(blob);
                    const item = await sp.web.getFileByServerRelativePath(pdfUrl).getItem("File_x0020_Type");
                    // Potential fix for edge cases where file type is not set correctly 

                    if (item["File_x0020_Type"] !== "pdf") {

                        await item.update({
                            "File_x0020_Type": "pdf"
                        });
                    }
                }

                else {
                    const error = await response.json();
                    console.log(error)
                }

            }

        }
        return isOk;

    }


    private async generatePdfUrls(listItemIds: string[], serverRelativeUrl: string, listName: string): Promise<SharePointFile[]> {

        //let web = Web(context.pageContext.web.absoluteUrl); 

        let options: RenderListDataOptions = RenderListDataOptions.EnableMediaTAUrls | RenderListDataOptions.ContextInfo | RenderListDataOptions.ListData | RenderListDataOptions.ListSchema;

        var values = listItemIds.map(i => { return `<Value Type='Counter'>${i}</Value>`; });

        const viewXml: string = ` 
        <View Scope='RecursiveAll'> 
            <Query> 
                <Where> 
                    <In> 
                        <FieldRef Name='ID' /> 
                        <Values> 
                            ${values.join("")} 
                        </Values> 
                    </In> 
                </Where> 
            </Query> 
            <RowLimit>${listItemIds.length}</RowLimit> 
        </View>`;

        //let listID = `2f08f1b2-a600-4a2f-b64c-80443725b364`//`67280d85-09b4-4f37-8c2b-0e42ea7a5fa1`;//`c93fb4cb-9dd1-4a5d-a757-617b1ba8b391`;//this.context.pageContext.list.id.toString() 


        let response = await sp.web.getList(`${serverRelativeUrl}/${Constants.sharedDocuments}`).renderListDataAsStream({ RenderOptions: options, ViewXml: viewXml }) as any;

        console.log(response);
        //"{.mediaBaseUrl}/transform/pdf?provider=spo&inputFormat={.fileType}&cs={.callerStack}&docid={.spItemUrl}&{.driveAccessToken}" 
        let pdfConversionUrl = response.ListSchema[".pdfConversionUrl"];
        let mediaBaseUrl = response.ListSchema[".mediaBaseUrl"];
        let callerStack = response.ListSchema[".callerStack"];
        let driveAccessToken = response.ListSchema[".driveAccessToken"];
        let pdfUrls: SharePointFile[] = [];
        response.ListData.Row.forEach(element => {

            let fileType = element[".fileType"];
            let spItemUrl = element[".spItemUrl"];
            let pdfUrl = pdfConversionUrl
                .replace("{.mediaBaseUrl}", mediaBaseUrl)
                .replace("{.fileType}", fileType)
                .replace("{.callerStack}", callerStack)
                .replace("{.spItemUrl}", spItemUrl)
                .replace("{.driveAccessToken}", driveAccessToken);

            let pdfFileName = element.FileLeafRef.replace(fileType, "pdf");
            pdfUrls.push({ serverRelativeUrl: element["FileRef"], pdfUrl: pdfUrl, fileType: fileType, pdfFileName: pdfFileName });

        });

        return pdfUrls;

    }



}