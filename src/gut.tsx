import Q = require("q");

import TFS_Wit_Contracts = require("TFS/WorkItemTracking/Contracts");
import TFS_Wit_Client = require("TFS/WorkItemTracking/RestClient");
import TFS_Wit_Services = require("TFS/WorkItemTracking/Services");

import { StoredFieldReferences } from "./gutModels";
 
function GetStoredFields(): IPromise<any> {
    var deferred = Q.defer();
    VSS.getService<IExtensionDataService>(VSS.ServiceIds.ExtensionData).then((dataService: IExtensionDataService) => {
        dataService.getValue<StoredFieldReferences>("storedFields").then((storedFields:StoredFieldReferences) => {
            if (storedFields) {
                console.log("Retrieved fields from storage");
                deferred.resolve(storedFields);
            }
            else {
                deferred.reject("Failed to retrieve fields from storage");
            }
        });
    });
    return deferred.promise;
}

function getWorkItemFormService()
{
    return TFS_Wit_Services.WorkItemFormService.getService();
}

function updateWSJFOnForm(storedFields:StoredFieldReferences) {
    getWorkItemFormService().then((service) => {
        service.getFields().then((fields: TFS_Wit_Contracts.WorkItemField[]) => {
            var matchingBusinessValueFields = fields.filter(field => field.referenceName === storedFields.gvField);
            var matchingTimeCriticalityFields = fields.filter(field => field.referenceName === storedFields.ugField);
            var matchingRROEValueFields = fields.filter(field => field.referenceName === storedFields.tdField);
            var matchingEffortFields = fields.filter(field => field.referenceName === storedFields.effortField); 
            var matchinggutFields = fields.filter(field => field.referenceName === storedFields.gutField);
            var roundTo: number = storedFields.roundTo;

            //If this work item type has WSJF, then update WSJF
            if ((matchingBusinessValueFields.length > 0) &&
                (matchingTimeCriticalityFields.length > 0) &&
                (matchingRROEValueFields.length > 0) &&
                (matchingEffortFields.length > 0) &&
                (matchinggutFields.length > 0)) {
                service.getFieldValues([storedFields.gvField, storedFields.ugField, storedFields.tdField, storedFields.effortField]).then((values) => {
                    var businessValue  = +values[storedFields.gvField];
                    var timeCriticality = +values[storedFields.ugField];
                    var rroevalue = +values[storedFields.tdField];
                    var effort = +values[storedFields.effortField];

                    var wsjf = 0;
                    if (effort > 0) {
                        wsjf = (businessValue + timeCriticality + rroevalue)/effort;
                        if(roundTo > -1) {
                            wsjf = Math.round(wsjf * Math.pow(10, roundTo)) / Math.pow(10, roundTo)
                        }
                    }
                    
                    service.setFieldValue(storedFields.gutField, wsjf);
                });
            }
        });
    });
}

function updateWSJFOnGrid(workItemId, storedFields:StoredFieldReferences):IPromise<any> {
    let gutFields = [
        storedFields.gvField,
        storedFields.ugField,
        storedFields.tdField,
        storedFields.effortField,
        storedFields.gutField
    ];

    var deferred = Q.defer();

    var client = TFS_Wit_Client.getClient();
    client.getWorkItem(workItemId, gutFields).then((workItem: TFS_Wit_Contracts.WorkItem) => {
        if (storedFields.gutField !== undefined && storedFields.tdField !== undefined) {     
            var businessValue = +workItem.fields[storedFields.gvField];
            var timeCriticality = +workItem.fields[storedFields.ugField];
            var rroevalue = +workItem.fields [storedFields.tdField];
            var effort = +workItem.fields[storedFields.effortField];
            var roundTo: number = storedFields.roundTo;

            var wsjf = 0;
            if (effort > 0) {
                wsjf = (businessValue + timeCriticality + rroevalue)/effort;
                if(roundTo > -1) {
                    wsjf = Math.round(wsjf * Math.pow(10, roundTo)) / Math.pow(10, roundTo)
                }
            }

            var document = [{
                from: null,
                op: "add",
                path: '/fields/' + storedFields.gutField,
                value: wsjf
            }];

            // Only update the work item if the WSJF has changed
            if (wsjf != workItem.fields[storedFields.gutField]) {
                client.updateWorkItem(document, workItemId).then((updatedWorkItem:TFS_Wit_Contracts.WorkItem) => {
                    deferred.resolve(updatedWorkItem);
                });
            }
            else {
                deferred.reject("No relevant change to work item");
            }
        }
        else
        {
            deferred.reject("Unable to calculate WSJF, please configure fields on the collection settings page.");
        }
    });

    return deferred.promise;
}

var formObserver = (context) => {
    return {
        onFieldChanged: function(args) {
            GetStoredFields().then((storedFields:StoredFieldReferences) => {
                if (storedFields && storedFields.gvField && storedFields.effortField && storedFields.ugField && storedFields.tdField && storedFields.gutField) {
                    //If one of fields in the calculation changes
                    if ((args.changedFields[storedFields.gvField] !== undefined) || 
                        (args.changedFields[storedFields.ugField] !== undefined) ||
                        (args.changedFields[storedFields.tdField] !== undefined) ||
                        (args.changedFields[storedFields.effortField] !== undefined)) {
                            updateWSJFOnForm(storedFields);
                        }
                }
                else {
                    console.log("Unable to calculate WSJF, please configure fields on the collection settings page.");    
                }
            }, (reason) => {
                console.log(reason);
            });
        },
        
        onLoaded: function(args) {
            GetStoredFields().then((storedFields:StoredFieldReferences) => {
                if (storedFields && storedFields.gvField && storedFields.effortField && storedFields.ugField && storedFields.tdField && storedFields.gutField) {
                    updateWSJFOnForm(storedFields);
                }
                else {
                    console.log("Unable to calculate WSJF, please configure fields on the collection settings page.");
                }
            }, (reason) => {
                console.log(reason);
            });
        }
    } 
}

var contextProvider = (context) => {
    return {
        execute: function(args) {
            GetStoredFields().then((storedFields:StoredFieldReferences) => {
                if (storedFields && storedFields.gvField && storedFields.effortField && storedFields.ugField && storedFields.tdField && storedFields.gutField) {
                    var workItemIds = args.workItemIds;
                    var promises = [];
                    $.each(workItemIds, function(index, workItemId) {
                        promises.push(updateWSJFOnGrid(workItemId, storedFields));
                    });

                    // Refresh view
                    Q.all(promises).then(() => {
                        VSS.getService(VSS.ServiceIds.Navigation).then((navigationService: IHostNavigationService) => {
                            navigationService.reload();
                        });
                    });
                }
                else {
                    console.log("Unable to calculate WSJF, please configure fields on the collection settings page.");
                    //TODO: Disable context menu item
                }
            }, (reason) => {
                console.log(reason);
            });
        }
    };
}

let extensionContext = VSS.getExtensionContext();
VSS.register(`${extensionContext.publisherId}.${extensionContext.extensionId}.wsjf-work-item-form-observer`, formObserver);
VSS.register(`${extensionContext.publisherId}.${extensionContext.extensionId}.wsjf-contextMenu`, contextProvider);