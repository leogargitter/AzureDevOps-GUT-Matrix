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

function updateGUTOnForm(storedFields:StoredFieldReferences) {
    getWorkItemFormService().then((service) => {
        service.getFields().then((fields: TFS_Wit_Contracts.WorkItemField[]) => {
            var matchingGravityFields  = fields.filter(field => field.referenceName === storedFields.gvField);
            var matchingUrgencyFields  = fields.filter(field => field.referenceName === storedFields.ugField);
            var matchingTendencyFields = fields.filter(field => field.referenceName === storedFields.tdField);
            var matchingGUTFields      = fields.filter(field => field.referenceName === storedFields.gutField);

            //If this work item type has GUT, then update GUT
            if ((matchingGravityFields.length  > 0) &&
                (matchingUrgencyFields.length  > 0) &&
                (matchingTendencyFields.length > 0) &&
                (matchingGUTFields.length      > 0)) {
                service.getFieldValues([storedFields.gvField, storedFields.ugField, storedFields.tdField, storedFields.multField]).then((values) => {
                    var Gravity  = +values[storedFields.gvField].toString().split('-')[0].trim();
                    var Urgency  = +values[storedFields.ugField].toString().split('-')[0].trim();
                    var Tendency = +values[storedFields.tdField].toString().split('-')[0].trim();
                    
                    var Multiplier = +values[storedFields.multField];

                    var gut = (Gravity * Urgency * Tendency);

                    if (gut <= 9){
                        gut = 4;
                    }
                    else if (gut <= 25){
                        gut = 3;
                    }
                    else if (gut < 50){
                        gut = 2;
                    }
                    else {
                        gut = 1;
                    }

                    service.setFieldValue(storedFields.gutField, gut);

                    let date: Date = new Date();
                    date.setDate(date.getDate() + (gut * Multiplier));
                    date.setHours(23);
                    date.setMinutes(59);
                    service.setFieldValue(storedFields.dateField, date)
                });
            }
        });
    });
}

function updateGUTOnGrid(workItemId, storedFields:StoredFieldReferences):IPromise<any> {
    let gutFields = [
        storedFields.gvField,
        storedFields.ugField,
        storedFields.tdField,
        storedFields.gutField
    ];
    
    var deferred = Q.defer();

    var client = TFS_Wit_Client.getClient();
    client.getWorkItem(workItemId, gutFields).then((workItem: TFS_Wit_Contracts.WorkItem) => {
        if (storedFields.gutField !== undefined && storedFields.tdField !== undefined) {     
            var Gravity  = +workItem.fields[storedFields.gvField].toString().split('-')[0].trim();
            var Urgency  = +workItem.fields[storedFields.ugField].toString().split('-')[0].trim();
            var Tendency = +workItem.fields[storedFields.tdField].toString().split('-')[0].trim();

            var gut = Gravity + Urgency + Tendency;

            if (gut <= 9){
                gut = 4;
            }
            else if (gut <= 25){
                gut = 3;
            }
            else if (gut < 50){
                gut = 2;
            }
            else {
                gut = 1;
            }

            var document = [{
                from: null,
                op: "add",
                path: '/fields/' + storedFields.gutField,
                value: gut
            }];

            // Only update the work item if the GUT has changed
            if (gut != workItem.fields[storedFields.gutField]) {
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
            deferred.reject("Unable to calculate GUT, please configure fields on the collection settings page.");
        }
    });

    return deferred.promise;
}

var formObserver = (context) => {
    return {
        onFieldChanged: function(args) {
            GetStoredFields().then((storedFields:StoredFieldReferences) => {
                if (storedFields && storedFields.gvField && storedFields.ugField && storedFields.tdField && storedFields.gutField) {
                    //If one of fields in the calculation changes
                    if ((args.changedFields[storedFields.gvField] !== undefined) || 
                        (args.changedFields[storedFields.ugField] !== undefined) ||
                        (args.changedFields[storedFields.tdField] !== undefined)) {
                            updateGUTOnForm(storedFields);
                        }
                }
                else {
                    console.log("Unable to calculate GUT, please configure fields on the collection settings page.");    
                }
            }, (reason) => {
                console.log(reason);
            });
        },
        
        onLoaded: function(args) {
            GetStoredFields().then((storedFields:StoredFieldReferences) => {
                if (storedFields && storedFields.gvField && storedFields.ugField && storedFields.tdField && storedFields.gutField && storedFields.dateField) {
                    updateGUTOnForm(storedFields);
                }
                else {
                    console.log("Unable to calculate GUT, please configure fields on the collection settings page.");
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
                if (storedFields && storedFields.gvField && storedFields.ugField && storedFields.tdField && storedFields.gutField && storedFields.dateField) {
                    var workItemIds = args.workItemIds;
                    var promises = [];
                    $.each(workItemIds, function(index, workItemId) {
                        promises.push(updateGUTOnGrid(workItemId, storedFields));
                    });

                    // Refresh view
                    Q.all(promises).then(() => {
                        VSS.getService(VSS.ServiceIds.Navigation).then((navigationService: IHostNavigationService) => {
                            navigationService.reload();
                        });
                    });
                }
                else {
                    console.log("Unable to calculate GUT, please configure fields on the collection settings page.");
                    //TODO: Disable context menu item
                }
            }, (reason) => {
                console.log(reason);
            });
        }
    };
}

let extensionContext = VSS.getExtensionContext();
VSS.register(`${extensionContext.publisherId}.${extensionContext.extensionId}.gut-work-item-form-observer`, formObserver);
VSS.register(`${extensionContext.publisherId}.${extensionContext.extensionId}.gut-contextMenu`, contextProvider);