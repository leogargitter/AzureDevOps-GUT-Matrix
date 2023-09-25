import Q = require("q");
import Controls = require("VSS/Controls");
import {Combo, IComboOptions} from "VSS/Controls/Combos";
import Menus = require("VSS/Controls/Menus");
import WIT_Client = require("TFS/WorkItemTracking/RestClient");
import Contracts = require("TFS/WorkItemTracking/Contracts");
import Utils_string = require("VSS/Utils/String");

import { StoredFieldReferences } from "./gutModels";

export class Settings {
    private _changeMade = false;
    private _selectedFields:StoredFieldReferences;
    private _fields:Contracts.WorkItemField[];
    private _menuBar = null;

    private getSortedFieldsList():IPromise<any> {
        var deferred = Q.defer();
        var client = WIT_Client.getClient();
        client.getFields().then((fields: Contracts.WorkItemField[]) => {
            this._fields = fields.filter(field => (field.type === Contracts.FieldType.Double  || 
                                                   field.type === Contracts.FieldType.Integer || 
                                                   field.type === Contracts.FieldType.String  ||
                                                   field.type === Contracts.FieldType.DateTime))
            var sortedFields = this._fields.map(field => field.name).sort((field1,field2) => {
                if (field1 > field2) {
                    return 1;
                }

                if (field1 < field2) {
                    return -1;
                }

                return 0;
            });
            deferred.resolve(sortedFields);
        });

        return deferred.promise;
    }

    private getFieldReferenceName(fieldName): string {
        let matchingFields = this._fields.filter(field => field.name === fieldName);
        return (matchingFields.length > 0) ? matchingFields[0].referenceName : null;
    }

    private getFieldName(fieldReferenceName): string {
        let matchingFields = this._fields.filter(field => field.referenceName === fieldReferenceName);
        return (matchingFields.length > 0) ? matchingFields[0].name : null;
    }

    private getComboOptions(id, fieldsList, initialField):IComboOptions {
        var that = this;
        return {
            id: id,
            mode: "drop",
            source: fieldsList,
            enabled: true,
            value: that.getFieldName(initialField),
            change: function () {
                that._changeMade = true;
                let fieldName = this.getText();
                let fieldReferenceName: string = (this.getSelectedIndex() < 0) ? null : that.getFieldReferenceName(fieldName);

                switch (this._id) {
                    case "Gravity":
                        that._selectedFields.gvField = fieldReferenceName;
                        break;
                    case "Urgency":
                        that._selectedFields.ugField = fieldReferenceName;
                        break;
                    case "Tendency":
                        that._selectedFields.tdField = fieldReferenceName;
                        break;
                    case "GUT":
                        that._selectedFields.gutField = fieldReferenceName;
                        break;
                    case "Date":
                        that._selectedFields.dateField = fieldReferenceName;
                        break;
                    case "Multiplier":
                        that._selectedFields.multField = fieldReferenceName;
                        break;
                }
                that.updateSaveButton();
            }
        };
    }

    private getNumeralComboOptions(id, source: number[], initialValue: number):IComboOptions {
        var that = this;
        const currentInitialValue = initialValue ? initialValue.toString(): null
        return {
            id: id,
            mode: "drop",
            source: source,
            enabled: true,
            value: currentInitialValue,
            change: function () {
                that._changeMade = true;
                let num: number = +(this.getText());
                that.updateSaveButton();
            }
        };
    }

    public initialize() {
        let hubContent = $(".hub-content");
        let uri = VSS.getWebContext().collection.uri + "_admin/_process";
        
        let descriptionText = "A Matriz GUT é uma ferramenta que auxilia na priorização de resolução de problemas (por isso é também conhecida como Matriz de Prioridades). A análise GUT é muito utilizada naquelas questões em que é preciso de uma orientação para tomar decisões complexas e que exigem a análise de vários problemas. Os fatores trabalhados com a Matriz GUT (Gravidade, Urgência e Tendência) são pontuados de 1 a 5:";
        let header = $("<div />").addClass("description-text bowtie").appendTo(hubContent);
        header = $("<div />").addClass("description-text bowtie").appendTo(hubContent);
        header.html(Utils_string.format(descriptionText));

        $("<img src='https://media.treasy.com.br/media/2018/02/como-montar-a-matriz-gut.png' />").addClass("description-image").appendTo(hubContent);
        
        descriptionText = "";
        header = $("<div />").addClass("description-text bowtie").appendTo(hubContent);
        header.html(Utils_string.format(descriptionText, "<a target='_blank' href='" + uri +"'>process hub</a>"));

        let container = $("<div />").addClass("gut-settings-container").appendTo(hubContent);

        var menubarOptions = {
            items: [
                { id: "save", icon: "icon-save", title: "Save the selected field" }   
            ],
            executeAction:(args) => {
                var command = args.get_commandName();
                switch (command) {
                    case "save":
                        this.save();
                        break;
                    default:
                        console.log("Unhandled action: " + command);
                        break;
                }
            }
        };
        this._menuBar = Controls.create<Menus.MenuBar, any>(Menus.MenuBar, container, menubarOptions);

        let gvContainer = $("<div />").addClass("settings-control").appendTo(container);
        $("<label />").text("Gravity").appendTo(gvContainer);

        let ugContainer = $("<div />").addClass("settings-control").appendTo(container);
        $("<label />").text("Urgency").appendTo(ugContainer);

        let tdContainer = $("<div />").addClass("settings-control").appendTo(container);
        $("<label />").text("Tendency").appendTo(tdContainer);

        let gutContainer = $("<div />").addClass("settings-control").appendTo(container);
        $("<label />").text("GUT Result Field").appendTo(gutContainer);

        let dateContainer = $("<div />").addClass("settings-control").appendTo(container);
        $("<label />").text("Target Date Field").appendTo(dateContainer);

        let multContainer = $("<div />").addClass("settings-control").appendTo(container);
        $("<label />").text("Multiplier Field").appendTo(multContainer);

        VSS.getService<IExtensionDataService>(VSS.ServiceIds.ExtensionData).then((dataService: IExtensionDataService) => {
            dataService.getValue<StoredFieldReferences>("storedFields").then((storedFields:StoredFieldReferences) => {
                if (storedFields) {
                    console.log("Retrieved fields from storage");
                    this._selectedFields = storedFields;
                }
                else {
                    console.log("Failed to retrieve fields from storage, defaulting values")
					//Enter in your config referenceName for "tdField" and "gutField"
                    this._selectedFields = {
                        gvField: null,
                        ugField: null,
                        tdField: null,
                        gutField: null,
                        dateField: null,
                        multField: null
                    };
                }

                this.getSortedFieldsList().then((fieldList) => {
                    Controls.create(Combo, gvContainer, this.getComboOptions("Gravity", fieldList, this._selectedFields.gvField));
                    Controls.create(Combo, ugContainer, this.getComboOptions("Urgency", fieldList, this._selectedFields.ugField));
                    Controls.create(Combo, tdContainer, this.getComboOptions("Tendency", fieldList, this._selectedFields.tdField));
                    Controls.create(Combo, gutContainer, this.getComboOptions("GUT", fieldList, this._selectedFields.gutField));
                    Controls.create(Combo, dateContainer, this.getComboOptions("Date", fieldList, this._selectedFields.dateField));
                    Controls.create(Combo, multContainer, this.getComboOptions("Multiplier", fieldList, this._selectedFields.multField));
                    this.updateSaveButton();

                    VSS.notifyLoadSucceeded();
                });
            });
        });  
    }

    private save() {
        VSS.getService<IExtensionDataService>(VSS.ServiceIds.ExtensionData).then((dataService: IExtensionDataService) => {
            dataService.setValue<StoredFieldReferences>("storedFields", this._selectedFields).then((storedFields:StoredFieldReferences) => {
                console.log("Storing fields completed");
                this._changeMade = false;
                this.updateSaveButton();
            });
        });
    } 

    private updateSaveButton() {
        var buttonState = (this._selectedFields.gvField && this._selectedFields.ugField && this._selectedFields.tdField &&
                           this._selectedFields.gutField && this._selectedFields.dateField && this._selectedFields.multField) && this._changeMade
                            ? Menus.MenuItemState.None : Menus.MenuItemState.Disabled;

        // Update the disabled state
        this._menuBar.updateCommandStates([
            { id: "save", disabled: buttonState },
        ]);
    }
}