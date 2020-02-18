import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { TextFieldBasicExample } from './PCFButton';
import { SearchEngineFormComponent } from './SearchEngineFormComponent';
import { IDropdownOption } from "office-ui-fabric-react";


export class SearchEngineForm implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    private theContainer: HTMLDivElement;
    private _context: ComponentFramework.Context<IInputs>;
    private _notifyOutputChanged: () => void;

    constructor() { }

    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement) {
        this.theContainer = container;
        this._context = context;
        this._notifyOutputChanged = notifyOutputChanged;

        let _this = this; 

        let req = new XMLHttpRequest();
        req.open("GET", (<any>this._context).page.getClientUrl() + "/api/data/v9.1/EntityDefinitions(LogicalName='msdyn_iotdevice')/Attributes/Microsoft.Dynamics.CRM.StatusAttributeMetadata?$select=LogicalName&$expand=OptionSet", true);
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
        req.onreadystatechange = function (t, c) {
            return function () {
                if (t.readyState === 4) {
                    req.onreadystatechange = null;
                    if (t.status === 200) {
                        let results = JSON.parse(t.response);
                        let options: IDropdownOption[] = [];

                        let statusCodeValues = results.value.filter((field: any) => field["LogicalName"] == "statuscode")[0]["OptionSet"]["Options"];

                        options.push({
                            key: "0",
                            text: "",
                            isSelected: false
                        });
                        for (let i = 0; i < statusCodeValues.length; i++) {
                            options.push({
                                key: statusCodeValues[i]["Value"],
                                text: statusCodeValues[i]["Label"]["UserLocalizedLabel"]["Label"],
                                isSelected: false
                            });
                        }

                        ReactDOM.render(
                            React.createElement(
                                SearchEngineFormComponent,
                                { 
                                    dropdownObjects: options,
                                    updateField: _this.updateField.bind(_this),
                                    context: _this._context
                                }
                            ),
                            c
                        );

                    }
                }
            }
        }(req, this.theContainer);
        return req.send();
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this._context = context;

    }

    public updateField() {
        this._notifyOutputChanged();
    }

    public getOutputs(): IOutputs {
        return {
            boundFieldValue: Date.now().toString()
        };
    }

    public destroy(): void {
        ReactDOM.unmountComponentAtNode(this.theContainer);
    }
}