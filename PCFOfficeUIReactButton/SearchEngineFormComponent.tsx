import * as React from 'react';
import { PrimaryButton } from 'office-ui-fabric-react';

import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
initializeIcons();

import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Stack, IStackStyles, IStackTokens, IStackItemStyles } from 'office-ui-fabric-react/lib/Stack';
import { Dropdown, IDropdownStyles, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { Label, ILabelStyles } from 'office-ui-fabric-react/lib/Label';

import { CommandBar, ICommandBarItemProps } from 'office-ui-fabric-react/lib/CommandBar';
import { IButtonProps, IButtonStyles } from 'office-ui-fabric-react/lib/Button';

import { DetailsList, DetailsListLayoutMode, Selection, IColumn, DetailsRow, IDetailsRowProps, IDetailsRowStyles } from 'office-ui-fabric-react/lib/DetailsList';
import { MarqueeSelection } from 'office-ui-fabric-react/lib/MarqueeSelection';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';

import { getTheme } from 'office-ui-fabric-react/lib/Styling';
import { IInputs } from './generated/ManifestTypes';

const theme = getTheme();

const DEVICE_HISTORY_STATUS_PENDING_INSTALL: number = 1;
const DEVICE_HISTORY_STATUS_INSTALLED: number = 964820000;

const DEVICE_HISTORY_STATUS_INACTIVE: number = 1;
const DEVICE_HISTORY_STATUS_ACTIVE: number = 0;

const DEVICE_STATUS_LOST: number = 964820008;
const DEVICE_STATUS_REPAIR: number = 964820009;
const DEVICE_STATUS_DAMAGED: number = 964820005;

export interface IDetailsListBasicExampleState {
    items: IDetailsListBasicExampleItem[];
    deviceId: string;
    deviceType: string;
    deviceModel: string;
    deviceStatus: string;
    deviceUnitAddress: string;
    deviceHeadendName: string;
    deviceNetworkCode: string;
    deviceLocationId: string;
    deviceAddress1: string;
    deviceAddress2: string;
    deviceCity: string;
    deviceState: string;
    deviceZip: string;
    deviceComment: string;

    displayGrid: boolean;
}

export interface IDetailsListBasicExampleItem {
    id: string;
    key: number;
    deviceId: string;
    deviceStatus: string;
    deviceModel: string;
    deviceType: string;
    headend: string;
    servicelocation: string;
    serviceaddress: string;
    unitaddress: string;
    statuscode: number;
    relatedHistoryStatus: number;
    pendingAddress: string;
}

const labelItemStyles: ILabelStyles = {
    root: {
        width: 120,
        textAlign: "right",
        marginRight: 5
    }
};

const stackRowStyles: IStackStyles = {
    root: {
        marginTop: 5,
        marginBottom: 5
    }
};

const textFieldItemStyles: ILabelStyles = {
    root: {
        width: 400
    }
};

const stackItemStyles: IStackItemStyles = {
    root: {
        padding: 5,
        width: 620
    }
};

const stackButtonsItemStyles: IStackItemStyles = {
    root: {
        padding: 5,
        marginTop: 10,
        marginLeft: 20
    }
};

const buttonsItemStyles: IButtonStyles = {
    root: {
        padding: 5,
        marginTop: 10,
        marginLeft: 20
    }
};

const itemAlignmentsStackTokens: IStackTokens = {
    childrenGap: 5,
    padding: 10
};

export interface IPCFContextProps {
    dropdownObjects: IDropdownOption[];
    updateField: () => void;
    context: ComponentFramework.Context<IInputs>;
}

export class SearchEngineFormComponent extends React.Component<IPCFContextProps, IDetailsListBasicExampleState> {
    private _selection: Selection;
    private _allItems: IDetailsListBasicExampleItem[];
    private _columns: IColumn[];
    private statusOptions: IDropdownOption[];

    private _items: ICommandBarItemProps[] = [
        {
            key: 'selectItem',
            text: 'Select',
            onClick: () => {
                this.performSelect();
            },
            cacheKey: 'myCacheKey', // changing this key will invalidate this item's cache
            iconProps: { iconName: 'Add' }
        }
    ];

    constructor(props: IPCFContextProps) {
        super(props);

        this.statusOptions = props.dropdownObjects;
        this._selection = new Selection();
        this._allItems = [];

        this._columns = [
            { key: 'column1', name: 'Device ID', fieldName: 'deviceId', minWidth: 100, maxWidth: 200, isResizable: true },
            { key: 'column2', name: 'Status', fieldName: 'deviceStatus', minWidth: 100, maxWidth: 200, isResizable: true },
            { key: 'column3', name: 'Model', fieldName: 'deviceModel', minWidth: 100, maxWidth: 200, isResizable: true },
            { key: 'column4', name: 'Type', fieldName: 'deviceType', minWidth: 100, maxWidth: 200, isResizable: true },
            { key: 'column5', name: 'Headend', fieldName: 'headend', minWidth: 100, maxWidth: 200, isResizable: true },
            { key: 'column6', name: 'Location', fieldName: 'servicelocation', minWidth: 100, maxWidth: 200, isResizable: true },
            { key: 'column7', name: 'Address', fieldName: 'serviceaddress', minWidth: 100, maxWidth: 200, isResizable: true },
            { key: 'column8', name: 'Unit Address', fieldName: 'unitaddress', minWidth: 100, maxWidth: 200, isResizable: true }
        ];

        this.state = {
            items: this._allItems,
            deviceId: "",
            deviceType: "",
            deviceModel: "",
            deviceStatus: "",
            deviceUnitAddress: "",
            deviceHeadendName: "",
            deviceNetworkCode: "",
            deviceLocationId: "",
            deviceAddress1: "",
            deviceAddress2: "",
            deviceCity: "",
            deviceState: "",
            deviceZip: "",
            deviceComment: "",

            displayGrid: false
        };
    }

    private performSelect() {        
        var selectedItems: IDetailsListBasicExampleItem[] = this._selection.getSelection() as IDetailsListBasicExampleItem[];

        if (selectedItems.length > 0) {
            var warningMessage = "";

            var j = 1;

            for (var i = 0; i < selectedItems.length; i++) {
                if (selectedItems[i].relatedHistoryStatus === DEVICE_HISTORY_STATUS_PENDING_INSTALL) {
                    warningMessage = warningMessage + j + ") " +
                        (selectedItems[i].deviceId ? selectedItems[i].deviceId : "") +
                        (selectedItems[i].deviceModel ? " - " + (selectedItems[i].deviceModel).trim() + " : " : " : ") +
                        "is currently in pending in " + selectedItems[i].pendingAddress + "\r\n";
                    j++;
                } else if (selectedItems[i].relatedHistoryStatus === DEVICE_HISTORY_STATUS_INSTALLED) {
                    warningMessage = warningMessage + j + ") " +
                        (selectedItems[i].deviceId ? selectedItems[i].deviceId : "") +
                        (selectedItems[i].deviceModel ? " - " + (selectedItems[i].deviceModel).trim() + " : " : " : ") +
                        "is currently installed in " + selectedItems[i].serviceaddress + "\r\n";
                    j++;
                } else if (selectedItems[i].statuscode === DEVICE_STATUS_LOST) {
                    warningMessage = warningMessage + j + ") " +
                        (selectedItems[i].deviceId ? selectedItems[i].deviceId : "") +
                        (selectedItems[i].deviceModel ? " - " + (selectedItems[i].deviceModel).trim() + " : " : " : ") + "lost" + "\r\n";
                    j++;
                } else if (selectedItems[i].statuscode === DEVICE_STATUS_REPAIR) {
                    warningMessage = warningMessage + j + ") " +
                        (selectedItems[i].deviceId ? selectedItems[i].deviceId : "") +
                        (selectedItems[i].deviceModel ? " - " + (selectedItems[i].deviceModel).trim() + " : " : " : ") + "under repair" + "\r\n";
                    j++;
                } else if (selectedItems[i].statuscode === DEVICE_STATUS_DAMAGED) {
                    warningMessage = warningMessage + j + ") " +
                        (selectedItems[i].deviceId ? selectedItems[i].deviceId : "") +
                        (selectedItems[i].deviceModel ? " - " + (selectedItems[i].deviceModel).trim() + " : " : " : ") + "damaged" + "\r\n";
                    j++;
                }
            }

            if (warningMessage != "") {
                // @ts-ignore
                this.props.context.navigation.openAlertDialog({ text: warningMessage, title: "Deselect devices which is not possible to add and try again :", confirmButtonLabel: "OK" }, { height: 450, width: 750 });
                return;
            }

            let currentRecordId = (this.props.context as any).page.entityId;

            let _this = this;

            this.getCurrentAccount(currentRecordId)
                .then(function (account: ComponentFramework.WebApi.Entity) {
                    let placeId = account["eti_ServiceLocation"]["eti_placeid"];
                    return _this.props.context.webAPI.retrieveRecord("eti_place", placeId, "?$select=eti_placeid&$expand=eti_ServiceAddress($select=eti_name)");
                })
                .then(function (place: ComponentFramework.WebApi.Entity) {
                    let data: any = {};
                    data["eti_Location@odata.bind"] = "/eti_places(" + place["eti_placeid"] + ")";
                    data["statuscode"] = DEVICE_HISTORY_STATUS_PENDING_INSTALL;
                    let today = new Date();
                    let date = today.getFullYear() + '-' + (today.getMonth() + 1) + '-' + today.getDate();
                    data["eti_requestdate"] = date;
                    data["eti_UserID@odata.bind"] = "/systemusers(" + _this.props.context.userSettings.userId.replace("{", "").replace("}", "") + ")";

                    for (var i = 0; i < selectedItems.length; i++) {
                        data["eti_Device@odata.bind"] = "/msdyn_iotdevices(" + selectedItems[i].id + ")";
                        _this.props.context.webAPI.createRecord("eti_devicelocationhistory", data);

                        selectedItems[i].relatedHistoryStatus = DEVICE_HISTORY_STATUS_PENDING_INSTALL;
                        selectedItems[i].pendingAddress = place["eti_ServiceAddress"]["eti_name"];
                    }

                    _this.setState({
                        items: [..._this.state.items]
                    });

                }).catch(console.log);
        }
    };

    private getCurrentAccount(accountId: string): Promise<any> {
        return this.props.context.webAPI.retrieveRecord("account", accountId, "?$select=accountid&$expand=eti_ServiceLocation($select=eti_name)");
    }

    private performSearch(e: any) {
        let deviceId = this.state.deviceId;
        let deviceType = this.state.deviceType;
        let deviceModel = this.state.deviceModel;
        let deviceStatus = this.state.deviceStatus;
        let deviceUnitAddress = this.state.deviceUnitAddress;
        let deviceHeadendName = this.state.deviceHeadendName;
        let deviceNetworkCode = this.state.deviceNetworkCode;
        let deviceLocationId = this.state.deviceLocationId;
        let deviceAddress1 = this.state.deviceAddress1;
        let deviceAddress2 = this.state.deviceAddress2;
        let deviceCity = this.state.deviceCity;
        let deviceState = this.state.deviceState;
        let deviceZip = this.state.deviceZip;
        let deviceComment = this.state.deviceComment;

        let resultFetch = "";
        let device_conditions = [];
        let linkEntity_serviceLocation_conditions = [];
        let linkEntity_serviceLocation_linkEntity_ServiceAddress_conditions = [];

        // field-filters for Device entity
        if (deviceId) device_conditions.push(`<condition attribute="msdyn_deviceid" operator="like" value="%` + deviceId + `%" />`);
        if (deviceType) device_conditions.push(`<condition attribute="msdyn_categoryname" operator="like" value="%` + deviceType + `%" />`);
        if (deviceModel) device_conditions.push(`<condition attribute="eti_devicemodelname" operator="like" value="%` + deviceModel + `%" />`);
        if (deviceStatus && deviceStatus != "0") device_conditions.push(`<condition attribute="statuscode" operator="eq" value="` + deviceStatus + `" />`);
        if (deviceUnitAddress) device_conditions.push(`<condition attribute="eti_unitaddress" operator="like" value="%` + deviceUnitAddress + `%" />`);

        // field-filters for related Place entity
        if (deviceHeadendName) linkEntity_serviceLocation_conditions.push(`<condition attribute="eti_headendname" operator="like" value="%` + deviceHeadendName + `%" />`);
        if (deviceLocationId) linkEntity_serviceLocation_conditions.push(`<condition attribute="eti_name" operator="like" value="%` + deviceLocationId + `%" />`);
        if (deviceAddress1) linkEntity_serviceLocation_conditions.push(`<condition attribute="eti_serviceaddressname" operator="like" value="%` + deviceAddress1 + `%" />`);
        if (deviceAddress2) linkEntity_serviceLocation_conditions.push(`<condition attribute="eti_serviceaddressname" operator="like" value="%` + deviceAddress2 + `%" />`);

        // field-filters for Service Address related to Place entity
        if (deviceCity) linkEntity_serviceLocation_linkEntity_ServiceAddress_conditions.push(`<condition attribute="eti_city" operator="like" value="%` + deviceCity + `%" />`);
        if (deviceState) linkEntity_serviceLocation_linkEntity_ServiceAddress_conditions.push(`<condition attribute="eti_stateorprovince" operator="like" value="%` + deviceState + `%" />`);
        if (deviceZip) linkEntity_serviceLocation_linkEntity_ServiceAddress_conditions.push(`<condition attribute="eti_postcode" operator="like" value="%` + deviceZip + `%" />`);

        let device_serviceLocation_serviceAddress_fetchXml = "";
        if (linkEntity_serviceLocation_linkEntity_ServiceAddress_conditions.length > 0)
            device_serviceLocation_serviceAddress_fetchXml = `<link-entity name="eti_geographicaddress" from="eti_geographicaddressid" to="eti_serviceaddress" link-type="inner" ><filter type="and" >` + linkEntity_serviceLocation_linkEntity_ServiceAddress_conditions.join(" ") + `</filter></link-entity>`;

        let device_serviceLocation_fetchXml = "";
        if (linkEntity_serviceLocation_conditions.length > 0 || linkEntity_serviceLocation_linkEntity_ServiceAddress_conditions.length > 0) {
            device_serviceLocation_fetchXml = `<link-entity name="eti_place" from="eti_placeid" to="eti_place" link-type="inner" ><filter type="and" >` + linkEntity_serviceLocation_conditions.join(" ") + `</filter>` + device_serviceLocation_serviceAddress_fetchXml + `</link-entity>`;
        }

        resultFetch = `<fetch page="1" mapping='logical'><entity name="msdyn_iotdevice" ><attribute name="msdyn_deviceid" /><attribute name="eti_place" /><attribute name="msdyn_iotdeviceid" /><attribute name="statuscode" /><attribute name="eti_unitaddress" /><attribute name="eti_devicemodel" /><attribute name="msdyn_category" /><filter type="and" >`
            + (device_conditions.length > 0 ? device_conditions.join(" ") : `<condition attribute="statecode" operator="eq" value="0" />`)
            + `</filter><link-entity name="eti_devicelocationhistory" from="eti_device" to="msdyn_iotdeviceid" link-type="outer" alias="lochist" ><attribute name="statuscode" alias="device_history_status" /><filter type="and" ><condition attribute="statecode" operator="eq" value="` + DEVICE_HISTORY_STATUS_ACTIVE + `" /><filter type="or" ><condition attribute="statuscode" operator="eq" value="` + DEVICE_HISTORY_STATUS_INSTALLED + `" /><condition attribute="statuscode" operator="eq" value="` + DEVICE_HISTORY_STATUS_PENDING_INSTALL + `" /></filter></filter><link-entity name="eti_place" from="eti_placeid" to="eti_location" link-type="outer" ><attribute name="eti_serviceaddress" alias="pending_address" /></link-entity></link-entity><link-entity name="eti_place" from="eti_placeid" to="eti_place" link-type="outer" ><attribute name="eti_headend" /><attribute name="eti_serviceaddress" /></link-entity>`
            + (device_serviceLocation_fetchXml == "" ? `` : device_serviceLocation_fetchXml)
            + `</entity></fetch>`;

        this.getDevicesFromFetch(resultFetch, 1, []).then(this.transformDataToNeededFormat.bind(null, this)).catch(console.log);
    };

    private transformDataToNeededFormat(_this: any, data: any[]) {
        let result: IDetailsListBasicExampleItem[] = [];

        for (var i = 0; i < data.length; i++) {
            let t: any = {};
            t["id"] = data[i]["msdyn_iotdeviceid"];
            t["key"] = i;
            t["deviceId"] = data[i]["msdyn_deviceid"];
            t["deviceStatus"] = data[i]["statuscode@OData.Community.Display.V1.FormattedValue"];
            t["deviceModel"] = data[i]["_eti_devicemodel_value@OData.Community.Display.V1.FormattedValue"];
            t["deviceType"] = data[i]["_msdyn_category_value@OData.Community.Display.V1.FormattedValue"];
            t["headend"] = data[i]["eti_place3.eti_headend@OData.Community.Display.V1.FormattedValue"];
            t["servicelocation"] = data[i]["_eti_place_value@OData.Community.Display.V1.FormattedValue"];
            t["serviceaddress"] = data[i]["eti_place3.eti_serviceaddress@OData.Community.Display.V1.FormattedValue"] || "";
            t["unitaddress"] = data[i]["eti_unitaddress"] || "";
            t["statuscode"] = data[i]["statuscode"] as number;
            t["relatedHistoryStatus"] = data[i]["device_history_status"] ? (data[i]["device_history_status"] as number) : 0;
            t["pendingAddress"] = data[i]["pending_address@OData.Community.Display.V1.FormattedValue"] || "";

            result.push(t);
        }

        _this.setState({
            items: result,
            displayGrid: true
        });
    }

    private getDevicesFromFetch(fetch: string, pageNumber: number, resultArray: any[]): Promise<any[]> {
        if (pageNumber > 1)
            fetch = fetch.replace('page="' + (pageNumber - 1) + '"', 'page="' + pageNumber + '"');

        // @ts-ignore    
        let requestUrl = this.props.context.page.getClientUrl() + `/api/data/v9.1/msdyn_iotdevices?fetchXml=` + encodeURIComponent(fetch);
        let _this = this;

        return this.makeGetRequest(requestUrl).then(function (result: any[]) {
            if (result.length < 5000) {
                resultArray = resultArray.concat(result);
                return resultArray;
            } else {
                resultArray = resultArray.concat(result);
                pageNumber++;
                return _this.getDevicesFromFetch(fetch, pageNumber, resultArray);
            }
        });
    }

    private makeGetRequest(url: string): Promise<any[]> {
        var req = new XMLHttpRequest();

        return new Promise(function (resolve, reject) {
            req.onreadystatechange = function () {
                if (req.readyState !== 4) return;

                if (req.status >= 200 && req.status < 300) {
                    let results = JSON.parse(this.response);
                    resolve(results.value as any[]);
                } else {
                    reject({
                        status: req.status,
                        statusText: req.statusText
                    });
                }
            };

            req.open("GET", url, true);
            req.setRequestHeader("OData-MaxVersion", "4.0");
            req.setRequestHeader("OData-Version", "4.0");
            req.setRequestHeader("Accept", "application/json");
            req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
            req.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
            req.send();
        });
    }

    private resetFilters(e: any) {
        this.setState({
            deviceId: "",
            deviceType: "",
            deviceModel: "",
            deviceStatus: "",
            deviceUnitAddress: "",
            deviceHeadendName: "",
            deviceNetworkCode: "",
            deviceLocationId: "",
            deviceAddress1: "",
            deviceAddress2: "",
            deviceCity: "",
            deviceState: "",
            deviceZip: "",
            deviceComment: ""
        });
    };

    private closeAdvFind(e: any) {
        this.resetFilters(null);

        this.setState({
            items: []
        });

        this.props.updateField();
    };

    private _onItemInvoked = (item: IDetailsListBasicExampleItem): void => {
        alert(`Item invoked: ${item.deviceId}`);
    };

    private setDeviceId(e: any) {
        this.setState({
            deviceId: e.target.value
        });
    }

    private setDeviceType(e: any) {
        this.setState({
            deviceType: e.target.value
        });
    }

    private setDeviceModel(e: any) {
        this.setState({
            deviceModel: e.target.value
        });
    }

    private setDeviceStatus(e: React.FormEvent<HTMLDivElement> | any, option: any = {}, index?: number) {
        this.setState({
            deviceStatus: option.key
        });
    }

    private setDeviceUnitAddress(e: any) {
        this.setState({
            deviceUnitAddress: e.target.value
        });
    }

    private setDeviceHeadendName(e: any) {
        this.setState({
            deviceHeadendName: e.target.value
        });
    }

    private setDeviceNetworkCode(e: any) {
        this.setState({
            deviceNetworkCode: e.target.value
        });
    }

    private setDeviceLocationId(e: any) {
        this.setState({
            deviceLocationId: e.target.value
        });
    }

    private setDeviceAddress1(e: any) {
        this.setState({
            deviceAddress1: e.target.value
        });
    }

    private setDeviceAddress2(e: any) {
        this.setState({
            deviceAddress2: e.target.value
        });
    }

    private setDeviceCity(e: any) {
        this.setState({
            deviceCity: e.target.value
        });
    }

    private setDeviceState(e: any) {
        this.setState({
            deviceState: e.target.value
        });
    }

    private setDeviceZip(e: any) {
        this.setState({
            deviceZip: e.target.value
        });
    }

    private setDeviceComment(e: any) {
        this.setState({
            deviceComment: e.target.value
        });
    }

    private _onRenderRow = (props?: IDetailsRowProps): JSX.Element => {
        if (!props) {
            return <DetailsRow item={undefined} itemIndex={-1} />;
        }
        const customStyles: Partial<IDetailsRowStyles> = {};
        if (props.item["relatedHistoryStatus"] === DEVICE_HISTORY_STATUS_PENDING_INSTALL) {
            customStyles.root = { backgroundColor: "#03A2F9", selectors: { ':hover': { background: '#03A2F9' } } };
        } else if (props.item["relatedHistoryStatus"] === DEVICE_HISTORY_STATUS_INSTALLED) {
            customStyles.root = { backgroundColor: "#1FF903", selectors: { ':hover': { background: '#1FF903' } } };
        } else if (props.item["statuscode"] === DEVICE_STATUS_LOST || props.item["statuscode"] === DEVICE_STATUS_REPAIR || props.item["statuscode"] === DEVICE_STATUS_DAMAGED) {
            customStyles.root = { backgroundColor: "#F91506", selectors: { ':hover': { background: '#F91506' } } };
        }

        return <DetailsRow {...props} styles={customStyles} />;
    };

    public render(): JSX.Element {
        const { items, deviceId, deviceType, deviceModel, deviceStatus, deviceUnitAddress, deviceHeadendName, deviceNetworkCode, deviceLocationId, deviceAddress1, deviceAddress2, deviceCity, deviceState, deviceZip, deviceComment } = this.state;

        return (
            <Fabric>
                <div className="ms-Grid" dir="ltr">
                    <Stack styles={stackRowStyles}>
                        <div className="ms-Grid-row">
                            <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg4"><Stack styles={stackRowStyles} horizontal><Label styles={labelItemStyles} htmlFor="deviceId">Device ID</Label><TextField value={deviceId} onChange={e => this.setDeviceId(e)} styles={textFieldItemStyles} id="deviceId" /></Stack></div>
                            <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg4"><Stack styles={stackRowStyles} horizontal><Label styles={labelItemStyles} htmlFor="deviceLocationId">Location ID</Label><TextField value={deviceLocationId} onChange={e => this.setDeviceLocationId(e)} styles={textFieldItemStyles} id="deviceLocationId" /></Stack></div>
                        </div>
                    </Stack>
                    <Stack styles={stackRowStyles}>
                        <div className="ms-Grid-row">
                            <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg4"><Stack styles={stackRowStyles} horizontal><Label styles={labelItemStyles} htmlFor="deviceType">Type</Label><TextField value={deviceType} onChange={e => this.setDeviceType(e)} styles={textFieldItemStyles} id="deviceType" /></Stack></div>
                            <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg4"><Stack styles={stackRowStyles} horizontal><Label styles={labelItemStyles} htmlFor="deviceAddress1">Address 1</Label><TextField value={deviceAddress1} onChange={e => this.setDeviceAddress1(e)} styles={textFieldItemStyles} id="deviceAddress1" /></Stack></div>
                        </div>
                    </Stack>
                    <Stack styles={stackRowStyles}>
                        <div className="ms-Grid-row">
                            <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg4"><Stack styles={stackRowStyles} horizontal><Label styles={labelItemStyles} htmlFor="deviceModel">Model</Label><TextField value={deviceModel} onChange={e => this.setDeviceModel(e)} styles={textFieldItemStyles} id="deviceModel" /></Stack></div>
                            <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg4"><Stack styles={stackRowStyles} horizontal><Label styles={labelItemStyles} htmlFor="deviceAddress2">Address 2</Label><TextField value={deviceAddress2} onChange={e => this.setDeviceAddress2(e)} styles={textFieldItemStyles} id="deviceAddress2" /></Stack></div>
                        </div>
                    </Stack>
                    <Stack styles={stackRowStyles}>
                        <div className="ms-Grid-row">
                            <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg4"><Stack styles={stackRowStyles} horizontal><Label styles={labelItemStyles} htmlFor="deviceStatus">Status</Label><Dropdown selectedKey={deviceStatus} onChange={(e, option, index) => this.setDeviceStatus(e, option, index)} options={this.statusOptions} styles={textFieldItemStyles} id="deviceStatus" /></Stack></div>
                            <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg4"><Stack styles={stackRowStyles} horizontal><Label styles={labelItemStyles} htmlFor="deviceCity">City</Label><TextField value={deviceCity} onChange={e => this.setDeviceCity(e)} styles={textFieldItemStyles} id="deviceCity" /></Stack></div>
                        </div>
                    </Stack>
                    <Stack styles={stackRowStyles}>
                        <div className="ms-Grid-row">
                            <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg4"><Stack styles={stackRowStyles} horizontal><Label styles={labelItemStyles} htmlFor="deviceUnitAddress">Unit Address</Label><TextField value={deviceUnitAddress} onChange={e => this.setDeviceUnitAddress(e)} styles={textFieldItemStyles} id="deviceUnitAddress" /></Stack></div>
                            <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg4"><Stack styles={stackRowStyles} horizontal><Label styles={labelItemStyles} htmlFor="deviceState">State</Label><TextField value={deviceState} onChange={e => this.setDeviceState(e)} styles={textFieldItemStyles} id="deviceState" /></Stack></div>
                        </div>
                    </Stack>
                    <Stack styles={stackRowStyles}>
                        <div className="ms-Grid-row">
                            <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg4"><Stack styles={stackRowStyles} horizontal><Label styles={labelItemStyles} htmlFor="deviceHeadendName">Headend Name</Label><TextField value={deviceHeadendName} onChange={e => this.setDeviceHeadendName(e)} styles={textFieldItemStyles} id="deviceHeadendName" /></Stack></div>
                            <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg4"><Stack styles={stackRowStyles} horizontal><Label styles={labelItemStyles} htmlFor="deviceZip">ZIP</Label><TextField value={deviceZip} onChange={e => this.setDeviceZip(e)} styles={textFieldItemStyles} id="deviceZip" /></Stack></div>
                        </div>
                    </Stack>
                    <Stack styles={stackRowStyles}>
                        <div className="ms-Grid-row">
                            <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg4"><Stack styles={stackRowStyles} horizontal><Label styles={labelItemStyles} htmlFor="deviceNetworkCode">Network Code</Label><TextField value={deviceNetworkCode} onChange={e => this.setDeviceNetworkCode(e)} styles={textFieldItemStyles} id="deviceNetworkCode" /></Stack></div>
                            <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg4"><Stack styles={stackRowStyles} horizontal><Label styles={labelItemStyles} htmlFor="deviceComment">Comment</Label><TextField value={deviceComment} onChange={e => this.setDeviceComment(e)} styles={textFieldItemStyles} id="deviceComment" /></Stack></div>
                        </div>
                    </Stack>
                </div>

                <div className="ms-Grid" dir="ltr">
                    <Stack styles={stackRowStyles}>
                        <div className="ms-Grid-row">
                            <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg4"><PrimaryButton styles={buttonsItemStyles} onClick={e => this.performSearch(e)}>Perform Search</PrimaryButton></div>
                            <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg4"><PrimaryButton styles={buttonsItemStyles} onClick={e => this.resetFilters(e)}>Reset Filters</PrimaryButton></div>
                            <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg4"><PrimaryButton styles={buttonsItemStyles} onClick={e => this.closeAdvFind(e)}>Close Adv Find</PrimaryButton></div>
                        </div>
                    </Stack>
                </div>


                {
                    this.state.displayGrid
                        ? <Stack>
                            <div>
                                <CommandBar
                                    items={this._items}
                                    /*overflowItems={_overflowItems}*/
                                    /*overflowButtonProps={overflowProps}*/
                                    /*farItems={_farItems}*/
                                    ariaLabel="Use left and right arrow keys to navigate between commands"
                                />
                            </div>
                            <MarqueeSelection selection={this._selection}>
                                <DetailsList
                                    items={items}
                                    columns={this._columns}
                                    setKey="set"
                                    layoutMode={DetailsListLayoutMode.justified}
                                    selection={this._selection}
                                    selectionPreservedOnEmptyClick={true}
                                    ariaLabelForSelectionColumn="Toggle selection"
                                    ariaLabelForSelectAllCheckbox="Toggle selection for all items"
                                    checkButtonAriaLabel="Row checkbox"
                                    onItemInvoked={this._onItemInvoked}
                                    onRenderRow={this._onRenderRow}
                                />
                            </MarqueeSelection>
                        </Stack>
                        : null
                }
            </Fabric >
        );

    };
};