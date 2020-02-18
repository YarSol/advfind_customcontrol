import * as React from 'react';
import { PrimaryButton } from 'office-ui-fabric-react';

import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
initializeIcons();

import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Stack, IStackStyles, IStackTokens, IStackItemStyles } from 'office-ui-fabric-react/lib/Stack';
import { Dropdown, IDropdownStyles, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { Label, ILabelStyles } from 'office-ui-fabric-react/lib/Label';

import { CommandBar, ICommandBarItemProps } from 'office-ui-fabric-react/lib/CommandBar';
import { IButtonProps } from 'office-ui-fabric-react/lib/Button';

import { DetailsList, DetailsListLayoutMode, Selection, IColumn, DetailsRow, IDetailsRowProps, IDetailsRowStyles } from 'office-ui-fabric-react/lib/DetailsList';
import { MarqueeSelection } from 'office-ui-fabric-react/lib/MarqueeSelection';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';

import { getTheme } from 'office-ui-fabric-react/lib/Styling';
import { IInputs } from './generated/ManifestTypes';

const theme = getTheme();

const DEVICE_HISTORY_STATUS_PENDING_INSTALL: number = 1;
const DEVICE_HISTORY_STATUS_INSTALLED: number = 964820000;

const DEVICE_STATUS_LOST: number = 964820008;
const DEVICE_STATUS_REPAIR: number = 964820009;
const DEVICE_STATUS_DAMAGED: number = 964820005;

const items: IDetailsListBasicExampleItem[] = [];

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


const labelItemStyles2: ILabelStyles = {
    root: {
        textAlign: "right",
        marginRight: 5
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

const itemAlignmentsStackTokens: IStackTokens = {
    childrenGap: 5,
    padding: 10
};

export interface IPCFContextProps {
    dropdownObjects: IDropdownOption[];
    updateField: () => void;
    context: ComponentFramework.Context<IInputs>;
}

export class ttttt extends React.Component {
    private _selection: Selection;
    private _allItems: IDetailsListBasicExampleItem[];
    private _columns: IColumn[];
    private statusOptions: IDropdownOption[];

    private _items: ICommandBarItemProps[] = [
        {
            key: 'selectItem',
            text: 'Select',
            cacheKey: 'myCacheKey', // changing this key will invalidate this item's cache
            iconProps: { iconName: 'Add' }
        }
    ];

    constructor(props: any) {
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



    public render(): JSX.Element {
        return (
            <Fabric>
                <div className="ms-Grid" dir="ltr">
                    <Stack styles={stackRowStyles}>
                        <div className="ms-Grid-row">
                            <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg4"><Stack styles={stackRowStyles} horizontal><Label styles={labelItemStyles} htmlFor="deviceId">Device ID</Label><TextField styles={textFieldItemStyles} id="deviceId" /></Stack></div>
                            <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg4"><Stack styles={stackRowStyles} horizontal><Label styles={labelItemStyles} htmlFor="deviceLocationId">Location ID</Label><TextField styles={textFieldItemStyles} id="deviceLocationId" /></Stack></div>
                        </div>
                    </Stack>

                    <Stack styles={stackRowStyles}>
                        <div className="ms-Grid-row">
                            <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg4"><Stack styles={stackRowStyles} horizontal><Label styles={labelItemStyles} htmlFor="deviceId">Device ID</Label><TextField styles={textFieldItemStyles} id="deviceId" /></Stack></div>
                            <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg4"><Stack styles={stackRowStyles} horizontal><Label styles={labelItemStyles} htmlFor="deviceLocationId">Location ID</Label><TextField styles={textFieldItemStyles} id="deviceLocationId" /></Stack></div>
                        </div>
                    </Stack>

                    <Stack styles={stackRowStyles}>
                        <div className="ms-Grid-row">
                            <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg4"><Stack styles={stackRowStyles} horizontal><Label styles={labelItemStyles} htmlFor="deviceId">Device ID</Label><TextField styles={textFieldItemStyles} id="deviceId" /></Stack></div>
                            <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg4"><Stack styles={stackRowStyles} horizontal><Label styles={labelItemStyles} htmlFor="deviceLocationId">Location ID</Label><TextField styles={textFieldItemStyles} id="deviceLocationId" /></Stack></div>
                        </div>
                    </Stack>

                    <Stack styles={stackRowStyles}>
                        <div className="ms-Grid-row">
                            <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg4"><Stack styles={stackRowStyles} horizontal><Label styles={labelItemStyles} htmlFor="deviceId">Device ID</Label><TextField styles={textFieldItemStyles} id="deviceId" /></Stack></div>
                            <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg4"><Stack styles={stackRowStyles} horizontal><Label styles={labelItemStyles} htmlFor="deviceLocationId">Location ID</Label><TextField styles={textFieldItemStyles} id="deviceLocationId" /></Stack></div>
                        </div>
                    </Stack>

                    <Stack styles={stackRowStyles}>
                        <div className="ms-Grid-row">
                            <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg4"><Stack styles={stackRowStyles} horizontal><Label styles={labelItemStyles} htmlFor="deviceId">Device ID</Label><TextField styles={textFieldItemStyles} id="deviceId" /></Stack></div>
                            <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg4"><Stack styles={stackRowStyles} horizontal><Label styles={labelItemStyles} htmlFor="deviceLocationId">Location ID</Label><TextField styles={textFieldItemStyles} id="deviceLocationId" /></Stack></div>
                        </div>
                    </Stack>

                    <Stack styles={stackRowStyles}>
                        <div className="ms-Grid-row">
                            <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg4"><Stack styles={stackRowStyles} horizontal><Label styles={labelItemStyles} htmlFor="deviceId">Device ID</Label><TextField styles={textFieldItemStyles} id="deviceId" /></Stack></div>
                            <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg4"><Stack styles={stackRowStyles} horizontal><Label styles={labelItemStyles} htmlFor="deviceLocationId">Location ID</Label><TextField styles={textFieldItemStyles} id="deviceLocationId" /></Stack></div>
                        </div>
                    </Stack>

                    <Stack styles={stackRowStyles}>
                        <div className="ms-Grid-row">
                            <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg4"><Stack styles={stackRowStyles} horizontal><Label styles={labelItemStyles} htmlFor="deviceId">Device ID</Label><TextField styles={textFieldItemStyles} id="deviceId" /></Stack></div>
                            <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg4"><Stack styles={stackRowStyles} horizontal><Label styles={labelItemStyles} htmlFor="deviceLocationId">Location ID</Label><TextField styles={textFieldItemStyles} id="deviceLocationId" /></Stack></div>
                        </div>
                    </Stack>
                </div>
            </Fabric >
        );

    };
};