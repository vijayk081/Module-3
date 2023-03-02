import * as React from 'react';
import { IEmployeeListingProps } from './IEmployeeListingProps';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { IDropdownOption } from 'office-ui-fabric-react';
export interface IDetailsListBasicExampleItem {
    ID: number;
    Name1: string;
    DOB: any;
<<<<<<< Updated upstream
    Manager: any[];
    Experience: number;
    Department: string;
    DepartmentId: any;
    ManagerId: any;
=======
    Experience: number;
    Department: string;
    DepartmentId: any;
>>>>>>> Stashed changes
}
export interface IDetailsListBasicExampleState {
    items: IDetailsListBasicExampleItem[];
    selectionDetails: string;
    announcedMessage: any;
    FilterData: IDetailsListBasicExampleItem[];
    id: any;
    text: string;
    Name1: any;
    ItemId: any;
<<<<<<< Updated upstream
    Manager: [];
    EMail: any[];
    DepartmentId: any;
    ManagerId: any;
    DOB: any;
    plpuser: any[];
    Department: any;
    SelectedItem: any;
    SelectedManager: any;
    selectedusers: string[];
    Experience: any;
=======
    Name1up: any;
    DepartmentId: any;
    DOB: any;
    DOBup: any;
    Department: any;
    SelectedItem: any;
    Experience: any;
    Experienceup: any;
>>>>>>> Stashed changes
    hideDialog: boolean;
    projectlookupvalues: IDropdownOption[];
    hideDialogup: boolean;
    EditMode: boolean;
    SelectedItemup: any;
<<<<<<< Updated upstream
    selectedItems: any;
=======
>>>>>>> Stashed changes
}
export default class EmployeeListing extends React.Component<IEmployeeListingProps, IDetailsListBasicExampleState> {
    private _selection;
    private _columns;
    constructor(props: IEmployeeListingProps);
    render(): React.ReactElement<IEmployeeListingProps>;
    updatedialog: () => Promise<void>;
    onDropdownchange(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void;
    componentDidMount: () => void;
    test: () => Promise<void>;
<<<<<<< Updated upstream
    private _getPeoplePicker;
=======
>>>>>>> Stashed changes
    reset: () => Promise<void>;
    createItem: () => Promise<void>;
    private _onclickdelete;
    editModeItems: (Id: any) => void;
    UpdateItem: (ItemId: any) => Promise<void>;
    _getdeplookupfield: () => Promise<void>;
    private _getSelectionDetails;
    private _onItemInvoked;
    private _onFilter;
    private _onColumnClick;
}
//# sourceMappingURL=EmployeeListing.d.ts.map