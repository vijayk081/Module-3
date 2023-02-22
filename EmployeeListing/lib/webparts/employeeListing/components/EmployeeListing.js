var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
import * as React from 'react';
// import { escape } from '@microsoft/sp-lodash-subset';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { DetailsList, DetailsListLayoutMode, Selection } from 'office-ui-fabric-react/lib/DetailsList';
import { MarqueeSelection } from 'office-ui-fabric-react/lib/MarqueeSelection';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { mergeStyles } from 'office-ui-fabric-react/lib/Styling';
import { Announced } from 'office-ui-fabric-react/lib/Announced';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { Dialog, DialogFooter } from '@fluentui/react/lib/Dialog';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Icon } from '@fluentui/react';
import { DatePicker, Dropdown } from 'office-ui-fabric-react';
import * as moment from 'moment';
// DatePicker,
// import { Dropdown, DropdownMenuItemType, IDropdownOption, IDropdownProps } from '@fluentui/react/lib/Dropdown';
// import { Fabric } from 'office-ui-fabric-react';
var exampleChildClass = mergeStyles({
    display: 'block',
    marginBottom: '10px',
});
var textFieldStyles = { root: { maxWidth: '300px' } };
var EmployeeListing = /** @class */ (function (_super) {
    __extends(EmployeeListing, _super);
    function EmployeeListing(props) {
        var _this = _super.call(this, props) || this;
        _this.updatedialog = function () { return __awaiter(_this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                this.setState({
                    hideDialogup: false,
                });
                return [2 /*return*/];
            });
        }); };
        _this.componentDidMount = function () {
            _this.test();
            _this._getdeplookupfield();
        };
        _this.test = function () { return __awaiter(_this, void 0, void 0, function () {
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, sp.web.lists.getByTitle("Employees").items.select("ID,Name1,DOB,Experience,Department/ID,Department/DepartmentName,Manager/ID,Manager/EMail").expand("Department,Manager").get().then(function (items) {
                            var AllData = []; // Manager: any;
                            items.map(function (data) {
                                var dataDOB = moment(data.DOB).format('MM/DD/YYYY');
                                var Allusers = [];
                                data.Manager.map(function (val) {
                                    Allusers.push(val.EMail);
                                });
                                AllData.push({
                                    ID: data.ID,
                                    Name1: data.Name1,
                                    DOB: dataDOB,
                                    Department: data.Department.DepartmentName,
                                    Experience: data.Experience,
                                    Manager: Allusers,
                                    DepartmentId: data.Department.ID,
                                    ManagerId: data.Manager.ID,
                                    Action: (React.createElement(React.Fragment, null,
                                        React.createElement(Icon, { iconName: 'delete', onClick: function () { _this._onclickdelete(data.ID); }, style: { marginRight: 30 } }),
                                        React.createElement(Icon, { iconName: 'EditSolid12', onClick: function () { _this.editModeItems(data.ID); } })))
                                });
                            });
                            _this.setState({
                                items: AllData,
                                selectionDetails: _this._getSelectionDetails(),
                                FilterData: AllData,
                            });
                        }).catch(function (e) {
                            console.log(e);
                        })];
                    case 1:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        }); };
        // People picker
        _this._getPeoplePicker = function (plpUser) {
            var AllManager = [];
            plpUser.map(function (val) {
                AllManager.push(val.id);
            });
            _this.setState({ plpuser: AllManager });
        };
        //for combination of dialog
        _this.reset = function () { return __awaiter(_this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                this.setState({ Name1: '', DOB: null, Experience: '', SelectedItemup: 0 });
                return [2 /*return*/];
            });
        }); };
        // create item
        _this.createItem = function () { return __awaiter(_this, void 0, void 0, function () {
            var _this = this;
            return __generator(this, function (_a) {
                sp.web.lists.getByTitle("Employees")
                    .items.add({
                    Name1: this.state.Name1,
                    DOB: this.state.DOB,
                    Experience: this.state.Experience,
                    DepartmentId: this.state.SelectedItem,
                    ManagerId: { results: this.state.plpuser }
                }).then(function () {
                    _this.setState({ hideDialog: true });
                    _this.test();
                }).catch(function (err) {
                    console.log(err);
                });
                return [2 /*return*/];
            });
        }); };
        _this._onclickdelete = function (ID) { return __awaiter(_this, void 0, void 0, function () {
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        console.log(ID);
                        return [4 /*yield*/, sp.web.lists.getByTitle("Employees").items.getById(ID).delete().then(function (data) {
                                console.log(data);
                                _this.test();
                            }).catch(function (err) {
                                console.log(err);
                            })];
                    case 1:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        }); };
        _this.editModeItems = function (Id) {
            var editItem = _this.state.items.filter(function (x) { return x.ID == Id; })[0];
            _this.setState({
                Name1: editItem.Name1,
                DOB: editItem.DOB,
                Experience: editItem.Experience,
                Manager: editItem.ManagerId,
                ItemId: editItem.ID,
                Department: editItem.Department,
                SelectedItemup: editItem.DepartmentId,
                selectedusers: editItem.Manager,
                // ManagerId:editItem.ManagerId,
                hideDialog: false,
                EditMode: true
            });
        };
        _this.UpdateItem = function (ItemId) { return __awaiter(_this, void 0, void 0, function () {
            var _this = this;
            return __generator(this, function (_a) {
                sp.web.lists.getByTitle("Employees").items.getById(ItemId)
                    .update({
                    Name1: this.state.Name1,
                    DOB: this.state.DOB,
                    Experience: this.state.Experience,
                    DepartmentId: this.state.SelectedItemup,
                    ManagerId: { results: this.state.plpuser }
                }).then(function () {
                    // location.reload();
                    _this.setState({ hideDialogup: true, hideDialog: true });
                    _this.reset(),
                        _this.test();
                })
                    .catch(function (err) {
                    console.log(err);
                });
                return [2 /*return*/];
            });
        }); };
        _this._getdeplookupfield = function () { return __awaiter(_this, void 0, void 0, function () {
            var allItems, dropdowndep;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, sp.web.lists.getByTitle("Department").items.getAll()];
                    case 1:
                        allItems = _a.sent();
                        dropdowndep = [];
                        allItems.forEach(function (Department) {
                            dropdowndep.push({ key: Department.ID, text: Department.DepartmentName });
                        });
                        this.setState({
                            projectlookupvalues: dropdowndep
                        });
                        return [2 /*return*/];
                }
            });
        }); };
        _this._onItemInvoked = function (item) {
            alert("Item invoked: ".concat(item.Name1));
        };
        _this._onFilter = function (ev, text) {
            _this.setState({
                items: text ? _this.state.FilterData.filter(function (i) { return i.Name1.toLowerCase().indexOf(text) > -1; }) : _this.state.FilterData,
            });
        };
        _this._onColumnClick = function (ev, column) {
            var items = _this.state.items;
            var newColumns = _this._columns.slice();
            var currColumn = newColumns.filter(function (currCol) { return column.key === currCol.key; })[0];
            newColumns.forEach(function (newCol) {
                if (newCol === currColumn) {
                    currColumn.isSortedDescending = !currColumn.isSortedDescending;
                    currColumn.isSorted = true;
                    _this.setState({
                        announcedMessage: "".concat(currColumn.name, " is sorted ").concat(currColumn.isSortedDescending ? 'descending' : 'ascending'),
                    });
                }
                else {
                    newCol.isSorted = false;
                    newCol.isSortedDescending = true;
                }
            });
            var newItems = _copyAndSort(items, currColumn.fieldName, currColumn.isSortedDescending);
            _this.setState({
                items: newItems,
            });
        };
        sp.setup({
            spfxContext: _this.props.spfxcontext
        });
        _this.state = {
            items: [],
            selectionDetails: '',
            announcedMessage: undefined,
            FilterData: [],
            id: '',
            plpuser: [],
            ItemId: [],
            DepartmentId: '',
            ManagerId: '',
            SelectedItem: 0,
            SelectedManager: [],
            selectedusers: [],
            text: 'string',
            Manager: [],
            Name1: '',
            DOB: null,
            Department: '',
            EMail: [],
            projectlookupvalues: [],
            Experience: '',
            hideDialog: true,
            hideDialogup: true,
            EditMode: false,
            SelectedItemup: '',
            selectedItems: [],
        };
        _this._getdeplookupfield();
        _this._selection = new Selection({
            onSelectionChanged: function () { return _this.setState({ selectionDetails: _this._getSelectionDetails() }); },
        });
        _this._columns = [
            { key: 'column1', name: 'Action', fieldName: 'Action', minWidth: 100, maxWidth: 200, isResizable: true },
            { key: 'column2', name: 'ID', fieldName: 'ID', minWidth: 100, maxWidth: 200, isResizable: true },
            { key: 'column3', name: 'Name', fieldName: 'Name1', minWidth: 100, maxWidth: 200, isResizable: true, onColumnClick: _this._onColumnClick },
            { key: 'column4', name: 'DOB', fieldName: 'DOB', minWidth: 100, maxWidth: 200, isResizable: true },
            { key: 'column5', name: 'Experience', fieldName: 'Experience', minWidth: 100, maxWidth: 200, isResizable: true },
            { key: 'column6', name: 'Department', fieldName: 'Department', minWidth: 100, maxWidth: 200, isResizable: true },
            { key: 'column7', name: 'Manager', fieldName: 'Manager', minWidth: 100, maxWidth: 200, isResizable: true },
        ];
        return _this;
    }
    EmployeeListing.prototype.render = function () {
        var _this = this;
        var _a = this.state, items = _a.items, selectionDetails = _a.selectionDetails;
        return (React.createElement("div", null,
            React.createElement(Fabric, null,
                React.createElement("div", { className: exampleChildClass }, selectionDetails),
                React.createElement(Announced, { message: selectionDetails }),
                React.createElement(TextField, { className: exampleChildClass, label: "Filter by name:", onChange: this._onFilter, styles: textFieldStyles }),
                React.createElement(Announced, { message: "Number of items after filter applied: ".concat(items.length, ".") }),
                React.createElement(PrimaryButton, { text: "Add Employee", onClick: function () { _this.setState({ hideDialog: false, EditMode: false }); } }),
                React.createElement(MarqueeSelection, { selection: this._selection },
                    React.createElement(DetailsList, { items: this.state.items, columns: this._columns, setKey: "set", layoutMode: DetailsListLayoutMode.justified, selection: this._selection, selectionPreservedOnEmptyClick: true, ariaLabelForSelectionColumn: "Toggle selection", ariaLabelForSelectAllCheckbox: "Toggle selection for all items", checkButtonAriaLabel: "select row", onItemInvoked: this._onItemInvoked }))),
            React.createElement("div", { className: 'adddialog' },
                React.createElement(Dialog, { hidden: this.state.hideDialog },
                    React.createElement("h1", null, this.state.EditMode ? "Update" : "Add Employee"),
                    React.createElement("div", { className: 'table' },
                        React.createElement("table", null,
                            React.createElement("tr", { className: 'Name1' },
                                React.createElement("td", null, "Name :"),
                                React.createElement("td", null,
                                    React.createElement("input", { type: "text", id: "Name1", value: this.state.Name1 ? this.state.Name1 : '', onChange: function (e) { _this.setState({ Name1: e.target.value }); } }))),
                            React.createElement("tr", { className: 'DOB' },
                                React.createElement("td", null, "DOB :"),
                                React.createElement("td", null,
                                    React.createElement(DatePicker, { id: "DOB", value: this.state.DOB ? new Date(this.state.DOB) : null, onSelectDate: function (selectedDate) {
                                            _this.setState({ DOB: selectedDate });
                                        } }))),
                            React.createElement("tr", { className: 'Experience' },
                                React.createElement("td", null, "Experience :"),
                                React.createElement("td", null,
                                    React.createElement("input", { type: "number", id: "Experience", value: this.state.Experience ? this.state.Experience : '', onChange: function (e) { _this.setState({ Experience: e.target.value }); } }))),
                            React.createElement("tr", { className: 'Department' },
                                React.createElement("td", null, "Department :"),
                                React.createElement("td", null,
                                    React.createElement(Dropdown, { placeholder: "Select a Department", options: this.state.projectlookupvalues, defaultSelectedKey: this.state.SelectedItemup, onChange: function (e, val) { _this.onDropdownchange(e, val); } }))),
                            React.createElement("tr", { className: 'Manager' },
                                React.createElement("td", null, "Manager:"),
                                React.createElement("td", null,
                                    React.createElement(PeoplePicker, { context: this.props.spfxcontext, personSelectionLimit: 3, showtooltip: true, required: true, disabled: false, defaultSelectedUsers: this.state.selectedusers, onChange: this._getPeoplePicker, showHiddenInUI: false, ensureUser: true, principalTypes: [PrincipalType.User], resolveDelay: 1000 }))))),
                    React.createElement(DialogFooter, null,
                        React.createElement(PrimaryButton, { onClick: function () { _this.state.EditMode ? _this.updatedialog() : _this.createItem(); } },
                            this.state.EditMode ? "Update" : "Save",
                            " "),
                        React.createElement(DefaultButton, { text: "Cancel", onClick: function () { _this.setState({ hideDialog: true }), _this.reset(); } }))),
                React.createElement(Dialog, { hidden: this.state.hideDialogup },
                    React.createElement("text", null, "Are you sure want to update details?"),
                    React.createElement(DialogFooter, null,
                        React.createElement(PrimaryButton, { text: "yes", onClick: function () { return _this.UpdateItem(_this.state.ItemId); } }),
                        React.createElement(DefaultButton, { text: "No", onClick: function () { _this.setState({ hideDialogup: true }), _this.reset(); } }))))));
    };
    EmployeeListing.prototype.onDropdownchange = function (event, item) {
        console.log();
        this.setState({ SelectedItem: item.key, SelectedItemup: item.key });
    };
    EmployeeListing.prototype._getSelectionDetails = function () {
        var selectionCount = this._selection.getSelectedCount();
        switch (selectionCount) {
            case 0:
                return 'No items selected';
            case 1:
                return '1 item selected: ' + this._selection.getSelection()[0].Name1;
            default:
                return "".concat(selectionCount, " items selected");
        }
    };
    return EmployeeListing;
}(React.Component));
export default EmployeeListing;
function _copyAndSort(items, columnKey, isSortedDescending) {
    var key = columnKey;
    return items.slice(0).sort(function (a, b) { return ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1); });
}
//# sourceMappingURL=EmployeeListing.js.map