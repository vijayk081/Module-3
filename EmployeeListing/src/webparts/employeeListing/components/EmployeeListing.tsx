import * as React from 'react';

// import styles from './EmployeeListing.module.scss';
import { IEmployeeListingProps } from './IEmployeeListingProps';
// import { escape } from '@microsoft/sp-lodash-subset';
import { TextField, ITextFieldStyles } from 'office-ui-fabric-react/lib/TextField';
import { DetailsList, DetailsListLayoutMode, Selection, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { MarqueeSelection } from 'office-ui-fabric-react/lib/MarqueeSelection';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { mergeStyles } from 'office-ui-fabric-react/lib/Styling';
import { Announced } from 'office-ui-fabric-react/lib/Announced';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { Dialog, DialogFooter } from '@fluentui/react/lib/Dialog';
import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs"
import "@pnp/sp/lists";
import "@pnp/sp/items";
import {  Icon } from '@fluentui/react';
import { DatePicker, Dropdown, IDropdownOption } from 'office-ui-fabric-react';
import * as moment from 'moment';
// DatePicker,
// import { Dropdown, DropdownMenuItemType, IDropdownOption, IDropdownProps } from '@fluentui/react/lib/Dropdown';

// import { Fabric } from 'office-ui-fabric-react';


const exampleChildClass = mergeStyles({
  display: 'block',
  marginBottom: '10px',
});


const textFieldStyles: Partial<ITextFieldStyles> = { root: { maxWidth: '300px' } };
export interface IDetailsListBasicExampleItem {
  ID: number;
  Name1: string;
  DOB: any;
  Experience: number;
  Department: string;
  DepartmentId:any;
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
  Name1up: any;
  DepartmentId:any;
  DOB: any;
  DOBup: any;
  Department: any;
  SelectedItem: any;
  Experience: any;
  Experienceup: any;
  hideDialog: boolean;
  projectlookupvalues: IDropdownOption[];
  hideDialogup: boolean;
  EditMode: boolean,
  SelectedItemup: any,



}



export default class EmployeeListing extends React.Component<IEmployeeListingProps, IDetailsListBasicExampleState> {
  private _selection: Selection;
  private _columns: IColumn[];



  constructor(props: IEmployeeListingProps) {
    super(props);
    sp.setup({
      spfxContext: this.props.spfxcontext
    });


    this.state = {
      items: [],
      selectionDetails: '',
      announcedMessage: undefined,
      FilterData: [],
      id: '',
      ItemId: [],
      DepartmentId:'',
      SelectedItem: 0,
      text: 'string',
      Name1: '',
      Name1up: '',
      DOB: null,
      DOBup: null,
      Department: '',
      projectlookupvalues: [],
      Experience: '',
      Experienceup: '',
      hideDialog: true,
      hideDialogup: true,
      EditMode: false,
      SelectedItemup: '',

    };
    this._getdeplookupfield();

    this._selection = new Selection({
      onSelectionChanged: () => this.setState({ selectionDetails: this._getSelectionDetails() }),
    });


    this._columns = [
      { key: 'column1', name: 'Action', fieldName: 'Action', minWidth: 100, maxWidth: 200, isResizable: true },
      { key: 'column2', name: 'ID', fieldName: 'ID', minWidth: 100, maxWidth: 200, isResizable: true },
      { key: 'column3', name: 'Name', fieldName: 'Name1', minWidth: 100, maxWidth: 200, isResizable: true, onColumnClick: this._onColumnClick },
      { key: 'column4', name: 'DOB', fieldName: 'DOB', minWidth: 100, maxWidth: 200, isResizable: true },
      { key: 'column5', name: 'Experience', fieldName: 'Experience', minWidth: 100, maxWidth: 200, isResizable: true },
      { key: 'column6', name: 'Department', fieldName: 'Department', minWidth: 100, maxWidth: 200, isResizable: true },
    ];


  }


  render(): React.ReactElement<IEmployeeListingProps> {
    const { items, selectionDetails } = this.state;


    return (
      <div>
        <Fabric>
          <div className={exampleChildClass}>{selectionDetails}</div>
          <Announced message={selectionDetails} />
          <TextField
            className={exampleChildClass}
            label="Filter by name:"
            onChange={this._onFilter}
            styles={textFieldStyles}
          />

          <Announced message={`Number of items after filter applied: ${items.length}.`} />
          <PrimaryButton text="Add Employee" onClick={() => { this.setState({ hideDialog: false , EditMode: false }) }} />
          <MarqueeSelection selection={this._selection}>
            <DetailsList
              items={this.state.items}
              columns={this._columns}
              setKey="set"
              layoutMode={DetailsListLayoutMode.justified}
              selection={this._selection}
              selectionPreservedOnEmptyClick={true}
              ariaLabelForSelectionColumn="Toggle selection"
              ariaLabelForSelectAllCheckbox="Toggle selection for all items"
              checkButtonAriaLabel="select row"
              onItemInvoked={this._onItemInvoked}
            />
          </MarqueeSelection>

        </Fabric>
        <div className='adddialog'>
          <Dialog
            hidden={this.state.hideDialog}>
            <h1>{this.state.EditMode ? "Update" : "Add Employee"}</h1>
            <div className='table'>
              <table>
                <tr className='Name1'>
                  <td>
                    Name :
                  </td>
                  <td>
                    <input type="text" id="Name1" value={this.state.Name1 ? this.state.Name1 : ''} onChange={(e) => { this.setState({ Name1: e.target.value }); }} />
                  </td>
                </tr>
                <tr className='DOB'>
                  <td>
                    DOB :
                  </td>
                  <td>
                  <DatePicker id="DOB"
                      value={new Date(this.state.DOB)}
                      onSelectDate={(selectedDate) => {
                        this.setState({ DOB: selectedDate });
                      }} />
                    {/* <input type="date" id="DOB" value={this.state.DOB ? this.state.DOB : ''} onChange={(e) => { this.setState({ DOB: e.target.value }); }} /> */}
                  </td>
                </tr>
                <tr className='Experience'>
                  <td>
                    Experience :
                  </td>
                  <td>
                    <input type="number" id="Experience" value={this.state.Experience ? this.state.Experience :''} onChange={(e) => { this.setState({ Experience: e.target.value }); }} />
                  </td>
                </tr>
                <tr className='Department'>
                  <td>
                    Department :
                  </td>
                  <td>

                    <Dropdown placeholder="Select a Department" options={this.state.projectlookupvalues} defaultSelectedKey={this.state.SelectedItemup} onChange={(e, val) => { this.onDropdownchange(e, val) }} ></Dropdown>
                  </td>
                </tr>
              </table>
            </div>
            <DialogFooter>
              <PrimaryButton  onClick={() => { this.state.EditMode ? this.updatedialog() : this.createItem()  }} >{this.state.EditMode ? "Update" : "Save"} </PrimaryButton>
              <DefaultButton text="Cancel" onClick={() => { this.setState({ hideDialog: true }), this.reset() }} />
            </DialogFooter>
          </Dialog>
          <Dialog hidden={this.state.hideDialogup}  >
            <text>
              Are you sure want to update details?
            </text>
          <DialogFooter>
              <PrimaryButton text="yes" onClick={() =>  this.UpdateItem(this.state.ItemId)} />
              <DefaultButton text="No" onClick={() => { this.setState({ hideDialogup: true }), this.reset() }} />
            </DialogFooter>
          </Dialog>
        </div>

        {/* <div className='updialog'>
          <Dialog
            hidden={this.state.hideDialogup}>
            <h1>update</h1>
            <div className='table'>
              <table>
                <tr className='Name1up'>
                  <td>
                    Name :
                  </td>
                  <td>
                    <input type="text" id="Name1" value={this.state.Name1 ? this.state.Name1 : ''} onChange={(e) => { this.setState({ Name1: e.target.value }); }} />
                  </td>
                </tr>
                <tr className='DOBup'>
                  <td>
                    DOB :
                  </td>
                  <td>
                    <DatePicker id="DOBup"
                      value={new Date(this.state.DOB)}
                      onSelectDate={(selectedDate) => {
                        this.setState({ DOB: selectedDate });
                      }} />
                  </td>
                </tr>
                <tr className='Experienceup'>
                  <td>
                    Experience :
                  </td>
                  <td>
                    <input type="number" id="Experienceup" value={this.state.Experience ? this.state.Experience : ''} onChange={(e) => { this.setState({ Experience: e.target.value }); }} />
                  </td>
                </tr>
                <tr className='Department'>
                  <td>
                    Department :
                  </td>
                  <td>
                    <Dropdown placeholder="Select a Department" options={this.state.projectlookupvalues} defaultSelectedKey={this.state.SelectedItemup} onChange={(e, val) => { this.onDropdownchange(e, val) }} ></Dropdown>
                  </td>
                </tr>
              </table>
            </div>
            <DialogFooter>
              <PrimaryButton text="Update" onClick={() => { this.UpdateItem(this.state.ItemId) }} />
              <DefaultButton text="Cancel" onClick={() => { this.setState({ hideDialogup: true }) }} />
            </DialogFooter>
          </Dialog>
        </div> */}
      </div>
    );
  }

  public updatedialog= async () =>{
    this.setState({
      hideDialogup:false,
    })
  }

  public onDropdownchange(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption) {
    console.log();
    this.setState({ SelectedItem: item.key , SelectedItemup: item.key })
  }

  public componentDidMount = () => {
    this.test();
    this._getdeplookupfield();
  }
  public test = async () => {
    await sp.web.lists.getByTitle("Employees").items.select("ID,Name1,DOB,Experience,Department/ID,Department/DepartmentName").expand("Department").get().then(items => {

      let AllData: { Action: any; ID: any; Name1: any; DOB: any; Experience: number; Department: any; DepartmentId:any;}[] = [];
      items.map((data) => {
        let dataDOB = moment(data.DOB).format('MM/DD/YYYY');

        AllData.push({
          ID: data.ID,
          Name1: data.Name1,
          DOB: dataDOB,
          Department: data.Department.DepartmentName,
          Experience: data.Experience,
          DepartmentId:data.Department.ID,
          Action: (
            <>
              <Icon
                iconName='delete' onClick={() => { this._onclickdelete(data.ID) }} style={{ marginRight: 30 }}>
              </Icon>
              <Icon
                iconName='EditSolid12' onClick={() => { this.editModeItems(data.ID) }} >
              </Icon>
            </>
          )
        })
      })
      this.setState({
        items: AllData,
        selectionDetails: this._getSelectionDetails(),
        FilterData: AllData,
      });
    }).catch((e) => {
      console.log(e);
    })
  }

  //for combination of dialog
    public reset = async() =>{
      this.setState({ Name1:'',DOB:null,Experience:'',SelectedItemup:0 })
    }

  // create item
  public createItem = async () => {

    sp.web.lists.getByTitle("Employees")
      .items.add({
        Name1: this.state.Name1,
        DOB: this.state.DOB,
        Experience: this.state.Experience,
        DepartmentId: this.state.SelectedItem,

      }).then(() => {
        this.setState({ hideDialog: true })
        this.test()

      }).catch((err) => {
        console.log(err);
      });
  }

  private _onclickdelete = async (ID: any) => {
    console.log(ID);
    await sp.web.lists.getByTitle("Employees").items.getById(ID).delete().then((data) => {
      console.log(data);
      this.test();
    }).catch((err) => {
      console.log(err);
    });
  }

  public editModeItems = (Id: any) => {
    let editItem = this.state.items.filter((x: any) => { return x.ID == Id; })[0];
    this.setState({
      Name1: editItem.Name1,
      DOB: editItem.DOB,
      Experience: editItem.Experience,
      ItemId: editItem.ID,
      Department: editItem.Department,
      SelectedItemup:editItem.DepartmentId,
      hideDialog: false,
      EditMode: true
    });
  }
  
  public UpdateItem = async (ItemId: any) => {
    sp.web.lists.getByTitle("Employees").items.getById(ItemId)
      .update({
        Name1: this.state.Name1,
        DOB: this.state.DOB,
        Experience: this.state.Experience,
        DepartmentId: this.state.SelectedItemup,

      }).then(() => {
        // location.reload();
        this.setState({ hideDialogup: true, hideDialog:true });
        this.reset(),
        this.test();
      })
      .catch((err) => {
        console.log(err);
      });
  }

  public _getdeplookupfield = async () => {
    const allItems: any[] = await sp.web.lists.getByTitle("Department").items.getAll();
    let dropdowndep: IDropdownOption[] = [];
    allItems.forEach(Department => {
      dropdowndep.push({ key: Department.ID, text: Department.DepartmentName });
    })
    this.setState({
      projectlookupvalues: dropdowndep
    });


  }

  private _getSelectionDetails(): string {
    const selectionCount = this._selection.getSelectedCount();

    switch (selectionCount) {
      case 0:
        return 'No items selected';
      case 1:
        return '1 item selected: ' + (this._selection.getSelection()[0] as IDetailsListBasicExampleItem).Name1;
      default:
        return `${selectionCount} items selected`;
    }
  }

  private _onItemInvoked = (item: IDetailsListBasicExampleItem): void => {
    alert(`Item invoked: ${item.Name1}`);
  };
  
  private _onFilter = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text: string): void => {
    this.setState({
      items: text ? this.state.FilterData.filter(i => i.Name1.toLowerCase().indexOf(text) > -1) : this.state.FilterData,
    });
  };

  private _onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
    const { items } = this.state;
    const newColumns: IColumn[] = this._columns.slice();
    const currColumn: IColumn = newColumns.filter(currCol => column.key === currCol.key)[0];
    newColumns.forEach((newCol: IColumn) => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;
        this.setState({
          announcedMessage: `${currColumn.name} is sorted ${currColumn.isSortedDescending ? 'descending' : 'ascending'
            }`,
        });
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });
    const newItems = _copyAndSort(items, currColumn.fieldName!, currColumn.isSortedDescending);
    this.setState({
      items: newItems,
    });
  };
}

function _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
  const key = columnKey as keyof T;
  return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
}
