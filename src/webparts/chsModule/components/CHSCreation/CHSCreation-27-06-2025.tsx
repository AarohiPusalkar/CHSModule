import * as React from 'react';
import { PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import styles from '../ChsModule.module.scss'
import * as moment from 'moment'
import { IChsModuleProps } from '../IChsModuleProps';
import UseUtilities, { IUtilities } from '../../../services/bal/utilities';
import Utilities from '../../../services/bal/utilities';
import { Formik, FormikProps, ErrorMessage, Field } from 'formik';
import * as yup from 'yup';
import { Web } from '@pnp/sp/presets/all';
import { BaseButton, Button, Checkbox, FontWeights, IPersonaProps } from 'office-ui-fabric-react';
import { Link, useHistory } from 'react-router-dom';
import useSPCRUD, { ISPCRUD } from '../../../services/bal/spcrud';
import SPCRUD from '../../../services/bal/spcrud';
import EmployeeOps from '../../../services/bal/EmployeeMaster';
import { IEmployeeMaster } from '../../../services/interface/IEmployeeMaster';
import { ICHSRequest } from '../../../services/interface/ICHSRequest';
import { keys } from '@microsoft/sp-lodash-subset';
import { IEmployeeCHSLimitMaster } from '../../../services/interface/IEmployeeCHSLimitMaster';
import EmployeeCHSLimitMasterOps from '../../../services/bal/EmployeeCHSLimitMaster';
import { Icon, DefaultButton, Dialog, DialogFooter, DialogType, Dropdown, IDropdownOption, PrimaryButton, IDropdown, } from 'office-ui-fabric-react';
import { Pivot, PivotItem, IPivotItemProps, PivotLinkSize, PivotLinkFormat } from 'office-ui-fabric-react/lib/Pivot';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { escape } from '@microsoft/sp-lodash-subset';
import { Items, sp } from 'sp-pnp-js';
import { CurrentUser } from 'sp-pnp-js/lib/sharepoint/siteusers';
import Swal from "sweetalert2";
// import Select from "react-select";
import Select from 'react-select-plus';
import 'react-select-plus/dist/react-select-plus.css';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { ENV_CONFIG } from '../../../../Enviroment/envConfig';
SPComponentLoader.loadCss('https://cdn.jsdelivr.net/npm/bootstrap@4.6.0/dist/css/bootstrap.min.css');
SPComponentLoader.loadCss('https://stackpath.bootstrapcdn.com/font-awesome/4.7.0/css/font-awesome.min.css');
import '../../assets/CHSCretaionstyle.scss'; 


const initialValues = {
  NoteTypeId: 0,
  GroupId: '',
  FinancialYearId: "",
  Subject: "",
}
const validate = yup.object().shape({
  FinancialYearId: yup.number().required('Financial Year Value is required'),
  Subject: yup.string().required('Subject is required'),
});

export interface ISelectState {
  selectedOption?: string;
}

// const useHistory =  useHistory();

const onbehalfoption: IDropdownOption[] = [
  { key: 'Yes', text: 'Yes' },
  { key: 'No', text: 'No' }
];

export default class CHSCreation extends React.Component<IChsModuleProps, any> {

  constructor(props: any) {
    super(props);
    this.state = {
      activeHRTab: localStorage.getItem("activeHRTab") || "tab1",
    };
    this.state = {
      activeTab: "Pending", // Set the initial active tab here
    };
    this.state = {
      GroupId: "",
      AllEmployeeCollObj: [],
      // allDashboardData:[],
      // NoteTypeId: "",
      // FinancialYearColl: [],
      // FinancialYearId: "",
      // Subject: "",
      // isMultiGrp: false,
      selectedOptionCHBx: null,
      isSubmitting: false,
      selectedOption: '',
      transformedOptions: [],
      IsSpouseEximEmployee: false,
      filteredData: [],
      showhideEmployeeNameLab: false,
      isOnBehalfDisabled: false,
      ShowHR1Tab: true,
      ShowHR2Tab: true,
      ShowHRTab: true,
      ActualClaimAmountLable: "",
      Currentuser: "",
      UserDashboard: [],
      userApprovedDashboard: [],
      userRejectedDashboard: [],
      HR1ApprPendingDashboard: [],
      HR1ApproverApprDashboard: [],
      HR1ApproverRejectDashboard: [],
      HR2ApprPendingDashboard: [],
      HR2ApproverApprDashboard: [],
      HR2ApproverRejectDashboard: [],
      DependentType: "ALL",
      allDashboardData: [],

      allDashboardDataHR1Pending: [],

      allDashboardDataHR1Rejected: [],

      allDashboardDataHR1Approved: [],

      allDashboardDataHR2Pending: [],

      allDashboardDataHR2Rejected: [],

      allDashboardDataHR2Approved: [],


      allDashboardData2: [],

      // HR2ApproverApprDashboard: [], // Original data
      filteredDashboard: [], // Filtered data for rendering
      // PerticularMaster: [],
      isDialogVisible: false,
      isDialogGH: false,
      isDialogHR1: false,
      isDialogHR2: false,
      isDialogViewHR1: false,
      isDialogViewHR2: false,
      EmployeeName: "",
      ExpenseDetailsAlert: false,
      searchValue: "",
      filteredEmployees: [],
      // Approver: "",
      // PaymentType: "NEFT",
      // AccountNo: "",
      // IFSCCode: "",
      // ClaimFor: "Self",
      // VendorDetails: "",
      EmployeeID: '',
      EmployeeIDId: '',
      DesignationId: '',
      CompanyEmail:'',

      // GradeId: '',
      // CurrentOfficeLocationId: '',
      // EmployeeSubGroupId: '',
      GHStatus: '',
      // Role: '',
      // ApproverId: '',
      // Level2Id: '',
      // Level1Id: '',
      // CashApproverId: '',
      // PettcashDetailsColl: [],
      // TotalValue: "",
      file: null,
      reqID: '',
      isClearable: true,
      isSearchable: true,
      CHSApproverView: [],


      filteredOptions: [],
      // filteredOptions: [],
      selectedId: null,
      isDropdownOpen: false, //  Track dropdown state
      // searchValue: "",
      // filteredOptions: this.state.AllEmployeeCollObj,
//  activeHR1Tab: 'Pending',
//  activeHRTab: 'Pending',
  selectedOuterTab: ''
    };
    // this.getSelectedEmployeeDetail=this.getSelectedEmployeeDetail.bind(this);  


  }
  async componentDidMount() {
    const savedTab1 = localStorage.getItem("activeHR1Tab"); // For HR1
    const savedTab2 = localStorage.getItem("activeHR2Tab"); // For HR2
    
    if (savedTab1) {
      this.setState({ activeHR1Tab: savedTab1 });
    }
    
    if (savedTab2) {
      this.setState({ activeHR2Tab: savedTab2 });
    }
   
 
    let tabLinks = document.querySelectorAll(".tablink");

// Get saved active tab from localStorage
let savedTabId = localStorage.getItem("activeTabId");

// If there's a saved tab ID, use it; otherwise, use the default
if(savedTabId){
  let activeTab = savedTabId? document.getElementById(savedTabId)
  : document.getElementById("defaultOpen");
  activeTab.classList.add("active");
}

else{
  let activeTab = document.getElementById("defaultOpen");
 // activeTab.classList.add("active");
}


tabLinks.forEach((tab) => {
  tab.addEventListener("click", function () {
    // Remove active class from all tabs
    tabLinks.forEach((t) => t.classList.remove("active"));

    // Add to the clicked tab
    this.classList.add("active");

    // Save the clicked tab ID to localStorage
    localStorage.setItem("activeTabId", this.id);
  });
});

     await this.getCurrentUser();
    await this.UserApprovedDashboards();
    await this.UserRejectedDashboards();
    await this.UserPendingDashboard();
    await this.HR1ApprovePendingDashboard();
    await this.HR1ApproveApprovedDashboards();
    await this.HR1ApproveRejectedDashboards();
    await this.HR2ApprovePendingDashboard();
    await this.HR2ApproveApprovedDashboards();
    await this.HR2ApproveRejectedDashboards();
    await this.GetEmployeelimit();
    await this.checkUserInGroups(["HR1_Group", "HR2_Group"]);
    await this.checkUserInGroupsForHR1Tab(["HR1_Group"]);
    await this.checkUserInGroupsForHR2Tab(["HR2_Group"]);
    await this.getEmployee();
    await this.GetEmployeelimit();
    
}
handleOuterTabClick = (tabName) => {
  this.setState({
    selectedOuterTab: tabName, 
    activeHRTab: 'Pending',
    activeHR1Tab: 'Pending',
    activeHR2Tab: 'Pending'
  }); 
}
handleTabClick = (event) => {
  const tabLinks = document.querySelectorAll(".tablink");
 
  // Remove the "active" class from all tab buttons
  tabLinks.forEach((tab) => {
    tab.classList.remove("active"); 
  });

  // Add the "active" class to the clicked button
  event.target.classList.add("active");
  }

  componentDidUpdate(prevProps, prevState) {
    if (prevState.showhideEmployeeNameLab !== this.state.showhideEmployeeNameLab && !this.state.showhideEmployeeNameLab) {
      this.setState({ selectedOption: null });
    }
  }


  private  handleChangeddl = (selectedOption) => {  
    this.setState({ selectedOption });  
    console.log(`Selected: ${selectedOption.label}`);  
  }  

  public getCurrentUser = async () => {
    const spCrudObj = await useSPCRUD();
    return await spCrudObj.currentUser(this.props).then(cuser => {
      this.setState({ Currentuser: cuser });
      return cuser;
    });
  }

  private handleChangeDrop = (selectedOption) => {
    this.setState({ selectedOption });
    console.log(`Selected: ${selectedOption.label}`);
  }

  public async checkUserInGroups(groups: any) {
    try {
      const spCrudObj = await useSPCRUD();
      const userGroups = await spCrudObj.currentUserGroup(this.props);
      if (!userGroups || userGroups.length === 0) {
        console.log("User is not part of any group.");
        return false;
      }
      const isUserInGroup = userGroups.some(group => groups.includes(group.Title));
      if (isUserInGroup) {
        this.setState({ showhideEmployeeNameLab: true, OnBehalf: 'No', isOnBehalfDisabled: true })
        console.log(`User exists in at least one of the specified groups.`);
      } else {
        this.setState({ showhideEmployeeNameLab: false, OnBehalf: 'No', isOnBehalfDisabled: false })
        console.log(`User does not exist in any of the specified groups.`);
      }
    } catch (error) {
      console.error("Error checking user in groups:", error);
    }
  }
  public async checkUserInGroupsForHR2Tab(groups: any) {
    try {
      const spCrudObj = await useSPCRUD();
      const userGroups = await spCrudObj.currentUserGroup(this.props);
      if (!userGroups || userGroups.length === 0) {
        console.log("User is not part of any group.");
        return false;
      }
      const isUserInGroup = userGroups.some(group => groups.includes(group.Title));
      if (isUserInGroup) {
        console.log(`User exists in at least one of the specified groups.`);
        this.setState({ ShowHR2Tab: true })
      }
    } catch (error) {
      console.error("Error checking user in groups:", error);
      return false;
    }
  }
  public async checkUserInGroupsForHR1Tab(groups: any) {
    try {
      const spCrudObj = await useSPCRUD();
      const userGroups = await spCrudObj.currentUserGroup(this.props);
      if (!userGroups || userGroups.length === 0) {
        console.log("User is not part of any group.");
        return false;
      }
      const isUserInGroup = userGroups.some(group => groups.includes(group.Title));
      if (isUserInGroup) {
        this.setState({ ShowHR1Tab: true })
      }
    } catch (error) {
      console.error("Error checking user in groups:", error);
      return false;
    }
  }
  public getEmployee = async (): Promise<IEmployeeMaster> => {
    return await EmployeeOps().getEmployeeMaster(this.props).then(async (results) => {
      let employeeData = results;
      let limitData: IEmployeeCHSLimitMaster[] = await this.GetEmployeelimit();
      let matchedLimit = limitData.filter((e) => e.Scale.Title == employeeData.Scale && e.Designation.Title == employeeData.DesignationTitle && e.EmployeeType == employeeData.EmpType);
      this.setState({
        EmployeeInfodb: employeeData,
        AllEmployeeCollObj: [],
        EmployeeName: employeeData.EmployeeName,
        EmployeeIDId: employeeData.Id,
        DependentType: "",
        ActualClaimAmountLable: "",
        // Approver: employeeData.LeaveLevel2.Title,
        // ApproverId: employeeData.LeaveLevel2Id,
        // Level2Id: employeeData.LeaveLevel2Id,
        // AccountNo: employeeData.AccountNo,
        // IFSCCode: employeeData.IFSCCode,
        CompanyEmail:employeeData.CompanyEmail,

        EmployeeID: employeeData.EmployeeId,
        DesignationId: employeeData.DesignationId,
        DesignationTitle: employeeData.DesignationTitle,
        DateofBirth: employeeData.DateofBirth,
        Scale: employeeData.Scale,
        Age: parseInt(employeeData.Age),
        EmpType: employeeData.EmpType,
        Limit: matchedLimit.length > 0 && matchedLimit !== undefined ? matchedLimit[0].Limit : "",
        // GradeId: employeeData.GradeId,
        // CurrentOfficeLocationId: employeeData.CurrentOfficeLocationId,
        // EmployeeSubGroupId: employeeData.SubGroupId,
        // Role: employeeData.Role
      });
      return employeeData;
    });
  };
  public GetEmployeelimit = async (): Promise<IEmployeeCHSLimitMaster[]> => {
    return await EmployeeCHSLimitMasterOps().getAllEmployeeCHSLimit(this.props).then(CHSLimitresults => {
      let formattedResults: IEmployeeCHSLimitMaster[] = Array.isArray(CHSLimitresults) ? CHSLimitresults : [CHSLimitresults];
      this.setState({
        EmployeeCHSInfodb: formattedResults
      });
      return formattedResults;
    });
  };

  public getSelectedEmployeeDetail = (e) => {
    if (e.key !== undefined) {
      const { EmployeeInfodb } = this.state;
      if (EmployeeInfodb.length > 0) {
        let selectedEmp = EmployeeInfodb.find(item => e.key == item.Id);
        if (selectedEmp !== undefined) {
          let matchedLimit = this.state.EmployeeCHSInfodb.filter(
            (lim) =>
              lim.Scale.Title === selectedEmp.Scale.Title &&
              lim.Designation.Title === selectedEmp.Designation.Title &&
              lim.EmployeeType === selectedEmp.EmployeeType.Title
          );
          this.setState({
            EmployeeName: selectedEmp.EmployeeName,
            // EmployeeIDId: selectedEmp.Id,
            // Approver: selectedEmp.LeaveLevel2.Title,
            // ApproverId: selectedEmp.LeaveLevel2Id,
            // Level2Id: selectedEmp.LeaveLevel2Id,
            // AccountNo: selectedEmp.AccountNo,
            // IFSCCode: selectedEmp.IFSCCode,
            isDropdownOpen: false, //  Close dropdown on selection
            CompanyEmail:selectedEmp.CompanyEmail,

            // this.setState({ EmployeeID: e.key });
            EmployeeID: selectedEmp.Title,
            DesignationId: selectedEmp.DesignationId,
            DesignationTitle: selectedEmp.Designation.Title,
            DateofBirth: selectedEmp.DOB
              ? `${new Date(selectedEmp.DOB).getDate()}-${new Date(selectedEmp.DOB).getMonth() + 1}-${new Date(selectedEmp.DOB).getFullYear()}`
              : null,
            Scale: selectedEmp.Scale.Title,
            Payscale: selectedEmp.Payscale.Title,
            Age: parseInt(selectedEmp.Age),
            EmpType: selectedEmp.EmployeeType.Title,
            DependentType: "",
            ActualClaimAmountLable: "",
            selectedOption:e,
            //   ExpenseDetails.Amount:""
            // GradeId: selectedEmp.GradeId,
            // CurrentOfficeLocationId: selectedEmp.CurrentOfficeLocationId,
            // EmployeeSubGroupId: selectedEmp.SubGroupId,
            // Role: selectedEmp.Role,
            Limit: matchedLimit.length > 0 ? matchedLimit[0].Limit : "",
          });
        }
      }
    }
    this.setState({ e });  

  };
  public getAllEmployee = async (): Promise<IEmployeeMaster> => {
    return await EmployeeOps().getAllEmployeeMaster(this.props).then(results => {
      let employeeOptions = results.map((item: any) => ({
        key: item.Id.toString(),
        text: item.EmployeeName,
        label:item.EmployeeName,
        value:item.EmployeeName,
      }));
      this.setState({
        EmployeeInfodb: results,
        // DependentType: "",
        // ActualClaimAmountLable: "",
        AllEmployeeCollObj: employeeOptions,
        filteredOptions: employeeOptions,

      });
      return results;
    });
  };
  public _getPeoplePickerItems = (item: IPersonaProps[]) => {
    console.log(item);
  }
  private _getPeoplePickerItems1(items: any[]) {
    console.log('Items:', items);
  }
  handleDropdownChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, field?: string) => {
    if (option && field) {
      this.setState({ [field]: option.key });
    }
  }
  public closeDialog = () => {
    this.setState({
      isDialogVisible: false,
      isDialogHR1: false,
      isDialogHR2: false,
      isDialogViewHR1: false,
      isDialogViewHR2: false
    });
  }
  public CreateRequest = () => {
    this.setState({ isDialogVisible: true,
      activeTab: 'Pending'
     })
    
  }
 
  public openPage(pageName, elmnt, color,tabName) {
    this.setState({ activeTab: tabName });
    var i, tabcontent, tablinks;
    tabcontent = document.getElementsByClassName("tabcontent");
    for (i = 0; i < tabcontent.length; i++) {
      tabcontent[i].style.display = "none";
    }
    tablinks = document.getElementsByClassName("tablink");
    for (i = 0; i < tablinks.length; i++) {
      tablinks[i].style.backgroundColor = "";
    }
    document.getElementById(pageName).style.display = "block";
    elmnt.style.backgroundColor = color;
  }
  public handleInputChangeadd = (e) => {
    const value = e.target.value;
    switch (e.target.name) {
      case "ExpenseDetails.Amount":
        this.setState({
          ExpenseDetails: {
            ...this.state.ExpenseDetails,
            Amount: +value,
          }
        });
        if (this.state.ActualClaimAmountLable != "") {
          if (value > this.state.ActualClaimAmountLable) {
            this.setState({
              ExpenseDetailsAlert: true
            });
          }
          else {
            this.setState({
              ExpenseDetailsAlert: false
            });
          }
        }
        break;
      case "ExpenseDetails.HR1Remarks":
        this.setState({
          ExpenseDetails: {
            ...this.state.ExpenseDetails,
            HR1Remarks: '' + value
          }
        });
        break;
      case "ExpenseDetails.HR2Remarks":
        this.setState({
          ExpenseDetails: {
            ...this.state.ExpenseDetails,
            HR2Remarks: '' + value
          }
        });
        break;
      case "ExpenseDetails.FinalAmount":
        this.setState({
          ExpenseDetails: {
            ...this.state.ExpenseDetails,
            FinalAmount: + value
          }
        });
        break;

      case "ExpenseDetails.HR1FinalAmount":
        this.setState({
          // ExpenseDetailsAlert:true,
          ExpenseDetails: {
            ...this.state.ExpenseDetails,
            HR1FinalAmount: + value
          }
        });

        if (value > this.state.CHSApproverView.EligibilityLimit) {
          this.setState({
            ExpenseDetailsAlert: true
          });
        }
        else
        // if(e.target.value <= parseFloat(this.state.CHSApproverView.EligibilityLimit))
        {
          this.setState({
            ExpenseDetailsAlert: false
          });
        }
        break;

      // case "ExpenseDetails.HR2FinalAmount":
      //   this.setState({
      //     ExpenseDetails: {
      //       ExpenseDetailsAlert:true,

      //       ...this.state.ExpenseDetails,
      //       FinalAmount: + value
      //     }
      //   });
      //   break;


      case "ExpenseDetails.HR2FinalAmount":
        this.setState({
          // ExpenseDetailsAlert:true,
          ExpenseDetails: {
            ...this.state.ExpenseDetails,
            HR2FinalAmount: + value
          }
        });

        if (value > this.state.CHSApproverView.EligibilityLimit) {
          this.setState({
            ExpenseDetailsAlert: true
          });
        }
        else
        // if(+e.target.value <= parseFloat(this.state.CHSApproverView.EligibilityLimit))
        {
          this.setState({
            ExpenseDetailsAlert: false
          });
        }
        break;

      default:
        break;
    }
  }
  // _handleFileChange = (event) => {
  //   this.setState({ files: event.target.files });
  // };
  _handleFileChange = (event) => {
    const files = event.target.files;
    const validFiles = [];
    const invalidFiles = [];
    // Updated regex to allow spaces, letters, numbers, dots, and underscores
    const fileNameRegex = /^[a-zA-Z0-9._\s]+$/;
    // Convert FileList to an array manually
    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      if (fileNameRegex.test(file.name)) {
        validFiles.push(file); // 
      } else {
        invalidFiles.push(file.name); // 
      }
    }
    // Show alert only if there are invalid files
    if (invalidFiles.length > 0) {
      alert(` These files have invalid names and won't be uploaded: ${invalidFiles.join(", ")}\nAllowed characters: letters, numbers, dots, underscores, and spaces.`);
    }
    // Manually reset file input to allow re-uploading the valid files
    const dataTransfer = new DataTransfer();
    validFiles.forEach(file => dataTransfer.items.add(file));
    event.target.files = dataTransfer.files;
    // Set only valid files in state
    this.setState({ files: validFiles });
  };
  handleChange = (e) => {
    const value = e.target.value;
    console.log('Input value:', value);
    this.setState({
      [e.target.name]: value,
      VendorDetails: value,
    });
  };
  public EligibleClaimAmount = (e) => {

    if (e == 'Spouse') {

      this.setState({

        IsSpouseEximEmployee: true
      })


    }
    if (e == 'Self') 
      {

        this.setState({

          IsSpouseEximEmployee: false
        })

    }
    let ActualClaimAmountLable;
    const LIMIT = this.state.Limit ? parseFloat(this.state.Limit) : 0;
    const scaleValue = this.state.Scale;
    if (this.state.EmpType != 'RETIRED') {
      if (this.state.Scale !== "GOVT") {
        const numberOnlyScale = parseFloat(scaleValue.match(/\d+$/)[0]);
        if (this.state.Age < 40 && numberOnlyScale <= 5) {
          let Scale5LIMIT = this.state.Limit ? parseFloat(this.state.Limit) : 0;
          const Scale5LIMITSelf = (Scale5LIMIT) * 0.9;
          ActualClaimAmountLable = Scale5LIMITSelf;
          if (e == 'Spouse') {
            let Scale5LIMIT = this.state.Limit ? parseFloat(this.state.Limit) : 0;
            const Scale5LIMITSpouse = (Scale5LIMIT) * 0.75
            ActualClaimAmountLable = Scale5LIMITSpouse;
            if (Scale5LIMITSpouse > 15000) {
              ActualClaimAmountLable = "15000";
            }
          }
        }
        if (this.state.Age > 40 && numberOnlyScale <= 5) {
          let Scale5LIMIT = this.state.Limit ? parseFloat(this.state.Limit) : 0;
          const Scale5LIMITSelf = (Scale5LIMIT) * 1;
          ActualClaimAmountLable = Scale5LIMITSelf;
          if (e == 'Spouse') {
            let Scale5LIMIT = this.state.Limit ? parseFloat(this.state.Limit) : 0;
            const Scale5LIMITSpouse = (Scale5LIMIT) * 0.75
            ActualClaimAmountLable = Scale5LIMITSpouse;
            if (Scale5LIMITSpouse > 15000) {
              ActualClaimAmountLable = "15000";
            }
          }
        }
        if (numberOnlyScale > 5) {
          const Scale5LIMITSelf = this.state.Limit ? parseFloat(this.state.Limit) : 0;
          ActualClaimAmountLable = Scale5LIMITSelf;
          if (e == 'Spouse') {
            const Scale5LIMITSpause = this.state.Limit ? parseFloat(this.state.Limit) : 0;
            const Scale5LIMITSpause1 = (Scale5LIMITSpause) * 0.75
            ActualClaimAmountLable = Scale5LIMITSpause1;
            if (Scale5LIMITSpause1 > 20000) {
              ActualClaimAmountLable = "20000";
            }
          }
        }
      }
      else {
        const Scale5LIMITSelf = this.state.Limit ? parseFloat(this.state.Limit) : 0;
        ActualClaimAmountLable = Scale5LIMITSelf;
        if (e == 'Spouse') {
          const Scale5LIMITSpause = this.state.Limit ? parseFloat(this.state.Limit) : 0;
          const Scale5LIMITSpause1 = (Scale5LIMITSpause) * 0.75
          ActualClaimAmountLable = Scale5LIMITSpause1;
          if (Scale5LIMITSpause1 > 20000) {
            ActualClaimAmountLable = "20000";
          }
        }
      }
    }
    if (this.state.EmpType == 'RETIRED') {
      const Scale5LIMITSelf = this.state.Limit ? parseFloat(this.state.Limit) : 0;
      const Scale5LIMITSelf1 = (Scale5LIMITSelf) * 0.9;
      ActualClaimAmountLable = Scale5LIMITSelf1;
      if (e == 'Spouse') {
        const Scale5LIMITSpause = this.state.Limit ? parseFloat(this.state.Limit) : 0;
        const Scale5LIMITSpause1 = (Scale5LIMITSpause) * 0.75
        ActualClaimAmountLable = Scale5LIMITSpause1;
        if (ActualClaimAmountLable > 15000) {
          ActualClaimAmountLable = "15000"
        }
      }
    }
    this.setState({
      DependentType: e,
      ActualClaimAmountLable: ActualClaimAmountLable,
      selectedOptionCHBx:null,
      ShowLableElibleClaimAmount: true
    });
  };

  
  public BtnSubmitRequest = async () => {
    // "EmployeeName": "Farm Admin",
    // "EmployeeID": "11",
    // if(this.state.EmployeeName)
    if(this.state.DependentType=='Spouse'){
    if(this.state.selectedOptionCHBx==null){
      alert("Please check on 'Is Spouse Exim Employee' ");
      return false;
    }
  }
    const currentDate = new Date();
    const currentYear = currentDate.getFullYear();
    const financialYearStart = new Date(currentDate.getMonth() < 3 ? currentYear - 1 : currentYear, 3, 1);
    const financialYearEnd = new Date(financialYearStart.getFullYear() + 1, 2, 31);
    let existingRequests
    await EmployeeOps().getEmployeeMasterById(this.props).then(results => {
      existingRequests = results;
    })
    let hasRequestThisYear = [];
    let counter = 0;
    // if (existingRequests.length > 0) {
    //   for (var e = 0; e < existingRequests.length; e++) {
    //     if ((existingRequests[e].DependentType == this.state.DependentType) && (existingRequests[e].EmployeeID == this.state.EmployeeID) && existingRequests[e].Status != "Rejected") {
    //       hasRequestThisYear = existingRequests.filter(
    //         (req) =>
    //           new Date(req.Created) >= financialYearStart &&
    //           new Date(req.Created) <= financialYearEnd && req.EmployeeID == this.state.EmployeeID
    //       );
    //       if (hasRequestThisYear.length > 0) {
    //         counter++;
    //       }
    //     }
    //   }
    // }

    if (existingRequests.length > 0) {
      // Function to normalize date (set time to 00:00:00)
      const normalizeDate = (date) => {
        let d = new Date(date);
        d.setHours(0, 0, 0, 0);
        return d;
      };
    
      let financialYearStartNormalized = normalizeDate(financialYearStart);
      let financialYearEndNormalized = normalizeDate(financialYearEnd);
    
      // Filter requests for the current employee and financial year
      let hasRequestThisYear = existingRequests.filter(
        (req) =>
          
          normalizeDate(req.Created) >= financialYearStartNormalized &&
          normalizeDate(req.Created) <= financialYearEndNormalized &&
          req.EmployeeID == this.state.EmployeeID
      );
    
      for (let req of existingRequests) {
        if (
          req.DependentType === this.state.DependentType &&
          req.EmployeeID === this.state.EmployeeID &&
          req.Status !== "Rejected"
        ) {
          if (hasRequestThisYear.length > 0) {
            counter++;
            break; // Exit loop early once a match is found
          }
        }
      }
    }

    
    console.log(counter);
    // if (counter > 0 && this.state.EmpType != "RETIRED") {
    if (counter > 0) {
      alert("You have already submitted a request this financial year!");
      return false;
    }
    const LIMIT = this.state.Limit ? parseFloat(this.state.Limit) : 0;
    const scaleValue = this.state.Scale;
    // const numberOnlyScale = parseFloat(scaleValue.match(/\d+$/)[0]);
    const CLAIMAMOUNT = this.state.ExpenseDetails ? parseFloat(this.state.ExpenseDetails.Amount) : 0;
    if (LIMIT == 0) {
      alert('Please Map Limit !')
      return false;
    }
    if (this.state.ActualClaimAmountLable == "") {
      alert('Please Select Dependent Type');
      return false;
    }
    if (!this.state.ExpenseDetails) {
      alert('Please Enter Claim Amount !!');
      return false;
    }
    if (this.state.ExpenseDetails.Amount == 0 || this.state.ExpenseDetails.Amount == null) {
      alert('Please Enter Claim Amount !!');
      return false;
    }
    if (CLAIMAMOUNT > LIMIT) {
      alert('Claim Amount should be less than Limit! ')
      return false;
    }
    if (this.state.ActualClaimAmountLable < this.state.ExpenseDetails.Amount) {
      alert('Claim Amount should be less than Limit! ')
      return false;
    }
    if (this.state.files == undefined || this.state.files == null) {
      alert('Please Attach Files! ')
      return false;
    }
    this.setState({ isSubmitting: true });
    const spCrudObj = await useSPCRUD();
    var CHSRequestItem;
    if (!this.state.ShowHR1Tab || !this.state.ShowHR2Tab) {
      CHSRequestItem = {
        OnBehalf: this.state.OnBehalf,
        EmployeeID: this.state.EmployeeID,
        Scale: this.state.Scale,
        EmployeeType: this.state.EmpType,
        DependentType: this.state.DependentType,
        Status: "Pending",
        HR1Response: "Pending with HR1",
        EligibilityLimit: +this.state.ActualClaimAmountLable,
        Limit: +this.state.Limit,
        HR2Response: "Pending with HR2",
        RequestorEmail	:this.state.CompanyEmail,
        
        IsSpouseEximMember:this.state.selectedOptionCHBx,
        DateofBirth: new Date(this.state.DateofBirth),
        Designation: this.state.DesignationTitle,
        Age: '' + this.state.Age,
        AmountClaimed: +this.state.ExpenseDetails.Amount,
        EmployeeName: this.state.EmployeeName,
      };
    }
    if (this.state.ShowHR1Tab) {
      CHSRequestItem = {
        OnBehalf: this.state.OnBehalf,
        EmployeeID: this.state.EmployeeID,
        Scale: this.state.Scale,
        EmployeeType: this.state.EmpType,
        DependentType: this.state.DependentType,
        Status: "Approved",
        HR1Response: "Approved by HR1",
        IsSpouseEximMember:this.state.selectedOptionCHBx,
        RequestorEmail	:this.state.CompanyEmail,

        // HR2Response: "",
        HR1ApproverNameId: this.state.Currentuser.Id,
        HR2Response: "Approved by HR2",
        HR1ResponseDate: new Date(),
        EligibilityLimit: +this.state.ActualClaimAmountLable,
        Limit: +this.state.Limit,
        DateofBirth: new Date(this.state.DateofBirth),
        Designation: this.state.DesignationTitle,
        Age: '' + this.state.Age,
        AmountClaimed: +this.state.ExpenseDetails.Amount,
        HRApprovedAmount: +this.state.ExpenseDetails.Amount,
        EmployeeName: this.state.EmployeeName,
      };
    }
    if (this.state.ShowHR2Tab) {
      CHSRequestItem = {
        OnBehalf: this.state.OnBehalf,
        EmployeeID: this.state.EmployeeID,
        Scale: this.state.Scale,
        EmployeeType: this.state.EmpType,
        DependentType: this.state.DependentType,
        Status: "Approved",
        // HR1Response: "",
        HR2Response: "Approved by HR2",
        HR1Response: "Approved by HR1",

        IsSpouseEximMember:this.state.selectedOptionCHBx,
        RequestorEmail	:this.state.CompanyEmail,

        HR2ApproverNameId: this.state.Currentuser.Id,
        // HR2Response: "Pending with HR2",
        HR2ResponseDate: new Date(),
        EligibilityLimit: +this.state.ActualClaimAmountLable,
        Limit: +this.state.Limit,
        DateofBirth: new Date(this.state.DateofBirth),
        Designation: this.state.DesignationTitle,
        Age: '' + this.state.Age,
        AmountClaimed: +this.state.ExpenseDetails.Amount,
        HRApprovedAmount: +this.state.ExpenseDetails.Amount,
        EmployeeName: this.state.EmployeeName,
      };
    }
    // return await spCrudObj.insertData("HealthCheckupService", CHSRequestItem, this.props).then(async (req) => {
    //   this.setState({ reqID: req.data.ID });
    //   const RequestNoGenerate = {
    //     Title: 'CHS000' + req.data.ID,
    //   };
    //   await spCrudObj.updateData("HealthCheckupService", req.data.ID, RequestNoGenerate, this.props)
    //   if (this.state.files && this.state.files.length > 0) {
    //     await this.uploadPRDoc("HealthCheckupService", req.data.ID, this.state.files);
    //     alert('CHS Request Submitted Successfully!');
    //     this.setState({ isSubmitting: false });
    //     this.closeDialog();
    //     location.reload();
    //   } else {
    //     alert('CHS Request Submitted without attachments.');
    //     return false;
    //   }
    //   return req;
    // });

    return await spCrudObj.insertData("HealthCheckupService", CHSRequestItem, this.props).then(async (req) => {
      this.setState({ reqID: req.data.ID });

      const RequestNoGenerate = {
        Title: 'CHS000' + req.data.ID,
      };

      await spCrudObj.updateData("HealthCheckupService", req.data.ID, RequestNoGenerate, this.props);

      if (this.state.files && this.state.files.length > 0) {
        await this.uploadPRDoc("HealthCheckupService", req.data.ID, this.state.files);
        alert('CHS Request Submitted Successfully!');
      } else {
        alert('CHS Request Submitted without attachments.');
      }

      // Ensure dialog closes and dashboard reloads
      this.setState({ isSubmitting: false });
      this.closeDialog();

      // Force dashboard reload
   //   window.location.href = "https://sharepointwebssse.eximbankindia.in/sites/hrm/SitePages/CHSModule.aspx";  // Update the URL path to your dashboard route
      window.location.href = `${ENV_CONFIG.siteUrl}/SitePages/CHSModule.aspx`;  // Update the URL path to your dashboard route

      return req;
    });
  };
  public BtnApproveHR1Request = async () => {
    if (!this.state.ExpenseDetails) {
      alert('Please mention Remarks ')
      return false;
    }
    if (!this.state.ExpenseDetails) {
      if (this.state.ExpenseDetails.HR1FinalAmount == 0) {
        alert('Please Enter Final Amount ')
        return false;
      }
    }
    if (this.state.ExpenseDetails.HR1FinalAmount > this.state.CHSApproverView.EligibilityLimit) {
      alert('Final Claim Amount should be less than Eligible Limit! ')
      return false;
    }
    const spCrudObj = await useSPCRUD();
    debugger;
    var CHSRequestItem = {
      Status: "Pending",
      HR1Response: "Approved by HR1",
      HR1ApproverNameId: this.state.Currentuser.Id,
      HR2Response: "Pending with HR2",
      HR1ResponseDate: new Date(),
      HR1Remark: this.state.ExpenseDetails.HR1Remarks,
      FinalAmount: +this.state.ExpenseDetails.HR1FinalAmount
    };
    return await spCrudObj.updateData("HealthCheckupService", this.state.CHSApproverView.ID, CHSRequestItem, this.props).then(async (req) => {
      alert('CHS Request Approved Successfully!');
      this.closeDialog();
      window.location.href = `${ENV_CONFIG.siteUrl}/SitePages/CHSModule.aspx`;  // Update the URL path to your dashboard route

      return req;
    });
  };
  public BtnApproveHR2Request = async () => {
    if (this.state.ExpenseDetails.HR2FinalAmount > this.state.CHSApproverView.EligibilityLimit) {
      alert('Final Claim Amount should be less than Eligible Limit! ')
      return false;
    }
    if (!this.state.ExpenseDetails) {
      alert('Please mention Remarks ')
      return false;
    }
    if (!this.state.ExpenseDetails) {
      if (this.state.ExpenseDetails.HR2FinalAmount == 0) {
        alert('Please Enter Final Amount ')
        return false;
      }
    }


    const spCrudObj = await useSPCRUD();
    var CHSRequestItem = {
      HR2ApproverNameId: this.state.Currentuser.Id,
      Status: "Approved",
      HR2Response: "Approved by HR2",
      HR2ResponseDate: new Date(),
      HRApprovedAmount: +this.state.ExpenseDetails.HR2FinalAmount,
      HR2Remark: this.state.ExpenseDetails.HR2Remarks
    };
    return await spCrudObj.updateData("HealthCheckupService", this.state.CHSApproverView.ID, CHSRequestItem, this.props).then(async (req) => {
      alert('CHS Request Approved Successfully!');
      this.closeDialog()
      //window.location.href = "https://sharepointwebssse.eximbankindia.in/sites/hrm/SitePages/CHSModule.aspx";  // Update the URL path to your dashboard route
      window.location.href = `${ENV_CONFIG.siteUrl}/SitePages/CHSModule.aspx`;  // Update the URL path to your dashboard route
      
      return req;
    });
  };
  public BtnRejectRequest = async (HRTYPE: string) => {
    const spCrudObj = await useSPCRUD();

    if (HRTYPE == 'HR1') {


      if (this.state.ExpenseDetails != undefined) {
        if (this.state.ExpenseDetails.HR1Remarks == undefined || this.state.ExpenseDetails.HR1Remarks == ""
        ) {
          alert('Please mention Remarks ')
          return false;
        }
      }
      else {
        alert('Please mention Remarks ')
        return false;



      }

      const CHSRequestItem = {
        Status: "Rejected",
        HR1Response: "Rejected by HR1",
        HR1ApproverNameId: this.state.Currentuser.Id,

        HR1Remark: this.state.ExpenseDetails.HR1Remarks
      };

      return await spCrudObj.updateData("HealthCheckupService", this.state.CHSApproverView.ID, CHSRequestItem, this.props).then(async (req) => {
        alert('CHS Request Rejected Successfully!');
        this.closeDialog();
       // window.location.href = "https://sharepointwebssse.eximbankindia.in/sites/hrm/SitePages/CHSModule.aspx";  // Update the URL path to your dashboard route
        window.location.href = `${ENV_CONFIG.siteUrl}/SitePages/CHSModule.aspx`;  // Update the URL path to your dashboard route
  
        return req;
      });

    }

    if (HRTYPE == 'HR2') {
      if (this.state.ExpenseDetails != undefined) {
        if (
          this.state.ExpenseDetails.HR2Remarks == undefined ||

          this.state.ExpenseDetails.HR2Remarks == "") {
          alert('Please mention Remarks ')
          return false;
        }
      }
      else {
        alert('Please mention Remarks ')
        return false;

      }

      const CHSRequestItem = {
        Status: "Rejected",
        HR2Response: "Rejected by HR2",
        HR2ApproverNameId: this.state.Currentuser.Id,

        HR2Remark: this.state.ExpenseDetails.HR2Remarks
      };
      return await spCrudObj.updateData("HealthCheckupService", this.state.CHSApproverView.ID, CHSRequestItem, this.props).then(async (req) => {
        alert('CHS Request Rejected Successfully!');
        this.closeDialog();
       // window.location.href = "https://sharepointwebssse.eximbankindia.in/sites/hrm/SitePages/CHSModule.aspx";  // Update the URL path to your dashboard route
        window.location.href = `${ENV_CONFIG.siteUrl}/SitePages/CHSModule.aspx`;  // Update the URL path to your dashboard route
  
        return req;
      });

    }

    // if( this.state.ExpenseDetails.FinalAmount==0 &&  !this.state.ExpenseDetails){
    //   alert('Please Enter Final Amount ')
    //   return false;
    // }

 
 
  };
  async uploadPRDoc(ListName, itemId, files) {
    const spCrudObj = await useSPCRUD();
    for (let i = 0; i < files.length; i++) {
      let file = files[i];
      let fileName = file.name;
      try {
        const brPlanFile = await spCrudObj.addAttchmentInList(
          file,
          ListName,
          itemId,
          fileName,
          this.props
        );
        console.log(`Uploaded: ${fileName}`, brPlanFile);
      } catch (error) {
        console.log(`Error uploading ${fileName}:`, error);
      }
    }
  }
  public UserPendingDashboard = async () => {
    return await EmployeeOps().getUserDashboard(this.props).then(UserPending => {
      this.setState({ UserDashboard: UserPending });
      return UserPending;
    });
  };
  public UserApprovedDashboards = async () => {
    return await EmployeeOps().getUserApprovedDashboard(this.props).then(UserApproved => {
      this.setState({ userApprovedDashboard: UserApproved });
      return UserApproved;
    });
  };
  public UserRejectedDashboards = async () => {
    return await EmployeeOps().getUserRejectedDashboard(this.props).then(UserRejected => {
      this.setState({ userRejectedDashboard: UserRejected });
      return UserRejected;
    });
  };
  public HR1ApprovePendingDashboard = async () => {
    return await EmployeeOps().HR1getApproveDashboard(this.props).then(UserPending => {
      this.setState({ HR1ApprPendingDashboard: UserPending, allDashboardDataHR1Pending: UserPending });
      return UserPending;
    });
  };


  // allDashboardDataHR1Pending:[],

  // allDashboardDataHR1Rejected:[],

  // allDashboardDataHR1Approved:[],

  public HR1ApproveApprovedDashboards = async () => {
    return await EmployeeOps().HR1getApproveApprovedDashboard(this.props).then(UserApproved => {
      this.setState({ HR1ApproverApprDashboard: UserApproved, allDashboardDataHR1Approved: UserApproved });
      return UserApproved;
    });
  };
  public HR1ApproveRejectedDashboards = async () => {
    return await EmployeeOps().HR1getApproveRejectedDashboard(this.props).then(UserRejected => {
      this.setState({ HR1ApproverRejectDashboard: UserRejected, allDashboardDataHR1Rejected: UserRejected });
      return UserRejected;
    });
  };
  public HR2ApprovePendingDashboard = async () => {
    return await EmployeeOps().HR2getApproveDashboard(this.props).then(UserPending => {
      this.setState({ HR2ApprPendingDashboard: UserPending, allDashboardDataHR2Pending: UserPending });
      return UserPending;
    });
  };
  public HR2ApproveApprovedDashboards = async () => {
    return await EmployeeOps().HR2getApproveApprovedDashboard(this.props).then(UserApproved => {
      this.setState({ HR2ApproverApprDashboard: UserApproved, allDashboardDataHR2Approved: UserApproved });
      return UserApproved;
    });
  };
  public HR2ApproveRejectedDashboards = async () => {
    return await EmployeeOps().HR2getApproveRejectedDashboard(this.props).then(UserRejected => {
      this.setState({ HR2ApproverRejectDashboard: UserRejected, allDashboardDataHR2Rejected: UserRejected });
      return UserRejected;
    });
  };
  public getHR1Approver = async (Items) => {
    console.log(Items);
    const ApproverViewReqItems = Items;
    var NewTotal = 0;
    this.setState({
      CHSApproverView: ApproverViewReqItems,
      isDialogHR1: true
    })
  }
  public getHR2Approver = async (Items) => {
    console.log(Items);
    const ApproverViewReqItems = Items;
    var NewTotal = 0;
    this.setState({
      CHSApproverView: ApproverViewReqItems,
      isDialogHR2: true
    })
  }
  public getHR1ApproverView = async (Items) => {
    console.log(Items);
    const ApproverViewReqItems = Items;
    var NewTotal = 0;
    this.setState({
      CHSApproverView: ApproverViewReqItems,
      isDialogViewHR1: true
    })
  }
  public getHR2ApproverView = async (Items) => {
    console.log(Items);
    const ApproverViewReqItems = Items;
    var NewTotal = 0;
    this.setState({
      CHSApproverView: ApproverViewReqItems,
      isDialogViewHR2: true
    })
  }


  //   // Dropdown change handler
  // handleDropdownChangeDash = (event, option) => {
  //   this.setState({ DependentType: event.key }, () => {
  //     this.filterDashboardData();
  //   });
  // };

  // // Additional function that runs after state updates
  // filterDashboardData = () => {
  //   console.log("Filtering data for:", this.state.DependentType);

  //   const filteredData =this.state.HR2ApproverApprDashboard.filter(
  //           (item) => item.EmployeeType === this.state.DependentType
  //         );

  //   this.setState({ HR2ApproverApprDashboard: filteredData });
  // };



  // handleDropdownChangeDashPendingHR2
  // handleDropdownChangeDashApprovedHR2
  // handleDropdownChangeDashRejectedHR2


  handleDropdownChangeDashHR2Approved = (event, option) => {
    const selectedType = event.key;
    this.setState({ DependentType: selectedType }, () => {
      // this.filterDashboardData();

      const { allDashboardDataHR2Approved, DependentType } = this.state;

      const filteredData = DependentType === "ALL" ? allDashboardDataHR2Approved : allDashboardDataHR2Approved.filter((item) => item.EmployeeType === DependentType);

      this.setState({
        HR2ApproverApprDashboard: filteredData
        // HR1ApproverApprDashboard: filteredData,

        // UserDashboard: filteredData,
        // userApprovedDashboard: filteredData,
        // userRejectedDashboard: filteredData,
        // HR1ApprPendingDashboard: filteredData,
        // HR1ApproverRejectDashboard: filteredData,
        // HR2ApprPendingDashboard: filteredData,
        // HR2ApproverRejectDashboard: filteredData
      });

    });
  };

  handleDropdownChangeDashHR2Rejected = (event, option) => {
    const selectedType = event.key;
    this.setState({ DependentType: selectedType }, () => {
      // this.filterDashboardData();

      const { allDashboardDataHR2Rejected, DependentType } = this.state;

      const filteredData = DependentType === "ALL" ? allDashboardDataHR2Rejected : allDashboardDataHR2Rejected.filter((item) => item.EmployeeType === DependentType);

      this.setState({
        HR2ApproverRejectDashboard: filteredData
        // HR1ApproverApprDashboard: filteredData,

        // UserDashboard: filteredData,
        // userApprovedDashboard: filteredData,
        // userRejectedDashboard: filteredData,
        // HR1ApprPendingDashboard: filteredData,
        // HR1ApproverRejectDashboard: filteredData,
        // HR2ApprPendingDashboard: filteredData,
        // HR2ApproverRejectDashboard: filteredData
      });

    });
  };

  handleDropdownChangeDashHR2Pending = (event, option) => {
    const selectedType = event.key;
    this.setState({ DependentType: selectedType }, () => {
      // this.filterDashboardData();

      const { allDashboardDataHR2Pending, DependentType } = this.state;

      const filteredData = DependentType === "ALL" ? allDashboardDataHR2Pending : allDashboardDataHR2Pending.filter((item) => item.EmployeeType === DependentType);

      this.setState({
        HR2ApprPendingDashboard: filteredData
        // HR1ApproverApprDashboard: filteredData,

        // UserDashboard: filteredData,
        // userApprovedDashboard: filteredData,
        // userRejectedDashboard: filteredData,
        // HR1ApprPendingDashboard: filteredData,
        // HR1ApproverRejectDashboard: filteredData,
        // HR2ApprPendingDashboard: filteredData,
        // HR2ApproverRejectDashboard: filteredData
      });

    });
  };


  handleDropdownChangeDashHR1Approved = (event, option) => {
    const selectedType = event.key;
    this.setState({ DependentType: selectedType }, () => {
      // this.filterDashboardData();

      const { allDashboardDataHR1Approved, DependentType } = this.state;

      const filteredData = DependentType === "ALL" ? allDashboardDataHR1Approved : allDashboardDataHR1Approved.filter((item) => item.EmployeeType === DependentType);

      this.setState({
        HR1ApproverApprDashboard: filteredData
        // HR1ApproverApprDashboard: filteredData,

        // UserDashboard: filteredData,
        // userApprovedDashboard: filteredData,
        // userRejectedDashboard: filteredData,
        // HR1ApprPendingDashboard: filteredData,
        // HR1ApproverRejectDashboard: filteredData,
        // HR2ApprPendingDashboard: filteredData,
        // HR2ApproverRejectDashboard: filteredData
      });

    });
  };

  handleDropdownChangeDashHR1Rejected = (event, option) => {
    const selectedType = event.key;
    this.setState({ DependentType: selectedType }, () => {
      // this.filterDashboardData();

      const { allDashboardDataHR1Rejected, DependentType } = this.state;

      const filteredData = DependentType === "ALL" ? allDashboardDataHR1Rejected : allDashboardDataHR1Rejected.filter((item) => item.EmployeeType === DependentType);

      this.setState({
        HR1ApproverRejectDashboard: filteredData
        // HR1ApproverApprDashboard: filteredData,

        // UserDashboard: filteredData,
        // userApprovedDashboard: filteredData,
        // userRejectedDashboard: filteredData,
        // HR1ApprPendingDashboard: filteredData,
        // HR1ApproverRejectDashboard: filteredData,
        // HR2ApprPendingDashboard: filteredData,
        // HR2ApproverRejectDashboard: filteredData
      });

    });
  };

  handleDropdownChangeDashHR1Pending = (event, option) => {
    const selectedType = event.key;
    this.setState({ DependentType: selectedType }, () => {
      // this.filterDashboardData();

      const { allDashboardDataHR1Pending, DependentType } = this.state;

      const filteredData = DependentType === "ALL" ? allDashboardDataHR1Pending : allDashboardDataHR1Pending.filter((item) => item.EmployeeType === DependentType);

      this.setState({
        HR1ApprPendingDashboard: filteredData
        // HR1ApproverApprDashboard: filteredData,

        // UserDashboard: filteredData,
        // userApprovedDashboard: filteredData,
        // userRejectedDashboard: filteredData,
        // HR1ApprPendingDashboard: filteredData,
        // HR1ApproverRejectDashboard: filteredData,
        // HR2ApprPendingDashboard: filteredData,
        // HR2ApproverRejectDashboard: filteredData
      });

    });
  };


  // filterDashboardData = () => {
  //   const { allDashboardData, DependentType } = this.state;

  //   const filteredData =
  //     DependentType === "ALL"
  //       ? allDashboardData
  //       : allDashboardData.filter((item) => item.EmployeeType === DependentType);

  //   this.setState({ HR1ApproverApprDashboard: filteredData 
  //     // HR1ApproverApprDashboard: filteredData,

  //     // UserDashboard: filteredData,
  //     // userApprovedDashboard: filteredData,
  //     // userRejectedDashboard: filteredData,
  //     // HR1ApprPendingDashboard: filteredData,
  //     // HR1ApproverRejectDashboard: filteredData,
  //     // HR2ApprPendingDashboard: filteredData,
  //     // HR2ApproverRejectDashboard: filteredData
  //   });
  // };


  onSearchChange = (newValue: string) => {
    const filtered = this.state.filteredOptions.filter((option) =>
      option.text.toLowerCase().includes(newValue.toLowerCase())
    );
    this.setState({
      searchValue: newValue,
      AllEmployeeCollObj: filtered,
      isDropdownOpen: !!filtered.length, //  Open dropdown if there are results

    });
  };

  // handleCheckboxChange = (
  //   event?: React.FormEvent<HTMLElement | HTMLInputElement>,
  //   checked?: boolean
  // ): void => {
  //   this.setState({ isChecked: !!checked });
  // };

  handleCheckboxChange = (option: string) => (
    event?: React.FormEvent<HTMLElement | HTMLInputElement>,
    checked?: boolean
  ): void => {
    if (checked) {
      this.setState({ selectedOptionCHBx: option });
    } else {
      this.setState({ selectedOptionCHBx: null });
    }
  };
  private openHR2InnerTab = (tabName: string): void => {
    // localStorage.setItem("activeHR2Tab", tabName);
    // this.setState({ activeHR2Tab: tabName });
  
  };
  private openHR1InnerTab = (tabName: string): void => {
    // localStorage.setItem("activeHR1Tab", tabName);
    // this.setState({ activeHR1Tab: tabName });
   
  };
  private openHRInnerTab = (tabName: string): void => {
    this.setState({ activeHR1Tab: tabName });
  localStorage.setItem("activeHR1Tab", tabName);
 
  };

  public render(): React.ReactElement<IChsModuleProps> {
    const { selectedOption } = this.state;
    const value = selectedOption;
    // const valueq = selectedOptionCHBx ;  

    const { activeTab } = this.state;
   
    return ( 
      
      <div className={styles.pettyCash} >
        <div className="" style={{border:'1px solid #ddd',  borderRadius:'4px'}}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}> </span>
              <div className='mb-2 text-right'><DefaultButton className='btn-primary' onClick={() => this.CreateRequest()}> <Icon iconName='' ></Icon> Create Request</DefaultButton></div>
             
              
              <Pivot linkSize={PivotLinkSize.large} linkFormat={PivotLinkFormat.tabs}>
               
                
<PivotItem linkText="User Dashboard" hidden={!this.state.ShowHRTab} className="tab-box"  onClick={() => this.handleOuterTabClick('UserDashboard')}>
  <div className="row">
    <div className={`${styles.tabnav} col-md-2`}>
      <button
        className={`tablink ${this.state.activeHR1Tab === 'Pending' ? 'active' : ''}`}
        onClick={() => this.openHRInnerTab('Pending')}
      >
        Pending
      </button>
      <button
        className={`tablink ${this.state.activeHR1Tab === 'Approved' ? 'active' : ''}`}
        onClick={() => this.openHRInnerTab('Approved')}
      >
        Approved
      </button>
      <button
        className={`tablink ${this.state.activeHR1Tab === 'Rejected' ? 'active' : ''}`}
        onClick={() => this.openHRInnerTab('Rejected')}
      >
        Rejected 
      </button>
    </div> 

    <div className="col-md-10 panelbodybox">
      {/* PENDING TAB */}
      <div className={`tabcontent ${this.state.activeHR1Tab === 'Pending' ? 'active' : ''}`} id="Pending">
      <table className="table ">
                            <tr>
                              <th>View</th>
                              <th>CHS ID</th>
                              <th>EmployeeID</th>
                              <th>EmployeeName</th>
                              <th>Employee Type</th>
                              <th>Date of Birth</th>
                              <th>Dependent Type</th>
                              <th>Claimed Ammount</th>
                              <th>Final Approved Ammount</th>
                              <th>Status</th>
                              {/* <th>View Doc.</th> */}
                            </tr>
                            {
                              this.state.UserDashboard.length > 0 ? this.state.UserDashboard.map((items) => {
                                return (
                                  <tr>
                                    <td><Icon iconName='View' onClick={() => this.getHR1ApproverView(items)} title='View' className={styles.iconcolor}></Icon></td>
                                    <td>{items.Title}</td>
                                    <td>{items.EmployeeID}</td>
                                    <td>{items.EmployeeName}</td>
                                    <td>{items.EmployeeType}</td>
                                    <td>{moment(items.DateofBirth).format("DD/MM/YYYY")}</td>
                                    <td>{items.DependentType}</td>
                                    <td>{items.AmountClaimed}</td>
                                    <td>{items.HRApprovedAmount}</td>
                                    <td>{items.Status}</td>

                                  </tr>
                                )
                              })
                                : ""
                            }
                          </table>
      </div>

      {/* APPROVED TAB */}
      <div className={` tabcontent ${this.state.activeHR1Tab === 'Approved' ? 'active' : ''}`} id="Approved">
      <table className="table ">
                            <tr>
                              <th>View</th>
                              <th>CHS ID</th>
                              <th>EmployeeID</th>
                              <th>EmployeeName</th>
                              <th>Employee Type</th>
                              <th>Date of Birth</th>
                              <th>Dependent Type</th>
                              <th>Claimed Ammount</th>
                              <th>Final Approved Ammount</th>
                              <th>Status</th>
                              {/* <th>View Doc.</th> */}

                            </tr>
                            {
                              this.state.userApprovedDashboard.length > 0 ? this.state.userApprovedDashboard.map((items) => {
                                return (
                                  <tr>
                                    <td><Icon iconName='View' onClick={() => this.getHR1ApproverView(items)} title='View' className={styles.iconcolor}></Icon></td>
                                    <td>{items.Title}</td>
                                    <td>{items.EmployeeID}</td>
                                    <td>{items.EmployeeName}</td>
                                    <td>{items.EmployeeType}</td>
                                    <td>{moment(items.DateofBirth).format("DD/MM/YYYY")}</td>
                                    <td>{items.DependentType}</td>
                                    <td>{items.AmountClaimed}</td>
                                    <td>{items.HRApprovedAmount}</td>
                                    <td>{items.Status}</td>
                                    {/* {
                                      items !== undefined && items !== null && items !== "" && items.AttachmentFiles && items.AttachmentFiles.length > 0 ? (
                                        items.AttachmentFiles.map((files) => (
                                          <li style={{ listStyle: 'decimal', color: '#428bca' }}><a href={files.ServerRelativeUrl} target='_blank'>{files.FileName}</a></li>
                                        ))
                                      ) : (
                                        <td>No Attachments</td>
                                      )
                                    } */}
                                  </tr>
                                )
                              })
                                : ""
                            }
                          </table>
      </div>

      {/* REJECTED TAB */}
      <div className={`tabcontent ${this.state.activeHR1Tab === 'Rejected' ? 'active' : ''}`} id="Rejected">
      <table className="table ">
                            <tr>
                              <th>View</th>
                              <th>CHS ID</th>
                              <th>EmployeeID</th>
                              <th>EmployeeName</th>
                              <th>Employee Type</th>
                              <th>Date of Birth</th>
                              <th>Dependent Type</th>
                              <th>Claimed Ammount</th>
                              <th>Final Approved Ammount</th>
                              <th>Status</th>
                              {/* <th>View Doc.</th> */}

                            </tr>
                            {
                              this.state.userRejectedDashboard.length > 0 ? this.state.userRejectedDashboard.map((items) => {
                                return (
                                  <tr>
                                    <td><Icon iconName='View' onClick={() => this.getHR1ApproverView(items)} title='View' className={styles.iconcolor}></Icon></td>
                                    <td>{items.Title}</td>
                                    <td>{items.EmployeeID}</td>
                                    <td>{items.EmployeeName}</td>
                                    <td>{items.EmployeeType}</td>
                                    <td>{moment(items.DateofBirth).format("DD/MM/YYYY")}</td>
                                    <td>{items.DependentType}</td>
                                    <td>{items.AmountClaimed}</td>
                                    <td>{items.HRApprovedAmount}</td>
                                    <td>{items.Status}</td>
                                    {/* {
                                      items !== undefined && items !== null && items !== "" && items.AttachmentFiles && items.AttachmentFiles.length > 0 ? (
                                        items.AttachmentFiles.map((files) => (
                                          <li style={{ listStyle: 'decimal', color: '#428bca' }}><a href={files.ServerRelativeUrl} target='_blank'>{files.FileName}</a></li>
                                        ))
                                      ) : (
                                        <td>No Attachments</td>
                                      )
                                    } */}
                                  </tr>
                                )
                              })
                                : ""
                            }
                          </table>
      </div>
    </div>
  </div>   
</PivotItem>
               
                <PivotItem linkText="HR1 Approver Dashboard" hidden={!this.state.ShowHR1Tab} className="tab-box"  onClick={() => this.handleOuterTabClick('HR1Dashboard')}>
  <div className="row">
    <div className={`${styles.tabnav} col-md-2`}>
      <button
        className={`tablink ${this.state.activeHR1Tab === 'Pending' ? 'active' : ''}`}
        onClick={() => this.openHR1InnerTab('Pending')}
      >
        Pending
      </button>
      <button
        className={`tablink ${this.state.activeHR1Tab === 'Approved' ? 'active' : ''}`}
        onClick={() => this.openHR1InnerTab('Approved')}
      >
        Approved
      </button>
      <button
        className={`tablink ${this.state.activeHR1Tab === 'Rejected' ? 'active' : ''}`}
        onClick={() => this.openHR1InnerTab('Rejected')}
      >
        Rejected 
      </button>
    </div> 

    <div className="col-md-10 panelbodybox">
      {/* PENDING TAB */}
      <div className={`tabcontent ${this.state.activeHR1Tab === 'Pending' ? 'active' : ''}`} id="Pending">
        
      <div className="col-md-5 plr-5">
                            <div className='row'>
                            <div className="col-md-4  pl-0" style={{paddingTop:'6px'}}>
                            <Label className="control-Label font-weight-bold" style={{display:'inline-block'}}>Employee Type </Label>
                            </div>
                            <div className="col-md-8">
                            <Dropdown
                              placeHolder="Select Employee Type"
                              options={[
                                { key: "ALL", text: "ALL" },
                                { key: "PERMANENT", text: "PERMANENT" },
                                { key: "RETIRED", text: "RETIRED" },
                              ]}
                              selectedKey={this.state.DependentType}
                              onChanged={this.handleDropdownChangeDashHR1Pending}
                              className="dropdown-style "
                            />
                            </div>
                            </div>
                           
                            {/* <Dropdown
                              placeHolder="Select Dependent"
                              // disabled={true}
                              options={[{ key: 'ALL', text: 'ALL' },{ key: 'PERMANENT', text: 'PERMANENT' }, { key: 'RETIRED', text: 'RETIRED' }]}
                              onChanged={(e, option) => this.setState({ DependentType: e.key })}
                              className="dropdown-style"
                            /> */}

                            {/* <Dropdown
                              placeHolder="Select Employee Type"
                              options={[
                                { key: "ALL", text: "ALL" },
                                { key: "PERMANENT", text: "PERMANENT" },
                                { key: "RETIRED", text: "RETIRED" },
                              ]}
                              selectedKey={this.state.DependentType}
                              onChanged={this.handleDropdownChangeDashHR1Pending}
                              className="dropdown-style"
                            /> */}
  
                          </div>

                          <table className="table HAD">
                            <tr>
                              <th>View</th>
                              <th>Action</th>
                              <th>CHS ID</th>
                              <th>EmployeeID</th>
                              <th>EmployeeName</th>
                              <th>Employee Type</th>
                              <th>Date of Birth</th>
                              <th>Dependent Type</th>
                              <th>Claimed Ammount</th>
                              <th>Final Approved Ammount</th>
                              <th>Status</th>
                              {/* <th>View Doc.</th> */}

                            </tr>
                            {
                              this.state.HR1ApprPendingDashboard.length > 0 ? this.state.HR1ApprPendingDashboard.map((items) => {
                                return (
                                  <tr>
                                    <td><Icon iconName='View' onClick={() => this.getHR1ApproverView(items)} title='View' className={styles.iconcolor}></Icon></td>
                                    <td><Icon iconName='CheckMark' onClick={() => this.getHR1Approver(items)} className={styles.iconcolor} title='Approver'></Icon></td>
                                    <td>{items.Title}</td>
                                    <td>{items.EmployeeID}</td>
                                    <td>{items.EmployeeName}</td>
                                    <td>{items.EmployeeType}</td>
                                    <td>{moment(items.DateofBirth).format("DD/MM/YYYY")}</td>
                                    <td>{items.DependentType}</td>
                                    <td>{items.AmountClaimed}</td>
                                    <td>{items.HRApprovedAmount}</td>
                                    <td>{items.Status}</td>
                                    {/* {
                                      items !== undefined && items !== null && items !== "" && items.AttachmentFiles && items.AttachmentFiles.length > 0 ? (
                                        items.AttachmentFiles.map((files) => (
                                          <li style={{ listStyle: 'decimal', color: '#428bca' }}><a href={files.ServerRelativeUrl} target='_blank'>{files.FileName}</a></li>
                                        ))
                                      ) : (
                                        <td>No Attachments</td>
                                      )
                                    } */}
                                  </tr>
                                )
                              })
                                : ""
                            }
                          </table>
      </div>

      {/* APPROVED TAB */}
      <div className={` tabcontent ${this.state.activeHR1Tab === 'Approved' ? 'active' : ''}`} id="Approved">
      <div className="col-md-5 plr-5">
                          <div className='row'>
                          <div className="col-md-4  pl-0"  style={{paddingTop:'6px'}}>
                          <Label className="control-Label font-weight-bold">Employee Type </Label>
                          </div>
                          <div className="col-md-8 ">
                          <Dropdown
                              placeHolder="Select Employee Type"
                              options={[
                                { key: "ALL", text: "ALL" },
                                { key: "PERMANENT", text: "PERMANENT" },
                                { key: "RETIRED", text: "RETIRED" },
                              ]}
                              selectedKey={this.state.DependentType}
                              onChanged={this.handleDropdownChangeDashHR1Approved}
                              className="dropdown-style"
                            />
                          </div>

                          </div>
                            
                            {/* <Dropdown
                              placeHolder="Select Dependent"
                              // disabled={true}
                              options={[{ key: 'ALL', text: 'ALL' },{ key: 'PERMANENT', text: 'PERMANENT' }, { key: 'RETIRED', text: 'RETIRED' }]}
                              onChanged={(e, option) => this.setState({ DependentType: e.key })}
                              className="dropdown-style"
                            /> */}

                            {/* <Dropdown
                              placeHolder="Select Employee Type"
                              options={[
                                { key: "ALL", text: "ALL" },
                                { key: "PERMANENT", text: "PERMANENT" },
                                { key: "RETIRED", text: "RETIRED" },
                              ]}
                              selectedKey={this.state.DependentType}
                              onChanged={this.handleDropdownChangeDashHR1Approved}
                              className="dropdown-style"
                            /> */}

                          </div>
                          <table className="table ">
                            <tr>
                              <th>View</th>
                              <th>CHS ID</th>
                              <th>EmployeeID</th>
                              <th>EmployeeName</th>
                              <th>Employee Type</th>
                              <th>Date of Birth</th>
                              <th>Dependent Type</th>
                              <th>Claimed Ammount</th>
                              <th>Final Approved Ammount</th>
                              <th>Status</th>
                              {/* <th>View Doc.</th> */}

                            </tr>
                            {
                              this.state.HR1ApproverApprDashboard.length > 0 ? this.state.HR1ApproverApprDashboard.map((items) => {
                                return (
                                  <tr>
                                    <td><Icon iconName='View' onClick={() => this.getHR1ApproverView(items)} title='View' className={styles.iconcolor}></Icon></td>
                                    <td>{items.Title}</td>
                                    <td>{items.EmployeeID}</td>
                                    <td>{items.EmployeeName}</td>
                                    <td>{items.EmployeeType}</td>
                                    <td>{moment(items.DateofBirth).format("DD/MM/YYYY")}</td>
                                    <td>{items.DependentType}</td>
                                    <td>{items.AmountClaimed}</td>
                                    <td>{items.HRApprovedAmount}</td>
                                    <td>{items.Status}</td>
                                    {/* {
                                      items !== undefined && items !== null && items !== "" && items.AttachmentFiles && items.AttachmentFiles.length > 0 ? (
                                        items.AttachmentFiles.map((files) => (
                                          <li style={{ listStyle: 'decimal', color: '#428bca' }}><a href={files.ServerRelativeUrl} target='_blank'>{files.FileName}</a></li>
                                        ))
                                      ) : (
                                        <td>No Attachments</td>
                                      )
                                    } */}
                                  </tr>
                                )
                              })
                                : ""
                            }
                          </table>
      </div>

      {/* REJECTED TAB */}
      <div className={`tabcontent ${this.state.activeHR1Tab === 'Rejected' ? 'active' : ''}`} id="Rejected">
      <div className="col-md-5 plr-5 ">
                            <div className="row">
                            <div className="col-md-4  pl-0"  style={{paddingTop:'6px'}}>
                               <Label className="control-Label font-weight-bold">Employee Type </Label>
                            </div>
                            <div className="col-md-8">
                            <Dropdown
                              placeHolder="Select Employee Type"
                              options={[
                                { key: "ALL", text: "ALL" },
                                { key: "PERMANENT", text: "PERMANENT" },
                                { key: "RETIRED", text: "RETIRED" },
                              ]}
                              selectedKey={this.state.DependentType}
                              onChanged={this.handleDropdownChangeDashHR1Rejected}
                              className="dropdown-style"
                            /> 
                            </div>
                            </div>
                           
                            {/* <Dropdown
                              placeHolder="Select Dependent"
                              // disabled={true}
                              options={[{ key: 'ALL', text: 'ALL' },{ key: 'PERMANENT', text: 'PERMANENT' }, { key: 'RETIRED', text: 'RETIRED' }]}
                              onChanged={(e, option) => this.setState({ DependentType: e.key })}
                              className="dropdown-style"
                            /> */}

                            {/* <Dropdown
                              placeHolder="Select Employee Type"
                              options={[
                                { key: "ALL", text: "ALL" },
                                { key: "PERMANENT", text: "PERMANENT" },
                                { key: "RETIRED", text: "RETIRED" },
                              ]}
                              selectedKey={this.state.DependentType}
                              onChanged={this.handleDropdownChangeDashHR1Rejected}
                              className="dropdown-style"
                            /> */}

                          </div>

                          <table className="table ">
                            <tr>
                              <th>View</th>
                              <th>CHS ID</th>
                              <th>EmployeeID</th>
                              <th>EmployeeName</th>
                              <th>Employee Type</th>
                              <th>Date of Birth</th>
                              <th>Dependent Type</th>
                              <th>Claimed Ammount</th>
                              <th>Final Approved Ammount</th>
                              <th>Status</th>
                              {/* <th>View Doc.</th> */}

                            </tr>
                            {
                              this.state.HR1ApproverRejectDashboard.length > 0 ? this.state.HR1ApproverRejectDashboard.map((items) => {
                                return (
                                  <tr>
                                    <td><Icon iconName='View' onClick={() => this.getHR1ApproverView(items)} title='View' className={styles.iconcolor}></Icon></td>
                                    <td>{items.Title}</td>
                                    <td>{items.EmployeeID}</td>
                                    <td>{items.EmployeeName}</td>
                                    <td>{items.EmployeeType}</td>
                                    <td>{moment(items.DateofBirth).format("DD/MM/YYYY")}</td>
                                    <td>{items.DependentType}</td>
                                    <td>{items.AmountClaimed}</td>
                                    <td>{items.HRApprovedAmount}</td>
                                    <td>{items.Status}</td>
                                    {/* {
                                      items !== undefined && items !== null && items !== "" && items.AttachmentFiles && items.AttachmentFiles.length > 0 ? (
                                        items.AttachmentFiles.map((files) => (
                                          <li style={{ listStyle: 'decimal', color: '#428bca' }}><a href={files.ServerRelativeUrl} target='_blank'>{files.FileName}</a></li>
                                        ))
                                      ) : (
                                        <td>No Attachments</td>
                                      )
                                    } */}
                                  </tr>
                                )
                              })
                                : ""
                            }
                          </table>
      </div>
    </div>
  </div>   
</PivotItem>
                {/* new code added for tabing */}
                <PivotItem linkText="HR2 Approver Dashboard" hidden={!this.state.ShowHR2Tab} className="tab-box"  onClick={() => this.handleOuterTabClick('HR2Dashboard')}>
               
  <div className="row">
    <div className={`${styles.tabnav} col-md-2`}>
      <button
        className={`tablink ${this.state.activeHR2Tab === 'Pending' ? 'active' : ''}`}
        onClick={() => this.openHR2InnerTab('Pending')}
      >
        Pending
      </button>
      <button
        className={`tablink ${this.state.activeHR2Tab === 'Approved' ? 'active' : ''}`}
        onClick={() => this.openHR2InnerTab('Approved')}
      >
        Approved
      </button>
      <button
        className={`tablink ${this.state.activeHR2Tab === 'Rejected' ? 'active' : ''}`}
        onClick={() => this.openHR2InnerTab('Rejected')}
      >
        Rejected
      </button>
    </div>

    <div className="col-md-10 panelbodybox">
      {/* PENDING TAB */}
      <div className={`tabcontent ${this.state.activeHR2Tab === 'Pending' ? 'active' : ''}`} id="Pending">
      <div className="col-md-5 plr-5">
                          <div className="row">
                          <div className="col-md-4  pl-0" style={{paddingTop:'6px'}}>
                            <Label className="control-Label font-weight-bold" style={{display:'inline-block'}}>Employee Type </Label>
                            </div>
                            <div className="col-md-8">
                            <Dropdown 
                              placeHolder="Select Employee Type"
                              options={[
                                { key: "ALL", text: "ALL" },
                                { key: "PERMANENT", text: "PERMANENT" },
                                { key: "RETIRED", text: "RETIRED" },
                              ]}
                              selectedKey={this.state.DependentType}
                              onChanged={this.handleDropdownChangeDashHR2Pending}
                              className="dropdown-style"
                            />
                            </div>

                          </div>
                          </div>
                         
                          <div className="col-md-3"  style={{paddingTop:'6px'}}>
                           
                            {/* <Dropdown
                              placeHolder="Select Dependent"
                              // disabled={true}
                              options={[{ key: 'ALL', text: 'ALL' },{ key: 'PERMANENT', text: 'PERMANENT' }, { key: 'RETIRED', text: 'RETIRED' }]}
                              onChanged={(e, option) => this.setState({ DependentType: e.key })}
                              className="dropdown-style"
                            /> */}



                           

                          </div>
                          <table className="table ">
                            <tr>
                              <th>View</th>
                              <th>Action</th>
                              <th>CHS ID</th>
                              <th>EmployeeID</th>
                              <th>EmployeeName</th>
                              <th>Employee Type</th>
                              <th>Date of Birth</th>
                              <th>Dependent Type</th>
                              <th>Claimed Ammount</th>
                              <th>Final Approved Ammount</th>
                              <th>Status</th>
                              {/* <th>View Doc.</th> */}

                            </tr>
                            {
                              this.state.HR2ApprPendingDashboard.length > 0 ? this.state.HR2ApprPendingDashboard.map((items) => {
                                return (
                                  <tr>
                                    <td><Icon iconName='View' onClick={() => this.getHR2ApproverView(items)} title='View' className={styles.iconcolor}></Icon></td>
                                    <td><Icon iconName='CheckMark' onClick={() => this.getHR2Approver(items)} className={styles.iconcolor} title='Approver'></Icon></td>
                                    <td>{items.Title}</td>
                                    <td>{items.EmployeeID}</td>
                                    <td>{items.EmployeeName}</td>
                                    <td>{items.EmployeeType}</td>
                                    <td>{moment(items.DateofBirth).format("DD/MM/YYYY")}</td>
                                    <td>{items.DependentType}</td>
                                    <td>{items.AmountClaimed}</td>
                                    <td>{items.HRApprovedAmount}</td>
                                    <td>{items.Status}</td>

                                  </tr>
                                )
                              })
                                : ""
                            }
                          </table>
      </div>

      {/* APPROVED TAB */}
      <div className={` tabcontent ${this.state.activeHR2Tab === 'Approved' ? 'active' : ''}`} id="Approved">
      <div className="col-md-5 plr-5">
                            <div className="row">
                            <div className="col-md-4  pl-0" style={{paddingTop:'6px'}}>
                            <Label className="control-Label font-weight-bold" style={{display:'inline-block'}}>Employee Type </Label>
                            </div>
                            <div className="col-md-8">
                            <Dropdown
                              placeHolder="Select Employee Type"
                              options={[
                                { key: "ALL", text: "ALL" },
                                { key: "PERMANENT", text: "PERMANENT" },
                                { key: "RETIRED", text: "RETIRED" },
                              ]}
                              selectedKey={this.state.DependentType}
                              onChanged={this.handleDropdownChangeDashHR2Approved}
                              className="dropdown-style"
                            />
                                  </div>
                            </div>
                          </div>
                          {/* <div className="col-md-3"  style={{paddingTop:'6px'}}>
                            <Label className="control-Label font-weight-bold">Employee Type </Label>


                            <Dropdown
                              placeHolder="Select Employee Type"
                              options={[
                                { key: "ALL", text: "ALL" },
                                { key: "PERMANENT", text: "PERMANENT" },
                                { key: "RETIRED", text: "RETIRED" },
                              ]}
                              selectedKey={this.state.DependentType}
                              onChanged={this.handleDropdownChangeDashHR2Approved}
                              className="dropdown-style"
                            />

                          </div> */}




                          <table className="table ">
                            <tr>
                              <th>View</th>
                              <th>CHS ID</th>
                              <th>EmployeeID</th>
                              <th>EmployeeName</th>
                              <th>Employee Type</th>
                              <th>Date of Birth</th>
                              <th>Dependent Type</th>
                              <th>Claimed Ammount</th>
                              <th>Final Approved Ammount</th>
                              <th>Status</th>
                              {/* <th>View Doc.</th> */}

                            </tr>
                            {
                              this.state.HR2ApproverApprDashboard.length > 0 ? this.state.HR2ApproverApprDashboard.map((items) => {
                                return (
                                  <tr>
                                    <td><Icon iconName='View' onClick={() => this.getHR2ApproverView(items)} title='View' className={styles.iconcolor}></Icon></td>
                                    <td>{items.Title}</td>
                                    <td>{items.EmployeeID}</td>
                                    <td>{items.EmployeeName}</td>
                                    <td>{items.EmployeeType}</td>
                                    <td>{moment(items.DateofBirth).format("DD/MM/YYYY")}</td>
                                    <td>{items.DependentType}</td>
                                    <td>{items.AmountClaimed}</td>
                                    <td>{items.HRApprovedAmount}</td>
                                    <td>{items.Status}</td>

                                  </tr>
                                )
                              })
                                : ""
                            }
                          </table>
      </div>

      {/* REJECTED TAB */}
      <div className={`tabcontent ${this.state.activeHR2Tab === 'Rejected' ? 'active' : ''}`} id="Rejected">
      <div className="col-md-5 plr-5">
                          <div className='row'>
                          <div className="col-md-4  pl-0" style={{paddingTop:'6px'}}>
                            <Label className="control-Label font-weight-bold" style={{display:'inline-block'}}>Employee Type </Label>
                            </div>
                            <div className="col-md-8">
                            <Dropdown
                              placeHolder="Select Employee Type"
                              options={[
                                { key: "ALL", text: "ALL" },
                                { key: "PERMANENT", text: "PERMANENT" },
                                { key: "RETIRED", text: "RETIRED" },
                              ]}
                              selectedKey={this.state.DependentType}
                              onChanged={this.handleDropdownChangeDashHR2Rejected}
                              className="dropdown-style"
                            />
                            </div>
                          </div>
                          </div>
                          
                          <table className="table ">
                            <tr>
                              <th>View</th>
                              <th>CHS ID</th>
                              <th>EmployeeID</th>
                              <th>EmployeeName</th>
                              <th>Employee Type</th>
                              <th>Date of Birth</th>
                              <th>Dependent Type</th>
                              <th>Claimed Ammount</th>
                              <th>Final Approved Ammount</th>
                              <th>Status</th>
                              {/* <th>View Doc.</th> */}

                            </tr>
                            {
                              this.state.HR2ApproverRejectDashboard.length > 0 ? this.state.HR2ApproverRejectDashboard.map((items) => {
                                return (
                                  <tr>
                                    <td><Icon iconName='View' onClick={() => this.getHR2ApproverView(items)} title='View' className={styles.iconcolor}></Icon></td>
                                    <td>{items.ID}</td>
                                    <td>{items.EmployeeID}</td>
                                    <td>{items.EmployeeName}</td>
                                    <td>{items.EmployeeType}</td>
                                    <td>{moment(items.DateofBirth).format("DD/MM/YYYY")}</td>
                                    <td>{items.DependentType}</td>
                                    <td>{items.AmountClaimed}</td>
                                    <td>{items.HRApprovedAmount}</td>
                                    <td>{items.Status}</td>
                                    {/* {
                                      items !== undefined && items !== null && items !== "" && items.AttachmentFiles && items.AttachmentFiles.length > 0 ? (
                                        items.AttachmentFiles.map((files) => (
                                          <li style={{ listStyle: 'decimal', color: '#428bca' }}><a href={files.ServerRelativeUrl} target='_blank'>{files.FileName}</a></li>
                                        ))
                                      ) : (
                                        <td>No Attachments</td>
                                      )
                                    } */}
                                  </tr>
                                )
                              })
                                : ""
                            }
                          </table>
      </div>
    </div>
  </div>
</PivotItem>

                {/* new code added end  */}
              </Pivot>
            </div>
          </div>
        </div>
        <Dialog
          hidden={!this.state.isDialogVisible}
          onDismiss={this.closeDialog}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'CHS Request',
            closeButtonAriaLabel: 'Close',
          }}
          containerClassName={'ms-dialogMainOverride ' + styles.textDialog}
        >
          <div className="card card-body">
            <div className="panel panel-default">
              <div className='panel-body'>
                <div className="row form-group" hidden={(!this.state.ShowHR1Tab) && (!this.state.ShowHR2Tab)} >
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">On behalf of:</Label>
                  </div>
                  <div className="col-sm-4" >
                    <Dropdown
                      placeHolder="Select Option"
                      options={onbehalfoption}
                      selectedKey={this.state.OnBehalf}
                      onChanged={(e, option) => {
                        if (e.key === "Yes") {
                          this.setState({
                            OnBehalf: e.key, DependentType: "",
                            ActualClaimAmountLable: "", showhideEmployeeNameLab: true
                          });
                          this.getAllEmployee();
                        }
                        if (e.key === "No") {
                          this.getEmployee();
                          this.setState({
                            OnBehalf: e.key, DependentType: "",
                            ActualClaimAmountLable: "", showhideEmployeeNameLab: false
                          });
                        }
                      }}
                      className="dropdown-style"
                    />
                  </div>
                  <div className="col-sm-2" hidden={!this.state.showhideEmployeeNameLab}
                  >
                    <Label className="control-Label font-weight-bold">Employee Name</Label>
                  </div>
                  <div className="col-sm-4" hidden={!this.state.showhideEmployeeNameLab}>
                    {/* <Dropdown
                      key={this.state.showhideEmployeeNameLab ? "dropdown-show" : "dropdown-hide"}
                      placeHolder="Select Employee"
                      hidden={!this.state.showhideEmployeeNameLab}
                      options={this.state.AllEmployeeCollObj}
                      selectedKey={this.state.Id}
                      onChanged={(e, option) => {
                        this.getSelectedEmployeeDetail(e);
                      }}
                      className="dropdown-style"
                    /> */}
                    
                      <Select
                      name="form-field-name"
                      autoFocus
                      clearable
                      value={selectedOption}
                      // value={this.state.AllEmployeeCollObj.find((item) => item.label === this.state.Id)}
                      onChange={(selectedOption) => this.getSelectedEmployeeDetail(selectedOption)}
                      options={this.state.AllEmployeeCollObj}
                      />
                  </div>

                  <div className="col-sm-4" hidden={!this.state.showhideEmployeeNameLab}>
                    {/* <TextField
                    placeholder="Search Employee"
                    value={this.state.searchValue}
                    onChanged={(newValue) => this.onSearchChange(newValue)} // Only 1 param
                  /> */}
                    {/* <Dropdown
                    key={this.state.showhideEmployeeNameLab ? "dropdown-show" : "dropdown-hide"}
                    placeHolder="Select Employee"
                    hidden={!this.state.showhideEmployeeNameLab}
                    selectedKey={this.state.Id}
                    options={this.state.AllEmployeeCollObj}

                    onChanged={(e, option) => this.getSelectedEmployeeDetail(e.key)}
                    className="dropdown-style"
                  /> */}

                    {/* <Select
                      // key={this.state.showhideEmployeeNameLab ? "dropdown-show" : "dropdown-hide"}

                      name="form-field-name"
                      // multi  //for Multi select
                      autofocus
                      value={value}
                      onChanged={(e, option) => {
                        this.getSelectedEmployeeDetail(e);
                      }}
                      // onChange={this.handleChange}
                      options={this.state.AllEmployeeCollObj}
                    /> */}
                  {/* <Select
                    name="form-field-name"
                    autoFocus
                    value={value}
                    onChange={(e, option) => this.getSelectedEmployeeDetail(option)}
                    options={this.state.AllEmployeeCollObj}
                  /> */}
                   {/* <div className="col-sm-6" hidden={!this.state.showhideEmployeeNameLab}> */}



                   {/* </div> */}
             


                    <div >

                    </div>

                  </div>

                </div>
                <div className="row form-group">
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Employee ID</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label">{this.state.EmployeeID}</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Employee Name</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.EmployeeName}</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Date of Birth </Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.DateofBirth}</Label>
                  </div>
                </div>
                <div className="row form-group">
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Scale</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.Scale}</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Employee Type </Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.EmpType}</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Designation</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.DesignationTitle}</Label>
                  </div>
                </div>
              </div>
            </div>
          </div>
          <div className="card card-body">
            <div className="panel panel-default">
              <div className='panel-body'>
                <div className="row form-group">
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Dependent Type </Label>
                  </div>
                  <div className="col-sm-2">
                    <Dropdown
                      placeHolder="Select Dependent"
                      options={[{ key: 'Self', text: 'Self' }, { key: 'Spouse', text: 'Spouse' }]}
                      selectedKey={this.state.DependentType}
                      onChanged={(e, options) => {
                        this.EligibleClaimAmount(e.key);
                      }}
                      className="dropdown-style"
                    />
                  </div>


                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Age</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.Age}</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">CHS Limit</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.Limit}</Label>
                  </div>
                </div>
                <div className="row form-group">
                  <div className="col-sm-2" >
                    <Label className="control-Label font-weight-bold">Amount Claimed<span >*</span></Label>
                  </div>
                  <div className="col-sm-2">
                    <TextField type='number'
                      min='0'
                      name="ExpenseDetails.Amount"
                      onChanged={(e: any) => this.handleInputChangeadd(event)}></TextField>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">CHS Eligiblity Limit</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.ActualClaimAmountLable}</Label>
                  </div>
                </div>
                <div className="row form-group" >
                <div className="col-sm-12" >
                <div className="col-md-7">
                  <span
                    hidden={!this.state.ExpenseDetailsAlert}
                    style={{
                      display: "block",
                      backgroundColor: "yellow",
                      color: "red",
                      padding: "10px",
                      border: "1px solid red",
                      borderRadius: "5px",
                      marginTop: "10px",
                      textAlign: "left",
                      fontWeight: "bold"
                    }}
                  >
                    Amount Claimed should not be more than the displayed CHS eligibility amount
                  </span>
                </div>
                </div>
                </div>


                <div className="row form-group" hidden={!this.state.IsSpouseEximEmployee}>
                  <div className="col-sm-2" >
                    <Label className="control-Label font-weight-bold">Is Spouse is Exim Employee<span>*</span></Label>
                  </div>
                  <div className='col-md-1'>
                    <Checkbox
                      label="Yes"
                      checked={this.state.selectedOptionCHBx === "Yes"}
                      onChange={this.handleCheckboxChange("Yes")}
                    />
                  </div>
                  <div className='col-md-1'>
                    <Checkbox
                      label="No"
                      checked={this.state.selectedOptionCHBx === "No"}
                      onChange={this.handleCheckboxChange("No")}
                    />
                  </div>
                </div>
              </div>
            </div>
            <div className="">
              <label className="col-md-2 control-Label font-weight-bold">Attachment <span >*</span></label>
              {/* <div className='col-md-2'>
                <input type="file" id="fileUpload" multiple onChange={this._handleFileChange} />
              </div> */}
              <div className="col-md-2">
                <input
                  type="file"
                  id="fileUpload"
                  multiple
                  onChange={this._handleFileChange}
                />
                <span
                  style={{
                    // display: "block",
                    // backgroundColor: "yellow",
                    color: "red",
                    // padding: "10px",
                    // border: "1px solid red",
                    // borderRadius: "5px",
                    // marginTop: "10px",
                    textAlign: "left",
                    fontSize: 10,
                    // fontWeight: "bold"
                  }}
                > Note : File name  Allowed characters: letters, numbers, dots, underscores, and spaces.</span>
                {/* <div>
    {this.state?.files.length > 0 && (
      <ul>
        {this.state.files.map((file, index) => (
          <li key={index}>{file.name}</li>
        ))}
      </ul>
    )}
  </div> */}
              </div>



            </div>



            {/* </div>  */}

          </div>
          <div className='text-center'>
            <PrimaryButton
              className={styles.custombtn + " mr-2"}
              onClick={() => this.BtnSubmitRequest()}
              disabled={this.state.isSubmitting}
            >
              {this.state.isSubmitting ? <Spinner size={SpinnerSize.small} /> : "Submit"}
            </PrimaryButton>
            {/* <PrimaryButton className={styles.custombtn + " " + "mr-2"} onClick={() => this.BtnSubmitRequest()} >Submit</PrimaryButton> */}
            <PrimaryButton onClick={() => this.closeDialog()} >Close</PrimaryButton>
          </div>
        </Dialog>
        <Dialog
          hidden={!this.state.isDialogHR1}
          onDismiss={this.closeDialog}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'CHS Approver',
            closeButtonAriaLabel: 'Close',
          }}
          containerClassName={'ms-dialogMainOverride ' + styles.textDialog}
        >
          <div className="card card-body">
            <div className="panel panel-default">
              <div className='panel-body'>
                <div className="row form-group">
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">CHS Request No</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label">{this.state.CHSApproverView.Title}</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Employee ID</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label">{this.state.CHSApproverView.EmployeeID}</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Employee Name</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.EmployeeName}</Label>
                  </div>
                </div>
                <div className="row form-group">
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Age</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.Age}</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Date of Birth </Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{moment(this.state.CHSApproverView.DateofBirth).format("DD/MM/YYYY")}</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Dependent Type </Label>
                  </div>
                  <div className="col-sm-2">
                    <Dropdown
                      placeHolder="Select Dependent"
                      disabled={true}
                      options={[{ key: 'Self', text: 'Self' }, { key: 'Spouse', text: 'Spouse' }]}
                      selectedKey={this.state.CHSApproverView.DependentType}
                      onChanged={(e, option) => this.setState({ DependentType: e.key })}
                      className="dropdown-style"
                    />
                  </div>
                </div>
                <div className="row form-group">
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Scale</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.Scale}</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Employee Type </Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.EmployeeType}</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Designation</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.Designation}</Label>
                  </div>
                </div>
              </div>
            </div>
          </div>
          <div className="card card-body">
            <div className="panel panel-default">
              <div className='panel-body'>
                <div className="row form-group">
                </div>
                <div className="row form-group">
                  <div className="col-sm-2" >
                    <Label className="control-Label font-weight-bold">Amount Claimed</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.AmountClaimed}</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">CHS Limit</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.Limit}</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">CHS Eligiblity Limit</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.EligibilityLimit}</Label>
                  </div>
                </div>
                <div className="row form-group">
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Final Amount</Label>
                  </div>
                  <div className="col-sm-2">
                    <TextField type='number'
                      name="ExpenseDetails.HR1FinalAmount"
                      onChanged={(e: any) => this.handleInputChangeadd(event)}></TextField>
                  </div>

                  {/* <div className="col-sm-2" >
                    <Label className="control-Label font-weight-bold">Is Spouse is Exim Employee<span>*</span></Label>
                  </div>
                  <div className='col-md-1'>
                    <Checkbox
                      label="Yes" 
                      checked={this.state.IsSpouseEximMember === "Yes"}
                      onChange={this.handleCheckboxChange("Yes")}
                    />
                  </div>
                  <div className='col-md-1'>
                    <Checkbox
                      label="No" 
                      checked={this.state.IsSpouseEximMember === "No"}
                      onChange={this.handleCheckboxChange("No")}
                    />
                    this.state.CHSApproverView.EligibilityLimit
                  </div> */}
                </div>

                
                <div className="row form-group" hidden={this.state.CHSApproverView.DependentType=='Self'} >
                  <div className="col-sm-2" >
                    <Label className="control-Label font-weight-bold">Is Spouse is Exim Employee<span>*</span></Label>
                  </div>
                  <div className='col-md-1'>
                    <Checkbox
                      label="Yes" 
                      disabled
                      checked={this.state.CHSApproverView.IsSpouseEximMember == "Yes"}
                      onChange={this.handleCheckboxChange("Yes")}
                    />
                  </div>
                  <div className='col-md-1'>
                    <Checkbox
                      label="No" 
                      disabled
                      checked={this.state.CHSApproverView.IsSpouseEximMember == "No"}
                      onChange={this.handleCheckboxChange("No")}
                    />
                  </div>
                </div>

                <div className="col-md-7">
                  <span
                    hidden={!this.state.ExpenseDetailsAlert}
                    style={{
                      display: "block",
                      backgroundColor: "yellow",
                      color: "red",
                      padding: "10px",
                      border: "1px solid red",
                      borderRadius: "5px",
                      marginTop: "10px",
                      textAlign: "left",
                      fontWeight: "bold"
                    }}
                  >
                    Amount Claimed should not be more than the displayed CHS eligibility amount
                  </span>
                </div>

              </div>
            </div>
            <div className="">
              <label className="col-md-2 control-Label font-weight-bold">Attachment</label>
              <div className='col-md-2'>
                {
                  this.state.CHSApproverView !== undefined && this.state.CHSApproverView !== null && this.state.CHSApproverView !== "" && this.state.CHSApproverView.AttachmentFiles && this.state.CHSApproverView.AttachmentFiles.length > 0 ? (
                    this.state.CHSApproverView.AttachmentFiles.map((files) => (
                      <li style={{ listStyle: 'decimal', color: '#428bca' }}><a href={files.ServerRelativeUrl} target='_blank'>{files.FileName}</a></li>
                    ))
                  ) : (
                    <div>No Attachments</div>
                  )
                }
              </div>
            </div>
            <div className="col-sm-12" style={{padding:0}}>
              <Label className="control-Label font-weight-bold">Remarks</Label>
            </div>
            <div className="col-sm-12" style={{padding:0}}>
              <TextField type='text'
                name="ExpenseDetails.HR1Remarks"
                onChanged={(e: any) => this.handleInputChangeadd(event)}></TextField>
            </div>
          </div>
          <div className='text-center'>
            <PrimaryButton className={styles.custombtn + " " + "mr-2"} onClick={() => this.BtnApproveHR1Request()} >Approve</PrimaryButton>
            <PrimaryButton className={styles.custombtn + " " + "mr-2"} onClick={() => this.BtnRejectRequest('HR1')} >Reject</PrimaryButton>
            <PrimaryButton onClick={() => this.closeDialog()} >Close</PrimaryButton>
          </div>
        </Dialog>
        <Dialog
          hidden={!this.state.isDialogViewHR1}
          onDismiss={this.closeDialog}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'CHS View Form',
            closeButtonAriaLabel: 'Close',
          }}
          containerClassName={'ms-dialogMainOverride ' + styles.textDialog}
        >
          <div className="card card-body">
            <div className="panel panel-default">
              <div className='panel-body'>
                <div className="row form-group">
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">CHS Request No</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label">{this.state.CHSApproverView.Title}</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Employee ID</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label">{this.state.CHSApproverView.EmployeeID}</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Employee Name</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.EmployeeName}</Label>
                  </div>
                </div>
                <div className="row form-group">
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Age</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.Age}</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Date of Birth </Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{moment(this.state.CHSApproverView.DateofBirth).format("DD/MM/YYYY")}</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Dependent Type </Label>
                  </div>
                  <div className="col-sm-2">
                    <Dropdown
                      placeHolder="Select Dependent"
                      disabled={true}
                      options={[{ key: 'Self', text: 'Self' }, { key: 'Spouse', text: 'Spouse' }]}
                      selectedKey={this.state.CHSApproverView.DependentType}
                      onChanged={(e, option) => this.setState({ DependentType: e.key })}
                      className="dropdown-style"
                    />
                  </div>
                </div>
                <div className="row form-group">
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Scale</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.Scale}</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Employee Type </Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.EmployeeType}</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Designation</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.Designation}</Label>
                  </div>
                </div>
              </div>
            </div>
          </div>
          <div className="card card-body">
            <div className="panel panel-default">
              <div className='panel-body'>
                <div className="row form-group">
                </div>
                <div className="row form-group">
                  <div className="col-sm-2" >
                    <Label className="control-Label font-weight-bold">Amount Claimed</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.AmountClaimed}</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">CHS Limit</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.Limit}</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">CHS Eligiblity Limit</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.EligibilityLimit}</Label>
                  </div>
                </div>
                <div className="row form-group">
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">HR1  Approved Amount</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.FinalAmount}</Label>
                  </div>
                  <div className="col-sm-2" >
                    <Label className="control-Label font-weight-bold">HR1 Remarks</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.HR1Remark}</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Final Approved Amount</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.HRApprovedAmount}</Label>
                  </div>
                </div>
                     
                <div className="row form-group" hidden={this.state.CHSApproverView.DependentType=='Self'} >
                  <div className="col-sm-2" >
                    <Label className="control-Label font-weight-bold">Is Spouse is Exim Employee<span>*</span></Label>
                  </div>
                  <div className='col-md-1'>
                    <Checkbox
                      label="Yes" 
                      disabled
                      checked={this.state.CHSApproverView.IsSpouseEximMember == "Yes"}
                      onChange={this.handleCheckboxChange("Yes")}
                    />
                  </div>
                  <div className='col-md-1'>
                    <Checkbox
                      label="No" 
                      disabled
                      checked={this.state.CHSApproverView.IsSpouseEximMember == "No"}
                      onChange={this.handleCheckboxChange("No")}
                    />
                  </div>
                </div>
              </div>
            </div>
            <div className="">
              <label className="col-md-2 control-Label font-weight-bold">Attachment</label>
              <div className='col-md-2'>
                {
                  this.state.CHSApproverView !== undefined && this.state.CHSApproverView !== null && this.state.CHSApproverView !== "" && this.state.CHSApproverView.AttachmentFiles && this.state.CHSApproverView.AttachmentFiles.length > 0 ? (
                    this.state.CHSApproverView.AttachmentFiles.map((files) => (
                      <li style={{ listStyle: 'decimal', color: '#428bca' }}><a href={files.ServerRelativeUrl} target='_blank'>{files.FileName}</a></li>
                    ))
                  ) : (
                    <div>No Attachments</div>
                  )
                }
              </div>
            </div>
          </div>
          <div className='text-center'>
            <PrimaryButton onClick={() => this.closeDialog()} >Close</PrimaryButton>
          </div>
        </Dialog>
        <Dialog
          hidden={!this.state.isDialogHR2}
          onDismiss={this.closeDialog}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'CHS Approver',
            closeButtonAriaLabel: 'Close',
          }}
          containerClassName={'ms-dialogMainOverride ' + styles.textDialog}
        >
          <div className="card card-body">
            <div className="panel panel-default">
              <div className='panel-body'>
                <div className="row form-group">
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">CHS Request No</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label">{this.state.CHSApproverView.Title}</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Employee ID</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label">{this.state.CHSApproverView.EmployeeID}</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Employee Name</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.EmployeeName}</Label>
                  </div>
                </div>
                <div className="row form-group">
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Age</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.Age}</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Date of Birth </Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{moment(this.state.CHSApproverView.DateofBirth).format("DD/MM/YYYY")}</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Dependent Type </Label>
                  </div>
                  <div className="col-sm-2">
                    <Dropdown
                      placeHolder="Select Dependent"
                      disabled={true}
                      options={[{ key: 'Self', text: 'Self' }, { key: 'Spouse', text: 'Spouse' }]}
                      selectedKey={this.state.CHSApproverView.DependentType}
                      onChanged={(e, option) => this.setState({ DependentType: e.key })}
                      className="dropdown-style"
                    />
                  </div>
                </div>
                <div className="row form-group">
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Scale</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.Scale}</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Employee Type </Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.EmployeeType}</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Designation</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.Designation}</Label>
                  </div>
                </div>
              </div>
            </div>
          </div>
          <div className="card card-body">
            <div className="panel panel-default">
              <div className='panel-body'>
                <div className="row form-group">
                </div>
                <div className="row form-group">
                  <div className="col-sm-2" >
                    <Label className="control-Label font-weight-bold">Amount Claimed</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.AmountClaimed}</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">CHS Limit</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.Limit}</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">CHS Eligiblity Limit</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.EligibilityLimit}</Label>
                  </div>
                </div>
                <div className="row form-group">
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">HR1  Approved Amount</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.FinalAmount}</Label>
                  </div>
                  <div className="col-sm-2" >
                    <Label className="control-Label font-weight-bold">HR1 Remarks</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.HR1Remark}</Label>
                  </div>
                </div>
                <div className="col-sm-2">
                  <Label className="control-Label font-weight-bold">Final  Approved Amount</Label>
                </div>
                <div className="col-sm-2">
                  <TextField type='number'
                    name="ExpenseDetails.HR2FinalAmount"
                    onChanged={(e: any) => this.handleInputChangeadd(event)}></TextField>
                </div>

                     
                <div className="row form-group" hidden={this.state.CHSApproverView.DependentType=='Self'} >
                  <div className="col-sm-2" >
                    <Label className="control-Label font-weight-bold">Is Spouse is Exim Employee<span>*</span></Label>
                  </div>
                  <div className='col-md-1'>
                    <Checkbox
                      label="Yes" 
                      disabled
                      checked={this.state.CHSApproverView.IsSpouseEximMember == "Yes"}
                      onChange={this.handleCheckboxChange("Yes")}
                    />
                  </div>
                  <div className='col-md-1'>
                    <Checkbox
                      label="No" 
                      disabled
                      checked={this.state.CHSApproverView.IsSpouseEximMember == "No"}
                      onChange={this.handleCheckboxChange("No")}
                    />
                  </div>
                </div>

                <div className="col-md-7">
                  <span
                    hidden={!this.state.ExpenseDetailsAlert}
                    style={{
                      display: "block",
                      backgroundColor: "yellow",
                      color: "red",
                      padding: "10px",
                      border: "1px solid red",
                      borderRadius: "5px",
                      marginTop: "10px",
                      textAlign: "left",
                      fontWeight: "bold"
                    }}
                  >
                    Amount Claimed should not be more than the displayed CHS eligibility amount
                  </span>
                </div>
              </div>
            </div>
            <div className="">
              <label className="col-md-2 control-Label font-weight-bold">Attachment</label>
              <div className='col-md-2'>
                {
                  this.state.CHSApproverView !== undefined && this.state.CHSApproverView !== null && this.state.CHSApproverView !== "" && this.state.CHSApproverView.AttachmentFiles && this.state.CHSApproverView.AttachmentFiles.length > 0 ? (
                    this.state.CHSApproverView.AttachmentFiles.map((files) => (
                      <li style={{ listStyle: 'decimal', color: '#428bca' }}><a href={files.ServerRelativeUrl} target='_blank'>{files.FileName}</a></li>
                    ))
                  ) : (
                    <div>No Attachments</div>
                  )
                }
              </div>
            </div>
            <div className="col-sm-12" >
              <Label className="control-Label font-weight-bold">Remarks</Label>
            </div>
            <div className="col-sm-12">
              <TextField type='text'
                name="ExpenseDetails.HR2Remarks"
                onChanged={(e: any) => this.handleInputChangeadd(event)}></TextField>
            </div>
          </div>
          <div className='text-center'>
            <PrimaryButton className={styles.custombtn + " " + "mr-2"} onClick={() => this.BtnApproveHR2Request()} >Approve</PrimaryButton>
            <PrimaryButton className={styles.custombtn + " " + "mr-2"} onClick={() => this.BtnRejectRequest('HR2')} >Reject</PrimaryButton>
            <PrimaryButton onClick={() => this.closeDialog()} >Close</PrimaryButton>
          </div>
        </Dialog>
        <Dialog
          hidden={!this.state.isDialogViewHR2}
          onDismiss={this.closeDialog}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'CHS View Form',
            closeButtonAriaLabel: 'Close',
          }}
          containerClassName={'ms-dialogMainOverride ' + styles.textDialog}
        >
          <div className="card card-body">
            <div className="panel panel-default">
              <div className='panel-body'>
                <div className="row form-group">
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">CHS Request No</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label">{this.state.CHSApproverView.Title}</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Employee ID</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label">{this.state.CHSApproverView.EmployeeID}</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Employee Name</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.EmployeeName}</Label>
                  </div>
                </div>
                <div className="row form-group">
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Age</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.Age}</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Date of Birth </Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{moment(this.state.CHSApproverView.DateofBirth).format("DD/MM/YYYY")}</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Dependent Type </Label>
                  </div>
                  <div className="col-sm-2">
                    <Dropdown
                      placeHolder="Select Dependent"
                      disabled={true}
                      options={[{ key: 'Self', text: 'Self' }, { key: 'Spouse', text: 'Spouse' }]}
                      selectedKey={this.state.CHSApproverView.DependentType}
                      onChanged={(e, option) => this.setState({ DependentType: e.key })}
                      className="dropdown-style"
                    />
                  </div>
                </div>
                <div className="row form-group">
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Scale</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.Scale}</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Employee Type </Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.EmployeeType}</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Designation</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.Designation}</Label>
                  </div>
                </div>
              </div>
            </div>
          </div>
          <div className="card card-body">
            <div className="panel panel-default">
              <div className='panel-body'>
                <div className="row form-group">
                </div>
                <div className="row form-group">
                  <div className="col-sm-2" >
                    <Label className="control-Label font-weight-bold">Amount Claimed</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.AmountClaimed}</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">CHS Limit</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.Limit}</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">CHS Eligiblity Limit</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.EligibilityLimit}</Label>
                  </div>
                </div>
                <div className="row form-group">
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">HR1  Approved Amount</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.FinalAmount}</Label>
                  </div>
                  <div className="col-sm-2" >
                    <Label className="control-Label font-weight-bold">HR1 Remarks</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.HR1Remark}</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Final  Approved Amount</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.HRApprovedAmount}</Label>
                  </div>
                </div>

                     
                <div className="row form-group" hidden={this.state.CHSApproverView.DependentType=='Self'} >
                  <div className="col-sm-2" >
                    <Label className="control-Label font-weight-bold">Is Spouse is Exim Employee<span>*</span></Label>
                  </div>
                  <div className='col-md-1'>
                    <Checkbox
                      label="Yes" 
                      disabled
                      checked={this.state.CHSApproverView.IsSpouseEximMember == "Yes"}
                      onChange={this.handleCheckboxChange("Yes")}
                    />
                  </div>
                  <div className='col-md-1'>
                    <Checkbox
                      label="No" 
                      disabled
                      checked={this.state.CHSApproverView.IsSpouseEximMember == "No"}
                      onChange={this.handleCheckboxChange("No")}
                    />
                  </div>
                </div>
              </div>
            </div>
            <div className="">
              <label className="col-md-2 control-Label font-weight-bold">Attachment</label>
              <div className='col-md-2'>
                {
                  this.state.CHSApproverView !== undefined && this.state.CHSApproverView !== null && this.state.CHSApproverView !== "" && this.state.CHSApproverView.AttachmentFiles && this.state.CHSApproverView.AttachmentFiles.length > 0 ? (
                    this.state.CHSApproverView.AttachmentFiles.map((files) => (
                      <li style={{ listStyle: 'decimal', color: '#428bca' }}><a href={files.ServerRelativeUrl} target='_blank'>{files.FileName}</a></li>
                    ))
                  ) : (
                    <div>No Attachments</div>
                  )
                }
              </div>
            </div>
            {/* <div className="col-sm-12" >
              <Label className="control-Label font-weight-bold">Remarks</Label>
            </div>
            <div className="col-sm-12">
              <TextField type='text'
                name="ExpenseDetails.HR2Remarks"
                onChanged={(e: any) => this.handleInputChangeadd(event)}></TextField>
            </div> */}
          </div>
          <div className='text-center'>
            <PrimaryButton onClick={() => this.closeDialog()} >Close</PrimaryButton>
          </div>
        </Dialog>
      </div>
    );
  }
}