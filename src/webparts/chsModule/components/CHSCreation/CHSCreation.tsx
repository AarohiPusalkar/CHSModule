import * as React from 'react';
import styles from '../ChsModule.module.scss'
import * as moment from 'moment'
import { IChsModuleProps } from '../IChsModuleProps';
import { Web } from '@pnp/sp/presets/all';
import { BaseButton, Button, Checkbox, FontWeights, IPersonaProps } from 'office-ui-fabric-react';
import useSPCRUD from '../../../services/bal/spcrud';
import EmployeeOps from '../../../services/bal/EmployeeMaster';
import { IEmployeeMaster } from '../../../services/interface/IEmployeeMaster';
import { IEmployeeCHSLimitMaster } from '../../../services/interface/IEmployeeCHSLimitMaster';
import EmployeeCHSLimitMasterOps from '../../../services/bal/EmployeeCHSLimitMaster';
import { Icon, DefaultButton, Dialog, DialogFooter, DialogType, Dropdown, IDropdownOption, PrimaryButton, IDropdown, } from 'office-ui-fabric-react';
import { Pivot, PivotItem, IPivotItemProps, PivotLinkSize, PivotLinkFormat } from 'office-ui-fabric-react/lib/Pivot';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { SPComponentLoader } from '@microsoft/sp-loader';
import Select from 'react-select-plus';
import 'react-select-plus/dist/react-select-plus.css';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { ENV_CONFIG } from '../../../../Enviroment/envConfig';
SPComponentLoader.loadCss('https://cdn.jsdelivr.net/npm/bootstrap@4.6.0/dist/css/bootstrap.min.css');
SPComponentLoader.loadCss('https://stackpath.bootstrapcdn.com/font-awesome/4.7.0/css/font-awesome.min.css');
import '../../assets/CHSCretaionstyle.scss';
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";

export interface ISelectState {
  selectedOption?: string;
}

const onbehalfoption: IDropdownOption[] = [
  { key: 'Yes', text: 'Yes' },
  { key: 'No', text: 'No' }
];

export default class CHSCreation extends React.Component<IChsModuleProps, any> {

  private readonly ADMIN_ONBEHALF_DESIGNATIONS: string[] = [
    "Administrative Officer (AO)",
    "Administrative Manager (AM)",
    "Chief Administrative Manager (CAM)",
    "Sr. Administrative Manager (SAM)",
    "Sr. Administrative Officer (SAO)"
  ];

  constructor(props: any) {
    super(props);

    this.state = {
      activeTab: "Pending", // Set the initial active tab here
      dependentRowFiles: {},
      dependentDocsByType: {},
      GroupId: "",
      AllEmployeeCollObj: [],
      selectedOptionCHBx: null,
      isSubmitting: false,
      isApproving: false,
      isRejecting: false,
      selectedOption: '',
      transformedOptions: [],
      IsSpouseEximEmployee: false,
      filteredData: [],
      showhideEmployeeNameLab: false,
      isOnBehalfDisabled: false,
      ShowHR1Tab: false,
      ShowHR2Tab: false,
      ShowTagApproverTab: false,

      ShowHRTab: true,
      CalculatedCHSEligibilityAmountLabel: "",
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

      TagApproverPendingDashboard: [],
      TagApproverApprDashboard: [],
      TagApproverRejectDashboard: [],

      DependentType: "ALL",
      AmountClaimed: "",
      IsMarried: false,
      TotalAmountClaimed: 0,
      CHSEligibilityAmount: 0,
      dependentitems: [],
      dropdownOptions: [{ key: 'Self', text: 'Self' }, { key: 'Spouse', text: 'Spouse' }],
      allDashboardData: [],

      allDashboardDataHR1Pending: [],

      allDashboardDataHR1Rejected: [],

      allDashboardDataHR1Approved: [],

      allDashboardDataHR2Pending: [],

      allDashboardDataHR2Rejected: [],

      allDashboardDataHR2Approved: [],

      allDashboardDataTagApproverPending: [],

      allDashboardDataTagApproverRejected: [],

      allDashboardDataTagApproverApproved: [],

      TagApproverApproverRejectDashboard: [],

      allDashboardData2: [],

      filteredDashboard: [], // Filtered data for rendering
      isDialogVisible: false,
      isDialogGH: false,
      isDialogHR1: false,
      isDialogHR2: false,
      isDialogTagApprover: false,


      isDialogViewHR1: false,
      isDialogViewHR2: false,
      isDialogViewTagApprover: false,

      EmployeeName: "",
      ExpenseDetailsAlert: false,
      searchValue: "",
      filteredEmployees: [],
      EmployeeID: '',
      EmployeeIDId: '',
      DesignationId: '',
      CompanyEmail: '',
      GHStatus: '',
      file: null,
      reqID: '',
      isClearable: true,
      isSearchable: true,
      CHSApproverView: [],


      filteredOptions: [],
      selectedId: null,
      isDropdownOpen: false, //  Track dropdown state
      activeHR1Tab: 'Pending',
      activeHR2Tab: 'Pending',
      activeHRTab: 'Pending',
      activeTagApproverTab: 'Pending',


      activeHRdashTab: 'Pending',
      selectedOuterTab: '',
      CurrentFinancialYear: this.getFinancialYear(),

      isOnBehalfandRetired: false,

      // --- Admin On Behalf (AO/AM/CAM/SAM/SAO) ---
      isAdminOnBehalfUser: false,
      EmployeeSubGroupIds: [],
      activeOnBehalfDashTab: localStorage.getItem("activeOnBehalfDashTab") || "Pending",
      OnBehalfDashboard: [],
      OnBehalfApprovedDashboard: [],
      OnBehalfRejectedDashboard: []
    };
  }

  async componentDidMount() {
    const savedTab3 = localStorage.getItem("activeHRdashTab");
    const savedTab1 = localStorage.getItem("activeHR1Tab");
    const savedTab2 = localStorage.getItem("activeHR2Tab");

    this.setState({
      ...(savedTab3 && { activeHRdashTab: savedTab3 }),
      ...(savedTab1 && { activeHR1Tab: savedTab1 }),
      ...(savedTab2 && { activeHR2Tab: savedTab2 }),
    });
    let tabLinks = document.querySelectorAll(".tablink");

    // Get saved active tab from localStorage
    let savedTabId = localStorage.getItem("activeTabId");

    // If there's a saved tab ID, use it; otherwise, use the default
    let activeTab = document.getElementById("defaultOpen");
    if (savedTabId) {
      activeTab = savedTabId ? document.getElementById(savedTabId)
        : document.getElementById("defaultOpen");
      activeTab.classList.add("active");
    }
    else {
      const urlParams = new URLSearchParams(window.location.search);
      const dashTabParam = urlParams.get("dashtab") || "Pending";
      const pivotTabParam = urlParams.get("ptab") || "User";
      const ishr1 = await this.checkUserInGroupsForHR1TabNav(['HR1_Group'])
      const ishr2 = await this.checkUserInGroupsForHR2TabNav(['HR2_Group'])
      ///const isTagApprover = await this.checkUserInGroupsForTagTabNav(['TagApprover'])
      const isTagApprover = await this.checkUserInGroupsForTagTabNav(['TADA_TAGUsers'])

      switch (pivotTabParam) {
        case "HR1":
          if (ishr1) {
            this.setState({
              ShowHR1Tab: true,
              activeHR1Tab: dashTabParam,
              selectedOuterTab: "HR1"
            });
          }
          else {
            this.setState({
              ShowHRTab: true,
              activeHRdashTab: dashTabParam,
              selectedOuterTab: "User"
            });
          }

          break;
        case "HR2":
          if (ishr2) {
            this.setState({
              ShowHR2Tab: true,
              activeHR2Tab: dashTabParam,
              selectedOuterTab: "HR2"
            });
          }
          else {
            this.setState({
              ShowHRTab: true,
              activeHRdashTab: dashTabParam,
              selectedOuterTab: "User"
            });
          }

          break;

        case "TagApprover":
          if (isTagApprover) {
            this.setState({
              ShowTagApproverTab: true,
              activeTagApproverTab: dashTabParam,
              selectedOuterTab: "TagApprover"
            });
          }
          else {
            this.setState({
              ShowHRTab: true,
              activeHRdashTab: dashTabParam,
              selectedOuterTab: "User"
            });
          }

          break;

        default:
          this.setState({
            ShowHRTab: true,
            activeHRdashTab: dashTabParam,
            selectedOuterTab: "User"
          });
          break;
      }
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

    await this.TagApproverPendingDashboard();
    await this.TagApproverApprovedDashboards();
    await this.TagApproverRejectedDashboards();

    await this.GetEmployeelimit();
    await this.checkUserInGroups(["HR1_Group", "HR2_Group"]);
    await this.checkUserInGroupsForHR1Tab(["HR1_Group"]);
    await this.checkUserInGroupsForHR2Tab(["HR2_Group"]);
    await this.checkUserInGroupsForTagApproverTab(["TADA_TAGUsers"]);

    await this.getEmployee();

    if (this.state.isAdminOnBehalfUser) {
      await this.OnBehalfApprovedDashboards();
      await this.OnBehalfRejectedDashboards();
      await this.OnBehalfPendingDashboard();
    }

    this.setState({ dropdownOptions: this.getDependentOptionsByDesignation() });
  }

  public getFinancialYear() {
    const today = new Date();
    const year = today.getFullYear();
    const month = today.getMonth() + 1; // Months are 0-based

    let startYear, endYear;

    if (month >= 4) {
      // From April to December → current FY starts this year
      startYear = year;
      endYear = year + 1;
    } else {
      // From January to March → current FY started last year
      startYear = year - 1;
      endYear = year;
    }

    return `April 1, ${startYear} – March 31, ${endYear}`;
  }

  private handleDependentRowFileChange = (rowIndex: number, ev: React.ChangeEvent<HTMLInputElement>) => {
    const files = ev.target.files;
    if (!files || files.length === 0) return;

    const file = this.validateAndFilterFiles(files)[0];
    if (!file) return;

    this.setState((prev) => ({
      dependentRowFile: { ...(prev.dependentRowFile || {}), [rowIndex]: file }
    }));
  };

  private validateAndFilterFiles(files: FileList): File[] {
    const validFiles: File[] = [];
    const invalidFiles: string[] = [];
    const fileNameRegex = /^[a-zA-Z0-9._\s]+$/;

    for (let i = 0; i < files.length; i++) {
      const f = files[i];
      if (fileNameRegex.test(f.name)) validFiles.push(f);
      else invalidFiles.push(f.name);
    }

    if (invalidFiles.length) {
      alert(`These files have invalid names and won't be uploaded: ${invalidFiles.join(", ")}\nAllowed: letters, numbers, dots, underscores, spaces.`);
    }
    return validFiles;
  }

  private async uploadDependentRowDocs(healthCheckupServiceId: number): Promise<void> {
    const web = Web(ENV_CONFIG.siteUrl);
    const titleFolder = `CHS000${healthCheckupServiceId}`;
    const baseFolder = `HealthCheckupServiceDocuments/${titleFolder}`;

    await web.folders.add(baseFolder).catch(() => null);

    for (let rowIndex = 0; rowIndex < this.state.dependentitems.length; rowIndex++) {
      var dep = this.state.dependentitems[rowIndex];
      var depType = "Row" + (rowIndex + 1);

      if (dep && dep.DependentType) {
        depType = dep.DependentType;
      }

      var file: File = undefined;

      if (this.state.dependentRowFile && this.state.dependentRowFile[rowIndex]) {
        file = this.state.dependentRowFile[rowIndex];
      }
      if (!file) continue;

      await web.folders.add(`${baseFolder}/${depType}`).catch(() => null);

      const targetFolder = web.getFolderByServerRelativeUrl(
        `${this.props.currentSPContext.pageContext.web.serverRelativeUrl}/${baseFolder}/${depType}`
      );

      await targetFolder.files.add(file.name, file, true);
    }
  }

  handleOuterTabClick = (tabName) => {
    this.setState({
      selectedOuterTab: tabName
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
    if (prevProps.selectedOuterTab !== this.props.selectedOuterTab) {
      this.setState({ activeHR1Tab: 'Pending' }); // Or your default
    }
  }

  public getCurrentUser = async () => {
    const spCrudObj = await useSPCRUD();
    return await spCrudObj.currentUser(this.props).then(cuser => {
      this.setState({ Currentuser: cuser });
      return cuser;
    });
  }

  public async checkUserInGroupsForHR2TabNav(groupsToCheck: string[]): Promise<boolean> {
    try {
      const spCrudObj = await useSPCRUD();
      const userGroups = await spCrudObj.currentUserGroup(this.props);

      if (!userGroups || userGroups.length === 0) {
        console.log(" User is not part of any SharePoint group.");
        return false;
      }

      // Normalize to lowercase for comparison
      const userGroupTitles = userGroups.map((g) => g.Title.toLowerCase());
      const targetGroups = groupsToCheck.map((g) => g.toLowerCase());

      // 🔹 Match ANY of the target groups
      const isUserInAnyGroup = targetGroups.some((target) =>
        userGroupTitles.includes(target)
      );

      console.log(" Is user in HR groups?", isUserInAnyGroup);

      // Optional: Set state if needed
      this.setState({ ShowHR2Tab: isUserInAnyGroup });

      return isUserInAnyGroup;
    } catch (error) {
      console.error(" Error checking user group membership:", error);
      return false;
    }
  }

  public async checkUserInGroupsForTagTabNav(groupsToCheck: string[]): Promise<boolean> {
    try {
      const spCrudObj = await useSPCRUD();
      const userGroups = await spCrudObj.currentUserGroupAnotherSiteCollection(this.props);

      if (!userGroups || userGroups.length === 0) {
        console.log(" User is not part of any SharePoint group.");
        return false;
      }

      // Normalize to lowercase for comparison
      const userGroupTitles = userGroups.map((g) => g.Title.toLowerCase());
      const targetGroups = groupsToCheck.map((g) => g.toLowerCase());

      // 🔹 Match ANY of the target groups
      const isUserInAnyGroup = targetGroups.some((target) =>
        userGroupTitles.includes(target)
      );

      console.log(" Is user in HR groups?", isUserInAnyGroup);

      // Optional: Set state if needed
      this.setState({ ShowTagApproverTab: isUserInAnyGroup });

      return isUserInAnyGroup;
    } catch (error) {
      console.error(" Error checking user group membership:", error);
      return false;
    }
  }

  public async checkUserInGroupsForHR1TabNav(groupsToCheck: string[]): Promise<boolean> {
    try {
      const spCrudObj = await useSPCRUD();
      const userGroups = await spCrudObj.currentUserGroup(this.props);

      if (!userGroups || userGroups.length === 0) {
        console.log(" User is not part of any SharePoint group.");
        return false;
      }

      // Normalize to lowercase for comparison
      const userGroupTitles = userGroups.map((g) => g.Title.toLowerCase());
      const targetGroups = groupsToCheck.map((g) => g.toLowerCase());

      // 🔹 Match ANY of the target groups
      const isUserInAnyGroup = targetGroups.some((target) =>
        userGroupTitles.includes(target)
      );

      console.log(" Is user in HR groups?", isUserInAnyGroup);

      // Optional: Set state if needed
      this.setState({ ShowHR1Tab: isUserInAnyGroup });

      return isUserInAnyGroup;
    } catch (error) {
      console.error(" Error checking user group membership:", error);
      return false;
    }
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
        this.setState({ showhideEmployeeNameLab: false, OnBehalf: 'No', isOnBehalfDisabled: true })
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

  public async checkUserInGroupsForTagApproverTab(groups: any) {
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
        this.setState({ ShowTagApproverTab: true })
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
      let matchedLimit = limitData.filter((e) => e.Scale.Title == employeeData.Scale && e.Designation.Title == employeeData.DesignationTitle && e.EmployeeType == employeeData.EmployeeType);

      let LIMIT = "";
      LIMIT = matchedLimit.length > 0 && matchedLimit !== undefined ? matchedLimit[0].Limit : "";

      const subGroupIds = this.parseMultiLookupIds(employeeData.SubGroupId || employeeData.SubGroup);

      this.setState({
        EmployeeInfodb: employeeData,
        AllEmployeeCollObj: [],
        EmployeeName: employeeData.EmployeeName,
        EmployeeIDId: employeeData.Id,
        DependentType: "",
        CalculatedCHSEligibilityAmountLabel: "",
        CompanyEmail: employeeData.CompanyEmail,

        EmployeeID: employeeData.EmployeeId,
        DesignationId: employeeData.DesignationId,
        DesignationTitle: employeeData.DesignationTitle,
        DateofBirth: employeeData.DateofBirth,
        Scale: employeeData.Scale,
        Age: parseInt(employeeData.Age),
        EmployeeType: employeeData.EmployeeType,
        isAdminOnBehalfUser: this.ADMIN_ONBEHALF_DESIGNATIONS.indexOf(employeeData.DesignationTitle) !== -1,
        EmployeeSubGroupIds: subGroupIds,
        Limit: LIMIT
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

          let LIMIT = "";
          LIMIT = matchedLimit.length > 0 && matchedLimit !== undefined ? matchedLimit[0].Limit : ""

          this.setState({
            EmployeeName: selectedEmp.EmployeeName,
            isDropdownOpen: false, //  Close dropdown on selection
            CompanyEmail: selectedEmp.CompanyEmail,
            EmployeeID: selectedEmp.Title,
            DesignationId: selectedEmp.DesignationId,
            DesignationTitle: selectedEmp.Designation.Title,
            DateofBirth: selectedEmp.DOB
              ? `${new Date(selectedEmp.DOB).getDate()}-${new Date(selectedEmp.DOB).getMonth() + 1}-${new Date(selectedEmp.DOB).getFullYear()}`
              : null,
            Scale: selectedEmp.Scale.Title,
            Payscale: selectedEmp.Payscale.Title,
            Age: parseInt(selectedEmp.Age),
            EmployeeType: selectedEmp.EmployeeType.Title,
            DependentType: "",
            CalculatedCHSEligibilityAmountLabel: "",
            selectedOption: e,
            Limit: LIMIT,
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
        label: item.EmployeeName,
        value: item.EmployeeName,
      }));
      this.setState({
        EmployeeInfodb: results,
        AllEmployeeCollObj: employeeOptions,
        filteredOptions: employeeOptions,

      });
      return results;
    });
  };

  private parseMultiLookupIds = (val: any): number[] => {
    // Case 1: already an array of numbers
    if (Array.isArray(val)) {
      return val
        .map((x) => parseInt(x, 10))
        .filter((n) => !isNaN(n));
    }

    // Case 2: array of objects with Id
    if (Array.isArray(val) && val.length && val[0] && val[0].Id) {
      return val.map((x) => parseInt(x.Id, 10)).filter((n) => !isNaN(n));
    }

    // Case 3: SharePoint multi-lookup string like "1;#SG1;#3;#SG3"
    if (typeof val === "string") {
      return (val || "")
        .split(";#")
        .filter((x) => /^\d+$/.test(x))
        .map((x) => parseInt(x, 10))
        .filter((n) => !isNaN(n));
    }

    // Case 4: single number
    if (typeof val === "number") return [val];

    // Default
    return [];
  };


  private loadAdminOnBehalfEmployees = async (subGroupIds: number[]): Promise<void> => {
    if (!subGroupIds || !subGroupIds.length) {
      this.setState({ AllEmployeeCollObj: [], filteredOptions: [], EmployeeInfodb: [] });
      return;
    }

    const sgFilter = subGroupIds.map((id) => `SubGroup/Id eq ${id}`).join(" or ");
    const filter =
      `(${sgFilter}) and ` +
      `(Designation/Title eq 'Managing Director (MD)' or Designation/Title eq 'Deputy Managing Director (DMD)') and (EmployeeType/Title eq 'PERMANENT')`;

    const results = await EmployeeOps().getEmployeeMasterId(
      filter,
      { column: "EmployeeName", isAscending: true },
      this.props
    );

    const options = (results || []).map((emp: any) => ({
      key: emp.Id.toString(),
      text: emp.EmployeeName,
      label: emp.EmployeeName,
      value: emp.EmployeeName
    }));

    this.setState({
      EmployeeInfodb: results,
      AllEmployeeCollObj: options,
      filteredOptions: options
    });
  };

  public _getPeoplePickerItems = (item: IPersonaProps[]) => {
    console.log(item);
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
      isDialogTagApprover: false,


      isDialogViewHR1: false,
      isDialogViewHR2: false,
      isDialogViewTagApprover: false

    });
  }
  public CreateRequest = () => {

    this.setState({
      IsSpouseEximEmployee: false,
      isDialogVisible: true,
      activeTab: 'Pending',
      dependentitems: [],
      DependentType: "",
      AmountClaimed: "",
      CalculatedCHSEligibilityAmountLabel: ""
    })
  }

  private exportToExcel = async (): Promise<void> => {
    await EmployeeOps().getAgeBasedHealthCheckupReportData(this.props).then(ReportData => {
      const ws = XLSX.utils.json_to_sheet(ReportData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "ExportedData");

      const excelBuffer: any = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      const blob: Blob = new Blob([excelBuffer], { type: "application/octet-stream" });

      saveAs(blob, "Report for Age based Health Checkup.xlsx");
    }
    );
  };

  public openPage(pageName, elmnt, color, tabName) {
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

        //working
        const LIMIT = parseFloat(this.state.Limit) || 0;
        if (this.state.DesignationTitle != "Managing Director (MD)" &&
          this.state.DesignationTitle != "Deputy Managing Director (DMD)") {
          if (LIMIT === 0) {
            return this.fail("Please Map Limit!");
          }
        }

        this.setState({
          ExpenseDetails: {
            ...this.state.ExpenseDetails,
            Amount: +value,
          },
          AmountClaimed: value
        });

        this.EligibleClaimAmount(this.state.DependentType, value);
        break;
      case "ExpenseDetails.HR1Remarks":
        this.setState({
          ExpenseDetails: {
            ...this.state.ExpenseDetails,
            HR1Remarks: '' + value
          }
        });
        break;
      case "ExpenseDetails.HRRemarkForRetired":
        this.setState({
          ExpenseDetails: {
            ...this.state.ExpenseDetails,
            HRRemarkForRetired: '' + value
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
      case "ExpenseDetails.TagApproverRemarks":
        this.setState({
          ExpenseDetails: {
            ...this.state.ExpenseDetails,
            TagApproverRemarks: '' + value
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
          ExpenseDetails: {
            ...this.state.ExpenseDetails,
            HR1FinalAmount: + value
          }
        });

        //working
        let hr1validAmount = 0;
        if ((this.state.CHSApproverView.Designation == "Managing Director (MD)" ||
          this.state.CHSApproverView.Designation == "Deputy Managing Director (DMD)")) {
          hr1validAmount = this.state.CHSApproverView.AmountClaimed;
        }
        else {
          hr1validAmount = this.state.CHSApproverView.CHSEligibilityAmount;
        }
        //if (value > this.state.CHSApproverView.EligibilityLimit) {
        if (value > hr1validAmount) { //30 Dec
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


      case "ExpenseDetails.HR2FinalAmount":
        this.setState({
          ExpenseDetails: {
            ...this.state.ExpenseDetails,
            HR2FinalAmount: + value
          }
        });


        let hr2validAmount = 0;
        if ((this.state.CHSApproverView.Designation == "Managing Director (MD)" ||
          this.state.CHSApproverView.Designation == "Deputy Managing Director (DMD)")) {
          hr2validAmount = this.state.CHSApproverView.AmountClaimed;
        }
        else {
          hr2validAmount = this.state.CHSApproverView.CHSEligibilityAmount;
        }
        if (value > hr2validAmount) { //30 Dec
          this.setState({
            ExpenseDetailsAlert: true
          });
        }
        else {
          this.setState({
            ExpenseDetailsAlert: false
          });
        }
        break;

      default:
        break;
    }
  }

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

  private getDependentOptionsByDesignation = () => {
    return [
      { key: 'Self', text: 'Self' },
      { key: 'Spouse', text: 'Spouse' }
    ];
  };

  public ddlDependentTypeChange = (e) => {

    const isMDorDMD =
      (this.state.DesignationTitle === "Managing Director (MD)" ||
        this.state.DesignationTitle === "Deputy Managing Director (DMD)");

    const LIMIT = parseFloat(this.state.Limit) || 0;
    if ((this.state.DesignationTitle != "Managing Director (MD)" &&
      this.state.DesignationTitle != "Deputy Managing Director (DMD)")) {
      if (LIMIT === 0) {
        return this.fail("Please Map Limit!");
      }
    }

    if (e === "Self") {
      this.setState({
        DependentType: e,
        IsSpouseEximEmployee: false,
        dropdownOptions: this.getDependentOptionsByDesignation(),
        AmountClaimed: "",
        CalculatedCHSEligibilityAmountLabel: "",
      });
      return;
    }

    // SPOUSE → existing validation remains
    if (e === "Spouse") {
      if (!isMDorDMD) {
        if (this.state.selectedOptionCHBx == null) {
          alert("Please check on 'Is Spouse Exim Employee'");
          this.setState({
            DependentType: null,
            dropdownOptions: this.getDependentOptionsByDesignation(),
            CalculatedCHSEligibilityAmountLabel: "",
            ShowLableElibleClaimAmount: false,
            AmountClaimed: ""
          });
          return false;
        }
      }

      this.setState({
        DependentType: e,
        IsSpouseEximEmployee: true,
        dropdownOptions: this.getDependentOptionsByDesignation(),
        AmountClaimed: "",
        CalculatedCHSEligibilityAmountLabel: ""
      });
    }
  };

  public EligibleClaimAmount = (e, amountClaimed) => {
    const isMDorDMD =
      (this.state.DesignationTitle === "Managing Director (MD)" ||
        this.state.DesignationTitle === "Deputy Managing Director (DMD)");

    if (this.state.selectedOptionCHBx != null) {
      let CalculatedCHSEligibilityAmountLabel;
      const LIMIT = this.state.Limit ? parseFloat(this.state.Limit) : 0;
      const scaleValue = this.state.Scale;
      if (this.state.EmployeeType != 'RETIRED') {
        if (this.state.Scale !== "GOVT") {
          const numberOnlyScale = parseFloat(scaleValue.match(/\d+$/)[0]);
          if (this.state.Age < 40 && numberOnlyScale <= 5) {
            let Scale5LIMIT = LIMIT ? LIMIT : 0;
            const Scale5LIMITSelf = (Scale5LIMIT) * 0.9;
            if (amountClaimed > LIMIT) {
              let calculatedAmountClaimed = (amountClaimed) * 0.9;
              if (calculatedAmountClaimed > LIMIT)
                CalculatedCHSEligibilityAmountLabel = LIMIT
              else
                CalculatedCHSEligibilityAmountLabel = calculatedAmountClaimed
            }
            else {
              if (amountClaimed > Scale5LIMITSelf) {
                CalculatedCHSEligibilityAmountLabel = Scale5LIMITSelf;
              }
              else {
                /////  CalculatedCHSEligibilityAmountLabel = amountClaimed
                CalculatedCHSEligibilityAmountLabel = (amountClaimed) * 0.9;
              }
            }

            if (e == 'Spouse') {
              if (amountClaimed > LIMIT) {
                let calculatedAmountClaimed = (amountClaimed) * 0.75;
                if (calculatedAmountClaimed > LIMIT)
                  CalculatedCHSEligibilityAmountLabel = LIMIT
                else
                  CalculatedCHSEligibilityAmountLabel = calculatedAmountClaimed
              }
              else {
                CalculatedCHSEligibilityAmountLabel = (amountClaimed) * 0.75;
              }

            }
          }
          if (this.state.Age >= 40 && numberOnlyScale <= 5) {
            let Scale5LIMIT = LIMIT ? LIMIT : 0;
            const Scale5LIMITSelf = (Scale5LIMIT) * 1;

            if (amountClaimed > LIMIT) {
              let calculatedAmountClaimed = (amountClaimed) * 1;
              if (calculatedAmountClaimed > LIMIT)
                CalculatedCHSEligibilityAmountLabel = LIMIT
              else
                CalculatedCHSEligibilityAmountLabel = calculatedAmountClaimed
            }
            else {
              if (amountClaimed > Scale5LIMITSelf) {
                CalculatedCHSEligibilityAmountLabel = Scale5LIMITSelf;
              }
              else {
                CalculatedCHSEligibilityAmountLabel = amountClaimed
              }
            }

            if (e == 'Spouse') {
              if (amountClaimed > LIMIT) {
                let calculatedAmountClaimed = (amountClaimed) * 0.75;
                if (calculatedAmountClaimed > LIMIT)
                  CalculatedCHSEligibilityAmountLabel = LIMIT
                else
                  CalculatedCHSEligibilityAmountLabel = calculatedAmountClaimed
              }
              else {
                CalculatedCHSEligibilityAmountLabel = (amountClaimed) * 0.75;
              }
            }
          }
          if (numberOnlyScale > 5) {
            const Scale5LIMITSelf = LIMIT ? LIMIT : 0;
            if (amountClaimed > LIMIT) {
              let calculatedAmountClaimed = (amountClaimed);
              if (calculatedAmountClaimed > LIMIT)
                CalculatedCHSEligibilityAmountLabel = LIMIT
              else
                CalculatedCHSEligibilityAmountLabel = calculatedAmountClaimed
            }
            else {
              if (amountClaimed > Scale5LIMITSelf) {
                CalculatedCHSEligibilityAmountLabel = Scale5LIMITSelf;
              }
              else {
                CalculatedCHSEligibilityAmountLabel = amountClaimed
              }
            }


            if (e == 'Spouse') {
              if (amountClaimed > LIMIT) {
                let calculatedAmountClaimed = (amountClaimed) * 0.75;
                if (calculatedAmountClaimed > LIMIT)
                  CalculatedCHSEligibilityAmountLabel = LIMIT
                else
                  CalculatedCHSEligibilityAmountLabel = calculatedAmountClaimed
              }
              else {
                CalculatedCHSEligibilityAmountLabel = (amountClaimed) * 0.75;
              }

            }
          }
        }
        else {
          const Scale5LIMITSelf = LIMIT ? LIMIT : 0;
          if (amountClaimed > LIMIT) {
            let calculatedAmountClaimed = (amountClaimed);
            if (calculatedAmountClaimed > LIMIT)
              CalculatedCHSEligibilityAmountLabel = LIMIT
            else
              CalculatedCHSEligibilityAmountLabel = calculatedAmountClaimed
          }
          else {
            if (amountClaimed > Scale5LIMITSelf) {
              CalculatedCHSEligibilityAmountLabel = Scale5LIMITSelf;
            }
            else {
              CalculatedCHSEligibilityAmountLabel = amountClaimed
            }
          }

          if (e == 'Spouse') {
            if (amountClaimed > LIMIT) {
              let calculatedAmountClaimed = (amountClaimed) * 0.75;
              if (calculatedAmountClaimed > LIMIT)
                CalculatedCHSEligibilityAmountLabel = LIMIT
              else
                CalculatedCHSEligibilityAmountLabel = calculatedAmountClaimed
            }
            else {
              CalculatedCHSEligibilityAmountLabel = (amountClaimed) * 0.75;
            }


            //  }
          }
        }
      }
      if (this.state.EmployeeType == 'RETIRED') {
        const Scale5LIMITSelf = LIMIT ? LIMIT : 0;
        if (amountClaimed > LIMIT) {
          let calculatedAmountClaimed = (amountClaimed) * 0.9;
          if (calculatedAmountClaimed > LIMIT)
            CalculatedCHSEligibilityAmountLabel = LIMIT
          else
            CalculatedCHSEligibilityAmountLabel = calculatedAmountClaimed
        }
        else {
          if (amountClaimed > Scale5LIMITSelf) {
            CalculatedCHSEligibilityAmountLabel = Scale5LIMITSelf;
          }
          else {
            ////CalculatedCHSEligibilityAmountLabel = amountClaimed
            CalculatedCHSEligibilityAmountLabel = (amountClaimed) * 0.9
          }
        }

        if (e == 'Spouse') {
          if (amountClaimed > LIMIT) {
            let calculatedAmountClaimed = (amountClaimed) * 0.75;
            if (calculatedAmountClaimed > LIMIT)
              CalculatedCHSEligibilityAmountLabel = LIMIT
            else
              CalculatedCHSEligibilityAmountLabel = calculatedAmountClaimed
          }
          else {
            CalculatedCHSEligibilityAmountLabel = (amountClaimed) * 0.75;
          }

          // }
        }
      }

      let updatedOptions = [{ key: 'Self', text: 'Self' }];

      if ((e !== 'Spouse' && this.state.selectedOptionCHBx !== 'Yes') || e == 'Spouse') {
        updatedOptions.push({ key: 'Spouse', text: 'Spouse' });
      }

      this.setState({
        DependentType: e,
        CalculatedCHSEligibilityAmountLabel: CalculatedCHSEligibilityAmountLabel,
        ////   selectedOptionCHBx: null,  //AP 7/7/25
        ShowLableElibleClaimAmount: true,
        dropdownOptions: updatedOptions
      });
    }
    else {
      if (!isMDorDMD) {
        alert("Please check on 'Is Spouse Exim Employee' ");

        this.setState({
          DependentType: null,
          dropdownOptions: this.getDependentOptionsByDesignation(),
          CalculatedCHSEligibilityAmountLabel: "",  // Optional: clear limit display
          ShowLableElibleClaimAmount: false
        });

        return false;
      }
    }
  };

  private async loadDependentDocs(healthCheckupServiceId: number): Promise<void> {
    const web = Web(ENV_CONFIG.siteUrl);
    const titleFolder = `CHS000${healthCheckupServiceId}`;
    const result: any = {};

    // We try per dependent type folder
    for (const dep of (this.state.dependentitems || [])) {
      var depType = "";

      if (dep && dep.DependentType) {
        depType = dep.DependentType;
      }

      if (!depType) {
        continue;
      }

      const folderUrl = `${this.props.currentSPContext.pageContext.web.serverRelativeUrl}/HealthCheckupServiceDocuments/${titleFolder}/${depType}`;

      try {
        const files = await web.getFolderByServerRelativeUrl(folderUrl).files.select("Name", "ServerRelativeUrl")();
        result[depType] = files || [];
      } catch {
        result[depType] = [];
      }
    }

    this.setState({ dependentDocsByType: result });
  }

  public BtnSubmitRequest = async () => {
    this.setState({ isSubmitting: true });

    const isMDorDMD =
      (this.state.DesignationTitle === "Managing Director (MD)" ||
        this.state.DesignationTitle === "Deputy Managing Director (DMD)");

    try {
      if (this.state.DependentType === "Spouse" && this.state.selectedOptionCHBx == null && !isMDorDMD) {
        return this.fail("Please check on 'Is Spouse Exim Employee'");
      }

      if (!this.state.dependentitems.length) {
        return this.fail("Please Enter Claim Amount !!");
      }

      const LIMIT = parseFloat(this.state.Limit) || 0;
      if (this.state.DesignationTitle != "Managing Director (MD)" &&
        this.state.DesignationTitle != "Deputy Managing Director (DMD)") {
        if (LIMIT === 0) {
          return this.fail("Please Map Limit!");
        }
      }

      var allRowsHaveFile = true;
      var deps = this.state.dependentitems ? this.state.dependentitems : [];

      for (var i = 0; i < deps.length; i++) {
        if (!this.state.dependentRowFile || !this.state.dependentRowFile[i]) {
          allRowsHaveFile = false;
          break;
        }
      }

      if (!allRowsHaveFile) return this.fail("Please upload 1 document for each dependent row!");

      const currentDate = new Date();
      const currentYear = currentDate.getFullYear();
      const fyStart = new Date(currentDate.getMonth() < 3 ? currentYear - 1 : currentYear, 3, 1);
      const fyEnd = new Date(fyStart.getFullYear() + 1, 2, 31);

      const normalizeDate = (date: any) => {
        const d = new Date(date);
        d.setHours(0, 0, 0, 0);
        return d;
      };

      const existingRequests = await EmployeeOps().getEmployeeMasterById(this.props);
      const hasRequestThisYear = existingRequests.filter(
        (req) =>
          normalizeDate(req.Created) >= normalizeDate(fyStart) &&
          normalizeDate(req.Created) <= normalizeDate(fyEnd) &&
          req.EmployeeID === this.state.EmployeeID
      );

      var requestedTypesMap: any = {};
      var requestedTypes: string[] = [];

      for (var i = 0; i < (this.state.dependentitems || []).length; i++) {
        var depItem = this.state.dependentitems[i];
        var depType = depItem && depItem.DependentType
          ? ("" + depItem.DependentType).trim()
          : "";
        if (depType && !requestedTypesMap[depType]) {
          requestedTypesMap[depType] = true;
          requestedTypes.push(depType);
        }
      }

      var hasRequestedSelf = requestedTypes.indexOf("Self") !== -1;
      var hasRequestedSpouse = requestedTypes.indexOf("Spouse") !== -1;


      var existingHasSelf = false;
      var existingHasSpouse = false;
      var existingHasCombinedSelfSpouse = false;

      for (var r = 0; r < hasRequestThisYear.length; r++) {
        var existingreq = hasRequestThisYear[r];
        if (!existingreq || existingreq.Status === "Rejected") continue;

        var existingTypesMap: any = {};
        var existingTypes: string[] = [];

        // Prefer JSON rows
        if (existingreq.DependentClaimDetails) {
          try {
            var parsed = JSON.parse(existingreq.DependentClaimDetails);
            if (parsed && parsed.length) {
              for (var j = 0; j < parsed.length; j++) {
                var row = parsed[j];
                var t = row && row.DependentType
                  ? ("" + row.DependentType).trim()
                  : "";
                if (t && !existingTypesMap[t]) {
                  existingTypesMap[t] = true;
                  existingTypes.push(t);
                }
              }
            }
          } catch (e) { }
        }

        // Fallback to single DependentType
        if (!existingTypes.length && existingreq.DependentType) {
          existingTypes.push(existingreq.DependentType.trim());
        }

        var hasSelf = existingTypes.indexOf("Self") !== -1;
        var hasSpouse = existingTypes.indexOf("Spouse") !== -1;

        if (hasSelf) existingHasSelf = true;
        if (hasSpouse) existingHasSpouse = true;
        if (hasSelf && hasSpouse) existingHasCombinedSelfSpouse = true;
      }

      // Rule 1: Combined request already exists
      if (existingHasCombinedSelfSpouse) {
        if (hasRequestedSelf && !hasRequestedSpouse) {
          return this.fail(
            "A combined request for Self + Spouse is already submitted for this financial year. Separate Self request is not allowed."
          );
        }

        if (hasRequestedSpouse && !hasRequestedSelf) {
          return this.fail(
            "A combined request for Self + Spouse is already submitted for this financial year. Separate Spouse request is not allowed."
          );
        }

        if (hasRequestedSelf && hasRequestedSpouse) {
          return this.fail(
            "A combined request for Self + Spouse is already submitted for this financial year."
          );
        }
      }

      // Rule 2: Partial overlaps
      if (hasRequestedSelf && hasRequestedSpouse) {
        if (existingHasSelf && existingHasSpouse) {
          return this.fail(
            "Requests for both Self and Spouse are already submitted for this financial year."
          );
        }
        if (existingHasSelf) {
          return this.fail(
            "A request for Self is already submitted for this financial year. Please submit only Spouse."
          );
        }
        if (existingHasSpouse) {
          return this.fail(
            "A request for Spouse is already submitted for this financial year. Please submit only Self."
          );
        }
      } else if (hasRequestedSelf && existingHasSelf) {
        return this.fail("A request for Self is already submitted for this financial year.");
      } else if (hasRequestedSpouse && existingHasSpouse) {
        return this.fail("A request for Spouse is already submitted for this financial year.");
      }

      const spCrudObj = await useSPCRUD();
      const { CHSRequestItem, pivotTab, dashTab } = this.buildCHSRequestItem();

      const req = await spCrudObj.insertData("HealthCheckupService", CHSRequestItem, this.props);
      this.setState({ reqID: req.data.ID });

      await spCrudObj.updateData("HealthCheckupService", req.data.ID, {
        Title: "CHS000" + req.data.ID,
      }, this.props);

      await this.uploadDependentRowDocs(req.data.ID);
      alert("Comprehensive Health Checkup Request Submitted Successfully!");

      this.setState({ isSubmitting: false });
      this.closeDialog();

      // Redirect to Dashboard
      window.location.href =
        `${ENV_CONFIG.siteUrl}/SitePages/CHSModule.aspx?dashtab=${dashTab}&ptab=${pivotTab}`;

      return req;
    } catch (err) {
      console.error("Submit error:", err);
      this.setState({ isSubmitting: false });
      alert("An error occurred while submitting the request." + err);
    }
  };

  private fail(message: string) {
    this.setState({ isSubmitting: false });
    alert(message);
    return false;
  }

  private buildCHSRequestItem() {
    let pivotTab = "User";
    let dashTab = "Pending";
    const selfItem = this.state.dependentitems.find((i) => i.DependentType === "Self");
    const spouseItem = this.state.dependentitems.find((i) => i.DependentType === "Spouse");
    var actualEligibilityLimit = (selfItem && selfItem.CalculatedCHSEligibilityAmountLabel) ||
      (spouseItem && spouseItem.CalculatedCHSEligibilityAmountLabel) || 0;
    const dob = this.state.DateofBirth; // "25-6-1956"
    const [day, month, year] = dob.split("-");

    // Convert to YYYY-MM-DD
    const formatted = `${year}-${month.padStart(2, "0")}-${day.padStart(2, "0")}`;


    let item: any = {
      OnBehalf: this.state.OnBehalf,
      EmployeeID: this.state.EmployeeID,
      Scale: this.state.Scale,
      EmployeeType: "" + this.state.EmployeeType,
      DependentType: this.state.DependentType,
      Status: "Pending",
      HR1Response: "Pending with HR1",
      HR2Response: "Pending with HR2",
      EligibilityLimit: actualEligibilityLimit,
      Limit: +this.state.Limit,
      FinancialYear: this.state.CurrentFinancialYear,
      RequestorEmail: this.state.CompanyEmail,
      IsSpouseEximMember: this.state.selectedOptionCHBx,
      DateofBirth: new Date(formatted), //this.state.DateofBirth?new Date(this.state.DateofBirth):this.state.DateofBirth,
      Designation: this.state.DesignationTitle,
      Age: "" + this.state.Age,
      AmountClaimed: this.state.TotalAmountClaimed,
      CHSEligibilityAmount: this.state.CHSEligibilityAmount,
      EmployeeName: this.state.EmployeeName,
      DependentClaimDetails: JSON.stringify(this.state.dependentitems),
      HRRemarkForRetired: this.state.ExpenseDetails.HRRemarkForRetired,
    };

    // HR1 Flow
    if (this.state.ShowHR1Tab) {
      pivotTab = "HR1";
      dashTab = "Pending";

      if (this.state.OnBehalf === "Yes") {
        const hrApprovedAmount =
          this.state.EmployeeType === "RETIRED"
            ? this.state.CHSEligibilityAmount
            : this.state.ExpenseDetails.Amount;

        if (this.state.EmployeeType === "RETIRED") { //23 Dec
          item = {
            ...item,
            Status: "Pending",
            HR1Response: "Approved by HR1",
            HR2Response: "Pending with HR2",
            TagApproverResponse: "Pending with TagApprover",
            HR1ApproverNameId: this.state.Currentuser.Id,
            HR1ResponseDate: new Date(),
          };
        }
        else {
          item = {
            ...item,
            Status: "Pending",
            HR1Response: "Approved by HR1",
            HR1ApproverNameId: this.state.Currentuser.Id,
            HR1ResponseDate: new Date(),
          };
        }
      }
    }

    // HR2 Flow
    if (this.state.ShowHR2Tab) {
      pivotTab = "HR2";
      dashTab = "Pending";

      if (this.state.OnBehalf === "Yes") {
        const hrApprovedAmount =
          this.state.EmployeeType === "RETIRED"
            ? this.state.CHSEligibilityAmount
            : this.state.ExpenseDetails.Amount;

        if (this.state.EmployeeType === "RETIRED") { //23 Dec
          item = {
            ...item,
            Status: "Pending",
            HR1Response: "Approved by HR1",
            TagApproverResponse: "Pending with TagApprover",

            HR2ApproverNameId: this.state.Currentuser.Id,
            HR2ResponseDate: new Date(),
            HRApprovedAmount: +hrApprovedAmount,
          };
        }
      }
    }
    return { CHSRequestItem: item, pivotTab, dashTab };
  }

  public BtnApproveHR1Request = async () => {
    this.setState({ isApproving: true });
    if (this.state.ExpenseDetails) {
      if (this.state.ExpenseDetails.HR1Remarks == "" || this.state.ExpenseDetails.HR1Remarks == undefined) {
        alert('Please mention Remarks ');
        this.setState({ isApproving: false });
        return false;
      }

      if (this.state.ExpenseDetails.HR1FinalAmount == undefined || this.state.ExpenseDetails.HR1FinalAmount == null || this.state.ExpenseDetails.HR1FinalAmount == "" || this.state.ExpenseDetails.HR1FinalAmount == 0) {
        alert('Please Enter Final Amount ');
        this.setState({ isApproving: false });
        return false;
      }
    }
    else {
      alert('Please enter required details ');
      this.setState({ isApproving: false });
      return false;
    }


    if (this.state.CHSApproverView.Designation != "Managing Director (MD)" &&
      this.state.CHSApproverView.Designation != "Deputy Managing Director (DMD)") {
      if (this.state.ExpenseDetails.HR1FinalAmount > this.state.CHSApproverView.CHSEligibilityAmount) {
        alert('Final Claim Amount should be less than Eligible Limit! ');
        this.setState({ isApproving: false });
        return false;
      }
    }
    const spCrudObj = await useSPCRUD();

    var CHSRequestItem = {
      Status: "Pending",
      HR1Response: "Approved by HR1",
      HR1ApproverNameId: this.state.Currentuser.Id,
      HR2Response: "Pending with HR2",
      TagApproverResponse: "Pending with TagApprover",

      HR1ResponseDate: new Date(),
      HR1Remark: this.state.ExpenseDetails.HR1Remarks,
      FinalAmount: +this.state.ExpenseDetails.HR1FinalAmount
    };
    return await spCrudObj.updateData("HealthCheckupService", this.state.CHSApproverView.ID, CHSRequestItem, this.props).then(async (req) => {
      alert('Comprehensive Health Checkup Request Approved Successfully!');
      this.setState({ isApproving: false });
      this.closeDialog();
      window.location.href = `${ENV_CONFIG.siteUrl}/SitePages/CHSModule.aspx?dashtab=Approved&ptab=HR1`;  // Update the URL path to your dashboard route

      return req;
    });
  };
  public BtnApproveHR2Request = async () => {
    this.setState({ isApproving: true });

    if (this.state.CHSApproverView.Designation != "Managing Director (MD)" &&
      this.state.CHSApproverView.Designation != "Deputy Managing Director (DMD)") {
      if (this.state.ExpenseDetails.HR2FinalAmount > this.state.CHSApproverView.CHSEligibilityAmount) {
        alert('Final Claim Amount should be less than Eligible Limit! ');
        this.setState({ isApproving: false });
        return false;
      }
    }
    if (this.state.ExpenseDetails) {
      if (this.state.ExpenseDetails.HR2Remarks == "" || this.state.ExpenseDetails.HR2Remarks == undefined) {
        alert('Please mention Remarks ');
        this.setState({ isApproving: false });
        return false;
      }
      if (this.state.ExpenseDetails.HR2FinalAmount == undefined || this.state.ExpenseDetails.HR2FinalAmount == null || this.state.ExpenseDetails.HR2FinalAmount == "" || this.state.ExpenseDetails.HR2FinalAmount == 0) {
        alert('Please Enter Final Amount ');
        this.setState({ isApproving: false });
        return false;
      }
    }
    else {
      alert('Please enter required details ');
      this.setState({ isApproving: false });
      return false;
    }


    const spCrudObj = await useSPCRUD();
    var CHSRequestItem = {
      HR2ApproverNameId: this.state.Currentuser.Id,
      Status: "Pending",
      HR2Response: "Approved by HR2",
      TagApproverResponse: "Pending with TagApprover",

      HR2ResponseDate: new Date(),
      HRApprovedAmount: +this.state.ExpenseDetails.HR2FinalAmount,
      HR2Remark: this.state.ExpenseDetails.HR2Remarks
    };
    return await spCrudObj.updateData("HealthCheckupService", this.state.CHSApproverView.ID, CHSRequestItem, this.props).then(async (req) => {
      alert('Comprehensive Health Checkup Request Approved Successfully!');
      this.setState({ isApproving: false });
      this.closeDialog()
      window.location.href = `${ENV_CONFIG.siteUrl}/SitePages/CHSModule.aspx?dashtab=Approved&ptab=HR2`;  // Update the URL path to your dashboard route

      return req;
    });
  };
  public BtnApproveTagApproverRequest = async () => {
    this.setState({ isApproving: true });
    const spCrudObj = await useSPCRUD();
    var CHSRequestItem = {
      TagApproverNameId: this.state.Currentuser.Id,
      Status: "Approved",
      TagApproverResponse: "Approved by TagApprover",
      TagApproverResponseDate: new Date(),
      TagApproverRemark: this.state.ExpenseDetails.TagApproverRemarks
    };
    return await spCrudObj.updateData("HealthCheckupService", this.state.CHSApproverView.ID, CHSRequestItem, this.props).then(async (req) => {
      alert('Comprehensive Health Checkup Request Approved Successfully!');
      this.setState({ isApproving: false });
      this.closeDialog();
      window.location.href = `${ENV_CONFIG.siteUrl}/SitePages/CHSModule.aspx?dashtab=Approved&ptab=HR2`;  // Update the URL path to your dashboard route

      return req;
    });
  };
  public BtnRejectRequest = async (HRTYPE: string) => {
    this.setState({ isRejecting: true });
    const spCrudObj = await useSPCRUD();
    if (HRTYPE == 'HR1') {
      if (this.state.ExpenseDetails != undefined) {
        if (this.state.ExpenseDetails.HR1Remarks == undefined || this.state.ExpenseDetails.HR1Remarks == ""
        ) {
          alert('Please mention Remarks ');
          this.setState({ isRejecting: false });
          return false;
        }
      }
      else {
        alert('Please mention Remarks ');
        this.setState({ isRejecting: false });
        return false;
      }

      const CHSRequestItem = {
        Status: "Rejected",
        HR1Response: "Rejected by HR1",
        HR1ApproverNameId: this.state.Currentuser.Id,

        HR1Remark: this.state.ExpenseDetails.HR1Remarks
      };

      return await spCrudObj.updateData("HealthCheckupService", this.state.CHSApproverView.ID, CHSRequestItem, this.props).then(async (req) => {
        alert('Comprehensive Health Checkup Request Rejected Successfully!');
        this.setState({ isRejecting: false });
        this.closeDialog();
        window.location.href = `${ENV_CONFIG.siteUrl}/SitePages/CHSModule.aspx?dashtab=Rejected&ptab=HR1`;  // Update the URL path to your dashboard route

        return req;
      });

    }

    if (HRTYPE == 'HR2') {
      if (this.state.ExpenseDetails != undefined) {
        if (
          this.state.ExpenseDetails.HR2Remarks == undefined ||

          this.state.ExpenseDetails.HR2Remarks == "") {
          alert('Please mention Remarks ');
          this.setState({ isRejecting: false });
          return false;
        }
      }
      else {
        alert('Please mention Remarks ');
        this.setState({ isRejecting: false });
        return false;

      }

      const CHSRequestItem = {
        Status: "Rejected",
        HR2Response: "Rejected by HR2",
        HR2ApproverNameId: this.state.Currentuser.Id,

        HR2Remark: this.state.ExpenseDetails.HR2Remarks
      };
      return await spCrudObj.updateData("HealthCheckupService", this.state.CHSApproverView.ID, CHSRequestItem, this.props).then(async (req) => {
        alert('Comprehensive Health Checkup Request Rejected Successfully!');
        this.setState({ isRejecting: false });
        this.closeDialog();
        window.location.href = `${ENV_CONFIG.siteUrl}/SitePages/CHSModule.aspx?dashtab=Rejected&ptab=HR2`;  // Update the URL path to your dashboard route

        return req;
      });

    }

    if (HRTYPE == 'TagApprover') {
      if (
        this.state.ExpenseDetails.TagApproverRemarks == undefined ||

        this.state.ExpenseDetails.TagApproverRemarks == "") {
        alert('Please mention Remarks ');
        this.setState({ isRejecting: false });
        return false;
      }
      const CHSRequestItem = {
        Status: "Rejected",
        TagApproverResponse: "Rejected by TagApprover",
        TagApproverNameId: this.state.Currentuser.Id,

        TagApproverRemark: this.state.ExpenseDetails.TagApproverRemarks
      };
      return await spCrudObj.updateData("HealthCheckupService", this.state.CHSApproverView.ID, CHSRequestItem, this.props).then(async (req) => {
        alert('Comprehensive Health Checkup Request Rejected Successfully!');
        this.setState({ isRejecting: false });
        this.closeDialog();
        window.location.href = `${ENV_CONFIG.siteUrl}/SitePages/CHSModule.aspx?dashtab=Rejected&ptab=TagApprover`;  // Update the URL path to your dashboard route

        return req;
      });

    }

  };
 
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

  // --- On Behalf Dashboards (Admin roles only) ---
  public OnBehalfPendingDashboard = async () => {
    const data = await EmployeeOps().getOnBehalfPendingDashboard(this.props);
    this.setState({ OnBehalfDashboard: data });
  };

  public OnBehalfApprovedDashboards = async () => {
    const data = await EmployeeOps().getOnBehalfApprovedDashboard(this.props);
    this.setState({ OnBehalfApprovedDashboard: data });
  };

  public OnBehalfRejectedDashboards = async () => {
    const data = await EmployeeOps().getOnBehalfRejectedDashboard(this.props);
    this.setState({ OnBehalfRejectedDashboard: data });
  };

  public openOnBehalfInnerTab = (tabName: string): void => {
    localStorage.setItem("activeOnBehalfDashTab", tabName);
    this.setState({ activeOnBehalfDashTab: tabName });
  };

  public HR1ApprovePendingDashboard = async () => {
    return await EmployeeOps().HR1getApproveDashboard(this.props).then(UserPending => {
      this.setState({ HR1ApprPendingDashboard: UserPending, allDashboardDataHR1Pending: UserPending });
      return UserPending;
    });
  };

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

  public TagApproverPendingDashboard = async () => {
    return await EmployeeOps().TagApprovergetApproveDashboard(this.props).then(UserPending => {
      this.setState({ TagApproverPendingDashboard: UserPending, allDashboardDataTagApproverPending: UserPending });
      return UserPending;
    });
  };

  public TagApproverApprovedDashboards = async () => {
    return await EmployeeOps().TagApprovergetApproveApprovedDashboard(this.props).then(UserApproved => {
      this.setState({ TagApproverApprDashboard: UserApproved, allDashboardDataTagApproverApproved: UserApproved });
      return UserApproved;
    });
  };
  public TagApproverRejectedDashboards = async () => {
    return await EmployeeOps().TagApprovergetApproveRejectedDashboard(this.props).then(UserRejected => {
      this.setState({ TagApproverApproverRejectDashboard: UserRejected, allDashboardDataTagApproverRejected: UserRejected });
      return UserRejected;
    });
  };
  public getHR1Approver = async (Items) => {
    console.log(Items);
    const ApproverViewReqItems = Items;
    var NewTotal = 0;
    this.setState({
      CHSApproverView: ApproverViewReqItems,
      isDialogHR1: true,
      dependentitems: ApproverViewReqItems.DependentClaimDetails
    },
      async () => {
        await this.loadDependentDocs(ApproverViewReqItems.ID);
      })
  }
  public getHR2Approver = async (Items) => {
    console.log(Items);
    const ApproverViewReqItems = Items;
    var NewTotal = 0;
    this.setState({
      CHSApproverView: ApproverViewReqItems,
      isDialogHR2: true,
      dependentitems: ApproverViewReqItems.DependentClaimDetails
    },
      async () => {
        await this.loadDependentDocs(ApproverViewReqItems.ID);
      })
  }
  public getHR1ApproverView = async (Items) => {
    console.log(Items);
    const ApproverViewReqItems = Items;
    var NewTotal = 0;
    if (Items.OnBehalf == 'Yes') {
      this.setState({
        CHSApproverView: ApproverViewReqItems,
        isDialogViewHR1: true,
        dependentitems: ApproverViewReqItems.DependentClaimDetails,
        isOnBehalfandRetired: true
      },
        async () => {
          await this.loadDependentDocs(ApproverViewReqItems.ID);
        })
    }
    else {
      this.setState({
        CHSApproverView: ApproverViewReqItems,
        isDialogViewHR1: true,
        dependentitems: ApproverViewReqItems.DependentClaimDetails,
        isOnBehalfandRetired: false
      },
        async () => {
          await this.loadDependentDocs(ApproverViewReqItems.ID);
        })
    }


  }
  public getHR2ApproverView = async (Items) => {
    console.log(Items);
    const ApproverViewReqItems = Items;
    var NewTotal = 0;
    this.setState({
      CHSApproverView: ApproverViewReqItems,
      isDialogViewHR2: true,
      dependentitems: ApproverViewReqItems.DependentClaimDetails,
    },
      async () => {
        await this.loadDependentDocs(ApproverViewReqItems.ID);
      })
  }

  public getTagApprover = async (Items) => {
    console.log(Items);
    const ApproverViewReqItems = Items;
    var NewTotal = 0;

    this.setState({
      CHSApproverView: ApproverViewReqItems,
      isDialogTagApprover: true,
      dependentitems: ApproverViewReqItems.DependentClaimDetails
    },
      async () => {
        await this.loadDependentDocs(ApproverViewReqItems.ID);
      })
  }
  public getTagApproverView = async (Items) => {
    console.log(Items);
    const ApproverViewReqItems = Items;
    var NewTotal = 0;
    this.setState({
      CHSApproverView: ApproverViewReqItems,
      isDialogViewTagApprover: true,
      dependentitems: ApproverViewReqItems.DependentClaimDetails
    },
      async () => {
        await this.loadDependentDocs(ApproverViewReqItems.ID);
      })
  }

  handleAdd = () => {
    const hasSelfDependent = this.state.dependentitems.some(
      (item) => item.DependentType === "Self"
    );

    const hasSpouseDependent = this.state.dependentitems.some(
      (item) => item.DependentType === "Spouse"
    );

    if (this.state.dependentitems.length < 2) {
      const { DependentType, AmountClaimed, CalculatedCHSEligibilityAmountLabel, dependentitems } = this.state;

      if (hasSelfDependent && DependentType == "Self") {
        alert("You have already added claim amount for Self!");
        this.setState({
          DependentType: "",
          AmountClaimed: "",
          CalculatedCHSEligibilityAmountLabel: ""
        });
        return false;
      }
      if (hasSpouseDependent && DependentType == "Spouse") {
        alert("You have already added claim amount for Spouse!");
        this.setState({
          DependentType: "",
          AmountClaimed: "",
          CalculatedCHSEligibilityAmountLabel: ""
        });
        return false;
      }

      if (DependentType && AmountClaimed > 0) {

        const LIMIT = this.state.Limit ? parseFloat(this.state.Limit) : 0;

        if (this.state.DesignationTitle != "Managing Director (MD)" &&
          this.state.DesignationTitle != "Deputy Managing Director (DMD)") {
          if (LIMIT == 0) {
            alert('Please Map Limit !');
            this.setState({
              DependentType: "",
              AmountClaimed: "",
              CalculatedCHSEligibilityAmountLabel: ""
            });
            return false;
          }
        }

        this.setState({
          dependentitems: [...dependentitems, { DependentType, AmountClaimed, CalculatedCHSEligibilityAmountLabel }],
          DependentType: "",
          AmountClaimed: "",
          CalculatedCHSEligibilityAmountLabel: ""

        }, () => {
          const dependentTypesCSV = this.state.dependentitems
            .map(item => item.DependentType)
            .join(', ');

          const totalClaimed = this.state.dependentitems.reduce((sum, item) => {
            return sum + parseFloat(item.AmountClaimed || 0);
          }, 0);

          const totalClaimedCHSEligibilityAmount = this.state.dependentitems.reduce((sum, item) => {
            return sum + parseFloat(item.CalculatedCHSEligibilityAmountLabel || 0);
          }, 0);

          this.setState({
            TotalAmountClaimed: totalClaimed,
            CHSEligibilityAmount: totalClaimedCHSEligibilityAmount,
            DependentType: dependentTypesCSV
          });
        });
      }
      else {
        alert("Please select Dependent Type and enter Amount Claimed!");
      }
    }
    else {
      alert("You have already added claim amount for Self and Spouse!");
      this.setState({
        DependentType: "",
        AmountClaimed: "",
        CalculatedCHSEligibilityAmountLabel: ""
      });
    }
  };

  handleDelete = (index: number) => {
    const items = [...this.state.dependentitems];
    items.splice(index, 1);

    let totalClaimed = 0;
    let totalEligibility = 0;
    let dependentTypesCSV = "";

    for (let i = 0; i < items.length; i++) {
      const amt = parseFloat(items[i].AmountClaimed || 0);
      const elig = parseFloat(items[i].CalculatedCHSEligibilityAmountLabel || 0);

      totalClaimed += isNaN(amt) ? 0 : amt;
      totalEligibility += isNaN(elig) ? 0 : elig;
    }

    dependentTypesCSV = items.map(x => x.DependentType).join(", ");

    this.setState({
      dependentitems: items,
      TotalAmountClaimed: totalClaimed,
      CHSEligibilityAmount: totalEligibility,

      // ✅ keep dependent type list based on remaining rows
      DependentType: dependentTypesCSV,

      // ✅ optional: clear only calculated label display (not mandatory)
      CalculatedCHSEligibilityAmountLabel: 0
    });
  };

  handleDropdownChangeDashHR2Approved = (event, option) => {
    const selectedType = event.key;
    this.setState({ DependentType: selectedType }, () => {
      // this.filterDashboardData();

      const { allDashboardDataHR2Approved, DependentType } = this.state;

      ///  const filteredData = DependentType === "ALL" ? allDashboardDataHR2Approved : allDashboardDataHR2Approved.filter((item) => item.EmployeeType === DependentType);
      const filteredData = DependentType === "ALL" ? allDashboardDataHR2Approved : DependentType === "GOVT" ? allDashboardDataHR2Approved.filter((item) => item.EmployeeType === 'PERMANENT' && (item.Designation == 'Managing Director (MD)' || item.Designation == 'Deputy Managing Director (DMD)')) : allDashboardDataHR2Approved.filter((item) => item.EmployeeType === DependentType);

      this.setState({
        HR2ApproverApprDashboard: filteredData
      });

    });
  };

  handleDropdownChangeDashHR2Rejected = (event, option) => {
    const selectedType = event.key;
    this.setState({ DependentType: selectedType }, () => {
      // this.filterDashboardData();

      const { allDashboardDataHR2Rejected, DependentType } = this.state;

      ////const filteredData = DependentType === "ALL" ? allDashboardDataHR2Rejected : allDashboardDataHR2Rejected.filter((item) => item.EmployeeType === DependentType);
      const filteredData = DependentType === "ALL" ? allDashboardDataHR2Rejected : DependentType === "GOVT" ? allDashboardDataHR2Rejected.filter((item) => item.EmployeeType === 'PERMANENT' && (item.Designation == 'Managing Director (MD)' || item.Designation == 'Deputy Managing Director (DMD)')) : allDashboardDataHR2Rejected.filter((item) => item.EmployeeType === DependentType);

      this.setState({
        HR2ApproverRejectDashboard: filteredData
      });

    });
  };


  handleDropdownChangeDashTagApproverApproved = (event, option) => {
    const selectedType = event.key;
    this.setState({ DependentType: selectedType }, () => {
      const { allDashboardDataTagApproverApproved, DependentType } = this.state;
      const filteredData = DependentType === "ALL" ? allDashboardDataTagApproverApproved : DependentType === "GOVT" ? allDashboardDataTagApproverApproved.filter((item) => item.EmployeeType === 'PERMANENT' && (item.Designation == 'Managing Director (MD)' || item.Designation == 'Deputy Managing Director (DMD)')) : allDashboardDataTagApproverApproved.filter((item) => item.EmployeeType === DependentType);
      this.setState({
        TagApproverApprDashboard: filteredData
      });

    });
  };

  handleDropdownChangeDashTagApproverRejected = (event, option) => {
    const selectedType = event.key;
    this.setState({ DependentType: selectedType }, () => {
      const { allDashboardDataTagApproverRejected, DependentType } = this.state;
      const filteredData = DependentType === "ALL" ? allDashboardDataTagApproverRejected : DependentType === "GOVT" ? allDashboardDataTagApproverRejected.filter((item) => item.EmployeeType === 'PERMANENT' && (item.Designation == 'Managing Director (MD)' || item.Designation == 'Deputy Managing Director (DMD)')) : allDashboardDataTagApproverRejected.filter((item) => item.EmployeeType === DependentType);
      this.setState({
        TagApproverApproverRejectDashboard: filteredData
      });

    });
  };

  handleDropdownChangeDashHR2Pending = (event, option) => {
    const selectedType = event.key;
    this.setState({ DependentType: selectedType }, () => {
      const { allDashboardDataHR2Pending, DependentType } = this.state;
      const filteredData = DependentType === "ALL" ? allDashboardDataHR2Pending : DependentType === "GOVT" ? allDashboardDataHR2Pending.filter((item) => item.EmployeeType === 'PERMANENT' && (item.Designation == 'Managing Director (MD)' || item.Designation == 'Deputy Managing Director (DMD)')) : allDashboardDataHR2Pending.filter((item) => item.EmployeeType === DependentType);
      this.setState({
        HR2ApprPendingDashboard: filteredData
      });

    });
  };

  handleDropdownChangeDashTagApproverPending = (event, option) => {
    const selectedType = event.key;
    this.setState({ DependentType: selectedType }, () => {
      const { allDashboardDataTagApproverPending, DependentType } = this.state;
      const filteredData = DependentType === "ALL" ? allDashboardDataTagApproverPending : DependentType === "GOVT" ? allDashboardDataTagApproverPending.filter((item) => item.EmployeeType === 'PERMANENT' && (item.Designation == 'Managing Director (MD)' || item.Designation == 'Deputy Managing Director (DMD)')) : allDashboardDataTagApproverPending.filter((item) => item.EmployeeType === DependentType);
      this.setState({
        TagApproverPendingDashboard: filteredData
      });

    });
  };

  handleDropdownChangeDashHR1Approved = (event, option) => {
    const selectedType = event.key;
    this.setState({ DependentType: selectedType }, () => {
      const { allDashboardDataHR1Approved, DependentType } = this.state;
      const filteredData = DependentType === "ALL" ? allDashboardDataHR1Approved : DependentType === "GOVT" ? allDashboardDataHR1Approved.filter((item) => item.EmployeeType === 'PERMANENT' && (item.Designation == 'Managing Director (MD)' || item.Designation == 'Deputy Managing Director (DMD)')) : allDashboardDataHR1Approved.filter((item) => item.EmployeeType === DependentType);
      this.setState({
        HR1ApproverApprDashboard: filteredData
      });

    });
  };

  handleDropdownChangeDashHR1Rejected = (event, option) => {
    const selectedType = event.key;
    this.setState({ DependentType: selectedType }, () => {
      const { allDashboardDataHR1Rejected, DependentType } = this.state;
      const filteredData = DependentType === "ALL" ? allDashboardDataHR1Rejected : DependentType === "GOVT" ? allDashboardDataHR1Rejected.filter((item) => item.EmployeeType === 'PERMANENT' && (item.Designation == 'Managing Director (MD)' || item.Designation == 'Deputy Managing Director (DMD)')) : allDashboardDataHR1Rejected.filter((item) => item.EmployeeType === DependentType);
      this.setState({
        HR1ApproverRejectDashboard: filteredData
      });

    });
  };

  handleDropdownChangeDashHR1Pending = (event, option) => {
    const selectedType = event.key;
    this.setState({ DependentType: selectedType }, () => {
      const { allDashboardDataHR1Pending, DependentType } = this.state;
      const filteredData = DependentType === "ALL" ? allDashboardDataHR1Pending : DependentType === "GOVT" ? allDashboardDataHR1Pending.filter((item) => item.EmployeeType === 'PERMANENT' && (item.Designation == 'Managing Director (MD)' || item.Designation == 'Deputy Managing Director (DMD)')) : allDashboardDataHR1Pending.filter((item) => item.EmployeeType === DependentType);
      this.setState({
        HR1ApprPendingDashboard: filteredData
      });

    });
  };

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

  handleCheckboxChange = (option: string) => (
    event?: React.FormEvent<HTMLElement | HTMLInputElement>,
    checked?: boolean
  ): boolean => {
    const LIMIT = parseFloat(this.state.Limit) || 0;
    if (this.state.DesignationTitle != "Managing Director (MD)" &&
      this.state.DesignationTitle != "Deputy Managing Director (DMD)") {
      if (LIMIT === 0) {
        return this.fail("Please Map Limit!");
      }
    }

    let updatedOptions = [{ key: 'Self', text: 'Self' }];

    this.setState({
      dependentitems: [],
      DependentType: null,
      dropdownOptions: updatedOptions,
      CalculatedCHSEligibilityAmountLabel: "",  // Optional: clear limit display
      ShowLableElibleClaimAmount: false,
      TotalAmountClaimed: 0
    })


    if (checked) {
      this.setState({
        selectedOptionCHBx: option
      });

      if (option == "Yes") {
        this.setState({
          IsSpouseEximEmployee: true
        });
      }
      else {
        this.setState({
          IsSpouseEximEmployee: false
        });
        updatedOptions.push({ key: 'Spouse', text: 'Spouse' });
      }
    } else {
      this.setState({
        selectedOptionCHBx: null,
        IsSpouseEximEmployee: false
      });
      updatedOptions.push({ key: 'Spouse', text: 'Spouse' });
    }

    this.setState({
      dropdownOptions: updatedOptions
    });
  };
  private openHR2InnerTab = (tabName: string): void => {
    localStorage.setItem("activeHR2Tab", tabName);
    this.setState({ activeHR2Tab: tabName });

  };

  private openTagApproverInnerTab = (tabName: string): void => {
    localStorage.setItem("activeTagApproverTab", tabName);
    this.setState({ activeTagApproverTab: tabName });

  };

  private openHR1InnerTab = (tabName: string): void => {
    localStorage.setItem("activeHR1Tab", tabName);
    this.setState({ activeHR1Tab: tabName });

  };
  private openHRInnerTab = (tabName: string): void => {
    console.log("Switching to tab:", tabName);
    this.setState({ activeHRdashTab: tabName }, () => {
      console.log("Tab set to:", this.state.activeHRdashTab);
      localStorage.setItem("activeHRdashTab", tabName);
    });
  };

  private handleRemarksPaste = (e: React.ClipboardEvent<HTMLInputElement | HTMLTextAreaElement>) => {
    const paste = e.clipboardData.getData('text');
    const invalidHtmlChars = /[<>&"'\/]/;

    if (invalidHtmlChars.test(paste)) {
      e.preventDefault();
      alert('Pasted text contains invalid characters like <, >, &, ", \', /.')
    }
  };

  public render(): React.ReactElement<IChsModuleProps> {
    const { selectedOption } = this.state;
    const value = selectedOption;
    // const valueq = selectedOptionCHBx ;  

    const { activeTab } = this.state;

    const showReportButton = this.state.ShowHR1Tab || this.state.ShowHR2Tab;

    const isMDorDMD =
      ((this.state.DesignationTitle === "Managing Director (MD)" ||
        this.state.DesignationTitle === "Deputy Managing Director (DMD)") && this.state.EmployeeType == "PERMANENT");

    const isMDorDMDApprover =
      (this.state.CHSApproverView.Designation === "Managing Director (MD)" ||
        this.state.CHSApproverView.Designation === "Deputy Managing Director (DMD)" && this.state.EmployeeType == "PERMANENT");
    return (

      <div className={styles.pettyCash} >
        <div className="" style={{ border: '1px solid #ddd', borderRadius: '4px' }}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}> </span>
              <div className='mb-2 text-right'>
                <DefaultButton className='btn-primary' onClick={() => this.CreateRequest()}> <Icon iconName='' ></Icon> Create Request </DefaultButton>
                <PrimaryButton className='btn-primary' style={{ marginLeft: "10px" }} text="Report for Age-based Health Checkup" onClick={this.exportToExcel} hidden={!showReportButton} />
              </div>

              <Pivot selectedKey={this.state.selectedOuterTab || "User"}
                onLinkClick={(item) => this.setState({ selectedOuterTab: item.props.itemKey })}
                linkSize={PivotLinkSize.large} linkFormat={PivotLinkFormat.tabs}>

                <PivotItem headerText="User Dashboard" itemKey="User" hidden={!this.state.ShowHRTab} className="tab-box"  >
                  <div className="row">
                    <div className={`${styles.tabnav} col-md-2`}>
                      <button
                        className={`tablink ${this.state.activeHRdashTab === 'Pending' ? 'active' : ''}`}
                        onClick={() => this.openHRInnerTab('Pending')}
                      >
                        Pending
                      </button>
                      <button
                        className={`tablink ${this.state.activeHRdashTab === 'Approved' ? 'active' : ''}`}
                        onClick={() => this.openHRInnerTab('Approved')}
                      >
                        Approved
                      </button>
                      <button
                        className={`tablink ${this.state.activeHRdashTab === 'Rejected' ? 'active' : ''}`}
                        onClick={() => this.openHRInnerTab('Rejected')}
                      >
                        Rejected
                      </button>
                    </div>

                    <div className="col-md-10 panelbodybox">
                      {/* PENDING TAB */}
                      <div className={`tabcontent ${this.state.activeHRdashTab === 'Pending' ? 'active' : ''}`} id="UserPending">

                        <table className="table ">
                          <tr>
                            <th>View</th>
                            <th>CHS ID</th>
                            <th>EmployeeID</th>
                            <th>EmployeeName</th>
                            <th>Employee Type</th>
                            <th>Date of Birth</th>
                            <th>Dependent Type</th>
                            <th>Claimed Amount</th>
                            <th>Final Approved Amount</th>
                            <th>Financial Year</th>
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
                                  <td>{items.HRApprovedAmount === null || items.HRApprovedAmount === undefined || items.HRApprovedAmount === '' ? 0 : items.HRApprovedAmount}</td>
                                  <td>{items.FinancialYear}</td>

                                  <td>{items.Status}</td>

                                </tr>
                              )
                            })
                              : ""
                          }
                        </table>
                      </div>

                      {/* APPROVED TAB */}
                      <div className={` tabcontent ${this.state.activeHRdashTab === 'Approved' ? 'active' : ''}`} id="UserApproved">

                        <table className="table ">
                          <tr>
                            <th>View</th>
                            <th>CHS ID</th>
                            <th>EmployeeID</th>
                            <th>EmployeeName</th>
                            <th>Employee Type</th>
                            <th>Date of Birth</th>
                            <th>Dependent Type</th>
                            <th>Claimed Amount</th>
                            <th>Final Approved Amount</th>
                            <th>Financial Year</th>
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
                                  <td>{items.HRApprovedAmount === null || items.HRApprovedAmount === undefined || items.HRApprovedAmount === '' ? 0 : items.HRApprovedAmount}</td>
                                  <td>{items.FinancialYear}</td>
                                  <td>{items.Status}</td>
                                </tr>
                              )
                            })
                              : ""
                          }
                        </table>
                      </div>

                      {/* REJECTED TAB */}
                      <div className={`tabcontent ${this.state.activeHRdashTab === 'Rejected' ? 'active' : ''}`} id="UserRejected">

                        <table className="table ">
                          <tr>
                            <th>View</th>
                            <th>CHS ID</th>
                            <th>EmployeeID</th>
                            <th>EmployeeName</th>
                            <th>Employee Type</th>
                            <th>Date of Birth</th>
                            <th>Dependent Type</th>
                            <th>Claimed Amount</th>
                            <th>Final Approved Amount</th>
                            <th>Financial Year</th>
                            <th>Status</th>
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
                                  <td>{items.HRApprovedAmount === null || items.HRApprovedAmount === undefined || items.HRApprovedAmount === '' ? 0 : items.HRApprovedAmount}</td>
                                  <td>{items.FinancialYear}</td>
                                  <td>{items.Status}</td>
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

                <PivotItem
                  headerText="On Behalf Dashboard"
                  itemKey="OnBehalfDashboard"
                  hidden={!this.state.isAdminOnBehalfUser}
                  className="tab-box"
                >
                  <div className="row">
                    <div className={`${styles.tabnav} col-md-2`}>
                      <button
                        className={`tablink ${this.state.activeOnBehalfDashTab === 'Pending' ? 'active' : ''}`}
                        onClick={() => this.openOnBehalfInnerTab('Pending')}
                      >
                        Pending
                      </button>
                      <button
                        className={`tablink ${this.state.activeOnBehalfDashTab === 'Approved' ? 'active' : ''}`}
                        onClick={() => this.openOnBehalfInnerTab('Approved')}
                      >
                        Approved
                      </button>
                      <button
                        className={`tablink ${this.state.activeOnBehalfDashTab === 'Rejected' ? 'active' : ''}`}
                        onClick={() => this.openOnBehalfInnerTab('Rejected')}
                      >
                        Rejected
                      </button>
                    </div>

                    <div className="col-md-10 panelbodybox">
                      {/* PENDING TAB */}
                      <div className={`tabcontent ${this.state.activeOnBehalfDashTab === 'Pending' ? 'active' : ''}`} id="OnBehalfPending">
                        <table className="table ">
                          <tr>
                            <th>View</th>
                            <th>CHS ID</th>
                            <th>EmployeeID</th>
                            <th>EmployeeName</th>
                            <th>Employee Type</th>
                            <th>Date of Birth</th>
                            <th>Dependent Type</th>
                            <th>Claimed Amount</th>
                            <th>Final Approved Amount</th>
                            <th>Financial Year</th>
                            <th>Status</th>
                          </tr>
                          {
                            this.state.OnBehalfDashboard.length > 0 ? this.state.OnBehalfDashboard.map((items) => {
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
                                  <td>{items.HRApprovedAmount === null || items.HRApprovedAmount === undefined || items.HRApprovedAmount === '' ? 0 : items.HRApprovedAmount}</td>
                                  <td>{items.FinancialYear}</td>
                                  <td>{items.Status}</td>

                                </tr>
                              )
                            })
                              : ""
                          }
                        </table>
                      </div>

                      {/* APPROVED TAB */}
                      <div className={` tabcontent ${this.state.activeOnBehalfDashTab === 'Approved' ? 'active' : ''}`} id="OnBehalfApproved">

                        <table className="table ">
                          <tr>
                            <th>View</th>
                            <th>CHS ID</th>
                            <th>EmployeeID</th>
                            <th>EmployeeName</th>
                            <th>Employee Type</th>
                            <th>Date of Birth</th>
                            <th>Dependent Type</th>
                            <th>Claimed Amount</th>
                            <th>Final Approved Amount</th>
                            <th>Financial Year</th>
                            <th>Status</th>

                          </tr>
                          {
                            this.state.OnBehalfApprovedDashboard.length > 0 ? this.state.OnBehalfApprovedDashboard.map((items) => {
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
                                  <td>{items.HRApprovedAmount === null || items.HRApprovedAmount === undefined || items.HRApprovedAmount === '' ? 0 : items.HRApprovedAmount}</td>
                                  <td>{items.FinancialYear}</td>
                                  <td>{items.Status}</td>
                                </tr>
                              )
                            })
                              : ""
                          }
                        </table>
                      </div>
                      {/* REJECTED TAB */}
                      <div className={`tabcontent ${this.state.activeOnBehalfDashTab === 'Rejected' ? 'active' : ''}`} id="OnBehalfRejected">
                        <table className="table ">
                          <tr>
                            <th>View</th>
                            <th>CHS ID</th>
                            <th>EmployeeID</th>
                            <th>EmployeeName</th>
                            <th>Employee Type</th>
                            <th>Date of Birth</th>
                            <th>Dependent Type</th>
                            <th>Claimed Amount</th>
                            <th>Final Approved Amount</th>
                            <th>Financial Year</th>
                            <th>Status</th>
                          </tr>
                          {
                            this.state.OnBehalfRejectedDashboard.length > 0 ? this.state.OnBehalfRejectedDashboard.map((items) => {
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
                                  <td>{items.HRApprovedAmount === null || items.HRApprovedAmount === undefined || items.HRApprovedAmount === '' ? 0 : items.HRApprovedAmount}</td>
                                  <td>{items.FinancialYear}</td>
                                  <td>{items.Status}</td>
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

                <PivotItem headerText="HR1 Approver Dashboard" itemKey="HR1" hidden={!this.state.ShowHR1Tab} className="tab-box"  >
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
                      <div className={`tabcontent ${this.state.activeHR1Tab === 'Pending' ? 'active' : ''}`} id="HR1Pending">

                        <div className="col-md-5 plr-5">
                          <div className='row'>
                            <div className="col-md-4  pl-0" style={{ paddingTop: '6px' }}>
                              <Label className="control-Label font-weight-bold" style={{ display: 'inline-block' }}>Employee Type </Label>
                            </div>
                            <div className="col-md-8">
                              <Dropdown
                                placeHolder="Select Employee Type"
                                options={[
                                  { key: "ALL", text: "ALL" },
                                  { key: "PERMANENT", text: "PERMANENT" },
                                  { key: "RETIRED", text: "RETIRED" },
                                  { key: "GOVT", text: "MD/DMD" }
                                ]}
                                selectedKey={this.state.DependentType}
                                onChanged={this.handleDropdownChangeDashHR1Pending}
                                className="dropdown-style "
                              />
                            </div>
                          </div>
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
                            <th>Claimed Amount</th>
                            <th>Final Approved Amount</th>
                            <th>Financial Year</th>
                            <th>Status</th>
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
                                  <td>{items.FinancialYear}</td>
                                  <td>{items.Status}</td>
                                </tr>
                              )
                            })
                              : ""
                          }
                        </table>
                      </div>

                      {/* APPROVED TAB */}
                      <div className={` tabcontent ${this.state.activeHR1Tab === 'Approved' ? 'active' : ''}`} id="HR1Approved">
                        <div className="col-md-5 plr-5">
                          <div className='row'>
                            <div className="col-md-4  pl-0" style={{ paddingTop: '6px' }}>
                              <Label className="control-Label font-weight-bold">Employee Type </Label>
                            </div>
                            <div className="col-md-8 ">
                              <Dropdown
                                placeHolder="Select Employee Type"
                                options={[
                                  { key: "ALL", text: "ALL" },
                                  { key: "PERMANENT", text: "PERMANENT" },
                                  { key: "RETIRED", text: "RETIRED" },
                                  { key: "GOVT", text: "MD/DMD" }
                                ]}
                                selectedKey={this.state.DependentType}
                                onChanged={this.handleDropdownChangeDashHR1Approved}
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
                            <th>Claimed Amount</th>
                            <th>Final Approved Amount</th>
                            <th>Financial Year</th>
                            <th>Status</th>
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
                                  <td>{items.FinancialYear}</td>
                                  <td>{items.Status}</td>
                                </tr>
                              )
                            })
                              : ""
                          }
                        </table>
                      </div>

                      {/* REJECTED TAB */}
                      <div className={`tabcontent ${this.state.activeHR1Tab === 'Rejected' ? 'active' : ''}`} id="HR1Rejected">
                        <div className="col-md-5 plr-5 ">
                          <div className="row">
                            <div className="col-md-4  pl-0" style={{ paddingTop: '6px' }}>
                              <Label className="control-Label font-weight-bold">Employee Type </Label>
                            </div>
                            <div className="col-md-8">
                              <Dropdown
                                placeHolder="Select Employee Type"
                                options={[
                                  { key: "ALL", text: "ALL" },
                                  { key: "PERMANENT", text: "PERMANENT" },
                                  { key: "RETIRED", text: "RETIRED" },
                                  { key: "GOVT", text: "MD/DMD" }
                                ]}
                                selectedKey={this.state.DependentType}
                                onChanged={this.handleDropdownChangeDashHR1Rejected}
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
                            <th>Claimed Amount</th>
                            <th>Final Approved Amount</th>
                            <th>Financial Year</th>
                            <th>Status</th>
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
                                  <td>{items.FinancialYear}</td>
                                  <td>{items.Status}</td>
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
                <PivotItem headerText="HR2 Approver Dashboard" itemKey="HR2" hidden={!this.state.ShowHR2Tab} className="tab-box"  >

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
                      <div className={`tabcontent ${this.state.activeHR2Tab === 'Pending' ? 'active' : ''}`} id="HR2Pending">
                        <div className="col-md-5 plr-5">
                          <div className="row">
                            <div className="col-md-4  pl-0" style={{ paddingTop: '6px' }}>
                              <Label className="control-Label font-weight-bold" style={{ display: 'inline-block' }}>Employee Type </Label>
                            </div>
                            <div className="col-md-8">
                              <Dropdown
                                placeHolder="Select Employee Type"
                                options={[
                                  { key: "ALL", text: "ALL" },
                                  { key: "PERMANENT", text: "PERMANENT" },
                                  { key: "RETIRED", text: "RETIRED" },
                                  { key: "GOVT", text: "MD/DMD" }
                                ]}
                                selectedKey={this.state.DependentType}
                                onChanged={this.handleDropdownChangeDashHR2Pending}
                                className="dropdown-style"
                              />
                            </div>

                          </div>
                        </div>

                        <div className="col-md-3" style={{ paddingTop: '6px' }}>
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
                            <th>Claimed Amount</th>
                            <th>Final Approved Amount</th>
                            <th>Financial Year</th>
                            <th>Status</th>
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
                                  <td>{items.FinancialYear}</td>
                                  <td>{items.Status}</td>

                                </tr>
                              )
                            })
                              : ""
                          }
                        </table>
                      </div>

                      {/* APPROVED TAB */}
                      <div className={` tabcontent ${this.state.activeHR2Tab === 'Approved' ? 'active' : ''}`} id="HR2Approved">
                        <div className="col-md-5 plr-5">
                          <div className="row">
                            <div className="col-md-4  pl-0" style={{ paddingTop: '6px' }}>
                              <Label className="control-Label font-weight-bold" style={{ display: 'inline-block' }}>Employee Type </Label>
                            </div>
                            <div className="col-md-8">
                              <Dropdown
                                placeHolder="Select Employee Type"
                                options={[
                                  { key: "ALL", text: "ALL" },
                                  { key: "PERMANENT", text: "PERMANENT" },
                                  { key: "RETIRED", text: "RETIRED" },
                                  { key: "GOVT", text: "MD/DMD" }
                                ]}
                                selectedKey={this.state.DependentType}
                                onChanged={this.handleDropdownChangeDashHR2Approved}
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
                            <th>Claimed Amount</th>
                            <th>Final Approved Amount</th>
                            <th>Financial Year</th>
                            <th>Status</th>
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
                                  <td>{items.FinancialYear}</td>

                                  <td>{items.Status}</td>

                                </tr>
                              )
                            })
                              : ""
                          }
                        </table>
                      </div>

                      {/* REJECTED TAB */}
                      <div className={`tabcontent ${this.state.activeHR2Tab === 'Rejected' ? 'active' : ''}`} id="HR2Rejected">
                        <div className="col-md-5 plr-5">
                          <div className='row'>
                            <div className="col-md-4  pl-0" style={{ paddingTop: '6px' }}>
                              <Label className="control-Label font-weight-bold" style={{ display: 'inline-block' }}>Employee Type </Label>
                            </div>
                            <div className="col-md-8">
                              <Dropdown
                                placeHolder="Select Employee Type"
                                options={[
                                  { key: "ALL", text: "ALL" },
                                  { key: "PERMANENT", text: "PERMANENT" },
                                  { key: "RETIRED", text: "RETIRED" },
                                  { key: "GOVT", text: "MD/DMD" }
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
                            <th>Claimed Amount</th>
                            <th>Final Approved Amount</th>
                            <th>Financial Year</th>
                            <th>Status</th>
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
                                  <td>{items.FinancialYear}</td>
                                  <td>{items.Status}</td>
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
                <PivotItem headerText="Tag Approver Dashboard" itemKey="TagApprover" hidden={!this.state.ShowTagApproverTab} className="tab-box"  >

                  <div className="row">
                    <div className={`${styles.tabnav} col-md-2`}>
                      <button
                        className={`tablink ${this.state.activeTagApproverTab === 'Pending' ? 'active' : ''}`}
                        onClick={() => this.openTagApproverInnerTab('Pending')}
                      >
                        Pending
                      </button>
                      <button
                        className={`tablink ${this.state.activeTagApproverTab === 'Approved' ? 'active' : ''}`}
                        onClick={() => this.openTagApproverInnerTab('Approved')}
                      >
                        Approved
                      </button>
                      <button
                        className={`tablink ${this.state.activeTagApproverTab === 'Rejected' ? 'active' : ''}`}
                        onClick={() => this.openTagApproverInnerTab('Rejected')}
                      >
                        Rejected
                      </button>
                    </div>

                    <div className="col-md-10 panelbodybox">
                      {/* PENDING TAB */}
                      <div className={`tabcontent ${this.state.activeTagApproverTab === 'Pending' ? 'active' : ''}`} id="TagPending">
                        <div className="col-md-5 plr-5">
                          <div className="row">
                            <div className="col-md-4  pl-0" style={{ paddingTop: '6px' }}>
                              <Label className="control-Label font-weight-bold" style={{ display: 'inline-block' }}>Employee Type </Label>
                            </div>
                            <div className="col-md-8">
                              <Dropdown
                                placeHolder="Select Employee Type"
                                options={[
                                  { key: "ALL", text: "ALL" },
                                  { key: "PERMANENT", text: "PERMANENT" },
                                  { key: "RETIRED", text: "RETIRED" },
                                  { key: "GOVT", text: "MD/DMD" }
                                ]}
                                selectedKey={this.state.DependentType}
                                onChanged={this.handleDropdownChangeDashTagApproverPending}
                                className="dropdown-style"
                              />
                            </div>

                          </div>
                        </div>

                        <div className="col-md-3" style={{ paddingTop: '6px' }}>

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
                            <th>Claimed Amount</th>
                            <th>Final Approved Amount</th>
                            <th>Financial Year</th>
                            <th>Status</th>
                            {/* <th>View Doc.</th> */}

                          </tr>
                          {
                            this.state.TagApproverPendingDashboard.length > 0 ? this.state.TagApproverPendingDashboard.map((items) => {
                              return (
                                <tr>
                                  <td><Icon iconName='View' onClick={() => this.getTagApproverView(items)} title='View' className={styles.iconcolor}></Icon></td>
                                  <td><Icon iconName='CheckMark' onClick={() => this.getTagApprover(items)} className={styles.iconcolor} title='Approver'></Icon></td>
                                  <td>{items.Title}</td>
                                  <td>{items.EmployeeID}</td>
                                  <td>{items.EmployeeName}</td>
                                  <td>{items.EmployeeType}</td>
                                  <td>{moment(items.DateofBirth).format("DD/MM/YYYY")}</td>
                                  <td>{items.DependentType}</td>
                                  <td>{items.AmountClaimed}</td>
                                  <td>{items.HRApprovedAmount}</td>
                                  <td>{items.FinancialYear}</td>


                                  <td>{items.Status}</td>

                                </tr>
                              )
                            })
                              : ""
                          }
                        </table>
                      </div>

                      {/* APPROVED TAB */}
                      <div className={` tabcontent ${this.state.activeTagApproverTab === 'Approved' ? 'active' : ''}`} id="TagApproved">
                        <div className="col-md-5 plr-5">
                          <div className="row">
                            <div className="col-md-4  pl-0" style={{ paddingTop: '6px' }}>
                              <Label className="control-Label font-weight-bold" style={{ display: 'inline-block' }}>Employee Type </Label>
                            </div>
                            <div className="col-md-8">
                              <Dropdown
                                placeHolder="Select Employee Type"
                                options={[
                                  { key: "ALL", text: "ALL" },
                                  { key: "PERMANENT", text: "PERMANENT" },
                                  { key: "RETIRED", text: "RETIRED" },
                                  { key: "GOVT", text: "MD/DMD" }
                                ]}
                                selectedKey={this.state.DependentType}
                                onChanged={this.handleDropdownChangeDashTagApproverApproved}
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
                            <th>Claimed Amount</th>
                            <th>Final Approved Amount</th>
                            <th>Financial Year</th>
                            <th>Status</th>
                          </tr>
                          {
                            this.state.TagApproverApprDashboard.length > 0 ? this.state.TagApproverApprDashboard.map((items) => {
                              return (
                                <tr>
                                  <td><Icon iconName='View' onClick={() => this.getTagApproverView(items)} title='View' className={styles.iconcolor}></Icon></td>
                                  <td>{items.Title}</td>
                                  <td>{items.EmployeeID}</td>
                                  <td>{items.EmployeeName}</td>
                                  <td>{items.EmployeeType}</td>
                                  <td>{moment(items.DateofBirth).format("DD/MM/YYYY")}</td>
                                  <td>{items.DependentType}</td>
                                  <td>{items.AmountClaimed}</td>
                                  <td>{items.HRApprovedAmount}</td>
                                  <td>{items.FinancialYear}</td>

                                  <td>{items.Status}</td>

                                </tr>
                              )
                            })
                              : ""
                          }
                        </table>
                      </div>

                      {/* REJECTED TAB */}
                      <div className={`tabcontent ${this.state.activeTagApproverTab === 'Rejected' ? 'active' : ''}`} id="TagRejected">
                        <div className="col-md-5 plr-5">
                          <div className='row'>
                            <div className="col-md-4  pl-0" style={{ paddingTop: '6px' }}>
                              <Label className="control-Label font-weight-bold" style={{ display: 'inline-block' }}>Employee Type </Label>
                            </div>
                            <div className="col-md-8">
                              <Dropdown
                                placeHolder="Select Employee Type"
                                options={[
                                  { key: "ALL", text: "ALL" },
                                  { key: "PERMANENT", text: "PERMANENT" },
                                  { key: "RETIRED", text: "RETIRED" },
                                  { key: "GOVT", text: "MD/DMD" }
                                ]}
                                selectedKey={this.state.DependentType}
                                onChanged={this.handleDropdownChangeDashTagApproverRejected}
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
                            <th>Claimed Amount</th>
                            <th>Final Approved Amount</th>
                            <th>Financial Year</th>
                            <th>Status</th>
                          </tr>
                          {
                            this.state.TagApproverApproverRejectDashboard.length > 0 ? this.state.TagApproverApproverRejectDashboard.map((items) => {
                              return (
                                <tr>
                                  <td><Icon iconName='View' onClick={() => this.getTagApproverView(items)} title='View' className={styles.iconcolor}></Icon></td>
                                  <td>{items.ID}</td>
                                  <td>{items.EmployeeID}</td>
                                  <td>{items.EmployeeName}</td>
                                  <td>{items.EmployeeType}</td>
                                  <td>{moment(items.DateofBirth).format("DD/MM/YYYY")}</td>
                                  <td>{items.DependentType}</td>
                                  <td>{items.AmountClaimed}</td>
                                  <td>{items.HRApprovedAmount}</td>
                                  <td>{items.FinancialYear}</td>

                                  <td>{items.Status}</td>
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
            title: 'Comprehensive Health Checkup Request',
            closeButtonAriaLabel: 'Close',
          }}
          containerClassName={'ms-dialogMainOverride ' + styles.textDialog}
        >
          <div className="card card-body">
            <div className="panel panel-default">
              <div className='panel-body'>
                <div className="row form-group" >
                  <div className="col-sm-2" >
                    <Label className="control-Label font-weight-bold">Financial Year:
                    </Label>
                  </div>
                  <div className="col-sm-6" >
                    {this.state.CurrentFinancialYear}
                  </div>
                </div>
                <div className="row form-group" hidden={(!this.state.ShowHR1Tab) && (!this.state.ShowHR2Tab) && (!this.state.isAdminOnBehalfUser)} >
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">On behalf of:</Label>
                  </div>
                  <div className="col-sm-4" >
                    <Dropdown
                      placeHolder="Select Option"
                      options={onbehalfoption}
                      selectedKey={this.state.OnBehalf}
                      onChanged={async (e, option) => {
                        if (e.key === "Yes") {
                          this.setState({
                            OnBehalf: e.key, DependentType: "",
                            CalculatedCHSEligibilityAmountLabel: "",
                            showhideEmployeeNameLab: true
                          });
                          if (this.state.isAdminOnBehalfUser) {
                            await this.loadAdminOnBehalfEmployees(this.state.EmployeeSubGroupIds);
                          } else {
                            this.getAllEmployee();
                          }
                        }
                        if (e.key === "No") {
                          this.getEmployee();
                          this.setState({
                            OnBehalf: e.key, DependentType: "",
                            CalculatedCHSEligibilityAmountLabel: "", showhideEmployeeNameLab: false
                          });
                        }
                      }}
                      className="dropdown-style"
                    />
                  </div>
                  <div className="col-sm-2" hidden={!this.state.showhideEmployeeNameLab}>
                    <Label className="control-Label font-weight-bold">Employee Name</Label>
                  </div>
                  <div className="col-sm-4" hidden={!this.state.showhideEmployeeNameLab}>
                    <Select
                      name="form-field-name"
                      autoFocus
                      clearable
                      value={selectedOption}
                      onChange={(selectedOption) => this.getSelectedEmployeeDetail(selectedOption)}
                      options={this.state.AllEmployeeCollObj}
                    />
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
                    <Label className="control-Label ">{this.state.EmployeeType}</Label>
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
                <div className="row form-group" style={{ display: isMDorDMD ? "none" : "block" }}>
                  <div className="col-sm-4" >
                    <Label className="control-Label font-weight-bold">Is Spouse an Exim Employee<span>*</span></Label>
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
                <div className="row form-group">

                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Age</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.Age}</Label>
                  </div>

                  <div className="col-sm-2" style={{ display: isMDorDMD ? "none" : "block" }}>
                    <Label className="control-Label font-weight-bold">CHS Limit</Label>
                  </div>
                  <div className="col-sm-2 d-flex align-items-center justify-content-between"
                    style={{ display: isMDorDMD ? "none" : "flex" }}>
                    <Label className="control-Label" style={{ display: isMDorDMD ? "none" : "block" }}>
                      {this.state.Limit}
                    </Label>
                  </div>

                </div>
                <div className="row form-group">
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Dependent Type </Label>
                  </div>
                  <div className="col-sm-2">
                    <Dropdown
                      placeHolder="Select Dependent"
                      options={this.state.dropdownOptions}
                      selectedKey={this.state.DependentType}
                      className="dropdown-style"
                      onChanged={(e, options) => {
                        this.ddlDependentTypeChange(e.key);
                      }}
                    />
                  </div>
                  <div className="col-sm-2" >
                    <Label className="control-Label font-weight-bold">Amount Claimed<span >*</span></Label>
                  </div>
                  <div className="col-sm-2">
                    <TextField type='number'
                      min='0'
                      name="ExpenseDetails.Amount" value={this.state.AmountClaimed}
                      onChanged={(e: any) => this.handleInputChangeadd(event)}></TextField>
                  </div>
                  <div
                    className="col-sm-2" style={{ display: isMDorDMD ? "none" : "block" }}>
                    <Label className="control-Label font-weight-bold">Calculated CHS Eligibility</Label>
                  </div>
                  <div
                    className="col-sm-2 d-flex align-items-center justify-content-between"
                    style={{ display: isMDorDMD ? "none" : "flex" }}>
                    <Label className="control-Label " style={{ display: isMDorDMD ? "none" : "block" }}>
                      {this.state.CalculatedCHSEligibilityAmountLabel}
                    </Label>
                  </div>
                  <PrimaryButton text="Add" onClick={this.handleAdd} />
                </div>
                {/* Table of Added Dependents */}
                {this.state.dependentitems.length > 0 && (
                  <div className="row form-group">
                    <div className="col-sm-12">
                      <table className="table table-bordered">
                        <thead>
                          <tr>
                            <th>Dependent Type</th>
                            <th>Amount Claimed</th>
                            <th style={{ display: isMDorDMD ? "none" : "table-cell" }}>Calculated CHS Eligibility</th>
                            <th>Attachment <br /><span
                              style={{
                                color: "red",
                                textAlign: "left",
                                fontSize: 10,
                              }}
                            > Note : File name  Allowed characters: letters, numbers, dots, underscores, and spaces.</span></th>
                            <th>Action</th>
                          </tr>
                        </thead>
                        <tbody>
                          {this.state.dependentitems.map((item, index) => (
                            <tr key={index}>
                              <td>{item.DependentType}</td>
                              <td>{item.AmountClaimed}</td>
                              <td style={{ display: isMDorDMD ? "none" : "table-cell" }}>{item.CalculatedCHSEligibilityAmountLabel}</td>
                              <td>
                                <input
                                  type="file"
                                  onChange={(ev) => this.handleDependentRowFileChange(index, ev as any)}
                                />
                                {this.state.dependentRowFile &&
                                  this.state.dependentRowFile[index] && (
                                    <div style={{ marginTop: 6, fontSize: 12 }}>
                                      {this.state.dependentRowFile[index].name}
                                    </div>
                                  )}
                              </td>
                              <td>
                                <button className="btn btn-danger btn-sm" onClick={() => this.handleDelete(index)}>
                                  Delete
                                </button>
                              </td>
                            </tr>
                          ))}
                        </tbody>
                        <tfoot>
                          <tr>
                            <td colSpan={1}><strong>Total</strong></td>
                            <td colSpan={1}>{this.state.TotalAmountClaimed}</td>
                            <td colSpan={1} style={{ display: isMDorDMD ? "none" : "table-cell" }}>{this.state.CHSEligibilityAmount}</td>
                            <td colSpan={1}></td>
                          </tr>
                        </tfoot>
                      </table>
                    </div>
                  </div>
                )}
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
              </div>
            </div>
            <div className="col-sm-12" style={{ padding: 0 }} hidden={(!this.state.showhideEmployeeNameLab && this.state.EmployeeType != "RETIRED") || this.state.isAdminOnBehalfUser} >
              <Label className="control-Label font-weight-bold col-md-2">HR Remarks For Retired Employee</Label>
              <TextField type='text' className='col-md-8'
                name="ExpenseDetails.HRRemarkForRetired"
                onChanged={(e: any) => this.handleInputChangeadd(event)}
                onPaste={this.handleRemarksPaste}
                onKeyPress={(e) => { if (/[<>&"'\/]/.test(e.key)) { e.preventDefault(); alert('Special characters like <, >, &, ", \', / are not allowed.'); } }}></TextField>
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
            <PrimaryButton onClick={() => this.closeDialog()} >Close</PrimaryButton>
          </div>
        </Dialog>
        <Dialog
          hidden={!this.state.isDialogHR1}
          onDismiss={this.closeDialog}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Comprehensive Health Checkup Approver',
            closeButtonAriaLabel: 'Close',
          }}
          containerClassName={'ms-dialogMainOverride ' + styles.textDialog}
        >
          <div className="card card-body">
            <div className="panel panel-default">
              <div className='panel-body'>

                <div className="row form-group">
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Finacial Year:</Label>
                  </div>
                  <div className="col-sm-6">
                    <Label className="control-Label">{this.state.CHSApproverView.FinancialYear}</Label>
                  </div>
                </div>
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
                    <Label className="control-Label font-weight-bold">Scale</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.Scale}</Label>
                  </div>
                </div>
                <div className="row form-group">
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
                <div className="row form-group" hidden={this.state.CHSApproverView.DependentType == 'Self' || isMDorDMDApprover} >
                  <div className="col-sm-4" >
                    <Label className="control-Label font-weight-bold">Is Spouse an Exim Employee<span>*</span></Label>
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
                <div className="row form-group">
                  <div className="col-sm-2" style={{ display: isMDorDMDApprover ? "none" : "block" }}>
                    <Label className="control-Label font-weight-bold">CHS Limit</Label>
                  </div>
                  <div className="col-sm-2" style={{ display: isMDorDMDApprover ? "none" : "flex" }}>
                    <Label className="control-Label ">{this.state.CHSApproverView.Limit}</Label>
                  </div>
                </div>
                {/* Table of Added Dependents */}
                {this.state.dependentitems.length > 0 && (
                  <div className="row form-group">
                    <div className="col-sm-12">
                      <table className="table table-bordered">
                        <thead>
                          <tr>
                            <th>Dependent Type</th>
                            <th>Amount Claimed</th>
                            <th style={{ display: isMDorDMDApprover ? "none" : "table-cell" }}>Calculated CHS Eligibility</th>
                            <th>Attachment </th>
                          </tr>
                        </thead>
                        <tbody>
                          {this.state.dependentitems.map((item, index) => (
                            <tr key={index}>
                              <td>{item.DependentType}</td>
                              <td>{item.AmountClaimed}</td>
                              <td style={{ display: isMDorDMDApprover ? "none" : "table-cell" }}>{item.CalculatedCHSEligibilityAmountLabel}</td>
                              <td>
                                {this.state.dependentDocsByType &&
                                  item &&
                                  item.DependentType &&
                                  this.state.dependentDocsByType[item.DependentType] &&
                                  this.state.dependentDocsByType[item.DependentType].length ? (
                                  this.state.dependentDocsByType[item.DependentType].map(function (f) {
                                    return (
                                      <a href={f.ServerRelativeUrl} target="_blank" rel="noreferrer">
                                        {f.Name}
                                      </a>
                                    );
                                  })
                                ) : (
                                  <span>-</span>
                                )}
                              </td>
                            </tr>
                          ))}
                        </tbody>
                        <tfoot>
                          <tr>
                            <td colSpan={1}><strong>Total</strong></td>
                            <td colSpan={1}>{this.state.CHSApproverView.AmountClaimed}</td>
                            <td colSpan={1} style={{ display: isMDorDMDApprover ? "none" : "table-cell" }}>{this.state.CHSApproverView.CHSEligibilityAmount}</td>
                          </tr>
                        </tfoot>
                      </table>
                    </div>
                  </div>
                )}
                <div className="row form-group">
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Final Amount</Label>
                  </div>
                  <div className="col-sm-2">
                    <TextField type='number'
                      name="ExpenseDetails.HR1FinalAmount"
                      onChanged={(e: any) => this.handleInputChangeadd(event)}></TextField>
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
            <div className="col-sm-12" style={{ padding: 0 }}>
              <Label className="control-Label font-weight-bold col-md-2">Remarks</Label>
              <TextField type='text' className='col-md-8'
                name="ExpenseDetails.HR1Remarks"
                onChanged={(e: any) => this.handleInputChangeadd(event)}
                onPaste={this.handleRemarksPaste}
                onKeyPress={(e) => { if (/[<>&"'\/]/.test(e.key)) { e.preventDefault(); alert('Special characters like <, >, &, ", \', / are not allowed.'); } }}></TextField>
            </div>
          </div>
          <div className='text-center'>
            <PrimaryButton className={styles.custombtn + " " + "mr-2"} onClick={() => this.BtnApproveHR1Request()} disabled={this.state.isApproving}>{this.state.isApproving ? <Spinner size={SpinnerSize.small} /> : "Approve"}</PrimaryButton>
            <PrimaryButton className={styles.custombtn + " " + "mr-2"} onClick={() => this.BtnRejectRequest('HR1')} disabled={this.state.isRejecting}>{this.state.isRejecting ? <Spinner size={SpinnerSize.small} /> : "Reject"}</PrimaryButton>
            <PrimaryButton onClick={() => this.closeDialog()} >Close</PrimaryButton>
          </div>
        </Dialog>
        <Dialog
          hidden={!this.state.isDialogViewHR1}
          onDismiss={this.closeDialog}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Comprehensive Health Checkup View Form',
            closeButtonAriaLabel: 'Close',
          }}
          containerClassName={'ms-dialogMainOverride ' + styles.textDialog}
        >
          <div className="card card-body">
            <div className="panel panel-default">
              <div className='panel-body'>
                <div className="row form-group">
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Finacial Year:</Label>
                  </div>
                  <div className="col-sm-6">
                    <Label className="control-Label">{this.state.CHSApproverView.FinancialYear}</Label>
                  </div>
                </div>
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
                    <Label className="control-Label font-weight-bold">Scale</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.Scale}</Label>
                  </div>
                </div>
                <div className="row form-group">
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
                <div className="row form-group" hidden={this.state.CHSApproverView.DependentType == 'Self' || isMDorDMDApprover} >
                  <div className="col-sm-4" >
                    <Label className="control-Label font-weight-bold">Is Spouse an Exim Employee<span>*</span></Label>
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
                <div className="row form-group">
                  <div className="col-sm-2" style={{ display: isMDorDMDApprover ? "none" : "block" }}>
                    <Label className="control-Label font-weight-bold">CHS Limit</Label>
                  </div>
                  <div className="col-sm-2" style={{ display: isMDorDMDApprover ? "none" : "flex" }}>
                    <Label className="control-Label ">{this.state.CHSApproverView.Limit}</Label>
                  </div>
                </div>
                {/* Table of Added Dependents */}
                {this.state.dependentitems.length > 0 && (
                  <div className="row form-group">
                    <div className="col-sm-12">
                      <table className="table table-bordered">
                        <thead>
                          <tr>
                            <th>Dependent Type</th>
                            <th>Amount Claimed</th>
                            <th style={{ display: isMDorDMDApprover ? "none" : "table-cell" }}>Calculated CHS Eligibility</th>
                            <th>Attachment</th>
                          </tr>
                        </thead>
                        <tbody>
                          {this.state.dependentitems.map((item, index) => (
                            <tr key={index}>
                              <td>{item.DependentType}</td>
                              <td>{item.AmountClaimed}</td>
                              <td style={{ display: isMDorDMDApprover ? "none" : "table-cell" }}>{item.CalculatedCHSEligibilityAmountLabel}</td>
                              <td>
                                {this.state.dependentDocsByType &&
                                  item &&
                                  item.DependentType &&
                                  this.state.dependentDocsByType[item.DependentType] &&
                                  this.state.dependentDocsByType[item.DependentType].length ? (
                                  this.state.dependentDocsByType[item.DependentType].map(function (f) {
                                    return (
                                      <a href={f.ServerRelativeUrl} target="_blank" rel="noreferrer">
                                        {f.Name}
                                      </a>
                                    );
                                  })
                                ) : (
                                  <span>-</span>
                                )}
                              </td>
                            </tr>
                          ))}
                        </tbody>
                        <tfoot>
                          <tr>
                            <td colSpan={1}><strong>Total</strong></td>
                            <td colSpan={1}>{this.state.CHSApproverView.AmountClaimed}</td>
                            <td colSpan={1} style={{ display: isMDorDMDApprover ? "none" : "table-cell" }}>{this.state.CHSApproverView.CHSEligibilityAmount}</td>
                          </tr>
                        </tfoot>
                      </table>
                    </div>
                  </div>
                )}
                <div className="row form-group">
                  <div className="col-sm-2" hidden={this.state.isOnBehalfandRetired && !this.state.isAdminOnBehalfUser}>
                    <Label className="control-Label font-weight-bold">HR1 Approved Amount</Label>
                  </div>
                  <div className="col-sm-2" hidden={this.state.isOnBehalfandRetired && !this.state.isAdminOnBehalfUser}>
                    <Label className="control-Label ">
                      {this.state.CHSApproverView.FinalAmount === null ||
                        this.state.CHSApproverView.FinalAmount === undefined ||
                        this.state.CHSApproverView.FinalAmount === ''
                        ? 0
                        : this.state.CHSApproverView.FinalAmount}
                    </Label>
                  </div>
                  <div className="col-sm-2" hidden={this.state.isOnBehalfandRetired && !this.state.isAdminOnBehalfUser} >
                    <Label className="control-Label font-weight-bold">HR1 Remarks</Label>
                  </div>
                  <div className="col-sm-6" hidden={this.state.isOnBehalfandRetired && !this.state.isAdminOnBehalfUser}>
                    <Label className="control-Label ">
                      {this.state.CHSApproverView.HR1Remark === null ||
                        this.state.CHSApproverView.HR1Remark === undefined ||
                        this.state.CHSApproverView.HR1Remark === ''
                        ? 'NA'
                        : this.state.CHSApproverView.HR1Remark}
                    </Label>
                  </div>
                </div>
                <div className="row form-group">
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Final Approved Amount</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label">
                      {this.state.CHSApproverView.HRApprovedAmount === null ||
                        this.state.CHSApproverView.HRApprovedAmount === undefined ||
                        this.state.CHSApproverView.HRApprovedAmount === ''
                        ? 0
                        : this.state.CHSApproverView.HRApprovedAmount}
                    </Label>
                  </div>
                  <div className="col-sm-2" hidden={!this.state.isOnBehalfandRetired && !this.state.isAdminOnBehalfUser} >
                    <Label className="control-Label font-weight-bold">HR2 Remarks</Label>
                  </div>
                  <div className="col-sm-6" hidden={!this.state.isOnBehalfandRetired && !this.state.isAdminOnBehalfUser}>
                    <Label className="control-Label ">
                      {this.state.CHSApproverView.HR2Remark === null ||
                        this.state.CHSApproverView.HR2Remark === undefined ||
                        this.state.CHSApproverView.HR2Remark === ''
                        ? 'NA'
                        : this.state.CHSApproverView.HR2Remark}
                    </Label>
                  </div>
                </div>

                <div className="row form-group">
                  <div className="col-sm-2" hidden={!this.state.isOnBehalfandRetired && !this.state.isAdminOnBehalfUser} >
                    <Label className="control-Label font-weight-bold">TAG Remarks</Label>
                  </div>
                  <div className="col-sm-6" hidden={!this.state.isOnBehalfandRetired && !this.state.isAdminOnBehalfUser}>
                    <Label className="control-Label ">
                      {this.state.CHSApproverView.TagApproverRemark === null ||
                        this.state.CHSApproverView.TagApproverRemark === undefined ||
                        this.state.CHSApproverView.TagApproverRemark === ''
                        ? 'NA'
                        : this.state.CHSApproverView.TagApproverRemark}
                    </Label>
                  </div>
                </div>


              </div>
            </div>
            <div className="panel panel-default" hidden={this.state.CHSApproverView.EmployeeType != "RETIRED"} >
              <div className='panel-body'>
                <div className="row form-group">
                  <div className="col-sm-2" hidden={this.state.CHSApproverView.EmployeeType != "RETIRED"} >
                    <Label className="control-Label font-weight-bold">HR Remarks For Retired Employee</Label>
                  </div>
                  <div className="col-sm-8" hidden={this.state.CHSApproverView.EmployeeType != "RETIRED"}>
                    <Label className="control-Label ">
                      {this.state.CHSApproverView.HRRemarkForRetired === null ||
                        this.state.CHSApproverView.HRRemarkForRetired === undefined ||
                        this.state.CHSApproverView.HRRemarkForRetired === ''
                        ? 'NA'
                        : this.state.CHSApproverView.HRRemarkForRetired}
                    </Label>
                  </div>
                </div>
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
            title: 'Comprehensive Health Checkup Approver',
            closeButtonAriaLabel: 'Close',
          }}
          containerClassName={'ms-dialogMainOverride ' + styles.textDialog}
        >
          <div className="card card-body">
            <div className="panel panel-default">
              <div className='panel-body'>
                <div className="row form-group">
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Finacial Year:</Label>
                  </div>
                  <div className="col-sm-6">
                    <Label className="control-Label">{this.state.CHSApproverView.FinancialYear}</Label>
                  </div>
                </div>
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
                    <Label className="control-Label font-weight-bold">Scale</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.Scale}</Label>
                  </div>
                </div>
                <div className="row form-group">
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
                <div className="row form-group" hidden={this.state.CHSApproverView.DependentType == 'Self' || isMDorDMDApprover} >
                  <div className="col-sm-4" >
                    <Label className="control-Label font-weight-bold">Is Spouse an Exim Employee<span>*</span></Label>
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
                <div className="row form-group">
                  <div className="col-sm-2" style={{ display: isMDorDMDApprover ? "none" : "block" }}>
                    <Label className="control-Label font-weight-bold">CHS Limit</Label>
                  </div>
                  <div className="col-sm-2" style={{ display: isMDorDMDApprover ? "none" : "flex" }}>
                    <Label className="control-Label ">{this.state.CHSApproverView.Limit}</Label>
                  </div>
                </div>
                {/* Table of Added Dependents */}
                {this.state.dependentitems.length > 0 && (
                  <div className="row form-group">
                    <div className="col-sm-12">
                      <table className="table table-bordered">
                        <thead>
                          <tr>
                            <th>Dependent Type</th>
                            <th>Amount Claimed</th>
                            <th style={{ display: isMDorDMDApprover ? "none" : "table-cell" }}>Calculated CHS Eligibility</th>
                            <th>Attachment</th>
                          </tr>
                        </thead>
                        <tbody>
                          {this.state.dependentitems.map((item, index) => (
                            <tr key={index}>
                              <td>{item.DependentType}</td>
                              <td>{item.AmountClaimed}</td>
                              <td style={{ display: isMDorDMDApprover ? "none" : "table-cell" }}>{item.CalculatedCHSEligibilityAmountLabel}</td>
                              <td>
                                {this.state.dependentDocsByType &&
                                  item &&
                                  item.DependentType &&
                                  this.state.dependentDocsByType[item.DependentType] &&
                                  this.state.dependentDocsByType[item.DependentType].length ? (
                                  this.state.dependentDocsByType[item.DependentType].map(function (f) {
                                    return (
                                      <a href={f.ServerRelativeUrl} target="_blank" rel="noreferrer">
                                        {f.Name}
                                      </a>
                                    );
                                  })
                                ) : (
                                  <span>-</span>
                                )}
                              </td>
                            </tr>
                          ))}
                        </tbody>
                        <tfoot>
                          <tr>
                            <td colSpan={1}><strong>Total</strong></td>
                            <td colSpan={1}>{this.state.CHSApproverView.AmountClaimed}</td>
                            <td colSpan={1} style={{ display: isMDorDMDApprover ? "none" : "table-cell" }}>{this.state.CHSApproverView.CHSEligibilityAmount}</td>
                          </tr>
                        </tfoot>
                      </table>
                    </div>
                  </div>
                )}
                <div className="row form-group">
                  <div className="col-sm-2" hidden={this.state.isOnBehalfandRetired && !this.state.isAdminOnBehalfUser}>
                    <Label className="control-Label font-weight-bold">HR1 Approved Amount</Label>
                  </div>
                  <div className="col-sm-2" hidden={this.state.isOnBehalfandRetired && !this.state.isAdminOnBehalfUser}>
                    <Label className="control-Label "> {this.state.CHSApproverView.FinalAmount === null ||
                      this.state.CHSApproverView.FinalAmount === undefined ||
                      this.state.CHSApproverView.FinalAmount === ''
                      ? 0
                      : this.state.CHSApproverView.FinalAmount}</Label>
                  </div>
                  <div className="col-sm-2" hidden={this.state.isOnBehalfandRetired && !this.state.isAdminOnBehalfUser} >
                    <Label className="control-Label font-weight-bold">HR1 Remarks</Label>
                  </div>
                  <div className="col-sm-6" hidden={this.state.isOnBehalfandRetired && !this.state.isAdminOnBehalfUser}>
                    <Label className="control-Label "> {this.state.CHSApproverView.HR1Remark === null ||
                      this.state.CHSApproverView.HR1Remark === undefined ||
                      this.state.CHSApproverView.HR1Remark === ''
                      ? 'NA'
                      : this.state.CHSApproverView.HR1Remark}</Label>
                  </div>
                </div>
                <div className="row form-group">
                  <div className="col-sm-2" >
                    <Label className="control-Label font-weight-bold">Final  Approved Amount</Label>
                  </div>
                  <div className="col-sm-2" >
                    <TextField type='number'
                      name="ExpenseDetails.HR2FinalAmount"
                      onChanged={(e: any) => this.handleInputChangeadd(event)}></TextField>
                  </div>
                  <div className="col-sm-8">
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
            </div>
            <div className="col-sm-12" >
              <Label className="control-Label font-weight-bold col-md-2">Remarks</Label>
              <TextField type='text' className='col-md-8'
                name="ExpenseDetails.HR2Remarks"
                onPaste={this.handleRemarksPaste}
                onKeyPress={(e) => { if (/[<>&"'\/]/.test(e.key)) { e.preventDefault(); alert('Special characters like <, >, &, ", \', / are not allowed.'); } }}
                onChanged={(e: any) => this.handleInputChangeadd(event)}></TextField>
            </div>
          </div>
          <div className='text-center'>
            <PrimaryButton className={styles.custombtn + " " + "mr-2"} onClick={() => this.BtnApproveHR2Request()} disabled={this.state.isApproving}>{this.state.isApproving ? <Spinner size={SpinnerSize.small} /> : "Approve"}</PrimaryButton>
            <PrimaryButton className={styles.custombtn + " " + "mr-2"} onClick={() => this.BtnRejectRequest('HR2')} disabled={this.state.isRejecting}>{this.state.isRejecting ? <Spinner size={SpinnerSize.small} /> : "Reject"}</PrimaryButton>
            <PrimaryButton onClick={() => this.closeDialog()} >Close</PrimaryButton>
          </div>
        </Dialog>
        <Dialog
          hidden={!this.state.isDialogViewHR2}
          onDismiss={this.closeDialog}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Comprehensive Health Checkup View Form',
            closeButtonAriaLabel: 'Close',
          }}
          containerClassName={'ms-dialogMainOverride ' + styles.textDialog}
        >
          <div className="card card-body">
            <div className="panel panel-default">
              <div className='panel-body'>
                <div className="row form-group">
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Finacial Year:</Label>
                  </div>
                  <div className="col-sm-6">
                    <Label className="control-Label">{this.state.CHSApproverView.FinancialYear}</Label>
                  </div>
                </div>
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
                    <Label className="control-Label font-weight-bold">Scale</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.Scale}</Label>
                  </div>
                </div>
                <div className="row form-group">
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
                <div className="row form-group" hidden={this.state.CHSApproverView.DependentType == 'Self' || isMDorDMDApprover} >
                  <div className="col-sm-4" >
                    <Label className="control-Label font-weight-bold">Is Spouse an Exim Employee<span>*</span></Label>
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
                <div className="row form-group">
                  <div className="col-sm-2" style={{ display: isMDorDMDApprover ? "none" : "block" }}>
                    <Label className="control-Label font-weight-bold">CHS Limit</Label>
                  </div>
                  <div className="col-sm-2" style={{ display: isMDorDMDApprover ? "none" : "flex" }}>
                    <Label className="control-Label ">{this.state.CHSApproverView.Limit}</Label>
                  </div>
                </div>
                {/* Table of Added Dependents */}
                {this.state.dependentitems.length > 0 && (
                  <div className="row form-group">
                    <div className="col-sm-12">
                      <table className="table table-bordered">
                        <thead>
                          <tr>
                            <th>Dependent Type</th>
                            <th>Amount Claimed</th>
                            <th style={{ display: isMDorDMDApprover ? "none" : "table-cell" }}>Calculated CHS Eligibility</th>
                            <th>Attachment</th>
                          </tr>
                        </thead>
                        <tbody>
                          {this.state.dependentitems.map((item, index) => (
                            <tr key={index}>
                              <td>{item.DependentType}</td>
                              <td>{item.AmountClaimed}</td>
                              <td style={{ display: isMDorDMDApprover ? "none" : "table-cell" }}>{item.CalculatedCHSEligibilityAmountLabel}</td>
                              <td>
                                {this.state.dependentDocsByType &&
                                  item &&
                                  item.DependentType &&
                                  this.state.dependentDocsByType[item.DependentType] &&
                                  this.state.dependentDocsByType[item.DependentType].length ? (
                                  this.state.dependentDocsByType[item.DependentType].map(function (f) {
                                    return (
                                      <a href={f.ServerRelativeUrl} target="_blank" rel="noreferrer">
                                        {f.Name}
                                      </a>
                                    );
                                  })
                                ) : (
                                  <span>-</span>
                                )}
                              </td>
                            </tr>
                          ))}
                        </tbody>
                        <tfoot>
                          <tr>
                            <td colSpan={1}><strong>Total</strong></td>
                            <td colSpan={1}>{this.state.CHSApproverView.AmountClaimed}</td>
                            <td colSpan={1} style={{ display: isMDorDMDApprover ? "none" : "table-cell" }}>{this.state.CHSApproverView.CHSEligibilityAmount}</td>
                          </tr>
                        </tfoot>
                      </table>
                    </div>
                  </div>
                )}
                <div className="row form-group">
                  <div className="col-sm-2" hidden={this.state.isOnBehalfandRetired && !this.state.isAdminOnBehalfUser}>
                    <Label className="control-Label font-weight-bold">HR1 Approved Amount</Label>
                  </div>
                  <div className="col-sm-2" hidden={this.state.isOnBehalfandRetired && !this.state.isAdminOnBehalfUser}>
                    <Label className="control-Label "> {this.state.CHSApproverView.FinalAmount === null ||
                      this.state.CHSApproverView.FinalAmount === undefined ||
                      this.state.CHSApproverView.FinalAmount === ''
                      ? 0
                      : this.state.CHSApproverView.FinalAmount}</Label>
                  </div>
                  <div className="col-sm-2" hidden={this.state.isOnBehalfandRetired && !this.state.isAdminOnBehalfUser}>
                    <Label className="control-Label font-weight-bold">HR1 Remarks</Label>
                  </div>
                  <div className="col-sm-6" hidden={this.state.isOnBehalfandRetired && !this.state.isAdminOnBehalfUser}>
                    <Label className="control-Label "> {this.state.CHSApproverView.HR1Remark === null ||
                      this.state.CHSApproverView.HR1Remark === undefined ||
                      this.state.CHSApproverView.HR1Remark === ''
                      ? 'NA'
                      : this.state.CHSApproverView.HR1Remark}</Label>
                  </div>
                </div>
                <div className="row form-group">
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Final  Approved Amount</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">
                      {this.state.CHSApproverView.HRApprovedAmount === null ||
                        this.state.CHSApproverView.HRApprovedAmount === undefined ||
                        this.state.CHSApproverView.HRApprovedAmount === ''
                        ? 0
                        : this.state.CHSApproverView.HRApprovedAmount}</Label>
                  </div>
                  <div className="col-sm-2" hidden={!this.state.isOnBehalfandRetired && !this.state.isAdminOnBehalfUser} >
                    <Label className="control-Label font-weight-bold">HR2 Remarks</Label>
                  </div>
                  <div className="col-sm-6" hidden={!this.state.isOnBehalfandRetired && !this.state.isAdminOnBehalfUser}>
                    <Label className="control-Label ">
                      {this.state.CHSApproverView.HR2Remark === null ||
                        this.state.CHSApproverView.HR2Remark === undefined ||
                        this.state.CHSApproverView.HR2Remark === ''
                        ? 'NA'
                        : this.state.CHSApproverView.HR2Remark}
                    </Label>
                  </div>
                </div>


                <div className="row form-group">
                  <div className="col-sm-2" hidden={!this.state.isOnBehalfandRetired && !this.state.isAdminOnBehalfUser} >
                    <Label className="control-Label font-weight-bold">TAG Remarks</Label>
                  </div>
                  <div className="col-sm-6" hidden={!this.state.isOnBehalfandRetired && !this.state.isAdminOnBehalfUser}>
                    <Label className="control-Label ">
                      {this.state.CHSApproverView.TagApproverRemark === null ||
                        this.state.CHSApproverView.TagApproverRemark === undefined ||
                        this.state.CHSApproverView.TagApproverRemark === ''
                        ? 'NA'
                        : this.state.CHSApproverView.TagApproverRemark}
                    </Label>
                  </div>
                </div>

              </div>
            </div>
            <div className="panel panel-default" hidden={this.state.CHSApproverView.EmployeeType != "RETIRED"}>
              <div className='panel-body'>
                <div className="row form-group">
                  <div className="col-sm-2" hidden={this.state.CHSApproverView.EmployeeType != "RETIRED"} >
                    <Label className="control-Label font-weight-bold">HR Remarks For Retired Employee</Label>
                  </div>
                  <div className="col-sm-8" hidden={this.state.CHSApproverView.EmployeeType != "RETIRED"}>
                    <Label className="control-Label ">
                      {this.state.CHSApproverView.HRRemarkForRetired === null ||
                        this.state.CHSApproverView.HRRemarkForRetired === undefined ||
                        this.state.CHSApproverView.HRRemarkForRetired === ''
                        ? 'NA'
                        : this.state.CHSApproverView.HRRemarkForRetired}
                    </Label>
                  </div>
                </div>
              </div></div>
          </div>
          <div className='text-center'>
            <PrimaryButton onClick={() => this.closeDialog()} >Close</PrimaryButton>
          </div>
        </Dialog>

        <Dialog
          hidden={!this.state.isDialogTagApprover}
          onDismiss={this.closeDialog}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Comprehensive Health Checkup Approver Form',
            closeButtonAriaLabel: 'Close',
          }}
          containerClassName={'ms-dialogMainOverride ' + styles.textDialog}
        >
          <div className="card card-body">
            <div className="panel panel-default">
              <div className='panel-body'>
                <div className="row form-group">
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Finacial Year:</Label>
                  </div>
                  <div className="col-sm-6">
                    <Label className="control-Label">{this.state.CHSApproverView.FinancialYear}</Label>
                  </div>
                </div>
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
                    <Label className="control-Label font-weight-bold">Scale</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.Scale}</Label>
                  </div>
                </div>
                <div className="row form-group">
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
                <div className="row form-group" hidden={this.state.CHSApproverView.DependentType == 'Self' || isMDorDMDApprover} >
                  <div className="col-sm-4" >
                    <Label className="control-Label font-weight-bold">Is Spouse an Exim Employee<span>*</span></Label>
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
                <div className="row form-group">
                  <div className="col-sm-2" style={{ display: isMDorDMDApprover ? "none" : "block" }}>
                    <Label className="control-Label font-weight-bold">CHS Limit</Label>
                  </div>
                  <div className="col-sm-2" style={{ display: isMDorDMDApprover ? "none" : "flex" }}>
                    <Label className="control-Label ">{this.state.CHSApproverView.Limit}</Label>
                  </div>
                </div>
                {/* Table of Added Dependents */}
                {this.state.dependentitems.length > 0 && (
                  <div className="row form-group">
                    <div className="col-sm-12">
                      <table className="table table-bordered">
                        <thead>
                          <tr>
                            <th>Dependent Type</th>
                            <th>Amount Claimed</th>
                            <th style={{ display: isMDorDMDApprover ? "none" : "table-cell" }}>Calculated CHS Eligibility</th>
                            <th>Attachment</th>
                          </tr>
                        </thead>
                        <tbody>
                          {this.state.dependentitems.map((item, index) => (
                            <tr key={index}>
                              <td>{item.DependentType}</td>
                              <td>{item.AmountClaimed}</td>
                              <td style={{ display: isMDorDMDApprover ? "none" : "table-cell" }}>{item.CalculatedCHSEligibilityAmountLabel}</td>
                              <td>
                                {this.state.dependentDocsByType &&
                                  item &&
                                  item.DependentType &&
                                  this.state.dependentDocsByType[item.DependentType] &&
                                  this.state.dependentDocsByType[item.DependentType].length ? (
                                  this.state.dependentDocsByType[item.DependentType].map(function (f) {
                                    return (
                                      <a href={f.ServerRelativeUrl} target="_blank" rel="noreferrer">
                                        {f.Name}
                                      </a>
                                    );
                                  })
                                ) : (
                                  <span>-</span>
                                )}
                              </td>
                            </tr>
                          ))}
                        </tbody>
                        <tfoot>
                          <tr>
                            <td colSpan={1}><strong>Total</strong></td>
                            <td colSpan={1}>{this.state.CHSApproverView.AmountClaimed}</td>
                            <td colSpan={1} style={{ display: isMDorDMDApprover ? "none" : "table-cell" }}>{this.state.CHSApproverView.CHSEligibilityAmount}</td>
                          </tr>
                        </tfoot>
                      </table>
                    </div>
                  </div>
                )}
                <div className="row form-group">
                  <div className="col-sm-2" hidden={this.state.isOnBehalfandRetired && !this.state.isAdminOnBehalfUser}>
                    <Label className="control-Label font-weight-bold">HR1 Approved Amount</Label>
                  </div>
                  <div className="col-sm-2" hidden={this.state.isOnBehalfandRetired && !this.state.isAdminOnBehalfUser}>
                    <Label className="control-Label ">
                      {this.state.CHSApproverView.FinalAmount === null ||
                        this.state.CHSApproverView.FinalAmount === undefined ||
                        this.state.CHSApproverView.FinalAmount === ''
                        ? 0
                        : this.state.CHSApproverView.FinalAmount}
                    </Label>
                  </div>
                  <div className="col-sm-2" hidden={this.state.isOnBehalfandRetired && !this.state.isAdminOnBehalfUser} >
                    <Label className="control-Label font-weight-bold">HR1 Remarks</Label>
                  </div>
                  <div className="col-sm-6" hidden={this.state.isOnBehalfandRetired && !this.state.isAdminOnBehalfUser}>
                    <Label className="control-Label ">
                      {this.state.CHSApproverView.HR1Remark === null ||
                        this.state.CHSApproverView.HR1Remark === undefined ||
                        this.state.CHSApproverView.HR1Remark === ''
                        ? 'NA'
                        : this.state.CHSApproverView.HR1Remark}
                    </Label>
                  </div>
                </div>
                <div className="row form-group">
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Final Approved Amount</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label">
                      {this.state.CHSApproverView.HRApprovedAmount === null ||
                        this.state.CHSApproverView.HRApprovedAmount === undefined ||
                        this.state.CHSApproverView.HRApprovedAmount === ''
                        ? 0
                        : this.state.CHSApproverView.HRApprovedAmount}
                    </Label>
                  </div>
                  <div className="col-sm-2" hidden={!this.state.isOnBehalfandRetired && !this.state.isAdminOnBehalfUser} >
                    <Label className="control-Label font-weight-bold">HR2 Remarks</Label>
                  </div>
                  <div className="col-sm-6" hidden={!this.state.isOnBehalfandRetired && !this.state.isAdminOnBehalfUser}>
                    <Label className="control-Label ">
                      {this.state.CHSApproverView.HR2Remark === null ||
                        this.state.CHSApproverView.HR2Remark === undefined ||
                        this.state.CHSApproverView.HR2Remark === ''
                        ? 'NA'
                        : this.state.CHSApproverView.HR2Remark}
                    </Label>
                  </div>
                </div>

                <div className="row form-group">
                  <div className="col-sm-2" hidden={!this.state.isOnBehalfandRetired && !this.state.isAdminOnBehalfUser} >
                    <Label className="control-Label font-weight-bold">TAG Remarks</Label>
                  </div>
                  <div className="col-sm-6" hidden={!this.state.isOnBehalfandRetired && !this.state.isAdminOnBehalfUser}>
                    <Label className="control-Label ">
                      {this.state.CHSApproverView.TagApproverRemark === null ||
                        this.state.CHSApproverView.TagApproverRemark === undefined ||
                        this.state.CHSApproverView.TagApproverRemark === ''
                        ? 'NA'
                        : this.state.CHSApproverView.TagApproverRemark}
                    </Label>
                  </div>
                </div>

              </div>
            </div>
            <div className="panel panel-default" hidden={this.state.CHSApproverView.EmployeeType != "RETIRED"}>
              <div className='panel-body'>
                <div className="row form-group">
                  <div className="col-sm-2" hidden={this.state.CHSApproverView.EmployeeType != "RETIRED"} >
                    <Label className="control-Label font-weight-bold">HR Remarks For Retired Employee</Label>
                  </div>
                  <div className="col-sm-8" hidden={this.state.CHSApproverView.EmployeeType != "RETIRED"}>
                    <Label className="control-Label ">
                      {this.state.CHSApproverView.HRRemarkForRetired === null ||
                        this.state.CHSApproverView.HRRemarkForRetired === undefined ||
                        this.state.CHSApproverView.HRRemarkForRetired === ''
                        ? 'NA'
                        : this.state.CHSApproverView.HRRemarkForRetired}
                    </Label>
                  </div>
                </div>
              </div>
            </div>
            <div className="col-sm-12" style={{ padding: 0 }}>
              <Label className="control-Label font-weight-bold col-md-2">Remarks</Label>
              <TextField type='text' className='col-md-8'
                name="ExpenseDetails.TagApproverRemarks"
                onChanged={(e: any) => this.handleInputChangeadd(event)}
                onPaste={this.handleRemarksPaste}
                onKeyPress={(e) => { if (/[<>&"'\/]/.test(e.key)) { e.preventDefault(); alert('Special characters like <, >, &, ", \', / are not allowed.'); } }}></TextField>
            </div>
          </div>
          <div className='text-center'>
            <PrimaryButton className={styles.custombtn + " " + "mr-2"} onClick={() => this.BtnApproveTagApproverRequest()} disabled={this.state.isApproving}>{this.state.isApproving ? <Spinner size={SpinnerSize.small} /> : "Approve"}</PrimaryButton>
            <PrimaryButton className={styles.custombtn + " " + "mr-2"} onClick={() => this.BtnRejectRequest('TagApprover')} disabled={this.state.isRejecting}>{this.state.isRejecting ? <Spinner size={SpinnerSize.small} /> : "Reject"}</PrimaryButton>
            <PrimaryButton onClick={() => this.closeDialog()} >Close</PrimaryButton>
          </div>
        </Dialog>
        <Dialog
          hidden={!this.state.isDialogViewTagApprover}
          onDismiss={this.closeDialog}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Comprehensive Health Checkup View Form',
            closeButtonAriaLabel: 'Close',
          }}
          containerClassName={'ms-dialogMainOverride ' + styles.textDialog}
        >
          <div className="card card-body">
            <div className="panel panel-default">
              <div className='panel-body'>
                <div className="row form-group">
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Finacial Year:</Label>
                  </div>
                  <div className="col-sm-6">
                    <Label className="control-Label">{this.state.CHSApproverView.FinancialYear}</Label>
                  </div>
                </div>
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
                    <Label className="control-Label font-weight-bold">Scale</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">{this.state.CHSApproverView.Scale}</Label>
                  </div>
                </div>
                <div className="row form-group">
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
                <div className="row form-group" hidden={this.state.CHSApproverView.DependentType == 'Self' || isMDorDMDApprover} >
                  <div className="col-sm-4" >
                    <Label className="control-Label font-weight-bold">Is Spouse an Exim Employee<span>*</span></Label>
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
                <div className="row form-group">
                  <div className="col-sm-2" style={{ display: isMDorDMDApprover ? "none" : "block" }}>
                    <Label className="control-Label font-weight-bold">CHS Limit</Label>
                  </div>
                  <div className="col-sm-2" style={{ display: isMDorDMDApprover ? "none" : "flex" }}>
                    <Label className="control-Label ">{this.state.CHSApproverView.Limit}</Label>
                  </div>
                </div>
                {/* Table of Added Dependents */}
                {this.state.dependentitems.length > 0 && (
                  <div className="row form-group">
                    <div className="col-sm-12">
                      <table className="table table-bordered">
                        <thead>
                          <tr>
                            <th>Dependent Type</th>
                            <th>Amount Claimed</th>
                            <th style={{ display: isMDorDMDApprover ? "none" : "table-cell" }}>Calculated CHS Eligibility</th>
                            <th>Attachment</th>
                          </tr>
                        </thead>
                        <tbody>
                          {this.state.dependentitems.map((item, index) => (
                            <tr key={index}>
                              <td>{item.DependentType}</td>
                              <td>{item.AmountClaimed}</td>
                              <td style={{ display: isMDorDMDApprover ? "none" : "table-cell" }}>{item.CalculatedCHSEligibilityAmountLabel}</td>
                              <td>
                                {this.state.dependentDocsByType &&
                                  item &&
                                  item.DependentType &&
                                  this.state.dependentDocsByType[item.DependentType] &&
                                  this.state.dependentDocsByType[item.DependentType].length ? (
                                  this.state.dependentDocsByType[item.DependentType].map(function (f) {
                                    return (
                                      <a href={f.ServerRelativeUrl} target="_blank" rel="noreferrer">
                                        {f.Name}
                                      </a>
                                    );
                                  })
                                ) : (
                                  <span>-</span>
                                )}
                              </td>
                            </tr>
                          ))}
                        </tbody>
                        <tfoot>
                          <tr>
                            <td colSpan={1}><strong>Total</strong></td>
                            <td colSpan={1}>{this.state.CHSApproverView.AmountClaimed}</td>
                            <td colSpan={1} style={{ display: isMDorDMDApprover ? "none" : "table-cell" }}>{this.state.CHSApproverView.CHSEligibilityAmount}</td>
                          </tr>
                        </tfoot>
                      </table>
                    </div>
                  </div>
                )}
                <div className="row form-group">
                  <div className="col-sm-2" hidden={this.state.isOnBehalfandRetired && !this.state.isAdminOnBehalfUser}>
                    <Label className="control-Label font-weight-bold">HR1 Approved Amount</Label>
                  </div>
                  <div className="col-sm-2" hidden={this.state.isOnBehalfandRetired && !this.state.isAdminOnBehalfUser}>
                    <Label className="control-Label "> {this.state.CHSApproverView.FinalAmount === null ||
                      this.state.CHSApproverView.FinalAmount === undefined ||
                      this.state.CHSApproverView.FinalAmount === ''
                      ? 0
                      : this.state.CHSApproverView.FinalAmount}</Label>
                  </div>
                  <div className="col-sm-2" hidden={this.state.isOnBehalfandRetired && !this.state.isAdminOnBehalfUser}>
                    <Label className="control-Label font-weight-bold">HR1 Remarks</Label>
                  </div>
                  <div className="col-sm-6" hidden={this.state.isOnBehalfandRetired && !this.state.isAdminOnBehalfUser}>
                    <Label className="control-Label "> {this.state.CHSApproverView.HR1Remark === null ||
                      this.state.CHSApproverView.HR1Remark === undefined ||
                      this.state.CHSApproverView.HR1Remark === ''
                      ? 'NA'
                      : this.state.CHSApproverView.HR1Remark}</Label>
                  </div>
                </div>
                <div className="row form-group">
                  <div className="col-sm-2">
                    <Label className="control-Label font-weight-bold">Final  Approved Amount</Label>
                  </div>
                  <div className="col-sm-2">
                    <Label className="control-Label ">
                      {this.state.CHSApproverView.HRApprovedAmount === null ||
                        this.state.CHSApproverView.HRApprovedAmount === undefined ||
                        this.state.CHSApproverView.HRApprovedAmount === ''
                        ? 0
                        : this.state.CHSApproverView.HRApprovedAmount}</Label>
                  </div>
                  <div className="col-sm-2" hidden={!this.state.isOnBehalfandRetired && !this.state.isAdminOnBehalfUser} >
                    <Label className="control-Label font-weight-bold">HR2 Remarks</Label>
                  </div>
                  <div className="col-sm-6" hidden={!this.state.isOnBehalfandRetired && !this.state.isAdminOnBehalfUser}>
                    <Label className="control-Label ">
                      {this.state.CHSApproverView.HR2Remark === null ||
                        this.state.CHSApproverView.HR2Remark === undefined ||
                        this.state.CHSApproverView.HR2Remark === ''
                        ? 'NA'
                        : this.state.CHSApproverView.HR2Remark}
                    </Label>
                  </div>
                </div>

                <div className="row form-group">
                  <div className="col-sm-2" hidden={!this.state.isOnBehalfandRetired && !this.state.isAdminOnBehalfUser} >
                    <Label className="control-Label font-weight-bold">TAG Remarks</Label>
                  </div>
                  <div className="col-sm-6" hidden={!this.state.isOnBehalfandRetired && !this.state.isAdminOnBehalfUser}>
                    <Label className="control-Label ">
                      {this.state.CHSApproverView.TagApproverRemark === null ||
                        this.state.CHSApproverView.TagApproverRemark === undefined ||
                        this.state.CHSApproverView.TagApproverRemark === ''
                        ? 'NA'
                        : this.state.CHSApproverView.TagApproverRemark}
                    </Label>
                  </div>
                </div>

              </div>
            </div>
            <div className="panel panel-default" hidden={this.state.CHSApproverView.EmployeeType != "RETIRED"}>
              <div className='panel-body'>
                <div className="row form-group">
                  <div className="col-sm-2" hidden={this.state.CHSApproverView.EmployeeType != "RETIRED"} >
                    <Label className="control-Label font-weight-bold">HR Remarks For Retired Employee</Label>
                  </div>
                  <div className="col-sm-8" hidden={this.state.CHSApproverView.EmployeeType != "RETIRED"}>
                    <Label className="control-Label ">
                      {this.state.CHSApproverView.HRRemarkForRetired === null ||
                        this.state.CHSApproverView.HRRemarkForRetired === undefined ||
                        this.state.CHSApproverView.HRRemarkForRetired === ''
                        ? 'NA'
                        : this.state.CHSApproverView.HRRemarkForRetired}
                    </Label>
                  </div>
                </div>
              </div></div>
          </div>
          <div className='text-center'>
            <PrimaryButton onClick={() => this.closeDialog()} >Close</PrimaryButton>
          </div>
        </Dialog>
      </div>
    );
  }
}
