import * as React from 'react';
import { PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import styles from '../ChsModule.module.scss'
import * as moment from 'moment'
import { IChsModuleProps } from '../IChsModuleProps';
import UseUtilities, { IUtilities } from '../../../services/bal/utilities';
import Utilities from '../../../services/bal/utilities';
import { Formik, FormikProps, ErrorMessage, Field } from 'formik';
import * as yup from 'yup';
import { sp } from '@pnp/sp';
import { Web } from '@pnp/sp/presets/all';
import {
    Stack, IStackTokens, ITag, TagPicker, IBasePickerSuggestionsProps, IBasePicker, IInputProps, Checkbox, TextField
    , DetailsList, IColumn, FontIcon, SelectionMode, DetailsListLayoutMode, IDetailsHeaderProps, DetailsHeader, ConstrainMode
    , PrimaryButton, ICheckboxProps, MessageBar, MessageBarType, DefaultButton
} from '@fluentui/react';
import { BaseButton, Button, FontWeights, IPersonaProps } from 'office-ui-fabric-react';

import { Link, useHistory } from 'react-router-dom';
import { Items } from '@pnp/sp/items';

import { Dropdown, IDropdown, DropdownMenuItemType, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import useSPCRUD, { ISPCRUD } from '../../../services/bal/spcrud';
import SPCRUD from '../../../services/bal/spcrud';

import EmployeeOps from '../../../services/bal/EmployeeMaster';

import { IEmployeeMaster } from '../../../services/interface/IEmployeeMaster';
// import { IEmployeeMaster2 } from '../../../services/interface/IEmployeeMaster2';

import { keys } from '@microsoft/sp-lodash-subset';
import { IEmployeeCHSLimitMaster } from '../../../services/interface/IEmployeeCHSLimitMaster';
import EmployeeCHSLimitMasterOps from '../../../services/bal/EmployeeCHSLimitMaster';

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
const options: IDropdownOption[] = [
    { key: 'Yes', text: 'Yes' },  // Hardcoded Yes
    { key: 'No', text: 'No' }    // Hardcoded No
];
export default class CHSCreation extends React.Component<IChsModuleProps, any> {
    constructor(props: any) {
        super(props);

        this.state = {
            GroupId: "", 
            AllEmployeeCollObj:[],
            NoteTypeId: "", 
            FinancialYearColl: [],
             FinancialYearId: "",
              Subject: "", 
              isMultiGrp: false, 
              showhideEmployeeNameLab: false,
              isOnBehalfDisabled:false
        };

    }
    async componentDidMount() {

        await this.getCurrentUser();
        // await this.currentUserGroup();


        
                 await this.GetEmployeelimit();


        await this.checkUserInGroups(["HR1_Group", "HR2_Group"]).then(isMember => {
            if (isMember) {
                this.setState({  showhideEmployeeNameLab : true })

            } else {
                this.setState({  showhideEmployeeNameLab : false, OnBehalf: 'No' , isOnBehalfDisabled : false})
            }
          });
        await this.getEmployee();
        await this.GetEmployeelimit();
        // await this.getAllEmployeeCHSLimit();

      

        
    }


    public getCurrentUser = async () => {
        const spCrudObj = await useSPCRUD();

        return await spCrudObj.currentUser(this.props).then(cuser => {
            this.setState({ Currentuser: cuser });
            return cuser;
        });
    }

    // public checkUserInGroups = async () => {
//  public async checkUserInGroups(groups: any[]): Promise<boolean> {


//         // const CheckUSerForGroup=['']
//         const spCrudObj = await useSPCRUD();

//         return await spCrudObj.currentUserGroup(this.props).then(cuser => {

//           for (const groupName of cuser) {
//             try {
//             //   const groupUsers = await web.siteGroups.getByName(groupName).users();
//             for (const groupsN of groups) {

//               const userExists = cuser.some(user => user.Title === groupsN);
    
//               if (userExists) {
//                 console.log(`User ${cuser.Title} exists in group "${groupName}"`);
//                 return true; // Return true if user is found in any group
//               }
//             }
//             } catch (error) {
//               console.warn(`Group "${groupName}" not found or inaccessible`);
//             }
//           }

      
//             // this.setState({ Currentuser: cuser });
//             // return cuser;
//         });
//     }


public async checkUserInGroups(groups: any): Promise<boolean> {
    try {
        const spCrudObj = await useSPCRUD();
        
        // Fetch the current user's groups
        const userGroups = await spCrudObj.currentUserGroup(this.props);

        // Ensure userGroups is not empty
        if (!userGroups || userGroups.length === 0) {
            console.log("User is not part of any group.");
            return false;
        }

        // Check if user exists in any of the specified groups
        const isUserInGroup = userGroups.some(group => groups.includes(group.Title));

        if (isUserInGroup) {
            console.log(`User exists in at least one of the specified groups.`);
            return true;
        } else {
            console.log(`User does not exist in any of the specified groups.`);
            return false;
        }

    } catch (error) {
        console.error("Error checking user in groups:", error);
        return false;
    }
}



    // public async checkUserInGroups(groupNames: string[]): Promise<boolean> {
    //     try {
    //         const spCrudObj = await useSPCRUD();

    //       // Get current user details
    //       const currentUser = await sp.web.currentUser.get();
    //       const userId = currentUser.Id;
          
    //       const web = Web(sp.web.toUrl());

    //      await this.getCurrentUser()
    
    //       // Loop through each group and check if user exists
    //       for (const groupName of groupNames) {
    //         try {
    //           const groupUsers = await web.siteGroups.getByName(groupName).users();
    //           const userExists = groupUsers.some(user => user.Id === userId);
    
    //           if (userExists) {
    //             console.log(`User ${currentUser.Title} exists in group "${groupName}"`);
    //             return true; // Return true if user is found in any group
    //           }
    //         } catch (error) {
    //           console.warn(`Group "${groupName}" not found or inaccessible`);
    //         }
    //       }
    
    //       console.log(`User ${currentUser.Title} does not exist in any specified groups`);
    //       return false; // Return false if user is not in any group
    
    //     } catch (error) {
    //       console.error("Error checking user in groups:", error);
    //       return false;
    //     }
    //   }

    //   public async checkUserInGroups(groupNames: string[]): Promise<boolean> {
    //     try {
    //         const spCrudObj = await useSPCRUD();
    
    //         // Get current logged-in user details using your existing function
    //         const currentUser = await this.getCurrentUser();
    //         if (!currentUser || !currentUser.LoginName) {
    //             console.error("Error fetching current user details.");
    //             return false;
    //         }
    
    //         const loginName = currentUser.LoginName.toLowerCase(); // Normalize for comparison
    //         const web = Web(sp.web.toUrl());
    
    //         // Loop through each group and check if user exists
    //         for (const groupName of groupNames) {
    //             try {
    //                 let web = Web(this.props.currentSPContext.pageContext.web.absoluteUrl);
    //                 const groupUsers = await web.siteGroups.getByName(groupName).users();
                    
    //                 const userExists = groupUsers.some(user => 
    //                     user.LoginName.toLowerCase() === loginName
    //                 );
    
    //                 if (userExists) {
    //                     console.log(`User ${currentUser.Title} exists in group "${groupName}"`);
    //                     return true; // User found in at least one group
    //                 }
    //             } catch (error) {
    //                 console.warn(`Group "${groupName}" not found or inaccessible.`);
    //             }
    //         }
    
    //         console.log(`User ${currentUser.Title} does not exist in any specified groups.`);
    //         return false; // User is not in any group
    
    //     } catch (error) {
    //         console.error("Error checking user in groups:", error);
    //         return false;
    //     }
    // }
    

    public getEmployee = async (): Promise<IEmployeeMaster> => {
        return await EmployeeOps().getEmployeeMaster(this.props).then(async (results) => {
            let employeeData = results;
            
            debugger;
            // Fetch Employee Limit List
            let limitData: IEmployeeCHSLimitMaster[] = await this.GetEmployeelimit();
    
            // Find matching limit based on Designation
            // let matchedLimit = limitData.find(limit => limit.Designation === employeeData.DesignationTitle);
            let matchedLimit = limitData.filter((e) => e.Scale.Title === employeeData.Scale &&  e.Designation.Title === employeeData.DesignationTitle &&e.EmployeeType === employeeData.EmpType);

            // Set state with the retrieved data
            // this.setState({
            //     EmployeeInfodb: employeeData,
            //     EmployeeName: employeeData.EmployeeName,
            //     EmployeeIDId: employeeData.Id,
            //     Approver: employeeData.LeaveLevel2?.Title || "N/A", 
            //     ApproverId: employeeData.LeaveLevel2Id,
            //     Level2Id: employeeData.LeaveLevel2Id,
            //     AccountNo: employeeData.AccountNo,
            //     IFSCCode: employeeData.IFSCCode,
            //     EmployeeID: employeeData.EmployeeId,
            //     DesignationId: employeeData.DesignationId,
            //     DesignationTitle: employeeData.DesignationTitle,
            //     DateofBirth: employeeData.DateofBirth,
            //     Scale: employeeData.Scale,
            //     Age: parseInt(employeeData.Age),
            //     EmpType: employeeData.EmpType,
            //     Limit: matchedLimit ? matchedLimit: "N/A", 
            //     GradeId: employeeData.GradeId,
            //     CurrentOfficeLocationId: employeeData.CurrentOfficeLocationId,
            //     EmployeeSubGroupId: employeeData.SubGroupId,
            //     Role: employeeData.Role
            // });

            this.setState({
                EmployeeInfodb: employeeData,
                EmployeeName: employeeData.EmployeeName,
                EmployeeIDId: employeeData.Id,
                Approver: employeeData.LeaveLevel2.Title, 
                ApproverId: employeeData.LeaveLevel2Id,
                Level2Id: employeeData.LeaveLevel2Id,
                AccountNo: employeeData.AccountNo,
                IFSCCode: employeeData.IFSCCode,
                EmployeeID: employeeData.EmployeeId,
                DesignationId: employeeData.DesignationId,
                DesignationTitle: employeeData.DesignationTitle,
                DateofBirth: employeeData.DateofBirth,
                Scale: employeeData.Scale,
                Age: parseInt(employeeData.Age),
                EmpType: employeeData.EmpType,
                Limit: matchedLimit.length > 0 && matchedLimit !== undefined ? matchedLimit[0].Limit :"",
                GradeId: employeeData.GradeId,
                CurrentOfficeLocationId: employeeData.CurrentOfficeLocationId,
                EmployeeSubGroupId: employeeData.SubGroupId,
                Role: employeeData.Role
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
                    // Get Employee Limit List (Use the same logic as in `getEmployee()`)
                    let matchedLimit = this.state.EmployeeCHSInfodb.filter(
                        (lim) =>
                            lim.Scale.Title === selectedEmp.Scale.Title &&
                            lim.Designation.Title === selectedEmp.Designation.Title &&
                            lim.EmployeeType === selectedEmp.EmpType
                    );
    
                    this.setState({
                        EmployeeName: selectedEmp.EmployeeName,
                        EmployeeIDId: selectedEmp.Id,
                        Approver: selectedEmp.LeaveLevel2.Title ,
                        ApproverId: selectedEmp.LeaveLevel2Id,
                        Level2Id: selectedEmp.LeaveLevel2Id,
                        AccountNo: selectedEmp.AccountNo,
                        IFSCCode: selectedEmp.IFSCCode,
                        EmployeeID: selectedEmp.EmployeeId,
                        DesignationId: selectedEmp.DesignationId,
                        DesignationTitle: selectedEmp.Designation.Title,
                        DateofBirth: selectedEmp.DOB
                            ? `${new Date(selectedEmp.DOB).getDate()}-${new Date(selectedEmp.DOB).getMonth() + 1}-${new Date(selectedEmp.DOB).getFullYear()}`
                            : null,
                        Scale: selectedEmp.Scale.Title,
                        Payscale: selectedEmp.Payscale.Title,
                        Age: parseInt(selectedEmp.Age),
                        EmpType: selectedEmp.EmpType,
                        GradeId: selectedEmp.GradeId,
                        CurrentOfficeLocationId: selectedEmp.CurrentOfficeLocationId,
                        EmployeeSubGroupId: selectedEmp.SubGroupId,
                        Role: selectedEmp.Role,
    
                        Limit: matchedLimit.length > 0 ? matchedLimit[0].Limit : "",
                    });
                }
            }
        }
    };
    


    // public getEmployee = async (): Promise<IEmployeeMaster> => {

    // //   await  this.GetEmployeelimit();

    //     return await EmployeeOps().getEmployeeMaster(this.props).then(results => {
            
    //         this.setState({
    //             EmployeeInfodb: results,
    //             EmployeeName: results.EmployeeName,
    //             EmployeeIDId: results.Id,
    //             Approver: results.LeaveLevel2.Title,
    //             ApproverId: results.LeaveLevel2Id,
    //             Level2Id: results.LeaveLevel2Id,
    //             AccountNo: results.AccountNo,
    //             IFSCCode: results.IFSCCode,
    //             EmployeeID: results.EmployeeId,
    //             DesignationId: results.DesignationId,
    //             DesignationTitle: results.DesignationTitle,
    //             DateofBirth: results.DateofBirth,
    //             Scale: results.Scale,
    //             Age:parseInt(results.Age),

    //             EmpType: results.EmpType,
    //             // Limit:results.Active,

    //             GradeId: results.GradeId,
    //             CurrentOfficeLocationId: results.CurrentOfficeLocationId,
    //             EmployeeSubGroupId: results.SubGroupId,
    //             Role: results.Role

    //         });
    //         return results;
    //     });
    // };


    //     public GetEmployeelimit = async (): Promise<IEmployeeCHSLimitMaster> => {
    //     return await EmployeeCHSLimitMasterOps().getAllEmployeeCHSLimit(this.props).then(CHSLimitresults => {

    //         this.setState({
    //             EmployeeCHSInfodb: CHSLimitresults,
    //             // Id: CHSLimitresults.Id,
    //             // Title: CHSLimitresults.Title,
    //             // Limit: CHSLimitresults.Limit,
    //             // EmployeeType: CHSLimitresults.EmployeeType,
    //             // Sacle: CHSLimitresults.Sacle.Title,
    //             // Designation: CHSLimitresults.Designation.Designation
    //         });
    //         return CHSLimitresults;
    //     });
    // };


    // public GetEmployeelimit = async (): Promise<IEmployeeCHSLimitMaster[]> => {
    //     return await EmployeeCHSLimitMasterOps().getAllEmployeeCHSLimit(this.props).then(CHSLimitresults => {
    //         let formattedResults: IEmployeeCHSLimitMaster[] = Array.isArray(CHSLimitresults) ? CHSLimitresults : [CHSLimitresults];
    
    //         this.setState({
    //             EmployeeCHSInfodb: formattedResults
    //         });
    
    //         return formattedResults;
    //     });
    // };





    
    // public getEmployeeWithLimit = async (): Promise<void> => {
    //     try {
    //         // Fetch Employee Data
    //         const employeeData: IEmployeeMaster = await EmployeeOps().getEmployeeMaster(this.props);
    
    //         if (!employeeData || !employeeData.Scale) {
    //             console.error("Employee data or Scale is missing.");
    //             return;
    //         }
    
    //         // Fetch Employee Limit Data
    //         let limitData: IEmployeeCHSLimitMaster[] = await EmployeeCHSLimitMasterOps().getAllEmployeeCHSLimit(this.props);
    
    //         if (!Array.isArray(limitData)) {
    //             console.warn("getAllEmployeeCHSLimit did not return an array, wrapping it into an array.");
    //             limitData = [limitData]; // Convert single object to an array
    //         }
    
    //         // Filter Limit Data Based on Employee Scale
    //         const filteredLimit: IEmployeeCHSLimitMaster | undefined = limitData.find((limit) => 
    //             limit.Scale?.Title === employeeData.Scale
    //         );
    
    //         // Update State with Employee Data and Limit
    //         this.setState({
    //             EmployeeInfodb: employeeData,
    //             EmployeeName: employeeData.EmployeeName,
    //             EmployeeIDId: employeeData.Id,
    //             Approver: employeeData.LeaveLevel2?.Title,
    //             ApproverId: employeeData.LeaveLevel2Id,
    //             Level2Id: employeeData.LeaveLevel2Id,
    //             AccountNo: employeeData.AccountNo,
    //             IFSCCode: employeeData.IFSCCode,
    //             EmployeeID: employeeData.EmployeeId,
    //             DesignationId: employeeData.DesignationId,
    //             DesignationTitle: employeeData.DesignationTitle,
    //             DateofBirth: employeeData.DateofBirth,
    //             Scale: employeeData.Scale,
    //             Age: parseInt(employeeData.Age),
    //             EmpType: employeeData.EmpType,
    //             GradeId: employeeData.GradeId,
    //             CurrentOfficeLocationId: employeeData.CurrentOfficeLocationId,
    //             EmployeeSubGroupId: employeeData.SubGroupId,
    //             Role: employeeData.Role,
    //             Limit: filteredLimit ? filteredLimit.Limit : null // Assign limit if found, else null
    //         });
    
    //     } catch (error) {
    //         console.error("Error fetching employee or limit data:", error);
    //     }
    // };
    


    
//     public getSelectedEmployeeDetail= (e)=> {
//          if(e.key!=undefined){

       
//         const { EmployeeInfodb } = this.state;
//         if(EmployeeInfodb.length>0){

//           let selectedEmp= EmployeeInfodb.filter(item=>{ return e.key==item.Id})[0];
// if(selectedEmp!=undefined){
//     this.setState({
//         // EmployeeInfodb: results,
//         EmployeeName: selectedEmp.EmployeeName,
//         EmployeeIDId: selectedEmp.Id,
//         Approver: selectedEmp.LeaveLevel2.Title,
//         ApproverId: selectedEmp.LeaveLevel2Id,
//         Level2Id: selectedEmp.LeaveLevel2Id,
//         AccountNo: selectedEmp.AccountNo,
//         IFSCCode: selectedEmp.IFSCCode,
//         EmployeeID: selectedEmp.EmployeeId,
//         DesignationId: selectedEmp.DesignationId,
//         DesignationTitle: selectedEmp.Designation.Title,
//         // DateofBirth: selectedEmp.DOB,
//         DateofBirth: selectedEmp.DOB
//         ? `${new Date(selectedEmp.DOB).getDate()}-${new Date(selectedEmp.DOB).getMonth() + 1}-${new Date(selectedEmp.DOB).getFullYear()}`
//         : null,
//         Scale: selectedEmp.Scale.Title,
//         Payscale: selectedEmp.Payscale.Title,    
//         Age:parseInt(selectedEmp.Age),
//         EmpType: selectedEmp.EmpType,

//         GradeId: selectedEmp.GradeId,
//         CurrentOfficeLocationId: selectedEmp.CurrentOfficeLocationId,
//         EmployeeSubGroupId: selectedEmp.SubGroupId,
//         Role: selectedEmp.Role

//     });


// }
            
//         }
//     }
//     };
 

    public getAllEmployee = async (): Promise<IEmployeeMaster> => {
        return await EmployeeOps().getAllEmployeeMaster(this.props).then(results => {
            debugger
            const employeeOptions = results.map((item: any) => ({
                key: item.Id.toString(), // Ensure key is string
                text: item.EmployeeName,
              }));
            // const employeeOptions = [{ key: results.Id, text: results.EmployeeName }]; // Create array with single employee

    
            this.setState({
                 EmployeeInfodb: results,
                AllEmployeeCollObj:employeeOptions
               
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


    // private onDropdownChange(event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void {
    //     this.setState({ OnBehalf: option?.key });
    //   }

    public render(): React.ReactElement<any> {
        return (
            <div className='container'>

                
                <Formik
                    initialValues={initialValues}
                    onSubmit={async values => {
                        await new Promise(resolve => setTimeout(resolve, 500));
                        alert(JSON.stringify(values, null, 2));

                    }}
                    validationSchema={validate}
                >
                    
                    {formik => (

                        
                        <div className='modalpane' style={{ margin: "0.5%", position: "absolute", width: "100%" }}>
                            <div className='panel panel-default'>
                                <div className='panel-body'>
                                    <div className='tab col-md-12 ng-scope'>
                                        <div className='form-horizontal'>
                                        <hr></hr>
                                            <h4 className={styles.label}>Employee Details </h4>
                                            <hr></hr>
                                            <br></br>
                                            <br></br>
                                            <div className="row">
                                                <div className="col-md-4">
                                                    <div><b className="header-style">On behalf of:</b></div>

                                                    
                                                    <Dropdown
                                                        placeHolder="Select Option"
                                                        options={options}
                                                        selectedKey={this.state.OnBehalf} // Binding selected value
                                                        disabled={this.state.isOnBehalfDisabled}
                                                        onChanged={(e, option) => {
                                                                if(e.key === "Yes")
                                                                {  
                                                                    this.getAllEmployee();
                                                                    this.setState({ OnBehalf: e.key, showhideEmployeeNameLab : true })
                                                                }
                                                                else{

                                                                    this.setState({ OnBehalf: e.key, showhideEmployeeNameLab : false })
                                                                }
                                                            }
                                                            
                                                        }
                                                        className="dropdown-style" // Add CSS if required
                                                    />
                                                </div>

                                                {/* <div className="col-md-4" hidden={this.state.showhideEmployeeNameLab}>
                                                    {
                                                        this.state.showhideEmployeeNameLab ?<>
                                                         <div><b className="header-style">Employee Name:</b></div>
                                                         <div>{this.state.EmployeeName}</div>
                                                        </>
                                                       
                                                        : ""
                                                    }
                                                    
                                                </div> */}

                                                <div className="col-md-4" hidden={this.state.showhideEmployeeNameLab}>
                                                    
                                                         <div><b className="header-style">Employee Name:</b></div>
                                                         <div>{this.state.EmployeeName}</div>
                                                        
                                                    
                                                    
                                                </div>
                                                <div className="col-md-4"  hidden={!this.state.showhideEmployeeNameLab}>
                                                    
                                                    <div><b className="header-style">Employee Name:</b></div>
                                                    <div>
                                                    <Dropdown
                                                            placeHolder="Select Employee"
                                                            options={this.state.AllEmployeeCollObj} // Dynamic options binding
                                                            selectedKey={this.state.EmployeeID} // Binding selected value
                                                            onChanged={(e, option) => {this.getSelectedEmployeeDetail(e);this.setState({ EmployeeID: e.key })}}
                                                            className="dropdown-style"
                                                        />


                                                    </div>
                                                   
                                               
                                               
                                           </div>



                                            </div>


                                            <br></br>
                                            <br></br>

                                            <div className="row">
                                                <div className="col-md-4">
                                                    <div><b className="header-style">Employee ID:</b></div>
                                                    <div>{this.state.EmployeeID}</div>
                                                </div>

                                                <div className="col-md-4">
                                                    <div><b className="header-style">Employee Name:</b></div>
                                                    <div>{this.state.EmployeeName}</div>
                                                </div>

                                                <div className="col-md-4">
                                                    <div><b className="header-style">Date of Birth :</b></div>
                                                    <div>{this.state.DateofBirth}</div>
                                                </div>
                                            </div>

                                            <div className="row">
                                                <div className="col-md-4">
                                                    <div><b className="header-style">Designation :</b></div>
                                                    <div>{this.state.DesignationTitle}</div>
                                                </div>

                                                <div className="col-md-4">
                                                    <div><b className="header-style">Scale :</b></div>
                                                    <div>{this.state.Scale}</div>
                                                </div>

                                                <div className="col-md-4">
                                                    <div><b className="header-style">Employee Type :</b></div>
                                                    <div>{this.state.EmpType}</div>
                                                </div>
                                            </div>

                                            <br></br><br></br>
                                            <br></br>
                                            <br></br>

                                            <hr></hr>
                                            <h4 className={styles.label}> Medical Health check up Details </h4>
                                            <hr></hr>

                                            <br></br>
                                            <br></br>

                                            <div>
                                                <div className="row">
                                                    <div className="col-md-4">
                                                        <div><b className="header-style">Dependent Type :</b></div>
                                                        <Dropdown
                                                            placeHolder="Select Dependent"
                                                            options={[{ key: 'Self', text: 'Self' }, { key: 'Spouse', text: 'Spouse' }]}
                                                            selectedKey={this.state.DependentType} // Binding selected value
                                                            onChanged={(e, option) => this.setState({ DependentType: e.key })}
                                                            className="dropdown-style"
                                                        />
                                                    </div>

                                                    <div className="col-md-4">
                                                        <div><b className="header-style">Age :</b></div>
                                                        <div>{this.state.Age}</div>
                                                    </div>

                                                    <div className="col-md-4">
                                                        <div><b className="header-style">CHS Limit :</b></div>
                                                        <div>{this.state.Limit}</div>
                                                    </div>
                                                </div>

<br></br><br></br>
<br></br>

                                                <div className="row">
                                                    <div className="col-md-4">
                                                    <div><b className="header-style">Amount Claimed :</b></div>
                                                    <input  type="number" id="quantity" name="quantity" min="1" />

                                                    </div>

                                                    <div className="col-md-4">
                                                        <div><b className="header-style">Final Amount :</b></div>
                                                        <div>{this.state.Limit}</div>
                                                    </div>


                                                </div>
                                            </div>

                                            <h4 className={styles.label}>Attachment Section</h4>
                                            <div className="col-sm-12">
                                                <div className="form-group">
                                                    <label className="control-label col-sm-4" style={{ textAlign: 'left' }}>Upload
                                                        File:</label>
                                                    <div className="col-sm-6" style={{ marginLeft: '-25%' }}>
                                                        <div> <span id="attach">
                                                            <input id="uploadFileID" type="file" multiple className="ng-isolate-scope" />

                                                        </span>
                                                        </div>
                                                    </div>
                                                </div>
                                            </div>

                                        </div>
                                    </div>

                                </div>


                            </div>
                        </div>
                    )}
                </Formik>

            </div>

        )
    }
}



