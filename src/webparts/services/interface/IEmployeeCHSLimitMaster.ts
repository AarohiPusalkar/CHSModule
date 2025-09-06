// export interface IEmployeeCHSLimitMaster{
//     length: number;
//     find(arg0: (limit: { Scale: { Title: string; }; }) => boolean): unknown;
//     // map(arg0: (item: any) => { key: any; text: any; }): unknown;
// //     map(arg0: (item: any) => { key: any; text: any }): unknown
//     Id?:number,
//     Title:string,
//     Limit:any,
//     EmployeeType?:any,
//     Sacle?:any,
//     Designation:any,
// }
export interface IEmployeeCHSLimitMaster {
    Id?: number;
    Title: string;
    Limit: any;
    EmployeeType?: any;
    Scale?: { Title: string }; // Corrected spelling from "Sacle" to "Scale"
    Designation: any;
}
