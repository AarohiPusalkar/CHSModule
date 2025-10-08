import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface IChsModuleProps {
  description: string;
  currentSPContext: WebPartContext;
  selectedOuterTab?: string; // Add this
  isHrGroup1?:boolean;
  isHrGroup2?:boolean;
  userLoginName?:string;


 

}
