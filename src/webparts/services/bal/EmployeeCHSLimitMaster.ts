import { IChsModuleProps } from '../../chsModule/components/IChsModuleProps';
import SPCRUDOPS from '../dal/spcrudops';
import {IEmployeeCHSLimitMaster  } from "../interface/IEmployeeCHSLimitMaster";
export interface IEmployeeCHSLimitMasterOps {
    getAllEmployeeCHSLimit(props: IChsModuleProps): Promise<IEmployeeCHSLimitMaster>;
}
export default function EmployeeCHSLimitMasterOps() {
    const spCrudOps = SPCRUDOPS();
    const getAllEmployeeCHSLimit = async (props: IChsModuleProps): Promise<IEmployeeCHSLimitMaster | null> => {
        try {
            const results = await (await spCrudOps).getDataAnotherSiteCollection(
                "EmployeeCHSLimitMaster",
                //  "*",
                "*,Scale/Title,Scale/Id,Designation/Title,Designation/Id",
                "Designation,Scale",
                // "",
                "",
                { column: "Id", isAscending: false },
                props
            );
            if (results && results.length > 0) {
                const firstResult = results;
                return firstResult;
            } else {
                console.warn("No employee found for the current user.");
                return null;
            }
        } catch (error) {
            console.error("Error in getEmployeeMaster:", error);
            return null;
        }
    };
    return {
        getAllEmployeeCHSLimit
    };
}