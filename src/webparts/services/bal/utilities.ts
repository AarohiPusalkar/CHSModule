import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IChsModuleProps } from '../../chsModule/components/IChsModuleProps';
import { IEmailProperties } from "@pnp/sp/sputilities";
import { sp } from "@pnp/sp";
export interface IMonth {
    Id: number;
    Title: string;
    ShortMonth: string;
    NarrowMonth: string;
}
export interface TypedHash<T> {
    [key: string]: T;
}
export interface IUtilities {
    filterData(jsonData: any, filterValue: string, excludeColumns: Array<string>): Promise<any>;
    MonthColl(): Array<IMonth>;
    hideShow(hideIt: boolean, props: WebPartContext, loadingMessage: string): void;
    sendEmail(to: string[], cc: string[], bcc: string[], subject: string, body: string, AdditionalHeaders: TypedHash<string>
        , props: IChsModuleProps, from?: string): void;
}
export default function Utilities() {
    const filterData = async (jsonData: any, filterValue: string, includeColumns: Array<string>) => {
        const lowercasedValue = filterValue.toLowerCase().trim();
        if (lowercasedValue === "") return jsonData;
        else {
            const filteredData = jsonData.filter(item => {
                // return Object.keys(item).some(key =>
                //     includeColumns.includes(key) ? item[key] !== undefined && item[key] !== null ? item[key].toString().toLowerCase().includes(lowercasedValue) : false : false
                // );
            });
            return filteredData;
        }
    };
    const MonthColl = () => {
        let Years: Array<IMonth> = new Array<IMonth>();
        let dateObj = new Date();
        for (let m = 0; m <= 11; m++) {
            let particularDate = new Date(dateObj.getFullYear(), m, 1);
            Years.push({
                Id: m + 1,
                Title: particularDate.toLocaleString('en-us', { month: 'long' }),
                ShortMonth: particularDate.toLocaleString('en-us', { month: 'short' }),
                NarrowMonth: particularDate.toLocaleString('en-us', { month: 'narrow' })
            });
        }
        return Years;
    };
    const hideShow = (hideIt: boolean, ctx: WebPartContext, loadingMessage: string) => {
        if (hideIt) {
            ctx.statusRenderer.clearLoadingIndicator(document.getElementById('divBlockRequestsLoader'));
            document.getElementById('divBlockRequestsLoader').setAttribute("style", "height: 0px !important;");
        }
        else {
            document.getElementById('divBlockRequestsLoader').setAttribute("style", "height: 1100px !important;");
            ctx.statusRenderer.displayLoadingIndicator(document.getElementById('divBlockRequestsLoader'), loadingMessage, 1);
        }
    };
    const sendEmail = async (to: string[], cc: string[], bcc: string[], subject: string, body: string
        , additionalHeaders: TypedHash<string>, props: IChsModuleProps, from?: string) => {
        const emailProps: IEmailProperties = {
            To: to,
            CC: cc,
            BCC: bcc,
            Subject: subject,
            Body: body
        };
        if (from !== null && from !== undefined && from.trim() !== "") {
            emailProps.From = from;
        }
        if (additionalHeaders !== null && additionalHeaders !== undefined) {
            emailProps.AdditionalHeaders = additionalHeaders;
        }
        return await sp.utility.sendEmail(emailProps);
    };
    const calculateAge = (dateOfBirth: Date )=> {
        const today = new Date();
        let age = today.getFullYear() - dateOfBirth.getFullYear();
        const m = today.getMonth() - dateOfBirth.getMonth();
        if (m < 0 || (m === 0 && today.getDate() < dateOfBirth.getDate())) {
            age--;
        }
        return age;
    };

    return {
        filterData,
        MonthColl,
        hideShow,
        sendEmail,
        calculateAge
    };
}
export async function isUserInGroup(userLoginName: string, groupName: string): Promise<boolean> {
    try {
      const groupUsers = await sp.web.siteGroups.getByName(groupName).users();
      return groupUsers.some(user => user.LoginName.toLowerCase() === userLoginName.toLowerCase());
    } catch (error) {
      console.error(`Error checking group membership for group "${groupName}":`, error);
      return false;
    }
  }

