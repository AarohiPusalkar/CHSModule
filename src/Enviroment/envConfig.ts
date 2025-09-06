const ENVIRONMENT = "UAT"; // Defaults to UAT if not set

const CONFIG = {

  UAT: {
    siteUrl: "https://sharepointwebssse.eximbankindia.in/sites/hrm",
    rootSiteUrl:"https://sharepointwebssse.eximbankindia.in/",
    listIds: {
      EmployeeList: "1234-uat-5678",
      ProjectList: "2345-uat-6789",
    },
  },
  PROD: {
    siteUrl: "https://sharepointweb.eximbankindia.in/sites/hrm",
    rootSiteUrl:"https://sharepointweb.eximbankindia.in/",

    listIds: {
      EmployeeList: "1234-prod-5678",
      ProjectList: "2345-prod-6789",
    },
  },
};

export const ENV_CONFIG = CONFIG[ENVIRONMENT];
