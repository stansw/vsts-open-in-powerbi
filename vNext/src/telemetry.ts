// Use this code to statically load ai.0.js library 
// https://github.com/Microsoft/ApplicationInsights-JS
// /// <reference path="../node_modules/applicationinsights-js/bundle/ai.module.d.ts" />
// let init = new (<any>Microsoft).ApplicationInsights.Initialization({
//     config: {
//         instrumentationKey: "478f349e-1267-43eb-a2aa-b70654cb6409"
//     }
// });
// let appInsights = <Microsoft.ApplicationInsights.IAppInsights>init.loadAppInsights();

import * as ai from "applicationinsights-js";

ai.AppInsights.downloadAndSetup({ instrumentationKey: "478f349e-1267-43eb-a2aa-b70654cb6409" });

let webContext = VSS.getWebContext();
if (webContext) {
    appInsights.setAuthenticatedUserContext(webContext.user.id, webContext.collection.id);
}

export var AppInsights = ai.AppInsights;