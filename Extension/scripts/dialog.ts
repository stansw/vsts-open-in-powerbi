import Contracts = require("TFS/WorkItemTracking/Contracts");
import RestClient = require("TFS/WorkItemTracking/RestClient");
import * as App from "./app";

const templateUrl = "http://vsts-open-in-powerbi.azurewebsites.net/_api/Template/";

let configuration = VSS.getConfiguration() as App.IDialogConfiguration;

let counter = 10;
let id = setInterval(() => {
    counter -= 1;
    if (counter < 0) {
        configuration.close();
        clearInterval(id);
    }
    else {
        $("#countdown").text(counter);
    }
}, 1000);

export function cancelAutoClose() {
    clearInterval(id);
    $("#countdown-message").hide();
}

export function openUrl(url: string) {
    // If you clicked a link, we will cancel auto close.
    cancelAutoClose();
    VSS.getService(VSS.ServiceIds.Navigation).then((navigationService: any) => {
        navigationService.openNewWindow(url);
    });
}

RestClient.getClient().getQuery(configuration.projectName, configuration.qid, Contracts.QueryExpand.Clauses).then(
    (query) => {
        let url = templateUrl
            + Contracts.QueryType[query.queryType]
            + "?" + $.param({
                url: configuration.hostUrl,
                project: configuration.projectName,
                team: configuration.teamName,
                qid: configuration.qid,
                qname: query.name
            });

        VSS.getService(VSS.ServiceIds.Navigation).then((navigationService: any) => {
            navigationService.navigate(url);
        });
    });