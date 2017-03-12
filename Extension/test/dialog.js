define(["require", "exports", "TFS/WorkItemTracking/Contracts", "TFS/WorkItemTracking/RestClient"], function (require, exports, Contracts, RestClient) {
    "use strict";
    var templateUrl = "http://vsts-open-in-powerbi.azurewebsites.net/_api/Template/";
    var configuration = VSS.getConfiguration();
    var counter = 10;
    var id = setInterval(function () {
        counter -= 1;
        if (counter < 0) {
            configuration.close();
            clearInterval(id);
        }
        else {
            $("#countdown").text(counter);
        }
    }, 1000);
    function cancelAutoClose() {
        clearInterval(id);
        $("#countdown-message").hide();
    }
    exports.cancelAutoClose = cancelAutoClose;
    function openUrl(url) {
        cancelAutoClose();
        VSS.getService(VSS.ServiceIds.Navigation).then(function (navigationService) {
            navigationService.openNewWindow(url);
        });
    }
    exports.openUrl = openUrl;
    RestClient.getClient().getQuery(configuration.projectName, configuration.qid, Contracts.QueryExpand.Clauses).then(function (query) {
        var url = templateUrl
            + Contracts.QueryType[query.queryType]
            + "?" + $.param({
            url: configuration.hostUrl,
            project: configuration.projectName,
            qid: configuration.qid,
            qname: query.name
        });
        VSS.getService(VSS.ServiceIds.Navigation).then(function (navigationService) {
            navigationService.navigate(url);
        });
    });
});
