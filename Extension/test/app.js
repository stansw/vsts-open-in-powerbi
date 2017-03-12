define(["require", "exports"], function (require, exports) {
    "use strict";
    var SupportedActions;
    (function (SupportedActions) {
        SupportedActions.OpenItems = "OpenItems";
        SupportedActions.OpenQuery = "OpenQuery";
    })(SupportedActions || (SupportedActions = {}));
    var WellKnownQueries;
    (function (WellKnownQueries) {
        WellKnownQueries.AssignedToMe = "A2108D31-086C-4FB0-AFDA-097E4CC46DF4";
        WellKnownQueries.UnsavedWorkItems = "B7A26A56-EA87-4C97-A504-3F028808BB9F";
        WellKnownQueries.FollowedWorkItems = "202230E0-821E-401D-96D1-24A7202330D0";
        WellKnownQueries.CreatedBy = "53FB153F-C52C-42F1-90B6-CA17FC3561A8";
        WellKnownQueries.SearchResults = "2CBF5136-1AE5-4948-B59A-36F526D9AC73";
        WellKnownQueries.CustomWiql = "08E20883-D56C-4461-88EB-CE77C0C7936D";
        WellKnownQueries.RecycleBin = "2650C586-0DE4-4156-BA0E-14BCFB664CCA";
    })(WellKnownQueries || (WellKnownQueries = {}));
    exports.queryExclusionList = [
        WellKnownQueries.AssignedToMe,
        WellKnownQueries.UnsavedWorkItems,
        WellKnownQueries.FollowedWorkItems,
        WellKnownQueries.CreatedBy,
        WellKnownQueries.SearchResults,
        WellKnownQueries.CustomWiql,
        WellKnownQueries.RecycleBin
    ];
    function isSupportedQueryId(queryId) {
        return queryId && exports.queryExclusionList.indexOf(queryId.toUpperCase()) === -1;
    }
    exports.isSupportedQueryId = isSupportedQueryId;
    function generateUrl(action, collection, project, qid, wids, columns) {
        var url = "tfs://ExcelRequirements/" + action + "?cn=" + collection + "&proj=" + project;
        if (action === SupportedActions.OpenQuery) {
            if (!qid) {
                throw new Error("'qid' must be provided for '" + SupportedActions.OpenQuery + "' action.");
            }
            url += "&qid=" + qid;
        }
        else {
            throw new Error("Unsupported action provided: " + action);
        }
        if (url.length > 2000) {
            throw new Error('Generated url is exceeds the maxlength, please reduce the number of work items you selected.');
        }
        return url;
    }
    exports.generateUrl = generateUrl;
    exports.openQueryAction = {
        getMenuItems: function (context) {
            if (!context || !context.query || !context.query.wiql || !isSupportedQueryId(context.query.id)) {
                return null;
            }
            else {
                return [{
                        title: "Open in Power BI",
                        text: "Open in Power BI",
                        icon: "img/powerbi_logo_16x16.png",
                        action: function (actionContext) {
                            if (actionContext && actionContext.query && actionContext.query.id) {
                                var context_1 = VSS.getWebContext();
                                showNotification({
                                    close: null,
                                    hostUrl: context_1.host.uri,
                                    projectName: context_1.project.name,
                                    qid: actionContext.query.id
                                });
                            }
                        }
                    }];
            }
        }
    };
    exports.openQueryOnToolbarAction = {
        getMenuItems: function (context) {
            return [{
                    title: "Open in Power BI",
                    text: "Open in Power BI",
                    icon: "img/powerbi_logo_16x16.png",
                    showText: true,
                    action: function (actionContext) {
                        if (actionContext && actionContext.query && actionContext.query.wiql && isSupportedQueryId(actionContext.query.id)) {
                            var context_2 = VSS.getWebContext();
                            showNotification({
                                close: null,
                                hostUrl: context_2.host.uri,
                                projectName: context_2.project.name,
                                qid: actionContext.query.id
                            });
                        }
                        else {
                            VSS.getService(VSS.ServiceIds.Dialog).then(function (hostDialogService) {
                                hostDialogService.openMessageDialog("In order to open query please save it in \"My Queries\" or \"Shared Queries\".", {
                                    title: "Unable to perform operation"
                                });
                            });
                        }
                    }
                }];
        }
    };
    function showNotification(configuration) {
        var extensionContext = VSS.getExtensionContext();
        var dialog;
        configuration.close = function () { return dialog.close(); };
        VSS.getService(VSS.ServiceIds.Dialog).then(function (hostDialogService) {
            hostDialogService.openDialog(extensionContext.publisherId + "." + extensionContext.extensionId + ".notificationDialog", {
                title: "Downloading Power BI file...",
                width: 400,
                height: 670,
                modal: true,
                draggable: false,
                resizable: false,
                buttons: {
                    "ok": {
                        id: "ok",
                        text: "Dismiss",
                        click: function () {
                            dialog.close();
                        },
                        class: "cta",
                    }
                }
            }, configuration)
                .then(function (d) {
                dialog = d;
            });
        });
    }
});
