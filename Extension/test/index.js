define(["require", "exports", "./app"], function (require, exports, App) {
    "use strict";
    var extensionContext = VSS.getExtensionContext();
    VSS.register(extensionContext.publisherId + "." + extensionContext.extensionId + ".openQueryAction", App.openQueryAction);
    VSS.register(extensionContext.publisherId + "." + extensionContext.extensionId + ".openQueryOnToolbarAction", App.openQueryOnToolbarAction);
});
