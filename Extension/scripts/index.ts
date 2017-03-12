import * as App from "./app";

let extensionContext = VSS.getExtensionContext();
VSS.register(`${extensionContext.publisherId}.${extensionContext.extensionId}.openQueryAction`, App.openQueryAction);
VSS.register(`${extensionContext.publisherId}.${extensionContext.extensionId}.openQueryOnToolbarAction`, App.openQueryOnToolbarAction);
