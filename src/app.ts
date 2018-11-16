import "../static/css/app.css";

// https://github.com/Microsoft/TypeScript/issues/4717
import "jquery-binarytransport"; // Load for side effects


import { AppInsights } from "./telemetry";
import { IConfiguration } from "./common";

// appInsights.trackEvent("test", { data1: "s" })
import * as JSZip from "jszip";

// Import file-saver and account for a bug in the type definitions.
import * as fileSaver from "file-saver";

import * as WorkItemTrackingContracts from "TFS/WorkItemTracking/Contracts";
import * as WorkItemTrackingClient from "TFS/WorkItemTracking/RestClient";
import { HostNavigationService } from "VSS/SDK/Services/Navigation";


const TemplateUrl = "http://vsts-open-in-powerbi.azurewebsites.net/_api/Template/";

let configuration = VSS.getConfiguration() as IConfiguration;

// Control the counter in the dialog to hide it automatically.
let counter = 10;
let counterId = setInterval(() => {
    counter -= 1;
    if (counter < 0) {
        clearInterval(counterId);
        AppInsights.flush();
        configuration.close();
    }
    else {
        $("#countdown").text(counter);
    }
}, 1000);
$("#countdown-cancel").click(() => {
    clearInterval(counterId);
    $("#countdown-message").hide();
});


function ajaxBlobAsync(url: string): Promise<Blob> {
    return new Promise<Blob>((resolve, reject) => {
        $.ajax({
            url: url,
            type: "GET",
            dataType: "binary"
        })
            .done((value) => resolve(value))
            .fail((jqXHR, textStatus, errorThrown) =>
                reject(errorThrown instanceof Error
                    ? errorThrown
                    : new Error(`Failed to get resource: ${url} (${errorThrown.toString()})`)
                )
            );
    });
}

async function transformDataMashupAsync(data: Blob, transform: (section: string) => string): Promise<Blob> {
    let pbitZip = await new JSZip().loadAsync(data);
    let mashupBuffer = await pbitZip.file("DataMashup").async("arraybuffer");

    let headerView = new Int32Array(mashupBuffer, 0, 2);
    let partsBytesCount = headerView[1];
    let partsBytes = new Uint8Array(mashupBuffer, 8, partsBytesCount);
    let otherBytesView = new Uint8Array(mashupBuffer, 8 + partsBytesCount);

    let partsZip = new JSZip();
    let parts = await partsZip.loadAsync(partsBytes);
    let section = await parts.file("Formulas/Section1.m").async("string");

    // Transform section.
    section = transform(section);

    // Replace section in the Parts archive.
    parts.remove("Formulas/Section1.m");
    parts.file("Formulas/Section1.m", section);

    // Generate new Part bytes.
    let partsBytesNew = await parts.generateAsync({ type: "uint8array" });

    // Add new Part bytes to the Mashup archive.
    let mashupBufferNew = new ArrayBuffer(8 + partsBytesNew.byteLength + otherBytesView.byteLength);
    let headerNewView = new Int32Array(mashupBufferNew, 0, 2);
    let partsBytesNewView = new Uint8Array(mashupBufferNew, 8, partsBytesNew.byteLength);
    let otherBytesNewView = new Uint8Array(mashupBufferNew, 8 + partsBytesNew.byteLength, otherBytesView.length);

    // Copy bytes to account for any change in the length.
    headerNewView[0] = 0;
    headerNewView[1] = partsBytesNew.byteLength;
    partsBytesNewView.set(partsBytesNew);
    otherBytesNewView.set(otherBytesView);

    // Replace Mashup bytes.
    pbitZip.remove("DataMashup");
    pbitZip.file("DataMashup", mashupBufferNew);

    return pbitZip.generateAsync({ type: "blob" });
}

export async function mainAsync() {
    try {
        let context = VSS.getWebContext();
        let hosted = context.account.name !== "TEAM FOUNDATION"
            && context.account.name !== "Team Foundation Server";
        let workItemTrackingClient = WorkItemTrackingClient.getClient();
        let query = await workItemTrackingClient.getQuery(
            context.project.name,
            configuration.queryId,
            WorkItemTrackingContracts.QueryExpand.All);
        let queryType = WorkItemTrackingContracts.QueryType[query.queryType];
        let queryMode = WorkItemTrackingContracts.LinkQueryMode[query.filterOptions] || "WorkItems";
        let success = false;

        // Diagnostics data.
        let scenario = "DownloadQueryFromExtension";
        let traceData = {
            collectionId: context.collection.id,
            projectId: context.project.id,
            teamId: context.team.id,
            contribution: configuration.contribution,
            queryId: query.id,
            queryMode: queryMode,
            queryType: queryType
        };

        try {
            AppInsights.startTrackEvent(scenario);

            let transform = (section) => {
                let url = context.host.uri
                    .replace(/dev\.azure\.com\/([^\/]*)/, (match, account) => `${account}.visualstudio.com`)
                    .replace(/\/$/, "");
                let replacements = {
                    // Do not include trailing forward slash in the URL.
                    '    url = "[^"]*",': `    url = "${url}",`,
                    // Define replacements for all parameters.
                    '    collection = "[^"]*",': `    collection = "${hosted ? "" : context.collection.name}",`,
                    '    project = "[^"]*",': `    project = "${context.project.name}",`,
                    '    team = "[^"]*",': `    team = "${context.team.name}",`,
                    '    id = "[^"]*",': `    id = "${configuration.queryId}",`
                };
                for (let pattern in replacements) {
                    section = section.replace(new RegExp(pattern), replacements[pattern]);
                }
                return section;
            };

            let deployment = hosted ? "VSTS" : "TFS";
            let templateBytes = await ajaxBlobAsync(`templates/${queryType}.${queryMode}.${deployment}.pbit`);
            let updatedBytes = await transformDataMashupAsync(templateBytes, transform);
            fileSaver.saveAs(updatedBytes, `${query.name}.pbit`);
            success = true;
        } catch (exception) {
            AppInsights.trackException(exception, null, traceData);
        } finally {
            AppInsights.stopTrackEvent(scenario, traceData);
        }
    } catch (exception) {
        AppInsights.trackException(exception);
    } finally {
        AppInsights.flush();
    }
}