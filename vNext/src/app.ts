// https://github.com/Microsoft/TypeScript/issues/4717
import "jquery-binarytransport"; // Load for side effects

import { AppInsights } from "./telemetry";

// appInsights.trackEvent("test", { data1: "s" })
import * as JSZip from "jszip";

// Import file-saver and account for a bug in the type definitions.
import * as fileSaver from "file-saver";

import * as WorkItemTrackingContracts from "TFS/WorkItemTracking/Contracts";
import * as WorkItemTrackingClient from "TFS/WorkItemTracking/RestClient";

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

export async function downloadAsync(replacement: string) {
    try {
        let pbitBytes = await ajaxBlobAsync("templates/Flat.pbit");
        let pbitZip = await new JSZip().loadAsync(pbitBytes);
        let mashupBuffer = await pbitZip.file("DataMashup").async("arraybuffer");

        let headerView = new Int32Array(mashupBuffer, 0, 2);
        let partsBytesCount = headerView[1];
        let partsBytes = new Uint8Array(mashupBuffer, 8, 8 + partsBytesCount);
        let otherBytesView = new Uint8Array(mashupBuffer, 1 + 8 + partsBytesCount);

        let partsZip = new JSZip();
        let parts = await partsZip.loadAsync(partsBytes);
        let section = await parts.file("Formulas/Section1.m").async("string");

        section = section.replace(/stansw/, replacement);
        parts.remove("Formulas/Section1.m");
        parts.file("Formulas/Section1.m", section);

        let partsBytesNew = await parts.generateAsync({ type: "uint8array" });

        let mashupBufferNew = new ArrayBuffer(8 + partsBytesNew.byteLength + otherBytesView.byteLength);
        let headerNewView = new Int32Array(mashupBufferNew, 0, 2);
        let partsBytesNewView = new Uint8Array(mashupBufferNew, 8, partsBytesNew.byteLength);
        let otherBytesNewView = new Uint8Array(mashupBufferNew, 1 + 8 + otherBytesView.byteLength);

        headerNewView[0] = 0;
        headerNewView[1] = partsBytesNew.byteLength;
        for (let i = 0; i < partsBytesNew.byteLength; i++) {
            partsBytesNewView[i] = partsBytesNew[i];
        }
        for (let i = 0; i < otherBytesNewView.byteLength; i++) {
            otherBytesNewView[i] = otherBytesNewView[i];
        }

        pbitZip.remove("DataMashup");
        pbitZip.file("DataMashup", mashupBufferNew);

        let pbitBytesNew = await pbitZip.generateAsync({ type: "blob" });

        fileSaver.saveAs(pbitBytesNew, "hello.pbit");
    } catch (exception) {
        AppInsights.trackException(exception, "sdf", { template: "Flat" });
        console.log("Operation failed");
    }
}

interface IConfiguration {
    close?: () => void;
    qid: string;
}

let configuration = VSS.getConfiguration() as IConfiguration;
let context = VSS.getWebContext();


let counter = 10;
let counterId = setInterval(() => {
    counter -= 1;
    if (counter < 0) {
        configuration.close();
        clearInterval(counterId);
    }
    else {
        $("#countdown").text(counter);
    }
}, 1000);
$("#countdown-cancel").click(() => {
    clearInterval(counterId);
    $("#countdown-message").hide();
});

async function tralala() {
    let workItemTrackingClient = WorkItemTrackingClient.getClient();
    let query = await workItemTrackingClient.getQuery(context.project.name, configuration.qid);

    let url = "templateUrl"
        + WorkItemTrackingContracts.QueryType[query.queryType]
        + "?" + $.param({
            hostUrl: context.host.uri,
            projectName: context.project.name,
            teamName: context.team.name,
            queryId: configuration.qid,
            queryName: query.name
        });

    downloadAsync(url);
}

tralala();
