/// <reference path="../node_modules/applicationinsights-js/bundle/ai.module.d.ts" />

let init = new (<any>Microsoft).ApplicationInsights.Initialization({
    config: {
        instrumentationKey: "478f349e-1267-43eb-a2aa-b70654cb6409"
    }
});
let appInsights = <Microsoft.ApplicationInsights.IAppInsights>init.loadAppInsights();

let webContext = VSS.getWebContext();
if (webContext !== undefined) {
    appInsights.setAuthenticatedUserContext(webContext.user.id, webContext.collection.id);
}

//appInsights.trackEvent("test", { data1: "s" })
import * as JSZip from "jszip";

// Import file-saver and account for a bug in the type definitions.
import * as fileSaver from "file-saver";
let saveAs: typeof fileSaver.saveAs = <any>fileSaver;

export class Greeter {
    element: HTMLElement;
    span: HTMLElement;
    timerToken: number;

    public static value = 11;

    constructor(element: HTMLElement) {
        this.element = element;
        this.element.innerHTML += "The time is: ";
        this.span = document.createElement("span");
        this.element.appendChild(this.span);
        this.span.innerText = new Date().toUTCString();
    }

    start() {
        var a = new JSZip();
        a.file("Hello.txt", "tralala");
        this.timerToken = setInterval(() => this.span.innerHTML = new Date().toUTCString(), 500);
    }

    stop() {
        clearTimeout(this.timerToken);
    }

}

const el = document.getElementById("content");


new Promise<Blob>((resolve, reject) => {
    $.ajax({
        url: "static/templates/Flat.pbitx",
        type: "GET",
        dataType: "binary"
    })
        .done((value) => resolve(value))
        .fail((jqXHR, textStatus, errorThrown) =>
            reject(errorThrown instanceof Error
                ? errorThrown
                : new Error(errorThrown.toString()))
        )
})
    .then(value => {
        var zip = new JSZip();
        return zip.loadAsync(value)
    })
    .then(pbit => {
        return pbit.file("DataMashup")
            .async("arraybuffer")
            .then((mashupBuffer: ArrayBuffer) => {
                let headerView = new Int32Array(mashupBuffer, 0, 2);
                let partsBytesCount = headerView[1];
                let partsBytes = new Uint8Array(mashupBuffer, 8, 8 + partsBytesCount);
                let otherBytesView = new Uint8Array(mashupBuffer, 1 + 8 + partsBytesCount);

                let partsZip = new JSZip();
                return partsZip.loadAsync(partsBytes)
                    .then(parts => {
                        return parts.file("Formulas/Section1.m")
                            .async("string")
                            .then((section: string) => {
                                section = section.replace(/stansw/, "dziala");
                                parts.remove("Formulas/Section1.m");
                                parts.file("Formulas/Section1.m", section);
                                return parts;
                            })
                    })
                    .then(parts => {
                        return <Promise<Uint8Array>>parts.generateAsync({ type: "uint8array" })
                    })
                    .then(partsBytesNew => {
                        console.log(partsBytes.byteLength);
                        console.log(partsBytesNew.byteLength);

                        let mashupBufferNew = new ArrayBuffer(8 + partsBytesNew.byteLength + otherBytesView.byteLength);
                        let headerNewView = new Int32Array(mashupBufferNew, 0, 2);
                        let partsBytesNewView = new Uint8Array(mashupBufferNew, 8, partsBytesNew.byteLength);
                        let otherBytesNewView = new Uint8Array(mashupBufferNew, 1 + 8 + otherBytesView.byteLength)

                        headerNewView[0] = 0;
                        headerNewView[1] = partsBytesNew.byteLength;
                        for (let i = 0; i < partsBytesNew.byteLength; i++) {
                            partsBytesNewView[i] = partsBytesNew[i];
                        }
                        for (let i = 0; i < otherBytesNewView.byteLength; i++) {
                            otherBytesNewView[i] = otherBytesNewView[i];
                        }

                        return mashupBufferNew;
                    });
            })
            .then(mashupBufferNew => {
                pbit.remove("DataMashup")
                pbit.file("DataMashup", mashupBufferNew);
                return pbit;
            });
    })
    .then(pbit => {
        return <Promise<Blob>>pbit.generateAsync({ type: "blob" })
    })
    .then(blob => {
        saveAs(blob, "hello.pbit");
    })
    .catch(reason => {
        appInsights.trackException(reason, "sdf", { template: "Flat" });
        console.log("Operation failed")
    });

//const greeter = new Greeter(el);
//greeter.start();
