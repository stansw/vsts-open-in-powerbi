/// <reference path="../node_modules/applicationinsights-js/bundle/ai.module.d.ts" />

var init = new (<any>Microsoft).ApplicationInsights.Initialization({
    config: {
        instrumentationKey: "478f349e-1267-43eb-a2aa-b70654cb6409"
    }
});
var appInsights = <Microsoft.ApplicationInsights.IAppInsights>init.loadAppInsights();
appInsights.trackEvent("test", { data1 : "s"} )


import * as zip from "jszip";



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
//const greeter = new Greeter(el);
//greeter.start();


var snippet = {
    config: {
        instrumentationKey: "478f349e-1267-43eb-a2aa-b70654cb6409"
    }
};


// AppInsights.

// var init = new ApplicationInsights.Initialization(snippet);   
// var appInsights = init.loadAppInsights();   
// appInsights.trackPageView("vsts-open-in-powerbi-tralala.Index");  



