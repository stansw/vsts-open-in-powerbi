 /// <reference types="vss-web-extension-sdk" />
import * as tc from "telemetryclient-team-services-extension";
import telemetryClientSettings = require("./telemetryClientSettings");
export class Greeter {
    element: HTMLElement;
    span: HTMLElement;
    timerToken: number;

    public static value = 10;

    constructor(element: HTMLElement) {
        this.element = element;
        this.element.innerHTML += "The time is: ";
        this.span = document.createElement("span");
        this.element.appendChild(this.span);
        this.span.innerText = new Date().toUTCString();
    }

    start() {
        this.timerToken = setInterval(() => this.span.innerHTML = new Date().toUTCString(), 500);
    }

    stop() {
        clearTimeout(this.timerToken);
    }

}

const el = document.getElementById("content");
//const greeter = new Greeter(el);
//greeter.start();

tc.TelemetryClient.getClient(telemetryClientSettings.settings).trackPageView("vsts-open-in-powerbi-tralala.Index");
