import * as tc from "telemetryclient-team-services-extension";

const telemetryClientSettings: tc.TelemetryClientSettings = {
    key: "478f349e-1267-43eb-a2aa-b70654cb6409",
    extensioncontext: "vsts-open-in-powerbi-tralala",
    disableTelemetry: "false",
    disableAjaxTracking: "false",
    enableDebug: "false"
};

export class Greeter {
    element: HTMLElement;
    span: HTMLElement;
    timerToken: number;

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
const greeter = new Greeter(el);
greeter.start();

tc.TelemetryClient.getClient(telemetryClientSettings).trackPageView("vsts-open-in-powerbi-tralala.Index");
