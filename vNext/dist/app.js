define(["require", "exports", "telemetryclient-team-services-extension", "./telemetryClientSettings"], function (require, exports, tc, telemetryClientSettings) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    var Greeter = (function () {
        function Greeter(element) {
            this.element = element;
            this.element.innerHTML += "The time is: ";
            this.span = document.createElement("span");
            this.element.appendChild(this.span);
            this.span.innerText = new Date().toUTCString();
        }
        Greeter.prototype.start = function () {
            var _this = this;
            this.timerToken = setInterval(function () { return _this.span.innerHTML = new Date().toUTCString(); }, 500);
        };
        Greeter.prototype.stop = function () {
            clearTimeout(this.timerToken);
        };
        return Greeter;
    }());
    var el = document.getElementById("content");
    var greeter = new Greeter(el);
    greeter.start();
    tc.TelemetryClient.getClient(telemetryClientSettings.settings).trackPageView("vsts-open-in-powerbi-tralala.Index");
});
//# sourceMappingURL=app.js.map