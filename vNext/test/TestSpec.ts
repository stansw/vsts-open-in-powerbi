import { expect } from "chai";

import * as app from "../src/app"

describe('A suite', function () {
    it('contains spec with an expectation', function () {
        expect(app.Greeter.value).to.equal(11);
    });
});