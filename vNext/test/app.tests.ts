import { expect } from "chai";

describe('A suite', function () {
    it('contains spec with an expectation', function () {
        let a = 10;
        expect(a).to.equal(9, "true doesn't equal tralala!")
    });
});