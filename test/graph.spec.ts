/// <reference path="../typings/tsd.d.ts" />
import { expect } from "chai";
import { mock } from "sinon";
import { Graph } from "../src/KurveGraph";

describe("Graph", () => {

    describe("meAsync", () => {

        it("should get /me url", (done) => {
            var graph = new Graph({
                defaultAccessToken: "access_token"
            });

            var mockedGet = mock(graph);

            mockedGet.expects("get")
                .withArgs("https://graph.microsoft.com/v1.0/me/")
                .yields(JSON.stringify({ displayName: "John" }), null);

            graph.meAsync().then((user) => {
                expect(user.data.displayName).to.be.equal("John");
                mockedGet.verify();
                done();
            });
        });
    });
});
