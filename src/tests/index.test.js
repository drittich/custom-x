"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const index_1 = require("../src/index");
const originalOffice = global.Office;
beforeEach(() => {
    global.Office = {
        AsyncResultStatus: {
            Succeeded: "succeeded",
        },
        context: {
            document: {
                customXmlParts: {
                    // Default mock implementation which can be overridden in tests
                    getByNamespaceAsync: jest.fn(),
                    addAsync: jest.fn(),
                },
            },
        },
    };
});
afterEach(() => {
    global.Office = originalOffice;
});
describe("custom-x library", () => {
    test("setCustomXmlPartByValue adds a new custom XML part", async () => {
        // Simulate that no existing custom XML part is found.
        global.Office.context.document.customXmlParts.getByNamespaceAsync.mockImplementation((namespace, callback) => {
            callback({ status: global.Office.AsyncResultStatus.Succeeded, value: [] });
        });
        // Simulate successful addition of a new XML part.
        global.Office.context.document.customXmlParts.addAsync.mockImplementation((xml, callback) => {
            callback({ status: global.Office.AsyncResultStatus.Succeeded });
        });
        await expect((0, index_1.setCustomXmlPartByValue)("test key", { dummy: true })).resolves.toBeUndefined();
    });
    test("getCustomXmlPartValue returns parsed value", async () => {
        // Create a fake custom XML part with a getXmlAsync method.
        const fakeXmlPart = {
            getXmlAsync: (callback) => {
                // The XML content includes the wrapper (customData) which your code strips away.
                callback({
                    status: global.Office.AsyncResultStatus.Succeeded,
                    value: '<customData><dummy>true</dummy></customData>',
                });
            },
        };
        global.Office.context.document.customXmlParts.getByNamespaceAsync.mockImplementation((namespace, callback) => {
            callback({ status: global.Office.AsyncResultStatus.Succeeded, value: [fakeXmlPart] });
        });
        const result = await (0, index_1.getCustomXmlPartValue)("test key");
        // Depending on the XML parser, the dummy value may be a string.
        expect(result).toEqual({ dummy: true });
    });
    test("removeCustomXmlPart deletes an existing custom XML part", async () => {
        const fakeCustomXmlPart = {
            deleteAsync: (callback) => {
                callback({ status: global.Office.AsyncResultStatus.Succeeded });
            },
        };
        global.Office.context.document.customXmlParts.getByNamespaceAsync.mockImplementation((namespace, callback) => {
            callback({ status: global.Office.AsyncResultStatus.Succeeded, value: [fakeCustomXmlPart] });
        });
        await expect((0, index_1.removeCustomXmlPart)("test key")).resolves.toBeUndefined();
    });
    // add a test that makes sure an error is throw when there is more than one custom xml part with the same namespace
    test("removeCustomXmlPart throws an error when multiple custom XML parts are found", async () => {
        const fakeCustomXmlPart1 = { id: "part1" };
        const fakeCustomXmlPart2 = { id: "part2" };
        global.Office.context.document.customXmlParts.getByNamespaceAsync.mockImplementation((namespace, callback) => {
            callback({
                status: global.Office.AsyncResultStatus.Succeeded,
                value: [fakeCustomXmlPart1, fakeCustomXmlPart2],
            });
        });
        await expect((0, index_1.removeCustomXmlPart)("test key")).rejects.toThrow("More than one custom XML part found for the namespace:");
    });
});
