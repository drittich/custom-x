import {
    setCustomXmlPartByValue,
    getCustomXmlPartValue,
    removeCustomXmlPart,
} from '../index';

// Create minimal mocks for Office objects and statuses.
declare var global: any;
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
        (global.Office.context.document.customXmlParts.getByNamespaceAsync as jest.Mock).mockImplementation(
            (namespace: string, callback: (result: any) => void) => {
                callback({ status: global.Office.AsyncResultStatus.Succeeded, value: [] });
            }
        );

        // Simulate successful addition of a new XML part.
        (global.Office.context.document.customXmlParts.addAsync as jest.Mock).mockImplementation(
            (xml: string, callback: (result: any) => void) => {
                callback({ status: global.Office.AsyncResultStatus.Succeeded });
            }
        );

        await expect(setCustomXmlPartByValue("test key", { dummy: true })).resolves.toBeUndefined();
    });

    test("getCustomXmlPartValue returns parsed value", async () => {
        // Create a fake custom XML part with a getXmlAsync method.
        const fakeXmlPart = {
            getXmlAsync: (callback: Function) => {
                // The XML content includes the wrapper (customData) which your code strips away.
                callback({
                    status: global.Office.AsyncResultStatus.Succeeded,
                    value: '<customData><dummy>true</dummy></customData>',
                });
            },
        };

        (global.Office.context.document.customXmlParts.getByNamespaceAsync as jest.Mock).mockImplementation(
            (namespace: string, callback: (result: any) => void) => {
                callback({ status: global.Office.AsyncResultStatus.Succeeded, value: [fakeXmlPart] });
            }
        );

        const result = await getCustomXmlPartValue("test key");
        // Depending on the XML parser, the dummy value may be a string.
        expect(result).toEqual({ dummy: true });
    });

    test("removeCustomXmlPart deletes an existing custom XML part", async () => {
        const fakeCustomXmlPart = {
            deleteAsync: (callback: Function) => {
                callback({ status: global.Office.AsyncResultStatus.Succeeded });
            },
        };

        (global.Office.context.document.customXmlParts.getByNamespaceAsync as jest.Mock).mockImplementation(
            (namespace: string, callback: (result: any) => void) => {
                callback({ status: global.Office.AsyncResultStatus.Succeeded, value: [fakeCustomXmlPart] });
            }
        );

        await expect(removeCustomXmlPart("test key")).resolves.toBeUndefined();
    });

	// add a test that makes sure an error is throw when there is more than one custom xml part with the same namespace
	test("removeCustomXmlPart throws an error when multiple custom XML parts are found", async () => {
		const fakeCustomXmlPart1 = { id: "part1" };
		const fakeCustomXmlPart2 = { id: "part2" };

		(global.Office.context.document.customXmlParts.getByNamespaceAsync as jest.Mock).mockImplementation(
			(namespace: string, callback: (result: any) => void) => {
				callback({
					status: global.Office.AsyncResultStatus.Succeeded,
					value: [fakeCustomXmlPart1, fakeCustomXmlPart2],
				});
			}
		);

		await expect(removeCustomXmlPart("test key")).rejects.toThrow("More than one custom XML part found for the namespace:");
	});
});