/// <reference types="office-js" />
import { XMLBuilder, XMLParser } from "fast-xml-parser";

const xmlBuilder = new XMLBuilder();
const xmlParser = new XMLParser();

/**
 * Sets a custom XML part for the specified key and value.
 *
 * This asynchronous function converts the provided value into an XML string and associates it
 * with a namespace derived from the given key. If a custom XML part with the same namespace already
 * exists, it is deleted; otherwise, a new custom XML part is added.
 *
 * @param key - A string used to determine the namespace for the custom XML part.
 * @param value - The value to be converted into an XML format and stored as a custom XML part.
 *
 * @returns A promise that resolves when the operation is complete.
 */
export async function setCustomXmlPartByValue(key: string, value: any): Promise<void> {
	const xmlValue = xmlBuilder.build(value);

	const namespace = getNameSpace(key);
	const customXmlPart = await getCustomXmlPartByNameSpace(namespace);

	if (customXmlPart != null) {
		await deleteCustomXmlPartByNameSpace(namespace);
	}

	await addCustomXmlPartByNameSpace(namespace, xmlValue);
}

/**
 * Retrieves a custom XML part associated with the specified key.
 *
 * @param key - The key used to derive a namespace for locating the custom XML part.
 * @returns A promise that resolves to the matching Office.CustomXmlPart if found, or null if not found.
 *
 * @remarks
 * This asynchronous function converts the provided key to a namespace using getNameSpace and then
 * obtains the related custom XML part by invoking getCustomXmlPartByNameSpace. It is primarily used
 * for accessing custom XML parts based on a key-to-namespace transformation.
 */
export async function getCustomXmlPart(key: string): Promise<Office.CustomXmlPart | null> {
	const namespace = getNameSpace(key);
	return await getCustomXmlPartByNameSpace(namespace);
}

/**
 * Retrieves and parses the value of a custom XML part by its key.
 *
 * This asynchronous function first attempts to locate the custom XML part identified by the provided key. If the XML part exists,
 * it fetches its XML content asynchronously using the Office API. The retrieved XML string is then parsed into an object. 
 *
 * @param key - The unique key that identifies the custom XML part.
 * @returns A promise that resolves to the parsed XML value, or null if the XML part is not found.
 * @throws An error if the asynchronous retrieval of the XML content fails.
 */
export async function getCustomXmlPartValue(key: string): Promise<any | null> {
	const xmlPart = await getCustomXmlPart(key);
	
	if (xmlPart == null) {
		return null;
	}
	
	return new Promise((resolve, reject) => {
		xmlPart.getXmlAsync((result) => {
			if (result.status != Office.AsyncResultStatus.Succeeded) {
				reject(new Error(`Error getting XML: ${result.error.message}`));
			} else {
				const xmlValue = result.value;
				const parsedValue = xmlParser.parse(xmlValue);
				// Return the contents of customData if it exists
				if (parsedValue.customData) {
					resolve(parsedValue.customData);
				} else {
					resolve(parsedValue);
				}
			}
		});
	});
}

/**
 * Removes a custom XML part associated with the provided key.
 *
 * This asynchronous function retrieves the namespace corresponding to the given key,
 * then fetches the custom XML part using that namespace. If the XML part exists,
 * it attempts to delete it using Office's asynchronous API.
 *
 * @param key - The identifier used to determine the namespace of the custom XML part.
 * @returns A promise that resolves if the XML part is successfully removed, or rejects with an error if deletion fails.
 */
export async function removeCustomXmlPart(key: string): Promise<void> {
	const namespace = getNameSpace(key);
	let xmlPart = await getCustomXmlPartByNameSpace(namespace);

	if (xmlPart == null) {
		return;
	}

	await new Promise<void>((resolve, reject) => {
		xmlPart.deleteAsync((result) => {
			if (result.status != Office.AsyncResultStatus.Succeeded) {
				reject(new Error(`Error deleting custom XML part: ${result.error.message}`));
			} else {
				resolve();
			}
		});
	});
}


/**
 * Retrieves the custom XML part for the specified namespace from the current Office document.
 *
 * This asynchronous function returns a Promise that:
 * - Resolves with the single matching custom XML part if one is found.
 * - Resolves with null if no matching custom XML parts are found.
 * - Rejects with an error if more than one matching custom XML part is found or if an
 *   error occurs during retrieval.
 *
 * @async
 * @param namespace - The XML namespace to search for within the custom XML parts.
 * @returns A Promise that resolves with the matching Office.CustomXmlPart, null, or rejects with an error.
 */
async function getCustomXmlPartByNameSpace(namespace: string): Promise<Office.CustomXmlPart | null> {
	return new Promise((resolve, reject) => {
		Office.context.document.customXmlParts.getByNamespaceAsync(namespace, (result) => {
			if (result.status === Office.AsyncResultStatus.Succeeded) {
				if (result.value.length === 0) {
					resolve(null);
				}
				else if (result.value.length > 1) {
					reject(new Error("More than one custom XML part found for the namespace: " + namespace));
				}
				else {
					resolve(result.value[0]);
				}
			} else {
				reject(new Error(`Error getting custom XML parts: ${result.error.message}`));
			}
		});
	});
}

/**
 * Adds a custom XML part with the specified namespace and XML content to the document.
 *
 * This function creates an XML string by wrapping the provided content inside a <customData> element
 * with the given namespace. It then asynchronously adds this XML part to the Office document.
 *
 * @param namespace - The XML namespace associated with the custom data.
 * @param xmlValue - The XML content to be encapsulated within the custom data element.
 * @returns A promise that resolves when the custom XML part is successfully added.
 *
 * @throws Error if the addition of the custom XML part fails.
 */
async function addCustomXmlPartByNameSpace(namespace: string, xmlValue: string): Promise<void> {
	const xmlWithNamespace = `<customData xmlns="${namespace}">${xmlValue}</customData>`;

	return await Office.context.document.customXmlParts.addAsync(xmlWithNamespace, (result) => {
		if (result.status != Office.AsyncResultStatus.Succeeded) {
			throw new Error(`Error adding custom XML part: ${result.error.message}`);
		}
	});
}

/**
 * Retrieves the custom XML parts for the specified namespace.
 *
 * This asynchronous function queries the document for custom XML parts that match the given namespace.
 * By convention, only one custom XML part is stored per namespace.
 *
 * @param namespace - The XML namespace used to locate the custom XML part.
 * @returns A Promise that resolves with an array of Office.CustomXmlPart instances if the operation succeeds.
 *          If the operation fails, the Promise is rejected with the encountered error.
 */
async function deleteCustomXmlPartByNameSpace(namespace: string): Promise<Office.CustomXmlPart[]> {
	// Get the first custom XML part in the document for the specified namespace
	// (By convention, we only ever store one part per namespace.)
	return new Promise((resolve, reject) => {
		Office.context.document.customXmlParts.getByNamespaceAsync(namespace, (result) => {
			if (result.status === Office.AsyncResultStatus.Succeeded) {
				resolve(result.value);
			} else {
				reject(new Error(`Error getting custom XML parts: ${result.error.message}`));
			}
		});
	});
}

/**
 * Sanitizes the provided key by replacing all non-alphanumeric characters with underscores and converting the result to lowercase.
 *
 * This function ensures that the resulting key is neither null, undefined, nor an empty string after sanitization. If the sanitized key is empty or only whitespace,
 * an error is thrown with details about the original and sanitized keys.
 *
 * @param key - The original key string to be sanitized. It is expected to contain only alphanumeric characters or underscores.
 * @returns The sanitized key string in lowercase.
 * @throws {Error} If the provided key is null, undefined, or empty after sanitization.
 */
function getNameSpace(key: string) {
	const sanitizedKey = key.replace(/[^a-zA-Z0-9]/g, "_").toLowerCase();

	// make sure key is not null or undefined or empty string
	if (sanitizedKey == null || sanitizedKey == undefined || sanitizedKey == "" || sanitizedKey.trim() == "") {
		throw new Error(`Key is case-insensitive and cannot be null, undefined, or empty string. Keys can only contain alphanumeric characters or underscores.\nOriginal key: ${key} Sanitized key: ${sanitizedKey}`);
	}

	return sanitizedKey;
}
