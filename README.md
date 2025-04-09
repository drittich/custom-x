# Custom-X

Custom-X is a TypeScript library that simplifies storing JavaScript objects in Office Custom XML parts. It provides functions to add, retrieve, parse, and remove custom XML data associated with a unique key.

## Features

- **Store Data as XML**: Convert JavaScript objects to XML and encapsulate them in a custom XML part.
- **Retrieve & Parse**: Fetch and parse stored XML data to retrieve JavaScript objects.
- **Namespace Handling**: Sanitize and use keys as namespaces ensuring consistency.
- **Seamless Office Integration**: Works with Office.js to manage Custom XML Parts in Office documents.

## Installation

Install the package via npm:

```sh
npm install custom-x
```

## Usage

Below is a simple example:

```ts
import { setCustomXmlPartByValue, getCustomXmlPartValue, removeCustomXmlPart } from 'custom-x';

// Set a custom XML part.
await setCustomXmlPartByValue("myData", { foo: "bar" });

// Retrieve and parse the custom XML part.
const myData = await getCustomXmlPartValue("myData");
console.log(myData); // { foo: "bar" }

// Remove the custom XML part.
await removeCustomXmlPart("myData");
```

## API Reference

- **setCustomXmlPartByValue(key: string, value: any): Promise<void>**  
  Converts a JavaScript object into XML and stores it as a custom XML part. If a part with the same key exists, it is first removed.

- **getCustomXmlPart(key: string): Promise<Office.CustomXmlPart | null>**  
  Retrieves the custom XML part associated with the key.

- **getCustomXmlPartValue(key: string): Promise<any | null>**  
  Retrieves and parses the XML content of the custom XML part back into a JavaScript object.

- **removeCustomXmlPart(key: string): Promise<void>**  
  Deletes the custom XML part corresponding to the specified key.

## Building & Testing

This project uses TypeScript and Jest for testing.

To build the project, run:

```sh
npm run build
```

To execute the tests, run:

```sh
npm test
```

The tests can be found in the [tests directory](c:\Users\darcy\source\repos\_Libs\Custom-X\tests\index.test.ts).

## License

Distributed under the MIT License. See [LICENSE](c:\Users\darcy\source\repos\_Libs\Custom-X\LICENSE) for more information.

## Contributing

Contributions are welcome! Please open issues and submit pull requests for bug fixes or feature enhancements.

## Support

For support, please refer to the [GitHub repository](https://github.com/drittich/custom-x).

---
Custom-X © 2025 D'Arcy Rittich