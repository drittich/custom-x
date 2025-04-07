# custom-x

A library to facilitate the storage of JavaScript objects in [Office Custom XML Parts](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/working-with-customxml).

## Features

- Serialize JavaScript objects to XML using [fast-xml-parser](https://github.com/NaturalIntelligence/fast-xml-parser).
- Add, retrieve, and delete custom XML parts in Office documents.
- Fully written in TypeScript with strict type checking.

## Installation

Install via npm:

```sh
npm install custom-x
```

## Usage

Import and use the library in your project:

```ts
import { setCustomXmlPart, getCustomXmlPart, removeCustomXmlPart } from 'custom-x';

// Set a custom XML part
setCustomXmlPart('exampleKey', { data: 123 })
  .then(() => console.log('Custom XML part set'))
  .catch(console.error);

// Retrieve a custom XML part
getCustomXmlPart('exampleKey')
  .then(xmlPart => console.log('Retrieved XML part:', xmlPart))
  .catch(console.error);

// Remove a custom XML part
removeCustomXmlPart('exampleKey')
  .then(() => console.log('Custom XML part removed'))
  .catch(console.error);
```

## Building the Library

Compile the TypeScript source code with:

```sh
npm run build
```

## License

MIT