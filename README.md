# office-templates

Template-based docx and pptx report creation for both Node and the browser.

## Installation

```
$ npm install office-templates
```

## Why?

* **Write documents and presentations naturally using Word or PowerPoint**, just adding some commands where needed for dynamic contents
* **Execute JavaScript snippets** (`EXEC` command, or `!` for short)
* **Insert the result of JavaScript snippets** in your document (`INS`, `=` or just *nothing*)
* **Embed images, hyperlinks and even HTML dynamically** (`IMAGE`, `LINK`, `HTML`). Dynamic images can be great for on-the-fly QR codes, downloading photos straight to your reports, charts… even maps!
* Add **loops** with `FOR`/`END-FOR` commands, with support for table rows, nested loops, and JavaScript processing of elements (filter, sort, etc)
* Include contents **conditionally**, `IF` a certain JavaScript expression is truthy
* Define custom **aliases** for some commands (`ALIAS`) — useful for writing table templates!
* Execute JavaScript directly, or provide your own sandbox through `runJs`.
* Include **literal XML**
* Written in TypeScript, so ships with type definitions.
* Plenty of **examples** in this repo (with Node, Webpack and Browserify)
* Supports `.docx` and `.pptx`, as well as macro-enabled `.docm` and `.pptm` templates.

Contributions are welcome!

# Table of contents

- [Installation](#installation)
- [Node usage](#node-usage)
- [Deno usage](#deno-usage)
- [Browser usage](#browser-usage)
  - [Polyfilled browser-ready bundle](#polyfilled-browser-ready-bundle)
  - [Browser template compatibility caveat](#browser-template-compatibility-caveat)
  - [Running within a Web Worker](#running-within-a-web-worker)
- [Writing templates](#writing-templates)
  - [PowerPoint templates](#powerpoint-templates)
  - [Migration](#migration)
  - [Custom command delimiters](#custom-command-delimiters)
  - [Supported commands](#supported-commands)
    - [`INS` (`=`, or nothing at all)](#ins--or-nothing-at-all)
    - [`EXEC` (`!`)](#exec-)
    - [`IMAGE`](#image)
    - [`LINK`](#link)
    - [`HTML`](#html)
    - [`FOR` and `END-FOR`](#for-and-end-for)
    - [`IF` and `END-IF`](#if-and-end-if)
    - [`ALIAS` (and alias resolution with `*`)](#alias-and-alias-resolution-with-)
  - [Inserting literal XML](#inserting-literal-xml)
- [Error handling](#error-handling)
  - [Error types](#error-types)
  - [Custom error handler](#custom-error-handler)
- [Inspecting templates](#inspecting-templates)
- [JavaScript execution and custom sandboxes](#javascript-execution-and-custom-sandboxes)
  - [Execution migration](#execution-migration)
- [Similar projects](#similar-projects)
- [License (MIT)](#license-mit)


# Node usage

Here is a simple example, with report data injected directly as an object:

```js
import createReport from 'office-templates';
import fs from 'fs';

const template = fs.readFileSync('myTemplate.docx');

const buffer = await createReport({
  template,
  data: {
    name: 'John',
    surname: 'Appleseed',
  },
});

fs.writeFileSync('report.docx', buffer)
```

You can also **provide a sync or Promise-returning data resolver function** instead of a `data` object:

```js
const report = await createReport({
  template,
  data: async () => fetchReportData(),
});
```

The resolver is called once without arguments; its resolved data is used throughout the template.

Other options (with defaults):

```js
const report = await createReport({
  // ...
  additionalJsContext: {
    // all of these will be available to JS snippets in your template commands (see below)
    foo: 'bar',
    qrCode: async url => {
      /* build QR and return image data */
    },
  },
  cmdDelimiter: ['{{', '}}'],
  literalXmlDelimiter: '||',
  processLineBreaks: true,

  /**
   * Whether to fail on the first error encountered in the template. Defaults to true. Can be used to collect all errors in a template (e.g. misspelled commands) before failing.
   */
  failFast: true,

  /**
   * When set to `true`, this setting ensures createReport throws an error when the result of an INS, HTML, IMAGE, or LINK command turns out to be null or undefined,
   * as this usually indicates a mistake in the template or the invoking code (defaults to `false`).
   */
  rejectNullish: false,

  /**
   * MS Word usually autocorrects JS string literal quotes with unicode 'smart' quotes ('curly' quotes). E.g. 'aubergine' -> ‘aubergine’.
   * This causes an error when evaluating commands containing these smart quotes, as they are not valid JavaScript.
   * If you set fixSmartQuotes to 'true', these smart quotes will automatically get replaced with straight quotes (') before command evaluation.
   * Defaults to false.
   */
  fixSmartQuotes: false,

  /**
   * Maximum loop iterations allowed when walking through the template.
   * You can increase this to generate reports with large amount of FOR loop elements.
   * Tip: You can disable infinite loop protection by using the `Infinity` constant.
   * This may be useful if you implement a process timeout instead.
   * (Default: 1,000,000)
   */
  maximumWalkingDepth: 1_000_000,
  /**
   * Whether to indent the generated XML to make it more human-readable.
   * Tip: Set this to true if you want to inspect the generated XML or if you want to use a diff tool to compare the generated Office file with another one.
   * Leaving this option to false will result in a smaller file size.
   * (Default: true)
   */
  indentXml: true,
  /**
   * Whether to preserve whitespace in the generated XML.*
   * Tip: Set this to true if your template contains significant whitespace that you want to preserve in the output document.
   * Leaving this option to false will result in a smaller file size.
   * (Default: true)
   */
  preserveSpace: true,
  /**
   * Compression level for the generated Office file.
   * Integer between 0 (no compression, fastest) and 9 (maximum compression, slowest).
   * (Default: 1)
   */
  compressionLevel: 1,
});
```

Check out the [Node examples folder](examples/example-node).

# Deno usage
For Deno, follow the Browser guide and import a local copy of the polyfilled bundle:

```ts
// @deno-types="./lib/bundled.d.ts"
import { createReport } from './lib/browser.js';
```

JavaScript executes directly by default. Use `runJs` if you want to supply a sandbox compatible with Deno.

# Browser usage

You can use office-templates in the browser (yay!). Just as when using office-templates in Node, you need to provide the template contents as a `Buffer`-like object.

For example when the template is on your server you can get it with something like:

```js
const template = await fetch('./template.docx').then(res => res.arrayBuffer())
```

Or if the user provides the template you can get a `File` object with:

```html
<input type="file">
```

Then read this file in an ArrayBuffer, feed it to office-templates, and download the result:

```js
import createReport from 'office-templates';

const onTemplateChosen = async () => {
  const template = await readFileIntoArrayBuffer(myFile);
  const report = await createReport({
    template,
    data: { name: 'John', surname: 'Appleseed' },
  });
  saveDataToFile(
    report,
    'report.docx',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
  );
};

// Load the user-provided file into an ArrayBuffer
const readFileIntoArrayBuffer = fd =>
  new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = reject;
    reader.onload = () => {
      resolve(reader.result);
    };
    reader.readAsArrayBuffer(fd);
  });
```

You can find an example implementation of `saveDataToFile()` [in the Webpack example](examples/example-webpack/client/index.js).

Check out the examples [using Webpack](examples/example-webpack) and [using Browserify](examples/example-browserify) or you can use the browserified bundle directly as discussed below.

## Polyfilled browser-ready bundle
The browser build includes the polyfills needed by the library, including `Buffer` and streams. There is no VM polyfill or built-in sandbox. If you bundle the library yourself, provide the polyfills required by its dependencies.

Host the `lib/browser.js` bundle yourself and import it directly as a module:

```ts
import { createReport } from './lib/browser.js';
```

The browser bundle targets `es2017`. Compile it to your target if you need to support older browsers.

## Browser template compatibility caveat
Note that the JavaScript code in your .docx or .pptx template will be run as-is by the browser. Transpilers like Babel can't see this code, and won't be able to polyfill it. This means that the JS code in your template needs to be compatible with the browsers you are targeting. In other words: don't use fancy modern syntax and functions in your template if you want older browsers, like IE11, to be able to render it.

## Running within a Web Worker
You can run `createReport` in a Web Worker without any execution-mode options. If you provide `runJs`, its implementation must work in that worker environment.

# Writing templates

You can find several template examples in this repo:

* [SWAPI](examples/example-node), a good example of what you can achieve loading GraphQL data in your application for a report, including a simple script for report generation. Uses the freak-ish online Star Wars GraphQL API.
* [Dynamic images](examples/example-node): with examples of images that are dynamically downloaded or created. Check out the _images-many-tiles_ example for a taste of this powerful feature.
* Browser-based examples [using Webpack](examples/example-webpack) and [using Browserify](examples/example-browserify).

## PowerPoint templates

The format is detected from the package manifest. The same API accepts `.docx` and `.pptx` templates and returns the same format as the input:

```js
const template = fs.readFileSync('myTemplate.pptx');
const buffer = await createReport({ template, data: { title: 'Quarterly update' } });
fs.writeFileSync('report.pptx', buffer);
```

Place `{{ title }}` or `{{ INS title }}` in a slide text box. Commands are processed in slides, speaker notes, layouts, and masters, including commands split across formatted text runs. `INS`, `EXEC`, `ALIAS`, `FOR`/`END-FOR`, `IF`/`END-IF`, `IMAGE`, and `LINK` work in PowerPoint. Loops repeat paragraphs or table rows within a part; they do not duplicate whole slides and cannot span slides. JavaScript context starts fresh for each part.

Use a dedicated text shape for an `IMAGE` command: the shape is replaced by a picture at the same position, with the requested dimensions. Image captions are Word-only. `LINK` creates a clickable text run. `HTML` is Word-only and raises an error in PowerPoint, which has no equivalent of Word's HTML import element.

For browser downloads, use the `.pptx` extension and `application/vnd.openxmlformats-officedocument.presentationml.presentation` MIME type.

## Migration

The default command wrappers are now `{{` and `}}`. Existing templates using `+++` remain usable by passing `cmdDelimiter: '+++'` to `createReport` or `listCommands`.

The `QUERY` command, `queryVars` option, and `QueryResolver` type have been removed. Fetch data in your application, or supply a zero-argument `DataResolver` function. Remove query tags from existing templates before rendering them.

## Custom command delimiters
The default wrappers are `{{` and `}}`. You can use different **left/right command delimiters** by passing an array to `cmdDelimiter`:

```js
const report = await createReport({
  // ...
  cmdDelimiter: ['{', '}'],
})
```

This allows much cleaner-looking templates!

Then you can add commands and JS snippets in your template like this: `{foo}`, `{project.name}`, `{FOR ...}`.

When choosing a delimiter, take care not to introduce conflicts with JS syntax, especially if you are planning to use larger JS code snippets in your templates. For example, with `['{', '}']` you may run into conflicts as the brackets in your JS code may be mistaken for command delimiters. As an alternative, consider using multi-character delimiters, like `{#` and `#}`.


## Supported commands
Currently supported commands are defined below.

### Insert data with the `INS` command ( or using `=`, or nothing at all)

Inserts the result of a given JavaScript snippet as follows.

Using code like this:
```js
 const report = await createReport({
    template,
    data: { name: 'John', surname: 'Appleseed' },
    cmdDelimiter: ['{{', '}}'],
  });
```
And a template like this:

```
{{name}} {{surname}}
```

Will produce a result docx file that looks like this:

```
John Appleseed
```

Alternatively, you can use the more explicit `INS` (insert) command syntax.
```
{{INS name}} {{INS surname}}
```

Note that the last evaluated expression is inserted into the document, so you can include more complex code if you wish:

```
{{INS
const a = Math.random();
const b = Math.round((a - 0.5) * 20);
`A number between -10 and 10: ${b}.`
}}
```

You can also use `=` as shorthand notation instead of `INS`:

```
{{= name}} {{= surname}}
```

Even shorter (and with custom `cmdDelimiter: ['{', '}']`):

```
{name} {surname}
```

You can also access functions in the `additionalJsContext` parameter to `createReport()`, which may even return a Promise. The resolved value of the Promise will be inserted in the document.

Use JavaScript's ternary operator to implement an _if-else_ structure:

```
{{= $details.year != null ? `(${$details.year})` : ''}}
```

### `EXEC` (`!`)

Executes a given JavaScript snippet, just like `INS` or `=`, but doesn't insert anything in the document. You can use `EXEC`, for example, to define functions or constants before using them elsewhere in your template.

```
{{EXEC
this.myFun = () => Math.random();
this.MY_CONSTANT = 3;
}}

{{! this.ANOTHER_CONSTANT = 5; }}
```

Usage elsewhere will then look like

```
{{= MY_CONSTANT }}
{{= ANOTHER_CONSTANT }}
{{= myFun() }}
```

**A note on scoping:**

By default, snippets execute with `with(this) { return eval(__code__); }`, using an execution context that is retained between commands within each document part. Use explicit context assignments such as `this.MY_CONSTANT = 3` to make new values available to later commands. Local declarations such as `const`, `let`, and function declarations do not persist between evaluations. A custom `runJs` implementation determines its own scoping behavior.

This example declares the test function in the context object, making it callable from another snippet.
```js
this.test = () => {};
```

While the below example only declares test in the local scope of the snippet, meaning it gets garbage collected after the snippet has executed.
```js
function test() {};
```

### `IMAGE`

Includes an image with the data resulting from evaluating a JavaScript snippet:

```
{{IMAGE qrCode(project.url)}}
```

In this case, we use a function from `additionalJsContext` object passed to `createReport()` that looks like this:

```js
  additionalJsContext: {
    qrCode: url => {
      const dataUrl = createQrImage(url, { size: 500 });
      const data = dataUrl.slice('data:image/gif;base64,'.length);
      return { width: 6, height: 6, data, extension: '.gif' };
    },
  }
```

The JS snippet must return an _image object_ or a Promise of an _image object_, containing:

* `width`: desired width of the image on the page _in cm_. Note that the aspect ratio should match that of the input image to avoid stretching.
* `height` desired height of the image on the page _in cm_.
* `data`: either an ArrayBuffer or a base64 string with the image data
* `extension`: one of `'.png'`, `'.gif'`, `'.jpg'`, `'.jpeg'`, `'.svg'`.
* `thumbnail` _[optional]_: when injecting an SVG image, a fallback non-SVG (png/jpg/gif, etc.) image can be provided. This thumbnail is used when SVG images are not supported (e.g. older versions of Word) or when the document is previewed by e.g. Windows Explorer. See usage example below.
* `alt` _[optional]_: optional alt text.
* `rotation` _[optional]_: optional rotation in degrees, with positive angles moving clockwise.
* `caption` _[optional]_: optional caption displayed below the image

In the .docx template:
```
{{IMAGE injectSvg()}}
```

Note that you can center the image by centering the IMAGE command in the template.

In the `createReport` call:
```js
additionalJsContext: {
  injectSvg: () => {
      const svg_data = Buffer.from(`<svg  xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink">
                                  <rect x="10" y="10" height="100" width="100" style="stroke:#ff0000; fill: #0000ff"/>
                                </svg>`, 'utf-8');

      // Providing a thumbnail is technically optional, as newer versions of Word will just ignore it.
      const thumbnail = {
        data: fs.readFileSync('sample.png'),
        extension: '.png',
      };
      return { width: 6, height: 6, data: svg_data, extension: '.svg', thumbnail };
    }
  }
```

### `LINK`

Includes a hyperlink with the data resulting from evaluating a JavaScript snippet:

```
{{LINK ({ url: project.url, label: project.name })}}
```

If the `label` is not specified, the URL is used as a label.

### `HTML`

Takes the HTML resulting from evaluating a JavaScript snippet and converts it to Word contents.

**Important:** This uses altchunk, which is only supported in Microsoft Word, and not in e.g. LibreOffice or Google Docs.

```
{{HTML `
<meta charset="UTF-8">
<body>
  <h1>${film.title}</h1>
  <h3>${film.releaseDate.slice(0, 4)}</h3>
  <p>
    <strong style="color: red;">This paragraph should be red and strong</strong>
  </p>
</body>
`}}
```

### `FOR` and `END-FOR`

Loop over a group of elements (resulting from the evaluation of a JavaScript expression):

```
{{FOR person IN project.people}}
{{INS $person.name}} (since {{INS $person.since}})
{{END-FOR person}}
```

Note that inside the loop, the variable relative to the current element being processed must be prefixed with `$`.

It is possible to get the current element index of the inner-most loop with the variable `$idx`, starting from `0`. For example:
```
{{FOR company IN companies}}
Company ({{$idx}}): {{INS $company.name}}
Executives:
{{FOR executive IN $company.executives}}
-	{{$idx}} {{$executive}}
{{END-FOR executive}}
{{END-FOR company}}
```

Since JavaScript expressions are supported, you can for example filter the loop domain:

```
{{FOR person IN project.people.filter(person => person.since > 2013)}}
...
```

`FOR` loops also work over table rows:

```
----------------------------------------------------------
| Name                         | Since                   |
----------------------------------------------------------
| {{FOR person IN             |                         |
| project.people}}            |                         |
----------------------------------------------------------
| {{INS $person.name}}       | {{INS $person.since}} |
----------------------------------------------------------
| {{END-FOR person}}         |                         |
----------------------------------------------------------
```

And let you dynamically generate columns:

```
+-------------------------------+--------------------+------------------------+
| {{ FOR row IN rows}}        |                    |                        |
+===============================+====================+========================+
| {{ FOR column IN columns }} | {{INS $row}}     | {{ END-FOR column }} |
|                               |                    |                        |
|                               | Some cell content  |                        |
|                               |                    |                        |
|                               | {{INS $column}}  |                        |
+-------------------------------+--------------------+------------------------+
| {{ END-FOR row}}            |                    |                        |
+-------------------------------+--------------------+------------------------+
```

Finally, you can nest loops (this example assumes a different data set):

```
{{FOR company IN companies}}
{{INS $company.name}}
{{FOR person IN $company.people}}
* {{INS $person.firstName}}
{{FOR project IN $person.projects}}
    - {{INS $project.name}}
{{END-FOR project}}
{{END-FOR person}}

{{END-FOR company}}
```

### `IF` and `END-IF`

Include contents conditionally (depending on the evaluation of a JavaScript expression):

```
{{IF person.name === 'Guillermo'}}
{{= person.fullName }}
{{END-IF}}
```

Similarly to the `FOR` command, it also works over table rows. You can also nest `IF` commands
and mix & match `IF` and `FOR` commands. In fact, for the technically inclined: the `IF` command
is implemented as a `FOR` command with 1 or 0 iterations, depending on the expression value.

### `ALIAS` (and alias resolution with `*`)

Define a name for a complete command (especially useful for formatting tables):

```
{{ALIAS name INS $person.name}}
{{ALIAS since INS $person.since}}

----------------------------------------------------------
| Name                         | Since                   |
----------------------------------------------------------
| {{FOR person IN             |                         |
| project.people}}            |                         |
----------------------------------------------------------
| {{*name}}                  | {{*since}}            |
----------------------------------------------------------
| {{END-FOR person}}         |                         |
----------------------------------------------------------
```

## Inserting literal XML
You can also directly insert Office Open XML markup into the document using the `literalXmlDelimiter`, which is by default set to `||`.

E.g. if you have a template like this:

```
{{INS text}}
```

```js
await createReport({
  template,
  data: { text: 'foo||<w:br/>||bar' },
}
```


# Error handling

By default, the Promise returned by `createReport` will reject with an error immediately when a problem is encountered in the template, such as a bad command (i.e. it 'fails fast'). In some cases, however, you may want to collect all errors that may exist in the template before failing. For example, this is useful when you are letting your users create templates interactively. You can disable fast-failing by providing the `failFast: false` parameter as shown below. This will make `createReport` reject with an array of errors instead of a single error so you can get a more complete picture of what is wrong with the template.

```typescript
try {
  const report = await createReport({
    template,
    data: {
      name: 'John',
      surname: 'Appleseed',
    },
    failFast: false,
  });
} catch (errors) {
  if (Array.isArray(errors)) {
    // An array of errors likely caused by bad commands in the template.
    console.log(errors);
  } else {
    // Not an array of template errors, indicating something more serious.
    throw errors;
  }
}
```
## Error types
The library exposes the following error types. See the `errors.ts` module for details.

```
NullishCommandResultError // thrown when rejectNullish is set to true and a command returns null or undefined
ObjectCommandResultError // thrown when the result of an `INS` command is an Object. This ensures you don't accidentally put `'[object Object]'` in your report.
CommandSyntaxError
InvalidCommandError
CommandExecutionError
ImageError
InternalError
TemplateParseError
IncompleteConditionalStatementError // thrown when an IF-statement has no corresponding END-IF command
UnterminatedForLoopError // thrown when a `FOR` loop is not properly terminated with an `END-FOR`
```

## Custom error handler
A custom error handler callback can be provided to handle any errors that may occur when executing commands from a template. The value returned by this callback will be inserted into the rendered document instead. The callback is provided with two arguments: the error that was caught and the raw code of the command.

```typescript
  const report = await createReport({
    template,
    data: {
      name: 'John',
      surname: 'Appleseed',
    },
    errorHandler: (err, command_code) => {
      return 'command failed!';
    },
  });
```

Using a custom `errorHandler` in combination with `rejectNullish = true` allows users to intelligently replace the result of commands that returned `null` or `undefined` (make sure to check for `NullishCommandResultError`).

# Inspecting templates
The `listCommands` function lets you list all the commands in a docx or pptx template using the same parser as `createReport`.

```typescript
import { listCommands } from 'office-templates';
const template_buffer = fs.readFileSync('template.docx');
const commands = await listCommands(template_buffer);

// `commands` will contain something like:
[
  { raw: 'INS some_variable', code: 'some_variable', type: 'INS' },
  { raw: 'IMAGE svgImgFile()', code: 'svgImgFile()', type: 'IMAGE' },
]
```

The `getMetadata` function lets you extract the metadata fields from a document, such as the number of pages or words. Note that this feature has a few limitations:
- Not all fields may be available, depending on the document.
- These metadata fields, including the number of pages, are only updated by MS Word (or LibreOffice) when saving the document. Office-templates does not alter these metadata fields, so the number of pages may not reflect the actual size of your rendered document. Office-templates can not reliably determine the number of pages in a document, as this requires a full-fledged docx renderer (e.g. MS Word).

```typescript
    import { getMetadata } from 'office-templates';
    const template = fs.readFileSync('template.docx');
    await getMetadata(template)
    // result:
      Object {
        "category": undefined,
        "characters": 24,
        "company": undefined,
        "created": "2015-08-16T18:55:00Z",
        "creator": "Someone Else",
        "description": undefined,
        "lastModifiedBy": "Grau Panea, Guillermo",
        "lastPrinted": undefined,
        "lines": 1,
        "modified": "2016-12-15T11:21:00Z",
        "pages": 1,
        "paragraphs": 1,
        "revision": "32",
        "subject": undefined,
        "template": "Normal.dotm",
        "title": undefined,
        "words": 4,
      }
```

# JavaScript execution and custom sandboxes

By default, template commands execute directly through JavaScript's `Function` and `eval`, without isolation. The library does not include a VM or a sandbox. Only render trusted templates with the default executor.

To provide your own sandbox, pass a `runJs` callback to `createReport`. The callback receives `{ sandbox, ctx }`:

- `sandbox.__code__` contains the JavaScript snippet to evaluate.
- `sandbox` also contains the report data, `additionalJsContext`, saved bindings from prior commands, and current loop variables such as `$item` and `$idx`.
- `ctx` contains the template execution context.

Return `{ modifiedSandbox, result }`. `modifiedSandbox` must contain the bindings you want to retain for later commands. The callback itself and its `result` may both be promises; the library awaits them. Each document part starts with a fresh execution context. The library resets the internal `__code__` and `__result__` fields before the next command.

For example, if your sandbox adapter exposes an asynchronous `evaluate(code, bindings)` method returning `{ bindings, value }`, connect it like this:

```js
import createReport from 'office-templates';
import { sandboxEngine } from './my-sandbox.js';

const report = await createReport({
  template,
  data: { name: 'John' },
  runJs: async ({ sandbox }) => {
    const { __code__, __result__, ...bindings } = sandbox;
    const evaluation = await sandboxEngine.evaluate(__code__, bindings);
    return {
      modifiedSandbox: {
        ...evaluation.bindings,
        __code__: undefined,
        __result__: undefined,
      },
      result: evaluation.value,
    };
  },
});
```

Adapt the callback to your sandbox's API. Errors thrown by the callback and rejected promises follow the same `errorHandler` and `failFast` behavior as the default executor. Your chosen sandbox is responsible for isolation, exposed capabilities, and execution limits.

## Execution migration

The built-in VM and `noSandbox` option have been removed. Remove `noSandbox` from existing calls. Direct execution is now the default; applications that need isolation must supply `runJs`.

# Similar projects

* docxtemplater supports template-based document generation.

* docx and similar ones - generate docx files from scratch, programmatically. Drawbacks of this approach: they typically do not support all Word features, and producing a complex document can be challenging.


# License (MIT)

This Project is licensed under the MIT License. See [LICENSE](LICENSE) for more information.
