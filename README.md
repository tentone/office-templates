# Office Templates

 - Fork of the [docx-templates]() project.
 - Template-based docx and pptx report creation for both Node and the browser.
 - The documents are created using Word or Powerpoint, placeholders can be added directly in the file.
 - Execute JavaScript snippets (`EXEC` command, or `!` for short)
   - Note that the JavaScript code in your .docx or .pptx template will be run as-is by the browser.
Transpilers  can't see this code, and won't be able to polyfill it.
 - Embed images, hyperlinks and even HTML dynamically (`IMAGE`, `LINK`, `HTML`).
   - Logic loops with `FOR`/`END-FOR` commands, with support for table rows, nested loops amnd conditionals using `IF`.
 - Possible to isolate the javascript execution with your own sandbox through `runJs`.
 - Include **literal XML** for custom attributes (e.g. color changes)
 - Supports `.docx` and `.pptx`, as well as macro-enabled `.docm` and `.pptm` templates.
 - Works in node.js and browser

## Usage

 - To install from npm run `npm install office-templates`
 - Here is a simple example, with report data injected directly as an object:

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

 - You can also **provide a sync or Promise-returning data resolver function** instead of a `data` object:

```js
const report = await createReport({
  template,
  data: async () => fetchReportData(),
});
```

## Options

 - The resolver is called once without arguments; its resolved data is used throughout the template.
 - Other options (with defaults):

```js
const report = await createReport({
  /* Additional data to be provided to the template.
   * 
   * All of these will be available to JS snippets in your template commands.
   */
  additionalJsContext: {
    foo: 'bar',
    qrCode: async url => {
      /* build QR and return image data */
    },
  },
  
  /*
   * Delimiter for template commands
   */
  cmdDelimiter: ['{{', '}}'],
  
  /*
   * Delimiter to wrap literal XML that will be injected in the model.
   */
  literalXmlDelimiter: '||',

  /*
   * Process line breaks
   */
  processLineBreaks: true,

  /* Whether to fail on the first error encountered in the template.
   *
   * Can be used to collect all errors in a template (e.g. misspelled commands) before failing.
   */
  failFast: true,

  /* This setting ensures createReport throws an error when the result of an INS, HTML, IMAGE, or LINK command turns out to be null or undefined, as this usually indicates a mistake in the template or the invoking code.
   */
  rejectNullish: false,

  /* MS Word usually autocorrects JS string literal quotes with unicode 'smart' quotes ('curly' quotes). E.g. 'aubergine' -> ‘aubergine’.
   *
   * This causes an error when evaluating commands containing these smart quotes, as they are not valid JavaScript.
   * 
   * If you set fixSmartQuotes to 'true', these smart quotes will automatically get replaced with straight quotes (') before command evaluation.
   */
  fixSmartQuotes: false,

  /**
   * Maximum loop iterations allowed when walking through the template.
   * 
   * You can increase this to generate reports with large amount of FOR loop elements.
   * 
   * Tip: You can disable infinite loop protection by using the `Infinity` constant.
   * 
   * This may be useful if you implement a process timeout instead.
   */
  maximumWalkingDepth: 1_000_000,

  /**
   * Whether to indent the generated XML to make it more human-readable.
   * 
   * Tip: Set this to true if you want to inspect the generated XML or if you want to use a diff tool to compare the generated Office file with another one.
   * 
   * Leaving this option to false will result in a smaller file size.
   */
  indentXml: true,

  /**
   * Whether to preserve whitespace in the generated XML.
   * 
   * Tip: Set this to true if your template contains significant whitespace that you want to preserve in the output document.
   * 
   * Leaving this option to false will result in a smaller file size.
   */
  preserveSpace: true,

  /**
   * Compression level for the generated Office file.
   * 
   * Integer between 0 (no compression, fastest) and 9 (maximum compression, slowest).
   */
  compressionLevel: 1,
});
```

## Templates

The format is detected from the package manifest. The same API accepts `.docx` and `.pptx` templates and returns the same format as the input:

```js
const template = fs.readFileSync('myTemplate.pptx');
const buffer = await createReport({ template, data: { title: 'Quarterly update' } });
fs.writeFileSync('report.pptx', buffer);
```

 - Place `{{ title }}` or `{{ INS title }}` in a slide text box. Commands are processed in slides, speaker notes, layouts, and masters, including commands split across formatted text runs. `INS`, `EXEC`, `ALIAS`, `FOR`/`END-FOR`, `IF`/`END-IF`, `IMAGE`, and `LINK` work in PowerPoint.
 - Loops repeat paragraphs or table rows within a part; they do not duplicate whole slides and cannot span slides. JavaScript context starts fresh for each part.
 - Use a dedicated text shape for an `IMAGE` command: the shape is replaced by a picture at the same position, with the requested dimensions. Image captions are Word-only. `LINK` creates a clickable text run. `HTML` is Word-only and raises an error in PowerPoint, which has no equivalent of Word's HTML import element.
 - For browser downloads, use the `.pptx` extension and `application/vnd.openxmlformats-officedocument.presentationml.presentation` MIME type.


## Supported Commands

## `INS` or `=`

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

`INS` automatically waits when its expression returns a promise, then inserts the resolved value. This also applies to `=` and omitted-command shorthand, promise-valued data fields, and async functions from `additionalJsContext`. You do not need to write `await` in the template.

For example, place these commands in your template:

```text
{{ INS customerName }}
{{ INS loadSummary() }}
```

```js
const report = await createReport({
  template,
  data: { customerName: Promise.resolve('John') },
  additionalJsContext: {
    loadSummary: async () => 'Quarterly summary',
  },
});
```

This behavior works in DOCX and PPTX with the default executor or a custom `runJs`. Rejected promises follow the usual command error handling. Validation such as `rejectNullish` applies to the resolved value.

Use JavaScript's ternary operator to implement an _if-else_ structure:

```
{{= $details.year != null ? `(${$details.year})` : ''}}
```

## `EXEC` (`!`)

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

## `IMAGE`

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

`IMAGE` also automatically waits for a promise and inserts the resolved image object. This works with promise-valued data fields and functions returning promises, without `await` in the template:

```text
{{ IMAGE loadImage() }}
```

```js
const report = await createReport({
  template,
  additionalJsContext: {
    loadImage: async () => ({
      data: await fs.promises.readFile('photo.png'),
      extension: '.png',
      width: 6,
      height: 4,
    }),
  },
});
```

The resolved image object is validated before insertion. Rejected promises follow the usual command error handling, and `rejectNullish` applies after resolution. For PPTX, place the command in a dedicated text shape as described above.

The JS snippet must return an _image object_ or a promise resolving to an _image object_, containing:

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

## `LINK`

Includes a hyperlink with the data resulting from evaluating a JavaScript snippet:

```
{{LINK ({ url: project.url, label: project.name })}}
```

If the `label` is not specified, the URL is used as a label.

## `HTML`

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

## `FOR` and `END-FOR`

Loop over a group of elements (resulting from the evaluation of a JavaScript expression):

```
{{FOR person IN project.people}}
{{INS $person.name}} (since {{INS $person.since}})
{{END-FOR person}}
```

Note that inside the loop, the variable relative to the current element being processed must be prefixed with `$`.

The `$idx` variable is the zero-based index of the innermost active `FOR` loop. It remains available in `IF` conditions and inside their contents; `IF` blocks do not replace or reset it. When a nested `FOR` finishes, `$idx` refers to the enclosing `FOR` again. For example:
```
{{FOR company IN companies}}
Company ({{$idx}}): {{INS $company.name}}
Executives:
{{FOR executive IN $company.executives}}
-	{{$idx}} {{$executive}}
{{END-FOR executive}}
{{END-FOR company}}
```

Use `$idx` in a condition to select specific iterations:

```text
{{ FOR person IN people }}{{ IF $idx === 1 }}{{ $person.name }} (index {{ $idx }}){{ END-IF }}{{ END-FOR person }}
```

This inserts only the second person's name. If the condition is inside multiple nested `FOR` loops, it uses the index of the innermost `FOR`, even when additional `IF` blocks are nested around it. This works in DOCX and PPTX, including with a custom `runJs` executor.

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

## `IF` and `END-IF`

Include contents conditionally (depending on the evaluation of a JavaScript expression):

```
{{IF person.name === 'Guillermo'}}
{{= person.fullName }}
{{END-IF}}
```

You can nest `IF` and `FOR` commands within the same paragraph or table row, including on a single line. Close each block in reverse order of opening it:

```text
{{ IF visible }}{{ FOR item IN items }}{{ IF $item.enabled }}{{ $item.name }} {{ END-IF }}{{ END-FOR item }}{{ END-IF }}
```

This works in both DOCX and PPTX. When an outer condition is false or an outer loop is empty, expressions inside the skipped block are not evaluated.

Similarly to the `FOR` command, it also works over table rows. You can also nest `IF` commands
and mix & match `IF` and `FOR` commands. In fact, for the technically inclined: the `IF` command
is implemented as a `FOR` command with 1 or 0 iterations, depending on the expression value.

## `ALIAS` (and alias resolution with `*`)

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

## Literal XML
 - You can also directly insert Office Open XML markup into the document using the `literalXmlDelimiter`, which is by default set to `||`.
 - Examples of usage for this feature can be 

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

 - By default, the Promise returned by `createReport` will reject with an error immediately when a problem is encountered in the template, such as a bad command (i.e. it 'fails fast'). In some cases, however, you may want to collect all errors that may exist in the template before failing. For example, this is useful when you are letting your users create templates interactively. You can disable fast-failing by providing the `failFast: false` parameter as shown below. This will make `createReport` reject with an array of errors instead of a single error so you can get a more complete picture of what is wrong with the template.

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

 - By default, template commands execute directly through JavaScript's `Function` and `eval`, without isolation.
 - The library does not include a VM or a sandbox.
 - Only render trusted templates with the default executor.
 - To provide your own sandbox, pass a `runJs` callback to `createReport`. The callback receives `{ sandbox, ctx }`:
    - `sandbox.__code__` contains the JavaScript snippet to evaluate.
    - `sandbox` also contains the report data, `additionalJsContext`, saved bindings from prior commands, and current loop variables such as `$item` and `$idx`.
    - `ctx` contains the template execution context.

Return `{ modifiedSandbox, result }`. `modifiedSandbox` must contain the bindings you want to retain for later commands. The callback itself and its `result` may both be promises; the library awaits them. Each document part starts with a fresh execution context. The library resets the internal `__code__` and `__result__` fields before the next command.

 - For example, if your sandbox adapter exposes an asynchronous `evaluate(code, bindings)` method returning `{ bindings, value }`, connect it like this:

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

 - Adapt the callback to your sandbox's API. Errors thrown by the callback and rejected promises follow the same `errorHandler` and `failFast` behavior as the default executor.
 - Your chosen sandbox is responsible for isolation, exposed capabilities, and execution limits.

# License (MIT)
 - This Project is licensed under the MIT License. See [LICENSE](LICENSE) for more information.
