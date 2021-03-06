# A Blazor WASM Project that can be used for Office Add-ins 

This project uses Outlook as an example using a demonstration for reading and writing to mail, however by changing the manifest and javascript code, it would most likely work for all types of Office add-ins!

## How to use:

Simply download the project and run it on visual studio. This project template shows an example of how Blazor WASM can be used for an Outlook Add-in, as the main page will allow you to write to the body of the mail through a textbox when composing a mail, and also read the current body of a mail when either composing or reading mail.

These features will only work if you use run them on an outlook client, so it is recommended that you use the taskpane in outlook to load the website in order for the demonstrations to work.

## Structure:

Everything is the same as a regular Blazor WASM project, except under the wwwroot folder there is a JavaScript folder which stores JavaScript files that contains office 
code (These scripts need to be imported into index.html). There is also the tsconfig.json which allows support for OfficeJS intellisense.

In the main project directory there is a Manifest folder which contains the manifest you must load into the Office Program (The one provided is meant for Outlook)


## Notes:
1. When creating new JavaScript files, it is important to import them to index.html located in the wwwroot folder, and if they require OfficeJS functionality, adding
the following code on top of the file will allow intellisense to work:

    /** @type {OfficeCore.Range} */

    /** @type {Office.Range} */

    /** @type {OfficeExtension.Range} */

2. When you need to call JavaScript functions when the program begins, go to App.razor and add an await JS.InvokeVoidAsync("functionName"); after Office.onReady() has been invoked

3. It should also be mentioned that at this current moment, the Office manifest must be sideloaded to the office program manually.

## Support

If there are any issues on running the project or any questions, please create an issue and I can try to resolve/answer any questions to the best of my ability.
