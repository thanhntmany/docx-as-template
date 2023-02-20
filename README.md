# docx-as-template


Instructions:


## Use as a js library
```js
// Short instruction:
app.setTemplate("path/to/template.docx");
app.setData("JsObject or path/to/data.json");
app.render("path/to/output.docx");


// chain functions style
app.setTemplate("path/to/template.docx").setData("JsObject or path/to/data.json").render("path/to/output.docx");

// equal:
app.render("path/to/output.docx", "JsObject or path/to/data.json", "path/to/template.docx");
```

## Or run as a command in shell:
```bash
nodejs docx-as-template.js "path/to/output.docx" "JsObject or path/to/data.json" "path/to/template.docx"
```