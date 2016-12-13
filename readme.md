# Office Javascript API

Office Javascript API enables you to access content in Office documents.

## Usage

### Use Office JavaScript API in Node.js Application
#### Using TypeScript
```typescript
import * as Excel from '@microsoft/office-js/excel'

var accessToken: string;
// getAccessToken() method is is implemented by you and it returns OAUTH access token
accessToken = getAccessToken();
// Assume that there is a file 'book.xlsx' in the OneDrive root folder
var workbookUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/book.xlsx:/workbook";
var session = new Excel.Session(workbookUrl, {Authorization: "Bearer " + accessToken});
Excel.run(session, (context) => {
    var r = context.workbook.worksheets.getItem('Sheet1').getRange("A1:B2");
    r.values = [["Hello", "World"], [1234, "=B2 + 100"]];
    r.load();
    return context.sync()
        .then(() => {
            console.log(JSON.stringify(r.values));
        });
});

```

#### Using JavaScript
```javascript
import * as Excel from '@microsoft/office-js/excel'

var accessToken;
// getAccessToken() method is is implemented by you and it returns OAUTH access token
accessToken = getAccessToken();
// Assume that there is a file 'book.xlsx' in the OneDrive root folder
var workbookUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/book.xlsx:/workbook";
var session = new Excel.Session(workbookUrl, {Authorization: "Bearer " + accessToken});
Excel.run(session, function(context){
    var r = context.workbook.worksheets.getItem('Sheet1').getRange("A1:B2");
    r.values = [["Hello", "World"], [1234, "=B2 + 100"]];
    r.load();
    return context.sync()
        .then(function() {
            console.log(JSON.stringify(r.values));
        });
});

```  
