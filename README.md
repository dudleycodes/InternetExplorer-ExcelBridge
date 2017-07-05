*Cleaning out my hard drives; this project is *not* maintained.*

**This project only works with depreciated technology and was written for a uniquely-limited Microsoft-based environment many years ago.**

# Excel Bridge for Internet Explorer

Provides a real-time link between Internet Explorer and Microsoft Excel via ActiveX.

# Documentation

To generate documentation:

```
npm install .
npm run build-docs
```

# Requirements
* Internet Explorer 8 or higher (ECMAScript 3rd Edition equivalent)
* Microsoft Office 2007 or higher
* Microsoft Windows 7 or higher with appropriate ActiveX permissions configured

# Usage
It is not recommended to create an instance directly - use an appropriate provided factory function.

## Example
```
var ExcelLink = new excelBridgeFromFile("\\server1\docs\abc.xlsx");
ExcelLink.setSelected('Sheet2');
var data = ExcelLink.getRow(1, 0, 'Sheet3');
ExcelLink.setCell('B4', 'Data has been read successfully.', 'Sheet1');
ExcelLink.terminate();
```

## Object Instantiation
To initiate a new instance use one of the following included factory functions. 
 ```
 excelBridgeFromFile(strUNCPath)
 excelBridgeFromNew()
 excelBridgeFromUserOpen()
 ```

 * In the event of  a critical error these factories will return null instead of an object. Creating
 * a new instance directly (w/o using a factory) is not recommended.

## Excel Application Busy
When Excel is busy its activeX connection will become unresponsive and direct calls will halt all javascript execution.  For the most part 'Excel application busy' will occur predictably such as when the user is actively manipulating a document or when a modular dialogue box is displayed. Occasionally 'application busy' will not be so predictable such as during an auto-save or VBA  script is modifying the document.  The function isReady() can be used to determine if Excel is available for read/writes.

To smooth basic retrieval operations around 'application busy' this object uses caching. All caching is stored in containers prefaced with 'cache' which should not be accessed or manipulated from outside of the object instance.

 ## Function Design
 Functions default to using a 'selected' workbook/sheet rather than the 'active' as the 'active' workbook/sheet depends on user selection and can be unpredictable.

## Common Function Returns
| Return Type | Purpose                                                                                                                                                                   |
|-------------|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| null        | Indicates the function cannot be preformed at this time.  Occurs when when the Excel Application is busy and caching is not available.                                    |
| false       | If not expecting a boolean return false indicates the function was successfully executed but cannot be  preformed.  Example: the worksheet to be modified does not exist. |
| varaint     | Function was successfully executed and is returning an appropriate value.                                                                                                 |

Functions that take parameters of 'worksheetName' and 'workbookName' do so in that order. While stark in contrast to Excel's document model tree the vast majority of use-cases involves only one workbook and as such is the more optional parameter.

 ## Error Control
 Internet Explorer 8's jscript implementation has issues catching extended ECMA standard errors. So rather than throwing errors to bubble messages up through the stack this object uses an event log. If a function returns null (be it from application busy or otherwise) or fails the member  variable 'logMessage' will contain a human-friendly string communicating the issue(s) that occurred immediately after a function call.  In the even no issues occurred logMessage will contain an empty string.  Functions that manipulate this variable should not be accessed outside of the object instance.

 ## General Notes
 * Functions with the verbs 'convertTo', 'format', and 'is' will never affect log messaging - all others will.
