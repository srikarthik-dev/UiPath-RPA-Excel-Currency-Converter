<img width="659" height="469" alt="image" src="https://github.com/user-attachments/assets/91d8d50b-8d80-43f0-b3a3-0fcc7d6a8472" />
<img width="1709" height="923" alt="image" src="https://github.com/user-attachments/assets/d1f4ac44-f71d-40b8-a795-764197e4ef50" />
# UiPath RPA: Excel INR to USD Currency Converter

An automated Robotic Process Automation (RPA) workflow developed with **UiPath Studio Web** on macOS. This workflow reads tabular data from an Excel workbook, extracts phone number values as INR figures, converts them to USD via Google Search using in-browser JavaScript injection, and writes the converted values into a dedicated worksheet (`Sheet2`).

## Features

- **Headless In-Memory Conversion**: Utilizes `Inject Js Script` to grab live Google exchange rates and compute conversions directly in the browser DOM, eliminating slow UI loops (`Type Into` / `Get Text`).
- **LINQ Batch Extraction**: Extracts the entire target column using a single LINQ expression (`dt_Input.AsEnumerable().Select(...)`).
- **Clean Excel Multi-Sheet Handling**: Appends converted data directly into `Sheet2` with structured headers (`USD Price`) ready for Apple Numbers and Microsoft Excel.

## Workflow Architecture

```text
[Read Range Workbook] (Reads challenge.xlsx: Sheet1 -> dt_Input)
       │
[Assign] (LINQ: Combines "Phone Number" column into comma-separated string)
       │
[Use Application/Browser: Chrome] (URL: [https://www.google.com/search?q=1+inr+to+usd](https://www.google.com/search?q=1+inr+to+usd))
       │
[Inject Js Script] (Passes numbers string -> evaluates rate -> returns CSV format)
       │
[Generate Data Table From Text] (Converts CSV string -> dt_USDOutput)
       │
[Write Range Workbook] (Writes dt_USDOutput -> challenge.xlsx: Sheet2)function(element, input) {
    var rateEl = document.querySelector('span[data-value]');
    var rate = rateEl ? parseFloat(rateEl.getAttribute('data-value')) : 0.012;
    var numbers = input.split(',');
    var result = ['USD Price'];
    for (var i = 0; i < numbers.length; i++) {
        var val = parseFloat(numbers[i]);
        var usd = isNaN(val) ? '' : (val * rate).toFixed(2);
        result.push('"' + '$' + usd + '"');
    }
    return result.join('\n');
}Prerequisites & Setup
macOS with UiPath Assistant / Remote Agent installed and running.

Google Chrome with the UiPath Web Automation extension enabled.

Target Excel file saved at /Users/srikarthikk/Documents/challenge.xlsx.

Running the Automation
Open UiPath Studio Web.

Open the project and ensure Chrome is accessible.

Click Debug on local machine.

Once completed, view results in Apple Numbers:open -a Numbers /Users/srikarthikk/Documents/challenge.xlsx
