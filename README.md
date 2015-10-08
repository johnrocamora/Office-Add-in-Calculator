# Office Add-in: Calculator

An Office task pane add-in that simulates a calculator. You can insert the data from the calculator display into the active selection. You can also get the text from the selection and insert it into the calculator display.

**Table of contents**

* [Summary](#summary)
* [Prerequisites](#prerequisites)
* [Description of the code](#codedescription)
* [Build and debug](#build)
* [Questions and comments](#questions)
* [Additional resources](#additional-resources)

<a name="summary"></a>
## Summary

In this sample, we show you how to use the [JavaScript API for Office](https://msdn.microsoft.com/en-us/library/office/fp142185.aspx) to create a simple calculator in a task pane add-in. 
The sample uses the JavaScript API for Office to interact with an Office document by getting selected text to calculate, or by inserting text into the document from the display. 
The calculator also has a memory set, recall, and clear option.

<a name="prerequisites"></a>
## Prerequisites

This sample requires the following:  

- Visual Studio 2013 with Update 5 or Visual Studio 2015.  
- Microsoft Office 2013
- Internet Explorer 9, which must be installed but doesn't have to be the default browser. 
- One of the following as the default browser: Internet Explorer 9, Safari 5.0.6, Firefox 5, Chrome 13, or a more recent version of one of these browsers.
- Familiarity with JavaScript programming

<a name="codedescription"></a>
## Description of the code

The calculator has a display panel that can be set by using the buttons in the task pane, or by selecting text from the document. 
To get or set the display panel to selected text from the document, the `setSelectedDataFromDisplay` and `getDataFromSelection` methods are provided. The [getSelectedDataAsync](https://msdn.microsoft.com/en-us/library/office/fp142294.aspx) and [setSelectedDataAsync](https://msdn.microsoft.com/EN-US/library/office/fp142145.aspx) methods are called to either get selected text, or set selected text. 
The [Settings](https://msdn.microsoft.com/EN-US/library/office/fp142179.aspx) object is used to store the string to calculate as a custom setting. As the user chooses a number, that number is appended to the custom setting string, and is then calculated once the user chooses `=`.

<a name="build"></a>
## Build and debug

1. Open the OfficeAddinCalculator.sln file in Visual Studio.
2. Press F5 to build and deploy the sample add-in.
3. Type some numbers to calculate in your document, select it, and then click 'Read'. You can also use the calculator's display panel.
4. To write what's in the display panel in the calculator to the active selection in your document, click 'Write'.

<a name="questions"></a>

## Questions and comments

- If you have any trouble running this sample, please [log an issue](https://github.com/OfficeDev/Office-Add-in-Calculator/issues).
- Questions about Office Add-ins development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins). Make sure that your questions or comments are tagged with [office-addins].

<a name="additional-resources"></a>
## Additional resources

- [More Add-in samples](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
- [Office Add-ins](http://msdn.microsoft.com/library/office/jj220060.aspx)
- [Anatomy of an Add-in](https://msdn.microsoft.com/library/office/jj220082.aspx#StartBuildingApps_AnatomyofApp)
- [Creating an Office add-in with Visual Studio](https://msdn.microsoft.com/library/office/fp179827.aspx#Tools_CreatingWithVS)

## Copyright
Copyright (c) 2015 Microsoft. All rights reserved.
