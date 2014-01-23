VSTOContrib
===========

VSTO Contrib is all about making VSTO development better for developers. It allows you to separate your concerns, use IoC and write more testable clean code in your VSTO add-ins.

## Getting Started
A getting started video is available at http://youtu.be/TxRjNsaVX6U, it goes through quite a few of the concepts in VSTO Contrib to help you get started

If you just want to dive in, install `VSTO.<OfficeProduct>` from nuget, for example `VSTO.Outlook` for an outlook package.

Included in the package is a readme with some code snippets to help you get started. Also check out the sample projects in this repo, there is a simple one for each of the Office applications and will help you get started

## Features
### Simplified Model
VSTO Contrib makes it easier to deal with things like:
 - Automativally Registering your custom task panes for new windows
 - Keeping ribbon controls in sync across windows
 - Keeping custom task pane state in sync across windows (size, visibility etc)
 - Allow your code to be contextually aware, for example in work, you know the current window and the document when the user clicks on a ribbon button

VSTO Contrib allows you to create a 'ViewModel' for a particular ribbon type, and you will get a new instance of your view model for each *context* (document, spreadsheet, mailitem etc).

### Convention over configuration
Your Ribbon XML will be discovered based on your ViewModels name, it will automatically be discovered and given to VSTO on demand.

The default convention can be overridden by providing your own `IViewLocationStrategy` 

### More powerful ribbon XML
Ribbon XML give you a lot more power and flexibility than the VSTO Designer, at the expense of lack of context. VSTO Contrib makes sure you have the current window, the context (document, worksheet etc), ribbon all available to you.

### Centralised error handling
Got try/catch's around every method in your add-in? VSTO Contrib allows you to write your own custom `IErrorHandler` which gives you a central place to handle the errors and stop them getting to Office.

### IoC Container Support
VSTO Contrib has full support for IoC containers, there is currently only an Autofac NuGet package, but it is easy to write a simple adapter to any container!

## Samples
There are sample applications for Word, Excel, PowerPoint and Outlook available at https://github.com/JakeGinnivan/VSTOContrib/tree/master/src/Samples, check them out and if there are more scenarios which you would like to see covered raise an issue