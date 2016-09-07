# Automation in finance: programming as a business user in an enterprise environment

Notes from a talk given in 2016 at Sydney's ModelOff Global Training Camp.

Files starting with 'U' are intended to be generic code that can be called upon when necessary.

## High-level Excel

Designing models to facilitate automation.

- Location references ('cell A4:B8', etc) introduce irrelevant information into your formulas that can potentially be a source of errors
- Structured references can be a clearer way to express what you mean, and are potentially more resilient to changes in your spreadsheets
- Named ranges are useful when you know exactly how many cells will be needed, such as a TRUE/FALSE setting, or the four suits in a deck of cards
- Excel tables (known as ListObjects in VBA) are useful when you don't know how much data you'll need to store (zero or more rows)
- Excel tables are also useful for separating inputs and calculations: you can set up a table so that all of its columns are calculated columns, and expand it to the size you need (see the file U_ListObject.bas)
- Be aware that errors can occur, and you should be aware of how to manage them in VBA code (particularly if you intend on fully automating a process).

## Extending VBA

How to talk to other applications.

- You can use VBA to send information between programs, including triggering actions (such as telling Outlook to create an email using information stored in a worksheet)
- This uses the 'Component Object Model' (**COM**), which is basically a way for programs to expose information about how other programs can talk to them (**type libraries**)
- You can explore type libraries using Excel's Object Browser (press F2 in Excel's VBA editor)
- You can also use COM to add extra functions to VBA (such as **regular expressions**, a powerful way to extract subsets of text based on whether it matches a pattern you specify - see U_RegExp.bas).

## Beyond VBA

How to overcome VBA’s limitations.

- Windows' Command Prompt provides access to some simple programs which automate tasks of finding text in files (**FINDSTR**) and comparing two files based on their text contents or binary representation (**FC**)
- You can use Command Prompt by opening it and typing one-off instructions, or alternatively, write those instructions in a file ending in '.bat' so they can be reused (a **batch file**)
- As a more powerful alternative to the Command Prompt, Windows' Scripting Host can be used to code your machine using languages very similar to VBA (**VBScript**) and JavaScript (**JScript**) - you can use them to script Excel and trigger VBA macros, and can use the same COM knowledge that you can use in VBA
- Windows' Task Scheduler can be used to schedule execution of programs or scripts, including batch files and VBScript/Jscript files (using **cscript**).

JScript has a lot of potential:

- it's more powerful than VBA or VBScript (you can use functions as inputs to other functions - see U_GlobalsMgmt.js)
- it can be called from within VBA using ScriptControl (see file example_sort_array_with_JScript.bas)
- JavaScript has a very active and open community
- it can be a basis on which to build other languages (such as TypeScript) - all the language needs is a compiler written in JavaScript
- it can be easily extended easily by downloading scripts as plain text (see the templating and Google Closure examples)
- it is highly resilient (all you really need is a web browser)
- it potentially offers a migration path if you ever want to move an Excel workbook's functions to the web.
