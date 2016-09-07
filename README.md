# GTC_2016_Sydney
Notes from a talk given in 2016 at Sydney's ModelOff Global Training Camp.

Topics:

## High-level Excel

Designing models to facilitate automation.

- Location references ('cell A4:B8', etc) introduce irrelevant information into your formulas that can potentially be a source of errors
- Structured references can be a clearer way to express what you mean, and are potentially more resilient to changes in your spreadsheets
- Named ranges are useful when you know exactly how many cells will be needed, such as a TRUE/FALSE setting, or the four suits in a deck of cards
- Excel tables (known as ListObjects in VBA) are useful when you don't know how much data you'll need to store (zero or more rows)
- Excel tables are also useful for separating inputs and calculations: you can set up a table so that all of its columns are calculated columns, and expand it to the size you need (see the file Utils_ListObject.bas).

## Extending VBA

How to talk to other applications.

- You can use VBA to send information between programs, including triggering actions (such as telling Outlook to create an email using information stored in a worksheet)
- This uses the 'Component Object Model' (**COM**), which is basically a way for programs to expose information about how other programs can talk to them (specifically called 'type libraries')
- You can use the same mechanism to add extra functions to VBA (such as **regular expressions**, a powerful way to extract subsets of text based on whether it matches a pattern you specify)
- Be aware that errors can occur, and you should be aware of how to deal with them in code (particularly if you intend on fully automating a process).

## Beyond VBA

How to overcome VBA’s limitations.

- Windows' Command Prompt provides access to some simple programs which automate tasks of finding text in files (**FINDSTR**) and comparing two files based on their text contents or binary representation (**FC**)
- You can type one-off messages at the Command Prompt, or alternatively, write instructions in advance that can be reused (a **batch file**)
- Windows' Task Scheduler can be used to set up a regular job that should be run
- Windows' Scripting Host can be used to code your machine using languages very similar to VBA (**VBScript**) and JavaScript (**JScript**) - you can use them to script Excel and trigger VBA macros, and can use the same COM knowledge that you can use in VBA.

JScript has a lot of potential:

- it's a more powerful language than VBA
- has a very active and open community
- can be a basis on which to build other languages that you can also run (such as TypeScript)
- can be easily extended easily by downloading scripts as plain text (see the templating and Google Closure examples)
- is highly resilient (all you really need is a web browser)
- potentially offers a migration path if you ever want to move an Excel workbook's functions to the web.
