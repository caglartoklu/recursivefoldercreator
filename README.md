# RecursiveFolderCreator

RecursiveFolderCreator is a [Microsoft Access](https://products.office.com/en-us/access) application that
creates recursive folders and keep their history.

It is based on [Basic Accessories](https://github.com/caglartoklu/basic-accessories)
which is a framework for Microsoft Access.

This software has been developed for repetitive and similar folder creation tasks.


# Structure

The software consists of 2 files:
* `RecursiveFolderCreator.accdb` The front end. When a new version is released,
it can be replaced unless there is a breaking change.
This file includes the forms, queries and active content
([VBA](https://github.com/OfficeDev/VBA-content) code
and [macro](https://msdn.microsoft.com/en-us/library/office/dn161227.aspx)).
* `RecursiveFolderCreatorDataBackEnd.accdb` The back end. It holds the user data.
It is supposed to change less frequently.

This architecture is called *Split Database Architecture*, making it easier to update the front end,
without overwriting the data.

This software uses an auto execution mechanism to link the front end to the back end in the same folder.
You need to have macros enabled for the software in the *Security Warning* section in this document.

The `src` folder includes the source code of the modules in the front end.
You can export all the code using the following call on immediate window of
[Microsoft Visual Basic for Applications IDE](https://en.wikipedia.org/wiki/Visual_Basic_for_Applications).
```vbnet
Call mdlFiles.ExportAllCode("src")
```
to a `src1 folder under your *Documents* directory.

## Updating

- If you are a new user, download both `.accdb` files since you have no previous data.
- Otherwise, backup your own front end, download the front end only, and overwrite your own copy with it.
Unless stated otherwise, the front end will be compatible with the back end.
Note that if you have some modifications on the front end, you have to migrate them to the new copy manually.


# Usage

- Launch the `RecursiveFolderCreator.accdb`.
- Type the name of the folder to the text box, and click the `Create Folder` button, you will see it being added to the history.
- You can copy/paste to/from the clipboard using `Copy to Clipboard` and `Paste from Clipboard` buttons.
- Double clicking an item from the history will bring it to the text box for modification.
- Clicking the `Open Folder` button will open it in Windows Explorer.
- Selecting a entry in the log and pressing `Delete Selected Log` will delete that log.
- `Delete All Logs` button deletes all the directory logs after a confirmation dialog.

![screenshot1.png](https://raw.githubusercontent.com/caglartoklu/recursivefoldercreator/master/media/screenshot1.png)

## Security Warning

When you first run the application, if you see the following message in a yellow warning bar:

```
SECURITY WARNING Some active content has been disabled. Click for more details. [Enable Content]
```

![security_warning1.png](https://raw.githubusercontent.com/caglartoklu/recursivefoldercreator/master/media/security_warning1.png)

Simply click on the `Enable Content` button, close the file and relaunch it.
This is a valid warning, and you should not run the every file downloaded from Internet on your computer.
When the content is not enabled, macros and VBA code will not work, and the software will be degraded.

If you are in doubt for this project, the source code is available for you to see or modify.
In this case, download the files, but do not click the `Enable Content` button.
Press `ALT` + `F11` to open
[Microsoft Visual Basic for Applications IDE](https://en.wikipedia.org/wiki/Visual_Basic_for_Applications),
to examine the code before allowing it.


# Requirements

[Microsoft Access](https://products.office.com/en-us/access) (tested with 2013 and 2016)
or
[Microsoft Access 2013 Runtime](https://www.microsoft.com/en-us/download/details.aspx?id=39358)
or
[Microsoft Access 2016 Runtime](https://www.microsoft.com/en-us/download/details.aspx?id=50040).


# License

Licensed under the Apache License, Version 2.0.
[LICENSE](https://github.com/caglartoklu/recursivefoldercreator/blob/master/LICENSE) file.


# Legal

All trademarks and registered trademarks are the property of their respective owners.
