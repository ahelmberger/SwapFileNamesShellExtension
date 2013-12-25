# SwapFileNamesShellExtension

A Windows shell extension, written in managed code, that lets you swap the names of two files using the Windows Explorer context menu

## Usage

Use the **two-step mode**:

* Right-click on the first file, then choose *Select file name*
* Right-click on the second file, then choose *Swap file name with '{NameOfFirstFile}'*

or use the **direct mode**:

* Select two file in Windows Explorer
* Rigth-click and choose *Swap file names*

## Installation

* Make sure you have automatic NuGet package retrieval enabled in Visual Studio
* Build the solution
* Copy the output directory into your tools/programs folder
* Execute `install.bat`/`install64.bat` (depending on your system) as Administrator

---

Copyright (c) 2013 Andreas Helmberger, licensed under [The MIT License](http://opensource.org/licenses/MIT).
This library has a dependency on [SharpShell](https://github.com/dwmkerr/sharpshell), which is licensed under [The Code Project Open License (CPOL)](http://www.codeproject.com/info/cpol10.aspx).
