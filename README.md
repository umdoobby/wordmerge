# WordMerge.exe
A programming challenge to make a program that can take in a csv and automatically make changes to a set of template documents based on that information.

#### Note from the author:
Hi, WordMerge was a technical challenge part of an interview. it was written in under 72 hours and most likely as a few bugs.
Feel free to play and use this in your project, just please leave the introduction and credit in the program. As well as a credit and link back to this Github page.
Thanks! :smile:

#### License:
This software is licensed under the GNU GLP-3.0 license.
More information here: https://github.com/umdoobby/wordmerge/blob/master/LICENSE

## About this project: 
 * Written using Visual Studio Community 2019
   * Version 16.4.5
   * VS Build Tools 2017 v15.9.20
 * Written for .Net Core v3.1
 * The main program is written in C#
 * You must include the following references in the project:
   * Microsoft Word 16.0 Object Library
   * Microsoft Office 16.0 Object Library
 * Due to a bug in VS the following two files are manually included:
   * Interop.Microsoft.Office.Core.dll
   * Interop.Microsoft.Office.Word.dll

## Compiling/Building:
 * Set up to work when build to a single executable
 * I like to include the framework in the file but thats just me
 * The platform is simple winx86 though winx64 shouldn't be an issue
 * Prebuilt executables are in the builds folder

 ## Documentation:
 * There is quite a bit of documentation in the code on the code
 * /help gets you the basic usage information
 * /help-all will open the documentation that is included in the project
   * The basic usage is on the introduction page
   * The developer information is on the developer's guide page
   * These are .htm files that it will open in your default browser

## Arguments, Returns, and Usage
#### Usage: 
`WordMerge.exe [/mrgsrc <path to file>] [/templates <path to folder>] [/outfolder <path to folder>] [/errorlog <path to file>] [/delimiter <character>] [/help] [/help-all]`

#### Options:
Argument | Overview | Required?
---------|----------|----------
`/mrgsrc <path to file>`|Specify the source data file for the merge. This argument cannot be left blank.|:heavy_check_mark:
`/templates <path to folder>`|Specify the source folder of all the templates. Templates must be in docx format. If ommitted the program will assume all templates are in the same directory as the program.|:x:
`/outfolder <path to folder>`|Specify the desitnation folder for all processed files. All resulting files will be named "[current time]-[number processed].docx". If ommitted the program will place all processed files into the same directory as the program.|:x:
`/errorlog <path to file>`|Specify the file where any errors that occure during the file processing will be saved. If ommitted the program will create a file named "[current time].errors.log" in the same directory as the program.|:x:
`/delimiter <character>`|Specify the delimiter in the source file. Must be a single character. If ommited the program will assume '\|' is the delimiter. Some characters like '\' and '"' are unavaliable.|:x:
`/help`|Prints out the usage information and short explanation.|:x:
`/help-all`|Opens full documentation.|:x:

#### Retuns:
Value | Definition
------|-----------
-1|Completed but there were warnings
0|Completed with no errors
1|An unknown argument was given
2|Specified file or directory doesn't exist
3|An argument was unrecognized
4|An argument was supplied without the required path to a file or folder or was incomplete
5|The arguments were too ambiguous
6|There was an error trying to create a file or folder
7|No source file was specified
8|There was an error trying to open and read a file
9|The specified file was not parsed correctly
10|The specified folder provided was empty
11|Could not edit the specified file or folder
12|An incorrect path was provided
