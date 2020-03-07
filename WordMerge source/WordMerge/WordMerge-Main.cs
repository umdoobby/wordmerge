// Programming challenge | WordMerge
// Created by Evan Spiker 3/3/20
// REQUIRES THE "Microsoft Word 16.0 Object Library" AND "Microsoft Office 16.0 Object Library" REFERENCES TO BE ADDED TO THE PROJECT

// vital imported resources
using System;
using System.IO;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using System.Diagnostics;

namespace WordMerge
{
    // a convenient place to store some useful functions and to help keep the memory clean
    class Functions
    {
        // print the usage information into the console
        public void printUsageInfo(String[] arguments) {
            // lets get all of this printed out nice and pretty
            String appName = AppDomain.CurrentDomain.FriendlyName;
            String countWhiteSpace = String.Empty;

            // just a quick counter to makes spacing a little more dynamic
            foreach (char i in ("Usage: " + appName + " ").ToCharArray()) {
                countWhiteSpace += " ";
            }

            // lets start writting out
            Console.WriteLine("Usage: " + appName +
                " [/" + arguments[0] + " <path to file>]" +   // MRGSCR
                " [/" + arguments[1] + " <path to folder>]" + //TEMPLATES
                " [/" + arguments[2] + " <path to folder>]"); //OUTFOLDER
            Console.WriteLine(countWhiteSpace + "[/" + arguments[3] + " <path to file>]" + //ERRORLOG
                " [/" + arguments[4] + " <character>]" +  //DELIMITER
                " [/" + arguments[5] + "]" + //SHORTHELP
                " [/" + arguments[6] + "]"); //LONGHELP

            Console.WriteLine("\nOptions:");

            Console.WriteLine("   /" + arguments[0] + " <path to file>       Specify the source data file for the merge. This argument cannot be left blank.");

            Console.WriteLine("   /" + arguments[1] + " <path to folder>  Specify the source folder of all the templates. Templates must be in docx format.");
            Console.WriteLine("                                If ommitted the program will assume all templates are in the same directory as");
            Console.WriteLine("                                the program.");

            Console.WriteLine("   /" + arguments[2] + " <path to folder>  Specify the desitnation folder for all processed files. All resulting files will");
            Console.WriteLine("                                be named \"[current time]-[number processed].docx\". If ommitted the program will");
            Console.WriteLine("                                place all processed files into the same directory as the program.");

            Console.WriteLine("   /" + arguments[3] + " <path to file>     Specify the file where any errors that occure during the file processing will be");
            Console.WriteLine("                                saved. If ommitted the program will create a file named \"[current time].errors.log\"");
            Console.WriteLine("                                in the same directory as the program.");

            Console.WriteLine("   /" + arguments[4] + " <character>       Specify the delimiter in the source file. Must be a single character. If ommited");
            Console.WriteLine("                                the program will assume '|' is the delimiter. Some characters like '\\' and '\"' are");
            Console.WriteLine("                                unavaliable.");

            Console.WriteLine("   /" + arguments[5] + "                        Prints out the usage information and short explanation.");

            Console.WriteLine("   /" + arguments[6] + "                    Opens full documentation.");

            Console.WriteLine("\nPossible exit codes: ");
            Console.WriteLine("   -1 = Completed but there were warnings");
            Console.WriteLine("    0 = Completed with no errors");
            Console.WriteLine("    1 = An unknown argument was given");
            Console.WriteLine("    2 = Specified file or directory doesn't exist");
            Console.WriteLine("    3 = An argument was unrecognized");
            Console.WriteLine("    4 = An argument was supplied without the required path to a file or folder or was incomplete");
            Console.WriteLine("    5 = The arguments were too ambiguous");
            Console.WriteLine("    6 = There was an error trying to create a file or folder");
            Console.WriteLine("    7 = No source file was specified");
            Console.WriteLine("    8 = There was an error trying to open and read a file");
            Console.WriteLine("    9 = The specified file was not parsed correctly");
            Console.WriteLine("   10 = The specified folder provided was empty");
            Console.WriteLine("   11 = Could not edit the specified file or folder");
            Console.WriteLine("   12 = An incorrect path was provided");

            Console.WriteLine("");
        }

        // checks if a file or directory exists, returns true if it does, false if not
        // returns: 1 = is a direcory and does exist
        //          0 = is a file and does exist
        //         -1 = cannot find the path 
        public int touch(String path) {
            if (File.Exists(path)) {
                // This path is a file
                return 0;
            } else if (Directory.Exists(path)) {
                // This path is a directory
                return 1;
            } else {
                // cannot find the file or directory
                return -1;
            }
        }

        // clean up the arguments fed into the program cause windows is bad at that
        public String[] cleanUpArguments(String[] arguments, char[] delimiters){
            // make an array list to make this easier
            List<string> newArgs = new List<string>();

            // lets go through the arguments
            foreach (string i in arguments){
                // set up some temporary variables for cleaning up the strings
                string[] preClean = i.Split(delimiters);
                string finalClean = String.Empty;

                //lets go through the results of split and remove the empty spaces
                foreach (string j in preClean) {
                    finalClean = j.Trim();
                    if (!String.IsNullOrEmpty(finalClean)) {
                        // add the totally scrubbed string to the list
                        newArgs.Add(finalClean);
                    }
                }
            }
            //return the list as an array for efficiency
            return newArgs.ToArray();
        }

        //// read in the the specified file into an array of arrays of strings
        public Array[] readAndParseMergeSrc(String pathToSource, char separator) {
            // make a temporary array list so that i can add to it freely
            List<String[]> parsedData = new List<String[]>();

            // let try to read this file
            try {
                // read in the file
                StreamReader file = new StreamReader(pathToSource);

                // read through every line of the file
                String line;
                int count = 0;
                while ((line = file.ReadLine()) != null) {
                    parsedData.Add(line.Split(separator));
                    count += 1;
                }

                //close the  file
                file.Dispose();

                // if there were no readable lines lets just make sure null returns
                if (count < 1) {
                    return null;
                }
            } catch {
                // unable to read the file
                return null;
            }

            // return the list as an array to save resources
            return parsedData.ToArray();
        }


        // we need to separate out the name of the file and th path that way we can choose the file based on the text file
        public Array[] buildFileGuide(String rootDir, String[] files, String fileExt) {
            List<String[]> fileGuide = new List<String[]>();
            String temp = String.Empty;
            
            // for each of the files lets trim that thing down
            foreach (String i in files) {
                // remove the path section
                temp = i.Replace(rootDir, "").Trim();

                // remove the extension
                temp = temp.Replace(fileExt, "").Trim();

                // add to the array list
                String[] tempArray = {temp, i};
                fileGuide.Add(tempArray);
            }
            // return it as an array so that we can dispose the array list
            return fileGuide.ToArray();
        }
    }

    // a unified place to send any error cases
    class ErrorHandler {

        ////global variables
        //private static String ERRORLOGPATH;

        // constructor that requires arguments
        public ErrorHandler() {
            // required to log errors automatically
            ERRORLOGPATH = null;

            // gets the time when the program was started
            STARTTIME = DateTime.Now;

            // set the errorlog as not open
            ISERRORLOGOPEN = false;
        }

        // constructor where you can set the error log location from the getgo
        public ErrorHandler(String errorLog) {
            // required to log errors automatically
            ERRORLOGPATH = errorLog;

            // gets the time when the program was started
            DateTime startTime = DateTime.Now;

            // set the errorlog as not open, you must manually do that
            ISERRORLOGOPEN = false;
        }

        // exited successfully
        public int exitSuccessfully(int numOfFilesWritten) {
            // grammer is important
            String fileGrammer = "files";
            if (numOfFilesWritten == 1) {
                fileGrammer = "file";
            }

            // print out the number of processed files
            return errorWriter("Successfully exited, wrote out " + numOfFilesWritten.ToString() + " " + fileGrammer + ".", 0);
        }

        // exited successfully but there were errors
        public int exitSuccessfully(int numOfFilesWritten, bool wereThereErrors, String log)
        {
            // grammer is important
            String fileGrammer = "files";
            if (numOfFilesWritten == 1) {
                fileGrammer = "file";
            }

            // get ready to print out the number of processed files as well as a warning message if needed
            String output;
            int returning;
            if (wereThereErrors) {
                output = "Successfully exited with errors, wrote out " + numOfFilesWritten.ToString() + " " + fileGrammer + ".\nPlease review the error log located at \"" + log + "\" for more details.";
                returning = -1;
            } else {
                output = "Successfully exited, wrote out " + numOfFilesWritten.ToString() + " " + fileGrammer + ".";
                returning = 0;
            }

            // write it all out
            return errorWriter(output, returning);
        }

        // exit with errors
        public int exitWithErrors(int exitCode, String error) {
            // report the runtime and return exit code
            DateTime endTime = DateTime.Now;
            TimeSpan diffTime = endTime - STARTTIME;
            String exit = error + "\nExitied with errors, total run time: " + diffTime.TotalMilliseconds.ToString() + "ms";
            return errorWriter(exit, exitCode);
            
        }

        // the error for when the specified directory could not be read
        public int errorReadingPath(String path) {
            return exitWithErrors(2, "ERROR: File or Directory \"" + path + "\" doesnt exist or an error occurs when trying to determine if the specified file or directory exists.");
        }

        // the error for when the wrong kind of path was supplied
        public int errorReadingPath(String path, bool isFile) {
            // see if we are talking about a file or directory, IsFile = true for file; false for directory
            // either way we are reporting the error to the user
            String tempError;
            if (isFile) {
                tempError = "ERROR: The file \"" + path + "\" was provided when a directory was required.";
            } else {
                tempError = "ERROR: The directory \"" + path + "\" was provided when a file was required.";
            }
            return exitWithErrors(12, tempError); ;
        }

        // the error for when the specified argument is not revognized
        public int errorInvaidArgument(String argument) {
            return exitWithErrors(3, "ERROR: The argument \"" + argument + "\" is unrecognized.");
        }

        // the error for when an argument is incomplete, a more generic version of the version below
        public int incompleteArgument(String argument) {
            return exitWithErrors(3, "ERROR: The argument \"" + argument + "\" was either incomplete or invalid options were entered.");
        }

        // the error for when an argument that needs a path was not given a path
        public int incompleteArgument(String argument, bool requireFile) {
            // lets see if the command needs a file or folder so give the user the best error
            String tempError;
            if (requireFile) {
                tempError = "ERROR: The argument \"" + argument + "\" requires a path to a file and none was specified.";
            } else {
                tempError = "ERROR: The argument \"" + argument + "\" requires a path to a folder and none was specified.";
            }
            return exitWithErrors(4, tempError);
        }

        // the error for when an argument is too ambiguous
        public int tooAmbiguous(String argument, String input) {
            // lets tell the user the error and how to avoid it
            String temp1 = "ERROR: The input \"" + input + "\" is too ambiguous and can be confused for another argument while being usd with \"" + argument + "\".";
            String temp2 = "To avoid this in the future please have trailing \"\\\" on directories and specify file extensions for files.";
            return exitWithErrors(5, temp1 + "\n" + temp2);
        }

        // the error for failing when trying to create a file or directory
        public int couldNotCreateFileOrFolder(String path, bool isFile) {
            // see if we are talking about a file or directory, IsFile = true for file; false for directory
            // either way we are reporting the error to the user
            String tempError;
            if (isFile) {
                tempError = "ERROR: There was an error while trying to create file \"" + path + "\".";
            }
            else {
                tempError = "ERROR: There was an error while trying to create directory at \"" + path + "\".";
            }
            return exitWithErrors(6, tempError);
        }

        // the error when the users did not specify a source file
        public int noMergeSourceSpecified() {
            return exitWithErrors(7, "ERROR: No merge source was specified, this argument is required and must specify a valid file.");
        }

        // the error for when the specified file could not be read or was blank
        public int emptyFileOrErrorReading(String path) {
            return exitWithErrors(8, "ERROR: The file \"" + path + "\" is either empty or unreadable please check the file.");
        }

        // the error for when the file was not able to be parsed
        public int errorParsingFile(String path, char delimiter) {
            return exitWithErrors(9, "ERROR: The file \"" + path + "\" could not be parsed correctly with '" + delimiter.ToString() + "' as the delimiter, please check the file."); ;
        }

        // error for when a folder is empty when we really need there to be files in that directory
        public int folderIsEmpty(String path) {
            return exitWithErrors(10, "ERROR: The folder \"" + path + "\" is empty and no data could be loaded.");
        }

        // the error for failing when trying to create a file or directory
        public int couldNotWriteFileOrFolder(String path, bool isFile) {
            // see if we are talking about a file or directory, IsFile = true for file; false for directory
            // either way we are reporting the error to the user
            String tempError;
            if (isFile) {
                tempError = "ERROR: There was an error while trying to edit the file \"" + path + "\".";
            }
            else {
                tempError = "ERROR: There was an error while trying to edit the folder at \"" + path + "\".";
            }
            return exitWithErrors(11, tempError);
        }

        // open the stream for the error log file
        // returns true if it worked out, false if there was an error writing to the file
        public bool openErrorLog() {
            // we need there to be an error log, return false for failing
            if (String.IsNullOrEmpty(ERRORLOGPATH)) {
                return false;
            }

            // lets try to open the stream
            try {
                // is if it exists
                if (File.Exists(ERRORLOGPATH)) {
                    // make a writer and test it with a quick header write
                    StreamWriter sw = File.AppendText(ERRORLOGPATH);
                    sw.WriteLine("========================");
                    sw.WriteLine("New run started at: " + STARTTIME.ToString("MM/dd/yyyy @ HH:mm:ss"));
                    sw.Close();
                    sw.Dispose();
                    ISERRORLOGOPEN = true;
                    return true;
                } else {
                    // the file didnt exist, return false
                    return false;
                }
            } catch {
                // there was an error opening and writting to the file
                return false;
            }
        }

        // opens the error log and sets the variable
        // returns true if it was able to, false if there was an error
        public bool openErrorLog(String logPath) {
            if (!String.IsNullOrEmpty(logPath)) {
                ERRORLOGPATH = logPath;
                return openErrorLog();
            }
            return false;
        }

        // standard place for errors to be output
        public int errorWriter(String error, int exitCode){
            if (ISERRORLOGOPEN) {
                try {
                    StreamWriter sw = File.AppendText(ERRORLOGPATH);
                    sw.WriteLine("");
                    sw.WriteLine(error);
                    sw.WriteLine("Exiting with code: " + exitCode);
                    sw.Close();
                    sw.Dispose();
                } catch {
                    // there was an errortyring to write to the error log
                    Console.WriteLine("ERROR: Could not write errors to the log file.");
                }
            }
            Console.WriteLine(error);
            return exitCode;
        }

        // this is just for logging errors that we can recover from, warnings, or information to the error log
        public bool warnWriter(String error) {
            // make sure we have a valid error log
            if (ISERRORLOGOPEN) {
                try {
                    // write the error|info|w/e to the file
                    StreamWriter sw = File.AppendText(ERRORLOGPATH);
                    sw.WriteLine("");
                    sw.WriteLine(error);
                    sw.Close();
                    sw.Dispose();
                    return true;
                } catch {
                    // there was an errortyring to write to the error log
                    Console.WriteLine("ERROR: Could not write errors to the log file.");
                }
            }
            // no file open cant write it
            return false;
        }

        public DateTime STARTTIME { get; }

        public String ERRORLOGPATH { get; set; }

        private bool ISERRORLOGOPEN { get; set; }

    }

    // here is the actuall application
    class Program
    {
        // start the program and capture the arguments
        // return values are the exit codes from the error handler
        // Return values:-1 = completed but there were errors 
        //                0 = completed without errors
        //                1 = unknown argument was passed into the application
        //                2 = file or directory doesnt exist
        //                3 = an argument was not recongnized
        //                4 = an argument was supplied without the required path to a file or folder or was incomplete
        //                5 = the arguments were too ambiguous
        //                6 = there was an error trying to create a file or folder
        //                7 = no source file was specified
        //                8 = there was an error trying to open and read a file
        //                9 = The file was not parsed correctly
        //               10 = the folder provided was empty
        //               11 = could not edit file or folder
        //               12 = incorrect path was provided

        // define the expected arguments as constants, if you want to change the arguments do it here
        private const string MRGSCR = "mrgsrc", TEMPLATES = "templates", OUTFOLDER = "outfolder", ERRORLOG = "errorlog", DELIMITER = "delimiter", SHORTHELP = "help", LONGHELP = "help-all";

        static int Main(string[] args)
        {
            // Intro and credit, please leave this in the program!
            Console.WriteLine("WordMerge v1.0");
            Console.WriteLine("(C) 2020 Evan Spiker [umdoobby]");
            Console.WriteLine("More info avaliable at https://github.com/umdoobby/wordmerge/.\n");

            // set up some of the needed variables for this
            Functions func = new Functions();      // what does the bulk of the work
            ErrorHandler eh = new ErrorHandler();  // set up for the easy to find error statments
            String mergeSource = String.Empty;     // the location of the source information to inject
            String templatesFolder = String.Empty; // where I can look to expect the 
            String outfileLocation = String.Empty; // where to drop the output
            String errorFile = String.Empty;       // where to log any errors
            char mergeSrcDelim = '|';              // create the delimiter variable, we just start out assuming '|'
            char startingDocDelim = '{';           // the first delimiter for the variables in the document
            char endingDocDelim = '}';             // the second delimiter for the variables in the document

            // define the delimiter between the directories, in this case '\' for windows anyway
            char dirDelimiter = '\\';
            // get the current working directory
            String currentWorkingDir = Directory.GetCurrentDirectory() + dirDelimiter;
            // get the root directory of the system
            String rootDir = Directory.GetDirectoryRoot(currentWorkingDir);
            // define the delimiter for a directory path, a directory should always have ":\" in it so thats what we are using
            String dirStartingDelim = ":\\";
            // get actual directory
            String actWorkingDir = System.AppContext.BaseDirectory;

            // lets define the current date/time string for when we start naming things
            String currentDateTime = eh.STARTTIME.ToString("MM.dd.yyyy-HH.mm.ss");

            // lets set the expected document type here
            String docExt = ".docx";

            // lets make an array of the argument options for easier logic - these are in a particular order
            String[] validArguments = { MRGSCR, TEMPLATES, OUTFOLDER, ERRORLOG, DELIMITER, SHORTHELP, LONGHELP };

            // see if we were pass any arguments if not just print the most basic help for usage
            if (args.Length == 0) {
                Console.WriteLine("For usage please use /help and for full documentation please use /help-all.");
                return eh.exitSuccessfully(0);
            }

            // lets clean up the inputs cause windows is bad at doing that automatically
            char[] argDelimiters = { '"', ' ', '/', '?' };  // hard code the delimiters
            String[] realArgs = func.cleanUpArguments(args, argDelimiters);

            // now we know all locations and arguments passed into now lets starts to take this apart and do work
            for (int i = 0; i < realArgs.Length; i++) {
                // grab only the arguments
                if (i%2 == 0) {
                    // normalize the arguments to make testing easier
                    realArgs[i] = realArgs[i].ToLower();

                    // test each of the arguments
                    switch (realArgs[i]) {
                        case MRGSCR:
                            // lets set up the path for the file that we are merging data from
                            // the next position in args is suppose to be the path to the directory but we will verify that

                            // just in case there is nothing after this argument we need to handle that 
                            if (!(i + 1 < realArgs.Length)) {
                                return eh.incompleteArgument(realArgs[i], true);
                            }

                            // we also need to make sure the next argument isnt just another argument instead of a folder
                            // unfortunately how im going to do this means that certain keywords will not be allowed as relavtive folders or .any files
                            foreach (String j in validArguments) {
                                if (j == realArgs[i + 1]) {
                                    // if we are here then the argument after this argument is too ambiguous and we are going stop and as for clarification
                                    return eh.tooAmbiguous(realArgs[i], realArgs[i + 1]);
                                }
                            }

                            // now lets see if the user gave me a global or relative file
                            if (!realArgs[i + 1].Contains(dirStartingDelim)) {
                                // if we are here then we were not given a full directory so lets continue assuming that this is suppose to be relative
                                realArgs[i + 1] = currentWorkingDir + realArgs[i + 1];
                            }

                            switch (func.touch(realArgs[i + 1])) {
                                case -1:
                                    // cannot read the file or directory
                                    return eh.errorReadingPath(realArgs[i + 1]);
                                case 0:
                                    // the file exists thats what we need
                                    mergeSource = realArgs[i + 1];
                                    break;
                                case 1:
                                    // this is a directory, we cant use that here
                                    return eh.errorReadingPath(realArgs[i + 1], false);
                            }
                            break;

                        case TEMPLATES:
                            // lets set up the path variable to the location of the templates
                            // the next position in args is suppose to be the path to the directory but we will verify that

                            // just in case there is nothing after this argument we need to handle that 
                            if (!(i + 1 < realArgs.Length)) {
                                return eh.incompleteArgument(realArgs[i], false);
                            }

                            // we also need to make sure the next argument isnt just another argument instead of a folder
                            // unfortunately how im going to do this means that certain keywords will not be allowed as relavtive folders or .any files
                            foreach (String j in validArguments)
                            {
                                if (j == realArgs[i + 1])
                                {
                                    // if we are here then the argument after this argument is too ambiguous and we are going stop and as for clarification
                                    return eh.tooAmbiguous(realArgs[i], realArgs[i + 1]);
                                }
                            }

                            // now lets see if the user gave me a global or relative directory
                            if (!realArgs[i + 1].Contains(dirStartingDelim)) {
                                // if we are here then we were not given a full directory so lets continue assuming that this is suppose to be relative
                                realArgs[i + 1] = currentWorkingDir + realArgs[i + 1];
                            }

                            // we need this to be a directory so we are going to make sure this read as a directory
                            if (!realArgs[i + 1].EndsWith(dirDelimiter)) {
                                // if we are here then we know that there isnt a trailing '\' and windows needs that
                                realArgs[i + 1] += "\\";
                            }


                            switch (func.touch(realArgs[i + 1])) {
                                case -1:
                                    // cannot read the file or directory
                                    return eh.errorReadingPath(realArgs[i + 1]);
                                case 0:
                                    // this is a file, we cant use that here
                                    return eh.errorReadingPath(realArgs[i + 1], true);
                                case 1:
                                    // the directory exists, thats what we need
                                    templatesFolder = realArgs[i + 1];
                                    break;
                            }
                            break;

                        case OUTFOLDER:
                            // setup the  output folder variable
                            // the next position in args is suppose to be the path to the directory but we will verify that

                            // just in case there is nothing after this argument we need to handle that 
                            if (!(i + 1 < realArgs.Length))
                            {
                                return eh.incompleteArgument(realArgs[i], false);
                            }

                            // we also need to make sure the next argument isnt just another argument instead of a folder
                            // unfortunately how im going to do this means that certain keywords will not be allowed as relavtive folders or .any files
                            foreach (String j in validArguments)
                            {
                                if (j == realArgs[i + 1])
                                {
                                    // if we are here then the argument after this argument is too ambiguous and we are going stop and as for clarification
                                    return eh.tooAmbiguous(realArgs[i], realArgs[i + 1]);
                                }
                            }

                            // now lets see if the user gave me a global or relative directory
                            if (!realArgs[i + 1].Contains(dirStartingDelim))
                            {
                                // if we are here then we were not given a full directory so lets continue assuming that this is suppose to be relative
                                realArgs[i + 1] = currentWorkingDir + realArgs[i + 1];
                            }

                            // we need this to be a directory so we are going to make sure this read as a directory
                            if (!realArgs[i + 1].EndsWith(dirDelimiter)) {
                                // if we are here then we know that there isnt a trailing '\' and windows needs that
                                realArgs[i + 1] += "\\";
                            }

                            switch (func.touch(realArgs[i + 1])) {
                                case -1:
                                    // cannot read the file or directory
                                    return eh.errorReadingPath(realArgs[i + 1]);
                                case 0:
                                    // this is a file, we cant use that here
                                    return eh.errorReadingPath(realArgs[i + 1], true);
                                case 1:
                                    // the directory exists, thats what we need
                                    outfileLocation = realArgs[i + 1];
                                    break;
                            }
                            break;

                        case ERRORLOG:
                            // setup the errorlog variable, this one is a bit special cause they could be refereing to a file that we need to make

                            // just in case there is nothing after this argument we need to handle that 
                            if (!(i + 1 < realArgs.Length))
                            {
                                return eh.incompleteArgument(realArgs[i], true);
                            }

                            // we also need to make sure the next argument isnt just another argument instead of a folder
                            // unfortunately how im going to do this means that certain keywords will not be allowed as relavtive folders or .any files
                            foreach (String j in validArguments)
                            {
                                if (j == realArgs[i + 1])
                                {
                                    // if we are here then the argument after this argument is too ambiguous and we are going stop and as for clarification
                                    return eh.tooAmbiguous(realArgs[i], realArgs[i + 1]);
                                }
                            }

                            // now lets see if the user gave me a global or relative directory
                            if (!realArgs[i + 1].Contains(dirStartingDelim)) {
                                // if we are here then we were not given a full directory so lets continue assuming that this is suppose to be relative
                                realArgs[i + 1] = currentWorkingDir + realArgs[i + 1];
                            }

                            // lets set up the correct variable, and make sure that we are dealing with a file not directory
                            // the next position in args is suppose to be the path to the directory
                            switch (func.touch(realArgs[i + 1])) {
                                case -1:
                                    // cannot read the file or directory
                                    // in this case we will just try to make it
                                    try {
                                        File.Create(realArgs[i + 1]).Dispose();
                                    } catch {
                                        // if here then there was an error making the file
                                        return eh.couldNotCreateFileOrFolder(realArgs[i + 1], true);
                                    }
                                    
                                    // now lets just verify that its here and that we can touch it
                                    if (func.touch(realArgs[i + 1]) == 0) {
                                        errorFile = realArgs[i + 1];
                                        break;
                                    } else {
                                        // if here then there was still an error, at this point lets stop
                                        return eh.errorReadingPath(realArgs[i + 1], false);
                                    }
                                case 0:
                                    // the file exists thats what we need
                                    errorFile = realArgs[i + 1];
                                    break;
                                case 1:
                                    // this is a directory, we cant use that here
                                    return eh.errorReadingPath(realArgs[i + 1], false);
                            }
                            break;

                        case DELIMITER:
                            // setup the delimiter variable, this one is also a bit special as we can only accept 1 character
                            // unfortunately some characters like \ that are being used as delimiters elsewere will not be avaliable

                            // just in case there is nothing after this argument we need to handle that 
                            if (!(i + 1 < realArgs.Length)) {
                                return eh.incompleteArgument(realArgs[i]);
                            }

                            // we also need to make sure the next argument isnt just another argument instead of a folder
                            // unfortunately how im going to do this means that certain keywords will not be allowed as relavtive folders or .any files
                            foreach (String j in validArguments) {
                                if (j == realArgs[i + 1]) {
                                    // if we are here then the argument after this argument is too ambiguous and we are going stop and as for clarification
                                    return eh.tooAmbiguous(realArgs[i], realArgs[i + 1]);
                                }
                            }

                            // lets make sure the user entered just one char
                            if (realArgs[i + 1].Length > 1) {
                                return eh.incompleteArgument(realArgs[i]);
                            }

                            // set the delimiter variable
                            mergeSrcDelim = realArgs[i + 1].ToCharArray()[0];

                            //finally done with this switch
                            break;

                        case SHORTHELP:
                            // print the usage instructions in the console and end the program
                            func.printUsageInfo(validArguments);
                            return eh.exitSuccessfully(0);

                        case LONGHELP:
                            // open the full documentation
                            Console.WriteLine("INFO: This action will open your web browser.");
                            switch (func.touch(actWorkingDir + "help-info\\intro.htm")) {
                                case 0:
                                    try {
                                        // try to open the web browser to show the help document
                                        var proc = Process.Start(@"cmd.exe ", @"/c " + actWorkingDir + "help-info\\intro.htm");
                                    }
                                    catch {
                                        // unable to open the browser
                                        Console.WriteLine("ERROR: Unable to launch web browser, you should be able to find the help info at \"" + actWorkingDir + "help-info\\intro.htm\".");
                                    }
                                    break;

                                default:
                                    // unable to find the help document 
                                    Console.WriteLine("ERROR: Could not find help document at \"" + actWorkingDir + "help-info\\intro.htm\" please check that there are no missing files and consider reinstalling the program.");
                                    break;
                            }
                            // either way we are just going t exit
                            return eh.exitSuccessfully(0);

                        default:
                            // if we are here then the argument is unknow throw error, end the program
                            return eh.errorInvaidArgument(realArgs[i]);


                    }
                }

            }

            // at this point all arguments should be checked and known valid
            // lets make sure all of the necessary variables are setup, if the user didnt manually specify one we will need to assume a path

            // if no merge source was specified then exit, no point in going further
            mergeSource = mergeSource.Trim();
            if (String.IsNullOrEmpty(mergeSource)) {
                return eh.noMergeSourceSpecified();
            }

            // if no templates folder was specified then we are going to assume that they are going to be in the current directory
            templatesFolder = templatesFolder.Trim();
            if (String.IsNullOrEmpty(templatesFolder)) {
                templatesFolder = currentWorkingDir;
            }

            // if no output folder location is specified then we are going to assume the want all results placed in the current directory
            outfileLocation = outfileLocation.Trim();
            if (String.IsNullOrEmpty(outfileLocation)) {
                outfileLocation = currentWorkingDir;
            }

            // if no error log file was specified then assume that we are making a new file in the current directory
            errorFile = errorFile.Trim();
            if(String.IsNullOrEmpty(errorFile)) {
                try {
                    File.Create(currentWorkingDir + currentDateTime + "-errors.log").Dispose();
                    errorFile = currentWorkingDir + currentDateTime + "-errors.log";
                    Console.WriteLine("Errors will be logged in \"" + errorFile + "\".");
                }
                catch
                {
                    // if here then there was an error making the file
                    return eh.couldNotCreateFileOrFolder(currentWorkingDir + currentDateTime + "-errors.log", true);
                }
            }

            // lets open the errorlog and return an error if we fail
            if (!eh.openErrorLog(errorFile)) {
                return eh.couldNotWriteFileOrFolder(errorFile, true);
            }

            // at this point all variables that need to be defined should be defined so all of the information should be gathered and lets rock'n'roll
            // lets start by trying to read the merge source file into memory
            Array[] parsedData = func.readAndParseMergeSrc(mergeSource, mergeSrcDelim);

            // parsedData is null then an error occured while reading the file, end the program
            if (parsedData == null || parsedData.Length == 0) {
                return eh.emptyFileOrErrorReading(mergeSource);
            }
            
            // lets make sure there is data in the arrays of strings in the array of parsed data
            // the first line sets up the expectations for the entire file, lets make sure that is met
            foreach (Array i in parsedData) {
                if (i.Length != parsedData[0].Length ) {
                    // if we are here then we got a blank line and thats not allowed
                    return eh.errorParsingFile(mergeSource, mergeSrcDelim);
                }
            }

            // lets try to read in the templates from the template  folder
            String[] templateFiles;
            try {
                templateFiles = Directory.GetFiles(templatesFolder, "*" + docExt);
            } catch {
                // if we are here then we couldnt read the dir for some reason
                return eh.folderIsEmpty(templatesFolder);
            }

            // now see if we got any files, if not we need to end the program
            if (templateFiles.Length < 1) {
                return eh.folderIsEmpty(templatesFolder);
            }

            //// now lets build the guide we are going to use to the logic
            //Array[] templateGuide = func.buildFileGuide(templatesFolder, templateFiles, docExt);

            // ===========================
            // start replacing
            //============================

            // lets start counting the files we are outputting
            // we also need a value to reference the template name column
            int fileCount = 0;
            int templateCol = -1;

            // open the word application and hide it
            var wordApp = new Microsoft.Office.Interop.Word.Application();
            wordApp.Visible = false;
            var wDoc = new Microsoft.Office.Interop.Word.Document();

            // lets find the template col and build the titles table, we are going to assume the first line is always the definition
            int tempTemplateFinder = 0;
            List<String> mergeTitles = new List<string>();
            foreach (String i in parsedData[0]) {
                if (i.ToLower() == "template") {
                    templateCol = tempTemplateFinder;
                }
                mergeTitles.Add(i.Trim());
                tempTemplateFinder += 1;
            }

            // just a boolean for if there were any warning during replacement
            bool warnings = false;

            // main replacement loop
            List<String> workingLine = new List<string>();
            // go through the merge data line by line
            for (int i = 1; i < parsedData.Length; i++) {
                // open the first line and throw it in an array list to keep my sanity
                foreach(String j in parsedData[i]) {
                    workingLine.Add(j);
                }

                bool found = false;

                // see if we can find the template
                switch (func.touch(templatesFolder + workingLine[templateCol] + docExt)) {
                    case 0:  // found the document
                        try {
                            wDoc = wordApp.Documents.Open(templatesFolder + workingLine[templateCol] + docExt);
                            wDoc.Activate();
                            found = true;
                        } catch {
                            eh.warnWriter("WARNING: Could not find the template \"" + templatesFolder + workingLine[templateCol] + docExt + "\" therefore line " + i.ToString() + " was skipped.");
                            warnings = true;
                        }
                        break;

                    default: // there was an error finding the document
                        eh.warnWriter("WARNING: Could not find the template \"" + templatesFolder + workingLine[templateCol] + docExt + "\" therefore line " + i.ToString() + " was skipped.");
                        warnings = true;
                        break;
                }

                // we only do the processing if we found the template previously
                if (found) {
                    // lets tally the spelling errors incase we make one
                    int numOfErrorsPreEdit = wDoc.SpellingErrors.Count;

                    // the length of the working line must be the same as the titles we got previously
                    if (workingLine.Count == mergeTitles.Count)
                    {
                        // go through the entire working line
                        for (int j = 0; j < workingLine.Count; j++)
                        {
                            // we arent looking for the template name in the document
                            if (j != templateCol)
                            {
                                // lets find the fields to replace
                                if (wDoc.Content.Find.Execute(FindText: startingDocDelim + mergeTitles[j] + endingDocDelim))
                                {
                                    // we have an issue, erronious white spaces can be made... lets at least try to avoid this
                                    if (workingLine[j].Trim().Length > 0) {
                                        wDoc.Content.Find.Execute(FindText: startingDocDelim + mergeTitles[j] + endingDocDelim, ReplaceWith: workingLine[j].Trim(), Replace: WdReplace.wdReplaceAll);
                                    } else {
                                        // then there is whitespace here and there could be an extra space or two!
                                        if (wDoc.Content.Find.Execute(FindText: " " + startingDocDelim + mergeTitles[j] + endingDocDelim + " ")) {
                                            // we could have two extra spaces at this location
                                            wDoc.Content.Find.Execute(FindText: " " + startingDocDelim + mergeTitles[j] + endingDocDelim + " ", ReplaceWith: workingLine[j].Trim(), Replace: WdReplace.wdReplaceAll);
                                            // but if we replaced something in the middle of a sentense we would have made a spelling error, lets try to fix it
                                            if (numOfErrorsPreEdit != wDoc.SpellingErrors.Count) {
                                                // we created one lets undo, and replace again leaving a space, and just add a space and log the issue
                                                eh.warnWriter("WARNING: Possible Gramatical or Spelling error was created in output document \"" + outfileLocation + currentDateTime + " - " + i.ToString() + docExt + "\" you may want to review it.");
                                                Console.WriteLine("WARNING: Possible Gramatical or Spelling error was created in output document \"" + outfileLocation + currentDateTime + " - " + i.ToString() + docExt + "\" you may want to review it.");
                                                warnings = true;
                                                wDoc.Undo();
                                                wDoc.Content.Find.Execute(FindText: " " + startingDocDelim + mergeTitles[j] + endingDocDelim + " ", ReplaceWith: workingLine[j].Trim() + " ", Replace: WdReplace.wdReplaceAll);
                                            }
                                        } else if (wDoc.Content.Find.Execute(FindText: " " + startingDocDelim + mergeTitles[j] + endingDocDelim)) {
                                            // we have a leading space issue
                                            wDoc.Content.Find.Execute(FindText: " " + startingDocDelim + mergeTitles[j] + endingDocDelim, ReplaceWith: workingLine[j].Trim(), Replace: WdReplace.wdReplaceAll);
                                        } else if (wDoc.Content.Find.Execute(FindText: startingDocDelim + mergeTitles[j] + endingDocDelim + " ")) {
                                            // we have a trailing space issue
                                            wDoc.Content.Find.Execute(FindText: startingDocDelim + mergeTitles[j] + endingDocDelim + " ", ReplaceWith: workingLine[j].Trim(), Replace: WdReplace.wdReplaceAll);
                                        } else {
                                            // this should never happend but just in case it does lets log it as a parsing issue
                                            eh.warnWriter("WARNING: There was an error parsing the data for output document \"" + outfileLocation + currentDateTime + " - " + i.ToString() + docExt + "\" for line " + i.ToString() + "you may want to review it.");
                                            Console.WriteLine("WARNING: There was an error parsing the data for output document \"" + outfileLocation + currentDateTime + " - " + i.ToString() + docExt + "\" for line " + i.ToString() + "you may want to review it.");
                                            warnings = true;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else {
                        // print that there was a parsing error on this line
                        eh.warnWriter("WARNING: There was an error parsing the data for output document \"" + outfileLocation + currentDateTime + " - " + i.ToString() + docExt + "\" for line " + i.ToString() + "you may want to review it.");
                        Console.WriteLine("WARNING: There was an error parsing the data for output document \"" + outfileLocation + currentDateTime + " - " + i.ToString() + docExt + "\" for line " + i.ToString() + "you may want to review it.");
                        warnings = true;
                    }

                    // we have checked for every string on this line lets empty the working line and start over
                    workingLine.Clear();

                    // lets try to output our hard work
                    try {
                        wDoc.SaveAs2(outfileLocation + currentDateTime + "-" + i.ToString() + docExt);
                        fileCount += 1;
                    } catch {
                        // there was an error trying to save the document
                        eh.warnWriter("ERROR: There was an issue creating the output document \"" + outfileLocation + currentDateTime + "-" + i.ToString() + docExt + "\" for line " + i.ToString() + " please review your perameters and try again.");
                        Console.WriteLine("ERROR: There was an issue creating the output document \"" + outfileLocation + currentDateTime + "-" + i.ToString() + docExt + "\" for line " + i.ToString() + " please review your perameters and try again.");
                        warnings = true;
                    }

                    // close the document
                    wDoc.Close();
                }
            }

            // ===================
            // work is done!
            // ===================

            // kill word and exit
            wordApp.Quit();
            Marshal.ReleaseComObject(wDoc);
            Marshal.ReleaseComObject(wordApp);
            return eh.exitSuccessfully(fileCount, warnings, errorFile);
        }
    }
}
