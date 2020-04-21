using System;
using System.IO;

using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.ApplicationServices;

namespace Tools
{
    // logger class should trace information
    // into different devices
    // - output into the console window
    // - output into a log file
    class Log
    {
        // log file name
        private static string _sFileName = "Autocad.log";

        // setter/getter function
        public static string sFileName
        {
            get { return _sFileName;  }     // ... = Log.sFileName
            set { _sFileName = value; }     // Log.sFileName = ...
        }

        // reset the log information
        public static void Reset()
        {
            File.Delete(_sFileName);
        }

        // append a message to logging information
        public static void Append(string msg)
        {
            // open for append and write log into log file
            StreamWriter f = new StreamWriter(_sFileName, true);
            f.WriteLine(DateTime.Now + "| " + msg);
            f.Close();

            // write msg into AutoCAD command line
            Editor edit = Application.DocumentManager.MdiActiveDocument.Editor;
            edit.WriteMessage(msg + "\n");
        }

    }
}
