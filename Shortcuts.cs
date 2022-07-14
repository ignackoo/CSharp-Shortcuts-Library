/******************************************************************************
**  Copyright(c) 2022 ignackoo. All rights reserved.
**
**  Licensed under the MIT license.
**  See LICENSE file in the project root for full license information.
**
**  This file is a part of the C# Library Shortcuts.
** 
**  To Create and Read .url and .lnk shortucts in windows
**
*/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using IWshRuntimeLibrary;


namespace Library
{
    public class Shortcuts
    {
        ///<summary>
        ///Create Url (*.url file).
        ///</summary>
        ///<param name="urlfilepath">Url file(*.url) path without .url extension.</param>
        ///<param name="url">Url address.</param>
        public static void CreateUrl(string urlfilepath, string url)
        {
            using (System.IO.StreamWriter writer = new System.IO.StreamWriter(urlfilepath + ".url"))
            {
                writer.WriteLine("[InternetShortcut]");
                writer.WriteLine("URL=" + url);
                writer.Flush();
                writer.Close();
            }
            return;
        }

        ///<summary>
        ///Get url target path from url (*.url file).
        ///</summary>
        ///<param name="urlfilepath">Url file(*.url) path without .url extension.</param>
        ///<returns>Target url.</returns>
        public static string GetTargetFromUrl(string urlfilepath)
        {
            string url = String.Empty;
            using (System.IO.StreamReader reader = new System.IO.StreamReader(urlfilepath + ".url"))
            {
                int found = 0;
                reader.ReadLine();
                url = reader.ReadLine();
                found = url.IndexOf("=");
                url = url.Substring(found + 1);
                reader.Close();
            }
            return(url);
        }


        ///<summary>
        ///Create link (*.lnk file).
        ///</summary>
        ///<param name="linkfilepath">Link file(*.lnk) path without .lnk extension.</param>
        ///<param name="targetpath">Target path.</param>
        ///<param name="workingdirectorypath">Working directory.</param>
        ///<param name="description">Description comment.</param>
        ///<param name="hotkey">Keyboard hot key.</param>
        public static void CreateLink(string linkfilepath, string targetpath, string workingdirectorypath, string description, string hotkey)
        {
            WshShell shell = new WshShell();
            IWshShortcut link = (IWshShortcut)shell.CreateShortcut(linkfilepath + ".lnk");
            link.TargetPath = targetpath;
            link.Description = description;
            link.WorkingDirectory = workingdirectorypath;
            link.Hotkey = hotkey;
            link.Save();
            return;
        }

        ///<summary>
        ///Get target path from link (*.lnk file).
        ///</summary>
        ///<param name="linkfilepath">Link file(*.lnk) path with .lnk extension to get target path.</param>
        ///<returns>Target path.</returns>
        public static string GetTargetPathFromLink(string linkfilepath)
        {
            if (System.IO.File.Exists(linkfilepath) == false) return (string.Empty);
            WshShell shell = new WshShell();
            IWshShortcut link = (IWshShortcut)shell.CreateShortcut(linkfilepath);
            return (link.TargetPath);
        }
    }
}
