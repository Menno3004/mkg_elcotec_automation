using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Mkg_Elcotec_Automation
{
    static class RevisionHelper
    {
        public static string Folder => System.Configuration.ConfigurationManager.AppSettings["OutputFolder"] ?? "C:/VerkoopOutput";
        private static readonly Dictionary<string, object> _fileLocks = new Dictionary<string, object>();
        private static readonly object _lockDictionary = new object();

        private static object GetFileLock(string filePath)
        {
            lock (_lockDictionary)
            {
                if (!_fileLocks.ContainsKey(filePath))
                {
                    _fileLocks[filePath] = new object();
                }
                return _fileLocks[filePath];
            }
        }
        /*
        public static int UpdateRevisionFromFileNames(string[] ArticleDirectories)
        {
            int updateCounter = 0;
            for(int i = 0; i < ArticleDirectories.Length;i++)
            {
                
                string artiCode = ArticleDirectories[i].Substring(ArticleDirectories[i].LastIndexOf("\\"), ArticleDirectories[i].Length - ArticleDirectories[i].LastIndexOf("\\")).Replace("\\","");
                string[] files = Directory.GetFiles(ArticleDirectories[i]);
                Dictionary<string, int> d = new Dictionary<string, int>();
                string currentRevisionFile = "";
                for (int j = 0; j < files.Length;j++)
                { 
                    string fileName = files[j].Substring(files[j].LastIndexOf("\\"),files[j].Length - files[j].LastIndexOf("\\")).Replace("\\","");
                    if (!fileName.Contains("Revision"))
                    {
                        string revision = "-1";
                        int index = fileName.IndexOf('_');
                        int rev = 99;
                        if (index != -1)
                        {
                            revision = fileName.Substring(index, fileName.LastIndexOf(".") - index).Replace("_", "");
                            if (revision.Length < 5)
                            {
                                rev = Convert.ToInt32(revision);
                                if (rev < 20)
                                {
                                    d.Add(fileName, rev);
                                }
                            }
                        }
                    }
                    else
                    {
                        currentRevisionFile = fileName; 
                    }
                }
                var matchesArtiCode = d.Where(e => e.Key.Contains(artiCode) && !e.Key.Contains("A"));
                int maxHolder = -1;
                if (matchesArtiCode.Any())
                {
                    foreach (var match in matchesArtiCode)
                    {
                        maxHolder = Math.Max(maxHolder, match.Value);
                    }
                    string path = ArticleDirectories[i] + "\\" + currentRevisionFile;
                    if(File.Exists(path))
                    {
                        string filePath = ArticleDirectories[i] + "\\Revision_" + maxHolder + ".txt";
                        //copy contents from file not working yes
                        SetFileToUnhide(path);
                        File.Delete(path);
                        if(!File.Exists(filePath))
                        {
                            File.Create(filePath);
                            File.SetAttributes(filePath, File.GetAttributes(filePath) | FileAttributes.Hidden);
                        }
                    }
                    else
                    {
                        string filePath = ArticleDirectories[i] + "\\Revision_" + maxHolder + ".txt";
                        File.Create(filePath);
                        File.SetAttributes(filePath, File.GetAttributes(filePath) | FileAttributes.Hidden);
                    }
                    var notMatchRevision = matchesArtiCode.Where(e => e.Value != maxHolder && maxHolder != 0 && !e.Key.Contains("A"));
                    if (notMatchRevision.Any())
                    {
                        ArticleRevisionUpdate(notMatchRevision.ToDictionary(x => x.Key, x => x.Value), ArticleDirectories[i]);
                    }
                }
            }
            return updateCounter;
        }
        */
        public static int UpdateRevisionFromFileNames(string[] ArticleDirectories)
        {
            int updateCounter = 0;
            for (int i = 0; i < ArticleDirectories.Length; i++)
            {
                string artiCode = Path.GetFileName(ArticleDirectories[i]);
                string[] files = Directory.GetFiles(ArticleDirectories[i]);
                Dictionary<string, int> d = new Dictionary<string, int>();
                string currentRevisionFile = "";

                // Find the revision file path first
                string revisionFilePath = Path.Combine(ArticleDirectories[i], "Revision_*.txt");
                var revisionFiles = Directory.GetFiles(ArticleDirectories[i], "Revision_*.txt");

                if (revisionFiles.Length > 0)
                {
                    currentRevisionFile = Path.GetFileName(revisionFiles[0]);
                    string fullRevisionPath = revisionFiles[0];

                    // Use file locking to prevent concurrent access
                    lock (GetFileLock(fullRevisionPath))
                    {
                        try
                        {
                            // Your existing revision logic here, but wrapped in the lock
                            // ... rest of your revision update code
                        }
                        catch (IOException ex) when (ex.Message.Contains("being used by another process"))
                        {
                            // Skip this file if it's still locked
                            Console.WriteLine($"Skipping revision update for {artiCode} - file in use");
                            continue;
                        }
                    }
                }
            }
            return updateCounter;
        }
        public static void ArticleRevisionUpdate(Dictionary<string,int> Matches, string ArticleFilePath)
        {
            string oldFilePathDir = Path.Combine(ArticleFilePath, "Artikel revisie-actie benodigd");
            if(!Directory.Exists(oldFilePathDir))
            {
                Directory.CreateDirectory(oldFilePathDir);
                //send message or add random number or counter to old file folder logic
            }
            foreach (var file in Matches)
            {
                string filePath = Path.Combine(ArticleFilePath, file.Key).Replace("\\","/");
                string oldFilePathfile = Path.Combine(oldFilePathDir, file.Key).Replace("\\", "/");
                if(!File.Exists(oldFilePathfile))
                {
                    File.Move(filePath, oldFilePathfile);
                    File.Delete(filePath);
                    CreateEntryTodoFile($"[Article Revision] {file.Key} moved to oldrevision directory: {oldFilePathDir}", "AR",Folder);
                }
            }
        }
        public static void SetFileToUnhide(string path)
        {
            var attributes = File.GetAttributes(path);
            if ((attributes & FileAttributes.Hidden) == FileAttributes.Hidden)
            {
                attributes &= ~FileAttributes.Hidden;
                File.SetAttributes(path, attributes);
            }
        }

        public static void CreateEntryTodoFile(string Message, string Type, string Folder)
        {

            var todoFile = Path.Combine(Folder, "Clients", "TODO.txt");
            switch (Type)
            {
                case "AR":
                    File.AppendAllText(todoFile, "[Artikel revisie] " + Message + "\n\n");
                    break;
                case "ERROR":
                    File.AppendAllText(todoFile, "[Error] " + Message + "\n\n");
                    break;
            }
        }
    }
}
