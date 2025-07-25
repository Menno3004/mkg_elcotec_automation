using HtmlAgilityPack;
using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Mkg_Elcotec_Automation
{
    class FileSystemHandler
    {
        public string Folder { get; set; }
        public string OutputFileLogger { get; set; }
        public List<string> AttachementData { get; set; }

        // Updates needed for FileSystemHandler.cs constructor:
        public FileSystemHandler() : this(GetDefaultFolder())
        {
        }

        public FileSystemHandler(string folder)
        {
            Folder = folder;
            EnsureBaseDirectoryExists();
            OutputFileLogger = Folder + "/Log/" + DateTime.Now.ToString("yyyy-mm-dd-hh-mm") + ".txt";
            File.Create(OutputFileLogger).Close();
            AttachementData = LoadAttachmentDataFromFile("Weir");
        }

        private static string GetDefaultFolder()
        {
            return System.Configuration.ConfigurationManager.AppSettings["OutputFolder"] ?? "C:/VerkoopOutput";
        }
        private void EnsureBaseDirectoryExists()
        {
            try
            {
                if (!Directory.Exists(Folder))
                {
                    Directory.CreateDirectory(Folder);
                }
                if (!Directory.Exists(Folder + "/Log"))
                {
                    Directory.CreateDirectory(Folder + "/Log");
                }
                if (!Directory.Exists(Folder + "/Orders/Processed"))
                {
                    Directory.CreateDirectory(Folder + "/Orders/Processed");
                }
                if (!Directory.Exists(Folder + "/Orders/Error"))
                {
                    Directory.CreateDirectory(Folder + "/Orders/Error");
                }
            }
            catch (Exception ex)
            {
                this.ConsoleWriteLineLogger($"\nFout bij het controleren van de basismap: {ex.Message}");
            }
        }
        private List<string> LoadAttachmentDataFromFile(string clientDomain)
        {
            var filePath = Path.Combine(Folder, "Clients", "Weir", "attachments.txt");
            var directoryPath = Path.Combine(Folder, "Clients", "Weir");
            List<string> attachementsInfo = new List<string>();

            if (!Directory.Exists(directoryPath))
            {
                Directory.CreateDirectory(directoryPath);
            }
            if (!File.Exists(filePath))
            {
                File.Create(filePath);
                File.SetAttributes(filePath, File.GetAttributes(filePath) | FileAttributes.Hidden);
            }
            else
            {
                using (var ostrm = new StreamReader(filePath))
                {
                    while (ostrm.Peek() >= 0)
                    {
                        attachementsInfo.Add(ostrm.ReadLine());
                    }
                }
            }
            return attachementsInfo;
        }

        //Helper Methods
        public string TrimFileName(string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
            {
                return "";
            }
            
            // Splits de bestandsnaam en extensie
            string extension = Path.GetExtension(fileName);
            string nameWithoutExtension = Path.GetFileNameWithoutExtension(fileName);

            // Zoek naar een kleurcode zoals A1 t/m A5 en revisienummer
            var match = Regex.Match(nameWithoutExtension, @"^(?<drawing>.+?)_(?<revision>[0-9]{1,3})");
            if (match.Success)
            {
                string drawing = match.Groups["drawing"].Value;
                string revision = match.Groups["revision"].Value;
                return $"{drawing}_{revision}{extension}";
            }

            //Console.WriteLine($"[WARNING] Bestandsnaam voldoet niet aan de verwachte structuur: {fileName}");
            return fileName; // Als geen match, retourneer originele naam
        }
        public string CleanupArticleIdFromColorCodeOrOtherStuff(string articleId)
        {
            if (articleId.Contains("A"))//A Colorcode
            {
                articleId = articleId.Substring(0, articleId.IndexOf('A') - 1);
            }
            if (articleId.Contains("B"))//BT code ?
            {
                articleId = articleId.Substring(0, articleId.IndexOf('B') - 1);
            }
            if (articleId.Contains("V"))//V code ?
            {
                articleId = articleId.Substring(0, articleId.IndexOf('V') - 1);
            }
            if (articleId.Contains("U"))//U code ?
            {
                articleId = articleId.Substring(0, articleId.IndexOf('U') - 1);
            }
            return articleId;
        }
        public void SetFileToUnhide(string path)
        {
            var attributes = File.GetAttributes(path);
            if ((attributes & FileAttributes.Hidden) == FileAttributes.Hidden)
            {
                attributes &= ~FileAttributes.Hidden;
                File.SetAttributes(path, attributes);
            }
        }
        public void SetFileReadOnlyFalse(string path)
        {
            var attributes = File.GetAttributes(path);
            attributes &= ~FileAttributes.ReadOnly;
            File.SetAttributes(path, attributes);
            
        }
        public void SaveErrorDataToFile(string message, string subject)
        {
            try
            {
                // Sanitize the subject for use as filename
                string fileName = SanitizeFileName(subject) + ".txt";
                var filePath = Path.Combine(Folder, "Orders", "Error", fileName);
                File.WriteAllText(filePath, message);
                this.ConsoleWriteLineLogger($"\n error logged in file {filePath}");
                CreateEntryTodoFile($"Error occured in message:{message}", "ERROR");
            }
            catch (Exception ex)
            {
                // Log to event log or console instead of trying to save file again
                Console.WriteLine($"[CRITICAL ERROR] Could not save error file: {ex.Message}");
                this.ConsoleWriteLineLogger($"\nFout bij het opslaan SaveErrorToDataFile {ex.Message}");
            }
        }

        private string SanitizeFileName(string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
                return "UnknownSubject";

            // Remove invalid characters
            char[] invalidChars = Path.GetInvalidFileNameChars();
            string sanitized = fileName;

            foreach (char c in invalidChars)
            {
                sanitized = sanitized.Replace(c, '_');
            }

            // Also replace some additional problematic characters
            sanitized = sanitized.Replace('|', '_')
                                .Replace('\\', '_')
                                .Replace('/', '_')
                                .Replace(':', '_')
                                .Replace('*', '_')
                                .Replace('?', '_')
                                .Replace('"', '_')
                                .Replace('<', '_')
                                .Replace('>', '_');

            // Limit length to prevent path too long errors
            if (sanitized.Length > 100)
                sanitized = sanitized.Substring(0, 100);

            return sanitized;
        }
        public void CreateEntryTodoFile(string message, string type)
        {
            try
            {
                var todoFile = Path.Combine(Folder, "Clients", "TODO.txt");
                switch (type)
                {
                    case "AR":
                        File.AppendAllText(todoFile, "[Artikel revisie] " + message + "\n\n");
                        break;
                    case "ERROR":
                        File.AppendAllText(todoFile, "[Error] " + message + "\n\n");
                        break;
                }
            }
            catch (Exception ex)
            {
                this.ConsoleWriteLineLogger($"\nFout bij het opslaan EntryTodoFile {ex.Message}");
            }
        }

        //Attachements.txt functionality (get data from desctiption field(memo-extern) of order processed with Htmlparser
        public void SaveAttachmentDataToFile(string clientDomain)
        {
            try
            {
                var filePath = Path.Combine(Folder, "Clients", "Weir", "Attachments" + ".txt");
                SetFileToUnhide(filePath);
                List<string> data = AttachementData.Distinct().ToList();
                File.WriteAllText(filePath, string.Empty);
                File.AppendAllLines(filePath, AttachementData);
                File.SetAttributes(filePath, File.GetAttributes(filePath) | FileAttributes.Hidden);
                this.ConsoleWriteLineLogger($"\nAttachement data opgeslagen in Client/Wier/attachements.txt");
            }
            catch (Exception ex)
            {
                //hier gekozen voor geen log to error file, is part of software funcionality, must function, if error program won't function correcttly
                this.ConsoleWriteLineLogger($"\n[CRITICAL ERROR]Fout bij het opslaan van attachmendata {ex.Message}");
            }
        }

        public void AddAttachementEntry(string attachmentName, string clientDomain, string articleId)
        {
            bool duplicate = false;
            try
            {
                string input = attachmentName.Trim() + ":" + articleId.Trim();
                if (attachmentName.Length < 8)
                {
                    return;
                }
                //this should be substrings fix later ....
                if (attachmentName.Contains("Conservation"))
                {
                    return;
                }
                // this should be substrings fix later....
                if (attachmentName.Contains("Conservated"))
                {
                    return;
                }
                // this should be substrings fix later....
                if (attachmentName.Contains("Deliver"))
                {
                    return;
                }
                // this should be substrings fix later....
                if (attachmentName.Contains("and"))
                {
                    return;
                }
                // this should be substrings fix later....
                if (attachmentName.Contains('-'))
                {
                    var index = attachmentName.IndexOf("-");
                    attachmentName = attachmentName.Substring(0, index);
                    input = attachmentName + ":" + articleId;
                }
                foreach (var line in AttachementData)
                {
                    if (line == input) { duplicate = true; }
                }
                if (!duplicate && articleId != "")
                {
                    AttachementData.Add(input.Trim());
                }
                this.ConsoleWriteLineLogger($"\nAttachement code: {attachmentName} toegevoegd voor ArticleId:{articleId}");
            }
            catch (Exception ex)
            {
                this.SaveErrorDataToFile($"Error bij invoeren [Attachment Entry] | attachementName: {attachmentName} | articleId: {articleId}", $"Error bij invoeren Attachment Entry {articleId}");
                this.ConsoleWriteLineLogger($"\nFout bij het opslaan van attachement invoer {ex.Message}");
            }
        }
        internal void RemoveDuplicateEntries(string clientDomain)
        {
            var filePath = Path.Combine(Folder, "Clients", "Weir", "attachments.txt");
            List<string> attachementsInfo = new List<string>();
            using (var ostrm = new StreamReader(filePath))
            {
                while (ostrm.Peek() >= 0)
                {
                    attachementsInfo.Add(ostrm.ReadLine());
                }
            }
            List<string> updatedList = attachementsInfo.Distinct().ToList();
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
                File.AppendAllLines(filePath, attachementsInfo);
            }
            this.ConsoleWriteLineLogger($"\n duplicate entries removed from attachments.txt");
        }
        //Create Folders and revision system logic
        public void CreateDirectoryforArticle(string articleId, string domain, string articleRevision)
        {
            try
            {
                if (articleId.Length > 6)
                {
                    if (articleId.Contains("CR"))
                    {
                        int index = articleId.IndexOf("-");
                        if (index != -1)
                        {
                            string id = articleId.Substring(index, articleId.Length - index);
                            articleId.Trim(' ');
                        }
                    }
                    articleId = CleanupArticleIdFromColorCodeOrOtherStuff(articleId);
                    var domainFilePath = Path.Combine(Folder, "Clients", "Weir", articleId);
                    if (!Directory.Exists(domainFilePath))
                    {
                        this.ConsoleWriteLineLogger($"\nAanmaken van map: {domainFilePath}");
                        Directory.CreateDirectory(domainFilePath);
                        this.ConsoleWriteLineLogger($"\nAanmaken van revisie file:{domainFilePath}/Revision_" + articleRevision + ".txt");
                    }
                }
            }
            catch (Exception ex)
            {
                this.SaveErrorDataToFile($"Error bij maken van [Directory Article] | articleId: {articleId} | articleRevision: {articleRevision}", $"Error bij maken van Directory Article {articleId}");
                this.ConsoleWriteLineLogger($"\nFout bij het aanmaken van de map voor artikel: {ex.Message}");
            }
        }

        //place files in article folder logic
        public void MatchAttachmentsForArticle(AttachmentCollectionResponse attachements, string clientDomain)
        {
            // Zet AttachementData om naar een snelle opzoekstructuur
            var attachmentDataSet = new HashSet<string>(AttachementData, StringComparer.OrdinalIgnoreCase);

            // Verzamel bestanden voor batchverwerking
            var filesToSave = new List<(string filePath, byte[] content)>();

            foreach (var attachment in attachements.Value)
            {
                string trimmedFilename = TrimFileName(attachment.Name);
                string fileHolder = trimmedFilename;
                if (Regex.IsMatch(trimmedFilename, @"^\d{1,5}.\d{1,5}.\d{1,5}.\d{1,5}"))
                {
                    int index = trimmedFilename.IndexOf("_");
                    if (index != -1)
                    {
                        trimmedFilename = trimmedFilename.Substring(0, index);
                    }
                    var strippedArticode = Regex.Replace(trimmedFilename, "[^0-9.]", "");
                    strippedArticode = CleanupArticleIdFromColorCodeOrOtherStuff(strippedArticode);

                    if (Regex.IsMatch(strippedArticode, @"^\d{1,5}\.\d{1,5}\.\d{1,5}") || Regex.IsMatch(strippedArticode, @"^\d{1,5}\.\d{1,5}\.\d{1,5}\.\d{1,5}"))
                    {
                        var matchingLines = attachmentDataSet.Where(line => line.Contains(trimmedFilename));
                        if (matchingLines.Any())
                        {
                            foreach (var line in matchingLines)
                            {
                                string[] splitLine = line.Split(":");
                                if (splitLine.Length < 2) continue;
                                string artiCode = splitLine[1];
                                artiCode = CleanupArticleIdFromColorCodeOrOtherStuff(artiCode);
                                string filePath = Path.Combine(Folder, "Clients", "Weir", artiCode, fileHolder);

                                if (!File.Exists(filePath) && attachment is FileAttachment fileAttachment && fileAttachment.ContentBytes != null)
                                {
                                    filesToSave.Add((filePath, fileAttachment.ContentBytes));
                                }
                             }
                        }
                    }

                }
            }
            foreach (var file in filesToSave)
            {
                File.WriteAllBytes(file.filePath, file.content);
                this.ConsoleWriteLineLogger($"\nBestand opgeslagen: {file.filePath}");          
            }
        }


        public void PlaceHighLevelFilesOnFinalOrderContainingA(AttachmentCollectionResponse attachments)
        {
            var filePathClients = Path.Combine(Folder, "Clients", "Weir");
            foreach (var attachment in attachments.Value)
            {
                string colorCode = "";
                int revisionIndex = attachment.Name.IndexOf("_") - 2;
                if (revisionIndex == -3)
                {
                    revisionIndex = attachment.Name.IndexOf("-") - 2;
                }
                if (revisionIndex > 0)
                {
                    colorCode = attachment.Name.Substring(revisionIndex, 2);
                }
                if (Regex.IsMatch(colorCode, @"^A\d{1}"))
                {
                    string finalFilename = TrimFileName(attachment.Name);
                    int index = finalFilename.IndexOf("_");
                    string artiCode = "";
                    if (index != -1)
                    {
                        artiCode = finalFilename.Substring(0, index - 2);
                    }
                    else
                    {
                        this.ConsoleWriteLineLogger("ERROR - index contains -1 meaning filename logic indexing on  _ not working");
                        continue;
                    }
                    var finalPath = Path.Combine(Folder, "Clients", "Weir", artiCode, finalFilename);
                    var finalFolder = Path.Combine(Folder, "Clients", "Weir", artiCode);
                    if (!Directory.Exists(finalFolder))
                    {
                        if (finalFolder.Contains("CR"))
                        {
                            int lineIndex1 = finalFilename.IndexOf('-') + 2;
                            int lineIndex2 = artiCode.IndexOf('-') + 2;
                            if (lineIndex1 != -1 && lineIndex2 != -1)
                            {
                                string adjustedFilename = finalFilename.Substring(lineIndex1, finalFilename.Length - lineIndex1);
                                string adjustedArticode = artiCode.Substring(lineIndex2, artiCode.Length - lineIndex2);
                                var adjustedFinalPath = Path.Combine(Folder, "Clients", "Weir", adjustedArticode, adjustedFilename);
                                var adjustedfinalFolder = Path.Combine(Folder, "Clients", "Weir", adjustedArticode);
                                if (!Directory.Exists(adjustedfinalFolder))
                                {
                                    Directory.CreateDirectory(adjustedfinalFolder);
                                }
                                if (!File.Exists(adjustedFinalPath))
                                {
                                    SaveAttachementToFileSystem("Weir", attachment, adjustedArticode);
                                    this.ConsoleWriteLineLogger($"file: {adjustedFilename} verplaats naar artikel map: {adjustedFinalPath}");
                                    this.AddAttachementEntry(adjustedArticode, "Weir", adjustedArticode);
                                }
                                return;
                            }
                        }
                        Directory.CreateDirectory(finalFolder);
                    }
                    if (!File.Exists(finalPath))
                    {
                        SaveAttachementToFileSystem("Weir", attachment, artiCode);
                        this.ConsoleWriteLineLogger($"file: {finalFilename} verplaats naar artikel map: {finalFolder}");
                        this.AddAttachementEntry(artiCode, "Weir", artiCode);
                        //this.CreatEntryRevisionFile(files[i], colorCode, "Weir", artiCode); need to add article revision here, might be hard, maybe dont need this method here.
                        //hier kun je geen entry maken 
                    }
                    else
                    {
                        //this.ConsoleWriteLineLogger("High level file al geplaatst in article map, geen actie ondernomen");
                    }
                }

            }
        }

        public void MatchAttachementsonArticleDirectories(AttachmentCollectionResponse attachments)
        {
            var filePathClients = Path.Combine(Folder, "Clients", "Weir");
            string[] directory = Directory.GetDirectories(filePathClients);

            foreach (var attachmement in attachments.Value)
            {
                for (int i = 0; i < directory.Length; i++)
                {
                    string directoryArticleName = directory[i].Substring(directory[i].IndexOf("\\") + 6);
                    if (attachmement.Name.Contains(directoryArticleName))
                    {
                        string trimmedFilename = TrimFileName(attachmement.Name);

                        if (Regex.IsMatch(trimmedFilename, @"^\d{1,5}.\d{1,5}.\d{1,5}.\d{1,5}"))
                        {
                            int index = trimmedFilename.IndexOf("_");
                            if (index != -1)
                            {
                                trimmedFilename = trimmedFilename.Substring(0, index);
                            }
                            var strippedArticode = Regex.Replace(trimmedFilename, "[^0-9.]", "");
                            if (Regex.IsMatch(strippedArticode, @"^\d{1,5}\.\d{1,5}\.\d{1,5}") || Regex.IsMatch(strippedArticode, @"^\d{1,5}\.\d{1,5}\.\d{1,5}\.\d{1,5}"))
                            {
                                
                                //if (strippedArticode.EndsWith('.')) {strippedArticode = strippedArticode.Substring(0, strippedArticode.Length - 1); }
                                strippedArticode = CleanupArticleIdFromColorCodeOrOtherStuff(strippedArticode);
                                var finalDirectory = Path.Combine(Folder, "Clients", "Weir", strippedArticode);
                                if (!Directory.Exists(finalDirectory))
                                {
                                    Directory.CreateDirectory(finalDirectory);
                                }
                                var finalFilePath = Path.Combine(Folder, "Clients", "Weir", strippedArticode, trimmedFilename);
                                if (!File.Exists(finalFilePath))
                                {
                                    SaveAttachementToFileSystem("Weir", attachmement, strippedArticode);
                                    //revision kan hier leeg zijn, word niet op gechecked, laat nu gewoon doorlopen wellicht niet slim of juist wel ........ XD
                                    //this.CreatEntryRevisionFile(files[i], revision, "Weir", directoryName);
                                    this.ConsoleWriteLineLogger($"file: {attachmement.Name} verplaats naar artikel map: {directoryArticleName}");
                                }
                                else
                                {
                                    //this.ConsoleWriteLineLogger("File in rest map staat al in artikel map, geen actie ondernomen");
                                }
                            }
                        }
                    }

                }
            }
        }
        public void SubjectContainArticode(string? subject, AttachmentCollectionResponse attachements)
        {
            if (string.IsNullOrEmpty(subject)) { return; }
            if (subject.StartsWith("CR"))
            {
                int index = subject.IndexOf("_");
                if (index != -1)
                {
                    subject = subject.Substring(4, subject.IndexOf("_") - 4);
                }
            }
            if (subject.StartsWith("RE: CR "))
            {
                int index = subject.IndexOf('-');
                if (index != -1)
                {
                    string dataString = subject;
                    subject = subject.Substring(7, subject.IndexOf('-') - 7);
                    int indexRevision = subject.IndexOf('_');
                    if (indexRevision != -1)
                    {
                        subject = subject.Substring(0, subject.IndexOf("_"));
                    }
                    if (!subject.Contains('.'))
                    {
                        int index2 = dataString.IndexOf('-');
                        if (index2 != -1)
                        {
                            subject = dataString.Substring(dataString.IndexOf('-'));
                        }

                    }
                }
            }

            string[] words = subject.Split(" ");
            int articleCount = MatchArrayforArticleIdCount(words);
            if (articleCount == 0) { return; }
            if (articleCount == 1)
            {
                for (int i = 0; i < words.Length; i++)
                {
                    if (Regex.IsMatch(words[i], @"^\d{1,5}.\d{1,5}.\d{1,5}.\d{1,5}"))
                    {
                        int index3 = words[i].IndexOf("_");
                        if (index3 != -1)
                        {
                            words[i] = words[i].Substring(0, index3);
                        }
                        int index4 = words[i].IndexOf("-");
                        if (index4 != -1)
                        {
                            words[i] = words[i].Substring(0, index4);
                        }
                        var strippedArticode = Regex.Replace(words[i], "[^0-9.]", "");
                        if (Regex.IsMatch(strippedArticode, @"^\d{1,5}\.\d{1,5}\.\d{1,5}") || Regex.IsMatch(strippedArticode, @"^\d{1,5}\.\d{1,5}\.\d{1,5}\.\d{1,5}"))
                        {
                            CreateDirectoryforArticle(strippedArticode, "Weir", "99");
                            foreach (var attachment in attachements.Value)
                            {
                                int dotIndex = attachment.Name.IndexOf('.');
                                if (dotIndex != -1)
                                {
                                    string extension = attachment.Name.Substring(dotIndex, attachment.Name.Length - dotIndex);
                                    extension = extension.ToLower().Trim('.');

                                    if (extension == "gif" || extension == "jpeg" || extension == "jpg" || extension == "png" || extension == "tiff" || extension == "bmp" || extension == "pdf")
                                    {
                                        if (attachment.Size > 35000)
                                        {
                                            SaveAttachementToFileSystem("Weir", attachment, strippedArticode);
                                        }
                                    }
                                    else
                                    {
                                        SaveAttachementToFileSystem("Weir", attachment, strippedArticode);
                                    }
                                }

                            }
                        }
                    }
                }
            }
            else
            {
                for (int i = 0; i < words.Length; i++)
                {
                    if (Regex.IsMatch(words[i], @"^\d{1,5}.\d{1,5}.\d{1,5}.\d{1,5}"))
                    {
                        int index3 = words[i].IndexOf("_");
                        if (index3 != -1)
                        {
                            words[i] = words[i].Substring(0, index3);
                        }
                        int index4 = words[i].IndexOf("-");
                        if (index4 != -1)
                        {
                            words[i] = words[i].Substring(0, index4);
                        }
                        var strippedArticode = Regex.Replace(words[i], "[^0-9.]", "");
                        if (Regex.IsMatch(strippedArticode, @"^\d{1,5}\.\d{1,5}\.\d{1,5}") || Regex.IsMatch(strippedArticode, @"^\d{1,5}\.\d{1,5}\.\d{1,5}\.\d{1,5}"))
                        {
                            CreateDirectoryforArticle(strippedArticode, "Weir", "99");
                            foreach (var attachment in attachements.Value)
                            {
                                int dotIndex = attachment.Name.IndexOf('.');
                                if (dotIndex != -1)
                                {
                                    string extension = attachment.Name.Substring(dotIndex, attachment.Name.Length - dotIndex);
                                    extension = extension.ToLower().Trim('.');

                                    if (extension == "gif" || extension == "jpeg" || extension == "jpg" || extension == "png" || extension == "tiff" || extension == "bmp" || extension == "pdf")
                                    {
                                        if (attachment.Size > 35000)
                                        {
                                            if (attachment.Name.Contains(words[i]))
                                            {
                                                SaveAttachementToFileSystem("Weir", attachment, strippedArticode);
                                            }

                                        }
                                    }
                                    else
                                    {
                                        if (attachment.Name.Contains(words[i]))
                                        {
                                            SaveAttachementToFileSystem("Weir", attachment, strippedArticode);
                                        }
                                    }
                                }

                            }
                        }
                    }
                }
            }
        }
        public int MatchArrayforArticleIdCount(string[] words)
        {
            int articleIdcounter = 0;
            for (int i = 0; i < words.Length; i++)
            {
                var strippedArticode = Regex.Replace(words[i], "[^0-9.]", "");
                if (Regex.IsMatch(strippedArticode, @"^\d{1,5}\.\d{1,5}\.\d{1,5}") || Regex.IsMatch(strippedArticode, @"^\d{1,5}\.\d{1,5}\.\d{1,5}\.\d{1,5}"))
                {
                    articleIdcounter++;
                }
            }
            return articleIdcounter;
        }
        public void SearchForBOM(string htmlBody, string? subject, AttachmentCollectionResponse attachements)
        {
            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(htmlBody);
            string body = "";
            var node = doc.DocumentNode.SelectSingleNode("//body");
            if (node != null)
            {
                body = doc.DocumentNode.SelectSingleNode("//body").InnerText;
            }
            else
            {
                body = htmlBody;
            }
            if (body.Contains("BOM"))
            {
                string[] words = body.Split(" ").Where(x => x != "").ToArray();
                foreach (var word in words)
                {
                    if (Regex.IsMatch(word, @"^\d{1,3}.\d{1,3}.\d{1,3}.\d{1,3}"))
                    {
                        var stripped = Regex.Replace(word, "[^0-9.]", ""); ;
                    }
                }
            }
            if (!string.IsNullOrEmpty(subject))
            {
                if (subject.Contains("BOM"))
                {
                }
                //extract articleID
                //paul bespreken wat doen we hier
            }
            foreach (var attachment in attachements.Value)
            {
                if (attachment.Name.Contains("BOM"))
                {
                    //paul bespreken wat doen we hier
                }
            }
        }
        public void SearchForSTEP(AttachmentCollectionResponse attachements)
        {
            foreach (var attachment in attachements.Value)
            {
                int dotIndex = attachment.Name.IndexOf('.');
                if (dotIndex != -1)
                {
                    string extension = attachment.Name.Substring(dotIndex, attachment.Name.Length - dotIndex);
                    if (extension == ".stp" || extension == ".step")
                    {
                        if (Regex.IsMatch(attachment.Name, @"^\d{1,3}.\d{1,3}.\d{1,3}.\d{1,3}"))
                        {
                            var stripped = Regex.Replace(attachment.Name, "[^0-9.A]", "");
                        }
                    }
                }

            }
        }
        //Save files to file system logic
        public void SaveAttachementToFileSystem(string clientDomain, Attachment attachment, string articleId)
        {
            try
            {
                if (attachment is FileAttachment)
                {
                    FileAttachment att = (FileAttachment)attachment;
                    if(attachment.Name.Contains("891.029.1635"))
                    {
                        //815.011.841 ////wrong filename
                        Console.WriteLine("TEST");
                    }
                    string finalFileName = TrimFileName(attachment.Name);
                    articleId = CleanupArticleIdFromColorCodeOrOtherStuff(articleId);
                    var filePath = Path.Combine(Folder, "Clients", clientDomain, articleId, finalFileName);
                    if (!File.Exists(filePath))
                    {
                        string extension = Path.GetExtension(filePath);
                        File.WriteAllBytes(filePath, att.ContentBytes);
                        this.ConsoleWriteLineLogger($"\nBestandspad voor opslag: {filePath}^%");
                    }
                    else
                    {
                        //this.ConsoleWriteLineLogger($"\nFile al opgeslagen ignore: {filePath}^%");
                    }
                }
                else
                {
                    //ignore file is itemattachent (outlook specific, can't find way to save atm.)
                }

            }
            catch (Exception ex)
            {
                this.SaveErrorDataToFile($"Error bij opslaan van [Save Attachment File System] | attachement (name): {attachment.Name} |  articleId: {articleId}", $"Error bij opslaan van Save Attachment File System {articleId}");
                this.ConsoleWriteLineLogger($"\nFout bij het opslaan van de bijlage in het bestandssysteem: {ex.Message}");
            }
        }
        public void SavePurchsaseOrderToFileSystem(Attachment attachment)
        {
            try
            {
                FileAttachment att = (FileAttachment)attachment;
                var filePath = Path.Combine(Folder, "Clients", "PO", attachment.Name);
                var directoryPath = Path.Combine(Folder, "Clients", "PO");
                if (!Directory.Exists(directoryPath))
                {
                    Directory.CreateDirectory(directoryPath);
                }
                if (!File.Exists(filePath))
                {
                    File.WriteAllBytes(filePath, att.ContentBytes);
                    this.ConsoleWriteLineLogger($"\nPurchase order voor opslag: {filePath}^%");
                }
                else
                {
                    //this.ConsoleWriteLineLogger($"\nPurchase order al opgeslagen ignore: {filePath}^%");
                }
            }
            catch (Exception ex)
            {
                this.SaveErrorDataToFile($"Error bij opslaan van [Purchase Order File System] | attachement (name): {attachment.Name}", $"Error bij opslaan van Save Attachment File System {attachment.Name}");
                this.ConsoleWriteLineLogger($"\nFout bij het opslaan van purchase order: {ex.Message}");
            }
        }

        public void SaveXmlOrder(string xml, string poNumber, string poDate)
        {
            try
            {
                var filePath = Path.Combine(Folder, "Orders");
                Directory.CreateDirectory(filePath);
                var words = poDate.Split("/");
                using (StreamWriter outputFile = new StreamWriter(Path.Combine(filePath, DateTime.Now.ToString(poNumber + "-" + words[2] + words[1] + words[0]) + ".xml")))
                {
                    outputFile.WriteLine(xml);
                    this.ConsoleWriteLineLogger($"Xml Order opgeslagen op: {filePath}");
                }
            }
            catch (Exception ex)
            {
                this.SaveErrorDataToFile($"Error bij opslaan van [XML Order] | xml: {xml} | poNumber {poNumber}", $"Error bij opslaan van Save Attachment File System {poNumber}");
                this.ConsoleWriteLineLogger($"\nFout bij het opslaan van xml order in het bestandssysteem: {ex.Message}");
            }
        }
        //Sumamry logic
        public void PrintoutDircectoryFileDataToFile(string[] articleDirectories)
        {
            string summaryPath = Path.Combine(Folder, "Clients", "WeirData.txt");
            string ClientDataPath = Path.Combine(Folder, "Clients","Weir");
            DirSearchPrintoutToFile(ClientDataPath, summaryPath);
        }
        public void DirSearchPrintoutToFile(string sDir,string summaryPath)
        {
            try
            {
                foreach (string d in Directory.GetDirectories(sDir))
                {
                    string articode = d.Substring(d.LastIndexOf("\\")).Replace("\\",(""));
                    Console.WriteLine(articode);
                    File.AppendAllText(summaryPath, articode + "\r\n");
                    foreach (string f in Directory.GetFiles(d))
                    {
                        string file = f.Substring(f.LastIndexOf("\\") + 1);
                        if(!file.Contains("Revision"))
                        {
                            Console.WriteLine("\t" + file);
                            File.AppendAllText(summaryPath, "\t" + file + "\r\n");
                        }
                        else
                        {
                            string revision = file.Substring(0, file.LastIndexOf("."));
                            Console.WriteLine("\t" + revision);
                            File.AppendAllText(summaryPath, "\t" + revision + "\r\n");
                        }
                    }
                    Console.WriteLine("");
                    File.AppendAllText(summaryPath,"\r\n");
                    DirSearchPrintoutToFile(d,summaryPath);
                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        public void AddDirectoryFileData(string[] articleDirectories)
        {
            for(int i = 0; i < articleDirectories.Length; i++)
            {
                int fileCount = Directory.GetFiles(articleDirectories[i]).Length;
                int dirCount = Directory.GetDirectories(articleDirectories[i]).Length;
                DirectoryInfo dir = new DirectoryInfo(articleDirectories[i]);
                string destination = articleDirectories[i] + $" [D {dirCount} F {fileCount}]";
                dir.MoveTo(destination.Replace("\\","/"));
                dir.Delete(false); 
            }
        }
        
        public void RemoveDirectoryFileData(string[] articleDirectories)   
        {
            for(int i = 0; i < articleDirectories.Length;i++)
            {
                string dirName = Path.GetFileName(articleDirectories[i]);
                string[] words = dirName.Split(" ");
                if(words.Length > 1)
                {
                    string root = articleDirectories[i].Substring(0, articleDirectories[i].Length - articleDirectories[i].LastIndexOf("\\"));
                    DirectoryInfo dir = new DirectoryInfo(articleDirectories[i]);
                    dir.MoveTo(root + "/" + words[0]);
                }
            }
        }
        public void ShowHideFolders(string[] articleDirectories)
        {
            for(int i = 0; i < articleDirectories.Length; i++)
            {
                int fileCount = Directory.GetFiles(articleDirectories[i]).Length;
                DirectoryInfo dir = new DirectoryInfo(articleDirectories[i]);
                FileInfo info = new FileInfo(articleDirectories[i]);
                if (fileCount > 0)
                {
                    dir.Attributes &= FileAttributes.Normal;
                    dir.Attributes &= ~FileAttributes.ReadOnly;
                }
                else
                {
                    dir.Attributes &= FileAttributes.Hidden;
                }
            }
        }
        //Logger logic
        public void ConsoleWriteLineLogger(string line)
        {
            using (var ostrm = new FileStream(OutputFileLogger, FileMode.Append, FileAccess.Write))
            using (var writer = new StreamWriter(ostrm))
            {
                Console.SetOut(writer);
                Console.WriteLine(line);
            }
            var standardOutput = new StreamWriter(Console.OpenStandardOutput());
            standardOutput.AutoFlush = true;
            Console.SetOut(standardOutput);
            Console.WriteLine(line);
        }
    }
}