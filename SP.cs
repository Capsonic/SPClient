using Microsoft.SharePoint.Client;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Threading.Tasks;

namespace SPClient
{
    public class SP
    {
        public string SiteURL { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }
        public Web baseWeb { get; set; }
        private ClientContext clientContext;
        public bool ContinueSearching { get; set; }

        public bool Login(string url, string userName, string password)
        {
            clientContext = new ClientContext(url);
            clientContext.Credentials = new NetworkCredential(userName, password, "capsonic");
            /*clientContext.AuthenticationMode = ClientAuthenticationMode.FormsAuthentication;

            clientContext.FormsAuthenticationLoginInfo = new FormsAuthenticationLoginInfo("apacheco", "Alfa0210");*/

            baseWeb = clientContext.Web;
            clientContext.Load(baseWeb, f => f.Folders, f => f.Title);
            clientContext.ExecuteQuery();
            return true;
        }

        Task loadFoldersAndFilesAsync(Folder parentFolder)
        {
            return Task.Factory.StartNew(() =>
            {
                clientContext.Load(parentFolder, f => f.Folders, f => f.Files);
                clientContext.ExecuteQuery();
            });
        }

        private async Task readFiles(Folder parentFolder, IList<FileItem> filesList,
            string folderLike = "", string fileLike = "")
        {
            //Console.WriteLine(parentFolder.ServerRelativeUrl);
            if (ContinueSearching)
            {
                if (filesList == null) throw new Exception("filesList is not initialized.");

                await loadFoldersAndFilesAsync(parentFolder);

                foreach (Folder folder in parentFolder.Folders)
                {
                    await readFiles(folder, filesList, folderLike, fileLike);
                }

                foreach (Microsoft.SharePoint.Client.File file in parentFolder.Files)
                {
                    if (file.Name.Contains(fileLike) || parentFolder.Name.Contains(folderLike))
                    {
                        if (file.Name.Substring(file.Name.Length - 4) == ".xls" || file.Name.Substring(file.Name.Length - 4) == "xlsx"
                            || file.Name.Substring(file.Name.Length - 4) == "xlsm")
                        {
                            filesList.Add(new FileItem()
                            {
                                FileName = file.Name,
                                FolderName = parentFolder.ServerRelativeUrl
                            });

                            Console.WriteLine(" Folder: " + parentFolder.ServerRelativeUrl + " \tFile Name: " + file.Name);
                        }
                    }
                }
            }
        }

        public async Task readWebs(IList<FileItem> FileItemList, string folderLike, string fileLike)
        {
            //FileItemList.Clear();

            ContinueSearching = true;


            foreach (Folder folder in baseWeb.Folders)
            {
                if (ContinueSearching)
                {
                    try
                    {
                        await readFiles(folder, FileItemList, folderLike, fileLike);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex);
                    }
                }
            }

            //clientContext.Load(parentWeb, f => f.Webs);
            //clientContext.ExecuteQuery();

            //foreach (Web web in parentWeb.Webs)
            //{
            //    Console.WriteLine(web.Title);
            //    readWebs(web);
            //}

        }

        public class FileItem
        {
            public FileItem()
            {
                IsSelected = true;
            }
            public string FileName { get; set; }
            public string FolderName { get; set; }
            public string Status { get; set; }
            public bool IsSelected { get; set; }
            public FileItem Clone()
            {
                return new FileItem()
                {
                    FileName = FileName,
                    FolderName = FolderName,
                    IsSelected = IsSelected,
                    Status = Status
                };
            }
        }


        public bool Process(IList<FileItem> fileItemsList, IList<IExcelUpdateProcess> processes)
        {
            foreach (var item in fileItemsList)
            {
                if (item.IsSelected)
                {
                    string fileAddress = item.FolderName + "/" + item.FileName;
                    Folder baseFolder = clientContext.Web.GetFolderByServerRelativeUrl(item.FolderName);
                    Microsoft.SharePoint.Client.File currentFile = clientContext.Web.GetFileByServerRelativeUrl(fileAddress);
                    clientContext.Load(baseFolder);
                    clientContext.Load(currentFile);
                    clientContext.ExecuteQuery();
                    if (currentFile.CheckOutType == CheckOutType.None)
                    {
                        currentFile.CheckOut();
                        clientContext.ExecuteQuery();

                        FileCreationInformation fileUpdated = new FileCreationInformation();
                        FileInformation fileInformation = Microsoft.SharePoint.Client.File.OpenBinaryDirect(clientContext, item.FolderName + "/" + item.FileName);
                        MemoryStream ms = new MemoryStream();
                        fileInformation.Stream.CopyTo(ms);

                        bool quitBecauseOfError = false;
                        using (var p = new ExcelPackage(ms))
                        {
                            foreach (var process in processes)
                            {
                                if (!process.Execute(p))
                                {
                                    item.Status = process.ErrorMessage;
                                    quitBecauseOfError = true;
                                    break;
                                }
                            }
                            if (!quitBecauseOfError)
                            {
                                fileUpdated.Content = p.GetAsByteArray();
                            }
                        }

                        if (quitBecauseOfError)
                        {
                            currentFile.UndoCheckOut();
                            clientContext.ExecuteQuery();
                        }
                        else
                        {
                            fileUpdated.Overwrite = true;
                            fileUpdated.Url = clientContext.Url.Substring(0, clientContext.Url.IndexOf(".com") + 4) + fileAddress;

                            baseFolder.Files.Add(fileUpdated);
                            currentFile.CheckIn("test", CheckinType.MinorCheckIn);
                            clientContext.ExecuteQuery();
                            item.Status = "Processed";
                        }
                    }
                    else
                    {
                        item.Status = "File is Checked Out by: " + currentFile.CheckedOutByUser.LoginName;
                    }
                }
            }
            return true;
        }

        private byte[] OpenFile(Stream fileStream)
        {
            using (var p = new ExcelPackage(fileStream))
            {
                ExcelWorksheet ws = p.Workbook.Worksheets[1];
                ws.Cells[1, 1].Value = "Desde sharepoint client.";
                return p.GetAsByteArray();
            }
        }
    }
}
