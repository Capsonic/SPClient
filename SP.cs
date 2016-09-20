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
                    if (file.Name.Contains(fileLike.Trim()) || parentFolder.Name.Contains(folderLike.Trim()))
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
            public bool IsSelected { get; set; }
            public string FolderName { get; set; }
            public string FileName { get; set; }
            public string Status { get; set; }
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


                        bool quitBecauseOfError = false;


                        FileInformation fileInformation = Microsoft.SharePoint.Client.File.OpenBinaryDirect(clientContext, item.FolderName + "/" + item.FileName);
                        

                        
                        MemoryStream ms = new MemoryStream();
                        fileInformation.Stream.CopyTo(ms);
                        
                        using (var p = new ExcelPackage(ms))
                        {
                            //p.Workbook.VbaProject.Protection.SetPassword(null);
                            //p.Workbook.VbaProject.Remove();
                            //p.Workbook.CreateVBAProject();
                            //var description = p.Workbook.VbaProject.Description;
                            //p.SaveAs(new FileInfo("second.xlsm"));

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
                                //var excelFile = new Microsoft.Office.Interop.Excel.Application();
                                //var workbook = excelFile.Workbooks.Open("tmpFile.xlsm", false, false, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, false, false, Type.Missing, false, true, Type.Missing);
                                //var project = workbook.VBProject;

                                fileUpdated.Content = p.GetAsByteArray();

                                //using (var newDocument = new ExcelPackage(new FileInfo("tmpFile.xlsm")))
                                //{
                                    //for (int i = 1; i < originalDocument.Workbook.Worksheets.Count; i++)
                                    //{
                                    //    originalDocument.Workbook.Worksheets[i].Cells.Copy(newDocument.Workbook.Worksheets[i].Cells); 
                                    //}


                                    //originalDocument.Workbook.Worksheets[1].Cells[1, 1].Value = originalDocument.Workbook.Worksheets[1].Cells[1, 1].Value;
                                    //var description = p.Workbook.VbaProject.Description;
                                    //p.SaveAs(new FileInfo("second.xlsm"));
                                    
                                //}

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
                        item.Status = "Not Processed: File is Checked Out";
                    }
                }
            }
            return true;
        }
    }
}
