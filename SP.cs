using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
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

        public void Login(string url, string userName, string password)
        {
            clientContext = new ClientContext(url);
            clientContext.Credentials = new NetworkCredential(userName, password, "capsonic");
            /*clientContext.AuthenticationMode = ClientAuthenticationMode.FormsAuthentication;

            clientContext.FormsAuthenticationLoginInfo = new FormsAuthenticationLoginInfo("apacheco", "Alfa0210");*/

            baseWeb = clientContext.Web;
            clientContext.Load(baseWeb, f => f.Folders, f => f.Title);
            clientContext.ExecuteQuery();
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

                foreach (File file in parentFolder.Files)
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
            public string FileName { get; set; }
            public string FolderName { get; set; }
            public string Status { get; set; }
        }
    }
}
