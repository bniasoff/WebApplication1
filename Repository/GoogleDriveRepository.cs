
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Application.Interfaces;
using Domain.Models;
using Google.Apis.Drive.v3;
using Google.Apis.Download;

namespace Application.Repository
{

    public class GoogleDriveRepository : IGoogleDriveRepository
    {

        private string[] scopes;
        private const string AppName = "MyAppName";    // Google Drive Application Name (setup on Google Cloud Platform)
                                                       // private const string folderID = "1pgzn6pohEqg_mjBJOjouOh69CB8DMRzs";           // NOTE: Get string from Google Drive Folder ID, https://drive.google.com/drive/folders/FOLDER-ID?usp=sharing
        private DriveService driveService = new DriveService();
        private const string folderID = "1Iu7xp71o8ric-e4Y1VbyFers2j817m1kmJ8gq8ADEXc";           // NOTE: Get string from Google Drive Folder ID, https://drive.google.com/drive/folders/FOLDER-ID?usp=sharing
        private const string FileID = "1Xo6seewjmT-Vk4sJY-vzdaF4CuSqRAME";                                                                                         //private const string folderID = "144Z3Q6Ya2_dDKuYbEKYNQa5-o1991Sk9";

        //private const string folderID = = "1JTrk1jleYAnvJymfheSbVdDXo8xPyV2x";
        // private const string FileID = "1GQ67RWrn_cL6pIcsfCyyaDrPi4mURRFT";
        // private const string FileID = "1Fc9asYIqPafU_JhOHWc-ufrqyKW-zrqL";


        public void Authorize()
        {
            try
            {

                scopes = new string[] {
                DriveService.Scope.Drive,
                DriveService.Scope.DriveAppdata,
                DriveService.Scope.DriveReadonly,
                DriveService.Scope.DriveFile,
                DriveService.Scope.DriveMetadataReadonly,
                DriveService.Scope.DriveReadonly,
                DriveService.Scope.DriveScripts };

                UserCredential credential;

                using (var stream = new FileStream("client_secrets.json", FileMode.Open, FileAccess.Read))
                {
                    // The file token.json stores the user's access and refresh tokens, and is created
                    // automatically when the authorization flow completes for the first time.
                    string credPath = "token.json";
                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.Load(stream).Secrets, scopes, "user", CancellationToken.None, new FileDataStore(credPath, true)).Result;
                    //Console.WriteLine("Credential file saved to: " + credPath);
                }

                // Create Drive API service.
                driveService = new DriveService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = AppName,
                });
            }
            catch (Exception e)
            {
                Console.WriteLine("ERROR: Unable to Authorise connection.\n" + e.Message);
            }
        }
        public (MemoryStream, Google.Apis.Drive.v3.Data.File) GetDrive()
        {
            Authorize();
            var FileList = GetFiles(folderID);
            var File = GetFile(FileID);
            var filestream = DownloadFile(driveService, File);
            return (filestream, File);
        }

        public Google.Apis.Drive.v3.Data.File DownloadFile(string fileID)
        {
            FilesResource.GetRequest fileSearchRequest = driveService.Files.Get(fileID);
            fileSearchRequest.Fields = "id, name, size, createdTime,parents,mimeType";

            Google.Apis.Drive.v3.Data.File file = fileSearchRequest.Execute();
            return file;
        }

        public MemoryStream DownloadFile(DriveService service, Google.Apis.Drive.v3.Data.File file)
        {

            var request = service.Files.Get(file.Id);
            var stream = new MemoryStream();
           
            // Add a handler which will be notified on progress changes.
            // It will notify on each chunk download and when the
            // download is completed or failed.
            request.MediaDownloader.ProgressChanged += (progress) =>
            {
                switch (progress.Status)
                {
                    case DownloadStatus.Downloading:
                        {
                            Console.WriteLine(progress.BytesDownloaded);
                            break;
                        }
                    case DownloadStatus.Completed:
                        {
                            // System.IO.File.WriteAllBytes("C:\\Pictures\\Test.xlsx", stream.ToArray());
                            // stream2 = new MemoryStream(System.IO.File.ReadAllBytes("C:\\Pictures\\Test.xlsx"));
                            Console.WriteLine("Download complete.");
                            break;
                        }
                    case DownloadStatus.Failed:
                        {
                            Console.WriteLine("Download failed.");
                            break;
                        }
                }
            };
            request.Download(stream);
            stream.Position = 0;

            return stream;
        }
 

        public Google.Apis.Drive.v3.Data.File GetFile(string fileID)
        {
            FilesResource.GetRequest fileSearchRequest = driveService.Files.Get(fileID);
            fileSearchRequest.Fields = "id, name, size, createdTime,parents,mimeType";

            Google.Apis.Drive.v3.Data.File file = fileSearchRequest.Execute();
            return file;
        }

        public IList<Google.Apis.Drive.v3.Data.File> GetFiles(int pageSize)
        {
            FilesResource.ListRequest listRequest = driveService.Files.List();
            listRequest.PageSize = pageSize;
            listRequest.Fields = "nextPageToken, files(id, name, size, createdTime)";

            // List files.
            IList<Google.Apis.Drive.v3.Data.File> files = listRequest.Execute().Files;
           
            return files;
            // Console.WriteLine("\nList of Files stored on Google Drive:");

            // if (files != null && files.Count > 0)
            // {
            //     foreach (var file in files)
            //     {
            //         Console.WriteLine("Name: {0}\t ID: ({1}\t Size (kB): {2}\t Uploaded: {3}", file.Name, file.Id, file.Size, file.CreatedTime);
            //     }
            // }
            // else
            // {
            //     Console.WriteLine("No files found.");
            // }
        }

        public IList<Google.Apis.Drive.v3.Data.File> GetFiles(string FolderID)
        {
            IList<Google.Apis.Drive.v3.Data.File> FileList;

            var request = driveService.Files.List();
            request.PageSize = 1000;
            request.Q = string.Format("parents in '{0}'", FolderID);
            request.Fields = "files(id, name, size, createdTime, mimeType)";
            request.OrderBy = "name";
            var results = request.Execute();

            FileList = results.Files;
            return FileList;
        }

        public void UploadFile(MemoryStream stream, Google.Apis.Drive.v3.Data.File origFile)
        {
            try
            {

                Google.Apis.Drive.v3.Data.File updatedFileMetadata = new Google.Apis.Drive.v3.Data.File();
                //updatedFileMetadata.Name = origFile.Name;

                FilesResource.UpdateMediaUpload updateRequest;
                string fileId = origFile.Id;


                updateRequest = driveService.Files.Update(updatedFileMetadata, fileId, stream, origFile.MimeType);
                updateRequest.Upload();
                var newfile = updateRequest.ResponseBody;

                updatedFileMetadata = null;
                newfile = null;
                stream.Close();
                driveService.Dispose();

            }
            catch (Exception)
            {

                throw;
            }
        }

    }

}

















