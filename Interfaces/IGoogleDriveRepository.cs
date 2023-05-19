using Domain.Models;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Drive.v3;

namespace Application.Interfaces
{
    public interface IGoogleDriveRepository
    {
        public void Authorize();

        public (MemoryStream, Google.Apis.Drive.v3.Data.File) GetDrive();        
         public MemoryStream DownloadFile(DriveService service, Google.Apis.Drive.v3.Data.File file);
        public Google.Apis.Drive.v3.Data.File DownloadFile(string fileID);
        public void UploadFile(MemoryStream stream, Google.Apis.Drive.v3.Data.File origFile);
        public Google.Apis.Drive.v3.Data.File GetFile(string fileID);
        public IList<Google.Apis.Drive.v3.Data.File> GetFiles(string FolderID);
        public IList<Google.Apis.Drive.v3.Data.File> GetFiles(int pageSize);
        
      
    
       



    }
}
