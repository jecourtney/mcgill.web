namespace McGill.Web
{
    //
    // Summary:
    //     Class to return file information to GUI
    public class ShareFileData
    {
        private string _fileName;
        private string _fileSize;
        private string _uploadDate;

        public ShareFileData( string sFileName, string sFileSize, string sUploadDate)
        {
            _fileName = sFileName;
            _fileSize = sFileSize;
            _uploadDate = sUploadDate;
        }

        public string FileName 
        {
            get { return _fileName; }
        }

        public string FileSize
        {
            get { return _fileSize; }
        }
        public string UploadDate
        {
            get { return _uploadDate; }
        }
        
    }
}