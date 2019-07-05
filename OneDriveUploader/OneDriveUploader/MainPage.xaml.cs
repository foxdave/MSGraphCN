using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Windows.Storage;
using Windows.Storage.Pickers;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;

// https://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x804 上介绍了“空白页”项模板

namespace OneDriveUploader
{
    /// <summary>
    /// 可用于自身或导航至 Frame 内部的空白页。
    /// </summary>
    public sealed partial class MainPage : Page
    {
        public MainPage()
        {
            this.InitializeComponent();
        }

        private const string AADClientId = "451912ea-1728-4fc5-b27d-453821c61aa6";
        private const string GraphAPIEndpointPrefix = "https://graph.microsoft.com/v1.0/";
        private string[] AADScopes = new string[] { "files.readwrite.all" };
        private IPublicClientApplication AADAppContext = null;
        private GraphServiceClient graphClient = null;
        private AuthenticationResult userCredentials;
        public AuthenticationResult UserCredentials
        {
            get { return userCredentials; }
            set { userCredentials = value; }
        }

        public void InitializeGraph()
        {
            if (userCredentials != null)
            {
                graphClient = new GraphServiceClient(
                GraphAPIEndpointPrefix,
                 new DelegateAuthenticationProvider(
                 async (requestMessage) =>
                 {
                     requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", userCredentials.AccessToken);
                 }
                 )
                 );
            }
        }

        /// <summary> 
        /// Log the user in to either Office 365 or OneDrive consumer 
        /// </summary> 
        /// <returns>A task to await on</returns> 
        public async Task<string> SignInUser()
        {
            string status = "Unknown";

            // Instantiate the app with AAD 
            AADAppContext = PublicClientApplicationBuilder.Create(AADClientId).Build();

            try
            {
                UserCredentials = await AADAppContext.AcquireTokenInteractive(AADScopes).ExecuteAsync();
                if (UserCredentials != null)
                {
                    status = "Signed in as " + UserCredentials.Account.Username;
                    InitializeGraph();
                }
            }

            catch (MsalServiceException serviceEx)
            {
                status = $"Could not sign in, error code: " + serviceEx.ErrorCode;
            }

            catch (Exception ex)
            {
                status = $"Error Acquiring Token: {ex}";
            }

            return (status);
        }

        ///<summary> 
        /// Take a file and upload it to the service 
        /// </summary> 
        /// <param name="fileToUpload">The file that we want to upload</param> 
        /// <param name="uploadToSharePoint">Should we upload to SharePoint or OneDrive?</param> 
        public async Task<DriveItem> UploadSmallFile(StorageFile fileToUpload, bool uploadToSharePoint = false)
        {
            Stream fileStream = (await fileToUpload.OpenReadAsync()).AsStreamForRead();
            DriveItem uploadedFile = null;

            // Do we want OneDrive for Business/Consumer or do we want a SharePoint Site? 
            if (uploadToSharePoint)
            {
                uploadedFile = await graphClient.Sites["root"].Drive.Root.ItemWithPath(fileToUpload.Name).Content.Request().PutAsync<DriveItem>(fileStream);
            }
            else
            {
                uploadedFile = await graphClient.Me.Drive.Root.ItemWithPath(fileToUpload.Name).Content.Request().PutAsync<DriveItem>(fileStream);
            }
            return (uploadedFile);
        }

        /// <summary> 
        /// Take a file greater than 4MB and upload it to the service 
        /// </summary> 
        /// <param name="fileToUpload">The file that we want to upload</param> 
        /// <param name="uploadToSharePoint">Should we upload to SharePoint or OneDrive?</param> 
        public async Task<DriveItem> UploadLargeFile(StorageFile fileToUpload, bool uploadToSharePoint = false)
        {
            Stream fileStream = (await fileToUpload.OpenReadAsync()).AsStreamForRead();
            DriveItem uploadedFile = null;
            UploadSession uploadSession = null;

            // Do we want OneDrive for Business/Consumer or do we want a SharePoint Site? 
            if (uploadToSharePoint)
            {
                uploadSession = await graphClient.Sites["root"].Drive.Root.ItemWithPath(fileToUpload.Name).CreateUploadSession().Request().PostAsync();
            }
            else
            {
                uploadSession = await graphClient.Me.Drive.Root.ItemWithPath(fileToUpload.Name).CreateUploadSession().Request().PostAsync();
            }

            if (uploadSession != null)
            {
                // Chunk size must be divisible by 320KiB, our chunk size will be slightly more than 1MB 
                int maxSizeChunk = (320 * 1024) * 4;
                ChunkedUploadProvider uploadProvider = new ChunkedUploadProvider(uploadSession, graphClient, fileStream, maxSizeChunk);
                var chunkRequests = uploadProvider.GetUploadChunkRequests();
                var exceptions = new List<Exception>();
                var readBuffer = new byte[maxSizeChunk];

                foreach (var request in chunkRequests)
                {
                    var result = await uploadProvider.GetChunkRequestResponseAsync(request, readBuffer, exceptions);

                    if (result.UploadSucceeded)
                    {
                        uploadedFile = result.ItemResponse;
                    }
                }
            }
            return (uploadedFile);
        }

        private async Task<StorageFile> PickFile()
        {
            var picker = new FileOpenPicker();
            picker.ViewMode = PickerViewMode.Thumbnail;
            picker.SuggestedStartLocation = PickerLocationId.PicturesLibrary;
            picker.FileTypeFilter.Add(".jpg");
            picker.FileTypeFilter.Add(".jpeg");
            picker.FileTypeFilter.Add(".png");

            StorageFile pickedFile = await picker.PickSingleFileAsync();
            return (pickedFile);
        }

        private async Task UploadFile(object whichButton)
        {
            if (this.UserCredentials == null)
            {
                await SignInUser();
            }

            StorageFile fileToUpload = await PickFile();
            DriveItem uploadedFile = null;

            if (whichButton == this.uploadSmallFileButton)
            {
                uploadedFile = await UploadSmallFile(fileToUpload);
            }
            else
            {
                uploadedFile = await UploadLargeFile(fileToUpload);
            }

            if (uploadedFile != null)
            {
                this.statusTextBlock.Text = "Uploaded file: " + uploadedFile.Name;
            }
            else
            {
                this.statusTextBlock.Text = "Upload failed";
            }
        }

        private async void uploadSmallFileButton_Click(object sender, RoutedEventArgs e)
        {
            await UploadFile(sender);
        }

        private async void uploadLargeFileButton_Click(object sender, RoutedEventArgs e)
        {
            await UploadFile(sender);
        }
    }
}
