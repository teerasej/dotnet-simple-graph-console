using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Graph;

namespace simple_graph_console
{
    public class OneDriveHelper
    {
        private static GraphServiceClient graphClient;

        public static void Initialize(GraphServiceClient client)
        {
            graphClient = client;
        }

        public static async Task<IDriveItemChildrenCollectionPage> GetUserDriveItemsAsync()
        {
            try
            {
                return await graphClient.Me.Drive.Root.Children.Request().GetAsync();
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting user's drive items: {ex.Message}");
                return null;
            }
        }

        public static async Task<DriveItem> CreateNewFolderAsync(string name = "temp")
        {
            var driveItem = new DriveItem
            {
                Name = name,
                Folder = new Folder
                {
                },
                AdditionalData = new Dictionary<string, object>()
                {
                    {"@microsoft.graph.conflictBehavior", "rename"}
                }
            };

            try
            {
                return await graphClient.Me.Drive.Root.Children
                .Request()
                .AddAsync(driveItem);
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error creating new folder in onedrive: {ex.Message}");
                return null;
            }
        }

        public static async Task DownloadFileAsync(string driveItemId)
        {
            try
            {
                var driveItem = await graphClient.Me.Drive.Items[driveItemId].Request().GetAsync();
                var stream = await graphClient.Me.Drive.Items[driveItemId].Content.Request().GetAsync();

                using (var fileStream = System.IO.File.Create(driveItem.Name))
                {
                    stream.Seek(0, SeekOrigin.Begin);
                    stream.CopyTo(fileStream);
                }
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error downloading a file: {ex.Message}");
            }

        }

        public static async Task UploadFileAsync(string fileName)
        {

            using (var fileStream = System.IO.File.OpenRead(fileName))
            {
                // Use properties to specify the conflict behavior
                // in this case, replace
                var uploadProps = new DriveItemUploadableProperties
                {
                    ODataType = null,
                    AdditionalData = new Dictionary<string, object>
                    {
                        { "@microsoft.graph.conflictBehavior", "replace" }
                    }
                };

                // Create the upload session
                // itemPath does not need to be a path to an existing item
                var uploadSession = await graphClient.Me.Drive.Root
                    .ItemWithPath(fileName)
                    .CreateUploadSession(uploadProps)
                    .Request()
                    .PostAsync();

                // Max slice size must be a multiple of 320 KiB
                int maxSliceSize = 320 * 1024;
                var fileUploadTask =
                    new LargeFileUploadTask<DriveItem>(uploadSession, fileStream, maxSliceSize);

                // Create a callback that is invoked after each slice is uploaded
                IProgress<long> progress = new Progress<long>(prog =>
                {
                    Console.WriteLine($"        Uploaded {prog} bytes of {fileStream.Length} bytes");
                });

                try
                {
                    // Upload the file
                    var uploadResult = await fileUploadTask.UploadAsync(progress);

                    if (uploadResult.UploadSucceeded)
                    {
                        // The ItemResponse object in the result represents the
                        // created item.
                        Console.WriteLine($"        Upload complete, item ID: {uploadResult.ItemResponse.Id}");
                    }
                    else
                    {
                        Console.WriteLine("         Upload failed");
                    }
                }
                catch (ServiceException ex)
                {
                    Console.WriteLine($"        Error uploading: {ex.ToString()}");
                }
            }

        }
    }
}