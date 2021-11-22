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
                Console.WriteLine($"Error creating new folder in onedrive: {ex.Message}");
            }

        }
    }
}