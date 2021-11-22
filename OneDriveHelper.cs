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
            if (!name.Contains("/"))
            {
                var folderNames = name.Split("/");

                String nameForNewFolder;
                String createdFolderId = "";
                DriveItem createdFolder;
                DriveItem newFolder;

                for (int i = 0; i < folderNames.Length; i++)
                {
                    nameForNewFolder = folderNames[i];

                    newFolder = new DriveItem
                    {
                        Name = nameForNewFolder,
                        Folder = new Folder
                        {
                        },
                        AdditionalData = new Dictionary<string, object>()
                        {
                            {"@microsoft.graph.conflictBehavior", "rename"}
                        }
                    };

                    if (i == 0)
                    {
                        createdFolder = await graphClient.Me.Drive.Root.Children
                        .Request()
                        .AddAsync(newFolder);

                    }
                    else
                    {
                        createdFolder = await graphClient.Me.Drive.Items[createdFolder].Children
                        .Request()
                        .AddAsync(newFolder);
                    }
                }
            }
            else 
            {

            }

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
    }
}