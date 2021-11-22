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
    }
}