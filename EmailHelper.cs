using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Graph;

namespace simple_graph_console
{
    public class EmailHelper
    {
        private static GraphServiceClient graphClient;

        public static void Initialize(GraphServiceClient client)
        {
            graphClient = client;
        }

        public static async Task<IUserMessagesCollectionPage> GetEmailsAsync()
        {
            try
            {
                return await graphClient.Me.Messages.Request().GetAsync();
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting user's message: {ex.Message}");
                return null;
            }
        }
    }
}