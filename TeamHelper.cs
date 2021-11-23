using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Graph;

namespace simple_graph_console
{
    public class TeamHelper
    {
        private static GraphServiceClient graphClient;

        public static void Initialize(GraphServiceClient client)
        {
            graphClient = client;
        }

        public static async Task<Team> CreateTeamAsync(string teamName, string description) 
        {
            var newTeam = new Team
            {
                DisplayName = teamName,
                Description = description,
                AdditionalData = new Dictionary<string, object>()
                {
                    {"template@odata.bind", "https://graph.microsoft.com/v1.0/teamsTemplates('standard')"}
                }
            };

            try
            {
                return await graphClient.Teams.Request().AddAsync(newTeam);
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error creating team: {ex.Message}");
                return null;
            }
        }
    }
}