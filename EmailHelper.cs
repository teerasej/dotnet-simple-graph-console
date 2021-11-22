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

    }
}