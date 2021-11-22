// See https://aka.ms/new-console-template for more information
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using simple_graph_console;



var appConfig = new ConfigurationBuilder()
        .AddUserSecrets<Program>()
        .Build();

if (string.IsNullOrEmpty(appConfig["appId"]) ||
    string.IsNullOrEmpty(appConfig["scopes"]))
{
    Console.WriteLine("Missing or invalid appsettings.json...exiting\n");
    return;
}

var appId = appConfig["appId"];
var scopesString = appConfig["scopes"];
var scopes = scopesString.Split(';');

GraphHelper.Initialize(appId, scopes, (code, cancellation) =>
{
    Console.WriteLine(code.Message);
    return Task.FromResult(0);
});

var accessToken = GraphHelper.GetAccessTokenAsync(scopes).Result;
Console.WriteLine("Signed In...\n");
Console.WriteLine($"Access token: {accessToken}\n");


// Email Helper
EmailHelper.Initialize(GraphHelper.graphClient);


int choice = -1;

while (choice != 0)
{
    Console.WriteLine("Please choose one of the following options:");
    Console.WriteLine("0. Exit");
    Console.WriteLine("1. Display access token");
    Console.WriteLine("2. Hello, Me.");
    Console.WriteLine("---- Email ----");
    Console.WriteLine("3. List Emails");

    try
    {
        choice = int.Parse(Console.ReadLine());
    }
    catch (System.FormatException)
    {
        choice = -1;
    }

    switch (choice)
    {
        case 0:
            Console.WriteLine("Goodbye...");
            break;
        case 1:
            Console.WriteLine($"Access token: {accessToken}\n");
            break;
        case 2:
            var user = await GraphHelper.GetMeAsync();
            Console.WriteLine($"Hi, {user.DisplayName}\n");
            break;

        case 3:
            var emails = await EmailHelper.GetEmailsAsync();
            for (int i = 0; i < emails.Count; i++)
            {
                Message message = emails[i];

                Console.WriteLine($"{i + 1}: message {message.Id}");
                Console.WriteLine($"   Sender: {message.Sender.EmailAddress.Name} ({message.Sender.EmailAddress.Address})");
                Console.WriteLine($"   Subject: {message.Subject}");
                Console.WriteLine("\n");
            }

            break;

        default:
            Console.WriteLine("Invalid choice! Please try again.");
            break;
    }
}
