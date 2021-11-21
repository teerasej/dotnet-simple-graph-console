﻿// See https://aka.ms/new-console-template for more information
using Microsoft.Extensions.Configuration;
using simple_graph_console;

Console.WriteLine("Hello, World!");

var appConfig = LoadAppSettings();

if (appConfig == null)
{
    Console.WriteLine("Missing or invalid appsettings.json...exiting");
    return;
}

var appId = appConfig["appId"];
var scopesString = appConfig["scopes"];
var scopes = scopesString.Split(';');

// Initialize Graph client
GraphHelper.Initialize(appId, scopes, (code, cancellation) => {
    Console.WriteLine(code.Message);
    return Task.FromResult(0);
});

var accessToken = GraphHelper.GetAccessTokenAsync(scopes).Result;
Console.WriteLine($"Access token: {accessToken}\n");

int choice = -1;

while (choice != 0)
{
    Console.WriteLine("Please choose one of the following options:");
    Console.WriteLine("0. Exit");
    Console.WriteLine("1. Display access token");
    Console.WriteLine("2. View this week's calendar");
    Console.WriteLine("3. Add an event");

    try
    {
        choice = int.Parse(Console.ReadLine());
    }
    catch (System.FormatException)
    {
        // Set to invalid value
        choice = -1;
    }

    switch (choice)
    {
        case 0:
            // Exit the program
            Console.WriteLine("Goodbye...");
            break;
        case 1:
            // Display access token
            break;
        case 2:
            // List the calendar
            break;
        case 3:
            // Create a new event
            break;
        default:
            Console.WriteLine("Invalid choice! Please try again.");
            break;
    }
}

static IConfigurationRoot LoadAppSettings()
{
    var appConfig = new ConfigurationBuilder()
        .AddUserSecrets<Program>()
        .Build();

    // Check for required settings
    if (string.IsNullOrEmpty(appConfig["appId"]) ||
        string.IsNullOrEmpty(appConfig["scopes"]))
    {
        return null;
    }

    return appConfig;
}