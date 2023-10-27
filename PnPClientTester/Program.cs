using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using PnP.Core.Auth;
using PnP.Core.Model.SharePoint;
using PnP.Core.Services;
using System.Xml;

// registered app id for PnPCoreSDKConsoleDemo (redirect url set to http://localhost) with delegated permissions:
// Microsoft Graph -> Sites.Manage.All
// SharePoint -> AllSites.Manage
string clientId = "<registered-app-id>";
string siteUrl = "<site-url>";

// Creates and configures the host
var host = Host.CreateDefaultBuilder()
    .ConfigureServices((context, services) =>
    {
        // Add PnP Core SDK
        services.AddPnPCore(options =>
        {
            // Configure the interactive authentication provider as default
            options.DefaultAuthenticationProvider = new InteractiveAuthenticationProvider()
            {
                ClientId = clientId,
                RedirectUri = new Uri("http://localhost")
            };
        });
    })
    .UseConsoleLifetime()
    .Build();

// Start the host
await host.StartAsync();

using (var scope = host.Services.CreateScope())
{
    // Ask an IPnPContextFactory from the host
    var pnpContextFactory = scope.ServiceProvider.GetRequiredService<IPnPContextFactory>();

    // Create a PnPContext
    using (var context = await pnpContextFactory.CreateAsync(new Uri(siteUrl)))
    {
        // Load the Title property of the site's root web
        await context.Web.LoadAsync(p => p.Title);
        Console.WriteLine($"The title of the web is {context.Web.Title}");
        
        var documentsLibrary = context.Web.Lists.GetByTitle("Documents");
        // Upload a file by adding it to the folder's files collection
        IFile addedFile = await documentsLibrary.RootFolder.Files.AddAsync("lorem.docx",
                  System.IO.File.OpenRead($".{Path.DirectorySeparatorChar}TestFilesFolder{Path.DirectorySeparatorChar}lorem.docx"));
        // note alternative for large files: https://github.com/pnp/pnpcore/blob/dev/docs/using-the-sdk/files-large.md#uploading-large-files
    }
}

// next up for managed metadata:
// https://github.com/pnp/pnpcore/blob/dev/docs/using-the-sdk/taxonomy-intro.md#working-with-taxonomy-data-an-introduction