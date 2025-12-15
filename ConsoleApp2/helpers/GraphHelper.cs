using Azure.Core;
using Azure.Identity;
using ConsoleApp2.config;
using ConsoleApp2.utils;
using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace ConsoleApp2.helpers;

public class GraphHelper
{
    private static Settings? settings;
    private static ClientSecretCredential? clientSecretCredential;
    private static GraphServiceClient? appClient;

    public static void InitializeGraphForAppOnlyAuth(Settings settings)
    {
        GraphHelper.settings = settings;

        _ = settings ??
            throw new NullReferenceException("Settings cannot be null");

        GraphHelper.settings = settings;

        clientSecretCredential ??= new ClientSecretCredential(
                GraphHelper.settings.TenantId, GraphHelper.settings.ClientId, GraphHelper.settings.ClientSecret);

        appClient ??= new GraphServiceClient(
                clientSecretCredential,
                ["https://graph.microsoft.com/.default"]);
    }

    public static async Task<string> GetAppOnlyTokenAsync()
    {
        _ = clientSecretCredential ??
            throw new NullReferenceException("Graph has not been initialized for app-only auth");

        var context = new TokenRequestContext(["https://graph.microsoft.com/.default"]);
        var response = await clientSecretCredential.GetTokenAsync(context);
        return response.Token;
    }
    public static Task<Microsoft.Graph.Models.UserCollectionResponse?> GetUsersAsync()
    {
        _ = appClient ??
            throw new NullReferenceException("Graph has not been initialized for app-only auth");

        return appClient.Users.GetAsync((config) =>
        {
            config.QueryParameters.Select = ["displayName", "id", "mail"];
            config.QueryParameters.Top = 25;
            config.QueryParameters.Orderby = ["displayName"];
        });
    }


    public static async Task<Microsoft.Graph.Models.Site?> GetSitesCallAsync()
    {
        _ = appClient ?? throw new NullReferenceException(
            "Graph has not been initialized for app-only auth");

        var response = await appClient.Sites.GetAsync(config =>
        {
            config.QueryParameters.Select = ["id", "webUrl"];
        });

        var site = response?.Value?
            .FirstOrDefault(s =>
            {
                var url = s.WebUrl?.ToString().TrimEnd('/') ?? "";
                return url.EndsWith("SPtraining", StringComparison.OrdinalIgnoreCase);
            });

        return site;
    }

    public static async Task<List<Microsoft.Graph.Models.List>> GetListsCallAsync()
    {
        _ = appClient ?? throw new NullReferenceException(
            "Graph has not been initialized for app-only auth");

        var response = await appClient.Sites[AppConstants.SPTrainingSiteId]
                                      .Lists
                                      .GetAsync(config =>
                                      {
                                          config.QueryParameters.Select = new[] { "id", "name", "webUrl" };
                                      });

        return response?.Value ?? new List<List>();
    }
    public static async Task<List<Microsoft.Graph.Models.Drive>> GetDrivesCallAsync()
    {
        _ = appClient ?? throw new NullReferenceException(
            "Graph has not been initialized for app-only auth");

        var response = await appClient.Sites[AppConstants.SPTrainingSiteId]
                                      .Drives
                                      .GetAsync(config =>
                                      {
                                          config.QueryParameters.Select = new[] { "id", "name", "webUrl" };
                                      });

        return response?.Value ?? new List<Drive>();
    }
    public static async Task<List<Microsoft.Graph.Models.DriveItem>> GetFilesInDriveAsync()
    {
        _ = appClient ?? throw new NullReferenceException(
            "Graph has not been initialized for app-only auth");

        var response = await appClient
            .Drives[AppConstants.SPTrainingDriveId]
            .Items["root"]
            .Children
            .GetAsync(config =>
            {
                config.QueryParameters.Select = new[]
                {
            "id", "name", "webUrl", "size", "file", "folder"
                };
            });

        return response?.Value?.ToList() ?? new List<DriveItem>();
    }
    public static async Task<List<Microsoft.Graph.Models.ColumnDefinition>> GetDocsFolderColumns()
    {
        _ = appClient ?? throw new NullReferenceException(
            "Graph has not been initialized for app-only auth");

        var response = await appClient.Sites[AppConstants.SPTrainingSiteId]
                                      .Lists[AppConstants.SPTrainingDocumentsFolderId]
                                      .Columns
                                      .GetAsync(config =>
                                      {
                                          config.QueryParameters.Select = new[] { "displayName" };
                                      });

        return response?.Value?
                        .ToList()
               ?? new List<ColumnDefinition>();
    }
    public static async Task<List<Microsoft.Graph.Models.ContentType>> GetDocsFolderContentTypes()
    {
        _ = appClient ?? throw new NullReferenceException(
            "Graph has not been initialized for app-only auth");

        var response = await appClient.Sites[AppConstants.SPTrainingSiteId]
                                      .Lists[AppConstants.SPTrainingDocumentsFolderId]
                                      .ContentTypes
                                      .GetAsync(config =>
                                      {
                                          config.QueryParameters.Select = new[]
                                          {
                                          "name", "description", "group"
                                          };
                                      });

        return response?.Value?.ToList()
               ?? new List<Microsoft.Graph.Models.ContentType>();
    }
    public static async Task<List<ListItem>> GetDocsFolderItemInfo()
    {
        _ = appClient ?? throw new NullReferenceException(
            "Graph has not been initialized for app-only auth");

        var response = await appClient.Sites[AppConstants.SPTrainingSiteId]
                                      .Lists[AppConstants.SPTrainingDocumentsFolderId]
                                      .Items
                                      .GetAsync(config =>
                                      {
                                          config.QueryParameters.Select = new[]
                                          {
                                          "id",
                                          "webUrl",
                                          "contentType",
                                          "fields"
                                          };

                                          config.QueryParameters.Expand = new[]
                                          {
                                          "fields"
                                          };
                                      });

        return response?.Value?.ToList() ?? new List<Microsoft.Graph.Models.ListItem>();
    }

    public static async Task<ColumnDefinition> CreateChoiceColumnAsync(string choiceColName, List<string> choiceList)
    {
        _ = appClient ?? throw new NullReferenceException(
            "Graph has not been initialized for app-only auth");

        var column = new ColumnDefinition
        {
            Description = "",
            EnforceUniqueValues = false,
            Hidden = false,
            Indexed = false,
            Name = choiceColName,
            Choice = new ChoiceColumn
            {
                AllowTextEntry = true,
                Choices = choiceList
            }
        };

        var response = await appClient.Sites[AppConstants.SPTrainingSiteId]
                                      .Lists[AppConstants.SPTrainingDocumentsFolderId]
                                      .Columns
                                      .PostAsync(column);

        return response ?? throw new InvalidOperationException(
        "Graph API returned an Exception.");
    }
    public static async Task<Microsoft.Graph.Models.ColumnDefinition> CreateNumberColumnAsync(string numberColName)
    {
        _ = appClient ?? throw new NullReferenceException(
            "Graph has not been initialized for app-only auth");

        var column = new ColumnDefinition
        {
            Description = "",
            EnforceUniqueValues = false,
            Hidden = false,
            Indexed = false,
            Name = numberColName,
            Number = new NumberColumn
            {
                DecimalPlaces = "automatic",
                DisplayAs = "number"
            }
        };

        var response = await appClient.Sites[AppConstants.SPTrainingSiteId]
                                      .Lists[AppConstants.SPTrainingDocumentsFolderId]
                                      .Columns
                                      .PostAsync(column);

        return response ?? throw new InvalidOperationException(
        "Graph API returned an Exception.");
    }
    public static async Task<Microsoft.Graph.Models.ColumnDefinition> CreateCurrencyColumnAsync(string currencyColName)
    {
        _ = appClient ?? throw new NullReferenceException(
            "Graph has not been initialized for app-only auth");

        var column = new ColumnDefinition
        {
            Description = "",
            EnforceUniqueValues = false,
            Hidden = false,
            Indexed = false,
            Name = currencyColName,
            Currency = new CurrencyColumn
            {
                Locale = "en-ca"
            }
        };

        var response = await appClient.Sites[AppConstants.SPTrainingSiteId]
                                      .Lists[AppConstants.SPTrainingDocumentsFolderId]
                                      .Columns
                                      .PostAsync(column);

        return response ?? throw new InvalidOperationException(
        "Graph API returned an Exception.");
    }
    public static async Task<Microsoft.Graph.Models.ColumnDefinition> CreateDateTimeColumnAsync(string dateTimeColName)
    {
        _ = appClient ?? throw new NullReferenceException(
            "Graph has not been initialized for app-only auth");

        var column = new ColumnDefinition
        {
            Description = "",
            EnforceUniqueValues = false,
            Hidden = false,
            Indexed = false,
            Name = dateTimeColName,
            DateTime = new DateTimeColumn
            {
                DisplayAs = "default",
                Format = "dateTime"
            }
        };

        var response = await appClient.Sites[AppConstants.SPTrainingSiteId]
                                      .Lists[AppConstants.SPTrainingDocumentsFolderId]
                                      .Columns
                                      .PostAsync(column);

        return response ?? throw new InvalidOperationException(
        "Graph API returned an Exception.");
    }
    public static async Task<Microsoft.Graph.Models.ColumnDefinition> CreateLookUpColumnAsync(string lookUpColName)
    {
        _ = appClient ?? throw new NullReferenceException(
            "Graph has not been initialized for app-only auth");

        var column = new ColumnDefinition
        {
            Description = "",
            EnforceUniqueValues = false,
            Hidden = false,
            Indexed = false,
            Name = lookUpColName,
            Lookup = new LookupColumn
            {
                AllowMultipleValues = false,
                AllowUnlimitedLength = false,
                ColumnName = "ID",
                ListId = "a27bf34e-444f-433b-8874-20eb73a7ef35" // TODO: Take to Constants or dinamically create it
            }
        };

        var response = await appClient.Sites[AppConstants.SPTrainingSiteId]
                                      .Lists[AppConstants.SPTrainingDocumentsFolderId]
                                      .Columns
                                      .PostAsync(column);

        return response ?? throw new InvalidOperationException(
        "Graph API returned an Exception.");
    }
    public static async Task<Microsoft.Graph.Models.ColumnDefinition> CreateBooleanColumnAsync(string boolColName)
    {
        _ = appClient ?? throw new NullReferenceException(
            "Graph has not been initialized for app-only auth");

        var column = new ColumnDefinition
        {
            Description = "",
            EnforceUniqueValues = false,
            Hidden = false,
            Indexed = false,
            Name = boolColName,
            Boolean = new BooleanColumn
            {
            }
        };

        var response = await appClient.Sites[AppConstants.SPTrainingSiteId]
                                      .Lists[AppConstants.SPTrainingDocumentsFolderId]
                                      .Columns
                                      .PostAsync(column);

        return response ?? throw new InvalidOperationException(
        "Graph API returned an Exception.");
    }
    public static async Task<Microsoft.Graph.Models.ColumnDefinition> CreatePersonGroupColumnAsync(string pergroupColName)
    {
        _ = appClient ?? throw new NullReferenceException(
            "Graph has not been initialized for app-only auth");

        var column = new ColumnDefinition
        {
            Description = "",
            EnforceUniqueValues = false,
            Hidden = false,
            Indexed = false,
            Name = pergroupColName,
            PersonOrGroup = new PersonOrGroupColumn
            {
                AllowMultipleSelection = true,
                DisplayAs = "account",
                ChooseFromType = "peopleAndGroups"
            }
        };

        var response = await appClient.Sites[AppConstants.SPTrainingSiteId]
                                      .Lists[AppConstants.SPTrainingDocumentsFolderId]
                                      .Columns
                                      .PostAsync(column);

        return response ?? throw new InvalidOperationException(
        "Graph API returned an Exception.");
    }
    public static async Task<Microsoft.Graph.Models.ColumnDefinition> CreateHyperlinkColumnAsync(string hyperlinkColName)
    {
        _ = appClient ?? throw new NullReferenceException(
            "Graph has not been initialized for app-only auth");

        var column = new ColumnDefinition
        {
            Description = "",
            EnforceUniqueValues = false,
            Hidden = false,
            Indexed = false,
            Name = hyperlinkColName,
            HyperlinkOrPicture = new HyperlinkOrPictureColumn
            {
                IsPicture = false,
            }
        };

        var response = await appClient.Sites[AppConstants.SPTrainingSiteId]
                                      .Lists[AppConstants.SPTrainingDocumentsFolderId]
                                      .Columns
                                      .PostAsync(column);

        return response ?? throw new InvalidOperationException(
        "Graph API returned an Exception.");
    }
    public static async Task<Microsoft.Graph.Models.ContentType> CreateCustomContentTypeAsync(string contentName,
                                                                                                string contentDescription,
                                                                                                string contentCategory)
    {
        _ = appClient ?? throw new NullReferenceException(
            "Graph has not been initialized for app-only auth");

        var newContentType = new Microsoft.Graph.Models.ContentType
        {
            Name = contentName,
            Description = contentDescription,
            Group = contentCategory,
            Base = new Microsoft.Graph.Models.ContentType
            {
                Id = "0x0101",
                Name = "Document2"
            }
        };

        var createdContentType = await appClient.Sites[AppConstants.SPTrainingSiteId]
                                                .ContentTypes
                                                .PostAsync(newContentType);

        return createdContentType ?? throw new InvalidOperationException(
        "Graph API returned an Exception.");
    }
    public static async Task<DriveItem> CreateDocumentSetFolderAsync(
    string documentSetName)
    {
        var folder = new DriveItem
        {
            Name = documentSetName,
            Folder = new Folder(),
            AdditionalData = new Dictionary<string, object>
        {
            { "@microsoft.graph.conflictBehavior", "fail" }
        }
        };

        return await appClient
            .Drives[AppConstants.SPTrainingDriveId]
            .Items
            .PostAsync(folder) ?? throw new InvalidOperationException(
        "Graph API returned an Exception.");
    }
    public static async Task UpdateListItemFieldAsync(string documentid, string documentNameNew)
    {
        _ = appClient ?? throw new NullReferenceException(
            "Graph has not been initialized for app-only auth");

        var updateItem = new DriveItem
        {
            Name = documentNameNew
        };

        await appClient
                    .Drives[AppConstants.SPTrainingDriveId]
                    .Items[documentid]
                    .PatchAsync(updateItem);
    }

    public static async Task<Microsoft.Graph.Models.ListItem> UpdateDocumentSetFieldAsync(
                                            string documentSetId,
                                            string itemId,
                                            string fieldName,
                                            object fieldValue)
    {
        _ = appClient ?? throw new NullReferenceException(
            "Graph has not been initialized for app-only auth");

        var updateItem = new ListItem
        {
            ContentType = new ContentTypeInfo
            {
                Id = documentSetId
            },
            Fields = new FieldValueSet
            {
                AdditionalData = new Dictionary<string, object>
            {
                { fieldName, fieldValue }
            }
            }
        };

        var updatedListItem = await appClient.Sites[AppConstants.SPTrainingSiteId]
                                                .Lists[AppConstants.SPTrainingDocumentsFolderId]
                                                .Items[itemId]
                                                .PatchAsync(updateItem);

        return updatedListItem ?? throw new InvalidOperationException(
        "Graph API returned an Exception.");
    }
    public static async Task<ListItem> UpdateDocumentSetMetadataAsync(
        string itemId,
        string documentSetContentTypeId,
        Dictionary<string, object?> fields)
    {
        _ = appClient ?? throw new InvalidOperationException(
            "Graph has not been initialized for app-only auth");

        fields["ContentTypeId"] = documentSetContentTypeId;
        fields["Title"] = fields.TryGetValue("Title", out var title) ? title : null;

        var updateItem = new ListItem
        {
            Fields = new FieldValueSet
            {
                AdditionalData = fields
            }
        };

        var response = await appClient
            .Sites[AppConstants.SPTrainingSiteId]
            .Lists[AppConstants.SPTrainingDocumentsFolderId]
            .Items[itemId]
            .PatchAsync(updateItem);

        return response ?? throw new InvalidOperationException(
            "Graph API returned null when updating the list item.");
    }

}