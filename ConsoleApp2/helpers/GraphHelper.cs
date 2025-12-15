using Azure.Core;
using Azure.Identity;
using ConsoleApp2.config;
using ConsoleApp2.utils;
using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace ConsoleApp2.helpers;

/// <summary>
/// Provides helper methods for authenticating with Microsoft Graph using app-only authentication and for performing
/// common operations on users, sites, lists, drives, and document libraries in Microsoft 365 environments.
/// </summary>
/// <remarks>This class is intended for use in applications that require app-only access to Microsoft Graph
/// resources, such as SharePoint sites and document libraries. Before calling any methods that interact with Graph
/// resources, you must initialize the class by calling InitializeGraphForAppOnlyAuth with valid application settings.
/// All methods are static and thread safety is not guaranteed; ensure proper synchronization if used in multi-threaded
/// scenarios.</remarks>
public class GraphHelper
{
    private static Settings? settings;
    private static ClientSecretCredential? clientSecretCredential;
    private static GraphServiceClient? appClient;

    /// <summary>
    /// Initializes the Microsoft Graph client for application-only authentication using the specified settings.
    /// </summary>
    /// <remarks>Call this method before making requests that require application-only authentication with
    /// Microsoft Graph. This method configures the Graph client to use the provided credentials for subsequent
    /// operations.</remarks>
    /// <param name="settings">The application settings containing authentication information such as Tenant ID, Client ID, and Client Secret.
    /// Cannot be null.</param>
    /// <exception cref="NullReferenceException">Thrown if the settings parameter is null.</exception>
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

    /// <summary>
    /// Asynchronously acquires an app-only access token for Microsoft Graph using the configured client credentials.
    /// </summary>
    /// <remarks>This method is intended for scenarios where the application needs to access Microsoft Graph
    /// without user context, using only its own identity. Ensure that the client credentials are properly configured
    /// before calling this method.</remarks>
    /// <returns>A task that represents the asynchronous operation. The task result contains the access token string for
    /// Microsoft Graph.</returns>
    /// <exception cref="NullReferenceException">Thrown if the client credentials have not been initialized for app-only authentication.</exception>
    public static async Task<string> GetAppOnlyTokenAsync()
    {
        _ = clientSecretCredential ??
            throw new NullReferenceException("Graph has not been initialized for app-only auth");

        var context = new TokenRequestContext(["https://graph.microsoft.com/.default"]);
        var response = await clientSecretCredential.GetTokenAsync(context);
        return response.Token;
    }

    /// <summary>
    /// Retrieves a collection of users from Microsoft Graph using app-only authentication.
    /// </summary>
    /// <remarks>The returned collection is ordered by display name and includes only the display name, ID,
    /// and mail properties for each user. This method requires that the Graph client is properly initialized with
    /// app-only permissions before calling.</remarks>
    /// <returns>A task that represents the asynchronous operation. The task result contains a <see
    /// cref="Microsoft.Graph.Models.UserCollectionResponse"/> with up to 25 users, including their display name, ID,
    /// and email address. Returns <see langword="null"/> if the request fails or no users are found.</returns>
    /// <exception cref="NullReferenceException">Thrown if the Graph client has not been initialized for app-only authentication.</exception>
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

    /// <summary>
    /// Asynchronously retrieves the first SharePoint site whose web URL ends with "SPtraining" using the Microsoft
    /// Graph API.
    /// </summary>
    /// <remarks>Only the site ID and web URL are retrieved for each site. This method requires that the Graph
    /// client is properly initialized before calling.</remarks>
    /// <returns>A task that represents the asynchronous operation. The task result contains a <see
    /// cref="Microsoft.Graph.Models.Site"/> object representing the matching site if found; otherwise, <see
    /// langword="null"/>.</returns>
    /// <exception cref="NullReferenceException">Thrown if the Graph client has not been initialized for app-only authentication.</exception>
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

    /// <summary>
    /// Retrieves a collection of SharePoint lists from the configured site using app-only authentication.
    /// </summary>
    /// <remarks>Only the list ID, name, and web URL are included in the returned list objects. Ensure that
    /// the Graph client is properly initialized before calling this method.</remarks>
    /// <returns>A list of SharePoint lists from the specified site. Returns an empty list if no lists are found.</returns>
    /// <exception cref="NullReferenceException">Thrown if the Graph client has not been initialized for app-only authentication.</exception>
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

    /// <summary>
    /// Asynchronously retrieves a list of drives from the configured SharePoint site using app-only authentication.
    /// </summary>
    /// <returns>A list of <see cref="Microsoft.Graph.Models.Drive"/> objects representing the drives available on the SharePoint
    /// site. Returns an empty list if no drives are found.</returns>
    /// <exception cref="NullReferenceException">Thrown if the Microsoft Graph client has not been initialized for app-only authentication.</exception>
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

    /// <summary>
    /// Asynchronously retrieves a list of files and folders from the root of the configured SharePoint drive.
    /// </summary>
    /// <remarks>The returned list includes basic metadata for each item, such as ID, name, web URL, size, and
    /// type information. Only items from the root directory of the configured drive are returned.</remarks>
    /// <returns>A task that represents the asynchronous operation. The task result contains a list of <see
    /// cref="Microsoft.Graph.Models.DriveItem"/> objects representing the files and folders in the root of the drive.
    /// Returns an empty list if no items are found.</returns>
    /// <exception cref="NullReferenceException">Thrown if the Microsoft Graph client has not been initialized for app-only authentication.</exception>
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

    /// <summary>
    /// Retrieves the list of column definitions for the training documents folder in the configured SharePoint site.
    /// </summary>
    /// <returns>A list of <see cref="Microsoft.Graph.Models.ColumnDefinition"/> objects representing the columns in the training
    /// documents folder. Returns an empty list if no columns are found.</returns>
    /// <exception cref="NullReferenceException">Thrown if the Microsoft Graph client has not been initialized for app-only authentication.</exception>
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

    /// <summary>
    /// Retrieves the list of content types defined for the training documents folder in the configured SharePoint site.
    /// </summary>
    /// <returns>A list of content types associated with the training documents folder. Returns an empty list if no content types
    /// are found.</returns>
    /// <exception cref="NullReferenceException">Thrown if the Microsoft Graph client has not been initialized for app-only authentication.</exception>
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

    /// <summary>
    /// Retrieves information about all items in the configured SharePoint documents folder using app-only
    /// authentication.
    /// </summary>
    /// <returns>A list of <see cref="Microsoft.Graph.Models.ListItem"/> objects representing the items in the SharePoint
    /// documents folder. Returns an empty list if no items are found.</returns>
    /// <exception cref="NullReferenceException">Thrown if the Graph client has not been initialized for app-only authentication.</exception>
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

    /// <summary>
    /// Creates a new choice column in the specified SharePoint list asynchronously.
    /// </summary>
    /// <param name="choiceColName">The name of the choice column to create. Cannot be null or empty.</param>
    /// <param name="choiceList">A list of string values representing the available choices for the column. Cannot be null.</param>
    /// <returns>A task that represents the asynchronous operation. The task result contains the definition of the created choice
    /// column.</returns>
    /// <exception cref="NullReferenceException">Thrown if the Graph client has not been initialized for app-only authentication.</exception>
    /// <exception cref="InvalidOperationException">Thrown if the Microsoft Graph API returns an error or does not return a column definition.</exception>
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

    /// <summary>
    /// Creates a new number column in the specified SharePoint list asynchronously.
    /// </summary>
    /// <remarks>The created number column will have automatic decimal places and will be displayed as a
    /// number. The column is added to the SharePoint list identified by application constants.</remarks>
    /// <param name="numberColName">The name to assign to the new number column. Cannot be null or empty.</param>
    /// <returns>A task that represents the asynchronous operation. The task result contains the created number column
    /// definition.</returns>
    /// <exception cref="NullReferenceException">Thrown if the Microsoft Graph client has not been initialized for app-only authentication.</exception>
    /// <exception cref="InvalidOperationException">Thrown if the Microsoft Graph API returns an error or does not return a column definition.</exception>
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

    /// <summary>
    /// Creates a new currency column in the specified SharePoint list using the Microsoft Graph API.
    /// </summary>
    /// <remarks>The created currency column uses the 'en-ca' locale. Ensure that the Microsoft Graph client
    /// is properly initialized before calling this method.</remarks>
    /// <param name="currencyColName">The name to assign to the new currency column. Cannot be null or empty.</param>
    /// <returns>A task that represents the asynchronous operation. The task result contains the definition of the created
    /// currency column.</returns>
    /// <exception cref="NullReferenceException">Thrown if the Microsoft Graph client has not been initialized for app-only authentication.</exception>
    /// <exception cref="InvalidOperationException">Thrown if the Microsoft Graph API returns an error or fails to create the column.</exception>
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

    /// <summary>
    /// Asynchronously creates a new date and time column in the specified SharePoint list using Microsoft Graph.
    /// </summary>
    /// <remarks>The column is created in the SharePoint list and configured with default date and time
    /// settings. Ensure that the Microsoft Graph client is properly initialized before calling this method.</remarks>
    /// <param name="dateTimeColName">The name to assign to the new date and time column. Cannot be null or empty.</param>
    /// <returns>A task that represents the asynchronous operation. The task result contains the definition of the created date
    /// and time column.</returns>
    /// <exception cref="NullReferenceException">Thrown if the Microsoft Graph client has not been initialized for app-only authentication.</exception>
    /// <exception cref="InvalidOperationException">Thrown if the Microsoft Graph API returns an error or does not return a column definition.</exception>
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

    /// <summary>
    /// Creates a new lookup column definition in the specified SharePoint list asynchronously.
    /// </summary>
    /// <remarks>The created lookup column references a specific list and column as defined in the method.
    /// Ensure that the target list and column exist and that the application has the necessary permissions to create
    /// columns in the specified SharePoint list.</remarks>
    /// <param name="lookUpColName">The name to assign to the new lookup column. Cannot be null or empty.</param>
    /// <returns>A task that represents the asynchronous operation. The task result contains the created lookup column
    /// definition.</returns>
    /// <exception cref="NullReferenceException">Thrown if the Microsoft Graph client has not been initialized for app-only authentication.</exception>
    /// <exception cref="InvalidOperationException">Thrown if the Microsoft Graph API returns an error or does not return a column definition.</exception>
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

    /// <summary>
    /// Asynchronously creates a new boolean column in the specified SharePoint list using Microsoft Graph.
    /// </summary>
    /// <param name="boolColName">The name of the boolean column to create. Cannot be null or empty.</param>
    /// <returns>A task that represents the asynchronous operation. The task result contains the definition of the created
    /// boolean column.</returns>
    /// <exception cref="NullReferenceException">Thrown if the Microsoft Graph client has not been initialized for app-only authentication.</exception>
    /// <exception cref="InvalidOperationException">Thrown if the Microsoft Graph API returns an error or does not return a column definition.</exception>
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

    /// <summary>
    /// Creates a new person or group column in the specified SharePoint list asynchronously.
    /// </summary>
    /// <remarks>The created column allows selection of multiple people or groups and is added to the
    /// configured SharePoint list. Ensure that the Microsoft Graph client is properly initialized before calling this
    /// method.</remarks>
    /// <param name="pergroupColName">The display name for the new person or group column. Cannot be null or empty.</param>
    /// <returns>A task that represents the asynchronous operation. The task result contains the created column definition.</returns>
    /// <exception cref="NullReferenceException">Thrown if the Microsoft Graph client has not been initialized for app-only authentication.</exception>
    /// <exception cref="InvalidOperationException">Thrown if the Microsoft Graph API returns an error or does not return a column definition.</exception>
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

    /// <summary>
    /// Creates a new SharePoint column of type Hyperlink in the specified list asynchronously.
    /// </summary>
    /// <remarks>The created column will be added to the configured SharePoint list using app-only
    /// authentication. The column is configured as a Hyperlink (not a Picture) and does not enforce unique values or
    /// indexing.</remarks>
    /// <param name="hyperlinkColName">The name to assign to the new Hyperlink column. Cannot be null or empty.</param>
    /// <returns>A task that represents the asynchronous operation. The task result contains the created ColumnDefinition for the
    /// new Hyperlink column.</returns>
    /// <exception cref="NullReferenceException">Thrown if the Microsoft Graph client has not been initialized for app-only authentication.</exception>
    /// <exception cref="InvalidOperationException">Thrown if the Microsoft Graph API returns an error or does not return a valid response.</exception>
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

    /// <summary>
    /// Creates a new custom content type in the specified SharePoint site using the Microsoft Graph API.
    /// </summary>
    /// <remarks>The created content type is based on the 'Document2' content type template. Ensure that the
    /// application has the necessary permissions to create content types in the target SharePoint site.</remarks>
    /// <param name="contentName">The display name of the custom content type to create. Cannot be null.</param>
    /// <param name="contentDescription">The description of the custom content type. This value provides additional information about the content type's
    /// purpose.</param>
    /// <param name="contentCategory">The group or category under which the content type will be organized. Cannot be null.</param>
    /// <returns>A task that represents the asynchronous operation. The task result contains the created ContentType object.</returns>
    /// <exception cref="NullReferenceException">Thrown if the Microsoft Graph client has not been initialized for app-only authentication.</exception>
    /// <exception cref="InvalidOperationException">Thrown if the Microsoft Graph API does not return a valid ContentType object.</exception>
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

    /// <summary>
    /// Creates a new document set folder in the configured SharePoint drive with the specified name.
    /// </summary>
    /// <remarks>If a folder with the specified name already exists at the target location, the operation will
    /// fail due to conflict behavior settings.</remarks>
    /// <param name="documentSetName">The name of the document set folder to create. Cannot be null or empty.</param>
    /// <returns>A task that represents the asynchronous operation. The task result contains the created DriveItem representing
    /// the new document set folder.</returns>
    /// <exception cref="InvalidOperationException">Thrown if the Microsoft Graph API does not return a valid DriveItem for the created folder.</exception>
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

    /// <summary>
    /// Asynchronously updates the name of a list item in the specified SharePoint drive.
    /// </summary>
    /// <param name="documentid">The unique identifier of the list item to update. Cannot be null.</param>
    /// <param name="documentNameNew">The new name to assign to the list item. Cannot be null.</param>
    /// <returns>A task that represents the asynchronous update operation.</returns>
    /// <exception cref="NullReferenceException">Thrown if the Graph client has not been initialized for app-only authentication.</exception>
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

    /// <summary>
    /// Asynchronously updates a specified field value for a list item within a SharePoint document set.
    /// </summary>
    /// <param name="documentSetId">The content type ID of the document set containing the list item to update. Cannot be null or empty.</param>
    /// <param name="itemId">The unique identifier of the list item whose field will be updated. Cannot be null or empty.</param>
    /// <param name="fieldName">The name of the field to update on the list item. Cannot be null or empty.</param>
    /// <param name="fieldValue">The new value to assign to the specified field. The value must be compatible with the field's data type.</param>
    /// <returns>A task that represents the asynchronous operation. The task result contains the updated list item. Throws an
    /// exception if the update fails.</returns>
    /// <exception cref="NullReferenceException">Thrown if the Microsoft Graph client has not been initialized for app-only authentication.</exception>
    /// <exception cref="InvalidOperationException">Thrown if the Microsoft Graph API returns an error or fails to update the list item.</exception>
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

    /// <summary>
    /// Updates the metadata fields of a document set item in the SharePoint list asynchronously.
    /// </summary>
    /// <param name="itemId">The unique identifier of the list item representing the document set to update. Cannot be null or empty.</param>
    /// <param name="documentSetContentTypeId">The content type ID to assign to the document set. This value is set as the 'ContentTypeId' field of the item.</param>
    /// <param name="fields">A dictionary containing the metadata fields and their values to update for the document set. Keys represent
    /// field names; values are the corresponding field values. Cannot be null.</param>
    /// <returns>A task that represents the asynchronous operation. The task result contains the updated <see cref="ListItem"/>
    /// object for the document set.</returns>
    /// <exception cref="InvalidOperationException">Thrown if the Graph client has not been initialized for app-only authentication, or if the update operation
    /// returns null.</exception>
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