using ConsoleApp2.config;
using ConsoleApp2.helpers;
using ConsoleApp2.utils;

/// <summary>
/// Provides methods for interacting with Microsoft Graph resources, including user, SharePoint site, list, drive, and
/// document set management using app-only authentication.
/// </summary>
/// <remarks>The GraphService class encapsulates common operations for working with Microsoft Graph in an
/// application context, such as listing users, accessing SharePoint sites and lists, managing files and folders, and
/// creating or updating columns and content types. All operations are performed using app-only authentication, and
/// results are typically written to the console. This class is intended for scenarios where automated or administrative
/// access to Microsoft 365 resources is required. Thread safety is not guaranteed; create separate instances if used
/// concurrently.</remarks>
public class GraphService
{
    /// <summary>
    /// Initializes a new instance of the GraphService class using the specified settings.
    /// </summary>
    /// <param name="settings">The configuration settings used to initialize the graph service. Cannot be null.</param>
    public GraphService(Settings settings)
    {
        InitializeGraph(settings);
    }

    /// <summary>
    /// Initializes the Microsoft Graph client for application-only authentication using the specified settings.
    /// </summary>
    /// <param name="settings">The settings to use for configuring application-only authentication. Cannot be null.</param>
    private void InitializeGraph(Settings settings)
    {
        GraphHelper.InitializeGraphForAppOnlyAuth(settings);
    }

    /// <summary>
    /// Asynchronously retrieves and displays the app-only access token in the console output.
    /// </summary>
    /// <remarks>If an error occurs while retrieving the access token, the exception message is written to the
    /// console output.</remarks>
    /// <returns>A task that represents the asynchronous operation.</returns>
    public async Task DisplayAccessTokenAsync()
    {
        try
        {
            var token = await GraphHelper.GetAppOnlyTokenAsync();
            Console.WriteLine($"App-only token: {token}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error getting app-only access token: {ex.Message}");
        }
    }

    /// <summary>
    /// Asynchronously retrieves a list of users and writes their details to the console output.
    /// </summary>
    /// <remarks>This method outputs each user's display name, ID, and email address to the console. If no
    /// users are found, a message is displayed. The method also indicates whether additional users are available for
    /// retrieval. This method is intended for interactive or diagnostic scenarios where console output is
    /// appropriate.</remarks>
    /// <returns>A task that represents the asynchronous operation.</returns>
    public async Task ListUsersAsync()
    {
        try
        {
            var userPage = await GraphHelper.GetUsersAsync();

            if (userPage?.Value == null)
            {
                Console.WriteLine("No results returned.");
                return;
            }

            foreach (var user in userPage.Value)
            {
                Console.WriteLine($"User: {user.DisplayName ?? "NO NAME"}");
                Console.WriteLine($"  ID: {user.Id}");
                Console.WriteLine($"  Email: {user.Mail ?? "NO EMAIL"}");
            }

            bool moreAvailable = !string.IsNullOrEmpty(userPage.OdataNextLink);
            Console.WriteLine($"\nMore users available? {moreAvailable}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error getting users: {ex.Message}");
        }
    }

    /// <summary>
    /// Retrieves and displays information about the root SharePoint site using Microsoft Graph.
    /// </summary>
    /// <remarks>This method writes the root site ID and web URL to the console if found. If no site is found
    /// or an error occurs, a message is written to the console. This method is intended for interactive or diagnostic
    /// scenarios where console output is appropriate.</remarks>
    /// <returns>A task that represents the asynchronous operation.</returns>
    public async Task ListSharepointRootSite()
    {
        try
        {
            var site = await GraphHelper.GetSitesCallAsync();

            if (site != null)
            {
                Console.WriteLine("-------------------------");
                Console.WriteLine($"ID: {site.Id}");
                Console.WriteLine($"WebUrl: {site.WebUrl}");
                Console.WriteLine("-------------------------");
            }
            else
            {
                Console.WriteLine("No matching site found.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error getting Root Site: {ex.Message}");
        }
    }

    /// <summary>
    /// Retrieves and displays a list of SharePoint sites and their associated lists to the console asynchronously.
    /// </summary>
    /// <remarks>This method writes the details of each SharePoint list, including its ID, name, and web URL,
    /// to the console output. If no lists are found, a message is displayed indicating that no lists are available.
    /// Errors encountered during the operation are written to the console.</remarks>
    /// <returns>A task that represents the asynchronous operation.</returns>
    public async Task ListSharepointListSites()
    {
        try
        {
            var lists = await GraphHelper.GetListsCallAsync();

            if (lists != null && lists.Count > 0)
            {
                foreach (var spList in lists)
                {
                    Console.WriteLine("-------------------------");
                    Console.WriteLine($"ID: {spList.Id}");
                    Console.WriteLine($"Name: {spList.Name}");
                    Console.WriteLine($"WebUrl: {spList.WebUrl}");
                    Console.WriteLine("-------------------------");
                }
            }
            else
            {
                Console.WriteLine("No lists found.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error getting lists: {ex.Message}");
        }
    }

    /// <summary>
    /// Retrieves and displays a list of SharePoint drive sites to the console.
    /// </summary>
    /// <remarks>This method writes the details of each SharePoint drive site, including its ID, name, and web
    /// URL, to the standard output. If no drives are found, a message is displayed indicating that no drives were
    /// found. Errors encountered during retrieval are written to the console.</remarks>
    /// <returns>A task that represents the asynchronous operation.</returns>
    public async Task ListSharepointDriveSites()
    {
        try
        {
            var drives = await GraphHelper.GetDrivesCallAsync();

            if (drives != null && drives.Count > 0)
            {
                foreach (var spDrives in drives)
                {
                    Console.WriteLine("-------------------------");
                    Console.WriteLine($"ID: {spDrives.Id}");
                    Console.WriteLine($"Name: {spDrives.Name}");
                    Console.WriteLine($"WebUrl: {spDrives.WebUrl}");
                    Console.WriteLine("-------------------------");
                }
            }
            else
            {
                Console.WriteLine("No Drives found.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error getting lists: {ex.Message}");
        }
    }

    /// <summary>
    /// Retrieves and displays a list of files from the SharePoint training drive to the console.
    /// </summary>
    /// <remarks>This method writes file details, including ID, name, URL, size, and type, to the standard
    /// output. If no files are found, a message is displayed. Errors encountered during retrieval are written to the
    /// console.</remarks>
    /// <returns>A task that represents the asynchronous operation.</returns>
    public async Task ListSPTrainingDriveFiles()
    {
        try
        {
            var files = await GraphHelper.GetFilesInDriveAsync();

            if (files != null && files.Count > 0)
            {
                foreach (var spFiles in files)
                {
                    Console.WriteLine("-------------------------");
                    Console.WriteLine($"ID: {spFiles.Id}");
                    Console.WriteLine($"Name: {spFiles.Name}");
                    Console.WriteLine($"WebUrl: {spFiles.WebUrl}");
                    Console.WriteLine($"Size: {spFiles.Size}");
                    Console.WriteLine($"File: {spFiles.File}");
                    Console.WriteLine($"Folder: {spFiles.Folder}");
                    Console.WriteLine("-------------------------");
                }
            }
            else
            {
                Console.WriteLine("No Files found.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error getting lists: {ex.Message}");
        }
    }

    /// <summary>
    /// Asynchronously retrieves and displays the columns of the SharePoint training folder in the console.
    /// </summary>
    /// <remarks>This method writes the display names of the columns to the standard output. If no columns are
    /// found, a message is displayed instead. Errors encountered during retrieval are also written to the console. This
    /// method is intended for diagnostic or informational purposes and does not return data to the caller.</remarks>
    /// <returns></returns>
    public async Task ListSPTrainingFolderColumns()
    {
        try
        {
            var columns = await GraphHelper.GetDocsFolderColumns();

            if (columns != null && columns.Count > 0)
            {
                foreach (var spColumns in columns)
                {
                    Console.WriteLine("-------------------------");
                    Console.WriteLine($"Display Name: {spColumns.DisplayName}");
                    Console.WriteLine("-------------------------");
                }
            }
            else
            {
                Console.WriteLine("No columns found.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error getting lists: {ex.Message}");
        }
    }

    /// <summary>
    /// Lists the content types available in the SharePoint training folder and writes their details to the console.
    /// </summary>
    /// <remarks>This method retrieves content types from the SharePoint training folder using the Graph API
    /// and outputs their name, group, and description to the console. If no content types are found, a message is
    /// displayed. Errors encountered during the operation are written to the console.</remarks>
    /// <returns>A task that represents the asynchronous operation.</returns>
    public async Task ListSPTrainingFolderContentTypes()
    {
        try
        {
            var contentTypes = await GraphHelper.GetDocsFolderContentTypes();

            if (contentTypes != null && contentTypes.Count > 0)
            {
                foreach (var spContentTypes in contentTypes)
                {
                    Console.WriteLine("-------------------------");
                    Console.WriteLine($"Name: {spContentTypes.Name}");
                    Console.WriteLine($"Group: {spContentTypes.Group}");
                    Console.WriteLine($"Description: {spContentTypes.Description}");
                    Console.WriteLine("-------------------------");
                }
            }
            else
            {
                Console.WriteLine("No Content Types found.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error getting lists: {ex.Message}");
        }
    }

    /// <summary>
    /// Retrieves and displays information about SharePoint training folder item types, including their content type
    /// names and associated fields.
    /// </summary>
    /// <remarks>This method writes details about each item type in the SharePoint training folder to the
    /// console. It is intended for diagnostic or informational purposes and does not return data to the caller. If no
    /// items are found, a message is displayed to the console. Any errors encountered during retrieval are also written
    /// to the console.</remarks>
    /// <returns></returns>
    public async Task ListSPTrainingFolderItemTypes()
    {
        try
        {
            var contentTypes = await GraphHelper.GetDocsFolderItemInfo();

            if (contentTypes != null && contentTypes.Count > 0)
            {
                foreach (var spContentTypes in contentTypes)
                {
                    string? ctName = spContentTypes.ContentType?.Name;

                    Console.WriteLine("-------------------------");
                    Console.WriteLine($"Id: {spContentTypes.Id}");
                    Console.WriteLine($"WebURL: {spContentTypes.WebUrl}");
                    Console.WriteLine($"Content Type Name: {spContentTypes.ContentType?.Name}");
                    Console.WriteLine("Fields:");
                    var fields = spContentTypes.Fields?.AdditionalData;

                    if (fields is null)
                    {
                        Console.WriteLine("No fields found.");
                    }
                    else
                    {
                        foreach (var kv in fields)
                        {
                            Console.WriteLine($"    {kv.Key}: {kv.Value}");
                        }
                    }

                    Console.WriteLine("-------------------------");
                }
            }
            else
            {
                Console.WriteLine("No Items found.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error getting lists: {ex.Message}");
        }
    }

    /// <summary>
    /// Creates a new choice column with the specified name and set of choices using the Microsoft Graph API.
    /// </summary>
    /// <remarks>If the column is created successfully, details about the new column and its choices are
    /// written to the console. If the operation fails, an error message is displayed. This method does not throw
    /// exceptions; errors are reported via console output.</remarks>
    /// <param name="choiceColName">The name of the choice column to create. Cannot be null or empty.</param>
    /// <param name="allChoices">A comma-separated list of choices to include in the column. Each choice will be trimmed of whitespace. Cannot be
    /// null or empty.</param>
    /// <returns>A task that represents the asynchronous operation.</returns>
    public async Task ResultCreatedChoiceColumn(string choiceColName, string allChoices)
    {
        try
        {
            var choiceList = allChoices
               .Split(',', StringSplitOptions.RemoveEmptyEntries)
               .Select(c => c.Trim())
               .ToList();

            var column = await GraphHelper.CreateChoiceColumnAsync(choiceColName, choiceList);

            if (column != null)
            {
                Console.WriteLine("-------------------------");
                Console.WriteLine("Choice Column Created!");
                Console.WriteLine($"Id: {column.Id}");
                Console.WriteLine($"Name: {column.Name}");
                Console.WriteLine($"Description: {column.Description}");
                Console.WriteLine($"Hidden: {column.Hidden}");
                Console.WriteLine($"Indexed: {column.Indexed}");
                Console.WriteLine($"EnforceUniqueValues: {column.EnforceUniqueValues}");

                if (column.Choice != null)
                {
                    Console.WriteLine("Choice options:");
                    Console.WriteLine($"    AllowTextEntry: {column.Choice.AllowTextEntry}");
                    Console.WriteLine($"    DisplayAs: {column.Choice.DisplayAs}");
                    if (column.Choice.Choices != null)
                    {
                        foreach (var choice in column.Choice.Choices)
                        {
                            Console.WriteLine($"    Choice: {choice}");
                        }
                    }
                }
                Console.WriteLine("-------------------------");
            }
            else
            {
                Console.WriteLine("Column was not created.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error creating column: {ex.Message}");
        }
    }

    /// <summary>
    /// Creates a new number column with the specified name and displays its details to the console.
    /// </summary>
    /// <remarks>If the column is successfully created, detailed information about the column is written to
    /// the console. If the creation fails or an error occurs, an appropriate message is displayed. This method is
    /// intended for interactive or diagnostic scenarios where console output is appropriate.</remarks>
    /// <param name="numberColName">The name to assign to the newly created number column. Cannot be null or empty.</param>
    /// <returns>A task that represents the asynchronous operation.</returns>
    public async Task ResultCreatedNumberColumn(string numberColName)
    {
        try
        {
            var column = await GraphHelper.CreateNumberColumnAsync(numberColName);

            if (column != null)
            {
                Console.WriteLine("-------------------------");
                Console.WriteLine("Number Column Created!");
                Console.WriteLine($"Id: {column.Id}");
                Console.WriteLine($"Name: {column.Name}");
                Console.WriteLine($"Description: {column.Description}");
                Console.WriteLine($"Hidden: {column.Hidden}");
                Console.WriteLine($"Indexed: {column.Indexed}");
                Console.WriteLine($"EnforceUniqueValues: {column.EnforceUniqueValues}");

                if (column.Number != null)
                {
                    Console.WriteLine("Choice options:");
                    Console.WriteLine($"    DecimalPlaces: {column.Number.DecimalPlaces}");
                    Console.WriteLine($"    DisplayAs: {column.Number.DisplayAs}");
                    Console.WriteLine($"    Maximum: {column.Number.Maximum}");
                    Console.WriteLine($"    Minimum: {column.Number.Minimum}");
                }
                Console.WriteLine("-------------------------");
            }
            else
            {
                Console.WriteLine("Column was not created.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error creating column: {ex.Message}");
        }
    }

    /// <summary>
    /// Creates a new currency column with the specified name and displays its details to the console.
    /// </summary>
    /// <remarks>If the column is successfully created, its details are written to the console. If creation
    /// fails or an error occurs, an error message is displayed. This method is intended for interactive or diagnostic
    /// scenarios where console output is appropriate.</remarks>
    /// <param name="currencyColName">The name of the currency column to create. Cannot be null or empty.</param>
    /// <returns>A task that represents the asynchronous operation.</returns>
    public async Task ResultCreatedCurrencyColumn(string currencyColName)
    {
        try
        {
            var column = await GraphHelper.CreateCurrencyColumnAsync(currencyColName);

            if (column != null)
            {
                Console.WriteLine("-------------------------");
                Console.WriteLine("Currency Column Created!");
                Console.WriteLine($"Id: {column.Id}");
                Console.WriteLine($"Name: {column.Name}");
                Console.WriteLine($"Description: {column.Description}");
                Console.WriteLine($"Hidden: {column.Hidden}");
                Console.WriteLine($"Indexed: {column.Indexed}");
                Console.WriteLine($"EnforceUniqueValues: {column.EnforceUniqueValues}");

                if (column.Currency != null)
                {
                    Console.WriteLine("Currency options:");
                    Console.WriteLine($"    Locale: {column.Currency.Locale}");
                }
                Console.WriteLine("-------------------------");
            }
            else
            {
                Console.WriteLine("Column was not created.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error creating column: {ex.Message}");
        }
    }

    /// <summary>
    /// Creates a new DateTime column with the specified name and outputs its details to the console.
    /// </summary>
    /// <remarks>If the column is successfully created, its properties are written to the console. If creation
    /// fails or the column is not created, a message is displayed instead.</remarks>
    /// <param name="dateTimeColName">The name of the DateTime column to create. Cannot be null or empty.</param>
    /// <returns>A task that represents the asynchronous operation.</returns>
    public async Task ResultCreatedDateTimeColumn(string dateTimeColName)
    {
        try
        {
            var column = await GraphHelper.CreateDateTimeColumnAsync(dateTimeColName);

            if (column != null)
            {
                Console.WriteLine("-------------------------");
                Console.WriteLine("DateTime Column Created!");
                Console.WriteLine($"Id: {column.Id}");
                Console.WriteLine($"Name: {column.Name}");
                Console.WriteLine($"Description: {column.Description}");
                Console.WriteLine($"Hidden: {column.Hidden}");
                Console.WriteLine($"Indexed: {column.Indexed}");
                Console.WriteLine($"EnforceUniqueValues: {column.EnforceUniqueValues}");

                if (column.DateTime != null)
                {
                    Console.WriteLine("DateTime options:");
                    Console.WriteLine($"    DisplayAs: {column.DateTime.DisplayAs}");
                    Console.WriteLine($"    Format: {column.DateTime.Format}");
                }
                Console.WriteLine("-------------------------");
            }
            else
            {
                Console.WriteLine("Column was not created.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error creating column: {ex.Message}");
        }
    }

    /// <summary>
    /// Creates a lookup column with the specified name and displays its details to the console.
    /// </summary>
    /// <remarks>If the lookup column is created successfully, its properties and lookup options are written
    /// to the console. If the creation fails or an error occurs, an error message is displayed.</remarks>
    /// <param name="lookUpColName">The name of the lookup column to create. Cannot be null or empty.</param>
    /// <returns>A task that represents the asynchronous operation.</returns>
    public async Task ResultCreatedLookUpColumn(string lookUpColName)
    {
        try
        {
            var column = await GraphHelper.CreateLookUpColumnAsync(lookUpColName);

            if (column != null)
            {
                Console.WriteLine("-------------------------");
                Console.WriteLine("LookUp Column Created!");
                Console.WriteLine($"Id: {column.Id}");
                Console.WriteLine($"Name: {column.Name}");
                Console.WriteLine($"Description: {column.Description}");
                Console.WriteLine($"Hidden: {column.Hidden}");
                Console.WriteLine($"Indexed: {column.Indexed}");
                Console.WriteLine($"EnforceUniqueValues: {column.EnforceUniqueValues}");

                if (column.Lookup != null)
                {
                    Console.WriteLine("LookUp options:");
                    Console.WriteLine($"    AllowMultipleValues: {column.Lookup.AllowMultipleValues}");
                    Console.WriteLine($"    AllowUnlimitedLenght: {column.Lookup.AllowUnlimitedLength}");
                    Console.WriteLine($"    ColumnName: {column.Lookup.ColumnName}");
                    Console.WriteLine($"    ListID: {column.Lookup.ListId}");
                }
                Console.WriteLine("-------------------------");
            }
            else
            {
                Console.WriteLine("Column was not created.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error creating column: {ex.Message}");
        }
    }

    /// <summary>
    /// Creates a new boolean column with the specified name and outputs its details to the console.
    /// </summary>
    /// <remarks>If the column is successfully created, its properties are written to the console. If creation
    /// fails or an error occurs, an appropriate message is displayed. This method is intended for interactive or
    /// diagnostic scenarios where console output is appropriate.</remarks>
    /// <param name="boolColName">The name of the boolean column to create. Cannot be null or empty.</param>
    /// <returns>A task that represents the asynchronous operation.</returns>
    public async Task ResultCreatedBooleanColumn(string boolColName)
    {
        try
        {
            var column = await GraphHelper.CreateBooleanColumnAsync(boolColName);

            if (column != null)
            {
                Console.WriteLine("-------------------------");
                Console.WriteLine("Boolean Column Created!");
                Console.WriteLine($"Id: {column.Id}");
                Console.WriteLine($"Name: {column.Name}");
                Console.WriteLine($"Description: {column.Description}");
                Console.WriteLine($"Hidden: {column.Hidden}");
                Console.WriteLine($"Indexed: {column.Indexed}");
                Console.WriteLine($"EnforceUniqueValues: {column.EnforceUniqueValues}");

                if (column.Boolean != null)
                {
                    Console.WriteLine("Boolean options:");
                    Console.WriteLine("     No options on booleans.");
                }
                Console.WriteLine("-------------------------");
            }
            else
            {
                Console.WriteLine("Column was not created.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error creating column: {ex.Message}");
        }
    }

    /// <summary>
    /// Creates a new Person or Group column with the specified name and displays its details to the console.
    /// </summary>
    /// <remarks>If the column is successfully created, detailed information about the column and its options
    /// is written to the console. If creation fails or the column is not created, an error message is
    /// displayed.</remarks>
    /// <param name="pergroupColName">The name of the Person or Group column to create. Cannot be null or empty.</param>
    /// <returns>A task that represents the asynchronous operation.</returns>
    public async Task ResultCreatedPersonGroupColumn(string pergroupColName)
    {
        try
        {
            var column = await GraphHelper.CreatePersonGroupColumnAsync(pergroupColName);

            if (column != null)
            {
                Console.WriteLine("-------------------------");
                Console.WriteLine("PersonOrGroup Column Created!");
                Console.WriteLine($"Id: {column.Id}");
                Console.WriteLine($"Name: {column.Name}");
                Console.WriteLine($"Description: {column.Description}");
                Console.WriteLine($"Hidden: {column.Hidden}");
                Console.WriteLine($"Indexed: {column.Indexed}");
                Console.WriteLine($"EnforceUniqueValues: {column.EnforceUniqueValues}");

                if (column.PersonOrGroup != null)
                {
                    Console.WriteLine("PersonOrGroup options:");
                    Console.WriteLine($"    AllowMultipleSelection: {column.PersonOrGroup.AllowMultipleSelection}");
                    Console.WriteLine($"    DisplayAs: {column.PersonOrGroup.DisplayAs}");
                    Console.WriteLine($"    ChooseFromType: {column.PersonOrGroup.ChooseFromType}");
                }
                Console.WriteLine("-------------------------");
            }
            else
            {
                Console.WriteLine("Column was not created.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error creating column: {ex.Message}");
        }
    }

    /// <summary>
    /// Creates a new hyperlink column with the specified name and displays its details to the console.
    /// </summary>
    /// <remarks>This method writes information about the created column to the console, including its
    /// properties and hyperlink options. If the column creation fails, an error message is displayed.</remarks>
    /// <param name="hyperlinkColName">The name to assign to the new hyperlink column. Cannot be null or empty.</param>
    /// <returns>A task that represents the asynchronous operation.</returns>
    public async Task ResultCreatedHyperlinkColumn(string hyperlinkColName)
    {
        try
        {
            var column = await GraphHelper.CreateHyperlinkColumnAsync(hyperlinkColName);

            if (column != null)
            {
                Console.WriteLine("-------------------------");
                Console.WriteLine("Hyperlink Column Created!");
                Console.WriteLine($"Id: {column.Id}");
                Console.WriteLine($"Name: {column.Name}");
                Console.WriteLine($"Description: {column.Description}");
                Console.WriteLine($"Hidden: {column.Hidden}");
                Console.WriteLine($"Indexed: {column.Indexed}");
                Console.WriteLine($"EnforceUniqueValues: {column.EnforceUniqueValues}");

                if (column.HyperlinkOrPicture != null)
                {
                    Console.WriteLine("Hyperlink options:");
                    Console.WriteLine($"    IsPicture: {column.HyperlinkOrPicture.IsPicture}");
                }
                Console.WriteLine("-------------------------");
            }
            else
            {
                Console.WriteLine("Column was not created.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error creating column: {ex.Message}");
        }
    }

    /// <summary>
    /// Creates a new custom content type with the specified name, description, and category.
    /// </summary>
    /// <remarks>This method outputs information about the created content type to the console. If the content
    /// type cannot be created, an error message is displayed. This method does not return the created content type
    /// object.</remarks>
    /// <param name="contentName">The name of the custom content type to create. Cannot be null or empty.</param>
    /// <param name="contentDescription">The description of the custom content type. This value provides additional details about the content type's
    /// purpose.</param>
    /// <param name="contentCategory">The category or group to which the custom content type belongs. Used to organize content types within the
    /// system.</param>
    /// <returns>A task that represents the asynchronous operation.</returns>
    public async Task ResultCreatedCustomContentType(string contentName, string contentDescription, string contentCategory)
    {
        try
        {
            var contentType = await GraphHelper.CreateCustomContentTypeAsync(contentName, contentDescription, contentCategory);

            if (contentType != null)
            {
                Console.WriteLine("-------------------------");
                Console.WriteLine("Custom Content Type Created!");
                Console.WriteLine($"Name: {contentType.Name}");
                Console.WriteLine($"Description: {contentType.Description}");
                Console.WriteLine($"Group: {contentType.Group}");

                if (contentType.Base != null)
                {
                    Console.WriteLine("Content Type options:");
                    Console.WriteLine($"    BaseId: {contentType.Base.Id}");
                    Console.WriteLine($"    Name: {contentType.Base.Name}");
                }
                Console.WriteLine("-------------------------");
            }
            else
            {
                Console.WriteLine("Content Type was not created.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error creating Content Type: {ex.Message}");
        }
    }

    /// <summary>
    /// Updates the name of a list item field in the drive by renaming the specified document.
    /// </summary>
    /// <remarks>If the specified document is not found, the method completes without making changes. This
    /// method writes status and error messages to the console during execution.</remarks>
    /// <param name="documentNameOld">The current name of the document to be updated. Cannot be null or empty.</param>
    /// <param name="documentNameNew">The new name to assign to the document. Cannot be null or empty.</param>
    /// <returns>A task that represents the asynchronous operation.</returns>
    /// <exception cref="InvalidOperationException">Thrown if the drive item does not have a valid identifier.</exception>
    public async Task ResultUpdatedItemField(
        string documentNameOld,
        string documentNameNew)
    {
        try
        {
            var files = await GraphHelper.GetFilesInDriveAsync();

            if (files is null)
            {
                Console.WriteLine("No files returned from Graph.");
                return;
            }

            var file = files.FirstOrDefault(f => f.Name == documentNameOld);

            if (file is null)
            {
                Console.WriteLine("No File with that name found.");
                return;
            }

            var fileId = file.Id
                ?? throw new InvalidOperationException("Drive item Id is null.");

            await GraphHelper.UpdateListItemFieldAsync(fileId, documentNameNew);

            Console.WriteLine("-------------------------");
            Console.WriteLine("File renamed successfully!");
            Console.WriteLine("-------------------------");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error getting lists: {ex.Message}");
        }
    }

    /// <summary>
    /// Updates the value of a specified field in a document set identified by name.
    /// </summary>
    /// <param name="setName">The name of the document set whose field value will be updated. Cannot be null.</param>
    /// <param name="fieldName">The name of the field within the document set to update. Cannot be null.</param>
    /// <param name="newValue">The new value to assign to the specified field. May be null or empty depending on the field's requirements.</param>
    /// <returns>A task that represents the asynchronous operation.</returns>
    /// <exception cref="InvalidOperationException">Thrown if the document set's content type ID or document set ID is null.</exception>
    public async Task ResultUpdatedSetField(
        string setName,
        string fieldName,
        string newValue)
    {
        try
        {
            var documentSets = await GraphHelper.GetDocsFolderItemInfo();

            if (documentSets is null)
            {
                Console.WriteLine("No document sets returned.");
                return;
            }

            var documentSet = documentSets.FirstOrDefault(f =>
            f.Fields?.AdditionalData is { } data &&
            data.TryGetValue("LinkFilename", out var value) &&
            value?.ToString() == setName);

            if (documentSet is null)
            {
                Console.WriteLine("No Document Sets found with that name.");
                return;
            }

            var contentTypeId = documentSet.ContentType?.Id
                ?? throw new InvalidOperationException("ContentType Id is null.");

            var documentSetId = documentSet.Id
                ?? throw new InvalidOperationException("DocumentSet Id is null.");

            await GraphHelper.UpdateDocumentSetFieldAsync(
                contentTypeId,
                documentSetId,
                fieldName,
                newValue);

            Console.WriteLine("-------------------------");
            Console.WriteLine("Document Set Updated successfully!");
            Console.WriteLine("-------------------------");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error creating Content Type: {ex.Message}");
        }
    }

    /// <summary>
    /// Creates a new Document Set with the specified name and updates its metadata with the provided field value.
    /// </summary>
    /// <remarks>If the Document Set folder cannot be created or its information cannot be retrieved, the
    /// method writes an error message to the console and returns without throwing an exception.</remarks>
    /// <param name="documentSetName">The name of the Document Set to create. This value is used as the title and to identify the folder.</param>
    /// <param name="commonFieldValue">The value to assign to the 'test' field in the Document Set's metadata.</param>
    /// <returns>A task that represents the asynchronous operation.</returns>
    /// <exception cref="InvalidOperationException">Thrown if the created Document Set folder item does not have a valid identifier.</exception>
    public async Task ResultCreatedDocumentSet(
        string documentSetName,
        string commonFieldValue)
    {
        var fields = new Dictionary<string, object?>
    {
        { "Title", documentSetName },
        { "test", commonFieldValue }
    };

        var folder = await GraphHelper.CreateDocumentSetFolderAsync(documentSetName);

        if (folder is null)
        {
            Console.WriteLine("Failed to create Document Set folder.");
            return;
        }

        var folderInfo = await GraphHelper.GetDocsFolderItemInfo();

        if (folderInfo is null)
        {
            Console.WriteLine("Could not retrieve folder information.");
            return;
        }

        var folderItem = folderInfo.FirstOrDefault(f =>
            f.Fields?.AdditionalData is { } data &&
            data.TryGetValue("LinkFilename", out var value) &&
            value?.ToString() == documentSetName);

        if (folderItem is null)
        {
            Console.WriteLine("There was a problem creating the Document Set!");
            return;
        }

        var folderItemId = folderItem.Id
            ?? throw new InvalidOperationException("Folder item Id is null.");

        await GraphHelper.UpdateDocumentSetMetadataAsync(
            folderItemId,
            AppConstants.DocumentSetContentTypeId,
            fields);

        Console.WriteLine("Document Set created!");
    }
}