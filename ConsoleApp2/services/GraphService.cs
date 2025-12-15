using ConsoleApp2.config;
using ConsoleApp2.helpers;
using ConsoleApp2.utils;

public class GraphService
{
    public GraphService(Settings settings)
    {
        InitializeGraph(settings);
    }

    private void InitializeGraph(Settings settings)
    {
        GraphHelper.InitializeGraphForAppOnlyAuth(settings);
    }

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

