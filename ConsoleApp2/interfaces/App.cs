public class App
{
    private readonly GraphService _graphService;

    public App(GraphService graphService)
    {
        _graphService = graphService;
    }

    public async Task RunAsync()
    {
        int choice = -1;

        while (choice != 0)
        {
            PrintMainMenu();
            choice = ReadUserChoice();

            switch (choice)
            {
                case 0:
                    Console.WriteLine("Goodbye...");
                    break;

                case 1:
                    await SharePointGeneralMenuAsync();
                    break;

                case 2:
                    await DocumentLibraryMenuAsync();
                    break;

                default:
                    Console.WriteLine("Invalid choice!");
                    break;
            }
        }
    }

    private void PrintMainMenu()
    {
        Console.WriteLine("\n=== MAIN MENU ===");
        Console.WriteLine("0. Exit");
        Console.WriteLine("1. View SPTraining SharePoint Site General Info");
        Console.WriteLine("2. Perform Operations in Documents Library");
        Console.Write("\nSelect an option: ");
    }

    private async Task SharePointGeneralMenuAsync()
    {
        int choice = -1;

        while (choice != 0)
        {
            Console.WriteLine("\n=== SHAREPOINT GENERAL MENU ===");
            Console.WriteLine("0. Back");
            Console.WriteLine("1. Display Access Token");
            Console.WriteLine("2. List Users");
            Console.WriteLine("3. List SPTraining Site Info");
            Console.WriteLine("4. List All SharePoint Lists");
            Console.WriteLine("5. List All Drive Libraries");
            Console.Write("\nSelect an option: ");

            choice = ReadUserChoice();

            switch (choice)
            {
                case 0: break;
                case 1: await _graphService.DisplayAccessTokenAsync(); break;
                case 2: await _graphService.ListUsersAsync(); break;
                case 3: await _graphService.ListSharepointRootSite(); break;
                case 4: await _graphService.ListSharepointListSites(); break;
                case 5: await _graphService.ListSharepointDriveSites(); break;

                default:
                    Console.WriteLine("Invalid option!");
                    break;
            }
        }
    }

    private async Task DocumentLibraryMenuAsync()
    {
        int choice = -1;

        while (choice != 0)
        {
            Console.WriteLine("\n=== DOCUMENT LIBRARY MENU ===");
            Console.WriteLine("0. Back");
            Console.WriteLine("1. List Files");
            Console.WriteLine("2. List Columns");
            Console.WriteLine("3. Create Columns");
            Console.WriteLine("4. List Content Types");
            Console.WriteLine("5. Create Content Type");
            Console.WriteLine("6. List Documents Info");
            Console.WriteLine("7. Create DocumentSet");
            Console.WriteLine("8. Update DocumentSet Field");
            Console.WriteLine("9. Rename Document");
            Console.Write("\nSelect an option: ");

            choice = ReadUserChoice();

            switch (choice)
            {
                case 0: break;
                case 1: await _graphService.ListSPTrainingDriveFiles(); break;
                case 2: await _graphService.ListSPTrainingFolderColumns(); break;
                case 3: await ColumnCreationMenuAsync(); break;
                case 4: await _graphService.ListSPTrainingFolderContentTypes(); break;
                case 5:  
                    string contentName = ReadRequiredString("Enter the name of the Content Type: ");
                    string contentDescription = ReadRequiredString("Enter a description for the Content Type: ");
                    string contentCategory = ReadRequiredString("Enter the name of the Content Type Category: ");
                    await _graphService.ResultCreatedCustomContentType(contentName, contentDescription, contentCategory); 
                    break;
                case 6: await _graphService.ListSPTrainingFolderItemTypes(); break;
                case 7: 
                    string newFolderName = ReadRequiredString("Enter the name of the new DocumentSet: ");
                    string commonFieldValue = ReadRequiredString("Enter it's default value on the common field: ");
                    await _graphService.ResultCreatedDocumentSet(newFolderName, commonFieldValue);
                    break;
                case 8:
                    string setName = ReadRequiredString("Enter the Document Set name: ");
                    string fieldName = ReadRequiredString("Enter field name: ");
                    string newValue = ReadRequiredString("Enter new value: ");
                    await _graphService.ResultUpdatedSetField(setName, fieldName, newValue);
                    break;

                case 9:
                    string documentNameOld = ReadRequiredString("Enter the current name of the document: ");
                    string documentNameNew = ReadRequiredString("Enter the new name of the document: ");
                    await _graphService.ResultUpdatedItemField(documentNameOld, documentNameNew); break;

                default:
                    Console.WriteLine("Invalid option");
                    break;
            }
        }
    }

    private async Task ColumnCreationMenuAsync()
    {
        int choice = -1;

        while (choice != 0)
        {
            Console.WriteLine("\n=== CREATE COLUMN ===");
            Console.WriteLine("0. Back");
            Console.WriteLine("1. Choice");
            Console.WriteLine("2. Number");
            Console.WriteLine("3. Currency");
            Console.WriteLine("4. DateTime");
            Console.WriteLine("5. Lookup");
            Console.WriteLine("6. Boolean");
            Console.WriteLine("7. Person/Group");
            Console.WriteLine("8. Hyperlink");
            Console.Write("\nSelect an option: ");

            choice = ReadUserChoice();

            switch (choice)
            {
                case 0: break;
                case 1:
                    string choiceColName = ReadRequiredString("Enter the name of the new Choice column: ");
                    string allChoices = ReadRequiredString("Enter the all possible choices separating them with a comma: ");
                    await _graphService.ResultCreatedChoiceColumn(choiceColName, allChoices); break;
                case 2:
                    string numberColName = ReadRequiredString("Enter the name of the new Number column: ");
                    await _graphService.ResultCreatedNumberColumn(numberColName); break;
                case 3:
                    string currencyColName = ReadRequiredString("Enter the name of the new Currency column: ");
                    await _graphService.ResultCreatedCurrencyColumn(currencyColName); break;
                case 4:
                    string dateTimeColName = ReadRequiredString("Enter the name of the new Date and Time column: ");
                    await _graphService.ResultCreatedDateTimeColumn(dateTimeColName); break;
                case 5:
                    string lookUpColName = ReadRequiredString("Enter the name of the new LookUp column: ");
                    await _graphService.ResultCreatedLookUpColumn(lookUpColName); break;
                case 6:
                    string boolColName = ReadRequiredString("Enter the name of the new Boolean column: ");
                    await _graphService.ResultCreatedBooleanColumn(boolColName); break;
                case 7:
                    string pergroupColName = ReadRequiredString("Enter the name of the new Person or Group column: ");
                    await _graphService.ResultCreatedPersonGroupColumn(pergroupColName); break;
                case 8:
                    string hyperlinkColName = ReadRequiredString("Enter the name of the new Hyperlink column: ");
                    await _graphService.ResultCreatedHyperlinkColumn(hyperlinkColName); break;

                default:
                    Console.WriteLine("Invalid option!");
                    break;
            }
        }
    }

    private int ReadUserChoice()
    {
        try
        {
            return int.Parse(Console.ReadLine() ?? string.Empty);
        }
        catch
        {
            return -1;
        }
    }

    private string ReadRequiredString(string prompt)
    {
        string? input;
        do
        {
            Console.Write(prompt);
            input = Console.ReadLine();
        } while (string.IsNullOrWhiteSpace(input));

        return input;
    }
}