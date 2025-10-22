using Azure.Identity;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Sites.Item.Lists.Item.Items;
using Microsoft.Kiota.Abstractions;


namespace Basys.EntApp.Core.SharePoint
{
    public class SharePointService : ISharePointService
    {
        private GraphServiceClient _graphServiceClient;
        private string _siteId;

        public SharePointService(SharepointSiteConfiguration config, string tenantId)
        {
            var devSpConfig = config;

            if (devSpConfig != null)
                this.Initialize(tenantId, devSpConfig);
        }

        public SharePointService(SharepointConfiguration config)
        {
            var devSpConfig = config.Sites.Where(p => p.Name == "TechSite").FirstOrDefault();

            if (devSpConfig != null)
                this.Initialize(config.TenantId, devSpConfig);
        }

        public SharePointService(IOptions<AzureAdConfiguration> azureAdConfig, IOptions<SharepointConfiguration> sharepointConfig)
        {

            if (azureAdConfig != null && sharepointConfig != null)
            {
                Console.WriteLine("azureAdConfig and sharepointConfig not null ");

                var config = sharepointConfig.Value;
                var spConfig = config.Sites?.FirstOrDefault();

                if (spConfig == null)
                {
                    Console.WriteLine($"Error config.Sites null, defaulting to empty");
                    spConfig = new SharepointSiteConfiguration();
                }

                Console.WriteLine($"azureAdConfig.Value.ClientId'{azureAdConfig.Value.ClientId}'  azureAdConfig.Value.TenantId ='{azureAdConfig.Value.TenantId}");

                spConfig.ClientId = azureAdConfig.Value.ClientId;
                spConfig.ClientSecret = azureAdConfig.Value.ClientSecret;

                Console.WriteLine($"Attempting to initalize");

                this.Initialize(azureAdConfig.Value.TenantId, spConfig);
            }
            else
            {
                Console.WriteLine("azureAdConfig or sharepointConfig null ");
            }
        }

        private void Initialize(string tenantId, SharepointSiteConfiguration config)
        {
            Console.WriteLine($"Creating ClientSecretCredential");

            var clientSecretCredential = new ClientSecretCredential(tenantId, config.ClientId, config.ClientSecret);
            _siteId = config.SiteId;

            Console.WriteLine($"Using siteid of {_siteId}");

            Console.WriteLine($"Creating GraphServiceClient");

            _graphServiceClient = new GraphServiceClient(clientSecretCredential, new[] { "https://graph.microsoft.com/.default" });
        }


        public async Task<DriveItem> UploadFileToSharePointAsync(string documentLibraryName, string folderPath, string fileName, Stream fileStream)
        {
            DriveItem uploadedFile = new DriveItem();

            try
            {
                // Ensure the file stream is at the beginning
                fileStream.Position = 0;

                var drives = await _graphServiceClient.Sites[_siteId].Drives.GetAsync();

                // Find the specified document library
                var drive = drives.Value?.FirstOrDefault(d => d.Name.Equals(documentLibraryName, StringComparison.OrdinalIgnoreCase));

                if (drive != null)
                {
                    // Create the full path for the file upload
                    string fullPath = string.IsNullOrEmpty(folderPath) ? fileName : $"{folderPath}/{fileName}";


                    try
                    {
                        var existingFile = await _graphServiceClient
                            .Drives[drive.Id]
                            .Root
                            .ItemWithPath(fullPath)
                            .GetAsync();

                        if (existingFile != null)
                        {
                            Console.WriteLine($"File '{fileName}' already exists in '{documentLibraryName}/{folderPath}'. Skipping upload.");
                            return existingFile;
                        }
                    }
                    catch (Exception ex)
                    {
                        // If a "not found" exception occurs, it means the file does not exist, and we can proceed with the upload.
                        if (!ex.Message.Contains("404") && !ex.Message.Contains("not be found"))
                        {
                            return null;
                        }
                    }

                    // Upload the file to the specified folder path
                    uploadedFile = await _graphServiceClient
                        .Drives[drive.Id]
                        .Root
                        .ItemWithPath(fullPath)
                        .Content
                        .PutAsync(fileStream);  // Remove the type argument <DriveItem>

                }
                else
                {
                    throw new Exception($"Document library '{documentLibraryName}' not found.");
                }


            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error uploading file to SharePoint: {ex.Message}");
            }

            return uploadedFile;

        }


        public GraphServiceClient GetGraphServiceClient() { return _graphServiceClient; }
        public async Task<List<ListItem>> GetListItemsAsync(
            string listName,
            Action<RequestConfiguration<ItemsRequestBuilder.ItemsRequestBuilderGetQueryParameters>> requestConfiguration = null)
        {
            var items = await _graphServiceClient
            .Sites[_siteId]
            .Lists[listName]
            .Items
            .GetAsync(requestConfiguration);

            return items?.Value ?? new List<ListItem>();
        }

        // Create a dictionary of internal name and display name for the columns
        public async Task<Dictionary<string, string>> GetListColumnsAsync(string listName)
        {
            var columns = await _graphServiceClient
                .Sites[_siteId]
                .Lists[listName]
                .Columns
                .GetAsync();

            var columnDictionary = new Dictionary<string, string>();

            foreach (var column in columns.Value)
            {
                // Add to dictionary with internal name as key and display name as value
                columnDictionary[column.Name] = column.DisplayName;
            }

            return columnDictionary;
        }

        public async Task<DriveItem> GetDriveItemFromSharingLinkAsync(string sharingUrl)
        {
            // Generate the token from the sharingUrl to retrieve it using Shares[encodedUrl]
            string base64Value = System.Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(sharingUrl));
            string encodedUrl = "u!" + base64Value.TrimEnd('=').Replace('/', '_').Replace('+', '-');

            try
            {
                var driveItem = await _graphServiceClient
                    .Shares[encodedUrl]
                    .DriveItem
                    .GetAsync();
                if (driveItem != null)
                {
                    return driveItem;
                }
                else
                {
                    var message = $"Drive item not found for URL: {encodedUrl}";
                    throw new FileNotFoundException(message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error retrieving DriveItem: {ex.Message}");
                throw;
            }
        }
        public async Task<byte[]> GetDriveItemContentAsync(DriveItem driveItem)
        {
            if (driveItem == null || driveItem.Id == null)
                throw new ArgumentNullException(nameof(driveItem), "DriveItem is null or invalid");

            var fileStream = await _graphServiceClient
                .Drives[driveItem.ParentReference.DriveId]
                .Items[driveItem.Id]
                .Content
                .GetAsync();

            // Convert the stream into a byte array
            using (var memoryStream = new MemoryStream())
            {
                await fileStream.CopyToAsync(memoryStream);
                return memoryStream.ToArray();
            }
        }
        public async Task<Stream?> GetFileFromSharePointAsync(string documentLibraryName, string folderPath, string fileName)
        {
            try
            {
                var drives = await _graphServiceClient.Sites[_siteId].Drives.GetAsync();
                var drive = drives?.Value?.FirstOrDefault(d => d.Name.Equals(documentLibraryName, StringComparison.OrdinalIgnoreCase));

                if (drive == null)
                {
                    Console.WriteLine("Drive was not found");
                    return null;
                }

                string fullPath = string.IsNullOrEmpty(folderPath) ? fileName : $"{folderPath}/{fileName}";

                try
                {
                    var file = await _graphServiceClient
                        .Drives[drive.Id]
                        .Root
                        .ItemWithPath(fullPath)
                        .Content
                        .GetAsync();

                    return file;
                }
                catch (ServiceException ex)
                {
                    Console.WriteLine("File does not exist at the specified path.");
                    return null;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Exception type: {ex.GetType()}");
                    Console.WriteLine($"Exception while trying to fetch file: {ex.Message}");
                    return null;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error retrieving file Drive/Drives from SharePoint: {ex.Message}");
                return null;
            }
        }

        public async Task<Stream?> GetFileFromSharePointById(string itemId, string siteId)
        {
            try
            {
                var drives = await _graphServiceClient.Sites[siteId].Drives.GetAsync();
                var drive = drives?.Value?.FirstOrDefault(d => d.Name.Equals("Documents", StringComparison.OrdinalIgnoreCase));

                if (drive == null)
                {
                    Console.WriteLine("Drive was not found");
                    return null;
                }
                var file = await _graphServiceClient
                    .Drives[drive.Id]
                    .Items[itemId]
                    .Content
                    .GetAsync();
                return file;
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"File does not exist with the specified ID: {ex.Message}");
                return null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception while trying to fetch file by ID: {ex.Message}");
                return null;
            }
        }

        public async Task<string> GetFolderPathAsync(string itemId, string siteId)
        {
            try
            {
                var drives = await _graphServiceClient.Sites[siteId].Drives.GetAsync();
                var drive = drives?.Value?.FirstOrDefault(d => d.Name.Equals("Documents", StringComparison.OrdinalIgnoreCase));

                if (drive == null)
                {
                    Console.WriteLine("Drive was not found");
                    return null;
                }
                var file = await _graphServiceClient
                    .Drives[drive.Id]
                    .Items[itemId].GetAsync();
                if (file == null)
                {
                    return null;
                }

                List<string> folderNames = new List<string>();
                var currentItem = file;

                while (currentItem.ParentReference != null)
                {
                    var parentItem = await _graphServiceClient
                        .Drives[drive.Id]
                        .Items[currentItem.ParentReference.Id]
                        .GetAsync();

                    if (parentItem == null || parentItem.Name == "root")
                    {
                        break;
                    }

                    // Add the parent folder's name to the list
                    folderNames.Add(parentItem.Name);

                    // Update the current item to the parent for the next iteration
                    currentItem = parentItem;
                }
                folderNames.Reverse();
                return string.Join("/", folderNames);

            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"File does not exist with the specified ID: {ex.Message}");
                return null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception while trying to fetch file by ID: {ex.Message}");
                return null;
            }
        }
    }
}
