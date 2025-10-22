using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Sites.Item.Lists.Item.Items;
using Microsoft.Kiota.Abstractions;

namespace Basys.EntApp.Core.SharePoint
{
    public interface ISharePointService
    {
        Task<DriveItem> UploadFileToSharePointAsync(string documentLibraryName, string folderPath, string fileName, Stream fileStream);
        GraphServiceClient GetGraphServiceClient();
        Task<List<ListItem>> GetListItemsAsync(
            string listName,
            Action<RequestConfiguration<ItemsRequestBuilder.ItemsRequestBuilderGetQueryParameters>> requestConfiguration = null);
        Task<Dictionary<string, string>> GetListColumnsAsync(string listName);
        Task<DriveItem> GetDriveItemFromSharingLinkAsync(string sharingUrl);
        Task<byte[]> GetDriveItemContentAsync(DriveItem driveItem);

        Task<Stream?> GetFileFromSharePointAsync(string documentLibraryName, string folderPath, string fileName);
        Task<Stream?> GetFileFromSharePointById(string itemId, string siteId);

        Task<string> GetFolderPathAsync(string itemId, string siteId);
    }
}
