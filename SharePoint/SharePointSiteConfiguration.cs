namespace Basys.EntApp.Core.SharePoint
{
    public class SharepointConfiguration
    {
        public string TenantId { get; set; } = string.Empty;
        public List<SharepointSiteConfiguration> Sites { get; set; }
    }

    public class SharepointSiteConfiguration
    {
        public string Name { get; set; }
        public string ClientId { get; set; } = string.Empty;
        public string ClientSecret { get; set; } = string.Empty;
        public string SiteId { get; set; } = string.Empty;
        public string SiteName { get; set; } = string.Empty;
        public string ListName { get; set; } = string.Empty;
    }
}
