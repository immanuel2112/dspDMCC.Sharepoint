namespace dspDMCC.Sharepoint
{
    public class SharepointValue
    {
        public string Id { get; set; }

        public string Name { get; set; }

        public string SiteURL { get; set; }

        public string SiteID { get; set; }

        public string ClientID { get; set; }

        public string ClientSecret { get; set; }

        public string FolderURL { get; set; }

        public string BaseFolderName { get; set; }

        public override string ToString()
        {
            return "Id: " + Id +
                ", Name: " + Name +
                ", SiteURL: " + SiteURL +
                ", SiteID: " + SiteID +
                ", ClientID: " + ClientID +
                ", ClientSecret: " + ClientSecret +
                ", FolderURL: " + FolderURL +
                ", BaseFolderName: " + BaseFolderName;
        }
    }
}
