namespace OutlookAddIn.Shared
{
    public class ProjectVM
    {
    public ProjectVM(string id, string name, string description, string url, string state)
    {
        this.id = id;
        this.name = name;
        this.description = description;
        this.url = url;
        this.state = state;
    }

    public string id { get; set; }
    public string name { get; set; }
    public string description { get; set; }
    public string url { get; set; }
    public string state { get; set; }
}

}

