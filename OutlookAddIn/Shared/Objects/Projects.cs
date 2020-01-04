namespace OutlookAddIn.Shared
{
    public class Projects
    {
        public int count { get; set; }
        public Value[] value { get; set; }
    }


    //Helper Class
    public class Value
    {
        public string id { get; set; }
        public string name { get; set; }
        public string description { get; set; }
        public string url { get; set; }
        public string state { get; set; }


    }
}

