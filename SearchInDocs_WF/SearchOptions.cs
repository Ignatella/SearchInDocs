namespace SearchInDocs_WF
{
    internal class SearchOptions
    {
        public string StrToSearchFor { get; }
        public string Path { get; }
        public SearchOptions(string strTosearchFor, string path)
        {
            StrToSearchFor = strTosearchFor;
            Path = path;
        }
    }
    
}
