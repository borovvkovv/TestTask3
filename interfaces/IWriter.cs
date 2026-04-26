namespace TestTask3.interfaces
{
    internal interface IWriter
    {
        IDictionary<string, string> GetParams();
        void SetParams(IDictionary<string, string> parameters);
    }
}
