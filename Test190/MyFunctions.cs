using ExcelDna.Integration;

namespace Test190
{
    public static class MyFunctions
    {
        [ExcelFunction(Description = "My first .NET function (from 190)")]
        public static string SayHello(string name)
        {
            return "Hello from 190 " + name;
        }
    }
}
