using ExcelDna.Integration;

namespace Test180
{
    public static class MyFunctions
    {
        [ExcelFunction(Description = "My first .NET function (from 180)")]
        public static string SayHello(string name)
        {
            return "Hello from 180 " + name;
        }
    }
}
