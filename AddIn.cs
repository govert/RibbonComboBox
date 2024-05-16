using ExcelDna.Integration;
using ExcelDna.IntelliSense;

namespace Test18
{
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            IntelliSenseServer.Install();
        }
        public void AutoClose()
        {
            IntelliSenseServer.Uninstall();
        }
    }

    public static class Functions
    {
        [ExcelFunction(Name = "HelloWorld", Description = "Returns 'Hello, World!'")]
        public static string HelloWorld()
        {
            return "Hello, World!";
        }

        [ExcelFunction(Name = "Add", Description = "Adds two numbers")]
        public static double Add(
            [ExcelArgument(Name = "First number", Description = "The first number")]
            double a,
            [ExcelArgument(Name = "Second number", Description = "The second number")]
            double b)
        {
            return a + b;
        }

        // Add a "SayHello" function
        [ExcelFunction(Name = "SayHello", Description = "Returns 'Hello, ' + name")]
        public static string SayHello(
            [ExcelArgument(Name = "Name", Description = "The name to say hello to")]
            string name)
        {
            return "Hello, " + name;
        }
    }
}
