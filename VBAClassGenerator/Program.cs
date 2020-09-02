using System;

namespace VBAClassGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            var prop1 = new VBAClassProperty("totalPayoff", VBADataType.Currency);
            var prop2 = new VBAClassProperty("totalPayoffInterests", VBADataType.Currency);

            var classBuilder = new VBAClassBuilder();
            classBuilder.AddProperty(prop1);
            classBuilder.AddProperty(prop2);




            Console.WriteLine(classBuilder.PrepareClass());
        }
    }
}
