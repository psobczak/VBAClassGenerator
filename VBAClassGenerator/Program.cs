using System;

namespace VBAClassGenerator
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            var prop1 = new VBAClassProperty("totalPayoff", VBADataType.Currency);
            var prop2 = new VBAClassProperty("totalPayoffInterests", VBADataType.Currency);
            var prop3 = new VBAClassProperty("schedulePlanNumber", VBADataType.Integer);
            var prop4 = new VBAClassProperty("sendScheduleToClient", VBADataType.Boolean);

            var classBuilder = new VBAClassBuilder();
            classBuilder
                .AddProperty(prop1)
                .AddProperty(prop2)
                .AddProperty(prop3)
                .AddProperty(prop4);

            Console.WriteLine(classBuilder.PrepareClass());
        }
    }
}