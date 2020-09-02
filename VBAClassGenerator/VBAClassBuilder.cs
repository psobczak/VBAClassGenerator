using System;
using System.Collections.Generic;
using System.Text;

namespace VBAClassGenerator
{
    internal class VBAClassBuilder
    {
        public string WholeClass { get; set; }
        public List<VBAClassProperty> Properties { get; set; } = new List<VBAClassProperty>();
        private const string OPTION = "Option Explicit";

        public VBAClassBuilder AddProperty(VBAClassProperty prop)
        {
            if (!Properties.Contains(prop))
            {
                Properties.Add(prop);
            }
            else
            {
                Console.WriteLine($"Class already contains property {prop}");
            }

            return this;
        }

        public string PrepareClass()
        {
            var builder = new StringBuilder();
            builder.Append(OPTION);
            builder.Append(Environment.NewLine);
            builder.Append(Environment.NewLine);
            foreach (var prop in Properties)
            {
                var field = prop.ToString();
                builder.Append(field).Append(Environment.NewLine);
            }

            builder.Append(Environment.NewLine);
            foreach (var prop in Properties)
            {
                builder.Append(PrepareGetProperty(prop));
                builder.Append(PrepareLetProperty(prop));
            }

            builder.Append(PrepareToStringMethod());

            return builder.ToString();
        }

        private string PrepareGetProperty(VBAClassProperty prop)
        {
            var builder = new StringBuilder();
            var propName = prop.Name[0].ToString().ToUpper() + prop.Name.Substring(1);
            builder
                .Append("Property Get ")
                .Append(propName)
                .Append("() As ")
                .Append(prop.DataType)
                .Append(Environment.NewLine);
            builder
                .Append("\t")
                .Append(propName)
                .Append(" = ")
                .Append(prop.PrivateMemberName)
                .Append(Environment.NewLine);
            builder
                .Append("End Property")
                .Append(Environment.NewLine);
            builder
                .Append(Environment.NewLine);

            return builder.ToString();
        }

        private string PrepareLetProperty(VBAClassProperty prop)
        {
            var builder = new StringBuilder();
            var propName = prop.Name[0].ToString().ToUpper() + prop.Name.Substring(1);
            var methodArgument = $"{prop.Name}_ ";

            builder.Append("Property Let ")
                .Append(propName)
                .Append("(")
                .Append(methodArgument)
                .Append($"As {prop.DataType})")
                .Append(Environment.NewLine);
            builder
                .Append("\t")
                .Append(prop.PrivateMemberName)
                .Append(" = ")
                .Append(methodArgument)
                .Append(Environment.NewLine);
            builder
                .Append("End Property")
                .Append(Environment.NewLine);
            builder
                .Append(Environment.NewLine);

            return builder.ToString();
        }

        private string PrepareToStringMethod()
        {
            var builder = new StringBuilder();

            builder
                .Append("Public Function ToString() As String")
                .Append(Environment.NewLine);

            builder
                .Append("\t")
                .Append("ToString = ");

            for (int i = 0; i < Properties.Count; i++)
            {
                if (i > 0)
                {
                    builder.Append("\t");
                }
                var propName = Properties[i].Name[0].ToString().ToUpper() + Properties[i].Name.Substring(1);
                builder
                    .Append("\"")
                    .Append(propName)
                    .Append(" = \" & ")
                    .Append(Properties[i].PrivateMemberName);

                if (i != Properties.Count - 1)
                {
                    builder.Append(" & vbNewLine & _");
                }
                else
                {
                    builder.Append(" & vbNewLine");
                }
                builder.Append(Environment.NewLine);
            }

            builder.Append("End Function");
            builder.Append(Environment.NewLine);

            return builder.ToString();
        }
    }
}