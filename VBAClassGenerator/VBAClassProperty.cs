namespace VBAClassGenerator
{
    internal class VBAClassProperty
    {
        public string Name { get; set; }

        public string PrivateMemberName
        {
            get => $"m_{Name}";
        }

        public VBADataType DataType { get; set; }

        public override string ToString()
        {
            return $"Private {PrivateMemberName} As {DataType}";
        }

        public VBAClassProperty(string name, VBADataType dataType)
        {
            Name = name;
            DataType = dataType;
        }
    }
}