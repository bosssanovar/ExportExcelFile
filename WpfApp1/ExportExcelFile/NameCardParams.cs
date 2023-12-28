using System.Drawing;

namespace ExportExcelFile
{
    public class NameCardParams(NameCardType nameCardType, Color color, string name)
    {
        public NameCardType NameCardType { get; } = nameCardType;
        public Color Color { get; } = color;
        public string Name { get; } = name;
    }
}
