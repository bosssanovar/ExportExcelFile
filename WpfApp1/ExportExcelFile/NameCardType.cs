using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportExcelFile
{
    public enum NameCardType
    {
        AAA,
        BBB,
        CCC,
        DDD,
    }

    public static partial class NameCardTypeEx
    {
        public static string GetButtonLabelLeft(this NameCardType type)
        {
            switch (type)
            {
                case NameCardType.AAA:
                    return "A Left";
                case NameCardType.BBB:
                    return "B Left";
                case NameCardType.CCC:
                    return "C Left";
                case NameCardType.DDD:
                    return "D Left";
                default:
                    throw new ArgumentException();
            }
        }
        public static string GetButtonLabelCenter(this NameCardType type)
        {
            switch (type)
            {
                case NameCardType.AAA:
                    return "A Center";
                case NameCardType.BBB:
                    return "B Center";
                case NameCardType.CCC:
                    return "C Center";
                case NameCardType.DDD:
                    return "D Center";
                default:
                    throw new ArgumentException();
            }
        }
        public static string GetButtonLabelRight(this NameCardType type)
        {
            switch (type)
            {
                case NameCardType.AAA:
                    return "A Right";
                case NameCardType.BBB:
                    return "B Right";
                case NameCardType.CCC:
                    return "C Right";
                case NameCardType.DDD:
                    return "D Right";
                default:
                    throw new ArgumentException();
            }
        }
    }
}
