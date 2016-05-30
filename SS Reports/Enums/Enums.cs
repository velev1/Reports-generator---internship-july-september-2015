using System.ComponentModel;
using System.Reflection;

namespace SS_Reports.Enums
{
    internal enum Stores : short
    {
        [Description("Technopolis")]
        Technopolis,
        [Description("Technomarket")]
        Technomarket
    }
    /// <summary>
    /// Abbreviations used for the platforms in the output file
    /// </summary>
    internal enum OutputAbbreviations : short
    {

        [Description("PS2")]
        PS2,
        [Description("PS3")]
        PS3,
        [Description("PS4")]
        PS4,
        [Description("XBOX360")]
        XBOX360,
        [Description("XBOXONE")]
        XBOXONE,
        [Description("WII")]
        WII,
        [Description("PSP")]
        PSP,
        [Description("3DS")]
        DS3,
        [Description("PSVITA")]
        PSVITA,
        [Description("PC")]
        PC,
        [Description("NDS")]
        NDS,
        [Description("Other")]
        Other
    }
    //Abbreviations used for the platforms in the technopolis source file.
    internal enum TechnopolisAbbreviations : short
    {
        [Description("P2")]
        PS2,
        [Description("P3")]
        PS3,
        [Description("P4")]
        PS4,
        [Description("XB3")]
        XBOX360,
        [Description("XBO")]
        XBOXONE,
        [Description("WII")]
        WII,
        [Description("PSP")]
        PSP,
        [Description("3D")]
        DS3,
        [Description("PSV")]
        PSVITA,
        [Description("PC")]
        PC,
        [Description("DS")]
        NDS,
        [Description("Other")]
        Other
    }
    /// <summary>
    /// Abbreviations used for the platforms in the technomarket source file.
    /// </summary>
    /// They are the same as the output abbreviations
    internal enum TechnomarketAbbreviations : short
    {
        [Description("PS2")]
        PS2,
        [Description("PS3")]
        PS3,
        [Description("PS4")]
        PS4,
        [Description("XBOX360")]
        XBOX360,
        [Description("XBOXONE")]
        XBOXONE,
        [Description("WII")]
        WII,
        [Description("PSP")]
        PSP,
        [Description("3DS")]
        DS3,
        [Description("PSVITA")]
        PSVITA,
        [Description("PC")]
        PC,
        [Description("NDS")]
        NDS,
        [Description("Other")]
        Other
    }
    /// <summary>
    /// Reads and returns the description off any of the above enums.
    /// </summary>
    internal class EnumHelper
    {
        public static string GetDescription(object enumValue)
        {
            string defDesc = "";
            FieldInfo fi = enumValue.GetType().GetField(enumValue.ToString());

            if (fi != null)
            {
                object[] attrs = fi.GetCustomAttributes(typeof(DescriptionAttribute), true);
                if (attrs != null && attrs.Length > 0)
                    return ((DescriptionAttribute)attrs[0]).Description;
            }

            return defDesc;
        }
    }
}
