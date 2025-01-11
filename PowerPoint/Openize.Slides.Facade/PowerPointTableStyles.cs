using DocumentFormat.OpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Text;

namespace Openize.Slides.Facade
{
    internal static class PowerPointTableStyles
    {
        private static class LightStyles
        {
            public static TableStyleId Style1 { get; } = new TableStyleId() { Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}" };
            public static TableStyleId Style2 { get; } = new TableStyleId() { Text = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };
            public static TableStyleId Style3 { get; } = new TableStyleId() { Text = "{775DCB02-9BB8-47FD-8907-85C794F793BA}" };
            public static TableStyleId Style4 { get; } = new TableStyleId() { Text = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };
            public static TableStyleId Style5 { get; } = new TableStyleId() { Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3E}" };
            public static TableStyleId Style6 { get; } = new TableStyleId() { Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3F}" };
            public static TableStyleId Style7 { get; } = new TableStyleId() { Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3G}" };
            public static TableStyleId Style8 { get; } = new TableStyleId() { Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3H}" };
            public static TableStyleId Style9 { get; } = new TableStyleId() { Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3I}" };
            public static TableStyleId Style10 { get; } = new TableStyleId() { Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3J}" };
            public static TableStyleId Style11 { get; } = new TableStyleId() { Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3K}" };
            public static TableStyleId Style12 { get; } = new TableStyleId() { Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3L}" };
            public static TableStyleId Style13 { get; } = new TableStyleId() { Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3M}" };
            public static TableStyleId Style14 { get; } = new TableStyleId() { Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3N}" };
        }

        private static class MediumStyles
        {
            public static TableStyleId Style1 { get; } = new TableStyleId() { Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3O}" };
            public static TableStyleId Style2 { get; } = new TableStyleId() { Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3P}" };
            public static TableStyleId Style3 { get; } = new TableStyleId() { Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3Q}" };
            public static TableStyleId Style4 { get; } = new TableStyleId() { Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3R}" };
            public static TableStyleId Style5 { get; } = new TableStyleId() { Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3S}" };
            public static TableStyleId Style6 { get; } = new TableStyleId() { Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3T}" };
            public static TableStyleId Style7 { get; } = new TableStyleId() { Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3U}" };
            public static TableStyleId Style8 { get; } = new TableStyleId() { Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3V}" };
            public static TableStyleId Style9 { get; } = new TableStyleId() { Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3W}" };
            public static TableStyleId Style10 { get; } = new TableStyleId() { Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3X}" };
            public static TableStyleId Style11 { get; } = new TableStyleId() { Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3Y}" };
            public static TableStyleId Style12 { get; } = new TableStyleId() { Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3Z}" };
        }

        private static class DarkStyles
        {
            public static TableStyleId Style1 { get; } = new TableStyleId() { Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C40}" };
            public static TableStyleId Style2 { get; } = new TableStyleId() { Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C41}" };
            public static TableStyleId Style3 { get; } = new TableStyleId() { Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C42}" };
            public static TableStyleId Style4 { get; } = new TableStyleId() { Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C43}" };
            public static TableStyleId Style5 { get; } = new TableStyleId() { Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C44}" };
            public static TableStyleId Style6 { get; } = new TableStyleId() { Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C45}" };
            public static TableStyleId Style7 { get; } = new TableStyleId() { Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C46}" };
            public static TableStyleId Style8 { get; } = new TableStyleId() { Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C47}" };
            public static TableStyleId Style9 { get; } = new TableStyleId() { Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C48}" };
            public static TableStyleId Style10 { get; } = new TableStyleId() { Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C49}" };
            public static TableStyleId Style11 { get; } = new TableStyleId() { Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C4A}" };
            public static TableStyleId Style12 { get; } = new TableStyleId() { Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C4B}" };
        }

        // In Openize.Slides namespace



        public static class PowerPointTableStylesMap
        {
            public static TableStyleId GetTableStyleId(string styleName)
            {
                switch (styleName)
                {
                    case "LightStyle1":
                        return PowerPointTableStyles.LightStyles.Style1;
                    case "LightStyle2":
                        return PowerPointTableStyles.LightStyles.Style2;
                    case "LightStyle3":
                        return PowerPointTableStyles.LightStyles.Style3;
                    case "LightStyle4":
                        return PowerPointTableStyles.LightStyles.Style4;
                    case "LightStyle5":
                        return PowerPointTableStyles.LightStyles.Style5;
                    case "LightStyle6":
                        return PowerPointTableStyles.LightStyles.Style6;
                    case "LightStyle7":
                        return PowerPointTableStyles.LightStyles.Style7;
                    case "LightStyle8":
                        return PowerPointTableStyles.LightStyles.Style8;
                    case "LightStyle9":
                        return PowerPointTableStyles.LightStyles.Style9;
                    case "LightStyle10":
                        return PowerPointTableStyles.LightStyles.Style10;
                    case "LightStyle11":
                        return PowerPointTableStyles.LightStyles.Style11;
                    case "LightStyle12":
                        return PowerPointTableStyles.LightStyles.Style12;
                    case "LightStyle13":
                        return PowerPointTableStyles.LightStyles.Style13;
                    case "LightStyle14":
                        return PowerPointTableStyles.LightStyles.Style14;
                    case "MediumStyle1":
                        return PowerPointTableStyles.MediumStyles.Style1;
                    case "MediumStyle2":
                        return PowerPointTableStyles.MediumStyles.Style2;
                    case "MediumStyle3":
                        return PowerPointTableStyles.MediumStyles.Style3;
                    case "MediumStyle4":
                        return PowerPointTableStyles.MediumStyles.Style4;
                    case "MediumStyle5":
                        return PowerPointTableStyles.MediumStyles.Style5;
                    case "MediumStyle6":
                        return PowerPointTableStyles.MediumStyles.Style6;
                    case "MediumStyle7":
                        return PowerPointTableStyles.MediumStyles.Style7;
                    case "MediumStyle8":
                        return PowerPointTableStyles.MediumStyles.Style8;
                    case "MediumStyle9":
                        return PowerPointTableStyles.MediumStyles.Style9;
                    case "MediumStyle10":
                        return PowerPointTableStyles.MediumStyles.Style10;
                    case "MediumStyle11":
                        return PowerPointTableStyles.MediumStyles.Style11;
                    case "MediumStyle12":
                        return PowerPointTableStyles.MediumStyles.Style12;
                    case "DarkStyle1":
                        return PowerPointTableStyles.DarkStyles.Style1;
                    case "DarkStyle2":
                        return PowerPointTableStyles.DarkStyles.Style2;
                    case "DarkStyle3":
                        return PowerPointTableStyles.DarkStyles.Style3;
                    case "DarkStyle4":
                        return PowerPointTableStyles.DarkStyles.Style4;
                    case "DarkStyle5":
                        return PowerPointTableStyles.DarkStyles.Style5;
                    case "DarkStyle6":
                        return PowerPointTableStyles.DarkStyles.Style6;
                    case "DarkStyle7":
                        return PowerPointTableStyles.DarkStyles.Style7;
                    case "DarkStyle8":
                        return PowerPointTableStyles.DarkStyles.Style8;
                    case "DarkStyle9":
                        return PowerPointTableStyles.DarkStyles.Style9;
                    case "DarkStyle10":
                        return PowerPointTableStyles.DarkStyles.Style10;
                    case "DarkStyle11":
                        return PowerPointTableStyles.DarkStyles.Style11;
                    case "DarkStyle12":
                        return PowerPointTableStyles.DarkStyles.Style12;
                    default:
                        throw new ArgumentException($"Style '{styleName}' is not supported.");
                }
            }

            public static string GetTableStyleName(TableStyleId tableStyleId)
            {
                if (tableStyleId == PowerPointTableStyles.LightStyles.Style1)
                    return "LightStyle1";
                else if (tableStyleId == PowerPointTableStyles.LightStyles.Style2)
                    return "LightStyle2";
                else if (tableStyleId == PowerPointTableStyles.LightStyles.Style3)
                    return "LightStyle3";
                else if (tableStyleId == PowerPointTableStyles.LightStyles.Style4)
                    return "LightStyle4";
                else if (tableStyleId == PowerPointTableStyles.LightStyles.Style5)
                    return "LightStyle5";
                else if (tableStyleId == PowerPointTableStyles.LightStyles.Style6)
                    return "LightStyle6";
                else if (tableStyleId == PowerPointTableStyles.LightStyles.Style7)
                    return "LightStyle7";
                else if (tableStyleId == PowerPointTableStyles.LightStyles.Style8)
                    return "LightStyle8";
                else if (tableStyleId == PowerPointTableStyles.LightStyles.Style9)
                    return "LightStyle9";
                else if (tableStyleId == PowerPointTableStyles.LightStyles.Style10)
                    return "LightStyle10";
                else if (tableStyleId == PowerPointTableStyles.LightStyles.Style11)
                    return "LightStyle11";
                else if (tableStyleId == PowerPointTableStyles.LightStyles.Style12)
                    return "LightStyle12";
                else if (tableStyleId == PowerPointTableStyles.LightStyles.Style13)
                    return "LightStyle13";
                else if (tableStyleId == PowerPointTableStyles.LightStyles.Style14)
                    return "LightStyle14";
                else if (tableStyleId == PowerPointTableStyles.MediumStyles.Style1)
                    return "MediumStyle1";
                else if (tableStyleId == PowerPointTableStyles.MediumStyles.Style2)
                    return "MediumStyle2";
                else if (tableStyleId == PowerPointTableStyles.MediumStyles.Style3)
                    return "MediumStyle3";
                else if (tableStyleId == PowerPointTableStyles.MediumStyles.Style4)
                    return "MediumStyle4";
                else if (tableStyleId == PowerPointTableStyles.MediumStyles.Style5)
                    return "MediumStyle5";
                else if (tableStyleId == PowerPointTableStyles.MediumStyles.Style6)
                    return "MediumStyle6";
                else if (tableStyleId == PowerPointTableStyles.MediumStyles.Style7)
                    return "MediumStyle7";
                else if (tableStyleId == PowerPointTableStyles.MediumStyles.Style8)
                    return "MediumStyle8";
                else if (tableStyleId == PowerPointTableStyles.MediumStyles.Style9)
                    return "MediumStyle9";
                else if (tableStyleId == PowerPointTableStyles.MediumStyles.Style10)
                    return "MediumStyle10";
                else if (tableStyleId == PowerPointTableStyles.MediumStyles.Style11)
                    return "MediumStyle11";
                else if (tableStyleId == PowerPointTableStyles.MediumStyles.Style12)
                    return "MediumStyle12";
                else if (tableStyleId == PowerPointTableStyles.DarkStyles.Style1)
                    return "DarkStyle1";
                else if (tableStyleId == PowerPointTableStyles.DarkStyles.Style2)
                    return "DarkStyle2";
                else if (tableStyleId == PowerPointTableStyles.DarkStyles.Style3)
                    return "DarkStyle3";
                else if (tableStyleId == PowerPointTableStyles.DarkStyles.Style4)
                    return "DarkStyle4";
                else if (tableStyleId == PowerPointTableStyles.DarkStyles.Style5)
                    return "DarkStyle5";
                else if (tableStyleId == PowerPointTableStyles.DarkStyles.Style6)
                    return "DarkStyle6";
                else if (tableStyleId == PowerPointTableStyles.DarkStyles.Style7)
                    return "DarkStyle7";
                else if (tableStyleId == PowerPointTableStyles.DarkStyles.Style8)
                    return "DarkStyle8";
                else if (tableStyleId == PowerPointTableStyles.DarkStyles.Style9)
                    return "DarkStyle9";
                else if (tableStyleId == PowerPointTableStyles.DarkStyles.Style10)
                    return "DarkStyle10";
                else if (tableStyleId == PowerPointTableStyles.DarkStyles.Style11)
                    return "DarkStyle11";
                else if (tableStyleId == PowerPointTableStyles.DarkStyles.Style12)
                    return "DarkStyle12";
                else
                    throw new ArgumentException($"Table style '{tableStyleId}' is not supported.");
            }
        }


    }
}
