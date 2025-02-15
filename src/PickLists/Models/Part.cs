namespace BrickAtHeart.LUGTools.PicklistGenerator.Models
{
    public class Part
    {
        public int Index { get; set; } = -1;

        public string BricklinkColorDescription { get; set; } = string.Empty;

        public string LegoColorDescription { get; set; } = string.Empty;

        public string LegoElementDescription { get; set; } = string.Empty;

        public string LegoElementId { get; set; }  = string.Empty;

        public static string MapLegoColor(string legoColor)
        {
            switch (legoColor.ToUpperInvariant().Trim())
            {
                case "AQUA":
                    return "LIGHT AQUA";
                case "BLACK":
                    return "BLACK";
                case "BRICK RED":
                    return "FABULAND BROWN";
                case "BRICK YELLOW":
                case "BRICK-YEL":
                    return "TAN";
                case "BRIGHT BLUE":
                case "BR.BLUE":
                    return "BLUE";
                case "BRIGHT BLUISH GREEN":
                    return "DARK TURQUOISE";
                case "BRIGHT BLUISH VIOLET":
                    return "VIOLET";
                case "BRIGHT GREEN":
                case "BR.GREEN":
                    return "BRIGHT GREEN";
                case "BRIGHT ORANGE":
                    return "ORANGE";
                case "BRIGHT PURPLE":
                    return "DARK PINK";
                case "BRIGHT RED":
                case "BR.RED":
                    return "RED";
                case "BRIGHT REDDISH LILAC":
                    return "LIGHT PURPLE";
                case "BRIGHT REDDISH VIOLET":
                    return "MAGENTA";
                case "BRIGHT VIOLET":
                    return "PURPLE";
                case "BRIGHT YELLOW":
                case "BR.YEL":
                    return "YELLOW";
                case "BRIGHT YELLOW GREEN":
                case "BRIGHT YELLOWISH GREEN":
                case "BR.YEL-GREEN":
                    return "LIME";
                case "BRIGHT YELLOWISH ORANGE":
                    return "MEDIUM ORANGE";
                case "BROWN":
                    return "DARK FLESH";
                case "COOL SILVER":
                    return "PEARL LIGHT GRAY";
                case "COOL SILVER, DIFFUSE":
                    return "SPECKLE BLACK-SILVER";
                case "COOL SILVER, DRUM LACQUERED":
                    return "METALLIC SILVER";
                case "COOL YELLOW":
                    return "BRIGHT LIGHT YELLOW";
                case "COPPER":
                case "COPPER METALLIC":
                    return "COPPER";
                case "COPPER, DIFFUSE":
                    return "SPECKLE BLACK-COPPER";
                case "CURRY":
                    return "DARK YELLOW";
                case "DARK AZUR":
                    return "DARK AZURE";
                case "DARK BROWN":
                case "DK. BROWN":
                    return "DARK BROWN";
                case "DARK GREEN":
                case "DK.GREEN":
                    return "GREEN";
                case "DARK GREY":
                    return "DARK GREY";
                case "DARK GREY METALLIC":
                    return "PEARL DARK GREY";
                case "DARK ORANGE":
                case "DK.ORA":
                    return "DARK ORANGE";
                case "DARK RED":
                    return "DARK RED";
                case "DARK ROYAL BLUE":
                    return "DARK BLUE-VIOLET";
                case "DARK STONE GREY":
                case "DK. ST. GREY":
                    return "DARK BLUISH GRAY";
                case "DOVE BLUE":
                    return "SKY BLUE";
                case "EARTH BLUE":
                    return "DARK BLUE";
                case "EARTH GREEN":
                    return "DARK GREEN";
                case "EARTH ORANGE":
                    return "BROWN";
                case "FLAME YELLOWISH ORANGE":
                    return "BRIGHT LIGHT ORANGE";
                case "GOLD":
                    return "PEARL LIGHT GOLD";
                case "GOLD INK":
                    return "METALLIC GOLD";
                case "GREY":
                    return "LIGHT GRAY";
                case "L.NOUGAT":
                    return "LIGHT NOUGAT";
                case "LAVENDER":
                    return "LAVENDER";
                case "LEMON METALLIC":
                    return "METALLIC GREEN";
                case "LIGHT BLUE":
                    return "LIGHT BLUE";
                case "LIGHT BLUISH GREEN":
                    return "AQUA";
                case "LIGHT BLUISH VIOLET":
                    return "LIGHT VIOLET";
                case "LGH.ROY.BLUE":
                    return "BRIGHT LIGHT BLUE";
                case "LIGHT BROWN":
                    return "FABULAND ORANGE";
                case "LIGHT GREEN":
                    return "LIGHT GREEN";
                case "LIGHT GREY":
                    return "VERY LIGHT GRAY";
                case "LIGHT GREY METALLIC":
                    return "PEARL VERY LIGHT GRAY";
                case "LIGHT NOUGAT":
                    return "LIGHT FLESH";
                case "LIGHT ORANGE BROWN":
                    return "EARTH ORANGE";
                case "LIGHT PURPLE":
                    return "BRIGHT PINK";
                case "LIGHT RED":
                    return "LIGHT SALMON";
                case "LIGHT REDDISH VIOLET":
                    return "PINK";
                case "LIGHT ROYAL BLUE":
                    return "BRIGHT LIGHT BLUE";
                case "LIGHT STONE GREY":
                    return "VERY LIGHT BLUISH GRAY";
                case "LIGHT YELLOW":
                    return "LIGHT YELLOW";
                case "LIGHT YELLOWISH GREEN":
                    return "LIGHT LIME";
                case "LIGHT YELLOWISH ORANGE":
                    return "VERY LIGHT ORANGE";
                case "LILAC":
                    return "MEDIUM VIOLET";
                case "MEDIUM AZUR":
                    return "MEDIUM AZURE";
                case "MEDIUM BLUE":
                    return "MEDIUM BLUE";
                case "MEDIUM BLUISH GREEN":
                    return "LIGHT TURQUOISE";
                case "MEDIUM BLUISH VIOLET":
                    return "MEDIUM VIOLET";
                case "MEDIUM GREEN":
                    return "MEDIUM GREEN";
                case "MEDIUM LAVENDER":
                    return "MEDIUM LAVENDER";
                case "MEDIUM LILAC":
                    return "DARK PURPLE";
                case "MEDIUM NOUGAT":
                case "M. NOUGAT":
                    return "MEDIUM NOUGAT";
                case "MEDIUM RED":
                    return "SALMON";
                case "MEDIUM REDDISH VIOLET":
                    return "DARK PINK";
                case "MEDIUM STONE GREY":
                case "MED. ST-GREY":
                    return "LIGHT BLUISH GRAY";
                case "MEDIUM YELLOWISH GREEN":
                    return "MEDIUM LIME";
                case "MEDIUM YELLOWISH ORANGE":
                    return "LIGHT ORANGE";
                case "METALIZED GOLD":
                    return "CHROME GOLD";
                case "METALIZED SILVER":
                    return "CHROME SILVER";
                case "METALLIC WHITE":
                    return "PEARL WHITE";
                case "NATURE":
                    return "MILKY WHITE";
                case "NEW DARK RED":
                    return "DARK RED";
                case "NOUGAT":
                    return "FLESH";
                case "OLIVE GREEN":
                    return "OLIVE GREEN";
                case "PASTEL BLUE":
                    return "MAERSK BLUE";
                case "PHOSPHORESCENT GREEN":
                    return "GLOW IN DARK TRANS";
                case "PHOSPHORESCENT WHITE":
                    return "GLOW IN DARK OPAQUE";
                case "PINK":
                    return "MEDIUM DARK PINK";
                case "REDDISH BROWN":
                case "RED. BROWN":
                    return "REDDISH BROWN";
                case "ROYAL BLUE":
                    return "BLUE-VIOLET";
                case "SAND BLUE":
                    return "SAND BLUE";
                case "SAND GREEN":
                    return "SAND GREEN";
                case "SAND YELLOW":
                    return "DARK TAN";
                case "SILVER METAL.":
                    return "FLAT SILVER";
                case "TITAN. METAL.":
                    return "PEARL DSRK GREY";
                case "TR.":
                    return "TRANSPARENT";
                case "TR. BROWN OPAL":
                    return "TRANSPARENT BROWN";
                case "TR.BLUE":
                    return "TRANS DARK BLUE";
                case "TR.BROWN":
                    return "TRANSPARENT BLACK";
                case "TR.L.BLUE":
                    return "TRANSPARENT LIGHT BLUE";
                case "W.GOLD":
                    return "PEARL GOLD";
                case "WHITE":
                    return "WHITE";
                default:
                    return "UNKNOWN";
            }
        }
    }
}