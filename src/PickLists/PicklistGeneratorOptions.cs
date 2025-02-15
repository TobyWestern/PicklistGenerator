namespace BrickAtHeart.LUGTools.PicklistGenerator
{
    public class PicklistGeneratorOptions
    {
        public const string Section = "PicklistGenerator";

        public string CsvFilePath { get; set; }

        public int PartStartRow { get; set; }

        public int PersonRow { get; set; }

        public int PersonRowStartColumn { get; set; }

        public int IndexColumn { get; set; }

        public int BricklinkColorDescriptionColumn { get; set; }

        public int LegoElementIdColumn { get; set; }

        public int LegoElementDescriptionColumn { get; set; }

        public int LegoColorDescriptionColumn { get; set; }

        public string PerPartFileName { get; set; }

        public string PerPersonFileName { get; set; }
    }
}