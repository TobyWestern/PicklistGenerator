using BrickAtHeart.LUGTools.PicklistGenerator.Models;
using Microsoft.Extensions.Options;
using Microsoft.VisualBasic.FileIO;
using System.Collections.Generic;
using System.Linq;

namespace BrickAtHeart.LUGTools.PicklistGenerator
{
    public class ExportFile
    {
        public ExportFile(IOptions<PicklistGeneratorOptions> options)
        {
            this.options = options.Value;
        }

        public List<Person> ReadPeople()
        {
            List<Person> people = new List<Person>();

            TextFieldParser csvParser = new TextFieldParser(options.CsvFilePath)
            {
                CommentTokens = new[] { "#" }
            };

            csvParser.SetDelimiters(",");
            csvParser.HasFieldsEnclosedInQuotes = true;

            while (csvParser.LineNumber < options.PersonRow)
            {
                csvParser.ReadFields();
            }

            string[] fields = csvParser.ReadFields();

            if (fields == null)
            {
                return people;
            }

            for (int column = options.PersonRowStartColumn; column < fields.Length; column++)
            {
                string[] person = fields[column].Split('.');

                if (person.Length < 2)
                {
                    continue;
                }

                if (person.Length > 2)
                {
                    person[1] = person[1].Trim() + ". " + person[2].Trim();
                }

                if (!string.IsNullOrWhiteSpace(person[1]))
                {
                    people.Add(new Person
                    {
                        Index = int.Parse(person[0]),
                        Column = column,
                        FullName = person[1].Trim()
                    });
                }
            }

            return people;
        }

        public Dictionary<int, Part> ReadParts()
        {
            Dictionary<int, Part> parts = new Dictionary<int, Part>();

            TextFieldParser csvParser = new TextFieldParser(options.CsvFilePath)
            {
                CommentTokens = new[] { "#" }
            };

            csvParser.SetDelimiters(",");
            csvParser.HasFieldsEnclosedInQuotes = true;

            while (csvParser.LineNumber < options.PartStartRow)
            {
                csvParser.ReadFields();
            }

            while (!csvParser.EndOfData)
            {
                string[] fields = csvParser.ReadFields();

                if (fields == null)
                {
                    return parts;
                }

                if (int.TryParse(fields[options.IndexColumn], out int index))
                {
                    parts.Add(index, new Part
                    {
                        Index = index,
                        BricklinkColorDescription = fields[options.BricklinkColorDescriptionColumn] == "#REF!" ?
                            Part.MapLegoColor(fields[options.LegoColorDescriptionColumn].ToUpperInvariant()) :
                            fields[options.BricklinkColorDescriptionColumn].ToUpperInvariant(),
                        LegoElementId = fields[options.LegoElementIdColumn],
                        LegoElementDescription = fields[options.LegoElementDescriptionColumn].ToUpperInvariant(),
                        LegoColorDescription = fields[options.LegoColorDescriptionColumn].ToUpperInvariant()
                    });
                }
            }

            return parts;
        }

        public List<Order> ReadOrders(List<Person> people, Dictionary<int, Part> parts)
        {
           List<Order> orders = new List<Order>();

            TextFieldParser csvParser = new TextFieldParser(options.CsvFilePath)
            {
                CommentTokens = new[] { "#" }
            };

            csvParser.SetDelimiters(",");
            csvParser.HasFieldsEnclosedInQuotes = true;

            while (csvParser.LineNumber < options.PartStartRow)
            {
                csvParser.ReadFields();
            }

            int rowCount = 1;

            while (!csvParser.EndOfData)
            {
                string[] fields = csvParser.ReadFields();

                if (fields == null)
                {
                    return orders;
                }

                if (rowCount <= parts.Count)
                {
                    for (int index = 0; index < people.Count + 5; index++)
                    {
                        if (int.TryParse(fields[index * 2 + options.PersonRowStartColumn], out int quantity) && quantity > 0)
                        {
                            if (GetPartIndex(csvParser.LineNumber, parts.Count) != -1)
                            {
                                orders.Add(new Order
                                {
                                    Person = people.First(p => p.Column == (index * 2 + options.PersonRowStartColumn)),
                                    Part = parts[GetPartIndex(csvParser.LineNumber, parts.Count)],
                                    Quantity = quantity
                                });
                            }
                        }
                    }
                }

                rowCount++;
            }

            return orders;
        }

        private int GetPartIndex(long lineNumber, int partCount)
        {
            if (lineNumber == -1)
            {
                return partCount;
            }

            return (int) (lineNumber - options.PartStartRow);
        }

        private readonly PicklistGeneratorOptions options;
    }
}