using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Repositories;
using ExcelCalibrationAddin.Core.Services;
using ExcelCalibrationAddin.Host.Recognition;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelCalibrationAddin.Core.Tests
{
    [TestClass]
    public sealed class RepeatedBlockRecognitionTests
    {
        [TestMethod]
        public void RecognizerAssignsOrderedBlockAndItemOrdinalsAndIgnoresBlockTitles()
        {
            var sheet = new SheetSnapshot
            {
                Name = "Summary",
                Cells = new List<CellMeta>
                {
                    new CellMeta { Row = 5, Column = 1, Text = "1气体类别：" },
                    new CellMeta { Row = 15, Column = 1, Text = "2气体类别：" },
                    new CellMeta { Row = 25, Column = 1, Text = "3气体类别：" },
                    new CellMeta { Row = 34, Column = 1, Text = "note" }
                }
            };
            var mappings = new List<TemplateRegionMapping>();
            foreach (var start in new[] { 5, 15, 25 })
            {
                mappings.Add(new TemplateRegionMapping
                {
                    ProjectName = "示值误差",
                    SectionRange = Range(start + 4, start + 6),
                    MeasurementValueRange = Range(start + 4, start + 6),
                    StandardValueRange = Range(start + 4, start + 6)
                });
                mappings.Add(new TemplateRegionMapping
                {
                    ProjectName = "重复性",
                    SectionRange = Range(start + 7, start + 7),
                    MeasurementValueRange = Range(start + 7, start + 7),
                    StandardValueRange = Range(start + 7, start + 7)
                });
            }

            RepeatedBlockRecognizer.Apply(sheet, mappings);

            CollectionAssert.AreEqual(new[] { 1, 1, 2, 2, 3, 3 }, mappings.Select(item => item.BlockOrdinal.Value).ToArray());
            CollectionAssert.AreEqual(new[] { 1, 2, 1, 2, 1, 2 }, mappings.Select(item => item.BlockRuleOrdinal.Value).ToArray());
            Assert.AreEqual(mappings[0].BlockStructureSignature, mappings[2].BlockStructureSignature);
            Assert.AreEqual(mappings[2].BlockStructureSignature, mappings[4].BlockStructureSignature);
            Assert.AreEqual(5, mappings[0].BlockRange.StartRow);
            Assert.AreEqual(14, mappings[1].BlockRange.EndRow);
        }

        [TestMethod]
        public void FieldMatcherPrefersCalibrationTitlesInsideNumberedGasBlocks()
        {
            var cells = new List<CellMeta>();
            foreach (var (start, ordinal) in new[] { (5, 1), (15, 2) })
            {
                cells.Add(new CellMeta { Row = start, Column = 1, Text = ordinal + "、气体类别：" });
                cells.Add(new CellMeta { Row = start + 1, Column = 1, Text = "Type of gas" });
                cells.Add(new CellMeta { Row = start + 2, Column = 1, Text = "项目" });
                cells.Add(MergedProjectTitle(start + 4, "示值误差"));
                cells.Add(MergedProjectTitle(start + 7, "重复性"));
                cells.Add(MergedProjectTitle(start + 8, "响应时间"));
                cells.Add(new CellMeta { Row = start + 4, Column = 6, Formula = "=OtherSheet!A1" });
                cells.Add(new CellMeta { Row = start + 7, Column = 6, Formula = "=OtherSheet!A2" });
                cells.Add(new CellMeta { Row = start + 8, Column = 6, Formula = "=OtherSheet!A3" });
            }
            var fields = new FieldMatcher().MatchMeasurementFields(new SheetSnapshot { Name = "Summary", Cells = cells });

            Assert.AreEqual(6, fields.Count);
            CollectionAssert.AreEqual(new[] { "示值误差", "重复性", "响应时间", "示值误差", "重复性", "响应时间" },
                fields.OrderBy(field => field.Range.StartRow).Select(field => field.Alias).ToArray());
        }

        [TestMethod]
        public void SubsetMatchFindsTwoMatchingBlocksInsideSixBlockTemplate()
        {
            var repository = CreateRepository();
            repository.SaveTemplate("Five-in-one", Fingerprint("five"), BuildRules(6, 3));

            var matched = repository.FindBestSubsetMatch(BuildRules(2, 2));

            Assert.IsNotNull(matched);
            Assert.AreEqual(18, matched.Rules.Count);
        }

        [TestMethod]
        public void SubsetMatchAllowsSameNonContiguousItemDeletionPatternInEveryBlock()
        {
            var repository = CreateRepository();
            repository.SaveTemplate("Five-in-one", Fingerprint("five"), BuildRules(6, 3));

            var matched = repository.FindBestSubsetMatch(BuildRulesWithItems(2, new[] { 1, 3 }, 3));

            Assert.IsNotNull(matched);
        }

        [TestMethod]
        public void SubsetMatchRejectsDifferentDeletionPatternAcrossBlocks()
        {
            var repository = CreateRepository();
            repository.SaveTemplate("Five-in-one", Fingerprint("five"), BuildRules(6, 3));
            var current = BuildRules(2, 2);
            current.Add(BuildRule(2, 3, 3));

            Assert.IsNull(repository.FindBestSubsetMatch(current));
        }

        private static List<MeasurementRule> BuildRules(int blockCount, int itemCount)
        {
            var result = new List<MeasurementRule>();
            for (var block = 1; block <= blockCount; block++)
                for (var item = 1; item <= itemCount; item++)
                    result.Add(BuildRule(block, item, itemCount));
            return result;
        }

        private static List<MeasurementRule> BuildRulesWithItems(int blockCount, int[] itemPositions, int itemCount)
        {
            var result = new List<MeasurementRule>();
            for (var block = 1; block <= blockCount; block++)
                foreach (var item in itemPositions)
                    result.Add(BuildRule(block, item, itemCount));
            return result;
        }

        private static MeasurementRule BuildRule(int block, int item, int itemCount)
        {
            return new MeasurementRule
            {
                FieldName = "calibration item " + item,
                FieldAlias = "calibration item " + item,
                TargetRange = Range(block * 10 + item, block * 10 + item),
                FixedStandardValue = 10,
                FixedMpe = 1,
                ErrorSource = new ParameterSource { Range = Range(block * 10 + item, block * 10 + item) },
                FormatRule = new FormatRule { DecimalPlaces = 2 },
                WritableCells = new List<CellAddress> { new CellAddress { Row = block * 10 + item, Column = 3 } },
                BlockOrdinal = block,
                BlockRuleOrdinal = item,
                BlockRange = Range(block * 10, block * 10 + itemCount + 2),
                BlockStructureSignature = "same-layout-v1",
                BlockItemStructureSignature = "item-layout-v1-" + item
            };
        }

        private static CellRange Range(int startRow, int endRow)
        {
            return new CellRange
            {
                SheetName = "Sheet1",
                StartRow = startRow,
                EndRow = endRow,
                StartColumn = 1,
                EndColumn = 4
            };
        }

        private static CellMeta MergedProjectTitle(int row, string title)
        {
            return new CellMeta
            {
                Row = row,
                Column = 1,
                Text = title,
                IsMerged = true,
                MergeRange = new CellRange
                {
                    SheetName = "Summary",
                    StartRow = row,
                    EndRow = row + 2,
                    StartColumn = 1,
                    EndColumn = 5
                }
            };
        }

        private static LocalTemplateRuleCacheRepository CreateRepository()
        {
            var path = Path.Combine(Path.GetTempPath(), "ExcelCalibrationAddin.Tests", Guid.NewGuid().ToString("N") + ".sqlite");
            var repository = new LocalTemplateRuleCacheRepository(path);
            repository.Initialize();
            return repository;
        }

        private static TemplateFingerprint Fingerprint(string id)
        {
            return new TemplateFingerprint { ExactFingerprint = id, FuzzyFingerprint = id + "-fuzzy" };
        }
    }
}
