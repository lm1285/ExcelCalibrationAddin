using System;
using System.Collections.Generic;
using System.Linq;
using ExcelCalibrationAddin.Contracts;

namespace ExcelCalibrationAddin.Core.Repositories
{
    public sealed partial class LocalTemplateRuleCacheRepository
    {
        public CachedTemplateRule FindBestSubsetMatch(IReadOnlyList<MeasurementRule> currentRules)
        {
            var current = ValidRules(currentRules);
            if (current.Count == 0) return null;

            var candidates = ListSavedTemplates()
                .Where(IsEligibleTemplate)
                .Select(template => new { Template = template, Saved = ValidRules(template.Rules) })
                .Where(item => item.Saved.Count >= current.Count)
                .Select(item => new { item.Template, item.Saved, Match = MatchSubset(current, item.Saved) })
                .Where(item => item.Match.Matched)
                .OrderBy(item => item.Saved.Count - current.Count)
                .ThenByDescending(item => item.Match.Score)
                .ThenByDescending(item => item.Template.UpdatedAt)
                .ToList();

            if (candidates.Count == 0) return null;
            var best = candidates[0];
            best.Template.MatchScore = 100;
            best.Template.MatchReason = best.Saved.Count == current.Count
                ? "ordered subset: equal rule count"
                : "ordered subset: fewer extra rules";
            return best.Template;
        }

        private static List<MeasurementRule> ValidRules(IEnumerable<MeasurementRule> rules)
        {
            return (rules ?? Enumerable.Empty<MeasurementRule>())
                .Where(rule => rule != null && !string.IsNullOrWhiteSpace(RuleName(rule)))
                .ToList();
        }

        private static bool IsEligibleTemplate(CachedTemplateRule template)
        {
            return template != null && template.Status == TemplateLifecycleStatus.Enabled && !template.DeletedAt.HasValue;
        }

        private static (bool Matched, int Score) MatchSubset(IReadOnlyList<MeasurementRule> current, IReadOnlyList<MeasurementRule> saved)
        {
            var hasBlocks = current.Any(rule => rule.BlockOrdinal.HasValue) || saved.Any(rule => rule.BlockOrdinal.HasValue);
            return hasBlocks ? MatchRepeatedBlocks(current, saved) : MatchOrderedRules(current, saved);
        }

        private static (bool Matched, int Score) MatchOrderedRules(IReadOnlyList<MeasurementRule> current, IReadOnlyList<MeasurementRule> saved)
        {
            var savedIndex = 0;
            var score = 0;
            foreach (var currentRule in current)
            {
                var found = -1;
                for (var index = savedIndex; index < saved.Count; index++)
                {
                    if (!SameRule(currentRule, saved[index])) continue;
                    found = index;
                    break;
                }
                if (found < 0) return (false, 0);
                score += 1 + (SameSheet(currentRule.TargetRange, saved[found].TargetRange) ? 2 : 0);
                savedIndex = found + 1;
            }
            return (true, score);
        }

        private static (bool Matched, int Score) MatchRepeatedBlocks(IReadOnlyList<MeasurementRule> current, IReadOnlyList<MeasurementRule> saved)
        {
            if (current.Any(InvalidBlockRule) || saved.Any(InvalidBlockRule)) return (false, 0);
            var currentBlocks = Blocks(current);
            var savedBlocks = Blocks(saved);
            if (currentBlocks.Count == 0 || currentBlocks.Count > savedBlocks.Count) return (false, 0);

            if (currentBlocks.Any(block => block.Any(rule => string.IsNullOrWhiteSpace(rule.BlockItemStructureSignature))) ||
                savedBlocks.Any(block => block.Any(rule => string.IsNullOrWhiteSpace(rule.BlockItemStructureSignature)))) return (false, 0);
            var expectedPattern = currentBlocks[0].Select(rule => rule.BlockItemStructureSignature).ToArray();
            if (currentBlocks.Any(block => !block.Select(rule => rule.BlockItemStructureSignature).SequenceEqual(expectedPattern, StringComparer.Ordinal))) return (false, 0);

            for (var offset = 0; offset <= savedBlocks.Count - currentBlocks.Count; offset++)
            {
                var score = 0;
                var compatible = true;
                int[] selectedPositions = null;
                for (var blockIndex = 0; blockIndex < currentBlocks.Count; blockIndex++)
                {
                    var candidateBlock = savedBlocks[offset + blockIndex];
                    var positions = new List<int>();
                    var savedIndex = 0;
                    foreach (var currentRule in currentBlocks[blockIndex])
                    {
                        var found = -1;
                        for (var index = savedIndex; index < candidateBlock.Count; index++)
                        {
                            if (!string.Equals(currentRule.BlockItemStructureSignature,
                                    candidateBlock[index].BlockItemStructureSignature, StringComparison.Ordinal)) continue;
                            found = index;
                            break;
                        }
                        if (found < 0)
                        {
                            compatible = false;
                            break;
                        }
                        positions.Add(candidateBlock[found].BlockRuleOrdinal.Value);
                        savedIndex = found + 1;
                        score += 3;
                    }
                    if (!compatible) break;
                    if (selectedPositions != null && !selectedPositions.SequenceEqual(positions))
                    {
                        compatible = false;
                        break;
                    }
                    selectedPositions = positions.ToArray();
                    score += 10;
                }
                if (compatible) return (true, score);
            }
            return (false, 0);
        }

        private static List<List<MeasurementRule>> Blocks(IReadOnlyList<MeasurementRule> rules)
        {
            return rules.GroupBy(rule => rule.BlockOrdinal.Value).OrderBy(group => group.Key).Select(group => group.OrderBy(rule => rule.BlockRuleOrdinal.Value).ToList()).ToList();
        }

        private static bool InvalidBlockRule(MeasurementRule rule)
        {
            return rule == null || !rule.BlockOrdinal.HasValue || !rule.BlockRuleOrdinal.HasValue;
        }

        private static bool SameRule(MeasurementRule left, MeasurementRule right)
        {
            return string.Equals(RuleName(left), RuleName(right), StringComparison.OrdinalIgnoreCase);
        }

        private static string RuleName(MeasurementRule rule)
        {
            return Normalize(rule?.FieldAlias ?? rule?.FieldName);
        }

        private static bool SameSheet(CellRange left, CellRange right)
        {
            return left != null && right != null && string.Equals(left.SheetName, right.SheetName, StringComparison.OrdinalIgnoreCase);
        }

    }
}
