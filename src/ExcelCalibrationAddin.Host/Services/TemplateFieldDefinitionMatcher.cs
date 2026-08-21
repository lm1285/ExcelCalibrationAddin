using System;
using System.Collections.Generic;
using System.Linq;
using ExcelCalibrationAddin.Contracts;

namespace ExcelCalibrationAddin.Host.Services
{
    public static class TemplateFieldDefinitionMatcher
    {
        private static readonly TemplateRegionRole[] CriticalRoles =
        {
            TemplateRegionRole.SetpointValue,
            TemplateRegionRole.StandardValue,
            TemplateRegionRole.MeasurementValue,
            TemplateRegionRole.ErrorValue,
            TemplateRegionRole.TechnicalRequirement,
            TemplateRegionRole.RangeValue
        };

        public static bool IsCompatible(TemplateFieldDefinition saved, TemplateFieldDefinition current)
        {
            if (saved == null || current == null)
            {
                return true;
            }

            if (!string.Equals(Normalize(saved.ProjectName), Normalize(current.ProjectName), StringComparison.Ordinal))
            {
                return false;
            }

            foreach (var role in CriticalRoles)
            {
                var savedRegion = FindRegion(saved, role);
                var currentRegion = FindRegion(current, role);
                if (savedRegion == null && currentRegion == null) continue;
                if (savedRegion == null || currentRegion == null) return false;
                if (!SameRelativeRange(saved.SectionRange, savedRegion.Range, current.SectionRange, currentRegion.Range))
                {
                    return false;
                }

                if (!SameHeaderPath(savedRegion.HeaderPath, currentRegion.HeaderPath))
                {
                    return false;
                }

                if (!SameFormulaSet(savedRegion, currentRegion))
                {
                    return false;
                }

                if (!HasCompatibleUnits(savedRegion, currentRegion))
                {
                    return false;
                }
            }

            return true;
        }

        private static bool SameRelativeRange(
            CellRange savedSection,
            CellRange saved,
            CellRange currentSection,
            CellRange current)
        {
            if (saved == null || current == null) return saved == null && current == null;
            var savedRowOffset = saved.StartRow - (savedSection?.StartRow ?? 1);
            var savedColumnOffset = saved.StartColumn - (savedSection?.StartColumn ?? 1);
            var currentRowOffset = current.StartRow - (currentSection?.StartRow ?? 1);
            var currentColumnOffset = current.StartColumn - (currentSection?.StartColumn ?? 1);
            return savedRowOffset == currentRowOffset &&
                savedColumnOffset == currentColumnOffset &&
                RowCount(saved) == RowCount(current) &&
                ColumnCount(saved) == ColumnCount(current);
        }

        private static bool SameHeaderPath(IReadOnlyList<string> saved, IReadOnlyList<string> current)
        {
            var savedPath = NormalizePath(saved);
            var currentPath = NormalizePath(current);
            return savedPath.Count == 0 || currentPath.Count == 0 || savedPath.SequenceEqual(currentPath);
        }

        private static bool SameFormulaSet(TemplateRegionDefinition saved, TemplateRegionDefinition current)
        {
            var savedFormulas = FormulaSet(saved);
            var currentFormulas = FormulaSet(current);
            return savedFormulas.Count == currentFormulas.Count && savedFormulas.SetEquals(currentFormulas);
        }

        private static HashSet<string> FormulaSet(TemplateRegionDefinition region)
        {
            var formulas = (region?.FormulaVariants ?? new List<TemplateFormulaDefinition>())
                .Where(formula => formula != null)
                .ToList();
            if (region?.Formula != null)
            {
                formulas.Add(region.Formula);
            }

            return new HashSet<string>(formulas.Select(formula => NormalizeFormula(
                    string.IsNullOrWhiteSpace(formula.FormulaR1C1) ? formula.Formula : formula.FormulaR1C1)),
                StringComparer.OrdinalIgnoreCase);
        }

        private static bool HasCompatibleUnits(TemplateRegionDefinition saved, TemplateRegionDefinition current)
        {
            var savedBranchUnits = BranchUnits(saved);
            var currentBranchUnits = BranchUnits(current);
            if (savedBranchUnits.Count > 0 || currentBranchUnits.Count > 0)
            {
                return savedBranchUnits.SetEquals(currentBranchUnits);
            }

            return TemplateUnitParser.SameUnitFamily(saved.Unit, current.Unit);
        }

        private static HashSet<string> BranchUnits(TemplateRegionDefinition region)
        {
            var formulas = (region?.FormulaVariants ?? new List<TemplateFormulaDefinition>()).ToList();
            if (formulas.Count == 0 && region?.Formula != null) formulas.Add(region.Formula);
            return new HashSet<string>(formulas
                .SelectMany(formula => formula?.Branches ?? new List<TemplateFormulaBranch>())
                .Select(branch => Normalize(branch?.Unit))
                .Where(unit => !string.IsNullOrWhiteSpace(unit)), StringComparer.OrdinalIgnoreCase);
        }

        private static TemplateRegionDefinition FindRegion(TemplateFieldDefinition definition, TemplateRegionRole role)
        {
            return (definition?.Regions ?? new List<TemplateRegionDefinition>())
                .FirstOrDefault(region => region?.Role == role);
        }

        private static List<string> NormalizePath(IEnumerable<string> path)
        {
            return (path ?? Enumerable.Empty<string>())
                .Select(Normalize)
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct()
                .ToList();
        }

        private static string NormalizeFormula(string value)
        {
            return Normalize(value).Replace(" ", string.Empty);
        }

        private static string Normalize(string value)
        {
            return new string((value ?? string.Empty)
                .Trim()
                .Replace("％", "%")
                .Where(character => !char.IsWhiteSpace(character))
                .ToArray())
                .ToUpperInvariant();
        }

        private static int RowCount(CellRange range) => range.EndRow - range.StartRow + 1;
        private static int ColumnCount(CellRange range) => range.EndColumn - range.StartColumn + 1;
    }
}
