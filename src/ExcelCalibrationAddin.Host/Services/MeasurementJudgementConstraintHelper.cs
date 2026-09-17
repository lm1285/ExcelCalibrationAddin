using System;
using System.Collections.Generic;
using System.Linq;
using ExcelCalibrationAddin.Contracts;

namespace ExcelCalibrationAddin.Host.Services
{
    internal static class MeasurementJudgementConstraintHelper
    {
        public static IReadOnlyList<MeasurementJudgementConstraint> Collect(MeasurementRule rule)
        {
            var constraints = new List<MeasurementJudgementConstraint>();
            if (rule == null)
            {
                return constraints;
            }

            if (HasPrimaryConstraint(rule))
            {
                constraints.Add(new MeasurementJudgementConstraint
                {
                    Name = string.IsNullOrWhiteSpace(rule.ErrorSource?.Name) ? "误差" : rule.ErrorSource.Name,
                    ErrorSource = rule.ErrorSource,
                    MpeSource = rule.MpeSource,
                    ResultSource = rule.ResultSource,
                    ErrorType = rule.ErrorType,
                    FixedMpe = rule.FixedMpe,
                    FixedNegativeTolerance = rule.FixedNegativeTolerance,
                    FixedPositiveTolerance = rule.FixedPositiveTolerance,
                    RequirementOperator = rule.RequirementOperator,
                    ErrorFormula = rule.ErrorFormula
                });
            }

            constraints.AddRange((rule.AdditionalJudgementConstraints ?? new List<MeasurementJudgementConstraint>())
                .Where(item => item != null && HasValidRange(item.ErrorSource?.Range)));
            return constraints;
        }

        public static MeasurementRule Overlay(MeasurementRule rule, MeasurementJudgementConstraint constraint)
        {
            if (rule == null || constraint == null)
            {
                return rule;
            }

            rule.ErrorSource = constraint.ErrorSource;
            rule.MpeSource = constraint.MpeSource;
            rule.ResultSource = constraint.ResultSource;
            rule.ErrorType = constraint.ErrorType;
            rule.FixedMpe = constraint.FixedMpe;
            rule.FixedNegativeTolerance = constraint.FixedNegativeTolerance;
            rule.FixedPositiveTolerance = constraint.FixedPositiveTolerance;
            rule.RequirementOperator = constraint.RequirementOperator;
            rule.ErrorFormula = constraint.ErrorFormula;
            return rule;
        }

        public static MeasurementJudgementConstraint Capture(MeasurementRule rule, MeasurementJudgementConstraint original)
        {
            if (rule == null)
            {
                return original;
            }

            return new MeasurementJudgementConstraint
            {
                Name = original?.Name ?? rule.ErrorSource?.Name ?? string.Empty,
                ErrorSource = rule.ErrorSource,
                MpeSource = rule.MpeSource,
                ResultSource = rule.ResultSource,
                ErrorType = rule.ErrorType,
                FixedMpe = rule.FixedMpe,
                FixedNegativeTolerance = rule.FixedNegativeTolerance,
                FixedPositiveTolerance = rule.FixedPositiveTolerance,
                RequirementOperator = rule.RequirementOperator,
                ErrorFormula = rule.ErrorFormula
            };
        }

        public static MeasurementJudgementConstraint Clone(MeasurementJudgementConstraint constraint)
        {
            return MeasurementRuleCloner.CloneJudgementConstraint(constraint);
        }

        public static List<MeasurementJudgementConstraint> CloneList(
            IEnumerable<MeasurementJudgementConstraint> constraints)
        {
            return MeasurementRuleCloner.CloneJudgementConstraints(constraints);
        }

        public static bool HasValidRange(CellRange range)
        {
            return range != null &&
                range.StartRow > 0 &&
                range.StartColumn > 0 &&
                range.EndRow >= range.StartRow &&
                range.EndColumn >= range.StartColumn;
        }

        private static bool HasPrimaryConstraint(MeasurementRule rule)
        {
            return HasValidRange(rule.ErrorSource?.Range) ||
                HasValidRange(rule.MpeSource?.Range) ||
                HasValidRange(rule.ResultSource?.Range);
        }
    }
}
