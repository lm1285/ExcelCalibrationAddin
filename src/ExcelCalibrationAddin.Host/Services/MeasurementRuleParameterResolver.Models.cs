using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Host.Recognition;

namespace ExcelCalibrationAddin.Host.Services
{
    public sealed partial class MeasurementRuleParameterResolver
    {
        private sealed class NumericCandidate
        {
            public double Value { get; set; }
            public int Index { get; set; }
            public string Context { get; set; } = string.Empty;
        }

        private sealed class SignedNumericCandidate
        {
            public double Value { get; set; }
            public bool HasExplicitPositiveSign { get; set; }
        }

        private sealed class MpeCandidate
        {
            public ErrorType ErrorType { get; set; }
            public double Mpe { get; set; }
            public double? NegativeTolerance { get; set; }
            public double? PositiveTolerance { get; set; }
            public double? ReferenceRange { get; set; }
            public TechnicalRequirementOperator RequirementOperator { get; set; }
            public string ValuePattern { get; set; } = string.Empty;
            public int Row { get; set; }
            public int Column { get; set; }
            public double Score { get; set; }
        }

        private sealed class ResolvedMpe
        {
            public ErrorType ErrorType { get; set; }
            public double Mpe { get; set; }
            public double? NegativeTolerance { get; set; }
            public double? PositiveTolerance { get; set; }
            public double? ReferenceRange { get; set; }
            public TechnicalRequirementOperator RequirementOperator { get; set; }
            public string ValuePattern { get; set; } = string.Empty;
        }

    }
}
