using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Host.Recognition;

namespace ExcelCalibrationAddin.Host.Services
{
    public sealed partial class MeasurementRuleDraftBuilder
    {
        private sealed class RefinedRange
        {
            public int StartColumn { get; set; }
            public int EndColumn { get; set; }
            public int HeaderBottomRow { get; set; }
        }

        private sealed class HeaderBand
        {
            public int TopRow { get; set; }
            public int BottomRow { get; set; }
        }

        private sealed class NumericHeaderRun
        {
            public int HeaderRow { get; set; }
            public int StartColumn { get; set; }
            public int EndColumn { get; set; }
            public int DataStartRow { get; set; }
            public int Score { get; set; }
        }

        private sealed class ColumnRun
        {
            public int StartColumn { get; set; }
            public int EndColumn { get; set; }
            public int Count { get; set; }
        }

        private sealed class ResultHeaderCandidate
        {
            public int StartColumn { get; set; }
            public int EndColumn { get; set; }
            public int DataStartRow { get; set; }
            public double Score { get; set; }
        }

        private sealed class HeaderCandidate
        {
            public int StartColumn { get; set; }
            public int EndColumn { get; set; }
            public int HeaderBottomRow { get; set; }
            public double Score { get; set; }
        }

    }
}
