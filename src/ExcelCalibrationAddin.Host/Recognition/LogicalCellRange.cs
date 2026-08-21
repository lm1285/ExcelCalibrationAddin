using ExcelCalibrationAddin.Contracts;

namespace ExcelCalibrationAddin.Host.Recognition
{
    public sealed class LogicalCellRange
    {
        public CellMeta Anchor { get; set; }
        public CellRange Range { get; set; }
    }
}
