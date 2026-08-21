namespace ExcelCalibrationAddin.Host.Interop
{
    public static class ExcelAddressHelper
    {
        public static string ToColumnName(int columnNumber)
        {
            if (columnNumber <= 0)
            {
                return "A";
            }

            var dividend = columnNumber;
            var columnName = string.Empty;

            while (dividend > 0)
            {
                var modulo = (dividend - 1) % 26;
                columnName = (char)('A' + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }
    }
}
