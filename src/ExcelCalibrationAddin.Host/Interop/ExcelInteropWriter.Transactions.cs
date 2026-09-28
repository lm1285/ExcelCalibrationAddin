using System;
using System.Collections.Generic;
using System.Linq;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Host.UseCases;

namespace ExcelCalibrationAddin.Host.Interop
{
    public sealed partial class ExcelInteropWriter : ITransactionalWorkbookWriter
    {
        public IWorkbookWriteTransaction BeginTransaction(IReadOnlyList<RulePreview> previews)
        {
            return new ExcelWriteTransaction(_workbook, previews);
        }

        private sealed class ExcelWriteTransaction : IWorkbookWriteTransaction
        {
            private readonly dynamic _workbook;
            private readonly List<CellState> _cells;
            private readonly ApplicationState _applicationState;
            private bool _committed;
            private bool _disposed;

            public ExcelWriteTransaction(dynamic workbook, IReadOnlyList<RulePreview> previews)
            {
                _workbook = workbook;
                _cells = CaptureCells(workbook, previews);
                _applicationState = CaptureApplicationState(workbook);
                try
                {
                    SetSafeTrialState(workbook, _applicationState);
                }
                catch
                {
                    RestoreApplicationState(workbook, _applicationState);
                    VerifyApplicationState(_applicationState);
                    throw;
                }
            }

            public void Commit()
            {
                _committed = true;
            }

            public void Dispose()
            {
                if (_disposed) return;
                _disposed = true;
                try
                {
                    if (!_committed)
                    {
                        RestoreCells(_workbook, _cells);
                        VerifyCells(_workbook, _cells);
                    }
                }
                finally
                {
                    RestoreApplicationState(_workbook, _applicationState);
                    VerifyApplicationState(_applicationState);
                }
            }

            private static List<CellState> CaptureCells(dynamic workbook, IReadOnlyList<RulePreview> previews)
            {
                var addresses = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var states = new List<CellState>();
                foreach (var preview in previews ?? new List<RulePreview>())
                {
                    var cells = preview?.WritableCells ?? new List<CellAddress>();
                    if (cells.Count == 0 && preview?.TargetRange != null)
                    {
                        for (var row = preview.TargetRange.StartRow; row <= preview.TargetRange.EndRow; row++)
                        for (var column = preview.TargetRange.StartColumn; column <= preview.TargetRange.EndColumn; column++)
                            cells = cells.Concat(new[] { new CellAddress { Row = row, Column = column } }).ToList();
                    }

                    foreach (var address in cells.Where(item => item != null))
                    {
                        var key = (preview.TargetRange?.SheetName ?? string.Empty) + "!" + address.Row + ":" + address.Column;
                        if (!addresses.Add(key)) continue;
                        dynamic worksheet = workbook.Worksheets[preview.TargetRange.SheetName];
                        dynamic cell = worksheet.Cells[address.Row, address.Column];
                        states.Add(new CellState
                        {
                            SheetName = preview.TargetRange.SheetName,
                            Row = address.Row,
                            Column = address.Column,
                            Value = TryGet(cell, "Value2"),
                            Formula = Convert.ToString(TryGet(cell, "Formula")) ?? string.Empty,
                            FormulaR1C1 = Convert.ToString(TryGet(cell, "FormulaR1C1")) ?? string.Empty,
                            NumberFormat = Convert.ToString(TryGet(cell, "NumberFormat")) ?? string.Empty
                        });
                    }
                }
                return states;
            }

            private static void RestoreCells(dynamic workbook, IEnumerable<CellState> states)
            {
                foreach (var state in states ?? Enumerable.Empty<CellState>())
                {
                    dynamic worksheet = workbook.Worksheets[state.SheetName];
                    dynamic cell = worksheet.Cells[state.Row, state.Column];
                    if (!string.IsNullOrWhiteSpace(state.Formula))
                    {
                        cell.Formula = state.Formula;
                        if (!string.IsNullOrWhiteSpace(state.FormulaR1C1)) cell.FormulaR1C1 = state.FormulaR1C1;
                    }
                    else
                    {
                        cell.Value2 = state.Value;
                    }
                    if (!string.IsNullOrWhiteSpace(state.NumberFormat)) cell.NumberFormat = state.NumberFormat;
                }
            }

            private static void VerifyCells(dynamic workbook, IEnumerable<CellState> states)
            {
                foreach (var state in states ?? Enumerable.Empty<CellState>())
                {
                    dynamic worksheet = workbook.Worksheets[state.SheetName];
                    dynamic cell = worksheet.Cells[state.Row, state.Column];
                    var value = TryGet(cell, "Value2");
                    var formula = Convert.ToString(TryGet(cell, "Formula")) ?? string.Empty;
                    var format = Convert.ToString(TryGet(cell, "NumberFormat")) ?? string.Empty;
                    if (!ValuesEqual(value, state.Value) ||
                        !string.Equals(formula, state.Formula, StringComparison.Ordinal) ||
                        !string.Equals(format, state.NumberFormat, StringComparison.Ordinal))
                    {
                        throw new InvalidOperationException(
                            $"Excel 试算恢复校验失败：{state.SheetName}!R{state.Row}C{state.Column} 的值、公式或数字格式与试算前不同。");
                    }
                }
            }

            private static bool ValuesEqual(object left, object right)
            {
                if (left == null || right == null) return left == null && right == null;
                if (left is Array leftArray && right is Array rightArray)
                    return leftArray.Cast<object>().SequenceEqual(rightArray.Cast<object>());
                return string.Equals(Convert.ToString(left, System.Globalization.CultureInfo.InvariantCulture),
                    Convert.ToString(right, System.Globalization.CultureInfo.InvariantCulture), StringComparison.Ordinal);
            }

            private static object TryGet(dynamic value, string property)
            {
                try { return value == null ? null : value.GetType().InvokeMember(property, System.Reflection.BindingFlags.GetProperty, null, value, null); }
                catch { return null; }
            }

            private static ApplicationState CaptureApplicationState(dynamic workbook)
            {
                dynamic app = null;
                try { app = workbook?.Application; } catch { }
                return new ApplicationState
                {
                    Application = app,
                    EnableEvents = TryGet(app, "EnableEvents"),
                    DisplayAlerts = TryGet(app, "DisplayAlerts"),
                    ScreenUpdating = TryGet(app, "ScreenUpdating"),
                    Calculation = TryGet(app, "Calculation"),
                    AskToUpdateLinks = TryGet(app, "AskToUpdateLinks"),
                    CalculateBeforeSave = TryGet(app, "CalculateBeforeSave")
                };
            }

            private static void SetSafeTrialState(dynamic workbook, ApplicationState state)
            {
                if (state?.Application == null)
                    throw new InvalidOperationException("无法取得 Excel Application 状态，已拒绝启动试算事务。");
                state.Application.EnableEvents = false;
                state.Application.DisplayAlerts = false;
                state.Application.ScreenUpdating = false;
                state.Application.AskToUpdateLinks = false;
                state.Application.CalculateBeforeSave = false;
                if (Convert.ToBoolean(TryGet(state.Application, "EnableEvents")) ||
                    Convert.ToBoolean(TryGet(state.Application, "DisplayAlerts")) ||
                    Convert.ToBoolean(TryGet(state.Application, "ScreenUpdating")) ||
                    Convert.ToBoolean(TryGet(state.Application, "AskToUpdateLinks")) ||
                    Convert.ToBoolean(TryGet(state.Application, "CalculateBeforeSave")))
                    throw new InvalidOperationException("无法确认 Excel 已进入受控试算状态。");
            }

            private static void RestoreApplicationState(dynamic workbook, ApplicationState state)
            {
                if (state?.Application == null) return;
                try { if (state.EnableEvents != null) state.Application.EnableEvents = state.EnableEvents; } catch { }
                try { if (state.DisplayAlerts != null) state.Application.DisplayAlerts = state.DisplayAlerts; } catch { }
                try { if (state.ScreenUpdating != null) state.Application.ScreenUpdating = state.ScreenUpdating; } catch { }
                try { if (state.Calculation != null) state.Application.Calculation = state.Calculation; } catch { }
                try { if (state.AskToUpdateLinks != null) state.Application.AskToUpdateLinks = state.AskToUpdateLinks; } catch { }
                try { if (state.CalculateBeforeSave != null) state.Application.CalculateBeforeSave = state.CalculateBeforeSave; } catch { }
            }

            private static void VerifyApplicationState(ApplicationState state)
            {
                if (state?.Application == null) return;
                VerifyProperty(state, "EnableEvents", state.EnableEvents);
                VerifyProperty(state, "DisplayAlerts", state.DisplayAlerts);
                VerifyProperty(state, "ScreenUpdating", state.ScreenUpdating);
                VerifyProperty(state, "Calculation", state.Calculation);
                VerifyProperty(state, "AskToUpdateLinks", state.AskToUpdateLinks);
                VerifyProperty(state, "CalculateBeforeSave", state.CalculateBeforeSave);
            }

            private static void VerifyProperty(ApplicationState state, string property, object expected)
            {
                if (expected == null) return;
                var actual = TryGet(state.Application, property);
                if (!ValuesEqual(actual, expected))
                    throw new InvalidOperationException($"Excel 试算后 {property} 状态恢复校验失败。");
            }

            private sealed class CellState
            {
                public string SheetName { get; set; }
                public int Row { get; set; }
                public int Column { get; set; }
                public object Value { get; set; }
                public string Formula { get; set; }
                public string FormulaR1C1 { get; set; }
                public string NumberFormat { get; set; }
            }

            private sealed class ApplicationState
            {
                public dynamic Application { get; set; }
                public object EnableEvents { get; set; }
                public object DisplayAlerts { get; set; }
                public object ScreenUpdating { get; set; }
                public object Calculation { get; set; }
                public object AskToUpdateLinks { get; set; }
                public object CalculateBeforeSave { get; set; }
            }
        }
    }
}
