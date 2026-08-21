using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Models;
using ExcelCalibrationAddin.Core.Services;
using ExcelCalibrationAddin.Host.ViewModels;
using ExcelCalibrationAddin.Host.Vsto;
using ExcelCalibrationAddin.Vsto.TaskPane;
using Microsoft.Office.Tools;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelCalibrationAddin.Vsto
{
    public partial class ThisAddIn
    {
        private const int PreferredTaskPaneWidth = 420;
        private const int MinimumTaskPaneWidth = 380;
        private const int MaximumTaskPaneWidth = 520;

        private static string SafeWorkbookName(Excel.Workbook workbook)
        {
            try
            {
                return Convert.ToString(workbook?.Name) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string ResolveWorkbookKey(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return string.Empty;
            }

            try
            {
                var fullName = Convert.ToString(workbook.FullName);
                if (!string.IsNullOrWhiteSpace(fullName))
                {
                    return fullName;
                }
            }
            catch
            {
            }

            try
            {
                return Convert.ToString(workbook.Name) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private void EnsureTaskPane()
        {
            if (IsTaskPaneUsable())
            {
                return;
            }

            try
            {
                if (_taskPane != null)
                {
                    this.CustomTaskPanes.Remove(_taskPane);
                }
            }
            catch
            {
            }

            _taskPaneControl = new TaskPane.CalibrationTaskPaneControl();
            try
            {
                _taskPane = this.CustomTaskPanes.Add(_taskPaneControl, TaskPaneTitle);
                _taskPane.Width = ResolveTaskPaneWidth();
                _taskPane.VisibleChanged += TaskPane_VisibleChanged;
                _taskPane.Visible = false;
            }
            catch (Exception ex)
            {
                _taskPaneControl.Dispose();
                _taskPaneControl = null;
                _taskPane = null;
                throw new InvalidOperationException("Failed to create the calibration task pane.", ex);
            }
        }

        private int ResolveTaskPaneWidth()
        {
            // CustomTaskPane.Width is expressed in points. Use a proportion
            // of the Excel window so large monitors get more editing room,
            // while retaining a minimum width for the fixed editor controls.
            int width = PreferredTaskPaneWidth;
            try
            {
                var activeWindow = Application?.ActiveWindow;
                if (activeWindow != null)
                {
                    width = (int)Math.Round(Convert.ToDouble(activeWindow.Width) * 0.30d);
                }
            }
            catch
            {
                // Keep the preferred width when Excel is not ready to expose
                // the active window during startup.
            }

            return Math.Max(MinimumTaskPaneWidth, Math.Min(MaximumTaskPaneWidth, width));
        }

        private void ApplyTaskPaneWidth()
        {
            if (!IsTaskPaneUsable())
            {
                return;
            }

            try
            {
                _taskPane.Width = ResolveTaskPaneWidth();
            }
            catch
            {
                // Excel can reject a resize while a workbook window is being
                // activated; the next activation will retry it.
            }
        }

        private void TaskPane_VisibleChanged(object sender, EventArgs e)
        {
            if (_restoringTaskPaneVisibility || _taskPane == null || _taskPane.Visible || _taskPaneControl == null)
            {
                return;
            }

            if (_taskPaneControl.ConfirmCloseIfDirty())
            {
                return;
            }

            try
            {
                _restoringTaskPaneVisibility = true;
                _taskPane.Visible = true;
            }
            finally
            {
                _restoringTaskPaneVisibility = false;
            }
        }

        private bool IsTaskPaneUsable()
        {
            if (_taskPane == null || _taskPaneControl == null || _taskPaneControl.IsDisposed)
            {
                return false;
            }

            try
            {
                var _ = _taskPane.Visible;
                return true;
            }
            catch
            {
                return false;
            }
        }

        private bool IsTaskPaneVisible()
        {
            EnsureTaskPane();

            try
            {
                return _taskPane.Visible;
            }
            catch
            {
                EnsureTaskPane();
                return _taskPane.Visible;
            }
        }

        private void SetTaskPaneVisible(bool visible)
        {
            EnsureTaskPane();

            if (!visible && !_taskPaneControl.ConfirmCloseIfDirty())
            {
                return;
            }

            try
            {
                _taskPane.Visible = visible;
            }
            catch
            {
                EnsureTaskPane();
                _taskPane.Visible = visible;
            }
        }

        private bool IsTaskPaneCurrentlyVisible()
        {
            if (_taskPane == null || _taskPaneControl == null || _taskPaneControl.IsDisposed)
            {
                return false;
            }

            try
            {
                return _taskPane.Visible;
            }
            catch
            {
                return false;
            }
        }

        private bool IsRangeFullyVisible(Excel.Range target)
        {
            try
            {
                var activeWindow = this.Application?.ActiveWindow;
                if (activeWindow == null)
                {
                    return false;
                }

                var visibleRange = activeWindow.VisibleRange as Excel.Range;
                if (visibleRange == null)
                {
                    return false;
                }

                var sameSheet = string.Equals(
                    Convert.ToString(((Excel.Worksheet)visibleRange.Worksheet).Name),
                    Convert.ToString(((Excel.Worksheet)target.Worksheet).Name),
                    StringComparison.OrdinalIgnoreCase);
                if (!sameSheet)
                {
                    return false;
                }

                return target.Row >= visibleRange.Row &&
                    target.Column >= visibleRange.Column &&
                    target.Row + target.Rows.Count - 1 <= visibleRange.Row + visibleRange.Rows.Count - 1 &&
                    target.Column + target.Columns.Count - 1 <= visibleRange.Column + visibleRange.Columns.Count - 1;
            }
            catch
            {
                return false;
            }
        }

    }
}
