using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace ExcelCalibrationAddin.Vsto
{
    [ComVisible(true)]
    public sealed class CalibrationRibbonXml : Office.IRibbonExtensibility
    {
        private const string RibbonXml = @"
<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" onLoad=""Ribbon_Load"">
  <commands>
    <command idMso=""Undo"" enabled=""true"" onAction=""OnUndo"" />
  </commands>
  <ribbon>
    <tabs>
      <tab idMso=""TabAddIns"">
        <group id=""CalibrationAssistantGroup"" label=""校准助手"">
          <labelControl id=""lblAddinVersion"" getLabel=""GetLabel"" />
          <labelControl id=""lblServiceConnection"" getLabel=""GetLabel"" />
          <button id=""btnQuickGenerate"" label=""生成随机数"" size=""large"" imageMso=""CalculateNow"" onAction=""OnButtonAction"" />
          <button id=""btnRecognize"" label=""识别模板"" size=""large"" imageMso=""TableDesign"" onAction=""OnButtonAction"" />
          <button id=""btnTemplateLibrary"" label=""模板库管理"" imageMso=""FileDocumentManageVersions"" onAction=""OnButtonAction"" />
          <button id=""btnTogglePane"" label=""侧边栏"" imageMso=""NavigationPane"" onAction=""OnButtonAction"" />
        </group>
        <group id=""RandomConfigurationGroup"" label=""随机数配置"">
          <box id=""RandomSummaryBox"" boxStyle=""vertical"">
            <labelControl id=""lblRandomRangeTitle"" getLabel=""GetLabel"" />
            <labelControl id=""lblRandomRangeDetail"" getLabel=""GetLabel"" />
          </box>
          <box id=""SingleUseOverrideBox"" boxStyle=""vertical"">
            <comboBox id=""cboOverrideRule"" label=""校准项"" getText=""GetText"" onChange=""OnTextChanged"" getItemCount=""GetItemCount"" getItemLabel=""GetItemLabel"" />
            <editBox id=""edtOverrideRange"" label=""系数区间"" getText=""GetText"" onChange=""OnTextChanged"" />
            <editBox id=""edtOverrideDecimals"" label=""小数位数"" getText=""GetText"" onChange=""OnTextChanged"" />
          </box>
          <button id=""btnRandomConfig"" label=""配置"" imageMso=""DefineName"" onAction=""OnButtonAction"" />
        </group>
        <group id=""AlarmValueGroup"" label=""报警值输入"">
          <editBox id=""edtAlarmValue"" label=""报警值"" getText=""GetText"" onChange=""OnTextChanged"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

        private readonly ThisAddIn _addIn;
        private readonly CalibrationRibbon _state;
        private Office.IRibbonUI _ribbon;

        public CalibrationRibbonXml(ThisAddIn addIn, CalibrationRibbon state)
        {
            _addIn = addIn ?? throw new ArgumentNullException(nameof(addIn));
            _state = state ?? throw new ArgumentNullException(nameof(state));
        }

        public string GetCustomUI(string ribbonId)
        {
            Trace.WriteLine($"[VSTO] Native Undo Ribbon XML requested. RibbonId={ribbonId}");
            return RibbonXml;
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUi)
        {
            _ribbon = ribbonUi;
            Trace.WriteLine("[VSTO] Native Undo Ribbon loaded.");
            _state.InitializeXmlState();
        }

        public string GetLabel(Office.IRibbonControl control)
        {
            switch (control?.Id)
            {
                case "lblAddinVersion": return _state.lblAddinVersion.Label;
                case "lblServiceConnection": return _state.lblServiceConnection.Label;
                case "lblRandomRangeTitle": return _state.lblRandomRangeTitle.Label;
                case "lblRandomRangeDetail": return _state.lblRandomRangeDetail.Label;
                default: return string.Empty;
            }
        }

        public string GetText(Office.IRibbonControl control)
        {
            switch (control?.Id)
            {
                case "cboOverrideRule": return _state.cboOverrideRule.Text;
                case "edtOverrideRange": return _state.edtOverrideRange.Text;
                case "edtOverrideDecimals": return _state.edtOverrideDecimals.Text;
                case "edtAlarmValue": return _state.edtAlarmValue.Text;
                default: return string.Empty;
            }
        }

        public void OnTextChanged(Office.IRibbonControl control, string text)
        {
            switch (control?.Id)
            {
                case "cboOverrideRule": _state.cboOverrideRule.Text = text ?? string.Empty; break;
                case "edtOverrideRange": _state.edtOverrideRange.Text = text ?? string.Empty; break;
                case "edtOverrideDecimals": _state.edtOverrideDecimals.Text = text ?? string.Empty; break;
                case "edtAlarmValue": _state.edtAlarmValue.Text = text ?? string.Empty; break;
            }
        }

        public int GetItemCount(Office.IRibbonControl control)
        {
            return _state.cboOverrideRule.Items.Count;
        }

        public string GetItemLabel(Office.IRibbonControl control, int index)
        {
            return index >= 0 && index < _state.cboOverrideRule.Items.Count
                ? _state.cboOverrideRule.Items[index].Label
                : string.Empty;
        }

        public void OnButtonAction(Office.IRibbonControl control)
        {
            _state.InvokeXmlButton(control?.Id);
        }

        public bool GetUndoEnabled(Office.IRibbonControl control)
        {
            var enabled = _addIn.CanUndoLastGeneration();
            Trace.WriteLine($"[VSTO] Native Undo enabled callback. Enabled={enabled}");
            return enabled;
        }

        public void OnUndo(Office.IRibbonControl control, ref bool cancelDefault)
        {
            var canUndoGeneration = _addIn.CanUndoLastGeneration();
            Trace.WriteLine($"[VSTO] Native Excel undo command invoked. CanUndoGeneration={canUndoGeneration}");
            if (!canUndoGeneration)
            {
                cancelDefault = false;
                return;
            }

            cancelDefault = true;
            _addIn.UndoLastGeneration();
        }

        internal void InvalidateAll()
        {
            _ribbon?.Invalidate();
        }

        internal void InvalidateUndo()
        {
            _ribbon?.InvalidateControlMso("Undo");
        }
    }
}
