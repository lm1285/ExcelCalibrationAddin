using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Host.ViewModels;

namespace ExcelCalibrationAddin.Vsto.TaskPane
{
    public partial class CalibrationTaskPaneControl
    {
	private static string BuildOverallStatusText(TaskPaneState state)
	{
		if (state == null)
		{
			return "等待识别";
		}

		if (state.IsFeatureBlocked || string.Equals(state.MatchStatus, "Failed", StringComparison.OrdinalIgnoreCase))
		{
			return "识别链路存在异常";
		}

		if (!string.IsNullOrWhiteSpace(state.RemoteTemplateName) || !string.IsNullOrWhiteSpace(state.RemoteTemplateId) || state.LocalTemplateStatus.HasValue)
		{
			return "模板库匹配成功";
		}

		return state.DraftRules != null && state.DraftRules.Count > 0
			? "新模板：已生成待确认草稿"
			: "未生成可用模板规则";
	}

	private static string BuildMatchChainText(TaskPaneState state)
	{
		if (state == null)
		{
			return string.Empty;
		}

		var lines = new List<string>
		{
			BuildChainStepText("识别", state.RecognitionStatusDetail),
			BuildChainStepText("指纹", state.FingerprintStatusDetail),
			BuildChainStepText("本地库", state.LocalMatchStatusDetail),
			BuildChainStepText("远端库", state.RemoteMatchStatusDetail),
			BuildChainStepText("规则", state.DraftRuleStatusDetail)
		};
		return string.Join(Environment.NewLine, lines);
	}

	private static string BuildChainStepText(string name, string detail)
	{
		string normalized = NormalizeChainDetail(detail);
		return ResolveChainIcon(normalized) + " " + name + "：" + BuildChainReason(normalized);
	}

	private static string ResolveChainIcon(string detail)
	{
		if (detail.StartsWith("通过", StringComparison.OrdinalIgnoreCase))
		{
			return "✓";
		}

		if (detail.StartsWith("异常", StringComparison.OrdinalIgnoreCase) ||
			detail.StartsWith("失败", StringComparison.OrdinalIgnoreCase))
		{
			return "✕";
		}

		if (detail.StartsWith("待确认", StringComparison.OrdinalIgnoreCase) ||
			detail.StartsWith("草稿", StringComparison.OrdinalIgnoreCase))
		{
			return "!";
		}

		return "○";
	}

	private static string BuildChainReason(string detail)
	{
		if (string.IsNullOrWhiteSpace(detail))
		{
			return "未执行";
		}

		var separatorIndex = detail.IndexOf('：');
		var reason = separatorIndex >= 0 && separatorIndex + 1 < detail.Length
			? detail.Substring(separatorIndex + 1)
			: detail;
		return reason.Trim();
	}

	private static string NormalizeChainDetail(string text)
	{
		return string.IsNullOrWhiteSpace(text) ? "未执行" : text.Trim();
	}

    }
}
