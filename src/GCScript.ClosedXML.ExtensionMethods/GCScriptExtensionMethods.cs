using ClosedXML.Excel;
using GCScript.ClosedXML.ExtensionMethods.Models;

namespace GCScript.ClosedXML.ExtensionMethods;

public static class GCScriptExtensionMethods {
	public static IXLCell SetTwoDecimalPlacesFormat(this IXLCell cell) { cell.Style.NumberFormat.Format = "_-* #,##0.00_-;\\-* #,##0.00_-;_-* \"-\"??_-;_-@_-"; return cell; }
	public static IXLRange SetTwoDecimalPlacesFormat(this IXLRange rng) { rng.Style.NumberFormat.Format = "_-* #,##0.00_-;\\-* #,##0.00_-;_-* \"-\"??_-;_-@_-"; return rng; }
	public static IXLCell SetZeroDecimalPlacesFormat(this IXLCell cell) { cell.Style.NumberFormat.Format = "_-* #,##0_-;\\-* #,##0_-;_-* \"-\"_-;_-@_-"; return cell; }
	public static IXLRange SetZeroDecimalPlacesFormat(this IXLRange rng) { rng.Style.NumberFormat.Format = "_-* #,##0_-;\\-* #,##0_-;_-* \"-\"_-;_-@_-"; return rng; }
	public static IXLCell SetTextFormat(this IXLCell cell) { cell.Style.NumberFormat.Format = "@"; return cell; }
	public static IXLRange SetTextFormat(this IXLRange rng) { rng.Style.NumberFormat.Format = "@"; return rng; }
	public static IXLCell SetDateFormat(this IXLCell cell) { cell.Style.NumberFormat.Format = "dd/mm/yyyy"; return cell; }
	public static IXLRange SetDateFormat(this IXLRange rng) { rng.Style.NumberFormat.Format = "dd/mm/yyyy"; return rng; }
	public static IXLCell SetDateTimeFormat(this IXLCell cell) { cell.Style.NumberFormat.Format = "dd/mm/yyyy hh:mm"; return cell; }
	public static IXLRange SetDateTimeFormat(this IXLRange rng) { rng.Style.NumberFormat.Format = "dd/mm/yyyy hh:mm"; return rng; }

	public static IXLCell SetBackgroundColor(this IXLCell cell, string hexColor) { cell.Style.Fill.BackgroundColor = XLColor.FromHtml(hexColor); return cell; }
	public static IXLRange SetBackgroundColor(this IXLRange rng, string hexColor) { rng.Style.Fill.BackgroundColor = XLColor.FromHtml(hexColor); return rng; }
	public static IXLCell SetFontColor(this IXLCell cell, string hexColor) { cell.Style.Font.FontColor = XLColor.FromHtml(hexColor); return cell; }
	public static IXLRange SetFontColor(this IXLRange rng, string hexColor) { rng.Style.Font.FontColor = XLColor.FromHtml(hexColor); return rng; }
	public static IXLCell SetFontBold(this IXLCell cell, bool bold = true) { cell.Style.Font.Bold = bold; return cell; }
	public static IXLRange SetFontBold(this IXLRange rng, bool bold = true) { rng.Style.Font.Bold = bold; return rng; }

	public static IXLCell SetTextRotation(this IXLCell cell, int degrees) { cell.Style.Alignment.TextRotation = degrees; return cell; }
	public static IXLRange SetTextRotation(this IXLRange rng, int degrees) { rng.Style.Alignment.TextRotation = degrees; return rng; }
	public static IXLCell SetHorizontalTextAlignment(this IXLCell cell, XLAlignmentHorizontalValues alignment) { cell.Style.Alignment.Horizontal = alignment; return cell; }
	public static IXLRange SetHorizontalTextAlignment(this IXLRange rng, XLAlignmentHorizontalValues alignment) { rng.Style.Alignment.Horizontal = alignment; return rng; }
	public static IXLCell SetVerticalTextAlignment(this IXLCell cell, XLAlignmentVerticalValues alignment) { cell.Style.Alignment.Vertical = alignment; return cell; }
	public static IXLRange SetVerticalTextAlignment(this IXLRange rng, XLAlignmentVerticalValues alignment) { rng.Style.Alignment.Vertical = alignment; return rng; }
	public static IXLCell SetTextAlignment(this IXLCell cell, XLAlignmentHorizontalValues horizontal, XLAlignmentVerticalValues vertical) { cell.Style.Alignment.Horizontal = horizontal; cell.Style.Alignment.Vertical = vertical; return cell; }
	public static IXLRange SetTextAlignment(this IXLRange rng, XLAlignmentHorizontalValues horizontal, XLAlignmentVerticalValues vertical) { rng.Style.Alignment.Horizontal = horizontal; rng.Style.Alignment.Vertical = vertical; return rng; }

	public static IXLCell SetSuccessStyle(this IXLCell cell) { cell.SetBackgroundColor(GCSColors.Success.BackgroundColor).SetFontColor(GCSColors.Success.FontColor); return cell; }
	public static IXLRange SetSuccessStyle(this IXLRange rng) { rng.SetBackgroundColor(GCSColors.Success.BackgroundColor).SetFontColor(GCSColors.Success.FontColor); return rng; }
	public static IXLCell SetWarningStyle(this IXLCell cell) { cell.SetBackgroundColor(GCSColors.Warning.BackgroundColor).SetFontColor(GCSColors.Warning.FontColor); return cell; }
	public static IXLRange SetWarningStyle(this IXLRange rng) { rng.SetBackgroundColor(GCSColors.Warning.BackgroundColor).SetFontColor(GCSColors.Warning.FontColor); return rng; }
	public static IXLCell SetErrorStyle(this IXLCell cell) { cell.SetBackgroundColor(GCSColors.Error.BackgroundColor).SetFontColor(GCSColors.Error.FontColor); return cell; }
	public static IXLRange SetErrorStyle(this IXLRange rng) { rng.SetBackgroundColor(GCSColors.Error.BackgroundColor).SetFontColor(GCSColors.Error.FontColor); return rng; }

	public static void SetColumnWidth(this IXLWorksheet ws, int column, double width) => ws.Column(column).Width = width;
	public static void SetRowHeight(this IXLWorksheet ws, int row, double height) => ws.Row(row).Height = height;

	public static List<GCSColumnTitle> GetColumnTitles(this IXLWorksheet ws, int row = 1) {
		List<GCSColumnTitle> result = [];
		if (ws.IsEmpty()) return result;
		var lastColumnUsed = ws.LastColumnUsed();
		if (lastColumnUsed is null) return result;
		var lastColumnUsedColumnNumber = lastColumnUsed.ColumnNumber();
		for (int position = 1; position <= lastColumnUsedColumnNumber; position++) {
			var cell = ws.Cell(row, position);
			if (cell.IsEmpty()) continue;
			string title = cell.Value.ToString();
			if (string.IsNullOrWhiteSpace(title)) continue;
			result.Add(new GCSColumnTitle(Position: position, Title: title));
		}
		return result;
	}

	public static int GetPositionByTitle(this List<GCSColumnTitle> columnTitles, string title) {
		return columnTitles.FirstOrDefault(ct => ct.Title.Equals(title, StringComparison.OrdinalIgnoreCase))?.Position ?? 0;
	}

	public static List<GCSWorksheetTitle> GetWorksheetTitles(this IXLWorkbook wb) {
		List<GCSWorksheetTitle> result = [];
		foreach (var ws in wb.Worksheets) {
			string title = ws.Name;
			if (string.IsNullOrWhiteSpace(title)) continue;
			result.Add(new GCSWorksheetTitle(Position: ws.Position, Title: title));
		}
		return result;
	}

	public static int GetPositionByTitle(this List<GCSWorksheetTitle> worksheetTitles, string title) {
		return worksheetTitles.FirstOrDefault(wt => wt.Title.Equals(title, StringComparison.OrdinalIgnoreCase))?.Position ?? 0;
	}
}
