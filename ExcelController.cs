using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace MPL1 {
	class ExcelController {
		private Excel.Application excelApp;
		private Excel.Workbook workbook;
		private void TerminateExcel() {
			workbook.Close(0);
			excelApp.Quit();
		}
		private int CreatePostcards(Excel.Worksheet worksheet, ContentController contentController) {
			int failedInsertions = 0;

			Excel.Range lastCell = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
			int lastRow = lastCell.Row;

			Excel.Range data = worksheet.Range[worksheet.Cells[2, "A"], worksheet.Cells[lastRow, "B"]];

			for (int i = 1; i <= data.Rows.Count; i++) {
				if (!contentController.AddNewPostCard(data.Cells[i, "A"].Value2.ToString(), data.Cells[i, "C"].Value2.ToString(), data.Cells[i, "B"].Value2.ToString())) {
					failedInsertions++;
				}
			}

			return failedInsertions;
		}

		private bool SetInfo(Excel.Worksheet worksheet, WordController wordController) {
			Excel.Range lastCell = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
			int lastRow = lastCell.Row;
			Excel.Range settingsNameRange = worksheet.Range["A2", worksheet.Cells[lastRow, "A"]];

			//Взятие шрифта
			string fontName = worksheet.Cells[settingsNameRange.Find("font").Row, "B"].Font.Name;
			wordController.SetFont(fontName);

			return true;
		}

		private void AddDictionary(Excel.Worksheet worksheet, ContentController contentController) {
			Excel.Range lastCell = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
			int lastRow = lastCell.Row;
			int lastColumn = lastCell.Column;

			if (!(lastRow > 1)){
				return;
			}

			Excel.Range data = worksheet.Range[worksheet.Cells[2, "A"], worksheet.Cells[lastRow, lastColumn]];
			for (int i = 1; i <= data.Columns.Count; i++) {
				for (int y = 1; y <= data.Rows.Count; y++) {
					if (data.Cells[y, i].Value == null) {
						continue;
					}

					contentController.AddNewPhrase(i, data.Cells[y, i].Value2.ToString());
				}
			}
		}

		private void AddCelebrations(Excel.Worksheet worksheet, ContentController contentController) {
			Excel.Range lastCell = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
			int lastRow = lastCell.Row;

			if (!(lastRow > 1)) {
				return;
			}

			Excel.Range data = worksheet.Range[worksheet.Cells[2, "A"], worksheet.Cells[lastRow, "C"]];
			for (int i = 1; i <= data.Rows.Count; i++) {
				if (data.Cells[i, "A"].Value == null || data.Cells[i, "B"].Value == null || data.Cells[i, "C"].Value == null) {
					continue;
				}

				contentController.AddNewCelebration(data.Cells[i, "A"].Value2.ToString(), data.Cells[i, "B"].Value2.ToString(), data.Cells[i, "C"].Value2.ToString());
			}
		}


		public void GetInfo(ContentController contentController, WordController wordController) {
			excelApp = new Excel.Application();

			try {
				workbook = excelApp.Workbooks.Open($@"{Directory.GetCurrentDirectory()}\settings.xlsx");
			} catch (Exception e){
				TerminateExcel();
				throw new Exception(e.Message);
			}

			//Считывание настроек
			Excel.Worksheet settingsWorksheet = workbook.Worksheets["Настройки"];
			SetInfo(settingsWorksheet, wordController);

			Excel.Worksheet celebrationsWorksheet = workbook.Worksheets["Праздники"];
			AddCelebrations(celebrationsWorksheet, contentController);

			Excel.Worksheet namesWorksheet = workbook.Worksheets["Имена"];
			CreatePostcards(namesWorksheet, contentController);

			Excel.Worksheet congratGroupsWorksheet = workbook.Worksheets["Поздравления"];
			AddDictionary(congratGroupsWorksheet, contentController);

			TerminateExcel();
		}
	}
}
