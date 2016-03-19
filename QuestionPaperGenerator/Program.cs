using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace QuestionPaperGenerator
{
	class Program
	{
		static void Main(string[] args)
		{
			Console.WriteLine("Opening workbook...");
			var questions = ParseQuestionItems();
			Console.WriteLine("Generating papers...");
			GenerateRandomPapers(questions, 25);
		}

		private static void GenerateRandomPapers(List<QuestionItem> questions, int questionCount = 25, int paperCount = 60, string outputDirectory = "Papers")
		{
			for (int i = 0; i < paperCount; i++)
			{
				var exportWorkbook = new XSSFWorkbook();
				var sheet1 = exportWorkbook.CreateSheet(i.ToString());

				//create header
				var headerRow = sheet1.CreateRow(0);
				headerRow.CreateCell(0).SetCellValue("答案（请填写A、B、C、D）");
				headerRow.CreateCell(1).SetCellValue("编号");
				headerRow.CreateCell(2).SetCellValue("问题");
				int maxOptionCount = questions.Max(q => q.Options.Count);
				for (int optionIndex = 0; optionIndex < maxOptionCount; optionIndex++)
				{
					headerRow.CreateCell(optionIndex + 3).SetCellValue("选项" + (char)('A' + optionIndex));
				}

				//randomize questions
				var rnd = new Random();
				var randomQuestions = questions.OrderBy(q => rnd.Next()).Take(questionCount).ToList();

				//generate question rows
				for (int rowIndex = 0; rowIndex < randomQuestions.Count; rowIndex++)
				{
					var newRow = sheet1.CreateRow(rowIndex + 1);
					var randomQuestion = randomQuestions[rowIndex];
					newRow.CreateCell(0, CellType.String).SetCellValue("");
					newRow.CreateCell(1, CellType.String).SetCellValue(randomQuestion.Id);
					newRow.CreateCell(2, CellType.String).SetCellValue(randomQuestion.Question);
					for (int columnIndex = 0; columnIndex < randomQuestion.Options.Count; columnIndex++)
					{
						newRow.CreateCell(columnIndex + 3, CellType.String).SetCellValue(randomQuestion.Options[columnIndex]);
					}
				}

				//sheet style
				sheet1.SetColumnHidden(1, true);
				sheet1.CreateFreezePane(1, 1);
				exportWorkbook.SetActiveSheet(0);

				//export to file
				Directory.CreateDirectory(outputDirectory);
				var filePath = Path.Combine(outputDirectory, i + ".xlsx");
				using (var fs = new FileStream(filePath, FileMode.Create))
				{
					exportWorkbook.Write(fs);
				}
			}
		}

		private static List<QuestionItem> ParseQuestionItems(string xlsxFilePath = "question.xlsx", string sheetName = "Sheet1")
		{
			var workbook = new XSSFWorkbook(File.OpenRead(xlsxFilePath));
			var sheet = workbook.GetSheet(sheetName);

			//read question from sheet1
			var questions = new List<QuestionItem>();
			for (int i = 0; i < sheet.LastRowNum + 1; i++)
			{
				var row = sheet.GetRow(i);
				var questionItem = new QuestionItem
				{
					Answer = row.Cells[0].ToString(),
					Id = row.Cells[1].ToString(),
					Question = row.Cells[2].ToString(),
					Options = Enumerable.Range(3, row.Cells.Count - 3).Select(columnIndex => row.Cells[columnIndex].ToString()).ToList()
				};
				questions.Add(questionItem);
			}
			return questions;
		}
	}
}
