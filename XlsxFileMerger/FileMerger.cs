using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace ExcelMerger;
public static class FileMerger
{

	private static Dictionary<XLWorkbook, string> c_workbookToFilenameDict = new Dictionary<XLWorkbook, string>();

	public static void StartMerge(FileInfo[] sourceFiles, string outFile, bool addHeader=true, bool adjustToContents=true)
	{
		if (EasterEgg.Enabled)
		{
			ConsoleWriter.WriteLineStatus("Starting magic...");
		}
		else
		{
			ConsoleWriter.WriteLineStatus("Starting merge...");
		}
		

		using var destXlsx = new XLWorkbook();
		var destSheet = destXlsx.AddWorksheet("Sheet1");

		List<IXLColumn> firstSourceColumns = new List<IXLColumn>();
		List<IXLWorksheet> workSheets = new List<IXLWorksheet>();

		var workbooks = ReadWorkbooks(sourceFiles);
		foreach (var book in workbooks)
		{

			if (book.Worksheets.Count == 0)
			{
				ConsoleWriter.WriteLineError($"File '{c_workbookToFilenameDict[book]}' has no worksheets in it");
				return;
			}
			var sheet = book.Worksheets.First();
			workSheets.Add(sheet);
			firstSourceColumns.Add(sheet.Column(1));

			
		}
		int end = 1;
		if (addHeader)
			end = 2;
		var contentToIndexDict = MergeColumnsAndBuildCTIDict(firstSourceColumns.ToArray(), destSheet.Column(1), end);
		CopyData(workSheets, destSheet, contentToIndexDict);


		workbooks.ForEach((book) => book.Dispose());
		workbooks.Clear();

		if (adjustToContents)
		{
			ConsoleWriter.WriteLineStatus("Adjusting columns to content");
			destSheet.ColumnsUsed().AdjustToContents();
		}
		

		try
		{
			destXlsx.SaveAs(outFile);
		}catch(Exception e)
		{
			ConsoleWriter.WriteLineError("Could not save output file. Maybe the file is open in excel?", e);
		}
		
	}


	private static List<XLWorkbook> ReadWorkbooks(FileInfo[] files)
	{
		List<XLWorkbook> books = new List<XLWorkbook>();
		foreach (var file in files)
		{
			ConsoleWriter.WriteLineStatus($"Reading file {file.FullName} ...");
			try
			{
				var book = new XLWorkbook(file.FullName);
				c_workbookToFilenameDict.Add(book, file.FullName);
				books.Add(book);
			}
			catch (Exception e)
			{
				ConsoleWriter.WriteLineError($"Could not read file '{file.FullName}'. This could happen when the file is open in another process. Or when the file is in onedrive and not downloaded", e);
				ConsoleWriter.WriteLineStatus("Continuing anyway");
			}

		}
		if (books.Count == 0)
		{
			ConsoleWriter.WriteLineError("No files were read. Exiting...");
			Environment.Exit(-1);
		}
		return books;

	}



	private static Dictionary<string, int> MergeColumnsAndBuildCTIDict(IXLColumn[] colls, IXLColumn dest, int destStartIndex=1)
	{
		int destIndex = destStartIndex;

		var contentToIndexDict = new Dictionary<string, int>();

		List<IEnumerator<IXLCell>> CollEnumerators = new List<IEnumerator<IXLCell>>();
		foreach (var col in colls)
		{
			CollEnumerators.Add(col.CellsUsed().GetEnumerator());
		}

		while (true)
		{
			bool unskippedFile = false;
			foreach (var collEnumerator in CollEnumerators)
			{
				if (!collEnumerator.MoveNext())
				{
					continue;
				}
				unskippedFile = true;
				var cell = collEnumerator.Current;


				if (Find(dest, (string)cell.Value, 1, destIndex) == null)
				{
					dest.Cell(destIndex).Value = cell.Value;
					contentToIndexDict.Add((string)cell.Value, destIndex);
					destIndex++;
				}
			}
			if (!unskippedFile)
			{
				break;
			}

			ConsoleWriter.WriteLineStatus($"Done with pass. Total lines written: {destIndex}");
		}
		
		return contentToIndexDict;
	}

	private static IXLCell? Find(IXLColumn column, string cellValue, int start, int end)
	{
		for (int i = start; i < end; i++)
		{
			var cell = column.Cell(i);
			if (cellValue == (string)cell.Value)
			{
				return cell;
			}

		}
		return null;
	}

	

	private static void CopyData(List<IXLWorksheet> sheets, IXLWorksheet dest, Dictionary<string, int> contentToIndexDict)
	{
		ConsoleWriter.WriteLineStatus("Copying data...");

		int destIndex = 2;

		for (int i = 0; i < sheets.Count; i++)
		{
			var sheet = sheets[i];
			var dataCol = sheet.Column(2);

			var fileName = Path.GetFileName(c_workbookToFilenameDict[sheet.Workbook]);
			dest.Cell(1, destIndex).Value = fileName;


			foreach (var cell in dataCol.CellsUsed())
			{
				var key = (string)sheet.Cell(cell.Address.RowNumber, 1).Value;

				var index = contentToIndexDict.GetValueOrDefault(key, -1);
				if (index == -1)
				{
					ConsoleWriter.WriteLineError($"Empty cell is at {cell.Address} in file '{fileName}'");
					continue;
				}
					
				dest.Cell(index, destIndex).Value = cell.Value;
			}

			ConsoleWriter.WriteLineStatus($"Copied sheet {i+1}/{sheets.Count}");
			destIndex++;
		}


	}



}
