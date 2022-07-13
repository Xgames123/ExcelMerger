using Cocona;
using ExcelMerger;
using System.Diagnostics;

CoconaLiteApp.Run<App>();

public class App
{
	[PrimaryCommand]
	public void Merge(
		[Argument(Description ="Path to the directory filled with *.xlsx that need to be merged")]
		string inputFiles,
		[Option(Description ="Path+name of the output file")]
		string outputFile="ExcelFileMerger_output.xlsx",

		[Option(Description = "If set threat the input-files argument as a list of file paths separated by ';'")]
		bool nodir=false,
		[Option(Description = "Answers yes to all prompts")]
		bool noprompts=false,
		[Option(Description = "If set don't open the output file when done processing")]
		bool noopenOutput=false,
		[Option(Description = "If set remove the line at the top of the excel document with the names of all the files")]
		bool noheader=false,
		[Option(Description = "If set don't adjust the with of the columns to the content size")]
		bool noadjustToContent=false,
		[Option(Description = "If set show the internal errors")]
		bool showInternalErrors=false
		)
	{
		ConsoleWriter.ShowInternalErrors = showInternalErrors;
		EasterEgg.EnableAtRandom();
		ConsoleWriter.WriteBanner();

		var fullOutFile = Path.GetFullPath(outputFile);

		FileInfo[] sourceFiles;

		if (nodir)
		{
			var sourceFilesStrArray = inputFiles.Split(';');

			var sourceFilesList = new List<FileInfo>();
			foreach (var sourceFileStr in sourceFilesStrArray)
			{
				try
				{
					var info = new FileInfo(sourceFileStr);
					if (!info.Exists)
					{
						ConsoleWriter.WriteLineError($"File '{info.FullName}' does not exist");
					}
					else
					{
						sourceFilesList.Add(info);
					}
				}
				catch (Exception e)
				{
					ConsoleWriter.WriteLineError($"Could not read input file '{sourceFileStr}'", e);
				}


			}

			sourceFiles = sourceFilesList.ToArray();


		}
		else
		{
			var sourceDir = new DirectoryInfo(inputFiles);
			if (!sourceDir.Exists)
			{
				ConsoleWriter.WriteLineError($"input directory '{sourceDir.FullName}' does not exist or is not a directory");
				Environment.Exit(-1);
			}
			sourceFiles = sourceDir.GetFiles();

			
		}

		if (sourceFiles.Length == 0)
		{

			ConsoleWriter.WriteLineError("There are no files read. Exiting...");
			Environment.Exit(-1);

		}
		
		


		if (File.Exists(fullOutFile) && !noprompts)
		{
			if(!ConsoleWriter.AskForYesNo("Destination file already exist. Do you want to overwrite it"))
			{
				Environment.Exit(-1);
			}
		}


		FileMerger.StartMerge(sourceFiles, fullOutFile, !noheader, !noadjustToContent);

		if (!noopenOutput)
		{
			ConsoleWriter.WriteLineStatus("Opening output file");
			try
			{
				Process.Start("explorer", $"\"{fullOutFile}\"").Dispose();
			}catch(Exception e)
			{
				ConsoleWriter.WriteLineError("Could not start process to open the output file", e);
			}
			
		}

	}

}

