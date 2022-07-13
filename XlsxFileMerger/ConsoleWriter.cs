using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMerger
{
	public static class ConsoleWriter
	{
		public static bool ShowInternalErrors;



		public static void WriteBanner()
		{
			var asmName = Assembly.GetExecutingAssembly().GetName();

			EasterEgg.SetConsoleColor();
			Console.WriteLine("*******************************************");
			EasterEgg.SetConsoleColor();
			Console.WriteLine($"{asmName.Name} V{asmName.Version}");
			EasterEgg.SetConsoleColor();
			Console.WriteLine("Made by XGames105");
			EasterEgg.SetConsoleColor();
			Console.WriteLine("*******************************************");

			Console.ResetColor();
		}


		public static void WriteCount(int value, int max)
		{
			Console.ResetColor();
			Console.WriteLine(value+"/"+max);

			Console.GetCursorPosition();
		}

		public static void WriteLineWarning(string value)
		{
			Console.ForegroundColor = ConsoleColor.Yellow;
			Console.WriteLine("WARNING: " + value);
			Console.ForegroundColor = ConsoleColor.White;
		}


		public static void WriteLineStatus(string value)
		{
			EasterEgg.SetConsoleColor(ConsoleColor.Green);
			Console.WriteLine("STATUS: " + value);
			Console.ResetColor();
		}

		public static void WriteLineError(string value, Exception? e=null)
		{
			Console.ForegroundColor = ConsoleColor.Red;
			Console.WriteLine("ERROR: "+value);

			if (!ShowInternalErrors)
			{
				Console.ResetColor();
				return;
			}


			if (e != null)
			{
				Console.WriteLine();
				Console.WriteLine("INTERNAL_ERROR: " + e.ToString());
			}
			Console.ResetColor();
		}


		public static bool AskForYesNo(string question)
		{
			Console.WriteLine(question+" (y/n)");
			while (true)
			{
				Console.Write(">");
				var awnser = Console.ReadLine().ToLower();
				if (awnser == "y" || awnser == "yes")
				{
					return true;

				}
				else if (awnser == "n" || awnser == "no")
				{
					return false;
				}
				Console.WriteLine($"'{awnser}' Was not a valid value (y/n)");

			}

		}


		public static string? AskForString(string question, Func<string?, bool>? validate=null)
		{
			while (true)
			{
				Console.WriteLine(question);
				Console.Write(">");

				var awnser = Console.ReadLine();

				if (validate == null)
				{
					return awnser;
				}
				if (validate.Invoke(awnser))
				{
					return awnser;
				}


				WriteLineError($"'{awnser}' was not valid");
			}
			

		}


	}
}
