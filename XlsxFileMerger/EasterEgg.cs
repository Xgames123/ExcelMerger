using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMerger;
public static class EasterEgg
{
	public static bool Enabled = false;

	

	public static void EnableAtRandom()
	{
		var rand = new Random(DateTime.Now.Millisecond);

		if (rand.Next(0, 30) == 3)
		{
			Enabled = true;
		}

	}

	public static void SetConsoleColor(ConsoleColor color)
	{
		Console.ForegroundColor = GetConsoleColor(color);
	}

	public static void SetConsoleColor()
	{
		if (!Enabled)
			return;

		Console.ForegroundColor = GetConsoleColor(ConsoleColor.White);
	}


	public static ConsoleColor GetConsoleColor(ConsoleColor color)
	{
		if (!Enabled)
			return color;

		switch(Random.Shared.Next(0, 7))
		{
			case 0:
				return ConsoleColor.Red;
			case 1:
				return ConsoleColor.Green;
			case 2:
				return ConsoleColor.Cyan;
			case 3:
				return ConsoleColor.Yellow;
			case 4:
				return ConsoleColor.Cyan;
			case 5:
				return ConsoleColor.Yellow;
			case 6:
				return ConsoleColor.Green;
		}

		return ConsoleColor.Cyan;
	}

}
