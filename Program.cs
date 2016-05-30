using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace zxc.excel.proc
{

	class Program
	{
		static Dictionary<string, int> STW_dict;

		static Program()
		{
			STW_dict = new Dictionary<string, int>();
		}

		static void Main(string[] args)
		{
			string mode = "insert";
			Boolean bShowPrefix = true;

			if(args.Count() < 4 || args.Count() > 6)
			{
				Console.WriteLine("\nUsage:\n\tzxc.excel.proc.EXE inputFilePath outputFilePath intBaseSTW intBaseTWT [insert|update] bShowPrefix [true|false]\n");
				return;
			}

			if(args.Count() > 4)
			{
				if(args[4].ToLower() == "insert")
					mode = "insert";
				else if(args[4].ToLower() == "update")
					mode = "update";
				else
				{
					Console.WriteLine("Incorrect mode {0}", args[4]);
					return;
				}
			if(args.Count() > 5 && args[5].ToLower() == "false")
			{
				bShowPrefix = false;
			}
			}

			Console.WriteLine("input  file: {0}", args[0]);
			Console.WriteLine("output file: {0}", args[1]);
			Console.WriteLine("base number for STW: {0}", int.Parse(args[2]));
			Console.WriteLine("base number for TWT: {0}", int.Parse(args[3]));
			Console.WriteLine("SQL Generation mode: {0}", mode);
			Console.WriteLine("present prefix in SQL: {0}", bShowPrefix.ToString());

			Reader rdr = new Reader(args[0], "toSOS_S_TEKSTOW");
			Dicts dicts = new Dicts(bShowPrefix);

			int cnt = rdr.ReadSTW(dicts, int.Parse(args[2]), int.Parse(args[3]));

			Console.WriteLine("Counter={0}", cnt);

			//Console.WriteLine("{0}", dicts.ToSqlString());

			// kodowanie polskich znakow w skrypcie : Ansi Windows-1250 
			System.IO.StreamWriter writer = new System.IO.StreamWriter( args[1], false, Encoding.GetEncoding(1250) );
			if(mode == "insert")
				writer.Write( dicts.to_insert_string() );
			else if(mode == "update")
				writer.Write( dicts.to_update_string() );
			else
				Console.WriteLine("Nieznany sql mode.");

			writer.Close();

			Console.Write("Koniec, naciś entera ...");
			Console.ReadLine();
		}
	}
}
