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

			if(args.Count() < 4 || args.Count() > 5)
			{
				Console.WriteLine("\nUsage:\n\tzxc.excel.proc.EXE inputFilePath outputFilePath intBaseSTW intBaseTWT [insert|update]\n");
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
			}

			Console.WriteLine("input  file: {0}", args[0]);
			Console.WriteLine("output file: {0}", args[1]);
			Console.WriteLine("base number for STW: {0}", int.Parse(args[2]));
			Console.WriteLine("base number for TWT: {0}", int.Parse(args[3]));
			Console.WriteLine("SQL Generation mode: {0}", mode);

			Reader rdr = new Reader(args[0], "toSOS_S_TEKSTOW");
			Dicts dicts = new Dicts();

			int cnt = rdr.ReadSTW(dicts, int.Parse(args[2]), int.Parse(args[3]));

			Console.WriteLine("Counter={0}", cnt);

			//Console.WriteLine("{0}", dicts.ToSqlString());


			System.IO.StreamWriter writer = new System.IO.StreamWriter(args[1]);
			if(mode == "insert")
				writer.Write(dicts.to_insert_string());
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
