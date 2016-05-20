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
			if(args.Count() != 4)
			{
				Console.WriteLine("\nUsage:\n\tzxc.excel.proc.EXE inputFilePath outputFilePath intBaseSTW intBaseTWT\n");
				return;
			}

			Console.WriteLine("input  file: {0}", args[0]);
			Console.WriteLine("output file: {0}", args[1]);
			Console.WriteLine("base number for STW: {0}", int.Parse(args[2]));
			Console.WriteLine("base number for TWT: {0}", int.Parse(args[3]));

			Reader rdr = new Reader(args[0], "toSOS_S_TEKSTOW");

			Dicts dicts = new Dicts();

			int cnt = rdr.ReadSTW(dicts, int.Parse(args[2]), int.Parse(args[3]));

			Console.WriteLine("Counter={0}", cnt);

			//Console.WriteLine("{0}", dicts.ToSqlString());


			System.IO.StreamWriter writer = new System.IO.StreamWriter(args[1]);
			writer.Write(dicts.to_insert_string());
			writer.Close();

			Console.Write("Koniec, naciś entera ...");
			Console.ReadLine();
		}
	}
}
