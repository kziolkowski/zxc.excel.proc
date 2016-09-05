using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Configuration;


namespace zxc.excel.proc
{

	class Program
	{
		static string mode        = "";

		static string sSTW_ID;
		public static int    nSTW_ID = 0;

		static string sTWT_ID;
		static int    nTWT_ID = 0;

		//static string sShowPrefix = "";
		static bool   bShowPrefix = true;

		static string inputFile;
		static string outputFile;

		static string outputDir   = "";

		static Dictionary<string, int> STW_dict;

		static Program()
		{
			STW_dict = new Dictionary<string, int>();
		}

		static void ReadConfig()
		{
			Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
			mode      = config.AppSettings.Settings["mode"].Value; 
			sSTW_ID   = config.AppSettings.Settings["STW_ID"].Value;
			nSTW_ID = int.Parse(sSTW_ID); 
			sTWT_ID   = config.AppSettings.Settings["TWT_ID"].Value; 
			nTWT_ID = int.Parse(sTWT_ID);
			outputDir  = config.AppSettings.Settings["outputDir" ].Value; 
			outputFile = config.AppSettings.Settings["outputFile"].Value; 
			inputFile  = config.AppSettings.Settings["inputFile" ].Value; 
		}

		static bool CheckRequiredParams(bool verbose)
		{
			bool res = false;

			if(verbose) {
			Console.WriteLine("input  file: {0}", inputFile);
			Console.WriteLine("output file: {0}", outputFile);
			Console.WriteLine("base id number for STW: {0}", nSTW_ID);
			Console.WriteLine("base id number for TWT: {0}", nTWT_ID);
			Console.WriteLine("SQL Generation mode: {0}", mode);
			//Console.WriteLine("present prefix in SQL: {0}", bShowPrefix.ToString());
			}

			if(!System.IO.File.Exists(inputFile)) {
				if(verbose) System.Console.WriteLine("Input file not exists: {0}", inputFile);
			}
			//if(!System.IO.Directory.Exists(outputDir)) {
			//	if(verbose) System.Console.WriteLine("Out dir not exists: {0}", outputDir);
			//	return false;
			//}
			else if(nSTW_ID <= 0) {
				if(verbose) System.Console.WriteLine("illegal STW_ID: {0}", nSTW_ID);
			}
			else if(nTWT_ID <= 0) {
				if(verbose) System.Console.WriteLine("illegal TWT_ID: {0}", nTWT_ID);
			}
			else if(mode !="insert" && mode!="update" && mode !="merge") {
				if(verbose) System.Console.WriteLine("illegal mode: {0}", mode);
			}
			else res = true;

			if(verbose && !res)
			{
				Console.WriteLine("Coś nie poszło ... naciś entera ...");
				Console.ReadLine();
			}

			return res;
		}

		static void CmdProcess(string[] args)
		{
			if(args.Count() > 0) inputFile  = args[0];
			if(args.Count() > 1) outputFile = args[1];
			if(args.Count() > 2) nSTW_ID = int.Parse(args[2]);
			if(args.Count() > 3) nTWT_ID = int.Parse(args[3]);
			if(args.Count() > 4) mode = args[4].ToLower();
			if(args.Count() > 5) bShowPrefix = bool.Parse(args[5]);

			if(!CheckRequiredParams(false))
				Console.WriteLine("\nUsage:\n\tzxc.excel.proc.EXE inputFilePath outputDirPath intBaseSTW intBaseTWT [insert|update|merge] bShowPrefix [true|false]\n");
		}

		static void Main(string[] args)
		{
			ReadConfig();
			CmdProcess(args);
			if(!CheckRequiredParams(true)) return;

			Reader rdr   = new Reader(inputFile, "toSOS_S_TEKSTOW");
			Dicts  dicts = new Dicts();

			int cnt = rdr.ReadSTW(dicts, nSTW_ID, nTWT_ID);

			int nSTW   = dicts.STW.Count();
			int nN_STW = dicts.CountNewSTW() ;
			int nTWT   = dicts.CountTWT();
			int nN_TWT = dicts.CountNewTWT();
			int nPRT   = dicts.CountPRT();
			Console.WriteLine("Counters: STW={0}/{1}, TWT={2}/{3}, PRT={4}, liczba insertów = {5} ", 
				nSTW, nN_STW, nTWT, nN_TWT, nPRT, nN_STW+nN_TWT+nPRT );

			//Console.WriteLine("{0}", dicts.ToSqlS`tring());

			string outFileName =  outputFile;
			string outFileNE = System.IO.Path.GetFileNameWithoutExtension(outFileName);
			string outExt  = System.IO.Path.GetExtension(outFileName);
			if(outputDir == "")
				outputDir  = System.IO.Path.GetDirectoryName(outFileName);
			string outFile    = System.IO.Path.Combine(outputDir, outFileNE+outExt);
			string outPreFile = System.IO.Path.Combine(outputDir, outFileNE+"_pre"+outExt);
			string delFile    = System.IO.Path.Combine(outputDir, outFileNE+"_del"+outExt);
			string delPreFile = System.IO.Path.Combine(outputDir, outFileNE+"_del_pre"+outExt);

			// kodowanie polskich znakow w skrypcie : Ansi Windows-1250 
			System.IO.StreamWriter writer_noPre   = new System.IO.StreamWriter( outFile,    false, Encoding.GetEncoding(1250) );
			System.IO.StreamWriter writer_pre     = new System.IO.StreamWriter( outPreFile, false, Encoding.GetEncoding(1250) );
			System.IO.StreamWriter writer_del     = new System.IO.StreamWriter( delFile,    false, Encoding.GetEncoding(1250) );
			System.IO.StreamWriter writer_delPre  = new System.IO.StreamWriter( delPreFile, false, Encoding.GetEncoding(1250) );

			if(mode == "insert")
			{
				writer_noPre.Write( dicts.to_insert_string(false) );
				writer_noPre.Close();

				writer_del.Write( dicts.to_delete_string(false) );
				writer_del.Close();

				writer_pre.Write( dicts.to_insert_string(true) );
				writer_pre.Close();

 				writer_delPre.Write( dicts.to_delete_string(true) );
				writer_delPre.Close();
			}
			else if(mode == "update")
			{
				writer_noPre.Write( dicts.to_update_string(false) ); writer_noPre.Close();
				writer_pre.Write(   dicts.to_update_string(true)  ); writer_pre.Close();
			}
			else if(mode == "merge")
			{
				writer_noPre.Write( dicts.to_merge_string(false) ); writer_noPre.Close();
				writer_pre.Write(   dicts.to_merge_string(true)  ); writer_pre.Close();
			}
			else
				Console.WriteLine("Nieznany sql mode.");

			Console.Write("Koniec, naciś entera ...");
			Console.ReadLine();
		}
	}
}
