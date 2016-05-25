using System;
using System.Reflection;

using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;


using ex = Microsoft.Office.Interop.Excel;

namespace zxc.excel
{

	class Reader
	{
		protected string          path;
		protected ex.Application  exApp;
		protected ex.Workbook     exWbk; 
		protected ex.Worksheet    exWks;

		protected Dictionary<string, int> dParamID;

		public Reader(string aPath, string sheet)
		{
			if(File.Exists(aPath))
				path = aPath;
			else
				throw new ArgumentException("File do not exists", aPath);
			
			exApp = new ex.Application();
			exWbk = exApp.Workbooks.Open(path);
			exWks = exWbk.Sheets[sheet];

			dParamID = new Dictionary<string, int>();
            dParamID.Add("M", 1);   // ("L", 1);
            dParamID.Add("N", 8);   // ("M", 8);
            dParamID.Add("O", 44);  // ("N", 44);
            dParamID.Add("P", 45);  // ("O", 45);
            dParamID.Add("Q", 72);  // ("P", 72);
            dParamID.Add("R", 75);  // ("Q", 75);
            dParamID.Add("S", 242); // ("R", 242);
            dParamID.Add("T", 245); // ("S", 245);
            dParamID.Add("U", 289); // ("T", 289);
            dParamID.Add("V", 304); // ("U", 324); 304 -> 324
            dParamID.Add("W", 305); // ("V", 325); 305 -> 325
            dParamID.Add("X", 306); // ("W", 326); 306 -> 326
            dParamID.Add("Y", 307); // ("X", 327); 307 -> 327
            dParamID.Add("Z", 308);  // ("Y", 328); 308 -> 328
            dParamID.Add("AA", 309); // ("Z", 329); 309 -> 329
            dParamID.Add("AB", 320); // ("AA", 320); 300 -> 320 zrobione wczesniej (nie wymaga zmiany)
            dParamID.Add("AC", 321); // ("AB", 321); 301 -> 321 j.w.
            dParamID.Add("AD", 322); // ("AC", 322); 302 -> 322 j.w.
            dParamID.Add("AE", 323); // ("AD", 323); 303 -> 323 j.w.
        }

		~Reader()
		{
			exWbk.Close(false);
		}

		public string Header(string col)
		{
			string r = col + "1";
			ex.Range hdr = exWks.Range[ r ];

			return hdr.Text;
		}

		public string Cell(string col, string row)
		{
            //return exWks.Range[ col + row ].Text;
            return exWks.Range[col + row].Text;
		}

		protected int ParseInt(string val)
		{
			if(val == "")
				return 0;
			else
				return int.Parse(val);
		}
		
		public rec_PRT MakePRT(int ID, string str_cell, rec_TWT twt)
		{
			rec_PRT prt = new rec_PRT();
			prt.ID_PARAMETRU   = ID;
			prt.ID_TYP_PISMA   = twt.ID_TYP_PISMA;
			prt.ID_SEKCJI      = twt.ID_SEKCJI;
			prt.ID_TEKST_PISMA = twt.ID_TEKST_PISMA;

			prt.WARTOSC        = str_cell;
			prt.WIELE_PARAM    = str_cell[0] == ','? "T" : "N";

			Console.Write("+{0}", prt.ID_PARAMETRU);
			return prt;
		}

		public rec_TWT ReadTWT(ref int ID, int ID_TEKST, string str_row)
		{
			rec_TWT twt = new rec_TWT();
			twt.ID_TEKST       = ID_TEKST;
			twt.ID_SEKCJI      = int.Parse(Cell("G", str_row));
			twt.kod_pisma      = Cell("H", str_row);
			twt.ID_TYP_PISMA   = int.Parse(Cell("I", str_row));
			twt.NR_KOLEJNY     = ParseInt(Cell("J", str_row));
   twt.SPOS_FORMAT    = Cell("K", str_row); // dla dodanej kolumny "K"
			twt.ID_TEKST_PISMA = ID++;

			Console.Write(" ={0}", twt.kod_pisma);
			foreach(KeyValuePair<string, int> par in dParamID)
			{
				string str_cell = Cell(par.Key, str_row);
				if(str_cell != "")
				{
					rec_PRT prt = MakePRT(par.Value, str_cell, twt);
					twt.PRT.Add(par.Value, prt);
				}
			}

			return twt;
		}

		public int ReadSTW(Dicts dicts, int baseSTW, int baseTWT)
		{
			int counter = 0;
			int twt_id = baseTWT;
			for(int row=2; row<172; row++)
			{
				string str_row = row.ToString();

				rec_STW stw;
				//stw.STW_ID_TEKST  = baseSTW + row - 2;
				//stw.STW_NAZWA     = Cell("D", str_row);
				//stw.STW_TEKST     = Cell("F", str_row);

				string key  = Cell("D", str_row);
				string name = Cell("E", str_row);
					int len = name.Length;
		                len = len > 149 ? 149 : len;
					name = name.Substring(0, len);
				string text = Cell("F", str_row);
				Console.Write("\n#{0}:{1}:{2}", key, row, twt_id);

				//rec_TWT twt = ReadTWT(ref twt_id, stw.STW_ID_TEKST, str_row);
				//rec_TWT twt = new rec_TWT();
				//twt.ID_TEKST   = stw.STW_ID_TEKST;
				//twt.ID_SEKCJI  = int.Parse(Cell("G", str_i));
				//twt.kod_pisma  = Cell("H", str_i);
				//twt.NR_KOLEJNY = ParseInt(Cell("J", str_i));
				//twt.ID_TEKST_PISMA = twt_id++;

				if(dicts.STW.ContainsKey(key))
				{
					stw = dicts.STW[key];
					rec_TWT twt = ReadTWT(ref twt_id, stw.STW_ID_TEKST, str_row);
					stw.TWT.Add(twt.ID_TEKST_PISMA, twt);
				}
				else
				{
					stw = new rec_STW();
					stw.STW_ID_TEKST = baseSTW + counter++;
					stw.STW_NAZWA    = Cell("D", str_row);// org name;
					stw.STW_TEKST    = text;
					rec_TWT twt = ReadTWT(ref twt_id, stw.STW_ID_TEKST, str_row);
					stw.TWT.Add(twt.ID_TEKST_PISMA, twt);
					dicts.STW.Add(key, stw);
				}

			}

			Console.WriteLine("TWT_ID={0}", twt_id);

			return counter;
		}
	}

}
