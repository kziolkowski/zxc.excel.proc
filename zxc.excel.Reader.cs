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
			dParamID.Add("M",    1); // ("L", 1);
			dParamID.Add("N",    8); // ("M", 8);
			dParamID.Add("O",   44); // ("N", 44);
			dParamID.Add("P",   45); // ("O", 45);
			dParamID.Add("Q",   72); // ("P", 72);
			dParamID.Add("R",   75); // ("Q", 75);
			dParamID.Add("S",  242); // ("R", 242);
			dParamID.Add("T",  245); // ("S", 245);
			dParamID.Add("U",  289); // ("T", 289);
			dParamID.Add("V",  324); // ("U", 324); 304 -> 324
			dParamID.Add("W",  325); // ("V", 325); 305 -> 325
			dParamID.Add("X",  326); // ("W", 326); 306 -> 326
			dParamID.Add("Y",  327); // ("X", 327); 307 -> 327
			dParamID.Add("Z",  328); // ("Y", 328); 308 -> 328
			dParamID.Add("AA", 329); // ("Z", 329); 309 -> 329
			dParamID.Add("AB", 320); // ("AA", 320); 300 -> 320 zrobione wczesniej (nie wymaga zmiany)
			dParamID.Add("AC", 321); // ("AB", 321); 301 -> 321 j.w.
			dParamID.Add("AD", 322); // ("AC", 322); 302 -> 322 j.w.
			dParamID.Add("AE", 323); // ("AD", 323); 303 -> 323 j.w.
			dParamID.Add("AF",  46); // ("AF", 46); dodal Adam A.
			dParamID.Add("AG",  28); // ("AD", 323);dodal Michal
			dParamID.Add("AH",  66); // ("AF", 46); dodal Michal
		}

		~Reader()
		{
			exWbk.Close(false);
		}

		public int CreateParamDict()
		{

			return 0;
		}

		public int RowCount(int nMax)
		{
			for(int i=2; i<nMax; i++)   // 2 pierwsze wiersze pomijamy
			{
				string str_i = i.ToString();
				if(Cell("A", str_i) == "" && Cell("B", str_i) == "")
					return i;
			}

			return nMax;
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
		
		public rec_PRT MakePRT(int ID, string str_cell, string czyReczny, rec_TWT twt)
		{
			rec_PRT prt = new rec_PRT();
			prt.ID_PARAMETRU   = ID;
			prt.ID_TYP_PISMA   = twt.ID_TYP_PISMA;
			prt.ID_SEKCJI      = twt.ID_SEKCJI;
			prt.ID_TEKST_PISMA = twt.ID_TEKST_PISMA;

			prt.WARTOSC        = str_cell;
			prt.RODZAJ         = czyReczny.ToUpper() == "R" ? "W" : "T";
			prt.WIELE_PARAM    = str_cell[0] == ','? "T" : "N";

			Console.Write("+{0}", prt.ID_PARAMETRU);
			return prt;
		}

		public rec_TWT ReadTWT(ref int ID, int ID_TEKST, string str_row)
		{
			rec_TWT twt = new rec_TWT();

			int    nID_TEKST_PISMA = 0;
			string sID_TEKST_PISMA = Cell("E", str_row);
			if(sID_TEKST_PISMA != "")
				nID_TEKST_PISMA = int.Parse(sID_TEKST_PISMA);

			if(nID_TEKST_PISMA == 0)
			{
				twt.ID_TEKST_PISMA = ID++;
				twt.rec_EXISTS = false;
			}
			else
			{
				twt.ID_TEKST_PISMA = nID_TEKST_PISMA;
				twt.rec_EXISTS = true;
			}

			twt.ID_TEKST       = ID_TEKST;
			twt.ID_SEKCJI      = int.Parse(Cell("G", str_row));
			twt.kod_pisma      = Cell("H", str_row);
			twt.ID_TYP_PISMA   = int.Parse(Cell("I", str_row));
			twt.NR_KOLEJNY     = ParseInt(Cell("J", str_row));
			twt.SPOS_FORMAT    = Cell("K", str_row);            // dla dodanej kolumny "K"

			string czyReczny   = Cell("L", str_row);

			Console.Write(" ={0}", twt.kod_pisma);
			foreach(KeyValuePair<string, int> par in dParamID)
			{
				string str_cell = Cell(par.Key, str_row);
				if(str_cell != "")
				{
					rec_PRT prt = MakePRT(par.Value, str_cell, czyReczny, twt);
					twt.PRT.Add(par.Value, prt);
				}
			}

			return twt;
		}

		public int ReadSTW(Dicts dicts, int baseSTW, int baseTWT)
		{
			int counter = 0;
			int twt_id = baseTWT;
			int max_row = RowCount(500);

			
			string ver_name = string.Format("VER_{0}_{1} - wersja ", 
				System.DateTime.Now.ToShortDateString(), 
				System.DateTime.Now.ToShortTimeString());

			string ver_desc = string.Format("user: {0}, date nad time of generation: {1} {2}",
				Environment.UserName,
				DateTime.Now.ToShortDateString(),
				DateTime.Now.ToShortTimeString()); 

			rec_STW ver_STW = new rec_STW(baseSTW-1, ver_name, ver_desc, false);
			dicts.STW.Add(ver_name, ver_STW);
		 
			System.Console.WriteLine("Liczba wierszy: {0}", max_row);
			for(int row=2; row<max_row; row++)
			{
				string str_row = row.ToString();

				rec_STW stw;
				//stw.STW_ID_TEKST  = baseSTW + row - 2;
				//stw.STW_NAZWA     = Cell("D", str_row);
				//stw.STW_TEKST     = Cell("F", str_row);

				int    nID_TEKST = 0;
				string sID_TEKST = Cell("C", str_row);

				if(sID_TEKST != "") nID_TEKST = int.Parse(Cell("C", str_row));

				string key;
				
				if(nID_TEKST == 0)
				{ 
					key = Cell("D", str_row);
					//string name = Cell("E", str_row);
					int len = key.Length;
						len = len > 149 ? 149 : len;
					key = key.Substring(0, len);
				}
				else
				{
					key = sID_TEKST;
				}

				string text = Cell("F", str_row);
				int key_len = Math.Min(key.Length, 32);
				Console.Write("\n#{0}..:{1}:{2}", key.Substring(0, key_len), row, twt_id);

				if(dicts.STW.ContainsKey(key))
				{
					stw = dicts.STW[key];
					rec_TWT twt = ReadTWT(ref twt_id, stw.STW_ID_TEKST, str_row);
					stw.TWT.Add(twt.ID_TEKST_PISMA, twt);
				}
				else
				{
					if(nID_TEKST == 0)
					{
						stw = new rec_STW(baseSTW + counter++, key, text, false);
						//stw.STW_ID_TEKST = baseSTW + counter++;
						//stw.STW_NAZWA    = key;
						//stw.STW_TEKST    = text;
					}
					else
					{
 						stw = new rec_STW(nID_TEKST, key, text, true);
					}
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
