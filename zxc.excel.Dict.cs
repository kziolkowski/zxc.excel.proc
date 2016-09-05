using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace zxc.excel
{
	class sql_table_generator
	{
		public string prefix; // = "PREFIX";
		public string table; //  = "TABLE";

		public string table_name(bool usePrefix)
		{
			StringBuilder sb = new StringBuilder(64);
			if(usePrefix && prefix.Count()>0 )
				sb.AppendFormat("{0}.{1} ", prefix, table);
			else
				sb.AppendFormat("{0}", table);

			return sb.ToString();
		}
	}

	/// <summary>
	/// row from ph.s2_s_par_tekstu table
	/// </summary>
	class rec_PRT : sql_table_generator
	{
		//const string prefix = "ph";
		//const string table  = "s2_s_par_tekstu";

		public int    ID_PARAMETRU;
		public int    ID_TYP_PISMA;
		public int    ID_TEKST_PISMA;
		public int    ID_SEKCJI;

		public string RODZAJ;
		public string WIELE_PARAM;
		public string WARTOSC;

		public rec_PRT()
		{
			prefix = "ph";
			table  = "s2_s_par_tekstu";

			ID_TYP_PISMA   = 0;	//pk
			ID_TEKST_PISMA = 0;	//pk
			ID_PARAMETRU   = 0;	//pk
			ID_SEKCJI      = 0;
			RODZAJ         = "T"; // T - automaTycznie, W - recznie ???
			WIELE_PARAM    = "T"; // T - jak sa przecinki wokol ,wartosci,
			//PARAM_FILTR = "NULL";
			WARTOSC        = "";
		}

		public override string ToString()
		{
			return String.Format("{ID_PARAMETRU[{0}],{1},{2},{3},{4},{5},{6}}", 
				ID_PARAMETRU, ID_TYP_PISMA, ID_TEKST_PISMA, ID_SEKCJI, RODZAJ, WIELE_PARAM, WARTOSC);
		}

		public string to_delete_string(Boolean presentPrefix)
		{
			StringBuilder sb = new StringBuilder();
			sb.AppendFormat("\n    delete from {0} where", table_name(presentPrefix) );
			sb.AppendFormat("\n    PRT_ID_TYP_PISMA={0} and PRT_ID_TEKST_PISMA={1} and", ID_TYP_PISMA, ID_TEKST_PISMA);
			sb.AppendFormat("\n    PRT_ID_PARAMETRU={0};", ID_PARAMETRU);

			return sb.ToString();
		}

		public string to_insert_string(Boolean presentPrefix)
		{
			StringBuilder sb = new StringBuilder();
			sb.AppendFormat("\n    insert into {0} ", table_name(presentPrefix) );
			sb.Append("(PRT_ID_TYP_PISMA,PRT_ID_TEKST_PISMA,\n    PRT_ID_PARAMETRU,");
			sb.Append("PRT_RODZAJ,PRT_ID_SEKCJI,PRT_PARAM_FILTR,\n    PRT_WARTOSC,PRT_WIELE_PARAM) \n    values \n");
			sb.AppendFormat("    ({0},{1},{2},'{3}',{4},NULL,'{5}','{6}');", ID_TYP_PISMA, ID_TEKST_PISMA, ID_PARAMETRU,
						RODZAJ, ID_SEKCJI, WARTOSC, WIELE_PARAM );

			return sb.ToString();
		}

		public string to_merge_string(Boolean presentPrefix)
		{
			StringBuilder sb = new StringBuilder();
			sb.AppendFormat("\n    merge into {0} as p using ( values ", table_name(presentPrefix) );
 			sb.AppendFormat("\n    ({0},{1},{2},'{3}',{4},NULL,'{5}','{6}') );", 
						ID_TYP_PISMA, ID_TEKST_PISMA, ID_PARAMETRU,
						RODZAJ, ID_SEKCJI, WARTOSC, WIELE_PARAM );
			sb.Append("\n    as r (PRT_ID_TYP_PISMA,PRT_ID_TEKST_PISMA,PRT_ID_PARAMETRU,");
			sb.Append("\n    PRT_RODZAJ,PRT_ID_SEKCJI,PRT_PARAM_FILTR,PRT_WARTOSC,PRT_WIELE_PARAM)");
			sb.Append("\n    on  p.PRT_ID_TYP_PISMA = r.PRT_ID_TYP_PISMA");
			sb.Append("\n    and p.PRT_ID_TEKST_PISMA = r.PRT_ID_TEKST_PISMA");
			sb.Append("\n    and p.PRT_ID_PARAMETRU = r.PRT_ID_PARAMETRU");
			sb.Append("\n    when mached then update set ");
			sb.Append("\n    p.PRT_ID_TYP_PISMA=r.PRT_ID_TYP_PISMA,");
			sb.Append("\n    p.PRT_ID_TEKST_PISMA=r.PRT_ID_TEKST_PISMA,");
			sb.Append("\n    p.PRT_ID_PARAMETRU=r.PRT_ID_PARAMETRU,");
			sb.Append("\n    p.PRT_RODZAJ=r.PRT_RODZAJ, p.PRT_ID_SEKCJI=r.PRT_ID_SEKCJI,");
			sb.Append("\n    p.PRT_PARAM_FILTR=r.PRT_PARAM_FILTR,");
			sb.Append("\n    p.PRT_WARTOSC=r.PRT_WARTOSC, p.PRT_WIELE_PARAM=r.PRT_WIELE_PARAM);");
			sb.Append("\n    when not mached then insert  ");
			sb.Append("\n    (PRT_ID_TYP_PISMA,PRT_ID_TEKST_PISMA,PRT_ID_PARAMETRU,");
			sb.Append("\n    PRT_RODZAJ,PRT_ID_SEKCJI,PRT_PARAM_FILTR,");
			sb.Append("\n    PRT_WARTOSC,PRT_WIELE_PARAM)");
			sb.Append("\n    values (r.PRT_ID_TYP_PISMA,r.PRT_ID_TEKST_PISMA,r.PRT_ID_PARAMETRU,");
			sb.Append("\n    r.PRT_RODZAJ,r.PRT_ID_SEKCJI,r.PRT_PARAM_FILTR,");
			sb.Append("\n    r.PRT_WARTOSC,r.PRT_WIELE_PARAM);");

			return sb.ToString();
		}

		public string to_update_string(Boolean presentPrefix)
		{
			StringBuilder sb = new StringBuilder();
			sb.AppendFormat("\n  update {0} set\n", table_name(presentPrefix) );
			sb.AppendFormat("  PRT_RODZAJ='{0}',PRT_ID_SEKCJI={1},PRT_WIELE_PARAM='{2}',\n", RODZAJ, ID_SEKCJI, WIELE_PARAM);
			sb.AppendFormat("  PRT_WARTOSC='{0}' \n", WARTOSC);
			sb.AppendFormat("  where PRT_ID_TYP_PISMA  ={0} and \n", ID_TYP_PISMA);
			sb.AppendFormat("        PRT_ID_TEKST_PISMA={0} and \n", ID_TEKST_PISMA);
			sb.AppendFormat("        PRT_ID_PARAMETRU  ={0};", ID_PARAMETRU);

			return sb.ToString();
		}

	}

	/// <summary>
	///  row from SL.SOS_S_TEKST_PISMA 
	/// </summary>
	class rec_TWT	: sql_table_generator
	{
		//static string prefix = "SL";
		//static string table  = "SOS_S_TEKST_PISMA";

		public bool rec_EXISTS;

		public int ID_TEKST_PISMA; //PK

		public string kod_pisma;
		public int    ID_TYP_PISMA;

 		public int ID_TEKST;

		public int ID_SEKCJI;
		public int NR_KOLEJNY;

		public string SPOS_FORMAT;

		public Dictionary<int, rec_PRT> PRT;

		public rec_TWT()
		{
			prefix = "SL";
			table  = "SOS_S_TEKST_PISMA";

			PRT = new Dictionary<int, rec_PRT>(16);
		}

		public override string ToString()
		{
			return String.Format("{ID[{0}]:ID_TEKST{1}:Typ:{2}({3}):Sek{4}}", ID_TEKST_PISMA, ID_TEKST, ID_TYP_PISMA, kod_pisma, ID_SEKCJI);
		}

		public string to_delete_string(Boolean presentPrefix)
		{
			StringBuilder sb = new StringBuilder();
			
			foreach(KeyValuePair<int, rec_PRT> prt in PRT)
			{
				sb.Append( prt.Value.to_delete_string(presentPrefix) );
			}

			if(!rec_EXISTS)
			{
				sb.AppendFormat("\n  DELETE FROM {0}", table_name(presentPrefix) );
				sb.AppendFormat("\n  WHERE TWT_ID_TEKST_PISMA={0};", ID_TEKST_PISMA);
			}
			else
			{
				sb.AppendFormat("\n  -- DELETE FROM {0}", table_name(presentPrefix) );
				sb.AppendFormat("\n  -- WHERE TWT_ID_TEKST_PISMA={0};", ID_TEKST_PISMA);
			}

			return sb.ToString();
		}

		public string to_insert_string(Boolean presentPrefix)
		{
			StringBuilder sb = new StringBuilder();
			if(!rec_EXISTS)
			{
				sb.AppendFormat("\n  INSERT INTO {0}", table_name(presentPrefix) );
				sb.AppendFormat("\n  (TWT_ID_TEKST_PISMA,TWT_ID_TYP_PISMA,TWT_ID_TEKST,TWT_ID_SEKCJI," );
				sb.Append(      "\n  TWT_NR_KOL_TEKSTU,TWT_CZY_DOMYSLNY,TWT_SPOS_FORMAT,");
				sb.Append(      "\n  TWT_DATA_OD,TWT_DATA_DO)\n");
				sb.AppendFormat("  VALUES\n  ({0},{1},{2},{3},{4},'T','{5}','2016-01-01','9999-12-31'); ", //-- {5}\n", 
					ID_TEKST_PISMA, ID_TYP_PISMA, ID_TEKST, ID_SEKCJI, NR_KOLEJNY, SPOS_FORMAT); //, kod_pisma);
			}
			else
			{
				sb.AppendFormat("\n  -- {0} TWT_ID_TEKST_PISMA={1}", table_name(presentPrefix), ID_TEKST_PISMA);
			}

			foreach(KeyValuePair<int, rec_PRT> prt in PRT)
			{
				sb.Append( prt.Value.to_insert_string(presentPrefix) );
			}
			
			return sb.ToString();
		}

		public string to_merge_string(Boolean presentPrefix)
		{
			StringBuilder sb = new StringBuilder();
			if(!rec_EXISTS)
			{
				sb.AppendFormat("\n  MERGE INTO {0} AS T USING ( VALUES", table_name(presentPrefix) );
				sb.AppendFormat("\n  ({0},{1},{2},{3},{4},'T','{5}','2016-01-01','9999-12-31')) AS R", //-- {5}\n", 
					ID_TEKST_PISMA, ID_TYP_PISMA, ID_TEKST, ID_SEKCJI, NR_KOLEJNY, SPOS_FORMAT); //, kod_pisma);
				sb.Append("\n  (TWT_ID_TEKST_PISMA,TWT_ID_TYP_PISMA,TWT_ID_TEKST,TWT_ID_SEKCJI," );
				sb.Append("\n  TWT_NR_KOL_TEKSTU,TWT_CZY_DOMYSLNY,TWT_SPOS_FORMAT,");
				sb.Append("\n  TWT_DATA_OD,TWT_DATA_DO)");
				sb.Append("\n  ON T.TWT_ID_TEKST_PISMA = R.TWT_ID_TEKST_PISMA");
				sb.Append("\n  WHEN MATCHED THEN update set ");
				sb.Append("\n  T.TWT_ID_TEKST_PISMA=R.TWT_ID_TEKST_PISMA,");
				sb.Append("\n  T.TWT_ID_TYP_PISMA=R.TWT_ID_TYP_PISMA,");
				sb.Append("\n  T.TWT_ID_TEKST=R.TWT_ID_TEKST,");
				sb.Append("\n  T.TWT_ID_SEKCJI=R.TWT_ID_SEKCJI," );
				sb.Append("\n  T.TWT_NR_KOL_TEKSTU=R.TWT_NR_KOL_TEKSTU,");
				sb.Append("\n  T.TWT_CZY_DOMYSLNY=R.TWT_CZY_DOMYSLNY,");
				sb.Append("\n  T.TWT_SPOS_FORMAT=R.TWT_SPOS_FORMAT,");
				sb.Append("\n  T.TWT_DATA_OD=R.TWT_DATA_OD, T.TWT_DATA_DO=R.TWT_DATA_DO");
				sb.Append("\n  WHEN NOT MATCHED THEN insert ");
				sb.Append("\n  (TWT_ID_TEKST_PISMA,TWT_ID_TYP_PISMA,TWT_ID_TEKST,TWT_ID_SEKCJI," );
				sb.Append("\n  TWT_NR_KOL_TEKSTU,TWT_CZY_DOMYSLNY,TWT_SPOS_FORMAT,");
				sb.Append("\n  TWT_DATA_OD,TWT_DATA_DO) VALUES");
				sb.Append("\n  (R.TWT_ID_TEKST_PISMA,R.TWT_ID_TYP_PISMA,R.TWT_ID_TEKST,R.TWT_ID_SEKCJI," );
				sb.Append("\n  R.TWT_NR_KOL_TEKSTU,R.TWT_CZY_DOMYSLNY,R.TWT_SPOS_FORMAT,");
				sb.Append("\n  R.TWT_DATA_OD,R.TWT_DATA_DO);");
			}
			else
			{
				sb.AppendFormat("\n  -- {0} TWT_ID_TEKST_PISMA={1}", table_name(presentPrefix), ID_TEKST_PISMA);
			}

			foreach(KeyValuePair<int, rec_PRT> prt in PRT)
			{
				sb.Append( prt.Value.to_merge_string(presentPrefix) );
			}
			
			return sb.ToString();
		}


		public string to_update_string(Boolean presentPrefix)
		{
			StringBuilder sb = new StringBuilder();
			sb.AppendFormat("\n UPDATE {0} SET \n", table_name(presentPrefix) );
			sb.AppendFormat(" TWT_ID_SEKCJI={0}, TWT_NR_KOL_TEKSTU={1}, ", ID_SEKCJI, NR_KOLEJNY);
			sb.AppendFormat("TWT_SPOS_FORMAT='{0}' , ", SPOS_FORMAT);
			sb.AppendFormat("TWT_CZY_DOMYSLNY='{0}' ", 'N');
			sb.AppendFormat("\n WHERE TWT_ID_TEKST_PISMA={0} AND\n", ID_TEKST_PISMA );
			sb.AppendFormat("    TWT_ID_TYP_PISMA={0} AND TWT_ID_TEKST={1};", ID_TYP_PISMA, ID_TEKST);

			foreach(KeyValuePair<int, rec_PRT> prt in PRT)
			{
				sb.Append( prt.Value.to_update_string(presentPrefix) );
			}
			
			return sb.ToString();
		}

	}

	/// <summary>
	/// row from SL.SOS_S_TEKSTOW
	/// </summary>
	class rec_STW : sql_table_generator
	{
//		static string prefix = "SL";
//		static string table  = "SOS_S_TEKSTOW";
		static string concatString = "||";       //uwaga MJ zamiana "concat" na "||" (double pipe)

		static String newLine = "\nCHR(13)||CHR(10)||\n";

		public bool   rec_EXISTS;   // rekord istnieje nie robimy insert
		public int    STW_ID_TEKST; //PK
		public string STW_NAZWA;
		public string STW_TEKST;

		public Dictionary<int, rec_TWT> TWT;

		public rec_STW()
		{
			prefix = "SL";
			table  = "SOS_S_TEKSTOW";

			TWT = new Dictionary<int, rec_TWT>(16);
			rec_EXISTS = false;
		}

		public rec_STW(int ID, string NAZWA, string TEXT, bool bExists)
		{
			prefix = "SL";
			table  = "SOS_S_TEKSTOW";

			STW_ID_TEKST = ID;
			STW_NAZWA    = NAZWA;
			STW_TEKST    = TEXT;
			rec_EXISTS   = bExists;

			TWT = new Dictionary<int, rec_TWT>(16);
		}


		protected string NewLineConvert(string s)
		{
			StringBuilder sb = new StringBuilder();
			for(int i=0; i<s.Length; i++)
			{
				if(s[i] == '\n')
					sb.Append(newLine);
				else
					sb.Append(s[i]);
			}

			return sb.ToString();
		}
	
		protected string LimitWidthConvert(string s, int width, string concat)
		{ 
			const string newLineSeq = "CHR(13)||CHR(10)||";

			StringBuilder sb = new StringBuilder();
			int len     = s.Length;
			int lineLen = 0;
			for(int i=0; i<len; i += lineLen)
			{
				if(i+width < len)
				{
					string substring = s.Substring(i, width);
					int newLinePos = substring.IndexOf('\n');
					if(newLinePos == -1)
					{
						sb.AppendFormat("'{0}'{1}\n", substring, concat);
						lineLen = width;
					}
					else
					{
						string subsub = substring.Substring(0, newLinePos);
						if(subsub == newLineSeq)
							sb.AppendFormat("{0}\n", subsub);
						else
							sb.AppendFormat("'{0}'{1}\n", subsub, concat);
						lineLen = newLinePos+1;
					}
				}
				else
				{
					string substring = s.Substring(i, len-i);
					int newLinePos = substring.IndexOf('\n');
					if(newLinePos == -1)
					{
						sb.AppendFormat("'{0}'\n", substring);
						lineLen = width;
					}
					else
					{
						string subsub = substring.Substring(0, newLinePos);
						if(subsub == newLineSeq)
							sb.AppendFormat("{0}\n", subsub);
						else
							sb.AppendFormat("'{0}'{1}\n", subsub, concat);
						lineLen = newLinePos+1;
					}
				}
			}

			return sb.ToString();
		}

		public override string ToString()
		{
			return "{" + STW_ID_TEKST + ":" + STW_NAZWA + ":" + STW_TEKST + "}";
		}

		public string to_delete_string(Boolean presentPrefix)
		{
			StringBuilder sb = new StringBuilder();

			foreach(KeyValuePair<int, rec_TWT> twt in TWT)
				sb.Append( twt.Value.to_delete_string(presentPrefix) );

			int len = Math.Min(STW_NAZWA.Length, 148);
			int short_len = Math.Min(STW_NAZWA.Length, 6);
			if(!rec_EXISTS)
			{ 
				sb.AppendFormat("\nDELETE FROM {0}\n", table_name(presentPrefix) );
				sb.AppendFormat("WHERE STW_ID_TEKST={0};", STW_ID_TEKST);
				// print version record - to allow see it
				if(STW_ID_TEKST == zxc.excel.proc.Program.nSTW_ID-1)
				{
					sb.AppendFormat("\n-- STW_NAZWA:'{0}'", STW_NAZWA);
 					sb.AppendFormat("\n-- STW_TEKST:'{0}'", STW_TEKST);
				}
			}
			else
			{
				sb.AppendFormat("\n--DELETE FROM {0}\n", table_name(presentPrefix) );
				sb.AppendFormat("-- WHERE STW_ID_TEKST={0};", STW_ID_TEKST);
			}

			return sb.ToString();
		}

		public string to_insert_string(Boolean presentPrefix)
		{
			int len = Math.Min(STW_NAZWA.Length, 148);
			int short_len = Math.Min(STW_NAZWA.Length, 6);
			StringBuilder sb = new StringBuilder();
			if(!rec_EXISTS)
			{ 
				sb.AppendFormat("\nINSERT INTO {0} \n", table_name(presentPrefix) );
				sb.Append("(STW_ID_TEKST,STW_NAZWA,STW_TEKST,STW_ID_JEDN_ZUS,STW_ID_ZNACZNIKA)\n");
				sb.AppendFormat("VALUES ({0},\n{1},\n{2},\nNULL,NULL);", 
					STW_ID_TEKST, 
					LimitWidthConvert( STW_NAZWA.Substring(0, len), 65, concatString), 
					LimitWidthConvert( NewLineConvert(STW_TEKST), 65, concatString)
				);
			}
			else
			{
				sb.AppendFormat("\n-- {0} STW_ID_TEKST={1} {2}", table_name(presentPrefix), 
					STW_ID_TEKST, STW_NAZWA.Substring(0, short_len));
			}

			foreach(KeyValuePair<int, rec_TWT> twt in TWT)
				sb.Append( twt.Value.to_insert_string(presentPrefix) );
			
			return sb.ToString();
		}

		public string to_merge_string(Boolean presentPrefix)
		{
			int len = Math.Min(STW_NAZWA.Length, 148);
			int short_len = Math.Min(STW_NAZWA.Length, 6);
			StringBuilder sb = new StringBuilder();
			if(!rec_EXISTS)
			{ 
				sb.AppendFormat("\nMERGE INTO {0} AS S USING ( ", table_name(presentPrefix) );
				sb.AppendFormat("VALUES({0},\n{1},\n{2}", 
					STW_ID_TEKST, 
					LimitWidthConvert( STW_NAZWA.Substring(0, len), 65, concatString), 
					LimitWidthConvert( NewLineConvert(STW_TEKST),   65, concatString)
				);
				sb.Append(")) AS R (ID_TEKST, NAZWA, TEKST) on R.ID_TEKST = S.STW_ID_TEKST" );
				sb.Append("\nWHEN MATCHED THEN UPDATE ");
				sb.Append("\nSET S.STW_NAZWA = R.NAZWA, S.STW_TEKST = R.TEKST");
				sb.Append("\nWHEN NOT MATCHED THEN INSERT");
				sb.Append("\n(STW_ID_TEKST,STW_NAZWA,STW_TEKST) \nVALUES (R.ID_TEKST, R.NAZWA, R.TEKST);");
			}
			else
			{
				sb.AppendFormat("\n-- {0} STW_ID_TEKST={1} {2}", table_name(presentPrefix), 
					STW_ID_TEKST, STW_NAZWA.Substring(0, short_len));
			}

			foreach(KeyValuePair<int, rec_TWT> twt in TWT)
				sb.Append( twt.Value.to_merge_string(presentPrefix) );
			
			return sb.ToString();
		}

		public string to_update_string(Boolean presentPrefix)
		{
			StringBuilder sb = new StringBuilder();
		if(presentPrefix)
			sb.Append("\nUPDATE SL.SOS_S_TEKSTOW SET");
		else
			sb.Append("\nUPDATE SOS_S_TEKSTOW SET");

			sb.AppendFormat(" \nSTW_NAZWA=\n{0},\nSTW_TEKST=\n{1} ", 
				LimitWidthConvert( STW_NAZWA, 65, concatString), 
				LimitWidthConvert( NewLineConvert(STW_TEKST), 65, concatString)
			);

			//sb.AppendFormat("STW_ID_JEDN_ZUS,STW_ID_ZNACZNIKA)\n");
			sb.AppendFormat("\nWHERE STW_ID_TEKST={0};", STW_ID_TEKST); 

			foreach(KeyValuePair<int, rec_TWT> twt in TWT)
				sb.Append( twt.Value.to_update_string(presentPrefix) );
			
			return sb.ToString();
		}
	}

	/// <summary>
	/// root dictionary 
	/// - of records with other dictionaries
	/// </summary>
	class Dicts
	{
		public Dictionary<string, rec_STW> STW;
		//public Boolean bshowPrefix;

		public Dicts()
		{
			STW = new Dictionary<string, rec_STW>(128);
		}

		public int CountNewSTW()
		{
			int sum = 0;
			foreach(var x in STW)
				sum += x.Value.rec_EXISTS ? 0 : 1;

			return sum;
		}
		public int CountNewTWT()
		{
			int sum = 0;
			foreach(var x in STW)
				sum += x.Value.TWT.Count(rec => !rec.Value.rec_EXISTS);

			return sum;
		}

		public int CountTWT()
		{
			int sum = 0;
			foreach(var x in STW)
				sum += x.Value.TWT.Count();

			return sum;
		}


		public int CountPRT()
		{
			int sum = 0;
			foreach(var x in STW)
				foreach(var y in x.Value.TWT)
					sum += y.Value.PRT.Count();

			return sum;
		}

		public override string ToString()
		{
			StringBuilder sb = new StringBuilder();
			sb.Append("===== STW ======\n");
			foreach(KeyValuePair<string, rec_STW> rec in STW)
			{
				 sb.AppendFormat("k:{0} = v:{1}\n", rec.Key, rec.Value);
			}
			return sb.ToString();
		}

		public string to_delete_string(bool bPrefix)
		{
			StringBuilder sb = new StringBuilder();
			foreach(KeyValuePair<string, rec_STW> rec in STW)
			{
				sb.Append( rec.Value.to_delete_string(bPrefix) );
			}
			return sb.ToString();
		}

		public string to_insert_string(bool bPrefix)
		{
			StringBuilder sb = new StringBuilder();
			foreach(KeyValuePair<string, rec_STW> rec in STW)
			{
				sb.Append( rec.Value.to_insert_string(bPrefix) );
			}
			return sb.ToString();
		}

		public string to_merge_string(bool bPrefix)
		{
			StringBuilder sb = new StringBuilder();
			foreach(KeyValuePair<string, rec_STW> rec in STW)
			{
				sb.Append( rec.Value.to_merge_string(bPrefix) );
			}
			return sb.ToString();
		}

 		public string to_update_string(bool bPrefix)
		{
			StringBuilder sb = new StringBuilder();
			foreach(KeyValuePair<string, rec_STW> rec in STW)
			{
				sb.Append( rec.Value.to_update_string(bPrefix) );
			}
			return sb.ToString();
		}

	}
}
