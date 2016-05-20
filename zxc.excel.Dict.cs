using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace zxc.excel
{
	/// <summary>
	/// row from ph.s2_s_par_tekstu table
	/// </summary>
	class rec_PRT
	{
		static string prefix = "ph";
		static string table  = "s2_s_par_tekstu";

		public int    ID_PARAMETRU;
		public int    ID_TYP_PISMA;
		public int    ID_TEKST_PISMA;
		public int    ID_SEKCJI;

		public string RODZAJ;
		public string WIELE_PARAM;
		public string WARTOSC;


		public rec_PRT()
		{
			ID_TYP_PISMA   = 0;	//pk
			ID_TEKST_PISMA = 0;	//pk
			ID_PARAMETRU   = 0;	//pk
			ID_SEKCJI      = 0;
			RODZAJ         = "T"; //todo check
			WIELE_PARAM    = "T"; //T - jak sa przecinki w wartosci
			//PARAM_FILTR = "NULL";
			WARTOSC        = "";
		}

		public override string ToString()
		{
			return String.Format("{ID_PARAMETRU[{0}],{1},{2},{3},{4},{5},{6}}", 
				ID_PARAMETRU, ID_TYP_PISMA, ID_TEKST_PISMA, ID_SEKCJI, RODZAJ, WIELE_PARAM, WARTOSC);
		}

		public string to_insert_string()
		{
			StringBuilder sb = new StringBuilder();
			sb.AppendFormat("\ninsert into {0}.{1} ", prefix, table);
			sb.Append("PRT_ID_TYP_PISMA,PRT_ID_TEKST_PISMA,\nPRT_ID_PARAMETRU,");
			sb.Append("PRT_RODZAJ,PRT_ID_SEKCJI,PRT_PARAM_FILTR,\nPRT_WARTOSC,PRT_WIELE_PARAM) \nVALUES \n");
			sb.AppendFormat("({0},{1},{2},'{3}',{4},NULL,'{5}','{6}');", ID_TYP_PISMA, ID_TEKST_PISMA, ID_PARAMETRU,
						RODZAJ, ID_SEKCJI, WARTOSC, WIELE_PARAM );

			return sb.ToString();
		}

		public string to_update_string()
		{
			StringBuilder sb = new StringBuilder();
			sb.AppendFormat("\nupdate {0}.{1} set ", prefix, table);
			sb.AppendFormat("  PRT_RODZAJ={0},PRT_ID_SEKCJI={1},\n", RODZAJ, ID_SEKCJI);
			sb.AppendFormat("  PRT_WARTOSC={0},PRT_WIELE_PARAM={1} \n", WARTOSC, WIELE_PARAM);
			sb.AppendFormat("where PRT_ID_TYP_PISMA = {0} and \n", ID_TYP_PISMA);
			sb.AppendFormat("  PRT_ID_TEKST_PISMA = {0} and \n", ID_TEKST_PISMA);
			sb.AppendFormat("  PRT_ID_PARAMETRU = {0}", ID_PARAMETRU);

			return sb.ToString();
		}

	}

	/// <summary>
	///  row from SL.SOS_S_TEKST_PISMA 
	/// </summary>
	class rec_TWT
	{
		static string prefix = "SL";
		static string table  = "SOS_S_TEKST_PISMA";

		public int ID_TEKST_PISMA; //PK

		public string kod_pisma;
		public int    ID_TYP_PISMA;

 		public int ID_TEKST;

		public int ID_SEKCJI;
		public int NR_KOLEJNY;

		public string SPOS_FORMAT;

		public Dictionary<int, rec_PRT> PRT;
		public rec_TWT() { PRT = new Dictionary<int, rec_PRT>(); }


		public override string ToString()
		{
			return String.Format("{ID[{0}]:ID_TEKST{1}:Typ:{2}({3}):Sek{4}}", ID_TEKST_PISMA, ID_TEKST, ID_TYP_PISMA, kod_pisma, ID_SEKCJI);
		}

		public string to_insert_string()
		{
			StringBuilder sb = new StringBuilder();
			sb.AppendFormat("\nINSERT INTO {0}.{1}", prefix, table);
			sb.AppendFormat("\nTWT_ID_TEKST_PISMA,TWT_ID_TYP_PISMA,TWT_ID_TEKST,TWT_ID_SEKCJI,\nTWT_NR_KOL_TEKSTU,");
			sb.Append("TWT_CZY_DOMYSLNY,TWT_SPOS_FORMAT,\nTWT_DATA_OD,TWT_DATA_DO)\n");
			sb.AppendFormat("VALUES ({0},{1},{2},{3},{4},'T','{5}','2016-01-01','9999-09-09'); ", //-- {5}\n", 
				ID_TEKST_PISMA, ID_TYP_PISMA, ID_TEKST, ID_SEKCJI, NR_KOLEJNY, SPOS_FORMAT); //, kod_pisma);

			foreach(KeyValuePair<int, rec_PRT> prt in PRT)
			{
				sb.Append( prt.Value.to_insert_string() );
			}
			
			return sb.ToString();
		}

		public string to_update_string()
		{
			StringBuilder sb = new StringBuilder();
			sb.AppendFormat("\nUPDATE {0}.{1} SET \n", prefix, table);
			sb.AppendFormat("TWT_ID_SEKCJI={0},TWT_NR_KOL_TEKSTU={1},", ID_SEKCJI, NR_KOLEJNY);
			sb.AppendFormat("TWT_SPOS_FORMAT={0} ", SPOS_FORMAT);
			sb.AppendFormat("WHERE TWT_ID_TEKST_PISMA={0} AND ", ID_TEKST_PISMA );
			sb.AppendFormat(" AND TWT_ID_TYP_PISMA={0} AND TWT_ID_TEKST={1};", ID_TYP_PISMA, ID_TEKST);

			foreach(KeyValuePair<int, rec_PRT> prt in PRT)
			{
				sb.Append( prt.Value.to_update_string() );
			}
			
			return sb.ToString();
		}

	}

	/// <summary>
	/// row from SL.SOS_S_TEKSTOW
	/// </summary>
	class rec_STW
	{
		static string prefix = "SL";
		static string table  = "SOS_S_TEKSTOW";

		public int    STW_ID_TEKST; //PK
		public string STW_NAZWA;
		public string STW_TEKST;

		public Dictionary<int, rec_TWT> TWT;

		public rec_STW()
		{
			TWT = new Dictionary<int, rec_TWT>();
		}

		protected string NewLineConvert(string s)
		{
			StringBuilder sb = new StringBuilder();
			for(int i=0; i<s.Length; i++)
			{
				if(s[i] == '\n')
					sb.Append("/n");
				else
					sb.Append(s[i]);
			}

			return sb.ToString();
		}
	
		protected string LimitWidthConvert(string s, int width, string concat)
		{ 
			StringBuilder sb2 = new StringBuilder();
			int len = s.Length;
			for(int i=0; i<len; i+=width)
			{
				if(i+width < len)
				{
					string substring = s.Substring(i, width);
					sb2.AppendFormat("'{0}'{1}\n", substring, concat);
					//sb2.AppendFormat("'{0}'concat\n", substring);
				}
				else
				{
					string substring = s.Substring(i, len-i);
					sb2.AppendFormat("'{0}'", substring);
				}
			}

			return sb2.ToString();
		}

		public override string ToString()
		{
			return "{" + STW_ID_TEKST + ":" + STW_NAZWA + ":" + STW_TEKST + "}";
		}

		public string to_insert_string()
		{
			StringBuilder sb = new StringBuilder();
			sb.AppendFormat("\n\nINSERT INTO {0}.{1} \n", prefix, table);
			sb.Append("(STW_ID_TEKST,STW_NAZWA,STW_TEKST,STW_ID_JEDN_ZUS,STW_ID_ZNACZNIKA)\n");
			sb.AppendFormat("VALUES ({0},\n{1}\n,\n{2}\n,NULL,NULL);", 
				STW_ID_TEKST, 
				LimitWidthConvert( STW_NAZWA, 60, "concat"), 
				LimitWidthConvert( NewLineConvert(STW_TEKST), 60, "concat")
			);

			foreach(KeyValuePair<int, rec_TWT> twt in TWT)
				sb.Append( twt.Value.to_insert_string() );
			
			return sb.ToString();
		}

		public string to_update_string()
		{
			StringBuilder sb = new StringBuilder();
			sb.Append("\nUPDATE SL.SOS_S_TEKSTOW SET\n");
			sb.AppendFormat("STW_NAZWA={0},\nSTW_TEKST={1},\n", 
				LimitWidthConvert( STW_NAZWA, 60, "||"), 
				LimitWidthConvert( NewLineConvert(STW_TEKST), 60, "||")
			);

			sb.AppendFormat("STW_ID_JEDN_ZUS,STW_ID_ZNACZNIKA)\n");
			sb.AppendFormat("WHERE STW_ID_TEKST={0};", STW_ID_TEKST); 

			foreach(KeyValuePair<int, rec_TWT> twt in TWT)
				sb.Append( twt.Value.to_update_string() );
			
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

		public Dicts()
		{
			STW = new Dictionary<string, rec_STW>();
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

		public string to_insert_string()
		{
			StringBuilder sb = new StringBuilder();
			foreach(KeyValuePair<string, rec_STW> rec in STW)
			{
				sb.Append( rec.Value.to_insert_string() );
			}
			return sb.ToString();
		}

 		public string to_update_string()
		{
			StringBuilder sb = new StringBuilder();
			foreach(KeyValuePair<string, rec_STW> rec in STW)
			{
				sb.Append( rec.Value.to_update_string() );
			}
			return sb.ToString();
		}

	}
}
