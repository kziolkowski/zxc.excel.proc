using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace zxc.excel
{
	class rec_PRT
	{
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
			return String.Format("{ID_PARAMETRU[{0}],{1},{2},{3},{4},}", ID_PARAMETRU, ID_TYP_PISMA, ID_TEKST_PISMA, ID_SEKCJI, RODZAJ, WIELE_PARAM, WARTOSC);
		}

		public string ToSqlString()
		{
			StringBuilder sb = new StringBuilder();
			sb.Append("\ninsert into ph.s2_s_par_tekstu (PRT_ID_TYP_PISMA,PRT_ID_TEKST_PISMA,\nPRT_ID_PARAMETRU,");
			sb.Append("PRT_RODZAJ,PRT_ID_SEKCJI,PRT_PARAM_FILTR,\nPRT_WARTOSC,PRT_WIELE_PARAM) \nVALUES \n");
			sb.AppendFormat("({0},{1},{2},'{3}',{4},NULL,'{5}','{6}');", ID_TYP_PISMA, ID_TEKST_PISMA, ID_PARAMETRU,
						RODZAJ, ID_SEKCJI, WARTOSC, WIELE_PARAM );

			return sb.ToString();
		}
	}

	class rec_TWT
	{
		public int ID_TEKST_PISMA; //PK

		public string kod_pisma;
		public int    ID_TYP_PISMA;

 		public int ID_TEKST;

		public int ID_SEKCJI;
		public int NR_KOLEJNY;

		public Dictionary<int, rec_PRT> PRT;
		public rec_TWT() { PRT = new Dictionary<int, rec_PRT>(); }


		public override string ToString()
		{
			return String.Format("{ID[{0}]:ID_TEKST{1}:Typ:{2}({3}):Sek{4}}", ID_TEKST_PISMA, ID_TEKST, ID_TYP_PISMA, kod_pisma, ID_SEKCJI);
		}

		public string ToSqlString()
		{
			StringBuilder sb = new StringBuilder();
			sb.Append("\nINSERT INTO SL.SOS_S_TEKST_PISMA\n(TWT_ID_TEKST_PISMA,TWT_ID_TYP_PISMA,TWT_ID_TEKST,TWT_ID_SEKCJI,\nTWT_NR_KOL_TEKSTU,");
			sb.Append("TWT_CZY_DOMYSLNY,TWT_SPOS_FORMAT,\nTWT_DATA_OD,TWT_DATA_DO)\n");
			sb.AppendFormat("VALUES ({0},{1},{2},{3},{4},'T','K','2016-01-01','9999-09-09'); ", //-- {5}\n", 
				ID_TEKST_PISMA, ID_TYP_PISMA, ID_TEKST, ID_SEKCJI, NR_KOLEJNY); //, kod_pisma);

			foreach(KeyValuePair<int, rec_PRT> prt in PRT)
			{
				//sb.AppendFormat("    -- par:{0}", prt.Key);
				sb.Append( prt.Value.ToSqlString() );
			}
			
			return sb.ToString();
		}

	}


	class rec_STW
	{
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
			string bez_newLine =  sb.ToString();

			StringBuilder sb2 = new StringBuilder();
			int len = bez_newLine.Length;
			for(int i=0; i<len; i+=60)
			{
				if(i+60 < len)
				{
					string substring = bez_newLine.Substring(i, 60);
					sb2.AppendFormat("'{0}'concat\n", substring);
				}
				else
				{
					string substring = bez_newLine.Substring(i, len-i);
					sb2.AppendFormat("'{0}'", substring);
				}
			}

			return sb2.ToString();
		}

		public override string ToString()
		{
			return "{" + STW_ID_TEKST + ":" + STW_NAZWA + ":" + STW_TEKST + "}";
		}

		public string ToSqlString()
		{
			StringBuilder sb = new StringBuilder();
			sb.Append("\n\nINSERT INTO SL.SOS_S_TEKSTOW \n(STW_ID_TEKST,STW_NAZWA,STW_TEKST,STW_ID_JEDN_ZUS,STW_ID_ZNACZNIKA)\n");
			sb.AppendFormat("VALUES ({0},'{1}',\n{2}\n,NULL,NULL);", STW_ID_TEKST, STW_NAZWA, NewLineConvert(STW_TEKST));

			foreach(KeyValuePair<int, rec_TWT> twt in TWT)
			{
				//sb.AppendFormat("\n  -- twt:{0}", twt.Key);
				sb.Append( twt.Value.ToSqlString() );
			}
			
			return sb.ToString();
		}
	}


	class Dicts
	{
		public Dictionary<string, rec_STW> STW;
		//public Dictionary<int, rec_TWT>    TWT;
		//public Dictionary<int, rec_PRT>    PRT;

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

		public string ToSqlString()
		{
			StringBuilder sb = new StringBuilder();
			foreach(KeyValuePair<string, rec_STW> rec in STW)
			{
				//sb.AppendFormat("\n-- STW:{0}", rec.Key);
				sb.Append( rec.Value.ToSqlString() );
			}
			return sb.ToString();
		}
	}
}
