using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.OleDb;

namespace LerExcel.Classes
{
    public class Excel
    {
        private int Linha { get; set; }
        private int Coluna { get; set; }

        public int GetLinhas()
        {
            return this.Linha;
        }

        public int GetColunas()
        {
            return this.Coluna;
        }

        private OleDbConnection Conexao(string Arq)
        {
            return new OleDbConnection(String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0 Xml;HDR=YES';", Arq));
        }

        public Planilha Ler(string Arq)
        {
            OleDbConnection conectar = Conexao(Arq);

            string textoComando = "SELECT * FROM [Plan1$]";
            OleDbCommand comando = new OleDbCommand(textoComando, conectar);

            try
            {
                conectar.Open();
                
                OleDbDataReader leitor = comando.ExecuteReader();

                this.Coluna = leitor.FieldCount;
                this.Linha = leitor.VisibleFieldCount;

                Planilha excel = new Planilha(Linha, Coluna);

                int j = 0;
                while (leitor.Read()) 
                {
                    for (int i = 0; i != Coluna; i++)
                    {
                        excel.planilha[j, i] = leitor[i].ToString();
                    }

                    j++;
                }

                conectar.Close();
                return excel;
            }
            catch 
            {
                conectar.Close();
                return null;
            }
        }
    }
}