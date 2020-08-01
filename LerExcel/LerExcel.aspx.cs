using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using LerExcel.Classes;
using System.Text;
using System.IO;

namespace LerExcel
{
    public partial class LerExcel : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void btnLer_Click(object sender, EventArgs e)
        {
            string arquivo = Server.MapPath("Temp") + "\\" + "planilha.xlsx";
            fuArquivo.SaveAs(arquivo);
            
            Excel excel = new Excel();

            Planilha planilha = excel.Ler(arquivo);

            if (planilha == null)
            {
                this.lTabela.Text = "Não foi possível ler a planilha selecionada";
            }
            else
            {
                StringBuilder tabela = new StringBuilder();

                tabela.Append("<table class='planilha' border='1px'>");

                for (int i = 0; i != excel.GetLinhas(); i++)
                {
                    tabela.Append("<tr>");

                    for (int j = 0; j != excel.GetColunas(); j++)
                    {
                        tabela.Append(String.Format("<td>{0}</td>", planilha.planilha[i, j]));
                    }

                    tabela.Append("</tr>");
                }

                tabela.Append("</table>");

                this.lTabela.Text = tabela.ToString();
            }
        }
    }
}