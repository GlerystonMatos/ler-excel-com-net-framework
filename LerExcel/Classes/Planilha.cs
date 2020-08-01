using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace LerExcel.Classes
{
    public class Planilha
    {
        public string[,] planilha;

        public Planilha(int Linha, int Coluna)
        {
            this.planilha = new string[Linha, Coluna];
        }
    }
}