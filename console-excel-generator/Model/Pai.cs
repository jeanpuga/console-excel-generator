using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace console_excel_generator.Model
{
    public class Pai
    {
        public int profundidade;

        public Pai(int profundidade, int largura, Random r,int nivel=1 )
        {
           
            nome = GetLetter(r) + GetLetter(r) + GetLetter(r) + GetLetter(r) + GetLetter(r) + GetLetter(r) + GetLetter(r) + GetLetter(r) + GetLetter(r) + GetLetter(r) + GetLetter(r) + GetLetter(r) + GetLetter(r) + GetLetter(r) + GetLetter(r);
            desc= GetLetter(r) + GetLetter(r) + GetLetter(r) + GetLetter(r) + GetLetter(r) + GetLetter(r) + GetLetter(r) + GetLetter(r) + GetLetter(r) + GetLetter(r) + GetLetter(r) + GetLetter(r) + GetLetter(r) + GetLetter(r) + GetLetter(r);
            valor1 = NextDouble(r, 9999, 99999);
            valor2 = NextDouble(r, 9999, 99999);
            valor3 = NextDouble(r, 9999, 99999);
            valor4 = NextDouble(r, 9999, 99999);
            valor5 = NextDouble(r, 9999, 99999);
            valor6 = NextDouble(r, 9999, 99999);
            valor7 = NextDouble(r, 9999, 99999);
            valor8 = NextDouble(r, 9999, 99999);
            this.profundidade = profundidade;
            this.nivel = nivel;
            if (profundidade > 0)
            {
                filhos = new List<Pai>();

                var ilimite = NextInt(r, 1, largura);
                ++nivel;
                for (var i = 0; i < ilimite; i++)
                {
                    filhos.Add(new Pai(--profundidade, largura, r, nivel));
                }
            }

        }

        public int id{ get; set; }
        public string nome { get; set; }
        public string desc{ get; set; }
        public decimal valor1 { get; set; }
        public decimal valor2 { get; set; }
        public decimal valor3 { get; set; }
        public decimal valor4 { get; set; }
        public decimal valor5 { get; set; }
        public decimal valor6 { get; set; }
        public decimal valor7 { get; set; }
        public decimal valor8 { get; set; }
        public int nivel { get; set; }
        public List<Pai> filhos { get; set; }

        private int NextInt(Random rnd, int min, int max)
        {
            return rnd.Next(min,max);
        }

        private decimal NextDouble(Random rnd, double min, double max)
        {
            return (decimal)(rnd.NextDouble() * (max - min) + min);
        }

        private string GetLetter(Random r)
        {
            string chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ";
            int num = r.Next(0, chars.Length - 1);
            return chars[num].ToString();
        }

    }


}
