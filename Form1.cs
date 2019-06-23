using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ceTe.DynamicPDF;
using ceTe.DynamicPDF.Merger;
using ceTe.DynamicPDF.PageElements;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        Document document = new Document();
        ceTe.DynamicPDF.Page page;
        String cab1, cab2, cab3, cab4;
        String func1, func2, func3, func4, func5, func6;
        Double venc, desc;
        int intLinha = 710;
        int intHolerite = 0;

        public Form1()
        {
            InitializeComponent();
        }

        private String format_value(String str1)
        {
            String str2;
            int maxlen = 0;

            str2 = Convert.ToString(Convert.ToDouble(str1.Substring(0, str1.Length)));

            if (str2.Length == 1)
            {
                return "               0,00";
            }

            if (str2.Length > 3)
            {
                str2 = str2.Substring(0,str2.Length -3) + "." + str2.Substring(str2.Length - 3, 3);
            }
                    
            switch  (str2.Length)
            {
                case 3:
                    maxlen = 6;
                    break;
                case 4:
                    maxlen = 6;
                    break;
                case 5:
                    maxlen = 5;
                    break;
                case 6:
                    maxlen = 4;
                    break;
                case 7:
                    maxlen = 3;
                    break;
                case 8:
                    maxlen = 2;
                    break;
                case 9:
                    maxlen = 1;
                    break;
            }
             
            while (str2.Length < 15 + maxlen)
            {
                str2 = " " + str2;
            }
            
            if (str2.IndexOf(",") < 1 )
            {
                str2 = str2 + ",00";
            }
            return str2;
        }

        private void CriarNovaPagina()
        {
            // Create page to place the imported PDF
            page = new ceTe.DynamicPDF.Page(1404, 1404, 0);

            intHolerite++;

            this.btnCriar.Text = Convert.ToString(intHolerite);

            // Parte de cima

            // Add rectangles to show dimensions of original          
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(1, 1, 1100, 200));  // Primeiro BOX
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(950, 570, 150, 30));
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(1, 120, 1100, 450));
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(1, 200, 1100, 15));
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(1, 200, 100, 400));
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(100, 200, 600, 400));
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(800, 200, 150, 400));
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Line(1, 600, 1100, 600)); // Linha dos Totais
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(800, 600, 150, 30)); // Mensagem Valor Líquido
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(950, 600, 150, 30)); // Valor Líquido
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(1, 630, 1100, 30)); // Bases

            // Cabeçalho das verbas

            ceTe.DynamicPDF.PageElements.Label lblVerbas1 = new ceTe.DynamicPDF.PageElements.Label("CÓD.", 5, 203, 800, 80);
            ceTe.DynamicPDF.PageElements.Label lblVerbas2 = new ceTe.DynamicPDF.PageElements.Label("DESCRIÇÃO", 150, 203, 800, 80);
            ceTe.DynamicPDF.PageElements.Label lblVerbas3 = new ceTe.DynamicPDF.PageElements.Label("REF.", 710, 203, 800, 80);
            ceTe.DynamicPDF.PageElements.Label lblVerbas4 = new ceTe.DynamicPDF.PageElements.Label("VECIMENTOS", 810, 203, 800, 80);
            ceTe.DynamicPDF.PageElements.Label lblVerbas5 = new ceTe.DynamicPDF.PageElements.Label("DESCONTOS", 960, 203, 800, 80);

            page.Elements.Add(lblVerbas1);
            page.Elements.Add(lblVerbas2);
            page.Elements.Add(lblVerbas3);
            page.Elements.Add(lblVerbas4);
            page.Elements.Add(lblVerbas5);


            // Propaganda
            ceTe.DynamicPDF.PageElements.Label lblp1 = new ceTe.DynamicPDF.PageElements.Label("facilitari.com", 1, 660, 800, 80);
            ceTe.DynamicPDF.PageElements.Label lblp2 = new ceTe.DynamicPDF.PageElements.Label("facilitari.com", 1, 660 + intLinha, 800, 80);
            page.Elements.Add(lblp1);
            page.Elements.Add(lblp2);

            // Separação via empregador, via empregado

            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Line(1, 680, 1100, 680)); // Linha de separação

            // Recibo do Empregado

            ceTe.DynamicPDF.PageElements.Label lblRecibo1 = new ceTe.DynamicPDF.PageElements.Label("DECLARO TER RECEBIDO A IMPORTÂNCIA LIQUÍDA DISCRIMINADA NESSE RECIBO", 1120, 510, 800, 80);
            ceTe.DynamicPDF.PageElements.Label lblRecibo2 = new ceTe.DynamicPDF.PageElements.Label("_____/_____/_____               ____________________________________", 1150, 510, 800, 80);
            ceTe.DynamicPDF.PageElements.Label lblRecibo3 = new ceTe.DynamicPDF.PageElements.Label("     Data                                      Assinatura           ", 1180, 510, 800, 80);
            ceTe.DynamicPDF.PageElements.Label lblRecibo4 = new ceTe.DynamicPDF.PageElements.Label("                             VIA EMPREGADOR                          ", 1200, 510, 800, 80);

            lblRecibo1.Angle = -90;
            lblRecibo2.Angle = -90;
            lblRecibo3.Angle = -90;
            lblRecibo4.Angle = -90;
            lblRecibo4.FontSize = 16;

            page.Elements.Add(lblRecibo1);
            page.Elements.Add(lblRecibo2);
            page.Elements.Add(lblRecibo3);
            page.Elements.Add(lblRecibo4);

            // Recibo do empregador

            ceTe.DynamicPDF.PageElements.Label lblRecibo5 = new ceTe.DynamicPDF.PageElements.Label("DECLARO TER RECEBIDO A IMPORTÂNCIA LIQUÍDA DISCRIMINADA NESSE RECIBO", 1120, 1260, 800, 80);
            ceTe.DynamicPDF.PageElements.Label lblRecibo6 = new ceTe.DynamicPDF.PageElements.Label("_____/_____/_____               ____________________________________", 1150, 1260, 800, 80);
            ceTe.DynamicPDF.PageElements.Label lblRecibo7 = new ceTe.DynamicPDF.PageElements.Label("     Data                                      Assinatura           ", 1180, 1260, 800, 80);
            ceTe.DynamicPDF.PageElements.Label lblRecibo8 = new ceTe.DynamicPDF.PageElements.Label("                       VIA EMPREGADO", 1200, 1260, 800, 80);

            lblRecibo5.Angle = -90;
            lblRecibo6.Angle = -90;
            lblRecibo7.Angle = -90;
            lblRecibo8.Angle = -90;
            lblRecibo8.FontSize = 16;

            page.Elements.Add(lblRecibo5);
            page.Elements.Add(lblRecibo6);
            page.Elements.Add(lblRecibo7);
            page.Elements.Add(lblRecibo8);

            // Parte de baixo

            // Add rectangles to show dimensions of original          
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(1, 1 + intLinha, 1100, 200));  // Primeiro BOX
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(950, 570 + intLinha, 150, 30));
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(1, 120 + intLinha, 1100, 450));
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(1, 200 + intLinha, 1100, 15));
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(1, 200 + intLinha, 100, 400));
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(100, 200 + intLinha, 600, 400));
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(800, 200 + intLinha, 150, 400));
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Line(1, 600 + intLinha, 900, 600 + intLinha)); // Linha dos Totais
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(800, 600 + intLinha, 150, 30)); // Mensagem Valor Líquido
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(950, 600 + intLinha, 150, 30)); // Valor Líquido
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(1, 630 + intLinha, 1100, 30)); // Bases

            // Cabeçalho das verbas

            ceTe.DynamicPDF.PageElements.Label lblVerbas21 = new ceTe.DynamicPDF.PageElements.Label("CÓD.", 5, 203 + intLinha, 800, 80);
            ceTe.DynamicPDF.PageElements.Label lblVerbas22 = new ceTe.DynamicPDF.PageElements.Label("DESCRIÇÃO", 150, 203 + intLinha, 800, 80);
            ceTe.DynamicPDF.PageElements.Label lblVerbas23 = new ceTe.DynamicPDF.PageElements.Label("REF.", 710, 203 + intLinha, 800, 80);
            ceTe.DynamicPDF.PageElements.Label lblVerbas24 = new ceTe.DynamicPDF.PageElements.Label("VECIMENTOS", 810, 203 + intLinha, 800, 80);
            ceTe.DynamicPDF.PageElements.Label lblVerbas25 = new ceTe.DynamicPDF.PageElements.Label("DESCONTOS", 960, 203 + intLinha, 800, 80);

            page.Elements.Add(lblVerbas21);
            page.Elements.Add(lblVerbas22);
            page.Elements.Add(lblVerbas23);
            page.Elements.Add(lblVerbas24);
            page.Elements.Add(lblVerbas25);
            
        }

        private void btnCriar_Click(object sender, EventArgs e)
        {
            // Create a merge document and set it's properties
            document.Creator = "Visual Studio 2015";
            document.Author = "Facilitari.COM - Eduardo F. de Melo (19) 3504-1629";
            document.Title = "Holerith";

            int Linha = 0;

            try
            {
                // Create an instance of StreamReader to read from a file.
                // The using statement also closes the StreamReader.
                StreamReader sr = new StreamReader("Holerite.txt");
                try
                {
                    String line;
                    // Read and display lines from the file until the end of 
                    // the file is reached.
                    while ((line = sr.ReadLine()) != null)
                    {
                        Linha = Linha + 1;
                        if (line.Substring(0,1) == "1")
                        {
                            cab1 = line.Substring(1, 50);
                            cab2 = line.Substring(70, 35).Trim() + ", " + line.Substring(108, 114).Trim() + " " + line.Substring(223, 50).Trim() + " " + line.Substring(273, 50).Trim() + " " + line.Substring(323, 2).Trim() + " " + line.Substring(325, 9).Trim();
                            cab3 = line.Substring(50, 20).Trim();
                            cab4 = line.Substring(332, 13);
                        }

                        if (line.Substring(0,1) == "2")
                        {
                            venc = 0;
                            desc = 0;
                            CriarNovaPagina();
                            // Parte de cima
                            ceTe.DynamicPDF.PageElements.Label lbl1 = new ceTe.DynamicPDF.PageElements.Label("Recibo Pagamento - Salário", 900, 40, 200, 35);
                            lbl1.FontSize = 14;
                            page.Elements.Add(lbl1);
                            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Label(cab1, 10, 40, 300, 35)); // Nome do condomínio
                            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Label(cab2, 10, 55, 700, 35));
                            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Label(cab3, 10, 100, 700, 35));
                            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Label(cab4, 1000, 100, 700, 35));
                            // Parte de baixo
                            ceTe.DynamicPDF.PageElements.Label lbl21 = new ceTe.DynamicPDF.PageElements.Label("Recibo Pagamento - Salário", 900, 40 + intLinha, 200, 35);
                            lbl1.FontSize = 14;
                            page.Elements.Add(lbl21);
                            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Label(cab1, 10, 40 + intLinha, 300, 35)); // Nome do condomínio
                            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Label(cab2, 10, 55 + intLinha, 700, 35));
                            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Label(cab3, 10, 100 + intLinha, 700, 35));
                            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Label(cab4, 1000, 100 + intLinha, 700, 35));
                            // Parte comum
                            func1 = "Funcionário: " + line.Substring(11,80);
                            func2 = "Cargo: " + line.Substring(111,50);
                            func3 = "Departamento: ";
                            func4 = "Seção: ";
                            func5 = "Data admissão: ";
                            func6 = "Data pagamento: ";
                            // Parte de cima
                            ceTe.DynamicPDF.PageElements.Label lbl6 = new ceTe.DynamicPDF.PageElements.Label(func1, 10, 140, 400, 35);
                            ceTe.DynamicPDF.PageElements.Label lbl7 = new ceTe.DynamicPDF.PageElements.Label(func2, 600, 140, 300, 35);
                            ceTe.DynamicPDF.PageElements.Label lbl8 = new ceTe.DynamicPDF.PageElements.Label(func3, 10, 160, 300, 35);
                            ceTe.DynamicPDF.PageElements.Label lbl9 = new ceTe.DynamicPDF.PageElements.Label(func4, 600, 160, 300, 35);
                            ceTe.DynamicPDF.PageElements.Label lbl10 = new ceTe.DynamicPDF.PageElements.Label(func5, 10, 180, 300, 35);
                            ceTe.DynamicPDF.PageElements.Label lbl11 = new ceTe.DynamicPDF.PageElements.Label(func6, 600, 180, 300, 35);
                            lbl6.FontSize = 14;
                            lbl7.FontSize = 14;
                            lbl8.FontSize = 14;
                            lbl9.FontSize = 14;
                            lbl10.FontSize = 14;
                            lbl11.FontSize = 14;
                            page.Elements.Add(lbl6);
                            page.Elements.Add(lbl7);
                            page.Elements.Add(lbl8);
                            page.Elements.Add(lbl9);
                            page.Elements.Add(lbl10);
                            page.Elements.Add(lbl11);
                            // Parte de baixo
                            ceTe.DynamicPDF.PageElements.Label lbl26 = new ceTe.DynamicPDF.PageElements.Label(func1, 10, 140 + intLinha, 400, 35);
                            ceTe.DynamicPDF.PageElements.Label lbl27 = new ceTe.DynamicPDF.PageElements.Label(func2, 600, 140 + intLinha, 300, 35);
                            ceTe.DynamicPDF.PageElements.Label lbl28 = new ceTe.DynamicPDF.PageElements.Label(func3, 10, 160 + intLinha, 300, 35);
                            ceTe.DynamicPDF.PageElements.Label lbl29 = new ceTe.DynamicPDF.PageElements.Label(func4, 600, 160 + intLinha, 300, 35);
                            ceTe.DynamicPDF.PageElements.Label lbl210 = new ceTe.DynamicPDF.PageElements.Label(func5, 10, 180 + intLinha, 300, 35);
                            ceTe.DynamicPDF.PageElements.Label lbl211 = new ceTe.DynamicPDF.PageElements.Label(func6, 600, 180 + intLinha, 300, 35);
                            lbl26.FontSize = 14;
                            lbl27.FontSize = 14;
                            lbl28.FontSize = 14;
                            lbl29.FontSize = 14;
                            lbl210.FontSize = 14;
                            lbl211.FontSize = 14;
                            page.Elements.Add(lbl26);
                            page.Elements.Add(lbl27);
                            page.Elements.Add(lbl28);
                            page.Elements.Add(lbl29);
                            page.Elements.Add(lbl210);
                            page.Elements.Add(lbl211);
                        }

                        if (line.Substring(0, 1) == "3")
                        {
                            String v1, v2, v3, v4, v5, v6;

                            v1 = line.Substring(1, 4);
                            v2 = line.Substring(5, 50);
                            v3 = line.Substring(51, 11);
                            v4 = line.Substring(64, 1);
                            v5 = line.Substring(65, 15);
                            v6 = Convert.ToString(Convert.ToDouble(v5) / 100);
                            // Parte de cima
                            ceTe.DynamicPDF.PageElements.Label lblv1 = new ceTe.DynamicPDF.PageElements.Label(v1, 10, (Linha*15)+170, 400, 35);
                            ceTe.DynamicPDF.PageElements.Label lblv2 = new ceTe.DynamicPDF.PageElements.Label(v2, 120, (Linha*15)+170, 300, 35);
                            ceTe.DynamicPDF.PageElements.Label lblv3 = new ceTe.DynamicPDF.PageElements.Label(v3, 700, (Linha*15)+170, 400, 35);
                            ceTe.DynamicPDF.PageElements.Label lblv4;
                            if (v4 == "P")
                            {
                                lblv4 = new ceTe.DynamicPDF.PageElements.Label(format_value(v6), 810, (Linha*15)+170, 700, 35);
                                venc = venc + Convert.ToDouble(v6);
                            } else
                            {
                                lblv4 = new ceTe.DynamicPDF.PageElements.Label(format_value(v6), 960,(Linha*15)+170, 800, 35);
                                desc = desc + Convert.ToDouble(v6);
                            }
                            page.Elements.Add(lblv1);
                            page.Elements.Add(lblv2);
                            page.Elements.Add(lblv3);
                            page.Elements.Add(lblv4);
                            // Parte de baixo
                            ceTe.DynamicPDF.PageElements.Label lblv21 = new ceTe.DynamicPDF.PageElements.Label(v1, 10, (Linha * 15) + 170 + intLinha, 400, 35);
                            ceTe.DynamicPDF.PageElements.Label lblv22 = new ceTe.DynamicPDF.PageElements.Label(v2, 120, (Linha * 15) + 170 + intLinha, 300, 35);
                            ceTe.DynamicPDF.PageElements.Label lblv23 = new ceTe.DynamicPDF.PageElements.Label(v3, 700, (Linha * 15) + 170 + intLinha, 400, 35);
                            ceTe.DynamicPDF.PageElements.Label lblv24;
                            if (v4 == "P")
                            {
                                lblv24 = new ceTe.DynamicPDF.PageElements.Label(format_value(v6), 810, (Linha * 15) + 170 + intLinha, 700, 35);
                            }
                            else
                            {
                                lblv24 = new ceTe.DynamicPDF.PageElements.Label(format_value(v6), 960, (Linha * 15) + 170 + intLinha, 800, 35);
                            }
                            page.Elements.Add(lblv21);
                            page.Elements.Add(lblv22);
                            page.Elements.Add(lblv23);
                            page.Elements.Add(lblv24);
                        }
                 
                        if (line.Substring(0, 1) == "4")
                        {

                            // Parte de cima
                            ceTe.DynamicPDF.PageElements.Label lbltotvenc = new ceTe.DynamicPDF.PageElements.Label(Convert.ToString(venc), 810,  580, 700, 35);
                            ceTe.DynamicPDF.PageElements.Label lbltotdesc = new ceTe.DynamicPDF.PageElements.Label(Convert.ToString(desc), 960,  580, 700, 35);
                            ceTe.DynamicPDF.PageElements.Label lblmsg1 = new ceTe.DynamicPDF.PageElements.Label("Valor Líquido", 810, 610, 700, 35);
                            ceTe.DynamicPDF.PageElements.Label lbltotliq = new ceTe.DynamicPDF.PageElements.Label(format_value(Convert.ToString(venc-desc)), 960, 610, 700, 35);

                            ceTe.DynamicPDF.PageElements.Label lblb1 = new ceTe.DynamicPDF.PageElements.Label("Salário Base",           5, 630, 700, 35);
                            ceTe.DynamicPDF.PageElements.Label lblb2 = new ceTe.DynamicPDF.PageElements.Label("Salário Contr. INSS",  150, 630, 700, 35);
                            ceTe.DynamicPDF.PageElements.Label lblb3 = new ceTe.DynamicPDF.PageElements.Label("Base de Calculo FGTS", 350, 630, 700, 35);
                            ceTe.DynamicPDF.PageElements.Label lblb4 = new ceTe.DynamicPDF.PageElements.Label("FGTS do Mês",          550, 630, 700, 35);
                            ceTe.DynamicPDF.PageElements.Label lblb5 = new ceTe.DynamicPDF.PageElements.Label("Base de Calculo IRRF", 750, 630, 700, 35);
                            ceTe.DynamicPDF.PageElements.Label lblb6 = new ceTe.DynamicPDF.PageElements.Label("Dependentes IRRF",     900, 630, 700, 35);

                            // Valores 

                            ceTe.DynamicPDF.PageElements.Label lblv1 = new ceTe.DynamicPDF.PageElements.Label(format_value(Convert.ToString(Convert.ToDouble(line.Substring(1,15))/100)),    5, 645, 700, 35);
                            ceTe.DynamicPDF.PageElements.Label lblv2 = new ceTe.DynamicPDF.PageElements.Label(format_value(Convert.ToString(Convert.ToDouble(line.Substring(16,15))/100)), 150, 645, 700, 35);
                            ceTe.DynamicPDF.PageElements.Label lblv3 = new ceTe.DynamicPDF.PageElements.Label(format_value(Convert.ToString(Convert.ToDouble(line.Substring(31,15))/100)), 350, 645, 700, 35);
                            ceTe.DynamicPDF.PageElements.Label lblv4 = new ceTe.DynamicPDF.PageElements.Label(format_value(Convert.ToString(Convert.ToDouble(line.Substring(46,15))/100)), 550, 645, 700, 35);
                            ceTe.DynamicPDF.PageElements.Label lblv5 = new ceTe.DynamicPDF.PageElements.Label(format_value(Convert.ToString(Convert.ToDouble(line.Substring(61,15))/100)), 750, 645, 700, 35);
                            ceTe.DynamicPDF.PageElements.Label lblv6 = new ceTe.DynamicPDF.PageElements.Label(format_value(line.Substring(77,15)), 900, 645, 700, 35);

                            page.Elements.Add(lbltotvenc);
                            page.Elements.Add(lbltotdesc);
                            page.Elements.Add(lblmsg1);
                            page.Elements.Add(lbltotliq);

                            page.Elements.Add(lblb1);
                            page.Elements.Add(lblb2);
                            page.Elements.Add(lblb3);
                            page.Elements.Add(lblb4);
                            page.Elements.Add(lblb5);
                            page.Elements.Add(lblb6);

                            page.Elements.Add(lblv1);
                            page.Elements.Add(lblv2);
                            page.Elements.Add(lblv3);
                            page.Elements.Add(lblv4);
                            page.Elements.Add(lblv5);
                            page.Elements.Add(lblv6);

                            // Parte de baixo
                            ceTe.DynamicPDF.PageElements.Label lbltotvenc2 = new ceTe.DynamicPDF.PageElements.Label(Convert.ToString(venc), 810, 580 + intLinha, 700, 35);
                            ceTe.DynamicPDF.PageElements.Label lbltotdesc2 = new ceTe.DynamicPDF.PageElements.Label(Convert.ToString(desc), 960, 580 + intLinha, 700, 35);
                            ceTe.DynamicPDF.PageElements.Label lblmsg21 = new ceTe.DynamicPDF.PageElements.Label("Valor Líquido", 810, 610 + intLinha, 700, 35);
                            ceTe.DynamicPDF.PageElements.Label lbltotliq2 = new ceTe.DynamicPDF.PageElements.Label(format_value(Convert.ToString(venc - desc)), 960, 610 + intLinha, 700, 35);

                            ceTe.DynamicPDF.PageElements.Label lblb21 = new ceTe.DynamicPDF.PageElements.Label("Salário Base", 5, 630 + intLinha, 700, 35);
                            ceTe.DynamicPDF.PageElements.Label lblb22 = new ceTe.DynamicPDF.PageElements.Label("Salário Contr. INSS", 150, 630 + intLinha, 700, 35);
                            ceTe.DynamicPDF.PageElements.Label lblb23 = new ceTe.DynamicPDF.PageElements.Label("Base de Calculo FGTS", 350, 630 + intLinha, 700, 35);
                            ceTe.DynamicPDF.PageElements.Label lblb24 = new ceTe.DynamicPDF.PageElements.Label("FGTS do Mês", 550, 630 + intLinha, 700, 35);
                            ceTe.DynamicPDF.PageElements.Label lblb25 = new ceTe.DynamicPDF.PageElements.Label("Base de Calculo IRRF", 750, 630 + intLinha, 700, 35);
                            ceTe.DynamicPDF.PageElements.Label lblb26 = new ceTe.DynamicPDF.PageElements.Label("Dependentes IRRF", 900, 630 + intLinha, 700, 35);

                            // Valores 

                            ceTe.DynamicPDF.PageElements.Label lblv21 = new ceTe.DynamicPDF.PageElements.Label(format_value(Convert.ToString(Convert.ToDouble(line.Substring(1, 15))/100)), 5, 645 + intLinha, 700, 35);
                            ceTe.DynamicPDF.PageElements.Label lblv22 = new ceTe.DynamicPDF.PageElements.Label(format_value(Convert.ToString(Convert.ToDouble(line.Substring(16, 15))/100)), 150, 645 + intLinha, 700, 35);
                            ceTe.DynamicPDF.PageElements.Label lblv23 = new ceTe.DynamicPDF.PageElements.Label(format_value(Convert.ToString(Convert.ToDouble(line.Substring(31, 15))/100)), 350, 645 + intLinha, 700, 35);
                            ceTe.DynamicPDF.PageElements.Label lblv24 = new ceTe.DynamicPDF.PageElements.Label(format_value(Convert.ToString(Convert.ToDouble(line.Substring(46, 15))/100)), 550, 645 + intLinha, 700, 35);
                            ceTe.DynamicPDF.PageElements.Label lblv25 = new ceTe.DynamicPDF.PageElements.Label(format_value(Convert.ToString(Convert.ToDouble(line.Substring(61, 15))/100)), 750, 645 + intLinha, 700, 35);
                            ceTe.DynamicPDF.PageElements.Label lblv26 = new ceTe.DynamicPDF.PageElements.Label(format_value(Convert.ToString(Convert.ToDouble(line.Substring(77, 15))/100)), 900, 645 + intLinha, 700, 35);

                            page.Elements.Add(lbltotvenc2);
                            page.Elements.Add(lbltotdesc2);
                            page.Elements.Add(lblmsg21);
                            page.Elements.Add(lbltotliq2);

                            page.Elements.Add(lblb21);
                            page.Elements.Add(lblb22);
                            page.Elements.Add(lblb23);
                            page.Elements.Add(lblb24);
                            page.Elements.Add(lblb25);
                            page.Elements.Add(lblb26);

                            page.Elements.Add(lblv21);
                            page.Elements.Add(lblv22);
                            page.Elements.Add(lblv23);
                            page.Elements.Add(lblv24);
                            page.Elements.Add(lblv25);
                            page.Elements.Add(lblv26);

                            // Add page to document
                            document.Pages.Add(page);
                            Linha = 1;
                        }
                    }
                }
                finally
                {
                    sr.Dispose();
                }
            }
            catch (System.Exception ee)
            {
                // Let the user know what went wrong.
                Console.WriteLine("The file could not be read:");
                Console.WriteLine(ee.Message.ToString());
            }

            // Outputs
            document.Draw( "Holerite.pdf" ); 
        }
    }
}