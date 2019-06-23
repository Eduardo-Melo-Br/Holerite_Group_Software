using System;
using System.IO; 
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ceTe.DynamicPDF;
using ceTe.DynamicPDF.Merger;
using ceTe.DynamicPDF.PageElements;
using ceTe.DynamicPDF.Text;
namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        Document document;
        ceTe.DynamicPDF.Page page;
        String cab1, cab2, cab3, cab4, cab5;
        String func1="Primeiro", func2, func3, func4, func5, func6;
        String nroBanco, nomBanco, ageBanco, ccBanco, cpfCli;
        Double venc, desc;
        String[] strParametrosCSV = new String[250];
        String strSalario_Base;
        String strBASE_INSS_ATE_O_TETO;
        String strBASE_CALCULO_FGTS;
        String strVALOR_FGTS;
        String strBASE_IRRF;
        String strDEPENDENTE_DE_IRRF;
        String strSALARIO_FAMILIA;
        String strNovaFolha = "N";
        String strUltimoHolerite = "S";
        String strDesconto = "";
        double vlr_liquido;
        int Linha = 3;
        int intTipoDoRegistro = 1;
        Boolean booNovoCondominio = true;
        // Matrizes dos créditos e débitos
        String[] creditosCodigo = new String[100];
        String[] creditosVerba = new String[100];
        String[] creditosReferencia = new String[100];
        String[] creditosValor = new String[100];
        String[] debitosCodigo = new String[100];
        String[] debitosVerba = new String[100];
        String[] debitosReferencia = new String[100];
        String[] debitosValor = new String[100];
        // Totalizador de créditos e débitos
        int totalItensCreditos;
        int TotalItensDebitos;
        // Dados do salário base
        String salarioBaseCodigo;
        String salarioBaseVerba;
        String salarioBaseReferencia;
        String salarioBaseValor;
        String strCargo = "";
        Boolean bPagina;
        OleDbConnection myConnection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.; Extended Properties = dBASE IV;"); // Conexão com a Tabela de Dados
        OleDbCommand _CondominiosCommand = new OleDbCommand();

        // String sDtAdmissao, sDtPagamento;

        int intLinha = 971; // 851, 951
        int intMargem = 85;

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            MessageBox.Show("Entre no sistema RHCorp, para gerar o arquivo de dados, salve na pasta e depois clique no botão ao lado para criar o arquivo PDF.");
        }

        int intHolerite = 0;

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Telefones: (19) 99648-6457 / 3897-4477 e-Mail efmelo@outlookcom Web Site: softwareandfaith.info");
        }

        int intMargemSuperior = 0;

        private void buttonTXTPDFMULTIPLO_Click(object sender, EventArgs e)
        {
            string[] files = {"1","2","3"};

            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                MessageBox.Show("Pasta selecionada: " + folderBrowserDialog1.SelectedPath);
            }
            else
            {
                MessageBox.Show("ERRO");
            }
                
            try
            {
                // Exception could occur due to insufficient permission.
                files = Directory.GetFiles(folderBrowserDialog1.SelectedPath, "Hole*.txt", SearchOption.TopDirectoryOnly);
            }
            catch (Exception)
            {
                MessageBox.Show("Não encontrei arquivos.");
            }

            // If matching files have been found, return the first one.
            if (files.Length > 0)
            {
                for (int i=0;i<=files.Length -1;i++)
                {
                    CriarPDF(files[i]);
                }
            }
        }


        public Form1()
        {
            InitializeComponent();
            myConnection.Open();
        }

        private String format_value(String str1)
        {
            String str2;

            str2 = Convert.ToString(Convert.ToDouble(str1.Substring(0, str1.Length)));

            if (str2.Length == 1)
            {
                return "               0,00";
            }
      
            while (str2.Length < 16)
            {
                str2 = " " + str2;
            }
            
            if (str2.IndexOf(",") < 1 )
            {
                str2 = str2 + ",00";
            }

            if (str2.Length <= str2.IndexOf(",") + 2)
            {
                str2 = str2 + "0";
            }

            return str2;
        }

        private void CriarPaginaDoVerso()
        {
            // Create page to place the PDF
            ceTe.DynamicPDF.Page pageVerso = new ceTe.DynamicPDF.Page(1404, 2000, 1);

            ceTe.DynamicPDF.PageElements.Label lbl1 = new ceTe.DynamicPDF.PageElements.Label(T_Unicode(cab1), 1404 -  intMargem, 1940, 700, 35); // Nome do condomínio
            ceTe.DynamicPDF.PageElements.Label lbl6 = new ceTe.DynamicPDF.PageElements.Label(func1, 1404 - intMargem, 1980, 500, 35); // Nome do funcionário
            ceTe.DynamicPDF.PageElements.Label lblTipo = new ceTe.DynamicPDF.PageElements.Label(" ", 700, 35, 400, 50);

            foreach (int indexChecked in checkedListBoxMSGTipoPagamento.CheckedIndices)
            {
                // The indexChecked variable contains the index of the item.
                lblTipo = new ceTe.DynamicPDF.PageElements.Label(checkedListBoxMSGTipoPagamento.Items[indexChecked].ToString() + " Referênte à " + T_Unicode(cab4) + "/" + cab5, 1404 - intMargem, 1900, 700, 35); // Tipo de Pagamento
            }
            lbl1.FontSize = 18;
            lbl6.FontSize = 18;
            lblTipo.FontSize = 18;
            lbl1.Angle = 180;
            lbl6.Angle = 180;
            lblTipo.Angle = 180;
            pageVerso.Elements.Add(lbl1);
            pageVerso.Elements.Add(lbl6);
            pageVerso.Elements.Add(lblTipo);
            document.Pages.Add(pageVerso);
        }


        private void CriarNovaPagina()
        {
            // Create page to place the PDF
            page = new ceTe.DynamicPDF.Page(1404, 2100, 1);

            if (strNovaFolha == "N")
            {
                intHolerite++;
            }
            
            this.btnCriar.Text = "Gerei " + Convert.ToString(intHolerite) + " Holerites.";

            // Linha de apoio a leitura visual
            // Linha de fundo para facilitar a visualização

            for (int intVisual=13;intVisual<44;intVisual=intVisual+2)
            {
                ceTe.DynamicPDF.PageElements.Rectangle retangulo1 = new ceTe.DynamicPDF.PageElements.Rectangle(10 + intMargem, intVisual * 19 - 11, 1100, 25, Grayscale.White, RgbColor.Snow, 1, LineStyle.DashSmall);
                ceTe.DynamicPDF.PageElements.Rectangle retangulo2 = new ceTe.DynamicPDF.PageElements.Rectangle(10 + intMargem, intVisual * 19 - 11 + intLinha, 1100, 25, Grayscale.White, RgbColor.Snow, 1, LineStyle.DashSmall);
                page.Elements.Add(retangulo1);
                page.Elements.Add(retangulo2);

            }

            // Parte de cima

            // Add rectangles to show dimensions of original          
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(1   + intMargem, 3 + intMargemSuperior,  1160, 220 + intMargemSuperior));              // Primeiro BOX
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(1   + intMargem, 120, 1160, 790));             // BOX do corpo
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Line(50 + intMargem, 200, 50 + intMargem, 810));         // Linha das referências vertical
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Line(1 + intMargem, 810, 890, 810));                     // Linha do final das verbas
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Line(710 + intMargem, 202, 710 + intMargem, 870));       // Linha vertical da referências
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Line(1   + intMargem, 200, 1160 + intMargem, 200));      // Linha Cabeçalho
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(800 + intMargem, 810, 361, 31));               // Box valor liquido
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(800 + intMargem, 200, 150, 670));              // Mensagem Valor Líquido
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(1161 + intMargem,  3, 125, 907));              // Recibo
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(1   + intMargem, 870, 1160, 40));              // Box dos Totais

            // Cabeçalho das verbas

            ceTe.DynamicPDF.PageElements.Label lblVerbas1 = new ceTe.DynamicPDF.PageElements.Label("CÓD.", 5 + intMargem, 203, 800, 80);
            ceTe.DynamicPDF.PageElements.Label lblVerbas2 = new ceTe.DynamicPDF.PageElements.Label("DESCRIÇÃO", 370 + intMargem, 203, 800, 80);
            ceTe.DynamicPDF.PageElements.Label lblVerbas3 = new ceTe.DynamicPDF.PageElements.Label("REF.", 730 + intMargem, 203, 800, 80);
            ceTe.DynamicPDF.PageElements.Label lblVerbas4 = new ceTe.DynamicPDF.PageElements.Label("VENCIMENTOS", 810 + intMargem, 203, 800, 80);
            ceTe.DynamicPDF.PageElements.Label lblVerbas5 = new ceTe.DynamicPDF.PageElements.Label("DESCONTOS", 980 + intMargem, 203, 800, 80);

            lblVerbas1.FontSize = 16;
            lblVerbas2.FontSize = 16;
            lblVerbas3.FontSize = 16;
            lblVerbas4.FontSize = 16;
            lblVerbas5.FontSize = 16;

            page.Elements.Add(lblVerbas1);
            page.Elements.Add(lblVerbas2);
            page.Elements.Add(lblVerbas3);
            page.Elements.Add(lblVerbas4);
            page.Elements.Add(lblVerbas5);


            // Suporte   
            ceTe.DynamicPDF.PageElements.Label lblp1 = new ceTe.DynamicPDF.PageElements.Label("Tecnologia: (19) 3897-4477", 10 + intMargem, 914, 800, 5);
            ceTe.DynamicPDF.PageElements.Label lblp2 = new ceTe.DynamicPDF.PageElements.Label("Tecnologia: (19) 3897-4477", 10 + intMargem, 914 + intLinha, 800, 5);
            lblp1.FontSize = 7;
            lblp2.FontSize = 7;
            page.Elements.Add(lblp1);
            page.Elements.Add(lblp2);

            // Recibo do Empregador

            ceTe.DynamicPDF.PageElements.Label lblr1 = new ceTe.DynamicPDF.PageElements.Label("DECLARO TER RECEBIDO A IMPORTÂNCIA LIQUÍDA DISCRIMINADA NESTE RECIBO", 1180 + intMargem, 770, 800, 80);
            lblr1.FontSize = 16;

            ceTe.DynamicPDF.PageElements.Label lblRecibo1 = lblr1;
            ceTe.DynamicPDF.PageElements.Label lblRecibo2 = new ceTe.DynamicPDF.PageElements.Label("..................../..................../....................               ..............................................................................................................", 1225 + intMargem, 760, 800, 80);
            ceTe.DynamicPDF.PageElements.Label lblRecibo3 = new ceTe.DynamicPDF.PageElements.Label("                         Data                                                                                        Assinatura", 1248 + intMargem, 760, 800, 80);
            ceTe.DynamicPDF.PageElements.Label lblRecibo4 = new ceTe.DynamicPDF.PageElements.Label("                             VIA EMPREGADOR", 1262 + intMargem, 730, 800, 80);

            lblRecibo1.Angle = -90;
            lblRecibo2.Angle = -90;
            lblRecibo3.Angle = -90;
            lblRecibo4.Angle = -90;
            lblRecibo4.FontSize = 15;

            page.Elements.Add(lblRecibo1);
            page.Elements.Add(lblRecibo2);
            page.Elements.Add(lblRecibo3);
            page.Elements.Add(lblRecibo4);

            // Recibo do empregado

            ceTe.DynamicPDF.PageElements.Label lblr5 = new ceTe.DynamicPDF.PageElements.Label("DECLARO TER RECEBIDO A IMPORTÂNCIA LIQUÍDA DISCRIMINADA NESTE RECIBO", 1180 + intMargem, 770 + intLinha, 800, 80);
            lblr5.FontSize = 16;

            ceTe.DynamicPDF.PageElements.Label lblRecibo5 = lblr5;
            ceTe.DynamicPDF.PageElements.Label lblRecibo6 = new ceTe.DynamicPDF.PageElements.Label("..................../..................../....................               ..............................................................................................................", 1230 + intMargem, 790 + intLinha, 800, 80);
            ceTe.DynamicPDF.PageElements.Label lblRecibo7 = new ceTe.DynamicPDF.PageElements.Label("               Data                                                                                                  Assinatura", 1248 + intMargem, 760 + intLinha, 800, 80);
            ceTe.DynamicPDF.PageElements.Label lblRecibo8 = new ceTe.DynamicPDF.PageElements.Label("                             VIA EMPREGADO", 1262 + intMargem, 730 + intLinha, 800, 80);

            lblRecibo5.Angle = -90;
            lblRecibo6.Angle = -90;
            lblRecibo7.Angle = -90;
            lblRecibo8.Angle = -90;
            lblRecibo8.FontSize = 15;

            page.Elements.Add(lblRecibo5);
            page.Elements.Add(lblRecibo6);
            page.Elements.Add(lblRecibo7);
            page.Elements.Add(lblRecibo8);

            // Parte de baixo

            // Add rectangles to show dimensions of original          
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(1 + intMargem, 3 + intLinha, 1160, 220));  // Primeiro BOX
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(1 + intMargem, 120 + intLinha, 1160, 790)); // BOX do corpo
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Line(50 + intMargem, 200 + intLinha, 50 + intMargem, 810 + intLinha)); // Linha das referências vertical
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Line(1 + intMargem, 810 + intLinha, 890, 810 + intLinha));             // Linha do final das verbas 
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Line(710 + intMargem, 202 + intLinha, 710 + intMargem, 870 + intLinha));        // Linha vertical da referências
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Line(1 + intMargem, 200 + intLinha, 1160 + intMargem, 200 + intLinha));         // Linha Cabeçalho
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(800 + intMargem, 810 + intLinha, 361, 31));                                      // Box valor liquido
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(800 + intMargem, 200 + intLinha, 150, 670));              // Mensagem Valor Líquido
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(800 + intMargem, 840 + intLinha, 361, 31));                            // Bases
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(1161 + intMargem, 3 + intLinha, 125, 907));                           // Recibo
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(1 + intMargem, 870 + intLinha, 1160, 40));                            // Box dos Totais
            
            // Cabeçalho das verbas

            ceTe.DynamicPDF.PageElements.Label lblVerbas21 = new ceTe.DynamicPDF.PageElements.Label("CÓD.", 5 + intMargem, 203 + intLinha, 800, 80);
            ceTe.DynamicPDF.PageElements.Label lblVerbas22 = new ceTe.DynamicPDF.PageElements.Label("DESCRIÇÃO", 370 + intMargem, 203 + intLinha, 800, 80);
            ceTe.DynamicPDF.PageElements.Label lblVerbas23 = new ceTe.DynamicPDF.PageElements.Label("REF.", 730 + intMargem, 203 + intLinha, 800, 80);
            ceTe.DynamicPDF.PageElements.Label lblVerbas24 = new ceTe.DynamicPDF.PageElements.Label("VENCIMENTOS", 810 + intMargem, 203 + intLinha, 800, 80);
            ceTe.DynamicPDF.PageElements.Label lblVerbas25 = new ceTe.DynamicPDF.PageElements.Label("DESCONTOS", 980 + intMargem, 203 + intLinha, 800, 80);

            page.Elements.Add(lblVerbas21);
            page.Elements.Add(lblVerbas22);
            page.Elements.Add(lblVerbas23);
            page.Elements.Add(lblVerbas24);
            page.Elements.Add(lblVerbas25);
        }
        
        private void LOG_Geracao(string strQuantidade, string strCondominio)
        {
            _CondominiosCommand.Connection = myConnection;
            _CondominiosCommand.CommandText = "INSERT INTO LOG (DATA_HORA, QUANTIDADE, COND) VALUES ('" + DateTime.Now.ToString() + "'," + strQuantidade + ",'" + strCondominio + "');";
            _CondominiosCommand.ExecuteNonQuery();
        }
        private string T_Unicode(string lbl)
        {
            String nresult = "";
            // Mêses
            if (lbl.IndexOf("Mar" + Convert.ToChar(65533) + "o") >= 0)
            {
                nresult += "Março ";
            }

            if (nresult.Length == 0)
            {
                nresult = lbl;
            }
            return nresult;
        }

        private void Limpa_Bases_de_Calculo()
        {
            // Base de cálculo
            strSalario_Base = "0";
            strBASE_INSS_ATE_O_TETO = "0";
            strBASE_CALCULO_FGTS = "0";
            strVALOR_FGTS = "0";
            strBASE_IRRF = "0";
            strDEPENDENTE_DE_IRRF = "0";
            strSALARIO_FAMILIA = "0";

        }

        private String Formata_Valor(String p1)
        {
            String r1 = "";

            if (Convert.ToDouble(p1) > 999.99)
            {
                p1 = p1.Trim();
                String s1 = p1.Substring(0,p1.Length - 6);
                String s2 = p1.Substring(p1.Length-6,6);
                r1 = s1 + "." + s2;
            }
            else
            {
                r1 = p1;
            }
            return r1;
        }
        
        private void CriarPDF(string PDFname)
        {
            
            // Create a merge document and set it's properties
            document = new Document();
            document.Creator = "Visual Studio 2015";
            document.Author = "efmelo@outlook.com (19) 3897-4477";
            document.Title = "Holerite";

            intHolerite = 0;
            
            try
            {
                // Create an instance of StreamReader to read from a file.
                // The using statement also closes the StreamReader.
                StreamReader sr = new StreamReader(PDFname);
                try
                {
                    String line;
                    
                    int intParameterCount = 0;
                    String strParameter;
                    
                    // Read and display lines from the file until the end of 
                    // the file is reached.
                    Linha = 3;
                    Limpa_Bases_de_Calculo();
                    strDesconto = "N";
                    strUltimoHolerite = "S";
                    booNovoCondominio = true;
                    while ((line = sr.ReadLine()) != null)
                    {
                        
                        // Limpar os parâmetros
                        for (int param = 0; param < 250; param++)
                        {
                            strParametrosCSV[param] = null;
                        }
                        intParameterCount = 0;
                        strParameter = "";

                        for (int charactercounter = 0;charactercounter<line.Length; charactercounter++)
                        {
                            String strCaracter1;
                            strCaracter1 = line.Substring(charactercounter,1);
                            Char c2 = (char)34;
                            String strCompare1 = Convert.ToString(c2);
                            if (strCaracter1.CompareTo(";")!=0 && strCaracter1.CompareTo(c2+"")!=0)
                            {
                                strParameter = String.Concat(strParameter, strCaracter1);
                            }
                            if (strCaracter1.CompareTo(";")==0 || charactercounter == line.Length - 2)
                            {
                                strParametrosCSV[intParameterCount] = strParameter;
                                intParameterCount++;
                                strParameter = "";
                            }
                        }
                        // Verificação se é o inicio de um novo Holerite, se for, imprime os dados do holerite anterior
                        if (func1 != "Funcionário: " + this.strParametrosCSV[4 + 15] && func1 != "Primeiro")
                        {
                            acaoImprimirVerbas();
                            // É uma nova linha com dados dos funcionário
                            acaoNovoFuncionario();
                            booNovoCondominio = true;
                        }
                        // Verificação se o valor está nulo
                        if (strParametrosCSV[44+15] == null)
                        {
                            strParametrosCSV[44+15] = "";
                        }
                        // Verificação se o valor está nulo
                        if (strParametrosCSV[55+15] == null)
                        {
                            strParametrosCSV[55+15] = "0";
                        }
                        /*
                        if (intHolerite == 12)
                        {
                            MessageBox.Show("63 : " + strParametrosCSV[50 + 13] + " 64 : " + strParametrosCSV[50 + 14] + " 65 : " + strParametrosCSV[50 + 15] + " 66 : " + strParametrosCSV[50 + 16]);
                        }
                        */
                        // Bases de Cálculos
                        if (Convert.ToInt16("0"+strParametrosCSV[50+14])==1) // Salário Base
                        {
                            strSalario_Base = strParametrosCSV[68];
                        }
                        if (Convert.ToInt16("0" + strParametrosCSV[58])==15) // Salário Liquido
                        {
                            vlr_liquido = Convert.ToDouble(strParametrosCSV[62])/100;
                        }
                        if (strParametrosCSV[50+15] == "BASE INSS ATE O TETO")
                        {
                            strBASE_INSS_ATE_O_TETO = strParametrosCSV[53+15];
                        }
                        if (strParametrosCSV[50+15] == "BASE FGTS")
                        {
                            strBASE_CALCULO_FGTS = strParametrosCSV[53+15];
                        }
                        if (strParametrosCSV[50+15] == "FGTS")
                        {
                            strVALOR_FGTS = strParametrosCSV[53+15];
                        }
                        if (strParametrosCSV[50+15] == "BASE IRRF")
                        {
                            strBASE_IRRF = strParametrosCSV[53+15];
                        }
                        if (strParametrosCSV[50+15] == "DEPENDENTE DE IRRF")
                        {
                           strDEPENDENTE_DE_IRRF=strParametrosCSV[52+15];
                           strSALARIO_FAMILIA=strParametrosCSV[53+15];
                        }
                        if (booNovoCondominio)
                        {
                            bPagina = false;
                            cab1 = this.strParametrosCSV[12 + 15]; // Nome do Condomínio
                            cab2 = this.strParametrosCSV[1] + ", " + this.strParametrosCSV[2] + ", " + this.strParametrosCSV[3]; // Endereço, número, bairro
                            cab3 = this.strParametrosCSV[0]; // CNPJ DO CONDOMÍNIO
                            cab4 = this.comboBoxMes.SelectedItem.ToString();
                            cab5 = this.comboBoxAno.SelectedItem.ToString();
                            intTipoDoRegistro = 2;
                            booNovoCondominio = false;
                            this.strUltimoHolerite = "S";
                        }
                        if (intTipoDoRegistro == 2)
                        {
                            if (strNovaFolha == "N")
                            {
                                venc = 0;
                                desc = 0;
                            }
                            // Verificar se buscou do arquivo de dados e se está nulo
                            if (cab1 == null)
                            {
                                cab1 = "";
                            }
                            if (cab2 == null)
                            {
                                cab2 = "";
                            }
                            if (cab3 == null)
                            {
                                cab3 = "";
                            }
                            if (cab4 == null)
                            {
                                cab4 = "";
                            }
                            // Parte comum
                            if (strParametrosCSV[10+15] == null)
                            {
                                strParametrosCSV[10+15] = "  /  /    ";
                            }
                            if (strParametrosCSV[11+15] == null)
                            {
                                strParametrosCSV[11+15] = "  /  /    ";
                            }
                            func1 = "Funcionário: " + this.strParametrosCSV[4+15];
                            func2 = "Cargo: " + this.strParametrosCSV[13+15];
                            func3 = "Departamento: " + this.strParametrosCSV[12+15];
                            func4 = "Seção: ";
                            func5 = "Data admissão: " + this.strParametrosCSV[10+15];
                            func6 = "Data pagamento: " + this.strParametrosCSV[30+15];
                            // Cargo Parte de cima
                            
                            strCargo = strParametrosCSV[14+15];
                            // Cargo Parte de baixo
                            strCargo = strParametrosCSV[14+15];
                            ceTe.DynamicPDF.PageElements.Label lblCargo2 = new ceTe.DynamicPDF.PageElements.Label("Cargo: " + strCargo, 30 + intMargem, 815 + intLinha, 800, 35);
                            // Dados bancários
                            nroBanco = this.strParametrosCSV[13];
                            nomBanco = this.strParametrosCSV[14];
                            ageBanco = this.strParametrosCSV[10] + "-" + this.strParametrosCSV[11] + "-" + this.strParametrosCSV[12];
                            ccBanco = this.strParametrosCSV[8] + "-" + this.strParametrosCSV[9];
                            this.cpfCli = this.strParametrosCSV[7];
                            acaoImprimirEtapa2();
                            intTipoDoRegistro = 3;
                        }
                        if (intTipoDoRegistro == 3)
                        {
                            String v1, v2, v5, v6, strValorVerba;
                            // O arquivo, contêm duas colunas de dados, na primeira estão os Créditos e na segunda estão os Débito
                            v1 = this.strParametrosCSV[37+15]; // Código da verba
                            v2 = this.strParametrosCSV[38+15]; // Descrição da verba
                            v5 = this.strParametrosCSV[40+15]; // Referência do Valor
                            strValorVerba = this.strParametrosCSV[41+15];
                            v6 = Convert.ToString(Convert.ToDouble(strValorVerba)/100);
                            // Parte de cima
                            if (v1 == null || v1 == "")
                            {
                                v1 = "0";
                            }
                            if (Convert.ToDouble(v6) != 0 && strParametrosCSV[57] != "T")
                            {
                                // Alimentar vetor de Créditos
                                if (Convert.ToInt16(v1) == 1)
                                {
                                    // Salário Mensal
                                    salarioBaseCodigo = v1;
                                    salarioBaseVerba = v2;
                                    salarioBaseReferencia = v5;
                                    salarioBaseValor = v6;
                                }
                                else
                                {
                                    totalItensCreditos++;
                                    creditosCodigo[totalItensCreditos - 1] = v1; // Código da verba
                                    creditosVerba[totalItensCreditos - 1] = v2; // Descrição da verba
                                    creditosReferencia[totalItensCreditos - 1] = v5;
                                    creditosValor[totalItensCreditos - 1] = v6;
                                }
                            }
                            strDesconto = "N";
                            // Descontos
                            // O arquivo, contêm duas colunas de dados, na primeira estão os Créditos e na segunda estão os Débito
                            v1 = this.strParametrosCSV[58]; // 42+15Código da verba
                            v2 = this.strParametrosCSV[59]; // Descrição da verba
                            v5 = this.strParametrosCSV[61]; // Valor de Referência
                            strValorVerba = this.strParametrosCSV[62];
                            v6 = Convert.ToString(Convert.ToDouble(strValorVerba) / 100);
                            String strTipoPagamento = this.strParametrosCSV[63];
                            if (Convert.ToDouble(v6) > 0 && strTipoPagamento!="T") // Linha de Débitos
                            {
                                TotalItensDebitos++;
                                debitosCodigo[TotalItensDebitos - 1] = v1; // Código da verba
                                debitosVerba[TotalItensDebitos - 1] = v2; // Descrição da verba
                                debitosReferencia[TotalItensDebitos - 1] = v5;
                                debitosValor[TotalItensDebitos - 1] = v6;
                            } // Linha de débitos
                        }
                    }
                }
                catch (System.Exception eee)
                {
                    MessageBox.Show(eee.Message.ToString());
                }
                finally
                {
                    //sr.Dispose();
                }
            }
            catch (System.Exception ee)
            {
                // Let the user know what went wrong.
                Console.WriteLine("The file could not be read:");
                Console.WriteLine(ee.Message.ToString());
            }

            PDFname = PDFname.Substring(0, PDFname.Length - 4);

            // Outputs

            if (strUltimoHolerite == "S")
            {
                acaoImprimirVerbas();
                this.acaoTotalizadores("");
                Linha = 3;
                CriarPaginaDoVerso();
            }
            document.Draw(PDFname + ".pdf");
            LOG_Geracao(intHolerite.ToString(), cab1);
            MessageBox.Show("Gerei o arquivo " + PDFname + ".pdf");
            //System.Diagnostics.Process.Start("acrord32 " + PDFname + ".pdf");
        }

        private void acaoImprimirVerbas()
        {
            int intProximaPagina = 0;
            // Objetos de Crédito
            ceTe.DynamicPDF.PageElements.Label lblv1;
            ceTe.DynamicPDF.PageElements.Label lblv2;
            ceTe.DynamicPDF.PageElements.Label lblv3;
            ceTe.DynamicPDF.PageElements.Label lblv4;
            // Objetos de Débito
            ceTe.DynamicPDF.PageElements.Label lblv21;
            ceTe.DynamicPDF.PageElements.Label lblv22;
            ceTe.DynamicPDF.PageElements.Label lblv23;
            ceTe.DynamicPDF.PageElements.Label lblv24;
            // Imprimir Salário Base
            if (this.salarioBaseValor != "0.00" && this.salarioBaseValor != null && intProximaPagina==0)
            {
                // Imprimir Salário Base Primeiro
                Linha = Linha + 1;
                lblv1 = new ceTe.DynamicPDF.PageElements.Label(this.salarioBaseCodigo, 15 + intMargem, (Linha * 19) + 165, 400, 35);
                lblv2 = new ceTe.DynamicPDF.PageElements.Label(this.salarioBaseVerba, 70 + intMargem, (Linha * 19) + 165, 300, 35);
                lblv3 = new ceTe.DynamicPDF.PageElements.Label(this.salarioBaseReferencia, 715 + intMargem, (Linha * 19) + 165, 400, 35); // 730
                lblv4 = new ceTe.DynamicPDF.PageElements.Label(Formata_Valor(format_value(this.salarioBaseValor)), 780 + intMargem, (Linha * 19) + 165, 700, 35);
                venc = venc + Convert.ToDouble(this.salarioBaseValor);
                // Ajustar a fonte
                lblv1.FontSize = 12;
                lblv2.FontSize = 12;
                lblv3.FontSize = 12;
                lblv4.FontSize = 12;
                // Ajustar o Alinhamneto da fonte
                lblv3.Width = 40;
                lblv4.Width = 120;
                lblv3.Align = ceTe.DynamicPDF.TextAlign.Right;
                lblv4.Align = ceTe.DynamicPDF.TextAlign.Right;
                // Adicionar na página
                page.Elements.Add(lblv1);
                page.Elements.Add(lblv2);
                page.Elements.Add(lblv3);
                page.Elements.Add(lblv4);
                // Parte de baixo
                lblv21 = new ceTe.DynamicPDF.PageElements.Label(this.salarioBaseCodigo, 15 + intMargem, (Linha * 19) + 165 + intLinha, 400, 35);
                lblv22 = new ceTe.DynamicPDF.PageElements.Label(this.salarioBaseVerba, 70 + intMargem, (Linha * 19) + 165 + intLinha, 300, 35);
                lblv23 = new ceTe.DynamicPDF.PageElements.Label(this.salarioBaseReferencia, 715 + intMargem, (Linha * 19) + 165 + intLinha, 400, 35); // 730
                lblv24 = new ceTe.DynamicPDF.PageElements.Label(Formata_Valor(format_value(this.salarioBaseValor)), 780 + intMargem, (Linha * 19) + 165 + intLinha, 700, 35);
                // Ajustar o tamanho da fonte
                lblv21.FontSize = 12;
                lblv22.FontSize = 12;
                lblv23.FontSize = 12;
                lblv24.FontSize = 12;
                // Ajustar o tipo da fonte
                lblv23.Width = 40;
                lblv24.Width = 120;
                lblv23.Align = ceTe.DynamicPDF.TextAlign.Right;
                lblv24.Align = ceTe.DynamicPDF.TextAlign.Right;
                // Adicionar na página
                page.Elements.Add(lblv21);
                page.Elements.Add(lblv22);
                page.Elements.Add(lblv23);
                page.Elements.Add(lblv24);
                venc = venc + Convert.ToDouble(this.salarioBaseValor);
                this.salarioBaseValor = "0.00"; // Zerar o Salário
            }
            // Imprimir Créditos
            for (int intCreditos = 1; intCreditos <= this.totalItensCreditos; intCreditos++)
            {
                Linha = Linha + 1;
                lblv1 = new ceTe.DynamicPDF.PageElements.Label(this.creditosCodigo[intCreditos - 1], 15 + intMargem, (Linha * 19) + 165, 400, 35);
                lblv2 = new ceTe.DynamicPDF.PageElements.Label(this.creditosVerba[intCreditos - 1], 70 + intMargem, (Linha * 19) + 165, 300, 35);
                lblv3 = new ceTe.DynamicPDF.PageElements.Label(this.creditosReferencia[intCreditos - 1], 715 + intMargem, (Linha * 19) + 165, 400, 35); // 730
                lblv4 = new ceTe.DynamicPDF.PageElements.Label(Formata_Valor(format_value(this.creditosValor[intCreditos - 1])), 780 + intMargem, (Linha * 19) + 165, 700, 35);
                venc = venc + Convert.ToDouble(this.creditosValor[intCreditos - 1]);
                // Ajustar a fonte
                lblv1.FontSize = 12;
                lblv2.FontSize = 12;
                lblv3.FontSize = 12;
                lblv4.FontSize = 12;
                // Ajustar o Alinhamneto da fonte
                lblv3.Width = 40;
                lblv4.Width = 120;
                lblv3.Align = ceTe.DynamicPDF.TextAlign.Right;
                lblv4.Align = ceTe.DynamicPDF.TextAlign.Right;
                // Adicionar na página
                page.Elements.Add(lblv1);
                page.Elements.Add(lblv2);
                page.Elements.Add(lblv3);
                page.Elements.Add(lblv4);
                // Parte de baixo
                lblv21 = new ceTe.DynamicPDF.PageElements.Label(this.creditosCodigo[intCreditos - 1], 15 + intMargem, (Linha * 19) + 165 + intLinha, 400, 35);
                lblv22 = new ceTe.DynamicPDF.PageElements.Label(this.creditosVerba[intCreditos - 1], 70 + intMargem, (Linha * 19) + 165 + intLinha, 300, 35);
                lblv23 = new ceTe.DynamicPDF.PageElements.Label(this.creditosReferencia[intCreditos - 1], 715 + intMargem, (Linha * 19) + 165 + intLinha, 400, 35); // 730
                lblv24 = new ceTe.DynamicPDF.PageElements.Label(Formata_Valor(format_value(this.creditosValor[intCreditos - 1])), 780 + intMargem, (Linha * 19) + 165 + intLinha, 700, 35);
                // Ajustar o tamanho da fonte
                lblv21.FontSize = 12;
                lblv22.FontSize = 12;
                lblv23.FontSize = 12;
                lblv24.FontSize = 12;
                // Ajustar o tipo da fonte
                lblv23.Width = 40;
                lblv24.Width = 120;
                lblv23.Align = ceTe.DynamicPDF.TextAlign.Right;
                lblv24.Align = ceTe.DynamicPDF.TextAlign.Right;
                // Adicionar na página
                page.Elements.Add(lblv21);
                page.Elements.Add(lblv22);
                page.Elements.Add(lblv23);
                page.Elements.Add(lblv24);
                if (Linha == 33)
                {
                    acaoTotalizadores("Próxima Página");
                    Linha = 3;
                    CriarPaginaDoVerso();
                    booNovoCondominio = true;
                    strNovaFolha = "S";
                    acaoImprimirEtapa2();

                }
            }
            this.totalItensCreditos = 0;
            // Imprimir Débitos
            for (int intDebitos = 1; intDebitos <= this.TotalItensDebitos; intDebitos++)
            {
                Linha = Linha + 1;
                lblv1 = new ceTe.DynamicPDF.PageElements.Label(this.debitosCodigo[intDebitos - 1], 15 + intMargem, (Linha * 19) + 165, 400, 35);
                lblv2 = new ceTe.DynamicPDF.PageElements.Label(this.debitosVerba[intDebitos - 1], 70 + intMargem, (Linha * 19) + 165, 300, 35);
                lblv3 = new ceTe.DynamicPDF.PageElements.Label(this.debitosReferencia[intDebitos - 1], 715 + intMargem, (Linha * 19) + 165, 400, 35); // 730
                lblv4 = new ceTe.DynamicPDF.PageElements.Label(Formata_Valor(format_value(this.debitosValor[intDebitos - 1])), 980 + intMargem, (Linha * 19) + 165, 700, 35);
                desc = desc + Convert.ToDouble(this.debitosValor[intDebitos - 1]);
                // Ajustar a fonte
                lblv1.FontSize = 12;
                lblv2.FontSize = 12;
                lblv3.FontSize = 12;
                lblv4.FontSize = 12;
                // Ajustar o Alinhamneto da fonte
                lblv3.Width = 40;
                lblv4.Width = 120;
                lblv3.Align = ceTe.DynamicPDF.TextAlign.Right;
                lblv4.Align = ceTe.DynamicPDF.TextAlign.Right;
                // Adicionar na página
                page.Elements.Add(lblv1);
                page.Elements.Add(lblv2);
                page.Elements.Add(lblv3);
                page.Elements.Add(lblv4);
                // Parte de baixo
                lblv21 = new ceTe.DynamicPDF.PageElements.Label(this.debitosCodigo[intDebitos - 1], 15 + intMargem, (Linha * 19) + 165 + intLinha, 400, 35);
                lblv22 = new ceTe.DynamicPDF.PageElements.Label(this.debitosVerba[intDebitos - 1], 70 + intMargem, (Linha * 19) + 165 + intLinha, 300, 35);
                lblv23 = new ceTe.DynamicPDF.PageElements.Label(this.debitosReferencia[intDebitos - 1], 715 + intMargem, (Linha * 19) + 165 + intLinha, 400, 35); // 730
                lblv24 = new ceTe.DynamicPDF.PageElements.Label(Formata_Valor(format_value(this.debitosValor[intDebitos - 1])), 980 + intMargem, (Linha * 19) + 165 + intLinha, 700, 35);
                // Ajustar o tamanho da fonte
                lblv21.FontSize = 12;
                lblv22.FontSize = 12;
                lblv23.FontSize = 12;
                lblv24.FontSize = 12;
                // Ajustar o tipo da fonte
                lblv23.Width = 40;
                lblv24.Width = 120;
                lblv23.Align = ceTe.DynamicPDF.TextAlign.Right;
                lblv24.Align = ceTe.DynamicPDF.TextAlign.Right;
                // Adicionar na página
                page.Elements.Add(lblv21);
                page.Elements.Add(lblv22);
                page.Elements.Add(lblv23);
                page.Elements.Add(lblv24);
                if (Linha == 33)
                {
                    acaoTotalizadores("Próxima Página");
                    Linha = 3;
                    CriarPaginaDoVerso();
                    booNovoCondominio = true;
                    strNovaFolha = "S";
                    acaoImprimirEtapa2();
                }
            }
            this.TotalItensDebitos = 0;
        }
        private void acaoImprimirEtapa2()
        {
            CriarNovaPagina();
            // Parte de cima
            ceTe.DynamicPDF.PageElements.Label lbl1 = new ceTe.DynamicPDF.PageElements.Label(" ", 700, 35, 400, 50);
            foreach (int indexChecked in checkedListBoxMSGTipoPagamento.CheckedIndices)
            {
                // The indexChecked variable contains the index of the item.
                lbl1 = new ceTe.DynamicPDF.PageElements.Label(checkedListBoxMSGTipoPagamento.Items[indexChecked].ToString(), 700, 25, 500, 50);
            }
            ceTe.DynamicPDF.PageElements.Label lbl2 = new ceTe.DynamicPDF.PageElements.Label(T_Unicode(cab1), 10 + intMargem, 25, 600, 50); // Nome do condomínio
            ceTe.DynamicPDF.PageElements.Label lbl3 = new ceTe.DynamicPDF.PageElements.Label(cab2, 10 + intMargem, 50, 750, 50);
            ceTe.DynamicPDF.PageElements.Label lbl4 = new ceTe.DynamicPDF.PageElements.Label(cab3, 10 + intMargem, 90, 750, 50);
            ceTe.DynamicPDF.PageElements.Label lbl5 = new ceTe.DynamicPDF.PageElements.Label(T_Unicode(cab4) + "/" + cab5, 1000 + intMargem, 90, 750, 50);
            lbl1.FontSize = 27;
            lbl2.FontSize = 26;
            lbl3.FontSize = 22;
            lbl4.FontSize = 22;
            lbl5.FontSize = 22;
            page.Elements.Add(lbl1);
            page.Elements.Add(lbl2);
            page.Elements.Add(lbl3);
            page.Elements.Add(lbl4);
            page.Elements.Add(lbl5);
            // Parte de baixo
            ceTe.DynamicPDF.PageElements.Label lbl21 = new ceTe.DynamicPDF.PageElements.Label(" ", 850, 35, 400, 50);

            foreach (int indexChecked in checkedListBoxMSGTipoPagamento.CheckedIndices)
            {
                // The indexChecked variable contains the index of the item.
                lbl21 = new ceTe.DynamicPDF.PageElements.Label(checkedListBoxMSGTipoPagamento.Items[indexChecked].ToString(), 700, 20 + intLinha, 500, 50);
            }

            lbl21.FontSize = 27;
            page.Elements.Add(lbl21);
            ceTe.DynamicPDF.PageElements.Label lbl22 = new ceTe.DynamicPDF.PageElements.Label(T_Unicode(cab1), 10 + intMargem, 20 + intLinha, 600, 50); // Nome do condomínio
            ceTe.DynamicPDF.PageElements.Label lbl23 = new ceTe.DynamicPDF.PageElements.Label(cab2, 10 + intMargem, 45 + intLinha, 750, 50);
            ceTe.DynamicPDF.PageElements.Label lbl24 = new ceTe.DynamicPDF.PageElements.Label(cab3, 10 + intMargem, 85 + intLinha, 750, 50);
            ceTe.DynamicPDF.PageElements.Label lbl25 = new ceTe.DynamicPDF.PageElements.Label(T_Unicode(cab4) + "/" + cab5, 1000 + intMargem, 85 + intLinha, 750, 50);
            lbl22.FontSize = 26;
            lbl23.FontSize = 22;
            lbl24.FontSize = 22;
            lbl25.FontSize = 22;
            page.Elements.Add(lbl21);
            page.Elements.Add(lbl22);
            page.Elements.Add(lbl23);
            page.Elements.Add(lbl24);
            page.Elements.Add(lbl25);
            // Parte de cima
            ceTe.DynamicPDF.PageElements.Label lbl6 = new ceTe.DynamicPDF.PageElements.Label(func1, 10 + intMargem, 135, 600, 35);
            ceTe.DynamicPDF.PageElements.Label lbl7 = new ceTe.DynamicPDF.PageElements.Label(func2, 720 + intMargem, 135, 300, 35);
            ceTe.DynamicPDF.PageElements.Label lbl8 = new ceTe.DynamicPDF.PageElements.Label(func3, 10 + intMargem, 155, 300, 35);
            ceTe.DynamicPDF.PageElements.Label lbl9 = new ceTe.DynamicPDF.PageElements.Label(func4, 720 + intMargem, 155, 300, 35);
            ceTe.DynamicPDF.PageElements.Label lbl10 = new ceTe.DynamicPDF.PageElements.Label(func5, 10 + intMargem, 175, 300, 35);
            ceTe.DynamicPDF.PageElements.Label lbl11 = new ceTe.DynamicPDF.PageElements.Label(func6, 720 + intMargem, 175, 300, 35);
            lbl6.FontSize = 22;
            lbl7.FontSize = 22;
            lbl8.FontSize = 22;
            lbl9.FontSize = 22;
            lbl10.FontSize = 22;
            lbl11.FontSize = 22;
            page.Elements.Add(lbl6);
            page.Elements.Add(lbl7);
            page.Elements.Add(lbl8);
            page.Elements.Add(lbl9);
            page.Elements.Add(lbl10);
            page.Elements.Add(lbl11);
            // Parte de baixo
            ceTe.DynamicPDF.PageElements.Label lbl26 = new ceTe.DynamicPDF.PageElements.Label(func1, 10 + intMargem, 135 + intLinha, 600, 35);
            ceTe.DynamicPDF.PageElements.Label lbl27 = new ceTe.DynamicPDF.PageElements.Label(func2, 720 + intMargem, 135 + intLinha, 300, 35);
            ceTe.DynamicPDF.PageElements.Label lbl28 = new ceTe.DynamicPDF.PageElements.Label(func3, 10 + intMargem, 155 + intLinha, 300, 35);
            ceTe.DynamicPDF.PageElements.Label lbl29 = new ceTe.DynamicPDF.PageElements.Label(func4, 720 + intMargem, 155 + intLinha, 300, 35);
            ceTe.DynamicPDF.PageElements.Label lbl210 = new ceTe.DynamicPDF.PageElements.Label(func5, 10 + intMargem, 175 + intLinha, 300, 35);
            ceTe.DynamicPDF.PageElements.Label lbl211 = new ceTe.DynamicPDF.PageElements.Label(func6, 720 + intMargem, 175 + intLinha, 300, 35);
            lbl26.FontSize = 22;
            lbl27.FontSize = 22;
            lbl28.FontSize = 22;
            lbl29.FontSize = 22;
            lbl210.FontSize = 22;
            lbl211.FontSize = 22;
            page.Elements.Add(lbl26);
            page.Elements.Add(lbl27);
            page.Elements.Add(lbl28);
            page.Elements.Add(lbl29);
            page.Elements.Add(lbl210);
            page.Elements.Add(lbl211);
            // Cargo Parte de cima
            ceTe.DynamicPDF.PageElements.Label lblCargo = new ceTe.DynamicPDF.PageElements.Label("Cargo: " + strCargo, 30 + intMargem, 815, 800, 35);
            lblCargo.FontSize = 14;
            page.Elements.Add(lblCargo);
            // Cargo Parte de baixo
            ceTe.DynamicPDF.PageElements.Label lblCargo2 = new ceTe.DynamicPDF.PageElements.Label("Cargo: " + strCargo, 30 + intMargem, 815 + intLinha, 800, 35);
            lblCargo.FontSize = 14;
            page.Elements.Add(lblCargo2);
            // Dados bancários
            ceTe.DynamicPDF.PageElements.Label lblnroBanco1 = new ceTe.DynamicPDF.PageElements.Label("Banco: " + nroBanco + " Nome: " + nomBanco + " Agência: " + ageBanco + " Conta: " + ccBanco, 30 + intMargem, 835, 800, 35);
            ceTe.DynamicPDF.PageElements.Label lblnroBanco2 = new ceTe.DynamicPDF.PageElements.Label("Banco: " + nroBanco + " Nome: " + nomBanco + " Agência: " + ageBanco + " Conta: " + ccBanco, 30 + intMargem, 835 + intLinha, 800, 35);
            ceTe.DynamicPDF.PageElements.Label lblcpf1 = new ceTe.DynamicPDF.PageElements.Label("CPF: " + cpfCli, 30 + intMargem, 850, 800, 35);
            ceTe.DynamicPDF.PageElements.Label lblcpf2 = new ceTe.DynamicPDF.PageElements.Label("CPF: " + cpfCli, 30 + intMargem, 850 + intLinha, 800, 35);
            lblnroBanco1.FontSize = 14;
            lblnroBanco2.FontSize = 14;
            lblcpf1.FontSize = 14;
            lblcpf2.FontSize = 14;
            page.Elements.Add(lblnroBanco1);
            page.Elements.Add(lblnroBanco2);
            page.Elements.Add(lblcpf1);
            page.Elements.Add(lblcpf2);
        }
        private void acaoNovoFuncionario()
        {
            acaoTotalizadores("");
            Linha = 3;
            CriarPaginaDoVerso();
            intTipoDoRegistro = 1;
            booNovoCondominio = true;
            Limpa_Bases_de_Calculo();
            strNovaFolha = "N";
            strUltimoHolerite = "N";
        }

        private void acaoTotalizadores(String strProxima)
        {
            if (strProxima != "Próxima Página")
            {
                // Parte de cima
                ceTe.DynamicPDF.PageElements.Label lbltotvenc = new ceTe.DynamicPDF.PageElements.Label(Formata_Valor(format_value(Convert.ToString(venc)).Trim()), 805 + intMargem, 820, 800, 35);
                ceTe.DynamicPDF.PageElements.Label lbltotdesc = new ceTe.DynamicPDF.PageElements.Label(Formata_Valor(format_value(Convert.ToString(desc)).Trim()), 967 + intMargem, 820, 800, 35);
                ceTe.DynamicPDF.PageElements.Label lblmsg1 = new ceTe.DynamicPDF.PageElements.Label("Valor Líquido", 830 + intMargem, 840, 800, 35);
                ceTe.DynamicPDF.PageElements.Label lbltotliq = new ceTe.DynamicPDF.PageElements.Label(Formata_Valor(format_value(Convert.ToString(vlr_liquido)).Trim()), 965 + intMargem, 845, 800, 35);

                lbltotvenc.Font = ceTe.DynamicPDF.Font.CourierBold;
                lbltotdesc.Font = ceTe.DynamicPDF.Font.CourierBold;
                lbltotliq.Font = ceTe.DynamicPDF.Font.CourierBold;
                lblmsg1.FontSize = 20;

                ceTe.DynamicPDF.PageElements.Label lblb1 = new ceTe.DynamicPDF.PageElements.Label("Salário Base", 35 + intMargem, 870, 800, 35);
                ceTe.DynamicPDF.PageElements.Label lblb2 = new ceTe.DynamicPDF.PageElements.Label("Sal. Contr. INSS", 180 + intMargem, 870, 800, 35);
                ceTe.DynamicPDF.PageElements.Label lblb3 = new ceTe.DynamicPDF.PageElements.Label("Base Cálc. FGTS", 380 + intMargem, 870, 800, 35);
                ceTe.DynamicPDF.PageElements.Label lblb4 = new ceTe.DynamicPDF.PageElements.Label("F.G.T.S. do Mês", 580 + intMargem, 870, 800, 35);
                ceTe.DynamicPDF.PageElements.Label lblb5 = new ceTe.DynamicPDF.PageElements.Label("Base de Cálc. IRRF", 735 + intMargem, 870, 800, 35);
                ceTe.DynamicPDF.PageElements.Label lblb6 = new ceTe.DynamicPDF.PageElements.Label("Faixa IRRF", 900 + intMargem, 870, 800, 35);
                ceTe.DynamicPDF.PageElements.Label lblb7 = new ceTe.DynamicPDF.PageElements.Label("No Dep IRRF/Sal Fam", 1010 + intMargem, 870, 800, 35);

                // Valores 
                ceTe.DynamicPDF.PageElements.Label lblv1 = new ceTe.DynamicPDF.PageElements.Label(Formata_Valor(format_value(Convert.ToString(Convert.ToDouble(strSalario_Base) / 100))), intMargem + 10, 890, 800, 35);
                ceTe.DynamicPDF.PageElements.Label lblv2 = new ceTe.DynamicPDF.PageElements.Label(Formata_Valor(format_value(Convert.ToString(Convert.ToDouble(strBASE_INSS_ATE_O_TETO) / 100))), 180 + intMargem, 885, 800, 35);
                ceTe.DynamicPDF.PageElements.Label lblv3 = new ceTe.DynamicPDF.PageElements.Label(Formata_Valor(format_value(Convert.ToString(Convert.ToDouble(strBASE_CALCULO_FGTS) / 100))), 370 + intMargem, 885, 800, 35);
                ceTe.DynamicPDF.PageElements.Label lblv4 = new ceTe.DynamicPDF.PageElements.Label(Formata_Valor(format_value(Convert.ToString(Convert.ToDouble(strVALOR_FGTS) / 100))), 580 + intMargem, 885, 800, 35);
                ceTe.DynamicPDF.PageElements.Label lblv5 = new ceTe.DynamicPDF.PageElements.Label(Formata_Valor(format_value(Convert.ToString(Convert.ToDouble(strBASE_IRRF) / 100))), 765 + intMargem, 885, 800, 35);
                ceTe.DynamicPDF.PageElements.Label lblv6 = new ceTe.DynamicPDF.PageElements.Label(Formata_Valor(Convert.ToString(Convert.ToDouble(strDEPENDENTE_DE_IRRF) / 100)), 1030 + intMargem, 885, 800, 35);
                ceTe.DynamicPDF.PageElements.Label lblv7 = new ceTe.DynamicPDF.PageElements.Label(Formata_Valor(format_value(Convert.ToString(Convert.ToDouble(strSALARIO_FAMILIA) / 100))), 1060 + intMargem, 885, 800, 35);
                // Desativado, Número de Dependentesndo IRRF ceTe.DynamicPDF.PageElements.Label lblv6 = new ceTe.DynamicPDF.PageElements.Label(format_value(line.Substring(77,15)), 900, 645, 700, 35);
                // Aumentar a fonte
                lblb1.FontSize = 14;
                lblb2.FontSize = 14;
                lblb3.FontSize = 14;
                lblb4.FontSize = 14;
                lblb5.FontSize = 14;
                lblb6.FontSize = 14;
                lblb7.FontSize = 14;

                lblv1.FontSize = 14;
                lblv2.FontSize = 14;
                lblv3.FontSize = 14;
                lblv4.FontSize = 14;
                lblv5.FontSize = 14;
                lblv6.FontSize = 14;
                lblv7.FontSize = 14;

                // Totais Vencimento, Descontos e Liquido
                lbltotvenc.FontSize = 20;
                lbltotdesc.FontSize = 20;
                lbltotliq.FontSize = 20;
                lbltotvenc.Width = 120;
                lbltotdesc.Width = 150;
                lbltotliq.Width = 150;
                lbltotvenc.Align = ceTe.DynamicPDF.TextAlign.Right;
                lbltotdesc.Align = ceTe.DynamicPDF.TextAlign.Right;
                lbltotliq.Align = ceTe.DynamicPDF.TextAlign.Right;

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
                page.Elements.Add(lblb7);

                page.Elements.Add(lblv1);
                page.Elements.Add(lblv2);
                page.Elements.Add(lblv3);
                page.Elements.Add(lblv4);
                page.Elements.Add(lblv5);
                page.Elements.Add(lblv6);
                //page.Elements.Add(lblv7);

                // Parte de baixo

                ceTe.DynamicPDF.PageElements.Label lbltotvenc2 = new ceTe.DynamicPDF.PageElements.Label(Formata_Valor(format_value(Convert.ToString(venc)).Trim()), 805 + intMargem, 820 + intLinha, 800, 35);
                ceTe.DynamicPDF.PageElements.Label lbltotdesc2 = new ceTe.DynamicPDF.PageElements.Label(Formata_Valor(format_value(Convert.ToString(desc)).Trim()), 967 + intMargem, 820 + intLinha, 800, 35);
                ceTe.DynamicPDF.PageElements.Label lblmsg21 = new ceTe.DynamicPDF.PageElements.Label("Valor Líquido", 830 + intMargem, 840 + intLinha, 800, 35);
                ceTe.DynamicPDF.PageElements.Label lbltotliq2 = new ceTe.DynamicPDF.PageElements.Label(Formata_Valor(format_value(Convert.ToString(vlr_liquido)).Trim()), 965 + intMargem, 850 + intLinha, 800, 35);

                lbltotvenc2.Font = ceTe.DynamicPDF.Font.CourierBold;
                lbltotdesc2.Font = ceTe.DynamicPDF.Font.CourierBold;
                lbltotliq2.Font = ceTe.DynamicPDF.Font.CourierBold;
                lblmsg21.FontSize = 20;

                ceTe.DynamicPDF.PageElements.Label lblb21 = new ceTe.DynamicPDF.PageElements.Label("Salário Base", 35 + intMargem, 870 + intLinha, 800, 35);
                ceTe.DynamicPDF.PageElements.Label lblb22 = new ceTe.DynamicPDF.PageElements.Label("Sal. Contr. INSS", 180 + intMargem, 870 + intLinha, 800, 35);
                ceTe.DynamicPDF.PageElements.Label lblb23 = new ceTe.DynamicPDF.PageElements.Label("Base Cálc. FGTS", 380 + intMargem, 870 + intLinha, 800, 35);
                ceTe.DynamicPDF.PageElements.Label lblb24 = new ceTe.DynamicPDF.PageElements.Label("F.G.T.S. do Mês", 580 + intMargem, 870 + intLinha, 800, 35);
                ceTe.DynamicPDF.PageElements.Label lblb25 = new ceTe.DynamicPDF.PageElements.Label("Base de Cálc. IRRF", 735 + intMargem, 870 + intLinha, 800, 35);
                ceTe.DynamicPDF.PageElements.Label lblb26 = new ceTe.DynamicPDF.PageElements.Label("Faixa IRRF", 900 + intMargem, 870 + intLinha, 800, 35);
                ceTe.DynamicPDF.PageElements.Label lblb27 = new ceTe.DynamicPDF.PageElements.Label("No Dep IRRF/Sal Fam", 1010 + intMargem, 870 + intLinha, 800, 35);

                // Valores 

                ceTe.DynamicPDF.PageElements.Label lblv21 = new ceTe.DynamicPDF.PageElements.Label(Formata_Valor(format_value(Convert.ToString(Convert.ToDouble(strSalario_Base) / 100))), intMargem + 10, 890 + intLinha, 800, 35);
                ceTe.DynamicPDF.PageElements.Label lblv22 = new ceTe.DynamicPDF.PageElements.Label(Formata_Valor(format_value(Convert.ToString(Convert.ToDouble(strBASE_INSS_ATE_O_TETO) / 100))), 180 + intMargem, 885 + intLinha, 800, 35);
                ceTe.DynamicPDF.PageElements.Label lblv23 = new ceTe.DynamicPDF.PageElements.Label(Formata_Valor(format_value(Convert.ToString(Convert.ToDouble(strBASE_CALCULO_FGTS) / 100))), 370 + intMargem, 885 + intLinha, 800, 35);
                ceTe.DynamicPDF.PageElements.Label lblv24 = new ceTe.DynamicPDF.PageElements.Label(Formata_Valor(format_value(Convert.ToString(Convert.ToDouble(strVALOR_FGTS) / 100))), 580 + intMargem, 885 + intLinha, 800, 35);
                ceTe.DynamicPDF.PageElements.Label lblv25 = new ceTe.DynamicPDF.PageElements.Label(Formata_Valor(format_value(Convert.ToString(Convert.ToDouble(strBASE_IRRF) / 100))), 765 + intMargem, 885 + intLinha, 800, 35);
                ceTe.DynamicPDF.PageElements.Label lblv26 = new ceTe.DynamicPDF.PageElements.Label(Formata_Valor(Convert.ToString(Convert.ToDouble(strDEPENDENTE_DE_IRRF) / 100)), 1030 + intMargem, 885 + intLinha, 800, 35);
                ceTe.DynamicPDF.PageElements.Label lblv27 = new ceTe.DynamicPDF.PageElements.Label(Formata_Valor(format_value(Convert.ToString(Convert.ToDouble(strSALARIO_FAMILIA) / 100))), 1060 + intMargem, 885 + intLinha, 800, 35);

                // Aumentar a fonte
                lblb21.FontSize = 14;
                lblb22.FontSize = 14;
                lblb23.FontSize = 14;
                lblb24.FontSize = 14;
                lblb25.FontSize = 14;
                lblb26.FontSize = 14;
                lblb27.FontSize = 14;

                lblv21.FontSize = 14;
                lblv22.FontSize = 14;
                lblv23.FontSize = 14;
                lblv24.FontSize = 14;
                lblv25.FontSize = 14;
                lblv26.FontSize = 14;
                lblv27.FontSize = 14;

                // Totais Vencimento, Descontos e Liquido
                lbltotvenc2.FontSize = 20;
                lbltotdesc2.FontSize = 20;
                lbltotliq2.FontSize = 20;
                lbltotvenc2.Width = 120;
                lbltotdesc2.Width = 150;
                lbltotliq2.Width = 150;
                lbltotvenc2.Align = ceTe.DynamicPDF.TextAlign.Right;
                lbltotdesc2.Align = ceTe.DynamicPDF.TextAlign.Right;
                lbltotliq2.Align = ceTe.DynamicPDF.TextAlign.Right;

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
                page.Elements.Add(lblb27);

                page.Elements.Add(lblv21);
                page.Elements.Add(lblv22);
                page.Elements.Add(lblv23);
                page.Elements.Add(lblv24);
                page.Elements.Add(lblv25);
                page.Elements.Add(lblv26);
                //page.Elements.Add(lblv27);
                vlr_liquido = 0;
            }

            // Add page to document
            document.Pages.Add(page);
        }

        private void btnCriar_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileBrowserDialog1 = new OpenFileDialog();
            if (fileBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                MessageBox.Show("Arquivo selecionado: " + fileBrowserDialog1.FileName);
            }
            else
            {
                MessageBox.Show("ERRO");
            }
            CriarPDF(fileBrowserDialog1.FileName);
        }

        private class CheckedItensColletion
        {
        }
    }
}