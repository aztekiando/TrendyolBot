using Newtonsoft.Json.Linq;
using RestSharp;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TrendyolBot
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            CheckForIllegalCrossThreadCalls = false;
        }

        private async  void button1_Click(object sender, EventArgs e)
        {
            await Task.Run(async () =>
            {
                int sayi = 0;
                for (int i = 1; i <= Convert.ToInt32(textBox2.Text); i++)
                {
                    try
                    {
                        RestClient restClient = new RestClient("https://public.trendyol.com/discovery-web-searchgw-service/v2/api/infinite-scroll/" + textBox1.Text + "?pi=" + i + "&categoryRelevancyEnabled=false&isLegalRequirementConfirmed=false&searchStrategyType=DEFAULT&productStampType=TypeA");

                        RestRequest restRequest = new RestRequest("", Method.GET);


                        IRestResponse restResponse = restClient.Execute(restRequest);
                        JObject MyParser = JObject.Parse(restResponse.Content);
                        for (int j = 0; j < 24; j++)
                        {
                            try
                            {
                                string isim = (string)MyParser["result"]["products"][j]["name"];
                                string fiyat = (string)MyParser["result"]["products"][j]["price"]["sellingPrice"];
                                string satici = (string)MyParser["result"]["products"][j]["brand"]["name"];
                                string resim = "";
                                string kategori = (string)MyParser["result"]["products"][j]["categoryHierarchy"];
                                bool kargo = false;
                                string url = "https://www.trendyol.com" + (string)MyParser["result"]["products"][j]["url"];
                                if ((string)MyParser["result"]["products"][j]["freeCargo"] == "True")
                                {
                                    kargo = true;
                                }

                                var resimler = MyParser["result"]["products"][j]["images"].Values<string>().ToArray();
                                foreach (var item in resimler)
                                {
                                    resim += "https://cdn.dsmcdn.com/" + item + "}";
                                }



                                var id = (string)MyParser["result"]["products"][j]["id"];

                                RestClient cli = new RestClient("https://public.trendyol.com/discovery-web-productgw-service/api/product-detail/" + id + "/html-content");
                                RestRequest req = new RestRequest("", Method.GET);
                                req.AddHeader("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8");


                                IRestResponse restResponsee = cli.Execute(req);
                                JObject MyParser2 = JObject.Parse(restResponsee.Content);
                                string conn = "";
                                if ((string)MyParser2["statusCode"] == "200")
                                {
                                    conn = (string)MyParser2["result"];

                                    conn = conn.Replace(@"id=""rich-content-wrapper""", "");
                                }


                                string aciklama = conn;
                                if (checkBox1.Checked)
                                {
                                    aciklama = HtmlToText(conn);

                                }

                                sayi++;

                                dataGridView1.Rows.Add(sayi, isim, fiyat, satici, aciklama, resim, kategori, kargo.ToString(), url);
                            }
                            catch (Exception ex)
                            {
                                break;
                            }
                        }

                    }
                    catch (Exception)
                    {
                        
                        break;

                    }


                }
                MessageBox.Show("Tüm Ürünler Çekildi");

            });

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.ColumnCount = 9;
            dataGridView1.Columns[0].Name = "Sayı";
            dataGridView1.Columns[1].Name = "Ürün İsmi";
            dataGridView1.Columns[2].Name = "Ürün Fiyatı";
            dataGridView1.Columns[3].Name = "Ürün Satıcısı";
            dataGridView1.Columns[4].Name = "Ürün Açıklaması";
            dataGridView1.Columns[5].Name = "Ürün Fotoğrafları";
            dataGridView1.Columns[6].Name = "Kategori";
            dataGridView1.Columns[7].Name = "Ücretsiz Kargo";
            dataGridView1.Columns[8].Name = "Ürün Url";

        }
        public static void ExcelD(DataGridView dataGridView1)
        {

            try
            {

                DialogResult dialog = new DialogResult();
                dialog = MessageBox.Show("Bu işlem, veri yoğunluğuna göre uzun sürebilir. Devam etmek istiyor musunuz?", "EXCEL'E AKTARMA", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (dialog == DialogResult.Yes)
                {
                    Microsoft.Office.Interop.Excel.Application uyg = new Microsoft.Office.Interop.Excel.Application();
                    uyg.Visible = true;
                    Microsoft.Office.Interop.Excel.Workbook kitap = uyg.Workbooks.Add(System.Reflection.Missing.Value);
                    Microsoft.Office.Interop.Excel.Worksheet sheet1 = (Microsoft.Office.Interop.Excel.Worksheet)kitap.Sheets[1];
                    for (int i = 0; i < dataGridView1.Columns.Count; i++)
                    {
                        Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[1, i + 1];
                        myRange.Value2 = dataGridView1.Columns[i].HeaderText;
                    }

                    for (int i = 0; i < dataGridView1.Columns.Count; i++)
                    {
                        for (int j = 0; j < dataGridView1.Rows.Count; j++)
                        {
                            Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[j + 2, i + 1];
                            myRange.Value2 = dataGridView1[i, j].Value;
                        }
                    }
                }
                else
                {
                    MessageBox.Show("İŞLEM İPTAL EDİLDİ.", "İşlem Sonucu", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception)
            {

                MessageBox.Show("İŞLEM TAMAMLANMADAN EXCEL PENCERESİNİ KAPATTINIZ.", "HATA", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExcelD(dataGridView1);
        }


        private static string HtmlToText(string html)
        {
            const string tagWhiteSpace = @"(>|$)(\W|\n|\r)+<";//matches one or more (white space or line breaks) between '>' and '<'
            const string stripFormatting = @"<[^>]*(>|$)";//match any character between '<' and '>', even when end tag is missing
            const string lineBreak = @"<(br|BR)\s{0,1}\/{0,1}>";//matches: <br>,<br/>,<br />,<BR>,<BR/>,<BR />
            var lineBreakRegex = new Regex(lineBreak, RegexOptions.Multiline);
            var stripFormattingRegex = new Regex(stripFormatting, RegexOptions.Multiline);
            var tagWhiteSpaceRegex = new Regex(tagWhiteSpace, RegexOptions.Multiline);

            var text = html;
            //Decode html specific characters
            text = System.Net.WebUtility.HtmlDecode(text);
            //Remove tag whitespace/line breaks
            text = tagWhiteSpaceRegex.Replace(text, "><");
            //Replace <br /> with line breaks
            text = lineBreakRegex.Replace(text, Environment.NewLine);
            //Strip formatting
            text = stripFormattingRegex.Replace(text, string.Empty);

            return text;
        }

    }
}
