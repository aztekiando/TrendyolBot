using Newtonsoft.Json.Linq;
using RestSharp;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
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
                                string resim = "";
                                string kategori = (string)MyParser["result"]["products"][j]["categoryHierarchy"];
                                bool kargo = false;
                                string url = "https://www.trendyol.com" + (string)MyParser["result"]["products"][j]["url"];
                                if ((string)MyParser["result"]["products"][j]["freeCargo"] == "True")
                                {
                                    kargo = true;
                                }

                                var resimler = MyParser["result"]["products"][0]["images"].Values<string>().ToArray();
                                foreach (var item in resimler)
                                {
                                    resim += "https://cdn.dsmcdn.com/" + item + ":";
                                }

                                sayi++;

                                dataGridView1.Rows.Add(sayi, isim, fiyat, resim, kategori, kargo.ToString(), url);
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
            dataGridView1.ColumnCount =7;
            dataGridView1.Columns[0].Name = "Sayı";
            dataGridView1.Columns[1].Name = "Ürün İsmi";
            dataGridView1.Columns[2].Name = "Ürün Fiyatı";
            dataGridView1.Columns[3].Name = "Ürün Fotoğrafları";
            dataGridView1.Columns[4].Name = "Kategori";
            dataGridView1.Columns[5].Name = "Ücretsiz Kargo";
            dataGridView1.Columns[6].Name = "Ürün Url";
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
    }
}
