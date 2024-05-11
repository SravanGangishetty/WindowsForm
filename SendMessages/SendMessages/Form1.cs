using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using IronXL;
using Syncfusion.WinForms.Core.Utils;


namespace SendMessages
{
    public partial class Form1 : Form
    {

        BusyIndicator busyIndicator = new BusyIndicator();
        public Form1()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }


        /// <summary>
        /// Browse
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            //Browse

            
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                
                string fileExt = Path.GetExtension(openFileDialog1.FileName); //get the file extension
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {
                    textBox1.Text = openFileDialog1.FileName;
                }
                else
                {
                    MessageBox.Show("Please choose .xls or .xlsx file only.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error
                    textBox1.Text = "";
                    openFileDialog1.Reset();
                }

            }

        }

        /// <summary>
            /// Preview
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            busyIndicator.Show(this);
            if (openFileDialog1.FileName == "" || openFileDialog1.FileName == "openFileDialog1")
            {
                busyIndicator.Hide();
                MessageBox.Show("Please select .xls or .xlsx file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error
            }
            else
            {
                try
                {
                    DataTable dtExcel = ReadExcel(openFileDialog1.FileName); //read excel file
                    dataGrdView.Visible = true;
                    dataGrdView.DataSource = dtExcel;
                }
                catch (Exception ex)
                {
                    busyIndicator.Hide();
                    MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            busyIndicator.Hide();
        }

        private DataTable ReadExcel(string fileName)
        {
            WorkBook workbook = WorkBook.Load(fileName);

            //// Work with a single WorkSheet.
            ////you can pass static sheet name like Sheet1 to get that sheet
            ////WorkSheet sheet = workbook.GetWorkSheet("Sheet1");

            //You can also use workbook.DefaultWorkSheet to get default in case you want to get first sheet only
            WorkSheet sheet = workbook.DefaultWorkSheet;

            //Convert the worksheet to System.Data.DataTable
            //Boolean parameter sets the first row as column names of your table.
            return sheet.ToDataTable(true);
        }


        /// <summary>
        /// Clear
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {

            dataGrdView.Visible = false;
            dataGrdView.DataSource = new DataTable();
            textBox1.Text = "";
            openFileDialog1.Reset();
            textBox2.Text = "";
        }


        /// <summary>
        /// Send Messages
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {

            busyIndicator.Show(this);

            if (textBox2.Text == "")
            {
                busyIndicator.Hide();
                MessageBox.Show("Please enter CURL URL", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error
                return;
            }

            if (openFileDialog1.FileName == "" || openFileDialog1.FileName == "openFileDialog1")
            {
                busyIndicator.Hide();
                MessageBox.Show("Please select .xls or .xlsx file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error
                return;
            }

            if (!IsValidCURL(textBox2.Text))
            {
                busyIndicator.Hide();
                MessageBox.Show("Please enter valid CURL", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error
                return;
            }



            try
            {

                DataTable dtExcel = ReadExcel(openFileDialog1.FileName); //read excel file

                if (dtExcel == null)
                {
                    busyIndicator.Hide();
                    MessageBox.Show("Please select valid .xls or .xlsx file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error
                    return;
                }
                else if (dtExcel != null && (dtExcel.Rows.Count < 1 && dtExcel.Columns.Count < 1))
                {
                    busyIndicator.Hide();
                    MessageBox.Show("Please upload valid .xls or .xlsx file. File should contain headers and atleast one column with Mobile Number", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error
                    return;
                }


                var curl = textBox2.Text;

                var splitObject = curl.Split(new char[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);

                var url = "";
                var headersDict = new Dictionary<string, string>();
                var payLoad = "";


                if (splitObject != null && splitObject.Any())
                {
                    foreach (var val in splitObject)
                    {

                        if (val.Contains("--location"))
                        {
                            Match match = Regex.Match(val, @"'([^']*)");
                            if (match.Success)
                            {
                                string yourValue = match.Groups[1].Value;

                                if (yourValue.Contains(""))
                                {
                                    yourValue = yourValue.Replace("<YOUR-API-DOMAIN>", "apisocial.telebu.com"); //Replacing domain 
                                }

                                url = yourValue;
                            }

                            if (url == "")
                            {
                                MessageBox.Show("Please enter valid CURL URL", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error  
                                return;
                            }
                        }

                        else if (val.Contains("--header"))
                        {
                            Match match = Regex.Match(val, @"'([^']*)");
                            if (match.Success)
                            {
                                string yourValue = match.Groups[1].Value;

                                if (!string.IsNullOrEmpty(yourValue) && yourValue.Contains(":") && yourValue.Contains("Basic"))
                                {
                                    var headersValues = yourValue.Split(new char[] { ':' }, StringSplitOptions.RemoveEmptyEntries);

                                    if (headersDict.ContainsKey(headersValues[0]))
                                    {
                                        headersDict[headersValues[0]] = headersValues[1];

                                    }
                                    else
                                    {
                                        headersDict.Add(headersValues[0], headersValues[1]);
                                    }
                                }
                            }

                            if (headersDict == null || headersDict.Count == 0)
                            {
                                busyIndicator.Hide();
                                MessageBox.Show("Please enter valid CURL URL", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error  
                                return;
                            }
                        }
                        else if (val.Contains("--data"))
                        {
                            Match match = Regex.Match(val, @"'([^']*)");
                            if (match.Success)
                            {
                                string yourValue = match.Groups[1].Value;
                                payLoad = yourValue;
                            }

                            if (payLoad == "" || !payLoad.Contains("<provide receiver phone number>"))
                            {
                                busyIndicator.Hide();
                                MessageBox.Show("Please enter valid CURL URL", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error  
                                return;
                            }
                        }
                    }
                }

                SendMessages(url, headersDict, payLoad, dtExcel);

                dataGrdView.Visible = false;
                dataGrdView.DataSource = new DataTable();
                textBox1.Text = "";
                openFileDialog1.Reset();
                textBox2.Text = "";

                busyIndicator.Hide();

                MessageBox.Show("Messages sent, please check the status in the response file", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information); //custom messageBox to show error  
            }
            catch (Exception ex)
            {
                busyIndicator.Hide();
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            busyIndicator.Hide();
        }

        private bool IsValidCURL(string curlURL)
        {
            if (curlURL.Contains("--location") && curlURL.Contains("--header") && curlURL.Contains("--data"))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private void SendMessages(string url, Dictionary<string, string> headers, string payLoadStr, DataTable excelDataTable)
        {

            //excelDataTable.Columns.Add(new DataColumn("Status", typeof(string)));
            excelDataTable.Columns.Add(new DataColumn("ResponseMessage", typeof(string)));

            Parallel.ForEach(excelDataTable.Rows.Cast<DataRow>(), dataRow =>
            {
                var payLoad = payLoadStr;

                if (payLoad.Contains("<provide receiver phone number>"))
                {
                    payLoad = payLoad.Replace("<provide receiver phone number>", dataRow[0].ToString());
                }

                for (var i = 1; i <= 100; i++)
                {
                    if (payLoad.Contains("<provide value for parameter " + i.ToString() + ">"))
                    {
                        payLoad = payLoad.Replace("<provide value for parameter " + i.ToString() + ">", dataRow[i].ToString());

                        //dataRow["ResponseMessage"] = payLoad;
                    }
                    else
                    {
                        break;
                    }
                }

                var result = Task.Run(() => PostAsync(url, headers, payLoad)).Result;
                dataRow["ResponseMessage"] = result;

            });


            //foreach (DataRow dataRow in excelDataTable.Rows)
            //{
            //    var payLoad = payLoadStr;

            //    if (payLoad.Contains("<provide receiver phone number>"))
            //    {
            //        payLoad = payLoad.Replace("<provide receiver phone number>", dataRow[0].ToString());
            //    }

            //    for (var i = 1; i <= 100; i++)
            //    {
            //        if (payLoad.Contains("<provide value for parameter " + i.ToString() + ">"))
            //        {
            //            payLoad = payLoad.Replace("<provide value for parameter " + i.ToString() + ">", dataRow[i].ToString());

            //            dataRow["ResponseMessage"] = payLoad;
            //        }
            //        else
            //        {
            //            break;
            //        }
            //    }



            //}

            //for (int j = 1; j < excelDataTable.Rows.Count; j++)
            //{
            //    if (excelDataTable.Rows[j][0] != null && excelDataTable.Rows[j][0].ToString() == "")
            //    {
            //        break;
            //    }

            //    if (payLoad.Contains("<provide receiver phone number>"))
            //    {
            //        payLoad = payLoad.Replace("<provide receiver phone number>", excelDataTable.Rows[j][0].ToString());
            //    }

            //    for (var i = 1; i <= 100; i++)
            //    {
            //        if (payLoad.Contains("<provide value for parameter " + i.ToString() + ">"))
            //        {
            //            payLoad = payLoad.Replace("<provide value for parameter " + i.ToString() + ">", excelDataTable.Rows[j][i].ToString());

            //            excelDataTable.Rows[j]["ResponseMessage"] = payLoad;
            //        }
            //        else
            //        {
            //            break;
            //        }
            //    }
            //}


            string extension = Path.GetExtension(openFileDialog1.FileName);
            var downloadFileName = openFileDialog1.FileName.Replace(extension, "_Response_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "" + extension);

            Excel.ExcelUtlity obj = new Excel.ExcelUtlity();

            obj.WriteDataTableToExcel(excelDataTable, "Response", downloadFileName, "Response");
           

        }


        private async Task<String> PostAsync(string url, Dictionary<string, string> headersDict, string payLoad)
        {
            using (HttpClient httpClient = new HttpClient())
            {

                HttpResponseMessage responseMessage = null;

                try
                {
                    Uri serviceUri = new Uri(url);
                    StringContent httpContent = new StringContent(payLoad, Encoding.UTF8, "application/json");

                    if (headersDict != null)
                    {
                        foreach (var keyValuePair in headersDict)
                        {
                            httpClient.DefaultRequestHeaders.Add(keyValuePair.Key, keyValuePair.Value);
                        }
                    }

                    responseMessage = await httpClient.PostAsync(serviceUri, httpContent);


                }
                catch (Exception ex)
                {
                    if (responseMessage == null)
                    {
                        responseMessage = new HttpResponseMessage();
                    }
                    responseMessage.StatusCode = HttpStatusCode.InternalServerError;
                    responseMessage.ReasonPhrase = string.Format("RestHttpClient.SendRequest failed: {0}", ex);
                }

                string responseString = await responseMessage.Content.ReadAsStringAsync();

                return responseString;
            }
        }

        private void ExportToExcel(DataSet table, string filePath)
        {

            int tablecount = table.Tables.Count;
            StreamWriter sw = new StreamWriter(filePath, false);
            sw.Write(@"<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN"">");
            sw.Write("<font style='font-size:10.0pt; font-family:Calibri;'>");

            for (int k = 0; k < tablecount; k++)
            {


                sw.Write("<BR><BR><BR>");
                sw.Write("<Table border='1' bgColor='#ffffff' borderColor='#000000' cellSpacing='0' cellPadding='0' style='font-size:10.0pt; font-family:Calibri; background:'#1E90FF'> <TR>");


                int columnscount = table.Tables[k].Columns.Count;

                for (int j = 0; j < columnscount; j++)
                {
                    sw.Write("<Td bgColor='#87CEFA'>");
                    sw.Write("");
                    //sw.Write(table.Columns[j].ToString());
                    sw.Write(table.Tables[k].Columns[j].ToString());

                    sw.Write("");
                    sw.Write("</Td>");
                }
                sw.Write("</TR>");
                foreach (DataRow row in table.Tables[k].Rows)
                {
                    sw.Write("<TR>");
                    for (int i = 0; i < table.Tables[k].Columns.Count; i++)
                    {
                        sw.Write("<Td>");
                        sw.Write(row[i].ToString());
                        sw.Write("</Td>");
                    }
                    sw.Write("</TR>");
                }
                sw.Write("</Table>");
                //sw.Write("<BR><BR><BR><BR>");
                //sw.Write("\n");
                //sw.Write(string.Format("Line1{0}Line2{0}", Environment.NewLine));


                sw.Write("</font>");

            }
            sw.Close();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            busyIndicator.Show(this);

            if (textBox2.Text == "")
            {
                busyIndicator.Hide();
                MessageBox.Show("Please enter CURL URL", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error
                return;
            }

            if (!IsValidCURL(textBox2.Text))
            {
                busyIndicator.Hide();
                MessageBox.Show("Please enter valid CURL", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error
                return;
            }

            var curl = textBox2.Text;

            var splitObject = curl.Split(new char[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);

            var payLoad = "";


            if (splitObject != null && splitObject.Any())
            {
                foreach (var val in splitObject)
                {

                    if (val.Contains("--data"))
                    {
                        Match match = Regex.Match(val, @"'([^']*)");
                        if (match.Success)
                        {
                            string yourValue = match.Groups[1].Value;
                            payLoad = yourValue;
                        }

                        if (payLoad == "" || !payLoad.Contains("<provide receiver phone number>"))
                        {
                            busyIndicator.Hide();
                            MessageBox.Show("Please enter valid CURL URL", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error  
                            return;
                        }
                    }
                }
            }

            int count = 0;

            for (var i = 1; i <= 100; i++)
            {
                if (payLoad.Contains("<provide value for parameter " + i.ToString() + ">"))
                {
                    count = i;
                }
                else
                {
                    break;
                }
            }


            busyIndicator.Hide();

            MessageBox.Show("Number of parameters is "+ count.ToString() +" ", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information); //custom messageBox to show error  

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }
    }
}
