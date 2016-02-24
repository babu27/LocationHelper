using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Net;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using LocationHelper.Google;
using LocationHelper.Helper;
using LocationHelper.Properties;
using Newtonsoft.Json;
using Excel = Microsoft.Office.Interop.Excel;

[assembly: SuppressIldasm]

namespace LocationHelper
{
    public partial class FrmLocationHelper : Form
    {
        public delegate void UpdateResult(string message);

        public delegate void UpdateProgress();

        private readonly Dictionary<string, string> _localTempCache = new Dictionary<string, string>();

        public FrmLocationHelper()
        {
            InitializeComponent();
            var showDate = new DateTime(2016, 3, 27);

            menuStrip1.Visible = DateTime.Now >= showDate;

        }

        private void btnInputFile_Click(object sender, EventArgs e)
        {
            ofDialog.ShowDialog();
        }

        private void ofDialog_FileOk(object sender, CancelEventArgs e)
        {
            txtFileName.Text = ofDialog.FileName;
        }

        private void btnProcess_Click(object sender, EventArgs e)
        {
            _localTempCache.Clear();

            txtResult.Text = string.Empty;

            if (string.IsNullOrEmpty(txtFileName.Text) || !File.Exists(txtFileName.Text))
            {
                MessageBox.Show(Resources.FrmLocationHelper_btnProcess_Click_Invalid_File_Path, Resources.FrmLocationHelper_btnProcess_Click_File_Path_Error);
                return;
            }

            new TaskFactory().StartNew(() => _ProcessData(txtFileName.Text));
        }

        private void _ProcessData(string filePath)
        {
            var xlApp = TryGetExistingExcelApplication() ?? new Excel.Application();
            var xlWorkbook = xlApp.Workbooks.Open(filePath, ReadOnly: false);

            if (xlWorkbook.Sheets.Count < ConfigHelper.DataSheetNumber)
            {
                xlWorkbook.Close();
                xlApp.Quit();
                MessageBox.Show(Resources.FrmLocationHelper__ProcessData_Configured_Sheet_number_does_not_exists_in_the_supplied_file);
                return;
            }

            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[ConfigHelper.DataSheetNumber];
            var xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            TryToggleProgressBarVisibility();
            TryInitProgressBar(rowCount - ConfigHelper.FirstDataRowNumber);

            for (int i = ConfigHelper.FirstDataRowNumber; i <= rowCount; i++)
            {
                try
                {
                    string lat = xlRange.Cells[i, ConfigHelper.LatitudeColumn].Value2.ToString();
                    string lng = xlRange.Cells[i, ConfigHelper.LongitudeColumn].Value2.ToString();

                    if (string.IsNullOrEmpty(lat) || string.IsNullOrEmpty(lng) || IsZero(lat) || IsZero(lng))
                    {
                        TryUpdateResultText(string.Format("RecordCount={0}, Invalid lat,lan={1} {2}", i - ConfigHelper.FirstDataRowNumber + 1, string.Format("{0},{1}", lat, lng), Environment.NewLine));
                        continue;
                    }

                    var latlng = string.Format("{0},{1}", lat, lng);

                    TryUpdateResultText(string.Format("RecordCount={0}, Processing latlan={1} {2}", i - ConfigHelper.FirstDataRowNumber + 1, latlng, Environment.NewLine));

                    var cacheKey = GetCacheKey(latlng);
                    string addressOutput;

                    if (_localTempCache.ContainsKey(cacheKey))
                    {
                        addressOutput = _localTempCache[cacheKey];
                    }
                    else
                    {
                        addressOutput = GetAddressFromCoOrdinate(latlng);
                        _localTempCache[cacheKey] = addressOutput;
                    }

                    xlRange.Cells[i, colCount + 1].Value2 = addressOutput;

                    TryUpdateProgressBar();
                    TryUpdateResultText(string.Format("Result={0}{1}", addressOutput, Environment.NewLine));
                }
                catch (Exception oex)
                {
                    TryUpdateResultText("There is an error while processing the location" + oex.Message + Environment.NewLine);
                    //MessageBox.Show(
                    //    Resources.FrmLocationHelper__ProcessData_Exception_occured_while_closing_the_excel_application_);
                }
            }

            TryUpdateResultText("Completed.....................");

            try
            {
                xlApp.DisplayAlerts = false;
                var outputFilePath = GetOutputFilePath(txtFileName.Text);

                xlRange = xlWorksheet.UsedRange;

                xlRange.Columns.AutoFit();

                if (File.Exists(outputFilePath))
                    File.Delete(outputFilePath);

                xlWorkbook.SaveAs(outputFilePath);
            }
            finally
            {
                xlRange.Clear();
                xlWorkbook.Close();
                xlApp.Quit();
                TryToggleProgressBarVisibility();
            }
        }

        private bool IsZero(string lat)
        {
            const double compVal = 0.0;

            var val = Convert.ToDouble(lat);

            return val <= compVal;
        }

        private string GetOutputFilePath(string inputFilePath)
        {
            var fileName = Path.GetFileNameWithoutExtension(inputFilePath);
            var fileExt = Path.GetExtension(inputFilePath);

            fileName = string.Format("{0}_output.{1}", fileName, fileExt);

            // ReSharper disable AssignNullToNotNullAttribute
            return Path.Combine(Path.GetDirectoryName(inputFilePath), fileName);
            // ReSharper restore AssignNullToNotNullAttribute
        }

        /// <summary>
        /// Expects comma seperated Langitude and latitude And Returns the Address string.
        /// </summary>
        /// <param name="langAndLat"></param>
        /// <returns></returns>
        private string GetAddressFromCoOrdinate(string langAndLat)
        {
            if (string.IsNullOrEmpty(langAndLat))
                return string.Empty;

            //var langAndLatAry = langAndLat.Split(",".ToCharArray());

            return CallGoogleApi(langAndLat);

        }

        private string CallGoogleApi(string latlng)
        {
            const string url = "http://maps.googleapis.com/maps/api/geocode/json?latlng={0}";

            var callingUrl = string.Format(url, latlng);

            var httpRequest = (HttpWebRequest)WebRequest.Create(callingUrl);
            httpRequest.Method = "GET";

            var response = (HttpWebResponse)httpRequest.GetResponse();
            if (response.StatusCode == HttpStatusCode.OK)
            {
                var resultString = GetResponseString(response);

                var result = JsonConvert.DeserializeObject<RootObject>(resultString);

                return result.results.Count >= 1 ? result.results[0].formatted_address : string.Empty;
            }

            return string.Empty;
        }

        public static string GetResponseString(HttpWebResponse response)
        {
            using (var resStream = response.GetResponseStream())
            {
                if (resStream != null)
                    using (var reader = new StreamReader(resStream))
                    {
                        var responseString = reader.ReadToEnd();
                        reader.Close();
                        resStream.Close();
                        return responseString;
                    }
            }

            return string.Empty;

        }

        private void UpdateResultText(object message)
        {
            txtResult.Text += message;
        }

        private void TryUpdateResultText(string message)
        {

            if (InvokeRequired)
            {
                Invoke(new UpdateResult(UpdateResultText), message);
            }
            else
                UpdateResultText(message);
        }

        private void tsMenuAbout_Click(object sender, EventArgs e)
        {
            var about = new FrmAbout { Location = Location };
            about.ShowDialog();
        }

        public Excel.Application TryGetExistingExcelApplication()
        {
            try
            {
                var o = Marshal.GetActiveObject("Excel.Application");
                return (Excel.Application)o;
            }
            catch (COMException)
            {
                // Probably there is no existing Excel instance running, return null
                return null;
            }
        }

        private void TryInitProgressBar(object count)
        {
            if (InvokeRequired)
                Invoke(new Action<object>(InitProgressBar), count);
            else
                InitProgressBar(count);
        }

        private void InitProgressBar(object count)
        {
            var val = int.Parse(count.ToString());
            progressBar1.Step = 1;
            progressBar1.Minimum = 0;
            progressBar1.Maximum = val > 0 ? val : 1;
            progressBar1.Value = 0;
        }

        private void TryUpdateProgressBar()
        {
            if (InvokeRequired)
                Invoke(new Action(UpdateProgressBar));
            else
                UpdateProgressBar();
        }

        private void UpdateProgressBar()
        {
            progressBar1.PerformStep();
            //progressBar1.Value += 1;
        }

        private void ToggleProgressBarVisiblity()
        {
            progressBar1.Visible = !progressBar1.Visible;
        }

        private void TryToggleProgressBarVisibility()
        {
            if (InvokeRequired)
                Invoke(new Action(ToggleProgressBarVisiblity));
            else
                ToggleProgressBarVisiblity();
        }

        private string GetCacheKey(string latlng)
        {
            return latlng.Trim();
        }
    }
}
