using CefSharp.Structs;
using Franklin_Templeton_DAL.Constants;
using Franklin_Templeton_DAL.Helpers;
using Franklin_Templeton_DAL.Models;
using Microsoft.Office.Interop.Excel;
using Microsoft.Web.WebView2.Core;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.AxHost;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Franklin_Templeton_DAL.InputForms
{
    public partial class WebWizardForm : Form
    {
        CoreWebView2DownloadOperation downloadOperation;
        public WebWizardForm()
        {
            InitializeComponent();
        }

        private async void WebWizardForm_Load(object sender, EventArgs e)
        {
            await InitializeAsync();
            if (DalSettings.dalEnvironment == ApiUrls.DalEnvironment.Production)
            {
                webView21.Source = new Uri("https://mosaiq.corp.frk.com/dal/webui/wizard?isFromDalPlugin=true");
            }
            if (DalSettings.dalEnvironment == ApiUrls.DalEnvironment.Development)
            {
                  webView21.Source = new Uri("https://mosaiqdev.corp.frk.com/dal/webui/wizard?isFromDalPlugin=true");
            }
            if (DalSettings.dalEnvironment == ApiUrls.DalEnvironment.Uat)
            {
                webView21.Source = new Uri("https://mosaiquat.corp.frk.com/dal/webui/wizard?isFromDalPlugin=true");
            }
            if (DalSettings.dalEnvironment == ApiUrls.DalEnvironment.Test)
            {
                webView21.Source = new Uri("https://mosaiqtest.corp.frk.com/dal/webui/wizard?isFromDalPlugin=true");
            }
        }
        public async Task InitializeAsync()
        {
            CoreWebView2EnvironmentOptions options = new CoreWebView2EnvironmentOptions("--isFromDalPlugin=true");
            options.AllowSingleSignOnUsingOSPrimaryAccount= true; 
            string tempwebcachedir = string.Empty;
            CoreWebView2Environment webview2environment = null;

            //set value

            if (string.IsNullOrEmpty(tempwebcachedir))
            {
                //get fully-qualified path to user's temp folder
                tempwebcachedir = System.IO.Path.GetTempPath();

                tempwebcachedir = System.IO.Path.Combine(tempwebcachedir, System.Guid.NewGuid().ToString("n"));
            }
            //use with webview2 fixedversionruntime
            webview2environment = await CoreWebView2Environment.CreateAsync(null, tempwebcachedir, options);

            //wait for corewebview2 initialization
            await webView21.EnsureCoreWebView2Async(webview2environment);
            webView21.CoreWebView2.Settings.UserAgent += ";dalplugin";
            webView21.CoreWebView2.Settings.IsScriptEnabled= true;
            webView21.CoreWebView2.WebMessageReceived += webView21_WebMessageReceived;
            webView21.CoreWebView2.DownloadStarting += webView21_DownloadStarting;
        }

        private void webView21_DownloadStarting(object sender, CoreWebView2DownloadStartingEventArgs e)
        {
            downloadOperation = e.DownloadOperation;
            e.ResultFilePath = System.IO.Path.GetTempPath() + Guid.NewGuid().ToString() + ".tmp";
            downloadOperation.StateChanged += downLoadStateChanged;
            e.Handled=true;
        }

        private void downLoadStateChanged(object sender, object e)
        {
            if (downloadOperation.State== CoreWebView2DownloadState.Completed)
            {
                ExcelHelper excelHelper = new ExcelHelper();
                excelHelper.AddBase64File(downloadOperation.ResultFilePath, ExcelHelper.ExcelFileType.FileAsPath);
                CloseMe();
               //((Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet).Range[1].Activate();
            }
            else if(downloadOperation.State==CoreWebView2DownloadState.Interrupted)
            {
                MessageBox.Show("Error while retrieving data");
                CloseMe();
            }
        }

        private void webView21_WebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {

            ExcelHelper excelHelper = new ExcelHelper();
            var strMessage = e.TryGetWebMessageAsString();
            WebMessage webMessage = null;
            var message = string.Empty;
            var calledfrom = string.Empty;
            if (strMessage.StartsWith("{") && strMessage.EndsWith("}"))
            {
                webMessage = JsonConvert.DeserializeObject<WebMessage>(strMessage);
                message = webMessage.Message;
                calledfrom = webMessage.CalledFrom;
            }
            else
            {
                message = strMessage;
            }
            if (message == "CancelButtonClicked")
            {
                CloseMe();
            }
            else if (message == "ExcelFileLoaded")
            {
                if (string.IsNullOrEmpty(webMessage.Data) == false)
                {
                    string fileUri = webMessage.Data;
                    if (!fileUri.StartsWith("http:") && !fileUri.StartsWith("https:") && !fileUri.StartsWith("file:"))
                    {
                        fileUri= "file://" + fileUri;
                    }
                    webView21.CoreWebView2.Navigate(fileUri);
                    //excelHelper.AddBase64File(webMessage.Data);
                }
            }
            else if (message == "SelectFromExcel")
            {               
                string Values = PopulateCellRange(excelHelper);
                if (!string.IsNullOrEmpty(Values))
                {
                    if (calledfrom == "dates")
                    {                        
                            string[] DateList = Values.Split(',');
                            string SelectedDate = "";
                            for (int i = 0; i < DateList.Length; i++)
                            {
                            DateTime dt;
                                if (DateTime.TryParse(DateList[i], out dt))
                                {
                                    //var date = Convert.ToDateTime(DateList[i]);
                                    var formateddate = dt.ToString("yyyy-MM-dd");
                                    if (i != DateList.Length - 1)
                                    {
                                        SelectedDate += formateddate + "|";
                                    }
                                    else
                                    {
                                        SelectedDate += formateddate;
                                    }
                                }
                                else
                                {
                                    if (i != DateList.Length - 1)
                                    {
                                        SelectedDate += DateList[i] + "|";
                                    }
                                    else
                                    {
                                        SelectedDate += DateList[i];
                                    }
                                }
                            }
                                webMessage = new WebMessage { Message = "SelectFromExcel", Data = SelectedDate, CalledFrom = webMessage.CalledFrom };
                                webView21.CoreWebView2.PostWebMessageAsJson(JsonConvert.SerializeObject(webMessage));
                                                 
                    }
                    else
                    {
                        Values = Values.Replace(',', '|');
                        webMessage = new WebMessage { Message = "SelectFromExcel", Data = Values, CalledFrom = webMessage.CalledFrom };
                        webView21.CoreWebView2.PostWebMessageAsJson(JsonConvert.SerializeObject(webMessage));
                    }
                    //webView21.CoreWebView2.PostWebMessageAsString(Values);
                    //await webView21.ExecuteScriptAsync("document.querySelector('#selectedexcelvalues').textContent = '" + Values + "';");
                }
            }
        }
        private string PopulateCellRange(ExcelHelper excelHelper)
        {

            string returnValue = string.Empty;
            try
            {
                this.Hide();
                returnValue = excelHelper.GetCellRange("Select cells containing entites");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                this.Show();
                this.Focus();
            }
            return returnValue;
        }
        public void CloseMe()
        {
            this.Close();
            Globals.ThisAddIn.Application.ActiveWindow.Activate();
        }

    }
}
