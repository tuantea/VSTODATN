using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Excel = Microsoft.Office.Interop.Excel;           

namespace DATN.FuntionExcel
{
    internal class ImportJson
    {
        public static void ImportJsonToExcel(Excel.Worksheet worksheet)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Json (*.json)|*.json|All files (*.*)|*.*";
            openFileDialog.Title = "Open File";
            if (openFileDialog.ShowDialog() != DialogResult.OK)
            {
                return;
            }
            string jsonContent = File.ReadAllText(openFileDialog.FileName);
            dynamic jsonData = JsonConvert.DeserializeObject(jsonContent);

            bool isVisible = (bool)jsonData["config"]["visible"];
            bool shouldPrintNow = (bool)jsonData["config"]["printnow"];
            bool shouldActivateSheet = (bool)jsonData["config"]["activatesheet"];
            bool shouldTerminate = (bool)jsonData["config"]["terminate"];

            JArray documentsArray = (JArray)jsonData["documents"];
            JObject firstDocument = (JObject)documentsArray.First;

            JArray contentControlsArray = (JArray)firstDocument["contentcontrols"];
            foreach (JObject contentControl in contentControlsArray)
            {
                string title = (string)contentControl["title"];
                string value = (string)contentControl["value"];

                // Do something with the title and value, such as populating them in Excel cells
            }
        }

    }
}
