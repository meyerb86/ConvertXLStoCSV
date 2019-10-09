using System;
using System.Data;
using System.IO;
using System.Web;


public partial class Default2 : System.Web.UI.Page
{
    static void Main() { }
    DataSet result = new DataSet();
    string filePath = @"C:\Users\BEM26331\Documents\AppDevProjects\9.10.19.xls";

    protected void UploadButton_Click(object sender, EventArgs e)
    {
    }

    protected void Page_Load(object sender, EventArgs e)
    {
        FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);

        Excel.IExcelDataReader excelReader = Excel.ExcelReaderFactory.CreateOpenXmlReader(stream);
        DataSet result = excelReader.AsDataSet();
        excelReader.Close();

        result.Tables[0].TableName.ToString();

        string csvData = "";
        int row_no = 0;
        int ind = 0;

        while (row_no < result.Tables[ind].Rows.Count)
        {
            for (int i = 0; i < result.Tables[ind].Columns.Count; i++)
            {
                csvData += result.Tables[ind].Rows[row_no][i].ToString() + ",";
            }
            row_no++;
            csvData += "\n";
        }

        string output = @"C:\Users\BEM26331\Documents\AppDevProjects\test.csv";
        StreamWriter csv = new StreamWriter(@output, false);
        csv.Write(csvData);
        csv.Close();
    }
}
