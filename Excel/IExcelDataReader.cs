using System.Data;

namespace Excel
{
    internal interface IExcelDataReader
    {
        DataSet AsDataSet();
        void Close();
    }
}