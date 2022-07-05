using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excelMultipleSheetDeneme
{
    public partial class ExcelControl
    {
        public DataSet dataSet;
        public string[] dataTableNames = new string[] { "Mode1", "Mode2", "Mode3", "Mode4", "Mode5", "Mode6" };

        public DataTable Sheet1_DataTable;
        public DataTable Sheet2_DataTable;
        public DataTable Sheet3_DataTable;
        public DataTable Sheet4_DataTable;
        public DataTable Sheet5_DataTable;
        public DataTable Sheet6_DataTable;

        public string firstColumn = "Num";
        public string secondColumn = "EVP ID";
        public string thirdColumn = "Test Name";
        public string fourthColumn = "Lower Limit";
        public string fifthColumn = "Upper Limit";
        public string sixthColumn = "Measured Value";
        public string seventhColumn = "Unit";
        public string eighthColumn = "PassOrFail";

        public int rowCounterExcel = 1;

        public void CreateExcelDataSet()
        {
            // Creating a new DataTable.
            Sheet1_DataTable = new DataTable(dataTableNames[0]);
            Sheet2_DataTable = new DataTable(dataTableNames[1]);
            Sheet3_DataTable = new DataTable(dataTableNames[2]);
            Sheet4_DataTable = new DataTable(dataTableNames[3]);
            Sheet5_DataTable = new DataTable(dataTableNames[4]);
            Sheet6_DataTable = new DataTable(dataTableNames[5]);

            CreateNewColumnDoubleTypeSheet1(firstColumn, "Num", false, false, true);
            CreateNewColumnStringTypeSheet1(secondColumn, "EVP ID", false, false, true);
            CreateNewColumnStringTypeSheet1(thirdColumn, "Test Name", false, false, false);
            CreateNewColumnDoubleTypeSheet1(fourthColumn, "Lower Limit", false, false, false);
            CreateNewColumnDoubleTypeSheet1(fifthColumn, "Upper Limit", false, false, false);
            CreateNewColumnDoubleTypeSheet1(sixthColumn, "Measured Value", false, false, false);
            CreateNewColumnStringTypeSheet1(seventhColumn, "Unit", false, false, false);
            CreateNewColumnStringTypeSheet1(eighthColumn, "PassOrFail", false, false, false);

            CreateNewColumnDoubleTypeSheet2(firstColumn, "Num", false, false, true);
            CreateNewColumnStringTypeSheet2(secondColumn, "EVP ID", false, false, true);
            CreateNewColumnStringTypeSheet2(thirdColumn, "Test Name", false, false, false);
            CreateNewColumnDoubleTypeSheet2(fourthColumn, "Lower Limit", false, false, false);
            CreateNewColumnDoubleTypeSheet2(fifthColumn, "Upper Limit", false, false, false);
            CreateNewColumnDoubleTypeSheet2(sixthColumn, "Measured Value", false, false, false);
            CreateNewColumnStringTypeSheet2(seventhColumn, "Unit", false, false, false);
            CreateNewColumnStringTypeSheet2(eighthColumn, "PassOrFail", false, false, false);

            CreateNewColumnDoubleTypeSheet3(firstColumn, "Num", false, false, true);
            CreateNewColumnStringTypeSheet3(secondColumn, "EVP ID", false, false, true);
            CreateNewColumnStringTypeSheet3(thirdColumn, "Test Name", false, false, false);
            CreateNewColumnDoubleTypeSheet3(fourthColumn, "Lower Limit", false, false, false);
            CreateNewColumnDoubleTypeSheet3(fifthColumn, "Upper Limit", false, false, false);
            CreateNewColumnDoubleTypeSheet3(sixthColumn, "Measured Value", false, false, false);
            CreateNewColumnStringTypeSheet3(seventhColumn, "Unit", false, false, false);
            CreateNewColumnStringTypeSheet3(eighthColumn, "PassOrFail", false, false, false);

            CreateNewColumnDoubleTypeSheet4(firstColumn, "Num", false, false, true);
            CreateNewColumnStringTypeSheet4(secondColumn, "EVP ID", false, false, true);
            CreateNewColumnStringTypeSheet4(thirdColumn, "Test Name", false, false, false);
            CreateNewColumnDoubleTypeSheet4(fourthColumn, "Lower Limit", false, false, false);
            CreateNewColumnDoubleTypeSheet4(fifthColumn, "Upper Limit", false, false, false);
            CreateNewColumnDoubleTypeSheet4(sixthColumn, "Measured Value", false, false, false);
            CreateNewColumnStringTypeSheet4(seventhColumn, "Unit", false, false, false);
            CreateNewColumnStringTypeSheet4(eighthColumn, "PassOrFail", false, false, false);

            CreateNewColumnDoubleTypeSheet5(firstColumn, "Num", false, false, true);
            CreateNewColumnStringTypeSheet5(secondColumn, "EVP ID", false, false, true);
            CreateNewColumnStringTypeSheet5(thirdColumn, "Test Name", false, false, false);
            CreateNewColumnDoubleTypeSheet5(fourthColumn, "Lower Limit", false, false, false);
            CreateNewColumnDoubleTypeSheet5(fifthColumn, "Upper Limit", false, false, false);
            CreateNewColumnDoubleTypeSheet5(sixthColumn, "Measured Value", false, false, false);
            CreateNewColumnStringTypeSheet5(seventhColumn, "Unit", false, false, false);
            CreateNewColumnStringTypeSheet5(eighthColumn, "PassOrFail", false, false, false);

            CreateNewColumnDoubleTypeSheet6(firstColumn, "Num", false, false, true);
            CreateNewColumnStringTypeSheet6(secondColumn, "EVP ID", false, false, true);
            CreateNewColumnStringTypeSheet6(thirdColumn, "Test Name", false, false, false);
            CreateNewColumnDoubleTypeSheet6(fourthColumn, "Lower Limit", false, false, false);
            CreateNewColumnDoubleTypeSheet6(fifthColumn, "Upper Limit", false, false, false);
            CreateNewColumnDoubleTypeSheet6(sixthColumn, "Measured Value", false, false, false);
            CreateNewColumnStringTypeSheet6(seventhColumn, "Unit", false, false, false);
            CreateNewColumnStringTypeSheet6(eighthColumn, "PassOrFail", false, false, false);

            // Make id column the primary key column.
            DataColumn[] PrimaryKeyColumns = new DataColumn[1];
            PrimaryKeyColumns[0] = Sheet1_DataTable.Columns["Num"];
            Sheet1_DataTable.PrimaryKey = PrimaryKeyColumns;

            // Create a new DataSet
            dataSet = new DataSet();

            // Add custTable to the DataSet.
            dataSet.Tables.Add(Sheet1_DataTable);
            dataSet.Tables.Add(Sheet2_DataTable);
            dataSet.Tables.Add(Sheet3_DataTable);
            dataSet.Tables.Add(Sheet4_DataTable);
            dataSet.Tables.Add(Sheet5_DataTable);
            dataSet.Tables.Add(Sheet6_DataTable);

        }

        public void CreateNewColumnStringTypeSheet1(string columnName, string captionText, bool autoIncrement, bool readOnly, bool Unique)
        {
            // Create a column
            DataColumn dataColumn = new DataColumn();
            dataColumn.DataType = typeof(string);
            dataColumn.ColumnName = columnName;
            dataColumn.Caption = captionText;
            dataColumn.AutoIncrement = autoIncrement;
            dataColumn.ReadOnly = readOnly;
            dataColumn.Unique = Unique;
            Sheet1_DataTable.Columns.Add(dataColumn);
        }

        public void CreateNewColumnDoubleTypeSheet1(string columnName, string captionText, bool autoIncrement, bool readOnly, bool Unique)
        {
            // Create a column
            DataColumn dataColumn = new DataColumn();
            dataColumn.DataType = typeof(double);
            dataColumn.ColumnName = columnName;
            dataColumn.Caption = captionText;
            dataColumn.AutoIncrement = autoIncrement;
            dataColumn.ReadOnly = readOnly;
            dataColumn.Unique = Unique;
            Sheet1_DataTable.Columns.Add(dataColumn);
        }

        public void CreateNewColumnStringTypeSheet2(string columnName, string captionText, bool autoIncrement, bool readOnly, bool Unique)
        {
            // Create a column
            DataColumn dataColumn = new DataColumn();
            dataColumn.DataType = typeof(string);
            dataColumn.ColumnName = columnName;
            dataColumn.Caption = captionText;
            dataColumn.AutoIncrement = autoIncrement;
            dataColumn.ReadOnly = readOnly;
            dataColumn.Unique = Unique;
            Sheet2_DataTable.Columns.Add(dataColumn);
        }

        public void CreateNewColumnDoubleTypeSheet2(string columnName, string captionText, bool autoIncrement, bool readOnly, bool Unique)
        {
            // Create a column
            DataColumn dataColumn = new DataColumn();
            dataColumn.DataType = typeof(double);
            dataColumn.ColumnName = columnName;
            dataColumn.Caption = captionText;
            dataColumn.AutoIncrement = autoIncrement;
            dataColumn.ReadOnly = readOnly;
            dataColumn.Unique = Unique;
            Sheet2_DataTable.Columns.Add(dataColumn);
        }

        public void CreateNewColumnStringTypeSheet3(string columnName, string captionText, bool autoIncrement, bool readOnly, bool Unique)
        {
            // Create a column
            DataColumn dataColumn = new DataColumn();
            dataColumn.DataType = typeof(string);
            dataColumn.ColumnName = columnName;
            dataColumn.Caption = captionText;
            dataColumn.AutoIncrement = autoIncrement;
            dataColumn.ReadOnly = readOnly;
            dataColumn.Unique = Unique;
            Sheet3_DataTable.Columns.Add(dataColumn);
        }

        public void CreateNewColumnDoubleTypeSheet3(string columnName, string captionText, bool autoIncrement, bool readOnly, bool Unique)
        {
            // Create a column
            DataColumn dataColumn = new DataColumn();
            dataColumn.DataType = typeof(double);
            dataColumn.ColumnName = columnName;
            dataColumn.Caption = captionText;
            dataColumn.AutoIncrement = autoIncrement;
            dataColumn.ReadOnly = readOnly;
            dataColumn.Unique = Unique;
            Sheet3_DataTable.Columns.Add(dataColumn);
        }

        public void CreateNewColumnStringTypeSheet4(string columnName, string captionText, bool autoIncrement, bool readOnly, bool Unique)
        {
            // Create a column
            DataColumn dataColumn = new DataColumn();
            dataColumn.DataType = typeof(string);
            dataColumn.ColumnName = columnName;
            dataColumn.Caption = captionText;
            dataColumn.AutoIncrement = autoIncrement;
            dataColumn.ReadOnly = readOnly;
            dataColumn.Unique = Unique;
            Sheet4_DataTable.Columns.Add(dataColumn);
        }

        public void CreateNewColumnDoubleTypeSheet4(string columnName, string captionText, bool autoIncrement, bool readOnly, bool Unique)
        {
            // Create a column
            DataColumn dataColumn = new DataColumn();
            dataColumn.DataType = typeof(double);
            dataColumn.ColumnName = columnName;
            dataColumn.Caption = captionText;
            dataColumn.AutoIncrement = autoIncrement;
            dataColumn.ReadOnly = readOnly;
            dataColumn.Unique = Unique;
            Sheet4_DataTable.Columns.Add(dataColumn);
        }

        public void CreateNewColumnStringTypeSheet5(string columnName, string captionText, bool autoIncrement, bool readOnly, bool Unique)
        {
            // Create a column
            DataColumn dataColumn = new DataColumn();
            dataColumn.DataType = typeof(string);
            dataColumn.ColumnName = columnName;
            dataColumn.Caption = captionText;
            dataColumn.AutoIncrement = autoIncrement;
            dataColumn.ReadOnly = readOnly;
            dataColumn.Unique = Unique;
            Sheet5_DataTable.Columns.Add(dataColumn);
        }

        public void CreateNewColumnDoubleTypeSheet5(string columnName, string captionText, bool autoIncrement, bool readOnly, bool Unique)
        {
            // Create a column
            DataColumn dataColumn = new DataColumn();
            dataColumn.DataType = typeof(double);
            dataColumn.ColumnName = columnName;
            dataColumn.Caption = captionText;
            dataColumn.AutoIncrement = autoIncrement;
            dataColumn.ReadOnly = readOnly;
            dataColumn.Unique = Unique;
            Sheet5_DataTable.Columns.Add(dataColumn);
        }

        public void CreateNewColumnStringTypeSheet6(string columnName, string captionText, bool autoIncrement, bool readOnly, bool Unique)
        {
            // Create a column
            DataColumn dataColumn = new DataColumn();
            dataColumn.DataType = typeof(string);
            dataColumn.ColumnName = columnName;
            dataColumn.Caption = captionText;
            dataColumn.AutoIncrement = autoIncrement;
            dataColumn.ReadOnly = readOnly;
            dataColumn.Unique = Unique;
            Sheet6_DataTable.Columns.Add(dataColumn);
        }

        public void CreateNewColumnDoubleTypeSheet6(string columnName, string captionText, bool autoIncrement, bool readOnly, bool Unique)
        {
            // Create a column
            DataColumn dataColumn = new DataColumn();
            dataColumn.DataType = typeof(double);
            dataColumn.ColumnName = columnName;
            dataColumn.Caption = captionText;
            dataColumn.AutoIncrement = autoIncrement;
            dataColumn.ReadOnly = readOnly;
            dataColumn.Unique = Unique;
            Sheet6_DataTable.Columns.Add(dataColumn);
        }

        // row structure => double string string double double double string string
        public void CreateNewRowSheet1(double rowCount, string evpId, string testName, double lowerLimit, double upperLimit, double measuredValue, string unit, string passFail)
        {
            DataRow myDataRow = Sheet1_DataTable.NewRow();
            myDataRow[firstColumn] = rowCount;
            myDataRow[secondColumn] = evpId;
            myDataRow[thirdColumn] = testName;
            myDataRow[fourthColumn] = lowerLimit;
            myDataRow[fifthColumn] = upperLimit;
            myDataRow[sixthColumn] = measuredValue;
            myDataRow[seventhColumn] = unit;
            myDataRow[eighthColumn] = passFail;
            Sheet1_DataTable.Rows.Add(myDataRow);
        }

        public void CreateNewRowSheet2(double rowCount, string evpId, string testName, double lowerLimit, double upperLimit, double measuredValue, string unit, string passFail)
        {
            DataRow myDataRow = Sheet2_DataTable.NewRow();
            myDataRow[firstColumn] = rowCount;
            myDataRow[secondColumn] = evpId;
            myDataRow[thirdColumn] = testName;
            myDataRow[fourthColumn] = lowerLimit;
            myDataRow[fifthColumn] = upperLimit;
            myDataRow[sixthColumn] = measuredValue;
            myDataRow[seventhColumn] = unit;
            myDataRow[eighthColumn] = passFail;
            Sheet2_DataTable.Rows.Add(myDataRow);
        }

        public void CreateNewRowSheet3(double rowCount, string evpId, string testName, double lowerLimit, double upperLimit, double measuredValue, string unit, string passFail)
        {
            DataRow myDataRow = Sheet3_DataTable.NewRow();
            myDataRow[firstColumn] = rowCount;
            myDataRow[secondColumn] = evpId;
            myDataRow[thirdColumn] = testName;
            myDataRow[fourthColumn] = lowerLimit;
            myDataRow[fifthColumn] = upperLimit;
            myDataRow[sixthColumn] = measuredValue;
            myDataRow[seventhColumn] = unit;
            myDataRow[eighthColumn] = passFail;
            Sheet3_DataTable.Rows.Add(myDataRow);
        }

        public void CreateNewRowSheet4(double rowCount, string evpId, string testName, double lowerLimit, double upperLimit, double measuredValue, string unit, string passFail)
        {
            DataRow myDataRow = Sheet4_DataTable.NewRow();
            myDataRow[firstColumn] = rowCount;
            myDataRow[secondColumn] = evpId;
            myDataRow[thirdColumn] = testName;
            myDataRow[fourthColumn] = lowerLimit;
            myDataRow[fifthColumn] = upperLimit;
            myDataRow[sixthColumn] = measuredValue;
            myDataRow[seventhColumn] = unit;
            myDataRow[eighthColumn] = passFail;
            Sheet4_DataTable.Rows.Add(myDataRow);
        }

        public void CreateNewRowSheet5(double rowCount, string evpId, string testName, double lowerLimit, double upperLimit, double measuredValue, string unit, string passFail)
        {
            DataRow myDataRow = Sheet5_DataTable.NewRow();
            myDataRow[firstColumn] = rowCount;
            myDataRow[secondColumn] = evpId;
            myDataRow[thirdColumn] = testName;
            myDataRow[fourthColumn] = lowerLimit;
            myDataRow[fifthColumn] = upperLimit;
            myDataRow[sixthColumn] = measuredValue;
            myDataRow[seventhColumn] = unit;
            myDataRow[eighthColumn] = passFail;
            Sheet5_DataTable.Rows.Add(myDataRow);
        }

        public void CreateNewRowSheet6(double rowCount, string evpId, string testName, double lowerLimit, double upperLimit, double measuredValue, string unit, string passFail)
        {
            DataRow myDataRow = Sheet6_DataTable.NewRow();
            myDataRow[firstColumn] = rowCount;
            myDataRow[secondColumn] = evpId;
            myDataRow[thirdColumn] = testName;
            myDataRow[fourthColumn] = lowerLimit;
            myDataRow[fifthColumn] = upperLimit;
            myDataRow[sixthColumn] = measuredValue;
            myDataRow[seventhColumn] = unit;
            myDataRow[eighthColumn] = passFail;
            Sheet6_DataTable.Rows.Add(myDataRow);
        }

    }
}
