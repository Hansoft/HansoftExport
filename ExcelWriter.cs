using System;
using System.Collections.Generic;
using System.Xml;
using System.Xml.Schema;
using OfficeOpenXml;

namespace Hansoft.SdkUtils
{
    public class ExcelWriter
    {

        public class Cell
        {
            private string value;

            public Cell(string value)
            {
                this.value = value;
            }

            public string Value
            { 
                get { return value; }
            }
        }

        public class Row
        {
            private List<Cell> cells;

            public Row()
            {
                cells = new List<Cell>();
            }

            public List<Cell> Cells
            {
                get { return cells; }
            }

            public Cell AddCell(string value)
            {
                Cell cell = new Cell(value);
                cells.Add(cell);
                return cell;
            }
        }


        private List<Row> rows;

        public ExcelWriter()
        {
            rows = new List<Row>();
        }
        
        public Row AddRow()
        {
            Row row = new Row();
            rows.Add(row);
            return row;
        }

        public void SaveAsOfficeOpenXml(string fileName)
        {
            System.IO.FileInfo fileInfo = new System.IO.FileInfo(fileName);
            if (fileInfo.Exists)
                fileInfo.Delete();
            ExcelPackage excelPackage = new ExcelPackage(fileInfo);
            ExcelWorksheet sheet1 = excelPackage.Workbook.Worksheets.Add("Sheet1");
            for (int rowInd= 0; rowInd <rows.Count; rowInd++)
                for (int colInd= 0; colInd<rows[rowInd].Cells.Count; colInd++)
                    sheet1.Cells[rowInd+1, colInd+1].Value = rows[rowInd].Cells[colInd].Value;
            excelPackage.Save();
        }
        
    }
}
