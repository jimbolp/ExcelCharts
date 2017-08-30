using System;
using System.Runtime.InteropServices;
using OfficeOpenXml;
using System.Collections.Generic;
using OfficeOpenXml.Drawing.Chart;

namespace ExcelCharts
{
    class ChartRange
    {
        const double chartHeigth = 521.0134; //18.23cm * 28.58
        const double chartWidth = 867.9746; //30.37cm * 28.58
        string topDateCell;
        string topDataCell;
        string bottomDateCell;
        string bottomDataCell;
        ExcelAddressBase usedRange;
        ExcelRange DateRange;
        ExcelRange DataRange;
        readonly bool printNeeded;
        readonly char col;
        readonly int type = 0;
        public int ChartNumber { get; set; } = 0;
        public int RowOfRange { get; set; } = 0;
        public ChartRange(char type,ExcelWorksheet xlWs, ExcelAddressBase usedRange, bool print, bool special)
        {
            string graphType = type.ToString().ToLower();
            if (!(graphType == "t" || graphType == "h"))
            {
                throw new ArgumentException("The given graph type is not supported! The supported graph types are 'T' for Temperatures or 'H' for Humidity.");
            }
            switch (graphType)
            {
                case "t":
                    col = 'B';
                    break;
                case "h":
                    {
                        col = special ? 'B' : 'C';
                    }
                    break;
            }
            //col = type.ToString().ToLower() == "t" ? 'B' : 'C';
            this.type = graphType == "t" ? 1 : 2;
            topDateCell = "A" + 1;
            topDataCell = col.ToString() + 1;
            bottomDateCell = topDateCell;
            bottomDataCell = topDataCell;
            this.usedRange = usedRange;
            printNeeded = print;
            DateRange = xlWs.Cells[topDateCell + ":" + bottomDateCell]; //.Range[topDateCell, bottomDateCell];
            DataRange = xlWs.Cells[topDataCell + ":" + bottomDataCell];
        }
        public void ExpandRange(int row)
        {
            RowOfRange++;
            bottomDateCell = "A" + row;
            bottomDataCell = col.ToString() + row;
        }
        public void StartNewRange(int row)
        {
            RowOfRange = 1;
            topDateCell = "A" + row;
            topDataCell = col.ToString() + row;
            bottomDateCell = "A" + row;
            bottomDataCell = col.ToString() + row;
        }
        public void CreateChart(ExcelWorksheet ws, List<ExcelChart> xlChartObjs, string Name, double startChartPositionLeft, double startChartPositionTop)
        {
            ChartNumber++;
            if (type == 2) { startChartPositionLeft += 100; } else { startChartPositionTop += 50; } 
            Name = type != 0 ? (type == 1 ? Name + "_T" : Name + "_H") : Name;
            DateRange = ws.Cells[topDateCell + ":" + bottomDateCell];
            DataRange = ws.Cells[topDataCell + ":" + bottomDataCell];

            ExcelChart xlChartObj = ws.Drawings.AddChart(Name+ChartNumber, eChartType.Line);//, startChartPositionLeft, startChartPositionTop, chartWidth, chartHeigth);

            var serie = xlChartObj.Series.Add(DataRange.Address,DateRange.Address);
            xlChartObj.Title.Text = Name;

            xlChartObj.SetSize((int)chartWidth, (int)chartHeigth);
            xlChartObj.SetPosition((int)startChartPositionTop, (int)startChartPositionLeft);

            var temp = DateRange.Value;
            //xlChartObjs.Add(xlChartObj);
            //Still no idea how to print the Charts?!
            //if (printNeeded)
               //PrintCharts();

            DateRange = null;
            DataRange = null;
        }//*/
        public bool EnoughDataForChart()
        {
            if (topDateCell == bottomDateCell || topDataCell == bottomDataCell)
            {
                return false;
            }
            return true;
        }
    }
}
