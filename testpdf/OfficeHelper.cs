using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Text;

namespace testpdf
{
    public static class OfficeHelper
    {
        /// <summary>
        /// 画列数据
        /// </summary>
        /// <param name="workSheet"></param>
        /// <param name="title"></param>
        /// <param name="rowIndex">第几行</param>
        /// <param name="columnIndex">第几列</param>
        public static bool DrawTitleLine(this ExcelWorksheet workSheet, List<string> titleList, int rowIndex, int columnIndex = 1)
        {
            if (titleList == null || titleList.Count == 0) return false;

            for (int i = 0; i < titleList.Count; i++)
                workSheet.Cells[rowIndex, columnIndex + i].Value = titleList[i];

            return true;
        }
    }
}
