using System;
using System.Collections.Generic;
using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel_Raqmana
{
    public static class SheetProcessor
    {
        public static void ProcessWorkbook(Excel.Workbook wb, List<string> errors, out bool processed)
        {
            processed = false;
            string localSep = CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;
            string invSep = (localSep == ".") ? "," : ".";

            foreach (Excel.Worksheet sheet in wb.Worksheets)
            {
                string header = Convert.ToString((sheet.Cells[8, 1] as Excel.Range).Text);
                if (string.IsNullOrEmpty(header) || !header.Contains("رقم")) continue;

                processed = true;
                int lastRow = sheet.Cells[sheet.Rows.Count, 2].End[Excel.XlDirection.xlUp].Row;
                int colEst = 0, colAdv = 0, startG = 5, endG = 0;

                for (int c = 5; c <= 25; c++)
                {
                    string h = Convert.ToString((sheet.Cells[8, c] as Excel.Range).Text);
                    if (h.Contains("التقديرات")) colEst = c;
                    if (h.Contains("الارشادات")) colAdv = c;
                    if (h.Contains("التقديرات") && endG == 0) endG = c - 1;
                }

                if (colEst == 0) continue;
                int subjectsCount = (endG - startG) + 1;

                for (int r = 9; r <= lastRow; r++)
                {
                    double sum = 0; bool isEmpty = true; bool hasError = false;

                    for (int c = startG; c <= endG; c++)
                    {
                        var cell = sheet.Cells[r, c] as Excel.Range;
                        string txt = Convert.ToString(cell.Value2);

                        if (string.IsNullOrWhiteSpace(txt)) { cell.Value2 = 0; txt = "0"; }
                        else { isEmpty = false; }

                        if (txt.Contains(invSep))
                        {
                            errors.Add($"ورقة: {sheet.Name} | خلية: {cell.get_Address(false, false)} | الخطأ: استخدم ({localSep})");
                            hasError = true; break;
                        }

                        if (!double.TryParse(txt, NumberStyles.Any, CultureInfo.CurrentCulture, out double v))
                        {
                            errors.Add($"ورقة: {sheet.Name} | خلية: {cell.get_Address(false, false)} | الخطأ: قيمة غير صالحة");
                            hasError = true; break;
                        }

                        // التحقق من صحة العلامة (0-20)
                        if (v < 0 || v > 20)
                        {
                            errors.Add($"ورقة: {sheet.Name} | خلية: {cell.get_Address(false, false)} | الخطأ: علامة خارج المجال ({v})");
                            hasError = true; break;
                        }

                        sum += v;
                    }

                    if (hasError) continue;
                    if (isEmpty)
                    {
                        sheet.Cells[r, colEst].Value2 = "غائب";
                        sheet.Cells[r, colAdv].Value2 = "غائب";
                    }
                    else
                    {
                        var res = EstimatesLogic.GetEvaluation(sum / subjectsCount);
                        sheet.Cells[r, colEst].Value2 = res.gradeText;
                        sheet.Cells[r, colAdv].Value2 = res.adviseText;
                    }
                }
            }
        }
    }
}