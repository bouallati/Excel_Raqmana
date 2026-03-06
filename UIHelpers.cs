using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel_Raqmana
{
    public static class UIHelpers
    {
        public static void ShowAboutBox()
        {
            string aboutText =
                            "🔹 Excel_Raqmana v1.0\n" +
                            "━━━━━━━━━━━━━━━━━━━━━\n" +
                            "تم تطوير هذا العمل لتسهيل مهام الأساتذة \n في إدراج التقديرات والإرشادات آلياً.\n\n" +
                            "━━━━━━━━━━━━━━━━━━━━━\n" +
                            "👤 المطور: Oussama Bouallati\n" +
                            "━━━━━━━━━━━━━━━━━━━━━\n" +
                            "🌿 عمل تطوعي مجاني - نسألكم الدعاء."+
                            "━━━━━━━━━━━━━━━━━━━━━\n" +
                            "https://github.com/bouallati/Excel-Raqmana-Addin";

            MessageBox.Show(aboutText, "حول البرنامج",
                MessageBoxButtons.OK, MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
        }

        public static void ShowFinalReport(List<string> errors, bool processed, Excel.Workbook wb)
        {
            if (!processed)
            {
                MessageBox.Show("لم يتم العثور على أي جدول متوافق مع نظام الرقمنة.", "تنبيه",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                return;
            }

            if (errors.Count > 0)
            {
                string report = $"تمت العملية مع اكتشاف {errors.Count} أخطاء:\n\n" + string.Join("\n", errors.GetRange(0, Math.Min(errors.Count, 5)));
                var res = MessageBox.Show(report + "\n\nهل تريد الذهاب لمكان أول خطأ؟", "تقرير التدقيق",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);

                if (res == DialogResult.Yes)
                {
                    try
                    {
                        string first = errors[0];
                        string sName = first.Split('|')[0].Replace("ورقة:", "").Trim();
                        string cAddr = first.Split('|')[1].Replace("خلية:", "").Trim();
                        ((Excel.Worksheet)wb.Sheets[sName]).Activate();
                        wb.Application.Range[cAddr].Select();
                    }
                    catch { }
                }
            }
            else
            {
                MessageBox.Show("تمت العملية بنجاح تام وبدون أخطاء!", "نجاح",
                    MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
            }
        }
    }
}