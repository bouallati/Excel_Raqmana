using System;

namespace Excel_Raqmana
{
    public static class EstimatesLogic
    {
        public static (string gradeText, string adviseText) GetEvaluation(double average)
        {
            if (average <= 0) return ("غائب", "غائب");
            if (average < 7) return ("نتائج غير مقبولة", "احذر التهاون");
            if (average < 10) return ("نتائج دون الوسط", "ينقصك الحرص و التركيز");
            if (average < 12) return ("نتائج متوسطة", "بمقدورك تحقيق نتائج افضل");
            if (average < 14) return ("نتائج حسنة", "نتائج مقبولة بامكانك تحسينه");
            if (average < 16) return ("نتائج جيدة", "واصل الاجتهاد و المثابرة");
            if (average < 18) return ("نتائج جيدة", "نتائج جيدة ومشجعة واصل");
            return ("نتائج ممتازة", "نتائج ممتازة ومرضية واصل");
        }
    }
}