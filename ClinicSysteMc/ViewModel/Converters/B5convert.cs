using ClinicSysteMc.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ClinicSysteMc.ViewModel.Converters
{
    internal class B5convert
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private readonly string _loadpath;

        public B5convert(string loadpath)
        {
            this._loadpath = loadpath;
        }

        internal async void Transform()
        {
            string[] Lines = System.IO.File.ReadAllLines(_loadpath, System.Text.Encoding.Default);

            int totalN = Lines.Length;  // -1 because line 1 is titles, so I should begin with 2 to total_N + 1
            // now I should divide the array into 500 lines each and store it into a list.

            int table_N = 10000;
            int total_div = totalN / table_N;
            int residual = totalN % table_N;

            log.Info($"  start async process.");
            List<Task<PTresult>> tasks = new List<Task<PTresult>>();

            // 將_data分拆成幾個小的Array
            for (int i = 0, idx = 0; i <= total_div; i++, idx += table_N)
            {
                string[] dummy;
                if (i < total_div)
                {
                    dummy = new string[table_N];
                    Array.Copy(Lines, idx, dummy, 0, table_N);
                }
                else
                {
                    dummy = new string[residual];
                    Array.Copy(Lines, idx, dummy, 0, residual);
                }
                tasks.Add(ImportB5_async(dummy));
            }

            PTresult[] result = await Task.WhenAll(tasks);

            int total_NewPT = (from p in result
                               select p.NewPT).Sum();
            int total_ChangePT = (from p in result
                                  select p.ChangePT).Sum();
            int total_AllPT = (from p in result
                               select p.AllPT).Sum();
            log.Info($"  end async process.");

            string output = $"共處理{total_AllPT}筆資料, 其中{total_NewPT}筆新資料, 修改{total_ChangePT}筆資料.";
            log.Info(output);
            Logging.Record_admin("B5 add/change", output);
        }

        private async Task<PTresult> ImportB5_async(string[] Lines)
        {
            int totalN = Lines.Length;
            int add_N = 0;
            int change_N = 0;
            int all_N = 0;

            await Task.Run(() =>
            {
                foreach (string Line in Lines)
                {
                    byte[] lineStr = System.Text.Encoding.Default.GetBytes(Line);
                    string sNewMark = System.Text.Encoding.Default.GetString(lineStr, 0, 2).Trim();          // 1 New_mark c 2 1 2
                    string sOralMed = System.Text.Encoding.Default.GetString(lineStr, 3, 10).Trim();         // 2 口服錠註記 c 10 4 13
                    string sComplex = System.Text.Encoding.Default.GetString(lineStr, 14, 2).Trim();         // 3 單 / 複方註記 c 2 15 16
                    string sNHI_code = System.Text.Encoding.Default.GetString(lineStr, 17, 10).Trim();       // 4 藥品代碼 c 10 18 27
                    string sNHI_price = System.Text.Encoding.Default.GetString(lineStr, 28, 9).Trim();       // 5 藥價參考金額 N 9,2 29 37
                    string sbegin_date = System.Text.Encoding.Default.GetString(lineStr, 38, 7).Trim();      // 6 藥價參考日期 D 7 39 45
                    string send_date = System.Text.Encoding.Default.GetString(lineStr, 46, 7).Trim();        // 7 藥價參考截止日期 D 7 47 53
                    string sename = System.Text.Encoding.Default.GetString(lineStr, 54, 120).Trim();         // 8 藥品英文名稱 c 120 55 174
                    string sspec_dose = System.Text.Encoding.Default.GetString(lineStr, 175, 7).Trim();      // 9 藥品規格量 N 7,2 176 182
                    string sspec_unit = System.Text.Encoding.Default.GetString(lineStr, 183, 52).Trim();     // 10 藥品規格單位 c 52 184 235
                    string scomp_name = System.Text.Encoding.Default.GetString(lineStr, 236, 56).Trim();     // 11 成份名稱 c 56 237 292
                    string scomp_dose = System.Text.Encoding.Default.GetString(lineStr, 293, 12).Trim();     // 12 成份含量 N 12,3 294 305
                    string scomp_unit = System.Text.Encoding.Default.GetString(lineStr, 306, 51).Trim();     // 13 成份含量單位 c 51 307 357
                    string sprep = System.Text.Encoding.Default.GetString(lineStr, 358, 86).Trim();          // 14 藥品劑型 c 86 359 444
                    string svendor = System.Text.Encoding.Default.GetString(lineStr, 604, 20).Trim();        // 16 藥商名稱 c 20 605 624
                    string sclas = System.Text.Encoding.Default.GetString(lineStr, 767, 1).Trim();           // 18 藥品分類 c 1 768 768
                    string squality = System.Text.Encoding.Default.GetString(lineStr, 769, 1).Trim();        // 19 品質分類碼 c 1 770 770
                    string scname = System.Text.Encoding.Default.GetString(lineStr, 771, 128).Trim();        // 20 藥品中文名稱 c 128 772 899
                    string sgroup_name = System.Text.Encoding.Default.GetString(lineStr, 900, 300).Trim();   // 21 分類分組名稱 c 300 901 1200
                    string scomp1 = System.Text.Encoding.Default.GetString(lineStr, 1200, 56).Trim();        // 22 （複方一）成份名稱 c 56 1201 1256
                    string scomp1_dose = System.Text.Encoding.Default.GetString(lineStr, 1258, 11).Trim();   // 23 （複方一）藥品成份含量 N 11,3 1259 1269
                    string scomp1_unit = System.Text.Encoding.Default.GetString(lineStr, 1270, 51).Trim();   // 24 （複方一）藥品成份含量單位 c 51 1271 1321
                    string scomp2 = System.Text.Encoding.Default.GetString(lineStr, 1322, 56).Trim();        // 25 （複方二）成份名稱 c 56 1323 1378
                    string scomp2_dose = System.Text.Encoding.Default.GetString(lineStr, 1379, 11).Trim();   // 26 （複方二）藥品成份含量 N 11,3 1380 1390
                    string scomp2_unit = System.Text.Encoding.Default.GetString(lineStr, 1391, 51).Trim();   // 27 （複方二）藥品成份含量單位 c 51 1392 1442
                    string scomp3 = System.Text.Encoding.Default.GetString(lineStr, 1443, 56).Trim();        // 28 （複方三）成份名稱 c 56 1444 1499
                    string scomp3_dose = System.Text.Encoding.Default.GetString(lineStr, 1500, 11).Trim();   // 29 （複方三）藥品成份含量 N 11,3 1501 1511
                    string scomp3_unit = System.Text.Encoding.Default.GetString(lineStr, 1512, 51).Trim();   // 30 （複方三）藥品成份含量單位 c 51 1513 1563
                    string scomp4 = System.Text.Encoding.Default.GetString(lineStr, 1564, 56).Trim();        // 31 （複方四）成份名稱 c 56 1565 1620
                    string scomp4_dose = System.Text.Encoding.Default.GetString(lineStr, 1621, 11).Trim();   // 32 （複方四）藥品成份含量 N 11,3 1622 1632
                    string scomp4_unit = System.Text.Encoding.Default.GetString(lineStr, 1633, 51).Trim();   // 33 （複方四）藥品成份含量單位 c 51 1634 1684
                    string scomp5 = System.Text.Encoding.Default.GetString(lineStr, 1685, 56).Trim();        // 34 （複方五）成份名稱 c 56 1686 1741
                    string scomp5_dose = System.Text.Encoding.Default.GetString(lineStr, 1742, 11).Trim();   // 35 （複方五）藥品成份含量 N 11,3 1743 1753
                    string scomp5_unit = System.Text.Encoding.Default.GetString(lineStr, 1754, 51).Trim();   // 36 （複方五）藥品成份含量單位 c 51 1755 1805
                    string smanufacturer = System.Text.Encoding.Default.GetString(lineStr, 1807, 42).Trim(); // 37 製造廠名稱 c 42 1807 1848
                    string sATC_code = System.Text.Encoding.Default.GetString(lineStr, 1849, 8).Trim();      // 38 ATC CODE c 8 1850 1857
                    string sNoProduce = System.Text.Encoding.Default.GetString(lineStr, 1858, 1).Trim();     // 39 未生產或未輸入達五年 c 1 1859 1859 108.5.21.新增

                    using (BSDataContext dc = new BSDataContext())
                    {
                        var q = from p in dc.NHI_med
                                where (p.NHI_code == sNHI_code) && (p.begin_date == sbegin_date)
                                select p;
                        if (q.Count() == 0)
                        {
                            try
                            {
                                NHI_med newNHI = new NHI_med()
                                {
                                    NewMark = sNewMark,         // 1 New_mark c 2 1 2
                                    OralMed = sOralMed,         // 2 口服錠註記 c 10 4 13
                                    Complex = sComplex,         // 3 單 / 複方註記 c 2 15 16
                                    NHI_code = sNHI_code,       // 4 藥品代碼 c 10 18 27
                                    NHI_price = sNHI_price,     // 5 藥價參考金額 N 9,2 29 37
                                    begin_date = sbegin_date,   // 6 藥價參考日期 D 7 39 45
                                    end_date = send_date,       // 7 藥價參考截止日期 D 7 47 53
                                    ename = sename,             // 8 藥品英文名稱 c 120 55 174
                                    spec_dose = sspec_dose,     // 9 藥品規格量 N 7,2 176 182
                                    spec_unit = sspec_unit,     // 10 藥品規格單位 c 52 184 235
                                    comp_name = scomp_name,     // 11 成份名稱 c 56 237 292
                                    comp_dose = scomp_dose,     // 12 成份含量 N 12,3 294 305
                                    comp_unit = scomp_unit,     // 13 成份含量單位 c 51 307 357
                                    prep = sprep,               // 14 藥品劑型 c 86 359 444
                                    vendor = svendor,           // 16 藥商名稱 c 20 605 624
                                    clas = sclas,               // 18 藥品分類 c 1 768 768
                                    quality = squality,         // 19 品質分類碼 c 1 770 770
                                    cname = scname,             // 20 藥品中文名稱 c 128 772 899
                                    group_name = sgroup_name,   // 21 分類分組名稱 c 300 901 1200
                                    comp1 = scomp1,             // 22 （複方一）成份名稱 c 56 1201 1256
                                    comp1_dose = scomp1_dose,   // 23 （複方一）藥品成份含量 N 11,3 1259 1269
                                    comp1_unit = scomp1_unit,   // 24 （複方一）藥品成份含量單位 c 51 1271 1321
                                    comp2 = scomp2,             // 25 （複方二）成份名稱 c 56 1323 1378
                                    comp2_dose = scomp2_dose,   // 26 （複方二）藥品成份含量 N 11,3 1380 1390
                                    comp2_unit = scomp2_unit,   // 27 （複方二）藥品成份含量單位 c 51 1392 1442
                                    comp3 = scomp3,             // 28 （複方三）成份名稱 c 56 1444 1499
                                    comp3_dose = scomp3_dose,   // 29 （複方三）藥品成份含量 N 11,3 1501 1511
                                    comp3_unit = scomp3_unit,   // 30 （複方三）藥品成份含量單位 c 51 1513 1563
                                    comp4 = scomp4,             // 31 （複方四）成份名稱 c 56 1565 1620
                                    comp4_dose = scomp4_dose,   // 32 （複方四）藥品成份含量 N 11,3 1622 1632
                                    comp4_unit = scomp4_unit,   // 33 （複方四）藥品成份含量單位 c 51 1634 1684
                                    comp5 = scomp5,             // 34 （複方五）成份名稱 c 56 1686 1741
                                    comp5_dose = scomp5_dose,   // 35 （複方五）藥品成份含量 N 11,3 1743 1753
                                    comp5_unit = scomp5_unit,   // 36 （複方五）藥品成份含量單位 c 51 1755 1805
                                    manufacturer = smanufacturer, // 37 製造廠名稱 c 42 1807 1848
                                    ATC_code = sATC_code,       // 38 ATC CODE c 8 1850 1857
                                    NoProduce = sNoProduce      // 39 未生產或未輸入達五年 c 1 1859 1859 108.5.21.新增
                                };
                                dc.NHI_med.InsertOnSubmit(newNHI);
                                dc.SubmitChanges();
                                add_N++;
                            }
                            catch (Exception ex)
                            {
                                Logging.Record_error(ex.Message);
                                log.Error(ex.Message);
                            }
                        }
                        else
                        {
                            try
                            {
                                // only one if any
                                bool bChanged = false;
                                string strChange = string.Empty;
                                NHI_med oldNHI = q.First();
                                if (oldNHI.NewMark != sNewMark)
                                {
                                    strChange += $"New Mark: {oldNHI.NewMark} => {sNewMark};";
                                    oldNHI.NewMark = sNewMark;
                                    bChanged = true;
                                }         // 1 New_mark c 2 1 2
                                if (oldNHI.OralMed != sOralMed)
                                {
                                    strChange += $"口服錠註記: {oldNHI.OralMed} => {sOralMed};";
                                    oldNHI.OralMed = sOralMed;
                                    bChanged = true;
                                }         // 2 口服錠註記 c 10 4 13
                                if (oldNHI.Complex != sComplex)
                                {
                                    strChange += $"單/複方註記: {oldNHI.Complex} => {sComplex};";
                                    oldNHI.Complex = sComplex;
                                    bChanged = true;
                                }         // 3 單 / 複方註記 c 2 15 16
                                if (oldNHI.NHI_price != sNHI_price)
                                {
                                    strChange += $"藥價參考金額: {oldNHI.NHI_price} => {sNHI_price};";
                                    oldNHI.NHI_price = sNHI_price;
                                    bChanged = true;
                                }         // 5 藥價參考金額 N 9,2 29 37
                                if (oldNHI.end_date != send_date)
                                {
                                    strChange += $"藥價參考截止日期: {oldNHI.end_date} => {send_date};";
                                    oldNHI.end_date = send_date;
                                    bChanged = true;
                                }       // 7 藥價參考截止日期 D 7 47 53
                                if (oldNHI.ename != sename)
                                {
                                    strChange += $"藥品英文名稱: {oldNHI.ename} => {sename};";
                                    oldNHI.ename = sename;
                                    bChanged = true;
                                }         // 8 藥品英文名稱 c 120 55 174
                                if (oldNHI.spec_dose != sspec_dose)
                                {
                                    strChange += $"藥品規格量: {oldNHI.spec_dose} => {sspec_dose};";
                                    oldNHI.spec_dose = sspec_dose;
                                    bChanged = true;
                                }     // 9 藥品規格量 N 7,2 176 182
                                if (oldNHI.spec_unit != sspec_unit)
                                {
                                    strChange += $"藥品規格單位: {oldNHI.spec_unit} => {sspec_unit};";
                                    oldNHI.spec_unit = sspec_unit;
                                    bChanged = true;
                                }     // 10 藥品規格單位 c 52 184 235
                                if (oldNHI.comp_name != scomp_name)
                                {
                                    strChange += $"成份名稱: {oldNHI.comp_name} => {scomp_name};";
                                    oldNHI.comp_name = scomp_name;
                                    bChanged = true;
                                }     // 11 成份名稱 c 56 237 292
                                if (oldNHI.comp_dose != scomp_dose)
                                {
                                    strChange += $"成份含量: {oldNHI.comp_dose} => {scomp_dose};";
                                    oldNHI.comp_dose = scomp_dose;
                                    bChanged = true;
                                }     // 12 成份含量 N 12,3 294 305
                                if (oldNHI.comp_unit != scomp_unit)
                                {
                                    strChange += $"成份含量單位: {oldNHI.comp_unit} => {scomp_unit};";
                                    oldNHI.comp_unit = scomp_unit;
                                    bChanged = true;
                                }     // 13 成份含量單位 c 51 307 357
                                if (oldNHI.prep != sprep)
                                {
                                    strChange += $"藥品劑型: {oldNHI.prep} => {sprep};";
                                    oldNHI.prep = sprep;
                                    bChanged = true;
                                }               // 14 藥品劑型 c 86 359 444
                                if (oldNHI.vendor != svendor)
                                {
                                    strChange += $"藥商名稱: {oldNHI.vendor} => {svendor};";
                                    oldNHI.vendor = svendor;
                                    bChanged = true;
                                }           // 16 藥商名稱 c 20 605 624
                                if (oldNHI.clas != sclas)
                                {
                                    strChange += $"藥品分類: {oldNHI.clas} => {sclas};";
                                    oldNHI.clas = sclas;
                                    bChanged = true;
                                }               // 18 藥品分類 c 1 768 768
                                if (oldNHI.quality != squality)
                                {
                                    strChange += $"品質分類碼: {oldNHI.quality} => {squality};";
                                    oldNHI.quality = squality;
                                    bChanged = true;
                                }         // 19 品質分類碼 c 1 770 770
                                if (oldNHI.cname != scname)
                                {
                                    strChange += $"藥品中文名稱: {oldNHI.cname} => {scname};";
                                    oldNHI.cname = scname;
                                    bChanged = true;
                                }             // 20 藥品中文名稱 c 128 772 899
                                if (oldNHI.group_name != sgroup_name)
                                {
                                    strChange += $"分類分組名稱: {oldNHI.group_name} => {sgroup_name};";
                                    oldNHI.group_name = sgroup_name;
                                    bChanged = true;
                                }   // 21 分類分組名稱 c 300 901 1200
                                if (oldNHI.comp1 != scomp1)
                                {
                                    strChange += $"（複方一）成份名稱: {oldNHI.comp1} => {scomp1};";
                                    oldNHI.comp1 = scomp1;
                                    bChanged = true;
                                }             // 22 （複方一）成份名稱 c 56 1201 1256
                                if (oldNHI.comp1_dose != scomp1_dose)
                                {
                                    strChange += $"（複方一）藥品成份含量: {oldNHI.comp1_dose} => {scomp1_dose};";
                                    oldNHI.comp1_dose = scomp1_dose;
                                    bChanged = true;
                                }   // 23 （複方一）藥品成份含量 N 11,3 1259 1269
                                if (oldNHI.comp1_unit != scomp1_unit)
                                {
                                    strChange += $"（複方一）藥品成份含量單位: {oldNHI.comp1_unit} => {scomp1_unit};";
                                    oldNHI.comp1_unit = scomp1_unit;
                                    bChanged = true;
                                }   // 24 （複方一）藥品成份含量單位 c 51 1271 1321
                                if (oldNHI.comp2 != scomp2)
                                {
                                    strChange += $"（複方二）成份名稱: {oldNHI.comp2} => {scomp2};";
                                    oldNHI.comp2 = scomp2;
                                    bChanged = true;
                                }             // 25 （複方二）成份名稱 c 56 1323 1378
                                if (oldNHI.comp2_dose != scomp2_dose)
                                {
                                    strChange += $"26 （複方二）藥品成份含量: {oldNHI.comp2_dose} => {scomp2_dose};";
                                    oldNHI.comp2_dose = scomp2_dose;
                                    bChanged = true;
                                }   // 26 （複方二）藥品成份含量 N 11) {}3 1380 1390
                                if (oldNHI.comp2_unit != scomp2_unit)
                                {
                                    strChange += $"（複方二）藥品成份含量單位: {oldNHI.comp2_unit} => {scomp2_unit};";
                                    oldNHI.comp2_unit = scomp2_unit;
                                    bChanged = true;
                                }   // 27 （複方二）藥品成份含量單位 c 51 1392 1442
                                if (oldNHI.comp3 != scomp3)
                                {
                                    strChange += $"（複方三）成份名稱: {oldNHI.comp3} => {scomp3};";
                                    oldNHI.comp3 = scomp3;
                                    bChanged = true;
                                }             // 28 （複方三）成份名稱 c 56 1444 1499
                                if (oldNHI.comp3_dose != scomp3_dose)
                                {
                                    strChange += $"（複方三）藥品成份含量: {oldNHI.comp3_dose} => {scomp3_dose};";
                                    oldNHI.comp3_dose = scomp3_dose;
                                    bChanged = true;
                                }   // 29 （複方三）藥品成份含量 N 11,3 1501 1511
                                if (oldNHI.comp3_unit != scomp3_unit)
                                {
                                    strChange += $"（複方三）藥品成份含量單位: {oldNHI.comp3_unit} => {scomp3_unit};";
                                    oldNHI.comp3_unit = scomp3_unit;
                                    bChanged = true;
                                }   // 30 （複方三）藥品成份含量單位 c 51 1513 1563
                                if (oldNHI.comp4 != scomp4)
                                {
                                    strChange += $"（複方四）成份名稱: {oldNHI.comp4} => {scomp4};";
                                    oldNHI.comp4 = scomp4;
                                    bChanged = true;
                                }             // 31 （複方四）成份名稱 c 56 1565 1620
                                if (oldNHI.comp4_dose != scomp4_dose)
                                {
                                    strChange += $"（複方四）藥品成份含量: {oldNHI.comp4_dose} => {scomp4_dose};";
                                    oldNHI.comp4_dose = scomp4_dose;
                                    bChanged = true;
                                }   // 32 （複方四）藥品成份含量 N 11,3 1622 1632
                                if (oldNHI.comp4_unit != scomp4_unit)
                                {
                                    strChange += $"（複方四）藥品成份含量單位: {oldNHI.comp4_unit} => {scomp4_unit};";
                                    oldNHI.comp4_unit = scomp4_unit;
                                    bChanged = true;
                                }   // 33 （複方四）藥品成份含量單位 c 51 1634 1684
                                if (oldNHI.comp5 != scomp5)
                                {
                                    strChange += $"（複方五）成份名稱: {oldNHI.comp5} => {scomp5};";
                                    oldNHI.comp5 = scomp5;
                                    bChanged = true;
                                }             // 34 （複方五）成份名稱 c 56 1686 1741
                                if (oldNHI.comp5_dose != scomp5_dose)
                                {
                                    strChange += $"（複方五）藥品成份含量: {oldNHI.comp5_dose} => {scomp5_dose};";
                                    oldNHI.comp5_dose = scomp5_dose;
                                    bChanged = true;
                                }   // 35 （複方五）藥品成份含量 N 11,3 1743 1753
                                if (oldNHI.comp5_unit != scomp5_unit)
                                {
                                    strChange += $"（複方五）藥品成份含量單位: {oldNHI.comp5_unit} => {scomp5_unit};";
                                    oldNHI.comp5_unit = scomp5_unit;
                                    bChanged = true;
                                }   // 36 （複方五）藥品成份含量單位 c 51 1755 1805
                                if (oldNHI.manufacturer != smanufacturer)
                                {
                                    strChange += $"製造廠名稱: {oldNHI.manufacturer} => {smanufacturer};";
                                    oldNHI.manufacturer = smanufacturer;
                                    bChanged = true;
                                } // 37 製造廠名稱 c 42 1807 1848
                                if (oldNHI.ATC_code != sATC_code)
                                {
                                    strChange += $"ATC CODE: {oldNHI.ATC_code} => {sATC_code};";
                                    oldNHI.ATC_code = sATC_code;
                                    bChanged = true;
                                }       // 38 ATC CODE c 8 1850 1857
                                if (oldNHI.NoProduce != sNoProduce)
                                {
                                    strChange += $"未生產或未輸入達五年: {oldNHI.NoProduce} => {sNoProduce};";
                                    oldNHI.NoProduce = sNoProduce;
                                    bChanged = true;
                                }      // 39 未生產或未輸入達五年 c 1 1859 1859 108.5.21.新增
                                if (bChanged)
                                {
                                    // 做實改變
                                    dc.SubmitChanges();
                                    // 做記錄
                                    // 20190929 加姓名, 病歷號
                                    Logging.Record_admin("Change b5 data", $"{sNHI_code}: {strChange}");
                                    log.Info($"Change b5 data: {sNHI_code}: {strChange}");
                                    change_N++;
                                }
                            }
                            catch (Exception ex)
                            {
                                Logging.Record_error(ex.Message);
                                log.Error(ex.Message);
                            }
                        }
                    };
                    all_N++;
                }
            });
            return new PTresult()
            {
                NewPT = add_N,
                ChangePT = change_N,
                AllPT = all_N
            };
        }
    }
}