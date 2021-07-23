using ClinicSysteMc.Model;
using Hardcodet.Wpf.TaskbarNotification;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ClinicSysteMc.ViewModel.Converters
{
    public partial class Dash
    {
        public List<sp_change_depResult> MakeCSV()
        {
            // Dim output As DEP_return = Change_DEP(strYM)
            // MessageBox.Show("修改了" + output.m.ToString + "筆, 請匯入門診資料")
            string savepath = @"C:\vpn\change_dep";
            int change_N;
            DateTime minD;
            DateTime maxD;
            List<sp_change_depResult> List_Change;

            // 存放目錄,不存在就要建立一個
            if (!(System.IO.Directory.Exists(savepath))) System.IO.Directory.CreateDirectory(savepath);

            #region Making CSV

            try
            {
                // 呼叫SQL stored procedure
                using (CSDataContext dc = new CSDataContext())
                {
                    List_Change = dc.sp_change_dep(_strYM).ToList();
                }

                // 自動產生名字
                string savefile = $"\\change_dep_{_strYM}_{DateTime.Now.Year}{(DateTime.Now.Month + 100).ToString().Substring(1)}";
                savefile += $"{(DateTime.Now.Day + 100).ToString().Substring(1)}_{DateTime.Now.TimeOfDay}";
                savefile = savefile.Replace(":", "").Replace(".", "");
                savepath += $"{savefile}.csv";

                // 製作csv檔 writing to csv
                System.IO.StreamWriter sw = new System.IO.StreamWriter(savepath);
                int i = 1;
                change_N = List_Change.Count;
                if (change_N == 0)
                {
                    tb.ShowBalloonTip("完成", "沒有什麼需要修改的", BalloonIcon.Info);
                    log.Info("change department: 沒有什麼需要修改的");
                    Logging.Record_admin("change department", "沒有什麼需要修改的");
                }
                else
                {
                    minD = DateTime.Parse("9999/12/31");
                    maxD = DateTime.Parse("0001/01/01");
                    foreach (var c in List_Change)
                    {
                        sw.Write(c.o); // 欄位名叫o
                        if (i < change_N) sw.Write(sw.NewLine);
                        DateTime tempD = DateTime.Parse($"{c.o.Substring(0, 4)}/{c.o.Substring(4, 2)}/{c.o.Substring(6, 2)}");
                        // 找尋最大的值
                        if (tempD.CompareTo(maxD) > 0) maxD = tempD;
                        // 找尋最小的值
                        if (tempD.CompareTo(minD) < 0) minD = tempD;
                        i++;
                    }
                    // 20200518 放在foreach的loop迴圈裡是錯誤的, 我把它放出來了
                    string output = $"{minD:d}~{maxD:d}, 共{change_N}筆需要修改";
                    tb.ShowBalloonTip("需修改:", output, BalloonIcon.Info);
                    log.Info($"change department: {output}");
                    Logging.Record_admin("change department", output);
                    sw.Close();
                }
                return List_Change;
            }
            catch (System.Exception ex)
            {
                string o = $"Something wrong in making CSV: {ex.Message}";
                Logging.Record_error(o);
                log.Error(o);
                return null;
            }
            #endregion Making CSV
        }
    }
}
