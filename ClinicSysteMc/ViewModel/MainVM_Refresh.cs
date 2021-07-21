using ClinicSysteMc.Model;
using Hardcodet.Wpf.TaskbarNotification;
using System.ComponentModel;
using System.Linq;

namespace ClinicSysteMc.ViewModel
{
    internal partial class MainVM : INotifyPropertyChanged
    {
        public void Refresh_Data()
        {
            CSDataContext dc = new CSDataContext();

            #region Function Page

            LogInOut = (from p in dc.log_Adm
                        where p.operation_name == "Log in" || p.operation_name == "Log out"
                        orderby p.regdate descending
                        select new { p.regdate, p.operation_name }).Take(100);
            // 20200522 add opd_import, for correct display
            OPD = (from p in dc.log_Adm
                   where p.operation_name == "add opd" || p.operation_name == "opd_import"
                   orderby p.regdate descending
                   select new { p.regdate }).Take(100);
            PT = (from p in dc.log_Adm
                  where p.operation_name == "病患檔案格式"
                  orderby p.regdate descending
                  select new { p.regdate }).Take(100);
            Order = (from p in dc.log_Adm
                     where p.operation_name == "計價檔格式"
                     orderby p.regdate descending
                     select new { p.regdate }).Take(100);
            Upload = (from p in dc.log_Adm
                      where p.operation_name == "健保上傳XML檔配對"
                      orderby p.regdate descending
                      select new { p.regdate }).Take(100);
            // 20200522 add PIJIA add/change, for correct display
            Pijia = (from p in dc.log_Adm
                     where p.operation_name == "新增批價檔: " || p.operation_name == "PIJIA add/change"
                     orderby p.regdate descending
                     select new { p.regdate }).Take(100);
            ChangeDepartment = (from p in dc.log_Adm
                                where p.operation_name == "change department"
                                orderby p.regdate descending
                                select new { p.regdate }).Take(100);
            Lab = (from p in dc.log_Adm
                   where p.operation_name == "Lab file format"
                   orderby p.regdate descending
                   select new { p.regdate }).Take(100);

            #endregion Function Page

            tb.ShowBalloonTip("完成", "主頁資料已更新", BalloonIcon.Info);
        }
    }
}
