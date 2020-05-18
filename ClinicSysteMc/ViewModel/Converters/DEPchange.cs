using AutoIt;
using ClinicSysteMc.Model;
using Hardcodet.Wpf.TaskbarNotification;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ClinicSysteMc.ViewModel.Converters
{
    internal class DEPchange
    {
        // 20190606 created, 目的再深化自動化
        // 20190608 加好了try, record_adm, record_err
        // 目前穩定,已經使用了大約一年, 意思是用前身AutoIt版本, 大概是201806中開始的
        // 20190607 created
        // 20200518 transcribed into c-sharp

        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private readonly TaskbarIcon tb = new TaskbarIcon();
        private readonly string _strYM;

        public DEPchange(string YM)
        {
            _strYM = YM;
        }

        public void Change()
        {
            // Dim output As DEP_return = Change_DEP(strYM)
            // MessageBox.Show("修改了" + output.m.ToString + "筆, 請匯入門診資料")
            string savepath = @"C:\vpn\change_dep";
            int change_N = 0;
            DateTime minD;
            DateTime maxD;

            // 存放目錄,不存在就要建立一個
            if (!(System.IO.Directory.Exists(savepath))) System.IO.Directory.CreateDirectory(savepath);

            #region Making CSV

            try
            {
                // 呼叫SQL stored procedure
                List<sp_change_depResult> ListChange;
                using (CSDataContext dc = new CSDataContext())
                {
                    ListChange = dc.sp_change_dep(_strYM).ToList();
                }

                // 自動產生名字
                string savefile = $"\\change_dep_{_strYM}_{DateTime.Now.Year}{(DateTime.Now.Month + 100).ToString().Substring(1)}";
                savefile += $"{(DateTime.Now.Day + 100).ToString().Substring(1)}_{DateTime.Now.TimeOfDay}.csv";
                savefile = savefile.Replace(":", "").Replace(".", "");
                savepath += savefile;

                // 製作csv檔 writing to csv
                System.IO.StreamWriter sw = new System.IO.StreamWriter(savepath);
                int i = 1;
                change_N = ListChange.Count;
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
                    foreach (var c in ListChange)
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
                    sw.Close();
                }
            }
            catch (System.Exception ex)
            {
                Logging.Record_error(ex.Message);
                log.Error(ex.Message);
                return;
            }

            #endregion Making CSV

            #region Environment

            try
            {
                // 營造環境
                if (AutoItX.WinExists("看診清單") == 1) //如果直接存在就直接叫用
                {
                    AutoItX.WinActivate("看診清單");
                }
                else
                {
                    Thesis.LogIN();
                    // 打開"看診清單"
                    AutoItX.Run(@"C:\Program Files (x86)\THESE\杏雲醫療資訊系統\THCClinic.exe", @"C:\Program Files (x86)\THESE\杏雲醫療資訊系統\");
                }
            }
            catch (Exception ex)
            {
                Logging.Record_error(ex.Message);
                log.Error(ex.Message);
                return;
            }

            #endregion Environment

            /*

#Region "Execute change department"
        Try
            aut.WinWaitActive("看診清單")
            Shell("C:\vpn\exe\changeDP.exe " + savepath, AppWinStyle.Hide, True)
            '            MessageBox.Show("修改了" + m.ToString + "筆, 請匯入門診資料")
            Record_adm("change department", "修改了" + return_value.m.ToString + "筆")
            Return return_value
        Catch ex As Exception
            Record_error(ex.ToString)
            Return return_value
        End Try
#End Region

        // 製造一個AutoIT VB程式, changeDP_DATE_VIST_ROOM, 針對"看診清單", [NAME:dtpSDate], [NAME:cmbVist], [NAME:cmbRmno]
        // 一個參數, 格式YYYYMMDDVRR

        // 製造一個AutoIT VB程式, changeDP_DEP, 針對"問診畫面"[NAME:cmbDept]
        // 一個參數, 格式DD

#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <File.au3>
#include <MsgBoxConstants.au3>
#include <Constants.au3>
#include <GuiComboBox.au3>
#include <GuiDateTimePicker.au3>

Local $iLineCount = 0
Local $file = $Cmdline[1] ;"C:\201905_should_be_in_family.csv"

If WinExists("看診清單") Then
	WinActivate("看診清單")
EndIf

Sleep(500)

$iLineCount = _FileCountLines($file)
;msgbox($MB_SYSTEMMODAL, "table", $iLineCount)
Local $otDate[7]=[False,"1973","1","7",0,0,0]
Local $oV="1"
Local $oR="01"
Local $bChange=False

for $iCount = 1 to $iLineCount
	$line = FileReadLine($file, $iCount)

	; 日期, what a strange way to solve the problem
	; 先設定好要的日期
	; 然後set focus, 再上下,才能把改變的資料傳入系統,否則,系統不知道
	Dim $tDate[7] = [False, StringMid($line,1,4), StringMid($line,5,2), StringMid($line,7,2), 0, 0, 0]
	if ($tDate[2]<>$otDate[2]) or  ($tDate[3]<>$otDate[3]) or  ($tDate[4]<>$otDate[4]) Then
		$hDTP=ControlGetHandle("看診清單", "", "[NAME:dtpSDate]")
		_GUICtrlDTP_SetSystemTime($hDTP,$tDate)
		ControlFocus("看診清單", "", "[NAME:dtpSDate]")
		Send("{Up}")
		Send("{Down}")
		$otDate[2]=$tDate[2]
		$otDate[3]=$tDate[3]
		$otDate[4]=$tDate[4]
		$bChange=True
		sleep(1000)
	endif
	; VIST
	;msgbox(0, "", "the line "&$iCount&" is "&StringMid($line,9,1))
	;午別
	; combobox1, 午別, 0 全部, (1 上午, 2 下午, 3 晚上)好像value 不固定,只是順序而已
	if $oV<>StringMid($line,9,1) then
		$hCB1=ControlGetHandle("看診清單", "", "[NAME:cmbVist]")
		;_GUICtrlComboBox_SetCurSel($hCB1,1)
		_GUICtrlComboBox_SelectString($hCB1, StringMid($line,9,1))
		$oV=StringMid($line,9,1)
		$bChange=True
		sleep(1000)
	endif
	sleep(1000)
	; Room NO
	;msgbox(0, "", "the line "&$iCount&" is "&StringMid($line,10,2))
	;診別
	; combobox2, 診間
	if $oR<>StringMid($line,10,2) then
		$hCB2=ControlGetHandle("看診清單", "", "[NAME:cmbRmno]")
		;_GUICtrlComboBox_SetCurSel($hCB2,1)
		_GUICtrlComboBox_SelectString($hCB2, StringMid($line,10,2))
		$oV=StringMid($line,10,2)
		$bChange=True
	endif

	sleep(1000)
	if $bChange=True Then
		ControlClick("看診清單", "", "[NAME:btnRefresh]")
		$bChange=False
	endif
	WinWaitActive("看診清單")

	; Seq NO
	; msgbox(0, "", "the line "&$iCount&" is "&StringMid($line,12,3))
	;輸入診號
	ControlSend("看診清單", "", "[NAME:txbSqno]",StringMid($line,12,3))

	; 按鈕
	ControlClick("看診清單", "", "[NAME:btnGo]")

	WinWaitActive("問診畫面")

	; 然後set focus, 再上下,才能把改變的資料傳入系統,否則,系統不知道
	$hCB1=ControlGetHandle("問診畫面", "", "[NAME:cmbDept]")
	;_GUICtrlComboBox_SetCurSel($hCB1,1)
	If StringLen($line)=16 Then
		if StringMid($line,15,2)="01" Then
			_GUICtrlComboBox_SetCurSel($hCB1,1)
		elseif StringMid($line,15,2)="13" Then
			_GUICtrlComboBox_SetCurSel($hCB1,13)
		endif
	Else
		_GUICtrlComboBox_SetCurSel($hCB1,1)
	endif

	Sleep(1500)
	Send("{F9}")
	;	ControlClick("問診畫面", "", "[NAME:btnCancel]")

	; 已經看過診了,是否繼續
	Sleep(1500)
	ControlClick("THCClinic", "", "[CLASSNN:Button1]")

	; 超過8種藥物(不一定有)
	Sleep(1500)
	ControlClick("THCClinic", "", "[CLASSNN:Button1]")

	; 是否重印收據
	WinWaitActive("These.CludCln.Accounting")
	ControlClick("These.CludCln.Accounting", "", "[CLASSNN:Button2]")

Next
fileclose($file)

             */
        }
    }
}