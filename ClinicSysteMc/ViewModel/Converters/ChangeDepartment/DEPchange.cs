namespace ClinicSysteMc.ViewModel.Converters
{
    /// <summary>
    /// 201806?? AutoIt version created.
    /// 20190606 VB version created, 目的再深化自動化
    /// 20190608 加好了try, record_adm, record_err
    /// 目前穩定,已經使用了大約一年, 意思是用前身AutoIt版本, 大概是201806中開始的
    /// 20190607 created
    /// 20200518 transcribed into c-sharp
    /// 20210721 要修正幾個地方: 1. 葉醫師部分不要動, 2. 要有感知功能, 3. 要能中斷, 4. 可以計算真實改變的數字
    /// </summary>
    public partial class DEPchange
    {
        private readonly string _strYM;

        public DEPchange(string YM)
        {
            _strYM = YM;
        }

        public void Change()
        {
            Dash d = new Dash(_strYM);

            d.Show();
        }
    }
}