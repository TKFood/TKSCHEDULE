using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;
using System.Media;

namespace TKSCHEDULE
{
    public partial class Form1 : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlConnection sqlConn2 = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        int result;
        string HRAUTO = "N";

        public Form1()
        {
            InitializeComponent();
            timer1.Enabled = true;
            timer1.Interval = 1000;
            timer1.Start();
        }

        #region FUNCTION
        private void timer1_Tick(object sender, EventArgs e)
        {
            label1.Text = DateTime.Now.ToString();
            HRAUTORUN();

        }
        public void SETbutton1()
        {
            if(button1.Text.Equals("啟動"))
            {
                HRAUTO = "Y";
                button1.Text = "停止";
                button1.ForeColor = Color.Red;
            }
            else if(button1.Text.Equals("停止"))
            {
                HRAUTO = "N";
                button1.Text = "啟動";
                button1.ForeColor = Color.Blue;
            }
        }
        public void HRAUTORUN()
        {
            string hh = "16";
            string mm = "45";
            if (HRAUTO.Equals("Y")&& DateTime.Now.Hour.ToString().Equals(hh)&&DateTime.Now.Minute.ToString().Equals(mm))
            {
                label3.Text = "YYYYYYY";
                ADDHRCARD();
            }
            else
            {
                label3.Text = "GO";
            }     
            
        }
        public void ADDHRCARD()
        {
            DateTime workdt1 = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day+1, 8, 30, 0);
            DateTime workdt2 = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day + 1, 17, 30, 0);
            DateTime carddt1 = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day + 1, 8, 20, 0);            DateTime carddt2 = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day + 1, 18, 30, 0);
            DateTime operdat = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day + 1, 09, 10, 0);

            sbSql.Clear();
            sbSql.Append(" INSERT INTO [HRMDB].[dbo].[AttendanceRollcall] ([AttendanceRollcallId],[EmployeeId],[Date],[BeginTime],[EndTime],[AttendanceRankId],[AttendanceTypeId],[Hours],[QuartersHours],[QuartersHoursUnit],[IsConfirm],[OperationDate],[UserId],[Recover],[Remark],[CreateDate],[LastModifiedDate],[CreateBy],[LastModifiedBy],[CorporationId],[Flag],[AssignReason],[OwnerId],[VirObjectId],[ActualBeginTime],[ActualEndTime],[Count],[DailyCards],[EmpRankCards],[CollectBegin],[CollectEnd],[IsAbnormal]) ");
            sbSql.AppendFormat(" SELECT  '{0}' AS [AttendanceRollcallId]", Guid.NewGuid());
            sbSql.Append(" , [Employee].EmployeeId AS [EmployeeId]");
            sbSql.AppendFormat(" ,'{0}' AS [Date]", workdt1.ToString("yyyy-MM-dd 00:00:00"));
            sbSql.AppendFormat(" ,'{0}' AS [BeginTime]", workdt1.ToString("yyyy-MM-dd HH:mm:ss"));
            sbSql.AppendFormat(" ,'{0}' AS [EndTime]", workdt2.ToString("yyyy-MM-dd HH:mm:ss"));
            sbSql.Append(" ,'13716c6bf390da08b46d1b85e1c1d24d986c9' AS[AttendanceRankId]");
            sbSql.Append(" ,'R13716c6bf390da08b46d1b85e1c1d24d986c9' AS[AttendanceTypeId]");
            sbSql.Append(" ,CASE datepart(weekday, getdate()) WHEN '1' THEN 0  WHEN '7' THEN 0 ELSE 480 END  AS [Hours]");
            sbSql.Append(" ,CASE datepart(weekday, getdate()) WHEN '1' THEN 0  WHEN '7' THEN 0 ELSE 480 END  AS [QuartersHours]");
            sbSql.Append(" ,'AttendanceUnit_003' AS [QuartersHoursUnit]");
            sbSql.Append(" ,'0' AS [IsConfirm]");
            sbSql.AppendFormat(" ,'{0}' AS [OperationDate]", operdat.ToString("yyyy-MM-dd HH:mm:ss"));
            sbSql.Append(" ,'FC0D07EA-E0FD-4BCF-B127-22F83D63D834' AS [UserId]");
            sbSql.Append(" ,'1' AS [Recover]");
            sbSql.Append(" ,NULL AS [Remark]");
            sbSql.AppendFormat(" ,'{0}' AS [CreateDate]", operdat.ToString("yyyy-MM-dd HH:mm:ss")); ;
            sbSql.AppendFormat(" ,'{0}' AS [LastModifiedDate]", operdat.ToString("yyyy-MM-dd HH:mm:ss"));
            sbSql.Append(" ,'FC0D07EA-E0FD-4BCF-B127-22F83D63D834' AS [CreateBy]");
            sbSql.Append(" ,'FC0D07EA-E0FD-4BCF-B127-22F83D63D834' AS [LastModifiedBy]");
            sbSql.Append(" ,NULL AS [CorporationId]");
            sbSql.Append(" ,'1' AS [Flag]");
            sbSql.Append(" ,NULL AS [AssignReason]");
            sbSql.Append(" ,'c9aec8c3-8889-4a8b-b718-9ce89aa84b22' AS [OwnerId]");
            sbSql.Append(" ,NULL AS [VirObjectId]");
            sbSql.Append(" ,NULL AS [ActualBeginTime]");
            sbSql.Append(" ,NULL AS [ActualEndTime]");
            sbSql.Append(" ,'1' AS [Count]");
            sbSql.Append(" ,CASE datepart(weekday, getdate()) WHEN '1' THEN NULL  WHEN '7' THEN NULL ELSE ' 08:20| 18:30' END  AS [DailyCards]");
            sbSql.Append(" ,CASE datepart(weekday, getdate()) WHEN '1' THEN NULL  WHEN '7' THEN NULL ELSE ' 08:20| 18:30' END AS [EmpRankCards]");
            sbSql.AppendFormat(" ,'{0}' AS [CollectBegin]",carddt1.ToString("yyyy-MM-dd HH:mm:ss"));
            sbSql.AppendFormat(" ,'{0}' AS [CollectEnd]", carddt2.ToString("yyyy-MM-dd HH:mm:ss"));
            sbSql.Append(" ,'0' AS [IsAbnormal]");
            sbSql.Append(" FROM [HRMDB].[dbo].[Employee]");
            sbSql.Append(" WHERE  [Employee].EmployeeId='C9AEC8C3-8889-4A8B-B718-9CE89AA84B22'");
            sbSql.Append(" ");
        }
        #endregion

        #region BUTTON

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            SETbutton1();
        }
    }
}
