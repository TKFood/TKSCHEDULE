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
            timer1.Interval = 1000*60;
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
            string hh = "8";
            string mm = "10";
            if (HRAUTO.Equals("Y")&& DateTime.Now.Hour.ToString().Equals(hh)&&DateTime.Now.Minute.ToString().Equals(mm))
            {                
                ADDHRCARD();
            }
                
            
        }
        public void ADDHRCARD()
        {
            DateTime workdt1 = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 8, 30, 0);
            DateTime workdt2 = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day , 17, 30, 0);
            DateTime carddt1 = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day , 8, 20, 0);
            DateTime carddt2 = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day , 18, 30, 0);
            DateTime operdat = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day , 09, 10, 0);
            DateTime operdat2 = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day , 09, 10, 0);

            

            try
            {

                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

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
                sbSql.AppendFormat(" ,'{0}' AS [CollectBegin]", carddt1.ToString("yyyy-MM-dd HH:mm:ss"));
                sbSql.AppendFormat(" ,'{0}' AS [CollectEnd]", carddt2.ToString("yyyy-MM-dd HH:mm:ss"));
                sbSql.Append(" ,'0' AS [IsAbnormal]");
                sbSql.Append(" FROM [HRMDB].[dbo].[Employee]");
                sbSql.Append(" WHERE  [Employee].EmployeeId='C9AEC8C3-8889-4A8B-B718-9CE89AA84B22'");
                sbSql.Append(" ");

                if (DateTime.Now.DayOfWeek.ToString("d").Equals("1") || DateTime.Now.DayOfWeek.ToString("d").Equals("1") || DateTime.Now.DayOfWeek.ToString("d").Equals("2") || DateTime.Now.DayOfWeek.ToString("d").Equals("3") || DateTime.Now.DayOfWeek.ToString("d").Equals("4") || DateTime.Now.DayOfWeek.ToString("d").Equals("5"))
                {
                    sbSql.Append(" INSERT INTO [HRMDB].[dbo].[AttendanceCollect] ([AttendanceCollectId],[MachineId],[MachineCode],[CardId],[CardCode],[EmployeeName],[EmployeeCode],[EmployeeId],[DepartmentName],[DepartmentId],[CostCenterId],[CostCenterCode],[Date],[Time],[IsManual],[Source],[IsUnusual],[UnusualTypeId],[Remark],[CreateDate],[LastModifiedDate],[CreateBy],[LastModifiedBy],[CorporationId],[Flag],[RepairId],[AttendanceTypeId],[OldLogIds],[AttendanceCollectLogId],[AssignReason],[OwnerId],[IsEss],[IsEF],[EssNo],[EssType],[ClassCode],[SubmitOperationDate],[SubmitUserId],[ConfirmOperationDate],[ConfirmUserId],[ApproveResultId],[FoundOperationDate],[FoundUserId],[ApproveDate],[ApproveEmployeeId],[ApproveEmployeeName],[ApproveRemark],[ApproveOperationDate],[ApproveUserId],[RepealOperationDate],[RepealUserId],[StateId],[IsFromEss],[IsForAttendance] )");
                    sbSql.AppendFormat(" SELECT '{0}' AS [AttendanceCollectId]", Guid.NewGuid());
                    sbSql.Append(" ,'CEC9EFA6-21E1-4222-A013-9E3F7D15B936' AS [MachineId]");
                    sbSql.Append(" ,NULL AS [MachineCode]");
                    sbSql.Append(" ,[CardId] AS [CardId]");
                    sbSql.Append(" ,[CardNo] AS [CardCode]");
                    sbSql.Append(" ,CnName AS [EmployeeName]");
                    sbSql.Append(" ,[Employee].[Code] AS [EmployeeCode]");
                    sbSql.Append(" ,[Employee].[EmployeeId] AS [EmployeeId]");
                    sbSql.Append(" ,[Department].[Name] AS [DepartmentName]");
                    sbSql.Append(" ,[Department].DepartmentId AS [DepartmentId]");
                    sbSql.Append(" ,'00000000-0000-0000-0000-000000000000' AS [CostCenterId]");
                    sbSql.Append(" ,NULL AS [CostCenterCode]");
                    sbSql.AppendFormat(" ,'{0}' AS [Date]", carddt1.ToString("yyyy-MM-dd HH:mm:ss"));
                    sbSql.AppendFormat(" ,'{0}' AS [Time]", carddt1.ToString("HH:mm"));
                    sbSql.Append(" ,'0' AS [IsManual]");
                    sbSql.AppendFormat(" ,'{0} '+[CardNo]+' 03' AS [Source]", carddt1.ToString("yyyy/MM/dd HH:mm"));
                    sbSql.Append(" ,'0' AS [IsUnusual]");
                    sbSql.Append(" ,NULL AS [UnusualTypeId]");
                    sbSql.Append(" ,NULL AS [Remark]");
                    sbSql.AppendFormat(" ,'{0}' AS [CreateDate]", carddt1.ToString("yyyy-MM-dd 09:10:00"));
                    sbSql.AppendFormat(" ,'{0}' AS [LastModifiedDate]", carddt1.ToString("yyyy-MM-dd  09:10:00"));
                    sbSql.Append(" ,'98385A19-5BA6-43E5-BD0A-6A727F2E9C35' AS [CreateBy]");
                    sbSql.Append(" ,'98385A19-5BA6-43E5-BD0A-6A727F2E9C35' AS [LastModifiedBy]");
                    sbSql.Append(" ,NULL AS [CorporationId]");
                    sbSql.Append(" ,'1' AS [Flag]");
                    sbSql.Append(" ,NULL AS [RepairId]");
                    sbSql.Append(" ,NULL AS [AttendanceTypeId]");
                    sbSql.Append(" ,NULL AS [OldLogIds]");
                    sbSql.Append(" ,'D6A48B3D-FE56-4556-B3B7-1A3FF26048C8' AS [AttendanceCollectLogId]");
                    sbSql.Append(" ,NULL AS [AssignReason]");
                    sbSql.Append(" ,'c9aec8c3-8889-4a8b-b718-9ce89aa84b22' AS [OwnerId]");
                    sbSql.Append(" ,'0' AS [IsEss]");
                    sbSql.Append(" ,NULL AS [IsEF]");
                    sbSql.Append(" ,NULL AS [EssNo]");
                    sbSql.Append(" ,NULL AS [EssType]");
                    sbSql.Append(" ,NULL AS [ClassCode]");
                    sbSql.Append(" ,NULL AS [SubmitOperationDate]");
                    sbSql.Append(" ,NULL AS [SubmitUserId]");
                    sbSql.Append(" ,NULL AS [ConfirmOperationDate]");
                    sbSql.Append(" ,NULL AS [ConfirmUserId]");
                    sbSql.Append(" ,'OperatorResult_001' AS [ApproveResultId]");
                    sbSql.Append(" ,NULL AS [FoundOperationDate]");
                    sbSql.Append(" ,NULL AS [FoundUserId]");
                    sbSql.AppendFormat(" ,'{0}' AS [ApproveDate]", carddt1.ToString("yyyy-MM-dd  09:10:00")); ;
                    sbSql.Append(" ,'9EF0F7A6-F2F0-491B-A730-CD66CAE430D4' AS [ApproveEmployeeId]");
                    sbSql.Append(" ,'蘇慧茹' AS [ApproveEmployeeName]");
                    sbSql.Append(" ,NULL AS [ApproveRemark]");
                    sbSql.AppendFormat(" ,'{0}' AS [ApproveOperationDate]", carddt1.ToString("yyyy-MM-dd  09:10:00"));
                    sbSql.Append(" ,'FC0D07EA-E0FD-4BCF-B127-22F83D63D834' AS [ApproveUserId]");
                    sbSql.Append(" ,NULL AS [RepealOperationDate] ");
                    sbSql.Append(" ,NULL AS [RepealUserId]");
                    sbSql.Append(" ,'PlanState_003' AS [StateId]");
                    sbSql.Append(" ,'0' AS [IsFromEss]");
                    sbSql.Append(" ,'1' AS [IsForAttendance]");
                    sbSql.Append(" FROM [HRMDB].[dbo].[Employee],[HRMDB].[dbo].[Card],[HRMDB].[dbo].[Department]");
                    sbSql.Append(" WHERE [Employee].EMPLOYEEID=[Card] .EMPLOYEEID AND [Employee].DepartmentId=[Department].DepartmentId");
                    sbSql.Append(" AND ([Employee].EMPLOYEEID='C9AEC8C3-8889-4A8B-B718-9CE89AA84B22' AND CardId='F70447D0-4553-4A8C-86DE-5D1B179B452E')");
                    sbSql.Append(" ");

                    sbSql.Append(" INSERT INTO [HRMDB].[dbo].[AttendanceCollect] ([AttendanceCollectId],[MachineId],[MachineCode],[CardId],[CardCode],[EmployeeName],[EmployeeCode],[EmployeeId],[DepartmentName],[DepartmentId],[CostCenterId],[CostCenterCode],[Date],[Time],[IsManual],[Source],[IsUnusual],[UnusualTypeId],[Remark],[CreateDate],[LastModifiedDate],[CreateBy],[LastModifiedBy],[CorporationId],[Flag],[RepairId],[AttendanceTypeId],[OldLogIds],[AttendanceCollectLogId],[AssignReason],[OwnerId],[IsEss],[IsEF],[EssNo],[EssType],[ClassCode],[SubmitOperationDate],[SubmitUserId],[ConfirmOperationDate],[ConfirmUserId],[ApproveResultId],[FoundOperationDate],[FoundUserId],[ApproveDate],[ApproveEmployeeId],[ApproveEmployeeName],[ApproveRemark],[ApproveOperationDate],[ApproveUserId],[RepealOperationDate],[RepealUserId],[StateId],[IsFromEss],[IsForAttendance] )");
                    sbSql.AppendFormat(" SELECT '{0}' AS [AttendanceCollectId]", Guid.NewGuid());
                    sbSql.Append(" ,'CEC9EFA6-21E1-4222-A013-9E3F7D15B936' AS [MachineId]");
                    sbSql.Append(" ,NULL AS [MachineCode]");
                    sbSql.Append(" ,[CardId] AS [CardId]");
                    sbSql.Append(" ,[CardNo] AS [CardCode]");
                    sbSql.Append(" ,CnName AS [EmployeeName]");
                    sbSql.Append(" ,[Employee].[Code] AS [EmployeeCode]");
                    sbSql.Append(" ,[Employee].[EmployeeId] AS [EmployeeId]");
                    sbSql.Append(" ,[Department].[Name] AS [DepartmentName]");
                    sbSql.Append(" ,[Department].DepartmentId AS [DepartmentId]");
                    sbSql.Append(" ,'00000000-0000-0000-0000-000000000000' AS [CostCenterId]");
                    sbSql.Append(" ,NULL AS [CostCenterCode]");
                    sbSql.AppendFormat(" ,'{0}' AS [Date]", carddt2.ToString("yyyy-MM-dd HH:mm:ss"));
                    sbSql.AppendFormat(" ,'{0}' AS [Time]", carddt2.ToString("HH:mm"));
                    sbSql.Append(" ,'0' AS [IsManual]");
                    sbSql.AppendFormat(" ,'{0} '+[CardNo]+' 03' AS [Source]", carddt2.ToString("yyyy/MM/dd  HH:mm"));
                    sbSql.Append(" ,'0' AS [IsUnusual]");
                    sbSql.Append(" ,NULL AS [UnusualTypeId]");
                    sbSql.Append(" ,NULL AS [Remark]");
                    sbSql.AppendFormat(" ,'{0}' AS [CreateDate]", operdat2.ToString("yyyy-MM-dd 09:10:00"));
                    sbSql.AppendFormat(" ,'{0}' AS [LastModifiedDate]", operdat2.ToString("yyyy-MM-dd  09:10:00"));
                    sbSql.Append(" ,'98385A19-5BA6-43E5-BD0A-6A727F2E9C35' AS [CreateBy]");
                    sbSql.Append(" ,'98385A19-5BA6-43E5-BD0A-6A727F2E9C35' AS [LastModifiedBy]");
                    sbSql.Append(" ,NULL AS [CorporationId]");
                    sbSql.Append(" ,'1' AS [Flag]");
                    sbSql.Append(" ,NULL AS [RepairId]");
                    sbSql.Append(" ,NULL AS [AttendanceTypeId]");
                    sbSql.Append(" ,NULL AS [OldLogIds]");
                    sbSql.Append(" ,'D6A48B3D-FE56-4556-B3B7-1A3FF26048C8' AS [AttendanceCollectLogId]");
                    sbSql.Append(" ,NULL AS [AssignReason]");
                    sbSql.Append(" ,'c9aec8c3-8889-4a8b-b718-9ce89aa84b22' AS [OwnerId]");
                    sbSql.Append(" ,'0' AS [IsEss]");
                    sbSql.Append(" ,NULL AS [IsEF]");
                    sbSql.Append(" ,NULL AS [EssNo]");
                    sbSql.Append(" ,NULL AS [EssType]");
                    sbSql.Append(" ,NULL AS [ClassCode]");
                    sbSql.Append(" ,NULL AS [SubmitOperationDate]");
                    sbSql.Append(" ,NULL AS [SubmitUserId]");
                    sbSql.Append(" ,NULL AS [ConfirmOperationDate]");
                    sbSql.Append(" ,NULL AS [ConfirmUserId]");
                    sbSql.Append(" ,'OperatorResult_001' AS [ApproveResultId]");
                    sbSql.Append(" ,NULL AS [FoundOperationDate]");
                    sbSql.Append(" ,NULL AS [FoundUserId]");
                    sbSql.AppendFormat(" ,'{0}' AS [ApproveDate]", operdat2.ToString("yyyy-MM-dd  09:10:00")); ;
                    sbSql.Append(" ,'9EF0F7A6-F2F0-491B-A730-CD66CAE430D4' AS [ApproveEmployeeId]");
                    sbSql.Append(" ,'蘇慧茹' AS [ApproveEmployeeName]");
                    sbSql.Append(" ,NULL AS [ApproveRemark]");
                    sbSql.AppendFormat(" ,'{0}' AS [ApproveOperationDate]", operdat2.ToString("yyyy-MM-dd  09:10:00"));
                    sbSql.Append(" ,'FC0D07EA-E0FD-4BCF-B127-22F83D63D834' AS [ApproveUserId]");
                    sbSql.Append(" ,NULL AS [RepealOperationDate] ");
                    sbSql.Append(" ,NULL AS [RepealUserId]");
                    sbSql.Append(" ,'PlanState_003' AS [StateId]");
                    sbSql.Append(" ,'0' AS [IsFromEss]");
                    sbSql.Append(" ,'1' AS [IsForAttendance]");
                    sbSql.Append(" FROM [HRMDB].[dbo].[Employee],[HRMDB].[dbo].[Card],[HRMDB].[dbo].[Department]");
                    sbSql.Append(" WHERE [Employee].EMPLOYEEID=[Card] .EMPLOYEEID AND [Employee].DepartmentId=[Department].DepartmentId");
                    sbSql.Append(" AND ([Employee].EMPLOYEEID='C9AEC8C3-8889-4A8B-B718-9CE89AA84B22' AND CardId='F70447D0-4553-4A8C-86DE-5D1B179B452E')");
                    sbSql.Append(" ");
                }

               
                //sbSql.AppendFormat("  UPDATE Member SET Cname='{1}',Mobile1='{2}' WHERE ID='{0}' ", list_Member[0].ID.ToString(), list_Member[0].Cname.ToString(), list_Member[0].Mobile1.ToString());

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                    label3.Text = DateTime.Now.ToString("yyyymmdd") + " FAIL";
                }
                else
                {
                    tran.Commit();      //執行交易
                    label3.Text = DateTime.Now.ToString("yyyymmdd") + " DONE";

                }

                sqlConn.Close();

            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
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
