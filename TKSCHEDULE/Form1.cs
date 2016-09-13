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
        StringBuilder sbSqlEXE = new StringBuilder();
        StringBuilder sbSqlEXE2 = new StringBuilder();
        StringBuilder sbSqlEXE3 = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlTransaction tran2;
        SqlCommand cmd = new SqlCommand();
        SqlCommand cmd2 = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();
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
                UPDATECARD();
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
                sbSqlEXE.Clear();
                sbSql.Clear();

                sbSql.Append(" SELECT [Employee].[EmployeeId],[AttendanceRank].[Name]");
                sbSql.Append(" FROM [HRMDB].[dbo].[Employee],[HRMDB].[dbo].[AttendanceEmpRank],[HRMDB].[dbo].[AttendanceRank]");
                sbSql.Append(" WHERE [Employee].[EmployeeId]=[AttendanceEmpRank].[EmployeeId]");
                sbSql.Append(" AND [AttendanceEmpRank].[AttendanceRankId]=[AttendanceRank].[AttendanceRankId] ");
                sbSql.Append(" AND CONVERT(varchar(100),[AttendanceEmpRank].[Date],112)=CONVERT(varchar(100),GETDATE(),112)");
                sbSql.Append(" AND [Employee].[EmployeeId] IN ('C9AEC8C3-8889-4A8B-B718-9CE89AA84B22' ,'6FBF39F6-4666-4941-9FAF-A9CBBC8B1E0B')");
                sbSql.Append(" ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds");
                sqlConn.Close();

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSqlEXE.Clear();

                if (ds.Tables["TEMPds"].Rows.Count >=1)
                {
                    foreach (DataRow od in ds.Tables["TEMPds"].Rows)
                    {
                        if(!od["Name"].ToString().Contains("休息"))
                        {
                            sbSqlEXE.Append(" ");
                            sbSqlEXE.Append(" INSERT INTO [HRMDB].[dbo].[AttendanceRollcall] ([AttendanceRollcallId],[EmployeeId],[Date],[BeginTime],[EndTime],[AttendanceRankId],[AttendanceTypeId],[Hours],[QuartersHours],[QuartersHoursUnit],[IsConfirm],[OperationDate],[UserId],[Recover],[Remark],[CreateDate],[LastModifiedDate],[CreateBy],[LastModifiedBy],[CorporationId],[Flag],[AssignReason],[OwnerId],[VirObjectId],[ActualBeginTime],[ActualEndTime],[Count],[DailyCards],[EmpRankCards],[CollectBegin],[CollectEnd],[IsAbnormal]) ");
                            sbSqlEXE.AppendFormat(" SELECT TOP 1 '{0}' AS [AttendanceRollcallId]", Guid.NewGuid());
                            sbSqlEXE.Append(" ,[EmployeeId]");
                            sbSqlEXE.Append(" ,CONVERT(varchar(100),GETDATE(),23)+' 00:00:00.000' AS [Date]");
                            sbSqlEXE.Append(" ,CONVERT(varchar(100),GETDATE(),23)+' '+CONVERT(varchar(100),[AttendanceRollcall].[BeginTime],114) AS [BeginTime] ");
                            sbSqlEXE.Append(" ,CONVERT(varchar(100),GETDATE(),23)+' '+CONVERT(varchar(100),[AttendanceRollcall].[EndTime],114) AS [EndTime]");
                            sbSqlEXE.Append(" ,[AttendanceRankId]");
                            sbSqlEXE.Append(" ,[AttendanceTypeId]");
                            sbSqlEXE.Append(" ,[Hours]");
                            sbSqlEXE.Append(" ,[QuartersHours]");
                            sbSqlEXE.Append(" ,[QuartersHoursUnit]");
                            sbSqlEXE.Append(" ,[IsConfirm]");
                            sbSqlEXE.Append(" ,CONVERT(varchar(100),GETDATE(),23)+' 09:30:00.000' AS [OperationDate]");
                            sbSqlEXE.Append(" ,[UserId]");
                            sbSqlEXE.Append(" ,[Recover]");
                            sbSqlEXE.Append(" ,[Remark]");
                            sbSqlEXE.Append(" ,CONVERT(varchar(100),GETDATE(),23)+' 09:30:00.000' AS [CreateDate]");
                            sbSqlEXE.Append(" ,CONVERT(varchar(100),GETDATE(),23)+' 09:30:00.000' AS [LastModifiedDate]");
                            sbSqlEXE.Append(" ,[CreateBy]");
                            sbSqlEXE.Append(" ,[LastModifiedBy]");
                            sbSqlEXE.Append(" ,[CorporationId]");
                            sbSqlEXE.Append(" ,[Flag]");
                            sbSqlEXE.Append(" ,[AssignReason]");
                            sbSqlEXE.Append(" ,[OwnerId]");
                            sbSqlEXE.Append(" ,[VirObjectId]");
                            sbSqlEXE.Append(" ,[ActualBeginTime]");
                            sbSqlEXE.Append(" ,[ActualEndTime]");
                            sbSqlEXE.Append(" ,[Count]");
                            sbSqlEXE.Append(" ,[DailyCards]");
                            sbSqlEXE.Append(" ,[EmpRankCards]");
                            sbSqlEXE.Append(" ,CONVERT(varchar(100),GETDATE(),23)+' '+CONVERT(varchar(100),[AttendanceRollcall].[CollectBegin],114) AS [CollectBegin]");
                            sbSqlEXE.Append(" ,CONVERT(varchar(100),GETDATE(),23)+' '+CONVERT(varchar(100),[AttendanceRollcall].[CollectEnd],114) AS [CollectEnd]");
                            sbSqlEXE.Append(" ,[IsAbnormal]");
                            sbSqlEXE.Append(" FROM [HRMDB].[dbo].[AttendanceRollcall] WITH (NOLOCK)");
                            sbSqlEXE.AppendFormat(" WHERE [Hours]>0 AND [EmployeeId]='{0}'", od["EmployeeId"].ToString());
                            sbSqlEXE.Append(" ORDER BY [AttendanceRollcall].[Date] DESC ");
                            sbSqlEXE.Append(" ");
                            sbSqlEXE.Append(" INSERT INTO [HRMDB].[dbo].[AttendanceCollect] ([AttendanceCollectId],[MachineId],[MachineCode],[CardId],[CardCode],[EmployeeName],[EmployeeCode],[EmployeeId],[DepartmentName],[DepartmentId],[CostCenterId],[CostCenterCode],[Date],[Time],[IsManual],[Source],[IsUnusual],[UnusualTypeId],[Remark],[CreateDate],[LastModifiedDate],[CreateBy],[LastModifiedBy],[CorporationId],[Flag],[RepairId],[AttendanceTypeId],[OldLogIds],[AttendanceCollectLogId],[AssignReason],[OwnerId],[IsEss],[IsEF],[EssNo],[EssType],[ClassCode],[SubmitOperationDate],[SubmitUserId],[ConfirmOperationDate],[ConfirmUserId],[ApproveResultId],[FoundOperationDate],[FoundUserId],[ApproveDate],[ApproveEmployeeId],[ApproveEmployeeName],[ApproveRemark],[ApproveOperationDate],[ApproveUserId],[RepealOperationDate],[RepealUserId],[StateId],[IsFromEss],[IsForAttendance] )");
                            sbSqlEXE.AppendFormat(" SELECT TOP 1  '{0}' AS [AttendanceCollectId]", Guid.NewGuid());
                            sbSqlEXE.Append(" ,[MachineId],[MachineCode],[CardId],[CardCode],[EmployeeName],[EmployeeCode],[EmployeeId],[DepartmentName],[DepartmentId],[CostCenterId],[CostCenterCode]");
                            sbSqlEXE.Append(" ,CONVERT(varchar(100),GETDATE(),23)+' '+CONVERT(varchar(100),[AttendanceCollect].[Date],114) AS [Date]");
                            sbSqlEXE.Append(" ,[Time],[IsManual]");
                            sbSqlEXE.Append(" ,CONVERT(varchar(100),GETDATE(),23)+' '+CONVERT(varchar(100),[AttendanceCollect].[Date],8)+' 000459 03'  AS [Source]");
                            sbSqlEXE.Append(" ,[IsUnusual],[UnusualTypeId],[Remark]");
                            sbSqlEXE.Append(" ,CONVERT(varchar(100),GETDATE(),23)+' '+CONVERT(varchar(100),[AttendanceCollect].[CreateDate],114) AS [CreateDate]");
                            sbSqlEXE.Append(" ,CONVERT(varchar(100),GETDATE(),23)+' '+CONVERT(varchar(100),[AttendanceCollect].[LastModifiedDate],114) AS [LastModifiedDate]");
                            sbSqlEXE.Append(" ,[CreateBy],[LastModifiedBy],[CorporationId],[Flag],[RepairId],[AttendanceTypeId],[OldLogIds],[AttendanceCollectLogId],[AssignReason],[OwnerId],[IsEss],[IsEF],[EssNo],[EssType],[ClassCode],[SubmitOperationDate],[SubmitUserId],[ConfirmOperationDate],[ConfirmUserId],[ApproveResultId],[FoundOperationDate],[FoundUserId]");
                            sbSqlEXE.Append(" ,CONVERT(varchar(100),GETDATE(),23)+' '+CONVERT(varchar(100),[AttendanceCollect].[ApproveDate],114) AS [ApproveDate]");
                            sbSqlEXE.Append(" ,[ApproveEmployeeId],[ApproveEmployeeName],[ApproveRemark],[ApproveOperationDate],[ApproveUserId]");
                            sbSqlEXE.Append(" ,CONVERT(varchar(100),GETDATE(),23)+' '+CONVERT(varchar(100),[AttendanceCollect].[RepealOperationDate],114) AS [RepealOperationDate]");
                            sbSqlEXE.Append(" ,[RepealUserId],[StateId],[IsFromEss],[IsForAttendance]");
                            sbSqlEXE.Append(" FROM  [HRMDB].[dbo].[AttendanceCollect] WITH (NOLOCK)");
                            sbSqlEXE.Append(" WHERE CONVERT(varchar(100),[AttendanceCollect].[Date],114) >='08:00:00' AND CONVERT(varchar(100),[AttendanceCollect].[Date],114) <='09:00:00'");
                            sbSqlEXE.AppendFormat(" AND  [EmployeeId]='{0}'", od["EmployeeId"].ToString());
                            sbSqlEXE.Append(" ORDER BY [AttendanceCollect].[Date] DESC ");
                            sbSqlEXE.Append(" ");
                            sbSqlEXE.Append(" INSERT INTO [HRMDB].[dbo].[AttendanceCollect] ([AttendanceCollectId],[MachineId],[MachineCode],[CardId],[CardCode],[EmployeeName],[EmployeeCode],[EmployeeId],[DepartmentName],[DepartmentId],[CostCenterId],[CostCenterCode],[Date],[Time],[IsManual],[Source],[IsUnusual],[UnusualTypeId],[Remark],[CreateDate],[LastModifiedDate],[CreateBy],[LastModifiedBy],[CorporationId],[Flag],[RepairId],[AttendanceTypeId],[OldLogIds],[AttendanceCollectLogId],[AssignReason],[OwnerId],[IsEss],[IsEF],[EssNo],[EssType],[ClassCode],[SubmitOperationDate],[SubmitUserId],[ConfirmOperationDate],[ConfirmUserId],[ApproveResultId],[FoundOperationDate],[FoundUserId],[ApproveDate],[ApproveEmployeeId],[ApproveEmployeeName],[ApproveRemark],[ApproveOperationDate],[ApproveUserId],[RepealOperationDate],[RepealUserId],[StateId],[IsFromEss],[IsForAttendance] )");
                            sbSqlEXE.AppendFormat(" SELECT TOP 1 '{0}' AS  [AttendanceCollectId]", Guid.NewGuid());
                            sbSqlEXE.Append(" ,[MachineId],[MachineCode],[CardId],[CardCode],[EmployeeName],[EmployeeCode],[EmployeeId],[DepartmentName],[DepartmentId],[CostCenterId],[CostCenterCode]");
                            sbSqlEXE.Append(" ,CONVERT(varchar(100),GETDATE(),23)+' '+CONVERT(varchar(100),[AttendanceCollect].[Date],114) AS [Date]");
                            sbSqlEXE.Append(" ,[Time],[IsManual]");
                            sbSqlEXE.Append(" ,CONVERT(varchar(100),GETDATE(),23)+' '+CONVERT(varchar(100),[AttendanceCollect].[Date],8)+' 000459 03'  AS [Source]");
                            sbSqlEXE.Append(" ,[IsUnusual],[UnusualTypeId],[Remark]");
                            sbSqlEXE.Append(" ,CONVERT(varchar(100),GETDATE()+1,23)+' '+CONVERT(varchar(100),[AttendanceCollect].[CreateDate],114) AS [CreateDate]");
                            sbSqlEXE.Append(" ,CONVERT(varchar(100),GETDATE()+1,23)+' '+CONVERT(varchar(100),[AttendanceCollect].[LastModifiedDate],114) AS [LastModifiedDate]");
                            sbSqlEXE.Append(" ,[CreateBy],[LastModifiedBy],[CorporationId],[Flag],[RepairId],[AttendanceTypeId],[OldLogIds],[AttendanceCollectLogId],[AssignReason],[OwnerId],[IsEss],[IsEF],[EssNo],[EssType],[ClassCode],[SubmitOperationDate],[SubmitUserId],[ConfirmOperationDate],[ConfirmUserId],[ApproveResultId],[FoundOperationDate],[FoundUserId]");
                            sbSqlEXE.Append(" ,CONVERT(varchar(100),GETDATE()+1,23)+' '+CONVERT(varchar(100),[AttendanceCollect].[ApproveDate],114) AS [ApproveDate]");
                            sbSqlEXE.Append(" ,[ApproveEmployeeId],[ApproveEmployeeName],[ApproveRemark],[ApproveOperationDate],[ApproveUserId]");
                            sbSqlEXE.Append(" ,CONVERT(varchar(100),GETDATE()+1,23)+' '+CONVERT(varchar(100),[AttendanceCollect].[RepealOperationDate],114) AS [RepealOperationDate]");
                            sbSqlEXE.Append(" ,[RepealUserId],[StateId],[IsFromEss],[IsForAttendance]  ");
                            sbSqlEXE.Append(" FROM  [HRMDB].[dbo].[AttendanceCollect] WITH (NOLOCK)");
                            sbSqlEXE.Append(" WHERE CONVERT(varchar(100),[AttendanceCollect].[Date],114) >='17:00:00'");
                            sbSqlEXE.AppendFormat(" AND  [EmployeeId]='{0}'", od["EmployeeId"].ToString());
                            sbSqlEXE.Append(" ORDER BY [AttendanceCollect].[Date] DESC ");
                            sbSqlEXE.Append(" ");
                        }
                        else
                        {
                            sbSqlEXE.Append(" ");
                            sbSqlEXE.Append(" INSERT INTO [HRMDB].[dbo].[AttendanceRollcall] ([AttendanceRollcallId],[EmployeeId],[Date],[BeginTime],[EndTime],[AttendanceRankId],[AttendanceTypeId],[Hours],[QuartersHours],[QuartersHoursUnit],[IsConfirm],[OperationDate],[UserId],[Recover],[Remark],[CreateDate],[LastModifiedDate],[CreateBy],[LastModifiedBy],[CorporationId],[Flag],[AssignReason],[OwnerId],[VirObjectId],[ActualBeginTime],[ActualEndTime],[Count],[DailyCards],[EmpRankCards],[CollectBegin],[CollectEnd],[IsAbnormal]) ");
                            sbSqlEXE.AppendFormat(" SELECT TOP 1 '{0}' AS [AttendanceRollcallId]", Guid.NewGuid());
                            sbSqlEXE.Append(" ,[EmployeeId]");
                            sbSqlEXE.Append(" ,CONVERT(varchar(100),GETDATE(),23)+' 00:00:00.000' AS [Date]");
                            sbSqlEXE.Append(" ,CONVERT(varchar(100),GETDATE(),23)+' '+CONVERT(varchar(100),[AttendanceRollcall].[BeginTime],114) AS [BeginTime] ");
                            sbSqlEXE.Append(" ,CONVERT(varchar(100),GETDATE(),23)+' '+CONVERT(varchar(100),[AttendanceRollcall].[EndTime],114) AS [EndTime]");
                            sbSqlEXE.Append(" ,[AttendanceRankId]");
                            sbSqlEXE.Append(" ,[AttendanceTypeId]");
                            sbSqlEXE.Append(" ,[Hours]");
                            sbSqlEXE.Append(" ,[QuartersHours]");
                            sbSqlEXE.Append(" ,[QuartersHoursUnit]");
                            sbSqlEXE.Append(" ,[IsConfirm]");
                            sbSqlEXE.Append(" ,CONVERT(varchar(100),GETDATE(),23)+' 09:30:00.000' AS [OperationDate]");
                            sbSqlEXE.Append(" ,[UserId]");
                            sbSqlEXE.Append(" ,[Recover]");
                            sbSqlEXE.Append(" ,[Remark]");
                            sbSqlEXE.Append(" ,CONVERT(varchar(100),GETDATE(),23)+' 09:30:00.000' AS [CreateDate]");
                            sbSqlEXE.Append(" ,CONVERT(varchar(100),GETDATE(),23)+' 09:30:00.000' AS [LastModifiedDate]");
                            sbSqlEXE.Append(" ,[CreateBy]");
                            sbSqlEXE.Append(" ,[LastModifiedBy]");
                            sbSqlEXE.Append(" ,[CorporationId]");
                            sbSqlEXE.Append(" ,[Flag]");
                            sbSqlEXE.Append(" ,[AssignReason]");
                            sbSqlEXE.Append(" ,[OwnerId]");
                            sbSqlEXE.Append(" ,[VirObjectId]");
                            sbSqlEXE.Append(" ,[ActualBeginTime]");
                            sbSqlEXE.Append(" ,[ActualEndTime]");
                            sbSqlEXE.Append(" ,[Count]");
                            sbSqlEXE.Append(" ,[DailyCards]");
                            sbSqlEXE.Append(" ,[EmpRankCards]");
                            sbSqlEXE.Append(" ,CONVERT(varchar(100),GETDATE(),23)+' '+CONVERT(varchar(100),[AttendanceRollcall].[CollectBegin],114) AS [CollectBegin]");
                            sbSqlEXE.Append(" ,CONVERT(varchar(100),GETDATE(),23)+' '+CONVERT(varchar(100),[AttendanceRollcall].[CollectEnd],114) AS [CollectEnd]");
                            sbSqlEXE.Append(" ,[IsAbnormal]");
                            sbSqlEXE.Append(" FROM [HRMDB].[dbo].[AttendanceRollcall] WITH (NOLOCK)");
                            sbSqlEXE.AppendFormat(" WHERE [Hours]=0 AND [EmployeeId]='{0}'", od["EmployeeId"].ToString());
                            sbSqlEXE.Append(" ORDER BY [AttendanceRollcall].[Date] DESC ");
                            sbSqlEXE.Append(" ");
                            
                        }
                        
                    }
                        
                }


                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSqlEXE.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                    label3.Text = DateTime.Now.ToString("yyyy/MM/dd") + "ADD FAIL";
                }
                else
                {
                    tran.Commit();      //執行交易
                    label3.Text = DateTime.Now.ToString("yyyy/MM/dd") + "ADD DONE";

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

        public void UPDATECARD()
        {
            try
            {

                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);
                sqlConn2 = new SqlConnection(connectionString);

                sqlConn.Close();
                sbSqlEXE2.Clear();
                sbSql.Clear();

                sbSql.Append(" SELECT [EmployeeId] FROM [HRMDB].[dbo].[Employee] WITH (NOLOCK) WHERE [CODE] IN ('160115')");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                sqlConn.Open();
                ds2.Clear();
                adapter.Fill(ds2, "TEMPds2");
                sqlConn.Close();
               
                sqlConn.Close();
                sqlConn.Open();
                tran2 = sqlConn.BeginTransaction();    

                if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                {
                    foreach (DataRow od in ds2.Tables["TEMPds2"].Rows)
                    {
                        sqlConn.Close();
                        sbSql.Clear();
                        sbSql.Append(" SELECT [AttendanceCollect].[AttendanceCollectId],*");
                        sbSql.Append(" FROM [HRMDB].[dbo].[AttendanceCollect] WITH (NOLOCK),[HRMDB].[dbo].[AttendanceLeaveInfo] WITH (NOLOCK)");
                        sbSql.Append(" WHERE CONVERT(varchar(100),[AttendanceLeaveInfo].[BeginDate],112)<=CONVERT(varchar(100),CONVERT(varchar(100),[AttendanceCollect].[Date],112),112)");
                        sbSql.Append(" AND CONVERT(varchar(100),[AttendanceLeaveInfo].[EndDate],112)>=CONVERT(varchar(100),CONVERT(varchar(100),[AttendanceCollect].[Date],112),112)");
                        sbSql.Append(" AND [AttendanceCollect].[EmployeeId]=[AttendanceLeaveInfo].[EmployeeId]");
                        sbSql.Append(" AND [AttendanceLeaveInfo].[Hours]=8");
                        sbSql.Append(" AND ISNULL([AttendanceCollect].[Date],'')<>''");
                        sbSql.AppendFormat(" AND [AttendanceCollect].[EmployeeId]='{0}' ", od["EmployeeId"].ToString());
                        sbSql.Append(" ");
                        adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
                        sqlCmdBuilder = new SqlCommandBuilder(adapter);

                        sqlConn.Open();
                        ds3.Clear();
                        adapter.Fill(ds3, "TEMPds3");
                        sqlConn.Close();

                        sqlConn.Close();
                        sbSql.Clear();
                        sbSql.Append(" SELECT [AttendanceRollcall].[AttendanceRollcallId],*");
                        sbSql.Append(" FROM [HRMDB].[dbo].[AttendanceRollcall] WITH (NOLOCK),[HRMDB].[dbo].[AttendanceLeaveInfo] WITH (NOLOCK)");
                        sbSql.Append(" WHERE CONVERT(varchar(100),[AttendanceLeaveInfo].[BeginDate],112)<=CONVERT(varchar(100),CONVERT(varchar(100),[AttendanceRollcall].[Date],112),112)");
                        sbSql.Append(" AND CONVERT(varchar(100),[AttendanceLeaveInfo].[EndDate],112)>=CONVERT(varchar(100),CONVERT(varchar(100),[AttendanceRollcall].[Date],112),112)");
                        sbSql.Append(" AND [AttendanceRollcall].[EmployeeId]=[AttendanceLeaveInfo].[EmployeeId]");
                        sbSql.Append(" AND [AttendanceLeaveInfo].[Hours]=8");
                        sbSql.Append(" AND [AttendanceRollcall].[Hours]<>0");
                        sbSql.AppendFormat(" AND [AttendanceRollcall].[EmployeeId]='{0}' ", od["EmployeeId"].ToString());
                        sbSql.Append(" ");
                        adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
                        sqlCmdBuilder = new SqlCommandBuilder(adapter);

                        sqlConn.Open();
                        ds4.Clear();
                        adapter.Fill(ds4, "TEMPds4");
                        sqlConn.Close();

                        if (ds3.Tables["TEMPds3"].Rows.Count >= 1)
                        {
                            sbSqlEXE2.Append(" UPDATE [HRMDB].[dbo].[AttendanceCollect]");
                            sbSqlEXE2.Append(" SET [Date]=NULL,[Time]=NULL,[Source]=NULL,[CreateDate]=NULL,[LastModifiedDate]=NULL,[ApproveDate]=NULL,[ApproveOperationDate]=NULL");
                            sbSqlEXE2.Append(" WHERE [AttendanceCollect].[AttendanceCollectId] IN ");
                            sbSqlEXE2.Append(" (SELECT [AttendanceCollect].[AttendanceCollectId] FROM [HRMDB].[dbo].[AttendanceCollect] WITH (NOLOCK),[HRMDB].[dbo].[AttendanceLeaveInfo] WITH (NOLOCK)");
                            sbSqlEXE2.Append(" WHERE CONVERT(varchar(100),[AttendanceLeaveInfo].[BeginDate],112)<=CONVERT(varchar(100),CONVERT(varchar(100),[AttendanceCollect].[Date],112),112)");
                            sbSqlEXE2.Append(" AND CONVERT(varchar(100),[AttendanceLeaveInfo].[EndDate],112)>=CONVERT(varchar(100),CONVERT(varchar(100),[AttendanceCollect].[Date],112),112)");
                            sbSqlEXE2.Append(" AND [AttendanceCollect].[EmployeeId]=[AttendanceLeaveInfo].[EmployeeId]");
                            sbSqlEXE2.Append(" AND [AttendanceLeaveInfo].[Hours]=8");
                            sbSqlEXE2.Append(" AND ISNULL([AttendanceCollect].[Date],'')<>''");
                            sbSqlEXE2.AppendFormat(" AND [AttendanceCollect].[EmployeeId]='{0}' )", od["EmployeeId"].ToString());
                            sbSqlEXE2.Append(" ");
                        }
                        if (ds4.Tables["TEMPds4"].Rows.Count >= 1)
                        {
                            sbSqlEXE2.Append(" UPDATE [HRMDB].[dbo].[AttendanceRollcall]");
                            sbSqlEXE2.Append(" SET [Hours]=0,[QuartersHours]=0,[DailyCards]=NULL,[EmpRankCards]=NULL,[CollectBegin]=NULL,[CollectEnd]=NULL");
                            sbSqlEXE2.Append(" WHERE [AttendanceRollcall].[AttendanceRollcallId] IN ");
                            sbSqlEXE2.Append(" (SELECT [AttendanceRollcall].[AttendanceRollcallId]");
                            sbSqlEXE2.Append(" FROM [HRMDB].[dbo].[AttendanceRollcall] WITH (NOLOCK),[HRMDB].[dbo].[AttendanceLeaveInfo] WITH (NOLOCK)");
                            sbSqlEXE2.Append(" WHERE CONVERT(varchar(100),[AttendanceLeaveInfo].[BeginDate],112)<=CONVERT(varchar(100),CONVERT(varchar(100),[AttendanceRollcall].[Date],112),112)");
                            sbSqlEXE2.Append(" AND CONVERT(varchar(100),[AttendanceLeaveInfo].[EndDate],112)>=CONVERT(varchar(100),CONVERT(varchar(100),[AttendanceRollcall].[Date],112),112)");
                            sbSqlEXE2.Append(" AND [AttendanceRollcall].[EmployeeId]=[AttendanceLeaveInfo].[EmployeeId]");
                            sbSqlEXE2.Append(" AND [AttendanceLeaveInfo].[Hours]=8");
                            sbSqlEXE2.Append(" AND [AttendanceRollcall].[Hours]<>0");
                            sbSqlEXE2.AppendFormat(" AND [AttendanceRollcall].[EmployeeId]='{0}' )", od["EmployeeId"].ToString());
                            sbSqlEXE2.Append(" ");
                        }
                    }

                    
                }

                sqlConn.Open();
                tran2 = sqlConn.BeginTransaction();
                cmd2.Connection = sqlConn;
                cmd2.CommandTimeout = 60;
                cmd2.CommandText = sbSqlEXE2.ToString();
                cmd2.Transaction = tran2;
                result = cmd2.ExecuteNonQuery();

                if (result == 0)
                {
                    tran2.Rollback();    //交易取消
                    label4.Text = DateTime.Now.ToString("yyyy/MM/dd") + "CARD FAIL";
                }
                else
                {
                    tran2.Commit();      //執行交易
                    label4.Text = DateTime.Now.ToString("yyyy/MM/dd") + "CARD DONE";

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
