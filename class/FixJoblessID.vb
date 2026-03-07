'Imports System.Data.SqlClient

' 轉換失業週數代碼(2010年度)
'****************************
'2010年開始,失業週數區間改為  
'---------------------------
'04	23週(含)以下
'05	24~51週
'06	52週(含)以上
'===========================
'2009年之前為 
'---------------------------
'01 26週(含)以下	
'02 27~52週	
'03 53週(含)以上	
'****************************
Public Class FixJoblessID
    'Dim connstr As String = System.Configuration.ConfigurationSettings.AppSettings("ConnectionString")
    'Dim conn As New OracleConnection
    'Dim da As New OracleDataAdapter

    'mode=1,儲存舊資料;mode=2,變更資料
    Public Sub chkJoblessID(ByVal mode As Int16, Optional ByVal OCID As String = "")
        Dim sqlcmd As New OracleCommand
        Dim dt_OCID As DataTable
        Dim trans As OracleTransaction
        Try
            'conn.ConnectionString = connstr
            Dim connstr As String = System.Configuration.ConfigurationSettings.AppSettings("ConnectionString")
            Dim conn As New OracleConnection(connstr)
            conn.Open()
            trans = conn.BeginTransaction
            sqlcmd.Connection = conn
            sqlcmd.CommandTimeout = 100000
            sqlcmd.Transaction = trans
            If OCID <> "" Then
                If mode = 1 Then     '備份代碼
                    Backup_ID(OCID, sqlcmd)
                ElseIf mode = 2 Then '轉換代碼
                    Change_ID(OCID, sqlcmd)
                ElseIf mode = 3 Then '還原代碼
                    Restore_ID(OCID, sqlcmd)
                ElseIf mode = 4 Then '找出有填寫實際失業週數的資料修正 失業週數代碼
                    'fixJobless(OCID, sqlcmd)
                    fixJobless(OCID, conn, trans) '資料太多更新不做trans
                End If
            Else
                dt_OCID = getOCID(sqlcmd)
                If dt_OCID.Rows.Count > 0 Then
                    For Each dr As DataRow In dt_OCID.Rows
                        If mode = 1 Then     '備份代碼
                            Backup_ID(Convert.ToString(dr("OCID")), sqlcmd)
                        ElseIf mode = 2 Then '轉換代碼
                            Change_ID(Convert.ToString(dr("OCID")), sqlcmd)
                        ElseIf mode = 3 Then '還原代碼
                            Restore_ID(Convert.ToString(dr("OCID")), sqlcmd)
                        ElseIf mode = 4 Then '找出有填寫實際失業週數的資料修正 失業週數代碼
                            'fixJobless(Convert.ToString(dr("OCID")), sqlcmd)
                            fixJobless(Convert.ToString(dr("OCID")), conn, trans)
                        End If
                    Next
                End If
            End If
            sqlcmd.Transaction.Commit()
            'trans.Commit()
            conn.Close()
        Catch ex As Exception
            sqlcmd.Transaction.Rollback()
            'trans.Rollback()
            Throw ex
        End Try
    End Sub

    Private Function IsInt(ByVal chkstr As String) As Boolean   '判斷是否為正整數
        Try
            If Int32.Parse(chkstr) = chkstr AndAlso Int32.Parse(chkstr) > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Sub update_JoblessID(ByVal RealJobless As String, ByVal IDNO As String, ByVal OJoblessID As String, ByVal conn As OracleConnection, ByVal trans As OracleTransaction) '   ByVal sqlcmd As OracleCommand)
        Dim sql As String
        Dim NJoblessID As String = ""
        Dim Weeks As Int16 = 0
        Dim sqlcmd As New OracleCommand
        If IsInt(Trim(RealJobless)) Then
            Weeks = CInt(Trim(RealJobless))
        End If
        sql = "  update Stud_StudentInfo "
        sql += " set  JoblessID= @JoblessID ,FixID='F' "
        sql += " where  idno= @IDNO "
        sqlcmd.Connection = conn
        sqlcmd.Transaction = trans
        sqlcmd.CommandText = sql
        sqlcmd.Parameters.Clear()
        sqlcmd.Parameters.Add("IDNO", OracleType.VarChar).Value = IDNO
        If Weeks <> 0 Then     '實際失業週數若有填寫則以實際失業週數為準
            Select Case Weeks
                Case Is <= 23                   '23週(含)以下
                    If Weeks > 0 Then
                        NJoblessID = "04"
                    End If
                Case Is >= 24
                    If Weeks <= 51 Then         '24~51週
                        NJoblessID = "05"
                    Else                        '52週(含)以上
                        NJoblessID = "06"
                    End If
            End Select
        Else
            If Trim(OJoblessID) <> "" Then     '若只有填寫代碼則以代碼為準
                If OJoblessID = "01" Then
                    NJoblessID = "04"
                ElseIf OJoblessID = "02" Then
                    NJoblessID = "05"
                ElseIf OJoblessID = "03" Then
                    NJoblessID = "06"
                End If
            End If
        End If
        If NJoblessID = "" Then
            sqlcmd.Parameters.Add("JoblessID", OracleType.VarChar).Value = Convert.DBNull
        Else
            sqlcmd.Parameters.Add("JoblessID", OracleType.VarChar).Value = NJoblessID
        End If
        sqlcmd.ExecuteNonQuery()
    End Sub

    '(轉換) 99年度代碼
    Private Function Change_ID(ByVal OCID As String, ByVal sqlcmd As OracleCommand) As Boolean
        Dim sql As String
        Dim result As Boolean = False
        sql = "  update  a "
        sql += "  set  a.JoblessID =b.JoblessID ,a.FixID='Y' "
        sql += "  from   Stud_StudentInfo  a  "
        sql += "  join   ( "
        sql += " select  aa.sid,"
        sql += "  case  rtrim(ltrim(aa.JoblessID))   when  '01'   then   '04'   when  '02'   then   '05'   when  '03'   then   '06'  end as  JoblessID  "
        sql += " from "
        sql += " (select  RealJobless, idno,JoblessID,sid   from   Stud_StudentInfo  )  aa "
        sql += " join ( select  sid  from   Class_StudentsOfClass   where  ocid = @OCID )  bb  on  bb.sid=aa.sid "
        sql += "  )   b    on   b.sid=a.sid "
        sql += " where  a.FixID is null  "
        sqlcmd.CommandText = sql
        sqlcmd.Parameters.Clear()
        sqlcmd.Parameters.Add("OCID", OracleType.VarChar).Value = OCID
        If sqlcmd.ExecuteNonQuery() Then
            result = True
        End If
        Return result
    End Function

    '(還原) 99年度代碼由 JoblessID_99 至 JoblessID 欄位
    Private Function Restore_ID(ByVal OCID As String, ByVal sqlcmd As OracleCommand) As Boolean
        Dim sql As String
        Dim result As Boolean = False
        sql = "  update   a "
        sql += "  set  a.JoblessID =b.JoblessID_99,a.FixID='Y' "
        sql += "  from   Stud_StudentInfo  a  "
        sql += "  join   ( "
        sql += " select  aa.sid, aa.JoblessID,aa.JoblessID_99  "
        sql += " from "
        sql += " (select  RealJobless, idno,JoblessID,sid,JoblessID_99    from   Stud_StudentInfo  )  aa "
        sql += " join ( select distinct   sid  from   Class_StudentsOfClass   where  ocid = @OCID )  bb  on  bb.sid=aa.sid "
        sql += "  )   b    on   b.sid=a.sid "
        sql += "  where   b.JoblessID_99 is not null "
        sqlcmd.CommandText = sql
        sqlcmd.Parameters.Clear()
        sqlcmd.Parameters.Add("OCID", OracleType.VarChar).Value = OCID
        If sqlcmd.ExecuteNonQuery() Then
            result = True
        End If
        Return result
    End Function

    '(保留) 99年度代碼至 JoblessID_99 欄位
    Private Function Backup_ID(ByVal OCID As String, ByVal sqlcmd As OracleCommand) As Boolean
        Dim sql As String
        Dim result As Boolean = False
        sql = "  update   a  "
        sql += "  set   a.JoblessID_99 =b.JoblessID "
        sql += "  from   Stud_StudentInfo  a  "
        sql += "  join   ( "
        sql += " select  aa.sid, aa.JoblessID,aa.JoblessID_99  "
        sql += " from "
        sql += " (select  RealJobless, idno,JoblessID,sid,JoblessID_99  from   Stud_StudentInfo  )  aa "
        sql += " join ( select  sid  from   Class_StudentsOfClass   where  ocid = @OCID )  bb  on  bb.sid=aa.sid "
        sql += "  )   b    on   b.sid=a.sid "
        sql += " where  a.JoblessID_99 is null  "
        sqlcmd.CommandText = sql
        sqlcmd.Parameters.Clear()
        sqlcmd.Parameters.Add("OCID", OracleType.VarChar).Value = OCID
        If sqlcmd.ExecuteNonQuery() Then
            result = True
        End If
        Return result
    End Function

    '依user填寫之實際失業週數更新失業週數代碼
    Private Sub fixJobless(ByVal OCID As String, ByVal conn As OracleConnection, ByVal trans As OracleTransaction)    'ByVal sqlcmd As OracleCommand)
        Dim dt As New DataTable
        Dim sql As String = ""
        Try
            sql = "select  a.RealJobless,a.JoblessID,a.idno,a.FixID  FROM  Stud_StudentInfo a"
            sql += " join (select  sid,ocid  from Class_StudentsOfClass)b  on  b.sid=a.sid"
            sql += " where  a.RealJobless  is not null  and  b.ocid= @OCID and (a.FixID is null  or  a.FixID <>'F' )    "
            Dim da As New OracleDataAdapter
            da.SelectCommand = New OracleCommand
            da.SelectCommand.Connection = conn
            da.SelectCommand.Transaction = trans
            da.SelectCommand.CommandText = sql
            da.SelectCommand.Parameters.Clear()
            da.SelectCommand.Parameters.Add("OCID", OracleType.VarChar).Value = OCID
            da.Fill(dt)
            If dt.Rows.Count > 0 Then
                For i As Int16 = 0 To dt.Rows.Count - 1
                    If IsInt(Trim(Convert.ToString((dt.Rows(i)("RealJobless"))))) Then
                        update_JoblessID(Convert.ToString((dt.Rows(i)("RealJobless"))), Trim(Convert.ToString(dt.Rows(i)("IDNO"))), Convert.ToString(dt.Rows(i)("JoblessID")), conn, trans) ', sqlcmd)
                    End If
                Next
            End If
            da.Dispose()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    '取得99年度班級代碼
    Private Function getOCID(ByVal sqlcmd As OracleCommand) As DataTable
        Dim dt As New DataTable
        Dim sql As String = ""
        sql = " select  distinct   ocid  from  Class_ClassInfo  " & vbCrLf
        sql += " where years='10' " & vbCrLf
        sql += " and  IsSuccess='Y' " & vbCrLf
        sql += " and  NotOpen='N' "
        Try
            sqlcmd.CommandText = sql
            Dim da As New OracleDataAdapter
            da.SelectCommand = sqlcmd
            da.Fill(dt)
            da.Dispose()
        Catch ex As Exception
            Throw ex
        End Try
        Return dt
    End Function

End Class
