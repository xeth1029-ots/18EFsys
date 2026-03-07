Public Class ojt1
    Inherits System.Web.UI.Page

    Const cst_TM1_now As String = "now"
    Dim vTM1 As String = ""
    Dim vOCID1 As String = ""

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            CCREATE1()
        End If
    End Sub
    Sub CCREATE1()
        Dim GV1 As String = TIMS.ClearSQM(Request("GV1")) '.ToLower
        If String.IsNullOrEmpty(GV1) Then
            labTITLE1.Text = "(參數有誤!)"
            Return
        End If
        vTM1 = TIMS.ClearSQM(Request("TM1")) '.ToLower
        vOCID1 = TIMS.ClearSQM(Request("OCID1"))
        'vOCID1 = TIMS.GetValue2(TIMS.ClearSQM(Request("OCID1"))) '.ToLower 'If vOCID1 <> "" AndAlso vOCID1 = "0" Then vOCID1 = ""
        'https://localhost:44383/ajax/ojt1.aspx?GV1=ojt7

        If vTM1 <> "" Then vTM1 = vTM1.ToLower()
        Select Case GV1.ToLower()
            Case "ojt1"
                QUERY_1()
            Case "ojt2"
                QUERY_2()
            Case "ojt3"
                QUERY_3()
            Case "ojt4"
                QUERY_4()
            Case "ojt5"
                QUERY_5()
            Case "ojt6"
                QUERY_6()
            Case "ojt7"
                QUERY_7()
            Case Else
                labTITLE1.Text = $"(參數有誤!), {GV1.ToLower}"
        End Select
    End Sub
    Sub QUERY_1()
        labTITLE1.Text = $"報名資料查詢-每分鐘報名人數(count1)-{Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}"
        Dim sSql As String = ""
        If vTM1 <> "" AndAlso vTM1 <> cst_TM1_now Then
            'CONVERT(DATETIME, '20250731 11:50', 120)
            sSql &= $"DECLARE @NTDATE1 DATETIME=CONVERT(DATETIME, '{vTM1} 11:50', 120);"
            sSql &= $"DECLARE @NTDATE2 DATETIME=CONVERT(DATETIME, '{vTM1} 13:00', 120);"
            sSql &= " WITH WC1 AS (SELECT RELENTERDATE,OCID1,SIGNNO,ENTERPATH FROM V_ENTERTYPE2 WITH(NOLOCK) WHERE RELENTERDATE>=@NTDATE1 AND RELENTERDATE<=@NTDATE2 )" & vbCrLf
            sSql &= " ,WC2 AS (SELECT RELENTERDATE,OCID1 FROM STUD_ENTERTYPE2DELDATA WITH(NOLOCK) WHERE RELENTERDATE>=@NTDATE1 AND OCID1 IN (SELECT OCID1 FROM WC1))" & vbCrLf
        Else
            sSql &= " WITH WC1 AS (SELECT RELENTERDATE,OCID1,SIGNNO,ENTERPATH FROM V_ENTERTYPE2 WITH(NOLOCK) WHERE RELENTERDATE>=GETDATE()-0.9 and RELENTERDATE<=GETDATE() )" & vbCrLf
            sSql &= " ,WC2 AS (SELECT RELENTERDATE,OCID1 FROM STUD_ENTERTYPE2DELDATA WITH(NOLOCK) WHERE MODIFYDATE>=GETDATE()-0.9 and OCID1 IN (SELECT OCID1 FROM WC1))" & vbCrLf
        End If
        sSql &= " SELECT SUBSTRING(FORMAT(RELENTERDATE,'HHmm'),1,4) HM1 ,COUNT(1) count1" & vbCrLf
        sSql &= " ,COUNT(CASE WHEN ENTERPATH='o' THEN 1 END ) CNT1_IN,COUNT(CASE WHEN ENTERPATH='O' THEN 1 END ) CNT1_OUT" & vbCrLf
        sSql &= " ,MIN(FORMAT(RELENTERDATE,'yyyyMMdd')) MINDATE,MAX(FORMAT(RELENTERDATE,'yyyyMMdd')) MAXDATE" & vbCrLf
        sSql &= " ,MIN(FORMAT(RELENTERDATE,'HHmmss')) MINTIME,MAX(FORMAT(RELENTERDATE,'HHmmss')) MAXTIME" & vbCrLf
        sSql &= " ,MIN(OCID1) MINOCID1,MAX(OCID1) MAXOCID1,COUNT(DISTINCT OCID1) CLASSCNT1" & vbCrLf
        sSql &= " ,(SELECT COUNT(1) FROM WC2 WHERE SUBSTRING(FORMAT(RELENTERDATE,'HHmm'),1,4)=SUBSTRING(FORMAT(WC1.RELENTERDATE,'HHmm'),1,4)) DELCOUNT1" & vbCrLf
        sSql &= " FROM WC1" & vbCrLf
        If vTM1 = cst_TM1_now Then
            sSql &= " WHERE SIGNNO IS NOT NULL AND FORMAT(RELENTERDATE,'HHmm')>=format(getdate()-0.042,'HHmm') AND FORMAT(RELENTERDATE,'HHmm')<=format(getdate(),'HHmm')" & vbCrLf
        ElseIf vTM1 <> "" AndAlso vTM1 <> cst_TM1_now Then
            sSql &= " WHERE SIGNNO IS NOT NULL" & vbCrLf
        Else
            sSql &= " WHERE SIGNNO IS NOT NULL AND FORMAT(RELENTERDATE,'HHmm')>='1150' AND FORMAT(RELENTERDATE,'HHmm')<='1300'" & vbCrLf
        End If
        sSql &= " GROUP BY SUBSTRING(FORMAT(RELENTERDATE,'HHmm'),1,4)" & vbCrLf
        sSql &= " ORDER BY 1" & vbCrLf
        Try
            Using objconn As SqlConnection = DbAccess.GetConnection()
                TIMS.OpenDbConn(objconn)
                Dim dt1 As New DataTable
                Dim oCmd As New SqlCommand(sSql, objconn)
                dt1.Load(oCmd.ExecuteReader())
                GridView1.DataSource = dt1
                GridView1.DataBind()
            End Using
        Catch ex As Exception
            labTITLE1.Text &= $"<BR>{ex.Message}"
        End Try
    End Sub
    Sub QUERY_2()
        labTITLE1.Text = $"報名資料查詢2-每10分鐘(count1)-{Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}"
        Dim sSql As String = ""
        If vTM1 <> "" AndAlso vTM1 <> cst_TM1_now Then
            'CONVERT(DATETIME, '20250731 11:50', 120)
            sSql &= $"DECLARE @NTDATE1 DATETIME=CONVERT(DATETIME, '{vTM1} 11:50', 120);"
            sSql &= $"DECLARE @NTDATE2 DATETIME=CONVERT(DATETIME, '{vTM1} 13:00', 120);"
            sSql &= " WITH WC1 AS (SELECT RELENTERDATE,OCID1,SIGNNO,ENTERPATH FROM V_ENTERTYPE2 WITH(NOLOCK) WHERE RELENTERDATE>=@NTDATE1 AND RELENTERDATE<=@NTDATE2 )" & vbCrLf
            sSql &= " ,WC2 AS (SELECT RELENTERDATE,OCID1 FROM STUD_ENTERTYPE2DELDATA WITH(NOLOCK) WHERE RELENTERDATE>=@NTDATE1 AND OCID1 IN (SELECT OCID1 FROM WC1))" & vbCrLf
        Else
            sSql &= " WITH WC1 AS (SELECT RELENTERDATE,OCID1,SIGNNO,ENTERPATH FROM V_ENTERTYPE2 WITH(NOLOCK) WHERE RELENTERDATE>=GETDATE()-0.9 and RELENTERDATE<=GETDATE())" & vbCrLf
            sSql &= " ,WC2 AS (SELECT RELENTERDATE,OCID1 FROM STUD_ENTERTYPE2DELDATA WITH(NOLOCK) WHERE MODIFYDATE>=GETDATE()-0.9 and OCID1 IN (SELECT OCID1 FROM WC1))" & vbCrLf
        End If
        sSql &= " SELECT SUBSTRING(FORMAT(RELENTERDATE,'HHmm'),1,3) HM1 ,COUNT(1) count1" & vbCrLf
        sSql &= " ,COUNT(CASE WHEN ENTERPATH='o' THEN 1 END ) CNT1_IN,COUNT(CASE WHEN ENTERPATH='O' THEN 1 END ) CNT1_OUT" & vbCrLf
        sSql &= " ,MIN(FORMAT(RELENTERDATE,'yyyyMMdd')) MINDATE,MAX(FORMAT(RELENTERDATE,'yyyyMMdd')) MAXDATE" & vbCrLf
        sSql &= " ,MIN(FORMAT(RELENTERDATE,'HHmm')) MINTIME,MAX(FORMAT(RELENTERDATE,'HHmm')) MAXTIME" & vbCrLf
        sSql &= " ,MIN(OCID1) MINOCID1,MAX(OCID1) MAXOCID1,COUNT(DISTINCT OCID1) CLASSCNT1" & vbCrLf
        sSql &= " ,(SELECT COUNT(1) FROM WC2 WHERE SUBSTRING(FORMAT(RELENTERDATE,'HHmm'),1,3)=SUBSTRING(FORMAT(WC1.RELENTERDATE,'HHmm'),1,3)) DELCOUNT1" & vbCrLf
        sSql &= " FROM WC1" & vbCrLf
        If vTM1 = cst_TM1_now Then
            sSql &= " WHERE SIGNNO IS NOT NULL AND FORMAT(RELENTERDATE,'HHmm')>=format(getdate()-0.042,'HHmm') AND FORMAT(RELENTERDATE,'HHmm')<=format(getdate(),'HHmm')" & vbCrLf
        ElseIf vTM1 <> "" AndAlso vTM1 <> cst_TM1_now Then
            sSql &= " WHERE SIGNNO IS NOT NULL" & vbCrLf
        Else
            sSql &= " WHERE SIGNNO IS NOT NULL AND FORMAT(RELENTERDATE,'HHmm')>='1150' AND FORMAT(RELENTERDATE,'HHmm')<='1300'" & vbCrLf
        End If
        sSql &= " GROUP BY SUBSTRING(FORMAT(RELENTERDATE,'HHmm'),1,3)" & vbCrLf
        sSql &= " ORDER BY 1" & vbCrLf
        Try
            Using objconn As SqlConnection = DbAccess.GetConnection()
                TIMS.OpenDbConn(objconn)
                Dim dt1 As New DataTable
                Dim oCmd As New SqlCommand(sSql, objconn)
                dt1.Load(oCmd.ExecuteReader())
                GridView1.DataSource = dt1
                GridView1.DataBind()
            End Using
        Catch ex As Exception
            labTITLE1.Text &= $"<BR>{ex.Message}"
        End Try
    End Sub
    ''' <summary>V_ENTERTYPE2</summary>
    Sub QUERY_3()
        labTITLE1.Text = $"報名資料查詢3-依班級看-count1:報名人數(扣掉取消)-DELCOUNT：取消人數)-{Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}"
        Dim sSql As String = ""
        If vTM1 <> "" AndAlso vTM1 <> cst_TM1_now Then
            'CONVERT(DATETIME, '20250731 11:50', 120)
            sSql &= $"DECLARE @NTDATE1 DATETIME=CONVERT(DATETIME, '{vTM1} 11:50', 120);"
            sSql &= $"DECLARE @NTDATE2 DATETIME=CONVERT(DATETIME, '{vTM1} 13:00', 120);"
            sSql &= " WITH WC1 AS (SELECT RELENTERDATE,OCID1,DISTID,CLASSCNAME,SIGNNO,ENTERPATH FROM V_ENTERTYPE2 WITH(NOLOCK) WHERE RELENTERDATE>=@NTDATE1 AND RELENTERDATE<=@NTDATE2 )" & vbCrLf
            sSql &= " ,WC2 AS (SELECT OCID1 FROM STUD_ENTERTYPE2DELDATA WITH(NOLOCK) WHERE RELENTERDATE>=@NTDATE1 AND OCID1 IN (SELECT OCID1 FROM WC1))" & vbCrLf
        Else
            sSql &= " WITH WC1 AS (SELECT RELENTERDATE,OCID1,DISTID,CLASSCNAME,SIGNNO,ENTERPATH FROM V_ENTERTYPE2 WITH(NOLOCK) WHERE RELENTERDATE>=GETDATE()-0.9 and RELENTERDATE<=GETDATE())" & vbCrLf
            sSql &= " ,WC2 AS (SELECT OCID1 FROM STUD_ENTERTYPE2DELDATA WITH(NOLOCK) WHERE MODIFYDATE>=GETDATE()-0.9 and OCID1 IN (SELECT OCID1 FROM WC1))" & vbCrLf
        End If
        sSql &= " SELECT ROW_NUMBER() OVER (ORDER BY COUNT(1) DESC,DISTID,OCID1) ROWNUM,OCID1,DISTID,CLASSCNAME,COUNT(1) count1" & vbCrLf
        sSql &= " ,COUNT(CASE WHEN ENTERPATH='o' THEN 1 END ) CNT1_IN,COUNT(CASE WHEN ENTERPATH='O' THEN 1 END ) CNT1_OUT" & vbCrLf
        sSql &= " ,MIN(FORMAT(RELENTERDATE,'HHmm')) MINTIME,MAX(FORMAT(RELENTERDATE,'HHmm')) MAXTIME" & vbCrLf
        sSql &= " ,COUNT(DISTINCT OCID1) CLASSCNT1,(SELECT COUNT(1) FROM WC2 WHERE OCID1=WC1.OCID1) DELCOUNT1" & vbCrLf
        sSql &= " FROM WC1" & vbCrLf
        If vTM1 = cst_TM1_now Then
            sSql &= " WHERE SIGNNO IS NOT NULL AND FORMAT(RELENTERDATE,'HHmm')>=format(getdate()-0.042,'HHmm') AND FORMAT(RELENTERDATE,'HHmm')<=format(getdate(),'HHmm')" & vbCrLf
        ElseIf vTM1 <> "" AndAlso vTM1 <> cst_TM1_now Then
            sSql &= " WHERE SIGNNO IS NOT NULL" & vbCrLf
        Else
            sSql &= " WHERE SIGNNO IS NOT NULL AND FORMAT(RELENTERDATE,'HHmm')>='1150' AND FORMAT(RELENTERDATE,'HHmm')<='1300'" & vbCrLf
        End If
        sSql &= " GROUP BY OCID1,DISTID,CLASSCNAME" & vbCrLf
        sSql &= " ORDER BY COUNT(1) DESC,DISTID,OCID1" & vbCrLf
        Try
            Using objconn As SqlConnection = DbAccess.GetConnection()
                TIMS.OpenDbConn(objconn)
                Dim dt1 As New DataTable
                Dim oCmd As New SqlCommand(sSql, objconn)
                dt1.Load(oCmd.ExecuteReader())
                GridView1.DataSource = dt1
                GridView1.DataBind()
            End Using
        Catch ex As Exception
            labTITLE1.Text &= $"<BR>{ex.Message}"
        End Try
    End Sub
    ''' <summary>STUD_ENTERTYPE2</summary>
    Sub QUERY_4()
        labTITLE1.Text = $"報名資料查詢4-學員依報名順序-{Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}"
        Dim fgUsePms As Boolean = False
        Dim pms1 As New Hashtable
        vOCID1 = TIMS.GetValue2(TIMS.ClearSQM(vOCID1))
        Dim sSql As String = ""
        If vTM1 <> "" AndAlso vTM1 <> cst_TM1_now Then
            sSql &= $"DECLARE @NTDATE1 DATETIME=CONVERT(DATETIME, '{vTM1} 11:50', 120);"
            sSql &= $"DECLARE @NTDATE2 DATETIME=CONVERT(DATETIME, '{vTM1} 13:00', 120);"
        End If
        sSql &= " SELECT TOP 888 ROW_NUMBER() OVER (ORDER BY RELENTERDATE) ROWNUM,format(ENTERDATE,'yyyy-MM-dd') ENTERDATE,format(RELENTERDATE,'HH-mm-ss-fff') RELENTETIME" & vbCrLf
        sSql &= " ,OCID1,ESETID,ESERNUM,SIGNNO,SIGNUPSTATUS,SIGNUPMEMO,ENTERPATH,dbo.FN_GET_MASK1(MODIFYACCT) MODIFYACCT,format(MODIFYDATE,'yyyy-MM-dd-HH-mm-ss-fff') MODIFYTIME" & vbCrLf
        sSql &= " ,DATEDIFF(second, RELENTERDATE, MODIFYDATE) DiffInSeconds,DATEDIFF(minute, RELENTERDATE, MODIFYDATE) DiffInMinutes" & vbCrLf
        sSql &= " FROM STUD_ENTERTYPE2 WITH(NOLOCK)" & vbCrLf
        If vTM1 = cst_TM1_now Then
            sSql &= " WHERE SIGNNO IS NOT NULL AND RELENTERDATE>=GETDATE()-0.9 AND FORMAT(RELENTERDATE,'HHmm')>=format(getdate()-0.042,'HHmm') AND FORMAT(RELENTERDATE,'HHmm')<=format(getdate(),'HHmm')" & vbCrLf
        ElseIf vTM1 <> "" AndAlso vTM1 <> cst_TM1_now Then '(日期查詢)
            sSql &= " WHERE SIGNNO IS NOT NULL AND RELENTERDATE>=GETDATE()-333 AND RELENTERDATE>=@NTDATE1 AND RELENTERDATE<=@NTDATE2 " & vbCrLf
        ElseIf vOCID1 <> "" AndAlso vOCID1 <> "0" Then
            fgUsePms = True
            pms1.Add("OCID1", CInt(vOCID1))
            sSql &= " WHERE OCID1=@OCID1" & vbCrLf
        Else
            sSql &= " WHERE SIGNNO IS NOT NULL AND RELENTERDATE>=GETDATE()-0.9 AND FORMAT(RELENTERDATE,'HHmm')>='1150' AND FORMAT(RELENTERDATE,'HHmm')<='1300'" & vbCrLf
        End If
        sSql &= " ORDER BY RELENTERDATE" & vbCrLf
        Try
            Using objconn As SqlConnection = DbAccess.GetConnection()
                If fgUsePms Then
                    Dim dt1 As DataTable = DbAccess.GetDataTable(sSql, objconn, pms1)
                    GridView1.DataSource = dt1
                    GridView1.DataBind()
                Else
                    TIMS.OpenDbConn(objconn)
                    Dim dt1 As New DataTable
                    Dim oCmd As New SqlCommand(sSql, objconn)
                    dt1.Load(oCmd.ExecuteReader())
                    GridView1.DataSource = dt1
                    GridView1.DataBind()
                End If
            End Using
        Catch ex As Exception
            labTITLE1.Text &= $"<BR>{ex.Message}"
        End Try
    End Sub
    ''' <summary>SYS_SIGNUP_STATISTICS</summary>
    Sub QUERY_5()
        'SELECT TOP 11 * FROM SYS_SIGNUP_CTL 'SELECT TOP 11 * FROM SYS_SIGNUP_STATISTICS WHERE idx>=format(getdate(),'mm')-1 and idx<=format(getdate(),'mm') order by host,idx
        labTITLE1.Text = $"報名資料查詢5-即時機器狀況-{Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}"
        Dim sSql As String = "SELECT Host,Idx,StatDate,IsCurrent,AvgWaitTime,AvgProcessTime,Wait,Process,Timeouts,Error,Fail,Success FROM SYS_SIGNUP_STATISTICS WITH(NOLOCK) WHERE Idx>=format(GETDATE(),'mm')-1 and Idx<=format(GETDATE(),'mm') ORDER BY Idx,Host"
        Try
            Using objconn As SqlConnection = DbAccess.GetConnection()
                TIMS.OpenDbConn(objconn)
                Dim dt1 As New DataTable
                Dim oCmd As New SqlCommand(sSql, objconn)
                dt1.Load(oCmd.ExecuteReader())
                GridView1.DataSource = dt1
                GridView1.DataBind()
            End Using
        Catch ex As Exception
            labTITLE1.Text &= $"<BR>{ex.Message}"
        End Try
    End Sub
    ''' <summary>SYS_SIGNUP_STATUS</summary>
    Sub QUERY_6()
        'SELECT TOP 11 * FROM SYS_SIGNUP_STATUS WITH(NOLOCK) WHERE [QueueTime]>=GETDATE()-0.9
        labTITLE1.Text = $"報名資料查詢6-即時狀況-{Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}"
        Dim sSql As String = "
WITH WC1 AS ( SELECT HOST,COUNT(1) CNT1,MIN([QueueTime]) QueueTN,MAX([QueueTime]) QueueTX,MIN(StartTime) StartTN,MAX(StartTime) StartTX
,MIN(FinishTime) FinishTN,MAX(FinishTime) FinishTX,COUNT(case [Timeout] when 1 then 1 end) TimeoutCNT
,MIN([WaitTimes]) WaitTN,MAX([WaitTimes]) WaitTX,MIN([ProcessTimes]) ProcessTN ,MAX([ProcessTimes]) ProcessTX 
,AVG([WaitTimes]) WaitTAVG,AVG([ProcessTimes]) ProcessTAVG
FROM SYS_SIGNUP_STATUS WITH(NOLOCK) WHERE [QueueTime]>=GETDATE()-0.0416 AND [STATUS] IN (2,3,9) GROUP BY HOST )
SELECT HOST,CNT1,format(QueueTN,'yyyy-MM-dd HH:mm:ss.fff') QueueTN,format(QueueTX,'yyyy-MM-dd HH:mm:ss.fff') QueueTX
,format(StartTN,'yyyy-MM-dd HH:mm:ss.fff') StartTN,format(StartTX,'yyyy-MM-dd HH:mm:ss.fff') StartTX
,format(FinishTN,'yyyy-MM-dd HH:mm:ss.fff') FinishTN,format(FinishTX,'yyyy-MM-dd HH:mm:ss.fff') FinishTX
,TimeoutCNT,WaitTN,WaitTX,ProcessTN,ProcessTX,FLOOR(WaitTAVG) WaitTAVG,FLOOR(ProcessTAVG) ProcessTAVG FROM WC1 ORDER BY 1"

        Try
            Using objconn As SqlConnection = DbAccess.GetConnection()
                TIMS.OpenDbConn(objconn)
                Dim dt1 As New DataTable
                Dim oCmd As New SqlCommand(sSql, objconn)
                dt1.Load(oCmd.ExecuteReader())
                GridView1.DataSource = dt1
                GridView1.DataBind()
            End Using
        Catch ex As Exception
            labTITLE1.Text &= $"<BR>{ex.Message}"
        End Try
    End Sub
    ''' <summary>SYS_SIGNUP_CTL</summary>
    Sub QUERY_7()
        labTITLE1.Text = $"報名控制表查詢7-即時狀況-{Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}"
        Dim sSql As String = "SELECT HOST,QDAY,SEQ,CURSEQ,PCOUNT,NCOUNT FROM SYS_SIGNUP_CTL WITH(NOLOCK) ORDER BY HOST"
        Try
            Using objconn As SqlConnection = DbAccess.GetConnection()
                TIMS.OpenDbConn(objconn)
                Dim dt1 As New DataTable
                Dim oCmd As New SqlCommand(sSql, objconn)
                dt1.Load(oCmd.ExecuteReader())
                GridView1.DataSource = dt1
                GridView1.DataBind()
            End Using
        Catch ex As Exception
            labTITLE1.Text &= $"<BR>{ex.Message}"
        End Try
    End Sub
End Class
