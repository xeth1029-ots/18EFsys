Partial Class CP_04_003_01_History
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload

        If Not IsPostBack Then
            Call create1()
        End If
    End Sub

    Sub create1()
        btnClose2.Attributes("onclick") = "window.close();"
        msg.Text = ""

        Dim rqOCID As String = TIMS.ClearSQM(Request("OCID"))
        Dim dt As DataTable = Nothing
        Dim dt1 As DataTable = Nothing
        Call sUtl_CreateDt1(dt)
        Dim sql As String = ""
        sql = " SELECT DISTINCT a.IDNO FROM Stud_StudentInfo a JOIN Class_StudentsOfClass b ON a.sid = b.sid WHERE b.OCID = '" & rqOCID & "' "
        dt1 = DbAccess.GetDataTable(sql, objconn)
        If dt1.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無資料!!")
            Exit Sub
        End If
        Dim sIDNO As String = ""
        For Each dr1 As DataRow In dt1.Rows
            If sIDNO <> "" Then sIDNO &= ","
            sIDNO &= "'" & dr1("IDNO") & "'"
        Next
        'sIDNO = "'G121386217','U220544423'"
        sql = "" & vbCrLf
        sql &= " SELECT a.STDID " & vbCrLf ' /*PK*/ 
        sql &= " ,a.YEARS ,a.DISTID ,a.DISTNAME ,a.COSUNIT ,a.TRINUNIT OrgName ,a.TPLANID ,a.PLANNAME ,a.CLASSNAME " & vbCrLf
        sql &= " ,CONVERT(VARCHAR, a.SDATE, 111) STDate ,CONVERT(VARCHAR, a.EDATE, 111) FTDate ,a.NAME " & vbCrLf
        sql &= " ,a.SID IDNO " & vbCrLf
        sql &= " ,CONVERT(VARCHAR, a.BIRTH, 111) Birthday " & vbCrLf
        sql &= " ,a.SEX ,a.IDENT ,a.ADDR ,a.TEL ,a.TRANDATE ,NULL TMID ,NULL THours ,NULL SkillName ,'結訓' TFlag " & vbCrLf
        sql &= " FROM STDALL a " & vbCrLf
        sql &= " WHERE a.SID IN (" & sIDNO & ") "
        dt1 = DbAccess.GetDataTable(sql, objconn)
        Call sUtl_dt1Data2dtData(dt1, dt)

        sql = "" & vbCrLf
        sql &= " SELECT a.SERIAL " & vbCrLf '  /*PK*/
        sql &= " ,a.DISTID ,a.DISTNAME ,a.COSUNIT ,a.TRINUNIT OrgName ,a.ORGKIND ,a.TPLANID ,a.PLANNAME ,a.CLASSNAME " & vbCrLf
        sql &= " ,CONVERT(VARCHAR, a.SDATE, 111) STDate ,CONVERT(VARCHAR, a.EDATE, 111) FTDate ,a.NAME" & vbCrLf
        sql &= " ,a.IDNO " & vbCrLf
        sql &= " ,CONVERT(VARCHAR, a.BIRTH, 111) Birthday " & vbCrLf
        sql &= " ,a.SEX ,a.IDENT ,a.DEGREEID ,a.ZIPCODE ,a.ADDR ,a.TEL ,a.TPROPERTYID ,a.JOBLESSID ,a.ISREAD ,a.ONETHREE " & vbCrLf
        sql &= " ,a.FOURSIX ,b.TrainName TMID ,NULL THours ,NULL SkillName ,'結訓' TFlag " & vbCrLf
        sql &= " FROM HISTORY_STUDENTINFO93 a " & vbCrLf
        sql &= " LEFT JOIN Key_TrainType b ON a.TMID = b.TMID "
        sql &= " WHERE a.IDNO IN (" & sIDNO & ") "
        dt1 = DbAccess.GetDataTable(sql, objconn)
        Call sUtl_dt1Data2dtData(dt1, dt)

        sql = "" & vbCrLf
        sql &= " SELECT B.SID ,b.IDNO ,b.Name ,b.Sex " & vbCrLf
        sql &= " ,CONVERT(VARCHAR, b.Birthday, 111) Birthday" & vbCrLf
        sql &= " ,f.Name DistName ,e.OrgName ,g.TrainName TMID " & vbCrLf
        sql &= " ,c.ClassCName ClassName ,c.THours ,CONVERT(VARCHAR, c.STDate, 111) STDate ,CONVERT(VARCHAR, c.FTDate, 111) FTDate " & vbCrLf
        sql &= " ,h.ExamName SkillName ,(CASE WHEN a.StudStatus = 1 THEN '在訓' WHEN a.StudStatus = 2 THEN '離訓' WHEN a.StudStatus = 3 THEN '退訓' WHEN a.StudStatus = 4 THEN '續訓' WHEN a.StudStatus = 5 THEN '結訓' ELSE ' ' END) TFlag " & vbCrLf
        sql &= " FROM Class_StudentsOfClass a " & vbCrLf
        sql &= " JOIN Stud_StudentInfo b ON b.sid = a.sid " & vbCrLf
        sql &= " JOIN Class_ClassInfo c ON a.OCID = c.OCID " & vbCrLf
        sql &= " JOIN ID_Plan i ON i.PlanID = c.PlanID " & vbCrLf
        sql &= " JOIN Auth_Relship d ON c.RID = d.RID " & vbCrLf
        sql &= " JOIN Org_OrgInfo e ON d.OrgID = e.OrgID " & vbCrLf
        sql &= " JOIN ID_District f ON d.DistID = f.DistID " & vbCrLf
        sql &= " JOIN Key_TrainType g ON c.TMID = g.TMID " & vbCrLf
        sql &= " LEFT JOIN Stud_TechExam h ON a.SOCID = h.SOCID " & vbCrLf
        sql &= " WHERE b.IDNO IN (" & sIDNO & ") "
        dt1 = DbAccess.GetDataTable(sql, objconn)
        Call sUtl_dt1Data2dtData(dt1, dt)

        msg.Text = "查無資料!"
        DataGridTable.Visible = False
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            DataGridTable.Visible = True

            dt.DefaultView.Sort = "IDNO,Birthday"
            DataGrid1.AllowPaging = False
            DataGrid1.DataSource = dt.DefaultView
            DataGrid1.DataBind()
        End If
    End Sub

    Sub sUtl_CreateDt1(ByRef dt As DataTable)
        '建立DataGird用的DataTable格式 Start
        dt = New DataTable
        dt.Columns.Add(New DataColumn("IDNO"))       '身分證號
        dt.Columns.Add(New DataColumn("Name"))       '姓名
        dt.Columns.Add(New DataColumn("Sex"))        '性別
        'dt.Columns.Add(New DataColumn("Birthday", System.Type.GetType("System.DateTime")))
        dt.Columns.Add(New DataColumn("Birthday")) '出生年月日
        dt.Columns.Add(New DataColumn("DistName"))   '轄區分署(轄區中心)
        dt.Columns.Add(New DataColumn("OrgName"))    '訓練機構
        dt.Columns.Add(New DataColumn("TMID"))       '訓練職類
        dt.Columns.Add(New DataColumn("ClassName"))  '班別
        dt.Columns.Add(New DataColumn("THours"))     '受訓時數
        dt.Columns.Add(New DataColumn("TRound"))     '受訓期間 'STDate,FTDate
        dt.Columns.Add(New DataColumn("SkillName"))  '技能檢定
        dt.Columns.Add(New DataColumn("TFlag"))      '訓練狀態
        '建立DataGird用的DataTable格式 End
    End Sub

    Sub sUtl_dt1Data2dtData(ByRef dt1 As DataTable, ByRef dt As DataTable)
        For Each dr1 As DataRow In dt1.Rows
            Dim dr As DataRow = dt.NewRow
            dt.Rows.Add(dr)
            dr("IDNO") = TIMS.ChangeIDNO(Convert.ToString(dr1("IDNO")))
            dr("Name") = dr1("Name")
            dr("Sex") = dr1("Sex")
            dr("Birthday") = dr1("Birthday")
            dr("DistName") = dr1("DistName")
            dr("OrgName") = dr1("OrgName")
            dr("TMID") = Convert.ToString(dr1("TMID"))
            dr("ClassName") = dr1("ClassName")
            dr("THours") = dr1("THours")
            If Convert.ToString(dr1("STDate")) <> "" OrElse Convert.ToString(dr1("FTDate")) <> "" Then
                Dim sTRound As String = ""
                sTRound = Convert.ToString(dr1("STDate"))
                sTRound &= "<BR>|<BR>"
                sTRound &= Convert.ToString(dr1("FTDate"))
                dr("TRound") = sTRound
            End If
            dr("SkillName") = dr1("SkillName")
            dr("TFlag") = dr("TFlag") '"結訓"
        Next
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)
        End Select
    End Sub
End Class