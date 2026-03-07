Partial Class SD_11_010_R
    Inherits AuthBasePage

    '受訓期間滿意度調查統計表
    Dim blnPrint2016 As Boolean = False
    Const cst_printFN_R1 As String = "SD_11_010_R_1" '班級
    Const cst_printFN_R As String = "SD_11_010_R" '不統計全轄區 RIDValue.Value (old)

    'DataGrid1
    'Columns
    'Cells
    Const Cst_序號 As Integer = 0
    Const Cst_縣市別 As Integer = 1
    Const Cst_訓練單位 As Integer = 2
    Const Cst_班別名稱 As Integer = 3 'ClassCName/CLASSCNAME2
    Const Cst_期別 As Integer = 4 'CYCLTYPE

    Const Cst_開訓日期 As Integer = 5
    Const Cst_結訓日期 As Integer = 6
    Const Cst_結訓人數 As Integer = 7
    Const Cst_填寫人數 As Integer = 8

    Const Cst_第1部分平均滿意度 As Integer = 9
    Const Cst_第2部分平均滿意度 As Integer = 10
    Const Cst_第3部分平均滿意度 As Integer = 11
    Const Cst_第4部分平均滿意度 As Integer = 12
    Const Cst_平均滿意度 As Integer = 13
    Const Cst_功能欄位 As Integer = 14

    'Dim vsQName As String = ""
    'Dim vsQID As String = ""
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '分頁設定
        PageControler1.PageDataGrid = DataGrid1

        If Not IsPostBack Then
            cCreate1()
        End If


    End Sub

    Sub cCreate1()
        'If TIMS.sUtl_ChkTest() Then Common.SetListItem(rblprtType1, Cst_defQA16) '20160501統計表
        yearlist = TIMS.GetSyear(yearlist)
        Common.SetListItem(yearlist, sm.UserInfo.Years)
        yearlist.Items.Remove(yearlist.Items.FindByValue(""))
        DistID = TIMS.Get_DistID(DistID)
        If DistID.Items.FindByValue("") Is Nothing Then DistID.Items.Insert(0, New ListItem("全部", ""))
        Tcitycode = TIMS.Get_CityName(Tcitycode, objconn)

        'center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID
        PlanID.Value = sm.UserInfo.PlanID

        'OCID.Style("display") = "none"
        Print.Visible = False
        btnExport1.Visible = False
        PageControler1.Visible = False
        'msg.Text = TIMS.cst_NODATAMsg11
        'Button3_Click(sender, e)
        'Call sSearch3()

        '2010/05/24 改成若是委訓單位登入下列欄位就不顯示
        Year_TR.Style("display") = If(sm.UserInfo.LID = "2", "none", "")
        DistID_TR.Style("display") = If(sm.UserInfo.LID = "2", "none", "")
        'Button2.Style("display") = If(sm.UserInfo.LID = "2", "none", "")

        'DistID.Attributes("onclick") = "ClearData();"
        Query.Attributes("OnClick") = "javascript:return chk()"
        '選擇全部轄區
        DistID.Attributes("onclick") = "SelectAll('DistID','DistHidden');"
        '辦訓地縣市
        Tcitycode.Attributes("onclick") = "SelectAll('Tcitycode','TcityHidden');"

        Common.SetListItem(DistID, sm.UserInfo.DistID)
        DistID.Enabled = If(sm.UserInfo.DistID = "000", True, False)
    End Sub

    '配合SQL 語法的WHERE條件 
    Function GET_SQLWHERE2_C() As String
        'Dim whereSql As String
        'Dim rst As Boolean = True
        Dim rst As String = "" '
        'Const cst_errMsg1 As String = "只能選擇一個計畫!"

        '辦訓地縣市
        Dim w_Tcitycode As String = TIMS.GetCblValue(Tcitycode)

        '選擇轄區
        Dim w_DistID As String = TIMS.CombiSQLIN(TIMS.GetCblValue(DistID))

        '班級選擇


        '年度選擇
        Dim v_yearlist As String = TIMS.GetListValue(yearlist)
        If v_yearlist = "" Then v_yearlist = sm.UserInfo.Years

        Dim sql As String = ""
        sql = ""
        sql &= String.Format(" AND ip.TPlanID='{0}'", sm.UserInfo.TPlanID) & vbCrLf '大計畫
        sql &= " AND ip.Years = '" & v_yearlist & "' " & vbCrLf

        If w_DistID <> "" Then sql &= " AND ip.DistID IN (" & w_DistID & ") " & vbCrLf '轄區選擇

        If w_Tcitycode <> "" Then sql &= " AND iz.CTID IN (" & w_Tcitycode & ") " & vbCrLf '縣市

        '開訓區間
        If STDate1.Text <> "" Then sql &= " AND a.STDate >= " & TIMS.To_date(STDate1.Text) & vbCrLf
        If STDate2.Text <> "" Then sql &= " AND a.STDate <= " & TIMS.To_date(STDate2.Text) & vbCrLf 'convert(datetime, '" & STDate2.Text & "', 111)" & vbCrLf
        '結訓區間
        If FTDate1.Text <> "" Then sql &= " AND a.FTDate >= " & TIMS.To_date(FTDate1.Text) & vbCrLf 'convert(datetime, '" & FTDate1.Text & "', 111)" & vbCrLf
        If FTDate2.Text <> "" Then sql &= " AND a.FTDate <= " & TIMS.To_date(FTDate2.Text) & vbCrLf 'convert(datetime, '" & FTDate2.Text & "', 111)" & vbCrLf

        rst = sql
        Return rst
    End Function

    '查詢 SQL (原)
    Sub sSearch1()
        Dim c_whereSql As String = GET_SQLWHERE2_C()
        If c_whereSql = "" Then Return

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " WITH WC1 AS (" & vbCrLf
        sql &= " SELECT a.OCID ,a.CyclType ,a.ClassCName" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(a.ClassCName,a.CyclType) CLASSCNAME2" & vbCrLf
        sql &= " ,a.STDate ,a.FTDate ,a.PlanID ,e.OrgName" & vbCrLf
        sql &= " ,d.RID ,d.DistID,ip.Years" & vbCrLf
        sql &= " ,iz.CTName ,iz.CTID" & vbCrLf
        sql &= " FROM dbo.CLASS_CLASSINFO a" & vbCrLf
        sql &= " JOIN dbo.AUTH_RELSHIP d ON d.RID=a.RID" & vbCrLf
        sql &= " JOIN dbo.ORG_ORGINFO e ON e.Orgid=d.OrgID" & vbCrLf
        sql &= " JOIN dbo.VIEW_PLAN ip ON ip.PlanID=a.PlanID" & vbCrLf
        sql &= " JOIN dbo.VIEW_ZIPNAME iz ON iz.ZipCode=a.TaddressZip" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= c_whereSql
        'sql &= " AND ip.Years = '2018'" & vbCrLf
        'sql &= " AND ip.DistID IN ('001')" & vbCrLf
        'sql &= " AND iz.CTID IN (1,2)" & vbCrLf
        'sql &= " AND ip.TPlanID IN ('06')" & vbCrLf
        sql &= " )" & vbCrLf

        sql &= " SELECT a.OCID" & vbCrLf
        sql &= " ,a.CyclType" & vbCrLf
        sql &= " ,a.ClassCName" & vbCrLf
        sql &= " ,a.CLASSCNAME2" & vbCrLf
        sql &= " ,format(a.STDate,'yyyy/MM/dd') STDate" & vbCrLf
        sql &= " ,format(a.FTDate,'yyyy/MM/dd') FTDate" & vbCrLf
        sql &= " ,a.PlanID" & vbCrLf
        sql &= " ,a.OrgName" & vbCrLf
        sql &= " ,a.RID" & vbCrLf
        sql &= " ,a.Years" & vbCrLf
        sql &= " ,a.DistID" & vbCrLf
        sql &= " ,a.CTName" & vbCrLf
        'sql &= " ,ISNULL(b.QID,2) QID" & vbCrLf
        sql &= " ,b.TOTAL" & vbCrLf
        sql &= " ,ISNULL(b.NUM1,0) NUM1" & vbCrLf
        sql &= " ,q4.Q1_AVERAGE Q1_AVERAGE" & vbCrLf
        sql &= " ,q4.Q2_AVERAGE Q2_AVERAGE" & vbCrLf
        sql &= " ,q4.Q3_AVERAGE Q3_AVERAGE" & vbCrLf
        sql &= " ,q4.Q4_AVERAGE Q4_AVERAGE" & vbCrLf
        sql &= " ,q4.AVERAGE AVERAGE" & vbCrLf
        sql &= " FROM WC1 a" & vbCrLf　'CLASS_CLASSINFO

        sql &= " JOIN (" & vbCrLf
        sql &= " SELECT a.OCID" & vbCrLf
        sql &= " ,COUNT(1) TOTAL" & vbCrLf '班級人數
        sql &= " ,COUNT(CASE WHEN q1.SOCID IS NOT NULL THEN 1 END) NUM1" & vbCrLf '填寫人數
        sql &= " FROM WC1 a" & vbCrLf
        sql &= " JOIN dbo.CLASS_STUDENTSOFCLASS cs ON a.OCID = cs.OCID" & vbCrLf
        sql &= " LEFT JOIN dbo.STUD_QUESTRAINING q1 ON q1.SOCID = CS.SOCID AND q1.OCID = CS.OCID" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        'sql &= " AND cs.StudStatus = 5" & vbCrLf
        sql &= " GROUP BY a.OCID" & vbCrLf
        sql &= " ) b ON a.ocid = b.ocid" & vbCrLf

        sql &= " JOIN dbo.V_QUESTRAINING4 q4 ON q4.ocid = a.ocid" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf

        sql &= " ORDER BY a.RID,a.PlanID,a.OCID " & vbCrLf

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)

        'Table4.Visible = True
        Table4.Style("display") = ""
        DataGrid1.Visible = False
        Print.Visible = False
        btnExport1.Visible = False
        PageControler1.Visible = False

        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無資料")
            Exit Sub
        End If

        'Table4.Visible = True
        'Table4.Style("display") = "inline"
        DataGrid1.Visible = True
        Print.Visible = True
        btnExport1.Visible = True
        PageControler1.Visible = True

        'PageControler1.SqlString = sqlstr_class
        'PageControler1.SqlPrimaryKeyDataCreate(sqlstr_class, "OCID", "DistID,PlanID,OCID,CyclType")
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    Private Sub DataGrid1_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Dim sCmdArg As String = e.CommandArgument
        If sCmdArg = "" Then Exit Sub

        Select Case e.CommandName
            Case "Detail"
                ' Const cst_printFN_R1 As String = "SD_11_010_R_1" '班級
                TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN_R1, sCmdArg)
        End Select
    End Sub

    'list 各各班級
    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                e.Item.Cells(Cst_序號).Text = TIMS.Get_DGSeqNo(sender, e) '序號
                Dim drv As DataRowView = e.Item.DataItem
                Dim BtnDetail As Button = e.Item.FindControl("BtnDetail") '列印明細
                '判斷有無填寫人數
                BtnDetail.Enabled = If(Val(drv("NUM1")) > 0, True, False)

                Dim sCmdArg As String = ""
                sCmdArg = ""
                sCmdArg &= "&Years=" & Convert.ToString(drv("Years"))
                sCmdArg &= "&OCID=" & Convert.ToString(drv("OCID"))
                'sCmdArg &= "&WriteNum=" & Convert.ToString(drv("NUM1"))
                '判斷有無填寫人數
                If Val(drv("NUM1")) > 0 Then BtnDetail.CommandArgument = sCmdArg
        End Select

    End Sub


    ''' <summary>
    ''' 查詢班級
    ''' </summary>
    Sub sSearch3()
        Return
        'Dim sql As String
        'Dim dt As DataTable
        'Dim dr As DataRow
        'Dim strSelected As String = ""

        PlanID.Value = TIMS.ClearSQM(PlanID.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        'SELECT * FROM AUTH_RELSHIP WHERE RID ='E1571'
        Dim relship As String = TIMS.GET_RelshipforRID(RIDValue.Value, objconn)


        Dim parms As New Hashtable
        parms.Add("PlanID", PlanID.Value)
        parms.Add("RID", RIDValue.Value)
        Dim dt As DataTable = Nothing
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT cc.OCID " & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.ClassCName,cc.CyclType) CLASSCNAME2" & vbCrLf
        sql &= " FROM CLASS_CLASSINFO cc " & vbCrLf
        sql &= " JOIN ID_PLAN ip ON ip.planid = cc.planid " & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        sql &= " AND cc.NotOpen = 'N' " & vbCrLf
        sql &= " AND cc.IsSuccess = 'Y' " & vbCrLf
        sql &= " AND cc.PlanID = @PlanID" & vbCrLf '" & PlanID.Value & "' " & vbCrLf 
        sql &= " AND cc.RID =@RID" & vbCrLf '" & RIDValue.Value & "' " & vbCrLf
        sql &= " ORDER BY cc.OCID" & vbCrLf

        Try
            dt = DbAccess.GetDataTable(sql, objconn, parms)
        Catch ex As Exception
            Common.MessageBox(Me, ex.ToString)
            Exit Sub
            'Common.RespWrite(Me, sqlstr_class)
        End Try

        'msg.Text = "查無此機構底下的班級"
        'OCID.Style("display") = "none"
        'If dt.Rows.Count = 0 Then Return

        'Dim strSelected As String = ""
        'If dt.Rows.Count > 0 Then
        '    msg.Text = ""
        '    OCID.Style("display") = ""
        '    TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
        '    OCID.Items.Clear()
        '    OCID.Items.Add(New ListItem("全選", "%"))
        '    For Each dr As DataRow In dt.Rows
        '        OCID.Items.Add(New ListItem(dr("CLASSCNAME2"), dr("OCID")))
        '        If Convert.ToString(dr("OCID")) = OCIDValue1.Value Then strSelected = Convert.ToString(dr("OCID"))
        '    Next
        '    If strSelected.ToString <> "" Then OCID.SelectedValue = strSelected
        'End If
    End Sub

    '查詢班級
    'Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
    '    Call sSearch3()
    'End Sub

#Region "hidden1"

    'Private Sub DataGrid1_Detail_1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1_Detail_1.ItemDataBound
    '    Dim dr As DataRowView
    '    Dim a As Integer = 0
    '    Dim a1, a2, a3, a4, a5, a6 As String
    '    dr = e.Item.DataItem

    '    If e.Item.ItemType = ListItemType.Footer Then
    '        For i As Integer = 1 To DataGrid1_Detail_1.Columns.Count - 1
    '            e.Item.Cells(i).Text = 0
    '            For Each Item As DataGridItem In DataGrid1_Detail_1.Items
    '                Select Case i
    '                    Case 1, 2, 3, 4, 5, 6
    '                        e.Item.Cells(i).Text = Int(e.Item.Cells(i).Text) + Int(Item.Cells(i).Text)
    '                End Select
    '            Next

    '            If i = 1 Then a1 = e.Item.Cells(1).Text
    '            If i = 2 Then a2 = e.Item.Cells(2).Text
    '            If i = 3 Then a3 = e.Item.Cells(3).Text
    '            If i = 4 Then a4 = e.Item.Cells(4).Text
    '            If i = 5 Then a5 = e.Item.Cells(5).Text
    '            If i = 6 Then a6 = e.Item.Cells(6).Text

    '        Next

    '        For j As Integer = 1 To 5
    '            a += Int(e.Item.Cells(j).Text)
    '        Next

    '        If a > 0 Then
    '            e.Item.Cells(7).Text = Math.Round(Convert.ToDouble(e.Item.Cells(6).Text) / a * 20, 2).ToString
    '            count1 = 1
    '        Else
    '            e.Item.Cells(7).Text = "0"
    '        End If
    '        Me.ViewState("Q1") = CDbl(e.Item.Cells(7).Text)
    '    End If
    'End Sub

    'Private Sub DataGrid1_Detail_2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1_Detail_2.ItemDataBound
    '    Dim dr As DataRowView
    '    Dim a As Integer = 0
    '    Dim a1, a2, a3, a4, a5, a6 As String
    '    'Dim count2 As Integer = 0
    '    dr = e.Item.DataItem

    '    If e.Item.ItemType = ListItemType.Footer Then
    '        For i As Integer = 1 To DataGrid1_Detail_2.Columns.Count - 1
    '            e.Item.Cells(i).Text = 0
    '            For Each Item As DataGridItem In DataGrid1_Detail_2.Items
    '                If (i = 1) Or (i = 2) Or (i = 3) Or (i = 4) Or (i = 5) Or (i = 6) Then
    '                    e.Item.Cells(i).Text = Int(e.Item.Cells(i).Text) + Int(Item.Cells(i).Text)
    '                End If
    '            Next

    '            If i = 1 Then a1 = e.Item.Cells(1).Text
    '            If i = 2 Then a2 = e.Item.Cells(2).Text
    '            If i = 3 Then a3 = e.Item.Cells(3).Text
    '            If i = 4 Then a4 = e.Item.Cells(4).Text
    '            If i = 5 Then a5 = e.Item.Cells(5).Text
    '            If i = 6 Then a6 = e.Item.Cells(6).Text

    '        Next

    '        For j As Integer = 1 To 5
    '            a += Int(e.Item.Cells(j).Text)
    '        Next

    '        If a > 0 Then
    '            e.Item.Cells(7).Text = Math.Round(Convert.ToDouble(e.Item.Cells(6).Text) / a * 20, 2).ToString
    '            count2 = 1
    '        Else
    '            e.Item.Cells(7).Text = "0"
    '        End If
    '        Me.ViewState("Q2") = CDbl(e.Item.Cells(7).Text)
    '    End If
    'End Sub

    'Private Sub DataGrid1_Detail_3_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1_Detail_3.ItemDataBound
    '    Dim dr As DataRowView
    '    Dim a As Integer = 0
    '    Dim a1, a2, a3, a4, a5, a6 As String
    '    'Dim count3 As Integer = 0
    '    dr = e.Item.DataItem

    '    If e.Item.ItemType = ListItemType.Footer Then
    '        For i As Integer = 1 To DataGrid1_Detail_3.Columns.Count - 1
    '            e.Item.Cells(i).Text = 0
    '            For Each Item As DataGridItem In DataGrid1_Detail_3.Items
    '                If (i = 1) Or (i = 2) Or (i = 3) Or (i = 4) Or (i = 5) Or (i = 6) Then
    '                    e.Item.Cells(i).Text = Int(e.Item.Cells(i).Text) + Int(Item.Cells(i).Text)
    '                End If
    '            Next

    '            If i = 1 Then a1 = e.Item.Cells(1).Text
    '            If i = 2 Then a2 = e.Item.Cells(2).Text
    '            If i = 3 Then a3 = e.Item.Cells(3).Text
    '            If i = 4 Then a4 = e.Item.Cells(4).Text
    '            If i = 5 Then a5 = e.Item.Cells(5).Text
    '            If i = 6 Then a6 = e.Item.Cells(6).Text

    '        Next

    '        For j As Integer = 1 To 5
    '            a += Int(e.Item.Cells(j).Text)
    '        Next

    '        If a > 0 Then
    '            e.Item.Cells(7).Text = Math.Round(Convert.ToDouble(e.Item.Cells(6).Text) / a * 20, 2).ToString
    '            count3 = 1
    '        Else
    '            e.Item.Cells(7).Text = "0"
    '        End If
    '        Me.ViewState("Q3") = CDbl(e.Item.Cells(7).Text)
    '    End If
    'End Sub

    'Private Sub DataGrid1_Detail_4_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1_Detail_4.ItemDataBound
    '    Dim dr As DataRowView
    '    Dim a As Integer = 0
    '    Dim a1, a2, a3, a4, a5, a6 As String
    '    'Dim count4 As Integer = 0
    '    dr = e.Item.DataItem

    '    If e.Item.ItemType = ListItemType.Footer Then
    '        For i As Integer = 1 To DataGrid1_Detail_4.Columns.Count - 1
    '            e.Item.Cells(i).Text = 0
    '            For Each Item As DataGridItem In DataGrid1_Detail_4.Items
    '                If (i = 1) Or (i = 2) Or (i = 3) Or (i = 4) Or (i = 5) Or (i = 6) Then
    '                    e.Item.Cells(i).Text = Int(e.Item.Cells(i).Text) + Int(Item.Cells(i).Text)
    '                End If
    '            Next

    '            If i = 1 Then a1 = e.Item.Cells(1).Text
    '            If i = 2 Then a2 = e.Item.Cells(2).Text
    '            If i = 3 Then a3 = e.Item.Cells(3).Text
    '            If i = 4 Then a4 = e.Item.Cells(4).Text
    '            If i = 5 Then a5 = e.Item.Cells(5).Text
    '            If i = 6 Then a6 = e.Item.Cells(6).Text

    '        Next

    '        For j As Integer = 1 To 5
    '            a += Int(e.Item.Cells(j).Text)
    '        Next

    '        If a > 0 Then
    '            e.Item.Cells(7).Text = Math.Round(Convert.ToDouble(e.Item.Cells(6).Text) / a * 20, 2).ToString
    '            count4 = 1
    '        Else
    '            e.Item.Cells(7).Text = "0"
    '        End If
    '        Me.ViewState("Q4") = CDbl(e.Item.Cells(7).Text)
    '    End If
    'End Sub

    'Private Sub DataGrid1_Detail_5_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1_Detail_5.ItemDataBound
    '    Dim dr As DataRowView
    '    Dim a As Integer = 0
    '    Dim a1, a2, a3, a4, a5, a6 As String
    '    'Dim count5 As Integer = 0
    '    dr = e.Item.DataItem

    '    If e.Item.ItemType = ListItemType.Footer Then
    '        For i As Integer = 1 To DataGrid1_Detail_5.Columns.Count - 1
    '            e.Item.Cells(i).Text = 0
    '            For Each Item As DataGridItem In DataGrid1_Detail_5.Items
    '                If (i = 1) Or (i = 2) Or (i = 3) Or (i = 4) Or (i = 5) Or (i = 6) Then
    '                    e.Item.Cells(i).Text = Int(e.Item.Cells(i).Text) + Int(Item.Cells(i).Text)
    '                End If
    '            Next

    '            If i = 1 Then a1 = e.Item.Cells(1).Text
    '            If i = 2 Then a2 = e.Item.Cells(2).Text
    '            If i = 3 Then a3 = e.Item.Cells(3).Text
    '            If i = 4 Then a4 = e.Item.Cells(4).Text
    '            If i = 5 Then a5 = e.Item.Cells(5).Text
    '            If i = 6 Then a6 = e.Item.Cells(6).Text

    '        Next

    '        For j As Integer = 1 To 5
    '            a += Int(e.Item.Cells(j).Text)
    '        Next

    '        If a > 0 Then
    '            e.Item.Cells(7).Text = Math.Round(Convert.ToDouble(e.Item.Cells(6).Text) / a * 20, 2).ToString
    '            count5 = 1
    '        Else
    '            e.Item.Cells(7).Text = "0"
    '        End If
    '        Me.ViewState("Q5") = CDbl(e.Item.Cells(7).Text)
    '    End If
    'End Sub

    'Private Sub DataGrid1_Detail_6_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1_Detail_6.ItemDataBound
    '    Dim dr As DataRowView
    '    Dim a As Integer = 0
    '    Dim a1, a2, a3, a4, a5, a6 As String
    '    'Dim count6 As Integer = 0
    '    dr = e.Item.DataItem

    '    If e.Item.ItemType = ListItemType.Footer Then
    '        For i As Integer = 1 To DataGrid1_Detail_6.Columns.Count - 1
    '            e.Item.Cells(i).Text = 0
    '            For Each Item As DataGridItem In DataGrid1_Detail_6.Items
    '                If (i = 1) Or (i = 2) Or (i = 3) Or (i = 4) Or (i = 5) Or (i = 6) Then
    '                    e.Item.Cells(i).Text = Int(e.Item.Cells(i).Text) + Int(Item.Cells(i).Text)
    '                End If
    '            Next

    '            If i = 1 Then a1 = e.Item.Cells(1).Text
    '            If i = 2 Then a2 = e.Item.Cells(2).Text
    '            If i = 3 Then a3 = e.Item.Cells(3).Text
    '            If i = 4 Then a4 = e.Item.Cells(4).Text
    '            If i = 5 Then a5 = e.Item.Cells(5).Text
    '            If i = 6 Then a6 = e.Item.Cells(6).Text
    '        Next

    '        For j As Integer = 1 To 5
    '            a += Int(e.Item.Cells(j).Text)
    '        Next

    '        If a > 0 Then
    '            e.Item.Cells(7).Text = Math.Round(Convert.ToDouble(e.Item.Cells(6).Text) / a * 20, 2).ToString
    '            count6 = 1
    '        Else
    '            e.Item.Cells(7).Text = "0"
    '        End If
    '        Me.ViewState("Q6") = CDbl(e.Item.Cells(7).Text)
    '    End If
    'End Sub

#End Region

    Private Sub Print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Print.Click
        '選擇轄區
        '報表要用的 轄區參數
        Dim DistID1 As String = ""
        For i As Integer = 1 To Me.DistID.Items.Count - 1
            If Me.DistID.Items(i).Selected Then
                If DistID1 <> "" Then DistID1 &= ","
                DistID1 &= Convert.ToString("\'" & Me.DistID.Items(i).Value & "\'")
            End If
        Next

        '報表要用的 辦訓地縣市
        Dim TCityCode2 As String = ""
        For i As Integer = 1 To Tcitycode.Items.Count - 1
            If Tcitycode.Items.Item(i).Selected = True Then
                If TCityCode2 <> "" Then TCityCode2 += ","
                TCityCode2 += Convert.ToString("\'" & Tcitycode.Items.Item(i).Value & "\'")
            End If
        Next

        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID

        Dim v_yearlist As String = TIMS.GetListValue(yearlist)
        Dim MyValue As String = ""
        MyValue = "k=r"
        MyValue += "&Years=" & v_yearlist
        MyValue += "&DistID=" & DistID1
        MyValue += "&TPlanID=" & sm.UserInfo.TPlanID 'TPlanID1
        MyValue += "&CTID=" & TCityCode2
        MyValue += "&OCID1=" '& OCIDStr
        MyValue += "&RID=" & RIDValue.Value
        MyValue += "&STTDate=" & Me.STDate1.Text
        MyValue += "&FTTDate=" & Me.STDate2.Text
        MyValue += "&SFTDate=" & Me.FTDate1.Text
        MyValue += "&FFTDate=" & Me.FTDate2.Text
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN_R, MyValue)


    End Sub

    '覆寫，不執行 MyBase.VerifyRenderingInServerForm 方法，解決執行 RenderControl 產生的錯誤   
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)
    End Sub

    Sub Utl_Export1()
        DataGrid1.AllowPaging = False
        DataGrid1.EnableViewState = False  '把ViewState給關了

        DataGrid1.AllowPaging = False
        DataGrid1.Columns(Cst_功能欄位).Visible = False
        DataGrid1.EnableViewState = False  '把ViewState給關了
        Dim objStringWriter As New System.IO.StringWriter
        Dim objHtmlTextWriter As New System.Web.UI.HtmlTextWriter(objStringWriter)
        Div1.RenderControl(objHtmlTextWriter)
        DataGrid1.AllowPaging = True
        DataGrid1.Columns(Cst_功能欄位).Visible = True
        '受訓期間滿意度調查統計表
        'Dim sFileName1 As String = "滿意度調查統計表"
        Dim sFileName1 As String = String.Format("EXPORT_{0}", TIMS.GetDateNo2(4))
        Dim strSTYLE As String = ""
        '套CSS值
        strSTYLE &= ("<style>")
        strSTYLE &= ("td{mso-number-format:""\@"";}")
        strSTYLE &= ("</style>")

        Dim strHTML As String = ""
        strHTML &= (TIMS.sUtl_AntiXss(Convert.ToString(objStringWriter)))

        Dim v_ExpType As String = TIMS.GetListValue(RBListExpType)
        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", v_ExpType) 'EXCEL/PDF/ODS
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", strHTML)
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)

        'Call TIMS.CloseDbConn(objconn)
        TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
    End Sub

    '匯出excel
    Private Sub btnExport1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExport1.Click
        Utl_Export1()
    End Sub

    '查詢
    Protected Sub Query_Click(sender As Object, e As EventArgs) Handles Query.Click
        sSearch1()
    End Sub



End Class