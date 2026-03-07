Partial Class SD_01_011
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在--------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在--------------------------End
        msg.Text = ""
        msg2.Text = ""
        PageControler1.PageDataGrid = DataGrid1

        If Not IsPostBack Then
            DataGridtable.Visible = False
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            If sm.UserInfo.LID <= 1 Then
                Button2.Disabled = False
                center.Enabled = True
            Else
                Button2.Disabled = True
                center.Enabled = False
            End If

            Page1.Visible = True
            Page2.Visible = False
            '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
            TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
        End If

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, Historytable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1")
        If Historytable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If
        If sm.UserInfo.RID = "A" Or sm.UserInfo.RoleID <= 1 Then
            Button2.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg.aspx');SetOneOCID();"
        Else
            Button2.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg1.aspx');SetOneOCID();"
        End If
        Button1.Attributes("onclick") = "return CheckSearch();"
        Button4.Attributes("onclick") = "ClearData();"
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'If Trim(Me.TxtPageSize.Text) <> "" And IsNumeric(Me.TxtPageSize.Text) Then
        '    If CInt(Me.TxtPageSize.Text) >= 1 Then
        '        Me.TxtPageSize.Text = Trim(Me.TxtPageSize.Text)
        '    Else
        '        Common.RespWrite(Me, "<script>alert('顯示列數不正確，以10 帶入');</script>")
        '        Me.TxtPageSize.Text = 10
        '    End If
        'Else
        '    Common.RespWrite(Me, "<script>alert('顯示列數不正確，以10 帶入');</script>")
        '    Me.TxtPageSize.Text = 10
        'End If
        'If Me.TxtPageSize.Text <> Me.DataGrid1.PageSize Then Me.DataGrid1.PageSize = Me.TxtPageSize.Text
        TIMS.sUtl_TxtPageSize(Me, Me.TxtPageSize, Me.DataGrid1)

        Dim sql As String = ""
        'Dim SearchStr As String
        'sql = "SELECT Distinct SID,IDNO,Name,Birthday FROM view_StudentBasicData WHERE 1=1" & SearchStr
        'sql += " and StudStatus not in (2,3) " & vbCrLf '排除離退訓學員輸入資料 by AMU 20090916
        sql = ""
        sql &= " SELECT Distinct se.SETID " & vbCrLf
        sql += " ,se.IDNO " & vbCrLf
        sql += " ,se.Name " & vbCrLf
        sql += " ,se.Birthday  " & vbCrLf
        sql += " FROM Class_ClassInfo cc " & vbCrLf
        sql += " join Stud_EnterType sy on cc. OCID = sy.OCID1 " & vbCrLf
        sql += " join Stud_EnterTemp se on se.SETID = sy.SETID " & vbCrLf
        sql += " WHERE 1=1"
        sql += " AND cc.RID ='" & RIDValue.Value & "'" & vbCrLf '& SearchStr
        If OCIDValue1.Value <> "" Then
            sql += " and sy.OCID1='" & OCIDValue1.Value & "'" & vbCrLf
        End If
        If IDNO.Text <> "" Then '身分證字號
            sql += " and se.IDNO LIKE '" & IDNO.Text & "%'" & vbCrLf
        End If
        If Name.Text <> "" Then '學員姓名
            sql += " and se.Name like '%" & Name.Text & "%'" & vbCrLf
        End If

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)
        DataGridtable.Visible = False
        msg.Text = "查無資料"
        If dt.Rows.Count > 0 Then
            DataGridtable.Visible = True
            msg.Text = ""

            'PageControler1.SqlPrimaryKeyDataCreate(sql, "SETID")
            PageControler1.PageDataTable = dt
            PageControler1.PrimaryKey = "SETID"
            PageControler1.ControlerLoad()
        End If
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = "SD_TD1"
            Case ListItemType.Item, ListItemType.AlternatingItem
                If e.Item.ItemType = ListItemType.Item Then
                    e.Item.CssClass = "SD_TD2"
                End If
                Dim drv As DataRowView = e.Item.DataItem
                Dim btn As LinkButton = e.Item.FindControl("Button3")
                btn.CommandArgument = drv("IDNO").ToString
        End Select
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        DataGrid2.Visible = False
        msg2.Text = "查無資料"
        If e.CommandArgument = "" Then Exit Sub
        Dim vIDNO As String = TIMS.ClearSQM(e.CommandArgument)
        If vIDNO = "" Then Exit Sub

        Dim parms As New Hashtable
        parms.Add("IDNO", vIDNO)

        'Dim sql As String
        'Dim dt As DataTable
        'Dim dr As DataRow

        Page1.Visible = False
        Page2.Visible = True

        Dim sql As String = ""
        sql = ""
        sql += " SELECT a.*" & vbCrLf
        sql += " FROM STUD_ENTERTEMP a" & vbCrLf
        sql += " WHERE a.IDNO=@IDNO"
        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn, parms)
        If dr Is Nothing Then Exit Sub

        LIDNO.Text = dr("IDNO").ToString
        LName.Text = dr("Name").ToString

        sql = "" & vbCrLf
        sql += " SELECT e.OrgName" & vbCrLf
        sql += " ,d.ClassCName + '第' + CyclType + '期' ClassCName" & vbCrLf
        sql += " ,d.STDate" & vbCrLf
        sql += " ,d.FTDate" & vbCrLf
        sql += " ,a.SumOfMoney " & vbCrLf
        sql += " ,a.AppliedStatus " & vbCrLf
        sql += " ,a.AppliedStatusM " & vbCrLf
        sql += " FROM Stud_SubsidyCost a " & vbCrLf
        sql += " JOIN Class_StudentsOfClass b ON a.SOCID=b.SOCID " & vbCrLf
        sql += " JOIN (SELECT SID,IDNO FROM Stud_StudentInfo WHERE IDNO='" & vIDNO & "') c ON b.SID=c.SID " & vbCrLf
        'sql += "JOIN (SELECT * FROM Stud_EnterTemp WHERE IDNO='" & vIDNO & "') f ON c.IDNO=f.IDNO" & vbCrLf
        sql += " JOIN Class_ClassInfo d ON b.OCID=d.OCID " & vbCrLf
        sql += " JOIN view_RIDName e ON d.RID=e.RID " & vbCrLf
        sql += " ORDER BY d.STDate " & vbCrLf
        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)

        RemainSub.Text = TIMS.Get_3Y_SupplyMoney(Me)  '50000
        '970508 Andy  學員補助金 依登入年度變更  
        '---------------------------------------
        '2007年前(補助金為2萬)
        '2007年(補助金為為3年3萬)
        '2008年(補助金為為3年5萬)
        '2012年(補助金為為3年7萬)
        '----------------------------------------
        'If sm.UserInfo.Years < "2007" Then
        '    RemainSub.Text = 20000
        'Else
        '    If sm.UserInfo.Years = "2007" Then
        '        RemainSub.Text = 30000
        '    Else
        '        If sm.UserInfo.Years >= "2008" Then
        '            RemainSub.Text = 50000
        '        End If
        '    End If
        'End If
        Me.LabTotal.Text = RemainSub.Text
        Me.LabTotal.ToolTip = TIMS.gTip_LabTotalSupplyMoney
        'Me.LabTotal.ToolTip = ""
        'Me.LabTotal.ToolTip += "2007年前，補助金為2萬"
        'Me.LabTotal.ToolTip += "2007年，補助金為為3年3萬" & vbCrLf
        'Me.LabTotal.ToolTip += "2008年，補助金為為3年5萬" & vbCrLf
        'Me.LabTotal.ToolTip += "2012年，補助金為為3年7萬" & vbCrLf

        Me.LabSumOfMoney.Text = 0
        'For Each dr In dt.Select("STDate>='" & FormatDateTime(Now.Date.AddYears(-3), 2) & "' and AppliedStatus=1")
        '    Me.LabSumOfMoney.Text += Int(dr("SumOfMoney"))
        '    RemainSub.Text = Int(RemainSub.Text) - dr("SumOfMoney")
        'Next
        '含職前webservice
        Me.LabSumOfMoney.Text += TIMS.Get_SubsidyCost(vIDNO, "", "", "Y", objconn)
        RemainSub.Text = Int(RemainSub.Text) - CInt(Me.LabSumOfMoney.Text)

        RemainSub.ForeColor = Color.Black
        If Int(RemainSub.Text) < 0 Then
            RemainSub.ForeColor = Color.Red
        End If

        If dt.Rows.Count = 0 Then
            DataGrid2.Visible = False
            msg2.Text = "查無資料"
            Exit Sub
        End If

        DataGrid2.Visible = True
        msg2.Text = ""
        DataGrid2.DataSource = dt
        DataGrid2.DataBind()

    End Sub

    Private Sub DataGrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Const Cst_AppliedStatusM As Integer = 5
        Const Cst_AppliedStatus As Integer = 6

        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = "SD_TD1"
            Case ListItemType.Item, ListItemType.AlternatingItem
                If e.Item.ItemType = ListItemType.Item Then
                    e.Item.CssClass = "SD_TD2"
                End If
                Dim drv As DataRowView = e.Item.DataItem

                '審核狀態
                Select Case drv("AppliedStatusM").ToString
                    Case "Y"
                        e.Item.Cells(Cst_AppliedStatusM).Text = "審核通過" '"申請成功"
                    Case "N"
                        e.Item.Cells(Cst_AppliedStatusM).Text = "審核不通過" '"申請失敗"
                    Case "R"
                        e.Item.Cells(Cst_AppliedStatusM).Text = "退件修正"
                    Case Else
                        e.Item.Cells(Cst_AppliedStatusM).Text = "審核中" '"未審核"
                End Select

                '撥款狀態
                If drv("AppliedStatus").ToString = "1" Then
                    e.Item.Cells(Cst_AppliedStatus).Text = "已撥款" '"申請成功"
                Else
                    Select Case drv("AppliedStatusM").ToString
                        Case "Y" '審核通過
                            e.Item.Cells(Cst_AppliedStatus).Text = "撥款中" '"申請中"
                        Case "N" '審核不通過
                            e.Item.Cells(Cst_AppliedStatus).Text = "不撥款" '"申請中"
                        Case "R" '退件修正
                            e.Item.Cells(Cst_AppliedStatus).Text = "未撥款" '"申請失敗"
                        Case Else '審核中
                            e.Item.Cells(Cst_AppliedStatus).Text = "未撥款"
                    End Select
                End If
        End Select
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Page1.Visible = True
        Page2.Visible = False
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim dr As DataRow
        '判斷機構是否只有一個班級
        dr = TIMS.GET_OnlyOne_OCID(RIDValue.Value)

        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        DataGridtable.Visible = False
        If Not dr Is Nothing Then
            If dr("total") = "1" Then '如果只有一個班級
                TMID1.Text = dr("trainname")
                OCID1.Text = dr("classname")
                TMIDValue1.Value = dr("trainid")
                OCIDValue1.Value = dr("ocid")
                DataGridtable.Visible = False
            End If
        End If
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub
End Class
