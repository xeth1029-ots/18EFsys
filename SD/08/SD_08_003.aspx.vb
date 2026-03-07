Partial Class SD_08_003
    Inherits AuthBasePage

    Dim Stud_SubsidyResult As DataTable
    Dim sql As String = ""
    Dim Days1 As Integer
    Dim Days2 As Integer

    'Dim FunDr As DataRow
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        'conn = DbAccess.GetConnection
        '取出設定天數檔 Start
        Dim dr As DataRow
        sql = "SELECT * FROM Sys_Days"
        dr = DbAccess.GetOneRow(sql, objconn)
        If Not dr Is Nothing Then
            Days1 = dr("Days1")
            Days2 = dr("Days2")
        End If
        '取出設定天數檔 End

        If Not IsPostBack Then
            msg.Text = ""
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            DetailTable.Visible = False
        End If

        Button1.Attributes("onclick") = "javascript:return search();"
        Button1.Style.Item("display") = "none"

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button5.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center", "Button6")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim dt As DataTable
        sql = "SELECT a.*,b.OCID FROM "
        sql += "(SELECT * FROM Stud_SubsidyResult WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID IN (" & OCIDValue.Value & "))) a "
        sql += "JOIN (SELECT * FROM Class_StudentsOfClass WHERE OCID IN (" & OCIDValue.Value & ")) b ON a.SOCID=b.SOCID "
        Stud_SubsidyResult = DbAccess.GetDataTable(sql, objconn)

        sql = "SELECT * FROM "
        sql += "(SELECT * FROM Class_ClassInfo WHERE OCID IN (" & OCIDValue.Value & ")) a "
        dt = DbAccess.GetDataTable(sql, objconn)

        msg.Text = "查無資料"
        DataGridTable.Visible = False
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            DataGridTable.Visible = True

            DataGrid1.DataSource = dt
            DataGrid1.DataKeyField = "OCID"
            DataGrid1.DataBind()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim sql As String = ""
        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing
        Dim da As SqlDataAdapter = Nothing

        Dim dt1 As DataTable = Nothing
        Dim dr1 As DataRow = Nothing
        Dim da1 As SqlDataAdapter = Nothing

        sql = "SELECT * FROM Stud_SubsidyResult WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID='" & NowOCID.Value & "') "
        dt = DbAccess.GetDataTable(sql, da, objconn)

        sql = "SELECT * FROM Class_StudentsOfClass WHERE OCID='" & NowOCID.Value & "'"
        dt1 = DbAccess.GetDataTable(sql, da1, objconn)

        For Each item As DataGridItem In DataGrid2.Items
            Dim AppliedStatusFin As DropDownList = item.FindControl("AppliedStatusFin")
            Dim LRID As HtmlInputHidden = item.FindControl("LRID")
            Dim FailReasonFin As TextBox = item.FindControl("FailReasonFin")

            dr = dt.Select("SUBID='" & DataGrid2.DataKeys(item.ItemIndex) & "'")(0)
            If AppliedStatusFin.SelectedIndex = 0 Then
                dr("AppliedStatusFin") = Convert.DBNull
                dr("FailReasonFin") = Convert.DBNull
            Else
                dr("AppliedStatusFin") = AppliedStatusFin.SelectedValue
                dr("FailReasonFin") = IIf(FailReasonFin.Text = "", Convert.DBNull, FailReasonFin.Text)

                dr1 = dt1.Select("SOCID='" & dr("SOCID") & "'")(0)
                dr1("GetSubsidy") = AppliedStatusFin.SelectedValue
                If AppliedStatusFin.SelectedValue = "Y" Then
                    dr("LRID") = Convert.DBNull
                Else
                    dr("LRID") = LRID.Value
                End If
                dr1("ModifyAcct") = sm.UserInfo.UserID
                dr1("ModifyDate") = Now
            End If

            dr("ModifyAcct") = sm.UserInfo.UserID
            dr("ModifyDate") = Now
        Next
        DbAccess.UpdateDataTable(dt, da)
        DbAccess.UpdateDataTable(dt1, da1)

        Common.MessageBox(Me, "資料儲存成功!")
        Button1_Click(sender, e)

        SearchTable.Visible = True
        DetailTable.Visible = False
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim btn As Button = e.Item.FindControl("Button4")
                e.Item.Cells(0).Text = e.Item.ItemIndex + 1
                If IsNumeric(drv("CyclType")) Then
                    If Int(drv("CyclType")) <> 0 Then
                        e.Item.Cells(1).Text += "第" & Int(drv("CyclType")) & "期"
                    End If
                End If
                e.Item.Cells(3).Text = Stud_SubsidyResult.Select("AppliedStatusS='Y' and OCID='" & drv("OCID") & "'").Length
                e.Item.Cells(4).Text = Stud_SubsidyResult.Select("AppliedStatusS='Y' and AppliedStatusFin IS NULL and OCID='" & drv("OCID") & "'").Length
                e.Item.Cells(5).Text = Stud_SubsidyResult.Select("AppliedStatusS='Y' and AppliedStatusFin='Y' and OCID='" & drv("OCID") & "'").Length
                e.Item.Cells(6).Text = Stud_SubsidyResult.Select("AppliedStatusS='Y' and AppliedStatusFin='N' and OCID='" & drv("OCID") & "'").Length
                If drv("IsClosed").ToString = "Y" Then
                    Select Case CInt(sm.UserInfo.RoleID)
                        Case 0, 1
                            '判斷計畫是否為補助辦理保母職業訓練(46)與辦理照顧服務員職業訓練(47)時,限製天數改成75天
                            '暫時先改這樣,以後還會再改
                            If sm.UserInfo.TPlanID = 46 Or sm.UserInfo.TPlanID = 47 Then
                                If DateDiff(DateInterval.Day, drv("FTDate"), Now.Date) > 75 Then
                                    e.Item.Cells(1).ForeColor = Color.Red
                                    btn.Text = "檢視"
                                End If
                            Else
                                If DateDiff(DateInterval.Day, drv("FTDate"), Now.Date) > Days1 Then
                                    e.Item.Cells(1).ForeColor = Color.Red
                                    btn.Text = "檢視"
                                End If
                            End If

                        Case Else
                            '判斷計畫是否為補助辦理保母職業訓練(46)與辦理照顧服務員職業訓練(47)時,限製天數改成60天
                            '暫時先改這樣,以後還會再改
                            If sm.UserInfo.TPlanID = 46 Or sm.UserInfo.TPlanID = 47 Then
                                If DateDiff(DateInterval.Day, drv("FTDate"), Now.Date) > 60 Then
                                    e.Item.Cells(1).ForeColor = Color.Red
                                    btn.Text = "檢視"
                                End If
                            Else
                                If DateDiff(DateInterval.Day, drv("FTDate"), Now.Date) > Days2 Then
                                    e.Item.Cells(1).ForeColor = Color.Red
                                    btn.Text = "檢視"
                                End If
                            End If

                    End Select
                End If
                If e.Item.Cells(3).Text = 0 Then
                    e.Item.Cells(3).ForeColor = Color.Red
                    btn.Enabled = False
                End If
        End Select
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Dim dt As DataTable
        Dim dr As DataRow
        Dim sql As String

        sql = "SELECT * FROM Class_ClassInfo WHERE OCID='" & DataGrid1.DataKeys(e.Item.ItemIndex) & "'"
        dr = DbAccess.GetOneRow(sql, objconn)

        If Not dr Is Nothing Then
            ClassName.Text = dr("ClassCName")
            If IsNumeric(dr("CyclType")) Then
                If Int(dr("CyclType")) <> 0 Then
                    ClassName.Text += "第" & Int(dr("CyclType")) & "期"
                End If
            End If

            If dr("IsClosed").ToString = "Y" Then
                Select Case CInt(sm.UserInfo.RoleID)
                    Case 0, 1
                        If DateDiff(DateInterval.Day, dr("FTDate"), Now.Date) > Days1 Then
                            Button2.Visible = False
                        End If
                    Case Else
                        If DateDiff(DateInterval.Day, dr("FTDate"), Now.Date) > Days2 Then
                            Button2.Visible = False
                        End If
                End Select
            Else
                Button2.Visible = True
            End If
        End If
        NowOCID.Value = DataGrid1.DataKeys(e.Item.ItemIndex)

        sql = "SELECT a.*,c.*,d.LRID+':'+d.LRName as LRName FROM "
        sql += "(SELECT * FROM Stud_SubsidyResult WHERE AppliedStatusS='Y') a "
        sql += "JOIN (SELECT * FROM Class_StudentsOfClass WHERE OCID='" & DataGrid1.DataKeys(e.Item.ItemIndex) & "') b ON a.SOCID=b.SOCID "
        sql += "JOIN Stud_StudentInfo c ON b.SID=c.SID "
        sql += "LEFT JOIN Key_LapmReason d ON a.LRID=d.LRID "
        sql += "Order By b.StudentID "
        dt = DbAccess.GetDataTable(sql, objconn)

        If dt.Rows.Count > 0 Then
            DataGrid2.DataSource = dt
            DataGrid2.DataKeyField = "SUBID"
            DataGrid2.DataBind()
        End If

        SearchTable.Visible = False
        DetailTable.Visible = True
    End Sub

    Private Sub DataGrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                Dim drop As DropDownList = e.Item.FindControl("DropDownList1")
                drop.Attributes("onchange") = "select_all(this.selectedIndex)"
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim AppliedStatusFin As DropDownList = e.Item.FindControl("AppliedStatusFin")
                Dim FailReasonFin As TextBox = e.Item.FindControl("FailReasonFin")
                Dim LRIDText As TextBox = e.Item.FindControl("LRIDText")
                Dim LRID As HtmlInputHidden = e.Item.FindControl("LRID")
                Dim DropDwon As HtmlGenericControl = e.Item.FindControl("DropDwon")

                e.Item.Cells(0).Text = e.Item.ItemIndex + 1

                Common.SetListItem(AppliedStatusFin, drv("AppliedStatusFin").ToString)
                If drv("AppliedStatusFin").ToString = "N" Then
                    LRIDText.Style("display") = "inline"
                    LRIDText.Text = drv("LRName").ToString
                Else
                    LRIDText.Style("display") = "none"
                End If
                DropDwon.Style("display") = "none"
                FailReasonFin.Text = drv("FailReasonFin").ToString

                If drv("AppliedStatusFin").ToString = "Y" Then
                    If sm.UserInfo.RoleID > 1 Then
                        AppliedStatusFin.Enabled = False
                        FailReasonFin.ReadOnly = True
                    End If

                    e.Item.Cells(9).Text = "勾稽通過"
                Else
                    If drv("AppliedStatusS").ToString = "Y" Then
                        e.Item.Cells(9).Text = "複審通過"
                    Else
                        If drv("AppliedStatusF").ToString = "Y" Then
                            e.Item.Cells(9).Text = "初審通過"
                        End If
                    End If
                End If

                AppliedStatusFin.Attributes("onchange") = "if(this.value=='N'){document.getElementById('" & LRIDText.ClientID & "').style.display='inline';}else{document.getElementById('" & LRIDText.ClientID & "').style.display='none';}"
                DropDwon.Attributes("src") = "SD_08_003_Reason.aspx?ValueField=" & LRID.ClientID & "&TextField=" & LRIDText.ClientID & "&FrameId=" & DropDwon.ClientID & ""
                LRIDText.Attributes("onclick") = "ShowReason('" & DropDwon.ClientID & "');"
                LRIDText.Attributes("onfocus") = "ShowReason('" & DropDwon.ClientID & "');"
                LRIDText.Style("CURSOR") = "hand"
        End Select
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        SearchTable.Visible = True
        DetailTable.Visible = False
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim sql As String
        Dim dt As DataTable
        Dim dr As DataRow
        Dim Flag As Boolean = False

        sql = "SELECT * FROM Cookie_Data WHERE Account='" & sm.UserInfo.UserID & "' and PlanID='" & sm.UserInfo.PlanID & "'"
        dt = DbAccess.GetDataTable(sql, objconn)

        For i As Integer = 1 To 5
            If dt.Select("ItemName='SubsidyRID" & i & "' and ItemValue='" & RIDValue.Value & "'").Length <> 0 Then
                dr = dt.Select("ItemName='SubsidyClass" & i & "'")(0)
                OCIDValue.Value = dr("ItemValue")
                Button1_Click(sender, e)
                Flag = True
            End If
        Next

        If Flag = False Then
            DataGridTable.Visible = False
        End If
    End Sub
End Class
