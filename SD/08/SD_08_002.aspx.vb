Partial Class SD_08_002
    Inherits AuthBasePage

    Dim Stud_SubsidyResult As DataTable
    Dim Days1 As Integer
    Dim Days2 As Integer
    Dim sql As String = ""

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

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button6.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        Button1.Attributes("onclick") = "javascript:return search()"
        Button1.Style.Item("display") = "none"
        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center", "Button7")
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
        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing
        Dim da As SqlDataAdapter = Nothing
        Dim sql As String = ""

        sql = "SELECT * FROM Stud_SubsidyResult WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID='" & NowOCID.Value & "') "
        dt = DbAccess.GetDataTable(sql, da, objconn)

        For Each item As DataGridItem In DataGrid2.Items
            Dim AppliedStatusS As DropDownList = item.FindControl("AppliedStatusS")
            Dim FailReasonS As TextBox = item.FindControl("FailReasonS")

            dr = dt.Select("SUBID='" & DataGrid2.DataKeys(item.ItemIndex) & "'")(0)
            If AppliedStatusS.SelectedIndex = 0 Then
                dr("AppliedStatusS") = Convert.DBNull
                dr("FailReasonS") = Convert.DBNull
            Else
                dr("AppliedStatusS") = AppliedStatusS.SelectedValue
                dr("FailReasonS") = IIf(FailReasonS.Text = "", Convert.DBNull, FailReasonS.Text)
            End If

            dr("ModifyAcct") = sm.UserInfo.UserID
            dr("ModifyDate") = Now
        Next
        DbAccess.UpdateDataTable(dt, da)

        Common.MessageBox(Me, "資料儲存成功!")
        Button1_Click(sender, e)

        SearchTable.Visible = True
        DetailTable.Visible = False
    End Sub

    '匯出
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim sql As String
        sql = "SELECT   a.IDNO, b.ApplyDate, d.Years, b.UnitCode,b.TrainCode, a.Name, a.Sex, "
        sql += "         a.Birthday, d.STDate, d.FTDate, e.ZipCode1, e.Address, "
        sql += "         e.ZipCode2, e.HouseholdAddress, e.PhoneD, d.ClassCName, "
        sql += "         d.CyclType, d.LevelType, d.TMID, b.TrainingMonth, "
        'change by nick 060523 c.IdentityID --> b.IdentityID
        sql += "         b.SumOfMoney, b.IdentityID, e.HandTypeID, e.HandLevelID,b.SUBID "
        sql += "FROM     Class_StudentsOfClass c, "
        sql += "         Stud_SubsidyResult b, "
        sql += "         Stud_StudentInfo a, "
        sql += "         Class_ClassInfo d, "
        sql += "         Stud_SubData e "
        sql += "WHERE    c.SOCID=b.SOCID "
        sql += "AND      c.SID=a.SID "
        sql += "AND      c.OCID=d.OCID "
        sql += "AND      e.SID=a.SID "
        sql += "AND      b.AppliedStatusS='Y' "
        sql += "AND      c.OCID='" & NowOCID.Value & "' "
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)

        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, "沒有資料匯出!")
            Exit Sub
        Else
            Response.Buffer = False

            'Dim i As Integer
            Dim dr As DataRow
            Dim Flag As Boolean = False
            For Each dr In dt.Rows
                If Flag = False Then
                    Response.AddHeader("content-disposition", "attachment; filename=" & sm.UserInfo.UserID & Second(Now) & ".txt")
                    Response.ContentType = "Application/octet-stream"
                    Response.ContentEncoding = System.Text.Encoding.GetEncoding("Big5")
                    Flag = True
                End If
                Common.RespWrite(Me, Get_ROCDate(FormatDateTime(Convert.ToString(dr("ApplyDate")), DateFormat.ShortDate)) & "|")
                Common.RespWrite(Me, Convert.ToString(dr("IDNO")) & "|")
                Common.RespWrite(Me, Get_ROCDate(FormatDateTime(Convert.ToString(dr("STDate")), DateFormat.ShortDate)) & "|")
                Common.RespWrite(Me, Get_ROCDate(FormatDateTime(Convert.ToString(dr("FTDate")), DateFormat.ShortDate)) & "|")
                Dim Syear As String = Int("20" & dr("Years").ToString) - 1911
                If Int(Syear) < 100 Then
                    Syear = "0" & Syear
                End If
                Common.RespWrite(Me, Syear & "|")
                Common.RespWrite(Me, Convert.ToString(dr("UnitCode")) & "|")
                If Convert.ToString(dr("Name")).Length > 5 Then
                    Common.RespWrite(Me, Left(Convert.ToString(dr("Name")), 5) & "|")
                Else
                    Common.RespWrite(Me, Convert.ToString(dr("Name")) & "|")
                End If
                Common.RespWrite(Me, Convert.ToString(dr("Sex")) & "|")
                Common.RespWrite(Me, Get_ROCDate(FormatDateTime(Convert.ToString(dr("Birthday")), DateFormat.ShortDate)) & "|")
                Dim now_home As String = TIMS.Get_ZipName(Convert.ToString(dr("ZipCode1"))) & Convert.ToString(dr("Address"))
                If now_home.Length > 30 Then
                    Common.RespWrite(Me, Left(now_home, 30) & "|")
                Else
                    Common.RespWrite(Me, now_home & "|")
                End If
                Dim old_home As String = TIMS.Get_ZipName(Convert.ToString(dr("ZipCode2"))) & Convert.ToString(dr("HouseholdAddress"))
                If old_home.Length > 30 Then
                    Common.RespWrite(Me, Left(old_home, 30) & "|")
                Else
                    Common.RespWrite(Me, old_home & "|")
                End If
                Common.RespWrite(Me, Left(Convert.ToString(dr("PhoneD")), 12) & "|")
                Dim Iden As String = ""
                Select Case Convert.ToString(dr("IdentityID"))
                    Case "03"
                        Iden = "A"
                    Case "04"
                        Iden = "B"
                    Case "05"
                        Iden = "D"
                    Case "06"
                        Iden = "C"
                    Case "07"
                        Iden = "E"
                    Case "10"
                        Iden = "G"
                End Select
                Common.RespWrite(Me, Iden & "|")

                Dim className As String = TIMS.GET_CLASSNAME(Convert.ToString(dr("ClassCName")), Convert.ToString(dr("CyclType")))

                Common.RespWrite(Me, Left(className, 10) & "|")
                Common.RespWrite(Me, Convert.ToString(dr("TrainCode")) & "|")
                Common.RespWrite(Me, Convert.ToString(dr("TrainingMonth")) & "|")
                Common.RespWrite(Me, Convert.ToString(dr("SumOfMoney")) & "|")
                Common.RespWrite(Me, "0|")
                Common.RespWrite(Me, "1|")
                Select Case dr("HandTypeID").ToString
                    Case "00"
                        Common.RespWrite(Me, "00|")
                    Case "01"
                        Common.RespWrite(Me, "01|")
                    Case "02"
                        Common.RespWrite(Me, "02|")
                    Case "03"
                        Common.RespWrite(Me, "03|")
                    Case "04"
                        Common.RespWrite(Me, "04|")
                    Case "05"
                        Common.RespWrite(Me, "05|")
                    Case "06"
                        Common.RespWrite(Me, "06|")
                    Case "07"
                        Common.RespWrite(Me, "07|")
                    Case "08"
                        Common.RespWrite(Me, "08|")
                    Case "10"
                        Common.RespWrite(Me, "17|")
                    Case "11"
                        Common.RespWrite(Me, "09|")
                    Case "12"
                        Common.RespWrite(Me, "14|")
                    Case "15"
                        Common.RespWrite(Me, "13|")
                    Case "17"
                        Common.RespWrite(Me, "16|")
                    Case Else
                        Common.RespWrite(Me, "|")
                End Select

                Common.RespWrite(Me, Convert.ToString(dr("HandLevelID")))
                Common.RespWrite(Me, TIMS.sUtl_AntiXss(vbCrLf))
            Next
            If Flag Then
                Response.End()
            End If
        End If
    End Sub

    Function Get_ROCDate(ByVal DateStr As String) As String
        Dim ROC As String = ""
        If DateStr <> "" Then
            If (Year(DateStr) - 1911) < 100 Then
                ROC = "0" & (Year(DateStr) - 1911).ToString
            Else
                ROC = (Year(DateStr) - 1911).ToString
            End If
            If Int(Month(DateStr)) < 10 Then
                ROC += "0" & Month(DateStr).ToString
            Else
                ROC += Month(DateStr).ToString
            End If
            If Int(Day(DateStr)) < 10 Then
                ROC += "0" & Day(DateStr).ToString
            Else
                ROC += Day(DateStr).ToString
            End If
        End If
        Return ROC
    End Function

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
                e.Item.Cells(3).Text = Stud_SubsidyResult.Select("AppliedStatusF='Y' and OCID='" & drv("OCID") & "'").Length
                e.Item.Cells(4).Text = Stud_SubsidyResult.Select("AppliedStatusF='Y' and AppliedStatusS IS NULL and OCID='" & drv("OCID") & "'").Length
                e.Item.Cells(5).Text = Stud_SubsidyResult.Select("AppliedStatusF='Y' and AppliedStatusS='Y' and OCID='" & drv("OCID") & "'").Length
                e.Item.Cells(6).Text = Stud_SubsidyResult.Select("AppliedStatusF='Y' and AppliedStatusS='N' and OCID='" & drv("OCID") & "'").Length
                If drv("IsClosed").ToString = "Y" Then
                    Select Case sm.UserInfo.RoleID
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
                    btn.Enabled = False
                    e.Item.Cells(3).ForeColor = Color.Red
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
                Select Case sm.UserInfo.RoleID
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

        sql = "SELECT a.*,c.* FROM "
        sql += "Stud_SubsidyResult a "
        sql += "JOIN (SELECT * FROM Class_StudentsOfClass WHERE OCID='" & DataGrid1.DataKeys(e.Item.ItemIndex) & "') b ON a.SOCID=b.SOCID "
        sql += "JOIN Stud_StudentInfo c ON b.SID=c.SID "
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
                Dim AppliedAll As DropDownList = e.Item.FindControl("AppliedAll")
                AppliedAll.Attributes("onchange") = "SelectAll();"
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim AppliedStatusS As DropDownList = e.Item.FindControl("AppliedStatusS")
                Dim FailReasonS As TextBox = e.Item.FindControl("FailReasonS")

                e.Item.Cells(0).Text = e.Item.ItemIndex + 1

                Common.SetListItem(AppliedStatusS, drv("AppliedStatusS").ToString)
                FailReasonS.Text = drv("FailReasonS").ToString

                If drv("AppliedStatusFin").ToString = "Y" Then
                    AppliedStatusS.Enabled = False
                    FailReasonS.ReadOnly = True

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
        End Select
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        SearchTable.Visible = True
        DetailTable.Visible = False
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
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


