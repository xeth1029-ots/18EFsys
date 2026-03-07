Partial Class SD_06_001
    Inherits AuthBasePage

    'Dim FunDr As DataRow
    Dim Days1 As Integer
    Dim Days2 As Integer

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
        '檢查Session是否存在 End

        If Not IsPostBack Then
            DataGridTable.Visible = False
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            Button3_Click(sender, e)
        End If
        msg.Text = ""

        '取出設定天數檔 Start
        Call TIMS.Get_SysDays(Days1, Days2)
        '取出設定天數檔 End

        Button1.Attributes("onclick") = "javascript:return search();"
        Button2.Attributes("onclick") = "javascript:return chkdata();"

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button8.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "", "", "TMIDValue1", "TMID1", , "Button1")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

        '檢查帳號的功能權限-----------------------------------Start
        'If sm.UserInfo.RoleID <> 0 Then
        '    If sm.UserInfo.FunDt Is Nothing Then
        '        Common.RespWrite(Me, "<script>alert('Session過期');</script>")
        '        Common.RespWrite(Me, "<script>top.location.href='../../logout.aspx';</script>")
        '    Else
        '        Dim FunDt As DataTable = sm.UserInfo.FunDt
        '        Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & Request("ID") & "'")

        '        If FunDrArray.Length = 0 Then
        '            Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
        '            Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
        '        Else
        '            FunDr = FunDrArray(0)
        '            If FunDr("Adds") = 1 Then
        '                Button2.Enabled = True
        '            Else
        '                Button2.Enabled = False
        '            End If
        '            If FunDr("Sech") = 1 Then
        '                Button1.Enabled = True
        '            Else
        '                Button1.Enabled = False
        '            End If
        '        End If
        '    End If
        'End If
        '檢查帳號的功能權限-----------------------------------End

    End Sub

    Function CheckData2(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        If Trim(FTDate1.Text) <> "" Then FTDate1.Text = Trim(FTDate1.Text) Else FTDate1.Text = ""
        If Trim(FTDate2.Text) <> "" Then FTDate2.Text = Trim(FTDate2.Text) Else FTDate2.Text = ""

        If FTDate1.Text <> "" Then
            If Not TIMS.IsDate1(FTDate1.Text) Then
                Errmsg += "結訓日期  起始日期格式有誤" & vbCrLf
            End If
            If Errmsg = "" Then
                FTDate1.Text = CDate(FTDate1.Text).ToString("yyyy/MM/dd")
            End If
        Else
            'Errmsg += "結訓日期 起始日期 為必填" & vbCrLf
        End If

        If FTDate2.Text <> "" Then
            If Not TIMS.IsDate1(FTDate2.Text) Then
                Errmsg += "結訓日期  迄止日期格式有誤" & vbCrLf
            End If
            If Errmsg = "" Then
                FTDate2.Text = CDate(FTDate2.Text).ToString("yyyy/MM/dd")
            End If
        Else
            ' Errmsg += "結訓日期 迄止日期 為必填" & vbCrLf
        End If

        If Errmsg = "" Then
            If FTDate1.Text.ToString <> "" AndAlso FTDate2.Text.ToString <> "" Then
                If CDate(FTDate1.Text) > CDate(FTDate2.Text) Then
                    Errmsg += "【結訓日期 】的起日不得大於【結訓日期 】的迄日!!" & vbCrLf
                End If
            End If
        End If

        'If Convert.ToString(OCID.SelectedValue) = "" _
        '    OrElse Not IsNumeric(OCID.SelectedValue) Then
        '    Errmsg += "統計對象 為必選" & vbCrLf
        'End If

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    '查詢
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Errmsg As String = ""
        Call CheckData2(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Me.labMsg2.Text = ""
        If OCIDValue1.Value <> "" Then
            Dim dr As DataRow
            dr = TIMS.GetOCIDDate(OCIDValue1.Value)
            If Not dr Is Nothing Then
                Dim CCmsg As String = ""
                CCmsg = ""
                CCmsg &= vbTab & "班級名稱：" & Convert.ToString(dr("ClassCName"))
                CCmsg &= vbTab & "開訓日期：" & Common.FormatDate(dr("STDate"))
                CCmsg &= vbTab & "結訓日期：" & Common.FormatDate(dr("FTDate"))
                Me.labMsg2.Text = CCmsg
            End If
        End If

        Dim dt As DataTable
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " SELECT   " & vbCrLf
        sql += " 	cc.OCID" & vbCrLf
        sql += " 	,substring(a.StudentID,len(a.StudentID)-1,2) as StudentID" & vbCrLf
        sql += " 	,c.InsureSalary" & vbCrLf
        sql += " 	,c.ApplyInsurance,c.DropoutInsurance,c.AppliedReason" & vbCrLf
        sql += " 	,b.IDNO,b.Name, a.SOCID,c.SINID,a.StudStatus, a.WorkSuppIdent, b.JobState " & vbCrLf
        sql += " FROM Class_StudentsOfClass a" & vbCrLf
        sql += " 	JOIN Stud_StudentInfo b ON b.SID=a.SID" & vbCrLf
        sql += " 	JOIN Class_Classinfo cc on cc.OCID =a.OCID " & vbCrLf
        sql += " 	LEFT JOIN Stud_Insurance c ON a.SOCID=c.SOCID " & vbCrLf
        sql += " WHERE 1=1" & vbCrLf

        '是否為在職者補助身分 46:補助辦理保母職業訓練'47:補助辦理照顧服務員職業訓練
        If TIMS.Cst_TPlanID46AppPlan5.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '排除在職者補助身分
            sql += " AND ( IsNull(a.WorkSuppIdent,' ') !='Y') " & vbCrLf
        End If

        If sm.UserInfo.RID = "A" Then
            'sql += " and cc.PlanID IN (SELECT PlanID FROM ID_Plan WHERE TPlanID='" & sm.UserInfo.TPlanID & "' and Years='" &
            sql += " and cc.PlanID IN (SELECT PlanID FROM ID_Plan WHERE TPlanID='" & sm.UserInfo.TPlanID & "'and YEARS = '" & sm.UserInfo.Years & "')" & vbCrLf
        Else
            sql += " and cc.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
        End If
        If OCIDValue1.Value <> "" Then
            sql += " AND a.OCID=" & OCIDValue1.Value & vbCrLf
        Else
            If RIDValue.Value <> "" Then
                sql += " AND cc.RID='" & RIDValue.Value & "'" & vbCrLf
            Else
                sql += " AND cc.RID='" & sm.UserInfo.RID & "'" & vbCrLf
            End If
        End If

        If FTDate1.Text <> "" Then
            sql += " AND cc.FTDate >= " & TIMS.To_date(FTDate1.Text) & vbCrLf
        End If
        If FTDate2.Text <> "" Then
            sql += " AND cc.FTDate <= " & TIMS.To_date(FTDate2.Text) & vbCrLf
        End If

        'ORDER BY
        sql += " ORDER BY a.OCID, substring(a.StudentID,len(a.StudentID)-1,2) " & vbCrLf
        dt = DbAccess.GetDataTable(sql, objconn)

        msg.Text = "查無資料!"
        DataGridTable.Visible = False
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            DataGridTable.Visible = True

            DataGrid1.DataSource = dt
            DataGrid1.DataKeyField = "SINID"
            DataGrid1.DataBind()
        End If

        If dt.Rows.Count > 0 Then
            Dim dr As DataRow
            sql = "SELECT * FROM Class_ClassInfo WHERE OCID='" & OCIDValue1.Value & "'"
            dr = DbAccess.GetOneRow(sql, objconn)
            If Not dr Is Nothing Then
                Button2.Enabled = True
                If dr("IsClosed") = "Y" Then
                    Select Case sm.UserInfo.RoleID
                        Case 0, 1
                            '判斷計畫是否為補助辦理保母職業訓練(46)與辦理照顧服務員職業訓練(47)時,限製天數改成75天

                            '是否為在職者補助身分 46:補助辦理保母職業訓練'47:補助辦理照顧服務員職業訓練
                            If TIMS.Cst_TPlanID46AppPlan5.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                                If DateDiff(DateInterval.Day, dr("FTDate"), Now.Date) > 75 Then
                                    '結訓日期超過系統權限天數
                                    TIMS.Tooltip(Button2, "結訓日期超過系統權限天數，不可修改")
                                    Button2.Enabled = False
                                End If
                            Else
                                If DateDiff(DateInterval.Day, dr("FTDate"), Now.Date) > Days1 Then
                                    '結訓日期超過系統權限天數
                                    TIMS.Tooltip(Button2, "結訓日期超過系統權限天數，不可修改")
                                    Button2.Enabled = False
                                End If
                            End If

                        Case Else
                            '判斷計畫是否為補助辦理保母職業訓練(46)與辦理照顧服務員職業訓練(47)時,限製天數改成60天
                            '是否為在職者補助身分 46:補助辦理保母職業訓練'47:補助辦理照顧服務員職業訓練
                            If TIMS.Cst_TPlanID46AppPlan5.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                                If DateDiff(DateInterval.Day, dr("FTDate"), Now.Date) > 60 Then
                                    TIMS.Tooltip(Button2, "結訓日期超過系統權限天數，不可修改")
                                    Button2.Enabled = False
                                End If
                            Else
                                If DateDiff(DateInterval.Day, dr("FTDate"), Now.Date) > Days2 Then
                                    TIMS.Tooltip(Button2, "結訓日期超過系統權限天數，不可修改")
                                    Button2.Enabled = False
                                End If
                            End If

                    End Select
                End If
            End If
        End If
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
            Case ListItemType.Item, ListItemType.AlternatingItem, ListItemType.EditItem
                Dim start_date As TextBox = e.Item.FindControl("start_date")
                Dim Checkbox1 As HtmlInputCheckBox = e.Item.FindControl("Checkbox1")
                Dim IMG1 As HtmlImage = e.Item.FindControl("IMG1")
                Dim Hidden1 As HtmlInputHidden = e.Item.FindControl("Hidden1")

                Dim end_date As TextBox = e.Item.FindControl("end_date")
                Dim Checkbox2 As HtmlInputCheckBox = e.Item.FindControl("Checkbox2")
                Dim IMG2 As HtmlImage = e.Item.FindControl("IMG2")
                Dim Hidden2 As HtmlInputHidden = e.Item.FindControl("Hidden2")

                start_date.Enabled = False
                Checkbox1.Checked = False
                IMG1.Style.Item("display") = "none"
                Hidden1.Value = "N"
                If start_date.Text <> "" Then
                    start_date.Enabled = True
                    start_date.Text = FormatDateTime(start_date.Text, DateFormat.ShortDate)

                    Checkbox1.Checked = True
                    IMG1.Style.Item("display") = "inline"
                    Hidden1.Value = "Y"
                End If
                Checkbox1.Attributes("onclick") = "Apply(" & e.Item.ItemIndex + 2 & ")"
                IMG1.Attributes("onclick") = "javascript:show_calendar('" & start_date.ClientID & "','','','CY/MM/DD');"

                end_date.Enabled = False
                Checkbox2.Checked = False
                IMG2.Style.Item("display") = "none"
                Hidden2.Value = "N"

                If end_date.Text <> "" Then
                    end_date.Enabled = True
                    end_date.Text = FormatDateTime(end_date.Text, DateFormat.ShortDate)

                    Checkbox2.Checked = True
                    IMG2.Style.Item("display") = "inline"
                    Hidden2.Value = "Y"
                End If
                Checkbox2.Attributes("onclick") = "Dropout(" & e.Item.ItemIndex + 2 & ")"
                IMG2.Attributes("onclick") = "javascript:show_calendar('" & end_date.ClientID & "','','','CY/MM/DD');"

        End Select

    End Sub

    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        Dim sql As String = ""
        For Each item As DataGridItem In DataGrid1.Items
            '取出所有欄位資料
            Dim MyTextbox As TextBox

            MyTextbox = item.Cells(3).Controls(1)
            Dim money As String = MyTextbox.Text
            If Trim(money) <> "" Then money = Trim(money) Else money = ""
            If Trim(money) <> "" Then
                If Not IsNumeric(money) Then
                    Errmsg += item.Cells(1).Text & "薪資應為數字 格式有誤!" & vbCrLf
                Else
                    Try
                        money = CInt(money)
                    Catch ex As Exception
                        Errmsg += item.Cells(1).Text & "薪資應為數字 格式有誤!" & vbCrLf
                    End Try
                End If
            End If

            MyTextbox = item.Cells(5).Controls(1)
            Dim Apply As String = MyTextbox.Text
            If Trim(Apply) <> "" Then Apply = Trim(Apply) Else Apply = ""
            If Trim(Apply) <> "" Then
                If Not IsDate(Apply) Then
                    Errmsg += item.Cells(1).Text & "加保日應為日期 格式有誤!" & vbCrLf
                Else
                    Try
                        Apply = CDate(Apply).ToString("yyyy/MM/dd")
                    Catch ex As Exception
                        Errmsg += item.Cells(1).Text & "加保日應為日期 格式有誤!" & vbCrLf
                    End Try
                End If
            End If

            MyTextbox = item.Cells(7).Controls(1)
            Dim Dropout As String = MyTextbox.Text
            If Trim(Dropout) <> "" Then Dropout = Trim(Dropout) Else Dropout = ""
            If Trim(Dropout) <> "" Then
                If Not IsDate(Dropout) Then
                    Errmsg += item.Cells(1).Text & "退保日應為日期 格式有誤!" & vbCrLf
                Else
                    Try
                        Dropout = CDate(Dropout).ToString("yyyy/MM/dd")
                    Catch ex As Exception
                        Errmsg += item.Cells(1).Text & "退保日應為日期 格式有誤!" & vbCrLf
                    End Try
                End If
            End If

            If Convert.ToString(item.Cells(10).Text) <> "" AndAlso OCIDValue1.Value <> "" Then
                sql = "select socid from class_studentsofclass where socid ='" & Convert.ToString(item.Cells(10).Text) & "' and ocid ='" & OCIDValue1.Value & "' "
                Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)
                If dr Is Nothing Then
                    Errmsg += item.Cells(1).Text & "加保資料異常!" & vbCrLf
                End If
            Else
                Errmsg += item.Cells(1).Text & "加保資料異常!" & vbCrLf
            End If

        Next

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    '存檔
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        'Dim da As SqlDataAdapter = nothing
        'Dim conn As SqlConnection
        'TIMS.TestDbConn(Me, conn)

        'Dim ds As New DataSet
        'sql = "SELECT * FROM Stud_Insurance"
        'da = New SqlDataAdapter(sql, conn)
        'da.Fill(ds, "Ins")
        'Dim drResult() As DataRow

        Dim dt As DataTable
        'Dim dtI As DataTable
        'Dim dtS As DataTable
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " select  cs.ocid ,ss.SOCID " & vbCrLf
        sql += " from Stud_Insurance ss" & vbCrLf
        sql += " join class_studentsofclass cs on cs.socid =ss.socid" & vbCrLf
        sql += " where cs.ocid ='" & OCIDValue1.Value & "'" & vbCrLf
        dt = DbAccess.GetDataTable(sql, objconn)
        Dim msgbox As String
        msgbox = ""

        Try
            Call TIMS.OpenDbConn(objconn)
            'Dim cmd1 As SqlCommand = Nothing 'INSERT INTO Stud_Insurance
            'Dim cmd2 As SqlCommand = Nothing
            ''sql += " INSERT INTO Stud_Insurance(SOCID,InsureSalary,ApplyInsurance,DropoutInsurance,AppliedReason,ModifyAcct,ModifyDate) "
            ''sql += "VALUES('" & item.Cells(10).Text & "'," & money & "," & Apply & "," & Dropout & "," & reason & ",'" & sm.UserInfo.UserID & "',getdate())"
            ''sql = "UPDATE Stud_Insurance SET InsureSalary=" & money & ",ApplyInsurance=" & Apply
            ''sql += ",DropoutInsurance=" & Dropout & ",AppliedReason=" & reason
            ''sql += ",ModifyAcct='" & sm.UserInfo.UserID & "',ModifyDate=getdate() where SOCID='" & item.Cells(10).Text & "'"
            'sql = ""
            'sql += " INSERT INTO Stud_Insurance(SOCID,InsureSalary,ApplyInsurance,DropoutInsurance,AppliedReason,ModifyAcct,ModifyDate) "
            'sql += " VALUES(@SOCID, @InsureSalary,@ApplyInsurance,@DropoutInsurance,@AppliedReason,@ModifyAcct,sydate) "
            'cmd1 = New SqlCommand(sql, objconn)
            'sql = ""
            'sql += " UPDATE Stud_Insurance "
            'sql += " SET InsureSalary=@InsureSalary "
            'sql += " ,ApplyInsurance=@ApplyInsurance"
            'sql += " ,DropoutInsurance=@DropoutInsurance"
            'sql += " ,AppliedReason=@AppliedReason"
            'sql += " ,ModifyAcct=@ModifyAcct"
            'sql += " ,ModifyDate=getdate()"
            'sql += " where SOCID=@SOCID "
            'cmd2 = New SqlCommand(sql, objconn)

            For Each item As DataGridItem In DataGrid1.Items
                '取出所有欄位資料
                Dim MyTextbox As TextBox

                MyTextbox = item.Cells(3).Controls(1)
                Dim money As String = MyTextbox.Text
                If Trim(money) = "" Then
                    money = "NULL"
                Else
                    money = CInt(money)
                End If

                MyTextbox = item.Cells(5).Controls(1)
                Dim Apply As String = MyTextbox.Text
                If Trim(Apply) <> "" AndAlso IsDate(Apply) Then
                    Apply = CDate(Apply).ToString("yyyy/MM/dd")
                    'Apply = "'" & Apply & "'"
                    Apply = TIMS.To_date(Apply)
                Else
                    Apply = "NULL"
                End If

                MyTextbox = item.Cells(7).Controls(1)
                Dim Dropout As String = MyTextbox.Text
                If Trim(Dropout) <> "" AndAlso IsDate(Dropout) Then
                    Dropout = CDate(Dropout).ToString("yyyy/MM/dd")
                    Dropout = TIMS.To_date(Dropout) '"'" & & "'"
                Else
                    Dropout = "NULL"
                End If

                MyTextbox = item.Cells(8).Controls(1)
                Dim reason As String = "'" & MyTextbox.Text.Replace("'", "''") & "'"
                If reason = "''" Then
                    reason = "NULL"
                End If

                Me.ViewState("StudName") = Convert.ToString(item.Cells(1).Text)
                If item.Cells(10).Text <> "" Then
                    'drResult = ds.Tables("Ins").Select("SOCID='" & item.Cells(10).Text & "'")
                    'drResult = dt.Select("SOCID='" & item.Cells(10).Text & "'")
                    '
                    '表示Stud_Insurance中沒有此班級的學員加退保紀錄(, 進行Insert的動作)
                    If dt.Select("SOCID='" & item.Cells(10).Text & "'").Length = 0 Then
                        sql = "INSERT INTO Stud_Insurance(SOCID,InsureSalary,ApplyInsurance,DropoutInsurance,AppliedReason,ModifyAcct,ModifyDate) "
                        sql += "VALUES('" & item.Cells(10).Text & "'," & money & "," & Apply & "," & Dropout & "," & reason & ",'" & sm.UserInfo.UserID & "',getdate())"
                    Else
                        sql = "UPDATE Stud_Insurance SET InsureSalary=" & money & ",ApplyInsurance=" & Apply
                        sql += ",DropoutInsurance=" & Dropout & ",AppliedReason=" & reason
                        sql += ",ModifyAcct='" & sm.UserInfo.UserID & "',ModifyDate=getdate() where SOCID='" & item.Cells(10).Text & "'"
                    End If
                    DbAccess.ExecuteNonQuery(sql, objconn)
                End If
            Next
        Catch ex As Exception
            msgbox += Me.ViewState("StudName") & "加保失敗!" & vbCrLf
        End Try

        If msgbox = "" Then
            Common.MessageBox(Me, "存檔成功!" & vbCrLf & "請記得於班級結訓後,完成加退保相關作業!")
        Else
            Common.MessageBox(Me, msgbox)
        End If

        Button1_Click(sender, e)
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub
End Class
