Partial Class SD_05_022
    Inherits AuthBasePage

    Const Cst_StudentID As Integer = 1
    Const Cst_Sex As Integer = 4
    Const Cst_StudStatus As Integer = 6

   'Dim au As New cAUTH
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
       'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在--------------------------End

        If Not IsPostBack Then
            msg.Text = ""
            msg.Visible = True

            bt_search.Attributes("onclick") = "return CheckData();"
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            TableShowData.Visible = False
            If sm.UserInfo.LID <> "2" Then
                TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
            Else
                Button3_Click(sender, e)
            End If
        End If

        If sm.UserInfo.RID = "A" Or sm.UserInfo.RoleID <= 1 Then
            BtnOrg.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg.aspx');"
        Else
            BtnOrg.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg1.aspx');"
        End If
        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1", True, "bt_search")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

        'btn_Save2.Enabled = False
        'If au.blnCanAdds OrElse au.blnCanMod Then btn_Save2.Enabled = True
        'bt_search.Enabled = False
        'If au.blnCanSech Then bt_search.Enabled = True
    End Sub

    Private Sub bt_search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_search.Click
        GetStudentData()
    End Sub

    'Function GetStudentData(ByVal SearchStr1 As String)
    Sub GetStudentData()
        'Dim RecordCountInt As Integer = 2000
        'Dim SearchStr1 As String = ""
        'If Me.OCIDValue1.Value <> "" Then
        '    SearchStr1 += " AND OCID='" & Me.OCIDValue1.Value & "' " & vbCrLf
        'End If
        'sql += "  " & vbCrLf
        'sql += " (select * from class_classinfo where 1=1 " & SearchStr1 & " ) cc " & vbCrLf

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " select st.Times" & vbCrLf
        sql &= " ,st.Rate" & vbCrLf
        sql &= " ,cc.ClassCName + cc.CyclTYPE Classname " & vbCrLf
        sql &= " ,cc.ClassCName+'第'+cc.CyclType+'期' Classname " & vbCrLf
        sql &= " from class_classinfo cc" & vbCrLf
        sql &= " left join Stud_TrainCost st on cc.ocid = st.ocid " & vbCrLf
        sql &= " where 1=1" & vbCrLf
        If Me.OCIDValue1.Value <> "" Then
            sql &= " AND cc.OCID='" & Me.OCIDValue1.Value & "' " & vbCrLf
        Else
            sql &= " and 1<>1"
        End If
        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)

        msg.Text = "查無資料"
        'msg.Visible = True

        TableShowData.Visible = False
        DataGrid2.Visible = False
        print_btn.Visible = False
        If dt.Rows.Count = 0 Then
            msg.Text = "查無資料"
            'msg.Visible = True

            TableShowData.Visible = False
            DataGrid2.Visible = False
            print_btn.Visible = False
            Exit Sub
        End If
        'msg.Visible = False
        msg.Text = ""

        TableSearch.Visible = True
        TableShowData.Visible = False
        DataGrid2.Visible = True
        print_btn.Visible = True
        DataGrid2.DataSource = dt
        DataGrid2.DataBind()

    End Sub

    Sub GetStudentData2(ByVal iTimes As Integer)
        If Me.OCIDValue1.Value = "" Then
            Common.MessageBox(Me, "條件異常，請選擇有效條件!!")
            Exit Sub
        End If

        Dim sTimes As String = TIMS.Get_StudTrainCostTimes(Me.OCIDValue1.Value, "T", objconn)

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT c.SID, c.StudentID" & vbCrLf
        sql &= " ,SUBSTR(c.StudentID, -2) StudID" & vbCrLf
        sql &= " ,c.OCID, c.SOCID, c.StudStatus" & vbCrLf
        sql &= " ,s.IDNO, s.Birthday, s.Name , s.Sex" & vbCrLf
        sql &= " ,d.JobCost, d.OtherJobCost" & vbCrLf
        sql &= " ,d.State ,t.Cost , Times" & vbCrLf
        sql &= " FROM (" & vbCrLf
        sql &= " SELECT * FROM CLASS_STUDENTSOFCLASS WHERE 1=1" & vbCrLf
        sql &= " AND OCID='" & Me.OCIDValue1.Value & "'" & vbCrLf
        sql &= " ) c" & vbCrLf
        sql &= " JOIN Stud_StudentInfo s ON c.SID=s.SID " & vbCrLf
        sql &= " JOIN Stud_TrainCostT t ON c.SOCID=t.SOCID " & vbCrLf
        sql &= " LEFT JOIN Stud_TrainCostD d ON c.SOCID=d.SOCID " & vbCrLf
        If iTimes = 99 Then '新增
            sql &= " AND t.Times=" & sTimes & vbCrLf
        Else
            sql &= " AND t.Times=" & CStr(iTimes) & vbCrLf
        End If
        sql &= " ORDER BY 3" & vbCrLf ',SUBSTR(c.StudentID, -2) StudID

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)

        msg.Text = "查無資料"
        'msg.Visible = True
        DataGrid1.Visible = False
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            'msg.Visible = False

            labRateAdd.Text = TIMS.Get_StudTrainRate(Me.OCIDValue1.Value, objconn)
            times2.Value = "99"
            If iTimes <> 99 Then
                times2.Value = iTimes
            End If

            DataGrid2.Visible = False
            TableSearch.Visible = False

            TableShowData.Visible = True
            DataGrid1.Visible = True
            DataGrid1.DataSource = dt
            DataGrid1.DataBind()
        End If

    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                'Dim dr As DataRowView = e.Item.DataItem
                e.Item.CssClass = "SD_TD1"
                Dim labTimes As Label = e.Item.FindControl("labTimes")
                Dim txtRate As TextBox = e.Item.FindControl("txtRate")
                txtRate.Attributes("onBlur") = "SET_Rate(this);"
                txtRate.Attributes("onChange") = "SET_Rate(this);"

                'labTimes.Text = TIMS.Get_StudTrainCostTimes(Me.OCIDValue1.Value)
                'txtRate.Text = TIMS.Get_StudTrainCostTimes(Me.OCIDValue1.Value, labTimes.Text)
                labTimes.Text = times2.Value
                If times2.Value = "99" Then
                    labTimes.Text = TIMS.Get_StudTrainCostTimes(Me.OCIDValue1.Value, "T", objconn)
                End If
                txtRate.Text = TIMS.Get_StudTrainCostTimes(Me.OCIDValue1.Value, labTimes.Text, objconn)

                Me.ViewState("sd_05_022_Times") = labTimes.Text
                Me.hidRate.Value = txtRate.Text
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                'If e.Item.ItemType = ListItemType.Item Then e.Item.CssClass = "SD_TD2"

                Dim Checkbox1 As HtmlInputCheckBox = e.Item.FindControl("Checkbox1")
                Dim labName As Label = e.Item.FindControl("labName")
                Dim txtJobCost As TextBox = e.Item.FindControl("txtJobCost")
                Dim txtOtherJobCost As TextBox = e.Item.FindControl("txtOtherJobCost")
                Dim txtState As TextBox = e.Item.FindControl("txtState")
                Dim txtCost As TextBox = e.Item.FindControl("txtCost")
                Dim labStudStatus As Label = e.Item.FindControl("labStudStatus")

                Checkbox1.Value = drv("SOCID")
                Checkbox1.Checked = True
                If drv("Cost").ToString <> "" Then
                    Checkbox1.Disabled = True
                End If

                e.Item.Cells(Cst_StudentID).Text = drv("StudID") 'Right(drv("StudentID"), 2)
                Select Case drv("Sex").ToString
                    Case "M"
                        e.Item.Cells(Cst_Sex).Text = "男"
                    Case "F"
                        e.Item.Cells(Cst_Sex).Text = "女"
                End Select

                Select Case drv("StudStatus")
                    Case 1
                        labStudStatus.Text = "在訓"
                    Case 2
                        labStudStatus.Text = "離訓"
                    Case 3
                        labStudStatus.Text = "退訓"
                    Case 4
                        labStudStatus.Text = "續訓"
                    Case 5
                        labStudStatus.Text = "結訓"
                End Select

                labName.ToolTip = "報名資料" & vbCrLf
                labName.ToolTip += "IDNO:" & drv("IDNO").ToString & vbCrLf
                labName.ToolTip += "Birthday:" & FormatDateTime(drv("Birthday"), DateFormat.ShortDate)

                labName.Style("CURSOR") = "hand"
                labName.Attributes("onclick") = "open_History('" & drv("IDNO").ToString & "');" '.CommandArgument = drv("IDNO").ToString

                Dim Post As String = TIMS.Get_StudTrainCostT(drv("SOCID"), Me.ViewState("sd_05_022_Times"), objconn)
                If Post <> "" Then
                    Post = labName.Text & vbCrLf & Post
                    labStudStatus.ToolTip = Post
                    txtCost.ToolTip = Post
                End If

        End Select
    End Sub

    Private Sub btn_back_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_back.Click
        TableSearch.Visible = True
        DataGrid2.Visible = True
        TableShowData.Visible = False
    End Sub

    Private Sub btn_Save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Save2.Click
        If Check_SaveDate() Then
            Select Case CType(sender, System.Web.UI.WebControls.Button).ClientID
                Case "btn_Save2"
                    '    Insert_Stud_TrainCost(2)
                    'Case Else
                    Insert_Stud_TrainCost()
            End Select
            'bt_search_Click(sender, e)
            GetStudentData()
        End If
    End Sub

    Function Check_SaveDate() As Boolean
        Dim Errmsg As String = ""
        Check_SaveDate = False

        Me.hidRate.Value = Me.hidRate.Value.Trim
        If Not Me.hidRate.Value.Trim <> "" Then
            Errmsg += "撥款百分比比率為必填" & vbCrLf
        Else

        End If
        If Not IsNumeric(Me.hidRate.Value.Trim) Then
            Errmsg += "撥款百分比比率必須為數字" & vbCrLf
        Else
            If CInt(Me.hidRate.Value.Trim) <= 0 Or CInt(Me.hidRate.Value.Trim) > 100 Then
                Errmsg += "撥款百分比比率必須大於0小於等於100" & vbCrLf
            End If
        End If
        If Errmsg = "" Then
            If CInt(labRateAdd.Text) + CInt(Me.hidRate.Value.Trim) > 100 Then
                Errmsg += "累計撥款百分比比率:" & CInt(labRateAdd.Text) + CInt(Me.hidRate.Value.Trim) & vbCrLf
                Errmsg += "累計撥款百分比比率必須小於等於100" & vbCrLf
            End If
        End If

        Dim Errmsg2 As String = ""
        Dim i As Integer = 0
        For Each Item As DataGridItem In DataGrid1.Items
            Dim Checkbox1 As HtmlInputCheckBox = Item.FindControl("Checkbox1")
            Dim IDNO As HtmlInputHidden = Item.FindControl("IDNO")

            Dim labName As Label = Item.FindControl("labName")
            Dim txtJobCost As TextBox = Item.FindControl("txtJobCost")
            Dim txtOtherJobCost As TextBox = Item.FindControl("txtOtherJobCost")
            Dim txtState As TextBox = Item.FindControl("txtState")
            Dim txtCost As TextBox = Item.FindControl("txtCost")

            If Checkbox1.Checked Then
                Errmsg2 = ""
                If txtJobCost.Text.Trim <> "" Then
                    If Not IsNumeric(txtJobCost.Text.Trim) Then
                        i += 1
                        Errmsg2 += "個人訓練費用(含就業輔導費)必須為數字" & vbCrLf
                    End If
                End If

                If txtOtherJobCost.Text.Trim <> "" Then
                    If Not IsNumeric(txtOtherJobCost.Text.Trim) Then
                        i += 1
                        Errmsg2 += "個人訓練費用(不含就業輔導費)必須為數字" & vbCrLf
                    End If
                End If

                If txtState.Text.Trim <> "" Then
                    If txtState.Text.Trim = "" Then
                        i += 1
                        Errmsg2 += "備註：受訓狀況(中長期失業週數)不可為空" & vbCrLf
                    End If
                End If

                If txtCost.Text.Trim <> "" Then
                    If Not IsNumeric(txtCost.Text.Trim) Then
                        i += 1
                        Errmsg2 += "撥款費用必須為數字" & vbCrLf
                    End If
                End If

                'If txtJobCost.Text.Trim = "" _
                '    And txtOtherJobCost.Text.Trim = "" _
                '    And txtState.Text.Trim = "" _
                '    And txtCost.Text.Trim = "" Then

                '    '輸入資料者為空, 清空資料
                '    Dim sql As String = ""
                '    sql = "DELETE Stud_TrainCostD WHERE SOCID='" & Checkbox1.Value & "' "
                '    DbAccess.ExecuteNonQuery(sql)

                '    sql = "DELETE Stud_TrainCostT WHERE SOCID='" & Checkbox1.Value & "' AND Times='" & Me.ViewState("sd_05_022_Times") & "'"
                '    DbAccess.ExecuteNonQuery(sql)

                '    Checkbox1.Checked = False
                'End If

                If Errmsg2 <> "" Then Errmsg += "學員[" & labName.Text & "]:" & vbCrLf & Errmsg2 & vbCrLf
            End If

            If i > 30 Then
                Errmsg += "...等等警告錯誤過多" & vbCrLf
                Exit For
            End If
        Next
        If Errmsg <> "" Then
            Common.MessageBox(Me, Errmsg)
        Else
            Check_SaveDate = True
        End If
    End Function

    'Function Insert_Stud_TrainCost(Optional ByVal sType As Integer = 1)
    Sub Insert_Stud_TrainCost()

        'objconn
        Dim trans As SqlTransaction = DbAccess.BeginTrans(objconn)
        Try
            'Dim labTimes As Label = Me.DataGrid1.FindControl("labTimes")
            'Dim txtRate As TextBox = Me.DataGrid1.FindControl("txtRate")
            'UPDATE Stud_TrainCost
            Dim dr As DataRow = Nothing
            Dim da As SqlDataAdapter = Nothing
            Dim dt As DataTable = Nothing
            Dim sql As String = ""
            sql = "SELECT * FROM Stud_TrainCost WHERE OCID='" & Me.OCIDValue1.Value & "' AND Times='" & Me.ViewState("sd_05_022_Times") & "'"
            dt = DbAccess.GetDataTable(sql, da, trans)
            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
            Else
                dr = dt.NewRow
                dt.Rows.Add(dr)
                dr("OCID") = Me.OCIDValue1.Value
                dr("Times") = Me.ViewState("sd_05_022_Times")
            End If
            dr("Rate") = Me.hidRate.Value.Trim
            'If sType <> 1 Then
            dr("Applied") = "Y"
            'End If
            dr("ModifyAcct") = sm.UserInfo.UserID
            dr("ModifyDate") = Now
            DbAccess.UpdateDataTable(dt, da, trans)

            For Each Item As DataGridItem In DataGrid1.Items
                Dim Checkbox1 As HtmlInputCheckBox = Item.FindControl("Checkbox1")
                Dim IDNO As HtmlInputHidden = Item.FindControl("IDNO")
                Dim labName As Label = Item.FindControl("labName")
                Dim txtJobCost As TextBox = Item.FindControl("txtJobCost")
                Dim txtOtherJobCost As TextBox = Item.FindControl("txtOtherJobCost")
                Dim txtState As TextBox = Item.FindControl("txtState")
                Dim txtCost As TextBox = Item.FindControl("txtCost")

                If Checkbox1.Checked Then
                    'UPDATE Stud_TrainCostD(SOCID)
                    sql = "SELECT * FROM Stud_TrainCostD WHERE SOCID='" & Checkbox1.Value & "' "
                    dt = DbAccess.GetDataTable(sql, da, trans)
                    If dt.Rows.Count > 0 Then
                        dr = dt.Rows(0)
                    Else
                        dr = dt.NewRow
                        dt.Rows.Add(dr)
                        dr("SOCID") = Checkbox1.Value
                    End If
                    dr("IDNO") = IDNO.Value
                    dr("OCID") = Me.OCIDValue1.Value
                    dr("JobCost") = IIf(txtJobCost.Text.Trim = "", Convert.DBNull, txtJobCost.Text.Trim)
                    dr("OtherJobCost") = IIf(txtOtherJobCost.Text.Trim = "", Convert.DBNull, txtOtherJobCost.Text.Trim)
                    dr("State") = IIf(txtState.Text.Trim = "", Convert.DBNull, txtState.Text.Trim)
                    dr("ModifyAcct") = sm.UserInfo.UserID
                    dr("ModifyDate") = Now
                    DbAccess.UpdateDataTable(dt, da, trans)

                    'UPDATE Stud_TrainCostT [SOCID] , [Times]
                    If txtCost.Text.Trim <> "" Then
                        sql = "SELECT * FROM Stud_TrainCostT WHERE SOCID='" & Checkbox1.Value & "' AND Times='" & Me.ViewState("sd_05_022_Times") & "'"
                        dt = DbAccess.GetDataTable(sql, da, trans)
                        If dt.Rows.Count > 0 Then
                            dr = dt.Rows(0)
                        Else
                            dr = dt.NewRow
                            dt.Rows.Add(dr)
                            dr("SOCID") = Checkbox1.Value
                            dr("Times") = Me.ViewState("sd_05_022_Times")
                        End If
                        dr("Cost") = IIf(txtCost.Text.Trim = "", Convert.DBNull, txtCost.Text.Trim)
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now
                        DbAccess.UpdateDataTable(dt, da, trans)
                    Else
                        sql = "DELETE Stud_TrainCostT WHERE SOCID='" & Checkbox1.Value & "' AND Times='" & Me.ViewState("sd_05_022_Times") & "'"
                        DbAccess.ExecuteNonQuery(sql, trans)
                    End If

                End If

            Next

            DbAccess.CommitTrans(trans)
        Catch ex As Exception
            DbAccess.RollbackTrans(trans)
            DbAccess.CloseDbConn(objconn)
            Throw ex
        End Try

        Common.MessageBox(Me, "儲存成功")
    End Sub

    Private Sub DataGrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                e.Item.Cells(2).Text = drv("Rate").ToString + "%"
                'Case ListItemType.Item, ListItemType.AlternatingItem
                'Dim drv As DataRowView = e.Item.DataItem
                'If e.Item.ItemType = ListItemType.Item Then e.Item.CssClass = "SD_TD2"
        End Select
    End Sub

    Private Sub DataGrid2_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid2.ItemCommand
        Select Case e.CommandName
            Case "edit"
                Dim times As Integer = 0
                If e.Item.Cells(1).Text <> "" Then
                    times = Val(e.Item.Cells(1).Text)
                End If
                Call GetStudentData2(times)
        End Select

    End Sub

    Private Sub btn_add_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btn_add.Click
        Call GetStudentData2(99)
    End Sub

    Private Sub print_btn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles print_btn.Click
        Dim Years As Integer
        Years = Me.sm.UserInfo.Years - 1911
        ReportQuery.PrintReport(Me, "Report", "SD_05_022", "OCID=" & Me.OCIDValue1.Value.ToString & "&Years=" & Years)
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '判斷機構是否只有一個班級
        Dim dr As DataRow
        dr = TIMS.GET_OnlyOne_OCID(RIDValue.Value)
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        TableShowData.Visible = False
        DataGrid2.Visible = False
        print_btn.Visible = False
        msg.Text = ""
        'msg.Visible = False
        If Not dr Is Nothing Then
            If dr("total") = "1" Then '如果只有一個班級
                TMID1.Text = dr("trainname")
                OCID1.Text = dr("classname")
                TMIDValue1.Value = dr("trainid")
                OCIDValue1.Value = dr("ocid")
                TableShowData.Visible = False
                DataGrid2.Visible = False
                print_btn.Visible = False
                'msg.Visible = False
            End If
        End If
    End Sub

End Class
