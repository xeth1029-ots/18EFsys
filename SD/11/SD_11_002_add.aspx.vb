Partial Class SD_11_002_add
    Inherits AuthBasePage

    'Dim sqlAdapter As SqlDataAdapter
    Dim objconn As SqlConnection = Nothing
    Dim stud_table As DataTable = Nothing
    Dim sql As String = ""
    'Dim FunDr As DataRow = Nothing

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁

        '檢查Session是否存在-------------------------- Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        Dim dt As DataTable
        ProcessType.Value = Request("ProcessType")
        RadioButtonList1_2.Attributes("onclick") = "disable_radio1();"
        RadioButtonList1_3.Attributes("onclick") = "disable_radio1();"
        RadioButtonList2_10.Attributes("onclick") = "disable_radio1();"
        'Button2.Attributes("onclick") = "history.go(-1);"

#Region "(No Use)"

        'If sm.UserInfo.RoleID <> 0 Then
        '    If sm.UserInfo.FunDt Is Nothing Then
        '        Common.RespWrite(Me, "<script>alert('Session過期');</script>")
        '        Common.RespWrite(Me, "<script>top.location.href='../../logout.aspx';</script>")
        '    Else
        '        Dim FunDt As DataTable = sm.UserInfo.FunDt
        '        Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & Request("ID") & "'")
        '        Re_ID.Value = Request("ID")
        '        If FunDrArray.Length = 0 Then
        '            Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
        '            Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
        '        Else
        '            FunDr = FunDrArray(0)
        '            If ProcessType.Value = "Update" Then
        '                If FunDr("Mod") = "0" And FunDr("Del") = "0" Then
        '                    Button1.Enabled = False
        '                Else
        '                    Button1.Enabled = True
        '                End If
        '            ElseIf ProcessType.Value = "Insert" Or ProcessType.Value = "Next" Then
        '                If FunDr("Adds") = "1" Then
        '                    Button1.Enabled = True
        '                Else
        '                    Button1.Enabled = False
        '                End If
        '            End If

        '        End If
        '    End If
        'End If

#End Region

        If Not IsPostBack Then
            If Session("QuestionarySearchStr2") IsNot Nothing Then ViewState("QuestionarySearchStr2") = Session("QuestionarySearchStr2")
            'Session("QuestionarySearchStr2") = Nothing
            If ProcessType.Value <> "Print" Then
                Re_OCID.Value = TIMS.ClearSQM(Request("ocid"))
                Re_Studentid.Value = TIMS.ClearSQM(Request("Stuedntid"))

                Dim parms As New Hashtable
                parms.Add("OCID", Val(Re_OCID.Value))
                parms.Add("studentid", Re_Studentid.Value)
                Dim sqlstr As String = " SELECT b.studentid, c.name, b.StudStatus, b.RejectTDate1, b.RejectTDate2 "
                sqlstr += " FROM class_classinfo a "
                sqlstr += " JOIN class_studentsofclass b ON a.ocid = b.ocid "
                sqlstr += " JOIN stud_studentinfo c ON b.sid = c.sid "
                sqlstr += " WHERE b.OCID =@OCID AND b.studentid =@studentid "
                Dim row As DataRow = DbAccess.GetOneRow(sqlstr, objconn, parms)
                Me.Label_Name.Text = row("name")
                Me.Label_Stud.Text = row("studentid")
                Me.Label_Status.Text = TIMS.GET_STUDSTATUS_N23(row("StudStatus"), row("RejectTDate1"), row("RejectTDate2"))

                next_but.Style("display") = "" '"inline"
                'Button2.Style("display") = "" '"inline"
                Button1.Style("display") = "" '"inline"
                'Else
                '    next_but.Style("display") = "none"
                '    Button2.Style("display") = "none"
                '    Button1.Style("display") = "none"
            End If

            If ProcessType.Value = "del" Then
                Dim parms_d As New Hashtable
                parms_d.Add("OCID", Val(Re_OCID.Value))
                parms_d.Add("StudID", Re_Studentid.Value)
                Dim sqlstrdel As String = " DELETE Stud_QuestionEpt WHERE OCID =@OCID AND StudID =@StudID "
                DbAccess.ExecuteNonQuery(sqlstrdel, objconn, parms_d)
                Dim sqlstrdel2 As String = " DELETE Stud_QuestionEpt2_3 WHERE OCID =@OCID AND StudID =@StudID "
                DbAccess.ExecuteNonQuery(sqlstrdel2, objconn, parms_d)
                Dim sqlstrdel3 As String = " DELETE Stud_QuestionEpt3_4 WHERE OCID =@OCID AND StudID =@StudID "
                DbAccess.ExecuteNonQuery(sqlstrdel3, objconn, parms_d)
                Dim sqlstrdel4 As String = " DELETE Org_StudRecord WHERE OCID =@OCID AND StudID =@StudID "
                DbAccess.ExecuteNonQuery(sqlstrdel4, objconn, parms_d)
            End If
            If ProcessType.Value = "check" Then
                Button1.Enabled = False

                Dim parms_ls As New Hashtable
                parms_ls.Add("OCID", Val(Re_OCID.Value))
                parms_ls.Add("StudID", Re_Studentid.Value)
                Dim sqlstr_list As String = " SELECT * FROM Stud_QuestionEpt WHERE OCID =@OCID AND StudID =@StudID "
                Dim row_list As DataRow = DbAccess.GetOneRow(sqlstr_list, objconn, parms_ls)

                Me.RadioButtonList1_1.SelectedValue = row_list("Q1_1")
                Me.RadioButtonList1_2.SelectedValue = row_list("Q1_2")
                Me.RadioButtonList1_3.SelectedValue = row_list("Q1_3")
                If Convert.IsDBNull(row_list("Q1_4")) Then
                Else
                    Me.RadioButtonList1_4.SelectedValue = row_list("Q1_4")
                End If
                If Convert.IsDBNull(row_list("Q1_5")) Then
                    Q1_5Other.Text = ""
                Else
                    Me.RadioButtonList1_5.SelectedValue = row_list("Q1_5")
                    If row_list("Q1_5") = "97" Then
                        Q1_5Other.Text = row_list("Q1_5Other")
                    Else
                        Q1_5Other.Text = ""
                    End If
                End If
                If Convert.IsDBNull(row_list("Q1_6")) Then
                    Q1_6Other.Text = ""
                Else
                    Me.RadioButtonList1_6.SelectedValue = row_list("Q1_6")
                    If row_list("Q1_6") = "97" Then
                        Q1_6Other.Text = row_list("Q1_6Other")
                    Else
                        Q1_6Other.Text = ""
                    End If
                End If
                If Convert.IsDBNull(row_list("Q1_7")) Then
                Else
                    Me.RadioButtonList1_7.SelectedValue = row_list("Q1_7")
                End If
                If Convert.IsDBNull(row_list("Q1_8")) Then
                    Q1_8Other.Text = ""
                Else
                    Me.RadioButtonList1_8.SelectedValue = row_list("Q1_8")
                    If row_list("Q1_8") = "97" Then
                        Q1_8Other.Text = row_list("Q1_8Other")
                    Else
                        Q1_8Other.Text = ""
                    End If
                End If
                '第二部分
                If Convert.IsDBNull(("Q2_1")) Then
                Else
                    RadioButtonList2_1.SelectedValue = row_list("Q2_1")
                End If
                If Convert.IsDBNull(row_list("Q2_2")) Then
                Else
                    RadioButtonList2_2.SelectedValue = row_list("Q2_2")
                End If
                If Convert.IsDBNull(row_list("Q2_4")) Then
                Else
                    RadioButtonList2_4.SelectedValue = row_list("Q2_4")
                End If
                If Convert.IsDBNull(row_list("Q2_5")) Then
                Else
                    RadioButtonList2_5.SelectedValue = row_list("Q2_5")
                End If
                If Convert.IsDBNull(row_list("Q2_6")) Then
                Else
                    RadioButtonList2_6.SelectedValue = row_list("Q2_6")
                End If
                If Convert.IsDBNull(row_list("Q2_7")) Then
                Else
                    RadioButtonList2_7.SelectedValue = row_list("Q2_7")
                End If
                If Convert.IsDBNull(row_list("Q2_8")) Then
                Else
                    RadioButtonList2_8.SelectedValue = row_list("Q2_8")
                End If
                If Convert.IsDBNull(row_list("Q2_9")) Then
                Else
                    RadioButtonList2_9.SelectedValue = row_list("Q2_9")
                End If
                If Convert.IsDBNull(row_list("Q2_10")) Then
                Else
                    RadioButtonList2_10.SelectedValue = row_list("Q2_10")
                End If
                If Convert.IsDBNull(row_list("Q2_11")) Then
                    Q2_11Other.Text = ""
                Else
                    RadioButtonList2_11.SelectedValue = row_list("Q2_11")
                    If row_list("Q2_11") = "97" Then
                        Q2_11Other.Text = row_list("Q2_11Other")
                    Else
                        Q2_11Other.Text = ""
                    End If
                End If
                '第三部分
                If Convert.IsDBNull(row_list("Q3_1")) Then
                Else
                    RadioButtonList3_1.SelectedValue = row_list("Q3_1")
                End If
                If Convert.IsDBNull(row_list("Q3_2")) Then
                Else
                    RadioButtonList3_2.SelectedValue = row_list("Q3_2")
                End If
                If Convert.IsDBNull(row_list("Q3_3")) Then
                Else
                    RadioButtonList3_3.SelectedValue = row_list("Q3_3")
                End If
                If Convert.IsDBNull(row_list("Q3_5")) Then
                Else
                    Q3_5.Text = row_list("Q3_5")
                End If

                'Stud_QuestionEpt2_3
                'Dim i, j As Integer
                Dim parms As New Hashtable
                parms.Add("OCID", Val(Re_OCID.Value))
                parms.Add("StudID", Re_Studentid.Value)
                sql = " SELECT * FROM Stud_QuestionEpt2_3 WHERE OCID =@OCID AND StudID =@StudID "
                dt = DbAccess.GetDataTable(sql, objconn, parms)
                For i As Integer = 0 To dt.Rows.Count - 1
                    For j As Integer = 0 To Me.CheckBoxList2_3.Items.Count - 1
                        If CheckBoxList2_3.Items(j).Value = dt.Rows(i).Item("Q2_3").ToString Then CheckBoxList2_3.Items(j).Selected = True
                    Next
                    If dt.Rows(i).Item("Q2_3").ToString = "97" Then
                        'dt.Rows(i).Item("Q2_3").ToString <> ""
                        Q2_3Other.Text = dt.Rows(i).Item("Q2_3Other").ToString
                    End If
                Next
                'Stud_QuestionEpt3_4
                'Dim parms As New Hashtable
                'parms.Add("OCID", Val(Re_OCID.Value))
                'parms.Add("StudID", Re_Studentid.Value)
                sql = " SELECT * FROM Stud_QuestionEpt3_4 WHERE OCID =@OCID AND StudID =@StudID "
                dt = DbAccess.GetDataTable(sql, objconn, parms)
                For i As Integer = 0 To dt.Rows.Count - 1
                    For j As Integer = 0 To Me.CheckBoxList3_4.Items.Count - 1
                        If CheckBoxList3_4.Items(j).Value = dt.Rows(i).Item("Q3_4").ToString Then CheckBoxList3_4.Items(j).Selected = True
                    Next
                    If dt.Rows(i).Item("Q3_4").ToString = "97" Then
                        'dt.Rows(i).Item("Q2_3").ToString <> ""
                        Q3_4Other.Text = dt.Rows(i).Item("Q3_4Other").ToString
                    End If
                Next
            End If
            If ProcessType.Value = "Next" Then check_next()
            If ProcessType.Value = "Print" Then Me.RegisterStartupScript("clientScript", "<script language=""javascript"">printDoc();history.back(1);</script>")
#Region "(目前不使用)"

            ''========== (若此頁面以開新視窗的方式開啟,需做額外設定，by:20180913)
            'If Not IsPostBack Then
            '    If ProcessType.Value = "Print" Then divPage.Style.Add("max-height", "880px")
            'End If
            ''================================================= End

#End Region
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Dim objTrans As SqlTransaction = Nothing
        'Stud_QuestionEpt;學員訓練成效調查檔
        Dim dr1 As DataRow = Nothing
        Dim dt1 As DataTable = Nothing
        Dim da1 As SqlDataAdapter = Nothing
        'Org_StudRecord
        Dim dr2 As DataRow = Nothing
        Dim dt2 As DataTable = Nothing
        Dim da2 As SqlDataAdapter = Nothing
        Dim sqlstr1 As String = ""
        Dim sqlstr2 As String = ""
        Dim cmd As New SqlCommand
        Dim cmd2 As New SqlCommand
        If Session("QuestionarySearchStr2") Is Nothing AndAlso ViewState("QuestionarySearchStr2") IsNot Nothing Then Session("QuestionarySearchStr2") = Me.ViewState("QuestionarySearchStr2")
        'Session("QuestionarySearchStr2") = Me.ViewState("QuestionarySearchStr2")
        Call TIMS.OpenDbConn(objconn)
        Dim objTrans As SqlTransaction = DbAccess.BeginTrans(objconn)
        Try
            'If objconn.State = ConnectionState.Closed Then objconn.Open()
            'Stud_QuestionEpt
            'sqlstr1 = "select * from Stud_QuestionEpt where 1<>1"
            sqlstr1 = " SELECT * FROM Stud_QuestionEpt WHERE OCID = '" & Re_OCID.Value & "' AND StudID = '" & Re_Studentid.Value & "' "
            dt1 = DbAccess.GetDataTable(sqlstr1, da1, objTrans)
            If dt1.Rows.Count = 0 Then dr1 = dt1.NewRow Else dr1 = dt1.Rows(0)
            'Org_StudRecord
            'sqlstr2 = "select * from Org_StudRecord where 1<>1"
            sqlstr2 = " SELECT * FROM Org_StudRecord WHERE OCID = '" & Re_OCID.Value & "' AND StudID = '" & Re_Studentid.Value & "' "
            dt2 = DbAccess.GetDataTable(sqlstr2, da2, objTrans)
            If dt2.Rows.Count = 0 Then dr2 = dt2.NewRow Else dr2 = dt2.Rows(0)
            'Stud_QuestionEpt
            dr1("OCID") = Re_OCID.Value
            dr1("StudID") = Re_Studentid.Value
            dr1("FillFormDate") = Now()
            '第一部分
            dr1("Q1_1") = Me.RadioButtonList1_1.SelectedValue
            dr1("Q1_2") = Me.RadioButtonList1_2.SelectedValue
            dr1("Q1_3") = Me.RadioButtonList1_3.SelectedValue
            If Me.RadioButtonList1_4.SelectedValue <> "" And Me.RadioButtonList1_4.Enabled = True Then
                dr1("Q1_4") = Me.RadioButtonList1_4.SelectedValue
            Else
                dr1("Q1_4") = Convert.DBNull
            End If
            If Me.RadioButtonList1_5.SelectedValue <> "" And Me.RadioButtonList1_5.Enabled = True Then
                dr1("Q1_5") = Me.RadioButtonList1_5.SelectedValue
                If Me.RadioButtonList1_5.SelectedValue = "97" Then
                    dr1("Q1_5Other") = Q1_5Other.Text
                Else
                    dr1("Q1_5Other") = Convert.DBNull
                End If
            Else
                dr1("Q1_5") = Convert.DBNull
                dr1("Q1_5Other") = Convert.DBNull
            End If
            If Me.RadioButtonList1_6.SelectedValue <> "" And Me.RadioButtonList1_6.Enabled = True Then
                dr1("Q1_6") = Me.RadioButtonList1_6.SelectedValue
                If Me.RadioButtonList1_6.SelectedValue = "97" Then
                    dr1("Q1_6Other") = Q1_6Other.Text
                Else
                    dr1("Q1_6Other") = Convert.DBNull
                End If
            Else
                dr1("Q1_6") = Convert.DBNull
                dr1("Q1_6Other") = Convert.DBNull
            End If
            If Me.RadioButtonList1_7.SelectedValue <> "" And Me.RadioButtonList1_7.Enabled = True Then
                dr1("Q1_7") = Me.RadioButtonList1_7.SelectedValue
            Else
                dr1("Q1_7") = Convert.DBNull
            End If
            If Me.RadioButtonList1_8.SelectedValue <> "" And Me.RadioButtonList1_8.Enabled = True Then
                dr1("Q1_8") = Me.RadioButtonList1_8.SelectedValue
                If Me.RadioButtonList1_8.SelectedValue = "97" Then
                    dr1("Q1_8Other") = Q1_8Other.Text
                Else
                    dr1("Q1_8Other") = Convert.DBNull
                End If
            Else
                dr1("Q1_8Other") = Convert.DBNull
                dr1("Q1_8") = Convert.DBNull
            End If
            '第二部分
            If Me.RadioButtonList2_1.SelectedValue <> "" Then
                dr1("Q2_1") = RadioButtonList2_1.SelectedValue
            Else
                dr1("Q2_1") = Convert.DBNull
            End If
            If Me.RadioButtonList2_2.SelectedValue <> "" Then
                dr1("Q2_2") = RadioButtonList2_2.SelectedValue
            Else
                dr1("Q2_2") = Convert.DBNull
            End If
            If Me.RadioButtonList2_4.SelectedValue <> "" Then
                dr1("Q2_4") = RadioButtonList2_4.SelectedValue
            Else
                dr1("Q2_4") = Convert.DBNull
            End If
            If Me.RadioButtonList2_5.SelectedValue <> "" Then
                dr1("Q2_5") = RadioButtonList2_5.SelectedValue
            Else
                dr1("Q2_5") = Convert.DBNull
            End If
            If Me.RadioButtonList2_6.SelectedValue <> "" Then
                dr1("Q2_6") = RadioButtonList2_6.SelectedValue
            Else
                dr1("Q2_6") = Convert.DBNull
            End If
            If Me.RadioButtonList2_7.SelectedValue <> "" AndAlso RadioButtonList2_7.Enabled = True Then
                dr1("Q2_7") = RadioButtonList2_7.SelectedValue
            Else
                dr1("Q2_7") = Convert.DBNull
            End If
            If Me.RadioButtonList2_8.SelectedValue <> "" AndAlso RadioButtonList2_8.Enabled = True Then
                dr1("Q2_8") = RadioButtonList2_8.SelectedValue
            Else
                dr1("Q2_8") = Convert.DBNull
            End If
            If Me.RadioButtonList2_9.SelectedValue <> "" AndAlso RadioButtonList2_9.Enabled = True Then
                dr1("Q2_9") = RadioButtonList2_9.SelectedValue
            Else
                dr1("Q2_9") = Convert.DBNull
            End If
            If Me.RadioButtonList2_10.SelectedValue <> "" AndAlso RadioButtonList2_10.Enabled = True Then
                dr1("Q2_10") = RadioButtonList2_10.SelectedValue
            Else
                dr1("Q2_10") = Convert.DBNull
            End If
            If Me.RadioButtonList2_11.SelectedValue <> "" AndAlso RadioButtonList2_11.Enabled = True Then
                dr1("Q2_11") = RadioButtonList2_11.SelectedValue
                If RadioButtonList2_11.SelectedValue = "97" Then
                    dr1("Q2_11Other") = Q2_11Other.Text
                Else
                    dr1("Q2_11Other") = Convert.DBNull
                End If
            Else
                dr1("Q2_11") = Convert.DBNull
                dr1("Q2_11Other") = Convert.DBNull
            End If
            '第三部分
            If Me.RadioButtonList3_1.SelectedValue <> "" Then
                dr1("Q3_1") = RadioButtonList3_1.SelectedValue
            Else
                dr1("Q3_1") = Convert.DBNull
            End If
            If Me.RadioButtonList3_2.SelectedValue <> "" Then
                dr1("Q3_2") = RadioButtonList3_2.SelectedValue
            Else
                dr1("Q3_2") = Convert.DBNull
            End If
            If Me.RadioButtonList3_3.SelectedValue <> "" Then
                dr1("Q3_3") = RadioButtonList3_3.SelectedValue
            Else
                dr1("Q3_3") = Convert.DBNull
            End If
            If Q3_5.Text <> "" Then
                dr1("Q3_5") = Q3_5.Text
            Else
                dr1("Q3_5") = Convert.DBNull
            End If
            dr1("ModifyAcct") = sm.UserInfo.UserID
            dr1("ModifyDate") = Now()
            If dt1.Rows.Count = 0 Then dt1.Rows.Add(dr1)
            DbAccess.UpdateDataTable(dt1, da1, objTrans)
            'Org_StudRecord
            dr2("OCID") = Re_OCID.Value
            dr2("StudID") = Re_Studentid.Value
            dr2("FillFormDate") = Now()
            dr2("RID") = sm.UserInfo.RID
            If Me.RadioButtonList2_1.SelectedValue <> "" Then
                Select Case RadioButtonList2_1.SelectedValue
                    Case 1
                        dr2("Point1") = "5"
                    Case 2
                        dr2("Point1") = "4"
                    Case 3
                        dr2("Point1") = "3"
                    Case 4
                        dr2("Point1") = "2"
                    Case 5
                        dr2("Point1") = "1"
                End Select
            Else
                dr2("Point1") = Convert.DBNull
            End If
            If Me.RadioButtonList3_1.SelectedValue <> "" Then
                Select Case RadioButtonList3_1.SelectedValue
                    Case 1
                        dr2("Point2") = "5"
                    Case 2
                        dr2("Point2") = "4"
                    Case 3
                        dr2("Point2") = "3"
                    Case 4
                        dr2("Point2") = "2"
                    Case 5
                        dr2("Point2") = "1"
                End Select
            Else
                dr2("Point2") = Convert.DBNull
            End If
            If Me.RadioButtonList3_2.SelectedValue <> "" Then
                Select Case RadioButtonList3_2.SelectedValue
                    Case 1
                        dr2("Point3") = "5"
                    Case 2
                        dr2("Point3") = "4"
                    Case 3
                        dr2("Point3") = "3"
                    Case 4
                        dr2("Point3") = "2"
                    Case 5
                        dr2("Point3") = "1"
                End Select
            Else
                dr2("Point3") = Convert.DBNull
            End If
            If Me.RadioButtonList1_1.SelectedValue <> "" Then
                Select Case RadioButtonList1_1.SelectedValue
                    Case 1
                        dr2("Point4") = "3"
                    Case 2
                        dr2("Point4") = "2"
                    Case 3
                        dr2("Point4") = "1"
                End Select
            Else
                dr2("Point4") = Convert.DBNull
            End If
            If Me.RadioButtonList1_2.SelectedValue <> "" Then
                Select Case RadioButtonList1_2.SelectedValue
                    Case 1
                        dr2("Point5") = "4"
                    Case 2
                        dr2("Point5") = "3"
                    Case 3
                        dr2("Point5") = "2"
                    Case 4
                        dr2("Point5") = "1"
                End Select
            Else
                dr2("Point5") = Convert.DBNull
            End If
            dr2("ModifyAcct") = sm.UserInfo.UserID
            dr2("ModifyDate") = Now()
            If dt2.Rows.Count = 0 Then dt2.Rows.Add(dr2)
            DbAccess.UpdateDataTable(dt2, da2, objTrans)
        Catch ex As Exception
            DbAccess.RollbackTrans(objTrans)
            Common.MessageBox(Me, "儲存失敗，有必填答案未填寫，請重新確認答案後，再次儲存!!謝謝")
            Exit Sub
            'Throw ex
        End Try
        Try
            'Stud_QuestionEpt2_3
            sql = " DELETE Stud_QuestionEpt2_3 where OCID =@OCID AND StudID =@StudID"
            cmd = New SqlCommand(sql, objconn, objTrans)
            With cmd
                .Parameters.Clear()
                .Parameters.Add("OCID", SqlDbType.BigInt).Value = Val(Re_OCID.Value)
                .Parameters.Add("StudID", SqlDbType.VarChar).Value = Re_Studentid.Value
                .ExecuteNonQuery()
            End With
            'Dim i As Integer
            For i As Integer = 0 To CheckBoxList2_3.Items.Count - 1
                If CheckBoxList2_3.Items(i).Selected = True Then
                    sql = " INSERT INTO Stud_QuestionEpt2_3 (OCID, StudID, Q2_3, Q2_3Other) VALUES (@OCID, @StudID,@Q2_3, @Q2_3Other)"
                    cmd2 = New SqlCommand(sql, objconn, objTrans)
                    With cmd2
                        .Parameters.Clear()
                        .Parameters.Add("OCID", SqlDbType.BigInt).Value = Val(Re_OCID.Value)
                        .Parameters.Add("StudID", SqlDbType.NVarChar).Value = Re_Studentid.Value
                        .Parameters.Add("Q2_3", SqlDbType.BigInt).Value = Val(CheckBoxList2_3.Items(i).Value)
                        .Parameters.Add("Q2_3Other", SqlDbType.NVarChar).Value = If(CheckBoxList2_3.Items(i).Value = "97", Q2_3Other.Text, "")
                        .ExecuteNonQuery()
                    End With
                End If
            Next
            'Stud_QuestionEpt3_4
            sql = " DELETE Stud_QuestionEpt3_4 WHERE OCID =@OCID AND StudID =@StudID"
            cmd = New SqlCommand(sql, objconn, objTrans)
            With cmd
                .Parameters.Clear()
                .Parameters.Add("OCID", SqlDbType.BigInt).Value = Val(Re_OCID.Value)
                .Parameters.Add("StudID", SqlDbType.VarChar).Value = Re_Studentid.Value
                .ExecuteNonQuery()
            End With
            For i As Integer = 0 To CheckBoxList3_4.Items.Count - 1
                If CheckBoxList3_4.Items(i).Selected = True Then
                    sql = " INSERT INTO Stud_QuestionEpt3_4 (OCID,StudID,Q3_4,Q3_4Other) VALUES (@OCID, @StudID,@Q3_4, @Q3_4Other)"
                    cmd2 = New SqlCommand(sql, objconn, objTrans)
                    With cmd2
                        .Parameters.Clear()
                        .Parameters.Add("OCID", SqlDbType.BigInt).Value = Val(Re_OCID.Value)
                        .Parameters.Add("StudID", SqlDbType.NVarChar).Value = Re_Studentid.Value
                        .Parameters.Add("Q3_4", SqlDbType.BigInt).Value = Val(CheckBoxList3_4.Items(i).Value)
                        .Parameters.Add("Q3_4Other", SqlDbType.NVarChar).Value = If(CheckBoxList3_4.Items(i).Value = "97", Q3_4Other.Text, "")
                        .ExecuteNonQuery()
                    End With
                End If
            Next
            DbAccess.CommitTrans(objTrans)
        Catch ex As Exception
            DbAccess.RollbackTrans(objTrans)
            Common.MessageBox(Me, "儲存失敗，有必填答案未填寫，請重新確認答案後，再次儲存!!謝謝")
            Exit Sub
            'Throw ex
        Finally
            If objconn.State = ConnectionState.Open Then objconn.Close()
        End Try
        Common.AddClientScript(Page, "insert_next();")
    End Sub

    Private Sub CustomValidator1_ServerValidate(ByVal source As System.Object, ByVal args As System.Web.UI.WebControls.ServerValidateEventArgs) Handles CustomValidator1.ServerValidate
        If RadioButtonList1_2.SelectedValue = "4" And RadioButtonList1_3.SelectedValue = "2" Then
            args.IsValid = True '通過驗證
        Else
            args.IsValid = False
            source.errormessage = ""
            If RadioButtonList1_4.SelectedValue = "" Then source.errormessage &= "請選擇第一部分的問題四" & vbCrLf
            If RadioButtonList1_5.SelectedValue = "" Then source.errormessage &= "請選擇第一部分的問題五" & vbCrLf
            If RadioButtonList1_6.SelectedValue = "" Then source.errormessage &= "請選擇第一部分的問題六" & vbCrLf
            If RadioButtonList1_7.SelectedValue = "" Then source.errormessage &= "請選擇第一部分的問題七" & vbCrLf
        End If
        If RadioButtonList1_2.SelectedValue = "4" And RadioButtonList1_3.SelectedValue = "1" Then
            args.IsValid = False '沒通過驗證
            If RadioButtonList1_8.SelectedValue = "" Then source.errormessage &= "請選擇第一部分的問題八" & vbCrLf
        End If
        If RadioButtonList1_5.SelectedValue = "97" And Q1_5Other.Text = "" Then
            args.IsValid = False
            source.errormessage &= "請輸入第一部分問題五中的其他選項" & vbCrLf
        End If
        If RadioButtonList1_6.SelectedValue = "97" And Q1_6Other.Text = "" Then
            args.IsValid = False
            source.errormessage &= "請輸入第一部分問題六中的其他選項" & vbCrLf
        End If
        If RadioButtonList1_8.SelectedValue = "97" And Q1_8Other.Text = "" Then
            args.IsValid = False
            source.errormessage &= "請輸入第一部分問題八中的其他選項" & vbCrLf
        End If
        If CheckBoxList2_3.SelectedValue = "97" And Q2_3Other.Text = "" Then
            args.IsValid = False
            source.errormessage &= "'請輸入第二部分問題三中的其他選項" & vbCrLf
        End If
        If RadioButtonList2_11.SelectedValue = "97" And Q2_11Other.Text = "" Then
            args.IsValid = False
            source.errormessage &= "請輸入第二部分問題十一中的其他選項" & vbCrLf
        End If
        If CheckBoxList3_4.SelectedValue = "97" And Q3_4Other.Text = "" Then
            args.IsValid = False
            source.errormessage &= "請輸入第三部分問題四中的其他選項" & vbCrLf
        End If
    End Sub

    Private Sub Page_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.PreRender
        Dim strMessage As String = ""
        For Each obj As WebControls.BaseValidator In Page.Validators
            If obj.IsValid = False Then strMessage &= obj.ErrorMessage & vbCrLf
        Next
        If strMessage <> "" Then Common.MessageBox(Page, strMessage)
    End Sub

#Region "(No Use)"

    'Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
    '    Session("QuestionarySearchStr2") = Me.ViewState("QuestionarySearchStr2")
    '    TIMS.Utl_Redirect1(Me, "SD_11_002.aspx?ProcessType=Back&ID=" & Re_ID.Value & "")
    'End Sub

#End Region

    Private Sub check_last()
        If Session("QuestionarySearchStr2") Is Nothing AndAlso ViewState("QuestionarySearchStr2") IsNot Nothing Then Session("QuestionarySearchStr2") = Me.ViewState("QuestionarySearchStr2")
        'Session("QuestionarySearchStr2") = Me.ViewState("QuestionarySearchStr2")
        Dim strScript As String
        strScript = "<script language=""javascript"">" + vbCrLf
        strScript += "alert('已為此班級中最後一筆學員!!');" + vbCrLf
        strScript += "location.href ='SD_11_002.aspx?ProcessType=Back&ocid='+document.getElementById('Re_OCID').value+'&ID='+document.getElementById('Re_ID').value;" + vbCrLf
        strScript += "</script>"
        Page.RegisterStartupScript("", strScript)
    End Sub

    Function Get_dtStudent2(ByVal ocid As String, ByVal StudID As String) As DataTable
        Dim sql As String = ""
        sql = ""
        sql &= " SELECT b.studentid, b.StudStatus, c.name, b.OCID, b.RejectTDate1, b.RejectTDate2 "
        If StudID <> "" Then
            sql &= " FROM class_classinfo a "
            sql &= " JOIN class_studentsofclass b ON a.ocid = b.ocid "
            sql &= " JOIN stud_studentinfo c ON b.sid = c.sid "
            sql &= " JOIN Stud_QuestionEpt ee ON b.ocid = ee.ocid AND b.studentid = ee.studid "
            sql &= " WHERE 1=1 "
            sql &= " AND ee.StudID = @StudID "
            sql &= " AND a.ocid = @ocid "
            'sql &= " ORDER BY b.studentid "
        Else
            sql &= " FROM class_classinfo a "
            sql &= " JOIN class_studentsofclass b ON a.ocid = b.ocid "
            sql &= " JOIN stud_studentinfo c ON b.sid = c.sid "
            sql &= " WHERE 1=1 AND a.ocid = @ocid "
            sql &= " ORDER BY b.studentid "
        End If
        Dim sCmd As New SqlCommand(sql, objconn)
        TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable '= Nothing
        With sCmd
            .Parameters.Clear()
            If StudID <> "" Then
                .Parameters.Add("StudID", SqlDbType.VarChar).Value = StudID
                .Parameters.Add("ocid", SqlDbType.VarChar).Value = ocid
            Else
                .Parameters.Add("ocid", SqlDbType.VarChar).Value = ocid
            End If
            dt.Load(.ExecuteReader())
        End With
        'dt = DbAccess.GetDataTable(sqlstr, objconn)
        Return dt
    End Function

    Private Sub check_next()
        If Session("QuestionarySearchStr2") Is Nothing AndAlso ViewState("QuestionarySearchStr2") IsNot Nothing Then Session("QuestionarySearchStr2") = Me.ViewState("QuestionarySearchStr2")


#Region "(No Use)"

        'Dim Student_Table2 As DataTable
        'Student_Table2 = Session("DTable_Stuednt2")
        'Dim rows() As DataRow
        'rows = dtStudent2.Select("studentid > '" & Re_Studentid.Value & "'")

#End Region
        Dim dtStudent2 As DataTable = Get_dtStudent2(Re_OCID.Value, "")
        Dim ff3 As String = "studentid > '" & Re_Studentid.Value & "'"
        If dtStudent2.Select(ff3).Length = 0 Then
            Call check_last()
        Else
            For Each dr As DataRow In dtStudent2.Select(ff3)
                Dim dtEpt1 As DataTable = Get_dtStudent2(dr("OCID"), dr("studentid"))
                If dtEpt1.Rows.Count = 0 Then
                    Re_Studentid.Value = dr("studentid")
                    Re_OCID.Value = dr("OCID")
                    Me.Label_Name.Text = dr("name")
                    Me.Label_Stud.Text = dr("studentid")
                    'Session("QuestionarySearchStr2") = Me.ViewState("QuestionarySearchStr2")
                    TIMS.Utl_Redirect1(Me, "SD_11_002_add.aspx?ocid=" & Me.Re_OCID.Value & "&Stuedntid=" & Re_Studentid.Value & "&ID=" & Re_ID.Value & "")
                    Exit For
                End If
            Next
        End If
#Region "(No Use)"

        'If rows.Length = 0 Then
        '    Call check_last()
        'Else
        '    Dim dr As DataRow
        '    Dim i As Integer
        '    For i As Integer = 0 To dtStudent2.Select(ff3).Length - 1
        '        Dim dr As DataRow = dtStudent2.Select(ff3)(i)
        '        Dim sqlstr_list As String = "select * from Stud_QuestionEpt where OCID='" & dr("OCID") & "' and StudID= '" & dr("studentid") & "'"
        '        If DbAccess.GetCount(sqlstr_list) = 0 Then '沒有資料
        '        ElseIf rows.Length = 0 Or rows.Length = 1 Then
        '            check_last()
        '        ElseIf i = rows.Length - 1 Then
        '            check_last()
        '        End If
        '    Next
        'End If

#End Region
    End Sub

    Private Sub next_but_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles next_but.Click
        check_next()
    End Sub

    Protected Sub BtnBack1_Click(sender As Object, e As EventArgs) Handles BtnBack1.Click
        If Session("QuestionarySearchStr2") Is Nothing AndAlso ViewState("QuestionarySearchStr2") IsNot Nothing Then Session("QuestionarySearchStr2") = Me.ViewState("QuestionarySearchStr2")

        Dim s_FUNID As String = TIMS.Get_MRqID(Me)
        TIMS.Utl_Redirect1(Me, "SD_11_002.aspx?ProcessType=Back&ID=" & s_FUNID & "")
    End Sub
End Class