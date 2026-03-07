Partial Class SD_11_005_add09
    Inherits AuthBasePage

    'Stud_QuestionFin 產投
    Dim objconn As SqlConnection

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
        '檢查Session是否存在--------------------------End

        ProcessType.Value = Request("ProcessType")
#Region "(No Use)"

        'If sm.UserInfo.RoleID <> 0 Then
        'End If

        'If sm.UserInfo.FunDt Is Nothing Then
        '    Common.RespWrite(Me, "<script>alert('Session過期');</script>")
        '    Common.RespWrite(Me, "<script>top.location.href='../../logout.aspx';</script>")
        'Else
        '    Dim FunDt As DataTable = sm.UserInfo.FunDt
        '    Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & Request("ID") & "'")
        '    Re_ID.Value = Request("ID")
        '    If FunDrArray.Length = 0 Then
        '        Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
        '        Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
        '    Else
        '       'Dim FunDr As DataRow
        '        FunDr = FunDrArray(0)
        '        Select Case ProcessType.Value
        '            Case "Update"
        '                Button1.Enabled = True

        '                If FunDr("Mod") = "0" And FunDr("Del") = "0" Then
        '                    Button1.Enabled = False
        '                    TIMS.Tooltip(Button1, "(Mod.Del)無權限使用該功能", True)
        '                End If
        '            Case "Insert", "next"
        '                Button1.Enabled = False
        '                If FunDr("Adds") = "1" Then
        '                    Button1.Enabled = True
        '                    'next_but.Enabled = True
        '                Else
        '                    TIMS.Tooltip(Button1, "(Adds)無權限使用該功能", True)
        '                    'TIMS.Tooltip(next_but, "(Adds)無權限使用該功能", True)
        '                End If
        '        End Select
        '    End If
        'End If

#End Region

        If Not IsPostBack Then
            If Convert.ToString(Request("OCID")) = "" Then
                Common.RespWrite(Me, "<script>alert('傳入資料有誤，請重新操作該功能');</script>")
                Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
                Exit Sub
            End If
#Region "(No Use)"

            'If Convert.ToString(Request("SOCID")) = "" Then
            '    Common.RespWrite(Me, "<script>alert('傳入資料有誤，請重新操作該功能');</script>")
            '    Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
            '    Exit Sub
            'End If

#End Region
            Re_OCID.Value = Convert.ToString(Request("OCID"))
            Re_SOCID.Value = Convert.ToString(Request("SOCID")) '不一定有資料
            Call LoadCreateData1()
            SOCID.Attributes("onchange") = "if(this.selectedIndex==0) return false;"
            Button1.Attributes.Add("OnClick", "return ChkData();")
            Q1.Attributes("onclick") = "Q1Is4();"
        End If
    End Sub

    Sub LoadCreateData1()
        If Not Session("QuestionFinSearchStr") = Nothing Then
            Me.ViewState("QuestionFinSearchStr") = Session("QuestionFinSearchStr")
            Session("QuestionFinSearchStr") = Nothing
        End If

        Dim sql As String
        Dim dt As DataTable
#Region "(No Use)"

        'sql = "SELECT StudentID, case "
        'sql += "when len(a.StudentID)=12 then b.Name+'('+RIGHT(a.StudentID,3)+')' "
        'sql += "else b.Name+'('+RIGHT(a.StudentID,2)+')' "
        'sql += "end as Name, a.SOCID "
        'sql += "FROM (SELECT * FROM Class_StudentsOfClass WHERE OCID='" & Request("OCID") & "') a "
        'sql += "JOIN (SELECT * FROM Stud_StudentInfo) b ON a.SID=b.SID"

#End Region
        '**by Milor 20080423----start
        sql = ""
        sql += " SELECT a.StudentID, CASE WHEN LEN(a.StudentID)=12 THEN b.Name COLLATE CHINESE_TAIWAN_STROKE_CI_AS + ISNULL('(' + SUBSTRING(a.StudentID,LEN(a.StudentID)-3,3) COLLATE CHINESE_TAIWAN_STROKE_CI_AS + ')', '') ELSE b.Name COLLATE CHINESE_TAIWAN_STROKE_CI_AS + ISNULL('(' + SUBSTRING(a.StudentID,LEN(a.StudentID)-2,2) COLLATE CHINESE_TAIWAN_STROKE_CI_AS + ')', '') END AS Name, a.SOCID " & vbCrLf
        sql += " FROM Class_StudentsOfClass a " & vbCrLf
        sql += " JOIN Stud_StudentInfo b ON a.SID = b.SID " & vbCrLf
        sql += " JOIN Class_ClassInfo c ON c.OCID = a.OCID " & vbCrLf
        sql += " JOIN Org_OrgInfo d ON d.ComIDNO = c.ComIDNO " & vbCrLf
        sql += " WHERE a.OCID = @OCID "
        '**by Milor 20080423----end

        Dim parms As Hashtable = New Hashtable()
        parms.Add("OCID", Re_OCID.Value)
        dt = DbAccess.GetDataTable(sql, objconn, parms)
        If dt.Rows.Count = 0 Then
            Common.RespWrite(Me, "<script>alert('傳入資料有誤，請重新操作該功能');</script>")
            Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
            Exit Sub
        End If
        dt.DefaultView.Sort = "StudentID"
        With SOCID
            .DataSource = dt
            .DataTextField = "Name"
            .DataValueField = "SOCID"
            .DataBind()
            .Items.Insert(0, New ListItem("===請選擇===", ""))
        End With

        If Re_SOCID.Value <> "" Then
            Common.SetListItem(SOCID, Re_SOCID.Value)
            Select Case ProcessType.Value
                Case "del" '刪除
                    Dim sqlstrdel As String
                    sqlstrdel = "DELETE Stud_QuestionFin WHERE SOCID = '" & Re_SOCID.Value & "' "
                    DbAccess.ExecuteNonQuery(sqlstrdel, objconn)
                    Common.MessageBox(Me, "資料刪除完成", "sd_11_005.aspx?ID=" & Request("ID"))
                Case "check" '查詢
                    create(Re_SOCID.Value)
                    Button1.Enabled = False
                    next_but.Enabled = False
                    TIMS.Tooltip(Button1, "僅供查詢", True)
                    TIMS.Tooltip(next_but, "僅供查詢", True)
                Case "Edit" '修改
                    create(Re_SOCID.Value)
                    Button1.Enabled = True
                    next_but.Enabled = False
                    TIMS.Tooltip(next_but, "僅供修改", True)
                Case "next"
                    '儲存 自動呼叫下一筆
                    Call MoveNext()
            End Select
        End If
    End Sub

    '建立學員顯示及答案 (依產投)
    Private Sub create(ByVal StrSOCID As String)
        Dim sqlstr As String = ""
        sqlstr = ""
        sqlstr = " SELECT b.StudentID, c.name, b.StudStatus, b.RejectTDate1, b.RejectTDate2, d.OrgKind "
        sqlstr += " FROM class_classinfo a "
        sqlstr += " JOIN class_studentsofclass b ON a.ocid = b.ocid "
        sqlstr += " JOIN plan_planinfo p ON a.PlanID = p.PlanID AND a.comIDNO = p.comIDNO AND a.SeqNO = p.SeqNO "
        sqlstr += " JOIN stud_studentinfo c ON b.sid = c.sid "
        sqlstr += " JOIN Org_OrgInfo d ON d.ComIDNO = a.ComIDNO "
        sqlstr += " WHERE 1=1 "
        sqlstr += "    AND p.TPlanID IN (" & TIMS.Cst_TPlanID28_2 & ") "
        'sqlstr += "   AND b.OCID = '" & Re_OCID.Value & "' "
        sqlstr += "    AND b.SOCID = '" & StrSOCID & "' "
        Dim row As DataRow = DbAccess.GetOneRow(sqlstr, objconn)
        If row Is Nothing Then
            Common.RespWrite(Me, "<script>alert('傳入資料有誤，請重新操作該功能');</script>")
            Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
            Exit Sub
        End If
        Me.Label_Name.Text = Convert.ToString(row("name"))
        If Convert.ToString(row("StudentID")) <> "" Then
            Me.Label_Stud.Text = Convert.ToString(row("StudentID"))
        Else
            Me.Label_Stud.Text = "--"
        End If
        'Dim okind As String = Convert.ToString(row("OrgKind"))  '**by Milor 20080424
        Dim str As String = ""
        Select Case row("StudStatus").ToString
            Case "1"
                Me.Label_Status.Text = "在訓"
            Case "2"
                str += "離訓"
                str += "(" + row("RejectTDate1") + ")"
                Me.Label_Status.Text = str
            Case "3"
                str += "退訓"
                str += "(" + row("RejectTDate2") + ")"
                Me.Label_Status.Text = str
            Case "4"
                Me.Label_Status.Text = "續訓"
            Case "5"
                Me.Label_Status.Text = "結訓"
        End Select
        'Dim sqlstr As String
        sqlstr = " SELECT * FROM Stud_QuestionFin WHERE SOCID = '" & StrSOCID & "' "
        Dim row_list As DataRow = DbAccess.GetOneRow(sqlstr, objconn)
        If Not row_list Is Nothing Then
            If Convert.ToString(row_list("Q1")) <> "" Then Common.SetListItem(Q1, Convert.ToString(row_list("Q1")))
            'If Not IsDBNull(row_list("Q1")) Then Common.SetListItem(Q1, row_list("Q1"))
            If Convert.ToString(row_list("Q1")) = "4" Then
                Q2.Enabled = False
                Q3.Enabled = False
                Q4.Enabled = False
                Q5.Enabled = False
            Else
                If Not IsDBNull(row_list("Q2")) Then Common.SetListItem(Q2, row_list("Q2"))
                If Not IsDBNull(row_list("Q3")) Then Common.SetListItem(Q3, row_list("Q3"))
                If Not IsDBNull(row_list("Q4")) Then Common.SetListItem(Q4, row_list("Q4"))
                If Not IsDBNull(row_list("Q5")) Then Common.SetListItem(Q5, row_list("Q5"))
            End If
            'If Not IsDBNull(row_list("Q2")) Then Common.SetListItem(Q2, row_list("Q2"))
            'If Not IsDBNull(row_list("Q3")) Then Common.SetListItem(Q3, row_list("Q3"))
            'If Not IsDBNull(row_list("Q4")) Then Common.SetListItem(Q4, row_list("Q4"))
            'If Not IsDBNull(row_list("Q5")) Then Common.SetListItem(Q5, row_list("Q5"))
            If Not IsDBNull(row_list("Q6")) Then Common.SetListItem(Q6, row_list("Q6"))
            If Not IsDBNull(row_list("Q7")) Then Common.SetListItem(Q7, row_list("Q7"))
            If Not IsDBNull(row_list("Q8")) Then Common.SetListItem(Q8, row_list("Q8"))
            If Not IsDBNull(row_list("Q9_1_Note")) Then Q9_1_Note.Text = row_list("Q9_1_Note").ToString Else Q9_1_Note.Text = ""
            If Not IsDBNull(row_list("Q9_2_Note")) Then Q9_2_Note.Text = row_list("Q9_2_Note").ToString Else Q9_2_Note.Text = ""
            If Not IsDBNull(row_list("Q9_3_Note")) Then Q9_3_Note.Text = row_list("Q9_3_Note").ToString Else Q9_3_Note.Text = ""
            If Not IsDBNull(row_list("Q10_1_Note")) Then Q10_1_Note.Text = row_list("Q10_1_Note").ToString Else Q10_1_Note.Text = ""
            If Not IsDBNull(row_list("Q10_2_Note")) Then Q10_2_Note.Text = row_list("Q10_2_Note").ToString Else Q10_2_Note.Text = ""
            If Not IsDBNull(row_list("Q10_3_Note")) Then Q10_3_Note.Text = row_list("Q10_3_Note").ToString Else Q10_3_Note.Text = ""
            'If Not IsDBNull(row_list("BusName")) Then BusName.Text = row_list("BusName").ToString Else BusName.Text = ""
            'If Not IsDBNull(row_list("Q11")) Then Common.SetListItem(Q11, row_list("Q11"))
            'If Not IsDBNull(row_list("Q12")) Then Common.SetListItem(Q12, row_list("Q12"))
            'If Not IsDBNull(row_list("Q13")) Then Common.SetListItem(Q13, row_list("Q13"))
            'If Not IsDBNull(row_list("Q14")) Then Common.SetListItem(Q14, row_list("Q14"))
            'If Not IsDBNull(row_list("Q15")) Then Common.SetListItem(Q15, row_list("Q15"))
            'If Not IsDBNull(row_list("Q16")) Then Common.SetListItem(Q16, row_list("Q16"))
        Else
            Q1.SelectedIndex = -1
            Q2.SelectedIndex = -1
            Q3.SelectedIndex = -1
            Q4.SelectedIndex = -1
            Q5.SelectedIndex = -1
            Q6.SelectedIndex = -1
            Q7.SelectedIndex = -1
            Q8.SelectedIndex = -1
            Q9_1_Note.Text = ""
            Q9_2_Note.Text = ""
            Q9_3_Note.Text = ""
            Q10_1_Note.Text = ""
            Q10_2_Note.Text = ""
            Q10_3_Note.Text = ""
            'BusName.Text = ""
            'Q11.SelectedIndex = -1
            'Q12.SelectedIndex = -1
            'Q13.SelectedIndex = -1
            'Q14.SelectedIndex = -1
            'Q15.SelectedIndex = -1
            'Q16.SelectedIndex = -1
        End If
    End Sub

    '檢查是否為 最後一筆學員
    Private Sub check_last()
        Session("QuestionFinSearchStr") = Me.ViewState("QuestionFinSearchStr")
        Dim strScript As String
        strScript = "<script language=""javascript"">" + vbCrLf
        strScript += "alert('已為此班級中最後一筆學員!!');" + vbCrLf
        strScript += "location.href ='SD_11_005.aspx?ProcessType=Back&ID='+document.getElementById('Re_ID').value;" + vbCrLf
        strScript += "</script>"
        Page.RegisterStartupScript("", strScript)
    End Sub

    '回上頁
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Session("QuestionFinSearchStr") = Me.ViewState("QuestionFinSearchStr")
        'Response.Redirect("SD_11_005.aspx?ProcessType=Back&ID=" & Re_ID.Value & "")
    End Sub

    '不儲存 呼叫下一筆
    Private Sub next_but_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles next_but.Click
        Call MoveNext()
    End Sub

    '下一筆
    Private Sub MoveNext()
        If SOCID.Items.Count > 0 Then
            Dim NowIndex As Integer
            Dim MaxIndex As Integer
            MaxIndex = SOCID.Items.Count - 1
            NowIndex = SOCID.SelectedIndex
            If NowIndex = MaxIndex Then
                check_last()
            Else
                SOCID.SelectedIndex = NowIndex + 1
                Re_SOCID.Value = SOCID.SelectedValue
                create(SOCID.SelectedValue)
            End If
        End If
    End Sub

    Private Sub SOCID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SOCID.SelectedIndexChanged
        Re_SOCID.Value = SOCID.SelectedValue
        create(SOCID.SelectedValue)
    End Sub

    '儲存
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim sql As String = ""
        Dim flagI As Integer = 0
        Dim tConn As SqlConnection = Nothing
        tConn = DbAccess.GetConnection()
        Call TIMS.OpenDbConn(tConn)
        'TIMS.TestDbConn(Me, tConn, True)
        Session("QuestionFinSearchStr") = Me.ViewState("QuestionFinSearchStr")
        Dim objTrans As SqlTransaction
        objTrans = DbAccess.BeginTrans(tConn)
        sql = " SELECT * FROM Stud_QuestionFin WHERE SOCID = '" & SOCID.SelectedValue & "' "
        Dim sqlAdapter As SqlDataAdapter = Nothing
        Dim dtTable As DataTable
        dtTable = DbAccess.GetDataTable(sql, sqlAdapter, objTrans)
        Try
            Dim dr_row As DataRow
            If dtTable.Rows.Count = 0 Then
                flagI = 1
                dr_row = dtTable.NewRow
                dr_row("SOCID") = SOCID.SelectedValue
            Else
                dr_row = dtTable.Rows(0)
            End If
            If Q1.SelectedIndex = -1 Then dr_row("Q1") = Convert.DBNull Else dr_row("Q1") = Q1.SelectedValue
            If Q2.SelectedIndex = -1 Then dr_row("Q2") = Convert.DBNull Else dr_row("Q2") = Q2.SelectedValue
            If Q3.SelectedIndex = -1 Then dr_row("Q3") = Convert.DBNull Else dr_row("Q3") = Q3.SelectedValue
            If Q4.SelectedIndex = -1 Then dr_row("Q4") = Convert.DBNull Else dr_row("Q4") = Q4.SelectedValue
            If Q5.SelectedIndex = -1 Then dr_row("Q5") = Convert.DBNull Else dr_row("Q5") = Q5.SelectedValue
            If Q6.SelectedIndex = -1 Then dr_row("Q6") = Convert.DBNull Else dr_row("Q6") = Q6.SelectedValue
            If Q7.SelectedIndex = -1 Then dr_row("Q7") = Convert.DBNull Else dr_row("Q7") = Q7.SelectedValue
            If Q8.SelectedIndex = -1 Then dr_row("Q8") = Convert.DBNull Else dr_row("Q8") = Q8.SelectedValue
            If Q9_1_Note.Text = "" Then dr_row("Q9_1_Note") = Convert.DBNull Else dr_row("Q9_1_Note") = Q9_1_Note.Text
            If Q9_2_Note.Text = "" Then dr_row("Q9_2_Note") = Convert.DBNull Else dr_row("Q9_2_Note") = Q9_2_Note.Text
            If Q9_3_Note.Text = "" Then dr_row("Q9_3_Note") = Convert.DBNull Else dr_row("Q9_3_Note") = Q9_3_Note.Text
            If Q10_1_Note.Text = "" Then dr_row("Q10_1_Note") = Convert.DBNull Else dr_row("Q10_1_Note") = Q10_1_Note.Text
            If Q10_2_Note.Text = "" Then dr_row("Q10_2_Note") = Convert.DBNull Else dr_row("Q10_2_Note") = Q10_2_Note.Text
            If Q10_3_Note.Text = "" Then dr_row("Q10_3_Note") = Convert.DBNull Else dr_row("Q10_3_Note") = Q10_3_Note.Text
            'If BusName.Text = "" Then dr_row("BusName") = Convert.DBNull Else dr_row("BusName") = BusName.Text
            'If Q11.SelectedIndex = -1 Then dr_row("Q11") = Convert.DBNull Else dr_row("Q11") = Q11.SelectedValue
            'If Q12.SelectedIndex = -1 Then dr_row("Q12") = Convert.DBNull Else dr_row("Q12") = Q12.SelectedValue
            'If Q13.SelectedIndex = -1 Then dr_row("Q13") = Convert.DBNull Else dr_row("Q13") = Q13.SelectedValue
            'If Q14.SelectedIndex = -1 Then dr_row("Q14") = Convert.DBNull Else dr_row("Q14") = Q14.SelectedValue
            'If Q15.SelectedIndex = -1 Then dr_row("Q15") = Convert.DBNull Else dr_row("Q15") = Q15.SelectedValue
            'If Q16.SelectedIndex = -1 Then dr_row("Q16") = Convert.DBNull Else dr_row("Q16") = Q16.SelectedValue
            dr_row("ModifyAcct") = sm.UserInfo.UserID
            dr_row("ModifyDate") = Now()
            If flagI = 1 Then dtTable.Rows.Add(dr_row)
            DbAccess.UpdateDataTable(dtTable, sqlAdapter, objTrans)
            DbAccess.CommitTrans(objTrans)
            TIMS.CloseDbConn(tConn)
            'objconn.Close()
            Common.AddClientScript(Page, "insert_next();")
        Catch ex As Exception
            DbAccess.RollbackTrans(objTrans)
            Common.MessageBox(Me, ex.ToString)
            'objconn.Close()
            Throw ex
        End Try
    End Sub
End Class