Partial Class SD_11_004_add08
    Inherits AuthBasePage

    '本程式異動Table: Stud_QuestionFac
    Dim sqlAdapter As SqlDataAdapter
    Dim objconn As SqlConnection
    Dim stud_table As DataTable
   'Dim FunDr As DataRow

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

        SOCID.Attributes("onchange") = "if(this.selectedIndex==0) return false;"
        'RadioButtonList3_4.Attributes("onclick") = "disable_radio3();"
        'RadioButtonList5_3.Attributes("onclick") = "disable_radio5();"
        ProcessType.Value = Request("ProcessType")

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
        '        FunDr = FunDrArray(0)
        '        If ProcessType.Value = "Update" Then
        '            If FunDr("Mod") = "0" And FunDr("Del") = "0" Then
        '                Button1.Enabled = False
        '            Else
        '                Button1.Enabled = True
        '            End If
        '        ElseIf ProcessType.Value = "Insert" Or ProcessType.Value = "Next" Then
        '            If FunDr("Adds") = "1" Then
        '                Button1.Enabled = True
        '                next_but.Enabled = True
        '            Else
        '                Button1.Enabled = False
        '            End If
        '            If ProcessType.Value = "next" Then MoveNext()
        '        End If
        '    End If
        'End If

        If Not IsPostBack Then
            Dim sql As String
            'Dim dr As DataRow
            Dim dt As DataTable
            sql = "" & vbCrLf
            sql += " SELECT cs.StudentID ,CASE WHEN LEN(cs.StudentID) = 12 THEN ss.Name COLLATE CHINESE_TAIWAN_STROKE_CI_AS + '(' + SUBSTRING(cs.StudentID, LEN(cs.StudentID)-3, 3) + ')' ELSE ss.Name COLLATE CHINESE_TAIWAN_STROKE_CI_AS + '(' + SUBSTRING(cs.StudentID, LEN(cs.StudentID)-2, 2) + ')' END AS Name , cs.SOCID " & vbCrLf
            sql += " FROM Class_StudentsOfClass cs " & vbCrLf
            sql += " JOIN Stud_StudentInfo ss ON ss.SID = cs.SID AND cs.OCID = '" & Request("OCID") & "' " & vbCrLf
            dt = DbAccess.GetDataTable(sql, objconn)
            dt.DefaultView.Sort = "StudentID"

            With SOCID
                .DataSource = dt
                .DataTextField = "Name"
                .DataValueField = "SOCID"
                .DataBind()
                .Items.Insert(0, New ListItem("===請選擇===", ""))
            End With
            Common.SetListItem(SOCID, Request("SOCID"))

            Me.ViewState("QuestionFacSearchStr") = Session("QuestionFacSearchStr")
            Session("QuestionFacSearchStr") = Nothing

            Re_OCID.Value = Request("ocid")
            Re_SOCID.Value = Request("SOCID")

            create2(Re_OCID.Value, Re_SOCID.Value)

            '判斷課程所歸屬的計畫**by Milor 20080414----start
            If (Request("orgkind") = "G") Then
                Label4.Text = "產業人才投資計畫"
                Label5.Text = "產業人才投資計畫"
            ElseIf (Request("orgkind") = "W") Then
                Label4.Text = "提升勞工自主學習計畫"
                Label5.Text = "提升勞工自主學習計畫"
            End If
            '判斷課程所歸屬的計畫**by Milor 20080414----end
            If ProcessType.Value = "del" Then
                Dim sqlstrdel As String = "delete Stud_QuestionFac where SOCID= '" & Re_SOCID.Value & "'"
                DbAccess.ExecuteNonQuery(sqlstrdel, objconn)
            End If
            If ProcessType.Value = "check" Then
                create(Re_SOCID.Value)
                Button1.Enabled = False
                next_but.Enabled = False
            ElseIf ProcessType.Value = "Edit" Then '修改
                create(Re_SOCID.Value)
                Button1.Enabled = True
                next_but.Enabled = False
            End If
            If ProcessType.Value = "next" Then MoveNext()
        End If
        Button1.Attributes.Add("OnClick", "return ChkData();")
    End Sub

    Private Sub create(ByVal StrSOCID As String)
        Dim sqlstr As String
        Dim I As Integer
        sqlstr = "select * from Stud_QuestionFac where SOCID = '" & StrSOCID & "'"
        Dim row_list As DataRow = DbAccess.GetOneRow(sqlstr, objconn)
        If Not row_list Is Nothing Then
            If Not IsDBNull(row_list("Q1_1")) Then Q1_1.SelectedValue = row_list("Q1_1") Else Q1_1.SelectedIndex = -1
            If Not IsDBNull(row_list("Q1_2")) Then Q1_2.SelectedValue = row_list("Q1_2") Else Q1_2.SelectedIndex = -1
            If Not IsDBNull(row_list("Q1_3")) Then Q1_3.SelectedValue = row_list("Q1_3") Else Q1_3.SelectedIndex = -1
            If Not IsDBNull(row_list("Q1_4")) Then Q1_4.SelectedValue = row_list("Q1_4") Else Q1_4.SelectedIndex = -1
            If Not IsDBNull(row_list("Q1_5")) Then Q1_5.SelectedValue = row_list("Q1_5") Else Q1_5.SelectedIndex = -1
            If Not IsDBNull(row_list("Q1_6")) Then
                If row_list("Q1_6") = 1 Then Q1_61.Checked = True
                If row_list("Q1_6") = 2 Then Q1_62.Checked = True
                If row_list("Q1_6") = 3 Then Q1_63.Checked = True
            End If
            If Not IsDBNull(row_list("Q1_6_CCourName")) Then Q1_6_CCourName.Text = row_list("Q1_6_CCourName") Else Q1_6_CCourName.Text = ""
            If Not IsDBNull(row_list("Q1_6_CHour")) Then Q1_6_CHour.Text = row_list("Q1_6_CHour") Else Q1_6_CHour.Text = ""
            If Not IsDBNull(row_list("Q1_6_MCourName")) Then Q1_6_MCourName.Text = row_list("Q1_6_MCourName") Else Q1_6_MCourName.Text = ""
            If Not IsDBNull(row_list("Q1_6_MHour")) Then Q1_6_MHour.Text = row_list("Q1_6_MHour") Else Q1_6_MHour.Text = ""
            If Not IsDBNull(row_list("Q2_1")) Then
                If row_list("Q2_1") = 1 Then Q2_11.Checked = True
                If row_list("Q2_1") = 2 Then Q2_12.Checked = True
                If row_list("Q2_1") = 3 Then Q2_13.Checked = True
            End If
            If Not IsDBNull(row_list("Q2_1_CCourName")) Then Q2_1_CCourName.Text = row_list("Q2_1_CCourName") Else Q2_1_CCourName.Text = ""
            If Not IsDBNull(row_list("Q2_1_CHour")) Then Q2_1_CHour.Text = row_list("Q2_1_CHour") Else Q2_1_CHour.Text = ""
            If Not IsDBNull(row_list("Q2_1_MCourName")) Then Q2_1_MCourName.Text = row_list("Q2_1_MCourName") Else Q2_1_MCourName.Text = ""
            If Not IsDBNull(row_list("Q2_1_MHour")) Then Q2_1_MHour.Text = row_list("Q2_1_MHour") Else Q2_1_MHour.Text = ""
            If Not IsDBNull(row_list("Q2_2")) Then Q2_2.SelectedValue = row_list("Q2_2") Else Q2_2.SelectedIndex = -1
            If Not IsDBNull(row_list("Q2_3")) Then Q2_3.SelectedValue = row_list("Q2_3") Else Q2_3.SelectedIndex = -1
            If Not IsDBNull(row_list("Q2_4")) Then Q2_4.SelectedValue = row_list("Q2_4") Else Q2_4.SelectedIndex = -1
            If Not IsDBNull(row_list("Q2_5")) Then Q2_5.SelectedValue = row_list("Q2_5") Else Q2_5.SelectedIndex = -1
            If Not IsDBNull(row_list("Q3_1")) Then Q3_1.SelectedValue = row_list("Q3_1") Else Q3_1.SelectedIndex = -1
            If Not IsDBNull(row_list("Q3_2")) Then Q3_2.SelectedValue = row_list("Q3_2") Else Q3_2.SelectedIndex = -1
            If Not IsDBNull(row_list("Q3_3")) Then Q3_3.SelectedValue = row_list("Q3_3") Else Q3_3.SelectedIndex = -1
            If Not IsDBNull(row_list("Q4")) Then Q4.SelectedValue = row_list("Q4") Else Q4.SelectedIndex = -1
            If Not IsDBNull(row_list("Q5")) Then
                If row_list("Q5") = 1 Then Q5_1.Checked = True
                If row_list("Q5") = 2 Then Q5_2.Checked = True
                If row_list("Q5") = 3 Then Q5_3.Checked = True
                If row_list("Q5") = 4 Then Q5_4.Checked = True
                If row_list("Q5") = 5 Then Q5_5.Checked = True
                If row_list("Q5") = 6 Then Q5_6.Checked = True
            End If
            If Not IsDBNull(row_list("Q5_Note_News")) Then Q5_Note_News.Text = row_list("Q5_Note_News") Else Q5_Note_News.Text = ""
            If Not IsDBNull(row_list("Q5_Note_Other")) Then Q5_Note_Other.Text = row_list("Q5_Note_Other") Else Q5_Note_Other.Text = ""
            If Not IsDBNull(row_list("Q6")) Then
                For I = 1 To Len(row_list("Q6").ToString)
                    Select Case Mid(row_list("Q6").ToString, I, 1)
                        Case "1"
                            Q6_1.Checked = True
                        Case "2"
                            Q6_2.Checked = True
                        Case "3"
                            Q6_3.Checked = True
                    End Select
                Next
            Else
                Q6_1.Checked = False
                Q6_2.Checked = False
                Q6_3.Checked = False
            End If
            If Not IsDBNull(row_list("Q6_Note1")) Then Q6_Note1.Text = row_list("Q6_Note1") Else Q6_Note1.Text = ""
            If Not IsDBNull(row_list("Q6_Note2")) Then Q6_Note2.Text = row_list("Q6_Note2") Else Q6_Note2.Text = ""
            If Not IsDBNull(row_list("Q7")) Then Q7.SelectedValue = row_list("Q7") Else Q7.SelectedIndex = -1
            If Not IsDBNull(row_list("Q8")) Then Q8.SelectedValue = row_list("Q8") Else Q8.SelectedIndex = -1
            '加入第9項到第12項**by Milor 20080414----start
            If Not IsDBNull(row_list("Q9_1")) Then Q9_1.SelectedValue = row_list("Q9_1") Else Q9_1.SelectedIndex = -1
            If Not IsDBNull(row_list("Q9_2")) Then Q9_2.SelectedValue = row_list("Q9_2") Else Q9_2.SelectedIndex = -1
            If Not IsDBNull(row_list("Q9_3")) Then Q9_3.SelectedValue = row_list("Q9_3") Else Q9_3.SelectedIndex = -1
            If Not IsDBNull(row_list("Q10")) Then Q10.SelectedValue = row_list("Q10") Else Q10.SelectedIndex = -1
            If Not IsDBNull(row_list("Q11")) Then
                Label3.Text = " " & row_list("Q11") & " "
                Q11.Visible = False
            Else
                Label3.Text = "     "
                Q11.Visible = False
            End If
            If Not IsDBNull(row_list("Q12")) Then Q12.Text = row_list("Q12") Else Q12.Text = ""
            '加入第9項到第12項**by Milor 20080414----end
        Else
            Q1_1.SelectedIndex = -1
            Q1_2.SelectedIndex = -1
            Q1_3.SelectedIndex = -1
            Q1_4.SelectedIndex = -1
            Q1_5.SelectedIndex = -1
            Q1_61.Checked = False
            Q1_62.Checked = False
            Q1_63.Checked = False
            Q1_6_CCourName.Text = ""
            Q1_6_CHour.Text = ""
            Q1_6_MCourName.Text = ""
            Q1_6_MHour.Text = ""
            Q2_11.Checked = False
            Q2_12.Checked = False
            Q2_13.Checked = False
            Q2_1_CCourName.Text = ""
            Q2_1_CHour.Text = ""
            Q2_1_MCourName.Text = ""
            Q2_1_MHour.Text = ""
            Q2_2.SelectedIndex = -1
            Q2_3.SelectedIndex = -1
            Q2_4.SelectedIndex = -1
            Q2_5.SelectedIndex = -1
            Q3_1.SelectedIndex = -1
            Q3_2.SelectedIndex = -1
            Q3_3.SelectedIndex = -1
            Q4.SelectedIndex = -1
            Q5_1.Checked = False
            Q5_2.Checked = False
            Q5_3.Checked = False
            Q5_4.Checked = False
            Q5_5.Checked = False
            Q5_6.Checked = False
            Q5_Note_News.Text = ""
            Q5_Note_Other.Text = ""
            Q6_1.Checked = False
            Q6_2.Checked = False
            Q6_3.Checked = False
            Q6_Note1.Text = ""
            Q6_Note2.Text = ""
            Q7.SelectedIndex = -1
            Q8.SelectedIndex = -1
            '加入第9項到第12項**by Milor 20080414----start
            Q9_1.SelectedIndex = -1
            Q9_2.SelectedIndex = -1
            Q9_3.SelectedIndex = -1
            Q10.SelectedIndex = -1
            Label3.Text = ""
            Q11.Text = ""
            Q11.Visible = True
            Q12.Text = ""
            '加入第9項到第12項**by Milor 20080414----end
        End If
    End Sub

    Sub create2(ByVal StrOCID As String, ByVal StrSOCID As String)
        Dim row As DataRow
        Dim sqlstr As String = ""

        sqlstr = "" & vbCrLf
        sqlstr += " SELECT b.SOCID ,b.studentid ,c.name ,b.StudStatus ,b.RejectTDate1 ,b.RejectTDate2 " & vbCrLf
        sqlstr += " 	   ,FORMAT(ISNULL(p.totalcost,0)/(CASE WHEN ISNULL(p.TNum, 1) <> 0 THEN ISNULL(p.TNum, 1) ELSE 1 END), '#######0.##') AS totalcost " & vbCrLf
        sqlstr += " 	   ,FORMAT(ISNULL(p.defstdcost,0)/(CASE WHEN ISNULL(p.TNum, 1) <> 0 THEN ISNULL(p.TNum, 1) ELSE 1 END), '#######0.##') AS defstdcost " & vbCrLf
        sqlstr += " FROM class_classinfo a " & vbCrLf
        sqlstr += " JOIN class_studentsofclass b ON a.ocid = b.ocid AND b.OCID = '" & StrOCID & "' AND b.SOCID = '" & StrSOCID & "' " & vbCrLf
        sqlstr += " JOIN plan_planinfo p ON a.PlanID = p.PlanID AND a.comIDNO = p.comIDNO AND a.SeqNO = p.SeqNO " & vbCrLf
        sqlstr += " JOIN stud_studentinfo c ON b.sid = c.sid " & vbCrLf
        row = DbAccess.GetOneRow(sqlstr, objconn)
        Me.Label_Name.Text = row("name")
        Me.Label_Stud.Text = row("studentid")
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
        Me.Label1.Text = row("totalcost").ToString
        Me.Label2.Text = row("defstdcost").ToString
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Session("QuestionFacSearchStr") = Me.ViewState("QuestionFacSearchStr")
        'Response.Redirect("SD_11_004.aspx?ProcessType=Back&ID=" & Re_ID.Value & "")
    End Sub

    Private Sub check_last()
        Session("QuestionFacSearchStr") = Me.ViewState("QuestionFacSearchStr")
        Dim strScript As String
        strScript = "<script language=""javascript"">" + vbCrLf
        strScript += "alert('已為此班級中最後一筆學員!!');" + vbCrLf
        strScript += "location.href ='SD_11_004.aspx?ProcessType=Back&ID='+document.getElementById('Re_ID').value;" + vbCrLf
        strScript += "</script>"
        Page.RegisterStartupScript("", strScript)
    End Sub

    Private Sub next_but_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles next_but.Click
        MoveNext()
    End Sub

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
                create2(Re_OCID.Value, Re_SOCID.Value)
            End If
        End If
    End Sub

    Private Sub SOCID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SOCID.SelectedIndexChanged
        Re_SOCID.Value = SOCID.SelectedValue
        create(SOCID.SelectedValue)
        create2(Re_OCID.Value, Re_SOCID.Value)
    End Sub

    Function Check_Stud_QuestionFac() As Boolean
        Dim msg As String = ""
        If Q1_6_CHour.Text <> "" Then
            Q1_6_CHour.Text = Trim(Q1_6_CHour.Text)
            If Not IsNumeric(Q1_6_CHour.Text) Then
                msg += "訓練時數應增加必須為數字" & vbCrLf
            Else
                If CInt(Q1_6_CHour.Text) <= 0 Then msg += "訓練時數應增加 必須為大於0的數字，沒有(-)號" & vbCrLf
                If CDbl(CInt(Q1_6_CHour.Text)) <> CDbl(Q1_6_CHour.Text) Then msg += "訓練時數應增加 必須為每小時，不可輸入浮點數" & vbCrLf
                If msg = "" Then Q1_6_CHour.Text = CInt(Q1_6_CHour.Text)
            End If
        End If

        If Q1_6_MHour.Text <> "" Then
            Q1_6_MHour.Text = Trim(Q1_6_MHour.Text)
            If Not IsNumeric(Q1_6_MHour.Text) Then
                msg += "訓練時數應減少 必須為數字" & vbCrLf
            Else
                If CInt(Q1_6_MHour.Text) <= 0 Then msg += "訓練時數應減少 必須為大於0的數字，沒有(-)號" & vbCrLf
                If CDbl(CInt(Q1_6_MHour.Text)) <> CDbl(Q1_6_MHour.Text) Then msg += "訓練時數應減少 必須為每小時，不可輸入浮點數" & vbCrLf
                If msg = "" Then Q1_6_MHour.Text = CInt(Q1_6_MHour.Text)
            End If
        End If

        If Q2_1_CHour.Text <> "" Then
            Q2_1_CHour.Text = Trim(Q2_1_CHour.Text)
            If Not IsNumeric(Q2_1_CHour.Text) Then
                msg += "術科時數應增加 必須為數字" & vbCrLf
            Else
                If CInt(Q2_1_CHour.Text) <= 0 Then msg += "術科時數應增加 必須為大於0的數字，沒有(-)號" & vbCrLf
                If CDbl(CInt(Q2_1_CHour.Text)) <> CDbl(Q2_1_CHour.Text) Then msg += "術科時數應增加 必須為每小時，不可輸入浮點數" & vbCrLf
                If msg = "" Then Q2_1_CHour.Text = CInt(Q2_1_CHour.Text)
            End If
        End If

        If Q2_1_MHour.Text <> "" Then
            Q2_1_MHour.Text = Trim(Q2_1_MHour.Text)
            If Not IsNumeric(Q2_1_MHour.Text) Then
                msg += "術科時數應減少 必須為數字" & vbCrLf
            Else
                If CInt(Q2_1_MHour.Text) <= 0 Then msg += "術科時數應減少 必須為大於0的數字，沒有(-)號" & vbCrLf
                If CDbl(CInt(Q2_1_MHour.Text)) <> CDbl(Q2_1_MHour.Text) Then msg += "術科時數應減少 必須為每小時，不可輸入浮點數" & vbCrLf
                If msg = "" Then Q2_1_MHour.Text = CInt(Q2_1_MHour.Text)
            End If
        End If

        '判斷第11項是否為數字**by Milor 20080414----start
        If Q11.Text <> "" Then
            Q11.Text = Trim(Q11.Text)
            If Not IsNumeric(Q11.Text) Then
                msg += "每年願意自費之金額 必須為數字" & vbCrLf
            Else
                If CInt(Q11.Text) <= 0 Then msg += "每年願意自費之金額 必須為大於0的數字，沒有(-)號" & vbCrLf
                If CDbl(CInt(Q11.Text)) <> CDbl(Q11.Text) Then msg += "每年願意自費之金額 不可輸入浮點數" & vbCrLf
                If msg = "" Then Q11.Text = CInt(Q11.Text)
            End If
        End If
        '判斷第11項是否為數字**by Milor 20080414----end

        If msg <> "" Then
            Common.MessageBox(Me, msg)
            Return False
        Else
            Return True
        End If
    End Function

    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        If Trim(Q12.Text) <> "" Then
            Q12.Text = Trim(Q12.Text)
            If Len(Q12.Text) > 200 Then Errmsg += "其他建議 長度超過系統範圍(200)" & vbCrLf
        Else
            'Errmsg += "請輸入 其他建議" & vbCrLf
        End If

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Dim sqlAdapter As SqlDataAdapter = Nothing
        Dim dr_row As DataRow = Nothing
        Dim dtTable As DataTable = Nothing
        Dim objTrans As SqlTransaction = Nothing

        'objconn.Open()
        Call TIMS.OpenDbConn(objconn)
        Dim sql As String = ""
        Dim I As Integer = 0
        'If Not Page.IsValid Then Exit Sub
        If Not Check_Stud_QuestionFac() Then Exit Sub
        Session("QuestionFacSearchStr") = Me.ViewState("QuestionFacSearchStr")

        objTrans = DbAccess.BeginTrans(objconn)
        sql = "SELECT * FROM Stud_QuestionFac WHERE SOCID='" & SOCID.SelectedValue & "'"
        dtTable = DbAccess.GetDataTable(sql, sqlAdapter, objTrans)
        If dtTable.Rows.Count = 0 Then
            I = 1
            dr_row = dtTable.NewRow
            dr_row("SOCID") = SOCID.SelectedValue
        Else
            I = 0
            dr_row = dtTable.Rows(0)
        End If

        If Q1_1.SelectedIndex = -1 Then dr_row("Q1_1") = Convert.DBNull Else dr_row("Q1_1") = Q1_1.SelectedValue
        If Q1_2.SelectedIndex = -1 Then dr_row("Q1_2") = Convert.DBNull Else dr_row("Q1_2") = Q1_2.SelectedValue
        If Q1_3.SelectedIndex = -1 Then dr_row("Q1_3") = Convert.DBNull Else dr_row("Q1_3") = Q1_3.SelectedValue
        If Q1_4.SelectedIndex = -1 Then dr_row("Q1_4") = Convert.DBNull Else dr_row("Q1_4") = Q1_4.SelectedValue
        If Q1_5.SelectedIndex = -1 Then dr_row("Q1_5") = Convert.DBNull Else dr_row("Q1_5") = Q1_5.SelectedValue
        If Q1_61.Checked = True Then
            dr_row("Q1_6") = 1
        ElseIf Q1_62.Checked = True Then
            dr_row("Q1_6") = 2
        ElseIf Q1_63.Checked = True Then
            dr_row("Q1_6") = 3
        Else
            dr_row("Q1_6") = Convert.DBNull
        End If
        If Q1_6_CCourName.Text = "" Then dr_row("Q1_6_CCourName") = Convert.DBNull Else dr_row("Q1_6_CCourName") = Q1_6_CCourName.Text
        If Q1_6_CHour.Text = "" Then dr_row("Q1_6_CHour") = Convert.DBNull Else dr_row("Q1_6_CHour") = Q1_6_CHour.Text
        If Q1_6_MCourName.Text = "" Then dr_row("Q1_6_MCourName") = Convert.DBNull Else dr_row("Q1_6_MCourName") = Q1_6_MCourName.Text
        If Q1_6_MHour.Text = "" Then dr_row("Q1_6_MHour") = Convert.DBNull Else dr_row("Q1_6_MHour") = Q1_6_MHour.Text
        If Q2_11.Checked = True Then
            dr_row("Q2_1") = 1
        ElseIf Q2_12.Checked = True Then
            dr_row("Q2_1") = 2
        ElseIf Q2_13.Checked = True Then
            dr_row("Q2_1") = 3
        Else
            dr_row("Q2_1") = Convert.DBNull
        End If
        If Q2_1_CCourName.Text = "" Then dr_row("Q2_1_CCourName") = Convert.DBNull Else dr_row("Q2_1_CCourName") = Q2_1_CCourName.Text
        If Q2_1_CHour.Text = "" Then dr_row("Q2_1_CHour") = Convert.DBNull Else dr_row("Q2_1_CHour") = Q2_1_CHour.Text
        If Q2_1_MCourName.Text = "" Then dr_row("Q2_1_MCourName") = Convert.DBNull Else dr_row("Q2_1_MCourName") = Q2_1_MCourName.Text
        If Q2_1_MHour.Text = "" Then dr_row("Q2_1_MHour") = Convert.DBNull Else dr_row("Q2_1_MHour") = Q2_1_MHour.Text
        If Q2_2.SelectedIndex = -1 Then dr_row("Q2_2") = Convert.DBNull Else dr_row("Q2_2") = Q2_2.SelectedValue
        If Q2_3.SelectedIndex = -1 Then dr_row("Q2_3") = Convert.DBNull Else dr_row("Q2_3") = Q2_3.SelectedValue
        If Q2_4.SelectedIndex = -1 Then dr_row("Q2_4") = Convert.DBNull Else dr_row("Q2_4") = Q2_4.SelectedValue
        If Q2_5.SelectedIndex = -1 Then dr_row("Q2_5") = Convert.DBNull Else dr_row("Q2_5") = Q2_5.SelectedValue
        If Q3_1.SelectedIndex = -1 Then dr_row("Q3_1") = Convert.DBNull Else dr_row("Q3_1") = Q3_1.SelectedValue
        If Q3_2.SelectedIndex = -1 Then dr_row("Q3_2") = Convert.DBNull Else dr_row("Q3_2") = Q3_2.SelectedValue
        If Q3_3.SelectedIndex = -1 Then dr_row("Q3_3") = Convert.DBNull Else dr_row("Q3_3") = Q3_3.SelectedValue
        If Q4.SelectedIndex = -1 Then dr_row("Q4") = Convert.DBNull Else dr_row("Q4") = Q4.SelectedValue
        If Q5_1.Checked = True Then
            dr_row("Q5") = 1
        ElseIf Q5_2.Checked = True Then
            dr_row("Q5") = 2
        ElseIf Q5_3.Checked = True Then
            dr_row("Q5") = 3
        ElseIf Q5_4.Checked = True Then
            dr_row("Q5") = 4
        ElseIf Q5_5.Checked = True Then
            dr_row("Q5") = 5
        ElseIf Q5_6.Checked = True Then
            dr_row("Q5") = 6
        Else
            dr_row("Q5") = Convert.DBNull
        End If
        If Q5_Note_News.Text = "" Then dr_row("Q5_Note_News") = Convert.DBNull Else dr_row("Q5_Note_News") = Q5_Note_News.Text
        If Q5_Note_Other.Text = "" Then dr_row("Q5_Note_Other") = Convert.DBNull Else dr_row("Q5_Note_Other") = Q5_Note_Other.Text
        If Q6_3.Checked = True Then dr_row("Q6") = 3
        If Q6_1.Checked = True And Q6_2.Checked = True Then dr_row("Q6") = 3
        If Q6_1.Checked = True And Q6_2.Checked = False Then dr_row("Q6") = 1
        If Q6_1.Checked = False And Q6_2.Checked = True Then dr_row("Q6") = 2
        If Q6_Note1.Text = "" Then dr_row("Q6_Note1") = Convert.DBNull Else dr_row("Q6_Note1") = Q6_Note1.Text
        If Q6_Note2.Text = "" Then dr_row("Q6_Note2") = Convert.DBNull Else dr_row("Q6_Note2") = Q6_Note2.Text
        If Q7.SelectedIndex = -1 Then dr_row("Q7") = Convert.DBNull Else dr_row("Q7") = Q7.SelectedValue
        If Q8.SelectedIndex = -1 Then dr_row("Q8") = Convert.DBNull Else dr_row("Q8") = Q8.SelectedValue
        '加入第9項到第12項**by Milor 20080414----start
        If Q9_1.SelectedIndex = -1 Then dr_row("Q9_1") = Convert.DBNull Else dr_row("Q9_1") = Q9_1.SelectedValue
        If Q9_2.SelectedIndex = -1 Then dr_row("Q9_2") = Convert.DBNull Else dr_row("Q9_2") = Q9_2.SelectedValue
        If Q9_3.SelectedIndex = -1 Then dr_row("Q9_3") = Convert.DBNull Else dr_row("Q9_3") = Q9_3.SelectedValue
        If Q10.SelectedIndex = -1 Then dr_row("Q10") = Convert.DBNull Else dr_row("Q10") = Q10.SelectedValue
        If Q11.Text = "" Then dr_row("Q11") = Convert.DBNull Else dr_row("Q11") = Q11.Text
        If Q12.Text = "" Then dr_row("Q12") = Convert.DBNull Else dr_row("Q12") = Q12.Text
        '加入第9項到第12項**by Milor 20080414----start
        dr_row("ModifyAcct") = sm.UserInfo.UserID
        dr_row("ModifyDate") = Now()
        If I = 1 Then dtTable.Rows.Add(dr_row)
        DbAccess.UpdateDataTable(dtTable, sqlAdapter, objTrans)
        DbAccess.CommitTrans(objTrans)
        Common.AddClientScript(Page, "insert_next();")
    End Sub
End Class