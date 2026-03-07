Partial Class SD_11_004_add12
    Inherits AuthBasePage

    '本程式異動Table: Stud_QuestionFac '受訓學員意見調查表
    '受訓學員意見調查表
    Dim flagLock As Boolean = False '(解)進行鎖定。
    Dim sqlAdapter As SqlDataAdapter
    Dim stud_table As DataTable
    'Dim FunDr As DataRow

    Const cst_errmsg2 As String = "查詢無該學員資料，請重新確認搜尋條件!!"
    Const ss_QuestionFacSearchStr As String = "QuestionFacSearchStr"

    'ProcessType/CommandName/.aspx
    Const cst_ptInsert As String = "Insert" 'e.CommandName'.aspx
    Const cst_ptDelete As String = "Delete" '.aspx
    Const cst_ptCheck As String = "Check" 'e.CommandName'.aspx
    Const cst_ptEdit As String = "Edit" 'e.CommandName'.aspx
    'Const cst_ptClear As String = "Clear" 'e.CommandName
    'Const cst_ptView As String = "View"
    'Const cst_ptSaveNext As String = "SaveNext" '儲存後移動下一筆。
    Const cst_ptNext As String = "Next" '單純移動下一筆。
    Const cst_ptBack As String = "Back" '.aspx(ProcessType)

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
        '檢查Session是否存在 End

        SOCID.Attributes("onchange") = "if(this.selectedIndex==0) return false;"
        'RadioButtonList3_4.Attributes("onclick") = "disable_radio3();"
        'RadioButtonList5_3.Attributes("onclick") = "disable_radio5();"
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
        '        FunDr = FunDrArray(0)
        '        Select Case ProcessType.Value
        '            Case cst_ptEdit
        '                Button1.Enabled = True
        '                If FunDr("Mod") = "0" And FunDr("Del") = "0" Then
        '                    Button1.Enabled = False '不可儲存
        '                    TIMS.Tooltip(Button1, "修改權限不足")
        '                End If
        '            Case cst_ptInsert
        '                Button1.Enabled = True '可儲存
        '                If FunDr("Adds") = "0" Then
        '                    Button1.Enabled = False '不可儲存
        '                    TIMS.Tooltip(Button1, "新增權限不足")
        '                End If
        '                'SAVE NEXT
        '                If ProcessType.Value = cst_ptNext Then MoveNext()
        '        End Select
        '    End If
        'End If

#End Region

        If Not IsPostBack Then
            Call Create1() 'LIST該班級所有學員。
            Common.SetListItem(SOCID, Request("SOCID")) '選定1學員。
            'Me.ViewState("QuestionFacSearchStr") = Session("QuestionFacSearchStr")
            'Session("QuestionFacSearchStr") = Nothing
            If Not Session(ss_QuestionFacSearchStr) Is Nothing Then
                Me.ViewState(ss_QuestionFacSearchStr) = Session(ss_QuestionFacSearchStr)
                Session(ss_QuestionFacSearchStr) = Session(ss_QuestionFacSearchStr)
            Else
                Session(ss_QuestionFacSearchStr) = Me.ViewState(ss_QuestionFacSearchStr)
            End If
            Re_OCID.Value = Request("ocid")
            Re_SOCID.Value = Request("SOCID")
            If Not create2(Re_OCID.Value, Re_SOCID.Value) Then
                'Common.MessageBox(Me, cst_errmsg2)
                '沒有資料，自動返回。
                Dim strScript As String
                strScript = "<script language=""javascript"">" + vbCrLf
                strScript += "alert('" & cst_errmsg2 & "');" + vbCrLf
                strScript += "location.href ='SD_11_004.aspx?ProcessType=" & cst_ptBack & "&ID='+document.getElementById('Re_ID').value;" + vbCrLf
                strScript += "</script>"
                Page.RegisterStartupScript("", strScript)
                Exit Sub
            End If

            '判斷課程所歸屬的計畫**by Milor 20080414----start
            Dim PNAME As String = ""
            Select Case Convert.ToString(Request("orgkind"))
                Case "G", "W"
                    PNAME = TIMS.Get_PName28(Me, Convert.ToString(Request("orgkind")), objconn)
            End Select
            Label4.Text = PNAME '"產業人才投資計畫" '"提升勞工自主學習計畫"
            Label5.Text = PNAME '"產業人才投資計畫" '"提升勞工自主學習計畫"
            '判斷課程所歸屬的計畫**by Milor 20080414----end

            Select Case ProcessType.Value
                Case cst_ptDelete '"del"
                    If Chk_DASOURCE(Re_SOCID.Value) = 1 Then
                        Common.MessageBox(Me, "無法逕行刪除，僅供單位查詢。")
                        Exit Sub
                    End If

                    '做刪除動作。
                    Dim sqldel As String = ""
                    sqldel = "DELETE STUD_QUESTIONFAC2 WHERE SOCID='" & Re_SOCID.Value & "'"
                    DbAccess.ExecuteNonQuery(sqldel, objconn)
                    sqldel = "DELETE STUD_QUESTIONFAC WHERE SOCID='" & Re_SOCID.Value & "'"
                    DbAccess.ExecuteNonQuery(sqldel, objconn)
                Case cst_ptCheck '"check"
                    create(Re_SOCID.Value) 'modify flagLock
                    Button1.Enabled = False '儲存。
                    next_but.Enabled = False  '移動下一筆。
                    TIMS.Tooltip(Button1, "目前動作不提供儲存。")
                    TIMS.Tooltip(next_but, "目前動作不提供移動下一筆。")
                Case cst_ptEdit '"Edit" '修改
                    create(Re_SOCID.Value) 'modify flagLock
                    'Button1.Enabled = False '儲存。
                    'next_but.Enabled = False  '移動下一筆。
                    'If Not flagLock Then Button1.Enabled = True '儲存。
                Case cst_ptNext  '"next"
                    '移動下一筆。
                    Call MoveNext()
            End Select
        End If
        Button1.Attributes.Add("OnClick", "return ChkData();")
    End Sub

    'LIST該班級所有學員。
    Sub Create1()
        '單位填寫'資料來源	Null:未知 1:報名網()(學員外網填寫。)2@TIMS系統()
        '新增「填寫來源」檢核機制，由系統判斷填寫來源為產投報名網或TIMS系統：
        '1.若為報名網，訓練單位端之各項功能(新增、修改、查詢、列印、清除重填等)將反灰，無法逕行修改，僅保留填寫狀態供訓練單位查詢。
        '2.若為TIMS系統，保留訓練單位各項功能(新增、修改、查詢、列印、清除重填)。
        '3.若有訓練單位先於TIMS系統協助填寫，學員後於報名網修正之情形者，訓練單位端之各項功能(新增、修改、查詢、列印、清除重填等)將反灰。

        Dim sql As String = ""
        'Dim dr As DataRow
        Dim dt As DataTable
        sql = "" & vbCrLf
        sql &= " SELECT cs.StudentID ,CASE WHEN LEN(cs.StudentID) = 12 THEN ss.Name + '(' + SUBSTRING(cs.StudentID, LEN(cs.StudentID)-3, 3) + ')' ELSE ss.Name + '(' + SUBSTRING(cs.StudentID, LEN(cs.StudentID)-2, 2) + ')' END Name ,cs.SOCID " & vbCrLf
        sql &= " FROM Class_StudentsOfClass cs " & vbCrLf
        sql &= " JOIN Stud_StudentInfo ss ON ss.SID = cs.SID " & vbCrLf
        sql &= " WHERE 1=1 AND cs.OCID = '" & Request("OCID") & "' " & vbCrLf
        '排除：1:報名網()(學員外網填寫。)
        sql &= " AND NOT EXISTS (SELECT 'x' FROM Stud_QuestionFac x WHERE x.DaSource = '1' AND x.SOCID = cs.SOCID) " & vbCrLf
        dt = DbAccess.GetDataTable(sql, objconn)
        dt.DefaultView.Sort = "StudentID"
        With SOCID
            .DataSource = dt
            .DataTextField = "Name"
            .DataValueField = "SOCID"
            .DataBind()
            .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        End With
    End Sub

#Region "(No Use)"

    'Sub obj_SelectedValue(ByRef RBLobj As RadioButtonList, ByVal rowVal As Object)
    '    Common.SetListItem(RBLobj, rowVal)
    'End Sub

#End Region

    '清除資料。
    Sub Clear_datalist()
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
        Q7_8.SelectedIndex = -1
        Q7_9.SelectedIndex = -1
        Q8.SelectedIndex = -1
        '加入第9項到第12項**by Milor 20080414----start
        Q9_1.SelectedIndex = -1
        Q9_2.SelectedIndex = -1
        Q9_3.SelectedIndex = -1
        Q10.SelectedIndex = -1
        Label3.Text = ""
        Q11.Text = ""
        Q11.Visible = True
        Q14.SelectedIndex = -1
        Q12.Text = ""
        '加入第9項到第12項**by Milor 20080414----end
    End Sub

    '查詢問卷資料。
    Private Sub create(ByVal StrSOCID As String)
        Dim sqlstr As String = ""
        sqlstr = "SELECT * FROM STUD_QUESTIONFAC WHERE SOCID = '" & StrSOCID & "'"
        Dim row_list As DataRow = DbAccess.GetOneRow(sqlstr, objconn)
        flagLock = False '判定是否要鎖定。
        HidDASOURCE.Value = ""
        If Not row_list Is Nothing Then
            'flagLock:有資料且為學員填寫進行問卷鎖定。
            Select Case Convert.ToString(row_list("DASOURCE"))
                Case "1"
                    flagLock = True '進行鎖定。
                    'Common.MessageBox(Me, "無法逕行修改，僅保留填寫狀態供訓練單位查詢。")
                    'Exit Sub
            End Select
            If Convert.ToString(row_list("DASOURCE")) <> "" Then HidDASOURCE.Value = Convert.ToString(row_list("DASOURCE"))
        End If

        'If TIMS.sUtl_ChkTest() Then flagLock = True 'flagLock:有資料且為學員填寫進行問卷鎖定。 '測試用

        '清除資料。
        Call Clear_datalist()
        If Not row_list Is Nothing Then
            '顯示資料。
            Try
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
                    For i As Integer = 1 To Len(row_list("Q6").ToString)
                        Select Case Mid(row_list("Q6").ToString, i, 1)
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
                If Not IsDBNull(row_list("Q7_8")) Then Q7_8.SelectedValue = row_list("Q7_8") Else Q7_8.SelectedIndex = -1
                If Not IsDBNull(row_list("Q7_9")) Then Q7_9.SelectedValue = row_list("Q7_9") Else Q7_9.SelectedIndex = -1
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
                If Not IsDBNull(row_list("Q14")) Then Q14.SelectedValue = row_list("Q14") Else Q14.SelectedIndex = -1
                If Not IsDBNull(row_list("Q12")) Then Q12.Text = row_list("Q12") Else Q12.Text = ""
                '加入第9項到第12項**by Milor 20080414----end
            Catch ex As Exception
                Call Clear_datalist()
            End Try
        End If

        '依 flagLock 判斷下列動作
        Me.Table3_Datalist.Disabled = False '(解)進行鎖定。
        If flagLock = True Then Me.Table3_Datalist.Disabled = True '進行鎖定。

        'Button1.Enabled = False '儲存。
        ''next_but.Enabled = False  '移動下一筆。
        'Select Case ProcessType.Value
        '    Case cst_ptEdit, cst_ptNext   '修改' 儲存後移動下一筆。
        '        If Not flagLock Then Button1.Enabled = True '儲存。(編輯模式下進行解鎖)
        'End Select
    End Sub

    '查詢基本資料。某班，某學員基本資料。沒有為false
    Function create2(ByVal StrOCID As String, ByVal StrSOCID As String) As Boolean
        Dim rst As Boolean = False
        Dim row As DataRow
        Dim sqlstr As String = ""
        sqlstr = "" & vbCrLf
        sqlstr += " SELECT b.SOCID ,b.studentid ,c.name ,b.StudStatus " & vbCrLf
        sqlstr += "       ,CONVERT(varchar, b.RejectTDate1, 111) RejectTDate1 " & vbCrLf
        sqlstr += "       ,CONVERT(varchar, b.RejectTDate2, 111) RejectTDate2 " & vbCrLf
        sqlstr += "       ,FORMAT(ISNULL(p.totalcost,0)/(CASE WHEN ISNULL(p.TNum, 1) <> 0 THEN ISNULL(p.TNum, 1) ELSE 1 END), '#######0.##') totalcost " & vbCrLf
        sqlstr += "       ,FORMAT(ISNULL(p.defstdcost,0)/(CASE WHEN ISNULL(p.TNum, 1) <> 0 THEN ISNULL(p.TNum, 1) ELSE 1 END), '#######0.##') defstdcost " & vbCrLf
        sqlstr += " FROM class_classinfo a " & vbCrLf
        sqlstr += " JOIN class_studentsofclass b ON a.ocid = b.ocid " & vbCrLf
        sqlstr += " JOIN plan_planinfo p ON a.PlanID = p.PlanID AND a.comIDNO = p.comIDNO AND a.SeqNO = p.SeqNO " & vbCrLf
        sqlstr += " JOIN stud_studentinfo c ON b.sid = c.sid " & vbCrLf
        sqlstr += " WHERE 1=1 " & vbCrLf
        sqlstr += "    AND b.OCID = '" & StrOCID & "' "
        sqlstr += "    AND b.SOCID = '" & StrSOCID & "' " & vbCrLf
        row = DbAccess.GetOneRow(sqlstr, objconn)
        If Not row Is Nothing Then
            Me.Label_Name.Text = row("name")
            Me.Label_Stud.Text = row("studentid")
            Me.Label_Status.Text = TIMS.GET_STUDSTATUS_N23(row("StudStatus"), row("RejectTDate1"), row("RejectTDate2"))

            Me.Label1.Text = row("totalcost").ToString
            Me.Label2.Text = row("defstdcost").ToString
            rst = True
        End If
        Return rst
    End Function

    '回上一頁。
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'Session("QuestionFacSearchStr") = Me.ViewState("QuestionFacSearchStr")
        If Not Session(ss_QuestionFacSearchStr) Is Nothing Then
            Me.ViewState(ss_QuestionFacSearchStr) = Session(ss_QuestionFacSearchStr)
            Session(ss_QuestionFacSearchStr) = Session(ss_QuestionFacSearchStr)
        Else
            Session(ss_QuestionFacSearchStr) = Me.ViewState(ss_QuestionFacSearchStr)
        End If
        TIMS.Utl_Redirect1(Me, "SD_11_004.aspx?ProcessType=" & cst_ptBack & "&ID=" & Re_ID.Value & "")
    End Sub

    '移動到最後一筆返回首頁。
    Private Sub check_last()
        'Session("QuestionFacSearchStr") = Me.ViewState("QuestionFacSearchStr")
        If Not Session(ss_QuestionFacSearchStr) Is Nothing Then
            Me.ViewState(ss_QuestionFacSearchStr) = Session(ss_QuestionFacSearchStr)
            Session(ss_QuestionFacSearchStr) = Session(ss_QuestionFacSearchStr)
        Else
            Session(ss_QuestionFacSearchStr) = Me.ViewState(ss_QuestionFacSearchStr)
        End If
        Dim strScript As String
        strScript = "<script language=""javascript"">" + vbCrLf
        strScript += "alert('已為此班級中最後一筆學員!!');" + vbCrLf
        strScript += "location.href ='SD_11_004.aspx?ProcessType=" & cst_ptBack & "&ID='+document.getElementById('Re_ID').value;" + vbCrLf
        strScript += "</script>"
        Page.RegisterStartupScript("", strScript)
    End Sub

    'MOVE NEXT
    Private Sub next_but_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles next_but.Click
        Call MoveNext()
    End Sub

    'SAVE NEXT OR MOVE NEXT
    Private Sub MoveNext()
        If SOCID.Items.Count > 0 Then
            Dim NowIndex As Integer
            Dim MaxIndex As Integer
            MaxIndex = SOCID.Items.Count - 1
            NowIndex = SOCID.SelectedIndex
            If NowIndex = MaxIndex Then
                Call check_last() '移動到最後一筆返回首頁。
            Else
                SOCID.SelectedIndex = NowIndex + 1
                Re_SOCID.Value = SOCID.SelectedValue
                create(SOCID.SelectedValue) 'modify flagLock
                If Not create2(Re_OCID.Value, Re_SOCID.Value) Then
                    'Common.MessageBox(Me, cst_errmsg2)
                    Dim strScript As String
                    strScript = "<script language=""javascript"">" + vbCrLf
                    strScript += "alert('" & cst_errmsg2 & "');" + vbCrLf
                    strScript += "location.href ='SD_11_004.aspx?ProcessType=" & cst_ptBack & "&ID='+document.getElementById('Re_ID').value;" + vbCrLf
                    strScript += "</script>"
                    Page.RegisterStartupScript("", strScript)
                    Exit Sub
                End If

            End If
        End If
    End Sub

    '直接選學員。
    Private Sub SOCID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SOCID.SelectedIndexChanged
        Re_SOCID.Value = SOCID.SelectedValue
        Call create(SOCID.SelectedValue)
        If Not create2(Re_OCID.Value, Re_SOCID.Value) Then
            'Common.MessageBox(Me, cst_errmsg2)
            Dim strScript As String
            strScript = "<script language=""javascript"">" + vbCrLf
            strScript += "alert('" & cst_errmsg2 & "');" + vbCrLf
            strScript += "location.href ='SD_11_004.aspx?ProcessType=" & cst_ptBack & "&ID='+document.getElementById('Re_ID').value;" + vbCrLf
            strScript += "</script>"
            Page.RegisterStartupScript("", strScript)
            Exit Sub
        End If
    End Sub

    '檢查。
    Function Check_Stud_QuestionFac() As Boolean
        Dim rst As Boolean = False
        Dim msg As String = ""
        If Q1_6_CHour.Text <> "" Then
            Q1_6_CHour.Text = Trim(Q1_6_CHour.Text)
            If Not IsNumeric(Q1_6_CHour.Text) Then
                msg += "訓練時數應增加必須為數字" & vbCrLf
            Else
                Try
                    If CInt(Q1_6_CHour.Text) <= 0 Then msg += "訓練時數應增加 必須為大於0的數字，沒有(-)號" & vbCrLf
                    If CDbl(CInt(Q1_6_CHour.Text)) <> CDbl(Q1_6_CHour.Text) Then msg += "訓練時數應增加 必須為每小時，不可輸入浮點數" & vbCrLf
                Catch ex As Exception
                    msg += "訓練時數應增加 有誤：" & vbCrLf
                    msg += ex.ToString
                End Try
                If msg = "" Then Q1_6_CHour.Text = CInt(Q1_6_CHour.Text)
            End If
        End If
        If Q1_6_MHour.Text <> "" Then
            Q1_6_MHour.Text = Trim(Q1_6_MHour.Text)
            If Not IsNumeric(Q1_6_MHour.Text) Then
                msg += "訓練時數應減少 必須為數字" & vbCrLf
            Else
                Try
                    If CInt(Q1_6_MHour.Text) <= 0 Then msg += "訓練時數應減少 必須為大於0的數字，沒有(-)號" & vbCrLf
                    If CDbl(CInt(Q1_6_MHour.Text)) <> CDbl(Q1_6_MHour.Text) Then msg += "訓練時數應減少 必須為每小時，不可輸入浮點數" & vbCrLf
                Catch ex As Exception
                    msg += "訓練時數應減少 有誤：" & vbCrLf
                    msg += ex.ToString
                End Try
                If msg = "" Then
                    Q1_6_MHour.Text = CInt(Q1_6_MHour.Text)
                End If
            End If
        End If
        If Q2_1_CHour.Text <> "" Then
            Q2_1_CHour.Text = Trim(Q2_1_CHour.Text)
            If Not IsNumeric(Q2_1_CHour.Text) Then
                msg += "術科時數應增加 必須為數字" & vbCrLf
            Else
                Try
                    If CInt(Q2_1_CHour.Text) <= 0 Then msg += "術科時數應增加 必須為大於0的數字，沒有(-)號" & vbCrLf
                    If CDbl(CInt(Q2_1_CHour.Text)) <> CDbl(Q2_1_CHour.Text) Then msg += "術科時數應增加 必須為每小時，不可輸入浮點數" & vbCrLf
                Catch ex As Exception
                    msg += "術科時數應增加 有誤：" & vbCrLf
                    msg += ex.ToString
                End Try
                If msg = "" Then Q2_1_CHour.Text = CInt(Q2_1_CHour.Text)
            End If
        End If
        If Q2_1_MHour.Text <> "" Then
            Q2_1_MHour.Text = Trim(Q2_1_MHour.Text)
            If Not IsNumeric(Q2_1_MHour.Text) Then
                msg += "術科時數應減少 必須為數字" & vbCrLf
            Else
                Try
                    If CInt(Q2_1_MHour.Text) <= 0 Then msg += "術科時數應減少 必須為大於0的數字，沒有(-)號" & vbCrLf
                    If CDbl(CInt(Q2_1_MHour.Text)) <> CDbl(Q2_1_MHour.Text) Then msg += "術科時數應減少 必須為每小時，不可輸入浮點數" & vbCrLf
                Catch ex As Exception
                    msg += "術科時數應減少 有誤：" & vbCrLf
                    msg += ex.ToString
                End Try
                If msg = "" Then Q2_1_MHour.Text = CInt(Q2_1_MHour.Text)
            End If
        End If

        '判斷第11項是否為數字**by Milor 20080414----start
        If Q11.Text <> "" Then
            Q11.Text = Trim(Q11.Text)
            If Not IsNumeric(Q11.Text) Then
                msg += "每年願意自費之金額 必須為數字" & vbCrLf
            Else
                Try
                    If CInt(Q11.Text) <= 0 Then msg += "每年願意自費之金額 必須為大於0的數字，沒有(-)號" & vbCrLf
                    If CDbl(CInt(Q11.Text)) <> CDbl(Q11.Text) Then msg += "每年願意自費之金額 不可輸入浮點數" & vbCrLf
                Catch ex As Exception
                    msg += "每年願意自費之金額 有誤：" & vbCrLf
                    msg += ex.ToString
                End Try
                If msg = "" Then Q11.Text = CInt(Q11.Text)
            End If
        End If
        '判斷第11項是否為數字**by Milor 20080414----end

        rst = True
        If msg <> "" Then
            rst = False
            Common.MessageBox(Me, msg)
        End If
        Return rst
    End Function

    '儲存前檢查。
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""
        If SOCID.SelectedValue = "" Then Errmsg += "學員資料有誤，請重新選擇。" & vbCrLf
        If Trim(Q12.Text) <> "" Then
            Q12.Text = Trim(Q12.Text)
            If Len(Q12.Text) > 200 Then Errmsg += "其他建議 長度超過系統範圍(200)" & vbCrLf
        Else
            'Errmsg += "請輸入 其他建議" & vbCrLf
        End If
        If HidDASOURCE.Value = "1" Then Errmsg += "無法逕行修改，僅供單位查詢。" & vbCrLf
        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    '儲存
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

#Region "(No Use)"

        'Dim sqlAdapter As SqlDataAdapter
        'Dim dr_row As DataRow
        'Dim dtTable As DataTable

        'Dim sql As String = ""
        'Dim I As Integer = 0
        'If Not Page.IsValid Then
        '    Exit Sub
        'End If

#End Region

        '檢查
        If Not Check_Stud_QuestionFac() Then Exit Sub '有問題離開。
        'Session("QuestionFacSearchStr") = Me.ViewState("QuestionFacSearchStr")
        If Not Session(ss_QuestionFacSearchStr) Is Nothing Then
            Me.ViewState(ss_QuestionFacSearchStr) = Session(ss_QuestionFacSearchStr)
            Session(ss_QuestionFacSearchStr) = Session(ss_QuestionFacSearchStr)
        Else
            Session(ss_QuestionFacSearchStr) = Me.ViewState(ss_QuestionFacSearchStr)
        End If
        Call SaveData1() '儲存
    End Sub

    '儲存
    Sub SaveData1()
        Call TIMS.OpenDbConn(objconn)
        'Dim dt As DataTable
        'Dim oCmd As SqlCommand
        Dim i_sql As String = ""
        Dim u_sql As String = ""

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT 'x' FROM STUD_QUESTIONFAC WHERE SOCID = @SOCID " & vbCrLf
        Dim sCmd As New SqlCommand(sql, objconn)

        sql = "" & vbCrLf
        sql &= " INSERT INTO STUD_QUESTIONFAC (" & vbCrLf
        sql &= "  SOCID ,Q1_1 ,Q1_2 ,Q1_3 ,Q1_4 ,Q1_5 ,Q1_6 ,Q1_6_CCOURNAME ,Q1_6_CHOUR ,Q1_6_MCOURNAME ,Q1_6_MHOUR " & vbCrLf
        sql &= "  ,Q2_1 ,Q2_1_CCOURNAME ,Q2_1_CHOUR ,Q2_1_MCOURNAME ,Q2_1_MHOUR ,Q2_2 ,Q2_3 ,Q2_4 ,Q2_5 " & vbCrLf
        sql &= "  ,Q3_1 ,Q3_2 ,Q3_3 ,Q4 ,Q5 ,Q5_NOTE_NEWS ,Q5_NOTE_OTHER ,Q6 ,Q6_NOTE1 ,Q6_NOTE2 ,Q7 ,Q7_8 ,Q7_9 " & vbCrLf
        sql &= "  ,Q8 ,Q9_1 ,Q9_2 ,Q9_3 ,Q10 ,Q11 ,Q12 ,Q14 ,DASOURCE ,MODIFYACCT ,MODIFYDATE " & vbCrLf
        sql &= "  ) VALUES (@SOCID ,@Q1_1 ,@Q1_2 ,@Q1_3 ,@Q1_4 ,@Q1_5 ,@Q1_6 ,@Q1_6_CCOURNAME ,@Q1_6_CHOUR ,@Q1_6_MCOURNAME ,@Q1_6_MHOUR " & vbCrLf
        sql &= "  ,@Q2_1 ,@Q2_1_CCOURNAME ,@Q2_1_CHOUR ,@Q2_1_MCOURNAME ,@Q2_1_MHOUR ,@Q2_2 ,@Q2_3 ,@Q2_4 ,@Q2_5 " & vbCrLf
        sql &= "  ,@Q3_1 ,@Q3_2 ,@Q3_3 ,@Q4 ,@Q5 ,@Q5_NOTE_NEWS ,@Q5_NOTE_OTHER ,@Q6 ,@Q6_NOTE1 ,@Q6_NOTE2 ,@Q7 ,@Q7_8 ,@Q7_9 " & vbCrLf
        sql &= "  ,@Q8 ,@Q9_1 ,@Q9_2 ,@Q9_3 ,@Q10 ,@Q11 ,@Q12 ,@Q14 ,@DASOURCE ,@MODIFYACCT ,GETDATE() ) " & vbCrLf
        i_sql = sql
        Dim iCmd As New SqlCommand(sql, objconn)

        sql = "" & vbCrLf
        sql &= " UPDATE STUD_QUESTIONFAC " & vbCrLf
        sql &= " SET Q1_1 = @Q1_1 ,Q1_2=@Q1_2 ,Q1_3 = @Q1_3 ,Q1_4 = @Q1_4 ,Q1_5 = @Q1_5 ,Q1_6 = @Q1_6 ,Q1_6_CCOURNAME = @Q1_6_CCOURNAME ,Q1_6_CHOUR = @Q1_6_CHOUR ,Q1_6_MCOURNAME = @Q1_6_MCOURNAME ,Q1_6_MHOUR = @Q1_6_MHOUR " & vbCrLf
        sql &= " ,Q2_1 = @Q2_1 ,Q2_1_CCOURNAME = @Q2_1_CCOURNAME ,Q2_1_CHOUR = @Q2_1_CHOUR ,Q2_1_MCOURNAME = @Q2_1_MCOURNAME ,Q2_1_MHOUR = @Q2_1_MHOUR ,Q2_2 = @Q2_2 ,Q2_3 = @Q2_3 ,Q2_4 = @Q2_4 ,Q2_5 = @Q2_5 " & vbCrLf
        sql &= " ,Q3_1 = @Q3_1 ,Q3_2 = @Q3_2 ,Q3_3 = @Q3_3 ,Q4 = @Q4 ,Q5 = @Q5 ,Q5_NOTE_NEWS = @Q5_NOTE_NEWS ,Q5_NOTE_OTHER = @Q5_NOTE_OTHER ,Q6 = @Q6 ,Q6_NOTE1 = @Q6_NOTE1 ,Q6_NOTE2 = @Q6_NOTE2 " & vbCrLf
        sql &= " ,Q7 = @Q7 ,Q7_8 = @Q7_8 ,Q7_9 = @Q7_9 ,Q8 = @Q8 ,Q9_1 = @Q9_1 ,Q9_2 = @Q9_2 ,Q9_3 = @Q9_3 ,Q10 = @Q10 ,Q11 = @Q11 ,Q12 = @Q12 ,Q14 = @Q14 ,DASOURCE = @DASOURCE ,MODIFYACCT = @MODIFYACCT ,MODIFYDATE = GETDATE() " & vbCrLf
        sql &= " WHERE SOCID = @SOCID " & vbCrLf
        u_sql = sql
        Dim uCmd As New SqlCommand(sql, objconn)

        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("SOCID", SqlDbType.VarChar).Value = SOCID.SelectedValue
            dt.Load(.ExecuteReader())
        End With

        Dim strQ1_6 As String = ""
        If Q1_61.Checked = True Then strQ1_6 = "1"
        If Q1_62.Checked = True Then strQ1_6 = "2"
        If Q1_63.Checked = True Then strQ1_6 = "3"

        Dim strQ2_1 As String = ""
        If Q2_11.Checked = True Then strQ2_1 = "1"
        If Q2_12.Checked = True Then strQ2_1 = "2"
        If Q2_13.Checked = True Then strQ2_1 = "3"

        Dim strQ5 As String = ""
        If Q5_1.Checked = True Then strQ5 = "1"
        If Q5_2.Checked = True Then strQ5 = "2"
        If Q5_3.Checked = True Then strQ5 = "3"
        If Q5_4.Checked = True Then strQ5 = "4"
        If Q5_5.Checked = True Then strQ5 = "5"
        If Q5_6.Checked = True Then strQ5 = "6"

        Dim ssQ6 As String = ""
        If Q6_3.Checked = True Then ssQ6 = 3
        If Q6_1.Checked = True And Q6_2.Checked = True Then ssQ6 = 3
        If Q6_1.Checked = True And Q6_2.Checked = False Then ssQ6 = 1
        If Q6_1.Checked = False And Q6_2.Checked = True Then ssQ6 = 2

        Try
            If dt.Rows.Count > 0 Then
                With uCmd
                    .Parameters.Clear()
                    .Parameters.Add("SOCID", SqlDbType.Decimal).Value = SOCID.SelectedValue
                    .Parameters.Add("Q1_1", SqlDbType.Decimal).Value = IIf(Q1_1.SelectedIndex = -1, Convert.DBNull, Q1_1.SelectedValue)
                    .Parameters.Add("Q1_2", SqlDbType.Decimal).Value = IIf(Q1_2.SelectedIndex = -1, Convert.DBNull, Q1_2.SelectedValue) 'Q1_2
                    .Parameters.Add("Q1_3", SqlDbType.Decimal).Value = IIf(Q1_3.SelectedIndex = -1, Convert.DBNull, Q1_3.SelectedValue) ' Q1_3
                    .Parameters.Add("Q1_4", SqlDbType.Decimal).Value = IIf(Q1_4.SelectedIndex = -1, Convert.DBNull, Q1_4.SelectedValue) ' Q1_4
                    .Parameters.Add("Q1_5", SqlDbType.Decimal).Value = IIf(Q1_5.SelectedIndex = -1, Convert.DBNull, Q1_5.SelectedValue) 'Q1_5
                    .Parameters.Add("Q1_6", SqlDbType.Decimal).Value = IIf(strQ1_6 = "", Convert.DBNull, Val(strQ1_6)) ' Q1_6
                    .Parameters.Add("Q1_6_CCOURNAME", SqlDbType.NVarChar).Value = IIf(Q1_6_CCourName.Text = "", Convert.DBNull, Q1_6_CCourName.Text) 'Q1_6_CCourName.Text
                    .Parameters.Add("Q1_6_CHOUR", SqlDbType.Decimal).Value = IIf(Q1_6_CHour.Text = "", Convert.DBNull, Q1_6_CHour.Text) 'Q1_6_CHour.Text
                    .Parameters.Add("Q1_6_MCOURNAME", SqlDbType.NVarChar).Value = IIf(Q1_6_MCourName.Text = "", Convert.DBNull, Q1_6_MCourName.Text) 'Q1_6_MCourName.Text
                    .Parameters.Add("Q1_6_MHOUR", SqlDbType.Decimal).Value = IIf(Q1_6_MHour.Text = "", Convert.DBNull, Q1_6_MHour.Text) 'Q1_6_MHour.Text
                    .Parameters.Add("Q2_1", SqlDbType.Decimal).Value = IIf(strQ2_1 = "", Convert.DBNull, strQ2_1)
                    .Parameters.Add("Q2_1_CCOURNAME", SqlDbType.NVarChar).Value = IIf(Q2_1_CCourName.Text = "", Convert.DBNull, Q2_1_CCourName.Text) 'Q2_1_CCourName.Text
                    .Parameters.Add("Q2_1_CHOUR", SqlDbType.Decimal).Value = IIf(Q2_1_CHour.Text = "", Convert.DBNull, Q2_1_CHour.Text) 'Q2_1_CHour.Text
                    .Parameters.Add("Q2_1_MCOURNAME", SqlDbType.NVarChar).Value = IIf(Q2_1_MCourName.Text = "", Convert.DBNull, Q2_1_MCourName.Text) 'Q2_1_MCourName.Text
                    .Parameters.Add("Q2_1_MHOUR", SqlDbType.Decimal).Value = IIf(Q2_1_MHour.Text = "", Convert.DBNull, Q2_1_MHour.Text) 'Q2_1_MHour.Text
                    .Parameters.Add("Q2_2", SqlDbType.Decimal).Value = IIf(Q2_2.SelectedIndex = -1, Convert.DBNull, Q2_2.SelectedValue) 'Q2_2
                    .Parameters.Add("Q2_3", SqlDbType.Decimal).Value = IIf(Q2_3.SelectedIndex = -1, Convert.DBNull, Q2_3.SelectedValue) ' Q2_3
                    .Parameters.Add("Q2_4", SqlDbType.Decimal).Value = IIf(Q2_4.SelectedIndex = -1, Convert.DBNull, Q2_4.SelectedValue) ' Q2_4
                    .Parameters.Add("Q2_5", SqlDbType.Decimal).Value = IIf(Q2_5.SelectedIndex = -1, Convert.DBNull, Q2_5.SelectedValue) ' Q2_5
                    .Parameters.Add("Q3_1", SqlDbType.Decimal).Value = IIf(Q3_1.SelectedIndex = -1, Convert.DBNull, Q3_1.SelectedValue) 'Q3_1
                    .Parameters.Add("Q3_2", SqlDbType.Decimal).Value = IIf(Q3_2.SelectedIndex = -1, Convert.DBNull, Q3_2.SelectedValue) 'Q3_2
                    .Parameters.Add("Q3_3", SqlDbType.Decimal).Value = IIf(Q3_3.SelectedIndex = -1, Convert.DBNull, Q3_3.SelectedValue) ' Q3_3
                    .Parameters.Add("Q4", SqlDbType.Decimal).Value = IIf(Q4.SelectedIndex = -1, Convert.DBNull, Q4.SelectedValue) 'Q4
                    .Parameters.Add("Q5", SqlDbType.Decimal).Value = IIf(strQ5 = "", Convert.DBNull, strQ5) 'Q4  'Q5
                    .Parameters.Add("Q5_NOTE_NEWS", SqlDbType.NVarChar).Value = IIf(Q5_Note_News.Text = "", Convert.DBNull, Q5_Note_News.Text) 'Q5_Note_News.Text
                    .Parameters.Add("Q5_NOTE_OTHER", SqlDbType.NVarChar).Value = IIf(Q5_Note_Other.Text = "", Convert.DBNull, Q5_Note_Other.Text) 'Q5_Note_Other.Text
                    .Parameters.Add("Q6", SqlDbType.VarChar).Value = IIf(ssQ6 = "", Convert.DBNull, ssQ6)
                    .Parameters.Add("Q6_NOTE1", SqlDbType.NVarChar).Value = IIf(Q6_Note1.Text = "", Convert.DBNull, Q6_Note1.Text) 'Q6_Note1.Text
                    .Parameters.Add("Q6_NOTE2", SqlDbType.NVarChar).Value = IIf(Q6_Note2.Text = "", Convert.DBNull, Q6_Note2.Text) 'Q6_Note2.Text
                    .Parameters.Add("Q7", SqlDbType.Decimal).Value = IIf(Q7.SelectedIndex = -1, Convert.DBNull, Q7.SelectedValue) 'Q7
                    .Parameters.Add("Q8", SqlDbType.Decimal).Value = IIf(Q8.SelectedIndex = -1, Convert.DBNull, Q8.SelectedValue) 'Q8
                    '.Parameters.Add("Q9_NOTE", SqlDbType.NVarChar).Value = Convert.DBNull ' IIf(Q9_NOTE.text = "", Convert.DBNull, Q9_NOTE.text) 'Q9_NOTE
                    '.Parameters.Add("Q9", SqlDbType.Decimal).Value = IIf(Q9.SelectedIndex = -1, Convert.DBNull, Q9.SelectedValue) 'Q9
                    .Parameters.Add("Q9_1", SqlDbType.VarChar).Value = IIf(Q9_1.SelectedIndex = -1, Convert.DBNull, Q9_1.SelectedValue) 'Q9_1
                    .Parameters.Add("Q9_2", SqlDbType.VarChar).Value = IIf(Q9_2.SelectedIndex = -1, Convert.DBNull, Q9_2.SelectedValue) 'Q9_2
                    .Parameters.Add("Q9_3", SqlDbType.VarChar).Value = IIf(Q9_3.SelectedIndex = -1, Convert.DBNull, Q9_3.SelectedValue) 'Q9_3
                    .Parameters.Add("Q10", SqlDbType.VarChar).Value = IIf(Q10.SelectedIndex = -1, Convert.DBNull, Q10.SelectedValue) 'Q10
                    .Parameters.Add("Q11", SqlDbType.VarChar).Value = IIf(Q11.Text = "", Convert.DBNull, Q11.Text) 'Q11
                    .Parameters.Add("Q12", SqlDbType.NVarChar).Value = IIf(Q12.Text = "", Convert.DBNull, Q12.Text) 'Q12
                    .Parameters.Add("Q7_8", SqlDbType.VarChar).Value = IIf(Q7_8.SelectedIndex = -1, Convert.DBNull, Q7_8.SelectedValue) 'Q7_8
                    .Parameters.Add("Q7_9", SqlDbType.VarChar).Value = IIf(Q7_9.SelectedIndex = -1, Convert.DBNull, Q7_9.SelectedValue) 'Q7_9
                    .Parameters.Add("Q14", SqlDbType.VarChar).Value = IIf(Q14.SelectedIndex = -1, Convert.DBNull, Q14.SelectedValue) 'Q14
                    .Parameters.Add("DASOURCE", SqlDbType.Decimal).Value = "2" '單位填寫'資料來源	Null:未知 1:報名網()(學員外網填寫。)2@TIMS系統()
                    .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID ' MODIFYACCT
                    '.Parameters.Add("MODIFYDATE", SqlDbType.VarChar).Value = MODIFYDATE
                    '.ExecuteNonQuery()
                End With
                DbAccess.ExecuteNonQuery(u_sql, objconn, uCmd.Parameters)
            Else
                With iCmd
                    .Parameters.Clear()
                    .Parameters.Add("Q1_1", SqlDbType.Decimal).Value = IIf(Q1_1.SelectedIndex = -1, Convert.DBNull, Q1_1.SelectedValue)
                    .Parameters.Add("Q1_2", SqlDbType.Decimal).Value = IIf(Q1_2.SelectedIndex = -1, Convert.DBNull, Q1_2.SelectedValue) 'Q1_2
                    .Parameters.Add("Q1_3", SqlDbType.Decimal).Value = IIf(Q1_3.SelectedIndex = -1, Convert.DBNull, Q1_3.SelectedValue) ' Q1_3
                    .Parameters.Add("Q1_4", SqlDbType.Decimal).Value = IIf(Q1_4.SelectedIndex = -1, Convert.DBNull, Q1_4.SelectedValue) ' Q1_4
                    .Parameters.Add("Q1_5", SqlDbType.Decimal).Value = IIf(Q1_5.SelectedIndex = -1, Convert.DBNull, Q1_5.SelectedValue) 'Q1_5
                    .Parameters.Add("Q1_6", SqlDbType.Decimal).Value = IIf(strQ1_6 = "", Convert.DBNull, Val(strQ1_6)) ' Q1_6
                    .Parameters.Add("Q1_6_CCOURNAME", SqlDbType.NVarChar).Value = IIf(Q1_6_CCourName.Text = "", Convert.DBNull, Q1_6_CCourName.Text) 'Q1_6_CCourName.Text
                    .Parameters.Add("Q1_6_CHOUR", SqlDbType.Decimal).Value = IIf(Q1_6_CHour.Text = "", Convert.DBNull, Q1_6_CHour.Text) 'Q1_6_CHour.Text
                    .Parameters.Add("Q1_6_MCOURNAME", SqlDbType.NVarChar).Value = IIf(Q1_6_MCourName.Text = "", Convert.DBNull, Q1_6_MCourName.Text) 'Q1_6_MCourName.Text
                    .Parameters.Add("Q1_6_MHOUR", SqlDbType.Decimal).Value = IIf(Q1_6_MHour.Text = "", Convert.DBNull, Q1_6_MHour.Text) 'Q1_6_MHour.Text
                    .Parameters.Add("Q2_1", SqlDbType.Decimal).Value = IIf(strQ2_1 = "", Convert.DBNull, strQ2_1)
                    .Parameters.Add("Q2_1_CCOURNAME", SqlDbType.NVarChar).Value = IIf(Q2_1_CCourName.Text = "", Convert.DBNull, Q2_1_CCourName.Text) 'Q2_1_CCourName.Text
                    .Parameters.Add("Q2_1_CHOUR", SqlDbType.Decimal).Value = IIf(Q2_1_CHour.Text = "", Convert.DBNull, Q2_1_CHour.Text) 'Q2_1_CHour.Text
                    .Parameters.Add("Q2_1_MCOURNAME", SqlDbType.NVarChar).Value = IIf(Q2_1_MCourName.Text = "", Convert.DBNull, Q2_1_MCourName.Text) 'Q2_1_MCourName.Text
                    .Parameters.Add("Q2_1_MHOUR", SqlDbType.Decimal).Value = IIf(Q2_1_MHour.Text = "", Convert.DBNull, Q2_1_MHour.Text) 'Q2_1_MHour.Text
                    .Parameters.Add("Q2_2", SqlDbType.Decimal).Value = IIf(Q2_2.SelectedIndex = -1, Convert.DBNull, Q2_2.SelectedValue) 'Q2_2
                    .Parameters.Add("Q2_3", SqlDbType.Decimal).Value = IIf(Q2_3.SelectedIndex = -1, Convert.DBNull, Q2_3.SelectedValue) ' Q2_3
                    .Parameters.Add("Q2_4", SqlDbType.Decimal).Value = IIf(Q2_4.SelectedIndex = -1, Convert.DBNull, Q2_4.SelectedValue) ' Q2_4
                    .Parameters.Add("Q2_5", SqlDbType.Decimal).Value = IIf(Q2_5.SelectedIndex = -1, Convert.DBNull, Q2_5.SelectedValue) ' Q2_5
                    .Parameters.Add("Q3_1", SqlDbType.Decimal).Value = IIf(Q3_1.SelectedIndex = -1, Convert.DBNull, Q3_1.SelectedValue) 'Q3_1
                    .Parameters.Add("Q3_2", SqlDbType.Decimal).Value = IIf(Q3_2.SelectedIndex = -1, Convert.DBNull, Q3_2.SelectedValue) 'Q3_2
                    .Parameters.Add("Q3_3", SqlDbType.Decimal).Value = IIf(Q3_3.SelectedIndex = -1, Convert.DBNull, Q3_3.SelectedValue) ' Q3_3
                    .Parameters.Add("Q4", SqlDbType.Decimal).Value = IIf(Q4.SelectedIndex = -1, Convert.DBNull, Q4.SelectedValue) 'Q4
                    .Parameters.Add("Q5", SqlDbType.Decimal).Value = IIf(strQ5 = "", Convert.DBNull, strQ5) 'Q4  'Q5
                    .Parameters.Add("Q5_NOTE_NEWS", SqlDbType.NVarChar).Value = IIf(Q5_Note_News.Text = "", Convert.DBNull, Q5_Note_News.Text) 'Q5_Note_News.Text
                    .Parameters.Add("Q5_NOTE_OTHER", SqlDbType.NVarChar).Value = IIf(Q5_Note_Other.Text = "", Convert.DBNull, Q5_Note_Other.Text) 'Q5_Note_Other.Text
                    .Parameters.Add("Q6", SqlDbType.VarChar).Value = IIf(ssQ6 = "", Convert.DBNull, ssQ6)
                    .Parameters.Add("Q6_NOTE1", SqlDbType.NVarChar).Value = IIf(Q6_Note1.Text = "", Convert.DBNull, Q6_Note1.Text) 'Q6_Note1.Text
                    .Parameters.Add("Q6_NOTE2", SqlDbType.NVarChar).Value = IIf(Q6_Note2.Text = "", Convert.DBNull, Q6_Note2.Text) 'Q6_Note2.Text
                    .Parameters.Add("Q7", SqlDbType.Decimal).Value = IIf(Q7.SelectedIndex = -1, Convert.DBNull, Q7.SelectedValue) 'Q7
                    .Parameters.Add("Q8", SqlDbType.Decimal).Value = IIf(Q8.SelectedIndex = -1, Convert.DBNull, Q8.SelectedValue) 'Q8
                    '.Parameters.Add("Q9_NOTE", SqlDbType.NVarChar).Value = Convert.DBNull ' IIf(Q9_NOTE.text = "", Convert.DBNull, Q9_NOTE.text) 'Q9_NOTE
                    '.Parameters.Add("Q9", SqlDbType.Decimal).Value = IIf(Q9.SelectedIndex = -1, Convert.DBNull, Q9.SelectedValue) 'Q9
                    .Parameters.Add("Q9_1", SqlDbType.VarChar).Value = IIf(Q9_1.SelectedIndex = -1, Convert.DBNull, Q9_1.SelectedValue) 'Q9_1
                    .Parameters.Add("Q9_2", SqlDbType.VarChar).Value = IIf(Q9_2.SelectedIndex = -1, Convert.DBNull, Q9_2.SelectedValue) 'Q9_2
                    .Parameters.Add("Q9_3", SqlDbType.VarChar).Value = IIf(Q9_3.SelectedIndex = -1, Convert.DBNull, Q9_3.SelectedValue) 'Q9_3
                    .Parameters.Add("Q10", SqlDbType.VarChar).Value = IIf(Q10.SelectedIndex = -1, Convert.DBNull, Q10.SelectedValue) 'Q10
                    .Parameters.Add("Q11", SqlDbType.VarChar).Value = IIf(Q11.Text = "", Convert.DBNull, Q11.Text) 'Q11
                    .Parameters.Add("Q12", SqlDbType.NVarChar).Value = IIf(Q12.Text = "", Convert.DBNull, Q12.Text) 'Q12
                    .Parameters.Add("Q7_8", SqlDbType.VarChar).Value = IIf(Q7_8.SelectedIndex = -1, Convert.DBNull, Q7_8.SelectedValue) 'Q7_8
                    .Parameters.Add("Q7_9", SqlDbType.VarChar).Value = IIf(Q7_9.SelectedIndex = -1, Convert.DBNull, Q7_9.SelectedValue) 'Q7_9
                    .Parameters.Add("Q14", SqlDbType.VarChar).Value = IIf(Q14.SelectedIndex = -1, Convert.DBNull, Q14.SelectedValue) 'Q14
                    .Parameters.Add("DASOURCE", SqlDbType.Decimal).Value = "2" '單位填寫'資料來源	Null:未知 1:報名網()(學員外網填寫。)2@TIMS系統()
                    .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID ' MODIFYACCT
                    '.Parameters.Add("MODIFYDATE", SqlDbType.VarChar).Value = MODIFYDATE
                    .Parameters.Add("SOCID", SqlDbType.Decimal).Value = SOCID.SelectedValue
                    '.ExecuteNonQuery()
                End With
                DbAccess.ExecuteNonQuery(i_sql, objconn, iCmd.Parameters)
            End If
            Common.AddClientScript(Page, "insert_next();")
        Catch ex As Exception
            Common.MessageBox(Me, "!!儲存失敗!!")
            Common.MessageBox(Me, ex.ToString)
            'Throw ex
        End Try
    End Sub

    '資料來源 0:未填寫或未知 1: 報名網(學員外網填寫。) 2: TIMS系統 
    Function Chk_DASOURCE(ByVal StrSOCID As String) As Integer
        Dim rst As Integer = 0
        Dim sqlstr As String = ""
        sqlstr = " SELECT * FROM Stud_QuestionFac WHERE SOCID = '" & StrSOCID & "' "
        Dim row_list As DataRow = DbAccess.GetOneRow(sqlstr, objconn)
        If Not row_list Is Nothing Then
            Select Case Convert.ToString(row_list("DASOURCE"))
                Case "1", "2"
                    rst = Val(row_list("DASOURCE"))
            End Select
        End If
        Return rst
    End Function
End Class