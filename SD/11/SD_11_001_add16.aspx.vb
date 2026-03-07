Partial Class SD_11_001_add16
    Inherits AuthBasePage

    'Stud_Survey,V_STUDQUESTION1 VIEW_QUESTIONARY4
    'SELECT * FROM KEY_SURVEYKIND WHERE SVID=4
    'SELECT * FROM ID_SURVEYQUESTION WHERE SVID=4
    'WITH WC1 AS (SELECT * FROM ID_SURVEYQUESTION WHERE SVID=4) SELECT * FROM ID_SURVEYANSWER WHERE SQID IN (SELECT SQID FROM WC1)
    Const cst_SVID4 As Integer = 4 '問卷編號
    Const cst_SV_EDIT As String = "E" 'E:修改 C:檢視
    Const cst_SV_VIEW As String = "C" 'E:修改 C:檢視
    Const cst_ptPrint As String = "Print" '列印空白表
    Const cst_ptInsert As String = "insert" '新增
    Const cst_ptCheck As String = "check" '查詢
    Const cst_ptEdit As String = "Edit" '修改
    Const cst_ptPrint2 As String = "Print2" '列印單1學員
    Const cst_ptDel As String = "del" '清除重填
    Const cst_errMsg2 As String = "學員資料有誤，請聯絡系統管理員！！"
    Const cst_rSD11001aspx As String = "SD_11_001.aspx"

    '只是列印
    'Dim iGy As Integer = 0  '判斷是否有學員填寫問卷預設為 0 (沒有)
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

#Region "(No Use)"

        'iGy = 0
        'Dim dt As DataTable = TIMS.Get_dtSdS(HID_SVID.Value, objconn)
        'If dt.Rows.Count <> 0 Then
        '    iGy = 1  '判斷是否有學員填寫答案 1表示有
        'End If
        'If Me.Request("Type") = "V" Then
        '    'RETop.Value = "關閉本頁"
        '    iGy = 2  '判斷是否顯示功能鈕 2表示外面呼叫
        'End If

#End Region
        Re_ID.Value = TIMS.sUtl_GetRqValue(Me, "ID", Re_ID.Value)
        HID_SVID.Value = TIMS.sUtl_GetRqValue(Me, "SVID", HID_SVID.Value)
        Hid_SOCID.Value = TIMS.sUtl_GetRqValue(Me, "SOCID", Hid_SOCID.Value)
        Hid_OCID.Value = TIMS.sUtl_GetRqValue(Me, "OCID", Hid_OCID.Value)
        Hid_Stuedntid.Value = TIMS.sUtl_GetRqValue(Me, "Stuedntid", Hid_Stuedntid.Value)
        'Print (列印空白表)
        'Hid_ProcessType insert/Edit/Print2 (新增/修改/列印) SOCID
        Hid_ProcessType.Value = TIMS.sUtl_GetRqValue(Me, "ProcessType", Hid_ProcessType.Value)
        'Re_OCID.Value = TIMS.sUtl_GetRqValue(Me, "OCID", Re_OCID.Value)

        If Not IsPostBack Then
            btnSave1.Visible = False
            tbControl1.Visible = False
            tbControl2.Visible = False
            'Call ShowKeySurveyKind(HID_SVID.Value, PlaceHolder1) ', Me
            Select Case Hid_ProcessType.Value
                Case cst_ptPrint '列印空白表
                    Call ShowKeySurveyKind(HID_SVID.Value, PlaceHolder1) ', Me
                    'Me.RegisterStartupScript("scripprint", "<script>printDoc();self.location.href='SD_11_001.aspx?ID=" & Me.Re_ID.Value & "&ProcessType=Back2';</script>")
                Case Else
#Region "(No Use)"

                    'Dim rst1 As Boolean = True
                    'rst1 = SHOW_STUDENT1(Hid_OCID.Value, Hid_SOCID.Value, Hid_Stuedntid.Value)
                    'If Not rst1 Then
                    '    Dim rUrl As String = cst_rSD11001aspx & "?ID=" & Me.Re_ID.Value & "&ProcessType=Back2"
                    '    Common.MessageBox(Me, cst_errMsg2, rUrl)
                    '    Exit Sub
                    'End If

                    'Dim dtStudAnswer As DataTable = Nothing
                    'PlaceHolder1.Visible = True
                    'Dim ss As String = ""
                    'TIMS.SetMyValue(ss, "SelType1", cst_SV_EDIT)
                    'TIMS.SetMyValue(ss, "SOCID", Hid_SOCID.Value)
                    'Call ShowKeySurveyKind(cst_SVID4, PlaceHolder1, dtStudAnswer, ss, objconn)

#End Region
            End Select
        End If

        Select Case Hid_ProcessType.Value
            Case cst_ptPrint '列印空白表
                'Call ShowKeySurveyKind(HID_SVID.Value, PlaceHolder1) ', Me
                'Me.RegisterStartupScript("scripprint", "<script>printDoc();self.location.href='SD_11_001.aspx?ID=" & Me.Re_ID.Value & "&ProcessType=Back2';</script>")
            Case Else
                Dim rst1 As Boolean = True
                rst1 = SHOW_STUDENT1(Hid_OCID.Value, Hid_SOCID.Value, Hid_Stuedntid.Value)
                If Not rst1 Then
                    Dim rUrl As String = cst_rSD11001aspx & "?ID=" & Me.Re_ID.Value & "&ProcessType=Back2"
                    Common.MessageBox(Me, cst_errMsg2, rUrl)
                    Exit Sub
                End If
                Dim dtStudAnswer As DataTable = Nothing
                PlaceHolder1.Visible = True
                Dim ss As String = ""
                TIMS.SetMyValue(ss, "SelType1", cst_SV_EDIT)
                TIMS.SetMyValue(ss, "SOCID", Hid_SOCID.Value)
                Call ShowKeySurveyKind(cst_SVID4, PlaceHolder1, dtStudAnswer, ss, objconn)
        End Select
        Select Case Hid_ProcessType.Value
            Case cst_ptPrint '列印空白表
                'Me.RegisterStartupScript("scripprint", "<script>printDoc();self.location.href='SD_11_001.aspx?ID=" & Me.Re_ID.Value & "&ProcessType=Back2';</script>")
                'Me.RegisterStartupScript("scripprint", "<script>printDoc();window.opener=null;window.open('','_self');window.close();</script>")
                TIMS.RegisterStartupScript(Me, TIMS.xBlockName(), "<script>printDoc();</script>")
            Case cst_ptPrint2 '列印學員
                tbControl2.Visible = True
                'Me.RegisterStartupScript("scripprint", "<script>printDoc();self.location.href='SD_11_001.aspx?ID=" & Me.Re_ID.Value & "&ProcessType=Back2';</script>")
                'Me.RegisterStartupScript("scripprint", "<script>printDoc();window.opener=null;window.open('','_self');window.close();</script>")
                TIMS.RegisterStartupScript(Me, TIMS.xBlockName(), "<script>printDoc();</script>")
            Case cst_ptInsert
                tbControl1.Visible = True
                tbControl2.Visible = True
                btnSave1.Visible = True
            Case cst_ptEdit
                tbControl1.Visible = True
                tbControl2.Visible = True
                btnSave1.Visible = True
            Case cst_ptCheck
                tbControl1.Visible = True
                tbControl2.Visible = True
            Case cst_ptDel
                tbControl1.Visible = True
                tbControl2.Visible = True
                btnSave1.Visible = True
                Dim rUrl As String = ""
                If Hid_SOCID.Value = "" Then
                    rUrl = cst_rSD11001aspx & "?ID=" & Me.Re_ID.Value & "&ProcessType=Back2"
                    Common.MessageBox(Me, cst_errMsg2, rUrl)
                    Exit Sub
                End If
                Dim sql As String = ""
                sql = "DELETE STUD_SURVEY WHERE SOCID = @SOCID"
                Dim dCmd As New SqlCommand(sql, objconn)
                With dCmd
#Region "(目前不使用)"

                    '.Parameters.Clear()
                    '.Parameters.Add("SOCID", SqlDbType.VarChar).Value = Hid_SOCID.Value
                    '.ExecuteNonQuery()

#End Region
                    Dim myParam As Hashtable = New Hashtable
                    myParam.Add("SOCID", Hid_SOCID.Value)
                    DbAccess.ExecuteNonQuery(sql, objconn, myParam)
                End With
                rUrl = cst_rSD11001aspx & "?ID=" & Me.Re_ID.Value & "&ProcessType=Back2"
                TIMS.Utl_Redirect1(Me, rUrl)
                Exit Sub
        End Select
    End Sub

    '顯示學員資料1
    Function SHOW_STUDENT1(ByVal OCID As String, ByVal SOCID As String, ByVal StudentID As String) As Boolean
        Dim rst As Boolean = True
        Dim sqlstr As String = ""
        sqlstr = "" & vbCrLf
        sqlstr &= " SELECT b.studentid ,c.name ,b.StudStatus" & vbCrLf
        sqlstr &= " ,CONVERT(VARCHAR, b.RejectTDate1, 111) RejectTDate1" & vbCrLf
        sqlstr &= " ,CONVERT(VARCHAR, b.RejectTDate2, 111) RejectTDate2 " & vbCrLf
        sqlstr &= " FROM class_classinfo a " & vbCrLf
        sqlstr &= " JOIN class_studentsofclass b ON a.ocid = b.ocid " & vbCrLf
        sqlstr &= " JOIN stud_studentinfo c ON b.sid = c.sid " & vbCrLf
        sqlstr &= " where 1=1 " & vbCrLf
        sqlstr &= " AND b.OCID = @OCID " & vbCrLf
        sqlstr &= " AND b.SOCID = @SOCID " & vbCrLf
        sqlstr &= " AND b.StudentID = @StudentID " & vbCrLf
        Dim sCmd As New SqlCommand(sqlstr, objconn)
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCID
            .Parameters.Add("SOCID", SqlDbType.VarChar).Value = SOCID
            .Parameters.Add("StudentID", SqlDbType.VarChar).Value = StudentID
        End With
        Dim row As DataRow = TIMS.GetOneRow(sCmd, objconn)
        If row Is Nothing Then Return False
        labName.Text = Convert.ToString(row("name"))
        labStud.Text = Convert.ToString(row("studentid"))
        labStatus.Text = TIMS.GET_STUDSTATUS_N23(row("StudStatus"), row("RejectTDate1"), row("RejectTDate2"))
        Return rst
    End Function

    '顯示問卷
    Sub ShowKeySurveyKind(ByVal RsSVID As String, ByRef PH1 As PlaceHolder, ByRef dtsysAnswer As DataTable, ByVal ss As String, ByRef objconn As SqlConnection)
        'dtsysAnswer (SQID,SAID)
        Const cst_widthMNpx As String = "700px"
        Dim vSelType1 As String = TIMS.GetMyValue(ss, "SelType1") 'E:修改 C:檢視
        Dim vSOCID As String = TIMS.GetMyValue(ss, "SOCID") 'E:修改 C:檢視
        Dim RadioButtonList2 As RadioButtonList
        Dim CheckBoxList2 As CheckBoxList
        Dim TBT As HtmlTable
        Dim Row As HtmlTableRow
        Dim Cell As HtmlTableCell
        Dim SKID As String = ""
        Dim SQID As String = ""
        Dim sql As String = ""
        Dim sql2 As String = ""
        Dim sql3 As String = ""
        Dim sql4 As String = ""
        Dim ff As String = "" 'find
        Dim dt As DataTable = Nothing 'KEY_SURVEYKIND
        Dim dt2 As DataTable = Nothing 'ID_SurveyQuestion
        Dim dt3 As DataTable = Nothing 'ID_SurveyAnswer
        Dim dt4 As DataTable = Nothing 'Stud_Survey

        sql = " SELECT * FROM KEY_SURVEYKIND WHERE SVID = " & RsSVID & " ORDER BY SERIAL " '取出標題
        dt = DbAccess.GetDataTable(sql, objconn)
        For i As Integer = 0 To dt.Rows.Count - 1 '利用迴圈新增標題
            SKID = dt.Rows(i).Item("SKID").ToString '　取出SKID
            TBT = New HtmlTable
            PH1.Controls.Add(TBT) '新增TABLE
            TBT.ID = "TBT" & i
            'TBT.Attributes.Add("width", "100%")
            TBT.Attributes.Add("width", cst_widthMNpx)
            TBT.Attributes.Add("class", "font")
            Row = New HtmlTableRow         '新增TR
            TBT.Controls.Add(Row)
            Cell = New HtmlTableCell       '新增TD
            Cell.InnerText = dt.Rows(i).Item("TOPIC").ToString   '代入標題
            Cell.Attributes.Add("Style", "color@White;background-color: #2aafc0")   '設定顏色
            Cell.Attributes.Add("width", "100%")
            Cell.Attributes.Add("class", "font")
            Row.Controls.Add(Cell)
            'SELECT * FROM ID_SURVEYQUESTION WHERE ROWNUM <=10
            sql2 = "SELECT * FROM ID_SURVEYQUESTION WHERE SKID = " & SKID & " ORDER BY SERIAL" '取出問題的題目
            dt2 = DbAccess.GetDataTable(sql2, objconn)

            For j As Integer = 0 To dt2.Rows.Count - 1 '利用迴圈新增題目
                SQID = dt2.Rows(j).Item("SQID").ToString '取出SQID
                TBT = New HtmlTable
                PH1.Controls.Add(TBT) '新增TABLE
                'TBT.Attributes.Add("width", "100%")
                TBT.Attributes.Add("width", cst_widthMNpx)
                TBT.Attributes.Add("Style", "background-color: #e9f2fc")
                TBT.Attributes.Add("class", "font")
                Row = New HtmlTableRow '新增TR
                TBT.Controls.Add(Row)
                Cell = New HtmlTableCell '新增td
                Cell.Attributes.Add("width", "100%")
                Cell.Attributes.Add("class", "font")
                Cell.InnerText = dt2.Rows(j).Item("QUESTION").ToString '題目內容
                Row.Controls.Add(Cell)
                sql3 = " SELECT * FROM ID_SURVEYANSWER WHERE SQID = " & SQID & " ORDER BY Serial " '取出答案內容
                dt3 = DbAccess.GetDataTable(sql3, objconn)

                Select Case vSelType1 'MyPage.Request("Type")
                    Case "E", "C" 'E:修改 C:檢視 
                        'Select * from Stud_Survey WHERE ROWNUM <=10 '(學員答案)
                        sql4 = "SELECT SAID FROM STUD_SURVEY WHERE SQID = " & SQID & " and SOCID = " & vSOCID
                        dt4 = DbAccess.GetDataTable(sql4, objconn)
                End Select

                If dt3.Rows.Count <> 0 Then     '如果答案選項不是零
                    Row = New HtmlTableRow '新增TR
                    TBT.Controls.Add(Row)
                    Cell = New HtmlTableCell
                    Row.Controls.Add(Cell) '新增TD

                    If dt2.Rows(j).Item("QTYPE").ToString = 1 Then  '判斷是那程型態的題目 1是radio2,2是checkbox
                        RadioButtonList2 = New RadioButtonList      '新增 RadioButtonList
                        RadioButtonList2.ID = "SQID" & dt2.Rows(j).Item("SQID").ToString   'id = SQID 的值
                        RadioButtonList2.Attributes.Add("runat", "server")
                        RadioButtonList2.Attributes.Add("class", "font")
                        RadioButtonList2.RepeatDirection = RepeatDirection.Horizontal
                        Cell.Controls.Add(RadioButtonList2)
                        'RadioButtonList2.AutoPostBack = True
                        'AddHandler RadioButtonList2.SelectedIndexChanged, AddressOf sUtl_RBL2SedIndexChanged
                        'RadioButtonList2.Items.Clear()
                        For z As Integer = 0 To dt3.Rows.Count - 1 '計算出有幾個答案選項
                            Dim SAID2 As String = dt3.Rows(z).Item("SAID")
                            If RadioButtonList2.Items.FindByValue(SAID2) Is Nothing Then
                                RadioButtonList2.Items.Add(dt3.Rows(z).Item("ANSWER").ToString)    '新增RadioButton的TEXT 為答案的內容ANSWER
                                RadioButtonList2.Items.Item(z).Value = SAID2 'dt3.Rows(z).Item("SAID").ToString 'RadioButton的值為SAID
                            End If
                            Select Case vSelType1 'MyPage.Request("Type")
                                Case "E", "C" 'E:修改 C:檢視
                                    If dtsysAnswer Is Nothing And Not dt4 Is Nothing Then   '如果是修改 還有答案都有作答
                                        ff = "SAID='" & RadioButtonList2.Items.Item(z).Value & "'"  '計算學員作答結果
                                        If dt4.Select(ff).Length > 0 Then
                                            RadioButtonList2.Items.Item(z).Selected = True
                                        End If
                                    End If
                            End Select
                            If Not dtsysAnswer Is Nothing Then  '如果有答案沒有作答的
                                ff = "SAID='" & RadioButtonList2.Items.Item(z).Value & "'"  '計算學員作答結果
                                If dtsysAnswer.Select(ff).Length > 0 Then RadioButtonList2.Items.Item(z).Selected = True
                            End If
                        Next
                    Else                                         '2是checkbox
                        CheckBoxList2 = New CheckBoxList          '新增 CheckBoxList
                        CheckBoxList2.ID = "SQID" & dt2.Rows(j).Item(0).ToString   'id = SQID 的值
                        CheckBoxList2.Attributes.Add("runat", "server")
                        CheckBoxList2.Attributes.Add("class", "font")
                        CheckBoxList2.RepeatDirection = RepeatDirection.Horizontal
                        Cell.Controls.Add(CheckBoxList2)
                        'CheckBoxList2.Items.Clear()
                        For z As Integer = 0 To dt3.Rows.Count - 1 '計算出有幾個答案選項
                            Dim SAID2 As String = dt3.Rows(z).Item("SAID")
                            If CheckBoxList2.Items.FindByValue(SAID2) Is Nothing Then
                                CheckBoxList2.Items.Add(dt3.Rows(z).Item("ANSWER").ToString)    '新增 CheckBoxList 的TEXT內容
                                CheckBoxList2.Items.Item(z).Value = SAID2 'CheckBoxList 的值 = SAID
                            End If
                            Select Case vSelType1 'MyPage.Request("Type")
                                Case "E", "C" 'E:修改 C:檢視
                                    If dtsysAnswer Is Nothing And Not dt4 Is Nothing Then   '如果是修改 還有答案都有作答
                                        ff = "SAID='" & CheckBoxList2.Items.Item(z).Value & "'"  '計算學員作答結果
                                        If dt4.Select(ff).Length > 0 Then
                                            CheckBoxList2.Items.Item(z).Selected = True
                                        End If
                                    End If

                            End Select
                            If Not dtsysAnswer Is Nothing Then  '如果有答案沒有作答的
                                ff = "SAID='" & CheckBoxList2.Items.Item(z).Value & "'"  '計算學員作答結果
                                If dtsysAnswer.Select(ff).Length > 0 Then CheckBoxList2.Items.Item(z).Selected = True
                            End If
                        Next
                    End If
                End If
            Next
        Next
    End Sub

    '組合 Web 網頁動態方式伺服器控制項
    Sub ShowKeySurveyKind(ByVal RsSVID As String, ByRef PH1 As PlaceHolder)
        ', ByRef MyPage As Page
        Dim CheckBoxList2 As CheckBoxList
        Dim RadioButtonList2 As RadioButtonList
        Dim Label2 As Label
        Dim SKID2 As String = ""
        Dim dt2 As DataTable = Nothing
        Dim dt3 As DataTable = Nothing
        Dim Tb As HtmlTable = Nothing
        Dim Tb2 As HtmlTable = Nothing
        Dim htRow As HtmlTableRow = Nothing
        Dim htCell As HtmlTableCell = Nothing
        'Const cst_maxwidth As String = "720px"
        Const cst_maxwidth As String = "100%"
        Dim dt As DataTable = Nothing
        Dim sql As String = ""
        sql = " SELECT SKID, TOPIC, SERIAL FROM KEY_SURVEYKIND WHERE SVID = '" & RsSVID & "' ORDER BY SERIAL "
        dt = DbAccess.GetDataTable(sql, objconn) '找出標題

        If dt.Rows.Count > 0 Then
            For v As Integer = 0 To dt.Rows.Count - 1  '利用迴圈新增標題
                SKID2 = dt.Rows(v).Item("SKID").ToString   '取出 SKID
                Tb = New HtmlTable                    '新增TB
                PH1.Controls.Add(Tb)
                'Tb.Attributes.Add("width", "100%")
                Tb.Attributes.Add("width", cst_maxwidth)
                Tb.Attributes.Add("class", "font12")
                htRow = New HtmlTableRow                '新增TR
                Tb.Controls.Add(htRow)
                htCell = New HtmlTableCell               '新增TD
                htRow.Controls.Add(htCell)
                With htCell
                    .InnerText = dt.Rows(v).Item("TOPIC").ToString '新增標題
                    '.Attributes.Add("Style", "color@White;background-color: #2aafc0")
                    .Attributes.Add("Style", "background-color: #2aafc0")
                    .Attributes.Add("width", "100%")
                    .Attributes.Add("height", "40px")
                    .Attributes.Add("class", "font14")
                End With
                htCell = New HtmlTableCell               '新增TD
                htRow.Controls.Add(htCell)
                With htCell
                    .Attributes.Add("width", "100%")
                    .Attributes.Add("align", "right")
                    .Attributes.Add("class", "font12")
                End With
                Tb2 = New HtmlTable                      '新增TB
                PH1.Controls.Add(Tb2)
                With Tb2
                    '.Attributes.Add("width", "100%")
                    .Attributes.Add("width", cst_maxwidth)
                    .Attributes.Add("CELLSPACING", 0)
                    .Attributes.Add("cellPadding", 0)
                    .Attributes.Add("BORDER", 0)
                    .Attributes.Add("class", "font12")
                    .ID = "Tb2" & v
                End With
                Dim sqlQ As String = ""
                sqlQ = ""
                sqlQ &= " SELECT SQID, QUESTION, QTYPE, SERIAL, SVID, SKID "
                sqlQ &= " FROM ID_SURVEYQUESTION "
                sqlQ &= " WHERE SKID = " & SKID2 & ""
                sqlQ &= " ORDER BY SERIAL " '找出問題題目
                dt2 = DbAccess.GetDataTable(sqlQ, objconn)

                If dt2.Rows.Count > 0 Then
                    For k As Integer = 0 To dt2.Rows.Count - 1
                        Dim SQID2 As String = dt2.Rows(k).Item("SQID").ToString 'SQID
                        htRow = New HtmlTableRow                          '新增tr
                        'MyPage.FindControl("Tb2" & v).Controls.Add(htRow)
                        Tb2.Controls.Add(htRow)
                        With htRow
                            .Attributes.Add("Style", "background-color: #e9f2fc")
                            .Attributes.Add("width", "100%")
                            .Attributes.Add("class", "font12")
                        End With
                        htCell = New HtmlTableCell                        '新增td
                        htRow.Controls.Add(htCell)
                        Label2 = New Label
                        With htCell
                            .Controls.Add(Label2)
                            .Attributes.Add("width", "90%")
                            .Attributes.Add("class", "font12")
                        End With
                        Label2.Text = dt2.Rows(k).Item("QUESTION").ToString '題目內容
                        htCell = New HtmlTableCell                         '新增td
                        htRow.Controls.Add(htCell)
                        htCell.Attributes.Add("align", "right")
                        Dim sqlA As String = ""
                        sqlA = ""
                        sqlA &= " SELECT SAID, ANSWER, SCORE, SERIAL, SQID "
                        sqlA &= " FROM ID_SurveyAnswer WHERE SQID = " & SQID2 & " ORDER BY Serial "
                        dt3 = DbAccess.GetDataTable(sqlA, objconn)

                        If dt3.Rows.Count > 0 Then
                            htRow = New HtmlTableRow                           '新增TR
                            'MyPage.FindControl("Tb2" & v).Controls.Add(htRow)
                            Tb2.Controls.Add(htRow)
                            With htRow
                                .Attributes.Add("width", "100%")
                                .Attributes.Add("Style", "background-color: #e9f2fc")
                            End With
                            htCell = New HtmlTableCell
                            htRow.Controls.Add(htCell)                           '新增TD
                            '判斷是那程型態的題目 1是radio ,2是checkbox
                            Select Case Convert.ToString(dt2.Rows(k).Item("QTYPE"))
                                Case "1"
                                    'RadioButtonList
                                    RadioButtonList2 = New RadioButtonList
                                    RadioButtonList2.Attributes.Add("class", "font12")
                                    RadioButtonList2.RepeatDirection = RepeatDirection.Horizontal
                                    htCell.Attributes.Add("colspan", 2)
                                    htCell.Controls.Add(RadioButtonList2)
                                    For z As Integer = 0 To dt3.Rows.Count - 1
                                        Dim sANSWER As String = dt3.Rows(z).Item("ANSWER").ToString & "&nbsp;&nbsp;"
                                        RadioButtonList2.Items.Add(sANSWER)
                                    Next
                                Case Else
                                    'checkbox
                                    CheckBoxList2 = New CheckBoxList
                                    CheckBoxList2.Attributes.Add("class", "font12")
                                    htCell.Attributes.Add("colspan", 2)
                                    htCell.Controls.Add(CheckBoxList2)
                                    For z As Integer = 0 To dt3.Rows.Count - 1
                                        Dim sANSWER As String = dt3.Rows(z).Item("ANSWER").ToString & "&nbsp;&nbsp;"
                                        CheckBoxList2.Items.Add(sANSWER)
                                    Next
                            End Select
                            htCell = New HtmlTableCell
                            htRow.Controls.Add(htCell)
                        End If
                    Next
                End If
            Next
        End If
    End Sub

    Protected Sub btnBack1_Click(sender As Object, e As EventArgs) Handles btnBack1.Click
        '"SD_11_001.aspx" 
        Dim sUrl As String = cst_rSD11001aspx & "?ID=" & Me.Re_ID.Value & "&ProcessType=Back2"
        TIMS.Utl_Redirect1(Me, sUrl)
    End Sub

    Protected Sub btnSave1_Click(sender As Object, e As EventArgs) Handles btnSave1.Click
        Dim ss As String = ""
        TIMS.SetMyValue(ss, "SOCID", Hid_SOCID.Value)
        TIMS.SetMyValue(ss, "UserID", sm.UserInfo.UserID)
        Call SaveStSurvey(cst_SVID4, PlaceHolder1, ss, objconn)
        Dim rUrl As String = cst_rSD11001aspx & "?ID=" & Me.Re_ID.Value & "&ProcessType=Back2"
        Common.MessageBox(Me, "儲存成功", rUrl)

#Region "(No Use)"

        'Exit Sub
        'Common.MessageBox(Me, "儲存成功")
        'Dim sUrl As String = "SD_11_001.aspx?ID=" & Me.Re_ID.Value & "&ProcessType=Back2"
        'Response.Redirect(sUrl)

#End Region
    End Sub

    '儲存2 (STUD_SURVEY)
    Sub SaveStSurvey(ByVal SVID As Integer, ByRef PH1 As PlaceHolder, ByVal SS As String, ByRef objconn As SqlConnection)
        Call TIMS.OpenDbConn(objconn)
        Dim vSOCID As String = TIMS.GetMyValue(SS, "SOCID")
        Dim vUserID As String = TIMS.GetMyValue(SS, "UserID")
        'Dim PN1 As Panel = MCLS1.Panel1
        'Dim a As Integer '標題的計數
        'Dim b As Integer '題目的計數
        'Dim c As Integer '答案的計數
        Dim SKID2 As String
        Dim SQID2 As String
        Dim SAID As String
        Dim RBL As RadioButtonList
        Dim CHL As CheckBoxList
        Dim sql4 As String = ""
        Dim sql3 As String = ""
        Dim sql2 As String = ""
        Dim dt As DataTable 'KEY_SURVEYKIND
        Dim dt2 As DataTable 'ID_SurveyQuestion
        Dim dt3 As DataTable 'ID_SurveyAnswer
        'Dim dt4 As DataTable 'Stud_Survey
        'If sm.UserInfo.UserID Is Nothing Then sm.UserInfo.UserID = Request("IDNO") '有可能是線上填寫

        Dim sql As String = ""
        'RadioButtonList (select)
        sql = "" & vbCrLf
        sql &= " SELECT SSID, SQID, SAID " & vbCrLf
        sql &= " FROM STUD_SURVEY " & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        sql &= " AND SOCID = @SOCID " & vbCrLf
        sql &= " AND SQID = @SQID " & vbCrLf
        Dim sCmd As New SqlCommand(sql, objconn)

        'checkboxlist (delete)
        sql = "" & vbCrLf
        sql &= " DELETE STUD_SURVEY " & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        sql &= " AND SOCID = @SOCID " & vbCrLf
        sql &= " AND SQID = @SQID " & vbCrLf
        Dim dCmd As New SqlCommand(sql, objconn)
        Dim dSql As String = sql

        'RadioButtonList /checkboxlist (INSERT)
        sql = "" & vbCrLf
        sql &= " INSERT INTO STUD_SURVEY(SSID, SOCID, DONEDATE, SVID, SKID, SQID, SAID, MODIFYACCT, MODIFYDATE) " & vbCrLf '/*PK*/
        sql &= " VALUES(@SSID, @SOCID, dbo.TRUNC_DATETIME(GETDATE()), @SVID, @SKID, @SQID, @SAID, @MODIFYACCT, GETDATE()) "
        Dim iCmd As New SqlCommand(sql, objconn)
        Dim iSql As String = sql

        'RadioButtonList (update)
        sql = "" & vbCrLf
        sql &= " UPDATE STUD_SURVEY " & vbCrLf
        sql &= " SET SAID = @SAID ,MODIFYACCT = @MODIFYACCT ,MODIFYDATE = GETDATE() " & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND SOCID=@SOCID" & vbCrLf
        sql &= " AND SQID=@SQID" & vbCrLf
        Dim uCmd As New SqlCommand(sql, objconn)
        Dim uSql As String = sql

        sql = "SELECT * FROM KEY_SURVEYKIND WHERE SVID = " & SVID & " order by serial " '取出標題
        dt = DbAccess.GetDataTable(sql, objconn)
        For a As Integer = 0 To dt.Rows.Count - 1
            SKID2 = dt.Rows(a).Item("SKID").ToString

            sql2 = "SELECT * FROM ID_SURVEYQUESTION Where SKID = " & SKID2 & " order by Serial" '取出問題的題目
            dt2 = DbAccess.GetDataTable(sql2, objconn)
            For b As Integer = 0 To dt2.Rows.Count - 1
                SQID2 = dt2.Rows(b).Item("SQID").ToString
                sql3 = "SELECT * FROM ID_SURVEYANSWER WHERE SQID = " & SQID2 & " order by Serial" '取出答案內容
                dt3 = DbAccess.GetDataTable(sql3, objconn)

                Dim objName As String = "SQID" & SQID2
                Select Case CStr(dt2.Rows(b).Item("QTYPE"))
                    Case "1" '如果是RadioButtonList
                        'ViewState("sysAnswer")
                        RBL = New RadioButtonList
                        RBL = CType(PH1.FindControl(objName), RadioButtonList)
                        'RBL = DirectCast(MCLS1.FindControl("SQID" & dt2.Rows(b).Item("SQID").ToString), RadioButtonList) '取得 RadioButtonList
                        For c As Integer = 0 To dt3.Rows.Count - 1
                            If Not RBL Is Nothing AndAlso RBL.Items.Item(c).Selected Then
                                SAID = RBL.Items.Item(c).Value
                                Dim dtSS As New DataTable
                                With sCmd
                                    .Parameters.Clear()
                                    .Parameters.Add("SOCID", SqlDbType.VarChar).Value = vSOCID
                                    .Parameters.Add("SQID", SqlDbType.VarChar).Value = SQID2
                                    dtSS.Load(.ExecuteReader())
                                End With
                                If dtSS.Rows.Count = 0 Then
                                    Dim iSSID As Integer = DbAccess.GetNewId(objconn, "STUD_SURVEY_SSID_SEQ,STUD_SURVEY,SSID")
                                    With iCmd

                                        Dim myParam As Hashtable = New Hashtable
                                        myParam.Add("SSID", iSSID)
                                        myParam.Add("SOCID", vSOCID)
                                        myParam.Add("SVID", SVID)
                                        myParam.Add("SKID", Val(SKID2))
                                        myParam.Add("SQID", Val(SQID2))
                                        myParam.Add("SAID", Val(SAID))
                                        myParam.Add("MODIFYACCT", vUserID)
                                        DbAccess.ExecuteNonQuery(iSql, objconn, myParam)
                                    End With
                                Else
                                    With uCmd

                                        Dim myParam As Hashtable = New Hashtable
                                        myParam.Add("SAID", Val(SAID))
                                        myParam.Add("MODIFYACCT", vUserID)
                                        myParam.Add("SOCID", vSOCID)
                                        myParam.Add("SQID", Val(SQID2))
                                        DbAccess.ExecuteNonQuery(uSql, objconn, myParam)
                                    End With
                                End If
                            End If
                        Next

                    Case "2" '如果是checkboxlist
                        CHL = New CheckBoxList
                        CHL = CType(PH1.FindControl(objName), CheckBoxList)
                        'CHL = DirectCast(MCLS1.FindControl("SQID" & dt2.Rows(b).Item("SQID").ToString), CheckBoxList) '取得checkboxlist
                        If Not CHL Is Nothing Then
                            With dCmd '刪除

                                Dim myParam As Hashtable = New Hashtable
                                myParam.Add("SOCID", vSOCID)
                                myParam.Add("SQID", SQID2)
                                DbAccess.ExecuteNonQuery(dSql, objconn, myParam)
                            End With
                        End If

                        For c As Integer = 0 To dt3.Rows.Count - 1
                            If Not CHL Is Nothing AndAlso CHL.Items.Item(c).Selected Then
                                SAID = CHL.Items.Item(c).Value
                                Dim iSSID As Integer = DbAccess.GetNewId(objconn, "STUD_SURVEY_SSID_SEQ,STUD_SURVEY,SSID")
                                With iCmd

                                    Dim myParam As Hashtable = New Hashtable
                                    myParam.Add("SSID", iSSID)
                                    myParam.Add("SOCID", vSOCID)
                                    myParam.Add("SVID", SVID)
                                    myParam.Add("SKID", Val(SKID2))
                                    myParam.Add("SQID", Val(SQID2))
                                    myParam.Add("SAID", Val(SAID))
                                    myParam.Add("MODIFYACCT", vUserID)
                                    DbAccess.ExecuteNonQuery(iSql, objconn, myParam)
                                End With
                            End If
                        Next
                End Select
            Next
        Next
    End Sub
End Class