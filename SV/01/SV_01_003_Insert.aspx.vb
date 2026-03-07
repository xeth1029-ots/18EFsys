Partial Class SV_01_003_Insert
    Inherits AuthBasePage

    'Dim j As Integer
    'Dim v As Integer '計算標題的筆數
    'Dim k As Integer
    'Dim i As Integer
    'Dim z As Integer
    'Dim y As Integer = 0  '判斷是否有學員填寫問卷預設為 0 (沒有)
    'Dim SKID2 As String
    'Dim SQID2 As String
    'Dim objConn As SqlConnection
    'Dim Tb As HtmlTable
    'Dim Tb2 As HtmlTable
    'Dim Row As HtmlTableRow
    'Dim Cell As HtmlTableCell
    'Dim Label As Label
    'Dim btn As Button
    'Dim btn2 As HtmlButton
    'Dim btn3 As HtmlButton
    'Dim btn4 As HtmlButton
    'Dim RadioButtonList As RadioButtonList
    'Dim CheckBoxList As CheckBoxList
    'Dim sql2 As String
    'Dim sql3 As String
    'Dim sqlQ As String
    'Dim sqlA As String
    'Dim da As SqlDataAdapter = nothing
    'Dim dt2 As DataTable
    'Dim dt3 As DataTable
    'Dim Trans As SqlTransaction
    'Dim dr As DataRow
    'Dim dr2 As DataRow
    'Dim dr3 As DataRow
    'Dim sql As String
    'Dim dt As DataTable

    '預設測試用 (功能隱藏)
    Const cst_vsSurveyAnswer As String = "SurveyAnswer"
    Dim iGy As Integer = 0  '判斷是否有學員填寫問卷預設為 0 (沒有)
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

        If Convert.ToString(Request("SVID")) <> "" Then
            HID_SVID.Value = TIMS.ClearSQM(Request("SVID"))
        End If
        hidIptName.Value = TIMS.ClearSQM(Request("IptName"))

        iGy = 0
        Dim dt As DataTable = TIMS.Get_dtSdS(HID_SVID.Value, objconn)
        If dt.Rows.Count <> 0 Then
            iGy = 1  '判斷是否有學員填寫答案 1表示有
        End If
        If Me.Request("Type") = "V" Then
            RETop.Value = "關閉本頁"
            iGy = 2  '判斷是否顯示功能鈕 2表示外面呼叫
        End If

        If Not IsPostBack Then
            Me.ViewState(cst_vsSurveyAnswer) = Nothing
            Table_I.Visible = False '問題答案的輸入畫面
            Save_Q.Attributes("onclick") = "return CheckData();"
        End If

        Call ShowKeySurveyKind(HID_SVID.Value, PlaceHolder1, Me)

    End Sub

    '組合 Web 網頁動態方式伺服器控制項
    Sub ShowKeySurveyKind(ByVal RsSVID As String, ByRef PH1 As PlaceHolder, ByRef MyPage As Page)
        Dim CheckBoxList As CheckBoxList
        Dim RadioButtonList As RadioButtonList
        Dim Label As Label

        Dim SKID2 As String = ""
        Dim SQID2 As String = ""

        Dim sql As String = ""
        Dim sqlQ As String = ""
        Dim sqlA As String = ""

        Dim dt As DataTable
        Dim dt2 As DataTable
        Dim dt3 As DataTable

        Dim Tb As HtmlTable
        Dim Tb2 As HtmlTable

        Dim Row As HtmlTableRow
        Dim Cell As HtmlTableCell

        Dim btn As Button
        Dim btn3 As HtmlButton
        Dim btn4 As HtmlButton

        sql = "select SKID,Topic,Serial from Key_surveykind where SVID = '" & RsSVID & "' order by Serial "
        dt = DbAccess.GetDataTable(sql, objconn) '找出標題

        If dt.Rows.Count <> 0 Then

            For v As Integer = 0 To dt.Rows.Count - 1           '利用迴圈新增標題

                SKID2 = dt.Rows(v).Item(0).ToString   '取出 SKID

                Tb = New HtmlTable                    '新增TB
                PH1.Controls.Add(Tb)
                Tb.Attributes.Add("width", "100%")
                Tb.Attributes.Add("class", "font")

                'Row = New HtmlTableRow                 '新增TR
                'Tb.Controls.Add(Row)

                'Cell = New HtmlTableCell               '新增TD
                'Row.Controls.Add(Cell)
                'Cell.Attributes.Add("width", "100%")
                'Cell.Attributes.Add("align", "right")


                'btn = New Button                        '新增Button
                'Cell.Controls.Add(btn)
                'btn.Text = "新增問卷題目"
                'btn.ID = v
                'btn.CommandArgument = dt.Rows(v).Item(0).ToString  'SKID
                'btn.CommandName = "I"
                'If y = 1 Then     '表示已有填寫問卷答案
                '    btn.Enabled = False
                '    btn.ToolTip = "【問卷資料填寫】已有資料不能新增"
                'End If
                'AddHandler btn.Click, AddressOf on_btnclick       '如果btn click 就跑 on_btnclick 

                'If v = 0 Then      '如果是第一筆就新增回上一頁鍵
                '    btn2 = New HtmlButton
                '    btn2.ID = "RETop"
                '    btn2.InnerText = "回上一頁"
                '    btn2.Attributes.Add("runat", "server")
                '    Cell.Controls.Add(btn2)
                '    AddHandler btn2.ServerClick, AddressOf on_btn2click    '如果btn2 click 就跑 on_btn2click
                'End If


                Row = New HtmlTableRow                '新增TR
                Tb.Controls.Add(Row)

                Cell = New HtmlTableCell               '新增TD
                Row.Controls.Add(Cell)
                Cell.InnerText = dt.Rows(v).Item(1).ToString
                Cell.Attributes.Add("Style", "color@White;background-color: #2aafc0")
                Cell.Attributes.Add("width", "100%")
                Cell.Attributes.Add("class", "font")

                Cell = New HtmlTableCell               '新增TD
                Row.Controls.Add(Cell)
                Cell.Attributes.Add("width", "100%")
                Cell.Attributes.Add("align", "right")
                Cell.Attributes.Add("class", "font")

                btn = New Button                        '新增Button
                Cell.Controls.Add(btn)
                btn.Text = "新增問卷題目"
                btn.ID = "SK_" & v
                btn.CommandArgument = dt.Rows(v).Item(0).ToString  'SKID
                btn.CommandName = "I"
                Select Case iGy
                    Case 1 '表示已有填寫問卷答案
                        btn.Enabled = False
                        btn.ToolTip = "【問卷資料填寫】已有資料不能新增"
                    Case 2
                        btn.Enabled = False
                        btn.Visible = False
                    Case Else
                        AddHandler btn.Click, AddressOf on_btnclick       '如果btn click 就跑 on_btnclick 
                End Select

                Tb2 = New HtmlTable                      '新增TB

                PH1.Controls.Add(Tb2)
                Tb2.Attributes.Add("width", "100%")
                Tb2.Attributes.Add("CELLSPACING", 0)
                Tb2.Attributes.Add("cellPadding", 0)
                Tb2.Attributes.Add("BORDER", 0)
                Tb2.Attributes.Add("class", "font")

                Tb2.ID = "Tb2" & v

                sqlQ = "select * from ID_SurveyQuestion where SKID = " & SKID2 & " order by Serial " '找出問題題目
                dt2 = DbAccess.GetDataTable(sqlQ, objconn)
                If dt2.Rows.Count <> 0 Then

                    For k As Integer = 0 To dt2.Rows.Count - 1
                        SQID2 = dt2.Rows(k).Item(0).ToString            'SQID

                        Row = New HtmlTableRow                          '新增tr

                        MyPage.FindControl("Tb2" & v).Controls.Add(Row)

                        Row.Attributes.Add("Style", "background-color: #e9f2fc")
                        Row.Attributes.Add("width", "100%")
                        Row.Attributes.Add("class", "font")

                        Cell = New HtmlTableCell                        '新增td
                        Row.Controls.Add(Cell)

                        Label = New Label
                        Cell.Controls.Add(Label)
                        Cell.Attributes.Add("width", "90%")
                        Cell.Attributes.Add("class", "font")
                        Label.Text = dt2.Rows(k).Item(1).ToString        '題目內容

                        Cell = New HtmlTableCell                         '新增td
                        Row.Controls.Add(Cell)
                        Cell.Attributes.Add("align", "right")

                        btn3 = New HtmlButton                             '新增修改button
                        btn3.ID = "e" & SQID2     '修改鍵的id,SQID的value + e 字元,因為id 為SQID 己使用過,所以加 e判別,方使傳SQID
                        btn3.InnerText = "修改"
                        btn3.Attributes.Add("runat", "server")
                        Select Case iGy
                            Case 1
                                btn3.Disabled = True
                                TIMS.Tooltip(btn3, "【問卷資料填寫】已有資料不能修改")
                                Cell.Controls.Add(btn3)
                            Case 2
                                btn3.Disabled = True
                                btn3.Style("display") = "none"
                                'Cell.Controls.Add(btn3)
                            Case Else
                                Cell.Controls.Add(btn3)
                                AddHandler btn3.ServerClick, AddressOf on_btn5click
                        End Select

                        btn4 = New HtmlButton                              '新增刪除button
                        btn4.ID = "d" & SQID2
                        btn4.InnerText = "刪除"
                        btn4.Attributes.Add("runat", "server")
                        Select Case iGy
                            Case 1
                                btn4.Disabled = True
                                TIMS.Tooltip(btn4, "【問卷資料填寫】已有資料不能刪除")
                                Cell.Controls.Add(btn4)
                            Case 2
                                btn4.Disabled = True
                                btn4.Style("display") = "none"
                                'Cell.Controls.Add(btn4)
                            Case Else
                                Cell.Controls.Add(btn4)
                                AddHandler btn4.ServerClick, AddressOf on_btn4click
                        End Select

                        sqlA = "select * from ID_SurveyAnswer where SQID= " & SQID2 & " order by Serial"
                        dt3 = DbAccess.GetDataTable(sqlA, objconn)

                        If dt3.Rows.Count <> 0 Then

                            Row = New HtmlTableRow                           '新增TR
                            MyPage.FindControl("Tb2" & v).Controls.Add(Row)
                            Row.Attributes.Add("width", "100%")
                            Row.Attributes.Add("Style", "background-color: #e9f2fc")

                            Cell = New HtmlTableCell
                            Row.Controls.Add(Cell)                           '新增TD

                            If dt2.Rows(k).Item(2).ToString = 1 Then    '判斷是那程型態的題目 1是radio2,2是checkbox

                                RadioButtonList = New RadioButtonList
                                RadioButtonList.Attributes.Add("class", "font")
                                Cell.Attributes.Add("colspan", 2)
                                Cell.Controls.Add(RadioButtonList)

                                For z As Integer = 0 To dt3.Rows.Count - 1
                                    RadioButtonList.Items.Add(dt3.Rows(z).Item(1).ToString)
                                Next

                            Else                                          'checkbox

                                CheckBoxList = New CheckBoxList
                                CheckBoxList.Attributes.Add("class", "font")
                                Cell.Attributes.Add("colspan", 2)
                                Cell.Controls.Add(CheckBoxList)

                                For z As Integer = 0 To dt3.Rows.Count - 1
                                    CheckBoxList.Items.Add(dt3.Rows(z).Item(1).ToString)
                                Next

                            End If

                            Cell = New HtmlTableCell
                            Row.Controls.Add(Cell)
                        End If
                    Next

                End If
            Next
        End If
    End Sub


    Private Sub DataGrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Select Case e.Item.ItemType

            Case ListItemType.AlternatingItem, ListItemType.Item

                Dim drv As DataRowView = e.Item.DataItem
                Dim Label1 As Label = e.Item.FindControl("Label1")
                Dim LAnswer As Label = e.Item.FindControl("LAnswer")

                Dim Edit As Button = e.Item.FindControl("Edit")
                Dim Del As Button = e.Item.FindControl("Del")

                Label1.Text = drv("Serial")
                LAnswer.Text = drv("Answer")

                If Label1.Text <> "" Then
                    Dim iNext As Integer = Val(Label1.Text) + 1
                    If HidSerial1.Value <> "" Then
                        If iNext > Val(HidSerial1.Value) Then
                            HidSerial1.Value = iNext
                        End If
                    Else
                        HidSerial1.Value = iNext
                    End If
                End If
                If drv("SAID").ToString <> "" Then
                    Edit.CommandArgument = CInt(drv("SAID"))
                End If

                If drv("SAID").ToString <> "" Then
                    Del.CommandArgument = CInt(drv("SAID"))
                End If
                Del.Attributes("onclick") = TIMS.cst_confirm_delmsg1 '刪除


            Case ListItemType.EditItem

                Dim drv As DataRowView = e.Item.DataItem
                Dim ENO1 As TextBox = e.Item.FindControl("ENO1")
                Dim EAnswer As TextBox = e.Item.FindControl("EAnswer")
                Dim Save As Button = e.Item.FindControl("Save")
                Dim Cancel As Button = e.Item.FindControl("Cancel")

                ENO1.Text = drv("Serial")
                EAnswer.Text = drv("Answer")
                If ENO1.Text <> "" Then
                    Dim iNext As Integer = Val(ENO1.Text) + 1
                    If HidSerial1.Value <> "" Then
                        If iNext > Val(HidSerial1.Value) Then
                            HidSerial1.Value = iNext
                        End If
                    Else
                        HidSerial1.Value = iNext
                    End If
                End If

                If drv("SAID").ToString <> "" Then
                    Save.CommandArgument = CInt(drv("SAID"))
                End If

                If drv("SAID").ToString <> "" Then
                    Cancel.CommandArgument = CInt(drv("SAID"))
                End If

                Save.Attributes("onclick") = "return CheckDescDataE('" & EAnswer.ClientID & "','" & ENO1.ClientID & "');" '存檔


            Case ListItemType.Footer

                Dim FNO1 As TextBox = e.Item.FindControl("FNO1")
                Dim FAnswer As TextBox = e.Item.FindControl("FAnswer")
                Dim ddlFAnswer As DropDownList = e.Item.FindControl("ddlFAnswer")
                Dim Save2 As Button = e.Item.FindControl("Save2")
                FNO1.Text = HidSerial1.Value
                ddlFAnswer = GetAnswer(ddlFAnswer, HID_SVID.Value, SKID.Value)
                ddlFAnswer.Attributes("onclick") = "RtnFAnswer('" & FNO1.ClientID & "','" & ddlFAnswer.ClientID & "');"
                'ddlFAnswer.Style("display") = "none" '預設測試用
                'ddlFAnswer.Attributes("onchange") = "RtnFAnswer('" & FNO1.ClientID & "','" & ddlFAnswer.ClientID & "');"
                Save2.Attributes("onclick") = "return CheckDescData('" & FAnswer.ClientID & "','" & FNO1.ClientID & "','" & ddlFAnswer.ClientID & "');" '新增

        End Select

    End Sub

    Function GetAnswer(ByVal obj As ListControl, ByVal SVID As String, Optional ByVal sKID As String = "0") As ListControl
        Dim dt As DataTable
        Dim sql As String
        sql = "" & vbCrLf
        sql += " SELECT  distinct  " & vbCrLf
        sql += " 	sa.Serial " & vbCrLf
        sql += " 	, sa.Answer" & vbCrLf
        sql += " 	, sa.Serial + '|' + sa.Answer as saValue" & vbCrLf
        sql += " from id_SurveyAnswer sa " & vbCrLf
        sql += "  join ID_SurveyQuestion sq on sq.SQID=sa.SQID" & vbCrLf
        sql += "  WHERE sq.SVID='" & SVID & "'" & vbCrLf
        If sKID <> "0" Then
            sql += "  AND sq.SKid='" & sKID & "'" & vbCrLf
        End If
        sql += " order by  sa.Serial " & vbCrLf
        sql += " " & vbCrLf

        dt = DbAccess.GetDataTable(sql, objconn)
        With obj
            .DataSource = dt
            .DataTextField = "Answer"
            .DataValueField = "saValue"
            .DataBind()
            .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        End With
        Return obj
    End Function

    Private Sub DataGrid2_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid2.ItemCommand
        Dim sql As String
        Dim dr As DataRow
        Dim dt As DataTable

        Select Case e.CommandName

            Case "Edit"    '修改

                DataGrid2.EditItemIndex = e.Item.ItemIndex

            Case "Del"      '刪除

                dt = Me.ViewState(cst_vsSurveyAnswer)
                If dt.Select("SAID='" & e.CommandArgument & "'").Length <> 0 Then
                    dt.Select("SAID='" & e.CommandArgument & "'")(0).Delete()
                End If
                DataGrid2.EditItemIndex = -1

            Case "update"  '修改存檔

                dt = Me.ViewState(cst_vsSurveyAnswer)
                Dim ENO1 As TextBox = e.Item.FindControl("ENO1")
                Dim EAnswer As TextBox = e.Item.FindControl("EAnswer")
                If dt.Select("SAID='" & e.CommandArgument & "'").Length <> 0 Then
                    dr = dt.Select("SAID='" & e.CommandArgument & "'")(0)
                    dr("Answer") = EAnswer.Text
                    dr("Serial") = ENO1.Text
                    dr("ModifyAcct") = sm.UserInfo.UserID
                    dr("ModifyDate") = Now
                End If
                DataGrid2.EditItemIndex = -1

            Case "Cancel"  '取消


                DataGrid2.EditItemIndex = -1

            Case "Save2"   '新增
                Dim FNO1 As TextBox = e.Item.FindControl("FNO1")
                Dim FAnswer As TextBox = e.Item.FindControl("FAnswer")
                Dim ddlFAnswer As DropDownList = e.Item.FindControl("ddlFAnswer")

                If Me.ViewState(cst_vsSurveyAnswer) Is Nothing Then '如果是第一筆

                    sql = "Select * From ID_SurveyAnswer Where 1<>1 "
                    dt = DbAccess.GetDataTable(sql, objconn)
                    dt.Columns("SAID").AutoIncrement = True
                    dt.Columns("SAID").AutoIncrementSeed = -1
                    dt.Columns("SAID").AutoIncrementStep = -1
                Else
                    dt = Me.ViewState(cst_vsSurveyAnswer)
                End If

                dr = dt.NewRow
                dt.Rows.Add(dr)
                If FAnswer.Text.ToString <> "" Then
                    dr("Answer") = FAnswer.Text
                Else
                    dr("Answer") = ddlFAnswer.SelectedItem.Text
                End If

                If FNO1.Text.ToString <> "" Then
                    dr("Serial") = FNO1.Text
                End If
                dr("ModifyAcct") = sm.UserInfo.UserID
                dr("ModifyDate") = Now
                Me.ViewState(cst_vsSurveyAnswer) = dt

        End Select

        dt = Me.ViewState(cst_vsSurveyAnswer)

        Me.DataGrid2.DataSource = dt

        If dt.Rows.Count <> 0 Then   '檢查是否有填答案

            Answercount.Value = "Y"
        Else
            Answercount.Value = "N"
        End If

        DataGrid2.Visible = True
        Me.DataGrid2.DataBind()

    End Sub

    'save
    Private Sub Save_Q_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Save_Q.Click
        If Convert.ToString(Request("SVID")) <> "" Then
            HID_SVID.Value = TIMS.ClearSQM(Request("SVID"))
        End If

        Dim sql As String = ""
        '新增問卷題目
        sql = ""
        sql &= " Insert Into ID_SURVEYQUESTION(SQID,Question,QType,Serial,SVID,SKID,ModifyAcct,ModifyDate)"
        sql &= " VALUES (@SQID,@Question,@QType,@Serial,@SVID,@SKID,@ModifyAcct,getdate())"
        Dim iCmd As New SqlCommand(sql, objconn)

        sql = ""
        sql &= " UPDATE ID_SURVEYQUESTION"
        sql &= " SET Question=@Question"
        sql &= " ,QType=@QType"
        sql &= " ,Serial=@Serial"
        sql &= " ,SVID=@SVID"
        sql &= " ,SKID=@SKID"
        sql &= " ,ModifyAcct=@ModifyAcct"
        sql &= " ,ModifyDate=getdate()"
        sql &= " WHERE 1=1 "
        sql &= " AND SQID=@SQID"
        Dim uCmd As New SqlCommand(sql, objconn)

        '找出這次新增加的SQID ,問卷題目的KEY
        'sql = "SELECT SQID FROM ID_SURVEYQUESTION WHERE SKID =@SKID ORDER BY SQID DESC"
        'Dim sCmd As New SqlCommand(sql, objconn)
        'dr2 = DbAccess.GetOneRow(sql2, objconn)

        Call TIMS.OpenDbConn(objconn)
        Select Case Type.Value
            Case "I" '如果是新增
                Dim iSQID As Integer = 0
                iSQID = DbAccess.GetNewId(objconn, "ID_SURVEYQUESTION_SQID_SEQ,ID_SURVEYQUESTION,SQID")
                With iCmd
                    .Parameters.Clear()
                    .Parameters.Add("SQID", SqlDbType.Int).Value = iSQID
                    .Parameters.Add("Question", SqlDbType.NVarChar).Value = Question.Text
                    .Parameters.Add("QType", SqlDbType.VarChar).Value = QTYPE.SelectedValue
                    .Parameters.Add("Serial", SqlDbType.Int).Value = CInt(SerialQ.Text)
                    .Parameters.Add("SVID", SqlDbType.Int).Value = Val(HID_SVID.Value)
                    .Parameters.Add("SKID", SqlDbType.Int).Value = Val(SKID.Value)
                    .Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                    .ExecuteNonQuery()
                End With
                'For Each in DataGrid2.

                SQID.Value = iSQID 'dr2("SQID")
                If Not Me.ViewState(cst_vsSurveyAnswer) Is Nothing Then '假如暫存的問卷答案TABLE裡有資料
                    Dim dtTemp2 As DataTable
                    Dim da As SqlDataAdapter = Nothing

                    dtTemp2 = Me.ViewState(cst_vsSurveyAnswer)

                    sql = "select * from ID_SURVEYANSWER where SQID = " & iSQID & ""
                    Dim dt As DataTable = DbAccess.GetDataTable(sql, da, objconn)
                    For Each dr As DataRow In dtTemp2.Select(Nothing, Nothing, DataViewRowState.CurrentRows)
                        dr("SQID") = iSQID 'dr2("SQID")   '問卷答案的SQID = 問卷題目的SQID
                        Dim iSAID As Integer = 0
                        If Convert.ToString(dr("SAID")) <> "" Then iSAID = Val(dr("SAID"))
                        If iSAID <= 0 Then iSAID = DbAccess.GetNewId(objconn, "ID_SURVEYANSWER_SAID_SEQ,ID_SURVEYANSWER,SAID")
                        dr("SAID") = iSAID
                    Next
                    dt = dtTemp2.Copy   ' copy dt
                    DbAccess.UpdateDataTable(dt, da)
                    Me.ViewState(cst_vsSurveyAnswer) = Nothing
                End If
                Common.MessageBox(Me, "新增成功")
            Case "E"
                If SQID.Value = "" Then
                    Common.MessageBox(Me, "查無修改資料!!")
                    Exit Sub
                End If

                With uCmd
                    .Parameters.Clear()
                    .Parameters.Add("Question", SqlDbType.NVarChar).Value = Question.Text
                    .Parameters.Add("QType", SqlDbType.VarChar).Value = QTYPE.SelectedValue
                    .Parameters.Add("Serial", SqlDbType.Int).Value = CInt(SerialQ.Text)
                    .Parameters.Add("SVID", SqlDbType.Int).Value = Val(HID_SVID.Value)
                    .Parameters.Add("SKID", SqlDbType.Int).Value = Val(SKID.Value)
                    .Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                    .Parameters.Add("SQID", SqlDbType.Int).Value = Val(SQID.Value)
                    .ExecuteNonQuery()
                End With

                If Not Me.ViewState(cst_vsSurveyAnswer) Is Nothing Then '假如暫存的問卷答案TABLE裡有資料
                    Dim dtTemp2 As DataTable
                    Dim da As SqlDataAdapter = Nothing
                    dtTemp2 = Me.ViewState(cst_vsSurveyAnswer)
                    sql = "select * from ID_SURVEYANSWER where SQID = " & Val(SQID.Value) & ""
                    Dim dt As DataTable = DbAccess.GetDataTable(sql, da, objconn)
                    For Each dr As DataRow In dtTemp2.Select(Nothing, Nothing, DataViewRowState.CurrentRows)
                        dr("SQID") = Val(SQID.Value) 'dr2("SQID")   '問卷答案的SQID = 問卷題目的SQID
                        Dim iSAID As Integer = 0
                        If Convert.ToString(dr("SAID")) <> "" Then iSAID = Val(dr("SAID"))
                        If iSAID <= 0 Then iSAID = DbAccess.GetNewId(objconn, "ID_SURVEYANSWER_SAID_SEQ,ID_SURVEYANSWER,SAID")
                        dr("SAID") = iSAID
                    Next
                    dt = dtTemp2.Copy   ' copy dt
                    DbAccess.UpdateDataTable(dt, da)
                    Me.ViewState(cst_vsSurveyAnswer) = Nothing
                End If
                Common.MessageBox(Me, "修改成功")
        End Select

        PlaceHolder1.Controls.Clear()
        Call ShowKeySurveyKind(HID_SVID.Value, PlaceHolder1, Me)
        'Page_Load(sender, e)

        PlaceHolder1.Visible = True
        RETop.Visible = True
        Table_I.Visible = False

    End Sub

#Region "FUN"
    '(新增)動態按鈕。
    Sub on_btnclick(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim SqlQ As String
        Dim Sql As String
        'Dim Sql2 As String
        Dim dt As DataTable
        Dim dr As DataRow

        Dim ThisBtn As Button = CType(sender, Button) '取出是那個BUTTON INSERT

        Type.Value = ThisBtn.CommandName   'INSERT TYPE = I
        PlaceHolder1.Visible = False
        RETop.Visible = False

        Answercount.Value = "N"  '判斷是否有輸入問題的答案項 預設為N
        Ivalue.Value = ThisBtn.ID '取得所按button的id
        SKID.Value = ThisBtn.CommandArgument '取出SKID

        SqlQ = "Select * from Key_SurveyKind where SKID = " & SKID.Value & ""
        dr = DbAccess.GetOneRow(SqlQ, objconn) '為了取出標題

        Sql = "Select * from ID_SurveyAnswer Where 1<>1"
        dt = DbAccess.GetDataTable(Sql, objconn)

        QLabel.InnerText = dr("Topic") '取出標題
        Question.Text = ""   '問題內容
        QTYPE.SelectedIndex = 0 '問題類型
        SerialQ.Text = "" '問題排序
        Table_I.Visible = True '新增畫面

        Me.DataGrid2.DataSource = dt
        DataGrid2.Visible = True
        Me.DataGrid2.DataBind()
    End Sub

    '(刪除)動態按鈕。
    Sub on_btn4click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim sql As String
        Dim SQID_d As String
        Dim dt As DataTable
        'Dim TYPE2 As String
        'Dim dr As DataRow
        'Dim SqlQ As String

        Dim ThisBtn As HtmlButton = CType(sender, HtmlButton) '取得是那個BUTTON 按刪除

        SQID_d = Replace(ThisBtn.ID, "d", "") ' ThisBtn.ID 之前設定 = SQID的值
        'SQID_d = ThisBtn.ID ' ThisBtn.ID 之前設定 = SQID的值

        sql = "select * from ID_SurveyQuestion IQ left join ID_SurveyAnswer IA on IQ.SQID = IA.SQID where IQ.SQID = '" & SQID_d & " '"
        dt = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count <> 0 Then

            sql = "Delete ID_SurveyQuestion where SQID = '" & SQID_d & " '"   '刪除問卷題目
            DbAccess.ExecuteNonQuery(sql, objconn)
            sql = "Delete ID_SurveyAnswer where SQID = '" & SQID_d & " '"     '刪除問卷題案項
            DbAccess.ExecuteNonQuery(sql, objconn)
            Common.MessageBox(Me, "刪除成功")

        End If

        PlaceHolder1.Controls.Clear()
        Call ShowKeySurveyKind(HID_SVID.Value, PlaceHolder1, Me)
        'Page_Load(sender, e)


    End Sub

    '(修改)動態按鈕。
    Sub on_btn5click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim sql As String
        'Dim SQID_d As String
        Dim dt As DataTable
        Dim dr As DataRow
        Dim SqlQ As String

        Dim ThisBtn As HtmlButton = CType(sender, HtmlButton) '取得是那個BUTTON 按修改

        Type.Value = "E" '修改

        PlaceHolder1.Visible = False   '問卷畫面隱藏
        RETop.Visible = False

        SQID.Value = Replace(ThisBtn.ID, "e", "") 'ThisBtn.ID = SQID.value + "e" 所以是取ThisBtn.ID左邊的ThisBtn.ID的長度-1,取SQID.value
        'SQID.Value = Left(ThisBtn.ID, Len(ThisBtn.ID) - 1)  'ThisBtn.ID = SQID.value + "e" 所以是取ThisBtn.ID左邊的ThisBtn.ID的長度-1,取SQID.value

        SqlQ = "Select KS.Topic,IQ.Question,IQ.Qtype,IQ.Serial from Key_SurveyKind KS Join ID_SurveyQuestion IQ on IQ.SKID = KS.SKID where IQ.SQID = " & SQID.Value & ""
        dr = DbAccess.GetOneRow(SqlQ, objconn)
        sql = "Select * from ID_SurveyAnswer Where SQID = '" & SQID.Value.ToString & "' order by Serial "
        dt = DbAccess.GetDataTable(sql, objconn)

        QLabel.InnerText = dr.Item(0).ToString '標題
        Question.Text = dr.Item(1).ToString    '問題
        QTYPE.SelectedValue = dr.Item(2).ToString '問題類型
        SerialQ.Text = dr.Item(3).ToString '序號
        Table_I.Visible = True

        Me.DataGrid2.DataSource = dt
        Me.ViewState(cst_vsSurveyAnswer) = dt
        DataGrid2.Visible = True
        Me.DataGrid2.DataBind()

        PlaceHolder1.Controls.Clear()
        Call ShowKeySurveyKind(HID_SVID.Value, PlaceHolder1, Me)
        'Page_Load(sender, e)

    End Sub


#End Region

    Private Sub return1_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles return1.ServerClick '修改或新增畫面的回上一頁
        Me.ViewState(cst_vsSurveyAnswer) = Nothing
        Table_I.Visible = False

        PlaceHolder1.Controls.Clear()
        Call ShowKeySurveyKind(HID_SVID.Value, PlaceHolder1, Me)
        'Page_Load(sender, e)

        PlaceHolder1.Visible = True
        RETop.Visible = True

    End Sub

    Private Sub RETop_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RETop.ServerClick
        Select Case iGy
            Case 2
                Common.RespWrite(Me, "<script>window.close();</script>")
            Case Else
                Dim sUrl As String = "SV_01_003.aspx?ID=" & Request("ID") & "&IptName=" & hidIptName.Value & ""
                TIMS.Utl_Redirect1(Me, sUrl)
        End Select


    End Sub

End Class
