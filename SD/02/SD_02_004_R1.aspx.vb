Partial Class SD_02_004_R1
    Inherits AuthBasePage

    'Maintest_list.jrxml
    Const cst_reportFN1 As String = "Maintest_list2"

    Dim rqPlanid As String = "" 'TIMS.sUtl_GetRqValue(Me, "planid")
    Dim rqOCID As String = "" 'TIMS.sUtl_GetRqValue(Me, "OCID")
    Dim rqMailtypeR1 As String = "" 'TIMS.sUtl_GetRqValue(Me, "Mailtype1")
    Dim rqMailtypeR2 As String = "" 'TIMS.sUtl_GetRqValue(Me, "Mailtype2")
    Dim rqMailtypeR3 As String = "" 'TIMS.sUtl_GetRqValue(Me, "Mailtype3")
    Dim rqMailtypeR4 As String = "" 'TIMS.sUtl_GetRqValue(Me, "Mailtype4")
    Dim rqMailtypeR5 As String = "" 'TIMS.sUtl_GetRqValue(Me, "Mailtype5")
    Dim rqChk1 As String = "" 'TIMS.sUtl_GetRqValue(Me, "chk1")
    Dim rqChk2 As String = "" 'TIMS.sUtl_GetRqValue(Me, "chk2")
    Dim rqChk3 As String = "" 'TIMS.sUtl_GetRqValue(Me, "chk3")
    Dim rqChk4 As String = "" 'TIMS.sUtl_GetRqValue(Me, "chk4")
    Dim rqChk5 As String = "" 'TIMS.sUtl_GetRqValue(Me, "chk5")
    Dim rqDistID As String = "" '2018-09-07 add 計畫所屬轄區代碼
    Dim gERRMSG1 As String = ""

    'Dim MailtypeR As Integer
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        '分頁設定 Start
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub

        rqPlanid = TIMS.sUtl_GetRqValue(Me, "planid")
        rqOCID = TIMS.sUtl_GetRqValue(Me, "OCID")
        rqMailtypeR1 = TIMS.sUtl_GetRqValue(Me, "Mailtype1")
        rqMailtypeR2 = TIMS.sUtl_GetRqValue(Me, "Mailtype2")
        rqMailtypeR3 = TIMS.sUtl_GetRqValue(Me, "Mailtype3")
        rqMailtypeR4 = TIMS.sUtl_GetRqValue(Me, "Mailtype4")
        rqMailtypeR5 = TIMS.sUtl_GetRqValue(Me, "Mailtype5")
        rqChk1 = TIMS.sUtl_GetRqValue(Me, "chk1")
        rqChk2 = TIMS.sUtl_GetRqValue(Me, "chk2")
        rqChk3 = TIMS.sUtl_GetRqValue(Me, "chk3")
        rqChk4 = TIMS.sUtl_GetRqValue(Me, "chk4")
        rqChk5 = TIMS.sUtl_GetRqValue(Me, "chk5")
        rqDistID = TIMS.sUtl_GetRqValue(Me, "distid")

        'request("OCID")
        'request("SEDIT"
        If Not IsPostBack Then
            Mailtype.Value = ""
            If rqMailtypeR1 <> "" Then Mailtype.Value += "&Mailtype1=" & rqMailtypeR1
            If rqMailtypeR2 <> "" Then Mailtype.Value += "&Mailtype2=" & rqMailtypeR2
            If rqMailtypeR3 <> "" Then Mailtype.Value += "&Mailtype3=" & rqMailtypeR3
            If rqMailtypeR4 <> "" Then Mailtype.Value += "&Mailtype4=" & rqMailtypeR4
            If rqMailtypeR5 <> "" Then Mailtype.Value += "&Mailtype5=" & rqMailtypeR5

            chkvalue.Value = ""
            If rqChk1 <> "" Then chkvalue.Value += "&chk1=" & rqChk1
            If rqChk2 <> "" Then chkvalue.Value += "&chk2=" & rqChk2
            If rqChk3 <> "" Then chkvalue.Value += "&chk3=" & rqChk3
            If rqChk4 <> "" Then chkvalue.Value += "&chk4=" & rqChk4
            If rqChk5 <> "" Then chkvalue.Value += "&chk5=" & rqChk5
            Call CreateTable(rqOCID)
        End If

        'Me.DistID.Value = sm.UserInfo.DistID
        Me.DistID.Value = rqDistID '2018-09-07 改成以該班所屬計畫轄區代碼
        Me.OCID.Value = rqOCID
        Me.OCID.Value = TIMS.ClearSQM(Me.OCID.Value)
        'Me.MailtypeR = Request("Mailtype1")
        'Button2.Attributes("onclick") = "history.go(-1);return false;"
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e.CommandArgument = "" Then Exit Sub
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        If Me.OCID.Value = "" Then
            Common.MessageBox(Me, "班級代碼有誤,請重新查詢.")
            Exit Sub
        End If
        If Me.OCID.Value <> rqOCID Then
            Common.MessageBox(Me, "班級代碼有誤,請重新查詢.")
            Exit Sub
        End If
        Dim sCmdArg As String = e.CommandArgument
        Dim cmSETID As String = TIMS.GetMyValue(sCmdArg, "SETID")
        Dim cmEnterDate As String = TIMS.GetMyValue(sCmdArg, "EnterDate")
        Dim cmSerNum As String = TIMS.GetMyValue(sCmdArg, "SerNum")
        If cmSETID = "" OrElse cmEnterDate = "" OrElse cmSerNum = "" Then Exit Sub

        Select Case e.CommandName
            Case "update"
                If sCmdArg <> "" Then
                    source.Visible = False
                    gERRMSG1 = ""
                    Dim NewExamNO As String = UpdateExamNo(sCmdArg) 'NewExamNO = UpdateExamNo(e.CommandArgument)
                    If gERRMSG1 <> "" Then
                        Common.MessageBox(Me, gERRMSG1)
                        Return
                    End If

                    If NewExamNO <> "" Then
                        Call CreateTable(rqOCID)
                        Dim js_str1 As String = Common.GetJsString(String.Concat("准考証更新完成!![", NewExamNO, "]."))
                        Dim strScript As String = String.Concat("<script language=""javascript""> alert('", js_str1, "');</script>")
                        Page.RegisterStartupScript("", strScript)
                    End If
                End If
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                'Dim MailtypeR1 As String
                'Dim MailtypeR2 As String
                'Dim MailtypeR3 As String
                'Dim MailtypeR4 As String
                'Dim MailtypeR5 As String
                Dim drv As DataRowView = e.Item.DataItem
                Dim ExamNO As HtmlInputHidden = e.Item.FindControl("ExamNO")
                Dim DoubleExamNo As Label = e.Item.FindControl("DoubleExamNo")
                Dim UpdateBtn As LinkButton = e.Item.FindControl("UpdateBtn")
                Dim Mailtype3 As CheckBoxList = e.Item.FindControl("Mailtype3") '郵寄型態
                Dim Button1 As LinkButton = e.Item.FindControl("Button1") '列印

                Mailtype3.Items(0).Selected = False
                Mailtype3.Items(1).Selected = False
                Mailtype3.Items(2).Selected = False
                Mailtype3.Items(3).Selected = False
                Mailtype3.Items(4).Selected = False
                If rqMailtypeR1 = "1" Then Mailtype3.Items(0).Selected = True
                If rqMailtypeR2 = "1" Then Mailtype3.Items(1).Selected = True
                If rqMailtypeR3 = "1" Then Mailtype3.Items(2).Selected = True
                If rqMailtypeR4 = "1" Then Mailtype3.Items(3).Selected = True
                If rqMailtypeR5 = "1" Then Mailtype3.Items(4).Selected = True

                Mailtype3.Enabled = False
                TIMS.Tooltip(Mailtype3, "在上個動作，請先選好!", True)
                ExamNO.Value = "'" & drv("ExamNO") & "'"
                Button1.CommandArgument = "?SETID='" & DataGrid1.DataKeys(e.Item.ItemIndex) & "'"

                Dim MyValue As String = "" '"ExamNO=" & drv("ExamNO") & "&planid=" & Me.Request("planid") & "&DistID=" & sm.UserInfo.DistID & "&OCID1=" & Request("OCID") & 
                '"&Mailtype1=" & MailtypeR1 & "&Mailtype2=" & MailtypeR2 & "&Mailtype3=" & MailtypeR3 & "&Mailtype4=" & MailtypeR4 & "&Mailtype5=" & MailtypeR5 & chkvalue.Value
                MyValue = "ExamNO=" & drv("ExamNO")
                MyValue &= "&planid=" & rqPlanid
                'MyValue &= "&DistID=" & sm.UserInfo.DistID
                MyValue &= "&DistID=" & Convert.ToString(drv("distid")) '2018-09-07 改傳計畫所屬轄區代碼，署的使用者才有得列印
                MyValue &= "&OCID1=" & rqOCID
                MyValue &= "&Mailtype1=" & rqMailtypeR1
                MyValue &= "&Mailtype2=" & rqMailtypeR2
                MyValue &= "&Mailtype3=" & rqMailtypeR3
                MyValue &= "&Mailtype4=" & rqMailtypeR4
                MyValue &= "&Mailtype5=" & rqMailtypeR5
                MyValue &= chkvalue.Value

                Button1.Attributes("onclick") = ReportQuery.ReportScript(Me, cst_reportFN1, MyValue)
                If e.Item.Cells(1).Text = "&nbsp;" Then Button1.Enabled = False
                TIMS.Tooltip(UpdateBtn, "因TIMS屬多人操作系統，故同一時間新增，會有此情況發生")
                TIMS.Tooltip(DoubleExamNo, "因TIMS屬多人操作系統，故同一時間新增，會有此情況發生")
                'UpdateBtn.Attributes("onclick") = "disable_btn(this);"
                '檢查有重複的准考證號
                Dim flag_dbl1 As Boolean = TIMS.CheckDblExamNo(drv("OCID1").ToString, drv("ExamNO").ToString, objconn)

                Button1.Enabled = True
                UpdateBtn.Visible = False
                If flag_dbl1 Then
                    Button1.Enabled = False 'True
                    UpdateBtn.Visible = True 'False
                    Dim sCmdArg As String = ""
                    TIMS.SetMyValue(sCmdArg, "SETID", "" & drv("SETID"))
                    TIMS.SetMyValue(sCmdArg, "EnterDate", "" & drv("EnterDate"))
                    TIMS.SetMyValue(sCmdArg, "SerNum", "" & drv("SerNum"))
                    'arg += "SETID=" & drv("SETID")
                    'arg += "&EnterDate=" & drv("EnterDate")
                    'arg += "&SerNum=" & drv("SerNum")
                    UpdateBtn.Visible = True
                    UpdateBtn.CommandArgument = sCmdArg 'arg
                    Button1.Enabled = False
                    TIMS.Tooltip(Button1, "重複無法列印")
                    'DoubleExamNo.Text = "有重複情況"
                End If
        End Select
    End Sub

    Sub CreateTable(ByVal OCID As String)
        If OCID = "" Then Exit Sub

        Dim drCC As DataRow = TIMS.GetOCIDDate(OCID, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "班級代碼有誤,請重新查詢.")
            Exit Sub
        End If
        label1.Text = Convert.ToString(drCC("CLASSCNAME2"))

        Dim parasm As New Hashtable() From {{"OCID1", OCID}}
        Dim sql As String = ""
        sql &= " SELECT a.Name" & vbCrLf
        sql &= " ,a.SETID" & vbCrLf
        sql &= " ,b.OCID1" & vbCrLf
        sql &= " ,CONVERT(varchar, b.EnterDate, 111) EnterDate" & vbCrLf
        sql &= " ,b.SerNum" & vbCrLf
        sql &= " ,b.eSerNum" & vbCrLf
        sql &= " ,b.ExamNo" & vbCrLf
        sql &= " ,b.RelEnterDate" & vbCrLf
        sql &= " ,b.EnterPath" & vbCrLf
        sql &= " ,d.DISTID" & vbCrLf '2018-09-07 add 查轄區代碼
        sql &= " FROM Stud_EnterTemp a" & vbCrLf
        sql &= " JOIN Stud_EnterType b ON a.SETID = b.SETID" & vbCrLf
        sql &= " JOIN class_classinfo c ON b.ocid1 = c.ocid" & vbCrLf
        sql &= " JOIN id_plan d ON c.planid = d.planid" & vbCrLf
        sql &= " WHERE b.OCID1 =@OCID1" & vbCrLf
        If Request("EnterPathY") = "W" Then sql &= " AND ISNULL(b.EnterPath,' ') = 'W'" & vbCrLf '★
        If Request("EnterPathN") = "W" Then sql &= " AND ISNULL(b.EnterPath,' ') != 'W'" & vbCrLf '★
        sql &= " ORDER BY b.ExamNo,b.RelEnterDate" & vbCrLf '2008/9/30 南區反應改成以准考證排序為主
        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn, parasm)

        DataGrid1.Visible = False
        PageControler1.Visible = False
        If dt.Rows.Count > 0 Then
            DataGrid1.Visible = True
            PageControler1.Visible = True
            'DataGrid1.Visible = True
            DataGrid1.DataSource = dt
            DataGrid1.DataKeyField = "SETID"
            DataGrid1.DataBind()
            'PageControler1.Visible = True
            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
        End If

        '檢查有重複的准考證號
        Dim flag_dbl1 As Boolean = TIMS.CheckDblExamNo(OCID, "", objconn)

        button3.Enabled = True '群組列印
        Me.button3.Attributes.Add("OnClick", "return CheckPrint();")
        Button4.Visible = False '自動修正准考試號
        If flag_dbl1 Then
            Me.button3.Attributes.Remove("OnClick")
            button3.Enabled = False
            TIMS.Tooltip(button3, "准考證號碼有誤或重複，無法列印資料，請先行修正")
            Button4.Visible = True
        End If
    End Sub

    '回上頁
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button2.Click
        Dim url As String = TIMS.GetFunIDUrl(Request("ID"), 0, objconn)
        Call TIMS.Utl_Redirect(Me, objconn, url)
        'Response.Redirect(url & "?ID=" & Request("ID"))
    End Sub

    '自動修正准考試號
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        OCID.Value = TIMS.ClearSQM(OCID.Value)
        If Me.OCID.Value = "" Then
            Common.MessageBox(Me, "班級代碼有誤,請重新查詢.")
            Exit Sub
        End If
        Dim dr As DataRow = TIMS.GetOCIDDate(Me.OCID.Value)
        If dr Is Nothing Then
            Common.MessageBox(Me, "班級代碼有誤,請重新查詢.")
            Exit Sub
        End If

        'Dim url As String = TIMS.GetFunIDUrl(Request("ID"), 1)

        Dim sPMS As New Hashtable() From {{"OCID1", OCID.Value}}
        Dim sql As String = ""
        sql &= " SELECT a.Name" & vbCrLf
        sql &= " ,a.SETID" & vbCrLf
        sql &= " ,b.ExamNo" & vbCrLf
        sql &= " ,b.RelEnterDate" & vbCrLf
        sql &= " ,b.OCID1" & vbCrLf
        sql &= " ,CONVERT(varchar, b.EnterDate, 111) EnterDate" & vbCrLf
        sql &= " ,b.SerNum" & vbCrLf
        sql &= " ,b.eSerNum" & vbCrLf
        sql &= " FROM Stud_EnterTemp a" & vbCrLf
        sql &= " JOIN Stud_EnterType b ON a.SETID = b.SETID" & vbCrLf
        sql &= " WHERE b.OCID1 =@OCID1" & vbCrLf
        sql &= " ORDER BY b.RelEnterDate" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, sPMS)

        If dt.Rows.Count > 0 Then
            For Each drv As DataRow In dt.Rows
                '檢查有重複的准考證號
                Dim flag_dbl1 As Boolean = TIMS.CheckDblExamNo(drv("OCID1").ToString, drv("ExamNO").ToString, objconn)
                If flag_dbl1 Then
                    Dim sArg As String = String.Concat("SETID=", drv("SETID"), "&EnterDate=", drv("EnterDate"), "&SerNum=", drv("SerNum"))
                    '自動更新為新准考證號
                    gERRMSG1 = ""
                    Dim ExamNo As String = UpdateExamNo(sArg)
                    If gERRMSG1 <> "" Then
                        Common.MessageBox(Me, gERRMSG1)
                        Return
                    End If
                End If
            Next
        End If

        Call CreateTable(rqOCID)

        Dim strScript As String = ""
        strScript &= "<script language=""javascript"">" + vbCrLf
        strScript += " alert('准考証更新完成!!');" + vbCrLf
        strScript += "</script>" + vbCrLf
        Page.RegisterStartupScript("", strScript)
    End Sub

    '自動更新為新准考證號
    Function UpdateExamNo(ByVal sCmdArg As String) As String
        Dim rstExamNo As String = ""
        If sCmdArg <> "" Then
            'Dim da As New SqlDataAdapter
            Dim oCmd As SqlCommand = Nothing

            Dim SETID As String = TIMS.GetMyValue(sCmdArg, "SETID")
            Dim EnterDate As String = TIMS.Cdate3(TIMS.GetMyValue(sCmdArg, "EnterDate"))
            Dim SerNum As String = TIMS.GetMyValue(sCmdArg, "SerNum")

            Call TIMS.OpenDbConn(objconn)
            Dim tPMS As New Hashtable() From {{"SETID", SETID}, {"EnterDate", EnterDate}, {"SerNum", SerNum}}
            Dim sql As String = ""
            sql &= " SELECT * FROM STUD_ENTERTYPE" & vbCrLf
            sql &= " WHERE SETID =@SETID AND EnterDate=@EnterDate AND SerNum =@SerNum" & vbCrLf
            Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, tPMS)
            If dt.Rows.Count > 0 Then
                Dim dr As DataRow = dt.Rows(0)
                Dim OCIDValue1 As String = dr("OCID1").ToString
                Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1, objconn)
                Dim ExamPlanID As String = Convert.ToString(drCC("PlanID"))
                Dim eSerNum As String = dr("eSerNum").ToString
                '取出准考證號   Start
                'Dim ExamNo1 As String = "" '取出班級的CLASSID +期別 成為准考證編碼的前面的固定碼
                Dim ExamOcid1 As String = OCIDValue1
                Dim ExamNo1 As String = TIMS.Get_ExamNo1(ExamOcid1, objconn)
                If ExamNo1 = "" OrElse ExamNo1.Length < 6 Then '防呆
                    gERRMSG1 = String.Concat("准考証更新有誤![班級的代號與期別有誤].")
                    Return ""
                End If
                'NewExamNO = TIMS.Chk_NewExamNOc("", ExamOcid1, objconn)
                Dim flgChkExamNo As Boolean = TIMS.Chk_NewExamNOc(ExamPlanID, ExamOcid1, objconn)
                Dim NewExamNO As String = ""
                If flgChkExamNo Then NewExamNO = TIMS.Get_NewExamNOc(ExamPlanID, ExamNo1, ExamOcid1, objconn) '准考證號
                '取出准考證號   End
                If NewExamNO <> "" Then
                    sql = ""
                    sql &= " UPDATE STUD_ENTERTYPE "
                    sql &= " SET ExamNO = @ExamNO "
                    sql &= " WHERE OCID1 = @OCID1 AND SETID = @SETID AND EnterDate = @EnterDate AND SerNum = @SerNum"
                    oCmd = New SqlCommand(sql, objconn)
                    With oCmd
                        .Parameters.Clear()
                        .Parameters.Add("ExamNO", SqlDbType.VarChar).Value = NewExamNO
                        .Parameters.Add("OCID1", SqlDbType.Int).Value = Val(OCIDValue1)
                        .Parameters.Add("SETID", SqlDbType.Int).Value = Val(SETID)
                        .Parameters.Add("EnterDate", SqlDbType.DateTime).Value = CDate(EnterDate)
                        .Parameters.Add("SerNum", SqlDbType.Int).Value = Val(SerNum)
                        '.ExecuteNonQuery()  'edit，by:20181016
                        DbAccess.ExecuteNonQuery(oCmd.CommandText, objconn, oCmd.Parameters)  'edit，by:20181016
                    End With

                    'dr("ExamNO") = NewExamNO
                    'DbAccess.UpdateDataTable(dt, da)
                    If eSerNum <> "" Then
                        'conn = DbAccess.GetConnection
                        sql = ""
                        sql &= " UPDATE STUD_ENTERTYPE2 "
                        sql &= " SET ExamNO = @ExamNO "
                        sql &= " WHERE eSerNum = @eSerNum AND OCID1 = @OCID1 "
                        oCmd = New SqlCommand(sql, objconn)
                        With oCmd
                            .Parameters.Add("ExamNO", SqlDbType.VarChar).Value = NewExamNO
                            .Parameters.Add("eSerNum", SqlDbType.Int).Value = Val(eSerNum)
                            .Parameters.Add("OCID1", SqlDbType.Int).Value = Val(OCIDValue1)
                            '.ExecuteNonQuery()  'edit，by:20181016
                            DbAccess.ExecuteNonQuery(oCmd.CommandText, objconn, oCmd.Parameters)  'edit，by:20181016
                        End With
                    End If
                End If
                rstExamNo = NewExamNO
            End If
        End If
        Return rstExamNo
    End Function
End Class