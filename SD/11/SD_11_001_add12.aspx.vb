Partial Class SD_11_001_add12
    Inherits AuthBasePage

    '訓練期末學員滿意度調查表
    'Stud_Questionary ( FROM STUD_QUESTIONARY )
    'Plan_Questionary
    'id_Questionary
    Dim sPage_Url As String = ""
    'Const Cst_Page_Url As String = "SD_11_001_add.aspx"
    Const Cst_Page_Url_A2 As String = "SD_11_001_add12.aspx"
    Const Cst_defQ As String = "A2" '預設問卷A2 'A'B
    Dim vsQName As String = ""
    Dim vsQID As String = ""
    'vsQName = dr("QName").ToString 'viewstate("QName")
    'vsQID = dr("QID").ToString 'viewstate("QID")

    Const cst_SsesQueSchStr1 As String = "QuestionarySearchStr"
#Region "(No Use)"

    ''判斷是否設定問卷類別
    'Function GetQType() As Boolean
    '    Dim Rst As Boolean = False
    '    '判斷是否設定問卷類別
    '    '搜尋計畫問卷類型是否設定
    '    '若有設定 viewstate("QName")=QName ; viewstate("QID")=QID
    '    Dim sqlstr As String
    '    Dim dt As DataTable
    '    Dim dr As DataRow
    '    sqlstr = ""
    '    sqlstr += " select  b.QName,b.QID " & vbCrLf
    '    sqlstr += " from Plan_Questionary a" & vbCrLf
    '    sqlstr += " left join ID_Questionary b on a.QID=b.QID " & vbCrLf
    '    sqlstr += " where TPlanID='" & sm.UserInfo.TPlanID & "'"
    '    dt = DbAccess.GetDataTable(sqlstr)
    '    If dt.Rows.Count > 0 Then
    '        '計畫已設定問卷類型
    '        dr = dt.Rows(0)
    '        viewstate("QName") = dr("QName").ToString
    '        viewstate("QID") = dr("QID").ToString
    '        Rst = True
    '    Else
    '        '計畫未設定問卷類型
    '        'Dim TD_Stud As HtmlTableCell = Me.FindControl("TD_Stud")
    '        'StdTr.Visible = False
    '        'Table3.Style("display") = "none"
    '        'Label1.Visible = True
    '        'Button1.Visible = False
    '        'Button2.Visible = False
    '        'next_but.Visible = False
    '        'Return False
    '        Rst = False
    '    End If
    '    Return Rst
    'End Function

    'OCID	NUMBER(10,0) 	班級代碼
    'STUDID	NVARCHAR2(12 CHAR) 學員班級編號
    'Q1_1	NUMBER(10,0) 請問您這次參加的職訓課程，對課程內容安排及銜接是否滿意
    'Q1_2	NUMBER(10,0) 請問您這次參加的職訓課程，對課程時數安排是否滿意
    'Q1_3	NUMBER(10,0) 請問您這次參加的職訓課程，對使用的上課教材與訓練設施（如工具 /材料）是否滿意
    'Q2_1	NUMBER(10,0) 請問您滿不滿意老師專業知識
    'Q2_2	NUMBER(10,0) 請問您滿不滿意老師教學態度及 教學耐心
    'Q2_3	NUMBER(10,0) 請問您滿不滿意老師實務操作之教導能力
    'Q2_4	NUMBER(10,0) 請問您滿不滿意老師與學員間之互動
    'Q2_5	NUMBER(10,0) 請問您對老師教學準備工作是否滿意
    'Q3_1	NUMBER(10,0) 請問您滿不滿意訓練單位的上課環境
    'Q3_2	NUMBER(10,0) 請問您這次參加的職訓課程，對訓練單位公共安全(如無障礙設施）是否滿意
    'Q3_3	NUMBER(10,0) 請問您對訓練單位行政支援（如求助導師及申訴管道）是否滿意
    'Q3_4	NUMBER(10,0) 請問您參加職訓的機構有提供就業輔導嗎
    'Q3_5	NUMBER(10,0) 那您滿不滿意其提供就業輔導服務？ (針對第四題回答有者)
    'Q3_6	NUMBER(10,0) 請問您這次參加的職訓課程，對訓練單位提供就業資訊是否滿意？(針對第四題回答有者)
    'Q3_7	NUMBER(10,0) 請問您這次參加的職訓課程，對訓練單位提供就業推介服務是否滿意？(針對第四題回答有者)
    'Q4_1	NUMBER(10,0) 請問您覺得自己上課內容吸收程度如何
    'Q4_2	NUMBER(10,0) 您對於自己職訓這段期間表現打幾分
    'Q4_3	NUMBER(10,0) 請問若用考試評估您的學習效果，您的學習效果如何
    'Q4_4	NUMBER(10,0) 請問若用交作業評估您的學習效果，您的學習效果如何
    'Q4_5	NUMBER(10,0) 請問若用實習評估您的學習效果，您的學習效果如何
    'Q4_6	NUMBER(10,0) 請問您受訓前的期待與受訓後的感受有沒有落差？感受如何
    'Q5_1	NUMBER(10,0) 請問您參加此職訓的目的之一是不是要考證照
    'Q5_2	NUMBER(10,0) 請問您這次參加的職訓課程， 對您的證照考試幫助大不大
    'Q5_3	NUMBER(10,0) 請問您這次參加的職訓課程， 有沒有考到證照
    'Q5_4	NUMBER(10,0) 考到證照後，對找工作有沒有幫助 ？(針對上一題回答有者)
    'Q6_1	NUMBER(10,0) 受訓所學知識技能，對找工作有沒有幫助
    'Q6_2	NUMBER(10,0) 受訓頒發的結業證書，對找工作有沒有幫助
    'Q6_3	NUMBER(10,0) 老師名氣，對找工作有沒有幫助
    'Q6_4	NUMBER(10,0) 職訓機構名聲，對找工作有沒有幫助

#End Region

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End
        'Button2.Attributes("onclick") = "history.go(-1);"

        hPageUrl.Value = Cst_Page_Url_A2 'Cst_Page_Url
        SOCID.Attributes("onchange") = "if(this.selectedIndex==0) return false;"
        ProcessType.Value = Request("ProcessType")
        ProcessType.Value = TIMS.ClearSQM(ProcessType.Value)

        '20080723 andy ------------------新增(問卷類型) start 
        If Not TIMS.GetQType(Me, vsQName, vsQID) Then
            '計畫未設定問卷類型，請先設定後，再進入問卷填寫
            Dim TD_Stud As HtmlTableCell = Me.FindControl("TD_Stud")
            StdTr.Visible = False
            Table3.Style("display") = "none"
            Label1.Visible = True
            Button1.Visible = False
            Button2.Visible = False
            next_but.Visible = False
            BtnBak.Visible = False
            Common.MessageBox(Me, "計畫未設定問卷類型，請先設定後，再進入問卷填寫")
            Exit Sub '離開此功能
        End If

        Call QuestionType() '問卷種類

        Select Case ProcessType.Value
            Case "Insert", "Next"
                Qtype_Value.Value = vsQName
                If vsQName = "B" Then
                    CustomValidator1.Enabled = False
                    RadioButtonList5_4.Enabled = True
                Else
                    RadioButtonList3_4.Attributes("onclick") = "disable_radio3();"
                    RadioButtonList5_3.Attributes("onclick") = "disable_radio5();"
                End If
        End Select

#Region "(No Use)"

        'Button2.Attributes("onclick") = "history.go(-1);"

        '20080723 andy ------------------新增(問卷類型)end
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
        '        If ProcessType.Value = "Update" Then
        '            If FunDr("Mod") = "0" AndAlso FunDr("Del") = "0" Then
        '                Button1.Enabled = False
        '            Else
        '                Button1.Enabled = True
        '            End If
        '        ElseIf ProcessType.Value = "Insert" Or ProcessType.Value = "Next" Then
        '            If FunDr("Adds") = "1" Then
        '                Button1.Enabled = True
        '            Else
        '                Button1.Enabled = False
        '            End If
        '        End If
        '    End If
        'End If

#End Region

        If Not IsPostBack Then
            cCreate1()
        End If
    End Sub

    ''' <summary>
    ''' 查詢班級學員 或學員
    ''' </summary>
    ''' <param name="OCID"></param>
    ''' <param name="StudentID"></param>
    ''' <param name="dt"></param>
    ''' <param name="oConn"></param>
    ''' <returns></returns>
    Public Shared Function GetClassStuddt2(ByVal OCID As String, ByVal StudentID As String, ByRef dt As DataTable, ByRef oConn As SqlConnection) As DataTable
        Dim parms As New Hashtable
        parms.Clear()

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT a.StudentID" & vbCrLf
        sql &= " ,concat(b.Name,'(',dbo.FN_CSTUDID2(a.STUDENTID),')') Name" & vbCrLf
        sql &= " ,a.SOCID" & vbCrLf
        sql &= " ,a.StudStatus" & vbCrLf
        sql &= " ,b.name cname" & vbCrLf
        sql &= ",CONVERT(VARCHAR, a.RejectTDate1, 111) RejectTDate1 " & vbCrLf
        sql &= ",CONVERT(VARCHAR, a.RejectTDate2, 111) RejectTDate2 " & vbCrLf
        sql &= " FROM CLASS_STUDENTSOFCLASS a" & vbCrLf
        sql &= " JOIN STUD_STUDENTINFO b ON b.SID=a.SID" & vbCrLf
        sql &= " JOIN STUD_SUBDATA c ON c.SID=a.SID" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND a.StudStatus NOT IN (2,3)" & vbCrLf '排除離退
        sql &= " AND A.OCID =@OCID" & vbCrLf
        parms.Add("OCID", OCID)
        If StudentID <> "" Then
            sql &= " AND a.StudentID=@StudentID" & vbCrLf
            parms.Add("StudentID", StudentID)
        End If
        dt = DbAccess.GetDataTable(sql, oConn, parms)
        Return dt
    End Function

    Sub cCreate1()
        Re_OCID.Value = TIMS.ClearSQM(Request("ocid"))
        Re_Studentid.Value = TIMS.ClearSQM(Request("Stuedntid"))

        If ProcessType.Value <> "Print" Then
            Dim dt As DataTable = Nothing
            dt = GetClassStuddt2(Re_OCID.Value, "", dt, objconn)
            dt.DefaultView.Sort = "StudentID"
            With SOCID
                .DataSource = dt
                .DataTextField = "Name"
                .DataValueField = "StudentID"
                .DataBind()
                .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
            End With
            Common.SetListItem(SOCID, Re_Studentid.Value)
            If Session(cst_SsesQueSchStr1) IsNot Nothing Then Me.ViewState(cst_SsesQueSchStr1) = Session(cst_SsesQueSchStr1)

            'Session(cst_SsesQueSchStr1) = Nothing

            dt = Nothing
            dt = GetClassStuddt2(Re_OCID.Value, Re_Studentid.Value, dt, objconn)
            If dt.Rows.Count = 0 Then Return

            Dim row As DataRow = dt.Rows(0) ' DbAccess.GetOneRow(Sql, objconn)
            Me.Label_Name.Text = row("cname")
            Me.Label_Stud.Text = row("studentid")
            Label_Status.Text = TIMS.GET_STUDSTATUS_N23(row("StudStatus"), row("RejectTDate1"), row("RejectTDate2"))
        End If

        '刪除學員問卷答案
        Select Case ProcessType.Value
            Case "clear"
                Button2.Visible = False
                BtnBak.Visible = True '清除專用因為若使用原上一頁btn會出現重覆的確認訊息
                Button1.Visible = True
                next_but.Visible = True
            Case "del"
                Re_OCID.Value = TIMS.ClearSQM(Re_OCID.Value)
                Re_Studentid.Value = TIMS.ClearSQM(Re_Studentid.Value)
                del_STDQUESTIONARY(Re_OCID.Value, Re_Studentid.Value, objconn)
            Case "check"              '檢視學員問卷答案
                SOCID.Enabled = False
                show_data2(Re_OCID.Value, Re_Studentid.Value, objconn)
                Button2.Visible = True '回上一頁
                BtnBak.Visible = False '清除專用因為若使用原上一頁btn會出現重覆的確認訊息
                Button1.Visible = False
                next_but.Visible = False
            Case "Edit"                '修改
                SOCID.Enabled = False
                show_data2(Re_OCID.Value, Re_Studentid.Value, objconn)
                Button2.Visible = True
                BtnBak.Visible = False '清除專用因為若使用原上一頁btn會出現重覆的確認訊息
                Button1.Visible = True
                next_but.Visible = False
            Case "Next"                '下一個
                check_next()
            Case "Print"               '列印空白
                next_but.Visible = False
                Button2.Visible = False
                BtnBak.Visible = False '清除專用因為若使用原上一頁btn會出現重覆的確認訊息
                Button1.Visible = False
                If Session(cst_SsesQueSchStr1) Is Nothing Then Session(cst_SsesQueSchStr1) = Me.ViewState(cst_SsesQueSchStr1)
                'Me.RegisterStartupScript("scripprint", "<script>printDoc();window.opener=null;window.open('','_self');window.close();</script>")
                TIMS.RegisterStartupScript(Me, TIMS.xBlockName(), "<script>printDoc();</script>")
            Case "Print2"              '列印
                show_data2(Re_OCID.Value, Re_Studentid.Value, objconn)
                next_but.Visible = False
                Button2.Visible = False
                BtnBak.Visible = False '清除專用因為若使用原上一頁btn會出現重覆的確認訊息
                Button1.Visible = False
                If Session(cst_SsesQueSchStr1) Is Nothing Then Session(cst_SsesQueSchStr1) = Me.ViewState(cst_SsesQueSchStr1)
                'Me.RegisterStartupScript("scripprint", "<script>printDoc();window.opener=null;window.open('','_self');window.close();</script>")
                TIMS.RegisterStartupScript(Me, TIMS.xBlockName(), "<script>printDoc();</script>")
            Case Else
                Button2.Visible = True
                BtnBak.Visible = False '清除專用因為若使用原上一頁btn會出現重覆的確認訊息
                Button1.Visible = True
                next_but.Visible = True
        End Select
    End Sub

    ''' <summary>
    ''' 刪除
    ''' </summary>
    ''' <param name="ReOCID"></param>
    ''' <param name="StrStudID"></param>
    ''' <param name="oConn"></param>
    Sub del_STDQUESTIONARY(ByVal ReOCID As String, ByVal StrStudID As String, ByRef oConn As SqlConnection)
        Re_Studentid.Value = TIMS.ClearSQM(Re_Studentid.Value)
        If Re_Studentid.Value = "" Then Return

        Dim sqlstrdel As String = "DELETE STUD_QUESTIONARY WHERE OCID=@OCID and StudID= @StudID"
        Dim parms As New Hashtable
        parms.Add("OCID", ReOCID)
        parms.Add("StudID", StrStudID)
        DbAccess.ExecuteNonQuery(sqlstrdel, oConn, parms)

        Button2.Visible = False
        BtnBak.Visible = True '清除專用因為若使用原上一頁btn會出現重覆的確認訊息
        Button1.Visible = True
        next_but.Visible = True
    End Sub

    ''' <summary> 取得學員問卷答案 </summary>
    ''' <param name="ReOCID"></param>
    ''' <param name="StrStudID"></param>
    ''' <param name="oConn"></param>
    ''' <returns></returns>
    Function GET_QUESTIONARY_ROW(ByVal ReOCID As String, ByVal StrStudID As String, ByRef oConn As SqlConnection) As DataRow
        Dim sqlstr As String
        sqlstr = " SELECT * FROM STUD_QUESTIONARY WHERE OCID =@OCID AND StudID =@StudID "
        Dim parms As New Hashtable
        parms.Add("OCID", ReOCID)
        parms.Add("StudID", StrStudID)
        Dim row_list As DataRow = DbAccess.GetOneRow(sqlstr, oConn, parms)
        Return row_list
    End Function

    ''' <summary>
    ''' 取得/顯示學員問卷答案
    ''' </summary>
    ''' <param name="ReOCID"></param>
    ''' <param name="StrStudID"></param>
    ''' <param name="oConn"></param>
    Sub show_data2(ByVal ReOCID As String, ByVal StrStudID As String, ByRef oConn As SqlConnection)
        '答案清除 -- Start
        RadioButtonList1_1.SelectedIndex = -1
        RadioButtonList1_2.SelectedIndex = -1
        RadioButtonList1_3.SelectedIndex = -1
        RadioButtonList2_1.SelectedIndex = -1
        RadioButtonList2_2.SelectedIndex = -1
        RadioButtonList2_3.SelectedIndex = -1
        RadioButtonList2_4.SelectedIndex = -1
        RadioButtonList2_5.SelectedIndex = -1
        RadioButtonList3_1.SelectedIndex = -1
        RadioButtonList3_2.SelectedIndex = -1
        RadioButtonList3_3.SelectedIndex = -1
        RadioButtonList3_4.SelectedIndex = -1
        RadioButtonList3_5.SelectedIndex = -1
        RadioButtonList3_6.SelectedIndex = -1
        RadioButtonList3_7.SelectedIndex = -1
        RadioButtonList4_1.SelectedIndex = -1
        RadioButtonList4_2.SelectedIndex = -1
        RadioButtonList4_3.SelectedIndex = -1
        RadioButtonList4_4.SelectedIndex = -1
        RadioButtonList4_5.SelectedIndex = -1
        RadioButtonList4_6.SelectedIndex = -1

        RadioButtonList5_1.SelectedIndex = -1
        RadioButtonList5_2.SelectedIndex = -1
        RadioButtonList5_3.SelectedIndex = -1
        RadioButtonList5_4.SelectedIndex = -1
        RadioButtonList6_1.SelectedIndex = -1
        RadioButtonList6_2.SelectedIndex = -1
        RadioButtonList6_3.SelectedIndex = -1
        RadioButtonList6_4.SelectedIndex = -1
        'RadioButtonList6_5.SelectedIndex = -1
        '答案清除 -- End

        '(取得學員問卷答案)
        Dim row_list As DataRow = get_QUESTIONARY_row(ReOCID, StrStudID, oConn)
        If row_list Is Nothing Then Return

        '有資料
        If row_list("Q1_1").ToString <> "" Then Common.SetListItem(RadioButtonList1_1, row_list("Q1_1").ToString)
        If row_list("Q1_2").ToString <> "" Then Common.SetListItem(RadioButtonList1_2, row_list("Q1_2").ToString)
        If row_list("Q1_3").ToString <> "" Then Common.SetListItem(RadioButtonList1_3, row_list("Q1_3").ToString)
        If row_list("Q2_1").ToString <> "" Then Common.SetListItem(RadioButtonList2_1, row_list("Q2_1").ToString)
        If row_list("Q2_2").ToString <> "" Then Common.SetListItem(RadioButtonList2_2, row_list("Q2_2").ToString)
        If row_list("Q2_3").ToString <> "" Then Common.SetListItem(RadioButtonList2_3, row_list("Q2_3").ToString)
        If row_list("Q2_4").ToString <> "" Then Common.SetListItem(RadioButtonList2_4, row_list("Q2_4").ToString)
        If row_list("Q2_5").ToString <> "" Then Common.SetListItem(RadioButtonList2_5, row_list("Q2_5").ToString)
        If row_list("Q3_1").ToString <> "" Then Common.SetListItem(RadioButtonList3_1, row_list("Q3_1").ToString)
        If row_list("Q3_2").ToString <> "" Then Common.SetListItem(RadioButtonList3_2, row_list("Q3_2").ToString)
        If row_list("Q3_3").ToString <> "" Then Common.SetListItem(RadioButtonList3_3, row_list("Q3_3").ToString)
        If row_list("Q3_4").ToString <> "" Then Common.SetListItem(RadioButtonList3_4, row_list("Q3_4").ToString)
        If row_list("Q3_5").ToString <> "" Then Common.SetListItem(RadioButtonList3_5, row_list("Q3_5").ToString)
        If row_list("Q3_6").ToString <> "" Then Common.SetListItem(RadioButtonList3_6, row_list("Q3_6").ToString)
        If row_list("Q3_7").ToString <> "" Then Common.SetListItem(RadioButtonList3_7, row_list("Q3_7").ToString)
        If row_list("Q4_1").ToString <> "" Then Common.SetListItem(RadioButtonList4_1, row_list("Q4_1").ToString)
        If row_list("Q4_2").ToString <> "" Then Common.SetListItem(RadioButtonList4_2, row_list("Q4_2").ToString)
        If row_list("Q4_3").ToString <> "" Then Common.SetListItem(RadioButtonList4_3, row_list("Q4_3").ToString)
        If row_list("Q4_4").ToString <> "" Then Common.SetListItem(RadioButtonList4_4, row_list("Q4_4").ToString)
        If row_list("Q4_5").ToString <> "" Then Common.SetListItem(RadioButtonList4_5, row_list("Q4_5").ToString)
        If row_list("Q4_6").ToString <> "" Then Common.SetListItem(RadioButtonList4_6, row_list("Q4_6").ToString)

        If row_list("Q5_1").ToString <> "" Then Common.SetListItem(RadioButtonList5_1, row_list("Q5_1").ToString)
        If row_list("Q5_2").ToString <> "" Then Common.SetListItem(RadioButtonList5_2, row_list("Q5_2").ToString)
        If row_list("Q5_3").ToString <> "" Then Common.SetListItem(RadioButtonList5_3, row_list("Q5_3").ToString)
        If row_list("Q5_4").ToString <> "" Then Common.SetListItem(RadioButtonList5_4, row_list("Q5_4").ToString)
        If Not IsDBNull(row_list("Q6_1")) Then Common.SetListItem(RadioButtonList6_1, row_list("Q6_1"))
        If Not IsDBNull(row_list("Q6_2")) Then Common.SetListItem(RadioButtonList6_2, row_list("Q6_2"))
        If Not IsDBNull(row_list("Q6_3")) Then Common.SetListItem(RadioButtonList6_3, row_list("Q6_3"))
        If Not IsDBNull(row_list("Q6_4")) Then Common.SetListItem(RadioButtonList6_4, row_list("Q6_4"))
        'If Not IsDBNull(row_list("Q6_5")) Then Common.SetListItem(RadioButtonList6_5, row_list("Q6_5"))

        Select Case ProcessType.Value
            Case "check", "Print2"
            Case Else '"Edit"
                '顯示提示字眼
                Dim strScript As String = ""
                strScript = "<script language=""javascript"">" + vbCrLf
                strScript += "alert('此學員，已經填寫!!');" + vbCrLf
                strScript += "</script>"
                Page.RegisterStartupScript("", strScript)
        End Select
    End Sub

    '驗證部份問題
    Private Sub CustomValidator1_ServerValidate(ByVal source As System.Object, ByVal args As System.Web.UI.WebControls.ServerValidateEventArgs) Handles CustomValidator1.ServerValidate
        Dim rIsValid As Boolean = False '(預設)未通過驗證
        If RadioButtonList3_4.SelectedValue = "2" Then
            rIsValid = True '通過驗證
            'args.IsValid = True '通過驗證
        Else
            args.IsValid = False
            source.errormessage = ""
            If RadioButtonList3_5.SelectedValue = "" Then source.errormessage &= "請選擇第三部分的問題五" & vbCrLf
            If RadioButtonList3_6.SelectedValue = "" Then source.errormessage &= "請選擇第三部分的問題六" & vbCrLf
            If RadioButtonList3_7.SelectedValue = "" Then source.errormessage &= "請選擇第三部分的問題七" & vbCrLf
            If source.errormessage = "" Then rIsValid = True '通過驗證
        End If
        If RadioButtonList5_3.SelectedValue = "2" Then
            rIsValid = True '通過驗證
            'args.IsValid = True '通過驗證
        ElseIf RadioButtonList5_3.SelectedValue = "" Then
            rIsValid = True '通過驗證
            'args.IsValid = True '通過驗證
        ElseIf RadioButtonList5_3.SelectedValue = "1" Then
            'args.IsValid = False
            source.errormessage = ""
            If RadioButtonList5_4.SelectedValue = "" Then source.errormessage &= "請選擇第五部分的問題四" & vbCrLf
            If source.errormessage = "" Then rIsValid = True '通過驗證
        End If
        args.IsValid = rIsValid '最後 是否 通過驗證
    End Sub

    '預先載入'頁面載入完成之前。
    Private Sub Page_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.PreRender
        Dim strMessage As String = ""
        For Each obj As WebControls.BaseValidator In Page.Validators
            If obj.IsValid = False Then strMessage &= obj.ErrorMessage & vbCrLf
        Next
        If strMessage <> "" Then Common.MessageBox(Page, strMessage)
    End Sub

    '已為此班級中最後一筆學員!!(顯示訊息)
    Private Sub check_last()
        If Session(cst_SsesQueSchStr1) Is Nothing Then Session(cst_SsesQueSchStr1) = Me.ViewState(cst_SsesQueSchStr1)

        Dim strScript As String
        strScript = "<script language=""javascript"">" + vbCrLf
        strScript += "alert('已為此班級中最後一筆學員!!');" + vbCrLf
        strScript += "location.href ='SD_11_001.aspx?ProcessType=Back&ID='+document.getElementById('Re_ID').value;" + vbCrLf
        strScript += "</script>"
        Page.RegisterStartupScript("", strScript)
    End Sub

    '取得下一位新增學員(若沒有才可新增，若直到最後一位，則顯示訊息)
    Private Sub check_next()
        Dim Student_Table As DataTable
        Dim rows() As DataRow

        If Session("DTable_Stuednt") Is Nothing Then
            Common.RespWrite(Me, "<script>alert('您的登入資訊已經遺失，請重新登入');top.location.href='../../MOICA_Login.aspx';</script>")
            Me.Response.End()
            Exit Sub
        End If
        Student_Table = Session("DTable_Stuednt")
        If Student_Table Is Nothing Then
            Common.RespWrite(Me, "<script>alert('您的登入資訊已經遺失，請重新登入');top.location.href='../../MOICA_Login.aspx';</script>")
            Me.Response.End()
            Exit Sub
        End If
        If Student_Table.Rows.Count = 0 Then
            Common.RespWrite(Me, "<script>alert('您的登入資訊已經遺失，請重新登入');top.location.href='../../MOICA_Login.aspx';</script>")
            Me.Response.End()
            Exit Sub
        End If

        If Student_Table.Select("studentid > '" & Re_Studentid.Value & "'").Length = 0 Then
            check_last() '已為此班級中最後一筆學員!!
        End If
        rows = Student_Table.Select("studentid > '" & Re_Studentid.Value & "'")
        If rows.Length = 0 Then
            check_last() '已為此班級中最後一筆學員!!
        End If

        Try
            For i As Integer = 0 To rows.Length - 1
                Dim dr As DataRow = rows(i) 'Stud_Questionary

                Dim ReOCID As String = dr("OCID").ToString()
                Dim StrStudID As String = dr("studentid").ToString()
                Dim row_list As DataRow = get_QUESTIONARY_row(ReOCID, StrStudID, objconn)
                'If row_list Is Nothing Then Return
                If row_list Is Nothing Then
                    '取得下1位學員基礎資料 (學員號與班級號資料)
                    Re_OCID.Value = dr("OCID").ToString()
                    Re_Studentid.Value = dr("studentid").ToString()
                    '存取搜尋頁面條件
                    If Session(cst_SsesQueSchStr1) Is Nothing Then Session(cst_SsesQueSchStr1) = Me.ViewState(cst_SsesQueSchStr1)

                    ''重新呼叫頁面(重新載入)
                    'Response.Redirect(Cst_Page_Url & "?ocid=" & Me.Re_OCID.Value & "&Stuedntid=" & Re_Studentid.Value & "&ID=" & Re_ID.Value & "")
                    '重新呼叫頁面(重新載入)
                    TIMS.Utl_Redirect1(Me, Cst_Page_Url_A2 & "?ocid=" & Me.Re_OCID.Value & "&Stuedntid=" & Re_Studentid.Value & "&ID=" & Re_ID.Value & "")
                    Exit For '執行不到這一句，但還是寫離開此迴圈
                End If
            Next

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    '選擇 另一學員
    Private Sub SOCID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SOCID.SelectedIndexChanged
        'Re_OCID.Value = TIMS.ClearSQM(Request("ocid"))
        'Re_Studentid.Value = TIMS.ClearSQM(Request("Stuedntid"))
        Re_Studentid.Value = SOCID.SelectedValue
        'create(SOCID.SelectedValue)
        show_data2(Re_OCID.Value, Re_Studentid.Value, objconn)
    End Sub

    '問卷種類 依 viewstate("QName")
    Sub QuestionType()
        'Dim TD_R3_4 As HtmlTableCell = Me.FindControl("TD_R3_4")
        'Dim TD_R3_5 As HtmlTableCell = Me.FindControl("TD_R3_5")
        'Dim TD_R3_6 As HtmlTableCell = Me.FindControl("TD_R3_6")
        'Dim TD_R3_7 As HtmlTableCell = Me.FindControl("TD_R3_7")
        'Dim TD_R6_1 As HtmlTableCell = Me.FindControl("TD_R6_1")
        'Dim TD_R6_2 As HtmlTableCell = Me.FindControl("TD_R6_2")
        'Dim TD_R6_3 As HtmlTableCell = Me.FindControl("TD_R6_3")
        'Dim TD_R6_4 As HtmlTableCell = Me.FindControl("TD_R6_4")
        'Dim TD_R6_5 As HtmlTableCell = Me.FindControl("TD_R6_5")
        'Dim TD_R6 As HtmlTableCell = Me.FindControl("TD_R6")
        Dim sqlstr As String = ""
        Dim qtype As String = vsQName
        Select Case qtype '問卷類型  ID_Questionary
            Case "B"
                '重新設定問卷題目 (在職)
                CustomValidator1.Enabled = False
                TD_R3_4.Style("display") = "none"
                TD_R3_5.Style("display") = "none"
                TD_R3_6.Style("display") = "none"
                TD_R3_7.Style("display") = "none"
                TD_R6_1.Style("display") = "none"
                TD_R6_2.Style("display") = "none"
                TD_R6_3.Style("display") = "none"
                TD_R6_4.Style("display") = "none"
                'TD_R6_5.Style("display") = "none"
                TD_R6.Style("display") = "none"
                Re_R3_4.Enabled = False
                Re_R6_1.Enabled = False
                Re_R6_2.Enabled = False
                Re_R6_3.Enabled = False
                Re_R6_4.Enabled = False
                'Re_R6_5.Enabled = False
                '第一部份
                Label_R1_1.Text = "1.請問您對這次的職訓課程內容安排是否滿意？"
                Label_R1_2.Text = "2.請問您對這次的職訓課程時數安排是否滿意？"
                Label_R1_3.Text = "3.請問您對使用的上課教材與訓練設施（如工具 /材料）是否滿意？"
                '第二部份
                Label_R2_1.Text = "1.請問您是否滿意老師之專業知識？"
                Label_R2_2.Text = "2.請問您是否滿意老師之教學態度？"
                Label_R2_3.Text = "3.請問您是否滿意老師之教學方法？"
                Label_R2_4.Text = "4.請問您是否滿意老師之教材內容？"
                Label_R2_5.Text = "5.請問您是否滿意老師與學員間之互動？"
                '第三部份
                Label_R3_1.Text = "1.請問您是否滿意上課環境？"
                Label_R3_2.Text = "2.請問您對這次職訓上課地點公共設施(如消防安全及無障礙設施）是否滿意？"
                Label_R3_3.Text = "3.請問您對訓練單位行政支援（如問題解決及申訴管道）是否滿意？"
                '第四部份
                Label_R4_1.Text = "1.請問您對上課內容吸收程度如何？"
                Label_R4_2.Text = "2.您對於自己參加職訓這段期間的總體學習表現打幾分？"
                Label_R4_3.Text = "3.若用考試評估您的學習效果，您的學習效果如何？"
                Label_R4_4.Text = "4.若用交作業評估您的學習效果，您的學習效果如何？"
                Label_R4_5.Text = "5.若用實習評估您的學習效果，您的學習效果如何？"
                Label_R4_6.Text = "6.相較於您參訓前對課程的期待，您對本次參訓結果是否滿意？"
                '第五部份
                Common.AddClientScript(Page, "ChgFont('B');")
                Label_R5_1.Text = "1.受訓之課程內容與目前的工作內容是否相關？"
                Label_R5_2.Text = "2.受訓所學知識技能，對目前工作或轉業有無幫助？"
                Label_R5_3.Text = "3.您對本次訓練認為最需要改進的地方為何？（單選）"
                Label_R5_4.Text = "4.您是否有意願繼續參加與工作有關之進修訓練活動？"

                '第一次
                With RadioButtonList5_1
                    If .SelectedIndex = -1 Then
                        .Items.Clear()
                        .Items.Insert(0, New ListItem("是", 1))
                        .Items.Insert(1, New ListItem("否", 2))
                        .AppendDataBoundItems = True
                    End If
                End With
                With RadioButtonList5_3
                    If .SelectedIndex = -1 Then
                        .Items.Clear()
                        .Items.Insert(0, New ListItem("很滿意，無需改進", 1))
                        .Items.Insert(1, New ListItem("參訓職類不符就業市場需求", 2))
                        .Items.Insert(2, New ListItem("訓練設備不符產業需求", 3))
                        .Items.Insert(3, New ListItem("訓練時數不足", 4))
                        .Items.Insert(4, New ListItem("教學課程安排不當", 5))
                        .Items.Insert(5, New ListItem("訓練師專業及熱忱不足", 6))
                        .AppendDataBoundItems = True
                        .RepeatColumns = 1
                        .Enabled = True
                    End If
                End With
                With RadioButtonList5_4
                    If .SelectedIndex = -1 Then
                        .Items.Clear()
                        .Items.Insert(0, New ListItem("政府無補助也願意", 1))
                        .Items.Insert(1, New ListItem("政府提供50%以上之補助才願意", 2))
                        .Items.Insert(2, New ListItem("政府有補助才願意，無補助就不願意", 3))
                        .Items.Insert(3, New ListItem("政府提供補助也不願意", 4))
                        .AppendDataBoundItems = True
                        .RepeatColumns = 1
                        .Enabled = True
                    End If
                End With

            Case Else
                '預設問卷為「A2」
                Re_R5_2.Enabled = False
                Re_R5_3.Enabled = False
                Re_R5_4.Enabled = False
                TD_R3_4.Style("display") = "inline"
                TD_R3_5.Style("display") = "inline"
                TD_R3_6.Style("display") = "inline"
                TD_R3_7.Style("display") = "inline"
                TD_R6_1.Style("display") = "inline"
                TD_R6_2.Style("display") = "inline"
                TD_R6_3.Style("display") = "inline"
                TD_R6_4.Style("display") = "inline"
                'TD_R6_5.Style("display") = "inline"
                TD_R6.Style("display") = "inline"
        End Select
    End Sub

    '儲存問卷答案
    Sub Fill_Questionary()

        Dim s_State As String = ""
        Const cst_update As String = "update"
        Const cst_add As String = "add"

        If Not Page.IsValid Then Exit Sub
        If Session(cst_SsesQueSchStr1) Is Nothing Then Session(cst_SsesQueSchStr1) = Me.ViewState(cst_SsesQueSchStr1)

        Dim ReOCID As String = TIMS.ClearSQM(Re_OCID.Value)
        Dim StrStudID As String = TIMS.ClearSQM(Re_Studentid.Value)

        Dim sqlstr_update As String = ""
        sqlstr_update = String.Format(" SELECT * FROM STUD_QUESTIONARY WHERE OCID = '{0}' AND StudID = '{1}' ", ReOCID, StrStudID)
        Dim sqldr As DataRow = DbAccess.GetOneRow(sqlstr_update, objconn)
        Dim dr_row As DataRow = Nothing
        Dim sqlAdapter As SqlDataAdapter = Nothing
        Dim dtTable As DataTable = Nothing
        If Not sqldr Is Nothing Then
            s_State = cst_update
            dr_row = DbAccess.GetUpdateRow(sqlstr_update, dtTable, sqlAdapter, objconn)
        Else
            s_State = cst_add
            dr_row = DbAccess.GetInsertRow("Stud_Questionary", dtTable, sqlAdapter, objconn)
            dr_row = dtTable.NewRow
            dr_row("OCID") = Re_OCID.Value
            dr_row("StudID") = Re_Studentid.Value
        End If
        dr_row("FillFormDate") = Now()
        Select Case vsQName
            Case Cst_defQ  '問卷A2
                dr_row("Q1_1") = RadioButtonList1_1.SelectedValue
                dr_row("Q1_2") = RadioButtonList1_2.SelectedValue
                dr_row("Q1_3") = RadioButtonList1_3.SelectedValue
                dr_row("Q2_1") = RadioButtonList2_1.SelectedValue
                dr_row("Q2_2") = RadioButtonList2_2.SelectedValue
                dr_row("Q2_3") = RadioButtonList2_3.SelectedValue
                dr_row("Q2_4") = RadioButtonList2_4.SelectedValue
                dr_row("Q2_5") = RadioButtonList2_5.SelectedValue
                dr_row("Q3_1") = RadioButtonList3_1.SelectedValue
                dr_row("Q3_2") = RadioButtonList3_2.SelectedValue
                dr_row("Q3_3") = RadioButtonList3_3.SelectedValue
                dr_row("Q3_4") = RadioButtonList3_4.SelectedValue
                If RadioButtonList3_5.SelectedValue <> "" And RadioButtonList3_5.Enabled = True Then
                    dr_row("Q3_5") = RadioButtonList3_5.SelectedValue
                Else
                    dr_row("Q3_5") = Convert.DBNull
                End If
                If RadioButtonList3_6.SelectedValue <> "" And RadioButtonList3_6.Enabled = True Then
                    dr_row("Q3_6") = RadioButtonList3_6.SelectedValue
                Else
                    dr_row("Q3_6") = Convert.DBNull
                End If
                If RadioButtonList3_7.SelectedValue <> "" And RadioButtonList3_7.Enabled = True Then
                    dr_row("Q3_7") = RadioButtonList3_7.SelectedValue
                Else
                    dr_row("Q3_7") = Convert.DBNull
                End If
                dr_row("Q4_1") = RadioButtonList4_1.SelectedValue
                dr_row("Q4_2") = RadioButtonList4_2.SelectedValue
                dr_row("Q4_3") = RadioButtonList4_3.SelectedValue
                dr_row("Q4_4") = RadioButtonList4_4.SelectedValue
                dr_row("Q4_5") = RadioButtonList4_5.SelectedValue
                dr_row("Q4_6") = RadioButtonList4_6.SelectedValue
                dr_row("Q5_1") = RadioButtonList5_1.SelectedValue
                dr_row("Q5_2") = If(RadioButtonList5_2.SelectedValue <> "", RadioButtonList5_2.SelectedValue, Convert.DBNull)
                dr_row("Q5_3") = If(RadioButtonList5_3.SelectedValue <> "", RadioButtonList5_3.SelectedValue, Convert.DBNull)
                dr_row("Q5_4") = If(RadioButtonList5_4.SelectedValue <> "", RadioButtonList5_4.SelectedValue, Convert.DBNull)

                dr_row("Q6_1") = RadioButtonList6_1.SelectedValue
                dr_row("Q6_2") = RadioButtonList6_2.SelectedValue
                dr_row("Q6_3") = RadioButtonList6_3.SelectedValue
                dr_row("Q6_4") = RadioButtonList6_4.SelectedValue
                'dr_row("Q6_5") = RadioButtonList6_5.SelectedValue
            Case "B"   '問卷B
                '第一部份
                dr_row("Q1_1") = RadioButtonList1_1.SelectedValue
                dr_row("Q1_2") = RadioButtonList1_2.SelectedValue
                dr_row("Q1_3") = RadioButtonList1_3.SelectedValue
                '第二部份
                dr_row("Q2_1") = RadioButtonList2_1.SelectedValue
                dr_row("Q2_2") = RadioButtonList2_2.SelectedValue
                dr_row("Q2_3") = RadioButtonList2_3.SelectedValue
                dr_row("Q2_4") = RadioButtonList2_4.SelectedValue
                dr_row("Q2_5") = RadioButtonList2_5.SelectedValue
                '第三部份
                dr_row("Q3_1") = RadioButtonList3_1.SelectedValue
                dr_row("Q3_2") = RadioButtonList3_2.SelectedValue
                dr_row("Q3_3") = RadioButtonList3_3.SelectedValue
                '第四部份
                dr_row("Q4_1") = RadioButtonList4_1.SelectedValue
                dr_row("Q4_2") = RadioButtonList4_2.SelectedValue
                dr_row("Q4_3") = RadioButtonList4_3.SelectedValue
                dr_row("Q4_4") = RadioButtonList4_4.SelectedValue
                dr_row("Q4_5") = RadioButtonList4_5.SelectedValue
                dr_row("Q4_6") = RadioButtonList4_6.SelectedValue
                '第五部份
                dr_row("Q5_1") = RadioButtonList5_1.SelectedValue
                dr_row("Q5_2") = RadioButtonList5_2.SelectedValue
                dr_row("Q5_3") = RadioButtonList5_3.SelectedValue
                dr_row("Q5_4") = RadioButtonList5_4.SelectedValue
        End Select
        dr_row("QID") = CInt(vsQID)
        dr_row("ModifyAcct") = sm.UserInfo.UserID
        dr_row("ModifyDate") = Now()
        Select Case s_State
            Case cst_add
                dtTable.Rows.Add(dr_row)
        End Select
        sqlAdapter.Update(dtTable)

        If ProcessType.Value <> "Edit" Then '是新增或是填下一個才跑這一個
            Common.AddClientScript(Page, "insert_next();")
        Else
            Common.AddClientScript(Page, "BAK();")
            ''Session(cst_SsesQueSchStr1) = Me.ViewState(cst_SsesQueSchStr1)
            'Response.Redirect("SD_11_001.aspx?ProcessType=Back&ID=" & Re_ID.Value & "&ocid=" & Re_OCID.Value & "")
        End If
    End Sub

    '儲存
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call Fill_Questionary() '儲存問卷答案
    End Sub

    '回上一頁
    Private Sub BtnBak_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnBak.Click
        Common.AddClientScript(Page, "BAK();")
    End Sub

    '不儲存填寫下一位
    Private Sub next_but_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles next_but.Click
        If SOCID.Items.Count > 0 Then
            Dim NowIndex As Integer
            Dim MaxIndex As Integer
            MaxIndex = SOCID.Items.Count - 1
            NowIndex = SOCID.SelectedIndex
            If NowIndex = MaxIndex Then
                check_last() '已為此班級中最後一筆學員!!
            Else
                SOCID.SelectedIndex = NowIndex + 1
                Re_Studentid.Value = SOCID.SelectedValue
                'create(SOCID.SelectedValue)
                show_data2(Re_OCID.Value, Re_Studentid.Value, objconn)
            End If
        End If
    End Sub
End Class