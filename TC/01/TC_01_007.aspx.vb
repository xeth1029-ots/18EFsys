Partial Class TC_01_007
    Inherits AuthBasePage

    'TC_01_027-產投使用,'TC_01_007-非產投/TIMS使用
    Const cst_printFN1 As String = "TC_01_007" 'TC_01_027
    Const cst_printFN2 As String = "Teach"

    Dim sMemo As String = "" '(查詢原因)
    Dim flag_File1_xls As Boolean = False
    Dim flag_File1_ods As Boolean = False
    Dim ff As String = ""
    'colArray(
    Const cst_i計劃階層 As Integer = 0
    Const cst_i講師代碼 As Integer = 1
    Const cst_i講師姓名 As Integer = 2
    Const cst_i講師英文姓名 As Integer = 3
    Const cst_i身份別 As Integer = 4
    Const cst_i身分證字號 As Integer = 5
    Const cst_i出生日期 As Integer = 6
    Const cst_i性別 As Integer = 7
    Const cst_i主要職類 As Integer = 8
    Const cst_i職稱 As Integer = 9
    Const cst_i內外聘 As Integer = 10
    Const cst_i師資別 As Integer = 11
    Const cst_i最高學歷 As Integer = 12
    Const cst_i畢業狀況 As Integer = 13
    Const cst_i學校名稱 As Integer = 14
    Const cst_i科系名稱 As Integer = 15
    Const cst_i聯絡電話 As Integer = 16
    Const cst_i行動電話 As Integer = 17
    Const cst_i電子郵件 As Integer = 18
    Const cst_i郵遞區號前3碼 As Integer = 19
    Const cst_i郵遞區號後6碼 As Integer = 20
    Const cst_i通訊地址 As Integer = 21
    Const cst_i服務單位名稱 As Integer = 22
    Const cst_i年資 As Integer = 23
    Const cst_i服務部門 As Integer = 24
    Const cst_i服務單位電話 As Integer = 25
    Const cst_i服務單位傳真 As Integer = 26
    Const cst_i服務單位郵遞區號前3碼 As Integer = 27
    Const cst_i服務單位郵遞區號後6碼 As Integer = 28
    Const cst_i服務單位地址 As Integer = 29
    Const cst_i服務單位一 As Integer = 30
    Const cst_i服務單位二 As Integer = 31
    Const cst_i服務單位三 As Integer = 32
    Const cst_i服務年資一 As Integer = 33
    Const cst_i服務年資二 As Integer = 34
    Const cst_i服務年資三 As Integer = 35
    Const cst_i服務期間一起日 As Integer = 36
    Const cst_i服務期間一迄日 As Integer = 37
    Const cst_i服務期間二起日 As Integer = 38
    Const cst_i服務期間二迄日 As Integer = 39
    Const cst_i服務期間三起日 As Integer = 40
    Const cst_i服務期間三迄日 As Integer = 41
    Const cst_i專長一 As Integer = 42
    Const cst_i專長二 As Integer = 43
    Const cst_i專長三 As Integer = 44
    Const cst_i專長四 As Integer = 45
    Const cst_i專長五 As Integer = 46
    Const cst_i譯著 As Integer = 47
    Const cst_i專業證照 As Integer = 48
    Const cst_i排課使用 As Integer = 49
    Const cst_i講師類別 As Integer = 50
    Const cst_i助教類別 As Integer = 51
    Const cst_i教師類別 As Integer = 50 '2018 add 在職教師類別
    Const cst_i第二教師類別 As Integer = 51 '2018 add 在職教師類別
    Const cst_i欄位長度 As Integer = 52

    'SELECT * FROM TEACH_TEACHERINFO WHERE ROWNUM <=10
    Dim ID_KindOfTeacher As DataTable = Nothing
    Dim ID_Invest As DataTable = Nothing
    Dim Key_Degree As DataTable = Nothing
    Dim Key_GradState As DataTable = Nothing
    Dim dtTRAINTYPE As DataTable = Nothing
    Dim iPYNum As Integer = 1 'iPYNum = TIMS.sUtl_GetPYNum(Me) '1:2017前 2:2017 3:2018
    'Dim au As New cAUTH
    Dim objconn As SqlConnection = Nothing

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload

        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Dim url1 As String = $"TC_01_027.aspx?ID={TIMS.Get_MRqID(Me)}"
            TIMS.Utl_Redirect(Me, objconn, url1)
        End If

        Call TIMS.OpenDbConn(objconn)

        Dim work2015 As String = TIMS.Utl_GetConfigSet("work2015")
        hidLockTime2.Value = If(work2015 = "Y", "1", "2") '1:啟用鎖定。2:未鎖定

        AddHandler Button1.Click, AddressOf SUtl_btnSearchData1 '查詢
        AddHandler Btn_XlsEmport.Click, AddressOf SUtl_btnSearchData1 '匯出
        AddHandler btndivPwdSubmit.Click, AddressOf SUtl_btnSearchData1 'hidSchBtnNum.value: 1.正常查詢 2.正常匯出

        '啟動個資法。
        'Button1.Attributes("onclick") = "aloader2on();"
        Button1.Attributes.Add("onclick", "return showLoginPwdDiv(1);")
        Button1.CommandName = "Button1"
        'Btn_XlsEmport_Click
        Btn_XlsEmport.Attributes.Add("onclick", "return showLoginPwdDiv(2);")
        Btn_XlsEmport.CommandName = "btnxlsemport"
        'btndivPwdSubmit.Attributes("onclick") = "aloader2on();"
        'Button1.Attributes("onclick") = "javascript:return search()"

        If Not IsPostBack Then
            msg.Text = ""
            'panelLoginDiv.Visible = False
            'panelLoginDiv.Style.Item("display") = "none"
            labChkMsg.Text = ""

            eMeng.Style("display") = HidVeMeng.Value 'VeMeng.Text = "none"
            center.Value = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID

            DropDownList3 = TIMS.Get_Invest(DropDownList3, objconn)

            '取出鍵詞-查詢原因-INQUIRY
            Dim V_INQUIRY As String = Session($"{TIMS.cst_GSE_V_INQUIRY}{TIMS.Get_MRqID(Me)}")
            If TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES) Then Call TIMS.GET_INQUIRY(ddl_INQUIRY_Sch, objconn, V_INQUIRY)

            '設定 資料與顯示 狀況
            Call CREATE1(0)
            DataGridTable.Visible = False
        End If

        Dim sql As String = ""
        sql = "SELECT KindID ,KINDNAME,KINDENGAGE FROM ID_KINDOFTEACHER ORDER BY KINDID"
        ID_KindOfTeacher = DbAccess.GetDataTable(sql, objconn)
        sql = "SELECT IVID,InvestName FROM ID_INVEST ORDER BY IVID"
        ID_Invest = DbAccess.GetDataTable(sql, objconn)
        sql = "SELECT DEGREEID,NAME FROM KEY_DEGREE ORDER BY DEGREEID"
        Key_Degree = DbAccess.GetDataTable(sql, objconn)
        '取出dt-畢業狀況代碼-師資資料設定
        Key_GradState = TIMS.Get_GradStateDt2(objconn)
        sql = "SELECT TMID,BUSID,BUSNAME,JOBID,JOBNAME,TRAINID,TRAINNAME FROM VIEW_TRAINTYPE WHERE TRAINID IS NOT NULL ORDER BY BUSID,JOBID,TMID"
        dtTRAINTYPE = DbAccess.GetDataTable(sql, objconn)
        '分頁設定 Start
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button5.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');ShowFrame();"
            HistoryRID.Attributes("onclick") = "ShowFrame();"
            center.Style("CURSOR") = "hand"
        End If

        '檢查帳號的功能權限-----------------------------------Start
        'Button2.Enabled = False
        'If au.blnCanAdds Then Button2.Enabled = True
        'Button1.Enabled = False
        'If au.blnCanSech Then Button1.Enabled = True
        '檢查帳號的功能權限-----------------------------------End

        'Button6.Attributes("onclick") = "if(document.form1.File1.value==''){alert('請選擇匯入檔案的路徑');return false;}"

        If Not IsPostBack Then
            lnkImpSample1.Visible = False
            lnkImpSample2.Visible = False
            If TIMS.Cst_TPlanID06Plan1.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                lnkImpSample2.Visible = True
            Else
                lnkImpSample1.Visible = True
            End If

            If Session("MySearchStr") IsNot Nothing Then
                Dim MyValue As String = ""
                Dim strSession As String = Session("MySearchStr")
                Session("MySearchStr") = Nothing

                center.Value = TIMS.GetMyValue(strSession, "center")
                RIDValue.Value = TIMS.GetMyValue(strSession, "RIDValue")
                'MyValue = TIMS.GetMyValue(strSession, "DropDownList1")
                'If MyValue <> "" Then Common.SetListItem(DropDownList1, MyValue)
                TextBox2.Text = TIMS.GetMyValue(strSession, "TextBox2")
                TextBox3.Text = TIMS.GetMyValue(strSession, "TextBox3")
                MyValue = TIMS.GetMyValue(strSession, "DropDownList2")
                If MyValue <> "" Then Common.SetListItem(DropDownList2, MyValue)
                TextBox4.Text = TIMS.GetMyValue(strSession, "TextBox4")
                MyValue = TIMS.GetMyValue(strSession, "DropDownList3")
                If MyValue <> "" Then Common.SetListItem(DropDownList3, MyValue)
                TB_career_id.Text = TIMS.GetMyValue(strSession, "TB_career_id")
                trainValue.Value = TIMS.GetMyValue(strSession, "trainValue")
                jobValue.Value = TIMS.GetMyValue(strSession, "jobValue")

                MyValue = TIMS.GetMyValue(strSession, "DropDownList4")
                If MyValue <> "" Then
                    Common.SetListItem(DropDownList4, MyValue)
                    Call Sub_DDL4Sel()
                End If
                MyValue = TIMS.GetMyValue(strSession, "DropDownList1")
                If MyValue <> "" Then Common.SetListItem(DropDownList1, MyValue)

                MyValue = TIMS.GetMyValue(strSession, "PageIndex")
                If MyValue <> "" Then PageControler1.PageIndex = MyValue
                MyValue = TIMS.GetMyValue(strSession, "Button1")
                'Button1_Click(sender, e)
                If MyValue = "True" Then Call GClickSearchButton()
            End If
        End If

    End Sub

    '設定 資料與顯示 狀況！
    Sub CREATE1(ByVal num As Integer)
        'num 0:第一次呼叫 --請選擇-- 1:內聘 2:外聘
        Select Case num
            Case 0
                DropDownList1.Items.Clear()
                DropDownList1.Items.Add(New ListItem("--請選擇內外聘--", 0))
                tr_techtype12.Visible = False
                tr_techtype34.Visible = False

                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    '產投
                    tr_techtype12.Visible = True '顯示 '講師 助教 (類別)
                ElseIf TIMS.Cst_TPlanID06Plan1.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    '2018 add 自辦在職
                    tr_techtype34.Visible = True ' 顯示 '教師 第二教師 (類別)
                End If
            Case Else '1.2
                DropDownList1 = TIMS.Get_KindOfTeacher(DropDownList1, CStr(num), "1", objconn)
        End Select
    End Sub

    Sub GetSearchStr()
        Session("MySearchStr") = Nothing

        Dim v_DropDownList1 As String = TIMS.GetListValue(DropDownList1)
        Dim v_DropDownList2 As String = TIMS.GetListValue(DropDownList2)
        Dim v_DropDownList3 As String = TIMS.GetListValue(DropDownList3)
        Dim v_DropDownList4 As String = TIMS.GetListValue(DropDownList4)

        Dim sMySearchStr As String = ""
        sMySearchStr &= "center=" & center.Value
        sMySearchStr &= "&RIDValue=" & RIDValue.Value
        sMySearchStr &= "&DropDownList1=" & v_DropDownList1 'DropDownList1.SelectedValue
        sMySearchStr &= "&TextBox2=" & TextBox2.Text
        sMySearchStr &= "&TextBox3=" & TextBox3.Text
        sMySearchStr &= "&DropDownList2=" & v_DropDownList2 'DropDownList2.SelectedValue
        sMySearchStr &= "&TextBox4=" & TextBox4.Text
        sMySearchStr &= "&DropDownList3=" & v_DropDownList3 'DropDownList3.SelectedValue
        sMySearchStr &= "&TB_career_id=" & TB_career_id.Text
        sMySearchStr &= "&trainValue=" & trainValue.Value
        sMySearchStr &= "&jobValue=" & jobValue.Value
        sMySearchStr &= "&DropDownList4=" & v_DropDownList4 'DropDownList4.SelectedValue
        sMySearchStr &= "&PageIndex=" & DataGrid1.CurrentPageIndex + 1
        sMySearchStr &= "&Button1=" & DataGrid1.Visible
        Session("MySearchStr") = sMySearchStr
    End Sub

    '刪除
    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e.CommandName = "" Then Exit Sub
        Dim MRqID As String = TIMS.Get_MRqID(Me)
        Select Case e.CommandName
            Case "edit"
                Call GetSearchStr()
                'e.CommandArgument@TechID
                'Response.Redirect("TC_01_007_add.aspx?proecess=edit&serial=" & e.CommandArgument & "&ID=" & MRqID)
                Dim url1 As String = "TC_01_007_add.aspx?proecess=edit&serial=" & e.CommandArgument & "&ID=" & MRqID
                TIMS.Utl_Redirect(Me, objconn, url1)
            Case "print"
                'Dim cGuid As String =   ReportQuery.GetGuid(Page)
                'Dim Url As String =   ReportQuery.GetUrl(Page)
                'Dim strScript As String
                'strScript = "<script language=""javascript"">" + vbCrLf
                'strScript &= "window.open('" & Url & "GUID=" + cGuid + "&SQ_AutoLogout=true&sys=list&filename=Teach&path=TIMS&TechID=" & e.CommandArgument & "');" + vbCrLf
                'strScript &= "</script>"
                'Page.RegisterStartupScript("window_onload", strScript)
            Case "del"
                'e.CommandArgument@TechID
                If Convert.ToString(e.CommandArgument) = "" Then
                    Common.MessageBox(Me, "傳入參數有誤，請重新查詢")
                    Exit Sub
                End If
                'e.CommandArgument@TechID
                Dim sTechID As String = e.CommandArgument
                Dim tmpTeacherName As String = TIMS.Get_TeachCName(sTechID, objconn) 'TIMS.Get_TeacherName(e.CommandArgument)
                If tmpTeacherName = "" Then
                    Common.MessageBox(Me, "查無該師姓名，請重新查詢")
                    Exit Sub
                End If

                If Not gDelTeach_TeacherInfo(e.CommandArgument) Then
                    Common.MessageBox(Me, "使用中，不可刪除")
                    Exit Sub
                End If
                sMemo = $"&動作=刪除&NAME={tmpTeacherName}"
                '寫入Log查詢(SubInsAccountLog1(Auth_Accountlog))
                Dim rqMID As String = TIMS.Get_MRqID(Me)
                Call TIMS.SubInsAccountLog1(Me, rqMID, TIMS.cst_wm刪除, TIMS.GetListValue(rblWorkMode), "", sMemo)

                Common.MessageBox(Me, "刪除完成")
                '查詢
                Call GClickSearchButton()

        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.Cells(0).Text = "序號"
                If cb_CourID.Checked Then e.Item.Cells(0).Text = "匯入用<BR>代碼"

            Case ListItemType.Item, ListItemType.AlternatingItem, ListItemType.EditItem
                Dim drv As DataRowView = e.Item.DataItem
                e.Item.Cells(0).Text = DataGrid1.CurrentPageIndex * DataGrid1.PageSize + e.Item.ItemIndex + 1
                If cb_CourID.Checked Then
                    e.Item.Cells(0).Text = CStr(drv("TechID"))
                End If
                'Dim row() As DataRow
                If Convert.ToString(drv("KindID")) <> "" Then
                    ff = "KindID='" & Convert.ToString(drv("KindID")) & "'"
                    If ID_KindOfTeacher.Select(ff).Length <> 0 Then
                        e.Item.Cells(4).Text = ID_KindOfTeacher.Select(ff)(0)("KindName")
                    End If
                End If
                Dim strKindEngage As String = Convert.ToString(drv("KindEngage"))
                Select Case Convert.ToString(drv("KindEngage"))
                    Case "1"
                        strKindEngage = "內聘(專任)"
                    Case "2"
                        strKindEngage = "外聘(兼任)"
                End Select
                e.Item.Cells(5).Text = strKindEngage

                Dim lbtEdit As LinkButton = e.Item.FindControl("lbtEdit")
                Dim lbtDel As LinkButton = e.Item.FindControl("lbtDel")
                Dim lbtPrt As LinkButton = e.Item.FindControl("lbtPrt")
                '修改/檢視
                'Dim but As Button = e.Item.Cells(6).FindControl("Button3")
                lbtEdit.CommandArgument = Convert.ToString(drv("TechID"))
                lbtEdit.Text = "檢視"
                If sm.UserInfo.RID = drv("RID") Then
                    lbtEdit.Text = "修改"
                ElseIf Len(sm.UserInfo.RID.ToString) = 1 Then
                    lbtEdit.Text = "修改"
                End If

                'If FunDr("Mod") = "1" Then lbtEdit.Enabled = True
                '2018 先 mark,todo... (按鈕權限控制尚未完成)
                'lbtEdit.Enabled = False
                'If au.blnCanMod Then lbtEdit.Enabled = True

                '刪除鈕
                'Dim btndelete As Button = e.Item.FindControl("btndelete")
                lbtDel.CommandArgument = Convert.ToString(drv("TechID"))
                lbtDel.Visible = False
                If sm.UserInfo.LID <= 1 Then
                    lbtDel.Attributes("onclick") = "javascript:return confirm('此動作會刪除師資資料，是否確定刪除?');"
                    lbtDel.Visible = True
                End If

                '列印師資資料
                'but = e.Item.Cells(6).FindControl("Button4")  '列印師資資料
                lbtPrt.CommandArgument = Convert.ToString(drv("TechID"))

                lbtPrt.Attributes("onclick") = ReportQuery.ReportScript(Me, cst_printFN2, "TechID=" & drv("TechID") & "")
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    '產投
                    lbtPrt.Visible = False '不顯示 '列印師資資料
                End If
        End Select

    End Sub

    '刪除 (含檢查使用狀況動作)
    Function gDelTeach_TeacherInfo(ByVal TechID As String) As Boolean
        Dim Rst As Boolean = False ' 刪除 有異常
        Dim flagCanDelete As Boolean = True '可以刪除
        Dim sql As String = ""
        'Dim dr As DataRow
        Dim dt As DataTable

        TechID = TIMS.ClearSQM(TechID)
        If TechID <> "" Then
            'TechID = TechID.Trim
            If IsNumeric(TechID) Then
                sql = "" & vbCrLf
                If flagCanDelete Then
                    '開班老師檔(產業人才投資方案)
                    sql = ""
                    sql &= " SELECT DISTINCT 'x1' x FROM CLASS_TEACHER WHERE TechID = '" & TechID & "'" & vbCrLf
                    dt = DbAccess.GetDataTable(sql, objconn)
                    If dt.Rows.Count > 0 Then
                        flagCanDelete = False '有資料，不可以刪除 
                    End If
                End If

                If flagCanDelete Then
                    '不預告實地抽查訪視記錄檔
                    'Sql &= " union" & vbCrLf
                    sql = ""
                    sql &= " SELECT distinct 'x21' x FROM CLASS_UNEXPECTVISITOR WHERE TechID = '" & TechID & "'" & vbCrLf
                    sql &= " union" & vbCrLf
                    sql &= " SELECT distinct 'x22' x FROM CLASS_UNEXPECTVISITOR WHERE TechID2 = '" & TechID & "'" & vbCrLf
                    dt = DbAccess.GetDataTable(sql, objconn)
                    If dt.Rows.Count > 0 Then
                        flagCanDelete = False '有資料，不可以刪除 
                    End If
                End If

                If flagCanDelete Then
                    '判斷是否有被使用
                    '-班級申請老師檔(產學訓)
                    'Sql &= " union" & vbCrLf
                    sql = ""
                    sql &= " SELECT distinct 'x3' x FROM PLAN_TEACHER WHERE TechID = '" & TechID & "'" & vbCrLf
                    dt = DbAccess.GetDataTable(sql, objconn)
                    If dt.Rows.Count > 0 Then
                        flagCanDelete = False '有資料，不可以刪除 
                    End If
                End If

                If flagCanDelete Then
                    '計畫訓練內容簡介(95年度)(97產學訓課程大綱)
                    'Sql &= " union" & vbCrLf
                    sql = ""
                    sql &= " SELECT distinct 'x4' x FROM Plan_TrainDesc WHERE TechID = '" & TechID & "'" & vbCrLf
                    sql &= " UNION" & vbCrLf
                    sql &= " SELECT distinct 'x5' x FROM Plan_TrainDesc WHERE TechID2 = '" & TechID & "'" & vbCrLf
                    dt = DbAccess.GetDataTable(sql, objconn)
                    If dt.Rows.Count > 0 Then
                        flagCanDelete = False '有資料，不可以刪除 
                    End If
                End If

                If flagCanDelete Then
                    '排課資訊
                    sql = "" & vbCrLf
                    sql &= " with x6 as (SELECT distinct 'x6' x FROM MVIEW_CLASS_SCHEDULE WHERE TechID = '" & TechID & "')" & vbCrLf
                    sql &= " ,x7 as (SELECT distinct 'x7' x FROM MVIEW_CLASS_SCHEDULE WHERE TechID2 = '" & TechID & "')" & vbCrLf
                    sql &= " select * from x6 union select * from x7" & vbCrLf
                    dt = DbAccess.GetDataTable(sql, objconn)
                    If dt.Rows.Count > 0 Then flagCanDelete = False '有資料，不可以刪除 
                End If

                If flagCanDelete Then
                    '排課資訊
                    sql = "" & vbCrLf
                    sql &= " with x6 as (SELECT distinct 'x6' x FROM VIEW_CLASS_SCHEDULE WHERE TechID = '" & TechID & "')" & vbCrLf
                    sql &= " ,x7 as (SELECT distinct 'x7' x FROM VIEW_CLASS_SCHEDULE WHERE TechID2 = '" & TechID & "')" & vbCrLf
                    sql &= " select * from x6 union select * from x7" & vbCrLf
                    dt = DbAccess.GetDataTable(sql, objconn)
                    If dt.Rows.Count > 0 Then flagCanDelete = False '有資料，不可以刪除 
                End If

                If flagCanDelete Then
                    '無使用資料 '可以刪除 
                    Try
                        sql = "DELETE TEACH_TEACHERINFO WHERE TechID= '" & TechID & "'"
                        DbAccess.ExecuteNonQuery(sql, objconn)
                        Rst = True ' 刪除 完成
                    Catch ex As Exception
                        Throw ex '刪除失敗
                    End Try
                End If
            End If
        End If
        Return Rst
    End Function

    '設定 資料與顯示 狀況
    Private Sub DropDownList4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DropDownList4.SelectedIndexChanged
        '設定 資料與顯示 狀況
        Call Sub_DDL4Sel()
    End Sub

    '設定 資料與顯示 狀況
    Sub Sub_DDL4Sel()
        Select Case DropDownList4.SelectedValue
            Case "1", "2"
                Call CREATE1(DropDownList4.SelectedValue)
            Case Else
                Call CREATE1(0)
        End Select
    End Sub

    '檢查輸入資料
    Function CheckImportData(ByRef colArray As Array) As String
        Dim Reason As String = ""
        'Dim i, j, subCount As Integer
        Dim sql As String
        'Dim intCol As Integer = 0
        Dim strCol1 As String = ""
        If colArray.Length < cst_i欄位長度 Then
            'Reason &= "欄位數量不正確(應該為47個欄位)<BR>"
            Reason &= "欄位對應有誤，資料欄位不足" & cst_i欄位長度 & "個欄位<BR>"
        Else
            colArray = TIMS.ChangeColArray(colArray)
            'If colArray(cst_i計劃階層).ToString = "" Then
            '    Reason &= "計劃階層必須填寫<Br>"
            'End If
            colArray(cst_i講師代碼) = TIMS.ClearSQM(colArray(cst_i講師代碼))
            If colArray(cst_i講師代碼).ToString = "" Then
                Reason &= "講師代碼必須填寫<Br>"
            Else
                If (colArray(cst_i講師代碼).ToString).Length > 10 Then
                    Reason &= "講師代碼不符合<BR>"
                End If
            End If

            colArray(cst_i講師姓名) = TIMS.ClearSQM(colArray(cst_i講師姓名))
            If colArray(cst_i講師姓名).ToString = "" Then
                Reason &= "講師姓名必須填寫<Br>"
            End If

            colArray(cst_i講師英文姓名) = TIMS.ClearSQM(colArray(cst_i講師英文姓名))
            If colArray(cst_i講師英文姓名).ToString <> "" Then
                colArray(cst_i講師英文姓名) = TIMS.ChangeIDNO(colArray(cst_i講師英文姓名), " ") '講師英文姓名
                If (colArray(cst_i講師英文姓名).ToString).Length > 30 Then
                    Reason &= "講師英文姓名 過長應小於等於30字字數<BR>"
                End If
            End If


            If colArray(cst_i身份別).ToString = "" Then
                Reason &= "身份別必須填寫<Br>"
            Else
                Select Case colArray(cst_i身份別).ToString
                    Case "1", "2"
                    Case Else
                        Reason &= "身份別必須輸入1或2(1.本國,2.外籍)<Br>"
                End Select
            End If
            If colArray(cst_i身分證字號).ToString = "" Then
                Reason &= "身分證必須填寫<Br>"
            Else
                If (colArray(cst_i身分證字號).ToString.Length <> 10) Then
                    Reason &= "身分證字數不符合<BR>"
                End If
            End If
            If Reason <> "" Then Return Reason '上述有錯誤離開

            colArray(cst_i身分證字號) = TIMS.ClearSQM(colArray(cst_i身分證字號))
            colArray(cst_i身分證字號) = TIMS.ChangeIDNO(colArray(cst_i身分證字號))
            Select Case colArray(cst_i身份別).ToString
                Case "1" '本國
                    If Not TIMS.CheckIDNO(colArray(cst_i身分證字號)) Then
                        Reason &= "身分證號碼有誤<BR>"
                    End If
                Case "2" '外藉
                    Dim nsIDNO As String = colArray(cst_i身分證字號)
                    '2:居留證 4:居留證2021
                    Dim flag2 As Boolean = TIMS.CheckIDNO2(nsIDNO, 2)
                    Dim flag4 As Boolean = TIMS.CheckIDNO2(nsIDNO, 4)
                    If Not flag2 AndAlso Not flag4 Then
                        Reason &= "身份別為外藉，居留證號有誤<BR>"
                    End If

                Case Else
                    Reason &= "身份別必須輸入1或2(1.本國,2.外籍)<Br>"
            End Select
            sql = ""
            sql &= " SELECT COUNT(1) CNT FROM Teach_TeacherInfo"
            sql &= " WHERE RID='" & IIf(RIDValue.Value = "", sm.UserInfo.RID, RIDValue.Value) & "'"
            sql &= " and IDNO='" & colArray(cst_i身分證字號).ToString & "' "
            Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)
            If CInt(dr("CNT").ToString) > 0 Then
                Reason &= "同一計劃，有相同身份號碼，重複輸入<BR>"
            End If


            If Convert.ToString(colArray(cst_i出生日期)) = "" Then
                'Reason &= "出生日期必須填寫<Br>"
            Else
                If IsDate(Convert.ToString(colArray(cst_i出生日期))) = False Then
                    Reason &= "出生日期必須是西元年格式(yyyy/mm/dd)<BR>"
                Else
                    Try
                        colArray(cst_i出生日期) = CDate(Convert.ToString(colArray(cst_i出生日期))).ToString("yyyy/MM/dd")
                        If CDate(Convert.ToString(colArray(cst_i出生日期))) < "1900/1/1" _
                            OrElse CDate(Convert.ToString(colArray(cst_i出生日期))) > "2100/1/1" Then
                            Reason &= "出生日期範圍有誤<BR>"
                        End If
                    Catch ex As Exception
                        Reason &= "出生日期必須是西元年格式(yyyy/mm/dd)<BR>"
                    End Try
                End If
            End If

            If colArray(cst_i性別).ToString = "" Then
                Reason &= "必須填寫性別<BR>"
            Else
                colArray(cst_i性別) = UCase(colArray(cst_i性別))
                Select Case colArray(cst_i性別).ToString
                    Case "M", "F"
                    Case Else
                        Reason &= "性別代號只能是M或者是F<BR>"
                End Select
            End If

            If colArray(cst_i主要職類).ToString = "" Then
                Reason &= "主要職類必須填寫<Br>"
            Else
                If Not IsNumeric(colArray(cst_i主要職類)) Then
                    Reason &= "主要職類必需為數字<BR>"
                Else
                    ff = "TMID='" & colArray(cst_i主要職類) & "'"
                    If dtTRAINTYPE.Select(ff).Length = 0 Then
                        Reason &= "主要職類不在鍵詞範圍內，請確認<BR>"
                    End If
                End If
            End If

            If colArray(cst_i職稱).ToString = "" Then
                Reason &= "職稱填寫是否正確<Br>"
            Else
                If RIDValue.Value.Length = 1 Then '分署(中心)以上單位
                    If IsNumeric(colArray(cst_i職稱)) = False Then
                        Reason &= "職稱必需為數字<BR>"
                    Else
                        colArray(cst_i職稱) = TIMS.ChangeIDNO(colArray(cst_i職稱))
                        If Len(colArray(cst_i職稱).ToString) < 2 Then
                            colArray(cst_i職稱) = "0" & colArray(cst_i職稱).ToString
                        End If
                        ff = "IVID='" & colArray(cst_i職稱) & "'"
                        If ID_Invest.Select(ff).Length = 0 Then
                            Reason &= "職稱不在鍵詞範圍內，請確認<BR>"
                        End If
                    End If
                Else
                    '委訓單位(輸入INVEST)
                End If
            End If

            If colArray(cst_i內外聘).ToString = "" Then
                Reason &= "內外聘必須填寫<Br>"
            Else
                If IsNumeric(colArray(cst_i內外聘)) = False Then
                    Reason &= "內外聘 必需為數字(1.內聘(專任)  2.外聘(兼任)(委訓單位))<BR>"
                Else
                    Select Case Convert.ToString(Val(colArray(cst_i內外聘)))
                        Case "1", "2"
                            colArray(cst_i內外聘) = Convert.ToString(Val(colArray(cst_i內外聘)))
                        Case Else
                            Reason &= "內外聘必需為數字(1.內聘(專任)  2.外聘(兼任)(委訓單位))<BR>"
                    End Select
                End If
            End If

            If colArray(cst_i師資別).ToString = "" Then
                Reason &= "師資別必須填寫<Br>"
            Else
                If IsNumeric(colArray(cst_i師資別)) = False Then
                    Reason &= "師資別必需為數字<BR>"
                Else
                    If Reason = "" Then
                        If RIDValue.Value.Length = 1 Then '分署(中心)以上單位
                            ff = "KindID='" & colArray(cst_i師資別) & "' AND KindEngage='" & Convert.ToString(CInt(colArray(cst_i內外聘))) & "'"
                            If ID_KindOfTeacher.Select(ff).Length = 0 Then
                                Reason &= "師資別不在鍵詞範圍內，請確認<BR>"
                            End If
                        Else
                            '委訓單位只能輸入130講師
                            If Convert.ToString(colArray(cst_i師資別)) <> "130" Then
                                Reason &= "委訓單位: 師資別只能輸入代碼:130(講師)<BR>"
                            End If
                        End If
                    End If
                End If
            End If

            If colArray(cst_i最高學歷).ToString = "" Then
                Reason &= "最高學歷必須填寫<Br>"
            Else
                If Len(colArray(cst_i最高學歷).ToString) < 2 Then
                    colArray(cst_i最高學歷) = "0" & colArray(cst_i最高學歷).ToString
                End If
                ff = "DegreeID='" & colArray(cst_i最高學歷) & "'"
                If Not Key_Degree.Select(ff).Length > 0 Then
                    Reason &= "最高學歷不在鍵詞範圍內，請確認<BR>"
                End If
            End If

            If colArray(cst_i畢業狀況).ToString = "" Then
                Reason &= "畢業狀況必須填寫<Br>"
            Else
                If Len(colArray(cst_i畢業狀況).ToString) < 2 Then
                    colArray(cst_i畢業狀況) = "0" & colArray(cst_i畢業狀況).ToString
                End If
                ff = String.Format("GradID='{0}'", colArray(cst_i畢業狀況))
                If Not Key_GradState.Select(ff).Length > 0 Then Reason &= "畢業狀況不在鍵詞範圍內，請確認<BR>"
            End If

            If colArray(cst_i學校名稱).ToString <> "" Then
                colArray(cst_i學校名稱) = TIMS.ClearSQM(colArray(cst_i學校名稱))
            End If
            If colArray(cst_i學校名稱).ToString <> "" Then
                If (colArray(cst_i學校名稱).ToString.Length > 30) Then
                    Reason &= "學校名稱必須小於等於中文25字<BR>"
                End If
            End If

            If colArray(cst_i科系名稱).ToString <> "" Then
                colArray(cst_i科系名稱) = TIMS.ClearSQM(colArray(cst_i科系名稱))
            End If
            If colArray(cst_i科系名稱).ToString <> "" Then
                If (colArray(cst_i科系名稱).ToString.Length > 25) Then
                    Reason &= "科系名稱必須小於等於中文25字<BR>"
                End If
            End If

            colArray(cst_i聯絡電話) = TIMS.ClearSQM(colArray(cst_i聯絡電話))
            If colArray(cst_i聯絡電話).ToString = "" Then
                Reason &= "聯絡電話必須填寫<Br>"
            Else
                If (colArray(cst_i聯絡電話).ToString.Length > 15) Then
                    Reason &= "聯絡電話必須小於等於15字字數<BR>"
                End If
            End If

            colArray(cst_i行動電話) = TIMS.ClearSQM(colArray(cst_i行動電話))
            If colArray(cst_i行動電話).ToString <> "" Then
                If (colArray(cst_i行動電話).ToString.Length > 20) Then
                    Reason &= "行動電話必須小於等於20字字數<BR>"
                End If
            End If

            colArray(cst_i電子郵件) = TIMS.ClearSQM(colArray(cst_i電子郵件))
            If colArray(cst_i電子郵件).ToString <> "" Then
                If (colArray(cst_i電子郵件).ToString.Length > 64) Then
                    Reason &= "E_mail必須小於等於64字字數<BR>"
                End If
            End If

            colArray(cst_i郵遞區號前3碼) = TIMS.ClearSQM(colArray(cst_i郵遞區號前3碼))
            If colArray(cst_i郵遞區號前3碼).ToString = "" Then
                Reason &= "通訊地址郵遞區號前3碼必須填寫<BR>"
            Else
                If IsNumeric(colArray(cst_i郵遞區號前3碼)) = False Then
                    Reason &= "通訊地址郵遞區號前3碼必須為數字<BR>"
                Else
                    If Len(Convert.ToString(colArray(cst_i郵遞區號前3碼)).Trim) <> 3 Then
                        Reason &= "通訊地址郵遞區號前3碼必須為3碼<BR>"
                    End If
                End If
            End If

            colArray(cst_i郵遞區號後6碼) = TIMS.ClearSQM(colArray(cst_i郵遞區號後6碼))
            If colArray(cst_i郵遞區號後6碼).ToString = "" Then
                Reason &= "通訊地址郵遞區號後6碼必須填寫<BR>"
            Else
                If Not IsNumeric(colArray(cst_i郵遞區號後6碼)) Then
                    Reason &= "通訊地址郵遞區號後6碼必須為數字<BR>"
                Else
                    Dim ilen56 As Integer = Len(Convert.ToString(colArray(cst_i郵遞區號後6碼)))
                    If ilen56 <> 5 AndAlso ilen56 <> 6 Then Reason &= "通訊地址郵遞區號後6碼 長度必須為5碼或6碼<BR>"
                End If
            End If

            colArray(cst_i通訊地址) = TIMS.ClearSQM(colArray(cst_i通訊地址))
            If colArray(cst_i通訊地址).ToString = "" Then
                Reason &= "通訊地址必須填寫<BR>"
            Else
                If (colArray(cst_i通訊地址).ToString.Length > 50) Then
                    Reason &= "通訊地址 必須小於等於 50字字數<BR>"
                End If
            End If

            colArray(cst_i服務單位名稱) = TIMS.ClearSQM(colArray(cst_i服務單位名稱))
            If colArray(cst_i服務單位名稱).ToString = "" Then
                Reason &= "服務單位名稱必須填寫<BR>"
            Else
                If (colArray(cst_i服務單位名稱).ToString.Length > 50) Then
                    Reason &= "服務單位名稱 必須小於等於 50字字數<BR>"
                End If
            End If

            If colArray(cst_i年資).ToString <> "" Then
                If IsNumeric(colArray(cst_i年資)) = False Then
                    Reason &= "服務年資必須為數字<BR>"
                End If
            End If

            If colArray(cst_i服務部門).ToString <> "" Then
                colArray(cst_i服務部門) = Trim(colArray(cst_i服務部門)) '服務部門
                If (colArray(cst_i服務部門).ToString).Length > 50 Then
                    Reason &= "服務部門 過長應小於等於50字字數<BR>"
                End If
            End If

            colArray(cst_i服務單位電話) = TIMS.ClearSQM(colArray(cst_i服務單位電話))
            If colArray(cst_i服務單位電話).ToString = "" Then
                Reason &= "服務單位電話 必須填寫<BR>"
            Else
                If (colArray(cst_i服務單位電話).ToString.Length > 20) Then
                    Reason &= "服務單位電話 必須小於等於 20字字數<BR>"
                End If
            End If

            If colArray(cst_i服務單位傳真).ToString <> "" Then
                colArray(cst_i服務單位傳真) = Trim(colArray(cst_i服務單位傳真)) '服務單位傳真
                If (colArray(cst_i服務單位傳真).ToString).Length > 20 Then
                    Reason &= "服務單位傳真 過長應小於等於20字字數<BR>"
                End If
            End If

            If colArray(cst_i服務單位郵遞區號前3碼).ToString <> "" Then
                If IsNumeric(colArray(cst_i服務單位郵遞區號前3碼)) = False Then
                    Reason &= "服務單位郵遞區號前3碼必須為數字<BR>"
                Else
                    If Len(Convert.ToString(colArray(cst_i服務單位郵遞區號前3碼)).Trim) <> 3 Then
                        Reason &= "服務單位郵遞區號前3碼必須為3碼<BR>"
                    End If
                End If
            End If

            colArray(cst_i服務單位郵遞區號後6碼) = TIMS.ClearSQM(colArray(cst_i服務單位郵遞區號後6碼))
            If colArray(cst_i服務單位郵遞區號後6碼).ToString <> "" Then
                'Reason &= "服務單位郵遞區號後6碼必須填寫<BR>"
                If Not IsNumeric(colArray(cst_i服務單位郵遞區號後6碼)) Then
                    Reason &= "服務單位郵遞區號後6碼必須為數字<BR>"
                Else
                    Dim ilen56 As Integer = Len(Convert.ToString(colArray(cst_i服務單位郵遞區號後6碼)))
                    If ilen56 <> 5 AndAlso ilen56 <> 6 Then Reason &= "服務單位郵遞區號後6碼 長度必須為5碼或6碼<BR>"
                End If
            End If

            If colArray(cst_i服務單位地址).ToString <> "" Then
                colArray(cst_i服務單位地址) = Trim(colArray(cst_i服務單位地址)) '服務單位地址
                If (colArray(cst_i服務單位地址).ToString).Length > 50 Then
                    Reason &= "服務單位地址 過長應小於等於50字字數<BR>"
                End If
            End If

            'If colArray(30).ToString = "" Then
            '    Reason &= "服務單位一必須填寫<BR>"
            'End If

            If colArray(cst_i服務單位一).ToString <> "" Then
                colArray(cst_i服務單位一) = Trim(colArray(cst_i服務單位一))
                If (colArray(cst_i服務單位一).ToString.Length > 50) Then
                    Reason &= "服務單位一 必須小於等於 50字字數<BR>"
                End If
            End If

            If colArray(cst_i服務單位二).ToString <> "" Then
                colArray(cst_i服務單位二) = Trim(colArray(cst_i服務單位二))
                If (colArray(cst_i服務單位二).ToString.Length > 50) Then
                    Reason &= "服務單位二 必須小於等於 50字字數<BR>"
                End If
            End If

            If colArray(cst_i服務單位三).ToString <> "" Then
                colArray(cst_i服務單位三) = Trim(colArray(cst_i服務單位三))
                If (colArray(cst_i服務單位三).ToString.Length > 50) Then
                    Reason &= "服務單位三 必須小於等於 50字字數<BR>"
                End If
            End If

            'If colArray(33).ToString = "" Then
            '    Reason &= "服務年資一必須填寫<BR>"
            'End If

            If colArray(cst_i服務年資一).ToString <> "" Then
                If Not IsNumeric(colArray(cst_i服務年資一).ToString) Then
                    Reason &= "服務年資一必須填寫整數數字格式<BR>"
                Else
                    If colArray(cst_i服務年資一).ToString.Trim.IndexOf(".") > -1 Then
                        Reason &= "服務年資一必須填寫整數數字格式<BR>"
                    End If
                    colArray(cst_i服務年資一) = CInt(colArray(cst_i服務年資一))
                End If
            End If

            If colArray(cst_i服務年資二).ToString <> "" Then
                If Not IsNumeric(colArray(cst_i服務年資二).ToString) Then
                    Reason &= "服務年資二必須填寫整數數字格式<BR>"
                Else
                    If colArray(cst_i服務年資二).ToString.Trim.IndexOf(".") > -1 Then
                        Reason &= "服務年資二必須填寫整數數字格式<BR>"
                    End If
                    colArray(cst_i服務年資二) = CInt(colArray(cst_i服務年資二))
                End If
            End If

            If colArray(cst_i服務年資三).ToString <> "" Then
                If Not IsNumeric(colArray(cst_i服務年資三).ToString) Then
                    Reason &= "服務年資三必須填寫整數數字格式<BR>"
                Else
                    If colArray(cst_i服務年資三).ToString.Trim.IndexOf(".") > -1 Then
                        Reason &= "服務年資三必須填寫整數數字格式<BR>"
                    End If
                    colArray(cst_i服務年資三) = CInt(colArray(cst_i服務年資三))
                End If
            End If

            '36~41
            For intCol As Integer = cst_i服務期間一起日 To cst_i服務期間三迄日
                'intCol = ji 'cst_i服務期間一起日
                Select Case intCol
                    Case cst_i服務期間一起日
                        strCol1 = "服務期間一起日"
                    Case cst_i服務期間一迄日
                        strCol1 = "服務期間一迄日"
                    Case cst_i服務期間二起日
                        strCol1 = "服務期間二起日"
                    Case cst_i服務期間二迄日
                        strCol1 = "服務期間二迄日"
                    Case cst_i服務期間三起日
                        strCol1 = "服務期間三起日"
                    Case cst_i服務期間三迄日
                        strCol1 = "服務期間三迄日"
                End Select
                If colArray(intCol).ToString = "" Then
                    'Reason += strCol1 & " 必須填寫<Br>"
                Else
                    If IsDate(colArray(intCol)) = False Then
                        Reason &= $"{strCol1} 必須是西元年格式(yyyy/mm/dd)<BR>"
                    Else
                        Try
                            colArray(intCol) = CDate(colArray(intCol)).ToString("yyyy/MM/dd")
                            If CDate(colArray(intCol)) < "1900/1/1" Or CDate(colArray(intCol)) > "2100/1/1" Then
                                Reason &= $"{strCol1} 範圍有誤<BR>"
                            End If
                        Catch ex As Exception
                            Reason &= $"{strCol1} 必須是西元年格式(yyyy/mm/dd)<BR>"
                        End Try
                    End If
                End If
            Next

            '42~46
            For intCol As Integer = cst_i專長一 To cst_i專長五
                Select Case intCol
                    Case cst_i專長一
                        strCol1 = "專長一"
                    Case cst_i專長二
                        strCol1 = "專長二"
                    Case cst_i專長三
                        strCol1 = "專長三"
                    Case cst_i專長四
                        strCol1 = "專長四"
                    Case cst_i專長五
                        strCol1 = "專長五"
                End Select

                If colArray(intCol).ToString <> "" Then
                    colArray(intCol) = Trim(colArray(intCol))
                    If colArray(intCol).ToString.Length > 250 Then
                        Reason &= $"{strCol1} 長度過長，限制為250個字元<BR>"
                    End If
                End If
            Next

            If colArray(cst_i譯著).ToString <> "" Then
                colArray(cst_i譯著) = Trim(colArray(cst_i譯著))
                If colArray(cst_i譯著).ToString.Length > 100 Then
                    Reason &= "譯著 長度過長，限制為100個字元<BR>"
                End If
            End If

            If colArray(cst_i專業證照).ToString <> "" Then
                colArray(cst_i專業證照) = Trim(colArray(cst_i專業證照))
                If colArray(cst_i專業證照).ToString.Length > 100 Then
                    Reason &= "專業證照長度過長，限制為100個字元<BR>"
                End If
            End If

            If colArray(cst_i排課使用).ToString = "" Then
                Reason &= "排課使用必須填寫<BR>"
            Else
                Select Case colArray(cst_i排課使用).ToString
                    Case "1", "2"
                    Case Else
                        Reason &= "排課使用必須輸入1或2(1.是,2.否)<BR>"
                End Select
            End If

            If TIMS.Cst_TPlanID06Plan1.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                '2018 add 在職教師類別(教師 / 第二教師)
                colArray(cst_i教師類別) = TIMS.ClearSQM(Convert.ToString(colArray(cst_i教師類別)))
                If colArray(cst_i教師類別).ToString <> "" Then
                    colArray(cst_i教師類別) = UCase(colArray(cst_i教師類別))
                    Select Case colArray(cst_i教師類別).ToString
                        Case "Y"
                        Case Else
                            Reason &= "教師類別 只能為(Y:是教師)或不填<BR>"
                    End Select
                End If
                colArray(cst_i第二教師類別) = TIMS.ClearSQM(Convert.ToString(colArray(cst_i第二教師類別)))
                If colArray(cst_i第二教師類別).ToString <> "" Then
                    colArray(cst_i第二教師類別) = UCase(colArray(cst_i第二教師類別))
                    Select Case colArray(cst_i第二教師類別).ToString
                        Case "Y"
                        Case Else
                            Reason &= "第二教師 只能為(Y:是第二教師)或不填<BR>"
                    End Select
                End If
            Else
                colArray(cst_i講師類別) = TIMS.ClearSQM(Convert.ToString(colArray(cst_i講師類別)))
                If colArray(cst_i講師類別).ToString <> "" Then
                    colArray(cst_i講師類別) = UCase(colArray(cst_i講師類別))
                    Select Case colArray(cst_i講師類別).ToString
                        Case "Y"
                        Case Else
                            Reason &= "講師類別 只能為(Y:是講師)或不填<BR>"
                    End Select
                End If
                colArray(cst_i助教類別) = TIMS.ClearSQM(Convert.ToString(colArray(cst_i助教類別)))
                If colArray(cst_i助教類別).ToString <> "" Then
                    colArray(cst_i助教類別) = UCase(colArray(cst_i助教類別))
                    Select Case colArray(cst_i助教類別).ToString
                        Case "Y"
                        Case Else
                            Reason &= "助教類別 只能為(Y:是助教)或不填<BR>"
                    End Select
                End If
            End If

            'If colArray.Length > 47 Then
            '    'If Not colArray(47).ToString = "" Then
            '    '    colArray(47) = Trim(colArray(47))
            '    '    If IsNumeric(colArray(47).ToString) Then
            '    '        Select Case colArray(47).ToString
            '    '            Case 1, 2
            '    '            Case Else
            '    '                Reason &= "身份別有誤(1:本國 2:外籍)<BR>"
            '    '        End Select
            '    '    End If
            '    'End If
            'End If

        End If
        Return Reason
    End Function

    '查詢原因-INQUIRY
    Private Function GET_SEARCH_MEMO() As String
        Dim RstMemo As String = ""
        center.Value = TIMS.ClearSQM(center.Value)
        TextBox2.Text = TIMS.ClearSQM(TextBox2.Text)
        TextBox3.Text = TIMS.ClearSQM(TextBox3.Text)
        TextBox4.Text = TIMS.ClearSQM(TextBox4.Text)
        TB_career_id.Text = TIMS.ClearSQM(TB_career_id.Text)
        Dim v_DropDownList1 As String = TIMS.GetListValue(DropDownList1)
        Dim v_DropDownList2 As String = TIMS.GetListValue(DropDownList2)
        Dim v_DropDownList3 As String = TIMS.GetListValue(DropDownList3)
        Dim v_DropDownList4 As String = TIMS.GetListValue(DropDownList4)

        If center.Value <> "" Then RstMemo &= String.Concat("&訓練機構=", center.Value)
        If TextBox2.Text <> "" Then RstMemo &= String.Concat("&講師姓名=", TextBox2.Text)
        If TextBox3.Text <> "" Then RstMemo &= String.Concat("&身分證號碼=", TextBox3.Text)
        If v_DropDownList4 <> "" Then RstMemo &= String.Concat("&內外聘=", v_DropDownList4)
        If v_DropDownList1 <> "" Then RstMemo &= String.Concat("&師資別=", v_DropDownList1)
        If TextBox4.Text <> "" Then RstMemo &= String.Concat("&講師代碼=", TextBox4.Text)
        If v_DropDownList2 <> "" Then RstMemo &= String.Concat("&排課使用=", v_DropDownList2)
        If TB_career_id.Text <> "" Then RstMemo &= String.Concat("&主要職類=", TB_career_id.Text)
        If v_DropDownList3 <> "" Then RstMemo &= String.Concat("&職稱=", v_DropDownList3)
        RstMemo &= String.Concat("&cb_techtype1=", cb_techtype1.Checked)
        RstMemo &= String.Concat("&cb_techtype2=", cb_techtype2.Checked)
        RstMemo &= String.Concat("&cb_techtype3=", cb_techtype3.Checked)
        RstMemo &= String.Concat("&cb_techtype4=", cb_techtype4.Checked)
        Return RstMemo
    End Function

    '查詢
    Sub SEARCH()
        Dim SearchStr As String = ""
        Dim parms As Hashtable = New Hashtable()

        If ViewState("KindID") <> "" Then
            SearchStr &= " AND KindID=@KindID" & vbCrLf
            parms.Add("KindID", ViewState("KindID"))
        End If
        If ViewState("WorkStatus") <> "" Then
            SearchStr &= " AND WorkStatus=@WorkStatus" & vbCrLf
            parms.Add("WorkStatus", ViewState("WorkStatus"))
        End If
        If ViewState("IVID") <> "" Then
            SearchStr &= " AND IVID=@IVID" & vbCrLf
            parms.Add("IVID", ViewState("IVID"))
        End If
        If ViewState("KindEngage") <> "" Then
            SearchStr &= " AND KindEngage=@KindEngage" & vbCrLf
            parms.Add("KindEngage", ViewState("KindEngage"))
        End If

        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '產投
            'Me.LabTMID.Text = "訓練業別"
            If ViewState("jobValue") <> "" Then
                SearchStr &= " AND ( TMID=@TMID OR TMID IN (" & vbCrLf
                SearchStr &= " select TMID from Key_TrainType where parent IN (" & vbCrLf '職類別
                SearchStr &= " select TMID from Key_TrainType where parent IN (" & vbCrLf '業別
                SearchStr &= " select TMID from Key_TrainType where busid ='G')" & vbCrLf '產業人才投資方案類
                SearchStr &= " AND TMID=@TMID )))" & vbCrLf

                parms.Add("TMID", ViewState("jobValue"))
            Else
                If ViewState("TMID") <> "" Then
                    SearchStr &= " AND TMID=@TMID" & vbCrLf
                    parms.Add("TMID", ViewState("TMID"))
                End If
            End If

            '產投
            'tr_techtype12.Visible = True '顯示 '講師 助教 (類別) 
            If ViewState("TechType1") <> "" OrElse ViewState("TechType2") <> "" Then
                If ViewState("TechType1") <> "" Then '講師
                    'SearchStr &= " AND TechType1=@TechType1" & vbCrLf
                    SearchStr &= " TechType1=@TechType1" & vbCrLf
                    parms.Add("TechType1", ViewState("TechType1"))
                End If
                If ViewState("TechType2") <> "" Then '助教
                    'SearchStr &= " AND TechType2=@TechType2" & vbCrLf
                    If ViewState("TechType1") <> "" Then SearchStr &= " or "
                    SearchStr &= " TechType2=@TechType2" & vbCrLf
                    parms.Add("TechType2", ViewState("TechType2"))
                End If
            End If

        Else
            If ViewState("TMID") <> "" Then
                SearchStr &= " AND TMID=@TMID" & vbCrLf
                parms.Add("TMID", ViewState("TMID"))
            End If
        End If

        If TIMS.Cst_TPlanID06Plan1.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            If ViewState("TechType3") <> "" OrElse ViewState("TechType4") <> "" Then
                SearchStr &= " and ( "
                '2018 add 自辦在職師資類別
                If ViewState("TechType3") <> "" Then '教師
                    SearchStr &= " TechType3=@TechType3" & vbCrLf
                    parms.Add("TechType3", ViewState("TechType3"))
                End If
                If ViewState("TechType4") <> "" Then '第二教師
                    If ViewState("TechType3") <> "" Then SearchStr &= " or "
                    SearchStr &= " TechType4=@TechType4" & vbCrLf
                    parms.Add("TechType4", ViewState("TechType4"))
                End If
                SearchStr &= " ) "
            End If
        End If

        If ViewState("RID") <> "" Then
            SearchStr &= " AND RID=@RID" & vbCrLf
            parms.Add("RID", ViewState("RID"))
        End If

        If ViewState("TeachCName") <> "" Then
            SearchStr &= " AND TeachCName like '%' + @TeachCName + '%'" & vbCrLf 'fix ORA-01722: invalid number
            parms.Add("TeachCName", ViewState("TeachCName"))
        End If
        If ViewState("IDNO") <> "" Then
            SearchStr &= " AND IDNO like '%' + @IDNO + '%'" & vbCrLf 'fix ORA-01722: invalid number
            parms.Add("IDNO", ViewState("IDNO"))
        End If
        If ViewState("TeacherID") <> "" Then
            SearchStr &= " AND TeacherID like '%' + @TeacherID + '%'" & vbCrLf
            parms.Add("TeacherID", ViewState("TeacherID"))
        End If

        Dim sql As String = $"SELECT * FROM TEACH_TEACHERINFO WHERE 1=1 {SearchStr}"

        msg.Text = "查無資料!!"
        DataGridTable.Visible = False

        Dim dt As DataTable = Nothing
        Try
            dt = DbAccess.GetDataTable(sql, objconn, parms)
        Catch ex As Exception
            Common.MessageBox(Me, $"查詢時發生錯誤，請重新輸入查詢值!!{ex.Message}")
            'Common.MessageBox(Me, ex.ToString)
            Dim strErrmsg As String = $"/* ex.Str: */{ex.ToString()}{vbCrLf}/* sql: */{sql}{vbCrLf}{TIMS.GetErrorMsg(Me)}" '取得錯誤資訊寫入
            strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)
            'Throw ex
            Exit Sub
        End Try

        '查詢原因
        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        Dim v_rblWorkMode As String = TIMS.GetListValue(rblWorkMode)
        Session(TIMS.gcst_rblWorkMode) = v_rblWorkMode
        Dim MRqID As String = TIMS.Get_MRqID(Me)
        sMemo = GET_SEARCH_MEMO()
        '查詢原因:INQNO--查詢結果筆數:RESCNT--查詢清單結果:RESDESC 
        Dim vRESDESC As String = TIMS.GET_RESDESCdt(dt, "TEACHERID,TEACHCNAME,IDNO,KINDID,KINDENGAGE")
        Call TIMS.SubInsAccountLog1(Me, MRqID, TIMS.cst_wm查詢, v_rblWorkMode, "", sMemo, objconn, v_INQUIRY, dt.Rows.Count, vRESDESC)

        If TIMS.dtHaveDATA(dt) Then
            For Each dr As DataRow In dt.Rows
                Dim idno As String = TIMS.ChangeIDNO($"{dr("IDNO")}")
                If v_rblWorkMode = TIMS.cst_wmdip1 Then dr("IDNO") = TIMS.strMask(idno, 1) '(身份證號MASK)
            Next

            msg.Text = ""
            DataGridTable.Visible = True

            PageControler1.PageDataTable = dt
            PageControler1.PrimaryKey = "TechID"
            PageControler1.Sort = "TeacherID"
            PageControler1.ControlerLoad()
        End If
    End Sub

    ''' <summary>
    ''' 查詢SQL
    ''' </summary>
    Sub GClickSearchButton()
        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        Dim SearchStr As String = "" & vbCrLf

        ViewState("KindID") = ""
        ViewState("WorkStatus") = ""
        ViewState("IVID") = ""
        ViewState("KindEngage") = ""

        ViewState("jobValue") = ""
        ViewState("TMID") = ""
        ViewState("RID") = ""
        ViewState("TeachCName") = ""
        ViewState("IDNO") = ""
        ViewState("TeacherID") = ""
        ViewState("TechType1") = ""
        ViewState("TechType2") = ""
        ViewState("TechType3") = ""
        ViewState("TechType4") = ""

        If tr_techtype12.Visible Then
            '產投(顯示) 才存取此功能
            'tr_techtype12.Visible = True '顯示 '講師 助教 (類別) 
            If cb_techtype1.Checked Then
                ViewState("TechType1") = "Y"
            End If
            If cb_techtype2.Checked Then
                ViewState("TechType2") = "Y"
            End If
        End If

        If tr_techtype34.Visible Then
            '2018 add 自辦在職(顯示) 才有此資訊可存取
            If cb_techtype3.Checked Then
                ViewState("TechType3") = "Y"
            End If

            If cb_techtype4.Checked Then
                ViewState("TechType4") = "Y"
            End If
        End If

        If DropDownList1.SelectedIndex <> 0 AndAlso DropDownList1.SelectedValue <> "" Then
            ViewState("KindID") = TIMS.ClearSQM(DropDownList1.SelectedValue)
        End If
        If DropDownList2.SelectedIndex <> 0 AndAlso DropDownList2.SelectedValue <> "" Then
            ViewState("WorkStatus") = TIMS.ClearSQM(DropDownList2.SelectedValue)
        End If
        If DropDownList3.SelectedIndex <> 0 AndAlso DropDownList3.SelectedValue <> "" Then
            ViewState("IVID") = TIMS.ClearSQM(DropDownList3.SelectedValue)
        End If
        If DropDownList4.SelectedIndex <> 0 AndAlso DropDownList4.SelectedValue <> "" Then
            ViewState("KindEngage") = TIMS.ClearSQM(DropDownList4.SelectedValue)
        End If

        If jobValue.Value <> "" Then
            ViewState("jobValue") = TIMS.ClearSQM(jobValue.Value)
        End If
        If trainValue.Value <> "" Then
            ViewState("TMID") = TIMS.ClearSQM(trainValue.Value)
        End If
        If RIDValue.Value.Trim = "" Then
            ViewState("RID") = sm.UserInfo.RID
        Else
            ViewState("RID") = TIMS.ClearSQM(RIDValue.Value)
        End If
        If TextBox2.Text.Trim <> "" Then
            ViewState("TeachCName") = TIMS.ClearSQM(TextBox2.Text)
        End If
        If TextBox3.Text.Trim <> "" Then
            ViewState("IDNO") = TIMS.ClearSQM(TextBox3.Text)
        End If
        If TextBox4.Text.Trim <> "" Then
            ViewState("TeacherID") = TIMS.ClearSQM(TextBox4.Text)
        End If

        '--------------- 查詢開始 '--------------- 
        Call SEARCH()
    End Sub

    '新增
    Private Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
        TextBox2.Text = TIMS.ClearSQM(TextBox2.Text)
        TextBox3.Text = TIMS.ClearSQM(TextBox3.Text)

        Call GetSearchStr()
        'Response.Redirect("TC_01_007_add.aspx?proecess=add&ID=" & MRqID)
        '20100208 按新增時代查詢之 講師名稱 & 身分證號
        'Response.Redirect("TC_01_007_add.aspx?proecess=Insert&ID=" & MRqID & "&TeachCName=" & TextBox2.Text & "&TeachIDNO=" & TextBox3.Text)
        Dim MRqID As String = TIMS.Get_MRqID(Me)
        Dim url1 As String = "TC_01_007_add.aspx?proecess=Insert&ID=" & MRqID & "&TeachCName=" & TextBox2.Text & "&TeachIDNO=" & TextBox3.Text
        TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    '列印排課匯入用的講師代碼
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, "RID=" & RIDValue.Value)
    End Sub

    '匯入名冊
    Private Sub Btn_XlsImport_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Btn_XlsImport.Click
        Dim dt_xls As DataTable = Nothing
        Dim MyFileName As String = ""
        Dim MyFileType As String = ""
        Dim Reason As String = "" '儲存錯誤的原因
        Dim dtWrong As New DataTable '儲存錯誤資料的DataTable
        Dim iRowIndex As Integer = 1

        Const cst_FileType As String = "xls"
        '檢核檔案狀況 異常為False (有錯誤訊息視窗)
        Dim MyPostedFile As HttpPostedFile = Nothing
        If Not TIMS.HttpCHKFile(Me, File2, MyPostedFile, cst_FileType) Then Return

        Const Cst_FileSavePath As String = "~/TC/01/Temp/"
        Call TIMS.MyCreateDir(Me, Cst_FileSavePath)
        If File2.Value = "" Then
            Common.MessageBox(Me, "未輸入匯入檔案位置")
            Exit Sub
        End If
        '檢查檔案格式
        If File2.PostedFile.ContentLength = 0 Then
            Common.MessageBox(Me, "檔案位置錯誤!")
            Exit Sub
        End If
        '取出檔案名稱
        MyFileName = Split(File2.PostedFile.FileName, "\")((Split(File2.PostedFile.FileName, "\")).Length - 1)
        '取出檔案類型
        If MyFileName.IndexOf(".") = -1 Then
            Common.MessageBox(Me, "檔案類型錯誤!")
            Exit Sub
        End If
        MyFileType = Split(MyFileName, ".")((Split(MyFileName, ".")).Length - 1)
        If Not MyFileType = cst_FileType Then
            Common.MessageBox(Me, "檔案類型錯誤，必須為XLS檔!")
            Exit Sub
        End If

        '3. 使用 Path.GetExtension 取得副檔名 (包含點，例如 .jpg)
        Dim fileNM_Ext As String = System.IO.Path.GetExtension(File2.PostedFile.FileName).ToLower()
        MyFileName = $"f{TIMS.GetDateNo()}_{TIMS.GetRandomS4()}{fileNM_Ext}"
        '上傳檔案
        File2.PostedFile.SaveAs(Server.MapPath(Cst_FileSavePath & MyFileName))

        '取得內容
        dt_xls = TIMS.GetDataTable_XlsFile(
                          Server.MapPath(Cst_FileSavePath & MyFileName).ToString,
                        "", Reason, "計劃階層", "講師代碼", "講師姓名")

        IO.File.Delete(Server.MapPath(Cst_FileSavePath & MyFileName)) '刪除檔案

        If Reason <> "" Then
            Common.MessageBox(Me, Reason)
            Common.MessageBox(Me, "資料有誤，故無法匯入，請修正Excel檔案，謝謝")
            Exit Sub
        End If


        'xls 方式 讀取寫入資料庫
        If dt_xls.Rows.Count > 0 Then '有資料
            '建立錯誤資料格式Table
            dtWrong.Columns.Add(New DataColumn("Index"))
            dtWrong.Columns.Add(New DataColumn("TeacherID"))
            dtWrong.Columns.Add(New DataColumn("Name"))
            dtWrong.Columns.Add(New DataColumn("IDNO"))
            dtWrong.Columns.Add(New DataColumn("Reason"))
            Reason = ""
            For i As Integer = 0 To dt_xls.Rows.Count - 1
                If iRowIndex <> 0 Then
                    Dim colArray As Array = dt_xls.Rows(i).ItemArray
                    Reason = CheckImportData(colArray)
                    If Reason <> "" Then
                        '錯誤資料，填入錯誤資料表
                        Dim drWrong As DataRow = dtWrong.NewRow
                        dtWrong.Rows.Add(drWrong)

                        drWrong("Index") = iRowIndex
                        If colArray.Length > 5 Then
                            drWrong("TeacherID") = colArray(cst_i講師代碼)
                            drWrong("Name") = colArray(cst_i講師姓名)
                            drWrong("IDNO") = colArray(cst_i身分證字號)
                            drWrong("Reason") = Reason
                        End If
                    Else '匯入資料
                        'Dim sql As String = ""
                        Dim dr As DataRow = Nothing
                        Dim dt As DataTable = Nothing
                        Dim da As SqlDataAdapter = Nothing

                        Dim s_TransType As String = TIMS.cst_TRANS_LOG_Update
                        Dim s_TargetTable As String = "TEACH_TEACHERINFO"
                        Dim s_FuncPath As String = "/TC/01/TC_01_007"
                        Const cst_fWHERE As String = "TECHID={0}"
                        Dim s_WHERE As String = ""

                        Using tConn As SqlConnection = DbAccess.GetConnection()
                            Dim trans As SqlTransaction = DbAccess.BeginTrans(tConn)
                            Try
                                Dim sql As String = ""
                                sql &= " SELECT * FROM TEACH_TEACHERINFO"
                                sql &= $" WHERE RID='{TIMS.ClearSQM(RIDValue.Value)}' AND TeacherID='{TIMS.ClearSQM(colArray(cst_i講師代碼))}' AND IDNO='{TIMS.ClearSQM(colArray(cst_i身分證字號))}'"
                                dt = DbAccess.GetDataTable(sql, da, trans)
                                Dim iTECHID As Integer = 0
                                If dt.Rows.Count = 0 Then
                                    s_TransType = TIMS.cst_TRANS_LOG_Insert
                                    iTECHID = DbAccess.GetNewId(trans, "TEACH_TEACHERINFO_TECHID_SEQ,TEACH_TEACHERINFO,TECHID")
                                    dr = dt.NewRow()
                                    dt.Rows.Add(dr)
                                    dr("TECHID") = iTECHID 'TEACH_TEACHERINFO_TECHID_SEQ
                                    dr("RID") = RIDValue.Value '機構
                                Else
                                    dr = dt.Rows(0)
                                    iTECHID = dr("TECHID")
                                End If
                                s_WHERE = String.Format(cst_fWHERE, iTECHID)
                                dr("TeacherID") = colArray(cst_i講師代碼).ToString '講師代碼
                                dr("TeachCName") = colArray(cst_i講師姓名).ToString '講師姓名
                                dr("TeachEName") = If(colArray(cst_i講師英文姓名).ToString <> "", colArray(cst_i講師英文姓名).ToString, Convert.DBNull) '講師英文姓名
                                '身份別 1/2
                                dr("PassPortNO") = If(colArray(cst_i身份別).ToString = "1", colArray(cst_i身份別), If(colArray(cst_i身份別).ToString = "2", colArray(cst_i身份別), "2"))
                                dr("IDNO") = colArray(cst_i身分證字號).ToString '身分證號碼
                                '出生日期(可輸入空白)
                                dr("Birthday") = If(colArray(cst_i出生日期).ToString() <> "", CDate(colArray(cst_i出生日期)), Convert.DBNull)
                                '性別
                                dr("Sex") = If(colArray(cst_i性別).ToString = "M", colArray(cst_i性別), If(colArray(cst_i性別).ToString = "F", colArray(cst_i性別), Convert.DBNull))
                                dr("TMID") = colArray(cst_i主要職類).ToString '職類代碼

                                If RIDValue.Value.Length = 1 Then
                                    'SELECT distinct IVID FROM TEACH_TEACHERINFO where 1=1 order by 1
                                    dr("IVID") = If(colArray(cst_i職稱).ToString <> "", colArray(cst_i職稱).ToString, Convert.DBNull) '職稱代碼
                                Else
                                    'SELECT distinct INVEST,trim(INVEST) FROM TEACH_TEACHERINFO where 1=1 and INVEST!=trim(INVEST) order by 1
                                    'update TEACH_TEACHERINFO set INVEST=trim(INVEST) where 1=1 and INVEST!=trim(INVEST)
                                    'SELECT distinct INVEST FROM TEACH_TEACHERINFO where 1=1 order by 1
                                    dr("INVEST") = If(colArray(cst_i職稱).ToString <> "", colArray(cst_i職稱).ToString, Convert.DBNull) '職稱代碼
                                End If
                                dr("KindEngage") = colArray(cst_i內外聘).ToString '內外聘
                                dr("KindID") = colArray(cst_i師資別).ToString '師資別
                                dr("DegreeID") = colArray(cst_i最高學歷).ToString '最高學歷
                                dr("GraduateStatus") = colArray(cst_i畢業狀況).ToString '畢業狀況
                                dr("SchoolName") = If(colArray(cst_i學校名稱).ToString <> "", colArray(cst_i學校名稱).ToString, Convert.DBNull) '學校名稱
                                dr("Department") = If(colArray(cst_i科系名稱).ToString <> "", colArray(cst_i科系名稱).ToString, Convert.DBNull) '科系名稱
                                dr("Phone") = colArray(cst_i聯絡電話).ToString '聯絡電話
                                dr("Mobile") = If(colArray(cst_i行動電話).ToString <> "", colArray(cst_i行動電話).ToString, Convert.DBNull) '行動電話
                                dr("Email") = If(colArray(cst_i電子郵件).ToString <> "", colArray(cst_i電子郵件).ToString, Convert.DBNull) 'E_Mail

                                dr("AddressZip") = colArray(cst_i郵遞區號前3碼).ToString '通訊地址Zip
                                dr("AddressZIP6W") = colArray(cst_i郵遞區號後6碼).ToString()
                                dr("Address") = colArray(cst_i通訊地址).ToString '通訊地址

                                dr("WorkOrg") = colArray(cst_i服務單位名稱).ToString '服務單位名稱
                                If Convert.ToString(colArray(cst_i年資)) <> "" Then
                                    dr("ExpYears") = colArray(cst_i年資).ToString '服務年資
                                End If
                                If colArray(cst_i服務部門).ToString <> "" Then
                                    dr("ServDept") = colArray(cst_i服務部門).ToString '服務部門
                                End If
                                dr("WorkPhone") = colArray(cst_i服務單位電話).ToString '服務單位電話
                                If colArray(cst_i服務單位傳真).ToString <> "" Then
                                    dr("Fax") = colArray(cst_i服務單位傳真).ToString '服務單位傳真
                                End If
                                If colArray(cst_i服務單位郵遞區號前3碼).ToString <> "" Then
                                    dr("WorkZip") = colArray(cst_i服務單位郵遞區號前3碼) '服務單位地址Zip
                                End If
                                If colArray(cst_i服務單位郵遞區號後6碼).ToString() <> "" Then
                                    dr("WorkZIP6W") = colArray(cst_i服務單位郵遞區號後6碼).ToString()
                                End If
                                If colArray(cst_i服務單位地址).ToString <> "" Then
                                    dr("Workaddr") = colArray(cst_i服務單位地址).ToString '服務單位地址
                                End If

                                If colArray(cst_i服務單位一).ToString <> "" Then
                                    dr("ExpUnit1") = colArray(cst_i服務單位一).ToString '服務單位一
                                End If
                                If colArray(cst_i服務單位二).ToString <> "" Then
                                    dr("ExpUnit2") = colArray(cst_i服務單位二).ToString '服務單位二
                                End If
                                If colArray(cst_i服務單位三).ToString <> "" Then
                                    dr("ExpUnit3") = colArray(cst_i服務單位三).ToString '服務單位三
                                End If

                                If colArray(cst_i服務年資一).ToString <> "" Then
                                    dr("ExpYears1") = colArray(cst_i服務年資一).ToString '服務年資一
                                End If
                                If colArray(cst_i服務年資二).ToString <> "" Then
                                    dr("ExpYears2") = colArray(cst_i服務年資二).ToString '服務年資二
                                End If
                                If colArray(cst_i服務年資三).ToString <> "" Then
                                    dr("ExpYears3") = colArray(cst_i服務年資三).ToString '服務年資三
                                End If

                                If colArray(cst_i服務期間一起日).ToString <> "" Then
                                    dr("ExpSDate1") = colArray(cst_i服務期間一起日).ToString '服務單位一起日
                                End If
                                If colArray(cst_i服務期間一迄日).ToString <> "" Then
                                    dr("ExpEDate1") = colArray(cst_i服務期間一迄日).ToString '服務單位一迄日
                                End If
                                If colArray(cst_i服務期間二起日).ToString <> "" Then
                                    dr("ExpSDate2") = colArray(cst_i服務期間二起日).ToString '服務單位二起日
                                End If
                                If colArray(cst_i服務期間二迄日).ToString <> "" Then
                                    dr("ExpEDate2") = colArray(cst_i服務期間二迄日).ToString '服務單位二迄日
                                End If
                                If colArray(cst_i服務期間三起日).ToString <> "" Then
                                    dr("ExpSDate3") = colArray(cst_i服務期間三起日).ToString '服務單位三起日
                                End If
                                If colArray(cst_i服務期間三迄日).ToString <> "" Then
                                    dr("ExpEDate3") = colArray(cst_i服務期間三迄日).ToString '服務單位三迄日
                                End If

                                Dim xi As Integer = 0
                                xi = 0
                                For ji As Integer = cst_i專長一 To cst_i專長五
                                    xi += 1
                                    Dim columnName As String = "Specialty" & CStr(xi)
                                    dr(columnName) = colArray(ji).ToString '專長一~專長五(42~46)
                                Next

                                If colArray(cst_i譯著).ToString <> "" Then
                                    dr("TransBook") = colArray(cst_i譯著).ToString '譯著
                                End If
                                If colArray(cst_i專業證照).ToString <> "" Then
                                    dr("ProLicense") = colArray(cst_i專業證照).ToString '專業證照
                                End If
                                If colArray(cst_i排課使用).ToString <> "" Then
                                    dr("WorkStatus") = colArray(cst_i排課使用).ToString '排課使用
                                End If

                                If TIMS.Cst_TPlanID06Plan1.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                                    '2018 add for 在職進修教師類別
                                    If colArray(cst_i教師類別).ToString <> "" Then
                                        dr("TECHTYPE3") = colArray(cst_i教師類別).ToString '講師類別
                                    End If
                                    If colArray(cst_i第二教師類別).ToString <> "" Then
                                        dr("TECHTYPE4") = colArray(cst_i第二教師類別).ToString '助教類別
                                    End If
                                Else
                                    If colArray(cst_i講師類別).ToString <> "" Then
                                        dr("TECHTYPE1") = colArray(cst_i講師類別).ToString '講師類別
                                    End If
                                    If colArray(cst_i助教類別).ToString <> "" Then
                                        dr("TECHTYPE2") = colArray(cst_i助教類別).ToString '助教類別
                                    End If
                                End If
                                dr("ModifyAcct") = sm.UserInfo.UserID '異動者
                                dr("ModifyDate") = Now() '異動時間

                                'Dim iPassPortNO As Integer = 1
                                'If UBound(colArray) >= 47 Then
                                '    If IsNumeric(colArray(47).ToString) And colArray(47).ToString <> "" Then
                                '        iPassPortNO = colArray(47).ToString '身份別 1:本國2:外籍
                                '    End If
                                'End If
                                'Select Case CStr(iPassPortNO)
                                '    Case "1", "2"
                                '        dr("PassPortNO") = iPassPortNO
                                '    Case Else
                                '        dr("PassPortNO") = "2"
                                'End Select
                                'Dim s_TransType, s_TargetTable, s_FuncPath, s_WHERE As String
                                Dim htPP As New Hashtable From {
                                    {"TransType", s_TransType},
                                    {"TargetTable", s_TargetTable},
                                    {"FuncPath", s_FuncPath},
                                    {"s_WHERE", s_WHERE}
                                }
                                TIMS.SaveTRANSLOG(sm, tConn, trans, dr, htPP)

                                DbAccess.UpdateDataTable(dt, da, trans)
                                DbAccess.CommitTrans(trans)

                            Catch ex As Exception
                                DbAccess.RollbackTrans(trans)
                                TIMS.CloseDbConn(tConn)

                                Const cst_errmsg1 As String = "意外錯誤：(請提供詳細資料，並連絡系統管理者協助處理)"

                                Dim strErrmsg As String = ""
                                strErrmsg &= "/*  匯入名冊 TC_01_007. Private Sub Btn_XlsImport_Click(ByVal sender As Object */" & vbCrLf
                                strErrmsg &= "/*  ex.ToString: */" & vbCrLf
                                strErrmsg += ex.ToString & vbCrLf
                                strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
                                strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                                Call TIMS.WriteTraceLog(strErrmsg)

                                Common.MessageBox(Me, cst_errmsg1)
                                Exit Sub
                                'Throw 'ex

                            End Try
                            Call TIMS.CloseDbConn(tConn)
                        End Using

                    End If
                End If
                iRowIndex += 1
            Next
        End If


        '判斷匯出資料是否有誤
        Dim explain As String = ""
        explain &= "匯入資料共" & dt_xls.Rows.Count & "筆" & vbCrLf
        explain &= "成功：" & (dt_xls.Rows.Count - dtWrong.Rows.Count) & "筆" & vbCrLf
        explain &= "失敗：" & dtWrong.Rows.Count & "筆" & vbCrLf

        Dim explain2 As String = ""
        explain2 &= "匯入資料共" & dt_xls.Rows.Count & "筆\n"
        explain2 &= "成功：" & (dt_xls.Rows.Count - dtWrong.Rows.Count) & "筆\n"
        explain2 &= "失敗：" & dtWrong.Rows.Count & "筆\n"

        If dtWrong.Rows.Count = 0 Then
            Common.MessageBox(Me, If(Reason <> "", String.Concat(explain, Reason), explain))
        Else
            'Session("MyWrongTable") = dtWrong
            Datagrid2.Style.Item("display") = "inline"
            Datagrid2.Visible = True
            Datagrid2.DataSource = dtWrong
            Datagrid2.DataBind()
            Common.MessageBox(Me, "資料匯入成功,但有錯誤資料請檢示原因!!!")
            For i As Integer = 1 To 100
                If i = 100 Then eMeng.Style.Item("display") = "inline"
                'Page.RegisterStartupScript("", "<script>{window.document.getElementById('eMeng').style.visibility='visible';}</script>")
            Next
            'Page.RegisterStartupScript("", "<script>if(confirm('資料匯入成功，但有錯誤的資料無法匯入，是否要檢視原因?')){window.open('TC_01_007_Wrong.aspx','','width=500,height=500,location=0,status=0,menubar=0,scrollbars=1,resizable=0');}</script>")
        End If
        'Button1_Click(sender, e)
        Call GClickSearchButton()

    End Sub

#Region "匯出教師資料 EXCEL檔"
#End Region

    '20080603  Andy 新增匯出教師資料 '匯出名冊
    Sub Export1()
        'Dim sql As String
        'Dim dt As DataTable
        'Dim dr As DataRow
        'copy一份sample資料---------------------   Start
        'Dim MyFile As System.IO.File
        'Dim MyDownload As System.IO.File
        Dim strErrmsg As String = ""

        center.Value = TIMS.ClearSQM(center.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If center.Value = "" Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If
        If RIDValue.Value = "" Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If
        'Dim ExpTitle As String = center.Value.ToString() & Format(Date.Now(), "yyyy-M-d")
        'Dim ExpTitle As String = center.Value.ToString() & Format(Date.Now(), "yyyy-M")
        Dim ExpTitle As String = center.Value
        Dim MyPath As String = ""
        Dim sFileName As String = ""
        sFileName &= "~\TC\01\Temp\"
        sFileName &= TIMS.ChangeIDNO(Replace(Replace(Replace(ExpTitle, ")", ""), "(", ""), "/", ""))
        sFileName &= TIMS.GetDateNo()
        sFileName &= ".xls"
        MyPath = Server.MapPath(sFileName)

        Dim MyFileName As String = TIMS.ChangeIDNO(Replace(Replace(Replace(ExpTitle, ")", ""), "(", ""), "/", "")) & ".xls"
        Const cst_Sample1xls As String = "~\TC\01\Temp\Sample1.xls"
        If Not IO.File.Exists(Server.MapPath(cst_Sample1xls)) Then
            Common.MessageBox(Me, "Sample檔案不存在")
            Exit Sub
        End If
        Try
            IO.File.Copy(Server.MapPath(cst_Sample1xls), MyPath, True)
            '除去sample檔的唯讀屬性
            'MyFile.SetAttributes(Server.MapPath("~\TC\01\Temp\" & Replace(Replace(Replace(ExpTitle, ")", ""), "(", ""), "/", "") & ".xls"), IO.FileAttributes.Normal)
            IO.File.SetAttributes(MyPath, IO.FileAttributes.Normal)
        Catch ex As Exception
            strErrmsg = ""
            strErrmsg &= "目錄名稱或磁碟區標籤語法錯誤!!!" & vbCrLf
            strErrmsg &= " (若連續出現多次(3次以上)請連絡系統管理者協助，謝謝)(造成您的不便深感抱歉)" & vbCrLf
            strErrmsg += ex.ToString & vbCrLf
            Common.MessageBox(Me, strErrmsg)
            'Exit Sub
        End Try
        'copy一份sample資料---------------------   End

        '根據路徑建立資料庫連線，並取出講師資料填入---------------   Start
        'RIDValue.Value.ToString()
        Dim dt As DataTable = Nothing
        Dim sql As String = ""
        sql &= " SELECT TECHID,RID,TEACHERID,TEACHCNAME,TEACHENAME,PASSPORTNO,IDNO,BIRTHDAY,SEX,TMID" & vbCrLf

        sql &= If(RIDValue.Value.Length = 1, " ,IVID", " ,Invest IVID") & vbCrLf

        sql &= " ,KindEngage,KindID,DegreeID , GraduateStatus, SchoolName, Department, Phone, Mobile, Email" & vbCrLf
        sql &= " ,AddressZip,AddressZIP6W,Address, WorkOrg, ExpYears ,ServDept , WorkPhone, Fax" & vbCrLf
        sql &= " ,WorkZip,WorkZIP6W,Workaddr, ExpUnit1, ExpUnit2, ExpUnit3, ExpYears1, ExpYears2, ExpYears3" & vbCrLf
        sql &= " ,ExpSDate1,ExpEDate1, ExpSDate2,ExpEDate2,ExpSDate3,ExpEDate3, Specialty1,Specialty2 , Specialty3 , Specialty4" & vbCrLf
        sql &= " ,Specialty5 ,TransBook , ProLicense ,PassPortNO, WorkStatus, ModifyAcct, ModifyDate" & vbCrLf
        sql &= " FROM Teach_TeacherInfo" & vbCrLf
        sql &= " WHERE RID=@RID" & vbCrLf

        Dim parms As New Hashtable From {{"RID", RIDValue.Value}}
        dt = DbAccess.GetDataTable(sql, objconn, parms)

        '查詢原因
        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        Dim v_rblWorkMode As String = TIMS.GetListValue(rblWorkMode)
        Session(TIMS.gcst_rblWorkMode) = v_rblWorkMode 'rblWorkMode.SelectedValue
        sMemo = GET_SEARCH_MEMO()
        '--查詢原因:INQNO--查詢結果筆數:RESCNT--查詢清單結果:RESDESC 
        Dim vRESDESC As String = TIMS.GET_RESDESCdt(dt, "TECHID,TEACHERID,TEACHCNAME,TEACHENAME,IDNO,BIRTHDAY,PHONE,MOBILE,EMAIL")
        Call TIMS.SubInsAccountLog1(Me, TIMS.Get_MRqID(Me), TIMS.cst_wm匯出, v_rblWorkMode, "", sMemo, objconn, v_INQUIRY, dt.Rows.Count, vRESDESC)

        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        Using MyConn As New OleDb.OleDbConnection
            MyConn.ConnectionString = TIMS.Get_OleDbStr(MyPath)
            Try
                MyConn.Open()
            Catch ex As Exception
                'Dim strErrmsg As String = ""
                strErrmsg = ""
                strErrmsg &= "/* ex.ToString: */" & vbCrLf & ex.ToString & vbCrLf
                strErrmsg &= "sql:" & vbCrLf & sql & vbCrLf
                strErrmsg &= "conn.ConnectionString:" & MyConn.ConnectionString & vbCrLf
                strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
                'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                Call TIMS.WriteTraceLog(strErrmsg)

                'Common.MessageBox(Me, "Excel資料無法開啟連線!")
                Common.MessageBox(Me, TIMS.cst_NODATAMsg2)
                Exit Sub
            End Try

            dt.DefaultView.Sort = "TechID"

            'Dim RID As String '機構
            Dim RIDLevel As String '計畫階層
            Dim TeacherID As String '講師代碼
            Dim TeachCName As String '講師姓名
            Dim TeachEName As String '講師英文姓名
            Dim PassPortNO As String '身份別
            Dim IDNO As String '身分證號碼
            Dim Birthday As String '出生日期
            Dim Sex As String '性別
            Dim TMID As String '職類代碼
            Dim IVID As String '職稱代碼
            Dim KindEngage As String '內外聘
            Dim KindID As String '師資別
            Dim DegreeID As String '學歷
            Dim GraduateStatus As String '畢業狀況
            Dim SchoolName As String '學校名稱
            Dim Department As String '科系名稱
            Dim Phone As String '聯絡電話
            Dim Mobile As String '行動電話
            Dim Email As String 'E_Mail
            Dim AddressZip As String '戶藉地址Zip
            Dim AddressZIP6W As String '戶藉地址Zip後2碼
            Dim Address As String '戶藉地址
            Dim WorkOrg As String '服務單位名稱
            Dim ExpYears As String '服務年資
            Dim ServDept As String '服務部門
            Dim WorkPhone As String '服務單位電話
            Dim Fax As String '服務單位傳真
            Dim WorkZip As String '服務單位地址Zip
            Dim WorkZIP6W As String '服務單位地址Zip後2碼
            Dim Workaddr As String ' 服務單位地址
            Dim ExpUnit1 As String '服務單位一
            Dim ExpUnit2 As String '服務單位二
            Dim ExpUnit3 As String '服務單位三
            Dim ExpYears1 As String '服務年資一
            Dim ExpYears2 As String '服務年資二
            Dim ExpYears3 As String '服務年資三
            Dim ExpSDate1 As String '服務單位一起日
            Dim ExpEDate1 As String '服務單位一迄日
            Dim ExpSDate2 As String '服務單位二起日
            Dim ExpEDate2 As String '服務單位二迄日
            Dim ExpSDate3 As String '服務單位三起日
            Dim ExpEDate3 As String '服務單位三迄日
            Dim Specialty1 As String '專長一
            Dim Specialty2 As String '專長二
            Dim Specialty3 As String '專長三
            Dim Specialty4 As String '專長四
            Dim Specialty5 As String '專長五
            Dim TransBook As String '譯著
            Dim ProLicense As String '專業證照
            Dim WorkStatus As String '任職狀況
            Dim ModifyAcct As String '異動者
            Dim ModifyDate As String '異動時間
            'Dim PassPortNO As String        ' 
            '----------------------------------
            For Each dr As DataRow In dt.Rows
                'RID = Right(dr("RID").ToString, 2)
                'If center.Value <> "" Then center.Value = Trim(center.Value)
                RIDLevel = TIMS.ClearSQM(center.Value)
                TeacherID = TIMS.ClearSQM(dr("TeacherID"))
                TeachCName = TIMS.ClearSQM(dr("TeachCName"))
                TeachEName = TIMS.ClearSQM(dr("TeachEName"))
                Select Case TIMS.ClearSQM(dr("PassPortNO"))
                    Case "1", "2"
                        PassPortNO = dr("PassPortNO").ToString
                    Case Else
                        PassPortNO = "2"
                End Select
                IDNO = TIMS.ClearSQM(dr("IDNO"))
                If rblWorkMode.SelectedValue = "1" Then IDNO = TIMS.strMask(IDNO, 1)
                Birthday = ""
                If Convert.ToString(dr("Birthday")) <> "" Then
                    Birthday = TIMS.Cdate3(dr("Birthday"))
                    If rblWorkMode.SelectedValue = "1" Then Birthday = TIMS.strMask(Birthday, 2)
                End If
                '   Birthday = dr("Birthday").ToString
                Sex = dr("Sex").ToString
                TMID = dr("TMID").ToString
                IVID = TIMS.ClearSQM(dr("IVID"))

                KindEngage = dr("KindEngage").ToString
                KindID = dr("KindID").ToString
                DegreeID = dr("DegreeID").ToString
                GraduateStatus = dr("GraduateStatus").ToString

                SchoolName = TIMS.ClearSQM(dr("SchoolName"))
                Department = TIMS.ClearSQM(dr("Department"))
                Phone = TIMS.ClearSQM(dr("Phone"))
                Mobile = TIMS.ClearSQM(dr("Mobile"))
                Email = TIMS.ClearSQM(dr("Email"))

                AddressZip = TIMS.ClearSQM(dr("AddressZip"))
                AddressZIP6W = TIMS.ClearSQM(dr("AddressZIP6W"))
                Address = TIMS.ClearSQM(dr("Address"))

                WorkOrg = TIMS.ClearSQM(dr("WorkOrg"))
                ExpYears = TIMS.ClearSQM(dr("ExpYears"))
                ServDept = TIMS.ClearSQM(dr("ServDept"))
                WorkPhone = TIMS.ClearSQM(dr("WorkPhone"))
                Fax = TIMS.ClearSQM(dr("Fax"))
                WorkZip = TIMS.ClearSQM(dr("WorkZip"))
                WorkZIP6W = TIMS.ClearSQM(dr("WorkZIP6W"))
                Workaddr = TIMS.ClearSQM(dr("Workaddr"))
                ExpUnit1 = TIMS.ClearSQM(dr("ExpUnit1"))
                ExpUnit2 = TIMS.ClearSQM(dr("ExpUnit2"))
                ExpUnit3 = TIMS.ClearSQM(dr("ExpUnit3"))
                ExpYears1 = dr("ExpYears1").ToString
                ExpYears2 = dr("ExpYears2").ToString
                ExpYears3 = dr("ExpYears3").ToString
                ExpSDate1 = ""
                If dr("ExpSDate1").ToString <> "" Then
                    ExpSDate1 = FormatDateTime(dr("ExpSDate1").ToString, DateFormat.ShortDate)
                Else
                    ExpSDate1 = ""
                End If
                If dr("ExpSDate2").ToString <> "" Then
                    ExpSDate2 = FormatDateTime(dr("ExpSDate2").ToString, DateFormat.ShortDate)
                Else
                    ExpSDate2 = ""
                End If
                If dr("ExpSDate3").ToString <> "" Then
                    ExpSDate3 = FormatDateTime(dr("ExpSDate3").ToString, DateFormat.ShortDate)
                Else
                    ExpSDate3 = ""
                End If
                If dr("ExpEDate1").ToString <> "" Then
                    ExpEDate1 = FormatDateTime(dr("ExpEDate1").ToString, DateFormat.ShortDate)
                Else
                    ExpEDate1 = ""
                End If
                If dr("ExpEDate2").ToString <> "" Then
                    ExpEDate2 = FormatDateTime(dr("ExpEDate2").ToString, DateFormat.ShortDate)
                Else
                    ExpEDate2 = ""
                End If
                If dr("ExpEDate3").ToString <> "" Then
                    ExpEDate3 = FormatDateTime(dr("ExpEDate3").ToString, DateFormat.ShortDate)
                Else
                    ExpEDate3 = ""
                End If
                'ExpSDate1 = dr("ExpSDate1").ToString
                'ExpSDate2 = dr("ExpSDate2").ToString
                'ExpSDate3 = dr("ExpSDate3").ToString
                'ExpEDate1 = dr("ExpEDate1").ToString
                'ExpEDate2 = dr("ExpEDate2").ToString
                'ExpEDate3 = dr("ExpEDate3").ToString
                Specialty1 = TIMS.ChangeSQM(dr("Specialty1")) '專長一
                Specialty2 = TIMS.ChangeSQM(dr("Specialty2")) '專長二
                Specialty3 = TIMS.ChangeSQM(dr("Specialty3")) '專長三
                Specialty4 = TIMS.ChangeSQM(dr("Specialty4")) '專長四
                Specialty5 = TIMS.ChangeSQM(dr("Specialty5")) '專長五
                TransBook = TIMS.ChangeSQM(dr("TransBook"))   '譯著
                ProLicense = TIMS.ChangeSQM(dr("ProLicense")) '專業證照

                WorkStatus = dr("WorkStatus").ToString '任職狀況
                ModifyAcct = dr("ModifyAcct").ToString
                ModifyDate = dr("Specialty1").ToString
                'PassPortNO = dr("PassPortNO").ToString
                '------------------------

                sql = "INSERT INTO [Sheet1$] ("
                sql &= "計劃階層,講師代碼,講師姓名,講師英文姓名,身份別,身分證字號,出生日期,性別,主要職類,職稱,"
                sql &= "內外聘,師資別,最高學歷,畢業狀況,學校名稱,科系名稱,聯絡電話,行動電話,電子郵件,"
                sql &= "郵遞區號前3碼,郵遞區號6碼,戶籍地址,服務單位名稱,年資,服務部門,服務單位電話,服務單位傳真,"
                sql &= "服務單位郵遞區號前3碼,服務單位郵遞區號6碼,服務單位地址,服務單位一,服務單位二,服務單位三,"
                sql &= "服務年資一,服務年資二,服務年資三,服務期間一起日,服務期間一迄日,服務期間二起日,"
                sql &= "服務期間二迄日,服務期間三起日,服務期間三迄日,專長一,專長二,專長三,專長四,專長五,譯著,專業證照,排課使用"
                sql &= ")"
                sql &= "VALUES ("
                sql &= "'" & RIDLevel & "','" & TeacherID & "','" & TeachCName & "','" & TeachEName & "','" & PassPortNO & "','" & IDNO & "','" & Birthday & "','" & Sex & "','" & TMID & "','" & IVID & "',"
                sql &= "'" & KindEngage & "','" & KindID & "','" & DegreeID & "','" & GraduateStatus & "' ,'" & SchoolName & "','" & Department & "','" & Phone & "','" & Mobile & "','" & Email & "',"
                sql &= "'" & AddressZip & "','" & AddressZIP6W & "','" & Address & "','" & WorkOrg & "','" & ExpYears & "','" & ServDept & "','" & WorkPhone & "','" & Fax & "',"
                sql &= "'" & WorkZip & "' , '" & WorkZIP6W & "' , '" & Workaddr & "' , '" & ExpUnit1 & "' , '" & ExpUnit2 & "' , '" & ExpUnit3 & "',"
                sql &= "'" & ExpYears1 & "' ,'" & ExpYears2 & "' ,'" & ExpYears3 & "','" & ExpSDate1 & "' , '" & ExpEDate1 & "' ,'" & ExpSDate2 & "',"
                sql &= "'" & ExpEDate2 & "' ,'" & ExpSDate3 & "' , '" & ExpEDate3 & "' , '" & Specialty1 & "' , '" & Specialty2 & "' ,"
                sql &= "'" & Specialty3 & "' , '" & Specialty4 & "' , '" & Specialty5 & "' , '" & TransBook & "' , '" & ProLicense & "','" & WorkStatus & "'"
                sql &= ")"

                Using ole_Cmd As New OleDb.OleDbCommand(sql, MyConn)
                    'cmd = New OleDb.OleDbCommand(sql, ole_conn)
                    Try
                        If MyConn.State = ConnectionState.Closed Then MyConn.Open()
                        ole_Cmd.ExecuteNonQuery()
                        'DbAccess.ExecuteNonQuery(ole_Cmd.CommandText, objconn, ole_Cmd.Parameters)
                        'If conn.State = ConnectionState.Open Then conn.Close()
                    Catch ex As Exception
                        'Dim strErrmsg As String = ""
                        strErrmsg &= "/* ex.ToString: */" & vbCrLf & ex.ToString & vbCrLf
                        strErrmsg &= "sql:" & vbCrLf & sql & vbCrLf
                        strErrmsg &= "conn.ConnectionString:" & MyConn.ConnectionString & vbCrLf
                        strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
                        strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                        Call TIMS.WriteTraceLog(strErrmsg)

                        If MyConn.State = ConnectionState.Open Then MyConn.Close()
                        Exit For
                        'Throw ex
                    End Try
                End Using
            Next
            If MyConn.State = ConnectionState.Open Then MyConn.Close()
            '根據路徑建立資料庫連線，並取出學員資料填入---------------   End
        End Using

        Dim V_ExpType As String = TIMS.GetListValue(RBListExpType)
        Select Case V_ExpType
            Case "EXCEL"
                ExpExccl_1(strErrmsg, MyPath)

            Case "ODS"
                Dim fr As New System.IO.FileStream(MyPath, IO.FileMode.Open)
                Dim br As New System.IO.BinaryReader(fr)
                Dim buf(fr.Length) As Byte
                fr.Read(buf, 0, fr.Length)
                fr.Close()

                Dim sFileName1 As String = "ExpFile" & TIMS.GetRnd6Eng()
                Dim parmsExp As New Hashtable
                parmsExp.Add("ExpType", V_ExpType) 'EXCEL/PDF/ODS
                parmsExp.Add("FileName", sFileName1)
                parmsExp.Add("xlsx_buf", buf)
                'parmsExp.Add("strHTML", strHTML)
                parmsExp.Add("ResponseNoEnd", "Y")
                TIMS.Utl_ExportRp1(Me, parmsExp)
            Case Else
                Dim s_log1 As String = ""
                s_log1 = String.Format("ExpType(參數有誤)!!{0}", V_ExpType)
                Common.MessageBox(Me, s_log1)
                Exit Sub
        End Select

        Call TIMS.MyFileDelete(MyPath)
        TIMS.Utl_RespWriteEnd(Me, objconn, "")
        Exit Sub
    End Sub

    Sub ExpExccl_1(ByRef strErrmsg As String, ByRef MyPath As String)
        '將新建立的excel存入記憶體下載-----   Start
        'Dim strErrmsg As String = ""
        'strErrmsg = ""
        Dim myFileName1 As String = TIMS.ClearSQM("TeacherList" & TIMS.GetRnd6Eng) & ".xls" '檔名
        Try
            Dim fr As New System.IO.FileStream(MyPath, IO.FileMode.Open)
            Dim br As New System.IO.BinaryReader(fr)
            Dim buf(fr.Length) As Byte
            fr.Read(buf, 0, fr.Length)
            fr.Close()

            Response.Clear()
            Response.ClearHeaders()
            Response.Buffer = True
            Response.AppendHeader("Content-Disposition", "attachment;filename=" & HttpUtility.UrlEncode(myFileName1, System.Text.Encoding.UTF8))
            'Response.AddHeader("content-disposition", "attachment;filename=" & HttpUtility.UrlEncode(MyFileName, System.Text.Encoding.UTF8))
            Response.ContentType = "Application/vnd.ms-Excel"
            'Common.RespWrite(Me, br.ReadBytes(fr.Length))
            Response.BinaryWrite(buf)
            'Response.End()
        Catch ex As Exception
            'Dim strErrmsg As String = ""
            'strErrmsg = ""
            'strErrmsg &= "/* ex.ToString: */" & vbCrLf & ex.ToString & vbCrLf
            'strErrmsg &= "sql:" & vbCrLf & Sql & vbCrLf
            'strErrmsg &= "MyPath:" & MyPath & vbCrLf
            'strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            'Call TIMS.WriteTraceLog(strErrmsg)

            strErrmsg = ""
            strErrmsg &= "無法存取該檔案!!!" & vbCrLf
            strErrmsg &= " (若連續出現多次(3次以上)請連絡系統管理者協助，謝謝)(造成您的不便深感抱歉)" & vbCrLf
            'strErrmsg += ex.ToString & vbCrLf
            Common.MessageBox(Me, strErrmsg)

            'Finally
            '刪除Temp中的資料
            'If MyFile.Exists(MyPath) Then MyFile.Delete(MyPath)
        End Try

    End Sub

    '查詢鈕  '匯出鈕 'hidSchBtnNum.value: 1.正常查詢 2.正常匯出
    Sub SUtl_btnSearchData1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim BtnObj As Button = CType(sender, Button)
        Const cst_button1 As String = "button1" '查詢
        Const cst_btnxlsemport As String = "btnxlsemport" '匯出'Btn_XlsEmport_Click
        Const cst_btndivPwdSubmit As String = "btndivpwdsubmit" 'hidSchBtnNum.value: 1.正常查詢 2.正常匯出
        Dim sMsg As String = ""
        eMeng.Style.Item("display") = "none"
        Datagrid2.Visible = False

        '取出鍵詞-查詢原因
        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        If (TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES)) Then
            If (v_INQUIRY = "") Then Common.MessageBox(Me, "請選擇「查詢原因」") : Return
            Session(String.Concat(TIMS.cst_GSE_V_INQUIRY, TIMS.Get_MRqID(Me))) = v_INQUIRY
        End If

        Select Case LCase(BtnObj.CommandName)
            Case cst_button1 '查詢鈕
                Call GClickSearchButton()
            Case cst_btnxlsemport '匯出鈕 'Btn_XlsEmport_Click
                Call Export1()
            Case cst_btndivPwdSubmit
                '正常顯示 '查詢或匯出。
                If Not TIMS.sUtl_ChkPlanPwd(sm.UserInfo.PlanID, objconn) Then
                    sMsg = "未設定計畫密碼!!"
                    labChkMsg.Text = sMsg
                    Common.MessageBox(Me, sMsg)
                    Exit Sub
                End If
                If Not TIMS.sUtl_ChkPlanPwdOK(objconn, sm.UserInfo.PlanID, txtdivPxssward.Text) Then
                    sMsg = "個資安全密碼錯誤!!"
                    labChkMsg.Text = sMsg
                    Common.MessageBox(Me, sMsg)
                    Exit Sub
                End If
                'If rblWorkMode.SelectedValue = "2" Then
                '    flgCIShow = True '可正常顯示個資。
                'End If
                txtdivPxssward.Text = ""
                Select Case hidSchBtnNum.Value
                    Case "1"
                        Call GClickSearchButton()
                    Case "2"
                        Call Export1()
                End Select
        End Select
    End Sub

    Protected Sub Btn_XlsEmport_Click(sender As Object, e As EventArgs) Handles Btn_XlsEmport.Click

    End Sub

    Protected Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

    End Sub


#Region "NO USE"
    'Protected Sub btnCloseLoginDiv_Click(sender As Object, e As EventArgs) Handles btnCloseLoginDiv.Click
    '    'panelLoginDiv.Visible = False
    '    panelLoginDiv.Style.Item("display") = "none"
    '    labChkMsg.Text = ""
    'End Sub

    'Function Check_Data_Protection_State() As Boolean
    '    Dim rst As Boolean = False
    '    If rblWorkMode.SelectedValue = "1" Then
    '        rst = False
    '    Else
    '        If hidCheckPasswordState.Value = "True" Then    'Password checked
    '            'panelLoginDiv.Visible = False
    '            panelLoginDiv.Style.Item("display") = "none"
    '        Else
    '            If hidLockTime1.Value = "1" Or hidLockTime1.Value = "" Then
    '                'panelLoginDiv.Visible = True
    '                panelLoginDiv.Style.Item("display") = "inline"
    '                rst = True
    '            End If
    '        End If
    '    End If
    '    If hidLockTime1.Value = "0" Then
    '        rblWorkMode.SelectedValue = "2"
    '    End If
    '    hidWorkMode.Value = rblWorkMode.SelectedValue
    '    Return rst
    'End Function

    'Protected Sub btnSubmit_Click(sender As Object, e As EventArgs) Handles btnSubmit.Click
    '    If Not TIMS.sUtl_ChkPlanPwd(sm.UserInfo.PlanID) Then
    '        labChkMsg.Text = "尚未設定個資安全密碼!!"
    '        'panelLoginDiv.Visible = True
    '        panelLoginDiv.Style.Item("display") = "inline"
    '        hidCheckPasswordState.Value = "False"
    '        Exit Sub
    '    End If

    '    If Not TIMS.sUtl_ChkPlanPwdOK(sm.UserInfo.PlanID, txtdivPxssward.Text) Then
    '        labChkMsg.Text = "個資安全密碼錯誤!!"
    '        'panelLoginDiv.Visible = True
    '        panelLoginDiv.Style.Item("display") = "inline"
    '        hidCheckPasswordState.Value = "False"
    '    Else
    '        'panelLoginDiv.Visible = False
    '        panelLoginDiv.Style.Item("display") = "none"
    '        hidCheckPasswordState.Value = "True"
    '        'Button1_Click(sender, e)
    '        Call gClickSearchButton()
    '    End If

    'End Sub
#End Region

    'sUtl_btnSearchData1
    'Protected Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
    'Call gClickSearchButton()
    'End Sub

End Class
