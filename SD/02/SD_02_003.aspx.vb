Partial Class SD_02_003
    Inherits AuthBasePage

    '未做 首頁>>學員動態管理>>招生作業>>甄試結果試算 
    '不可使用該功能(首頁>>學員動態管理>>招生作業>>錄取作業 )。SELECT * FROM STUD_SELRESULT WHERE ROWNUM <=10

    Dim s_Msg1 As String = ""

    Dim gm1_DataGrid As DataGrid
    Const cst_printFN1 As String = "SD02003_RXC"
    Const vs_SD02003_OCID1 As String = "SD02003_OCID1"

    Const cst_watingcount As String = "watingcount"
    Const cst_vsSort As String = "Sort"
    Const cst_tip2 As String = "未做 甄試結果試算 可能查無錄取作業資料。"
    Const cst_errmsg1 As String = "資料儲存時發生錯誤，請重新查詢!!"
    Const cst_msg0001 As String = "學員資料已經存在不可修改"
    Const cst_msg0002 As String = "已設定為正取，但該學員資料尚未建立!"
    Const cst_msg0003 As String = "學員資料已經存在!!"
    Const cst_msg0004 As String = "請設定為正取,學員資料已經存在!!"
    Const cst_msg0005 As String = "錄取完成!!,取消備取名次儲存"
    Const cst_msg0006 As String = "錄取完成!!" & vbCrLf
    'Const cst_msg0007 As String="已按過「送出」鈕，不可再進行錄取作業修改!!" & vbCrLf
    Const cst_msg0008 As String = "錄訓名單審核，尚未完成，不可列印!!"
    'Const cst_msg0009 As String="列印鍵非產投方案適用功能，故無需使用。"

    'Const cst_msg0010 As String="已完成[送出]作業,故無法再次送出"
    Const cst_msg0011 As String = "登入權限與計劃權限不符，停用此功能"
    Const cst_msg0012 As String = "報名班級已過開訓14日，不可再進行錄取作業。"
    Const cst_msg0013 As String = "報名班級已過開訓日，鎖定功能填寫。"
    Const cst_msg0014 As String = "目前尚有e網報名學員尚未審核，需先將【e網報名】的所有學員審核完成後，始能登錄【錄取作業】!"
    Const cst_msg0015 As String = "針對【重大災害提醒處理】功能，有未進行提醒處理，無法進行該作業，請先完成 重大災害提醒處理!"
    Const cst_msg0016 As String = "確認錄取作業是否完成?"
    Const cst_msg0017 As String = "當「名次」 位於「訓練人數」名額內，若錄訓結果為「備取」或「未錄取」請增加填寫「備取或未錄取原因」"

    Const cst_OCIDValue1 As String = "OCIDValue1"
    Const cst_Desc As String = " Desc"
    'Const Cst_TotalResult As String="WSort,TotalResult"
    'Const Cst_ExamNo As String="WSort,ExamNo"
    Const cst_TotalResult As String = "TotalResult"
    Const cst_ExamNo As String = "ExamNo"
    Const cst_SIGNNO As String = "SIGNNO"

    Const cst_rsort As String = "rsort"
    Const cst_totalresult2 As String = "totalresult"
    Const cst_examno2 As String = "examno"

    '增修需求 OJT-21012501：<系統> 產投 - 錄訓作業：介面欄位調整
    Const cst_ceExamNo As Integer = 0 '自辦用准考證序號
    Const cst_ceSIGNNO As Integer = 1 '產投用e網報名序號
    'Const cst_ceName As Integer=2 '姓名
    Const cst_ceWriteResult As Integer = 3 '筆試
    Const cst_ceOralResult As Integer = 4 '口試
    Const cst_ceTotalResult As Integer = 5 '總成績
    Const cst_ceEnterDate As Integer = 6 '報名日期
    Const cst_ceRSort As Integer = 7 '名次
    Const cst_ceddlResult As Integer = 8 '甄試結果 甄試結果 > 錄訓結果(產投用)
    Const cst_ceselsort As Integer = 9 '備取名次
    Const cst_cenotes2 As Integer = 10 '備註

    'Const Cst_ceTRNDType As Integer=7
    'Const Cst_ceSelResultID As Integer=9
    'Const Cst_ceSETID As Integer=10
    'Const Cst_ceSerNum As Integer=11
    'Const Cst_cenotes2 As Integer=13 '備註

    Const cst_msg219 As String = "※ 姓名前標記「x-」表示民眾已註銷推介"
    Const cst_fgb219 As String = "x-"
    Const cst_Mgc219 As String = "民眾已註銷推介"

    '屆退官兵者 (依系統日期判斷) 判斷計畫為自辦職前。
    'Dim flagTPlanID02Plan2 As Boolean=False

    Dim flag_no_use_print_btn As Boolean = False '列印鍵非產投方案適用功能，故無需使用。
    Dim flgROLEIDx0xLIDx0 As Boolean = False   '判斷登入者的權限。

    'WSort: 1.就服單位協助報名
    'Const cst_EnterPathW As String="W" '就服站代碼'WSort
    Const cst_EnterPathNameW As String = "(就服單位協助報名)" '說明

    '#Region "(No Use)"
    'SELECT * FROM Stud_SelResult WHERE OCID=38536
    'SELECT * FROM Stud_EnterType WHERE OCID1=38536
    'SELECT * 
    'FROM Stud_EnterTemp C
    'WHERE EXISTS (
    '	SELECT 'X' FROM Stud_EnterType X WHERE X.OCID1=38536
    '	AND X.SETID=C.SETID
    ')
    '委訓單位 將報名學員設定為正取、備取、未錄取。
    '查詢、e網公告、完成錄取、挑選其他志願、備取儲存。

    'Dim oflag_Test As Boolean=False '測試 (false:正式)
    'Dim dtArc As DataTable '暫時權限Table

    'Dim au As New cAUTH
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
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        '是否為超級使用者'是否為(後台)系統管理者 
        flgROLEIDx0xLIDx0 = TIMS.IsSuperUser(Me, 1)

        btnSAVE2.Visible = False '解鎖不顯示
        '屆退官兵者 (依系統日期判斷)
        'flagTPlanID02Plan2=False '判斷計畫為自辦職前。
        'If TIMS.Cst_TPlanID02Plan2.IndexOf(sm.UserInfo.TPlanID) > -1 Then flagTPlanID02Plan2=True '判斷計畫為自辦職前。

        'oflag_Test=TIMS.sUtl_ChkTest() '測試
        tr_rblsortmode.Visible = True
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then tr_rblsortmode.Visible = False

        Dim flagS1 As Boolean = flgROLEIDx0xLIDx0 ' TIMS.IsSuperUser(Me, 1) '是否為(後台)系統管理者 
        '檢查帳號的功能權限-----------------------------------Start
        Select Case sm.UserInfo.LID
            Case 1 '分署不鎖
                btnSEND1.Enabled = False
                TIMS.Tooltip(btnSEND1, "")
            Case Else
                If RIDValue.Value <> "" AndAlso RIDValue.Value <> sm.UserInfo.RID Then
                    If Not flagS1 Then
                        btnSEND1.Enabled = False
                        TIMS.Tooltip(btnSEND1, cst_msg0011, True)
                    End If
                    'If Not flagS1 Then,'    btnSAVE1.Enabled=False,'    TIMS.Tooltip(btnSAVE1, cst_msg0011, True),'End If,'button3.Disabled=True,'TIMS.Tooltip(button3, "登入權限與計劃權限不符，停用此功能", True),
                End If
        End Select

        'If Not flagS1 Then
        '    button1.Enabled=False
        '    If au.blnCanSech Then button1.Enabled=True
        '    If Not au.blnCanSech Then TIMS.Tooltip(button1, "無查詢功能權限", True)
        'End If
        '檢查帳號的功能權限-----------------------------------End
        'If oflag_Test Then
        '    '測試用
        '    btnSEND1.Enabled=True
        '    TIMS.Tooltip(btnSEND1, "測試開啟", True)
        '    'button3.Disabled=False
        '    'TIMS.Tooltip(button3, "測試開啟", True)
        'End If

        'labmsg219.Text=cst_msg219   '(依照承辦人需求,將此訊息隱藏起來，by:20180919)
        If Not IsPostBack Then Call Create1() '第1次載入執行

        Call EveryCreate2() '每次載入執行
    End Sub

    '每次載入執行
    Sub EveryCreate2()
        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        button8.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, historyrid, "HistoryList2", "RIDValue", "center")
        If historyrid.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

        '確認機構是否為黑名單
        Dim vsMsg2 As String = "" '確認機構是否為黑名單
        'vsMsg2=""
        If Chk_OrgBlackList(vsMsg2) Then
            'button6.Enabled=False
            'TIMS.Tooltip(button6, vsMsg2)
            btnSEND1.Enabled = False
            TIMS.Tooltip(btnSEND1, vsMsg2)

            'button3.Disabled=True
            'TIMS.Tooltip(button3, vsMsg2)
            Dim vsStrScript As String = $"<script>alert('{vsMsg2}');</script>"
            Page.RegisterStartupScript("", vsStrScript)
        End If
    End Sub

    '第1次載入執行
    Sub Create1()
        'btnSAVE1.Attributes.Add("onclick", "return savedataCHK1();")
        lab_msg1.Text = ""
        lab_msg_r1.Text = ""
        Hid_CAN_IGNORE_RULE1_CNT.Value = ""
        ViewState(vs_SD02003_OCID1) = Nothing
        flag_no_use_print_btn = If(TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1, True, False)
        If flag_no_use_print_btn Then
            'btnPrint1.Enabled=False
            'TIMS.Tooltip(btnPrint1, cst_msg0009)
            btnPrint1.Visible = False
            lab_msg_r1.Text = cst_msg0017 ' cst_msg0009
        End If

        table4.Visible = False
        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID

        If sm.UserInfo.LID <> "2" Then
            TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
        Else
            'center.Enabled=False
            Call TIMS.GET_OnlyOne_OCID2(Me, RIDValue.Value, TMID1, OCID1, TMIDValue1, OCIDValue1, objconn)
            'Button4_Click(sender, e)
        End If
        button1.Attributes("onclick") = "javascript:return search()"
        btnSEND1.Attributes("onclick") = "javascript:return confirm('" & cst_msg0016 & "')"
        btnsave.Visible = False '備取設定的儲存
        btncancel.Visible = False '備取設定的取消

        '------------------判斷是否可以備取排名----------------------
        'Dim dr2 As DataRow
        'sql="SELECT * FROM Sys_GlobalVar where Distid ='" & sm.UserInfo.DistID & "' and Tplanid ='" & sm.UserInfo.TPlanID & "' and GVID=22 and itemVar1='Y'"
        'dr2=DbAccess.GetOneRow(sql)
        Hid_ChangeWating.Value = TIMS.GetGlobalVar(Me, "24", "1", objconn)
        If btnSEND1.Enabled Then
            If Hid_ChangeWating.Value = "Y" Then
                'BtnWatingSet.Visible=True
                btnSEND1.Text = "完成錄取及備取排名設定"
                btnSEND1.CssClass = "asp_button_L"
                Hid_ChangeWating.Value = "Y"
            Else
                'BtnWatingSet.Visible=FALSE
                btnSEND1.Text = "完成錄取"
                btnSEND1.CssClass = "asp_button_M"
                Hid_ChangeWating.Value = "N"
            End If
        End If
        '-----------------判斷是否可以備取排名End----------------------------

        'If ViewState("open")=1 Then
        '    ViewState("open")=0
        '    'Button1_Click(sender, e)
        '    Search1()
        'End If

    End Sub

    '機構黑名單內容(訓練單位處分功能)
    Function Chk_OrgBlackList(ByRef Errmsg As String) As Boolean
        Dim rst As Boolean = False
        Errmsg = ""
        Dim vsComIDNO As String = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, objconn)
        If TIMS.Check_OrgBlackList(Me, vsComIDNO, objconn) Then
            rst = True
            Errmsg = sm.UserInfo.OrgName & "，已列入處分名單!!"
            Me.isBlack.Value = "Y"
            Me.Blackorgname.Value = sm.UserInfo.OrgName
        End If
        Return rst
    End Function

    '查詢 'SQL
    Sub Search1()
        lab_msg1.Text = ""
        table4.Visible = False
        Datagrid2.Visible = False
        btnsave.Visible = False '備取設定的儲存
        btncancel.Visible = False '備取設定的取消
        btnSEND1.Enabled = True
        'btnSAVE1.Enabled=True

        Call TIMS.OpenDbConn(objconn)  ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub 'SESSION不完整

        'If Hid_DG_SORT1.Value="" Then
        '    Select Case TIMS.GetListValue(rblsortmode)'.SelectedValue
        '        Case "1"
        '            Hid_DG_SORT1.Value=cst_TotalResult & cst_Desc
        '        Case "2"
        '            Hid_DG_SORT1.Value=cst_ExamNo
        '    End Select
        'End If

        Hid_StudTNum.Value = "0"
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If OCIDValue1.Value = "" Then
            Common.MessageBox(Me, "班級選擇有誤!!")
            Exit Sub
        End If
        ' If OCIDValue1.Value="" Then Exit Sub
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "班級選擇有誤!!")
            Exit Sub
        End If

        'If Convert.ToString(sm.UserInfo.Years)="" Then Exit Sub
        'Dim dtStud As New DataTable
        Dim strMsg As String = ""
        Hid_OCID1.Value = OCIDValue1.Value

        '如果是產學訓就要檢查是否尚有e網報名未審學員
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '檢查是否尚有e網報名未審學員
            Dim flag_check_et2 As Boolean = Chk_EsignUpStatus(OCIDValue1.Value, objconn)
            If flag_check_et2 Then
                '如果有e網報名未審學員就顯示訊息
                strMsg = cst_msg0014 '"目前尚有e網報名學員尚未審核，需先將【e網報名】的所有學員審核完成後，始能登錄【錄取作業】!"
                btnSEND1.Enabled = False '(完成錄取) '失效
                TIMS.Tooltip(btnSEND1, strMsg, True)
                Common.MessageBox(Me, strMsg)
                Exit Sub
            End If
        End If

        'https://jira.turbotech.com.tw/browse/TIMSC-153
        'If TIMS.Chk_DISASTERNG(Me, OCIDValue1.Value, objconn) Then
        '    '針對”重大災害提醒處理”功能中，若有未進行提醒處理者，則無法進行報到作業，並顯示告警訊息。
        '    strMsg=cst_msg0015 '"針對【重大災害提醒處理】功能，有未進行提醒處理，無法進行該作業，請先完成 重大災害提醒處理!"
        '    Common.MessageBox(Me, strMsg)
        '    Exit Sub
        'End If

        '過N關 存 ViewState(cst_OCIDValue1) 
        ViewState(cst_OCIDValue1) = OCIDValue1.Value
        Hid_OCID1.Value = OCIDValue1.Value

        'Dim drItemVar As DataRow=Nothing '取成績計算比例
        Dim iSort As Integer = 0 '學員人數
        argrole.Text = ""
        DataGrid1.Columns(cst_ceWriteResult).Visible = False '筆試
        DataGrid1.Columns(cst_ceOralResult).Visible = False '口試
        DataGrid1.Columns(cst_ceTotalResult).Visible = False '總成績
        DataGrid1.Columns(cst_ceEnterDate).Visible = False  '報名日期

        Dim dtChk As DataTable = Nothing 'GET_SELRESULTdt(OCIDValue1.Value, objconn)
        Dim flag_can_show As Boolean = False '不可顯示
        '接受企業委託訓練
        If TIMS.Cst_TPlanID07.IndexOf(sm.UserInfo.TPlanID) > -1 Then flag_can_show = True '可顯示

        If Not flag_can_show Then
            dtChk = GET_SELRESULTdt(OCIDValue1.Value, objconn)
            strMsg = ""
            If dtChk.Rows.Count = 0 Then
                '沒資料，可能問題-TIMS-筆數為0
                strMsg = ""
                strMsg += "尚有流程未完成，可能原因如下：" & vbCrLf
                strMsg += "1.尚未有報名且審核成功的學員" & vbCrLf
                strMsg += "2.尚有成績未登錄(如免試可直接跳過做「甄試結果試算」即可)" & vbCrLf
                strMsg += "3.尚未執行甄試結果試算" & vbCrLf
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    '沒資料，可能問題-產投-筆數為0
                    strMsg = ""
                    strMsg += "尚有流程未完成，可能原因如下：" & vbCrLf
                    strMsg += "1.尚未有報名且審核成功的學員" & vbCrLf
                End If
                table4.Visible = False
                Common.MessageBox(Me, strMsg)
                Exit Sub
            End If
        End If

        'Dim sql As String=""
        If dtChk IsNot Nothing Then
            iSort = TIMS.CINT1(dtChk.Rows(0)("total")) '有填 甄試成績  /產投沒有 甄試成績
        End If
        If iSort > 0 Then
            '加入只能查登入年度的年度限制。
            'NotExam 是否免試
            'TRNDType Dtype 職訓卷種類
            'SumOfGrad	甄試成績 
            'ExamNo	證號-准考證號碼

            '有填 甄試成績 

            'NotExam 是否免試
            'TRNDType Dtype 職訓卷種類
            'SumOfGrad	甄試成績 
            'ExamNo	證號
            DataGrid1.Columns(cst_ceWriteResult).Visible = True '筆試
            DataGrid1.Columns(cst_ceOralResult).Visible = True '口試
            DataGrid1.Columns(cst_ceTotalResult).Visible = True '總成績
        Else
            '未填 甄試成績 
            DataGrid1.Columns(cst_ceEnterDate).Visible = True '報名日期
        End If

        'Call TIMS.OpenDbConn(objconn)
        'parms.Clear()
        Dim parms As Hashtable = New Hashtable From {{"CLASS1", OCIDValue1.Value}, {"OCID", OCIDValue1.Value}, {"YEARS", Right(sm.UserInfo.Years, 2)}}

        Dim ORDER_sql As String = "ORDER BY a.TotalResult DESC,a.RelEnterDate,a.ExamNo"
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then ORDER_sql = "ORDER BY a.NotExam DESC,d2.SIGNNO"

        Dim sql As String = ""
        sql &= " SELECT a.SETID, CONVERT(VARCHAR, a.EnterDate, 111) EnterDate ,a.SerNum" & vbCrLf
        sql &= " ,a.ExamNo ,a.WriteResult ,a.OralResult ,a.TotalResult ,a.NotExam" & vbCrLf
        sql &= " ,b.notes2,format(b.MODIFYDATE,'yyyy-MM-dd HH:mm') MODIFYDATE" & vbCrLf
        sql &= " ,b.OCID ,b.SumOfGrad , b.SelResultID ,b.Admission" & vbCrLf
        sql &= " ,ISNULL(d.Name, '') Name ,d.IDNO" & vbCrLf
        sql &= " ,CASE WHEN b.selsort IS NOT NULL THEN concat('備取',b.selsort) ELSE '未設定' END selsort" & vbCrLf
        sql &= $" ,ROW_NUMBER() OVER ({ORDER_sql}) RSort" & vbCrLf '名次 'sql &= " ,0 RSort" & vbCrLf '名次
        sql &= " ,CASE WHEN a.EnterPath='W' THEN 1 ELSE 2 END WSort" & vbCrLf
        sql &= " ,'' GOVKILL" & vbCrLf
        sql &= " ,d2.SIGNNO" & vbCrLf
        sql &= " FROM STUD_SELRESULT b" & vbCrLf
        sql &= " JOIN STUD_ENTERTYPE a ON a.SETID=b.SETID AND a.EnterDate=b.EnterDate AND a.SerNum=b.SerNum AND a.OCID1=b.OCID" & vbCrLf
        sql &= " JOIN STUD_ENTERTEMP d ON d.SETID=a.SETID" & vbCrLf
        sql &= " JOIN CLASS_CLASSINFO e ON e.OCID=b.OCID" & vbCrLf
        sql &= " LEFT JOIN STUD_ENTERTYPE2 d2 on d2.SETID=b.SETID AND d2.EnterDate=b.EnterDate AND d2.SerNum=b.SerNum AND d2.OCID1=b.OCID" & vbCrLf
        sql &= " WHERE e.OCID=@OCID AND e.Years=@YEARS" & vbCrLf 'and a.CCLID IS NULL '加入只能查登入年度的年度限制。
        sql &= $"{ORDER_sql}" & vbCrLf

        DataGrid1.Columns(cst_ceExamNo).Visible = True  '自辦用准考證序號
        DataGrid1.Columns(cst_ceSIGNNO).Visible = False  '產投用e網報名序號
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            DataGrid1.Columns(cst_ceExamNo).Visible = False  '自辦用准考證序號
            DataGrid1.Columns(cst_ceSIGNNO).Visible = True  '產投用e網報名序號
        End If

        'Dim sCmd3 As New SqlCommand(sql, objconn)
        Dim dtStud As DataTable = Nothing
        Try
            dtStud = DbAccess.GetDataTable(sql, objconn, parms)
        Catch ex As Exception
            Dim strErrmsg As String = ""
            strErrmsg += "/*  ex.ToString: */" & vbCrLf
            strErrmsg += ex.ToString & vbCrLf
            strErrmsg += "sql.Value:" & sql & vbCrLf
            strErrmsg += "OCIDValue1.Value:" & OCIDValue1.Value & vbCrLf
            strErrmsg += "years.Value:" & Right(sm.UserInfo.Years, 2) & vbCrLf
            strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)
            Common.MessageBox(Me, TIMS.cst_NODATAMsg2)
            Return ' Exit Sub 'Throw ex
            'Common.MessageBox(Me, ex.ToString)
        End Try

        Hid_StudTNum.Value = If(TIMS.dtHaveDATA(dtStud), dtStud.Rows.Count, 0)

        lab_msg1.Visible = True
        lab_msg1.Text = "查無資料!"
        table4.Visible = False
        Datagrid2.Visible = False
        btnsave.Visible = False '備取設定的儲存
        btncancel.Visible = False '備取設定的取消

        If TIMS.dtHaveDATA(dtStud) Then
            For Each dc As DataColumn In dtStud.Columns
                dc.ReadOnly = False
            Next

            'Dim i As Integer=0
            'Dim dr1 As DataRow=dtStud.Rows(0)
            'For Each dr As DataRow In dtStud.Rows
            '    i += 1
            '    dr("RSort")=i '(排序)
            'Next
            'dtStud.AcceptChanges()
            'msg.Visible=True
            lab_msg1.Text = ""
            table4.Visible = True
            'Datagrid2.Visible=False
            'btnsave.Visible=False '備取設定的儲存
            'btncancel.Visible=False '備取設定的取消
            'SORT ORDER BY 
            If Hid_DG_SORT1.Value <> "" Then dtStud.DefaultView.Sort = Hid_DG_SORT1.Value
            DataGrid1.DataSource = dtStud.DefaultView.Table
            DataGrid1.DataBind()

            If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) = -1 Then
                '非產投
                '取得成績計算比例,ItemVarFlag(0=>全計畫設定, 1=>只單一班設定)
                'DataGridTable.Visible=False
                argrole.Text = ""
                Dim inParms As New Hashtable
                inParms.Add("RIDValue", Convert.ToString(drCC("RID")))
                inParms.Add("PlanID", Convert.ToString(drCC("PlanID")))
                inParms.Add("DistID", Convert.ToString(drCC("DistID")))
                inParms.Add("TPLANID", Convert.ToString(drCC("TPLANID")))
                Dim outParms As New Hashtable
                If Not TIMS.getItemVar23(sm, Me, objconn, inParms, outParms) Then
                    Dim s_ErrorMsg1 As String = TIMS.GetMyValue2(outParms, "ErrorMsg1")
                    Common.MessageBox(Me, s_ErrorMsg1)
                    'Exit Sub
                End If
                ItemVar1.Value = TIMS.GetMyValue2(outParms, "ItemVar1")
                ItemVar2.Value = TIMS.GetMyValue2(outParms, "ItemVar2")
                argrole.Text = TIMS.GetMyValue2(outParms, "ArgRole")

                'drItemVar=getItemVar23(objconn)
                'argrole.Text=""
                'If drItemVar Is Nothing Then
                '    s_Msg1=""
                '    s_Msg1 &= "系統尚未設定筆試與口試的參數,請聯絡系統管理員!!" & vbCrLf & "該業務可能不屬於該登入使用者，請勿任意存取!!" & vbCrLf
                '    Common.MessageBox(Me, s_Msg1)
                '    'Common.MessageBox(Me, "系統尚未設定筆試與口試的參數,請聯絡系統管理員!!")
                '    'Common.MessageBox(Me, "該業務可能不屬於該登入使用者，請勿任意存取!!")
                '    argrole.Text=""
                'End If
                'If drItemVar IsNot Nothing Then End If
                'argrole.Text="[成績計算方式：(筆試*" & drItemVar(0) & "%)+(口試*" & drItemVar(1) & "%)=總成績]"
                '屆退官兵者 (依系統日期判斷)
                'If flagTPlanID02Plan2 Then
                '    argrole.Text="成績計算方式：[(筆試*" & ItemVar1.Value & "%)+(口試*" & ItemVar2.Value & "%)]*3%=總成績，總成績最高為100分。<br /> 「*」為屆退官兵身分者"
                'End If
                TIMS.Tooltip(argrole, cst_tip2)
            End If
        End If

        'CLASS_CONFIRM
        Dim drCF As DataRow = TIMS.GET_CLSCONFIRM(OCIDValue1.Value, objconn)
        If drCF IsNot Nothing Then
            '有資料
            Hid_OCID1.Value = CStr(drCF("OCID"))
            Hid_CFSEQNO.Value = CStr(drCF("CFSEQNO"))
            Hid_CFGUID.Value = CStr(drCF("CFGUID"))
            'flag_USE_NEW_CONFIRM=False
        End If

        '#Region "(No Use)"

        'button6.Visible=True
        'ptitle6.Visible=True 'Title文字一起消失吧
        'If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '    button6.Visible=False '產業人才投資方案，不顯示e網公告的按鈕
        '    ptitle6.Visible=False 'Title文字一起消失吧
        'End If

        classname.Text = ""
        If drCC IsNot Nothing Then
            classname.Text = String.Format("{0}，訓練人數:{1}人", drCC("ClassCName2"), drCC("TNum"))
        End If

        '按過後鎖定 true:鎖定  false:不鎖定
        'https://jira.turbotech.com.tw/browse/TIMSC-255
        'Dim drCF As DataRow=TIMS.CHK_CONFIRM_NOLOCK(OCIDValue1.Value, objconn)
        'If Not drCF Is Nothing Then
        '    Hid_OCID1.Value=CStr(drCF("OCID"))
        '    Hid_CFSEQNO.Value=CStr(drCF("CFSEQNO"))
        '    Hid_CFGUID.Value=CStr(drCF("CFGUID"))
        '    btnSEND1.Enabled=False
        '    TIMS.Tooltip(btnSEND1, cst_msg0007, True)
        '    btnSAVE1.Enabled=False
        '    TIMS.Tooltip(btnSAVE1, cst_msg0007, True)
        '    TIMS.IsSuperUser(1)
        '    If flgROLEIDx0xLIDx0 Then btnSAVE2.Visible=True '有權限者可解鎖
        '    Common.MessageBox(Me, cst_msg0007)
        '    Exit Sub
        'End If

        'Try
        'Catch ex As Exception
        '    Dim strErrmsg As String=""
        '    strErrmsg += "/*  ex.ToString: */" & vbCrLf
        '    strErrmsg += ex.ToString & vbCrLf
        '    strErrmsg += "OCIDValue1.Value:" & OCIDValue1.Value & vbCrLf
        '    strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
        '    strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
        '    Call TIMS.SendMailTest(strErrmsg)
        '    'Common.MessageBox(Me, ex.ToString)
        'End Try
    End Sub

    '查詢
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button1.Click
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        ViewState(cst_OCIDValue1) = ""
        Hid_OCID1.Value = ""
        'Hid_DG_SORT1.Value=""
        Select Case TIMS.GetListValue(rblsortmode)'.SelectedValue
            Case "1"
                Hid_DG_SORT1.Value = cst_TotalResult & cst_Desc
            Case "2"
                Hid_DG_SORT1.Value = cst_ExamNo
        End Select
        btnSEND1.Enabled = True

        Call Search1()
    End Sub

    '檢查學員是否存在
    Function CheckStudentsOfClass(ByVal OCID As String, ByVal IDNO As String) As Boolean
        Dim rst As Boolean = False 'FALSE:未存在學員／TRUE: 有存在學員
        Dim pms_1 As New Hashtable From {{"OCID", TIMS.CINT1(OCID)}, {"IDNO", IDNO}}
        Dim sql As String = ""
        sql &= " SELECT cs.SOCID" & vbCrLf
        sql &= " FROM dbo.CLASS_STUDENTSOFCLASS cs WITH(NOLOCK)" & vbCrLf
        sql &= " JOIN dbo.STUD_STUDENTINFO ss WITH(NOLOCK) ON cs.SID=ss.SID" & vbCrLf
        sql &= " WHERE cs.OCID=@OCID AND ss.IDNO=@IDNO" & vbCrLf
        Dim dr1 As DataRow = DbAccess.GetOneRow(sql, objconn, pms_1)
        If dr1 Is Nothing Then Return rst
        Return (Convert.ToString(dr1("SOCID")) <> "")
    End Function

    '取得成績計算比例
    'Private Function getItemVar23(ByVal tConn As SqlConnection) As DataRow
    '    Dim drRtn As DataRow=Nothing
    '    Try
    '        Dim sql As String=""
    '        Dim sda As New SqlDataAdapter
    '        Dim dt As New DataTable
    '        Dim iVal As String=""
    '        iVal=TIMS.GetGlobalVar(Me, "23", "1", objconn)
    '        If iVal="Y" Then
    '            '開放設定,取機構(班級)設定資料
    '            sql=""
    '            sql &= " SELECT writeresult, oralresult FROM org_writeoral WHERE planid=@planid "
    '            sql &= " AND orgid IN (SELECT a.orgid FROM auth_relship a WHERE a.rid=@rid) AND (ocid=@ocid) ORDER BY ocid "
    '            With sda
    '                .SelectCommand=New SqlCommand(sql, tConn)
    '                .SelectCommand.Parameters.Clear()
    '                .SelectCommand.Parameters.Add("@planid", SqlDbType.VarChar).Value=sm.UserInfo.PlanID
    '                .SelectCommand.Parameters.Add("@rid", SqlDbType.VarChar).Value=RIDValue.Value
    '                .SelectCommand.Parameters.Add("@ocid", SqlDbType.VarChar).Value=OCIDValue1.Value
    '                dt=New DataTable
    '                .Fill(dt)
    '            End With
    '            If dt.Rows.Count=0 Then
    '                '開放設定,取機構設定資料
    '                sql=""
    '                sql &= " SELECT writeresult, oralresult FROM org_writeoral WHERE planid=@planid "
    '                sql &= " AND orgid IN (SELECT a.orgid FROM auth_relship a WHERE a.rid=@rid) AND (ocid IS NULL) ORDER BY ocid "
    '                With sda
    '                    .SelectCommand=New SqlCommand(sql, tConn)
    '                    .SelectCommand.Parameters.Clear()
    '                    .SelectCommand.Parameters.Add("@planid", SqlDbType.VarChar).Value=sm.UserInfo.PlanID
    '                    .SelectCommand.Parameters.Add("@rid", SqlDbType.VarChar).Value=RIDValue.Value
    '                    dt=New DataTable
    '                    .Fill(dt)
    '                End With
    '            End If
    '            If dt.Rows.Count > 0 Then drRtn=dt.Rows(0)
    '        Else
    '            Dim wVal As String=TIMS.GetGlobalVar(Me, "2", "1", objconn)
    '            Dim oVal As String=TIMS.GetGlobalVar(Me, "2", "2", objconn)
    '            'sql="select '" & wVal & "' writeresult, '" & oVal & "' oralresult" 
    '            sql="SELECT '" & wVal & "' writeresult, '" & oVal & "' oralresult  "
    '            With sda
    '                .SelectCommand=New SqlCommand(sql, tConn)
    '                .SelectCommand.Parameters.Clear()
    '                dt=New DataTable
    '                .Fill(dt)
    '            End With
    '            If dt.Rows.Count > 0 Then drRtn=dt.Rows(0)
    '        End If
    '        If Not sda Is Nothing Then sda.Dispose()
    '        If Not dt Is Nothing Then dt.Dispose()
    '    Catch ex As Exception
    '        Dim strErrmsg As String=""
    '        strErrmsg += "/*  ex.ToString: */" & vbCrLf
    '        strErrmsg += ex.ToString & vbCrLf
    '        strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
    '        strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
    '        Call TIMS.WriteTraceLog(strErrmsg)
    '        Common.MessageBox(Me, ex.ToString)
    '        Throw ex
    '    End Try
    '    Return drRtn
    'End Function

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                If Hid_DG_SORT1.Value <> "" Then
                    Dim img As New UI.WebControls.Image
                    Select Case Hid_DG_SORT1.Value
                        Case cst_SIGNNO, cst_ExamNo, cst_TotalResult, cst_examno2, cst_totalresult2, cst_rsort
                            img.ImageUrl = "../../images/SortUp.gif"
                        Case cst_SIGNNO & cst_Desc, cst_ExamNo & cst_Desc, cst_TotalResult & cst_Desc, cst_examno2 & cst_Desc, cst_totalresult2 & cst_Desc, cst_rsort & cst_Desc
                            img.ImageUrl = "../../images/SortDown.gif"
                    End Select
                    Select Case Hid_DG_SORT1.Value
                        Case cst_ExamNo, cst_ExamNo & cst_Desc
                            e.Item.Cells(cst_ceExamNo).Controls.Add(img)
                        Case cst_examno2, cst_examno2 & cst_Desc
                            e.Item.Cells(cst_ceExamNo).Controls.Add(img)
                        Case cst_SIGNNO, cst_SIGNNO & cst_Desc
                            e.Item.Cells(cst_ceSIGNNO).Controls.Add(img)
                        Case cst_TotalResult, cst_TotalResult & cst_Desc, cst_totalresult2, cst_totalresult2 & cst_Desc
                            e.Item.Cells(cst_ceTotalResult).Controls.Add(img)
                        Case cst_rsort, cst_rsort & cst_Desc
                            e.Item.Cells(cst_ceRSort).Controls.Add(img)
                    End Select
                End If
                e.Item.Cells(cst_ceselsort).Visible = False
                If Hid_ChangeWating.Value = "Y" Then e.Item.Cells(cst_ceselsort).Visible = True '能

                e.Item.Cells(cst_cenotes2).Visible = False '備註
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    e.Item.Cells(cst_cenotes2).Visible = True
                    TIMS.Tooltip(e.Item.Cells(cst_cenotes2), "該欄位資訊外網不顯示")
                End If
                'Header Text
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then e.Item.Cells(cst_ceddlResult).Text = "錄訓結果"

            Case ListItemType.AlternatingItem, ListItemType.EditItem, ListItemType.Item, ListItemType.SelectedItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim hidIDNO As HtmlInputHidden = e.Item.FindControl("hidIDNO")

                Dim labPathW As Label = e.Item.FindControl("labPathW")
                Dim labname As Label = e.Item.FindControl("labname")
                Dim labstarSOLDER As Label = e.Item.FindControl("labstarSOLDER")
                'ddlResult
                Dim notes2 As TextBox = e.Item.FindControl("notes2") '備註  'notes2 備取或未錄取原因
                Dim ddlResult As DropDownList = e.Item.FindControl("ddlResult") 'e.Item.Cells(8).Controls(1)
                ddlResult = TIMS.Get_SelResult(ddlResult, 2, objconn) '甄試結果 '增加審核中-'01','02','03','05'

                Dim Hid_rsort As HiddenField = e.Item.FindControl("Hid_rsort") 'rsort 名次
                Dim Hid_SETID As HiddenField = e.Item.FindControl("Hid_SETID")
                Dim Hid_EnterDate As HiddenField = e.Item.FindControl("Hid_EnterDate")
                Dim Hid_SerNum As HiddenField = e.Item.FindControl("Hid_SerNum")
                Hid_rsort.Value = Convert.ToString(drv("rsort"))
                Hid_SETID.Value = Convert.ToString(drv("SETID"))
                Hid_EnterDate.Value = Convert.ToString(drv("EnterDate"))
                Hid_SerNum.Value = Convert.ToString(drv("SerNum"))
                labname.Text = Convert.ToString(drv("name"))
                hidIDNO.Value = Convert.ToString(drv("IDNO"))
                labstarSOLDER.Visible = False

                'If flagTPlanID02Plan2 Then
                '    '屆退官兵者 (依系統日期判斷)
                '    labstarSOLDER.Visible=True
                '    labstarSOLDER.Text=""
                '    If TIMS.CheckRESOLDER(objconn, hidIDNO.Value, sm.UserInfo.DistID, "") Then labstarSOLDER.Text="*"
                'End If
                'WSort 1.就服單位協助報名
                labPathW.Visible = False
                Select Case Convert.ToString(drv("WSort"))
                    Case "1"
                        labPathW.Visible = True
                        labPathW.Text = cst_EnterPathNameW
                End Select
                'e.Item.Cells(Cst_ceTRNDType).Text="-"
                'If Convert.ToString(drv("TRNDType")) <> "" Then
                '    Select Case Convert.ToString(drv("TRNDType"))  'e.Item.Cells(Cst_ceTRNDType).Text
                '        Case "1"
                '            e.Item.Cells(Cst_ceTRNDType).Text="甲式"
                '        Case "2"
                '            e.Item.Cells(Cst_ceTRNDType).Text="乙式"
                '        Case "3"
                '            e.Item.Cells(Cst_ceTRNDType).Text="推介單" '星光幫專用（青年職涯啟動計劃）
                '            'e.Item.Cells(Cst_ceTRNDType).ToolTip=drv("PlanID").ToString '被使用，請勿更改寫法
                '    End Select
                'End If
                notes2.Text = Convert.ToString(drv("notes2")) '備註
                Dim v_SelResultID As String = "" 'ddlResult 01:正取 02:備取 03:未錄取
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    '產投-'尚未做選擇時，預設為審核中
                    v_SelResultID = TIMS.cst_SelResultID_審核中 '"03"
                    If Convert.ToString(drv("SelResultID")) <> "" Then v_SelResultID = Convert.ToString(drv("SelResultID"))
                    'Common.SetListItem(ddlResult, Convert.ToString(drv("SelResultID")))
                    'If Convert.ToString(drv("Admission"))="Y" Then Result.SelectedIndex=0
                Else
                    'TIMS-'尚未做選擇時，預設為未錄取
                    v_SelResultID = TIMS.cst_SelResultID_未錄取 '"03"
                    If Convert.ToString(drv("SelResultID")) <> "" Then v_SelResultID = Convert.ToString(drv("SelResultID"))
                End If
                Common.SetListItem(ddlResult, v_SelResultID)

                Dim s_title_msg1 As String = ""
                '檢查學員是否存在 '且不等於未錄取
                If CheckStudentsOfClass(drv("OCID"), drv("idno")) Then
                    '有建立學員資料
                    'Const Cst_vTitle1 As String="學員資料已經存在!!"
                    'Const Cst_vTitle As String="請設定為正取,學員資料已經存在!!"
                    'ddlResult 01:正取 02:備取 03:未錄取
                    Select Case TIMS.GetListValue(ddlResult)'.SelectedValue '01:正取 02:備取 03:未錄取
                        Case TIMS.cst_SelResultID_正取 '"01"
                            ddlResult.Enabled = False
                            s_title_msg1 = String.Concat(cst_msg0001, ",最後異動日：", drv("MODIFYDATE"))
                        Case TIMS.cst_SelResultID_備取 '"02" '備取
                            If Not ddlResult.Enabled = False Then ddlResult.Enabled = True
                            ddlResult.ForeColor = Color.DarkBlue
                            s_title_msg1 = String.Concat(cst_msg0003, ",最後異動日：", drv("MODIFYDATE"))
                        Case TIMS.cst_SelResultID_未錄取 '"03" '未錄取
                            If Not ddlResult.Enabled = False Then ddlResult.Enabled = True
                            ddlResult.ForeColor = Color.DarkBlue
                            s_title_msg1 = String.Concat(cst_msg0004, ",最後異動日：", drv("MODIFYDATE"))
                        Case Else
                            '無值或其他狀況！
                            If Not ddlResult.Enabled = False Then ddlResult.Enabled = True
                            ddlResult.ForeColor = Color.DarkBlue
                            s_title_msg1 = String.Concat(cst_msg0004, ",最後異動日：", drv("MODIFYDATE"))
                    End Select
                Else
                    '未建立學員資料
                    'Const Cst_vTitle As String="已設定為正取，但該學員資料尚未建立!"
                    '01:正取 02:備取 03:未錄取
                    Select Case TIMS.GetListValue(ddlResult)'.SelectedValue
                        Case TIMS.cst_SelResultID_正取 '"01" '正取
                            ddlResult.ForeColor = Color.DarkBlue
                            If Convert.ToString(drv("MODIFYDATE")) <> "" Then
                                s_title_msg1 = String.Concat(cst_msg0002, ",最後異動日：", drv("MODIFYDATE"))
                            Else
                                s_title_msg1 = cst_msg0002
                            End If
                    End Select
                End If
                If s_title_msg1 <> "" Then
                    TIMS.Tooltip(e.Item.Cells(cst_ceddlResult), s_title_msg1, True)
                    TIMS.Tooltip(e.Item, s_title_msg1, True)
                    TIMS.Tooltip(ddlResult, s_title_msg1, True)
                End If

                If Convert.ToString(drv("GOVKILL")) = "Y" Then
                    Common.SetListItem(ddlResult, TIMS.cst_SelResultID_未錄取) '"03"
                    ddlResult.Enabled = False
                    'e.Item.Enabled=False
                    TIMS.Tooltip(e.Item, cst_Mgc219)
                End If
                '因為不會存取，所以要開放同等屬性
                notes2.Enabled = ddlResult.Enabled '備註
                'If CheckStudentsOfClass(drv("OCID"), drv("idno")) _
                '    AndAlso Result.SelectedValue <> "03" Then
                '    Result.Enabled=False
                '    TIMS.Tooltip(Result, "學員資料已經存在不可修改", True)
                'End If
                e.Item.Cells(cst_ceselsort).Visible = False
                If Hid_ChangeWating.Value = "Y" Then e.Item.Cells(cst_ceselsort).Visible = True
                e.Item.Cells(cst_cenotes2).Visible = False
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    e.Item.Cells(cst_cenotes2).Visible = True
                    TIMS.Tooltip(e.Item.Cells(cst_cenotes2), "該欄位資訊外網不顯示")
                End If
        End Select
    End Sub

    Private Sub DataGrid1_SortCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridSortCommandEventArgs) Handles DataGrid1.SortCommand
        If Hid_DG_SORT1.Value <> "" Then
            Hid_DG_SORT1.Value = If(Hid_DG_SORT1.Value.IndexOf(cst_Desc) > -1, e.SortExpression, e.SortExpression & cst_Desc)
        End If

        Call Search1()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn) '判斷機構是否只有一個班級
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        table4.Visible = False
        lab_msg1.Visible = False
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        '如果只有一個班級
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
        table4.Visible = False
        lab_msg1.Visible = False
    End Sub

    Private Sub Datagrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles Datagrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.EditItem, ListItemType.Item, ListItemType.SelectedItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim DrpWatingS As DropDownList = e.Item.FindControl("DrpWatingS")
                Dim Hid_SETID As HiddenField = e.Item.FindControl("Hid_SETID")
                Hid_SETID.Value = Convert.ToString(drv("SETID"))
                If ViewState(cst_watingcount) > 0 Then
                    For i As Integer = 0 To ViewState(cst_watingcount)
                        If i = 0 Then
                            DrpWatingS.Items.Add(New ListItem(TIMS.cst_ddl_PleaseChoose2, ""))
                        Else
                            DrpWatingS.Items.Add(New ListItem(i, i))
                        End If
                    Next
                End If
                If Convert.ToString(drv("selsort")) <> "" Then   '代資料
                    ' DrpWatingS.SelectedValue=Convert.ToString(drv("selsort"))
                    Common.SetListItem(DrpWatingS, Convert.ToString(drv("selsort")))
                End If
        End Select
    End Sub

    '備取名次存檔(儲存)
    Private Sub BtnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnsave.Click
        '#Region "(No Use)"

        ''Dim OKflag As Boolean=False
        'Dim sqlAdp As New SqlDataAdapter
        'Dim sqlStr As String=""
        ''OKflag=False
        ''Dim Ditem As DataGridItem
        ''Dim Ditem2 As DataGridItem
        'Dim SETID As String=""
        'Dim WatingSort As DropDownList
        'Dim strALL As String=""
        'Dim strS As String=""
        'Dim i As Integer=1
        ''-------------判斷備取名次設定是否有重復------------------
        'For Each Ditem In Datagrid2.Items
        '    WatingSort=Ditem.Cells(2).Controls(1)
        '    If WatingSort.SelectedIndex > 0 Then   '排除沒選名次的
        '        strS="'" & WatingSort.SelectedValue & "',"   '取備取名次的值
        '        If i=1 Then        '第一筆資料
        '            strALL += "'" & WatingSort.SelectedValue & "',"   '比對的字串
        '        ElseIf i >= 2 Then   '第二筆資料才開始比對
        '            If InStr(strALL, strS) <= 0 Then    '沒有找到相同的
        '                strALL += "'" & WatingSort.SelectedValue & "',"    '比對的字串
        '            Else                                '有找到相同的
        '                Common.MessageBox(Me, "備取名次設定不能重覆!!")
        '                Exit Sub
        '            End If
        '        End If
        '        i=i + 1       '計算次數
        '    End If
        'Next
        ''-------------判斷備取名次設定是否有重復-end----------------

        ''判斷備取名次設定是否有重復 'False:沒有重複，True:重複
        'If sUtl_DoubleCheck1() Then
        '    Common.MessageBox(Me, "備取名次設定不能重覆!!")
        '    Exit Sub
        'End If


        'Dim sqlStr As String=""
        'Dim sqlAdp As New SqlDataAdapter
        For Each eItem As DataGridItem In Datagrid2.Items
            'Dim WatingSort As DropDownList
            'Dim SETID As String=""
            'WatingSort=Ditem.Cells(2).Controls(1)
            'SETID=Ditem.Cells(3).Text
            'DrpWatingS
            Dim DrpWatingS As DropDownList = eItem.FindControl("DrpWatingS")
            Dim Hid_SETID As HiddenField = eItem.FindControl("Hid_SETID")
            Dim v_DrpWatingS As String = TIMS.GetListValue(DrpWatingS)
            Dim sqlStr As String = ""
            sqlStr &= " UPDATE STUD_SELRESULT" & vbCrLf
            sqlStr &= " SET selsort=@selsort" & vbCrLf
            sqlStr &= " WHERE OCID=@OCIDValue1 AND SETID=@SETID" & vbCrLf
            Dim myParam As Hashtable = New Hashtable
            myParam.Add("OCIDValue1", ViewState(cst_OCIDValue1))
            myParam.Add("selsort", If(v_DrpWatingS <> "", Val(v_DrpWatingS), Convert.DBNull)) '.SelectedValue))
            myParam.Add("SETID", Hid_SETID.Value)
            DbAccess.ExecuteNonQuery(sqlStr, objconn, myParam)
        Next

        Common.MessageBox(Me, cst_msg0006)
        Call Search1()
    End Sub

    '備取名次取消(取消)
    Private Sub BtnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btncancel.Click
        Common.MessageBox(Me, cst_msg0005)
        'Button1_Click(sender, e)
        Call Search1()
        'Table3.Visible=True
        'Table4.Visible=True
        'Datagrid2.Visible=False
        'BtnSave.Visible=False '備取設定的儲存
        'BtnCancel.Visible=False '備取設定的取消
    End Sub

    'UPDATE_STUD_SELRESULT(objconn, htss,drocid) (此FUNC 已經包在TRY CATCH)
    Sub UPDATE_STUD_SELRESULT(ByVal oTrans As SqlTransaction, ByRef htSS As Hashtable, ByRef drOC As DataRow, ByRef rErrMsg As String)
        'UPDATE STUD_SELRESULT
        'UPDATE STUD_ENTERTYPE2
        Dim strSetId As String = TIMS.GetMyValue2(htSS, "strSetId")
        Dim strSerNum As String = TIMS.GetMyValue2(htSS, "strSerNum")
        Dim strEnterDate As String = TIMS.GetMyValue2(htSS, "strEnterDate")
        Dim sNotes2 As String = TIMS.GetMyValue2(htSS, "sNotes2") 'notes2.Text
        Dim sAdmission As String = TIMS.GetMyValue2(htSS, "sAdmission") 'sAdmission Y/N
        Dim sResultSelValue As String = TIMS.GetMyValue2(htSS, "sResultSelValue") 'Result.SelectedValue 
        rErrMsg = ""
        'Dim objTrans As SqlTransaction=Nothing
        'objTrans=DbAccess.BeginTrans(tConn)
        Dim sql As String = ""
        sql = " SELECT * FROM STUD_SELRESULT WHERE SETID='" & strSetId & "' AND SerNum='" & strSerNum & "' AND EnterDate=" & TIMS.To_date(strEnterDate)
        Dim da As SqlDataAdapter = Nothing
        Dim dr As DataRow = Nothing
        Dim dt As DataTable = DbAccess.GetDataTable(sql, da, oTrans)
        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)
            If IsDBNull(dr("AppliedStatus")) Then dr("AppliedStatus") = "N" '是否報到
        Else
            dr = dt.NewRow
            dt.Rows.Add(dr)
            dr("SETID") = strSetId
            dr("EnterDate") = CDate(Common.FormatDate(strEnterDate)) ', DateFormat.ShortDate)
            dr("SerNum") = strSerNum
            dr("OCID") = ViewState(cst_OCIDValue1)
            'dr("SumOfGrad")=""
            dr("AppliedStatus") = "N" '是否報到:未報到
            'dr("TRNDType")="3" '職訓卷種類
        End If
        dr("RID") = drOC("RID") 'RIDValue.Value
        dr("PLANID") = drOC("PlanID")
        dr("notes2") = If(sNotes2 <> "", sNotes2, Convert.DBNull) '備註
        dr("selsort") = If(sResultSelValue <> "02", Convert.DBNull, dr("selsort")) '(CLEAR@selsort)
        dr("Admission") = sAdmission 'Y/N
        '甄試結果代碼(未選擇):0
        dr("SelResultID") = If(sResultSelValue <> "", sResultSelValue, "0")
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now()
        DbAccess.UpdateDataTable(dt, da, oTrans)

        '2007/11/14 Ellen E網報名結果
        sql = ""
        sql &= " SELECT * FROM STUD_ENTERTYPE2" & vbCrLf
        sql &= " WHERE OCID1='" & ViewState(cst_OCIDValue1) & "'" & vbCrLf
        sql &= " AND SETID='" & strSetId & "'" & vbCrLf
        sql &= " AND SerNum='" & strSerNum & "'" & vbCrLf
        sql &= " AND EnterDate=" & TIMS.To_date(Common.FormatDate(strEnterDate)) & vbCrLf
        Dim dt1 As DataTable = Nothing
        dt1 = DbAccess.GetDataTable(sql, da, oTrans)
        If dt1.Rows.Count = 0 Then Return

        Dim dr1 As DataRow = dt1.Rows(0)
        Dim i_signUpStatus As Integer = TIMS.CINT1(dr1("signUpStatus"))
        'signUpStatus 0:收件完成 1:報名成功 2:報名失敗 3:正取(Key_SelResult) 4:備取 5:未錄取
        Select Case sResultSelValue 'Result.SelectedValue
            Case TIMS.cst_SelResultID_正取 '"01" '正取
                i_signUpStatus = 3
            Case TIMS.cst_SelResultID_備取 '"02" '備取
                i_signUpStatus = 4
            Case TIMS.cst_SelResultID_未錄取
                i_signUpStatus = 5
            Case TIMS.cst_SelResultID_缺考 '"02" '備取
                i_signUpStatus = 1
            Case TIMS.cst_SelResultID_審核中
                i_signUpStatus = 1
        End Select
        dr1("signUpStatus") = i_signUpStatus
        DbAccess.UpdateDataTable(dt1, da, oTrans)
    End Sub

    ''' <summary>系統使用者，可增加儲存警告次數</summary>
    Sub Utl_CAN_IGNORE_RULE1_CNT()
        Dim flagS1 As Boolean = flgROLEIDx0xLIDx0 ' TIMS.IsSuperUser(Me, 1) '是否為(後台)系統管理者 
        If Not flagS1 Then
            ViewState(vs_SD02003_OCID1) = Nothing
            Hid_CAN_IGNORE_RULE1_CNT.Value = ""
            Return
        End If
        If (Hid_CAN_IGNORE_RULE1_CNT.Value = "") Then
            ViewState(vs_SD02003_OCID1) = Hid_OCID1.Value
            Hid_CAN_IGNORE_RULE1_CNT.Value = "1"
            Return
        End If
        If (Hid_CAN_IGNORE_RULE1_CNT.Value <> "") AndAlso ViewState(vs_SD02003_OCID1) = Hid_OCID1.Value Then
            Hid_CAN_IGNORE_RULE1_CNT.Value = TIMS.CINT1(Hid_CAN_IGNORE_RULE1_CNT.Value) + 1
        End If
    End Sub

    ''' <summary>取得儲存警告次數</summary>
    ''' <returns></returns>
    Function Get_CAN_IGNORE_RULE1_CNT() As Integer
        Dim rst As Integer = 0
        If (Hid_CAN_IGNORE_RULE1_CNT.Value = "") Then Return rst
        If (Convert.ToString(ViewState(vs_SD02003_OCID1)) <> Hid_OCID1.Value) Then Return rst
        rst = TIMS.CINT1(Hid_CAN_IGNORE_RULE1_CNT.Value)
        Return rst
    End Function

    ''' <summary>完成錄取(儲存)</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub btnSEND1_Click(sender As Object, e As EventArgs) Handles btnSEND1.Click
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        '班級資料
        Dim drOC As DataRow = Nothing
        Dim sErrmsg As String = ""
        Call CheckCONF1(sErrmsg, drOC)
        If sErrmsg <> "" Then
            Utl_CAN_IGNORE_RULE1_CNT()
            Common.MessageBox(Me, sErrmsg)
            Exit Sub
        End If

        'dt9 已有學員資料
        Dim dt9 As DataTable = Nothing
        '儲存前 檢核可能的錯誤
        'Dim sErrmsg As String=""
        Call CheckData1(sErrmsg, dt9, drOC)
        If sErrmsg <> "" Then
            Utl_CAN_IGNORE_RULE1_CNT()
            Common.MessageBox(Me, sErrmsg)
            Exit Sub
        End If

        '完成錄取(儲存)
        Call SaveData1(dt9, drOC)

        '單位確認-不鎖定 (NOLOCK='Y') 儲存-CLASS_CONFIRM (NOLOCK='Y') 儲存但不鎖定
        Call SaveCONF2()
    End Sub

    ''' <summary>檢查是否尚有e網報名未審學員</summary>
    ''' <param name="OCID1"></param>
    ''' <param name="oConn"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Chk_EsignUpStatus(ByVal OCID1 As String, ByRef oConn As SqlConnection) As Boolean
        Dim flag As Boolean = False '預設為沒有e網報名未審學員
        If OCID1 = "" Then Return flag
        'dt=DbAccess.GetDataTable(sql, oConn)
        '" & OCID1 & "'" & vbCrLf
        Dim parms As New Hashtable From {{"OCID1", OCID1}}
        Dim sql As String = ""
        sql &= " SELECT e.Name, e.IDNO ,a.eSETID, a.RelEnterDate, a.signUpStatus ,b1.OCID, b1.ClassCName, d.OrgName" & vbCrLf
        sql &= " FROM Stud_EnterType2 a" & vbCrLf
        sql &= " JOIN Stud_EnterTemp2 e ON a.eSETID=e.eSETID" & vbCrLf
        sql &= " JOIN Class_ClassInfo b1 ON a.OCID1=b1.OCID" & vbCrLf
        sql &= " JOIN Auth_Relship c ON b1.RID=c.RID" & vbCrLf
        sql &= " JOIN Org_OrgInfo d ON c.OrgID=d.Orgid" & vbCrLf
        sql &= " WHERE a.signUpStatus=0 AND a.OCID1=@OCID1" & vbCrLf
        'Dim sCmd As New SqlCommand(sql, oConn) 'Dim dt As New DataTable
        'With sCmd
        '    .Parameters.Clear()
        '    .Parameters.Add("OCID1", SqlDbType.VarChar).Value=OCID1
        '    dt.Load(.ExecuteReader())
        'End With
        'parms.Clear()
        Dim dt As DataTable = DbAccess.GetDataTable(sql, oConn, parms)
        If TIMS.dtHaveDATA(dt) Then flag = True
        'Common.MessageBox(Me, "目前尚有e網報名學員尚未審核，需先將【e網報名】的所有學員審核完成後，始能登錄【錄取作業】！！")
        Return flag
    End Function

    ''' <summary>取得儲存後的錄訓結果</summary>
    ''' <param name="OCID1"></param>
    ''' <param name="tConn"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Get_SELRESULTdt2(ByVal OCID1 As String, ByRef tConn As SqlConnection) As DataTable
        Dim dtT As DataTable = Nothing
        If OCID1 = "" Then Return dtT

        'dt9 已有 錄訓結果資料
        'parms.Clear()
        Dim parms As New Hashtable From {{"OCID", OCID1}}
        Dim sql As String = ""
        sql &= " SELECT se.name ,ss.SETID ,CONVERT(VARCHAR, ss.EnterDate, 111) EnterDate ,ss.serNum ,ss.SelResultID" & vbCrLf
        sql &= " ,dbo.FN_SELRESULTID(ss.SelResultID,1) SELRESULTID1" & vbCrLf
        sql &= " ,dbo.FN_SELRESULTID(ss.SelResultID,2) SELRESULTID2" & vbCrLf
        sql &= " FROM STUD_ENTERTEMP se" & vbCrLf
        sql &= " JOIN STUD_SELRESULT ss ON ss.SETID=se.SETID" & vbCrLf
        sql &= " JOIN STUD_ENTERTYPE sr ON sr.SETID=ss.SETID AND sr.EnterDate=ss.EnterDate AND sr.SerNum=ss.SerNum" & vbCrLf
        sql &= " WHERE ss.OCID=@OCID" & vbCrLf
        sql &= " AND ss.SelResultID IN ('01','02','03')" & vbCrLf '正取/不錄取(備取)
        dtT = DbAccess.GetDataTable(sql, tConn, parms)
        'If oflag_Test Then Return True '測試啟用，不作判斷
        Return dtT
    End Function

    ''' <summary>判斷式以成績或者是以報名順序來區分</summary>
    ''' <param name="OCID1"></param>
    ''' <param name="oConn"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GET_SELRESULTdt(ByVal OCID1 As String, ByRef oConn As SqlConnection) As DataTable
        Dim dtChk As New DataTable
        If OCID1 = "" Then Return dtChk
        'SumOfGrad	甄試成績 
        Dim parms As New Hashtable From {{"OCID1", OCID1}}
        Dim sql As String = ""
        sql &= " SELECT OCID, ISNULL(SUM(SumOfGrad),0) total "
        sql &= " FROM STUD_SELRESULT "
        sql &= " WHERE OCID=@OCID1 "
        sql &= " GROUP BY OCID "
        'Dim sCmd As New SqlCommand(sql, oConn)
        'With sCmd
        '    .Parameters.Clear()
        '    .Parameters.Add("OCID1", SqlDbType.VarChar).Value=OCID1 'OCIDValue1.Value
        '    dtChk.Load(.ExecuteReader())
        'End With
        dtChk = DbAccess.GetDataTable(sql, oConn, parms)
        Return dtChk
    End Function

    ''' <summary>儲存前 檢核可能的錯誤</summary>
    ''' <param name="Errmsg"></param>
    ''' <param name="dt9"></param>
    ''' <param name="drOC"></param>
    ''' <returns></returns>
    Function CheckData1(ByRef Errmsg As String, ByRef dt9 As DataTable, ByRef drOC As DataRow) As Boolean
        Dim rst As Boolean = True
        Errmsg = ""
        ViewState(cst_OCIDValue1) = TIMS.ClearSQM(ViewState(cst_OCIDValue1))
        If Convert.ToString(ViewState(cst_OCIDValue1)) = "" Then
            Errmsg &= "請重新查詢班級資料!!" & vbCrLf
            Return False
        End If
        If Convert.ToString(ViewState(cst_OCIDValue1)) = "" _
            OrElse Convert.ToString(ViewState(cst_OCIDValue1)) = "0" Then
            Errmsg &= "班級選擇有誤!!" & vbCrLf
            Return False
        End If
        If ViewState(cst_OCIDValue1) <> OCIDValue1.Value Then
            Errmsg &= "班級選擇有誤!!" & vbCrLf
            Return False
        End If
        If Hid_OCID1.Value <> OCIDValue1.Value Then
            Errmsg &= "班級選擇有誤!!" & vbCrLf
            Return False
        End If
        'Dim drOC As DataRow=TIMS.GetOCIDDate(ViewState(cst_OCIDValue1), objconn)
        If drOC Is Nothing Then drOC = TIMS.GetOCIDDate(ViewState(cst_OCIDValue1), objconn)
        If drOC Is Nothing Then
            Errmsg &= "班級選擇有誤!!" & vbCrLf
            Return False
        End If

        'dt9 已有學員資料
        Dim pParms As New Hashtable From {{"OCID", ViewState(cst_OCIDValue1)}}
        Dim sql As String = ""
        sql &= " SELECT cs.SOCID ,se.NAME ,ss.SETID ,CONVERT(VARCHAR, ss.EnterDate, 111) EnterDate ,ss.serNum ,ss.SelResultID" & vbCrLf
        sql &= " ,dbo.FN_SELRESULTID(ss.SelResultID,1) SELRESULTID1" & vbCrLf
        sql &= " ,dbo.FN_SELRESULTID(ss.SelResultID,2) SELRESULTID2" & vbCrLf
        sql &= " FROM STUD_ENTERTEMP se" & vbCrLf
        sql &= " JOIN CLASS_STUDENTSOFCLASS cs ON se.setid=cs.setid" & vbCrLf
        sql &= " JOIN STUD_SELRESULT ss ON cs.SETID=ss.SETID AND cs.ETEnterDate=ss.EnterDate AND cs.SerNum=ss.serNum" & vbCrLf
        sql &= " WHERE cs.OCID=@OCID" & vbCrLf
        sql &= " AND ss.SelResultID NOT IN ('02')" & vbCrLf '正取:01/不錄取:03 (排除備取:02)
        'sql &= " AND ss.SelResultID IN ('01','03')" & vbCrLf '正取/不錄取 (排除備取)
        dt9 = DbAccess.GetDataTable(sql, objconn, pParms)
        'If oflag_Test Then Return True '測試啟用，不作判斷

        '非系統管理者(超過14日)
        'Dim flagS1 As Boolean=TIMS.IsSuperUser(Me, 1)
        Dim flagS1 As Boolean = flgROLEIDx0xLIDx0 ' TIMS.IsSuperUser(Me, 1) '是否為(後台)系統管理者 
        Dim flag_can_ignore_control As Boolean = False
        If TIMS.Cst_TPlanID70.IndexOf(sm.UserInfo.TPlanID) > -1 AndAlso DateDiff(DateInterval.Day, Now, CDate(TIMS.cst_TPlanID70_end_date_1)) >= 0 Then
            '70: 區域產業據點職業訓練計畫(在職)
            flag_can_ignore_control = True '忽視卡關-暫時
        End If

        '忽視卡關
        Dim iIGNORE_RULE1_CNT As Integer = Get_CAN_IGNORE_RULE1_CNT()
        If iIGNORE_RULE1_CNT >= 4 Then Return True

        '暫時權限Table
        Dim dtArc As DataTable = TIMS.Get_Auth_REndClass(Me, objconn)
        Dim fg_RULE2870 As Boolean = (TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1) OrElse (TIMS.Cst_TPlanID70.IndexOf(sm.UserInfo.TPlanID) > -1)
        If fg_RULE2870 Then
            If Not flagS1 AndAlso $"{drOC("InputOK14")}" <> "Y" Then
                Dim flagInputOK14NG As Boolean = False '開訓日後14日鎖定功能填寫!(預設:不鎖定)
                'https://jira.turbotech.com.tw/browse/TIMSC-161
                If TIMS.ChkIsEndDate(OCIDValue1.Value, TIMS.cst_FunID_錄訓作業, dtArc) Then
                    '過了使用期限 True(不可使用) False(可使用) '開訓日後14日鎖定功能填寫!
                    flagInputOK14NG = True '(鎖定)
                End If
                If flag_can_ignore_control Then flagInputOK14NG = False '可忽略
                If flagInputOK14NG Then
                    Errmsg &= $"{cst_msg0012}{vbCrLf}"
                    Return False
                End If
            End If
        Else
            If flag_can_ignore_control Then flagS1 = True '可忽略
            'https://jira.turbotech.com.tw/browse/TIMSC-161
            If Not flagS1 Then
                '(超過開訓日)(未超過14日)
                If TIMS.ChkIsEndDate(OCIDValue1.Value, TIMS.cst_FunID_錄訓作業, dtArc) Then
                    '過了使用期限 True(不可使用)   False(可使用)
                    If $"{drOC("InputOK")}" <> "Y" Then
                        Errmsg &= $"{cst_msg0013}{vbCrLf}"
                        Return False
                    End If
                End If
            End If
        End If

        'WebRequest物件如何忽略憑證問題
        System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf TIMS.ValidateServerCertificate)
        'TLS 1.2-基礎連接已關閉: 傳送時發生未預期的錯誤 
        System.Net.ServicePointManager.SecurityProtocol = Net.SecurityProtocolType.Tls12
        '檢核學員重複參訓。
        'http://163.29.199.211/TIMSWS/timsService1.asmx
        'https://wltims.wda.gov.tw/TIMSWS/timsService1.asmx
        Dim timsSer1 As New timsService1.timsService1

        Dim stdBLACK2TPLANID As String = ""
        Dim iStdBlackType As Integer = TIMS.Chk_StdBlackType(Me, objconn, stdBLACK2TPLANID)
        Dim sBlackIDNO As String = TIMS.Get_StdBlackIDNO(Me, iStdBlackType, stdBLACK2TPLANID, objconn) '學員處分

        'Dim vs_Alertmsg As String=""
        For Each eItem As DataGridItem In Me.DataGrid1.Items
            Dim hidIDNO As HtmlInputHidden = eItem.FindControl("hidIDNO") 'IDNO
            Dim notes2 As TextBox = eItem.FindControl("notes2") '備註  'notes2 備取或未錄取原因
            Dim Hid_rsort As HiddenField = eItem.FindControl("Hid_rsort") 'rsort 名次

            Dim ddlResult As DropDownList = eItem.FindControl("ddlResult") '=myitem.Cells(8).Controls(1)
            Dim v_ddlResult As String = TIMS.GetListValue(ddlResult) 'Result.SelectedValue '01:正取 /02:備取 /03:未錄取 '

            Dim Hid_SETID As HiddenField = eItem.FindControl("Hid_SETID")
            Dim Hid_EnterDate As HiddenField = eItem.FindControl("Hid_EnterDate")
            Dim Hid_SerNum As HiddenField = eItem.FindControl("Hid_SerNum")
            Dim strSetId As String = TIMS.ClearSQM(Hid_SETID.Value)
            Dim strEnterDate As String = TIMS.Cdate3(Hid_EnterDate.Value)
            Dim strSerNum As String = TIMS.ClearSQM(Hid_SerNum.Value)

            Dim labname As Label = eItem.FindControl("labname")
            If strEnterDate = "" Then
                Errmsg &= "請重新查詢班級資料!!" & vbCrLf
                Return False
            End If
            Dim ffSCH1 As String = ""
            ffSCH1 = "SETID='" & strSetId & "' AND SerNum='" & strSerNum & "' AND EnterDate='" & strEnterDate & "' "
            Dim dr As DataRow = If(dt9.Select(ffSCH1).Length > 0, dt9.Select(ffSCH1)(0), Nothing)

            'If dt9.Select(ff).Length > 0 Then dr=dt9.Select(ff)(0)
            'If dr IsNot Nothing Then
            '    If Convert.ToString(dr("SelResultID")) <> v_ddlResult Then
            '        原甄試結果
            '        v_ddlResult=Convert.ToString(dr("SelResultID"))
            '        Msg += "學員【 " & dr("name") & "】己在【學員資料維護】功能有學員資料,故不能改變甄試結果,保留原甄試結果為【 " & dr("SelResultID2") & "】" & vbCrLf
            '        只是提醒。
            '        vs_Alertmsg += "學員【 " & dr("name") & "】己在【學員資料維護】功能有學員資料,故不能改變甄試結果,保留原甄試結果為【 " & dr("SelResultID2") & "】" & vbCrLf
            '    End If
            'End If
            If dr Is Nothing Then
                '01:正取 02:備取 03:未錄取
                Select Case v_ddlResult'Result.SelectedValue
                    Case TIMS.cst_SelResultID_備取, TIMS.cst_SelResultID_未錄取 '"01", "03"
                        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                            'rsort 名次 / 'notes2 備取或未錄取原因
                            Dim flag_show_msg0017 As Boolean = False
                            If Hid_rsort.Value <> "" AndAlso TIMS.CINT1(Hid_rsort.Value) <= TIMS.CINT1(drOC("TNum")) AndAlso TIMS.ClearSQM(notes2.Text) = "" Then flag_show_msg0017 = True
                            If flag_show_msg0017 Then
                                'Const cst_msg0017 As String="當「名次」 位於「訓練人數」名額內，若錄訓結果為「備取」或「未錄取」請增加填寫「備取或未錄取原因」"
                                Errmsg &= labname.Text & cst_msg0017 & vbCrLf
                                If Errmsg <> "" Then Return False
                            End If
                        End If
                End Select

                '01:正取 02:備取 03:未錄取 (ddlResult.Enabled 確認檢核)
                Select Case v_ddlResult 'TIMS.GetListValue(ddlResult)'.SelectedValue
                    Case TIMS.cst_SelResultID_正取, TIMS.cst_SelResultID_備取
                        'If oflag_Test Then
                        '    Errmsg &= labname.Text & "，依處分日期及年限，仍在處分期間者，錄取作業，被處分者不可設定為正取或備取。" & vbCrLf '加入單名單暫存(2009/07/28 判斷黑名單)
                        '    Return False
                        'End If
                        If sBlackIDNO <> "" AndAlso sBlackIDNO.IndexOf(hidIDNO.Value) > -1 AndAlso ddlResult.Enabled Then
                            Errmsg &= $"{labname.Text}，依處分日期及年限，仍在處分期間者，錄取作業，被處分者不可設定為正取或備取。{vbCrLf}"  '加入單名單暫存(2009/07/28 判斷黑名單)
                            If Errmsg <> "" Then Return False
                        End If
                        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 AndAlso ddlResult.Enabled Then
                            '修改說明：
                            '1.針對報名民眾於不同階段發生參訓時段重疊(報名)情形，設立無法儲存階段及彈跳提醒視窗，於「錄取作業」點選「正取」，無法儲存，給予彈跳視窗提醒(如圖3.jpg)。此外，參訓時段重疊報名情形比對，「錄取作業」項下備取生和未選取者&「e網報名審核」項下報名審核未點選成功或失敗者，於該課程開訓後第15天取消比對。
                            '2.課程開訓第15天起，如透過系統人員將學員從備取生轉至正取生，應再次檢核學員重參情形，建議新增檢核功能鈕(按下時，以班級或正取生為單位，即時檢核，並不設「課程開訓後第15天取消比對」的限制)。
                            '3.報名資料於「錄取作業」未完成點選(仍在請選擇)時，民眾仍可於報名網站自行取消報名;如已點選正取、備取、未錄取則無法於報名網站自行取消。
                            'by AMU 20160202 'Dim drET2 As DataRow=TIMS.Get_ENTERTYPE2(Hid_eSerNum.Value, objconn)

                            hidIDNO.Value = TIMS.ClearSQM(hidIDNO.Value)
                            ViewState(cst_OCIDValue1) = TIMS.ClearSQM(ViewState(cst_OCIDValue1))
                            Dim aIDNO As String = hidIDNO.Value
                            Dim aOCID1 As String = ViewState(cst_OCIDValue1)
                            Dim xStudInfo As String = ""
                            TIMS.SetMyValue(xStudInfo, "IDNO", aIDNO)
                            TIMS.SetMyValue(xStudInfo, "OCID1", aOCID1)
                            '檢核學員OOO有參訓時段重疊報名情形，無法儲存，請再確認！
                            Call TIMS.ChkStudDouble(timsSer1, Errmsg, labname.Text, xStudInfo)
                            If Errmsg <> "" Then Return False
                        End If
                End Select
            End If
        Next
        If Errmsg <> "" Then rst = False
        Return rst
    End Function

    ''' <summary>儲存前 檢核可能的錯誤</summary>
    ''' <param name="Errmsg"></param>
    ''' <returns></returns>
    Function CheckCONF1(ByRef Errmsg As String, ByRef drOC As DataRow) As Boolean
        Dim rst As Boolean = True
        Errmsg = ""
        ViewState(cst_OCIDValue1) = TIMS.ClearSQM(ViewState(cst_OCIDValue1))
        If Convert.ToString(ViewState(cst_OCIDValue1)) = "" Then
            Errmsg &= "請重新查詢班級資料!!" & vbCrLf
            Return False
        End If
        If Convert.ToString(ViewState(cst_OCIDValue1)) = "" _
            OrElse Convert.ToString(ViewState(cst_OCIDValue1)) = "0" Then
            Errmsg &= "班級選擇有誤!!" & vbCrLf
            Return False
        End If
        If ViewState(cst_OCIDValue1) <> OCIDValue1.Value Then
            Errmsg &= "班級選擇有誤!!" & vbCrLf
            Return False
        End If
        If Hid_OCID1.Value <> OCIDValue1.Value Then
            Errmsg &= "班級選擇有誤!!" & vbCrLf
            Return False
        End If
        'Dim drOC As DataRow=TIMS.GetOCIDDate(ViewState(cst_OCIDValue1), objconn)
        If drOC Is Nothing Then drOC = TIMS.GetOCIDDate(ViewState(cst_OCIDValue1), objconn)
        If drOC Is Nothing Then
            Errmsg &= "班級選擇有誤!!" & vbCrLf
            Return False
        End If

        '非系統管理者(超過14日)
        'Dim flagS1 As Boolean=TIMS.IsSuperUser(Me, 1)
        Dim flagS1 As Boolean = flgROLEIDx0xLIDx0 ' TIMS.IsSuperUser(Me, 1) '是否為(後台)系統管理者 
        Dim flag_can_ignore_control As Boolean = False
        If TIMS.Cst_TPlanID70.IndexOf(sm.UserInfo.TPlanID) > -1 AndAlso DateDiff(DateInterval.Day, Now, CDate(TIMS.cst_TPlanID70_end_date_1)) >= 0 Then
            '70: 區域產業據點職業訓練計畫(在職)
            flag_can_ignore_control = True '忽視卡關-暫時
        End If

        '忽視卡關
        Dim iIGNORE_RULE1_CNT As Integer = Get_CAN_IGNORE_RULE1_CNT()
        If iIGNORE_RULE1_CNT >= 4 Then Return True

        '暫時權限Table
        Dim dtArc As DataTable = TIMS.Get_Auth_REndClass(Me, objconn)
        Dim fg_RULE2870 As Boolean = (TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1) OrElse (TIMS.Cst_TPlanID70.IndexOf(sm.UserInfo.TPlanID) > -1)
        If fg_RULE2870 Then
            If Not flagS1 AndAlso Convert.ToString(drOC("InputOK14")) <> "Y" Then
                Dim flagInputOK14NG As Boolean = False '開訓日後14日鎖定功能填寫!(預設:不鎖定)
                'https://jira.turbotech.com.tw/browse/TIMSC-161
                If TIMS.ChkIsEndDate(OCIDValue1.Value, TIMS.cst_FunID_錄訓作業, dtArc) Then
                    '過了使用期限 True(不可使用) False(可使用) '開訓日後14日鎖定功能填寫!
                    flagInputOK14NG = True '(鎖定)
                End If
                If flag_can_ignore_control Then flagInputOK14NG = False '可忽略
                If flagInputOK14NG Then
                    Errmsg &= $"{cst_msg0012}{vbCrLf}"
                    If Errmsg <> "" Then Return False 'Return False
                End If
            End If
        Else
            If flag_can_ignore_control Then flagS1 = True '可忽略
            'https://jira.turbotech.com.tw/browse/TIMSC-161
            If Not flagS1 Then
                '(超過開訓日)(未超過14日)
                If TIMS.ChkIsEndDate(OCIDValue1.Value, TIMS.cst_FunID_錄訓作業, dtArc) Then
                    '過了使用期限 True(不可使用)   False(可使用)
                    If Convert.ToString(drOC("InputOK")) <> "Y" Then
                        Errmsg &= $"{cst_msg0013}{vbCrLf}"
                        If Errmsg <> "" Then Return False
                    End If
                End If
            End If
        End If

        'WebRequest物件如何忽略憑證問題
        System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf TIMS.ValidateServerCertificate)
        'TLS 1.2-基礎連接已關閉: 傳送時發生未預期的錯誤 
        System.Net.ServicePointManager.SecurityProtocol = Net.SecurityProtocolType.Tls12
        '檢核學員重複參訓。
        'http://163.29.199.211/TIMSWS/timsService1.asmx
        'https://wltims.wda.gov.tw/TIMSWS/timsService1.asmx
        Dim timsSer1 As New timsService1.timsService1

        Dim stdBLACK2TPLANID As String = ""
        Dim iStdBlackType As Integer = TIMS.Chk_StdBlackType(Me, objconn, stdBLACK2TPLANID)
        Dim sBlackIDNO As String = TIMS.Get_StdBlackIDNO(Me, iStdBlackType, stdBLACK2TPLANID, objconn) '學員處分
        Dim dtX2 As DataTable = Get_SELRESULTdt2(ViewState(cst_OCIDValue1), objconn)
        'Dim vs_Alertmsg As String=""
        For Each eItem As DataGridItem In Me.DataGrid1.Items
            Dim hidIDNO As HtmlInputHidden = eItem.FindControl("hidIDNO") 'IDNO
            Dim notes2 As TextBox = eItem.FindControl("notes2")
            'Dim StrSql As String=""
            'Dim sAdmission As String="" 'Y/N
            Dim ddlResult As DropDownList = eItem.FindControl("ddlResult") '=myitem.Cells(8).Controls(1)
            Dim v_ddlResult As String = TIMS.GetListValue(ddlResult) '.SelectedValue '01:正取 /02:備取 /03:未錄取 '05:審核中
            'Dim strName As String=eItem.Cells(Cst_ceName).Text 'Name
            'Dim strSetId As String=eItem.Cells(Cst_ceSETID).Text 'SETID
            'Dim strSerNum As String=eItem.Cells(Cst_ceSerNum).Text 'SerNum 
            Dim strEnterDate As String = eItem.Cells(cst_ceEnterDate).Text 'EnterDate
            strEnterDate = TIMS.Cdate3(strEnterDate)
            Dim labname As Label = eItem.FindControl("labname")
            If strEnterDate = "" Then
                Errmsg &= "請重新查詢班級資料!!" & vbCrLf
                If Errmsg <> "" Then Return False
            End If
            'Dim drRS As DataRow=Nothing
            'Dim ff3 As String=""
            'ff3="SETID='" & strSetId & "' and SerNum='" & strSerNum & "' and EnterDate='" & strEnterDate & "'"
            'If dtX2.Select(ff3).Length <> 1 Then
            '    Errmsg &= "請先完成錄取作業!!" & vbCrLf
            '    Return False
            'End If
            'If dtX2.Select(ff3).Length > 0 Then
            '    有結果資料
            '    drRS=dtX2.Select(ff3)(0)
            '    If Convert.ToString(drRS("SelResultID")) <> sResult Then
            '        Errmsg &= "送出時不可改變 原錄取作業 甄試結果,保留原甄試結果為【 " & drRS("SelResultID2") & "】，請先完成錄取作業!!" & vbCrLf
            '        Return False
            '    End If
            'End If
            '01:正取 02:備取 03:未錄取
            Select Case v_ddlResult'Result.SelectedValue
                Case TIMS.cst_SelResultID_正取, TIMS.cst_SelResultID_備取
                    If sBlackIDNO <> "" AndAlso sBlackIDNO.IndexOf(hidIDNO.Value) > -1 AndAlso ddlResult.Enabled Then
                        Errmsg &= $"{labname.Text}，依處分日期及年限，仍在處分期間者，錄取作業，被處分者不可設定為正取或備取。{vbCrLf}"  '加入單名單暫存(2009/07/28 判斷黑名單)
                        Return False
                    End If
                    If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 AndAlso ddlResult.Enabled Then
                        '修改說明：
                        '1.針對報名民眾於不同階段發生參訓時段重疊(報名)情形，設立無法儲存階段及彈跳提醒視窗，於「錄取作業」點選「正取」，無法儲存，給予彈跳視窗提醒(如圖3.jpg)。此外，參訓時段重疊報名情形比對，「錄取作業」項下備取生和未選取者&「e網報名審核」項下報名審核未點選成功或失敗者，於該課程開訓後第15天取消比對。
                        '2.課程開訓第15天起，如透過系統人員將學員從備取生轉至正取生，應再次檢核學員重參情形，建議新增檢核功能鈕(按下時，以班級或正取生為單位，即時檢核，並不設「課程開訓後第15天取消比對」的限制)。
                        '3.報名資料於「錄取作業」未完成點選(仍在請選擇)時，民眾仍可於報名網站自行取消報名;如已點選正取、備取、未錄取則無法於報名網站自行取消。
                        'by AMU 20160202
                        'Dim drET2 As DataRow=TIMS.Get_ENTERTYPE2(Hid_eSerNum.Value, objconn)
                        hidIDNO.Value = TIMS.ClearSQM(hidIDNO.Value)
                        ViewState(cst_OCIDValue1) = TIMS.ClearSQM(ViewState(cst_OCIDValue1))
                        Dim aIDNO As String = hidIDNO.Value
                        Dim aOCID1 As String = ViewState(cst_OCIDValue1)
                        Dim xStudInfo As String = ""
                        TIMS.SetMyValue(xStudInfo, "IDNO", aIDNO)
                        TIMS.SetMyValue(xStudInfo, "OCID1", aOCID1)
                        '檢核學員OOO有參訓時段重疊報名情形，無法儲存，請再確認！
                        Call TIMS.ChkStudDouble(timsSer1, Errmsg, labname.Text, xStudInfo)
                        If Errmsg <> "" Then Return False
                    End If
            End Select
        Next
        If Errmsg <> "" Then rst = False
        Return rst
    End Function

    ''' <summary>SAVE STUD_SELRESULT -SqlTransaction</summary>
    ''' <param name="eItem"></param>
    ''' <param name="drOC"></param>
    ''' <param name="sErrmsg"></param>
    ''' <param name="oTrans"></param>
    Sub Utl_SaveDataRow(ByRef eItem As DataGridItem, ByRef drOC As DataRow, ByRef sErrmsg As String, ByRef oTrans As SqlTransaction)
        'Dim sary As Array 'sary=myitem
        Dim hidIDNO As HtmlInputHidden = eItem.FindControl("hidIDNO")
        Dim notes2 As TextBox = eItem.FindControl("notes2")
        'Dim StrSql As String=""
        'Dim sAdmission As String="" 'Y/N
        'Dim sResult As String="" '01/02/03
        'Dim strSetId As String=""
        'Dim strSerNum As String=""
        'Dim strEnterDate As String=""
        'Dim strSetId As String=eItem.Cells(Cst_ceSETID).Text
        'Dim strEnterDate As String=eItem.Cells(Cst_ceEnterDate).Text
        'Dim strSerNum As String=eItem.Cells(Cst_ceSerNum).Text
        'drop=myitem.Cells(8).Controls(1)
        'RSort=Convert.ToString(myitem.Cells(8).Text)
        'Admission= Result.SelectedValue
        Dim Hid_SETID As HiddenField = eItem.FindControl("Hid_SETID")
        Dim Hid_EnterDate As HiddenField = eItem.FindControl("Hid_EnterDate")
        Dim Hid_SerNum As HiddenField = eItem.FindControl("Hid_SerNum")
        Dim strSetId As String = TIMS.ClearSQM(Hid_SETID.Value)
        Dim strEnterDate As String = TIMS.Cdate3(Hid_EnterDate.Value)
        Dim strSerNum As String = TIMS.ClearSQM(Hid_SerNum.Value)
        Dim ddlResult As DropDownList = eItem.FindControl("ddlResult") '=myitem.Cells(8).Controls(1)
        Dim v_ddlResult As String = TIMS.GetListValue(ddlResult) ' Result.SelectedValue

        Dim sAdmission As String = If(v_ddlResult = "01", "Y", "N") '是否錄取:錄取 (除了正取才是錄取) N" '不錄取
        Dim htSS As New Hashtable
        htSS.Add("strSetId", strSetId)
        htSS.Add("strSerNum", strSerNum)
        htSS.Add("strEnterDate", strEnterDate)
        htSS.Add("sNotes2", notes2.Text)
        htSS.Add("sAdmission", sAdmission)
        htSS.Add("sResultSelValue", v_ddlResult)
        Call UPDATE_STUD_SELRESULT(oTrans, htSS, drOC, sErrmsg)
    End Sub

    ''' <summary> 儲存-UPDATE STUD_SELRESULT </summary>
    Sub SaveData1(ByRef dt9 As DataTable, ByRef drOC As DataRow)
        ''dt9 已有學員資料
        'Dim dt9 As DataTable=Nothing
        ''班級資料
        'Dim drOC As DataRow=Nothing
        ''儲存前 檢核可能的錯誤
        Dim sErrmsg As String = ""
        'Call CheckData1(sErrmsg, dt9, drOC)
        'If sErrmsg <> "" Then
        '    Common.MessageBox(Me, sErrmsg)
        '    Exit Sub
        'End If

        ViewState(cst_watingcount) = 0 '集合備取數
        Dim strSETIDALL As String = "" '集合備取
        Dim vs_Alertmsg As String = ""
        Dim tConn As SqlConnection = DbAccess.GetConnection()
        Dim oTrans As SqlTransaction = DbAccess.BeginTrans(tConn)
        Try
            For Each eItem As DataGridItem In Me.DataGrid1.Items
                'Dim sary As Array 'sary=myitem
                Dim hidIDNO As HtmlInputHidden = eItem.FindControl("hidIDNO")
                Dim notes2 As TextBox = eItem.FindControl("notes2")
                'Dim StrSql As String=""
                'Dim sAdmission As String="" 'Y/N
                'Dim sResult As String="" '01/02/03
                'Dim strSetId As String=""
                'Dim strSerNum As String=""
                'Dim strEnterDate As String=""
                'Dim strSetId As String=eItem.Cells(Cst_ceSETID).Text
                'Dim strEnterDate As String=eItem.Cells(Cst_ceEnterDate).Text
                'Dim strSerNum As String=eItem.Cells(Cst_ceSerNum).Text
                'drop=myitem.Cells(8).Controls(1)
                'RSort=Convert.ToString(myitem.Cells(8).Text)
                'Admission= Result.SelectedValue
                Dim Hid_SETID As HiddenField = eItem.FindControl("Hid_SETID")
                Dim Hid_EnterDate As HiddenField = eItem.FindControl("Hid_EnterDate")
                Dim Hid_SerNum As HiddenField = eItem.FindControl("Hid_SerNum")
                Dim strSetId As String = TIMS.ClearSQM(Hid_SETID.Value)
                Dim strEnterDate As String = TIMS.Cdate3(Hid_EnterDate.Value)
                Dim strSerNum As String = TIMS.ClearSQM(Hid_SerNum.Value)
                Dim ddlResult As DropDownList = eItem.FindControl("ddlResult") '=myitem.Cells(8).Controls(1)
                Dim v_ddlResult As String = TIMS.GetListValue(ddlResult) ' Result.SelectedValue

                'PlanID=eItem.Cells(Cst_ceTRNDType).ToolTip '特別使用
                '------------------找出那些學員是備取及幾個備取---------
                'Result=myitem.Cells(8).Controls(1)    '是否錄取
                '        strSETID=myitem.Cells(10).Text        'SETID
                Select Case v_ddlResult'Result.SelectedValue
                    Case TIMS.cst_SelResultID_正取, TIMS.cst_SelResultID_未錄取'"01", "03"
                        '正取、未錄取
                    Case TIMS.cst_SelResultID_備取 '"02" 'Case Else
                        '若是備取
                        '集合備取
                        ViewState(cst_watingcount) += 1
                        If strSETIDALL <> "" Then strSETIDALL &= ","
                        strSETIDALL &= strSetId
                End Select
                '------------------找出那些學員是備取及幾個備取---end------

                Dim dr As DataRow = Nothing
                Dim ff As String = String.Concat("SETID='", strSetId, "' AND SerNum='", strSerNum, "' AND EnterDate='", Common.FormatDate(strEnterDate), "'")
                If dt9.Select(ff).Length > 0 Then dr = dt9.Select(ff)(0)
                If dr IsNot Nothing Then
                    'ddlResult 01:正取 02:備取 03:未錄取
                    Dim flag_can_save As Boolean = True 'true:可儲存 false:不可儲存
                    If v_ddlResult <> TIMS.cst_SelResultID_正取 AndAlso Convert.ToString(dr("SelResultID")) <> v_ddlResult Then
                        '非正取，且改變 原甄試結果
                        v_ddlResult = Convert.ToString(dr("SelResultID"))
                        Dim vSELRESULTID2 As String = Convert.ToString(dr("SELRESULTID2"))
                        'Msg += "學員【 " & dr("name") & "】己在【學員資料維護】功能有學員資料,故不能改變甄試結果,保留原甄試結果為【 " & dr("SelResultID2") & "】" & vbCrLf
                        vs_Alertmsg &= String.Concat("學員【 ", dr("name"), "】己在【學員資料維護】功能有學員資料,故不能改變甄試結果,保留原甄試結果為【 ", vSELRESULTID2, "】", vbCrLf)
                        flag_can_save = False 'true:可儲存 false:不可儲存
                    End If
                    If flag_can_save Then
                        Call Utl_SaveDataRow(eItem, drOC, sErrmsg, oTrans)
                        If sErrmsg <> "" Then
                            'sErrmsg=Replace(sErrmsg, vbCrLf, "<br>" & vbCrLf)
                            Call TIMS.WriteTraceLog(sErrmsg)
                            DbAccess.RollbackTrans(oTrans)
                            Call TIMS.CloseDbConn(tConn)
                            Common.MessageBox(Me, cst_errmsg1)
                            Exit Sub
                            'Common.MessageBox(Me, ex.ToString)
                        End If
                    End If

                Else
                    If ddlResult.SelectedIndex <> 0 AndAlso v_ddlResult <> "" Then
                        Call Utl_SaveDataRow(eItem, drOC, sErrmsg, oTrans)
                        If sErrmsg <> "" Then
                            'sErrmsg=Replace(sErrmsg, vbCrLf, "<br>" & vbCrLf)
                            Call TIMS.WriteTraceLog(sErrmsg)
                            DbAccess.RollbackTrans(oTrans)
                            Call TIMS.CloseDbConn(tConn)
                            Common.MessageBox(Me, cst_errmsg1)
                            Exit Sub
                            'Common.MessageBox(Me, ex.ToString)
                        End If
                    End If
                End If
            Next
            DbAccess.CommitTrans(oTrans)

        Catch ex As Exception
            Dim strErrmsg As String = ""
            strErrmsg &= String.Format("ex.Message: {0}", ex.Message) & vbCrLf
            strErrmsg &= String.Format("ex.ToString: {0}", ex.ToString) & vbCrLf
            strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)
            DbAccess.RollbackTrans(oTrans)
            Call TIMS.CloseDbConn(tConn)
            Common.MessageBox(Me, cst_errmsg1)
            Exit Sub
            'Common.MessageBox(Me, ex.ToString)
        End Try
        Call TIMS.CloseDbConn(tConn)

        If Hid_ChangeWating.Value = "Y" Then
            If ViewState(cst_watingcount) > 0 Then
                '備取資料
                'strSETIDALL=Left(strSETIDALL, Len(strSETIDALL) - 1)
                'sql="select Name,SETID from Stud_EnterTemp where SETID in (" & strSETIDALL & ")"
                Dim sql As String = ""
                sql = ""
                sql &= " SELECT se.SETID" & vbCrLf
                sql &= " ,se.Name" & vbCrLf
                sql &= " ,CASE WHEN st.selsort IS NOT NULL THEN CONVERT(VARCHAR, st.selsort) ELSE '未設定' END selsort" & vbCrLf
                sql &= " FROM STUD_ENTERTEMP se" & vbCrLf
                sql &= " JOIN STUD_SELRESULT st ON se.SETID=st.SETID" & vbCrLf
                sql &= " WHERE OCID='" & ViewState(cst_OCIDValue1) & "'" & vbCrLf
                sql &= " AND st.SETID IN (" & strSETIDALL & ")" & vbCrLf
                Dim dt As DataTable = Nothing
                dt = DbAccess.GetDataTable(sql, objconn)
                If dt.Rows.Count > 0 Then
                    Table_3.Visible = False
                    table4.Visible = False
                    Datagrid2.Visible = True
                    Datagrid2.DataSource = dt
                    Datagrid2.DataBind()
                    btnsave.Visible = True '備取設定的儲存
                    btncancel.Visible = True '備取設定的取消
                Else
                    Datagrid2.Visible = False
                    btnsave.Visible = False '備取設定的儲存
                    btncancel.Visible = False '備取設定的取消
                    Common.MessageBox(Me, "查無備取學員基本資料!!")
                End If
            ElseIf ViewState(cst_watingcount) = 0 Then
                Datagrid2.Visible = False
                btnsave.Visible = False '備取設定的儲存
                btncancel.Visible = False '備取設定的取消
                Common.MessageBox(Me, "沒有備取學員!!")
            End If
        Else
            If Convert.ToString(vs_Alertmsg) <> "" Then
                Common.MessageBox(Me, "錄取完成!!" & vbCrLf & vs_Alertmsg)
            Else
                Common.MessageBox(Me, "錄取完成!!" & vbCrLf)
            End If
            Call Search1()
        End If
        '#Region "(No Use)"

        'If Not objTrans Is Nothing Then DbAccess.RollbackTrans(objTrans)
        'Call TIMS.CloseDbConn(tConn)
        'Try
        'Catch ex As Exception
        '    Dim strErrmsg As String=""
        '    strErrmsg += "/*  ex.ToString: */" & vbCrLf
        '    strErrmsg += ex.ToString & vbCrLf
        '    strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
        '    strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
        '    Call TIMS.SendMailTest(strErrmsg)
        '    Common.MessageBox(Me, "資料輸入出錯!!")
        '    Common.MessageBox(Me, ex.ToString)
        'End Try
    End Sub

    ''' <summary>
    ''' 儲存-CLASS_CONFIRM (NOLOCK='Y') 儲存但不鎖定
    ''' </summary>
    Sub SaveCONF2()
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        Hid_OCID1.Value = TIMS.ClearSQM(Hid_OCID1.Value)
        If OCIDValue1.Value = "" AndAlso Hid_OCID1.Value = "" Then Return

        '按過後鎖定 true:鎖定  false:不鎖定
        'https://jira.turbotech.com.tw/browse/TIMSC-255
        'CLASS_CONFIRM
        Dim drCF As DataRow = TIMS.GET_CLSCONFIRM(OCIDValue1.Value, objconn)
        Dim flag_USE_NEW_CONFIRM As Boolean = (drCF Is Nothing) '資料要新增

        If Not flag_USE_NEW_CONFIRM Then
            '有資料
            Hid_OCID1.Value = CStr(drCF("OCID"))
            Hid_CFSEQNO.Value = CStr(drCF("CFSEQNO"))
            Hid_CFGUID.Value = CStr(drCF("CFGUID"))
        End If

        Dim iCFSEQNO As Integer = 0 'TIMS.Get_CFSEQNO1(Hid_OCID1.Value, objconn)
        Dim sCFGUID As String = "" 'TIMS.GetGUID()
        If flag_USE_NEW_CONFIRM Then
            '產生一組新的GUID / 取得確認次數
            sCFGUID = TIMS.GetGUID()
            iCFSEQNO = TIMS.Get_CFSEQNO1(Hid_OCID1.Value, objconn)
        Else
            '使用舊資料 如果有的話
            iCFSEQNO = TIMS.CINT1(Hid_CFSEQNO.Value)
            sCFGUID = Hid_CFGUID.Value
        End If

        'ODNUMBER.Text=TIMS.ClearSQM(ODNUMBER.Text)
        '#Region "(No Use)"
        'sql=""
        'sql &= " SELECT STUDMODE" 'A:遞補學員 2/3:離退訓 NULL:正常清單
        'sql &= " FROM STUD_CONFIRM"
        'sql &= " WHERE OCID=@vOCID AND CFSEQNO=@CFSEQNO AND SOCID=@SOCID"

        If Not flag_USE_NEW_CONFIRM Then
            '使用舊資料 如果有的話
            Dim uParms As New Hashtable From {
                {"MODIFYACCT", sm.UserInfo.UserID},
                {"OCID", Hid_OCID1.Value},
                {"CFGUID", sCFGUID},
                {"CFSEQNO", iCFSEQNO}
            }
            Dim u_sql As String = ""
            u_sql &= " UPDATE CLASS_CONFIRM" & vbCrLf
            u_sql &= " SET MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf 'u_sql &= " ,NOLOCK=NULL" & vbCrLf '不管如何還是鎖定
            u_sql &= " WHERE OCID=@OCID AND CFGUID=@CFGUID AND CFSEQNO=@CFSEQNO" & vbCrLf
            'uParms.Clear() 'get_CFSEQNO(Hid_OCID.Value, objConn)
            DbAccess.ExecuteNonQuery(u_sql, objconn, uParms)
        Else
            '寫入一筆新資料
            Dim c_Sql As String = ""
            c_Sql &= " INSERT INTO CLASS_CONFIRM (CFGUID ,OCID, CFSEQNO, ODNUMBER, CREATEACCT, CREATEDATE, CONFIRACCT, CONFIRDATE, MODIFYACCT, MODIFYDATE)" & vbCrLf
            c_Sql &= " VALUES (@CFGUID ,@OCID ,@CFSEQNO ,NULL ,@CREATEACCT ,GETDATE() ,@CONFIRACCT ,GETDATE() ,@MODIFYACCT ,GETDATE())" & vbCrLf
            'cParms.Clear()
            'cParms.Add("ODNUMBER", SqlDbType.VarChar).Value=ODNUMBER.Text
            'cParms.Add("ODNUMBER", SqlDbType.VarChar).Value="N"
            Dim cParms As New Hashtable From {
                {"CFGUID", sCFGUID},
                {"OCID", Hid_OCID1.Value},
                {"CFSEQNO", iCFSEQNO}, 'get_CFSEQNO(Hid_OCID.Value, objConn)
                {"CREATEACCT", sm.UserInfo.UserID},
                {"CONFIRACCT", sm.UserInfo.UserID}, '確認者 
                {"MODIFYACCT", sm.UserInfo.UserID}
            }
            DbAccess.ExecuteNonQuery(c_Sql, objconn, cParms)
        End If

        '--STUD_CONFIRM--
        Dim ss_Sql As String = ""
        ss_Sql &= " SELECT SFID FROM STUD_CONFIRM" & vbCrLf
        ss_Sql &= " WHERE OCID=@OCID AND CFSEQNO=@CFSEQNO" & vbCrLf
        ss_Sql &= " AND SETID=@SETID AND ENTERDATE=@ENTERDATE AND SERNUM=@SERNUM" & vbCrLf

        '依錄取學員 儲存參訓學員名單
        For Each eItem As DataGridItem In DataGrid1.Items
            Dim Hid_SETID As HiddenField = eItem.FindControl("Hid_SETID")
            Dim Hid_EnterDate As HiddenField = eItem.FindControl("Hid_EnterDate")
            Dim Hid_SerNum As HiddenField = eItem.FindControl("Hid_SerNum")
            'Dim SOCID As HtmlInputHidden=eItem.FindControl("SOCID")
            'Dim HStudStatus As HtmlInputHidden=eItem.FindControl("HStudStatus")

            'sParms.Clear()
            Dim sParms As New Hashtable From {
                {"OCID", Hid_OCID1.Value},
                {"CFSEQNO", iCFSEQNO}, 'get_CFSEQNO(Hid_OCID.Value, objConn)
                {"SETID", Hid_SETID.Value},
                {"EnterDate", TIMS.Cdate2(Hid_EnterDate.Value)},
                {"SERNUM", Hid_SerNum.Value}
            }
            Dim dtS As DataTable = DbAccess.GetDataTable(ss_Sql, objconn, sParms)

            Dim iSFID As Integer = 0
            If dtS.Rows.Count = 0 Then
                '新增-INSERT
                iSFID = DbAccess.GetNewId(objconn, "STUD_CONFIRM_SFID_SEQ,STUD_CONFIRM,SFID")
                'sParms.Clear()
                Dim isParms As New Hashtable From {
                    {"SFID", iSFID},
                    {"OCID", Hid_OCID1.Value},
                    {"CFSEQNO", iCFSEQNO}, 'get_CFSEQNO(Hid_OCID.Value, objConn)
                    {"SETID", Hid_SETID.Value},
                    {"EnterDate", TIMS.Cdate2(Hid_EnterDate.Value)},
                    {"SERNUM", Hid_SerNum.Value},
                    {"MODIFYACCT", sm.UserInfo.UserID}
                }
                Dim si_Sql As String = ""
                si_Sql &= " INSERT INTO STUD_CONFIRM (SFID ,OCID ,CFSEQNO ,SETID ,ENTERDATE ,SERNUM ,MODIFYACCT ,MODIFYDATE)" & vbCrLf
                si_Sql &= " VALUES (@SFID ,@OCID ,@CFSEQNO ,@SETID ,@ENTERDATE ,@SERNUM ,@MODIFYACCT ,GETDATE())" & vbCrLf
                DbAccess.ExecuteNonQuery(si_Sql, objconn, isParms)
            Else
                'UPDATE-有資料
                Dim drS As DataRow = dtS.Rows(0)
                iSFID = TIMS.CINT1(drS("SFID"))
                Dim usParms As New Hashtable From {{"MODIFYACCT", sm.UserInfo.UserID}, {"SFID", iSFID}}
                Dim su_Sql As String = ""
                su_Sql &= " UPDATE STUD_CONFIRM" & vbCrLf
                su_Sql &= " SET MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
                su_Sql &= " WHERE SFID=@SFID" & vbCrLf
                DbAccess.ExecuteNonQuery(su_Sql, objconn, usParms)
            End If
        Next
    End Sub

    ''' <summary>列印</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub btnPrint1_Click(sender As Object, e As EventArgs) Handles btnPrint1.Click
        flag_no_use_print_btn = If(TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1, True, False)
        If flag_no_use_print_btn Then
            'btnPrint1.Enabled=False
            'TIMS.Tooltip(btnPrint1, cst_msg0009)
            btnPrint1.Visible = False
            lab_msg_r1.Text = cst_msg0017 ' cst_msg0009
            Return
        End If
        'If flag_no_use_print_btn Then
        '    Common.MessageBox(Me, cst_msg0009)
        '    btnPrint1.Enabled=False
        '    Exit Sub
        'End If

        Hid_OCID1.Value = TIMS.ClearSQM(Hid_OCID1.Value)
        Hid_CFGUID.Value = TIMS.ClearSQM(Hid_CFGUID.Value)
        Hid_CFSEQNO.Value = TIMS.ClearSQM(Hid_CFSEQNO.Value)
        If Hid_CFGUID.Value = "" Then
            Common.MessageBox(Me, cst_msg0008)
            Exit Sub
        End If
        If Hid_CFSEQNO.Value = "" Then
            Common.MessageBox(Me, cst_msg0008)
            Exit Sub
        End If
        Dim myValue As String = "" 'myValue=""
        myValue &= "&OCID=" & Hid_OCID1.Value
        myValue &= "&CFGUID=" & Hid_CFGUID.Value
        myValue &= "&CFSEQNO=" & Hid_CFSEQNO.Value
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, myValue)
    End Sub

    ''' <summary>
    ''' 送出-解鎖 btnSAVE2 暫不提供解鎖功能 CLASS_CONFIRM
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub btnSAVE2_Click(sender As Object, e As EventArgs) Handles btnSAVE2.Click
        '解鎖-檢核
        'Dim drOC As DataRow
        Dim Errmsg As String = ""
        Hid_OCID1.Value = TIMS.ClearSQM(Hid_OCID1.Value)
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If Hid_OCID1.Value = "" OrElse OCIDValue1.Value = "" Then
            Errmsg &= "班級選擇有誤!!" & vbCrLf
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If
        If Hid_OCID1.Value <> OCIDValue1.Value Then
            Errmsg &= "班級選擇有誤!!" & vbCrLf
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If
        Dim drOC As DataRow = TIMS.GetOCIDDate(Hid_OCID1.Value, objconn)
        If drOC Is Nothing Then
            Errmsg &= "班級選擇有誤!!" & vbCrLf
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If

        'CLASS_CONFIRM
        Dim drCF As DataRow = TIMS.GET_CLSCONFIRM(OCIDValue1.Value, objconn)
        If drCF Is Nothing Then
            Errmsg &= "班級選擇有誤!!" & vbCrLf
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If
        '有資料
        Hid_OCID1.Value = CStr(drCF("OCID"))
        Hid_CFGUID.Value = CStr(drCF("CFGUID"))
        Hid_CFSEQNO.Value = CStr(drCF("CFSEQNO"))

        '解鎖-最後一筆解鎖
        'uParms.Clear()
        Dim uParms As New Hashtable From {
            {"OCID", OCIDValue1.Value},
            {"CFGUID", Hid_CFGUID.Value},
            {"CFSEQNO", Hid_CFSEQNO.Value},
            {"MODIFYACCT", sm.UserInfo.UserID}
        }
        Dim sql As String = ""
        sql &= " UPDATE CLASS_CONFIRM" & vbCrLf
        sql &= " SET [NOLOCK]='Y'" & vbCrLf 'NO LOCK 不鎖定!
        sql &= " ,MODIFYACCT=@MODIFYACCT" & vbCrLf
        sql &= " ,MODIFYDATE=GETDATE()" & vbCrLf
        sql &= " WHERE OCID=@OCID" & vbCrLf
        sql &= " AND CFGUID=@CFGUID" & vbCrLf
        sql &= " AND CFSEQNO=@CFSEQNO" & vbCrLf
        sql &= " AND [NOLOCK] IS NULL" & vbCrLf
        DbAccess.ExecuteNonQuery(sql, objconn, uParms)
    End Sub

#Region "(No Use)"
    '挑選其他志願
    'Private Sub Button3_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button3.ServerClick
    '    'SD_02_003_other.aspx 'JS function choose_other(){
    '    'ViewState("open")=1
    '    Call Search1()
    'End Sub

    '職訓e網將顯示甄試的錄取結果與成績結果
    'Sub Check_Button6(ByVal OKflag As Boolean)
    '    If button6.Visible Then
    '        If OKflag Then
    '            button6.Enabled=False
    '            TIMS.Tooltip(button6, TIMS.Cst_GPRJName1 & "將顯示甄試的錄取結果與成績結果", True)
    '        Else
    '            button6.Enabled=True
    '            TIMS.Tooltip(button6, TIMS.Cst_GPRJName1 & "暫不顯示甄試的錄取結果與成績結果(公告後將顯示)", True)
    '        End If
    '    End If
    'End Sub

    ''#Region "TPlanID36_NoUse"
    '    Sub Search36()
    '    End Sub
    ' 

    '備取名次設定
    'Private Sub BtnWatingSet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnWatingSet.Click
    '    'Dim watingcount As Integer=0
    '    Dim myitem As DataGridItem
    '    Dim sql As String=""
    '    Dim strSETID As String=""
    '    Dim strSETIDALL As String=""
    '    Dim Result As DropDownList
    '    Dim dt As DataTable
    '    Dim dr As DataRow
    '    ViewState(cst_watingcount)=0
    '    For Each myitem In Me.DataGrid1.Items
    '        Result=myitem.Cells(8).Controls(1)    '是否錄取
    '        strSETID=myitem.Cells(10).Text        'SETID
    '        If Result.SelectedValue="02" Then     '若是備取
    '            ViewState(cst_watingcount)=ViewState(cst_watingcount) + 1
    '            strSETIDALL += strSETID + ","
    '        End If
    '    Next
    '    If ViewState(cst_watingcount) >= 1 Then
    '        strSETIDALL=Left(strSETIDALL, Len(strSETIDALL) - 1)
    '        sql="select Name from Stud_EnterTemp where SETID in (" & strSETIDALL & ")"
    '        dt=DbAccess.GetDataTable(sql)
    '        If dt.Rows.Count > 0 Then
    '            Table3.Visible=False
    '            Table4.Visible=False
    '            Datagrid2.Visible=True
    '            Datagrid2.DataSource=dt
    '            Datagrid2.DataBind()
    '        Else
    '            Common.MessageBox(Me, "查無備取學員基本資料!!")
    '        End If
    '    ElseIf ViewState(cst_watingcount)=0 Then
    '        Common.MessageBox(Me, "沒有備取學員無法設定!!")
    '    End If
    'End Sub

    'e網公告
    'Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button6.Click
    '    If Convert.ToString(sm.UserInfo.UserID)="" Then
    '        Exit Sub
    '    End If
    '    If ViewState(cst_OCIDValue1) Is Nothing Then
    '        Common.MessageBox(Me, "班級選擇有誤!!")
    '        Exit Sub
    '    End If
    '    If Convert.ToString(ViewState(cst_OCIDValue1))="" _
    '        OrElse Convert.ToString(ViewState(cst_OCIDValue1))="0" Then
    '        Common.MessageBox(Me, "班級選擇有誤!!")
    '        Exit Sub
    '    End If
    '    If ViewState(cst_OCIDValue1) <> OCIDValue1.Value Then
    '        Common.MessageBox(Me, "班級選擇有誤!!")
    '        Exit Sub
    '    End If

    '    Dim OKflag As Boolean=False
    '    Dim sqlStr As String=""
    '    sqlStr=""
    '    sqlStr &= " UPDATE Class_ClassInfo" & vbCrLf
    '    sqlStr &= " SET ModifyDate=getdate()" & vbCrLf
    '    sqlStr &= " ,ModifyAcct='" & sm.UserInfo.UserID & "'" & vbCrLf
    '    sqlStr &= " ,eTrain_Show='Y'" & vbCrLf
    '    sqlStr &= " WHERE OCID=@OCIDValue1" & vbCrLf
    '    Dim sCmd As New SqlCommand(sqlStr, objconn)
    '    Call TIMS.OpenDbConn(objconn)
    '    Try
    '        With sCmd
    '            .Parameters.Clear()
    '            .Parameters.Add("@OCIDValue1", SqlDbType.Int).Value=ViewState(cst_OCIDValue1)
    '            .ExecuteNonQuery()
    '        End With
    '        OKflag=True
    '    Catch ex As Exception
    '        OKflag=False
    '        Dim strErrmsg As String=""
    '        strErrmsg += "/*  ex.ToString: */" & vbCrLf
    '        strErrmsg += ex.ToString & vbCrLf
    '        strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
    '        strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
    '        Call TIMS.SendMailTest(strErrmsg)
    '    End Try
    '    Call Check_Button6(OKflag)
    'End Sub

    '判斷備取名次設定是否有重復 'False:沒有重複，True:重複
    'Function sUtl_DoubleCheck1() As Boolean
    '    Dim rst As Boolean=False
    '    Dim strALL As String=""
    '    Dim strS As String=""
    '    Dim i As Integer=1
    '    '-------------判斷備取名次設定是否有重復------------------
    '    For Each Ditem As DataGridItem In Datagrid2.Items
    '        Dim WatingSort As DropDownList=Ditem.Cells(2).Controls(1)
    '        If WatingSort.SelectedIndex > 0 Then   '排除沒選名次的
    '            strS="'" & WatingSort.SelectedValue & "',"   '取備取名次的值
    '            If i=1 Then        '第一筆資料
    '                strALL += "'" & WatingSort.SelectedValue & "',"   '比對的字串
    '            ElseIf i >= 2 Then   '第二筆資料才開始比對
    '                If InStr(strALL, strS) <= 0 Then    '沒有找到相同的
    '                    strALL += "'" & WatingSort.SelectedValue & "',"    '比對的字串
    '                Else                                '有找到相同的
    '                    Common.MessageBox(Me, "備取名次設定不能重覆!!")
    '                    'Exit Function
    '                    rst=True
    '                    Exit For
    '                End If
    '            End If
    '            i=i + 1       '計算次數
    '        End If
    '    Next
    '    '-------------判斷備取名次設定是否有重復-end----------------
    '    Return rst
    'End Function

#End Region

End Class