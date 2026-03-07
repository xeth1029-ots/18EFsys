Partial Class SD_11_004_add17
    Inherits AuthBasePage

    '本程式異動Table: STUD_QUESTIONFAC '受訓學員意見調查表
    '本程式異動Table: STUD_QUESTIONFAC2 '受訓學員意見調查表(2)2017
    '受訓學員意見調查表
    'Dim flagLock As Boolean = False '(解)進行鎖定。
    'Dim sqlAdapter As SqlDataAdapter
    'Dim stud_table As DataTable
    'Dim FunDr As DataRow

    'Dim ff3 As String = ""
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
    'Private LOG As ILog = LogManager.GetLogger("TIMS")
    'Dim au As New cAUTH
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
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objconn) '開啟連線
        '檢查Session是否存在 End

        Call SHOW_LIT_MSG()

        SOCID.Attributes("onchange") = "if(this.selectedIndex==0) return false;"
        'RadioButtonList3_4.Attributes("onclick") = "disable_radio3();"
        'RadioButtonList5_3.Attributes("onclick") = "disable_radio5();"
        'ProcessType.Value = Request("ProcessType")
        'Re_OCID.Value = Request("ocid")
        'Re_SOCID.Value = Request("SOCID")
        'Re_ID.Value = Request("ID")

        ProcessType.Value = TIMS.ClearSQM(Request("ProcessType"))
        Re_OCID.Value = TIMS.ClearSQM(Request("ocid"))
        Dim rqSOCID As String = TIMS.ClearSQM(Request("SOCID"))
        If rqSOCID <> "" AndAlso Re_SOCID.Value = "" Then Re_SOCID.Value = rqSOCID 'TIMS.ClearSQM(Request("SOCID"))
        'If (Re_SOCID.Value <> rqSOCID) Then Re_SOCID.Value = ""
        Re_ID.Value = TIMS.ClearSQM(Request("ID"))
        'Dim s_TK1_DEC As String = TIMS.DecryptAes(Request("TK1"))
        'Dim s_TK1_DEC_SOCID As String = TIMS.GetMyValue(s_TK1_DEC, "SOCID")
        'If s_TK1_DEC <> "" AndAlso (Re_SOCID.Value <> TIMS.GetMyValue(s_TK1_DEC, "SOCID")) Then Re_SOCID.Value = ""
        'If s_TK1_DEC <> "" AndAlso (Re_OCID.Value <> TIMS.GetMyValue(s_TK1_DEC, "ocid")) Then Re_OCID.Value = ""
        'Dim logMsg1 As String = String.Format("SD_11_004_add17: rqSOCID={0},Re_SOCID.Value={1},s_TK1_DEC_SOCID={2},Re_OCID.Value={3}", rqSOCID, Re_SOCID.Value, s_TK1_DEC_SOCID, Re_OCID.Value)
        'LOG.Info(logMsg1)

        Select Case ProcessType.Value
            Case cst_ptEdit
                Button1.Enabled = True
            Case cst_ptInsert
                Button1.Enabled = True '可儲存
            Case cst_ptNext
                Button1.Enabled = True '可儲存
                'MoveNext()
                'If ProcessType.Value = cst_ptNext Then MoveNext()
        End Select

        If Not IsPostBack Then
            'LIST該班級所有學員。
            'Call Create1()
            Call LoadCreateData1()
            Common.SetListItem(SOCID, Re_SOCID.Value) '選定1學員。
            Call CHECK_SESS1()

            '確認學員資料正確性!!
            Dim chkS1 As Boolean = Chk_STUDENTSData1(Re_OCID.Value, Re_SOCID.Value)
            If Not chkS1 Then '(不正確)
                'Common.MessageBox(Me, cst_errmsg2)
                '沒有資料，自動返回。
                Dim strScript As String = ""
                strScript = "<script language=""javascript"">" + vbCrLf
                strScript += "alert('" & cst_errmsg2 & "');" + vbCrLf
                strScript += "location.href ='SD_11_004.aspx?ProcessType=" & cst_ptBack & "&ID='+document.getElementById('Re_ID').value;" + vbCrLf
                strScript += "</script>"
                Page.RegisterStartupScript(TIMS.xBlockName, strScript)
                Exit Sub
            End If

            ''判斷課程所歸屬的計畫**by Milor 20080414----start
            'Dim PNAME As String = ""
            'Select Case Convert.ToString(Request("orgkind"))
            '    Case "G", "W"
            '        PNAME = TIMS.Get_PName28(Me, Convert.ToString(Request("orgkind")), objconn)
            'End Select
            'Label4.Text = PNAME '"產業人才投資計畫" '"提升勞工自主學習計畫"
            'Label5.Text = PNAME '"產業人才投資計畫" '"提升勞工自主學習計畫"
            ''判斷課程所歸屬的計畫**by Milor 20080414----end

            Button1.Visible = True '儲存。
            next_but.Visible = True '移動下一筆。
            Select Case ProcessType.Value
                Case cst_ptDelete '"del"
                    Re_SOCID.Value = TIMS.ClearSQM(Re_SOCID.Value)
                    If Re_SOCID.Value = "" Then
                        Common.MessageBox(Me, "刪除資訊有誤!!。")
                        Exit Sub
                    End If
                    Dim i_DAS As Integer = Chk_DASOURCE(Re_SOCID.Value)
                    Select Case i_DAS
                        Case 1
                            Common.MessageBox(Me, "無法逕行刪除，僅供單位查詢。")
                            Exit Sub
                        Case 2
                        Case Else
                            Common.MessageBox(Me, "刪除資訊有誤!!。")
                            Exit Sub
                    End Select

                    '做刪除動作。
                    Dim dParms1 As New Hashtable
                    dParms1.Add("SOCID", Re_SOCID.Value)
                    Dim sqldel As String = ""
                    sqldel = "DELETE STUD_QUESTIONFAC2 WHERE SOCID=@SOCID"
                    DbAccess.ExecuteNonQuery(sqldel, objconn, dParms1)
                    sqldel = "DELETE STUD_QUESTIONFAC WHERE SOCID=@SOCID"
                    DbAccess.ExecuteNonQuery(sqldel, objconn, dParms1)

                Case cst_ptCheck '"check"
                    Call Shw_DataList1(Re_SOCID.Value) 'modify flgLock
                    Button1.Enabled = False '儲存。
                    next_but.Enabled = False  '移動下一筆。
                    TIMS.Tooltip(Button1, "目前動作不提供儲存。")
                    TIMS.Tooltip(next_but, "目前動作不提供移動下一筆。")
                    Button1.Visible = False '儲存。
                    next_but.Visible = False '移動下一筆。
                    '進行鎖定。
                    SOCID.Enabled = False
                    tb3_Datalist.Disabled = True '進行鎖定。

                Case cst_ptEdit '"Edit" '修改
                    Call Shw_DataList1(Re_SOCID.Value) 'modify flagLock
                    'Button1.Enabled = False '儲存。
                    'next_but.Enabled = False  '移動下一筆。
                    'If Not flagLock Then Button1.Enabled = True '儲存。
                Case cst_ptNext  '"next"
                    '移動下一筆。
                    Call MoveNext()
            End Select
            Button1.Attributes.Add("OnClick", "return ChkData1();")
        End If
    End Sub

    Sub SHOW_LIT_MSG()
        Const cst_title_1 As String = "在職訓練網(https://ojt.wda.gov.tw/StudQuestion)"
        Dim str_Lit_msg_1 As String = ""
        str_Lit_msg_1 &= "本問卷係本署為瞭解參訓學員接受此項訓練的意見和建議，請以打√方式填答每一題項"
        str_Lit_msg_1 &= "，如有其他意見請於第三部份：（二）其他建議欄以文字敘述；您亦可至"
        str_Lit_msg_1 &= cst_title_1
        str_Lit_msg_1 &= "填寫，上網填答的內容僅供本署內部查閱，不對外公開，請安心填寫，感謝您的配合！ "
        Lit_msg_1.Text = str_Lit_msg_1
    End Sub

    'LIST該班級所有學員。
    Sub LoadCreateData1()
        'ProcessType.Value = Request("ProcessType")
        'Re_OCID.Value = Request("ocid")
        'Re_SOCID.Value = Request("SOCID")
        'Re_ID.Value = Request("ID")
        'ProcessType.Value = TIMS.ClearSQM(ProcessType.Value)
        'Re_OCID.Value = TIMS.ClearSQM(Re_OCID.Value)
        'Re_SOCID.Value = TIMS.ClearSQM(Re_SOCID.Value)
        'Re_ID.Value = TIMS.ClearSQM(Re_ID.Value)

        '單位填寫'資料來源	Null:未知 1:報名網()(學員外網填寫。)2@TIMS系統()
        '新增「填寫來源」檢核機制，由系統判斷填寫來源為產投報名網或TIMS系統：
        '1.若為報名網，訓練單位端之各項功能(新增、修改、查詢、列印、清除重填等)將反灰，無法逕行修改，僅保留填寫狀態供訓練單位查詢。
        '2.若為TIMS系統，保留訓練單位各項功能(新增、修改、查詢、列印、清除重填)。
        '3.若有訓練單位先於TIMS系統協助填寫，學員後於報名網修正之情形者，訓練單位端之各項功能(新增、修改、查詢、列印、清除重填等)將反灰。

        Re_OCID.Value = TIMS.ClearSQM(Re_OCID.Value)
        Dim sParms1 As New Hashtable
        sParms1.Add("OCID", Re_OCID.Value)

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT cs.SOCID ,cs.StudentID ,ss.NAME" & vbCrLf
        sql &= " ,dbo.fn_CSTUDID2(cs.STUDENTID) STUDID2 " & vbCrLf
        sql &= " ,ss.NAME + '(' + dbo.fn_CSTUDID2(cs.STUDENTID) + ')' NAME2" & vbCrLf
        sql &= " ,ISNULL(sf.DaSource,0) DaSource " & vbCrLf
        sql &= " FROM CLASS_STUDENTSOFCLASS cs " & vbCrLf
        sql &= " JOIN STUD_STUDENTINFO ss ON ss.SID = cs.SID " & vbCrLf
        sql &= " LEFT JOIN STUD_QUESTIONFAC2 sf ON sf.SOCID = cs.SOCID " & vbCrLf
        sql &= " WHERE cs.OCID=@OCID" & vbCrLf
        sql &= " AND ISNULL(sf.DASOURCE, 0) != 1" & vbCrLf '排除(1:報名網()(學員外網填寫。))
        '資料來源(DASOURCE) 0:未填寫或未知 1: 報名網(學員外網填寫。) 2: TIMS系統 
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, sParms1)
        dt.DefaultView.Sort = "StudentID"
        With SOCID
            .DataSource = dt
            .DataTextField = "NAME2"
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
    Sub Clr_DataList1()
        TIMS.SetCblValue(S1chk, "") '學員基本資料：(一)參加產投方案動機（可複選）：
        S16_NOTE.Text = ""
        S2.SelectedIndex = -1 '學員基本資料：(二)是否為第1次參加產業人才投資方案課程？
        S3.SelectedIndex = -1 '學員基本資料：(三)服務單位員工人數：
        TIMS.SetCblValue(A1chk, "") '第一部份：(一)您獲得本次課程的訊息來源（可複選）：
        A2.SelectedIndex = -1 '第一部份：(二)參加本次課程的主要原因：
        A3.SelectedIndex = -1 '第一部份：(三)選擇本訓練單位的主要原因：
        A4.SelectedIndex = -1 '第一部份：(四)沒有參加本方案訓練之前，每年參加訓練支出的費用？
        A5.SelectedIndex = -1 '第一部份：(五)如果沒有補助訓練費用，你每年願意自費參加訓練課程的金額？
        A6.SelectedIndex = -1 '第一部份：(六)您認為本次課程的訓練費用是否合理？
        A7.SelectedIndex = -1 '第一部份：(七)結訓後對於工作的規劃？
        A1_10_NOTE.Text = ""
        A2_7_NOTE.Text = ""
        A3_5_NOTE.Text = ""
        B11.SelectedIndex = -1 '第二部份：(一)訓練課程1.課程內容符合期望
        B12.SelectedIndex = -1 '第二部份：(一)訓練課程2.課程難易安排適當
        B13.SelectedIndex = -1 '第二部份：(一)訓練課程3.課程總時數適當
        B14.SelectedIndex = -1 '第二部份：(一)訓練課程4.課程符合實務需求
        B15.SelectedIndex = -1 '第二部份：(一)訓練課程5.課程符合產業發展趨勢
        B21.SelectedIndex = -1 '第二部份：(二)講師1.滿意講師的教學態度
        B22.SelectedIndex = -1 '第二部份：(二)講師2.滿意講師的教學方法
        B23.SelectedIndex = -1 '第二部份：(二)講師3.滿意講師的課程專業度
        B31.SelectedIndex = -1 '第二部份：(三)教材1.對於訓練教材感到滿意
        B32.SelectedIndex = -1 '第二部份：(三)教材2.訓練教材能夠輔助課程學習
        B41.SelectedIndex = -1 '第二部份：(四)訓練環境1.您對於訓練場地感到滿意
        B42.SelectedIndex = -1 '第二部份：(四)訓練環境2.您對於訓練設備感到滿意
        B43.SelectedIndex = -1 '第二部份：(四)訓練環境3.您認為實作設備的數量適當
        B44.SelectedIndex = -1 '第二部份：(四)訓練環境4.您認為實作設備新穎
        B51.SelectedIndex = -1 '第二部份：(五)訓練評量:訓練評量（如：訓後測驗、專題報告、作品展示等）能促進學習效果
        B61.SelectedIndex = -1 '第二部份：(六)立即學習效果1.您認為在訓練課程中，課程內容能讓您專注
        B62.SelectedIndex = -1 '第二部份：(六)立即學習效果2.您在完成訓練後，已充份學習訓練課程所教授知識或技能
        B63.SelectedIndex = -1 '第二部份：(六)立即學習效果3.您在完成訓練後，有學習到新的知識或技能
        B71.SelectedIndex = -1 '第二部份：(七)整體意見1.您對於訓練單位的課程安排與授課情形感到滿意
        B72.SelectedIndex = -1 '第二部份：(七)整體意見2.您對於訓練單位的行政服務感到滿意
        B73.SelectedIndex = -1 '第二部份：(七)整體意見3.您對於產業人才投資方案感到滿意
        B74.SelectedIndex = -1 '第二部份：(七)整體意見4.您認為完成本訓練課程對於目前或未來工作有幫助
        C11.SelectedIndex = -1 '第三部份：(一)若本訓練課程沒有補助，是否會全額自費參訓？
        C21_NOTE.Text = ""
    End Sub

    '查詢問卷資料。
    Sub Shw_DataList1(ByVal strSOCID As String)
        Call Clr_DataList1() '清除資料。

        strSOCID = TIMS.ClearSQM(strSOCID)
        If Not TIMS.IsNumeric2(strSOCID) Then strSOCID = ""
        If strSOCID = "" Then Return '無有效資料

        Call TIMS.OpenDbConn(objconn)
        Dim sql As String = ""
        sql = " SELECT * FROM STUD_QUESTIONFAC2 WHERE SOCID = @SOCID "
        If sm.UserInfo.LID = 2 Then
            '資料來源(DASOURCE) 0:未填寫或未知 1: 報名網(學員外網填寫。) 2: TIMS系統 
            sql &= " AND ISNULL(DASOURCE,0) != 1 " & vbCrLf '排除(1:報名網()(學員外網填寫。))
        End If
        Dim sCmd As New SqlCommand(sql, objconn)

        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("SOCID", SqlDbType.VarChar).Value = strSOCID
            dt.Load(.ExecuteReader())
        End With
        If dt.Rows.Count = 0 Then Exit Sub

        Dim dr As DataRow = dt.Rows(0)

        HidDASOURCE.Value = Convert.ToString(dr("DASOURCE"))
        'flagLock:有資料且為學員填寫進行問卷鎖定 / True '進行鎖定。
        Dim flgLock As Boolean = If(Convert.ToString(dr("DASOURCE")) = "1", True, False) '(解)進行鎖定。false:不鎖定 True'進行鎖定。

        '依 flagLock 判斷下列動作
        SOCID.Enabled = If(flgLock, False, True)  '進行鎖定/'(解)進行鎖定
        tb3_Datalist.Disabled = If(flgLock, True, False) '進行鎖定/'(解)進行鎖定

        'If TIMS.sUtl_ChkTest() Then flagLock = True 'flagLock:有資料且為學員填寫進行問卷鎖定。 '測試用

        Dim ff3 As String = ""
        For i As Integer = 1 To 6
            ff3 = "S1" & CStr(i)
            If Convert.ToString(dr(ff3)) = "Y" Then
                Dim oLItem As ListItem = S1chk.Items.FindByValue(CStr(i))
                If Not oLItem Is Nothing Then oLItem.Selected = True
            End If
        Next

        S16_NOTE.Text = Convert.ToString(dr("S16_NOTE"))
        Common.SetListItem(S2, Convert.ToString(dr("S2")))
        Common.SetListItem(S3, Convert.ToString(dr("S3")))

        For i As Integer = 1 To 10
            ff3 = "A1_" & CStr(i)
            If Convert.ToString(dr(ff3)) = "Y" Then
                Dim oLItem As ListItem = A1chk.Items.FindByValue(CStr(i))
                If Not oLItem Is Nothing Then oLItem.Selected = True
            End If
        Next

        Common.SetListItem(A2, Convert.ToString(dr("A2")))
        Common.SetListItem(A3, Convert.ToString(dr("A3")))
        Common.SetListItem(A4, Convert.ToString(dr("A4")))
        Common.SetListItem(A5, Convert.ToString(dr("A5")))
        Common.SetListItem(A6, Convert.ToString(dr("A6")))
        Common.SetListItem(A7, Convert.ToString(dr("A7")))
        A1_10_NOTE.Text = Convert.ToString(dr("A1_10_NOTE"))
        A2_7_NOTE.Text = Convert.ToString(dr("A2_7_NOTE"))
        A3_5_NOTE.Text = Convert.ToString(dr("A3_5_NOTE"))
        Common.SetListItem(B11, Convert.ToString(dr("B11")))
        Common.SetListItem(B12, Convert.ToString(dr("B12")))
        Common.SetListItem(B13, Convert.ToString(dr("B13")))
        Common.SetListItem(B14, Convert.ToString(dr("B14")))
        Common.SetListItem(B15, Convert.ToString(dr("B15")))
        Common.SetListItem(B21, Convert.ToString(dr("B21")))
        Common.SetListItem(B22, Convert.ToString(dr("B22")))
        Common.SetListItem(B23, Convert.ToString(dr("B23")))
        Common.SetListItem(B31, Convert.ToString(dr("B31")))
        Common.SetListItem(B32, Convert.ToString(dr("B32")))
        Common.SetListItem(B41, Convert.ToString(dr("B41")))
        Common.SetListItem(B42, Convert.ToString(dr("B42")))
        Common.SetListItem(B43, Convert.ToString(dr("B43")))
        Common.SetListItem(B44, Convert.ToString(dr("B44")))
        Common.SetListItem(B51, Convert.ToString(dr("B51")))
        Common.SetListItem(B61, Convert.ToString(dr("B61")))
        Common.SetListItem(B62, Convert.ToString(dr("B62")))
        Common.SetListItem(B63, Convert.ToString(dr("B63")))
        Common.SetListItem(B71, Convert.ToString(dr("B71")))
        Common.SetListItem(B72, Convert.ToString(dr("B72")))
        Common.SetListItem(B73, Convert.ToString(dr("B73")))
        Common.SetListItem(B74, Convert.ToString(dr("B74")))
        Common.SetListItem(C11, Convert.ToString(dr("C11")))
        C21_NOTE.Text = Convert.ToString(dr("C21_NOTE"))

        'Button1.Enabled = False '儲存。
        ''next_but.Enabled = False  '移動下一筆。
        'Select Case ProcessType.Value
        '    Case cst_ptEdit, cst_ptNext   '修改' 儲存後移動下一筆。
        '        If Not flagLock Then Button1.Enabled = True '儲存。(編輯模式下進行解鎖)
        'End Select
    End Sub

    '查詢基本資料。某班，某學員基本資料。沒有為false
    Function Chk_STUDENTSData1(ByVal strOCID As String, ByVal strSOCID As String) As Boolean
        Dim rst As Boolean = False
        strOCID = TIMS.ClearSQM(strOCID)
        strSOCID = TIMS.ClearSQM(strSOCID)
        If strOCID = "" Then Return rst
        If strSOCID = "" Then Return rst
        If Not TIMS.IsNumeric2(strOCID) Then Return rst
        If Not TIMS.IsNumeric2(strSOCID) Then Return rst

        Dim sqlstr As String = ""
        sqlstr = "" & vbCrLf
        sqlstr &= " SELECT b.SOCID ,b.studentid ,c.name ,b.StudStatus " & vbCrLf
        sqlstr &= " ,CONVERT(varchar, b.RejectTDate1, 111) RejectTDate1 " & vbCrLf
        sqlstr &= " ,CONVERT(varchar, b.RejectTDate2, 111) RejectTDate2 " & vbCrLf
        sqlstr &= " FROM CLASS_CLASSINFO a WITH(NOLOCK)" & vbCrLf
        sqlstr &= " JOIN CLASS_STUDENTSOFCLASS b WITH(NOLOCK) ON a.ocid = b.ocid " & vbCrLf
        sqlstr &= " JOIN PLAN_PLANINFO p WITH(NOLOCK) ON a.PlanID = p.PlanID AND a.comIDNO = p.comIDNO AND a.SeqNO = p.SeqNO " & vbCrLf
        sqlstr &= " JOIN STUD_STUDENTINFO c WITH(NOLOCK) ON b.sid = c.sid " & vbCrLf
        sqlstr &= " WHERE 1=1 " & vbCrLf
        sqlstr &= " AND b.OCID = '" & strOCID & "' "
        sqlstr &= " AND b.SOCID = '" & strSOCID & "' " & vbCrLf
        Dim row As DataRow
        row = DbAccess.GetOneRow(sqlstr, objconn)
        If row Is Nothing Then Return rst

        rst = True
        Me.Label_Name.Text = Convert.ToString(row("name"))
        Me.Label_Stud.Text = Convert.ToString(row("studentid"))
        Me.Label_Status.Text = TIMS.GET_STUDSTATUS_N23(row("StudStatus"), row("RejectTDate1"), row("RejectTDate2"))
        'Me.Label_Status.Text = sStatus
        'Me.Label1.Text = Convert.ToString(row("totalcost"))
        'Me.Label2.Text = Convert.ToString(row("defstdcost"))
        Return rst
    End Function

    Sub CHECK_SESS1()
        'call CHECK_SESS1()
        'Session("QuestionFacSearchStr") = Me.ViewState("QuestionFacSearchStr")
        If Session(ss_QuestionFacSearchStr) IsNot Nothing Then
            Me.ViewState(ss_QuestionFacSearchStr) = Session(ss_QuestionFacSearchStr)
            Session(ss_QuestionFacSearchStr) = Session(ss_QuestionFacSearchStr)
        Else
            If ViewState(ss_QuestionFacSearchStr) IsNot Nothing Then
                Session(ss_QuestionFacSearchStr) = Me.ViewState(ss_QuestionFacSearchStr)
            End If
        End If
    End Sub


    '回上一頁。
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Call CHECK_SESS1()
        TIMS.CloseDbConn(objconn)
        TIMS.Utl_Redirect1(Me, "SD_11_004.aspx?ProcessType=" & cst_ptBack & "&ID=" & Re_ID.Value & "")
    End Sub

    '移動到最後一筆返回首頁。
    Sub check_last()
        Call CHECK_SESS1()

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
    Sub MoveNext()
        If SOCID.Items.Count > 0 Then
            Dim NowIndex As Integer
            Dim MaxIndex As Integer
            MaxIndex = SOCID.Items.Count - 1
            NowIndex = SOCID.SelectedIndex
            If NowIndex = MaxIndex Then
                Call check_last() '移動到最後一筆返回首頁。
            Else
                SOCID.SelectedIndex = NowIndex + 1
                Re_SOCID.Value = TIMS.GetListValue(SOCID) 'SOCID.SelectedValue
                Call Shw_DataList1(Re_SOCID.Value) 'modify flagLock
                Dim chkS1 As Boolean = Chk_STUDENTSData1(Re_OCID.Value, Re_SOCID.Value) '確認學員資料正確性!!
                If Not chkS1 Then
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
        Re_SOCID.Value = TIMS.GetListValue(SOCID) 'SOCID.SelectedValue
        Call Shw_DataList1(Re_SOCID.Value) 'modify flagLock
        Dim chkS1 As Boolean = Chk_STUDENTSData1(Re_OCID.Value, Re_SOCID.Value) '確認學員資料正確性!!
        If Not chkS1 Then
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

    ''檢查。
    'Function Chk_STUD_QUESTIONFAC2() As Boolean
    '    Dim rst As Boolean = False
    '    Dim msg As String = ""
    '    rst = True
    '    If msg <> "" Then
    '        rst = False
    '        Common.MessageBox(Me, msg)
    '    End If
    '    Return rst
    'End Function

    '儲存前檢查。
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        If tb3_Datalist.Disabled Then
            Errmsg &= "資料已進行鎖定，不可更動!!" & vbCrLf
            Return Errmsg
            'Common.MessageBox(Me, "資料已進行鎖定，不可更動!!") 'Exit Sub
        End If

        Re_SOCID.Value = TIMS.GetListValue(SOCID)
        If Re_SOCID.Value = "" Then
            Errmsg &= "學員資料有誤，請重新選擇。" & vbCrLf
        End If

        Const cst_iC21_NOTE_MaxLen As Integer = 700
        C21_NOTE.Text = TIMS.ClearSQM(C21_NOTE.Text)
        If C21_NOTE.Text <> "" AndAlso Len(C21_NOTE.Text) > cst_iC21_NOTE_MaxLen Then
            'Q12.Text = Trim(Q12.Text)
            Errmsg &= String.Concat("其他建議 長度超過系統範圍(", cst_iC21_NOTE_MaxLen, ")") & vbCrLf
        End If

        HidDASOURCE.Value = Get_HidDASOURCE_VALUE(Re_SOCID.Value)
        If HidDASOURCE.Value = "1" Then Errmsg &= "無法逕行修改，僅供單位查詢。" & vbCrLf
        If Errmsg <> "" Then Return False

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    ''' <summary>'資料來源(DASOURCE) 0:未填寫或未知 1: 報名網(學員外網填寫。) 2: TIMS系統</summary>
    ''' <param name="v_SOCID1"></param>
    Function Get_HidDASOURCE_VALUE(ByVal v_SOCID1 As String) As String
        '資料來源(DASOURCE) 0:未填寫或未知 1: 報名網(學員外網填寫。) 2: TIMS系統 
        If v_SOCID1 = "" Then Return ""

        Dim parms As New Hashtable
        parms.Add("SOCID", v_SOCID1)
        Dim sql As String = " SELECT ISNULL(DASOURCE,0) DASOURCE FROM STUD_QUESTIONFAC2 WHERE SOCID=@SOCID"
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return ""

        Return Convert.ToString(dt.Rows(0)("DASOURCE"))
    End Function

    '儲存
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call CHECK_SESS1()

        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If
        'If tb3_Datalist.Disabled Then
        '    Common.MessageBox(Me, "資料已進行鎖定，不可更動!!")
        '    Exit Sub
        'End If

        '儲存
        Call SaveData1()
        Common.AddClientScript(Page, "insert_next();")

        'Try
        '    '儲存
        '    Call SaveData1()
        '    Common.AddClientScript(Page, "insert_next();")
        'Catch ex As Exception
        '    Common.MessageBox(Me, "!!儲存失敗!!")
        '    'Common.MessageBox(Me, ex.ToString)

        '    Dim strErrmsg As String = ""
        '    strErrmsg = "SD_11_004_add17!!"
        '    strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
        '    strErrmsg &= "/* ex.ToString: */" & vbCrLf
        '    strErrmsg &= ex.ToString & vbCrLf
        '    strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
        '    Call TIMS.SendMailTest(strErrmsg)
        '    Exit Sub
        'End Try

    End Sub

    '儲存
    Sub SaveData1()
        Dim sMODIFYACCT As String = sm.UserInfo.UserID
        Const cst_DASOURCE As String = "2" ' "2" '單位填寫'資料來源	Null:未知 1:報名網()(學員外網填寫。)2@TIMS系統()

        Call TIMS.OpenDbConn(objconn)

        Dim i_sql As String = ""
        Dim u_sql As String = ""
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT 'x' FROM STUD_QUESTIONFAC2 WHERE SOCID =@SOCID" & vbCrLf
        If sm.UserInfo.LID = 2 Then
            '資料來源(DASOURCE) 0:未填寫或未知 1: 報名網(學員外網填寫。) 2: TIMS系統 
            sql &= " AND ISNULL(DASOURCE,0) != 1" & vbCrLf '排除(1:報名網()(學員外網填寫。))
        End If
        Dim sCmd As New SqlCommand(sql, objconn)

        sql = ""
        sql &= " INSERT INTO STUD_QUESTIONFAC2 (" & vbCrLf
        sql &= " SOCID" & vbCrLf
        sql &= " ,S11" & vbCrLf
        sql &= " ,S12" & vbCrLf
        sql &= " ,S13" & vbCrLf
        sql &= " ,S14" & vbCrLf
        sql &= " ,S15" & vbCrLf
        sql &= " ,S16" & vbCrLf
        sql &= " ,S16_NOTE" & vbCrLf
        sql &= " ,S2" & vbCrLf
        sql &= " ,S3" & vbCrLf
        sql &= " ,A1_1" & vbCrLf
        sql &= " ,A1_2" & vbCrLf
        sql &= " ,A1_3" & vbCrLf
        sql &= " ,A1_4" & vbCrLf
        sql &= " ,A1_5" & vbCrLf
        sql &= " ,A1_6" & vbCrLf
        sql &= " ,A1_7" & vbCrLf
        sql &= " ,A1_8" & vbCrLf
        sql &= " ,A1_9" & vbCrLf
        sql &= " ,A1_10" & vbCrLf
        sql &= " ,A1_10_NOTE" & vbCrLf
        sql &= " ,A2" & vbCrLf
        sql &= " ,A3" & vbCrLf
        sql &= " ,A4" & vbCrLf
        sql &= " ,A5" & vbCrLf
        sql &= " ,A6" & vbCrLf
        sql &= " ,A7" & vbCrLf
        sql &= " ,B11" & vbCrLf
        sql &= " ,B12" & vbCrLf
        sql &= " ,B13" & vbCrLf
        sql &= " ,B14" & vbCrLf
        sql &= " ,B15" & vbCrLf
        sql &= " ,B21" & vbCrLf
        sql &= " ,B22" & vbCrLf
        sql &= " ,B23" & vbCrLf
        sql &= " ,B31" & vbCrLf
        sql &= " ,B32" & vbCrLf
        sql &= " ,B41" & vbCrLf
        sql &= " ,B42" & vbCrLf
        sql &= " ,B43" & vbCrLf
        sql &= " ,B44" & vbCrLf
        sql &= " ,B51" & vbCrLf
        sql &= " ,B61" & vbCrLf
        sql &= " ,B62" & vbCrLf
        sql &= " ,B63" & vbCrLf
        sql &= " ,B71" & vbCrLf
        sql &= " ,B72" & vbCrLf
        sql &= " ,B73" & vbCrLf
        sql &= " ,B74" & vbCrLf
        sql &= " ,C11" & vbCrLf
        sql &= " ,C21_NOTE" & vbCrLf
        sql &= " ,A2_7_NOTE" & vbCrLf
        sql &= " ,A3_5_NOTE" & vbCrLf
        sql &= " ,MODIFYACCT" & vbCrLf
        sql &= " ,MODIFYDATE" & vbCrLf
        sql &= " ,DASOURCE" & vbCrLf

        sql &= " ) VALUES (" & vbCrLf
        sql &= " @SOCID" & vbCrLf
        sql &= " ,@S11" & vbCrLf
        sql &= " ,@S12" & vbCrLf
        sql &= " ,@S13" & vbCrLf
        sql &= " ,@S14" & vbCrLf
        sql &= " ,@S15" & vbCrLf
        sql &= " ,@S16" & vbCrLf
        sql &= " ,@S16_NOTE" & vbCrLf
        sql &= " ,@S2" & vbCrLf
        sql &= " ,@S3" & vbCrLf
        sql &= " ,@A1_1" & vbCrLf
        sql &= " ,@A1_2" & vbCrLf
        sql &= " ,@A1_3" & vbCrLf
        sql &= " ,@A1_4" & vbCrLf
        sql &= " ,@A1_5" & vbCrLf
        sql &= " ,@A1_6" & vbCrLf
        sql &= " ,@A1_7" & vbCrLf
        sql &= " ,@A1_8" & vbCrLf
        sql &= " ,@A1_9" & vbCrLf
        sql &= " ,@A1_10" & vbCrLf
        sql &= " ,@A1_10_NOTE" & vbCrLf
        sql &= " ,@A2" & vbCrLf
        sql &= " ,@A3" & vbCrLf
        sql &= " ,@A4" & vbCrLf
        sql &= " ,@A5" & vbCrLf
        sql &= " ,@A6" & vbCrLf
        sql &= " ,@A7" & vbCrLf
        sql &= " ,@B11" & vbCrLf
        sql &= " ,@B12" & vbCrLf
        sql &= " ,@B13" & vbCrLf
        sql &= " ,@B14" & vbCrLf
        sql &= " ,@B15" & vbCrLf
        sql &= " ,@B21" & vbCrLf
        sql &= " ,@B22" & vbCrLf
        sql &= " ,@B23" & vbCrLf
        sql &= " ,@B31" & vbCrLf
        sql &= " ,@B32" & vbCrLf
        sql &= " ,@B41" & vbCrLf
        sql &= " ,@B42" & vbCrLf
        sql &= " ,@B43" & vbCrLf
        sql &= " ,@B44" & vbCrLf
        sql &= " ,@B51" & vbCrLf
        sql &= " ,@B61" & vbCrLf
        sql &= " ,@B62" & vbCrLf
        sql &= " ,@B63" & vbCrLf
        sql &= " ,@B71" & vbCrLf
        sql &= " ,@B72" & vbCrLf
        sql &= " ,@B73" & vbCrLf
        sql &= " ,@B74" & vbCrLf
        sql &= " ,@C11" & vbCrLf
        sql &= " ,@C21_NOTE" & vbCrLf
        sql &= " ,@A2_7_NOTE" & vbCrLf
        sql &= " ,@A3_5_NOTE" & vbCrLf
        sql &= " ,@MODIFYACCT" & vbCrLf
        sql &= " ,GETDATE()" & vbCrLf
        sql &= " ,@DASOURCE" & vbCrLf ' "2" '單位填寫'資料來源	Null:未知 1:報名網()(學員外網填寫。)2@TIMS系統()
        sql &= " )" & vbCrLf
        i_sql = sql
        Dim iCmd As New SqlCommand(sql, objconn)

        'Dim sql As String = ""
        sql = ""
        sql &= " UPDATE STUD_QUESTIONFAC2"
        sql &= " SET S11=@S11" & vbCrLf
        sql &= " ,S12=@S12" & vbCrLf
        sql &= " ,S13=@S13" & vbCrLf
        sql &= " ,S14=@S14" & vbCrLf
        sql &= " ,S15=@S15" & vbCrLf
        sql &= " ,S16=@S16" & vbCrLf
        sql &= " ,S16_NOTE=@S16_NOTE" & vbCrLf
        sql &= " ,S2=@S2" & vbCrLf
        sql &= " ,S3=@S3" & vbCrLf

        sql &= " ,A1_1=@A1_1" & vbCrLf
        sql &= " ,A1_2=@A1_2" & vbCrLf
        sql &= " ,A1_3=@A1_3" & vbCrLf
        sql &= " ,A1_4=@A1_4" & vbCrLf
        sql &= " ,A1_5=@A1_5" & vbCrLf
        sql &= " ,A1_6=@A1_6" & vbCrLf
        sql &= " ,A1_7=@A1_7" & vbCrLf
        sql &= " ,A1_8=@A1_8" & vbCrLf
        sql &= " ,A1_9=@A1_9" & vbCrLf
        sql &= " ,A1_10=@A1_10" & vbCrLf
        sql &= " ,A1_10_NOTE=@A1_10_NOTE" & vbCrLf
        sql &= " ,A2=@A2" & vbCrLf
        sql &= " ,A3=@A3" & vbCrLf
        sql &= " ,A4=@A4" & vbCrLf
        sql &= " ,A5=@A5" & vbCrLf
        sql &= " ,A6=@A6" & vbCrLf
        sql &= " ,A7=@A7" & vbCrLf

        sql &= " ,B11=@B11" & vbCrLf
        sql &= " ,B12=@B12" & vbCrLf
        sql &= " ,B13=@B13" & vbCrLf
        sql &= " ,B14=@B14" & vbCrLf
        sql &= " ,B15=@B15" & vbCrLf

        sql &= " ,B21=@B21" & vbCrLf
        sql &= " ,B22=@B22" & vbCrLf
        sql &= " ,B23=@B23" & vbCrLf

        sql &= " ,B31=@B31" & vbCrLf
        sql &= " ,B32=@B32" & vbCrLf
        sql &= " ,B41=@B41" & vbCrLf
        sql &= " ,B42=@B42" & vbCrLf
        sql &= " ,B43=@B43" & vbCrLf
        sql &= " ,B44=@B44" & vbCrLf

        sql &= " ,B51=@B51" & vbCrLf
        sql &= " ,B61=@B61" & vbCrLf
        sql &= " ,B62=@B62" & vbCrLf
        sql &= " ,B63=@B63" & vbCrLf
        sql &= " ,B71=@B71" & vbCrLf
        sql &= " ,B72=@B72" & vbCrLf
        sql &= " ,B73=@B73" & vbCrLf
        sql &= " ,B74=@B74" & vbCrLf
        sql &= " ,C11=@C11" & vbCrLf

        sql &= " ,C21_NOTE=@C21_NOTE" & vbCrLf
        sql &= " ,A2_7_NOTE=@A2_7_NOTE" & vbCrLf
        sql &= " ,A3_5_NOTE=@A3_5_NOTE" & vbCrLf

        sql &= " ,MODIFYACCT=@MODIFYACCT" & vbCrLf
        sql &= " ,MODIFYDATE=GETDATE()" & vbCrLf
        sql &= " ,DASOURCE=@DASOURCE" & vbCrLf '"2" '單位填寫'資料來源	Null:未知 1:報名網()(學員外網填寫。)2@TIMS系統()
        sql &= " WHERE SOCID=@SOCID" & vbCrLf
        If sm.UserInfo.LID = 2 Then
            '資料來源(DASOURCE) 0:未填寫或未知 1: 報名網(學員外網填寫。) 2: TIMS系統 
            sql &= " AND ISNULL(DASOURCE,0) != 1 " & vbCrLf '排除(1:報名網()(學員外網填寫。))
        End If
        u_sql = sql
        Dim uCmd As New SqlCommand(sql, objconn)

        Dim v_SOCID As String = TIMS.GetListValue(SOCID)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("SOCID", SqlDbType.VarChar).Value = TIMS.GetValue1(v_SOCID)
            dt.Load(.ExecuteReader())
        End With

        If dt.Rows.Count = 0 Then
            With iCmd
                .Parameters.Clear()
                .Parameters.Add("SOCID", SqlDbType.VarChar).Value = TIMS.GetValue1(v_SOCID)
                .Parameters.Add("S11", SqlDbType.VarChar).Value = If(S1chk.Items(0).Selected, "Y", Convert.DBNull)
                .Parameters.Add("S12", SqlDbType.VarChar).Value = If(S1chk.Items(1).Selected, "Y", Convert.DBNull)
                .Parameters.Add("S13", SqlDbType.VarChar).Value = If(S1chk.Items(2).Selected, "Y", Convert.DBNull)
                .Parameters.Add("S14", SqlDbType.VarChar).Value = If(S1chk.Items(3).Selected, "Y", Convert.DBNull)
                .Parameters.Add("S15", SqlDbType.VarChar).Value = If(S1chk.Items(4).Selected, "Y", Convert.DBNull)
                .Parameters.Add("S16", SqlDbType.VarChar).Value = If(S1chk.Items(5).Selected, "Y", Convert.DBNull)
                .Parameters.Add("S16_NOTE", SqlDbType.VarChar).Value = TIMS.GetValue1(S16_NOTE.Text)

                .Parameters.Add("S2", SqlDbType.VarChar).Value = TIMS.GetValue1(TIMS.GetListValue(S2)) '.SelectedValue)
                .Parameters.Add("S3", SqlDbType.VarChar).Value = TIMS.GetValue1(TIMS.GetListValue(S3)) '.SelectedValue)

                .Parameters.Add("A1_1", SqlDbType.VarChar).Value = If(A1chk.Items(0).Selected, "Y", Convert.DBNull)
                .Parameters.Add("A1_2", SqlDbType.VarChar).Value = If(A1chk.Items(1).Selected, "Y", Convert.DBNull)
                .Parameters.Add("A1_3", SqlDbType.VarChar).Value = If(A1chk.Items(2).Selected, "Y", Convert.DBNull)
                .Parameters.Add("A1_4", SqlDbType.VarChar).Value = If(A1chk.Items(3).Selected, "Y", Convert.DBNull)
                .Parameters.Add("A1_5", SqlDbType.VarChar).Value = If(A1chk.Items(4).Selected, "Y", Convert.DBNull)
                .Parameters.Add("A1_6", SqlDbType.VarChar).Value = If(A1chk.Items(5).Selected, "Y", Convert.DBNull)
                .Parameters.Add("A1_7", SqlDbType.VarChar).Value = If(A1chk.Items(6).Selected, "Y", Convert.DBNull)
                .Parameters.Add("A1_8", SqlDbType.VarChar).Value = If(A1chk.Items(7).Selected, "Y", Convert.DBNull)
                .Parameters.Add("A1_9", SqlDbType.VarChar).Value = If(A1chk.Items(8).Selected, "Y", Convert.DBNull)
                .Parameters.Add("A1_10", SqlDbType.VarChar).Value = If(A1chk.Items(9).Selected, "Y", Convert.DBNull)
                .Parameters.Add("A1_10_NOTE", SqlDbType.VarChar).Value = TIMS.GetValue1(A1_10_NOTE.Text)

                .Parameters.Add("A2", SqlDbType.VarChar).Value = TIMS.GetValue1(A2.SelectedValue)
                .Parameters.Add("A3", SqlDbType.VarChar).Value = TIMS.GetValue1(A3.SelectedValue) 'A3
                .Parameters.Add("A4", SqlDbType.VarChar).Value = TIMS.GetValue1(A4.SelectedValue) 'A4
                .Parameters.Add("A5", SqlDbType.VarChar).Value = TIMS.GetValue1(A5.SelectedValue) 'A5
                .Parameters.Add("A6", SqlDbType.VarChar).Value = TIMS.GetValue1(A6.SelectedValue) 'A6
                .Parameters.Add("A7", SqlDbType.VarChar).Value = TIMS.GetValue1(A7.SelectedValue) 'A7

                .Parameters.Add("B11", SqlDbType.VarChar).Value = TIMS.GetValue1(B11.SelectedValue) 'B11
                .Parameters.Add("B12", SqlDbType.VarChar).Value = TIMS.GetValue1(B12.SelectedValue) 'B12
                .Parameters.Add("B13", SqlDbType.VarChar).Value = TIMS.GetValue1(B13.SelectedValue) 'B13
                .Parameters.Add("B14", SqlDbType.VarChar).Value = TIMS.GetValue1(B14.SelectedValue) 'B14
                .Parameters.Add("B15", SqlDbType.VarChar).Value = TIMS.GetValue1(B15.SelectedValue) 'B15

                .Parameters.Add("B21", SqlDbType.VarChar).Value = TIMS.GetValue1(B21.SelectedValue) 'B21
                .Parameters.Add("B22", SqlDbType.VarChar).Value = TIMS.GetValue1(B22.SelectedValue) 'B22
                .Parameters.Add("B23", SqlDbType.VarChar).Value = TIMS.GetValue1(B23.SelectedValue) 'B23

                .Parameters.Add("B31", SqlDbType.VarChar).Value = TIMS.GetValue1(B31.SelectedValue) 'B31
                .Parameters.Add("B32", SqlDbType.VarChar).Value = TIMS.GetValue1(B32.SelectedValue) 'B32
                .Parameters.Add("B41", SqlDbType.VarChar).Value = TIMS.GetValue1(B41.SelectedValue) 'B41
                .Parameters.Add("B42", SqlDbType.VarChar).Value = TIMS.GetValue1(B42.SelectedValue) 'B42
                .Parameters.Add("B43", SqlDbType.VarChar).Value = TIMS.GetValue1(B43.SelectedValue) 'B43
                .Parameters.Add("B44", SqlDbType.VarChar).Value = TIMS.GetValue1(B44.SelectedValue) 'B44

                .Parameters.Add("B51", SqlDbType.VarChar).Value = TIMS.GetValue1(B51.SelectedValue)
                .Parameters.Add("B61", SqlDbType.VarChar).Value = TIMS.GetValue1(B61.SelectedValue)
                .Parameters.Add("B62", SqlDbType.VarChar).Value = TIMS.GetValue1(B62.SelectedValue)
                .Parameters.Add("B63", SqlDbType.VarChar).Value = TIMS.GetValue1(B63.SelectedValue)
                .Parameters.Add("B71", SqlDbType.VarChar).Value = TIMS.GetValue1(B71.SelectedValue)
                .Parameters.Add("B72", SqlDbType.VarChar).Value = TIMS.GetValue1(B72.SelectedValue)
                .Parameters.Add("B73", SqlDbType.VarChar).Value = TIMS.GetValue1(B73.SelectedValue)
                .Parameters.Add("B74", SqlDbType.VarChar).Value = TIMS.GetValue1(B74.SelectedValue)
                .Parameters.Add("C11", SqlDbType.VarChar).Value = TIMS.GetValue1(C11.SelectedValue)

                .Parameters.Add("C21_NOTE", SqlDbType.VarChar).Value = TIMS.GetValue1(C21_NOTE.Text)
                .Parameters.Add("A2_7_NOTE", SqlDbType.VarChar).Value = TIMS.GetValue1(A2_7_NOTE.Text)
                .Parameters.Add("A3_5_NOTE", SqlDbType.VarChar).Value = TIMS.GetValue1(A3_5_NOTE.Text)

                .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sMODIFYACCT 'sm.UserInfo.UserID 'MODIFYACCT
                '.Parameters.Add("MODIFYDATE", SqlDbType.VarChar).Value = MODIFYDATE
                .Parameters.Add("DASOURCE", SqlDbType.VarChar).Value = cst_DASOURCE '單位填寫'資料來源	Null:未知 1:報名網()(學員外網填寫。)2@TIMS系統()
                'dt.Load(.ExecuteReader())'rst = .ExecuteScalar()
                '.ExecuteNonQuery()
            End With
            DbAccess.ExecuteNonQuery(i_sql, objconn, iCmd.Parameters)
        Else
            With uCmd
                .Parameters.Clear()

                .Parameters.Add("S11", SqlDbType.VarChar).Value = If(S1chk.Items(0).Selected, "Y", Convert.DBNull)
                .Parameters.Add("S12", SqlDbType.VarChar).Value = If(S1chk.Items(1).Selected, "Y", Convert.DBNull)
                .Parameters.Add("S13", SqlDbType.VarChar).Value = If(S1chk.Items(2).Selected, "Y", Convert.DBNull)
                .Parameters.Add("S14", SqlDbType.VarChar).Value = If(S1chk.Items(3).Selected, "Y", Convert.DBNull)
                .Parameters.Add("S15", SqlDbType.VarChar).Value = If(S1chk.Items(4).Selected, "Y", Convert.DBNull)
                .Parameters.Add("S16", SqlDbType.VarChar).Value = If(S1chk.Items(5).Selected, "Y", Convert.DBNull)
                .Parameters.Add("S16_NOTE", SqlDbType.VarChar).Value = TIMS.GetValue1(S16_NOTE.Text)
                .Parameters.Add("S2", SqlDbType.VarChar).Value = TIMS.GetValue1(S2.SelectedValue)
                .Parameters.Add("S3", SqlDbType.VarChar).Value = TIMS.GetValue1(S3.SelectedValue)

                .Parameters.Add("A1_1", SqlDbType.VarChar).Value = If(A1chk.Items(0).Selected, "Y", Convert.DBNull)
                .Parameters.Add("A1_2", SqlDbType.VarChar).Value = If(A1chk.Items(1).Selected, "Y", Convert.DBNull)
                .Parameters.Add("A1_3", SqlDbType.VarChar).Value = If(A1chk.Items(2).Selected, "Y", Convert.DBNull)
                .Parameters.Add("A1_4", SqlDbType.VarChar).Value = If(A1chk.Items(3).Selected, "Y", Convert.DBNull)
                .Parameters.Add("A1_5", SqlDbType.VarChar).Value = If(A1chk.Items(4).Selected, "Y", Convert.DBNull)
                .Parameters.Add("A1_6", SqlDbType.VarChar).Value = If(A1chk.Items(5).Selected, "Y", Convert.DBNull)
                .Parameters.Add("A1_7", SqlDbType.VarChar).Value = If(A1chk.Items(6).Selected, "Y", Convert.DBNull)
                .Parameters.Add("A1_8", SqlDbType.VarChar).Value = If(A1chk.Items(7).Selected, "Y", Convert.DBNull)
                .Parameters.Add("A1_9", SqlDbType.VarChar).Value = If(A1chk.Items(8).Selected, "Y", Convert.DBNull)
                .Parameters.Add("A1_10", SqlDbType.VarChar).Value = If(A1chk.Items(9).Selected, "Y", Convert.DBNull)
                .Parameters.Add("A1_10_NOTE", SqlDbType.VarChar).Value = TIMS.GetValue1(A1_10_NOTE.Text)
                .Parameters.Add("A2", SqlDbType.VarChar).Value = TIMS.GetValue1(A2.SelectedValue)
                .Parameters.Add("A3", SqlDbType.VarChar).Value = TIMS.GetValue1(A3.SelectedValue) 'A3
                .Parameters.Add("A4", SqlDbType.VarChar).Value = TIMS.GetValue1(A4.SelectedValue) 'A4
                .Parameters.Add("A5", SqlDbType.VarChar).Value = TIMS.GetValue1(A5.SelectedValue) 'A5
                .Parameters.Add("A6", SqlDbType.VarChar).Value = TIMS.GetValue1(A6.SelectedValue) 'A6
                .Parameters.Add("A7", SqlDbType.VarChar).Value = TIMS.GetValue1(A7.SelectedValue) 'A7

                .Parameters.Add("B11", SqlDbType.VarChar).Value = TIMS.GetValue1(B11.SelectedValue) 'B11
                .Parameters.Add("B12", SqlDbType.VarChar).Value = TIMS.GetValue1(B12.SelectedValue) 'B12
                .Parameters.Add("B13", SqlDbType.VarChar).Value = TIMS.GetValue1(B13.SelectedValue) 'B13
                .Parameters.Add("B14", SqlDbType.VarChar).Value = TIMS.GetValue1(B14.SelectedValue) 'B14
                .Parameters.Add("B15", SqlDbType.VarChar).Value = TIMS.GetValue1(B15.SelectedValue) 'B15

                .Parameters.Add("B21", SqlDbType.VarChar).Value = TIMS.GetValue1(B21.SelectedValue) 'B21
                .Parameters.Add("B22", SqlDbType.VarChar).Value = TIMS.GetValue1(B22.SelectedValue) 'B22
                .Parameters.Add("B23", SqlDbType.VarChar).Value = TIMS.GetValue1(B23.SelectedValue) 'B23

                .Parameters.Add("B31", SqlDbType.VarChar).Value = TIMS.GetValue1(B31.SelectedValue) 'B31
                .Parameters.Add("B32", SqlDbType.VarChar).Value = TIMS.GetValue1(B32.SelectedValue) 'B32
                .Parameters.Add("B41", SqlDbType.VarChar).Value = TIMS.GetValue1(B41.SelectedValue) 'B41
                .Parameters.Add("B42", SqlDbType.VarChar).Value = TIMS.GetValue1(B42.SelectedValue) 'B42
                .Parameters.Add("B43", SqlDbType.VarChar).Value = TIMS.GetValue1(B43.SelectedValue) 'B43
                .Parameters.Add("B44", SqlDbType.VarChar).Value = TIMS.GetValue1(B44.SelectedValue) 'B44

                .Parameters.Add("B51", SqlDbType.VarChar).Value = TIMS.GetValue1(B51.SelectedValue)
                .Parameters.Add("B61", SqlDbType.VarChar).Value = TIMS.GetValue1(B61.SelectedValue)
                .Parameters.Add("B62", SqlDbType.VarChar).Value = TIMS.GetValue1(B62.SelectedValue)
                .Parameters.Add("B63", SqlDbType.VarChar).Value = TIMS.GetValue1(B63.SelectedValue)
                .Parameters.Add("B71", SqlDbType.VarChar).Value = TIMS.GetValue1(B71.SelectedValue)
                .Parameters.Add("B72", SqlDbType.VarChar).Value = TIMS.GetValue1(B72.SelectedValue)
                .Parameters.Add("B73", SqlDbType.VarChar).Value = TIMS.GetValue1(B73.SelectedValue)
                .Parameters.Add("B74", SqlDbType.VarChar).Value = TIMS.GetValue1(B74.SelectedValue)
                .Parameters.Add("C11", SqlDbType.VarChar).Value = TIMS.GetValue1(C11.SelectedValue)

                .Parameters.Add("C21_NOTE", SqlDbType.VarChar).Value = TIMS.GetValue1(C21_NOTE.Text)
                .Parameters.Add("A2_7_NOTE", SqlDbType.VarChar).Value = TIMS.GetValue1(A2_7_NOTE.Text)
                .Parameters.Add("A3_5_NOTE", SqlDbType.VarChar).Value = TIMS.GetValue1(A3_5_NOTE.Text)

                .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sMODIFYACCT 'sm.UserInfo.UserID 'MODIFYACCT
                '.Parameters.Add("MODIFYDATE", SqlDbType.VarChar).Value = MODIFYDATE
                .Parameters.Add("DASOURCE", SqlDbType.VarChar).Value = cst_DASOURCE '單位填寫'資料來源	Null:未知 1:報名網()(學員外網填寫。)2@TIMS系統()
                'dt.Load(.ExecuteReader())
                .Parameters.Add("SOCID", SqlDbType.VarChar).Value = TIMS.GetValue1(v_SOCID) 'SOCID.SelectedValue)
                '.ExecuteNonQuery() 'rst = .ExecuteScalar()
            End With
            DbAccess.ExecuteNonQuery(u_sql, objconn, uCmd.Parameters)
        End If

    End Sub

    ''' <summary>資料來源(DASOURCE) 0:未填寫或未知 1: 報名網(學員外網填寫。) 2: TIMS系統</summary>
    ''' <param name="strSOCID"></param>
    ''' <returns></returns>
    Function Chk_DASOURCE(ByVal strSOCID As String) As Integer
        Dim rst As Integer = 0
        strSOCID = TIMS.ClearSQM(strSOCID)
        If strSOCID = "" Then Return rst

        Call TIMS.OpenDbConn(objconn)
        Dim sql As String = "SELECT * FROM STUD_QUESTIONFAC2 WHERE SOCID =@SOCID"
        Dim sCmd As New SqlCommand(sql, objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("SOCID", SqlDbType.VarChar).Value = strSOCID
            dt.Load(.ExecuteReader())
        End With
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return rst
        Dim dr As DataRow = dt.Rows(0)
        Select Case Convert.ToString(dr("DASOURCE"))
            Case "1", "2" '資料來源(DASOURCE) 0:未填寫或未知 1: 報名網(學員外網填寫。) 2: TIMS系統 
                rst = Val(dr("DASOURCE"))
        End Select
        Return rst
    End Function

End Class