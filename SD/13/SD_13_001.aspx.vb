Partial Class SD_13_001
    Inherits AuthBasePage

    Dim dteStart As DateTime = Now 'dteStart = Now
    Dim s_logmsg1 As String = ""

    Dim ff As String = ""
    Const cst_SAVE_OK_msg1 As String = "「參訓學員出席紀錄一覽表」已連動更新，前已列印者請重新列印；未列印者請忽略此訊息。"

    Const cst_學號 As Integer = 0
    Const cst_姓名 As Integer = 1
    Const cst_身分證號碼 As Integer = 2
    Const cst_是否獲得學分 As Integer = 3
    'Const cst_出席達2分之3 = 4
    'Const cst_出席達3分之4 As Integer = 4
    '缺席未超過1/5
    '3.1/5的判斷依學員出缺勤作業（首頁>>學員動態管理>>教務管理>>學員出缺勤作業）
    '「缺席時數」扣除「喪假」時數未超過1/5（＞＝）者，顯示「是」；時數 超過1/5者，顯示「否」。 
    Const cst_缺席未超過5分之1 As Integer = 4
    Const cst_是否補助 As Integer = 5
    Const cst_補助比例 As Integer = 6
    Const cst_總費用 As Integer = 7
    Const cst_補助費用 As Integer = 8
    Const cst_個人支付 As Integer = 9
    Const cst_剩餘可用餘額 As Integer = 10
    Const cst_其他申請中金額 As Integer = 11
    'Const cst_是否提出申請 As Integer = 12
    Const cst_申請狀態 As Integer = 13
    Const cst_撥款狀態 As Integer = 14
    Const cst_預算別 As Integer = 15

    '年度大於2011啟用。
#Region "Functions"

    ''' <summary>此班級學員資料尚未審核</summary>
    ''' <param name="tmpOCID"></param>
    ''' <returns></returns>
    Function Check_AppliedResultR(ByVal tmpOCID As String) As Boolean
        Dim rst As Boolean = False
        Dim sqlStr As String = "SELECT ISNULL(AppliedResultR,'N') AppliedResultR FROM CLASS_CLASSINFO WHERE OCID=@OCID"
        Dim sCmd As New SqlCommand(sqlStr, objconn)
        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("OCID", SqlDbType.VarChar).Value = tmpOCID
            dt.Load(.ExecuteReader())
        End With
        If dt.Rows.Count > 0 Then
            If dt.Rows(0)("AppliedResultR") = "Y" Then rst = True
        End If
        Return rst
    End Function

    '完成結訓動作
    'Private Function Check_IsClosed(ByVal t_OCID As Integer) As Boolean
    '    Dim rst As Boolean = False
    '    Call TIMS.OpenDbConn(objconn)
    '    Dim sqlStr As String = "SELECT ISNULL(IsClosed,'N') IsClosed FROM CLASS_CLASSINFO WHERE OCID=@OCID"
    '    Dim dt As New DataTable
    '    Dim s_cmd As New SqlCommand(sqlStr, objconn)
    '    With s_cmd
    '        .Parameters.Clear()
    '        .Parameters.Add("OCID", SqlDbType.Int).Value = t_OCID
    '        dt.Load(.ExecuteReader())
    '    End With

    '    If dt IsNot Nothing AndAlso dt.Rows.Count > 0 AndAlso Convert.ToString(dt.Rows(0)("IsClosed")) = "Y" Then rst = True

    '    Return rst
    'End Function

#End Region

    Dim gflagTest1 As Boolean = False 'TIMS.sUtl_ChkTest() '測試環境參數
    Dim gsBlackIDNO As String = "" '學員處分
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
        gflagTest1 = TIMS.sUtl_ChkTest() '測試環境

        'Dim work2016 As String = TIMS.Utl_GetConfigSet("work2016")
        'If work2016 <> "Y" Then
        '    Select Case sm.UserInfo.Years
        '        Case Is <= "2011"
        '            Server.Transfer("SD_13_001_00.aspx?ID=" & Request("ID"))
        '            Exit Sub
        '        Case Is <= "2015"
        '            Server.Transfer("SD_13_001_15.aspx?ID=" & Request("ID"))
        '            Exit Sub
        '    End Select
        'End If

        If Not IsPostBack Then
            CCREATE1()
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');SetOneOCID();"
        Button2.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

        Button1.Attributes("onclick") = "return CheckSearch();"
        Button3.Attributes("onclick") = "return CheckData();"
    End Sub

    Sub CCREATE1()
        msg.Text = ""
        DataGridTable.Style("display") = "none"
        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID
        Dim V_INQUIRY As String = Session(String.Concat(TIMS.cst_GSE_V_INQUIRY, TIMS.Get_MRqID(Me)))
        If TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES) Then Call TIMS.GET_INQUIRY(ddl_INQUIRY_Sch, objconn, V_INQUIRY)
        '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub

    Private Function GET_SEARCH_MEMO() As String
        Dim RstMemo As String = ""
        center.Text = TIMS.ClearSQM(center.Text)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If center.Text <> "" Then RstMemo &= String.Concat("&center=", center.Text)
        If RIDValue.Value <> "" Then RstMemo &= String.Concat("&RID=", RIDValue.Value)
        Return RstMemo
    End Function

    '查詢
    Sub sUtl_Search1()
        dteStart = Now
        '20090907 針對尚未完成班級結訓動作的班級，是不能查出資料的。
        msg.Text = ""
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If Not gflagTest1 Then '(非)測試環境參數
            '此班級學員資料尚未審核
            If Not Check_AppliedResultR(OCIDValue1.Value) Then
                DataGridTable.Style("display") = "none"
                msg.Text = "查無資料"
                Common.MessageBox(Me, "該班學員資料複審結果尚未通過")
                Exit Sub
            End If
        End If
        s_logmsg1 = "##SD_13_001: Check_AppliedResultR(OCIDValue1.Value)," & TIMS.GetTsEnd(dteStart)
        TIMS.LOG.Debug(s_logmsg1)

        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "請選擇有效班級資料!!")
            Exit Sub
        End If
        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        If (TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES)) Then
            If (v_INQUIRY = "") Then Common.MessageBox(Me, "請選擇「查詢原因」") : Return
            Session(String.Concat(TIMS.cst_GSE_V_INQUIRY, TIMS.Get_MRqID(Me))) = v_INQUIRY
        End If

        Dim cPlanid As String = CStr(drCC("planid"))
        Dim cComIDNO As String = CStr(drCC("comidno"))
        Dim cSeqNo As String = CStr(drCC("seqno"))
        '完成結訓動作
        If Not TIMS.Check_IsClosed(Val(OCIDValue1.Value), objconn) Then
            DataGridTable.Style("display") = "none"
            msg.Text = "(本班尚未完成結訓動作)查無資料"
            Exit Sub
        ElseIf TIMS.CHK_STUDENTSOFCLASS_S1(Val(OCIDValue1.Value), objconn) Then
            DataGridTable.Style("display") = "none"
            msg.Text = "(本班尚未完成結訓動作)有學員參訓狀態仍為在訓中"
            Exit Sub
        End If

        '職前參訓歷史查詢-WEB-SERVICE-依OCID
        TIMS.GetTrainingList2OCID(sm, objconn, OCIDValue1.Value)

        hidBlackMsg.Value = "" '清空黑名單暫存記錄(2009/07/28 判斷黑名單)

        Dim sqlstr As String = ""
        sqlstr &= " SELECT a.OCID,d.SETID,d.SOCID" & vbCrLf
        sqlstr &= " ,d.IdentityID" & vbCrLf
        sqlstr &= " ,dbo.FN_CSTUDID2(d.StudentID) StudentID" & vbCrLf
        sqlstr &= " ,ISNULL(g.BudID,d.BudgetID) BudgetID" & vbCrLf
        sqlstr &= " ,d.SupplyID ESupplyID" & vbCrLf
        '判斷是否填寫問卷 (STUD_QUESTIONFAC/STUD_QUESTIONFAC2) 'dbo.FN_GET_GOVCNT(意見調查表) 
        sqlstr &= " ,dbo.FN_GET_GOVCNT(d.SOCID) GovCnt" & vbCrLf
        sqlstr &= " ,e.Name ,e.IDNO" & vbCrLf
        sqlstr &= " ,d.CreditPoints" & vbCrLf
        '除數可能有溢位問題，無條件捨去餘2位數。
        sqlstr &= " ,FLOOR(ISNULL(b.TotalCost,0)/ISNULL(b.TNum,1)) Total" & vbCrLf
        sqlstr &= " ,b.TotalCost TotalCostX" & vbCrLf
        sqlstr &= " ,b.TNum ,a.THours ,ar.DistID ,ISNULL(f.COUNTHOURS,0) COUNTHOURS" & vbCrLf
        'sqlstr += " ,ISNULL(f.COUNTHOURS2,0) COUNTHOURS2" & vbCrLf ' 扣除「喪假」時數
        sqlstr &= " ,e.DegreeID ,d.StudStatus ,d.MIdentityID ,a.STDate" & vbCrLf
        sqlstr &= " ,a.AppliedResultM ,d.AppliedResult" & vbCrLf
        sqlstr &= " ,g.SOCID Exist" & vbCrLf
        sqlstr &= " ,g.SumOfMoney ,g.PayMoney ,g.AppliedStatus" & vbCrLf
        sqlstr &= " ,g.AppliedNote" & vbCrLf
        sqlstr &= " ,g.SupplyID" & vbCrLf
        '其他申請中金額
        'sqlstr &= " ,dbo.FN_GET_GOVAPPL2(e.IDNO,a.STDate) GovAppl2" & vbCrLf
        sqlstr &= "  ,dbo.FN_GET_GOVCOST2(e.IDNO, convert(varchar,a.STDate,111)) GovAppl2"
        '其他申請中金額(並排除本班)
        'sqlstr += " ,dbo.FN_GET_GOVAPPL22(e.IDNO,a.STDate,a.OCID) GovAppl2" & vbCrLf
        sqlstr &= " ,g.AppliedStatusM" & vbCrLf
        sqlstr &= " FROM Class_ClassInfo a WITH(NOLOCK)" & vbCrLf
        sqlstr &= " JOIN Plan_PlanInfo b WITH(NOLOCK) ON a.PlanID = b.PlanID AND a.ComIDNO = b.ComIDNO AND a.SeqNo = b.SeqNo" & vbCrLf
        sqlstr &= " JOIN Auth_Relship ar WITH(NOLOCK) ON a.RID = ar.RID" & vbCrLf
        sqlstr &= " JOIN Class_StudentsOfClass d WITH(NOLOCK) ON a.OCID = d.OCID" & vbCrLf
        sqlstr &= " JOIN Stud_StudentInfo e WITH(NOLOCK) ON d.SID = e.SID" & vbCrLf
        '喪假(LEAVEID:05)。99:(使用者輸入)
        sqlstr &= " LEFT JOIN (" & vbCrLf
        sqlstr &= "   SELECT t.SOCID ,SUM(CASE WHEN t.LEAVEID IS NULL THEN t.Hours END) COUNTHOURS" & vbCrLf
        sqlstr &= "   FROM STUD_TURNOUT2 t WITH(NOLOCK)" & vbCrLf
        sqlstr &= "   JOIN CLASS_STUDENTSOFCLASS cs WITH(NOLOCK) on cs.socid = t.socid" & vbCrLf
        sqlstr &= "   WHERE cs.OCID = '" & OCIDValue1.Value & "'" & vbCrLf
        sqlstr &= "   GROUP BY t.SOCID) f ON f.SOCID = d.SOCID" & vbCrLf
        sqlstr &= " LEFT JOIN STUD_SUBSIDYCOST g WITH(NOLOCK) ON d.SOCID=g.SOCID" & vbCrLf
        sqlstr &= " WHERE a.OCID = '" & OCIDValue1.Value & "'" & vbCrLf
        If Not gflagTest1 Then '(非)測試環境參數
            sqlstr &= " AND a.AppliedResultR = 'Y'" & vbCrLf
        End If
        '產業人才投資方案 Y:通過 C:全班學員資料確認
        ' ORDER BY StudentID ASC"
        If InStr(Me.ViewState("sort"), "IDNO") > 0 Then
            sqlstr &= " ORDER BY e." & Me.ViewState("sort").ToString
        ElseIf InStr(Me.ViewState("sort"), "StudentID") > 0 Then
            sqlstr &= " ORDER BY dbo.FN_CSTUDID2(d.StudentID) " & Replace(Me.ViewState("sort").ToString, "StudentID", "") & vbCrLf
        Else
            sqlstr &= " ORDER BY dbo.FN_CSTUDID2(d.StudentID)" & vbCrLf
        End If

        dteStart = Now
        Dim dt As DataTable = Nothing
        Try
            'SQL
            dt = DbAccess.GetDataTable(sqlstr, objconn)
        Catch ex As Exception
            Dim strErrmsg As String = ""
            strErrmsg += "/*  ex.ToString: */" & vbCrLf
            strErrmsg += ex.ToString & vbCrLf
            strErrmsg += "/* sqlstr: */" & vbCrLf
            strErrmsg += sqlstr & vbCrLf
            'strErrmsg += TIMS.GetErrorMsg(MyPage) '取得錯誤資訊寫入
            'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)
            Common.MessageBox(Me, "資料庫效能異常，請重新查詢")
            Exit Sub
        End Try
        s_logmsg1 = "##SD_13_001: DbAccess.GetDataTable(sqlstr, objconn)," & TIMS.GetTsEnd(dteStart)
        TIMS.LOG.Debug(s_logmsg1)

        'START 黑名單之身分證記錄 2009/07/28 by waiming
        Dim stdBLACK2TPLANID As String = ""
        Dim iStdBlackType As Integer = TIMS.Chk_StdBlackType(Me, objconn, stdBLACK2TPLANID)
        gsBlackIDNO = TIMS.Get_StdBlackIDNO(Me, iStdBlackType, stdBLACK2TPLANID, objconn) '學員處分

        Dim sMemo As String = GET_SEARCH_MEMO()
        '--查詢原因:INQNO--查詢結果筆數:RESCNT--查詢清單結果:RESDESC 
        Dim vRESDESC As String = TIMS.GET_RESDESCdt(dt, "NAME,IDNO")
        Call TIMS.SubInsAccountLog1(Me, TIMS.Get_MRqID(Me), TIMS.cst_wm查詢, TIMS.cst_wmdip2, OCIDValue1.Value, sMemo, objconn, v_INQUIRY, dt.Rows.Count, vRESDESC)

        DataGridTable.Style("display") = "none"
        msg.Text = "查無資料"
        If dt.Rows.Count > 0 Then
            dteStart = Now
            '顯示有效資料
            DataGridTable.Style("display") = "inline"
            msg.Text = ""
            If ViewState("sort") = "" Then ViewState("sort") = "StudentID"
            DataGrid1.DataKeyField = "SOCID"
            DataGrid1.DataSource = dt
            DataGrid1.DataBind()
            '顯示有效資料
            s_logmsg1 = "##SD_13_001: DataGrid1.DataSource = dt," & TIMS.GetTsEnd(dteStart)
            TIMS.LOG.Debug(s_logmsg1)

            '檢查是否有重複參訓學員排除產學訓計畫
            Dim dr As DataRow = dt.Rows(0)
            Button3.Enabled = True '儲存鈕
            If Convert.ToString(dr("AppliedResultM")) = "Y" Then
                'Button3.Enabled = False '永遠可申請
                TIMS.Tooltip(Button3, "班級學員經費審核結果，已完成")
            End If

            dteStart = Now
            'If Not gflagTest1 Then '(非)測試環境參數
            dt = TIMS.GET_Duplicate_Student(OCIDValue1.Value, 1, objconn) '檢查是否有重複參訓學員排除產學訓計畫
            Dclass.Value = If(dt IsNot Nothing, 1, 2) '1 '有重複參訓學員 2'沒重複
            s_logmsg1 = "##SD_13_001: TIMS.GET_Duplicate_Student(OCIDValue1.Value, 1, objconn)," & TIMS.GetTsEnd(dteStart)
            TIMS.LOG.Debug(s_logmsg1)
        End If
    End Sub

    '查詢
    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call sUtl_Search1() '查詢
    End Sub

    Private Sub DataGrid1_SortCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridSortCommandEventArgs) Handles DataGrid1.SortCommand
        Dim s_sort1 As String = Convert.ToString(Me.ViewState("sort"))
        Me.ViewState("sort") = If(Me.ViewState("sort") <> e.SortExpression, e.SortExpression, e.SortExpression & " DESC")
        Dim s_sort2 As String = Convert.ToString(Me.ViewState("sort"))
        If s_sort1 <> s_sort2 Then
            'Button1_Click(Me, e)
            Call sUtl_Search1()
        End If
    End Sub

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Const cst_總費用_msg As String = "非學分班的訓練費用項目"
        Const cst_是否補助_msg As String = "是否有請領補助津貼的資格"
        Const cst_補助費用_msg As String = "預定要補助的金額(可自行變動，未申請前系統會根據可用餘額推算)"
        Const cst_個人支付_msg As String = "學員自行要支付的金額(會根據補助費用所輸入的值來調動)"
        'Const cst_剩餘可用餘額_msg = "學員目前可用餘額-這次預定補助費用的剩餘金額(成為負數時會以紅字表示)"
        Const cst_剩餘可用餘額_msg As String = "學員目前可用餘額-已審核通過費用的剩餘金額(成為負數時會以紅字表示)"
        Const cst_目前申請總額_msg As String = "學員目前已申請未審核補助金總額(含本次補助費用)合併後 超過剩餘可用餘額以紅字表示" '其他申請中金額
        'Dim totalleft As Decimal = 0   ' 剩餘可用餘額   
        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = "head_navy"
                'e.Item.Cells(5).ToolTip = "是否有請領補助津貼的資格"
                'e.Item.Cells(7).ToolTip = "預定要補助的金額(可自行變動，未申請前系統會根據可用餘額推算)"
                'e.Item.Cells(8).ToolTip = "學員自行要支付的金額(會根據補助費用所輸入的值來調動)"
                'e.Item.Cells(9).ToolTip = "學員目前可用餘額-這次預定補助費用的剩餘金額(成為負數時會以紅字表示)"
                e.Item.Cells(cst_是否補助).ToolTip = cst_是否補助_msg
                e.Item.Cells(cst_總費用).ToolTip = cst_總費用_msg
                e.Item.Cells(cst_補助費用).ToolTip = cst_補助費用_msg
                e.Item.Cells(cst_個人支付).ToolTip = cst_個人支付_msg
                e.Item.Cells(cst_剩餘可用餘額).ToolTip = cst_剩餘可用餘額_msg
                e.Item.Cells(cst_其他申請中金額).ToolTip = cst_目前申請總額_msg '其他申請中的金額
                If Me.ViewState("sort") <> "" Then
                    'Dim mylabel As String
                    Dim mysort As New System.Web.UI.WebControls.Image
                    Dim i As Integer = -1
                    Select Case Me.ViewState("sort")
                        Case "StudentID", "StudentID DESC"
                            'mylabel = "IDNO"
                            i = 0
                            mysort.ImageUrl = If(Me.ViewState("sort") = "StudentID", "../../images/SortUp.gif", "../../images/SortDown.gif")
                        Case "IDNO", "IDNO DESC"
                            'mylabel = "StudentID"
                            i = 2
                            mysort.ImageUrl = If(Me.ViewState("sort") = "IDNO", "../../images/SortUp.gif", "../../images/SortDown.gif")
                    End Select
                    If i <> -1 Then e.Item.Cells(i).Controls.Add(mysort)
                End If

            Case ListItemType.Item, ListItemType.AlternatingItem
                If e.Item.ItemType = ListItemType.Item Then e.Item.CssClass = ""

                Dim Flag As Integer = 0  '得到學分  
                Dim FlagStudy As Integer = 0 '出席達到資格 (3分之4)
                Dim iFlag3 As Integer = 0  '填寫調查表 0:未填寫調查表 1:有填寫。
                Dim drv As DataRowView = e.Item.DataItem
                Dim SupplyID As DropDownList = e.Item.FindControl("SupplyID") 'ESupplyID 'SupplyID.Enabled
                Dim BudID As DropDownList = e.Item.FindControl("BudID") 'BudgetID 'BudID.Enabled
                'Dim HidSupplyID As HiddenField = e.Item.FindControl("HidSupplyID") 'ESupplyID 'SupplyID.Enabled
                'Dim HidBudID As HiddenField = e.Item.FindControl("HidBudID") 'BudgetID 'BudID.Enabled
                'Dim labSupplyID As Label = e.Item.FindControl("labSupplyID") 'ESupplyID 'SupplyID.Enabled
                'Dim labBudID As Label = e.Item.FindControl("labBudID") 'BudgetID 'BudID.Enabled

                '補助比例和預算別改唯讀
                SupplyID.Enabled = False '補助比例
                BudID.Enabled = False '預算別'暫設不可更改 預算別，根據審核狀況來開放 預算別
                'SupplyID.Style("display") = "none" '隱藏不顯示
                'BudID.Style("display") = "none" '隱藏不顯示

                Dim DataGrid2 As DataGrid = e.Item.FindControl("DataGrid2")
                Dim CreditPoints As Label = e.Item.FindControl("CreditPoints")
                Dim SumOfMoney As TextBox = e.Item.FindControl("SumOfMoney")
                Dim Checkbox1 As HtmlInputCheckBox = e.Item.FindControl("Checkbox1")
                'Attribute SelectItemClear(headchkbox) SelectItemAll('headCheckbox1');
                Checkbox1.Attributes("onclick") = "SelectItemClear('headCheckbox1');"

                Dim RemainSub As HtmlInputHidden = e.Item.FindControl("RemainSub")
                Dim MaxSub As HtmlInputHidden = e.Item.FindControl("MaxSub")
                Dim PayMoney As HtmlInputHidden = e.Item.FindControl("PayMoney")
                Dim balancemoney As HtmlInputHidden = e.Item.FindControl("balancemoney")

                '20090201 andy edit
                '-------------
                Dim setid As HtmlInputHidden = e.Item.FindControl("setid")
                Dim ocid As HtmlInputHidden = e.Item.FindControl("ocid")
                Dim socid As HtmlInputHidden = e.Item.FindControl("socid")
                setid.Value = drv("setid").ToString
                ocid.Value = drv("ocid").ToString
                socid.Value = drv("socid").ToString
                '-------------

                Dim star1 As TextBox = e.Item.FindControl("star1") '未填寫 意見調查表
                Dim stud_1 As TextBox = e.Item.FindControl("stud1") '學號
                star1.Visible = True
                stud_1.Visible = True '學號
                SupplyID = TIMS.Get_SupplyID(SupplyID)

                e.Item.Cells(cst_是否補助).ToolTip = cst_是否補助_msg
                e.Item.Cells(cst_總費用).ToolTip = cst_總費用_msg
                e.Item.Cells(cst_補助費用).ToolTip = cst_補助費用_msg
                e.Item.Cells(cst_個人支付).ToolTip = cst_個人支付_msg
                e.Item.Cells(cst_剩餘可用餘額).ToolTip = cst_剩餘可用餘額_msg
                e.Item.Cells(cst_其他申請中金額).ToolTip = cst_目前申請總額_msg

                If Convert.ToString(drv("IDNO")) <> "" Then
                    e.Item.Cells(cst_姓名).ToolTip = TIMS.Search_Stud_SubsidyCost(Convert.ToString(drv("IDNO")), objconn)
                    e.Item.Cells(cst_學號).ToolTip = e.Item.Cells(cst_姓名).ToolTip
                    e.Item.Cells(cst_身分證號碼).ToolTip = e.Item.Cells(cst_姓名).ToolTip
                End If
                Common.SetListItem(SupplyID, Convert.ToString(drv("ESupplyID")))
                'If drv("ESupplyID").ToString <> "" Then
                '    'labSupplyID.Text = SupplyID.SelectedItem.Text
                '    'HidSupplyID.Value = Convert.ToString(drv("ESupplyID"))
                'End If

                '判斷是否填寫問卷 star1  (意見調查表)
                star1.Text = "" '已填寫(意見調查表)
                iFlag3 = 1 '已填寫(意見調查表)

                '填寫(意見調查表)
                Dim flag_GovCnt_0 As Boolean = False '已填寫
                If Not flag_GovCnt_0 AndAlso IsDBNull(drv("GovCnt")) Then flag_GovCnt_0 = True '未填寫(意見調查表)
                If Not flag_GovCnt_0 AndAlso Convert.ToString(drv("GovCnt")) = "" Then flag_GovCnt_0 = True '未填寫(意見調查表)
                If Not flag_GovCnt_0 AndAlso Convert.ToString(drv("GovCnt")) = "0" Then flag_GovCnt_0 = True '未填寫(意見調查表)
                '未填寫(意見調查表)
                If flag_GovCnt_0 Then star1.Text = "*" '未填寫(意見調查表)
                If flag_GovCnt_0 Then iFlag3 = 0 '未填寫(意見調查表)
                If flag_GovCnt_0 Then TIMS.Tooltip(star1, "未填寫調查表。") 'GovCnt

                stud_1.Text = drv("StudentID").ToString '學號
                BudID = TIMS.Get_Budget(BudID, 2, objconn)
                '沒有公務可用。
                If BudID.Items.FindByValue("01") IsNot Nothing Then
                    If drv("DistID").ToString <> "001" Then BudID.Items.Remove(BudID.Items.FindByValue("01"))
                End If

                'If drv("SupplyID").ToString <> "" Then Common.SetListItem(SupplyID, drv("SupplyID").ToString)

                '規則改為show class_studentsofclass.budgetid
                If drv("BudgetID").ToString <> "" Then
                    Common.SetListItem(BudID, Convert.ToString(drv("BudgetID")))
                    'labBudID.Text = BudID.SelectedItem.Text
                    'HidBudID.Value = Convert.ToString(drv("BudgetID")) '.Text = BudID.SelectedItem.Text
                End If

                'If drv("BudID").ToString <> "" Then
                '    Common.SetListItem(BudID, drv("BudID").ToString)
                'Else
                '    If drv("BudgetID").ToString <> "" Then Common.SetListItem(BudID, drv("BudgetID").ToString)
                'End If

                '20090123 andy  edit 產投、在職 2009年 身分別為「就業保險被保險人非自願失業者」時
                '1.預算來源設定為 02:就安基金 ； 2.補助比例為100%  
                '090423(將090123修改之程式mark起來)直接帶前端所輸入之原始值即可，不用再做一次判斷重新給值
                '--------------------------  start
                'If sm.UserInfo.TPlanID = 28 Then
                '    If CInt(Me.sm.UserInfo.Years) > 2008 Then
                '        For i As Integer = 0 To Split(Convert.ToString(drv("IdentityID")), ",").Length - 1
                '            If Split(drv("IdentityID").ToString, ",")(i) = "02" Then
                '                Common.SetListItem(BudID, "02")
                '                Common.SetListItem(SupplyID, "2")
                '                drv("ESupplyID") = "2"
                '            End If
                '        Next
                '    End If
                'End If
                '----------------------------- end
                '是否得到學分
                CreditPoints.Text = If(Convert.ToString(drv("CreditPoints")) = "1", "是", "<font color='RED'>否</font>")
                If Convert.ToString(drv("CreditPoints")) = "1" Then Flag = 1

                '缺席時數
                e.Item.Cells(cst_缺席未超過5分之1).Text = "否"
                If Convert.ToString(drv("THours")) <> "" Then
                    If drv("THours") > 0 Then
                        Dim iVal1 As Double = Val(drv("COUNTHOURS")) '- Val(drv("COUNTHOURS2"))
                        TIMS.Tooltip(e.Item.Cells(cst_缺席未超過5分之1), "出席時數:" & (drv("THours") - iVal1) & "/" & drv("THours"))
                        If iVal1 / drv("THours") <= 1 / 5 Then
                            e.Item.Cells(cst_缺席未超過5分之1).Text = "是"
                            FlagStudy = 1
                        End If
                    End If
                End If

                'e.Item.Cells(cst_出席達3分之4).Text = "否"
                'If drv("THours") > 0 Then
                '    If (drv("THours") - drv("COUNTHOURS")) / drv("THours") >= 3 / 4 Then
                '        e.Item.Cells(cst_出席達3分之4).Text = "是"
                '        FlagStudy = 1
                '    End If
                'End If
                'If drv("THours") > 0 Then
                '    If (drv("THours") - drv("COUNTHOURS")) / drv("THours") >= 2 / 3 Then
                '        e.Item.Cells(cst_出席達2分之3).Text = "是"
                '        FlagStudy = 1
                '    Else
                '        e.Item.Cells(cst_出席達2分之3).Text = "否"
                '    End If
                'Else
                '    e.Item.Cells(cst_出席達2分之3).Text = "否"
                'End If

                'If Int(drv("DegreeID")) <= 2 Then
                'End If
                '可用補助額
                Dim Total As Integer = TIMS.Get_3Y_SupplyMoney()
                '20080609  Andy  可用補助額
                '含職前webservice
                Dim SubsidyCost As Double = TIMS.Get_SubsidyCost(drv("IDNO").ToString(), drv("STDate").ToString(), "", "Y", objconn)
                Total -= SubsidyCost

                If Total < 0 Then Total = 0
                RemainSub.Value = Total '30000 '可用補助額

                'Dim ESupplyPercent, SupplyPercent As Double
                If IsDBNull(drv("Exist")) Then      '表示沒資料,以新增的型態顯示
                    ' If Flag = 1 Then '得到學分  
                    If (Flag = 1 AndAlso FlagStudy = 1) Then   '970513 Andy  得到學分且出席達2/3  
                        e.Item.Cells(cst_是否補助).Text = "是"
                        If drv("MIdentityID").ToString <> "" Then
                            '20080806 andy 原程式補助費用判斷有納入主要身分為一般身分別時只能補助80%的條件==>改為只依據產學訓(補助比例代碼)來做判斷
                            '-----------------------------------------------------
                            'If drv("MIdentityID").ToString = "01" Then '一般身分者
                            '    'drv("Total") 此次費用
                            '    If Total >= Decimal.Truncate(drv("Total") * 0.8) Then '可用補助額 > '計算補助費用
                            '        SumOfMoney.Text = Decimal.Truncate(drv("Total") * 0.8) '此次可用補助額=(課程費用*0.8)計算補助費用 
                            '    Else
                            '        SumOfMoney.Text = Total '此次可用補助額=可用補助額 
                            '    End If
                            'Else '其他身分者
                            '    If Total >= drv("Total") Then '可用補助額 > '課程費用*1(計算補助費用)
                            '        SumOfMoney.Text = drv("Total") '此次可用補助額=課程費用 
                            '    Else
                            '        SumOfMoney.Text = Total '此次可用補助額=可用補助額 
                            '    End If
                            'End If
                            '--------   Start
                            'Select Case drv("MIdentityID")  '主要身分代碼
                            '    Case "01"                  '一般身分者,補助比例80%    
                            '        SupplyPercent = 0.8
                            '    Case Else                   '其它身分,補助比例100%
                            '        SupplyPercent = 1
                            'End Select
                            '(有其他狀況)暫'補助比例0%
                            Dim ESupplyPercent As Double = 0
                            If drv("ESupplyID").ToString <> "" Then
                                Select Case drv("ESupplyID")  '產學訓(補助比例代碼)
                                    Case 1  '補助比例80%
                                        ESupplyPercent = 0.8
                                    Case 2  '補助比例100%
                                        ESupplyPercent = 1
                                    Case 9  '補助比例0%
                                        ESupplyPercent = 0
                                    Case Else '(有其他狀況)暫'補助比例0%
                                        ESupplyPercent = 0
                                End Select
                            End If
                            If Total >= Decimal.Truncate(drv("Total") * ESupplyPercent) Then
                                SumOfMoney.Text = Decimal.Truncate(drv("Total") * ESupplyPercent)
                            Else
                                SumOfMoney.Text = Total '此次可用補助額=可用補助額 
                            End If
                            '------   End
                            MaxSub.Value = SumOfMoney.Text '此次最大可用補助額
                            e.Item.Cells(cst_個人支付).Text = CStr(CInt(drv("Total")) - CInt(IIf(Trim(SumOfMoney.Text) = "", "0", Trim(SumOfMoney.Text))))   '課程費用-'可用補助額=個人支付費用
                            PayMoney.Value = CInt(drv("Total")) - CInt(SumOfMoney.Text) '課程費用-'可用補助額=個人支付費用
                            e.Item.Cells(cst_剩餘可用餘額).Text = Total.ToString()  '可用補助額=剩餘可用餘額
                            'totalleft = Total  '20091221 andy 
                        Else
                            SumOfMoney.Enabled = False '不可填入補助費用
                            Checkbox1.Disabled = True '不可提出申請
                            'e.Item.Cells(cst_是否補助).Text = "<font color='RED'>否</font>"
                            e.Item.Cells(cst_剩餘可用餘額).Text = Total
                            'totalleft = Total  '20091221 andy 
                        End If
                    Else
                        SumOfMoney.Enabled = False '不可填入補助費用
                        Checkbox1.Disabled = True  '不可提出申請
                        Checkbox1.Checked = False
                        e.Item.Cells(cst_是否補助).Text = "<font color='RED'>否</font>"
                        e.Item.Cells(cst_剩餘可用餘額).Text = Total
                        '20080606 andy 是否補助為「否」,個人支付=總費用
                        e.Item.Cells(cst_個人支付).Text = e.Item.Cells(cst_總費用).Text
                        SumOfMoney.Text = "0"
                        'totalleft = Total  '20091221 andy 
                    End If
                Else
                    ' If Flag = 1 Then '得到學分
                    If (Flag = 1 AndAlso FlagStudy = 1) Then '970513 Andy  得到學分且出席達2分之3  
                        e.Item.Cells(cst_是否補助).Text = "是"
                    Else '尚未得到學分
                        e.Item.Cells(cst_是否補助).Text = "<font color='RED'>否</font>"
                        '20080606 andy 是否補助為「否」,個人支付=總費用
                        e.Item.Cells(cst_個人支付).Text = e.Item.Cells(cst_總費用).Text
                        SumOfMoney.Text = "0"
                    End If
                    'If drv("MIdentityID").ToString = "01" Then  '身分別,01:一般身分者 
                    '    If Total >= Decimal.Truncate(drv("Total") * 0.8) Then
                    '        MaxSub.Value = Decimal.Truncate(drv("Total") * 0.8) '此次最大可用補助額
                    '    Else
                    '        MaxSub.Value = Total '此次最大可用補助額
                    '    End If
                    'Else
                    '    If Total >= drv("Total") Then
                    '        MaxSub.Value = drv("Total")
                    '    Else
                    '        MaxSub.Value = Total
                    '    End If
                    'End If
                    '20080806  Andy 原程式補助費用判斷有納入主要身分為一般身分別時只能補助80%的條件==>改為只依據產學訓(補助比例代碼)來做判斷
                    '-----------   Start
                    '(有其他狀況)暫'補助比例0%
                    Dim ESupplyPercent As Double = 0
                    Select Case drv("ESupplyID")  '產學訓(補助比例代碼)
                        Case 1  '補助比例80%
                            ESupplyPercent = 0.8
                        Case 2  '比例100%
                            ESupplyPercent = 1
                        Case 9  '比例0%
                            ESupplyPercent = 0
                    End Select
                    If Total >= Decimal.Truncate(drv("Total") * ESupplyPercent) Then
                        MaxSub.Value = Decimal.Truncate(drv("Total") * ESupplyPercent)
                    Else
                        MaxSub.Value = Total '此次可用補助額=可用補助額 
                    End If
                    '-----------   End
                    SumOfMoney.Text = drv("SumOfMoney").ToString '可用補助額
                    PayMoney.Value = drv("PayMoney").ToString '個人支付費用
                    e.Item.Cells(cst_個人支付).Text = drv("PayMoney").ToString '個人支付費用

                    Checkbox1.Checked = True '有提出申請
                    'BudID.Enabled = False '暫設不可更改 預算別，根據審核狀況來開放 預算別
                    If IsDBNull(drv("AppliedStatusM")) Then
                        Checkbox1.Disabled = False '可更改提出申請
                        e.Item.Cells(cst_申請狀態).Text = "審核中"
                        e.Item.Cells(cst_撥款狀態).Text = "未撥款"
                        'BudID.Enabled = True '可更改 預算別
                    Else
                        'If drv("AppliedStatusM") Then
                        '    Checkbox1.Disabled = True
                        '    SumOfMoney.ReadOnly = True
                        '    e.Item.Cells(cst_申請狀態).Text = "審核通過"
                        'Else
                        '    Checkbox1.Disabled = False
                        '    e.Item.Cells(cst_申請狀態).Text = "審核失敗"
                        'End If
                        Select Case drv("AppliedStatusM").ToString
                            Case "Y"
                                Checkbox1.Disabled = True
                                SumOfMoney.ReadOnly = True
                                e.Item.Cells(cst_申請狀態).Text = "審核通過"
                                If IsDBNull(drv("AppliedStatus")) Then '撥款審核狀態
                                    'Checkbox1.Disabled = True 'False '審核通過後不可放棄申請
                                    e.Item.Cells(cst_撥款狀態).Text = "待撥款" '"撥款審核中"
                                Else
                                    If drv("AppliedStatus") Then '=1
                                        'Checkbox1.Disabled = True '審核通過後不可放棄申請
                                        'SumOfMoney.ReadOnly = True
                                        e.Item.Cells(cst_撥款狀態).Text = "已撥款"
                                    Else
                                        'Checkbox1.Disabled = False
                                        e.Item.Cells(cst_撥款狀態).Text = "不撥款"
                                    End If
                                End If
                            Case "N"
                                Checkbox1.Disabled = False '審核失敗  提出申請
                                e.Item.Cells(cst_申請狀態).Text = "審核不通過" '"審核失敗"
                                e.Item.Cells(cst_撥款狀態).Text = "不撥款"
                            Case "R"
                                Checkbox1.Disabled = False '退件修正  提出申請
                                e.Item.Cells(cst_申請狀態).Text = "退件修正"
                                e.Item.Cells(cst_撥款狀態).Text = "未撥款"
                                'BudID.Enabled = True '可更改 預算別
                            Case ""
                                Checkbox1.Disabled = False
                        End Select
                    End If

                    If Not e.Item.Cells(cst_申請狀態).Text = "審核通過" Then
                        If Total - CInt(SumOfMoney.Text) >= 0 Then
                            e.Item.Cells(cst_剩餘可用餘額).Text = Total
                        Else
                            e.Item.Cells(cst_剩餘可用餘額).Text = String.Concat("<font color=Red>", Total, "</font>")
                        End If
                        If drv("GovAppl2") > Total - CInt(SumOfMoney.Text) Then
                            e.Item.Cells(cst_其他申請中金額).Text = String.Concat("<font color=Red>", drv("GovAppl2"), "</font>")
                        End If
                    Else
                        SumOfMoney.Enabled = False
                        If Total >= 0 Then
                            e.Item.Cells(cst_剩餘可用餘額).Text = Total
                        Else
                            e.Item.Cells(cst_剩餘可用餘額).Text = String.Concat("<font color=Red>", Total, "</font>")
                        End If
                        If drv("GovAppl2") > Total Then e.Item.Cells(cst_其他申請中金額).Text = String.Concat("<font color=Red>", drv("GovAppl2"), "</font>")
                    End If
                End If

                'SupplyID.Enabled = SumOfMoney.Enabled
                'BudID.Enabled = SumOfMoney.Enabled
                '補助比例和預算別改唯讀
                SupplyID.Enabled = False
                BudID.Enabled = False '預算別

                '學員補助不能提出申請
                If (Not IsDBNull(drv("AppliedResult")) AndAlso drv("AppliedResult").ToString = "N") OrElse (SupplyID.SelectedItem.Selected = True AndAlso SupplyID.SelectedValue.ToString() = "9") Then         '審核結果 或 補助比例代碼=9 補助比例0%
                    '970513 Andy ,若審核結果為不補助 
                    Checkbox1.Disabled = True
                    '970717 Andy ,若審核結果為不補助則是否提出申請應該是無法勾選的,且為不勾選的
                    Checkbox1.Checked = False
                    If (drv("AppliedResult").ToString = "N") Then e.Item.Cells(cst_是否補助).ToolTip += vbCrLf & "(學員資料審核)預算別：不補助"
                    SumOfMoney.Enabled = False '不可填入補助費用
                    'BudID.Enabled = False '預算別
                    e.Item.Cells(cst_是否補助).Text = "<font color='RED'>否</font>"
                    '20080606 andy 是否補助為「否」,個人支付=總費用
                    e.Item.Cells(cst_個人支付).Text = e.Item.Cells(cst_總費用).Text
                    SumOfMoney.Text = "0"
                End If

                '970513 Andy 學分為 0 或 出勤未達標準。
                '或 未填寫調查表。201407 AMU
                If Flag = 0 OrElse FlagStudy = 0 OrElse iFlag3 = 0 Then
                    Checkbox1.Disabled = True
                    Checkbox1.Checked = False
                    e.Item.Cells(cst_是否補助).Text = "<font color='RED'>否</font>"
                    SumOfMoney.Enabled = False '不可填入補助費用
                    'BudID.Enabled = False '預算別
                    '20080606 andy 是否補助為「否」,個人支付=總費用
                    e.Item.Cells(cst_個人支付).Text = e.Item.Cells(cst_總費用).Text
                    SumOfMoney.Text = "0"
                End If

                'For i As Integer = 0 To 2
                '    If DataGrid2.Visible Then
                '        e.Item.Cells(i).Attributes("onmouseover") = "if(document.getElementsById('" & DataGrid2.ClientID & "')){document.getElementById('" & DataGrid2.ClientID & "').style.display='inline';}"
                '        e.Item.Cells(i).Attributes("onmouseout") = "if(document.getElementById('" & DataGrid2.ClientID & "')){document.getElementById('" & DataGrid2.ClientID & "').style.display='none';}"
                '        e.Item.Cells(i).Style("CURSOR") = "hand"
                '    End If
                'Next

                SumOfMoney.Attributes("onchange") = "ChangeMoney(" & e.Item.ItemIndex + 1 & ",'" & SumOfMoney.ClientID & "','" & RemainSub.ClientID & "','" & MaxSub.ClientID & "','" & PayMoney.ClientID & "');"
                SumOfMoney.Attributes("onblur") = "ChangeMoney(" & e.Item.ItemIndex + 1 & ",'" & SumOfMoney.ClientID & "','" & RemainSub.ClientID & "','" & MaxSub.ClientID & "','" & PayMoney.ClientID & "');"
                SumOfMoney.Attributes("onFocus") = "ChangeMoney(" & e.Item.ItemIndex + 1 & ",'" & SumOfMoney.ClientID & "','" & RemainSub.ClientID & "','" & MaxSub.ClientID & "','" & PayMoney.ClientID & "');"

                'START 黑名單為不補助鎖定特定選項 2009/07/28 by waiming
                'Dim arr() As String
                'arr = Split(Me.ViewState("BlackIDNO"), ",")
                'For i As Int16 = 0 To arr.Length - 1
                '    If Convert.ToString(drv("IDNO")) = arr(i) Then
                '    End If
                'Next

                If gsBlackIDNO <> "" AndAlso gsBlackIDNO.IndexOf(Convert.ToString(drv("IDNO"))) > -1 Then
                    Dim s_BlackTip As String = String.Format("學號{0}.{1} {2}已受學員處分", stud_1.Text, drv("IDNO"), drv("Name"))
                    e.Item.Cells(cst_是否補助).Text = "<font color='RED'>否</font>" '是否補助
                    TIMS.Tooltip(e.Item.Cells(cst_是否補助), s_BlackTip)
                    SupplyID.SelectedValue = "9" '補助比例
                    SupplyID.Enabled = False
                    SumOfMoney.Text = "0" '補助費用
                    SumOfMoney.Enabled = False
                    Checkbox1.Checked = False '是否提出申請
                    Checkbox1.Disabled = True
                    BudID.SelectedIndex = 0 '預算別
                    BudID.Enabled = False
                    hidBlackMsg.Value += s_BlackTip & vbCrLf '加入單名單暫存(2009/07/28 判斷黑名單)
                End If
                'END 黑名單為不補助鎖定特定選項
                'Dim hid_totLeft As HtmlInputHidden = e.Item.FindControl("hid_totLeft")  '20091221 andy 
                'hid_totLeft.Value = Convert.ToString(totalleft)

                If Convert.ToString(drv("DegreeID")) = "" Then '檢查學歷欄位
                    SupplyID.Enabled = False    '補助比例不可更改 
                    Checkbox1.Disabled = True   '不可提出申請
                    SumOfMoney.Enabled = False  '補助費用
                    BudID.Enabled = False       '預算別 不可更改 
                    e.Item.Cells(cst_姓名).ForeColor = Color.Red
                    e.Item.Cells(cst_姓名).ToolTip = "此學員學歷資料未提供完備!"
                End If
        End Select

    End Sub

    '儲存
    Sub SaveData1()
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        'Dim da As SqlDataAdapter = Nothing
        'Const cst_是否補助 As Integer = 6
        Dim sql As String = ""
        Dim dtC As DataTable = Nothing '確定班級成員
        sql = "SELECT SOCID FROM CLASS_STUDENTSOFCLASS WHERE OCID='" & OCIDValue1.Value & "'"
        dtC = DbAccess.GetDataTable(sql, objconn)
        If dtC.Rows.Count = 0 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg2)
            Exit Sub
        End If

        Dim flagNoError As Boolean = True '沒有意外錯誤為true 
        Dim flagSaveOk As Boolean = False '正常結束為true
        Dim oConn As SqlConnection = DbAccess.GetConnection()
        Dim oTrans As SqlTransaction = DbAccess.BeginTrans(oConn)
        sql = " SELECT * FROM STUD_SUBSIDYCOST WHERE SOCID=@SOCID"
        Dim sCmd As New SqlCommand(sql, oConn, oTrans)

        Dim isSql As String = ""
        isSql &= " INSERT INTO STUD_SUBSIDYCOST (SOCID ,SUMOFMONEY ,PAYMONEY ,SUPPLYID ,BUDID ,MODIFYACCT ,MODIFYDATE)" & vbCrLf
        isSql &= " VALUES (@SOCID ,@SUMOFMONEY ,@PAYMONEY ,@SUPPLYID ,@BUDID ,@MODIFYACCT ,GETDATE())" & vbCrLf
        Dim iCmd As New SqlCommand(isSql, oConn, oTrans)

        Dim usSql As String = ""
        usSql &= " UPDATE STUD_SUBSIDYCOST" & vbCrLf
        usSql &= " SET SUMOFMONEY = @SUMOFMONEY" & vbCrLf
        usSql &= " ,PAYMONEY = @PAYMONEY" & vbCrLf
        usSql &= " ,SUPPLYID = @SUPPLYID" & vbCrLf
        usSql &= " ,BUDID = @BUDID" & vbCrLf
        usSql &= " ,MODIFYACCT = @MODIFYACCT" & vbCrLf
        usSql &= " ,MODIFYDATE = GETDATE()" & vbCrLf
        usSql &= " WHERE SOCID = @SOCID" & vbCrLf
        Dim uCmd As New SqlCommand(usSql, oConn, oTrans)

        Dim dsSql As String = ""
        dsSql &= " DELETE STUD_SUBSIDYCOST" & vbCrLf
        dsSql &= " WHERE SOCID = @SOCID" & vbCrLf
        Dim dCmd As New SqlCommand(dsSql, oConn, oTrans)

        Try
            For Each item As DataGridItem In DataGrid1.Items
                'client 判斷
                Dim iTotal As Integer = CInt(IIf(Trim(item.Cells(cst_總費用).Text) = "", 0, Trim(item.Cells(cst_總費用).Text)))
                Dim SumOfMoney As TextBox = item.FindControl("SumOfMoney") '補助費用
                Dim Checkbox1 As HtmlInputCheckBox = item.FindControl("Checkbox1")
                Dim RemainSub As HtmlInputHidden = item.FindControl("RemainSub")
                Dim PayMoney As HtmlInputHidden = item.FindControl("PayMoney") '個人支付費用
                Dim SupplyID As DropDownList = item.FindControl("SupplyID")
                Dim BudID As DropDownList = item.FindControl("BudID")
                Dim setid As HtmlInputHidden = item.FindControl("setid")
                Dim ocid As HtmlInputHidden = item.FindControl("ocid")
                Dim socid As HtmlInputHidden = item.FindControl("socid")
                Dim delFlag As Boolean = False '是否已執行刪除。 true:已執行
                'Dim iSOCID As Integer = Val(DataGrid1.DataKeys(item.ItemIndex)) 'DataKeys
                socid.Value = TIMS.ClearSQM(socid.Value)
                ff = "SOCID='" & socid.Value & "'"
                If dtC.Select(ff).Length = 0 Then
                    flagNoError = False '成員異常'異常
                    Common.MessageBox(Me, TIMS.cst_NODATAMsg2)
                    Exit For
                End If

                If Not Checkbox1.Disabled Then
                    If Checkbox1.Checked Then
                        Dim iSUMOFMONEY As Integer = Val(SumOfMoney.Text) '此次可用補助額
                        Dim iPAYMONEY As Integer = iTotal - iSUMOFMONEY '個人支付費用
                        'iSOCID  'DataGrid1.DataKeys(item.ItemIndex)
                        Dim dt As New DataTable
                        With sCmd
                            .Parameters.Clear()
                            .Parameters.Add("SOCID", SqlDbType.Int).Value = Val(socid.Value)
                            dt.Load(.ExecuteReader())
                        End With
                        If dt.Rows.Count = 0 Then
                            With iCmd
                                .Parameters.Clear()
                                .Parameters.Add("SOCID", SqlDbType.Int).Value = Val(socid.Value)
                                .Parameters.Add("SUMOFMONEY", SqlDbType.Int).Value = iSUMOFMONEY
                                .Parameters.Add("PAYMONEY", SqlDbType.Int).Value = iPAYMONEY
                                .Parameters.Add("SUPPLYID", SqlDbType.VarChar).Value = TIMS.GetValue1(SupplyID.SelectedValue) '補助比例
                                .Parameters.Add("BUDID", SqlDbType.VarChar).Value = TIMS.GetValue1(BudID.SelectedValue) '預算別
                                .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                                .ExecuteNonQuery()
                            End With
                        Else
                            With uCmd
                                .Parameters.Clear()
                                .Parameters.Add("SUMOFMONEY", SqlDbType.Int).Value = iSUMOFMONEY
                                .Parameters.Add("PAYMONEY", SqlDbType.Int).Value = iPAYMONEY
                                .Parameters.Add("SUPPLYID", SqlDbType.VarChar).Value = TIMS.GetValue1(SupplyID.SelectedValue) '補助比例
                                .Parameters.Add("BUDID", SqlDbType.VarChar).Value = TIMS.GetValue1(BudID.SelectedValue) '預算別
                                .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)

                                .Parameters.Add("SOCID", SqlDbType.Int).Value = Val(socid.Value)
                                .ExecuteNonQuery()
                            End With
                        End If
                    Else
                        '執行刪除一筆資料。
                        With dCmd
                            .Parameters.Clear()
                            .Parameters.Add("SOCID", SqlDbType.Int).Value = Val(socid.Value)
                            .ExecuteNonQuery()
                        End With
                        delFlag = True
                    End If
                End If

                '20080717  Andy  不補助則刪除
                If Not delFlag AndAlso Checkbox1.Disabled AndAlso Not Checkbox1.Checked Then
                    With dCmd
                        .Parameters.Clear()
                        .Parameters.Add("SOCID", SqlDbType.Int).Value = Val(socid.Value)
                        .ExecuteNonQuery()
                    End With
                    delFlag = True
                End If
                'i += 1
            Next
            'DbAccess.UpdateDataTable(dt, da)
            DbAccess.CommitTrans(oTrans)
            flagSaveOk = True
        Catch ex As Exception
            Dim strErrmsg As String = ""
            strErrmsg += "/*  ex.ToString: */" & vbCrLf
            strErrmsg += ex.ToString & vbCrLf
            strErrmsg += "/* sql: */" & vbCrLf
            strErrmsg += sql & vbCrLf
            strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)
            Common.MessageBox(Me.Page, "發生錯誤：" & ex.ToString)
            DbAccess.RollbackTrans(oTrans)
            TIMS.CloseDbConn(oConn)
            Exit Sub
        End Try
        TIMS.CloseDbConn(oConn)

        If flagSaveOk AndAlso flagNoError Then
            'Common.MessageBox(Me, "儲存成功")
            Common.MessageBox(Me, cst_SAVE_OK_msg1)
            'Button1_Click(sender, e)
            Call sUtl_Search1()
        End If
    End Sub

    '儲存
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim drC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drC Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg2)
            Exit Sub
        End If

        Call SaveData1()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn) '判斷機構是否只有一個班級
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        DataGridTable.Style("display") = "none"
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        '如果只有一個班級
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
        DataGridTable.Style("display") = "none"
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub

End Class
