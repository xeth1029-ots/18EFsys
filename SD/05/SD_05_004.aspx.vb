Partial Class SD_05_004
    Inherits AuthBasePage

    'CELLS
    Const cst_aCLASSCNAME2 As Integer = 0
    Const cst_aStudentID As Integer = 1
    Const cst_aName As Integer = 2
    Const cst_aStudStatus As Integer = 3
    Const cst_aRejectTDate As Integer = 4 ' 離退訓日期
    Const cst_aReason As Integer = 5 '離退訓原因
    Const cst_aRejectCDate As Integer = 6 ' 申請日期
    Const cst_a功能 As Integer = 7

    'SELECT * FROM KEY_PLAN
    Const cst_tplanid28_aspx As String = "SD_05_004_add.aspx" '產投專用 '28 / 54
    Const cst_tplanid02_aspx As String = "SD_05_004_add2.aspx" '非產投專用(委訓單位) '70：區域產業據點職業訓練計畫(在職)
    Const cst_tplanid06_aspx As String = "SD_05_004_add3.aspx" ' '06：自辦在職 / 07：接受企業委託訓練

    'sql = "SELECT * FROM KEY_REJECTTREASON ORDER BY RTReasonID"
    'Optional ByVal OldRTReasonID As String = ""
    'SELECT * FROM KEY_REJECTTREASON WHERE SORT2 IS NOT NULL ORDER BY RTReasonID 
    'SELECT * FROM KEY_REJECTTREASON WHERE SORT3 IS NOT NULL ORDER BY RTReasonID 
    'Stud_LeaveTraining
    'Class_StudentsOfClass RejectDayIn14 WkAheadOfSch
    'Dim FunDr As DataRow
    Dim Days1 As Integer = 0
    Dim Days2 As Integer = 0

    'Dim au As New cAUTH
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
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End
        '分頁設定 Start
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End
        '取出設定天數檔 Start
        TIMS.Get_SysDays(Days1, Days2)
        '取出設定天數檔 End

        If Not IsPostBack Then
            msg.Text = ""
            Table4.Visible = False
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');SetOneOCID();"
        Button5.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

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

        Button1.Attributes("onclick") = "javascript:return search()"
        '產業人才投資方案
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Button2.Attributes("onclick") = "javascript:add_ShowAlert();"
        End If

        '檢查帳號的功能權限-----------------------------------Start
        Button2.Enabled = True '新增   ----------------新增功能 因測試暫時拿掉
        'If Not au.blnCanAdds Then
        '    Button2.Enabled = False
        '    TIMS.Tooltip(Button2, "")
        '    TIMS.Tooltip(Button2, "無新增權限")
        'End If
        Button1.Enabled = True '查詢  ----------------查詢功能 因測試暫時拿掉
        'If Not au.blnCanSech Then
        '    Button1.Enabled = False
        '    TIMS.Tooltip(Button1, "")
        '    TIMS.Tooltip(Button1, "無查詢權限")
        'End If
        '檢查帳號的功能權限-----------------------------------End

        If Not Session("_search") Is Nothing Then
            Dim MyValue As String = ""
            MyValue = TIMS.GetMyValue(Session("_search"), "prg")
            If MyValue = "SD_05_004" Then
                center.Text = TIMS.GetMyValue(Session("_search"), "center")
                RIDValue.Value = TIMS.GetMyValue(Session("_search"), "RIDValue")
                TMID1.Text = TIMS.GetMyValue(Session("_search"), "TMID1")
                OCID1.Text = TIMS.GetMyValue(Session("_search"), "OCID1")
                TMIDValue1.Value = TIMS.GetMyValue(Session("_search"), "TMIDValue1")
                OCIDValue1.Value = TIMS.GetMyValue(Session("_search"), "OCIDValue1")

                MyValue = TIMS.GetMyValue(Session("_search"), "PageIndex")
                PageControler1.PageIndex = MyValue
                MyValue = TIMS.GetMyValue(Session("_search"), "submit")
                If MyValue = "1" Then
                    'Button1_Click(sender, e)
                    Call sUtl_Search1()
                End If
            End If
            Session("_search") = Nothing
        End If

        If Not IsPostBack Then
            '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
            TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
        End If
    End Sub

    '查詢Sql
    Sub sUtl_Search1()
        'cst_aClassCName
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        DataGrid1.Columns(cst_aCLASSCNAME2).Visible = If(OCIDValue1.Value = "", True, False) '班別顯示/班別不顯示

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT  e.RID" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(e.ClassCName,e.CyclType) CLASSCNAME2" & vbCrLf
        sql &= " ,e.CyclType" & vbCrLf
        sql &= " ,e.OCID" & vbCrLf
        sql &= " ,e.IsClosed" & vbCrLf
        sql &= " ,e.FTDate" & vbCrLf
        sql &= " ,f.ClassID" & vbCrLf
        sql &= " ,FORMAT(cs.RejectTDate1,'yyyy/MM/dd') RejectTDate1" & vbCrLf
        sql &= " ,FORMAT(cs.RejectTDate2,'yyyy/MM/dd') RejectTDate2" & vbCrLf
        sql &= " ,FORMAT(case cs.StudStatus when 2 then cs.RejectTDate1 when 3 then cs.RejectTDate2 end,'yyyy/MM/dd') RejectTDateN" & vbCrLf
        sql &= " ,cs.StudStatus" & vbCrLf
        sql &= " ,case cs.StudStatus when 2 then '離訓' when 3 then '退訓' end StudStatusN" & vbCrLf
        sql &= " ,dbo.FN_CSTUDID2(cs.StudentID) StudentID" & vbCrLf
        sql &= " ,FORMAT(cs.RejectCDate,'yyyy/MM/dd') RejectCDate" & vbCrLf
        sql &= " ,d.Name" & vbCrLf
        sql &= " ,a.SLTID" & vbCrLf
        sql &= " ,case when a.NeedPay='Y' then '是' else '否' end NeedPay" & vbCrLf
        sql &= " ,ISNULL(a.SumOfPay,0) SumOfPay" & vbCrLf
        sql &= " ,ISNULL(a.HadPay,0) HadPay" & vbCrLf
        sql &= " ,b.Reason" & vbCrLf
        sql &= " ,cs.RejectDayIn14" & vbCrLf
        sql &= " ,cs.RejectSOCID" & vbCrLf
        sql &= " ,cs.MakeSOCID" & vbCrLf
        sql &= " ,cs.WkAheadOfSch" & vbCrLf
        sql &= " FROM CLASS_CLASSINFO e" & vbCrLf
        sql &= " JOIN ID_CLASS f on f.CLSID=e.CLSID" & vbCrLf
        sql &= " JOIN ID_Plan ip ON ip.PlanID =e.PlanID" & vbCrLf
        sql &= " JOIN Class_StudentsOfClass cs on cs.OCID=e.OCID" & vbCrLf
        sql &= " JOIN Stud_StudentInfo d ON d.SID=cs.SID" & vbCrLf
        sql &= " JOIN Stud_LeaveTraining a ON a.SOCID=cs.SOCID" & vbCrLf
        sql &= " LEFT JOIN Key_RejectTReason b ON b.RTReasonID=cs.RTReasonID" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        If sm.UserInfo.RID = "A" Then
            sql &= " and ip.TPlanID ='" & sm.UserInfo.TPlanID & "'" & vbCrLf
            sql &= " and ip.Years ='" & sm.UserInfo.Years & "'" & vbCrLf
        Else
            sql &= " and ip.PlanID ='" & sm.UserInfo.PlanID & "'" & vbCrLf
        End If
        cjobValue.Value = TIMS.ClearSQM(cjobValue.Value)
        If cjobValue.Value <> "" Then sql &= " and e.CJOB_UNKEY=" & cjobValue.Value & vbCrLf
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value <> "" Then sql &= " and e.RID='" & RIDValue.Value & "'" & vbCrLf

        If OCIDValue1.Value <> "" Then
            'sqlStr &= " and e.OCID='66577'" & vbCrLf
            sql &= " and e.OCID='" & OCIDValue1.Value & "'" & vbCrLf
        End If

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)

        msg.Text = "查無資料!"
        Table4.Visible = False
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            Table4.Visible = True

            PageControler1.PageDataTable = dt
            PageControler1.PrimaryKey = "SLTID"
            PageControler1.Sort = "ClassID,CyclType,StudentID"
            PageControler1.ControlerLoad()
        End If

    End Sub

    '查詢
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '查詢Sql
        Call sUtl_Search1()
    End Sub

    '新增動作
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'Dim dr As DataRow
        Dim vsErrorMsg As String = ""
        vsErrorMsg = ""

        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If OCID1.Text = "" OrElse OCIDValue1.Value = "" Then
            Common.MessageBox(Me, "請選擇班別！")
            Exit Sub
        End If

        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "請重新選擇班別！")
            Exit Sub
        End If

        'Class_ClassInfo.APPLIEDRESULTM --select top 10 APPLIEDRESULTM from Class_ClassInfo
        If Convert.ToString(drCC("AppliedResultM")) = "Y" Then
            vsErrorMsg += "學員經費審核結果已經通過，不可新增" & vbCrLf
        End If

        '回傳錯誤訊息，超過結訓日期n天，停用新增功能
        Dim tmpErrMsg1 As String = TIMS.Chk_StopUseDate(Me, Days1, Days2, Convert.ToString(drCC("ISCLOSED")), CDate(drCC("FTDATE")))
        ''測試用
        'If TIMS.sUtl_ChkTest() Then
        '    tmpErrMsg1 = "" '測試用
        'End If
        If tmpErrMsg1 <> "" Then vsErrorMsg &= tmpErrMsg1

        Dim url1 As String = ""

        If TIMS.sUtl_ChkTest Then
            Common.MessageBox(Me, vsErrorMsg)
            '沒有錯誤訊息可繼續
            KeepSearchStr()
            'Response.Redirect(sUtl_GetUrl1() & "&Proecess=add" & "&TMID=" & TMIDValue1.Value & "&OCID=" & OCIDValue1.Value)
            url1 = sUtl_GetUrl1() & "&Proecess=add" & "&TMID=" & TMIDValue1.Value & "&OCID=" & OCIDValue1.Value
            Call TIMS.Utl_Redirect(Me, objconn, url1)
        End If

        '有錯誤訊息將無法繼續
        If vsErrorMsg <> "" Then
            Common.MessageBox(Me, vsErrorMsg)
            Exit Sub
        End If

        '沒有錯誤訊息可繼續
        KeepSearchStr()
        'Response.Redirect(sUtl_GetUrl1() & "&Proecess=add" & "&TMID=" & TMIDValue1.Value & "&OCID=" & OCIDValue1.Value)
        url1 = sUtl_GetUrl1() & "&Proecess=add" & "&TMID=" & TMIDValue1.Value & "&OCID=" & OCIDValue1.Value
        Call TIMS.Utl_Redirect(Me, objconn, url1)

    End Sub

    '刪除
    Public Shared Function Del_LeaveTraining(ByVal MyPage As Page, ByVal tConn As SqlConnection, ByVal eCmdArg As String) As Boolean
        Dim Rst As Boolean = False
        Dim dt As DataTable
        Dim da As SqlDataAdapter = Nothing
        Dim trans As SqlTransaction = Nothing
        Dim sql As String = ""
        Dim sm As SessionModel = SessionModel.Instance()

        sql = "SELECT * FROM Stud_LeaveTraining" & eCmdArg
        dt = DbAccess.GetDataTable(sql, tConn)
        If dt.Rows.Count <> 1 Then
            Common.MessageBox(MyPage, "查無資料，刪除失敗!")
            Return Rst
        End If
        Dim dr As DataRow = Nothing
        dr = DbAccess.GetOneRow(sql, tConn)
        If Not dr Is Nothing Then
            Dim SOCID As Integer = dr("SOCID")

            '取得學員資料。
            Dim drS As DataRow = TIMS.Get_StudData(SOCID, tConn)
            Dim sDelNote As String = ""
            sDelNote = "刪除離退訓作業 Stud_LeaveTraining [" & drS("PlanName") & "]-[" & drS("OrgName") & "]-[" & drS("CLASSCNAME") & "]"
            Dim iMRqID As Integer = Val(MyPage.Request("ID"))
            TIMS.InsertDelLog(sm.UserInfo.UserID, iMRqID, sm.UserInfo.DistID, sDelNote,
                sm.UserInfo.OrgID, drS("RID"), drS("PlanID"), drS("ComIDNO"), drS("SeqNo"), drS("OCID"))
            Try
                sql = "SELECT 'x' FROM Class_StudentsOfClass WHERE SOCID='" & SOCID & "'"
                dt = DbAccess.GetDataTable(sql, tConn)

                trans = DbAccess.BeginTrans(tConn)
                If dt.Rows.Count > 0 Then
                    sql = "SELECT * FROM Class_StudentsOfClass WHERE SOCID='" & SOCID & "'"
                    dt = DbAccess.GetDataTable(sql, da, trans)
                    dr = dt.Rows(0)
                    dr("RejectTDate1") = Convert.DBNull '離訓日期 
                    dr("RejectTDate2") = Convert.DBNull '退訓日期
                    dr("RTReasonID") = Convert.DBNull '離退訓原因代碼 
                    dr("RTReasoOther") = Convert.DBNull '離退訓原因(其他)
                    dr("TrainHours") = Convert.DBNull '參訓時數
                    dr("WkAheadOfSch") = Convert.DBNull '提前就業

                    dr("StudStatus") = 1 '學員狀態 
                    dr("ModifyAcct") = sm.UserInfo.UserID
                    dr("ModifyDate") = Now
                    DbAccess.UpdateDataTable(dt, da, trans)
                End If

                sql = "DELETE Stud_LeaveTraining" & eCmdArg
                DbAccess.ExecuteNonQuery(sql, trans)

                '如果是三合一的學生，要更新狀態
                sql = "SELECT * FROM Adp_TRNData WHERE SOCID='" & SOCID & "'"
                dt = DbAccess.GetDataTable(sql, da, trans)
                If dt.Rows.Count <> 0 Then
                    dr = dt.Rows(0)
                    dr("ARVL_STATE") = 1
                    dr("ARVL_STP_DATE") = Convert.DBNull
                    dr("ARVL_STP_REASON") = Convert.DBNull
                    dr("SEND_DATE") = Convert.DBNull
                    dr("TIMSModifyDate") = Now
                    DbAccess.UpdateDataTable(dt, da, trans)
                End If
                sql = "SELECT * FROM Adp_DGTRNData WHERE SOCID='" & SOCID & "'"
                dt = DbAccess.GetDataTable(sql, da, trans)
                If dt.Rows.Count <> 0 Then
                    dr = dt.Rows(0)
                    dr("ARVL_STATE") = 1
                    dr("STOP_DATE") = Convert.DBNull
                    dr("STOP_REASON") = Convert.DBNull
                    dr("STOP_STATE") = Convert.DBNull
                    dr("TIMSModifyDate") = Now
                    DbAccess.UpdateDataTable(dt, da, trans)
                End If
                sql = "SELECT * FROM Adp_GOVTRNData WHERE SOCID='" & SOCID & "'"
                dt = DbAccess.GetDataTable(sql, da, trans)
                If dt.Rows.Count <> 0 Then
                    dr = dt.Rows(0)
                    dr("ARVL_STATE") = 1
                    dr("ARVL_STP_DATE") = Convert.DBNull
                    dr("ARVL_STP_REASON") = Convert.DBNull
                    dr("TIMSModifyDate") = Now
                    DbAccess.UpdateDataTable(dt, da, trans)
                End If
                DbAccess.CommitTrans(trans)
                Common.MessageBox(MyPage, "刪除成功!")
                Rst = True
            Catch ex As Exception
                DbAccess.RollbackTrans(trans)
                Common.MessageBox(MyPage, "刪除失敗!" & ex.ToString)
                Return Rst
                'Throw ex
                'Common.RespWrite(Me, ex.ToString)
            End Try
        End If
        Return Rst
    End Function

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Select Case e.CommandName
            Case "edit"
                KeepSearchStr()
                'Response.Redirect(sUtl_GetUrl1() & "&" & e.CommandArgument)
                Dim url1 As String = sUtl_GetUrl1() & "&" & e.CommandArgument
                Call TIMS.Utl_Redirect(Me, objconn, url1)

            Case "del"
                If Del_LeaveTraining(Me, objconn, e.CommandArgument) Then
                    'Button1_Click(Button1, e)
                    Call sUtl_Search1()
                End If
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem

                Dim btn1 As Button = e.Item.FindControl("Button3") '修改／查詢
                Dim btn2 As Button = e.Item.FindControl("Button4") '刪除
                btn1.CommandArgument = "Proecess=edit&SLTID=" & drv("SLTID") & "&TMID=" & TMIDValue1.Value & "&OCID=" & Convert.ToString(drv("OCID"))
                btn2.Attributes("onclick") = "return confirm('確定要刪除這一筆離退訓紀錄嗎?');"
                btn2.CommandArgument = " WHERE SLTID = " & Convert.ToString(drv("SLTID"))

                If drv("IsClosed") = "Y" Then
                    If sm.UserInfo.RoleID <= 1 Then
                        If DateDiff(DateInterval.Day, drv("FTDate"), Now) > Days2 Then
                            'If TIMS.sUtl_ChkTest() Then
                            '    btn1.Text = "查看(可修改) (測試中!!)" '測試
                            '    btn2.Enabled = True '測試
                            '    TIMS.Tooltip(btn2, "只可查看，不可刪除 (測試中!!)") '正式
                            'Else
                            '    btn1.Text = "查看" '正式
                            '    btn2.Enabled = False '正式
                            '    TIMS.Tooltip(btn2, "只可查看，不可刪除(時間限制)") '正式
                            'End If
                            btn1.Text = "查看" '正式
                            btn2.Enabled = False '正式
                            TIMS.Tooltip(btn2, "只可查看，不可刪除") '正式
                        Else
                            btn2.Enabled = True
                        End If
                    Else
                        btn2.Enabled = False
                        TIMS.Tooltip(btn2, "無刪除權限")
                        If DateDiff(DateInterval.Day, drv("FTDate"), Now) > Days2 Then
                            btn1.Text = "查看"
                        End If
                    End If
                Else
                    btn1.Enabled = True
                    'If Not au.blnCanMod Then
                    '    btn1.Enabled = False
                    '    TIMS.Tooltip(btn1, "無修改權限")
                    '    If Val(sm.UserInfo.LID) < 2 Then
                    '        btn1.Text = "查看"
                    '        btn1.Enabled = True
                    '        TIMS.Tooltip(btn1, "提供查看權限")
                    '    End If
                    'End If
                    btn2.Enabled = True
                    'If Not au.blnCanDel Then
                    '    btn2.Enabled = False
                    '    TIMS.Tooltip(btn2, "無刪除權限")
                    'End If
                End If

                '被遞補學員 為正式學員
                If Convert.ToString(drv("MakeSOCID")) <> "" Then
                    btn2.Enabled = False
                    TIMS.Tooltip(btn2, " 已有被遞補學員：" & TIMS.GetSOCIDName(Convert.ToString(drv("MakeSOCID")), objconn))
                End If
                'Case ListItemType.Header, ListItemType.Footer
                'Case Else
        End Select

    End Sub

    '保留Session
    Sub KeepSearchStr()
        Session("_search") = Nothing

        Dim str_search1 As String = ""
        str_search1 = "prg=SD_05_004"
        str_search1 &= "&center=" & center.Text
        str_search1 &= "&RIDValue=" & RIDValue.Value
        str_search1 &= "&TMID1=" & TMID1.Text
        str_search1 &= "&OCID1=" & OCID1.Text
        str_search1 &= "&TMIDValue1=" & TMIDValue1.Value
        str_search1 &= "&OCIDValue1=" & OCIDValue1.Value
        str_search1 &= "&PageIndex=" & DataGrid1.CurrentPageIndex + 1
        str_search1 &= If(DataGrid1.Visible, "&submit=1", "&submit=0")

        Session("_search") = str_search1
    End Sub

    '單一班級搜尋
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        '判斷機構是否只有一個班級
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn)
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        Table4.Visible = False
        '如果只有一個班級
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
        Table4.Visible = False
    End Sub

    '單一班級搜尋
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub

    '覆寫，不執行 MyBase.VerifyRenderingInServerForm 方法，解決執行 RenderControl 產生的錯誤   
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)

    End Sub

    '匯出
    Protected Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        SUB_EXP1()
    End Sub

    Sub SUB_EXP1()
        'Dim cst_功能 As Integer = 10
        DataGrid1.AllowPaging = False
        DataGrid1.Columns(cst_a功能).Visible = False
        'DataGrid1.Columns(0).Visible = False '班別不顯示
        'If OCIDValue1.Value = "" Then
        '    DataGrid1.Columns(0).Visible = True '班別顯示
        'End If
        DataGrid1.EnableViewState = False  '把ViewState給關了

        Call sUtl_Search1()

        Dim sFileName1 As String = "離退訓作業"

        Dim strSTYLE As String = ""
        '套CSS值
        strSTYLE &= ("<style>")
        strSTYLE &= ("td{mso-number-format:""\@"";}")
        strSTYLE &= ("</style>")

        DataGrid1.AllowPaging = False
        DataGrid1.Columns(cst_a功能).Visible = False
        DataGrid1.EnableViewState = False  '把ViewState給關了

        Dim objStringWriter As New System.IO.StringWriter
        Dim objHtmlTextWriter As New System.Web.UI.HtmlTextWriter(objStringWriter)
        Div1.RenderControl(objHtmlTextWriter)

        Dim strHTML As String = ""
        strHTML &= (TIMS.sUtl_AntiXss(Convert.ToString(objStringWriter)))

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", strHTML)
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)

        DataGrid1.AllowPaging = True
        DataGrid1.Columns(cst_a功能).Visible = True
        TIMS.Utl_RespWriteEnd(Me, objconn, "") '  Response.End()
    End Sub

    '分頁 SD_05_004_add
    Function sUtl_GetUrl1() As String
        Dim rst As String = ""
        '非產投用 
        Dim sMRqID As String = "?ID=" & TIMS.ClearSQM(Request("ID"))
        Select Case Convert.ToString(sm.UserInfo.TPlanID)
            Case TIMS.Cst_TPlanID06 '"06" '在職 
                rst = cst_tplanid06_aspx & sMRqID 'Request("ID")
            Case TIMS.Cst_TPlanID07 '計畫：07:接受企業委託訓練
                rst = cst_tplanid06_aspx & sMRqID 'Request("ID")
            Case Else '職前
                rst = cst_tplanid02_aspx & sMRqID 'Request("ID")
        End Select
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '產投專用
            rst = cst_tplanid28_aspx & sMRqID  'Request("ID")
        End If
        Return rst
    End Function
End Class