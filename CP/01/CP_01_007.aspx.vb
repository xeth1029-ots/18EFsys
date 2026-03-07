Partial Class CP_01_007
    Inherits AuthBasePage

#Region "參數/變數"

    'Dim iPYNum14 As Integer = 1 'iPYNum14 = TIMS.sUtl_GetPYNum14(Me)
    'Dim prtFilename As String = "" '列印表件名稱
    'CP_01_007*.jrxml
    'CP_01_007
    'CP_01_007_1
    'CP_01_007_b
    'CP_01_007_1_b

    'Dim blnCanAdds As Boolean = False '新增
    'Dim blnCanMod As Boolean = False '修改
    'Dim blnCanDel As Boolean = False '刪除
    'Dim blnCanSech As Boolean = False '查詢
    'Dim blnCanPrnt As Boolean = False '列印

#End Region

#Region "NOUSE"
    'Const cst_printFN2 As String = "CP_01_007_1_b"
    'Const cst_printFN1 As String = "CP_01_007_b"
    'Const cst_printFN2_2 As String = "CP_01_007_1_b2" '空白 資料
    'Const cst_printFN1_2 As String = "CP_01_007_b2" '有資料
    'Const cst_printFN2_3 As String = "CP_01_007_1_b3" '空白 資料
    'Const cst_printFN1_3 As String = "CP_01_007_b3" '有資料
#End Region

    Const cst_printFN2_4 As String = "CP_01_007_1_b4" '空白 資料
    Const cst_printFN1_4 As String = "CP_01_007_b4" '有資料

    Const cst_add As String = "ADD"
    Const cst_edit As String = "EDIT"
    Const cst_del As String = "DEL"

    Dim sURL_CP01007ADD As String = ""
    'Dim STR_PRT_FILE_N1 As String = "" '有資料
    'Dim STR_PRT_FILE_N2 As String = "" '空白 資料
    Dim fg_SUPER1 As Boolean = False 'TIMS.IsSuperUser(Me, 1) '是否為(後台)系統管理者 
    Dim fg_TEST1 As Boolean = False 'TIMS.sUtl_ChkTest() '測試

    Dim objconn As SqlConnection
    Dim flag_ROC As Boolean = TIMS.CHK_REPLACE2ROC_YEARS()

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call TIMS.GetAuth(Me, blnCanAdds, blnCanMod, blnCanDel, blnCanSech, blnCanPrnt) '2011 取得功能按鈕權限值
        '檢查Session是否存在 End
        fg_SUPER1 = TIMS.IsSuperUser(Me, 1) '是否為(後台)系統管理者 
        fg_TEST1 = TIMS.sUtl_ChkTest() '測試

        Dim rqMID As String = TIMS.Get_MRqID(Me)
        sURL_CP01007ADD = String.Concat(TIMS.URL_CP01007ADD, rqMID)
        'STR_PRT_FILE_N1 = cst_printFN1_3 '有資料
        'STR_PRT_FILE_N2 = cst_printFN2_3 '空白 資料
        PageControler1.PageDataGrid = DataGrid1

        If Not IsPostBack Then
            Ccreate1()
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
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

#Region "(No Use)"

        '檢查帳號的功能權限-----------------------------------Start
        'If sm.UserInfo.FunDt Is Nothing Then
        '    Common.RespWrite(Me, "<script>alert('Session過期');</script>")
        '    Common.RespWrite(Me, "<script>top.location.href='../../logout.aspx';</script>")
        'Else
        '    Dim FunDt As DataTable = sm.UserInfo.FunDt
        '    Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & Request("ID") & "'")

        '    If FunDrArray.Length = 0 Then
        '        Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
        '        Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
        '    Else
        '        FunDr = FunDrArray(0)
        '        If FunDr("Adds") = 1 Then
        '            Button2.Enabled = True
        '        Else
        '            Button2.Enabled = False
        '        End If
        '        If FunDr("Sech") = 1 Then
        '            Button1.Enabled = True
        '        Else
        '            Button1.Enabled = False
        '        End If
        '    End If
        'End If

        'Button2.Enabled = False
        'If blnCanAdds Then Button2.Enabled = True
        'Button1.Enabled = False
        'If blnCanSech Then Button1.Enabled = True
        '檢查帳號的功能權限-----------------------------------End

#End Region


    End Sub

    Sub Ccreate1()
        msg.Text = ""
        'Button1.Attributes("onclick") = "javascript:return search()"
        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID
        end_date.Text = If(flag_ROC, TIMS.Cdate17(Now.Date), TIMS.Cdate3(Now.Date)) '民國日期，by:20181018 / '西元日期，by:20181018
        DataGridTable.Visible = False

        If Session("SearchStr") IsNot Nothing Then
            Dim myValue As String = ""
            Dim sSearchStr As String = Convert.ToString(Session("SearchStr"))

            center.Text = TIMS.UrlDecode1(TIMS.GetMyValue(sSearchStr, "center")) 'Replace(MyValue, "%26", "&")
            RIDValue.Value = TIMS.GetMyValue(sSearchStr, "RIDValue")
            TMID1.Text = TIMS.UrlDecode1(TIMS.GetMyValue(sSearchStr, "TMID1"))
            OCID1.Text = TIMS.UrlDecode1(TIMS.GetMyValue(sSearchStr, "OCID1"))

            TMIDValue1.Value = TIMS.GetMyValue(sSearchStr, "TMIDValue1")
            OCIDValue1.Value = TIMS.GetMyValue(sSearchStr, "OCIDValue1")
            start_date.Text = TIMS.GetMyValue(sSearchStr, "start_date")
            end_date.Text = TIMS.GetMyValue(sSearchStr, "end_date")
            myValue = TIMS.GetMyValue(sSearchStr, "VisitItem")
            Common.SetListItem(VisitItem, myValue)
            myValue = TIMS.GetMyValue(sSearchStr, "PageIndex")
            If myValue <> "" Then PageControler1.PageIndex = Val(myValue)
            myValue = TIMS.GetMyValue(sSearchStr, "Button1")
            If myValue = "true" Then Call SSearch1()

            Session("SearchStr") = Nothing
        End If
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)  '若只有管理一個班級，自動協助帶出班級--by andy 2009-02-25
        PrintBlank.Attributes("onclick") = "javascript:return CheckAdd()"
    End Sub
    Sub SSearch1()
        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        start_date.Text = TIMS.ClearSQM(start_date.Text)
        end_date.Text = TIMS.ClearSQM(end_date.Text)
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then Exit Sub

        Dim Relship As String = TIMS.GET_RelshipforRID(RIDValue.Value, objconn)
        If Relship = "" Then Exit Sub

        Dim sql As String = ""
        sql &= " SELECT cc.OCID, cc.RID, cc.PLANID" & vbCrLf
        sql &= " ,a.SeqNo, a.Item10, a.Item10_1" & vbCrLf
        sql &= " ,a.Item10_Note ,a.RID RIDValue" & vbCrLf
        sql &= " ,FORMAT(a.ApplyDate,'yyyy/MM/dd') ApplyDate" & vbCrLf
        sql &= " ,cc.ClassCName" & vbCrLf
        sql &= " ,cc.CyclType" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME ,cc.CYCLTYPE) CLASSNAME2" & vbCrLf
        sql &= " ,d.OrgName" & vbCrLf
        sql &= " ,ic.ClassID" & vbCrLf
        sql &= " ,d.Relship " & vbCrLf
        sql &= " FROM CLASS_CLASSINFO cc " & vbCrLf
        sql &= " JOIN ID_PLAN ip ON ip.PLANID = cc.PLANID " & vbCrLf
        sql &= " JOIN VIEW_RIDNAME d ON cc.RID = d.RID " & vbCrLf
        sql &= " JOIN ID_CLASS ic ON cc.CLSID = ic.CLSID " & vbCrLf
        sql &= " LEFT JOIN CLASS_UNEXPECTTEL a ON cc.ocid = a.ocid " & vbCrLf
        sql &= " WHERE d.Relship LIKE @Relship " & vbCrLf

        Dim myParam As New Hashtable From {{"Relship", Relship + "%"}}
        If sm.UserInfo.RID = "A" Then
            sql &= " AND ip.TPlanID = @TPlanID " & vbCrLf
            sql &= " AND ip.Years = @Years " & vbCrLf
            myParam.Add("TPlanID", sm.UserInfo.TPlanID)
            myParam.Add("Years", sm.UserInfo.Years)
        Else
            sql &= " AND ip.PlanID = @PlanID " & vbCrLf
            myParam.Add("PlanID", sm.UserInfo.PlanID)
        End If
        Select Case VisitItem.SelectedValue
            Case "1" '未抽訪
                sql &= " AND a.OCID IS NULL " & vbCrLf
            Case "2" '已抽訪
                sql &= " AND a.OCID IS NOT NULL " & vbCrLf
            Case Else '"3" '全部
        End Select
        If OCIDValue1.Value <> "" Then
            sql &= "  AND cc.OCID = @OCID " & vbCrLf
            myParam.Add("OCID", OCIDValue1.Value)
        End If
        If start_date.Text <> "" Then
            sql &= " AND CONVERT(VARCHAR, a.ApplyDate, 111) >= @ApplyDate1 " & vbCrLf
            If flag_ROC Then
                myParam.Add("ApplyDate1", TIMS.Cdate18(start_date.Text.Trim))  'edit，by:20181018
            Else
                myParam.Add("ApplyDate1", start_date.Text.Trim)  'edit，by:20181018
            End If
        End If
        If end_date.Text <> "" Then
            sql &= " AND CONVERT(VARCHAR, a.ApplyDate, 111) <= @ApplyDate2 " & vbCrLf
            If flag_ROC Then
                myParam.Add("ApplyDate2", TIMS.Cdate18(end_date.Text.Trim))  'edit，by:20181018
            Else
                myParam.Add("ApplyDate2", end_date.Text.Trim)  'edit，by:20181018
            End If
        End If

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, myParam)
        msg.Text = "查無資料!"
        DataGridTable.Visible = False

        If dt.Rows.Count > 0 Then
            msg.Text = ""
            DataGridTable.Visible = True
            If VisitItem.SelectedIndex = 0 Then '未抽訪
                DataGrid1.Columns(3).Visible = False
                DataGrid1.Columns(4).Visible = False
                DataGrid1.Columns(5).Visible = False
                DataGrid1.Columns(6).Visible = True
            Else
                DataGrid1.Columns(3).Visible = True
                DataGrid1.Columns(4).Visible = True
                DataGrid1.Columns(5).Visible = True
                DataGrid1.Columns(6).Visible = True
            End If
            PageControler1.Sort = "RIDValue,ClassID,CyclType"
            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
        End If
    End Sub

    '查詢鈕
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call SSearch1()
    End Sub

    '新增
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If RIDValue.Value = "" Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg11)
            Return
        End If
        If OCIDValue1.Value = "" Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg11B)
            Return
        End If
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg11B)
            Return
        End If

        GetSearchStr()
        'Call TIMS.CloseDbConn(objconn)

        Dim MyValue1 As String = ""
        TIMS.SetMyValue(MyValue1, "RIDValue", RIDValue.Value)
        TIMS.SetMyValue(MyValue1, "TMIDValue1", TMIDValue1.Value)
        TIMS.SetMyValue(MyValue1, "OCIDValue1", OCIDValue1.Value)
        TIMS.SetMyValue(MyValue1, "State", "Add")
        TIMS.Utl_Redirect1(Me, sURL_CP01007ADD & MyValue1, objconn)
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Select Case UCase(e.CommandName)
            Case UCase(cst_edit)
                GetSearchStr()
                Session("_SearchStr") = Me.ViewState("_SearchStr")
                Call TIMS.CloseDbConn(objconn)
                TIMS.Utl_Redirect1(Me, e.CommandArgument)
            Case UCase(cst_add)
                GetSearchStr()
                Session("_SearchStr") = Me.ViewState("_SearchStr")
                Call TIMS.CloseDbConn(objconn)
                TIMS.Utl_Redirect1(Me, e.CommandArgument)
            Case UCase(cst_del)
                Dim vOCID As String = TIMS.GetMyValue(e.CommandArgument, "OCID")
                Dim vSeqNo As String = TIMS.GetMyValue(e.CommandArgument, "SeqNo")

                Dim myParam_d As New Hashtable From {{"OCID", vOCID}, {"SeqNo", vSeqNo}}
                Dim sql_d As String = ""
                sql_d &= " DELETE Class_UnexpectTelApply WHERE OCID = @OCID " & vbCrLf
                sql_d &= " AND SeqNo = @SeqNo " & vbCrLf
                DbAccess.ExecuteNonQuery(sql_d, objconn, myParam_d)

                Dim myParam_d2 As New Hashtable From {{"OCID", vOCID}, {"SeqNo", vSeqNo}}
                Dim sql_d2 As String = ""
                sql_d2 &= " DELETE Class_UnexpectTel WHERE OCID = @OCID " & vbCrLf
                sql_d2 &= " AND SeqNo = @SeqNo " & vbCrLf
                DbAccess.ExecuteNonQuery(sql_d2, objconn, myParam_d2)

                Common.MessageBox(Me, "刪除成功!!")
                Call SSearch1()
        End Select
    End Sub

    '列印
    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)
                Dim ParentName As String = TIMS.Get_ParentRID(drv("Relship"), objconn)

                If ParentName <> "" Then e.Item.Cells(1).Text = String.Concat(ParentName, "-", drv("OrgName"))

                e.Item.Cells(2).Text = Convert.ToString(drv("CLASSNAME2"))

                Dim mybut1 As Button = e.Item.Cells(3).FindControl("Button3") '修改/查詢 State=Edit/View""
                Dim mybut2 As Button = e.Item.Cells(3).FindControl("Button4") '刪除
                Dim Button8 As Button = e.Item.Cells(3).FindControl("Button8") '列印 資料表
                Dim mybut4 As Button = e.Item.Cells(3).FindControl("Button7") '新增
                Dim PrintBlank2 As Button = e.Item.Cells(3).FindControl("PrintBlank2") '列印 空白表

                If VisitItem.SelectedIndex = 0 Then
                    '未抽訪只需顯示新增
                    mybut1.Visible = False
                    mybut2.Visible = False
                    Button8.Visible = False
                    mybut4.Visible = True
                    PrintBlank2.Visible = True
                    mybut4.Enabled = If(Left(Convert.ToString(drv("RID")), 1) = sm.UserInfo.RID, True, False) '新增

                Else
                    mybut4.Visible = False
                    Dim fg_VIEW1 As Boolean = (Convert.ToString(drv("RID")) <> sm.UserInfo.RID)
                    '(檢視狀況)且超級測試環境，不提供檢視
                    If fg_VIEW1 AndAlso (fg_SUPER1 AndAlso fg_TEST1) Then fg_VIEW1 = False
                    If fg_VIEW1 Then
                        mybut1.Text = "檢視"
                        mybut2.Visible = False
                    End If

                    '未抽訪/正常/不正常
                    Dim v_CELLS_4_TEXT As String = If(Convert.ToString(drv("Item10")) = "1", "正常", If(Convert.ToString(drv("ApplyDate")) = "", "未抽訪", "不正常"))
                    e.Item.Cells(4).Text = v_CELLS_4_TEXT  '未抽訪/正常/不正常

                    Dim Reason As HtmlControls.HtmlTextArea = e.Item.FindControl("Reason")
                    '結論其他附加說明
                    TIMS.Tooltip(Reason, "結論其他附加說明..")
                    Reason.Value = Convert.ToString(drv("Item10_Note"))

                    mybut4.Visible = True
                    mybut1.Visible = False
                    mybut2.Visible = False
                    Button8.Visible = False
                    PrintBlank2.Visible = True
                    If Convert.ToString(drv("ApplyDate")) <> "" Then
                        e.Item.Cells(3).Text = If(flag_ROC, TIMS.Cdate17(drv("ApplyDate")), TIMS.Cdate3(drv("ApplyDate")))

                        mybut4.Visible = False
                        mybut1.Visible = True
                        mybut2.Visible = True
                        Button8.Visible = True
                        PrintBlank2.Visible = True
                    End If

                    Dim s_State As String = If(fg_VIEW1, "View", "Edit")
                    Dim s_StateTxt As String = If(fg_VIEW1, "查詢", "修改")
                    Dim MyValue1 As String = ""
                    TIMS.SetMyValue(MyValue1, "OCID", drv("OCID").ToString())
                    TIMS.SetMyValue(MyValue1, "SeqNo", drv("SeqNo").ToString())
                    TIMS.SetMyValue(MyValue1, "State", s_State)

                    Dim str_CmdArg_B1 As String = String.Concat(sURL_CP01007ADD, MyValue1)
                    mybut1.Text = s_StateTxt 'If(Convert.ToString(drv("RID")) <> sm.UserInfo.RID, "查詢", mybut1.Text)
                    mybut1.CommandArgument = str_CmdArg_B1

                    Dim CmdArg As String = ""
                    CmdArg &= "&OCID=" & drv("OCID").ToString
                    CmdArg &= "&SeqNo=" & drv("SeqNo").ToString
                    mybut2.Attributes("onclick") = TIMS.cst_confirm_delmsg1
                    mybut2.CommandArgument = CmdArg
                End If

                '列印 資料表
                '"OCID=" & drv("OCID") & "&SeqNo=" & drv("SeqNo") & "&Years2=" & sm.UserInfo.Years
                Dim MyValue3 As String = ""
                TIMS.SetMyValue(MyValue3, "RID", drv("RID"))
                TIMS.SetMyValue(MyValue3, "PLANID", drv("PLANID"))
                TIMS.SetMyValue(MyValue3, "OCID", drv("OCID"))
                TIMS.SetMyValue(MyValue3, "SeqNo", drv("SeqNo"))
                TIMS.SetMyValue(MyValue3, "Years2", sm.UserInfo.Years)
                Button8.Attributes("onclick") = ReportQuery.ReportScript(Me, cst_printFN1_4, MyValue3)

                '"&OCID=" & drv("OCID") & "&SeqNo=" & drv("SeqNo") & "&State=Add"
                Dim MyValue4 As String = ""
                TIMS.SetMyValue(MyValue4, "RID", drv("RID"))
                TIMS.SetMyValue(MyValue4, "PLANID", drv("PLANID"))
                TIMS.SetMyValue(MyValue4, "OCID", drv("OCID"))
                TIMS.SetMyValue(MyValue4, "SeqNo", drv("SeqNo").ToString())
                TIMS.SetMyValue(MyValue4, "State", "Add")
                mybut4.CommandArgument = sURL_CP01007ADD & MyValue4

                '列印 空白表
                '"OCID=" & drv("OCID") & "&Years2=" & sm.UserInfo.Years & "&OrgName=" & sm.UserInfo.OrgName & ""
                Dim MyValue2 As String = ""
                TIMS.SetMyValue(MyValue2, "RID", drv("RID"))
                TIMS.SetMyValue(MyValue2, "PLANID", drv("PLANID"))
                TIMS.SetMyValue(MyValue2, "OCID", drv("OCID"))
                TIMS.SetMyValue(MyValue2, "Years2", sm.UserInfo.Years)
                TIMS.SetMyValue(MyValue2, "OrgName", sm.UserInfo.OrgName)
                PrintBlank2.Attributes("onclick") = ReportQuery.ReportScript(Me, cst_printFN2_4, MyValue2)

        End Select

    End Sub

    ''' <summary>keep session UrlEncode1</summary>
    Sub GetSearchStr()
        Dim SearchStr As String = "pj=CP01007"

        TIMS.SetMyValue(SearchStr, "center", TIMS.UrlEncode1(center.Text))
        TIMS.SetMyValue(SearchStr, "RIDValue", RIDValue.Value)
        TIMS.SetMyValue(SearchStr, "TMID1", TIMS.UrlEncode1(TMID1.Text)) 'TMID1.Text)
        TIMS.SetMyValue(SearchStr, "OCID1", TIMS.UrlEncode1(OCID1.Text)) ' OCID1.Text)
        TIMS.SetMyValue(SearchStr, "TMIDValue1", TMIDValue1.Value)
        TIMS.SetMyValue(SearchStr, "OCIDValue1", OCIDValue1.Value)
        TIMS.SetMyValue(SearchStr, "start_date", start_date.Text)
        TIMS.SetMyValue(SearchStr, "end_date", end_date.Text)
        TIMS.SetMyValue(SearchStr, "VisitItem", VisitItem.SelectedValue)
        TIMS.SetMyValue(SearchStr, "PageIndex", DataGrid1.CurrentPageIndex + 1)
        TIMS.SetMyValue(SearchStr, "Button1", If(DataGrid1.Visible, "true", "false"))

        Session("SearchStr") = SearchStr
    End Sub

    '列印 列印空白 實地訪查紀錄表
    Private Sub PrintBlank_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintBlank.Click
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If RIDValue.Value = "" Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg11)
            Return
        End If
        If OCIDValue1.Value = "" Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg11B)
            Return
        End If
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg11B)
            Return
        End If
        'Dim Years As String,'Dim RID As String,'Dim OCID As String,
        '列印空白實地訪查紀錄表
        Dim Roc_Years As String = (sm.UserInfo.Years - 1911).ToString()

        '"Years=" & Years & "&OCID=" & OCID & "&Years2=" & sm.UserInfo.Years & "&OrgName=" & sm.UserInfo.OrgName
        Dim MyValue1 As String = ""
        TIMS.SetMyValue(MyValue1, "Years", Roc_Years)
        TIMS.SetMyValue(MyValue1, "RID", drCC("RID"))
        TIMS.SetMyValue(MyValue1, "PLANID", drCC("PLANID"))
        TIMS.SetMyValue(MyValue1, "OCID", OCIDValue1.Value)
        TIMS.SetMyValue(MyValue1, "Years2", sm.UserInfo.Years)
        TIMS.SetMyValue(MyValue1, "OrgName", sm.UserInfo.OrgName)
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN2_4, MyValue1)
    End Sub

End Class
