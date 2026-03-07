Partial Class CP_01_006
    Inherits AuthBasePage

    'Dim FunDr As DataRow=Nothing,'ReportQuery,'SQControl.aspx,
    '/**NEW 2014**/ CLASS_UNEXPECTVISITOR
    'CP_01_006*.jrxml'CP_01_006_1'CP_01_006'CP_01_006_1_b
    '紀錄表->紀錄表
    'CP_01_006_b (尚未完成)
    'Dim iPYNum14 As Integer=1 'TIMS.sUtl_GetPYNum14(Me)
    'http://192.168.0.76:8080/ReportServer3/report?RptID=CP_01_006_C2&OCID=156430&OrgName=%e5%8b%9e%e5%8b%95%e5%8a%9b%e7%99%bc%e5%b1%95%e7%bd%b2&UserID=snoopy

    'e.Item.Cells/DataGrid1.Columns
    Const cst_dgf序號 As Integer = 0
    Const cst_dgf訓練機構 As Integer = 1
    Const cst_dgf班別名稱 As Integer = 2
    Const cst_dgf開訓日期 As Integer = 3
    Const cst_dgf結訓日期 As Integer = 4
    Const cst_dgf訪查日期 As Integer = 5
    Const cst_dgf抽訪結果 As Integer = 6
    Const cst_dgf原因及追 As Integer = 7
    Const cst_dgf功能 As Integer = 8

    Dim prtFilename As String = "" '列印表件名稱

    ' cst_printFN1 "CP_01_006_1_b" '(Old)-空白表單
    ' cst_printFN2 "CP_01_006_b2" '2017-空白表單
    ' cst_printFN3 "CP_01_006_C" '2018'空白表單
    ' cst_printFN4 "CP_01_006_C2" '2019'空白表單 CP_01_006_C2*.jrxml
    Const cst_printFN4 As String = "CP_01_006_C3" '2024'空白表單 CP_01_006_C3*.jrxml 

    Const cst_CP_01_006_add_aspx_new As String = "CP_01_006_add9.aspx?ID=" 'NEW 不預告實地抽訪紀錄表 - 班級
    Const cst_CP_01_006_add9t_aspx_edt As String = "CP_01_006_add9t.aspx?ID=" 'EDIT '抽訪學員紀錄-AddStd

    'State
    Const cst_State_Add_新增 As String = "Add"
    Const cst_State_View_檢視 As String = "View" '查詢
    Const cst_State_Edit_修改 As String = "Edit"

    Const cst_xx26 As String = "%26"
    Dim iPYNum As Integer = 1 'iPYNum=TIMS.sUtl_GetPYNum(Me) '1:2017前 2:2017 3:2018
    'Dim au As New cAUTH
    Dim objconn As SqlConnection
    Dim flag_ROC As Boolean = TIMS.CHK_REPLACE2ROC_YEARS()

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
        Call TIMS.OpenDbConn(objconn) '開啟連線
        '檢查Session是否存在 End
        PageControler1.PageDataGrid = DataGrid1
        iPYNum = TIMS.sUtl_GetPYNum(Me)

        If Not IsPostBack Then
            msg.Text = ""
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID

            end_date.Text = If(flag_ROC, TIMS.Cdate17(Now.Date), TIMS.Cdate3(Now.Date)) '民國日期，by:20181018

            DataGridTable.Visible = False
            Button1.Attributes("onclick") = "javascript:return search()"

            Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
            Button5.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))
        End If

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

        If Not IsPostBack Then
            If Session("SearchStr") IsNot Nothing Then
                Dim MyValue As String
                MyValue = TIMS.GetMyValue(Session("SearchStr"), "prgid")
                If MyValue = "CP_01_006" Then
                    Dim str_SearchStr As String = Session("SearchStr")
                    MyValue = TIMS.GetMyValue(str_SearchStr, "center")
                    center.Text = Replace(MyValue, cst_xx26, "&")
                    RIDValue.Value = TIMS.GetMyValue(str_SearchStr, "RIDValue")

                    MyValue = TIMS.GetMyValue(str_SearchStr, "TMID1")
                    TMID1.Text = Replace(MyValue, cst_xx26, "&")
                    MyValue = TIMS.GetMyValue(str_SearchStr, "OCID1")
                    OCID1.Text = Replace(MyValue, cst_xx26, "&")
                    TMIDValue1.Value = TIMS.GetMyValue(str_SearchStr, "TMIDValue1")
                    OCIDValue1.Value = TIMS.GetMyValue(str_SearchStr, "OCIDValue1")
                    start_date.Text = TIMS.GetMyValue(str_SearchStr, "start_date")
                    end_date.Text = TIMS.GetMyValue(str_SearchStr, "end_date")

                    MyValue = TIMS.GetMyValue(str_SearchStr, "VisitItem")
                    If MyValue <> "" Then Common.SetListItem(VisitItem, MyValue)

                    'PageControler1.PageIndex=TIMS.GetMyValue(str_SearchStr , "PageIndex")
                    MyValue = TIMS.GetMyValue(str_SearchStr, "PageIndex")
                    If MyValue <> "" AndAlso IsNumeric(MyValue) Then
                        MyValue = CInt(MyValue)
                        PageControler1.PageIndex = MyValue
                    End If

                    MyValue = TIMS.GetMyValue(str_SearchStr, "Button1")
                    If MyValue = "true" Then
                        Button1_Click(sender, e)
                    End If
                End If
                Session("SearchStr") = Nothing
                '若只有管理一個班級，自動協助帶出班級--by andy 2009-02-25
                TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
            End If

            Button2.Attributes("onclick") = "javascript:return CheckAdd()"
            Button10.Attributes("onclick") = "javascript:return CheckAdd()"
        End If

    End Sub

    Sub sSearch1()

        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then
            Common.MessageBox(Me, "未選擇有效單位，停止查詢!!")
            Exit Sub '查無輸入單位。
        End If

        Dim parms As New Hashtable() From {{"RID", RIDValue.Value}}

        Dim sql As String = ""
        sql &= " SELECT b.OCID" & vbCrLf
        sql &= " ,a.SeqNo, a.RID, a.LItem1, a.LItem2, a.LItem2_1, a.LItem2_2" & vbCrLf
        sql &= " ,a.LItem2_2_Note" & vbCrLf
        sql &= " ,a.LItem2_3_Note" & vbCrLf
        sql &= " ,b.STDATE ,b.FTDATE ,a.APPLYDATE" & vbCrLf
        sql &= " ,dbo.FN_TW_DATE(b.STDATE) STDATE_TW" & vbCrLf
        sql &= " ,dbo.FN_TW_DATE(b.FTDATE) FTDATE_TW" & vbCrLf
        sql &= " ,dbo.FN_TW_DATE(a.APPLYDATE) APPLYDATE_TW" & vbCrLf
        sql &= " ,d.OrgName" & vbCrLf
        sql &= " ,d.Relship" & vbCrLf
        sql &= " ,c.ClassID" & vbCrLf
        sql &= " ,b.RID RIDValue" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(b.CLASSCNAME,b.CYCLTYPE) CLASSCNAME" & vbCrLf
        sql &= " FROM dbo.CLASS_CLASSINFO b" & vbCrLf
        sql &= " JOIN dbo.VIEW_RIDNAME d on d.rid =b.rid" & vbCrLf
        sql &= " JOIN dbo.ID_CLASS c on b.clsid=c.clsid" & vbCrLf
        sql &= " JOIN dbo.ID_PLAN ip on ip.PlanID=b.PlanID" & vbCrLf
        sql &= " LEFT JOIN dbo.CLASS_UNEXPECTVISITOR a on a.OCID=b.OCID" & vbCrLf
        'AND a.OCID IS NULL AND d.RID='B5322' AND ip.PlanID='4494' AND b.OCID='112751' AND d.Relship like '" & Relship & "'
        sql &= " WHERE b.NOTOPEN='N' AND d.RID=@RID" & vbCrLf

        Select Case VisitItem.SelectedIndex
            Case 0
                '未抽訪
                sql &= " AND a.OCID IS NULL" & vbCrLf
            Case 1
                '已抽訪
                sql &= " AND a.OCID IS NOT NULL" & vbCrLf
            Case 2
                '全部
        End Select
        Select Case VisitItem.SelectedIndex
            Case 0 '未抽訪
            Case 2 '全部
            Case Else
                If start_date.Text <> "" Then
                    If flag_ROC Then
                        sql &= " AND a.ApplyDate >= @ApplyDate1" & vbCrLf  'edit，by:20181018
                        parms.Add("ApplyDate1", TIMS.Cdate18(start_date.Text))  'edit，by:20181018
                    Else
                        sql &= " AND a.ApplyDate >= " & TIMS.To_date(start_date.Text) & vbCrLf  'edit，by:20181018
                    End If
                End If
                If end_date.Text <> "" Then
                    If flag_ROC Then
                        sql &= " AND a.ApplyDate <= @ApplyDate2" & vbCrLf  'edit，by:20181018
                        parms.Add("ApplyDate2", TIMS.Cdate18(end_date.Text))  'edit，by:20181018
                    Else
                        sql &= " AND a.ApplyDate <= " & TIMS.To_date(end_date.Text) & vbCrLf  'edit，by:20181018
                    End If
                End If
        End Select

        If sm.UserInfo.RID = "A" Then
            sql &= " AND ip.TPlanID=@TPlanID" & vbCrLf
            sql &= " AND ip.Years=@Years" & vbCrLf

            parms.Add("TPlanID", sm.UserInfo.TPlanID)
            parms.Add("Years", sm.UserInfo.Years)
        Else
            sql &= " AND ip.PlanID=@PlanID" & vbCrLf
            parms.Add("PlanID", sm.UserInfo.PlanID)
        End If
        If OCIDValue1.Value <> "" Then
            sql &= " AND b.OCID=@OCID "
            parms.Add("OCID", OCIDValue1.Value)
        End If
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)

        labmsg1.Text = ""
        msg.Text = "查無資料!"
        DataGridTable.Visible = False

        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return

        labmsg1.Text = Get_RtnMsgValue(RIDValue.Value)
        msg.Text = ""
        DataGridTable.Visible = True
        If VisitItem.SelectedIndex = 0 Then
            '未抽訪
            DataGrid1.Columns(cst_dgf訪查日期).Visible = False
            DataGrid1.Columns(cst_dgf抽訪結果).Visible = False
            DataGrid1.Columns(cst_dgf原因及追).Visible = False
            DataGrid1.Columns(cst_dgf功能).Visible = True
        Else
            '非'未抽訪
            DataGrid1.Columns(cst_dgf訪查日期).Visible = True
            DataGrid1.Columns(cst_dgf抽訪結果).Visible = True
            DataGrid1.Columns(cst_dgf原因及追).Visible = True
            DataGrid1.Columns(cst_dgf功能).Visible = True
        End If

        PageControler1.Sort = "RIDValue,ClassID"
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()

    End Sub

    ''' <summary>OJT-24031304 判斷單位是否曾經有處分過x訪視異常之提示訊息</summary>
    ''' <param name="RID"></param>
    ''' <returns></returns>
    Function Get_RtnMsgValue(ByRef RID As String) As String
        Dim rst As String = ""
        If RID = "" Then Return rst
        Const cst_alertMsg1 As String = "*注意：該單位曾遭處分紀錄或登錄訪視異常，請分署加強訪視次數!"
        Const cst_alertMsg2 As String = "*注意：該單位曾遭處分紀錄或登錄訪視異常，請分署加強訪視次數!!"
        '系統：在職系統 '計畫：產業人才投資方案 '功能：首頁>>學員動態管理>>教務管理>>不預告實地抽訪紀錄表
        '需求： 'PS：這個需求是副分署長要求的，要讓分署知道哪些單位要多去訪視
        '為協助分署強化「精準訪視」，分署於不預告實地抽訪紀錄表功能選擇【機構】時，
        '即該班於「綜合查詢統計表」之【累計不預告實地抽訪異常次數】 > 0        '(判斷的是下面勾選2這個)
        '若有上述兩種情況任一種，系統於查詢結果上方顯示提示訊息： '「注意：該單位曾遭處分紀錄或登錄訪視異常，請分署加強訪視次數」

        '若單位曾經發生過以下任一狀況

        '1.於「訓練單位處分」功能，產投計畫，任一年(不管是否還在處分期間)，曾經有處分停權資料
        Dim hPMS1 As New Hashtable From {{"RID", RID}}
        Dim sSql1 As String = ""
        sSql1 &= " SELECT TOP 3 b.OBSN,b.COMIDNO,b.OBSDATE,b.OBYEARS,b.MODIFYACCT,b.MODIFYDATE FROM ORG_BLACKLIST b" & vbCrLf
        sSql1 &= " WHERE b.COMIDNO IN (SELECT COMIDNO FROM VIEW_RIDNAME WHERE RID=@RID)" & vbCrLf
        Dim dt1 As DataTable = DbAccess.GetDataTable(sSql1, objconn, hPMS1)
        If dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0 Then Return cst_alertMsg1

        '2.於「不預告實地抽訪紀錄表」功能，任一年任一班之抽訪結果曾經被註記「不正常」時，
        Dim hPMS2 As New Hashtable From {{"RID", RID}}
        Dim sSql2 As String = ""
        sSql2 &= " SELECT TOP 3 U.OCID,U.SEQNO,U.APPLYDATE,u.LItem1,u.VISITWAY,u.MODIFYACCT,u.MODIFYDATE" & vbCrLf
        sSql2 &= " FROM CLASS_CLASSINFO cc" & vbCrLf
        sSql2 &= " JOIN CLASS_UNEXPECTVISITOR U on U.OCID =cc.OCID" & vbCrLf
        sSql2 &= " WHERE U.LItem1='2' AND ISNULL(u.VISITWAY,'1')='1'" & vbCrLf
        sSql2 &= " AND cc.COMIDNO IN (SELECT COMIDNO FROM VIEW_RIDNAME WHERE RID=@RID)" & vbCrLf
        Dim dt2 As DataTable = DbAccess.GetDataTable(sSql2, objconn, hPMS2)
        If dt2 IsNot Nothing AndAlso dt2.Rows.Count > 0 Then Return cst_alertMsg2

        Return rst
    End Function

    '查詢
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        sSearch1()
    End Sub

    '新增
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Call KeepSearchStr(OCIDValue1.Value)
        Call TIMS.CloseDbConn(objconn)
        Dim rqMID As String = TIMS.Get_MRqID(Me)
        'Dim sUrl1 As String="CP_01_006_add.aspx?ID=" & rqMID
        'If iPYNum >= 3 Then sUrl1="CP_01_006_add8.aspx?ID=" & rqMID
        'If TIMS.GetReportQueryPath=TIMS.cst_Report_TEST Then sUrl1="CP_01_006_add9.aspx?ID=" & rqMID
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        Dim sUrl1 As String = cst_CP_01_006_add_aspx_new & rqMID
        If (OCIDValue1.Value <> "") Then sUrl1 &= "&OCID=" & OCIDValue1.Value
        sUrl1 &= "&State=" & cst_State_Add_新增
        TIMS.Utl_Redirect1(Me, sUrl1)
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e.CommandArgument = "" Then Exit Sub
        Dim sCmdArg As String = e.CommandArgument
        Dim gOCID As String = TIMS.GetMyValue(sCmdArg, "OCID")
        Dim gSeqNo As String = TIMS.GetMyValue(sCmdArg, "SeqNo")
        Dim gState As String = TIMS.GetMyValue(sCmdArg, "State")

        Select Case e.CommandName
            Case "edit", "view"
                Call KeepSearchStr(gOCID)
                Dim rqMID As String = TIMS.Get_MRqID(Me)
                'Dim sUrl1 As String="CP_01_006_add.aspx?ID=" & rqMID
                'If iPYNum >= 3 Then sUrl1="CP_01_006_add8.aspx?ID=" & rqMID
                'If TIMS.GetReportQueryPath=TIMS.cst_Report_TEST Then sUrl1="CP_01_006_add9.aspx?ID=" & rqMID
                Dim sUrl1 As String = cst_CP_01_006_add_aspx_new & rqMID
                sUrl1 &= "&OCID=" & gOCID
                sUrl1 &= "&SeqNo=" & gSeqNo
                sUrl1 &= "&State=" & gState
                TIMS.Utl_Redirect(Me, objconn, sUrl1)
            Case "Add"
                Call KeepSearchStr(gOCID)
                Dim rqMID As String = TIMS.Get_MRqID(Me)
                'Dim sUrl1 As String="CP_01_006_add.aspx?ID=" & rqMID
                'If iPYNum >= 3 Then sUrl1="CP_01_006_add8.aspx?ID=" & rqMID
                'If TIMS.GetReportQueryPath=TIMS.cst_Report_TEST Then sUrl1="CP_01_006_add9.aspx?ID=" & rqMID
                Dim sUrl1 As String = cst_CP_01_006_add_aspx_new & rqMID
                sUrl1 &= "&OCID=" & gOCID
                sUrl1 &= "&SeqNo=" & gSeqNo
                sUrl1 &= "&State=" & gState

                TIMS.Utl_Redirect(Me, objconn, sUrl1)
            Case "del"
                Dim drCC As DataRow = TIMS.GetOCIDDate(gOCID, objconn)
                If drCC Is Nothing Then
                    Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
                    Exit Sub
                End If

                Dim pms_s1 As New Hashtable From {{"OCID", gOCID}, {"SeqNo", gSeqNo}}
                Dim sql As String = ""
                sql &= " SELECT 'X' FROM CLASS_UNEXPECTVISITOR WHERE OCID=@OCID AND SeqNo=@SeqNo"
                Dim drCU As DataRow = DbAccess.GetOneRow(sql, objconn, pms_s1)
                If drCU Is Nothing Then
                    Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
                    Exit Sub
                End If

                'Dim sql As String=""
                Dim pms_d1 As New Hashtable From {{"OCID", gOCID}, {"SeqNo", gSeqNo}}
                Dim sql_D As String = ""
                sql_D &= " DELETE CLASS_UNEXPECTVISITOR"
                sql_D &= " WHERE OCID=@OCID AND SeqNo=@SeqNo"
                DbAccess.ExecuteNonQuery(sql_D, objconn, pms_d1)

                Dim rqMID As String = TIMS.Get_MRqID(Me)
                Dim sUrl1 As String = "CP/01/CP_01_006.aspx?ID=" & rqMID
                TIMS.BlockAlert(Me, "刪除成功!!", sUrl1)

            Case "AddStd"
                Call KeepSearchStr(gOCID)
                Dim rqMID As String = TIMS.Get_MRqID(Me)
                'Dim sUrl1 As String="CP_01_006_add.aspx?ID=" & rqMID
                'If iPYNum >= 3 Then sUrl1="CP_01_006_add8.aspx?ID=" & rqMID
                'If TIMS.GetReportQueryPath=TIMS.cst_Report_TEST Then sUrl1="CP_01_006_add9.aspx?ID=" & rqMID
                Dim sUrl1 As String = cst_CP_01_006_add9t_aspx_edt & rqMID
                sUrl1 &= "&OCID=" & gOCID
                sUrl1 &= "&SeqNo=" & gSeqNo
                'sUrl1 &= "&State=" & gState
                TIMS.Utl_Redirect(Me, objconn, sUrl1)

        End Select
    End Sub

    '逐筆列印
    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem, ListItemType.EditItem
                Dim drv As DataRowView = e.Item.DataItem
                e.Item.Cells(cst_dgf序號).Text = TIMS.Get_DGSeqNo(sender, e)

                Dim ParentName As String = TIMS.Get_ParentRID(drv("Relship"), objconn)
                If ParentName <> "" Then e.Item.Cells(cst_dgf訓練機構).Text = ParentName & "-" & drv("OrgName")

                Dim mybtnEdit As Button = e.Item.FindControl("Button3e") '修改'edit
                Dim mybtnView As Button = e.Item.FindControl("Button3v") '查詢'view
                Dim mybtnDel As Button = e.Item.FindControl("Button4") '刪除'del
                Dim BtnAddStd As Button = e.Item.FindControl("BtnAddStd") '抽訪學員紀錄-AddStd
                Dim mybtnPrint As Button = e.Item.FindControl("Button8") '列印'prt1
                Dim myBtnAdd As Button = e.Item.FindControl("Button7") '新增'Add
                Dim mybtnPrtEmpty As Button = e.Item.FindControl("Button11") '列印空白表單'prt2
                Dim Reason As HtmlControls.HtmlTextArea = e.Item.FindControl("Reason") '現場處理說明的其他 'Reason.Disabled=True 
                'Dim LabApplyDate As Label=e.Item.FindControl("LabApplyDate")
                'If Convert.ToString(drv("ApplyDate")) <> "" Then
                '    LabApplyDate.Text=If(flag_ROC, TIMS.cdate17(drv("ApplyDate")), TIMS.cdate3(drv("ApplyDate")))
                'End If
                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "OCID", Convert.ToString(drv("OCID")))
                TIMS.SetMyValue(sCmdArg, "SeqNo", Convert.ToString(drv("SeqNo")))

                mybtnEdit.Visible = False
                mybtnView.Visible = False
                mybtnDel.Visible = False
                BtnAddStd.Visible = False '抽訪學員紀錄
                mybtnPrint.Visible = False
                myBtnAdd.Visible = True
                mybtnPrtEmpty.Visible = True

                If VisitItem.SelectedIndex = 0 Then
                    '未抽訪只需顯示新增、列印空白
                    myBtnAdd.Enabled = True
                    myBtnAdd.CommandArgument = sCmdArg & "&State=" & cst_State_Add_新增

                    '列印-列印空白表單+OrgName
                    'prtFilename=cst_printFN1 '"CP_01_006_1_b"
                    'If iPYNum=2 Then prtFilename=cst_printFN2 '"CP_01_006_b2"
                    'If iPYNum >= 3 Then prtFilename=cst_printFN3 '"CP_01_006_C"
                    'If TIMS.GetReportQueryPath()=TIMS.cst_Report_TEST Then prtFilename=cst_printFN4 '"CP_01_006_C2"
                    mybtnPrtEmpty.Enabled = True
                    prtFilename = cst_printFN4 '"CP_01_006_C2"
                    Dim s_val As String = String.Format("OCID={0}&OrgName={1}", drv("OCID"), sm.UserInfo.OrgName)
                    mybtnPrtEmpty.Attributes("onclick") = ReportQuery.ReportScript(Me, prtFilename, s_val)
                Else
                    myBtnAdd.Visible = False '不顯示新增
                    mybtnEdit.Visible = True
                    mybtnView.Visible = False
                    If Not IsDBNull(drv("RID")) Then
                        If drv("RID") <> sm.UserInfo.RID Then
                            'mybtnEdit.Text="檢視"
                            mybtnEdit.Visible = False
                            mybtnView.Visible = True
                            mybtnDel.Visible = False
                        End If
                    End If

                    If Not IsDBNull(drv("LItem1")) Then
                        If drv("LItem1").ToString = "1" Then e.Item.Cells(cst_dgf抽訪結果).Text = "正常"
                    End If

                    If Not IsDBNull(drv("LItem2")) Then
                        If drv("LItem2").ToString = "1" Then e.Item.Cells(cst_dgf抽訪結果).Text = "不正常"
                    End If

                    'Dim Reason As HtmlControls.HtmlTextArea=e.Item.FindControl("Reason")
                    'Reason.Disabled=True
                    '現場處理說明的其他
                    TIMS.Tooltip(Reason, "現場處理說明的其他與補充說明..")
                    Reason.Value = ""
                    If Not IsDBNull(drv("LItem2_2_Note")) _
                        AndAlso Convert.ToString(drv("LItem2_2_Note")) <> "" Then
                        Reason.Value = drv("LItem2_2_Note").ToString
                    End If
                    If Not IsDBNull(drv("LItem2_3_Note")) _
                       AndAlso Convert.ToString(drv("LItem2_3_Note")) <> "" Then
                        If Reason.Value <> "" Then Reason.Value += ","
                        Reason.Value += drv("LItem2_3_Note").ToString
                    End If
                    If Not IsDBNull(drv("LItem2_1")) Then
                        If drv("LItem2_1").ToString = "1" Then
                            If Reason.Value <> "" Then Reason.Value += ","
                            Reason.Value += "[學員資料有誤或填寫錯誤]"
                        End If
                    End If

                    myBtnAdd.Visible = False
                    'mybtnEdit.Visible=True
                    'mybtnView.Visible=False
                    mybtnDel.Visible = True
                    BtnAddStd.Visible = True
                    mybtnPrint.Visible = True
                    If Convert.ToString(drv("ApplyDate")) = "" Then 'cst_dgf訪查日期e.Item.Cells(3).Text
                        e.Item.Cells(cst_dgf抽訪結果).Text = "未抽訪"
                        myBtnAdd.Visible = True
                        mybtnEdit.Visible = False
                        mybtnView.Visible = False
                        mybtnDel.Visible = False
                        BtnAddStd.Visible = False
                        mybtnPrint.Visible = False
                    End If

                    mybtnView.CommandArgument = sCmdArg & "&State=" & cst_State_View_檢視
                    mybtnEdit.CommandArgument = sCmdArg & "&State=" & cst_State_Edit_修改
                    mybtnDel.Attributes("onclick") = TIMS.cst_confirm_delmsg1
                    mybtnDel.CommandArgument = sCmdArg
                    myBtnAdd.CommandArgument = sCmdArg & "&State=" & cst_State_Add_新增
                    BtnAddStd.CommandArgument = sCmdArg '抽訪學員紀錄
                    'ReportQuery(列印)
                    'prtFilename=cst_printFN1 '"CP_01_006_1_b"
                    'If iPYNum=2 Then prtFilename=cst_printFN2 '"CP_01_006_b2"
                    'If iPYNum >= 3 Then prtFilename=cst_printFN3 '"CP_01_006_C"
                    'If TIMS.GetReportQueryPath()=TIMS.cst_Report_TEST Then prtFilename=cst_printFN4 '"CP_01_006_C2"
                    prtFilename = cst_printFN4 '"CP_01_006_C2"
                    mybtnPrint.Attributes("onclick") = ReportQuery.ReportScript(Me, prtFilename, "OCID=" & drv("OCID") & "&SeqNo=" & drv("SeqNo"))

                    'ReportQuery(列印)-列印空白表單+OrgName
                    'prtFilename=cst_printFN1 '"CP_01_006_1_b"
                    'If iPYNum=2 Then prtFilename=cst_printFN2 '"CP_01_006_b2"
                    'If iPYNum >= 3 Then prtFilename=cst_printFN3 '"CP_01_006_C"
                    'If TIMS.GetReportQueryPath()=TIMS.cst_Report_TEST Then prtFilename=cst_printFN4 '"CP_01_006_C2"
                    prtFilename = cst_printFN4 '"CP_01_006_C2"
                    mybtnPrtEmpty.Attributes("onclick") = ReportQuery.ReportScript(Me, prtFilename, "OCID=" & drv("OCID") & "&OrgName=" & sm.UserInfo.OrgName & "")
                End If

                Select Case sm.UserInfo.RID
                    Case "A"
                    Case Else
                        If Left(Convert.ToString(drv("RIDValue")), 1) <> sm.UserInfo.RID Then
                            myBtnAdd.Enabled = False
                            TIMS.Tooltip(myBtnAdd, "業務資訊不相同，不提供新增")
                            mybtnEdit.Enabled = False
                            TIMS.Tooltip(mybtnEdit, "業務資訊不相同，不提供修改")
                            mybtnView.Enabled = False
                            TIMS.Tooltip(mybtnView, "業務資訊不相同，不提供檢視")
                            mybtnDel.Enabled = False
                            TIMS.Tooltip(mybtnDel, "業務資訊不相同，不提供刪除")
                        End If

                End Select

                If Convert.ToString(drv("ApplyDate")) <> "" Then
                    Dim flgROLEIDx0xLIDx0 As Boolean = TIMS.IsSuperUser(sm, 1)
                    If flgROLEIDx0xLIDx0 Then
                        myBtnAdd.Visible = True
                        mybtnEdit.Visible = True
                        myBtnAdd.Enabled = True
                        mybtnEdit.Enabled = True
                    End If
                End If
        End Select

    End Sub

    ''' <summary>
    ''' SESSION SAVE
    ''' </summary>
    Sub KeepSearchStr(ByVal vOCIDVal As String)
        Session("SearchStr") = Nothing
        Dim drCC As DataRow = TIMS.GetOCIDDate(vOCIDVal, objconn)
        'If drCC Is Nothing Then Return

        '文字類可能有 "&" 符號，試著轉換 "%26" 字眼
        Dim str_SearchStr As String = ""
        str_SearchStr = "prgid=CP_01_006"
        str_SearchStr &= "&RIDValue=" & TIMS.ClearSQM(RIDValue.Value)
        str_SearchStr &= "&center=" & Replace(center.Text, "&", cst_xx26)
        str_SearchStr &= "&TMID1=" & Replace(TMID1.Text, "&", cst_xx26)
        str_SearchStr &= "&OCID1=" & Replace(OCID1.Text, "&", cst_xx26)
        str_SearchStr &= "&TMIDValue1=" & TIMS.ClearSQM(TMIDValue1.Value)
        str_SearchStr &= "&OCIDValue1=" & TIMS.ClearSQM(OCIDValue1.Value) 'OCID
        If drCC IsNot Nothing Then
            str_SearchStr &= String.Format("&SFDATE_TW={0}~{1}", TIMS.Cdate17(drCC("STDATE")), TIMS.Cdate17(drCC("FTDATE")))
        End If

        str_SearchStr &= "&start_date=" & TIMS.ClearSQM(start_date.Text)
        str_SearchStr &= "&end_date=" & TIMS.ClearSQM(end_date.Text)
        str_SearchStr &= "&VisitItem=" & TIMS.GetListValue(VisitItem)
        str_SearchStr &= "&PageIndex=" & DataGrid1.CurrentPageIndex + 1
        str_SearchStr &= If(DataGrid1.Visible, "&Button1=true", "&Button1=false")
        Session("SearchStr") = str_SearchStr
    End Sub

    '列印空白表單
    Private Sub Button10_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button10.Click
        Dim myValue As String = ""
        myValue = "OCID=" & OCIDValue1.Value
        myValue += "&OrgName=" & Server.UrlEncode(sm.UserInfo.OrgName)
        'prtFilename=cst_printFN1 '"CP_01_006_1_b"
        'If iPYNum=2 Then prtFilename=cst_printFN2 '"CP_01_006_b2"
        'If iPYNum >= 3 Then prtFilename=cst_printFN3 '"CP_01_006_C"
        'If TIMS.GetReportQueryPath()=TIMS.cst_Report_TEST Then prtFilename=cst_printFN4 '"CP_01_006_C2"
        prtFilename = cst_printFN4 '"CP_01_006_C2"
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, prtFilename, myValue)
    End Sub

End Class
