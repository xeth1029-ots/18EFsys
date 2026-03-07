Partial Class SD_01_013
    Inherits AuthBasePage

    'Dim CPdt As DataTable
    'Dim ProcessType As String
    'Dim RelshipTable As DataTable
    Dim flagExportExcel As Boolean = False '匯出Excel

    'Cells 'DG_ClassInfo 非產投 (TIMS)
    'Const Cst_管控單位 As Integer=1
    Const Cst_訓練機構 As Integer = 1
    'Const Cst_班別代碼 As Integer=2
    Const Cst_開結訓日 As Integer = 3
    Const Cst_colspan As String = "8" '依dataGrid資料欄
    Const cst_excelFN1 As String = "學員甄試人數統計表.xls"

    Dim iClassCnt As Integer = 0
    Dim iStudCnt1 As Integer = 0
    Dim iStudCnt2 As Integer = 0

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), titlelab1, titlelab2)
        'TIMS.TestDbConn(Me, objConn, True)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        '分頁設定 Start
        'msg2_28.Visible=False
        DG_Classinfo.Visible = False
        'DG_Classinfo2.Visible=False
        '非產投(TIMS)
        DG_Classinfo.Visible = True
        PageControler1.PageDataGrid = DG_Classinfo
        btu_sel.Attributes("onclick") = "openTrain(document.getElementById('trainValue').value);"
        '分頁設定 End

        'ProcessType=Request("ProcessType")
        If sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1 Then
            org.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg.aspx');"
        Else
            org.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg1.aspx');"
        End If
        TIMS.ShowHistoryRID(Me, historyrid, "HistoryList2", "RIDValue", "center")
        If historyrid.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

        'Dim sql As String=""
        'sql=""
        'sql += " SELECT a.RID,a.Relship,b.OrgName "
        'sql += " FROM Auth_Relship a "
        'sql += " JOIN Org_OrgInfo b ON a.OrgID=b.OrgID "
        'RelshipTable=DbAccess.GetDataTable(sql, objconn)

        If Not Page.IsPostBack Then
            btnExport1.Visible = False '預設要查詢一次再顯示匯出鍵
            msg.Text = ""
            table4.Visible = False
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
        End If
    End Sub

    '查詢 '非產投查詢 (現場?)
    Sub Search1()
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub

        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID

        '依sm.UserInfo.PlanID取得PlanKind
        Dim sPlanKind As String = TIMS.Get_PlanKind(Me, objconn)

        Dim bRWClassFlag As Boolean = False
        Select Case sm.UserInfo.RID
            Case "A"
            Case Else
                '非署(局)才有此限制
                If sPlanKind = "1" AndAlso sm.UserInfo.RoleID > 1 Then bRWClassFlag = True '自辦者只能列出賦予給此帳號的班級
        End Select

        Dim parms As Hashtable = New Hashtable()
        Dim sql As String = ""
        sql &= " WITH WC1 AS ( SELECT cc.Years" & vbCrLf
        sql &= "  ,cc.CyclType,cc.OCLASSID,cc.OCID,cc.PlanID,cc.ComIDNO,cc.SeqNO" & vbCrLf
        sql &= "  ,cc.ClassCName + '(第' + cc.CyclType + '期)' ClassCName ,cc.TPropertyID" & vbCrLf
        sql &= "  ,CONVERT(varchar, cc.STDate, 111) STDate" & vbCrLf
        sql &= "  ,CONVERT(varchar, cc.FTDate, 111) FTDate" & vbCrLf
        sql &= "  ,cc.RID,cc.TNum,cc.ClassID,cc.OrgName,cc.TrainName" & vbCrLf
        sql &= "  FROM VIEW2 cc " & vbCrLf
        'sql &= "  JOIN VIEW_TRAINTYPE tt ON tt.tmid=cc.tmid " & vbCrLf
        sql &= "  WHERE 1=1 " & vbCrLf
        If sm.UserInfo.RID = "A" Then
            sql &= " AND cc.TPlanID=@TPlanID AND cc.Years=@Years " & vbCrLf
        Else
            sql &= " AND cc.PlanID=@PlanID " & vbCrLf
        End If
        sql &= " AND cc.RID=@RID " & vbCrLf

        If trainValue.Value <> "" Then
            sql &= " AND cc.TMID=@TMID " & vbCrLf
        End If
        If tb_classname.Text <> "" Then
            sql &= " AND cc.ClassCName LIKE @ClassCName " & vbCrLf
        End If
        If txtCJOB_NAME.Text <> "" Then   '通俗職類
            sql &= " AND cc.CJOB_UNKEY=@CJOB_UNKEY " & vbCrLf
        End If
        If start_date.Text <> "" Then
            'sql &= " and cc.STDate >= " & TIMS.to_date(Me.start_date.Text) & vbCrLf
            sql &= " AND cc.STDate >= @STDate1 " & vbCrLf
        End If
        If end_date.Text <> "" Then
            'sql &= " and cc.STDate <= @STDate2" & TIMS.to_date(Me.end_date.Text) & vbCrLf
            sql &= " AND cc.STDate <= @STDate2 " & vbCrLf
        End If
        If bRWClassFlag Then sql &= " AND EXISTS (SELECT 'x' FROM Auth_AccRWClass x WHERE x.OCID=cc.OCID AND x.Account=@Account)" & vbCrLf
        sql &= " ) " & vbCrLf

        'sql &= " ,WS1 AS (" & vbCrLf
        'sql &= "   SELECT cc.OCID ,COUNT(1) StEnTeNum2 " & vbCrLf
        'sql &= "   FROM WC1 cc " & vbCrLf
        'sql &= "   JOIN STUD_ENTERTYPE b ON b.OCID1=cc.OCID " & vbCrLf
        'sql &= "   JOIN STUD_ENTERTEMP a ON a.setid=b.setid " & vbCrLf
        'sql &= "   WHERE 1=1 " & vbCrLf
        'sql &= "   AND (b.totalresult >= 0) " & vbCrLf
        'sql &= " GROUP BY cc.OCID " & vbCrLf
        'sql &= " ) " & vbCrLf

        sql &= " SELECT cc.* " & vbCrLf
        '甄試人數查詢 '有任1分數(筆試、口試、總分)，大於等於0，即為甄試名單
        'sql &= " AND (b.writeresult >=0 OR b.oralresult >=0 OR b.totalresult >= 0) " & vbCrLf
        'sql &= " AND (b.writeresult >=0 OR b.oralresult >=0) " & vbCrLf
        '若總分，大於等於0，即為甄試名單'https://jira.turbotech.com.tw/browse/TIMSC-3
        sql &= " ,dbo.FN_STUDCOUNT(cc.OCID,1) StEnTeNum2 " & vbCrLf
        sql &= " FROM WC1 cc " & vbCrLf
        'sql &= " JOIN WS1 ss ON ss.OCID=cc.OCID " & vbCrLf
        Dim sCmd As New SqlCommand(sql, objconn)
        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable

        Try
            With sCmd
                .Parameters.Clear()
                If sm.UserInfo.RID = "A" Then
                    'sql &= " AND cc.TPlanID='" & sm.UserInfo.TPlanID & "'" & vbCrLf
                    'sql &= " AND cc.Years='" & sm.UserInfo.Years & "'" & vbCrLf
                    .Parameters.Add("TPlanID", SqlDbType.VarChar).Value = sm.UserInfo.TPlanID
                    .Parameters.Add("Years", SqlDbType.Int).Value = sm.UserInfo.Years
                Else
                    'sql &= " AND cc.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
                    .Parameters.Add("PlanID", SqlDbType.Int).Value = Convert.ToInt32(sm.UserInfo.PlanID)
                End If
                'sql &= "and cc.RID='" & RIDValue.Value & "' " & vbCrLf
                .Parameters.Add("RID", SqlDbType.VarChar).Value = RIDValue.Value

                If trainValue.Value <> "" Then
                    'sql &= "and cc.TMID='" & trainValue.Value & "' " & vbCrLf
                    .Parameters.Add("TMID", SqlDbType.Int).Value = Convert.ToInt32(Me.trainValue.Value)
                End If
                If tb_classname.Text <> "" Then
                    'sql &= "and  cc.ClassCName like '%" & tb_classname.Text & "%'" & vbCrLf
                    .Parameters.Add("ClassCName", SqlDbType.NVarChar).Value = "%" & tb_classname.Text & "%"
                End If
                If txtCJOB_NAME.Text <> "" Then   '通俗職類
                    'sql &= " and cc.CJOB_UNKEY=" & cjobValue.Value & "" & vbCrLf
                    .Parameters.Add("CJOB_UNKEY", SqlDbType.Int).Value = Convert.ToInt32(cjobValue.Value)
                End If
                If start_date.Text <> "" Then
                    'sql &= " and cc.STDate >= " & TIMS.to_date(Me.start_date.Text) & vbCrLf
                    .Parameters.Add("STDate1", SqlDbType.VarChar).Value = start_date.Text
                End If
                If end_date.Text <> "" Then
                    'sql &= " and cc.STDate <= " & TIMS.to_date(Me.end_date.Text) & vbCrLf
                    .Parameters.Add("STDate2", SqlDbType.VarChar).Value = end_date.Text
                End If
                If bRWClassFlag Then
                    'sql &= " and EXISTS (select 'x' from Auth_AccRWClass x where x.OCID =cc.OCID AND x.Account='" & sm.UserInfo.UserID & "')" & vbCrLf
                    .Parameters.Add("Account", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                End If
                'dt.Load(.ExecuteReader())
                dt = DbAccess.GetDataTable(sCmd.CommandText, objconn, sCmd.Parameters)
            End With
        Catch ex As Exception
            TIMS.WriteTraceLog(Me, ex, ex.ToString)
            Common.MessageBox(Me, ex.ToString)
            Exit Sub
            'Common.RespWrite(Me, sqlstr)
            'Throw ex
        End Try

        iClassCnt = 0 'Dim iClassCnt As Integer=0
        iStudCnt1 = 0 'Dim iStudCnt1 As Integer=0
        iStudCnt2 = 0 'Dim iStudCnt2 As Integer=0

        '匯出Excel
        If Not flagExportExcel Then
            '一般輸出
            msg.Text = "查無資料!!"
            table4.Visible = False
            DG_Classinfo.Visible = False
            'DG_ClassInfo2.Visible=False '產投
            btnExport1.Visible = False
            PageControler1.Visible = False

            If dt.Rows.Count = 0 Then
                Common.MessageBox(Me, "查無資料")
                Exit Sub
            End If

            msg.Text = ""
            table4.Visible = True

            DG_Classinfo.Visible = True
            btnExport1.Visible = True
            PageControler1.Visible = True

            'PageControler1.SqlString=sqlstr_class
            'PageControler1.ControlerLoad()
            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
        Else
            If dt.Rows.Count = 0 Then
                Common.MessageBox(Me, "查無資料")
                Exit Sub
            End If

            '匯出Excel
            'If dt.Rows.Count > 0 Then
            iClassCnt = Convert.ToString(dt.Rows.Count)
            iStudCnt1 = 0
            iStudCnt2 = 0
            For Each dr As DataRow In dt.Rows
                If Convert.ToString(dr("TNum")) <> "" Then iStudCnt1 += Val(dr("TNum"))
                If Convert.ToString(dr("StEnTeNum2")) <> "" Then iStudCnt2 += Val(dr("StEnTeNum2"))
            Next
            'Me.ViewState("StudCnt1")=Convert.ToString(StudCnt1)
            'Me.ViewState("StudCnt2")=Convert.ToString(StudCnt2)

            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
        End If
    End Sub

    Private Sub DG_ClassInfo_ItemDataBound(ByVal sender As System.Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DG_Classinfo.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                If Not flagExportExcel Then
                    If Not ViewState("sort") Is Nothing Then
                        Dim i As Integer = 0
                        Select Case ViewState("sort")
                            Case "OrgName", "OrgName desc"
                                i = Cst_訓練機構
                        End Select
                        Dim img As New UI.WebControls.Image
                        img.ImageUrl = "../../images/SortDown.gif"
                        If ViewState("sort").ToString.IndexOf("desc") = -1 Then img.ImageUrl = "../../images/SortUp.gif"
                        e.Item.Cells(i).Controls.Add(img)
                    End If
                Else
                    'Dim i As Integer=Cst_訓練機構
                    e.Item.Cells(Cst_訓練機構).Text = "訓練機構"
                End If
            Case ListItemType.Item, ListItemType.AlternatingItem, ListItemType.EditItem
                Dim drv As DataRowView = e.Item.DataItem
                'Dim is_parent As String=""
                'Dim Result, Result1 As String
                'Dim myTableCell, myTableCell1 As TableCell
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e) '序號
                'If Convert.ToString(drv("OrgName2")) <> "" Then e.Item.Cells(Cst_管控單位).Text=Convert.ToString(drv("OrgName2"))
                'myTableCell=e.Item.Cells(Cst_班別代碼)

                'Result=""
                'Dim courName As String
                'If Len(drv("ClassID").ToString) < 4 Then
                '    courName=drv("Years").ToString & 0 & drv("ClassID").ToString & drv("CyclType").ToString
                'Else
                '    courName=drv("Years").ToString & drv("ClassID").ToString & drv("CyclType").ToString
                'End If
                'Result=courName
                'myTableCell.Text=Result

                'myTableCell1=e.Item.Cells(Cst_開結訓日)
                'Result1=""
                'Dim date_str As String
                'date_str=Convert.ToDateTime(drv("STDate")) & "<br>" & Convert.ToDateTime(drv("FTDate"))
                'Result1=date_str
                Dim sTMP2 As String = ""
                sTMP2 = CStr(drv("STDate")) & "<br>" & CStr(drv("FTDate"))
                e.Item.Cells(Cst_開結訓日).Text = sTMP2
        End Select
    End Sub

    Private Sub DG_ClassInfo_SortCommand(ByVal source As System.Object, ByVal e As System.Web.UI.WebControls.DataGridSortCommandEventArgs) Handles DG_Classinfo.SortCommand
        If Not flagExportExcel Then
            If e.SortExpression = ViewState("sort") Then
                ViewState("sort") = e.SortExpression & " desc"
            Else
                ViewState("sort") = e.SortExpression
            End If
            PageControler1.Sort = ViewState("sort")
            PageControler1.ChangeSort()
        Else
            ViewState("sort") = ""
            PageControler1.Sort = ViewState("sort")
            PageControler1.ChangeSort()
        End If
    End Sub

    '覆寫，不執行 MyBase.VerifyRenderingInServerForm 方法，解決執行 RenderControl 產生的錯誤   
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)
    End Sub

    '匯出 TIMS
    Sub Export1()
        flagExportExcel = True '匯出Excel

        DG_Classinfo.AllowPaging = False
        'DG_ClassInfo.Columns(8).Visible=False
        DG_Classinfo.EnableViewState = False  '把ViewState給關了

        Call Search1()

        Dim sFileName1 As String = "學員甄試人數統計表"

        '套CSS值
        Dim strSTYLE As String = ""
        strSTYLE &= ("<style>")
        strSTYLE &= ("td{mso-number-format:""\@"";}")
        strSTYLE &= ("</style>")

        Dim strHTML As String = ""

        '加title條件區
        Dim tdstyle As String = "" '依dataGrid資料欄
        tdstyle = " colspan='" & Cst_colspan & "'" '依dataGrid資料欄

        Dim myTable As String = ""
        myTable = ""
        myTable &= "<table border='0' cellspacing='0' cellpadding='0' align='center' style='width:100%;border-collapse@collapse;'>"
        myTable &= "<tr><td " & tdstyle & ">" & "查詢條件如下：" & "</td></tr>"
        myTable &= "<tr><td " & tdstyle & ">" & "訓練機構：" & center.Text & "</td></tr>"
        myTable &= "<tr><td " & tdstyle & ">" & "訓練職類：" & TB_career_id.Text & "</td></tr>"
        myTable &= "<tr><td " & tdstyle & ">" & "班級名稱：" & tb_classname.Text & "</td></tr>"
        myTable &= "<tr><td " & tdstyle & ">" & "開訓日期：" & start_date.Text & "~" & end_date.Text & "</td></tr>"
        'myTable &= "<tr><td " & tdstyle & ">" & "報名日期：" & redate_start.Text & "~" & redate_end.Text & "</td></tr>"
        'myTable &= "<tr><td " & tdstyle & ">" & "開班狀態：" & NotOpen.SelectedItem.Text & "</td></tr>"
        myTable &= "<table>"
        strHTML &= (myTable)
        'Common.RespWrite(Me, myTable)

        DG_Classinfo.AllowPaging = False
        DG_Classinfo.EnableViewState = False  '把ViewState給關了

        Dim objStringWriter As New System.IO.StringWriter
        Dim objHtmlTextWriter As New System.Web.UI.HtmlTextWriter(objStringWriter)
        div1.RenderControl(objHtmlTextWriter)
        strHTML &= (Convert.ToString(objStringWriter))
        'Common.RespWrite(Me, Convert.ToString(objStringWriter))

        '加title條件區
        myTable = ""
        myTable &= "<table border='0' cellspacing='0' cellpadding='0' align='center' style='width:100%;border-collapse@collapse;'>"
        myTable &= "<tr><td " & tdstyle & ">" & "統計數量：" & "</td></tr>"
        myTable &= "<tr><td " & tdstyle & ">" & "班級：" & iClassCnt & "筆</td></tr>"
        myTable &= "<tr><td " & tdstyle & ">" & "訓練人數：" & iStudCnt1 & "筆</td></tr>"
        myTable &= "<tr><td " & tdstyle & ">" & "報名人數：" & iStudCnt2 & "筆</td></tr>"
        myTable &= "<table>"
        strHTML &= (myTable)
        'Common.RespWrite(Me, myTable)

        '加title條件區
        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", strHTML)
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)

        DG_Classinfo.AllowPaging = True
        TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
        'DG_ClassInfo.Columns(8).Visible=True
    End Sub

    '查詢鈕
    Private Sub bt_search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_search.Click
        flagExportExcel = False ' 非 匯出Excel
        Call TIMS.SUtl_TxtPageSize(Me, txtpagesize, DG_Classinfo)
        '非產投查詢(TIMS)
        Call Search1()
    End Sub

    '匯出EXCEL
    Private Sub btnExport1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExport1.Click
        'TIMS
        Call Export1()
    End Sub
End Class