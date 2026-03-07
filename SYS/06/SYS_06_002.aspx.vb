Public Class SYS_06_002
    Inherits AuthBasePage

    'AUTH_ACCOUNTLOG
    Const cst_TOP_MAX_ROWS As String = "9999"

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
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End
        PageControler1.PageDataGrid = Me.DataGrid1

        If Not Me.IsPostBack Then
            Call Create1()
            'panelSearch.Visible = True '搜尋功能啟動 'PanelEdit1.Visible = False '修改功能關閉
        End If
    End Sub

    Sub Create1()
        DataGridTable1.Visible = False '預設搜尋資料不顯示

        lab_TOP_MAX_ROWS.Text = "最大查詢筆數：" & cst_TOP_MAX_ROWS

        WorkDate1.Text = TIMS.Cdate3(DateAdd(DateInterval.Day, -3, Now.Date))
        WorkDate2.Text = TIMS.Cdate3(Now.Date)

        ddlKind = TIMS.Get_ddlFunction(ddlKind)
        '作業方式
        ddlWorkMethod = TIMS.Get_WorkMethod(ddlWorkMethod, objconn)

        Dim V_INQUIRY As String = Session(String.Concat(TIMS.cst_GSE_V_INQUIRY, TIMS.Get_MRqID(Me))) 'LOG查詢 
        If TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES) Then Call TIMS.GET_INQUIRY(ddl_INQUIRY_Sch, objconn, V_INQUIRY)
    End Sub

    '查詢
    Private Sub btnQuery_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnQuery.Click

        TIMS.SUtl_TxtPageSize(Me, Me.TxtPageSize, Me.DataGrid1)

        Call search1()
    End Sub

    'SQL
    Sub search1()
        Dim dt As DataTable = Search1dt("")
        'Dim dt As DataTable  'dt = DbAccess.GetDataTable(sql, objconn)
        DataGridTable1.Visible = False
        msg.Text = "查無資料"
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            DataGridTable1.Visible = True

            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
        End If
    End Sub

    Function Search1dt(SCHTYPE As String) As DataTable
        Call sUtl_ViewStateValue()  '設定 ViewState Value

        'panelSearch.Visible = True '搜尋功能啟動 'PanelEdit1.Visible = False '修改功能關閉
        Dim sql As String = ""
        sql &= " SELECT TOP " & cst_TOP_MAX_ROWS & " a.ACCOUNT" & vbCrLf
        If SCHTYPE = "EXP" Then sql &= " ,a.RESDESC"
        sql &= " ,aa.NAME CNAME,f.NAME FUNNAME
,ISNULL(m.WNAME,'查詢') WorkMethod
,case when a.WorkMode='1' then '1:模糊顯示' when a.WorkMode='2' then '2:正常顯示' else '其他' end WorkMode
,format(a.WorkDate,'yyyy/MM/dd') WorkDate
,a.LAID,a.INQNO,a.RESCNT,a.NOTE
,(SELECT RNAME FROM KEY_INQUIRY WHERE INQNO=a.INQNO) INQNO_N
,(SELECT ORGNAME FROM ORG_ORGINFO WHERE ORGID=a.ORGID) ORGNAME
FROM AUTH_ACCOUNTLOG a WITH(NOLOCK)
LEFT JOIN AUTH_ACCOUNT aa WITH(NOLOCK) on aa.account=a.account
LEFT JOIN VIEW_FUNCTION f WITH(NOLOCK) on f.funid =a.funid
LEFT JOIN V_WORKMETHOD m WITH(NOLOCK) on m.WID =a.WorkMethod COLLATE Chinese_Taiwan_Stroke_CI_AS" & vbCrLf
        sql &= " WHERE f.Valid='Y'" & vbCrLf
        If SCHTYPE = "EXP" AndAlso ViewState("LAIDVALS") <> "" Then
            sql &= String.Concat(" AND a.LAID IN (", ViewState("LAIDVALS"), ")")
        End If
        If ViewState("tUserID") <> "" Then sql &= " AND a.account=@tUserID" & vbCrLf
        If ViewState("tIDNO") <> "" Then sql &= " AND aa.idno=@tIDNO" & vbCrLf
        If ViewState("Kind") <> "" Then sql &= " AND f.kind=@Kind" & vbCrLf
        '有子層。
        If ViewState("ddlFunID") <> "" Then
            sql &= " AND a.FUNID=@ddlFunID" & vbCrLf
        Else
            '沒有子層。
            If ViewState("FunParent") <> "" Then sql &= " AND f.pfunid=@FunParent" & vbCrLf
        End If
        If ViewState("ddlWorkMethod") <> "" Then sql &= " AND a.WorkMethod=@ddlWorkMethod" & vbCrLf
        If ViewState("WorkDate1") <> "" Then sql &= " and dbo.TRUNC_DATETIME(a.WorkDate)>=@WorkDate1" & vbCrLf
        If ViewState("WorkDate2") <> "" Then sql &= " and dbo.TRUNC_DATETIME(a.WorkDate)<=@WorkDate2" & vbCrLf
        If ViewState("rblWorkMode") <> "" Then sql &= " and a.WorkMode=@rblWorkMode" & vbCrLf
        If ViewState("INQUIRY") <> "" Then sql &= " and a.INQNO=@INQNO" & vbCrLf

        Dim dt As New DataTable
        Dim oCmd As New SqlCommand(sql, objconn)
        With oCmd
            .Parameters.Clear()
            If ViewState("tUserID") <> "" Then .Parameters.Add("tUserID", SqlDbType.VarChar).Value = ViewState("tUserID")
            If ViewState("tIDNO") <> "" Then .Parameters.Add("tIDNO", SqlDbType.VarChar).Value = ViewState("tIDNO")
            If ViewState("Kind") <> "" Then .Parameters.Add("Kind", SqlDbType.VarChar).Value = ViewState("Kind")
            If ViewState("ddlFunID") <> "" Then
                .Parameters.Add("ddlFunID", SqlDbType.VarChar).Value = ViewState("ddlFunID")
            Else
                If ViewState("FunParent") <> "" Then .Parameters.Add("FunParent", SqlDbType.VarChar).Value = ViewState("FunParent")
            End If
            If ViewState("ddlWorkMethod") <> "" Then .Parameters.Add("ddlWorkMethod", SqlDbType.VarChar).Value = ViewState("ddlWorkMethod")

            If ViewState("WorkDate1") <> "" Then .Parameters.Add("WorkDate1", SqlDbType.DateTime).Value = TIMS.Cdate2(ViewState("WorkDate1"))
            If ViewState("WorkDate2") <> "" Then .Parameters.Add("WorkDate2", SqlDbType.DateTime).Value = TIMS.Cdate2(ViewState("WorkDate2"))
            If ViewState("rblWorkMode") <> "" Then .Parameters.Add("rblWorkMode", SqlDbType.VarChar).Value = ViewState("rblWorkMode")
            If ViewState("INQUIRY") <> "" Then .Parameters.Add("INQNO", SqlDbType.VarChar).Value = ViewState("INQUIRY")
            dt.Load(.ExecuteReader())
        End With

        Return dt
    End Function

    '設定 ViewState Value
    Sub sUtl_ViewStateValue()
        tUserID.Text = TIMS.ClearSQM(tUserID.Text)
        tIDNO.Text = TIMS.ClearSQM(tIDNO.Text)
        Dim v_ddlKind As String = TIMS.GetListValue(ddlKind)
        '父層。
        Dim v_ddlFunP As String = TIMS.GetListValue(ddlFunP)
        '子層。
        Dim v_ddlFunC As String = TIMS.GetListValue(ddlFunC)
        Dim v_ddlWorkMethod As String = TIMS.GetListValue(ddlWorkMethod)
        WorkDate1.Text = TIMS.ClearSQM(WorkDate1.Text)
        WorkDate2.Text = TIMS.ClearSQM(WorkDate2.Text)
        Dim v_rblWorkMode As String = TIMS.GetListValue(rblWorkMode)
        Select Case v_rblWorkMode
            Case "1", "2"
            Case Else
                v_rblWorkMode = "" '(其它資訊清空)
        End Select
        Dim v_ddl_INQUIRY_Sch As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        Dim V_LAIDVALS As String = GET_LAIDVAL()

        ViewState("tUserID") = tUserID.Text
        ViewState("tIDNO") = TIMS.ChangeIDNO(tIDNO.Text)
        ViewState("Kind") = v_ddlKind
        ViewState("FunParent") = v_ddlFunP
        ViewState("ddlFunID") = v_ddlFunC
        ViewState("ddlWorkMethod") = v_ddlWorkMethod
        ViewState("WorkDate1") = If(WorkDate1.Text <> "", Common.FormatDate(WorkDate1.Text), "")
        ViewState("WorkDate2") = If(WorkDate2.Text <> "", Common.FormatDate(WorkDate2.Text), "")
        ViewState("rblWorkMode") = v_rblWorkMode
        ViewState("INQUIRY") = v_ddl_INQUIRY_Sch
        ViewState("LAIDVALS") = V_LAIDVALS
    End Sub

    Sub sUtl_ShowLevel0Func(ByVal objDDL As DropDownList, ByVal sKindFunID As String, ByVal Levels As String)
        'Levels: 0:父層 1:子層。'Dim objDDL As DropDownList'Dim textField As String'Dim valueField As String'objDDL = FunParent
        objDDL.Items.Clear()
        sKindFunID = TIMS.ClearSQM(sKindFunID)
        If sKindFunID = "" Then Exit Sub

        Dim textField As String = "fName"
        Dim valueField As String = "FunID"

        Dim sql As String = ""
        sql &= " SELECT a.FunID" & vbCrLf
        sql &= " ,a.Spage,a.LEVELS" & vbCrLf
        sql &= " ,a.Valid,a.KINDName" & vbCrLf
        sql &= " ,(CASE WHEN a.Spage IS NOT NULL THEN '*' ELSE '' END)+a.NAME fName" & vbCrLf
        sql &= " FROM VIEW_FUNCTION a WITH(NOLOCK)" & vbCrLf
        sql &= " WHERE a.Valid='Y'" & vbCrLf
        Select Case Levels
            Case "0"
                sql &= " AND a.Kind='" & sKindFunID & "'" & vbCrLf
                sql &= " AND (a.LEVELS='" & Levels & "')" & vbCrLf
            Case "1"
                sql &= " AND a.PFUNID='" & sKindFunID & "'" & vbCrLf
                sql &= " AND (a.LEVELS='" & Levels & "')" & vbCrLf
        End Select
        'sql &= " AND (a.Spage IS NULL)" & vbCrLf
        sql &= " ORDER BY a.Sort" & vbCrLf

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count > 0 Then
            objDDL.DataSource = dt
            objDDL.DataTextField = textField
            objDDL.DataValueField = valueField
            objDDL.DataBind()
        End If

    End Sub

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim lab_SNO As Label = e.Item.FindControl("lab_SNO")
                Dim Hid_LAID As HiddenField = e.Item.FindControl("Hid_LAID")
                Dim CB_SNO As HtmlInputCheckBox = e.Item.FindControl("CB_SNO")
                CB_SNO.Value = Convert.ToString(drv("LAID"))
                'CB_SNO.Attributes("onclick") = "InsertValue(this.checked,this.value)"
                lab_SNO.Text = TIMS.Get_DGSeqNo(sender, e) '序號
                Hid_LAID.Value = Convert.ToString(drv("LAID"))
        End Select
        'Case ListItemType.Header'    Dim Checkbox3 As HtmlInputCheckBox = e.Item.FindControl("Checkbox3")'    Checkbox3.Attributes("onclick") = "ChangeAll(this);"
    End Sub

    Protected Sub ddlKind_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlKind.SelectedIndexChanged
        Dim v_ddlKind As String = TIMS.GetListValue(ddlKind)
        '依KIND查詢子功能。
        Call sUtl_ShowLevel0Func(ddlFunP, v_ddlKind, "0")

        Dim v_ddlFunP As String = TIMS.GetListValue(ddlFunP)
        Call sUtl_ShowLevel0Func(ddlFunC, v_ddlFunP, "1")
    End Sub

    Protected Sub ddlFunP_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlFunP.SelectedIndexChanged
        '查詢底下功能
        Dim v_ddlFunP As String = TIMS.GetListValue(ddlFunP)
        '依KIND查詢子功能。
        Call sUtl_ShowLevel0Func(ddlFunC, v_ddlFunP, "1")
    End Sub

    Protected Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        UTL_EXPORT1()
        'Common.MessageBox(Me, "開發測試中!") 'Return
    End Sub

    Sub UTL_EXPORT1()
        Dim V_LAIDVALS As String = GET_LAIDVAL()
        If V_LAIDVALS = "" Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg14)
            Return
        End If
        Dim dtXls As DataTable = Search1dt("EXP")
        If TIMS.dtNODATA(dtXls) Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Return
        End If
        '編號,查詢時間,查詢功能路徑,使用人單位,使用人,使用人代號,查詢條件,查核結果,備註(查詢事由)
        Dim s_title1 As String = "編號,查詢時間,查詢功能路徑,使用人單位,使用人,使用人代號,查詢條件,查核結果,備註(查詢事由)"
        Dim s_data1 As String = "LAID,WORKDATE,FUNNAME,ORGNAME,CNAME,ACCOUNT,NOTE,RESDESC,INQNO_N"

        Dim AS_TITLE1() As String = s_title1.Split(",")
        Dim AS_DATA1() As String = s_data1.Split(",")
        Dim iColSpanCount As Integer = AS_TITLE1.Length + 1

        Const cst_ColFmt1 As String = "<td>{0}</td>"
        'Const cst_ColFmt2 As String = "<td class=""noDecFormat"">{0}</td>" '(純數字)
        'Const cst_ColFmt3 As String = "<td class=""DateFormat"">{0}</td>" '(文字)/(日期)

        Dim vROC_YERS As String = Now.Year - 1911
        Dim vROC_MONTH As String = Now.Month
        '勞動部勞動力發展署OO分署 OO年 在職進修訓練 年度執行成效   
        '匯出表頭名稱
        Dim sFileName1 As String = String.Format("export_{0}", TIMS.GetDateNo2())
        Dim s_TitleName As String = String.Format("{0} 年 {1} 月查詢紀錄單", vROC_YERS, vROC_MONTH)

        '套CSS值
        'mso-number-format:"0" 
        Dim strSTYLE As String = ""
        strSTYLE &= ("<style>")
        'strSTYLE &= ("td{mso-number-format:""\@"";}") '(文字)
        strSTYLE &= (".noDecFormat{mso-number-format:""0"";}") '(數字)
        strSTYLE &= (".DateFormat{mso-number-format:""\@"";}") '(文字)/(日期)
        strSTYLE &= ("</style>")

        Dim ExportStr As String '建立輸出文字
        Dim sbHTML As New StringBuilder
        sbHTML.Append("<div>")
        sbHTML.Append("<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")

        '表頭及查詢條件列
        ExportStr = String.Format("<tr><td align='center' colspan='{0}'>{1}</td></tr>", iColSpanCount, s_TitleName) & vbCrLf
        sbHTML.Append(ExportStr)

        ExportStr = "<tr>"
        ExportStr &= String.Format(cst_ColFmt1, "序號")
        For Each s_T1 As String In AS_TITLE1
            ExportStr &= String.Format(cst_ColFmt1, s_T1) '& vbTab
        Next
        ExportStr &= "</tr>"
        sbHTML.Append(ExportStr)

        '建立資料面
        Dim i_rows As Integer = 0
        For Each oDr1 As DataRow In dtXls.Rows
            i_rows += 1
            ExportStr = "<tr>"
            ExportStr &= String.Format(cst_ColFmt1, i_rows) '序號
            For Each s_D1 As String In AS_DATA1
                ExportStr &= String.Format(cst_ColFmt1, oDr1(s_D1)) '& vbTab
            Next
            ExportStr &= "</tr>"
            sbHTML.Append(ExportStr)
        Next

        sbHTML.Append("</table>")
        sbHTML.Append("</div>")

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", "EXCEL")
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", sbHTML.ToString())
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)
        TIMS.Utl_RespWriteEnd(Me, objconn, "")
    End Sub

    Function GET_LAIDVAL() As String
        Dim RST As String = ""
        For Each eItem As DataGridItem In DataGrid1.Items
            Dim lab_SNO As Label = eItem.FindControl("lab_SNO")
            Dim Hid_LAID As HiddenField = eItem.FindControl("Hid_LAID")
            Dim CB_SNO As HtmlInputCheckBox = eItem.FindControl("CB_SNO")
            If CB_SNO IsNot Nothing AndAlso CB_SNO.Checked Then
                Hid_LAID.Value = TIMS.ClearSQM(Hid_LAID.Value)
                CB_SNO.Value = TIMS.ClearSQM(CB_SNO.Value)
                If (CB_SNO.Value <> "" AndAlso CB_SNO.Value = Hid_LAID.Value) Then
                    RST &= String.Concat(If(RST <> "", ",", ""), CB_SNO.Value)
                End If
            End If
        Next
        RST = TIMS.CombiSQLINM3(RST)
        Return RST
    End Function

End Class
