Public Class SYS_03_035
    Inherits AuthBasePage

    Const cst_title1 As String = "功能使用盤點統計資料"
    Const cst_title2 As String = "功能使用盤點明細資料"

    Const cst_sys03035Scope As String = "sys03035Scope"
    Dim ss_sys03035Scope As String = "" '儲存查詢值。
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        If Not IsPostBack Then
            Call sUtl_Show(0)
            Call Create1()
        End If

    End Sub

    Sub Create1()
        btnExport1.Visible = False
        btnExport2.Visible = False
        ddlY1 = TIMS.GetSyear(ddlY1, 2013, 0, False)
        ddlY2 = TIMS.GetSyear(ddlY2, 2013, 0, False)

        ddlMY1 = TIMS.GetSyear(ddlMY1, 2013, 0, False)
        ddlMM1 = TIMS.Get_Month(ddlMM1, "")
        ddlMY2 = TIMS.GetSyear(ddlMY2, 2013, 0, False)
        ddlMM2 = TIMS.Get_Month(ddlMM2, "")

        Common.SetListItem(ddlY1, Now.Year.ToString)
        Common.SetListItem(ddlMY1, Now.Year.ToString)
        Common.SetListItem(ddlMM1, Now.Month.ToString)

        rblScope.Attributes.Add("onclick", "return chkSchScope();")
    End Sub

    '顯示
    Sub sUtl_Show(ByVal iType As Integer)
        SchTable.Visible = False
        DataTable1.Visible = False
        DataTable2.Visible = False
        Select Case iType
            Case 0
                SchTable.Visible = True
            Case 1
                DataTable1.Visible = True
            Case 2
                DataTable2.Visible = True
        End Select

    End Sub

    '查詢
    Protected Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        Call sUtl_Show(1)
        Call Show_DataGrid1()
    End Sub

    '查詢1
    Sub Show_DataGrid1(Optional ByVal tmpPage As Integer = 0)
        If txtFunName.Text <> "" Then txtFunName.Text = Trim(txtFunName.Text)
        If txtFunName.Text <> "" Then txtFunName.Text = TIMS.ClearSQM(txtFunName.Text)

        Dim SYM1 As String = ""
        Dim SYM2 As String = ""
        ss_sys03035Scope = ""

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " select f.funid" & vbCrLf
        sql += " ,f.name funname" & vbCrLf
        sql += " ,replace(replace(f.funpath2,'//','/'),'/','>>') funpath" & vbCrLf
        sql += " ,f.spage" & vbCrLf
        sql += " ,f.kind" & vbCrLf
        sql += " ,f.memo" & vbCrLf
        sql += " ,dbo.NVL(h.count1,0) count1" & vbCrLf
        sql += " FROM VIEW_FUNCTION f " & vbCrLf
        sql += " LEFT JOIN  (" & vbCrLf
        sql += "  select funid " & vbCrLf
        sql += "  ,count(1) count1 " & vbCrLf
        sql += "  from SYS_HISFUNCCHK " & vbCrLf
        sql += "  WHERE 1=1" & vbCrLf
        Select Case rblScope.SelectedValue
            Case "Y"
                If ddlY1.SelectedValue <> "" Then
                    sql += " AND DATEPART(YEAR, MODIFYDATE) >=@ddlY1" & vbCrLf
                End If
                If ddlY2.SelectedValue <> "" Then
                    sql += " AND DATEPART(YEAR, MODIFYDATE) <=@ddlY2" & vbCrLf
                End If

            Case "M"
                SYM1 = ddlMY1.SelectedValue & Right("0" & ddlMM1.SelectedValue, 2)
                SYM2 = ddlMY2.SelectedValue & Right("0" & ddlMM2.SelectedValue, 2)
                sql += " AND convert(varchar(6), MODIFYDATE, 112) >=@SYM1" & vbCrLf
                sql += " AND convert(varchar(6), MODIFYDATE, 112) <=@SYM2" & vbCrLf
            Case "D"
                If MDATE1.Text <> "" Then
                    sql += " AND MODIFYDATE >= @MDATE1" & vbCrLf
                End If
                If MDATE2.Text <> "" Then
                    sql += " AND MODIFYDATE <= @MDATE2" & vbCrLf
                End If

        End Select
        sql += "  group by funid " & vbCrLf
        sql += " ) h  on f.funid=h.funid" & vbCrLf
        sql += " WHERE 1=1" & vbCrLf
        If txtFunName.Text <> "" Then
            sql += " AND f.name like '%'+@FunName+'%'" & vbCrLf
        End If
        sql += "  ORDER BY  f.funid " & vbCrLf
        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        Dim oCmd As New SqlCommand(sql, objconn)
        With oCmd
            .Parameters.Clear()
            '.Parameters.Add("xxx", SqlDbType.VarChar).Value = ""
            TIMS.SetMyValue(ss_sys03035Scope, "rblScope", rblScope.SelectedValue)
            Select Case rblScope.SelectedValue
                Case "Y"
                    If ddlY1.SelectedValue <> "" Then
                        'sql += " AND DATEPART(YEAR, MODIFYDATE) >=@ddlY1" & vbCrLf
                        .Parameters.Add("ddlY1", SqlDbType.VarChar).Value = ddlY1.SelectedValue
                        TIMS.SetMyValue(ss_sys03035Scope, "ddlY1", ddlY1.SelectedValue)
                    End If
                    If ddlY2.SelectedValue <> "" Then
                        'sql += " AND DATEPART(YEAR, MODIFYDATE) >=@ddlY2" & vbCrLf
                        .Parameters.Add("ddlY2", SqlDbType.VarChar).Value = ddlY2.SelectedValue
                        TIMS.SetMyValue(ss_sys03035Scope, "ddlY2", ddlY2.SelectedValue)
                    End If

                Case "M"
                    'sql += " AND convert(varchar(6), MODIFYDATE, 112) >=@SYM1" & vbCrLf
                    .Parameters.Add("SYM1", SqlDbType.VarChar).Value = SYM1
                    'sql += " AND convert(varchar(6), MODIFYDATE, 112) <=@SYM2" & vbCrLf
                    .Parameters.Add("SYM2", SqlDbType.VarChar).Value = SYM2
                    TIMS.SetMyValue(ss_sys03035Scope, "SYM1", SYM1)
                    TIMS.SetMyValue(ss_sys03035Scope, "SYM2", SYM2)
                Case "D"
                    If MDATE1.Text <> "" Then
                        'sql += " AND MODIFYDATE >= @MDATE1" & vbCrLf
                        .Parameters.Add("MDATE1", SqlDbType.DateTime).Value = TIMS.Cdate2(MDATE1.Text)
                        TIMS.SetMyValue(ss_sys03035Scope, "MDATE1", MDATE1.Text)
                    End If
                    If MDATE2.Text <> "" Then
                        'sql += " AND MODIFYDATE <= @MDATE2" & vbCrLf
                        .Parameters.Add("MDATE2", SqlDbType.DateTime).Value = TIMS.Cdate2(MDATE2.Text)
                        TIMS.SetMyValue(ss_sys03035Scope, "MDATE2", MDATE2.Text)
                    End If
            End Select
            If txtFunName.Text <> "" Then
                .Parameters.Add("FunName", SqlDbType.VarChar).Value = txtFunName.Text
                TIMS.SetMyValue(ss_sys03035Scope, "FunName", txtFunName.Text)
            End If
            dt.Load(.ExecuteReader())
        End With

        btnExport1.Visible = False
        btnExport2.Visible = False

        lab_Msg1.Visible = True
        DataGrid1.Visible = False
        If dt.Rows.Count > 0 Then
            Me.ViewState(cst_sys03035Scope) = ss_sys03035Scope
            btnExport1.Visible = True
            btnExport2.Visible = True

            lab_Msg1.Visible = False
            DataGrid1.Visible = True

            DataGrid1.DataSource = dt
            DataGrid1.CurrentPageIndex = tmpPage
            DataGrid1.DataBind()
        End If

    End Sub

    '查詢2
    Sub Show_DataGrid2(ByVal ifunid As Integer, ByVal sys03035Scope As String, Optional ByVal tmpPage As Integer = 0)
        'ifunid:若為0 則是顯示全部功能範圍。
        'If txtFunName.Text <> "" Then txtFunName.Text = Trim(txtFunName.Text)
        Const cst_funPath As Integer = 1
        Const cst_funName As Integer = 2

        Dim srblScope As String = TIMS.GetMyValue(sys03035Scope, "rblScope")
        Dim SYM1 As String = TIMS.GetMyValue(sys03035Scope, "SYM1")
        Dim SYM2 As String = TIMS.GetMyValue(sys03035Scope, "SYM2")
        Dim sddlY1 As String = TIMS.GetMyValue(sys03035Scope, "ddlY1")
        Dim sddlY2 As String = TIMS.GetMyValue(sys03035Scope, "ddlY2")
        Dim sMDATE1 As String = TIMS.GetMyValue(sys03035Scope, "MDATE1")
        Dim sMDATE2 As String = TIMS.GetMyValue(sys03035Scope, "MDATE2")
        Dim sFunName As String = TIMS.GetMyValue(sys03035Scope, "FunName")

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " select h.HISID" & vbCrLf
        sql += " ,f.funid" & vbCrLf
        sql += " ,replace(replace(f.funpath2,'//','/'),'/','>>') funpath" & vbCrLf
        sql += " ,f.name funname" & vbCrLf
        sql += " ,f.spage" & vbCrLf
        sql += " ,f.kind" & vbCrLf
        sql += " ,f.memo" & vbCrLf
        sql += " ,ip.Years" & vbCrLf
        sql += " ,ip.distName" & vbCrLf
        sql += " ,ip.tplanid" & vbCrLf
        sql += " ,ip.planname tplanname" & vbCrLf
        sql += " ,ip.years+ip.distname+ip.planname+ip.seq planname" & vbCrLf
        sql += " ,oo.OrgName" & vbCrLf
        sql += " ,aa.Name AcctName" & vbCrLf
        sql += " ,aa.Account" & vbCrLf
        sql += " ,convert(varchar, h.modifydate, 120) MDATE" & vbCrLf
        sql += " FROM VIEW_FUNCTION f " & vbCrLf
        sql += " JOIN SYS_HISFUNCCHK h on f.funid=h.funid " & vbCrLf
        sql += " JOIN VIEW_PLAN ip ON ip.PLANID =h.PLANID " & vbCrLf
        sql += " LEFT JOIN AUTH_ACCOUNT aa ON aa.ACCOUNT =h.MODIFYACCT" & vbCrLf
        sql += " LEFT JOIN ORG_ORGINFO oo ON oo.ORGID =aa.ORGID " & vbCrLf
        sql += " WHERE 1=1" & vbCrLf
        'sql +=   and rownum <=10" & vbCrLf
        'sql +=   --AND convert(varchar(6), h.MODIFYDATE, 112) >=@SYM1  AND convert(varchar(6), h.MODIFYDATE, 112) <=@SYM2" & vbCrLf
        Select Case srblScope
            Case "Y"
                'sddlY1 = TIMS.GetMyValue(sys03035Scope, "ddlY1")
                'sddlY2 = TIMS.GetMyValue(sys03035Scope, "ddlY2")
                If sddlY1 <> "" Then
                    sql += " AND DATEPART(YEAR, h.MODIFYDATE) >=@ddlY1" & vbCrLf
                End If
                If sddlY2 <> "" Then
                    sql += " AND DATEPART(YEAR, h.MODIFYDATE) <=@ddlY2" & vbCrLf
                End If
            Case "M"
                'SYM1 = ddlMY1.SelectedValue & ddlMM1.SelectedValue
                'SYM2 = ddlMY2.SelectedValue & ddlMM2.SelectedValue
                sql += " AND convert(varchar(6), h.MODIFYDATE, 112) >=@SYM1" & vbCrLf
                sql += " AND convert(varchar(6), h.MODIFYDATE, 112) <=@SYM2" & vbCrLf
            Case "D"
                'sMDATE1 = TIMS.GetMyValue(sys03035Scope, "MDATE1")
                'sMDATE2 = TIMS.GetMyValue(sys03035Scope, "MDATE2")
                If sMDATE1 <> "" Then
                    sql += " AND h.MODIFYDATE >= @MDATE1" & vbCrLf
                End If
                If sMDATE2 <> "" Then
                    sql += " AND h.MODIFYDATE <= @MDATE2" & vbCrLf
                End If
        End Select
        If sFunName <> "" Then
            sql += " AND f.name like '%'+@FunName+'%'" & vbCrLf
        End If
        If ifunid <> 0 Then
            sql += " AND f.funid=@funid" & vbCrLf
        End If
        sql += " ORDER BY h.HISID" & vbCrLf

        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        Dim oCmd As New SqlCommand(sql, objconn)
        With oCmd
            .Parameters.Clear()
            '.Parameters.Add("xxx", SqlDbType.VarChar).Value = ""
            Select Case srblScope
                Case "Y"
                    'sddlY1 = TIMS.GetMyValue(sys03035Scope, "ddlY1")
                    'sddlY2 = TIMS.GetMyValue(sys03035Scope, "ddlY2")
                    If sddlY1 <> "" Then
                        .Parameters.Add("ddlY1", SqlDbType.VarChar).Value = sddlY1
                    End If
                    If sddlY2 <> "" Then
                        .Parameters.Add("ddlY2", SqlDbType.VarChar).Value = sddlY2
                    End If
                Case "M"
                    'SYM1 = ddlMY1.SelectedValue & ddlMM1.SelectedValue
                    'SYM2 = ddlMY2.SelectedValue & ddlMM2.SelectedValue
                    .Parameters.Add("SYM1", SqlDbType.VarChar).Value = SYM1
                    .Parameters.Add("SYM2", SqlDbType.VarChar).Value = SYM2
                Case "D"
                    'sMDATE1 = TIMS.GetMyValue(sys03035Scope, "MDATE1")
                    'sMDATE2 = TIMS.GetMyValue(sys03035Scope, "MDATE2")
                    If sMDATE1 <> "" Then
                        .Parameters.Add("MDATE1", SqlDbType.DateTime).Value = TIMS.Cdate2(sMDATE1)
                    End If
                    If sMDATE2 <> "" Then
                        .Parameters.Add("MDATE2", SqlDbType.DateTime).Value = TIMS.Cdate2(sMDATE2)
                    End If
            End Select
            If sFunName <> "" Then
                .Parameters.Add("FunName", SqlDbType.VarChar).Value = sFunName
            End If
            If ifunid <> 0 Then
                .Parameters.Add("funid", SqlDbType.VarChar).Value = ifunid
            End If
            dt.Load(.ExecuteReader())
        End With

        'btnExport1.Visible = False
        'btnExport2.Visible = False

        lab_Msg2.Visible = True
        DataGrid2.Visible = False
        If dt.Rows.Count > 0 Then
            'btnExport1.Visible = True
            'btnExport2.Visible = True

            lab_Msg2.Visible = False
            DataGrid2.Visible = True

            DataGrid2.Columns(cst_funPath).Visible = False
            DataGrid2.Columns(cst_funName).Visible = False
            If ifunid = 0 Then
                DataGrid2.Columns(cst_funPath).Visible = True
                DataGrid2.Columns(cst_funName).Visible = True
            End If

            DataGrid2.DataSource = dt
            DataGrid2.CurrentPageIndex = tmpPage
            DataGrid2.DataBind()


        End If

    End Sub

    Private Sub DataGrid1_ItemCommand(source As Object, e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Select Case e.CommandName
            Case "ListData1"
                Dim cmdArg As String = e.CommandArgument
                hidfunid.Value = TIMS.GetMyValue(cmdArg, "FunID")
                lFunPath.Text = TIMS.GetMyValue(cmdArg, "FunPath")
                lFunName.Text = TIMS.GetMyValue(cmdArg, "FunName")
                Call sUtl_Show(2)
                Show_DataGrid2(Val(hidfunid.Value), Me.ViewState(cst_sys03035Scope))
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.EditItem, ListItemType.Item, ListItemType.SelectedItem
                Dim dr_Data As DataRowView = e.Item.DataItem

                'Dim labFunID As Label = e.Item.FindControl("labFunID")
                Dim labFunPath As Label = e.Item.FindControl("labFunPath")
                Dim labFunName As Label = e.Item.FindControl("labFunName")
                Dim labCount As Label = e.Item.FindControl("labCount")
                Dim btnListData1 As LinkButton = e.Item.FindControl("btnListData1")

                e.Item.Cells(0).Text = sender.PageSize * sender.CurrentPageIndex + e.Item.ItemIndex + 1

                Dim cmdArg As String = ""
                Call TIMS.SetMyValue(cmdArg, "FunID", Convert.ToString(dr_Data("FunID")))
                Call TIMS.SetMyValue(cmdArg, "FunPath", Convert.ToString(dr_Data("FunPath")))
                Call TIMS.SetMyValue(cmdArg, "FunName", Convert.ToString(dr_Data("FunName")))

                'Call TIMS.SetMyValue(cmdArg, "Scope", Me.rblScope.SelectedValue)
                btnListData1.CommandArgument = cmdArg

                'labFunID.Text = Convert.ToString(dr_Data("FunID"))
                'labFunID.Text = Convert.ToString(sender.PageSize * sender.CurrentPageIndex + e.Item.ItemIndex + 1)
                labFunPath.Text = Convert.ToString(dr_Data("FunPath"))
                labFunName.Text = Convert.ToString(dr_Data("funname"))
                labCount.Text = Convert.ToString(dr_Data("count1"))

        End Select
    End Sub

    Private Sub DataGrid1_PageIndexChanged(source As Object, e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles DataGrid1.PageIndexChanged
        Show_DataGrid1(e.NewPageIndex)
    End Sub

    Private Sub DataGrid2_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.EditItem, ListItemType.Item, ListItemType.SelectedItem
                Dim dr_Data As DataRowView = e.Item.DataItem

                'Dim labFunID As Label = e.Item.FindControl("labFunID")
                'Dim labFunPath As Label = e.Item.FindControl("labFunPath")
                'Dim labFunName As Label = e.Item.FindControl("labFunName")
                'Dim labCount As Label = e.Item.FindControl("labCount")

                Dim labFunPath As Label = e.Item.FindControl("labFunPath")
                Dim labFunName As Label = e.Item.FindControl("labFunName")
                Dim labYears As Label = e.Item.FindControl("labYears")
                Dim labDistName As Label = e.Item.FindControl("labDistName")
                Dim labTPlanName As Label = e.Item.FindControl("labTPlanName")
                Dim labPlanName As Label = e.Item.FindControl("labPlanName")
                Dim labOrgName As Label = e.Item.FindControl("labOrgName")
                Dim labAcctName As Label = e.Item.FindControl("labAcctName")
                Dim labAccount As Label = e.Item.FindControl("labAccount")
                Dim labMDATE As Label = e.Item.FindControl("labMDATE")

                'Dim btnListData1 As LinkButton = e.Item.FindControl("btnListData1")

                e.Item.Cells(0).Text = sender.PageSize * sender.CurrentPageIndex + e.Item.ItemIndex + 1
                'Dim cmdArg As String = ""
                'Call TIMS.SetMyValue(cmdArg, "FunID", Convert.ToString(dr_Data("FunID")))
                ''Call TIMS.SetMyValue(cmdArg, "Scope", Me.rblScope.SelectedValue)
                'btnListData1.CommandArgument = cmdArg

                labFunPath.Text = Convert.ToString(dr_Data("FunPath"))
                labFunName.Text = Convert.ToString(dr_Data("FunName"))
                labYears.Text = Convert.ToString(dr_Data("Years"))
                labDistName.Text = Convert.ToString(dr_Data("DistName"))
                labTPlanName.Text = Convert.ToString(dr_Data("TPlanName"))
                labPlanName.Text = Convert.ToString(dr_Data("PlanName"))
                labOrgName.Text = Convert.ToString(dr_Data("OrgName"))
                labAcctName.Text = Convert.ToString(dr_Data("AcctName"))
                labAccount.Text = Convert.ToString(dr_Data("Account"))
                labMDATE.Text = Convert.ToString(dr_Data("MDATE"))

        End Select
    End Sub

    Private Sub DataGrid2_PageIndexChanged(source As Object, e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles DataGrid2.PageIndexChanged
        Show_DataGrid2(Val(hidfunid.Value), Me.ViewState(cst_sys03035Scope), e.NewPageIndex)
    End Sub

    Protected Sub btnBack2_Click(sender As Object, e As EventArgs) Handles btnBack2.Click
        Call sUtl_Show(1)
    End Sub

    Protected Sub btnBack1_Click(sender As Object, e As EventArgs) Handles btnBack1.Click
        Call sUtl_Show(0)
    End Sub

    '匯出統計
    Protected Sub btnExport1_Click(sender As Object, e As EventArgs) Handles btnExport1.Click
        Call sExport1()
    End Sub

    '匯出明細
    Protected Sub btnExport2_Click(sender As Object, e As EventArgs) Handles btnExport2.Click
        Call sExport2()
    End Sub

    '覆寫，不執行 MyBase.VerifyRenderingInServerForm 方法，解決執行 RenderControl 產生的錯誤   
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)

    End Sub

    '匯出 統計 資料 1
    Sub sExport1()
        Const cst_fun As Integer = 4
        'Dim Errmsg As String = ""
        'Call CheckData1(Errmsg)
        'If Errmsg <> "" Then
        '    Common.MessageBox(Page, Errmsg)
        '    Exit Sub
        'End If

        DataGrid1.AllowPaging = False
        DataGrid1.Columns(cst_fun).Visible = False
        DataGrid1.EnableViewState = False  '把ViewState給關了

        Call Show_DataGrid1()

        Dim sFileName As String = ""
        sFileName = HttpUtility.UrlEncode(cst_title1 & ".xls", System.Text.Encoding.UTF8)

        Response.Clear()
        Response.Buffer = True
        Response.Charset = "UTF-8" '設定字集

        Response.AppendHeader("Content-Disposition", "attachment;filename=" & sFileName)

        Response.ContentEncoding = System.Text.Encoding.GetEncoding("UTF-8")
        'Response.ContentType = "application/vnd.ms-excel " '內容型態設為Excel
        '文件內容指定為Excel
        Response.ContentType = "application/ms-excel;charset=utf-8"

        Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=UTF-8>")

        ''套CSS值
        Common.RespWrite(Me, "<style>")
        Common.RespWrite(Me, "td{mso-number-format:""\@"";}")
        Common.RespWrite(Me, "</style>")

        DataGrid1.AllowPaging = False
        DataGrid1.Columns(cst_fun).Visible = False
        DataGrid1.EnableViewState = False  '把ViewState給關了

        Dim objStringWriter As New System.IO.StringWriter
        Dim objHtmlTextWriter As New System.Web.UI.HtmlTextWriter(objStringWriter)
        Div1.RenderControl(objHtmlTextWriter)

        Common.RespWrite(Me, Convert.ToString(objStringWriter))
        Response.End()

        DataGrid1.AllowPaging = True
        DataGrid1.Columns(cst_fun).Visible = True

    End Sub

    '匯出 明細 資料 2
    Sub sExport2()
        'Dim Errmsg As String = ""
        'Call CheckData1(Errmsg)
        'If Errmsg <> "" Then
        '    Common.MessageBox(Page, Errmsg)
        '    Exit Sub
        'End If

        DataGrid2.AllowPaging = False
        'DataGrid1.Columns(8).Visible = False
        DataGrid2.EnableViewState = False  '把ViewState給關了

        Show_DataGrid2(0, Me.ViewState(cst_sys03035Scope), 0)

        Dim sFileName As String = ""
        sFileName = HttpUtility.UrlEncode(cst_title2 & ".xls", System.Text.Encoding.UTF8)

        Response.Clear()
        Response.Buffer = True
        Response.Charset = "UTF-8" '設定字集

        Response.AppendHeader("Content-Disposition", "attachment;filename=" & sFileName)

        Response.ContentEncoding = System.Text.Encoding.GetEncoding("UTF-8")
        'Response.ContentType = "application/vnd.ms-excel " '內容型態設為Excel
        '文件內容指定為Excel
        Response.ContentType = "application/ms-excel;charset=utf-8"

        Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=UTF-8>")

        ''套CSS值
        Common.RespWrite(Me, "<style>")
        Common.RespWrite(Me, "td{mso-number-format:""\@"";}")
        Common.RespWrite(Me, "</style>")

        DataGrid2.AllowPaging = False
        'DataGrid1.Columns(8).Visible = False
        DataGrid2.EnableViewState = False  '把ViewState給關了

        Dim objStringWriter As New System.IO.StringWriter
        Dim objHtmlTextWriter As New System.Web.UI.HtmlTextWriter(objStringWriter)
        Div2.RenderControl(objHtmlTextWriter)

        Common.RespWrite(Me, Convert.ToString(objStringWriter))
        Response.End()

        DataGrid2.AllowPaging = True
        'DataGrid1.Columns(8).Visible = True

    End Sub

End Class