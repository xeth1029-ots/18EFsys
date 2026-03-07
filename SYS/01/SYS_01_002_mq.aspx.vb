Partial Class SYS_01_002_mq
    Inherits AuthBasePage

    Dim s_CHKERRMSG As String = ""
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload

        If Not Page.IsPostBack Then
            Create1()
        End If

    End Sub

    Sub Create1()
        labNAME1.Text = ""
        DataGrid1.Visible = False
        msg.Text = ""

    End Sub

    Private Function CHK_Search1(ByRef ACCOUNT As String) As Boolean
        Dim rst As Boolean = False
        ACCOUNT = ""

        txt_ACCTID.Text = TIMS.ClearSQM(txt_ACCTID.Text)
        txt_ACCNTNAME.Text = TIMS.ClearSQM(txt_ACCNTNAME.Text)
        txt_ORGNAME.Text = TIMS.ClearSQM(txt_ORGNAME.Text)
        Dim v_rdoIsUsed As String = TIMS.GetListValue(rdoIsUsed)

        If txt_ACCTID.Text = "" AndAlso txt_ACCNTNAME.Text = "" AndAlso txt_ORGNAME.Text = "" Then
            Common.MessageBox(Me, "請輸入搜尋條件!")
            Return rst
        End If

        Dim hDIC As New Hashtable
        Dim sSql As String = ""
        sSql &= " SELECT a.ACCOUNT,COUNT(1) cnt1" & vbCrLf
        sSql &= " FROM AUTH_ACCOUNT a" & vbCrLf
        sSql &= " JOIN ID_ROLE b on a.RoleID=b.RoleID" & vbCrLf
        sSql &= " JOIN AUTH_ACCRWPLAN c on c.ACCOUNT=a.ACCOUNT" & vbCrLf
        sSql &= " JOIN AUTH_RELSHIP R ON R.RID=c.RID" & vbCrLf
        sSql &= " JOIN VIEW_PLAN ip on ip.PLANID=c.PLANID" & vbCrLf
        sSql &= " JOIN ORG_ORGINFO oo on oo.ORGID=a.ORGID" & vbCrLf
        sSql &= " WHERE 1=1" & vbCrLf

        If CB_LIKE11.Checked Then

            If txt_ACCTID.Text <> "" Then
                hDIC.Add("ACCOUNT_LK", txt_ACCTID.Text)
                sSql &= " AND a.ACCOUNT LIKE '%'+@ACCOUNT_LK+'%'" & vbCrLf
            End If
            If txt_ACCNTNAME.Text <> "" Then
                hDIC.Add("ACCNTNAME_LK", txt_ACCNTNAME.Text)
                sSql &= " AND a.NAME LIKE '%'+@ACCNTNAME_LK+'%'" & vbCrLf
            End If
            If txt_ORGNAME.Text <> "" Then
                hDIC.Add("ORGNAME_LK", txt_ORGNAME.Text)
                sSql &= " AND oo.ORGNAME LIKE '%'+@ORGNAME_LK+'%'" & vbCrLf
            End If
        Else

            If txt_ACCTID.Text <> "" Then
                hDIC.Add("ACCOUNT_LK", txt_ACCTID.Text)
                sSql &= " AND a.ACCOUNT=@ACCOUNT_LK" & vbCrLf
            End If
            If txt_ACCNTNAME.Text <> "" Then
                hDIC.Add("ACCNTNAME_LK", txt_ACCNTNAME.Text)
                sSql &= " AND a.NAME=@ACCNTNAME_LK" & vbCrLf
            End If
            If txt_ORGNAME.Text <> "" Then
                hDIC.Add("ORGNAME_LK", txt_ORGNAME.Text)
                sSql &= " AND oo.ORGNAME=@ORGNAME_LK" & vbCrLf
            End If
        End If
        Select Case v_rdoIsUsed
            Case "Y", "N"
                hDIC.Add("ISUSED", v_rdoIsUsed)
                sSql &= " AND a.ISUSED=@ISUSED" & vbCrLf
        End Select
        sSql &= " GROUP BY a.ACCOUNT" & vbCrLf

        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objconn, hDIC)

        '只能查詢1筆資料
        If dt IsNot Nothing AndAlso dt.Rows.Count = 1 Then
            rst = True
            ACCOUNT = Convert.ToString(dt.Rows(0)("ACCOUNT"))
        End If

        s_CHKERRMSG = ""
        If dt Is Nothing Then
            s_CHKERRMSG &= "資料為空!"
        ElseIf dt.Rows.Count = 0 Then
            s_CHKERRMSG &= "查無帳號資料!"
        ElseIf dt.Rows.Count > 1 Then
            s_CHKERRMSG &= String.Concat("帳號資訊多筆(", dt.Rows.Count, ")")
        End If

        Return rst
    End Function

    Function SCH_ACCOUNT1_dt(ByRef ACCOUNT As String) As DataTable

        txt_ACCTID.Text = TIMS.ClearSQM(txt_ACCTID.Text)
        txt_ACCNTNAME.Text = TIMS.ClearSQM(txt_ACCNTNAME.Text)
        txt_ORGNAME.Text = TIMS.ClearSQM(txt_ORGNAME.Text)
        Dim v_rdoIsUsed As String = TIMS.GetListValue(rdoIsUsed)

        Dim hDIC As New Hashtable From {{"ACCOUNT", ACCOUNT}}
        Dim sSql As String = ""
        sSql &= " SELECT a.ACCOUNT,a.NAME ACCTNAME,b.NAME ROLENAME,a.LID" & vbCrLf
        sSql &= " ,ip.DISTID,ip.DISTNAME,ip.YEARS,ip.PLANNAME" & vbCrLf
        sSql &= " ,a.ORGID,oo.ORGNAME,a.ISUSED, CASE a.ISUSED WHEN 'Y' THEN '啟用' WHEN 'N' THEN '停用' ELSE a.ISUSED END ISUSED_N" & vbCrLf
        sSql &= " ,concat(ip.Years,ip.DISTNAME,ip.PlanName,ip.SEQ) USERPLAN2" & vbCrLf
        sSql &= " FROM AUTH_ACCOUNT a" & vbCrLf
        sSql &= " JOIN ID_ROLE b on a.RoleID=b.RoleID" & vbCrLf
        sSql &= " JOIN AUTH_ACCRWPLAN c on c.ACCOUNT=a.ACCOUNT" & vbCrLf
        sSql &= " JOIN AUTH_RELSHIP R ON R.RID=c.RID" & vbCrLf
        sSql &= " JOIN VIEW_PLAN ip on ip.PLANID=c.PLANID" & vbCrLf
        sSql &= " JOIN ORG_ORGINFO oo on oo.ORGID=a.ORGID" & vbCrLf
        sSql &= " WHERE a.ACCOUNT=@ACCOUNT" & vbCrLf

        If CB_LIKE11.Checked Then
            If txt_ACCTID.Text <> "" Then
                hDIC.Add("ACCOUNT_LK", txt_ACCTID.Text)
                sSql &= " AND a.ACCOUNT LIKE '%'+@ACCOUNT_LK+'%'" & vbCrLf
            End If
            If txt_ACCNTNAME.Text <> "" Then
                hDIC.Add("ACCNTNAME_LK", txt_ACCNTNAME.Text)
                sSql &= " AND a.NAME LIKE '%'+@ACCNTNAME_LK+'%'" & vbCrLf
            End If
            If txt_ORGNAME.Text <> "" Then
                hDIC.Add("ORGNAME_LK", txt_ORGNAME.Text)
                sSql &= " AND oo.ORGNAME LIKE '%'+@ORGNAME_LK+'%'" & vbCrLf
            End If
        Else
            If txt_ACCTID.Text <> "" Then
                hDIC.Add("ACCOUNT_LK", txt_ACCTID.Text)
                sSql &= " AND a.ACCOUNT=@ACCOUNT_LK" & vbCrLf
            End If
            If txt_ACCNTNAME.Text <> "" Then
                hDIC.Add("ACCNTNAME_LK", txt_ACCNTNAME.Text)
                sSql &= " AND a.NAME=@ACCNTNAME_LK" & vbCrLf
            End If
            If txt_ORGNAME.Text <> "" Then
                hDIC.Add("ORGNAME_LK", txt_ORGNAME.Text)
                sSql &= " AND oo.ORGNAME=@ORGNAME_LK" & vbCrLf
            End If
        End If

        Select Case v_rdoIsUsed
            Case "Y", "N"
                hDIC.Add("ISUSED", v_rdoIsUsed)
                sSql &= " AND a.ISUSED=@ISUSED" & vbCrLf
        End Select
        sSql &= " ORDER BY a.ACCOUNT, ip.DISTID,ip.YEARS,oo.ORGID" & vbCrLf

        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objconn, hDIC)

        Return dt
    End Function

    Private Sub Search1(ACCOUNT As String)
        labNAME1.Text = ""
        DataGrid1.Visible = False
        msg.Text = "查無資料"

        Dim dt As DataTable = SCH_ACCOUNT1_dt(ACCOUNT)

        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return

        Dim dr1 As DataRow = dt.Rows(0)

        Dim str_labNAME As String = String.Concat("姓名：", dr1("ACCTNAME"), "，", "帳號：", dr1("ACCOUNT"))

        labNAME1.Text = str_labNAME
        DataGrid1.Visible = True
        msg.Text = ""

        DataGrid1.DataSource = dt
        DataGrid1.DataBind()
    End Sub

    Protected Sub btn_SEARCH1_Click(sender As Object, e As EventArgs) Handles btn_SEARCH1.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Dim ACCOUNT As String = ""
        Dim fg_cango As Boolean = False

        labNAME1.Text = ""
        DataGrid1.Visible = False
        msg.Text = ""

        Try
            fg_cango = CHK_Search1(ACCOUNT)
        Catch ex As Exception
            msg.Text = "查詢資料有誤！"
            DataGrid1.Visible = False
            Call TIMS.WriteTraceLog(Me, ex, ex.Message)
        End Try

        If Not fg_cango OrElse ACCOUNT = "" Then
            Common.MessageBox(Me, "搜尋帳號有誤，請調整搜尋資訊!")
            Return
        End If

        Try
            If fg_cango Then Search1(ACCOUNT)
        Catch ex As Exception
            msg.Text = "查詢資料有誤！！"
            DataGrid1.Visible = False
            Call TIMS.WriteTraceLog(Me, ex, ex.Message)
        End Try

    End Sub

    Private Sub CheckData1(ByRef errmsg As String)
        errmsg = ""

        txt_ACCTID.Text = TIMS.ClearSQM(txt_ACCTID.Text)
        txt_ACCNTNAME.Text = TIMS.ClearSQM(txt_ACCNTNAME.Text)
        txt_ORGNAME.Text = TIMS.ClearSQM(txt_ORGNAME.Text)
        Dim v_rdoIsUsed As String = TIMS.GetListValue(rdoIsUsed)

        If txt_ACCTID.Text = "" AndAlso txt_ACCNTNAME.Text = "" AndAlso txt_ORGNAME.Text = "" Then
            errmsg &= "請輸入搜尋條件!" & vbCrLf
            Return
        End If

        Dim ACCOUNT As String = ""
        Dim fg_cango As Boolean = False
        s_CHKERRMSG = ""

        Try
            fg_cango = CHK_Search1(ACCOUNT)
        Catch ex As Exception
            'msg.Text = "查詢資料有誤！"
            errmsg &= "查詢資料有誤！" & ex.Message & vbCrLf
            Return
            'DataGrid1.Visible = False
            'Call TIMS.WriteTraceLog(Me, ex, ex.Message)
        End Try

        If Not fg_cango OrElse ACCOUNT = "" Then
            errmsg &= String.Concat("搜尋帳號有誤，請調整搜尋資訊!", s_CHKERRMSG, vbCrLf)
            Return
        End If
    End Sub

    Sub ExportX1(ByRef dtX1 As DataTable)

        If dtX1 Is Nothing OrElse dtX1.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無資料!")
            Return
        End If

        Dim str_TITLE1 As String = "帳號-計畫賦予(跨年度)"

        Dim dr1 As DataRow = dtX1.Rows(0)
        Dim str_labNAME As String = String.Concat("姓名：", dr1("ACCTNAME"), "，", "帳號：", dr1("ACCOUNT"))

        Dim sPattern As String = "" '序號,
        sPattern &= "年度,計畫,詳細計畫名稱,分署,單位"
        Dim sColumn As String = "" '序號,
        sColumn &= "YEARS,PLANNAME,USERPLAN2,DISTNAME,ORGNAME"

        Dim sPatternA() As String = Split(sPattern, ",")
        Dim sColumnA() As String = Split(sColumn, ",")

        Dim i_Colspan As Integer = sColumnA.Length + 1

        'Dim sFileName1 As String = String.Concat("帳號計畫賦予(跨年度)", TIMS.GetDateNo2())
        Dim s_FILENAME1 As String = String.Concat("帳號計畫賦予(跨年度)", "_", TIMS.GetDateNo2(3))

        '套CSS值
        Dim strSTYLE As String = ""
        strSTYLE &= "<style>"
        strSTYLE &= "td{mso-number-format:""\@"";}"
        strSTYLE &= ".noDecFormat{mso-number-format:""0"";}"
        strSTYLE &= "</style>"

        Dim strHTML As String = ""
        strHTML &= "<div>"
        strHTML &= "<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">"
        'Common.RespWrite(Me, "<tr>")

        '標題抬頭
        Dim ExportStr As String = "" '建立輸出文字

        ExportStr = String.Format("<tr><td colspan='{0}'>{1}</td></tr>", i_Colspan, str_TITLE1) & vbCrLf
        strHTML &= ExportStr

        ExportStr = String.Format("<tr><td colspan='{0}'>{1}</td></tr>", i_Colspan, str_labNAME) & vbCrLf
        strHTML &= ExportStr

        ExportStr = "<tr>"
        ExportStr &= String.Format("<td>{0}</td>", "序號") '& vbTab
        For i As Integer = 0 To sPatternA.Length - 1
            ExportStr &= String.Format("<td>{0}</td>", sPatternA(i)) '& vbTab
        Next
        ExportStr &= "</tr>" & vbCrLf
        strHTML &= ExportStr

        '建立資料面
        Dim iNum As Integer = 0
        ExportStr = ""
        For Each dr As DataRow In dtX1.Rows
            iNum += 1
            ExportStr = "<tr>"
            ExportStr &= String.Format("<td>{0}</td>", iNum) '& vbTab
            For i As Integer = 0 To sColumnA.Length - 1
                ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr(sColumnA(i))))
            Next
            ExportStr &= "</tr>" & vbCrLf
            strHTML &= ExportStr
        Next
        strHTML &= "</table>"
        strHTML &= "</div>"

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType)) 'EXCEL/PDF/ODS
        parmsExp.Add("FileName", s_FILENAME1)
        parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", strHTML)
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)
        TIMS.Utl_RespWriteEnd(Me, objconn, "") 'TIMS.CloseDbConn(objconn) 'Response.End()
    End Sub

    Protected Sub btn_EXPORT1_Click(sender As Object, e As EventArgs) Handles btn_EXPORT1.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        labNAME1.Text = ""
        DataGrid1.Visible = False
        msg.Text = ""

        Dim ACCOUNT As String = ""
        Dim fg_cango As Boolean = CHK_Search1(ACCOUNT)

        If Not fg_cango OrElse ACCOUNT = "" Then
            Common.MessageBox(Me, "搜尋帳號有誤，請調整搜尋資訊!")
            Return
        End If

        Dim dtExp As DataTable = SCH_ACCOUNT1_dt(ACCOUNT)

        Call ExportX1(dtExp)
    End Sub

    Protected Sub btnBACK1_Click(sender As Object, e As EventArgs) Handles btnBACK1.Click
        TIMS.Utl_Redirect1(Me, String.Concat("SYS_01_002.aspx?ID=", TIMS.Get_MRqID(Me)))
    End Sub

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)
        End Select
    End Sub
End Class
