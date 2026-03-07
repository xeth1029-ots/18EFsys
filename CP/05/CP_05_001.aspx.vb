Partial Class CP_05_001
    Inherits AuthBasePage

    Dim SFunID1 As String = "" '實地訪查紀錄表:136
    'Dim Auth_Relship As DataTable
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
        '檢查Session是否存在 End
        PageControler1.PageDataGrid = DataGrid1
        PageControler2.PageDataGrid = DataGrid2

        SFunID1 = GetFunctionID()
        If SFunID1 = "" Then
            Common.MessageBox(Me, "功能異常，請連絡系統管理者!!")
            Exit Sub
        End If

        If Not IsPostBack Then
            msg1.Text = ""
            msg2.Text = ""
            DataGridTable.Visible = False
            DetailTable.Visible = False
            CreateItem()
            CheckBox1.Checked = True
            end_date.Text = Now.Date
        End If
        CheckBox1.Attributes("onclick") = "SelectAll(this.checked," & CTID.Items.Count & ");"

        If Not Session("_SearchStr") Is Nothing Then
            CheckBox1.Checked = False
            Dim MyArray As Array
            Dim MyItem As String
            Dim MyValue As String

            MyArray = Split(Session("_SearchStr"), "&")
            For i As Integer = 0 To MyArray.Length - 1
                MyItem = Split(MyArray(i), "=")(0)
                MyValue = Split(MyArray(i), "=")(1)

                Select Case MyItem
                    Case "CTID"
                        For Each item As ListItem In CTID.Items
                            For j As Integer = 0 To Split(MyValue, ",").Length - 1
                                If item.Value = Split(MyValue, ",")(j) Then
                                    item.Selected = True
                                End If
                            Next
                        Next
                    Case "SearchOrgName"
                        SearchOrgName.Text = MyValue
                End Select
            Next
            Session("_SearchStr") = Nothing
            Button1_Click(sender, e)
        Else
            If Not IsPostBack Then
                Page.RegisterStartupScript("First", "<script>SelectAll(true," & CTID.Items.Count & ");</script>")
            End If
        End If

        Button1.Attributes("onclick") = "javascript:return check_data()"
        'Button1.Attributes("onclick") = "return check_data();"
    End Sub

    Sub CreateItem()
        Dim sql As String = ""
        sql &= "SELECT a.CTID,b.CTName FROM "
        sql += "(SELECT * FROM Auth_DistCity WHERE DistID='" & sm.UserInfo.DistID & "') a "
        sql += "JOIN ID_City b ON a.CTID=b.CTID "
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)

        With CTID
            .DataSource = dt
            .DataTextField = "CTName"
            .DataValueField = "CTID"
            .DataBind()
        End With
    End Sub

    '查詢按鈕
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1) '顯示列數不正確

        DataGrid1.CurrentPageIndex = 0

        Dim CTStr As String = ""
        For Each item As ListItem In CTID.Items
            If item.Selected = True Then
                If CTStr <> "" Then CTStr += ","
                CTStr += "'" & item.Value & "'"
            End If
        Next

        Dim sql As String = ""
        sql &= " SELECT distinct a.RSID,a.RID,a.PlanID,a.Relship,c.OrgName "
        sql &= " ,dbo.NVL(r3.OrgName2,i2.name) OrgName2" & vbCrLf 'OrgName2管控單位
        sql &= " FROM Auth_Relship a"
        sql &= " JOIN Org_OrgPlanInfo b ON a.RSID=b.RSID "
        sql &= " JOIN Org_OrgInfo c ON a.OrgID=c.OrgID "
        '計畫有限定
        '10	新興科技人才培訓
        '11	資訊軟體人才培訓
        sql &= " JOIN ID_Plan d ON a.PlanID=d.PlanID and d.TPlanID IN ('10','11') AND d.TPlanID='" & sm.UserInfo.TPlanID & "'"
        sql &= " JOIN Class_ClassInfo cc  on cc.RID=a.RID "
        sql &= " JOIN ID_Class f ON cc.CLSID=f.CLSID "
        sql &= " JOIN ID_DISTRICT i2 ON i2.DistID=d.DistID" & vbCrLf
        sql &= " LEFT JOIN MVIEW_RELSHIP23 r3 on r3.RID3=a.RID" & vbCrLf
        sql &= " WHERE 1=1"
        If CTStr <> "" Then
            sql &= " and cc.TaddressZip IN (SELECT ZipCode FROM ID_ZIP WHERE CTID IN (" & CTStr & "))"
        End If
        If SearchOrgName.Text <> "" Then
            sql &= " and c.OrgName like '%" & SearchOrgName.Text & "%'"
        End If
        If start_date.Text <> "" Then
            sql &= " and cc.STDate>= " & TIMS.To_date(start_date.Text)
        End If
        If end_date.Text <> "" Then
            sql &= " and cc.FTDate<= " & TIMS.To_date(end_date.Text) '" & end_date.Text & "'"
        End If

        msg1.Text = "查無資料"
        DataGridTable.Visible = False

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)
        If TIMS.dtNODATA(dt) Then Return

        msg1.Text = ""
        DataGridTable.Visible = True

        PageControler1.PageDataTable = dt
        PageControler1.PrimaryKey = "RSID"
        'PageControler1.Sort = ""
        PageControler1.ControlerLoad()
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        'If Trim(Me.TxtPageSize.Text) <> "" And IsNumeric(Me.TxtPageSize.Text) Then
        '    If CInt(Me.TxtPageSize.Text) >= 1 Then
        '        Me.TxtPageSize.Text = Trim(Me.TxtPageSize.Text)
        '    Else
        '        Common.RespWrite(Me, "<script>alert('顯示列數不正確，以10 帶入');</script>")
        '        Me.TxtPageSize.Text = 10
        '    End If
        'Else
        '    Common.RespWrite(Me, "<script>alert('顯示列數不正確，以10 帶入');</script>")
        '    Me.TxtPageSize.Text = 10
        'End If
        'Me.DataGrid1.PageSize = Me.TxtPageSize.Text

        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim btn1 As Button = e.Item.FindControl("Button2")
                Dim btn2 As Button = e.Item.FindControl("Button3")
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)

                'Dim ParentRID As String
                'If Split(drv("Relship"), "/").Length > 2 Then
                '    ParentRID = Split(drv("Relship"), "/")(Split(drv("Relship"), "/").Length - 3)
                '    If Auth_Relship.Select("RID='" & ParentRID & "'").Length <> 0 Then
                '        e.Item.Cells(1).Text = Auth_Relship.Select("RID='" & ParentRID & "'")(0)("OrgName")
                '    End If
                'End If
                e.Item.Cells(1).Text = Convert.ToString(drv("OrgName2"))

                btn1.CommandArgument = "RID='" & drv("RID") & "' and PlanID='" & drv("PlanID") & "'"
                btn2.CommandArgument = "RID=" & drv("RID") & " & PlanID=" & drv("PlanID") & "&start_date=" & start_date.Text & "&end_date2=" & end_date.Text & ""
                btn2.Attributes("onclick") = ReportQuery.ReportScript(Me, "list", "SchoolBegins_List", "RID=" & drv("RID") & " & PlanID=" & drv("PlanID") & "&start_date=" & start_date.Text & "&end_date2=" & end_date.Text & "")
        End Select
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Select Case e.CommandName
            Case "ShowClass"
                Dim sql As String
                'Dim NewSql As String
                Dim dt As DataTable
                Dim dr As DataRow

                sql = "SELECT b.OrgName FROM "
                sql += "(SELECT * FROM Auth_Relship WHERE " & e.CommandArgument & ") a "
                sql += "JOIN Org_OrgInfo b ON a.OrgID=b.OrgID "
                dr = DbAccess.GetOneRow(sql, objconn)
                OrgName.Text = dr("OrgName").ToString

                sql = ""
                sql &= " SELECT * FROM Class_ClassInfo WHERE " & e.CommandArgument & ""
                If start_date.Text <> "" Then
                    sql &= " and STDate>= " & TIMS.To_date(start_date.Text)
                End If
                If end_date.Text <> "" Then
                    sql &= " and FTDate<= " & TIMS.To_date(end_date.Text) '" & end_date.Text & "'"
                End If
                sql &= " ORDER BY OCID "
                'NewSql = TIMS.Get_SQLPAGE(sql, 1, DataGrid2.PageSize, "OCID")

                dt = DbAccess.GetDataTable(sql, objconn)

                msg2.Text = "查無資料"
                DataGridTable2.Visible = False
                If dt.Rows.Count > 0 Then
                    msg2.Text = ""
                    DataGridTable2.Visible = True
                    'PageControler2.SqlPrimaryKeyDataCreate(sql, "OCID")
                    PageControler2.PageDataTable = dt
                    PageControler2.PrimaryKey = "OCID"
                    PageControler2.Sort = "OCID"
                    PageControler2.ControlerLoad()
                End If

                SearchTable.Visible = False
                DetailTable.Visible = True
            Case "PrintClass"
                'Dim cGuid As String =   ReportQuery.GetGuid(Page)
                'Dim Url As String =   ReportQuery.GetUrl(Page)
                'Dim PlanID As String
                'Dim TPlanID As String

                ''If sm.UserInfo.DistID = "000" Then
                ''    TPlanID = sm.UserInfo.TPlanID
                ''Else
                ''    PlanID = sm.UserInfo.PlanID
                ''End If
                ''&TPlanID=" & TPlanID & "
                'Dim strScript As String
                'strScript = "<script language=""javascript"">" + vbCrLf
                'strScript += "window.open('" & Url & "GUID=" + cGuid + "&SQ_AutoLogout=true&sys=list&filename=SchoolBegins_List&path=TIMS&" & e.CommandArgument & "');" + vbCrLf
                'strScript += "</script>"
                'Page.RegisterStartupScript("window_onload", strScript)
        End Select
    End Sub

    Private Sub DataGrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim Link As LinkButton = e.Item.FindControl("LinkButton1")
                Dim btn1 As Button = e.Item.FindControl("Button5")
                Dim btn2 As Button = e.Item.FindControl("Button6")
                Dim btn3 As Button = e.Item.FindControl("Button7")
                Dim btn4 As Button = e.Item.FindControl("Button8")
                Dim btn5 As Button = e.Item.FindControl("Button9")

                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)
                Link.ForeColor = Color.Blue
                Link.Text = drv("ClassCName").ToString
                Link.CommandArgument = "../01/CP_01_001.aspx?ID=" & SFunID1 & "&DOCID=" & drv("OCID") & "&DID=" & Request("ID")
                If IsNumeric(drv("CyclType")) Then
                    Link.Text += "第" & Int(drv("CyclType")) & "期"
                End If
                If drv("STDate").ToString <> "" And drv("FTDate").ToString <> "" Then
                    e.Item.Cells(2).Text = FormatDateTime(drv("STDate"), 2) & "<BR>|<BR>" & FormatDateTime(drv("FTDate"), 2)
                End If

                btn1.CommandArgument = "OCID=" & drv("OCID")            '課程表
                btn2.CommandArgument = "OCID=" & drv("OCID")            '學員名冊
                btn3.CommandArgument = "OCID=" & drv("OCID")            '出缺勤
                btn4.CommandArgument = "OCID=" & drv("OCID")            '生活津貼印領清冊
                btn5.CommandArgument = "OCID=" & drv("OCID")            '生活津貼統計明細表
                btn1.Attributes("onclick") = ReportQuery.ReportScript(Me, "list", "course_list", "OCID=" & drv("OCID"))
                btn2.Attributes("onclick") = ReportQuery.ReportScript(Me, "MultiBlock", "Student_Report", "OCID=" & drv("OCID"))
                btn3.Attributes("onclick") = ReportQuery.ReportScript(Me, "list", "fall_vacant_list", "OCID=" & drv("OCID"))
                btn4.Attributes("onclick") = ReportQuery.ReportScript(Me, "MultiBlock", "living_Report", "OCID=" & drv("OCID"))
                btn5.Attributes("onclick") = ReportQuery.ReportScript(Me, "MultiBlock", "Subsidy_Report", "OCID=" & drv("OCID"))
        End Select
    End Sub

    '查詢FUNID 實地訪查紀錄表
    Function GetFunctionID() As String
        Dim RST As String = ""
        'SELECT * FROM ID_Function WHERE Name ='實地訪查紀錄表'
        Const cst_planname As String = "實地訪查紀錄表" '136
        Dim sql As String
        sql = "SELECT FunID FROM ID_Function WHERE Name ='" & cst_planname & "'"
        RST = DbAccess.ExecuteScalar(sql, objconn)
        Return RST
    End Function

    Private Sub DataGrid2_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid2.ItemCommand
        Select Case e.CommandName
            Case "record"
                GetSearch()
                'Response.Redirect(e.CommandArgument)
                Dim url1 As String = e.CommandArgument
                Call TIMS.Utl_Redirect(Me, objconn, url1)

                'Case "Course"               '課程表
                '    Dim cGuid As String =   ReportQuery.GetGuid(Page)
                '    Dim Url As String =   ReportQuery.GetUrl(Page)
                '    Dim strScript As String
                '    strScript = "<script language=""javascript"">" + vbCrLf
                '    strScript += "window.open('" & Url & "GUID=" + cGuid + "&SQ_AutoLogout=true&sys=list&filename=course_list&path=TIMS&" & e.CommandArgument & "');" + vbCrLf
                '    strScript += "</script>"
                '    Page.RegisterStartupScript("window_onload", strScript)
                'Case "StudentInfo"          '學員名冊
                '    Dim cGuid As String =   ReportQuery.GetGuid(Page)
                '    Dim Url As String =   ReportQuery.GetUrl(Page)
                '    Dim strScript As String
                '    strScript = "<script language=""javascript"">" + vbCrLf
                '    strScript += "window.open('" & Url & "GUID=" + cGuid + "&SQ_AutoLogout=true&sys=MultiBlock&path=TIMS&filename=Student_Report&" & e.CommandArgument & "');" + vbCrLf
                '    strScript += "</script>"
                '    Page.RegisterStartupScript("window_onload", strScript)
                'Case "TurnOut"              '出缺勤
                '    Dim cGuid As String =   ReportQuery.GetGuid(Page)
                '    Dim Url As String =   ReportQuery.GetUrl(Page)
                '    Dim strScript As String
                '    strScript = "<script language=""javascript"">" + vbCrLf
                '    strScript += "window.open('" & Url & "GUID=" + cGuid + "&SQ_AutoLogout=true&sys=list&filename=fall_vacant_list&path=TIMS&" & e.CommandArgument & "');" + vbCrLf
                '    strScript += "</script>"
                '    Page.RegisterStartupScript("window_onload", strScript)
                'Case "Subsidy"              '生活津貼印領清冊
                '    Dim cGuid As String =   ReportQuery.GetGuid(Page)
                '    Dim Url As String =   ReportQuery.GetUrl(Page)
                '    Dim strScript As String
                '    strScript = "<script language=""javascript"">" + vbCrLf
                '    strScript += "window.open('" & Url & "GUID=" + cGuid + "&AutoLogout=true&sys=MultiBlock&filename=living_Report&path=TIMS&" & e.CommandArgument & "');" + vbCrLf
                '    strScript += "</script>"
                '    Page.RegisterStartupScript("window_onload", strScript)
                'Case "Subsidy2"             '生活津貼統計明細表
                '    Dim cGuid As String =   ReportQuery.GetGuid(Page)
                '    Dim Url As String =   ReportQuery.GetUrl(Page)
                '    Dim strScript As String
                '    strScript = "<script language=""javascript"">" + vbCrLf
                '    strScript += "window.open('" & Url & "GUID=" + cGuid + "&SQ_AutoLogout=true&sys=MultiBlock&path=TIMS&filename=Subsidy_Report&" & e.CommandArgument & "');" + vbCrLf
                '    strScript += "</script>"
                '    Page.RegisterStartupScript("window_onload", strScript)
        End Select
    End Sub

    Sub GetSearch()
        Dim CTIDStr As String = ""
        For Each item As ListItem In CTID.Items
            If item.Value <> "" AndAlso item.Selected Then
                If CTIDStr <> "" Then CTIDStr &= ","
                CTIDStr &= item.Value
            End If
        Next
        Session("_SearchStr") = "CTID=" & CTIDStr
        Session("_SearchStr") += "&SearchOrgName=" & SearchOrgName.Text
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        SearchTable.Visible = True
        DetailTable.Visible = False
    End Sub
End Class
