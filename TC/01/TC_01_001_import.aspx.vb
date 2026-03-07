Partial Class TC_01_001_import
    Inherits AuthBasePage

    Dim FF3 As String = ""
    Dim dtTPlan As DataTable
    Dim dtDist As DataTable
    Dim objconn As SqlConnection
    Dim flag_ROC As Boolean = TIMS.CHK_REPLACE2ROC_YEARS()

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

        '分頁設定 Start
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End

        If Not IsPostBack Then
            Button1.Attributes("onclick") = "return chkdata();"
            Button2.Attributes("onclick") = "window.close();"
            Fromyear = TIMS.GetSyear(Fromyear)
            Toyear = TIMS.GetSyear(Toyear)
            PageControler1.Visible = False
            Table2.Visible = False

            Dim sql As String
            sql = " SELECT DISTID,NAME FROM ID_DISTRICT ORDER BY DISTID "
            dtDist = DbAccess.GetDataTable(sql, objconn)
            'Me.ViewState("DistID") = dt
            sql = " SELECT TPLANID,PLANNAME FROM KEY_PLAN ORDER BY TPLANID "
            dtTPlan = DbAccess.GetDataTable(sql, objconn)

            '要是署(局)的身分，要產生所有的轄區代碼
            If sm.UserInfo.LID = 0 Then
                With DistID
                    .DataSource = dtDist
                    .DataTextField = "NAME"
                    .DataValueField = "DISTID"
                    .DataBind()
                    .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
                End With
            End If
            'Me.ViewState("TPlan") = dt
        End If

        Table3.Style.Item("display") = "none"
        If sm.UserInfo.LID = 0 Then Table3.Style.Item("display") = ""
        If Not Session("_search") Is Nothing Then Me.ViewState("_search") = Session("_search") 'Session("_search") = Nothing
        'If Not Me.ViewState("dt") Is Nothing Then PageControler1.PageDataTable = Me.ViewState("dt")
    End Sub

    Function GET_INSERT_SQL() As String
        'Sql &= "    ,FLEXTURNOUTKIND ,EMAILSEND ,SUBTITLE ,PCOMMENT " & vbCrLf
        'Sql &= "    ,@FLEXTURNOUTKIND ,@EMAILSEND ,@SUBTITLE ,@PCOMMENT " & vbCrLf
        Dim sql As String = ""
        sql &= " INSERT INTO ID_PLAN (PLANID ,YEARS ,DISTID ,TPLANID ,SEQ ,SPONSOR ,COSPONSOR ,SDATE ,EDATE ,PLANKIND ,MODIFYACCT ,MODIFYDATE )" & vbCrLf
        sql &= " VALUES (@PLANID ,@YEARS ,@DISTID ,@TPLANID ,@SEQ ,@SPONSOR ,@COSPONSOR ,@SDATE ,@EDATE ,@PLANKIND ,@MODIFYACCT ,GETDATE() )" & vbCrLf
        Return sql
    End Function

    '匯入
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim sql As String = ""
        If sm.UserInfo.LID = 0 Then
            sql = " SELECT * FROM ID_Plan WHERE DistID = '" & DistID.SelectedValue & "' AND Years = '" & Fromyear.SelectedValue & "' "
        Else
            sql = " SELECT * FROM ID_Plan WHERE DistID = '" & sm.UserInfo.DistID & "' AND Years = '" & Fromyear.SelectedValue & "' "
        End If
        Dim dt1 As DataTable = DbAccess.GetDataTable(sql, objconn)

        If sm.UserInfo.LID = 0 Then
            sql = " SELECT * FROM ID_Plan WHERE DistID = '" & DistID.SelectedValue & "' AND Years = '" & Toyear.SelectedValue & "' "
        Else
            sql = " SELECT * FROM ID_Plan WHERE DistID = '" & sm.UserInfo.DistID & "' AND Years = '" & Toyear.SelectedValue & "' "
        End If
        'Dim da As SqlDataAdapter = Nothing
        Dim dt2 As DataTable = DbAccess.GetDataTable(sql, objconn)
        'Dim dt3 As DataTable = dt2.Clone
        Dim iRow3 As Integer = 0
        For Each dr1 As DataRow In dt1.Rows
            Select Case dr1("TPlanID")
                Case "36" '青年職涯計劃---排除
                Case Else
                    If dt2.Select("TPlanID='" & dr1("TPlanID") & "' and Seq='" & dr1("Seq") & "'").Length = 0 Then
                        Dim sqlins As String = GET_INSERT_SQL()
                        Dim iPLANID As Integer = DbAccess.GetNewId(objconn, "ID_PLAN_PLANID_SEQ,ID_PLAN,PLANID")
                        Dim parms As New Hashtable From {
                            {"PLANID", iPLANID},
                            {"YEARS", Toyear.SelectedValue},
                            {"DISTID", dr1("DistID")},
                            {"TPLANID", dr1("TPlanID")},
                            {"SEQ", dr1("Seq")},
                            {"SPONSOR", dr1("Sponsor")},
                            {"COSPONSOR", dr1("Cosponsor")},
                            {"SDATE", DateAdd(DateInterval.Year, 1, dr1("SDate"))},
                            {"EDATE", DateAdd(DateInterval.Year, 1, dr1("EDate"))},
                            {"PLANKIND", dr1("PlanKind")},
                            {"MODIFYACCT", sm.UserInfo.UserID}
                        }
                        DbAccess.ExecuteNonQuery(sqlins, objconn, parms)
                    Else
                        'dt3.ImportRow(dr1)
                        iRow3 += 1
                    End If
            End Select
        Next
        'DbAccess.UpdateDataTable(dt2, da)

        If iRow3 = 0 Then
            'Dim m_url1 As String = "TC\01\TC_01_001.aspx?ID=" & TIMS.ClearSQM(Request("ID"))
            'Common.MessageBox(Me, "匯入成功!", m_url1)
            'Page.RegisterStartupScript("", "<script>window.close();</script>")
            If ViewState("_search") IsNot Nothing AndAlso Session("_search") Is Nothing Then Session("_search") = Me.ViewState("_search")
            Common.MessageBox(Me, "匯入成功!")
            Exit Sub
        Else
#Region "(No Use)"

            'Table1.Visible = False
            'Table2.Visible = False
            'Table3.Visible = False
            'PageControler1.PageDataTable = dt3
            'PageControler1.ControlerLoad()
            'DataGrid1.DataSource = dt3
            'DataGrid1.DataBind()
            'Me.ViewState("dt") = dt3
            ''分頁用-   Start
            'DataGridPage1.MyDataTable = dt3
            'DataGridPage1.FirstTime()
            ''分頁用-   End
            'Page.RegisterStartupScript("", "<script>window.resizeBy(400,250);</script>")
            'Dim m_url1 As String = "TC\01\TC_01_001.aspx?ID=" & TIMS.ClearSQM(Request("ID"))
            'Common.MessageBox(Me, "匯入成功，但匯入有失敗的資料，可能是訓練計畫或序號重複!", m_url1)

#End Region
            If ViewState("_search") IsNot Nothing AndAlso Session("_search") Is Nothing Then Session("_search") = Me.ViewState("_search")
            Common.MessageBox(Me, "匯入成功，但匯入有失敗的資料，可能是訓練計畫或序號重複!")
            Exit Sub
            'TIMS.Utl_Redirect1(Me, "TC_01_001.aspx?ID=" & Request("ID") & "&editid=" & e.CommandArgument)
        End If
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                'Dim dt As DataTable
                'dt = Me.ViewState("DistID")
                FF3 = "DistID='" & drv("DistID").ToString & "'"
                If dtDist.Select(FF3).Length <> 0 Then e.Item.Cells(1).Text = dtDist.Select(FF3)(0)("Name")
                'dt = Me.ViewState("TPlan")
                FF3 = "TPlanID='" & drv("TPlanID").ToString & "'"
                If dtTPlan.Select(FF3).Length <> 0 Then e.Item.Cells(2).Text = dtTPlan.Select(FF3)(0)("PlanName")
                Dim PlanKind_N As String = ""
                If Convert.ToString(drv("PlanKind")) = "1" Then PlanKind_N = "自辦"
                If Convert.ToString(drv("PlanKind")) = "2" Then PlanKind_N = "委外"
                e.Item.Cells(8).Text = PlanKind_N '"自辦/委外/"
                If flag_ROC Then
                    e.Item.Cells(6).Text = TIMS.Cdate17(drv("SDate"))  '(將原先的西年日期改為民國日期，by:20180928、20181001)
                    e.Item.Cells(7).Text = TIMS.Cdate17(drv("EDate"))  '(將原先的西年日期改為民國日期，by:20180928、20181001)
                End If
        End Select
    End Sub

    '回上頁
    Protected Sub BTN_BACK1_Click(sender As Object, e As EventArgs) Handles BTN_BACK1.Click
        If ViewState("_search") IsNot Nothing AndAlso Session("_search") Is Nothing Then Session("_search") = Me.ViewState("_search")
        TIMS.Utl_Redirect1(Me, "TC_01_001.aspx?ID=" & TIMS.ClearSQM(Request("ID")))
    End Sub
End Class