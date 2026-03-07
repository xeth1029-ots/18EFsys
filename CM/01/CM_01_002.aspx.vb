Partial Class CM_01_002
    Inherits AuthBasePage

    Dim processtype As String = ""
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在--------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在--------------------------End

#Region "(No Use)"

        'If sm.UserInfo.FunDt Is Nothing Then
        '    Common.RespWrite(Me, "<script>alert('Session過期');</script>")
        '    Common.RespWrite(Me, "<script>top.location.href='../../logout.aspx';</script>")
        'Else
        '    Dim FunDt As DataTable = sm.UserInfo.FunDt
        '    Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & Request("ID") & "'")
        '    If FunDrArray.Length = 0 Then
        '        Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
        '        Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
        '    End If
        'End If

#End Region

        processtype = Request("processtype")
        If Not Page.IsPostBack Then
            'syear.Items.Insert(0, New ListItem("===請選擇===", ""))
            syear = TIMS.GetSyear(syear)
            If processtype = "back" Then
                syear.SelectedValue = Request("year")
                syear_SelectedIndexChanged(sender, e)
            End If
        End If
    End Sub

#Region "(No Use)"

    'Sub x()
    '    Sql = "select name as distname,district.distid,isnull(b.advance_class,0)as advance_class,isnull(a.advance_total,0)as advance_total ,isnull(c.real_class,0) as real_class,isnull(d.real_total,0)as real_total from " & vbCrLf
    '    Sql += " (SELECT * FROM ID_District) district  left join (" & vbCrLf
    '    Sql += " ( select isnull( sum( isnull(DefGovCost,0)+isnull(DefUnitCost,0)+isnull(DefStdCost,0)),0) as advance_total,b.distid from (select * from   Plan_PlanInfo where Planyear=" & syear.SelectedValue & " ) a " & vbCrLf
    '    Sql += "join  id_plan b on a.planid=b.planid  " & vbCrLf
    '    Sql += " where AppliedResult='Y'    group by b.distid)a left join " & vbCrLf
    '    Sql += " (select b.distid,count(1)as advance_class  from (select * from   Plan_PlanInfo where Planyear=" & syear.SelectedValue & " )  a " & vbCrLf
    '    Sql += " join  id_plan b on a.planid=b.planid   where AppliedResult='Y'   group by b.distid)b on a.distid=b.distid left join  " & vbCrLf
    '    Sql += " (select count(1) as real_class,b.distid  from (select * from Plan_PlanInfo where  Planyear=" & syear.SelectedValue & " )p_table " & vbCrLf
    '    Sql += "  join  id_plan b on  p_table.planid=b.planid  " & vbCrLf
    '    Sql += "  join  Class_ClassInfo c on  p_table.planid=c.planid and  p_table.ComIDNO=c.ComIDNO and p_table.SeqNO=c.SeqNO and p_table.rid=c.rid  where c.NotOpen='N' and IsSuccess='Y'  and p_table.AppliedResult='Y' and c.STDate<=getdate()  group by b.distid) c on c.distid=b.distid  left join  " & vbCrLf
    '    Sql += " (select isnull( sum( isnull(DefGovCost,0)+isnull(DefUnitCost,0)+isnull(DefStdCost,0)),0) as real_total,b.distid  from (select * from Plan_PlanInfo where  Planyear=" & syear.SelectedValue & " )p_table  " & vbCrLf
    '    Sql += "  join  id_plan b on  p_table.planid=b.planid   " & vbCrLf
    '    Sql += "  join  Class_ClassInfo c on  p_table.planid=c.planid and  p_table.ComIDNO=c.ComIDNO and p_table.SeqNO=c.SeqNO and p_table.rid=c.rid " & vbCrLf
    '    Sql += " where NotOpen='N'  and IsSuccess='Y' and p_table.AppliedResult='Y' and c.STDate<=getdate() group by b.distid)d on d.distid=c.distid)  on a.distid=district.distid order by district. distid" & vbCrLf
    'End Sub

#End Region

    Private Sub syear_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles syear.SelectedIndexChanged
        Dim sql As String = ""
        Dim dt As DataTable

        If Me.syear.SelectedValue <> "" Then
            Panel.Visible = True
            sql = "" & vbCrLf
            sql += " SELECT dd.name distname ,dd.distid " & vbCrLf
            sql += "        ,ISNULL(b.advance_class,0) advance_class " & vbCrLf '/*年度預定開班 */
            sql += "        ,ISNULL(b.advance_total,0) advance_total " & vbCrLf '/*年度預定總預算 */
            sql += "        ,ISNULL(c.real_class,0) real_class " & vbCrLf '/* 實際開班數*/
            sql += "        ,ISNULL(c.real_total,0) real_total " & vbCrLf '/*實際總經費*/
            sql += " FROM ID_District dd " & vbCrLf
            sql += " LEFT JOIN ( " & vbCrLf
            sql += "   SELECT b.distid ,COUNT(1) advance_class ,SUM(ISNULL(a.DefGovCost,0)+ISNULL(a.DefUnitCost,0)+ISNULL(a.DefStdCost,0)) AS advance_total " & vbCrLf
            sql += "   FROM Plan_PlanInfo a " & vbCrLf
            sql += "   JOIN id_plan b ON a.planid = b.planid AND a.Planyear = '" & syear.SelectedValue & "' " & vbCrLf
            sql += "   WHERE AppliedResult = 'Y' " & vbCrLf
            sql += "   GROUP by b.distid " & vbCrLf
            sql += " ) b ON dd.distid = b.distid " & vbCrLf
            sql += " LEFT JOIN ( " & vbCrLf
            sql += "   SELECT b.distid ,COUNT(1) real_class ,SUM(ISNULL(p_table.DefGovCost,0) + ISNULL(p_table.DefUnitCost,0) + ISNULL(p_table.DefStdCost,0)) AS real_total " & vbCrLf
            sql += "   FROM Plan_PlanInfo p_table " & vbCrLf
            sql += "   JOIN id_plan b ON p_table.planid = b.planid AND p_table.Planyear = '" & syear.SelectedValue & "' " & vbCrLf
            sql += "   JOIN Class_ClassInfo c ON p_table.planid = c.planid AND p_table.ComIDNO = c.ComIDNO AND p_table.SeqNO = c.SeqNO AND p_table.rid = c.rid " & vbCrLf
            sql += "   WHERE 1=1 " & vbCrLf
            sql += "      AND c.NotOpen = 'N' " & vbCrLf
            sql += "      AND c.IsSuccess = 'Y' " & vbCrLf
            sql += "      AND p_table.AppliedResult = 'Y' " & vbCrLf
            sql += "      AND c.STDate <= dbo.TRUNC_DATETIME(GETDATE()) " & vbCrLf
            sql += "   GROUP BY b.distid " & vbCrLf
            sql += " ) c ON dd.distid = c.distid " & vbCrLf
            sql += " ORDER BY dd.distid " & vbCrLf
            dt = DbAccess.GetDataTable(sql, objconn)
            DG_Grid1.DataSource = dt
            DG_Grid1.DataBind()
            DG_Grid1.Visible = True

            '細目
            sql = "" & vbCrLf
            sql += " SELECT dd.name distname ,dd.distid " & vbCrLf
            sql += "        ,ISNULL(a.real_total,0) real_total " & vbCrLf ' /*實際總經費*/
            sql += "        ,ISNULL(a.cancel_total,0) cancel_total " & vbCrLf '/*實際開班已付金額*/
            sql += "        ,ISNULL(a.real_total,0)-ISNULL(a.cancel_total,0) money_all " & vbCrLf
            sql += " FROM ID_District dd " & vbCrLf
            sql += " LEFT JOIN ( " & vbCrLf
            sql += "   SELECT b.distid ,SUM(ISNULL(p_table.DefGovCost,0) + ISNULL(p_table.DefUnitCost,0) + ISNULL(p_table.DefStdCost,0)) real_total ,SUM(bug.CancelCost) cancel_total " & vbCrLf
            sql += "   FROM Plan_PlanInfo p_table " & vbCrLf
            sql += "   JOIN id_plan b ON p_table.planid = b.planid AND p_table.Planyear = '" & syear.SelectedValue & "' " & vbCrLf
            sql += "   JOIN Class_ClassInfo c ON p_table.planid = c.planid AND p_table.ComIDNO = c.ComIDNO AND p_table.SeqNO = c.SeqNO AND p_table.rid = c.rid " & vbCrLf
            sql += "   LEFT JOIN Budget_ClassCancel bug ON bug.ocid = c.ocid " & vbCrLf
            sql += "   WHERE 1=1 " & vbCrLf
            sql += "      AND c.NotOpen = 'N'  " & vbCrLf
            sql += "      AND c.IsSuccess = 'Y' " & vbCrLf
            sql += "      AND p_table.AppliedResult = 'Y' " & vbCrLf
            sql += "      AND c.STDate <= dbo.TRUNC_DATETIME(GETDATE()) " & vbCrLf
            sql += "   GROUP BY b.distid " & vbCrLf
            sql += " ) a on a.distid = dd.distid " & vbCrLf
            sql += " ORDER BY dd.distid " & vbCrLf
            dt = DbAccess.GetDataTable(sql, objconn)
            DataGrid1.DataSource = dt
            DataGrid1.DataBind()
        End If
    End Sub

    Private Sub DG_Grid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DG_Grid1.ItemDataBound
        Dim dr As DataRowView
        Dim MyLabel As Label
        dr = e.Item.DataItem

        Select Case e.Item.ItemType
            Case ListItemType.Header
                Dim btn As HyperLink = e.Item.FindControl("HyperLink1")
                btn.NavigateUrl = "javascript:show_dg();"
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim MyLink As LinkButton = e.Item.FindControl("LinkButton1")
                MyLabel = e.Item.FindControl("Label1")
                MyLink.Text = drv("DistName").ToString
                MyLink.ForeColor = Color.Blue
                MyLink.CommandArgument = drv("DistID")
                MyLink.ToolTip = "點選可以觀看詳細資料"
                If drv("advance_class") = 0 Then
                    MyLink.Attributes("onclick") = "return false;"
                    MyLink.ForeColor = Color.Black
                    MyLink.ToolTip = "此轄區無詳細資料"
                End If
                '年度總預算數(元)
                e.Item.Cells(2).Text = Format(CDbl(dr("advance_total")), "#,##0.00")
                '實際開班經費(元)
                e.Item.Cells(4).Text = Format(CDbl(MyLabel.Text), "#,##0.00")
                '結餘金額:年度總預算數(-實際開班經費)
                e.Item.Cells(5).Text = Format(CDbl(Convert.ToDouble(e.Item.Cells(2).Text) - (Convert.ToDouble(MyLabel.Text))), "#,##0.00")
            Case ListItemType.Footer
                For i As Integer = 1 To DG_Grid1.Columns.Count - 1
                    e.Item.Cells(i).Text = 0
                    For Each Item As DataGridItem In DG_Grid1.Items
                        If i = 1 Or i = 3 Then
                            e.Item.Cells(i).Text = Int(e.Item.Cells(i).Text) + Int(Item.Cells(i).Text)
                        Else
                            e.Item.Cells(i).Text = Format(CDbl(Convert.ToDouble(e.Item.Cells(i).Text) + Convert.ToDouble(Item.Cells(i).Text)), "#,##0.00")
                        End If
                    Next
                Next
        End Select
    End Sub

    Private Sub DG_Grid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DG_Grid1.ItemCommand
        TIMS.Utl_Redirect1(Me, "CM_01_002_Detail2.aspx?ID=" & Request("ID") & "&DistID=" & e.CommandArgument & "&year=" & syear.SelectedValue & "&type=T")
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Dim dr As DataRowView
        dr = e.Item.DataItem

        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                e.Item.Cells(1).Text = Format(CDbl(dr("cancel_total")), "#,##0.00")
                e.Item.Cells(2).Text = Format(CDbl(dr("money_all")), "#,##0.00")
            Case ListItemType.Footer
                For i As Integer = 1 To DataGrid1.Columns.Count - 1
                    e.Item.Cells(i).Text = 0
                    For Each Item As DataGridItem In DataGrid1.Items
                        e.Item.Cells(i).Text = Format(CDbl(Convert.ToDouble(e.Item.Cells(i).Text) + Convert.ToDouble(Item.Cells(i).Text)), "#,##0.00")
                    Next
                Next
        End Select
    End Sub
End Class