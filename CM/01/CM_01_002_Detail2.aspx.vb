Partial Class CM_01_002_Detail2
    Inherits AuthBasePage

    Dim re_year As String = ""
    Dim re_distid As String = ""
    Dim re_type As String = ""
    Dim re_tplanid As String = ""
    Dim re_rid As String = ""
    Dim re_sort As String = ""

    'Dim sqlstr As String = ""
    Dim objconn As SqlConnection
    Dim flag_ROC As Boolean = TIMS.CHK_REPLACE2ROC_YEARS()

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

        'If sm.UserInfo.RoleID <> 0 Then
        'End If
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

        re_year = Request("year")
        re_distid = Request("DistID")
        re_type = Request("type")
        re_tplanid = Request("TPlanID")
        re_rid = Request("rid")
        re_sort = Request("sort")
        re_year = TIMS.ClearSQM(re_year)
        re_distid = TIMS.ClearSQM(re_distid)
        re_type = TIMS.ClearSQM(re_type)
        re_tplanid = TIMS.ClearSQM(re_tplanid)
        re_rid = TIMS.ClearSQM(re_rid)
        re_sort = TIMS.ClearSQM(re_sort)

        'Dim Sqlstr As String = "select * from Plan_CostItem"
        'CostTable = DbAccess.GetDataTable(Sqlstr)
        If Not Page.IsPostBack Then create1()
    End Sub

    Sub create1()
        Dim sqlstr As String = ""
        sqlstr = " SELECT name FROM ID_District WHERE distid = '" & re_distid & "' "
        Dim dr As DataRow = DbAccess.GetOneRow(sqlstr, objconn)
        Dim list As String = ""

        Select Case re_type
            Case "T" '訓練計畫
                Label1.Text = IIf(flag_ROC, (CInt(re_year) - 1911).ToString, re_year) & "年" & dr("name")  'edit，by:20181022
                loaddata()
            Case "O" '機構
                Select Case re_sort
                    Case "0"
                        list = "公務"
                    Case "1"
                        list = "公務_就安"
                    Case "2"
                        list = "公務_就安_就保"
                    Case "3"
                        list = "公務_就保"
                    Case "4"
                        list = "就安"
                    Case "5"
                        list = "就安_就保"
                    Case "6"
                        list = "就保"
                End Select
                Label1.Text = IIf(flag_ROC, (CInt(re_year) - 1911).ToString, re_year) & "年" & dr("name") & "   " & Request("planname") & "(" & list & ")"  'edit，by:20181022
                loaddata2()
            Case "C"
                Select Case re_sort 'Request("sort")
                    Case "0"
                        list = "公務"
                    Case "1"
                        list = "公務_就安"
                    Case "2"
                        list = "公務_就安_就保"
                    Case "3"
                        list = "公務_就保"
                    Case "4"
                        list = "就安"
                    Case "5"
                        list = "就安_就保"
                    Case "6"
                        list = "就保"
                End Select
                Label1.Text = IIf(flag_ROC, (CInt(re_year) - 1911).ToString, re_year) & "年" & dr("name") & "   " & Request("planname") & "(" & list & ")"  'edit，by:20181022
                Label2.Text = "訓練機構：" & Request("orgname")
                loaddata3()
        End Select
    End Sub

    '訓練計畫
    Sub loaddata()
        Dim dt As DataTable
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " SELECT k.TPlanID ,k.PlanName " & vbCrLf
        sql += " 	,ISNULL(SUM(advance_class),0) advance_class " & vbCrLf
        sql += " 	,ISNULL(SUM(advance_total),0) advance_total " & vbCrLf
        sql += " 	,ISNULL(SUM(real_class),0) real_class " & vbCrLf
        sql += " 	,ISNULL(SUM(real_total),0) real_total " & vbCrLf
        sql += " 	,CASE WHEN SUM(b1_budget)>0 THEN '1' ELSE '0' END AS b1_budget " & vbCrLf
        sql += " 	,CASE WHEN SUM(b2_budget)>0 THEN '1' ELSE '0' END AS b2_budget " & vbCrLf
        sql += " 	,CASE WHEN SUM(b3_budget)>0 THEN '1' ELSE '0' END AS b3_budget " & vbCrLf
        sql += "    ,CASE " & vbCrLf
        sql += " 	   WHEN SUM(b1_budget)>0 AND SUM(b2_budget)=0 AND SUM(b3_budget)=0 THEN 0 " & vbCrLf
        sql += " 	   WHEN SUM(b1_budget)>0 AND SUM(b2_budget)>0 AND SUM(b3_budget)=0 THEN 1 " & vbCrLf
        sql += " 	   WHEN SUM(b1_budget)>0 AND SUM(b2_budget)>0 AND SUM(b3_budget)>0 THEN 2 " & vbCrLf
        sql += " 	   WHEN SUM(b1_budget)>0 AND SUM(b2_budget)=0 AND SUM(b3_budget)>0 THEN 3 " & vbCrLf
        sql += " 	   WHEN SUM(b1_budget)=0 AND SUM(b2_budget)>0 AND SUM(b3_budget)=0 THEN 4 " & vbCrLf
        sql += " 	   WHEN SUM(b1_budget)=0 AND SUM(b2_budget)>0 AND SUM(b3_budget)>0 THEN 5 " & vbCrLf
        sql += " 	   WHEN SUM(b1_budget)=0 AND SUM(b2_budget)=0 AND SUM(b3_budget)>0 THEN 6 " & vbCrLf
        sql += " 	   END AS SORT " & vbCrLf
        sql += " FROM key_Plan k " & vbCrLf
        sql += " 	 JOIN (" & vbCrLf
        sql += " 	SELECT p.TPlanID ,ISNULL(a.DefGovCost,0) + ISNULL(a.DefUnitCost,0) + ISNULL(a.DefStdCost,0) AS advance_total " & vbCrLf
        sql += " 		,CASE WHEN a.planid IS NOT NULL THEN 1 ELSE 0 END AS advance_class " & vbCrLf
        sql += " 		,CASE WHEN (c.ocid IS NOT NULL) AND (c.NotOpen='N' AND c.IsSuccess='Y' AND c.STDate <= dbo.TRUNC_DATETIME(GETDATE())) THEN 1 ELSE 0 END AS real_class " & vbCrLf
        sql += " 		,CASE WHEN (c.ocid IS NOT NULL) AND (c.NotOpen='N' AND c.IsSuccess='Y' AND c.STDate <= dbo.TRUNC_DATETIME(GETDATE())) THEN ISNULL(a.DefGovCost,0) + ISNULL(a.DefUnitCost,0) + ISNULL(a.DefStdCost,0) ELSE 0 END AS real_total " & vbCrLf
        sql += " 		,CASE WHEN b1.d IS NOT NULL THEN 1 ELSE 0 END AS b1_budget " & vbCrLf
        sql += " 		,CASE WHEN b2.d IS NOT NULL THEN 1 ELSE 0 END AS b2_budget " & vbCrLf
        sql += " 		,CASE WHEN b3.d IS NOT NULL THEN 1 ELSE 0 END AS b3_budget " & vbCrLf
        sql += " 	FROM Plan_PlanInfo a " & vbCrLf
        sql += "    JOIN ID_Plan p ON a.planid = p.planid " & vbCrLf
        sql += "         AND a.Planyear = '" & IIf(flag_ROC, (CInt(re_year) + 1911).ToString, re_year) & "' " & vbCrLf  'edit，by:20181022
        sql += "         AND p.DistID = '" & re_distid & "' " & vbCrLf
        sql += " 	LEFT JOIN Class_ClassInfo c ON a.planid = c.planid AND a.ComIDNO = c.ComIDNO AND a.SeqNO = c.SeqNO AND a.rid = c.rid " & vbCrLf
        sql += " 	LEFT JOIN (SELECT 'x' d, mb.TPlanID,mb.Syear FROM plan_budget mb WHERE mb.budid = '01') b1 ON p.TPlanID = b1.TPlanID AND p.Years = b1.Syear " & vbCrLf
        sql += " 	LEFT JOIN (SELECT 'x' d, mb.TPlanID,mb.Syear FROM plan_budget mb WHERE mb.budid = '02') b2 ON p.TPlanID = b2.TPlanID AND p.Years = b2.Syear " & vbCrLf
        sql += " 	LEFT JOIN (SELECT 'x' d, mb.TPlanID,mb.Syear FROM plan_budget mb WHERE mb.budid = '03') b3 ON p.TPlanID = b3.TPlanID AND p.Years = b3.Syear " & vbCrLf
        sql += " 	WHERE a.AppliedResult = 'Y' ) g ON g.TPlanID = k.TPlanID " & vbCrLf
        sql += " GROUP BY k.tplanid,k.PlanName" & vbCrLf
        sql += " ORDER BY k.tplanid" & vbCrLf

        dt = DbAccess.GetDataTable(sql, objconn)
        DG_Grid1.DataSource = dt
        DG_Grid1.DataBind()
        DG_Grid1.Visible = True
    End Sub

    '機構
    Sub loaddata2()
        Dim dt As DataTable

        '/*年度預定開班 */
        '/*年度預定總預算 */
        '/*實際開班數*/
        '/*實際總經費*/
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " SELECT org.orgname ,auth.rid " & vbCrLf
        sql += "        ,ISNULL(a.advance_class,0) advance_class " & vbCrLf
        sql += "        ,ISNULL(a.advance_total,0) advance_total " & vbCrLf
        sql += "        ,ISNULL(c.real_class,0) real_class " & vbCrLf
        sql += "        ,ISNULL(c.real_total,0) real_total " & vbCrLf
        sql += " FROM Auth_Relship auth " & vbCrLf
        sql += " JOIN org_orginfo org ON auth.orgid = org.orgid " & vbCrLf
        sql += " JOIN (" & vbCrLf
        sql += "   SELECT a.rid ,COUNT(1) advance_class ,sum(ISNULL(a.DefGovCost,0) + ISNULL(a.DefUnitCost,0) + ISNULL(a.DefStdCost,0)) advance_total " & vbCrLf
        sql += "   FROM Plan_PlanInfo a " & vbCrLf
        sql += "   JOIN id_plan b ON a.planid = b.planid AND a.AppliedResult = 'Y' " & vbCrLf
        sql += "       AND a.Planyear = '" & IIf(flag_ROC, (CInt(re_year) + 1911).ToString, re_year) & "' " & vbCrLf  'edit，by:20181022
        sql += "       AND a.tplanid = '" & re_tplanid & "' " & vbCrLf
        sql += "       AND b.distid = '" & re_distid & "' " & vbCrLf
        sql += "   GROUP BY a.rid " & vbCrLf
        sql += " ) a ON a.rid = auth.rid " & vbCrLf
        sql += " LEFT JOIN (" & vbCrLf
        sql += "   SELECT a.rid ,COUNT(1) AS real_class ,SUM(ISNULL(DefGovCost,0) + ISNULL(DefUnitCost,0) + ISNULL(DefStdCost,0)) AS real_total " & vbCrLf
        sql += "   FROM Plan_PlanInfo a " & vbCrLf
        sql += "   JOIN id_plan b ON a.planid = b.planid AND a.AppliedResult = 'Y' " & vbCrLf
        sql += "   JOIN Class_ClassInfo c ON a.planid = c.planid AND a.ComIDNO = c.ComIDNO AND a.SeqNO = c.SeqNO AND a.rid = c.rid " & vbCrLf
        sql += "        AND c.NotOpen = 'N'  " & vbCrLf
        sql += "        AND c.IsSuccess = 'Y'  " & vbCrLf
        sql += "        AND c.STDate <= dbo.TRUNC_DATETIME(GETDATE())" & vbCrLf
        sql += "   WHERE 1=1" & vbCrLf
        sql += "       AND a.Planyear = '" & IIf(flag_ROC, (CInt(re_year) + 1911).ToString, re_year) & "' " & vbCrLf  'edit，by:20181022
        sql += "       AND a.tplanid = '" & re_tplanid & "' " & vbCrLf
        sql += "       AND b.distid = '" & re_distid & "' " & vbCrLf
        sql += "   GROUP BY a.rid " & vbCrLf
        sql += " ) c ON c.rid = auth.rid " & vbCrLf

        dt = DbAccess.GetDataTable(sql, objconn)
        DataGrid1.DataSource = dt
        DataGrid1.DataBind()
        DataGrid1.Visible = True
    End Sub

    '班級
    Sub loaddata3()
        'Dim planid_str As String = ""
        Dim dt As DataTable
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " WITH WAA AS ( " & vbCrLf
        sql &= "   SELECT a.ocid ,SUM(ISNULL(b.DefGovCost,0) + ISNULL(b.DefUnitCost,0) + ISNULL(b.DefStdCost,0)) no1 " & vbCrLf
        sql &= "   FROM class_classinfo a " & vbCrLf
        sql &= "   JOIN plan_planinfo b ON a.planid = b.planid AND a.comidno = b.comidno AND a.seqno = b.seqno " & vbCrLf
        sql &= "   JOIN id_plan ip ON ip.planid = a.planid " & vbCrLf
        sql &= "   WHERE 1=1 " & vbCrLf
        sql &= "      AND a.IsSuccess = 'Y' " & vbCrLf
        sql &= "      AND a.NotOpen = 'N' " & vbCrLf
        sql &= "      AND a.STDate <= dbo.TRUNC_DATETIME(GETDATE()) " & vbCrLf
        sql += "      AND b.Planyear = '" & IIf(flag_ROC, (CInt(re_year) + 1911).ToString, re_year) & "' " & vbCrLf  'edit，by:20181022
        sql += "      AND b.TPLANID = '" & re_tplanid & "' " & vbCrLf
        sql += "      AND a.rid = '" & re_rid & "' " & vbCrLf
        sql += "      AND b.rid = '" & re_rid & "' " & vbCrLf
        sql &= "   GROUP BY a.ocid " & vbCrLf
        sql &= " )" & vbCrLf
        sql &= " ,WBB AS ( " & vbCrLf
        sql &= "   SELECT a.ocid " & vbCrLf
        sql &= "          ,SUM(CASE WHEN cs.budgetid='01' THEN 1 ELSE 0 END) x01 " & vbCrLf
        sql &= "          ,SUM(CASE WHEN cs.budgetid='02' THEN 1 ELSE 0 END) x02 " & vbCrLf
        sql &= "          ,SUM(CASE WHEN cs.budgetid='03' THEN 1 ELSE 0 END) x03 " & vbCrLf
        sql &= "   FROM class_classinfo a " & vbCrLf
        sql &= "   JOIN plan_planinfo b ON a.planid = b.planid AND a.comidno = b.comidno AND a.seqno = b.seqno " & vbCrLf
        sql &= "   JOIN class_studentsofclass cs ON cs.ocid = a.ocid " & vbCrLf
        sql &= "   JOIN id_plan ip ON ip.planid = a.planid " & vbCrLf
        sql &= "   WHERE 1=1 " & vbCrLf
        sql &= "     AND a.IsSuccess = 'Y' " & vbCrLf
        sql &= "     AND a.NotOpen = 'N' " & vbCrLf
        sql &= "     AND a.STDate <= dbo.TRUNC_DATETIME(GETDATE()) " & vbCrLf
        sql += "     AND b.Planyear = '" & IIf(flag_ROC, (CInt(re_year) + 1911).ToString, re_year) & "' " & vbCrLf  'edit，by:20181022
        sql += "     AND b.TPLANID = '" & re_tplanid & "' " & vbCrLf
        sql += "     AND a.rid = '" & re_rid & "' " & vbCrLf
        sql += "     AND b.rid = '" & re_rid & "' " & vbCrLf
        sql &= "   GROUP BY a.ocid " & vbCrLf
        sql &= " ) " & vbCrLf
        sql &= " ,WCC AS ( " & vbCrLf
        sql &= "   select a.ocid ,SUM(c.CancelCost) TotalCancelCost " & vbCrLf
        sql &= "   FROM class_classinfo a " & vbCrLf
        sql &= "   JOIN plan_planinfo b ON a.planid = b.planid AND a.comidno = b.comidno AND a.seqno = b.seqno " & vbCrLf
        sql &= "   JOIN Budget_ClassCancel c ON c.ocid = a.ocid " & vbCrLf
        sql &= "   JOIN id_plan ip ON ip.planid = a.planid " & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        sql &= "    AND a.IsSuccess = 'Y' " & vbCrLf
        sql &= "    AND a.NotOpen = 'N' " & vbCrLf
        sql &= "    AND a.STDate <= dbo.TRUNC_DATETIME(GETDATE()) " & vbCrLf
        sql += "    AND b.Planyear = '" & IIf(flag_ROC, (CInt(re_year) + 1911).ToString, re_year) & "' " & vbCrLf  'edit，by:20181022
        sql += "    AND b.TPLANID = '" & re_tplanid & "' " & vbCrLf
        sql += "    AND a.rid = '" & re_rid & "' " & vbCrLf
        sql += "    AND b.rid = '" & re_rid & "' " & vbCrLf
        sql &= " GROUP BY a.ocid" & vbCrLf
        sql &= " )" & vbCrLf
        sql &= " SELECT a.OCID" & vbCrLf
        sql &= "    ,dbo.FN_GET_CLASSCNAME(a.CLASSCNAME,a.CYCLTYPE) CLASSCNAME" & vbCrLf
        sql &= "    ,a.STDate ,a.FTDate ,a.THours ,a.Tnum " & vbCrLf
        sql &= "    ,ISNULL(cc.TotalCancelCost,0) TotalCancelCost ,d.orgname ,a.planid ,a.ComIDNO ,a.SeqNo " & vbCrLf
        sql &= "    ,ISNULL(aa.no1,0) no1" & vbCrLf
        sql &= "    ,ISNULL(bb.x01,0) x01" & vbCrLf
        sql &= "    ,ISNULL((bb.x01*aa.no1),0) no2" & vbCrLf
        sql &= "    ,ISNULL(bb.x02,0) x02" & vbCrLf
        sql &= "    ,ISNULL((bb.x02*aa.no1),0) no3" & vbCrLf
        sql &= "    ,ISNULL(bb.x03,0) x03" & vbCrLf
        sql &= "    ,ISNULL((bb.x03*aa.no1),0) no4" & vbCrLf
        sql &= " FROM class_classinfo a " & vbCrLf
        sql &= " JOIN plan_planinfo b ON a.planid = b.planid AND a.comidno = b.comidno AND a.seqno = b.seqno " & vbCrLf
        sql &= " JOIN Org_OrgInfo d ON d.comidno = a.comidno " & vbCrLf
        sql &= " JOIN id_plan ip ON ip.planid = a.planid " & vbCrLf
        sql &= " JOIN WAA aa ON aa.ocid = a.ocid " & vbCrLf
        sql &= " LEFT join WBB bb ON bb.ocid = a.ocid " & vbCrLf
        sql &= " LEFT join WCC cc ON cc.ocid = a.ocid " & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        sql &= "    AND a.IsSuccess = 'Y'" & vbCrLf
        sql &= "    AND a.NotOpen = 'N'" & vbCrLf
        sql &= "    AND a.STDate <= dbo.TRUNC_DATETIME(GETDATE()) " & vbCrLf
        sql += "    AND b.Planyear = '" & IIf(flag_ROC, (CInt(re_year) + 1911).ToString, re_year) & "' " & vbCrLf  'edit，by:20181022
        sql += "    AND b.TPLANID = '" & re_tplanid & "' " & vbCrLf
        sql += "    AND a.rid = '" & re_rid & "' " & vbCrLf
        sql += "    AND b.rid = '" & re_rid & "' " & vbCrLf
        sql += "    AND ip.distid = '" & re_distid & "' " & vbCrLf
        sql += " ORDER BY CLASSCNAME " & vbCrLf
        dt = DbAccess.GetDataTable(sql, objconn)
        'For Each dr As DataRow In dt.Rows
        '    planid_str = dr("PlanID")
        '    Exit For
        'Next
        DataGrid2.DataSource = dt
        DataGrid2.DataBind()
        DataGrid2.Visible = True

        '依照plan_budbet判斷顯示哪一種預算別
        Dim bugid_str As String = "SELECT BUDID FROM PLAN_BUDGET WHERE TPLANID=" & re_tplanid & " and syear='" & re_year & "'" & vbCrLf
        Dim datatable As DataTable = DbAccess.GetDataTable(bugid_str, objconn)
        For Each dr As DataRow In datatable.Rows
            Select Case dr("budid")
                Case "01" '公務
                    DataGrid2.Columns(7).Visible = Not DataGrid2.Columns(7).Visible
                Case "02" '就安
                    DataGrid2.Columns(5).Visible = Not DataGrid2.Columns(5).Visible
                Case "03" '就保
                    DataGrid2.Columns(6).Visible = Not DataGrid2.Columns(6).Visible
            End Select
        Next
    End Sub

    Private Sub DG_Grid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DG_Grid1.ItemDataBound
        Dim dr As DataRowView
        dr = e.Item.DataItem
        Select Case e.Item.ItemType
            Case ListItemType.Header
                Dim Head As DataGridItem = e.Item
                Dim cell As TableCell
                Head.Cells.Clear()
                cell = New TableCell
                cell.Text = "公務"
                Head.Cells.Add(cell)
                cell = New TableCell
                cell.Text = "就安"
                Head.Cells.Add(cell)
                cell = New TableCell
                cell.Text = "就保"
                Head.Cells.Add(cell)
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim MyLink As LinkButton = e.Item.FindControl("LinkButton1")
                MyLink.Text = drv("planname").ToString
                MyLink.ForeColor = Color.Blue
                MyLink.CommandArgument = "planname=" & drv("planname") & "&sort=" & drv("sort") & "&TPlanID=" & drv("TPlanID")
                MyLink.ToolTip = "點選可以觀看詳細資料"
                If drv("advance_class") = 0 Then
                    MyLink.Attributes("onclick") = "return false;"
                    MyLink.ForeColor = Color.Black
                    MyLink.ToolTip = "此轄區無詳細資料"
                End If
                ' 0 公務 ' 1 就安 ' 2 就保
                Const cst_公務 As Integer = 0
                Const cst_就安 As Integer = 1
                Const cst_就保 As Integer = 2

                Select Case dr("b1_budget")
                    Case 1
                        e.Item.Cells(cst_公務).Text = "◎"
                End Select

                Select Case dr("b2_budget")
                    Case 1
                        e.Item.Cells(cst_就安).Text = "◎"
                End Select

                Select Case dr("b3_budget")
                    Case 1
                        e.Item.Cells(cst_就保).Text = "◎"
                End Select

#Region "(No Use)"

                'Select Case dr("sort")
                '    Case "0"
                '        e.Item.Cells(cst_公務).Text = "◎"
                '    Case "1"
                '        e.Item.Cells(cst_公務).Text = "◎"
                '        e.Item.Cells(cst_就安).Text = "◎"
                '    Case "2"
                '        e.Item.Cells(cst_公務).Text = "◎"
                '        e.Item.Cells(cst_就安).Text = "◎"
                '        e.Item.Cells(cst_就保).Text = "◎"
                '    Case "3"
                '        e.Item.Cells(cst_公務).Text = "◎"
                '        e.Item.Cells(cst_就保).Text = "◎"
                '    Case "4"
                '        e.Item.Cells(cst_就安).Text = "◎"
                '    Case "5"
                '        e.Item.Cells(cst_就安).Text = "◎"
                '        e.Item.Cells(cst_就保).Text = "◎"
                '    Case "6"
                '        e.Item.Cells(cst_就保).Text = "◎"
                'End Select

#End Region

                '年度總預算數(元)
                e.Item.Cells(5).Text = Format(CDbl(dr("advance_total")), "#,##0.00")
                '實際開班經費(元)
                e.Item.Cells(7).Text = Format(CDbl(dr("real_total")), "#,##0.00")
                '結餘金額:年度總預算數(-實際開班經費)
                e.Item.Cells(8).Text = Format(CDbl(Convert.ToDouble(e.Item.Cells(5).Text) - (Convert.ToDouble(e.Item.Cells(7).Text))), "#,##0.00")

            Case ListItemType.Footer
                For i As Integer = 4 To DG_Grid1.Columns.Count - 1
                    e.Item.Cells(i).Text = 0
                    For Each Item As DataGridItem In DG_Grid1.Items
                        If i = 4 Or i = 6 Then
                            e.Item.Cells(i).Text = Int(e.Item.Cells(i).Text) + Int(Item.Cells(i).Text)
                        Else
                            e.Item.Cells(i).Text = Format(CDbl(Convert.ToDouble(e.Item.Cells(i).Text) + Convert.ToDouble(Item.Cells(i).Text)), "#,##0.00")
                        End If
                    Next
                Next
        End Select
    End Sub

    Private Sub DG_Grid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DG_Grid1.ItemCommand
        TIMS.Utl_Redirect1(Me, "CM_01_002_Detail2.aspx?ID=" & Request("ID") & "&DistID=" & re_distid & "&year=" & re_year & "&type=O&" & e.CommandArgument)
    End Sub

    Private Sub DG_Grid1_ItemCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DG_Grid1.ItemCreated
        Select Case e.Item.ItemType
            Case ListItemType.Pager
                Dim Head As DataGridItem = e.Item
                Dim cell As TableCell
                Head.Cells.Clear()
                Head.Attributes("Class") = "head_navy"
                'Head.BackColor = Color.FromName("#CC6666")
                'Head.ForeColor = Color.White

                cell = New TableCell
                cell.Text = "預算別"
                cell.ColumnSpan = 3
                Head.Cells.Add(cell)

                cell = New TableCell
                cell.RowSpan = 2
                cell.Text = "訓練計畫"
                Head.Cells.Add(cell)

                cell = New TableCell
                cell.RowSpan = 2
                cell.Text = "年度預定開班數"
                Head.Cells.Add(cell)

                cell = New TableCell
                cell.RowSpan = 2
                cell.Text = "年度總預算數(元)"
                Head.Cells.Add(cell)

                cell = New TableCell
                cell.RowSpan = 2
                cell.Text = "實際開班數"
                Head.Cells.Add(cell)

                cell = New TableCell
                cell.RowSpan = 2
                cell.Text = "實際開班經費(元)"
                Head.Cells.Add(cell)

                cell = New TableCell
                cell.RowSpan = 2
                cell.Text = "結餘金額(元)"
                Head.Cells.Add(cell)
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Dim dr As DataRowView
        dr = e.Item.DataItem

        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                e.Item.Cells(0).Text = e.Item.ItemIndex + 1
                '年度總預算數(元)
                e.Item.Cells(3).Text = Format(CDbl(dr("advance_total")), "#,##0.00")
                '實際開班經費(元)
                e.Item.Cells(5).Text = Format(CDbl(dr("real_total")), "#,##0.00")
                '結餘金額:年度總預算數(-實際開班經費)
                e.Item.Cells(6).Text = Format(CDbl(Convert.ToDouble(e.Item.Cells(3).Text) - (Convert.ToDouble(e.Item.Cells(5).Text))), "#,##0.00")
                Dim drv As DataRowView = e.Item.DataItem
                Dim MyLink As LinkButton = e.Item.FindControl("LinkButton2")
                MyLink.Text = drv("orgname").ToString
                MyLink.ForeColor = Color.Blue
                MyLink.CommandArgument = "orgname=" & drv("orgname") & "&rid=" & drv("rid")
                MyLink.ToolTip = "點選可以觀看詳細資料"
                If drv("real_class") = 0 Or e.Item.Cells(4).Text = 0 Then
                    MyLink.Attributes("onclick") = "return false;"
                    MyLink.ForeColor = Color.Black
                    MyLink.ToolTip = "此轄區無詳細資料"
                End If
            Case ListItemType.Footer
                For i As Integer = 2 To DataGrid1.Columns.Count - 1
                    e.Item.Cells(i).Text = 0
                    For Each Item As DataGridItem In DataGrid1.Items
                        If i = 2 Or i = 4 Then
                            e.Item.Cells(i).Text = Int(e.Item.Cells(i).Text) + Int(Item.Cells(i).Text)
                        Else
                            e.Item.Cells(i).Text = Format(CDbl(Convert.ToDouble(e.Item.Cells(i).Text) + Convert.ToDouble(Item.Cells(i).Text)), "#,##0.00")
                        End If
                    Next
                Next
        End Select
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        TIMS.Utl_Redirect1(Me, "CM_01_002_Detail2.aspx?ID=" & Request("ID") & "&DistID=" & re_distid & "&year=" & re_year & "&TPlanID=" & re_tplanid & "&planname=" & Request("planname") & "&sort=" & Request("sort") & "&type=C&" & e.CommandArgument)
    End Sub

    Private Sub DataGrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        'Dim cost, sum_a, sum_b As Double
        Dim sum_a As Double = 0
        Dim sum_b As Double = 0
        Dim dr As DataRowView = e.Item.DataItem
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                e.Item.Cells(0).Text = e.Item.ItemIndex + 1
                e.Item.Cells(2).Text = "" & dr("STDate") & "<br>" & dr("FTDate") & ""
                '計畫人數/每人費用
                '每人費用
                'cost = Convert.ToDouble(dr("no1") / dr("Tnum")) '(no1/tnum) 
                e.Item.Cells(4).Text = "" & dr("Tnum") & "/" & Format(CDbl(dr("no1")), "#,##0.00") & ""
                '就安人數/金額
                e.Item.Cells(5).Text = "" & dr("x02") & "/" & Format(CDbl(dr("no1") * dr("x02")), "#,##0.00") & " "
                ''就保人數/金額
                e.Item.Cells(6).Text = "" & dr("x03") & "/" & Format(CDbl(dr("no1") * dr("x03")), "#,##0.00") & ""
                '公務人數/金額
                e.Item.Cells(7).Text = "" & dr("x01") & "/" & Format(CDbl(dr("no1") * dr("x01")), "#,##0.00") & ""
                '計畫總經費
                e.Item.Cells(8).Text = Format(CDbl(dr("no1") * dr("Tnum")), "#,##0.00")
                '已核銷總金額
                e.Item.Cells(9).Text = Format(CDbl(dr("TotalCancelCost")), "#,##0.00")
                '結餘總金額
                e.Item.Cells(10).Text = Format(CDbl((dr("no1") * dr("Tnum")) - dr("TotalCancelCost")), "#,##0.00")
            Case ListItemType.Footer
                For i As Integer = 4 To 7
                    e.Item.Cells(i).Text = 0
                    For Each item As DataGridItem In DataGrid2.Items
                        Dim MyArray As Array = Split(item.Cells(i).Text, "/")
                        sum_a = sum_a + Int(MyArray(0))
                        sum_b = sum_b + Convert.ToDouble(MyArray(1))
                    Next
                    e.Item.Cells(i).Text = "" & sum_a & " / " & Format(CDbl(Math.Round(sum_b, 2)), "#,##0.00") & ""
                    sum_a = 0
                    sum_b = 0
                Next

                For j As Integer = 8 To DataGrid2.Columns.Count - 1
                    e.Item.Cells(j).Text = 0
                    For Each item As DataGridItem In DataGrid2.Items
                        e.Item.Cells(j).Text = Format(CDbl(Math.Round(Convert.ToDouble(e.Item.Cells(j).Text), 2) + Math.Round(Convert.ToDouble(item.Cells(j).Text), 2)), "#,##0.00")
                    Next
                Next
        End Select
    End Sub

    Private Sub Button1_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.ServerClick
        Select Case re_type
            Case "T" '訓練計畫
                TIMS.Utl_Redirect1(Me, "CM_01_002.aspx?ID=" & Request("ID") & "&year=" & re_year & "&processtype=back")
            Case "O" '機構
                TIMS.Utl_Redirect1(Me, "CM_01_002_Detail2.aspx?ID=" & Request("ID") & "&DistID=" & re_distid & "&year=" & re_year & "&type=T")
            Case "C"
                TIMS.Utl_Redirect1(Me, "CM_01_002_Detail2.aspx?ID=" & Request("ID") & "&DistID=" & re_distid & "&year=" & re_year & "&type=O&TPlanID=" & re_tplanid & "&planname=" & Request("planname") & "&sort=" & Request("sort"))
        End Select
    End Sub
End Class