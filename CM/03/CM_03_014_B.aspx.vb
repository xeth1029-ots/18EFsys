Partial Class CM_03_014_B
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        If Not IsPostBack Then
            Call GetSessionSearch()
            Call Search1()
        End If

    End Sub

    '取出Session值
    Sub GetSessionSearch()
        Const Cst_MySearch As String = "_MySearch"
        Dim MyValue As String = ""

        If Not Session(Cst_MySearch) Is Nothing Then
            MyValue = TIMS.GetMyValue(Session(Cst_MySearch), "prgid")
            If MyValue = "CM_03_014" Then
                Me.hYears.Value = TIMS.GetMyValue(Session(Cst_MySearch), "Years")
                Me.hSTDate1.Value = TIMS.GetMyValue(Session(Cst_MySearch), "STDate1")
                Me.hSTDate2.Value = TIMS.GetMyValue(Session(Cst_MySearch), "STDate2")
                Me.hFTDate1.Value = TIMS.GetMyValue(Session(Cst_MySearch), "FTDate1")
                Me.hFTDate2.Value = TIMS.GetMyValue(Session(Cst_MySearch), "FTDate2")

                Me.hDistID1.Value = TIMS.GetMyValue(Session(Cst_MySearch), "DistID1")
                Me.hTPlanID1.Value = TIMS.GetMyValue(Session(Cst_MySearch), "TPlanID1")
                Me.hBudgetID.Value = TIMS.GetMyValue(Session(Cst_MySearch), "BudgetID")

                Me.Yearsb.Value = TIMS.GetMyValue(Session(Cst_MySearch), "Yearsb")
                Me.DistIDb.Value = TIMS.GetMyValue(Session(Cst_MySearch), "DistIDb")
                Me.TPlanIDb.Value = TIMS.GetMyValue(Session(Cst_MySearch), "TPlanIDb")
                Me.PlanIDb.Value = TIMS.GetMyValue(Session(Cst_MySearch), "PlanIDb")
            End If
        End If

    End Sub

    ' （'回上一層的搜尋值）
    Sub KeepSearch3()
        Const Cst_MySearch As String = "_MySearch"
        Session(Cst_MySearch) = Nothing

        Dim MySearch As String = ""
        MySearch = "prgid=" & "CM_03_014"
        MySearch += "&Years=" & Me.hYears.Value
        MySearch += "&STDate1=" & Me.hSTDate1.Value
        MySearch += "&STDate2=" & Me.hSTDate2.Value
        MySearch += "&FTDate1=" & Me.hFTDate1.Value
        MySearch += "&FTDate2=" & Me.hFTDate2.Value
        MySearch += "&DistID1=" & hDistID1.Value
        MySearch += "&TPlanID1=" & hTPlanID1.Value
        MySearch += "&BudgetID=" & hBudgetID.Value

        'If ArgValue <> "" Then MySearch += ArgValue
        Session(Cst_MySearch) = MySearch
    End Sub

    'SQL查詢
    Sub Search1()
        'Dim dt As DataTable
        'Dim da As SqlDataAdapter = nothing
        'da = TIMS.GetOneDA()

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " select Vp.planname, vp.TPlanID, vp.Years " & vbCrLf
        sql += " ,vp.DistID ,id1.Name DistName ,Vp.PlanID,cc.ocid,cc.ClassCName " & vbCrLf
        sql += "  	,sum (dbo.NVL(gc.trainClass,0)) trainClass" & vbCrLf
        sql += "  	,sum (dbo.NVL(gs.trainNum,0)) trainNum" & vbCrLf
        sql += "  	,sum (dbo.NVL(gs.closeNum,0)) closeNum" & vbCrLf
        sql += "  	,sum (dbo.NVL(gs.JobNum,0)) JobNum" & vbCrLf
        sql += "  	,sum (dbo.NVL(gs.NJobNum,0)) NJobNum" & vbCrLf
        sql += "  	,sum (dbo.NVL(gs.xJobNum,0)) xJobNum" & vbCrLf
        sql += " from view_LoginPlan vp" & vbCrLf
        sql += " join key_plan kp on kp.TPlanID =vp.TPlanID " & vbCrLf
        sql += " join id_district id1 on id1.DistID =vp.DistID " & vbCrLf
        sql += " join class_classinfo cc on cc.PlanID=vp.PlanID " & vbCrLf
        sql += "  join (" & vbCrLf
        sql += "  	select ip.PlanID,ip.TPlanID,cc.OCID " & vbCrLf
        sql += "  	,count(*) trainClass " & vbCrLf
        sql += "  	from class_classinfo cc" & vbCrLf
        sql += "  	join plan_planinfo pp on pp.planid =cc.planid and pp.comidno =cc.comidno and pp.seqno =cc.seqno " & vbCrLf
        sql += "  	join id_plan ip on ip.planid=cc.planid" & vbCrLf
        sql += "  	where 1=1" & vbCrLf

        If Me.hYears.Value <> "" Then
            sql += " and ip.Years='" & Me.hYears.Value & "'" & vbCrLf
        End If
        If Me.hSTDate1.Value <> "" Then
            sql += " and cc.STDate>= " & TIMS.To_date(Me.hSTDate1.Value) & vbCrLf '" & Me.hSTDate1.Value & "'" & vbCrLf
        End If
        If Me.hSTDate2.Value <> "" Then
            sql += " and cc.STDate<= " & TIMS.To_date(Me.hSTDate2.Value) & vbCrLf '" & Me.hSTDate2.Value & "'" & vbCrLf
        End If
        If Me.hFTDate1.Value <> "" Then
            sql += " and cc.FTDate>= " & TIMS.To_date(Me.hFTDate1.Value) & vbCrLf '" & Me.hFTDate1.Value & "'" & vbCrLf
        End If
        If Me.hFTDate2.Value <> "" Then
            sql += " and cc.FTDate<= " & TIMS.To_date(Me.hFTDate2.Value) & vbCrLf '" & Me.hFTDate2.Value & "'" & vbCrLf
        End If
        If Me.hDistID1.Value <> "" Then
            sql += " and ip.DistID in (" & Me.hDistID1.Value & ")" & vbCrLf
        End If
        If Me.hTPlanID1.Value <> "" Then
            sql += " and ip.TPlanID in (" & Me.hTPlanID1.Value & ")" & vbCrLf
        End If

        If Me.Yearsb.Value <> "" Then
            sql += " and ip.Years='" & Me.Yearsb.Value & "'" & vbCrLf
        End If
        If Me.DistIDb.Value <> "" Then
            sql += " and ip.DistID in (" & Me.DistIDb.Value & ")" & vbCrLf
        End If
        If Me.TPlanIDb.Value <> "" Then
            sql += " and ip.TPlanID in (" & Me.TPlanIDb.Value & ")" & vbCrLf
        End If
        If Me.PlanIDb.Value <> "" Then
            sql += " and ip.PlanID in (" & Me.PlanIDb.Value & ")" & vbCrLf
        End If

        sql += "  	and exists (" & vbCrLf
        sql += "  		select 'x' from class_studentsofclass xcs where xcs.ocid =cc.ocid and xcs.MIdentityID='05'" & vbCrLf
        If Me.hBudgetID.Value <> "" Then
            sql += " and xcs.BudgetID in (" & Me.hBudgetID.Value & ")" & vbCrLf
        End If

        sql += "  	)" & vbCrLf
        sql += "  	group by ip.PlanID,ip.TPlanID,cc.OCID  " & vbCrLf
        sql += " ) gc on vp.PlanID=gc.PlanID AND cc.OCID=gc.OCID  " & vbCrLf
        sql += " left join (" & vbCrLf
        sql += "  	select ip.PlanID,ip.TPlanID,cc.OCID " & vbCrLf
        sql += "  	,count(*) trainNum " & vbCrLf
        sql += "  	,sum(case when cs.studstatus in (5) then 1 else 0 end) closeNum" & vbCrLf
        sql += "  	,sum(case when cs.studstatus in (5) and sg.IsGetJob='1' then 1 else 0 end) JobNum --1.就業" & vbCrLf
        sql += "  	,sum(case when cs.studstatus in (5) and sg.IsGetJob='2' then 1 else 0 end) NJobNum --2.不就業" & vbCrLf
        sql += "  	,sum(case when cs.studstatus in (5) and dbo.NVL(sg.IsGetJob,'99') not in ('1','2') then 1 else 0 end) xJobNum --x1x2.未就業" & vbCrLf
        sql += "  	from class_classinfo cc" & vbCrLf
        sql += "  	join plan_planinfo pp on pp.planid =cc.planid and pp.comidno =cc.comidno and pp.seqno =cc.seqno " & vbCrLf
        sql += "  	join id_plan ip on ip.planid=cc.planid" & vbCrLf
        sql += "  	join class_studentsofclass cs on cs.ocid =cc.ocid and cs.MIdentityID='05'" & vbCrLf
        sql += "  	left join Stud_GetJobState3 sg on sg.CPoint=1 and sg.socid =cs.socid " & vbCrLf
        sql += "  	where 1=1" & vbCrLf

        If Me.hYears.Value <> "" Then
            sql += " and ip.Years='" & Me.hYears.Value & "'" & vbCrLf
        End If
        If Me.hSTDate1.Value <> "" Then
            sql += " and cc.STDate>= " & TIMS.To_date(Me.hSTDate1.Value) & vbCrLf '" & Me.hSTDate1.Value & "'" & vbCrLf
        End If
        If Me.hSTDate2.Value <> "" Then
            sql += " and cc.STDate<= " & TIMS.To_date(Me.hSTDate2.Value) & vbCrLf '" & Me.hSTDate2.Value & "'" & vbCrLf
        End If
        If Me.hFTDate1.Value <> "" Then
            sql += " and cc.FTDate>= " & TIMS.To_date(Me.hFTDate1.Value) & vbCrLf '" & Me.hFTDate1.Value & "'" & vbCrLf
        End If
        If Me.hFTDate2.Value <> "" Then
            sql += " and cc.FTDate<= " & TIMS.To_date(Me.hFTDate2.Value) & vbCrLf '" & Me.hFTDate2.Value & "'" & vbCrLf
        End If
        If Me.hDistID1.Value <> "" Then
            sql += " and ip.DistID in (" & Me.hDistID1.Value & ")" & vbCrLf
        End If
        If Me.hTPlanID1.Value <> "" Then
            sql += " and ip.TPlanID in (" & Me.hTPlanID1.Value & ")" & vbCrLf
        End If
        If Me.hBudgetID.Value <> "" Then
            sql += " and cs.BudgetID in (" & Me.hBudgetID.Value & ")" & vbCrLf
        End If

        If Me.Yearsb.Value <> "" Then
            sql += " and ip.Years='" & Me.Yearsb.Value & "'" & vbCrLf
        End If
        If Me.DistIDb.Value <> "" Then
            sql += " and ip.DistID in (" & Me.DistIDb.Value & ")" & vbCrLf
        End If
        If Me.TPlanIDb.Value <> "" Then
            sql += " and ip.TPlanID in (" & Me.TPlanIDb.Value & ")" & vbCrLf
        End If
        If Me.PlanIDb.Value <> "" Then
            sql += " and ip.PlanID in (" & Me.PlanIDb.Value & ")" & vbCrLf
        End If

        sql += "  	and exists (" & vbCrLf
        sql += "  		select 'x' from class_studentsofclass xcs where xcs.ocid =cc.ocid and xcs.socid =cs.socid and xcs.MIdentityID='05'" & vbCrLf
        If Me.hBudgetID.Value <> "" Then
            sql += " and xcs.BudgetID in (" & Me.hBudgetID.Value & ")" & vbCrLf
        End If

        sql += "  	)" & vbCrLf
        sql += "  	group by ip.PlanID,ip.TPlanID,cc.OCID " & vbCrLf
        sql += "  ) gs on vp.PlanID=gs.PlanID AND cc.OCID=gs.OCID " & vbCrLf
        sql += "  where 1=1" & vbCrLf

        If Me.hYears.Value <> "" Then
            sql += " and vp.Years='" & Me.hYears.Value & "'" & vbCrLf
        End If
        If Me.hDistID1.Value <> "" Then
            sql += " and vp.DistID in (" & Me.hDistID1.Value & ")" & vbCrLf
        End If
        If Me.hTPlanID1.Value <> "" Then
            sql += " and vp.TPlanID in (" & Me.hTPlanID1.Value & ")" & vbCrLf
        End If
        If Me.PlanIDb.Value <> "" Then
            sql += " and vp.PlanID in (" & Me.PlanIDb.Value & ")" & vbCrLf
        End If

        sql += " GROUP BY " & vbCrLf
        sql += "   Vp.planname, vp.TPlanID, vp.Years " & vbCrLf
        sql += "  ,vp.DistID ,id1.Name ,Vp.PlanID,cc.ocid,cc.ClassCName " & vbCrLf
        sql += " ORDER BY " & vbCrLf
        sql += "   Vp.planname, vp.TPlanID, vp.Years " & vbCrLf
        sql += "  ,vp.DistID ,id1.Name ,Vp.PlanID,cc.ocid,cc.ClassCName " & vbCrLf

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)
        'Call TIMS.Fill(Sql, da, dt)
        'DataGrid1.AllowPaging = False
        DataGrid1.DataSource = dt
        DataGrid1.DataBind()
    End Sub

    '回上層鈕
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Call KeepSearch3()
        TIMS.Utl_Redirect1(Me, "CM_03_014_A.aspx?ID=" & Request("ID"))

    End Sub

    Private Sub DataGrid1_ItemCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemCreated
        Select Case e.Item.ItemType
            Case ListItemType.Pager
                Dim Head As DataGridItem = e.Item
                Dim cell As TableCell

                Head.Cells.Clear()

                Head.BackColor = Color.FromName("#CC6666")
                Head.ForeColor = Color.White

                cell = New TableCell
                cell.Text = "計畫名稱"
                cell.RowSpan = 2
                Head.Cells.Add(cell)

                cell = New TableCell
                cell.RowSpan = 2
                cell.Text = "班級名稱"
                Head.Cells.Add(cell)

                cell = New TableCell
                cell.RowSpan = 2
                cell.Text = "訓練人數"
                Head.Cells.Add(cell)

                cell = New TableCell
                cell.RowSpan = 2
                cell.Text = "結訓人數"
                Head.Cells.Add(cell)

                cell = New TableCell
                cell.ColumnSpan = 3 'ColSpan
                cell.Text = "訓後三個月內就業輔導情形"
                Head.Cells.Add(cell)

        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                Dim Head As DataGridItem = e.Item
                Dim cell As TableCell

                Head.Cells.Clear()

                cell = New TableCell
                cell.Text = "就業人數"
                Head.Cells.Add(cell)

                cell = New TableCell
                cell.Text = "不就業人數"
                Head.Cells.Add(cell)

                cell = New TableCell
                cell.Text = "未就業人數" '含未填寫
                Head.Cells.Add(cell)

                'cell = New TableCell
            Case ListItemType.Item, ListItemType.AlternatingItem
                'Dim drv As DataRowView = e.Item.DataItem
                'Dim LinkButton1 As LinkButton = e.Item.FindControl("LinkButton1")

                'LinkButton1.Text = ""
                'LinkButton1.Text &= drv("Years").ToString
                'LinkButton1.Text &= drv("planname").ToString
                'LinkButton1.ForeColor = Color.Blue
                'Dim ArgValue As String = ""
                'ArgValue = ""
                'ArgValue += "&planname=" & drv("planname")
                'ArgValue += "&TPlanID=" & drv("TPlanID")
                'ArgValue += "&Years=" & drv("Years")
                'LinkButton1.CommandArgument = ArgValue

            Case ListItemType.Footer
                For i As Integer = 1 To DataGrid1.Columns.Count - 1
                    e.Item.Cells(i).Text = 0
                    For Each Item As DataGridItem In DataGrid1.Items
                        If IsNumeric(Item.Cells(i).Text) Then
                            e.Item.Cells(i).Text = CInt(e.Item.Cells(i).Text) + CInt(Item.Cells(i).Text)
                        Else
                            e.Item.Cells(i).Text += 1
                            'Exit For
                        End If
                    Next
                Next

        End Select
    End Sub

End Class
