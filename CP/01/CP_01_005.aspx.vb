Partial Class CP_01_005
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        msg1.Text = ""
        If Not IsPostBack Then
            CreateItem()
            DataGridTable1.Visible = False
            Page2.Visible = False
            Page3.Visible = False
            EDate.Text = Now.Date
            Button9.Visible = False
        End If
        PageControler1.PageDataGrid = DataGrid1
        PageControler2.PageDataGrid = DataGrid2
        PageControler3.PageDataGrid = DataGrid3

        Button2.Attributes("onclick") = "return search();"
        Button7.Attributes("onclick") = "return check_data();"
        LinkButton3.Attributes("onclick") = "return HiddenTable('SearchTable','TableState1');"
        LinkButton2.Attributes("onclick") = "return HiddenTable('FilterTable','TableState2');"
        ShowMode.Attributes("onclick") = "ChangeMode();"

        If ShowMode.SelectedIndex = 0 Then
            ShowModeTable1.Style("display") = ""
            ShowModeTable2.Style("display") = "none"
        Else
            ShowModeTable1.Style("display") = "none"
            ShowModeTable2.Style("display") = ""
        End If
        If TableState1.Value = "0" Then
            SearchTable.Style("display") = "none"
        Else
            'SearchTable.Style("display")="inline"
            '上面為原寫法
            SearchTable.Style("display") = ""
        End If
        If TableState2.Value = "0" Then
            FilterTable.Style("display") = "none"
        Else
            FilterTable.Style("display") = ""
        End If
        '2006/03/28 add conn by matt
    End Sub

    Sub CreateItem()
        rblDistID = TIMS.Get_DistID(rblDistID)
        rblDistID.Items.Insert(0, New ListItem("全部"))
        If sm.UserInfo.RID = "A" Then
            rblDistID.SelectedIndex = 0
        Else
            TR1.Visible = False
            Common.SetListItem(rblDistID, sm.UserInfo.DistID)
        End If

        rblTPlanID = TIMS.Get_TPlan(rblTPlanID)
        rblTPlanID.Items.Insert(0, New ListItem("全部"))
        rblTPlanID.SelectedIndex = 0

        ddlYears = TIMS.GetSyear(ddlYears)
        Common.SetListItem(ddlYears, Now.Year)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        Dim v_ddlYears As String = TIMS.GetListValue(ddlYears)
        Dim v_rblTPlanID As String = TIMS.GetListValue(rblTPlanID)
        Dim v_rblDistID As String = TIMS.GetListValue(rblDistID)
        SOrgName.Text = TIMS.ClearSQM(SOrgName.Text)

        Dim sql As String = ""
        sql &= " SELECT ip.TPlanID" & vbCrLf
        sql &= " ,ip.PlanID" & vbCrLf
        sql &= " ,ip.Years" & vbCrLf
        sql &= " ,ip.PlanName" & vbCrLf
        sql &= " ,ip.DistID" & vbCrLf
        sql &= " ,ip.DistName" & vbCrLf
        sql &= " ,d.RID" & vbCrLf
        sql &= " ,d.OrgName" & vbCrLf
        sql &= " ,ISNULL(g.ClassCount,0) ClassCount" & vbCrLf
        sql &= " ,ISNULL(g.FinCount,0) FinCount" & vbCrLf
        sql &= " ,ISNULL(g.VisitCount,0) VisitCount" & vbCrLf
        sql &= " ,i.Mode1" & vbCrLf
        sql &= " ,i.Mode1Rate" & vbCrLf
        sql &= " ,j.GRate2" & vbCrLf
        sql &= " ,j.YRate2" & vbCrLf
        sql &= " ,j.RRate2" & vbCrLf
        sql &= " FROM VIEW_RIDNAME d" & vbCrLf
        sql &= " JOIN VIEW_PLAN ip ON ip.PlanID=d.PlanID" & vbCrLf

        sql &= " LEFT JOIN (" & vbCrLf
        sql &= " 	SELECT ip.PlanID ,d.RID" & vbCrLf
        sql &= " 	,COUNT(1) ClassCount" & vbCrLf
        sql &= " 	,SUM(CASE WHEN cc.FTDate <= GETDATE() THEN 1 END) FinCount" & vbCrLf
        sql &= " 	,SUM(CASE WHEN cc.FTDate <= GETDATE() AND h2.h2cnt > 0 THEN h2.h2cnt END) VisitCount" & vbCrLf
        sql &= " 	FROM CLASS_CLASSINFO cc" & vbCrLf
        sql &= " 	JOIN ID_PLAN ip ON ip.planid=cc.planid" & vbCrLf
        sql &= " 	JOIN VIEW_RIDNAME d ON d.RID=cc.RID" & vbCrLf
        sql &= " 	LEFT JOIN (SELECT ocid ,COUNT(1) h2cnt FROM VIEW_VISITOR h2 GROUP BY ocid) h2 ON h2.ocid=cc.ocid" & vbCrLf
        sql &= " 	WHERE 1=1" & vbCrLf
        sql &= " 	AND cc.NotOpen='N'" & vbCrLf
        sql &= " 	AND cc.IsSuccess='Y'" & vbCrLf
        '計畫年度機構。
        If v_rblTPlanID <> "" Then sql &= " AND ip.TPlanID='" & v_rblTPlanID & "'" & vbCrLf
        '計畫年度機構。
        If ddlYears.SelectedIndex <> 0 AndAlso v_ddlYears <> "" Then sql &= " AND ip.Years='" & v_ddlYears & "' " & vbCrLf
        '計畫年度分署。
        If rblDistID.SelectedIndex <> 0 AndAlso v_rblDistID <> "" Then sql &= " AND ip.DistID='" & v_rblDistID & "'" & vbCrLf
        '機構名
        If SOrgName.Text <> "" Then sql &= " AND d.OrgName LIKE '%" & SOrgName.Text & "%' "
        'sql &= " 	AND ip.Years='2019'" & vbCrLf
        'sql &= " 	AND ip.TPlanID='70'" & vbCrLf
        sql &= " 	GROUP BY ip.PlanID ,d.RID" & vbCrLf

        sql &= " ) g ON g.PlanID=ip.PlanID AND g.RID=d.RID" & vbCrLf
        sql &= " LEFT JOIN Sys_VisitRate i ON i.TPlanID=ip.TPlanID AND i.Years=ip.Years" & vbCrLf
        sql &= " CROSS JOIN Sys_VisitAlert j" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf

        '只顯示有查核資料()
        If chkVisirOnly.Checked Then sql &= " AND g.VisitCount > 0 " & vbCrLf
        '計畫年度機構。
        If v_rblTPlanID <> "" Then sql &= " AND ip.TPlanID='" & v_rblTPlanID & "'" & vbCrLf
        '計畫年度機構。
        If ddlYears.SelectedIndex <> 0 AndAlso v_ddlYears <> "" Then sql &= " AND ip.Years='" & v_ddlYears & "' " & vbCrLf
        '計畫年度分署。
        If rblDistID.SelectedIndex <> 0 AndAlso v_rblDistID <> "" Then sql &= " AND ip.DistID='" & v_rblDistID & "'" & vbCrLf
        '機構名
        If SOrgName.Text <> "" Then sql &= " AND d.OrgName LIKE '%" & SOrgName.Text & "%' "
        'sql &= " AND ip.Years='2019'" & vbCrLf
        'sql &= " AND ip.TPlanID='70'" & vbCrLf

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)

        DataGridTable1.Visible = False
        Button9.Visible = False
        msg1.Text = "查無資料"
        If dt.Rows.Count > 0 Then
            DataGridTable1.Visible = True
            Button9.Visible = True
            msg1.Text = ""

            PageControler1.PageDataTable = dt
            'PageControler1.Sort="RIDValue,ClassID,CyclType"
            PageControler1.ControlerLoad()
        End If
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = "head_navy"

            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim btn As LinkButton = e.Item.FindControl("LinkButton1")
                If e.Item.ItemType = ListItemType.Item Then e.Item.CssClass = "white_smoke"

                Dim v_ddlYears As String = TIMS.GetListValue(ddlYears)
                Dim v_rblTPlanID As String = TIMS.GetListValue(rblTPlanID)
                Dim v_rblDistID As String = TIMS.GetListValue(rblDistID)

                Const cst_v_ddlYears_2007 As String = " ddlYears_2007"
                Const cst_v_ddlYears_2008 As String = " ddlYears_2008"
                Dim str_v_ddlYears As String = ""
                If v_ddlYears <> "" Then
                    If v_ddlYears <= 2007 Then
                        str_v_ddlYears = cst_v_ddlYears_2007
                    Else
                        If v_ddlYears <= 2008 Then
                            str_v_ddlYears = cst_v_ddlYears_2008
                        End If
                    End If
                End If
                Select Case str_v_ddlYears
                    Case cst_v_ddlYears_2007
                        Select Case drv("Mode1").ToString
                            Case "1", "2"
                                If drv("Mode1").ToString = "1" Then
                                    e.Item.Cells(6).Text = drv("FinCount") * drv("Mode1Rate")
                                    e.Item.Cells(2).Text = "結訓班數*" & drv("Mode1Rate")
                                Else
                                    e.Item.Cells(6).Text = Math.Ceiling(drv("FinCount") * drv("Mode1Rate") / 100)
                                    e.Item.Cells(2).Text = "結訓班數*" & drv("Mode1Rate") & "%"
                                End If
                                If Int(drv("VisitCount")) < Int(e.Item.Cells(6).Text) Then
                                    e.Item.Cells(8).Text = "<font color=Red>●</font>"
                                Else
                                    e.Item.Cells(8).Text = "<font color=Green>●</font>"
                                End If
                            Case Else
                                e.Item.Cells(2).Text = "尚未設定參數"
                                e.Item.Cells(6).Text = "尚未設定參數"
                                e.Item.Cells(8).Text = "尚未設定參數"
                        End Select

                    Case cst_v_ddlYears_2008
                        Select Case drv("Mode1").ToString
                            Case "1", "2", "3"
                                If drv("Mode1").ToString = "1" Then
                                    e.Item.Cells(6).Text = Math.Ceiling(drv("FinCount") * drv("Mode1Rate") / 100)
                                    e.Item.Cells(2).Text = "結訓班數*" & drv("Mode1Rate") & "%"
                                End If
                                If Int(drv("VisitCount")) < Int(e.Item.Cells(6).Text) Then
                                    e.Item.Cells(8).Text = "<font color=Red>●</font>"
                                Else
                                    e.Item.Cells(8).Text = "<font color=Green>●</font>"
                                End If
                            Case Else
                                e.Item.Cells(2).Text = "尚未設定參數"
                                e.Item.Cells(6).Text = "尚未設定參數"
                                e.Item.Cells(8).Text = "尚未設定參數"
                        End Select

                    Case Else
                        Dim run_flag As Boolean = False '繼續執行
                        Select Case drv("Mode1").ToString
                            Case "1", "2", "3", "4"
                                run_flag = True
                            Case "5", "6", "7", "8"
                                run_flag = True
                            Case "9", "10", "11", "12", "13"
                                run_flag = True
                            Case Else
                                run_flag = False
                                e.Item.Cells(2).Text = "尚未設定參數"
                                e.Item.Cells(6).Text = "尚未設定參數"
                                e.Item.Cells(8).Text = "尚未設定參數"
                        End Select
                        If run_flag Then
                            Select Case drv("Mode1Rate").ToString '
                                Case "1", "2", "3", "4" '次數
                                    e.Item.Cells(6).Text = Math.Ceiling(drv("FinCount") * drv("Mode1Rate"))
                                    e.Item.Cells(2).Text = "結訓班數*" & drv("Mode1Rate") & "次"
                                    If Int(drv("VisitCount")) < Int(e.Item.Cells(6).Text) Then
                                        e.Item.Cells(8).Text = "<font color=Red>●</font>"
                                    Else
                                        e.Item.Cells(8).Text = "<font color=Green>●</font>"
                                    End If
                                Case "5", "6", "7", "8" ' 5%,15%,25%,50%
                                    e.Item.Cells(6).Text = Math.Ceiling(drv("FinCount") * drv("Mode1Rate") / 100)
                                    e.Item.Cells(2).Text = "結訓班數*" & drv("Mode1Rate") & "%"
                                    If Int(drv("VisitCount")) < Int(e.Item.Cells(6).Text) Then
                                        e.Item.Cells(8).Text = "<font color=Red>●</font>"
                                    Else
                                        e.Item.Cells(8).Text = "<font color=Green>●</font>"
                                    End If
                                Case Else
                                    e.Item.Cells(2).Text = "尚未設定參數"
                                    e.Item.Cells(6).Text = "尚未設定參數"
                                    e.Item.Cells(8).Text = "尚未設定參數"
                            End Select

                        End If
                End Select



                btn.Text = drv("OrgName").ToString
                btn.CommandArgument = ""
                btn.CommandArgument += "PlanID=" & drv("PlanID")
                btn.CommandArgument += "&RID=" & drv("RID")
                btn.CommandArgument += "&PlanYears=" & v_ddlYears ' Years.SelectedValue

                'If drv("VisitCount")=0 Then
                '    btn.Attributes("onclick")="return false;"
                '    btn.ForeColor=Color.Black
                'End If
        End Select
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Page1.Visible = False
        Page2.Visible = True
        Dim cmdArg As String = e.CommandArgument
        'Dim RID As String
        'Dim PlanID As String
        Me.ViewState("PlanID") = TIMS.GetMyValue(cmdArg, "PlanID") ' Split(Split(e.CommandArgument, "&")(0), "=")(1)
        Me.ViewState("RID") = TIMS.GetMyValue(cmdArg, "RID") 'Split(Split(e.CommandArgument, "&")(1), "=")(1)
        Me.ViewState("PlanYears") = TIMS.GetMyValue(e.CommandArgument, "PlanYears")
        CreateVisit()
        CreateDisVisit()
        CreateName(e.Item.ItemIndex, Me.ViewState("PlanYears"))
    End Sub

    Sub CreateName(ByVal num As Integer, ByVal PlanYears As Integer)
        Dim sql As String
        sql = " SELECT OrgName From Org_OrgInfo WHERE OrgID=(SELECT OrgID FROM Auth_Relship WHERE RID='" & Me.ViewState("RID") & "') "
        OrgName.Text = DbAccess.ExecuteScalar(sql, objconn)

        sql = " SELECT PlanName From Key_Plan WHERE TPlanID=(SELECT TPlanID FROM ID_Plan WHERE PlanID='" & Me.ViewState("PlanID") & "') "
        PlanName.Text = DbAccess.ExecuteScalar(sql, objconn)

        If PlanYears <= 2008 Then
            sql = "" & vbCrLf
            sql &= " SELECT CASE WHEN Mode1=1 THEN '依照次數訪查*' + CONVERT(varchar, Mode1Rate)  " & vbCrLf
            sql &= " WHEN Mode1=2 THEN '依照百分比訪查*' + CONVERT(varchar, Mode1Rate) + '%' END AS Mode1 " & vbCrLf
            sql &= " FROM Sys_VisitRate " & vbCrLf
            sql &= " WHERE TPlanID=(SELECT TPlanID FROM ID_Plan WHERE PlanID='" & Me.ViewState("PlanID") & "') " & vbCrLf
            sql &= " AND Years='" & PlanYears & "' " & vbCrLf
            CheckMode.Text = DbAccess.ExecuteScalar(sql, objconn)
        Else
            sql = "" & vbCrLf
            sql &= " SELECT CASE WHEN Mode1Rate IN (1,2,3,4) THEN '依照訪查次數*' + CONVERT(varchar, Mode1Rate) " & vbCrLf
            sql &= "  WHEN Mode1Rate=5 THEN '依照訪查百分比*' + CONVERT(varchar, 5) + '%' " & vbCrLf
            sql &= "  WHEN Mode1Rate=6 THEN '依照訪查百分比*' + CONVERT(varchar, 15) + '%' " & vbCrLf
            sql &= "  WHEN Mode1Rate=7 THEN '依照訪查百分比*' + CONVERT(varchar, 25) + '%' " & vbCrLf
            sql &= "  WHEN Mode1Rate=8 THEN '依照訪查百分比*' + CONVERT(varchar, 50) + '%' " & vbCrLf
            sql &= "  ELSE '未設定' " & vbCrLf
            sql &= "  END AS Mode1 " & vbCrLf
            sql &= " FROM Sys_VisitRate " & vbCrLf
            sql &= " WHERE TPlanID=(SELECT TPlanID FROM ID_Plan WHERE PlanID='" & Me.ViewState("PlanID") & "') " & vbCrLf
            sql &= " AND Years='" & PlanYears & "' " & vbCrLf

            CheckMode.Text = DbAccess.ExecuteScalar(sql, objconn)
        End If
        DeCount.Text = DataGrid1.Items(num).Cells(6).Text
        RelCount.Text = DataGrid1.Items(num).Cells(7).Text
        If CheckMode.Text = "" Then CheckMode.Text = "尚未設定參數"
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize2, DataGrid2)
        CreateVisit()
    End Sub

    '建立尚未訪查的資料
    Sub CreateDisVisit()
        Dim PMS1 As New Hashtable From {
            {"PlanID", TIMS.CINT1(ViewState("PlanID"))},
            {"RID", ViewState("RID")}
        }
        Dim sql As String = ""
        sql &= " SELECT a.OCID,a.PLANID,a.RID,a.CLASSCNAME2,a.STDATE,a.FTDATE FROM VIEW2 a" & vbCrLf
        sql &= " WHERE NOT EXISTS (SELECT 'x' FROM VIEW_VISITOR x WHERE x.OCID=a.OCID) " & vbCrLf
        sql &= " AND a.PlanID=@PlanID AND a.RID=@RID" & vbCrLf

        msg3.Text = "查無資料"
        DataGridTable3.Visible = False

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, PMS1)
        If TIMS.dtNODATA(dt) Then Return

        msg3.Text = ""
        DataGridTable3.Visible = True
        PageControler3.PageDataTable = dt
        PageControler3.PrimaryKey = "OCID"
        'PageControler3.Sort="OrgID,OCID,StudentID,LeaveDate"
        PageControler3.ControlerLoad()
#Region "(No Use)"

        'If TIMS.Get_SQLRecordCount(sql)=0 Then
        '    msg3.Text="查無資料"
        '    DataGridTable3.Visible=False
        'Else
        '    DataGridTable3.Visible=True
        '    PageControler3.SqlPrimaryKeyDataCreate(sql, "OCID")
        'End If

#End Region
    End Sub

    '建立以訪查的資料
    Sub CreateVisit()
        Dim sql As String = ""
        'Dim dt As DataTable
        Dim SearchStr As String = ""
        Dim WrongCount As String = ""

        If ddlYears.SelectedValue <= 2007 Then
            WrongCount = "CASE WHEN Data1 <> 1 THEN 1 ELSE 0 END + CASE WHEN Data2 <> 1 THEN 1 ELSE 0 END + CASE WHEN Data3 <> 1 THEN 1 ELSE 0 END + CASE WHEN Data4 <> 1 THEN 1 ELSE 0 END + "
            WrongCount += "CASE WHEN Data5 <> 1 THEN 1 ELSE 0 END + CASE WHEN Data6 <> 1 THEN 1 ELSE 0 END + CASE WHEN Data7 <> 1 THEN 1 ELSE 0 END + "
            WrongCount += "CASE WHEN Item1_1='Y' THEN 0 ELSE 1 END + CASE WHEN Item1_2='Y' THEN 0 ELSE 1 END + CASE WHEN Item2_1='Y' THEN 0 ELSE 1 END + CASE WHEN Item2_2='Y' THEN 0 ELSE 1 END + "
            WrongCount += "CASE WHEN Item3_1='Y' THEN 0 ELSE 1 END + CASE WHEN Item3_2='Y' THEN 0 ELSE 1 END + CASE WHEN Item4_1='Y' THEN 0 ELSE 1 END + CASE WHEN Item5_1='Y' THEN 0 ELSE 1 END + "
            WrongCount += "CASE WHEN Item6_1='Y' THEN 0 ELSE 1 END AS WrongCount "

            If SDate.Text <> "" Then SearchStr += " AND ApplyDate >= " & TIMS.To_date(SDate.Text) & vbCrLf
            If EDate.Text <> "" Then SearchStr += " AND ApplyDate <=" & TIMS.To_date(EDate.Text) & vbCrLf
            If IsClear.SelectedIndex <> 0 Then SearchStr += " AND IsClear='" & IsClear.SelectedValue & "' "

            'Y: <=2007
            sql = ""
            sql &= " SELECT * FROM ( "
            sql &= "   SELECT a.CyclType ,a.ClassCName ,a.FTDate ,b.OCID ,b.SeqNo ,b.ApplyDate ,b.WrongCount ,b.IsClear "
            sql &= "    ,CASE WHEN WrongCount/16.0*100 > c.RRate1 THEN 1 WHEN WrongCount/16.0*100 > c.YRate1 THEN 2 ELSE 3 END AS VisitResult "
            sql &= "   FROM (SELECT * FROM Class_ClassInfo WHERE RID='" & Me.ViewState("RID") & "' AND PlanID='" & Me.ViewState("PlanID") & "') a "
            sql &= "   JOIN (SELECT " & WrongCount & ", x.* FROM Class_Visitor x WHERE 1=1 " & SearchStr & ") b ON a.OCID=b.OCID "
            sql &= " CROSS JOIN Sys_VisitAlert c ) TempTable "
            If VisitResult.SelectedIndex <> 0 Then sql &= " WHERE VisitResult='" & VisitResult.SelectedValue & "' "
        Else
            'Y: >2007
            WrongCount = "  CASE WHEN Data1 IN (2,3) THEN 1 ELSE 0 END "
            WrongCount += " + CASE WHEN Data3 IN (2,3) THEN 1 ELSE 0 END "
            WrongCount += " + CASE WHEN Data5 IN (2,3) THEN 1 ELSE 0 END "
            WrongCount += " + CASE WHEN Data6 IN (2,3) THEN 1 ELSE 0 END "
            WrongCount += " + CASE WHEN Data10 IN (2,3) THEN 1 ELSE 0 END "
            WrongCount += " + CASE WHEN Item1_1='2' THEN 1 ELSE 0 END "
            WrongCount += " + CASE WHEN Item1_2='2' THEN 1 ELSE 0 END "
            WrongCount += " + CASE WHEN Item2_1='2' THEN 1 ELSE 0 END "
            WrongCount += " + CASE WHEN Item2_2='2' THEN 1 ELSE 0 END "
            WrongCount += " + CASE WHEN Item3_1='2' THEN 1 ELSE 0 END "
            WrongCount += " + CASE WHEN Item6_1='2' THEN 1 ELSE 0 END "
            WrongCount += " + CASE WHEN Item14='2' THEN 1 ELSE 0 END "
            WrongCount += " + CASE WHEN Item15='2' THEN 1 ELSE 0 END "
            WrongCount += " + CASE WHEN Item16='2' THEN 1 ELSE 0 END "
            WrongCount += " + CASE WHEN Item17='2' THEN 1 ELSE 0 END "
            WrongCount += " + CASE WHEN Item18='2' THEN 1 ELSE 0 END "
            WrongCount += " + CASE WHEN Item19='2' THEN 1 ELSE 0 END "
            WrongCount += " + CASE WHEN Item20='2' THEN 1 ELSE 0 END "
            WrongCount += " + CASE WHEN Item21='2' THEN 1 ELSE 0 END "
            WrongCount += " + CASE WHEN Item22='2' THEN 1 ELSE 0 END "
            WrongCount += " + CASE WHEN Item23='2' THEN 1 ELSE 0 END "
            WrongCount += " + CASE WHEN Item24='2' THEN 1 ELSE 0 END "
            WrongCount += " + CASE WHEN Item25='2' THEN 1 ELSE 0 END "
            WrongCount += " + CASE WHEN Item26='2' THEN 1 ELSE 0 END "
            WrongCount += " + CASE WHEN Item27='2' THEN 1 ELSE 0 END "
            If ddlYears.SelectedValue > "2007" And ddlYears.SelectedValue < "2010" Then
                WrongCount += " + CASE WHEN Data7 IN (2,3) THEN 1 ELSE 0 END "
                WrongCount += " + CASE WHEN Data9 IN (2,3) THEN 1 ELSE 0 END "
                WrongCount += " + CASE WHEN Data11 IN (2,3) THEN 1 ELSE 0 END "
                WrongCount += " + CASE WHEN Item28='2' THEN 1 ELSE 0 END "
                WrongCount += " + CASE WHEN Item28_2='2' THEN 1 ELSE 0 END"
                WrongCount += " + CASE WHEN Item29='2' THEN 1 ELSE 0 END "
                WrongCount += " + CASE WHEN Item30='2' THEN 1 ELSE 0 END "
            End If
            WrongCount += " AS WrongCount "

            If SDate.Text <> "" Then SearchStr += " AND ApplyDate >= " & TIMS.To_date(SDate.Text) & vbCrLf '" & SDate.Text & "'"
            If EDate.Text <> "" Then SearchStr += " AND ApplyDate <= " & TIMS.To_date(EDate.Text) & vbCrLf '" & EDate.Text & "'"
            If IsClear.SelectedIndex > 0 Then SearchStr += " AND IsClear='" & IsClear.SelectedValue & "' "

            sql = " SELECT * FROM ( "
            sql &= "   SELECT a.CyclType ,a.ClassCName ,a.FTDate ,b.OCID ,b.SeqNo ,b.ApplyDate ,b.WrongCount ,b.IsClear "
            sql &= "    ,CASE WHEN WrongCount/32.0*100 > c.RRate1 THEN 1 WHEN WrongCount/32.0*100 > c.YRate1 THEN 2 ELSE 3 END AS VisitResult "
            sql &= "   FROM (SELECT * FROM Class_ClassInfo WHERE RID='" & Me.ViewState("RID") & "' AND PlanID='" & Me.ViewState("PlanID") & "') a "
            sql &= "   JOIN (SELECT " & WrongCount & " ,x.* FROM Class_Visitor x WHERE 1=1 " & SearchStr & ") b ON a.OCID=b.OCID "
            sql &= " LEFT JOIN Sys_VisitAlert c ON 1=1) TempTable "
            If VisitResult.SelectedIndex <> 0 Then sql &= " WHERE VisitResult='" & VisitResult.SelectedValue & "' "
        End If

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)

        DataGridTable2.Visible = False
        msg2.Text = "查無資料"
        If dt.Rows.Count > 0 Then
            DataGridTable2.Visible = True
            msg2.Text = ""
            PageControler2.PageDataTable = dt
            'PageControler2.PrimaryKey=""
            'PageControler2.Sort=""
            PageControler2.ControlerLoad()
        End If

        'If TIMS.Get_SQLRecordCount(sql)=0 Then
        '    DataGridTable2.Visible=False
        '    msg2.Text="查無資料"
        'Else
        '    DataGridTable2.Visible=True
        '    PageControler2.SqlDataCreate(sql)
        'End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Me.ViewState("RID") = ""
        Me.ViewState("PlanID") = ""
        Page2.Visible = False
        Page1.Visible = True
        msg2.Text = ""
        msg3.Text = ""
        CheckMode.Text = ""
        DeCount.Text = ""
        RelCount.Text = ""
        ShowMode.SelectedIndex = 0
        SDate.Text = ""
        EDate.Text = Now.Date
        IsClear.SelectedIndex = -1
        VisitResult.SelectedIndex = -1
    End Sub

    Private Sub DataGrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize2, DataGrid2)

        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = "head_navy"
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim btn1 As Button = e.Item.FindControl("Button3")
                Dim btn2 As Button = e.Item.FindControl("Button4")
                If e.Item.ItemType = ListItemType.Item Then e.Item.CssClass = ""

                'If IsNumeric(drv("CyclType")) Then
                '    If Int(drv("CyclType")) <> 0 Then e.Item.Cells(1).Text += "第" & Int(drv("CyclType")) & "期"
                'End If

                Select Case TIMS.CINT1(drv("VisitResult"))
                    Case 1
                        e.Item.Cells(3).Text = "<font color=Red>●</font>"
                    Case 2
                        e.Item.Cells(3).Text = "<font color=Yellow>●</font>"
                    Case 3
                        e.Item.Cells(3).Text = "<font color=Green>●</font>"
                End Select
                If drv("IsClear") Then
                    e.Item.Cells(5).Text = "是"
                Else
                    e.Item.Cells(5).Text = "否"
                End If

                If drv("WrongCount") = 0 And Not drv("IsClear") Then
                    btn1.Visible = False
                    btn2.Visible = True
                Else
                    If drv("WrongCount") = 0 Then
                        btn1.Visible = False
                        btn2.Visible = False
                    Else
                        btn1.Visible = True
                        btn2.Visible = False
                    End If
                End If

                btn1.CommandArgument = "WHERE OCID='" & drv("OCID") & "' and SeqNo='" & drv("SeqNo") & "'"
                btn2.CommandArgument = "WHERE OCID='" & drv("OCID") & "' and SeqNo='" & drv("SeqNo") & "'"
        End Select
    End Sub

    Private Sub DataGrid2_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid2.ItemCommand
        Dim sql As String
        Dim dt As DataTable
        Dim da As SqlDataAdapter = Nothing
        Dim dr As DataRow

        Select Case e.CommandName
            Case "clear"
                sql = " SELECT * FROM Class_Visitor " & e.CommandArgument
                dt = DbAccess.GetDataTable(sql, da, objconn)
                If dt.Rows.Count <> 0 Then
                    dr = dt.Rows(0)
                    dr("IsClear") = 1
                    dr("ModifyAcct") = sm.UserInfo.UserID
                    dr("ModifyDate") = Now
                    DbAccess.UpdateDataTable(dt, da)
                End If
                Common.MessageBox(Me, "結案成功!")
                CreateVisit()
            Case "view"
                VisitResult1.Text = DataGrid2.Items(e.Item.ItemIndex).Cells(2).Text
                ClassCName1.Text = DataGrid2.Items(e.Item.ItemIndex).Cells(1).Text
                CleanItem()
                Dim WrongCount As String
                WrongCount = " CASE WHEN Data1 <> 1 THEN 1 ELSE 0 END + CASE WHEN Data2 <> 1 THEN 1 ELSE 0 END + CASE WHEN Data3 <>1 THEN 1 ELSE 0 END + CASE WHEN Data4 <>1 THEN 1 ELSE 0 END + "
                WrongCount += " CASE WHEN Data5 <> 1 THEN 1 ELSE 0 END + CASE WHEN Data6 <> 1 THEN 1 ELSE 0 END + CASE WHEN Data7 <>1 THEN 1 ELSE 0 END + "
                WrongCount += " CASE WHEN Item1_1='Y' THEN 0 ELSE 1 END + CASE WHEN Item1_2='Y' THEN 0 ELSE 1 END + CASE WHEN Item2_1='Y' THEN 0 ELSE 1 END + CASE WHEN Item2_2='Y' THEN 0 ELSE 1 END + "
                WrongCount += " CASE WHEN Item3_1='Y' THEN 0 ELSE 1 END + CASE WHEN Item3_2='Y' THEN 0 ELSE 1 END + CASE WHEN Item4_1='Y' THEN 0 ELSE 1 END + CASE WHEN Item5_1='Y' THEN 0 ELSE 1 END + "
                WrongCount += " CASE WHEN Item6_1='Y' THEN 0 ELSE 1 END AS WrongCount "

                Page2.Visible = False
                Page3.Visible = True

                sql = "SELECT * ,CASE WHEN WrongCount/16 > e.RRate1 THEN 1 WHEN WrongCount/16 > e.YRate1 THEN 2 ELSE 3 END AS VisitResult "
                sql &= " FROM (SELECT " & WrongCount & " ,x.* FROM Class_Visitor x " & e.CommandArgument & ") a "
                sql &= " LEFT JOIN Class_VisitorTrace b ON a.OCID=b.OCID AND a.SeqNo=b.SeqNo "
                sql &= " JOIN Class_ClassInfo c ON a.OCID=c.OCID "
                sql &= " JOIN Org_OrgInfo d ON c.ComIDNO=d.ComIDNO "
                sql &= " LEFT JOIN Sys_VisitAlert e ON 1=1 "
                dr = DbAccess.GetOneRow(sql, objconn)

                If Not dr Is Nothing Then
                    Me.ViewState("OCID") = dr("OCID")
                    Me.ViewState("SeqNo") = dr("SeqNo")
                    OrgName1.Text = dr("OrgName").ToString
                    ClassCName1.Text = dr("ClassCName").ToString
                    IsClear1.Text = IIf(dr("IsClear"), "是", "否")
                    ApplyDate1.Text = FormatDateTime(dr("ApplyDate"), 2)
                    If dr("TraceDate").ToString = "" Then
                        TraceDate.Text = Now.Date
                    Else
                        TraceDate.Text = FormatDateTime(dr("TraceDate"), 2)
                    End If

                    If ddlYears.SelectedValue <= 2007 Then '2007年(含)之前
                        Table7.Visible = True
                        Table7_97.Visible = False

                        For i As Integer = 1 To 7
                            Select Case dr("Data" & i).ToString
                                Case "1"
                                    CType(Me.FindControl("Data" & i), Label).Text = "備齊"
                                    CType(Me.FindControl("Data" & i & "Trace"), DropDownList).Enabled = False
                                    CType(Me.FindControl("Data" & i & "Trace"), DropDownList).SelectedIndex = 0
                                Case "2"
                                    CType(Me.FindControl("Data" & i), Label).Text = "未備齊"
                                    CType(Me.FindControl("Data" & i), Label).ForeColor = Color.Red
                                    CType(Me.FindControl("Data" & i & "Trace"), DropDownList).Enabled = True
                                    Common.SetListItem(CType(Me.FindControl("Data" & i & "Trace"), DropDownList), dr("Data" & i & "Trace").ToString)
                                Case "3"
                                    CType(Me.FindControl("Data" & i), Label).Text = "部份備齊"
                                    CType(Me.FindControl("Data" & i), Label).ForeColor = Color.Red
                                    CType(Me.FindControl("Data" & i & "Trace"), DropDownList).Enabled = True
                                    Common.SetListItem(CType(Me.FindControl("Data" & i & "Trace"), DropDownList), dr("Data" & i & "Trace").ToString)
                            End Select

                            CType(Me.FindControl("DataCopy" & i), Label).Text = dr("DataCopy" & i).ToString
                            CType(Me.FindControl("Data" & i & "TNote"), TextBox).Text = dr("Data" & i & "TNote").ToString
                        Next

                        Item1_1.Text = IIf(dr("Item1_1").ToString = "Y", "是", "<font color=Red>否</font>")
                        Item1_2.Text = IIf(dr("Item1_2").ToString = "Y", "是", "<font color=Red>否</font>")
                        Item2_1.Text = IIf(dr("Item2_1").ToString = "Y", "是", "<font color=Red>否</font>")
                        Item2_2.Text = IIf(dr("Item2_2").ToString = "Y", "是", "<font color=Red>否</font>")
                        Item3_1.Text = IIf(dr("Item3_1").ToString = "Y", "是", "<font color=Red>否</font>")
                        Item3_2.Text = IIf(dr("Item3_2").ToString = "Y", "是", "<font color=Red>否</font>")
                        Item4_1.Text = IIf(dr("Item4_1").ToString = "Y", "是", "<font color=Red>否</font>")
                        Item5_1.Text = IIf(dr("Item5_1").ToString = "Y", "是", "<font color=Red>否</font>")
                        Item6_1.Text = IIf(dr("Item6_1").ToString = "Y", "是", "<font color=Red>否</font>")

                        Item1Pros.Text = dr("Item1Pros").ToString
                        Item2Pros.Text = dr("Item2Pros").ToString
                        Item3Pros.Text = dr("Item3Pros").ToString
                        Item4Pros.Text = dr("Item4Pros").ToString
                        Item5Pros.Text = dr("Item5Pros").ToString

                        If dr("Item1_1").ToString = "Y" Then
                            Item1_1Trace.SelectedIndex = 0
                            Item1_1Trace.Enabled = False
                        Else
                            Common.SetListItem(Item1_1Trace, dr("Item1_1Trace").ToString)
                            Item1_1Trace.Enabled = True
                        End If
                        If dr("Item1_2").ToString = "Y" Then
                            Item1_2Trace.SelectedIndex = 0
                            Item1_2Trace.Enabled = False
                        Else
                            Common.SetListItem(Item1_2Trace, dr("Item1_2Trace").ToString)
                            Item1_2Trace.Enabled = True
                        End If
                        If dr("Item2_1").ToString = "Y" Then
                            Item2_1Trace.SelectedIndex = 0
                            Item2_1Trace.Enabled = False
                        Else
                            Common.SetListItem(Item2_1Trace, dr("Item2_1Trace").ToString)
                            Item2_1Trace.Enabled = True
                        End If
                        If dr("Item2_2").ToString = "Y" Then
                            Item2_2Trace.SelectedIndex = 0
                            Item2_2Trace.Enabled = False
                        Else
                            Common.SetListItem(Item2_2Trace, dr("Item2_2Trace").ToString)
                            Item2_2Trace.Enabled = True
                        End If
                        If dr("Item3_1").ToString = "Y" Then
                            Item3_1Trace.SelectedIndex = 0
                            Item3_1Trace.Enabled = False
                        Else
                            Common.SetListItem(Item3_1Trace, dr("Item3_1Trace").ToString)
                            Item3_1Trace.Enabled = True
                        End If
                        If dr("Item3_2").ToString = "Y" Then
                            Item3_2Trace.SelectedIndex = 0
                            Item3_2Trace.Enabled = False
                        Else
                            Common.SetListItem(Item3_2Trace, dr("Item3_2Trace").ToString)
                            Item3_2Trace.Enabled = True
                        End If
                        If dr("Item4_1").ToString = "Y" Then
                            Item4_1Trace.SelectedIndex = 0
                            Item4_1Trace.Enabled = False
                        Else
                            Common.SetListItem(Item4_1Trace, dr("Item4_1Trace").ToString)
                            Item4_1Trace.Enabled = True
                        End If
                        If dr("Item5_1").ToString = "Y" Then
                            Item5_1Trace.SelectedIndex = 0
                            Item5_1Trace.Enabled = False
                        Else
                            Common.SetListItem(Item5_1Trace, dr("Item5_1Trace").ToString)
                            Item5_1Trace.Enabled = True
                        End If
                        If dr("Item6_1").ToString = "Y" Then
                            Item6_1Trace.SelectedIndex = 0
                            Item6_1Trace.Enabled = False
                        Else
                            Common.SetListItem(Item6_1Trace, dr("Item6_1Trace").ToString)
                            Item6_1Trace.Enabled = True
                        End If

                        Item1_1TNote.Text = dr("Item1_1TNote").ToString
                        Item1_2TNote.Text = dr("Item1_2TNote").ToString
                        Item2_1TNote.Text = dr("Item2_1TNote").ToString
                        Item2_2TNote.Text = dr("Item2_2TNote").ToString
                        Item3_1TNote.Text = dr("Item3_1TNote").ToString
                        Item3_2TNote.Text = dr("Item3_2TNote").ToString
                        Item4_1TNote.Text = dr("Item4_1TNote").ToString
                        Item5_1TNote.Text = dr("Item5_1TNote").ToString
                        Item6_1TNote.Text = dr("Item6_1TNote").ToString
                        'Else
                        '    '沒資料的例外狀況
                        '    Button7.Visible=False
                        'End If
                    Else
                        Table7.Visible = False
                        Table7_97.Visible = True
                        If ddlYears.SelectedValue >= "2010" Then
                            Data7_TR.Style("display") = "none"
                            Data9_TR.Style("display") = "none"
                            Data11_TR.Style("display") = "none"
                            Item28_TR.Style("display") = "none"
                            Item28_2_TR.Style("display") = "none"
                            Item29_TR.Style("display") = "none"
                            Item30_TR.Style("display") = "none"
                        Else
                            Data7_TR.Style("display") = ""
                            Data9_TR.Style("display") = ""
                            Data11_TR.Style("display") = ""
                            Item28_TR.Style("display") = ""
                            Item28_2_TR.Style("display") = ""
                            Item29_TR.Style("display") = ""
                            Item30_TR.Style("display") = ""
                        End If
                        '(書)
                        Select Case dr("Data1").ToString      '第一題
                            Case "1"
                                Data1_97.Text = "備齊"
                                Data1Trace_97.Enabled = False
                                Data1Trace_97.SelectedIndex = 0
                            Case "2"
                                Data1_97.Text = "未備齊"
                                Data1_97.ForeColor = Color.Red
                                Data1Trace_97.Enabled = True
                                If dr("Data1Trace").ToString <> "" Then Data1Trace_97.SelectedValue = dr("Data1Trace").ToString
                            Case "3"
                                Data1_97.Text = "部份備齊"
                                Data1_97.ForeColor = Color.Red
                                Data1Trace_97.Enabled = True
                                If dr("Data1Trace").ToString <> "" Then Data1Trace_97.SelectedValue = dr("Data1Trace").ToString
                            Case "4"
                                Data1_97.Text = "免提供"
                                Data1Trace_97.Enabled = False
                                Data1Trace_97.SelectedIndex = 0
                        End Select

                        Select Case dr("Data3").ToString   '第二題
                            Case "1"
                                Data3_97.Text = "備齊"
                                Data3Trace_97.Enabled = False
                                Data3Trace_97.SelectedIndex = 0
                            Case "2"
                                Data3_97.Text = "未備齊"
                                Data3_97.ForeColor = Color.Red
                                Data3Trace_97.Enabled = True
                                If dr("Data2Trace").ToString <> "" Then Data3Trace_97.SelectedValue = dr("Data2Trace").ToStringIf
                            Case "3"
                                Data3_97.Text = "部份備齊"
                                Data3_97.ForeColor = Color.Red
                                Data3Trace_97.Enabled = True
                                If dr("Data2Trace").ToString <> "" Then Data3Trace_97.SelectedValue = dr("Data2Trace").ToString
                            Case "4"
                                Data3_97.Text = "免提供"
                                Data3Trace_97.Enabled = False
                                Data3Trace_97.SelectedIndex = 0
                        End Select

                        Select Case dr("Data5").ToString   '第三題
                            Case "1"
                                Data5_97.Text = "備齊"
                                Data5Trace_97.Enabled = False
                                Data5Trace_97.SelectedIndex = 0
                            Case "2"
                                Data5_97.Text = "未備齊"
                                Data5_97.ForeColor = Color.Red
                                Data5Trace_97.Enabled = True
                                If dr("Data3Trace").ToString <> "" Then Data5Trace_97.SelectedValue = dr("Data3Trace").ToString
                            Case "3"
                                Data5_97.Text = "部份備齊"
                                Data5_97.ForeColor = Color.Red
                                Data5Trace_97.Enabled = True
                                If dr("Data3Trace").ToString <> "" Then Data5Trace_97.SelectedValue = dr("Data3Trace").ToString
                            Case "4"
                                Data5_97.Text = "免提供"
                                Data5Trace_97.Enabled = False
                                Data5Trace_97.SelectedIndex = 0
                        End Select

                        Select Case dr("Data6").ToString   '第四題
                            Case "1"
                                Data6_97.Text = "備齊"
                                Data6Trace_97.Enabled = False
                                Data6Trace_97.SelectedIndex = 0
                            Case "2"
                                Data6_97.Text = "未備齊"
                                Data6_97.ForeColor = Color.Red
                                Data6Trace_97.Enabled = True
                                If dr("Data4Trace").ToString <> "" Then Data6Trace_97.SelectedValue = dr("Data4Trace").ToString
                            Case "3"
                                Data6_97.Text = "部份備齊"
                                Data6_97.ForeColor = Color.Red
                                Data6Trace_97.Enabled = True
                                If dr("Data4Trace").ToString <> "" Then Data6Trace_97.SelectedValue = dr("Data4Trace").ToString
                            Case "4"
                                Data6_97.Text = "免提供"
                                Data6Trace_97.Enabled = False
                                Data6Trace_97.SelectedIndex = 0
                        End Select

                        Select Case dr("Data7").ToString   '第五題
                            Case "1"
                                Data7_97.Text = "備齊"
                                Data7Trace_97.Enabled = False
                                Data7Trace_97.SelectedIndex = 0
                            Case "2"
                                Data7_97.Text = "未備齊"
                                Data7_97.ForeColor = Color.Red
                                Data7Trace_97.Enabled = True
                                If dr("Data5Trace").ToString <> "" Then Data7Trace_97.SelectedValue = dr("Data5Trace").ToString
                            Case "3"
                                Data7_97.Text = "部份備齊"
                                Data7_97.ForeColor = Color.Red
                                Data7Trace_97.Enabled = True
                                If dr("Data5Trace").ToString <> "" Then Data7Trace_97.SelectedValue = dr("Data5Trace").ToString
                            Case "4"
                                Data7_97.Text = "免提供"
                                Data7Trace_97.Enabled = False
                                Data7Trace_97.SelectedIndex = 0
                        End Select

                        Select Case dr("Data9").ToString   '第六題
                            Case "1"
                                Data9_97.Text = "備齊"
                                Data9Trace_97.Enabled = False
                                Data9Trace_97.SelectedIndex = 0
                            Case "2"
                                Data9_97.Text = "未備齊"
                                Data9_97.ForeColor = Color.Red
                                Data9Trace_97.Enabled = True
                                If dr("Data6Trace").ToString <> "" Then Data9Trace_97.SelectedValue = dr("Data6Trace").ToString
                            Case "3"
                                Data9_97.Text = "部份備齊"
                                Data9_97.ForeColor = Color.Red
                                Data9Trace_97.Enabled = True
                                If dr("Data6Trace").ToString <> "" Then Data9Trace_97.SelectedValue = dr("Data6Trace").ToString
                            Case "4"
                                Data9_97.Text = "免提供"
                                Data9Trace_97.Enabled = False
                                Data9Trace_97.SelectedIndex = 0
                        End Select

                        Select Case dr("Data10").ToString   '第七題
                            Case "1"
                                Data10_97.Text = "備齊"
                                Data10Trace_97.Enabled = False
                                Data10Trace_97.SelectedIndex = 0
                            Case "2"
                                Data10_97.Text = "未備齊"
                                Data10_97.ForeColor = Color.Red
                                Data10Trace_97.Enabled = True
                                If dr("Data7Trace").ToString <> "" Then Data10Trace_97.SelectedValue = dr("Data7Trace").ToString
                            Case "3"
                                Data10_97.Text = "部份備齊"
                                Data10_97.ForeColor = Color.Red
                                Data10Trace_97.Enabled = True
                                If dr("Data7Trace").ToString <> "" Then Data10Trace_97.SelectedValue = dr("Data7Trace").ToString
                            Case "4"
                                Data10_97.Text = "免提供"
                                Data10Trace_97.Enabled = False
                                Data10Trace_97.SelectedIndex = 0
                        End Select

                        Select Case dr("Data11").ToString   '第八題
                            Case "1"
                                Data11_97.Text = "備齊"
                                Data11Trace_97.Enabled = False
                                Data11Trace_97.SelectedIndex = 0
                            Case "2"
                                Data11_97.Text = "未備齊"
                                Data11_97.ForeColor = Color.Red
                                Data11Trace_97.Enabled = True
                                If dr("Data8Trace").ToString <> "" Then Data11Trace_97.SelectedValue = dr("Data8Trace").ToString
                            Case "3"
                                Data11_97.Text = "部份備齊"
                                Data11_97.ForeColor = Color.Red
                                Data11Trace_97.Enabled = True
                                If dr("Data8Trace").ToString <> "" Then Data11Trace_97.SelectedValue = dr("Data8Trace").ToString
                            Case "4"
                                Data11_97.Text = "免提供"
                                Data11Trace_97.Enabled = False
                                Data11Trace_97.SelectedIndex = 0
                        End Select

                        If dr("DataCopy1").ToString <> "" Then DataCopy1_97.Text = dr("DataCopy1") '第一題的影本
                        If dr("DataCopy3").ToString <> "" Then DataCopy3_97.Text = dr("DataCopy3") '第二題的影本
                        If dr("DataCopy5").ToString <> "" Then DataCopy5_97.Text = dr("DataCopy5") '第三題的影本
                        If dr("DataCopy6").ToString <> "" Then DataCopy6_97.Text = dr("DataCopy6") '第四題的影本
                        If dr("DataCopy7").ToString <> "" Then DataCopy7_97.Text = dr("DataCopy7") '第五題的影本
                        If dr("DataCopy9").ToString <> "" Then DataCopy9_97.Text = dr("DataCopy9") '第六題的影本
                        If dr("DataCopy10").ToString <> "" Then DataCopy10_97.Text = dr("DataCopy10") '第七題的影本
                        If dr("DataCopy11").ToString <> "" Then DataCopy11_97.Text = dr("DataCopy11") '第八題的影本
                        If dr("Data1TNote").ToString <> "" Then Data1TNote_97.Text = dr("Data1TNote") '第一題的備註
                        If dr("Data2TNote").ToString <> "" Then Data3TNote_97.Text = dr("Data2TNote") '第二題的備註
                        If dr("Data3TNote").ToString <> "" Then Data5TNote_97.Text = dr("Data3TNote") '第三題的備註
                        If dr("Data4TNote").ToString <> "" Then Data6TNote_97.Text = dr("Data4TNote") '第四題的備註
                        If dr("Data5TNote").ToString <> "" Then Data7TNote_97.Text = dr("Data5TNote") '第五題的備註
                        If dr("Data6TNote").ToString <> "" Then Data9TNote_97.Text = dr("Data6TNote") '第六題的備註
                        If dr("Data7TNote").ToString <> "" Then Data10TNote_97.Text = dr("Data7TNote") '第七題的備註
                        If dr("Data8TNote").ToString <> "" Then Data11TNote_97.Text = dr("Data8TNote") '第八題的備註

#Region "(No Use)"

                        'For i As Integer=5 To 7                   '第3~5題
                        '    Select Case dr("Data" & i).ToString
                        '        Case "1"
                        '            CType(Me.FindControl("Data" & i & "_97"), Label).Text="備齊"
                        '            CType(Me.FindControl("Data" & i & "Trace_97"), DropDownList).Enabled=False
                        '            CType(Me.FindControl("Data" & i & "Trace_97"), DropDownList).SelectedIndex=0
                        '        Case "2"
                        '            CType(Me.FindControl("Data" & i & "_97"), Label).Text="未備齊"
                        '            CType(Me.FindControl("Data" & i & "_97"), Label).ForeColor=Color.Red
                        '            CType(Me.FindControl("Data" & i & "Trace_97"), DropDownList).Enabled=True
                        '            Common.SetListItem(CType(Me.FindControl("Data" & i & "Trace_97"), DropDownList), dr("Data" & i & "Trace").ToString)
                        '        Case "3"
                        '            CType(Me.FindControl("Data" & i & "_97"), Label).Text="部份備齊"
                        '            CType(Me.FindControl("Data" & i & "_97"), Label).ForeColor=Color.Red
                        '            CType(Me.FindControl("Data" & i & "Trace_97"), DropDownList).Enabled=True
                        '            Common.SetListItem(CType(Me.FindControl("Data" & i & "Trace_97"), DropDownList), dr("Data" & i & "Trace").ToString)
                        '    End Select

                        '    CType(Me.FindControl("DataCopy" & i & "_97"), Label).Text=dr("DataCopy" & i).ToString
                        '    CType(Me.FindControl("Data" & i & "TNote_97"), TextBox).Text=dr("Data" & i & "Note").ToString
                        'Next

                        'For k As Integer=9 To 11
                        '    Select Case dr("Data" & k).ToString
                        '        Case "1"
                        '            CType(Me.FindControl("Data" & k & "_97"), Label).Text="備齊"
                        '            CType(Me.FindControl("Data" & k & "Trace_97"), DropDownList).Enabled=False
                        '            CType(Me.FindControl("Data" & k & "Trace_97"), DropDownList).SelectedIndex=0
                        '        Case "2"
                        '            CType(Me.FindControl("Data" & k & "_97"), Label).Text="未備齊"
                        '            CType(Me.FindControl("Data" & k & "_97"), Label).ForeColor=Color.Red
                        '            CType(Me.FindControl("Data" & k & "Trace_97"), DropDownList).Enabled=True
                        '            Common.SetListItem(CType(Me.FindControl("Data" & k & "Trace_97"), DropDownList), dr("Data" & k & "Trace").ToString)
                        '        Case "3"
                        '            CType(Me.FindControl("Data" & k & "_97"), Label).Text="部份備齊"
                        '            CType(Me.FindControl("Data" & k & "_97"), Label).ForeColor=Color.Red
                        '            CType(Me.FindControl("Data" & k & "Trace_97"), DropDownList).Enabled=True
                        '            Common.SetListItem(CType(Me.FindControl("Data" & k & "Trace_97"), DropDownList), dr("Data" & k & "Trace").ToString)
                        '    End Select

                        '    CType(Me.FindControl("DataCopy" & k & "_97"), Label).Text=dr("DataCopy" & k).ToString
                        '    CType(Me.FindControl("Data" & k & "TNote_97"), TextBox).Text=dr("Data" & k & "Note").ToString
                        'Next

#End Region

                        For j As Integer = 14 To 30
                            Select Case dr("Item" & j).ToString
                                Case "1"
                                    CType(FindControl("Item" & j & "_97"), Label).Text = "是"
                                    CType(FindControl("Item" & j & "Trace_97"), DropDownList).SelectedIndex = 0
                                    CType(FindControl("Item" & j & "Trace_97"), DropDownList).Enabled = False
                                Case "2"
                                    CType(FindControl("Item" & j & "_97"), Label).Text = "<font color=Red>否</font>"
                                    'CType(FindControl("Item" & j & "Trace_97"), DropDownList).SelectedIndex=1
                                    Common.SetListItem(CType(Me.FindControl("Item" & j & "Trace_97"), DropDownList), dr("Item" & j & "Trace").ToString)
                                    CType(FindControl("Item" & j & "Trace_97"), DropDownList).Enabled = True
                                Case "3"
                                    CType(FindControl("Item" & j & "_97"), Label).Text = "免填"
                                    Common.SetListItem(CType(Me.FindControl("Item" & j & "Trace_97"), DropDownList), dr("Item" & j & "Trace").ToString)
                                    'CType(FindControl("Item" & j & "Trace_97"), DropDownList).SelectedIndex=2
                                    CType(FindControl("Item" & j & "Trace_97"), DropDownList).Enabled = False
                            End Select
                            CType(FindControl("Item" & j & "TNote_97"), TextBox).Text = dr("Item" & j & "TNote").ToString
                        Next

                        Select Case dr("Item1_1").ToString
                            Case 1
                                Item1_1_97.Text = "是"
                                Item1_1Trace_97.SelectedIndex = 0
                                Item1_1Trace_97.Enabled = False
                            Case 2
                                Item1_1_97.Text = "<font color=Red>否</font>"
                                'Item1_1Trace_97.SelectedIndex=1
                                If dr("Item1_1Trace").ToString <> "" Then Item1_1Trace_97.SelectedValue = dr("Item1_1Trace").ToString
                                Item1_1Trace_97.Enabled = True
                            Case 3
                                Item1_1_97.Text = "免填"
                                'Item1_1Trace_97.SelectedIndex=2
                                If dr("Item1_1Trace").ToString <> "" Then Item1_1Trace_97.SelectedValue = dr("Item1_1Trace").ToString
                                Item1_1Trace_97.Enabled = False
                        End Select

                        Select Case dr("Item1_2").ToString
                            Case 1
                                Item1_2_97.Text = "是"
                                Item1_2Trace_97.SelectedIndex = 0
                                Item1_2Trace_97.Enabled = False
                            Case 2
                                Item1_2_97.Text = "<font color=Red>否</font>"
                                If dr("Item1_2Trace").ToString <> "" Then Item1_2Trace_97.SelectedValue = dr("Item1_2Trace").ToString
                                'Item1_2Trace_97.SelectedIndex=1
                                Item1_2Trace_97.Enabled = True
                            Case 3
                                Item1_2_97.Text = "免填"
                                If dr("Item1_2Trace").ToString <> "" Then Item1_2Trace_97.SelectedValue = dr("Item1_2Trace").ToString
                                'Item1_2Trace_97.SelectedIndex=2
                                Item1_2Trace_97.Enabled = False
                        End Select

                        Select Case dr("Item3_1").ToString
                            Case 1
                                Item3_1_97.Text = "是"
                                Item3_1Trace_97.SelectedIndex = 0
                                Item3_1Trace_97.Enabled = False
                            Case 2
                                Item3_1_97.Text = "<font color=Red>否</font>"
                                If dr("Item3_1Trace").ToString <> "" Then Item3_1Trace_97.SelectedValue = dr("Item3_1Trace").ToString
                                'Item3_1Trace_97.SelectedIndex=1
                                Item3_1Trace_97.Enabled = True
                            Case 3
                                Item3_1_97.Text = "免填"
                                If dr("Item3_1Trace").ToString <> "" Then Item3_1Trace_97.SelectedValue = dr("Item3_1Trace").ToString
                                'Item3_1Trace_97.SelectedIndex=2
                                Item3_1Trace_97.Enabled = False
                        End Select

                        Select Case dr("Item2_1").ToString
                            Case 1
                                Item2_1_97.Text = "是"
                                Item2_1Trace_97.SelectedIndex = 0
                                Item2_1Trace_97.Enabled = False
                            Case 2
                                Item2_1_97.Text = "<font color=Red>否</font>"
                                If dr("Item2_1Trace").ToString <> "" Then Item2_1Trace_97.SelectedValue = dr("Item2_1Trace").ToString
                                'Item2_1Trace_97.SelectedIndex=1
                                Item2_1Trace_97.Enabled = True
                            Case 3
                                Item2_1_97.Text = "免填"
                                If dr("Item2_1Trace").ToString <> "" Then Item2_1Trace_97.SelectedValue = dr("Item2_1Trace").ToString
                                'Item2_1Trace_97.SelectedIndex=2
                                Item2_1Trace_97.Enabled = False
                        End Select

                        Select Case dr("Item2_2").ToString
                            Case 1
                                Item2_2_97.Text = "是"
                                Item2_2Trace_97.SelectedIndex = 0
                                Item2_2Trace_97.Enabled = False
                            Case 2
                                Item2_2_97.Text = "<font color=Red>否</font>"
                                If dr("Item2_2Trace").ToString <> "" Then Item2_2Trace_97.SelectedValue = dr("Item2_2Trace").ToString
                                'Item2_2Trace_97.SelectedIndex=1
                                Item2_2Trace_97.Enabled = True
                            Case 3
                                Item2_2_97.Text = "免填"
                                If dr("Item2_2Trace").ToString <> "" Then Item2_2Trace_97.SelectedValue = dr("Item2_2Trace").ToString
                                'Item2_2Trace_97.SelectedIndex=2
                                Item2_2Trace_97.Enabled = False
                        End Select

                        Select Case dr("Item6_1").ToString
                            Case 1
                                Item6_1_97.Text = "是"
                                Item6_1Trace_97.SelectedIndex = 0
                                Item6_1Trace_97.Enabled = False
                            Case 2
                                Item6_1_97.Text = "<font color=Red>否</font>"
                                If dr("Item6_1Trace").ToString <> "" Then Item6_1Trace_97.SelectedValue = dr("Item6_1Trace").ToString
                                'Item6_1Trace_97.SelectedIndex=1
                                Item6_1Trace_97.Enabled = True
                            Case 3
                                Item6_1_97.Text = "免填"
                                If dr("Item6_1Trace").ToString <> "" Then Item6_1Trace_97.SelectedValue = dr("Item6_1Trace").ToString
                                'Item6_1Trace_97.SelectedIndex=2
                                Item6_1Trace_97.Enabled = False
                        End Select

                        Select Case dr("Item28_2").ToString
                            Case 1
                                Item28_2_97.Text = "是"
                                Item28_2Trace_97.SelectedIndex = 0
                                Item28_2Trace_97.Enabled = False
                            Case 2
                                Item28_2_97.Text = "<font color=Red>否</font>"
                                If dr("Item5_1Trace").ToString <> "" Then Item28_2Trace_97.SelectedValue = dr("Item5_1Trace").ToString
                                'Item28_2Trace_97.SelectedIndex=1
                                Item28_2Trace_97.Enabled = True
                            Case 3
                                Item28_2_97.Text = "免填"
                                If dr("Item5_1Trace").ToString <> "" Then Item28_2Trace_97.SelectedValue = dr("Item5_1Trace").ToString
                                'Item28_2Trace_97.SelectedIndex=2
                                Item28_2Trace_97.Enabled = False
                        End Select

                        '備註
                        Item1_1TNote_97.Text = dr("Item1_1TNote").ToString
                        Item1_2TNote_97.Text = dr("Item1_2TNote").ToString
                        Item2_1TNote_97.Text = dr("Item2_1TNote").ToString
                        Item2_2TNote_97.Text = dr("Item2_2TNote").ToString
                        Item3_1TNote_97.Text = dr("Item3_1TNote").ToString
                        Item6_1TNote_97.Text = dr("Item6_1TNote").ToString
                        Item28_2TNote_97.Text = dr("Item5_1TNote").ToString

                        '處理情形
                        Item1Pros_97.Text = dr("Item1Pros").ToString
                        Item2Pros_97.Text = dr("Item2Pros").ToString
                        Item3Pros_97.Text = dr("Item3Pros").ToString
                        Item4Pros_97.Text = dr("Item4Pros").ToString
                        Item5Pros_97.Text = dr("Item5Pros").ToString
                    End If
                Else
                    Button7.Visible = False '沒資料的例外狀況
                End If
        End Select
    End Sub

    Sub CleanItem()
        If ddlYears.SelectedValue <= 2007 Then
            Data1Trace.SelectedIndex = -1
            Data2Trace.SelectedIndex = -1
            Data3Trace.SelectedIndex = -1
            Data4Trace.SelectedIndex = -1
            Data5Trace.SelectedIndex = -1
            Data6Trace.SelectedIndex = -1
            Data7Trace.SelectedIndex = -1
            Item1_1Trace.SelectedIndex = -1
            Item1_2Trace.SelectedIndex = -1
            Item2_1Trace.SelectedIndex = -1
            Item2_2Trace.SelectedIndex = -1
            Item3_1Trace.SelectedIndex = -1
            Item3_2Trace.SelectedIndex = -1
            Item4_1Trace.SelectedIndex = -1
            Item5_1Trace.SelectedIndex = -1
            Item6_1Trace.SelectedIndex = -1

            Data1TNote.Text = ""
            Data2TNote.Text = ""
            Data3TNote.Text = ""
            Data4TNote.Text = ""
            Data5TNote.Text = ""
            Data6TNote.Text = ""
            Data7TNote.Text = ""
            Item1_1TNote.Text = ""
            Item1_2TNote.Text = ""
            Item2_1TNote.Text = ""
            Item2_2TNote.Text = ""
            Item3_1TNote.Text = ""
            Item3_2TNote.Text = ""
            Item4_1TNote.Text = ""
            Item5_1TNote.Text = ""
            Item6_1TNote.Text = ""
        Else
            Data1Trace_97.SelectedIndex = -1
            'Data2Trace.SelectedIndex=-1
            Data3Trace_97.SelectedIndex = -1
            'Data4Trace.SelectedIndex=-1
            Data5Trace_97.SelectedIndex = -1
            Data6Trace_97.SelectedIndex = -1
            Data7Trace_97.SelectedIndex = -1
            Data9Trace_97.SelectedIndex = -1
            Data10Trace_97.SelectedIndex = -1
            Data11Trace_97.SelectedIndex = -1

            Item1_1Trace_97.SelectedIndex = -1
            Item1_2Trace_97.SelectedIndex = -1
            Item2_1Trace_97.SelectedIndex = -1
            Item2_2Trace_97.SelectedIndex = -1
            Item3_1Trace_97.SelectedIndex = -1
            'Item3_2Trace.SelectedIndex=-1
            'Item4_1Trace.SelectedIndex=-1
            Item28_2Trace_97.SelectedIndex = -1
            Item6_1Trace_97.SelectedIndex = -1

            Item14Trace_97.SelectedIndex = -1
            Item15Trace_97.SelectedIndex = -1
            Item16Trace_97.SelectedIndex = -1
            Item17Trace_97.SelectedIndex = -1
            Item18Trace_97.SelectedIndex = -1
            Item19Trace_97.SelectedIndex = -1
            Item20Trace_97.SelectedIndex = -1
            Item21Trace_97.SelectedIndex = -1
            Item22Trace_97.SelectedIndex = -1
            Item23Trace_97.SelectedIndex = -1
            Item24Trace_97.SelectedIndex = -1
            Item25Trace_97.SelectedIndex = -1
            Item26Trace_97.SelectedIndex = -1
            Item27Trace_97.SelectedIndex = -1
            Item28Trace_97.SelectedIndex = -1
            Item29Trace_97.SelectedIndex = -1
            Item30Trace_97.SelectedIndex = -1

            Data1TNote_97.Text = ""
            Data3TNote_97.Text = ""
            Data5TNote_97.Text = ""
            Data6TNote_97.Text = ""
            Data7TNote_97.Text = ""
            Data9TNote_97.Text = ""
            Data10TNote_97.Text = ""
            Data11TNote_97.Text = ""

            Item1_1TNote_97.Text = ""
            Item1_2TNote_97.Text = ""
            Item2_1TNote_97.Text = ""
            Item2_2TNote_97.Text = ""
            Item3_1TNote_97.Text = ""
            'Item3_2TNote.Text=""
            'Item4_1TNote.Text=""
            Item28_2TNote_97.Text = ""
            Item6_1TNote_97.Text = ""
            Item14TNote_97.Text = ""
            Item15TNote_97.Text = ""
            Item16TNote_97.Text = ""
            Item17TNote_97.Text = ""
            Item18TNote_97.Text = ""
            Item19TNote_97.Text = ""
            Item20TNote_97.Text = ""
            Item21TNote_97.Text = ""
            Item22TNote_97.Text = ""
            Item23TNote_97.Text = ""
            Item24TNote_97.Text = ""
            Item25TNote_97.Text = ""
            Item26TNote_97.Text = ""
            Item27TNote_97.Text = ""
            Item28TNote_97.Text = ""
            Item29TNote_97.Text = ""
            Item30TNote_97.Text = ""
        End If
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim sql As String
        Dim dt As DataTable
        Dim da As SqlDataAdapter = Nothing
        Dim dr As DataRow
        Dim success As Integer = 0

        sql = " SELECT * FROM Class_VisitorTrace WHERE OCID='" & Me.ViewState("OCID") & "' AND SeqNo='" & Me.ViewState("SeqNo") & "' "
        dt = DbAccess.GetDataTable(sql, da, objconn)
        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
            dt.Rows.Add(dr)
            dr("OCID") = Me.ViewState("OCID")
            dr("SeqNo") = Me.ViewState("SeqNo")
        Else
            dr = dt.Rows(0)
        End If

        dr("TraceDate") = TraceDate.Text

        If ddlYears.SelectedValue <= 2007 Then
            If Data1Trace.Enabled = False Then
                dr("Data1Trace") = Convert.DBNull
                dr("Data1TNote") = Convert.DBNull
                success += 1
            Else
                dr("Data1Trace") = Data1Trace.SelectedValue
                dr("Data1TNote") = IIf(Data1TNote.Text = "", Convert.DBNull, Data1TNote.Text)
                If Data1Trace.SelectedIndex = 1 Then success += 1
            End If
            If Data2Trace.Enabled = False Then
                dr("Data2Trace") = Convert.DBNull
                dr("Data2TNote") = Convert.DBNull
                success += 1
            Else
                dr("Data2Trace") = Data2Trace.SelectedValue
                dr("Data2TNote") = IIf(Data2TNote.Text = "", Convert.DBNull, Data2TNote.Text)
                If Data2Trace.SelectedIndex = 1 Then success += 1
            End If
            If Data3Trace.Enabled = False Then
                dr("Data3Trace") = Convert.DBNull
                dr("Data3TNote") = Convert.DBNull
                success += 1
            Else
                dr("Data3Trace") = Data3Trace.SelectedValue
                dr("Data3TNote") = IIf(Data3TNote.Text = "", Convert.DBNull, Data3TNote.Text)
                If Data3Trace.SelectedIndex = 1 Then success += 1
            End If
            If Data4Trace.Enabled = False Then
                dr("Data4Trace") = Convert.DBNull
                dr("Data4TNote") = Convert.DBNull
                success += 1
            Else
                dr("Data4Trace") = Data4Trace.SelectedValue
                dr("Data4TNote") = IIf(Data4TNote.Text = "", Convert.DBNull, Data4TNote.Text)
                If Data4Trace.SelectedIndex = 1 Then success += 1
            End If
            If Data5Trace.Enabled = False Then
                dr("Data5Trace") = Convert.DBNull
                dr("Data5TNote") = Convert.DBNull
                success += 1
            Else
                dr("Data5Trace") = Data5Trace.SelectedValue
                dr("Data5TNote") = IIf(Data5TNote.Text = "", Convert.DBNull, Data5TNote.Text)
                If Data5Trace.SelectedIndex = 1 Then success += 1
            End If
            If Data6Trace.Enabled = False Then
                dr("Data6Trace") = Convert.DBNull
                dr("Data6TNote") = Convert.DBNull
                success += 1
            Else
                dr("Data6Trace") = Data6Trace.SelectedValue
                dr("Data6TNote") = IIf(Data6TNote.Text = "", Convert.DBNull, Data6TNote.Text)
                If Data6Trace.SelectedIndex = 1 Then success += 1
            End If
            If Data7Trace.Enabled = False Then
                dr("Data7Trace") = Convert.DBNull
                dr("Data7TNote") = Convert.DBNull
                success += 1
            Else
                dr("Data7Trace") = Data7Trace.SelectedValue
                dr("Data7TNote") = IIf(Data7TNote.Text = "", Convert.DBNull, Data7TNote.Text)
                If Data7Trace.SelectedIndex = 1 Then success += 1
            End If
            If Item1_1Trace.Enabled = False Then
                dr("Item1_1Trace") = Convert.DBNull
                dr("Item1_1TNote") = Convert.DBNull
                success += 1
            Else
                dr("Item1_1Trace") = Item1_1Trace.SelectedValue
                dr("Item1_1TNote") = IIf(Item1_1TNote.Text = "", Convert.DBNull, Item1_1TNote.Text)
                dr("Item1_2TNote") = IIf(Item1_2TNote.Text = "", Convert.DBNull, Item1_2TNote.Text)
                If Item1_1Trace.SelectedIndex = 1 Then success += 1
            End If
            If Item1_2Trace.Enabled = False Then
                dr("Item1_2Trace") = Convert.DBNull
                dr("Item1_2TNote") = Convert.DBNull
                success += 1
            Else
                dr("Item1_2Trace") = Item1_2Trace.SelectedValue
                dr("Item1_2TNote") = IIf(Item1_2TNote.Text = "", Convert.DBNull, Item1_2TNote.Text)
                If Item1_2Trace.SelectedIndex = 1 Then success += 1
            End If
            If Item2_1Trace.Enabled = False Then
                dr("Item2_1Trace") = Convert.DBNull
                dr("Item2_1TNote") = Convert.DBNull
                success += 1
            Else
                dr("Item2_1Trace") = Item2_1Trace.SelectedValue
                dr("Item2_1TNote") = IIf(Item2_1TNote.Text = "", Convert.DBNull, Item2_1TNote.Text)
                If Item2_1Trace.SelectedIndex = 1 Then success += 1
            End If
            If Item2_2Trace.Enabled = False Then
                dr("Item2_2Trace") = Convert.DBNull
                dr("Item2_2TNote") = Convert.DBNull
                success += 1
            Else
                dr("Item2_2Trace") = Item2_2Trace.SelectedValue
                dr("Item2_2TNote") = IIf(Item2_2TNote.Text = "", Convert.DBNull, Item2_2TNote.Text)
                If Item2_2Trace.SelectedIndex = 1 Then success += 1
            End If
            If Item3_1Trace.Enabled = False Then
                dr("Item3_1Trace") = Convert.DBNull
                dr("Item3_1TNote") = Convert.DBNull
                success += 1
            Else
                dr("Item3_1Trace") = Item3_1Trace.SelectedValue
                dr("Item3_1TNote") = IIf(Item3_1TNote.Text = "", Convert.DBNull, Item3_1TNote.Text)
                If Item3_1Trace.SelectedIndex = 1 Then success += 1
            End If
            If Item3_2Trace.Enabled = False Then
                dr("Item3_2Trace") = Convert.DBNull
                dr("Item3_2TNote") = Convert.DBNull
                success += 1
            Else
                dr("Item3_2Trace") = Item3_2Trace.SelectedValue
                dr("Item3_2TNote") = IIf(Item3_2TNote.Text = "", Convert.DBNull, Item3_2TNote.Text)
                If Item3_2Trace.SelectedIndex = 1 Then success += 1
            End If
            If Item4_1Trace.Enabled = False Then
                dr("Item4_1Trace") = Convert.DBNull
                dr("Item4_1TNote") = Convert.DBNull
                success += 1
            Else
                dr("Item4_1Trace") = Item4_1Trace.SelectedValue
                dr("Item4_1TNote") = IIf(Item4_1TNote.Text = "", Convert.DBNull, Item4_1TNote.Text)
                If Item4_1Trace.SelectedIndex = 1 Then success += 1
            End If
            If Item5_1Trace.Enabled = False Then
                dr("Item5_1Trace") = Convert.DBNull
                dr("Item5_1TNote") = Convert.DBNull
                success += 1
            Else
                dr("Item5_1Trace") = Item5_1Trace.SelectedValue
                dr("Item5_1TNote") = IIf(Item5_1TNote.Text = "", Convert.DBNull, Item5_1TNote.Text)
                If Item5_1Trace.SelectedIndex = 1 Then success += 1
            End If
            If Item6_1Trace.Enabled = False Then
                dr("Item6_1Trace") = Convert.DBNull
                dr("Item6_1TNote") = Convert.DBNull
                success += 1
            Else
                dr("Item6_1Trace") = Item6_1Trace.SelectedValue
                dr("Item6_1TNote") = IIf(Item6_1TNote.Text = "", Convert.DBNull, Item6_1TNote.Text)
                If Item6_1Trace.SelectedIndex = 1 Then success += 1
            End If
        Else
            If Data1Trace_97.Enabled = False Then '第一題
                dr("Data1Trace") = Convert.DBNull
                dr("Data1TNote") = Convert.DBNull
                success += 1
            Else
                dr("Data1Trace") = Data1Trace_97.SelectedValue
                dr("Data1TNote") = IIf(Data1TNote_97.Text = "", Convert.DBNull, Data1TNote_97.Text)
                If Data1Trace_97.SelectedIndex = 1 Then success += 1
            End If
            If Data3Trace_97.Enabled = False Then '第二題
                dr("Data2Trace") = Convert.DBNull
                dr("Data2TNote") = Convert.DBNull
                success += 1
            Else
                dr("Data2Trace") = Data2Trace.SelectedValue
                dr("Data2TNote") = IIf(Data3TNote_97.Text = "", Convert.DBNull, Data3TNote_97.Text)
                If Data3Trace_97.SelectedIndex = 1 Then success += 1
            End If
            If Data5Trace_97.Enabled = False Then  '第三題
                dr("Data3Trace") = Convert.DBNull
                dr("Data3TNote") = Convert.DBNull
                success += 1
            Else
                dr("Data3Trace") = Data3Trace.SelectedValue
                dr("Data3TNote") = IIf(Data5TNote_97.Text = "", Convert.DBNull, Data5TNote_97.Text)
                If Data5Trace_97.SelectedIndex = 1 Then success += 1
            End If
            If Data6Trace_97.Enabled = False Then     '第四題
                dr("Data4Trace") = Convert.DBNull
                dr("Data4TNote") = Convert.DBNull
                success += 1
            Else
                dr("Data4Trace") = Data6Trace_97.SelectedValue
                dr("Data4TNote") = IIf(Data6TNote_97.Text = "", Convert.DBNull, Data6TNote_97.Text)
                If Data6Trace_97.SelectedIndex = 1 Then success += 1
            End If
            If Data7Trace_97.Enabled = False Then       '第五題
                dr("Data5Trace") = Convert.DBNull
                dr("Data5TNote") = Convert.DBNull
                success += 1
            Else
                dr("Data5Trace") = Data5Trace.SelectedValue
                dr("Data5TNote") = IIf(Data7TNote_97.Text = "", Convert.DBNull, Data7TNote_97.Text)
                If Data7Trace_97.SelectedIndex = 1 Then success += 1
            End If
            If Data9Trace_97.Enabled = False Then      '第六題
                dr("Data6Trace") = Convert.DBNull
                dr("Data6TNote") = Convert.DBNull
                success += 1
            Else
                dr("Data6Trace") = Data6Trace.SelectedValue
                dr("Data6TNote") = IIf(Data9TNote_97.Text = "", Convert.DBNull, Data9TNote_97.Text)
                If Data9Trace_97.SelectedIndex = 1 Then success += 1
            End If
            If Data10Trace_97.Enabled = False Then   '第七題
                dr("Data7Trace") = Convert.DBNull
                dr("Data7TNote") = Convert.DBNull
                success += 1
            Else
                dr("Data7Trace") = Data7Trace.SelectedValue
                dr("Data7TNote") = IIf(Data10TNote_97.Text = "", Convert.DBNull, Data10TNote_97.Text)
                If Data10Trace_97.SelectedIndex = 1 Then success += 1
            End If
            If Data11Trace_97.Enabled = False Then   '第八題
                dr("Data8Trace") = Convert.DBNull
                dr("Data8TNote") = Convert.DBNull
                success += 1
            Else
                dr("Data8Trace") = Data7Trace.SelectedValue
                dr("Data8TNote") = IIf(Data11TNote_97.Text = "", Convert.DBNull, Data11TNote_97.Text)
                If Data11Trace_97.SelectedIndex = 1 Then success += 1
            End If
            If Item1_1Trace_97.Enabled = False Then             '第一題(查)
                dr("Item1_1Trace") = Convert.DBNull
                dr("Item1_1TNote") = Convert.DBNull
                success += 1
            Else
                dr("Item1_1Trace") = Item1_1Trace_97.SelectedValue
                dr("Item1_1TNote") = IIf(Item1_1TNote_97.Text = "", Convert.DBNull, Item1_1TNote_97.Text)
                'dr("Item1_2TNote")=IIf(Item1_2TNote.Text="", Convert.DBNull, Item1_2TNote.Text)
                If Item1_1Trace_97.SelectedIndex = 1 Then success += 1
            End If
            If Item1_2Trace_97.Enabled = False Then
                dr("Item1_2Trace") = Convert.DBNull
                dr("Item1_2TNote") = Convert.DBNull
                success += 1
            Else
                dr("Item1_2Trace") = Item1_2Trace_97.SelectedValue
                dr("Item1_2TNote") = IIf(Item1_2TNote_97.Text = "", Convert.DBNull, Item1_2TNote_97.Text)
                If Item1_2Trace_97.SelectedIndex = 1 Then success += 1
            End If
            If Item2_1Trace_97.Enabled = False Then
                dr("Item2_1Trace") = Convert.DBNull
                dr("Item2_1TNote") = Convert.DBNull
                success += 1
            Else
                dr("Item2_1Trace") = Item2_1Trace_97.SelectedValue
                dr("Item2_1TNote") = IIf(Item2_1TNote_97.Text = "", Convert.DBNull, Item2_1TNote_97.Text)
                If Item2_1Trace_97.SelectedIndex = 1 Then success += 1
            End If
            If Item2_2Trace_97.Enabled = False Then
                dr("Item2_2Trace") = Convert.DBNull
                dr("Item2_2TNote") = Convert.DBNull
                success += 1
            Else
                dr("Item2_2Trace") = Item2_2Trace_97.SelectedValue
                dr("Item2_2TNote") = IIf(Item2_2TNote_97.Text = "", Convert.DBNull, Item2_2TNote_97.Text)
                If Item2_2Trace_97.SelectedIndex = 1 Then success += 1
            End If
            If Item3_1Trace_97.Enabled = False Then
                dr("Item3_1Trace") = Convert.DBNull
                dr("Item3_1TNote") = Convert.DBNull
                success += 1
            Else
                dr("Item3_1Trace") = Item3_1Trace_97.SelectedValue
                dr("Item3_1TNote") = IIf(Item3_1TNote_97.Text = "", Convert.DBNull, Item3_1TNote_97.Text)
                If Item3_1Trace_97.SelectedIndex = 1 Then success += 1
            End If
#Region "(No Use)"

            'If Item3_2Trace.Enabled=False Then
            '    dr("Item3_2Trace")=Convert.DBNull
            '    dr("Item3_2TNote")=Convert.DBNull
            '    success += 1
            'Else
            '    dr("Item3_2Trace")=Item3_2Trace.SelectedValue
            '    dr("Item3_2TNote")=IIf(Item3_2TNote.Text="", Convert.DBNull, Item3_2TNote.Text)
            '    If Item3_2Trace.SelectedIndex=1 Then
            '        success += 1
            '    End If
            'End If
            'If Item4_1Trace.Enabled=False Then
            '    dr("Item4_1Trace")=Convert.DBNull
            '    dr("Item4_1TNote")=Convert.DBNull
            '    success += 1
            'Else
            '    dr("Item4_1Trace")=Item4_1Trace.SelectedValue
            '    dr("Item4_1TNote")=IIf(Item4_1TNote.Text="", Convert.DBNull, Item4_1TNote.Text)
            '    If Item4_1Trace.SelectedIndex=1 Then
            '        success += 1
            '    End If
            'End If

#End Region
            If Item28_2Trace_97.Enabled = False Then
                dr("Item5_1Trace") = Convert.DBNull
                dr("Item5_1TNote") = Convert.DBNull
                success += 1
            Else
                dr("Item5_1Trace") = Item28_2Trace_97.SelectedValue
                dr("Item5_1TNote") = IIf(Item28_2TNote_97.Text = "", Convert.DBNull, Item28_2TNote_97.Text)
                If Item28_2Trace_97.SelectedIndex = 1 Then
                    success += 1
                End If
            End If
            If Item6_1Trace_97.Enabled = False Then
                dr("Item6_1Trace") = Convert.DBNull
                dr("Item6_1TNote") = Convert.DBNull
                success += 1
            Else
                dr("Item6_1Trace") = Item6_1Trace_97.SelectedValue
                dr("Item6_1TNote") = IIf(Item6_1TNote_97.Text = "", Convert.DBNull, Item6_1TNote_97.Text)
                If Item6_1Trace_97.SelectedIndex = 1 Then success += 1
            End If
        End If

        For s As Integer = 14 To 30
            If CType(FindControl("Item" & s & "Trace_97"), DropDownList).Enabled = False Then
                dr("Item" & s & "Trace") = Convert.DBNull
                dr("Item" & s & "TNote") = Convert.DBNull
                success += 1
            Else
                dr("Item" & s & "Trace") = CType(FindControl("Item" & s & "Trace_97"), DropDownList).SelectedValue
                dr("Item" & s & "TNote") = IIf(CType(FindControl("Item" & s & "TNote_97"), TextBox).Text = "", Convert.DBNull, CType(FindControl("Item" & s & "TNote_97"), TextBox).Text)
                If CType(FindControl("Item" & s & "Trace_97"), DropDownList).SelectedIndex = 1 Then success += 1
            End If
        Next

        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now
        DbAccess.UpdateDataTable(dt, da)

        sql = " SELECT * FROM Class_Visitor WHERE OCID='" & Me.ViewState("OCID") & "' AND SeqNo='" & Me.ViewState("SeqNo") & "' "
        '2006/03/28 add conn by matt
        dt = DbAccess.GetDataTable(sql, da, objconn)
        If dt.Rows.Count <> 0 Then
            dr = dt.Rows(0)
            If ddlYears.SelectedValue <= 2007 Then
                If success = 16 Then
                    dr("IsClear") = 1
                Else
                    dr("IsClear") = 0
                End If
            Else
                If success = 32 Then
                    dr("IsClear") = 1
                Else
                    dr("IsClear") = 0
                End If
            End If
            dr("ModifyAcct") = sm.UserInfo.UserID
            dr("ModifyDate") = Now
            DbAccess.UpdateDataTable(dt, da)
        End If

        CreateVisit()
        Common.MessageBox(Me, "儲存成功!")
        Page2.Visible = True
        Page3.Visible = False
        Me.ViewState("OCID") = ""
        Me.ViewState("SeqNo") = ""
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Me.ViewState("OCID") = ""
        Me.ViewState("SeqNo") = ""
        Page2.Visible = True
        Page3.Visible = False
    End Sub

    Private Sub DataGrid3_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid3.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = "CP_TD1"
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                If e.Item.ItemType = ListItemType.Item Then e.Item.CssClass = "CP_TD2"
                'If IsNumeric(drv("CyclType")) Then
                '    If Int(drv("CyclType")) <> 0 Then e.Item.Cells(0).Text += "第" & Int(drv("CyclType")) & "期"
                'End If
        End Select
    End Sub

    Private Sub ddlYears_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlYears.SelectedIndexChanged
        Dim v_ddlYears As String = TIMS.GetListValue(ddlYears)
        If ddlYears.SelectedIndex <> 0 AndAlso v_ddlYears <> "" Then
            rblTPlanID = TIMS.Get_YearTPlan(rblTPlanID, v_ddlYears, "", objconn)
            rblTPlanID.Items.Insert(0, New ListItem("全部"))
            rblTPlanID.SelectedIndex = 0
        End If
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        DataGridTable1.Visible = False
        SearchTable.Visible = True
        Button9.Visible = False
    End Sub

End Class
