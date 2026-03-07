Partial Class CM_01_001_add
    Inherits AuthBasePage

    Dim CancelID As Integer
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload

        Dim sql As String
        'Dim dr As DataRow
        'CancelID
        '1:成本加工費
        '2:多期核銷
        '3:總平均核銷
        '4:學習單元
        '5:主要身分別

        sql = "SELECT CancelID FROM Plan_BudgetCan WHERE TPlanID='" & sm.UserInfo.TPlanID & "'"
        CancelID = DbAccess.ExecuteScalar(sql, objconn)
        CancelMode1.Visible = False
        CancelMode2.Visible = False
        CancelMode3.Visible = False
        CancelMode4.Visible = False
        CancelMode5.Visible = False
        If Not IsPostBack Then
            CancelDate1.Text = Now.Date
            CancelDate2.Text = Now.Date
            CancelDate3.Text = Now.Date
            CancelDate4.Text = Now.Date
            CancelDate5.Text = Now.Date
        End If

        CancelModeValue.Value = CancelID
        Select Case CancelID
            Case 1
                CancelMode1.Visible = True
            Case 2
                CancelMode2.Visible = True
            Case 3
                CancelMode3.Visible = True
            Case 4
                CancelMode4.Visible = True
            Case 5
                CancelMode5.Visible = True
        End Select


        If Not IsPostBack Then
            CreateBasic()
            If CreateItem() = True Then
                CreateCancelData()
            End If

            Me.ViewState("BudgetSearchStr") = Session("BudgetSearchStr")
            Session("BudgetSearchStr") = Nothing
        End If

        Button1.Attributes("onclick") = "return CheckData1();"
        Button5.Attributes("onclick") = "return CheckData2();"
        Button7.Attributes("onclick") = "return CheckData3();"
        Button15.Attributes("onclick") = "return CheckData4();"
        Button11.Attributes("onclick") = "return CheckData5();"
        Button10.Attributes("onclick") = "wopen('CM_01_001_StudList.aspx?ID=" & Request("ID") & "OCID=" & Request("OCID") & "&Mode=3','ClassList',450,500,1);"
        Button17.Attributes("onclick") = "wopen('CM_01_001_StudList.aspx?ID=" & Request("ID") & "OCID=" & Request("OCID") & "&Mode=4','ClassList',450,500,1);"
        Button13.Attributes("onclick") = "wopen('CM_01_001_StudList.aspx?ID=" & Request("ID") & "OCID=" & Request("OCID") & "&Mode=5','ClassList',450,500,1);"
        Button18.Style("display") = "none"
    End Sub

    Function CreateItem() As Boolean
        Dim rst As Boolean = True
        Dim sql As String
        Dim dt As DataTable
        'CancelID
        '1:成本加工費
        '2:多期核銷
        '3:總平均核銷
        '4:學習單元
        '5:主要身分別

        Select Case CancelID
            Case 1

                sql = "" & vbCrLf
                sql += " SELECT  " & vbCrLf
                sql += " 	a.PCID,b.CostName+CONVERT(varchar, a.OPrice)+'元*'+CONVERT(varchar, a.Itemage)+b.ItemageName+'*'+CONVERT(varchar, a.ItemCost)+b.ItemCostName as CostName " & vbCrLf
                sql += " FROM Plan_CostItem a" & vbCrLf
                sql += " JOIN Key_CostItem b ON a.CostID=b.CostID and a.PlanID='" & Request("PlanID") & "' and a.ComIDNO='" & Request("ComIDNO") & "' and a.SeqNo='" & Request("SeqNo") & "' " & vbCrLf
                dt = DbAccess.GetDataTable(sql, objconn)

                If dt.Rows.Count = 0 Then
                    PCID.Items.Clear()
                    PCID.Items.Add(New ListItem("查無可核銷的項目"))
                Else
                    With PCID
                        .DataSource = dt
                        .DataTextField = "CostName"
                        .DataValueField = "PCID"
                        .DataBind()
                        .Items.Insert(0, New ListItem("請選擇"))
                    End With
                End If

                '1.自辦 2.委辦 3.合辦 4.補助
                sql = "SELECT PlanType FROM Key_Plan WHERE TPlanID='" & sm.UserInfo.TPlanID & "'"
                Common.SetListItem(PlanType1, DbAccess.ExecuteScalar(sql, objconn))
            Case 2
                sql = "SELECT dbo.NVL(count(1),0) AS StudentCount FROM Class_StudentsOfClass WHERE OCID='" & Request("OCID") & "' and StudStatus IN (1,4,5)"
                PNum.Text = DbAccess.ExecuteScalar(sql, objconn)

                sql = "" & vbCrLf
                sql += " SELECT a.BudID,a.BudName " & vbCrLf
                sql += " FROM view_Budget a " & vbCrLf
                sql += " JOIN (SELECT * FROM Plan_Budget WHERE TPlanID='" & sm.UserInfo.TPlanID & "' and Syear=(SELECT Years FROM ID_Plan WHERE PlanID='" & sm.UserInfo.PlanID & "')) b ON a.BudID=b.BudID "
                dt = DbAccess.GetDataTable(sql, objconn)
                With BudID
                    .DataSource = dt
                    .DataTextField = "BudName"
                    .DataValueField = "BudID"
                    .DataBind()
                End With
                If BudID.Items.Count = 1 Then
                    BudID.SelectedIndex = 0
                End If

                sql = "SELECT PlanType FROM Key_Plan WHERE TPlanID='" & sm.UserInfo.TPlanID & "'"
                Common.SetListItem(PlanType2, DbAccess.ExecuteScalar(sql, objconn))
            Case 3
                sql = "SELECT * FROM Class_StudentsOfClass WHERE OCID='" & Request("OCID") & "' and StudStatus IN (1,4,5)"
                dt = DbAccess.GetDataTable(sql, objconn)
                Num1.Text = dt.Select("BudgetID='03' and PMode='1'").Length
                Num2.Text = dt.Select("BudgetID='03' and PMode='2'").Length
                Num3.Text = dt.Select("BudgetID='02' and PMode='1'").Length
                Num4.Text = dt.Select("BudgetID='02' and PMode='2'").Length

                sql = "SELECT PlanType FROM Key_Plan WHERE TPlanID='" & sm.UserInfo.TPlanID & "'"
                Common.SetListItem(PlanType3, DbAccess.ExecuteScalar(sql, objconn))
            Case 4
                sql = "SELECT * FROM Class_StudentsOfClass WHERE OCID='" & Request("OCID") & "' and StudStatus IN (1,4,5)"
                dt = DbAccess.GetDataTable(sql, objconn)
                CreateDGData(dt)
                sql = "SELECT PlanType FROM Key_Plan WHERE TPlanID='" & sm.UserInfo.TPlanID & "'"
                Common.SetListItem(PlanType4, DbAccess.ExecuteScalar(sql, objconn))
            Case 5
                sql = "SELECT * FROM Class_StudentsOfClass WHERE OCID='" & Request("OCID") & "' and StudStatus IN (1,4,5)"
                dt = DbAccess.GetDataTable(sql, objconn)
                GNum.Text = dt.Select("MIdentityID Not IN ('02','03','04','05','06','07','08','09','10','13','14','17','18')").Length
                SNum.Text = dt.Select("MIdentityID IN ('02','03','04','05','06','07','08','09','10','13','14','17','18')").Length

                Dim Cst_msg23 As String = "請先至【參數設定】功能設定" & vbCrLf & " 『訓用合一』計畫的【核銷%數】項目，" & vbCrLf & "機構別為 【" & Me.ViewState("OrgTypeName") & "】"
                Dim Cst_msg34 As String = "請先至【參數設定】功能設定" & vbCrLf & " 『推動事業單位辦理職前培訓計畫(原與企業合作辦理職前訓練)』計畫的【核銷%數】項目，" & vbCrLf & "機構別為 【" & Me.ViewState("OrgTypeName") & "】"
                Dim Cst_msg41 As String = "請先至【參數設定】功能設定" & vbCrLf & " 『推動營造業事業單位辦理職前培訓』計畫的【核銷%數】項目，" & vbCrLf & "機構別為 【" & Me.ViewState("OrgTypeName") & "】"
                Select Case Convert.ToString(sm.UserInfo.TPlanID)
                    Case "23", "34", "41"

                        Dim Errmsg1 As String = ""
                        Dim dr As DataRow

                        '設定核銷%數 核銷數
                        '23:訓用合一 
                        '34:與企業合作辦理職前訓練 
                        '41:推動營造業事業單位辦理職前培訓計畫
                        If sm.UserInfo.Years <= "2009" Then
                            sql = "SELECT * FROM Sys_GlobalVar WHERE DistID='" & sm.UserInfo.DistID & "' and TPlanID='" & sm.UserInfo.TPlanID & "' and GVID='9'"
                        ElseIf sm.UserInfo.Years >= "2010" Then
                            sql = "SELECT * FROM Sys_OrgType WHERE DistID='" & sm.UserInfo.DistID & "' and TPlanID='" & sm.UserInfo.TPlanID & "' and OrgTypeID = '" & Me.ViewState("orgkind") & "'"
                        End If

                        dr = DbAccess.GetOneRow(sql, objconn)
                        If dr Is Nothing Then
                            If sm.UserInfo.Years >= "2010" Then    '如果是>=2010年沒有設定參數才出現訊息
                                Select Case Convert.ToString(sm.UserInfo.TPlanID)
                                    Case "23"
                                        Errmsg1 = Cst_msg23
                                    Case "34"
                                        Errmsg1 = Cst_msg34
                                    Case "41"
                                        Errmsg1 = Cst_msg41
                                End Select
                                Common.MessageBox(Me, Errmsg1)
                                TIMS.Tooltip(Button11, Errmsg1)
                            End If
                            Button11.Enabled = False '新增鈕失效
                            rst = False
                            Return rst
                        Else
                            Button11.Enabled = True '新增鈕有效
                            Var1.Text = "*" & dr("ItemVar1") & "%"
                            Var2.Text = "*" & dr("ItemVar2") & "%"
                            ItemVar1.Value = dr("ItemVar1").ToString
                            ItemVar2.Value = dr("ItemVar2").ToString
                        End If
                End Select

                'If sm.UserInfo.TPlanID = "23" Then
                '    sql = "SELECT * FROM Sys_GlobalVar WHERE DistID='" & sm.UserInfo.DistID & "' and TPlanID='" & sm.UserInfo.TPlanID & "' and GVID='9'"
                '    Dim dr As DataRow
                '    dr = DbAccess.GetOneRow(sql)
                '    If dr Is Nothing Then
                '        Common.MessageBox(Me, "請先設定訓用合一核銷%數")
                '        Button11.Enabled = False
                '    Else
                '        Button11.Enabled = True
                '        Var1.Text = "*" & dr("ItemVar1") & "%"
                '        Var2.Text = "*" & dr("ItemVar2") & "%"
                '        ItemVar1.Value = dr("ItemVar1").ToString
                '        ItemVar2.Value = dr("ItemVar2").ToString
                '    End If
                'End If

                sql = "SELECT PlanType FROM Key_Plan WHERE TPlanID='" & sm.UserInfo.TPlanID & "'"
                Common.SetListItem(PlanType5, DbAccess.ExecuteScalar(sql, objconn))
                rst = True
        End Select
        Return rst
    End Function

    '建立班級的基本資料
    Sub CreateBasic()
        Dim sql As String
        Dim dr As DataRow
        'Dim dt As DataTable
        Dim TotalCost As Double
        Dim AdmPercent As Integer

        Select Case Request("CostMode")
            Case "1"
                sql = "SELECT dbo.NVL(Sum(OPrice*Itemage*dbo.NVL(ItemCost,1)),0) as TotalCost FROM Plan_CostItem WHERE PlanID='" & Request("PlanID") & "' and ComIDNO='" & Request("ComIDNO") & "' and SeqNo='" & Request("SeqNo") & "'"
                TotalCost = Math.Round(DbAccess.ExecuteScalar(sql, objconn))
                sql = "SELECT dbo.NVL(AdmPercent,0) as AdmPercent FROM Plan_PlanInfo WHERE PlanID='" & Request("PlanID") & "' and ComIDNO='" & Request("ComIDNO") & "' and SeqNo='" & Request("SeqNo") & "'"
                AdmPercent = DbAccess.ExecuteScalar(sql, objconn)
                sql = "SELECT dbo.NVL(Sum(OPrice*Itemage*dbo.NVL(ItemCost,1)),0) as TotalCost FROM Plan_CostItem WHERE PlanID='" & Request("PlanID") & "' and ComIDNO='" & Request("ComIDNO") & "' and SeqNo='" & Request("SeqNo") & "' and AdmFlag='Y'"
                TotalCost += Math.Round(DbAccess.ExecuteScalar(sql, objconn) * AdmPercent / 100)
            Case "2"
                sql = "SELECT dbo.NVL(Sum(OPrice*Itemage*dbo.NVL(ItemCost,1)),0) as TotalCost FROM Plan_CostItem WHERE PlanID='" & Request("PlanID") & "' and ComIDNO='" & Request("ComIDNO") & "' and SeqNo='" & Request("SeqNo") & "'"
                TotalCost = Math.Round(DbAccess.ExecuteScalar(sql, objconn))
            Case "3"
                sql = "SELECT dbo.NVL(Sum(OPrice*Itemage),0) as TotalCost FROM Plan_CostItem WHERE PlanID='" & Request("PlanID") & "' and ComIDNO='" & Request("ComIDNO") & "' and SeqNo='" & Request("SeqNo") & "'"
                TotalCost = Math.Round(DbAccess.ExecuteScalar(sql, objconn))
            Case "4"
                sql = "SELECT dbo.NVL(Sum(OPrice*Itemage),0) as TotalCost FROM Plan_CostItem WHERE PlanID='" & Request("PlanID") & "' and ComIDNO='" & Request("ComIDNO") & "' and SeqNo='" & Request("SeqNo") & "'"
                TotalCost = Math.Round(DbAccess.ExecuteScalar(sql, objconn))
                sql = "SELECT dbo.NVL(AdmPercent,0) as AdmPercent FROM Plan_PlanInfo WHERE PlanID='" & Request("PlanID") & "' and ComIDNO='" & Request("ComIDNO") & "' and SeqNo='" & Request("SeqNo") & "'"
                AdmPercent = DbAccess.ExecuteScalar(sql, objconn)
                sql = "SELECT dbo.NVL(Sum(OPrice*Itemage*dbo.NVL(ItemCost,1)),0) as TotalCost FROM Plan_CostItem WHERE PlanID='" & Request("PlanID") & "' and ComIDNO='" & Request("ComIDNO") & "' and SeqNo='" & Request("SeqNo") & "' and AdmFlag='Y'"
                TotalCost += Math.Round(DbAccess.ExecuteScalar(sql, objconn) * AdmPercent / 100)
        End Select
        TrainCost.Text = Format(TotalCost, "#,##0.00")


        sql = "" & vbCrLf
        sql += " SELECT " & vbCrLf
        sql += " 	c.OrgName,c.orgkind" & vbCrLf
        sql += " 	,d.Name as OrgTypeName" & vbCrLf
        sql += " 	,a.ClassCName,a.CyclType,a.Tnum,a.STDate,a.FTDate " & vbCrLf
        sql += " FROM Class_ClassInfo a" & vbCrLf
        sql += " 	JOIN Auth_Relship b ON a.RID=b.RID and a.OCID='" & Request("OCID") & "'" & vbCrLf
        sql += " 	JOIN Org_OrgInfo c ON b.OrgID=c.OrgID " & vbCrLf
        sql += " 	JOIN Key_OrgType d on c.orgkind = d.orgTypeID " & vbCrLf
        dr = DbAccess.GetOneRow(sql, objconn)

        OrgName.Text = dr("OrgName")
        Me.ViewState("orgkind") = dr("orgkind")
        Me.ViewState("OrgTypeName") = dr("OrgTypeName")
        ClassCName.Text = dr("ClassCName")
        If IsNumeric(dr("CyclType")) Then
            If Int(dr("CyclType")) <> 0 Then
                ClassCName.Text += "第" & dr("CyclType") & "期"
            End If
        End If

        TDate.Text = FormatDateTime(dr("STDate"), 2) & "~" & FormatDateTime(dr("FTDate"), 2)
        Tnum.Text = dr("Tnum").ToString
    End Sub

    Sub CreateCost()
        Dim sql As String
        Dim dr As DataRow
        sql = "select dbo.NVL(sum(CancelCost),0)  as CancelCost  from Budget_ClassCancel where OCID=" & Request("OCID") & ""
        dr = DbAccess.GetOneRow(sql, objconn)
        CancelCost.Text = Format(CDbl(Math.Round(Int(TrainCost.Text) - dr("CancelCost"), 2)), "#,##0.00")
        If CDbl(CancelCost.Text) >= 0 Then
            CancelCost.ForeColor = Color.Black
        Else
            CancelCost.ForeColor = Color.Red
        End If
    End Sub

    '建立核銷資料
    Sub CreateCancelData()
        Dim sql As String
        Dim dt As DataTable
        'CancelID
        '1:成本加工費
        '2:多期核銷
        '3:總平均核銷
        '4:學習單元
        '5:主要身分別

        Select Case CancelID
            Case 1
                sql = "" & vbCrLf
                sql += " SELECT " & vbCrLf
                sql += " 	a.*,c.CostName " & vbCrLf
                sql += " FROM " & vbCrLf
                sql += " 	Budget_ClassCancel a " & vbCrLf
                sql += " 	JOIN Plan_CostItem b ON a.PCID=b.PCID and a.OCID='" & Request("OCID") & "'" & vbCrLf
                sql += " 	JOIN Key_CostItem c ON b.CostID=c.CostID  " & vbCrLf
                sql += " Order By a.PCID,a.BCCID " & vbCrLf
                dt = DbAccess.GetDataTable(sql, objconn)
                DataGridTable1.Visible = False
                If dt.Rows.Count > 0 Then
                    DataGridTable1.Visible = True
                    DataGrid1.DataKeyField = "BCCID"
                    DataGrid1.DataSource = dt
                    DataGrid1.DataBind()
                End If
            Case 2

                sql = "" & vbCrLf
                sql += " SELECT a.*,b.BudName " & vbCrLf
                sql += " FROM " & vbCrLf
                sql += " 	Budget_ClassCancel a" & vbCrLf
                sql += " 	JOIN view_Budget b ON a.BudID=b.BudID and a.OCID='" & Request("OCID") & "'" & vbCrLf
                sql += " Order By a.Times" & vbCrLf
                dt = DbAccess.GetDataTable(sql, objconn)
                DataGridTable2.Visible = False
                If dt.Rows.Count > 0 Then
                    DataGridTable2.Visible = True
                    DataGrid2.DataKeyField = "BCCID"
                    DataGrid2.DataSource = dt
                    DataGrid2.DataBind()
                End If
            Case 3
                sql = "SELECT * FROM Budget_ClassCancel WHERE OCID='" & Request("OCID") & "' "
                sql += "Order By Times"
                dt = DbAccess.GetDataTable(sql, objconn)
                DataGridTable3.Visible = False
                If dt.Rows.Count > 0 Then
                    DataGridTable3.Visible = True
                    DataGrid3.DataKeyField = "BCCID"
                    DataGrid3.DataSource = dt
                    DataGrid3.DataBind()
                End If
            Case 4
                sql = "SELECT * FROM Budget_ClassCancel WHERE OCID='" & Request("OCID") & "' "
                sql += "Order By Times"
                dt = DbAccess.GetDataTable(sql, objconn)
                DataGridTable4.Visible = False
                If dt.Rows.Count > 0 Then
                    DataGridTable4.Visible = True
                    DataGrid4.DataKeyField = "BCCID"
                    DataGrid4.DataSource = dt
                    DataGrid4.DataBind()
                End If
            Case 5
                sql = "SELECT * FROM Budget_ClassCancel WHERE OCID='" & Request("OCID") & "' "
                sql += "Order By Times"
                dt = DbAccess.GetDataTable(sql, objconn)
                DataGridTable5.Visible = False
                If dt.Rows.Count > 0 Then
                    Dim sql2 As String
                    Dim dr2 As DataRow

                    sql2 = "SELECT * FROM Sys_GlobalVar WHERE DistID='" & sm.UserInfo.DistID & "' and TPlanID='" & sm.UserInfo.TPlanID & "' and GVID='9'"
                    dr2 = DbAccess.GetOneRow(sql2, objconn)
                    Me.ViewState("OldItemVar1") = dr2("ItemVar1")
                    Me.ViewState("OldItemVar2") = dr2("ItemVar2")
                    DataGridTable5.Visible = True
                    DataGrid5.DataKeyField = "BCCID"
                    DataGrid5.DataSource = dt
                    DataGrid5.DataBind()
                End If
        End Select

        CreateCost()
    End Sub

    '建立學習券資料表
    Sub CreateDGData(ByVal dt As DataTable)
        Dim MyRow As TableRow = Nothing
        Dim MyCell As TableCell = Nothing
        Dim Dtt As DataTable = Nothing

        Dim sql As String
        sql = "SELECT * FROM KEY_DGTHOUR ORDER BY DGID"
        Dtt = DbAccess.GetDataTable(sql, objconn)

        CreateRow(DGTable, MyRow, "CM_TD1")
        CreateCell(MyRow, MyCell, "就安/就保", "CM_TD1")
        MyCell.RowSpan = 2
        CreateCell(MyRow, MyCell, "單元", "CM_TD1")
        MyCell.ColumnSpan = 3
        CreateCell(MyRow, MyCell, "人數", "CM_TD1")
        MyCell.RowSpan = 2
        CreateCell(MyRow, MyCell, "單價", "CM_TD1")
        MyCell.RowSpan = 2
        CreateCell(MyRow, MyCell, "總時數", "CM_TD1")

        MyCell.RowSpan = 2
        CreateRow(DGTable, MyRow, "CM_TD1")
        CreateCell(MyRow, MyCell, Dtt.DefaultView(0)(1) & "(" & Dtt.DefaultView(0)(2) & ")", "CM_TD1")
        CreateCell(MyRow, MyCell, Dtt.DefaultView(1)(1) & "(" & Dtt.DefaultView(1)(2) & ")", "CM_TD1")
        CreateCell(MyRow, MyCell, Dtt.DefaultView(2)(1) & "(" & Dtt.DefaultView(2)(2) & ")", "CM_TD1")

        If dt.Select("BudgetID='02' and RelClass_Unit like '100'").Length <> 0 Then
            CreateRow(DGTable, MyRow)
            CreateCell(MyRow, MyCell, "就安")
            CreateCell(MyRow, MyCell, "ˇ")
            CreateCell(MyRow, MyCell, "")
            CreateCell(MyRow, MyCell, "")
            CreateCell(MyRow, MyCell, dt.Select("BudgetID='02' and RelClass_Unit like '100'").Length)
            CreateCell(MyRow, MyCell, 110)
            CreateCell(MyRow, MyCell, 12)
        End If
        If dt.Select("BudgetID='02' and RelClass_Unit like '010'").Length <> 0 Then
            CreateRow(DGTable, MyRow)
            CreateCell(MyRow, MyCell, "就安")
            CreateCell(MyRow, MyCell, "")
            CreateCell(MyRow, MyCell, "ˇ")
            CreateCell(MyRow, MyCell, "")
            CreateCell(MyRow, MyCell, dt.Select("BudgetID='02' and RelClass_Unit like '010'").Length)
            CreateCell(MyRow, MyCell, 110)
            CreateCell(MyRow, MyCell, 18)
        End If
        If dt.Select("BudgetID='02' and RelClass_Unit like '001'").Length <> 0 Then
            CreateRow(DGTable, MyRow)
            CreateCell(MyRow, MyCell, "就安")
            CreateCell(MyRow, MyCell, "")
            CreateCell(MyRow, MyCell, "")
            CreateCell(MyRow, MyCell, "ˇ")
            CreateCell(MyRow, MyCell, dt.Select("BudgetID='02' and RelClass_Unit like '001'").Length)
            CreateCell(MyRow, MyCell, 110)
            CreateCell(MyRow, MyCell, 6)
        End If
        If dt.Select("BudgetID='02' and RelClass_Unit like '110'").Length <> 0 Then
            CreateRow(DGTable, MyRow)
            CreateCell(MyRow, MyCell, "就安")
            CreateCell(MyRow, MyCell, "ˇ")
            CreateCell(MyRow, MyCell, "ˇ")
            CreateCell(MyRow, MyCell, "")
            CreateCell(MyRow, MyCell, dt.Select("BudgetID='02' and RelClass_Unit like '110'").Length)
            CreateCell(MyRow, MyCell, 110)
            CreateCell(MyRow, MyCell, 30)
        End If
        If dt.Select("BudgetID='02' and RelClass_Unit like '101'").Length <> 0 Then
            CreateRow(DGTable, MyRow)
            CreateCell(MyRow, MyCell, "就安")
            CreateCell(MyRow, MyCell, "ˇ")
            CreateCell(MyRow, MyCell, "", "")
            CreateCell(MyRow, MyCell, "ˇ")
            CreateCell(MyRow, MyCell, dt.Select("BudgetID='02' and RelClass_Unit like '101'").Length)
            CreateCell(MyRow, MyCell, 110)
            CreateCell(MyRow, MyCell, 18)
        End If
        If dt.Select("BudgetID='02' and RelClass_Unit like '011'").Length <> 0 Then
            CreateRow(DGTable, MyRow)
            CreateCell(MyRow, MyCell, "就安")
            CreateCell(MyRow, MyCell, "")
            CreateCell(MyRow, MyCell, "ˇ")
            CreateCell(MyRow, MyCell, "ˇ")
            CreateCell(MyRow, MyCell, dt.Select("BudgetID='02' and RelClass_Unit like '011'").Length)
            CreateCell(MyRow, MyCell, 110)
            CreateCell(MyRow, MyCell, 24)
        End If
        If dt.Select("BudgetID='02' and RelClass_Unit like '111'").Length <> 0 Then
            CreateRow(DGTable, MyRow)
            CreateCell(MyRow, MyCell, "就安")
            CreateCell(MyRow, MyCell, "ˇ")
            CreateCell(MyRow, MyCell, "ˇ")
            CreateCell(MyRow, MyCell, "ˇ")
            CreateCell(MyRow, MyCell, dt.Select("BudgetID='02' and RelClass_Unit like '111'").Length)
            CreateCell(MyRow, MyCell, 110)
            CreateCell(MyRow, MyCell, 36)
        End If

        If dt.Select("BudgetID='03' and RelClass_Unit like '100'").Length <> 0 Then
            CreateRow(DGTable, MyRow)
            CreateCell(MyRow, MyCell, "就保", "")
            CreateCell(MyRow, MyCell, "ˇ", "")
            CreateCell(MyRow, MyCell, "", "")
            CreateCell(MyRow, MyCell, "", "")
            CreateCell(MyRow, MyCell, dt.Select("BudgetID='03' and RelClass_Unit like '100'").Length, "")
            CreateCell(MyRow, MyCell, 110, "")
            CreateCell(MyRow, MyCell, 10, "")
        End If
        If dt.Select("BudgetID='03' and RelClass_Unit like '010'").Length <> 0 Then
            CreateRow(DGTable, MyRow)
            CreateCell(MyRow, MyCell, "就保", "")
            CreateCell(MyRow, MyCell, "", "")
            CreateCell(MyRow, MyCell, "ˇ", "")
            CreateCell(MyRow, MyCell, "", "")
            CreateCell(MyRow, MyCell, dt.Select("BudgetID='03' and RelClass_Unit like '010'").Length, "")
            CreateCell(MyRow, MyCell, 110, "")
            CreateCell(MyRow, MyCell, 12, "")
        End If
        If dt.Select("BudgetID='03' and RelClass_Unit like '001'").Length <> 0 Then
            CreateRow(DGTable, MyRow)
            CreateCell(MyRow, MyCell, "就保", "")
            CreateCell(MyRow, MyCell, "", "")
            CreateCell(MyRow, MyCell, "", "")
            CreateCell(MyRow, MyCell, "ˇ", "")
            CreateCell(MyRow, MyCell, dt.Select("BudgetID='03' and RelClass_Unit like '001'").Length, "")
            CreateCell(MyRow, MyCell, 110, "")
            CreateCell(MyRow, MyCell, 24, "")
        End If
        If dt.Select("BudgetID='03' and RelClass_Unit like '110'").Length <> 0 Then
            CreateRow(DGTable, MyRow)
            CreateCell(MyRow, MyCell, "就保", "")
            CreateCell(MyRow, MyCell, "ˇ", "")
            CreateCell(MyRow, MyCell, "ˇ", "")
            CreateCell(MyRow, MyCell, "", "")
            CreateCell(MyRow, MyCell, dt.Select("BudgetID='03' and RelClass_Unit like '110'").Length, "")
            CreateCell(MyRow, MyCell, 110, "")
            CreateCell(MyRow, MyCell, 22, "")
        End If
        If dt.Select("BudgetID='03' and RelClass_Unit like '101'").Length <> 0 Then
            CreateRow(DGTable, MyRow)
            CreateCell(MyRow, MyCell, "就保", "")
            CreateCell(MyRow, MyCell, "ˇ", "")
            CreateCell(MyRow, MyCell, "", "")
            CreateCell(MyRow, MyCell, "ˇ", "")
            CreateCell(MyRow, MyCell, dt.Select("BudgetID='03' and RelClass_Unit like '101'").Length, "")
            CreateCell(MyRow, MyCell, 110, "")
            CreateCell(MyRow, MyCell, 34, "")
        End If
        If dt.Select("BudgetID='03' and RelClass_Unit like '011'").Length <> 0 Then
            CreateRow(DGTable, MyRow)
            CreateCell(MyRow, MyCell, "就保", "")
            CreateCell(MyRow, MyCell, "", "")
            CreateCell(MyRow, MyCell, "ˇ", "")
            CreateCell(MyRow, MyCell, "ˇ", "")
            CreateCell(MyRow, MyCell, dt.Select("BudgetID='03' and RelClass_Unit like '011'").Length, "")
            CreateCell(MyRow, MyCell, 110, "")
            CreateCell(MyRow, MyCell, 36, "")
        End If
        If dt.Select("BudgetID='03' and RelClass_Unit like '111'").Length <> 0 Then
            CreateRow(DGTable, MyRow)
            CreateCell(MyRow, MyCell, "就保", "")
            CreateCell(MyRow, MyCell, "ˇ", "")
            CreateCell(MyRow, MyCell, "ˇ", "")
            CreateCell(MyRow, MyCell, "ˇ", "")
            CreateCell(MyRow, MyCell, dt.Select("BudgetID='03' and RelClass_Unit like '111'").Length, "")
            CreateCell(MyRow, MyCell, 110, "")
            CreateCell(MyRow, MyCell, 36, "")
        End If
    End Sub

    Sub CreateRow(ByRef MyTable As Table, ByRef MyRow As TableRow, Optional ByVal CssStyle As String = "")
        MyRow = New TableRow
        MyTable.Rows.Add(MyRow)
        If CssStyle <> "" Then
            MyRow.CssClass = CssStyle
        End If
    End Sub

    Sub CreateCell(ByRef MyRow As TableRow, ByRef MyCell As TableCell, ByVal CellText As String, Optional ByVal CssStyle As String = "CM_TD2")
        MyCell = New TableCell
        MyRow.Cells.Add(MyCell)

        MyCell.Text = CellText
        MyCell.CssClass = CssStyle
        MyCell.BorderWidth = Unit.Pixel(1)
    End Sub


    '方法一新增(成本加工法)
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim sql As String
        Dim dt As DataTable
        Dim dr As DataRow
        Dim OPrice As Double
        Dim BudgetCost As Double
        Dim AddUpCost As Double

        'sql = "SELECT OPrice,OPrice*Itemage*dbo.NVL(ItemCost,1) as BudgetCost FROM Plan_CostItem WHERE PCID='" & PCID.SelectedValue & "'"
        sql = "SELECT OPrice,OPrice*Itemage*dbo.NVL(ItemCost,1) as BudgetCost FROM Plan_CostItem WHERE PCID='" & PCID.SelectedValue & "'"
        dr = DbAccess.GetOneRow(sql, objconn)
        BudgetCost = dr("BudgetCost")
        OPrice = dr("OPrice")

        sql = "SELECT dbo.NVL(Sum(CancelCost),0) as CancelCost FROM Budget_ClassCancel WHERE PCID='" & PCID.SelectedValue & "'"
        AddUpCost = DbAccess.ExecuteScalar(sql, objconn)

        Dim da As SqlDataAdapter = Nothing
        Dim conn As SqlConnection = DbAccess.GetConnection()
        Call TIMS.OpenDbConn(conn)
        'TIMS.TestDbConn(Me, conn, True)
        sql = "SELECT * FROM Budget_ClassCancel WHERE 1<>1"
        '2006/03/28 add conn by matt
        dt = DbAccess.GetDataTable(sql, da, conn)

        dr = dt.NewRow
        dt.Rows.Add(dr)
        dr("OCID") = Request("OCID")
        dr("CancelDate") = CancelDate1.Text
        dr("CancelMode") = 1
        dr("BudgetCost") = BudgetCost
        dr("PCID") = PCID.SelectedValue
        dr("Itemage") = Itemage.Text
        If ItemCost.Text <> "" Then
            dr("ItemCost") = ItemCost.Text
            dr("CancelCost") = OPrice * Itemage.Text * ItemCost.Text
        Else
            dr("ItemCost") = 1
            dr("CancelCost") = OPrice * Itemage.Text * 1
        End If
        dr("AddUpCost") = AddUpCost
        dr("PlanType") = PlanType1.SelectedValue
        dr("Note") = IIf(Note1.Text = "", Convert.DBNull, Note1.Text)
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now

        DbAccess.UpdateDataTable(dt, da)
        'If conn.State = ConnectionState.Open Then conn.Close()
        Call TIMS.CloseDbConn(conn)
        CreateCancelData()
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = "CM_TD1"
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim btn As Button = e.Item.FindControl("Button2")

                If e.Item.ItemType = ListItemType.Item Then
                    e.Item.CssClass = "CM_TD2"
                End If

                e.Item.Cells(0).Text = e.Item.ItemIndex + 1
                e.Item.Cells(8).Text = Format(drv("BudgetCost") - drv("AddUpCost") - drv("CancelCost"), "#,##0.00")
                If CDbl(e.Item.Cells(8).Text) < 0 Then
                    e.Item.Cells(8).ForeColor = Color.Red
                End If
                btn.Attributes("onclick") = "return confirm('您確定要刪除這一筆資料?');"
                btn.CommandArgument = drv("BCCID")
        End Select
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Dim sql As String
        Dim dt As DataTable
        Dim dr As DataRow
        Dim da As SqlDataAdapter = Nothing
        Dim PCID As Integer

        sql = "SELECT PCID FROM Budget_ClassCancel WHERE BCCID='" & e.CommandArgument & "'"
        PCID = DbAccess.ExecuteScalar(sql, objconn)

        sql = "DELETE Budget_ClassCancel WHERE BCCID='" & e.CommandArgument & "'"
        DbAccess.ExecuteNonQuery(sql, objconn)

        Dim conn As SqlConnection = DbAccess.GetConnection()
        Call TIMS.OpenDbConn(conn)

        'TIMS.TestDbConn(Me, conn, True)
        sql = "SELECT * FROM Budget_ClassCancel WHERE PCID='" & PCID & "' Order BY BCCID DESC"
        '2006/03/28 add conn by matt
        dt = DbAccess.GetDataTable(sql, da, conn)
        For i As Integer = 0 To dt.Rows.Count - 1
            dr = dt.Rows(i)
            If i = dt.Rows.Count - 1 Then
                dr("AddUpCost") = 0
            Else
                dr("AddUpCost") = 0
                For j As Integer = i + 1 To dt.Rows.Count - 1
                    dr("AddUpCost") += dt.Rows(j)("CancelCost")
                Next
            End If
        Next
        DbAccess.UpdateDataTable(dt, da)
        Call TIMS.CloseDbConn(conn)
        'If conn.State = ConnectionState.Open Then conn.Close()

        CreateCancelData()
        Common.MessageBox(Me, "刪除成功")
    End Sub

    '回上一頁
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Session("BudgetSearchStr") = Me.ViewState("BudgetSearchStr")
        TIMS.Utl_Redirect1(Me, "CM_01_001.aspx?ID=" & Request("ID") & "")
    End Sub

    Private Sub DataGrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = "CM_TD1"
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim btn As Button = e.Item.FindControl("Button6")

                If e.Item.ItemType = ListItemType.Item Then
                    e.Item.CssClass = "CM_TD2"
                End If

                btn.Attributes("onclick") = "return confirm('您確定要刪除這一筆資料?');"
                btn.CommandArgument = drv("BCCID")
        End Select
    End Sub

    Private Sub DataGrid2_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid2.ItemCommand
        Dim sql As String = ""

        sql = "DELETE Budget_ClassCancel WHERE BCCID='" & e.CommandArgument & "'"
        DbAccess.ExecuteNonQuery(sql, objconn)

        CreateCancelData()
        Common.MessageBox(Me, "刪除成功")
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim sql As String
        Dim dt As DataTable
        Dim dr As DataRow
        Dim da As SqlDataAdapter = Nothing
        Dim conn As SqlConnection = DbAccess.GetConnection()
        Call TIMS.OpenDbConn(conn)
        'TIMS.TestDbConn(Me, conn, True)

        sql = "SELECT * FROM Budget_ClassCancel WHERE 1<>1"
        '2006/03/28 add conn by matt
        dt = DbAccess.GetDataTable(sql, da, conn)

        dr = dt.NewRow
        dt.Rows.Add(dr)
        dr("OCID") = Request("OCID")
        dr("CancelDate") = CancelDate2.Text
        dr("CancelMode") = 2
        dr("BudID") = BudID.SelectedValue
        dr("Times") = Times1.Text
        dr("PNum") = PNum.Text
        dr("PMoney") = PMoney.Text
        dr("CancelCost") = Int(PNum.Text) * CDbl(PMoney.Text)
        dr("PlanType") = PlanType2.SelectedValue
        dr("Note") = IIf(Note2.Text = "", Convert.DBNull, Note2.Text)
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now

        DbAccess.UpdateDataTable(dt, da)
        Call TIMS.CloseDbConn(conn)
        'If conn.State = ConnectionState.Open Then conn.Close()
        CreateCancelData()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Session("BudgetSearchStr") = Me.ViewState("BudgetSearchStr")
        TIMS.Utl_Redirect1(Me, "CM_01_001.aspx?ID=" & Request("ID") & "")
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim sql As String = ""
        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing
        Dim da As SqlDataAdapter = Nothing

        Dim conn As SqlConnection = DbAccess.GetConnection()
        Call TIMS.OpenDbConn(conn)
        'TIMS.TestDbConn(Me, conn, True)
        sql = "SELECT * FROM Budget_ClassCancel WHERE 1<>1"
        '2006/03/28 add conn by matt
        dt = DbAccess.GetDataTable(sql, da, conn)

        dr = dt.NewRow
        dt.Rows.Add(dr)
        dr("OCID") = Request("OCID")
        dr("CancelDate") = CancelDate3.Text
        dr("CancelMode") = 3
        dr("Times") = Times2.Text
        dr("SelfPrice") = SelfPrice.Text
        dr("Num1") = IIf(Num1.Text = "", Convert.DBNull, Num1.Text)
        dr("Percent1") = IIf(Percent1.Text = "", Convert.DBNull, Percent1.Text)
        dr("Num2") = IIf(Num2.Text = "", Convert.DBNull, Num2.Text)
        dr("Percent2") = IIf(Percent2.Text = "", Convert.DBNull, Percent2.Text)
        dr("Num3") = IIf(Num3.Text = "", Convert.DBNull, Num3.Text)
        dr("Percent3") = IIf(Percent3.Text = "", Convert.DBNull, Percent3.Text)
        dr("Num4") = IIf(Num4.Text = "", Convert.DBNull, Num4.Text)
        dr("Percent4") = IIf(Percent4.Text = "", Convert.DBNull, Percent4.Text)
        dr("CancelCost") = CDbl(IIf(SelfPrice.Text = "", 0, SelfPrice.Text)) * Int(IIf(Num1.Text = "", 0, Num1.Text)) * Int(IIf(Percent1.Text = "", 0, Percent1.Text)) / 100 + CDbl(IIf(SelfPrice.Text = "", 0, SelfPrice.Text)) * Int(IIf(Num2.Text = "", 0, Num2.Text)) * Int(IIf(Percent2.Text = "", 0, Percent2.Text)) / 100 + CDbl(IIf(SelfPrice.Text = "", 0, SelfPrice.Text)) * Int(IIf(Num3.Text = "", 0, Num3.Text)) * Int(IIf(Percent3.Text = "", 0, Percent3.Text)) / 100 + CDbl(IIf(SelfPrice.Text = "", 0, SelfPrice.Text)) * Int(IIf(Num4.Text = "", 0, Num4.Text)) * Int(IIf(Percent4.Text = "", 0, Percent4.Text)) / 100
        dr("PlanType") = PlanType3.SelectedValue
        dr("Note") = IIf(Note3.Text = "", Convert.DBNull, Note3.Text)
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now

        DbAccess.UpdateDataTable(dt, da)
        Call TIMS.CloseDbConn(conn)
        'If conn.State = ConnectionState.Open Then conn.Close()
        CreateCancelData()
    End Sub

    Private Sub DataGrid3_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid3.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = "CM_TD1"
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim btn As Button = e.Item.FindControl("Button9")
                Dim PValue1 As HtmlInputHidden = e.Item.FindControl("PValue1")
                Dim PValue2 As HtmlInputHidden = e.Item.FindControl("PValue2")
                Dim PValue3 As HtmlInputHidden = e.Item.FindControl("PValue3")
                Dim PValue4 As HtmlInputHidden = e.Item.FindControl("PValue4")

                If e.Item.ItemType = ListItemType.Item Then
                    e.Item.CssClass = "CM_TD2"
                End If

                If Not IsDBNull(drv("Num1")) And Not IsDBNull(drv("Percent1")) Then
                    e.Item.Cells(5).Text = drv("Num1") & "人*" & drv("Percent1") & "%=" & Math.Round(CDbl(drv("SelfPrice")) * Int(drv("Num1")) * Int(drv("Percent1")) / 100, 2)
                End If
                If Not IsDBNull(drv("Num2")) And Not IsDBNull(drv("Percent2")) Then
                    e.Item.Cells(6).Text = drv("Num2") & "人*" & drv("Percent2") & "%=" & Math.Round(CDbl(drv("SelfPrice")) * Int(drv("Num2")) * Int(drv("Percent2")) / 100, 2)
                End If
                If Not IsDBNull(drv("Num3")) And Not IsDBNull(drv("Percent3")) Then
                    e.Item.Cells(3).Text = drv("Num3") & "人*" & drv("Percent3") & "%=" & Math.Round(CDbl(drv("SelfPrice")) * Int(drv("Num3")) * Int(drv("Percent3")) / 100, 2)
                End If
                If Not IsDBNull(drv("Num4")) And Not IsDBNull(drv("Percent4")) Then
                    e.Item.Cells(4).Text = drv("Num4") & "人*" & drv("Percent4") & "%=" & Math.Round(CDbl(drv("SelfPrice")) * Int(drv("Num4")) * Int(drv("Percent4")) / 100, 2)
                End If

                If Not IsDBNull(drv("Percent1")) Then
                    PValue1.Value = drv("Percent1")
                Else
                    PValue1.Value = 0
                End If
                If Not IsDBNull(drv("Percent2")) Then
                    PValue2.Value = drv("Percent2")
                Else
                    PValue2.Value = 0
                End If
                If Not IsDBNull(drv("Percent3")) Then
                    PValue3.Value = drv("Percent3")
                Else
                    PValue3.Value = 0
                End If
                If Not IsDBNull(drv("Percent4")) Then
                    PValue4.Value = drv("Percent4")
                Else
                    PValue4.Value = 0
                End If

                btn.Attributes("onclick") = "return confirm('您確定要刪除這一筆資料?');"
                btn.CommandArgument = drv("BCCID")
        End Select
    End Sub

    Private Sub DataGrid3_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid3.ItemCommand
        Dim sql As String
        sql = "DELETE Budget_ClassCancel WHERE BCCID='" & e.CommandArgument & "'"
        DbAccess.ExecuteNonQuery(sql, objconn)

        CreateCancelData()
        Common.MessageBox(Me, "刪除成功")
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Session("BudgetSearchStr") = Me.ViewState("BudgetSearchStr")
        TIMS.Utl_Redirect1(Me, "CM_01_001.aspx?ID=" & Request("ID") & "")
    End Sub

    '新增-方法五
    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Dim sql As String
        Dim dt As DataTable
        Dim dr As DataRow
        Dim da As SqlDataAdapter = Nothing

        Dim conn As SqlConnection = DbAccess.GetConnection()
        Call TIMS.OpenDbConn(conn)
        'TIMS.TestDbConn(Me, conn, True)
        sql = "SELECT * FROM Budget_ClassCancel WHERE 1<>1"
        '2006/03/28 add conn by matt
        dt = DbAccess.GetDataTable(sql, da, conn)

        dr = dt.NewRow
        dt.Rows.Add(dr)
        dr("OCID") = Request("OCID")
        dr("CancelDate") = CancelDate5.Text
        dr("CancelMode") = 5
        dr("Times") = Times4.Text
        dr("GNum") = IIf(GNum.Text = "", Convert.DBNull, GNum.Text)
        dr("GPrice") = IIf(GPrice.Text = "", Convert.DBNull, GPrice.Text)
        dr("SNum") = IIf(SNum.Text = "", Convert.DBNull, SNum.Text)
        dr("SPrice") = IIf(SPrice.Text = "", Convert.DBNull, SPrice.Text)

        Select Case Convert.ToString(sm.UserInfo.TPlanID)
            Case "23", "34", "41"
                '設定核銷%數 核銷數
                '23:訓用合一 
                '34:與企業合作辦理職前訓練 
                '41:推動營造業事業單位辦理職前培訓計畫
                dr("CancelCost") = CDbl(IIf(GNum.Text = "", 0, GNum.Text)) * CDbl(IIf(GPrice.Text = "", 0, GPrice.Text)) * CDbl(IIf(ItemVar1.Value = "", 0, ItemVar1.Value)) / 100 + CDbl(IIf(SNum.Text = "", 0, SNum.Text)) * CDbl(IIf(SPrice.Text = "", 0, SPrice.Text)) * CDbl(IIf(ItemVar2.Value = "", 0, ItemVar2.Value)) / 100
            Case Else
                dr("CancelCost") = CDbl(IIf(GNum.Text = "", 0, GNum.Text)) * CDbl(IIf(GPrice.Text = "", 0, GPrice.Text)) + CDbl(IIf(SNum.Text = "", 0, SNum.Text)) * CDbl(IIf(SPrice.Text = "", 0, SPrice.Text))
        End Select

        'If sm.UserInfo.TPlanID = "23" Then
        '    dr("CancelCost") = CDbl(IIf(GNum.Text = "", 0, GNum.Text)) * CDbl(IIf(GPrice.Text = "", 0, GPrice.Text)) * CDbl(IIf(ItemVar1.Value = "", 0, ItemVar1.Value)) / 100 + CDbl(IIf(SNum.Text = "", 0, SNum.Text)) * CDbl(IIf(SPrice.Text = "", 0, SPrice.Text)) * CDbl(IIf(ItemVar2.Value = "", 0, ItemVar2.Value)) / 100
        'Else
        '    dr("CancelCost") = CDbl(IIf(GNum.Text = "", 0, GNum.Text)) * CDbl(IIf(GPrice.Text = "", 0, GPrice.Text)) + CDbl(IIf(SNum.Text = "", 0, SNum.Text)) * CDbl(IIf(SPrice.Text = "", 0, SPrice.Text))
        'End If

        dr("PlanType") = PlanType5.SelectedValue
        dr("Note") = IIf(Note5.Text = "", Convert.DBNull, Note5.Text)
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now

        DbAccess.UpdateDataTable(dt, da)
        Call TIMS.CloseDbConn(conn)
        'If conn.State = ConnectionState.Open Then conn.Close()
        CreateCancelData()
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Session("BudgetSearchStr") = Me.ViewState("BudgetSearchStr")
        TIMS.Utl_Redirect1(Me, "CM_01_001.aspx?ID=" & Request("ID") & "")
    End Sub

    Private Sub DataGrid5_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid5.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = "CM_TD1"
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim btn As Button = e.Item.FindControl("Button14")

                If e.Item.ItemType = ListItemType.Item Then
                    e.Item.CssClass = "CM_TD2"
                End If

                If drv("CancelDate") < "2010/08/10" Then  '如果是舊資料

                    If Not IsDBNull(drv("GNum")) And Not IsDBNull(drv("GPrice")) Then
                        Select Case Convert.ToString(sm.UserInfo.TPlanID)
                            Case "23", "34", "41" '設定核銷%數 核銷數
                                e.Item.Cells(2).Text = drv("GNum") & "人*" & drv("GPrice") & "*" & Me.ViewState("OldItemVar1") & "%" & "=" & Math.Round(CDbl(drv("GNum")) * Int(drv("GPrice") * Int(Me.ViewState("OldItemVar1")) / 100), 2)
                            Case Else
                                e.Item.Cells(2).Text = drv("GNum") & "人*" & drv("GPrice") & "=" & Math.Round(CDbl(drv("GNum")) * Int(drv("GPrice")), 2)
                        End Select

                    End If
                    If Not IsDBNull(drv("SNum")) And Not IsDBNull(drv("SPrice")) Then
                        Select Case Convert.ToString(sm.UserInfo.TPlanID)
                            Case "23", "34", "41" '設定核銷%數 核銷數
                                e.Item.Cells(3).Text = drv("SNum") & "人*" & drv("SPrice") & "*" & Me.ViewState("OldItemVar2") & "%" & "=" & Math.Round(CDbl(drv("SNum")) * Int(drv("SPrice") * Int(Me.ViewState("OldItemVar2")) / 100), 2)
                            Case Else
                                e.Item.Cells(3).Text = drv("SNum") & "人*" & drv("SPrice") & "=" & Math.Round(CDbl(drv("SNum")) * Int(drv("SPrice")), 2)
                        End Select

                    End If

                Else                                        '如果是新資料
                    If Not IsDBNull(drv("GNum")) And Not IsDBNull(drv("GPrice")) Then
                        Select Case Convert.ToString(sm.UserInfo.TPlanID)
                            Case "23", "34", "41" '設定核銷%數 核銷數
                                e.Item.Cells(2).Text = drv("GNum") & "人*" & drv("GPrice") & "*" & ItemVar1.Value & "%" & "=" & Math.Round(CDbl(drv("GNum")) * Int(drv("GPrice") * Int(ItemVar1.Value) / 100), 2)
                            Case Else
                                e.Item.Cells(2).Text = drv("GNum") & "人*" & drv("GPrice") & "=" & Math.Round(CDbl(drv("GNum")) * Int(drv("GPrice")), 2)
                        End Select

                    End If
                    If Not IsDBNull(drv("SNum")) And Not IsDBNull(drv("SPrice")) Then
                        Select Case Convert.ToString(sm.UserInfo.TPlanID)
                            Case "23", "34", "41" '設定核銷%數 核銷數
                                e.Item.Cells(3).Text = drv("SNum") & "人*" & drv("SPrice") & "*" & ItemVar2.Value & "%" & "=" & Math.Round(CDbl(drv("SNum")) * Int(drv("SPrice") * Int(ItemVar2.Value) / 100), 2)
                            Case Else
                                e.Item.Cells(3).Text = drv("SNum") & "人*" & drv("SPrice") & "=" & Math.Round(CDbl(drv("SNum")) * Int(drv("SPrice")), 2)
                        End Select

                    End If

                End If

                btn.Attributes("onclick") = "return confirm('您確定要刪除這一筆資料?');"
                btn.CommandArgument = drv("BCCID")
        End Select
    End Sub

    Private Sub DataGrid5_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid5.ItemCommand
        Dim sql As String

        sql = "DELETE Budget_ClassCancel WHERE BCCID='" & e.CommandArgument & "'"
        DbAccess.ExecuteNonQuery(sql, objconn)

        CreateCancelData()
        Common.MessageBox(Me, "刪除成功")
    End Sub

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        Dim dt As DataTable
        Dim sql As String
        sql = "SELECT * FROM Class_StudentsOfClass WHERE OCID='" & Request("OCID") & "' and StudStatus IN (1,4,5)"
        dt = DbAccess.GetDataTable(sql, objconn)

        CreateDGData(dt)
    End Sub

    '新增-方法4
    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        Dim sql As String
        Dim dt As DataTable
        Dim dr As DataRow
        Dim da As SqlDataAdapter = Nothing
        Dim TNum1 As Integer
        Dim TPrice1 As Integer
        Dim TNum2 As Integer
        Dim TPrice2 As Integer
        'Dim conn As SqlConnection
        'TIMS.TestDbConn(Me, conn, True)
        Dim conn As SqlConnection = DbAccess.GetConnection()
        Call TIMS.OpenDbConn(conn)

        sql = "SELECT * FROM Class_StudentsOfClass WHERE OCID='" & Request("OCID") & "' and StudStatus IN (1,4,5)"
        dt = DbAccess.GetDataTable(sql)
        TNum1 = dt.Select("BudgetID='02'").Length
        TPrice1 = dt.Select("BudgetID='02' and RelClass_Unit IN ('100','110','101','111')").Length * 110 * 10 + dt.Select("BudgetID='02' and RelClass_Unit IN ('010','011','110','111')").Length * 110 * 12 + dt.Select("BudgetID='02' and RelClass_Unit IN ('001','101','011','111')").Length * 110 * 24
        TNum2 = dt.Select("BudgetID='03'").Length
        TPrice2 = dt.Select("BudgetID='03' and RelClass_Unit IN ('100','110','101','111')").Length * 110 * 10 + dt.Select("BudgetID='03' and RelClass_Unit IN ('010','011','110','111')").Length * 110 * 12 + dt.Select("BudgetID='03' and RelClass_Unit IN ('001','101','011','111')").Length * 110 * 24
        sql = "SELECT * FROM Budget_ClassCancel WHERE 1<>1"
        '2006/03/28 add conn by matt
        dt = DbAccess.GetDataTable(sql, da, conn)

        dr = dt.NewRow
        dt.Rows.Add(dr)
        dr("OCID") = Request("OCID")
        dr("CancelDate") = CancelDate4.Text
        dr("CancelMode") = 4
        dr("Times") = Times3.Text
        dr("TNum1") = TNum1
        dr("TPrice1") = TPrice1
        dr("TNum2") = TNum2
        dr("TPrice2") = TPrice2
        dr("CancelCost") = TPrice1 + TPrice2
        dr("PlanType") = PlanType4.SelectedValue
        dr("Note") = IIf(Note4.Text = "", Convert.DBNull, Note4.Text)
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now

        DbAccess.UpdateDataTable(dt, da)
        'If conn.State = ConnectionState.Open Then conn.Close()
        Call TIMS.CloseDbConn(conn)

        Button18_Click(sender, e)
        CreateCancelData()
    End Sub

    '回上一頁
    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        Session("BudgetSearchStr") = Me.ViewState("BudgetSearchStr")
        TIMS.Utl_Redirect1(Me, "CM_01_001.aspx?ID=" & Request("ID") & "")
    End Sub

    Private Sub DataGrid4_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid4.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = "CM_TD1"
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim btn As Button = e.Item.FindControl("Button19")

                If e.Item.ItemType = ListItemType.Item Then
                    e.Item.CssClass = "CM_TD2"
                End If

                btn.Attributes("onclick") = "return confirm('您確定要刪除這一筆資料?');"
                btn.CommandArgument = drv("BCCID")
        End Select
    End Sub

    Private Sub DataGrid4_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid4.ItemCommand
        Dim sql As String

        sql = "DELETE Budget_ClassCancel WHERE BCCID='" & e.CommandArgument & "'"
        DbAccess.ExecuteNonQuery(sql, objconn)

        Button18_Click(Button18, Nothing)
        CreateCancelData()
        Common.MessageBox(Me, "刪除成功")
    End Sub
End Class
