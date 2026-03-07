Partial Class SD_13_001_00
    Inherits AuthBasePage

    Const cst_學號 As Integer = 0
    Const cst_姓名 As Integer = 1
    Const cst_身分證號碼 As Integer = 2
    'Const cst_是否獲得學分 = 3
    Const cst_出席達3分之2 As Integer = 4
    Const cst_是否補助 As Integer = 5
    'Const cst_補助比例 = 6
    Const cst_總費用 As Integer = 7
    Const cst_補助費用 As Integer = 8
    Const cst_個人支付 As Integer = 9
    Const cst_剩餘可用餘額 As Integer = 10
    Const cst_其他申請中金額 As Integer = 11
    'Const cst_是否提出申請  As Integer = 12
    Const cst_申請狀態 As Integer = 13
    Const cst_撥款狀態 As Integer = 14
    Const cst_預算別 As Integer = 15
    '年度小於等於2011啟用。

#Region "Functions"

    Function Check_AppliedResultR(ByVal tmpOCID As String) As Boolean
        Dim rst As Boolean = False
        Dim sqlAdp As New SqlDataAdapter
        Dim sqlStr As String = ""

        Try
            If objconn.State = ConnectionState.Closed Then objconn.Open()
            Dim tAppliedResultR As String = ""
            sqlStr = "select ISNULL(AppliedResultR,'N') as AppliedResultR from Class_ClassInfo where OCID= @OCID"
            sqlAdp.SelectCommand = New SqlCommand(sqlStr, objconn)
            sqlAdp.SelectCommand.Parameters.Clear()
            sqlAdp.SelectCommand.Parameters.Add("OCID", SqlDbType.VarChar).Value = tmpOCID
            tAppliedResultR = sqlAdp.SelectCommand.ExecuteScalar()
            If tAppliedResultR = "Y" Then rst = True
        Catch ex As Exception
            Dim strErrmsg As String = ""
            strErrmsg += "/*  ex.ToString: */" & vbCrLf
            strErrmsg += ex.ToString & vbCrLf
            strErrmsg += "/* sqlstr: */" & vbCrLf
            strErrmsg += sqlStr & vbCrLf
            'strErrmsg += TIMS.GetErrorMsg(MyPage) '取得錯誤資訊寫入
            strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)
        End Try

        Return rst
    End Function

    Private Function Check_IsClosed(ByVal tmpID As Integer) As Boolean
        Dim rst As Boolean = False
        Dim sqlAdp As New SqlDataAdapter
        Dim sqlStr As String = ""

        Try
            If objconn.State = ConnectionState.Closed Then objconn.Open()
            'Dim sqlStr As String = String.Empty
            sqlStr = "select ISNULL(IsClosed,'N') as IsClosed from Class_ClassInfo where OCID= @OCID "
            With sqlAdp
                .SelectCommand = New SqlCommand(sqlStr, objconn)
                .SelectCommand.Parameters.Clear()
                .SelectCommand.Parameters.Add("OCID", SqlDbType.Int).Value = tmpID
                If .SelectCommand.ExecuteScalar() = "Y" Then rst = True
            End With
        Catch ex As Exception
            Dim strErrmsg As String = ""
            strErrmsg += "/*  ex.ToString: */" & vbCrLf
            strErrmsg += ex.ToString & vbCrLf
            strErrmsg += "/* sqlstr: */" & vbCrLf
            strErrmsg += sqlStr & vbCrLf
            'strErrmsg += TIMS.GetErrorMsg(MyPage) '取得錯誤資訊寫入
            strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)
            Common.MessageBox(Me, ex.ToString)
        End Try
        Return rst
    End Function

#End Region

    Dim gsBlackIDNO As String = "" '學員處分
    Dim objconn As SqlConnection

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

#Region "(No Use)"

        'Select Case sm.UserInfo.Years
        '    Case Is <= "2006"
        '        Server.Transfer("SD_13_001_95.aspx?ID=" & Request("ID"))
        '        Exit Sub
        '    Case Is <= "2011"
        '        Server.Transfer("SD_13_001_00.aspx?ID=" & Request("ID"))
        '        Exit Sub
        'End Select

#End Region

        If Not IsPostBack Then
            msg.Text = ""
            DataGridTable.Style("display") = "none"
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
            TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
        End If

#Region "(No Use)"

        'If sm.UserInfo.LID = 1 Then
        '    '分署(中心)審核
        '    Me.DataGrid1.Columns(cst_預算別).Visible = True
        'Else
        '    '委訓單位不顯示 cst_預算別
        '    Me.DataGrid1.Columns(cst_預算別).Visible = False
        'End If

#End Region

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');SetOneOCID();"
        Button2.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

        Button1.Attributes("onclick") = "return CheckSearch();"
        Button3.Attributes("onclick") = "return CheckData();"
    End Sub

    '查詢 (SQL)
    Sub Search1()
        '20090907 針對尚未完成班級結訓動作的班級，是不能查出資料的。
        msg.Text = ""
        If Not Check_AppliedResultR(OCIDValue1.Value) Then
            DataGridTable.Style("display") = "none"
            msg.Text = "查無資料"
            Common.MessageBox(Me, "該班學員資料複審結果尚未通過")
            Exit Sub
        End If

        If Check_IsClosed(OCIDValue1.Value) = True Then
            hidBlackMsg.Value = "" '清空黑名單暫存記錄(2009/07/28 判斷黑名單)
            Dim dt As DataTable
            Dim dr As DataRow

            Dim sqlstr As String = ""
            sqlstr = "" & vbCrLf
            sqlstr &= " SELECT a.ocid ,d.setid ,d.SOCID ,d.IdentityID ,dbo.FN_CSTUDID2(d.StudentID) StudentID " & vbCrLf
            sqlstr &= "  ,ISNULL(g.BudID ,d.BudgetID) BudgetID ,d.SupplyID ESupplyID " & vbCrLf
            sqlstr &= "  ,dbo.FN_GET_GOVCNT(d.SOCID) GovCnt " & vbCrLf
            sqlstr &= "  ,e.Name ,e.IDNO " & vbCrLf
            sqlstr &= "  ,d.CreditPoints " & vbCrLf
            '除數可能有溢位問題，無條件捨去餘2位數。
            sqlstr &= "  ,CASE WHEN b.TotalCost >= ISNULL(c.Total2,0) THEN FLOOR(ISNULL(b.TotalCost,0)/ISNULL(b.TNum,1)) ELSE FLOOR(ISNULL(c.Total2,0)/ISNULL(b.TNum,1) ) END Total " & vbCrLf
            sqlstr &= "  ,a.THours ,ar.DistID " & vbCrLf
            sqlstr &= "  ,ISNULL(f.CountHours,0) CountHours " & vbCrLf
            sqlstr &= "  ,e.DegreeID ,d.StudStatus " & vbCrLf
            sqlstr &= "  ,d.MIdentityID " & vbCrLf
            sqlstr &= "  ,a.STDate " & vbCrLf
            sqlstr &= "  ,a.AppliedResultM " & vbCrLf
            sqlstr &= "  ,d.AppliedResult " & vbCrLf
            sqlstr &= "  ,g.SOCID Exist " & vbCrLf
            sqlstr &= "  ,g.SumOfMoney " & vbCrLf
            sqlstr &= "  ,g.PayMoney " & vbCrLf
            sqlstr &= "  ,g.AppliedStatus " & vbCrLf
            sqlstr &= "  ,g.AppliedNote " & vbCrLf
            sqlstr &= "  ,g.SupplyID " & vbCrLf
            '其他申請中金額
            'sqlstr &= "  ,dbo.FN_GET_GOVAPPL2(e.IDNO,a.STDate) GovAppl2 " & vbCrLf
            sqlstr &= "  ,dbo.FN_GET_GOVCOST2(e.IDNO, convert(varchar,a.STDate,111)) GovAppl2"
            '其他申請中金額(並排除本班)
            'sqlstr += " ,dbo.FN_GET_GOVAPPL22(e.IDNO,a.STDate,a.OCID) GovAppl2 " & vbCrLf
            sqlstr &= "  ,g.AppliedStatusM " & vbCrLf
            sqlstr &= " FROM Class_ClassInfo a " & vbCrLf
            sqlstr &= " JOIN Plan_PlanInfo b ON a.PlanID = b.PlanID AND a.ComIDNO = b.ComIDNO AND a.SeqNo = b.SeqNo " & vbCrLf
            sqlstr &= " JOIN Auth_Relship ar ON a.RID = ar.RID " & vbCrLf
            sqlstr &= " LEFT JOIN (" & vbCrLf
            sqlstr &= "   SELECT PlanID ,ComIDNO ,SeqNo ,SUM(ISNULL(OPrice,1)*ISNULL(ItemCost,1)) Total ,SUM(ISNULL(OPrice,1)*ISNULL(Itemage,1)) Total2 " & vbCrLf
            sqlstr &= "   FROM Plan_CostItem " & vbCrLf
            If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                sqlstr &= " WHERE COSTMODE = 5 " & vbCrLf
            Else
                sqlstr &= " WHERE COSTMODE <> 5 " & vbCrLf
            End If
            sqlstr &= " Group By PlanID ,ComIDNO ,SeqNo) c ON a.PlanID = c.PlanID AND a.ComIDNO = c.ComIDNO AND a.SeqNo = c.SeqNo " & vbCrLf
            sqlstr &= " JOIN Class_StudentsOfClass d ON a.OCID = d.OCID " & vbCrLf
            sqlstr &= " JOIN Stud_StudentInfo e ON d.SID = e.SID " & vbCrLf
            sqlstr &= " LEFT JOIN (SELECT SOCID ,SUM(Hours) CountHours FROM Stud_Turnout2 GROUP BY SOCID) f ON d.SOCID = f.SOCID " & vbCrLf
            sqlstr &= " LEFT JOIN Stud_SubsidyCost g ON d.SOCID = g.SOCID " & vbCrLf
            sqlstr &= " WHERE a.OCID = '" & OCIDValue1.Value & "' AND a.AppliedResultR = 'Y' " & vbCrLf '產業人才投資方案 Y:通過 C:全班學員資料確認
            ' order by StudentID ASC"
            If InStr(Me.ViewState("sort"), "IDNO") > 0 Then
                sqlstr &= " ORDER BY e." & Me.ViewState("sort").ToString
            ElseIf InStr(Me.ViewState("sort"), "StudentID") > 0 Then
                sqlstr &= " ORDER BY dbo.FN_CSTUDID2(d.StudentID) " & Replace(Me.ViewState("sort").ToString, "StudentID", "") & vbCrLf
            Else
                sqlstr &= " ORDER BY dbo.FN_CSTUDID2(d.StudentID) " & vbCrLf
            End If

            Try
                'SQL
                dt = DbAccess.GetDataTable(sqlstr, objconn)
            Catch ex As Exception
                Dim strErrmsg As String = ""
                strErrmsg += "/*  ex.ToString: */" & vbCrLf
                strErrmsg += ex.ToString & vbCrLf
                strErrmsg += "/* sqlstr: */" & vbCrLf
                strErrmsg += sqlstr & vbCrLf
                'strErrmsg += TIMS.GetErrorMsg(MyPage) '取得錯誤資訊寫入
                strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                Call TIMS.WriteTraceLog(strErrmsg)
                Common.MessageBox(Me, "資料庫效能異常，請重新查詢")
                Exit Sub
            End Try

#Region "(No Use)"

            'START 黑名單之身分證記錄 2009/07/28 by waiming
            'Dim txtIDNO As String
            'Dim dt_Blacklist As DataTable
            'Dim sql_Blacklist As String
            'sql_Blacklist = "" & vbCrLf
            'sql_Blacklist += " select IDNO" & vbCrLf
            'sql_Blacklist += " from Stud_Blacklist" & vbCrLf
            'sql_Blacklist += " where Avail='Y'" & vbCrLf
            'sql_Blacklist += " and getdate() between SBSdate and DATEADD(month, 12*SBYears, SBSdate)" & vbCrLf
            'dt_Blacklist = DbAccess.GetDataTable(sql_Blacklist, objconn)
            'txtIDNO = ""
            'For i As Int16 = 0 To dt_Blacklist.Rows.Count - 1
            '    If dt.Select("IDNO='" & Convert.ToString(dt_Blacklist.Rows(i)("IDNO")) & "'").Length > 0 Then
            '        If txtIDNO <> "" Then
            '            If txtIDNO.IndexOf(Convert.ToString(dt_Blacklist.Rows(i)("IDNO"))) = -1 Then
            '                If txtIDNO <> "" Then txtIDNO &= ","
            '                txtIDNO &= Convert.ToString(dt_Blacklist.Rows(i)("IDNO"))
            '            End If
            '        Else
            '            If txtIDNO <> "" Then txtIDNO &= ","
            '            txtIDNO &= Convert.ToString(dt_Blacklist.Rows(i)("IDNO"))
            '        End If
            '    End If
            'Next

#End Region

            Dim stdBLACK2TPLANID As String = ""
            Dim iStdBlackType As Integer = TIMS.Chk_StdBlackType(Me, objconn, stdBLACK2TPLANID)
            gsBlackIDNO = TIMS.Get_StdBlackIDNO(Me, iStdBlackType, stdBLACK2TPLANID, objconn) '學員處分
            'Me.ViewState("BlackIDNO") = txtIDNO
            'END 黑名單之身分證記錄

            DataGridTable.Style("display") = "none"
            msg.Text = "查無資料"
            If dt.Rows.Count > 0 Then
                '顯示有效資料
                DataGridTable.Style("display") = "inline"
                msg.Text = ""
                If ViewState("sort") = "" Then ViewState("sort") = "StudentID"
                DataGrid1.DataKeyField = "SOCID"
                DataGrid1.DataSource = dt
                DataGrid1.DataBind()

                '檢查是否有重複參訓學員排除產學訓計畫
                dr = dt.Rows(0)
                Button3.Enabled = True '儲存鈕
                If dr("AppliedResultM").ToString = "Y" Then
                    'Button3.Enabled = False '永遠可申請
                    TIMS.Tooltip(Button3, "班級學員經費審核結果，已完成")
                End If
                dt = TIMS.GET_Duplicate_Student(OCIDValue1.Value, 1, objconn) '檢查是否有重複參訓學員排除產學訓計畫
                If dt IsNot Nothing Then '有重複參訓學員
                    Dclass.Value = 1
                Else
                    Dclass.Value = 2 '沒重複
                End If

#Region "(No Use)"

                '20080606 Andy
                'Dim da As SqlDataAdapter = nothing
                'Dim dr2 As DataRow
                'Dim conn As SqlConnection = DbAccess.GetConnection
                'sql = "SELECT * FROM Stud_SubsidyCost WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID='" & OCIDValue1.Value & "' and  SupplyID=9 )"
                'dt = DbAccess.GetDataTable(sql, da, conn)
                'For Each dr2 In dt.Rows
                '    dr2("PayMoney") = dr("total")
                '    dr2("SumOfMoney") = 0
                'Next
                'DbAccess.UpdateDataTable(dt, da)

#End Region
            End If
        Else
            DataGridTable.Style("display") = "none"
            msg.Text = "(本班尚未完成結訓動作)查無資料"
        End If
    End Sub

    '查詢
    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call Search1() '查詢 (SQL)
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Const cst_總費用_msg As String = "非學分班的訓練費用項目"
        Const cst_是否補助_msg As String = "是否有請領補助津貼的資格"
        Const cst_補助費用_msg As String = "預定要補助的金額(可自行變動，未申請前系統會根據可用餘額推算)"
        Const cst_個人支付_msg As String = "學員自行要支付的金額(會根據補助費用所輸入的值來調動)"
        'Const cst_剩餘可用餘額_msg = "學員目前可用餘額-這次預定補助費用的剩餘金額(成為負數時會以紅字表示)"
        Const cst_剩餘可用餘額_msg As String = "學員目前可用餘額-已審核通過費用的剩餘金額(成為負數時會以紅字表示)"
        Const cst_目前申請總額_msg As String = "學員目前已申請未審核補助金總額(含本次補助費用)合併後 超過剩餘可用餘額以紅字表示" '其他申請中金額
        'Dim totalleft As Decimal = 0   ' 剩餘可用餘額   
        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = "SD_TD1"
                'e.Item.Cells(5).ToolTip = "是否有請領補助津貼的資格"
                'e.Item.Cells(7).ToolTip = "預定要補助的金額(可自行變動，未申請前系統會根據可用餘額推算)"
                'e.Item.Cells(8).ToolTip = "學員自行要支付的金額(會根據補助費用所輸入的值來調動)"
                'e.Item.Cells(9).ToolTip = "學員目前可用餘額-這次預定補助費用的剩餘金額(成為負數時會以紅字表示)"
                e.Item.Cells(cst_是否補助).ToolTip = cst_是否補助_msg
                e.Item.Cells(cst_總費用).ToolTip = cst_總費用_msg
                e.Item.Cells(cst_補助費用).ToolTip = cst_補助費用_msg
                e.Item.Cells(cst_個人支付).ToolTip = cst_個人支付_msg
                e.Item.Cells(cst_剩餘可用餘額).ToolTip = cst_剩餘可用餘額_msg
                e.Item.Cells(cst_其他申請中金額).ToolTip = cst_目前申請總額_msg '其他申請中的金額
                If Me.ViewState("sort") <> "" Then
                    'Dim mylabel As String
                    Dim mysort As New System.Web.UI.WebControls.Image
                    Dim i As Integer = -1
                    Select Case Me.ViewState("sort")
                        Case "StudentID", "StudentID DESC"
                            'mylabel = "IDNO"
                            i = 0
                            If Me.ViewState("sort") = "StudentID" Then
                                mysort.ImageUrl = "../../images/SortUp.gif"
                            Else
                                mysort.ImageUrl = "../../images/SortDown.gif"
                            End If
                        Case "IDNO", "IDNO DESC"
                            'mylabel = "StudentID"
                            i = 2
                            If Me.ViewState("sort") = "IDNO" Then
                                mysort.ImageUrl = "../../images/SortUp.gif"
                            Else
                                mysort.ImageUrl = "../../images/SortDown.gif"
                            End If
                    End Select
                    If i <> -1 Then
                        e.Item.Cells(i).Controls.Add(mysort)
                    End If
                End If
            Case ListItemType.Item, ListItemType.AlternatingItem
                If e.Item.ItemType = ListItemType.Item Then e.Item.CssClass = "SD_TD2"
                Dim Flag As Integer = 0  '得到學分  
                Dim FlagStudy As Integer = 0 '出席達3/2
                Dim drv As DataRowView = e.Item.DataItem
                Dim SupplyID As DropDownList = e.Item.FindControl("SupplyID") 'ESupplyID 'SupplyID.Enabled
                Dim BudID As DropDownList = e.Item.FindControl("BudID") 'BudgetID 'BudID.Enabled

                '補助比例和預算別改唯讀
                SupplyID.Enabled = False '補助比例
                BudID.Enabled = False '預算別'暫設不可更改 預算別，根據審核狀況來開放 預算別
                Dim DataGrid2 As DataGrid = e.Item.FindControl("DataGrid2")
                Dim CreditPoints As Label = e.Item.FindControl("CreditPoints")
                Dim SumOfMoney As TextBox = e.Item.FindControl("SumOfMoney")
                Dim Checkbox1 As HtmlInputCheckBox = e.Item.FindControl("Checkbox1")
                Dim RemainSub As HtmlInputHidden = e.Item.FindControl("RemainSub")
                Dim MaxSub As HtmlInputHidden = e.Item.FindControl("MaxSub")
                Dim PayMoney As HtmlInputHidden = e.Item.FindControl("PayMoney")
                Dim balancemoney As HtmlInputHidden = e.Item.FindControl("balancemoney")

                '20090201 andy edit
                '-------------
                Dim setid As HtmlInputHidden = e.Item.FindControl("setid")
                Dim ocid As HtmlInputHidden = e.Item.FindControl("ocid")
                Dim hid_socid As HtmlInputHidden = e.Item.FindControl("socid")
                ocid.Value = drv("ocid").ToString
                setid.Value = drv("setid").ToString
                hid_socid.Value = drv("setid").ToString
                '-------------

                Dim star As TextBox = e.Item.FindControl("star1")
                Dim stud As TextBox = e.Item.FindControl("stud1")
                star.Visible = True
                stud.Visible = True
                SupplyID = TIMS.Get_SupplyID(SupplyID)

                e.Item.Cells(cst_是否補助).ToolTip = cst_是否補助_msg
                e.Item.Cells(cst_總費用).ToolTip = cst_總費用_msg
                e.Item.Cells(cst_補助費用).ToolTip = cst_補助費用_msg
                e.Item.Cells(cst_個人支付).ToolTip = cst_個人支付_msg
                e.Item.Cells(cst_剩餘可用餘額).ToolTip = cst_剩餘可用餘額_msg
                e.Item.Cells(cst_其他申請中金額).ToolTip = cst_目前申請總額_msg
                If drv("IDNO").ToString <> "" Then
                    e.Item.Cells(cst_姓名).ToolTip = TIMS.Search_Stud_SubsidyCost(drv("IDNO").ToString, objconn)
                    e.Item.Cells(cst_學號).ToolTip = e.Item.Cells(cst_姓名).ToolTip
                    e.Item.Cells(cst_身分證號碼).ToolTip = e.Item.Cells(cst_姓名).ToolTip
                End If
                If drv("ESupplyID").ToString <> "" Then Common.SetListItem(SupplyID, drv("ESupplyID").ToString)

                '判斷是否填寫問卷
                If drv("GovCnt") = 0 Then star.Text = "*" Else star.Text = ""
                stud.Text = drv("StudentID").ToString

                BudID = TIMS.Get_Budget(BudID, 2)
                If drv("DistID").ToString <> "001" Then BudID.Items.Remove(BudID.Items.FindByValue("01"))

                'If drv("SupplyID").ToString <> "" Then Common.SetListItem(SupplyID, drv("SupplyID").ToString)

                '規則改為show class_studentsofclass.budgetid
                If drv("BudgetID").ToString <> "" Then
                    Common.SetListItem(BudID, drv("BudgetID").ToString)
                End If

#Region "(No Use)"

                'If drv("BudID").ToString <> "" Then
                '    Common.SetListItem(BudID, drv("BudID").ToString)
                'Else
                '    If drv("BudgetID").ToString <> "" Then
                '        Common.SetListItem(BudID, drv("BudgetID").ToString)
                '    End If
                'End If

                '20090123 andy  edit 產投、在職 2009年 身分別為「就業保險被保險人非自願失業者」時
                '1.預算來源設定為 02:就安基金 ； 2.補助比例為100%  
                '090423(將090123修改之程式mark起來)直接帶前端所輸入之原始值即可，不用再做一次判斷重新給值
                '--------------------------  start
                'If sm.UserInfo.TPlanID = 28 Then
                '    If CInt(Me.sm.UserInfo.Years) > 2008 Then
                '        For i As Integer = 0 To Split(Convert.ToString(drv("IdentityID")), ",").Length - 1
                '            If Split(drv("IdentityID").ToString, ",")(i) = "02" Then
                '                Common.SetListItem(BudID, "02")
                '                Common.SetListItem(SupplyID, "2")
                '                drv("ESupplyID") = "2"
                '            End If
                '        Next
                '    End If
                'End If
                '----------------------------- end

#End Region

                If IsDBNull(drv("CreditPoints")) Then
                    CreditPoints.Text = "<font color='RED'>否</font>"
                Else
                    If drv("CreditPoints") Then '是否得到學分
                        CreditPoints.Text = "是"
                        Flag = 1
                    Else
                        CreditPoints.Text = "<font color='RED'>否</font>"
                    End If
                End If

                e.Item.Cells(cst_出席達3分之2).Text = "否"
                If drv("THours") > 0 Then
                    If (drv("THours") - drv("CountHours")) / drv("THours") >= 2 / 3 Then
                        e.Item.Cells(cst_出席達3分之2).Text = "是"
                        FlagStudy = 1
                    End If
                End If

                'Dim sql As String
                'Dim dr As DataRow
                'Dim dt As DataTable
                Dim Total As Integer = 0 '可用補助額
                '可用補助額(2007年3年3萬)
                '可用補助額(2008年3年5萬)
                '可用補助額(2012年3年7萬)
                Total = TIMS.Get_3Y_SupplyMoney(Me)

#Region "(No Use)"

                'If sm.UserInfo.Years < 2008 Then
                '    ''2007年前(含2007)
                '    Total = 30000
                'Else
                '    ''2008年後(含2008)
                '    Total = 50000
                'End If
                'Total = Total - drv("GovCost") '可用補助額-'政府已補助費用(三年一段補助)

#End Region

                '20080609  Andy  可用補助額
                Dim SubsidyCost As Double
                SubsidyCost = TIMS.Get_SubsidyCost(drv("IDNO").ToString(), drv("STDate").ToString(), "", "Y", objconn)
                Total -= SubsidyCost
                If Total < 0 Then Total = 0

                RemainSub.Value = Total '30000 '可用補助額

                'Dim ESupplyPercent, SupplyPercent As Double
                If IsDBNull(drv("Exist")) Then      '表示沒資料,以新增的型態顯示
                    ' If Flag = 1 Then '得到學分  
                    If (Flag = 1 And FlagStudy = 1) Then   '970513 Andy  得到學分且出席達2/3  
                        e.Item.Cells(cst_是否補助).Text = "是"
                        If drv("MIdentityID").ToString <> "" Then
#Region "(No Use)"

                            '20080806 andy 原程式補助費用判斷有納入主要身分為一般身分別時只能補助80%的條件==>改為只依據產學訓(補助比例代碼)來做判斷
                            '-----------------------------------------------------
                            'If drv("MIdentityID").ToString = "01" Then '一般身分者
                            '    'drv("Total") 此次費用
                            '    If Total >= Decimal.Truncate(drv("Total") * 0.8) Then '可用補助額 > '計算補助費用
                            '        SumOfMoney.Text = Decimal.Truncate(drv("Total") * 0.8) '此次可用補助額=(課程費用*0.8)計算補助費用 
                            '    Else
                            '        SumOfMoney.Text = Total '此次可用補助額=可用補助額 
                            '    End If
                            'Else '其他身分者
                            '    If Total >= drv("Total") Then '可用補助額 > '課程費用*1(計算補助費用)
                            '        SumOfMoney.Text = drv("Total") '此次可用補助額=課程費用 
                            '    Else
                            '        SumOfMoney.Text = Total '此次可用補助額=可用補助額 
                            '    End If
                            'End If
                            '--------   Start
                            'Select Case drv("MIdentityID")  '主要身分代碼
                            '    Case "01"                  '一般身分者,補助比例80%    
                            '        SupplyPercent = 0.8
                            '    Case Else                   '其它身分,補助比例100%
                            '        SupplyPercent = 1
                            'End Select

#End Region

                            Dim ESupplyPercent As Double = 0
                            '(有其他狀況)暫'補助比例0%
                            If drv("ESupplyID").ToString <> "" Then
                                Select Case drv("ESupplyID")  '產學訓(補助比例代碼)
                                    Case 1  '補助比例80%
                                        ESupplyPercent = 0.8
                                    Case 2  '補助比例100%
                                        ESupplyPercent = 1
                                    Case 9  '補助比例0%
                                        ESupplyPercent = 0
                                    Case Else '(有其他狀況)暫'補助比例0%
                                        ESupplyPercent = 0
                                End Select
                            End If

                            If Total >= Decimal.Truncate(drv("Total") * ESupplyPercent) Then
                                SumOfMoney.Text = Decimal.Truncate(drv("Total") * ESupplyPercent)
                            Else
                                SumOfMoney.Text = Total '此次可用補助額=可用補助額 
                            End If
                            '------   End
                            MaxSub.Value = SumOfMoney.Text '此次最大可用補助額
                            e.Item.Cells(cst_個人支付).Text = CStr(CInt(drv("Total")) - CInt(IIf(Trim(SumOfMoney.Text) = "", "0", Trim(SumOfMoney.Text))))   '課程費用-'可用補助額=個人支付費用
                            PayMoney.Value = CInt(drv("Total")) - CInt(SumOfMoney.Text) '課程費用-'可用補助額=個人支付費用
                            e.Item.Cells(cst_剩餘可用餘額).Text = Total.ToString()  '可用補助額=剩餘可用餘額
                            'totalleft = Total  '20091221 andy 
                        Else
                            SumOfMoney.Enabled = False '不可填入補助費用
                            Checkbox1.Disabled = True '不可提出申請
                            'e.Item.Cells(cst_是否補助).Text = "<font color='RED'>否</font>"
                            e.Item.Cells(cst_剩餘可用餘額).Text = Total
                            'totalleft = Total  '20091221 andy 
                        End If
                    Else
                        SumOfMoney.Enabled = False '不可填入補助費用
                        Checkbox1.Disabled = True  '不可提出申請
                        Checkbox1.Checked = False
                        e.Item.Cells(cst_是否補助).Text = "<font color='RED'>否</font>"
                        e.Item.Cells(cst_剩餘可用餘額).Text = Total
                        '20080606 andy 是否補助為「否」,個人支付=總費用
                        e.Item.Cells(cst_個人支付).Text = e.Item.Cells(cst_總費用).Text
                        SumOfMoney.Text = "0"
                        'totalleft = Total  '20091221 andy 
                    End If
                Else
                    ' If Flag = 1 Then '得到學分
                    If (Flag = 1 And FlagStudy = 1) Then '970513 Andy  得到學分且出席達2分之3  
                        e.Item.Cells(cst_是否補助).Text = "是"
                    Else '尚未得到學分
                        e.Item.Cells(cst_是否補助).Text = "<font color='RED'>否</font>"
                        '20080606 andy 是否補助為「否」,個人支付=總費用
                        e.Item.Cells(cst_個人支付).Text = e.Item.Cells(cst_總費用).Text
                        SumOfMoney.Text = "0"
                    End If
#Region "(No Use)"

                    'If drv("MIdentityID").ToString = "01" Then  '身分別,01:一般身分者 
                    '    If Total >= Decimal.Truncate(drv("Total") * 0.8) Then
                    '        MaxSub.Value = Decimal.Truncate(drv("Total") * 0.8) '此次最大可用補助額
                    '    Else
                    '        MaxSub.Value = Total '此次最大可用補助額
                    '    End If
                    'Else
                    '    If Total >= drv("Total") Then
                    '        MaxSub.Value = drv("Total")
                    '    Else
                    '        MaxSub.Value = Total
                    '    End If
                    'End If

#End Region
                    '20080806  Andy 原程式補助費用判斷有納入主要身分為一般身分別時只能補助80%的條件==>改為只依據產學訓(補助比例代碼)來做判斷
                    '-----------   Start
                    Dim ESupplyPercent As Double = 0
                    '(有其他狀況)暫'補助比例0%
                    Select Case drv("ESupplyID")  '產學訓(補助比例代碼)
                        Case 1  '補助比例80%
                            ESupplyPercent = 0.8
                        Case 2  '比例100%
                            ESupplyPercent = 1
                        Case 9  '比例0%
                            ESupplyPercent = 0
                    End Select
                    If Total >= Decimal.Truncate(drv("Total") * ESupplyPercent) Then
                        MaxSub.Value = Decimal.Truncate(drv("Total") * ESupplyPercent)
                    Else
                        MaxSub.Value = Total '此次可用補助額=可用補助額 
                    End If
                    '-----------   End
                    SumOfMoney.Text = drv("SumOfMoney").ToString '可用補助額
                    PayMoney.Value = drv("PayMoney").ToString '個人支付費用
                    e.Item.Cells(cst_個人支付).Text = drv("PayMoney").ToString '個人支付費用
                    Checkbox1.Checked = True '有提出申請
                    'BudID.Enabled = False '暫設不可更改 預算別，根據審核狀況來開放 預算別
                    If IsDBNull(drv("AppliedStatusM")) Then
                        Checkbox1.Disabled = False '可更改提出申請
                        e.Item.Cells(cst_申請狀態).Text = "審核中"
                        e.Item.Cells(cst_撥款狀態).Text = "未撥款"
                        'BudID.Enabled = True '可更改 預算別
                    Else
#Region "(No Use)"

                        'If drv("AppliedStatusM") Then
                        '    Checkbox1.Disabled = True
                        '    SumOfMoney.ReadOnly = True

                        '    e.Item.Cells(cst_申請狀態).Text = "審核通過"
                        'Else
                        '    Checkbox1.Disabled = False
                        '    e.Item.Cells(cst_申請狀態).Text = "審核失敗"
                        'End If

#End Region
                        Select Case drv("AppliedStatusM").ToString
                            Case "Y"
                                Checkbox1.Disabled = True
                                SumOfMoney.ReadOnly = True
                                e.Item.Cells(cst_申請狀態).Text = "審核通過"
                                If IsDBNull(drv("AppliedStatus")) Then '撥款審核狀態
                                    'Checkbox1.Disabled = True 'False '審核通過後不可放棄申請
                                    e.Item.Cells(cst_撥款狀態).Text = "撥款中" '"撥款審核中"
                                Else
                                    If drv("AppliedStatus") Then '=1
                                        'Checkbox1.Disabled = True '審核通過後不可放棄申請
                                        'SumOfMoney.ReadOnly = True
                                        e.Item.Cells(cst_撥款狀態).Text = "已撥款"
                                    Else
                                        'Checkbox1.Disabled = False
                                        e.Item.Cells(cst_撥款狀態).Text = "不撥款"
                                    End If
                                End If
                            Case "N"
                                Checkbox1.Disabled = False '審核失敗  提出申請
                                e.Item.Cells(cst_申請狀態).Text = "審核不通過" '"審核失敗"
                                e.Item.Cells(cst_撥款狀態).Text = "不撥款"
                            Case "R"
                                Checkbox1.Disabled = False '退件修正  提出申請
                                e.Item.Cells(cst_申請狀態).Text = "退件修正"
                                e.Item.Cells(cst_撥款狀態).Text = "未撥款"
                                'BudID.Enabled = True '可更改 預算別
                            Case ""
                                Checkbox1.Disabled = False
                        End Select
                    End If
                    If Not e.Item.Cells(cst_申請狀態).Text = "審核通過" Then
                        If Total - CInt(SumOfMoney.Text) >= 0 Then
                            'e.Item.Cells(cst_剩餘可用餘額).Text = Total - cint(SumOfMoney.Text)
                            e.Item.Cells(cst_剩餘可用餘額).Text = Total
                        Else
                            'e.Item.Cells(cst_剩餘可用餘額).Text = "<font color=Red>" & Total - cint(SumOfMoney.Text) & "</font>"
                            e.Item.Cells(cst_剩餘可用餘額).Text = "<font color=Red>" & Total & "</font>"
                        End If
                        'totalleft = Total  '20091221 andy 
                        If drv("GovAppl2") > Total - CInt(SumOfMoney.Text) Then
                            e.Item.Cells(cst_其他申請中金額).Text = "<font color=Red>" & drv("GovAppl2").ToString & "</font>"
                        End If
                    Else
                        SumOfMoney.Enabled = False
                        'e.Item.Cells(cst_剩餘可用餘額).Enabled = False
                        If Total >= 0 Then
                            e.Item.Cells(cst_剩餘可用餘額).Text = Total
                        Else
                            e.Item.Cells(cst_剩餘可用餘額).Text = "<font color=Red>" & Total & "</font>"
                        End If
                        'totalleft = Total  '20091221 andy 
                        If drv("GovAppl2") > Total Then
                            e.Item.Cells(cst_其他申請中金額).Text = "<font color=Red>" & drv("GovAppl2").ToString & "</font>"
                        End If
                    End If
                End If
                'SupplyID.Enabled = SumOfMoney.Enabled
                'BudID.Enabled = SumOfMoney.Enabled
                '補助比例和預算別改唯讀
                SupplyID.Enabled = False
                BudID.Enabled = False '預算別

                '學員補助不能提出申請
                If (Not IsDBNull(drv("AppliedResult")) And drv("AppliedResult").ToString = "N") Or (SupplyID.SelectedItem.Selected = True And SupplyID.SelectedValue.ToString() = "9") Then         '審核結果 或 補助比例代碼=9 補助比例0%
                    '970513 Andy ,若審核結果為不補助 
                    Checkbox1.Disabled = True
                    '970717 Andy ,若審核結果為不補助則是否提出申請應該是無法勾選的,且為不勾選的
                    Checkbox1.Checked = False
                    If (drv("AppliedResult").ToString = "N") Then e.Item.Cells(cst_是否補助).ToolTip += vbCrLf & "(學員資料審核)預算別：不補助"
                    SumOfMoney.Enabled = False '不可填入補助費用
                    'BudID.Enabled = False '預算別
                    e.Item.Cells(cst_是否補助).Text = "<font color='RED'>否</font>"
                    '20080606 andy 是否補助為「否」,個人支付=總費用
                    e.Item.Cells(cst_個人支付).Text = e.Item.Cells(cst_總費用).Text
                    SumOfMoney.Text = "0"
                End If

                '970513 Andy 學分為 0 或 出勤未滿2/3
                If Flag = 0 Or FlagStudy = 0 Then
                    Checkbox1.Disabled = True
                    Checkbox1.Checked = False
                    e.Item.Cells(cst_是否補助).Text = "<font color='RED'>否</font>"
                    SumOfMoney.Enabled = False '不可填入補助費用
                    'BudID.Enabled = False '預算別
                    '20080606 andy 是否補助為「否」,個人支付=總費用
                    e.Item.Cells(cst_個人支付).Text = e.Item.Cells(cst_總費用).Text
                    SumOfMoney.Text = "0"
                End If

#Region "(No Use)"

                'For i As Integer = 0 To 2
                '    If DataGrid2.Visible Then
                '        e.Item.Cells(i).Attributes("onmouseover") = "if(document.getElementsById('" & DataGrid2.ClientID & "')){document.getElementById('" & DataGrid2.ClientID & "').style.display='inline';}"
                '        e.Item.Cells(i).Attributes("onmouseout") = "if(document.getElementById('" & DataGrid2.ClientID & "')){document.getElementById('" & DataGrid2.ClientID & "').style.display='none';}"
                '        e.Item.Cells(i).Style("CURSOR") = "hand"
                '    End If
                'Next

#End Region

                SumOfMoney.Attributes("onchange") = "ChangeMoney(" & e.Item.ItemIndex + 1 & ",'" & SumOfMoney.ClientID & "','" & RemainSub.ClientID & "','" & MaxSub.ClientID & "','" & PayMoney.ClientID & "');"
                SumOfMoney.Attributes("onblur") = "ChangeMoney(" & e.Item.ItemIndex + 1 & ",'" & SumOfMoney.ClientID & "','" & RemainSub.ClientID & "','" & MaxSub.ClientID & "','" & PayMoney.ClientID & "');"
                SumOfMoney.Attributes("onFocus") = "ChangeMoney(" & e.Item.ItemIndex + 1 & ",'" & SumOfMoney.ClientID & "','" & RemainSub.ClientID & "','" & MaxSub.ClientID & "','" & PayMoney.ClientID & "');"

#Region "(No Use)"

                'START 黑名單為不補助鎖定特定選項 2009/07/28 by waiming
                'Dim arr() As String
                'arr = Split(Me.ViewState("BlackIDNO"), ",")
                'For i As Int16 = 0 To arr.Length - 1
                '    If Convert.ToString(drv("IDNO")) = arr(i) Then
                '    End If
                'Next

#End Region

                If gsBlackIDNO <> "" _
                    AndAlso gsBlackIDNO.IndexOf(Convert.ToString(drv("IDNO"))) > -1 Then
                    e.Item.Cells(cst_是否補助).Text = "<font color='RED'>否</font>" '是否補助
                    SupplyID.SelectedValue = "9" '補助比例
                    SupplyID.Enabled = False
                    SumOfMoney.Text = "0" '補助費用
                    SumOfMoney.Enabled = False
                    Checkbox1.Checked = False '是否提出申請
                    Checkbox1.Disabled = True
                    BudID.SelectedIndex = 0 '預算別
                    BudID.Enabled = False
                    hidBlackMsg.Value += "學號" + stud.Text + "." + drv("IDNO") + " " + drv("Name") + "已受處分" & vbCrLf '加入單名單暫存(2009/07/28 判斷黑名單)
                End If

                'END 黑名單為不補助鎖定特定選項
                'Dim hid_totLeft As HtmlInputHidden = e.Item.FindControl("hid_totLeft")  '20091221 andy 
                'hid_totLeft.Value = Convert.ToString(totalleft)

                If Convert.ToString(drv("DegreeID")) = "" Then '檢查學歷欄位
                    SupplyID.Enabled = False    '補助比例不可更改 
                    Checkbox1.Disabled = True   '不可提出申請
                    SumOfMoney.Enabled = False  '補助費用
                    BudID.Enabled = False       '預算別 不可更改 
                    e.Item.Cells(1).ForeColor = Color.Red
                    e.Item.Cells(1).ToolTip = "此學員學歷資料未提供完備!"
                End If
        End Select
    End Sub

    '儲存
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim conn As SqlConnection = DbAccess.GetConnection
        Dim sql As String = ""
        Dim dr As DataRow = Nothing
        'Dim dr1 As DataRow = Nothing
        Dim dt As DataTable = Nothing
        Dim da As SqlDataAdapter = Nothing
        'Dim i As Integer

        'Const cst_是否補助 = 6
        Try
            'Dim errStr As String = ""
            'errStr = chkStud_SubsidyCost()  '20091217 andy edit  檢查是否補助費用>剩餘可用餘額 (因javascript 已有擋，同時金額會變動，所以不可用vb來擋) 
            'If errStr <> "" Then
            '    Me.Page.RegisterStartupScript("msg", "<script> alert('" & errStr & "'); </script>")
            '    Exit Sub
            'End If

            sql = "SELECT * FROM Stud_SubsidyCost WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID='" & OCIDValue1.Value & "')"
            dt = DbAccess.GetDataTable(sql, da, conn)

            For Each item As DataGridItem In DataGrid1.Items
                'client 判斷
                'If item.Cells(cst_剩餘可用餘額).Text <> "" Then
                '    If CInt(item.Cells(cst_剩餘可用餘額).Text) < 0 Then
                '        Common.MessageBox(Me, "學員申請補助總額超過可用餘額，請確認申請資料正確~")
                '        Exit Sub
                '    End If
                'End If
                Dim Total As Integer = CInt(IIf(Trim(item.Cells(cst_總費用).Text) = "", 0, Trim(item.Cells(cst_總費用).Text)))
                Dim SumOfMoney As TextBox = item.FindControl("SumOfMoney") '補助費用
                Dim Checkbox1 As HtmlInputCheckBox = item.FindControl("Checkbox1")
                Dim RemainSub As HtmlInputHidden = item.FindControl("RemainSub")
                Dim PayMoney As HtmlInputHidden = item.FindControl("PayMoney") '個人支付費用
                Dim SupplyID As DropDownList = item.FindControl("SupplyID")
                Dim BudID As DropDownList = item.FindControl("BudID")
                Dim setid As HtmlInputHidden = item.FindControl("setid")
                Dim ocid As HtmlInputHidden = item.FindControl("ocid")
                Dim hid_socid As HtmlInputHidden = item.FindControl("socid")
                If Checkbox1.Disabled = False Then
                    If Checkbox1.Checked = True Then
                        If dt.Select("SOCID='" & DataGrid1.DataKeys(item.ItemIndex) & "'").Length = 0 Then
                            dr = dt.NewRow()
                            dt.Rows.Add(dr)
                            dr("SOCID") = DataGrid1.DataKeys(item.ItemIndex) 'hid_socid.Value 'DataGrid1.DataKeys(item.ItemIndex)
                        Else
                            dr = dt.Select("SOCID='" & DataGrid1.DataKeys(item.ItemIndex) & "'")(0)
                        End If
                        If SumOfMoney.Text <> "" Then
                            dr("SumOfMoney") = SumOfMoney.Text '此次可用補助額
                            dr("PayMoney") = Total - CInt(SumOfMoney.Text) '個人支付費用
                        Else
                            dr("SumOfMoney") = Convert.DBNull
                            dr("PayMoney") = Total '個人支付費用
                        End If
                        'If PayMoney.Value <> "" Then
                        '    dr("PayMoney") = PayMoney.Value '個人支付費用
                        'Else
                        '    dr("PayMoney") = Convert.DBNull
                        'End If
                        If SupplyID.SelectedValue <> "" Then
                            dr("SupplyID") = SupplyID.SelectedValue '補助比例
                        Else
                            dr("SupplyID") = Convert.DBNull
                        End If
                        If BudID.SelectedValue <> "" Then
                            dr("BudID") = BudID.SelectedValue '預算別
                        Else
                            dr("BudID") = Convert.DBNull
                        End If
                        '20090123 andy  edit 產投、在職 2009年 身分別為「就業保險被保險人非自願失業者」時
                        '1.預算來源設定為 02:就安基金 ； 2.補助比例為100%
                        '--------------------------  start
                        'If sm.UserInfo.TPlanID = 28 Then
                        '    If CInt(Me.sm.UserInfo.Years) > 2008 Then
                        '        sql = " select PayRate  from   Stud_EnterType2  where  "
                        '        sql += "  setid=" & Trim(setid.Value.ToString) & "  and   signUpStatus in (1,3)  and  ocid1='" & Trim(ocid.Value.ToString) & "'"
                        '        dr1 = DbAccess.GetOneRow(sql)
                        '        If Not IsDBNull(dr1("PayRate")) Then
                        '            If Convert.ToInt16(dr1("PayRate")) = 1 Then
                        '                dr("AppliedNote") = "學費須預繳100%"
                        '            End If
                        '        End If
                        '    End If
                        'End If
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now
                    Else
                        If dt.Select("SOCID='" & DataGrid1.DataKeys(item.ItemIndex) & "'").Length <> 0 Then
                            dt.Select("SOCID='" & DataGrid1.DataKeys(item.ItemIndex) & "'")(0).Delete()
                        End If
                    End If
                End If
                '20080717  Andy  不補助則刪除
                If Checkbox1.Disabled = True AndAlso Checkbox1.Checked = False Then
                    If dt.Select("SOCID='" & DataGrid1.DataKeys(item.ItemIndex) & "'").Length <> 0 Then
                        dt.Select("SOCID='" & DataGrid1.DataKeys(item.ItemIndex) & "'")(0).Delete()
                    End If
                End If
                'i += 1
            Next
            DbAccess.UpdateDataTable(dt, da)
            Common.MessageBox(Me, "儲存成功")
            'Button1_Click(sender, e)
            Call Search1() '查詢 (SQL)
        Catch ex As Exception
            Dim strErrmsg As String = ""
            strErrmsg += "/*  ex.ToString: */" & vbCrLf
            strErrmsg += ex.ToString & vbCrLf
            strErrmsg += "/* sql: */" & vbCrLf
            strErrmsg += sql & vbCrLf
            'strErrmsg += TIMS.GetErrorMsg(MyPage) '取得錯誤資訊寫入
            strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)
            Common.MessageBox(Me.Page, "發生錯誤：" & ex.ToString)
        End Try
        'End If
    End Sub

    Private Sub DataGrid1_SortCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridSortCommandEventArgs) Handles DataGrid1.SortCommand
        If Me.ViewState("sort") <> e.SortExpression Then
            Me.ViewState("sort") = e.SortExpression
        Else
            Me.ViewState("sort") = e.SortExpression & " DESC"
        End If
        'Button1_Click(Me, e)
        Call Search1() '查詢 (SQL)
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim dr As DataRow
        dr = TIMS.GET_OnlyOne_OCID(RIDValue.Value) '判斷機構是否只有一個班級
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        DataGridTable.Style("display") = "none"
        If Not dr Is Nothing Then
            If dr("total") = "1" Then '如果只有一個班級
                TMID1.Text = dr("trainname")
                OCID1.Text = dr("classname")
                TMIDValue1.Value = dr("trainid")
                OCIDValue1.Value = dr("ocid")
                'DataGridTable.Style("display") = "none"
            End If
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub
End Class