Partial Class SD_17_001
    Inherits AuthBasePage

#Region "Functions"
    '檢查該班是否已做結訓動作
    'Private Function Check_IsClosed(ByVal OCIDValue As String) As Boolean
    '    Dim rst As Boolean = False
    '    Dim sqlStr As String = ""
    '    sqlStr = "select dbo.NVL(IsClosed,'N') IsClosed from Class_ClassInfo where OCID='" & OCIDValue & "' "
    '    Dim IsClosed As String = DbAccess.ExecuteScalar(sqlStr, objconn)
    '    If IsClosed = "Y" Then
    '        rst = True
    '    End If
    '    Return rst
    'End Function

    ''取得黑名單之身分證記錄
    'Function GET_Blacklist(ByVal dt As DataTable) As String
    '    'START 黑名單之身分證記錄 2009/07/28 by waiming
    '    Dim txtIDNO As String = ""
    '    Dim sql_Blacklist As String = ""
    '    Dim dt_Blacklist As DataTable
    '    sql_Blacklist = "select IDNO" & vbCrLf
    '    sql_Blacklist += "from Stud_Blacklist" & vbCrLf
    '    sql_Blacklist += "where Avail='Y'" & vbCrLf
    '    sql_Blacklist += "and getdate() between SBSdate and DATEADD(month, SBYears*12, SBSdate)" & vbCrLf
    '    dt_Blacklist = DbAccess.GetDataTable(sql_Blacklist, objconn)
    '    For i As Int16 = 0 To dt_Blacklist.Rows.Count - 1
    '        If dt.Select("IDNO='" & Convert.ToString(dt_Blacklist.Rows(i)("IDNO")) & "'").Length > 0 Then
    '            If txtIDNO = "" Then
    '                txtIDNO += Convert.ToString(dt_Blacklist.Rows(i)("IDNO"))
    '            Else
    '                txtIDNO += Convert.ToString(dt_Blacklist.Rows(i)("IDNO")) + ","
    '            End If
    '        End If
    '    Next
    '    'Me.ViewState("BlackIDNO") = txtIDNO
    '    'END 黑名單之身分證記錄
    '    Return txtIDNO
    'End Function

#End Region

    Const cst_學號 As Integer = 0
    Const cst_姓名 As Integer = 1
    Const cst_身分證號碼 As Integer = 2
    'Const cst_是否獲得學分 As Integer = 3
    Const cst_出席達80X As Integer = 4 '出席達80%
    Const cst_是否為在職者 As Integer = 5

    Const cst_是否補助 As Integer = 6
    Const cst_總費用 As Integer = 7
    Const cst_補助費用 As Integer = 8
    Const cst_個人支付 As Integer = 9
    Const cst_剩餘可用餘額 As Integer = 10
    Const cst_其他申請中金額 As Integer = 11
    'Const cst_是否提出申請 As Integer = 12
    Const cst_申請狀態 As Integer = 13
    Const cst_撥款狀態 As Integer = 14
    'Dim tipMsg As String = ""
    Dim gsBlackIDNO As String = "" '全域被處分學員IDNO

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
        '檢查Session TPlanID是否設定正確-46.47.58
        TIMS.CheckSessionTPlanID4647(Me, objconn)

        If Not IsPostBack Then
            DataGridTable.Style("display") = "none"
            msg.Text = "" '"查無資料"
            Dclass.Value = ""

            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID

            TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
        End If

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

    '查詢鈕
    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        ''20090907 針對尚未完成班級結訓動作的班級，是不能查出資料的。
        'If Not Check_IsClosed(OCIDValue1.Value) Then
        '    DataGridTable.Style("display") = "none"
        '    msg.Text = "查無資料(尚未完成班級結訓動作)"
        '    Exit Sub
        'End If
        hidBlackMsg.Value = "" '清空黑名單暫存記錄(2009/07/28 判斷黑名單)

        Dim dt As DataTable
        Dim dr As DataRow

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " SELECT a.OCID,a.THours" & vbCrLf
        sql += " ,a.STDate,a.AppliedResultM" & vbCrLf
        sql += " ,case when Len(d.StudentID) = 12 then dbo.SUBSTR(d.StudentID,-3) else '0'+substr(d.StudentID,-2) end as StudentID" & vbCrLf
        sql += " ,d.SETID,d.SOCID, d.StudStatus,d.IdentityID,d.MIdentityID" & vbCrLf
        sql += " ,d.AppliedResult, d.WorkSuppIdent" & vbCrLf
        sql += " ,e.Name, e.IDNO ,e.DegreeID" & vbCrLf
        sql += " ,d.CreditPoints" & vbCrLf
        sql += " ,case when b.TotalCost >= dbo.NVL(c.Total2,0) then TRUNC(dbo.NVL(b.TotalCost,0)/dbo.NVL(b.TNum,1))" & vbCrLf
        sql += " else TRUNC(dbo.NVL(c.Total2,0)/dbo.NVL(b.TNum,1)) end Total" & vbCrLf
        sql += " ,b.TotalCost,c.Total2" & vbCrLf
        sql += " ,ar.DistID, ip.TPlanID" & vbCrLf
        sql += " ,dbo.NVL(f.CountHours,0) as CountHours" & vbCrLf
        '20080805 andy 當學員輔助金撥款檔有資料時則預算別顯示申請變更後的資料
        sql += " ,case when  g.BudID is null then d.BudgetID else g.BudID end  BudgetID" & vbCrLf
        sql += " ,g.SOCID Exist,g.SumOfMoney,g.PayMoney,g.AppliedStatus,g.AppliedNote, g.SupplyID" & vbCrLf
        sql += " ,g.AppliedStatusM" & vbCrLf
        sql += " ,dbo.FN_GET_GOVAPPL2(e.IDNO,a.STDate) GovAppl2" & vbCrLf
        sql += " FROM Class_ClassInfo a" & vbCrLf
        sql += " JOIN Plan_PlanInfo b ON a.PlanID=b.PlanID and a.ComIDNO=b.ComIDNO and a.SeqNo=b.SeqNo" & vbCrLf
        sql += " JOIN Auth_Relship ar ON a.RID = ar.RID" & vbCrLf
        sql += " LEFT JOIN (" & vbCrLf
        sql += "    SELECT PlanID,ComIDNO,SeqNo" & vbCrLf
        sql += " 	,Sum(dbo.NVL(OPrice,1)*dbo.NVL(Itemage,1)) as Total2" & vbCrLf
        sql += " 	FROM Plan_CostItem p" & vbCrLf
        sql += " 	WHERE COSTMODE =3 AND exists (" & vbCrLf
        sql += " 	    select 'x' from Class_ClassInfo c" & vbCrLf
        sql += " 		where 1=1" & vbCrLf
        sql += " 		and c.ocid ='" & OCIDValue1.Value & "'" & vbCrLf
        sql += " 		AND p.PlanID=c.PlanID and p.ComIDNO=c.ComIDNO and p.SeqNo=c.SeqNo" & vbCrLf
        sql += " 	)" & vbCrLf
        sql += " 	Group By PlanID,ComIDNO,SeqNo" & vbCrLf
        sql += " ) c ON a.PlanID=c.PlanID and a.ComIDNO=c.ComIDNO and a.SeqNo=c.SeqNo" & vbCrLf
        sql += " JOIN Class_StudentsOfClass d ON a.OCID=d.OCID" & vbCrLf
        sql += " JOIN Stud_StudentInfo e ON d.SID=e.SID" & vbCrLf
        sql += " JOIN id_plan ip on ip.PlanID =b.PlanID" & vbCrLf
        sql += " LEFT JOIN (SELECT SOCID,Sum(Hours) CountHours FROM Stud_Turnout2 Group By SOCID) f ON d.SOCID=f.SOCID" & vbCrLf
        sql += " LEFT JOIN Stud_SubsidyCost g ON d.SOCID=g.SOCID" & vbCrLf
        sql += " WHERE 1=1" & vbCrLf

        sql += " AND ip.TPlanID in (" & TIMS.Cst_TPlanID46wSql & ")"
        sql += " AND a.OCID='" & OCIDValue1.Value & "'" & vbCrLf
        Select Case rblWork.SelectedValue
            Case "Y"
                sql += " AND d.WorkSuppIdent='Y' " & vbCrLf
            Case "N"
                sql += " AND (d.WorkSuppIdent='N' OR d.WorkSuppIdent IS NULL) " & vbCrLf
        End Select

        If InStr(Me.ViewState("sort"), "IDNO") > 0 Then
            sql += "order by e." & Me.ViewState("sort").ToString
        ElseIf InStr(Me.ViewState("sort"), "StudentID") > 0 Then
            sql += "order by CONVERT(numeric, case when Len(d.StudentID) = 12 then dbo.SUBSTR(d.StudentID,-3) else '0'+substr(d.StudentID,-2) end) " & Replace(Me.ViewState("sort").ToString, "StudentID", "") & vbCrLf
        Else
            sql += "order by CONVERT(numeric, case when Len(d.StudentID) = 12 then dbo.SUBSTR(d.StudentID,-3) else '0'+substr(d.StudentID,-2) end) " & vbCrLf
        End If
        dt = DbAccess.GetDataTable(sql, objconn)

        DataGridTable.Style("display") = "none"
        msg.Text = "查無資料"

        If dt.Rows.Count > 0 Then
            Dim stdBLACK2TPLANID As String = ""
            Dim iStdBlackType As Integer = TIMS.Chk_StdBlackType(Me, objconn, stdBLACK2TPLANID)
            gsBlackIDNO = TIMS.Get_StdBlackIDNO(Me, iStdBlackType, stdBLACK2TPLANID, objconn) '學員處分

            'START 黑名單之身分證記錄 2009/07/28 by waiming
            'Me.ViewState("BlackIDNO") = GET_Blacklist(dt)
            'END 黑名單之身分證記錄

            DataGridTable.Style("display") = "inline"
            msg.Text = ""

            If ViewState("sort") = "" Then
                ViewState("sort") = "StudentID"
            End If
            DataGrid1.DataKeyField = "SOCID"
            DataGrid1.DataSource = dt
            DataGrid1.DataBind()

            dr = dt.Rows(0)
            If dr("AppliedResultM").ToString = "Y" Then
                '學員經費審核結果 (通過後不可再儲存)
                Button3.Enabled = False
                TIMS.Tooltip(Button3, "學員經費審核，已經通過不可再儲存", True)
            Else
                Button3.Enabled = True
                TIMS.Tooltip(Button3, "學員經費審核，尚未通過", True)
            End If

            dt = TIMS.GET_Duplicate_Student(OCIDValue1.Value, 1, objconn) '檢查是否有重複參訓學員排除產學訓計畫
            If dt IsNot Nothing Then '有重複參訓學員
                Dclass.Value = 1 '有重複參訓學員
            Else
                Dclass.Value = 2 '沒重複
            End If
        End If
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        'Const Cst_TP46_max01 = 6000 '46:補助辦理保母職業訓練，一般身分者最高補助額度
        'Const Cst_TP46_max02 = 8000 '46:補助辦理保母職業訓練，特殊身分者最高補助額度
        'Const Cst_TP47_max01 = 5000 '47:補助辦理照顧服務員職業訓練，一般身分者最高補助額度
        'Const Cst_TP47_max02 = 8000 '47:補助辦理照顧服務員職業訓練，特殊身分者最高補助額度
        '2012 調高為80%，但改為不限制最高金額上限 即最高輔助金為可用輔助金 by AMU 豪哥 20120428

        Const cst_總費用_msg As String = "非學分班的訓練費用項目"
        Const cst_是否補助_msg As String = "是否有請領補助津貼的資格"
        Const cst_補助費用_msg As String = "預定要補助的金額(可自行變動，未申請前系統會根據可用餘額推算)"
        Const cst_個人支付_msg As String = "學員自行要支付的金額(會根據補助費用所輸入的值來調動)"
        'Const cst_剩餘可用餘額_msg AS String = "學員目前可用餘額-這次預定補助費用的剩餘金額(成為負數時會以紅字表示)"
        Const cst_剩餘可用餘額_msg As String = "學員目前可用餘額-已審核通過費用的剩餘金額(成為負數時會以紅字表示)"
        Const cst_目前申請總額_msg As String = "學員目前已申請未審核補助金總額(超過剩餘可用餘額以紅字表示)" '其他申請中金額

        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = "SD_TD1"
                e.Item.Cells(cst_是否補助).ToolTip = cst_是否補助_msg
                e.Item.Cells(cst_總費用).ToolTip = cst_總費用_msg
                e.Item.Cells(cst_補助費用).ToolTip = cst_補助費用_msg
                e.Item.Cells(cst_個人支付).ToolTip = cst_個人支付_msg
                e.Item.Cells(cst_剩餘可用餘額).ToolTip = cst_剩餘可用餘額_msg
                e.Item.Cells(cst_其他申請中金額).ToolTip = cst_目前申請總額_msg '其他申請中的金額
                If Me.ViewState("sort") <> "" Then
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

                Dim Flag As Integer = 0  '獲得結訓資格 '得到學分  
                Dim Flag2 As Integer = 0 '出席達80%
                Dim Flag3 As Integer = 0 '在職者
                Dim drv As DataRowView = e.Item.DataItem

                'Dim SupplyID As DropDownList = e.Item.FindControl("SupplyID")
                Dim BudID As DropDownList = e.Item.FindControl("BudID")
                Dim DataGrid2 As DataGrid = e.Item.FindControl("DataGrid2")
                Dim CreditPoints As Label = e.Item.FindControl("CreditPoints") '是否得到學分
                Dim SumOfMoney As TextBox = e.Item.FindControl("SumOfMoney") '申請金額
                Dim Checkbox1 As HtmlInputCheckBox = e.Item.FindControl("Checkbox1")
                Dim RemainSub As HtmlInputHidden = e.Item.FindControl("RemainSub") '決議後可用補助額
                Dim MaxSub As HtmlInputHidden = e.Item.FindControl("MaxSub") '此次最大可用補助額
                Dim PayMoney As HtmlInputHidden = e.Item.FindControl("PayMoney") '課程費用-'可用補助額=個人支付費用

                '20090201 andy edit
                '-------------
                Dim setid As HtmlInputHidden = e.Item.FindControl("setid")
                Dim ocid As HtmlInputHidden = e.Item.FindControl("ocid")
                ocid.Value = drv("ocid").ToString
                setid.Value = drv("setid").ToString
                '-------------

                Dim star1 As TextBox = e.Item.FindControl("star1") '未填寫調查表(該計畫不填寫調查表)
                Dim stud1 As TextBox = e.Item.FindControl("stud1") '學員號
                star1.Visible = True '未填寫調查表(該計畫不填寫調查表)
                star1.Text = "" '未填寫調查表(該計畫不填寫調查表)
                star1.Style("display") = "none" '未填寫調查表(該計畫不填寫調查表)
                stud1.Visible = True

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

                '受訓學員意見調查表 'SD_11_004.aspx
                '判斷是否填寫問卷(表示為該學員未填寫調查表)'dbo.fn_GET_GovCnt 'Stud_QuestionFac
                'If drv("GovCnt") = 0 Then star1.Text = "*" Else star1.Text = ""
                stud1.Text = drv("StudentID").ToString '學員號

                BudID = TIMS.Get_Budget(BudID, 2)
                If drv("DistID").ToString <> "001" Then BudID.Items.Remove(BudID.Items.FindByValue("01"))
                '規則改為show class_studentsofclass.budgetid
                If drv("BudgetID").ToString <> "" Then
                    Common.SetListItem(BudID, drv("BudgetID").ToString)
                End If

                '獲得結訓資格
                If IsDBNull(drv("CreditPoints")) Then '是否獲得結訓資格
                    CreditPoints.Text = "<font color='RED'>否</font>"
                    TIMS.Tooltip(CreditPoints, "未獲得結訓資格(為空)")
                Else
                    If CBool(drv("CreditPoints")) Then  '是否獲得結訓資格
                        CreditPoints.Text = "是"
                        Flag = 1
                    Else
                        CreditPoints.Text = "<font color='RED'>否</font>"
                        TIMS.Tooltip(CreditPoints, "未獲得結訓資格(為否)")
                    End If
                End If

                '出席達80%
                e.Item.Cells(cst_出席達80X).Text = "否"
                If drv("THours") > 0 Then
                    If (drv("THours") - drv("CountHours")) / drv("THours") >= 8 / 10 Then
                        e.Item.Cells(cst_出席達80X).Text = "是"
                        TIMS.Tooltip(e.Item.Cells(cst_出席達80X), "", True)
                        Flag2 = 1
                    End If
                End If
                If Flag2 = 0 Then
                    TIMS.Tooltip(CreditPoints, "出席未達80%")
                End If

                e.Item.Cells(cst_是否為在職者).Text = "否"
                If drv("WorkSuppIdent").ToString = "Y" Then
                    e.Item.Cells(cst_是否為在職者).Text = "是"
                    TIMS.Tooltip(e.Item.Cells(cst_是否為在職者), "", True)
                    Flag3 = 1
                End If
                If Flag3 = 0 Then
                    TIMS.Tooltip(e.Item.Cells(cst_是否為在職者), "不是在職者，不可申請")
                End If

                'Dim sql As String
                'Dim dr As DataRow
                'Dim dt As DataTable
                '20080609  Andy  可用補助額
                '含職前webservice
                Dim SubsidyCost As Double = TIMS.Get_SubsidyCost(drv("IDNO").ToString(), drv("STDate").ToString(), "", "Y", objconn)

                '產投 政府補助經費
                Dim Total As Integer = TIMS.Get_3Y_SupplyMoney()
                Total -= SubsidyCost
                If Total < 0 Then
                    Total = 0
                End If
                RemainSub.Value = Total '決議後可用補助額'50000 (3年5萬)

                Dim strTooltip As String = ""
                strTooltip = ""
                '1:獲得結訓資格 2:出席達80% 3:是否為在職者
                If Not Flag = 1 Then
                    strTooltip += "1:尚未獲得結訓資格" & vbCrLf
                End If
                '1:獲得結訓資格 2:出席達80% 3:是否為在職者
                If Not Flag2 = 1 Then
                    strTooltip += "2:出席未達80%" & vbCrLf
                End If
                '1:獲得結訓資格 2:出席達80% 3:是否為在職者
                If Not Flag3 = 1 Then
                    strTooltip += "3:補助身分不是在職者" & vbCrLf
                End If

                If IsDBNull(drv("Exist")) Then      '表示沒資料,以新增的型態顯示
                    If (Flag = 1 AndAlso Flag2 = 1 AndAlso Flag3 = 1) Then  ' 1:得到學分 2:出席達80% 3:是否為在職者
                        e.Item.Cells(cst_是否補助).Text = "是"

                        SumOfMoney.Text = Total '此次可用補助額=可用補助額 (依 個人補助額)
                        'SumOfMoney.Text = CInt(Total * 0.8) '此次可用補助額=可用補助額 (依 個人補助額)
                        'Select Case drv("TPlanID").ToString
                        '    Case "46"
                        '        Select Case drv("MIdentityID").ToString 'select top 10 * from key_Identity
                        '            Case "01" '01:一般身分者
                        '                If Total >= Cst_TP46_max01 Then SumOfMoney.Text = Cst_TP46_max01
                        '            Case Else '特殊身分者
                        '                If Total >= Cst_TP46_max02 Then SumOfMoney.Text = Cst_TP46_max02
                        '        End Select
                        '    Case "47"
                        '        Select Case drv("MIdentityID").ToString 'select top 10 * from key_Identity
                        '            Case "01" '01:一般身分者
                        '                If Total >= Cst_TP47_max01 Then SumOfMoney.Text = Cst_TP47_max01
                        '            Case Else '特殊身分者
                        '                If Total >= Cst_TP47_max02 Then SumOfMoney.Text = Cst_TP47_max02
                        '        End Select
                        'End Select
                        If drv("Total") <= SumOfMoney.Text Then  '班級補助額若 小於 個人補助額
                            SumOfMoney.Text = drv("Total")  '此次可用補助額=可用補助額 (依班級補助額)
                        End If
                        '------   End
                        MaxSub.Value = SumOfMoney.Text '此次最大可用補助額
                        e.Item.Cells(cst_個人支付).Text = CStr(CInt(drv("Total")) - CInt(IIf(Trim(SumOfMoney.Text) = "", "0", Trim(SumOfMoney.Text))))   '課程費用-'可用補助額=個人支付費用
                        PayMoney.Value = CInt(drv("Total")) - CInt(SumOfMoney.Text) '課程費用-'可用補助額=個人支付費用
                        e.Item.Cells(cst_剩餘可用餘額).Text = Total.ToString()  '可用補助額=剩餘可用餘額
                    Else
                        SumOfMoney.Enabled = False '不可填入補助費用
                        Checkbox1.Disabled = True  '不可提出申請
                        Checkbox1.Checked = False  '不可提出申請

                        e.Item.Cells(cst_剩餘可用餘額).Text = Total
                        e.Item.Cells(cst_個人支付).Text = e.Item.Cells(cst_總費用).Text
                        SumOfMoney.Text = "0" '此次可用補助額=可用補助額 
                        MaxSub.Value = SumOfMoney.Text '此次最大可用補助額
                    End If
                Else
                    If (Flag = 1 AndAlso Flag2 = 1 AndAlso Flag3 = 1) Then
                        e.Item.Cells(cst_是否補助).Text = "是"
                    Else
                        '20080606 andy 是否補助為「否」,個人支付=總費用
                        e.Item.Cells(cst_個人支付).Text = e.Item.Cells(cst_總費用).Text
                        SumOfMoney.Text = "0"
                    End If

                    SumOfMoney.Text = drv("SumOfMoney").ToString '可用補助額
                    PayMoney.Value = drv("PayMoney").ToString '個人支付費用
                    e.Item.Cells(cst_個人支付).Text = drv("PayMoney").ToString '個人支付費用

                    Checkbox1.Checked = True '有提出申請
                    BudID.Enabled = False '暫設不可更改 預算別，根據審核狀況來開放 預算別

                    If IsDBNull(drv("AppliedStatusM")) Then
                        Checkbox1.Disabled = False '可更改提出申請
                        e.Item.Cells(cst_申請狀態).Text = "審核中"
                        e.Item.Cells(cst_撥款狀態).Text = "未撥款"
                        BudID.Enabled = True '可更改 預算別
                    Else
                        Select Case drv("AppliedStatusM").ToString
                            Case "Y"
                                Checkbox1.Disabled = True
                                SumOfMoney.ReadOnly = True
                                e.Item.Cells(cst_申請狀態).Text = "審核通過"
                                If IsDBNull(drv("AppliedStatus")) Then '撥款審核狀態
                                    'Checkbox1.Disabled = True 'False '審核通過後不可放棄申請
                                    e.Item.Cells(cst_撥款狀態).Text = "待撥款" '"撥款審核中"
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
                                BudID.Enabled = True '可更改 預算別
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

                        If drv("GovAppl2") > Total Then
                            e.Item.Cells(cst_其他申請中金額).Text = "<font color=Red>" & drv("GovAppl2").ToString & "</font>"
                        End If
                    End If
                End If

                'SupplyID.Enabled = SumOfMoney.Enabled
                'BudID.Enabled = SumOfMoney.Enabled
                '補助比例和預算別改唯讀
                '學員補助不能提出申請
                If (Not IsDBNull(drv("AppliedResult")) And drv("AppliedResult").ToString = "N") Then
                    '970513 Andy ,若審核結果為不補助 
                    Checkbox1.Disabled = True
                    '970717 Andy ,若審核結果為不補助則是否提出申請應該是無法勾選的,且為不勾選的
                    Checkbox1.Checked = False

                    If (drv("AppliedResult").ToString = "N") Then
                        e.Item.Cells(cst_是否補助).ToolTip += vbCrLf & "(學員資料審核)預算別：不補助"
                    End If
                    SumOfMoney.Enabled = False '不可填入補助費用
                    BudID.Enabled = False '預算別
                    e.Item.Cells(cst_是否補助).Text = "<font color='RED'>否</font>"
                    TIMS.Tooltip(e.Item.Cells(cst_是否補助), "審核結果為不補助")
                    '20080606 andy 是否補助為「否」,個人支付=總費用
                    e.Item.Cells(cst_個人支付).Text = e.Item.Cells(cst_總費用).Text
                    SumOfMoney.Text = "0"
                End If

                '970513 Andy 學分為 0 或 出勤未滿2/3
                If Flag = 0 OrElse Flag2 = 0 OrElse Flag3 = 0 Then
                    Checkbox1.Disabled = True
                    Checkbox1.Checked = False

                    e.Item.Cells(cst_是否補助).Text = "<font color='RED'>否</font>"
                    SumOfMoney.Enabled = False '不可填入補助費用
                    BudID.Enabled = False '預算別

                    '20080606 andy 是否補助為「否」,個人支付=總費用
                    e.Item.Cells(cst_個人支付).Text = e.Item.Cells(cst_總費用).Text
                    SumOfMoney.Text = "0"
                End If

                SumOfMoney.Attributes("onchange") = "ChangeMoney(" & e.Item.ItemIndex + 1 & ",'" & SumOfMoney.ClientID & "','" & RemainSub.ClientID & "','" & MaxSub.ClientID & "','" & PayMoney.ClientID & "');"
                SumOfMoney.Attributes("onblur") = "ChangeMoney(" & e.Item.ItemIndex + 1 & ",'" & SumOfMoney.ClientID & "','" & RemainSub.ClientID & "','" & MaxSub.ClientID & "','" & PayMoney.ClientID & "');"
                SumOfMoney.Attributes("onFocus") = "ChangeMoney(" & e.Item.ItemIndex + 1 & ",'" & SumOfMoney.ClientID & "','" & RemainSub.ClientID & "','" & MaxSub.ClientID & "','" & PayMoney.ClientID & "');"

                'START 黑名單為不補助鎖定特定選項 2009/07/28 by waiming
                If gsBlackIDNO <> "" AndAlso gsBlackIDNO.IndexOf(Convert.ToString(drv("IDNO"))) > -1 Then
                    e.Item.Cells(cst_是否補助).Text = "<font color='RED'>否</font>" '是否補助

                    'SupplyID.SelectedValue = "9" '補助比例
                    'SupplyID.Enabled = False

                    SumOfMoney.Text = "0" '補助費用
                    SumOfMoney.Enabled = False

                    Checkbox1.Checked = False '是否提出申請
                    Checkbox1.Disabled = True

                    BudID.SelectedIndex = 0 '預算別
                    BudID.Enabled = False

                    hidBlackMsg.Value += "學號" + stud1.Text + "." + drv("IDNO") + " " + drv("Name") + "已受處分" & vbCrLf '加入單名單暫存(2009/07/28 判斷黑名單)
                End If
                'END 黑名單為不補助鎖定特定選項

                If strTooltip <> "" And Not Checkbox1.Checked Then
                    TIMS.Tooltip(SumOfMoney, strTooltip, True)
                    TIMS.Tooltip(BudID, strTooltip, True)
                    TIMS.Tooltip(e.Item.Cells(cst_是否補助), strTooltip)
                    TIMS.Tooltip(Checkbox1, strTooltip, True)
                End If
        End Select

    End Sub

    Private Sub DataGrid1_SortCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridSortCommandEventArgs) Handles DataGrid1.SortCommand
        If Me.ViewState("sort") <> e.SortExpression Then
            Me.ViewState("sort") = e.SortExpression
        Else
            Me.ViewState("sort") = e.SortExpression & " DESC"
        End If
        Button1_Click(Me, e)
    End Sub

    '儲存鈕
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim conn As SqlConnection = DbAccess.GetConnection
        Dim sql As String = ""
        Dim dr As DataRow = Nothing
        Dim dt As DataTable = Nothing
        Dim da As SqlDataAdapter = Nothing
        Dim i As Integer = 0

        'Const cst_是否補助 As Integer = 6
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
                'Dim SupplyID As DropDownList = item.FindControl("SupplyID")
                Dim BudID As DropDownList = item.FindControl("BudID")
                Dim setid As HtmlInputHidden = item.FindControl("setid")
                Dim ocid As HtmlInputHidden = item.FindControl("ocid")

                If Checkbox1.Disabled = False Then
                    If Checkbox1.Checked = True Then
                        If dt.Select("SOCID='" & DataGrid1.DataKeys(i) & "'").Length = 0 Then
                            dr = dt.NewRow()
                            dt.Rows.Add(dr)
                            dr("SOCID") = DataGrid1.DataKeys(i)
                        Else
                            dr = dt.Select("SOCID='" & DataGrid1.DataKeys(i) & "'")(0)
                        End If
                        If SumOfMoney.Text <> "" Then
                            dr("SumOfMoney") = SumOfMoney.Text '此次可用補助額
                            dr("PayMoney") = Total - CInt(SumOfMoney.Text) '個人支付費用
                        Else
                            dr("SumOfMoney") = Convert.DBNull
                            dr("PayMoney") = Total '個人支付費用
                        End If
                        dr("SupplyID") = Convert.DBNull '補助比例
                        'dr("BudID") = Convert.DBNull

                        ''If PayMoney.Value <> "" Then
                        ''    dr("PayMoney") = PayMoney.Value '個人支付費用
                        ''Else
                        ''    dr("PayMoney") = Convert.DBNull
                        ''End If
                        'If SupplyID.SelectedValue <> "" Then
                        '    dr("SupplyID") = SupplyID.SelectedValue '補助比例
                        'Else
                        '    dr("SupplyID") = Convert.DBNull
                        'End If
                        If BudID.SelectedValue <> "" Then
                            dr("BudID") = BudID.SelectedValue '預算別
                        Else
                            dr("BudID") = Convert.DBNull
                        End If

                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now
                    Else
                        If dt.Select("SOCID='" & DataGrid1.DataKeys(i) & "'").Length <> 0 Then
                            dt.Select("SOCID='" & DataGrid1.DataKeys(i) & "'")(0).Delete()
                        End If
                    End If
                End If
                '20080717  Andy  不補助則刪除
                If Checkbox1.Disabled = True And Checkbox1.Checked = False Then
                    If dt.Select("SOCID='" & DataGrid1.DataKeys(i) & "'").Length <> 0 Then
                        dt.Select("SOCID='" & DataGrid1.DataKeys(i) & "'")(0).Delete()
                    End If
                End If

                i += 1
            Next
            DbAccess.UpdateDataTable(dt, da)

            Common.MessageBox(Me, "儲存成功")
            Button1_Click(sender, e)
        Catch ex As Exception
            Common.MessageBox(Me.Page, "發生錯誤：" & ex.ToString)
        End Try
        'End If
    End Sub

    '判斷機構是否只有一個班級
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '判斷機構是否只有一個班級
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn)
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        DataGridTable.Style("display") = "none"
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        '如果只有一個班級
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
        DataGridTable.Style("display") = "none"

    End Sub

    '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub

End Class