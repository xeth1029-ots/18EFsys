Partial Class SD_13_003_16
    Inherits AuthBasePage

    Const cst_學號 As Integer = 0
    Const cst_姓名 As Integer = 1
    Const cst_身分證號碼 As Integer = 2
    'Const cst_是否取得結訓資格 = 3
    ''Const cst_出席達3分之2 = 4
    ''Const cst_出席達4分之3 = 4

    'Const cst_是否補助 = 5
    'Const cst_總費用 = 6
    'Const cst_補助費用 = 7
    'Const cst_個人支付 = 8
    'Const cst_剩餘可用餘額 = 9
    Const cst_其他申請中金額 As Integer = 10
    'Const cst_審核狀態 = 11
    'Const cst_審核備註 = 12
    'Const cst_保險證號 = 13
    'Const cst_預算別代碼 = 14
    'Dim objgConn As SqlConnection

    '年度大於2011啟用。

#Region "Functions"

    '查詢資料 [SQL]
    Private Function Get_ClassStudentsOfClass(ByVal ocid As Integer, ByVal orderCol As String) As DataTable
        Dim rst As DataTable = Nothing
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " select a.OCID" & vbCrLf
        sql &= " ,a.SOCID" & vbCrLf
        sql &= " ,substr(a.StudentID,-2) StudentID" & vbCrLf
        sql &= " ,b.IDNO" & vbCrLf
        sql &= " ,b.Name" & vbCrLf
        sql &= " ,d.TotalCost" & vbCrLf
        sql &= " ,f.SumOfMoney" & vbCrLf
        sql &= " ,f.PayMoney" & vbCrLf
        sql &= " ,f.AppliedStatusM" & vbCrLf
        sql &= " ,f.AppliedNote" & vbCrLf
        sql &= " ,a.ActNO" & vbCrLf
        sql &= " ,f.BudID" & vbCrLf
        sql &= " ,dbo.NVL(a.CreditPoints,0) CreditPoints " & vbCrLf
        sql &= " ,d.THours" & vbCrLf
        'sql += " ,dbo.NVL(t.TOHours,0) TOHours" & vbCrLf
        'sql += " ,dbo.NVL(ff.COUNTHOURS,0)-dbo.NVL(ff.COUNTHOURS2,0) TOHours" & vbCrLf
        sql &= " ,dbo.NVL(ff.COUNTHOURS,0) TOHours" & vbCrLf
        sql &= " ,CONVERT(varchar, c.STDate, 111) as STDate" & vbCrLf
        '政府已補助經費
        sql &= " ,0 GovCost " & vbCrLf
        '其他申請中金額
        sql &= " ,0 GovAppl2" & vbCrLf
        '檢查是否存在只有加保沒有退保的紀錄，有的話代表加保中。
        sql &= " ,'N' bliFlag" & vbCrLf
        ''政府已補助經費
        'sql += " ,dbo.fn_GET_GOVCOST(b.IDNO, CONVERT(varchar, c.STDate, 111)) GovCost " & vbCrLf
        ''其他申請中金額
        'sql += " ,dbo.FN_GET_GOVAPPL2(b.IDNO,c.STDate) GovAppl2" & vbCrLf
        sql &= " from Class_StudentsOfClass a" & vbCrLf
        sql &= " join Stud_StudentInfo b on b.SID=a.SID" & vbCrLf
        sql &= " join Class_ClassInfo c on c.OCID=a.OCID" & vbCrLf
        sql &= " join Plan_PlanInfo d on d.ComIDNO=c.ComIDNO and d.PlanID=c.PlanID and d.SeqNO=c.SeqNO" & vbCrLf
        sql &= " join Stud_SubSidyCost f on f.SOCID=a.SOCID" & vbCrLf

        sql &= " LEFT JOIN ( " & vbCrLf
        'select * from STUD_TURNOUT2 WHERE rownum <=10
        '喪假(LEAVEID:05)。99:(使用者輸入)
        sql &= " SELECT t.SOCID" & vbCrLf
        sql &= " ,SUM(CASE WHEN t.LEAVEID IS NULL THEN t.Hours END) COUNTHOURS " & vbCrLf
        sql &= " FROM STUD_TURNOUT2 t" & vbCrLf
        sql &= " JOIN Class_StudentsOfClass cs on cs.socid =t.socid " & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " and cs.OCID ='" & OCIDValue1.Value & "'" & vbCrLf
        sql &= " Group By t.SOCID" & vbCrLf
        sql &= " ) ff ON ff.SOCID=a.SOCID" & vbCrLf
        sql &= " where a.OCID='" & ocid & "' " & vbCrLf

        If orderCol = "" Then sql &= " order by StudentID " Else sql &= " order by " & orderCol
        rst = DbAccess.GetDataTable(sql, objconn)

        For i As Integer = 0 To rst.Rows.Count - 1
            Dim rst2 As DataTable
            Dim dr As DataRow = rst.Rows(i)
            sql = ""
            '政府已補助經費
            sql &= " select dbo.fn_GET_GOVCOST('" & dr("IDNO") & "', '" & dr("STDate") & "') GovCost"
            '其他申請中金額
            sql &= " , dbo.FN_GET_GOVAPPL2('" & dr("IDNO") & "'," & TIMS.to_date(dr("STDate")) & ") GovAppl2"
            '其他申請中金額(並排除本班)
            'sql &= " ,dbo.FN_GET_GOVAPPL22('" & dr("IDNO") & "'," & TIMS.to_date(dr("STDate")) & "," & dr("OCID") & ") GovAppl2"
            sql &= " "
            rst2 = DbAccess.GetDataTable(sql, objconn)
            Dim dr2 As DataRow = rst2.Rows(0)
            dr("GovCost") = dr2("GovCost")
            dr("GovAppl2") = dr2("GovAppl2")

            Dim bliFlag As Boolean = False
            '檢查是否存在只有加保沒有退保的紀錄，有的話代表加保中。
            bliFlag = Get_StudBligateData(dr("IDNO"), CDate(dr("STDate")), objconn)
            If bliFlag Then
                dr("bliFlag") = "Y"
            End If
        Next

        Return rst
    End Function

    '顯示查詢後的資料LIST
    Private Sub Show_DataGrid(ByVal dg As DataGrid, ByVal dt As DataTable, Optional ByVal key As String = "", Optional ByVal page As Integer = 0)
        AuditNumPanel.Visible = False
        msg.Text = "查無資料"
        DataGridTable.Style("display") = "none"

        If Not dt Is Nothing Then
            dg.DataSource = dt
            dg.CurrentPageIndex = page
            If key <> "" Then dg.DataKeyField = key
            dg.DataBind()
            AuditNum()

            AuditNumPanel.Visible = True
            msg.Text = ""
            DataGridTable.Style("display") = "inline"
        End If
    End Sub

    '檢查是否存在只有加保沒有退保的紀錄，有的話代表加保中。
    Function Get_StudBligateData(ByVal idno As String,
                                 ByVal sdate As DateTime,
                                 ByVal tmpConn As SqlConnection) As Boolean
        Dim rst As Boolean = False  '沒有投保
        Dim sqlStr As String = String.Empty
        idno = TIMS.ChangeIDNO(idno)
        Call TIMS.OpenDbConn(tmpConn)

        '先檢查是否存在只有加保沒有退保的紀錄，有的話代表加保中。
        'sqlStr = " select count(1) cnt " & vbCrLf
        'sqlStr += " from (select ActNo,ChangeMode,Max(MDate) as MDate from Stud_BligateData where ChangeMode=4 and upper(IDNO)='" & idno & "' group by ActNo,ChangeMode) a " & vbCrLf
        'sqlStr += " left join (select ActNo,ChangeMode,Max(MDate) as MDate from Stud_BligateData where ChangeMode=2 and upper(IDNO)='" & idno & "' group by ActNo,ChangeMode) b on b.ActNo=a.ActNo and b.MDate>=a.MDate " & vbCrLf
        'sqlStr += " where b.MDate is null and a.MDate<=" & TIMS.to_date(sdate)
        'Dim cntBli As Integer = DbAccess.ExecuteScalar(sqlStr, tmpConn)

        sqlStr = ""
        sqlStr &= " select a.ActNo,a.ChangeMode,a.MDate,b.ActNo as ActNo2,b.ChangeMode as ChangeMode2,b.MDate as MDate2 " & vbCrLf
        sqlStr &= " from (select ActNo,ChangeMode,Max(MDate) MDate from Stud_BligateData where ChangeMode=4 and IDNO='" & idno & "' and MDate<= " & TIMS.to_date(sdate) & " group by ActNo,ChangeMode ) a "
        sqlStr &= " join (select ActNo,ChangeMode,Max(MDate) MDate from Stud_BligateData where ChangeMode=2 and IDNO='" & idno & "' group by ActNo,ChangeMode) b on b.ActNo=a.ActNo and b.MDate>=a.MDate "
        sqlStr &= " order by a.MDate desc "
        Dim dt As New DataTable
        Dim oCmd As New SqlCommand(sqlStr, tmpConn)
        With oCmd
            .Parameters.Clear()
            dt.Load(.ExecuteReader())
        End With
        If dt.Rows.Count = 0 Then Return False
        'If dt.Rows.Count > 0 Then rst = True
        '開訓前不存在只有加保紀錄的資料時，檢查開訓前最後一筆加保的退保紀錄是否在開訓後
        If Not rst Then
            If dt.Rows.Count > 0 Then
                For Each dr As DataRow In dt.Rows
                    If CDate(dr("MDate2")) >= sdate Then
                        rst = False
                        Exit For
                    Else
                        rst = True
                    End If
                Next
            End If
        End If
        Return rst
    End Function

    Sub KeepSearch()
        Session("_Search") = ""
        Session("_Search") += "center=" & center.Text
        Session("_Search") += "&RIDValue=" & RIDValue.Value
        Session("_Search") += "&TMID1=" & TMID1.Text
        Session("_Search") += "&OCID1=" & OCID1.Text
        Session("_Search") += "&TMIDValue1=" & TMIDValue1.Value
        Session("_Search") += "&OCIDValue1=" & OCIDValue1.Value
        If DataGridTable.Style("display") = "inline" Then
            Session("_Search") += "&Button1=TRUE"
        End If
    End Sub

    '整班經費審核通過及不補助
    '儲存、經費審核確認、審核選單開放使用
    Sub OpenButtons(ByVal OCIDValue As String)
        Dim sql As String = ""

        '查看學員資料
        sql = "" & vbCrLf
        sql &= " SELECT cs.socid,f.AppliedStatusM" & vbCrLf
        sql &= " FROM Stud_SubSidyCost f" & vbCrLf
        sql &= " JOIN Class_StudentsOfClass cs on cs.SOCID=f.SOCID " & vbCrLf
        sql &= " join Stud_StudentInfo ss on ss.SID=cs.SID " & vbCrLf
        sql &= " where cs.OCID='" & OCIDValue & "' and f.AppliedStatusM is null" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)

        '整班經費審核通過及不補助
        '儲存、經費審核確認、審核選單開放使用
        sql = "SELECT AppliedResultM FROM Class_ClassInfo WHERE OCID='" & OCIDValue & "'"
        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)

        Me.AuditCheckR.Visible = False
        If IsDBNull(dr("AppliedResultM")) OrElse (dr("AppliedResultM").ToString = "R") Then
            Me.AuditCheck.Enabled = True '經費審核確認
            TIMS.Tooltip(AuditCheck, "整班經費審核，尚未確認", True)
        Else
            If dt.Rows.Count = 0 Then
                If sm.UserInfo.RoleID = 0 Then
                    '只有系統管理者開啟此功能。
                    Me.AuditCheckR.Visible = True
                    TIMS.Tooltip(AuditCheckR, "提供還原整班經費審核通過(補助及不補助)", True)
                End If
                Me.AuditCheck.Enabled = False '經費審核確認
                TIMS.Tooltip(AuditCheck, "整班經費審核通過(補助及不補助)", True)
            Else
                Me.AuditCheck.Enabled = True '經費審核確認
                TIMS.Tooltip(AuditCheck, "整班經費審核通過,但尚有學員未審核通過，尚未確認", True)
            End If
        End If
    End Sub

    Sub AuditNum()
        'Dim sql As String
        'Dim dr As DataRow
        'Dim dt As DataTable
        'sql = "SELECT Sum(Case When AppliedStatusM = 'Y' Then 1 Else 0 End) as SNum, " '審核成功筆數
        'sql += "Sum(Case When AppliedStatusM = 'N' Then 1 Else 0 End) as FNum, " '審核失敗筆數
        'sql += "Sum(Case When AppliedStatusM = 'R' Then 1 Else 0 End) as RNum, " '退件修正筆數
        'sql += "Sum(Case When AppliedStatusM is Null Then 1 Else 0 End) as ANum " '未審核(請選擇、還原)筆數
        'sql += "FROM Stud_SubsidyCost WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID='" & OCIDValue1.Value & "') AND Budid is not null "
        Me.SNum.Text = "" 'dr("SNum").ToString
        Me.FNum.Text = "" 'dr("FNum").ToString
        Me.ANum.Text = "" 'dr("ANum").ToString
        Me.RNum.Text = "" 'dr("RNum").ToString
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT dbo.NVL(COUNT(Case When c.AppliedStatusM = 'Y' Then 1 End),0) SNum" & vbCrLf
        sql &= " ,dbo.NVL(COUNT(Case When c.AppliedStatusM = 'N' Then 1 End),0) FNum" & vbCrLf
        sql &= " ,dbo.NVL(COUNT(Case When c.AppliedStatusM = 'R' Then 1 End),0) RNum" & vbCrLf
        sql &= " ,dbo.NVL(COUNT(Case When c.AppliedStatusM is Null Then 1 End),0) ANum" & vbCrLf
        sql &= " FROM Class_StudentsOfClass cs" & vbCrLf
        sql &= " JOIN STUD_SUBSIDYCOST c on c.SOCID =cs.SOCID" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND c.Budid IS NOT NULL" & vbCrLf
        sql &= " and cs.OCID='" & OCIDValue1.Value & "'" & vbCrLf
        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)
        If dr Is Nothing Then Exit Sub
        Me.SNum.Text = dr("SNum").ToString
        Me.FNum.Text = dr("FNum").ToString
        Me.ANum.Text = dr("ANum").ToString
        Me.RNum.Text = dr("RNum").ToString
    End Sub

#End Region

#Region "NO USE"
    'Function chkStud_SubsidyCost() As String  '檢查是否  補助費用>剩餘可用餘額  
    '    Dim errMsg As String = ""
    '    'Dim i As Int16 = 0
    '    For Each itm As DataGridItem In Datagrid2.Items
    '        Dim lab_SubSidyCost As Label = itm.FindControl("lab_SubSidyCost")   '補助費用
    '        Dim lab_Balance As Label = itm.FindControl("lab_Balance")            '剩餘可用餘額
    '        Dim Supply As Integer = 0     '補助費用
    '        Dim totLeft As Integer = 0    '剩餘可用餘額
    '        'i = i + 1
    '        If Trim(lab_SubSidyCost.Text) <> "" Then
    '            Supply = CInt(lab_SubSidyCost.Text)
    '        End If
    '        If Trim(lab_Balance.Text) <> "" Then
    '            totLeft = CInt(lab_Balance.Text)
    '        End If
    '        'errMsg = errMsg + "(" & i.ToString & ") " & Supply.ToString & "/" & totLeft.ToString & "\n"
    '        If Supply > totLeft Then
    '            errMsg = "補助費用 不可大於 剩餘可用餘額！"
    '            lab_SubSidyCost.Style("background-color") = "#FFCCFF"
    '        Else
    '            lab_SubSidyCost.Style("background-color") = "#FFFFFF"
    '        End If
    '    Next
    '    'Common.MessageBox(Me.Page, "<script>alert('" & errMsg & "'); </script>")
    '    Return errMsg
    'End Function
#End Region

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

        Dim work2016 As String = TIMS.Utl_GetConfigSet("work2016")
        If work2016 <> "Y" Then
            Select Case sm.UserInfo.Years
                Case Is <= "2011"
                    Call TIMS.CloseDbConn(objconn)
                    Server.Transfer("SD_13_003_00.aspx?ID=" & Request("ID"))
                    Exit Sub
                Case Is <= "2015"
                    Call TIMS.CloseDbConn(objconn)
                    Server.Transfer("SD_13_003_15.aspx?ID=" & Request("ID"))
                    Exit Sub
            End Select
        End If

        'TIMS.TestDbConn(Me, objconn, True)
        Call TIMS.OpenDbConn(objconn)

        If Not IsPostBack Then
            msg.Text = ""
            DataGridTable.Style("display") = "none"
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            'Button3.Attributes("onclick") = "return confirm('這樣會改變審核狀態，是否確定？');"

            '帶入查詢參數
            If Not Session("_Search") Is Nothing Then
                center.Text = TIMS.GetMyValue(Session("_Search"), "center")
                RIDValue.Value = TIMS.GetMyValue(Session("_Search"), "RIDValue")
                TMID1.Text = TIMS.GetMyValue(Session("_Search"), "TMID1")
                OCID1.Text = TIMS.GetMyValue(Session("_Search"), "OCID1")
                TMIDValue1.Value = TIMS.GetMyValue(Session("_Search"), "TMIDValue1")
                OCIDValue1.Value = TIMS.GetMyValue(Session("_Search"), "OCIDValue1")
                If TIMS.GetMyValue(Session("_Search"), "Button1") = "TRUE" Then
                    Session("_Search") = Nothing
                    Button1_Click(sender, e)
                Else
                    Session("_Search") = Nothing
                End If
            End If

            '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
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
        'AuditCheck.Attributes("onclick") = "return CheckDate(); return false;" '確認重複參訓是否繼續
        AuditCheck.Attributes("onclick") = "if (confirm('經費審核確認') ==true){ return chkMoney(); } else {return false;}  "
        AuditNumPanel.Visible = False

        '暫不提供 還原經費審核確認
        AuditCheckR.Visible = False
        AuditCheckR.Attributes("onclick") = "return confirm('這樣會還原經費審核確認，是否確定？');"

        '整班經費審核通過及不補助
        If OCIDValue1.Value <> "" Then Call OpenButtons(OCIDValue1.Value)

    End Sub

    '查詢鈕
    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        If OCIDValue1.Value <> "" Then
            Dim sqlStr As String = ""
            sqlStr = " select AppliedResultM from Class_ClassInfo where OCID='" & OCIDValue1.Value & "' "
            Me.ViewState("appRst") = Convert.ToString(DbAccess.ExecuteScalar(sqlStr, objconn))
            Dim odt As DataTable = Get_ClassStudentsOfClass(OCIDValue1.Value, Me.ViewState("sort"))

            Call Show_DataGrid(Datagrid2, odt, "SOCID")

            '整班經費審核通過及不補助
            Call OpenButtons(OCIDValue1.Value)
        End If
    End Sub

    '審核確認鈕
    Private Sub AuditCheck_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AuditCheck.Click
        Dim sql As String
        Dim CNT As Integer = 0 '資料總數
        Dim Y, N, R, S As Integer 'Y為審核成功數,N為審核失敗數,R為退件修正數,S為請選擇數

        '職類/班別 OCIDValue1.Value
        If OCIDValue1.Value = "" Then
            Common.MessageBox(Me, "職類/班別 有誤，請重新選擇／查詢！")
            Exit Sub
        End If

        Dim chk_overpay As Boolean = False '是否溢撥(預設值=false 否)
        Dim Overpay_Name As String = ""
        For Each eItem As DataGridItem In Datagrid2.Items
            Dim AppliedStatusM As DropDownList = eItem.FindControl("list_Verify")
            Dim hid_OverPay As HtmlInputHidden = eItem.FindControl("hid_OverPay")
            Dim link_Name As LinkButton = eItem.FindControl("link_Name")
            Dim lab_Balance As Label = eItem.FindControl("lab_Balance") '剩餘可用餘額
            Dim lab_SubSidyCost As Label = eItem.FindControl("lab_SubSidyCost") '補助費用(本次補助費用)

            If (AppliedStatusM.Enabled = True) Then         '目前尚未審過
                If AppliedStatusM.SelectedValue = "Y" Then  '選擇審核通過者
                    If hid_OverPay.Value = "" Then
                        '未執行計算重新計算
                        hid_OverPay.Value = (IIf(lab_Balance.Text = "", 0, CInt(lab_Balance.Text)) - IIf(lab_SubSidyCost.Text = "", 0, CInt(lab_SubSidyCost.Text)))
                    End If

                    If CInt(hid_OverPay.Value) < 0 Then     '剩餘可用餘額 減去 本次補助費用之後的金額是否有負數
                        chk_overpay = True
                        If Overpay_Name <> "" Then Overpay_Name &= "\n"
                        Overpay_Name &= "學員:" & link_Name.Text & ""
                    End If
                End If
            End If
        Next

        If chk_overpay = True Then
            Me.Page.RegisterStartupScript("Errmsg", "<script>alert('" & ("本班學員,於本次補助審核後,其剩餘可用額度將會變成負數,造成溢撥狀況,煩請再確認補助金額!" & "\n" & Overpay_Name).ToString.Replace("'", "\'").Replace(Convert.ToChar(10), "\n").Replace(Convert.ToChar(13), "") & "');</script>")
            Exit Sub
        End If

        CNT = 0
        Y = 0 : N = 0 : R = 0 : S = 0
        For Each item As DataGridItem In Datagrid2.Items
            Dim AppliedStatusM As DropDownList = item.FindControl("list_Verify")
            'Select Case Convert.ToString(AppliedStatusM.SelectedIndex)
            '    Case "1"
            '        Y += 1
            '    Case "2"
            '        N += 1
            '    Case "3"
            '        R += 1
            '    Case Else '"0"
            '        S += 1
            'End Select

            Select Case Convert.ToString(AppliedStatusM.SelectedValue)
                Case "Y"
                    Y += 1
                Case "N"
                    N += 1
                Case "R"
                    R += 1
                Case Else '"0"
                    S += 1
            End Select
            CNT += 1
        Next

        AuditNumPanel.Visible = True '經費審核確認 開啟
        If CNT = 0 OrElse (Y + N + R) = 0 OrElse S > 0 Then
            If S > 0 Then
                '學員審核只要一筆為請選擇，就給警告
                Common.MessageBox(Me, "學員審核狀態有「請選擇」！")
                Exit Sub
            End If
            Common.MessageBox(Me, "學員資料數 異常！")
            Exit Sub
        End If

        'Button3_Click(sender, e)

        '儲存區如下順序有別
        Dim dr As DataRow = Nothing
        Dim dt As DataTable = Nothing
        Dim da As SqlDataAdapter = Nothing
        Dim dr2 As DataRow = Nothing
        Dim dt2 As DataTable = Nothing
        Dim da2 As SqlDataAdapter = Nothing

        Dim objTrans As SqlTransaction = DbAccess.BeginTrans(objconn)
        sql = "SELECT * FROM STUD_SUBSIDYCOST p WHERE EXISTS (SELECT 'x' FROM Class_StudentsOfClass c WHERE c.OCID='" & OCIDValue1.Value & "' AND c.SOCID=p.SOCID )"
        dt = DbAccess.GetDataTable(sql, da, objTrans)
        For Each item As DataGridItem In Datagrid2.Items
            Dim AppliedStatusM As DropDownList = item.FindControl("list_Verify")
            Dim AppliedNote As TextBox = item.FindControl("txt_VerifyNote")
            Dim BudID As DropDownList = item.FindControl("list_BudID")

            'UPDATE Stud_SubsidyCost.BudgetId
            dr = dt.Select("SOCID='" & Datagrid2.DataKeys(item.ItemIndex) & "'")(0)
            Select Case AppliedStatusM.SelectedIndex
                Case 0
                    dr("AppliedStatusM") = Convert.DBNull
                    dr("AppliedStatus") = Convert.DBNull
                Case 1
                    dr("AppliedStatusM") = "Y"
                    dr("AppliedStatus") = Convert.DBNull
                Case 2
                    dr("AppliedStatusM") = "N"
                    dr("AppliedStatus") = "0" '不撥款
                Case 3
                    dr("AppliedStatusM") = "R"
                    dr("AppliedStatus") = "0" '不撥款
            End Select
            If BudID.SelectedValue <> "" Then
                dr("BudID") = BudID.SelectedValue '預算別
            Else
                dr("BudID") = Convert.DBNull
            End If
            dr("AppliedNote") = IIf(AppliedNote.Text = "", Convert.DBNull, AppliedNote.Text)
            dr("ModifyAcct") = sm.UserInfo.UserID
            dr("ModifyDate") = Now
        Next
        DbAccess.UpdateDataTable(dt, da, objTrans)

        sql = "SELECT * FROM Class_StudentsOfClass WHERE OCID='" & OCIDValue1.Value & "'"
        dt2 = DbAccess.GetDataTable(sql, da2, objTrans)
        For Each item As DataGridItem In Datagrid2.Items
            Dim AppliedStatusM As DropDownList = item.FindControl("list_Verify")
            Dim AppliedNote As TextBox = item.FindControl("txt_VerifyNote")
            Dim BudID As DropDownList = item.FindControl("list_BudID")

            'UPDATE Class_StudentsOfClass.BudgetId
            dr2 = dt2.Select("SOCID='" & Datagrid2.DataKeys(item.ItemIndex) & "'")(0)
            If BudID.SelectedValue <> "" Then
                dr2("BudgetId") = BudID.SelectedValue '預算別
            Else
                dr2("BudgetId") = Convert.DBNull
            End If
            dr2("ModifyAcct") = sm.UserInfo.UserID
            dr2("ModifyDate") = Now
        Next
        DbAccess.UpdateDataTable(dt2, da2, objTrans)


        '若有一筆選退件修正，則該班開班基本資料的學員經費審核狀態為退件修正
        Dim ARMValue As String = ""
        If R > 0 Then
            ARMValue = "R"
            'sql += " AppliedResultM = 'R' "
        Else
            '學員只有審核成功和失敗，只要一筆為成功，則該班開班基本資料的學員經費審核狀態為審核成功
            If Y > 0 Then
                ARMValue = "Y"
                'sql += " AppliedResultM = 'Y' "
            Else
                ARMValue = "N"
                '為N
                '學員全部為審核失敗，則該班開班基本資料的學員經費審核狀態為審核失敗
                'sql += " AppliedResultM = 'N' "
            End If
        End If
        'sql += " WHERE OCID= '" & OCIDValue1.Value & "'"
        'DbAccess.ExecuteNonQuery(sql, objTrans)

        Dim oCmd As SqlCommand = Nothing
        sql = " UPDATE CLASS_CLASSINFO SET AppliedResultM=@AppliedResultM WHERE OCID=@OCID "
        Call TIMS.OpenDbConn(objconn)
        oCmd = New SqlCommand(sql, objconn, objTrans)
        With oCmd
            .Parameters.Clear()
            .Parameters.Add("AppliedResultM", SqlDbType.VarChar).Value = ARMValue
            .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCIDValue1.Value
            .ExecuteNonQuery()
        End With
        DbAccess.CommitTrans(objTrans)

        '整班經費審核通過及不補助 LOG
        'Call UPDATE_Stud_SubsidyCostLOG(OCIDValue1.Value)

        'Try
        '    'Dim objconn As SqlConnection
        '    'TIMS.TestDbConn(Me, objconn, True)
        '    'objConn = DbAccess.GetConnection()
        '    'If objConn.State = ConnectionState.Closed Then objConn.Open()

        'Catch ex As Exception
        '    DbAccess.RollbackTrans(objTrans)
        '    Common.MessageBox(Me, "儲存失敗，請重新執行，補助審核！")
        '    Exit Sub
        'End Try

        '查詢鈕
        Button1_Click(sender, e)
        'End If
    End Sub

    '還原審核確認鈕
    Protected Sub AuditCheckR_Click(sender As Object, e As EventArgs) Handles AuditCheckR.Click
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If OCIDValue1.Value <> "" Then
            '學員經費審核結果 (Null未審核)
            Dim sqlStr As String = ""
            sqlStr = "UPDATE CLASS_CLASSINFO SET AppliedResultM=null where OCID='" & OCIDValue1.Value & "' "
            DbAccess.ExecuteNonQuery(sqlStr, objconn)
        End If
    End Sub

    '單一班級查詢1
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim dr As DataRow
        '判斷機構是否只有一個班級
        dr = TIMS.GET_OnlyOne_OCID(RIDValue.Value)
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
                DataGridTable.Style("display") = "none"
            End If
        End If
    End Sub

    '單一班級查詢2
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub

    Private Sub Datagrid2_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles Datagrid2.ItemCommand
        Select Case e.CommandName
            Case "back" '還原功能
                Dim sqlStr As String = String.Empty
                Me.ViewState("SOCID") = e.CommandArgument

                If Convert.ToString(Me.ViewState("SOCID")) <> "" Then
                    '學員經費審核狀態-申請 (NULL未審核)
                    sqlStr = "UPDATE STUD_SUBSIDYCOST SET AppliedStatusM=null where SOCID='" & Me.ViewState("SOCID") & "' "
                    DbAccess.ExecuteNonQuery(sqlStr, objconn)
                End If

                If OCIDValue1.Value <> "" Then
                    '學員經費審核結果 (Null未審核)
                    sqlStr = "UPDATE CLASS_CLASSINFO SET AppliedResultM=null where OCID='" & OCIDValue1.Value & "' "
                    DbAccess.ExecuteNonQuery(sqlStr, objconn)
                End If

                Common.MessageBox(Me, "還原成功")
                '重新執行查詢
                Button1_Click(Me, e)

            Case "Link" 'linkName (link_Name)
                KeepSearch()
                TIMS.Utl_Redirect1(Me, "../13/SD_13_003_Bligate.aspx?ID=" & Request("ID") & "&" & e.CommandArgument)

        End Select
    End Sub

    Private Sub Datagrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles Datagrid2.ItemDataBound
        Const Msg_目前申請總額 As String = "學員目前已申請未核准補助金總額(超過剩餘可用餘額以紅字表示)"
        Select Case e.Item.ItemType
            Case ListItemType.Header
                Dim listVerifyAll As DropDownList = e.Item.FindControl("list_VerifyAll")
                'Me.ViewState("appRst")
                If Me.ViewState("appRst") = "Y" Or Me.ViewState("appRst") = "N" Then
                    listVerifyAll.Enabled = False
                Else
                    listVerifyAll.Enabled = True
                    listVerifyAll.Attributes.Add("onChange", "SelectAll();")
                End If

                Dim mysort As New System.Web.UI.WebControls.Image
                Select Case Me.ViewState("sort")
                    Case "StudentID"
                        mysort.ImageUrl = "../../images/SortUp.gif"
                        e.Item.Cells(cst_學號).Controls.Add(mysort)
                    Case "StudentID DESC"
                        mysort.ImageUrl = "../../images/SortDown.gif"
                        e.Item.Cells(cst_學號).Controls.Add(mysort)
                    Case "IDNO"
                        mysort.ImageUrl = "../../images/SortUp.gif"
                        e.Item.Cells(cst_身分證號碼).Controls.Add(mysort)
                    Case "IDNO DESC"
                        mysort.ImageUrl = "../../images/SortDown.gif"
                        e.Item.Cells(cst_身分證號碼).Controls.Add(mysort)
                End Select

            Case ListItemType.Item, ListItemType.AlternatingItem, ListItemType.EditItem
                Dim drData As DataRowView = e.Item.DataItem
                Dim labStudentID As Label = e.Item.FindControl("lab_StudentID")
                Dim linkName As LinkButton = e.Item.FindControl("link_Name")
                Dim labIDNO As Label = e.Item.FindControl("lab_IDNO")
                Dim listVerify As DropDownList = e.Item.FindControl("list_Verify")
                Dim btn_BackVerify As LinkButton = e.Item.FindControl("btn_BackVerify") '還原鈕
                Dim txtVerifyNote As TextBox = e.Item.FindControl("txt_VerifyNote")
                Dim labActNO As Label = e.Item.FindControl("lab_ActNO")

                Dim listBudID As DropDownList = e.Item.FindControl("list_BudID")
                Dim labEndClass As Label = e.Item.FindControl("lab_EndClass")
                Dim lab_OnClassRate As Label = e.Item.FindControl("lab_OnClassRate") '出席達3/4
                Dim lab_IsSubSidy As Label = e.Item.FindControl("lab_IsSubSidy") '是否補助
                Dim labTotal As Label = e.Item.FindControl("lab_Total") '總費用
                Dim lab_SubSidyCost As Label = e.Item.FindControl("lab_SubSidyCost") '補助費用
                Dim labPersonalCost As Label = e.Item.FindControl("lab_PersonalCost") '個人支付
                Dim lab_Balance As Label = e.Item.FindControl("lab_Balance") '剩餘可用餘額
                Dim labOtherGovApply As Label = e.Item.FindControl("lab_OtherGovApply") '其他申請中金額
                Dim lab_Star As Label = e.Item.FindControl("lab_Star")
                'Dim totalSubSidy, sumSubSidy As Integer
                'Dim sDate, eDate As String
                Dim hid_OverPay As HtmlInputHidden = e.Item.FindControl("hid_OverPay")

                listBudID.Enabled = False '預算別鎖定
                labStudentID.Text = Convert.ToString(drData("StudentID"))
                linkName.Text = Convert.ToString(drData("Name"))
                labIDNO.Text = Convert.ToString(drData("IDNO"))

                Dim strCmdArg As String = String.Empty
                strCmdArg += "&IDNO=" & Convert.ToString(drData("IDNO"))
                strCmdArg += "&Name=" & Convert.ToString(drData("Name"))
                strCmdArg += "&STDate=" & Common.FormatDate(Convert.ToString(drData("STDate")))
                strCmdArg += "&ActNo=" & Convert.ToString(drData("ActNO"))
                strCmdArg += "&SOCID=" & Convert.ToString(drData("SOCID"))
                linkName.CommandArgument = strCmdArg
                '總費用、補助費用、個人支付
                lab_SubSidyCost.Text = drData("SumOfMoney") '本次補助費用。

                labPersonalCost.Text = drData("PayMoney")
                labTotal.Text = drData("SumOfMoney") + drData("PayMoney")
                '餘額、其他補助
                'If sm.UserInfo.Years < 2008 Then totalSubSidy = 30000 Else totalSubSidy = 50000
                '可用補助額(2007年3年3萬)
                '可用補助額(2008年3年5萬)
                '可用補助額(2012年3年7萬)
                Dim totalSubSidy As Integer = TIMS.Get_3Y_SupplyMoney(Me)

                'labBalance.Text = totalSubSidy - Get_SubSidyCost(drData("IDNO"), drData("STDate"), "Y", sDate, eDate, objConn)
                '剩餘可用餘額= '可用補助額 - '政府已補助經費
                'labBalance.Text = totalSubSidy - TIMS.Get_SubsidyCost(drData("IDNO"), drData("STDate"), , , objconn)
                '剩餘可用餘額= '可用補助額 - '政府已補助經費
                lab_Balance.Text = totalSubSidy - Val(drData("GovCost"))

                '尚未審核者需檢查剩餘可用餘額 減去 本次補助費用之後的金額 (是否有負數)
                hid_OverPay.Value = ""
                If Convert.ToString(drData("AppliedStatusM")) <> "N" Then
                    '除了不通過之外，其餘金額檢查一次
                    hid_OverPay.Value = (IIf(lab_Balance.Text = "", 0, CInt(lab_Balance.Text)) - IIf(lab_SubSidyCost.Text = "", 0, CInt(lab_SubSidyCost.Text)))
                End If

                If lab_Balance.Text <> "" Then
                    If Val(lab_Balance.Text) < 0 Then
                        lab_Balance.ForeColor = Color.Red
                    End If
                End If

                'labOtherGovApply.Text = Get_SubSidyCost(drData("IDNO"), drData("STDate"), "N", sDate, eDate, objConn)
                '其他申請中金額
                'labOtherGovApply.Text = TIMS.Get_SubsidyCost(drData("IDNO"), drData("STDate"), drData("SOCID"), "N", objconn)
                'labOtherGovApply.Text = drData("GovAppl2") 'Get_SubSidyCost(drData("IDNO"), drData("STDate"), "N", sDate, eDate, objConn)
                '其他申請中金額
                labOtherGovApply.Text = Val(drData("GovAppl2"))

                Dim sDate As String = ""
                Dim eDate As String = ""
                e.Item.Cells(cst_其他申請中金額).ToolTip = Msg_目前申請總額 ' "學員目前已申請未核准補助金總額(超過剩餘可用餘額以紅字表示)"
                e.Item.Cells(cst_學號).ToolTip = TIMS.Get_StudSubSidyCostTooltip(labIDNO.Text, sDate, eDate, objconn)
                e.Item.Cells(cst_姓名).ToolTip = e.Item.Cells(cst_學號).ToolTip
                e.Item.Cells(cst_身分證號碼).ToolTip = e.Item.Cells(cst_學號).ToolTip
                '審核
                txtVerifyNote.Text = Convert.ToString(drData("AppliedNote"))
                labActNO.Text = Convert.ToString(drData("ActNO"))
                listBudID = TIMS.Get_Budget(listBudID, 2)

                btn_BackVerify.Visible = False '還原鈕
                If IsDBNull(drData("AppliedStatusM")) = False Then
                    listVerify.SelectedValue = drData("AppliedStatusM")
                    btn_BackVerify.Visible = True
                    listVerify.Enabled = False
                    txtVerifyNote.Enabled = False
                    listBudID.Enabled = False '預算別鎖定
                End If

                btn_BackVerify.CommandArgument = Convert.ToString(drData("SOCID"))

                If Not listBudID.Items.FindByValue(Convert.ToString(drData("BudID"))) Is Nothing Then listBudID.SelectedValue = drData("BudID")
                '是否結訓
                'If drData("CreditPoints") = True Then labEndClass.Text = "是" Else labEndClass.Text = "否"
                labEndClass.Text = "否"
                If Not Convert.IsDBNull(drData("CreditPoints")) AndAlso Convert.ToInt32(drData("CreditPoints")) = 1 Then
                    labEndClass.Text = "是"
                End If
                '出席達2/3
                '出席達3/4 2012 
                '缺席未超過1/5 2016
                ''''e.Item.Cells(cst_出席達4分之3).Text = "否"
                lab_OnClassRate.Text = "是"
                If drData("TOHours") > 0 Then '出缺勤檔 請假時數
                    If CDbl(drData("TOHours")) > CDbl(drData("THours") / 5) Then
                        lab_OnClassRate.Text = "否"
                    End If
                End If
                '是否補助
                If lab_OnClassRate.Text = "是" AndAlso labEndClass.Text = "是" Then lab_IsSubSidy.Text = "是" Else lab_OnClassRate.Text = "<font color='red'>否</font>"

                '判斷開訓日是否有落在加退保期間，沒有的話打上星號。
                'labStar.Visible = Get_StudBligateData(labIDNO.Text, drData("STDate"), objconn).Equals(False)
                '判斷開訓日是否有落在加退保期間，沒有的話打上星號。
                lab_Star.Visible = False
                If Convert.ToString(drData("bliFlag")) = "N" Then
                    lab_Star.Visible = True
                End If
                '20091230 andy add
                Dim hid_SubSidyCost As HtmlInputHidden = e.Item.FindControl("hid_SubSidyCost")
                'Dim hid_Name As HtmlInputHidden = e.Item.FindControl("hid_Name")
                'Dim hid_Balance As HtmlInputHidden = e.Item.FindControl("hid_Balance")
                Dim hid_vstatus As HtmlInputHidden = e.Item.FindControl("hid_vstatus")
                hid_SubSidyCost.Value = Val(lab_SubSidyCost.Text) '轉為數字

                'hid_Balance.Value = labBalance.Text
                'hid_Name.Value = Convert.ToString(drData("Name"))
                If listVerify.Enabled = True Then
                    hid_vstatus.Value = "1"
                Else
                    hid_vstatus.Value = "2"
                End If

        End Select

    End Sub

    Private Sub Datagrid2_SortCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridSortCommandEventArgs) Handles Datagrid2.SortCommand
        If Me.ViewState("sort") <> e.SortExpression Then
            Me.ViewState("sort") = e.SortExpression
        Else
            Me.ViewState("sort") = e.SortExpression & " DESC"
        End If

        '重新執行查詢
        Button1_Click(Me, e)
    End Sub

    '整班經費審核通過及不補助 LOG
    Sub UPDATE_Stud_SubsidyCostLOG(ByVal OCIDValue As String)
        If OCIDValue = "" Then Exit Sub

        Dim oCmd As SqlCommand = Nothing
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT * FROM V_STUDENTINFO WHERE OCID =@OCID " & vbCrLf
        Call TIMS.OpenDbConn(objconn)
        Dim dt1 As New DataTable
        oCmd = New SqlCommand(sql, objconn)
        With oCmd
            .Parameters.Clear()
            .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCIDValue
            dt1.Load(.ExecuteReader())
        End With

        sql = "" & vbCrLf
        sql &= " SELECT f.SOCID " & vbCrLf
        sql &= " ,f.SUMOFMONEY" & vbCrLf
        sql &= " ,f.PAYMONEY" & vbCrLf
        sql &= " ,f.APPLIEDSTATUS" & vbCrLf
        sql &= " ,f.APPLIEDNOTE" & vbCrLf
        sql &= " ,f.SUPPLYID" & vbCrLf
        sql &= " ,f.BUDID" & vbCrLf
        sql &= " ,f.MODIFYACCT" & vbCrLf
        sql &= " ,f.MODIFYDATE" & vbCrLf
        sql &= " ,f.APPLIEDSTATUSM" & vbCrLf
        sql &= " ,f.ALLOTDATE" & vbCrLf
        sql &= " ,cs.OCID" & vbCrLf
        sql &= " ,ss.IDNO" & vbCrLf
        sql &= " FROM Stud_SubSidyCost f" & vbCrLf
        sql &= " JOIN Class_StudentsOfClass cs on cs.SOCID=f.SOCID " & vbCrLf
        sql &= " JOIN Stud_StudentInfo ss on ss.SID=cs.SID " & vbCrLf
        sql &= " where cs.OCID=@OCID" & vbCrLf
        Call TIMS.OpenDbConn(objconn)
        Dim dt2 As New DataTable
        oCmd = New SqlCommand(sql, objconn)
        With oCmd
            .Parameters.Clear()
            .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCIDValue
            dt2.Load(.ExecuteReader())
        End With
    End Sub

End Class