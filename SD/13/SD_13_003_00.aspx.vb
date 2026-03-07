Partial Class SD_13_003_00
    Inherits AuthBasePage

    Const cst_學號 As Integer = 0
    Const cst_姓名 As Integer = 1
    Const cst_身分證號碼 As Integer = 2
    'Const cst_是否取得結訓資格 = 3
    ''Const cst_出席達3分之2 = 4

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

    '年度小於等於2011啟用。
#Region "Functions"

    '查詢資料
    Private Function Get_ClassStudentsOfClass(ByVal ocid As Integer, ByVal orderCol As String) As DataTable
        Dim rst As DataTable = Nothing

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " select cs.socid" & vbCrLf
        sql += " ,sum(t.Hours) TOHours " & vbCrLf
        sql += " from class_classinfo cc" & vbCrLf
        sql += " join Class_StudentsOfClass cs  on cs.ocid =cc.ocid " & vbCrLf
        sql += " join Stud_Turnout t on t.socid =cs.socid " & vbCrLf
        sql += " where 1=1" & vbCrLf
        sql += " and cc.ocid =@ocid" & vbCrLf
        sql += " group by cs.socid" & vbCrLf
        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        Dim oCmd As New SqlCommand(sql, objconn)
        With oCmd
            .Parameters.Clear()
            .Parameters.Add("ocid", SqlDbType.VarChar).Value = ocid
            dt.Load(.ExecuteReader())
        End With

        Dim sqlstr As String = ""
        sqlstr = "" & vbCrLf
        sqlstr += " select a.SOCID" & vbCrLf
        sqlstr += " ,substr(a.StudentID,-2) StudentID" & vbCrLf
        sqlstr += " ,b.IDNO" & vbCrLf
        sqlstr += " ,b.Name" & vbCrLf
        sqlstr += " ,d.TotalCost" & vbCrLf
        sqlstr += " ,f.SumOfMoney" & vbCrLf
        sqlstr += " ,f.PayMoney" & vbCrLf
        sqlstr += " ,f.AppliedStatusM" & vbCrLf
        sqlstr += " ,f.AppliedNote" & vbCrLf
        sqlstr += " ,a.ActNO" & vbCrLf
        sqlstr += " ,f.BudID" & vbCrLf
        sqlstr += " ,dbo.NVL(a.CreditPoints,0) CreditPoints" & vbCrLf
        sqlstr += " ,d.THours" & vbCrLf
        sqlstr += " ,0 TOHours" & vbCrLf
        sqlstr += " ,CONVERT(varchar, c.STDate, 111)  STDate" & vbCrLf
        sqlstr += " from Class_StudentsOfClass a " & vbCrLf
        sqlstr += " join Stud_StudentInfo b on b.SID=a.SID" & vbCrLf
        sqlstr += " join Class_ClassInfo c on c.OCID=a.OCID" & vbCrLf
        sqlstr += " join Plan_PlanInfo d on d.ComIDNO=c.ComIDNO and d.PlanID=c.PlanID and d.SeqNO=c.SeqNO" & vbCrLf
        sqlstr += " join Stud_SubSidyCost f on f.SOCID=a.SOCID" & vbCrLf
        sqlstr += " where a.OCID='" & ocid & "' " & vbCrLf
        If orderCol = "" Then sqlstr += "order by StudentID " Else sqlstr += "order by " & orderCol
        rst = DbAccess.GetDataTable(sqlstr, objconn)
        Dim ff As String = ""
        For Each dr As DataRow In rst.Rows '循環RST
            ff = "SOCID=" & Convert.ToString(dr("SOCID")) '循環RST 搜尋DT
            If dt.Select(ff).Length > 0 Then
                dr("TOHours") = dt.Select(ff)(0)("TOHours") '寫入DR
            End If
        Next
        rst.AcceptChanges()
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

#Region "NO USE"
    'Private Function Get_GovAppl2(ByVal idno As String, ByVal sdate As DateTime, ByVal tmpConn As SqlConnection) As Integer
    '    Dim sqlAdp As New SqlDataAdapter
    '    Dim sqlStr As String = String.Empty
    '    Dim rst As Integer = 0

    '    Try
    '        sqlStr = "select dbo.FN_GET_GOVAPPL2(@idno,@sdate) "
    '        With sqlAdp
    '            .SelectCommand = New SqlCommand(sqlStr, tmpConn)
    '            .SelectCommand.Parameters.Clear()
    '            .SelectCommand.Parameters.Add("@idno", SqlDbType.VarChar).Value = idno
    '            .SelectCommand.Parameters.Add("@sdate", SqlDbType.DateTime).Value = sdate
    '            rst = .SelectCommand.ExecuteScalar()
    '        End With
    '    Catch ex As Exception
    '        Common.MessageBox(Me, ex.ToString)
    '        tmpConn.Close()
    '    Finally
    '        sqlAdp.Dispose()
    '    End Try
    '    Return rst
    'End Function

    'Private Sub Get_YearsPeriod(ByVal idno As String, ByRef sdate As String, ByRef edate As String, ByVal applied As String, ByVal tmpConn As SqlConnection)
    '    Dim dt As DataTable = Nothing
    '    Dim Sql As String = ""
    '    Sql = "" & vbCrLf
    '    Sql += " select convert(varchar,a.STDate,111) as SDate" & vbCrLf
    '    Sql += "   ,convert(varchar,dateadd(day,-1,dateadd(year,3,a.STDate)),111) as EDate" & vbCrLf
    '    Sql += " from Class_ClassInfo a " & vbCrLf
    '    Sql += " join Class_StudentsOfClass b on b.OCID=a.OCID" & vbCrLf
    '    Sql += " join Stud_SubSidyCost c on c.SOCID=b.SOCID " & vbCrLf
    '    If applied = "N" Then
    '        Sql += " and isnull(c.AppliedStatusM,'R')='R'" & vbCrLf
    '    Else
    '        Sql += " and c.AppliedStatusM='Y'" & vbCrLf
    '    End If
    '    Sql += " join Stud_StudentInfo d on d.SID=b.SID" & vbCrLf
    '    Sql += " where 1=1 " & vbCrLf
    '    Sql += " and d.IDNO='" & idno & "' " & vbCrLf
    '    If edate <> "" Then
    '        If IsDate(edate) Then
    '            Sql += " and a.STDate>'" & edate & "' " & vbCrLf
    '        End If
    '    End If
    '    Try
    '        dt = DbAccess.GetDataTable(Sql, tmpConn)
    '        If dt.Rows.Count > 0 Then
    '            sdate = Convert.ToString(dt.Rows(0).Item(0))
    '            edate = Convert.ToString(dt.Rows(0).Item(1))
    '        End If
    '    Catch ex As Exception
    '        Common.MessageBox(Me, ex.ToString)

    '    End Try
    'End Sub

    'Private Function Get_SubSidyCost(ByVal idno As String, ByVal stdate As String, ByVal applied As String, ByRef tmpSDate As String, ByRef tmpEDate As String, ByVal tmpConn As SqlConnection) As Integer
    '    Dim rst As Integer = 0
    '    Dim sqlStr As String = String.Empty
    '    Dim sDate As String = String.Empty
    '    Dim eDate As String = String.Empty
    '    Dim chk As Boolean = False 'True:有傳入開結訓日 'False:沒有傳入開結訓日

    '    Try
    '        If tmpSDate = "" And tmpEDate = "" Then
    '            Do While chk = False
    '                '先試著取審核通過的三年區間
    '                Get_YearsPeriod(idno, sDate, eDate, "Y", tmpConn)
    '                '沒取到就試者取申請紀錄的三年區間
    '                If sDate = String.Empty Or eDate = String.Empty Then Get_YearsPeriod(idno, sDate, eDate, "N", tmpConn)
    '                '再沒取到 就直接用目前的開訓日期的三年區間
    '                If sDate = String.Empty Or eDate = String.Empty Then
    '                    sDate = stdate
    '                    eDate = DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Year, 3, CDate(sDate))).ToString("yyyy/MM/dd")
    '                End If
    '                '判斷開訓日期是否落入取得區間，不是的話就重新取得
    '                If sDate <> String.Empty And eDate <> String.Empty Then
    '                    If CDate(stdate) >= CDate(sDate) And CDate(stdate) <= CDate(eDate) Then chk = True
    '                    If CDate(eDate) >= Today Then chk = True
    '                    If chk = False Then sDate = String.Empty
    '                End If
    '            Loop
    '        Else
    '            chk = True
    '            sDate = tmpSDate
    '            eDate = tmpEDate
    '        End If
    '        If chk Then
    '            sqlStr = "select isnull(sum(a.SumOfMoney),0) as SumOfMoney " & vbCrLf
    '            sqlStr += "from Stud_SubSidyCost a join Class_StudentsOfClass b on b.SOCID=a.SOCID " & vbCrLf
    '            sqlStr += "join Stud_StudentInfo c on c.SID=b.SID " & vbCrLf
    '            sqlStr += "join Class_ClassInfo d on d.OCID=b.OCID " & vbCrLf
    '            sqlStr += "where 1=1" & vbCrLf
    '            If applied = "N" Then
    '                sqlStr += " and isnull(a.AppliedStatusM,'R')='R' " & vbCrLf
    '            Else
    '                sqlStr += " and a.AppliedStatusM='Y' " & vbCrLf
    '            End If
    '            sqlStr += "and d.STDate between '" & sDate & "' and '" & eDate & "' " & vbCrLf
    '            sqlStr += "and IDNO='" & idno & "' " & vbCrLf
    '            rst = DbAccess.ExecuteScalar(sqlStr, tmpConn)
    '        End If
    '        tmpSDate = sDate
    '        tmpEDate = eDate
    '    Catch ex As Exception
    '        Common.MessageBox(Me, ex.ToString)

    '    End Try
    '    Return rst
    'End Function

    'Private Function Get_StudSubSidyCostTooltip(ByVal idno As String, ByVal sdate As String, ByVal edate As String, ByVal tmpConn As SqlConnection) As String
    '    Dim rst As String = ""
    '    Dim sqlStr As String = String.Empty
    '    Dim dt As DataTable = Nothing

    '    Try
    '        sqlStr = "select '20'+d.Years as Years,e.OrgName,d.ClassCName,d.CyclType" & vbCrLf
    '        sqlStr += ",a.AppliedStatusM,a.SumOfMoney,a.AppliedStatus " & vbCrLf
    '        sqlStr += "from Stud_SubSidyCost a join Class_StudentsOfClass b on b.SOCID=a.SOCID " & vbCrLf
    '        sqlStr += "join Stud_StudentInfo c on c.SID=b.SID " & vbCrLf
    '        sqlStr += "join Class_ClassInfo d on d.OCID=b.OCID " & vbCrLf
    '        sqlStr += "join Org_OrgInfo e on e.ComIDNO=d.ComIDNO " & vbCrLf
    '        sqlStr += "where c.IDNO='" & idno & "' " & vbCrLf
    '        sqlStr += " and d.STDate between '" & sdate & "' and '" & edate & "' " & vbCrLf
    '        dt = DbAccess.GetDataTable(sqlStr, tmpConn)
    '        If dt.Rows.Count > 0 Then
    '            For Each dr As DataRow In dt.Rows
    '                rst += "(" & Convert.ToString(dr("Years")) & ")" & Convert.ToString(dr("OrgName")) & "-"
    '                rst += Convert.ToString(dr("ClassCName"))
    '                rst += IIf(Convert.ToString(dr("CyclType")) = "", "", "第" & Convert.ToString(dr("CyclType")) & "期") & ": "
    '                Select Case Convert.ToString(dr("AppliedStatusM"))
    '                    Case "Y"
    '                        rst += "審核通過"
    '                    Case "N"
    '                        rst += "審核失敗"
    '                    Case "R"
    '                        rst += "退件修正"
    '                    Case Else
    '                        rst += "審核中"
    '                End Select
    '                rst += " $" & Convert.ToString(dr("SumOfMoney")) & " :"
    '                If IsDBNull(dr("AppliedStatus")) = True Then
    '                    rst += "撥款中"
    '                Else
    '                    If dr("AppliedStatus") = True Then rst += "已撥款" Else rst += "撥款失敗"
    '                End If
    '                rst += vbCrLf
    '            Next
    '        End If

    '    Catch ex As Exception
    '        Common.MessageBox(Me, ex.ToString)

    '    End Try
    '    Return rst
    'End Function
#End Region

    Private Function Get_StudBligateData(ByVal idno As String, ByVal sdate As DateTime, ByVal tmpConn As SqlConnection) As Boolean
        Dim rst As Boolean = False  '沒有投保
        Dim sqlStr As String = String.Empty
        Dim dt As DataTable
        Dim cntBli As Integer = 0
        idno = TIMS.ChangeIDNO(idno)
        Try
            '先檢查是否存在只有加保沒有退保的紀錄，有的話代表加保中。
            sqlStr = ""
            sqlStr &= "select count(1) cnt " & vbCrLf
            sqlStr += "from (select ActNo,ChangeMode,Max(MDate) as MDate from Stud_BligateData where ChangeMode=4 and IDNO='" & idno & "' group by ActNo,ChangeMode) a " & vbCrLf
            sqlStr += "left join (select ActNo,ChangeMode,Max(MDate) as MDate from Stud_BligateData where ChangeMode=2 and IDNO='" & idno & "' group by ActNo,ChangeMode) b on b.ActNo=a.ActNo and b.MDate>=a.MDate " & vbCrLf
            sqlStr += "where b.MDate is null and a.MDate<='" & sdate & "' "
            cntBli = DbAccess.ExecuteScalar(sqlStr, tmpConn)
            If cntBli > 0 Then rst = True

            '開訓前不存在只有加保紀錄的資料時，檢查開訓前最後一筆加保的退保紀錄是否在開訓後
            If Not rst Then
                sqlStr = ""
                sqlStr &= "select a.ActNo,a.ChangeMode,a.MDate,b.ActNo as ActNo2,b.ChangeMode as ChangeMode2,b.MDate as MDate2 " & vbCrLf
                sqlStr += "from (select ActNo,ChangeMode,Max(MDate) as MDate from Stud_BligateData where ChangeMode=4 and IDNO='" & idno & "' group by ActNo,ChangeMode) a "
                sqlStr += "left join (select ActNo,ChangeMode,Max(MDate) as MDate from Stud_BligateData where ChangeMode=2 and IDNO='" & idno & "' group by ActNo,ChangeMode) b on b.ActNo=a.ActNo and b.MDate>=a.MDate "
                sqlStr += "where b.ActNo is not null and a.MDate<='" & sdate & "' order by a.MDate desc "
                dt = DbAccess.GetDataTable(sqlStr, tmpConn)
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
        Catch ex As Exception
            Dim strErrmsg As String = ""
            strErrmsg += "/*  ex.ToString: */" & vbCrLf
            strErrmsg += ex.ToString & vbCrLf
            strErrmsg += "/* sqlStr: */" & vbCrLf
            strErrmsg += sqlStr & vbCrLf
            'strErrmsg += TIMS.GetErrorMsg(MyPage) '取得錯誤資訊寫入
            strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)

            Common.MessageBox(Me, ex.ToString)
        End Try
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
        sql += " SELECT cs.socid,f.AppliedStatusM" & vbCrLf
        sql += " FROM Stud_SubSidyCost f" & vbCrLf
        sql += " JOIN Class_StudentsOfClass cs on cs.SOCID=f.SOCID " & vbCrLf
        sql += " join Stud_StudentInfo ss on ss.SID=cs.SID " & vbCrLf
        sql += " where cs.OCID='" & OCIDValue & "' and f.AppliedStatusM is null" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)

        '整班經費審核通過及不補助
        '儲存、經費審核確認、審核選單開放使用
        sql = "SELECT AppliedResultM FROM Class_ClassInfo WHERE OCID='" & OCIDValue & "'"
        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)

        If IsDBNull(dr("AppliedResultM")) OrElse (dr("AppliedResultM").ToString = "R") Then
            Me.AuditCheck.Enabled = True '經費審核確認
            TIMS.Tooltip(AuditCheck, "整班經費審核，尚未確認", True)
        Else
            If dt.Rows.Count = 0 Then
                Me.AuditCheck.Enabled = False '經費審核確認
                TIMS.Tooltip(AuditCheck, "整班經費審核通過(補助及不補助)", True)
            Else
                Me.AuditCheck.Enabled = True '經費審核確認
                TIMS.Tooltip(AuditCheck, "整班經費審核通過,但尚有學員未審核通過，尚未確認", True)
            End If
        End If
    End Sub

    Sub AuditNum()
        Dim sql As String
        Dim dr As DataRow
        'Dim dt As DataTable

        sql = ""
        sql &= " SELECT Sum(Case When AppliedStatusM = 'Y' Then 1 Else 0 End) as SNum, " '審核成功筆數
        sql += " Sum(Case When AppliedStatusM = 'N' Then 1 Else 0 End) as FNum, " '審核失敗筆數
        sql += " Sum(Case When AppliedStatusM = 'R' Then 1 Else 0 End) as RNum, " '退件修正筆數
        sql += " Sum(Case When AppliedStatusM is Null Then 1 Else 0 End) as ANum " '未審核(請選擇、還原)筆數
        sql += " FROM Stud_SubsidyCost WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID='" & OCIDValue1.Value & "') AND Budid is not null "
        dr = DbAccess.GetOneRow(sql, objconn)

        Me.SNum.Text = dr("SNum").ToString
        Me.FNum.Text = dr("FNum").ToString
        Me.ANum.Text = dr("ANum").ToString
        Me.RNum.Text = dr("RNum").ToString
    End Sub

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

        If OCIDValue1.Value <> "" Then
            OpenButtons(OCIDValue1.Value)
        End If
    End Sub

    '查詢鈕
    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim sqlStr As String = ""
        sqlStr = "select AppliedResultM from Class_ClassInfo where OCID='" & OCIDValue1.Value & "' "
        Me.ViewState("appRst") = Convert.ToString(DbAccess.ExecuteScalar(sqlStr, objconn))

        Show_DataGrid(Datagrid2, Get_ClassStudentsOfClass(OCIDValue1.Value, Me.ViewState("sort")), "SOCID")

        If OCIDValue1.Value <> "" Then
            OpenButtons(OCIDValue1.Value)
        End If
    End Sub

    '審核確認鈕
    Private Sub AuditCheck_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AuditCheck.Click
        Dim sql As String
        Dim CNT As Integer = 0 '資料總數
        Dim Y, N, R, S As Integer 'Y為審核成功數,N為審核失敗數,R為退件修正數,S為請選擇數


        'Dim i As Integer 'INDEX 值
        'Dim errStr As String = ""
        'errStr = chkStud_SubsidyCost()  '20091217 andy edit  檢查是否補助費用>剩餘可用餘額  
        '    Me.Page.RegisterStartupScript("msg", "<script> alert('" & errStr & "'); </script>")
        '    Exit Sub
        'End If

        '職類/班別 OCIDValue1.Value
        If OCIDValue1.Value = "" Then
            Common.MessageBox(Me, "職類/班別 有誤，請重新選擇／查詢！")
            Exit Sub
        End If

        Dim chk_overpay As Boolean = False '是否溢撥(預設值=false 否)
        Dim Overpay_Name As String = ""
        For Each item As DataGridItem In Datagrid2.Items
            Dim AppliedStatusM As DropDownList = item.FindControl("list_Verify")
            Dim hid_OverPay As HtmlInputHidden = item.FindControl("hid_OverPay")
            Dim link_Name As LinkButton = item.FindControl("link_Name")

            If (AppliedStatusM.Enabled = True) Then         '目前尚未審過
                If AppliedStatusM.SelectedValue = "Y" Then  '選擇審核通過者
                    If IIf(hid_OverPay.Value = "", 0, CInt(hid_OverPay.Value)) < 0 Then     '剩餘可用餘額 減去 本次補助費用之後的金額是否有負數
                        chk_overpay = True
                        If Overpay_Name = "" Then
                            Overpay_Name = "學員:" & link_Name.Text & ""
                        Else
                            Overpay_Name = Overpay_Name & "\n" & "學員:" & link_Name.Text & ""
                        End If
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
        Dim objTrans As SqlTransaction = Nothing
        Dim conn As SqlConnection = DbAccess.GetConnection()
        Call TIMS.OpenDbConn(conn)
        'TIMS.TestDbConn(Me, conn, True)
        objTrans = DbAccess.BeginTrans(conn)

        sql = "SELECT * FROM Stud_SubsidyCost p WHERE exists (SELECT 'x'  FROM Class_StudentsOfClass c WHERE c.OCID='" & OCIDValue1.Value & "' AND c.SOCID=p.SOCID )"
        dt = DbAccess.GetDataTable(sql, da, objTrans)

        sql = "SELECT * FROM Class_StudentsOfClass WHERE OCID='" & OCIDValue1.Value & "'"
        dt2 = DbAccess.GetDataTable(sql, da2, objTrans)

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

        DbAccess.UpdateDataTable(dt, da, objTrans)
        DbAccess.UpdateDataTable(dt2, da2, objTrans)

        sql = ""
        sql += "Update Class_ClassInfo set "
        '若有一筆選退件修正，則該班開班基本資料的學員經費審核狀態為退件修正
        If R > 0 Then
            sql += " AppliedResultM = 'R' "
        Else
            '學員只有審核成功和失敗，只要一筆為成功，則該班開班基本資料的學員經費審核狀態為審核成功
            If Y > 0 Then
                sql += " AppliedResultM = 'Y' "
            Else
                '為N
                '學員全部為審核失敗，則該班開班基本資料的學員經費審核狀態為審核失敗
                sql += " AppliedResultM = 'N' "
            End If
        End If
        sql += " WHERE OCID= '" & OCIDValue1.Value & "'"
        DbAccess.ExecuteNonQuery(sql, objTrans)
        DbAccess.CommitTrans(objTrans)
        'Try


        'Catch ex As Exception
        '    DbAccess.RollbackTrans(objTrans)
        '    Call TIMS.CloseDbConn(conn)
        '    Common.MessageBox(Me, "儲存失敗，請重新執行，補助審核！")
        '    Exit Sub
        'End Try
        Call TIMS.CloseDbConn(conn)

        '查詢鈕
        Button1_Click(sender, e)
        'End If
    End Sub

    '單一班級查詢1
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim dr As DataRow
        '判斷機構是否只有一個班級
        dr = TIMS.GET_OnlyOne_OCID(RIDValue.Value)
        If Not dr Is Nothing Then
            If dr("total") = "1" Then '如果只有一個班級
                TMID1.Text = dr("trainname")
                OCID1.Text = dr("classname")
                TMIDValue1.Value = dr("trainid")
                OCIDValue1.Value = dr("ocid")
                DataGridTable.Style("display") = "none"
            Else '不只一個班級
                TMID1.Text = ""
                OCID1.Text = ""
                TMIDValue1.Value = ""
                OCIDValue1.Value = ""
                DataGridTable.Style("display") = "none"
            End If
        Else
            TMID1.Text = ""
            OCID1.Text = ""
            TMIDValue1.Value = ""
            OCIDValue1.Value = ""
            DataGridTable.Style("display") = "none"
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
                'Dim tmpConn As SqlConnection = DbAccess.GetConnection()
                'Dim sqlAdp As New SqlDataAdapter
                Dim sqlStr As String = String.Empty
                Me.ViewState("SOCID") = e.CommandArgument

                If Convert.ToString(Me.ViewState("SOCID")) <> "" Then
                    '學員經費審核狀態-申請 (NULL未審核)
                    sqlStr = "update Stud_SubSidyCost set AppliedStatusM=null where SOCID='" & Me.ViewState("SOCID") & "' "
                    DbAccess.ExecuteNonQuery(sqlStr, objconn)
                End If

                If OCIDValue1.Value <> "" Then
                    '學員經費審核結果 (Null未審核)
                    sqlStr = "update Class_ClassInfo set AppliedResultM=null where OCID='" & OCIDValue1.Value & "' "
                    DbAccess.ExecuteNonQuery(sqlStr, objconn)
                End If

                Common.MessageBox(Me, "還原成功")
                '重新執行查詢
                Button1_Click(Me, e)

            Case "Link"
                KeepSearch()
                TIMS.Utl_Redirect1(Me, "../13/SD_13_003_Bligate.aspx?ID=" & Request("ID") & "&" & e.CommandArgument)

        End Select
    End Sub

    Private Sub Datagrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles Datagrid2.ItemDataBound
        Dim drData As DataRowView = e.Item.DataItem
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
                Dim labStudentID As Label = e.Item.FindControl("lab_StudentID")
                Dim linkName As LinkButton = e.Item.FindControl("link_Name")
                Dim labIDNO As Label = e.Item.FindControl("lab_IDNO")
                Dim listVerify As DropDownList = e.Item.FindControl("list_Verify")
                Dim btn_BackVerify As LinkButton = e.Item.FindControl("btn_BackVerify") '還原鈕
                Dim txtVerifyNote As TextBox = e.Item.FindControl("txt_VerifyNote")
                Dim labActNO As Label = e.Item.FindControl("lab_ActNO")

                Dim listBudID As DropDownList = e.Item.FindControl("list_BudID")
                Dim labEndClass As Label = e.Item.FindControl("lab_EndClass")
                Dim labOnClassRate As Label = e.Item.FindControl("lab_OnClassRate")
                Dim labIsSubSidy As Label = e.Item.FindControl("lab_IsSubSidy") '是否補助
                Dim labTotal As Label = e.Item.FindControl("lab_Total") '總費用
                Dim labSubSidyCost As Label = e.Item.FindControl("lab_SubSidyCost") '補助費用
                Dim labPersonalCost As Label = e.Item.FindControl("lab_PersonalCost") '個人支付
                Dim labBalance As Label = e.Item.FindControl("lab_Balance") '剩餘可用餘額
                Dim labOtherGovApply As Label = e.Item.FindControl("lab_OtherGovApply") '其他申請中金額
                Dim labStar As Label = e.Item.FindControl("lab_Star")
                Dim totalSubSidy As Integer 'Dim totalSubSidy, sumSubSidy As Integer
                Dim sDate As String = ""
                Dim eDate As String = ""
                Dim strCommandArg As String = String.Empty
                Dim hid_OverPay As HtmlInputHidden = e.Item.FindControl("hid_OverPay")

                listBudID.Enabled = False '預算別鎖定
                labStudentID.Text = Convert.ToString(drData("StudentID"))
                linkName.Text = Convert.ToString(drData("Name"))
                labIDNO.Text = Convert.ToString(drData("IDNO"))
                strCommandArg += "&IDNO=" & Convert.ToString(drData("IDNO"))
                strCommandArg += "&Name=" & Convert.ToString(drData("Name"))
                strCommandArg += "&STDate=" & Common.FormatDate(Convert.ToString(drData("STDate")))
                strCommandArg += "&ActNo=" & Convert.ToString(drData("ActNO"))
                linkName.CommandArgument = strCommandArg
                '總費用、補助費用、個人支付
                labSubSidyCost.Text = drData("SumOfMoney")
                labPersonalCost.Text = drData("PayMoney")
                labTotal.Text = drData("SumOfMoney") + drData("PayMoney")
                '餘額、其他補助
                'If sm.UserInfo.Years < 2008 Then totalSubSidy = 30000 Else totalSubSidy = 50000
                '可用補助額(2007年3年3萬)
                '可用補助額(2008年3年5萬)
                '可用補助額(2012年3年7萬)
                totalSubSidy = TIMS.Get_3Y_SupplyMoney(Me)

                'labBalance.Text = totalSubSidy - Get_SubSidyCost(drData("IDNO"), drData("STDate"), "Y", sDate, eDate, objConn)
                labBalance.Text = totalSubSidy - TIMS.Get_SubsidyCost(drData("IDNO").ToString(), drData("STDate").ToString(), "", "Y", objconn)

                '尚未審核者需檢查剩餘可用餘額 減去 本次補助費用之後的金額 (是否有負數)
                If Convert.ToString(drData("AppliedStatusM")) = "" Then
                    hid_OverPay.Value = (IIf(labBalance.Text = "", 0, CInt(labBalance.Text)) - IIf(labSubSidyCost.Text = "", 0, CInt(labSubSidyCost.Text)))
                End If

                'labOtherGovApply.Text = Get_SubSidyCost(drData("IDNO"), drData("STDate"), "N", sDate, eDate, objConn)
                labOtherGovApply.Text = TIMS.Get_SubsidyCost(drData("IDNO").ToString(), drData("STDate").ToString(), drData("SOCID").ToString(), "N", objconn)

                'labOtherGovApply.Text = drData("GovAppl2") 'Get_SubSidyCost(drData("IDNO"), drData("STDate"), "N", sDate, eDate, objConn)
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
                If drData("CreditPoints") = True Then labEndClass.Text = "是" Else labEndClass.Text = "否"

                '出席達2/3
                '出席達3/4 2012 
                'e.Item.Cells(cst_出席達4分之3).Text = "否"
                'labOnClassRate.Text = "是"
                'If drData("TOHours") > 0 Then '出缺勤檔 請假時數
                '    If CDbl(drData("TOHours")) > CDbl(drData("THours") / 4) Then
                '        labOnClassRate.Text = "否"
                '    End If
                'End If
                'e.Item.Cells(cst_出席達3分之2).Text = "否"
                labOnClassRate.Text = "是"
                If drData("TOHours") > 0 Then '出缺勤檔 請假時數
                    If CDbl(drData("TOHours")) > CDbl(drData("THours") / 3) Then
                        labOnClassRate.Text = "否"
                    End If
                End If

                '是否補助
                If labOnClassRate.Text = "是" AndAlso labEndClass.Text = "是" Then labIsSubSidy.Text = "是" Else labIsSubSidy.Text = "<font color='red'>否</font>"

                '判斷開訓日是否有落在加退保期間，沒有的話打上星號。
                labStar.Visible = Get_StudBligateData(labIDNO.Text, drData("STDate"), objconn).Equals(False)
                '20091230 andy add
                Dim hid_SubSidyCost As HtmlInputHidden = e.Item.FindControl("hid_SubSidyCost")
                'Dim hid_Name As HtmlInputHidden = e.Item.FindControl("hid_Name")
                'Dim hid_Balance As HtmlInputHidden = e.Item.FindControl("hid_Balance")
                Dim hid_vstatus As HtmlInputHidden = e.Item.FindControl("hid_vstatus")
                hid_SubSidyCost.Value = labSubSidyCost.Text
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
End Class
