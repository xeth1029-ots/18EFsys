Partial Class SD_03_006
    Inherits System.Web.UI.Page

    Dim FunDr As DataRow
    Dim MySqlStr As String
    Dim MySOCID As String
    Dim objControl As CheckBox
    'Dim DataPage As SD_03_006_add

    Dim Key_Degree As DataTable
    Dim Key_GradState As DataTable
    Dim Key_Military As DataTable
    Dim Key_Identity As DataTable
    Dim Key_Subsidy As DataTable
    Dim Key_HandicatType As DataTable
    Dim Key_HandicatLevel As DataTable
    Dim Key_JoblessWeek As DataTable
    Dim Plan_Budget As DataTable
    Dim dtArc As DataTable
    Dim PageControler1 As New PageControler
    Dim objconn As OracleConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在---------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在---------------------------End
        '分頁設定---------------Start
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定---------------End
        PageControler1 = Me.FindControl("PageControler1")
        dtArc = TIMS.Get_Auth_REndClass(Me, objconn)
        '檢查是否有使用權-----------------------------Start
        If dtArc Is Nothing Then
            Common.RespWrite(Me, "<script>alert('抱歉!!您無權限使用本程式');</script>")
            Common.RespWrite(Me, "<script>location.href='../../main.aspx';</script>")
            Exit Sub
        End If
        Dim ss As String = "IsEndDate='N'" '尚有授權設定
        If dtArc.Select(ss).Length = 0 Then
            Common.RespWrite(Me, "<script>alert('抱歉!!您無權限使用本程式');</script>")
            Common.RespWrite(Me, "<script>location.href='../../main.aspx';</script>")
            Exit Sub
        End If
        'Dim dr As DataRow
        'Dim sql As String
        'sql = ""
        'sql &= " SELECT count(1) as cnt FROM Auth_REndClass WHERE Account = '" & sm.UserInfo.UserID & "'"
        'sql += " And UseAble ='Y'"
        'dr = DbAccess.GetOneRow(sql, objconn)
        'If dr("cnt") = 0 Then
        '    Common.RespWrite(Me, "<script>alert('抱歉!!您無權限使用本程式');</script>")
        '    Common.RespWrite(Me, "<script>location.href='../../main.aspx';</script>")
        'End If
        '檢查是否有使用權-----------------------------End

        msg.Text = ""
        Page.RegisterStartupScript("ZipCcript1", TIMS.Get_ZipNameJScript(objconn))

        If Not IsPostBack Then
            ImportTable.Style.Item("display") = "none"
            DataGridTable.Style.Item("display") = "none"

            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
        End If

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1", True, "Button1")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

        Button1.Attributes("onclick") = "javascript@return search()"
        Button4.Attributes("onclick") = "CheckPrint();return false;"

        Dim Plankind As String = TIMS.Get_PlanKind(Me, objconn)
        'sql = "SELECT PlanKind FROM ID_Plan WHERE PlanID='" & sm.UserInfo.PlanID & "'"
        'dr = DbAccess.GetOneRow(sql, objconn)
        If Plankind = "1" Then
            Button5.Attributes("onclick") = "choose_class(2);"
        Else
            Button5.Attributes("onclick") = "choose_class(1);"
        End If

        Button6.Attributes("onclick") = "if(document.getElementById('DataGridTable').style.display=='none'){alert('無學員資料可以匯出!');return false;}"
        Button7.Attributes("onclick") = "if(document.form1.File1.value==''){alert('請選擇匯入檔案的路徑');return false;}"
        If sm.UserInfo.RID = "A" Or sm.UserInfo.RoleID <= 1 Then
            Button8.Attributes("onclick") = "javascript@openOrg('../../Common/LevOrg.aspx');"
        Else
            Button8.Attributes("onclick") = "javascript@openOrg('../../Common/LevOrg1.aspx');"
        End If

        '檢查帳號的功能權限-----------------------------------Start
        If sm.UserInfo.FunDt Is Nothing Then
            Common.RespWrite(Me, "<script>alert('Session過期');</script>")
            Common.RespWrite(Me, "<script>top.location.href='../../logout.aspx';</script>")
        Else
            Dim FunDt As DataTable = sm.UserInfo.FunDt
            Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & Request("ID") & "'")

            If FunDrArray.Length = 0 Then
                Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
                Common.RespWrite(Me, "<script>location.href='../../main.aspx';</script>")
            Else
                FunDr = FunDrArray(0)
                If FunDr("Sech") = 1 Then
                    Button1.Enabled = True
                    Button4.Enabled = True
                Else
                    Button1.Enabled = False
                    Button4.Enabled = False
                End If
            End If
        End If
        '檢查帳號的功能權限-----------------------------------End

        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            HyperLink1.NavigateUrl = "../../Doc/ClassStudentForTS.zip"
            Button4.Visible = False
        Else
            HyperLink1.NavigateUrl = "../../Doc/ClassStudent.zip"
        End If

        If Not IsPostBack Then
            If Not Session("_SearchStr") Is Nothing Then
                center.Text = TIMS.GetMyValue(Session("_SearchStr"), "center")
                RIDValue.Value = TIMS.GetMyValue(Session("_SearchStr"), "RIDValue")
                TMID1.Text = TIMS.GetMyValue(Session("_SearchStr"), "TMID1")
                TMIDValue1.Value = TIMS.GetMyValue(Session("_SearchStr"), "TMIDValue1")
                OCID1.Text = TIMS.GetMyValue(Session("_SearchStr"), "OCID1")
                OCIDValue1.Value = TIMS.GetMyValue(Session("_SearchStr"), "OCIDValue1")

                PageControler1.PageIndex = 0
                'PageControler1.PageIndex = TIMS.GetMyValue(Session("SearchStr"), "PageIndex")
                Dim MyValue As String = TIMS.GetMyValue(Session("_SearchStr"), "PageIndex")
                If MyValue <> "" AndAlso IsNumeric(MyValue) Then
                    MyValue = CInt(MyValue)
                    PageControler1.PageIndex = MyValue
                End If

                If TIMS.GetMyValue(Session("_SearchStr"), "submit") = "1" Then
                    Button1_Click(sender, e)
                End If

                Session("_SearchStr") = Nothing
            End If
        End If
    End Sub

    Sub Show_ClassInfo()
        Dim dr As DataRow
        Dim sql As String
        sql = "SELECT * FROM Class_ClassInfo WHERE OCID='" & OCIDValue1.Value & "'"
        dr = DbAccess.GetOneRow(sql, objconn)
        If Not dr Is Nothing Then
            DateRound.Text = FormatDateTime(dr("STDate"), 2) & "~" & FormatDateTime(dr("FTDate"), 2)
            If dr("CTName").ToString = "" Then
                CTName.Text = "無"
            Else
                If IsNumeric(dr("CTName")) Then
                    sql = "SELECT TeachCName FROM Teach_TeacherInfo WHERE CONVERT(char, TechID)='" & dr("CTName") & "'"
                    CTName.Text = DbAccess.ExecuteScalar(sql, objconn).ToString
                End If
                If CTName.Text = "" Then
                    CTName.Text = dr("CTName")
                End If
            End If
            Tnum.Text = dr("TNum")
        End If
    End Sub

    Sub search1()
        '060727 加入判斷是否有必填的欄位未填 by nick
        Dim sql As String
        Dim dt As DataTable

        sql = "" & vbCrLf
        sql += " SELECT a.SOCID,a.StudentID,b.SID,b.Name,b.IDNO,b.Sex,b.Birthday,a.StudStatus," & vbCrLf
        sql += " a.SubsidyID,a.levelNo,b.EngName,b.DegreeID,c.school,c.Department,b.MilitaryID,c.ServiceID," & vbCrLf
        sql += " c.MilitaryRank,c.ServiceOrg,c.ServicePhone,c.SServiceDate,c.FServiceDate,c.PhoneD,c.PhoneN," & vbCrLf
        sql += " c.address,a.MidentityID,a.IdentityID,c.EmergencyContact,c.EmergencyRelation,c.EmergencyAddress," & vbCrLf
        sql += " c.ZipCode3,c.ShowDetail,a.budgetID,b.IsAgree" & vbCrLf
        sql += " FROM Class_StudentsOfClass a" & vbCrLf
        sql += " JOIN Stud_StudentInfo b ON a.SID=b.SID and a.OCID='" & OCIDValue1.Value & "'" & vbCrLf
        sql += " JOIN Stud_SubData c on a.SID = c.SID" & vbCrLf
        sql += " JOIN Auth_REndClass d on a.OCID = d.OCID" & vbCrLf
        sql += " WHERE 1=1" & vbCrLf
        sql += " and d.Account = '" & sm.UserInfo.UserID & "'" & vbCrLf
        sql += " and d.UseAble = 'Y'" & vbCrLf
        Me.ViewState("SD03002_SearchSqlStr") = sql
        dt = DbAccess.GetDataTable(sql, objconn)

        msg.Text = "查無學生資料!"
        DataGridTable.Style.Item("display") = "none"
        StdNum.Text = dt.Rows.Count
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            DataGridTable.Style.Item("display") = "inline"

            PageControler1.PageDataTable = dt
            PageControler1.PrimaryKey = "SOCID"
            PageControler1.Sort = "StudentID"
            PageControler1.ControlerLoad()
        End If

    End Sub

    '查詢按鈕
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ImportTable.Style.Item("display") = "inline"

        Button2.ToolTip = "按此按鈕可以進行新增學員的動作"
        Button7.ToolTip = "按此按鈕可以進行匯入學員的動作"

        If sm.UserInfo.TPlanID <> 15 Then
            If FunDr("Adds") = 1 Then
                Button2.Enabled = True
                Button7.Enabled = True
            Else
                Button2.Enabled = False
                Button7.Enabled = False
            End If
        Else
            If sm.UserInfo.RoleID = 1 Then
                Button2.Enabled = True
                Button7.Enabled = True
            Else
                Button2.Enabled = False
                Button7.Enabled = False
            End If
        End If

        Call search1()
    End Sub

    '匯出學員資料
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim sql As String
        Dim dt As DataTable
        Dim dr As DataRow

        'copy一份sample資料----------------------------------------------------------------Start
        'Dim MyFile As System.IO.File
        'Dim MyDownload As System.IO.File
        Dim MyPath As String = ""
        Dim sFileName As String = ""
        sFileName = ""
        sFileName += "~\SD\03\Temp\"
        'sFileName += TIMS.ChangeIDNO(Replace(Replace(Replace(OCID1.Text, ")", ""), "(", ""), "/", ""))
        sFileName += TIMS.GetDateNo()
        sFileName += ".xls"
        MyPath = Server.MapPath(sFileName)

        'Dim MyFileName As String = OCID1.Text & ".xls"
        Dim MyFileName As String = TIMS.ChangeIDNO(Replace(Replace(Replace(OCID1.Text, ")", ""), "(", ""), "/", "")) & ".xls"

        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            If IO.File.Exists(Server.MapPath("~\SD\03\Sample2.xls")) Then
                IO.File.Copy(Server.MapPath("~\SD\03\Sample2.xls"), MyPath, True)
            Else
                Common.MessageBox(Me, "Sample檔案不存在")
                Exit Sub
            End If
        Else
            If IO.File.Exists(Server.MapPath("~\SD\03\Sample.xls")) Then
                IO.File.Copy(Server.MapPath("~\SD\03\Sample.xls"), MyPath, True)
            Else
                Common.MessageBox(Me, "Sample檔案不存在")
                Exit Sub
            End If
        End If

        '除去sample檔的唯讀屬性
        IO.File.SetAttributes(Server.MapPath("~\SD\03\Temp\" & Replace(Replace(Replace(OCID1.Text, ")", ""), "(", ""), "/", "") & ".xls"), IO.FileAttributes.Normal)
        'copy一份sample資料----------------------------------------------------------------End


        '根據路徑建立資料庫連線，並取出學員資料填入----------------------------------------------------------Start
        sql = "" & vbCrLf
        sql += " SELECT b.SOCID,b.StudentID,c.Name,c.EngName,c.IDNO,c.Sex,case c.Sex when 'M' then '男' when 'F' then '女' end as SexName," & vbCrLf
        sql += " c.PassPortNO,case c.PassPortNO when 1 then '本國' when 2 then '外籍' end as PassPortName,c.ChinaOrNot,case c.ChinaOrNot when 1 then '是' else '否' end as ChinaOrNotName,c.Nationality,c.PPNO" & vbCrLf
        sql += " ,case dbo.NVL(CONVERT(char, c.PPNO),'0') when '1' then '護照號碼' when '2' then '居留證號' end as PPNOName" & vbCrLf
        sql += " ,c.Birthday,c.DegreeID,e.DegreeName" & vbCrLf
        sql += " ,c.MaritalStatus,case CONVERT(char, c.MaritalStatus) when '1' then '已婚' when '2' then '未婚' end as MaritalStatusName," & vbCrLf
        sql += " d.School,d.Department,f.GradID,f.GradName,g.MilitaryID,g.MilitaryName,d.ServiceID,MilitaryAppointment,d.MilitaryRank," & vbCrLf
        sql += " d.ServiceOrg,d.ChiefRankName,d.ServicePhone,d.ServiceAddress,d.SServiceDate,d.FServiceDate,d.ZipCode4," & vbCrLf
        sql += " q.CTName4,p.ZipName4,d.ServiceAddress as Address4,d.PhoneD,d.PhoneN,d.CellPhone," & vbCrLf
        sql += " d.ZipCode1,i.CTName1,h.ZipName1,d.Address as Address1,d.ZipCode2,k.CTName2,j.ZipName2,d.HouseholdAddress as Address2," & vbCrLf
        sql += " d.Email,b.IdentityID,b.MIdentityID,l.SubsidyID,l.SubsidyName,b.OpenDate,b.CloseDate,b.EnterDate," & vbCrLf
        sql += " d.HandTypeID,r.HandTypeName,d.HandLevelID,s.HandLevelName," & vbCrLf
        sql += " d.EmergencyContact,d.EmergencyRelation,d.EmergencyPhone,d.ZipCode3,o.CTName3,n.ZipName3,d.EmergencyAddress as Address3," & vbCrLf
        sql += " b.RejectTDate1,b.RejectTDate2,m.RTReasonID,m.Reason," & vbCrLf
        sql += " d.PriorWorkOrg1,d.Title1,d.SOfficeYM1,d.FOfficeYM1,d.PriorWorkOrg2,d.Title2,d.SOfficeYM2,d.FOfficeYM2," & vbCrLf
        sql += " d.PriorWorkPay,c.RealJobless,c.JoblessID,t.JoblessName," & vbCrLf
        sql += " d.Traffic,case d.Traffic when 1 then '住宿' when 2 then '通勤' end as TrafficName," & vbCrLf
        sql += " d.ShowDetail,b.LevelNo,b.EnterChannel" & vbCrLf
        sql += " ,case CONVERT(char, b.EnterChannel) when '1' then '網路' when '2' then '現場' when '3' then '通訊' when '4' then '推介' end AS EnterChannelName," & vbCrLf
        sql += " b.TRNDMode,case CONVERT(char, b.TRNDMode) when '1' then '職訓券' when '2' then '學習券' when '3' then '推介券' end as TRNDModeName," & vbCrLf
        sql += " b.TRNDType,case CONVERT(char, b.TRNDType) when '1' then '甲式' when '2' then '乙式' end as TRNDTypeName" & vbCrLf
        sql += " ,b.BudgetID,u.BudName,c.IsAgree,b.PMode," & vbCrLf
        sql += " d.ForeName,d.ForeTitle,d.ForeSex,case d.ForeSex when 'M' then '男' when 'F' then '女' end as ForeSexName,d.ForeBirth,d.ForeIDNO,d.ForeZip,v.ZipName as ForeZipName,d.ForeAddr,kn.KNID,kn.Name as knName" & vbCrLf
        sql += " FROM Class_ClassInfo a" & vbCrLf
        sql += " JOIN Class_StudentsOfClass b ON a.OCID=b.OCID" & vbCrLf
        sql += " and a.OCID='" & OCIDValue1.Value & "'" & vbCrLf
        sql += " and b.OCID='" & OCIDValue1.Value & "'" & vbCrLf
        sql += " JOIN Stud_StudentInfo c ON b.SID=c.SID" & vbCrLf
        sql += " JOIN Stud_SubData d ON c.SID=d.SID" & vbCrLf
        sql += " LEFT JOIN (SELECT DegreeID,Name as DegreeName FROM Key_Degree) e ON c.DegreeID=e.DegreeID" & vbCrLf
        sql += " LEFT JOIN (SELECT GradID,Name as GradName FROM Key_GradState) f ON c.GraduateStatus=f.GradID" & vbCrLf
        sql += " LEFT JOIN (SELECT MilitaryID,Name as MilitaryName FROM Key_Military) g ON c.MilitaryID=g.MilitaryID" & vbCrLf
        sql += " LEFT JOIN (SELECT ZipCode,ZipName as ZipName1,CTID FROM ID_ZIP) h ON h.ZipCode=d.ZipCode1" & vbCrLf
        sql += " LEFT JOIN (SELECT CTName as CTName1,CTID FROM ID_City) i ON h.CTID=i.CTID" & vbCrLf
        sql += " LEFT JOIN (SELECT ZipCode,ZipName as ZipName2,CTID FROM ID_ZIP) j ON j.ZipCode=d.ZipCode2" & vbCrLf
        sql += " LEFT JOIN (SELECT CTName as CTName2,CTID FROM ID_City) k ON j.CTID=k.CTID" & vbCrLf
        sql += " LEFT JOIN (SELECT SubsidyID,Name as SubsidyName FROM Key_Subsidy) l ON b.SubsidyID=l.SubsidyID" & vbCrLf
        sql += " LEFT JOIN (SELECT RTReasonID,Reason FROM Key_RejectTReason) m ON b.RTReasonID=m.RTReasonID" & vbCrLf
        sql += " LEFT JOIN (SELECT ZipCode,ZipName as ZipName3,CTID FROM ID_ZIP) n ON n.ZipCode=d.ZipCode3" & vbCrLf
        sql += " LEFT JOIN (SELECT CTName as CTName3,CTID FROM ID_City) o ON n.CTID=o.CTID" & vbCrLf
        sql += " LEFT JOIN (SELECT ZipCode,ZipName as ZipName4,CTID FROM ID_ZIP) p ON p.ZipCode=d.ZipCode4" & vbCrLf
        sql += " LEFT JOIN (SELECT CTName as CTName4,CTID FROM ID_City) q ON p.CTID=q.CTID" & vbCrLf
        sql += " LEFT JOIN (SELECT HandTypeID,Name as HandTypeName FROM Key_HandicatType) r ON d.HandTypeID=r.HandTypeID" & vbCrLf
        sql += " LEFT JOIN (SELECT HandLevelID,Name as HandLevelName FROM Key_HandicatLevel) s ON d.HandLevelID=s.HandLevelID" & vbCrLf
        sql += " LEFT JOIN (SELECT JoblessID,Name as JoblessName FROM Key_JoblessWeek) t ON c.JoblessID=t.JoblessID" & vbCrLf
        sql += " LEFT JOIN (SELECT BudID,BudName FROM Key_Budget) u ON b.BudgetID=u.BudID" & vbCrLf
        sql += " LEFT JOIN view_ZipName v ON d.ForeZip=v.ZipCode" & vbCrLf
        sql += " LEFT JOIN Key_Native kn ON kn.KNID = b.Native" & vbCrLf

        dt = DbAccess.GetDataTable(sql, objconn)

        sql = "SELECT * FROM Key_Identity"
        Dim dt1 As DataTable = DbAccess.GetDataTable(sql, objconn)

        'Dim i As Integer
        If dt.Rows.Count <> 0 Then
            Dim dtSample As New DataTable
            'Dim drSample As DataRow

            Dim conn As New OleDb.OleDbConnection
            Dim cmd As OleDb.OleDbCommand

            conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & MyPath & ";Extended Properties=Excel 8.0;"
            conn.Open()

            dt.DefaultView.Sort = "StudentID"

            Dim StudentID As String = ""
            Dim Name As String = ""
            Dim LastName As String = ""
            Dim FirstName As String = ""
            Dim IDNO As String = ""
            Dim SEX As String = ""
            Dim PassPortNO As String = ""
            Dim ChinaOrNot As String = ""
            Dim Nationality As String = ""
            Dim PPNO As String = ""
            Dim Birthday As String = ""
            Dim MaritalStatus As String = ""
            Dim DegreeID As String = ""
            Dim School As String = ""
            Dim Department As String = ""
            Dim GradID As String = ""
            Dim MilitaryID As String = ""
            Dim ServiceID As String = ""
            Dim MilitaryAppointment As String = ""
            Dim MilitaryRank As String = ""
            Dim ServiceOrg As String = ""
            Dim ChiefRankName As String = ""
            Dim ServicePhone As String = ""
            Dim SServiceDate As String = ""
            Dim FServiceDate As String = ""
            Dim ZipCode4 As String = ""
            Dim Address4 As String = ""
            Dim PhoneD As String = ""
            Dim PhoneN As String = ""
            Dim CellPhone As String = ""
            Dim ZipCode1 As String = ""
            Dim Address1 As String = ""
            Dim ZipCode2 As String = ""
            Dim Address2 As String = ""
            Dim Email As String = ""
            Dim IdentityID As String = ""
            Dim MIdentityID As String = ""
            Dim SubsidyID As String = ""
            Dim OpenDate As String = ""
            Dim CloseDate As String = ""
            Dim EnterDate As String = ""
            Dim HandTypeID As String = ""
            Dim HandLevelID As String = ""
            Dim EmergencyContact As String = ""
            Dim EmergencyRelation As String = ""
            Dim EmergencyPhone As String = ""
            Dim ZipCode3 As String = ""
            Dim Address3 As String = ""
            Dim PriorWorkOrg1 As String = ""
            Dim Title1 As String = ""
            Dim SOfficeYM1 As String = ""
            Dim FOfficeYM1 As String = ""
            Dim PriorWorkOrg2 As String = ""
            Dim Title2 As String = ""
            Dim SOfficeYM2 As String = ""
            Dim FOfficeYM2 As String = ""
            Dim PriorWorkPay As String = ""
            Dim Traffic As String = ""
            Dim RealJobless As String = ""
            Dim JoblessID As String = ""
            Dim ShowDetail As String = ""
            Dim LevelNo As String = ""
            Dim EnterChannel As String = ""
            Dim TRNDMode As String = ""
            Dim TRNDType As String = ""
            Dim BudgetID As String = ""
            Dim IsAgree As String = ""
            Dim PMode As String = ""
            Dim ForeName As String = ""
            Dim ForeTitle As String = ""
            Dim ForeSex As String = ""
            Dim ForeBirth As String = ""
            Dim ForeIDNO As String = ""
            Dim ForeZip As String = ""
            Dim ForeAddr As String = ""
            Dim KNID As String = ""

            Dim AcctMode As String = ""
            Dim PostNo As String = ""
            Dim AcctHeadNo As String = ""
            Dim AcctExNo As String = ""
            Dim AcctNo As String = ""
            Dim BankName As String = ""
            Dim ExBankName As String = ""
            Dim FirDate As String = ""
            Dim Uname As String = ""
            Dim Intaxno As String = ""
            Dim Tel As String = ""
            Dim Fax As String = ""
            Dim Zip As String = ""
            Dim Addr As String = ""
            Dim ServDept As String = ""
            Dim JobTitle As String = ""
            Dim SDate As String = ""
            Dim SJDate As String = ""
            Dim SPDate As String = ""
            Dim Q1 As String = ""
            Dim Q2 As String = ""
            Dim Q3 As String = ""
            Dim Q3_Other As String = ""
            Dim Q4 As String = ""
            Dim Q5 As String = ""
            Dim Q61 As String = ""
            Dim Q62 As String = ""
            Dim Q63 As String = ""
            Dim Q64 As String = ""

            For Each dr In dt.Rows
                StudentID = Right(dr("StudentID").ToString, 2)
                Name = dr("Name").ToString()
                If dr("EngName").ToString <> "" Then
                    If dr("EngName").ToString.IndexOf(" ") <> -1 Then
                        LastName = Trim(Left(dr("EngName").ToString, dr("EngName").ToString.IndexOf(" ")))
                        FirstName = Trim(Right(dr("EngName").ToString, dr("EngName").ToString.Length - dr("EngName").ToString.IndexOf(" ") - 1))
                    Else
                        LastName = dr("EngName").ToString
                    End If
                Else
                    LastName = ""
                    FirstName = ""
                End If
                IDNO = dr("IDNO").ToString()
                SEX = dr("SEX").ToString()
                If shiftsort.SelectedValue = "1" Then
                    SEX = dr("SEX").ToString()
                Else
                    SEX = dr("SEXName").ToString()
                End If
                PassPortNO = dr("PassPortNO").ToString
                If shiftsort.SelectedValue = "1" Then
                    PassPortNO = dr("PassPortNO").ToString
                    ChinaOrNot = dr("ChinaOrNot").ToString
                    Nationality = dr("Nationality").ToString
                    PPNO = dr("PPNO").ToString
                Else
                    PassPortNO = dr("PassPortName").ToString
                    ChinaOrNot = dr("ChinaOrNotName").ToString
                    Nationality = dr("Nationality").ToString
                    PPNO = dr("PPNOName").ToString
                End If

                If dr("Birthday").ToString <> "" Then
                    Birthday = FormatDateTime(dr("Birthday").ToString, DateFormat.ShortDate)
                Else
                    Birthday = ""
                End If

                If shiftsort.SelectedValue = "1" Then
                    MaritalStatus = dr("MaritalStatus").ToString
                Else
                    MaritalStatus = dr("MaritalStatusName").ToString
                End If
                If shiftsort.SelectedValue = "1" Then
                    DegreeID = dr("DegreeID").ToString()
                Else
                    DegreeID = dr("DegreeName").ToString()
                End If

                School = dr("School").ToString()

                '11~20
                Department = dr("Department").ToString()
                GradID = dr("GradID").ToString()
                If shiftsort.SelectedValue = "1" Then
                    GradID = dr("GradID").ToString()
                Else
                    GradID = dr("GradName").ToString()
                End If
                MilitaryID = dr("MilitaryID").ToString()
                If shiftsort.SelectedValue = "1" Then
                    MilitaryID = dr("MilitaryID").ToString()
                Else
                    MilitaryID = dr("MilitaryName").ToString()
                End If
                ServiceID = dr("ServiceID").ToString
                MilitaryAppointment = dr("MilitaryAppointment").ToString
                MilitaryRank = dr("MilitaryRank").ToString
                ServiceOrg = dr("ServiceOrg").ToString
                ChiefRankName = dr("ChiefRankName").ToString
                ServicePhone = dr("ServicePhone").ToString
                SServiceDate = dr("SServiceDate").ToString

                '21~30
                FServiceDate = dr("FServiceDate").ToString
                If dr("ZipCode4").ToString <> "" Then
                    ZipCode4 = dr("ZipCode4").ToString
                    If shiftsort.SelectedValue = "2" Then
                        ZipCode4 = "(" & dr("ZipCode4").ToString & ")" & dr("CTName4").ToString & dr("ZipName4").ToString
                    End If
                Else
                    ZipCode4 = ""
                End If
                Address4 = dr("Address4").ToString
                PhoneD = dr("PhoneD").ToString()
                PhoneN = dr("PhoneN").ToString()
                CellPhone = dr("CellPhone").ToString()

                If dr("ZipCode1").ToString <> "" Then
                    ZipCode1 = dr("ZipCode1").ToString()
                    If shiftsort.SelectedValue = "2" Then
                        ZipCode1 = "(" & dr("ZipCode1").ToString & ")" & dr("CTName1").ToString & dr("ZipName1").ToString
                    End If
                Else
                    ZipCode1 = ""
                End If
                Address1 = dr("Address1").ToString()
                If dr("ZipCode2").ToString <> "" Then
                    ZipCode2 = dr("ZipCode2").ToString()
                    If shiftsort.SelectedValue = "2" Then
                        ZipCode2 = "(" & dr("ZipCode2").ToString & ")" & dr("CTName2").ToString & dr("ZipName2").ToString
                    End If
                Else
                    ZipCode2 = ""
                End If
                Address2 = dr("Address2").ToString()

                '31~41
                Email = dr("Email").ToString()
                IdentityID = dr("IdentityID").ToString
                MIdentityID = dr("MIdentityID").ToString
                If shiftsort.SelectedValue = "1" Then
                    IdentityID = Replace(dr("IdentityID").ToString, ",", "，")
                Else
                    IdentityID = TIMS.Get_IdentityName(dr("IdentityID").ToString, dt1, "，")
                    MIdentityID = TIMS.Get_IdentityName(dr("IdentityID").ToString, dt1, "，")
                End If
                'by Vicient
                If shiftsort.SelectedValue = "1" Then
                    KNID = dr("KNID").ToString()
                Else
                    KNID = dr("knName").ToString()
                End If

                SubsidyID = dr("SubsidyID").ToString()
                If shiftsort.SelectedValue = "1" Then
                    SubsidyID = dr("SubsidyID").ToString()
                Else
                    SubsidyID = dr("SubsidyName").ToString()
                End If
                If dr("OpenDate").ToString <> "" Then
                    OpenDate = FormatDateTime(dr("OpenDate"), DateFormat.ShortDate)
                Else
                    OpenDate = ""
                End If
                If dr("CloseDate").ToString <> "" Then
                    CloseDate = FormatDateTime(dr("CloseDate"), DateFormat.ShortDate)
                Else
                    CloseDate = ""
                End If
                If dr("EnterDate").ToString <> "" Then
                    EnterDate = FormatDateTime(dr("EnterDate"), DateFormat.ShortDate)
                Else
                    EnterDate = ""
                End If
                HandTypeID = dr("HandTypeID").ToString
                If shiftsort.SelectedValue = "2" Then
                    HandTypeID = dr("HandTypeName").ToString
                End If
                HandLevelID = dr("HandLevelID").ToString
                If shiftsort.SelectedValue = "2" Then
                    HandLevelID = dr("HandLevelName").ToString
                End If
                EmergencyContact = dr("EmergencyContact").ToString()
                EmergencyRelation = dr("EmergencyRelation").ToString()

                '42~50
                EmergencyPhone = dr("EmergencyPhone").ToString()
                If dr("ZipCode3").ToString <> "" Then
                    ZipCode3 = dr("ZipCode3").ToString()
                    If shiftsort.SelectedValue = "2" Then
                        ZipCode3 = "(" & dr("ZipCode3").ToString & ")" & dr("CTName3").ToString & dr("ZipName3").ToString
                    End If
                Else
                    ZipCode3 = ""
                End If
                Address3 = dr("Address3").ToString()
                PriorWorkOrg1 = dr("PriorWorkOrg1").ToString()
                Title1 = dr("Title1").ToString
                SOfficeYM1 = ""
                If dr("SOfficeYM1").ToString <> "" Then
                    SOfficeYM1 = FormatDateTime(dr("SOfficeYM1"), DateFormat.ShortDate)
                End If
                FOfficeYM1 = ""
                If dr("FOfficeYM1").ToString <> "" Then
                    FOfficeYM1 = FormatDateTime(dr("FOfficeYM1"), DateFormat.ShortDate)
                End If
                PriorWorkOrg2 = dr("PriorWorkOrg2").ToString()
                Title2 = dr("Title2").ToString
                SOfficeYM2 = ""
                If dr("SOfficeYM2").ToString <> "" Then
                    SOfficeYM2 = FormatDateTime(dr("SOfficeYM2"), DateFormat.ShortDate)
                End If

                '51~56
                FOfficeYM2 = ""
                If dr("FOfficeYM2").ToString <> "" Then
                    FOfficeYM2 = FormatDateTime(dr("FOfficeYM2"), DateFormat.ShortDate)
                End If
                PriorWorkPay = dr("PriorWorkPay").ToString
                Traffic = dr("Traffic").ToString()
                If shiftsort.SelectedValue = "1" Then
                    Traffic = dr("Traffic").ToString()
                Else
                    Traffic = dr("TrafficName").ToString()
                End If
                RealJobless = dr("RealJobless").ToString()
                JoblessID = dr("JoblessID").ToString
                If shiftsort.SelectedValue = "2" Then
                    JoblessID = dr("JoblessName").ToString
                End If
                ShowDetail = dr("ShowDetail").ToString
                If shiftsort.SelectedValue = "2" Then
                    If dr("ShowDetail").ToString = "Y" Then
                        ShowDetail = "是"
                    Else
                        ShowDetail = "否"
                    End If
                End If
                LevelNo = dr("LevelNo").ToString
                If shiftsort.SelectedValue = "1" Then
                    EnterChannel = dr("EnterChannel").ToString
                    TRNDMode = dr("TRNDMode").ToString
                    TRNDType = dr("TRNDType").ToString
                    BudgetID = dr("BudgetID").ToString
                Else
                    EnterChannel = dr("EnterChannelName").ToString
                    TRNDMode = dr("TRNDModeName").ToString
                    TRNDType = dr("TRNDTypeName").ToString
                    BudgetID = dr("BudName").ToString
                End If
                IsAgree = dr("IsAgree").ToString
                PMode = dr("PMode").ToString

                ForeName = dr("ForeName").ToString
                ForeTitle = dr("ForeTitle").ToString
                If shiftsort.SelectedValue = "1" Then
                    ForeSex = dr("ForeSex").ToString
                Else
                    ForeSex = dr("ForeSexName").ToString
                End If
                If IsDate(dr("ForeBirth")) Then
                    ForeBirth = FormatDateTime(dr("ForeBirth"), 2)
                Else
                    ForeBirth = ""
                End If
                ForeIDNO = dr("ForeIDNO").ToString
                If shiftsort.SelectedValue = "1" Then
                    ForeZip = dr("ForeZip").ToString
                Else
                    ForeZip = dr("ForeZipName").ToString
                End If
                ForeAddr = dr("ForeAddr").ToString

                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    '企訓專用
                    Dim SOCID As Integer = dr("SOCID")
                    Dim dr1 As DataRow
                    Dim dt2 As DataTable

                    sql = "SELECT * FROM Stud_ServicePlace WHERE SOCID='" & SOCID & "'"
                    dr1 = DbAccess.GetOneRow(sql, objconn)

                    AcctMode = ""
                    AcctHeadNo = ""
                    AcctExNo = ""
                    PostNo = ""
                    AcctNo = ""
                    BankName = ""
                    ExBankName = ""
                    Uname = ""
                    Intaxno = ""
                    ServDept = ""
                    JobTitle = ""
                    Zip = ""
                    Addr = ""
                    Tel = ""
                    Fax = ""
                    SDate = ""
                    SJDate = ""
                    SPDate = ""
                    If Not dr1 Is Nothing Then
                        If Not IsDBNull(dr1("AcctMode")) Then
                            If dr1("AcctMode") Then
                                If shiftsort.SelectedValue = "1" Then
                                    AcctMode = 1
                                Else
                                    AcctMode = "金融"
                                End If
                                AcctHeadNo = dr1("AcctHeadNo").ToString
                                AcctExNo = dr1("AcctExNo").ToString
                                AcctNo = dr1("AcctNo").ToString
                                BankName = dr1("BankName").ToString
                                ExBankName = dr1("ExBankName").ToString
                            Else
                                If shiftsort.SelectedValue = "1" Then
                                    AcctMode = 0
                                Else
                                    AcctMode = "郵政"
                                End If
                                PostNo = dr1("PostNo").ToString
                                AcctNo = dr1("AcctNo").ToString
                            End If
                        End If

                        If dr1("FirDate").ToString <> "" Then
                            FirDate = FormatDateTime(dr1("FirDate"), 2)
                        End If
                        Uname = dr1("Uname").ToString
                        Intaxno = dr1("Intaxno").ToString
                        ServDept = dr1("ServDept").ToString
                        JobTitle = dr1("JobTitle").ToString
                        Zip = dr1("Zip").ToString
                        Addr = dr1("Addr").ToString
                        Tel = dr1("Tel").ToString
                        Fax = dr1("Fax").ToString
                        If dr1("SDate").ToString <> "" Then
                            SDate = FormatDateTime(dr1("SDate"), 2)
                        End If
                        If dr1("SJDate").ToString <> "" Then
                            SJDate = FormatDateTime(dr1("SJDate"), 2)
                        End If
                        If dr1("SPDate").ToString <> "" Then
                            SPDate = FormatDateTime(dr1("SPDate"), 2)
                        End If
                    End If

                    Q1 = ""
                    Q3 = ""
                    Q3_Other = ""
                    Q4 = ""
                    Q5 = ""
                    Q61 = ""
                    Q62 = ""
                    Q63 = ""
                    Q64 = ""
                    sql = "SELECT * FROM Stud_TrainBG WHERE SOCID='" & SOCID & "'"
                    dr1 = DbAccess.GetOneRow(sql, objconn)
                    If Not dr1 Is Nothing Then
                        If dr1("Q1") Then
                            Q1 = "Y"
                        Else
                            Q1 = "N"
                        End If
                        Q3 = dr1("Q3").ToString
                        Q3_Other = dr1("Q3_Other").ToString
                        Q4 = dr1("Q4").ToString
                        If Not IsDBNull(dr1("Q5")) Then
                            If dr1("Q5") Then
                                Q5 = "Y"
                            Else
                                Q5 = "N"
                            End If
                        End If
                        Q61 = dr1("Q61").ToString
                        Q62 = dr1("Q62").ToString
                        Q63 = dr1("Q63").ToString
                        Q64 = dr1("Q64").ToString
                    End If

                    Q2 = ""
                    sql = "SELECT * FROM Stud_TrainBGQ2 WHERE SOCID='" & SOCID & "'"
                    dt2 = DbAccess.GetDataTable(sql, objconn)
                    For Each dr1 In dt2.Rows
                        If Q2 = "" Then
                            Q2 = dr1("Q2").ToString
                        Else
                            Q2 += "，" & dr1("Q2").ToString
                        End If
                    Next
                End If

                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    '企訓專用
                    sql = "INSERT INTO [Sheet1$]"
                    sql += "("
                    sql += "學號,中文姓名,LastName,FirstName,身分證號碼,性別,"
                    sql += "出生日期,最高學歷,學校名稱,"
                    sql += "科系,畢業狀況,"
                    sql += "聯絡電話_日,聯絡電話_夜,"
                    sql += "行動電話,通訊地址郵遞區號,通訊地址,"
                    sql += "電子郵件帳號,參訓身分別,主要參訓身分別,開訓日期,結訓日期,"
                    sql += "報到日期,障礙類別,障礙等級,緊急通知人姓名,緊急通知人關係,"
                    sql += "緊急通知人電話,緊急通知人地址郵遞區號,緊急通知人地址,"
                    sql += "報名管道,個資法意願,"
                    sql += "撥款方式,郵政_局號,金融_總代號,金融_分支代號,帳號,銀行名稱,分行名稱,第一次投保勞保日,公司名稱,統編,公司電話,"
                    sql += "公司傳真,公司地址郵遞區號,公司地址,目前任職部門,職稱,"
                    sql += "個人到任目前任職公司起日,個人到任目前職務起日,最近升遷日期,是否由公司推薦參訓,參訓動機,"
                    sql += "訓後動向,訓後動向其他說明,服務單位行業別,服務單位是否屬於中小企業,個人工作年資,"
                    sql += "在這家公司的年資,在這職位的年資,最近升遷離本職幾年,是否提供基本資料查詢"
                    sql += ")VALUES ("
                    sql += "'" & StudentID & "' , '" & Name & "' , '" & LastName & "', '" & FirstName & "' , '" & IDNO & "' , '" & SEX & "' ,"
                    sql += "'" & Birthday & "' , '" & DegreeID & "' , '" & School & "' ,"
                    sql += "'" & Department & "' , '" & GradID & "',"
                    sql += "'" & PhoneD & "' , '" & PhoneN & "' ,"
                    sql += "'" & CellPhone & "' , '" & ZipCode1 & "' , '" & Address1 & "' ,"
                    sql += "'" & Email & "' , '" & IdentityID & "' , '" & MIdentityID & "' ,'" & OpenDate & "' , '" & CloseDate & "' ,"
                    sql += "'" & EnterDate & "' , '" & HandTypeID & "' , '" & HandLevelID & "' , '" & EmergencyContact & "' , '" & EmergencyRelation & "' ,"
                    sql += "'" & EmergencyPhone & "' , '" & ZipCode3 & "' , '" & Address3 & "',"
                    sql += "'" & EnterChannel & "','" & IsAgree & "',"
                    sql += "'" & AcctMode & "','" & PostNo & "','" & AcctHeadNo & "','" & AcctExNo & "','" & AcctNo & "' ,'" & BankName & "','" & ExBankName & "', '" & FirDate & "' , '" & Uname & "' , '" & Intaxno & "' , '" & Tel & "' ,"
                    sql += "'" & Fax & "' , '" & Zip & "' , '" & Addr & "' , '" & ServDept & "' , '" & JobTitle & "' ,"
                    sql += "'" & SDate & "' , '" & SJDate & "' , '" & SPDate & "' , '" & Q1 & "' , '" & Q2 & "' ,"
                    sql += "'" & Q3 & "' , '" & Q3_Other & "' , '" & Q4 & "' , '" & Q5 & "' , '" & Q61 & "' ,"
                    sql += "'" & Q62 & "' , '" & Q63 & "' , '" & Q64 & "' , '" & ShowDetail & "'"
                    sql += ")"
                Else
                    sql = "INSERT INTO [Sheet1$]"
                    sql += "("
                    sql += "學號,中文姓名,LastName,FirstName,身分證號碼,性別,"
                    sql += "身分別,非本國人身分別,原屬國籍,護照或工作證號,出生日期,婚姻狀況,最高學歷,學校名稱,"
                    sql += "科系,畢業狀況,兵役,軍種,職務_兵役,"
                    sql += "階級,服務單位名稱,主管階級姓名,單位電話,服役起日期,"
                    sql += "服役迄日期,服役單位地址郵遞區號,服役單位地址,聯絡電話_日,聯絡電話_夜,"
                    sql += "手機,通訊地址郵遞區號,通訊地址,戶籍地址郵遞區號,戶籍地址,"
                    sql += "電子郵件帳號,參訓身分別,主要參訓身分別,津貼類別,開訓日期,結訓日期,"
                    sql += "報到日期,障礙類別,障礙等級,緊急通知人姓名,緊急通知人關係,"
                    sql += "緊急通知人電話,緊急通知人地址郵遞區號,緊急通知人地址,受訓前服務單位1,受訓前服務單位1職稱,"
                    sql += "受訓前服務單位1任職起日,受訓前服務單位1任職迄日,受訓前服務單位2,受訓前服務單位2職稱,受訓前服務單位2任職起日,"
                    sql += "受訓前服務單位2任職迄日,受訓前薪資,受訓前真正失業週數,受訓前失業週數,交通方式,"
                    sql += "是否提供基本資料查詢,報名階段,報名管道,推介種類,券別種類,預算別,個資法意願,自費公費, "
                    sql += "國內親屬資料_姓名,國內親屬資料_稱謂,國內親屬資料_性別,國內親屬資料_生日,國內親屬資料_身分證號碼,國內親屬資料_郵遞區號,國內親屬資料_地址,原住民民族別 "
                    sql += ")VALUES ("
                    sql += "'" & StudentID & "' , '" & Name & "' , '" & LastName & "', '" & FirstName & "' , '" & IDNO & "' , '" & SEX & "' ,"
                    sql += "'" & PassPortNO & "' ,'" & ChinaOrNot & "','" & Nationality & "','" & PPNO & "', '" & Birthday & "' , '" & MaritalStatus & "' , '" & DegreeID & "' , '" & School & "' ,"
                    sql += "'" & Department & "' , '" & GradID & "' , '" & MilitaryID & "' , '" & ServiceID & "' , '" & MilitaryAppointment & "' ,"
                    sql += "'" & MilitaryRank & "' , '" & ServiceOrg & "' , '" & ChiefRankName & "' , '" & ServicePhone & "' , '" & SServiceDate & "' ,"
                    sql += "'" & FServiceDate & "' , '" & ZipCode4 & "' , '" & Address4 & "' , '" & PhoneD & "' , '" & PhoneN & "' ,"
                    sql += "'" & CellPhone & "' , '" & ZipCode1 & "' , '" & Address1 & "' , '" & ZipCode2 & "' , '" & Address2 & "' ,"
                    sql += "'" & Email & "' , '" & IdentityID & "' , '" & MIdentityID & "' , '" & SubsidyID & "' , '" & OpenDate & "' , '" & CloseDate & "' ,"
                    sql += "'" & EnterDate & "' , '" & HandTypeID & "' , '" & HandLevelID & "' , '" & EmergencyContact & "' , '" & EmergencyRelation & "' ,"
                    sql += "'" & EmergencyPhone & "' , '" & ZipCode3 & "' , '" & Address3 & "' , '" & PriorWorkOrg1 & "' , '" & Title1 & "' ,"
                    sql += "'" & SOfficeYM1 & "' , '" & FOfficeYM1 & "' , '" & PriorWorkOrg2 & "' , '" & Title2 & "' , '" & SOfficeYM2 & "' ,"
                    sql += "'" & FOfficeYM2 & "' , '" & PriorWorkPay & "' , '" & RealJobless & "' , '" & JoblessID & "' , '" & Traffic & "' ,"
                    sql += "'" & ShowDetail & "','" & LevelNo & "','" & EnterChannel & "','" & TRNDMode & "','" & TRNDType & "','" & BudgetID & "','" & IsAgree & "','" & PMode & "',"
                    sql += "'" & ForeName & "','" & ForeTitle & "','" & ForeSex & "','" & ForeBirth & "','" & ForeIDNO & "','" & ForeZip & "','" & ForeAddr & "','" & KNID & "'"
                    sql += ")"
                End If


                cmd = New OleDb.OleDbCommand(sql, conn)
                Try
                    If conn.State = ConnectionState.Closed Then conn.Open()
                    cmd.ExecuteNonQuery()
                    'If conn.State = ConnectionState.Open Then conn.Close()
                Catch ex As Exception
                    If conn.State = ConnectionState.Open Then conn.Close()
                    Throw ex
                End Try
            Next
            If conn.State = ConnectionState.Open Then conn.Close()
            '根據路徑建立資料庫連線，並取出學員資料填入----------------------------------------------------------End

            '將新建立的excel存入記憶體下載------------------------------------------------Start
            Dim strErrmsg As String = ""
            strErrmsg = ""
            Try
                Dim fr As New System.IO.FileStream(MyPath, IO.FileMode.Open)
                Dim br As New System.IO.BinaryReader(fr)
                Dim buf(fr.Length) As Byte

                fr.Read(buf, 0, fr.Length)
                fr.Close()

                Response.Clear()
                Response.ClearHeaders()
                Response.Buffer = True
                Response.AddHeader("content-disposition", "attachment;filename=" & HttpUtility.UrlEncode(MyFileName, System.Text.Encoding.UTF8))
                Response.ContentType = "Application/vnd.ms-Excel"
                'Common.RespWrite(Me, br.ReadBytes(fr.Length))
                Response.BinaryWrite(buf)
            Catch ex As Exception
                strErrmsg = ""
                strErrmsg += "無法存取該檔案!!!" & vbCrLf
                strErrmsg += " (若連續出現多次(3次以上)請連絡系統管理者協助，謝謝)(造成您的不便深感抱歉) " & vbCrLf
                strErrmsg += ex.ToString & vbCrLf
            Finally
                '刪除Temp中的資料
                If IO.File.Exists(MyPath) Then IO.File.Delete(MyPath)
                If strErrmsg = "" Then Response.End()
            End Try
            If strErrmsg <> "" Then
                Common.MessageBox(Me, strErrmsg)
            End If
            '將新建立的excel存入記憶體下載------------------------------------------------End
        End If
    End Sub

    Dim StudentIDBasic As String
    '匯入學員資料
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        'Dim MyFile As System.IO.File
        Dim MyFileName As String
        Dim MyFileType As String
        Dim flag As String
        If File1.Value <> "" Then
            '檢查檔案格式與大小-----------------------------------------------------Start
            If File1.PostedFile.ContentLength = 0 Then
                Common.MessageBox(Me, "檔案位置錯誤!")
                Exit Sub
            Else
                '取出檔案名稱
                MyFileName = Split(File1.PostedFile.FileName, "\")((Split(File1.PostedFile.FileName, "\")).Length - 1)
                '取出檔案類型
                If MyFileName.IndexOf(".") = -1 Then

                    Common.MessageBox(Me, "檔案類型錯誤!")
                    Exit Sub
                Else
                    MyFileType = Split(MyFileName, ".")((Split(MyFileName, ".")).Length - 1)
                    If MyFileType = "csv" Then
                        flag = ","
                    Else
                        Common.MessageBox(Me, "檔案類型錯誤，必須為CSV檔!")
                        Exit Sub
                    End If
                End If
            End If
            '檢查檔案格式與大小-----------------------------------------------------End

            '上傳檔案
            File1.PostedFile.SaveAs(Server.MapPath("~/SD/03/Temp/" & MyFileName))
            'Common.MessageBox(Me, Request.BinaryRead(File1.PostedFile.ContentLength).ToString)

            '將檔案讀出放入記憶體
            Dim sr As System.IO.Stream
            Dim srr As System.IO.StreamReader
            sr = IO.File.OpenRead(Server.MapPath("~/SD/03/Temp/" & MyFileName))
            srr = New System.IO.StreamReader(sr, System.Text.Encoding.Default)

            Dim RowIndex As Integer = 0
            Dim OneRow As String
            'Dim col As String                       '欄位
            Dim colArray As Array

            '取出資料庫的所有欄位---------------------------------------------------Start
            Dim sql As String
            'Dim dtClass As DataTable
            'Dim dtMain As DataTable
            'Dim dtSub As DataTable
            'Dim dtStuOfClass As DataTable
            Dim dr As DataRow
            Dim da As OracleDataAdapter = Nothing
            'Dim da1 As OracleDataAdapter
            'Dim da2 As OracleDataAdapter
            'Dim da3 As OracleDataAdapter
            Dim trans As OracleTransaction = Nothing
            'Dim conn As OracleConnection = DbAccess.GetConnection
            Dim dt As DataTable
            Dim STDate As Date
            Dim i As Integer

            sql = "SELECT STDate FROM Class_ClassInfo WHERE OCID='" & OCIDValue1.Value & "'"
            STDate = DbAccess.ExecuteScalar(sql, objconn)

            Dim BasicSID As String = TIMS.Get_DateNo
            Dim SIDNum As Integer = 1
            Dim SID As String
            Dim Reason As String                '儲存錯誤的原因
            Dim dtWrong As New DataTable            '儲存錯誤資料的DataTable
            Dim drWrong As DataRow

            '建立錯誤資料格式Table----------------Start
            dtWrong.Columns.Add(New DataColumn("Index"))
            dtWrong.Columns.Add(New DataColumn("Name"))
            dtWrong.Columns.Add(New DataColumn("StudentID"))
            dtWrong.Columns.Add(New DataColumn("IDNO"))
            dtWrong.Columns.Add(New DataColumn("Reason"))
            '建立錯誤資料格式Table----------------End


            '取出所有鍵值當判斷-----------------------------------Start
            sql = "SELECT BudID FROM Plan_Budget WHERE TPlanID='" & sm.UserInfo.TPlanID & "' and Syear='" & sm.UserInfo.Years & "'"
            Plan_Budget = DbAccess.GetDataTable(sql, objconn)

            'sql = "SELECT * FROM Key_Degree"
            sql = "SELECT * FROM Key_Degree WHERE 1=1 AND DegreeType IN ('0','1')"
            Key_Degree = DbAccess.GetDataTable(sql, objconn)
            sql = "SELECT * FROM Key_GradState"
            Key_GradState = DbAccess.GetDataTable(sql, objconn)
            sql = "SELECT * FROM Key_Military"
            Key_Military = DbAccess.GetDataTable(sql, objconn)
            sql = "SELECT * FROM Key_Identity"
            Key_Identity = DbAccess.GetDataTable(sql, objconn)
            sql = "SELECT * FROM Key_Subsidy"
            Key_Subsidy = DbAccess.GetDataTable(sql, objconn)
            sql = "SELECT * FROM Key_HandicatType"
            Key_HandicatType = DbAccess.GetDataTable(sql, objconn)
            sql = "SELECT * FROM Key_HandicatLevel"
            Key_HandicatLevel = DbAccess.GetDataTable(sql, objconn)
            'sql = "SELECT * FROM Key_JoblessWeek"
            If CInt(Me.sm.UserInfo.Years) >= 2010 Then
                sql = " SELECT * FROM Key_JoblessWeek  where  joblessid in  ('04','05','06') "
            Else
                sql = " SELECT * FROM Key_JoblessWeek  where  joblessid in  ('01','02','03') "
            End If
            Key_JoblessWeek = DbAccess.GetDataTable(sql, objconn)
            '取出所有鍵值當判斷-----------------------------------End

            '建立StudentID值
            sql = "" & vbCrLf
            sql += " SELECT a.Years,b.ClassID,a.CyclType " & vbCrLf
            sql += " FROM Class_ClassInfo a " & vbCrLf
            sql += " join ID_Class b ON a.CLSID=b.CLSID" & vbCrLf
            sql += " where 1=1" & vbCrLf
            sql += " and a.OCID ='" & OCIDValue1.Value & "'" & vbCrLf
            dr = DbAccess.GetOneRow(sql, objconn)
            Dim StudentID As String
            StudentIDBasic = dr("Years").ToString & "0" & dr("ClassID").ToString & dr("CyclType").ToString

            If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                '企訓專用
                Do While srr.Peek >= 0
                    OneRow = srr.ReadLine
                    If Replace(",", OneRow, "") = "" Then
                        Exit Do
                    End If
                    If RowIndex <> 0 Then
                        Reason = ""
                        colArray = Split(OneRow, flag)
                        '建立StudentID欄位值

                        Reason += CheckImportData(colArray)

                        '通過檢查，開始輸入資料---------------------Start
                        If Reason = "" Then
                            Dim StudentIDNum As String = colArray(0).ToString
                            Dim Name As String = colArray(1).ToString
                            Dim LastName As String = colArray(2).ToString
                            Dim FirstName As String = colArray(3).ToString
                            Dim IDNO As String = colArray(4).ToString
                            Dim Sex As String = colArray(5).ToString
                            Dim Birthday As String = colArray(6).ToString
                            Dim DegreeID As String = colArray(7).ToString

                            Dim School As String
                            If colArray(8).ToString = Nothing Then
                                School = "不詳"
                            Else
                                School = colArray(8).ToString
                            End If

                            Dim Department As String
                            If colArray(9).ToString = Nothing Then
                                Department = "不詳"
                            Else
                                Department = colArray(9).ToString
                            End If
                            Dim GraduateStatus As String = colArray(10).ToString
                            'Dim MilitaryID As String = colArray(11).ToString
                            Dim PhoneD As String = colArray(11).ToString
                            Dim PhoneN As String = colArray(12).ToString
                            Dim CellPhone As String = colArray(13).ToString
                            Dim ZipCode1 As String = colArray(14).ToString
                            Dim Address As String = colArray(15).ToString
                            Dim Email As String = colArray(16).ToString
                            Dim IdentityID As String = colArray(17).ToString
                            Dim MIdentityID As String = colArray(18).ToString
                            Dim OpenDate As String = colArray(19).ToString
                            Dim CloseDate As String = colArray(20).ToString
                            Dim EnterDate As String = colArray(21).ToString
                            Dim HandTypeID As String = colArray(22).ToString
                            Dim HandLevelID As String = colArray(23).ToString
                            Dim EmergencyContact As String = colArray(24).ToString
                            Dim EmergencyRelation As String = colArray(25).ToString
                            Dim EmergencyPhone As String = colArray(26).ToString
                            Dim ZipCode3 As String = colArray(27).ToString
                            Dim EmergencyAddress As String = colArray(28).ToString
                            Dim EnterChannel As String = colArray(29).ToString
                            Dim IsAgree As String = colArray(30).ToString
                            Dim AcctMode As String = colArray(31).ToString
                            Dim PostNo As String = colArray(32).ToString
                            Dim AcctHeadNo As String = colArray(33).ToString
                            Dim AcctExNo As String = colArray(34).ToString
                            Dim AcctNo As String = colArray(35).ToString
                            Dim BankName As String = colArray(36).ToString
                            Dim ExBankName As String = colArray(37).ToString
                            Dim FirDate As String = colArray(38).ToString
                            Dim Uname As String = colArray(39).ToString
                            Dim Intaxno As String = colArray(40).ToString
                            Dim Tel As String = colArray(41).ToString
                            Dim Fax As String = colArray(42).ToString
                            Dim Zip As String = colArray(43).ToString
                            Dim Addr As String = colArray(44).ToString
                            Dim ServDept As String = colArray(45).ToString
                            Dim JobTitle As String = colArray(46).ToString
                            Dim SDate As String = colArray(47).ToString
                            Dim SJDate As String = colArray(48).ToString
                            Dim SPDate As String = colArray(49).ToString
                            Dim Q1 As String = colArray(50).ToString
                            Dim Q2 As String = colArray(51).ToString
                            Dim Q3 As String = colArray(52).ToString
                            Dim Q3_Other As String = colArray(53).ToString
                            Dim Q4 As String = colArray(54).ToString
                            Dim Q5 As String = colArray(55).ToString
                            Dim Q61 As String = colArray(56).ToString
                            Dim Q62 As String = colArray(57).ToString
                            Dim Q63 As String = colArray(58).ToString
                            Dim Q64 As String = colArray(59).ToString
                            Dim ShowDetail As String = colArray(60).ToString
                            Dim SOCID As Integer

                            '建立StudentID欄位值
                            If Int(StudentIDNum) < 10 Then
                                StudentID = StudentIDBasic & "0" & Int(StudentIDNum)
                            Else
                                StudentID = StudentIDBasic & Int(StudentIDNum)
                            End If

                            '建立SID欄位值
                            sql = "SELECT * FROM Stud_StudentInfo WHERE upper(IDNO)= upper('" & TIMS.ChangeIDNO(IDNO) & "') and Birthday=" & TIMS.to_date(Birthday)
                            dr = DbAccess.GetOneRow(sql, objconn)
                            If dr Is Nothing Then
                                If SIDNum < 10 Then
                                    SID = BasicSID & "0" & SIDNum
                                Else
                                    SID = BasicSID & SIDNum
                                End If
                            Else
                                SID = dr("SID")
                            End If
                            '假如此班無個人參加紀錄
                            Try
                                '2006/03/28 add conn by matt
                                trans = DbAccess.BeginTrans(objconn)

                                '檢查班級學員檔是否有此人
                                '有的話跳過匯入程序
                                '沒有則新增匯入
                                sql = "SELECT * FROM Class_StudentsOfClass WHERE SID='" & SID & "' and OCID='" & OCIDValue1.Value & "'"
                                dt = DbAccess.GetDataTable(sql, da, trans)

                                If dt.Rows.Count = 0 Then
                                    dr = dt.NewRow
                                    dt.Rows.Add(dr)

                                    dr("OCID") = OCIDValue1.Value
                                    dr("SID") = SID
                                    dr("StudentID") = StudentID
                                    dr("EnterDate") = IIf(EnterDate = "", Convert.DBNull, EnterDate)
                                    dr("OpenDate") = IIf(OpenDate = "", STDate, OpenDate)
                                    dr("CloseDate") = IIf(CloseDate = "", Convert.DBNull, CloseDate)
                                    dr("StudStatus") = 1
                                    dr("BudgetID") = "03"
                                    dr("EnterChannel") = IIf(EnterChannel = "", Convert.DBNull, EnterChannel)
                                    For i = 0 To Split(IdentityID, "，").Length - 1
                                        If dr("IdentityID").ToString = "" Then
                                            If Split(IdentityID, "，")(i).ToString.Length < 2 Then
                                                dr("IdentityID") = "0" & Split(IdentityID, "，")(i)
                                            Else
                                                dr("IdentityID") = Split(IdentityID, "，")(i)
                                            End If
                                        Else
                                            If Split(IdentityID, "，")(i).ToString.Length < 2 Then
                                                dr("IdentityID") += ",0" & Split(IdentityID, "，")(i)
                                            Else
                                                dr("IdentityID") += "," & Split(IdentityID, "，")(i)
                                            End If
                                        End If
                                    Next
                                    If MIdentityID.Length < 2 Then
                                        dr("MIdentityID") = "0" & MIdentityID
                                    Else
                                        dr("MIdentityID") = MIdentityID
                                    End If
                                    dr("SubsidyID") = "01"
                                    dr("ModifyAcct") = sm.UserInfo.UserID
                                    dr("ModifyDate") = Now

                                    DbAccess.UpdateDataTable(dt, da, trans)
                                    SOCID = DbAccess.GetId(trans, "CLASS_STUDENTSOFCLASS_SOCID_SE")

                                    '檢查學員個人基本資料是否有此人
                                    sql = "SELECT * FROM Stud_StudentInfo WHERE SID='" & SID & "'"
                                    dt = DbAccess.GetDataTable(sql, da, trans)
                                    If dt.Rows.Count = 0 Then
                                        dr = dt.NewRow
                                        dt.Rows.Add(dr)

                                        dr("SID") = SID
                                        dr("IDNO") = TIMS.ChangeIDNO(IDNO)
                                        dr("Name") = Name
                                        dr("EngName") = LastName & " " & FirstName
                                        dr("PassPortNO") = 1 '預設為本國
                                        dr("Sex") = Sex
                                        dr("Birthday") = Birthday
                                        dr("MaritalStatus") = Convert.DBNull
                                        If DegreeID.Length < 2 Then
                                            dr("DegreeID") = "0" & DegreeID
                                        Else
                                            dr("DegreeID") = DegreeID
                                        End If
                                        If GraduateStatus.Length < 2 Then
                                            dr("GraduateStatus") = "0" & GraduateStatus
                                        Else
                                            dr("GraduateStatus") = GraduateStatus
                                        End If
                                        'If MilitaryID.Length < 2 Then
                                        '    dr("MilitaryID") = "0" & MilitaryID
                                        'Else
                                        '    dr("MilitaryID") = MilitaryID
                                        'End If
                                        dr("MilitaryID") = "03"
                                        dr("IdentityID") = Convert.DBNull
                                        dr("JoblessID") = "01"
                                        dr("RealJobless") = Convert.DBNull

                                        SIDNum += 1
                                    Else
                                        dr = dt.Rows(0)
                                    End If
                                    dr("IsAgree") = IsAgree
                                    dr("ModifyAcct") = sm.UserInfo.UserID
                                    dr("ModifyDate") = Now
                                    DbAccess.UpdateDataTable(dt, da, trans)

                                    sql = "SELECT * FROM Stud_SubData WHERE SID='" & SID & "'"
                                    dt = DbAccess.GetDataTable(sql, da, trans)

                                    If dt.Rows.Count = 0 Then
                                        dr = dt.NewRow
                                        dt.Rows.Add(dr)

                                        dr("SID") = SID
                                        dr("Name") = Name
                                        dr("School") = School
                                        dr("Department") = Department
                                        dr("ZipCode1") = ZipCode1
                                        dr("Address") = Address
                                        dr("ZipCode2") = Convert.DBNull
                                        dr("HouseholdAddress") = Convert.DBNull
                                        dr("Email") = IIf(Email = "", Convert.DBNull, Email)
                                        dr("PhoneD") = IIf(PhoneD = "", Convert.DBNull, PhoneD)
                                        dr("PhoneN") = IIf(PhoneN = "", Convert.DBNull, PhoneN)
                                        dr("CellPhone") = IIf(CellPhone = "", Convert.DBNull, CellPhone)
                                        dr("EmergencyContact") = IIf(EmergencyContact = "", Convert.DBNull, EmergencyContact)
                                        dr("EmergencyRelation") = IIf(EmergencyRelation = "", Convert.DBNull, EmergencyRelation)
                                        dr("EmergencyPhone") = IIf(EmergencyPhone = "", Convert.DBNull, EmergencyPhone)
                                        dr("ZipCode3") = IIf(ZipCode3 = "", Convert.DBNull, ZipCode3)
                                        dr("EmergencyAddress") = IIf(EmergencyAddress = "", Convert.DBNull, EmergencyAddress)
                                        dr("PriorWorkOrg1") = Convert.DBNull
                                        dr("Title1") = Convert.DBNull
                                        dr("SOfficeYM1") = Convert.DBNull
                                        dr("FOfficeYM1") = Convert.DBNull
                                        dr("SOfficeYM2") = Convert.DBNull
                                        dr("FOfficeYM2") = Convert.DBNull
                                        dr("PriorWorkPay") = Convert.DBNull
                                        dr("Traffic") = Convert.DBNull
                                        dr("ShowDetail") = ShowDetail
                                        dr("ServiceID") = Convert.DBNull
                                        dr("MilitaryAppointment") = Convert.DBNull
                                        dr("MilitaryRank") = Convert.DBNull
                                        dr("SServiceDate") = Convert.DBNull
                                        dr("FServiceDate") = Convert.DBNull
                                        dr("ServiceOrg") = Convert.DBNull
                                        dr("ChiefRankName") = Convert.DBNull
                                        dr("ZipCode4") = Convert.DBNull
                                        dr("ServiceAddress") = Convert.DBNull
                                        dr("ServicePhone") = Convert.DBNull
                                        If HandTypeID = "" Then
                                            dr("HandTypeID") = "0" & HandTypeID
                                        Else
                                            If HandTypeID.Length < 2 Then
                                                dr("HandTypeID") = "0" & HandTypeID
                                            Else
                                                dr("HandTypeID") = HandTypeID
                                            End If
                                        End If
                                        If HandLevelID.ToString <> "" Then
                                            If HandLevelID.Length < 2 Then
                                                dr("HandLevelID") = "0" & HandLevelID
                                            Else
                                                dr("HandLevelID") = HandLevelID
                                            End If
                                        End If
                                        dr("ModifyAcct") = sm.UserInfo.UserID
                                        dr("ModifyDate") = Now

                                        DbAccess.UpdateDataTable(dt, da, trans)
                                    End If

                                    sql = "SELECT * FROM Stud_ServicePlace WHERE SOCID='" & SOCID & "'"
                                    dt = DbAccess.GetDataTable(sql, da, trans)

                                    If dt.Rows.Count = 0 Then
                                        dr = dt.NewRow
                                        dt.Rows.Add(dr)

                                        dr("SOCID") = SOCID
                                        If AcctMode = "1" Then
                                            dr("AcctMode") = True
                                        Else
                                            dr("AcctMode") = False
                                        End If
                                        dr("PostNo") = IIf(PostNo = "", Convert.DBNull, PostNo)
                                        dr("AcctHeadNo") = IIf(AcctHeadNo = "", Convert.DBNull, AcctHeadNo)
                                        dr("AcctExNo") = IIf(AcctExNo = "", Convert.DBNull, AcctExNo)
                                        dr("AcctNo") = AcctNo
                                        dr("BankName") = IIf(BankName = "", Convert.DBNull, BankName)
                                        dr("ExBankName") = IIf(ExBankName = "", Convert.DBNull, ExBankName)
                                        dr("FirDate") = IIf(FirDate = "", Convert.DBNull, FirDate)
                                        dr("Uname") = IIf(Uname = "", Convert.DBNull, Uname)
                                        dr("Intaxno") = IIf(Intaxno = "", Convert.DBNull, Intaxno)
                                        dr("ServDept") = IIf(ServDept = "", Convert.DBNull, ServDept)
                                        dr("JobTitle") = IIf(JobTitle = "", Convert.DBNull, JobTitle)
                                        dr("Zip") = Zip
                                        dr("Addr") = Addr
                                        dr("Tel") = Tel
                                        dr("Fax") = IIf(Fax = "", Convert.DBNull, Fax)
                                        dr("SDate") = IIf(SDate = "", Convert.DBNull, SDate)
                                        dr("SJDate") = IIf(SJDate = "", Convert.DBNull, SJDate)
                                        dr("SPDate") = IIf(SPDate = "", Convert.DBNull, SPDate)

                                        dr("ModifyAcct") = sm.UserInfo.UserID
                                        dr("ModifyDate") = Now

                                        DbAccess.UpdateDataTable(dt, da, trans)
                                    End If

                                    sql = "SELECT * FROM Stud_TrainBG WHERE SOCID='" & SOCID & "'"
                                    dt = DbAccess.GetDataTable(sql, da, trans)

                                    If dt.Rows.Count = 0 Then
                                        dr = dt.NewRow
                                        dt.Rows.Add(dr)

                                        dr("SOCID") = SOCID
                                        dr("Q1") = IIf(Q1 = "Y", 1, 0)
                                        dr("Q3") = IIf(Q3 = "", Convert.DBNull, Q3)
                                        dr("Q3_Other") = IIf(Q3_Other = "", Convert.DBNull, Q3_Other)
                                        dr("Q4") = IIf(Q4 = "", Convert.DBNull, Q4)
                                        If Q5 = "" Then
                                            dr("Q5") = Convert.DBNull
                                        Else
                                            dr("Q5") = IIf(Q5 = "Y", 1, 0)
                                        End If
                                        dr("Q61") = IIf(Q61 = "", Convert.DBNull, Q61)
                                        dr("Q62") = IIf(Q62 = "", Convert.DBNull, Q62)
                                        dr("Q63") = IIf(Q63 = "", Convert.DBNull, Q63)
                                        dr("Q64") = IIf(Q64 = "", Convert.DBNull, Q64)

                                        dr("ModifyAcct") = sm.UserInfo.UserID
                                        dr("ModifyDate") = Now

                                        DbAccess.UpdateDataTable(dt, da, trans)
                                    End If

                                    sql = "DELETE Stud_TrainBGQ2 WHERE SOCID='" & SOCID & "'"
                                    DbAccess.ExecuteNonQuery(sql, trans)
                                    If Q2 <> "" Then
                                        sql = "SELECT * FROM Stud_TrainBGQ2 WHERE SOCID='" & SOCID & "'"
                                        dt = DbAccess.GetDataTable(sql, da, trans)
                                        If Split(Q2, "，").Length <> 0 Then
                                            For i = 0 To Split(Q2, "，").Length - 1
                                                If dt.Select("Q2='" & Split(Q2, "，")(i) & "'").Length = 0 Then
                                                    dr = dt.NewRow
                                                    dt.Rows.Add(dr)
                                                    dr("SOCID") = SOCID
                                                    dr("Q2") = Split(Q2, "，")(i)
                                                End If
                                            Next
                                            DbAccess.UpdateDataTable(dt, da, trans)
                                        End If
                                    End If
                                    DbAccess.CommitTrans(trans)

                                End If
                            Catch ex As Exception
                                DbAccess.RollbackTrans(trans)
                                Throw ex
                                'drWrong = dtWrong.NewRow
                                'dtWrong.Rows.Add(drWrong)

                                'drWrong("Index") = RowIndex
                                'If colArray.Length > 5 Then
                                '    drWrong("Name") = Name
                                '    drWrong("StudentID") = StudentID
                                '    drWrong("IDNO") = ChangeIDNO(IDNO)
                                '    drWrong("Reason") = "資料庫上傳錯誤"
                                'End If
                            End Try
                        Else
                            '錯誤資料，填入錯誤資料表
                            drWrong = dtWrong.NewRow
                            dtWrong.Rows.Add(drWrong)

                            drWrong("Index") = RowIndex
                            If colArray.Length > 5 Then
                                drWrong("Name") = colArray(1)
                                drWrong("StudentID") = colArray(0)
                                drWrong("IDNO") = TIMS.ChangeIDNO(colArray(4))
                                drWrong("Reason") = Reason
                            End If
                        End If
                    End If
                    RowIndex += 1
                Loop
            Else
                '一般狀況匯入學員
                'conn = DbAccess.GetConnection

                '開始判別欄位存入-------------------------------------------------------Start
                Do While srr.Peek >= 0
                    OneRow = srr.ReadLine
                    If RowIndex <> 0 Then
                        Reason = ""
                        colArray = Split(OneRow, flag)
                        '建立StudentID欄位值

                        Reason += CheckImportData(colArray)

                        '通過檢查，開始輸入資料---------------------Start
                        If Reason = "" Then
                            If colArray.Length <> 1 Then
                                Try

                                    '建立StudentID欄位值
                                    If Int(colArray(0)) < 10 Then
                                        StudentID = StudentIDBasic & "0" & Int(colArray(0))
                                    Else
                                        StudentID = StudentIDBasic & Int(colArray(0))
                                    End If
                                    '建立SID欄位值
                                    sql = "SELECT * FROM Stud_StudentInfo WHERE IDNO='" & TIMS.ChangeIDNO(colArray(4).ToString) & "'"
                                    dr = DbAccess.GetOneRow(sql, objconn)
                                    If dr Is Nothing Then
                                        If SIDNum < 10 Then
                                            SID = BasicSID & "0" & SIDNum
                                        Else
                                            SID = BasicSID & SIDNum
                                        End If
                                    Else
                                        SID = dr("SID")
                                    End If

                                    '2006/03/28 add conn by matt
                                    trans = DbAccess.BeginTrans(objconn)
                                    sql = "SELECT * FROM Class_StudentsOfClass WHERE SID='" & SID & "' and OCID='" & OCIDValue1.Value & "'"
                                    dt = DbAccess.GetDataTable(sql, da, trans)

                                    If dt.Rows.Count = 0 Then
                                        dr = dt.NewRow
                                        dt.Rows.Add(dr)

                                        dr("OCID") = OCIDValue1.Value
                                        dr("SID") = SID
                                        dr("StudentID") = StudentID
                                        If colArray(40).ToString <> "" Then
                                            dr("EnterDate") = colArray(40)
                                        End If
                                        If Trim(colArray(38).ToString) = "" Then
                                            dr("OpenDate") = STDate
                                        Else
                                            dr("OpenDate") = colArray(38)
                                        End If
                                        If colArray(39).ToString <> "" Then
                                            dr("CloseDate") = colArray(39)
                                        End If
                                        dr("StudStatus") = 1
                                        If colArray(61).ToString <> "" Then
                                            dr("LevelNo") = colArray(61).ToString
                                        End If
                                        If colArray(63).ToString <> "" Then
                                            dr("TRNDMode") = colArray(63).ToString
                                        End If
                                        If colArray(64).ToString <> "" Then
                                            dr("TRNDType") = colArray(64).ToString
                                        End If
                                        If colArray(62).ToString <> "" Then
                                            dr("EnterChannel") = colArray(62).ToString
                                        End If
                                        If colArray(65).ToString <> "" Then
                                            If Int(colArray(65).ToString) < 10 Then
                                                dr("BudgetID") = 0 & Int(colArray(65).ToString)
                                            Else
                                                dr("BudgetID") = Int(colArray(65).ToString)
                                            End If
                                        End If
                                        For i = 0 To Split(colArray(35), "，").Length - 1
                                            If dr("IdentityID").ToString = "" Then
                                                If Split(colArray(35), "，")(i).ToString.Length < 2 Then
                                                    dr("IdentityID") = "0" & Split(colArray(35), "，")(i)
                                                Else
                                                    dr("IdentityID") = Split(colArray(35), "，")(i)
                                                End If
                                            Else
                                                If Split(colArray(35), "，")(i).ToString.Length < 2 Then
                                                    dr("IdentityID") += ",0" & Split(colArray(35), "，")(i)
                                                Else
                                                    dr("IdentityID") += "," & Split(colArray(35), "，")(i)
                                                End If
                                            End If
                                        Next
                                        If colArray(36).ToString.Length < 2 Then
                                            dr("MIdentityID") = "0" & colArray(36).ToString
                                        Else
                                            dr("MIdentityID") = colArray(36).ToString
                                        End If
                                        'by Vicient 原住民別
                                        If colArray(75).ToString <> "" Then
                                            If colArray(75).ToString.Length < 2 Then
                                                dr("Native") = "0" & colArray(75).ToString
                                            Else
                                                dr("Native") = colArray(75).ToString
                                            End If
                                        End If
                                        If colArray(37).ToString <> "" Then
                                            If colArray(37).ToString.Length < 2 Then
                                                dr("SubsidyID") = "0" & colArray(37)
                                            Else
                                                dr("SubsidyID") = colArray(37)
                                            End If
                                        End If
                                        dr("ModifyAcct") = sm.UserInfo.UserID
                                        dr("ModifyDate") = Now
                                        DbAccess.UpdateDataTable(dt, da, trans)

                                        Dim SOCID As Integer = DbAccess.GetId(trans, "CLASS_STUDENTSOFCLASS_SOCID_SE")
                                        sql = "SELECT * FROM Stud_StudentInfo WHERE SID='" & SID & "'"
                                        dt = DbAccess.GetDataTable(sql, da, trans)
                                        If dt.Rows.Count = 0 Then
                                            dr = dt.NewRow
                                            dt.Rows.Add(dr)

                                            dr("SID") = SID
                                            dr("IDNO") = TIMS.ChangeIDNO(colArray(4))
                                            dr("Name") = colArray(1)
                                            dr("EngName") = colArray(2) & " " & colArray(3)
                                            Select Case Convert.ToString(colArray(6))
                                                Case "1", "2"
                                                    dr("PassPortNO") = Convert.ToString(colArray(6))
                                                Case Else
                                                    dr("PassPortNO") = "2"
                                            End Select
                                            If colArray(6).ToString = "2" Then
                                                dr("ChinaOrNot") = colArray(7)
                                                dr("Nationality") = colArray(8)
                                                dr("PPNO") = colArray(9)
                                            Else
                                                dr("ChinaOrNot") = Convert.DBNull
                                                dr("Nationality") = Convert.DBNull
                                                dr("PPNO") = Convert.DBNull
                                            End If
                                            dr("Sex") = colArray(5)
                                            dr("Birthday") = colArray(10)
                                            If colArray(11).ToString = "" Then
                                                dr("MaritalStatus") = Convert.DBNull
                                            Else
                                                dr("MaritalStatus") = colArray(11)
                                            End If
                                            If colArray(12).ToString.Length < 2 Then
                                                dr("DegreeID") = "0" & colArray(12)
                                            Else
                                                dr("DegreeID") = colArray(12)
                                            End If
                                            If colArray(15).ToString.Length < 2 Then
                                                dr("GraduateStatus") = "0" & colArray(15)
                                            Else
                                                dr("GraduateStatus") = colArray(15)
                                            End If
                                            If colArray(16).ToString.Length < 2 Then
                                                dr("MilitaryID") = "0" & colArray(16)
                                            Else
                                                dr("MilitaryID") = colArray(16)
                                            End If
                                            dr("IdentityID") = ""
                                            If colArray(58).ToString <> "" Then
                                                If colArray(58).ToString.Length < 2 Then
                                                    dr("JoblessID") = "0" & colArray(58)
                                                Else
                                                    dr("JoblessID") = colArray(58)
                                                End If
                                            End If
                                            If colArray(57).ToString <> "" Then
                                                dr("RealJobless") = colArray(57)
                                            End If
                                            dr("IsAgree") = colArray(66).ToString
                                            dr("ModifyAcct") = sm.UserInfo.UserID
                                            dr("ModifyDate") = Now

                                            DbAccess.UpdateDataTable(dt, da, trans)
                                            SIDNum += 1
                                        End If

                                        sql = "SELECT * FROM Stud_SubData WHERE SID='" & SID & "'"
                                        dt = DbAccess.GetDataTable(sql, da, trans)
                                        If dt.Rows.Count = 0 Then
                                            dr = dt.NewRow
                                            dt.Rows.Add(dr)
                                            dr("SID") = SID
                                            dr("Name") = colArray(1)
                                            dr("School") = colArray(13)
                                            dr("Department") = colArray(14)
                                            dr("ZipCode1") = colArray(30)
                                            dr("Address") = colArray(31)
                                            If colArray(32).ToString <> "" Then
                                                dr("ZipCode2") = colArray(32)
                                            End If
                                            If colArray(33).ToString <> "" Then
                                                dr("HouseholdAddress") = colArray(33)
                                            End If
                                            If colArray(34).ToString <> "" Then
                                                dr("Email") = colArray(34)
                                            End If
                                            dr("PhoneD") = colArray(27)
                                            dr("PhoneN") = colArray(28)
                                            dr("CellPhone") = colArray(29)
                                            dr("EmergencyContact") = colArray(43)
                                            dr("EmergencyRelation") = colArray(44)
                                            dr("EmergencyPhone") = colArray(45)
                                            If colArray(46).ToString <> "" Then
                                                dr("ZipCode3") = colArray(46)
                                            End If
                                            If colArray(47).ToString <> "" Then
                                                dr("EmergencyAddress") = colArray(47)
                                            End If
                                            If colArray(48).ToString <> "" Then
                                                dr("PriorWorkOrg1") = colArray(48)
                                            End If
                                            If colArray(49).ToString <> "" Then
                                                dr("Title1") = colArray(49)
                                            End If
                                            If colArray(50).ToString <> "" Then
                                                dr("SOfficeYM1") = colArray(50)
                                            End If
                                            If colArray(51).ToString <> "" Then
                                                dr("FOfficeYM1") = colArray(51)
                                            End If
                                            dr("PriorWorkOrg2") = colArray(52)
                                            dr("Title2") = colArray(53)
                                            If colArray(54).ToString <> "" Then
                                                dr("SOfficeYM2") = colArray(54)
                                            End If
                                            If colArray(55).ToString <> "" Then
                                                dr("FOfficeYM2") = colArray(55)
                                            End If
                                            If colArray(56).ToString <> "" Then
                                                dr("PriorWorkPay") = colArray(56)
                                            End If
                                            If colArray(59).ToString <> "" Then
                                                dr("Traffic") = colArray(59)
                                            End If
                                            dr("ShowDetail") = colArray(60)
                                            If Int(colArray(16)) = 4 Then
                                                If Trim(colArray(17).ToString) <> "" Then
                                                    dr("ServiceID") = colArray(17)
                                                End If
                                                If Trim(colArray(18).ToString) <> "" Then
                                                    dr("MilitaryAppointment") = colArray(18)
                                                End If
                                                If Trim(colArray(19).ToString) <> "" Then
                                                    dr("MilitaryRank") = colArray(19)
                                                End If
                                                If Trim(colArray(23).ToString) <> "" Then
                                                    dr("SServiceDate") = colArray(23)
                                                End If
                                                If Trim(colArray(24).ToString) <> "" Then
                                                    dr("FServiceDate") = colArray(24)
                                                End If
                                                If Trim(colArray(20).ToString) <> "" Then
                                                    dr("ServiceOrg") = colArray(20)
                                                End If
                                                If Trim(colArray(21).ToString) <> "" Then
                                                    dr("ChiefRankName") = colArray(21)
                                                End If
                                                If Trim(colArray(25).ToString) <> "" Then
                                                    dr("ZipCode4") = colArray(25)
                                                End If
                                                If Trim(colArray(26).ToString) <> "" Then
                                                    dr("ServiceAddress") = colArray(26)
                                                End If
                                                If Trim(colArray(22).ToString) <> "" Then
                                                    dr("ServicePhone") = colArray(22)
                                                End If
                                            End If

                                            If colArray(41).ToString <> "" Then
                                                If colArray(41).ToString.Length < 2 Then
                                                    dr("HandTypeID") = "0" & colArray(41)
                                                Else
                                                    dr("HandTypeID") = colArray(41)
                                                End If
                                            End If
                                            If colArray(42).ToString <> "" Then
                                                If colArray(42).ToString.Length < 2 Then
                                                    dr("HandLevelID") = "0" & colArray(42)
                                                Else
                                                    dr("HandLevelID") = colArray(42)
                                                End If
                                            End If
                                            If colArray(6).ToString = "2" Then
                                                dr("ForeName") = colArray(68)
                                                dr("ForeTitle") = colArray(69)
                                                dr("ForeSex") = colArray(70)
                                                dr("ForeBirth") = colArray(71)
                                                dr("ForeIDNO") = colArray(72)
                                                dr("ForeZip") = colArray(73)
                                                dr("ForeAddr") = colArray(74)
                                            End If
                                            dr("ModifyAcct") = sm.UserInfo.UserID
                                            dr("ModifyDate") = Now

                                            DbAccess.UpdateDataTable(dt, da, trans)
                                        End If
                                        DbAccess.CommitTrans(trans)

                                    End If
                                Catch ex As Exception
                                    DbAccess.RollbackTrans(trans)
                                    drWrong = dtWrong.NewRow
                                    dtWrong.Rows.Add(drWrong)

                                    drWrong("Index") = RowIndex
                                    If colArray.Length > 5 Then
                                        drWrong("Name") = colArray(1)
                                        drWrong("StudentID") = colArray(0)
                                        drWrong("IDNO") = TIMS.ChangeIDNO(colArray(4))
                                        drWrong("Reason") = "資料庫上傳錯誤"
                                    End If
                                End Try
                            End If
                        Else
                            '錯誤資料，填入錯誤資料表
                            drWrong = dtWrong.NewRow
                            dtWrong.Rows.Add(drWrong)

                            drWrong("Index") = RowIndex
                            If colArray.Length > 5 Then
                                drWrong("Name") = colArray(1)
                                drWrong("StudentID") = colArray(0)
                                drWrong("IDNO") = TIMS.ChangeIDNO(colArray(4))
                                drWrong("Reason") = Reason
                            End If
                        End If
                    End If
                    RowIndex += 1
                Loop
            End If

            '開始判別欄位存入-------------------------------------------------------End

            If dtWrong.Rows.Count = 0 Then
                Common.MessageBox(Me, "資料匯入成功")
            Else
                Session("MyWrongTable") = dtWrong
                Page.RegisterStartupScript("", "<script>if(confirm('資料匯入成功，但有錯誤的資料無法匯入，是否要檢視原因?')){window.open('SD_03_006_Wrong.aspx','','width=500,height=500,location=0,status=0,menubar=0,scrollbars=1,resizable=0');}</script>")
            End If
            sr.Close()
            srr.Close()
            IO.File.Delete(Server.MapPath("~/SD/03/Temp/" & MyFileName))

            Button1_Click(sender, e)
        End If
    End Sub

    Dim IDNOArray As New ArrayList
    '檢查輸入資料
    Function CheckImportData(ByVal colArray As Array) As String
        Dim Reason As String = ""
        Dim SearchEngStr As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ- "
        Dim sql As String
        Dim dr As DataRow

        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '企訓專用
            If colArray.Length <> 61 Then
                'Reason += "欄位數量不正確(應該為61個欄位)<BR>"
                Reason += "欄位對應有誤<BR>"
                Reason += "請注意欄位中是否有半形逗點<BR>"
            Else
                Dim StudentIDNum As String = colArray(0).ToString
                Dim Name As String = colArray(1).ToString
                Dim LastName As String = colArray(2).ToString
                Dim FirstName As String = colArray(3).ToString
                Dim IDNO As String = colArray(4).ToString
                Dim Sex As String = colArray(5).ToString
                Dim Birthday As String = colArray(6).ToString
                Dim DegreeID As String = colArray(7).ToString
                Dim school As String
                '如果是產學訓且未填學校名則預設不詳
                school = colArray(8).ToString
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    If colArray(8).ToString = "" Then
                        school = "不詳"
                    End If
                End If

                '如果是產學訓且未填科系名則預設不詳 add by nick
                Dim Department As String
                Department = colArray(9).ToString
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    If colArray(9).ToString = "" Then
                        Department = "不詳"
                    End If
                End If

                Dim GraduateStatus As String = colArray(10).ToString
                'Dim MilitaryID As String = colArray(11).ToString
                Dim PhoneD As String = colArray(11).ToString
                Dim PhoneN As String = colArray(12).ToString
                Dim CellPhone As String = colArray(13).ToString
                Dim ZipCode1 As String = colArray(14).ToString
                Dim Address As String = colArray(15).ToString
                Dim Email As String = colArray(16).ToString
                Dim IdentityID As String = colArray(17).ToString
                Dim MIdentityID As String = colArray(18).ToString
                Dim OpenDate As String = colArray(19).ToString
                Dim CloseDate As String = colArray(20).ToString
                Dim EnterDate As String = colArray(21).ToString
                Dim HandTypeID As String = colArray(22).ToString
                Dim HandLevelID As String = colArray(23).ToString
                Dim EmergencyContact As String = colArray(24).ToString
                Dim EmergencyRelation As String = colArray(25).ToString
                Dim EmergencyPhone As String = colArray(26).ToString
                Dim ZipCode3 As String = colArray(27).ToString
                Dim EmergencyAddress As String = colArray(28).ToString
                Dim EnterChannel As String = colArray(29).ToString
                Dim IsAgree As String = colArray(30).ToString
                Dim AcctMode As String = colArray(31).ToString
                Dim PostNo As String = colArray(32).ToString
                Dim AcctHeadNo As String = colArray(33).ToString
                Dim AcctExNo As String = colArray(34).ToString
                Dim AcctNo As String = colArray(35).ToString
                Dim BankName As String = colArray(36).ToString
                Dim ExBankName As String = colArray(37).ToString
                Dim FirDate As String = colArray(38).ToString
                Dim Uname As String = colArray(39).ToString
                Dim Intaxno As String = colArray(40).ToString
                Dim Tel As String = colArray(41).ToString
                Dim Fax As String = colArray(42).ToString
                Dim Zip As String = colArray(43).ToString
                Dim Addr As String = colArray(44).ToString
                Dim ServDept As String = colArray(45).ToString
                Dim JobTitle As String = colArray(46).ToString
                Dim SDate As String = colArray(47).ToString
                Dim SJDate As String = colArray(48).ToString
                Dim SPDate As String = colArray(49).ToString
                Dim Q1 As String = colArray(50).ToString
                Dim Q2 As String = colArray(51).ToString
                Dim Q3 As String = colArray(52).ToString
                Dim Q3_Other As String = colArray(53).ToString
                Dim Q4 As String = colArray(54).ToString
                Dim Q5 As String = colArray(55).ToString
                Dim Q61 As String = colArray(56).ToString
                Dim Q62 As String = colArray(57).ToString
                Dim Q63 As String = colArray(58).ToString
                Dim Q64 As String = colArray(59).ToString
                Dim ShowDetail As String = colArray(60).ToString

                If StudentIDNum = "" Then
                    Reason += "必須填寫學號<Br>"
                Else
                    If IsNumeric(StudentIDNum) Then
                        Dim MyKey As String = Int(StudentIDNum)
                        If Int(MyKey) >= 100 Then
                            Reason += "學號必須為在(01~99)範圍內<BR>"
                        Else
                            If Int(MyKey) < 10 Then
                                MyKey = "0" & Int(MyKey)
                            End If

                            sql = "SELECT * FROM Class_StudentsOfClass WHERE OCID='" & OCIDValue1.Value & "' and  StudentID='" & StudentIDBasic & MyKey & "'"
                            dr = DbAccess.GetOneRow(sql, objconn)
                            If Not dr Is Nothing Then
                                Reason += "學號重複<BR>"
                            End If
                        End If
                    Else
                        Reason += "學號必須為數字(01~99)<BR>"
                    End If
                End If

                If Name = "" Then
                    Reason += "必須填寫中文姓名<BR>"
                End If
                '產學訓不擋以下判斷 ===== add by nick
                If Not TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    If LastName = "" Then
                        Reason += "必須填寫英文姓名(LastName)<BR>"
                    Else
                        For i As Integer = 0 To LastName.Length - 1
                            If SearchEngStr.IndexOf(LastName.ToUpper.Chars(i)) = -1 Then
                                Reason += "英文姓名必須只有英文字(LastName)<BR>"
                            End If
                        Next
                    End If
                    If FirstName = "" Then
                        Reason += "必須填寫英文姓名(FirstName)<BR>"
                    Else
                        For i As Integer = 0 To FirstName.Length - 1
                            If SearchEngStr.IndexOf(FirstName.ToUpper.Chars(i)) = -1 Then
                                Reason += "英文姓名必須只有英文字(FirstName)<BR>"
                            End If
                        Next
                    End If
                    If EmergencyContact = "" Then
                        Reason += "必須填寫緊急通知人姓名<BR>"
                    End If
                    If EmergencyRelation = "" Then
                        Reason += "必須填寫緊急通知人關係<BR>"
                    End If
                    If EmergencyPhone = "" Then
                        Reason += "必須填寫緊急通知人電話<BR>"
                    End If
                    If ZipCode3 = "" Then
                        Reason += "必須填寫緊急通知人地址郵遞區號<BR>"
                    Else
                        If IsNumeric(ZipCode3) = False Then
                            Reason += "郵遞區號必須為數字<BR>"
                        End If
                    End If
                    If EmergencyAddress = "" Then
                        Reason += "必須填寫緊急通知人地址<BR>"
                    End If

                End If
                '---------- end -------快樂的產學訓結束

                If IDNO = "" Then
                    Reason += "必須填寫身分證號碼<BR>"
                Else
                    If sm.UserInfo.RoleID = 1 Then
                        Dim IDNOFlag As Boolean = True
                        Dim EngStr As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
                        If IDNO.Length <> 10 Then
                            IDNOFlag = False
                        ElseIf IDNO.Chars(1) <> "1" And IDNO.Chars(1) <> "2" Then
                            IDNOFlag = False
                        ElseIf EngStr.IndexOf(IDNO.ToUpper.Chars(0)) = -1 Then
                            IDNOFlag = False
                        ElseIf IDNO = "A123456789" Then
                            IDNOFlag = False
                        End If

                        If IDNOFlag Then
                            sql = "SELECT * FROM Stud_StudentInfo WHERE IDNO='" & IDNO & "' and SID IN (SELECT SID FROM Class_StudentsOfClass WHERE OCID='" & OCIDValue1.Value & "')"
                            dr = DbAccess.GetOneRow(sql, objconn)
                            If Not dr Is Nothing Then
                                Reason += "此班已經有相同的身分證號碼<BR>"
                            Else
                                Dim Flag As Boolean = True
                                For i As Integer = 0 To IDNOArray.Count - 1
                                    If IDNOArray(i) = IDNO Then
                                        Reason += "檔案中有相同的身分證號碼<BR>"
                                        Flag = False
                                    End If
                                Next
                                If Flag Then
                                    IDNOArray.Add(IDNO)
                                End If
                            End If
                        Else
                            Reason += "身分證號碼錯誤!<BR>"
                        End If
                    Else
                        If TIMS.CheckIDNO(IDNO) Then
                            sql = "SELECT * FROM Stud_StudentInfo WHERE IDNO='" & IDNO & "' and SID IN (SELECT SID FROM Class_StudentsOfClass WHERE OCID='" & OCIDValue1.Value & "')"
                            dr = DbAccess.GetOneRow(sql, objconn)
                            If Not dr Is Nothing Then
                                Reason += "此班已經有相同的身分證號碼<BR>"
                            Else
                                Dim Flag As Boolean = True
                                For i As Integer = 0 To IDNOArray.Count - 1
                                    If IDNOArray(i) = IDNO Then
                                        Reason += "檔案中有相同的身分證號碼<BR>"
                                        Flag = False
                                    End If
                                Next
                                If Flag Then
                                    IDNOArray.Add(IDNO)
                                End If
                            End If
                        Else
                            Reason += "身分證號碼錯誤!請聯絡系統管理員<BR>"
                        End If
                    End If
                End If
                If Sex = "" Then
                    Reason += "必須填寫性別<BR>"
                Else
                    Select Case Sex
                        Case "M", "F"
                        Case Else
                            Reason += "性別代號只能是M或者是F<BR>"
                    End Select
                End If

                If Reason = "" Then
                    If Not TIMS.checkMemberSex(IDNO, Sex) Then
                        Reason += "依身分證號判斷 性別選項 不正確！<BR>"
                    End If
                End If

                If Birthday = "" Then
                    Reason += "必須填寫出生日期<BR>"
                Else
                    If IsDate(Birthday) = False Then
                        Reason += "出生日期必須是西元年格式(yyyy/mm/dd)<BR>"
                    Else
                        If CDate(Birthday) < "1900/1/1" Or CDate(Birthday) > "2100/1/1" Then
                            Reason += "出生日期範圍有誤<BR>"
                        End If
                    End If
                End If
                If DegreeID = "" Then
                    Reason += "必須填寫最高學歷<BR>"
                Else
                    Dim MyKey As String = DegreeID
                    If DegreeID.Length < 2 Then
                        MyKey = "0" & DegreeID
                    End If
                    If Key_Degree.Select("DegreeID='" & MyKey & "'").Length = 0 Then
                        Reason += "學歷值有錯，不符合鍵詞<BR>"
                    End If
                End If

                If school = "" Then
                    Reason += "必須填寫學校名稱<BR>"
                End If
                If Department = "" Then
                    Reason += "必須填寫科系<BR>"
                End If
                If GraduateStatus = "" Then
                    Reason += "必須填寫畢業狀況<BR>"
                Else
                    Dim MyKey As String = GraduateStatus
                    If GraduateStatus.Length < 2 Then
                        MyKey = "0" & GraduateStatus
                    End If
                    If Key_GradState.Select("GradID='" & MyKey & "'").Length = 0 Then
                        Reason += "畢業狀況有錯，不符合鍵詞<BR>"
                    End If
                End If
                'If MilitaryID = "" Then
                '    Reason += "必須填寫兵役狀況<BR>"
                'Else
                '    Dim MyKey As String = MilitaryID
                '    If MilitaryID.Length < 2 Then
                '        MyKey = "0" & MilitaryID
                '    End If
                '    If Key_Military.Select("MilitaryID='" & MyKey & "'").Length = 0 Then
                '        Reason += "兵役狀況有錯，不符合鍵詞<BR>"
                '    End If
                'End If

                If PhoneD = "" Then
                    Reason += "必須填寫聯絡電話_日<BR>"
                End If
                If ZipCode1 = "" Then
                    Reason += "必須填寫通訊地址郵遞區號<BR>"
                Else
                    If IsNumeric(ZipCode1) = False Then
                        Reason += "通訊地址郵遞區號必須要是數字<BR>"
                    End If
                End If
                If Address = "" Then
                    Reason += "必須填寫通訊地址<BR>"
                End If
                If IdentityID = "" Then
                    Reason += "必須填寫參訓身分別<BR>"
                Else
                    If IdentityID.ToString.IndexOf("，") = -1 Then
                        Dim MyKey As String = IdentityID
                        If IdentityID.Length < 2 Then
                            MyKey = "0" & IdentityID
                        Else
                            MyKey = IdentityID
                        End If
                        If Key_Identity.Select("IdentityID='" & MyKey & "'").Length = 0 Then
                            Reason += "參訓身分別不符合鍵詞<BR>"
                        End If
                    Else
                        For i As Integer = 0 To Split(IdentityID, "，").Length - 1
                            Dim MyKey As String = Split(IdentityID, "，")(i)
                            If Split(IdentityID, "，")(i).Length < 2 Then
                                MyKey = "0" & Split(IdentityID, "，")(i)
                            Else
                                MyKey = Split(IdentityID, "，")(i)
                            End If
                            If Key_Identity.Select("IdentityID='" & MyKey & "'").Length = 0 Then
                                Reason += "參訓身分別不符合鍵詞<BR>"
                            End If
                        Next
                        If Split(IdentityID, "，").Length > 3 Then
                            Reason += "參訓身分別只能選擇三種<BR>"
                        End If
                    End If
                End If
                If MIdentityID = "" Then
                    Reason += "必須填寫主要參訓身分別<BR>"
                Else
                    Dim MyKey As String = MIdentityID
                    If MIdentityID.Length < 2 Then
                        MyKey = "0" & MIdentityID
                    Else
                        MyKey = MIdentityID
                    End If
                    If Key_Identity.Select("IdentityID='" & MyKey & "'").Length = 0 Then
                        Reason += "主要參訓身分別不符合鍵詞<BR>"
                    Else
                        Dim flag As Boolean = False
                        Dim MyArray As Array = Split(MIdentityID, "，")
                        For i As Integer = 0 To MyArray.Length - 1
                            If Int(MyKey) = Int(MyArray(i)) Then
                                flag = True
                            End If
                        Next
                        If flag = False Then
                            Reason += "主要參訓身分別必須在參訓身分別的身分中<BR>"
                        End If
                    End If
                End If
                If OpenDate <> "" Then
                    If Not IsDate(OpenDate) Then
                        Reason += "開訓日期必須為正確的日期格式<BR>"
                    End If
                End If
                If CloseDate <> "" Then
                    If Not IsDate(OpenDate) Then
                        Reason += "結訓日期必須為正確的日期格式<BR>"
                    End If
                End If
                If EnterDate <> "" Then
                    If Not IsDate(EnterDate) Then
                        Reason += "報到日期必須為正確的日期格式<BR>"
                    End If
                End If

                If IsAgree = "" Then
                    Reason += "必須填寫願意是否提供個人資料給 勞動部勞動力發展署 暨所屬機關運用(Y/N)<BR>"
                Else
                    Select Case IsAgree
                        Case "Y", "N"
                        Case Else
                            Reason += "願意是否提供個人資料給 勞動部勞動力發展署 暨所屬機關運用必須為Y或N值<BR>"
                    End Select
                End If
                If AcctMode = "" Then
                    Reason += "請輸入撥款方式(0郵政,1金融)<BR>"
                Else
                    Select Case AcctMode
                        Case "0"
                            If PostNo = "" Then
                                Reason += "請輸入郵政_局號<BR>"
                            End If
                            If AcctNo = "" Then
                                Reason += "請輸入帳號<BR>"
                            End If
                        Case "1"
                            If AcctHeadNo = "" Then
                                Reason += "請輸入金融_總代號<BR>"
                            End If
                            'mark by nick 取消金融機構分支 20060414
                            ' If AcctExNo = "" Then
                            'Reason += "請輸入金融_分支代號<BR>"
                            'End If
                            'If ExBankName = "" Then
                            'Reason += "請輸入分行名稱<BR>"
                            'End If

                            If AcctNo = "" Then
                                Reason += "請輸入帳號<BR>"
                            End If
                            If BankName = "" Then
                                Reason += "請輸入銀行名稱<BR>"
                            End If

                        Case Else
                            Reason += "撥款方式超過參數範圍(0郵政,1金融)<BR>"
                    End Select
                End If
                If FirDate <> "" Then
                    If IsDate(FirDate) = False Then
                        Reason += "第一次投保勞保日必須為正確的日期格式(YYYY/MM/DD)<BR>"
                    End If
                End If
                If Tel = "" Then
                    Reason += "請輸入公司電話<BR>"
                End If
                If Zip = "" Then
                    Reason += "必須填寫公司地址郵遞區號<BR>"
                Else
                    If IsNumeric(Zip) = False Then
                        Reason += "公司地址郵遞區號必須為數字<BR>"
                    End If
                End If
                If Addr = "" Then
                    Reason += "必須填寫公司地址<BR>"
                End If
                If SDate <> "" Then
                    If IsDate(SDate) = False Then
                        Reason += "個人到任目前任職公司起日必須為正確的日期格式(YYYY/MM/DD)<BR>"
                    End If
                End If
                If SJDate <> "" Then
                    If IsDate(SJDate) = False Then
                        Reason += "個人到任目前職務起日必須為正確的日期格式(YYYY/MM/DD)<BR>"
                    End If
                End If
                If SPDate <> "" Then
                    If IsDate(SPDate) = False Then
                        Reason += "最近升遷日期必須為正確的日期格式(YYYY/MM/DD)<BR>"
                    End If
                End If
                If Q1 = "" Then
                    Reason += "是否由公司推薦參訓(Y/N值)<BR>"
                Else
                    Select Case Q1
                        Case "Y", "N"
                        Case Else
                            Reason += "是否由公司推薦參訓必須為Y/N值<BR>"
                    End Select
                End If
                If Q2 = "" Then
                    Reason += "必須填寫參訓動機(1~4)<BR>"
                Else
                    If Q2.IndexOf("，") = -1 Then
                        If Not IsNumeric(Q2) Then
                            Reason += "參訓動機必須為數字(1~4)<BR>"
                        Else
                            If Int(Q2) > 4 Or Int(Q2) < 1 Then
                                Reason += "參訓動機範圍1~4<BR>"
                            End If
                        End If
                    Else
                        For i As Integer = 0 To Split(Q2, "，").Length - 1
                            If Not IsNumeric(Split(Q2, "，")(i)) Then
                                Reason += "參訓動機必須為數字(1~4)<BR>"
                            Else
                                If Int(Split(Q2, "，")(i)) > 4 Or Int(Split(Q2, "，")(i)) < 1 Then
                                    Reason += "參訓動機範圍1~4<BR>"
                                End If
                            End If
                        Next
                    End If
                End If
                If Q3 <> "" Then
                    If Not IsNumeric(Q3) Then
                        Reason += "訓後動向必須為數字(1~3)"
                    Else
                        If Int(Q3) > 3 Or Int(Q3) < 1 Then
                            Reason += "訓後動向範圍1~3<BR>"
                        End If
                    End If
                End If
                If Q4 = "" Then
                    Reason += "必須填寫服務單位行業別<BR>"
                Else
                    If Not IsNumeric(Q4) Then
                        Reason += "服務單位行業別必須為數字(01~31)"
                    Else
                        If Int(Q4) > 31 Or Int(Q4) < 1 Then
                            Reason += "服務單位行業別範圍01~31<BR>"
                        End If
                    End If
                End If
                If Q5 <> "" Then
                    Select Case Q5
                        Case "Y", "N"
                        Case Else
                            Reason += "服務單位是否屬於中小企業只能輸入Y或N<BR>"
                    End Select
                End If
                If Q61 <> "" Then
                    If Not IsNumeric(Q61) Or Q61.IndexOf(".") <> -1 Then
                        Reason += "個人工作年資必須為數字<BR>"
                    End If
                End If
                If Q62 <> "" Then
                    If Not IsNumeric(Q62) Or Q62.IndexOf(".") <> -1 Then
                        Reason += "在這家公司的年資必須為數字<BR>"
                    End If
                End If
                If Q63 <> "" Then
                    If Not IsNumeric(Q63) Or Q63.IndexOf(".") <> -1 Then
                        Reason += "在這職位的年資必須為數字<BR>"
                    End If
                End If
                If Q64 <> "" Then
                    If Not IsNumeric(Q64) Or Q64.IndexOf(".") <> -1 Then
                        Reason += "最近升遷離本職幾年必須為數字<BR>"
                    End If
                End If
                If ShowDetail = "" Then
                    Reason += "必須填寫是否提供基本資料查詢<BR>"
                Else
                    Select Case ShowDetail
                        Case "Y", "N"
                        Case Else
                            Reason += "是否提供基本資料查詢必須為Y或N值<BR>"
                    End Select
                End If
            End If
        Else
            '產業人才投資方案
            If colArray.Length <> 76 Then
                'Reason += "欄位數量不正確(應該為76個欄位)<BR>"
                Reason += "欄位對應有誤<BR>"
                Reason += "請注意欄位中是否有半形逗點<BR>"
            Else
                If colArray(0).ToString = "" Then
                    Reason += "必須填寫學號<Br>"
                Else
                    If IsNumeric(colArray(0)) Then
                        Dim MyKey As String = Int(colArray(0))
                        If Int(MyKey) >= 100 Then
                            Reason += "學號必須為在(01~99)範圍內<BR>"
                        Else
                            If Int(MyKey) < 10 Then
                                MyKey = "0" & Int(MyKey)
                            End If

                            sql = "SELECT * FROM Class_StudentsOfClass WHERE OCID='" & OCIDValue1.Value & "' and  StudentID='" & StudentIDBasic & MyKey & "'"
                            dr = DbAccess.GetOneRow(sql, objconn)
                            If Not dr Is Nothing Then
                                Reason += "學號重複<BR>"
                            End If
                        End If
                    Else
                        Reason += "學號必須為數字(01~99)<BR>"
                    End If
                End If

                If colArray(1).ToString = "" Then
                    Reason += "必須填寫中文姓名<BR>"
                End If
                If colArray(2).ToString = "" Then
                    Reason += "必須填寫英文姓名(LastName)<BR>"
                Else
                    For i As Integer = 0 To colArray(2).ToString.Length - 1
                        If SearchEngStr.IndexOf(colArray(2).ToString.ToUpper.Chars(i)) = -1 Then
                            Reason += "英文姓名必須只有英文字(LastName)<BR>"
                        End If
                    Next
                End If
                If colArray(3).ToString = "" Then
                    Reason += "必須填寫英文姓名(FirstName)<BR>"
                Else
                    For i As Integer = 0 To colArray(3).ToString.Length - 1
                        If SearchEngStr.IndexOf(colArray(3).ToString.ToUpper.Chars(i)) = -1 Then
                            Reason += "英文姓名必須只有英文字(FirstName)<BR>"
                        End If
                    Next
                End If
                If colArray(4).ToString = "" Then
                    Reason += "必須填寫身分證號碼<BR>"
                Else
                    Dim IDNO As String = colArray(4).ToString
                    If colArray(6).ToString = "1" Then
                        If sm.UserInfo.RoleID <> 5 Then
                            Dim IDNOFlag As Boolean = True
                            Dim EngStr As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
                            If IDNO.Length <> 10 Then
                                IDNOFlag = False
                            ElseIf IDNO.Chars(1) <> "1" And IDNO.Chars(1) <> "2" Then
                                IDNOFlag = False
                            ElseIf EngStr.IndexOf(IDNO.ToUpper.Chars(0)) = -1 Then
                                IDNOFlag = False
                            ElseIf IDNO = "A123456789" Then
                                IDNOFlag = False
                            End If

                            If IDNOFlag = False Then
                                Reason += "身分證號碼錯誤!<BR>"
                            End If
                        Else
                            If TIMS.CheckIDNO(IDNO) = False Then
                                Reason += "身分證號碼錯誤!請聯絡系統管理員<BR>"
                            End If
                        End If
                    End If

                    sql = "SELECT * FROM Stud_StudentInfo WHERE IDNO='" & IDNO & "' and SID IN (SELECT SID FROM Class_StudentsOfClass WHERE OCID='" & OCIDValue1.Value & "')"
                    dr = DbAccess.GetOneRow(sql, objconn)
                    If Not dr Is Nothing Then
                        Reason += "此班已經有相同的身分證號碼<BR>"
                    Else
                        Dim Flag As Boolean = True
                        For i As Integer = 0 To IDNOArray.Count - 1
                            If IDNOArray(i) = IDNO Then
                                Reason += "檔案中有相同的身分證號碼<BR>"
                                Flag = False
                            End If
                        Next
                        If Flag Then
                            IDNOArray.Add(IDNO)
                        End If
                    End If
                End If
                If colArray(5).ToString = "" Then
                    Reason += "必須填寫性別<BR>"
                Else
                    Select Case colArray(5).ToString
                        Case "M", "F"
                        Case Else
                            Reason += "性別代號只能是M或者是F<BR>"
                    End Select
                End If

                If Reason = "" Then
                    If Not TIMS.checkMemberSex(colArray(4).ToString, colArray(5).ToString) Then
                        Reason += "依身分證號判斷 性別選項 不正確！<BR>"
                    End If
                End If

                If colArray(6).ToString = "" Then
                    Reason += "必須填寫身分別(1~2)<BR>"
                Else
                    Select Case colArray(6).ToString
                        Case "1"
                        Case "2"
                            If colArray(7).ToString = "" Then
                                Reason += "請輸入非本國人身分別"
                            Else
                                Select Case colArray(7).ToString
                                    Case "1"
                                    Case "2"
                                    Case Else
                                        Reason += "非本國人身分別只能輸入1(大陸人士)或2(非大陸人士)"
                                End Select
                            End If
                            If colArray(8).ToString = "" Then
                                Reason += "請輸入原屬國籍"
                            End If
                            If colArray(9).ToString = "" Then
                                Reason += "請輸入護照或工作證號"
                            Else
                                Select Case colArray(9).ToString
                                    Case "1"
                                    Case "2"
                                    Case Else
                                        Reason += "非本國人身分別只能輸入1(護照號碼)或2(工作證號)"
                                End Select
                            End If
                        Case Else
                            Reason += "身分別只能輸入1(本國)或2(外籍)<BR>"
                    End Select
                End If
                If colArray(10).ToString = "" Then
                    Reason += "必須填寫出生日期<BR>"
                Else
                    If IsDate(colArray(10)) = False Then
                        Reason += "出生日期必須是西元年格式(yyyy/mm/dd)<BR>"
                    Else
                        If CDate(colArray(10)) < "1900/1/1" Or CDate(colArray(10)) > "2100/1/1" Then
                            Reason += "出生日期範圍有誤<BR>"
                        End If
                    End If
                End If
                If colArray(11).ToString <> "" Then
                    Select Case colArray(11).ToString
                        Case "1", "2"
                        Case Else
                            Reason += "婚姻狀況必須是1(已婚)或2(未婚)<BR>"
                    End Select
                End If
                If colArray(12).ToString = "" Then
                    Reason += "必須填寫最高學歷<BR>"
                Else
                    Dim MyKey As String = colArray(12)
                    If colArray(12).ToString.Length < 2 Then
                        MyKey = "0" & colArray(12)
                    End If
                    If Key_Degree.Select("DegreeID='" & MyKey & "'").Length = 0 Then
                        Reason += "學歷值有錯，不符合鍵詞<BR>"
                    End If
                End If
                If colArray(13).ToString = "" Then
                    Reason += "必須填寫學校名稱<BR>"
                End If
                If colArray(14).ToString = "" Then
                    Reason += "必須填寫科系<BR>"
                End If
                If colArray(15).ToString = "" Then
                    Reason += "必須填寫畢業狀況<BR>"
                Else
                    Dim MyKey As String = colArray(15)
                    If colArray(15).ToString.Length < 2 Then
                        MyKey = "0" & colArray(15)
                    End If
                    If Key_GradState.Select("GradID='" & MyKey & "'").Length = 0 Then
                        Reason += "畢業狀況有錯，不符合鍵詞<BR>"
                    End If
                End If
                If colArray(16).ToString = "" Then
                    Reason += "必須填寫兵役狀況<BR>"
                Else
                    Dim MyKey As String = colArray(16)
                    If colArray(16).ToString.Length < 2 Then
                        MyKey = "0" & colArray(16)
                    End If
                    If Key_Military.Select("MilitaryID='" & MyKey & "'").Length = 0 Then
                        Reason += "兵役狀況有錯，不符合鍵詞<BR>"
                    Else
                        If Int(colArray(16)) = "4" Then
                            If colArray(17).ToString = "" Then
                                Reason += "必須填寫軍種<BR>"
                            End If
                            If colArray(19).ToString = "" Then
                                Reason += "必須填寫階級<BR>"
                            End If
                            If colArray(20).ToString = "" Then
                                Reason += "必須填寫服務單位名稱<BR>"
                            End If
                            If colArray(22).ToString = "" Then
                                Reason += "必須填寫單位電話<BR>"
                            End If
                            If colArray(23).ToString = "" Then
                                Reason += "必須填寫服役起日期<BR>"
                            Else
                                If IsDate(colArray(23)) = False Then
                                    Reason += "服役起日期不是正確的日期格式<BR>"
                                Else
                                    If CDate(colArray(23)) < "1900/1/1" Or CDate(colArray(23)) > "2100/1/1" Then
                                        Reason += "服役起日期範圍有誤<BR>"
                                    End If
                                End If
                            End If
                            If colArray(24).ToString = "" Then
                                Reason += "必須填寫服役迄日期<BR>"
                            Else
                                If IsDate(colArray(24)) = False Then
                                    Reason += "服役迄日期不是正確的日期格式<BR>"
                                Else
                                    If CDate(colArray(24)) < "1900/1/1" Or CDate(colArray(24)) > "2100/1/1" Then
                                        Reason += "服役迄日期範圍有誤<BR>"
                                    End If
                                End If
                            End If
                            If colArray(25).ToString <> "" Then
                                If IsNumeric(colArray(25)) = False Then
                                    Reason += "服役單位地址郵遞區號必須為數字<BR>"
                                End If
                            End If
                        End If
                    End If
                End If

                If colArray(27).ToString = "" Then
                    Reason += "必須填寫聯絡電話_日<BR>"
                End If
                If colArray(30).ToString = "" Then
                    Reason += "必須填寫通訊地址郵遞區號<BR>"
                Else
                    If IsNumeric(colArray(30)) = False Then
                        Reason += "通訊地址郵遞區號必須要是數字<BR>"
                    End If
                End If
                If colArray(31).ToString = "" Then
                    Reason += "必須填寫通訊地址<BR>"
                End If
                If colArray(32).ToString <> "" Then
                    If IsNumeric(colArray(32)) = False Then
                        Reason += "戶籍地址郵遞區號必須要是數字<BR>"
                    End If
                End If
                If colArray(35).ToString = "" Then
                    Reason += "必須填寫參訓身分別<BR>"
                Else
                    Dim MyKey As String = colArray(35)
                    If colArray(35).ToString.Length < 2 Then
                        MyKey = "0" & colArray(35)
                    Else
                        MyKey = colArray(35)
                    End If
                    If colArray(35).ToString.IndexOf("，") = -1 Then
                        If Key_Identity.Select("IdentityID='" & MyKey & "'").Length = 0 Then
                            Reason += "參訓身分別不符合鍵詞<BR>"
                        End If
                    Else
                        For i As Integer = 0 To Split(colArray(35), "，").Length - 1
                            If Split(colArray(35), "，")(i).Length < 2 Then
                                MyKey = "0" & Split(colArray(35), "，")(i)
                            Else
                                MyKey = Split(colArray(35), "，")(i)
                            End If
                            If Key_Identity.Select("IdentityID='" & MyKey & "'").Length = 0 Then
                                Reason += "參訓身分別不符合鍵詞<BR>"
                            End If
                        Next
                        If Split(colArray(35), "，").Length > 3 Then
                            Reason += "參訓身分別只能選擇三種<BR>"
                        End If
                    End If
                End If
                If colArray(36).ToString = "" Then
                    Reason += "必須填寫主要參訓身分別<BR>"
                Else
                    Dim MyKey As String = colArray(36)
                    If colArray(36).ToString.Length < 2 Then
                        MyKey = "0" & colArray(36)
                    Else
                        MyKey = colArray(36)
                    End If
                    If Key_Identity.Select("IdentityID='" & MyKey & "'").Length = 0 Then
                        Reason += "主要參訓身分別不符合鍵詞<BR>"
                    Else
                        Dim flag As Boolean = False
                        Dim MyArray As Array = Split(colArray(35), "，")
                        For i As Integer = 0 To MyArray.Length - 1
                            If Int(MyKey) = Int(MyArray(i)) Then
                                flag = True
                            End If
                        Next
                        If flag = False Then
                            Reason += "主要參訓身分別必須在參訓身分別的身分中<BR>"
                        End If
                    End If
                    'by Vicient
                    If MyKey = "05" Or MyKey = "5" Then
                        If colArray(75).ToString = "" Then
                            Reason += "必須填寫原住民別<BR>"
                        Else
                            Dim MyKey1 As String = colArray(75)
                            Dim Key_Native As DataTable

                            sql = "select * from Key_Native where KNID = " & MyKey1
                            Key_Native = DbAccess.GetDataTable(sql, objconn)
                            If Key_Native.Rows.Count = 0 Then
                                Reason += "民族別有錯，不符合鍵詞<BR>"
                            End If
                        End If
                    End If

                End If
                If colArray(37).ToString = "" Then
                    Reason += "必須填寫生活津貼代碼<BR>"
                Else
                    Dim MyKey As String = colArray(37)
                    If colArray(37).ToString.Length < 2 Then
                        MyKey = "0" & colArray(37)
                    End If
                    If Key_Subsidy.Select("SubsidyID='" & MyKey & "'").Length = 0 Then
                        Reason += "生活津貼不符合鍵詞<BR>"
                    End If
                End If
                If colArray(38).ToString <> "" Then
                    If IsDate(colArray(38)) = False Then
                        Reason += "開訓日期不符合日期格式<BR>"
                    Else
                        If CDate(colArray(38)) < "1900/1/1" Or CDate(colArray(38)) > "2100/1/1" Then
                            Reason += "開訓日期範圍有誤<BR>"
                        End If
                    End If
                End If
                If colArray(39).ToString <> "" Then
                    If IsDate(colArray(39)) = False Then
                        Reason += "結訓日期不符合日期格式<BR>"
                    Else
                        If CDate(colArray(39)) < "1900/1/1" Or CDate(colArray(39)) > "2100/1/1" Then
                            Reason += "結訓日期範圍有誤<BR>"
                        End If
                    End If
                End If
                If colArray(40).ToString <> "" Then
                    If IsDate(colArray(40)) = False Then
                        Reason += "報到日期不符合日期格式<BR>"
                    Else
                        If CDate(colArray(40)) < "1900/1/1" Or CDate(colArray(40)) > "2100/1/1" Then
                            Reason += "報到日期範圍有誤<BR>"
                        End If
                    End If
                End If
                If colArray(41).ToString <> "" Then
                    Dim MyKey As String = colArray(41)
                    If colArray(41).ToString.Length < 2 Then
                        MyKey = "0" & colArray(41)
                    End If
                    If Key_HandicatType.Select("HandTypeID='" & MyKey & "'").Length = 0 Then
                        Reason += "障礙類別有錯，不符合鍵詞<BR>"
                    End If
                End If
                If colArray(42).ToString <> "" Then
                    Dim MyKey As String = colArray(42)
                    If colArray(42).ToString.Length < 2 Then
                        MyKey = "0" & colArray(42)
                    End If
                    If Key_HandicatLevel.Select("HandLevelID='" & MyKey & "'").Length = 0 Then
                        Reason += "障礙等級有錯，不符合鍵詞<BR>"
                    End If
                End If
                If colArray(43).ToString = "" Then
                    Reason += "必須填寫緊急通知人姓名<BR>"
                End If
                If colArray(44).ToString = "" Then
                    Reason += "必須填寫緊急通知人關係<BR>"
                End If
                If colArray(45).ToString = "" Then
                    Reason += "必須填寫緊急通知人電話<BR>"
                End If
                If colArray(46).ToString = "" Then
                    Reason += "必須填寫緊急通知人地址郵遞區號<BR>"
                Else
                    If IsNumeric(colArray(46)) = False Then
                        Reason += "郵遞區號必須為數字<BR>"
                    End If
                End If
                If colArray(47).ToString = "" Then
                    Reason += "必須填寫緊急通知人地址<BR>"
                End If
                If colArray(50).ToString <> "" Then
                    If IsDate(colArray(50)) = False Then
                        Reason += "受訓前服務單位1任職起日不符合日期格式<BR>"
                    Else
                        If CDate(colArray(50)) < "1900/1/1" Or CDate(colArray(50)) > "2100/1/1" Then
                            Reason += "受訓前服務單位1任職起日範圍有誤<BR>"
                        End If
                    End If
                End If
                If colArray(51).ToString <> "" Then
                    If IsDate(colArray(51)) = False Then
                        Reason += "受訓前服務單位1任職迄日不符合日期格式<BR>"
                    Else
                        If CDate(colArray(51)) < "1900/1/1" Or CDate(colArray(51)) > "2100/1/1" Then
                            Reason += "受訓前服務單位1任職迄日範圍有誤<BR>"
                        End If
                    End If
                End If
                If colArray(54).ToString <> "" Then
                    If IsDate(colArray(54)) = False Then
                        Reason += "受訓前服務單位2任職起日不符合日期格式<BR>"
                    Else
                        If CDate(colArray(54)) < "1900/1/1" Or CDate(colArray(54)) > "2100/1/1" Then
                            Reason += "受訓前服務單位2任職起日範圍有誤<BR>"
                        End If
                    End If
                End If
                If colArray(55).ToString <> "" Then
                    If IsDate(colArray(55)) = False Then
                        Reason += "受訓前服務單位2任職迄日不符合日期格式<BR>"
                    Else
                        If CDate(colArray(55)) < "1900/1/1" Or CDate(colArray(55)) > "2100/1/1" Then
                            Reason += "受訓前服務單位2任職迄日範圍有誤<BR>"
                        End If
                    End If
                End If
                If colArray(56).ToString <> "" Then
                    If IsNumeric(colArray(56)) = False Then
                        Reason += "受訓前薪資必須為數字<BR>"
                    End If
                End If
                If colArray(57).ToString <> "" Then
                    If IsNumeric(colArray(57)) = False Then
                        Reason += "受訓前真正失業週數必須為數字<BR>"
                    End If
                End If
                If colArray(58).ToString = "" Then
                    Reason += "必須填寫失業週數代碼<BR>"
                Else
                    Dim MyKey As String = colArray(58)
                    If colArray(58).ToString.Length < 2 Then
                        MyKey = "0" & colArray(58)
                    End If
                    If Key_JoblessWeek.Select("JoblessID='" & MyKey & "'").Length = 0 Then
                        Reason += "失業週數代碼有錯，不符合鍵詞<BR>"
                    End If
                End If
                If colArray(59).ToString <> "" Then
                    Select Case colArray(59).ToString
                        Case "1", "2"
                        Case Else
                            Reason += "交通方式必須為1(住宿)或2(通勤)<BR>"
                    End Select
                End If
                If colArray(60).ToString = "" Then
                    Reason += "必須填寫是否提供基本資料查詢<BR>"
                Else
                    Select Case colArray(60).ToString
                        Case "Y", "y", "N", "n"
                        Case Else
                            Reason += "是否提供基本資料查詢必須為Y或N值<BR>"
                    End Select
                End If
                If colArray(62).ToString <> "" Then
                    Select Case colArray(62).ToString
                        Case "1", "2", "3"
                        Case "4"
                            If colArray(63).ToString = "" Then
                                Reason += "報名管道為推介時，必須選擇卷別<BR>"
                            Else
                                Select Case colArray(63).ToString
                                    Case "1", "3"
                                        If colArray(64).ToString = "" Then
                                            Reason += "券別種類必須填入甲乙式<BR>"
                                        Else
                                            Select Case colArray(64).ToString
                                                Case "1", "2"
                                                Case Else
                                                    Reason += "券別種類只有1(甲式)2(乙式)<BR>"
                                            End Select
                                        End If
                                    Case "2"
                                        If colArray(64).ToString <> "" Then
                                            Reason += "學習券不區分甲乙式<BR>"
                                        End If
                                    Case Else
                                        Reason += "推介種類只有1(職訓券)2(學習券)3(推介券)<BR>"
                                End Select
                            End If
                        Case Else
                            Reason += "報名管道只有1(網路)2(現場)3(通訊)4(推介)<BR>"
                    End Select
                End If
                If colArray(65).ToString = "" And Plan_Budget.Rows.Count <> 0 Then
                    Reason += "必須填寫預算別<BR>"
                Else
                    Dim MyKey As String = colArray(65).ToString
                    If MyKey.Length < 2 Then
                        MyKey = "0" & MyKey
                    End If
                    If Plan_Budget.Select("BudID='" & MyKey & "'").Length = 0 Then
                        Reason += "預算別不符合此訓練計畫<BR>"
                    End If
                End If
                If colArray(66).ToString = "" Then
                    Reason += "必須填寫願意是否提供個人資料給 勞動部勞動力發展署 暨所屬機關運用(Y/N)<BR>"
                Else
                    Select Case colArray(66).ToString
                        Case "Y", "N"
                        Case Else
                            Reason += "願意是否提供個人資料給 勞動部勞動力發展署 暨所屬機關運用必須為Y或N值<BR>"
                    End Select
                End If
                If colArray(67).ToString = "" Then
                    If sm.UserInfo.TPlanID = "12" Then
                        Reason += "必須填寫公費(1)或自費(2)<BR>"
                    End If
                ElseIf colArray(67).ToString <> "" Then
                    Select Case colArray(67).ToString
                        Case "1", "2"
                        Case Else
                            Reason += "公費(1)/自費(2)值超出範圍<BR>"
                    End Select
                End If
                If colArray(70).ToString <> "" Then
                    Select Case colArray(70).ToString
                        Case "M", "F"
                        Case Else
                            Reason += "國內親屬資料_性別只能輸入M(男性)F(女性)<BR>"
                    End Select
                End If
                If colArray(71).ToString <> "" Then
                    If Not IsDate(colArray(71).ToString) Then
                        Reason += "國內親屬資料_生日必須為正確的日期格式<BR>"
                    End If
                End If
                If colArray(72).ToString <> "" Then
                    If Not TIMS.CheckIDNO(colArray(72).ToString) Then
                        Reason += "國內親屬資料_身分證號碼不是正確的身分證號碼<BR>"
                    End If
                End If
                If colArray(73).ToString <> "" Then
                    If Not IsNumeric(colArray(73)) Then
                        Reason += "國內親屬資料_郵遞區號必須為數字<BR>"
                    End If
                End If
            End If
        End If

        Return Reason
    End Function

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound


        '"a.StudentID,b.Name,b.IDNO,b.Sex,b.Birthday"
        '",a.levelNo,b.EngName,b.DegreeID,c.school,c.Department,b.MilitaryID,c.ServiceID,c.MilitaryRank,
        'c.ServiceOrg,c.ServicePhone,c.SServiceDate,c.FServiceDate,c.PhoneD,c.PhoneN,c.address,a.MidentityID,
        'a.IdentityID,c.EmergencyContact,c.EmergencyRelation,c.EmergencyAddress,c.ZipCode3,c.ShowDetail,a.budgetID,b.IsAgree "


        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then
            Dim drv As DataRowView = e.Item.DataItem
            objControl = e.Item.FindControl("Checkbox1")
            Dim btn1 As LinkButton = e.Item.FindControl("Button3")
            Dim btn2 As LinkButton = e.Item.FindControl("Button9b")
            Dim Checkbox2 As HtmlInputCheckBox = e.Item.FindControl("Checkbox2")

            '加入判斷是否有資料未填 by nick
            Dim star As Label = e.Item.FindControl("star1")

            If Not IsDBNull(drv("SubsidyID")) Then
                If CStr(drv("SubsidyID")).Trim = "" Then
                    star.Visible = True
                End If
            Else
                star.Visible = True
            End If
            If Not IsDBNull(drv("MIdentityID")) Then
                If CStr(drv("MIdentityID")).Trim = "" Then
                    star.Visible = True
                End If
            Else
                star.Visible = True
            End If
            If Not IsDBNull(drv("school")) Then
                If CStr(drv("school")).Trim = "" Then
                    star.Visible = True
                End If
            Else
                star.Visible = True
            End If
            If Not IsDBNull(drv("Department")) Then
                If CStr(drv("Department")).Trim = "" Then
                    star.Visible = True
                End If
            Else
                star.Visible = True
            End If
            '兵役狀況如果為在役中04 則要填軍種階級等
            If Not IsDBNull(drv("MilitaryID")) Then
                If CStr(drv("MilitaryID")).Trim = "" Then
                    star.Visible = True
                ElseIf CStr(drv("MilitaryID")).Trim = "04" Then
                    If Not IsDBNull(drv("ServiceID")) Then
                        If CStr(drv("ServiceID")).Trim = "" Then
                            star.Visible = True
                        End If
                    Else
                        star.Visible = True
                    End If
                    If Not IsDBNull(drv("ServiceOrg")) Then
                        If CStr(drv("ServiceOrg")).Trim = "" Then
                            star.Visible = True
                        End If
                    Else
                        star.Visible = True
                    End If
                    If Not IsDBNull(drv("ServicePhone")) Then
                        If CStr(drv("ServicePhone")).Trim = "" Then
                            star.Visible = True
                        End If
                    Else
                        star.Visible = True
                    End If
                    If Not IsDBNull(drv("MilitaryRank")) Then
                        If CStr(drv("MilitaryRank")).Trim = "" Then
                            star.Visible = True
                        End If
                    Else
                        star.Visible = True
                    End If

                End If
            Else
                star.Visible = True
            End If


            If Not IsDBNull(drv("PhoneD")) Then
                If CStr(drv("PhoneD")).Trim = "" Then
                    star.Visible = True
                End If
            Else
                star.Visible = True
            End If
            If Not IsDBNull(drv("address")) Then
                If CStr(drv("address")).Trim = "" Then
                    star.Visible = True
                End If
            Else
                star.Visible = True
            End If


            If Not IsDBNull(drv("IdentityID")) Then
                If CStr(drv("IdentityID")).Trim = "" Then
                    star.Visible = True
                End If
            Else
                star.Visible = True
            End If
            '如果是產學訓則不擋
            If Not TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                If Not IsDBNull(drv("EngName")) Then
                    If CStr(drv("EngName")).Trim = "" Then
                        star.Visible = True
                    End If
                Else
                    star.Visible = True
                End If
                If Not IsDBNull(drv("EmergencyContact")) Then
                    If CStr(drv("EmergencyContact")).Trim = "" Then
                        star.Visible = True
                    End If
                Else
                    star.Visible = True
                End If
                If Not IsDBNull(drv("EmergencyRelation")) Then
                    If CStr(drv("EmergencyRelation")).Trim = "" Then
                        star.Visible = True
                    End If
                Else
                    star.Visible = True
                End If
                If Not IsDBNull(drv("ZipCode3")) Then
                    If CStr(drv("ZipCode3")).Trim = "" Then
                        star.Visible = True
                    End If
                Else
                    star.Visible = True
                End If
                If Not IsDBNull(drv("budgetID")) Then
                    If CStr(drv("budgetID")).Trim = "" Then
                        star.Visible = True
                    End If
                Else
                    star.Visible = True
                End If
            End If

            If Not IsDBNull(drv("ShowDetail")) Then
                If CStr(drv("ShowDetail")).Trim = "" Then
                    star.Visible = True
                End If
            Else
                star.Visible = True
            End If

            If Not IsDBNull(drv("IsAgree")) Then
                If CStr(drv("IsAgree")).Trim = "" Then
                    star.Visible = True
                End If
            Else
                star.Visible = True
            End If


            'end 加入判斷是否有資料未填
            Checkbox2.Value = drv("StudentID")
            Checkbox2.Attributes("onclick") = "InsertValue(this.checked,this.value)"
            If PrintValue.Value.IndexOf(drv("StudentID")) <> -1 Then
                Checkbox2.Checked = True
            End If
            e.Item.Cells(1).Text = Right(e.Item.Cells(1).Text, 2)
            If drv("Sex") = "M" Then
                e.Item.Cells(4).Text = "男"
            ElseIf drv("Sex") = "F" Then
                e.Item.Cells(4).Text = "女"
            End If

            Select Case drv("StudStatus").ToString
                Case "1"
                    e.Item.Cells(6).Text = "在訓"
                Case "2"
                    e.Item.Cells(6).Text = "離訓"
                Case "3"
                    e.Item.Cells(6).Text = "退訓"
                Case "4"
                    e.Item.Cells(6).Text = "續訓"
                Case "5"
                    e.Item.Cells(6).Text = "結訓"
            End Select

            'btn.CommandArgument = "SID=" & drv("SID") & "&SOCID=" & drv("SOCID") & ""
            btn1.CommandArgument = drv("SOCID")
            btn2.CommandArgument = drv("SOCID")

            If FunDr("Mod") = 1 Then
                btn1.Enabled = True
            Else
                btn1.Enabled = False
            End If

            If sm.UserInfo.RoleID = 1 Or sm.UserInfo.RoleID = 5 Then
                btn2.Visible = True
                btn2.Attributes("onclick") = "return confirm('這樣會刪除此學員的相關班級資料,\n但不會刪除此學員的個人基本資料,\n確定要繼續刪除?');"
            Else
                btn2.Visible = False
            End If
        End If
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Select Case e.CommandName
            Case "edit"
                Session("SearchSOCID") = e.CommandArgument
                GetSearchStr()
                'Response.Redirect("SD_03_006_add.aspx?ID=" & Request("ID") & "&OCID=" & OCIDValue1.Value)
            Case "del"
                Dim sql As String = ""
                Dim MsgBox As String = ""
                Dim dr As DataRow
                '津貼
                sql = "SELECT * FROM Stud_SubsidyResult WHERE SOCID='" & e.CommandArgument & "'"
                dr = DbAccess.GetOneRow(sql, objconn)
                If Not dr Is Nothing Then
                    MsgBox += "此學員現在有津貼資料，不能刪除" & vbCrLf
                End If
                '技能檢定
                sql = "SELECT * FROM Stud_TechExam WHERE SOCID='" & e.CommandArgument & "'"
                dr = DbAccess.GetOneRow(sql, objconn)
                If Not dr Is Nothing Then
                    MsgBox += "此學員現在有技能檢定資料，不能刪除" & vbCrLf
                End If
                '結訓成績
                sql = "SELECT * FROM Stud_TrainingResults WHERE SOCID='" & e.CommandArgument & "'"
                dr = DbAccess.GetOneRow(sql, objconn)
                If Not dr Is Nothing Then
                    MsgBox += "此學員現在有結訓成績資料，不能刪除" & vbCrLf
                End If
                '操行
                sql = "SELECT * FROM Stud_Conduct WHERE SOCID='" & e.CommandArgument & "'"
                dr = DbAccess.GetOneRow(sql, objconn)
                If Not dr Is Nothing Then
                    MsgBox += "此學員現在有操行成績資料，不能刪除" & vbCrLf
                End If
                '轉班
                sql = "SELECT * FROM Stud_TranClassRecord WHERE SOCID='" & e.CommandArgument & "'"
                dr = DbAccess.GetOneRow(sql, objconn)
                If Not dr Is Nothing Then
                    MsgBox += "此學員現在有轉班資料，不能刪除" & vbCrLf
                End If
                '出缺勤
                sql = "SELECT * FROM Stud_Turnout WHERE SOCID='" & e.CommandArgument & "'"
                dr = DbAccess.GetOneRow(sql, objconn)
                If Not dr Is Nothing Then
                    MsgBox += "此學員現在有出缺勤資料，不能刪除" & vbCrLf
                End If
                '獎懲
                sql = "SELECT * FROM Stud_Sanction WHERE SOCID='" & e.CommandArgument & "'"
                dr = DbAccess.GetOneRow(sql, objconn)
                If Not dr Is Nothing Then
                    MsgBox += "此學員現在有獎懲資料，不能刪除" & vbCrLf
                End If
                '結訓學員資料卡
                sql = "SELECT * FROM Stud_ResultStudData WHERE SOCID='" & e.CommandArgument & "'"
                dr = DbAccess.GetOneRow(sql, objconn)
                If Not dr Is Nothing Then
                    MsgBox += "此學員現在有填寫結訓學員資料卡，不能刪除" & vbCrLf
                End If

                If MsgBox = "" Then
                    Page.RegisterStartupScript("del", "<script>wopen('SD_03_006_del.aspx?ID=" & Request("ID") & "&SOCID=" & e.CommandArgument & "','del',350,250,0)</script>")
                Else
                    Common.MessageBox(Me, MsgBox)
                End If
        End Select

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        GetSearchStr()
        If sm.UserInfo.TPlanID = 15 Then
            'Response.Redirect("SD_03_006_3in1.aspx?ID=" & Request("ID") & "&OCID=" & OCIDValue1.Value & "")
        Else
            'Response.Redirect("SD_03_006_add.aspx?ID=" & Request("ID") & "&OCID=" & OCIDValue1.Value)
        End If
    End Sub

    Sub GetSearchStr()
        Session("_SearchStr") = "center=" & center.Text & "&"
        Session("_SearchStr") += "RIDValue=" & RIDValue.Value & "&"
        Session("_SearchStr") += "TMID1=" & TMID1.Text & "&"
        Session("_SearchStr") += "TMIDValue1=" & TMIDValue1.Value & "&"
        Session("_SearchStr") += "OCID1=" & OCID1.Text & "&"
        Session("_SearchStr") += "OCIDValue1=" & OCIDValue1.Value & "&"
        Session("_SearchStr") += "PageIndex=" & DataGrid1.CurrentPageIndex + 1 & "&"
        If DataGrid1.Visible Then
            Session("_SearchStr") += "submit=1"
        Else
            Session("_SearchStr") += "submit=0"
        End If
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Dim sql As String
        Dim dt As DataTable
        Dim IDNOArray As New ArrayList

        sql = Me.ViewState("SD03002_SearchSqlStr")
        dt = DbAccess.GetDataTable(sql, objconn)
        For Each dr As DataRow In dt.Rows
            IDNOArray.Add(dr("IDNO").ToString)
        Next

        Session("IDNOArray") = IDNOArray
        Page.RegisterStartupScript("History", "<script>window.open('../../SD/01/SD_01_001_old.aspx','history','width=700,height=500,scrollbars=1')</script>")
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Dim dr As DataRow
        '判斷機構是否只有一個班級
        dr = TIMS.GET_OnlyOne_OCID(RIDValue.Value)
        If Not dr Is Nothing Then
            If dr("total") = "1" Then '如果只有一個班級
                TMID1.Text = dr("trainname")
                OCID1.Text = dr("classname")
                TMIDValue1.Value = dr("trainid")
                OCIDValue1.Value = dr("ocid")
            Else '不只一個班級
                TMID1.Text = ""
                OCID1.Text = ""
                TMIDValue1.Value = ""
                OCIDValue1.Value = ""
            End If
        End If
    End Sub

End Class
