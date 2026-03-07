Partial Class SD_03_006_add
    Inherits System.Web.UI.Page

    'Dim SearchPage As SD_03_006
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

        SOCID.Attributes("onchange") = "if(this.selectedIndex==0) return false;"
        Page.RegisterStartupScript("ZipCcript1", TIMS.Get_ZipNameJScript(objconn))
        IDNO.Attributes("onblur") = "if(this.value.length==10){document.getElementById('IDNO').value=this.value.toUpperCase();}"
        City1.Attributes("onblur") = "getzipname(this.value,'City1','ZipCode1');"
        City2.Attributes("onblur") = "getzipname(this.value,'City2','ZipCode2');"
        City3.Attributes("onblur") = "getzipname(this.value,'City3','ZipCode3');"
        City4.Attributes("onblur") = "getzipname(this.value,'City4','ZipCode4');"
        City5.Attributes("onblur") = "getzipname(this.value,'City5','Zip');"
        City6.Attributes("onblur") = "getzipname(this.value,'City6','ForeZip');"
        If Not TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            MIdentityID.Attributes("onchange") = "checkNativeID()"
        End If
        Button1.Attributes("onclick") = "return chkdata();"
        Button2.Attributes("onclick") = "return chkdata();"
        EnterChannel.Attributes("onchange") = "EnterChannelChange();"
        TRNDMode.Attributes("onchange") = "TRNDModeChange();"
        StudentID.Attributes("onblur") = "chk_studentID(this.value,this);"
        Button4.Attributes("onclick") = "if(document.getElementById('IDNO').value==''){alert('請輸入身分證號碼');return false;}"
        PassPortNO.Attributes("onclick") = "ChangePassPort();"
        ChinaOrNot.Attributes("onclick") = "if(getRadioValue(document.form1.ChinaOrNot)==1){document.getElementById('Nationality').value='中國';}else{document.getElementById('Nationality').value='';}"
        AcctMode.Attributes("onclick") = "ChangeBank();"

        DGTR.Style.Item("display") = "none"
        GovTR.Style.Item("display") = "none"

        'by Vicient 20060804 民族別欄位
        Tr1.Style("display") = "none"

        '如果是產學訓則英文姓名 緊急通知人 失業週數部份不顯示要求強制輸入的*號 add by nick
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            star1.Visible = False
            star2.Visible = False
            star3.Visible = False
            star4.Visible = False
            star5.Visible = False
            star6.Visible = False
        End If

        'end
        If Not IsPostBack Then
            Me.ViewState("_SearchStr") = Session("_SearchStr")
            Session("_SearchStr") = Nothing
            TPlanID.Value = sm.UserInfo.TPlanID
            RoleID.Value = sm.UserInfo.RoleID
            SolTR.Style.Item("display") = "none"
            TRNDTR.Style.Item("display") = "none"
            Add_Items()
            GetOpenDate()
            If Not Session("SearchSOCID") Is Nothing Then           '處裡狀態
                Process.Value = "edit"
                create(Session("SearchSOCID"))
                Session("SearchSOCID") = Nothing
                StdTr.Visible = True
                Button2.Visible = True
                'Button3.Visible = False
                Button4.Visible = False
            Else
                ChinaOrNotTable.Style("display") = "none"
                PPNO.Style("display") = "none"
                ForeTr1.Style("display") = "none"
                ForeTr2.Style("display") = "none"
                ForeTr3.Style("display") = "none"
                ForeTr4.Style("display") = "none"
                ForeTr5.Style("display") = "none"
                PortTR.Style("display") = "none"
                BankTR1.Style("display") = "none"
                BankTR2.Style("display") = "none"
                BankTR3.Style("display") = "none"
                Process.Value = "add"
                Button4.Visible = True
                StdTr.Visible = False
                Button2.Visible = False
                If Request("TICKET_NO") <> "" Then
                    createDG()
                End If
            End If
        End If
        'add by nick 060316 加入時數限制
        Dim sql As String
        Dim DGHR As DataTable

        sql = "SELECT * FROM Key_DGTHour"
        DGHR = DbAccess.GetDataTable(sql, objconn)

        Label1.Text = DGHR.DefaultView(0)(2)
        Label2.Text = DGHR.DefaultView(1)(2)
        Label3.Text = DGHR.DefaultView(2)(2)
        Label4.Text = DGHR.DefaultView(3)(2)
        'end

        LearnTR1.Style("display") = "none"
        LearnTR2.Style("display") = "none"
        LearnTR3.Style("display") = "none"
        LearnTR4.Style("display") = "none"
        LearnTR5.Style("display") = "none"

        TPlan23TR.Visible = False
        MenuTable.Visible = False
        BackTable.Visible = False
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            MenuTable.Visible = True
            BackTable.Visible = True
            If Not IsPostBack Then
                Page.RegisterStartupScript("11111", "<script>ChangeMode(1)</script>")
            End If
        Else
            Select Case sm.UserInfo.TPlanID
                Case "15"
                    LearnTR1.Style("display") = "inline"
                    LearnTR2.Style("display") = "inline"
                    LearnTR3.Style("display") = "inline"
                    LearnTR4.Style("display") = "inline"
                    LearnTR5.Style("display") = "inline"
                Case "23", "34", "41"
                    '23:訓用合一 
                    '34:與企業合作辦理職前訓練 
                    '41:推動營造業事業單位辦理職前培訓計畫
                    TPlan23TR.Visible = True
            End Select
        End If

        GetScript()
    End Sub

    Function createDG()
        Dim rqOCID As String = TIMS.ClearSQM(Request("OCID"))
        Dim rqTICKET_NO As String = TIMS.ClearSQM(Request("TICKET_NO"))
        Dim sql As String
        Dim dr As DataRow

        sql = "" & vbCrLf
        sql += " SELECT a.IDNO" & vbCrLf
        sql += " ,b.Name,b.Sex,b.Birth,b.Marri,b.Edgr,b.Gradu,b.School,b.DeptName" & vbCrLf
        sql += " ,b.Solder,b.Addr_Zip,b.Addr,b.Tel,b.Mobile,b.Email,c.Share_Name " & vbCrLf
        sql += " FROM Adp_DGTRNData a" & vbCrLf
        sql += " join Adp_StdData b on a.IDNO=b.IDNO" & vbCrLf
        sql += " LEFT JOIN (SELECT Share_Name,Share_ID FROM Adp_ShareSource WHERE Share_Type='301') c ON a.OBJECT_TYPE=c.Share_ID" & vbCrLf
        sql += " LEFT JOIN Adp_WorkStation d ON a.CREATE_RGSTN=d.Station_Scheme_ID+d.Station_Unit_ID+d.Station_ID" & vbCrLf
        sql += " where 1=1" & vbCrLf
        sql += " and a.TICKET_NO='" & rqTICKET_NO & "'" & vbCrLf
        dr = DbAccess.GetOneRow(sql, objconn)

        If Not dr Is Nothing Then
            IDNO.Text = TIMS.ChangeIDNO(dr("IDNO").ToString)
            Name.Text = dr("Name").ToString
            If dr("Sex").ToString = "1" Then
                Common.SetListItem(Sex, "M")
            ElseIf dr("Sex").ToString = "2" Then
                Common.SetListItem(Sex, "F")
            End If
            If dr("Birth").ToString <> "" Then
                Birthday.Text = FormatDateTime(dr("Birth"), DateFormat.ShortDate)
            End If

            Select Case Convert.ToString(dr("Marri"))
                Case "1", "2"
                    Common.SetListItem(MaritalStatus, dr("Marri").ToString)
                Case Else
                    Common.SetListItem(MaritalStatus, "3")
            End Select

            Common.SetListItem(DegreeID, dr("Edgr").ToString)
            Common.SetListItem(GraduateStatus, dr("Gradu").ToString)
            School.Text = dr("School").ToString
            Department.Text = dr("DeptName").ToString
            Common.SetListItem(MilitaryID, dr("Solder").ToString)
            If dr("Addr_Zip").ToString <> "" Then
                City1.Text = "(" & dr("Addr_Zip").ToString & ")" & TIMS.Get_ZipName(dr("Addr_Zip").ToString)
                ZipCode1.Value = dr("Addr_Zip").ToString
            End If
            Address.Text = dr("Addr").ToString
            PhoneD.Text = dr("Tel").ToString
            CellPhone.Text = dr("Mobile").ToString
            Email.Text = dr("Email").ToString
            DGIdentValue.Text = dr("Share_Name").ToString

            Page.RegisterStartupScript("DG", "<script>sol(" & MilitaryID.SelectedValue & ");</script>")
            DGTR.Style.Item("display") = "inline"
        End If
    End Function

    '新增資料時取得開訓日期
    Function GetOpenDate()
        Dim rqOCID As String = TIMS.ClearSQM(Request("OCID"))
        Dim rqTICKET_NO As String = TIMS.ClearSQM(Request("TICKET_NO"))

        Dim sql As String
        Dim dr As DataRow

        sql = "" & vbCrLf
        sql += " SELECT a.ClassCName" & vbCrLf
        sql += " ,a.CyclType,a.LevelCount,a.STDate,a.FTDate" & vbCrLf
        sql += " ,b.TPlanID,d.ActNo " & vbCrLf
        sql += " FROM Class_ClassInfo a" & vbCrLf
        sql += " JOIN ID_Plan b ON a.PlanID=b.PlanID" & vbCrLf
        sql += " JOIN Auth_Relship c ON a.RID=c.RID" & vbCrLf
        sql += " JOIN Org_OrgPlanInfo d ON c.RSID=d.RSID" & vbCrLf
        sql += " WHERE 1=1" & vbCrLf
        sql += " AND a.OCID='" & rqOCID & "'" & vbCrLf
        dr = DbAccess.GetOneRow(sql, objconn)

        If Not dr Is Nothing Then
            ClassName.Text = dr("ClassCName").ToString
            If Int(dr("CyclType")) <> 0 Then
                ClassName.Text += "第" & Int(dr("CyclType")) & "期"
            End If
            If Not IsPostBack Then
                LevelNo.Items.Clear()
                If dr("LevelCount").ToString <> "" Then
                    If Int(dr("LevelCount")) <> 0 Then
                        For i As Integer = 1 To Int(dr("LevelCount"))
                            LevelNo.Items.Add(New ListItem("第" & i & "階段", i))
                        Next
                        LevelNo.Items.Insert(0, "====請選擇====")
                    Else
                        LevelNo.Items.Add(New ListItem("無區分階段", 0))
                        LevelNo.Enabled = False
                    End If
                Else
                    LevelNo.Items.Add(New ListItem("無區分階段", 0))
                    LevelNo.Enabled = False
                End If
            End If

            If OpenDate.Text = "" Then
                OpenDate.Text = FormatDateTime(dr("STDate"), 2)
            End If
            If CloseDate.Text = "" Then
                CloseDate.Text = FormatDateTime(dr("FTDate"), 2)
            End If

            If ActNo.Text = "" Then
                ActNo.Text = dr("ActNo").ToString
            End If
        End If
    End Function

    Function GetScript()
        Dim rqOCID As String = TIMS.ClearSQM(Request("OCID"))
        Dim rqTICKET_NO As String = TIMS.ClearSQM(Request("TICKET_NO"))

        Dim javascript As String
        Dim sql As String
        Dim dt As DataTable
        Dim dr As DataRow
        Dim i As Integer
        sql = "SELECT StudentID FROM Class_StudentsOfClass WHERE OCID='" & rqOCID & "'"
        dt = DbAccess.GetDataTable(sql, objconn)

        javascript = "<script language='javascript'>" & vbCrLf
        javascript += "   function chk_studentID(num,obj){" & vbCrLf
        javascript += "      var all=new Array("
        For i = 0 To dt.Rows.Count - 1
            If i = 0 Then
                javascript += "'" & Right(dt.Rows(i).Item("StudentID"), 2) & "'"
            Else
                javascript += ",'" & Right(dt.Rows(i).Item("StudentID"), 2) & "'"
            End If
        Next
        javascript += ");" & vbCrLf
        javascript += "      for(var i=0;i<all.length;i++){" & vbCrLf
        javascript += "         if(document.form1.StudentID.value==all[i] && all[i]!=document.form1.StudentIDstring.value){" & vbCrLf
        javascript += "            alert('學號重複');" & vbCrLf
        javascript += "            obj.focus();" & vbCrLf
        javascript += "         }" & vbCrLf
        javascript += "      }" & vbCrLf
        javascript += "   }" & vbCrLf
        javascript += "</script>"

        Page.RegisterStartupScript("SID", javascript)
        '        Me.ViewState("script") = javascript
    End Function

    Function create(ByVal SOCIDStr As String)
        Dim sql As String
        Dim dr As DataRow

        'sql = "SELECT *,c.IdentityID as IdentityIDEX,c.SubsidyID as SubsidyIDEX "
        'sql += "FROM (SELECT * FROM Stud_StudentInfo) a,"
        'sql += "(SELECT * FROM Stud_SubData) b,"
        'sql += "(SELECT * FROM Class_StudentsOfClass WHERE SOCID='" & SOCIDStr & "') c "
        'sql += "WHERE a.SID=b.SID AND c.SID=b.SID"

        sql = "" & vbCrLf
        sql += " SELECT c.SOCID" & vbCrLf
        sql += " ,c.IdentityID as IdentityIDEX" & vbCrLf
        sql += " ,c.SubsidyID as SubsidyIDEX" & vbCrLf
        sql += " ,c.LevelNo" & vbCrLf
        sql += " ,c.StudentID" & vbCrLf
        sql += " ,a.Name" & vbCrLf
        sql += " ,a.EngName" & vbCrLf
        sql += " ,a.PassPortNO" & vbCrLf
        sql += " ,a.ChinaOrNot" & vbCrLf
        sql += " ,a.Nationality" & vbCrLf
        sql += " ,a.PPNO" & vbCrLf
        sql += " ,b.ForeName" & vbCrLf
        sql += " ,b.ForeTitle" & vbCrLf
        sql += " ,b.ForeSex" & vbCrLf
        sql += " ,b.ForeBirth" & vbCrLf
        sql += " ,b.ForeIDNO" & vbCrLf
        sql += " ,b.ForeZip" & vbCrLf
        sql += " ,b.ForeAddr" & vbCrLf
        sql += " ,a.IDNO" & vbCrLf
        sql += " ,a.Sex" & vbCrLf
        sql += " ,a.Birthday" & vbCrLf
        sql += " ,a.MaritalStatus" & vbCrLf
        sql += " ,c.EnterChannel" & vbCrLf
        sql += " ,c.TRNDMode" & vbCrLf
        sql += " ,c.OpenDate" & vbCrLf
        sql += " ,c.CloseDate" & vbCrLf
        sql += " ,c.EnterDate" & vbCrLf
        sql += " ,a.DegreeID" & vbCrLf
        sql += " ,b.School" & vbCrLf
        sql += " ,b.Department" & vbCrLf
        sql += " ,a.GraduateStatus" & vbCrLf
        sql += " ,a.MilitaryID" & vbCrLf
        sql += " ,b.ServiceID" & vbCrLf
        sql += " ,b.MilitaryAppointment" & vbCrLf
        sql += " ,b.MilitaryRank" & vbCrLf
        sql += " ,b.ServiceOrg" & vbCrLf
        sql += " ,b.ChiefRankName" & vbCrLf
        sql += " ,b.ServicePhone" & vbCrLf
        sql += " ,b.SServiceDate" & vbCrLf
        sql += " ,b.FServiceDate" & vbCrLf
        sql += " ,b.ZipCode4" & vbCrLf
        sql += " ,b.ServiceAddress" & vbCrLf
        sql += " ,b.PhoneD" & vbCrLf
        sql += " ,b.PhoneN" & vbCrLf
        sql += " ,b.CellPhone" & vbCrLf
        sql += " ,b.ZipCode1" & vbCrLf
        sql += " ,b.Address" & vbCrLf
        sql += " ,b.ZipCode2" & vbCrLf
        sql += " ,b.HouseholdAddress" & vbCrLf
        sql += " ,b.Email" & vbCrLf
        sql += " ,c.MIdentityID" & vbCrLf
        sql += " ,c.Native" & vbCrLf
        sql += " ,b.HandTypeID" & vbCrLf
        sql += " ,b.HandLevelID" & vbCrLf
        sql += " ,c.RejectTDate1" & vbCrLf
        sql += " ,c.RejectTDate2" & vbCrLf
        sql += " ,b.EmergencyContact" & vbCrLf
        sql += " ,b.EmergencyPhone" & vbCrLf
        sql += " ,b.EmergencyRelation" & vbCrLf
        sql += " ,b.ZipCode3" & vbCrLf
        sql += " ,b.EmergencyAddress" & vbCrLf
        sql += " ,b.PriorWorkOrg1" & vbCrLf
        sql += " ,b.Title1" & vbCrLf
        sql += " ,b.SOfficeYM1" & vbCrLf
        sql += " ,b.FOfficeYM1" & vbCrLf
        sql += " ,b.PriorWorkOrg2" & vbCrLf
        sql += " ,b.Title2" & vbCrLf
        sql += " ,b.SOfficeYM2" & vbCrLf
        sql += " ,b.FOfficeYM2" & vbCrLf
        sql += " ,b.PriorWorkPay" & vbCrLf
        sql += " ,a.RealJobless" & vbCrLf
        sql += " ,a.JoblessID" & vbCrLf
        sql += " ,b.Traffic" & vbCrLf
        sql += " ,b.ShowDetail" & vbCrLf
        sql += " ,a.IsAgree" & vbCrLf
        sql += " ,c.BudgetID" & vbCrLf
        sql += " ,c.PMode" & vbCrLf
        sql += " ,c.RelClass_Unit" & vbCrLf
        sql += " ,c.Unit1Hour" & vbCrLf
        sql += " ,c.Unit2Hour" & vbCrLf
        sql += " ,c.Unit3Hour" & vbCrLf
        sql += " ,c.Unit4Hour" & vbCrLf
        sql += " ,c.Unit1Score" & vbCrLf
        sql += " ,c.Unit2Score" & vbCrLf
        sql += " ,c.Unit3Score" & vbCrLf
        sql += " ,c.Unit4Score" & vbCrLf
        sql += " ,c.ActNo" & vbCrLf
        sql += " FROM Class_StudentsOfClass c " & vbCrLf
        sql += " join Stud_StudentInfo a on a.SID=c.SID" & vbCrLf
        sql += " join Stud_SubData b on b.SID=c.SID" & vbCrLf
        sql += " where 1=1" & vbCrLf
        sql += " and c.SOCID='" & SOCIDStr & "'" & vbCrLf
        dr = DbAccess.GetOneRow(sql, objconn)

        If dr Is Nothing Then
            Common.MessageBox(Me, "找不到此學員!")
            Page.RegisterStartupScript("", "<script>location.herf='SD_03_006.aspx?ID=" & Request("ID") & "';</script>")
        Else
            Common.SetListItem(SOCID, dr("SOCID").ToString)
            Common.SetListItem(LevelNo, dr("LevelNo").ToString)
            Name.Text = Convert.ToString(dr("Name"))
            StudentID.Text = Right(Convert.ToString(dr("StudentID")), 2)
            StudentIDstring.Value = Right(Convert.ToString(dr("StudentID")), 2)
            If Convert.ToString(dr("EngName")) <> "" Then
                If Split(dr("EngName"), " ", , CompareMethod.Text).Length = 1 Then
                    LName.Text = dr("EngName").ToString
                Else
                    LName.Text = Trim(Left(dr("EngName").ToString, dr("EngName").ToString.IndexOf(" ")))
                    FName.Text = Trim(Right(dr("EngName").ToString, dr("EngName").ToString.Length - 1 - dr("EngName").ToString.IndexOf(" ")))
                End If
            End If
            Common.SetListItem(PassPortNO, dr("PassPortNO").ToString)
            If dr("PassPortNO").ToString = "1" Then
                ChinaOrNotTable.Style("display") = "none"
                PPNO.Style("display") = "none"
                ForeTr1.Style("display") = "none"
                ForeTr2.Style("display") = "none"
                ForeTr3.Style("display") = "none"
                ForeTr4.Style("display") = "none"
                ForeTr5.Style("display") = "none"
                ChinaOrNot.SelectedIndex = -1
                Nationality.Text = ""
                PPNO.SelectedIndex = -1
            Else
                ChinaOrNotTable.Style("display") = "inline"
                PPNO.Style("display") = "inline"
                ForeTr1.Style("display") = "inline"
                ForeTr2.Style("display") = "inline"
                ForeTr3.Style("display") = "inline"
                ForeTr4.Style("display") = "inline"
                ForeTr5.Style("display") = "inline"
                Common.SetListItem(ChinaOrNot, dr("ChinaOrNot").ToString)
                Nationality.Text = dr("Nationality").ToString
                Common.SetListItem(PPNO, dr("PPNO").ToString)
                ForeName.Text = dr("ForeName").ToString
                ForeTitle.Text = dr("ForeTitle").ToString
                Common.SetListItem(ForeSex, dr("ForeSex").ToString)
                If IsDate(dr("ForeBirth")) Then
                    ForeBirth.Text = FormatDateTime(dr("ForeBirth"), 2)
                End If
                ForeIDNO.Text = TIMS.ChangeIDNO(dr("ForeIDNO").ToString)
                If dr("ForeZip").ToString <> "" Then
                    City6.Text = "(" & dr("ForeZip").ToString & ")" & TIMS.Get_ZipName(dr("ForeZip").ToString)
                    ForeZip.Value = dr("ForeZip").ToString
                    ForeAddr.Text = dr("ForeAddr").ToString
                End If
            End If
            IDNO.Text = TIMS.ChangeIDNO(dr("IDNO").ToString)
            Common.SetListItem(Sex, dr("Sex").ToString)
            If Convert.ToString(dr("Birthday")) <> "" Then
                Birthday.Text = FormatDateTime(Convert.ToString(dr("Birthday")), DateFormat.ShortDate)
            End If
            Select Case Convert.ToString(dr("MaritalStatus"))
                Case "1", "2"
                    Common.SetListItem(MaritalStatus, dr("MaritalStatus").ToString)
                Case Else
                    Common.SetListItem(MaritalStatus, "3")
            End Select

            Common.SetListItem(EnterChannel, dr("EnterChannel").ToString)
            Page.RegisterStartupScript("1111", "<script>EnterChannelChange();</script>")
            Common.SetListItem(TRNDMode, dr("TRNDMode").ToString)

            Dim count As Integer
            sql = "SELECT count(*) FROM Adp_TRNData WHERE SOCID='" & SOCIDStr & "'"
            count = DbAccess.ExecuteScalar(sql, objconn)
            sql = "SELECT count(*) FROM Adp_DGTRNData WHERE SOCID='" & SOCIDStr & "'"
            count += DbAccess.ExecuteScalar(sql, objconn)
            sql = "SELECT count(*) FROM Adp_GOVTRNData WHERE SOCID='" & SOCIDStr & "'"
            count += DbAccess.ExecuteScalar(sql, objconn)
            If count > 0 Then
                '表示從三合一報名
                EnterChannel.Enabled = False
                TRNDMode.Enabled = False
                TRNDType.Enabled = False

                TRNDTR.Style.Item("display") = "inline"
                Select Case dr("TRNDMode").ToString
                    Case "1"
                        Common.SetListItem(TRNDType, dr("TRNDType").ToString)
                    Case "2"
                        DGTR.Style.Item("display") = "inline"
                        GetDGIdent(SOCIDStr)
                        Page.RegisterStartupScript("0000", "<script>TRNDModeChange();</script>")
                    Case "3"
                        GovTR.Style.Item("display") = "inline"
                        GetGovIdent(SOCIDStr)
                        Common.SetListItem(TRNDType, dr("TRNDType").ToString)
                        If dr("TRNDType").ToString = "2" Then
                            EnterChannel.Enabled = True
                            TRNDMode.Enabled = True
                            TRNDType.Enabled = True
                        End If
                End Select
            Else
                Select Case dr("TRNDMode").ToString
                    Case "1", "3"
                        Common.SetListItem(TRNDType, dr("TRNDType").ToString)
                    Case Else
                        Page.RegisterStartupScript("1112", "<script>TRNDModeChange();</script>")
                End Select
            End If

            If Convert.ToString(dr("OpenDate")) <> "" Then
                OpenDate.Text = FormatDateTime(Convert.ToString(dr("OpenDate")), DateFormat.ShortDate)
            End If
            If Convert.ToString(dr("CloseDate")) <> "" Then
                CloseDate.Text = FormatDateTime(Convert.ToString(dr("CloseDate")), DateFormat.ShortDate)
            End If
            If Convert.ToString(dr("EnterDate")) <> "" Then
                EnterDate.Text = FormatDateTime(Convert.ToString(dr("EnterDate")), DateFormat.ShortDate)
            End If
            Common.SetListItem(DegreeID, dr("DegreeID").ToString)
            School.Text = Convert.ToString(dr("School"))
            Department.Text = Convert.ToString(dr("Department"))
            Common.SetListItem(GraduateStatus, dr("GraduateStatus").ToString)
            Common.SetListItem(MilitaryID, dr("MilitaryID").ToString)
            If dr("MilitaryID").ToString = "04" Then
                SolTR.Style.Item("display") = "inline"
            End If
            ServiceID.Text = Convert.ToString(dr("ServiceID"))
            MilitaryAppointment.Text = Convert.ToString(dr("MilitaryAppointment"))
            MilitaryRank.Text = Convert.ToString(dr("MilitaryRank"))
            ServiceOrg.Text = Convert.ToString(dr("ServiceOrg"))
            ChiefRankName.Text = Convert.ToString(dr("ChiefRankName"))
            ServicePhone.Text = Convert.ToString(dr("ServicePhone"))
            If Convert.ToString(dr("SServiceDate")) <> "" Then
                SServiceDate.Text = FormatDateTime(Convert.ToString(dr("SServiceDate")), DateFormat.ShortDate)
            End If
            If Convert.ToString(dr("FServiceDate")) <> "" Then
                FServiceDate.Text = FormatDateTime(Convert.ToString(dr("FServiceDate")), DateFormat.ShortDate)
            End If
            If Convert.ToString(dr("ZipCode4")) <> "" Then
                City4.Text = "(" & Convert.ToString(dr("ZipCode4")) & ")" & TIMS.Get_ZipName(dr("ZipCode4"))
            End If
            ZipCode4.Value = Convert.ToString(dr("ZipCode4"))
            ServiceAddress.Text = Convert.ToString(dr("ServiceAddress"))
            PhoneD.Text = Convert.ToString(dr("PhoneD"))
            PhoneN.Text = Convert.ToString(dr("PhoneN"))
            CellPhone.Text = Convert.ToString(dr("CellPhone"))
            If dr("ZipCode1").ToString <> "" Then
                City1.Text = "(" & Convert.ToString(dr("ZipCode1")) & ")" & TIMS.Get_ZipName(Convert.ToString(dr("ZipCode1")))
                ZipCode1.Value = Convert.ToString(dr("ZipCode1"))
            End If
            Address.Text = Convert.ToString(dr("Address"))
            If Convert.ToString(dr("ZipCode2")) <> "" Then
                City2.Text = "(" & Convert.ToString(dr("ZipCode2")) & ")" & TIMS.Get_ZipName(Convert.ToString(dr("ZipCode2")))
                ZipCode2.Value = Convert.ToString(dr("ZipCode2"))
            End If
            HouseholdAddress.Text = Convert.ToString(dr("HouseholdAddress"))
            If Convert.ToString(dr("ZipCode1")) = Convert.ToString(dr("ZipCode2")) And Convert.ToString(dr("Address")) = Convert.ToString(dr("HouseholdAddress")) Then
                CheckBox1.Checked = True
            End If
            Email.Text = Convert.ToString(dr("Email"))


            Common.SetListItem(SubsidyID, dr("SubsidyIDEX").ToString)
            If dr("SubsidyIDEX").ToString = "03" Then
                SubsidyHidden.Value = "1"
            Else
                SubsidyHidden.Value = "0"
            End If
            SubsidyID.Attributes("onchange") = "ChangeSubsidy();"

            Common.SetListItem(MIdentityID, dr("MIdentityID").ToString)
            'by Vicient
            If dr("MIdentityID").ToString = "05" Then
                Tr1.Style("display") = "inline"
                If IsDBNull(dr("Native")) Then
                    NativeID.Items(0).Selected = True
                End If
            Else
                Tr1.Style("display") = "none"
            End If
            If Not IsDBNull(dr("Native")) Then
                Common.SetListItem(NativeID, dr("Native").ToString)
            End If

            If Convert.ToString(dr("IdentityIDEX")) <> "" Then
                If InStr(Convert.ToString(dr("IdentityIDEX")), "06", CompareMethod.Binary) = 0 Then
                    HandTypeID.Enabled = False
                    HandLevelID.Enabled = False
                Else
                    HandTypeID.Enabled = True
                    HandLevelID.Enabled = True
                End If
                Dim all() = Split(Convert.ToString(dr("IdentityIDEX")), ",", , CompareMethod.Text)
                For i As Integer = 0 To IdentityID.Items.Count - 1
                    For j As Integer = 0 To all.Length - 1
                        If IdentityID.Items(i).Value = all(j) Then
                            IdentityID.Items(i).Selected = True
                        End If
                    Next
                Next
            End If
            Common.SetListItem(HandTypeID, dr("HandTypeID").ToString)
            Common.SetListItem(HandLevelID, dr("HandLevelID").ToString)
            If Convert.ToString(dr("RejectTDate1")) <> "" Then
                RejectTDate1.Text = FormatDateTime(Convert.ToString(dr("RejectTDate1")), DateFormat.ShortDate)
            End If
            If Convert.ToString(dr("RejectTDate2")) <> "" Then
                RejectTDate2.Text = FormatDateTime(Convert.ToString(dr("RejectTDate2")), DateFormat.ShortDate)
            End If
            EmergencyContact.Text = Convert.ToString(dr("EmergencyContact"))
            EmergencyPhone.Text = Convert.ToString(dr("EmergencyPhone"))
            EmergencyRelation.Text = Convert.ToString(dr("EmergencyRelation"))
            If Convert.ToString(dr("ZipCode3")) <> "" Then
                City3.Text = "(" & Convert.ToString(dr("ZipCode3")) & ")" & TIMS.Get_ZipName(Convert.ToString(dr("ZipCode3")))
                ZipCode3.Value = Convert.ToString(dr("ZipCode3"))
            End If
            EmergencyAddress.Text = Convert.ToString(dr("EmergencyAddress"))
            PriorWorkOrg1.Text = Convert.ToString(dr("PriorWorkOrg1"))
            Title1.Text = Convert.ToString(dr("Title1"))
            If Convert.ToString(dr("SOfficeYM1")) <> "" Then
                SOfficeYM1.Text = FormatDateTime(Convert.ToString(dr("SOfficeYM1")), DateFormat.ShortDate)
            End If
            If Convert.ToString(dr("FOfficeYM1")) <> "" Then
                FOfficeYM1.Text = FormatDateTime(Convert.ToString(dr("FOfficeYM1")), DateFormat.ShortDate)
            End If
            PriorWorkOrg2.Text = Convert.ToString(dr("PriorWorkOrg2"))
            Title2.Text = Convert.ToString(dr("Title2"))
            Convert.ToString(dr("SOfficeYM2"))
            If Convert.ToString(dr("SOfficeYM2")) <> "" Then
                SOfficeYM2.Text = FormatDateTime(Convert.ToString(dr("SOfficeYM2")), DateFormat.ShortDate)
            End If
            If Convert.ToString(dr("FOfficeYM2")) <> "" Then
                FOfficeYM2.Text = FormatDateTime(Convert.ToString(dr("FOfficeYM2")), DateFormat.ShortDate)
            End If
            PriorWorkPay.Text = Convert.ToString(dr("PriorWorkPay"))
            RealJobless.Text = Convert.ToString(dr("RealJobless"))

            lb_msg.Text = ""
            RealJobless.Style.Add("background-color", "fffff")
            If TIMS.IsInt(Trim(Convert.ToString(dr("RealJobless")))) Then
                If chkJobless(Convert.ToString(dr("RealJobless")), Convert.ToString(dr("JoblessID"))) = False Then
                    lb_msg.Text = "*所填寫之受訓前失業週數與<br/>所選擇下拉式選單選項不一致!"
                    RealJobless.Style.Add("background-color", "LightPink")
                End If
            End If

            Common.SetListItem(JoblessID, dr("JoblessID").ToString)
            Common.SetListItem(Traffic, dr("Traffic").ToString)
            Common.SetListItem(ShowDetail, dr("ShowDetail").ToString)
            Common.SetListItem(IsAgree, dr("IsAgree").ToString)

            Common.SetListItem(BudID, dr("BudgetID").ToString)
            If BudID.Items.Count = 1 Then
                BudID.SelectedIndex = -1
                BudID.Items(0).Selected = True
            End If
            Common.SetListItem(PMode, dr("PMode").ToString)
            For i As Integer = 0 To dr("RelClass_Unit").ToString.Length - 1
                If dr("RelClass_Unit").ToString.Chars(i) = "1" Then
                    RelClass_Unit.Items(i).Selected = True
                End If
            Next
            Unit1Hour.Text = dr("Unit1Hour").ToString
            Unit2Hour.Text = dr("Unit2Hour").ToString
            Unit3Hour.Text = dr("Unit3Hour").ToString
            Unit4Hour.Text = dr("Unit4Hour").ToString
            'add by nick
            Unit1Score.Text = dr("Unit1Score").ToString
            Unit2Score.Text = dr("Unit2Score").ToString
            Unit3Score.Text = dr("Unit3Score").ToString
            Unit4Score.Text = dr("Unit4Score").ToString
            'end
            ActNo.Text = dr("ActNo").ToString

            If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                '企訓專用
                sql = "SELECT * FROM Stud_ServicePlace WHERE SOCID='" & SOCIDStr & "'"
                dr = DbAccess.GetOneRow(sql, objconn)
                If Not dr Is Nothing Then
                    If IsDBNull(dr("AcctMode")) Then
                        AcctMode.SelectedIndex = 1
                        AcctNo2.Text = dr("AcctNo").ToString
                        PortTR.Style("display") = "none"
                        BankTR1.Style("display") = "inline"
                        BankTR2.Style("display") = "inline"
                        BankTR3.Style("display") = "inline"
                    Else
                        If dr("AcctMode") = False Then
                            AcctMode.SelectedIndex = 0
                            If dr("PostNo").ToString.IndexOf("-") = -1 Then
                                PostNo_1.Text = dr("PostNo").ToString
                            Else
                                PostNo_1.Text = Left(dr("PostNo").ToString, dr("PostNo").ToString.IndexOf("-"))
                                PostNo_2.Text = Right(dr("PostNo").ToString, dr("PostNo").ToString.Length - dr("PostNo").ToString.IndexOf("-") - 1)
                            End If
                            If dr("AcctNo").ToString.IndexOf("-") = -1 Then
                                AcctNo1_1.Text = dr("AcctNo").ToString
                            Else
                                AcctNo1_1.Text = Left(dr("AcctNo").ToString, dr("AcctNo").ToString.IndexOf("-"))
                                AcctNo1_2.Text = Right(dr("AcctNo").ToString, dr("AcctNo").ToString.Length - dr("AcctNo").ToString.IndexOf("-") - 1)
                            End If

                            PortTR.Style("display") = "inline"
                            BankTR1.Style("display") = "none"
                            BankTR2.Style("display") = "none"
                            BankTR3.Style("display") = "none"
                        Else
                            AcctMode.SelectedIndex = 1
                            BankName.Text = dr("BankName").ToString
                            '    ExBankName.Text = dr("ExBankName").ToString
                            AcctHeadNo.Text = dr("AcctHeadNo").ToString
                            '   AcctExNo.Text = dr("AcctExNo").ToString
                            AcctNo2.Text = dr("AcctNo").ToString

                            PortTR.Style("display") = "none"
                            BankTR1.Style("display") = "inline"
                            BankTR2.Style("display") = "inline"
                            BankTR3.Style("display") = "inline"
                        End If
                    End If
                    If IsDate(dr("FirDate")) Then
                        FirDate.Text = FormatDateTime(dr("FirDate"), 2)
                    End If
                    Uname.Text = dr("Uname").ToString
                    Intaxno.Text = dr("Intaxno").ToString
                    ServDept.Text = dr("ServDept").ToString
                    JobTitle.Text = dr("JobTitle").ToString
                    City5.Text = "(" & dr("Zip").ToString & ")" & TIMS.Get_ZipName(dr("Zip").ToString)
                    Zip.Value = dr("Zip").ToString
                    Addr.Text = dr("Addr").ToString
                    Tel.Text = dr("Tel").ToString
                    Fax.Text = dr("Fax").ToString
                    If IsDate(dr("SDate")) Then
                        SDate.Text = FormatDateTime(dr("SDate"), 2)
                    End If
                    If IsDate(dr("SJDate")) Then
                        SJDate.Text = FormatDateTime(dr("SJDate"), 2)
                    End If
                    If IsDate(dr("SPDate")) Then
                        SPDate.Text = FormatDateTime(dr("SPDate"), 2)
                    End If
                End If

                sql = "SELECT * FROM Stud_TrainBG WHERE SOCID='" & SOCIDStr & "'"
                dr = DbAccess.GetOneRow(sql, objconn)
                If Not dr Is Nothing Then
                    If dr("Q1") Then
                        Q1.SelectedIndex = 0
                    Else
                        Q1.SelectedIndex = 1
                    End If
                    Common.SetListItem(Q3, dr("Q3"))
                    Common.SetListItem(Q4, dr("Q4"))
                    If IsDBNull(dr("Q5")) Then
                        Q5.SelectedIndex = -1
                    Else
                        If dr("Q5") Then
                            Q5.SelectedIndex = 0
                        Else
                            Q5.SelectedIndex = 1
                        End If
                    End If
                    Q61.Text = dr("Q61").ToString
                    Q62.Text = dr("Q62").ToString
                    Q63.Text = dr("Q63").ToString
                    Q64.Text = dr("Q64").ToString
                End If

                sql = "SELECT * FROM Stud_TrainBGQ2 WHERE SOCID='" & SOCIDStr & "'"
                Dim dt As DataTable
                dt = DbAccess.GetDataTable(sql, objconn)
                For Each dr In dt.Rows
                    For Each item As ListItem In Q2.Items
                        If dr("Q2") = item.Value Then
                            item.Selected = True
                        End If
                    Next
                Next
            End If
        End If
    End Function

    '取出學習卷的學員身分資料
    Function GetDGIdent(ByVal SOCID As Integer)
        Dim sql As String
        Dim dr As DataRow

        sql = ""
        sql += " SELECT b.Share_Name  "
        sql += " FROM Adp_DGTRNData a "
        sql += " JOIN (SELECT Share_Name,Share_ID FROM Adp_ShareSource WHERE Share_Type='301') b ON a.OBJECT_TYPE=b.Share_ID "
        sql += " where 1=1"
        sql += " and a.SOCID='" & SOCID & "'"
        dr = DbAccess.GetOneRow(sql, objconn)
        DGIdentValue.Text = dr("Share_Name")
    End Function

    '取出學習卷的學員身分資料
    Function GetGovIdent(ByVal SOCID As Integer)
        Dim sql As String
        Dim dr As DataRow

        sql = ""
        sql += " SELECT b.Share_Name  "
        sql += " FROM Adp_GOVTRNData a "
        sql += " JOIN (SELECT Share_Name,Share_ID FROM Adp_ShareSource WHERE Share_Type='527') b ON a.OBJECT_TYPE=b.Share_ID "
        sql += " where 1=1"
        sql += " and a.SOCID='" & SOCID & "'"
        dr = DbAccess.GetOneRow(sql, objconn)
        If Not dr Is Nothing Then
            GovObject_Type.Text = dr("Share_Name").ToString
        End If

        sql = ""
        sql += " SELECT b.Share_Name  "
        sql += " FROM Adp_GOVTRNData a "
        sql += " JOIN (SELECT Share_Name,Share_ID FROM Adp_ShareSource WHERE Share_Type='528') b ON a.OBJECT_TYPE=b.Share_ID "
        sql += " where 1=1"
        sql += " and a.SOCID='" & SOCID & "'"
        dr = DbAccess.GetOneRow(sql, objconn)
        If Not dr Is Nothing Then
            GovSpecial_Type.Text = dr("Share_Name").ToString
        End If
    End Function

    Function Add_Items()
        Dim rqOCID As String = TIMS.ClearSQM(Request("OCID"))
        Dim rqTICKET_NO As String = TIMS.ClearSQM(Request("TICKET_NO"))

        Dim sql As String
        Dim dr As DataRow
        Dim dt As DataTable

        'by Vicient 20060804 增加民族別選項
        sql = "SELECT KNID,Name FROM Key_Native ORDER BY KNID "
        dt = DbAccess.GetDataTable(sql, objconn)
        With NativeID
            .DataSource = dt
            .DataTextField = "Name"
            .DataValueField = "KNID"
            .DataBind()
            .Items.Insert(0, New ListItem("===請選擇===", 0))
        End With

        DegreeID = TIMS.Get_Degree(DegreeID)
        ''by Vicient
        ''sql = "select * from Key_Degree where DegreeID IN ('01','02','03','04','05','06')"
        'dt = DbAccess.GetDataTable(sql)
        'If dt.Rows.Count <> 0 Then
        '    With DegreeID
        '        .DataSource = dt
        '        .DataTextField = "Name"
        '        .DataValueField = "DegreeID"
        '        .DataBind()
        '        .Items.Insert(0, New ListItem("===請選擇===", 0)) '請選擇
        '    End With
        'End If
        GraduateStatus = TIMS.Get_GradState(GraduateStatus)

        'MilitaryID = TIMS.Get_Military(MilitaryID)
        '列出兵役下拉選單資料-by Vicient
        sql = "SELECT MilitaryID,NAME FROM Key_Military where MilitaryID <> '00' ORDER BY MilitaryID"
        dt = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count <> 0 Then
            With MilitaryID
                .DataSource = dt
                .DataTextField = "Name"
                .DataValueField = "MilitaryID"
                .DataBind()
                .Items.Insert(0, New ListItem("===請選擇===", 0))
            End With
        End If

        MilitaryID.Attributes("onchange") = "sol(this.value)"
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            sql = "SELECT * FROM Key_Identity WHERE IdentityID IN ('01','03','04','05','06','07','10') ORDER BY IdentityID"
            dt = DbAccess.GetDataTable(sql, objconn)
            With MIdentityID
                .DataSource = dt
                .DataTextField = "Name"
                .DataValueField = "IdentityID"
                .DataBind()
                .Items.Insert(0, New ListItem("===請選擇===", ""))
            End With
            With IdentityID
                .DataSource = dt
                .DataTextField = "Name"
                .DataValueField = "IdentityID"
                .DataBind()
            End With
        Else
            MIdentityID = TIMS.Get_Identity(MIdentityID, 2)
            IdentityID = TIMS.Get_Identity(IdentityID, 2)
        End If
        With IdentityID
            .Attributes("onclick") = "hard(" & IdentityID.ClientID & ")"
        End With
        SubsidyID = TIMS.Get_SubsidyID(SubsidyID)
        HandTypeID = TIMS.Get_HandicatType(HandTypeID)
        HandLevelID = TIMS.Get_HandicatLevel(HandLevelID)
        JoblessID = TIMS.Get_JoblessID(JoblessID, Nothing, Me.sm.UserInfo.Years)
        RelClass_Unit = TIMS.Get_DGTHour(RelClass_Unit)

        'add by nick 20060327
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            SubsidyID.Items.Insert(0, New ListItem("未申請", "01"))
        End If

        Call TIMS.Get_Trade(Q4)

        '建立StudentID值
        sql = "" & vbCrLf
        sql += " SELECT a.Years,b.ClassID,a.CyclType " & vbCrLf
        sql += " FROM Class_ClassInfo a " & vbCrLf
        sql += " join ID_Class b ON a.CLSID=b.CLSID" & vbCrLf
        sql += " where 1=1" & vbCrLf
        sql += " and a.OCID ='" & rqOCID & "'" & vbCrLf
        dr = DbAccess.GetOneRow(sql, objconn)
        If dr Is Nothing Then
            Common.MessageBox(Me, "沒有班別學號基本碼")
        Else
            StudentIDValue.Value = dr("Years").ToString & "0" & dr("ClassID").ToString & dr("CyclType").ToString
        End If

        sql = "" & vbCrLf
        sql += " SELECT a.StudentID , b.Name+'('+substr(a.StudentID,-2)+')' Name,a.SOCID" & vbCrLf
        sql += " FROM Class_StudentsOfClass a" & vbCrLf
        sql += " JOIN Stud_StudentInfo b ON a.SID=b.SID" & vbCrLf
        sql += " where 1=1" & vbCrLf
        sql += " and a.OCID ='" & rqOCID & "'" & vbCrLf
        dt = DbAccess.GetDataTable(sql, objconn)
        dt.DefaultView.Sort = "StudentID"
        With SOCID
            .DataSource = dt
            .DataTextField = "Name"
            .DataValueField = "SOCID"
            .DataBind()
            .Items.Insert(0, New ListItem("===請選擇===", ""))
        End With

        sql = "" & vbCrLf
        sql += " SELECT k.budname,k.budid " & vbCrLf
        sql += " FROM Plan_Budget b" & vbCrLf
        sql += " join Key_Budget k on k.BudID=b.BudID" & vbCrLf
        sql += " join id_plan ip on ip.TPlanID =b.TPlanID AND ip.Years =b.SYear" & vbCrLf
        sql += " where 1=1" & vbCrLf
        sql += " and ip.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
        dt = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count = 0 Then
            BudIDMsg.Text = "尚未設定預算"
        Else
            With BudID
                .DataSource = dt
                .DataTextField = "BudName"
                .DataValueField = "BudID"
                .DataBind()
            End With

            If BudID.Items.Count = 1 Then
                BudID.Items(0).Selected = True
            End If
        End If
    End Function

    Function clear_data()
        Name.Text = ""
        StudentID.Text = ""
        LName.Text = ""
        FName.Text = ""
        IDNO.Text = ""
        Birthday.Text = ""
        OpenDate.Text = ""
        CloseDate.Text = ""
        EnterDate.Text = ""
        School.Text = ""
        Department.Text = ""
        ServiceID.Text = ""
        MilitaryAppointment.Text = ""
        MilitaryRank.Text = ""
        ServiceOrg.Text = ""
        ChiefRankName.Text = ""
        ServicePhone.Text = ""
        SServiceDate.Text = ""
        FServiceDate.Text = ""
        City4.Text = ""
        ZipCode4.Value = ""
        ServiceAddress.Text = ""
        PhoneD.Text = ""
        PhoneN.Text = ""
        CellPhone.Text = ""
        City1.Text = ""
        ZipCode1.Value = ""
        Address.Text = ""
        City2.Text = ""
        ZipCode2.Value = ""
        HouseholdAddress.Text = ""
        Email.Text = ""
        RejectTDate1.Text = ""
        RejectTDate2.Text = ""
        EmergencyContact.Text = ""
        EmergencyPhone.Text = ""
        EmergencyRelation.Text = ""
        City3.Text = ""
        ZipCode3.Value = ""
        EmergencyAddress.Text = ""
        PriorWorkOrg1.Text = ""
        Title1.Text = ""
        PriorWorkOrg2.Text = ""
        Title2.Text = ""
        SOfficeYM1.Text = ""
        FOfficeYM1.Text = ""
        SOfficeYM2.Text = ""
        FOfficeYM2.Text = ""
        PriorWorkPay.Text = ""
        RealJobless.Text = ""
        PostNo_1.Text = ""
        PostNo_2.Text = ""
        AcctNo1_1.Text = ""
        AcctNo1_2.Text = ""
        BankName.Text = ""
        '  ExBankName.Text = ""
        AcctHeadNo.Text = ""
        '    AcctExNo.Text = ""
        AcctNo2.Text = ""
        FirDate.Text = ""
        Uname.Text = ""
        Intaxno.Text = ""
        Tel.Text = ""
        Fax.Text = ""
        City5.Text = ""
        Zip.Value = ""
        Addr.Text = ""
        ServDept.Text = ""
        JobTitle.Text = ""
        SDate.Text = ""
        SJDate.Text = ""
        SPDate.Text = ""

        CheckBox1.Checked = False

        If Me.ViewState("ADD") <> 1 Then
            If Not PassPortNO.SelectedItem Is Nothing Then
                PassPortNO.SelectedItem.Selected = False
            End If
            If Not Sex.SelectedItem Is Nothing Then
                Sex.SelectedItem.Selected = False
            End If
            If Not MaritalStatus.SelectedItem Is Nothing Then
                MaritalStatus.SelectedItem.Selected = False
            End If
        End If
        If Not DegreeID.SelectedItem Is Nothing Then
            DegreeID.SelectedItem.Selected = False
        End If
        If Not GraduateStatus.SelectedItem Is Nothing Then
            GraduateStatus.SelectedItem.Selected = False
        End If
        If Not MilitaryID.SelectedItem Is Nothing Then
            MilitaryID.SelectedItem.Selected = False
        End If
        If Not NativeID.SelectedItem Is Nothing Then
            NativeID.SelectedItem.Selected = False
        End If
        If Not SubsidyID.SelectedItem Is Nothing Then
            SubsidyID.SelectedItem.Selected = False
        End If
        If Not HandTypeID.SelectedItem Is Nothing Then
            HandTypeID.SelectedItem.Selected = False
        End If
        If Not HandLevelID.SelectedItem Is Nothing Then
            HandLevelID.SelectedItem.Selected = False
        End If
        If Not JoblessID.SelectedItem Is Nothing Then
            JoblessID.SelectedItem.Selected = False
        End If
        If Not Traffic.SelectedItem Is Nothing Then
            Traffic.SelectedItem.Selected = False
        End If
        If Not ShowDetail.SelectedItem Is Nothing Then
            ShowDetail.SelectedItem.Selected = False
        End If
        If Not TRNDMode.SelectedItem Is Nothing Then
            TRNDMode.SelectedItem.Selected = False
        End If
        If Not TRNDType.SelectedItem Is Nothing Then
            TRNDType.SelectedItem.Selected = False
        End If
        If Not EnterChannel.SelectedItem Is Nothing Then
            EnterChannel.SelectedItem.Selected = False
        End If

        For i As Integer = 0 To IdentityID.Items.Count - 1
            IdentityID.Items(i).Selected = False
        Next
        If Not BudID.SelectedItem Is Nothing Then
            BudID.SelectedItem.Selected = False
        End If
        For i As Integer = 0 To RelClass_Unit.Items.Count - 1
            RelClass_Unit.Items(i).Selected = False
        Next
        'For i As Integer = 0 To BudID.Items.Count - 1
        '    BudID.Items(i).Selected = False
        'Next
        AcctMode.SelectedIndex = -1
        PMode.SelectedIndex = -1
        Q1.SelectedIndex = -1
        For Each item As ListItem In Q2.Items
            item.Selected = False
        Next
        Q3.SelectedIndex = -1
        Q3_Other.Text = ""
        Q4.SelectedIndex = -1
        Q5.SelectedIndex = -1
        Q61.Text = ""
        Q62.Text = ""
        Q63.Text = ""
        Q64.Text = ""

        SolTR.Style.Item("display") = "none"
        TRNDTR.Style.Item("display") = "none"
    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click, Button2.Click
        Call Savedata1()

        Common.RespWrite(Me, "<script>alert('儲存成功');</script>")
        If sender Is Button1 Then
            '儲存回查詢頁面
            Session("_SearchStr") = Session("_SearchStr")
            If Not Me.ViewState("_SearchStr") Is Nothing Then
                Session("_SearchStr") = Me.ViewState("_SearchStr")
            End If
            Common.RespWrite(Me, "<script>location.href='SD_03_006.aspx?ID=" & Request("ID") & "'</script>")
        ElseIf sender Is Button2 Then
            '維護下一位學員
            Dim Index As Integer = SOCID.SelectedIndex
            If Index < SOCID.Items.Count - 1 Then
                clear_data()
                SOCID.SelectedItem.Selected = False
                SOCID.Items(Index + 1).Selected = True
                create(SOCID.SelectedValue)
                GetOpenDate()
            Else
                Common.MessageBox(Me, "已經到最後一位學員!")
            End If
        End If
    End Sub

    '儲存
    Sub Savedata1()
        Dim iSOCIDValue As Integer = 0 '班級學員PK@SOCID
        Dim SID As String = "" '學員序號一組。
        Dim rqOCID As String = TIMS.ClearSQM(Request("OCID"))
        Dim rqTICKET_NO As String = TIMS.ClearSQM(Request("TICKET_NO"))
        IDNO.Text = TIMS.ChangeIDNO(IDNO.Text)
        IDNO.Text = TIMS.ClearSQM(IDNO.Text)
        Birthday.Text = TIMS.ClearSQM(Birthday.Text)
        Birthday.Text = TIMS.cdate3(Birthday.Text)

        '先檢查是否有資料
        If StdTr.Visible = False Then
            '新增狀態，檢查有沒有個人資料存在
            If TIMS.Chk_ClassStudent2(rqOCID, IDNO.Text, objconn) Then
                Common.MessageBox(Me, "此班級已經有相同的身分證號碼!")
                Page.RegisterStartupScript("hard", "<script>hard();</script>")
                Exit Sub
            End If
            If TIMS.Chk_StudentInfo(IDNO.Text, Birthday.Text, objconn) Then
                SID = TIMS.Get_StudentInfoSID(IDNO.Text, Birthday.Text, objconn)
                Common.MessageBox(Me, "此學員個人資料已存在，將更新您所輸入的資料") '提醒
            Else
                SID = TIMS.Get_DateNo & "01" '預設值
            End If
        Else
            '修改狀態
            iSOCIDValue = SOCID.SelectedValue
            If iSOCIDValue <= 0 Then
                Common.MessageBox(Me, "沒有選擇正確的學員，請重新選擇!") '錯誤
                Page.RegisterStartupScript("hard", "<script>hard();</script>")
                Exit Sub
            End If
            If Not TIMS.Chk_ClassStudent(rqOCID, iSOCIDValue, objconn) Then
                Common.MessageBox(Me, "沒有選擇正確的學員，請重新選擇!") '錯誤
                Page.RegisterStartupScript("hard", "<script>hard();</script>")
                Exit Sub
            End If
            '檢查是否有填寫津貼申請
            If Not chk_SubsidyResult(Me, CStr(iSOCIDValue), IdentityID, objconn) Then
                Exit Sub '檢查是否有填寫津貼申請 不通過停止儲存
            End If
            If TIMS.Chk_StudentInfo(IDNO.Text, Birthday.Text, objconn) Then
                SID = TIMS.Get_StudentInfoSID(IDNO.Text, Birthday.Text, objconn)
                Common.MessageBox(Me, "此學員個人資料已存在，將更新您所輸入的資料") '提醒
            Else
                SID = TIMS.Get_DateNo & "01" '預設值
                Common.MessageBox(Me, "此學員個人資料不存在，將要把您輸入的資料新增存入")
            End If
        End If

        '其他基本資料檢查。
        Dim ErrMessage As String = ""
        If ErrMessage = "" Then
            If Not TIMS.checkMemberSex(IDNO.Text, Sex.SelectedValue) Then
                ErrMessage += "依身分證號判斷 性別選項 不正確！" & vbCrLf
            End If
        End If
        If ErrMessage <> "" Then
            Common.MessageBox(Me, ErrMessage)
            Exit Sub
        End If

        '學習券新增，更新三合一資料
        Dim dt As DataTable = Nothing '學員背景資料
        Dim da As OracleDataAdapter = Nothing
        Dim dr As DataRow = Nothing
        Dim sql As String = ""
        Dim objTrans As OracleTransaction = Nothing
        Try
            objTrans = DbAccess.BeginTrans(objconn)
            '學員基本資料主檔
            Call UPDATE_StudentInfo(SID, objTrans)
            '學員基本資料副檔
            Call UPDATE_SubData(SID, objTrans)
            '學員班級資料 
            Call UPDATE_StudentsOfClass(iSOCIDValue, SID, rqOCID, objTrans)
            'iSOCIDValue 若是修改為INPUT 若是新增為OUTPUT

            If SubsidyID.SelectedValue <> "03" Then
                If StdTr.Visible = True Then
                    '修改狀態
                    sql = "DELETE Stud_SubsidyResult WHERE SOCID='" & iSOCIDValue & "'"
                    DbAccess.ExecuteNonQuery(sql, objTrans)
                End If
            End If

            '企訓專用------------------------------------------------------Start
            If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                '學員背景資料
                Call UPDATE_ServicePlace(CStr(iSOCIDValue), objTrans)
            End If
            '企訓專用------------------------------------------------------End

            If rqTICKET_NO <> "" Then
                Dim OrgName As String = ""
                Dim ComIDNO As String = ""
                Dim ContactName As String = ""
                Dim ContactPhone As String = ""

                sql = "" & vbCrLf
                sql += " SELECT b.OrgName,b.ComIDNO,d.ContactName,d.Phone" & vbCrLf
                sql += " FROM Class_ClassInfo a" & vbCrLf
                sql += " JOIN Org_OrgInfo b ON a.ComIDNO=b.ComIDNO" & vbCrLf
                sql += " JOIN Auth_Relship c ON a.RID=c.RID" & vbCrLf
                sql += " JOIN Org_OrgPlanInfo d ON c.RSID=d.RSID" & vbCrLf
                sql += " where a.OCID='" & rqOCID & "'" & vbCrLf
                dr = DbAccess.GetOneRow(sql, objTrans)
                OrgName = dr("OrgName").ToString
                ComIDNO = dr("ComIDNO").ToString
                ContactName = dr("ContactName").ToString
                ContactPhone = dr("Phone").ToString

                sql = "SELECT * FROM Adp_DGTRNData WHERE TICKET_NO='" & rqTICKET_NO & "'"
                dt = DbAccess.GetDataTable(sql, da, objTrans)
                dr = dt.Rows(0)
                dr("SOCID") = iSOCIDValue
                dr("ARVL_STATE") = 1
                dr("ARVL_DATE") = IIf(EnterDate.Text = "", Convert.DBNull, EnterDate.Text)
                dr("ARVL_UNIT_NAME") = OrgName
                dr("ARVL_ORG_NAME") = OrgName
                dr("ARVL_ORG_DOCNO") = ComIDNO
                dr("ARVL_SDATE") = IIf(OpenDate.Text = "", Convert.DBNull, OpenDate.Text)
                dr("ARVL_EDATE") = IIf(CloseDate.Text = "", Convert.DBNull, CloseDate.Text)
                dr("ARVL_UNIT_PROMOTER") = ContactName
                dr("ARVL_UNIT_TEL") = ContactPhone
                dr("ACT_END_DATE") = IIf(CloseDate.Text = "", Convert.DBNull, CloseDate.Text)
                dr("ARVL_CLASS_NAME") = ClassName.Text
                dr("ARVL_CLASS_NO") = rqOCID
                dr("TransToTIMS") = "Y"
                DbAccess.UpdateDataTable(dt, da, objTrans)
            End If


            '結訓學員資料卡更新----------------------------------Start
            If StdTr.Visible = True Then
                sql = "SELECT * FROM Stud_ResultStudData WHERE SOCID='" & iSOCIDValue & "'"
                dt = DbAccess.GetDataTable(sql, da, objTrans)
                If dt.Rows.Count <> 0 Then
                    dr = dt.Rows(0)
                    Dim DLID As Integer = dr("DLID")
                    Dim SubNo As Integer = dr("SubNo")
                    dr("StdName") = Name.Text
                    dr("StudentID") = StudentID.Text
                    dr("StdPID") = TIMS.ChangeIDNO(IDNO.Text)
                    dr("Sex") = Sex.SelectedIndex + 1
                    dr("BirthYear") = CDate(Birthday.Text).Year
                    dr("BirthMonth") = CDate(Birthday.Text).Month
                    dr("BirthDate") = CDate(Birthday.Text).Day
                    dr("DegreeID") = DegreeID.SelectedValue
                    dr("MilitaryID") = MilitaryID.SelectedValue
                    dr("ModifyAcct") = sm.UserInfo.UserID
                    dr("ModifyDate") = Now
                    DbAccess.UpdateDataTable(dt, da, objTrans)

                    'Stud_ResultIdentData不在匯入此 TABLE 改用 Class_StudentsOfClass.IdentityID	參訓身分別代碼
                    'BY AMU 2009-07-30
                    '非局屬狀況加入 Stud_ResultIdentData  BY AMU 2009-08-25
                    sql = "SELECT * FROM Stud_ResultIdentData WHERE DLID='" & DLID & "' and SubNo='" & SubNo & "'"
                    dt = DbAccess.GetDataTable(sql, da, objTrans)
                    For Each item As ListItem In IdentityID.Items
                        If item.Value <> "" Then
                            If item.Selected = True Then
                                If dt.Select("IdentityID='" & item.Value & "'").Length = 0 Then
                                    dr = dt.NewRow
                                    dt.Rows.Add(dr)
                                    dr("DLID") = DLID
                                    dr("SubNo") = SubNo
                                    dr("IdentityID") = item.Value
                                End If
                            Else
                                If dt.Select("IdentityID='" & item.Value & "'").Length <> 0 Then
                                    dt.Select("IdentityID='" & item.Value & "'")(0).Delete()
                                End If
                            End If
                        End If
                    Next
                    DbAccess.UpdateDataTable(dt, da, objTrans)
                End If
            End If
            '結訓學員資料卡更新----------------------------------End
            DbAccess.CommitTrans(objTrans)
        Catch ex As Exception
            DbAccess.RollbackTrans(objTrans)
            Throw ex
        End Try
    End Sub

    Private Sub SOCID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SOCID.SelectedIndexChanged
        If SOCID.SelectedIndex <> 0 Then
            clear_data()
            create(SOCID.SelectedValue)
            GetOpenDate()
        End If
    End Sub

    Public ReadOnly Property MySearch() As String
        Get
            Return Me.ViewState("search")
        End Get
    End Property

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Session("_SearchStr") = Me.ViewState("_SearchStr")
        If Request("TICKET_NO") <> "" Then
            TIMS.Utl_Redirect1(Me, "SD_03_006_3in1.aspx?ID=" & Request("ID"))
        Else
            TIMS.Utl_Redirect1(Me, "SD_03_006.aspx?ID=" & Request("ID"))
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim sql As String
        Dim dr As DataRow

        sql = "" & vbCrLf
        sql += " SELECT  " & vbCrLf
        sql += " a.SID  /*PK*/ " & vbCrLf
        sql += " ,a.IDNO" & vbCrLf
        sql += " ,a.Name" & vbCrLf
        sql += " ,a.EngName" & vbCrLf
        sql += " ,a.PassPortNO" & vbCrLf
        sql += " ,a.Sex" & vbCrLf
        sql += " ,a.Birthday" & vbCrLf
        sql += " ,a.MaritalStatus" & vbCrLf
        sql += " ,a.DegreeID" & vbCrLf
        sql += " ,a.GraduateStatus" & vbCrLf
        sql += " ,a.MilitaryID" & vbCrLf
        sql += " ,a.IdentityID" & vbCrLf
        sql += " ,a.SubsidyID" & vbCrLf
        sql += " ,a.JoblessID" & vbCrLf
        sql += " ,a.RealJobless" & vbCrLf
        sql += " ,a.GetCertificate" & vbCrLf
        sql += " ,a.GetSubsidy" & vbCrLf
        sql += " ,a.IsAgree" & vbCrLf
        sql += " ,a.LaInFlag" & vbCrLf
        sql += " ,a.ChinaOrNot" & vbCrLf
        sql += " ,a.Nationality" & vbCrLf
        sql += " ,a.PPNO" & vbCrLf
        sql += " ,a.JobState" & vbCrLf
        sql += " ,a.FType" & vbCrLf
        sql += " ,a.ActNo" & vbCrLf
        sql += " ,a.MDate" & vbCrLf
        sql += " ,a.SalID" & vbCrLf
        sql += " ,a.FixID" & vbCrLf
        sql += " ,a.JoblessID_99" & vbCrLf
        sql += " ,a.GraduateY" & vbCrLf
        sql += " ,b.School" & vbCrLf
        sql += " ,b.Department" & vbCrLf
        sql += " ,b.ZipCode1" & vbCrLf
        sql += " ,b.Address" & vbCrLf
        sql += " ,b.ZipCode2" & vbCrLf
        sql += " ,b.HouseholdAddress" & vbCrLf
        sql += " ,b.Email" & vbCrLf
        sql += " ,b.PhoneD" & vbCrLf
        sql += " ,b.PhoneN" & vbCrLf
        sql += " ,b.CellPhone" & vbCrLf
        sql += " ,b.EmergencyContact" & vbCrLf
        sql += " ,b.EmergencyRelation" & vbCrLf
        sql += " ,b.EmergencyPhone" & vbCrLf
        sql += " ,b.ZipCode3" & vbCrLf
        sql += " ,b.EmergencyAddress" & vbCrLf
        sql += " ,b.PriorWorkOrg1" & vbCrLf
        sql += " ,b.Title1" & vbCrLf
        sql += " ,b.SOfficeYM1" & vbCrLf
        sql += " ,b.FOfficeYM1" & vbCrLf
        sql += " ,b.PriorWorkOrg2" & vbCrLf
        sql += " ,b.Title2" & vbCrLf
        sql += " ,b.SOfficeYM2" & vbCrLf
        sql += " ,b.FOfficeYM2" & vbCrLf
        sql += " ,b.PriorWorkPay" & vbCrLf
        sql += " ,b.Traffic" & vbCrLf
        sql += " ,b.ShowDetail" & vbCrLf
        sql += " ,b.ServiceID" & vbCrLf
        sql += " ,b.MilitaryAppointment" & vbCrLf
        sql += " ,b.MilitaryRank" & vbCrLf
        sql += " ,b.SServiceDate" & vbCrLf
        sql += " ,b.FServiceDate" & vbCrLf
        sql += " ,b.ServiceOrg" & vbCrLf
        sql += " ,b.ChiefRankName" & vbCrLf
        sql += " ,b.ZipCode4" & vbCrLf
        sql += " ,b.ServiceAddress" & vbCrLf
        sql += " ,b.ServicePhone" & vbCrLf
        sql += " ,b.HandTypeID" & vbCrLf
        sql += " ,b.HandLevelID" & vbCrLf
        sql += " ,b.ForeName" & vbCrLf
        sql += " ,b.ForeTitle" & vbCrLf
        sql += " ,b.ForeSex" & vbCrLf
        sql += " ,b.ForeBirth" & vbCrLf
        sql += " ,b.ForeIDNO" & vbCrLf
        sql += " ,b.ForeZip" & vbCrLf
        sql += " ,b.ForeAddr" & vbCrLf
        sql += " ,b.ZipCode1_2w" & vbCrLf
        sql += " ,b.ZipCode2_2w" & vbCrLf
        sql += " ,b.ZipCode3_2w" & vbCrLf
        sql += " ,b.ZipCode4_2w" & vbCrLf
        sql += " ,b.ForeZip2w" & vbCrLf
        sql += " FROM Stud_StudentInfo a" & vbCrLf
        sql += " join Stud_SubData b on a.SID =b.SID" & vbCrLf
        sql += " WHERE 1=1" & vbCrLf
        sql += " AND upper(a.IDNO)=upper('" & TIMS.ChangeIDNO(IDNO.Text) & "')" & vbCrLf
        dr = DbAccess.GetOneRow(sql, objconn)
        If Not dr Is Nothing Then
            Name.Text = dr("Name").ToString
            If dr("EngName").ToString.IndexOf(" ") = -1 Then
                LName.Text = dr("EngName").ToString
            Else
                LName.Text = Trim(Left(dr("EngName").ToString, dr("EngName").ToString.IndexOf(" ")))
                FName.Text = Trim(Right(dr("EngName").ToString, dr("EngName").ToString.Length - 1 - dr("EngName").ToString.IndexOf(" ")))
            End If
            Common.SetListItem(PassPortNO, dr("PassPortNO").ToString)
            IDNO.Text = TIMS.ChangeIDNO(dr("IDNO").ToString)
            Common.SetListItem(Sex, dr("Sex").ToString)
            If dr("Birthday").ToString <> "" Then
                Birthday.Text = FormatDateTime(dr("Birthday"), 2)
            End If
            Select Case Convert.ToString(dr("MaritalStatus"))
                Case "1", "2"
                    Common.SetListItem(MaritalStatus, dr("MaritalStatus").ToString)
                Case Else
                    Common.SetListItem(MaritalStatus, "3")
            End Select

            Common.SetListItem(DegreeID, dr("DegreeID").ToString)
            School.Text = dr("School").ToString
            Department.Text = dr("Department").ToString
            Common.SetListItem(GraduateStatus, dr("GraduateStatus").ToString)
            Common.SetListItem(MilitaryID, dr("MilitaryID").ToString)
            If dr("MilitaryID").ToString = "04" Then
                SolTR.Style.Item("display") = "inline"
            End If
            ServiceID.Text = dr("ServiceID").ToString
            MilitaryAppointment.Text = dr("MilitaryAppointment").ToString
            MilitaryRank.Text = dr("MilitaryRank").ToString
            ServiceOrg.Text = dr("ServiceOrg").ToString
            If dr("SServiceDate").ToString <> "" Then
                SServiceDate.Text = FormatDateTime(dr("SServiceDate"), 2)
            End If
            If dr("FServiceDate").ToString <> "" Then
                FServiceDate.Text = FormatDateTime(dr("FServiceDate"), 2)
            End If
            If dr("ZipCode4").ToString <> "" Then
                City4.Text = "(" & dr("ZipCode4").ToString & ")" & TIMS.Get_ZipName(dr("ZipCode4").ToString)
                ZipCode4.Value = dr("ZipCode4").ToString
            End If
            ServiceAddress.Text = dr("ServiceAddress").ToString
            PhoneD.Text = dr("PhoneD").ToString
            PhoneN.Text = dr("PhoneN").ToString
            CellPhone.Text = dr("CellPhone").ToString
            If dr("ZipCode1").ToString <> "" Then
                City1.Text = "(" & dr("ZipCode1").ToString & ")" & TIMS.Get_ZipName(dr("ZipCode1").ToString)
                ZipCode1.Value = dr("ZipCode1").ToString
            End If
            Address.Text = dr("Address").ToString
            If dr("ZipCode2").ToString <> "" Then
                City2.Text = "(" & dr("ZipCode2").ToString & ")" & TIMS.Get_ZipName(dr("ZipCode2").ToString)
                ZipCode2.Value = dr("ZipCode2").ToString
            End If
            HouseholdAddress.Text = dr("HouseholdAddress").ToString
            Email.Text = dr("Email").ToString
            Page.RegisterStartupScript("hard", "<script>hard();</script>")
            Common.SetListItem(HandTypeID, dr("HandTypeID").ToString)
            Common.SetListItem(HandLevelID, dr("HandLevelID").ToString)
            EmergencyContact.Text = dr("EmergencyContact").ToString
            EmergencyPhone.Text = dr("EmergencyPhone").ToString
            EmergencyRelation.Text = dr("EmergencyRelation").ToString
            If dr("ZipCode3").ToString <> "" Then
                City3.Text = "(" & dr("ZipCode3").ToString & ")" & TIMS.Get_ZipName(dr("ZipCode3").ToString)
                ZipCode3.Value = dr("ZipCode3").ToString
            End If
            EmergencyAddress.Text = dr("EmergencyAddress").ToString
            PriorWorkOrg1.Text = dr("PriorWorkOrg1").ToString
            PriorWorkOrg1.Text = dr("PriorWorkOrg1").ToString
            Title1.Text = dr("Title1").ToString
            Title2.Text = dr("Title2").ToString
            If dr("SOfficeYM1").ToString <> "" Then
                SOfficeYM1.Text = FormatDateTime(dr("SOfficeYM1"), 2)
            End If
            If dr("FOfficeYM1").ToString <> "" Then
                FOfficeYM1.Text = FormatDateTime(dr("FOfficeYM1"), 2)
            End If
            If dr("SOfficeYM2").ToString <> "" Then
                SOfficeYM2.Text = FormatDateTime(dr("SOfficeYM2"), 2)
            End If
            If dr("FOfficeYM2").ToString <> "" Then
                FOfficeYM2.Text = FormatDateTime(dr("FOfficeYM2"), 2)
            End If
            PriorWorkPay.Text = dr("PriorWorkPay").ToString
            RealJobless.Text = dr("RealJobless").ToString

            lb_msg.Text = ""
            RealJobless.Style.Add("background-color", "fffff")
            If TIMS.IsInt(Trim(Convert.ToString(dr("RealJobless")))) Then
                If chkJobless(Convert.ToString(dr("RealJobless")), Convert.ToString(dr("JoblessID"))) = False Then
                    lb_msg.Text = "*所填寫之受訓前失業週數與<br/>所選擇下拉式選單選項不一致!"
                    RealJobless.Style.Add("background-color", "LightPink")
                End If
            End If

            Common.SetListItem(JoblessID, dr("JoblessID").ToString)
            Common.SetListItem(Traffic, dr("Traffic").ToString)
            Common.SetListItem(ShowDetail, dr("ShowDetail").ToString)
            If BudID.Items.Count = 1 Then
                BudID.Items(0).Selected = True
            End If
            Common.SetListItem(IsAgree, dr("IsAgree").ToString)
        Else
            Common.MessageBox(Me, "查無相關參訓個人資料!")
        End If

        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '企訓專用
            Page.RegisterStartupScript("11111", "<script>ChangeMode(1)</script>")
        End If
    End Sub

#Region "Function"
    '學員基本資料主檔
    Sub UPDATE_StudentInfo(ByVal SID As String, ByRef tTrans As OracleTransaction)
        Dim MyTable1 As DataTable = Nothing '學員基本資料主檔
        Dim da1 As OracleDataAdapter = Nothing
        Dim Mydr1 As DataRow = Nothing

        Dim sql As String = ""
        sql = "SELECT * from Stud_StudentInfo WHERE SID='" & SID & "'"
        MyTable1 = DbAccess.GetDataTable(sql, da1, tTrans)
        If MyTable1.Rows.Count = 0 Then
            Mydr1 = MyTable1.NewRow
            MyTable1.Rows.Add(Mydr1)
            Mydr1("SID") = SID
        Else
            Mydr1 = MyTable1.Rows(0)
        End If
        '更新學員基本資料檔----------------------------------------------Strat
        Mydr1("IDNO") = TIMS.ChangeIDNO(IDNO.Text)
        Mydr1("Name") = Name.Text
        Mydr1("EngName") = LName.Text & " " & FName.Text
        Select Case PassPortNO.SelectedValue
            Case "1", "2"
                Mydr1("PassPortNO") = PassPortNO.SelectedValue
            Case Else
                Mydr1("PassPortNO") = "2"
        End Select
        If PassPortNO.SelectedValue = "1" Then
            Mydr1("ChinaOrNot") = Convert.DBNull
            Mydr1("Nationality") = Convert.DBNull
            Mydr1("PPNO") = Convert.DBNull
        Else
            Mydr1("ChinaOrNot") = IIf(ChinaOrNot.SelectedIndex = -1, Convert.DBNull, ChinaOrNot.SelectedValue)
            Mydr1("Nationality") = IIf(Nationality.Text = "", Convert.DBNull, Nationality.Text)
            Mydr1("PPNO") = IIf(PPNO.SelectedIndex = -1, Convert.DBNull, PPNO.SelectedValue)
        End If
        Mydr1("Sex") = Sex.SelectedValue
        Mydr1("Birthday") = Birthday.Text

        Select Case MaritalStatus.SelectedValue
            Case "1", "2"
                Mydr1("MaritalStatus") = MaritalStatus.SelectedValue
            Case Else
                Mydr1("MaritalStatus") = Convert.DBNull
        End Select

        Mydr1("DegreeID") = DegreeID.SelectedValue
        Mydr1("GraduateStatus") = GraduateStatus.SelectedValue
        Mydr1("MilitaryID") = MilitaryID.SelectedValue
        If JoblessID.SelectedValue = "0" Then
            Mydr1("JoblessID") = Convert.DBNull
        Else
            Mydr1("JoblessID") = JoblessID.SelectedValue
        End If
        If RealJobless.Text = "" Then
            Mydr1("RealJobless") = Convert.DBNull
        Else
            Mydr1("RealJobless") = RealJobless.Text
        End If
        Mydr1("IsAgree") = IsAgree.SelectedValue
        Mydr1("ModifyAcct") = sm.UserInfo.UserID
        Mydr1("ModifyDate") = Now()
        DbAccess.UpdateDataTable(MyTable1, da1, tTrans)
        '更新學員基本資料檔----------------------------------------------End
    End Sub

    '學員基本資料副檔
    Sub UPDATE_SubData(ByVal SID As String, ByRef tTrans As OracleTransaction)
        Dim MyTable2 As DataTable = Nothing
        Dim da2 As OracleDataAdapter = Nothing
        Dim Mydr2 As DataRow = Nothing

        Dim sql As String = ""
        sql = "SELECT * FROM Stud_SubData WHERE SID='" & SID & "'"
        MyTable2 = DbAccess.GetDataTable(sql, da2, tTrans)
        If MyTable2.Rows.Count = 0 Then
            Mydr2 = MyTable2.NewRow
            MyTable2.Rows.Add(Mydr2)
            Mydr2("SID") = SID
        Else
            Mydr2 = MyTable2.Rows(0)
        End If

        '更新學員資料副檔----------------------------------------------Start
        Mydr2("Name") = Name.Text
        Mydr2("School") = School.Text
        Mydr2("Department") = Department.Text
        If ZipCode1.Value = "" Then
            Mydr2("ZipCode1") = Convert.DBNull
        Else
            Mydr2("ZipCode1") = ZipCode1.Value
        End If
        Mydr2("Address") = Address.Text
        If CheckBox1.Checked = True Then
            Mydr2("ZipCode2") = ZipCode1.Value
            Mydr2("HouseholdAddress") = Address.Text
        Else
            If ZipCode2.Value = "" Then
                Mydr2("ZipCode2") = Convert.DBNull
            Else
                Mydr2("ZipCode2") = ZipCode2.Value
            End If
            Mydr2("HouseholdAddress") = HouseholdAddress.Text
        End If
        Mydr2("Email") = Email.Text
        Mydr2("PhoneD") = PhoneD.Text
        Mydr2("PhoneN") = PhoneN.Text
        Mydr2("CellPhone") = CellPhone.Text
        Mydr2("EmergencyContact") = EmergencyContact.Text
        Mydr2("EmergencyRelation") = EmergencyRelation.Text
        Mydr2("EmergencyPhone") = EmergencyPhone.Text
        If ZipCode3.Value = "" Then
            Mydr2("ZipCode3") = Convert.DBNull
        Else
            Mydr2("ZipCode3") = ZipCode3.Value
        End If
        Mydr2("EmergencyAddress") = EmergencyAddress.Text
        Mydr2("PriorWorkOrg1") = PriorWorkOrg1.Text
        Mydr2("Title1") = Title1.Text
        If SOfficeYM1.Text = "" Then
            Mydr2("SOfficeYM1") = Convert.DBNull
        Else
            Mydr2("SOfficeYM1") = SOfficeYM1.Text
        End If
        If FOfficeYM1.Text = "" Then
            Mydr2("FOfficeYM1") = Convert.DBNull
        Else
            Mydr2("FOfficeYM1") = FOfficeYM1.Text
        End If
        Mydr2("PriorWorkOrg2") = PriorWorkOrg2.Text
        Mydr2("Title2") = Title2.Text
        If SOfficeYM2.Text = "" Then
            Mydr2("SOfficeYM2") = Convert.DBNull
        Else
            Mydr2("SOfficeYM2") = SOfficeYM2.Text
        End If
        If FOfficeYM2.Text = "" Then
            Mydr2("FOfficeYM2") = Convert.DBNull
        Else
            Mydr2("FOfficeYM2") = FOfficeYM2.Text
        End If

        If PriorWorkPay.Text = "" Then
            Mydr2("PriorWorkPay") = Convert.DBNull
        Else
            Mydr2("PriorWorkPay") = PriorWorkPay.Text
        End If
        If Traffic.SelectedValue = "0" Then
            Mydr2("Traffic") = Convert.DBNull
        Else
            Mydr2("Traffic") = Traffic.SelectedValue
        End If
        If ShowDetail.SelectedValue = "Y" Then
            Mydr2("ShowDetail") = ShowDetail.SelectedValue
        Else
            Mydr2("ShowDetail") = "N"
        End If
        Mydr2("ServiceID") = ServiceID.Text
        Mydr2("MilitaryAppointment") = MilitaryAppointment.Text
        Mydr2("MilitaryRank") = MilitaryRank.Text
        If SServiceDate.Text = "" Then
            Mydr2("SServiceDate") = Convert.DBNull
        Else
            Mydr2("SServiceDate") = SServiceDate.Text
        End If

        If FServiceDate.Text = "" Then
            Mydr2("FServiceDate") = Convert.DBNull
        Else
            Mydr2("FServiceDate") = FServiceDate.Text
        End If
        Mydr2("ServiceOrg") = ServiceOrg.Text
        Mydr2("ChiefRankName") = ChiefRankName.Text
        If ZipCode4.Value = "" Then
            Mydr2("ZipCode4") = Convert.DBNull
        Else
            Mydr2("ZipCode4") = ZipCode4.Value
        End If
        Mydr2("ServiceAddress") = ServiceAddress.Text
        Mydr2("ServicePhone") = ServicePhone.Text
        If HandTypeID.SelectedIndex = 0 Then
            Mydr2("HandTypeID") = Convert.DBNull
        Else
            Mydr2("HandTypeID") = HandTypeID.SelectedValue
        End If
        If HandLevelID.SelectedIndex = 0 Then
            Mydr2("HandLevelID") = Convert.DBNull
        Else
            Mydr2("HandLevelID") = HandLevelID.SelectedValue
        End If

        '外國籍新增部分2005/12/16---------Start
        If PassPortNO.SelectedValue = "1" Then
            Mydr2("ForeName") = Convert.DBNull
            Mydr2("ForeTitle") = Convert.DBNull
            Mydr2("ForeSex") = Convert.DBNull
            Mydr2("ForeBirth") = Convert.DBNull
            Mydr2("ForeIDNO") = Convert.DBNull
            Mydr2("ForeZip") = Convert.DBNull
            Mydr2("ForeAddr") = Convert.DBNull
        Else
            Mydr2("ForeName") = IIf(ForeName.Text = "", Convert.DBNull, ForeName.Text)
            Mydr2("ForeTitle") = IIf(ForeTitle.Text = "", Convert.DBNull, ForeTitle.Text)
            Mydr2("ForeSex") = IIf(ForeSex.SelectedIndex = -1, Convert.DBNull, ForeSex.SelectedValue)
            Mydr2("ForeBirth") = IIf(ForeBirth.Text = "", Convert.DBNull, ForeBirth.Text)
            Mydr2("ForeIDNO") = IIf(ForeIDNO.Text = "", Convert.DBNull, TIMS.ChangeIDNO(ForeIDNO.Text))
            Mydr2("ForeZip") = IIf(ForeZip.Value = "", Convert.DBNull, ForeZip.Value)
            Mydr2("ForeAddr") = IIf(ForeAddr.Text = "", Convert.DBNull, ForeAddr.Text)
        End If
        '外國籍新增部分2005/12/16---------End

        Mydr2("ModifyAcct") = sm.UserInfo.UserID
        Mydr2("ModifyDate") = Now()
        DbAccess.UpdateDataTable(MyTable2, da2, tTrans)
        '更新學員資料副檔----------------------------------------------End
    End Sub

    '班級學員檔
    Sub UPDATE_StudentsOfClass(ByRef iSOCID As Integer, ByVal SID As String, ByVal rqOCID As String, ByRef tTrans As OracleTransaction)
        If iSOCID <= 0 Then
            '小於等於0表示新增SOCID
            iSOCID = DbAccess.GetNewId(tTrans, "CLASS_STUDENTSOFCLASS_SOCID_SE,CLASS_STUDENTSOFCLASS,SOCID")
        End If

        Dim MyTable3 As DataTable = Nothing '學員基本資料主檔
        Dim da3 As OracleDataAdapter = Nothing
        Dim Mydr3 As DataRow = Nothing

        Dim sql As String = ""
        sql = "SELECT * from Class_StudentsOfClass WHERE SOCID='" & iSOCID & "'"
        MyTable3 = DbAccess.GetDataTable(sql, da3, tTrans)
        If MyTable3.Rows.Count = 0 Then
            Mydr3 = MyTable3.NewRow
            MyTable3.Rows.Add(Mydr3)
            Mydr3("SOCID") = iSOCID
            Mydr3("SID") = SID '不相同的SID 修正為輸入的SID
            Mydr3("StudStatus") = 1
        Else
            Mydr3 = MyTable3.Rows(0)
            Mydr3("SID") = SID '不相同的SID 修正為輸入的SID
        End If

        '更新班級學員檔------------------------------------------------Start
        Mydr3("StudentID") = StudentIDValue.Value & StudentID.Text
        If LevelNo.Enabled = True Then
            Mydr3("LevelNo") = LevelNo.SelectedValue
        Else
            Mydr3("LevelNo") = 0
        End If
        Mydr3("OCID") = rqOCID
        If EnterDate.Text = "" Then
            Mydr3("EnterDate") = Convert.DBNull
        Else
            Mydr3("EnterDate") = EnterDate.Text
        End If
        If OpenDate.Text = "" Then
            Mydr3("OpenDate") = Convert.DBNull
        Else
            Mydr3("OpenDate") = OpenDate.Text
        End If
        If CloseDate.Text = "" Then
            Mydr3("CloseDate") = Convert.DBNull
        Else
            Mydr3("CloseDate") = CloseDate.Text
        End If
        If RejectTDate1.Text = "" Then
            Mydr3("RejectTDate1") = Convert.DBNull
        Else
            Mydr3("RejectTDate1") = RejectTDate1.Text
        End If
        If RejectTDate2.Text = "" Then
            Mydr3("RejectTDate2") = Convert.DBNull
        Else
            Mydr3("RejectTDate2") = RejectTDate2.Text
        End If
        If TRNDMode.SelectedIndex = 0 Then
            Mydr3("TRNDMode") = Convert.DBNull
        Else
            Mydr3("TRNDMode") = TRNDMode.SelectedValue
        End If
        If TRNDType.SelectedItem Is Nothing Then
            Mydr3("TRNDType") = Convert.DBNull
        Else
            Mydr3("TRNDType") = TRNDType.SelectedValue
        End If
        If EnterChannel.SelectedIndex = 0 Then
            Mydr3("EnterChannel") = Convert.DBNull
        Else
            Mydr3("EnterChannel") = EnterChannel.SelectedValue
        End If
        Mydr3("MIdentityID") = MIdentityID.SelectedValue
        'by Vicient
        If NativeID.SelectedIndex = 0 Then
            Mydr3("Native") = Convert.DBNull
        Else
            Mydr3("Native") = NativeID.SelectedValue
        End If

        Dim all_Identity As String = ""
        For i As Integer = 0 To IdentityID.Items.Count - 1
            If IdentityID.Items(i).Selected AndAlso IdentityID.Items(i).Value <> "" Then
                If all_Identity <> "" Then all_Identity &= ","
                all_Identity &= IdentityID.Items(i).Value
            End If
        Next
        Mydr3("IdentityID") = all_Identity
        Mydr3("SubsidyID") = SubsidyID.SelectedValue

        '學習券要判斷上課單元--------------------------------------------------------------Start
        If sm.UserInfo.TPlanID = "15" Then
            Mydr3("RelClass_Unit") = ""
            Mydr3("RelClass_Hour") = ""
            Mydr3("Unit1Hour") = IIf(Unit1Hour.Text = "", 0, Unit1Hour.Text)
            Mydr3("Unit2Hour") = IIf(Unit2Hour.Text = "", 0, Unit2Hour.Text)
            Mydr3("Unit3Hour") = IIf(Unit3Hour.Text = "", 0, Unit3Hour.Text)
            Mydr3("Unit4Hour") = IIf(Unit4Hour.Text = "", 0, Unit4Hour.Text)

            'add by nick 060316
            Mydr3("Unit1Score") = IIf(Unit1Score.Text = "", 0, Unit1Score.Text)
            Mydr3("Unit2Score") = IIf(Unit2Score.Text = "", 0, Unit2Score.Text)
            Mydr3("Unit3Score") = IIf(Unit3Score.Text = "", 0, Unit3Score.Text)
            Mydr3("Unit4Score") = IIf(Unit4Score.Text = "", 0, Unit4Score.Text)
            Dim i As Integer = 0
            For Each item As ListItem In RelClass_Unit.Items
                i += 1
                If item.Selected = True Then
                    Mydr3("RelClass_Unit") = Mydr3("RelClass_Unit") & "1"
                    Select Case i
                        Case 1
                            Mydr3("RelClass_Hour") = Mydr3("RelClass_Hour") & IIf(Int(Unit1Hour.Text) < 10, "0" & Unit1Hour.Text, Unit1Hour.Text)
                        Case 2
                            Mydr3("RelClass_Hour") = Mydr3("RelClass_Hour") & IIf(Int(Unit2Hour.Text) < 10, "0" & Unit2Hour.Text, Unit2Hour.Text)
                        Case 3
                            Mydr3("RelClass_Hour") = Mydr3("RelClass_Hour") & IIf(Int(Unit3Hour.Text) < 10, "0" & Unit3Hour.Text, Unit3Hour.Text)
                        Case 4
                            Mydr3("RelClass_Hour") = Mydr3("RelClass_Hour") & IIf(Int(Unit4Hour.Text) < 10, "0" & Unit4Hour.Text, Unit4Hour.Text)
                    End Select
                Else
                    Mydr3("RelClass_Unit") = Mydr3("RelClass_Unit") & "0"
                    Mydr3("RelClass_Hour") = Mydr3("RelClass_Hour") & "00"
                End If
            Next
        Else
            Mydr3("RelClass_Unit") = Convert.DBNull
            Mydr3("RelClass_Hour") = Convert.DBNull
            Mydr3("Unit1Hour") = 0
            Mydr3("Unit2Hour") = 0
            Mydr3("Unit3Hour") = 0
            Mydr3("Unit4Hour") = 0

            Mydr3("Unit1Score") = 0
            Mydr3("Unit2Score") = 0
            Mydr3("Unit3Score") = 0
            Mydr3("Unit4Score") = 0
        End If
        '學習券要判斷上課單元--------------------------------------------------------------End

        If BudID.Items.Count = 0 Then
            Mydr3("BudgetID") = Convert.DBNull
        Else
            Mydr3("BudgetID") = BudID.SelectedValue
        End If
        If PMode.SelectedIndex <> -1 Then
            Mydr3("PMode") = PMode.SelectedValue
        End If
        Mydr3("ActNo") = IIf(ActNo.Text = "", Convert.DBNull, ActNo.Text)
        Mydr3("ModifyAcct") = sm.UserInfo.UserID
        Mydr3("ModifyDate") = Now
        DbAccess.UpdateDataTable(MyTable3, da3, tTrans)

        'If StdTr.Visible = False Then
        '    SOCIDValue = DbAccess.GetId(objTrans, "CLASS_STUDENTSOFCLASS_SOCID_SE")
        'Else
        '    SOCIDValue = SOCID.SelectedValue
        'End If
        '更新班級學員檔------------------------------------------------End
    End Sub

    '學員背景資料
    Sub UPDATE_ServicePlace(ByVal SOCIDValue As String, ByRef tTrans As OracleTransaction)
        Dim dt As DataTable = Nothing '學員背景資料
        Dim da As OracleDataAdapter = Nothing
        Dim dr As DataRow = Nothing

        Dim sql As String = ""
        sql = "SELECT * FROM Stud_ServicePlace WHERE SOCID='" & SOCIDValue & "'"
        dt = DbAccess.GetDataTable(sql, da, tTrans)
        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
            dt.Rows.Add(dr)
            dr("SOCID") = Val(SOCIDValue)
        Else
            dr = dt.Rows(0)
        End If

        If AcctMode.SelectedValue = 0 Then
            dr("AcctMode") = False
            dr("PostNo") = PostNo_1.Text & "-" & PostNo_2.Text
            dr("AcctNo") = AcctNo1_1.Text & "-" & AcctNo1_2.Text
            dr("BankName") = Convert.DBNull
            '  dr("ExBankName") = Convert.DBNull
            dr("AcctHeadNo") = Convert.DBNull
            '  dr("AcctExNo") = Convert.DBNull
        ElseIf AcctMode.SelectedValue = 1 Then
            dr("AcctMode") = True
            dr("PostNo") = Convert.DBNull
            dr("BankName") = BankName.Text
            '  dr("ExBankName") = ExBankName.Text
            dr("AcctHeadNo") = AcctHeadNo.Text
            '  dr("AcctExNo") = AcctExNo.Text
            dr("AcctNo") = AcctNo2.Text
        End If
        dr("FirDate") = IIf(FirDate.Text = "", Convert.DBNull, FirDate.Text)
        dr("Uname") = IIf(Uname.Text = "", Convert.DBNull, Uname.Text)
        dr("Intaxno") = IIf(Intaxno.Text = "", Convert.DBNull, Intaxno.Text)
        dr("ServDept") = IIf(ServDept.Text = "", Convert.DBNull, ServDept.Text)
        dr("JobTitle") = IIf(JobTitle.Text = "", Convert.DBNull, JobTitle.Text)
        dr("Zip") = Zip.Value
        dr("Addr") = Addr.Text
        dr("Tel") = Tel.Text
        dr("Fax") = IIf(Fax.Text = "", Convert.DBNull, Fax.Text)
        dr("SDate") = IIf(SDate.Text = "", Convert.DBNull, SDate.Text)
        dr("SJDate") = IIf(SJDate.Text = "", Convert.DBNull, SJDate.Text)
        dr("SPDate") = IIf(SPDate.Text = "", Convert.DBNull, SPDate.Text)
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now
        DbAccess.UpdateDataTable(dt, da, tTrans)


        dt = Nothing '學員背景資料
        da = Nothing
        dr = Nothing
        sql = "SELECT * FROM Stud_TrainBG WHERE SOCID='" & SOCIDValue & "'"
        dt = DbAccess.GetDataTable(sql, da, tTrans)
        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
            dt.Rows.Add(dr)
            dr("SOCID") = Val(SOCIDValue)
        Else
            dr = dt.Rows(0)
        End If
        If Q1.SelectedIndex = 0 Then
            dr("Q1") = 1
        ElseIf Q1.SelectedIndex = 1 Then
            dr("Q1") = 0
        End If
        If Q3.SelectedIndex <> -1 Then
            dr("Q3") = Q3.SelectedValue
        End If
        dr("Q3_Other") = IIf(Q3_Other.Text = "", Convert.DBNull, Q3_Other.Text)
        If Q4.SelectedIndex <> 0 Then
            dr("Q4") = Q4.SelectedValue
        End If
        If Q5.SelectedIndex = 0 Then
            dr("Q5") = 1
        ElseIf Q5.SelectedIndex = 1 Then
            dr("Q5") = 0
        End If
        dr("Q61") = IIf(Q61.Text = "", Convert.DBNull, Q61.Text)
        dr("Q62") = IIf(Q62.Text = "", Convert.DBNull, Q62.Text)
        dr("Q63") = IIf(Q63.Text = "", Convert.DBNull, Q63.Text)
        dr("Q64") = IIf(Q64.Text = "", Convert.DBNull, Q64.Text)

        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now
        DbAccess.UpdateDataTable(dt, da, tTrans)

        sql = "DELETE Stud_TrainBGQ2 WHERE SOCID='" & SOCIDValue & "'"
        DbAccess.ExecuteNonQuery(sql, tTrans)

        dt = Nothing '學員背景資料
        da = Nothing
        dr = Nothing
        sql = "SELECT * FROM Stud_TrainBGQ2 WHERE SOCID='" & SOCIDValue & "'"
        dt = DbAccess.GetDataTable(sql, da, tTrans)
        For Each item As ListItem In Q2.Items
            If item.Selected = True Then
                dr = dt.NewRow
                dt.Rows.Add(dr)
                dr("SOCID") = SOCIDValue
                dr("Q2") = item.Value
            End If
        Next
        DbAccess.UpdateDataTable(dt, da, tTrans)
        '企訓專用------------------------------------------------------End

    End Sub

    '檢查是否有填寫津貼申請
    Public Shared Function chk_SubsidyResult(ByRef MyPage As Page, ByVal SOCID As String, ByRef IdentityID As CheckBoxList, ByVal tConn As OracleConnection)
        Dim rst As Boolean = True 'true '可繼續執行儲存 ('是否可繼續執行儲存)
        '檢查是否有填寫津貼申請---------------------------------------------Start
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " SELECT a.IdentityID,b.Name " & vbCrLf
        sql += " FROM Stud_SubsidyResult a" & vbCrLf
        sql += " JOIN Key_Identity b ON a.IdentityID=b.IdentityID" & vbCrLf
        sql += " where 1=1 and a.SOCID=@SOCID" & vbCrLf
        Dim sCmd As New OracleCommand(sql, tConn)

        TIMS.OpenDbConn(tConn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("SOCID", OracleType.VarChar).Value = SOCID
            dt.Load(.ExecuteReader())
        End With
        If dt.Rows.Count > 0 Then
            Dim dr As DataRow = dt.Rows(0)
            Dim bFlag As Boolean = False
            For i As Integer = 0 To IdentityID.Items.Count - 1
                If IdentityID.Items(i).Selected = True AndAlso IdentityID.Items(i).Value <> "" Then
                    If dr("IdentityID") = IdentityID.Items(i).Value Then
                        bFlag = True
                        Exit For
                    End If
                End If
            Next
            If Not bFlag Then
                Common.MessageBox(MyPage, "此學員已經用身分別[" & dr("Name") & "]申請生活津貼，不可以取消此身分別")
                MyPage.RegisterStartupScript("hard", "<script>hard();</script>")
                rst = False 'Exit Sub 'false '停止執行儲存
            End If
        End If
        '檢查是否有填寫津貼申請---------------------------------------------End
        Return rst '是否可繼續執行儲存
    End Function

    '判斷(參訓前)真正失業週數是否填寫與失業週數代碼相符
    Function chkJobless(ByVal RealJobless As String, ByVal JoblessID As String) As Boolean
        Dim IsOK As Boolean = True
        Dim weeks As Int16 = CInt(Trim(RealJobless))

        If CInt(sm.UserInfo.Years) >= 2010 Then
            If weeks <= 23 Then                                                    '23週(含)以下   '04'
                If JoblessID <> "04" Then
                    IsOK = False
                End If
            Else
                If weeks >= 24 And weeks <= 51 And JoblessID <> "05" Then          '24~51週       '05'
                    IsOK = False
                ElseIf weeks >= 52 And JoblessID <> "06" Then                      '52週(含)以上  '06'
                    IsOK = False
                End If
            End If
        Else
            '99年之前 下拉選單顯示 週數區間 代碼
            If weeks <= 26 Then                                                    '26週(含)以下  '01'
                If JoblessID <> "01" Then
                    IsOK = False
                End If
            Else
                If weeks >= 27 And weeks <= 52 And JoblessID <> "02" Then          '27~52週        02'
                    IsOK = False
                ElseIf weeks >= 53 And JoblessID <> "03" Then                      '53週(含)以上  '06'
                    IsOK = False
                End If
            End If
        End If
        Return IsOK
    End Function

#End Region
End Class
