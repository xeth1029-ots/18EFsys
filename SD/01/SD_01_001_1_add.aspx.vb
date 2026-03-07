Partial Class SD01_001_1_add
    Inherits System.Web.UI.Page

    '此頁功能似乎暫無使用，待有需要時再更新
    '產業人才投資方案
    Dim objconn As OracleConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在--------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), titlelab1, titlelab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在--------------------------End
        Call TIMS.OpenDbConn(objconn)

        If Not Page.IsPostBack Then
            '列出學歷下拉選單資料
            degreeid = TIMS.Get_Degree(degreeid)
            '列出兵役下拉選單資料
            MilitaryID = TIMS.Get_Military(MilitaryID)
            MilitaryID.Items.Remove(MilitaryID.Items.FindByValue("00"))
            '列出畢業狀況下拉選單資料
            graduatestatus = TIMS.Get_GradState(graduatestatus)
            '列出主要參訓身分別下拉選單資料
            midentityid = TIMS.Get_Identity(midentityid, 1)
            '列出參訓身分別選單
            identityid = TIMS.Get_Identity(identityid, 3)
            '列出障礙類別下拉選單資料
            handtypeid = TIMS.Get_HandicatType(handtypeid)
            '列出障礙等級下拉選單資料
            handlevelid = TIMS.Get_HandicatLevel(handlevelid)
            '列出行業別下拉選單資料
            q4 = TIMS.Get_Trade(q4)
            '列出失業週數下拉選單資料
            joblessid = TIMS.Get_JoblessID(joblessid, Nothing, Me.sm.UserInfo.Years)

            create()

            'SearchHistory("")
            Me.ViewState("IDNO") = Me.idno.Text
            SearchHistory(Me.ViewState("IDNO"))
        End If

        If Not Me.ViewState("proecess") = "view" Then
            If Me.identityid.Items(4).Selected = True Then
                Me.handtypeid.Enabled = True
                Me.handlevelid.Enabled = True
            End If
        End If

        backtable.Style("display") = "none"
        porttr.Style("display") = "none"
        banktr1.Style("display") = "none"
        banktr2.Style("display") = "none"
        banktr3.Style("display") = "none"
        If Me.acctmode.Items(0).Selected = True Then
            porttr.Style("display") = "inline"
            banktr1.Style("display") = "none"
            banktr2.Style("display") = "none"
            banktr3.Style("display") = "none"
        End If
        If Me.acctmode.Items(1).Selected = True Then
            porttr.Style("display") = "none"
            banktr1.Style("display") = "inline"
            banktr2.Style("display") = "inline"
            banktr3.Style("display") = "inline"
        End If

        Me.send.Attributes.Add("OnClick", "return ChkData();")
    End Sub

    'SERVER端 檢查
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim blnRst As Boolean = True
        Errmsg = ""

        If lociddate.Value = "" Then
            Errmsg += "報名班級資料有誤(沒有開訓日期)，請重新確認資料!!" & vbCrLf
        End If

        If Trim(birthday.Text) <> "" Then
            birthday.Text = Trim(birthday.Text)
            If TIMS.IsDate1(birthday.Text) Then
                birthday.Text = CDate(birthday.Text).ToString("yyyy/MM/dd")
                'Common.FormatDate(Me.birthDay.Text)
            Else
                Errmsg += "出生日期格式有誤!!" & vbCrLf
            End If
        Else
            Errmsg += "請填入出生日期" & vbCrLf
        End If

        If midentityid.SelectedValue = "" Then
            Errmsg += "請選擇 主要參訓身分別 " & vbCrLf
        End If

        If Errmsg = "" AndAlso midentityid.SelectedValue = "04" Then
            '檢測此學員是否 屬於中高齡資格 45歲~65歲
            If Not TIMS.Check_YearsOld45(birthday.Text, lociddate.Value) Then
                Errmsg += "學員資格 年齡非介於45歲~65歲之間 不符合中高齡資格！" & vbCrLf
            End If
        End If

        If Errmsg = "" AndAlso midentityid.SelectedValue = "37" Then
            '檢測此學員是否 屬於六十五歲以上者資格 65歲以上 BY AMU 20121212
            If Not TIMS.Check_YearsOld65(birthday.Text, lociddate.Value) Then
                Errmsg += "學員資格 年齡非65歲以上 不符合 六十五歲以上者 資格！" & vbCrLf
            End If
        End If

        If email.Text <> "" Then
            If Not TIMS.CheckEmail(email.Text) Then
                Errmsg += "電子郵件 EMail格式錯誤。" & vbCrLf
            End If
        End If

        If Errmsg <> "" Then blnRst = False
        Return blnRst
    End Function

    '送出
    Private Sub send_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles send.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        'Dim objadapter As OracleDataAdapter
        'Dim objadapter1 As OracleDataAdapter
        'Dim i As Integer
        'Dim IdentityIDs As String = ""
        'Dim SETID, SerNum, eSerNum, SEID, CHSEID As Integer
        'Dim EnterDate As Date
        'Dim str As String = ""
        'Dim str1 As String = ""

        Dim objstr As String = ""
        Dim objada As New OracleDataAdapter
        Dim objada1 As New OracleDataAdapter
        Dim table As DataTable
        Dim table1 As DataTable
        'Dim objtable As DataTable

        Dim dr As DataRow
        Dim dr1 As DataRow
        Dim sqldr As DataRow

        Dim i As Integer
        Dim IdentityIDs As String = ""
        'Dim SETID, SerNum, eSerNum, SEID, CHSEID As Integer
        Dim EnterDate As Date

        Dim str As String = ""
        Dim str1 As String = ""

        Dim eSerNum As Integer = 0
        Dim SEID As Integer = 0
        Dim CHSEID As Integer = 0 'Stud_EnterTrain2@SEID
        Dim SETID As Integer = 0
        Dim SerNum As Integer = 0

        '線上報名資料寫入Stud_EnterTemp中---start
        If Request("serial") <> "" Then
            str = ""
            str += " SELECT ModifyDate,SETID from Stud_EnterTemp where SETID='" & Request("serial") & "'"
            str += " ORDER BY  ModifyDate desc"

        Else
            str = ""
            str += " SELECT 'X' x from Stud_EnterTemp where IDNO='" & Request("IDNO") & "'"
        End If
        table = DbAccess.GetDataTable(str, objada, objconn)

        If table.Rows.Count = 0 Then
            '如未有此學員的報名資料.則新增一筆報名學員資料
            dr = table.NewRow()
            'objstr = "select Max(SETID)+1 as MaxID from Stud_EnterTemp"
            'objtable = DbAccess.GetDataTable(objstr)
            'sqldr = objtable.Rows(0)
            'SETID = sqldr("MaxID")
            Call TIMS.OpenDbConn(objconn)
            SETID = DbAccess.GetNewId(objconn, "STUD_ENTERTEMP_SETID_SEQ,STUD_ENTERTEMP,SETID")
            dr("SETID") = SETID  'sqldr("MaxID")
            i = 0
        Else
            dr = table.Rows(0)
            SETID = dr("SETID")
            i = 1
        End If

        dr("IDNO") = TIMS.ChangeIDNO(Me.idno.Text)
        dr("Name") = Me.name.Text
        If Me.sex.Items(0).Selected = True Then
            dr("Sex") = "M"
        Else
            dr("Sex") = "F"
        End If
        dr("Birthday") = CDate(Me.birthday.Text)
        If Me.passportno.Items(0).Selected = True Then
            dr("PassPortNo") = 1
        Else
            dr("PassPortNo") = 2
        End If

        Select Case Me.MaritalStatus.SelectedValue
            Case "1", "2"
                dr("MaritalStatus") = Me.MaritalStatus.SelectedValue
            Case Else
                dr("MaritalStatus") = Convert.DBNull
        End Select

        dr("DegreeID") = Me.degreeid.SelectedValue
        dr("GradID") = Me.graduatestatus.SelectedValue
        dr("School") = Me.school.Text
        dr("Department") = Me.department.Text
        dr("MilitaryID") = IIf(MilitaryID.SelectedValue = "", Convert.DBNull, MilitaryID.SelectedValue) 'Me.MilitaryID.SelectedValue
        dr("ZipCode") = Me.zipcode1.Value
        dr("Address") = Me.address.Text
        dr("Phone1") = Me.phoned.Text
        If Trim(Me.phonen.Text) <> "" Then Me.phonen.Text = Trim(Me.phonen.Text) Else Me.phonen.Text = ""
        If Trim(Me.cellphone.Text) <> "" Then Me.cellphone.Text = Trim(Me.cellphone.Text) Else Me.cellphone.Text = ""
        dr("Phone2") = Me.phonen.Text
        dr("CellPhone") = Me.cellphone.Text
        dr("Email") = ""
        If Me.email.Text <> "" Then
            dr("Email") = Me.email.Text
        End If
        dr("IsAgree") = "N"
        If Me.isagree.Items(0).Selected = True Then
            dr("IsAgree") = "Y"
        End If
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now()
        If i = 0 Then
            table.Rows.Add(dr)
        End If
        DbAccess.UpdateDataTable(table, objada)
        '---end

        '線上報名資料寫入Stud_EnterType中---start
        '新增報名資料
        objada = New OracleDataAdapter
        If Request("proecess") = "add" Then
            str = "select * from Stud_EnterType WHERE 1<>1"
            table = DbAccess.GetDataTable(str, objada, objconn)
            dr = table.NewRow()
            dr("SETID") = SETID
            EnterDate = FormatDateTime(Now(), DateFormat.ShortDate)
            dr("EnterDate") = FormatDateTime(Now(), DateFormat.ShortDate)

            objstr = "select Max(SerNum) as maxnum from Stud_EnterType where SETID = '" & SETID & "' and EnterDate =  " & TIMS.to_date(FormatDateTime(Now(), DateFormat.ShortDate))
            sqldr = DbAccess.GetOneRow(objstr)
            If Not sqldr Is Nothing Then
                If IsDBNull(sqldr("maxnum")) Then
                    dr("SerNum") = 1
                    SerNum = 1
                Else
                    dr("SerNum") = sqldr("maxnum") + 1
                    SerNum = sqldr("maxnum") + 1
                End If
            End If

            Dim ExamNo1 As String '取出班級的CLASSID +期別 成為准考證編碼的前面的固定碼
            Dim NewExamNO As String
            Dim ExamOcid1 As String = OCIDValue1.Value
            Dim ExamPlanID As String = sm.UserInfo.PlanID

            '取出准考證號-------------------------------------------Start
            ExamNo1 = TIMS.Get_ExamNo1(ExamOcid1)
            If ExamNo1 = "" Then '防呆
                Common.MessageBox(Me, "班級的代號 與期別有誤，請確認班級狀態")
                Exit Sub
            Else
                If ExamPlanID <> "" Then
                    NewExamNO = TIMS.Get_NewExamNO(ExamPlanID, ExamNo1, ExamOcid1)
                End If
            End If
            If NewExamNO = "" Then
                Common.MessageBox(Me, "班級的代號 與計畫不符，請確認班級狀態(取出准考證號)")
                Exit Sub
            End If
            '取出准考證號-------------------------------------------End


            ''------取出准考證號碼最大號-Start
            'Dim Exam, ExamNoStr As String
            'If ExamNo.Value = "" Then
            '    sql = "SELECT b.ClassID+a.CyclType as Exam FROM "
            '    sql += "(SELECT * FROM Class_ClassInfo WHERE OCID='" & OCIDValue1.Value & "') a "
            '    sql += "Join ID_Class b ON a.CLSID=b.CLSID"

            '    sqldr = DbAccess.GetOneRow(sql)

            '    If sqldr Is Nothing Then
            '        Common.RespWrite(Me, "<script language=javascript>window.alert('查無此班級的序號')</script>")
            '        Exit Sub
            '    Else
            '        Exam = sqldr("Exam").ToString
            '        sql = "select ExamNo from Stud_EnterType WITH (TABLOCK) WHERE OCID1='" & OCIDValue1.Value & "'"
            '        objtable = DbAccess.GetDataTable(sql)
            '        If objtable.Select("ExamNo Like '" & Exam & "%'").Length = 0 Then
            '            ExamNoStr = Exam & "001"
            '        Else
            '            sqldr = objtable.Select("ExamNo Like '" & Exam & "%'", "ExamNo desc")(0)
            '            Dim num As Integer = Int(Right(sqldr("ExamNo"), 3)) + 1
            '            If num >= 100 Then
            '                ExamNoStr = Exam & num
            '            ElseIf num < 100 And num >= 10 Then
            '                ExamNoStr = Exam & "0" & num
            '            ElseIf num < 10 Then
            '                ExamNoStr = Exam & "00" & num
            '            End If
            '        End If
            '    End If
            'Else
            '    ExamNoStr = ExamNo.Value
            'End If

            dr("ExamNo") = NewExamNO 'ExamNoStr
            dr("EnterChannel") = 2
            dr("EnterPath") = "P" '個人報名基本資料送出
            dr("RID") = sm.UserInfo.RID
            dr("PlanID") = sm.UserInfo.PlanID
        End If
        '修改報名資料
        If Request("proecess") = "edit" Then
            'Dim tdEnterDate As String = TIMS.to_date(Request("EnterDate"))
            'str = "select "
            'str += " Q2_3,Q2_4,dbo.SUBSTR(Q2_5OTHER, 1, 4000) Q2_5OTHER,MODIFYACCT,  " & vbCrLf
            'str += " MODIFYDATE,TICKET_NO,RELENTERDATE,NOTEXAM,CCLID,ESETID,ESERNUM,  " & vbCrLf
            'str += " TRANSDATE,SEID,SUPPLYID,BUDID,ENTERPATH,HIGHEDUBG,WORKSUPPIDENT,  " & vbCrLf
            'str += " USERNOSHOW,NOTES,PRIORWORKTYPE1,PRIORWORKORG1,SOFFICEYM1,FOFFICEYM1,  " & vbCrLf
            'str += " ACTNO,SETID,ENTERDATE,SERNUM,EXAMNO,OCID1,TMID1,OCID2,TMID2,  " & vbCrLf
            'str += " OCID3,TMID3,WRITERESULT,ORALRESULT,TOTALRESULT,ENTERCHANNEL,  " & vbCrLf
            'str += " IDENTITYID,RID,PLANID,TRNDMODE,TRNDTYPE,Q1_1,Q1_2,Q1_2OTHER,  " & vbCrLf
            'str += " Q1_3,Q1_3OTHER,Q1_4,Q1_4OTHER,Q1_5  " & vbCrLf
            str = ""
            str += " SELECT *" & vbCrLf
            str += " from Stud_EnterType where SETID='" & SETID & "' and EnterDate= " & TIMS.to_date(Request("EnterDate")) & " and SerNum='" & Request("SerNum") & "'"
            table = DbAccess.GetDataTable(str, objada, objconn)
            dr = table.Rows(0)
            If Not IsDBNull(dr("eSerNum")) Then
                eSerNum = dr("eSerNum")
            End If
            If Not IsDBNull(dr("SEID")) Then
                SEID = dr("SEID")
            End If
            EnterDate = Request("EnterDate")
            SerNum = Request("SerNum")
        End If

        '--------檢查之前的志願是否有相衝-Start
        Dim sql As String = ""
        Dim dt As DataTable
        If Request("proecess") = "edit" Then
            'sql = "SELECT"
            'sql += " Q2_3,Q2_4,dbo.SUBSTR(Q2_5OTHER, 1, 4000) Q2_5OTHER,MODIFYACCT,  " & vbCrLf
            'sql += " MODIFYDATE,TICKET_NO,RELENTERDATE,NOTEXAM,CCLID,ESETID,ESERNUM,  " & vbCrLf
            'sql += " TRANSDATE,SEID,SUPPLYID,BUDID,ENTERPATH,HIGHEDUBG,WORKSUPPIDENT,  " & vbCrLf
            'sql += " USERNOSHOW,NOTES,PRIORWORKTYPE1,PRIORWORKORG1,SOFFICEYM1,FOFFICEYM1,  " & vbCrLf
            'sql += " ACTNO,SETID,ENTERDATE,SERNUM,EXAMNO,OCID1,TMID1,OCID2,TMID2,  " & vbCrLf
            'sql += " OCID3,TMID3,WRITERESULT,ORALRESULT,TOTALRESULT,ENTERCHANNEL,  " & vbCrLf
            'sql += " IDENTITYID,RID,PLANID,TRNDMODE,TRNDTYPE,Q1_1,Q1_2,Q1_2OTHER,  " & vbCrLf
            'sql += " Q1_3,Q1_3OTHER,Q1_4,Q1_4OTHER,Q1_5  " & vbCrLf
            sql = ""
            sql += " SELECT *" & vbCrLf
            sql += " FROM Stud_EnterType WHERE SETID='" & SETID & "' and EnterDate<> " & TIMS.to_date(Request("EnterDate")) & " and SerNum<>'" & Request("SerNum") & "'"
        Else
            'sql = "SELECT "
            'sql += " Q2_3,Q2_4,dbo.SUBSTR(Q2_5OTHER, 1, 4000) Q2_5OTHER,MODIFYACCT,  " & vbCrLf
            'sql += " MODIFYDATE,TICKET_NO,RELENTERDATE,NOTEXAM,CCLID,ESETID,ESERNUM,  " & vbCrLf
            'sql += " TRANSDATE,SEID,SUPPLYID,BUDID,ENTERPATH,HIGHEDUBG,WORKSUPPIDENT,  " & vbCrLf
            'sql += " USERNOSHOW,NOTES,PRIORWORKTYPE1,PRIORWORKORG1,SOFFICEYM1,FOFFICEYM1,  " & vbCrLf
            'sql += " ACTNO,SETID,ENTERDATE,SERNUM,EXAMNO,OCID1,TMID1,OCID2,TMID2,  " & vbCrLf
            'sql += " OCID3,TMID3,WRITERESULT,ORALRESULT,TOTALRESULT,ENTERCHANNEL,  " & vbCrLf
            'sql += " IDENTITYID,RID,PLANID,TRNDMODE,TRNDTYPE,Q1_1,Q1_2,Q1_2OTHER,  " & vbCrLf
            'sql += " Q1_3,Q1_3OTHER,Q1_4,Q1_4OTHER,Q1_5  " & vbCrLf
            sql = ""
            sql += " SELECT *" & vbCrLf
            sql += " FROM Stud_EnterType WHERE SETID='" & SETID & "'"
        End If
        dt = DbAccess.GetDataTable(sql, objconn)

        For Each sqldr In dt.Rows
            If Convert.ToString(sqldr("OCID1")) = OCIDValue1.Value Then
                Common.RespWrite(Me, "<script language=javascript>window.alert('" & OCID1.Text & "已報名過在志願中!')</script>")
                Exit Sub
            End If
        Next
        '--------檢查之前的志願是否有相衝-End
        For i = 0 To identityid.Items.Count - 1
            If identityid.Items(i).Selected = True Then
                If IdentityIDs <> "" Then IdentityIDs += ","
                IdentityIDs = IdentityIDs + identityid.Items(i).Value
            End If
        Next
        dr("IdentityID") = IdentityIDs
        dr("RelEnterDate") = Me.relenterdate.Text
        dr("OCID1") = Val(Me.OCIDValue1.Value)
        dr("TMID1") = Val(Me.TMIDValue1.Value)
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now()

        '線上報名資料寫入Stud_EnterTrain2中---start
        If SEID <> 0 Then
            str1 = "select * from Stud_EnterTrain2 where SEID = '" & SEID & "'"
            table1 = DbAccess.GetDataTable(str1, objada1, objconn)
            If table1.Rows.Count <> 0 Then
                dr1 = table1.Rows(0)
                CHSEID = 0
            Else
                dr1 = table1.NewRow()
                dr1("eSerNum") = 0
                CHSEID = 1
                SEID = 0
            End If
        ElseIf eSerNum <> 0 Then
            str1 = "select * from Stud_EnterTrain2 where eSerNum = '" & eSerNum & "'"
            table1 = DbAccess.GetDataTable(str1, objada1, objconn)
            If table1.Rows.Count <> 0 Then
                dr1 = table1.Rows(0)
                CHSEID = 0
                SEID = dr1("SEID")
            Else
                dr1 = table1.NewRow()
                dr1("eSerNum") = 0
                CHSEID = 1
            End If
        Else
            str1 = "select * from Stud_EnterTrain2 where 1<>1"
            table1 = DbAccess.GetDataTable(str1, objada1, objconn)
            dr1 = table1.NewRow()
            dr1("eSerNum") = 0
            CHSEID = 1
        End If

        If Me.zipcode2.Value <> "" Then
            dr1("ZipCode2") = Me.zipcode2.Value
            dr1("HouseholdAddress") = Me.householdaddress.Text
        Else
            If checkbox1.Checked = True Then
                dr1("ZipCode2") = Me.zipcode1.Value
                dr1("HouseholdAddress") = Me.address.Text
            End If
        End If
        If Me.midentityid.SelectedValue <> "" Then
            dr1("MidentityID") = Me.midentityid.SelectedValue
        Else
            dr1("MidentityID") = "01"
        End If
        If Me.handtypeid.SelectedValue <> "" Then
            dr1("HandTypeID") = Me.handtypeid.SelectedValue
        End If
        If Me.handlevelid.SelectedValue <> "" Then
            dr1("HandLevelID") = Me.handlevelid.SelectedValue
        End If
        If Me.priorworkorg1.Text <> "" Then
            dr1("PriorWorkOrg1") = Me.priorworkorg1.Text
        End If
        If Me.title1.Text <> "" Then
            dr1("Title1") = Me.title1.Text
        End If
        If Me.priorworkorg2.Text <> "" Then
            dr1("PriorWorkOrg2") = Me.priorworkorg2.Text
        End If
        If Me.title2.Text <> "" Then
            dr1("Title2") = Me.title2.Text
        End If
        If Me.sofficeym1.Text <> "" Then
            dr1("SOfficeYM1") = CDate(Me.sofficeym1.Text)
        End If
        If Me.fofficeym1.Text <> "" Then
            dr1("FOfficeYM1") = CDate(Me.fofficeym1.Text)
        End If
        If Me.sofficeym2.Text <> "" Then
            dr1("SOfficeYM2") = CDate(Me.sofficeym2.Text)
        End If
        If Me.fofficeym2.Text <> "" Then
            dr1("FOfficeYM2") = CDate(Me.fofficeym2.Text)
        End If
        If Me.priorworkpay.Text <> "" Then
            dr1("PriorWorkPay") = Me.priorworkpay.Text
        End If
        If Me.realjobless.Text <> "" Then
            dr1("RealJobless") = Me.realjobless.Text
        End If

        If Me.joblessid.SelectedValue <> "" Then
            dr1("JoblessID") = Me.joblessid.SelectedValue
        End If
        If Me.traffic.SelectedValue <> "" Then
            dr1("Traffic") = Me.traffic.SelectedValue
        End If
        If Me.showdetail.SelectedValue = "Y" Then
            dr1("ShowDetail") = Me.showdetail.SelectedValue
        Else
            dr1("ShowDetail") = "N"
        End If
        If Me.acctmode.Items(0).Selected = True Then '
            dr1("AcctMode") = 0
            dr1("PostNo") = Me.postno_1.Text + "-" + Me.postno_2.Text
            dr1("AcctNo") = Me.acctno1_1.Text + "-" + Me.acctno1_2.Text
        Else
            dr1("AcctMode") = 1
            dr1("AcctHeadNo") = Me.acctheadno.Text
            dr1("BankName") = Me.bankname.Text
            dr1("AcctNo") = Me.acctno2.Text
        End If
        If Me.firdate.Text <> "" Then
            dr1("FirDate") = CDate(Me.firdate.Text)
        End If
        If Me.uname.Text <> "" Then
            dr1("Uname") = Me.uname.Text
        End If
        If Me.intaxno.Text <> "" Then
            dr1("Intaxno") = Me.intaxno.Text
        End If
        If Me.servdept.Text <> "" Then
            dr1("ServDept") = Me.servdept.Text
        End If
        If Me.jobtitle.Text <> "" Then
            dr1("JobTitle") = Me.jobtitle.Text
        End If
        dr1("Zip") = Me.zip.Value
        dr1("Addr") = Me.addr.Text
        dr1("Tel") = Me.tel.Text
        If Me.fax.Text <> "" Then
            dr1("Fax") = Me.fax.Text
        End If
        If Me.sdate.Text <> "" Then
            dr1("SDate") = CDate(Me.sdate.Text)
        End If
        If Me.sjdate.Text <> "" Then
            dr1("SJDate") = CDate(Me.sjdate.Text)
        End If
        If Me.spdate.Text <> "" Then
            dr1("SPDate") = CDate(Me.spdate.Text)
        End If
        If Me.q1.Items(0).Selected = True Then
            dr1("Q1") = 1
        Else
            dr1("Q1") = 0
        End If
        If Me.q2.Items(0).Selected = True Then
            dr1("Q2_1") = 1
        Else
            dr1("Q2_1") = 2
        End If
        If Me.q2.Items(1).Selected = True Then
            dr1("Q2_2") = 1
        Else
            dr1("Q2_2") = 2
        End If
        If Me.q2.Items(2).Selected = True Then
            dr1("Q2_3") = 1
        Else
            dr1("Q2_3") = 2
        End If
        If Me.q2.Items(3).Selected = True Then
            dr1("Q2_4") = 1
        Else
            dr1("Q2_4") = 2
        End If
        If Me.q3.SelectedValue <> "" Then
            dr1("Q3") = Me.q3.SelectedValue
            If Me.q3.SelectedValue = "3" Then
                If Me.q3_other.Text <> "" Then
                    dr1("Q3_Other") = Me.q3_other.Text
                End If
            End If
        End If
        dr1("Q4") = Me.q4.SelectedValue
        If Me.q5.Items(0).Selected = True Then
            dr1("Q5") = 1
        Else
            dr1("Q5") = 0
        End If
        If Me.q61.Text <> "" Then
            dr1("Q61") = Me.q61.Text
        End If
        If Me.q62.Text <> "" Then
            dr1("Q62") = Me.q62.Text
        End If
        If Me.q63.Text <> "" Then
            dr1("Q63") = Me.q63.Text
        End If
        If Me.q64.Text <> "" Then
            dr1("Q64") = Me.q64.Text
        End If

        dr1("ModifyAcct") = sm.UserInfo.UserID
        dr1("ModifyDate") = Now()
        If CHSEID = 1 Then
            '新增
            table1.Rows.Add(dr1)
        End If
        DbAccess.UpdateDataTable(table1, objada1)
        '線上報名資料寫入Stud_EnterTrain2中---end

        If SEID = 0 Then
            str1 = "select Max(SEID) as MaxID from Stud_EnterTrain2"
            table1 = DbAccess.GetDataTable(str1)
            If table1.Rows.Count <> 0 Then
                dr1 = table1.Rows(0)
                SEID = dr1("MaxID")
            End If
        End If
        If SEID <> 0 Then
            dr("SEID") = SEID
        End If

        If Request("proecess") = "add" Then
            table.Rows.Add(dr)
        End If
        DbAccess.UpdateDataTable(table, objada)
        '線上報名資料寫入Stud_EnterType中----end

        '產學訓新增報名學員..直接代入甄試結果
        Dim blnAddFlag As Boolean = False
        Dim blnUpdateFlag As Boolean = False
        str = "SELECT * FROM Stud_SelResult where SETID='" & SETID & "' and EnterDate=convert(datetime, '" & EnterDate & "', 111) and SerNum='" & SerNum & "'"
        table = DbAccess.GetDataTable(str, objada, objconn)
        If table.Rows.Count = 0 Then
            '可新增
            If SerNum <> 0 Then blnAddFlag = True
        Else
            '可修改
            blnUpdateFlag = True
        End If
        '可新增 或 可修改
        If blnAddFlag OrElse blnUpdateFlag Then
            If blnAddFlag Then
                '可新增
                dr = table.NewRow()
                table.Rows.Add(dr)
                dr("SETID") = SETID
                dr("EnterDate") = EnterDate
                dr("SerNum") = SerNum
            Else
                '可修改
                dr = table.Rows(0)
            End If
            dr("OCID") = Val(Me.OCIDValue1.Value)

            dr("Admission") = "Y"
            dr("SelResultID") = "01"

            dr("RID") = sm.UserInfo.RID
            dr("PlanID") = sm.UserInfo.PlanID
            dr("ModifyAcct") = sm.UserInfo.UserID
            dr("ModifyDate") = Now
            DbAccess.UpdateDataTable(table, objada)
        End If

        '是否試算
        'str = "SELECT "
        'str += " NORID,OTHERREASON,ISBUSINESS,APPLIEDRESULTR,APPLIEDRESULTM,PNUM,   " & vbCrLf
        'str += " ISCONT,QAYSDATE,QAYFDATE,LASTSTATE,TADDRESSZIP2W,EVTA_NOSHOW,   " & vbCrLf
        'str += " ETRAIN_SHOW,ECOMMENT,COMPANYNAME,NOTICE,CJOB_UNKEY,EXAMPERIOD,   " & vbCrLf
        'str += " OCID,CLSID,PLANID,YEARS,CYCLTYPE,LEVELTYPE,RID,CLASSCNAME,CLASSENGNAME,   " & vbCrLf
        'str += " dbo.SUBSTR(CONTENT, 1, 4000) CONTENT,dbo.SUBSTR(PURPOSE, 1, 4000) PURPOSE,   " & vbCrLf
        'str += " TPROPERTYID,TMID,CLID,SENTERDATE,FENTERDATE,CHECKINDATE,EXAMDATE,STDATE,   " & vbCrLf
        'str += " FTDATE,TADDRESSZIP,TADDRESS,THOURS,TNUM,TDEADLINE,TPERIOD,NOTOPEN,   " & vbCrLf
        'str += " ISAPPLIC,RELSHIP,COMIDNO,SEQNO,ISCALCULATE,ISSUCCESS,CTNAME,MODIFYACCT,   " & vbCrLf
        'str += " MODIFYDATE,CLASSNUM,LEVELCOUNT,ISFULLDATE,CLASS_UNIT,ISCLOSED,BGTIME   " & vbCrLf
        'str += " FROM Class_ClassInfo WHERE OCID='" & Val(Me.ocidvalue1.Value) & "'"
        'table = DbAccess.GetDataTable(str, objadapter, objconn)
        'If table.Rows.Count <> 0 Then
        '    dr = table.Rows(0)
        '    dr("IsCalculate") = "Y"
        '    DbAccess.UpdateDataTable(table, objadapter)
        'End If
        '是否試算
        Call TIMS.Update_ClassInfoIsCalculate("Y", Val(Me.OCIDValue1.Value), objconn)

        Dim MsgBox As String
        Select Case Request("proecess")
            Case "add"
                MsgBox += "資料新增成功!\n"
            Case "edit"
                MsgBox += "資料更新成功!\n"
        End Select
        Common.RespWrite(Me, "<script language=javascript>window.alert('" & MsgBox & "');</script>")
        Common.RespWrite(Me, "<script language=javascript>window.location.href='SD_01_001.aspx?ID=" & Request("ID") & "';</script>")
    End Sub

    '輸入控制( Disabled/Visible/Enabled )
    Function SHOW_Disabled(Optional ByVal tmpON As Integer = 0)
        '只能看的時候就顯示灰階吧
        If tmpON = 0 Then
            trbasic1.Disabled = True
            imgrelenterdate.Visible = False
            imgbirthday.Visible = False
            degreeid.Enabled = False
            graduatestatus.Enabled = False
            MilitaryID.Enabled = False
            MaritalStatus.Enabled = False
            btnzipcode1.Visible = False
            btnzipcode2.Visible = False
            checkbox1.Visible = False
            midentityid.Enabled = False
            button5.Visible = False
            button2.Visible = False
            button3.Visible = False
            btnclear1.Visible = False
            btnclear2.Visible = False
            btnclear3.Visible = False
            imgsofficeym1.Visible = False
            imgsofficeym2.Visible = False
            imgfofficeym1.Visible = False
            imgfofficeym2.Visible = False
            joblessid.Enabled = False
            traffic.Enabled = False
            showdetail.Enabled = False

            imgfirdate.Visible = False
            btnzip5.Visible = False
            imgsdate.Visible = False
            imgsjdate.Visible = False
            imgspdate.Visible = False
            q4.Enabled = False
        Else
            trbasic1.Disabled = False
            imgrelenterdate.Visible = True
            imgbirthday.Visible = True
            degreeid.Enabled = True
            graduatestatus.Enabled = True
            MilitaryID.Enabled = True
            MaritalStatus.Enabled = True
            btnzipcode1.Visible = True
            btnzipcode2.Visible = True
            checkbox1.Visible = True
            midentityid.Enabled = True
            button5.Visible = True
            button2.Visible = True
            button3.Visible = True
            btnclear1.Visible = True
            btnclear2.Visible = True
            btnclear3.Visible = True
            imgsofficeym1.Visible = True
            imgsofficeym2.Visible = True
            imgfofficeym1.Visible = True
            imgfofficeym2.Visible = True
            joblessid.Enabled = True
            traffic.Enabled = True
            showdetail.Enabled = True

            imgfirdate.Visible = True
            btnzip5.Visible = True
            imgsdate.Visible = True
            imgsjdate.Visible = True
            imgspdate.Visible = True
            q4.Enabled = True
        End If

    End Function

    Function create()
        Dim sqldr As DataRow
        Dim objtable As DataTable
        Dim objstr As String = ""
        Dim Table As DataTable
        Dim str As String = ""
        Dim dt As DataTable
        Dim dr As DataRow

        Select Case Request("proecess")
            Case "add"
                '找出現場報名的學員資料(Stud_EnterTemp)
                create1()
                Me.idno.Text = TIMS.ChangeIDNO(Request("IDNO"))
            Case Else '"edit"
                If Request("proecess") <> "edit" Then
                    If Request("IDNO") <> "" Then
                        Me.ViewState("proecess") = "view"
                        Dim strScript As String = ""
                        strScript += "<script language=""javascript"" type=""text/javascript"">"
                        strScript += "    top.document.title=""個人報名基本資料"";"
                        strScript += "</script> "
                        Page.RegisterStartupScript("document_title", strScript)
                    Else
                        trbasic1.Disabled = True
                        SHOW_Disabled()
                        send.Visible = False
                        button22.Visible = False
                        Common.MessageBox(Me, "資料有誤!!")
                        Exit Function
                    End If
                End If

                '找出現場報名的學員資料(Stud_EnterTemp)
                create1()

                If Me.ViewState("proecess") = "view" Then
                    '只能看的時候就顯示灰階吧
                    trbasic1.Disabled = True
                    SHOW_Disabled()
                    send.Visible = False
                    button22.Visible = False
                    Me.handtypeid.Enabled = False
                    Me.handlevelid.Enabled = False

                    str = "" & vbCrLf
                    str += " select  "
                    str += " s.SETID,s.IDNO,s.NAME,s.SEX,s.BIRTHDAY,s.PASSPORTNO,s.MARITALSTATUS,  " & vbCrLf
                    str += " s.DEGREEID,s.GRADID,s.SCHOOL,s.DEPARTMENT,s.MILITARYID,s.ZIPCODE,  " & vbCrLf
                    str += " s.ADDRESS,s.PHONE1,s.PHONE2,s.CELLPHONE,s.EMAIL,s.PASSWORD,  " & vbCrLf
                    'str += " dbms_lob.substr( s.NOTES, 4000, 1 ) NOTES,s.ISAGREE,s.LAINFLAG,  " & vbCrLf
                    str += " s.NOTES," & vbCrLf
                    str += " s.ISAGREE,s.LAINFLAG,  " & vbCrLf
                    str += " s.MODIFYACCT,s.MODIFYDATE,s.ESETID,s.ZIPCODE2W  " & vbCrLf
                    str += "  ,t.SETID,t.ENTERDATE,t.SERNUM,t.EXAMNO,t.OCID1,t.TMID1,t.OCID2,t.TMID2,t.OCID3,t.TMID3,  " & vbCrLf
                    str += " t.WRITERESULT,t.ORALRESULT,t.TOTALRESULT,t.ENTERCHANNEL,t.IDENTITYID,t.RID,t.PLANID,  " & vbCrLf
                    str += " t.TRNDMODE,t.TRNDTYPE,t.Q1_1,t.Q1_2,t.Q1_2OTHER,t.Q1_3,t.Q1_3OTHER,t.Q1_4,t.Q1_4OTHER,  " & vbCrLf
                    str += " t.Q1_5,t.Q2_3,t.Q2_4," & vbCrLf
                    'str += " dbms_lob.substr( t.Q2_5OTHER, 4000, 1 ) Q2_5OTHER,t.MODIFYACCT,  " & vbCrLf
                    str += " t.Q2_5OTHER, " & vbCrLf
                    str += " t.MODIFYACCT, " & vbCrLf
                    str += " t.MODIFYDATE,t.TICKET_NO,t.RELENTERDATE,t.NOTEXAM,t.CCLID,t.ESETID,t.ESERNUM,t.TRANSDATE,  " & vbCrLf
                    str += " t.SEID,t.SUPPLYID,t.BUDID,t.ENTERPATH,t.HIGHEDUBG,t.WORKSUPPIDENT,t.USERNOSHOW,t.NOTES,  " & vbCrLf
                    str += " t.PRIORWORKTYPE1,t.PRIORWORKORG1,t.SOFFICEYM1,t.FOFFICEYM1,t.ACTNO,t.WSID  " & vbCrLf
                    str += " from Stud_EnterTemp s" & vbCrLf
                    str += " join Stud_EnterType t on s.SETID=t.SETID" & vbCrLf
                    str += " where 1=1" & vbCrLf
                    If Request("IDNO") <> "" Then
                        str += " and s.IDNO ='" & Request("IDNO") & "'" & vbCrLf
                    End If
                    If Request("OCID") <> "" Then
                        str += " and t.ocid1 ='" & Request("OCID") & "'" & vbCrLf
                    End If
                    str += " " & vbCrLf
                Else
                    'str = "select * from (select * from Stud_EnterTemp where SETID='" & Request("serial") & "') s,(select * from Stud_EnterType where EnterDate='" & Request("EnterDate") & "' and SerNum='" & Request("SerNum") & "') t where s.SETID=t.SETID" '★
                    'str = "select * from (select "
                    'str += " SETID,IDNO,NAME,SEX,BIRTHDAY,PASSPORTNO,MARITALSTATUS,  " & vbCrLf
                    'str += " DEGREEID,GRADID,SCHOOL,DEPARTMENT,MILITARYID,ZIPCODE,  " & vbCrLf
                    'str += " ADDRESS,PHONE1,PHONE2,CELLPHONE,EMAIL,PASSWORD,  " & vbCrLf
                    'str += " dbo.SUBSTR(NOTES, 1, 4000) NOTES,ISAGREE,LAINFLAG,  " & vbCrLf
                    'str += " MODIFYACCT,MODIFYDATE,ESETID,ZIPCODE2W  " & vbCrLf
                    'str += " from Stud_EnterTemp where SETID='" & Request("serial") & "') s,( " & vbCrLf
                    'str += " select  Q2_3,Q2_4,dbo.SUBSTR(Q2_5OTHER, 1, 4000) Q2_5OTHER,MODIFYACCT,  " & vbCrLf
                    'str += " MODIFYDATE,TICKET_NO,RELENTERDATE,NOTEXAM,CCLID,ESETID,ESERNUM,  " & vbCrLf
                    'str += " TRANSDATE,SEID,SUPPLYID,BUDID,ENTERPATH,HIGHEDUBG,WORKSUPPIDENT,  " & vbCrLf
                    'str += " USERNOSHOW,NOTES,PRIORWORKTYPE1,PRIORWORKORG1,SOFFICEYM1,FOFFICEYM1,  " & vbCrLf
                    'str += " ACTNO,SETID,ENTERDATE,SERNUM,EXAMNO,OCID1,TMID1,OCID2,TMID2,  " & vbCrLf
                    'str += " OCID3,TMID3,WRITERESULT,ORALRESULT,TOTALRESULT,ENTERCHANNEL,  " & vbCrLf
                    'str += " IDENTITYID,RID,PLANID,TRNDMODE,TRNDTYPE,Q1_1,Q1_2,Q1_2OTHER,  " & vbCrLf
                    'str += " Q1_3,Q1_3OTHER,Q1_4,Q1_4OTHER,Q1_5  " & vbCrLf

                    str = ""
                    str += " SELECT *" & vbCrLf
                    str += " from Stud_EnterType where EnterDate= " & TIMS.to_date(Request("EnterDate")) & " and SerNum='" & Request("SerNum") & "') t where s.SETID=t.SETID"
                End If

                Table = DbAccess.GetDataTable(str, objconn)
                If Table.Rows.Count <> 0 Then
                    dr = Table.Rows(0)
                    '代入報名職類資料
                    Me.relenterdate.Text = dr("RelEnterDate")
                    If Not IsDBNull(dr("IsAgree")) Then
                        If dr("IsAgree") = "Y" Then
                            Me.isagree.Items(0).Selected = True
                        Else
                            Me.isagree.Items(1).Selected = True
                        End If
                    End If
                    Dim all As Array
                    all = Split(dr("IdentityID"), ",", , CompareMethod.Text)
                    For i As Integer = 0 To all.Length - 1
                        For j As Integer = 0 To identityid.Items.Count - 1
                            If all(i) = identityid.Items(j).Value Then
                                identityid.Items(j).Selected = True
                            End If
                        Next
                    Next
                    Me.TMIDValue1.Value = dr("TMID1")
                    Me.OCIDValue1.Value = dr("OCID1")
                    'objstr = "select CONVERT(VARCHAR,STDATE,111) STDATE,ClassCName + '第' + CyclType + '期'as ClassCName from Class_ClassInfo where OCID = '" & dr("OCID1") & "'" '★
                    objstr = "select CONVERT(char, STDATE, 111) STDATE,ClassCName + '第' + CyclType + '期'as ClassCName from Class_ClassInfo where OCID = '" & dr("OCID1") & "'"
                    objtable = DbAccess.GetDataTable(objstr)
                    If objtable.Rows.Count <> 0 Then
                        sqldr = objtable.Rows(0)
                        Me.OCID1.Text = sqldr("ClassCName")
                        lociddate.Value = Convert.ToString(sqldr("STDATE"))
                    End If
                    'objstr = "select '[' + TrainID + ']' + TrainName as TrainName from Key_TrainType where TMID = '" & dr("TMID1") & "'"
                    objstr = "" & vbCrLf
                    'objstr += " select '[' + ISNULL(TrainID,JobID) + ']' + ISNULL(TrainName,JobName) as TrainName " & vbCrLf '★
                    objstr += " select '[' + dbo.NVL(TrainID,JobID) + ']' + dbo.NVL(TrainName,JobName) as TrainName " & vbCrLf
                    objstr += " from Key_TrainType" & vbCrLf
                    objstr += " where TMID = '" & dr("TMID1") & "'" & vbCrLf
                    objtable = DbAccess.GetDataTable(objstr)
                    If objtable.Rows.Count <> 0 Then
                        sqldr = objtable.Rows(0)
                        If Not IsDBNull(sqldr("TrainName")) Then
                            Me.TMID1.Text = sqldr("TrainName")
                        End If
                    End If
                    '代入線上報名資料(產訓資料)
                    Dim objstr1 As String
                    Dim objtable1 As DataTable
                    Dim sqldr1 As DataRow
                    If Not IsDBNull(dr("SEID")) Then
                        objstr = "select * from Stud_EnterTrain2 where SEID = '" & dr("SEID") & "'"
                    ElseIf IsDBNull(dr("eSerNum")) Then
                        objstr = "select * from Stud_EnterTrain2 where eSerNum = '" & dr("eSerNum") & "'"
                    Else
                        objstr = ""
                    End If
                    If objstr <> "" Then
                        objtable = DbAccess.GetDataTable(objstr)
                        If objtable.Rows.Count <> 0 Then
                            sqldr = objtable.Rows(0)
                            If Not IsDBNull(sqldr("ZipCode2")) Then
                                objstr1 = "select ic.CTName,iz.ZipName from ID_ZIP iz JOIN ID_City ic ON ic.CTID = iz.CTID where iz.ZipCode = '" & sqldr("ZipCode2") & "'"
                                objtable1 = DbAccess.GetDataTable(objstr1)
                                If objtable1.Rows.Count <> 0 Then
                                    sqldr1 = objtable1.Rows(0)
                                    Me.city2.Text = "(" & sqldr("ZipCode2") & ")" & sqldr1("CTName") & sqldr1("ZipName")
                                    Me.zipcode2.Value = sqldr("ZipCode2")
                                    Me.householdaddress.Text = sqldr("HouseholdAddress")
                                End If
                            End If

                            objstr1 = "SELECT * FROM Key_Identity where IdentityID = '" & sqldr("MIdentityID") & "'"
                            objtable1 = DbAccess.GetDataTable(objstr1)
                            sqldr1 = objtable1.Rows(0)
                            Me.midentityid.Items.Insert(0, New ListItem(sqldr1("Name"), sqldr1("IdentityID")))
                            Me.midentityid.Items.Insert(1, New ListItem("請選擇", ""))
                            If Me.identityid.Items(4).Selected = True Then
                                If Not IsDBNull(sqldr("HandTypeID")) Then
                                    objstr1 = "SELECT * FROM Key_HandicatType where HandTypeID = '" & sqldr("HandTypeID") & "'"
                                    objtable1 = DbAccess.GetDataTable(objstr1)
                                    sqldr1 = objtable1.Rows(0)
                                    Me.handtypeid.Items.Insert(0, New ListItem(sqldr1("Name"), sqldr1("HandTypeID")))
                                    Me.handtypeid.Items.Insert(1, New ListItem("請選擇", ""))
                                Else
                                    Me.handtypeid.Items.Insert(0, New ListItem("請選擇", ""))
                                End If

                                If Not IsDBNull(sqldr("HandLevelID")) Then
                                    objstr1 = "SELECT * FROM Key_HandicatLevel where HandLevelID = '" & sqldr("HandLevelID") & "'"
                                    objtable1 = DbAccess.GetDataTable(objstr1)
                                    sqldr1 = objtable1.Rows(0)
                                    Me.handlevelid.Items.Insert(0, New ListItem(sqldr1("Name"), sqldr1("HandLevelID")))
                                    Me.handlevelid.Items.Insert(1, New ListItem("請選擇", ""))
                                Else
                                    Me.handlevelid.Items.Insert(0, New ListItem("請選擇", ""))
                                End If
                            Else
                                Me.handtypeid.Items.Insert(0, New ListItem("請選擇", ""))
                                Me.handlevelid.Items.Insert(0, New ListItem("請選擇", ""))
                            End If

                            If Not IsDBNull(sqldr("PriorWorkOrg1")) Then
                                Me.priorworkorg1.Text = sqldr("PriorWorkOrg1")
                            End If
                            If Not IsDBNull(sqldr("Title1")) Then
                                Me.title1.Text = sqldr("Title1")
                            End If
                            If Not IsDBNull(sqldr("PriorWorkOrg2")) Then
                                Me.priorworkorg2.Text = sqldr("PriorWorkOrg2")
                            End If
                            If Not IsDBNull(sqldr("Title2")) Then
                                Me.title2.Text = sqldr("Title2")
                            End If
                            If Not IsDBNull(sqldr("SOfficeYM1")) Then
                                Me.sofficeym1.Text = sqldr("SOfficeYM1")
                            End If
                            If Not IsDBNull(sqldr("FOfficeYM1")) Then
                                Me.fofficeym1.Text = sqldr("FOfficeYM1")
                            End If
                            If Not IsDBNull(sqldr("SOfficeYM2")) Then
                                Me.sofficeym2.Text = sqldr("SOfficeYM2")
                            End If
                            If Not IsDBNull(sqldr("FOfficeYM2")) Then
                                Me.fofficeym2.Text = sqldr("FOfficeYM2")
                            End If
                            If Not IsDBNull(sqldr("PriorWorkPay")) Then
                                Me.priorworkpay.Text = sqldr("PriorWorkPay")
                            End If
                            If Not IsDBNull(sqldr("RealJobless")) Then
                                Me.realjobless.Text = sqldr("RealJobless")
                            End If
                            If Not IsDBNull(sqldr("JoblessID")) Then
                                objstr1 = "SELECT * FROM Key_JoblessWeek where JoblessID = '" & sqldr("JoblessID") & "'"
                                objtable1 = DbAccess.GetDataTable(objstr1)
                                sqldr1 = objtable1.Rows(0)
                                Me.joblessid.Items.Insert(0, New ListItem(sqldr1("Name"), sqldr1("JoblessID")))
                                Me.joblessid.Items.Insert(1, New ListItem("請選擇", ""))
                            Else
                                Me.joblessid.Items.Insert(0, New ListItem("請選擇", ""))
                            End If
                            If Not IsDBNull(sqldr("Traffic")) Then
                                Me.traffic.SelectedValue = sqldr("Traffic")
                            End If
                            If Not IsDBNull(sqldr("ShowDetail")) Then
                                If sqldr("ShowDetail").ToString <> "" Then
                                    Common.SetListItem(showdetail, sqldr("ShowDetail"))
                                End If
                                'Me.ShowDetail.SelectedValue = sqldr("ShowDetail")
                            End If
                            If Not IsDBNull(sqldr("AcctMode")) Then
                                If sqldr("AcctMode") = "0" Then
                                    Me.acctmode.Items(0).Selected = True
                                    If Not IsDBNull(sqldr("PostNo")) Then
                                        If sqldr("PostNo").ToString.IndexOf("-") = -1 Then
                                            Me.postno_1.Text = sqldr("PostNo").ToString
                                        Else
                                            Me.postno_1.Text = Left(sqldr("PostNo").ToString, sqldr("PostNo").ToString.IndexOf("-"))
                                            Me.postno_2.Text = Right(sqldr("PostNo").ToString, sqldr("PostNo").ToString.Length - sqldr("PostNo").ToString.IndexOf("-") - 1)
                                        End If
                                    End If
                                    If Not IsDBNull(sqldr("AcctNo")) Then
                                        If sqldr("AcctNo").ToString.IndexOf("-") = -1 Then
                                            Me.acctno1_1.Text = sqldr("AcctNo").ToString
                                        Else
                                            Me.acctno1_1.Text = Left(sqldr("AcctNo").ToString, sqldr("AcctNo").ToString.IndexOf("-"))
                                            Me.acctno1_2.Text = Right(sqldr("AcctNo").ToString, sqldr("AcctNo").ToString.Length - sqldr("AcctNo").ToString.IndexOf("-") - 1)
                                        End If
                                    End If
                                Else
                                    Me.acctmode.Items(1).Selected = True
                                    If Not IsDBNull(sqldr("BanKName")) Then
                                        Me.bankname.Text = sqldr("BanKName")
                                    End If
                                    If Not IsDBNull(sqldr("AcctHeadNo")) Then
                                        Me.acctheadno.Text = sqldr("AcctHeadNo")
                                    End If
                                    If Not IsDBNull(sqldr("AcctNo")) Then
                                        Me.acctno2.Text = sqldr("AcctNo")
                                    End If
                                End If
                            End If

                            If Not IsDBNull(sqldr("FirDate")) Then
                                Me.firdate.Text = sqldr("FirDate")
                            End If
                            If Not IsDBNull(sqldr("Uname")) Then
                                Me.uname.Text = sqldr("Uname")
                            End If
                            If Not IsDBNull(sqldr("Intaxno")) Then
                                Me.intaxno.Text = sqldr("Intaxno")
                            End If
                            If Not IsDBNull(sqldr("ServDept")) Then
                                Me.servdept.Text = sqldr("ServDept")
                            End If
                            If Not IsDBNull(sqldr("JobTitle")) Then
                                Me.jobtitle.Text = sqldr("JobTitle")
                            End If
                            If Not IsDBNull(sqldr("Zip")) Then
                                objstr1 = "select ic.CTName,iz.ZipName from ID_ZIP iz JOIN ID_City ic ON ic.CTID = iz.CTID where iz.ZipCode = '" & sqldr("Zip") & "'"
                                objtable1 = DbAccess.GetDataTable(objstr1)
                                If objtable1.Rows.Count <> 0 Then
                                    sqldr1 = objtable1.Rows(0)
                                    Me.city5.Text = "(" & sqldr("Zip") & ")" & sqldr1("CTName") & sqldr1("ZipName")
                                    Me.zip.Value = sqldr("Zip")
                                    Me.addr.Text = sqldr("Addr")
                                End If
                            End If
                            If Not IsDBNull(sqldr("Tel")) Then
                                Me.tel.Text = sqldr("Tel")
                            End If
                            If Not IsDBNull(sqldr("Fax")) Then
                                Me.fax.Text = sqldr("Fax")
                            End If
                            If Not IsDBNull(sqldr("SDate")) Then
                                Me.sdate.Text = sqldr("SDate")
                            End If
                            If Not IsDBNull(sqldr("SJDate")) Then
                                Me.sjdate.Text = sqldr("SJDate")
                            End If
                            If Not IsDBNull(sqldr("SPDate")) Then
                                Me.spdate.Text = sqldr("SPDate")
                            End If
                            If Not IsDBNull(sqldr("Q1")) Then
                                If sqldr("Q1") Then
                                    Me.q1.SelectedIndex = 0
                                Else
                                    Me.q1.SelectedIndex = 1
                                End If
                            End If
                            If Not IsDBNull(sqldr("Q2_1")) Then
                                If sqldr("Q2_1") = 1 Then
                                    Me.q2.Items(0).Selected = True
                                End If
                            End If
                            If Not IsDBNull(sqldr("Q2_2")) Then
                                If sqldr("Q2_2") = 1 Then
                                    Me.q2.Items(1).Selected = True
                                End If
                            End If
                            If Not IsDBNull(sqldr("Q2_3")) Then
                                If sqldr("Q2_3") = 1 Then
                                    Me.q2.Items(2).Selected = True
                                End If
                            End If
                            If Not IsDBNull(sqldr("Q2_4")) Then
                                If sqldr("Q2_4") = 1 Then
                                    Me.q2.Items(3).Selected = True
                                End If
                            End If
                            If Not IsDBNull(sqldr("Q3")) Then
                                If sqldr("Q3") = 1 Then
                                    Me.q3.Items(0).Selected = True
                                End If
                                If sqldr("Q3") = 2 Then
                                    Me.q3.Items(1).Selected = True
                                End If
                                If sqldr("Q3") = 3 Then
                                    Me.q3.Items(2).Selected = True
                                    If Not IsDBNull(sqldr("Q3_Other")) Then
                                        Me.q3_other.Text = sqldr("Q3_Other")
                                        Me.q3_other.Enabled = True
                                    End If
                                End If
                            End If
                            If Not IsDBNull(sqldr("Q4")) Then
                                objstr1 = "SELECT TradeID,'[' + TradeID + ']' + TradeName as TradeName FROM Key_Trade where TradeID = '" & sqldr("Q4") & "'"
                                objtable1 = DbAccess.GetDataTable(objstr1)
                                sqldr1 = objtable1.Rows(0)
                                Me.q4.Items.Insert(0, New ListItem(sqldr1("TradeName"), sqldr1("TradeID")))
                                Me.q4.Items.Insert(1, New ListItem("請選擇", ""))
                            Else
                                Me.q4.Items.Insert(0, New ListItem("請選擇", ""))
                            End If
                            If Not IsDBNull(sqldr("Q5")) Then
                                If sqldr("Q5") Then
                                    Me.q5.SelectedIndex = 0
                                Else
                                    Me.q5.SelectedIndex = 1
                                End If
                            End If
                            If Not IsDBNull(sqldr("Q61")) Then
                                Me.q61.Text = sqldr("Q61")
                            End If
                            If Not IsDBNull(sqldr("Q62")) Then
                                Me.q62.Text = sqldr("Q62")
                            End If
                            If Not IsDBNull(sqldr("Q63")) Then
                                Me.q63.Text = sqldr("Q63")
                            End If
                            If Not IsDBNull(sqldr("Q64")) Then
                                Me.q64.Text = sqldr("Q64")
                            End If
                        Else
                            Me.handtypeid.Items.Insert(0, New ListItem("請選擇", ""))
                            Me.handlevelid.Items.Insert(0, New ListItem("請選擇", ""))
                            Me.midentityid.Items.Insert(0, New ListItem("請選擇", ""))
                            Me.joblessid.Items.Insert(0, New ListItem("請選擇", ""))
                            Me.q4.Items.Insert(0, New ListItem("請選擇", ""))
                        End If
                    Else
                        Me.handtypeid.Items.Insert(0, New ListItem("請選擇", ""))
                        Me.handlevelid.Items.Insert(0, New ListItem("請選擇", ""))
                        Me.midentityid.Items.Insert(0, New ListItem("請選擇", ""))
                        Me.joblessid.Items.Insert(0, New ListItem("請選擇", ""))
                        Me.q4.Items.Insert(0, New ListItem("請選擇", ""))
                    End If
                Else
                    Me.handtypeid.Items.Insert(0, New ListItem("請選擇", ""))
                    Me.handlevelid.Items.Insert(0, New ListItem("請選擇", ""))
                    Me.midentityid.Items.Insert(0, New ListItem("請選擇", ""))
                    Me.joblessid.Items.Insert(0, New ListItem("請選擇", ""))
                    Me.q4.Items.Insert(0, New ListItem("請選擇", ""))
                End If
        End Select
    End Function

    '找出現場報名的學員資料(Stud_EnterTemp)
    Function create1()
        Dim sqldr As DataRow
        Dim Table As DataTable
        Dim objtable As DataTable

        Dim Str As String = ""
        Dim objstr As String = ""
        Dim dr1 As DataRow

        '檢查是否曾經報名過.如有.基本資料直接代入欄位中
        If Request("serial") <> "" Then
            'Str = "select "
            'Str += " SETID,IDNO,NAME,SEX,BIRTHDAY,PASSPORTNO,MARITALSTATUS,  " & vbCrLf
            'Str += " DEGREEID,GRADID,SCHOOL,DEPARTMENT,MILITARYID,ZIPCODE,  " & vbCrLf
            'Str += " ADDRESS,PHONE1,PHONE2,CELLPHONE,EMAIL,PASSWORD,  " & vbCrLf
            'Str += " dbo.SUBSTR(NOTES, 1, 4000) NOTES,ISAGREE,LAINFLAG,  " & vbCrLf
            'Str += " MODIFYACCT,MODIFYDATE,ESETID,ZIPCODE2W  " & vbCrLf

            Str = ""
            Str += " SELECT *" & vbCrLf
            Str += " from Stud_EnterTemp where SETID='" & Request("serial") & "'"
            Str += " order by ModifyDate desc"
        ElseIf Request("IDNO") <> "" Then
            'Str = "select "
            'Str += " SETID,IDNO,NAME,SEX,BIRTHDAY,PASSPORTNO,MARITALSTATUS,  " & vbCrLf
            'Str += " DEGREEID,GRADID,SCHOOL,DEPARTMENT,MILITARYID,ZIPCODE,  " & vbCrLf
            'Str += " ADDRESS,PHONE1,PHONE2,CELLPHONE,EMAIL,PASSWORD,  " & vbCrLf
            'Str += " dbo.SUBSTR(NOTES, 1, 4000) NOTES,ISAGREE,LAINFLAG,  " & vbCrLf
            'Str += " MODIFYACCT,MODIFYDATE,ESETID,ZIPCODE2W  " & vbCrLf

            Str = ""
            Str += " SELECT *" & vbCrLf
            Str += " from Stud_EnterTemp where IDNO='" & Request("IDNO") & "'"
            Str += " order by ModifyDate desc"
        Else
            Exit Function
        End If
        ''檢查是否曾經報名過.如有.基本資料直接代入欄位中
        'If Request("serial") <> "" Then
        '    Str = "select * from Stud_EnterTemp where SETID='" & Request("serial") & "'"
        '    Str = Str & " order by ModifyDate desc"
        'ElseIf Request("IDNO") <> "" Then
        '    Str = "select * from Stud_EnterTemp where IDNO='" & Request("IDNO") & "'"
        '    Str = Str & " order by ModifyDate desc"
        'Else
        '    Exit Function
        'End If
        Table = DbAccess.GetDataTable(Str)
        If Table.Rows.Count <> 0 Then
            sqldr = Table.Rows(0)
            '姓名
            Me.name.Text = sqldr("Name")
            '身分證
            Me.idno.Text = TIMS.ChangeIDNO(sqldr("IDNO"))
            '生日
            Me.birthday.Text = FormatDateTime(sqldr("Birthday"), 2)
            '學校名稱
            Me.school.Text = sqldr("School")
            '科系名稱
            Me.department.Text = sqldr("Department")
            '身分別
            If sqldr("PassPortNO") = 1 Then
                passportno.Items(0).Selected = True
            Else
                passportno.Items(1).Selected = True
            End If
            '性別
            If sqldr("Sex") = "M" Then
                sex.Items(0).Selected = True
            Else
                sex.Items(1).Selected = True
            End If

            '學歷
            If degreeid.Items.Count > 0 Then
                If sqldr("DegreeID").ToString <> "" Then
                    Common.SetListItem(degreeid, sqldr("DegreeID"))
                End If
            End If
            '兵役
            If MilitaryID.Items.Count > 0 Then
                If sqldr("MilitaryID").ToString <> "" Then
                    Common.SetListItem(MilitaryID, sqldr("MilitaryID"))
                End If
            End If
            '畢業狀況
            If graduatestatus.Items.Count > 0 Then
                If sqldr("GradID").ToString <> "" Then
                    Common.SetListItem(graduatestatus, sqldr("GradID"))
                End If
            End If

            ''學歷
            'objstr = "select * from Key_Degree where DegreeID = '" & sqldr("DegreeID") & "'"
            'objtable = DbAccess.GetDataTable(objstr)
            'dr1 = objtable.Rows(0)
            'Me.DegreeID.Items.Insert(0, New ListItem(dr1("Name"), dr1("DegreeID")))
            'Me.DegreeID.Items.Insert(1, New ListItem("請選擇", ""))
            ''兵役
            'objstr = "select * from Key_Military where MilitaryID = '" & sqldr("MilitaryID") & "'"
            'objtable = DbAccess.GetDataTable(objstr)
            'dr1 = objtable.Rows(0)
            'Me.MilitaryID.Items.Insert(0, New ListItem(dr1("Name"), dr1("MilitaryID")))
            'Me.MilitaryID.Items.Insert(1, New ListItem("請選擇", ""))
            ''畢業狀況
            'objstr = "SELECT * FROM Key_GradState where GradID = '" & sqldr("GradID") & "'"
            'objtable = DbAccess.GetDataTable(objstr)
            'dr1 = objtable.Rows(0)
            'Me.GraduateStatus.Items.Insert(0, New ListItem(dr1("Name"), dr1("GradID")))
            'Me.GraduateStatus.Items.Insert(1, New ListItem("請選擇", ""))

            '婚姻狀況
            If Not IsDBNull(sqldr("MaritalStatus")) Then
                If sqldr("MaritalStatus") = 1 Then
                    Me.MaritalStatus.Items.Insert(0, New ListItem("已婚", sqldr("MaritalStatus")))
                Else
                    Me.MaritalStatus.Items.Insert(0, New ListItem("未婚", sqldr("MaritalStatus")))
                End If
            End If
            '聯洛電話(日)、(夜)、行動電話
            phoned.Text = sqldr("Phone1")
            If Not IsDBNull(sqldr("Phone2")) Then
                phonen.Text = sqldr("Phone2")
            End If
            If Not IsDBNull(sqldr("CellPhone")) Then
                cellphone.Text = sqldr("CellPhone")
            End If
            If Trim(cellphone.Text) <> "" Then
                Common.SetListItem(rblmobil, "Y")
            Else
                Common.SetListItem(rblmobil, "N")
            End If

            '通訊地址
            objstr = "select ic.CTName,iz.ZipName from ID_ZIP iz JOIN ID_City ic ON ic.CTID = iz.CTID where iz.ZipCode = '" & sqldr("ZipCode") & "'"
            objtable = DbAccess.GetDataTable(objstr)
            dr1 = objtable.Rows(0)
            city1.Text = "(" & sqldr("ZipCode") & ")" & dr1("CTName") & dr1("ZipName")
            zipcode1.Value = sqldr("ZipCode")
            address.Text = sqldr("Address")
            If Not IsDBNull(sqldr("Email")) Then
                email.Text = sqldr("Email")
            End If
            If Request("proecess") = "add" Then
                '新增動作
                'Me.DegreeID.Items.Insert(0, New ListItem("請選擇", ""))
                'Me.MilitaryID.Items.Insert(0, New ListItem("請選擇", ""))
                'Me.GraduateStatus.Items.Insert(0, New ListItem("請選擇", ""))
                Me.midentityid.Items.Insert(0, New ListItem("請選擇", ""))
                Me.handtypeid.Items.Insert(0, New ListItem("請選擇", ""))
                Me.handlevelid.Items.Insert(0, New ListItem("請選擇", ""))
                Me.joblessid.Items.Insert(0, New ListItem("請選擇", ""))
                Me.q4.Items.Insert(0, New ListItem("請選擇", ""))
            End If
        Else
            '沒資料
            Me.degreeid.Items.Insert(0, New ListItem("請選擇", ""))
            Me.MilitaryID.Items.Insert(0, New ListItem("請選擇", ""))
            Me.graduatestatus.Items.Insert(0, New ListItem("請選擇", ""))
            Me.midentityid.Items.Insert(0, New ListItem("請選擇", ""))
            Me.handtypeid.Items.Insert(0, New ListItem("請選擇", ""))
            Me.handlevelid.Items.Insert(0, New ListItem("請選擇", ""))
            Me.joblessid.Items.Insert(0, New ListItem("請選擇", ""))
            Me.q4.Items.Insert(0, New ListItem("請選擇", ""))
        End If
    End Function

    '回報名登錄
    Private Sub Button22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button22.Click
        TIMS.Utl_Redirect1(Me, "SD_01_001.aspx?ID=" & Request("ID") & "")
    End Sub

    Private Sub DataGrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles datagrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                'If Not Me.ViewState("sort") Is Nothing Then
                '    Dim i As Integer = -1
                '    Dim MyImage As New Web.UI.WebControls.Image
                '    Select Case Me.ViewState("sort")
                '        Case "DistName", "DistName DESC"
                '            i = 1
                '        Case "OrgName", "OrgName DESC"
                '            i = 4
                '        Case "TMID", "TMID DESC"
                '            i = 5
                '        Case "ClassName", "ClassName DESC"
                '            i = 6
                '        Case "TRound", "TRound DESC"
                '            i = 8
                '    End Select
                '    If Me.ViewState("sort").ToString.IndexOf(" DESC") = -1 Then
                '        MyImage.ImageUrl = "../../images/SortUp.gif"
                '    Else
                '        MyImage.ImageUrl = "../../images/SortDown.gif"
                '    End If
                '    If i <> -1 Then
                '        e.Item.Cells(i).Controls.Add(MyImage)
                '    End If
                'End If
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                e.Item.Cells(0).Text = e.Item.ItemIndex + 1 + datagrid2.PageSize * datagrid2.CurrentPageIndex

                'Dim MyTable As HtmlTable = e.Item.FindControl("Table4")

                'Dim LName As Label = e.Item.FindControl("LName")
                'Dim LIDNO As Label = e.Item.FindControl("LIDNO")
                'Dim LBirthday As Label = e.Item.FindControl("LBirthday")
                'Dim LSex As Label = e.Item.FindControl("LSex")
                'Dim LIdent As Label = e.Item.FindControl("LIdent")
                'Dim LTel As Label = e.Item.FindControl("LTel")
                'Dim LAddress As Label = e.Item.FindControl("LAddress")

                'Select Case sm.UserInfo.LID
                '    Case 0, 1
                '        e.Item.Cells(1).Style("CURSOR") = "hand"
                '        e.Item.Cells(1).Attributes("onmouseover") = "ShowPersonData('" & MyTable.ClientID & "');"
                '        e.Item.Cells(1).Attributes("onmouseout") = "HidPersonData('" & MyTable.ClientID & "');"
                'End Select

                'LName.Text = drv("Name").ToString
                'LIDNO.Text = TIMS.ChangeIDNO(drv("IDNO").ToString)
                'If drv("Birthday").ToString <> "" Then
                '    LBirthday.Text = FormatDateTime(drv("Birthday"), 2)
                'End If
                'If drv("Sex").ToString = "M" Then
                '    LSex.Text = "男"
                'ElseIf drv("Sex").ToString = "F" Then
                '    LSex.Text = "女"
                'End If
                'LIdent.Text = drv("Ident").ToString
                'LTel.Text = drv("Tel").ToString
                'LAddress.Text = drv("Address").ToString
        End Select
    End Sub

    'Private Sub DataGrid2_SortCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridSortCommandEventArgs) Handles DataGrid2.SortCommand
    '    If Me.ViewState("sort") = e.SortExpression Then
    '        Me.ViewState("sort") = e.SortExpression & " DESC"
    '    Else
    '        Me.ViewState("sort") = e.SortExpression
    '    End If
    '    SearchHistory(Me.ViewState("IDNO"))
    'End Sub

    Sub SearchHistory(ByVal IDNO_Val As String)
        Dim SearchStr1 As String = "" 'StdAll a
        Dim SearchStr2 As String = "" 'History_StudentInfo93 a
        Dim SearchStr3 As String

        datagrid2.CurrentPageIndex = 0
        If IDNO_Val <> "" Then
            SearchStr1 = " and a.SID='" & IDNO_Val & "'"
            SearchStr2 = " and a.IDNO='" & IDNO_Val & "'"
            SearchStr3 = " and b.IDNO='" & IDNO_Val & "'"
        Else
            Exit Sub
        End If
        Call GetStudentData(SearchStr1, SearchStr2, SearchStr3)
    End Sub

    Function GetStudentData(ByVal SearchStr1 As String, ByVal SearchStr2 As String, ByVal SearchStr3 As String)
        Dim sql As String
        Dim dt As DataTable
        Dim dr As DataRow
        Dim dt1 As DataTable
        Dim dr1 As DataRow
        Dim dt2 As DataTable
        Dim dr2 As DataRow
        Dim dt3 As DataTable
        Dim dr3 As DataRow
        Dim RecordCountInt As Integer = 2000
        Dim Key_Identity As DataTable

        Key_Identity = TIMS.Get_KeyTable("Key_Identity")

        '建立DataGird用的DataTable格式------------------Start
        dt = New DataTable
        dt.Columns.Add(New DataColumn("IDNO"))
        dt.Columns.Add(New DataColumn("Name"))
        dt.Columns.Add(New DataColumn("Sex"))
        dt.Columns.Add(New DataColumn("Birthday", System.Type.GetType("System.DateTime")))
        dt.Columns.Add(New DataColumn("DistName"))                  '轄區中心
        dt.Columns.Add(New DataColumn("Years"))
        dt.Columns.Add(New DataColumn("PlanName"))
        dt.Columns.Add(New DataColumn("OrgName"))                   '訓練機構
        dt.Columns.Add(New DataColumn("TMID"))                      '訓練職類
        dt.Columns.Add(New DataColumn("ClassName"))                 '班別
        dt.Columns.Add(New DataColumn("THours"))                    '受訓時數
        dt.Columns.Add(New DataColumn("TRound"))                    '受訓期間
        dt.Columns.Add(New DataColumn("SkillName"))                 '技能檢定
        dt.Columns.Add(New DataColumn("TFlag"))                     '訓練狀態

        dt.Columns.Add(New DataColumn("Ident"))                     '身分別
        dt.Columns.Add(New DataColumn("Tel"))                       '聯絡電話
        dt.Columns.Add(New DataColumn("Address"))                   '聯絡地址
        '建立DataGird用的DataTable格式------------------End


        'sql = "SELECT TOP " & RecordCountInt & " * FROM StdAll WHERE 1=1" & SearchStr1 '★
        sql = "SELECT a.* FROM StdAll a WHERE 1=1 " & SearchStr1
        dt1 = DbAccess.GetDataTable(sql, objconn)
        For Each dr1 In dt1.Rows
            If RecordCountInt > 0 Then
                RecordCountInt -= 1
            Else
                Exit For
            End If

            dr = dt.NewRow
            dt.Rows.Add(dr)
            dr("IDNO") = TIMS.ChangeIDNO(dr1("SID"))
            dr("Name") = dr1("Name")
            dr("Sex") = dr1("Sex")
            dr("Birthday") = dr1("Birth")
            dr("DistName") = dr1("DistName")
            dr("Years") = dr1("Years")
            dr("PlanName") = dr1("PlanName")
            dr("OrgName") = dr1("TrinUnit")
            'dr("TMID") = dr1("")
            dr("ClassName") = dr1("ClassName")
            'dr("THours") = dr1("")
            If dr1("SDate").ToString <> "" And dr1("EDate").ToString <> "" Then
                dr("TRound") = FormatDateTime(dr1("SDate"), DateFormat.ShortDate) & "<BR>|<BR>" & FormatDateTime(dr1("EDate"), DateFormat.ShortDate)
            End If
            'dr("SkillName") = dr1("")
            dr("TFlag") = "結訓"

            dr("Ident") = IIf(IsNumeric(dr1("Ident")), "無法辨別", dr1("Ident").ToString)
            dr("Tel") = dr1("Tel").ToString
            dr("Address") = dr1("Addr").ToString
        Next

        ''sql = "SELECT TOP " & RecordCountInt & " a.*,b.TrainName FROM " '★
        'sql = "SELECT a.*,b.TrainName FROM "
        'sql += "(SELECT * FROM History_StudentInfo93 WHERE 1=1" & SearchStr2 & ") a "
        'sql += "LEFT JOIN Key_TrainType b ON a.TMID=b.TMID"
        ''sql += " where ROWNUM < (" & RecordCountInt & "+1) " '★

        sql = "" & vbCrLf
        sql += " select a.*,b.TrainName" & vbCrLf
        sql += " from History_StudentInfo93 a " & vbCrLf
        sql += " LEFT JOIN Key_TrainType b ON a.TMID=b.TMID" & vbCrLf
        sql += " WHERE 1=1" & vbCrLf
        sql += SearchStr2

        dt2 = DbAccess.GetDataTable(sql, objconn)
        For Each dr2 In dt2.Rows
            If RecordCountInt > 0 Then
                RecordCountInt -= 1
            Else
                Exit For
            End If
            dr = dt.NewRow
            dt.Rows.Add(dr)
            dr("IDNO") = TIMS.ChangeIDNO(dr2("IDNO"))
            dr("Name") = dr2("Name")
            dr("Sex") = dr2("Sex")
            dr("Birthday") = dr2("Birth")
            dr("DistName") = dr2("DistName")
            dr("PlanName") = dr2("PlanName")
            dr("OrgName") = dr2("TrinUnit")
            dr("TMID") = dr2("TrainName")
            dr("ClassName") = dr2("ClassName")
            dr("TRound") = FormatDateTime(dr2("SDate"), DateFormat.ShortDate) & "<BR>|<BR>" & FormatDateTime(dr2("EDate"), DateFormat.ShortDate)
            dr("TFlag") = "結訓"

            If Key_Identity.Select("IdentityID='" & dr2("Ident") & "'").Length > 0 Then
                dr("Ident") = Key_Identity.Select("IdentityID='" & dr2("Ident") & "'")(0)("Name")
            Else
                dr("Ident") = "無身分別"
            End If
            dr("Tel") = dr2("Tel").ToString
            dr("Address") = dr2("Addr").ToString
        Next

        '/*
        '資料重複請刪除()
        'select * 
        'from Stud_TechExam 
        'where socid in ( SELECT SOCID FROM Stud_TechExam  group by SOCID having count(*) > 1 ) 
        'order by socid 

        'delete Stud_TechExam 
        'where socid in ( SELECT SOCID  FROM Stud_TechExam  group by SOCID having count(*) > 1 ) 
        'and steid in (SELECT max(steid) setid  FROM Stud_TechExam  group by SOCID having count(*) > 1 )
        '*/

        sql = ""
        'sql += "SELECT TOP " & RecordCountInt & " b.IDNO,b.Name,b.Sex,b.Birthday " & vbCrLf '★
        sql += "SELECT b.IDNO,b.Name,b.Sex,b.Birthday " & vbCrLf
        sql += " ,f.Name as DistName,e.OrgName,g.TrainName as TMID " & vbCrLf
        sql += " ,c.ClassCName + '第' + c.cyclType + '期' as ClassName" & vbCrLf
        sql += " ,case when a.TrainHours is null then c.THours else a.TrainHours end THours " & vbCrLf
        sql += " ,case when a.OpenDate is null then c.STDate else a.OpenDate end STDate " & vbCrLf
        sql += " ,case when a.CloseDate is null then c.FTDate else a.CloseDate end FTDate " & vbCrLf
        sql += " ,a.TrainHours ,a.RejectTDate1, a.RejectTDate2 " & vbCrLf
        sql += " ,h.ExamName,a.StudStatus,a.MIdentityID " & vbCrLf
        sql += " ,j.PhoneD,j.ZipCode1,j.Address,k.PlanName,i.Years " & vbCrLf
        sql += " FROM Class_StudentsOfClass a " & vbCrLf
        sql += " JOIN Stud_StudentInfo b ON a.SID=b.SID " & vbCrLf
        sql += " JOIN Class_ClassInfo c ON a.OCID=c.OCID " & vbCrLf
        sql += " JOIN Auth_Relship d ON c.RID=d.RID " & vbCrLf
        sql += " JOIN Org_OrgInfo e ON d.OrgID=e.OrgID " & vbCrLf
        sql += " LEFT JOIN ID_District f ON d.DistID=f.DistID " & vbCrLf
        sql += " LEFT JOIN Key_TrainType g ON c.TMID=g.TMID " & vbCrLf
        sql += " LEFT JOIN Stud_TechExam h ON a.SOCID=h.SOCID " & vbCrLf
        sql += " JOIN ID_Plan i ON i.PlanID=c.PlanID " & vbCrLf
        sql += " JOIN KEY_plan k ON k.TPlanID=i.TPlanID " & vbCrLf
        sql += " JOIN Stud_SubData j ON j.SID=b.SID " & vbCrLf
        sql += " Where 1=1 " & vbCrLf
        sql += SearchStr3
        dt3 = DbAccess.GetDataTable(sql, objconn)
        For Each dr3 In dt3.Rows
            If RecordCountInt > 0 Then
                RecordCountInt -= 1
            Else
                Exit For
            End If
            dr = dt.NewRow
            dt.Rows.Add(dr)
            dr("IDNO") = TIMS.ChangeIDNO(dr3("IDNO"))
            dr("Name") = dr3("Name")
            dr("Sex") = dr3("Sex")
            dr("Birthday") = dr3("Birthday")
            dr("DistName") = dr3("DistName")
            dr("Years") = dr3("Years")
            dr("PlanName") = dr3("PlanName")
            dr("OrgName") = dr3("OrgName")
            dr("TMID") = dr3("TMID")
            dr("ClassName") = dr3("ClassName")
            Select Case dr3("StudStatus").ToString '訓練狀態，以 Class_StudentsOfClass 為優先資料顯示 Class_ClassInfo 為副
                Case "2" '"離訓"
                    dr("THours") = "<FONT color='Red'>" & dr3("TrainHours") & "</FONT>" '參訓時數，以 Class_StudentsOfClass 為主
                    dr("TRound") = FormatDateTime(dr3("STDate"), DateFormat.ShortDate) & "<BR>|<BR>" & FormatDateTime(dr3("RejectTDate1"), DateFormat.ShortDate)
                Case "3" '"退訓"
                    dr("THours") = "<FONT color='Red'>" & dr3("TrainHours") & "</FONT>" '參訓時數，以 Class_StudentsOfClass 為主
                    dr("TRound") = FormatDateTime(dr3("STDate"), DateFormat.ShortDate) & "<BR>|<BR>" & FormatDateTime(dr3("RejectTDate2"), DateFormat.ShortDate)
                Case Else
                    dr("THours") = dr3("THours") '參訓時數，以 Class_StudentsOfClass 為優先資料顯示 Class_ClassInfo 為副
                    dr("TRound") = FormatDateTime(dr3("STDate"), DateFormat.ShortDate) & "<BR>|<BR>" & FormatDateTime(dr3("FTDate"), DateFormat.ShortDate)
            End Select
            dr("SkillName") = dr3("ExamName")
            Select Case dr3("StudStatus").ToString
                Case "1"
                    dr("TFlag") = "在訓"
                Case "2"
                    dr("TFlag") = "離訓"
                Case "3"
                    dr("TFlag") = "退訓"
                Case "4"
                    dr("TFlag") = "續訓"
                Case "5"
                    dr("TFlag") = "結訓"
            End Select

            If Key_Identity.Select("IdentityID='" & dr3("MIdentityID") & "'").Length > 0 Then
                dr("Ident") = Key_Identity.Select("IdentityID='" & dr3("MIdentityID") & "'")(0)("Name")
            Else
                dr("Ident") = "無身分別"
            End If
            dr("Tel") = dr3("PhoneD").ToString
            dr("Address") = TIMS.Get_ZipName(dr3("ZipCode1")) & dr3("Address").ToString
        Next

        If dt.Rows.Count = 0 Then
            recordcount.Text = ""
            msg.Text = "查無資料!"
            historytable.Style.Item("display") = "none"
        Else
            msg.Text = ""
            recordcount.Text = dt.Rows.Count

            historytable.Style.Item("display") = "inline"
            If Me.ViewState("sort") Is Nothing Then
                Me.ViewState("sort") = "IDNO,Birthday,TRound"
            End If
            dt.DefaultView.Sort = Me.ViewState("sort")

            'PageControler1.PageDataTable = dt
            'PageControler1.Sort = "IDNO,Birthday,TRound"
            'PageControler1.ControlerLoad()

            datagrid2.DataSource = dt
            datagrid2.DataBind()

            dt.Dispose()
            dt = Nothing
        End If
    End Function


End Class


