Public Class SD_01_012
    Inherits AuthBasePage

    '小table
    Dim Key_Identity As DataTable = Nothing
    Dim Key_Trade As DataTable = Nothing
    Dim dtZip As DataTable = Nothing

    '取得郵遞區號資料
    'Function Get_ZipName() As DataTable
    '    Dim rst As New DataTable
    '    Dim sql As String = "SELECT CTID ,ZIPCODE ,ZIPNAME ,CTNAME ,ZNAME ,LCID ,KLNAME ,LNAME FROM VIEW_ZIPNAME ORDER BY 1,2 "
    '    Call TIMS.OpenDbConn(objconn)
    '    Dim oCmd As New SqlCommand(sql, objconn)
    '    With oCmd
    '        .Parameters.Clear()
    '        rst.Load(.ExecuteReader())
    '    End With
    '    Return rst
    'End Function

    'Function Get_TradeDt() As DataTable
    '    Dim rst As New DataTable
    '    Dim sql As String = "SELECT TradeID , '['+TradeID+']'+TradeName TradeName FROM Key_Trade ORDER BY 1"
    '    Call TIMS.OpenDbConn(objconn)
    '    Dim oCmd As New SqlCommand(sql, objconn)
    '    With oCmd
    '        .Parameters.Clear()
    '        rst.Load(.ExecuteReader())
    '    End With
    '    Return rst
    'End Function

    Function Get_IdentityDt() As DataTable
        Dim rst As New DataTable
        Dim strSql As String = ""
        '20090123 andy  edit 產投 2009年 身分別「就業保險被保險人非自願失業者」 名稱改為「非自願離職者」
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 AndAlso CInt(Me.sm.UserInfo.Years) > 2008 Then
            strSql = ""
            strSql &= " SELECT IdentityID"
            strSql &= "  ,CASE WHEN IdentityID = '02' THEN N'非自願離職者' ELSE Name END Name "
            strSql &= " FROM Key_Identity"
            'strSql &= " ORDER BY 1 "
        Else
            strSql = " SELECT IDENTITYID ,NAME FROM KEY_IDENTITY"
        End If
        'Call TIMS.OpenDbConn(objconn)
        Dim oCmd As New SqlCommand(strSql, objconn)
        With oCmd
            .Parameters.Clear()
            rst.Load(.ExecuteReader())
        End With
        Return rst
    End Function

    '檢查 內網已有資料
    Function Check_EnterType(ByVal sIDNO As String, ByVal signUpStatus As String, ByVal OCID1 As String) As Boolean
        Dim rst As Boolean = False
        If signUpStatus = "1" Then Return rst

        Dim eParms As New Hashtable
        eParms.Add("IDNO", sIDNO)
        eParms.Add("OCID1", OCID1)
        Dim sql As String = ""
        sql = " SELECT 1 FROM STUD_ENTERTEMP a" & vbCrLf
        sql &= " JOIN STUD_ENTERTYPE b ON a.SETID = b.SETID" & vbCrLf
        sql &= " WHERE a.IDNO=@IDNO AND b.OCID1 =@OCID1" & vbCrLf
        Dim dt1 As New DataTable
        'TIMS.OpenDbConn(objconn)
        Dim oCmd As New SqlCommand(sql, objconn)
        Call DbAccess.HashParmsChange(oCmd, eParms)
        dt1.Load(oCmd.ExecuteReader())
        If dt1.Rows.Count > 0 Then rst = True
        Return rst
    End Function

    Function loaddata1(ByVal eSerNum As String) As DataRow
        Dim dr As DataRow = Nothing
        Dim stdBLACK2TPLANID As String = ""
        Dim iStdBlackType As Integer = TIMS.Chk_StdBlackType(Me, objconn, stdBLACK2TPLANID)
        Dim sqlWSB As String = TIMS.Get_StdBlackWSB(Me, iStdBlackType, stdBLACK2TPLANID, 1) 'WSB

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= sqlWSB
        sql &= " SELECT a.ESETID " & vbCrLf '  /*PK*/ 
        sql &= " ,a.SETID " & vbCrLf
        sql &= " ,a.IDNO " & vbCrLf
        sql &= " ,a.NAME " & vbCrLf
        sql &= " ,a.SEX " & vbCrLf
        sql &= " ,a.BIRTHDAY " & vbCrLf
        sql &= " ,a.PASSPORTNO " & vbCrLf
        sql &= " ,a.MARITALSTATUS " & vbCrLf
        sql &= " ,a.DEGREEID " & vbCrLf
        sql &= " ,a.GRADID " & vbCrLf
        sql &= " ,a.SCHOOL " & vbCrLf
        sql &= " ,a.DEPARTMENT " & vbCrLf
        sql &= " ,a.MILITARYID " & vbCrLf
        sql &= " ,a.ZIPCODE " & vbCrLf
        sql &= " ,a.ZIPCODE6W " & vbCrLf
        sql &= " ,a.ADDRESS " & vbCrLf
        sql &= " ,a.PHONE1 " & vbCrLf
        sql &= " ,a.PHONE2 " & vbCrLf
        sql &= " ,a.CELLPHONE " & vbCrLf
        sql &= " ,a.EMAIL " & vbCrLf
        sql &= " ,a.ISAGREE " & vbCrLf
        'sql += "       ,a.MODIFYACCT " & vbCrLf
        'sql += "       ,a.MODIFYDATE " & vbCrLf
        'sql += "       ,a.LAINFLAG " & vbCrLf
        'sql += "       ,b.ESERNUM " & vbCrLf
        'sql += "       ,b.ESETID " & vbCrLf
        'sql += "       ,b.SETID " & vbCrLf
        'sql += "       ,b.ENTERDATE " & vbCrLf
        'sql += "       ,b.SERNUM " & vbCrLf
        sql &= " ,b.RELENTERDATE " & vbCrLf
        'sql += "       ,b.EXAMNO " & vbCrLf
        'sql += "       ,b.OCID1 " & vbCrLf
        'sql += "       ,b.TMID1 " & vbCrLf
        'sql += "       ,b.OCID2 " & vbCrLf
        'sql += "       ,b.TMID2 " & vbCrLf
        'sql += "       ,b.OCID3 " & vbCrLf
        'sql += "       ,b.TMID3 " & vbCrLf
        sql &= " ,b.IDENTITYID " & vbCrLf
        'sql += "       ,b.RID " & vbCrLf
        'sql += "       ,b.PLANID " & vbCrLf
        'sql += "       ,b.CCLID " & vbCrLf
        'sql += "       ,b.SIGNUPSTATUS " & vbCrLf
        'sql += "       ,b.SIGNUPMEMO " & vbCrLf
        'sql += "       ,b.ISOUT " & vbCrLf
        'sql += "       ,b.SUPPLYID " & vbCrLf
        'sql += "       ,b.BUDID " & vbCrLf
        'sql += "       ,b.MODIFYACCT " & vbCrLf
        'sql += "       ,b.MODIFYDATE " & vbCrLf
        'sql += "       ,b.ENTERPATH " & vbCrLf
        'sql += "       ,b.WORKSUPPIDENT " & vbCrLf
        'sql += "       ,b.USERNOSHOW " & vbCrLf
        'sql += "       ,b.NOTE S" & vbCrLf
        'sql += "       ,b.ISEMAILFAIL " & vbCrLf
        'sql += "       ,b.SIGNNO " & vbCrLf
        'sql += "       ,b.INVOLLEAVER " & vbCrLf
        'sql += "       ,b.CFIRE1 " & vbCrLf
        'sql += "       ,b.CFIRE1NS " & vbCrLf
        'sql += "       ,b.CFIRE1REASON " & vbCrLf
        'sql += "       ,b.CFIRE1MACCT " & vbCrLf
        'sql += "       ,b.CFIRE1MDATE " & vbCrLf
        'sql += "       ,b.CMASTER1 " & vbCrLf
        'sql += "       ,b.CMASTER1NS " & vbCrLf
        'sql += "       ,b.CMASTER1REASON " & vbCrLf
        'sql += "       ,b.CMASTER1MACCT " & vbCrLf
        'sql += "       ,b.CMASTER1MDATE " & vbCrLf
        'sql += "       ,b.CMASTER1NT " & vbCrLf
        'sql += "       ,b.CFIRE1R2 " & vbCrLf
        sql &= " ,f.examdate " & vbCrLf
        sql &= " ,c.Name DegreeName " & vbCrLf
        sql &= " ,d.Name GradName " & vbCrLf
        sql &= " ,e.Name MilitaryName " & vbCrLf
        sql &= " ,f.ClassCName ClassCName1 " & vbCrLf
        sql &= " ,f.CyclType CyclType1 " & vbCrLf
        sql &= " ,f.STDate " & vbCrLf
        sql &= " ,f.FTDate " & vbCrLf
        'sql += "       ,DATEADD(MONTH, -6, f.STDate) BFDate " & vbCrLf
        'sql += "       ,g.ClassCName AS ClassCName2 ,g.CyclType AS CyclType2 ,h.ClassCName AS ClassCName3 ,h.CyclType AS CyclType3 " & vbCrLf
        '有多筆 學員處分資料
        sql &= " ,CASE WHEN sb.IDNO IS NOT NULL THEN 'Y' ELSE 'N' END ISblack " & vbCrLf 'WSB
        sql &= " ,b.OCID1 " & vbCrLf
        sql &= " ,b.OCID2 " & vbCrLf
        sql &= " ,b.OCID3 " & vbCrLf
        sql &= " ,b.signUpStatus " & vbCrLf
        sql &= " ,b.signUpMemo " & vbCrLf
        sql &= " ,i.LevelName " & vbCrLf
        sql &= " ,se.ActNo " & vbCrLf
        sql &= " ,se.PriorWorkType1 PriorWorkType2 " & vbCrLf
        sql &= " ,se.PriorWorkOrg1 PriorWorkOrg2" & vbCrLf
        sql &= " ,se.SOfficeYM1 SOfficeYM2 " & vbCrLf
        sql &= " ,se.FOfficeYM1 FOfficeYM2 " & vbCrLf
        sql &= " FROM STUD_ENTERTEMP2 a " & vbCrLf
        'sql += " JOIN Stud_EnterType2 b ON a.eSETID = b.eSETID " & vbCrLf
        sql &= " JOIN STUD_ENTERTYPE2DELDATA b ON a.eSETID = b.eSETID " & vbCrLf
        sql &= " LEFT JOIN Key_Degree c ON a.DegreeID = c.DegreeID " & vbCrLf
        sql &= " LEFT JOIN Key_GradState d ON a.GradID = d.GradID " & vbCrLf
        sql &= " LEFT JOIN Key_Military e ON a.MilitaryID = e.MilitaryID " & vbCrLf
        sql &= " JOIN Class_ClassInfo f ON b.OCID1 = f.OCID" & vbCrLf
        sql &= " JOIN ID_PLAN ip ON ip.planid = f.planid" & vbCrLf
        sql &= " LEFT JOIN Class_ClassLevel i ON b.OCID1 = i.OCID AND b.CCLID = i.CCLID " & vbCrLf
        '有多筆 學員處分資料
        sql &= " LEFT JOIN WSB sb ON sb.IDNO = a.IDNO " & vbCrLf
        sql &= " LEFT JOIN STUD_ENTERSUBDATA2 se ON se.eSerNum = b.eSerNum " & vbCrLf
        sql &= " WHERE b.eSerNum = " & eSerNum & vbCrLf

        'If Not TIMS.sUtl_ChkTest Then
        'End If

        Select Case sm.UserInfo.LID
            Case 0 '(不限制)
            Case 1 '分署只能查該計畫/登入年度
                sql &= " AND ip.TPlanID = '" & sm.UserInfo.TPlanID & "' " & vbCrLf
                sql &= " AND ip.Years = '" & sm.UserInfo.Years & "' " & vbCrLf
            Case Else
                '(正式機) 使用者只能查自已計畫
                sql &= " AND ip.PlanID = '" & sm.UserInfo.PlanID & "' " & vbCrLf
        End Select

        dr = DbAccess.GetOneRow(sql, objconn)
        Return dr
    End Function

    'list Stud_EnterType2.eSerNum
    Sub create1(ByVal eSerNum As String)
        msg.Text = ""
        Dim dr As DataRow = loaddata1(eSerNum)
        If dr Is Nothing Then
            HidIDNO.Value = ""
            HidSTDate.Value = ""
            Common.MessageBox(Me, "查無資料!!")
            Exit Sub
        End If

        'If dr Is Nothing Then
        '    Common.RespWrite(Me, "<script>alert('查無資料');history.go(-1);</script>")
        'End If

        ''WorkSuppIdent
        'If IsDBNull(dr("WorkSuppIdent")) = False Then Common.SetListItem(WorkSuppIdent, dr("WorkSuppIdent"))
        Dim vsExamDate As String = ""
        HidIDNO.Value = ""
        Dim vsOCID1val As String = ""
        If Convert.ToString(dr("examdate")) <> "" Then vsExamDate = Common.FormatDate(dr("examdate"))
        HidIDNO.Value = Convert.ToString(dr("IDNO"))
        'If IDNOValue.Value <> "" Then IDNOValue.Value = Trim(IDNOValue.Value)
        'If IDNOValue.Value <> "" Then IDNOValue.Value = UCase(IDNOValue.Value)
        If HidIDNO.Value <> "" Then HidIDNO.Value = TIMS.ChangeIDNO(HidIDNO.Value)
        vsOCID1val = Convert.ToString(dr("OCID1"))

        HidSTDate.Value = ""
        HidFTDate.Value = ""
        If IsDBNull(dr("STDate")) = False Then HidSTDate.Value = TIMS.Cdate3(Convert.ToString(dr("STDate")))
        If IsDBNull(dr("FTDate")) = False Then HidFTDate.Value = TIMS.Cdate3(Convert.ToString(dr("FTDate")))
        'If IsDBNull(dr("BFDate")) = False Then HidFTDate.Value = FormatDateTime(Convert.ToString(dr("BFDate")), DateFormat.ShortDate) Else BFDateValue.Value = ""

        'IDNOValue.Value = TIMS.ChangeIDNO(dr("IDNO").ToString)
        LabIDNO.Text = TIMS.ChangeIDNO(dr("IDNO").ToString)
        'If IDNO.Text <> "" Then IDNO.Text = Trim(IDNO.Text)
        'If IDNO.Text <> "" Then IDNO.Text = UCase(IDNO.Text)
        If LabIDNO.Text <> "" Then LabIDNO.Text = TIMS.ChangeIDNO(LabIDNO.Text)
        'LabIDNO.Text = TIMS.strMask(LabIDNO.Text, 1)

        '--------------------start受訓前學員任職資料--------
        Select Case dr("PriorWorkType2").ToString
            Case "1"
                PriorWorkType1.Text = "曾工作過"
            Case "2"
                PriorWorkType1.Text = "未曾工作過"
            Case "3"
                PriorWorkType1.Text = "先前從事為非勞保性質工作"
            Case Else
                PriorWorkType1.Text = "無資料"
        End Select

        If Convert.ToString(dr("PriorWorkOrg2")) <> "" Then
            PriorWorkOrg1.Text = Convert.ToString(dr("PriorWorkOrg2"))
        Else
            PriorWorkOrg1.Text = "無資料"
        End If

        If Convert.ToString(dr("ActNo")) <> "" Then
            ActNo.Text = Convert.ToString(dr("ActNo"))
        Else
            ActNo.Text = "無資料"
        End If

        If Convert.ToString(dr("SOfficeYM2")) <> "" AndAlso Convert.ToString(dr("FOfficeYM2")) <> "" Then
            OfficeDate.Text = Format(dr("SOfficeYM2"), "yyyy/MM/dd") & "~" & Format(dr("FOfficeYM2"), "yyyy/MM/dd")
        Else
            OfficeDate.Text = "無資料"
        End If
        '--------------------end 受訓前學員任職資料-----

        LabNAME.Text = dr("Name").ToString
        Birthday.Text = TIMS.Cdate3(dr("Birthday"))
        'Birthday.Text = TIMS.strMask(Birthday.Text, 2)
        'If IsDate(dr("Birthday")) Then Birthday.Text = FormatDateTime(dr("Birthday"), 2)

        If dr("PassPortNO") = 1 Then
            PassPortNO.Text = "本國"
        Else
            PassPortNO.Text = "外國"
        End If

        If dr("Sex").ToString = "M" Then
            Sex.Text = "男"
        Else
            Sex.Text = "女"
        End If
        Select Case dr("MaritalStatus").ToString
            Case "1"
                MaritalStatus.Text = "已婚"
            Case "2"
                MaritalStatus.Text = "未婚"
            Case ""
                MaritalStatus.Text = "無資料"
        End Select
        DegreeID.Text = If(dr("DegreeName").ToString = "", "無資料", dr("DegreeName").ToString)
        GradID.Text = If(dr("GradName").ToString = "", "無資料", dr("GradName").ToString)
        School.Text = If(dr("School").ToString = "", "無資料", dr("School").ToString)
        Department.Text = If(dr("Department").ToString = "", "無資料", dr("Department").ToString)
        MilitaryID.Text = If(dr("MilitaryName").ToString = "", "無資料", dr("MilitaryName").ToString)

        If dtZip Is Nothing Then dtZip = TIMS.Get_VZipName(objconn)

        Dim sAddress As String = Convert.ToString(dr("Address"))
        Dim sZipCODE As String = If(Convert.ToString(dr("ZipCODE6W")) <> "", Convert.ToString(dr("ZipCODE6W")), Convert.ToString(dr("ZipCode")))
        Address.Text = TIMS.getZipName6(sZipCODE, sAddress, "", dtZip)

        Phone1.Text = If(dr("Phone1").ToString = "", "無資料", dr("Phone1").ToString)
        Phone2.Text = If(dr("Phone2").ToString = "", "無資料", dr("Phone2").ToString)
        Email.Text = If(dr("Email").ToString = "", "無資料", dr("Email").ToString)
        CellPhone.Text = If(dr("CellPhone").ToString = "", "無資料", dr("CellPhone").ToString)

        'TRIdentityID.Style.Item("display") = "inline"
        TRIdentityID.Style.Item("display") = "none" '新的e網報名，完全不用顯示參訓身分別(暫時) ---2008-11-24 by AMU
        TRHandTypeID.Style.Item("display") = "none" '障礙類別行

        '不管是不是用e網報名的資料，只要有 障礙類別 就顯示喔 ---2008-11-23 by AMU
        Dim sqlE As String = " SELECT * FROM E_MEMBER WHERE MEM_IDNO = '" & dr("IDNO").ToString & "' "
        Dim drE As DataRow
        drE = DbAccess.GetOneRow(sqlE, objconn)
        If Not drE Is Nothing Then
            If drE("HandTypeID").ToString <> "" Then
                TRIdentityID.Style.Item("display") = "none"
                TRHandTypeID.Style.Item("display") = ""
                Me.labHandTypeID.Text = TIMS.Get_HandTypeName(drE("HandTypeID").ToString)
                Me.labHandLevelID.Text = TIMS.Get_HandLevelName(drE("HandLevelID").ToString)
            End If
            If drE("HandTypeID2").ToString <> "" Then
                TRIdentityID.Style.Item("display") = "none"
                TRHandTypeID.Style.Item("display") = ""
                Me.labHandTypeID.Text = TIMS.Get_HandTypeName2(drE("HandTypeID2").ToString)
                Me.labHandLevelID.Text = TIMS.Get_HandLevelName2(drE("HandLevelID2").ToString)
            End If
        End If

        'Select Case dr("EnterPath").ToString.ToUpper
        '    Case "E" '使用e網報名的資料喔
        '        Dim sqlE As String = "SELECT * FROM E_Member WHERE mem_idno='" & dr("IDNO").ToString & "'"
        '        drE = DbAccess.GetOneRow(sqlE)
        '        If Not drE Is Nothing Then
        '            If drE("HandTypeID").ToString <> "" Then
        '                TRIdentityID.Style.Item("display") = "none"
        '                TRHandTypeID.Style.Item("display") = "inline"
        '                Me.labHandTypeID.Text = TIMS.Get_HandTypeName(drE("HandTypeID").ToString)
        '                Me.labHandLevelID.Text = TIMS.Get_HandLevelName(drE("HandLevelID").ToString)
        '            End If
        '        End If
        'End Select

        If Key_Identity Is Nothing Then Key_Identity = Get_IdentityDt()
        IdentityID.Text = TIMS.Get_IdentityName(Convert.ToString(dr("IdentityID")), Key_Identity, ",")

        'For i As Integer = 0 To Split(dr("IdentityID"), ",").Length - 1
        '    sql = "SELECT Name FROM Key_Identity WHERE IdentityID='" & Split(dr("IdentityID"), ",")(i) & "'"
        '    '20090123 andy  edit 產投 2009年 身分別「就業保險被保險人非自願失業者」 名稱改為「非自願離職者」
        '    '-----------------------------------
        '    If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '        If IdentityID.Text = "" Then
        '            If Split(dr("IdentityID"), ",")(i) = "02" Then
        '                IdentityID.Text = "非自願離職者"
        '            Else
        '                IdentityID.Text = DbAccess.ExecuteScalar(sql, objconn)
        '            End If
        '        Else
        '            If Split(dr("IdentityID"), ",")(i) = "02" Then
        '                IdentityID.Text += "," & "非自願離職者"
        '            Else
        '                IdentityID.Text += "," & DbAccess.ExecuteScalar(sql, objconn)
        '            End If
        '        End If
        '    Else
        '        If IdentityID.Text = "" Then
        '            IdentityID.Text = DbAccess.ExecuteScalar(sql, objconn)
        '        Else
        '            IdentityID.Text += "," & DbAccess.ExecuteScalar(sql, objconn)
        '        End If
        '    End If
        'Next
        If IsDate(dr("RelEnterDate")) Then RelEnterDate.Text = FormatDateTime(dr("RelEnterDate"), DateFormat.GeneralDate)

        'LabOCID1.Text
        LabOCID1.Text = dr("ClassCName1").ToString
        If IsNumeric(dr("CyclType1")) Then
            If Int(dr("CyclType1")) <> 0 Then LabOCID1.Text += "第" & Int(dr("CyclType1")) & "期"
        End If
        LabClassCname.Text = LabOCID1.Text

        If dr("LevelName").ToString <> "" Then
            If Int(dr("LevelName")) <> 0 Then LabOCID1.Text += "(第" & Int(dr("LevelName")) & "階段)"
        End If

        '檢查 內網已有資料
        If Check_EnterType(Convert.ToString(dr("IDNO")), Convert.ToString(dr("signUpStatus")), Convert.ToString(dr("OCID1"))) Then
            LabOCID1.Text += "(內網已有資料)"
            LabOCID1.ForeColor = Color.Red
        End If

        'If dr("ClassCName2").ToString = "" Then
        '    OCID2.Text = "無資料"
        'Else
        '    OCID2.Text = dr("ClassCName2").ToString
        '    If IsNumeric(dr("CyclType2")) Then
        '        If Int(dr("CyclType2")) <> 0 Then OCID3.Text += "第" & Int(dr("CyclType2")) & "期"
        '    End If
        '    If dr("signUpStatus") <> 1 Then
        '        For Each dr1 In dt1.Rows
        '            If dr1("OCID1").ToString = dr("OCID2").ToString Or dr1("OCID2").ToString = dr("OCID2").ToString Or dr1("OCID3").ToString = dr("OCID2").ToString Then
        '                'OCID2.Text += "(已重複報名)"
        '                OCID2.Text += "(內網已有資料)"
        '                OCID2.ForeColor = Color.Red
        '                Exit For
        '            End If
        '        Next
        '    End If
        'End If
        'If dr("ClassCName3").ToString = "" Then
        '    OCID3.Text = "無資料"
        'Else
        '    OCID3.Text = dr("ClassCName3").ToString
        '    If IsNumeric(dr("CyclType3")) Then
        '        If Int(dr("CyclType3")) <> 0 Then OCID3.Text += "第" & Int(dr("CyclType3")) & "期"
        '    End If
        '    If dr("signUpStatus") <> 1 Then
        '        For Each dr1 In dt1.Rows
        '            If dr1("OCID1").ToString = dr("OCID3").ToString Or dr1("OCID2").ToString = dr("OCID3").ToString Or dr1("OCID3").ToString = dr("OCID3").ToString Then
        '                'OCID3.Text += "(已重複報名)"
        '                OCID3.Text += "(內網已有資料)"
        '                OCID3.ForeColor = Color.Red
        '                Exit For
        '            End If
        '        Next
        '    End If
        'End If

        'signUpMemo.Text = dr("signUpMemo").ToString
        'If dr("signUpStatus") <> 0 Then
        '    Button1.Visible = False
        '    Button2.Visible = False
        'Else
        '    If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 AndAlso dr("ISblack") = "Y" Then
        '        signUpMemo.Text = "此學員己遭處分,系統帶審核失敗"
        '        signUpMemo.ReadOnly = True
        '        Button1.Enabled = False
        '        Button2.Enabled = True
        '    End If
        'End If

        '2006/06/21 by Vicient----產學訓---start
        'Me.LabOCID1.Text = Me.OCID1.Text
        Me.LabEnterDate.Text = Me.RelEnterDate.Text
        Me.LabName2.Text = Me.LabNAME.Text
        Me.ViewState("Birthday") = Me.Birthday.Text
        LabBirthDay.Text = Me.Birthday.Text
        'LabBirthDay.Text = TIMS.strMask(LabBirthDay.Text, 2)
        'If Session(TIMS.gcst_rblWorkMode) = "1" Then
        '    Me.Birthday.Text = TIMS.strMask(Me.Birthday.Text, 2)
        '    Me.LabBirthDay.Text = TIMS.strMask(Me.LabBirthDay.Text, 2)
        'End If

        Me.Label5.Text = Me.PassPortNO.Text
        'If Me.IDNO.Text <> "" Then Me.IDNO.Text = TIMS.ChangeIDNO(Me.IDNO.Text)
        Me.Label6.Text = Me.LabIDNO.Text 'TIMS.ChangeIDNO(Me.IDNO.Text)
        'Me.ViewState("IDNO") = Me.IDNO.Text
        'If Session(TIMS.gcst_rblWorkMode) = "1" Then
        '    Me.LabIDNO.Text = TIMS.strMask(Me.LabIDNO.Text, 1)
        '    Me.Label6.Text = TIMS.strMask(Me.Label6.Text, 1)
        'End If

        Me.Label7.Text = Me.Sex.Text
        'Me.Label8.Text = Me.MaritalStatus.Text
        Me.Label9.Text = Me.DegreeID.Text
        'Me.Label10.Text = Me.GradID.Text
        'Me.Label11.Text = Me.School.Text
        'Me.Label12.Text = Me.Department.Text
        'Me.Label13.Text = Me.MilitaryID.Text
        Me.Label14.Text = Address.Text
        Me.Label16.Text = Me.Phone1.Text
        Me.Label17.Text = Me.Phone2.Text
        Me.Label18.Text = Me.Email.Text
        Me.Label19.Text = Me.CellPhone.Text
        'Me.Label21.Text = Me.IdentityID.Text
    End Sub

    'list Stud_EnterTrain2.eSerNum
    Sub create2(ByVal eSerNum As String)
        Dim sql As String = ""
        Dim Table As DataTable
        eSerNum = TIMS.ClearSQM(eSerNum)

        Dim eParms As New Hashtable
        eParms.Add("eSerNum", eSerNum)
        sql = " SELECT * FROM STUD_ENTERTRAIN2 WHERE eSerNum =@eSerNum"
        Table = DbAccess.GetDataTable(sql, objconn, eParms)

        If Table.Rows.Count = 0 Then
            Me.Label15.Text = "無資料"
            Me.Label20.Text = "無資料"
            'Me.Label22.Text = "無資料"
            'Me.Label23.Text = "無資料"
            'Me.Label24.Text = "無資料"
            'Me.Label25.Text = "無資料"
            'Me.Label26.Text = "無資料"
            'Me.Label27.Text = "無資料"
            'Me.Label28.Text = "無資料"
            'Me.Label29.Text = "無資料"
            Me.Label30.Text = "無資料"
            'Me.Label31.Text = "無資料"
            'Me.Label32.Text = "無資料"
            Me.Label33.Text = "無資料"
            'Me.Label39.Text = "無資料"
            Me.Label40.Text = "無資料"
            Me.Label41.Text = "無資料"
            'Me.Label42.Text = "無資料"
            'Me.Label43.Text = "無資料"
            'Me.Label44.Text = "無資料"
            'Me.Label45.Text = "無資料"
            Me.Label46.Text = "無資料"
            'Me.Label47.Text = "無資料"
            'Me.Label48.Text = "無資料"
            'Me.Label49.Text = "無資料"
            'Me.Label50.Text = "無資料"
            Me.Label51.Text = "無資料"
            Me.Label52.Text = "無資料"
            Me.Label53.Text = "無資料"
            Me.Label54.Text = "無資料"
            Me.Label55.Text = "無資料"
            Me.Label56.Text = "無資料"
            Me.Label57.Text = "無資料"
            Me.Label58.Text = "無資料"
            Me.Label59.Text = "無資料"
            Me.Label60.Text = "無資料"
            Me.Label61.Text = "無資料"
            Me.Label62.Text = "無資料"
            Me.Label63.Text = "無資料"
            Me.Label64.Text = "無資料"
            Me.Label65.Text = "無資料"
            'Me.ActComidno.Text = "無資料"
        Else
            Dim dr As DataRow
            dr = Table.Rows(0)
            Me.Label15.Text = "無資料"

            If dtZip Is Nothing Then dtZip = TIMS.Get_VZipName(objconn)
            'Dim sZ1 As String = Convert.ToString(dr("ZipCode2"))
            'Dim sZ2 As String = Convert.ToString(dr("ZipCode2_6W"))
            Dim sAddress As String = Convert.ToString(dr("HouseholdAddress"))
            Dim sZipCode2 As String = If(Convert.ToString(dr("ZipCode2_6W")) <> "", Convert.ToString(dr("ZipCode2_6W")), Convert.ToString(dr("ZipCode2")))
            Label15.Text = TIMS.getZipName6(sZipCode2, sAddress, "", dtZip)

            If Key_Identity Is Nothing Then Key_Identity = Get_IdentityDt()
            Me.Label20.Text = "無資料"
            If Not IsDBNull(dr("MIdentityID")) Then
                Me.Label20.Text = TIMS.GetMyValue(Key_Identity, "IdentityID", "Name", Convert.ToString(dr("MIdentityID")))
            End If

            'If Not IsDBNull(dr("HandTypeID")) Then
            '    str = "select * from Key_HandicatType where HandTypeID = '" & dr("HandTypeID") & "'"
            '    table1 = DbAccess.GetDataTable(str)
            '    If table1.Rows.Count <> 0 Then
            '        dr1 = table1.Rows(0)
            '        Me.Label22.Text = dr1("Name")
            '    End If
            'Else
            '    Me.Label22.Text = "無資料"
            'End If

            'If Not IsDBNull(dr("HandLevelID")) Then
            '    str = "select * from Key_HandicatLevel where HandLevelID = '" & dr("HandLevelID") & "'"
            '    table1 = DbAccess.GetDataTable(str)
            '    If table1.Rows.Count <> 0 Then
            '        dr1 = table1.Rows(0)
            '        Me.Label23.Text = dr1("Name")
            '    End If
            'Else
            '    Me.Label23.Text = "無資料"
            'End If

            'If Not IsDBNull(dr("PriorWorkOrg1")) Then
            '    Me.Label24.Text = dr("PriorWorkOrg1")
            'Else
            '    Me.Label24.Text = "無資料"
            'End If

            'If Not IsDBNull(dr("Title1")) Then
            '    Me.Label26.Text = dr("Title1")
            'Else
            '    Me.Label26.Text = "無資料"
            'End If

            'If Not IsDBNull(dr("PriorWorkOrg2")) Then
            '    Me.Label25.Text = dr("PriorWorkOrg2")
            'Else
            '    Me.Label25.Text = "無資料"
            'End If

            'If Not IsDBNull(dr("Title2")) Then
            '    Me.Label27.Text = dr("Title2")
            'Else
            '    Me.Label27.Text = "無資料"
            'End If

            'If Not IsDBNull(dr("SOfficeYM1")) Then
            '    Me.Label28.Text = dr("SOfficeYM1")
            '    Me.Label28.Text = Me.Label28.Text & " ~ "
            'End If
            'If Not IsDBNull(dr("FOfficeYM1")) Then
            '    If Not IsDBNull(dr("SOfficeYM1")) Then
            '        Me.Label28.Text = Me.Label28.Text & dr("FOfficeYM1")
            '    Else
            '        Me.Label28.Text = " ~ " & dr("FOfficeYM1")
            '    End If
            'End If
            'If Me.Label28.Text = "" Then Me.Label28.Text = "無資料"

            'If Not IsDBNull(dr("SOfficeYM2")) Then
            '    Me.Label29.Text = dr("SOfficeYM2")
            '    Me.Label29.Text = Me.Label29.Text & " ~ "
            'End If
            'If Not IsDBNull(dr("FOfficeYM2")) Then
            '    If Not IsDBNull(dr("SOfficeYM2")) Then
            '        Me.Label29.Text = Me.Label29.Text & dr("FOfficeYM2")
            '    Else
            '        Me.Label29.Text = " ~ " & dr("FOfficeYM2")
            '    End If
            'End If
            'If Me.Label29.Text = "" Then Me.Label29.Text = "無資料"
            Me.Label30.Text = "無資料"
            If Not IsDBNull(dr("PriorWorkPay")) Then Me.Label30.Text = dr("PriorWorkPay")

            'If Not IsDBNull(dr("RealJobless")) Then
            '    Me.Label31.Text = dr("RealJobless")
            'Else
            '    Me.Label31.Text = "無資料"
            'End If

            'If Not IsDBNull(dr("Traffic")) Then
            '    If dr("Traffic") = 1 Then
            '        Me.Label32.Text = "住宿"
            '    ElseIf dr("Traffic") = 2 Then
            '        Me.Label32.Text = "通勤"
            '    Else
            '        Me.Label32.Text = "無資料"
            '    End If
            'Else
            '    Me.Label32.Text = "無資料"
            'End If
            '**by Milor 20080512--加入項目2訓練單位代轉現金----start
            If Not IsDBNull(dr("AcctMode")) Then
                If dr("AcctMode") = 1 Then
                    Me.Label33.Text = "銀行"
                    Table22.Style("display") = "none"
                    Table23.Style("display") = ""
                    If Not IsDBNull(dr("BankName")) Then
                        Me.Label37.Text = dr("BankName")
                    End If
                    If Not IsDBNull(dr("AcctHeadNo")) Then
                        Me.Label38.Text = dr("AcctHeadNo")
                    End If
                    If Not IsDBNull(dr("ExBankName")) Then
                        Me.Label61.Text = dr("ExBankName")
                    End If
                    If Not IsDBNull(dr("AcctExNo")) Then
                        Me.Label62.Text = dr("AcctExNo")
                    End If
                    If Not IsDBNull(dr("AcctNo")) Then
                        Me.Label34.Text = dr("AcctNo")
                    End If
                ElseIf dr("AcctMode") = 0 Then
                    Me.Label33.Text = "郵局"
                    Table22.Style("display") = ""
                    Table23.Style("display") = "none"
                    If Not IsDBNull(dr("PostNo")) Then
                        Me.Label35.Text = dr("PostNo")
                    End If
                    If Not IsDBNull(dr("AcctNo")) Then
                        Me.Label36.Text = dr("AcctNo")
                    End If
                ElseIf dr("AcctMode") = 2 Then
                    Me.Label33.Text = "訓練單位代轉現金"
                    Table22.Style("display") = "none"
                    Table23.Style("display") = "none"
                End If
            Else
                Me.Label33.Text = "無資料"
                Table22.Style("display") = "none"
                Table23.Style("display") = "none"
            End If
            '**by Milor 20080512----end
            'If Not IsDBNull(dr("FirDate")) Then
            '    Me.Label39.Text = dr("FirDate")
            'Else
            '    Me.Label39.Text = "無資料"
            'End If

            If Not IsDBNull(dr("Uname")) Then
                Me.Label40.Text = dr("Uname")
            Else
                Me.Label40.Text = "無資料"
            End If

            If Not IsDBNull(dr("Intaxno")) Then
                Me.Label41.Text = dr("Intaxno")
            Else
                Me.Label41.Text = "無資料"
            End If

            'If Not IsDBNull(dr("Tel")) Then
            '    Me.Label42.Text = dr("Tel")
            'Else
            '    Me.Label42.Text = "無資料"
            'End If

            'If Not IsDBNull(dr("Fax")) Then
            '    Me.Label43.Text = dr("Fax")
            'Else
            '    Me.Label43.Text = "無資料"
            'End If

            'If dr("Zip").ToString <> "" And Trim(dr("Zip").ToString) <> "-1" Then
            '    Me.Label44.Text = "[" & dr("Zip") & "]" & TIMS.Get_ZipName(dr("Zip"))
            '    If dr("Addr").ToString <> "" Then
            '        Me.Label44.Text = Me.Label44.Text & dr("Addr")
            '    End If
            'Else
            '    Me.Label44.Text = "無資料"
            'End If

            If Not IsDBNull(dr("ServDept")) Then
                Me.Label45.Text = dr("ServDept")
            Else
                Me.Label45.Text = "無資料"
            End If

            If Not IsDBNull(dr("JobTitle")) Then
                Me.Label46.Text = dr("JobTitle")
            Else
                Me.Label46.Text = "無資料"
            End If

            'If Not IsDBNull(dr("SDate")) Then
            '    Me.Label47.Text = dr("SDate")
            'Else
            '    Me.Label47.Text = "無資料"
            'End If

            'If Not IsDBNull(dr("SJDate")) Then
            '    Me.Label48.Text = dr("SJDate")
            'Else
            '    Me.Label48.Text = "無資料"
            'End If

            'If Not IsDBNull(dr("SPDate")) Then
            '    Me.Label49.Text = dr("SPDate")
            'Else
            '    Me.Label49.Text = "無資料"
            'End If

            If Not IsDBNull(dr("Q1")) Then
                If dr("Q1") Then
                    Me.Label50.Text = "是"
                Else
                    Me.Label50.Text = "否"
                End If
            Else
                Me.Label50.Text = "無資料"
            End If

            Dim z As Integer
            z = 1
            If Not IsDBNull(dr("Q2_1")) Then
                Me.Label51.Text = z & ". 為補充與原專長相關之技能 "
                z = z + 1
            End If
            If Not IsDBNull(dr("Q2_2")) Then
                Me.Label51.Text = Me.Label51.Text & z & ". 轉換其他行職業所需技能 "
                z = z + 1
            End If
            If Not IsDBNull(dr("Q2_3")) Then
                Me.Label51.Text = Me.Label51.Text & z & " .拓展工作領域及視野 "
                z = z + 1
            End If
            If Not IsDBNull(dr("Q2_4")) Then
                Me.Label51.Text = Me.Label51.Text & z & " .其他"
            End If
            If Me.Label51.Text = "" Then
                Me.Label51.Text = "無資料"
            End If

            If Not IsDBNull(dr("Q3")) Then
                Select Case Convert.ToString(dr("Q3"))
                    Case "1"
                        Me.Label52.Text = "轉換工作"
                    Case "2"
                        Me.Label52.Text = "留任"
                    Case "3"
                        Me.Label52.Text = "其他"
                        If Not IsDBNull(dr("Q3_Other")) Then
                            Me.Label52.Text &= "(" & dr("Q3_Other") & ")"
                        End If
                End Select
            Else
                Me.Label52.Text = "無資料"
            End If

            If Not IsDBNull(dr("Q5")) Then
                If Convert.ToString(dr("Q5")) = "1" Then
                    Me.Label53.Text = "是" '是
                Else
                    Me.Label53.Text = "否" '否
                End If
            Else
                Me.Label53.Text = "無資料"
            End If

            If Not IsDBNull(dr("Q61")) Then
                Me.Label54.Text = dr("Q61")
            Else
                Me.Label54.Text = "無資料"
            End If

            If Not IsDBNull(dr("Q62")) Then
                Me.Label55.Text = dr("Q62")
            Else
                Me.Label55.Text = "無資料"
            End If

            If Not IsDBNull(dr("Q63")) Then
                Me.Label56.Text = dr("Q63")
            Else
                Me.Label56.Text = "無資料"
            End If

            If Not IsDBNull(dr("Q64")) Then
                Me.Label57.Text = dr("Q64")
            Else
                Me.Label57.Text = "無資料"
            End If

            If Key_Trade Is Nothing Then Key_Trade = TIMS.Get_TradeDt(objconn)
            Me.Label58.Text = "無資料"
            If Not IsDBNull(dr("Q4")) Then
                Me.Label58.Text = TIMS.GetMyValue(Key_Trade, "TradeID", "TradeName", Convert.ToString(dr("Q4")))
            End If

            If Not IsDBNull(dr("ActNo")) Then
                Me.Label60.Text = dr("ActNo")
            Else
                Me.Label60.Text = "無資料"
            End If
            If Not IsDBNull(dr("Actname")) Then
                Me.Label59.Text = dr("Actname")
            Else
                Me.Label59.Text = "無資料"
            End If
            If Not IsDBNull(dr("ActType")) Then
                Select Case Convert.ToString(dr("ActType"))
                    Case "1"
                        Me.Label63.Text = "勞"
                    Case "2"
                        Me.Label63.Text = "農"
                End Select
            Else
                Me.Label63.Text = "無資料"
            End If

            'If Not IsDBNull(dr("ActComidno")) Then
            '    Me.ActComidno.Text = dr("ActComidno")
            'Else
            '    Me.ActComidno.Text = "無資料"
            'End If

            If Not IsDBNull(dr("ActTel")) Then
                Me.Label64.Text = dr("ActTel")
            Else
                Me.Label64.Text = "無資料"
            End If

            Me.Label65.Text = "無資料"
            If dtZip Is Nothing Then dtZip = TIMS.Get_VZipName(objconn)
            sAddress = Convert.ToString(dr("ActAddress"))
            Dim sZipCode3 As String = If(Convert.ToString(dr("ZipCode3_6W")) <> "", Convert.ToString(dr("ZipCode3_6W")), Convert.ToString(dr("ZipCode3")))
            Me.Label65.Text = TIMS.getZipName6(sZipCode3, sAddress, "", dtZip)
        End If
    End Sub

    Dim objconn As SqlConnection

    Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End
        PageControler1.PageDataGrid = DataGrid1

        If Not IsPostBack Then
            msg.Text = ""
            tbSch.Visible = True
            tbList.Visible = False
            tbView.Visible = False

            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID

            If sm.UserInfo.LID <= 1 Then
                Button2.Disabled = False
                center.Enabled = True
            Else
                Button2.Disabled = True
                center.Enabled = False
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
        TIMS.ShowHistoryClass(Me, Historytable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1")
        If Historytable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

        Button1.Attributes("onclick") = "return CheckSearch();"
        Button4.Attributes("onclick") = "ClearData();"
    End Sub

    '查詢 SQL
    Sub Search1()

        TIMS.SUtl_TxtPageSize(Me, Me.TxtPageSize, Me.DataGrid1)
        'select esernum 
        ',modifydate
        ',count(1) cnt 
        'from stud_enterType2DelData 
        'group by esernum 
        ',modifydate
        'having count(1) >1
        txtQIDNO.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(txtQIDNO.Text))
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        txtQName.Text = TIMS.ClearSQM(txtQName.Text)

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " WITH WP2 AS (" & vbCrLf
        sql &= " SELECT isnull(a.SETID,b.SETID) SETID,a.IDNO,a.NAME,a.ESETID,a.BIRTHDAY" & vbCrLf
        sql &= " FROM STUD_ENTERTEMP2 a WITH(NOLOCK)" & vbCrLf
        sql &= " LEFT JOIN STUD_ENTERTEMP b WITH(NOLOCK) on b.IDNO=a.IDNO AND b.BIRTHDAY =a.BIRTHDAY" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        '身分證字號
        If txtQIDNO.Text <> "" Then sql &= " AND a.IDNO LIKE '" & txtQIDNO.Text & "%'" & vbCrLf
        '學員姓名
        If txtQName.Text <> "" Then sql &= " AND a.Name LIKE '%" & txtQName.Text & "%'" & vbCrLf
        'sql &= " AND A.IDNO='U201013029'" & vbCrLf
        sql &= " )" & vbCrLf

        'sql &= " ,WT1 AS (" & vbCrLf
        'sql &= " select b.eSernum,b.modifyDate,b.EnterDate,b.RelEnterDate,b.OCID1,b.modifyAcct,b.SERNUM" & vbCrLf
        'sql &= " ,(select x.SIGNNO from STUD_ENTERTYPE2 x WITH(NOLOCK) where x.ocid1=b.ocid1 and x.esetid=b.esetid) SIGNNO" & vbCrLf
        'sql &= " ,a.IDNO" & vbCrLf
        'sql &= " FROM WP2 a" & vbCrLf
        'sql &= " JOIN STUD_ENTERTYPEDELDATA b WITH(NOLOCK) ON b.esetid=a.esetid and b.modifyAcct=a.idno" & vbCrLf
        'sql &= " )" & vbCrLf

        sql &= " ,WT2 AS (" & vbCrLf
        sql &= " select b.eSernum,b.modifyDate,b.EnterDate,b.RelEnterDate,b.OCID1,b.modifyAcct,b.SERNUM" & vbCrLf
        sql &= " ,b.SIGNNO" & vbCrLf
        sql &= " ,a.IDNO" & vbCrLf
        sql &= " FROM WP2 a" & vbCrLf
        sql &= " JOIN STUD_ENTERTYPE2DELDATA b WITH(NOLOCK) ON b.esetid=a.esetid and b.modifyAcct=a.idno" & vbCrLf
        '班級
        If OCIDValue1.Value <> "" Then sql &= " AND b.OCID1='" & OCIDValue1.Value & "'" & vbCrLf
        sql &= " )" & vbCrLf

        sql &= " SELECT a.SETID ,a.idno ,a.name" & vbCrLf
        sql &= " ,b.SIGNNO ,b.eSernum ,b.SERNUM ,b.modifyAcct" & vbCrLf
        sql &= " ,convert(varchar, b.modifyDate, 120) modifyDate" & vbCrLf
        sql &= " ,CONVERT(varchar, b.EnterDate, 111) EnterDate" & vbCrLf
        sql &= " ,convert(varchar, b.RelEnterDate, 120) RelEnterDate" & vbCrLf
        sql &= " ,oo.orgname ,ip.distname" & vbCrLf
        sql &= " ,cc.ocid" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.ClassCName,cc.CyclType) ClassCName" & vbCrLf
        sql &= " ,cc.CyclType" & vbCrLf
        sql &= " ,cc.RID" & vbCrLf
        sql &= " ,cc.Relship" & vbCrLf
        sql &= " ,CONVERT(varchar, cc.STDate, 111)+'<br>|<br>'+CONVERT(varchar, cc.FTDate, 111) TRound" & vbCrLf
        sql &= " ,CONVERT(varchar, cc.STDate, 111) STDate" & vbCrLf
        sql &= " ,CONVERT(varchar, cc.FTDate, 111) FTDate" & vbCrLf
        sql &= " " & vbCrLf
        sql &= " FROM WP2 a" & vbCrLf
        'sql &= " JOIN (SELECT * FROM WT1 UNION SELECT * FROM WT2) b on b.IDNO=a.IDNO" & vbCrLf
        sql &= " JOIN WT2 b on b.IDNO=a.IDNO" & vbCrLf
        sql &= " JOIN CLASS_CLASSINFO cc WITH(NOLOCK) ON cc.ocid =b.ocid1" & vbCrLf
        sql &= " LEFT JOIN ORG_ORGINFO oo WITH(NOLOCK) ON oo.comidno=cc.comidno" & vbCrLf
        sql &= " LEFT JOIN VIEW_PLAN ip ON ip.planid =cc.planid" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND cc.NotOpen='N'" & vbCrLf
        sql &= " AND cc.IsSuccess='Y'" & vbCrLf
        '(正式的限制-未結訓-未開訓+14)
        'sql += " AND cc.FTDate >= dbo.TRUNC_DATETIME(GETDATE()) " & vbCrLf '未結訓 
        'sql += " AND DATEADD(DAY, 14, cc.STDate) >= dbo.TRUNC_DATETIME(GETDATE()) " & vbCrLf '未開訓+14
        sql &= " AND cc.STDate >= GETDATE()-400 " & vbCrLf '(可查詢一年內刪除的班)

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)

        Dim iLen As Integer = 0 'Len(RIDValue.Value)

        Select Case sm.UserInfo.LID
            Case 0 '(不限制) '發展署無限
            Case 1 '分署只能查該計畫/登入年度
                If RIDValue.Value <> "" Then
                    iLen = Len(RIDValue.Value) '搜尋單位
                    If iLen = 1 Then '各分署
                        sql &= " AND cc.RID LIKE '" & RIDValue.Value & "%'" & vbCrLf '& SearchStr
                    Else '各機構
                        sql &= " AND cc.RID ='" & RIDValue.Value & "'" & vbCrLf '& SearchStr
                    End If
                Else
                    sql &= " AND ip.TPlanID = '" & sm.UserInfo.TPlanID & "' " & vbCrLf
                    sql &= " AND ip.Years = '" & sm.UserInfo.Years & "' " & vbCrLf
                End If

            Case Else
                '(正式機) 使用者只能查自已計畫
                sql &= " AND ip.PlanID = '" & sm.UserInfo.PlanID & "' " & vbCrLf
                If RIDValue.Value <> "" Then
                    sql &= " AND cc.RID ='" & RIDValue.Value & "'" & vbCrLf '& SearchStr
                Else
                    sql &= " AND cc.RID ='" & sm.UserInfo.RID & "'" & vbCrLf '& SearchStr
                End If
        End Select

        '班級
        If OCIDValue1.Value <> "" Then sql &= " AND cc.OCID='" & OCIDValue1.Value & "'" & vbCrLf
        '身分證字號
        If txtQIDNO.Text <> "" Then sql &= " AND a.IDNO LIKE '" & txtQIDNO.Text & "%'" & vbCrLf
        '學員姓名
        If txtQName.Text <> "" Then sql &= " AND a.Name like '%" & txtQName.Text & "%'" & vbCrLf

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)
        tbList.Visible = False
        msg.Text = "查無資料"
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return

        'If dt.Rows.Count > 0 Then End If
        TIMS.Tooltip(Button1, "(查詢開訓日1年內的刪除資料)", True)
        tbList.Visible = True
        msg.Text = ""

        'PageControler1.SqlPrimaryKeyDataCreate(sql, "SETID")
        PageControler1.PageDataTable = dt
        'PageControler1.PrimaryKey = "SETID"
        PageControler1.ControlerLoad()
    End Sub

    '查詢
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim flag_can_sch As Boolean = False
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        txtQIDNO.Text = TIMS.ClearSQM(txtQIDNO.Text)
        txtQName.Text = TIMS.ClearSQM(txtQName.Text)
        If OCIDValue1.Value <> "" Then flag_can_sch = True
        If txtQIDNO.Text <> "" Then flag_can_sch = True
        If txtQName.Text <> "" Then flag_can_sch = True
        If Not flag_can_sch Then
            'alert('至少要輸入一項條件');
            Common.MessageBox(Me, "至少要輸入一項條件!")
            Return
        End If

        Call Search1()
    End Sub

    Private Sub DataGrid1_ItemCommand(source As Object, e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Select Case e.CommandName
            Case "vie"
                tbSch.Visible = False
                tbList.Visible = False
                tbView.Visible = True
                Dim sCmdArg As String = e.CommandArgument
                Dim eSernum As String = TIMS.GetMyValue(sCmdArg, "eSernum")
                Call create1(eSernum)
                Call create2(eSernum)
                'Call create3()
                'Call Createhistory3()
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = "head_navy"
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim labIDNO As Label = e.Item.FindControl("labIDNO")
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e) '序號

                If Convert.ToString(drv("SIGNNO")) <> "" Then
                    e.Item.Cells(0).Text = "<font color=Red>" & CStr(drv("SIGNNO")) & "</font>"
                    TIMS.Tooltip(e.Item.Cells(0), "此為報名序號!")
                End If
                labIDNO.Text = TIMS.strMask(Convert.ToString(drv("IDNO")), 1)

                Dim BtnView As LinkButton = e.Item.FindControl("BtnView")
                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "SETID", Convert.ToString(drv("SETID")))
                TIMS.SetMyValue(sCmdArg, "eSernum", Convert.ToString(drv("eSernum")))
                TIMS.SetMyValue(sCmdArg, "modifyDate", Convert.ToString(drv("modifyDate")))
                TIMS.SetMyValue(sCmdArg, "IDNO", Convert.ToString(drv("IDNO")))
                BtnView.CommandArgument = sCmdArg 'drv("IDNO").ToString

                BtnView.Visible = True
                BtnView.Enabled = True '委訓單位無法檢視
                Select Case Convert.ToString(sm.UserInfo.LID)
                    Case "2"
                        BtnView.Visible = False
                        BtnView.Enabled = False
                        TIMS.Tooltip(BtnView, "委訓單位暫不提供檢視功能!!")
                End Select
        End Select
    End Sub

    '該帳號有賦于班級時(只有一個時)帶出該班級
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        '判斷機構是否只有一個班級
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn)
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        tbSch.Visible = True
        tbList.Visible = False
        tbView.Visible = False
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        '如果只有一個班級
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
        'tbSch.Visible = True
        'tbList.Visible = False
        'tbView.Visible = False
    End Sub

    '該帳號有賦于班級時(只有一個時)帶出該班級
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub

    '回上頁
    Protected Sub BtnBack1_Click(sender As Object, e As EventArgs) Handles BtnBack1.Click
        tbSch.Visible = True
        tbList.Visible = True
        tbView.Visible = False
    End Sub
End Class