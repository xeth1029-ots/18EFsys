Partial Class SD_01_010_add
    Inherits AuthBasePage

#Region "參數/變數"

    'SD_01_001_1_add.aspx
    '修改此程式之相關程式為
    'editstd_2.aspx Online_3.aspx SD_01_010_add.aspx

    Const cst_msgSuper1 As String = "該使用者，可強行報名!!!"

    Dim intTmp2wVal As Integer = 0 '暫存值。
    Dim ff As String = "" '過濾字
    'Dim dtDegree As DataTable=Nothing
    'Dim dtIdentity As DataTable=Nothing
    Dim dtZipCode As DataTable = Nothing
    Dim dtGradState As DataTable = Nothing
    Dim dtSERVDEPT As DataTable = Nothing
    Dim dtJOBTITLE As DataTable = Nothing

#End Region

#Region "Function_1"
    'Dim rblEmail As ListControl

    Function GET_E_MEMBER(ByVal IDNO As String, ByVal birthDay As String) As DataTable
        Dim dt As New DataTable
        Call TIMS.OpenDbConn(objconn)
        If IDNO <> "" AndAlso birthDay <> "" AndAlso IsDate(birthDay) Then
            Dim sql As String = " SELECT * FROM dbo.E_MEMBER WHERE mem_idno=@IDNO AND mem_birth=@BIRTHDAY"
            Dim sCmd As New SqlCommand(sql, objconn)
            With sCmd
                .Parameters.Clear()
                .Parameters.Add("IDNO", SqlDbType.VarChar).Value = IDNO
                .Parameters.Add("BIRTHDAY", SqlDbType.DateTime).Value = TIMS.Cdate2(birthDay)
                dt.Load(.ExecuteReader())
            End With
            Return dt
        End If

        Dim sql_E As String = " SELECT * FROM dbo.E_MEMBER WHERE 1<>1 "
        dt = DbAccess.GetDataTable(sql_E, objconn)
        Return dt
    End Function

    Sub SHOW_E_MEMBER(ByVal dt As DataTable)
        If TIMS.dtNODATA(dt) Then Return

        Dim sqldr As DataRow = dt.Rows(0)
        '婚姻狀況 '1.已;2.未 3.暫不提供(預設) 
        Select Case Convert.ToString(sqldr("mem_marry"))
            Case "1", "2"
                Common.SetListItem(MaritalStatus, sqldr("mem_marry").ToString)
            Case Else
                Common.SetListItem(MaritalStatus, "3")
        End Select
        Name.Text = sqldr("mem_name").ToString '會員姓名

        '身分別
        If sqldr("mem_foreign") = 0 Then
            PassPortNO.Items(0).Selected = True '本國
        Else
            PassPortNO.Items(1).Selected = True '外籍
        End If

        '性別
        If sqldr("mem_sex") = "M" Then
            Sex.Items(0).Selected = True
        Else
            Sex.Items(1).Selected = True
        End If

        '學歷
        If Not IsDBNull(sqldr("mem_edu")) Then Common.SetListItem(DegreeID, sqldr("mem_edu"))

        '畢業狀況 GraduateStatus
        If Convert.ToString(sqldr("mem_GradUATE")) <> "" Then
            Common.SetListItem(GraduateStatus, sqldr("mem_GradUATE"))
        Else
            Common.SetListItem(GraduateStatus, "01")
        End If

        '學校名稱 School
        If Convert.ToString(sqldr("mem_School")) <> "" Then
            School.Text = Convert.ToString(sqldr("mem_School"))
        ElseIf School.Text = "" Then
            School.Text = TIMS.cst_未填寫 '"不詳"
        End If

        '科系名稱 Department
        If Convert.ToString(sqldr("mem_Depart")) <> "" Then
            Department.Text = Convert.ToString(sqldr("mem_Depart"))
        ElseIf Department.Text = "" Then
            Department.Text = TIMS.cst_未填寫 '"不詳"
        End If

        '聯洛電話(日)、(夜)、行動電話
        If Not IsDBNull(sqldr("mem_tel")) Then PhoneD.Text = sqldr("mem_tel")
        If Not IsDBNull(sqldr("mem_teln")) Then PhoneN.Text = sqldr("mem_teln")
        If Not IsDBNull(sqldr("mem_mobile")) Then CellPhone.Text = sqldr("mem_mobile")
        CellPhone.Text = TIMS.ClearSQM(CellPhone.Text)
        Dim V_TMP As String = "N"
        If CellPhone.Text <> "" Then V_TMP = "Y"
        Common.SetListItem(rblMobil, V_TMP)

        '通訊地址 ZipCode
        If Not IsDBNull(sqldr("mem_zip")) Then ZipCode1.Value = String.Format("{0:000}", CInt(sqldr("mem_zip")))
        If Convert.ToString(sqldr("mem_ZIP6W")) <> "" Then hidZipCode1_6W.Value = Convert.ToString(sqldr("mem_ZIP6W"))
        If Convert.ToString(sqldr("mem_ZIP6W")) <> "" Then ZipCode1_B3.Value = TIMS.GetZIPCODEB3(hidZipCode1_6W.Value)
        City1.Text = TIMS.getZipName2(ZipCode1.Value, hidZipCode1_6W.Value, dtZipCode)

        If Not IsDBNull(sqldr("mem_addr")) Then Address.Text = sqldr("mem_addr")
        'If Email.Text <> "" Then Email.Text=Trim(Email.Text)
        If Not IsDBNull(sqldr("mem_email")) Then Email.Text = TIMS.ClearSQM(sqldr("mem_email"))

        'Dim V_TMP As String="N"
        V_TMP = If(Email.Text <> "" AndAlso Email.Text <> "無", "Y", "N")
        Common.SetListItem(rblEmail, V_TMP)

        If Not IsDBNull(sqldr("ePaper")) Then
            If sqldr("ePaper") = "1" Then IseMail.Items(0).Selected = True Else IseMail.Items(0).Selected = False
        End If

    End Sub

    '使用STUD_ENTERTEMP3 報名資料維護。
    Sub SHOW_STUD_ENTERTEMP3()
        IDNO.Text = TIMS.ClearSQM(IDNO.Text)
        Birthday.Text = TIMS.ClearSQM(Birthday.Text)
        If IDNO.Text = "" OrElse Birthday.Text = "" Then Return

        'parms.Clear()
        Dim parms As New Hashtable From {{"IDNO", IDNO.Text}, {"Birthday", CDate(Birthday.Text)}}
        Dim sql As String = ""
        '身分別PassPortNO 1:本國 2:外籍
        sql &= " SELECT a.ESETID3 ,a.IDNO ,a.NAME ,a.BIRTHDAY ,a.SEX,a.PASSPORTNO" & vbCrLf
        sql &= " ,a.MARITALSTATUS ,a.DEGREEID ,a.GRADID ,a.SCHOOL ,a.DEPARTMENT ,a.MILITARYID" & vbCrLf
        sql &= " ,dbo.FN_GZIP3(a.ZIPCODE1) ZIPCODE1,a.ZIPCODE1_6W ,a.ADDRESS" & vbCrLf
        sql &= " ,dbo.FN_GZIP3(a.ZIPCODE2) ZIPCODE2,a.ZIPCODE2_6W ,a.HOUSEHOLDADDRESS" & vbCrLf
        sql &= " ,a.PHONE1 ,a.PHONE2 ,a.CELLPHONE ,a.EMAIL ,a.MIDENTITYID ,a.PRIORWORKPAY" & vbCrLf
        'str &= " ,a.HANDTYPEID ,a.HANDLEVELID" & vbCrLf
        sql &= " ,a.ACCTMODE ,a.POSTNO ,a.ACCTHEADNO ,a.BANKNAME ,a.ACCTEXNO ,a.EXBANKNAME ,a.ACCTNO" & vbCrLf
        sql &= " ,a.FIRDATE ,a.UNAME ,a.INTAXNO ,a.SERVDEPT ,a.ACTNAME ,a.ACTTYPE ,a.ACTNO ,a.ACTTEL" & vbCrLf
        sql &= " ,dbo.FN_GZIP3(a.ZIPCODE3) ZIPCODE3,ZIPCODE3_6W ,a.ACTADDRESS" & vbCrLf
        sql &= " ,a.JOBTITLE ,a.SERVDEPTID ,a.JOBTITLEID" & vbCrLf
        sql &= " ,a.Q1 ,a.Q2_1 ,a.Q2_2 ,a.Q2_3 ,a.Q2_4 ,a.Q3 ,a.Q3_OTHER ,a.Q4 ,a.Q5 ,a.Q61 ,a.Q62 ,a.Q63 ,a.Q64" & vbCrLf
        sql &= " ,a.ISEMAIL ,a.ISAGREE ,a.MODIFYACCT ,a.MODIFYDATE" & vbCrLf
        sql &= " FROM dbo.STUD_ENTERTEMP3 a WITH(NOLOCK)" & vbCrLf
        '基本資料代入功能，加上出生年月日，已確保資訊安全
        sql &= " WHERE a.IDNO=@IDNO AND a.Birthday=@Birthday" & vbCrLf

        Dim flag_error As Boolean = True '預設為錯誤 ! 查詢正確時為false 
        Dim dt As DataTable = Nothing
        Try
            dt = DbAccess.GetDataTable(sql, objconn, parms)
            flag_error = False
        Catch ex As Exception
            TIMS.LOG.Error(ex.Message, ex)
            Dim cst_fun_page_name As String = "##SD_01_010_add.aspx, "
            Dim slogMsg1 As String = ""
            slogMsg1 &= cst_fun_page_name & "sql: " & sql & vbCrLf
            slogMsg1 &= cst_fun_page_name & "parms: " & TIMS.GetMyValue3(parms) & vbCrLf
            'Call TIMS.SendMailTest(slogMsg1)
            Dim strErrmsg As String = ""
            strErrmsg &= "ex.Message:" & vbCrLf & ex.Message & vbCrLf
            strErrmsg &= "ex.ToString:" & vbCrLf & ex.ToString & vbCrLf
            strErrmsg &= "slogMsg1:" & vbCrLf & slogMsg1 & vbCrLf
            'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.SendMailTest(strErrmsg)
        End Try

        If TIMS.dtNODATA(dt) Then Return

        '取得最新資料。
        Dim ff As String = "IDNO='" & IDNO.Text & "'"
        Dim ss As String = "ModifyDate DESC"

        If dt.Select(ff, ss).Length = 0 Then Return

        If dt.Select(ff, ss).Length > 0 Then
            Dim sqldr As DataRow = dt.Select(ff, ss)(0)
            hid_eSETID3.Value = Convert.ToString(sqldr("ESETID3"))

            '性別
            If Not IsDBNull(sqldr("Sex")) Then
                If Convert.ToString(sqldr("Sex")) = "M" Then
                    Sex.Items(0).Selected = True
                ElseIf Convert.ToString(sqldr("Sex")) = "F" Then
                    Sex.Items(1).Selected = True
                End If
            End If

            '身分別PassPortNO 1:本國 2:外籍
            If Not IsDBNull(sqldr("PassPortNO")) Then
                If sqldr("PassPortNO") = 1 Then
                    PassPortNO.Items(0).Selected = True
                Else
                    PassPortNO.Items(1).Selected = True
                End If
            End If

            '婚姻狀況 '1.已;2.未 3.暫不提供(預設) 
            Select Case Convert.ToString(sqldr("MaritalStatus"))
                Case "1", "2"
                    Common.SetListItem(MaritalStatus, sqldr("MaritalStatus").ToString)
                Case Else
                    Common.SetListItem(MaritalStatus, "3")
            End Select

            '學歷
            If Not IsDBNull(sqldr("DegreeID")) Then Common.SetListItem(DegreeID, sqldr("DegreeID"))

            '畢業狀況 GraduateStatus
            If Convert.ToString(sqldr("GradID")) <> "" Then
                Common.SetListItem(GraduateStatus, sqldr("GradID"))
            Else
                Common.SetListItem(GraduateStatus, "01")
            End If

            '學校名稱 School
            If Convert.ToString(sqldr("School")) <> "" Then
                School.Text = Convert.ToString(sqldr("School"))
            ElseIf School.Text = "" Then
                School.Text = TIMS.cst_未填寫 '"不詳"
            End If

            '科系名稱 Department
            If Convert.ToString(sqldr("Department")) <> "" Then
                Department.Text = Convert.ToString(sqldr("Department"))
            ElseIf Department.Text = "" Then
                Department.Text = TIMS.cst_未填寫 '"不詳"
            End If

            '通訊地址 ZipCode
            ZipCode1.Value = TIMS.TrimZipCode(Convert.ToString(sqldr("ZipCode1")), dtZipCode)
            hidZipCode1_6W.Value = Convert.ToString(sqldr("ZipCode1_6W"))
            ZipCode1_B3.Value = TIMS.GetZIPCODEB3(hidZipCode1_6W.Value)
            City1.Text = TIMS.getZipName2(ZipCode1.Value, hidZipCode1_6W.Value, dtZipCode)
            Address.Text = Convert.ToString(sqldr("Address"))

            '戶籍地址
            ZipCode2.Value = TIMS.TrimZipCode(Convert.ToString(sqldr("ZipCode2")), dtZipCode)
            hidZipCode2_6W.Value = Convert.ToString(sqldr("ZipCode2_6W"))
            ZipCode2_B3.Value = TIMS.GetZIPCODEB3(hidZipCode2_6W.Value)
            City2.Text = TIMS.getZipName2(ZipCode2.Value, hidZipCode2_6W.Value, dtZipCode)
            HouseholdAddress.Text = Convert.ToString(sqldr("HouseholdAddress"))

            '聯洛電話(日)、(夜)、行動電話
            If Not IsDBNull(sqldr("Phone1")) Then PhoneD.Text = sqldr("Phone1")
            If Not IsDBNull(sqldr("Phone2")) Then PhoneN.Text = sqldr("Phone2")
            Dim v_rblMobil As String = "N"
            If Not IsDBNull(sqldr("CellPhone")) Then
                CellPhone.Text = TIMS.ClearSQM(sqldr("CellPhone"))
                If Trim(CellPhone.Text) <> "" Then v_rblMobil = "Y"
            End If
            Common.SetListItem(rblMobil, v_rblMobil)

            If Not IsDBNull(sqldr("Email")) Then
                Email.Text = TIMS.ClearSQM(sqldr("Email"))
                'If Email.Text <> "" Then Email.Text=Trim(Email.Text)
            End If

            Email.Text = TIMS.ClearSQM(Email.Text)
            Dim V_TMP As String = "N"
            V_TMP = "N"
            If Email.Text <> "" AndAlso Email.Text <> "無" Then V_TMP = "Y"
            Common.SetListItem(rblEmail, V_TMP)

            'If Email.Text <> "" AndAlso Trim(Email.Text) <> "無" Then
            '    Common.SetListItem(rblEmail, "Y")
            'Else
            '    Common.SetListItem(rblEmail, "N")
            'End If
            'Convert.ToString(sqldr("MidentityID"))
            'If Not IsDBNull(sqldr("MidentityID")) Then Common.SetListItem(MIdentityID, sqldr("MidentityID"))
            If Convert.ToString(sqldr("MidentityID")) <> "" Then Common.SetListItem(MIdentityID, sqldr("MidentityID"))
            'PriorWorkPay.Text=Convert.ToString(sqldr("PRIORWORKPAY")) '受訓前薪資

            If AcctMode.Items.Count > 0 Then AcctMode.Items(0).Selected = False
            If AcctMode.Items.Count > 1 Then AcctMode.Items(1).Selected = False
            If AcctMode.Items.Count > 2 Then AcctMode.Items(2).Selected = False

            Porttr.Style("display") = "none"
            Banktr1.Style("display") = "none"
            Banktr2.Style("display") = "none"
            Banktr3.Style("display") = "none"

            '郵局帳號
            PostNo_1.Text = ""
            'PostNo_2.Text=""
            AcctNo1_1.Text = ""
            'AcctNo1_2.Text=""

            '銀行帳號
            BankName.Text = ""
            AcctheadNo.Text = ""
            ExBankName.Text = ""
            AcctExNo.Text = ""
            AcctNo2.Text = ""

            If Convert.ToString(sqldr("AcctMode")) <> "" Then
                Select Case Convert.ToString(sqldr("AcctMode"))
                    Case "0" '郵局帳號
                        If AcctMode.Items.Count > 0 Then AcctMode.Items(0).Selected = True
                        Porttr.Style("display") = ""
                        If Not IsDBNull(sqldr("PostNo")) Then
                            Dim sPostNo As String = Convert.ToString(sqldr("PostNo"))
                            If InStr(sPostNo, "-") = 0 Then
                                PostNo_1.Text = sPostNo
                            Else
                                sPostNo = Replace(sPostNo, "-", "")
                                PostNo_1.Text = sPostNo
                            End If
                        End If
                        If Not IsDBNull(sqldr("AcctNo")) Then
                            Dim sAcctNo As String = Convert.ToString(sqldr("AcctNo"))
                            If InStr(sAcctNo, "-") = 0 Then
                                AcctNo1_1.Text = sAcctNo
                            Else
                                sAcctNo = Replace(sAcctNo, "-", "")
                                AcctNo1_1.Text = sAcctNo
                            End If
                        End If

                    Case "1" '銀行帳號
                        If AcctMode.Items.Count > 1 Then AcctMode.Items(1).Selected = True
                        Banktr1.Style("display") = ""
                        Banktr2.Style("display") = ""
                        Banktr3.Style("display") = ""
                        If Not IsDBNull(sqldr("AcctHeadNo")) Then AcctheadNo.Text = sqldr("AcctHeadNo").ToString
                        If Not IsDBNull(sqldr("BankName")) Then BankName.Text = sqldr("BankName").ToString
                        If Not IsDBNull(sqldr("AcctExNo")) Then AcctExNo.Text = sqldr("AcctExNo").ToString
                        If Not IsDBNull(sqldr("ExBankName")) Then ExBankName.Text = sqldr("ExBankName").ToString
                        If Not IsDBNull(sqldr("AcctNo")) Then AcctNo2.Text = sqldr("AcctNo").ToString

                    Case "2" '訓練單位代轉現金
                        If AcctMode.Items.Count > 2 Then AcctMode.Items(2).Selected = True

                End Select
            End If
            Select Case Convert.ToString(sqldr("IsAgree"))
                Case "Y", "N"
                    Common.SetListItem(IsAgree, Convert.ToString(sqldr("IsAgree")))
                Case Else
                    Common.SetListItem(IsAgree, "Y") '預設為同意。
            End Select

            'If Not IsDBNull(sqldr("IsAgree")) Then
            '    If sqldr("IsAgree")="Y" Then  IsAgree.Items(0).Selected=True Else  IsAgree.Items(0).Selected=False
            'End If

            If Not IsDBNull(sqldr("IseMail")) Then
                If sqldr("IseMail") = "Y" Then IseMail.Items(0).Selected = True Else IseMail.Items(0).Selected = False
            End If

            '服務單位
            If Not IsDBNull(sqldr("Uname")) Then Uname.Text = sqldr("Uname").ToString
            If Not IsDBNull(sqldr("Intaxno")) Then Intaxno.Text = sqldr("Intaxno").ToString

            '服務部門 ServDept 30 CHAR
            If Not IsDBNull(sqldr("ServDept")) Then ServDept.Text = sqldr("ServDept").ToString
            If Convert.ToString(sqldr("SERVDEPTID")) <> "" Then Common.SetListItem(ddlSERVDEPTID, sqldr("SERVDEPTID"))

            If Not IsDBNull(sqldr("Actname")) Then ActName.Text = sqldr("Actname").ToString
            If Convert.ToString(sqldr("ActType")) <> "" Then Common.SetListItem(ActType, Convert.ToString(sqldr("ActType")))
            If Not IsDBNull(sqldr("ActNo")) Then ActNo.Text = sqldr("ActNo").ToString
            If Not IsDBNull(sqldr("JobTitle")) Then JobTitle.Text = sqldr("JobTitle").ToString
            If Convert.ToString(sqldr("JOBTITLEID")) <> "" Then Common.SetListItem(ddlJOBTITLEID, sqldr("JOBTITLEID"))

            '*by Milor 20080904--投保單位電話地址----start
            If Not IsDBNull(sqldr("ActTel")) Then ActTel.Text = sqldr("ActTel").ToString

            '投保單位電話地址
            ZipCode3.Value = TIMS.TrimZipCode(Convert.ToString(sqldr("ZipCode3")), dtZipCode)
            hidZipCode3_6W.Value = Convert.ToString(sqldr("ZipCode3_6W"))
            ZipCode3_B3.Value = TIMS.GetZIPCODEB3(hidZipCode3_6W.Value)
            City3.Text = TIMS.getZipName2(ZipCode3.Value, ZipCode3_B3.Value, dtZipCode)
            If Not IsDBNull(sqldr("ActAddress")) Then ActAddress.Text = sqldr("ActAddress").ToString
            '**by Milor 20080904--投保單位電話地址----end

            '參訓背景資料,'是否由公司推薦參訓 Q1,'參訓動機 * Q2,'訓後動向 Q3 Q3_Other,'服務單位行業別 * Q4,'服務單位是否 屬於中小企業 Q5,
            '個人工作年資 Q61,'在這家公司 的年資 Q62,'在這職位的年資 Q63,'最近升遷離 本職幾年 Q64,

            '參訓背景資料 '是否由公司推薦參訓 Q1
            If Convert.ToString(sqldr("Q1")) <> "" Then Common.SetListItem(Q1, Convert.ToString(sqldr("Q1")))

            '參訓動機 * Q2
            If Convert.ToString(sqldr("Q2_1")) = "1" Then Q2.Items(0).Selected = True
            If Convert.ToString(sqldr("Q2_2")) = "1" Then Q2.Items(1).Selected = True
            If Convert.ToString(sqldr("Q2_3")) = "1" Then Q2.Items(2).Selected = True
            If Convert.ToString(sqldr("Q2_4")) = "1" Then Q2.Items(3).Selected = True

            '訓後動向 Q3 Q3_Other
            If Convert.ToString(sqldr("Q3")) <> "" Then Common.SetListItem(Q3, Convert.ToString(sqldr("Q3")))
            Q3_Other.Text = Convert.ToString(sqldr("Q3_Other"))

            '服務單位行業別 * Q4
            If Convert.ToString(sqldr("Q4")) <> "" Then Common.SetListItem(Q4, Convert.ToString(sqldr("Q4")))

            '服務單位是否 屬於中小企業 Q5
            If $"{sqldr("Q5")}" = "1" Then
                Q5.Items(0).Selected = True '是
            Else
                Q5.Items(1).Selected = True '否
            End If

            '個人工作年資 Q61
            If Convert.ToString(sqldr("Q61")) <> "" Then Q61.Text = TIMS.VAL1(sqldr("Q61"))

            '在這家公司 的年資 Q62
            If Convert.ToString(sqldr("Q62")) <> "" Then Q62.Text = TIMS.VAL1(sqldr("Q62"))

            '在這職位的年資 Q63
            If Convert.ToString(sqldr("Q63")) <> "" Then Q63.Text = TIMS.VAL1(sqldr("Q63"))

            '最近升遷離 本職幾年 Q64
            If Convert.ToString(sqldr("Q64")) <> "" Then Q64.Text = TIMS.VAL1(sqldr("Q64"))
        End If
    End Sub

    ''' <summary>無值回傳0</summary>
    ''' <param name="oConn"></param>
    ''' <param name="IDNOvalue"></param>
    ''' <returns></returns>
    Public Shared Function Get_STUD_ENTERTEMP_SETID(ByRef oConn As SqlConnection, ByRef IDNOvalue As String) As Integer
        Dim iSETID As Integer = 0 'STUD_ENTERTEMP
        Dim objstr As String = $" SELECT SETID FROM dbo.STUD_ENTERTEMP WHERE IDNO='{IDNOvalue}'"
        Dim dr1 As DataRow = DbAccess.GetOneRow(objstr, oConn)
        If dr1 Is Nothing Then Return iSETID
        iSETID = TIMS.CINT1(dr1("SETID"))
        Return iSETID
    End Function

    'Function SAVE_STUD_ENTERTEMP2_lock(ByVal IDNOvalue As String, ByVal OCID1_Value As String,
    '                                   ByRef Errmsg As String, ByVal Redirect1 As String, ByRef JaveScriptAlertMsg As String) As Boolean
    '    Dim flag As Boolean=False
    '    'Dim objLock As New Object
    '    SyncLock TIMS.objLock_SD01010ADD
    '        Dim flag_save As Boolean=SAVE_STUD_ENTERTEMP2(IDNO.Text, OCIDValue1.Value, Errmsg, Redirect1, JaveScriptAlertMsg)
    '        flag=flag_save
    '    End SyncLock
    '    Return flag
    'End Function

    ''' <summary> 存報名資料 </summary>
    ''' <param name="IDNOvalue"></param>
    ''' <param name="OCID1_Value"></param>
    ''' <param name="Errmsg"></param>
    ''' <param name="Redirect1"></param>
    ''' <param name="JaveScriptAlertMsg"></param>
    ''' <returns></returns>
    Function SAVE_STUD_ENTERTEMP2(ByVal IDNOvalue As String, ByVal OCID1_Value As String,
                                  ByRef Errmsg As String, ByVal Redirect1 As String, ByRef JaveScriptAlertMsg As String) As Boolean
        Errmsg = ""
        Const Cst_SuccessMsg1 As String = "收件成功!!!"

        Dim OKFlag As Boolean = False '正常為True，異常為False
        'If IDNOvalue="" OrElse OCID1_Value="" Then Return False

        Dim iSETID As Integer = 0 'STUD_ENTERTEMP
        Dim iSerNum As Integer = 1 'STUD_ENTERTYPE2
        Dim ieSETID As Integer = 0 'STUD_ENTERTEMP2 
        Dim ieSerNum As Integer = 0 'STUD_ENTERTYPE2
        'Dim dr As DataRow=Nothing 'Dim dt As DataTable=Nothing 'Dim da As SqlDataAdapter=Nothing
        Dim iSignNo As Integer = -1

        Dim objstr As String = ""
        Dim vMsg As String = "" '存取 ex.tostring
        '此月決議以身分證號為搜尋條件 by AMU 20090417 
        IDNOvalue = TIMS.ChangeIDNO(TIMS.ClearSQM(IDNOvalue)) 'UCase()
        OCID1_Value = TIMS.ClearSQM(OCID1_Value)
        If (IDNOvalue = "" OrElse IDNOvalue.Length <> 10) Then
            Errmsg = String.Concat("報名失敗，該學員報名資料異常!!", vbCrLf, "(若持續出現此問題，請聯絡系統管理者)!!", vbCrLf, vMsg)
            Return False
        End If

        Dim drCC1 As DataRow = TIMS.GetOCIDDate(OCID1_Value, objconn)
        If drCC1 Is Nothing Then
            Errmsg = String.Concat("報名失敗，該學員報名資料異常!", vbCrLf, "(若持續出現此問題，請聯絡系統管理者)!", vbCrLf, vMsg)
            Return False
        End If

        iSETID = Get_STUD_ENTERTEMP_SETID(objconn, IDNOvalue)

        ' MIdentityID.SelectedValue 'IdentityID.SelectedValue
        Dim v_MIdentityID As String = TIMS.GetListValue(MIdentityID)
        'Threading.Thread.Sleep(1) '假設處理某段程序需花費1毫秒 (避免機器不同步)

        aNow = TIMS.GetSysDateNow(objconn)
        Dim s_aToday As String = Common.FormatDate(aNow) 'Today()

        Dim dr2_EMAIL As String = ""
        Dim dr2_ESETID As String = ""
        Dim SSQL_2 As String = $"SELECT * FROM dbo.STUD_ENTERTEMP2 WHERE IDNO='{IDNOvalue}'" 'ORDER BY MODIFYDATE DESC
        Dim dt2 As DataTable = DbAccess.GetDataTable(SSQL_2, objconn)
        Dim dr2 As DataRow = Nothing
        If TIMS.dtHaveDATA(dt2) Then dr2 = dt2.Rows(0)
        If dr2 IsNot Nothing Then
            dr2_EMAIL = Convert.ToString(dr2("Email"))
            dr2_ESETID = dr2("ESETID")
        End If


        '婚姻狀況 '1.已;2.未 3.暫不提供(預設) 
        Dim v_MaritalStatus As String = TIMS.GetListValue(MaritalStatus)
        Dim v_DegreeID As String = TIMS.GetListValue(DegreeID)
        '畢業狀況 '01:畢業' GraduateStatus.SelectedValue
        Dim v_GraduateStatus As String = TIMS.GetListValue(GraduateStatus)
        If v_GraduateStatus <> "" Then
            ff = String.Format("GradID='{0}'", v_GraduateStatus)
            If dtGradState.Select(ff).Length = 0 Then v_GraduateStatus = "01"
        Else
            v_GraduateStatus = "01" '畢業' GraduateStatus.SelectedValue
        End If
        hidZipCode1_6W.Value = TIMS.GetZIPCODE6W(ZipCode1.Value, ZipCode1_B3.Value)
        Address.Text = TIMS.ClearSQM(Address.Text)
        PhoneD.Text = TIMS.ChangeIDNO(PhoneD.Text)
        PhoneN.Text = TIMS.ChangeIDNO(PhoneN.Text)
        CellPhone.Text = TIMS.ChangeIDNO(CellPhone.Text)
        Email.Text = TIMS.ClearSQM(Email.Text)
        Email.Text = If(Email.Text <> "", Email.Text, dr2_EMAIL)
        dr2_EMAIL = Email.Text
        '勞動部勞動力發展署 暨所屬機關，為本人提供職業訓練及就業服務時使用本人資料
        Dim v_IsAgree As String = TIMS.GetListValue(IsAgree) '預設為Y  IsAgree.SelectedValue

        Using TransConn As SqlConnection = DbAccess.GetConnection()
            Dim Trans As SqlTransaction = DbAccess.BeginTrans(TransConn)
            Try
                '線上報名資料寫入Stud_EnterTemp2中---start '此月決議以身分證號為搜尋條件 by AMU 20090417
                'Dim sStr As String="SELECT * FROM dbo.STUD_ENTERTEMP2 WHERE IDNO='" & IDNOvalue & "'" 'ORDER BY MODIFYDATE DESC
                'dt=DbAccess.GetDataTable(sStr, da, trans)
                If TIMS.dtNODATA(dt2) Then
                    ieSETID = TIMS.CINT1(TIMS.Get_eSETID_MaxID(IDNOvalue, TransConn, Trans))
                    'iParms.Add("MILITARYID", MILITARYID)
                    Dim iParms As New Hashtable From {
                        {"ESETID", ieSETID},
                        {"SETID", If(iSETID > 0, iSETID, Convert.DBNull)},
                        {"IDNO", IDNOvalue},
                        {"NAME", Name.Text},
                        {"SEX", If(Sex.Items(0).Selected = True, "M", "F")},
                        {"BIRTHDAY", Birthday.Text},
                        {"PASSPORTNO", If(PassPortNO.Items(0).Selected = True, 1, 2)},
                        {"MARITALSTATUS", If(v_MaritalStatus = "1", v_MaritalStatus, If(v_MaritalStatus = "2", v_MaritalStatus, Convert.DBNull))},
                        {"DEGREEID", v_DegreeID},
                        {"GRADID", v_GraduateStatus},
                        {"SCHOOL", If(School.Text <> "", School.Text, TIMS.cst_未填寫)},
                        {"DEPARTMENT", If(Department.Text <> "", Department.Text, TIMS.cst_未填寫)},
                        {"ZIPCODE", ZipCode1.Value},
                        {"ZIPCODE6W", If(hidZipCode1_6W.Value <> "", hidZipCode1_6W.Value, Convert.DBNull)},
                        {"ADDRESS", Address.Text},
                        {"PHONE1", PhoneD.Text},
                        {"PHONE2", PhoneN.Text},
                        {"CELLPHONE", CellPhone.Text},
                        {"EMAIL", If(dr2_EMAIL <> "", dr2_EMAIL, Convert.DBNull)},
                        {"ISAGREE", If(v_IsAgree = "N", v_IsAgree, "Y")},
                        {"MODIFYACCT", sm.UserInfo.UserID}
                    }
                    'iParms.Add("MODIFYDATE", MODIFYDATE)'iParms.Add("ZIPCODE2W", ZIPCODE2W)'iParms.Add("LAINFLAG", LAINFLAG)'iParms.Add("ZIPCODE_N", ZIPCODE_N)
                    Dim SSQL_i As String = ""
                    SSQL_i &= " INSERT INTO STUD_ENTERTEMP2(ESETID,SETID,IDNO,NAME,SEX,BIRTHDAY,PASSPORTNO,MARITALSTATUS,DEGREEID,GRADID,SCHOOL,DEPARTMENT"
                    SSQL_i &= " ,ZIPCODE,ZIPCODE6W,ADDRESS,PHONE1,PHONE2,CELLPHONE,EMAIL,ISAGREE,MODIFYACCT,MODIFYDATE)" & vbCrLf
                    SSQL_i &= " VALUES (@ESETID,@SETID,@IDNO,@NAME,@SEX,@BIRTHDAY,@PASSPORTNO,@MARITALSTATUS,@DEGREEID,@GRADID,@SCHOOL,@DEPARTMENT"
                    SSQL_i &= " ,@ZIPCODE,@ZIPCODE6W,@ADDRESS,@PHONE1,@PHONE2,@CELLPHONE,@EMAIL,@ISAGREE,@MODIFYACCT,GETDATE())" & vbCrLf
                    ''新增  '如未有此學員的線上報名資料.則新增一筆報名學員資料
                    'dr=dt.NewRow() 'dt.Rows.Add(dr) ''取得 eSETID_MaxID
                    'ieSETID=Val(TIMS.Get_eSETID_MaxID(IDNOvalue, TransConn, trans))
                    'dr("eSETID")=ieSETID 'dr("IDNO")=IDNOvalue
                    DbAccess.ExecuteNonQuery(SSQL_i, Trans, iParms)
                Else
                    '修改 'dr=dt.Rows(0) 'ieSETID=Val(dr("eSETID"))
                    ieSETID = dr2_ESETID
                    'uParmss.Add("MILITARYID", MILITARYID)
                    Dim uParms As New Hashtable From {
                        {"SETID", If(iSETID > 0, iSETID, Convert.DBNull)},
                        {"IDNO", IDNOvalue},
                        {"NAME", Name.Text},
                        {"SEX", If(Sex.Items(0).Selected = True, "M", "F")},
                        {"BIRTHDAY", Birthday.Text},
                        {"PASSPORTNO", If(PassPortNO.Items(0).Selected = True, 1, 2)},
                        {"MARITALSTATUS", If(v_MaritalStatus = "1", v_MaritalStatus, If(v_MaritalStatus = "2", v_MaritalStatus, Convert.DBNull))},
                        {"DEGREEID", v_DegreeID},
                        {"GRADID", v_GraduateStatus},
                        {"SCHOOL", If(School.Text <> "", School.Text, TIMS.cst_未填寫)},
                        {"DEPARTMENT", If(Department.Text <> "", Department.Text, TIMS.cst_未填寫)},
                        {"ZIPCODE", ZipCode1.Value},
                        {"ZIPCODE6W", If(hidZipCode1_6W.Value <> "", hidZipCode1_6W.Value, Convert.DBNull)},
                        {"ADDRESS", Address.Text},
                        {"PHONE1", PhoneD.Text},
                        {"PHONE2", PhoneN.Text},
                        {"CELLPHONE", CellPhone.Text},
                        {"EMAIL", If(dr2_EMAIL <> "", dr2_EMAIL, Convert.DBNull)},
                        {"ISAGREE", If(v_IsAgree = "N", v_IsAgree, "Y")},
                        {"MODIFYACCT", sm.UserInfo.UserID},
                        {"ESETID", ieSETID}
                    }
                    Dim SSQL_u As String = ""
                    SSQL_u &= " UPDATE STUD_ENTERTEMP2" & vbCrLf
                    SSQL_u &= " SET SETID=@SETID,IDNO=@IDNO,NAME=@NAME,SEX=@SEX,BIRTHDAY=@BIRTHDAY,PASSPORTNO=@PASSPORTNO,MARITALSTATUS=@MARITALSTATUS" & vbCrLf
                    SSQL_u &= " ,DEGREEID=@DEGREEID,GRADID=@GRADID,SCHOOL=@SCHOOL,DEPARTMENT=@DEPARTMENT,ZIPCODE=@ZIPCODE,ZIPCODE6W=@ZIPCODE6W" & vbCrLf
                    SSQL_u &= " ,ADDRESS=@ADDRESS,PHONE1=@PHONE1,PHONE2=@PHONE2,CELLPHONE=@CELLPHONE,EMAIL=@EMAIL,ISAGREE=@ISAGREE" & vbCrLf
                    SSQL_u &= " ,MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf ',ZIPCODE2W=@ZIPCODE2W,LAINFLAG=@LAINFLAG,ZIPCODE_N=@ZIPCODE_N
                    SSQL_u &= " WHERE ESETID=@ESETID" & vbCrLf
                    DbAccess.ExecuteNonQuery(SSQL_u, Trans, uParms)
                End If

                aNow = TIMS.GetSysDateNow(Trans)
                s_aToday = Common.FormatDate(aNow) 'Today()
                'SAVE_STUD_ENTERTYPE2 'htPP1.Add("iSignNo", iSignNo)
                Dim htPP1 As New Hashtable From {
                    {"MIdentityID", v_MIdentityID},
                    {"IDNOvalue", IDNOvalue},
                    {"OCID1_Value", OCID1_Value},
                    {"iSETID", iSETID},
                    {"iSerNum", iSerNum},
                    {"ieSETID", ieSETID},
                    {"ModifyAcct", sm.UserInfo.UserID}
                }

                OKFlag = SAVE_STUD_ENTERTYPE2(Trans, htPP1, drCC1, ieSerNum, iSignNo, Errmsg, aNow)

                DbAccess.CommitTrans(Trans)
                OKFlag = True '正常
            Catch ex As Exception
                TIMS.LOG.Error(ex.Message, ex)
                vMsg = ex.Message
                OKFlag = False '異常
                ieSETID = 0
                DbAccess.RollbackTrans(Trans)
                DbAccess.CloseDbConn(TransConn) '(下面應該是用不到了)
                Errmsg = "報名失敗，該學員報名資料異常 或 資料庫異常，請重試!" & vbCrLf
                Errmsg &= "請再試一次，造成您不便之處，還請見諒。" & vbCrLf
                Errmsg &= String.Concat("(若持續出現此問題，請聯絡系統管理者)!", vbCrLf, vMsg)
                Return False
            End Try


            '(無意外，但有錯誤提醒資訊／或不寫入，直接離開)
            If Errmsg <> "" OrElse Not OKFlag OrElse iSignNo <= 0 Then
                DbAccess.RollbackTrans(Trans)
                DbAccess.CloseDbConn(TransConn) '(下面應該是用不到了)
                Errmsg = "報名失敗，該學員報名資料異常 或 資料庫異常，請重試!!" & vbCrLf
                Errmsg &= "請再試一次，造成您不便之處，還請見諒。" & vbCrLf
                Errmsg &= String.Concat("(若持續出現此問題，請聯絡系統管理者)!!", vbCrLf, vMsg)
                Return False
            End If

        End Using

        Try
            'SAVE_STUD_ENTERTRAIN2
            Dim htPP1 As New Hashtable From {{"MIdentityID", v_MIdentityID}}
            Call SAVE_STUD_ENTERTRAIN2(objconn, htPP1, ieSerNum)
            OKFlag = True
            Errmsg = String.Concat(Cst_SuccessMsg1, vbCrLf)
        Catch ex As Exception
            TIMS.LOG.Error(ex.Message, ex)
            vMsg = ex.Message
            OKFlag = False
            Errmsg = "報名失敗，該學員報名資料異常 或 資料庫異常，請重試!!!" & vbCrLf
            Errmsg &= "請再試一次，造成您不便之處，還請見諒。" & vbCrLf
            Errmsg &= String.Concat("(若持續出現此問題，請聯絡系統管理者)!!!", vbCrLf, vMsg)
            Return False
        End Try

        Dim TxtMessage As String = ""
        Try
            Call Check_GovCost(IDNOvalue, OCID1_Value, TxtMessage)

            If OKFlag Then
                If TxtMessage <> "" Then
                    JaveScriptAlertMsg = String.Concat("<script>", "alert('", Cst_SuccessMsg1, "');", "location.href='", Redirect1, "';", "</script>")
                End If
                Errmsg = String.Concat(Cst_SuccessMsg1, vbCrLf)
            End If
        Catch ex As Exception
            TIMS.LOG.Error(ex.Message, ex)
            If OKFlag Then
                OKFlag = False
                Errmsg &= "報名資料上傳成功，但儲存時有其他資料異常狀況(學員資料)!!!" & vbCrLf
            Else
                'OKFlag=False
                Errmsg &= "報名資料上傳失敗!!!" & vbCrLf
            End If
        End Try
        Return OKFlag
    End Function

    ''' <summary>線上報名資料寫入 STUD_ENTERTYPE2 </summary>
    ''' <param name="trans"></param>
    ''' <param name="htPP1"></param>
    ''' <param name="drCC1"></param>
    ''' <param name="ieSerNum"></param>
    ''' <param name="Errmsg"></param>
    ''' <returns></returns>
    Public Shared Function SAVE_STUD_ENTERTYPE2(ByRef trans As SqlTransaction, ByRef htPP1 As Hashtable, ByRef drCC1 As DataRow,
                                                ByRef ieSerNum As Integer, ByRef iSignNo As Integer, ByRef Errmsg As String, ByRef aNow As Date) As Boolean
        Dim rst As Boolean = False 'false:異常/true:正常
        Dim v_MIdentityID As String = TIMS.GetMyValue2(htPP1, "MIdentityID")
        Dim IDNOvalue As String = TIMS.GetMyValue2(htPP1, "IDNOvalue")
        Dim OCID1_Value As String = TIMS.CINT1(TIMS.GetMyValue2(htPP1, "OCID1_Value"))

        Dim iSETID As Integer = TIMS.CINT1(TIMS.GetMyValue2(htPP1, "iSETID"))
        Dim iSerNum As Integer = TIMS.CINT1(TIMS.GetMyValue2(htPP1, "iSerNum"))
        Dim ieSETID As Integer = TIMS.CINT1(TIMS.GetMyValue2(htPP1, "ieSETID"))
        Dim s_ModifyAcct As String = TIMS.GetMyValue2(htPP1, "ModifyAcct")
        'Dim ieSerNum As Integer=0 '線上報名資料寫入Stud_EnterType2中---start
        Dim dr As DataRow = Nothing
        Dim da As SqlDataAdapter = Nothing
        Dim sStr As String = $"SELECT * FROM dbo.STUD_ENTERTYPE2 WHERE eSETID={ieSETID} AND OCID1={OCID1_Value}"
        Dim dt As DataTable = DbAccess.GetDataTable(sStr, da, trans)
        If TIMS.dtHaveDATA(dt) Then
            'DbAccess.CommitTrans(trans) ' ViewState("aNow")=TIMS.CssFormatDate(aNow)
            'Errmsg &= " 目前系統時間為(" &  ViewState("aNow") & ")!!" 'Common.MessageBox(Me, Errmsg)
            Errmsg = " 報名失敗，您已經有報名此班級了!!!" & vbCrLf '\n\n
            Return False 'false:異常/true:正常
        End If

        Const cst_err1 As String = "無法取得報名序號!"
        Dim OKFlag As Boolean = False
        Try
            iSignNo = TIMS.GetSignNoxEnterType3(trans, trans.Connection, OCID1_Value)
        Catch ex As Exception
            TIMS.LOG.Error(String.Concat(cst_err1, ex.Message), ex)
            iSignNo = -1
            Dim excep As New Exception(cst_err1)
            Throw excep
            'Dim excep As New Exception("無法取得報名序號!") Throw ex
        End Try
        'OKFlag=If(iSignNo > 0, True, False)  'false:異常/true:正常
        'If Not OKFlag Then,'DbAccess.RollbackTrans(trans),Dim excep As New Exception("無法取得報名序號!"),Throw excep,End If,
        '新增
        dr = dt.NewRow()
        dt.Rows.Add(dr)
        ieSerNum = TIMS.GETMAX_ESERNUM_SEQ(trans)
        dr("eSerNum") = ieSerNum 'PK
        dr("eSETID") = ieSETID 'FK
        dr("SETID") = If(iSETID > 0, iSETID, Convert.DBNull) 'FK
        dr("SerNum") = iSerNum
        dr("RelEnterDate") = aNow 'Now()
        dr("OCID1") = TIMS.CINT1(OCID1_Value)
        dr("EnterDate") = Common.FormatDate(aNow) 'Today()
        dr("SIGNNO") = iSignNo '報名序號
        '=====START=====取得班級資料，填入學員報名職類基本資料
        dr("TMID1") = drCC1("TMID")
        dr("RID") = drCC1("RID")
        dr("PlanID") = drCC1("PLANID")
        dr("MIDENTITYID") = v_MIdentityID
        dr("IDENTITYID") = v_MIdentityID ' MIDENTITYID.SELECTEDVALUE 'IDENTITYID.SELECTEDVALUE
        dr("EnterPath") = "o" '產投外網(內部小寫o) by AMU 201212
        dr("ModifyAcct") = s_ModifyAcct 'sm.UserInfo.UserID 'IDNOvalue
        dr("ModifyDate") = aNow 'Now()
        '=====END=====取得班級資料，填入學員報名職類基本資料
        DbAccess.UpdateDataTable(dt, da, trans)
        'DbAccess.CommitTrans(trans)
        Return True
    End Function

    ''' <summary>>線上報名資料寫入 STUD_ENTERTRAIN2 </summary>
    ''' <param name="tConn"></param>
    ''' <param name="htPP1"></param>
    ''' <param name="ieSerNum"></param>
    Sub SAVE_STUD_ENTERTRAIN2(ByRef tConn As SqlConnection, ByRef htPP1 As Hashtable, ByRef ieSerNum As Integer)
        Dim v_MIdentityID As String = TIMS.GetMyValue2(htPP1, "MIdentityID")

        '線上報名資料寫入Stud_EnterTrain2中---start
        Dim dr As DataRow = Nothing
        Dim da As SqlDataAdapter = Nothing
        Dim sStr As String = $" SELECT * FROM dbo.STUD_ENTERTRAIN2 WHERE ESERNUM={ieSerNum}"
        Dim dt As DataTable = DbAccess.GetDataTable(sStr, da, tConn)

        Dim iSEID As Integer = 0
        If TIMS.dtNODATA(dt) Then
            iSEID = DbAccess.GetNewId(tConn, "STUD_ENTERTRAIN2_SEID_SEQ,STUD_ENTERTRAIN2,SEID")
            dr = dt.NewRow()
            dt.Rows.Add(dr)
            dr("SEID") = iSEID
            dr("eSerNum") = ieSerNum
        Else
            dr = dt.Rows(0)
            iSEID = TIMS.CINT1(dr("SEID"))
        End If
        dr("SERVDEPTID") = TIMS.GetValue1(TIMS.GetListValue(ddlSERVDEPTID))
        dr("JOBTITLEID") = TIMS.GetValue1(TIMS.GetListValue(ddlJOBTITLEID))

        If ZipCode2.Value <> "" Then
            dr("ZipCode2") = ZipCode2.Value
            hidZipCode2_6W.Value = TIMS.GetZIPCODE6W(ZipCode2.Value, ZipCode2_B3.Value)
            dr("ZipCode2_6W") = If(hidZipCode2_6W.Value <> "", hidZipCode2_6W.Value, Convert.DBNull)
            dr("HouseholdAddress") = HouseholdAddress.Text
        Else
            If CheckBox1.Checked Then
                dr("ZipCode2") = ZipCode1.Value
                hidZipCode2_6W.Value = TIMS.GetZIPCODE6W(ZipCode1.Value, ZipCode1_B3.Value)
                dr("ZipCode2_6W") = If(hidZipCode2_6W.Value <> "", hidZipCode2_6W.Value, Convert.DBNull)
                dr("HouseholdAddress") = Address.Text
            End If
        End If
        dr("MidentityID") = v_MIdentityID ' MIdentityID.SelectedValue
        'dr("PriorWorkPay")=If(PriorWorkPay.Text <> "", Val(PriorWorkPay.Text), Convert.DBNull) '受訓前薪資

        For int_i As Integer = 0 To AcctMode.Items.Count - 1
            Select Case int_i
                Case 0
                    PostNo_1.Text = TIMS.ClearSQM(PostNo_1.Text)
                    AcctNo1_1.Text = TIMS.ClearSQM(AcctNo1_1.Text)
                    If AcctMode.Items(0).Selected = True Then
                        dr("AcctMode") = 0
                        dr("PostNo") = PostNo_1.Text '+ "-" +  PostNo_2.Text
                        dr("AcctNo") = AcctNo1_1.Text '+ "-" +  AcctNo1_2.Text
                    End If
                Case 1
                    If AcctMode.Items(1).Selected = True Then
                        dr("AcctMode") = 1
                        dr("AcctHeadNo") = AcctheadNo.Text
                        dr("BankName") = BankName.Text
                        dr("AcctExNo") = AcctExNo.Text
                        dr("ExBankName") = ExBankName.Text
                        dr("AcctNo") = AcctNo2.Text
                    End If
                Case 2
                    If AcctMode.Items(2).Selected = True Then
                        dr("AcctMode") = 2
                    End If
            End Select
        Next

        If Uname.Text <> "" Then dr("Uname") = Uname.Text
        If Intaxno.Text <> "" Then dr("Intaxno") = Intaxno.Text

        '服務部門 ServDept 30 CHAR
        Dim t_ddlSERVDEPTID As String = TIMS.GetListText(ddlSERVDEPTID)
        dr("ServDept") = If(t_ddlSERVDEPTID <> "", TIMS.GetValue1(t_ddlSERVDEPTID), TIMS.GetValue1(ServDept.Text))
        dr("ActNo") = If(ActNo.Text <> "", ActNo.Text, Convert.DBNull)
        dr("ActName") = If(ActName.Text <> "", ActName.Text, Convert.DBNull)
        '**by Milor 20080904--投保單位電話與地址----start 
        dr("ActTel") = Convert.DBNull
        dr("ZipCode3") = Convert.DBNull
        hidZipCode3_6W.Value = TIMS.GetZIPCODE6W(ZipCode3.Value, ZipCode3_B3.Value)
        dr("ZipCode3_6W") = If(hidZipCode3_6W.Value <> "", hidZipCode3_6W.Value, Convert.DBNull)
        dr("ActAddress") = Convert.DBNull

        ActTel.Text = TIMS.ClearSQM(ActTel.Text)
        ActAddress.Text = TIMS.ClearSQM(ActAddress.Text)
        'If ActTel.Text <> "" Then ActTel.Text=Trim(ActTel.Text)
        'If ActAddress.Text <> "" Then ActAddress.Text=Trim(ActAddress.Text)
        If ActTel.Text <> "" Then dr("ActTel") = ActTel.Text '.Trim.Replace("'", "''")

        ZipCode3.Value = TIMS.TrimZipCode(ZipCode3.Value, dtZipCode)
        hidZipCode3_6W.Value = TIMS.GetZIPCODE6W(ZipCode3.Value, ZipCode3_B3.Value)
        dr("ZipCode3") = If(ZipCode3.Value <> "", ZipCode3.Value, Convert.DBNull)
        dr("ZipCode3_6W") = If(hidZipCode3_6W.Value <> "", hidZipCode3_6W.Value, Convert.DBNull)
        dr("ActAddress") = If(ActAddress.Text <> "", ActAddress.Text, Convert.DBNull)
        '**by Milor 20080904--投保單位電話與地址----end
        Dim v_ActType As String = TIMS.GetListValue(ActType)
        dr("ActType") = If(v_ActType <> "", v_ActType, Convert.DBNull)

        JobTitle.Text = TIMS.ClearSQM(JobTitle.Text)
        Dim v_ddlJOBTITLEID As String = TIMS.GetListValue(ddlJOBTITLEID)
        Dim t_ddlJOBTITLEID As String = TIMS.GetListText(ddlJOBTITLEID)
        Dim v_JobTitle As String = ""
        v_JobTitle = If(t_ddlJOBTITLEID <> "" AndAlso v_ddlJOBTITLEID <> "", TIMS.GetValue1(t_ddlJOBTITLEID), TIMS.GetValue1(JobTitle.Text))
        dr("JobTitle") = If(v_JobTitle <> "", v_JobTitle, Convert.DBNull)
        'If  JobTitle.Text <> "" Then dr("JobTitle")= JobTitle.Text

        '=====任職公司其他資料地址=====
        dr("Zip") = "-1" '97產業人才投資方案取消輸入
        dr("Addr") = " " '97產業人才投資方案取消輸入
        dr("Tel") = " " '97產業人才投資方案取消輸入
        dr("ShowDetail") = " " '97產業人才投資方案取消輸入
        'dr("Q1")=0
        '是否由公司推薦參訓 Q1
        Select Case Q1.SelectedValue
            Case "0", "1"
                dr("Q1") = Q1.SelectedValue
            Case Else
                dr("Q1") = 0 'Convert.DBNull
        End Select
        dr("Q2_1") = If(Q2.Items(0).Selected, 1, 2)
        dr("Q2_2") = If(Q2.Items(1).Selected, 1, 2)
        dr("Q2_3") = If(Q2.Items(2).Selected, 1, 2)
        dr("Q2_4") = If(Q2.Items(3).Selected, 1, 2)
        Dim v_Q3 As String = TIMS.GetListValue(Q3)
        If v_Q3 <> "" Then
            dr("Q3") = v_Q3
            If Q3.SelectedValue = "3" Then
                If Q3_Other.Text <> "" Then dr("Q3_Other") = Q3_Other.Text
            End If
        End If

        Dim v_Q4 As String = TIMS.GetListValue(Q4)
        dr("Q4") = If(v_Q4 <> "", v_Q4, Convert.DBNull)
        dr("Q5") = If(Q5.Items(0).Selected, 1, 0)
        Q61.Text = TIMS.ClearSQM(Q61.Text)
        Q62.Text = TIMS.ClearSQM(Q62.Text)
        Q63.Text = TIMS.ClearSQM(Q63.Text)
        Q64.Text = TIMS.ClearSQM(Q64.Text)
        dr("Q61") = If(Q61.Text <> "", TIMS.VAL1(Q61.Text), Convert.DBNull)
        dr("Q62") = If(Q62.Text <> "", TIMS.VAL1(Q62.Text), Convert.DBNull)
        dr("Q63") = If(Q63.Text <> "", TIMS.VAL1(Q63.Text), Convert.DBNull)
        dr("Q64") = If(Q64.Text <> "", TIMS.VAL1(Q64.Text), Convert.DBNull)

        '定期收到產業人才投資方案最新課程資訊
        Dim v_IseMail As String = TIMS.GetListValue(IseMail)
        dr("IseMail") = If(v_IseMail <> "", v_IseMail, Convert.DBNull)
        'If IseMail.Items(0).Selected=True Then dr("IseMail")="Y" Else dr("IseMail")="N"
        'If  IsAgreedata.Items(0).Selected=True Then dr("IseMail")="Y" Else dr("IseMail")="N"
        'If  IsAgreedata.SelectedValue="Y" Then dr("IseMail")="Y" Else dr("IseMail")="N"
        dr("ModifyAcct") = sm.UserInfo.UserID 'IDNOvalue  '修改者為學員本身
        dr("ModifyDate") = aNow 'Now()
        DbAccess.UpdateDataTable(dt, da)
        'DbAccess.UpdateDataTable(dt, da, trans)
        'DbAccess.CommitTrans(trans)
    End Sub

    '輔助金使用檢核
    Function Check_GovCost(ByVal aIDNO As String, ByVal OCID1 As String, ByRef TxtMessage As String) As Boolean
        Dim Check_GovCostFlag As Boolean = True 'True-正常:可以用補助/異常:補助額不足，將另尋解決途徑
        TxtMessage = ""
        'AlertType 1-正常:可以用補助(進入繼續報名程序)
        'AlertType 2-異常:補助額度不足，但仍有部份補助額(進入繼續報名程序／終止報名程序)
        'AlertType 3-異常:補助額度已滿(終止報名程序)

        'Dim TxtMessage As String 'message
        ''(限定產業人才投資方案) 20090325 BY AMU
        ''Dim ActSubsidyCost As String=TIMS.Get_ActSubsidyCost28(IDNO) '已實際請領補助費(限定產業人才投資方案)
        ''Dim SignSubsidyCost As String=TIMS.Get_SignSubsidyCost28(IDNO) '已報名申請補助費(限定產業人才投資方案)
        ''Dim DefGovCost As String=TIMS.Get_DefGeoCost28(IDNO) '(=線上報名預算的政府補助)(全部)(限定產業人才投資方案)

        'Const Cst_MaxCanUseCost=50000 '三年內最大可用餘額2007年前為３萬, 2008年改為５萬
        'Const Cst_AlertCost=40000 '警示額
        Dim LOCIDdate As String = "" '本班的結訓日期 '本班的開訓日期 (最後使用經費日期)
        Dim ActSubsidyCost As String = "" '已實際請領補助費(限定產業人才投資方案)
        Dim SignSubsidyCost As String = "" '已報名申請補助費(限定產業人才投資方案)
        Dim DefGovCost As String = "" '(=線上報名預算的政府補助)(全部)(限定產業人才投資方案)
        Dim ccSTDate As String = ""
        Dim ccFTDate As String = ""
        Dim ccClsDefGovCost As String = ""
        Dim ccClsTtlCost As String = ""
        Dim LimitCost As Double ''最後可用政府補助經費
        'Dim ClsDefGovCost As String="" '此班級的政府補助額－每人費用－本班政府補助預算
        'Dim LimitCost As Double ''最後可用政府補助經費
        'Dim GovCost As String '''您已使用政府補助經費
        'Dim dr1 As DataRow
        'Dim objtable As DataTable
        '代入初始資料(Class_ClassInfo)
        '此班級的結訓日期 (最後使用經費日期)
        '此班級 每人政府所補助的費用
        '加入要開班與尚未達到結訓日的條件( notopen='N' OR STDate>=getdate() )
        Dim OKFlag2 As Boolean = True '資料庫連結正常 True/ 異常 False
        Dim PMS1 As New Hashtable From {{"OCID1", TIMS.CINT1(OCID1)}}
        Dim str As String = ""
        str &= " SELECT CONVERT(varchar, cc.STDate, 111) STDate" & vbCrLf
        str &= " ,CONVERT(varchar, cc.FTDate, 111) FTDate" & vbCrLf
        'str &= " ,format((pp.DefGovCost/cc.TNum), '#######0.##') ClsDefGovCost " & vbCrLf
        str &= " ,CASE WHEN cc.TNum <> 0 THEN FORMAT((pp.DefGovCost/cc.TNum), '#######0.##') ELSE '0' END ClsDefGovCost" & vbCrLf
        'str &= " ,format(TRUNC(pp.TotalCost/cc.TNum), '#######0.##') ClsTtlCost " & vbCrLf
        str &= " ,CASE WHEN cc.TNum <> 0 THEN FORMAT(FLOOR(pp.TotalCost/cc.TNum), '#######0.##') ELSE '0' END ClsTtlCost" & vbCrLf
        str &= " FROM dbo.CLASS_CLASSINFO cc WITH(NOLOCK)" & vbCrLf
        str &= " JOIN dbo.PLAN_PLANINFO pp WITH(NOLOCK) ON pp.ComIDNO=cc.ComIDNO AND pp.PlanID=cc.PlanID AND pp.SeqNO=cc.SeqNO" & vbCrLf
        str &= " WHERE cc.OCID=@OCID1" & vbCrLf
        Dim sqldr As DataRow = DbAccess.GetOneRow(str, objconn, PMS1)
        If sqldr IsNot Nothing Then
            ccSTDate = Convert.ToString(sqldr("STDate"))
            ccFTDate = Convert.ToString(sqldr("FTDate"))
            ccClsDefGovCost = Convert.ToString(sqldr("ClsDefGovCost"))
            ccClsTtlCost = Convert.ToString(sqldr("ClsTtlCost"))
        End If
        If sqldr Is Nothing Then OKFlag2 = False

        LOCIDdate = ccSTDate  '本班的開訓日期 (最後使用經費日期)
        Dim sDate As String = String.Empty
        Dim eDate As String = String.Empty
        Dim ClsDefGovCost As String = "" '此班級的政府補助額－每人費用－本班政府補助預算
        Dim ClsTtlCost As String = "" '課程總費用
        'Dim GovCost As String '''您已使用政府補助經費

        'Dim objtable As DataTable
        'Dim dr1 As DataRow
        Call TIMS.Get_SubSidyCostDay(aIDNO, LOCIDdate, sDate, eDate, objconn)
        ActSubsidyCost = TIMS.Get_ActSubsidyCost28(aIDNO, sDate, eDate, objconn) '(本期) 已實際請領補助費(限定產業人才投資方案)
        SignSubsidyCost = TIMS.Get_SignSubsidyCost28(aIDNO, sDate, eDate, objconn) '(本期) 已報名申請補助費(限定產業人才投資方案)
        DefGovCost = TIMS.Get_DefGeoCost28(aIDNO, sDate, eDate, objconn) '(本期) (=線上報名預算的政府補助)(全部)(限定產業人才投資方案)

        ClsDefGovCost = ccClsDefGovCost '此班級的政府補助額－每人費用－本班政府補助預算
        ClsTtlCost = ccClsTtlCost

        '產投 政府補助經費
        If ActSubsidyCost < TIMS.Get_3Y_SupplyMoney() Then
            LimitCost = TIMS.Get_3Y_SupplyMoney() - (CInt(ActSubsidyCost) + CInt(DefGovCost))  '可用政府補助經費(剩餘可用餘額)
            'If LimitCost < 0 Then LimitCost=0
            ViewState("LimitCost") = LimitCost.ToString
            'KeepSearch()
            If LimitCost >= ClsDefGovCost Then
                Check_GovCost = True
                'Check_GovCostFlag=True 'AlertType=1
                TxtMessage = ""
                TxtMessage &= " 已報名申請補助費 " & CInt(SignSubsidyCost) + CInt(DefGovCost) & "元\n"
                TxtMessage &= " 已實際請領補助費 " & ActSubsidyCost & "元\n"
                'TxtMessage &= " 該班級補助申請款為" & ClsDefGovCost & "元，未超出補助額度\n"
                TxtMessage &= " 尚餘經費 " & LimitCost.ToString & " 元，可供申請\n"
            Else
                Check_GovCost = False
                Check_GovCostFlag = False
                'AlertType=2
                TxtMessage = ""
                TxtMessage &= " 已報名申請補助費 " & CInt(SignSubsidyCost) + CInt(DefGovCost) & "元\r\n"
                TxtMessage &= " 已實際請領補助費 " & ActSubsidyCost & "元\r\n"
                TxtMessage &= " 尚餘經費 " & LimitCost.ToString & " 元，可供申請\r\n"
                TxtMessage &= " 該班級補助申請款為" & ClsDefGovCost & "元，將超出補助額度\r\n"
                TxtMessage &= " 可使用補助餘額為 " & LimitCost.ToString & " 元, 是否同意繼續報名\r\n"
            End If
        Else
            LimitCost = 0
            Check_GovCost = False
            Check_GovCostFlag = False
            'AlertType=3
            TxtMessage = ""
            TxtMessage &= " 已報名申請補助費 " & CInt(SignSubsidyCost) + CInt(DefGovCost) & "元\r\n"
            TxtMessage &= " 已實際請領補助費 " & ActSubsidyCost & "元\r\n"
            'TxtMessage &= " 尚餘經費 " & LimitCost.ToString & " 元 \r\n"
            TxtMessage &= " 3年5萬補助額度已滿。\r\n"
        End If

        Return Check_GovCostFlag
    End Function

    Sub UPDATE_STUD_ENTERTEMP3(ByVal eSETID3 As Integer, ByVal aIDNO As String, ByRef tConn As SqlConnection)
        If eSETID3 <= 0 OrElse aIDNO = "" Then Return

        PhoneD.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(PhoneD.Text))
        PhoneN.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(PhoneN.Text))
        CellPhone.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(CellPhone.Text))
        Email.Text = TIMS.ChangeEmail(TIMS.ClearSQM(Email.Text))
        Dim vEmail As String = Email.Text
        'If Email.Text <> "" AndAlso Email.Text <> "無" Then '有資料時修改／沒資料時不變
        '    vEmail=Email.Text
        'Else
        '    vEmail="無"
        '    Email.Text=vEmail
        'End If
        Dim v_MIdentityID As String = TIMS.GetListValue(MIdentityID)
        ' PostNo_1.Text=TIMS.ClearSQM( PostNo_1.Text)
        ' AcctNo1_1.Text=TIMS.ClearSQM( AcctNo1_1.Text)

        Dim v_MaritalStatus As String = TIMS.GetListValue(MaritalStatus)

        Call TIMS.OpenDbConn(tConn)

        Dim dt1 As New DataTable
        Dim sql As String = "SELECT 1 FROM STUD_ENTERTEMP3 WHERE IDNO=@IDNO AND ESETID3=@ESETID3"
        Dim sCmd1 As New SqlCommand(sql, tConn)
        With sCmd1
            .Parameters.Add("IDNO", SqlDbType.VarChar).Value = aIDNO
            .Parameters.Add("ESETID3", SqlDbType.VarChar).Value = eSETID3
            dt1.Load(.ExecuteReader())
        End With
        If TIMS.dtNODATA(dt1) Then Return

        Dim uSql As String = ""
        uSql &= " UPDATE STUD_ENTERTEMP3" & vbCrLf
        uSql &= " SET NAME=@NAME ,SEX=@SEX ,BIRTHDAY=@BIRTHDAY ,PASSPORTNO=@PASSPORTNO ,MARITALSTATUS=@MARITALSTATUS" & vbCrLf
        uSql &= " ,DEGREEID=@DEGREEID ,GRADID=@GRADID ,SCHOOL=@SCHOOL ,DEPARTMENT=@DEPARTMENT" & vbCrLf
        uSql &= " ,ZIPCODE1=@ZIPCODE1 ,ZIPCODE1_6W=@ZIPCODE1_6W ,ADDRESS=@ADDRESS ,ZIPCODE2=@ZIPCODE2 ,ZIPCODE2_6W=@ZIPCODE2_6W" & vbCrLf
        uSql &= " ,HOUSEHOLDADDRESS=@HOUSEHOLDADDRESS ,PHONE1=@PHONE1 ,PHONE2=@PHONE2 ,CELLPHONE=@CELLPHONE ,EMAIL=@EMAIL ,MIDENTITYID=@MIDENTITYID" & vbCrLf
        'sql &= " ,HANDTYPEID=@HANDTYPEID ,HANDLEVELID=@HANDLEVELID,PRIORWORKPAY=@PRIORWORKPAY" & vbCrLf
        uSql &= " ,ACCTMODE=@ACCTMODE ,POSTNO=@POSTNO ,ACCTHEADNO=@ACCTHEADNO ,BANKNAME=@BANKNAME" & vbCrLf
        uSql &= " ,ACCTEXNO=@ACCTEXNO ,EXBANKNAME=@EXBANKNAME ,ACCTNO=@ACCTNO ,UNAME=@UNAME ,INTAXNO=@INTAXNO ,SERVDEPT=@SERVDEPT" & vbCrLf
        uSql &= " ,ACTNAME=@ACTNAME ,ACTTYPE=@ACTTYPE ,ACTNO=@ACTNO ,ACTTEL=@ACTTEL ,ZIPCODE3=@ZIPCODE3 ,ZIPCODE3_6W=@ZIPCODE3_6W ,ACTADDRESS=@ACTADDRESS" & vbCrLf
        'sql &= " ,SERVDEPT=@SERVDEPT" & vbCrLf
        uSql &= " ,JOBTITLE=@JOBTITLE ,SERVDEPTID=@SERVDEPTID ,JOBTITLEID=@JOBTITLEID" & vbCrLf
        '是否由公司推薦參訓 Q1
        uSql &= " ,Q1=@Q1 ,Q2_1=@Q2_1 ,Q2_2=@Q2_2 ,Q2_3=@Q2_3 ,Q2_4=@Q2_4 ,Q3=@Q3 ,Q3_OTHER=@Q3_OTHER ,Q4=@Q4 ,Q5=@Q5" & vbCrLf
        uSql &= " ,Q61=@Q61 ,Q62=@Q62 ,Q63=@Q63 ,Q64=@Q64 ,ISEMAIL=@ISEMAIL ,ISAGREE=@ISAGREE ,MODIFYACCT=@MODIFYACCT ,MODIFYDATE=GETDATE()" & vbCrLf
        uSql &= " WHERE IDNO=@IDNO AND ESETID3=@ESETID3"
        Dim uoCmd As New SqlCommand(sql, tConn)
        With uoCmd
            .Parameters.Clear()
            .Parameters.Add("NAME", SqlDbType.NVarChar).Value = Name.Text
            .Parameters.Add("SEX", SqlDbType.VarChar).Value = If(Sex.Items(0).Selected = True, "M", "F")
            .Parameters.Add("BIRTHDAY", SqlDbType.DateTime).Value = CDate(Birthday.Text)
            .Parameters.Add("PASSPORTNO", SqlDbType.Int).Value = If(PassPortNO.Items(0).Selected = True, 1, 2)
            '婚姻狀況 '1.已;2.未 3.暫不提供(預設) 
            Select Case v_MaritalStatus'MaritalStatus.SelectedValue
                Case "1", "2"
                Case Else
                    v_MaritalStatus = ""
            End Select
            .Parameters.Add("MARITALSTATUS", SqlDbType.Int).Value = If(v_MaritalStatus <> "", v_MaritalStatus, Convert.DBNull)
            .Parameters.Add("DEGREEID", SqlDbType.VarChar).Value = DegreeID.SelectedValue
            ', GRADID VARCHAR2(3 CHAR) NOT NULL 
            ', SCHOOL NVARCHAR2(30) 
            ', DEPARTMENT NVARCHAR2(128) 
            If GraduateStatus.SelectedValue <> "" Then
                .Parameters.Add("GRADID", SqlDbType.VarChar).Value = GraduateStatus.SelectedValue
            Else
                .Parameters.Add("GRADID", SqlDbType.VarChar).Value = "01"
            End If
            .Parameters.Add("SCHOOL", SqlDbType.NVarChar).Value = School.Text
            .Parameters.Add("DEPARTMENT", SqlDbType.NVarChar).Value = Department.Text

            .Parameters.Add("ZIPCODE1", SqlDbType.Int).Value = If(ZipCode1.Value <> "", CInt(ZipCode1.Value), Convert.DBNull)
            hidZipCode1_6W.Value = TIMS.GetZIPCODE6W(ZipCode1.Value, ZipCode1_B3.Value)
            .Parameters.Add("ZIPCODE1_6W", SqlDbType.VarChar).Value = If(hidZipCode1_6W.Value <> "", hidZipCode1_6W.Value, Convert.DBNull)
            .Parameters.Add("ADDRESS", SqlDbType.VarChar).Value = Address.Text

            .Parameters.Add("ZIPCODE2", SqlDbType.Int).Value = If(ZipCode2.Value <> "", CInt(ZipCode2.Value), Convert.DBNull)
            hidZipCode2_6W.Value = TIMS.GetZIPCODE6W(ZipCode2.Value, ZipCode2_B3.Value)
            .Parameters.Add("ZIPCODE2_6W", SqlDbType.VarChar).Value = If(hidZipCode2_6W.Value <> "", hidZipCode2_6W.Value, Convert.DBNull)
            .Parameters.Add("HOUSEHOLDADDRESS", SqlDbType.VarChar).Value = HouseholdAddress.Text

            .Parameters.Add("PHONE1", SqlDbType.VarChar).Value = PhoneD.Text
            .Parameters.Add("PHONE2", SqlDbType.VarChar).Value = PhoneN.Text
            .Parameters.Add("CELLPHONE", SqlDbType.VarChar).Value = CellPhone.Text
            .Parameters.Add("EMAIL", SqlDbType.VarChar).Value = vEmail
            .Parameters.Add("MIDENTITYID", SqlDbType.VarChar).Value = v_MIdentityID '.SelectedValue
            '.Parameters.Add("HANDTYPEID", SqlDbType.VarChar).Value=HANDTYPEID
            '.Parameters.Add("HANDLEVELID", SqlDbType.VarChar).Value=HANDLEVELID
            '.Parameters.Add("PRIORWORKPAY", SqlDbType.Int).Value=If(PriorWorkPay.Text <> "", CInt(PriorWorkPay.Text), Convert.DBNull) '受訓前薪資

            Dim strACCTMODE As String = ""
            '0~1或0~2
            For i As Integer = 0 To AcctMode.Items.Count - 1
                If AcctMode.Items(i).Selected Then
                    strACCTMODE = CStr(i)
                    Exit For
                End If
            Next
            If strACCTMODE <> "" Then
                .Parameters.Add("ACCTMODE", SqlDbType.Int).Value = CInt(strACCTMODE)
                Select Case CInt(strACCTMODE)
                    Case 0 '郵局帳號
                        .Parameters.Add("POSTNO", SqlDbType.VarChar).Value = PostNo_1.Text '+ "-" +  PostNo_2.Text
                        .Parameters.Add("ACCTHEADNO", SqlDbType.VarChar).Value = Convert.DBNull
                        .Parameters.Add("BANKNAME", SqlDbType.VarChar).Value = Convert.DBNull
                        .Parameters.Add("ACCTEXNO", SqlDbType.VarChar).Value = Convert.DBNull
                        .Parameters.Add("EXBANKNAME", SqlDbType.VarChar).Value = Convert.DBNull
                        .Parameters.Add("ACCTNO", SqlDbType.VarChar).Value = AcctNo1_1.Text '+ "-" +  AcctNo1_2.Text
                    Case 1 '銀行帳號
                        .Parameters.Add("POSTNO", SqlDbType.VarChar).Value = Convert.DBNull
                        .Parameters.Add("ACCTHEADNO", SqlDbType.VarChar).Value = AcctheadNo.Text
                        .Parameters.Add("BANKNAME", SqlDbType.VarChar).Value = BankName.Text
                        .Parameters.Add("ACCTEXNO", SqlDbType.VarChar).Value = AcctExNo.Text
                        .Parameters.Add("EXBANKNAME", SqlDbType.VarChar).Value = ExBankName.Text
                        .Parameters.Add("ACCTNO", SqlDbType.VarChar).Value = AcctNo2.Text
                    Case Else '2'訓練單位代轉現金
                        .Parameters.Add("POSTNO", SqlDbType.VarChar).Value = Convert.DBNull
                        .Parameters.Add("ACCTHEADNO", SqlDbType.VarChar).Value = Convert.DBNull
                        .Parameters.Add("BANKNAME", SqlDbType.VarChar).Value = Convert.DBNull
                        .Parameters.Add("ACCTEXNO", SqlDbType.VarChar).Value = Convert.DBNull
                        .Parameters.Add("EXBANKNAME", SqlDbType.VarChar).Value = Convert.DBNull
                        .Parameters.Add("ACCTNO", SqlDbType.VarChar).Value = Convert.DBNull
                End Select
            Else
                .Parameters.Add("ACCTMODE", SqlDbType.Int).Value = Convert.DBNull
                .Parameters.Add("POSTNO", SqlDbType.VarChar).Value = Convert.DBNull
                .Parameters.Add("ACCTHEADNO", SqlDbType.VarChar).Value = Convert.DBNull
                .Parameters.Add("BANKNAME", SqlDbType.VarChar).Value = Convert.DBNull
                .Parameters.Add("ACCTEXNO", SqlDbType.VarChar).Value = Convert.DBNull
                .Parameters.Add("EXBANKNAME", SqlDbType.VarChar).Value = Convert.DBNull
                .Parameters.Add("ACCTNO", SqlDbType.VarChar).Value = Convert.DBNull
            End If

            .Parameters.Add("UNAME", SqlDbType.VarChar).Value = Uname.Text
            .Parameters.Add("INTAXNO", SqlDbType.VarChar).Value = Intaxno.Text

            Dim t_ddlSERVDEPTID As String = TIMS.GetListText(ddlSERVDEPTID)
            .Parameters.Add("SERVDEPT", SqlDbType.NVarChar).Value = If(t_ddlSERVDEPTID <> "", TIMS.GetValue1(t_ddlSERVDEPTID), TIMS.GetValue1(ServDept.Text))

            .Parameters.Add("ACTNAME", SqlDbType.VarChar).Value = ActName.Text
            .Parameters.Add("ACTTYPE", SqlDbType.Char, 1).Value = ActType.SelectedValue
            .Parameters.Add("ACTNO", SqlDbType.VarChar).Value = ActNo.Text

            If ActTel.Text <> "" Then ActTel.Text = Trim(ActTel.Text)
            .Parameters.Add("ACTTEL", SqlDbType.VarChar).Value = ActTel.Text

            ZipCode3.Value = TIMS.ClearSQM(ZipCode3.Value)
            ZipCode3_B3.Value = TIMS.ClearSQM(ZipCode3_B3.Value)
            hidZipCode3_6W.Value = TIMS.GetZIPCODE6W(ZipCode3.Value, ZipCode3_B3.Value)
            .Parameters.Add("ZIPCODE3", SqlDbType.VarChar).Value = If(ZipCode3.Value <> "", CInt(ZipCode3.Value), Convert.DBNull)
            .Parameters.Add("ZIPCODE3_6W", SqlDbType.VarChar).Value = If(hidZipCode3_6W.Value <> "", hidZipCode3_6W.Value, Convert.DBNull)
            .Parameters.Add("ACTADDRESS", SqlDbType.VarChar).Value = ActAddress.Text

            '.Parameters.Add("SERVDEPT", SqlDbType.VarChar).Value=SERVDEPT
            Dim t_ddlJOBTITLEID As String = TIMS.GetListText(ddlJOBTITLEID)
            .Parameters.Add("JOBTITLE", SqlDbType.NVarChar).Value = If(t_ddlJOBTITLEID <> "", TIMS.GetValue1(t_ddlJOBTITLEID), TIMS.GetValue1(JobTitle.Text))
            .Parameters.Add("SERVDEPTID", SqlDbType.VarChar).Value = TIMS.GetValue1(ddlSERVDEPTID.SelectedValue)
            .Parameters.Add("JOBTITLEID", SqlDbType.VarChar).Value = TIMS.GetValue1(ddlJOBTITLEID.SelectedValue)
            '是否由公司推薦參訓 Q1
            Select Case Q1.SelectedValue
                Case "0", "1"
                    .Parameters.Add("Q1", SqlDbType.VarChar).Value = Q1.SelectedValue
                Case Else
                    .Parameters.Add("Q1", SqlDbType.VarChar).Value = Convert.DBNull
            End Select
            Dim tmpValue As Integer = 0
            For i As Integer = 1 To 4
                tmpValue = 2
                If Q2.Items(i - 1).Selected Then
                    tmpValue = 1
                End If
                .Parameters.Add(CStr("Q2_" & i), SqlDbType.Int).Value = tmpValue
            Next
            If Q3.SelectedValue <> "" Then
                .Parameters.Add("Q3", SqlDbType.Int).Value = Q3.SelectedValue
                Select Case Q3.SelectedValue
                    Case "3"
                        .Parameters.Add("Q3_OTHER", SqlDbType.NVarChar).Value = Q3_Other.Text
                    Case Else
                        .Parameters.Add("Q3_OTHER", SqlDbType.NVarChar).Value = Convert.DBNull
                End Select
            Else
                .Parameters.Add("Q3", SqlDbType.Int).Value = Convert.DBNull
                .Parameters.Add("Q3_OTHER", SqlDbType.NVarChar).Value = Convert.DBNull
            End If

            .Parameters.Add("Q4", SqlDbType.VarChar).Value = If(Q4.SelectedValue <> "", Q4.SelectedValue, Convert.DBNull)

            tmpValue = 0
            If Q5.Items(0).Selected Then tmpValue = 1 Else tmpValue = 0
            .Parameters.Add("Q5", SqlDbType.Int).Value = tmpValue

            .Parameters.Add("Q61", SqlDbType.Decimal).Value = TIMS.GetNumber(Q61.Text)
            .Parameters.Add("Q62", SqlDbType.Decimal).Value = TIMS.GetNumber(Q62.Text) 'Q62
            .Parameters.Add("Q63", SqlDbType.Decimal).Value = TIMS.GetNumber(Q63.Text) 'Q63
            .Parameters.Add("Q64", SqlDbType.Decimal).Value = TIMS.GetNumber(Q64.Text) 'Q64

            Dim tmpStr As String = ""
            tmpStr = "N"
            If IseMail.Items(0).Selected Then tmpStr = "Y" Else tmpStr = "N"
            .Parameters.Add("ISEMAIL", SqlDbType.Char, 1).Value = tmpStr
            tmpStr = "Y" '預設為同意。
            'If  IsAgree.Items(0).Selected Then tmpStr="Y" Else tmpStr="N"
            .Parameters.Add("ISAGREE", SqlDbType.Char, 1).Value = tmpStr
            .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = aIDNO
            .Parameters.Add("IDNO", SqlDbType.VarChar).Value = aIDNO
            .Parameters.Add("ESETID3", SqlDbType.VarChar).Value = eSETID3

            '.ExecuteNonQuery()
            '為了統一寫入資料異動記錄, 一律改為呼叫 DBAccess.ExecuteNonQuery()
            'DbAccess.ExecuteNonQuery(oCmd.CommandText, Trans, oCmd.Parameters)  '20181009
        End With
        DbAccess.ExecuteNonQuery(uSql, tConn, uoCmd.Parameters)
    End Sub

    Sub UPDATE_STUDENTINFO_ISAGREE(ByVal aIDNO As String, ByVal sIsAgree As String, ByRef tConn As SqlConnection)
        sIsAgree = TIMS.ClearSQM(sIsAgree)
        aIDNO = TIMS.ChangeIDNO(aIDNO)
        If sIsAgree <> "Y" AndAlso sIsAgree <> "N" Then Return
        If aIDNO = "" Then Return

        Dim u_parms As New Hashtable From {{"IsAgree", sIsAgree}, {"IDNO", aIDNO}}
        Dim u_oStr As String = "UPDATE STUD_STUDENTINFO SET IsAgree=@IsAgree WHERE IDNO=@IDNO"
        Call DbAccess.ExecuteNonQuery(u_oStr, tConn, u_parms)
    End Sub

#End Region

    Dim save_entertemp2_LOCK As New Object
    Dim aNow As Date 'Dim aToday As Date
    Dim objconn As SqlConnection

    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload

        '金融機構代碼查詢 Financial institution code query
        HL_finaCodeQuery.NavigateUrl = TIMS.str_finaCodeQueryUrl
        HL_finaCodeQuery.Target = "_blank"
        HL_finaCodeQuery.ForeColor = Color.Blue

        aNow = TIMS.GetSysDateNow(objconn)
        Dim sql As String = ""

        Dim Y2016Start1 As String = "Y" 'TIMS.Utl_GetConfigSet("Y2016Start1") '測試2016年 啟用該功能。(Y/N)
        ServDept.Visible = True
        JobTitle.Visible = True
        ddlSERVDEPTID.Visible = False
        ddlJOBTITLEID.Visible = False
        If Y2016Start1 = "Y" Then
            ServDept.Visible = False
            JobTitle.Visible = False
            ddlSERVDEPTID.Visible = True
            ddlJOBTITLEID.Visible = True
            sql = " SELECT SERVDEPTID,SDNAME FROM dbo.KEY_SERVDEPT ORDER BY SERVDEPTID "
            dtSERVDEPT = DbAccess.GetDataTable(sql, objconn)
            sql = " SELECT JOBTITLEID,JTNAME FROM dbo.KEY_JOBTITLE ORDER BY JOBTITLEID "
            dtJOBTITLE = DbAccess.GetDataTable(sql, objconn)
        End If
        sql = " SELECT * FROM dbo.VIEW_ZIPNAME ORDER BY ZIPCODE "
        dtZipCode = DbAccess.GetDataTable(sql, objconn)
        sql = " SELECT * FROM dbo.KEY_GRADSTATE ORDER BY GRADID "
        dtGradState = DbAccess.GetDataTable(sql, objconn)
        'a_apost3.HRef=TIMS.cst_PostCodeQry3
        'a_hpost3.HRef=TIMS.cst_PostCodeQry3
        'a_cpost3.HRef=TIMS.cst_PostCodeQry3
        'a_apost2.HRef=TIMS.cst_PostCodeQry2
        'a_hpost2.HRef=TIMS.cst_PostCodeQry2
        'a_cpost2.HRef=TIMS.cst_PostCodeQry2

        '列出學歷下拉選單資料
        'objstr="SELECT * FROM Key_Degree where DegreeID IN ('01','02','03','04','05','06')"
        'sql="SELECT DEGREEID,NAME,DegreeType,SORT FROM Key_Degree where DegreeType=1 ORDER BY SORT"
        'dtDegree=DbAccess.GetDataTable(sql, objconn)

        '身分別
        ''列出主要參訓身分別下拉選單資料  
        ''20090922 andy  原本的03負擔家計婦女拿掉改成代碼28--獨力負擔家計者並加入 29天然災害受災民眾
        ''20090805 andy  edit 與tims學員資料維護選項一致加入-- 其他(就服法24條) 
        ''20100301 AMU 取消 02非自願離職者 
        ''20100310 AMU 取消 29天然災害受災民眾
        ''20100325 AMU 改回 其他(就服法24條)="更生保護人"(鍵值代碼10) 
        ''20121212 AMU 新增 37:六十五歲以上者資格 65歲以上 BY AMU 20121212
        'objstr &= " SELECT IdentityID,case when IdentityID='10' then '其他(就服法24條)' else Name end as Name" & vbCrLf
        'sql="" & vbCrLf
        'sql &= " SELECT IdentityID, Name" & vbCrLf
        'sql &= " FROM Key_Identity" & vbCrLf
        'sql &= " WHERE IdentityID IN  (" & TIMS.cst_Identity28 & ")" & vbCrLf
        'sql &= " ORDER BY IdentityID" & vbCrLf
        'dtIdentity=DbAccess.GetDataTable(sql, objconn)

        Dim sAltMsg As String = "" '訊息
        Dim flag_stopEnter2 As Boolean = TIMS.StopEnterTempMsg2(objconn, sAltMsg)
        If flag_stopEnter2 Then
            Common.MessageBox(Me, sAltMsg)
            Exit Sub
        End If

        'BackTable.Style("display")=""
        btnSend1.Attributes.Add("OnClick", "if(confirm ('資料確認無誤，確定送出？')) { return ChkData(); }")
        CheckBox1.Attributes("onclick") = "Clear_Zip2()"

        Dim xErrmsg As String = ""

        If Not Page.IsPostBack Then
            If TIMS.StopEnterTempMsg1(Me, objconn, True) Then Exit Sub

            ''Dim xErrmsg As String=""
            ''檢測是否停止報名
            ''SE:'產投停止報名
            'Const cst_StopFlag As String="SE"
            'Dim sAltMsg As String="" '訊息
            'Dim AltMsgSDate As String="" '訊息公佈日
            'Dim AltMsgEDate As String="" '訊息結束日
            'sAltMsg=TIMS.Get_System_Msg("AltMsg", cst_StopFlag, objconn)
            'AltMsgSDate=TIMS.Get_System_Msg("AltMsgSDate", cst_StopFlag, objconn)
            'AltMsgEDate=TIMS.Get_System_Msg("AltMsgEDate", cst_StopFlag, objconn)
            'xErrmsg=TIMS.Get_AltMsg_System_Msg(sAltMsg, AltMsgSDate, AltMsgEDate, aNow)
            'If xErrmsg <> "" Then
            '    '因網路系統維護，將於2012年8月14日12:10至13:10中斷服務1小時，造成不便，敬請見諒！
            '    Common.AddClientScript(Page, "alert('" & xErrmsg & "');location.href='../../main2.aspx';")
            '    Exit Sub
            'End If

            Call CCreate1()

            If Session("_SearchStr") IsNot Nothing Then
                ViewState("_SearchStr") = Session("_SearchStr")
                Session("_SearchStr") = Nothing
                RIDValue.Value = TIMS.GetMyValue(ViewState("_SearchStr"), "RIDValue")
                OCIDValue1.Value = TIMS.GetMyValue(ViewState("_SearchStr"), "OCIDValue1")
                'STDate.Value=TIMS.GetMyValue( ViewState("_SearchStr"), "STDate")
                TMIDValue1.Value = TIMS.GetMyValue(ViewState("_SearchStr"), "TMIDValue1")
                IDNO.Text = TIMS.GetMyValue(ViewState("_SearchStr"), "IDNO")
                Birthday.Text = TIMS.GetMyValue(ViewState("_SearchStr"), "birthDay")
                'RelEnterDate.Text=TIMS.GetMyValue( ViewState("_SearchStr"), "EnterDate")
            End If

            If OCIDValue1.Value = "" Then
                If ViewState("_SearchStr") IsNot Nothing Then
                    Session("_SearchStr") = ViewState("_SearchStr")
                    ViewState("_SearchStr") = Nothing
                End If
                'Response.Redirect("SD_01_010.aspx?ID=" & Request("ID") & "")
                Dim url1 As String = "SD_01_010.aspx?ID=" & Request("ID") & ""
                Call TIMS.Utl_Redirect(Me, objconn, url1)
            End If

            '代入初始資料
            Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
            ClassName.Text = If(InStr(drCC("ClassCName"), "班") = 0, drCC("ClassCName") & "  班", drCC("ClassCName"))

            'LOCIDdate.Text=FormatDateTime(Convert.ToString(sqldr("FTDate")), DateFormat.ShortDate) '本班的結訓日期 (最後使用經費日期)
            Hid_LOCIDdate.Value = Common.FormatDate(Convert.ToString(drCC("STDate"))) '本班的開訓日期 (最後使用經費日期)

            '函算政府補助經費--身分證號，依據此報名班別的結訓日期(開訓日期)
            IDNO.Text = TIMS.ChangeIDNO(IDNO.Text)
            Hid_LOCIDdate.Value = TIMS.ClearSQM(Hid_LOCIDdate.Value)

            Dim dr1 As DataRow = Nothing
            Dim sStr As String = $" SELECT dbo.FN_GET_GOVCOST('{IDNO.Text}','{Hid_LOCIDdate.Value}') GovCost "
            Dim objtable As DataTable = DbAccess.GetDataTable(sStr, objconn)
            If TIMS.dtHaveDATA(objtable) Then
                dr1 = objtable.Rows(0)
                GovCost.Text = "   您已使用政府補助經費 " & dr1("GovCost") & "元 "
            End If

            IDNO.Text = TIMS.ClearSQM(IDNO.Text)
            Birthday.Text = TIMS.ClearSQM(Birthday.Text)
            objtable = GET_E_MEMBER(IDNO.Text, Birthday.Text)
            If TIMS.dtHaveDATA(objtable) Then
                '採用 E_Member資料(基本資料)
                Call SHOW_E_MEMBER(objtable)
                '使用Stud_EnterTemp3 報名資料維護。
                Call SHOW_STUD_ENTERTEMP3()
                '有些特別碼的轉換
                Call CHG_HTMLCODE()
            End If
            '取消帶入報名詳細資料，避免委訓單位黑箱作業(宥妤) BY AMU 20120906
            'Call SHOW_Stud_EnterTemp12(IDNO.Text)
            If Mid(IDNO.Text, 2, 1) = "1" Then Sex.Items(0).Selected = True Else Sex.Items(1).Selected = True
        End If

        '判斷身分證號與生日組合 是否異常於 原系統資訊
        If xErrmsg = "" Then xErrmsg = TIMS.ChkDataIdnoBirth(objconn, IDNO.Text, Birthday.Text)
        '最後警告
        If xErrmsg <> "" Then Common.MessageBox(Me, xErrmsg)

    End Sub

    '有些特別碼的轉換
    Sub CHG_HTMLCODE()
        If Name.Text <> "" Then Name.Text = HttpUtility.HtmlDecode(Name.Text)
        If School.Text <> "" Then School.Text = HttpUtility.HtmlDecode(School.Text)
        If Department.Text <> "" Then Department.Text = HttpUtility.HtmlDecode(Department.Text)
        If Address.Text <> "" Then Address.Text = HttpUtility.HtmlDecode(Address.Text)
        If HouseholdAddress.Text <> "" Then HouseholdAddress.Text = HttpUtility.HtmlDecode(HouseholdAddress.Text)
    End Sub

    Sub CCreate1()
        If OCIDValue1.Value <> "" Then
            If TIMS.Get_OrgKind2(OCIDValue1.Value, TIMS.c_OCID, objconn) = "G" Then AcctMode.Items.RemoveAt(2)
        End If

        '郵遞區號查詢
        LitZip1.Text = TIMS.Get_WorkZIPB3Link2()
        LitZip2.Text = TIMS.Get_WorkZIPB3Link2()
        LitZip3.Text = TIMS.Get_WorkZIPB3Link2()

        Dim bt1_Attr_VAL As String = TIMS.GET_ZipCodeWOAspxUrl(ZipCode1, ZipCode1_B3, hidZipCode1_6W, City1, Address)
        Bt1_city_zip.Attributes.Add("onclick", bt1_Attr_VAL)
        Dim bt2_Attr_VAL As String = TIMS.GET_ZipCodeWOAspxUrl(ZipCode2, ZipCode2_B3, hidZipCode2_6W, City2, HouseholdAddress)
        Button1.Attributes.Add("onclick", bt2_Attr_VAL)
        Dim bt3_Attr_VAL As String = TIMS.GET_ZipCodeWOAspxUrl(ZipCode3, ZipCode3_B3, hidZipCode3_6W, City3, ActAddress)
        Button2.Attributes.Add("onclick", bt3_Attr_VAL)

        '列出學歷下拉選單資料
        DegreeID = TIMS.Get_Degree(DegreeID, 1, objconn)

        ''列出畢業狀況下拉選單資料
        'GraduateStatus=TIMS.Get_GradState(GraduateStatus)

        ''列出兵役下拉選單資料
        'MilitaryID=TIMS.Get_Military(MilitaryID) 'MilitaryID.Items.Remove(MilitaryID.Items.FindByValue("00"))

        'Get_Identity(4): ('01','04','05','06','07','26','10','28')
        'MIdentityID=TIMS.Get_Identity(MIdentityID, 4, objconn)
        'MIdentityID=TIMS.Get_Identity(MIdentityID, 5, sm.UserInfo.TPlanID, sm.UserInfo.Years, objconn)
        MIdentityID = TIMS.Get_Identity(MIdentityID, 52, sm.UserInfo.TPlanID, sm.UserInfo.Years, objconn)
        'MIdentityID.Items.Add(New ListItem("非自願離職者", "02"))
        ddlSERVDEPTID = TIMS.Get_SERVDEPTID(ddlSERVDEPTID, dtSERVDEPT)
        ddlJOBTITLEID = TIMS.Get_JOBTITLEID(ddlJOBTITLEID, dtJOBTITLE)

        '服務單位行業別
        Q4 = TIMS.Get_Trade(Q4)

        '畢業狀況
        With GraduateStatus
            .DataSource = dtGradState
            .DataTextField = "Name"
            .DataValueField = "GRADID"
            .DataBind()
        End With
    End Sub

    '回報名作業
    Private Sub BtnBack1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBack1.Click
        If Not ViewState("_SearchStr") Is Nothing Then
            Session("_SearchStr") = ViewState("_SearchStr")
            ViewState("_SearchStr") = Nothing
        End If
        'Response.Redirect("SD_01_010.aspx?ID=" & Request("ID") & "")
        Dim url1 As String = "SD_01_010.aspx?ID=" & Request("ID") & ""
        Call TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    ''' <summary>檢核報名狀況</summary>
    ''' <returns></returns>
    Function CheckData2() As String
        Dim rst As String = ""
        IDNO.Text = TIMS.ClearSQM(IDNO.Text)
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)

        Dim parms As New Hashtable From {{"IDNO", IDNO.Text}, {"OCID1", TIMS.CINT1(OCIDValue1.Value)}}
        '20090325(Milor)取消生日的判斷，只要身分證號重複就擋掉。
        Dim objstr As String = ""
        objstr &= " SELECT 'x'" & vbCrLf
        objstr &= " FROM dbo.STUD_ENTERTEMP2 se1" & vbCrLf
        objstr &= " JOIN dbo.STUD_ENTERTYPE2 se2 ON se1.eSETID=se2.eSETID" & vbCrLf
        objstr &= " WHERE se1.IDNO=@IDNO AND se2.OCID1=@OCID1" & vbCrLf
        Dim dt1 As DataTable = DbAccess.GetDataTable(objstr, objconn, parms)
        If TIMS.dtHaveDATA(dt1) Then rst &= "您已經有報名此班級了!!" & vbCrLf
        Return rst
    End Function

    ''' <summary> 檢核資料正確性 </summary>
    ''' <returns></returns>
    Function CheckEnterData() As String
        Dim rst As String = ""

        Name.Text = TIMS.ClearSQM(Name.Text)
        School.Text = TIMS.ClearSQM(School.Text)
        Department.Text = TIMS.ClearSQM(Department.Text)
        Address.Text = TIMS.ClearSQM(Address.Text)
        HouseholdAddress.Text = TIMS.ClearSQM(HouseholdAddress.Text)

        Dim aIDNO As String = ""
        Dim aPassPortNO As String = ""

        If Name.Text = "" Then rst &= "請填入中文姓名" & vbCrLf

        aPassPortNO = PassPortNO.SelectedValue
        Select Case PassPortNO.SelectedValue
            Case "1", "2"
            Case Else
                rst &= "請選擇 身分別" & vbCrLf
        End Select

        '身分證驗証
        IDNO.Text = TIMS.ChangeIDNO(IDNO.Text)
        aIDNO = TIMS.ChangeIDNO(IDNO.Text)
        If aIDNO = "" Then
            rst &= "必須填寫身分證號碼<BR>"
        Else
            Select Case aPassPortNO
                Case "2" '身分別為外籍 

                    'If rst="" Then
                    '    '=========== 驗証匯入檔案時不要有相同的身分證號碼 Start=============
                    '    Dim Flag As Boolean=True
                    '    For i As Integer=0 To IDNOArray.Count - 1
                    '        If IDNOArray(i)=aIDNO Then
                    '            Reason &= "檔案中有相同的身分證號碼<BR>"
                    '            Flag=False
                    '        End If
                    '    Next
                    '    If Flag Then IDNOArray.Add(aIDNO)
                    '    '=========== 驗証匯入檔案時不要有相同的身分證號碼 -End-=============
                    'End If

                Case "1"
                    If TIMS.CheckIDNO(aIDNO) Then '一般驗証
                        'If sm.UserInfo.RoleID=1 Then '角色代碼為1 可執行安全性規則確認
                        Dim IDNOFlag As Boolean = True
                        Dim EngStr As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
                        If aIDNO.Length <> 10 Then
                            IDNOFlag = False
                        ElseIf aIDNO.Chars(1) <> "1" And aIDNO.Chars(1) <> "2" Then
                            IDNOFlag = False
                        ElseIf EngStr.IndexOf(aIDNO.ToUpper.Chars(0)) = -1 Then
                            IDNOFlag = False
                        ElseIf aIDNO = "A123456789" Then
                            IDNOFlag = False
                        End If
                        If Not IDNOFlag Then rst &= "身分證號碼錯誤!" & vbCrLf

                        'If rst="" Then
                        '    '=========== 驗証匯入檔案時不要有相同的身分證號碼 Start=============
                        '    Dim Flag As Boolean=True
                        '    For i As Integer=0 To IDNOArray.Count - 1
                        '        If IDNOArray(i)=aIDNO Then
                        '            Reason &= "檔案中有相同的身分證號碼<BR>"
                        '            Flag=False
                        '        End If
                        '    Next
                        '    If Flag Then IDNOArray.Add(aIDNO)
                        '    '=========== 驗証匯入檔案時不要有相同的身分證號碼 -End-=============
                        'End If

                    Else
                        rst &= "身分證號碼錯誤!(如果有此身分證號碼，請聯絡系統管理者協助)!!!" & vbCrLf
                    End If
                Case Else
                    rst &= "請選擇身分別" & vbCrLf
            End Select
        End If

        'OK_IDNO()
        Select Case Sex.SelectedValue
            Case "F", "M"
            Case Else
                rst &= "請選擇性別" & vbCrLf
        End Select

        'If rst="" Then
        '    If Not TIMS.checkMemberSex(aIDNO, Sex.SelectedValue) Then rst &= "依身分證號判斷 性別選項 不正確，請確認" & vbCrLf
        'End If

        'C_Year()
        If Trim(Birthday.Text) <> "" Then
            Birthday.Text = Trim(Birthday.Text)
            If TIMS.IsDate1(Birthday.Text) Then
                Birthday.Text = CDate(Birthday.Text).ToString("yyyy/MM/dd")
                'Common.FormatDate(birthDay.Text)
            Else
                rst &= "出生日期格式有誤!!" & vbCrLf
            End If
        Else
            rst &= "請填入出生日期" & vbCrLf
        End If

        ActNo.Text = TIMS.ClearSQM(ActNo.Text)
        If ActNo.Text <> "" Then
            'ActNo.Text=ActNo.Text.Trim
            '投保單位保險證號為09開頭者，為訓字保，亦不可報名
            '根據參訓學員於e網所填列之保險證號前2碼判讀, 前2碼為
            '01、04、05、15及08者其補助經費來源歸屬為 03:就保基金
            '02、03、06、07者其經費來源歸屬為 02:就安基金
            '09與無法辨視者為 99:不予補助對象
            Select Case Left(ActNo.Text, 2)
                Case "09"
                    rst &= "學員資格 投保單位保險證號 為09開頭者為訓字保 不符合可參訓條件！" & vbCrLf
            End Select
        End If

        If rst = "" Then
            '檢測此學員是否 可參訓 產業人才投資方案 (大於15歲者)
            If Not TIMS.Check_YearsOld15(Birthday.Text, Hid_LOCIDdate.Value) Then rst &= "學員資格 年齡不滿15歲 不符合可參訓條件！" & vbCrLf
        End If

        If DegreeID.SelectedValue = "" Then rst &= "請選擇最高學歷" & vbCrLf
        If GraduateStatus.SelectedValue = "" Then rst &= "請選擇畢業狀況" & vbCrLf  '畢業狀況 GraduateStatus GradID

        '學校名稱 School School
        If School.Text <> "" Then School.Text = TIMS.ClearSQM(School.Text)
        If School.Text <> "" Then School.Text = HttpUtility.HtmlDecode(School.Text)
        If School.Text = "" Then rst &= "請輸入學校名稱" & vbCrLf

        '科系名稱 Department Department
        If Department.Text <> "" Then Department.Text = TIMS.ClearSQM(Department.Text)
        If Department.Text <> "" Then Department.Text = HttpUtility.HtmlDecode(Department.Text)
        If Department.Text = "" Then rst &= "請輸入科系名稱" & vbCrLf

        CellPhone.Text = TIMS.ClearSQM(CellPhone.Text) 'Else CellPhone.Text=""
        PhoneD.Text = TIMS.ClearSQM(PhoneD.Text) 'Else PhoneD.Text=""
        PhoneN.Text = TIMS.ClearSQM(PhoneN.Text) 'Else PhoneN.Text=""

        'If PhoneD.Text="" Then rst &= "請填入聯絡電話(日)" & vbCrLf
        'If PhoneN.Text="" Then rst &= "請填入聯絡電話(夜)" & vbCrLf

        Select Case rblMobil.SelectedValue
            Case "Y"
                'If CellPhone.Text="" Then rst &= "有行動電話 請輸入行動電話" & vbCrLf
                If Not TIMS.CheckPhone(CellPhone.Text) Then rst &= "有行動電話 請輸入行動電話" & vbCrLf
            Case Else
                If PhoneD.Text = "" AndAlso PhoneN.Text = "" Then rst &= "請填入聯絡電話(日)或電話(夜)" & vbCrLf
                If CellPhone.Text <> "" Then rst &= "有輸入行動電話,請選擇[有行動電話]" & vbCrLf
        End Select

        Address.Text = TIMS.ClearSQM(Address.Text)
        HouseholdAddress.Text = TIMS.ClearSQM(HouseholdAddress.Text)

        ZipCode1.Value = TIMS.ClearSQM(ZipCode1.Value)
        'If ZipCode1.Value.Trim <> "" Then ZipCode1.Value=ZipCode1.Value.Trim
        If ZipCode1.Value = "" Then
            rst &= "請選擇通訊地址(縣市) 郵遞區號" & vbCrLf
        ElseIf TIMS.IsNumberStr(ZipCode1.Value) Then
            ff = "ZipCode='" & ZipCode1.Value & "'"
            If dtZipCode.Select(ff).Length = 0 Then rst &= "請選擇正確的通訊地址(縣市) 郵遞區號!" & vbCrLf
        Else
            rst &= "請選擇正確的通訊地址(縣市) 郵遞區號!!" & vbCrLf
        End If

        ZipCode1_B3.Value = TIMS.ClearSQM(ZipCode1_B3.Value)
        If ZipCode1_B3.Value = "" Then
            rst &= "請輸入通訊地址(縣市) 郵遞區號後2碼或後3碼" & vbCrLf
        Else
            hidZipCode1_6W.Value = TIMS.GetZIPCODE6W(ZipCode1.Value, ZipCode1_B3.Value)
            Call TIMS.CheckZipCODEB3(ZipCode1_B3.Value, "通訊地址 郵遞區號後2碼或後3碼", False, rst)
        End If
        If Address.Text = "" Then
            rst &= "請輸入通訊地址 資料" & vbCrLf
        End If

        ZipCode2.Value = TIMS.ClearSQM(ZipCode2.Value)
        ZipCode2_B3.Value = TIMS.ClearSQM(ZipCode2_B3.Value)
        'If ZipCode2.Value.Trim <> "" Then ZipCode2.Value=ZipCode2.Value.Trim
        If CheckBox1.Checked Then
            If rst = "" Then
                ZipCode2.Value = ZipCode1.Value
                ZipCode2_B3.Value = ZipCode1_B3.Value
                HouseholdAddress.Text = Address.Text
            End If
        Else
            If ZipCode2.Value = "" Then
                rst &= "請選擇戶籍地址(縣市) 郵遞區號" & vbCrLf
            ElseIf TIMS.IsNumberStr(ZipCode2.Value) Then
                ff = "ZipCode='" & ZipCode2.Value & "'"
                If dtZipCode.Select(ff).Length = 0 Then rst &= "請選擇正確的戶籍地址(縣市) 郵遞區號!" & vbCrLf
            Else
                rst &= "請選擇正確的戶籍地址(縣市) 郵遞區號!!" & vbCrLf
            End If
            If ZipCode2_B3.Value = "" Then
                rst &= "請輸入戶籍地址(縣市) 郵遞區號後2碼或後3碼" & vbCrLf
            Else
                hidZipCode2_6W.Value = TIMS.GetZIPCODE6W(ZipCode2.Value, ZipCode2_B3.Value)
                Call TIMS.CheckZipCODEB3(ZipCode2_B3.Value, "戶籍地址 郵遞區號後2碼或後3碼", False, rst)
            End If
            If HouseholdAddress.Text = "" Then rst &= "請輸入戶籍地址 資料" & vbCrLf
        End If

        If Trim(Email.Text) <> "" Then Email.Text = TIMS.ClearSQM(Email.Text) Else Email.Text = ""
        Email.Text = TIMS.ChangeEmail(Email.Text)
        Select Case rblEmail.SelectedValue
            Case "Y" '選有
                If Email.Text = "" OrElse Email.Text = "無" Then
                    rst &= "有電子郵件 請輸入電子郵件" & vbCrLf
                Else
                    If Not TIMS.CheckEmail(Email.Text) Then rst &= "請輸入有效電子郵件 格式有誤" & vbCrLf
                End If
            Case Else '選無
                If Not Email.Text = "無" Then '非無(其他、空白、輸入EMAIL)
                    If Email.Text <> "" Then rst &= "有輸入電子郵件,請選擇有電子郵件" & vbCrLf
                End If
        End Select

        If MIdentityID.SelectedValue = "" Then rst &= "請選擇 主要參訓身分別" & vbCrLf

        If rst = "" AndAlso MIdentityID.SelectedValue = "04" Then
            '檢測此學員是否 屬於中高齡資格 45歲~65歲
            If Not TIMS.Check_YearsOld45(Birthday.Text, Hid_LOCIDdate.Value) Then rst &= "學員資格 年齡非介於45歲~65歲之間 不符合中高齡資格！" & vbCrLf
        End If

        If rst = "" AndAlso MIdentityID.SelectedValue = "37" Then
            '檢測此學員是否 屬於六十五歲以上者資格 65歲以上 BY AMU 20121212
            If Not TIMS.Check_YearsOld65(Birthday.Text, Hid_LOCIDdate.Value) Then rst &= "學員資格 年齡非65歲以上 不符合 六十五歲以上者 資格！" & vbCrLf
        End If

        'PriorWorkPay.Text=TIMS.ClearSQM(PriorWorkPay.Text)
        'If PriorWorkPay.Text <> "" Then
        '    If Not IsNumeric(PriorWorkPay.Text) Then
        '        rst &= "受訓前薪資 請填寫整數數字" & vbCrLf
        '    Else
        '        Try
        '            If CDbl(CInt(PriorWorkPay.Text)) <> CDbl(PriorWorkPay.Text) Then rst &= "受訓前薪資 請填寫整數數字" & vbCrLf
        '        Catch ex As Exception
        '            Call TIMS.SendMailTestx(Me, "PriorWorkPay", PriorWorkPay.Text)
        '            rst &= "受訓前薪資 請填寫整數數字" & vbCrLf
        '        End Try
        '    End If
        'End If

        '改非必填 201008 by AMU 
        'If AcctMode.SelectedValue="" Then rst &= "請選擇 郵政/銀行帳號 種類" & vbCrLf

        PostNo_1.Text = TIMS.ClearSQM(PostNo_1.Text)
        AcctNo1_1.Text = TIMS.ClearSQM(AcctNo1_1.Text)
        'PostNo_1.Text=PostNo_1.Text.Trim'PostNo_2.Text=PostNo_2.Text.Trim'AcctNo1_1.Text=AcctNo1_1.Text.Trim'AcctNo1_2.Text=AcctNo1_2.Text.Trim
        BankName.Text = BankName.Text.Trim
        AcctheadNo.Text = AcctheadNo.Text.Trim
        ExBankName.Text = ExBankName.Text.Trim
        AcctExNo.Text = AcctExNo.Text.Trim
        AcctNo2.Text = AcctNo2.Text.Trim

        '改非必填 201008 by AMU 
        'Select Case AcctMode.SelectedValue
        '    Case "0"
        '        If PostNo_1.Text="" Then rst &= "請輸入 局號 1" & vbCrLf
        '        If PostNo_2.Text="" Then rst &= "請輸入 局號 2" & vbCrLf
        '        If AcctNo1_1.Text="" Then rst &= "請輸入 帳號 1" & vbCrLf
        '        If AcctNo1_2.Text="" Then rst &= "請輸入 帳號 2" & vbCrLf
        '    Case "1"
        '        If BankName.Text="" Then rst &= "請輸入 總行名稱" & vbCrLf
        '        If AcctHeadNo.Text="" Then rst &= "請輸入 總行代號" & vbCrLf
        '        If ExBankName.Text="" Then rst &= "請輸入 分行名稱" & vbCrLf
        '        If AcctExNo.Text="" Then rst &= "請輸入 分行代號" & vbCrLf
        '        If AcctNo2.Text="" Then rst &= "請輸入 帳號" & vbCrLf
        'End Select

        Uname.Text = TIMS.ClearSQM(Uname.Text)
        Intaxno.Text = TIMS.ClearSQM(Intaxno.Text)
        ServDept.Text = TIMS.ClearSQM(ServDept.Text)
        ActName.Text = TIMS.ClearSQM(ActName.Text)
        ActNo.Text = TIMS.ClearSQM(ActNo.Text)
        ActTel.Text = TIMS.ClearSQM(ActTel.Text)

        If Uname.Text = "" Then rst &= "請輸入服務單位" & vbCrLf
        If ActName.Text = "" Then rst &= "請輸入投保單位名稱" & vbCrLf

        Select Case ActType.SelectedValue
            Case "1", "2"
            Case Else
                rst &= "請選擇投保類別" & vbCrLf
        End Select

        If Trim(Intaxno.Text) <> "" Then
            If TIMS.LENB(Intaxno.Text) >= 10 Then rst &= "統一編號 長度超過系統範圍(10)" & vbCrLf
        End If
        If ServDept.Visible AndAlso ServDept.Text = "" Then rst &= "請輸入服務部門" & vbCrLf  '服務部門 ServDept 30 CHAR
        'If ActNo.Text <> "" Then ActNo.Text=Trim(ActNo.Text)
        ActNo.Text = TIMS.ClearSQM(ActNo.Text)
        If ActNo.Text <> "" Then
            If TIMS.LENB(ActNo.Text) >= 20 Then
                rst &= "投保單位保險證號 長度超過系統範圍(20)" & vbCrLf
            End If
        End If
        'If ActNo.Text="" Then
        '    rst &= "請輸入投保單位保險證號" & vbCrLf
        'End If

        If ActNo.Text <> "" Then
            '投保單位保險證號為09開頭者，為訓字保，亦不可報名
            '根據參訓學員於e網所填列之保險證號前2碼判讀, 前2碼為
            '01、04、05、15及08者其補助經費來源歸屬為 03:就保基金
            '02、03、06、07者其經費來源歸屬為 02:就安基金
            '09與無法辨視者為 99:不予補助對象
            Select Case Left(ActNo.Text, 2)
                Case "09"
                    rst &= "學員資格 投保單位保險證號 為09開頭者為訓字保 不符合可參訓條件！" & vbCrLf
            End Select
        End If

        If ActTel.Text <> "" Then ActTel.Text = Trim(ActTel.Text)
        'If ActTel.Text="" Then rst &= "請輸入投保單位電話" & vbCrLf
        If ZipCode3.Value <> "" Then ZipCode3.Value = Trim(ZipCode3.Value)
        If ZipCode3.Value = "" Then
            'rst &= "請選擇投保單位地址(縣市) 郵遞區號" & vbCrLf
        End If
        ZipCode3_B3.Value = TIMS.ClearSQM(ZipCode3_B3.Value)
        'If ZipCode3_B3.Value <> "" Then ZipCode3_B3.Value=Trim(ZipCode3_B3.Value)
        'If ZipCode3.Value.Trim <> "" Then ZipCode3.Value=ZipCode3.Value.Trim 
        'If ZipCode3.Value="" Then
        '    rst &= "請選擇投保單位地址(縣市) 郵遞區號" & vbCrLf
        'Else
        '    If IsNumeric(ZipCode3.Value) Then
        '        ff="ZipCode='" & ZipCode3.Value & "'"
        '        If dtZipCode.Select(ff).Length=0 Then
        '            rst &= "請選擇正確的保單位地址(縣市) 郵遞區號" & vbCrLf
        '        End If
        '    Else
        '        rst &= "請選擇正確的保單位地址(縣市) 郵遞區號" & vbCrLf
        '    End If
        'End If

        If ZipCode3_B3.Value = "" Then
            'rst &= "請輸入投保單位地址(縣市) 郵遞區號後2碼" & vbCrLf
        Else
            ZipCode3_B3.Value = TIMS.ClearSQM(ZipCode3_B3.Value)
            hidZipCode3_6W.Value = TIMS.GetZIPCODE6W(ZipCode3.Value, ZipCode3_B3.Value)
            Call TIMS.CheckZipCODEB3(ZipCode3_B3.Value, "投保單位地址(縣市)  郵遞區號後2碼或後3碼", False, rst)
        End If
        If ActAddress.Text <> "" Then ActAddress.Text = Trim(ActAddress.Text)
        If ActAddress.Text = "" Then
            'rst &= "請輸入投保單位地址 資料" & vbCrLf
        End If

        Dim q2cnt As Integer = 0
        For i As Int16 = 0 To Q2.Items.Count - 1
            If Q2.Items(i).Selected Then q2cnt += 1
        Next
        If q2cnt = 0 Then rst &= "請選擇參訓動機" & vbCrLf

        Q61.Text = TIMS.ClearSQM(Q61.Text)
        Q62.Text = TIMS.ClearSQM(Q62.Text)
        Q63.Text = TIMS.ClearSQM(Q63.Text)
        Q64.Text = TIMS.ClearSQM(Q64.Text)
        'If Trim(Q61.Text) <> "" Then Q61.Text=Trim(Q61.Text) Else Q61.Text=""
        'If Trim(Q62.Text) <> "" Then Q62.Text=Trim(Q62.Text) Else Q62.Text=""
        'If Trim(Q63.Text) <> "" Then Q63.Text=Trim(Q63.Text) Else Q63.Text=""
        'If Trim(Q64.Text) <> "" Then Q64.Text=Trim(Q64.Text) Else Q64.Text=""
        Q61.Text = TIMS.CHECK_Q61TXTVAL("個人工作年資", Q61.Text, rst)
        Q62.Text = TIMS.CHECK_Q61TXTVAL("在這家公司的年資", Q62.Text, rst)
        Q63.Text = TIMS.CHECK_Q61TXTVAL("在這職位的年資", Q63.Text, rst)
        Q64.Text = TIMS.CHECK_Q61TXTVAL("最近升遷離本職幾年", Q64.Text, rst)

        'If IseMail.SelectedValue="" Then rst &= "請選擇是否希望 定期收到產業人才投資方案最新課程資訊。" & vbCrLf
        Dim v_IseMail As String = TIMS.GetListValue(IseMail)
        Dim v_rblEmail As String = TIMS.GetListValue(rblEmail)
        If v_IseMail = "" Then
            rst &= "(下方)請選擇是否希望 定期收到產業人才投資方案最新課程資訊。" & vbCrLf
        Else
            If v_rblEmail <> "Y" AndAlso v_IseMail = "Y" Then rst &= "希望收到最新課程資訊; 電子郵件 請選擇「有」並填寫有效資料。" & vbCrLf
        End If

        'If IsAgree.SelectedValue="" Then Common.SetListItem(IsAgree, "Y")
        Common.SetListItem(IsAgree, "Y")
        Common.SetListItem(IsAgreedata, "Y")
        'If IsAgreedata.SelectedValue="Y" Then rst &= "由於您不同意您的資料用於上開所列蒐集目的，因此無法報名!" & vbCrLf
        Common.SetListItem(IsCorrect, "Y")
        'If IsCorrect.SelectedValue <> "Y" Then rst &= "請填寫正確且最新的個人資料" & vbCrLf
        'If IsAgree.SelectedValue="" Then rst &= "請選擇是否同意 提供職業訓練及就業服務時使用本人資料。" & vbCrLf

        'rst &= "測試錯誤 1" & vbCrLf
        'rst &= "測試錯誤 2" & vbCrLf
        Return rst
    End Function

    ''' <summary>送出</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnSend1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSend1.Click
        ' btnSend1.Attributes.Add("OnClick", "if(confirm ('資料確認無誤，確定送出？')) { return ChkData(); }")
        Dim Errmsg As String = ""
        If TIMS.StopEnterTempMsg1(Me, objconn, True) Then Exit Sub

        Dim sAltMsg As String = "" '訊息
        Dim flag_stopEnter2 As Boolean = TIMS.StopEnterTempMsg2(objconn, sAltMsg)
        If flag_stopEnter2 Then
            Common.MessageBox(Me, sAltMsg)
            Exit Sub
        End If

        '報名資料再確認
        Dim xErrmsg As String = ""
        Dim drCC As DataRow = TIMS.GetOCIDDate2854(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            xErrmsg &= " 報名班級資料有誤，請重新查詢！" & vbCrLf
            Common.MessageBox(Me, xErrmsg)
            Exit Sub
        End If

        Dim flag_Chk_OnShellDate As Boolean = False
        If TIMS.Cst_TPlanID28.IndexOf(Convert.ToString(drCC("TPlanID"))) > -1 Then flag_Chk_OnShellDate = True
        If flag_Chk_OnShellDate Then
            '上架日期-ONSHELLDATE
            If Convert.ToString(drCC("ONSHELLDATE")) = "" Then
                xErrmsg &= " 此班級尚未開始報名!!!" & vbCrLf
                Common.MessageBox(Me, xErrmsg)
                Exit Sub
            End If
            '上架日期-ONSHELLDATE
            Dim ChkTime3 As Long = 0
            ChkTime3 = DateDiff(DateInterval.Minute, CDate(aNow), CDate(drCC("ONSHELLDATE"))) '未到結束報名時間大於0
            If ChkTime3 > 0 Then
                xErrmsg &= " 此班級尚未開始報名!!!" & vbCrLf
                Common.MessageBox(Me, xErrmsg)
                Exit Sub
            End If
        End If

        ViewState("SEnterDate") = TIMS.CssFormatDate(drCC("SEnterDate"))
        ViewState("FEnterDate") = TIMS.CssFormatDate(drCC("FEnterDate"))
        Dim ChkTime1 As Long = 0
        Dim ChkTime2 As Long = 0
        ChkTime1 = DateDiff(DateInterval.Second, CDate(drCC("SEnterDate")), CDate(aNow))  '過報名時間大於0
        ChkTime2 = DateDiff(DateInterval.Minute, CDate(aNow), CDate(drCC("FEnterDate"))) '未到結束報名時間大於0

        'If TestStr="AmuTest" Then '測試'    ChkTime1=1'    ChkTime2=1'End If '測試
        '為配合2016年度課程公告作業，擬將2016年上半年度核可並轉班完成之課程統一於2016年1月23日0:01起方才能於產投報名網站上查詢到公告課程。
        If ChkTime1 >= 0 AndAlso ChkTime2 >= 0 Then
            Dim bln_noEnter As Boolean = False
            If DateDiff(DateInterval.Second, aNow, CDate(TIMS.cst_SEnterDate2016_28)) >= 0 Then bln_noEnter = True
            'If Convert.ToString(drCC("Years"))="2016" AndAlso bln_noEnter Then
            '    ChkTime1=-1
            '    ChkTime2=-1
            '     ViewState("SEnterDate")=TIMS.CssFormatDate(CDate(TIMS.cst_SEnterDate2016_28))
            'End If
        End If

        '在報名時間內
        If ChkTime1 > 0 AndAlso ChkTime2 >= 0 Then
            'vsOCIDvalue 'Session(tims.cst_OCID)=Trim(Li_Class.Text)  '記錄報名課程資料 '在報名時間內
        ElseIf ChkTime1 < 0 Then
            '此班級將於(該班可報名時間)開始報名!! 'Common.MessageBox(Me,  ViewState("SEnterDate") & " 此班級尚未開始報名!!!")
            xErrmsg &= " 此班級將於(" & ViewState("SEnterDate") & ")開始報名!!!" & vbCrLf
        Else
            '報名時間已過
            xErrmsg &= ViewState("FEnterDate") & " 此班報名時間已過!!!" & vbCrLf
        End If

        If xErrmsg <> "" Then
            ' ViewState("aNow")=TIMS.CssFormatDate(aNow)
            xErrmsg &= String.Concat(" ", vbCrLf, " 目前系統時間為(", TIMS.CssFormatDate(aNow), ")!!", vbCrLf)
            'Common.MessageBox(Me, xErrmsg) 'Exit Sub
            Dim flagS1 As Boolean = TIMS.IsSuperUser(Me, 1) '是否為(後台)系統管理者 
            Dim flagS2 As Boolean = TIMS.sUtl_ChkTest() '測試環境
            If flagS1 OrElse flagS2 Then
                xErrmsg &= cst_msgSuper1 & vbCrLf
                Common.MessageBox(Me, xErrmsg)
            Else
                Common.MessageBox(Me, xErrmsg)
                Exit Sub
            End If
        End If

        '檢核資料正確性
        Try
            Errmsg = CheckEnterData()
        Catch ex As Exception
            TIMS.LOG.Error(ex.Message, ex)
            Errmsg = ex.Message

            Dim cst_fun_page_name As String = "##SD_01_010_add.aspx, "
            Dim slogMsg1 As String = ""
            slogMsg1 &= cst_fun_page_name & "ff: " & ff & vbCrLf
            'Call TIMS.SendMailTest(slogMsg1)
            Dim strErrmsg As String = ""
            strErrmsg &= "ex.Message:" & vbCrLf & ex.Message & vbCrLf
            strErrmsg &= "ex.ToString:" & vbCrLf & ex.ToString & vbCrLf
            'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.SendMailTest(strErrmsg)
        End Try
        If Errmsg <> "" Then
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If
        Errmsg = CheckData2()
        If Errmsg <> "" Then
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If

        'WebRequest物件如何忽略憑證問題
        System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf TIMS.ValidateServerCertificate)
        'TLS 1.2-基礎連接已關閉: 傳送時發生未預期的錯誤 
        System.Net.ServicePointManager.SecurityProtocol = Net.SecurityProtocolType.Tls12
        '檢核學員重複參訓。
        'http://163.29.199.211/TIMSWS/timsService1.asmx 'https://wltims.wda.gov.tw/TIMSWS/timsService1.asmx

        Dim timsSer1 As New timsService1.timsService1
        '檢核學員重複參訓。
        Dim aIDNO1 As String = TIMS.ClearSQM(IDNO.Text)
        Dim aOCID1 As String = TIMS.ClearSQM(drCC("OCID"))
        Dim xStudInfo As String = ""
        TIMS.SetMyValue(xStudInfo, "IDNO", aIDNO1)
        TIMS.SetMyValue(xStudInfo, "OCID1", aOCID1)
        'TIMS.SetMyValue(xStudInfo, "STEST", "Y")
        '檢核學員OOO有參訓時段重疊報名情形，無法儲存，請再確認！
        Call TIMS.ChkStudDouble(timsSer1, Errmsg, "", xStudInfo)
        If Errmsg <> "" Then
            Common.MessageBox(Me, Errmsg)
            Return 'Exit Sub
        End If

        'Request("ID") 'Errmsg=""
        Dim s_Redirect1 As String = String.Concat("SD_01_010.aspx?ID=", TIMS.Get_MRqID(Me)) '課程代號輸入頁面
        Dim s_AlertMsg As String = ""
        '儲存 (報名)
        Dim flag_ok_save As Boolean = False
        SyncLock save_entertemp2_LOCK
            Try
                flag_ok_save = SAVE_STUD_ENTERTEMP2(IDNO.Text, OCIDValue1.Value, Errmsg, s_Redirect1, s_AlertMsg)
            Catch ex As Exception
                TIMS.LOG.Error(ex.Message, ex)
                Errmsg = "報名失敗，學員報名異常 或 資料庫異常，請重試!!" & vbCrLf
                Errmsg &= "請再試一次，造成您不便之處，還請見諒。" & vbCrLf
                Errmsg &= String.Concat("(若持續出現此問題，請聯絡系統管理者)!!", vbCrLf, ex.Message)
                'Return False
                Common.MessageBox(Me, Errmsg) '需要顯示狀況
                Return
            End Try
        End SyncLock

        '儲存失敗
        If Not flag_ok_save Then
            'Common.showMsg(Me, Errmsg, Request.FilePath)
            Common.MessageBox(Me, Errmsg)
            Return 'Exit Sub
        End If

        '報名成功    
        Dim aIDNO As String = IDNO.Text
        Dim sIsAgree As String = "Y" '預設為同意。
        'If  IsAgree.Items(0).Selected=False Then sIsAgree="N" '不同意
        Call UPDATE_STUDENTINFO_ISAGREE(aIDNO, sIsAgree, objconn)
        If hid_eSETID3.Value <> "" AndAlso TIMS.IsNumeric1(hid_eSETID3.Value) AndAlso TIMS.CINT1(hid_eSETID3.Value) > 0 Then
            Call UPDATE_STUD_ENTERTEMP3(hid_eSETID3.Value, aIDNO, objconn)
        End If
        'check.Value="Y" 'Common.MessageBox(Me, Errmsg) '儲存成功
        Common.RespWrite(Me, s_AlertMsg) '有需要顯示狀況(可以報名)
    End Sub
End Class