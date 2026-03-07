Partial Class SD_01_001_add
    Inherits AuthBasePage

#Region "WEBFORM"
    Sub SUtl_PageInit1()
        Dim dt As DataTable = TIMS.Get_USERTABCOLUMNS("STUD_ENTERTEMP", objconn)
        Call TIMS.sUtl_SetMaxLen(dt, "IDNO", IDNO)
        Call TIMS.sUtl_SetMaxLen(dt, "NAME", Name)
        Call TIMS.sUtl_SetMaxLen(dt, "SCHOOL", School)
        Call TIMS.sUtl_SetMaxLen(dt, "DEPARTMENT", Department)
        Call TIMS.sUtl_SetMaxLen(dt, "ADDRESS", Address)
        Call TIMS.sUtl_SetMaxLen(dt, "PHONE1", Phone1)
        Call TIMS.sUtl_SetMaxLen(dt, "PHONE2", Phone2)
        Call TIMS.sUtl_SetMaxLen(dt, "CELLPHONE", CellPhone)
        Call TIMS.sUtl_SetMaxLen(dt, "EMAIL", Email)
    End Sub
#End Region

    Dim prtFilename As String = ""             '列印表件名稱
    Dim iPYNum17 As Integer = 1                'iPYNum17=TIMS.sUtl_GetPYNum17(Me)
    Dim flgROLEIDx0xLIDx0 As Boolean = False   '判斷登入者的權限。
    Dim ivsExistence As Integer = 0
    Dim ivsSETID As Integer = 0
    Dim gsEnterDate As String = ""             '系統報名日期(全域)

    Dim rqFrom_type As String = ""             'Request("from_type")
    Dim rqProecess As String = ""              'Request("proecess") 'add/edit/shift
    'Const cst_master1 As String="具公司/商業負責人身分，認定為在職者"                                          '(檢查SD_01_001.aspx)
    'Const cst_msgERR2 As String="查詢該民眾 具公司/商業負責人身分 連線有誤，請重新查詢!!"                         '(檢查SD_01_001.aspx)
    'Const cst_msgAlt2 As String="請先確認民眾是否非公司／商業負責人及非就保非自願離職者，須符合參訓資格才能繼續報名。"   '(檢查SD_01_001.aspx)
    Const cst_msgAlt4 As String = "投保證號為075、175（裁減續保）、076、176（職災續保）、09（訓）皆為不予補助對象，惠請查明該筆民眾身分是否符合本計畫參訓資格。"
    Dim dtZipCode As DataTable = Nothing
    'Dim dtGradState As DataTable=Nothing
    Dim dtSERVDEPT As DataTable = Nothing
    Dim dtJOBTITLE As DataTable = Nothing
    Dim blnTestflag As Boolean = False '測試中

    Dim aNow As Date
    Dim objconn As SqlConnection

    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        Call SUtl_PageInit1()
        aNow = TIMS.GetSysDateNow(objconn)

        iPYNum17 = TIMS.sUtl_GetPYNum17(Me)  '若是登入年度為 2017年以後，則傳回2，其餘為1
        'If blnTestflag Then iPYNum17=2
        blnTestflag = TIMS.sUtl_ChkTest()
        '是否為超級使用者
        flgROLEIDx0xLIDx0 = TIMS.IsSuperUser(Me, 1)

        CheckBox1.Attributes("onclick") = "Clear_Zip2()"

        '(職前邏輯)
        ''2017年後 '2017職前 使用勞保明細檢查鈕
        'BtnCheckBli.Visible=False
        'If iPYNum17=2 AndAlso TIMS.Cst_TPlanID_PreUseLimited17f.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '    HidPreUseLimited17f.Value=TIMS.cst_YES
        '    BtnCheckBli.Visible=False
        '    Button8.Visible=False
        '    lab2017_1.Text="" '"任職<br />單位名稱"
        '    lab2017_2.Text="" '"投保單位<br />加退保日期"
        '    lab2017_3.Text="" '"投保單位<br />保險證號"

        '    BtnCheckBli.Visible=True
        '    lab2017_1.Text="任職<br />單位名稱"
        '    lab2017_2.Text="投保單位<br />加退保日期"
        '    lab2017_3.Text="投保單位<br />保險證號"
        '    'Select Case iPYNum17
        '    '    Case 2 '2017年後
        '    '    Case 1 '2016年前
        '    'End Select
        'Else
        '    '2016年前
        '    Button8.Visible=True
        '    lab2017_1.Text="最後一次任<br />職單位名稱"
        '    lab2017_2.Text="最後投保單<br />位起迄日期"
        '    lab2017_3.Text="最後投保單<br />位保險證號"
        'End If

        Select Case Convert.ToString(Request("ID"))
            Case TIMS.cst_FunID_報名登錄 '報名登錄(SD_01_001_add)
            Case TIMS.cst_FunID_專案核定報名登錄 '專案核定報名登錄
            Case TIMS.cst_FunID_特例專案核定報名登錄 '專案核定報名登錄
            Case Else
                Common.MessageBox(Me, "基本輸入參數有誤，請重新選擇功能查詢!!", "SD_01_001.aspx?ID=" & TIMS.cst_FunID_報名登錄)
                Exit Sub
        End Select

        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub

        Page.RegisterStartupScript("ZipCcript1", TIMS.Get_ZipNameJScript(objconn))
        'TBCity.Attributes("onblur")="getzipname(this.value,'TBCity','city_code');"
        'conn=DbAccess.GetConnection
        EnterChannel.Attributes("onchange") = "EnterChannelChange();"
        '1.網;2.現;3.通;4.推

        '2017年後 
        trAVTCP.Visible = False '獲得職訓課程管道 cblAVTCP1
        If iPYNum17 = 2 Then trAVTCP.Visible = True

        If Not IsPostBack Then
            If TIMS.StopEnterTempMsg_ANM2(Me, objconn, True) Then Exit Sub
            rblWorkSuppIdent.Attributes("onclick") = "return CHK_WKSI();"

            '(職前邏輯)
            ''勞保及3合1就業資料查詢 '勞保及三合一就業資料查詢(MDate)
            'BtnCheckBli.Attributes("onclick")="open_SD01001sch();return false;"

            '2015年執行(未決定執行)。
            'If TIMS.Utl_GetConfigSet("work2015")="Y" Then
            '    Common.RespWrite(Me, "<script language=javascript>window.alert('請先確認民眾是否非就保非自願離職者，須非就保自願離職者才能繼續報名。')</script>")
            'End If
            'Common.RespWrite(Me, "<script language=javascript>window.alert('請先確認民眾是否非就保非自願離職者，須非就保自願離職者才能繼續報名。')</script>")

            '(職前邏輯)
            ''排除
            ''請先確認民眾是否非就保非自願離職者，須非就保自願離職者才能繼續報名
            'If TIMS.Cst_NotTPlanID5.IndexOf(sm.UserInfo.TPlanID)=-1 Then
            '    Dim sScript1 As String=""
            '    sScript1=""
            '    sScript1 &= "<script language=javascript>"
            '    sScript1 &= "window.alert('" & cst_msgAlt2 & "');"
            '    sScript1 &= "</script>"
            '    Common.RespWrite(Me, TIMS.sUtl_AntiXss(sScript1))
            'End If
            ''20090330專上畢業學歷失業者
            'Dim flagHighEduBg As Boolean=False
            'flagHighEduBg=TIMS.Check_OptOptions("專上畢業學歷失業者", Convert.ToString(sm.UserInfo.TPlanID), objconn)
            'HGTR.Visible=flagHighEduBg

            '是否為在職者補助身分
            TPlanid.Value = Convert.ToString(sm.UserInfo.TPlanID)
            WSITR.Visible = False
            '是否為在職者補助身分 46:補助辦理保母職業訓練'47:補助辦理照顧服務員職業訓練
            If TIMS.Cst_TPlanID46AppPlan5.IndexOf(sm.UserInfo.TPlanID) > -1 Then WSITR.Visible = True

            Dim rqTicket As String = TIMS.ClearSQM(Request("ticket"))
            If rqTicket <> "" Then Hid_rqTicket.Value = rqTicket
            Dim ticketType As String = TIMS.ClearSQM(Request("ticketType"))
            If ticketType <> "" Then Hid_ticketType.Value = ticketType
            Dim ticketNo As String = TIMS.ClearSQM(Request("TICKET_NO"))
            If ticketNo <> "" Then Hid_ticketNo.Value = ticketNo

            rqFrom_type = TIMS.ClearSQM(Request("from_type"))
            rqProecess = TIMS.ClearSQM(Request("proecess"))
            ptype.Value = rqProecess '2007/08/20 by mick 將shift 和 add 換此欄位判斷
            '--------------------------------------三合一修改規則：--------------------------------------
            '修改日：2007/08/19-2007/08/20 'by mick
            '新增時檢查此idno是否有三合一資料，並帶入學生三合一資料(班別不帶入)
            '志願一、二、三以志願一為主
            '如志願一為三合一班別時，直接走三合一流程 Request("proecess")="shift"
            '如志願二、三為三合一班別時，則提示使用者，如確定繼續則改走一般流程 Request("proecess")="add"
            '如志願一、二、三均不是三合一班別，則直接走一般流程
            '--------------------------------------三合一修改規則：--------------------------------------
            'Common.MessageBox(Me, "此學員含有推介單報名資料" & vbCrLf & "推介券編號@A31990A02200700017" & vbCrLf & "發卷日期:2007/3/20")
            'Exit Sub
            Hid_PreUseLimited18a.Value = ""
            If TIMS.Cst_TPlanID_PreUseLimited18a.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                Hid_PreUseLimited18a.Value = "Y"
            End If

            'DGTR.Visible=False
            'GovTR.Visible=False
            Call Add_Items() '增加下拉選單資料
            Call SU_Create1() '依傳入資料建立 (報名者資料檢核)
            Call SU_Create2() '依傳入資料 後續動作執行

            If Session("_SearchStr") IsNot Nothing Then
                ViewState("_SearchStr") = Session("_SearchStr")
                Session("_SearchStr") = Nothing
            End If
        End If

        IDNO.Attributes("onchange") += "javascript:chkidnosex();"
        IDNO.Attributes("onblur") += "javascript:chkidnosex();"

        '送出(隱藏)
        Button1.Style("display") = "none"
        '送出 //document.getElementById('Button1').click();
        Button7.Attributes("onclick") = "javascript:return chkdata();"
        'Button4 / 回報名登錄 / Button4_ServerClick

        '查詢歷史紀錄//onclick() SD_05_010.aspx
        'Button6.Attributes("onclick")="wopen('SD_01_001_old.aspx?IDNO='+document.form1.IDNO.value,'history',600,600,1);"
        'Button9.Attributes.Add("onclick", "return open_StudentList('" & IDNO.Text & "');")
        Dim rqID As String = TIMS.Get_MRqID(Me)
        Button9.Attributes.Add("onclick", $"return open_StudentList('{rqID}');")
        If IDNO.Text <> "" Then
            Dim s_ENCIDNO As String = RSA20031.AesEncrypt2(IDNO.Text)
            Hid_ENCIDNO.Value = s_ENCIDNO
        End If
        hdatenow.Value = FormatDateTime(aNow, DateFormat.ShortDate)
        'Sex.Attributes("onclick")="autoMilitary();"

        ''2017年後 '2017職前 使用勞保明細檢查鈕
        'If iPYNum17=2 AndAlso TIMS.Cst_TPlanID_PreUseLimited17f.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '    HidPreUseLimited17f.Value=TIMS.cst_YES
        '    PriorWorkType1.Attributes("onclick")="chgPriorWorkType1();"
        'End If

    End Sub

    '檢查 是否列入處分名單
    Sub Chk_OrgBlackList()
        Dim msg2 As String = ""
        'Me.ViewState("msg2")=""
        Dim strComIDNO As String = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, objconn)
        If TIMS.Check_OrgBlackList(Me, strComIDNO, objconn) Then
            'Me.ViewState("msg2")=sm.UserInfo.OrgName & "，已列入處分名單!!"
            msg2 = sm.UserInfo.OrgName & "，已列入處分名單!!"
            isBlack.Value = "Y"
            orgname.Value = sm.UserInfo.OrgName
        End If
    End Sub

    ''' <summary> 增加下拉選單資料 </summary>
    Sub Add_Items()
        '檢查 是否列入處分名單
        Call Chk_OrgBlackList()
        '20060517 by Vicient start
        DegreeID = TIMS.Get_Degree(DegreeID, 1, objconn)

        If trAVTCP.Visible Then cblAVTCP1 = TIMS.Get_AVTCP(cblAVTCP1, objconn)

        MilitaryID = TIMS.Get_Military(MilitaryID, 1, objconn)
        MilitaryID.Items.Remove(MilitaryID.Items.FindByValue("00"))

        'MIdentityID 'SELECT * FROM Key_Identity
        '參訓身分別鍵詞檔2010/08/12 改為用  Plan_Identity table 可依計畫設定不用顯示
        'MIdentityID=TIMS.Get_Identity(MIdentityID, 5, sm.UserInfo.TPlanID, sm.UserInfo.Years, objconn)
        'Identity=TIMS.Get_Identity(Identity, 5, sm.UserInfo.TPlanID, sm.UserInfo.Years, objconn)
        MIdentityID = TIMS.Get_Identity(MIdentityID, 52, sm.UserInfo.TPlanID, sm.UserInfo.Years, objconn)
        IdentityID = TIMS.Get_Identity(IdentityID, 53, sm.UserInfo.TPlanID, sm.UserInfo.Years, objconn)

        Dim rqSTDate As String = TIMS.ClearSQM(Request("STDate"))
        If rqSTDate <> "" AndAlso IdentityID.Items.FindByValue("03") IsNot Nothing Then
            If CDate(rqSTDate) > CDate("2009/06/01") Then IdentityID.Items.Remove(IdentityID.Items.FindByValue("03"))
        End If
        GradID = TIMS.Get_GradState(GradID, objconn)

        For Each li As ListItem In IdentityID.Items
            li.Attributes.Add("ChkValue", li.Text.Trim)
        Next li

        '郵遞區號查詢
        LitZip1.Text = TIMS.Get_WorkZIPB3Link2()
        LitZip2.Text = TIMS.Get_WorkZIPB3Link2()
        LitZip3.Text = TIMS.Get_WorkZIPB3Link2()

        Dim bt1_Attr_VAL As String = TIMS.GET_ZipCodeWOAspxUrl(city_code, ZipCODEB3, hidZipCODE6W, TBCity, Address)
        bt1_zipcode.Attributes.Add("onclick", bt1_Attr_VAL)
        Dim bt2_Attr_VAL As String = TIMS.GET_ZipCodeWOAspxUrl(ZipCode2, ZipCode2_B3, HidZipCode2_6W, City2, HouseholdAddress)
        Button6.Attributes.Add("onclick", bt2_Attr_VAL)
        Dim bt3_Attr_VAL As String = TIMS.GET_ZipCodeWOAspxUrl(ZipCode3, ZipCode3_B3, HidZipCode3_6W, City3, ActAddress)
        Button8.Attributes.Add("onclick", bt3_Attr_VAL)

    End Sub

    '採用檢視查詢資料
    Sub ItemVisibleState()
        Button1.Visible = False
        Button7.Visible = False
        Button5.Disabled = True
        'Button2.Disabled=True
        'Button3.Disabled=True
    End Sub

    '依傳入資料建立 (報名者資料檢核)
    Sub SU_Create1()
        aNow = TIMS.GetSysDateNow(objconn)
        Dim rqSerial As String = TIMS.ClearSQM(Request("serial"))

        Dim sql As String = ""
        sql = " SELECT * FROM dbo.VIEW_ZIPNAME ORDER BY ZIPCODE "
        dtZipCode = DbAccess.GetDataTable(sql, objconn)

        sql = " SELECT SERVDEPTID,SDNAME FROM dbo.KEY_SERVDEPT ORDER BY SERVDEPTID "
        dtSERVDEPT = DbAccess.GetDataTable(sql, objconn)
        ddlSERVDEPTID = TIMS.Get_SERVDEPTID(ddlSERVDEPTID, dtSERVDEPT)

        sql = " SELECT JOBTITLEID,JTNAME FROM dbo.KEY_JOBTITLE ORDER BY JOBTITLEID "
        dtJOBTITLE = DbAccess.GetDataTable(sql, objconn)
        ddlJOBTITLEID = TIMS.Get_JOBTITLEID(ddlJOBTITLEID, dtJOBTITLE)

        R_serial.Value = TIMS.ClearSQM(Request("serial"))
        R_EnterDate.Value = TIMS.Cdate3(TIMS.ClearSQM(Request("EnterDate")))
        R_SerNum.Value = TIMS.ClearSQM(Request("SerNum"))

        Dim rqIDNO As String = TIMS.ChangeIDNO(TIMS.ClearSQM(Request("IDNO")))
        IDNO.Text = If(rqIDNO <> "", rqIDNO, "")

        If rqProecess = "" Then rqProecess = TIMS.ClearSQM(Request("proecess"))
        Dim rqTRN_UNKEY As String = TIMS.ClearSQM(Request("TRN_UNKEY"))
        Dim rqBIRTH As String = TIMS.ClearSQM(Request("BIRTH"))
        Dim rqTICKET_NO As String = TIMS.ClearSQM(Request("TICKET_NO"))
        Dim rqAPPLY_DATE As String = TIMS.ClearSQM(Request("APPLY_DATE"))

        Select Case rqProecess 'Proeces add/edit/shift
            Case "add"
            Case "edit"
                If R_serial.Value = "" Then
                    Common.MessageBox(Me, "傳入參數有誤，請重新查詢。")
                    Exit Sub
                End If
                If R_EnterDate.Value = "" Then
                    Common.MessageBox(Me, "傳入參數有誤，請重新查詢。")
                    Exit Sub
                End If
                If R_SerNum.Value = "" Then
                    Common.MessageBox(Me, "傳入參數有誤，請重新查詢。")
                    Exit Sub
                End If
                If TIMS.Cst_TPlanID06Plan1.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    Dim strSS As String = BliClass1.Get_SSValue(R_serial.Value, R_EnterDate.Value, R_SerNum.Value, objconn)
                    If strSS <> "" Then BliClass1.Get_SELRESULTBNG(Me, strSS, objconn)
                End If
            Case "shift"
            Case Else
                '偵錯用儲存欄
                Dim strErrmsg As String = ""
                strErrmsg += "/* SD_01_001_add: */" & vbCrLf
                strErrmsg += "Request.RawUrl: " & Request.RawUrl & vbCrLf
                strErrmsg += "rqProecess:" & rqProecess & vbCrLf
                strErrmsg += "R_serial:" & R_serial.Value & vbCrLf
                strErrmsg += "R_EnterDate:" & R_EnterDate.Value & vbCrLf
                strErrmsg += "R_SerNum:" & R_SerNum.Value & vbCrLf
                'strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
                'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                'Call TIMS.SendMailTest(strErrmsg)
                Common.MessageBox(Me, "傳入參數有誤，請重新查詢。")
                Exit Sub
        End Select

        '限定計畫執行 '鎖定為在職者
        '原TIMS邏輯, 只有公司負責人身份才會跑這一段(職前課程邏輯), 
        '產投/在職 不用判斷是否為公司負責人
        If TIMS.Cst_TPlanID46AppPlan5.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Common.SetListItem(rblWorkSuppIdent, "Y")
            rblWorkSuppIdent.Enabled = False '鎖定為在職者
            'TIMS.Tooltip(rblWorkSuppIdent, cst_master1)
        End If

        ivsExistence = 0
        ivsSETID = 0
        Dim ss As String = ""
        TIMS.SetMyValue(ss, "rqSerial", rqSerial)
        TIMS.SetMyValue(ss, "rqIDNO", rqIDNO)
        TIMS.SetMyValue(ss, "rqBIRTH", rqBIRTH)
        TIMS.SetMyValue(ss, "rqTRN_UNKEY", rqTRN_UNKEY)
        TIMS.SetMyValue(ss, "rqAPPLY_DATE", rqAPPLY_DATE)
        TIMS.SetMyValue(ss, "rqTICKET_NO", rqTICKET_NO)

        Select Case rqProecess 'proecess add/edit/shift
            Case "add" '新增動作
                Call SU_Proecess_add(ss)
            Case "edit"
                Call SU_Proecess_edit(ss)
                'Case "shift" 
                '    Call sU_Proecess_shift(ss)
        End Select
    End Sub

    '依傳入資料 後續動作執行
    Sub SU_Create2()
        Select Case rqProecess
            Case "shift"
                If rqFrom_type <> "add" Then 'by mick
                    'EnterChannel.Enabled=True
                    Common.SetListItem(EnterChannel, "4")
                    EnterChannel.Enabled = False '1.網;2.現;3.通;4.推
                    TIMS.Tooltip(EnterChannel, "由三合一轉入，方式為推介")
                    'EnterChannel.Items(4).Selected=True
                    IDNO.ReadOnly = True
                    '由三合一轉入，不能挑選第二第三志願 'ticket=" & TRNDMode
                    'If Hid_rqTicket.Value <> "2" Then
                    '    Button5.Disabled=True
                    '    'Button2.Disabled=True
                    '    'Button3.Disabled=True
                    '    TIMS.Tooltip(Button5, "由三合一轉入，不能挑選第二第三志願")
                    '    'TIMS.Tooltip(Button2, "由三合一轉入，不能挑選第二第三志願")
                    '    'TIMS.Tooltip(Button3, "由三合一轉入，不能挑選第二第三志願")
                    'End If
                Else
                    '如新增時檢查到idno有三合一資料時，加入此檢查 by mick
                    Button1.Attributes("onclick") = "javascript:return chkadp();"
                End If
            Case Else
                If Request("view") = "1" Then
                    If TIMS.Cst_TPlanID06Plan1.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                        'ACTNObli
                        '不予補助對象，惠請查明該筆民眾身分是否符合本計畫參訓資格
                        Dim flagAct60 As Boolean = True
                        Dim msgAct60 As String = ""
                        flagAct60 = TIMS.Chk_ActNoALTERMSG1(Hid_ACTNObli.Value, msgAct60)
                        If flagAct60 AndAlso msgAct60 <> "" Then Common.MessageBox(Me, msgAct60)
                    End If
                    Call ItemVisibleState()
                End If
        End Select
    End Sub

    Sub SU_Proecess_add(ByRef ss As String)
        Dim rqSerial As String = TIMS.GetMyValue(ss, "rqSerial")
        Dim rqIDNO As String = TIMS.GetMyValue(ss, "rqIDNO")
        rqIDNO = TIMS.ChangeIDNO(rqIDNO)
        Dim sql As String = " SELECT * FROM STUD_ENTERTEMP a where 1<>1 "
        If rqSerial <> "" Then
            sql = $" SELECT * FROM STUD_ENTERTEMP a where a.SETID={TIMS.CINT1(rqSerial)}"
        ElseIf rqIDNO <> "" Then
            sql = $" SELECT * FROM STUD_ENTERTEMP a where a.IDNO='{rqIDNO}'"
        End If
        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)

        If rqIDNO.Length = 10 Then      '假如為身分證號碼
            Select Case rqIDNO.Chars(1)
                Case "1"
                    Sex.Items(0).Selected = True
                Case "2"
                    Sex.Items(1).Selected = True
            End Select
            PassPortNO.Items(0).Selected = True
        End If
        ivsExistence = 0
        ivsSETID = 0
        EnterDate.Value = aNow.Date
        RelEnterDate.Text = aNow.Date
        If dr Is Nothing Then Exit Sub

        ivsExistence = 1           '表示資料存在於Stud_EnterTemp
        ivsSETID = dr("SETID")

        '純文字欄
        ExamID.Text = "(系統儲存自動產生)"
        Name.Text = dr("Name")
        birthday.Text = ""
        If Convert.ToString(dr("Birthday")) <> "" AndAlso IsDate(dr("Birthday")) Then birthday.Text = CDate(dr("Birthday")).ToString("yyyy/MM/dd")

        IDNO.Text = TIMS.ChangeIDNO(TIMS.ClearSQM($"{dr("IDNO")}"))
        IDNO.ReadOnly = True
        School.Text = Convert.ToString(dr("School"))
        Department.Text = Convert.ToString(dr("Department"))

        '---------------
        TBCity.Text = ""
        city_code.Value = ""
        ZipCODEB3.Value = ""
        If Convert.ToString(dr("ZipCode")) <> "" Then
            city_code.Value = Convert.ToString(dr("ZipCode")) 'TIMS.AddZero(Convert.ToString(dr("ZipCode")), 3)
            hidZipCODE6W.Value = Convert.ToString(dr("ZipCODE6W"))
            ZipCODEB3.Value = TIMS.GetZIPCODEB3(Convert.ToString(dr("ZipCODE6W")))
            TBCity.Text = TIMS.GET_FullCCTName(objconn, Convert.ToString(dr("ZipCode")), Convert.ToString(dr("ZipCODE6W")))
        End If
        Address.Text = Convert.ToString(dr("Address"))
        '---------------
        Phone1.Text = dr("Phone1").ToString
        Phone2.Text = dr("Phone2").ToString
        Email.Text = dr("Email").ToString
        CellPhone.Text = dr("CellPhone").ToString
        Dim v_rblMobil_YN As String = If(TIMS.ClearSQM(CellPhone.Text) <> "", "Y", "N")
        Common.SetListItem(rblMobil, v_rblMobil_YN)

        'notes.Text=dr("notes").ToString
        Common.SetListItem(PassPortNO, dr("PassPortNO").ToString)
        Common.SetListItem(Sex, dr("Sex").ToString)

        Dim TMPVAL As String = Convert.ToString(dr("MaritalStatus"))
        Common.SetListItem(MaritalStatus, TMPVAL)
        Common.SetListItem(GradID, dr("GradID").ToString)
        Common.SetListItem(DegreeID, dr("DegreeID").ToString)
        If Convert.ToString(dr("MilitaryID")) <> "" Then Common.SetListItem(MilitaryID, dr("MilitaryID").ToString)
        'Common.SetListItem(IsAgree, dr("IsAgree").ToString)

        EnterDate.Value = aNow.Date
        RelEnterDate.Text = aNow.Date

        'rblWorkSuppIdent.Enabled=True
        'If HidMaster.Value="Y" Then
        '    '具公司/商業負責人身分 '限定計畫執行 '鎖定為在職者
        '    Common.SetListItem(rblWorkSuppIdent, "Y")
        '    rblWorkSuppIdent.Enabled=False
        '    TIMS.Tooltip(rblWorkSuppIdent, cst_master1)
        'End If
        'If Convert.ToString(dr("CMASTER1"))="Y" Then
        '    '具公司/商業負責人身分 '限定計畫執行 '鎖定為在職者
        '    Common.SetListItem(rblWorkSuppIdent, "Y")
        '    rblWorkSuppIdent.Enabled=False
        'End If

    End Sub

    ''' <summary>
    ''' 報名登錄附加資料1/2
    ''' </summary>
    ''' <param name="htSS"></param>
    ''' <returns></returns>
    Function GET_ENTERTRAIN(ByRef htSS As Hashtable) As DataRow
        Dim uSETID As String = TIMS.GetMyValue2(htSS, "SETID")
        Dim uEnterDate As String = TIMS.GetMyValue2(htSS, "EnterDate")
        Dim uSerNum As String = TIMS.GetMyValue2(htSS, "SerNum")
        Dim uESERNUM As String = TIMS.GetMyValue2(htSS, "ESERNUM")

        Dim PMS1 As New Hashtable From {
            {"SETID", TIMS.CINT1(uSETID)},
            {"EnterDate", TIMS.Cdate2(uEnterDate)},
            {"SerNum", TIMS.CINT1(uSerNum)}
        }
        Dim sql As String = ""
        sql &= " SELECT a.SENID  /*PK*/" & vbCrLf
        sql &= " ,a.SETID" & vbCrLf
        sql &= " ,a.ENTERDATE" & vbCrLf
        sql &= " ,a.SERNUM" & vbCrLf
        sql &= " ,a.ZIPCODE2" & vbCrLf
        sql &= " ,a.ZIPCODE2_6W" & vbCrLf
        sql &= " ,a.HOUSEHOLDADDRESS" & vbCrLf
        sql &= " ,a.HANDTYPEID" & vbCrLf
        sql &= " ,a.HANDLEVELID" & vbCrLf
        sql &= " ,a.UNAME" & vbCrLf
        sql &= " ,a.INTAXNO" & vbCrLf
        sql &= " ,a.SERVDEPT" & vbCrLf
        sql &= " ,a.JOBTITLE" & vbCrLf
        sql &= " ,a.ACTNAME" & vbCrLf
        sql &= " ,a.ACTTYPE" & vbCrLf
        sql &= " ,a.ACTNO" & vbCrLf
        sql &= " ,a.ACTTEL" & vbCrLf
        sql &= " ,a.ZIPCODE3" & vbCrLf
        sql &= " ,a.ZIPCODE3_6W" & vbCrLf
        sql &= " ,a.ACTADDRESS" & vbCrLf
        sql &= " ,a.SERVDEPTID" & vbCrLf
        sql &= " ,a.JOBTITLEID" & vbCrLf
        sql &= " ,a.MODIFYACCT" & vbCrLf
        sql &= " ,a.MODIFYDATE" & vbCrLf
        sql &= " FROM STUD_ENTERTRAIN a WITH(NOLOCK)" & vbCrLf
        sql &= " WHERE a.SETID=@SETID AND a.EnterDate=@EnterDate AND a.SerNum=@SerNum"
        Dim sqldr As DataRow = DbAccess.GetOneRow(sql, objconn, PMS1)
        If sqldr IsNot Nothing Then Return sqldr
        If uESERNUM = "" Then Return sqldr

        Dim parms2 As New Hashtable From {{"ESERNUM", TIMS.CINT1(uESERNUM)}}
        Dim sql2 As String = ""
        sql2 &= String.Format(" SELECT {0} SETID", uSETID) & vbCrLf
        sql2 &= String.Format(" ,{0} ENTERDATE", TIMS.To_date(uEnterDate)) & vbCrLf
        sql2 &= String.Format(" ,{0} SERNUM", uSerNum) & vbCrLf
        sql2 &= String.Format(" ,{0} SENID", 0) & vbCrLf 'SENID(0)
        sql2 &= " ,a.ZIPCODE2" & vbCrLf
        sql2 &= " ,a.ZIPCODE2_6W" & vbCrLf
        sql2 &= " ,a.HOUSEHOLDADDRESS" & vbCrLf
        sql2 &= " ,a.HANDTYPEID" & vbCrLf
        sql2 &= " ,a.HANDLEVELID" & vbCrLf
        sql2 &= " ,a.UNAME" & vbCrLf
        sql2 &= " ,a.INTAXNO" & vbCrLf
        sql2 &= " ,a.SERVDEPT" & vbCrLf
        sql2 &= " ,a.JOBTITLE" & vbCrLf
        sql2 &= " ,a.ACTNAME" & vbCrLf
        sql2 &= " ,a.ACTTYPE" & vbCrLf
        sql2 &= " ,a.ACTNO" & vbCrLf
        sql2 &= " ,a.ACTTEL" & vbCrLf
        sql2 &= " ,a.ZIPCODE3" & vbCrLf
        sql2 &= " ,a.ZIPCODE3_6W" & vbCrLf
        sql2 &= " ,a.ACTADDRESS" & vbCrLf
        sql2 &= " ,a.SERVDEPTID" & vbCrLf
        sql2 &= " ,a.JOBTITLEID" & vbCrLf
        sql2 &= " ,a.MODIFYACCT" & vbCrLf
        sql2 &= " ,a.MODIFYDATE" & vbCrLf
        sql2 &= " FROM STUD_ENTERTRAIN2 a WITH(NOLOCK)" & vbCrLf
        sql2 &= " WHERE a.ESERNUM=@ESERNUM" & vbCrLf

        sqldr = DbAccess.GetOneRow(sql2, objconn, parms2)
        If sqldr IsNot Nothing Then Return sqldr

        Return sqldr
    End Function

    Sub SU_Proecess_edit(ByRef ss As String)
        Dim rqIDNO As String = TIMS.GetMyValue(ss, "rqIDNO")
        '修改動作
        Dim PMS1 As New Hashtable From {
            {"SETID", TIMS.CINT1(R_serial.Value)},
            {"EnterDate", TIMS.Cdate2(R_EnterDate.Value)},
            {"SerNum", TIMS.CINT1(R_SerNum.Value)}
        }
        Dim sql As String = ""
        sql &= " SELECT s.IDNO,s.NAME,s.SEX,s.BIRTHDAY,s.PASSPORTNO,s.MARITALSTATUS" & vbCrLf
        sql &= " ,s.DEGREEID,s.GRADID,s.SCHOOL,s.DEPARTMENT,s.MILITARYID,s.ZIPCODE,s.ZIPCODE6W" & vbCrLf
        sql &= " ,s.ADDRESS,s.PHONE1,s.PHONE2,s.CELLPHONE,s.EMAIL" & vbCrLf
        sql &= " ,s.NOTES,s.ISAGREE,s.LAINFLAG,s.MODIFYACCT,s.MODIFYDATE,s.ESETID" & vbCrLf
        sql &= " ,t.SETID,t.EnterDate,t.SerNum,t.ESERNUM,t.RelEnterDate,t.ExamNo,t.EnterChannel,t.OCID1,t.TMID1,t.OCID2,t.TMID2,t.OCID3,t.TMID3" & vbCrLf
        sql &= " ,t.IdentityID ,t.MIDENTITYID" & vbCrLf
        sql &= " ,t.HighEduBg,t.WorkSuppIdent,t.CCLID,t.TRNDMode,t.TICKET_NO,t.PriorWorkType1,t.PriorWorkOrg1,t.ActNo,t.SOfficeYM1,t.FOfficeYM1,t.notes tNotes,t.CMASTER1,t.APID1" & vbCrLf
        sql &= " ,c.ACTNO ACTNObli" & vbCrLf ''ACTNObli
        sql &= " FROM STUD_ENTERTEMP s" & vbCrLf
        sql &= " JOIN STUD_ENTERTYPE t ON t.SETID=s.SETID" & vbCrLf 'ESETID/ESERNUM
        sql += " LEFT JOIN STUD_SELRESULTBNG c ON c.SETID=t.SETID and c.EnterDate=t.EnterDate and c.SerNum=t.SerNum and c.OCID=t.OCID1" & vbCrLf
        sql &= " WHERE t.SETID=@SETID AND t.EnterDate=@EnterDate AND t.SerNum=@SerNum"
        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn, PMS1)
        If dr Is Nothing Then
            Common.MessageBox(Me, "傳入參數有誤，請重新查詢。")
            Exit Sub
        End If

        '(2)「報名志願」選完，報名資料儲存後，不可再修改。
        '(報名登錄作業，如屬於新增學員報名資料，在新增狀態時「報名志願」該欄位是可以被修正
        '如果正式送出後，欄位是不得開放修正的。
        If dr IsNot Nothing Then
            If Convert.ToString(dr("OCID1")) <> "" Then
                BtnClear1.Visible = False
                'BtnClear2.Visible=False
                'BtnClear3.Visible=False
                Button5.Disabled = True
                'Button2.Disabled=True
                'Button3.Disabled=True
                Dim sTitle As String = "報名班級已送出，不能再修改報名班級" '自願"
                TIMS.Tooltip(Button5, sTitle)
                'TIMS.Tooltip(Button2, sTitle)
                'TIMS.Tooltip(Button3, sTitle)
            End If
        End If

        Dim s_ESERNUM As String = Convert.ToString(dr("ESERNUM"))
        Dim parms1 As New Hashtable
        parms1.Add("SETID", R_serial.Value)
        parms1.Add("EnterDate", R_EnterDate.Value)
        parms1.Add("SerNum", R_SerNum.Value)
        parms1.Add("ESERNUM", s_ESERNUM)
        '報名登錄附加資料1
        Dim sqldr As DataRow = GET_ENTERTRAIN(parms1)

        ServDept.Visible = False
        JobTitle.Visible = False
        ddlSERVDEPTID.Visible = True
        ddlJOBTITLEID.Visible = True
        If sqldr IsNot Nothing Then
            '*戶籍地址
            ZipCode2.Value = TIMS.TrimZipCode(Convert.ToString(sqldr("ZIPCODE2")), dtZipCode)
            ZipCode2_B3.Value = TIMS.GetZIPCODEB3(sqldr("ZIPCODE2_6W")) 'ZipCode2_B3.Value=TIMS.TrimZipCODEB3(Convert.ToString(sqldr("ZIPCODE2_6W")))
            HidZipCode2_6W.Value = Convert.ToString(sqldr("ZIPCODE2_6W"))
            City2.Text = TIMS.getZipName2(ZipCode2.Value, ZipCode2_B3.Value, dtZipCode)
            HouseholdAddress.Text = Convert.ToString(sqldr("HouseholdAddress"))
            '服務單位資料 '*服務單位
            If Not IsDBNull(sqldr("Uname")) Then Uname.Text = sqldr("Uname").ToString
            '統一編號
            If Not IsDBNull(sqldr("Intaxno")) Then Intaxno.Text = sqldr("Intaxno").ToString
            '*服務部門 ServDept 30 CHAR
            If Not IsDBNull(sqldr("ServDept")) Then ServDept.Text = sqldr("ServDept").ToString
            If Convert.ToString(sqldr("SERVDEPTID")) <> "" Then Common.SetListItem(ddlSERVDEPTID, sqldr("SERVDEPTID"))
            '*投保單位名稱
            If Not IsDBNull(sqldr("Actname")) Then ActName.Text = sqldr("Actname").ToString
            '*投保類別
            If Convert.ToString(sqldr("ActType")) <> "" Then Common.SetListItem(ActType, Convert.ToString(sqldr("ActType")))
            '投保單位保險證號
            If Not IsDBNull(sqldr("ActNo")) Then ActNo.Text = sqldr("ActNo").ToString
            '*職稱/職務
            If Not IsDBNull(sqldr("JobTitle")) Then JobTitle.Text = sqldr("JobTitle").ToString
            If Convert.ToString(sqldr("JOBTITLEID")) <> "" Then Common.SetListItem(ddlJOBTITLEID, sqldr("JOBTITLEID"))
            '投保單位電話
            If Not IsDBNull(sqldr("ActTel")) Then ActTel.Text = sqldr("ActTel").ToString
            '投保單位地址
            ZipCode3.Value = TIMS.TrimZipCode(Convert.ToString(sqldr("ZipCode3")), dtZipCode)
            ZipCode3_B3.Value = TIMS.GetZIPCODEB3(sqldr("ZipCode3_6W"))
            HidZipCode3_6W.Value = Convert.ToString(sqldr("ZipCode3_6W"))
            City3.Text = TIMS.getZipName2(ZipCode3.Value, ZipCode3_B3.Value, dtZipCode)
            ActAddress.Text = Convert.ToString(sqldr("ActAddress"))
        End If

        '表示資料存在於Stud_EnterTemp
        ivsExistence = 1
        ivsSETID = dr("SETID") '純文字欄
        ExamID.Text = Convert.ToString(dr("ExamNo")) 'asp:Label
        ExamNo.Value = Convert.ToString(dr("ExamNo")) 'hidden
        Name.Text = dr("Name")
        birthday.Text = ""
        If IsDate(dr("Birthday")) Then birthday.Text = CDate(dr("Birthday")).ToString("yyyy/MM/dd")

        'IDNO.Text=TIMS.ChangeIDNO(dr("IDNO"))
        IDNO.Text = TIMS.ChangeIDNO(TIMS.ClearSQM($"{dr("IDNO")}"))
        IDNO.ReadOnly = True '20090511(Milor)鎖定身分證號不能修改。
        IDNO.Attributes("onchange") += "if(document.form1.IDNO.value=='" & dr("IDNO") & "'){document.form1.IDNOChange.value='0'}else{document.form1.IDNOChange.value='1'};"
        IDNO.Attributes("onblur") += "if(document.form1.IDNO.value=='" & dr("IDNO") & "'){document.form1.IDNOChange.value='0'}else{document.form1.IDNOChange.value='1'};"

        School.Text = Convert.ToString(dr("School"))
        Department.Text = Convert.ToString(dr("Department"))

        Phone1.Text = dr("Phone1").ToString
        Phone2.Text = dr("Phone2").ToString
        Email.Text = dr("Email").ToString
        CellPhone.Text = dr("CellPhone").ToString
        Dim v_MobilYN As String = If(TIMS.ClearSQM(CellPhone.Text) <> "", "Y", "N")
        Common.SetListItem(rblMobil, v_MobilYN)
        notes.Text = dr("tNotes").ToString

        TBCity.Text = ""
        city_code.Value = ""
        ZipCODEB3.Value = ""
        If Convert.ToString(dr("ZipCode")) <> "" Then
            city_code.Value = Convert.ToString(dr("ZipCode")) 'TIMS.AddZero(Convert.ToString(dr("ZipCode")), 3)
            ZipCODEB3.Value = TIMS.GetZIPCODEB3(Convert.ToString(dr("ZipCODE6W")))
            hidZipCODE6W.Value = Convert.ToString(dr("ZipCODE6W"))
            TBCity.Text = TIMS.GET_FullCCTName(objconn, Convert.ToString(dr("ZipCode")), Convert.ToString(dr("ZipCODE6W")))
        End If
        Address.Text = Convert.ToString(dr("Address"))

        Common.SetListItem(PassPortNO, dr("PassPortNO").ToString)
        Common.SetListItem(Sex, dr("Sex").ToString)

        Dim TMPVAL As String = Convert.ToString(dr("MaritalStatus"))
        Common.SetListItem(MaritalStatus, TMPVAL)
        Common.SetListItem(GradID, dr("GradID").ToString)
        Common.SetListItem(DegreeID, dr("DegreeID").ToString)
        If Convert.ToString(dr("MilitaryID")) <> "" Then Common.SetListItem(MilitaryID, dr("MilitaryID").ToString)
        'Common.SetListItem(IsAgree, dr("IsAgree").ToString)
        Common.SetListItem(EnterChannel, dr("EnterChannel").ToString)

        '(1)「報名管道」如果是『1.網路』與『4.推介』的不可修改，
        '只有『2.現場』與『3.通訊』的，可以修改。
        If rqProecess = "edit" Then
            '1.網;2.現;3.通;4.推
            Select Case EnterChannel.SelectedValue
                Case "1"
                    EnterChannel.Enabled = False
                    Dim sTitle As String = "報名管道「網路」不能再修改"
                    TIMS.Tooltip(EnterChannel, sTitle)
                    IMG1.Visible = False
                    RelEnterDate.ReadOnly = True
                Case "4"
                    EnterChannel.Enabled = False
                    Dim sTitle As String = "報名管道「推介」不能再修改"
                    TIMS.Tooltip(EnterChannel, sTitle)
            End Select
        End If

        hide_EnterChannel.Value = Convert.ToString(dr("EnterChannel"))
        OCIDValue1.Value = TIMS.Change0(dr("OCID1").ToString)
        TMIDValue1.Value = TIMS.Change0(dr("TMID1").ToString)
        'OCIDValue2.Value=TIMS.Change0(dr("OCID2").ToString)
        'TMIDValue2.Value=TIMS.Change0(dr("TMID2").ToString)
        'OCIDValue3.Value=TIMS.Change0(dr("OCID3").ToString)
        'TMIDValue3.Value=TIMS.Change0(dr("TMID3").ToString)
        'hide_TrainMode.Value=TIMS.Get_GOVTRNData(Convert.ToString(dr("OCID1")), IDNO.Text, objconn)
        hide_TrainMode.Value = ""
        If hide_TrainMode.Value <> "" Then Common.SetListItem(EnterChannel, "4")

        ''12:屆退官兵(須單位將級以上長官薦送函)(職前課程邏輯)
        'If TIMS.Cst_TPlanID02Plan2.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '    Const cst_id12 As String="12"
        '    Dim flag_ChkID12 As Boolean=False '未勾選
        '    If Convert.ToString(dr("IdentityID")).IndexOf(cst_id12) > -1 Then
        '        flag_ChkID12=True '已勾選 屆退官兵
        '    End If
        '    If Not flag_ChkID12 Then '未勾選 屆退官兵
        '        Dim flag_SRSOLDIERS As Boolean=False '是否為屆退官兵
        '        flag_SRSOLDIERS=TIMS.CheckRESOLDER(objconn, rqIDNO, sm.UserInfo.DistID, TIMS.cdate3(aNow))
        '        Dim MyValue As String=Convert.ToString(dr("IdentityID"))
        '        If flag_SRSOLDIERS Then
        '            '為屆退官兵 且資料並無勾選 屆退官兵
        '            If MyValue <> "" Then MyValue &= ","
        '            MyValue &= cst_id12
        '            dr("IdentityID")=MyValue '重新填入身分別
        '        End If
        '    End If
        'End If

        Dim drET2 As DataRow = If(s_ESERNUM <> "", TIMS.Get_ENTERTYPE2(s_ESERNUM, objconn), Nothing)
        Dim s_MIdentityID As String = Convert.ToString(dr("MIDENTITYID"))
        If s_MIdentityID = "" AndAlso drET2 IsNot Nothing Then s_MIdentityID = Convert.ToString(drET2("MIDENTITYID"))

        Common.SetListItem(MIdentityID, s_MIdentityID)

        If Convert.ToString(dr("IdentityID")) <> "" Then Call TIMS.SetCblValue(IdentityID, Convert.ToString(dr("IdentityID")))
        EnterDate.Value = ""
        If Convert.ToString(dr("EnterDate")) <> "" AndAlso IsDate(dr("EnterDate")) Then EnterDate.Value = CDate(dr("EnterDate")).ToString("yyyy/MM/dd")
        RelEnterDate.Text = ""
        If Convert.ToString(dr("RelEnterDate")) <> "" AndAlso IsDate(dr("RelEnterDate")) Then RelEnterDate.Text = CDate(dr("RelEnterDate")).ToString("yyyy/MM/dd")
        'If Convert.ToString(dr("HighEduBg")) <> "" Then Common.SetListItem(rdo_HighEduBg, dr("HighEduBg"))
        If Convert.ToString(dr("WorkSuppIdent")) <> "" Then Common.SetListItem(rblWorkSuppIdent, dr("WorkSuppIdent"))

        If TIMS.Cst_TPlanID46AppPlan5.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Common.SetListItem(rblWorkSuppIdent, "Y")
            rblWorkSuppIdent.Enabled = False '鎖定為在職者
            'TIMS.Tooltip(rblWorkSuppIdent, cst_master1)
        End If

        '獲得職訓課程管道 cblAVTCP1
        Dim s_APID1 As String = Convert.ToString(dr("APID1"))
        If s_APID1 = "" AndAlso drET2 IsNot Nothing Then s_APID1 = Convert.ToString(drET2("APID1"))
        If s_APID1 <> "" Then Call TIMS.SetCblValue(cblAVTCP1, s_APID1)

        CCLID.Value = TIMS.Change0(dr("CCLID").ToString)

        'If dr("CCLID").ToString <> "" Then
        '    Button2.Disabled=True
        '    Button3.Disabled=True
        'End If

        '(職前邏輯)
        'DGTR.Visible=False
        'GovTR.Visible=False
        'Select Case dr("TRNDMode").ToString
        '    Case "2"
        '        If dr("TICKET_NO").ToString <> "" Then
        '            sql=""
        '            sql &= " SELECT b.Share_Name"
        '            sql &= " FROM Adp_DGTRNData a"
        '            sql &= " LEFT JOIN (SELECT * FROM Adp_ShareSource  WHERE Share_Type='301') b ON a.OBJECT_TYPE=b.Share_ID "
        '            sql &= " WHERE a.TICKET_NO='" & dr("TICKET_NO").ToString & "'"
        '            OBJECT_TYPE.Text=DbAccess.ExecuteScalar(sql, objconn).ToString
        '            DGTR.Visible=True
        '        Else
        '            'sql="SELECT b.Share_Name FROM "
        '            'sql += "Adp_DGTRNData a "
        '            'sql += "LEFT JOIN (SELECT * FROM Adp_ShareSource WHERE Share_Type='301') b ON a.OBJECT_TYPE=b.Share_ID "
        '            'sql += "WHERE a.TICKET_NO='" & dr("TICKET_NO").ToString & "'"
        '            OBJECT_TYPE.Text="(無學習卷號碼)狀況異常請查明" '特殊狀況'DbAccess.ExecuteScalar(sql).ToString
        '            OBJECT_TYPE.ForeColor=Color.Red
        '            OBJECT_TYPE.Font.Italic=True
        '            OBJECT_TYPE.Font.Bold=True
        '            DGTR.Visible=True
        '        End If
        '    Case "3"
        '        If dr("TICKET_NO").ToString <> "" Then
        '            sql=""
        '            sql &= " SELECT b.Share_Name "
        '            sql &= " FROM Adp_GOVTRNData a  "
        '            sql &= " LEFT JOIN (SELECT * FROM Adp_ShareSource WHERE Share_Type='527') b ON a.OBJECT_TYPE=b.Share_ID "
        '            sql &= " WHERE a.TICKET_NO='" & dr("TICKET_NO").ToString & "'"
        '            OBJECT_TYPE2.Text=DbAccess.ExecuteScalar(sql, objconn).ToString

        '            sql=""
        '            sql &= " SELECT b.Share_Name "
        '            sql &= " FROM Adp_GOVTRNData a  "
        '            sql &= " LEFT JOIN (SELECT * FROM Adp_ShareSource WHERE Share_Type='528') b ON a.SPECIAL_TYPE=b.Share_ID "
        '            sql &= " WHERE a.TICKET_NO='" & dr("TICKET_NO").ToString & "'"
        '            SPECIAL_TYPE.Text=DbAccess.ExecuteScalar(sql, objconn).ToString
        '        Else
        '            Dim drG As DataRow
        '            drG=TIMS.Get_Adp_GOVTRNData(dr("IDNO").ToString, dr("OCID1").ToString)
        '            If Not drG Is Nothing Then
        '                If drG("TICKET_NO").ToString <> "" Then
        '                    sql=""
        '                    sql &= " UPDATE STUD_ENTERTYPE" & vbCrLf
        '                    sql &= " SET TICKET_NO='" & drG("TICKET_NO").ToString & "'" & vbCrLf
        '                    sql &= " where EnterDate= " & TIMS.to_date(R_EnterDate.Value) & " and SerNum='" & R_SerNum.Value & "' AND SETID='" & R_serial.Value & "'" & vbCrLf
        '                    DbAccess.ExecuteNonQuery(sql, objconn)

        '                    sql=""
        '                    sql &= " SELECT b.Share_Name"
        '                    sql &= " FROM Adp_GOVTRNData a  "
        '                    sql &= " LEFT JOIN (SELECT * FROM Adp_ShareSource WHERE Share_Type='527') b ON a.OBJECT_TYPE=b.Share_ID "
        '                    sql &= " WHERE a.TICKET_NO='" & drG("TICKET_NO").ToString & "'"
        '                    OBJECT_TYPE2.Text=DbAccess.ExecuteScalar(sql, objconn).ToString

        '                    sql=""
        '                    sql &= " SELECT b.Share_Name "
        '                    sql &= " FROM Adp_GOVTRNData a  "
        '                    sql &= " LEFT JOIN (SELECT * FROM Adp_ShareSource WHERE Share_Type='528') b ON a.SPECIAL_TYPE=b.Share_ID "
        '                    sql &= " WHERE a.TICKET_NO='" & drG("TICKET_NO").ToString & "'"
        '                    SPECIAL_TYPE.Text=DbAccess.ExecuteScalar(sql, objconn).ToString
        '                End If
        '            End If
        '        End If

        '        GovTR.Visible=True
        'End Select

        ''--------------------受訓前任職資料start----------------
        'If Convert.ToString(dr("PriorWorkType1")) <> "" Then
        '    Common.SetListItem(PriorWorkType1, dr("PriorWorkType1"))
        'End If
        'PriorWorkOrg1.Text=Convert.ToString(dr("PriorWorkOrg1"))
        'ActNo.Text=Convert.ToString(dr("ActNo"))
        ''ACTNObli
        'Hid_ACTNObli.Value=Convert.ToString(dr("ACTNObli"))
        'If Convert.ToString(dr("SOfficeYM1")) <> "" Then
        '    SOfficeYM1.Text=TIMS.cdate3(dr("SOfficeYM1"))
        'End If
        'If Convert.ToString(dr("FOfficeYM1")) <> "" Then
        '    FOfficeYM1.Text=TIMS.cdate3(dr("FOfficeYM1"))
        'End If
        ''--------------------受訓前任職資料end------------------

        '異常郵遞區號修正
        Dim i_zipcodeLen As Integer = Convert.ToString(dr("ZipCode")).Length
        If i_zipcodeLen = 5 OrElse i_zipcodeLen = 6 Then
            Try
                Dim xSql As String = ""
                xSql &= " UPDATE STUD_ENTERTEMP SET ZipCODE6W=ZipCode ,ZipCode=LEFT(ZipCode,3)" & vbCrLf
                xSql &= " WHERE Len(ZipCode)=5 AND ZipCODE6W IS NULL" & vbCrLf
                DbAccess.ExecuteNonQuery(xSql, objconn)

                Dim xSql2 As String = ""
                xSql2 &= " UPDATE STUD_ENTERTEMP SET ZipCODE6W=ZipCode ,ZipCode=LEFT(ZipCode,3)" & vbCrLf
                xSql2 &= " WHERE Len(ZipCode)=6 AND ZipCODE6W IS NULL" & vbCrLf
                DbAccess.ExecuteNonQuery(xSql2, objconn)
            Catch ex As Exception
                TIMS.WriteLog(Me, ex.Message)
                'Throw ex
            End Try
        End If

        '假如報名班級被試算過，則不能修改自願
        Dim pms2 As New Hashtable From {
            {"SETID", TIMS.CINT1(dr("SETID"))},
            {"EnterDate", TIMS.Cdate2(dr("EnterDate"))},
            {"SerNum", TIMS.CINT1(dr("SerNum"))}
        }
        sql = " SELECT * FROM STUD_SELRESULT WHERE SETID=@SETID AND EnterDate=@EnterDate AND SerNum=@SerNum"
        Dim drRE As DataRow = DbAccess.GetOneRow(sql, objconn, pms2)
        If drRE IsNot Nothing Then
            BtnClear1.Visible = False
            'BtnClear2.Visible=False
            'BtnClear3.Visible=False
            Button5.Disabled = True
            'Button2.Disabled=True
            'Button3.Disabled=True
            Dim sTitle As String = "報名班級被試算過，不能再修改自願"
            TIMS.Tooltip(Button5, sTitle)
            'TIMS.Tooltip(Button2, sTitle)
            'TIMS.Tooltip(Button3, sTitle)
        End If

        '先找出三個志願班的統一廠商編號
        'Dim drOC1 As DataRow=TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        'Dim drOC2 As DataRow=TIMS.GetOCIDDate(OCIDValue2.Value, objconn)
        'Dim drOC3 As DataRow=TIMS.GetOCIDDate(OCIDValue3.Value, objconn)
        Dim drOC1 As DataRow = Nothing  'edit，by:20181025
        If OCIDValue1.Value.Trim <> "" Then drOC1 = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)  'edit，by:20181025
        'Dim drOC2 As DataRow=Nothing  'edit，by:20181025
        'If OCIDValue2.Value.Trim <> "" Then drOC2=TIMS.GetOCIDDate(OCIDValue2.Value, objconn)  'edit，by:20181025
        'Dim drOC3 As DataRow=Nothing  'edit，by:20181025
        'If OCIDValue3.Value.Trim <> "" Then drOC3=TIMS.GetOCIDDate(OCIDValue3.Value, objconn)  'edit，by:20181025

        If drOC1 IsNot Nothing Then
            ComIDNO1.Value = Convert.ToString(drOC1("COMIDNO"))
            SeqNO1.Value = Convert.ToString(drOC1("SeqNO"))
            TMID1.Text = "[" & drOC1("TrainID") & "]" & drOC1("TrainName")
            OCID1.Text = Convert.ToString(drOC1("ClassCName2"))
            Session("wish1_date") = TIMS.Cdate3(drOC1("STDate")) '.ToString("yyyy/MM/dd")
        End If
        'If Not drOC2 Is Nothing Then
        '    ComIDNO2.Value=Convert.ToString(drOC2("COMIDNO"))
        '    SeqNO2.Value=Convert.ToString(drOC2("SeqNO"))
        '    TMID2.Text="[" & drOC2("TrainID") & "]" & drOC2("TrainName")
        '    OCID2.Text=Convert.ToString(drOC2("ClassCName2")) '
        'End If
        'If Not drOC3 Is Nothing Then
        '    ComIDNO3.Value=Convert.ToString(drOC3("COMIDNO"))
        '    SeqNO3.Value=Convert.ToString(drOC3("SeqNO"))
        '    TMID3.Text="[" & drOC3("TrainID") & "]" & drOC3("TrainName")
        '    OCID3.Text=Convert.ToString(drOC3("ClassCName2")) '
        'End If
        If CCLID.Value <> "" AndAlso OCIDValue1.Value <> "" Then Call SUtl_ShowCL() '課程階段顯示
    End Sub

#Region "NOUSE"
    ' (職前訓練)ADP_STDDATA 就服三合一相關邏輯
    'Sub sU_Proecess_shift(ByRef ss As String)
    '    Dim rqIDNO As String=TIMS.GetMyValue(ss, "rqIDNO")
    '    Dim rqBIRTH As String=TIMS.GetMyValue(ss, "rqBIRTH")
    '    Dim rqTRN_UNKEY As String=TIMS.GetMyValue(ss, "rqTRN_UNKEY")
    '    Dim rqAPPLY_DATE As String=TIMS.GetMyValue(ss, "rqAPPLY_DATE")
    '    Dim rqTICKET_NO As String=TIMS.GetMyValue(ss, "rqTICKET_NO")

    '    Dim sql As String=""
    '    sql="SELECT * FROM ADP_STDDATA WHERE IDNO='" & rqIDNO & "'"
    '    Dim dr As DataRow=DbAccess.GetOneRow(sql, objconn)
    '    IDNO.ReadOnly=True
    '    If Not dr Is Nothing Then
    '        Name.Text=dr("Name").ToString
    '        birthday.Text=""
    '        If IsDate(dr("Birth")) Then
    '            birthday.Text=CDate(dr("Birth")).ToString("yyyy/MM/dd")
    '        End If

    '        'IDNO.Text=TIMS.ChangeIDNO(dr("IDNO").ToString)
    '        IDNO.Text=Convert.ToString(dr("IDNO"))
    '        If IDNO.Text <> "" Then IDNO.Text=Trim(IDNO.Text)
    '        If IDNO.Text <> "" Then IDNO.Text=UCase(IDNO.Text)
    '        If IDNO.Text <> "" Then IDNO.Text=TIMS.ChangeIDNO(IDNO.Text)

    '        If Convert.ToString(dr("School")) <> "" Then
    '            School.Text=dr("School").ToString
    '            If School.Text.IndexOf("&#") > -1 Then
    '                School.Text=Replace(School.Text, "&#", "& #")
    '            End If
    '            If Len(School.Text) > 30 Then
    '                School.Text=Left(School.Text, 30)
    '            End If
    '        Else
    '            School.Text=""
    '        End If
    '        Department.Text=dr("DeptName").ToString

    '        Phone1.Text=dr("Tel").ToString
    '        Email.Text=dr("Email").ToString
    '        CellPhone.Text=dr("Mobile").ToString
    '        If Trim(CellPhone.Text) <> "" Then
    '            Common.SetListItem(rblMobil, "Y")
    '        Else
    '            Common.SetListItem(rblMobil, "N")
    '        End If

    '        Common.SetListItem(Sex, dr("Sex").ToString)
    '        Select Case Convert.ToString(dr("Marri"))
    '            Case "1", "2"
    '                Common.SetListItem(MaritalStatus, dr("Marri").ToString)
    '            Case Else
    '                Common.SetListItem(MaritalStatus, "3")
    '        End Select

    '        Common.SetListItem(GradID, dr("Gradu").ToString)
    '        Common.SetListItem(DegreeID, dr("Edgr").ToString)
    '        If Convert.ToString(dr("Solder")) <> "" Then
    '            Common.SetListItem(MilitaryID, dr("Solder").ToString)
    '        End If

    '        '檢查三合一轉檔是否存在於temp檔中
    '        sql=""
    '        sql &= " select a.SETID,a.IDNO,a.NAME,a.SEX,a.BIRTHDAY,a.PASSPORTNO,a.MARITALSTATUS" & vbCrLf
    '        sql &= " ,a.DEGREEID,a.GRADID,a.SCHOOL,a.DEPARTMENT,a.MILITARYID,a.ZIPCODE" & vbCrLf
    '        sql &= " ,a.ADDRESS,a.PHONE1,a.PHONE2,a.CELLPHONE,a.EMAIL" & vbCrLf
    '        'Sql += " ,dbms_lob.substr( a.NOTES, 4000, 1 ) NOTES,a.ISAGREE,a.LAINFLAG" & vbCrLf
    '        sql &= " ,a.NOTES" & vbCrLf
    '        sql &= " ,a.ISAGREE,a.LAINFLAG" & vbCrLf
    '        sql &= " ,a.MODIFYACCT,a.MODIFYDATE,a.ESETID,a.ZipCODEB3" & vbCrLf
    '        sql &= " from Stud_EnterTemp a where a.IDNO='" & rqIDNO & "'"
    '        dr=DbAccess.GetOneRow(sql, objconn)
    '        If Not dr Is Nothing Then
    '            ivsSETID=dr("SETID")
    '            ivsExistence=1
    '        End If
    '    Else
    '        '檢查三合一轉檔是否存在於temp檔中
    '        sql=""
    '        sql &= " select a.SETID,a.IDNO,a.NAME,a.SEX,a.BIRTHDAY,a.PASSPORTNO,a.MARITALSTATUS" & vbCrLf
    '        sql &= " ,a.DEGREEID,a.GRADID,a.SCHOOL,a.DEPARTMENT,a.MILITARYID,a.ZIPCODE" & vbCrLf
    '        sql &= " ,a.ADDRESS,a.PHONE1,a.PHONE2,a.CELLPHONE,a.EMAIL" & vbCrLf
    '        'Sql += " ,dbms_lob.substr( a.NOTES, 4000, 1 ) NOTES,a.ISAGREE,a.LAINFLAG" & vbCrLf
    '        sql &= " ,a.NOTES" & vbCrLf
    '        sql &= " ,a.ISAGREE,a.LAINFLAG" & vbCrLf
    '        sql &= " ,a.MODIFYACCT,a.MODIFYDATE,a.ESETID,a.ZipCODEB3" & vbCrLf
    '        sql &= " FROM STUD_ENTERTEMP a" & vbCrLf
    '        sql &= " WHERE a.IDNO='" & rqIDNO & "'"
    '        dr=DbAccess.GetOneRow(sql, objconn)
    '        If Not dr Is Nothing Then
    '            ivsExistence=1
    '            ivsSETID=dr("SETID")

    '            Name.Text=dr("Name")
    '            birthday.Text=""
    '            If IsDate(dr("Birthday")) Then
    '                birthday.Text=CDate(dr("Birthday")).ToString("yyyy/MM/dd") 'FormatDateTime(dr("Birthday"), 2)
    '            End If

    '            'IDNO.Text=dr("IDNO").ToString
    '            IDNO.Text=Convert.ToString(dr("IDNO"))
    '            If IDNO.Text <> "" Then IDNO.Text=Trim(IDNO.Text)
    '            If IDNO.Text <> "" Then IDNO.Text=UCase(IDNO.Text)
    '            If IDNO.Text <> "" Then IDNO.Text=TIMS.ChangeIDNO(IDNO.Text)

    '            School.Text=dr("School").ToString
    '            Department.Text=dr("Department").ToString

    '            Phone1.Text=dr("Phone1").ToString
    '            Phone2.Text=dr("Phone2").ToString
    '            Email.Text=dr("Email").ToString
    '            CellPhone.Text=dr("CellPhone").ToString
    '            If Trim(CellPhone.Text) <> "" Then
    '                Common.SetListItem(rblMobil, "Y")
    '            Else
    '                Common.SetListItem(rblMobil, "N")
    '            End If
    '            'notes.Text=dr("notes").ToString
    '            '-----------------------
    '            TBCity.Text=""
    '            city_code.Value=""
    '            ZipCODEB3.Value=""
    '            Address.Text=""
    '            If dr("ZipCode").ToString <> "" Then
    '                city_code.Value=TIMS.AddZero(Convert.ToString(dr("ZipCode")), 3)
    '                If Convert.ToString(dr("ZipCODEB3")) <> "" Then
    '                    ZipCODEB3.Value=TIMS.AddZero(Convert.ToString(dr("ZipCODEB3")), 2)
    '                End If
    '                TBCity.Text=TIMS.Get_FullCTName(Convert.ToString(dr("ZipCode")), Convert.ToString(dr("ZipCODEB3")))
    '            End If
    '            Address.Text=dr("Address").ToString
    '            '--------------------------
    '            Common.SetListItem(PassPortNO, dr("PassPortNO").ToString)
    '            Common.SetListItem(Sex, dr("Sex").ToString)

    '            Select Case Convert.ToString(dr("MaritalStatus"))
    '                Case "1", "2"
    '                    Common.SetListItem(MaritalStatus, dr("MaritalStatus").ToString)
    '                Case Else
    '                    Common.SetListItem(MaritalStatus, "3")
    '            End Select

    '            Common.SetListItem(GradID, dr("GradID").ToString)
    '            Common.SetListItem(DegreeID, dr("DegreeID").ToString)
    '            If Convert.ToString(dr("MilitaryID")) <> "" Then
    '                Common.SetListItem(MilitaryID, dr("MilitaryID").ToString)
    '            End If
    '            'Common.SetListItem(IsAgree, dr("IsAgree").ToString)
    '        End If
    '    End If

    '    EnterDate.Value=aNow.Date
    '    RelEnterDate.Text=aNow.Date
    '    If rqFrom_type <> "add" Then
    '        Select Case rqTicket
    '            Case "1", "3"
    '                DGTR.Visible=False
    '                GovTR.Visible=True
    '                If rqTicket="1" Then
    '                    '職訓卷已經停用 Request("ticket")=1
    '                    sql="SELECT * FROM Adp_TRNData WHERE TRN_UNKEY='" & rqTRN_UNKEY & "' and IDNO='" & rqIDNO & "' and BIRTH= " & TIMS.to_date(rqBIRTH) & " and TICKET_NO='" & rqTICKET_NO & "' and APPLY_DATE= " & TIMS.to_date(rqAPPLY_DATE)
    '                Else
    '                    '其餘為推介單 Request("ticket")=3
    '                    sql="SELECT * FROM Adp_GOVTRNData WHERE TRN_UNKEY='" & rqTRN_UNKEY & "' and IDNO='" & rqIDNO & "' and BIRTH= " & TIMS.to_date(rqBIRTH) & " and TICKET_NO='" & rqTICKET_NO & "' and APPLY_DATE= " & TIMS.to_date(rqAPPLY_DATE)
    '                End If
    '                dr=DbAccess.GetOneRow(sql, objconn)
    '                If Not dr Is Nothing Then
    '                    If dr("TRN_CLASS").ToString <> "" Then
    '                        Dim drTemp As DataRow=TIMS.GetOCIDDate(dr("TRN_CLASS").ToString)
    '                        If Not drTemp Is Nothing Then
    '                            'OCID1.Text=drTemp("ClassCName")
    '                            
    '                            OCID1.Text=drTemp("ClassCName2")
    '                            TMID1.Text="[" & drTemp("TrainID") & "]" & drTemp("TrainName")
    '                            OCIDValue1.Value=dr("TRN_CLASS").ToString
    '                            TMIDValue1.Value=dr("JOB_TYPE").ToString
    '                        End If
    '                    End If
    '                End If

    '            Case "2" '學習卷
    '                sql=""
    '                sql &= " SELECT b.Share_Name "
    '                sql &= " FROM Adp_DGTRNData a"
    '                sql &= " LEFT JOIN (SELECT * FROM Adp_ShareSource  WHERE Share_Type='301') b ON a.OBJECT_TYPE=b.Share_ID "
    '                sql &= " WHERE a.TICKET_NO='" & rqTICKET_NO & "'"
    '                OBJECT_TYPE.Text=DbAccess.ExecuteScalar(sql, objconn).ToString
    '                DGTR.Visible=True
    '                GovTR.Visible=False
    '        End Select
    '    Else
    '        '比對3合1資料返回(check_class)
    '        Dim AGTDataT As DataTable
    '        Dim AGTDataR As DataRow
    '        sql="SELECT IDNO,TRN_CLASS FROM Adp_GOVTRNData  WHERE IDNO='" & rqIDNO & "' AND TICKET_STATE='1' AND TransToTIMS='N' ORDER BY CREATE_DATE"
    '        AGTDataT=DbAccess.GetDataTable(sql, objconn)
    '        For Each AGTDataR In AGTDataT.Rows
    '            '將身分證號碼和班級代號先帶入欄位中，以便之後檢查
    '            If check_idno.Value <> AGTDataR.Item("IDNO").ToString Then
    '                check_idno.Value=AGTDataR.Item("IDNO").ToString
    '                If check_idno.Value <> "" Then check_idno.Value=Trim(check_idno.Value)
    '                If check_idno.Value <> "" Then check_idno.Value=UCase(check_idno.Value)
    '                If check_idno.Value <> "" Then check_idno.Value=TIMS.ChangeIDNO(check_idno.Value)
    '            End If

    '            If check_class.Value <> "" Then check_class.Value &= ";"
    '            check_class.Value &= AGTDataR.Item("TRN_CLASS").ToString
    '        Next

    '    End If
    'End Sub
#End Region

    '檢查報名者的課程限制資格是符合
    Function Chk_CLASS(ByVal iNum As Integer, ByRef sErrMsg As String) As Boolean
        Dim rst As Boolean = True 'True:符合 (預設無資料) FALSE:不符合
        aNow = TIMS.GetSysDateNow(objconn)
        'Dim iDis As Integer=0 '0:符合 1:不符合
        Dim sPlanID As String = ""
        Dim sComIDNO As String = ""
        Dim sSeqNO As String = ""

        Select Case iNum
            Case 1
                Dim drCX As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
                If drCX Is Nothing Then
                    sErrMsg &= String.Concat(TMID1.Text, "[", OCID1.Text, "]報名資格不符合!") & vbCrLf
                    Return False
                End If
                If flgROLEIDx0xLIDx0 Then
                    sPlanID = Convert.ToString(drCX("PlanID"))
                    sComIDNO = ComIDNO1.Value
                    sSeqNO = SeqNO1.Value
                End If
                If Not flgROLEIDx0xLIDx0 Then
                    If ComIDNO1.Value <> "" AndAlso SeqNO1.Value <> "" Then
                        sPlanID = Convert.ToString(sm.UserInfo.PlanID)
                        sComIDNO = ComIDNO1.Value
                        sSeqNO = SeqNO1.Value
                    End If
                End If
                Dim drP As DataRow = TIMS.GetPCSDate(sPlanID, sComIDNO, sSeqNO, objconn)
                If drP Is Nothing Then
                    sErrMsg &= String.Concat(TMID1.Text, "[", OCID1.Text, "]報名資格不符合!!") & vbCrLf
                    Return False
                End If
                '年齡檢查
                'Dim iAge As Integer=0
                birthday.Text = TIMS.Cdate3(birthday.Text)
                If Convert.ToString(drP("CapAge1")) <> "" Then
                    '依生日加了限制年齡 
                    Dim LastSTDate As String = TIMS.Cdate3(DateAdd(DateInterval.Year, Val(drP("CapAge1")), CDate(birthday.Text)))
                    '未超過開訓日ok /超過不ok
                    If DateDiff(DateInterval.Day, CDate(LastSTDate), CDate(drCX("STDate"))) < 0 Then
                        sErrMsg &= "年齡不符合此課程報名資格的最小年齡要求" & vbCrLf
                        Return False
                    End If
                End If
                'Case 2
                '    If ComIDNO2.Value <> "" AndAlso SeqNO2.Value <> "" Then
                '        sPlanID=Convert.ToString(sm.UserInfo.PlanID)
                '        sComIDNO=ComIDNO2.Value
                '        sSeqNO=SeqNO2.Value
                '    End If
                '    Dim drP As DataRow=TIMS.GetPCSDate(sPlanID, sComIDNO, sSeqNO, objconn)
                '    If drP Is Nothing Then
                '        sErrMsg &= TMID2.Text & "[" & OCID2.Text & "]報名資格不符合" & vbCrLf
                '        Return False
                '    End If
                'Case 3
                '    '應該沒什麼用
                '    If ComIDNO3.Value <> "" AndAlso SeqNO3.Value <> "" Then
                '        sPlanID=Convert.ToString(sm.UserInfo.PlanID)
                '        sComIDNO=ComIDNO3.Value
                '        sSeqNO=SeqNO3.Value
                '    End If
                '    Dim drP As DataRow=TIMS.GetPCSDate(sPlanID, sComIDNO, sSeqNO, objconn)
                '    If drP Is Nothing Then
                '        sErrMsg &= TMID3.Text & "[" & OCID3.Text & "]報名資格不符合" & vbCrLf
                '        Return False
                '    End If
        End Select

        '有發現錯誤
        If sErrMsg <> "" Then rst = False
        Return rst
    End Function

    '送出前檢核 'SERVER端 檢查
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        '報名登錄: 所有規則都要通過 含4項不得報名/4項不予錄訓
        '專案核定報名登錄: 卡4項不得報名/4項不予錄訓 (前3項)
        '特例專案核定報名登錄: 不檢查規則
        Errmsg = ""
        rqProecess = TIMS.ClearSQM(Request("proecess")) 'add/shift/edit
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '產投不可使用 報名登錄 功能。
            Errmsg += "該計畫不提供使用此功能!!" & vbCrLf
            Return False
        End If

        Select Case Convert.ToString(Request("ID"))
            Case TIMS.cst_FunID_報名登錄 '報名登錄(SD_01_001_add)
            Case TIMS.cst_FunID_專案核定報名登錄 '專案核定報名登錄
            Case TIMS.cst_FunID_特例專案核定報名登錄 '專案核定報名登錄
            Case Else
                'If Not blnTestflag Then
                '    Errmsg += "基本輸入參數有誤，請重新選擇功能查詢!!" & vbCrLf
                '    Return False
                'End If
                Errmsg += "基本輸入參數有誤，請重新選擇功能查詢!!" & vbCrLf
                Return False
        End Select

        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        'OCIDValue2.Value=TIMS.ClearSQM(OCIDValue2.Value)
        'OCIDValue3.Value=TIMS.ClearSQM(OCIDValue3.Value)
        If OCIDValue1.Value = "" Then
            Errmsg += "報名班級 職類： 不可為空，請重新輸入。" & vbCrLf
            Return False
        End If
        Dim drOCID1 As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If OCIDValue1.Value.Trim <> "" Then
            If drOCID1 Is Nothing Then
                Errmsg += "報名班級 職類： 不可為空，請重新輸入。" & vbCrLf
                'Errmsg += "報名志願 志願1：班級資料有誤，請改選其他班級報名。" & vbCrLf
                Return False
            End If
            Dim flag1 As Boolean = True '要檢驗
            Select Case Convert.ToString(Request("ID"))
                Case TIMS.cst_FunID_報名登錄 '報名登錄(SD_01_001_add)
                Case TIMS.cst_FunID_專案核定報名登錄 '專案核定報名登錄
                Case TIMS.cst_FunID_特例專案核定報名登錄 '專案核定報名登錄
                    flag1 = False '停止檢驗
            End Select
            '超級使用者
            If flgROLEIDx0xLIDx0 Then flag1 = False '停止詳細檢驗/使用大方向檢驗
            If flag1 Then
                If Convert.ToString(drOCID1("PlanID")) <> Convert.ToString(sm.UserInfo.PlanID) Then Errmsg += "報名班級 與登入計畫不同不可報名。" & vbCrLf
                If Convert.ToString(drOCID1("RID")) <> Convert.ToString(sm.UserInfo.RID) Then Errmsg += "報名班級 與登入業務權限不同不可報名。" & vbCrLf
            Else
                Dim drP1 As DataRow = TIMS.GetPlanID1(Convert.ToString(drOCID1("PlanID")), objconn)
                If Convert.ToString(drP1("TPlanID")) <> Convert.ToString(sm.UserInfo.TPlanID) Then Errmsg += "報名班級 與登入計畫不同不可報名。" & vbCrLf
                If Convert.ToString(drP1("Years")) <> Convert.ToString(sm.UserInfo.Years) Then Errmsg += "報名班級 與登入年度不同不可報名。" & vbCrLf
            End If
            If Errmsg <> "" Then Return False
        End If

        Select Case Convert.ToString(Request("ID"))
            Case TIMS.cst_FunID_特例專案核定報名登錄 '專案核定報名登錄
            Case Else
                'If OCIDValue1.Value <> "" Then
                '    Dim drOCID As DataRow=TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
                '    If drOCID Is Nothing Then
                '        Errmsg += "報名志願 志願一：職類： 不可為空，請重新輸入。" & vbCrLf
                '        'Errmsg += "報名志願 志願1：班級資料有誤，請改選其他班級報名。" & vbCrLf
                '        Return False
                '    End If
                'End If


                If drOCID1 Is Nothing Then
                    Errmsg += "報名班級 職類： 不可為空，請重新輸入。" & vbCrLf
                    'Errmsg += "報名志願 志願1：班級資料有誤，請改選其他班級報名。" & vbCrLf
                    Return False
                End If
                '非系統管理者(超過開訓日)
                Dim flagS1 As Boolean = TIMS.IsSuperUser(Me, 1)
                'https://jira.turbotech.com.tw/browse/TIMSC-161
                If Not flagS1 Then
                    '暫時權限Table
                    Dim dtArc As DataTable = TIMS.Get_Auth_REndClass(Me, objconn)
                    '(超過開訓日)(未超過14日)
                    If TIMS.ChkIsEndDate(OCIDValue1.Value, TIMS.cst_FunID_報名登錄, dtArc) Then
                        '過了使用期限 True(不可使用)   False(可使用)
                        If $"{drOCID1("InputOK")}" <> "Y" Then
                            Errmsg += "報名班級 已過開訓日，鎖定功能填寫。" & vbCrLf
                            Return False
                        End If
                    End If
                End If
                If $"{drOCID1("InputOK14")}" = "" Then
                    '請在報名資料新增儲存時，檢核選擇報名的班級，若已過報名班級的開訓14日，則
                    '不可新增該筆報名資料，請使用者改選其他班級報名。
                    If Not blnTestflag Then
                        Errmsg += "已過報名班級的開訓14日，不可新增該筆報名資料，請改選其他班級報名。" & vbCrLf
                        Return False
                    End If
                    'Errmsg += "已過報名班級的開訓14日，不可新增該筆報名資料，請改選其他班級報名。" & vbCrLf
                    'Return False
                End If
                '1.新增儲存時，判斷是否過選擇報名班級的最晚報名時間，若已過， '則不可儲存，顯示"已過報名班級的最晚報名時間，不可報名此班級!"。
                '2.若是修改已存在之資料時，報名班級項目不提供修改，且不需判斷該班的最晚報名時間。
                If rqProecess = "add" OrElse rqProecess = "shift" Then
                    If TIMS.Cst_TPlanID07.IndexOf(sm.UserInfo.TPlanID) = -1 Then
                        If Convert.ToString(drOCID1("FENTERDATE2")) <> "" AndAlso Convert.ToString(drOCID1("FENTERDATE2Y")) <> "Y" Then
                            If Not blnTestflag Then
                                Errmsg += "已過報名班級的最晚報名時間，不可報名此班級!" & vbCrLf
                                Return False
                            End If
                        End If
                    End If
                End If
        End Select

        'If OCIDValue2.Value.Trim <> "" Then
        '    Dim drOCID2 As DataRow=TIMS.GetOCIDDate(OCIDValue2.Value, objconn)
        '    If Convert.ToString(drOCID2("InputOK14"))="" Then
        '        '請在報名資料新增儲存時，檢核選擇報名的班級，若已過報名班級的開訓14日，則
        '        '不可新增該筆報名資料，請使用者改選其他班級報名。
        '        Errmsg += "報名志願 志願2：已過報名班級的開訓14日，不可新增該筆報名資料，請改選其他班級報名。" & vbCrLf
        '        Return False
        '    End If
        'End If

        'If OCIDValue3.Value.Trim <> "" Then
        '    Dim drOCID3 As DataRow=TIMS.GetOCIDDate(OCIDValue3.Value, objconn)
        '    If Convert.ToString(drOCID3("InputOK14"))="" Then
        '        '請在報名資料新增儲存時，檢核選擇報名的班級，若已過報名班級的開訓14日，則
        '        '不可新增該筆報名資料，請使用者改選其他班級報名。
        '        Errmsg += "報名志願 志願3：已過報名班級的開訓14日，不可新增該筆報名資料，請改選其他班級報名。" & vbCrLf
        '        Return False
        '    End If
        'End If

        IDNO.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(IDNO.Text))
        If IDNO.Text = "" Then Errmsg += "錯誤的身分證號碼，請重新輸入。" & vbCrLf

        Select Case PassPortNO.SelectedValue
            Case "1"
                '1:國民身分證 2:居留證 4:居留證2021
                Dim rqIDNO As String = IDNO.Text
                Dim flag1 As Boolean = TIMS.CheckIDNO(rqIDNO)
                Dim flag2 As Boolean = TIMS.CheckIDNO2(rqIDNO, 2)
                Dim flag4 As Boolean = TIMS.CheckIDNO2(rqIDNO, 4)
                If Not flag1 AndAlso Not flag2 AndAlso Not flag4 Then
                    Errmsg += "錯誤的身分證號碼，請重新輸入。" & vbCrLf
                End If
                'If Not TIMS.CheckIDNO(IDNO.Text) Then Errmsg += "錯誤的身分證號碼，請重新輸入。" & vbCrLf
            Case "2"
                'If Not TIMS.CheckIDNO2(IDNO.Text) Then
                '    Errmsg += "身分證號碼有誤(應為正確的居留證號碼)，請重新輸入。" & vbCrLf
                'End If
            Case Else
                Errmsg += "未選擇身分別，請重新選擇。" & vbCrLf
        End Select

        If notes.Text <> "" Then notes.Text = Trim(notes.Text)
        If notes.Text <> "" Then
            If Len(notes.Text) > 500 Then Errmsg += "備註 長度超過系統範圍(500)" & vbCrLf
            'Else 'notes.Text="" 'Errmsg += "請輸入 其他建議" & vbCrLf
        End If

        '職前課程邏輯
        'ActNo.Text=TIMS.ClearSQM(ActNo.Text)
        ''If ActNo.Text <> "" Then ActNo.Text=Trim(ActNo.Text)
        'If ActNo.Text <> "" Then
        '    If Len(ActNo.Text) > 9 Then Errmsg += "最後投保單位保險證號 長度超過系統範圍(9)" & vbCrLf
        '    If Not TIMS.CheckABC123(ActNo.Text) Then Errmsg += "最後投保單位保險證號 只可輸入英數字" & vbCrLf
        '    'Else ActNo.Text="" 'Errmsg += "請輸入 其他建議" & vbCrLf
        'End If
        If Errmsg = "" Then
            '足歲45歲則選取中高齡者(45歲)。跟開訓日期比較
            Call Over_45YearOld(OCIDValue1.Value, birthday.Text, IdentityID, Errmsg, objconn)
        End If

        School.Text = TIMS.ClearSQM(School.Text)
        Department.Text = TIMS.ClearSQM(Department.Text)
        If School.Text = "" Then Errmsg += "請填寫學校名稱。" & vbCrLf
        If Department.Text = "" Then Errmsg += "請填寫科系名稱。" & vbCrLf

        'If Email.Text <> "" Then Email.Text=Trim(Email.Text)
        Email.Text = TIMS.ClearSQM(Email.Text)
        If Email.Text <> "" AndAlso Email.Text <> "無" Then
            If Not TIMS.CheckEmail(Email.Text) Then Errmsg += "電子信箱 EMail格式錯誤。" & vbCrLf
        End If
        ZipCODEB3.Value = TIMS.ClearSQM(ZipCODEB3.Value)
        Call TIMS.CheckZipCODEB3(ZipCODEB3.Value, "聯絡地址郵遞區號後2碼或後3碼", False, Errmsg)

        '獲得職訓 課程管道
        'If trAVTCP.Visible Then
        'End If
        If Errmsg <> "" Then Return False

        '(職前課程邏輯)若為下列計畫, 則依4項不予錄訓規定設定邏輯判斷學員是否可參訓:
        ' https://jira.turbotech.com.tw/browse/TIMSC-142
        ' 呼叫 TIMS.Get_ChkIsJobsCounse44() 進行檢查
        'Select Case Convert.ToString(Request("ID"))
        '    Case TIMS.cst_FunID_專案核定報名登錄 '專案核定報名登錄
        '        '依4項不予錄訓規定設定邏輯判斷學員是否可參訓
        '        Dim htSS As New Hashtable 'htSS Hashtable() '
        '        htSS.Add("IDNOt", IDNO.Text)
        '        htSS.Add("OCIDVal", OCIDValue1.Value)
        '        htSS.Add("SENTERDATE", TIMS.cdate3(drOCID1("SENTERDATE")))
        '        Errmsg &= TIMS.Get_ChkIsJobsCounse44(Me, htSS, TIMS.cst_FunID_專案核定報名登錄, objconn)
        '    Case TIMS.cst_FunID_特例專案核定報名登錄 '專案核定報名登錄
        '    Case Else
        '        'Case cst_funid報名登錄 '報名登錄(SD_01_001_add)
        '        '依4項不予錄訓規定設定邏輯判斷學員是否可參訓
        '        Dim htSS As New Hashtable 'htSS Hashtable() '
        '        htSS.Add("IDNOt", IDNO.Text)
        '        htSS.Add("OCIDVal", OCIDValue1.Value)
        '        htSS.Add("SENTERDATE", TIMS.cdate3(drOCID1("SENTERDATE")))
        '        Errmsg &= TIMS.Get_ChkIsJobsCounse44(Me, htSS, TIMS.cst_FunID_報名登錄, objconn)
        'End Select

        ZipCode2.Value = TIMS.ClearSQM(ZipCode2.Value)
        ZipCode2_B3.Value = TIMS.ClearSQM(ZipCode2_B3.Value) '後2碼或後3碼
        HidZipCode2_6W.Value = TIMS.ClearSQM(HidZipCode2_6W.Value)
        HouseholdAddress.Text = TIMS.ClearSQM(HouseholdAddress.Text)
        If (ZipCode2.Value = "") Then Errmsg &= "請輸入戶籍地址郵遞區號前3碼。" & vbCrLf
        If (ZipCode2_B3.Value = "") Then Errmsg &= "請輸入戶籍地址郵遞區號後2碼或後3碼。" & vbCrLf
        If (HouseholdAddress.Text = "") Then Errmsg &= "請輸入服務單位。" & vbCrLf

        Uname.Text = TIMS.ClearSQM(Uname.Text)
        Dim v_ddlSERVDEPTID As String = TIMS.GetListValue(ddlSERVDEPTID)
        ActName.Text = TIMS.ClearSQM(ActName.Text)
        Dim v_ActType As String = TIMS.GetListValue(ActType)
        Dim v_ddlJOBTITLEID As String = TIMS.GetListValue(ddlJOBTITLEID)

        If (Uname.Text = "") Then Errmsg &= "請輸入服務單位。" & vbCrLf
        If (v_ddlSERVDEPTID = "") Then Errmsg &= "請選擇服務部門。" & vbCrLf
        If (ActName.Text = "") Then Errmsg &= "請輸入投保單位名稱。" & vbCrLf
        If (v_ActType = "") Then Errmsg &= "請選擇投保類別。" & vbCrLf
        If (v_ddlJOBTITLEID = "") Then Errmsg &= "請選擇職稱。" & vbCrLf

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    '儲存
    Sub SSaveData1()
        'SAVE
        'If TestStr="AmuTest" Then Exit Sub '測試用
        Dim rqIDNO As String = TIMS.ChangeIDNO(TIMS.ClearSQM(Request("IDNO")))
        IDNO.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(IDNO.Text))

        Dim ErrorFlag1 As Boolean = False '有異常為true
        If IDNO.Text = "" Then ErrorFlag1 = True
        If rqIDNO = "" Then ErrorFlag1 = True
        If rqIDNO <> IDNO.Text Then ErrorFlag1 = True
        If ErrorFlag1 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg3)
            Exit Sub
        End If

        Dim sql As String = ""
        Dim fg_eData As Boolean = False   '用來判斷是否有E網報名要通過
        Dim chkMsg As String = ""
        Dim strErrmsg As String = "" '偵錯用儲存欄
        Dim strfield As String = "" '偵錯用儲存欄
        Dim strSql2 As String = "" '偵錯用儲存欄

        Dim drCC1 As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        Select Case ptype.Value 'rqProecess 'Request("proecess")
            Case "add" '新增時特別檢查
                chkMsg = ""
                Dim ss As String = ""
                TIMS.SetMyValue(ss, "tmpIDNO", IDNO.Text)
                TIMS.SetMyValue(ss, "tmpOCID1", OCIDValue1.Value)
                'TIMS.SetMyValue(ss, "tmpOCID2", OCIDValue2.Value)
                'TIMS.SetMyValue(ss, "tmpOCID3", OCIDValue3.Value)
                If Not Check_E_ClsTrace(chkMsg, ss, objconn) Then
                    If chkMsg <> "" Then
                        Dim vsScriptStr As String = String.Concat(vbCrLf, chkMsg, vbCrLf)
                        Common.MessageBox(Me, vsScriptStr)
                        Return ' Exit Sub
                    End If
                End If
        End Select
        '檢查報名日期??

        '0--------檢查欄位
        '1--------檢查報名者的課程限制資格是符合
        '2--------取出最大SETID並且準備建立  or  使用的原始SETID
        '3--------檢查廠商是否提供線上報名
        '4--------建立SerNum序號(Stud_EnterTye)
        '5--------檢查之前的志願是否有相衝
        '6--------開始儲存個人資料
        '7--------開始儲存報名資料

        '0--------檢查欄位-Start
        Button7.Disabled = True
        'Button7.Enabled=False

        '(職前邏輯)
        'If ptype.Value="shift" _
        '    AndAlso rqFrom_type="add" Then
        '    'EnterChannel.SelectedValue=4 '1.網;2.現;3.通;4.推
        '    Common.SetListItem(EnterChannel, "4")
        'End If

        Dim sErrMsg As String = ""
        Dim v_MIdentityID As String = TIMS.GetListValue(MIdentityID)
        If v_MIdentityID = "" Then sErrMsg &= "請選擇 主要參訓身分別!" & vbCrLf

        Dim v_IdentityID As String = TIMS.GetCblValue(IdentityID)
        If v_IdentityID = "" Then sErrMsg &= "請選擇 參訓身分別!" & vbCrLf
        'Common.MessageBox(Me, sErrMsg) 'Exit Sub
        'Common.RespWrite(Me, "<script language=javascript>window.alert('" & Errmsg & "')</script>") 'Exit Sub
        '職前課程邏輯
        'Call TIMS.CheckDateErr(SOfficeYM1.Text, "投保單位 加保日期", False, sErrMsg)
        'Call TIMS.CheckDateErr(FOfficeYM1.Text, "投保單位 退保日期", False, sErrMsg)
        If sErrMsg <> "" Then
            Common.MessageBox(Me, sErrMsg)
            Exit Sub
        End If
        '0--------檢查欄位-End

        'Dim conn As SqlConnection=DbAccess.GetConnection()
        'Call TIMS.OpenDbConn(conn)
        'TIMS.TestDbConn(Me, conn, True)
        'conn.Open()
        'Dim mainTrain As SqlTransaction
        Dim cmd As New SqlCommand               '建立Stud_EnterTemp專用command
        Dim cmd9 As New SqlCommand              '建立Stud_EnterTemp2專用command
        Dim cmd2 As New SqlCommand              '建立STUD_ENTERTYPE專用command
        'Dim cmd3 As SqlCommand
        Dim cmd4 As New SqlCommand               '三合一轉換用
        'Dim cmd5 As SqlCommand                  '填入問卷複選1
        'Dim cmd6 As SqlCommand                  '填入問卷複選2
        Dim cmd7 As New SqlCommand  '回填STUD_ENTERTEMP2
        Dim cmd8 As New SqlCommand  '回填STUD_ENTERTYPE2
        Dim dr As DataRow = Nothing
        'Dim dr1 As DataRow

        '檢查如果改變的身分證，是否有存在於資料庫中
        'Dim rqProecess As String=Request("proecess")
        rqProecess = TIMS.ClearSQM(Request("proecess"))
        'Dim ivsExistence As Integer=0 '沒有資料(STUD_ENTERTEMP)
        'Dim ivsSETID As Integer=0 'SETID(報名登錄 學員報名基本資料 STUD_ENTERTEMP)

        ivsExistence = 0 '沒有資料(STUD_ENTERTEMP)
        ivsSETID = 0 'SETID(報名登錄
        'IDNO.Text=TIMS.ChangeIDNO(IDNO.Text) 'If IDNO.Text="" Then Exit Sub
        Dim sParmsID As New Hashtable From {{"IDNO", IDNO.Text}}
        sql = ""
        sql &= " SELECT a.SETID,a.IDNO,a.NAME,a.SEX,a.BIRTHDAY,a.PASSPORTNO,a.MARITALSTATUS" & vbCrLf
        sql &= " ,a.DEGREEID,a.GRADID,a.SCHOOL,a.DEPARTMENT,a.MILITARYID,a.ZIPCODE,a.ZIPCODE6W" & vbCrLf
        sql &= " ,a.ADDRESS,a.PHONE1,a.PHONE2,a.CELLPHONE,a.EMAIL" & vbCrLf
        'sql += " ,dbms_lob.substr(a.NOTES, 4000, 1) NOTES ,a.ISAGREE ,a.LAINFLAG," & vbCrLf
        sql &= " ,a.NOTES,a.ISAGREE ,a.LAINFLAG" & vbCrLf
        sql &= " ,a.MODIFYACCT ,a.MODIFYDATE ,a.ESETID" & vbCrLf
        sql &= " FROM STUD_ENTERTEMP a" & vbCrLf
        sql &= " WHERE a.IDNO=@IDNO"
        dr = DbAccess.GetOneRow(sql, objconn, sParmsID)
        If dr IsNot Nothing Then
            ivsExistence = 1 '有資料(STUD_ENTERTEMP)
            ivsSETID = dr("SETID")
        End If
        If rqProecess = "edit" AndAlso IDNOChange.Value = "1" Then
            If dr IsNot Nothing Then
                Dim Errmsg As String = ""
                Errmsg = "此身分證已存在!"
                Common.RespWrite(Me, "<script language=javascript>window.alert('" & Errmsg & "')</script>")
                Exit Sub
            End If
        End If

        '1--------檢查報名者的課程限制資格是符合-Start
        If CCLID.Value = "" Then
            '插班不檢查資格
            Dim Errmsg As String = ""
            If OCIDValue1.Value <> "" AndAlso Not Chk_CLASS(1, Errmsg) Then
                Common.MessageBox(Me, Errmsg)
                Exit Sub '離開
            End If
            'If OCIDValue2.Value <> "" AndAlso Not chk_class(2, Errmsg) Then
            '    Common.MessageBox(Me, Errmsg)
            '    Exit Sub '離開
            'End If
            'If OCIDValue3.Value <> "" AndAlso Not chk_class(3, Errmsg) Then
            '    Common.MessageBox(Me, Errmsg)
            '    Exit Sub '離開
            'End If
        End If
        '1--------檢查報名者的課程限制資格是符合-End

        '2--------取出最大SETID並且準備建立  or  使用的原始SETID-Start
        If ivsExistence = 0 AndAlso ivsSETID = 0 Then
            Call TIMS.OpenDbConn(objconn)
            ivsSETID = DbAccess.GetNewId(objconn, "STUD_ENTERTEMP_SETID_SEQ,STUD_ENTERTEMP,SETID")
        End If
        '2--------取出最大SETID並且準備建立  or  使用的原始SETID-End

        '3--------檢查廠商是否提供線上報名-Start
        'Dim IsOnLine As Integer=1'提供線上報名
        'If ViewState("Tplan1") <> "" Then
        '    sql="select * from Key_Plan where TPlanID='" & ViewState("Tplan1") & "'"
        '    dr=DbAccess.GetOneRow(sql)
        '    If Not dr Is Nothing Then
        '        If dr("IsOnLine")="Y" Or dr("IsOnLine")="y" Then
        '            IsOnLine=1
        '        Else
        '            IsOnLine=0
        '        End If
        '    End If
        'End If
        '3--------檢查廠商是否提供線上報名-End

        '4--------建立SerNum序號(Stud_EnterTye)-Start
        Dim iSerNum As Integer = 1 '0
        aNow = TIMS.GetSysDateNow(objconn)
        If R_SerNum.Value = "" Then
            '新增'沒有接收值
            If ivsExistence = 0 Then
                iSerNum = 1 '沒有資料(STUD_ENTERTEMP)
            Else
                '有資料(STUD_ENTERTEMP)
                iSerNum = TIMS.GET_ENTERTYPESERNUM(ivsSETID, objconn)
            End If
        Else
            '修改'有接收值
            'Dim rqSerNum As String=Request("SerNum")
            'rqSerNum=TIMS.ClearSQM(rqSerNum )
            iSerNum = Val(TIMS.ClearSQM(R_SerNum.Value)) 'Request("SerNum")
        End If
        '4--------建立SerNum序號(Stud_EnterTye)-End

        '5--------檢查之前的志願是否有相衝-Start
        Dim dt As DataTable
        If ivsExistence <> 0 Then
            'Me.ivsSETID=TIMS.ClearSQM(ivsSETID)
            'If rqProecess="edit" Then
            '    R_serial.Value=Request("serial")
            '    R_EnterDate.Value=Request("EnterDate")
            '    R_SerNum.Value=Request("SerNum")
            '    R_serial.Value=TIMS.ClearSQM(R_serial.Value)
            '    R_EnterDate.Value=TIMS.ClearSQM(R_EnterDate.Value)
            '    R_SerNum.Value=TIMS.ClearSQM(R_SerNum.Value)
            '    sql="SELECT OCID1,OCID2,OCID3 FROM STUD_ENTERTYPE WHERE SETID='" & ivsSETID & "' and EnterDate<> " & TIMS.to_date(R_EnterDate.Value) & " and SerNum<>'" & R_SerNum.Value & "'"
            'Else
            '    sql="SELECT OCID1,OCID2,OCID3 FROM STUD_ENTERTYPE WHERE SETID='" & ivsSETID & "'"
            'End If

            sql = "SELECT OCID1,OCID2,OCID3 FROM STUD_ENTERTYPE WHERE SETID='" & ivsSETID & "' AND OCID1=" & OCIDValue1.Value
            dt = DbAccess.GetDataTable(sql, objconn)
            If dt.Rows.Count > 0 Then
                Dim sAltMsg As String = String.Concat(OCID1.Text, "已報名過該班級!")
                Common.RespWrite(Me, "<script language=javascript>window.alert('" & sAltMsg & "')</script>")
                'Exit Sub
            End If

        End If
        '5--------檢查之前的志願是否有相衝-End

        '5.1------取出准考證號碼最大號-Start
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        Dim NewExamNo As String = ""
        Dim ExamOcid1 As String = OCIDValue1.Value
        Dim ExamPlanID As String = CStr(drCC1("PlanID"))

        ExamNo.Value = TIMS.ClearSQM(ExamNo.Value)
        If ExamNo.Value = "" Then
            '取出准考證號   Start '取出班級的CLASSID +期別 成為准考證編碼的前面的固定碼
            Dim ExamNo1 As String = TIMS.Get_ExamNo1(ExamOcid1, objconn)
            If ExamNo1 = "" OrElse ExamNo1.Length < 6 Then '防呆
                Common.MessageBox(Me, "班級的代號 與期別有誤，請確認班級狀態")
                Exit Sub
            End If
            Dim flgChkExamNo As Boolean = TIMS.Chk_NewExamNOc(ExamPlanID, ExamOcid1, objconn)
            If Not flgChkExamNo Then
                Common.MessageBox(Me, "班級的代號 與計畫不符，請確認班級狀態(取出准考證號)!")
                Exit Sub
            End If
            '准考證號
            'NewExamNo=TIMS.Get_NewExamNOt(ExamPlanID, ExamNo1, ExamOcid1, mainTrain)
            NewExamNo = TIMS.Get_NewExamNOc(ExamPlanID, ExamNo1, ExamOcid1, objconn)
            If NewExamNo = "" Then
                Common.MessageBox(Me, "班級的代號 與計畫不符，請確認班級狀態(取出准考證號)!!")
                Exit Sub
            End If
            '取出准考證號   End
        Else
            NewExamNo = ExamNo.Value
        End If

        '**by Milor 20080506--檢查是否存在網報名資料----start
        'Dim strIDNO As String=TIMS.ClearSQM(Request("IDNO"))
        'If strIDNO <> "" Then strIDNO=Trim(strIDNO)
        'If strIDNO <> "" Then strIDNO=UCase(strIDNO)
        'If strIDNO <> "" Then strIDNO=TIMS.ChangeIDNO(strIDNO)

        '當資料新增或是推介轉入時，進行判斷E網是否有資料。
        '先把E網報名重複的資料取出，符合條件是IDNO與OCID1、OCID2、OCID3，且沒有SETID
        Dim i_myOCID1 As Integer = Val(If(Len(OCIDValue1.Value) = 0, "0", OCIDValue1.Value))
        Dim r_parms As New Hashtable From {{"IDNO", rqIDNO}, {"OCID1", i_myOCID1}}
        'Dim sql As String=""
        Dim Rsql As String = ""
        Rsql &= " SELECT CONVERT(VARCHAR, a.RelEnterDate, 111) REDate" & vbCrLf
        Rsql &= " ,a.*" & vbCrLf
        Rsql &= " FROM STUD_ENTERTYPE2 a" & vbCrLf
        Rsql &= " JOIN STUD_ENTERTEMP2 b ON a.eSETID=b.eSETID" & vbCrLf
        Rsql &= " WHERE b.IDNO=@IDNO and a.OCID1=@OCID1" & vbCrLf
        Dim chkdr As DataRow = DbAccess.GetOneRow(Rsql, objconn, r_parms)

        'sql="select max(SerNum)+1 SerNum FROM STUD_ENTERTYPE2 where eSETID=" & chkdr("eSETID") & " and RelEnterDate=convert(datetime, '" & chkdr("REDate") & "', 111) and SerNum is not NULL"
        'Dim cntSerNum As Integer=DbAccess.GetCount(sql, objconn)
        'Dim newSerNum As String="1"
        'newSerNum="1"
        'If cntSerNum <> 0 Then newSerNum=DbAccess.ExecuteScalar(sql, objconn)
        '當沒有被賦予過序號時從1開始；否則即採用取出的最大值
        Dim newSerNum As String = "1"
        If chkdr IsNot Nothing Then
            '取得新的序號SerNum
            'sql="select max(SerNum)+1 as SerNum from STUD_ENTERTYPE2 where eSETID=" & chkdr("eSETID") & " and convert(RelEnterDate,111)=convert(varchar(10),convert(datetime,'" & chkdr("REDate") & "'),111) and SerNum is not NULL" '★
            Dim pms_ty2 As New Hashtable From {{"eSETID", chkdr("eSETID")}, {"RelEnterDate", chkdr("REDate")}}
            Dim sql_ty2 As String = ""
            sql_ty2 &= " SELECT 1 FROM STUD_ENTERTYPE2"
            sql_ty2 &= " WHERE eSETID=@eSETID AND RelEnterDate=CONVERT(DATETIME, @RelEnterDate, 111) "
            sql_ty2 &= " AND SerNum IS NOT NULL "
            Dim dtTY2 As DataTable = DbAccess.GetDataTable(sql_ty2, objconn, pms_ty2)
            If dtTY2.Rows.Count > 0 Then
                Dim pms_ty3 As New Hashtable From {{"eSETID", chkdr("eSETID")}, {"RelEnterDate", chkdr("REDate")}}
                Dim sql_ty3 As String = ""
                sql_ty3 &= " SELECT MAX(ISNULL(SerNum, 0))+1 SerNum FROM STUD_ENTERTYPE2"
                sql_ty3 &= " WHERE eSETID=@eSETID AND RelEnterDate=CONVERT(DATETIME, @RelEnterDate, 111) "
                sql_ty3 &= " AND SerNum IS NOT NULL "
                newSerNum = DbAccess.ExecuteScalar(sql_ty3, objconn, pms_ty3)
            End If
        End If

        Dim tConn As SqlConnection = DbAccess.GetConnection()
        Dim mainTrain As SqlTransaction = DbAccess.BeginTrans(tConn)
        Dim JSmsgbox As String = ""
        If ptype.Value = "add" OrElse ptype.Value = "shift" Then
            '當在E網報名檢查出有重覆且未審核的報名資料時，直接將E網報名審核通過
            If chkdr IsNot Nothing Then
                '回填E網報名學員資料暫存檔
                Dim sql7 As String = "" & vbCrLf
                sql7 &= " UPDATE STUD_ENTERTEMP2" & vbCrLf
                sql7 &= $" SET SETID={ivsSETID} ,ModifyAcct='{sm.UserInfo.UserID}',ModifyDate=GETDATE()" & vbCrLf
                sql7 &= $" WHERE eSETID={chkdr("eSETID")}"
                cmd7 = New SqlCommand(sql7, tConn, mainTrain)

                '回填E網報名職類檔
                sql = ""
                sql &= " UPDATE STUD_ENTERTYPE2" & vbCrLf
                sql &= " SET SETID=" & ivsSETID & vbCrLf
                sql &= " ,ExamNo='" & NewExamNo & "'" & vbCrLf
                sql &= " ,SerNum='" & newSerNum & "'" & vbCrLf
                'signUpStatus-0:收件完成 1:報名成功 2:報名失敗 3:正取(Key_SelResult) 4:備取 5:未錄取
                sql &= " ,signUpStatus=1" & vbCrLf
                sql &= " ,ModifyAcct='" & sm.UserInfo.UserID & "'" & vbCrLf
                sql &= " ,ModifyDate=GETDATE()" & vbCrLf
                sql &= " ,BudID=NULL" & vbCrLf
                sql &= " ,SupplyID=NULL" & vbCrLf
                sql &= " WHERE eSETID=" & chkdr("eSETID") & vbCrLf
                sql &= " AND eSerNum=" & chkdr("eSerNum") & vbCrLf
                cmd8 = New SqlCommand(sql, tConn, mainTrain)
                '將eData賦予ture，在後面判斷是否須執行E網報名審核通過
                JSmsgbox += "此學員含有E網報名資料，已將此資料通過!\n"
                fg_eData = True

                '當不具備有推介身分時，報名管道改為網路報名
                '，如果是推介身分，則仍維持報名管道為推介
                'EnterChannel: 1.網;2.現;3.通;4.推
                If hide_TrainMode.Value <> "" Then Common.SetListItem(EnterChannel, "4")
                If ptype.Value <> "shift" Then
                    Select Case EnterChannel.SelectedValue
                        Case "4"
                        Case Else
                            Common.SetListItem(EnterChannel, "1")
                    End Select
                End If
                'If ptype.Value <> "shift" Then
                '    'EnterChannel.SelectedValue=1 '1.網;2.現;3.通;4.推
                '    Common.SetListItem(EnterChannel, "1")
                'End If
            End If
        End If
        '**by Milor 20080506----end

        '6--------開始儲存個人資料-Start
        Try

            If ivsExistence = 0 Then
                '新增ivsSETID
                Call SUtl_INSERT_STUD_ENTERTEMP(ivsSETID, tConn, mainTrain)
            Else
                '使用IDNO
                Call SUtl_UPDATE_STUD_ENTERTEMP(tConn, mainTrain)
            End If

        Catch ex As Exception
            DbAccess.RollbackTrans(mainTrain)
            strErrmsg = String.Concat(ex.ToString, TIMS.GetErrorMsg(Me), vbCrLf) '偵錯用儲存欄
            DbAccess.CloseDbConn(tConn)
            Call TIMS.WriteTraceLog(strErrmsg)
            Throw ex
        End Try
        '6--------開始儲存個人資料-End

        '-------------------update  Stud_EnterTemp2 資料表-------------------- start
        Try

            If IDNO.Text <> "" Then Call SUtl_UPDATE_STUD_ENTERTEMP2(tConn, mainTrain)

        Catch ex As Exception
            DbAccess.RollbackTrans(mainTrain)
            strErrmsg = String.Concat(ex.ToString, TIMS.GetErrorMsg(Me), vbCrLf) '偵錯用儲存欄
            DbAccess.CloseDbConn(tConn)
            Call TIMS.WriteTraceLog(strErrmsg)
            Throw ex
        End Try
        '-------------------update  Stud_EnterTemp2 資料表--------------------end 

        '7--------開始儲存報名資料-Start
        Dim rqTRN_UNKEY As String = TIMS.ClearSQM(Request("TRN_UNKEY"))
        Dim rqBIRTH As String = TIMS.ClearSQM(Request("BIRTH"))
        Dim rqTICKET_NO As String = TIMS.ClearSQM(Request("TICKET_NO"))
        Dim rqAPPLY_DATE As String = TIMS.ClearSQM(Request("APPLY_DATE"))
        'rqTRN_UNKEY=TIMS.ClearSQM(rqTRN_UNKEY)
        'rqBIRTH=TIMS.ClearSQM(rqBIRTH)
        'rqTICKET_NO=TIMS.ClearSQM(rqTICKET_NO)
        'rqAPPLY_DATE=TIMS.ClearSQM(rqAPPLY_DATE)
        'rqEnterDate
        If R_EnterDate.Value = "" Then R_EnterDate.Value = CDate(aNow).ToString("yyyy/MM/dd")

        Dim IsOnLine As Integer = 1 '提供線上報名
        If IsOnLine = 1 Then
            Try

                If rqProecess = "edit" Then
                    Call SUtl_UPDATE_STUD_ENTERTYPE(ivsSETID, R_EnterDate.Value, iSerNum, NewExamNo, tConn, mainTrain, drCC1)
                Else
                    Call SUtl_INSERT_STUD_ENTERTYPE(ivsSETID, R_EnterDate.Value, iSerNum, NewExamNo, tConn, mainTrain, drCC1)
                End If
                Call SUtl_UPDATE_STUD_ENTERTRAIN(ivsSETID, R_EnterDate.Value, iSerNum, tConn, mainTrain, drCC1)

            Catch ex As Exception
                DbAccess.RollbackTrans(mainTrain)
                strErrmsg = String.Concat(ex.ToString, TIMS.GetErrorMsg(Me), vbCrLf) '偵錯用儲存欄
                DbAccess.CloseDbConn(tConn)
                Call TIMS.WriteTraceLog(strErrmsg)
                Throw ex
            End Try

            'Dim htPC As New Hashtable
            'htPC.Add("rqIDNO", rqIDNO)
            'htPC.Add("rqTRN_UNKEY", rqTRN_UNKEY)
            'htPC.Add("rqBIRTH", rqBIRTH)
            'htPC.Add("rqTICKET_NO", rqTICKET_NO)
            'htPC.Add("rqAPPLY_DATE", rqAPPLY_DATE)
            'Call GOV_STOP_1(htPC)

            '**by Milor 20080509--有E網報名資料可以被審核通過時，進行E網報名Table的更新----start
            If fg_eData Then
                Try
                    cmd7.ExecuteNonQuery()
                Catch ex As Exception
                    strErrmsg = ex.ToString & vbCrLf '偵錯用儲存欄
                    If mainTrain IsNot Nothing Then mainTrain.Rollback()
                    Call TIMS.CloseDbConn(tConn)
                    'Common.MessageBox(Me, strErrmsg)

                    '偵錯用儲存欄
                    strErrmsg += "/* strSql2: */" & vbCrLf
                    strErrmsg += strSql2 & vbCrLf
                    strErrmsg += "/* field: */" & vbCrLf
                    strErrmsg += strfield & vbCrLf
                    strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
                    'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                    Call TIMS.WriteTraceLog(strErrmsg)
                    Throw ex
                End Try

                Try
                    cmd8.ExecuteNonQuery()
                Catch ex As Exception
                    strErrmsg = ex.ToString & vbCrLf '偵錯用儲存欄
                    If mainTrain IsNot Nothing Then mainTrain.Rollback()
                    Call TIMS.CloseDbConn(tConn)
                    'Common.MessageBox(Me, strErrmsg)

                    '偵錯用儲存欄
                    strErrmsg += "/* strSql2: */" & vbCrLf
                    strErrmsg += strSql2 & vbCrLf
                    strErrmsg += "/* field: */" & vbCrLf
                    strErrmsg += strfield & vbCrLf
                    strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
                    'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                    Call TIMS.WriteTraceLog(strErrmsg)
                    Throw ex
                End Try
            End If
            '**by Milor 20080509----end

            Try
                mainTrain.Commit()
                '寄發E-mail給報名會員
                'If Not conn Is Nothing Then
                '    If conn.State=ConnectionState.Open Then conn.Close()
                'End If
            Catch ex As Exception
                strErrmsg = ex.ToString & vbCrLf '偵錯用儲存欄
                If mainTrain IsNot Nothing Then mainTrain.Rollback()
                Call TIMS.CloseDbConn(tConn)
                'Common.MessageBox(Me, strErrmsg)

                '偵錯用儲存欄
                'strErrmsg += "/* strSql2: */" & vbCrLf
                'strErrmsg += strSql2 & vbCrLf
                'strErrmsg += "/* field: */" & vbCrLf
                'strErrmsg += strfield & vbCrLf
                strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
                'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                Call TIMS.WriteTraceLog(strErrmsg)
                Throw ex
            End Try
            Session.Remove("wish1_date")    '移除第一志願的session所存時間
            TIMS.CloseDbConn(tConn)

            '假如是學習券,產學訓或插班-等,直接進入參訓狀態
            Dim flag_wr_STUD_SELRESULT As Boolean = False 'false:不可 / true:可 '直接進入參訓狀態
            If TIMS.Cst_TPlanID28AppPlan3.IndexOf(sm.UserInfo.TPlanID) > -1 OrElse CCLID.Value <> "" Then flag_wr_STUD_SELRESULT = True
            If TIMS.Cst_TPlanID07.IndexOf(sm.UserInfo.TPlanID) > -1 Then flag_wr_STUD_SELRESULT = True
            If flag_wr_STUD_SELRESULT Then

                'TIMS.to_date(R_EnterDate.Value)
                Dim sParms As New Hashtable From {{"SETID", ivsSETID}, {"ENTERDATE", TIMS.Cdate3(R_EnterDate.Value)}, {"SERNUM", iSerNum}}
                sql = ""
                sql &= " SELECT * FROM STUD_SELRESULT "
                sql &= " WHERE SETID=@SETID AND convert(date,EnterDate)=convert(date,@ENTERDATE) AND SERNUM=@SERNUM"
                dt = DbAccess.GetDataTable(sql, objconn, sParms)

                If dt.Rows.Count = 0 Then
                    Dim iParms As New Hashtable
                    iParms.Add("SETID", ivsSETID)
                    iParms.Add("ENTERDATE", TIMS.Cdate3(R_EnterDate.Value))
                    iParms.Add("SERNUM", iSerNum)
                    iParms.Add("OCID", OCIDValue1.Value)
                    iParms.Add("RID", drCC1("RID")) 'sm.UserInfo.RID
                    iParms.Add("PLANID", drCC1("PlanID")) 'sm.UserInfo.PlanID
                    iParms.Add("MODIFYACCT", sm.UserInfo.UserID)
                    Dim iSql As String = ""
                    iSql = " INSERT INTO STUD_SELRESULT(SETID ,ENTERDATE ,SERNUM ,OCID,SUMOFGRAD,APPLIEDSTATUS,ADMISSION,SELRESULTID,TRNDTYPE,RID,PLANID,MODIFYACCT,MODIFYDATE)" & vbCrLf
                    iSql &= " VALUES(@SETID ,convert(date,@ENTERDATE),@SERNUM ,@OCID,-1,'Y','Y','01',NULL,@RID,@PLANID,@MODIFYACCT,GETDATE())" & vbCrLf
                    DbAccess.ExecuteNonQuery(iSql, objconn, iParms)

                    Dim uParms As New Hashtable From {{"OCID", OCIDValue1.Value}}
                    Dim u_sql_cc As String = ""
                    u_sql_cc &= " UPDATE Class_ClassInfo "
                    u_sql_cc &= " SET IsCalculate='Y' "
                    u_sql_cc &= " WHERE OCID=@OCID AND ISNULL(IsCalculate,'N')='N'"
                    DbAccess.ExecuteNonQuery(u_sql_cc, objconn, uParms)
                End If
            End If

            Select Case ptype.Value
                Case "add"
                    JSmsgbox += "資料新增成功!\n"
                Case "edit"
                    JSmsgbox += "資料更新成功!\n"
                Case "shift"
                    JSmsgbox += "資料轉入成功!\n"
            End Select
        Else
            '寄發E-mail給廠商
            JSmsgbox += "已發送E-mail!\n"
        End If

        Session("_SearchStr") = ViewState("_SearchStr")

        'city_code.Value 'Cst_TV1_TPlanID
        '該民眾，符合「勞動部因應重大災害職業訓練協助計畫」受災者，請其提供證明文件，得免試入訓。
        Dim sZIPCODE As String = city_code.Value
        Hid_MSG1.Value = TIMS.SHOW_ZIP2MSG(Me, sZIPCODE, R_EnterDate.Value, objconn)
        If Hid_MSG1.Value <> "" Then
            Common.RespWrite(Me, "<script language=javascript>window.alert('" & Hid_MSG1.Value & "');</script>")
        End If

        'https://jira.turbotech.com.tw/browse/TIMSC-151
        'sZIPCODE=city_code.Value
        'Dim iADID As Integer=0
        'Dim flagMSG2 As Boolean=TIMS.CHK_DIS2MSG(Me, sZIPCODE, R_EnterDate.Value, objconn, iADID)
        'If flagMSG2 Then
        '    Hid_MSGADIDN.Value=iADID
        '    Hid_MSG2.Value=TIMS.SHOW_DIS2MSG(Me, sZIPCODE, R_EnterDate.Value, iADID, objconn)
        '    If Hid_MSG2.Value <> "" Then Common.RespWrite(Me, "<script language=javascript>window.alert('" & Hid_MSG2.Value & "');</script>")
        '    Dim ss As String=""
        '    Call TIMS.SetMyValue(ss, "ADID", Hid_MSGADIDN.Value)
        '    Call TIMS.SetMyValue(ss, "OCID", OCIDValue1.Value)
        '    Call TIMS.SetMyValue(ss, "IDNO", rqIDNO)
        '    Call TIMS.SetMyValue(ss, "SETID", ivsSETID)
        '    Call TIMS.SUtl_AddDISASTER(Me, ss, objconn)
        'End If

        If JSmsgbox <> "" Then
            Common.RespWrite(Me, "<script language=javascript>window.alert('" & JSmsgbox & "');</script>")
            Common.RespWrite(Me, "<script language=javascript>window.location.href='SD_01_001.aspx?ID=" & Request("ID") & "';</script>")
        End If

        '7--------開始儲存報名資料-End        
        Button7.Disabled = False
        'Button7.Enabled=True
        'Catch ex As Exception
        '    Throw ex
        'End Try
    End Sub

    'Sub GOV_STOP_1(ByRef htPC As Hashtable)
    '    Dim rqIDNO As String= TIMS.GetMyValue2(htPC, "rqIDNO")
    '    Dim rqTRN_UNKEY As String= TIMS.GetMyValue2(htPC, "rqTRN_UNKEY")
    '    Dim rqBIRTH As String= TIMS.GetMyValue2(htPC, "rqBIRTH")
    '    Dim rqTICKET_NO As String= TIMS.GetMyValue2(htPC, "rqTICKET_NO")
    '    Dim rqAPPLY_DATE As String= TIMS.GetMyValue2(htPC, "rqAPPLY_DATE")

    '    Dim sql As String= ""
    '    If ptype.Value= "shift" Then
    '        If rqFrom_type= "add" Then
    '            '如為三合一資料時，則將需要的三合一資料撈出 'by mick
    '            Dim adp_dr As DataRow
    '            sql= "SELECT * FROM Adp_GOVTRNData WHERE IDNO='" & rqIDNO & "' AND TRN_CLASS='" & select_id.Value & "' ORDER BY TICKET_STATE DESC,CREATE_DATE "
    '            adp_dr= DbAccess.GetOneRow(sql, objconn)
    '            If adp_dr IsNot Nothing Then
    '                sql= ""
    '                sql &= " UPDATE Adp_GOVTRNData "
    '                sql &= " SET TransToTIMS='Y' ,TIMSModifyDate=GETDATE() "
    '                sql &= " WHERE 1=1 "
    '                sql &= " AND TRN_UNKEY='" & adp_dr("TRN_UNKEY") & "' "
    '                sql &= " AND IDNO='" & rqIDNO & "' "
    '                sql &= " AND BIRTH=" & TIMS.To_date(TIMS.Cdate3(adp_dr("BIRTH")))
    '                sql &= " AND TICKET_NO='" & adp_dr("TICKET_NO") & "' "
    '                sql &= " AND APPLY_DATE=" & TIMS.To_date(TIMS.Cdate3(adp_dr("APPLY_DATE")))
    '            End If
    '        End If

    '        Select Case Hid_rqTicket.Value 'rqTicket/Request("ticket")
    '            Case "1"
    '                sql= ""
    '                sql &= " UPDATE Adp_TRNData "
    '                sql &= " SET TransToTIMS='Y',TIMSModifyDate=GETDATE() "
    '                sql &= " WHERE 1=1 "
    '                sql &= " AND TRN_UNKEY='" & rqTRN_UNKEY & "' "
    '                sql &= " AND IDNO='" & rqIDNO & "' "
    '                sql &= " AND BIRTH=" & TIMS.To_date(rqBIRTH)
    '                sql &= " AND TICKET_NO='" & rqTICKET_NO & "' "
    '                sql &= " AND APPLY_DATE=" & TIMS.To_date(rqAPPLY_DATE)
    '            Case "2"
    '                Dim rqDIGITRAIN_UNKEY As String= Request("DIGITRAIN_UNKEY")
    '                rqDIGITRAIN_UNKEY= TIMS.ClearSQM(rqDIGITRAIN_UNKEY)
    '                sql= ""
    '                sql &= " UPDATE Adp_DGTRNData "
    '                sql &= " SET TransToTIMS='Y' ,TIMSModifyDate=GETDATE() "
    '                sql &= " WHERE 1=1 "
    '                sql &= " AND DIGITRAIN_UNKEY='" & rqDIGITRAIN_UNKEY & "' "
    '                sql &= " AND IDNO='" & rqIDNO & "' "
    '                sql &= " AND BIRTH=" & TIMS.To_date(rqBIRTH)
    '                sql &= " AND TICKET_NO='" & rqTICKET_NO & "' "
    '                sql &= " AND APPLY_DATE=" & TIMS.To_date(rqAPPLY_DATE)
    '            Case "3"
    '                sql= ""
    '                sql &= " UPDATE ADP_GOVTRNDATA "
    '                sql &= " SET TransToTIMS='Y' ,TIMSModifyDate=GETDATE() "
    '                sql &= " WHERE 1=1 "
    '                sql &= " AND TRN_UNKEY='" & rqTRN_UNKEY & "' "
    '                sql &= " AND IDNO='" & rqIDNO & "' "
    '                sql &= " AND BIRTH=" & TIMS.To_date(rqBIRTH)
    '                sql &= " AND TICKET_NO='" & rqTICKET_NO & "' "
    '                sql &= " AND APPLY_DATE=" & TIMS.To_date(rqAPPLY_DATE)
    '                'JSmsgbox += "此學員含有推介報名資料，已將此資料轉入!\n"
    '        End Select
    '        cmd4= New SqlCommand(sql, tConn, mainTrain)
    '    End If

    '    Try
    '    Catch ex As Exception
    '        Dim strErrmsg As String= ""
    '        strErrmsg= ex.ToString & vbCrLf '偵錯用儲存欄
    '        strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
    '        Call TIMS.WriteTraceLog(strErrmsg)
    '        Throw ex
    '    End Try
    'End Sub

    '重載數據(JS無法讀取有效值)
    'Sub ReLoad_SB4IDx()
    '    IDNO.Text=TIMS.ClearSQM(IDNO.Text)
    '    hidSB4ID.Value=TIMS.ClearSQM(hidSB4ID.Value)
    '    If hidSB4ID.Value="" Then Exit Sub '為空離開
    '    Select Case PriorWorkType1.SelectedValue
    '        Case "1" '曾工作過
    '        Case Else
    '            Exit Sub '未選擇 (曾工作過) 離開
    '    End Select
    '    Dim drSB4ID As DataRow=TIMS.Get_BLIGATEDATA4(hidSB4ID.Value, IDNO.Text, objconn)
    '    If drSB4ID Is Nothing Then
    '        'Common.MessageBox(Me, "查無資料，無法回傳值")
    '        Exit Sub
    '    End If
    '    '任職單位名稱
    '    PriorWorkOrg1.Text=Convert.ToString(drSB4ID("COMNAME"))
    '    '投保單位保險證號
    '    ActNo.Text=Convert.ToString(drSB4ID("actno"))
    'End Sub

    ''' <summary>
    ''' 送出(隱藏)
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim aNow As Date = TIMS.GetSysDateNow(objconn)
        Dim sAltMsg As String = "" '訊息
        Dim flag_stopEnterH4 As Boolean = TIMS.StopEnterTempMsgH4(objconn, sAltMsg)
        If flag_stopEnterH4 Then
            Common.MessageBox(Me, sAltMsg)
            Return 'Exit Sub
        End If

        '重載數據(JS無法讀取有效值)
        'Call ReLoad_SB4IDx()
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        '儲存
        Call SSaveData1()
    End Sub

    '回報名登錄
    Private Sub Button4_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.ServerClick
        Session("_SearchStr") = ViewState("_SearchStr")
        Call TIMS.CloseDbConn(objconn)
        TIMS.Utl_Redirect1(Me, "SD_01_001.aspx?ID=" & Request("ID") & "")
    End Sub

    '檢查e網前台 報名重複資料 MEM_060
    'TIMS SD_01_001_add 報名登錄功能
    'TIMS SD_01_001 報名登錄功能
    Public Shared Function Check_E_ClsTrace(ByRef Errmsg As String, ByVal ss As String, ByVal oConn As SqlConnection) As Boolean
        Dim rst As Boolean = False 'false:異常
        '加入 201008 排除 不開班 (NotOpen='Y')  等同 NotOpen='N'" & vbCrLf '要開班的課程 BY AMU 
        'Check_E_ClsTrace=False

        Dim tmpIDNO As String = TIMS.GetMyValue(ss, "tmpIDNO")
        Dim tmpOCID1 As String = TIMS.GetMyValue(ss, "tmpOCID1")
        'Dim tmpOCID2 As String=TIMS.GetMyValue(ss, "tmpOCID2")
        'Dim tmpOCID3 As String=TIMS.GetMyValue(ss, "tmpOCID3")

        Dim OCIDs As String = "" '準備要報名的班級
        OCIDs = ""
        If tmpOCID1 <> "" Then
            If OCIDs <> "" Then OCIDs &= ","
            OCIDs &= tmpOCID1
        End If
        'If tmpOCID2 <> "" Then
        '    If OCIDs <> "" Then OCIDs &= ","
        '    OCIDs &= tmpOCID2
        'End If
        'If tmpOCID3 <> "" Then
        '    If OCIDs <> "" Then OCIDs &= ","
        '    OCIDs &= tmpOCID3
        'End If

        Dim sql As String = ""
        Dim dt As DataTable
        Dim dr As DataRow
        Dim dt2 As DataTable

        Errmsg = ""
        If OCIDs <> "" Then
            sql = "" & vbCrLf
            sql &= " SELECT i.ORGID ,i.OrgName ,a.PlanID, a.ComIDNO, a.SeqNo, a.RID, a.OCID, a.ExamDate, a.ExamPeriod" & vbCrLf
            sql &= " ,dbo.FN_GET_CLASSCNAME(a.CLASSCNAME,a.CYCLTYPE) CLASSNAME" & vbCrLf
            sql &= " FROM CLASS_CLASSINFO a" & vbCrLf
            sql &= " JOIN ORG_ORGINFO i ON i.ComIDNO=a.ComIDNO" & vbCrLf
            sql &= $" WHERE a.OCID IN ({OCIDs})" & vbCrLf
            dt = DbAccess.GetDataTable(sql, oConn)

            If dt.Rows.Count > 0 Then
                '狀況一(當本次報名課程清單中,有同一培訓單位,且甄試日期為同一天的報名課程),顯示訊息如下
                'For i As Integer=0 To dt.Rows.Count - 1 '本次報名課程，各個課程判斷
                '    dr=dt.Rows(i)
                '    If dr("ExamDate").ToString <> "" And IsDate(dr("ExamDate")) Then
                '        dt2=Nothing
                '        sql="" & vbCrLf
                '        sql &= " 	SELECT  i.ORGID, i.OrgName" & vbCrLf
                '        sql &= "  	,a.PlanID, a.ComIDNO, a.SeqNo, a.RID, a.OCID, a.ExamDate, a.ExamPeriod" & vbCrLf
                '        sql &= " ,dbo.FN_GET_CLASSCNAME(a.CLASSCNAME,a.CYCLTYPE) CLASSNAME" & vbCrLf
                '        sql &= " 	FROM Class_ClassInfo a" & vbCrLf
                '        sql &= " 	join Org_OrgInfo i on i.ComIDNO=a.ComIDNO" & vbCrLf
                '        sql &= " 	WHERE 1=1" & vbCrLf
                '        sql &= " 	AND a.OCID IN (" & OCIDs & ")" & vbCrLf
                '        sql &= " 	and i.ORGID=" & dr("ORGID") & vbCrLf
                '        sql &= " 	and a.ExamDate='" & Common.FormatDate(dr("ExamDate")) & "'" & vbCrLf
                '        Select Case Convert.ToString(dr("ExamPeriod"))
                '            Case "02" '02:上午 (01:全天或02:同上午)
                '                sql &= " 	and a.ExamPeriod IN ('01','02')" & vbCrLf
                '            Case "03" '03:下午 (01:全天或03:同下午)
                '                sql &= " 	and a.ExamPeriod IN ('01','03')" & vbCrLf
                '            Case Else '01:全天 (其他全天)
                '        End Select
                '        dt2=DbAccess.GetDataTable(sql)
                '        If dt2.Rows.Count > 1 Then '大於兩筆,有同一培訓單位,且甄試日期為同一天的報名課程
                '            Errmsg += " 您所選擇由 「" & dt2.Rows(0)("OrgName").ToString & "」 (" & dt2.Rows(0)("ORGID").ToString & ")(培訓單位)\r\n"
                '            Errmsg += " 開訓之課程 \r\n 報名課程1: 「" & dt2.Rows(0)("ClassName").ToString & "」 (" & dt2.Rows(0)("OCID").ToString & ")\r\n"
                '            Errmsg += " 報名課程2: 「" & dt2.Rows(1)("ClassName").ToString & "」 (" & dt2.Rows(1)("OCID").ToString & ")\r\n"
                '            Errmsg += " 甄試作業為同一天舉辦 (" & Common.FormatDate(dr("ExamDate")) & ")，請擇一報名，以確保您後續的權益，謝謝!"
                '            Exit For
                '        End If
                '    End If
                'Next

                ' IDNO@Session("UsrMemberID")@tmpIDNO

                '狀況二(當本次報名課程清單中,與過去已報名課程,為同一培訓單位
                ',且甄試日期為同一天的報名課程,顯示訊息如下

                '您所選擇由XXXXX(訓練單位名稱)開訓之課程XXXXX(課程名稱)
                '，與您日前已完成報名之XXXXX(課程名稱)甄試作業為同一天舉辦
                '，因故無法受理您此次報名需求，請見諒。

                '狀況三(當本次報名課程清單中,與過去已報名課程,為同一培訓單位
                ',且甄試日期為同一天的報名課程,但已報名課程之報名資料被審核失敗,則容許報名本次)
                '(e網審核成功,錄取作業為未選擇,正取與備取者.)
                '請查閱： Get_STUD_ENTERTYPE2_OCIDs
                Dim OCIDs2 As String = ""
                If Errmsg = "" Then
                    'OCIDs2=Get_STUD_ENTERTYPE2_OCIDs(tmpIDNO, oConn) '收件完成(審核中)，報名成功
                    'For i As Integer=0 To dt.Rows.Count - 1
                    '    dr=dt.Rows(i)
                    '    If OCIDs2 <> "" Then
                    '        If dr("ExamDate").ToString <> "" And IsDate(dr("ExamDate")) Then
                    '            sql="" & vbCrLf
                    '            sql &= " 	SELECT  i.ORGID, i.OrgName" & vbCrLf
                    '            sql &= "  	,a.PlanID, a.ComIDNO, a.SeqNo, a.RID, a.OCID, a.ExamDate, a.ExamPeriod" & vbCrLf
                    '            sql &= " ,dbo.FN_GET_CLASSCNAME(a.CLASSCNAME,a.CYCLTYPE) CLASSNAME" & vbCrLf
                    '            sql &= " 	FROM Class_ClassInfo a" & vbCrLf
                    '            sql &= " 	join Org_OrgInfo i on i.ComIDNO=a.ComIDNO" & vbCrLf
                    '            sql &= " 	WHERE 1=1" & vbCrLf
                    '            sql &= " 	AND a.OCID IN (" & OCIDs2 & ")" & vbCrLf
                    '            sql &= " 	and i.ORGID=" & dr("ORGID") & vbCrLf
                    '            sql &= " 	and a.ExamDate='" & Common.FormatDate(dr("ExamDate")) & "'" & vbCrLf
                    '            Select Case Convert.ToString(dr("ExamPeriod"))
                    '                Case "02" '02:上午 (01:全天或02:同上午)
                    '                    sql &= " 	and a.ExamPeriod IN ('01','02')" & vbCrLf
                    '                Case "03" '03:下午 (01:全天或03:同下午)
                    '                    sql &= " 	and a.ExamPeriod IN ('01','03')" & vbCrLf
                    '                Case Else '01:全天 (其他全天)
                    '            End Select
                    '            dt2=DbAccess.GetDataTable(sql)
                    '            If dt2.Rows.Count > 0 Then '大於兩筆,有同一培訓單位,且甄試日期為同一天的報名課程
                    '                Errmsg += " 您所選擇由 「" & dr("OrgName").ToString & "」 (" & dr("ORGID").ToString & ")(培訓單位)\r\n"
                    '                Errmsg += " 開訓之課程 \r\n 報名課程1: 「" & dr("ClassName").ToString & "」 (" & dr("OCID").ToString & ") \r\n"
                    '                Errmsg += " 與您日前已完成報名之 \r\n 報名課程2: 「" & dt2.Rows(0)("ClassName").ToString & "」 (" & dt2.Rows(0)("OCID").ToString & ") \r\n"
                    '                Errmsg += " 甄試作業為同一天舉辦 (" & Common.FormatDate(dr("ExamDate")) & ")，因故無法受理您此次報名需求，請見諒。"
                    '                Exit For
                    '            End If
                    '        End If
                    '    End If
                    'Next


                End If

                '狀況四(當本次報名課程清單中,與過去已報名課程,為同一培訓單位
                ',但已報名課程之報名資料被審核失敗,則容許報名本次，否則不可報名)
                If Errmsg = "" Then
                    If OCIDs2 = "" Then OCIDs2 = Get_STUD_ENTERTYPE2_OCIDs(tmpIDNO, oConn) '收件完成(審核中)，報名成功
                    For i As Integer = 0 To dt.Rows.Count - 1
                        dr = dt.Rows(i) '要報名的班級
                        If OCIDs2 <> "" Then
                            sql = "" & vbCrLf
                            sql &= " SELECT i.ORGID, i.OrgName" & vbCrLf
                            sql &= " ,a.PlanID, a.ComIDNO, a.SeqNo, a.RID, a.OCID, a.ExamDate" & vbCrLf
                            sql &= " ,dbo.FN_GET_CLASSCNAME(a.CLASSCNAME,a.CYCLTYPE) CLASSNAME" & vbCrLf
                            sql &= " FROM Class_ClassInfo a" & vbCrLf
                            sql &= " JOIN Org_OrgInfo i ON i.ComIDNO=a.ComIDNO" & vbCrLf
                            sql &= $" WHERE a.OCID IN ({OCIDs2})" & vbCrLf
                            sql &= " AND i.ORGID=" & dr("ORGID") & vbCrLf
                            sql &= " AND a.OCID='" & dr("OCID").ToString & "'" & vbCrLf
                            dt2 = DbAccess.GetDataTable(sql, oConn)

                            If dt2.Rows.Count > 0 Then '大於兩筆,為同一培訓單位,尚在審核中或報名成功
                                Errmsg = ""
                                Errmsg += " 您所選擇由 「" & dr("OrgName").ToString & "」 (" & dr("ORGID").ToString & ")(培訓單位) <br/> "
                                Errmsg += " 開訓之課程 <br/> 報名課程1: 「" & dr("ClassName").ToString & "」 (" & dr("OCID").ToString & ") <br/> "
                                Errmsg += " 與您日前已完成報名之 <br/> 報名課程2: 「" & dt2.Rows(0)("ClassName").ToString & "」 (" & dt2.Rows(0)("OCID").ToString & ") <br/> "
                                Errmsg += " 已在e網有報名資料，尚在審核中或報名成功，因故無法受理您此次報名需求，請見諒。"
                                Exit For
                            End If
                        End If
                    Next
                End If

                '同狀況三
                '狀況四-2(當本次報名課程清單中,與過去已報名課程,為同一培訓單位(甄試作業為同一天舉辦)
                ',但已報名課程之報名資料被審核失敗,則容許報名本次，否則不可報名)
                'If Errmsg="" Then
                '    If OCIDs2="" Then
                '        OCIDs2=Get_STUD_ENTERTYPE2_OCIDs(TMPIDNO) '收件完成(審核中)，報名成功
                '    End If
                '    For i As Integer=0 To dt.Rows.Count - 1
                '        dr=dt.Rows(i) '要報名的班級
                '        If OCIDs2 <> "" Then
                '            If dr("ExamDate").ToString <> "" And IsDate(dr("ExamDate")) Then
                '                sql="" & vbCrLf
                '                sql &= " 	SELECT  i.ORGID, i.OrgName" & vbCrLf
                '                sql &= "  	,a.PlanID, a.ComIDNO, a.SeqNo, a.RID, a.OCID, a.ExamDate" & vbCrLf
                '                sql &= " ,dbo.FN_GET_CLASSCNAME(a.CLASSCNAME,a.CYCLTYPE) CLASSNAME" & vbCrLf
                '                sql &= " 	FROM Class_ClassInfo a" & vbCrLf
                '                sql &= " 	join Org_OrgInfo i on i.ComIDNO=a.ComIDNO" & vbCrLf
                '                sql &= " 	WHERE 1=1" & vbCrLf
                '                sql &= " 	AND a.OCID IN (" & OCIDs2 & ")" & vbCrLf
                '                sql &= " 	and i.ORGID=" & dr("ORGID") & vbCrLf
                '                'sql += " 	AND a.OCID='" & dr("OCID").ToString & "'" & vbCrLf
                '                sql &= " 	and a.ExamDate='" & Common.FormatDate(dr("ExamDate")) & "'" & vbCrLf
                '                dt2=DbAccess.GetDataTable(sql)
                '                If dt2.Rows.Count > 0 Then '大於兩筆,為同一培訓單位,尚在審核中或報名成功
                '                    Errmsg += " 您所選擇由 「" & dr("OrgName").ToString & "」 (" & dr("ORGID").ToString & ")(培訓單位)\r\n"
                '                    Errmsg += " 開訓之課程 \r\n 報名課程1: 「" & dr("ClassName").ToString & "」 (" & dr("OCID").ToString & ") \r\n"
                '                    Errmsg += " 與您日前已完成報名之 \r\n 報名課程2: 「" & dt2.Rows(0)("ClassName").ToString & "」 (" & dt2.Rows(0)("OCID").ToString & ") \r\n"
                '                    Errmsg += " 甄試作業為同一天舉辦 (" & Common.FormatDate(dr("ExamDate")) & ") "
                '                    Errmsg += " 已在e網有報名資料，尚在審核中或報名成功，因故無法受理您此次報名需求，請見諒。"
                '                    Exit For
                '                End If
                '            End If
                '        End If
                '    Next
                'End If



                '狀況五(已在內網有報名資料)
                Dim OCIDs1 As String = ""
                OCIDs1 = ""
                If Errmsg = "" Then
                    OCIDs1 = Get_STUD_ENTERTYPE_OCIDs(tmpIDNO, oConn) '已在內網有報名資料
                    For i As Integer = 0 To dt.Rows.Count - 1
                        dr = dt.Rows(i) '要報名的班級
                        If OCIDs1 <> "" Then
                            sql = "" & vbCrLf
                            sql &= " SELECT i.ORGID, i.OrgName" & vbCrLf
                            sql &= " ,a.PlanID, a.ComIDNO, a.SeqNo, a.RID, a.OCID, a.ExamDate" & vbCrLf
                            sql &= " ,dbo.FN_GET_CLASSCNAME(a.CLASSCNAME,a.CYCLTYPE) CLASSNAME" & vbCrLf
                            sql &= " FROM Class_ClassInfo a" & vbCrLf
                            sql &= " JOIN Org_OrgInfo i ON i.ComIDNO=a.ComIDNO" & vbCrLf
                            sql &= $" WHERE a.OCID IN ({OCIDs1})" & vbCrLf
                            sql &= " AND i.ORGID=" & dr("ORGID") & vbCrLf
                            sql &= " AND a.OCID='" & dr("OCID").ToString & "'" & vbCrLf
                            dt2 = DbAccess.GetDataTable(sql, oConn)

                            If dt2.Rows.Count > 0 Then '大於兩筆,為同一培訓單位,已在內網有報名資料
                                Errmsg += " 您所選擇由 「" & dr("OrgName").ToString & "」 (" & dr("ORGID").ToString & ")(培訓單位) <br/> "
                                Errmsg += " 開訓之課程 <br/> 報名課程1: 「" & dr("ClassName").ToString & "」 (" & dr("OCID").ToString & ") <br/> "
                                Errmsg += " 與您日前已完成報名之 <br/> 報名課程2: 「" & dt2.Rows(0)("ClassName").ToString & "」 (" & dt2.Rows(0)("OCID").ToString & ") <br/> "
                                Errmsg += " 已在內網有報名資料，尚在審核中或報名成功，因故無法受理您此次報名需求，請見諒。"
                                Exit For
                            End If
                        End If
                    Next
                End If

                '狀況五-2(已在內網有報名資料)(甄試作業為同一天舉辦)
                If Errmsg = "" Then
                    'If OCIDs1="" Then OCIDs1=Get_STUD_ENTERTYPE_OCIDs(tmpIDNO, oConn) '已在內網有報名資料
                    'For i As Integer=0 To dt.Rows.Count - 1
                    '    dr=dt.Rows(i) '要報名的班級
                    '    If OCIDs1 <> "" Then
                    '        If dr("ExamDate").ToString <> "" And IsDate(dr("ExamDate")) Then
                    '            sql="" & vbCrLf
                    '            sql &= " 	SELECT  i.ORGID, i.OrgName" & vbCrLf
                    '            sql &= "  	,a.PlanID, a.ComIDNO, a.SeqNo, a.RID, a.OCID, a.ExamDate" & vbCrLf
                    '            sql &= " ,dbo.FN_GET_CLASSCNAME(a.CLASSCNAME,a.CYCLTYPE) CLASSNAME" & vbCrLf
                    '            sql &= " 	FROM Class_ClassInfo a" & vbCrLf
                    '            sql &= " 	join Org_OrgInfo i on i.ComIDNO=a.ComIDNO" & vbCrLf
                    '            sql &= " 	WHERE 1=1" & vbCrLf
                    '            sql &= " 	AND a.OCID IN (" & OCIDs1 & ")" & vbCrLf
                    '            sql &= " 	and i.ORGID=" & dr("ORGID") & vbCrLf
                    '            'sql += " 	AND a.OCID='" & dr("OCID").ToString & "'" & vbCrLf
                    '            sql &= " 	and a.ExamDate='" & Common.FormatDate(dr("ExamDate")) & "'" & vbCrLf
                    '            Select Case Convert.ToString(dr("ExamPeriod"))
                    '                Case "02" '02:上午 (01:全天或02:同上午)
                    '                    sql &= " 	and a.ExamPeriod IN ('01','02')" & vbCrLf
                    '                Case "03" '03:下午 (01:全天或03:同下午)
                    '                    sql &= " 	and a.ExamPeriod IN ('01','03')" & vbCrLf
                    '                Case Else '01:全天 (其他全天)
                    '            End Select
                    '            dt2=DbAccess.GetDataTable(sql)
                    '            If dt2.Rows.Count > 0 Then '大於兩筆,為同一培訓單位,已在內網有報名資料
                    '                Errmsg += " 您所選擇由 「" & dr("OrgName").ToString & "」 (" & dr("ORGID").ToString & ")(培訓單位)\r\n"
                    '                Errmsg += " 開訓之課程 \r\n 報名課程1: 「" & dr("ClassName").ToString & "」 (" & dr("OCID").ToString & ") \r\n"
                    '                Errmsg += " 與您日前已完成報名之 \r\n 報名課程2: 「" & dt2.Rows(0)("ClassName").ToString & "」 (" & dt2.Rows(0)("OCID").ToString & ") \r\n"
                    '                Errmsg += " 甄試作業為同一天舉辦 (" & Common.FormatDate(dr("ExamDate")) & ") "
                    '                Errmsg += " 已在內網有報名資料，尚在審核中或報名成功，因故無法受理您此次報名需求，請見諒。"
                    '                Exit For
                    '            End If
                    '        End If
                    '    End If
                    'Next
                End If

            End If
        End If

        If Errmsg <> "" Then Replace(Errmsg, "\r\n", vbCrLf)
        If Errmsg <> "" Then rst = False
        If Errmsg = "" Then rst = True
        Return rst
        'Check_E_ClsTrace=True
    End Function

    Public Shared Function Get_STUD_ENTERTYPE2_OCIDs(ByVal IDNO As String, ByVal oConn As SqlConnection) As String
        '狀況三(當本次報名課程清單中,與過去已報名課程,為同一培訓單位
        ',且甄試日期為同一天的報名課程,但已報名課程之報名資料被審核失敗,則容許報名本次)
        '(e網審核成功,錄取作業為 0:未選擇, 1:正取, 4:備取者.) 排除：(NOT IN) 2:報名失敗/5:未錄取
        'signUpStatus: 2:報名失敗/5:未錄取
        'signUpStatus-0:收件完成 1:報名成功 2:報名失敗 3:正取(Key_SelResult) 4:備取 5:未錄取
        Dim OCIDs As String = ""
        Dim sql As String
        sql = "" & vbCrLf
        'sql += "  SELECT DISTINCT b.ocid1 OCID" & vbCrLf
        sql &= "  SELECT b.OCID1 OCID" & vbCrLf
        sql &= "  FROM STUD_ENTERTEMP2 a" & vbCrLf
        sql &= "  JOIN STUD_ENTERTYPE2 b ON a.eSETID=b.eSETID AND b.signUpStatus NOT IN (2,5)" & vbCrLf
        sql &= "  WHERE a.IDNO='" & IDNO & "'" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, oConn)
        For Each dr1 As DataRow In dt.Rows
            If OCIDs.IndexOf("'" & Convert.ToString(dr1("OCID")) & "'") = -1 Then
                If OCIDs <> "" Then OCIDs &= ","
                OCIDs &= "'" & Convert.ToString(dr1("OCID")) & "'"
            End If
        Next
        Return OCIDs
    End Function

    Public Shared Function Get_STUD_ENTERTYPE_OCIDs(ByVal IDNO As String, ByVal oConn As SqlConnection) As String
        '狀況五(已在內網有報名資料)
        Dim OCIDs As String = ""
        Dim sql As String
        sql = "" & vbCrLf
        sql &= " SELECT DISTINCT b.OCID1 OCID" & vbCrLf
        sql &= " FROM Stud_EnterTemp a" & vbCrLf
        sql &= " JOIN STUD_ENTERTYPE b ON a.SETID=b.SETID" & vbCrLf
        sql &= " LEFT JOIN STUD_SELRESULT f ON b.SETID=f.SETID AND b.EnterDate=f.EnterDate AND b.SerNum=f.SerNum" & vbCrLf
        sql &= " WHERE a.IDNO='" & IDNO & "'" & vbCrLf
        'f.Admission 是否錄取 (N:不通過, Y:通過, null:尚未審核、審核中)
        sql &= " AND (f.Admission='Y' OR f.Admission IS NULL)" & vbCrLf
        'b2.signUpStatus 排除 報名狀態 2:報名失敗 5:未錄取 
        'signUpStatus-0:收件完成 1:報名成功 2:報名失敗 3:正取(Key_SelResult) 4:備取 5:未錄取
        sql &= " AND NOT EXISTS (" & vbCrLf
        sql &= " 	SELECT 'x'" & vbCrLf
        sql &= " 	FROM Stud_EnterTemp2 a2" & vbCrLf
        sql &= " 	JOIN STUD_ENTERTYPE2 b2 ON a2.eSETID=b2.eSETID AND b2.signUpStatus IN (2,5)" & vbCrLf
        sql &= " 	WHERE a2.IDNO='" & IDNO & "' AND b2.ocid1=b.ocid1" & vbCrLf
        sql &= " )" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, oConn)
        For Each dr1 As DataRow In dt.Rows
            If OCIDs.IndexOf("'" & Convert.ToString(dr1("OCID")) & "'") = -1 Then
                If OCIDs <> "" Then OCIDs &= ","
                OCIDs &= "'" & Convert.ToString(dr1("OCID")) & "'"
            End If
        Next
        Return OCIDs
    End Function

    Sub SUtl_INSERT_STUD_ENTERTYPE(ByVal vSETID As String, ByVal rqEnterDate As String, ByVal vSerNum As String, ByVal NewExamNo As String, ByVal conn As SqlConnection, ByVal mainTrain As SqlTransaction, ByVal drCC1 As DataRow)
        If vSETID = "" Then Return
        If Val(vSETID) <= 0 Then Return

        Dim strErrmsg As String = "" '偵錯用儲存欄
        Dim strfield As String = "" '偵錯用儲存欄
        Dim strSql2 As String = "" '偵錯用儲存欄
        '建立STUD_ENTERTYPE專用command
        Dim param As SqlParameter
        Dim myParam As Hashtable = New Hashtable
        Dim sql As String = ""
        Dim MyKey As String = ""

        If ptype.Value = "shift" Then
            sql = ""
            sql &= " INSERT INTO STUD_ENTERTYPE (RelEnterDate,ExamNo,OCID1 ,TMID1 ,EnterChannel ,EnterPath ,EnterPath2 " ',OCID2 ,TMID2 ,OCID3 ,TMID3 
            sql &= " ,IdentityID ,MIDENTITYID ,RID ,PlanID ,CCLID ,ModifyAcct ,ModifyDate "
            'sql &= " ,HighEduBg ,WorkSuppIdent ,PriorWorkType1 ,PriorWorkOrg1 ,ActNo ,SOfficeYM1 ,FOfficeYM1 ,Notes "
            sql &= " ,WorkSuppIdent ,Notes "  'edit，by:20181024
            sql &= " ,TRNDMode,TRNDType,Ticket_NO "
            'sql &= " ,CMASTER1 ,APID1 "
            sql &= " ,APID1 "  'edit，by:20181024
            sql &= " ,SETID,EnterDate,SerNum) "
            sql &= "  VALUES(@RelEnterDate, @ExamNo,@OCID1, @TMID1, @EnterChannel ,@EnterPath ,@EnterPath2 " ', @OCID2, @TMID2, @OCID3, @TMID3
            sql &= " ,@IdentityID ,@MIDENTITYID , @RID, @PlanID, @CCLID, @ModifyAcct, GETDATE() "
            'sql &= " ,@HighEduBg, @WorkSuppIdent, @PriorWorkType1, @PriorWorkOrg1, @ActNo, @SOfficeYM1, @FOfficeYM1, @Notes "
            sql &= " ,@WorkSuppIdent, @Notes "  'edit，by:20181024
            sql &= " ,@TRNDMode, @TRNDType, @Ticket_NO "
            'sql &= " ,@CMASTER1 ,@APID1 "
            sql &= " ,@APID1 "  'edit，by:20181024
            sql &= " ,@SETID, " & TIMS.To_date(rqEnterDate) & ", @SerNum) "
        Else
            sql = ""
            sql &= " INSERT INTO STUD_ENTERTYPE (RelEnterDate ,ExamNo,OCID1 ,TMID1,EnterChannel ,EnterPath ,EnterPath2 " ',OCID2 ,TMID2 ,OCID3 ,TMID3 "
            sql &= " ,IdentityID ,MIDENTITYID ,RID ,PlanID ,CCLID ,ModifyAcct ,ModifyDate "
            'sql &= " ,HighEduBg ,WorkSuppIdent ,PriorWorkType1 ,PriorWorkOrg1 ,ActNo ,SOfficeYM1 ,FOfficeYM1 ,Notes "
            sql &= " ,WorkSuppIdent ,Notes "  'edit，by:20181024
            'sql += " ,TRNDMode ,TRNDType ,Ticket_NO ,CMASTER1 ,APID1 "
            sql &= " ,APID1 "  'edit，by:20181024
            sql &= " ,SETID ,EnterDate ,SerNum) "
            sql &= " VALUES(@RelEnterDate, @ExamNo ,@OCID1, @TMID1, @EnterChannel, @EnterPath ,@EnterPath2 " ', @OCID2, @TMID2, @OCID3, @TMID3
            sql &= " ,@IdentityID ,@MIDENTITYID, @RID, @PlanID, @CCLID, @ModifyAcct, GETDATE() "
            'sql &= " ,@HighEduBg, @WorkSuppIdent, @PriorWorkType1, @PriorWorkOrg1, @ActNo, @SOfficeYM1, @FOfficeYM1, @Notes "
            sql &= " ,@WorkSuppIdent, @Notes "  'edit，by:20181024
            'sql += " ,@TRNDMode ,@TRNDType ,@Ticket_NO ,@CMASTER1 ,@APID1 "
            sql &= " ,@APID1 "  'edit，by:20181024
            sql &= " ,@SETID, " & TIMS.To_date(rqEnterDate) & ", @SerNum) "
        End If

        strErrmsg = "" '偵錯用儲存欄
        strSql2 = sql '偵錯用儲存欄
        strfield = "" '偵錯用儲存欄
        Dim cmd2 As New SqlCommand(sql, conn, mainTrain)
        Try
            If RelEnterDate.Text <> "" _
                AndAlso TIMS.IsDate1(RelEnterDate.Text) Then
                RelEnterDate.Text = CDate(RelEnterDate.Text).ToString("yyyy/MM/dd")
            Else
                RelEnterDate.Text = CDate(aNow).ToString("yyyy/MM/dd") '異常 帶入當日
            End If
        Catch ex As Exception
        End Try

        param = cmd2.Parameters.Add("RelEnterDate", SqlDbType.DateTime)
        param.Value = CDate(RelEnterDate.Text & " " & FormatDateTime(aNow, DateFormat.ShortTime))
        Call TIMS.Set_STR2STR(strfield, param.Value, "RelEnterDate")
        myParam.Add("RelEnterDate", param.Value)

        param = cmd2.Parameters.Add("ExamNo", SqlDbType.VarChar)
        param.Value = NewExamNo 'ExamNoStr
        Call TIMS.Set_STR2STR(strfield, param.Value, "ExamNo")
        myParam.Add("ExamNo", param.Value)

        param = cmd2.Parameters.Add("OCID1", SqlDbType.Int)
        param.Value = Val(OCIDValue1.Value)
        Call TIMS.Set_STR2STR(strfield, param.Value, "OCID1")
        myParam.Add("OCID1", param.Value)

        param = cmd2.Parameters.Add("TMID1", SqlDbType.Int)
        param.Value = Val(TMIDValue1.Value)
        Call TIMS.Set_STR2STR(strfield, param.Value, "TMID1")
        myParam.Add("TMID1", param.Value)

        'param=cmd2.Parameters.Add("OCID2", SqlDbType.Int)
        'param.Value=If(OCIDValue2.Value <> "", Val(OCIDValue2.Value), Convert.DBNull)
        'Call TIMS.Set_STR2STR(strfield, param.Value, "OCID2")
        'myParam.Add("OCID2", param.Value)

        'param=cmd2.Parameters.Add("TMID2", SqlDbType.Int)
        'param.Value=If(TMIDValue2.Value <> "", Val(TMIDValue2.Value), Convert.DBNull)
        'Call TIMS.Set_STR2STR(strfield, param.Value, "TMID2")
        'myParam.Add("TMID2", param.Value)

        'param=cmd2.Parameters.Add("OCID3", SqlDbType.Int)
        'param.Value=If(OCIDValue3.Value <> "", Val(OCIDValue3.Value), Convert.DBNull)
        'Call TIMS.Set_STR2STR(strfield, param.Value, "OCID3")
        'myParam.Add("OCID3", param.Value)

        'param=cmd2.Parameters.Add("TMID3", SqlDbType.Int)
        'param.Value=If(TMIDValue3.Value <> "", Val(TMIDValue3.Value), Convert.DBNull)
        'Call TIMS.Set_STR2STR(strfield, param.Value, "TMID3")
        'myParam.Add("TMID3", param.Value)

        MyKey = "2" '1.網;2.現;3.通;4.推
        If EnterChannel.SelectedValue <> "" Then
            MyKey = EnterChannel.SelectedValue '1.網;2.現;3.通;4.推
        End If
        param = cmd2.Parameters.Add("EnterChannel", SqlDbType.VarChar)
        param.Value = MyKey 'EnterChannel.SelectedValue
        Call TIMS.Set_STR2STR(strfield, param.Value, "EnterChannel")
        myParam.Add("EnterChannel", param.Value)

        param = cmd2.Parameters.Add("EnterPath", SqlDbType.Char, 1)
        param.Value = "S" '報名登錄送出
        Call TIMS.Set_STR2STR(strfield, param.Value, "EnterPath")
        myParam.Add("EnterPath", param.Value)

        param = cmd2.Parameters.Add("EnterPath2", SqlDbType.VarChar)
        Select Case Convert.ToString(Request("ID"))
            Case TIMS.cst_FunID_報名登錄 '報名登錄(SD_01_001_add)
                param.Value = "N" '報名登錄
            Case TIMS.cst_FunID_專案核定報名登錄 '專案核定報名登錄
                param.Value = "P" '專案核定報名登錄
            Case TIMS.cst_FunID_特例專案核定報名登錄
                param.Value = "S"
            Case Else
                param.Value = Convert.DBNull
                'Common.MessageBox(Me, "請重新查詢登入!!", "SD_01_001.aspx?ID=" & cst_funid報名登錄)
                'Exit Sub
        End Select
        Call TIMS.Set_STR2STR(strfield, param.Value, "EnterPath2")
        myParam.Add("EnterPath2", param.Value)

        param = cmd2.Parameters.Add("IdentityID", SqlDbType.VarChar, 50)
        'Dim MyKey As String=""
        MyKey = TIMS.GetCblValue(IdentityID)
        param.Value = MyKey
        Call TIMS.Set_STR2STR(strfield, param.Value, "IdentityID")
        myParam.Add("IdentityID", param.Value)

        Dim v_MIdentityID As String = TIMS.GetListValue(MIdentityID)
        param = cmd2.Parameters.Add("MIDENTITYID", SqlDbType.VarChar, 3)
        param.Value = If(v_MIdentityID <> "", v_MIdentityID, Convert.DBNull)
        Call TIMS.Set_STR2STR(strfield, param.Value, "MIDENTITYID")
        myParam.Add("MIDENTITYID", param.Value)

        param = cmd2.Parameters.Add("RID", SqlDbType.VarChar, 10)
        param.Value = drCC1("RID") 'sm.UserInfo.RID
        Call TIMS.Set_STR2STR(strfield, param.Value, "RID")
        myParam.Add("RID", param.Value)

        param = cmd2.Parameters.Add("PlanID", SqlDbType.Int)
        param.Value = drCC1("PlanID") 'sm.UserInfo.PlanID
        Call TIMS.Set_STR2STR(strfield, param.Value, "PlanID")
        myParam.Add("PlanID", param.Value)

        param = cmd2.Parameters.Add("CCLID", SqlDbType.Int)
        param.Value = If(CCLID.Value = "", Convert.DBNull, Val(CCLID.Value))
        Call TIMS.Set_STR2STR(strfield, param.Value, "CCLID")
        myParam.Add("CCLID", param.Value)

        param = cmd2.Parameters.Add("ModifyAcct", SqlDbType.VarChar, 15)
        param.Value = sm.UserInfo.UserID
        Call TIMS.Set_STR2STR(strfield, param.Value, "ModifyAcct")
        myParam.Add("ModifyAcct", param.Value)

        ''20090330專上畢業學歷失業者
        'If rdo_HighEduBg.SelectedValue <> "" Then
        '    cmd2.Parameters.Add("HighEduBg", SqlDbType.Char).Value=rdo_HighEduBg.SelectedValue
        '    Call TIMS.Set_STR2STR(strfield, rdo_HighEduBg.SelectedValue, "HighEduBg")
        'Else
        '    cmd2.Parameters.Add("HighEduBg", SqlDbType.Char).Value=Convert.DBNull
        '    Call TIMS.Set_STR2STR(strfield, Convert.DBNull, "HighEduBg")
        'End If

        If WSITR.Visible Then
            If rblWorkSuppIdent.SelectedValue <> "" Then
                cmd2.Parameters.Add("WorkSuppIdent", SqlDbType.Char).Value = rblWorkSuppIdent.SelectedValue
                Call TIMS.Set_STR2STR(strfield, rblWorkSuppIdent.SelectedValue, "WorkSuppIdent")
                myParam.Add("WorkSuppIdent", rblWorkSuppIdent.SelectedValue)
            Else
                cmd2.Parameters.Add("WorkSuppIdent", SqlDbType.Char).Value = Convert.DBNull
                Call TIMS.Set_STR2STR(strfield, Convert.DBNull, "WorkSuppIdent")
                myParam.Add("WorkSuppIdent", Convert.DBNull)
            End If
        Else
            cmd2.Parameters.Add("WorkSuppIdent", SqlDbType.Char).Value = Convert.DBNull
            Call TIMS.Set_STR2STR(strfield, Convert.DBNull, "WorkSuppIdent")
            myParam.Add("WorkSuppIdent", Convert.DBNull)
        End If

        ''----------(職前課程邏輯)受訓前任職資料start-----------
        'param=cmd2.Parameters.Add("PriorWorkType1", SqlDbType.Char, 1)
        'param.Value=If(PriorWorkType1.SelectedValue="", Convert.DBNull, PriorWorkType1.SelectedValue)
        'Call TIMS.Set_STR2STR(strfield, param.Value, "PriorWorkType1")

        'param=cmd2.Parameters.Add("PriorWorkOrg1", SqlDbType.NVarChar, 30)
        'param.Value=If(PriorWorkOrg1.Text="", Convert.DBNull, PriorWorkOrg1.Text)
        'Call TIMS.Set_STR2STR(strfield, param.Value, "PriorWorkOrg1")

        'param=cmd2.Parameters.Add("ActNo", SqlDbType.VarChar, 9)
        'param.Value=If(ActNo.Text="", Convert.DBNull, ActNo.Text)
        'Call TIMS.Set_STR2STR(strfield, param.Value, "ActNo")

        'param=cmd2.Parameters.Add("SOfficeYM1", SqlDbType.DateTime)
        'param.Value=If(SOfficeYM1.Text="", Convert.DBNull, TIMS.cdate2(SOfficeYM1.Text))
        'Call TIMS.Set_STR2STR(strfield, param.Value, "SOfficeYM1")

        'param=cmd2.Parameters.Add("FOfficeYM1", SqlDbType.DateTime)
        'param.Value=If(FOfficeYM1.Text="", Convert.DBNull, TIMS.cdate2(FOfficeYM1.Text))
        'Call TIMS.Set_STR2STR(strfield, param.Value, "FOfficeYM1")
        ''----------受訓前任職資料end-------------

        notes.Text = Trim(notes.Text)
        If notes.Text <> "" Then
            cmd2.Parameters.Add("Notes", SqlDbType.NVarChar).Value = notes.Text
            Call TIMS.Set_STR2STR(strfield, notes.Text, "Notes")
            myParam.Add("Notes", notes.Text)
        Else
            cmd2.Parameters.Add("Notes", SqlDbType.NVarChar).Value = Convert.DBNull
            Call TIMS.Set_STR2STR(strfield, Convert.DBNull, "Notes")
            myParam.Add("Notes", Convert.DBNull)
        End If

        If ptype.Value = "shift" Then
            param = cmd2.Parameters.Add("TRNDMode", SqlDbType.Int)
            param.Value = If(Hid_rqTicket.Value = "", Convert.DBNull, Val(Hid_rqTicket.Value))
            Call TIMS.Set_STR2STR(strfield, param.Value, "TRNDMode")
            myParam.Add("TRNDMode", param.Value)

            param = cmd2.Parameters.Add("TRNDType", SqlDbType.Int)
            param.Value = If(Hid_ticketType.Value = "", Convert.DBNull, Val(Hid_ticketType.Value))
            Call TIMS.Set_STR2STR(strfield, param.Value, "TRNDType")
            myParam.Add("TRNDType", param.Value)

            param = cmd2.Parameters.Add("Ticket_NO", SqlDbType.VarChar)
            param.Value = If(Hid_ticketType.Value = "", Convert.DBNull, Val(Hid_ticketType.Value))
            Call TIMS.Set_STR2STR(strfield, param.Value, "Ticket_NO")
            myParam.Add("Ticket_NO", param.Value)
        End If

        'param=cmd2.Parameters.Add("CMASTER1", SqlDbType.VarChar)
        ''If HidMaster.Value="Y" Then param.Value=Me.HidMaster.Value Else param.Value=Convert.DBNull
        'Call TIMS.Set_STR2STR(strfield, param.Value, "CMASTER1")

        param = cmd2.Parameters.Add("APID1", SqlDbType.VarChar)
        MyKey = TIMS.GetCblValue(cblAVTCP1)
        If MyKey <> "" Then param.Value = MyKey Else param.Value = Convert.DBNull
        Call TIMS.Set_STR2STR(strfield, param.Value, "APID1")
        myParam.Add("APID1", param.Value)

        param = cmd2.Parameters.Add("SETID", SqlDbType.Int)
        param.Value = Val(vSETID)
        Call TIMS.Set_STR2STR(strfield, param.Value, "SETID")
        myParam.Add("SETID", param.Value)

        'param=cmd2.Parameters.Add("EnterDate", SqlDbType.DateTime)
        'param.Value=rqEnterDate 'CDate(Now()).ToString("yyyy/MM/dd")
        'Call TIMS.Set_STR2STR(strfield, param.Value, "EnterDate")

        param = cmd2.Parameters.Add("SerNum", SqlDbType.Int)
        param.Value = Val(vSerNum)
        Call TIMS.Set_STR2STR(strfield, param.Value, "SerNum")
        myParam.Add("SerNum", param.Value)

        Try
            'cmd2.ExecuteNonQuery()
            DbAccess.ExecuteNonQuery(sql, myParam)
        Catch ex As Exception
            strErrmsg = ex.ToString & vbCrLf '偵錯用儲存欄
            If Not mainTrain Is Nothing Then mainTrain.Rollback()
            Call TIMS.CloseDbConn(conn)
            Common.MessageBox(Me, strErrmsg)

            '偵錯用儲存欄
            strErrmsg += "/* strSql2: */" & vbCrLf
            strErrmsg += strSql2 & vbCrLf
            strErrmsg += "/* field: */" & vbCrLf
            strErrmsg += strfield & vbCrLf
            strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)

            Throw ex
        End Try
    End Sub

    Sub SUtl_UPDATE_STUD_ENTERTRAIN(ByVal vSETID As String, ByVal rqEnterDate As String, ByVal vSerNum As String,
                                    ByVal conn As SqlConnection, ByVal mainTrain As SqlTransaction, ByVal drCC1 As DataRow)
        If vSETID = "" Then Return
        If Val(vSETID) <= 0 Then Return

        Dim parms As New Hashtable From {
            {"SETID", TIMS.CINT1(vSETID)},
            {"EnterDate", TIMS.Cdate2(rqEnterDate)},
            {"SerNum", TIMS.CINT1(vSerNum)}
        }
        Dim sql As String = ""
        sql &= " SELECT 'X'" & vbCrLf ' a.SENID  /*PK*/
        sql &= " FROM STUD_ENTERTRAIN a" & vbCrLf
        sql &= " WHERE a.SETID=@SETID AND a.EnterDate=@EnterDate AND a.SerNum=@SerNum" & vbCrLf
        Dim dt1 As DataTable = DbAccess.GetDataTable(sql, mainTrain, parms)

        Dim t_ddlSERVDEPTID As String = TIMS.GetListText(ddlSERVDEPTID)
        JobTitle.Text = TIMS.ClearSQM(JobTitle.Text)
        Dim v_ddlJOBTITLEID As String = TIMS.GetListValue(ddlJOBTITLEID)

        ZipCode2.Value = TIMS.ClearSQM(ZipCode2.Value)
        ZipCode2_B3.Value = TIMS.ClearSQM(ZipCode2_B3.Value)
        HidZipCode2_6W.Value = TIMS.GetZIPCODE6W(ZipCode2.Value, ZipCode2_B3.Value) 'String.Concat(ZipCode2.Value, ZipCode2_B3.Value)
        HouseholdAddress.Text = TIMS.ClearSQM(HouseholdAddress.Text)

        Uname.Text = TIMS.ClearSQM(Uname.Text)
        Intaxno.Text = TIMS.ClearSQM(Intaxno.Text)
        ActName.Text = TIMS.ClearSQM(ActName.Text)

        Dim v_ActType As String = TIMS.GetListValue(ActType)
        ActNo.Text = TIMS.ClearSQM(ActNo.Text)
        ActTel.Text = TIMS.ClearSQM(ActTel.Text)

        ZipCode3.Value = TIMS.ClearSQM(ZipCode3.Value)
        ZipCode3_B3.Value = TIMS.ClearSQM(ZipCode3_B3.Value)
        HidZipCode3_6W.Value = TIMS.GetZIPCODE6W(ZipCode3.Value, ZipCode3_B3.Value) 'String.Concat(ZipCode3.Value, ZipCode3_B3.Value)
        ActAddress.Text = TIMS.ClearSQM(ActAddress.Text)

        If dt1.Rows.Count = 0 Then
            Dim i_sql As String = ""
            i_sql &= " INSERT INTO STUD_ENTERTRAIN (SENID,SETID,ENTERDATE,SERNUM,ZIPCODE2,ZIPCODE2_6W,HOUSEHOLDADDRESS" & vbCrLf
            i_sql &= " ,UNAME,INTAXNO,SERVDEPT,JOBTITLE,ACTNAME,ACTTYPE,ACTNO,ACTTEL" & vbCrLf
            i_sql &= " ,ZIPCODE3,ZIPCODE3_6W,ACTADDRESS,SERVDEPTID,JOBTITLEID" & vbCrLf
            i_sql &= " ,MODIFYACCT,MODIFYDATE)" & vbCrLf
            i_sql &= " VALUES (@SENID,@SETID,@ENTERDATE,@SERNUM,@ZIPCODE2,@ZIPCODE2_6W,@HOUSEHOLDADDRESS" & vbCrLf
            i_sql &= " ,@UNAME,@INTAXNO,@SERVDEPT,@JOBTITLE,@ACTNAME,@ACTTYPE,@ACTNO,@ACTTEL" & vbCrLf
            i_sql &= " ,@ZIPCODE3,@ZIPCODE3_6W,@ACTADDRESS,@SERVDEPTID,@JOBTITLEID" & vbCrLf
            i_sql &= " ,@MODIFYACCT,GETDATE())" & vbCrLf

            Dim iSENID As Integer = DbAccess.GetNewId(mainTrain, "STUD_ENTERTRAIN_SENID_SEQ,STUD_ENTERTRAIN,SENID")
            Dim i_parms As New Hashtable
            i_parms.Add("SENID", iSENID)
            i_parms.Add("SETID", Val(vSETID))
            i_parms.Add("EnterDate", CDate(rqEnterDate))
            i_parms.Add("SerNum", Val(vSerNum))
            i_parms.Add("ZIPCODE2", TIMS.GetValue1(ZipCode2.Value))
            i_parms.Add("ZIPCODE2_6W", TIMS.GetValue1(HidZipCode2_6W.Value))
            i_parms.Add("HOUSEHOLDADDRESS", If(HouseholdAddress.Text <> "", HouseholdAddress.Text, Convert.DBNull))

            i_parms.Add("UNAME", If(Uname.Text <> "", Uname.Text, Convert.DBNull))
            i_parms.Add("INTAXNO", If(Intaxno.Text <> "", Intaxno.Text, Convert.DBNull))
            i_parms.Add("SERVDEPT", If(t_ddlSERVDEPTID <> "", TIMS.GetValue1(t_ddlSERVDEPTID), TIMS.GetValue1(ServDept.Text)))
            i_parms.Add("JOBTITLE", If(ddlJOBTITLEID.SelectedItem.Text <> "" AndAlso v_ddlJOBTITLEID <> "", TIMS.GetValue1(ddlJOBTITLEID.SelectedItem.Text), TIMS.GetValue1(JobTitle.Text)))
            i_parms.Add("ACTNAME", If(ActName.Text <> "", ActName.Text, Convert.DBNull))
            i_parms.Add("ACTTYPE", If(v_ActType <> "", v_ActType, Convert.DBNull))
            i_parms.Add("ACTNO", If(ActNo.Text <> "", ActNo.Text, Convert.DBNull))
            i_parms.Add("ACTTEL", If(ActTel.Text <> "", ActTel.Text, Convert.DBNull))

            i_parms.Add("ZIPCODE3", TIMS.GetValue1(ZipCode3.Value))
            i_parms.Add("ZIPCODE3_6W", TIMS.GetValue1(HidZipCode3_6W.Value))
            i_parms.Add("ACTADDRESS", If(ActAddress.Text <> "", ActAddress.Text, Convert.DBNull))
            i_parms.Add("SERVDEPTID", TIMS.GetValue1(TIMS.GetListValue(ddlSERVDEPTID)))
            i_parms.Add("JOBTITLEID", TIMS.GetValue1(TIMS.GetListValue(ddlJOBTITLEID)))

            i_parms.Add("MODIFYACCT", sm.UserInfo.UserID)
            Try
                DbAccess.ExecuteNonQuery(i_sql, mainTrain, i_parms)
            Catch ex As Exception
                Dim cst_fun_page_name As String = "##SD_01_001_add.aspx, "
                Dim slogMsg1 As String = String.Concat(cst_fun_page_name, ",i_sql: ", i_sql, vbCrLf, ",i_parms: ", TIMS.GetMyValue5(i_parms), vbCrLf)
                Dim strErrmsg As String = String.Concat("ex.Message: ", ex.Message, vbCrLf, "ex.ToString: ", ex.ToString, vbCrLf, "slogMsg1: ", slogMsg1)
                Call TIMS.SendMailTest(strErrmsg)
                Throw ex
            End Try
        Else
            Dim u_sql As String = ""
            u_sql &= " UPDATE STUD_ENTERTRAIN" & vbCrLf ' a.SENID  /*PK*/
            u_sql &= " SET ZIPCODE2=@ZIPCODE2" & vbCrLf
            u_sql &= " ,ZIPCODE2_6W=@ZIPCODE2_6W" & vbCrLf
            u_sql &= " ,HOUSEHOLDADDRESS=@HOUSEHOLDADDRESS" & vbCrLf
            'u_sql &= " ,HANDTYPEID=@HANDTYPEID" & vbCrLf
            'u_sql &= " ,HANDLEVELID=@HANDLEVELID" & vbCrLf
            u_sql &= " ,UNAME=@UNAME" & vbCrLf
            u_sql &= " ,INTAXNO=@INTAXNO" & vbCrLf
            u_sql &= " ,SERVDEPT=@SERVDEPT" & vbCrLf
            u_sql &= " ,JOBTITLE=@JOBTITLE" & vbCrLf
            u_sql &= " ,ACTNAME=@ACTNAME" & vbCrLf
            u_sql &= " ,ACTTYPE=@ACTTYPE" & vbCrLf
            u_sql &= " ,ACTNO=@ACTNO" & vbCrLf
            u_sql &= " ,ACTTEL=@ACTTEL" & vbCrLf

            u_sql &= " ,ZIPCODE3=@ZIPCODE3" & vbCrLf
            u_sql &= " ,ZIPCODE3_6W=@ZIPCODE3_6W" & vbCrLf
            u_sql &= " ,ACTADDRESS=@ACTADDRESS" & vbCrLf
            u_sql &= " ,SERVDEPTID=@SERVDEPTID" & vbCrLf
            u_sql &= " ,JOBTITLEID=@JOBTITLEID" & vbCrLf
            u_sql &= " ,MODIFYACCT=@MODIFYACCT" & vbCrLf
            u_sql &= " ,MODIFYDATE=GETDATE()" & vbCrLf
            u_sql &= " FROM STUD_ENTERTRAIN a" & vbCrLf
            u_sql &= " WHERE SETID=@SETID" & vbCrLf
            u_sql &= " AND EnterDate=@EnterDate" & vbCrLf
            u_sql &= " AND SerNum=@SerNum" & vbCrLf

            Dim u_parms As New Hashtable
            u_parms.Add("ZIPCODE2", If(ZipCode2.Value <> "", TIMS.GetValue1(ZipCode2.Value), Convert.DBNull))
            u_parms.Add("ZIPCODE2_6W", If(HidZipCode2_6W.Value <> "", TIMS.GetValue1(HidZipCode2_6W.Value), Convert.DBNull))
            u_parms.Add("HOUSEHOLDADDRESS", If(HouseholdAddress.Text <> "", HouseholdAddress.Text, Convert.DBNull))

            u_parms.Add("UNAME", If(Uname.Text <> "", Uname.Text, Convert.DBNull))
            u_parms.Add("INTAXNO", If(Intaxno.Text <> "", Intaxno.Text, Convert.DBNull))
            u_parms.Add("SERVDEPT", If(t_ddlSERVDEPTID <> "", TIMS.GetValue1(t_ddlSERVDEPTID), TIMS.GetValue1(ServDept.Text)))
            u_parms.Add("JOBTITLE", If(ddlJOBTITLEID.SelectedItem.Text <> "" AndAlso v_ddlJOBTITLEID <> "", TIMS.GetValue1(ddlJOBTITLEID.SelectedItem.Text), TIMS.GetValue1(JobTitle.Text)))
            u_parms.Add("ACTNAME", If(ActName.Text <> "", ActName.Text, Convert.DBNull))
            u_parms.Add("ACTTYPE", If(v_ActType <> "", v_ActType, Convert.DBNull))
            u_parms.Add("ACTNO", If(ActNo.Text <> "", ActNo.Text, Convert.DBNull))
            u_parms.Add("ACTTEL", If(ActTel.Text <> "", ActTel.Text, Convert.DBNull))

            u_parms.Add("ZIPCODE3", If(ZipCode3.Value <> "", TIMS.GetValue1(ZipCode3.Value), Convert.DBNull))
            u_parms.Add("ZIPCODE3_6W", If(HidZipCode3_6W.Value <> "", TIMS.GetValue1(HidZipCode3_6W.Value), Convert.DBNull))
            u_parms.Add("ACTADDRESS", If(ActAddress.Text <> "", ActAddress.Text, Convert.DBNull))
            u_parms.Add("SERVDEPTID", TIMS.GetValue1(TIMS.GetListValue(ddlSERVDEPTID)))
            u_parms.Add("JOBTITLEID", TIMS.GetValue1(TIMS.GetListValue(ddlJOBTITLEID)))
            u_parms.Add("MODIFYACCT", sm.UserInfo.UserID)

            u_parms.Add("SETID", Val(vSETID))
            u_parms.Add("EnterDate", CDate(rqEnterDate))
            u_parms.Add("SerNum", Val(vSerNum))
            Try
                DbAccess.ExecuteNonQuery(u_sql, mainTrain, u_parms)
            Catch ex As Exception
                Dim cst_fun_page_name As String = "##SD_01_001_add.aspx, "
                Dim slogMsg1 As String = String.Concat(cst_fun_page_name, ",u_sql: ", u_sql, vbCrLf, ",u_parms: ", TIMS.GetMyValue5(u_parms), vbCrLf)
                Dim strErrmsg As String = String.Concat("ex.Message: ", ex.Message, vbCrLf, "ex.ToString: ", ex.ToString, vbCrLf, "slogMsg1: ", slogMsg1)
                Call TIMS.SendMailTest(strErrmsg)
                Throw ex
            End Try
        End If
    End Sub

    Sub SUtl_UPDATE_STUD_ENTERTYPE(ByVal vSETID As String, ByVal rqEnterDate As String, ByVal vSerNum As String, ByVal NewExamNo As String,
                                   ByVal conn As SqlConnection, ByVal mainTrain As SqlTransaction, ByVal drCC1 As DataRow)

        If vSETID = "" Then Return
        If Val(vSETID) <= 0 Then Return

        Dim MyKey As String = ""
        Dim strErrmsg As String = ""              '偵錯用儲存欄
        Dim strfield As String = ""               '偵錯用儲存欄
        Dim strSql2 As String = ""                '偵錯用儲存欄
        Dim cmd2 As New SqlCommand                '建立STUD_ENTERTYPE專用command
        Dim param As SqlParameter
        Dim myParam As Hashtable = New Hashtable
        Dim sql As String = ""
        sql &= " UPDATE STUD_ENTERTYPE "
        sql &= " SET RelEnterDate=@RelEnterDate ,ExamNo=@ExamNo"
        sql &= " ,OCID1=@OCID1 ,TMID1=@TMID1" ' ,OCID2=@OCID2 ,TMID2=@TMID2 ,OCID3=@OCID3 ,TMID3=@TMID3 "
        '異動不應該寫入此值。EnterPath=@EnterPath,
        sql &= " ,EnterChannel=@EnterChannel "
        sql &= " ,IdentityID=@IdentityID "
        sql &= " ,MIDENTITYID=@MIDENTITYID "
        sql &= " ,RID=@RID "
        sql &= " ,PlanID=@PlanID ,CCLID=@CCLID ,ModifyAcct=@ModifyAcct ,ModifyDate=GETDATE() "  'edit，by:20181024
        sql &= " ,WorkSuppIdent=@WorkSuppIdent "  'edit，by:20181024
        sql &= " ,Notes=@Notes "  'edit，by:20181024
        sql &= " ,APID1=@APID1 "  'edit，by:20181024
        sql &= " WHERE SETID=@SETID AND EnterDate=@EnterDate AND SerNum=@SerNum "
        'sql &= "    ,PlanID=@PlanID ,CCLID=@CCLID ,ModifyAcct=@ModifyAcct ,ModifyDate=GETDATE() ,HighEduBg=@HighEduBg "
        'sql &= "    ,WorkSuppIdent=@WorkSuppIdent ,PriorWorkType1=@PriorWorkType1 ,PriorWorkOrg1=@PriorWorkOrg1 "
        'sql &= "    ,ActNo=@ActNo ,SOfficeYM1=@SOfficeYM1 ,FOfficeYM1=@FOfficeYM1 ,Notes=@Notes "

        strErrmsg = "" '偵錯用儲存欄
        strSql2 = sql '偵錯用儲存欄
        strfield = "" '偵錯用儲存欄
        cmd2 = New SqlCommand(sql, conn, mainTrain)

        Try
            If RelEnterDate.Text <> "" AndAlso TIMS.IsDate1(RelEnterDate.Text) Then
                RelEnterDate.Text = CDate(RelEnterDate.Text).ToString("yyyy/MM/dd")
            Else
                RelEnterDate.Text = CDate(aNow).ToString("yyyy/MM/dd") '異常 帶入當日
            End If
        Catch ex As Exception
        End Try

        param = cmd2.Parameters.Add("RelEnterDate", SqlDbType.DateTime)
        param.Value = RelEnterDate.Text & " " & FormatDateTime(aNow, DateFormat.ShortTime)
        Call TIMS.Set_STR2STR(strfield, param.Value, "RelEnterDate")
        myParam.Add("RelEnterDate", param.Value)

        param = cmd2.Parameters.Add("ExamNo", SqlDbType.VarChar)
        param.Value = NewExamNo 'ExamNoStr
        Call TIMS.Set_STR2STR(strfield, param.Value, "ExamNo")
        myParam.Add("ExamNo", param.Value)

        OCIDValue1.Value = Trim(OCIDValue1.Value)
        param = cmd2.Parameters.Add("OCID1", SqlDbType.Int)
        param.Value = Val(OCIDValue1.Value)
        Call TIMS.Set_STR2STR(strfield, param.Value, "OCID1")
        myParam.Add("OCID1", param.Value)

        TMIDValue1.Value = Trim(TMIDValue1.Value)
        param = cmd2.Parameters.Add("TMID1", SqlDbType.Int)
        param.Value = Val(TMIDValue1.Value)
        Call TIMS.Set_STR2STR(strfield, param.Value, "TMID1")
        myParam.Add("TMID1", param.Value)

        'OCIDValue2.Value=Trim(OCIDValue2.Value)
        'param=cmd2.Parameters.Add("OCID2", SqlDbType.Int)
        'param.Value=If(OCIDValue2.Value="", Convert.DBNull, OCIDValue2.Value)
        'Call TIMS.Set_STR2STR(strfield, param.Value, "OCID2")
        'myParam.Add("OCID2", param.Value)

        'TMIDValue2.Value=Trim(TMIDValue2.Value)
        'param=cmd2.Parameters.Add("TMID2", SqlDbType.Int)
        'param.Value=If(TMIDValue2.Value="", Convert.DBNull, TMIDValue2.Value)
        'Call TIMS.Set_STR2STR(strfield, param.Value, "TMID2")
        'myParam.Add("TMID2", param.Value)

        'OCIDValue3.Value=Trim(OCIDValue3.Value)
        'param=cmd2.Parameters.Add("OCID3", SqlDbType.Int)
        'param.Value=If(OCIDValue3.Value="", Convert.DBNull, OCIDValue3.Value)
        'Call TIMS.Set_STR2STR(strfield, param.Value, "OCID3")
        'myParam.Add("OCID3", param.Value)

        'TMIDValue3.Value=Trim(TMIDValue3.Value)
        'param=cmd2.Parameters.Add("TMID3", SqlDbType.Int)
        'param.Value=If(TMIDValue3.Value="", Convert.DBNull, TMIDValue3.Value)
        'Call TIMS.Set_STR2STR(strfield, param.Value, "TMID3")
        'myParam.Add("TMID3", param.Value)

        MyKey = "2" '1.網;2.現;3.通;4.推
        If EnterChannel.SelectedValue <> "" Then
            MyKey = EnterChannel.SelectedValue '1.網;2.現;3.通;4.推
        End If
        param = cmd2.Parameters.Add("EnterChannel", SqlDbType.VarChar)
        param.Value = MyKey
        Call TIMS.Set_STR2STR(strfield, param.Value, "EnterChannel")
        myParam.Add("EnterChannel", param.Value)

        '異動不寫入此值
        'param=cmd2.Parameters.Add("EnterPath", SqlDbType.Char, 1)
        'param.Value="S" '報名登錄送出
        'Call TIMS.Set_STR2STR(strfield, param.Value, "EnterPath")

        'MyKey=""
        MyKey = TIMS.GetCblValue(IdentityID)
        param = cmd2.Parameters.Add("IdentityID", SqlDbType.VarChar, 50)
        param.Value = MyKey
        Call TIMS.Set_STR2STR(strfield, param.Value, "IdentityID")
        myParam.Add("IdentityID", param.Value)

        Dim v_MIdentityID As String = TIMS.GetListValue(MIdentityID)
        param = cmd2.Parameters.Add("MIDENTITYID", SqlDbType.VarChar, 3)
        param.Value = If(v_MIdentityID <> "", v_MIdentityID, Convert.DBNull)
        Call TIMS.Set_STR2STR(strfield, param.Value, "MIDENTITYID")
        myParam.Add("MIDENTITYID", param.Value)

        param = cmd2.Parameters.Add("RID", SqlDbType.VarChar, 10)
        param.Value = drCC1("RID") 'sm.UserInfo.RID
        Call TIMS.Set_STR2STR(strfield, param.Value, "RID")
        myParam.Add("RID", param.Value)

        param = cmd2.Parameters.Add("PlanID", SqlDbType.Int)
        param.Value = drCC1("PlanID") 'sm.UserInfo.PlanID
        Call TIMS.Set_STR2STR(strfield, param.Value, "PlanID")
        myParam.Add("PlanID", param.Value)

        If CCLID.Value <> "" Then CCLID.Value = Trim(CCLID.Value)
        If CCLID.Value <> "" Then CCLID.Value = Val(CCLID.Value)
        param = cmd2.Parameters.Add("CCLID", SqlDbType.Int)
        param.Value = If(CCLID.Value = "", Convert.DBNull, CCLID.Value)
        Call TIMS.Set_STR2STR(strfield, param.Value, "CCLID")
        myParam.Add("CCLID", param.Value)

        param = cmd2.Parameters.Add("ModifyAcct", SqlDbType.VarChar, 15)
        param.Value = sm.UserInfo.UserID
        Call TIMS.Set_STR2STR(strfield, param.Value, "ModifyAcct")
        myParam.Add("ModifyAcct", param.Value)

        'param=cmd2.Parameters.Add("ModifyDate", SqlDbType.DateTime)
        'param.Value=Now()
        'Call TIMS.Set_STR2STR(strfield, param.Value, "ModifyDate")

        ''20090330專上畢業學歷失業者(職前課程邏輯)
        'If rdo_HighEduBg.SelectedValue <> "" Then
        '    cmd2.Parameters.Add("HighEduBg", SqlDbType.Char).Value=rdo_HighEduBg.SelectedValue
        '    Call TIMS.Set_STR2STR(strfield, rdo_HighEduBg.SelectedValue, "HighEduBg")
        'Else
        '    cmd2.Parameters.Add("HighEduBg", SqlDbType.Char).Value=Convert.DBNull
        '    Call TIMS.Set_STR2STR(strfield, Convert.DBNull, "HighEduBg")
        'End If

        If WSITR.Visible Then
            If rblWorkSuppIdent.SelectedValue <> "" Then
                cmd2.Parameters.Add("WorkSuppIdent", SqlDbType.Char).Value = rblWorkSuppIdent.SelectedValue
                Call TIMS.Set_STR2STR(strfield, rblWorkSuppIdent.SelectedValue, "WorkSuppIdent")
                myParam.Add("WorkSuppIdent", rblWorkSuppIdent.SelectedValue)
            Else
                cmd2.Parameters.Add("WorkSuppIdent", SqlDbType.Char).Value = Convert.DBNull
                Call TIMS.Set_STR2STR(strfield, Convert.DBNull, "WorkSuppIdent")
                myParam.Add("WorkSuppIdent", Convert.DBNull)
            End If
        Else
            cmd2.Parameters.Add("WorkSuppIdent", SqlDbType.Char).Value = Convert.DBNull
            Call TIMS.Set_STR2STR(strfield, Convert.DBNull, "WorkSuppIdent")
            myParam.Add("WorkSuppIdent", Convert.DBNull)
        End If

        ''----------受訓前任職資料start(職前課程邏輯)-----------
        'param=cmd2.Parameters.Add("PriorWorkType1", SqlDbType.Char, 1)
        'param.Value=If(PriorWorkType1.SelectedValue="", Convert.DBNull, PriorWorkType1.SelectedValue)
        'Call TIMS.Set_STR2STR(strfield, param.Value, "PriorWorkType1")

        'param=cmd2.Parameters.Add("PriorWorkOrg1", SqlDbType.NVarChar, 30)
        'param.Value=If(PriorWorkOrg1.Text="", Convert.DBNull, PriorWorkOrg1.Text)
        'Call TIMS.Set_STR2STR(strfield, param.Value, "PriorWorkOrg1")

        'param=cmd2.Parameters.Add("ActNo", SqlDbType.VarChar, 9)
        'param.Value=If(ActNo.Text="", Convert.DBNull, ActNo.Text)
        'Call TIMS.Set_STR2STR(strfield, param.Value, "ActNo")

        'param=cmd2.Parameters.Add("SOfficeYM1", SqlDbType.DateTime)
        'param.Value=If(SOfficeYM1.Text="", Convert.DBNull, TIMS.cdate2(SOfficeYM1.Text))
        'Call TIMS.Set_STR2STR(strfield, param.Value, "SOfficeYM1")

        'param=cmd2.Parameters.Add("FOfficeYM1", SqlDbType.DateTime)
        'param.Value=If(FOfficeYM1.Text="", Convert.DBNull, TIMS.cdate2(FOfficeYM1.Text))
        'Call TIMS.Set_STR2STR(strfield, param.Value, "FOfficeYM1")
        ''----------受訓前任職資料end-------------

        notes.Text = Trim(notes.Text)
        If notes.Text <> "" Then
            cmd2.Parameters.Add("Notes", SqlDbType.NVarChar).Value = notes.Text
            Call TIMS.Set_STR2STR(strfield, notes.Text, "Notes")
            myParam.Add("Notes", notes.Text)
        Else
            cmd2.Parameters.Add("Notes", SqlDbType.NVarChar).Value = Convert.DBNull
            Call TIMS.Set_STR2STR(strfield, Convert.DBNull, "Notes")
            myParam.Add("Notes", Convert.DBNull)
        End If

        param = cmd2.Parameters.Add("APID1", SqlDbType.VarChar)
        MyKey = TIMS.GetCblValue(cblAVTCP1)
        If MyKey <> "" Then param.Value = MyKey Else param.Value = Convert.DBNull
        Call TIMS.Set_STR2STR(strfield, param.Value, "APID1")
        myParam.Add("APID1", param.Value)

        param = cmd2.Parameters.Add("SETID", SqlDbType.Int)
        param.Value = vSETID
        Call TIMS.Set_STR2STR(strfield, param.Value, "SETID")
        myParam.Add("SETID", param.Value)

        param = cmd2.Parameters.Add("EnterDate", SqlDbType.DateTime)
        param.Value = rqEnterDate 'CDate(Now()).ToString("yyyy/MM/dd")
        Call TIMS.Set_STR2STR(strfield, param.Value, "EnterDate")
        myParam.Add("EnterDate", param.Value)

        param = cmd2.Parameters.Add("SerNum", SqlDbType.Int)
        param.Value = vSerNum
        Call TIMS.Set_STR2STR(strfield, param.Value, "SerNum")
        myParam.Add("SerNum", param.Value)

        Try
            'cmd2.ExecuteNonQuery()
            DbAccess.ExecuteNonQuery(sql, myParam)
        Catch ex As Exception
            strErrmsg = ex.ToString & vbCrLf '偵錯用儲存欄
            If Not mainTrain Is Nothing Then mainTrain.Rollback()
            Call TIMS.CloseDbConn(conn)
            Common.MessageBox(Me, strErrmsg)

            '偵錯用儲存欄
            strErrmsg += "/* strSql2: */" & vbCrLf
            strErrmsg += strSql2 & vbCrLf
            strErrmsg += "/* field: */" & vbCrLf
            strErrmsg += strfield & vbCrLf
            strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)
            Throw ex
        End Try
    End Sub

    Sub SUtl_UPDATE_STUD_ENTERTEMP(ByVal conn As SqlConnection, ByVal mainTrain As SqlTransaction)
        'Dim MyKey As String=""
        Dim strErrmsg As String = ""  '偵錯用儲存欄
        Dim strfield As String = ""   '偵錯用儲存欄
        Dim strSql2 As String = ""    '偵錯用儲存欄
        Dim cmd As New SqlCommand     '建立STUD_ENTERTYPE專用command
        Dim param As SqlParameter
        Dim myParam As Hashtable = New Hashtable
        Dim sql As String = ""
        sql &= " UPDATE Stud_EnterTemp" & vbCrLf
        sql &= " SET Name=@Name, Birthday=@Birthday, PassPortNO=@PassPortNO" & vbCrLf
        sql &= " ,Sex=@Sex, MaritalStatus=@MaritalStatus, DegreeID=@DegreeID, GradID=@GradID" & vbCrLf
        sql &= " ,School=@School, Department=@Department, MilitaryID=@MilitaryID ,ZipCode=@ZipCode ,ZipCODE6W=@ZipCODE6W" & vbCrLf
        sql &= " ,Address=@Address, Phone1=@Phone1, Phone2=@Phone2, Email=@Email, CellPhone=@CellPhone" & vbCrLf
        sql &= " ,IsAgree=@IsAgree, ModifyAcct=@ModifyAcct, ModifyDate=getdate()" & vbCrLf
        sql &= " where IDNO=@IDNO "

        strErrmsg = "" '偵錯用儲存欄
        strSql2 = sql '偵錯用儲存欄
        strfield = "" '偵錯用儲存欄
        cmd = New SqlCommand(sql, conn, mainTrain)

        param = cmd.Parameters.Add("Name", SqlDbType.NVarChar, 30)
        param.Value = Name.Text
        Call TIMS.Set_STR2STR(strfield, param.Value, "Name")
        myParam.Add("Name", param.Value)

        param = cmd.Parameters.Add("Birthday", SqlDbType.DateTime)
        param.Value = TIMS.Cdate2(birthday.Text)
        Call TIMS.Set_STR2STR(strfield, param.Value, "Birthday")
        myParam.Add("Birthday", param.Value)

        param = cmd.Parameters.Add("PassPortNO", SqlDbType.Int, 4)
        Select Case PassPortNO.SelectedValue
            Case "1", "2"
                param.Value = Val(PassPortNO.SelectedValue)
            Case Else
                param.Value = 2
        End Select
        Call TIMS.Set_STR2STR(strfield, param.Value, "PassPortNO")
        myParam.Add("PassPortNO", param.Value)

        param = cmd.Parameters.Add("Sex", SqlDbType.Char, 1)
        Select Case Sex.SelectedValue
            Case "M", "F"
                param.Value = Sex.SelectedValue
            Case Else
                param.Value = TIMS.GetMemberSex(IDNO.Text)
        End Select
        Call TIMS.Set_STR2STR(strfield, param.Value, "Sex")
        myParam.Add("Sex", param.Value)

        param = cmd.Parameters.Add("MaritalStatus", SqlDbType.Int)
        Dim v_MaritalStatus As String = TIMS.GetListValue(MaritalStatus)
        Dim MaritalStatuslist As New List(Of String) From {"1", "2", "3"} '未填補0或NULL
        param.Value = If(MaritalStatuslist.Contains(v_MaritalStatus), Val(v_MaritalStatus), 0) : Call TIMS.Set_STR2STR(strfield, param.Value, "MaritalStatus")
        myParam.Add("MaritalStatus", param.Value)

        param = cmd.Parameters.Add("DegreeID", SqlDbType.VarChar, 3)
        param.Value = DegreeID.SelectedValue
        Call TIMS.Set_STR2STR(strfield, param.Value, "DegreeID")
        myParam.Add("DegreeID", param.Value)

        param = cmd.Parameters.Add("GradID", SqlDbType.VarChar, 3)
        param.Value = GradID.SelectedValue
        Call TIMS.Set_STR2STR(strfield, param.Value, "GradID")
        myParam.Add("GradID", param.Value)

        param = cmd.Parameters.Add("School", SqlDbType.NVarChar)
        param.Value = School.Text
        Call TIMS.Set_STR2STR(strfield, param.Value, "School")
        myParam.Add("School", param.Value)

        param = cmd.Parameters.Add("Department", SqlDbType.NVarChar)
        param.Value = Department.Text
        Call TIMS.Set_STR2STR(strfield, param.Value, "Department")
        myParam.Add("Department", param.Value)

        param = cmd.Parameters.Add("MilitaryID", SqlDbType.VarChar)
        param.Value = If(MilitaryID.SelectedValue = "", Convert.DBNull, MilitaryID.SelectedValue) '
        Call TIMS.Set_STR2STR(strfield, param.Value, "MilitaryID")
        myParam.Add("MilitaryID", param.Value)

        param = cmd.Parameters.Add("ZipCode", SqlDbType.Int)
        param.Value = If(city_code.Value <> "", Val(city_code.Value), Convert.DBNull)
        Call TIMS.Set_STR2STR(strfield, param.Value, "ZipCode")
        myParam.Add("ZipCode", param.Value)
        hidZipCODE6W.Value = TIMS.GetZIPCODE6W(city_code.Value, ZipCODEB3.Value)
        param = cmd.Parameters.Add("ZipCODE6W", SqlDbType.VarChar)
        param.Value = If(hidZipCODE6W.Value <> "", hidZipCODE6W.Value, Convert.DBNull)
        Call TIMS.Set_STR2STR(strfield, param.Value, "ZipCODE6W")
        myParam.Add("ZipCODE6W", param.Value)

        param = cmd.Parameters.Add("Address", SqlDbType.NVarChar)
        param.Value = Address.Text
        Call TIMS.Set_STR2STR(strfield, param.Value, "Address")
        myParam.Add("Address", param.Value)

        param = cmd.Parameters.Add("Phone1", SqlDbType.VarChar)
        param.Value = Phone1.Text
        Call TIMS.Set_STR2STR(strfield, param.Value, "Phone1")
        myParam.Add("Phone1", param.Value)

        param = cmd.Parameters.Add("Phone2", SqlDbType.VarChar)
        param.Value = Phone2.Text
        Call TIMS.Set_STR2STR(strfield, param.Value, "Phone2")
        myParam.Add("Phone2", param.Value)

        param = cmd.Parameters.Add("Email", SqlDbType.VarChar)
        param.Value = Email.Text
        Call TIMS.Set_STR2STR(strfield, param.Value, "Email")
        myParam.Add("Email", param.Value)

        param = cmd.Parameters.Add("CellPhone", SqlDbType.VarChar)
        param.Value = CellPhone.Text
        Call TIMS.Set_STR2STR(strfield, param.Value, "CellPhone")
        myParam.Add("CellPhone", param.Value)

        'param=cmd.Parameters.Add("Notes", SqlDbType.Text)
        'param.Value=notes.Text
        param = cmd.Parameters.Add("IsAgree", SqlDbType.Char)
        param.Value = "Y" 'IsAgree.SelectedValue
        Call TIMS.Set_STR2STR(strfield, param.Value, "IsAgree")
        myParam.Add("IsAgree", param.Value)

        param = cmd.Parameters.Add("ModifyAcct", SqlDbType.VarChar, 30)
        param.Value = sm.UserInfo.UserID
        Call TIMS.Set_STR2STR(strfield, param.Value, "ModifyAcct")
        myParam.Add("ModifyAcct", param.Value)

        'If IDNO.Text <> "" Then IDNO.Text=Trim(IDNO.Text) 'If IDNO.Text <> "" Then IDNO.Text=UCase(IDNO.Text)
        IDNO.Text = TIMS.ChangeIDNO(IDNO.Text)
        param = cmd.Parameters.Add("IDNO", SqlDbType.VarChar, 15)
        param.Value = IDNO.Text
        Call TIMS.Set_STR2STR(strfield, param.Value, "IDNO")
        myParam.Add("IDNO", param.Value)

        Try
            'cmd.ExecuteNonQuery()
            DbAccess.ExecuteNonQuery(sql, myParam)
        Catch ex As Exception
            strErrmsg = ex.ToString & vbCrLf '偵錯用儲存欄
            If Not mainTrain Is Nothing Then mainTrain.Rollback()
            Call TIMS.CloseDbConn(conn)
            Common.MessageBox(Me, strErrmsg)

            '偵錯用儲存欄
            strErrmsg += "/* strSql2: */" & vbCrLf
            strErrmsg += strSql2 & vbCrLf
            strErrmsg += "/* field: */" & vbCrLf
            strErrmsg += strfield & vbCrLf
            strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)

            Throw ex
        End Try
    End Sub

    Sub SUtl_INSERT_STUD_ENTERTEMP(ByVal ivsSETID As Integer, ByVal conn As SqlConnection, ByVal mainTrain As SqlTransaction)
        'Dim MyKey As String=""
        'select * from Stud_EnterTemp where SETID<=0
        'If ivsSETID=0 Then Exit Sub
        If ivsSETID <= 0 Then Return '只能新增大於0的資料Exit Sub 
        Dim strErrmsg As String = ""           '偵錯用儲存欄
        Dim strfield As String = ""            '偵錯用儲存欄
        Dim strSql2 As String = ""             '偵錯用儲存欄
        Dim cmd As New SqlCommand              '建立STUD_ENTERTYPE專用command
        Dim param As SqlParameter
        Dim myParam As Hashtable = New Hashtable
        Dim sql As String = ""

        sql = "" & vbCrLf
        sql &= " INSERT INTO Stud_EnterTemp(SETID, Name,Birthday,PassPortNO,IDNO,Sex" & vbCrLf
        sql &= " ,MaritalStatus,DegreeID,GradID,School,Department,MilitaryID,ZipCode,ZipCODE6W" & vbCrLf
        sql &= " ,Address,Phone1,Phone2,Email,CellPhone, IsAgree,ModifyAcct,ModifyDate)" & vbCrLf
        sql &= " VALUES (@SETID, @Name, @Birthday, @PassPortNO, @IDNO, @Sex" & vbCrLf
        sql &= " ,@MaritalStatus, @DegreeID, @GradID, @School, @Department, @MilitaryID, @ZipCode, @ZipCODE6W" & vbCrLf
        sql &= " ,@Address, @Phone1, @Phone2, @Email, @CellPhone, @IsAgree, @ModifyAcct, GETDATE())" & vbCrLf

        strErrmsg = "" '偵錯用儲存欄
        strSql2 = sql '偵錯用儲存欄
        strfield = "" '偵錯用儲存欄
        cmd = New SqlCommand(sql, conn, mainTrain)

        param = cmd.Parameters.Add("SETID", SqlDbType.Int)
        param.Value = ivsSETID
        Call TIMS.Set_STR2STR(strfield, param.Value, "SETID")
        myParam.Add("SETID", param.Value)

        param = cmd.Parameters.Add("Name", SqlDbType.NVarChar, 30)
        param.Value = Name.Text
        Call TIMS.Set_STR2STR(strfield, param.Value, "Name")
        myParam.Add("Name", param.Value)

        param = cmd.Parameters.Add("Birthday", SqlDbType.DateTime)
        param.Value = birthday.Text
        Call TIMS.Set_STR2STR(strfield, param.Value, "Birthday")
        myParam.Add("Birthday", param.Value)

        param = cmd.Parameters.Add("PassPortNO", SqlDbType.Int, 4)
        Select Case PassPortNO.SelectedValue
            Case "1", "2"
                param.Value = Val(PassPortNO.SelectedValue)
            Case Else
                param.Value = 2 'Val("2")
        End Select
        Call TIMS.Set_STR2STR(strfield, param.Value, "PassPortNO")
        myParam.Add("PassPortNO", param.Value)

        'If IDNO.Text <> "" Then IDNO.Text=Trim(IDNO.Text) 'If IDNO.Text <> "" Then IDNO.Text=UCase(IDNO.Text)
        IDNO.Text = TIMS.ChangeIDNO(IDNO.Text)
        param = cmd.Parameters.Add("IDNO", SqlDbType.VarChar, 15)
        param.Value = IDNO.Text
        Call TIMS.Set_STR2STR(strfield, param.Value, "IDNO")
        myParam.Add("IDNO", param.Value)

        param = cmd.Parameters.Add("Sex", SqlDbType.Char, 1)
        param.Value = Sex.SelectedValue
        Call TIMS.Set_STR2STR(strfield, param.Value, "Sex")
        myParam.Add("Sex", param.Value)

        param = cmd.Parameters.Add("MaritalStatus", SqlDbType.Int)
        Dim v_MaritalStatus As String = TIMS.GetListValue(MaritalStatus)
        Dim MaritalStatuslist As New List(Of String) From {"1", "2", "3"} '未填補0或NULL
        param.Value = If(MaritalStatuslist.Contains(v_MaritalStatus), Val(v_MaritalStatus), 0) : Call TIMS.Set_STR2STR(strfield, param.Value, "MaritalStatus")
        myParam.Add("MaritalStatus", param.Value)

        param = cmd.Parameters.Add("DegreeID", SqlDbType.VarChar, 3)
        param.Value = DegreeID.SelectedValue
        Call TIMS.Set_STR2STR(strfield, param.Value, "DegreeID")
        myParam.Add("DegreeID", param.Value)

        param = cmd.Parameters.Add("GradID", SqlDbType.VarChar, 3)
        param.Value = GradID.SelectedValue
        Call TIMS.Set_STR2STR(strfield, param.Value, "GradID")
        myParam.Add("GradID", param.Value)

        param = cmd.Parameters.Add("School", SqlDbType.NVarChar)
        param.Value = School.Text
        Call TIMS.Set_STR2STR(strfield, param.Value, "School")
        myParam.Add("School", param.Value)

        param = cmd.Parameters.Add("Department", SqlDbType.NVarChar)
        param.Value = Department.Text
        Call TIMS.Set_STR2STR(strfield, param.Value, "Department")
        myParam.Add("Department", param.Value)

        param = cmd.Parameters.Add("MilitaryID", SqlDbType.VarChar)
        param.Value = If(MilitaryID.SelectedValue = "", Convert.DBNull, MilitaryID.SelectedValue) 'MilitaryID.SelectedValue
        Call TIMS.Set_STR2STR(strfield, param.Value, "MilitaryID")
        myParam.Add("MilitaryID", param.Value)

        param = cmd.Parameters.Add("ZipCode", SqlDbType.Int)
        param.Value = If(city_code.Value <> "", Val(city_code.Value), Convert.DBNull)
        Call TIMS.Set_STR2STR(strfield, param.Value, "ZipCode")
        myParam.Add("ZipCode", param.Value)
        hidZipCODE6W.Value = TIMS.GetZIPCODE6W(city_code.Value, ZipCODEB3.Value)
        param = cmd.Parameters.Add("ZipCODE6W", SqlDbType.VarChar)
        param.Value = If(hidZipCODE6W.Value <> "", hidZipCODE6W.Value, Convert.DBNull)
        Call TIMS.Set_STR2STR(strfield, param.Value, "ZipCODE6W")
        myParam.Add("ZipCODE6W", param.Value)

        param = cmd.Parameters.Add("Address", SqlDbType.NVarChar)
        param.Value = Address.Text
        Call TIMS.Set_STR2STR(strfield, param.Value, "Address")
        myParam.Add("Address", param.Value)

        param = cmd.Parameters.Add("Phone1", SqlDbType.VarChar)
        param.Value = Phone1.Text
        Call TIMS.Set_STR2STR(strfield, param.Value, "Phone1")
        myParam.Add("Phone1", param.Value)

        param = cmd.Parameters.Add("Phone2", SqlDbType.VarChar)
        param.Value = Phone2.Text
        Call TIMS.Set_STR2STR(strfield, param.Value, "Phone2")
        myParam.Add("Phone2", param.Value)

        param = cmd.Parameters.Add("Email", SqlDbType.VarChar)
        param.Value = Email.Text
        Call TIMS.Set_STR2STR(strfield, param.Value, "Email")
        myParam.Add("Email", param.Value)

        param = cmd.Parameters.Add("CellPhone", SqlDbType.VarChar)
        param.Value = CellPhone.Text
        Call TIMS.Set_STR2STR(strfield, param.Value, "CellPhone")
        myParam.Add("CellPhone", param.Value)

        param = cmd.Parameters.Add("IsAgree", SqlDbType.Char)
        param.Value = "Y" 'IsAgree.SelectedValue
        Call TIMS.Set_STR2STR(strfield, param.Value, "IsAgree")
        myParam.Add("IsAgree", param.Value)

        param = cmd.Parameters.Add("ModifyAcct", SqlDbType.VarChar, 30)
        param.Value = sm.UserInfo.UserID
        Call TIMS.Set_STR2STR(strfield, param.Value, "ModifyAcct")
        myParam.Add("ModifyAcct", param.Value)

        Try
            'cmd.ExecuteNonQuery()
            DbAccess.ExecuteNonQuery(sql, myParam)
        Catch ex As Exception
            strErrmsg = ex.ToString & vbCrLf '偵錯用儲存欄
            If Not mainTrain Is Nothing Then mainTrain.Rollback()
            TIMS.CloseDbConn(conn)
            'If conn Is Nothing Then Exit Sub
            'If conn.State=ConnectionState.Open Then conn.Close()
            Common.MessageBox(Me, strErrmsg)

            '偵錯用儲存欄
            strErrmsg += "/* strSql2: */" & vbCrLf
            strErrmsg += strSql2 & vbCrLf
            strErrmsg += "/* field: */" & vbCrLf
            strErrmsg += strfield & vbCrLf
            strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)

            Throw ex
        End Try
    End Sub

    Sub SUtl_UPDATE_STUD_ENTERTEMP2(ByVal conn As SqlConnection, ByVal mainTrain As SqlTransaction)
        Dim strErrmsg As String = ""            '偵錯用儲存欄
        Dim strfield As String = ""             '偵錯用儲存欄
        Dim strSql2 As String = ""              '偵錯用儲存欄
        Dim cmd9 As New SqlCommand              '建立STUD_ENTERTYPE專用command
        Dim param As SqlParameter
        Dim myParam As Hashtable = New Hashtable
        Dim sql As String = ""

        sql = ""
        sql &= " UPDATE Stud_EnterTemp2"
        sql &= " SET Name=@Name ,Birthday=@Birthday ,PassPortNO=@PassPortNO "
        sql &= " ,Sex=@Sex ,MaritalStatus=@MaritalStatus ,DegreeID=@DegreeID ,GradID=@GradID "
        sql &= " ,School=@School ,Department=@Department ,MilitaryID=@MilitaryID ,ZipCode=@ZipCode,ZipCODE6W=@ZipCODE6W  "
        sql &= " ,Address=@Address ,Phone1=@Phone1 ,Phone2=@Phone2 ,Email=@Email ,CellPhone=@CellPhone "
        sql &= " ,IsAgree=@IsAgree ,ModifyAcct=@ModifyAcct ,ModifyDate=GETDATE() "
        sql &= " WHERE IDNO=@IDNO " '" & TIMS.ChangeIDNO(IDNO.Text.ToString) & "'"

        strErrmsg = "" '偵錯用儲存欄
        strSql2 = sql '偵錯用儲存欄
        strfield = "" '偵錯用儲存欄
        cmd9 = New SqlCommand(sql, conn, mainTrain)

        'param=cmd9.Parameters.Add("SETID", SqlDbType.Int, 8)
        'param.Value=Me.ivsSETID
        'Call TIMS.Set_STR2STR(strfield, param.Value, "SETID")

        param = cmd9.Parameters.Add("Name", SqlDbType.NVarChar, 30)
        param.Value = Name.Text
        Call TIMS.Set_STR2STR(strfield, param.Value, "Name")
        myParam.Add("Name", param.Value)

        param = cmd9.Parameters.Add("Birthday", SqlDbType.DateTime)
        param.Value = birthday.Text
        Call TIMS.Set_STR2STR(strfield, param.Value, "Birthday")
        myParam.Add("Birthday", param.Value)

        param = cmd9.Parameters.Add("PassPortNO", SqlDbType.Int)
        Select Case PassPortNO.SelectedValue
            Case "1", "2"
                param.Value = Val(PassPortNO.SelectedValue)
            Case Else
                param.Value = 2 '"2"
        End Select
        Call TIMS.Set_STR2STR(strfield, param.Value, "PassPortNO")
        myParam.Add("PassPortNO", param.Value)

        param = cmd9.Parameters.Add("Sex", SqlDbType.Char, 1)
        param.Value = Sex.SelectedValue
        Call TIMS.Set_STR2STR(strfield, param.Value, "Sex")
        myParam.Add("Sex", param.Value)

        param = cmd9.Parameters.Add("MaritalStatus", SqlDbType.Int)
        Dim v_MaritalStatus As String = TIMS.GetListValue(MaritalStatus)
        Dim MaritalStatuslist As New List(Of String) From {"1", "2", "3"} '未填補0或NULL
        param.Value = If(MaritalStatuslist.Contains(v_MaritalStatus), Val(v_MaritalStatus), 0) : Call TIMS.Set_STR2STR(strfield, param.Value, "MaritalStatus")
        myParam.Add("MaritalStatus", param.Value)

        param = cmd9.Parameters.Add("DegreeID", SqlDbType.VarChar, 3)
        param.Value = DegreeID.SelectedValue : Call TIMS.Set_STR2STR(strfield, param.Value, "DegreeID")
        myParam.Add("DegreeID", param.Value)

        param = cmd9.Parameters.Add("GradID", SqlDbType.VarChar, 3)
        param.Value = GradID.SelectedValue : Call TIMS.Set_STR2STR(strfield, param.Value, "GradID")
        myParam.Add("GradID", param.Value)

        param = cmd9.Parameters.Add("School", SqlDbType.NVarChar)
        param.Value = School.Text : Call TIMS.Set_STR2STR(strfield, param.Value, "School")
        myParam.Add("School", param.Value)

        param = cmd9.Parameters.Add("Department", SqlDbType.NVarChar)
        param.Value = Department.Text : Call TIMS.Set_STR2STR(strfield, param.Value, "Department")
        myParam.Add("Department", param.Value)

        Dim v_MilitaryID As String = TIMS.GetListValue(MilitaryID)
        param = cmd9.Parameters.Add("MilitaryID", SqlDbType.VarChar)
        param.Value = If(v_MilitaryID = "", Convert.DBNull, v_MilitaryID) 'MilitaryID.SelectedValue
        Call TIMS.Set_STR2STR(strfield, param.Value, "MilitaryID")
        myParam.Add("MilitaryID", param.Value)

        param = cmd9.Parameters.Add("ZipCode", SqlDbType.Int)
        param.Value = If(city_code.Value <> "", Val(city_code.Value), Convert.DBNull)
        Call TIMS.Set_STR2STR(strfield, param.Value, "ZipCode")
        myParam.Add("ZipCode", param.Value)
        hidZipCODE6W.Value = TIMS.GetZIPCODE6W(city_code.Value, ZipCODEB3.Value)
        param = cmd9.Parameters.Add("ZipCODE6W", SqlDbType.VarChar)
        param.Value = If(hidZipCODE6W.Value <> "", hidZipCODE6W.Value, Convert.DBNull)
        Call TIMS.Set_STR2STR(strfield, param.Value, "ZipCODE6W")
        myParam.Add("ZipCODE6W", param.Value)

        param = cmd9.Parameters.Add("Address", SqlDbType.NVarChar)
        param.Value = Address.Text
        Call TIMS.Set_STR2STR(strfield, param.Value, "Address")
        myParam.Add("Address", param.Value)

        param = cmd9.Parameters.Add("Phone1", SqlDbType.VarChar)
        param.Value = Phone1.Text
        Call TIMS.Set_STR2STR(strfield, param.Value, "Phone1")
        myParam.Add("Phone1", param.Value)

        param = cmd9.Parameters.Add("Phone2", SqlDbType.VarChar)
        param.Value = Phone2.Text
        Call TIMS.Set_STR2STR(strfield, param.Value, "Phone2")
        myParam.Add("Phone2", param.Value)

        param = cmd9.Parameters.Add("Email", SqlDbType.VarChar)
        param.Value = Email.Text
        Call TIMS.Set_STR2STR(strfield, param.Value, "Email")
        myParam.Add("Email", param.Value)

        param = cmd9.Parameters.Add("CellPhone", SqlDbType.VarChar)
        param.Value = CellPhone.Text
        Call TIMS.Set_STR2STR(strfield, param.Value, "CellPhone")
        myParam.Add("CellPhone", param.Value)

        'param=cmd9.Parameters.Add("Notes", SqlDbType.Text)
        'param.Value=notes.Text
        param = cmd9.Parameters.Add("IsAgree", SqlDbType.Char)
        param.Value = "Y" 'IsAgree.SelectedValue
        Call TIMS.Set_STR2STR(strfield, param.Value, "IsAgree")
        myParam.Add("IsAgree", param.Value)

        param = cmd9.Parameters.Add("ModifyAcct", SqlDbType.VarChar, 30)
        param.Value = sm.UserInfo.UserID
        Call TIMS.Set_STR2STR(strfield, param.Value, "ModifyAcct")
        myParam.Add("ModifyAcct", param.Value)

        'If IDNO.Text <> "" Then IDNO.Text = Trim(IDNO.Text) 'If IDNO.Text <> "" Then IDNO.Text = UCase(IDNO.Text)
        IDNO.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(IDNO.Text))
        param = cmd9.Parameters.Add("IDNO", SqlDbType.VarChar, 15)
        param.Value = IDNO.Text
        Call TIMS.Set_STR2STR(strfield, param.Value, "IDNO")
        myParam.Add("IDNO", param.Value)

        Try
            'cmd9.ExecuteNonQuery()
            DbAccess.ExecuteNonQuery(sql, myParam)
        Catch ex As Exception
            strErrmsg = ex.ToString & vbCrLf '偵錯用儲存欄
            If Not mainTrain Is Nothing Then mainTrain.Rollback()
            Call TIMS.CloseDbConn(conn)
            Common.MessageBox(Me, strErrmsg)

            '偵錯用儲存欄
            strErrmsg += "/* strSql2: */" & vbCrLf
            strErrmsg += strSql2 & vbCrLf
            strErrmsg += "/* field: */" & vbCrLf
            strErrmsg += strfield & vbCrLf
            strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)

            Throw ex
        End Try
    End Sub

    '勾選身分別 (中高齡者) 且傳回 出生日期是否有誤 正確@True; 錯誤@False;
    Public Shared Function Over_45YearOld(ByVal OCID As String, ByVal sBirthDay As String, ByRef objIdent As CheckBoxList, ByRef sErrmsg As String, ByVal tConn As SqlConnection) As Boolean
        Dim Rst As Boolean = False '有誤
        sErrmsg = ""

        Try
            sBirthDay = CDate(sBirthDay).ToString("yyyy/MM/dd")
            If Not TIMS.IsDate1(sBirthDay) Then
                Rst = False '有誤
                sErrmsg += "出生日期格式有誤，請確認" & vbCrLf
            End If
        Catch ex As Exception
            Rst = False '有誤
            sErrmsg += "出生日期格式有誤，請確認" & vbCrLf
        End Try

        Dim sqlstr As String = ""
        Dim STDateTemp As String = ""
        If OCID = "" Then
            Rst = False '有誤
            sErrmsg += "班級選擇有誤，請確認" & vbCrLf
        Else
            '取出報名志願一班別的開訓日期
            sqlstr = "SELECT CONVERT(varchar, STDate, 111) STDate FROM Class_ClassInfo WHERE OCID='" & OCID & "'"
            STDateTemp = Convert.ToString(DbAccess.ExecuteScalar(sqlstr, tConn))
            If STDateTemp = "" Then
                Rst = False '有誤
                sErrmsg += "班級選擇有誤，請確認，開訓日期" & vbCrLf
            End If
        End If
        If sErrmsg = "" Then
            Rst = True '正常
            Dim STDateTemp45 As String = ""
            Dim STDateTemp65 As String = ""
            If STDateTemp <> "" Then STDateTemp = CDate(STDateTemp).ToString("yyyy/MM/dd")

            '因為2009/06/01 拿掉負擔家計負女,所以選項少1
            '足歲45歲則選取中高齡者(45歲)。跟開訓日期比較
            '未滿65歲則選取中高齡者。跟開訓日期比較
            STDateTemp45 = CDate(DateAdd(DateInterval.Day, -45, CDate(STDateTemp))).ToString("yyyy/MM/dd")
            STDateTemp65 = CDate(DateAdd(DateInterval.Day, -65, CDate(STDateTemp))).ToString("yyyy/MM/dd")

            If DateDiff(DateInterval.Year, CDate(sBirthDay), CDate(STDateTemp45)) >= 0 _
                AndAlso DateDiff(DateInterval.Year, CDate(sBirthDay), CDate(STDateTemp65)) <= 0 Then
                For i As Integer = 0 To objIdent.Items.Count - 1
                    If objIdent.Items(i).Value = "04" Then
                        objIdent.Items(i).Selected = True
                        Exit For
                    End If
                Next
            End If
        End If

        Return Rst
    End Function

    '課程階段顯示
    Sub SUtl_ShowCL()
        If CCLID.Value = "" Then Exit Sub
        If OCIDValue1.Value = "" Then Exit Sub
        If OCID1.Text = "" Then Exit Sub

        Dim pms1 As New Hashtable From {
            {"OCID", TIMS.CINT1(OCIDValue1.Value)},
            {"CCLID", CCLID.Value}
        }
        Dim sql As String = ""
        sql &= " SELECT a.ClassCName ,a.CyclType ,b.LevelName" & vbCrLf
        sql &= " FROM Class_ClassInfo a" & vbCrLf
        sql &= " JOIN Class_ClassLevel b ON a.OCID=b.OCID" & vbCrLf
        sql &= " WHERE a.OCID=@OCID AND b.CCLID=@CCLID" & vbCrLf
        Dim drCL As DataRow = DbAccess.GetOneRow(sql, objconn, pms1)

        If drCL IsNot Nothing Then
            Dim sCLASSN As String = OCID1.Text
            Dim sMVALUE As String = ""
            Select Case Int(drCL("LevelName"))
                Case 1
                    sMVALUE = sCLASSN & "(第一階段)"
                    'OCID1.Text += "(第一階段)"
                Case 2
                    sMVALUE = sCLASSN & "(第二階段)"
                    'OCID1.Text += "(第二階段)"
                Case 3
                    sMVALUE = sCLASSN & "(第三階段)"
                    'OCID1.Text += "(第三階段)"
                Case 4
                    sMVALUE = sCLASSN & "(第四階段)"
                    'OCID1.Text += "(第四階段)"
                Case 5
                    sMVALUE = sCLASSN & "(第五階段)"
                    'OCID1.Text += "(第五階段)"
                Case Else
                    sMVALUE = sCLASSN
            End Select
            OCID1.Text = sMVALUE
        End If
    End Sub

    '查詢歷史紀錄 Button9_Click
    'Protected Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
    'End Sub

    '送出 / Button1_Click /Button7_ServerClick
    'Private Sub Button7_ServerClick(sender As Object, e As EventArgs) Handles Button7.ServerClick
    'End Sub
End Class
