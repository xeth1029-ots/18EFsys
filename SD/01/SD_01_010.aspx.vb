Partial Class SD_01_010
    Inherits AuthBasePage

#Region "參數/變數"

    Const cst_msgSuper1 As String = "該使用者，可強行報名!!!"

    'Dim blnCanAdds As Boolean=False '新增
    'Dim blnCanMod As Boolean=False '修改
    'Dim blnCanDel As Boolean=False '刪除
    'Dim blnCanSech As Boolean=False '查詢
    'Dim blnCanPrnt As Boolean=False '列印
    'Dim FunDr As DataRow
    'Dim TestStr As String="" '測試用
    Dim objconn As SqlConnection

#End Region

    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload

        'Call TIMS.GetAuth(Me, blnCanAdds, blnCanMod, blnCanDel, blnCanSech, blnCanPrnt) '2011 取得功能按鈕權限值

        'btnAdd1.Enabled=False
        'If blnCanAdds Then btnAdd1.Enabled=True
        'If Not btnAdd1.Enabled Then TIMS.Tooltip(btnAdd1, "權限無法使用 報名功能")

#Region "(No Use)"

        'Dim sql As String=""
        ''Dim dr As DataRow
        'Try
        '    If sm.UserInfo.FunDt Is Nothing Then
        '        Common.RespWrite(Me, "<script>alert('Session過期');</script>")
        '        Common.RespWrite(Me, "<script>top.location.href='../../logout.aspx';</script>")
        '    Else
        '        Dim FunDt As DataTable=sm.UserInfo.FunDt
        '        Dim FunDrArray() As DataRow=FunDt.Select("FunID='" & Request("ID") & "'")
        '        If FunDrArray.Length=0 Then
        '            Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
        '            Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
        '        Else
        '            FunDr=FunDrArray(0)
        '            If btnAdd1.Enabled Then
        '                If FunDr("Adds")="1" Then btnAdd1.Enabled=True Else btnAdd1.Enabled=False
        '            End If
        '            'If btnView1.Enabled Then
        '            '    If FunDr("Sech")="1" Then btnView1.Enabled=True Else btnView1.Enabled=False
        '            'End If
        '        End If
        '    End If
        'Catch ex As Exception
        '    Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
        '    Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
        'End Try

#End Region

        Dim sAltMsg As String = "" '訊息
        Dim flag_stopEnter2 As Boolean = TIMS.StopEnterTempMsg2(objconn, sAltMsg)
        If flag_stopEnter2 Then
            Common.MessageBox(Me, sAltMsg)
            Exit Sub
        End If

        If Not IsPostBack Then
            btnAdd1.Enabled = False '(請先確認狀態後開啟此按鍵)
            TIMS.Tooltip(btnAdd1, "(請先確認狀態後開啟此按鍵)")

            If TIMS.StopEnterTempMsg1(Me, objconn, True) Then Exit Sub

            If Session("_SearchStr") IsNot Nothing Then
                ViewState("_SearchStr") = Session("_SearchStr")
                Session("_SearchStr") = Nothing
                center.Text = TIMS.GetMyValue(ViewState("_SearchStr"), "center")
                RIDValue.Value = TIMS.GetMyValue(ViewState("_SearchStr"), "RIDValue")
                OCID1.Text = TIMS.GetMyValue(ViewState("_SearchStr"), "OCID1")
                TMID1.Text = TIMS.GetMyValue(ViewState("_SearchStr"), "TMID1")
                OCIDValue1.Value = TIMS.GetMyValue(ViewState("_SearchStr"), "OCIDValue1")
                TMIDValue1.Value = TIMS.GetMyValue(ViewState("_SearchStr"), "TMIDValue1")
                IDNO.Text = TIMS.GetMyValue(ViewState("_SearchStr"), "IDNO")
                birthDay.Text = TIMS.GetMyValue(ViewState("_SearchStr"), "birthDay")
                'EnterDate.Text=TIMS.GetMyValue(ViewState("_SearchStr"), "EnterDate")
            Else
                'EnterDate.Text=Now.Date
                center.Text = sm.UserInfo.OrgName
                RIDValue.Value = sm.UserInfo.RID
            End If
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx?BtnName=btnCheckClass');"
        Button8.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, Historytable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1", True, "btnCheckClass")
        If Historytable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

        Button5.Attributes("onclick") = "choose_class();"
        btnAdd1.Attributes.Add("OnClick", "if( confirm ('資料確認無誤，確定送出？')) { return ChkData(); }")
        IDNO.Attributes.Add("onclick", "sChkData1();")
        IDNO.Attributes.Add("onblur", "sChkData1();")
        IDNO.Attributes.Add("onFocus", "sChkData1();")
        IDNO.Attributes.Add("onKeyPress", "sChkData1();")
        IDNO.Attributes.Add("onKeyDown", "sChkData1();")

        '確認機構是否為黑名單
        Dim vsMsg2 As String = "" '確認機構是否為黑名單
        vsMsg2 = ""
        If Chk_OrgBlackList(vsMsg2) Then
            btnCheckClass.Enabled = False
            TIMS.Tooltip(btnCheckClass, vsMsg2)
            btnAdd1.Enabled = False
            TIMS.Tooltip(btnAdd1, vsMsg2)

            btnCheck1.Enabled = False
            TIMS.Tooltip(btnCheck1, vsMsg2)
            Dim vsStrScript As String = $"<script>alert('{vsMsg2}');</script>"
            Page.RegisterStartupScript("", vsStrScript)
        End If
    End Sub

    '機構黑名單內容(訓練單位處分功能)
    Function Chk_OrgBlackList(ByRef Errmsg As String) As Boolean
        Dim rst As Boolean = False
        Errmsg = ""
        Dim vsComIDNO As String = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, objconn)
        If TIMS.Check_OrgBlackList(Me, vsComIDNO, objconn) Then
            rst = True
            Errmsg = $"{sm.UserInfo.OrgName}，已列入處分名單!!"
            isBlack.Value = "Y"
            Blackorgname.Value = sm.UserInfo.OrgName
        End If
        Return rst
    End Function

    '輔助金使用檢核
    Function Check_GovCost(ByVal aIDNO As String, ByVal OCID1 As String, ByVal Redirect1 As String, ByVal Redirect2 As String, ByRef JaveScriptAlertMsg As String, ByRef AlertType As String) As Boolean
        'Optional ByRef JaveScriptAlertMsg As String="", Optional ByRef AlertType As String="1"

        Dim Check_GovCostFlag As Boolean = True 'True-正常:可以用補助/異常:補助額不足，將另尋解決途徑
        'AlertType 1-正常:可以用補助(進入繼續報名程序)
        'AlertType 2-異常:補助額度不足，但仍有部份補助額(進入繼續報名程序／終止報名程序)
        'AlertType 3-異常:補助額度已滿(終止報名程序)

        Dim TxtMessage As String 'message
        ''(限定產業人才投資方案) 20090325 BY AMU
        ''Dim ActSubsidyCost As String=TIMS.Get_ActSubsidyCost28(IDNO) '已實際請領補助費(限定產業人才投資方案)
        ''Dim SignSubsidyCost As String=TIMS.Get_SignSubsidyCost28(IDNO) '已報名申請補助費(限定產業人才投資方案)
        ''Dim DefGovCost As String=TIMS.Get_DefGeoCost28(IDNO) '(=線上報名預算的政府補助)(全部)(限定產業人才投資方案)

        'Const Cst_MaxCanUseCost=50000 '三年內最大可用餘額2007年前為３萬, 2008年改為５萬
        'Const Cst_AlertCost=40000 '警示額
        Dim LOCIDdate As String '本班的結訓日期 '本班的開訓日期 (最後使用經費日期)
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
        'str += ",FORMAT((pp.DefGovCost/cc.TNum), '#######0.##') ClsDefGovCost" & vbCrLf
        str &= " ,CASE WHEN cc.TNum <> 0 THEN FORMAT((pp.DefGovCost/cc.TNum), '#######0.##') ELSE '0' END ClsDefGovCost" & vbCrLf
        'str += ",FORMAT(TRUNC(pp.TotalCost/cc.TNum), '#######0.##') as ClsTtlCost" & vbCrLf
        str &= " ,CASE WHEN cc.TNum <> 0 THEN FORMAT(FLOOR(pp.TotalCost/cc.TNum), '#######0.##') ELSE '0' END ClsTtlCost" & vbCrLf
        str &= " FROM dbo.CLASS_CLASSINFO cc" & vbCrLf
        str &= " JOIN dbo.PLAN_PLANINFO pp ON pp.ComIDNO=cc.ComIDNO and pp.PlanID=cc.PlanID and pp.SeqNO=cc.SeqNO" & vbCrLf
        str &= " WHERE cc.OCID=@OCID1" & vbCrLf
        Dim sqldr As DataRow = DbAccess.GetOneRow(str, objconn, PMS1)
        If Not sqldr Is Nothing Then
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

        '可用補助額'產投 政府補助經費
        If ActSubsidyCost < TIMS.Get_3Y_SupplyMoney() Then
            LimitCost = TIMS.Get_3Y_SupplyMoney() - (CInt(ActSubsidyCost) + CInt(DefGovCost))  '可用政府補助經費(剩餘可用餘額)
            'If LimitCost < 0 Then LimitCost=0
            ViewState("LimitCost") = LimitCost.ToString
            'KeepSearch()
            If LimitCost >= ClsDefGovCost Then
                Check_GovCost = True
                'Check_GovCostFlag=True
                AlertType = 1
                TxtMessage = ""
                TxtMessage &= " 已報名申請補助費 " & CInt(SignSubsidyCost) + CInt(DefGovCost) & "元\n"
                TxtMessage &= " 已實際請領補助費 " & ActSubsidyCost & "元\n"
                'If Cst_MaxCanUseCost - ActSubsidyCost > 0 Then
                '    TxtMessage &= " (目前剩餘補助經費 " & Cst_MaxCanUseCost - ActSubsidyCost & "元)\n"
                'Else
                '    TxtMessage &= " (目前剩餘補助經費 0 元)\n"
                'End If
                'TxtMessage &= " 該班級補助申請款為" & ClsDefGovCost & "元\n"
                TxtMessage &= " 該班級補助申請款為" & ClsDefGovCost & "元，未超出補助額度\n"
                TxtMessage &= " 尚餘經費 " & LimitCost.ToString & " 元，可供申請\n"
                'JaveScriptAlertMsg="<script>alert('" & TxtMessage & "');</script>"
                'JaveScriptAlertMsg += "alert('填寫基本資料時，請務必填寫可以聯繫到您本人的資料，方便訓練單位與您聯繫!!');"
                JaveScriptAlertMsg = ""
                JaveScriptAlertMsg += "<script>"
                JaveScriptAlertMsg += "alert('" & TxtMessage & "');"
                JaveScriptAlertMsg += "location.href='" & Redirect2 & "';"
                JaveScriptAlertMsg += "</script>"
            Else
                Check_GovCost = False
                Check_GovCostFlag = False
                AlertType = 2
                TxtMessage = ""
                TxtMessage &= " 已報名申請補助費 " & CInt(SignSubsidyCost) + CInt(DefGovCost) & "元\r\n"
                TxtMessage &= " 已實際請領補助費 " & ActSubsidyCost & "元\r\n"
                'If Cst_MaxCanUseCost - ActSubsidyCost > 0 Then TxtMessage &= " (目前剩餘補助經費 " & Cst_MaxCanUseCost - ActSubsidyCost & "元)\r\n"
                TxtMessage &= " 尚餘經費 " & LimitCost.ToString & " 元，可供申請\r\n"
                TxtMessage &= " 該班級補助申請款為" & ClsDefGovCost & "元，將超出補助額度\r\n"
                TxtMessage &= " 可使用補助餘額為 " & LimitCost.ToString & " 元, 是否同意繼續報名\r\n"
                JaveScriptAlertMsg = VB_JaveConfirm(TxtMessage, "" & Redirect2 & "", "" & Redirect1 & "")
                'JaveScriptAlertMsg=TIMS.VB_JaveConfirm(TxtMessage, "" & Cst_Online_3 & "", "" & Cst_index_0 & "")
                'JaveScriptAlertMsg=VB_JaveConfirm(TxtMessage, "Online_3.aspx?LimitCost=" & ViewState("LimitCost"), "index.aspx")
            End If
        Else
            LimitCost = 0
            Check_GovCost = False
            Check_GovCostFlag = False
            AlertType = 3
            TxtMessage = ""
            TxtMessage &= " 已報名申請補助費 " & CInt(SignSubsidyCost) + CInt(DefGovCost) & "元\r\n"
            TxtMessage &= " 已實際請領補助費 " & ActSubsidyCost & "元\r\n"
            'If Cst_MaxCanUseCost - ActSubsidyCost > 0 Then TxtMessage &= " (目前剩餘補助經費 " & Cst_MaxCanUseCost - ActSubsidyCost & "元)\r\n"
            TxtMessage &= " 尚餘經費 " & LimitCost.ToString & " 元 \r\n"
            TxtMessage &= " 該班級補助申請款為" & ClsDefGovCost & "元，將超出補助額度\r\n"
            TxtMessage &= " " & TIMS.Get_3YSupplyMsg(Me) & "補助額度已滿, 是否選擇自費：\r\n"
            Const Cst_Msg1 As String = "請逕洽訓練單位報名即可"
            Const Cst_Msg2 As String = "謝謝您對97年度產業人才投資方案／提升在職勞工學習方案的支持"
            JaveScriptAlertMsg = VB_JaveConfirm2(TxtMessage, Cst_Msg1, Cst_Msg2, "" & Redirect1 & "")
            'JaveScriptAlertMsg=TIMS.VB_JaveConfirm2(TxtMessage, Cst_Msg1, Cst_Msg2, "" & Cst_index_0 & "")
            'JaveScriptAlertMsg="<script>confirm('" & TxtMessage & "');</script>"
        End If
        Return Check_GovCostFlag
    End Function

    '回答是 (連線1) 或否 (連線2) 
    Function VB_JaveConfirm(ByVal TxtMessage As String, ByVal YesUrl As String, ByVal NoUrl As String) As String
        Dim JaveScriptAlertMsg As String = ""
        JaveScriptAlertMsg &= " <script language=""javascript"">" & vbCrLf
        JaveScriptAlertMsg &= "  //<!--" & vbCrLf
        JaveScriptAlertMsg &= "  /*@cc_on   @*/" & vbCrLf
        JaveScriptAlertMsg &= "  /*@if (@_win32 && @_jscript_version>=5)" & vbCrLf
        JaveScriptAlertMsg &= "  function window.confirm(str)" & vbCrLf
        JaveScriptAlertMsg &= "  {" & vbCrLf
        JaveScriptAlertMsg &= "     str=str.replace(/\'/g, ""' & chr(39) & '"").replace(/\r\n/g, ""' & VBCrLf & '"");" & vbCrLf
        JaveScriptAlertMsg &= "  	execScript(""n=msgbox('"" + str + ""',257,'提示訊息')"",""vbscript"");" & vbCrLf
        JaveScriptAlertMsg &= "  	return (n==1);" & vbCrLf
        JaveScriptAlertMsg &= "  }" & vbCrLf
        JaveScriptAlertMsg &= "  @end @*/" & vbCrLf
        JaveScriptAlertMsg &= "  //debugger;" & vbCrLf
        JaveScriptAlertMsg &= "  //alert(confirm('1.\'第一行\';\r\n2.第二行;\r\n'));" & vbCrLf
        JaveScriptAlertMsg &= "  // -->" & vbCrLf
        JaveScriptAlertMsg &= " " & vbCrLf
        JaveScriptAlertMsg &= " if (confirm('" & TxtMessage & "')==true) {" & vbCrLf
        JaveScriptAlertMsg &= "     location.href='" & YesUrl & "';"
        JaveScriptAlertMsg &= " }" & vbCrLf
        JaveScriptAlertMsg &= " else {" & vbCrLf
        JaveScriptAlertMsg &= "     location.href='" & NoUrl & "';"
        JaveScriptAlertMsg &= " }" & vbCrLf
        JaveScriptAlertMsg &= " </script>" & vbCrLf
        Return JaveScriptAlertMsg
    End Function

    '回答是或否時都可 連線
    Function VB_JaveConfirm2(ByVal TxtMessage As String, ByVal YesMsg As String, ByVal NoMsg As String, Optional ByVal Url As String = "") As String
        Dim JaveScriptAlertMsg As String
        JaveScriptAlertMsg = ""
        JaveScriptAlertMsg &= " <script language=""javascript"">" & vbCrLf
        JaveScriptAlertMsg &= "  //<!--" & vbCrLf
        JaveScriptAlertMsg &= "  /*@cc_on   @*/" & vbCrLf
        JaveScriptAlertMsg &= "  /*@if   (@_win32 && @_jscript_version>=5)" & vbCrLf
        JaveScriptAlertMsg &= "  function window.confirm(str)" & vbCrLf
        JaveScriptAlertMsg &= "  {" & vbCrLf
        JaveScriptAlertMsg &= "    str=str.replace(/\'/g, ""' & chr(39) & '"").replace(/\r\n/g, ""' & VBCrLf & '"");" & vbCrLf
        JaveScriptAlertMsg &= "    execScript(""n=msgbox('"" + str + ""', 257, '提示訊息')"",""vbscript"");" & vbCrLf
        JaveScriptAlertMsg &= "    return(n==1);" & vbCrLf
        JaveScriptAlertMsg &= "  }" & vbCrLf
        JaveScriptAlertMsg &= "  @end   @*/" & vbCrLf
        JaveScriptAlertMsg &= "  //debugger;" & vbCrLf
        JaveScriptAlertMsg &= "  //alert(confirm('1.\'第一行\';\r\n2.第二行;\r\n'));" & vbCrLf
        JaveScriptAlertMsg &= "  //   -->" & vbCrLf
        JaveScriptAlertMsg &= " " & vbCrLf
        JaveScriptAlertMsg &= " if (confirm('" & TxtMessage & "')==true) {" & vbCrLf
        JaveScriptAlertMsg &= "     alert('" & YesMsg & "');" & vbCrLf
        JaveScriptAlertMsg &= " }" & vbCrLf
        JaveScriptAlertMsg &= " else {" & vbCrLf
        JaveScriptAlertMsg &= "     alert('" & NoMsg & "');" & vbCrLf
        JaveScriptAlertMsg &= " }" & vbCrLf
        If Url <> "" Then JaveScriptAlertMsg &= " location.href='" & Url & "';"
        JaveScriptAlertMsg &= " </script>" & vbCrLf
        Return JaveScriptAlertMsg
    End Function

    '黑名單判斷訊息顯示(學員處分)
    Function CheckData4(ByVal MyPage As Page, ByVal sIDNO As String, ByRef sErrmsg As String) As Boolean
        'START 黑名單判斷 
        Dim rst As Boolean = True '沒問題(沒有黑名單資訊)
        sErrmsg = ""
        '取得任一筆黑名單資訊

        Dim stdBLACK2TPLANID As String = ""
        Dim iStdBlackType As Integer = TIMS.Chk_StdBlackType(MyPage, objconn, stdBLACK2TPLANID)
        'Dim sqlWSB As String=TIMS.Get_StdBlackWSB(Me, iStdBlackType, stdBLACK2TPLANID, 1)
        Dim sTPlanID As String = sm.UserInfo.TPlanID '判斷計畫。

        Dim sql As String = ""
        sql &= " SELECT sb.IDNO,sb.SBComment,dbo.fn_CDate(sb.SBSdate) SBSdate" & vbCrLf
        sql &= " ,dbo.fn_CDate(DATEADD(month, 12*sb.SBYears, sb.SBSdate)) SBEdate" & vbCrLf
        sql &= " FROM dbo.STUD_BLACKLIST sb" & vbCrLf
        sql &= " WHERE sb.Avail='Y'" & vbCrLf
        sql &= " AND (GETDATE() >= sb.SBSdate AND GETDATE() <= DATEADD(month, sb.SBYears*12, sb.SBSdate) )" & vbCrLf
        sql &= " AND sb.IDNO=@IDNO" & vbCrLf
        Select Case iStdBlackType
            Case 0 '0:回傳(不啟用)
                sql &= " AND 1<>1" & vbCrLf  'Return rstOrgBlackList '回傳(不啟用)
            Case 1 '1：各計畫自行限制處分紀錄
                sql &= " AND sb.TPlanID='" & sTPlanID & "'" & vbCrLf '有效的(含不在限期內的)
            Case 2 '2：跨計畫合併限制處分紀錄，因跨計畫合併限制處分紀錄可能會有不同組合，需要另外一個欄位紀錄組合喔
                If stdBLACK2TPLANID <> "" Then
                    sql &= " AND sb.TPlanID IN (" & stdBLACK2TPLANID & ")" & vbCrLf '有效的(含不在限期內的)
                Else
                    sql &= " AND sb.TPlanID='" & sTPlanID & "'" & vbCrLf '有效的(含不在限期內的)
                End If
            Case 3 '3：所有計畫合併限制處分紀錄
            Case 4 '4:回傳(無處分限制)
                sql &= " AND 1<>1" & vbCrLf  'Return rstOrgBlackList '回傳(無處分限制)
        End Select
        sql &= " ORDER BY sb.SBSdate" & vbCrLf
        Dim sCmd As New SqlCommand(sql, objconn)

        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("IDNO", SqlDbType.VarChar).Value = sIDNO
            dt.Load(.ExecuteReader())
        End With

        'dt=DbAccess.GetDataTable(sqlstr)
        If TIMS.dtHaveDATA(dt) Then
            rst = False
            sErrmsg = $"查台端於{dt.Rows(0)("SBSdate")}因{dt.Rows(0)("SBComment")}"

            '字串長度切換-start
            Dim num As Int16 = Len(sErrmsg) \ 30
            Dim last As Int16 = Len(sErrmsg) Mod 30
            Dim j As Integer = 1
            Dim ErrMsg_TMP As String = ""
            For i As Integer = 1 To num + 1
                If i = num + 1 Then
                    ErrMsg_TMP = ErrMsg_TMP + Mid(sErrmsg, j, last)
                Else
                    ErrMsg_TMP = ErrMsg_TMP + Mid(sErrmsg, j, 30) & vbCrLf
                    j = j + 30
                End If
            Next
            sErrmsg = $"{ErrMsg_TMP}{vbCrLf}至{dt.Rows(0)("SBEdate")}前將不得再申領本署補助 。"
            '字串長度切換-end'Common.AddClientScript(Page, "alert('" & ErrMsg_TMP & "');")'Exit Function
        End If
        'END 黑名單判斷 
        Return rst
    End Function

    ''' <summary>
    ''' '不可報名或可報名訊息顯示 0:不可報名 1:可報名
    ''' </summary>
    ''' <param name="AlertType"></param>
    ''' <returns></returns>
    Function CheckData3(ByRef AlertType As Integer) As String
        AlertType = 0 ' 0:不可報名 1:可報名
        Dim rst As String = ""

        '限定開班。'GetOCIDDate2854
        If OCIDValue1.Value = "" Then
            ViewState("STDate") = ""
            rst &= "請選擇正確的課程名稱代號!!" & vbCrLf
            Return rst
        End If
        '檢查是否有此課程代碼
        Dim drCLASS As DataRow = TIMS.GetOCIDDate2854(OCIDValue1.Value, objconn)
        If drCLASS Is Nothing Then
            ViewState("STDate") = ""
            rst &= "請選擇正確的課程名稱代號!!" & vbCrLf
            Return rst
        End If

        '檢核 上架日期 (確認報名可否)
        Dim flag_Chk_OnShellDate As Boolean = If(TIMS.Cst_TPlanID28.IndexOf(Convert.ToString(drCLASS("TPlanID"))) > -1, True, False)
        If flag_Chk_OnShellDate Then
            '上架日期-ONSHELLDATE
            If Convert.ToString(drCLASS("ONSHELLDATE")) = "" Then
                rst &= " 此班級尚未開始報名!!!" & vbCrLf
                Return rst
            End If
            '上架日期-ONSHELLDATE
            Dim ChkTime3 As Long = 0
            ChkTime3 = DateDiff(DateInterval.Minute, CDate(drCLASS("today")), CDate(drCLASS("ONSHELLDATE"))) '未到結束報名時間大於0
            If ChkTime3 > 0 Then
                rst &= " 此班級尚未開始報名!!!" & vbCrLf
                Return rst
            End If
        End If

        ViewState("SEnterDate") = TIMS.CssFormatDate(drCLASS("SEnterDate"))
        ViewState("FEnterDate") = TIMS.CssFormatDate(drCLASS("FEnterDate"))
        If DateDiff(DateInterval.Minute, CDate(drCLASS("today")), CDate(drCLASS("SEnterDate"))) > 0 Then
            'Label99.Text="尚未開始招生"
            rst += ViewState("SEnterDate") & " 此班級尚未開始報名!!!" & vbCrLf
        ElseIf DateDiff(DateInterval.Minute, CDate(drCLASS("FEnterDate")), CDate(drCLASS("today"))) > 0 Then
            'Label99.Text="已招生結束"
            rst += ViewState("FEnterDate") & " 此班報名時間已過!!!" & vbCrLf
        ElseIf Convert.ToString(drCLASS("IsClosed")) = "Y" Then
            'Label99.Text="此班已結訓"
            rst += ViewState("FEnterDate") & " 此班已結訓!!!" & vbCrLf
        ElseIf Convert.ToString(drCLASS("IsClosed")) = "N" _
                AndAlso DateDiff(DateInterval.Minute, CDate(drCLASS("today")), CDate(drCLASS("FEnterDate"))) > 0 Then
            'AndAlso EnterCount >= CInt(drClass("TNum")) Then
            'Label99.Text="報名額滿 (接受以備取身分報名)"
            AlertType = 1 ' 0:不可報名 1:可報名
            Dim iEnterCount As Integer = TIMS.Get_EnterCount(OCIDValue1.Value, objconn) '取得報名人數
            If iEnterCount >= drCLASS("TNum") Then '如果報名人數大於班級的核定訓練人數
                rst &= " 本班級已報名額滿，現在報名將列為備取!!!" & vbCrLf
            End If
        End If

        ''為配合2016年度課程公告作業，擬將2016年上半年度核可並轉班完成之課程統一於2016年1月23日0:01起方才能於產投報名網站上查詢到公告課程。
        'If ChkTime1 >= 0 AndAlso ChkTime2 >= 0 Then
        '    Dim bln_noEnter As Boolean=False
        '    If DateDiff(DateInterval.Second, aNow, CDate(TIMS.cst_SEnterDate2016_28)) >= 0 Then
        '        bln_noEnter=True
        '    End If
        '    If Convert.ToString(dr("Years"))="2016" AndAlso bln_noEnter Then
        '        ChkTime1=-1
        '        ChkTime2=-1
        '        ViewState("SEnterDate")=TIMS.CssFormatDate(CDate(TIMS.cst_SEnterDate2016_28))
        '    End If
        'End If

        ViewState("STDate") = TIMS.Cdate3(drCLASS("STDate")) '班級確認塞值

        Return rst
    End Function

    '重複報名訊息顯示 依身分證號判斷
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

#Region "(No Use)"

    ''判斷身分證號與生日組合 是否異常於 原系統資訊
    'Function checkData3() As String
    '    Dim rst As String
    '    Const cst_errmsg1 As String="判斷身分證號與生日組合 與系統不吻合!! 若確定該報名資料正確，請先連絡系統管理者，修改系統資訊!!"
    '    Dim objstr As String=""
    '    Dim dt1 As DataTable
    '    Dim dt2 As DataTable
    '    Dim dt3 As DataTable

    '    objstr="select IDNO,convert(varchar,BirthDay,111) BirthDay from Stud_EnterTemp   where IDNO='" & IDNO.Text & "'" & vbCrLf
    '    dt1=DbAccess.GetDataTable(objstr, objconn)
    '    If rst="" AndAlso dt1.Rows.Count > 0 Then
    '        For i As Integer=0 To dt1.Rows.Count - 1
    '            Dim dr As DataRow=dt1.Rows(i)
    '            If DateDiff(DateInterval.Day, CDate(dr("BirthDay")), CDate(birthDay.Text)) <> 0 Then
    '                rst += cst_errmsg1 & vbCrLf
    '                Exit For
    '            End If
    '        Next
    '    End If

    '    If rst="" Then
    '        objstr="select IDNO,convert(varchar,BirthDay,111) BirthDay from Stud_EnterTemp2   where IDNO='" & IDNO.Text & "'" & vbCrLf
    '        dt2=DbAccess.GetDataTable(objstr, objconn)
    '    End If
    '    If rst="" AndAlso dt2.Rows.Count > 0 Then
    '        For i As Integer=0 To dt2.Rows.Count - 1
    '            Dim dr As DataRow=dt2.Rows(i)
    '            If DateDiff(DateInterval.Day, CDate(dr("BirthDay")), CDate(birthDay.Text)) <> 0 Then
    '                rst += cst_errmsg1 & vbCrLf
    '                Exit For
    '            End If
    '        Next
    '    End If

    '    If rst="" Then
    '        objstr="select IDNO,convert(varchar,BirthDay,111) BirthDay from Stud_studentinfo   where IDNO='" & IDNO.Text & "'" & vbCrLf
    '        dt3=DbAccess.GetDataTable(objstr, objconn)
    '    End If
    '    If rst="" AndAlso dt3.Rows.Count > 0 Then
    '        For i As Integer=0 To dt3.Rows.Count - 1
    '            Dim dr As DataRow=dt3.Rows(i)
    '            If DateDiff(DateInterval.Day, CDate(dr("BirthDay")), CDate(birthDay.Text)) <> 0 Then
    '                rst += cst_errmsg1 & vbCrLf
    '                Exit For
    '            End If
    '        Next
    '    End If

    '    Return rst
    'End Function

#End Region

    '基本資料驗證訊息顯示
    Function CheckData1(Optional ByVal sType As Integer = 0) As String
        'sType 0:全部判斷 1:排除生日判斷
        Dim rst As String = ""

        If OCID1.Text = "" OrElse OCIDValue1.Value = "" Then rst += "請選擇報名班級!!" & vbCrLf

        IDNO.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(IDNO.Text))
        If IDNO.Text = "" Then
            rst += "請輸入身分證號碼!!" & vbCrLf
        Else
            'Dim strIDNO As String=""
            Dim strIDNO As String = IDNO.Text
            If strIDNO.Length >= 10 Then
                Select Case Convert.ToString(strIDNO).Substring(1, 1)
                    Case "1", "2"
                        If Not TIMS.CheckIDNO(strIDNO) Then
                            rst += "身分證號碼有誤(第2碼為 1、2 依國民身分證規則判斷)!!" & vbCrLf
                        End If
                    Case "A", "C", "B", "D"
                        If Not TIMS.CheckIDNO2(strIDNO, 2) Then
                            rst += "居留證號碼有誤(第2碼為 A、C、B、D 依國民居留證號碼規則判斷)!!" & vbCrLf
                        End If
                    Case "8", "9"
                        If Not TIMS.CheckIDNO2(strIDNO, 4) Then
                            rst += "居留證號碼有誤(第2碼為 8、9 依國民居留證號碼規則判斷)!!" & vbCrLf
                        End If
                    Case Else
                        rst += "身分證號碼或居留證號碼有誤，第2碼，應為(1.2.或A.C.B.D.或8.9.)" & vbCrLf
                End Select
            Else
                rst += "身分證號碼或居留證號碼有誤!!" & vbCrLf
            End If
            'If Not TIMS.CheckIDNO(IDNO.Text) Then
            '    rst += "身分證號碼有誤!!" & vbCrLf
            'End If
        End If

        birthDay.Text = TIMS.ClearSQM(birthDay.Text)
        If sType = 0 Then
            If Trim(birthDay.Text) <> "" Then
                birthDay.Text = Trim(birthDay.Text)
                If TIMS.IsDate1(birthDay.Text) Then
                    birthDay.Text = CDate(birthDay.Text).ToString("yyyy/MM/dd")
                    'Common.FormatDate(birthDay.Text)
                Else
                    rst += "出生日期格式有誤!!" & vbCrLf
                End If
            Else
                rst += "請輸入出生日期!!" & vbCrLf
            End If
        End If

#Region "(No Use)"

        'If Me.EnterDate.Text="" Then
        '    rst += "請輸入報名日期!!" & vbCrLf
        'Else
        '    If IsDate(Me.EnterDate.Text) Then
        '        Me.EnterDate.Text=Common.FormatDate(Me.EnterDate.Text)
        '    Else
        '        rst += "報名日期格式有誤!!" & vbCrLf
        '    End If
        'End If

#End Region

        Return rst
    End Function

    Sub GetSearchStr()
        Dim s_SearchStr As String = ""
        TIMS.SetMyValue(s_SearchStr, "center", center.Text)
        TIMS.SetMyValue(s_SearchStr, "RIDValue", RIDValue.Value)
        TIMS.SetMyValue(s_SearchStr, "OCID1", OCID1.Text)
        TIMS.SetMyValue(s_SearchStr, "TMID1", TMID1.Text)
        TIMS.SetMyValue(s_SearchStr, "OCIDValue1", OCIDValue1.Value)
        TIMS.SetMyValue(s_SearchStr, "TMIDValue1", TMIDValue1.Value)
        TIMS.SetMyValue(s_SearchStr, "IDNO", TIMS.ChangeIDNO(IDNO.Text))
        TIMS.SetMyValue(s_SearchStr, "birthDay", TIMS.Cdate3(birthDay.Text))
        Session("_SearchStr") = s_SearchStr
        'Session("_SearchStr") += "&EnterDate=" & EnterDate.Text
    End Sub

    ''' <summary> 報名 </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnAdd1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd1.Click
        Dim aNow As Date = TIMS.GetSysDateNow(objconn)
        If TIMS.StopEnterTempMsg1(Me, objconn, True) Then Exit Sub
        Dim sAltMsg As String = "" '訊息
        Dim flag_stopEnter2 As Boolean = TIMS.StopEnterTempMsg2(objconn, sAltMsg)
        If flag_stopEnter2 Then
            Common.MessageBox(Me, sAltMsg)
            Exit Sub
        End If

        '報名資料再確認
        Dim xErrmsg As String = ""
        IDNO.Text = TIMS.ClearSQM(IDNO.Text)
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        birthDay.Text = TIMS.ClearSQM(birthDay.Text)

        If IDNO.Text = "" Then xErrmsg &= " 身分證號碼-資料有誤，請重新輸入！" & vbCrLf
        If birthDay.Text = "" OrElse Not TIMS.IsDate1(birthDay.Text) Then xErrmsg &= " 出生日期-資料有誤，請重新輸入！" & vbCrLf
        If xErrmsg <> "" Then
            Common.MessageBox(Me, xErrmsg)
            Exit Sub
        End If

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

        'If TestStr="AmuTest" Then '測試
        '    ChkTime1=1
        '    ChkTime2=1
        'End If '測試

        '為配合2016年度課程公告作業，擬將2016年上半年度核可並轉班完成之課程統一於2016年1月23日0:01起方才能於產投報名網站上查詢到公告課程。
        If ChkTime1 >= 0 AndAlso ChkTime2 >= 0 Then
            Dim bln_noEnter As Boolean = False
            If DateDiff(DateInterval.Second, aNow, CDate(TIMS.cst_SEnterDate2016_28)) >= 0 Then
                bln_noEnter = True
            End If
            'If Convert.ToString(drCC("Years"))="2016" AndAlso bln_noEnter Then
            '    ChkTime1=-1
            '    ChkTime2=-1
            '    ViewState("SEnterDate")=TIMS.CssFormatDate(CDate(TIMS.cst_SEnterDate2016_28))
            'End If
        End If

        '在報名時間內
        If ChkTime1 > 0 AndAlso ChkTime2 >= 0 Then
            'vsOCIDvalue
            'Session(tims.cst_OCID)=Trim(Li_Class.Text)  '記錄報名課程資料
            '在報名時間內
        ElseIf ChkTime1 < 0 Then
            '此班級將於(該班可報名時間)開始報名!!
            'Common.MessageBox(Me, ViewState("SEnterDate") & " 此班級尚未開始報名!!!")
            xErrmsg &= " 此班級將於(" & ViewState("SEnterDate") & ")開始報名!!!" & vbCrLf
        Else
            '報名時間已過
            xErrmsg += ViewState("FEnterDate") & " 此班報名時間已過!!!" & vbCrLf
        End If

        If xErrmsg <> "" Then
            'ViewState("aNow")=TIMS.CssFormatDate(aNow)
            'aNow=TIMS.GetSysDateNow(objconn)
            'aToday=Common.FormatDate(aNow) 'Today()
            xErrmsg += "" & vbCrLf
            'xErrmsg &= " 目前系統時間為(" & ViewState("aNow") & ")!!" & vbCrLf
            xErrmsg &= " 目前系統時間為(" & TIMS.CssFormatDate(aNow) & ")!!" & vbCrLf
            Dim flagS1 As Boolean = TIMS.IsSuperUser(Me, 1) '是否為(後台)系統管理者 
            Dim flagS2 As Boolean = TIMS.sUtl_ChkTest() '測試環境
            If flagS1 OrElse flagS2 Then
                xErrmsg += cst_msgSuper1 & vbCrLf
                Common.MessageBox(Me, xErrmsg)
            Else
                Common.MessageBox(Me, xErrmsg)
                Exit Sub
            End If
        End If

        IDNO.Text = TIMS.ChangeIDNO(IDNO.Text)
        'ViewState("IDNO")=Convert.ToString(IDNO.Text)

        Dim Redirect1 As String = String.Concat("SD_01_010.aspx?ID=", Request("ID")) '課程代號輸入頁面
        Dim Redirect2 As String = String.Concat("SD_01_010_add.aspx?ID=", Request("ID"), "&proecess=add&IDNO=", TIMS.ChangeIDNO(IDNO.Text)) '輸入個人資料報名頁面

        'Dim ssScript As String
        Dim Errmsg As String = ""
        '基本資料驗證
        Errmsg = CheckData1()
        '無錯誤，判斷 是否重複報名 '重複報名訊息顯示 依身分證號判斷
        If Errmsg = "" Then Errmsg = CheckData2()
        '判斷身分證號與生日組合 是否異常於 原系統資訊
        If Errmsg = "" Then Errmsg = TIMS.ChkDataIdnoBirth(objconn, IDNO.Text, birthDay.Text)
        If hidCheck1.Value = "" Then
            Common.MessageBox(Me, "請先按台端資料確認!!")
            Exit Sub
        End If
        If Convert.ToString(ViewState("STDate")) = "" Then
            Common.MessageBox(Me, "請按 班級確認 進入報名!!")
            Exit Sub
        End If
        '檢測此學員是否 可參訓 產業人才投資方案 (大於15歲者)
        If Errmsg = "" Then
            If Not TIMS.Check_YearsOld15(birthDay.Text, CDate(ViewState("STDate"))) Then Errmsg += "學員資格 年齡不滿15歲 不符合可參訓條件！" & vbCrLf
        End If

        '有基本錯誤
        If Errmsg <> "" Then
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If

        'WebRequest物件如何忽略憑證問題
        System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf TIMS.ValidateServerCertificate)
        'TLS 1.2-基礎連接已關閉: 傳送時發生未預期的錯誤 
        System.Net.ServicePointManager.SecurityProtocol = Net.SecurityProtocolType.Tls12
        '檢核學員重複參訓。
        'http://163.29.199.211/TIMSWS/timsService1.asmx
        'https://wltims.wda.gov.tw/TIMSWS/timsService1.asmx
        Dim timsSer1 As New timsService1.timsService1
        Dim sIDNO As String = TIMS.ClearSQM(IDNO.Text)
        Dim aOCID1 As String = TIMS.ClearSQM(drCC("OCID"))
        '檢核學員重複參訓。
        Dim xStudInfo As String = ""
        TIMS.SetMyValue(xStudInfo, "IDNO", sIDNO)
        TIMS.SetMyValue(xStudInfo, "OCID1", aOCID1)
        '檢核學員OOO有參訓時段重疊報名情形，無法儲存，請再確認！
        Call TIMS.ChkStudDouble(timsSer1, Errmsg, "", xStudInfo)
        If Errmsg <> "" Then
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If

        '基本沒有錯誤
        Call GetSearchStr()

        Try
            Dim AlertMsg As String = ""
            '無錯誤，判斷 補助 額是否不足
            If Check_GovCost(IDNO.Text, OCIDValue1.Value, Redirect1, Redirect2, AlertMsg, "1") Then
                '有需要顯示狀況(可以報名)
                Common.RespWrite(Me, AlertMsg)
            Else
                '有需要顯示的狀況(可能還是可以報名)
                Common.RespWrite(Me, AlertMsg)
            End If
        Catch ex As Exception
            TIMS.WriteTraceLog(Nothing, ex, ex.ToString)
            Common.MessageBox(Me, "!!儲存失敗!!")
            Common.MessageBox(Me, ex.ToString)
        End Try
    End Sub

    ''' <summary>
    ''' 班級確認鈕
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnCheckClass_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCheckClass.Click
        Dim xErrmsg As String = ""
        'SE:'產投停止報名
        Const cst_StopFlag As String = "SE"
        Dim aNow As Date = TIMS.GetSysDateNow(objconn)
        Dim sAltMsg As String = "" '訊息
        Dim AltMsgSDate As String = "" '訊息公佈日
        Dim AltMsgEDate As String = "" '訊息結束日
        sAltMsg = TIMS.Get_System_Msg("AltMsg", cst_StopFlag, objconn)
        AltMsgSDate = TIMS.Get_System_Msg("AltMsgSDate", cst_StopFlag, objconn)
        AltMsgEDate = TIMS.Get_System_Msg("AltMsgEDate", cst_StopFlag, objconn)
        xErrmsg = TIMS.Get_AltMsg_System_Msg(sAltMsg, AltMsgSDate, AltMsgEDate, aNow)
        If xErrmsg <> "" Then
            '因網路系統維護，將於2012年8月14日12:10至13:10中斷服務1小時，造成不便，敬請見諒！
            Common.AddClientScript(Page, "alert('" & xErrmsg & "');location.href='../../main2.aspx';")
            Exit Sub
        End If

        Dim AlertType As Integer = 0 '0:不可報名/1:可報名
        Dim Errmsg As String = ""

        If OCIDValue1.Value <> "" Then
            If Errmsg = "" Then Errmsg = CheckData3(AlertType) '報名時間判斷式
            btnAdd1.Enabled = False
            'If btnAdd1.Enabled Then
            '    If FunDr("Adds")="1" Then btnAdd1.Enabled=True Else btnAdd1.Enabled=False
            'End If
            If AlertType = 1 Then
                '===== 可開放報名(AlertType=1) =====
                btnAdd1.Enabled = True
                'If btnAdd1.Enabled Then
                '    If FunDr("Adds")="1" Then btnAdd1.Enabled=True Else btnAdd1.Enabled=False
                'End If
                '===== 再次檢驗權限是否可開放報名 =====
                'btnAdd1.Enabled=False
                'If blnCanAdds Then btnAdd1.Enabled=True
                'If Not btnAdd1.Enabled Then TIMS.Tooltip(btnAdd1, "權限無法使用 報名功能", True)
            End If

            If Not btnAdd1.Enabled Then
                Dim flagS1 As Boolean = TIMS.IsSuperUser(Me, 1) '是否為(後台)系統管理者 
                Dim flagS2 As Boolean = TIMS.sUtl_ChkTest() '測試環境
                If flagS1 Then
                    btnAdd1.Visible = True
                    TIMS.Tooltip(btnAdd1, "(已過報名時間)-" & cst_msgSuper1)
                End If
                If flagS2 Then '測試環境
                    btnAdd1.Enabled = True
                    TIMS.Tooltip(btnAdd1, "(已過報名時間)-測試環境，測試用!!!")
                End If
            End If
            If Errmsg <> "" Then
                '有需告知訊息的顯示
                Common.MessageBox(Me, Errmsg)
                Exit Sub
            End If
        End If
    End Sub

    Private Sub BtnCheck1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCheck1.Click
        'Stud_Blacklist
        hidCheck1.Value = "1"
        IDNO.Text = TIMS.ChangeIDNO(IDNO.Text)
        'ViewState("IDNO")=Convert.ToString(IDNO.Text)

        Dim Errmsg As String = ""
        Errmsg = CheckData1(1)
        '無錯誤，判斷 是否重複報名'重複報名訊息顯示 依身分證號判斷
        If Errmsg = "" Then Errmsg = CheckData2()
        '判斷身分證號與生日組合 是否異常於 原系統資訊
        If Errmsg = "" Then Errmsg = TIMS.ChkDataIdnoBirth(objconn, IDNO.Text, birthDay.Text)
        If Errmsg <> "" Then
            '有基本錯誤
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If

        If IDNO.Text = "" Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        Dim AlertMsg As String = ""
        '黑名單測試
        Dim oflagblack As Boolean = CheckData4(Me, IDNO.Text, AlertMsg)
        'If Not checkData4(Me, ViewState("IDNO"), AlertMsg) Then
        '    '有黑名單資料
        'End If
        If AlertMsg <> "" Then
            '有黑名單資料
            'Common.AddClientScript(Page, "alert('" & AlertMsg & "');")
            Common.MessageBox(Me, AlertMsg)
        Else
            Common.MessageBox(Me, "台端資料確認!!")
        End If
    End Sub
End Class