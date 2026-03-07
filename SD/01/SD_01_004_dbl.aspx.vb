Public Class SD_01_004_dbl
    Inherits System.Web.UI.Page

    'STUD_ENTERDOUBLE
    'http://163.29.199.211/TIMSWS/timsService1.asmx (TimsWebserviceo)
    'Dim timsSer1 As New timsService1.timsService1

    '隢?敹?甈∠Ⅱ隤?衣隞乩??摩
    '?單?瘥?嚗蝬脣?祟????撖拇??????SD_01_004)
    '??雿平(??甇????鞊???SD_03_001)
    '摮詨??(??閰脫活??詨?啗? (SD_02_003) '?單??脰???蝝???畾菟???撠?銝西???
    '1.??芷?銵?瘨??剔??
    '2.?e蝬脣?????蝝?
    '3.?撌脤?閮??剔???
    '4.??芷????剔???
    '5.??剔???敺洵15?亦??剔??蝬脣祟?豢撖押?
    '6.??剔???敺洵15?亦??剔????隢??????
    '7.??剔???敺洵15?亦??剔???勗(甇????鞊???
    '8.?撌脩?閮??剔???

    Dim ff As String = "" '?蕪摮?
    Dim ss As String = "" '?蕪摮?
    Dim gssOCID As String = ""
    Dim gsPTDID As String = "" '???銴?隤脩?
    Dim rqOCID1 As String = ""
    Dim rqIDNO As String = ""
    Dim rqBirth As String = ""
    Dim dtStdCost As DataTable
    Dim dtStud As DataTable
    Dim dtTrain As DataTable
    'Dim rOver6w As Boolean = False
    Dim iOver6w As Double = 0
    'Const cst_msg1 As String = "?箏????刻?蝺渲?皞?隢???挾????????閮?畾菟???隤脩?嚗?????嚗蒂??擗玨蝔???? & vbCrLf & "憒歇?⊥??芾???隢蜓?晾閰Ｚ?蝺游雿??抬?雓???
    Const cst_msg6 As String = "???剁??摯?函???抵祥雿輻撌脤?6?砍?嚗??怠歇?豢??閮葉?歇?勗??玨蝔?嚗??函???"
    Const cst_msgover6w As String = "餈?撟游鋆鞎颱蝙???恍?隡?嚗?
    Const cst_msgTxt1 As String = "?亦?Ｘ??寞?銝玨?挾??蝝??!"
    Dim sDate As String = String.Empty
    Dim eDate As String = String.Empty
    Dim oSTDate As String = ""

    Dim objconn As OracleConnection

    Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
        ' (?湔??AuthBasePage ??, 銝?瑼Ｘ Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)

        PageControler1.PageDataGrid = DataGrid2bb

        If Convert.ToString(Request("A")) = "Y" Then
            If initObj28() Then Call Search1() '?寞?皜祈岫
        Else
            '銝?砍?雿?
            If Not IsPostBack Then
                If initObj28() Then Call Search1()
            End If
        End If

#Region "(No Use)"

        ''?寞?皜祈岫
        'If Convert.ToString(Request("A")) = "Y" Then
        '    If initObj() Then
        '        Call Search1()
        '    End If
        '    Call TIMS.Get_SubSidyCostDay(rqIDNO, oSTDate, sDate, eDate, objconn)
        '    Common.RespWrite(Me, TIMS.Get_DefGeoCost28s(rqIDNO, sDate, eDate, objconn))
        'End If

#End Region
    End Sub

    '?瑁?甇?Ⅱ?摭rue ?蚌alse
    Function initObj28() As Boolean
        Dim rst As Boolean = False '?蚌alse

        Dim ReqeSETID As String = Request("eSETID")
        Dim ReqeSerNum As String = Request("eSerNum")
        ReqeSETID = TIMS.ClearSQM(ReqeSETID)
        ReqeSerNum = TIMS.ClearSQM(ReqeSerNum)

        Dim dr As DataRow = Nothing
        If ReqeSETID <> "" Then dr = TIMS.Get_ENTERTEMP2(ReqeSETID, objconn)
        If dr Is Nothing Then
            Common.MessageBox(Me, "?亦??鞈?")
            Return rst 'Exit Function
        End If
        Dim dr1 As DataRow = Nothing
        If ReqeSerNum <> "" Then dr1 = TIMS.Get_ENTERTYPE2(ReqeSerNum, objconn)
        If dr1 Is Nothing Then
            Common.MessageBox(Me, "?亦??鞈?")
            Return rst 'Exit Function
        End If
        Hid_eSerNum.Value = ReqeSerNum
        rqOCID1 = Request("OCID1")
        rqIDNO = Request("IDNO")
        rqOCID1 = TIMS.ClearSQM(rqOCID1)
        rqIDNO = TIMS.ClearSQM(rqIDNO)
        If rqIDNO = "" OrElse rqIDNO <> Convert.ToString(dr("IDNO")) Then
            Common.MessageBox(Me, "?亦??鞈?")
            Return rst 'Exit Function
        End If
        If rqOCID1 = "" OrElse rqOCID1 <> Convert.ToString(dr1("OCID1")) Then
            Common.MessageBox(Me, "?亦??鞈?")
            Return rst 'Exit Function
        End If
        If ReqeSETID <> Convert.ToString(dr1("eSETID")) Then
            Common.MessageBox(Me, "?亦??鞈?")
            Return rst 'Exit Function
        End If
        rqBirth = TIMS.cdate3(dr("BIRTHDAY"))
        LabName.Text = Convert.ToString(dr("NAME"))
        'TPlanName.Text = TIMS.GetPlanName(sm.UserInfo.PlanID, objconn)

        Dim ERRMSG As String = ""
        Dim xStudInfo As String = ""
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '瑼Ｘ摮詨??????
            'http://163.29.199.211/TIMSWS/timsService1.asmx
            Dim timsSer1 As New timsService1.timsService1
            TIMS.SetMyValue(xStudInfo, "IDNO", rqIDNO)
            TIMS.SetMyValue(xStudInfo, "OCID1", rqOCID1)
            '瑼Ｘ摮詨OOO??閮?畾菟????敶ｇ??⊥??脣?嚗??Ⅱ隤?
            Call TIMS.ChkStudDouble(timsSer1, ERRMSG, "", xStudInfo)
        End If
        rst = True '甇?Ⅱ?摭rue
        Return rst 'Exit Function
    End Function

    '????
    Sub Search1()
        Dim aNow As Date
        aNow = TIMS.GetSysDateNow(objconn)
        Dim sMsg As String = ""
        Dim dr1 As DataRow = TIMS.GetOCIDDate(rqOCID1, objconn)
        oSTDate = Convert.ToString(TIMS.cdate3(dr1("STDate")))
        sDate = String.Empty
        eDate = String.Empty
        Call TIMS.Get_SubSidyCostDay(rqIDNO, oSTDate, sDate, eDate, objconn)
        sMsg &= "嚗qOCID1:" & rqOCID1
        sMsg &= "嚗STDate:" & oSTDate
        sMsg &= "嚗D:" & sDate
        sMsg &= "嚗D:" & eDate

        '? 鋆?隢閰?[SQL] (閰脣飛?∠?????拚?鞈?)
        dtStdCost = TIMS.GetHistoryDefStdCostDt2(rqIDNO, rqBirth, objconn)
        sMsg &= "嚗ost_CNT:" & dtStdCost.Rows.Count
        sMsg &= "(" & TIMS.Get_ssOCID(dtStdCost) & ")"

        '(餈?撟? ??閰脣飛?⊥??????Ｘ???)
        '1.sDate eDate ?箏?憛怠?2.???賣???閮????隞亥????3撟?
        Dim dt As DataTable
        dt = TIMS.GetRelEnterDateDtY3(rqIDNO, rqBirth, sDate, eDate, objconn)
        sMsg &= "嚗Y3_CNT:" & dt.Rows.Count
        sMsg &= "(" & TIMS.Get_ssOCID(dt) & ")"

        '??摮詨鞈?銵具?閰脣飛?∠??????
        dtStud = TIMS.Get_StudInfo(rqIDNO, objconn)
        sMsg &= "嚗tud_CNT:" & dtStud.Rows.Count
        sMsg &= "(" & TIMS.Get_ssOCID(dtStud) & ")"

        'HidDouble.Value = "" '???玨蝔???1.??銴?"".瘝?銴?
        'HidMoney6.Value = "" '雿輻撌脤?6?砍? 1.?" ".瘝?
        Me.labMoneyShow1.Text = ""
        'Me.labMoneyShow2.Text = ""
        'Me.labMoneyShow3.Text = ""

        Dim flagStd As Boolean = True '閰脣飛?∠銝玨鞈?嚗隢?鋆 ?true 瘝?@false
        If dtStdCost.Rows.Count = 0 Then
            '餈?撟湔??????
            gssOCID = TIMS.Get_ssOCID(dt) '???桀?table???cid隞仿???
            dtTrain = TIMS.Get_TRAINDESC(gssOCID, dt, objconn)   '??閬?撠?隤脩?鞈?
            '餈?撟湔??????靽???嚗?斤?????勗?憭望??)
            gssOCID = TIMS.Get_ssOCID(dtTrain) '???桀?table???cid隞仿???(?銝鈭??玨蝔?鞈?)
            flagStd = False
            '閰脣飛?∠銝玨鞈?
            If dt.Rows.Count > 0 Then
                Dim dr As DataRow
                '???鞈?蝑('6???抒??勗?鞈?嚗?閮?雿輻??瘥?雿輻??
                '3撟渡?蝭?鞈?
                ff = "XDAY2=1 AND XDAY3=1"
                ss = "STDate"
                If dt.Select(ff, ss).Length > 0 Then
                    '??隞嗥?1蝑???
                    dr = dt.Select(ff, ss)(0)
                    If Convert.ToString(Request("A")) = "Y" Then
                        oSTDate = Convert.ToString(dr("STDate"))
                        sDate = String.Empty
                        eDate = String.Empty
                        Call TIMS.Get_SubSidyCostDay(rqIDNO, dr("STDate"), sDate, eDate, objconn)
                        sMsg &= "嚗?.o:" & CStr(dr("OCID"))
                        sMsg &= "嚗TDate:" & CStr(dr("STDate"))
                        sMsg &= "嚗D:" & sDate
                        sMsg &= "嚗D:" & eDate
                    End If
                    Me.labMoneyShow1.Text = TIMS.Check_GovCost(rqIDNO, rqOCID1, iOver6w, objconn, "")
                Else
                    '?⊥?隞嗥??敺?蝑???
                    dr = dt.Rows(dt.Rows.Count - 1)
                    If Convert.ToString(Request("A")) = "Y" Then
                        oSTDate = Convert.ToString(dr("STDate"))
                        sDate = String.Empty
                        eDate = String.Empty
                        Call TIMS.Get_SubSidyCostDay(rqIDNO, dr("STDate"), sDate, eDate, objconn)
                        sMsg &= "嚗?.o:" & CStr(dr("OCID"))
                        sMsg &= "嚗TDate:" & CStr(dr("STDate"))
                        sMsg &= "嚗D:" & sDate
                        sMsg &= "嚗D:" & eDate
                    End If
                    Me.labMoneyShow1.Text = TIMS.Check_GovCost(rqIDNO, rqOCID1, iOver6w, objconn, "")
                End If
            End If
        End If
        sMsg &= "嚗lagStd:" & Convert.ToString(flagStd)

        If flagStd AndAlso dt.Rows.Count > 0 Then
            '餈?撟湔??????
            gssOCID = TIMS.Get_ssOCID(dt) '???桀?table???cid隞仿???
            dtTrain = TIMS.Get_TRAINDESC(gssOCID, dt, objconn)   '??閬?撠?隤脩?鞈?
            '餈?撟湔??????靽???嚗?斤?????勗?憭望??)
            gssOCID = TIMS.Get_ssOCID(dtTrain) '???桀?table???cid隞仿??? 
            'labMoneyShow1.Text = "" '撖阡??豢鋆鞎餌
            'labMoneyShow2.Text = "" '??銝剛玨蝔?隡啗??抵祥??
            'labMoneyShow3.Text = "" '?勗?銝剛玨蝔?隡啗??抵祥??
            Dim dr As DataRow
            '???鞈?蝑('6???抒??勗?鞈?嚗?閮?雿輻??瘥?雿輻??
            '3撟渡?蝭?鞈?
            ff = "XDAY2=1 AND XDAY3=1"
            ss = "STDate"
            If dt.Select(ff, ss).Length > 0 Then
                dr = dt.Select(ff, ss)(0) '蝚?蝑?
                If Convert.ToString(Request("A")) = "Y" Then
                    '皜祈岫?
                    oSTDate = Convert.ToString(dr("STDate"))
                    sDate = String.Empty
                    eDate = String.Empty
                    Call TIMS.Get_SubSidyCostDay(rqIDNO, dr("STDate"), sDate, eDate, objconn)
                    sMsg &= "嚗?.o:" & CStr(dr("OCID"))
                    sMsg &= "嚗TDate:" & CStr(dr("STDate"))
                    sMsg &= "嚗D:" & sDate
                    sMsg &= "嚗D:" & eDate
                End If
                '憿舐內?桀?雿輻????
                Me.labMoneyShow1.Text = TIMS.Check_GovCost(rqIDNO, rqOCID1, iOver6w, objconn, "")
            Else
                '6???抒鞈???敺?蝑???
                dr = dt.Rows(dt.Rows.Count - 1)
                If Convert.ToString(Request("A")) = "Y" Then
                    '皜祈岫?
                    oSTDate = Convert.ToString(dr("STDate"))
                    sDate = String.Empty
                    eDate = String.Empty
                    Call TIMS.Get_SubSidyCostDay(rqIDNO, dr("STDate"), sDate, eDate, objconn)
                    sMsg &= "嚗?.(m6N)o:" & CStr(dr("OCID"))
                    sMsg &= "嚗TDate:" & CStr(dr("STDate"))
                    sMsg &= "嚗D:" & sDate
                    sMsg &= "嚗D:" & eDate
                End If
                Me.labMoneyShow1.Text = TIMS.Check_GovCost(rqIDNO, rqOCID1, iOver6w, objconn, "")
            End If
        End If

        'If sYear2015Test = "Y" Then Me.labMoneyShow3.Text = "(皜祈岫?啣???皜祈岫)" & cst_titleS1 & cst_titleS2

        'RecordCount.Text = dt.Rows.Count
        'msg.Text = "?∪?閮???!"
        'PageControler1.Visible = False
        'Me.DataGrid2.Visible = False

        Dim chkDouble As Boolean = False '?斗???鞈???
        If dt.Rows.Count > 0 Then
            For Each drv As DataRow In dt.Rows
                'signUpStatus 0:?嗡辣摰? 1:?勗??? 2:?芷???3:甇?? 4:?? 5:?芷???
                Select Case Convert.ToString(drv("signUpStatus"))
                    Case "0", "1", "3"
                        'FROM 
                        '瑼Ｘ?臬??銴?隤脩??? true:??false:瘝? '?? gsPTDID
                        If TIMS.Chk_DoubleDESC(drv("OCID"), gssOCID, dtTrain, gsPTDID) Then chkDouble = True
                End Select
            Next
        End If
        sMsg &= "嚗d:" & Convert.ToString(chkDouble)
        If Convert.ToString(Request("A")) = "Y" Then
            'sDate = String.Empty
            'eDate = String.Empty
            'Call TIMS.Get_SubSidyCostDay(rqIDNO, oSTDate, sDate, eDate, objconn)
            Me.labMoneyShow1.Text = TIMS.Check_GovCost(rqIDNO, rqOCID1, iOver6w, objconn, "Y")
            'TIMS.Tooltip(labMsg2, sMsg2)
        End If

        Dim dtCC As DataTable
        dtCC = TIMS.GetRelEnterDateDtY3d(rqIDNO, rqBirth, gsPTDID, objconn)
        RecordCount.Text = dtCC.Rows.Count

        msgbb.Text = cst_msgTxt1 '"?⊿?銴?閮???!"
        'msg.Text = "?∪?閮???!"
        PageControler1.Visible = False
        Me.DataGrid2bb.Visible = False
        If dtCC.Rows.Count > 0 Then
            msgbb.Text = ""
            PageControler1.Visible = True
            Me.DataGrid2bb.Visible = True
            PageControler1.PageDataTable = dtCC
            PageControler1.ControlerLoad()
        End If

        If dtCC.Rows.Count = 0 Then
            '?⊿?銴?閮???
            Dim iCnt As Integer = dtCC.Rows.Count '(?桀???鞈?蝑)
            Call SaveDa1Del28DBL(iCnt, rqIDNO, rqOCID1)
        End If
        'If gsPTDID <> "" AndAlso chkDouble AndAlso dtCC.Rows.Count > 0 Then
        '    msg.Text = ""
        '    PageControler1.Visible = True
        '    Me.DataGrid2.Visible = True
        '    PageControler1.PageDataTable = dtCC
        '    PageControler1.ControlerLoad()
        'End If
        'If HidDouble.Value = "1" Then
        '    Common.MessageBox(Me, cst_msg1)
        '    'Exit Sub
        'End If
        labOver6w.Text = cst_msgover6w '"餈?撟游鋆鞎颱蝙???恍?隡?嚗?

        Const Cst_AlertCost As Integer = 60000 '霅衣內憿?
        If iOver6w >= Cst_AlertCost Then
            labOver6w.Text &= "<font color=Red>"
            labOver6w.Text &= CStr(iOver6w) & "??
            'labOver6w.Text &= CStr(Cst_AlertCost) & "??
            labOver6w.Text &= "</font>"
            'labOver6w.Visible = True
            'labOver6w.ForeColor = Color.Red
            Common.MessageBox(Me, cst_msg6)
            'Exit Sub
        Else
            labOver6w.Text &= CStr(iOver6w) & "??
        End If
        If Convert.ToString(Request("A")) = "Y" Then
            'TIMS.Tooltip(labOver6w, sMsg)
            labOver6w.Text &= sMsg
        End If
    End Sub

    '?芷??鞈?(?桀??亥岷鞈?銝衣??鞈?)
    Sub SaveDa1Del28DBL(ByVal iCnt As Integer, ByVal aIDNO As String, ByVal aOCID1 As String)
        If iCnt > 0 Then Exit Sub '??銴?閮??
        If Not initObj28() Then Exit Sub '?啣虜?瘜??

        Dim sql As String = ""
        sql = ""
        sql &= " SELECT ESERNUM FROM STUD_ENTERDOUBLE WHERE IDNO = @IDNO AND OCID1 = @OCID1 AND ETYPE1 = 'Y' "
        sql &= " UNION SELECT ESERNUM FROM STUD_ENTERDOUBLE WHERE IDNO = @IDNO AND OCID2 = @OCID1 AND ETYPE1 = 'Y' "
        Dim sCmd As New OracleCommand(sql, objconn)
        Call TIMS.OpenDbConn(objconn)
        Dim dtS As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("IDNO", OracleType.VarChar).Value = aIDNO
            .Parameters.Add("OCID1", OracleType.VarChar).Value = aOCID1
            dtS.Load(.ExecuteReader())
        End With
        If dtS.Rows.Count = 0 Then Exit Sub '?⊿?銴????

        Call TIMS.OpenDbConn(objconn)
        Dim tmpTrans As OracleTransaction = objconn.BeginTransaction()
        For Each dr As DataRow In dtS.Rows
            Try
                Dim sqlAdp As New OracleDataAdapter
                Dim sqlStr As String = ""
                sqlStr = " UPDATE STUD_ENTERDOUBLE SET MODIFYACCT = @MODIFYACCT, MODIFYDATE = GETDATE() WHERE ESERNUM = @ESERNUM "
                With sqlAdp
                    .UpdateCommand = New OracleCommand(sqlStr, objconn, tmpTrans)
                    .UpdateCommand.Parameters.Clear()
                    .UpdateCommand.Parameters.Add("MODIFYACCT", OracleType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                    .UpdateCommand.Parameters.Add("ESERNUM", OracleType.VarChar).Value = dr("ESERNUM")
                    .UpdateCommand.ExecuteNonQuery()
                End With
                'sqlStr = "INSERT INTO STUD_ENTERDOUBLEDELDATA SELECT * FROM STUD_ENTERDOUBLE WHERE ESERNUM= @ESERNUM"
                sqlStr = "" & vbCrLf
                sqlStr &= " INSERT INTO STUD_ENTERDOUBLEDELDATA (ESERNUM ,ESETID,IDNO,OCID1,OCID2,PTDID1,PTDID2,PNAME1,PNAME2,SUMOFMONEY,ETYPE1,ETYPE2,EMAIL,ISSEND1,ISSEND2,ISSEND3,CREATEACCT,CREATEDATE,MODIFYACCT,MODIFYDATE,SENDMAILDATE) " & vbCrLf
                sqlStr &= " SELECT ESERNUM,ESETID,IDNO,OCID1,OCID2,PTDID1,PTDID2,PNAME1,PNAME2,SUMOFMONEY,ETYPE1,ETYPE2,EMAIL,ISSEND1,ISSEND2,ISSEND3,CREATEACCT,CREATEDATE,MODIFYACCT,MODIFYDATE,SENDMAILDATE " & vbCrLf
                sqlStr &= " FROM STUD_ENTERDOUBLE WHERE ESERNUM = @ESERNUM "
                With sqlAdp
                    .InsertCommand = New OracleCommand(sqlStr, objconn, tmpTrans)
                    .InsertCommand.Parameters.Clear()
                    .InsertCommand.Parameters.Add("ESERNUM", OracleType.VarChar).Value = dr("ESERNUM")
                    .InsertCommand.ExecuteNonQuery()
                End With
                sqlStr = " DELETE STUD_ENTERDOUBLE WHERE ESERNUM = @ESERNUM "
                With sqlAdp
                    .DeleteCommand = New OracleCommand(sqlStr, objconn, tmpTrans)
                    .DeleteCommand.Parameters.Clear()
                    .DeleteCommand.Parameters.Add("ESERNUM", OracleType.VarChar).Value = dr("ESERNUM")
                    .DeleteCommand.ExecuteNonQuery()
                End With
            Catch ex As Exception
                Dim strErrmsg As String = ""
                strErrmsg += "/* Sub SaveDa1Del28DBL(ByVal iCnt As Integer, ByVal aIDNO As String, ByVal aOCID1 As String) */" & vbCrLf
                strErrmsg += "/* ex.ToString: */" & ex.ToString & vbCrLf
                strErrmsg += "aIDNO: " & aIDNO & vbCrLf
                strErrmsg += "aOCID1: " & aOCID1 & vbCrLf
                strErrmsg += "ESERNUM: " & dr("ESERNUM") & vbCrLf
                strErrmsg += TIMS.GetErrorMsg(Me) '???航炊鞈?撖怠
                strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                Call TIMS.SendMailTest(strErrmsg)
                tmpTrans.Rollback()
                Exit Sub
            End Try
        Next
        tmpTrans.Commit()
    End Sub

    '??
    Private Sub Btnclose_ServerClick(sender As Object, e As System.EventArgs) Handles Btnclose.ServerClick
        Page.RegisterStartupScript("History", "<script>window.close();</script>")
    End Sub

    '蝯? '?交?-銝玨??
    Function Get_TRAINDESCtb(ByVal OCID As String, ByVal sPTDID As String, ByVal dtTrain As DataTable, ByVal iRow As Integer) As String
        Dim rst As String = ""
        If sPTDID = "" Then Return rst
        Dim ff As String = ""
        ff = ""
        ff &= " OCID = " & OCID
        ff &= " AND PTDID IN (" & sPTDID & ") "
        Dim ss As String = "PTDID"
        If dtTrain.Select(ff).Length > 0 Then
            rst &= "<table class=""font"" cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">"
            For Each dr As DataRow In dtTrain.Select(ff, ss)
                rst &= "<tr>"
                'If iRow Mod 2 = 0 Then
                '    rst &= "<tr style=""background-color@WhiteSmoke;"">"
                'Else
                '    rst &= "<tr>"
                'End If
                rst &= "<td>" & Convert.ToString(dr("STRAINDATE")) & "</td>"
                rst &= "<td>" & Convert.ToString(dr("PNAME")) & "</td>"
                rst &= "</tr>"
            Next
            rst &= "</table>"
        End If
        Return rst
    End Function

    Private Sub DataGrid2bb_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2bb.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                'e.Item.Cells(0).Text = e.Item.ItemIndex + 1 + sender.PageSize * sender.CurrentPageIndex
                Dim Labsignno As Label = e.Item.FindControl("Labsignno")
                Dim Labdouble As Label = e.Item.FindControl("Labdouble")
                Dim Literal1 As Literal = e.Item.FindControl("Literal1") '?交?-銝玨??
                Dim Labstudstatus As Label = e.Item.FindControl("Labstudstatus") '閮毀&lt;BR&gt;???
                Labsignno.Text = e.Item.ItemIndex + 1 + sender.PageSize * sender.CurrentPageIndex
                Labdouble.Visible = False '????閮???
                Literal1.Text = Get_TRAINDESCtb(CStr(drv("OCID")), gsPTDID, dtTrain, Val(Labsignno.Text))
                If Literal1.Text <> "" Then Labdouble.Visible = True 'False'????閮???
                Labstudstatus.Text = "?勗?銝?
                ff = "OCID=" & CStr(drv("OCID"))
                If dtStud.Select(ff).Length > 0 Then Labstudstatus.Text = CStr(dtStud.Select(ff)(0)("STUDSTATUS2"))
        End Select
    End Sub
End Class