Partial Class SD_01_004_add
    Inherits AuthBasePage

    'Dim OCID_value, IDNO_value, BFDate_value, STDate_Value As String 'Dim OCID_value As String
    '#Region "TPlanID28DBL" '參考 SD_01_004_dbl.aspx
    Dim ff3 As String = "" '過濾字
    Dim fss As String = "" '過濾字
    Dim gssOCID As String = ""
    Dim gsPTDID As String = "" '所有重複的課程
    Dim iOverSUMOFMONEY As Double = 0 '使用補助金(產投充飛)

    Dim flg_ShowTID28DBL As Boolean = False '是否要顯示產投重疊資訊。

    Dim flag_show_actno_budid As Boolean = False '保險證號/預算別代碼 false:不顯示 true:顯示

    Const cst_Regfull As String = "本班報名已額滿"

    Dim rqOCID1 As String = ""
    Dim rqIDNO As String = ""
    Dim rqBirth As String = ""

    Dim sDate As String = String.Empty
    Dim eDate As String = String.Empty
    Dim oSTDate As String = ""

    Const cst_msgTxt1 As String = "查無產投方案上課時段重疊紀錄!!"

    Dim dtStdCost As DataTable
    Dim dtStud As DataTable
    Dim dtTrain As DataTable

    Dim cst_dg4_甄試日期 As Int16 = 4
    Dim cst_dg4_開訓日期 As Int16 = 5
    Dim cst_dg4_結訓日期 As Int16 = 6

    '#Region "TPlanID28DBL 2"
    '執行正確為True 異False
    Function initObj28() As Boolean
        Dim rst As Boolean = False '異False

        Dim ReqeSETID As String = TIMS.ClearSQM(Request("eSETID"))
        Dim ReqeSerNum As String = TIMS.ClearSQM(Request("eSerNum"))
        Dim dr As DataRow = Nothing
        If ReqeSETID <> "" Then dr = TIMS.Get_ENTERTEMP2(ReqeSETID, objconn)
        If dr Is Nothing Then
            Common.MessageBox(Me, cst_errmsg4)
            Return rst 'Exit Function
        End If
        Dim dr1 As DataRow = Nothing
        If ReqeSerNum <> "" Then dr1 = TIMS.Get_ENTERTYPE2(ReqeSerNum, objconn)
        If dr1 Is Nothing Then
            Common.MessageBox(Me, cst_errmsg4)
            Return rst 'Exit Function
        End If
        Hid_eSerNum.Value = ReqeSerNum
        rqOCID1 = TIMS.ClearSQM(Request("OCID1"))
        rqIDNO = TIMS.ClearSQM(Request("IDNO"))
        If rqIDNO = "" OrElse rqIDNO <> Convert.ToString(dr("IDNO")) Then
            Common.MessageBox(Me, cst_errmsg4)
            Return rst 'Exit Function
        End If
        If rqOCID1 = "" OrElse rqOCID1 <> Convert.ToString(dr1("OCID1")) Then
            Common.MessageBox(Me, cst_errmsg4)
            Return rst 'Exit Function
        End If
        If ReqeSETID <> Convert.ToString(dr1("eSETID")) Then
            Common.MessageBox(Me, cst_errmsg4)
            Return rst 'Exit Function
        End If
        rqBirth = TIMS.Cdate3(dr("BIRTHDAY")) '西元年 yyyy/MM/dd
        'LabName.Text=Convert.ToString(dr("NAME"))
        'TPlanName.Text=TIMS.GetPlanName(sm.UserInfo.PlanID, objconn)

        'WebRequest物件如何忽略憑證問題
        System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf TIMS.ValidateServerCertificate)
        'TLS 1.2-基礎連接已關閉: 傳送時發生未預期的錯誤 
        System.Net.ServicePointManager.SecurityProtocol = Net.SecurityProtocolType.Tls12
        '檢核學員重複參訓。
        'http://163.29.199.211/TIMSWS/timsService1.asmx
        'https://wltims.wda.gov.tw/TIMSWS/timsService1.asmx
        Dim timsSer1 As New timsService1.timsService1

        Dim ERRMSG As String = ""
        Dim xStudInfo As String = ""
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            TIMS.SetMyValue(xStudInfo, "IDNO", rqIDNO)
            TIMS.SetMyValue(xStudInfo, "OCID1", rqOCID1)
            '檢核學員OOO有參訓時段重疊報名情形，無法儲存，請再確認！
            Call TIMS.ChkStudDouble(timsSer1, ERRMSG, "", xStudInfo)
        End If
        rst = True '正確為True
        Return rst 'Exit Function
    End Function

    '搜尋(重複資訊)。
    Sub Search28DBL()
        Dim aNow As Date = TIMS.GetSysDateNow(objconn)
        Dim sMsg As String = ""
        Dim dr1 As DataRow = TIMS.GetOCIDDate(rqOCID1, objconn)
        oSTDate = Convert.ToString(TIMS.Cdate3(dr1("STDate"))) '西元年yyyy/MM/dd
        sDate = String.Empty
        eDate = String.Empty
        Call TIMS.Get_SubSidyCostDay(rqIDNO, oSTDate, sDate, eDate, objconn)
        sMsg &= String.Concat("，rqOCID1:", rqOCID1, "，oSTDate:", oSTDate, "，sD:", sDate, "，eD:", eDate)

        '取出 補助金申請查詢 [SQL] (該學員的所有補助金資料)
        dtStdCost = TIMS.GetHistoryDefStdCostDt2(rqIDNO, rqBirth, objconn)
        sMsg &= String.Concat("，Cost_CNT:", dtStdCost.Rows.Count, "(", TIMS.Get_ssOCID(dtStdCost), ")")

        '(近3年) 取得該學員所有報名日期。(產投限定)
        '1.sDate eDate 為必填值 2.因為有可能沒有參訓資料，所以資料往前推3年
        Dim dt As DataTable = TIMS.GetRelEnterDateDtY3(rqIDNO, rqBirth, sDate, eDate, objconn)
        sMsg &= String.Concat("，EY3_CNT:", dt.Rows.Count, "(", TIMS.Get_ssOCID(dt), ")")

        '取得學員資料表。(該學員的所有資料)
        dtStud = TIMS.Get_StudInfo(rqIDNO, objconn)
        sMsg &= String.Concat("，Stud_CNT:", dtStud.Rows.Count, "(", TIMS.Get_ssOCID(dtStud), ")")

        'HidDouble.Value="" '重複的課程時間 1.有重複 "".沒重複 '使用已達n萬元 1.達到" ".沒事
        labMoneyShow1.Text = ""
        Dim ReqA As String = TIMS.ClearSQM(Request("A"))

        Dim flagStd As Boolean = True '該學員無上課資料，無請領補助 有@true 沒有@false
        If dtStdCost.Rows.Count = 0 Then
            '近3年所有報名資料。
            gssOCID = TIMS.Get_ssOCID(dt) '取得目前table的所有ocid以逗點分隔
            dtTrain = TIMS.Get_TRAINDESC(gssOCID, dt, objconn)   '取得要比對的課程資料

            '近3年所有報名資料。(保留有效，刪除無效資料(報名失敗排除)
            gssOCID = TIMS.Get_ssOCID(dtTrain) '取得目前table的所有ocid以逗點分隔(排除一些沒有課程的資料)

            flagStd = False
            '該學員無上課資料
            If dt.Rows.Count > 0 Then
                Dim dr As DataRow
                '某限期內資料筆數('6個月內的報名資料，供計算使用金額比較使用。)
                '3年的範圍資料
                ff3 = "XDAY2=1 AND XDAY3=1"
                fss = "STDate"
                If dt.Select(ff3, fss).Length > 0 Then
                    '有條件的1筆資料
                    dr = dt.Select(ff3, fss)(0)
                    If ReqA = "Y" Then
                        oSTDate = Convert.ToString(dr("STDate"))
                        sDate = String.Empty
                        eDate = String.Empty
                        Call TIMS.Get_SubSidyCostDay(rqIDNO, dr("STDate"), sDate, eDate, objconn)
                        sMsg &= String.Concat("，1.o:", dr("OCID"), "，STDate:", dr("STDate"), "，sD:", sDate, "，eD:", eDate)
                    End If
                    labMoneyShow1.Text = TIMS.Check_GovCost(rqIDNO, rqOCID1, iOverSUMOFMONEY, objconn, "")
                Else
                    '無條件的最後1筆資料
                    dr = dt.Rows(dt.Rows.Count - 1)
                    If ReqA = "Y" Then
                        oSTDate = Convert.ToString(dr("STDate"))
                        sDate = String.Empty
                        eDate = String.Empty
                        Call TIMS.Get_SubSidyCostDay(rqIDNO, dr("STDate"), sDate, eDate, objconn)
                        sMsg &= String.Concat("，2.o:", dr("OCID"), "，STDate:", dr("STDate"), "，sD:", sDate, "，eD:", eDate)
                    End If
                    labMoneyShow1.Text = TIMS.Check_GovCost(rqIDNO, rqOCID1, iOverSUMOFMONEY, objconn, "")
                End If
            End If
        End If
        sMsg &= "，flagStd:" & Convert.ToString(flagStd)

        If flagStd AndAlso dt.Rows.Count > 0 Then
            '近3年所有報名資料。
            gssOCID = TIMS.Get_ssOCID(dt) '取得目前table的所有ocid以逗點分隔
            dtTrain = TIMS.Get_TRAINDESC(gssOCID, dt, objconn)   '取得要比對的課程資料

            '近3年所有報名資料。(保留有效，刪除無效資料(報名失敗排除)
            gssOCID = TIMS.Get_ssOCID(dtTrain) '取得目前table的所有ocid以逗點分隔 

            'labMoneyShow1.Text="" '實際核撥補助費用
            'labMoneyShow2.Text="" '參訓中課程預估補助費用
            'labMoneyShow3.Text="" '報名中課程預估補助費用
            Dim dr As DataRow
            '某限期內資料筆數('6個月內的報名資料，供計算使用金額比較使用。)
            '3年的範圍資料
            ff3 = "XDAY2=1 AND XDAY3=1 AND STUDSTATUS NOT IN (2,3)" 'FILTER
            fss = "STDate" 'SORT
            If dt.Select(ff3, fss).Length > 0 Then
                dr = dt.Select(ff3, fss)(0) '第1筆
                If ReqA = "Y" Then
                    '測試功能
                    oSTDate = Convert.ToString(dr("STDate"))
                    sDate = String.Empty
                    eDate = String.Empty
                    Call TIMS.Get_SubSidyCostDay(rqIDNO, dr("STDate"), sDate, eDate, objconn)
                    sMsg &= String.Concat("，3.o:", dr("OCID"), "，STDate:", dr("STDate"), "，sD:", sDate, "，eD:", eDate)
                End If
                labMoneyShow1.Text = TIMS.Check_GovCost(rqIDNO, rqOCID1, iOverSUMOFMONEY, objconn, "") '顯示目前使用金額。
            Else
                '6個月內無資料挑最後一筆資料
                dr = dt.Rows(dt.Rows.Count - 1)
                If ReqA = "Y" Then
                    '測試功能
                    oSTDate = Convert.ToString(dr("STDate"))
                    sDate = String.Empty
                    eDate = String.Empty
                    Call TIMS.Get_SubSidyCostDay(rqIDNO, dr("STDate"), sDate, eDate, objconn)
                    sMsg &= String.Concat("，4.(m6N)o:", dr("OCID"), "，STDate:", dr("STDate"), "，sD:", sDate, "，eD:", eDate)
                End If
                labMoneyShow1.Text = TIMS.Check_GovCost(rqIDNO, rqOCID1, iOverSUMOFMONEY, objconn, "")
            End If
        End If

        'If sYear2015Test="Y" Then  labMoneyShow3.Text="(測試環境文字測試)" & cst_titleS1 & cst_titleS2
        'RecordCount.Text=dt.Rows.Count
        'msg.Text="無參訓紀錄!!"
        'PageControler1.Visible=False
        'DataGrid2.Visible=False

        Dim chkDouble As Boolean = False '判斷有無重複資料。
        If dt.Rows.Count > 0 Then
            For Each drv As DataRow In dt.Rows
                Dim canDelete As Boolean = False
                Dim canUseDoubleCheck As Boolean = False
                'signUpStatus 0:收件完成 1:報名成功 2:報名失敗 3:正取(Key_SelResult) 4:備取 5:未錄取
                Select Case Convert.ToString(drv("signUpStatus"))
                    Case "0", "1", "3"
                        Select Case Convert.ToString(drv("STUDSTATUS"))
                            Case "2", "3"
                                canDelete = True
                        End Select
                    Case Else
                        canDelete = True
                End Select
                If Not canDelete Then canUseDoubleCheck = True '不刪除要比對是否重複
                If canUseDoubleCheck Then '比對是否重複
                    'FROM 
                    '檢查是否有重複的課程時間 true:有 false:沒有 '取得 gsPTDID
                    If TIMS.Chk_DoubleDESC(drv("OCID"), gssOCID, dtTrain, gsPTDID) Then
                        chkDouble = True
                    End If
                End If
            Next
        End If
        sMsg &= "，dd:" & Convert.ToString(chkDouble)
        If ReqA = "Y" Then
            'sDate=String.Empty
            'eDate=String.Empty
            'Call TIMS.Get_SubSidyCostDay(rqIDNO, oSTDate, sDate, eDate, objconn)
            labMoneyShow1.Text = TIMS.Check_GovCost(rqIDNO, rqOCID1, iOverSUMOFMONEY, objconn, "Y")
            'TIMS.Tooltip(labMsg2, sMsg2)
        End If

        Dim dtCC As DataTable = TIMS.GetRelEnterDateDtY3d(rqIDNO, rqBirth, gsPTDID, objconn)
        'RecordCount.Text=dtCC.Rows.Count
        msgbb.Text = cst_msgTxt1 '"無重複參訓紀錄!!"
        'msg.Text="無參訓紀錄!!"
        'PageControler1.Visible=False
        DataGrid2bb.Visible = False
        If dtCC.Rows.Count > 0 Then
            msgbb.Text = ""
            'PageControler1.Visible=True
            With DataGrid2bb
                .Visible = True
                .DataSource = dtCC
                .DataBind()
            End With
            'DataGrid2bb.Visible=True
            'PageControler1.PageDataTable=dtCC
            'PageControler1.ControlerLoad()
        End If

        If dtCC.Rows.Count = 0 Then
            '無重複參訓紀錄
            Dim iCnt As Integer = dtCC.Rows.Count '(目前重複資訊筆數)
            Call SaveDa1Del28DBL(iCnt, rqIDNO, rqOCID1)
        End If

        Dim sMessage As String = ""
        Dim fg_test As Boolean = TIMS.CHK_IS_TEST_ENVC() 'fg_test OrElse fg_use1 
        Dim fg_use1 As Boolean = TIMS.CanUse3Y10WCost()
        If fg_use1 OrElse fg_test Then
            labOver6w.Text = GET_labOver6w_TXT(iOverSUMOFMONEY, sMessage)
            'Else labOver6w.Text = GET_labOver6w_TXT_old1(iOverSUMOFMONEY, sMessage)
        End If
        If sMessage <> "" Then Common.MessageBox(Me, sMessage)

        If ReqA = "Y" Then
            'TIMS.Tooltip(labOver6w, sMsg)
            labOver6w.Text &= sMsg
        End If
    End Sub

    '提醒您，預估您目前補助費使用已達7萬元
    'Public Shared Function GET_labOver6w_TXT_old1(ByRef iOverSUMOFMONEY As Double, ByRef sMessage As String) As String
    '    Const cst_msg6_over As String = "提醒您，預估您目前補助費使用已達6萬元（包含已核撥、參訓中、已報名的課程），請您留意！"
    '    Const cst_msg7_over As String = "提醒您，預估您目前補助費使用已達7萬元（包含已核撥、參訓中、已報名的課程），請您留意！"
    '    Const cst_msgover6w As String = "近3年內補助費使用(含預估)："

    '    Const Cst_AlertCost_60k As Integer = 60000 '警示額 60k
    '    Const Cst_AlertCost_70k As Integer = 70000 '警示額 70k
    '    Const cst_ETYPE2_60K As String = "6" '達60K (6萬)
    '    Const cst_ETYPE2_70K As String = "7" '達70K (7萬)

    '    Dim rst_labOver6w As String = cst_msgover6w '"近3年內補助費使用(含預估)："

    '    Dim v_ETYPE2 As String = "" '補助費使用已達 XX 暫存欄 
    '    If Val(iOverSUMOFMONEY) >= Cst_AlertCost_70k Then v_ETYPE2 = cst_ETYPE2_70K
    '    If Val(iOverSUMOFMONEY) >= Cst_AlertCost_60k AndAlso Val(iOverSUMOFMONEY) < Cst_AlertCost_70k Then v_ETYPE2 = cst_ETYPE2_60K
    '    sMessage = ""
    '    Select Case v_ETYPE2
    '        Case cst_ETYPE2_60K
    '            rst_labOver6w &= String.Concat("<font color=Red>", iOverSUMOFMONEY, "元", "</font>")
    '            sMessage = cst_msg6_over
    '        'Exit Sub
    '        Case cst_ETYPE2_70K
    '            rst_labOver6w &= String.Concat("<font color=Red>", iOverSUMOFMONEY, "元", "</font>")
    '            sMessage = cst_msg7_over
    '        Case Else
    '            rst_labOver6w &= String.Concat(iOverSUMOFMONEY, "元")
    '    End Select
    '    Return rst_labOver6w
    'End Function

    ''' <summary>提醒您，預估您目前補助費使用已達10萬元</summary>
    ''' <param name="iOverSUMOFMONEY"></param>
    ''' <param name="sMessage"></param>
    ''' <returns></returns>
    Public Shared Function GET_labOver6w_TXT(ByRef iOverSUMOFMONEY As Double, ByRef sMessage As String) As String
        Const cst_msg9_over As String = "提醒您，預估您目前補助費使用已達9萬元（包含已核撥、參訓中、已報名的課程），請您留意！"
        Const cst_msg10_over As String = "提醒您，預估您目前補助費使用已達10萬元（包含已核撥、參訓中、已報名的課程），請您留意！"
        Const cst_msgover6w As String = "近3年內補助費使用(含預估)："
        Dim rst_labOver6w As String = cst_msgover6w '"近3年內補助費使用(含預估)："
        Const Cst_AlertCost_90k As Integer = 90000 '警示額 90k
        Const Cst_AlertCost_100k As Integer = 100000 '警示額 100k
        Const cst_ETYPE2_90K As String = "9" '達90K (9萬)
        Const cst_ETYPE2_100K As String = "10" '達100K (10萬)
        Dim v_ETYPE2 As String = "" '補助費使用已達 XX 暫存欄 
        If Val(iOverSUMOFMONEY) >= Cst_AlertCost_100k Then v_ETYPE2 = cst_ETYPE2_100K
        If Val(iOverSUMOFMONEY) >= Cst_AlertCost_90k AndAlso Val(iOverSUMOFMONEY) < Cst_AlertCost_100k Then v_ETYPE2 = cst_ETYPE2_90K
        sMessage = ""
        Select Case v_ETYPE2
            Case cst_ETYPE2_90K
                rst_labOver6w &= String.Concat("<font color=Red>", iOverSUMOFMONEY, "元", "</font>")
                sMessage = cst_msg9_over
            'Exit Sub
            Case cst_ETYPE2_100K
                rst_labOver6w &= String.Concat("<font color=Red>", iOverSUMOFMONEY, "元", "</font>")
                sMessage = cst_msg10_over
            Case Else
                rst_labOver6w &= String.Concat(iOverSUMOFMONEY, "元")
        End Select
        Return rst_labOver6w
    End Function

    '刪除重複資訊(目前查詢資料並無動複資訊)
    Sub SaveDa1Del28DBL(ByVal iCnt As Integer, ByVal aIDNO As String, ByVal aOCID1 As String)
        If iCnt > 0 Then Exit Sub '有重複資訊離開
        If Not initObj28() Then Exit Sub '異常狀況離開。
        Dim sql As String = ""
        sql &= " SELECT ESERNUM FROM STUD_ENTERDOUBLE WHERE IDNO=@IDNO AND OCID1=@OCID1 AND ETYPE1='Y'"
        sql &= " UNION SELECT ESERNUM FROM STUD_ENTERDOUBLE WHERE IDNO=@IDNO AND OCID2=@OCID1 AND ETYPE1='Y'"
        Dim sCmd As New SqlCommand(sql, objconn)
        Call TIMS.OpenDbConn(objconn)
        Dim dtS As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("IDNO", SqlDbType.VarChar).Value = aIDNO
            .Parameters.Add("OCID1", SqlDbType.VarChar).Value = aOCID1
            dtS.Load(.ExecuteReader())
        End With
        If dtS.Rows.Count = 0 Then Exit Sub '無重複資料離開。
        'Call TIMS.OpenDbConn(objconn)
        Dim oConn As SqlConnection = DbAccess.GetConnection()
        Dim oTrans As SqlTransaction = DbAccess.BeginTrans(oConn)
        Try
            Dim sqlAdp As New SqlDataAdapter
            Dim sqlStr As String = ""
            sqlStr = " UPDATE STUD_ENTERDOUBLE SET MODIFYACCT=@MODIFYACCT, MODIFYDATE=GETDATE() WHERE ESERNUM=@ESERNUM "
            sqlAdp.UpdateCommand = New SqlCommand(sqlStr, oConn, oTrans)
            'sqlStr="INSERT INTO STUD_ENTERDOUBLEDELDATA SELECT * FROM STUD_ENTERDOUBLE WHERE ESERNUM= @ESERNUM"
            sqlStr = "" & vbCrLf
            sqlStr &= " INSERT INTO STUD_ENTERDOUBLEDELDATA (ESERNUM,ESETID,IDNO,OCID1,OCID2,PTDID1,PTDID2,PNAME1,PNAME2,SUMOFMONEY,ETYPE1,ETYPE2,EMAIL,ISSEND1,ISSEND2,ISSEND3,CREATEACCT,CREATEDATE,MODIFYACCT,MODIFYDATE,SENDMAILDATE)" & vbCrLf
            sqlStr &= " SELECT ESERNUM,ESETID,IDNO,OCID1,OCID2,PTDID1,PTDID2,PNAME1,PNAME2,SUMOFMONEY,ETYPE1,ETYPE2,EMAIL,ISSEND1,ISSEND2,ISSEND3,CREATEACCT,CREATEDATE,MODIFYACCT,MODIFYDATE,SENDMAILDATE" & vbCrLf
            sqlStr &= " FROM STUD_ENTERDOUBLE WHERE ESERNUM=@ESERNUM "
            sqlAdp.InsertCommand = New SqlCommand(sqlStr, oConn, oTrans)

            sqlStr = "DELETE STUD_ENTERDOUBLE WHERE ESERNUM=@ESERNUM "
            sqlAdp.DeleteCommand = New SqlCommand(sqlStr, oConn, oTrans)

            For Each dr As DataRow In dtS.Rows
                With sqlAdp.UpdateCommand
                    .Parameters.Clear()
                    .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                    .Parameters.Add("ESERNUM", SqlDbType.VarChar).Value = dr("ESERNUM")
                    .ExecuteNonQuery()
                End With
                With sqlAdp.InsertCommand
                    .Parameters.Clear()
                    .Parameters.Add("ESERNUM", SqlDbType.VarChar).Value = dr("ESERNUM")
                    .ExecuteNonQuery()
                End With
                With sqlAdp.DeleteCommand
                    .Parameters.Clear()
                    .Parameters.Add("ESERNUM", SqlDbType.VarChar).Value = dr("ESERNUM")
                    .ExecuteNonQuery()
                End With
            Next
            DbAccess.CommitTrans(oTrans)
        Catch ex As Exception
            Dim strErrmsg As String = ""
            strErrmsg &= "/* Sub SaveDa1Del28DBL(ByVal iCnt As Integer, ByVal aIDNO As String, ByVal aOCID1 As String) */" & vbCrLf
            strErrmsg &= "/* ex.ToString: */" & ex.ToString & vbCrLf
            strErrmsg &= "aIDNO: " & aIDNO & vbCrLf
            strErrmsg &= "aOCID1: " & aOCID1 & vbCrLf
            strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)

            DbAccess.RollbackTrans(oTrans)
            Call TIMS.CloseDbConn(oConn)
        End Try
        Call TIMS.CloseDbConn(oConn)
    End Sub

    ''' <summary>'組合'日期-上課時間</summary>
    ''' <param name="OCID"></param>
    ''' <param name="sPTDID"></param>
    ''' <param name="dtTrain"></param>
    ''' <param name="iRow"></param>
    ''' <returns></returns>
    Function Get_TRAINDESCtb(ByVal OCID As String, ByVal sPTDID As String, ByVal dtTrain As DataTable, ByVal iRow As Integer) As String
        Dim rst As String = ""
        If sPTDID = "" Then Return rst
        Dim ff As String = String.Concat(" OCID=", OCID, " AND PTDID IN (", sPTDID, ")")
        Dim ss As String = "PTDID"
        If dtTrain.Select(ff).Length > 0 Then
            rst &= "<table class=""font"" cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">"
            For Each dr As DataRow In dtTrain.Select(ff, ss)
                rst &= String.Concat("<tr><td>", dr("STRAINDATE"), "</td><td>", dr("PNAME"), "</td></tr>")
            Next
            rst &= "</table>"
        End If
        Return rst
    End Function

    Const cst_inline1 As String = "" 'inline
    Const cst_errFlag As String = "Y" 'HiderrFlag.Value '檢核若產生錯誤為Y
    Dim vsExamDate As String = "" '(跨 sub 使用)
    Dim vsOCID1val As String = "" '(跨 sub 使用)
    '小table
    Dim dtIdentity As DataTable = Nothing
    Dim dtTrade As DataTable = Nothing
    Dim dtZip As DataTable = Nothing
    'Dim dtSCJOB As DataTable=Nothing
    Dim gsEnterDate As String = "" '系統報名日期(全域)
    Const cst_errmsg1 As String = "報名資料異常(請重新操作)，若持續發生錯誤，請將資料提供給系統管理者。"
    Const cst_errmsg2 As String = "報名資料已被刪除(請重新操作)，若持續發生錯誤，請將資料提供給系統管理者。"
    Const cst_errmsg3 As String = "查無資料"
    Const cst_errmsg4 As String = "查無有效資料"
    Dim flag_ROC As Boolean = False '是否啟用民國年日期顯示機制

    '取得郵遞區號資料
    'Function Get_ZipName() As DataTable
    '    Dim rst As New DataTable
    '    Dim sql As String="SELECT CTID,ZIPCODE,ZIPNAME,CTNAME,ZNAME,LCID,KLNAME,LNAME FROM VIEW_ZIPNAME ORDER BY 1,2"
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
    '    Dim sql As String="SELECT TRADEID , '['+TRADEID+']'+TRADENAME TRADENAME FROM KEY_TRADE ORDER BY 1"
    '    Call TIMS.OpenDbConn(objconn)
    '    Dim oCmd As New SqlCommand(sql, objconn)
    '    With oCmd
    '        .Parameters.Clear()
    '        rst.Load(.ExecuteReader())
    '    End With
    '    Return rst
    'End Function

    Function Get_IdentityDt() As DataTable
        Dim rDt As DataTable = Nothing
        Dim strSql As String = ""
        '20090123 andy  edit 產投 2009年 身分別「就業保險被保險人非自願失業者」 名稱改為「非自願離職者」
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 _
            AndAlso CInt(sm.UserInfo.Years) > 2008 Then
            rDt = TIMS.Get_dtIdentity(2, objconn)
            'strSql=""
            'strSql &= " select IdentityID"
            'strSql &= " ,case when IdentityID='02' then  N'非自願離職者' else Name end Name "
            'strSql &= " from Key_Identity "
            'strSql &= " ORDER BY 1"
        Else
            rDt = TIMS.Get_dtIdentity(0, objconn)
            'strSql=""
            'strSql &= " select IdentityID"
            'strSql &= " ,Name "
            'strSql &= " from Key_Identity "
            'strSql &= " ORDER BY 1"
        End If
        'Call TIMS.OpenDbConn(objconn)
        'Dim oCmd As New SqlCommand(strSql, objconn)
        'With oCmd
        '    .Parameters.Clear()
        '    rst.Load(.ExecuteReader())
        'End With
        Return rDt
    End Function

    '檢查 內網已有資料
    Function Check_EnterType(ByVal sIDNO As String, ByVal signUpStatus As String, ByVal OCID1 As String) As String
        Dim rst As Boolean = False
        'signUpStatus 0:收件完成 1:報名成功 2:報名失敗 3:正取(Key_SelResult) 4:備取 5:未錄取
        If signUpStatus = "1" Then Return rst
        Dim sql As String = ""
        sql = ""
        sql &= " SELECT 'x' x" & vbCrLf
        sql &= " FROM Stud_EnterTemp a" & vbCrLf
        sql &= " JOIN Stud_EnterType b ON a.SETID=b.SETID" & vbCrLf
        sql &= " WHERE a.IDNO=@IDNO AND b.OCID1=@OCID1" & vbCrLf
        Dim dt1 As New DataTable
        TIMS.OpenDbConn(objconn)
        Dim oCmd As New SqlCommand(sql, objconn)
        With oCmd
            .Parameters.Clear()
            .Parameters.Add("IDNO", SqlDbType.VarChar).Value = sIDNO
            .Parameters.Add("OCID1", SqlDbType.VarChar).Value = OCID1
            dt1.Load(.ExecuteReader())
        End With
        If dt1.Rows.Count > 0 Then rst = True
        Return rst
    End Function

    '刪除 (含還原刪除)
    Sub Del_StudSelResult(ByVal tmpSETID As Integer, ByVal tmpOCID As Integer, ByVal tmpConn As SqlConnection)
        Dim myParam As Hashtable = New Hashtable
        Dim sqlStr As String = ""
        'Dim sqlAdp As New SqlDataAdapter
        sqlStr = ""
        sqlStr &= " UPDATE STUD_SELRESULT SET ModifyAcct=@ModifyAcct, ModifyDate=GETDATE()" & vbCrLf
        sqlStr &= " WHERE SETID=@SETID AND OCID=@OCID" & vbCrLf
        myParam.Clear()
        myParam.Add("ModifyAcct", Convert.ToString(sm.UserInfo.UserID))
        myParam.Add("SETID", tmpSETID)
        myParam.Add("OCID", tmpOCID)
        DbAccess.ExecuteNonQuery(sqlStr, tmpConn, myParam)

        sqlStr = ""
        sqlStr &= " INSERT INTO Stud_SelResultDelData" & vbCrLf
        sqlStr &= " SELECT * FROM Stud_SelResult WHERE SETID=@SETID AND OCID=@OCID" & vbCrLf
        'Dim myParam As Hashtable=New Hashtable
        myParam.Clear()
        myParam.Add("SETID", tmpSETID)
        myParam.Add("OCID", tmpOCID)
        DbAccess.ExecuteNonQuery(sqlStr, tmpConn, myParam)

        sqlStr = " DELETE STUD_SELRESULT WHERE SETID=@SETID AND OCID=@OCID "
        myParam.Clear()
        myParam.Add("SETID", tmpSETID)
        myParam.Add("OCID", tmpOCID)
        DbAccess.ExecuteNonQuery(sqlStr, tmpConn, myParam)
    End Sub

    'list Stud_EnterType2.eSerNum STUD_ENTERSUBDATA2
    Sub SHOW_ENTERTYPE2()
        HiderrFlag.Value = ""
        If Request("eSerNum") <> "" Then heSerNum.Value = Request("eSerNum")
        If Request("eSETID") <> "" Then heSETID.Value = Request("eSETID")
        heSerNum.Value = TIMS.ClearSQM(heSerNum.Value)
        heSETID.Value = TIMS.ClearSQM(heSETID.Value)
        Dim ERRMSG As String = ""
        If heSerNum.Value = "" OrElse heSETID.Value = "" Then
            HiderrFlag.Value = cst_errFlag
            ERRMSG = "資料有誤請重新查詢!" & vbCrLf
            Common.MessageBox(Me, ERRMSG)
            Exit Sub
        End If
        Dim drET2 As DataRow = TIMS.Get_ENTERTYPE2(heSerNum.Value, objconn)
        If drET2 Is Nothing Then
            HiderrFlag.Value = cst_errFlag
            ERRMSG = "資料有誤請重新查詢!" & vbCrLf
            Common.MessageBox(Me, ERRMSG)
            Exit Sub
        End If
        If Convert.ToString(drET2("eSETID")) <> heSETID.Value Then
            HiderrFlag.Value = cst_errFlag
            ERRMSG = "資料有誤請重新查詢!" & vbCrLf
            Common.MessageBox(Me, ERRMSG)
            Exit Sub
        End If
        If ERRMSG <> "" Then
            HiderrFlag.Value = cst_errFlag
            Common.MessageBox(Me, ERRMSG)
            Exit Sub
        End If
        'heSerNum.Value=Request("eSerNum")
        'heSerNum.Value=TIMS.ClearSQM(heSerNum.Value)
        'sql=""
        'sql &= " SELECT a.*" & vbCrLf
        'sql &= " ,b.*" & vbCrLf
        Dim stdBLACK2TPLANID As String = ""
        Dim iStdBlackType As Integer = TIMS.Chk_StdBlackType(Me, objconn, stdBLACK2TPLANID)
        Dim sqlWSB As String = TIMS.Get_StdBlackWSB(Me, iStdBlackType, stdBLACK2TPLANID, 1)
        'sql &= sqlWSB
        Dim dr As DataRow = Nothing
        Dim sql As String = ""
        sql &= sqlWSB
        sql &= " SELECT a.ESETID" & vbCrLf '  /*PK*/
        sql &= " ,a.SETID ,a.IDNO ,a.NAME ,a.SEX ,a.BIRTHDAY ,a.PASSPORTNO ,a.MARITALSTATUS ,a.DEGREEID" & vbCrLf
        sql &= " ,a.GRADID ,a.SCHOOL ,a.DEPARTMENT ,a.MILITARYID ,a.ZIPCODE ,a.ADDRESS ,a.PHONE1 ,a.PHONE2" & vbCrLf
        sql &= " ,a.CELLPHONE ,a.EMAIL ,a.ISAGREE ,a.ZIPCODE6W" & vbCrLf
        sql &= " ,b.ESERNUM" & vbCrLf '/*PK*/
        sql &= " ,b.ENTERDATE ,b.RELENTERDATE ,b.EXAMNO ,b.OCID1 ,b.OCID2 ,b.OCID3 ,b.IDENTITYID" & vbCrLf
        sql &= " ,b.SIGNUPSTATUS ,b.SIGNUPMEMO ,b.SUPPLYID ,b.BUDID" & vbCrLf
        sql &= " ,format(b.ModifyDate,'yyyy-MM-dd HH:mm') LastModifyDate" & vbCrLf
        'sql &= " ,b.WORKSUPPIDENT" & vbCrLf'sql &= " ,b.USERNOSHOW" & vbCrLf'sql &= " ,b.NOTES" & vbCrLf'sql &= " ,b.ISEMAILFAIL" & vbCrLf'sql &= " ,b.SIGNNO" & vbCrLf
        'sql &= " ,b.INVOLLEAVER" & vbCrLf'sql &= " ,b.CFIRE1" & vbCrLf'sql &= " ,b.CFIRE1NS" & vbCrLf'sql &= " ,b.CFIRE1REASON" & vbCrLf'sql &= " ,b.CFIRE1MACCT" & vbCrLf
        'sql &= " ,b.CFIRE1MDATE" & vbCrLf'sql &= " ,b.CMASTER1" & vbCrLf'sql &= " ,b.CMASTER1NS" & vbCrLf'sql &= " ,b.CMASTER1REASON" & vbCrLf
        'sql &= " ,b.CMASTER1MACCT" & vbCrLf'sql &= " ,b.CMASTER1MDATE" & vbCrLf'sql &= " ,b.CMASTER1NT" & vbCrLf'sql &= " ,b.CFIRE1R2" & vbCrLf 'sql &= " ,b.PREEXDATE" & vbCrLf
        '屆退官兵身分者 'Session retreat Soldiers The identity of persons
        sql &= " ,CASE WHEN CONVERT(DATE, b.PREEXDATE) > CONVERT(DATE, GETDATE()) THEN 'Y' END SRSOLDIERS" & vbCrLf
        sql &= " ,f.examdate ,c.Name DegreeName ,d.Name GradName ,e.Name MilitaryName" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(f.ClassCName,f.CyclType) CLASSCNAME2" & vbCrLf
        'sql &= " ,f.ClassCName ClassCName1" & vbCrLf 'sql &= " ,f.CyclType CyclType1" & vbCrLf
        sql &= " ,f.STDate" & vbCrLf
        sql &= " ,DATEADD(month, -6, f.STDate) BFDate" & vbCrLf

        '有多筆 學員處分資料
        sql &= " ,CASE WHEN sb.IDNO IS NOT NULL THEN 'Y' ELSE 'N' END IsStdBlack" & vbCrLf
        sql &= " ,i.LevelName" & vbCrLf
        sql &= " ,se.ActNo ,se.PriorWorkType1 PriorWorkType2" & vbCrLf
        sql &= " ,se.PriorWorkOrg1 PriorWorkOrg2 ,se.SOfficeYM1 SOfficeYM2 ,se.FOfficeYM1 FOfficeYM2 ,sb6.ACTNO ACTNObli" & vbCrLf 'ACTNObli
        sql &= " FROM STUD_ENTERTEMP2 a WITH(NOLOCK)" & vbCrLf
        sql &= " JOIN STUD_ENTERTYPE2 b WITH(NOLOCK) ON a.eSETID=b.eSETID" & vbCrLf
        sql &= " JOIN CLASS_CLASSINFO f WITH(NOLOCK) ON b.OCID1=f.OCID" & vbCrLf
        sql &= " JOIN ID_PLAN ip WITH(NOLOCK) ON ip.planid=f.planid" & vbCrLf

        sql &= " LEFT JOIN KEY_DEGREE c WITH(NOLOCK) ON a.DegreeID=c.DegreeID" & vbCrLf
        sql &= " LEFT JOIN KEY_GRADSTATE d WITH(NOLOCK) ON a.GradID=d.GradID" & vbCrLf
        sql &= " LEFT JOIN KEY_MILITARY e WITH(NOLOCK) ON a.MilitaryID=e.MilitaryID" & vbCrLf
        sql &= " LEFT JOIN CLASS_CLASSLEVEL i WITH(NOLOCK) ON b.OCID1=i.OCID AND b.CCLID=i.CCLID" & vbCrLf

        '有多筆 學員處分資料
        sql &= " LEFT JOIN WSB sb ON sb.IDNO=a.IDNO" & vbCrLf
        sql &= " LEFT JOIN STUD_ENTERSUBDATA2 se ON se.eSerNum=b.eSerNum" & vbCrLf
        sql &= " LEFT JOIN STUD_BLIGATEDATA06 sb6 ON sb6.ESERNUM=b.ESERNUM" & vbCrLf
        sql &= $" WHERE b.eSerNum={heSerNum.Value}" & vbCrLf
        If sm.UserInfo.RID = "A" Then
            sql &= " AND ip.TPlanID='" & sm.UserInfo.TPlanID & "'" & vbCrLf
            sql &= " AND ip.Years='" & sm.UserInfo.Years & "'" & vbCrLf
        Else
            sql &= " AND b.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
        End If
        dr = DbAccess.GetOneRow(sql, objconn)
        If dr Is Nothing Then
            heSerNum.Value = ""
            IDNOValue.Value = ""
            STDateValue.Value = ""
            'Common.RespWrite(Me, "<script>alert('查無資料');history.go(-1);</script>")
            Label1.Text = "(請確認業務權限)" '報名班級
            Label1.ForeColor = Color.Red
            Common.MessageBox(Me, cst_errmsg3)
            Exit Sub
        End If

        'NOT dr Is Nothing 
        Dim sZIPCODE As String = Convert.ToString(dr("ZIPCODE"))
        gsEnterDate = TIMS.Cdate3(dr("ENTERDATE"))
        If Hid_MSG1.Value = "" Then Hid_MSG1.Value = TIMS.SHOW_ZIP2MSG(Me, sZIPCODE, gsEnterDate, objconn)

        'https://jira.turbotech.com.tw/browse/TIMSC-150
        Dim iADID As Integer = 0
        Dim flagMSG2 As Boolean = TIMS.CHK_DIS2MSG(Me, sZIPCODE, gsEnterDate, objconn, iADID)
        If flagMSG2 AndAlso Hid_MSG2.Value = "" Then
            Hid_MSGADIDN.Value = iADID
            Hid_MSG2.Value = TIMS.SHOW_DIS2MSG(Me, sZIPCODE, gsEnterDate, iADID, objconn)
        End If

        'WebRequest物件如何忽略憑證問題
        System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf TIMS.ValidateServerCertificate)
        'TLS 1.2-基礎連接已關閉: 傳送時發生未預期的錯誤 
        System.Net.ServicePointManager.SecurityProtocol = Net.SecurityProtocolType.Tls12
        '檢核學員重複參訓。
        'http://163.29.199.211/TIMSWS/timsService1.asmx
        'https://wltims.wda.gov.tw/TIMSWS/timsService1.asmx
        Dim timsSer1 As New timsService1.timsService1

        'Dim ERRMSG As String=""
        Dim rqIDNO As String = TIMS.ClearSQM(dr("IDNO"))
        Dim rqOCID1 As String = TIMS.ClearSQM(dr("OCID1"))
        Dim xStudInfo As String = ""
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            If rqIDNO <> "" AndAlso rqOCID1 <> "" Then
                Dim tERRMSG As String = ""
                '檢核學員重複參訓。
                TIMS.SetMyValue(xStudInfo, "IDNO", rqIDNO)
                TIMS.SetMyValue(xStudInfo, "OCID1", rqOCID1)
                '檢核學員OOO有參訓時段重疊報名情形，無法儲存，請再確認！
                Call TIMS.ChkStudDouble(timsSer1, tERRMSG, "", xStudInfo)
                Common.MessageBox(Me, tERRMSG)
            End If
        End If

        ''WorkSuppIdent
        'If IsDBNull(dr("WorkSuppIdent"))=False Then Common.SetListItem(WorkSuppIdent, dr("WorkSuppIdent"))
        'Dim vsExamDate As String=""
        vsExamDate = ""
        IDNOValue.Value = ""
        'Dim vsOCID1val As String=""
        vsOCID1val = ""
        If Convert.ToString(dr("examdate")) <> "" Then vsExamDate = Common.FormatDate(dr("examdate"))
        IDNOValue.Value = Convert.ToString(dr("IDNO"))
        'If IDNOValue.Value <> "" Then IDNOValue.Value=Trim(IDNOValue.Value)
        'If IDNOValue.Value <> "" Then IDNOValue.Value=UCase(IDNOValue.Value)
        If IDNOValue.Value <> "" Then IDNOValue.Value = TIMS.ChangeIDNO(IDNOValue.Value)
        vsOCID1val = Convert.ToString(dr("OCID1"))
        If IsDBNull(dr("STDate")) = False Then STDateValue.Value = FormatDateTime(Convert.ToString(dr("STDate")), DateFormat.ShortDate) Else STDateValue.Value = ""
        If IsDBNull(dr("BFDate")) = False Then BFDateValue.Value = FormatDateTime(Convert.ToString(dr("BFDate")), DateFormat.ShortDate) Else BFDateValue.Value = ""
        'IDNOValue.Value=TIMS.ChangeIDNO(dr("IDNO").ToString)
        IDNO.Text = TIMS.ChangeIDNO(dr("IDNO").ToString)
        'If IDNO.Text <> "" Then IDNO.Text=Trim(IDNO.Text)
        'If IDNO.Text <> "" Then IDNO.Text=UCase(IDNO.Text)
        'If IDNO.Text <> "" Then IDNO.Text=TIMS.ChangeIDNO(IDNO.Text)

        'start受訓前學員任職資料
        Select Case dr("PriorWorkType2").ToString
            Case "1"
                PriorWorkType1.Text = "曾工作過"
            Case "2"
                PriorWorkType1.Text = "未曾工作過"
            Case "3"
                PriorWorkType1.Text = "先前從事為非勞保性質工作"
            Case Else
                PriorWorkType1.Text = TIMS.cst_NODATAMsg12
        End Select

        PriorWorkOrg1.Text = TIMS.cst_NODATAMsg12
        If Convert.ToString(dr("PriorWorkOrg2")) <> "" Then PriorWorkOrg1.Text = Convert.ToString(dr("PriorWorkOrg2"))
        ActNo.Text = TIMS.cst_NODATAMsg12
        If Convert.ToString(dr("ActNo")) <> "" Then ActNo.Text = Convert.ToString(dr("ActNo"))
        'ACTNObli
        Hid_ACTNObli.Value = Convert.ToString(dr("ACTNObli"))
        OfficeDate.Text = TIMS.cst_NODATAMsg12
        If Convert.ToString(dr("SOfficeYM2")) <> "" AndAlso Convert.ToString(dr("FOfficeYM2")) <> "" Then OfficeDate.Text = Format(dr("SOfficeYM2"), "yyyy/MM/dd") & "~" & Format(dr("FOfficeYM2"), "yyyy/MM/dd")
        'end 受訓前學員任職資料

        Name.Text = dr("Name").ToString
        Birthday.Text = If(flag_ROC, TIMS.Cdate17(dr("Birthday")), TIMS.Cdate3(dr("Birthday")))
        'If IsDate(dr("Birthday")) Then Birthday.Text=FormatDateTime(dr("Birthday"), 2)

        PassPortNO.Text = "外國"
        If dr("PassPortNO") = 1 Then PassPortNO.Text = "本國"

        Sex.Text = "女"
        If dr("Sex").ToString = "M" Then Sex.Text = "男"

        Dim vMStatus As String = TIMS.cst_NODATAMsg12
        Select Case Convert.ToString(dr("MaritalStatus"))
            Case "1"
                vMStatus = "已婚"
            Case "2"
                vMStatus = "未婚"
        End Select
        MaritalStatus.Text = vMStatus

        DegreeID.Text = If(dr("DegreeName").ToString = "", TIMS.cst_NODATAMsg12, dr("DegreeName").ToString)
        GradID.Text = If(dr("GradName").ToString = "", TIMS.cst_NODATAMsg12, dr("GradName").ToString)
        School.Text = If(dr("School").ToString = "", TIMS.cst_NODATAMsg12, dr("School").ToString)
        Department.Text = If(dr("Department").ToString = "", TIMS.cst_NODATAMsg12, dr("Department").ToString)
        MilitaryID.Text = If(dr("MilitaryName").ToString = "", TIMS.cst_NODATAMsg12, dr("MilitaryName").ToString)

        If dtZip Is Nothing Then dtZip = TIMS.Get_VZipName(objconn)
        Dim sAddress As String = Convert.ToString(dr("Address"))
        'If Session(TIMS.gcst_rblWorkMode)=TIMS.cst_wmdip1 Then sAddress=TIMS.strMask(sAddress, 3)
        Dim s_ZipCODE As String = If(Convert.ToString(dr("ZipCODE6W")) <> "", Convert.ToString(dr("ZipCODE6W")), Convert.ToString(dr("ZipCODE")))
        Address.Text = TIMS.getZipName6(s_ZipCODE, sAddress, "", dtZip)

        Phone1.Text = If(dr("Phone1").ToString = "", TIMS.cst_NODATAMsg12, dr("Phone1").ToString)
        Phone2.Text = If(dr("Phone2").ToString = "", TIMS.cst_NODATAMsg12, dr("Phone2").ToString)
        Email.Text = If(dr("Email").ToString = "", TIMS.cst_NODATAMsg12, dr("Email").ToString)
        CellPhone.Text = If(dr("CellPhone").ToString = "", TIMS.cst_NODATAMsg12, dr("CellPhone").ToString)

        'TRIdentityID.Style.Item("display")="inline"
        TRIdentityID.Style.Item("display") = "none" '新的e網報名，完全不用顯示參訓身分別(暫時) ---2008-11-24 by AMU
        TRHandTypeID.Style.Item("display") = "none" '障礙類別行

        '屆退官兵身分者 'Session retreat Soldiers The identity of persons
        'HidSRSOLDIERS.Value=Convert.ToString(dr("SRSOLDIERS"))
        HidIdentityID.Value = Convert.ToString(dr("IdentityID"))
        If Convert.ToString(dr("SRSOLDIERS")) = "Y" Then
            '12:屆退官兵(須單位將級以上長官薦送函)
            Const cst_id12 As String = "12"
            If HidIdentityID.Value.IndexOf(cst_id12) = -1 Then
                If HidIdentityID.Value <> "" Then HidIdentityID.Value &= ","
                HidIdentityID.Value &= cst_id12
            End If
        End If
        If dtIdentity Is Nothing Then dtIdentity = Get_IdentityDt()
        IdentityID.Text = TIMS.Get_IdentityName(HidIdentityID.Value, dtIdentity, ",")

        '不管是不是用e網報名的資料，只要有 障礙類別 就顯示喔 ---2008-11-23 by AMU
        Dim PMSE As New Hashtable From {{"MEM_IDNO", $"{dr("IDNO")}"}}
        Dim sqlE As String = "SELECT * FROM E_MEMBER WHERE MEM_IDNO=@MEM_IDNO"
        Dim drE As DataRow = DbAccess.GetOneRow(sqlE, objconn, PMSE)
        If drE IsNot Nothing Then
            If $"{drE("HandTypeID")}" <> "" Then
                TRIdentityID.Style.Item("display") = "none"
                TRHandTypeID.Style.Item("display") = cst_inline1 '"inline"
                labHandTypeID.Text = TIMS.Get_HandTypeName($"{drE("HandTypeID")}")
                labHandLevelID.Text = TIMS.Get_HandLevelName($"{drE("HandLevelID")}")
            End If
            If $"{drE("HandTypeID2")}" <> "" Then
                TRIdentityID.Style.Item("display") = "none"
                TRHandTypeID.Style.Item("display") = cst_inline1 '"inline"
                labHandTypeID.Text = TIMS.Get_HandTypeName2($"{drE("HandTypeID2")}")
                labHandLevelID.Text = TIMS.Get_HandLevelName2($"{drE("HandLevelID2")}")
            End If
        End If

        'For i As Integer=0 To Split(dr("IdentityID"), ",").Length - 1
        '    sql="SELECT Name FROM Key_Identity WHERE IdentityID='" & Split(dr("IdentityID"), ",")(i) & "'"
        '    '20090123 andy  edit 產投 2009年 身分別「就業保險被保險人非自願失業者」 名稱改為「非自願離職者」
        '    
        '    If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '        If IdentityID.Text="" Then
        '            If Split(dr("IdentityID"), ",")(i)="02" Then
        '                IdentityID.Text="非自願離職者"
        '            Else
        '                IdentityID.Text=DbAccess.ExecuteScalar(sql, objconn)
        '            End If
        '        Else
        '            If Split(dr("IdentityID"), ",")(i)="02" Then
        '                IdentityID.Text &= "," & "非自願離職者"
        '            Else
        '                IdentityID.Text &= "," & DbAccess.ExecuteScalar(sql, objconn)
        '            End If
        '        End If
        '    Else
        '        If IdentityID.Text="" Then
        '            IdentityID.Text=DbAccess.ExecuteScalar(sql, objconn)
        '        Else
        '            IdentityID.Text &= "," & DbAccess.ExecuteScalar(sql, objconn)
        '        End If
        '    End If
        'Next

        If IsDate(dr("RelEnterDate")) Then
            RelEnterDate.Text = FormatDateTime(dr("RelEnterDate"), DateFormat.GeneralDate)

            If flag_ROC Then
                Dim relEnterDateAry As String() = FormatDateTime(dr("RelEnterDate"), DateFormat.GeneralDate).Split(" ")
                RelEnterDate.Text = TIMS.Cdate17(dr("RelEnterDate")) + " " + relEnterDateAry(1) + " " + relEnterDateAry(2)
            End If
        End If

        OCID1.Text = dr("CLASSCNAME2").ToString
        'If IsNumeric(dr("CyclType1")) Then
        '    If Int(dr("CyclType1")) <> 0 Then OCID1.Text &= "第" & Int(dr("CyclType1")) & "期"
        'End If
        If Convert.ToString(dr("LevelName")) <> "" AndAlso Int(dr("LevelName")) <> 0 Then OCID1.Text += String.Concat("(第", dr("LevelName"), "階段)")

        '檢查 內網已有資料
        'signUpStatus 0:收件完成 1:報名成功 2:報名失敗 3:正取(Key_SelResult) 4:備取 5:未錄取
        If Check_EnterType(Convert.ToString(dr("IDNO")), Convert.ToString(dr("signUpStatus")), Convert.ToString(dr("OCID1"))) Then
            OCID1.Text &= "(內網已有資料)"
            OCID1.ForeColor = Color.Red
        End If

        signUpMemo.Text = dr("signUpMemo").ToString
        If Convert.ToString(dr("LastModifyDate")) <> "" Then lab_LastModifyDate.Text = String.Concat("最後異動日期：", dr("LastModifyDate"))
        'signUpStatus 0:收件完成 1:報名成功 2:報名失敗 3:正取(Key_SelResult) 4:備取 5:未錄取
        Dim flag_signUpStatus_isNot_0 As Boolean = False
        flag_signUpStatus_isNot_0 = If(dr("signUpStatus") <> 0, True, False)
        If flag_signUpStatus_isNot_0 Then
            Button1.Visible = False
            Button2.Visible = False
        Else
            If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 AndAlso Convert.ToString(dr("IsStdBlack")) = "Y" Then
                signUpMemo.Text = "此學員己遭處分,系統帶審核失敗"
                signUpMemo.ReadOnly = True
                Button1.Enabled = False
                Button2.Enabled = True
            End If
        End If

        '2006/06/21 by Vicient 產學訓 start
        Label1.Text = OCID1.Text '報名班級
        Label2.Text = RelEnterDate.Text
        Label3.Text = Name.Text
        Label4.Text = Birthday.Text
        'Hid_Birthday.Value=Birthday.Text
        ViewState("Birthday") = Birthday.Text
        'If Session(TIMS.gcst_rblWorkMode)=TIMS.cst_wmdip1 Then
        '    Birthday.Text=TIMS.strMask(Birthday.Text, 2)
        '    Label4.Text=TIMS.strMask(Label4.Text, 2)
        'End If
        Label5.Text = PassPortNO.Text
        'If IDNO.Text <> "" Then IDNO.Text=Trim(IDNO.Text)
        'If IDNO.Text <> "" Then IDNO.Text=UCase(IDNO.Text)
        'If IDNO.Text <> "" Then IDNO.Text=TIMS.ChangeIDNO(IDNO.Text)
        Label6.Text = IDNO.Text 'TIMS.ChangeIDNO(IDNO.Text)
        ViewState("IDNO") = IDNO.Text
        'If Session(TIMS.gcst_rblWorkMode)=TIMS.cst_wmdip1 Then
        '    IDNO.Text=TIMS.strMask(IDNO.Text, 1)
        '    Label6.Text=TIMS.strMask(Label6.Text, 1)
        'End If
        Label7.Text = Sex.Text
        'Label8.Text=MaritalStatus.Text
        Label9.Text = DegreeID.Text
        'Label10.Text=GradID.Text
        'Label11.Text=School.Text
        'Label12.Text=Department.Text
        'Label13.Text=MilitaryID.Text
        Label14.Text = Address.Text

        Label16.Text = Phone1.Text
        Label17.Text = Phone2.Text
        Label18.Text = Email.Text
        Label19.Text = CellPhone.Text
        'Label21.Text=IdentityID.Text
    End Sub

    'list STUD_ENTERTRAIN2.eSerNum
    Sub SHOW_ENTERTRAIN2()
        If HiderrFlag.Value = cst_errFlag Then
            'Dim sErrMsg As String="資料有誤請重新查詢!" & vbCrLf
            'Common.MessageBox(Me, sErrMsg)
            Exit Sub
        End If

        heSerNum.Value = TIMS.ClearSQM(Request("eSerNum"))
        If heSerNum.Value = "" Then Exit Sub
        Dim sql As String = ""
        sql = " Select * FROM STUD_ENTERTRAIN2 WHERE ESERNUM=@ESERNUM "
        Dim dtSE2 As DataTable = Nothing
        Dim parms As Hashtable = New Hashtable()
        parms.Clear()
        parms.Add("ESERNUM", heSerNum.Value)
        dtSE2 = DbAccess.GetDataTable(sql, objconn, parms)
        If dtSE2.Rows.Count = 0 Then
            Label15.Text = TIMS.cst_NODATAMsg12
            Label20.Text = TIMS.cst_NODATAMsg12

            'Label22.Text=TIMS.cst_NODATAMsg12
            'Label23.Text=TIMS.cst_NODATAMsg12
            'Label24.Text=TIMS.cst_NODATAMsg12
            'Label25.Text=TIMS.cst_NODATAMsg12
            'Label26.Text=TIMS.cst_NODATAMsg12
            'Label27.Text=TIMS.cst_NODATAMsg12
            'Label28.Text=TIMS.cst_NODATAMsg12
            'Label29.Text=TIMS.cst_NODATAMsg12
            'Label31.Text=TIMS.cst_NODATAMsg12
            'Label32.Text=TIMS.cst_NODATAMsg12
            'Label30.Text = TIMS.cst_NODATAMsg12
            'Label33.Text = TIMS.cst_NODATAMsg12
            'Table22.Style("display") = "none"
            'Table23.Style("display") = "none"
            'Label39.Text=TIMS.cst_NODATAMsg12
            'Label42.Text=TIMS.cst_NODATAMsg12
            'Label43.Text=TIMS.cst_NODATAMsg12
            'Label44.Text=TIMS.cst_NODATAMsg12
            'Label45.Text=TIMS.cst_NODATAMsg12
            'Label47.Text=TIMS.cst_NODATAMsg12
            'Label48.Text=TIMS.cst_NODATAMsg12
            'Label49.Text=TIMS.cst_NODATAMsg12
            'Label50.Text=TIMS.cst_NODATAMsg12
            Label40.Text = TIMS.cst_NODATAMsg12
            Label41.Text = TIMS.cst_NODATAMsg12
            Label46.Text = TIMS.cst_NODATAMsg12
            Label51.Text = TIMS.cst_NODATAMsg12
            Label52.Text = TIMS.cst_NODATAMsg12
            Label53.Text = TIMS.cst_NODATAMsg12
            Label54.Text = TIMS.cst_NODATAMsg12
            Label55.Text = TIMS.cst_NODATAMsg12
            Label56.Text = TIMS.cst_NODATAMsg12
            Label57.Text = TIMS.cst_NODATAMsg12
            Label58.Text = TIMS.cst_NODATAMsg12
            Label59.Text = TIMS.cst_NODATAMsg12
            ACTNO60.Text = TIMS.cst_NODATAMsg12 'ActNo 
            'Label61.Text = TIMS.cst_NODATAMsg12
            'Label62.Text = TIMS.cst_NODATAMsg12
            Label63.Text = TIMS.cst_NODATAMsg12
            Label64.Text = TIMS.cst_NODATAMsg12
            Label65.Text = TIMS.cst_NODATAMsg12
            'Table13
            '戶籍地址
            LabHouseholdAddress.Text = TIMS.cst_NODATAMsg12
            '主要參訓身分別 Lab1311 Label20
            Lab1311.Text = TIMS.cst_NODATAMsg12
            '投保單位名稱  Lab1301 - Label59
            Lab1301.Text = TIMS.cst_NODATAMsg12 'If(Convert.ToString(dr("Actname")) <> "", Convert.ToString(dr("Actname")), TIMS.cst_NODATAMsg12)
            '投保單位保險證號    Lab1302 ACTNO60-
            Lab1302.Text = TIMS.cst_NODATAMsg12 'If(Convert.ToString(dr("ActNo")) <> "", Convert.ToString(dr("ActNo")), TIMS.cst_NODATAMsg12)
            '投保單位類別  Lab1303	 Label63
            Lab1303.Text = TIMS.cst_NODATAMsg12 's_ActType
            '投保單位電話  Lab1304 Label64
            Lab1304.Text = TIMS.cst_NODATAMsg12 'If(Convert.ToString(dr("ActTel")) <> "", Convert.ToString(dr("ActTel")), TIMS.cst_NODATAMsg12)
            '投保單位地址  Lab1305 Label65
            Lab1305.Text = TIMS.cst_NODATAMsg12 'If(s_ActAddress <> "", s_ActAddress, TIMS.cst_NODATAMsg12)
            '目前公司名稱  Lab1307	 Label40
            Lab1307.Text = TIMS.cst_NODATAMsg12 'If(Convert.ToString(dr("Uname")) <> "", Convert.ToString(dr("Uname")), TIMS.cst_NODATAMsg12)
            '統一編號    Lab1308 Label41
            Lab1308.Text = TIMS.cst_NODATAMsg12 'If(Convert.ToString(dr("Intaxno")) <> "", Convert.ToString(dr("Intaxno")), TIMS.cst_NODATAMsg12)
            '目前任職部門   Lab1309	 Label45
            Lab1309.Text = TIMS.cst_NODATAMsg12 'If(Convert.ToString(dr("ServDept")) <> "", Convert.ToString(dr("ServDept")), TIMS.cst_NODATAMsg12)
            '職稱  Lab1310 Label46
            Lab1310.Text = TIMS.cst_NODATAMsg12 'If(Convert.ToString(dr("JobTitle")) <> "", Convert.ToString(dr("JobTitle")), TIMS.cst_NODATAMsg12)
            'ActComidno.Text=TIMS.cst_NODATAMsg12
            Exit Sub
        End If
        'If dtSE2.Rows.Count=0 Then Exit Sub
        Dim dr As DataRow = dtSE2.Rows(0)

        'NOT dr Is Nothing 
        Dim sZIPCODE As String = Convert.ToString(dr("ZIPCODE2"))
        'gsEnterDate=TIMS.cdate3(dr("ENTERDATE"))
        If Hid_MSG1.Value = "" Then Hid_MSG1.Value = TIMS.SHOW_ZIP2MSG(Me, sZIPCODE, gsEnterDate, objconn)

        'https://jira.turbotech.com.tw/browse/TIMSC-150
        Dim iADID As Integer = 0
        Dim flagMSG2 As Boolean = TIMS.CHK_DIS2MSG(Me, sZIPCODE, gsEnterDate, objconn, iADID)
        If flagMSG2 AndAlso Hid_MSG2.Value = "" Then
            Hid_MSGADIDN.Value = iADID
            Hid_MSG2.Value = TIMS.SHOW_DIS2MSG(Me, sZIPCODE, gsEnterDate, iADID, objconn)
        End If

        Label15.Text = TIMS.cst_NODATAMsg12

        If dtZip Is Nothing Then dtZip = TIMS.Get_VZipName(objconn)
        'Dim sZ1 As String=Convert.ToString(dr("ZipCode2"))
        Dim sAddress As String = Convert.ToString(dr("HouseholdAddress"))
        'If Session(TIMS.gcst_rblWorkMode)=TIMS.cst_wmdip1 Then sAddress=TIMS.strMask(sAddress, 3)
        Dim s_ZipCode2 As String = If(Convert.ToString(dr("ZipCode2_6W")) <> "", Convert.ToString(dr("ZipCode2_6W")), Convert.ToString(dr("ZipCode2")))
        Dim s_HouseholdAddress As String = TIMS.getZipName6(s_ZipCode2, sAddress, "", dtZip)
        Label15.Text = s_HouseholdAddress

        If dtIdentity Is Nothing Then dtIdentity = Get_IdentityDt()
        Label20.Text = TIMS.cst_NODATAMsg12
        Dim s_MIdentityID As String = TIMS.GetMyValue(dtIdentity, "IdentityID", "Name", Convert.ToString(dr("MIdentityID")))
        If s_MIdentityID <> "" Then Label20.Text = s_MIdentityID

        'If Not IsDBNull(dr("HandTypeID")) Then
        '    str="Select * from Key_HandicatType where HandTypeID='" & dr("HandTypeID") & "'"
        '    table1=DbAccess.GetDataTable(str)
        '    If table1.Rows.Count <> 0 Then
        '        dr1=table1.Rows(0)
        '        Label22.Text=dr1("Name")
        '    End If
        'Else
        '    Label22.Text=TIMS.cst_NODATAMsg12
        'End If

        'If Not IsDBNull(dr("HandLevelID")) Then
        '    str="select * from Key_HandicatLevel where HandLevelID='" & dr("HandLevelID") & "'"
        '    table1=DbAccess.GetDataTable(str)
        '    If table1.Rows.Count <> 0 Then
        '        dr1=table1.Rows(0)
        '        Label23.Text=dr1("Name")
        '    End If
        'Else
        '    Label23.Text=TIMS.cst_NODATAMsg12
        'End If

        'If Not IsDBNull(dr("PriorWorkOrg1")) Then
        '    Label24.Text=dr("PriorWorkOrg1")
        'Else
        '    Label24.Text=TIMS.cst_NODATAMsg12
        'End If

        'If Not IsDBNull(dr("Title1")) Then
        '    Label26.Text=dr("Title1")
        'Else
        '    Label26.Text=TIMS.cst_NODATAMsg12
        'End If

        'If Not IsDBNull(dr("PriorWorkOrg2")) Then
        '    Label25.Text=dr("PriorWorkOrg2")
        'Else
        '    Label25.Text=TIMS.cst_NODATAMsg12
        'End If

        'If Not IsDBNull(dr("Title2")) Then
        '    Label27.Text=dr("Title2")
        'Else
        '    Label27.Text=TIMS.cst_NODATAMsg12
        'End If

        'If Not IsDBNull(dr("SOfficeYM1")) Then
        '    Label28.Text=dr("SOfficeYM1")
        '    Label28.Text=Label28.Text & " ~ "
        'End If
        'If Not IsDBNull(dr("FOfficeYM1")) Then
        '    If Not IsDBNull(dr("SOfficeYM1")) Then
        '        Label28.Text=Label28.Text & dr("FOfficeYM1")
        '    Else
        '        Label28.Text=" ~ " & dr("FOfficeYM1")
        '    End If
        'End If
        'If Label28.Text="" Then
        '    Label28.Text=TIMS.cst_NODATAMsg12
        'End If

        'If Not IsDBNull(dr("SOfficeYM2")) Then
        '    Label29.Text=dr("SOfficeYM2")
        '    Label29.Text=Label29.Text & " ~ "
        'End If
        'If Not IsDBNull(dr("FOfficeYM2")) Then
        '    If Not IsDBNull(dr("SOfficeYM2")) Then
        '        Label29.Text=Label29.Text & dr("FOfficeYM2")
        '    Else
        '        Label29.Text=" ~ " & dr("FOfficeYM2")
        '    End If
        'End If
        'If Label29.Text="" Then
        '    Label29.Text=TIMS.cst_NODATAMsg12
        'End If

        'Label30.Text = TIMS.cst_NODATAMsg12
        'If Not IsDBNull(dr("PriorWorkPay")) Then
        '    Label30.Text = dr("PriorWorkPay")
        'End If

        'If Not IsDBNull(dr("RealJobless")) Then
        '    Label31.Text=dr("RealJobless")
        'Else
        '    Label31.Text=TIMS.cst_NODATAMsg12
        'End If

        'If Not IsDBNull(dr("Traffic")) Then
        '    If dr("Traffic")=1 Then
        '        Label32.Text="住宿"
        '    ElseIf dr("Traffic")=2 Then
        '        Label32.Text="通勤"
        '    Else
        '        Label32.Text=TIMS.cst_NODATAMsg12
        '    End If
        'Else
        '    Label32.Text=TIMS.cst_NODATAMsg12
        'End If

        '**by Milor 20080512 加入項目2訓練單位代轉現金 start
        'Label33.Text = TIMS.cst_NODATAMsg12
        'Table22.Style("display") = "none"
        'Table23.Style("display") = "none"
        'Label37.Text = "" 'Convert.ToString(dr("BankName"))
        'Label38.Text = "" 'Convert.ToString(dr("AcctHeadNo"))
        'Label61.Text = "" 'Convert.ToString(dr("ExBankName"))
        'Label62.Text = "" 'Convert.ToString(dr("AcctExNo"))
        'Label34.Text = "" 'Convert.ToString(dr("AcctNo"))
        'Label35.Text = "" 'Convert.ToString(dr("PostNo"))
        'Label36.Text = "" 'Convert.ToString(dr("AcctNo"))
        'Select Case Convert.ToString(dr("AcctMode"))
        '    Case "1"
        '        Label33.Text = "銀行"
        '        Table22.Style("display") = "none"
        '        Table23.Style("display") = cst_inline1 '"inline"
        '        Label37.Text = Convert.ToString(dr("BankName"))
        '        Label38.Text = Convert.ToString(dr("AcctHeadNo"))
        '        Label61.Text = Convert.ToString(dr("ExBankName"))
        '        Label62.Text = Convert.ToString(dr("AcctExNo"))
        '        Label34.Text = Convert.ToString(dr("AcctNo"))
        '    Case "0"
        '        Label33.Text = "郵局"
        '        Table22.Style("display") = cst_inline1 '"inline"
        '        Table23.Style("display") = "none"
        '        Label35.Text = Convert.ToString(dr("PostNo"))
        '        Label36.Text = Convert.ToString(dr("AcctNo"))
        '    Case "2"
        '        Label33.Text = "訓練單位代轉現金"
        '        Table22.Style("display") = "none"
        '        Table23.Style("display") = "none"
        'End Select

        '**by Milor 20080512 end
        'If Not IsDBNull(dr("FirDate")) Then
        '    Label39.Text=dr("FirDate")
        'Else
        '    Label39.Text=TIMS.cst_NODATAMsg12
        'End If
        Label40.Text = TIMS.cst_NODATAMsg12
        If Not IsDBNull(dr("Uname")) Then Label40.Text = dr("Uname")

        Label41.Text = TIMS.cst_NODATAMsg12
        If Not IsDBNull(dr("Intaxno")) Then Label41.Text = dr("Intaxno")

        'If Not IsDBNull(dr("Tel")) Then
        '    Label42.Text=dr("Tel")
        'Else
        '    Label42.Text=TIMS.cst_NODATAMsg12
        'End If

        'If Not IsDBNull(dr("Fax")) Then
        '    Label43.Text=dr("Fax")
        'Else
        '    Label43.Text=TIMS.cst_NODATAMsg12
        'End If

        'If dr("Zip").ToString <> "" And Trim(dr("Zip").ToString) <> "-1" Then
        '    Label44.Text="[" & dr("Zip") & "]" & TIMS.Get_ZipName(dr("Zip"))
        '    If dr("Addr").ToString <> "" Then
        '        Label44.Text=Label44.Text & dr("Addr")
        '    End If
        'Else
        '    Label44.Text=TIMS.cst_NODATAMsg12
        'End If

        Label45.Text = TIMS.cst_NODATAMsg12
        If Not IsDBNull(dr("ServDept")) Then Label45.Text = dr("ServDept")

        Label46.Text = TIMS.cst_NODATAMsg12
        If Not IsDBNull(dr("JobTitle")) Then Label46.Text = dr("JobTitle")

        '(1: 是 ,0:否)
        Label50.Text = TIMS.cst_NODATAMsg12
        If Convert.ToString(dr("Q1")) <> "" Then
            Label50.Text = "否"
            If Convert.ToString(dr("Q1")) = "1" Then Label50.Text = "是"
        End If
        Dim tmpQ2 As String = ""
        Dim z As Integer = 0
        If Not IsDBNull(dr("Q2_1")) Then
            z += 1
            tmpQ2 &= z & ". 為補充與原專長相關之技能 "
        End If
        If Not IsDBNull(dr("Q2_2")) Then
            z += 1
            tmpQ2 &= z & ". 轉換其他行職業所需技能 "
        End If
        If Not IsDBNull(dr("Q2_3")) Then
            z += 1
            tmpQ2 &= z & " .拓展工作領域及視野 "
        End If
        If Not IsDBNull(dr("Q2_4")) Then
            z += 1
            tmpQ2 &= z & " .其他"
        End If
        Label51.Text = TIMS.cst_NODATAMsg12
        If tmpQ2 <> "" Then
            Label51.Text = tmpQ2
        End If
        Label52.Text = TIMS.cst_NODATAMsg12
        If Not IsDBNull(dr("Q3")) Then
            Select Case Convert.ToString(dr("Q3"))
                Case "1"
                    Label52.Text = "轉換工作"
                Case "2"
                    Label52.Text = "留任"
                Case "3"
                    Label52.Text = "其他"
                    If Not IsDBNull(dr("Q3_Other")) Then Label52.Text &= "(" & dr("Q3_Other") & ")"
            End Select
        End If
        Label53.Text = If(Not IsDBNull(dr("Q5")), If(Convert.ToString(dr("Q5")) = "1", "是", "否"), TIMS.cst_NODATAMsg12)

        Label54.Text = TIMS.cst_NODATAMsg12
        If Not IsDBNull(dr("Q61")) Then Label54.Text = dr("Q61")
        Label55.Text = TIMS.cst_NODATAMsg12
        If Not IsDBNull(dr("Q62")) Then Label55.Text = dr("Q62")
        Label56.Text = TIMS.cst_NODATAMsg12
        If Not IsDBNull(dr("Q63")) Then Label56.Text = dr("Q63")
        Label57.Text = TIMS.cst_NODATAMsg12
        If Not IsDBNull(dr("Q64")) Then Label57.Text = dr("Q64")

        If dtTrade Is Nothing Then dtTrade = TIMS.Get_TradeDt(objconn)
        Label58.Text = TIMS.cst_NODATAMsg12
        If Not IsDBNull(dr("Q4")) Then Label58.Text = TIMS.GetMyValue(dtTrade, "TradeID", "TradeName", Convert.ToString(dr("Q4")))
        ACTNO60.Text = TIMS.cst_NODATAMsg12
        If Not IsDBNull(dr("ActNo")) Then ACTNO60.Text = dr("ActNo")
        Label59.Text = TIMS.cst_NODATAMsg12
        If Not IsDBNull(dr("Actname")) Then Label59.Text = dr("Actname")

        Dim s_ActType As String = TIMS.cst_NODATAMsg12
        If Not IsDBNull(dr("ActType")) Then
            Select Case Convert.ToString(dr("ActType"))
                Case "1"
                    s_ActType = "勞"
                Case "2"
                    s_ActType = "農"
            End Select
        End If
        Label63.Text = s_ActType

        'If Not IsDBNull(dr("ActComidno")) Then
        '    ActComidno.Text=dr("ActComidno")
        'Else
        '    ActComidno.Text=TIMS.cst_NODATAMsg12
        'End If

        Label64.Text = TIMS.cst_NODATAMsg12
        If Not IsDBNull(dr("ActTel")) Then Label64.Text = dr("ActTel")

        Label65.Text = TIMS.cst_NODATAMsg12
        If dtZip Is Nothing Then dtZip = TIMS.Get_VZipName(objconn)
        'sZ1=Convert.ToString(dr("ZipCode3"))
        sAddress = Convert.ToString(dr("ActAddress"))
        'If Session(TIMS.gcst_rblWorkMode)=TIMS.cst_wmdip1 Then sAddress=TIMS.strMask(sAddress, 3)
        Dim sZipCode3 As String = If(Convert.ToString(dr("ZipCode3_6W")) <> "", Convert.ToString(dr("ZipCode3_6W")), Convert.ToString(dr("ZipCode3")))
        Dim s_ActAddress As String = TIMS.getZipName6(sZipCode3, sAddress, "", dtZip)
        Label65.Text = s_ActAddress 'TIMS.getZipName3(sZ1, sZ2, sAddress, "", dtZip)

        '戶籍地址
        LabHouseholdAddress.Text = If(s_HouseholdAddress <> "", s_HouseholdAddress, TIMS.cst_NODATAMsg12)
        '主要參訓身分別 Lab1311 Label20
        Lab1311.Text = If(s_MIdentityID <> "", s_MIdentityID, TIMS.cst_NODATAMsg12)
        '投保單位名稱  Lab1301 - Label59
        Lab1301.Text = If(Convert.ToString(dr("Actname")) <> "", Convert.ToString(dr("Actname")), TIMS.cst_NODATAMsg12)
        '投保單位保險證號    Lab1302 ACTNO60-
        Lab1302.Text = If(Convert.ToString(dr("ActNo")) <> "", Convert.ToString(dr("ActNo")), TIMS.cst_NODATAMsg12)
        '投保單位類別  Lab1303	 Label63
        Lab1303.Text = s_ActType
        '投保單位電話  Lab1304 Label64
        Lab1304.Text = If(Convert.ToString(dr("ActTel")) <> "", Convert.ToString(dr("ActTel")), TIMS.cst_NODATAMsg12)
        '投保單位地址  Lab1305 Label65
        Lab1305.Text = If(s_ActAddress <> "", s_ActAddress, TIMS.cst_NODATAMsg12)
        '目前公司名稱  Lab1307	 Label40
        Lab1307.Text = If(Convert.ToString(dr("Uname")) <> "", Convert.ToString(dr("Uname")), TIMS.cst_NODATAMsg12)
        '統一編號    Lab1308 Label41
        Lab1308.Text = If(Convert.ToString(dr("Intaxno")) <> "", Convert.ToString(dr("Intaxno")), TIMS.cst_NODATAMsg12)
        '目前任職部門   Lab1309	 Label45
        Lab1309.Text = If(Convert.ToString(dr("ServDept")) <> "", Convert.ToString(dr("ServDept")), TIMS.cst_NODATAMsg12)
        '職稱  Lab1310 Label46
        Lab1310.Text = If(Convert.ToString(dr("JobTitle")) <> "", Convert.ToString(dr("JobTitle")), TIMS.cst_NODATAMsg12)
    End Sub

    Sub SHOW_BIEPTBL_DG1()
        If HiderrFlag.Value = cst_errFlag Then
            'Dim sErrMsg As String="資料有誤請重新查詢!" & vbCrLf
            'Common.MessageBox(Me, sErrMsg)
            Exit Sub
        End If
        Call DataGridTable_Show(IDNOValue.Value, STDateValue.Value, BFDateValue.Value)
        If IDNOValue.Value <> "" Then Call DataGrid2_Show(IDNOValue.Value)
        Call DataGrid4_Show(vsExamDate, IDNOValue.Value, vsOCID1val) '已報名同一甄試日期 
    End Sub

    '已申請失業給付資料 'BIEPTBL '領取失業補助檔(產學訓)  'Stud_SubsidyCost-學員輔助金撥款檔(產學訓)
    Sub DataGridTable_Show(ByVal vIDNO As String, ByVal vSTDate1 As String, ByVal vBFDate1 As String)
        'Dim sql As String=""
        DataGridTable.Visible = False
        Return
        'If vIDNO <> "" AndAlso vSTDate1 <> "" AndAlso vBFDate1 <> "" Then
        '    Dim sql As String=""
        '    sql=""
        '    sql &= " SELECT b.APPLY_DATE ,b.APPLY_MONEY ,aw.STATION_NAME" & vbCrLf
        '    sql &= " FROM (" & vbCrLf
        '    sql &= " SELECT APPLY_DATE ,APPLY_MONEY ,STATION" & vbCrLf
        '    sql &= " FROM dbo.BIEPTBL" & vbCrLf
        '    sql &= " WHERE idno=@idno" & vbCrLf
        '    sql &= " AND APPLY_DATE >= convert(date,@BFDate1)" & vbCrLf
        '    sql &= " AND APPLY_DATE <= convert(date,@STDate1)" & vbCrLf
        '    sql &= " ) b" & vbCrLf
        '    sql &= " LEFT JOIN dbo.ADP_WORKSTATION aw ON aw.Station_Scheme_ID + aw.Station_Unit_ID + aw.Station_ID=b.STATION" & vbCrLf
        '    Dim parms As Hashtable=New Hashtable()
        '    parms.Clear()
        '    parms.Add("idno", vIDNO)
        '    parms.Add("BFDate1", vBFDate1)
        '    parms.Add("STDate1", vSTDate1)
        '    Dim dt As DataTable=Nothing
        '    dt=DbAccess.GetDataTable(sql, objconn, parms)
        '    If dt.Rows.Count > 0 Then
        '        DataGridTable.Visible=True
        '        DataGrid1.DataSource=dt
        '        DataGrid1.DataBind()
        '    End If
        'End If
    End Sub

    '已申請職訓生活津貼 '就業安定基金特定對象生活津貼 (sub_subsidyapply)
    Private Sub DataGrid2_Show(ByVal IDNO As String)
        DataGridTable2.Visible = False
        Return
        'If IDNO="" Then Return

        'Dim sql As String=""
        'Dim dt As DataTable
        ''Dim dr As DataRow
        'sql=""
        'sql &= " SELECT a.IDNO ,e.OrgName" & vbCrLf
        'sql &= " ,dbo.FN_GET_CLASSCNAME(d.CLASSCNAME,d.CYCLTYPE) CLASSCNAME" & vbCrLf
        'sql &= " ,d.STDate ,d.FTDate ,a.TrainingMoney ,a.PayMoney" & vbCrLf
        'sql &= " FROM SUB_SUBSIDYAPPLY a" & vbCrLf
        'sql &= " JOIN Class_StudentsOfClass b ON a.SOCID=b.SOCID" & vbCrLf
        'sql &= " JOIN Stud_StudentInfo c ON b.SID=c.SID AND c.IDNO=@IDNO" & vbCrLf
        'sql &= " JOIN Class_ClassInfo d ON b.OCID=d.OCID" & vbCrLf
        'sql &= " JOIN VIEW_RIDNAME e ON d.RID=e.RID" & vbCrLf
        'Dim parms As Hashtable=New Hashtable()
        'parms.Clear()
        'parms.Add("IDNO", IDNO)
        'dt=DbAccess.GetDataTable(sql, objconn, parms)
        ''2007年前，補助金為2萬之後為3萬" '2008產業人才投資方案，改為3年5萬
        ''DataGridTable2.Visible=False
        'If dt.Rows.Count > 0 Then
        '    DataGridTable2.Visible=True

        '    DataGrid2.DataSource=dt
        '    DataGrid2.DataBind()
        'End If
    End Sub

    '已報名同一甄試日期 已報名同一時間甄試
    Sub DataGrid4_Show(ByVal ExamDate As String, ByVal IDNO As String, ByVal OCID1 As String)
        'Dim sql As String=""
        'signUpStatus 0:收件完成 1:報名成功 2:報名失敗 3:正取(Key_SelResult) 4:備取 5:未錄取
        Dim flagNG As Boolean = True '有異常
        ExamDate = TIMS.Cdate3(ExamDate)
        If Convert.ToString(ExamDate) <> "" AndAlso IDNO <> "" AndAlso OCID1 <> "" Then flagNG = False '沒有異常。
        If flagNG Then
            '異常
            DataGridTable4.Visible = False
            Exit Sub
        End If

        Dim parms As Hashtable = New Hashtable()
        parms.Clear()
        parms.Add("IDNO", IDNO)
        parms.Add("OCID1", OCID1)
        parms.Add("ExamDate", ExamDate)

        'signUpStatus 0:收件完成 1:報名成功 2:報名失敗 3:正取(Key_SelResult) 4:備取 5:未錄取
        Dim sql As String = ""
        sql = ""
        sql &= " WITH WE2 AS (" & vbCrLf
        sql &= "  SELECT DISTINCT b.OCID1 ,a.IDNO ,(CASE WHEN b.signUpStatus=2 THEN '審核失敗' WHEN b.signUpStatus=5 THEN '審核失敗' END) signUpStatus2" & vbCrLf
        sql &= "  FROM Stud_EnterTemp2 a WITH(NOLOCK)" & vbCrLf
        sql &= "  JOIN Stud_EnterType2 b WITH(NOLOCK) ON a.eSETID=b.eSETID" & vbCrLf
        sql &= "  WHERE a.IDNO=@IDNO AND b.OCID1 != @OCID1" & vbCrLf
        sql &= " )" & vbCrLf
        sql &= " ,WE1 AS (" & vbCrLf
        sql &= "  SELECT DISTINCT b.OCID1 ,a.IDNO ,null signUpStatus2" & vbCrLf
        sql &= "  FROM Stud_EnterTemp a WITH(NOLOCK)" & vbCrLf
        sql &= "  JOIN Stud_EnterType b WITH(NOLOCK) ON a.SETID=b.SETID" & vbCrLf
        sql &= "  WHERE a.IDNO=@IDNO AND b.OCID1 != @OCID1" & vbCrLf
        sql &= " )" & vbCrLf
        sql &= " ,WE3 AS (" & vbCrLf
        sql &= "  select g.OCID1,g.IDNO" & vbCrLf
        sql &= "  ,MAX(g.signUpStatus2) signUpStatus2" & vbCrLf
        sql &= "  from (SELECT * FROM WE1 UNION SELECT * FROM WE2) g" & vbCrLf
        sql &= "  GROUP BY g.OCID1,g.IDNO" & vbCrLf
        sql &= " )" & vbCrLf

        sql &= " SELECT ip.Years + ip.DistName + ip.PlanName + ip.Seq PlanName" & vbCrLf
        sql &= " ,oo.OrgName" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE) CLASSCNAME" & vbCrLf
        sql &= " ,CONVERT(VARCHAR, cc.examdate, 111) examdate" & vbCrLf
        sql &= " ,CONVERT(VARCHAR, cc.stdate, 111) stdate" & vbCrLf
        sql &= " ,CONVERT(VARCHAR, cc.ftdate, 111) ftdate" & vbCrLf
        sql &= " ,b.signUpStatus2" & vbCrLf
        sql &= " ,b.idno" & vbCrLf
        sql &= " FROM CLASS_CLASSINFO cc WITH(NOLOCK)" & vbCrLf
        sql &= " JOIN VIEW_PLAN ip ON ip.planid=cc.planid AND cc.ExamDate IS NOT NULL" & vbCrLf
        sql &= " JOIN ORG_ORGINFO oo WITH(NOLOCK) ON oo.comidno=cc.comidno" & vbCrLf
        sql &= " JOIN WE3 b ON b.OCID1=cc.OCID" & vbCrLf
        sql &= " WHERE convert(date,cc.ExamDate)=convert(date,@ExamDate)" & vbCrLf
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '產投或充飛，只判斷該計畫
            sql &= " AND ip.TPlanID=@TPlanID" & vbCrLf
            parms.Add("TPlanID", sm.UserInfo.TPlanID)
        Else
            'TIMS計畫 排除產投 充飛
            sql &= " AND ip.TPlanID NOT IN ('28','54')" & vbCrLf
        End If
        sql &= " AND cc.OCID != @OCID1" & vbCrLf

        'Dim sCmd As New SqlCommand(sql, objconn)
        Call TIMS.OpenDbConn(objconn)
        Dim dt2 As DataTable = DbAccess.GetDataTable(sql, objconn, parms)

        DataGridTable4.Visible = False
        If dt2.Rows.Count > 0 Then
            DataGridTable4.Visible = True
            DataGrid4.DataSource = dt2
            DataGrid4.DataBind()
        End If
    End Sub

#Region "(No Use)"

    'Private Sub Update_StudEnterTemp(ByVal IDNO2 As String, _
    '    ByVal tmpConn As SqlConnection, ByVal tmpTrans As SqlTransaction)

    '    'Dim sqlAdp As New SqlDataAdapter
    '    Dim sqlStr As String
    '    Dim BKID As Integer=0

    '    Try
    '        IDNO2=TIMS.ChangeIDNO(IDNO2)
    '        Dim tmpSETID As Integer=0
    '        sqlStr="" & vbCrLf
    '        sqlStr &= " select" & vbCrLf
    '        sqlStr &= "  SETID,Name" & vbCrLf
    '        sqlStr &= " ,Sex" & vbCrLf
    '        sqlStr &= " ,Birthday,PassPortNO" & vbCrLf
    '        sqlStr &= " ,MaritalStatus,DegreeID" & vbCrLf
    '        sqlStr &= " ,GradID,School" & vbCrLf
    '        sqlStr &= " ,Department,MilitaryID" & vbCrLf
    '        sqlStr &= " ,ZipCode,Address" & vbCrLf
    '        sqlStr &= " ,Phone1,Phone2" & vbCrLf
    '        sqlStr &= " ,CellPhone,Email" & vbCrLf
    '        sqlStr &= " ,IsAgree,ZipCODE6W" & vbCrLf
    '        sqlStr &= " ,ModifyAcct,ModifyDate" & vbCrLf
    '        sqlStr &= " from Stud_EnterTemp2" & vbCrLf
    '        sqlStr &= " where IDNO=@IDNO2" & vbCrLf '有IDNO
    '        sqlStr &= " and SETID IS NOT NULL" & vbCrLf 'SETID有值
    '        sqlStr &= " order by ModifyDate desc" & vbCrLf '最新1筆料。
    '        Dim dt As New DataTable
    '        Dim cmd As New SqlCommand(sqlStr, tmpConn, tmpTrans)
    '        With cmd
    '            .Parameters.Clear()
    '            .Parameters.Add("@IDNO2", SqlDbType.VarChar).Value=IDNO2
    '            dt.Load(.ExecuteReader())
    '        End With
    '        If dt.Rows.Count > 0 Then
    '            Dim dr As DataRow=dt.Rows(0)
    '            tmpSETID=dr("SETID")
    '        End If

    '        If tmpSETID > 0 Then
    '            '將Stud_EnterTemp2的資料更新到Stud_EnterTemp
    '            sqlStr="  UPDATE Stud_EnterTemp SET (Name,Sex" & vbCrLf
    '            sqlStr &= " ,Birthday,PassPortNO "
    '            sqlStr &= " ,MaritalStatus,DegreeID "
    '            sqlStr &= " ,GradID,School" & vbCrLf
    '            sqlStr &= " ,Department,MilitaryID "
    '            sqlStr &= " ,ZipCode,Address "
    '            sqlStr &= " ,Phone1,Phone2 "
    '            sqlStr &= " ,CellPhone,Email" & vbCrLf
    '            sqlStr &= " ,IsAgree,ZipCODE6W "
    '            sqlStr &= " ,ModifyAcct,ModifyDate "
    '            sqlStr &= " )=(select b.Name ,b.Sex "
    '            sqlStr &= " ,b.Birthday,b.PassPortNO "
    '            sqlStr &= " ,b.MaritalStatus,b.DegreeID "
    '            sqlStr &= " ,b.GradID,b.School" & vbCrLf
    '            sqlStr &= " ,b.Department,b.MilitaryID "
    '            sqlStr &= " ,b.ZipCode,b.Address "
    '            sqlStr &= " ,b.Phone1,b.Phone2 "
    '            sqlStr &= " ,b.CellPhone,b.Email" & vbCrLf
    '            sqlStr &= " ,b.IsAgree,b.ZipCODE6W "
    '            sqlStr &= " ,@ModifyAcct ModifyAcct ,getdate() ModifyDate" & vbCrLf
    '            sqlStr &= " from Stud_EnterTemp2 b" & vbCrLf
    '            sqlStr &= " where b.IDNO=@IDNO2 and b.SETID=@SETID) " '同IDNO同SETID 執行更新
    '            sqlStr &= " where IDNO=@IDNO1 " '同IDNO
    '            Dim sqlAdp As New SqlDataAdapter
    '            With sqlAdp
    '                .UpdateCommand=New SqlCommand(sqlStr, tmpConn, tmpTrans)
    '                .UpdateCommand.Parameters.Clear()
    '                .UpdateCommand.Parameters.Add("@ModifyAcct", SqlDbType.VarChar).Value=sm.UserInfo.UserID
    '                '.UpdateCommand.Parameters.Add("@eSETID", SqlDbType.VarChar).Value=tmpSETID
    '                .UpdateCommand.Parameters.Add("@IDNO2", SqlDbType.VarChar).Value=IDNO2
    '                .UpdateCommand.Parameters.Add("@SETID", SqlDbType.Int).Value=tmpSETID
    '                .UpdateCommand.Parameters.Add("@IDNO1", SqlDbType.VarChar).Value=IDNO2
    '                .UpdateCommand.ExecuteNonQuery()
    '            End With
    '        End If

    '    Catch ex As Exception
    '        'tmpConn.Close()
    '        tmpTrans.Rollback()
    '        Common.MessageBox(Me, ex.ToString)
    '        Throw ex
    '    End Try
    'End Sub

    'If IDNOValue.Value <> "" And STDateValue.Value <> "" And BFDateValue.Value <> "" Then
    '    Dim dt As DataTable
    '    Dim sql As String
    '    sql=" Select b.APPLY_DATE,b.APPLY_MONEY,aw.station_name" & vbCrLf
    '    sql &= " from (select * from BIEPTBL where idno='" & IDNOValue.Value & "' and APPLY_DATE BETWEEN '" & BFDateValue.Value & "' AND '" & STDateValue.Value & "') b" & vbCrLf
    '    sql &= " left join Adp_WorkStation aw" & vbCrLf
    '    sql &= "   on aw.Station_Scheme_ID+aw.Station_Unit_ID+aw.Station_ID=b.STATION" & vbCrLf

    '    dt=DbAccess.GetDataTable(sql)
    '    If dt.Rows.Count=0 Then
    '        DataGridTable.Visible=False
    '    Else
    '        DataGridTable.Visible=True
    '        DataGrid1.DataSource=dt
    '        DataGrid1.DataBind()
    '    End If
    'Else
    '    DataGridTable.Visible=False
    'End If

#End Region

    'Dim oTest_flag As Boolean=False
    Dim aNow As Date
    'Dim au As New cAUTH
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        Call sUtl_Create0()
        If Not IsPostBack Then
            Call cCreate1()
        End If
    End Sub

    Sub sUtl_Create0()
        Dim flag_SHOW_2020x70 As Boolean = TIMS.SHOW_2020x70(sm)
        Dim flag_SHOW_2020x06 As Boolean = TIMS.SHOW_2020x06(sm)
        'Dim flag_show_actno_budid As Boolean=False '保險證號/預算別代碼 false:不顯示 true:顯示
        flag_show_actno_budid = False '保險證號/預算別代碼 false:不顯示 true:顯示
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then flag_show_actno_budid = True
        If flag_SHOW_2020x70 Then flag_show_actno_budid = True
        If flag_SHOW_2020x06 Then flag_show_actno_budid = True
        Hid_show_actno_budid.Value = ""
        If (flag_show_actno_budid) Then Hid_show_actno_budid.Value = "Y"

        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        aNow = TIMS.GetSysDateNow(objconn)
        '是否啟用西元轉民國顯示
        flag_ROC = TIMS.CHK_REPLACE2ROC_YEARS()

        'If TIMS.sUtl_ChkTest() Then oTest_flag=True
        Button4.Visible = False '已申請失業給付資料顯示
        Dim rqSD03002classver As String = Request("SD_03_002_classver")
        rqSD03002classver = TIMS.ClearSQM(rqSD03002classver)
        If rqSD03002classver = "VIEW" Then
            '已申請失業給付資料顯示
            vsOCID1val = TIMS.ClearSQM(Request("OCID"))
            IDNOValue.Value = TIMS.ChangeIDNO(Request("IDNO"))
            BFDateValue.Value = TIMS.ClearSQM(Request("BFDate"))
            STDateValue.Value = TIMS.ClearSQM(Request("STDate"))
            'Table1.Style("display")="none"
            'Table11.Style("display")="none"
            'Table12.Style("display")="none"
            'Table22.Style("display")="none"
            'Table23.Style("display")="none"
            'Table10.Style("display")="none"
            Table1.Visible = False
            'Table46.Visible=False '是否為在職者補助身分
            Table11.Visible = False
            Table12.Visible = False
            'Table22.Visible = False
            'Table23.Visible = False
            Table10.Visible = False
            'menuLab1.Text="首頁>>學員動態管理>>報到>>"
            'menuLab2.Text="學員資料審核作業"
            Button1.Visible = False
            Button2.Visible = False
            Button3.Visible = False
            Button4.Visible = True
            DataGridTable_Show(IDNOValue.Value, STDateValue.Value, BFDateValue.Value)
            Exit Sub '直接離開
        End If

        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        Hid_PreUseLimited18a.Value = ""
        If TIMS.Cst_TPlanID_PreUseLimited18a.IndexOf(sm.UserInfo.TPlanID) > -1 Then Hid_PreUseLimited18a.Value = "Y"

        '產投、充飛計畫
        Table11.Style("display") = "none"
        Table12.Style("display") = "none"
        'Table13.Style("display")="none" '"inline"
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            'Table46.Style("display")="none" '是否為在職者補助身分
            Table1.Style("display") = "none"
            Table11.Style("display") = cst_inline1 '"inline"
            Table12.Style("display") = cst_inline1 '"inline"
            TablePWTYPE.Style("display") = "none"
            'DataGridTable2.Style("display")="inline"
            'DataGridTable2.Visible=True
        Else
            'Table46.Style("display")="none" '是否為在職者補助身分(WorkSuppIdent)
            Table1.Style("display") = cst_inline1 '"inline"
            'Table22.Style("display") = "none"
            'Table23.Style("display") = "none"
            Select Case sm.UserInfo.TPlanID
                Case "06", "70"
                    '06:在職進修計畫／70:區域產業據點計畫
                    TablePWTYPE.Style("display") = "none"
                Case Else
                    TablePWTYPE.Style("display") = cst_inline1 '"inline"
                    'DataGridTable2.Style("display")="none"
                    'DataGridTable2.Visible=False
            End Select
            'If Hid_show_actno_budid.Value="Y" Then
            '    Table13.Style("display")=cst_inline1 '"inline"
            'End If
        End If

        '產投計畫
        'Dim flg_TPlanID28DBL As Boolean 
        flg_ShowTID28DBL = False '是否要顯示產投重疊資訊。
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then flg_ShowTID28DBL = True '是否要顯示產投重疊資訊。
        trTPlanID28DBL2.Visible = False
        trTPlanID28DBL3.Visible = False
        tbTPlanID28DBL1.Visible = False
        tbTPlanID28DBL1.Style("display") = "none"
        If flg_ShowTID28DBL Then
            trTPlanID28DBL2.Visible = True
            trTPlanID28DBL3.Visible = True
            tbTPlanID28DBL1.Visible = True
            tbTPlanID28DBL1.Style("display") = cst_inline1 '"inline"
        End If
    End Sub

    Sub cCreate1()
        Button2.Attributes("onclick") = "return CheckData();" '審核失敗 CheckData
        Button1.Attributes("onclick") = "return CheckData1();" '審核成功 CheckData1
        '預算別改變
        'ddlBudID.Attributes("onchange")="return Change();"
        Call SHOW_ENTERTYPE2()
        Call SHOW_ENTERTRAIN2()
        Call SHOW_BIEPTBL_DG1()
        Call SHOW_STUD_HISTORY()
        Call SHOW_DG34(IDNOValue.Value)

        'If oTest_flag Then '測
        '    Button1.Visible=True
        '    Hid_MSG1.Value="該民眾為臺東地區民眾，若符合「勞動部因應重大災害職業訓練協助計畫」受災者，請其提供證明文件，得免試入訓。"
        'End If
        If HiderrFlag.Value = cst_errFlag Then '有任何錯誤就離開不要再RUN下去了
            'Dim sErrMsg As String="資料有誤請重新查詢!" & vbCrLf
            'Common.MessageBox(Me, sErrMsg)
            Exit Sub
        End If
        If flg_ShowTID28DBL Then
            If initObj28() Then Call Search28DBL() '是否要顯示產投重疊資訊。
        End If
        If Session("_SearchStr") IsNot Nothing Then
            ViewState("_SearchStr") = Session("_SearchStr")
            Session("_SearchStr") = Nothing
        End If
        'ViewState("_SearchStr")=Session("_SearchStr")
        'Session("_SearchStr")=Nothing
        '不予補助對象，惠請查明該筆民眾身分是否符合本計畫參訓資格
        'Dim flag_Not_Subsidized_AlertMsg As Boolean=False
        'Dim rqBudID As String=TIMS.ClearSQM(Request("BudID")) '預設值
        'Dim rqSupplyID As String=TIMS.ClearSQM(Request("SupplyID")) '預設值
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            'Table23.Style("display")="inline"加入預算別產業人才投資計畫時務必顯示，可供選擇
            'If rqBudID <> "" Then Common.SetListItem(ddlBudID, rqBudID)
            'If rqSupplyID <> "" Then Common.SetListItem(ddlSupplyID, rqSupplyID)
            'ActNo
            '不予補助對象，惠請查明該筆民眾身分是否符合本計畫參訓資格
            'flag_Not_Subsidized_AlertMsg=True
            If ACTNO60.Text <> "" Then
                Dim msgAct60 As String = ""
                Dim flagAct60 As Boolean = False
                '不予補助對象，惠請查明該筆民眾身分是否符合本計畫參訓資格
                flagAct60 = TIMS.Chk_ActNoALTERMSG1(ACTNO60.Text, msgAct60)
                If flagAct60 AndAlso msgAct60 <> "" Then Common.MessageBox(Me, msgAct60)
            End If
        End If
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) = -1 Then
            If TIMS.Cst_TPlanID06Plan1.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                '不予補助對象，惠請查明該筆民眾身分是否符合本計畫參訓資格
                'flag_Not_Subsidized_AlertMsg=True
                If ACTNO60.Text <> "" Then
                    Dim msgAct60 As String = ""
                    Dim flagAct60 As Boolean = False
                    '不予補助對象，惠請查明該筆民眾身分是否符合本計畫參訓資格
                    flagAct60 = TIMS.Chk_ActNoALTERMSG1(Hid_ACTNObli.Value, msgAct60)
                    If flagAct60 AndAlso msgAct60 <> "" Then Common.MessageBox(Me, msgAct60)
                End If
            End If
            'Select Case Convert.ToString(sm.UserInfo.TPlanID)
            '    Case "06"
            'End Select
        End If
    End Sub
    '檢核
    Sub sUtl_ChkData1(ByRef ERRMSG As String)
        'CHECK
        'Dim ERRMSG As String=""
        ERRMSG = ""
        If Request("eSerNum") <> "" Then heSerNum.Value = Request("eSerNum")
        If Request("eSETID") <> "" Then heSETID.Value = Request("eSETID")
        heSerNum.Value = TIMS.ClearSQM(heSerNum.Value)
        heSETID.Value = TIMS.ClearSQM(heSETID.Value)
        If heSerNum.Value = "" OrElse heSETID.Value = "" Then
            ERRMSG &= "資料有誤請重新查詢!" & vbCrLf
            Exit Sub
        End If
        Dim drET2 As DataRow = TIMS.Get_ENTERTYPE2(heSerNum.Value, objconn)
        If Convert.ToString(drET2("eSETID")) <> heSETID.Value Then
            ERRMSG &= "資料有誤請重新查詢!" & vbCrLf
            Exit Sub
        End If
        'Dim flagYearsOld65 As Boolean=False '判斷是否為 六十五歲以上者資格
        'BUDID 02
        'Const Cst_Msg65 As String="參訓學員為65歲以上者, 其預算別一律運用就安預算!!預算別，(非就安)有誤!"  '產投
        Dim stdBLACK2TPLANID As String = ""
        Dim iStdBlackType As Integer = TIMS.Chk_StdBlackType(Me, objconn, stdBLACK2TPLANID)
        Dim aIDNO1 As String = TIMS.ClearSQM(drET2("IDNO"))
        Dim flgStdBlack As Boolean = TIMS.Get_StdBlackIDNO1(Me, iStdBlackType, stdBLACK2TPLANID, aIDNO1, objconn)
        If flgStdBlack Then
            '身分證號被處分了
            '依處分日期及年限，仍在處分期間者，e網報名審核，只能審核為失敗。
            ERRMSG &= "依處分日期及年限，仍在處分期間者，e網報名審核，只能審核為失敗。" & vbCrLf
            Exit Sub
        End If

        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            'WebRequest物件如何忽略憑證問題
            System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf TIMS.ValidateServerCertificate)
            'TLS 1.2-基礎連接已關閉: 傳送時發生未預期的錯誤 
            System.Net.ServicePointManager.SecurityProtocol = Net.SecurityProtocolType.Tls12
            '檢核學員重複參訓。
            'http://163.29.199.211/TIMSWS/timsService1.asmx
            'https://wltims.wda.gov.tw/TIMSWS/timsService1.asmx
            Dim timsSer1 As New timsService1.timsService1

            '檢核學員重複參訓。
            'Dim aIDNO1 As String=CStr(drET2("IDNO"))
            Dim aOCID1 As String = TIMS.ClearSQM(drET2("OCID1"))
            Dim xStudInfo As String = ""
            xStudInfo = ""
            TIMS.SetMyValue(xStudInfo, "IDNO", aIDNO1)
            TIMS.SetMyValue(xStudInfo, "OCID1", aOCID1)
            '檢核學員OOO有參訓時段重疊報名情形，無法儲存，請再確認！
            Call TIMS.ChkStudDouble(timsSer1, ERRMSG, "", xStudInfo)

            '28:產業人才投資計劃
            'flagYearsOld65=False 'false:否 判斷是否為 六十五歲以上者資格
            'If TIMS.Check_YearsOld65(ViewState("Birthday"), STDateValue.Value) Then
            '    flagYearsOld65=True 'true:是 判斷是否為 六十五歲以上者資格
            'End If

            'If ddlBudID.SelectedValue="請選擇" OrElse ddlBudID.SelectedValue="" Then
            '    ERRMSG &= "審核清單中,有報名學員的預算別與補助比例尚未選擇,請設定!" & vbCrLf
            'End If
            'If ddlSupplyID.SelectedValue="請選擇" OrElse ddlSupplyID.SelectedValue="" Then
            '    ERRMSG &= "審核清單中,有報名學員的預算別與補助比例尚未選擇,請設定!" & vbCrLf
            'End If
            'If ddlBudID.SelectedValue="97" AndAlso ddlSupplyID.SelectedValue <> "2" Then
            '    ERRMSG &= "預算別為協助,補助比例應為100% !" & vbCrLf
            'End If
            'If ddlBudID.SelectedValue="99" AndAlso ddlSupplyID.SelectedValue <> "9" Then
            '    ERRMSG &= "預算別為不補助,補助比例應為0% !" & vbCrLf
            'End If
            '修改說明:有關參訓學員已65歲以上者（依該參訓學員出生年月日及開訓日期判斷），其預算別一律運用就安預算 2013/11/20
            'true:是 判斷是否為 六十五歲以上者資格
            'If flagYearsOld65 AndAlso ddlBudID.SelectedValue <> "02" Then
            '    ERRMSG += Cst_Msg65 & vbCrLf
            'End If

        End If
        'https://jira.turbotech.com.tw/browse/TIMSC-58
        '投保證號為075、175（裁減續保）、076（職災續保）、09（訓）、176皆為不予補助對象，惠請查明該筆民眾身分是否符合本計畫參訓資格。
        'If ERRMSG <> "" Then
        '    Common.MessageBox(Me, ERRMSG)
        '    Exit Sub
        'End If
    End Sub

    Function GET_STUD_ENTERTYPE2_ESETID(eSerNum As String) As String
        Dim PMS1 As New Hashtable From {{"eSerNum", eSerNum}}
        Dim SSQL As String = "SELECT * FROM STUD_ENTERTYPE2 WITH(NOLOCK) WHERE eSerNum=@eSerNum"
        Dim dr As DataRow = DbAccess.GetOneRow(SSQL, objconn, PMS1)
        If dr Is Nothing Then Return ""
        Return Convert.ToString(dr("ESETID")) 'heSETID.Value =
    End Function

    '審核成功
    Sub SaveDataOK1()
        'Dim flagYearsOld65 As Boolean=False '判斷是否為 六十五歲以上者資格 ''BUDID 02
        'Const Cst_Msg65 As String="參訓學員為65歲以上者, 其預算別一律運用就安預算!!預算別，(非就安)有誤!"  '產投
        heSerNum.Value = TIMS.ClearSQM(Request("eSerNum"))
        heSETID.Value = TIMS.ClearSQM(Request("eSETID"))
        Dim ERRMSG As String = ""
        Dim drET2 As DataRow = Nothing
        If heSerNum.Value <> "" Then drET2 = TIMS.Get_ENTERTYPE2(heSerNum.Value, objconn)
        If heSerNum.Value = "" OrElse heSETID.Value = "" Then
            ERRMSG = "資料有誤請重新查詢!" & vbCrLf
            Common.MessageBox(Me, ERRMSG)
            Exit Sub
        ElseIf drET2 Is Nothing Then
            ERRMSG = "資料有誤請重新查詢!!" & vbCrLf
            Common.MessageBox(Me, ERRMSG)
            Exit Sub
        ElseIf Convert.ToString(drET2("eSETID")) <> heSETID.Value Then
            ERRMSG = "資料有誤請重新查詢!" & vbCrLf
            Common.MessageBox(Me, ERRMSG)
            Exit Sub
        End If
        If ERRMSG <> "" Then
            Common.MessageBox(Me, ERRMSG)
            Exit Sub
        End If
        'sUtl_ChkData1
        Call sUtl_ChkData1(ERRMSG)
        If ERRMSG <> "" Then
            Common.MessageBox(Me, ERRMSG)
            Exit Sub
        End If
        '2006/03/ add conn by matt
        aNow = TIMS.GetSysDateNow(objconn)
        'SAVE
        ViewState("Subject") = "" 'Nothing
        ViewState("ExamNo") = "" 'Nothing
        ViewState("RelEnterDate") = "" 'Nothing
        vsExamDate = "" 'Nothing
        ViewState("CheckInDate") = "" 'Nothing
        ViewState("STUD_NAME") = "" 'Nothing
        ViewState("Email") = "" 'Nothing
        Dim sql As String = ""
        Dim iSEID As Integer = 0
        Dim iSETID As Integer = 0
        Dim iSerNum As Integer = 0
        Dim tmpOCID1 As String = ""

        'STUD_ENTERTRAIN2
        Dim pmsTRAIN2 As New Hashtable From {{"eSerNum", Val(heSerNum.Value)}}
        sql = " SELECT * FROM STUD_ENTERTRAIN2 WITH(NOLOCK) WHERE eSerNum=@eSerNum"
        Dim drTRAIN2 As DataRow = DbAccess.GetOneRow(sql, objconn, pmsTRAIN2)
        If drTRAIN2 IsNot Nothing Then iSEID = drTRAIN2("SEID")

        'STUD_ENTERSUBDATA2
        Dim pmsSUBDATA2 As New Hashtable From {{"eSerNum", Val(heSerNum.Value)}}
        sql = " SELECT * FROM STUD_ENTERSUBDATA2 WITH(NOLOCK) WHERE eSerNum=@eSerNum" '受訓前任職資料
        Dim drSUBDATA2 As DataRow = DbAccess.GetOneRow(sql, objconn, pmsSUBDATA2)

        'STUD_ENTERTEMP2,(STUD_ENTERTYPE2) 
        Dim pmsTEMP2 As New Hashtable From {{"eSerNum", Val(heSerNum.Value)}}
        sql = ""
        sql &= " SELECT a.*" & vbCrLf
        sql &= " ,b.OCID1 ,b.RID" & vbCrLf
        sql &= " FROM STUD_ENTERTEMP2 a WITH(NOLOCK)" & vbCrLf
        sql &= " JOIN STUD_ENTERTYPE2 b WITH(NOLOCK) ON b.eSETID=a.eSETID" & vbCrLf
        sql &= " WHERE b.eSerNum=@eSerNum" & vbCrLf
        Dim dtTEMP2 As DataTable = DbAccess.GetDataTable(sql, objconn, pmsTEMP2)
        Dim drTEMP2 As DataRow = Nothing
        If TIMS.dtHaveDATA(dtTEMP2) AndAlso dtTEMP2.Rows.Count > 0 Then
            drTEMP2 = dtTEMP2.Rows(0) 'STUD_ENTERTEMP2
            heSETID.Value = Convert.ToString(drTEMP2("eSETID"))
        End If
        If TIMS.dtNODATA(dtTEMP2) OrElse drTEMP2 Is Nothing Then
            heSETID.Value = ""
            'ERRMSG="學員資料有誤, 無法儲存, 請再確認!" & vbCrLf
            Dim strErrmsg As String = ""
            'strErrmsg &= "/* ex.ToString */" & vbCrLf
            'strErrmsg &= ex.ToString & vbCrLf
            strErrmsg &= cst_errmsg1 & vbCrLf
            strErrmsg &= " AND eSerNum=" & heSerNum.Value & vbCrLf
            heSETID.Value = GET_STUD_ENTERTYPE2_ESETID(Val(heSerNum.Value))
            strErrmsg &= " AND eSETID=" & heSETID.Value & vbCrLf
            'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)

            Common.MessageBox(Me, cst_errmsg1)
            Exit Sub
        End If
        If heSETID.Value = "" Then
            Common.MessageBox(Me, cst_errmsg1)
            Exit Sub
        End If
        ViewState("STUD_NAME") = drTEMP2("Name") '報考人姓名
        ViewState("Email") = Convert.ToString(drTEMP2("Email")) '報考人Email
        Dim drOC1 As DataRow = TIMS.GetOCIDDate(drTEMP2("OCID1"), objconn)
        If drOC1 Is Nothing Then
            Common.MessageBox(Me, cst_errmsg1)
            Exit Sub
        End If
        Dim ExamOcid1 As String = drTEMP2("OCID1").ToString
        'Dim ExamNo1 As String '取出班級的CLASSID +期別 成為准考證編碼的前面的固定碼
        Dim ExamNo1 As String = TIMS.Get_ExamNo1(ExamOcid1, objconn)
        If ExamNo1 = "" OrElse ExamNo1.Length < 6 Then '防呆
            Common.MessageBox(Me, "班級的代號 與期別有誤，請確認班級狀態")
            Exit Sub
        End If

        '    '(職前課程邏輯)若為下列計畫, 則依4項不予錄訓規定設定邏輯判斷學員是否可參訓:
        '    ' https://jira.turbotech.com.tw/browse/TIMSC-142
        '    ' 呼叫 TIMS.Get_ChkIsJobsCounse44() 進行檢查
        'Dim IDNOt As String=Convert.ToString(drTEMP2("IDNO"))
        'Dim OCIDVal As String=Convert.ToString(drTEMP2("OCID1"))
        'If OCIDVal <> "" Then
        '    Dim htSS As New Hashtable 'htSS Hashtable() '
        '    htSS.Add("IDNOt", IDNOt)
        '    htSS.Add("OCIDVal", OCIDVal)
        '    htSS.Add("SENTERDATE", TIMS.cdate3(drOC1("SENTERDATE")))
        '    ERRMSG &= TIMS.Get_ChkIsJobsCounse44(Me, htSS, TIMS.cst_FunID_e網報名審核, objconn)
        '    If ERRMSG <> "" Then
        '        Common.MessageBox(Me, ERRMSG)
        '        Exit Sub
        '    End If
        'End If

        '2006/03/ add conn by matt
        'Call TIMS.OpenDbConn(conn)
        Dim drTYPE2 As DataRow = Nothing 'STUD_ENTERTYPE2
        Using TransConn As SqlConnection = DbAccess.GetConnection()
            Dim Trans As SqlTransaction = DbAccess.BeginTrans(TransConn)
            Try
                Dim da As SqlDataAdapter = Nothing 'Stud_EnterTemp
                Dim dr As DataRow = Nothing 'STUD_ENTERTEMP
                Dim dt As DataTable = Nothing 'Stud_EnterTemp
                sql = " SELECT * FROM STUD_ENTERTEMP WHERE IDNO='" & TIMS.ChangeIDNO(drTEMP2("IDNO")) & "'"
                dt = DbAccess.GetDataTable(sql, da, Trans)
                If dt.Rows.Count = 0 Then
                    '無資料新增1筆
                    dr = dt.NewRow
                    dt.Rows.Add(dr)
                    'dr("SETID")=SETID 'STUD_ENTERTEMP_SETID_SEQ
                    iSETID = DbAccess.GetNewId(Trans, "STUD_ENTERTEMP_SETID_SEQ,STUD_ENTERTEMP,SETID")
                    dr("SETID") = iSETID
                    dr("IDNO") = TIMS.ChangeIDNO(drTEMP2("IDNO"))
                    dr("Name") = drTEMP2("Name")
                    dr("Sex") = drTEMP2("Sex")
                    dr("Birthday") = drTEMP2("Birthday")

                    Dim vPassPortNO As String = ""
                    Select Case Convert.ToString(drTEMP2("PassPortNO"))
                        Case "1", "2"
                            vPassPortNO = Convert.ToString(drTEMP2("PassPortNO"))
                        Case Else
                            vPassPortNO = "2"
                    End Select
                    dr("PassPortNO") = If(vPassPortNO <> "", vPassPortNO, "2")
                    dr("MaritalStatus") = drTEMP2("MaritalStatus")
                    dr("DegreeID") = drTEMP2("DegreeID")
                    dr("GradID") = drTEMP2("GradID")
                    dr("School") = drTEMP2("School")
                    dr("Department") = drTEMP2("Department")
                    dr("MilitaryID") = drTEMP2("MilitaryID")
                    dr("ZipCode") = drTEMP2("ZipCode")
                    dr("ZIPCODE6W") = drTEMP2("ZIPCODE6W")
                    dr("Address") = drTEMP2("Address")

                    dr("Phone1") = drTEMP2("Phone1")
                    dr("Phone2") = drTEMP2("Phone2")
                    dr("CellPhone") = drTEMP2("CellPhone")
                    dr("Email") = drTEMP2("Email")
                    dr("IsAgree") = If(Convert.ToString(drTEMP2("IsAgree")) <> "", "Y", "N")
                    dr("eSETID") = drTEMP2("eSETID")
                    dr("ModifyAcct") = sm.UserInfo.UserID
                    dr("ModifyDate") = aNow 'Now
                    DbAccess.UpdateDataTable(dt, da, Trans) 'Stud_EnterTemp
                Else
                    '有資料做 UPDATE
                    dr = dt.Rows(0)
                    iSETID = dr("SETID")
                    'Update_StudEnterTemp(Convert.ToString(drTEMP2("eSETID")), SETID, conn, Trans)
                    'Update_StudEnterTemp(Convert.ToString(drTEMP2("eSETID")), drTEMP2("IDNO"), SETID, conn, Trans)
                    If Convert.ToString(drTEMP2("IDNO")) <> "" Then
                        'Stud_EnterTemp
                        Call TIMS.UPDATE_STUDENTERTEMP(Me, Convert.ToString(drTEMP2("IDNO")), TransConn, Trans)
                    End If
                End If
                If iSETID <> 0 Then
                    sql = " UPDATE STUD_ENTERTEMP2 SET SETID=" & iSETID & " WHERE eSETID='" & heSETID.Value & "'"
                    DbAccess.ExecuteNonQuery(sql, Trans)
                    'Update Stud_EnterTemp2
                    'Dim drT2 As DataRow=Nothing
                    'dt1=DbAccess.GetDataTable(sql, da1, Trans)
                    'drT2=dt1.Rows(0)
                    'drT2("SETID")=iSETID
                    'drT2("ModifyAcct")=sm.UserInfo.UserID
                    'drT2("ModifyDate")=aNow 'Now
                    'DbAccess.UpdateDataTable(dt1, da1, Trans)
                End If
                Dim da1 As SqlDataAdapter = Nothing 'Stud_EnterType2
                'STUD_ENTERTYPE2
                'Dim pmsTYPE2 As New Hashtable From {{"eSerNum", Val(heSerNum.Value)}}
                sql = String.Concat("SELECT * FROM STUD_ENTERTYPE2 WHERE eSerNum=", Val(heSerNum.Value))
                Dim dtTYPE2 As DataTable = DbAccess.GetDataTable(sql, da1, Trans)
                drTYPE2 = dtTYPE2.Rows(0)
                tmpOCID1 = Convert.ToString(drTYPE2("OCID1"))
                drTYPE2("CCLID") = TIMS.Change0(drTYPE2("CCLID"))

                ViewState("Subject") = TIMS.Get_Subject(objconn, drTYPE2("OCID1").ToString)
                'ViewState("CheckInDate")=TIMS.GET_CheckInDate(drTYPE2("OCID1").ToString)
                'vsExamDate=TIMS.GET_ExamDate(drTYPE2("OCID1").ToString)
                ViewState("CheckInDate") = If(Convert.ToString(drOC1("CheckInDate")) <> "", If(flag_ROC, Common.FormatDate2Roc(drOC1("CheckInDate")), drOC1("CheckInDate")), "")
                vsExamDate = If(Convert.ToString(drOC1("ExamDate")) <> "", If(flag_ROC, Common.FormatDate2Roc(drOC1("ExamDate")), drOC1("ExamDate")), "")
                'Stud_EnterType2  drTYPE2 dt1
                'Stud_EnterType  dr dt

                '取出准考證號   Start
                Dim ExamPlanID As String = If(Convert.ToString(drTYPE2("PlanID")) <> "", Convert.ToString(drTYPE2("PlanID")), sm.UserInfo.PlanID)
                Dim flgChkExamNo As Boolean = TIMS.Chk_NewExamNOc(ExamPlanID, ExamOcid1, objconn)
                If Not flgChkExamNo Then
                    Common.MessageBox(Me, "班級的代號 與計畫不符，請確認班級狀態(取出准考證號)!")
                    DbAccess.RollbackTrans(Trans)
                    Call TIMS.CloseDbConn(TransConn)
                    Exit Sub
                End If
                '准考證號
                Dim NewExamNO As String = TIMS.Get_NewExamNOt(ExamPlanID, ExamNo1, ExamOcid1, Trans)
                If NewExamNO = "" Then
                    Common.MessageBox(Me, "班級的代號 與計畫不符，請確認班級狀態(取出准考證號)!!")
                    DbAccess.RollbackTrans(Trans)
                    Call TIMS.CloseDbConn(TransConn)
                    Exit Sub
                End If
                '取出准考證號   End

                '儲存報名班級資料--
                sql = String.Concat(" SELECT * FROM STUD_ENTERTYPE WHERE SETID=", iSETID, " AND EnterDate=", TIMS.To_date(drTYPE2("EnterDate")), " ORDER BY SerNum DESC")
                dt = DbAccess.GetDataTable(sql, da, Trans)
                'Stud_EnterType2  drTYPE2 dt1
                'Stud_EnterType  dr dt
                If dt.Rows.Count = 0 Then
                    iSerNum = 1
                    dr = dt.NewRow
                    dt.Rows.Add(dr)
                    dr("SETID") = iSETID
                    dr("EnterDate") = drTYPE2("EnterDate")
                    dr("SerNum") = iSerNum
                Else
                    ff3 = String.Concat("OCID1=", drTYPE2("OCID1"))
                    If dt.Select(ff3).Length > 0 Then
                        'double_OCID=True ''同一SETID(學員)，產生重複報名同一班(OCID) 'SerNum=dt.Select("OCID1='" & drTYPE2("OCID1") & "'")(0)("SerNum") + 1
                        dr = dt.Select(ff3)(0)
                        iSerNum = dt.Rows(0)("SerNum")
                    Else
                        'Not 同一SETID(學員)，產生重複報名同一班(OCID)
                        iSerNum = dt.Rows(0)("SerNum") + 1
                        dr = dt.NewRow
                        dt.Rows.Add(dr)
                        dr("SETID") = iSETID
                        dr("EnterDate") = drTYPE2("EnterDate")
                        dr("SerNum") = iSerNum
                    End If
                End If
                dr("ExamNo") = NewExamNO
                ViewState("ExamNo") = NewExamNO '准考證號
                dr("OCID1") = drTYPE2("OCID1")
                dr("TMID1") = drTYPE2("TMID1")
                dr("OCID2") = drTYPE2("OCID2")
                dr("TMID2") = drTYPE2("TMID2")
                dr("OCID3") = drTYPE2("OCID3")
                dr("TMID3") = drTYPE2("TMID3")
                '就服站不可異動
                If Convert.ToString(dr("EnterPath")) <> "W" Then
                    dr("EnterChannel") = 1
                    dr("EnterPath") = "E" 'E網報名審核成功
                End If
                dr("IdentityID") = If(HidIdentityID.Value <> "", HidIdentityID.Value, drTYPE2("IdentityID")) 'drTYPE2("IdentityID")
                dr("RID") = drTYPE2("RID")
                dr("PlanID") = drTYPE2("PlanID")
                dr("RelEnterDate") = drTYPE2("RelEnterDate")

                '報名日期
                ViewState("RelEnterDate") = If(flag_ROC, Common.FormatDate2Roc(drTYPE2("RelEnterDate")), Convert.ToString(drTYPE2("RelEnterDate")))
                dr("CCLID") = drTYPE2("CCLID")

                If IsDBNull(dr("eSerNum")) AndAlso Convert.ToString(drTYPE2("eSerNum")) <> "" Then dr("eSerNum") = Convert.ToString(drTYPE2("eSerNum"))
                If IsDBNull(dr("eSETID")) AndAlso Convert.ToString(drTYPE2("eSETID")) <> "" Then dr("eSETID") = Convert.ToString(drTYPE2("eSETID"))
                '受訓前任職資料start 'STUD_ENTERSUBDATA2
                If drSUBDATA2 IsNot Nothing Then
                    dr("PriorWorkType1") = If(Convert.ToString(drSUBDATA2("PriorWorkType1")) = "", Convert.DBNull, Convert.ToString(drSUBDATA2("PriorWorkType1")))
                    dr("PriorWorkOrg1") = If(Convert.ToString(drSUBDATA2("PriorWorkOrg1")) = "", Convert.DBNull, Convert.ToString(drSUBDATA2("PriorWorkOrg1")))
                    dr("ActNo") = If(Convert.ToString(drSUBDATA2("ActNo")) = "", Convert.DBNull, Convert.ToString(drSUBDATA2("ActNo")))
                    dr("SOfficeYM1") = If(Convert.ToString(drSUBDATA2("SOfficeYM1")) = "", Convert.DBNull, Convert.ToString(drSUBDATA2("SOfficeYM1")))
                    dr("FOfficeYM1") = If(Convert.ToString(drSUBDATA2("FOfficeYM1")) = "", Convert.DBNull, Convert.ToString(drSUBDATA2("FOfficeYM1")))
                End If
                '受訓前任職資料end
                '把SEID(線上報名資料的流水號-產學訓)寫入Stud_Entertype中---start
                If Not IsDBNull(drTYPE2("eSerNum")) Then
                    If iSEID > 0 Then dr("SEID") = iSEID
                End If
                dr("TransDate") = aNow 'Now
                dr("ModifyAcct") = sm.UserInfo.UserID
                dr("ModifyDate") = aNow 'Now
                Dim sTemp1 As String = ""
                sTemp1 = "CFIRE1,CFIRE1NS,CFIRE1REASON,CFIRE1MACCT,CFIRE1MDATE"
                sTemp1 &= ",CMASTER1,CMASTER1NS,CMASTER1REASON,CMASTER1MACCT,CMASTER1MDATE,CMASTER1NT,CFIRE1R2"
                Dim aTemp1 As String() = sTemp1.Split(",")
                For Each tmpCSTR1 As String In aTemp1
                    'type2 有值 type 無值 才動作
                    If Convert.ToString(drTYPE2(tmpCSTR1)) <> "" AndAlso Convert.ToString(dr(tmpCSTR1)) = "" Then
                        dr(tmpCSTR1) = drTYPE2(tmpCSTR1)
                    End If
                Next
                'For i As Integer=0 To aTemp1.Length - 1
                '    Dim tmpCT1 As String=aTemp1(i) 'type2 有值 type 無值 才動作
                '    If Convert.ToString(drTYPE2(tmpCT1)) <> "" AndAlso Convert.ToString(dr(tmpCT1))="" Then
                '        dr(tmpCT1)=drTYPE2(tmpCT1)
                '    End If
                'Next

                'If Convert.ToString(drTYPE2("CFIRE1")) <> "" AndAlso Convert.ToString(dr("CFIRE1"))="" Then dr("CFIRE1")=drTYPE2("CFIRE1")
                'CFIRE1 CFIRE1NS CFIRE1REASON CFIRE1MACCT CFIRE1MDATE 
                'CMASTER1 CMASTER1NS CMASTER1REASON CMASTER1MACCT CMASTER1MDATE CMASTER1NT CFIRE1R2 
                'EXAMPLUS PREEXDATE
                'CFIRE1 CFIRE1NS CFIRE1REASON CFIRE1MACCT CFIRE1MDATE 
                'CMASTER1 CMASTER1NS CMASTER1REASON CMASTER1MACCT CMASTER1MDATE CMASTER1NT CFIRE1R2 PREEXDATE
                DbAccess.UpdateDataTable(dt, da, Trans) 'Stud_EnterType

                '---end
                drTYPE2("SETID") = iSETID
                drTYPE2("SerNum") = iSerNum
                drTYPE2("ExamNo") = NewExamNO
                'signUpStatus 0:收件完成 1:報名成功 2:報名失敗 3:正取(Key_SelResult) 4:備取 5:未錄取
                drTYPE2("signUpStatus") = 1
                drTYPE2("ModifyAcct") = sm.UserInfo.UserID
                drTYPE2("ModifyDate") = aNow 'Now
                DbAccess.UpdateDataTable(dtTYPE2, da1, Trans) 'Stud_EnterType2

                '假如是插班,則直接進入參訓狀態(STUD_SELRESULT)
                If Convert.ToString(drTYPE2("CCLID")) <> "" OrElse TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    'sql="SELECT * FROM Stud_SelResult WHERE SETID='" & SETID & "' and EnterDate='" & drTYPE2("EnterDate") & "' and SerNum='" & SerNum & "'" '★
                    sql = ""
                    sql &= " SELECT * FROM STUD_SELRESULT"
                    sql &= " WHERE SETID='" & iSETID & "' AND EnterDate=" & TIMS.To_date(drTYPE2("EnterDate")) & " AND SerNum='" & iSerNum & "'"
                    dt = DbAccess.GetDataTable(sql, da, Trans)
                    If dt.Rows.Count = 0 Then
                        dr = dt.NewRow
                        dt.Rows.Add(dr)
                        dr("SETID") = iSETID
                        dr("EnterDate") = drTYPE2("EnterDate")
                        dr("SerNum") = iSerNum
                    Else
                        dr = dt.Rows(0)
                    End If
                    dr("OCID") = drTYPE2("OCID1")
                    dr("Admission") = "Y" '錄取
                    dr("SelResultID") = "01" '01:正取
                    'SELRESULTID: 01:正取 02:備取 03:未錄取 04:缺考 05:審核中
                    If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                        dr("Admission") = Convert.DBNull '是否錄取 未填寫
                        dr("SelResultID") = TIMS.cst_SelResultID_審核中 '"05:審核中 03:不錄取(未錄取)
                    End If
                    dr("RID") = sm.UserInfo.RID
                    dr("PlanID") = sm.UserInfo.PlanID
                    dr("ModifyAcct") = sm.UserInfo.UserID
                    dr("ModifyDate") = aNow 'Now
                    DbAccess.UpdateDataTable(dt, da, Trans)

                    sql = ""
                    Dim UC_sql As String = " UPDATE Class_ClassInfo SET IsCalculate='Y' WHERE OCID='" & drTYPE2("OCID1") & "'"
                    DbAccess.ExecuteNonQuery(UC_sql, Trans)
                End If

                'https://jira.turbotech.com.tw/browse/TIMSC-150
                'If Hid_MSGADIDN.Value <> "" Then
                '    Dim ss As String = ""
                '    Call TIMS.SetMyValue(ss, "ADID", Hid_MSGADIDN.Value)
                '    Call TIMS.SetMyValue(ss, "OCID", drTYPE2("OCID1"))
                '    Call TIMS.SetMyValue(ss, "IDNO", drTEMP2("IDNO"))
                '    Call TIMS.SetMyValue(ss, "ESETID", drTYPE2("ESETID"))
                '    Call TIMS.SUtl_AddDISASTER(Me, ss, Trans)
                'End If

                DbAccess.CommitTrans(Trans)
            Catch ex As Exception
                Dim strErrmsg As String = ""
                strErrmsg &= "/* ex.ToString */" & vbCrLf
                strErrmsg &= ex.ToString & vbCrLf
                strErrmsg &= cst_errmsg1 & vbCrLf
                strErrmsg &= " AND eSerNum=" & heSerNum.Value & vbCrLf
                strErrmsg &= " AND eSETID=" & heSETID.Value & vbCrLf
                strErrmsg &= " AND Subject=" & ViewState("Subject") & vbCrLf
                strErrmsg &= " AND ExamNo=" & ViewState("ExamNo") & vbCrLf
                strErrmsg &= " AND RelEnterDate=" & ViewState("RelEnterDate") & vbCrLf
                strErrmsg &= " AND CheckInDate=" & ViewState("CheckInDate") & vbCrLf
                strErrmsg &= " AND vsExamDate=" & vsExamDate & vbCrLf
                strErrmsg &= " AND stud_name=" & ViewState("STUD_NAME") & vbCrLf
                strErrmsg &= " AND Email=" & ViewState("Email") & vbCrLf
                strErrmsg &= " AND isEmailFail=" & ViewState("isEmailFail") & vbCrLf '已發送失敗E-MAIL
                strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
                'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                Call TIMS.WriteTraceLog(strErrmsg)

                DbAccess.RollbackTrans(Trans)
                Call TIMS.CloseDbConn(TransConn)
            End Try
            Call TIMS.CloseDbConn(TransConn)
        End Using

        If drTYPE2 Is Nothing Then Return '(審核有誤離開儲存)

        Session("_SearchStr") = ViewState("_SearchStr")
        '20090601 by Jimmy 取得e網報名審核成功「說明事項」內容 -- begin
        ViewState("EComment") = ""
        Dim tmpdt As DataTable = Nothing
        Dim tmpdr As DataRow = Nothing
        Dim tmpDistID As String = ""
        Dim tmpOrgID As String = ""
        Dim sqlstr As String = ""
        sqlstr = " SELECT DistID FROM Auth_Relship WHERE RID='" & Convert.ToString(drTYPE2("RID")) & "' "
        tmpDistID = DbAccess.ExecuteScalar(sqlstr, objconn)
        If tmpDistID Is Nothing Then tmpDistID = ""
        sqlstr = " SELECT b.OrgID FROM Auth_Relship a JOIN Org_OrgInfo b ON a.orgid=b.orgid WHERE a.RID='" & Convert.ToString(drTYPE2("RID")) & "' "
        tmpOrgID = DbAccess.ExecuteScalar(sqlstr, objconn)
        If tmpOrgID Is Nothing Then tmpOrgID = ""
        'Dim oConn As SqlConnection=DbAccess.GetConnection() 'Call TIMS.OpenDbConn(oConn)
        tmpdt = TIMS.Get_FinalEComment("Class", Convert.ToString(tmpOrgID), Convert.ToString(drTYPE2("OCID1")), Convert.ToString(drTYPE2("RID")), Convert.ToString(tmpDistID), Convert.ToString(sm.UserInfo.PlanID), Nothing, objconn) '20090618 by Jimmy 修改 依需求訓練機構(中心)都要for不同年度、計畫、轄區
        'Call TIMS.CloseDbConn(oConn)
        Call TIMS.OpenDbConn(objconn)
        If TIMS.dtHaveDATA(tmpdt) Then
            tmpdr = tmpdt.Rows(0)
            ViewState("EComment") = Convert.ToString(tmpdr("eComment"))
        End If
        '20090601 by Jimmy 取得e網報名審核成功「說明事項」內容 -- end
        'ViewState("EmailSend") 為發送或不發送
        ViewState("EmailSend") = TIMS.CheckEmailSend(Me, "", "", objconn)
        Dim path3_from_emailaddress As String = TIMS.Utl_GetConfigSet("from_emailaddress")
        If Convert.ToString(path3_from_emailaddress) = "" Then path3_from_emailaddress = TIMS.Cst_SendMail3_from_emailaddress
        'Dim mail_msg As String=""
        If ViewState("Email").ToString <> "" AndAlso ViewState("EmailSend") Then
            '20090601 by Jimmy add 依需求將「甄試日期」改為「說明事項」，新增 eComment 參數
            Dim htSS As New Hashtable From {
                {"TPlanID", Convert.ToString(sm.UserInfo.TPlanID)},
                {"Stud_Name", Convert.ToString(ViewState("STUD_NAME"))},
                {"Subject", Convert.ToString(ViewState("Subject"))},
                {"ExamNo", Convert.ToString(ViewState("ExamNo"))},
                {"RelEnterDate", Convert.ToString(ViewState("RelEnterDate"))},
                {"ExamDate", Convert.ToString(vsExamDate)},
                {"CheckInDate", Convert.ToString(ViewState("CheckInDate"))},
                {"EComment", Convert.ToString(ViewState("EComment"))},
                {"Email", TIMS.ChangeIDNO(ViewState("Email"))},
                {"from_emailaddress", path3_from_emailaddress},
                {"signUpMemo", ""},
                {"sRIDOrgName", ""},
                {"sType", TIMS.Cst_SendMail3_CheckedOK}
            } 'htSS Hashtable() 
            Dim mail_msg As String = TIMS.SendMail3(htSS)
            'If mail_msg <> "" Then Common.RespWrite(Me, "<script>alert('" & mail_msg & "');</script>")
            If mail_msg <> "" Then
                Common.RespWrite(Me, "<script>alert('" & mail_msg & "');</script>")
            Else
                Dim pmsU2400 As New Hashtable From {{"eSerNum", Val(heSerNum.Value)}}
                Dim sqlU2400 As String = " UPDATE STUD_ENTERTYPE2 SET isEmailFail='O' WHERE eSerNum=@eSerNum"
                DbAccess.ExecuteNonQuery(sqlU2400, objconn, pmsU2400)
            End If
        End If
        Dim sMemo As String = String.Concat("&動作=審核成功", "&NAME=", TIMS.ClearSQM(Name.Text))
        '寫入Log查詢 SubInsAccountLog1 (Auth_Accountlog)
        Call TIMS.SubInsAccountLog1(Me, Request("ID"), TIMS.cst_wm修改, TIMS.cst_wmdip2, tmpOCID1, sMemo)
        Common.RespWrite(Me, "<script>alert('儲存成功');location.href='SD_01_004.aspx?ID=" & Request("ID") & "';</script>")
    End Sub
    ''' <summary>'審核成功 </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim objLock As New Object
        SyncLock objLock
            SyncLock TIMS.objLock_SD01004
                Call SaveDataOK1() '審核成功
            End SyncLock
        End SyncLock
    End Sub

    '審核失敗
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'Dim flagError As Boolean=False
        Dim sql As String = ""
        'Dim dt As DataTable
        'Dim dr As DataRow
        'Dim da As SqlDataAdapter=Nothing
        'Dim Trans As SqlTransaction=Nothing
        If Request("eSerNum") <> "" Then heSerNum.Value = TIMS.ClearSQM(Request("eSerNum"))
        If Request("eSETID") <> "" Then heSETID.Value = TIMS.ClearSQM(Request("eSETID"))
        heSerNum.Value = TIMS.ClearSQM(heSerNum.Value)
        heSETID.Value = TIMS.ClearSQM(heSETID.Value)
        'ViewState("Request_eSerNum")=TIMS.ClearSQM(Request("eSerNum")).Trim.Replace("'", "''")
        If heSerNum.Value = "" OrElse Not IsNumeric(heSerNum.Value) Then
            Dim strErrmsg As String = ""
            strErrmsg &= cst_errmsg1 & vbCrLf
            strErrmsg &= " AND eSerNum=" & heSerNum.Value & vbCrLf
            strErrmsg &= " AND eSETID=" & heSETID.Value & vbCrLf
            strErrmsg &= " AND Subject=" & ViewState("Subject") & vbCrLf
            strErrmsg &= " AND ExamNo=" & ViewState("ExamNo") & vbCrLf
            strErrmsg &= " AND RelEnterDate=" & ViewState("RelEnterDate") & vbCrLf
            strErrmsg &= " AND CheckInDate=" & ViewState("CheckInDate") & vbCrLf
            strErrmsg &= " AND vsExamDate=" & vsExamDate & vbCrLf
            strErrmsg &= " AND stud_name=" & ViewState("STUD_NAME") & vbCrLf
            strErrmsg &= " AND Email=" & ViewState("Email") & vbCrLf
            strErrmsg &= " AND isEmailFail=" & ViewState("isEmailFail") & vbCrLf '已發送失敗E-MAIL
            'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)

            Common.MessageBox(Me, cst_errmsg1)
            Exit Sub
        End If

        '2006/03/ add conn by matt
        aNow = TIMS.GetSysDateNow(objconn)
        ViewState("Subject") = "" 'Nothing
        ViewState("ExamNo") = "" 'Nothing
        ViewState("RelEnterDate") = "" 'Nothing
        ViewState("CheckInDate") = "" 'Nothing
        vsExamDate = "" 'Nothing
        ViewState("STUD_NAME") = "" 'Nothing
        ViewState("Email") = "" 'Nothing
        ViewState("isEmailFail") = ""          '已發送失敗E-MAIL
        Dim tmpOCID1 As String = ""
        Dim dt1 As DataTable
        Dim dr1 As DataRow
        TIMS.OpenDbConn(objconn)
        Try
            sql = ""
            sql &= " SELECT a.* ,b.OCID1" & vbCrLf
            sql &= " FROM Stud_EnterTemp2 a" & vbCrLf
            sql &= " JOIN Stud_EnterType2 b ON b.eSETID=a.eSETID" & vbCrLf
            sql &= " WHERE b.eSerNum='" & heSerNum.Value & "'" & vbCrLf
            dt1 = DbAccess.GetDataTable(sql, objconn)
            If dt1.Rows.Count = 0 Then
                Common.MessageBox(Me, cst_errmsg2)
                Exit Sub
            End If
            dr1 = dt1.Rows(0)
            Dim drX As DataRow = TIMS.GetOCIDDate(dr1("OCID1"), objconn)
            sql = ""
            sql &= " SELECT b.SETID, b.enterdate, b.sernum, b.OCID1" & vbCrLf
            sql &= " FROM Stud_EnterType b" & vbCrLf
            sql &= " JOIN Stud_EnterTemp a ON a.setid=b.setid" & vbCrLf
            sql &= " WHERE b.ocid1=@ocid1 AND a.idno=@idno" & vbCrLf
            Dim sCmd As New SqlCommand(sql, objconn)
            'TIMS.OpenDbConn(objconn)
            Dim dt3 As New DataTable
            With sCmd
                .Parameters.Clear()
                .Parameters.Add("ocid1", SqlDbType.VarChar).Value = dr1("ocid1")
                .Parameters.Add("idno", SqlDbType.VarChar).Value = dr1("idno")
                dt3.Load(.ExecuteReader())
            End With
            If dt3.Rows.Count > 0 Then
                For Each dr3 As DataRow In dt3.Rows
                    '審核失敗刪除錄取資料。
                    Call Del_StudSelResult(dr3("SETID"), dr3("OCID1"), objconn)
                Next
            End If
            ViewState("STUD_NAME") = dr1("Name") '報考人姓名
            ViewState("Email") = Convert.ToString(dr1("Email")) '報考人Email
            sql = " SELECT * FROM Stud_EnterType2 WHERE eSerNum='" & heSerNum.Value & "' "
            dt1 = DbAccess.GetDataTable(sql, objconn)
            dr1 = dt1.Rows(0)
            tmpOCID1 = Convert.ToString(dr1("OCID1"))
            ViewState("Subject") = TIMS.Get_Subject(objconn, dr1("OCID1").ToString)
            'ViewState("CheckInDate")=TIMS.GET_CheckInDate(dr1("OCID1").ToString)
            'vsExamDate=TIMS.GET_ExamDate(dr1("OCID1").ToString)

            ViewState("CheckInDate") = ""
            If Convert.ToString(drX("CheckInDate")) <> "" Then
                If flag_ROC Then
                    ViewState("CheckInDate") = Common.FormatDate2Roc(drX("CheckInDate"))
                Else
                    ViewState("CheckInDate") = drX("CheckInDate")
                End If
            End If

            vsExamDate = ""
            If Convert.ToString(drX("ExamDate")) <> "" Then
                If flag_ROC Then
                    vsExamDate = Common.FormatDate2Roc(drX("ExamDate"))
                Else
                    vsExamDate = drX("ExamDate")
                End If
            End If

            ViewState("RIDOrgName") = TIMS.Get_OrgNameInputRID(dr1("RID").ToString, objconn)
            ViewState("isEmailFail") = Convert.ToString(dr1("isEmailFail"))          '已發送失敗E-MAIL
            '20090601 by Jimmy 取得e網報名審核成功「說明事項」內容 -- begin
            ViewState("EComment") = ""
            Dim tmpdt As DataTable = Nothing
            Dim tmpdr As DataRow = Nothing
            Dim tmpDistID As String = ""
            Dim tmpOrgID As String = ""
            Dim sqlstr As String = ""
            sqlstr = " SELECT DistID FROM Auth_Relship WHERE RID='" & Convert.ToString(dr1("RID")) & "' "
            tmpDistID = DbAccess.ExecuteScalar(sqlstr, objconn)
            If tmpDistID Is Nothing Then tmpDistID = ""
            sqlstr = " SELECT b.OrgID FROM Auth_Relship a JOIN Org_OrgInfo b ON a.orgid=b.orgid WHERE a.RID='" & Convert.ToString(dr1("RID")) & "' "
            tmpOrgID = DbAccess.ExecuteScalar(sqlstr, objconn)
            If tmpOrgID Is Nothing Then tmpOrgID = ""
            tmpdt = TIMS.Get_FinalEComment("Class", Convert.ToString(tmpOrgID), Convert.ToString(dr1("OCID1")), Convert.ToString(dr1("RID")), Convert.ToString(tmpDistID), Convert.ToString(sm.UserInfo.PlanID), Nothing, objconn) '20090618 by Jimmy 修改 依需求訓練機構(中心)都要for不同年度、計畫、轄區
            If Not tmpdt Is Nothing AndAlso tmpdt.Rows.Count > 0 Then
                tmpdr = tmpdt.Rows(0)
                ViewState("EComment") = Convert.ToString(tmpdr("eComment"))
            End If
        Catch ex As Exception
            Dim strErrmsg As String = ""
            strErrmsg &= "/* ex.ToString */" & vbCrLf
            strErrmsg &= ex.ToString & vbCrLf
            strErrmsg &= cst_errmsg1 & vbCrLf
            strErrmsg &= " and eSerNum=" & heSerNum.Value & vbCrLf
            strErrmsg &= " and eSETID=" & heSETID.Value & vbCrLf
            'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)

            Common.MessageBox(Me, cst_errmsg1)
            Exit Sub
        End Try
        Dim da As SqlDataAdapter = Nothing
        Dim dt As DataTable = Nothing
        sql = " SELECT * FROM Stud_EnterType2 WHERE eSerNum='" & heSerNum.Value & "' "
        dt = DbAccess.GetDataTable(sql, objconn)
        Session("_SearchStr") = ViewState("_SearchStr")
        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, cst_errmsg1)
            Exit Sub
        End If
        Using oConn As SqlConnection = DbAccess.GetConnection()
            Dim oTrans As SqlTransaction = DbAccess.BeginTrans(oConn)
            Try
                sql = " SELECT * FROM STUD_ENTERTYPE2 WHERE eSerNum='" & heSerNum.Value & "' "
                dt = DbAccess.GetDataTable(sql, da, oTrans)
                Dim dr As DataRow = dt.Rows(0)
                'signUpStatus 0:收件完成 1:報名成功 2:報名失敗 3:正取(Key_SelResult) 4:備取 5:未錄取
                dr("signUpStatus") = 2
                signUpMemo.Text = TIMS.ClearSQM(signUpMemo.Text)
                dr("signUpMemo") = Left(signUpMemo.Text.Trim, 150)
                ViewState("signUpMemo") = Replace(Left(signUpMemo.Text.Trim, 150), "'", "''")
                dr("ModifyAcct") = Convert.ToString(sm.UserInfo.UserID)
                dr("ModifyDate") = aNow 'Now '異動
                DbAccess.UpdateDataTable(dt, da, oTrans)
                DbAccess.CommitTrans(oTrans)
            Catch ex As Exception
                Dim strErrmsg As String = ""
                strErrmsg &= "/* ex.ToString */" & vbCrLf
                strErrmsg &= ex.ToString & vbCrLf
                strErrmsg &= cst_errmsg1 & vbCrLf
                strErrmsg &= " and eSerNum=" & heSerNum.Value & vbCrLf
                strErrmsg &= " and eSETID=" & heSETID.Value & vbCrLf
                strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
                'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                Call TIMS.WriteTraceLog(strErrmsg)

                DbAccess.RollbackTrans(oTrans)
                Call TIMS.CloseDbConn(oConn)
            End Try
            Call TIMS.CloseDbConn(oConn)
        End Using

        '20090601 by Jimmy 取得e網報名審核成功「說明事項」內容 -- end
        'ViewState("EmailSend") 為發送或不發送
        ViewState("EmailSend") = TIMS.CheckEmailSend(Me, "", "", objconn)
        Dim path3 As String = TIMS.Utl_GetConfigSet("from_emailaddress")
        Dim vpath3 As String = If(String.IsNullOrEmpty(path3), TIMS.Cst_SendMail3_from_emailaddress, path3)
        'Dim mail_msg As String=""
        If ViewState("Email").ToString <> "" And ViewState("EmailSend") And ViewState("isEmailFail") <> "Y" Then
            '修正錯誤的EMAIL增加發信成功率 BY AMU
            ViewState("Email") = TIMS.ChangeEmail(ViewState("Email"))
            '20090601 by Jimmy add 依需求將「甄試日期」改為「說明事項」，新增 eComment 參數
            Dim htSS As New Hashtable From {
                {"TPlanID", Convert.ToString(sm.UserInfo.TPlanID)},
                {"Stud_Name", Convert.ToString(ViewState("STUD_NAME"))},
                {"Subject", Convert.ToString(ViewState("Subject"))},
                {"ExamNo", Convert.ToString(ViewState("ExamNo"))},
                {"RelEnterDate", Convert.ToString(ViewState("RelEnterDate"))},
                {"ExamDate", Convert.ToString(vsExamDate)},
                {"CheckInDate", Convert.ToString(ViewState("CheckInDate"))},
                {"EComment", Convert.ToString(ViewState("EComment"))},
                {"Email", Convert.ToString(ViewState("Email"))},
                {"from_emailaddress", vpath3},
                {"signUpMemo", Convert.ToString(ViewState("signUpMemo"))},
                {"sRIDOrgName", Convert.ToString(ViewState("RIDOrgName"))},
                {"sType", TIMS.Cst_SendMail3_CheckedFalse}
            } 'htSS Hashtable() 
            Dim mail_msg As String = TIMS.SendMail3(htSS)
            If mail_msg <> "" Then Common.RespWrite(Me, "<script>alert('" & mail_msg & "');</script>")
            If mail_msg = "" Then
                Dim pms_u As New Hashtable From {{"eSerNum", Val(heSerNum.Value)}}
                Dim sqlstr_u As String = " UPDATE STUD_ENTERTYPE2 SET isEmailFail='Y' WHERE eSerNum=@eSerNum"
                DbAccess.ExecuteNonQuery(sqlstr_u, objconn, pms_u)
            End If
        End If

        Dim sMemo As String = String.Concat("&動作=審核失敗", "&NAME=", Name.Text)
        '寫入Log查詢 SubInsAccountLog1 (Auth_Accountlog)
        Call TIMS.SubInsAccountLog1(Me, Request("ID"), TIMS.cst_wm修改, TIMS.cst_wmdip2, tmpOCID1, sMemo)
        Common.RespWrite(Me, "<script>alert('儲存成功');location.href='SD_01_004.aspx?ID=" & Request("ID") & "';</script>")
    End Sub

    '回上一頁
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Session("_SearchStr") = ViewState("_SearchStr")
        'Response.Redirect("SD_01_004.aspx?ID=" & Request("ID"))
        Dim url1 As String = "SD_01_004.aspx?ID=" & Request("ID")
        Call TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    '回上一頁
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        'Response.Redirect("../03/SD_03_002_classver.aspx?ID=" & Request("ID") & "&OCID=" & vsOCID1val)
        Dim url1 As String = "../03/SD_03_002_classver.aspx?ID=" & Request("ID") & "&OCID=" & vsOCID1val
        Call TIMS.Utl_Redirect(Me, objconn, url1)
        'SD_03_002_classver=VIEW
    End Sub

    '近兩年參訓資料查詢
    Sub SHOW_STUD_HISTORY()
        If HiderrFlag.Value = cst_errFlag Then
            'Dim sErrMsg As String="資料有誤請重新查詢!" & vbCrLf
            'Common.MessageBox(Me, sErrMsg)
            Exit Sub
        End If
        Tablehistory3.Visible = True

        IDNOValue.Value = TIMS.ChangeIDNO(TIMS.ClearSQM(IDNOValue.Value))
        If IDNOValue.Value = "" Then Exit Sub

        Dim sParms As New Hashtable
        sParms.Add("IDNO", IDNOValue.Value)
        Dim sql As String = ""
        sql &= " SELECT b.IDNO,b.Name,b.Sex,b.Birthday" & vbCrLf
        sql &= " ,i.years + i.DistName + i.planname + i.seq planname" & vbCrLf
        sql &= " ,e.OrgName" & vbCrLf
        sql &= " ,ISNULL(g.TrainName, g.JobName) TMID" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(c.ClassCName,c.CyclType) ClassName" & vbCrLf
        sql &= " ,CASE WHEN a.TrainHours IS NULL THEN c.THours ELSE a.TrainHours END THours" & vbCrLf
        sql &= " ,c.STDate ,c.FTDate ,a.TrainHours" & vbCrLf
        sql &= " ,a.RejectTDate1 ,a.RejectTDate2" & vbCrLf
        sql &= " ,h.ExamName SkillName" & vbCrLf
        sql &= " ,a.StudStatus" & vbCrLf
        sql &= " ,j.PhoneD Tel ,j.ZipCode1 ,j.Address ,i.Years ,s.CJOB_NAME" & vbCrLf
        sql &= " ,dbo.fn_GET_PLAN_ONCLASS(pp.PlanID,pp.ComIDNO,pp.SeqNo,'WEEKTIME') WEEKS" & vbCrLf
        'sql &= " ,dbo.fn_GET_JOBSTATUS(sg3.IsGetJob,sg3.PUBLICRESCUE) JobStatus" & vbCrLf 'else '未填寫'
        sql &= " FROM Class_StudentsOfClass a" & vbCrLf
        sql &= " JOIN Stud_StudentInfo b ON a.SID=b.SID" & vbCrLf
        sql &= " JOIN Stud_SubData j ON j.SID=b.SID" & vbCrLf
        sql &= " JOIN Class_ClassInfo c ON a.OCID=c.OCID" & vbCrLf
        sql &= " JOIN Plan_PlanInfo pp ON c.planid=pp.planid AND pp.comidno=c.comidno AND pp.seqno=c.seqno" & vbCrLf
        sql &= " JOIN Auth_Relship d ON c.RID=d.RID" & vbCrLf
        sql &= " JOIN Org_OrgInfo e ON d.OrgID=e.OrgID" & vbCrLf
        sql &= " JOIN VIEW_PLAN i ON i.PlanID=c.PlanID" & vbCrLf
        sql &= " LEFT JOIN Key_TrainType g ON c.TMID=g.TMID" & vbCrLf
        sql &= " LEFT JOIN SHARE_CJOB s ON s.CJOB_UNKEY=c.CJOB_UNKEY" & vbCrLf
        sql &= " LEFT JOIN Stud_TechExam h ON a.SOCID=h.SOCID" & vbCrLf
        'sql &= " LEFT JOIN Stud_GetJobState3 sg3 ON sg3.SOCID=a.SOCID AND sg3.CPoint=1" & vbCrLf
        '近兩年參訓資料查詢
        sql &= " WHERE DATEPART(YEAR, c.STDate) >= (DATEPART(YEAR, GETDATE())-1)" & vbCrLf
        sql &= " AND b.IDNO=@IDNO" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, sParms)

        Tablehistory3.Visible = False
        If dt.Rows.Count > 0 Then
            'Dim dtSCJOB As DataTable=TIMS.Get_SHARECJOBdt(Me, objconn)
            'dtSCJOB=TIMS.Get_SHARECJOBdt(Me, objconn)
            Tablehistory3.Visible = True
            DataGrid3.DataSource = dt
            DataGrid3.DataBind()
        End If

    End Sub

    Sub SHOW_DG34(ByVal IDNOStr As String)
        'Public Shared Sub SHOW_DG34(ByRef sm As SessionModel, ByRef DGd34 As DataGrid, ByVal IDNOStr As String)
        Dim dr As DataRow = Nothing
        Dim dt As New DataTable
        TIMS.INIT_SPECdt(dt)

        IDNOStr = TIMS.ClearSQM(TIMS.ChangeIDNO(IDNOStr))
        'If IDNOStr="" Then Return dt 'Exit Sub
        If IDNOStr = "" Then Exit Sub

        Dim dt4B As DataTable = Nothing
        Dim flagGs4 As Boolean = True '查詢正常 true:正常 / false:異常
        Try
            dt4B = TIMS.GetTrainingListS(IDNOStr)
        Catch ex As Exception
            flagGs4 = False 'false:異常
        End Try

        If dt4B Is Nothing Then flagGs4 = False 'false:異常
        If dt4B IsNot Nothing Then
            If dt4B.Rows.Count = 0 Then flagGs4 = False 'false:異常
        End If
        'If Not flagGs4 Then Return dt 'Exit Sub
        If Not flagGs4 Then Exit Sub

        Dim i_SORT1 As Integer = 0
        Dim fff3 As String = "Years >= '" & CInt(sm.UserInfo.Years) - 1 & "'"
        dt4B.DefaultView.RowFilter = fff3
        dt4B = TIMS.dv2dt(dt4B.DefaultView)

        For Each dr4 As DataRow In dt4B.Rows 'For Each dr3 In dt3.Rows
            i_SORT1 += 1
            dr = dt.NewRow
            dt.Rows.Add(dr)
            dr("VSSORT") = i_SORT1 '改為序號顯示

            dr("IDNO") = TIMS.ChangeIDNO(dr4("IDNO")) '1.身分證號
            dr("Name") = Convert.ToString(dr4("NAME")) '1.姓名
            dr("Sex") = " " 'dr3("SEX") '性別
            dr("Birthday") = TIMS.Cdate3(dr4("Birthday")) '2.出生年月日
            dr("DistName") = Convert.ToString(dr4("DISTNAME")) '2.分署
            dr("Years") = Convert.ToString(dr4("YEARS")) '3.訓練年度
            dr("PlanName") = "<FONT color='Red'>" & Convert.ToString(dr4("PLANNAME")) & "</FONT>" '4.訓練計畫
            'dr("OrgName")=Convert.ToString(dr4("ORGNAME")) '5.訓練機構
            dr("OrgName") = Convert.ToString(dr4("ORGNAME")) '5.訓練機構
            dr("TMID") = Convert.ToString(dr4("TRAINNAME")) '6.訓練職類
            dr("CJOB_NAME") = Convert.ToString(dr4("CJOB_NAME")) '7.通俗職類
            dr("ClassName") = Convert.ToString(dr4("CLASSCNAME")) '8.班別名稱

            'THours: '9.受訓時數
            'TRound: '10.受訓期間
            dr("THours") = Convert.ToString(dr4("THOURS"))
            'dr("TRound")=Common.FormatDate(dr3("STDate")) & "<BR>|<BR>" & Common.FormatDate(dr3("FTDATE"))
            Dim strSTDate As String = "" 'If(flag_Roc, TIMS.cdate17(dr3("STDate")), Common.FormatDate(dr3("STDate")))
            Dim strFTDate As String = "" 'If(flag_Roc, TIMS.cdate17(dr3("FTDate")), Common.FormatDate(dr3("FTDate")))
            strSTDate = Convert.ToString(dr4("TRound")).Split("-")(0)
            strFTDate = Convert.ToString(dr4("TRound")).Split("-")(1)
            'If (flag_ROC) Then strSTDate=TIMS.cdate17(CDate(strSTDate))
            'If (flag_ROC) Then strFTDate=TIMS.cdate17(CDate(strFTDate))
            dr("TRound") = strSTDate & "<BR>|<BR>" & strFTDate
            'dr("SkillName")=dr3("ExamName") '11.技能檢定
            'dr("WEEKS")=dr3("WEEKS")  '12.上課時間
            '13.訓練狀態
            dr("TFlag") = TIMS.CHG_TFLAG(Convert.ToString(dr4("TFlag"))) 'Convert.ToString(dr4("TFlag")) 'dr4("StudStatus"))
            '補離退資訊
            'dr("JobStatus")=dr3("JobStatus") '15.訓後就業狀況
            ''參訓身分
            'If Key_Identity.Select("IdentityID='" & dr3("MIdentityID") & "'").Length > 0 Then
            '    dr("Ident")=Key_Identity.Select("IdentityID='" & dr3("MIdentityID") & "'")(0)("Name")
            'Else
            '    dr("Ident")="無身分別"
            'End If
            '電話1
            dr("Tel") = " " 'dr3("Tel").ToString
            '地址。
            dr("Address") = " " 'dr3("Address").ToString
            dr("WEEKS") = Convert.ToString(dr4("WEEKS")) '12.上課時間
            dr("MEMO1") = Convert.ToString(dr4("MEMO1"))
        Next

        DataGrid34.Visible = False
        If dt.Rows.Count > 0 Then
            Tablehistory3.Visible = True '共用dg3 table
            DataGrid34.Visible = True
            DataGrid34.DataSource = dt
            DataGrid34.DataBind()
        End If

        'DGd34.Visible=False
        'If dt.Rows.Count > 0 Then
        '    Tablehistory3.Visible=True '共用dg3 table
        '    DGd34.Visible=True
        '    DGd34.DataSource=dt
        '    DGd34.DataBind()
        'End If
        'Return dt
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                e.Item.Cells(0).Text = e.Item.ItemIndex + 1
        End Select
        'If e.Item.ItemType=ListItemType.Item Or e.Item.ItemType=ListItemType.AlternatingItem Then e.Item.Cells(0).Text=e.Item.ItemIndex + 1
    End Sub

    Private Sub Datagrid3_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid3.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                Const Cst_TMID As Integer = 3
                Const Cst_TMID_Name As String = "訓練業別"
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then e.Item.Cells(Cst_TMID).Text = Cst_TMID_Name '"訓練業別" '產投 訓練職類不採用 改為 訓練業別

            Case ListItemType.AlternatingItem, ListItemType.EditItem, ListItemType.Item, ListItemType.SelectedItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim TFlag As Label = e.Item.FindControl("TFlag")
                Dim THours As Label = e.Item.FindControl("THours")
                Dim TRound As Label = e.Item.FindControl("TRound")
                Dim cjob_Name As Label = e.Item.FindControl("cjob_Name")
                e.Item.Cells(0).Text = e.Item.ItemIndex + 1
                'cjob_Name.Text=TIMS.Get_CJOBNAME(dtSCJOB, Convert.ToString(drv("CJOB_UNKEY")))
                cjob_Name.Text = Convert.ToString(drv("CJOB_NAME"))

                Dim strSTDate As String = If(flag_ROC, TIMS.Cdate17(drv("STDate")), FormatDateTime(drv("STDate"), DateFormat.ShortDate))
                Dim strFTDate As String = If(flag_ROC, TIMS.Cdate17(drv("FTDate")), FormatDateTime(drv("FTDate"), DateFormat.ShortDate))
                Dim STUDSTATUS_N As String = TIMS.GET_STUDSTATUS_N(drv("StudStatus"))
                TFlag.Text = STUDSTATUS_N '"在訓"
                Select Case drv("StudStatus").ToString
                    Case "1"
                        'TFlag.Text="在訓"
                        THours.Text = drv("THours") '參訓時數，以 Class_StudentsOfClass 為優先資料顯示 Class_ClassInfo 為副
                        'TRound.Text=FormatDateTime(drv("STDate"), DateFormat.ShortDate) & "<BR>|<BR>" & FormatDateTime(drv("FTDate"), DateFormat.ShortDate)
                        TRound.Text = strSTDate & "<BR>|<BR>" & strFTDate
                    Case "2"
                        Dim vRejectTDate1 As String = TIMS.Cdate3(drv("RejectTDate1")) '有可能為空
                        Dim strRejectTDate1 As String = If(flag_ROC, TIMS.Cdate17(vRejectTDate1), vRejectTDate1)
                        'TFlag.Text="離訓"
                        THours.Text = "<FONT color='Red'>" & drv("TrainHours") & "</FONT>" '參訓時數，以 Class_StudentsOfClass 為主
                        'TRound.Text=FormatDateTime(drv("STDate"), DateFormat.ShortDate) & "<BR>|<BR>" & FormatDateTime(drv("RejectTDate1"), DateFormat.ShortDate)
                        TRound.Text = strSTDate & "<BR>|<BR>" & If(strRejectTDate1 <> "", strRejectTDate1, TIMS.cst_NODATAMsg93)
                        TIMS.Tooltip(THours, "離訓")
                    Case "3"
                        Dim vRejectTDate2 As String = TIMS.Cdate3(drv("RejectTDate2")) '有可能為空
                        Dim strRejectTDate2 As String = If(flag_ROC, TIMS.Cdate17(vRejectTDate2), vRejectTDate2)
                        'TFlag.Text="退訓"
                        THours.Text = "<FONT color='Red'>" & drv("TrainHours") & "</FONT>" '參訓時數，以 Class_StudentsOfClass 為主
                        'TRound.Text=FormatDateTime(drv("STDate"), DateFormat.ShortDate) & "<BR>|<BR>" & FormatDateTime(drv("RejectTDate2"), DateFormat.ShortDate)
                        TRound.Text = strSTDate & "<BR>|<BR>" & If(strRejectTDate2 <> "", strRejectTDate2, TIMS.cst_NODATAMsg93)
                        TIMS.Tooltip(THours, "退訓")
                    Case "4"
                        'TFlag.Text="續訓"
                        THours.Text = drv("THours") '參訓時數，以 Class_StudentsOfClass 為優先資料顯示 Class_ClassInfo 為副
                        'TRound.Text=FormatDateTime(drv("STDate"), DateFormat.ShortDate) & "<BR>|<BR>" & FormatDateTime(drv("FTDate"), DateFormat.ShortDate)
                        TRound.Text = strSTDate & "<BR>|<BR>" & strFTDate
                    Case "5"
                        'TFlag.Text="結訓"
                        THours.Text = drv("THours") '參訓時數，以 Class_StudentsOfClass 為優先資料顯示 Class_ClassInfo 為副
                        'TRound.Text=FormatDateTime(drv("STDate"), DateFormat.ShortDate) & "<BR>|<BR>" & FormatDateTime(drv("FTDate"), DateFormat.ShortDate)
                        TRound.Text = strSTDate & "<BR>|<BR>" & strFTDate
                End Select
        End Select
    End Sub

    Private Sub Datagrid4_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid4.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.EditItem, ListItemType.Item, ListItemType.SelectedItem
                e.Item.Cells(0).Text = e.Item.ItemIndex + 1
                Dim drv As DataRowView = e.Item.DataItem
                If flag_ROC Then
                    e.Item.Cells(cst_dg4_甄試日期).Text = TIMS.Cdate17(drv("examdate"))
                    e.Item.Cells(cst_dg4_開訓日期).Text = TIMS.Cdate17(drv("stdate"))
                    e.Item.Cells(cst_dg4_結訓日期).Text = TIMS.Cdate17(drv("ftdate"))
                Else
                    e.Item.Cells(cst_dg4_甄試日期).Text = drv("examdate")
                    e.Item.Cells(cst_dg4_開訓日期).Text = drv("stdate")
                    e.Item.Cells(cst_dg4_結訓日期).Text = drv("ftdate")
                End If
        End Select
    End Sub

    Private Sub DataGrid2bb_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2bb.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim Labsignno As Label = e.Item.FindControl("Labsignno")
                Dim Labdouble As Label = e.Item.FindControl("Labdouble")
                Dim labCLASSCNAME As Label = e.Item.FindControl("labCLASSCNAME")
                Dim Literal1 As Literal = e.Item.FindControl("Literal1") '日期-上課時間
                Dim Labstudstatus As Label = e.Item.FindControl("Labstudstatus") '訓練&lt;BR&gt;狀態
                Labsignno.Text = TIMS.Get_DGSeqNo(sender, e)

                Labdouble.Visible = False '重複參訓記錄。
                Literal1.Text = Get_TRAINDESCtb(CStr(drv("OCID")), gsPTDID, dtTrain, Val(Labsignno.Text))
                If Literal1.Text <> "" Then Labdouble.Visible = True 'False'重複參訓記錄。
                Labstudstatus.Text = "報名中"
                ff3 = "OCID=" & CStr(drv("OCID"))
                If dtStud.Select(ff3).Length > 0 Then Labstudstatus.Text = CStr(dtStud.Select(ff3)(0)("STUDSTATUS2"))
                labCLASSCNAME.Text = Convert.ToString(drv("CLASSCNAME"))

                Dim HrLk1 As HyperLink = e.Item.FindControl("HrLk1")
                ff3 = String.Format(TIMS.cst_ClassSearchUrl1, CStr(drv("OCID")), "1") '& "OCID=" & CStr(drv("OCID"))
                HrLk1.Target = "_blank"
                HrLk1.NavigateUrl = ff3
        End Select
    End Sub
End Class
