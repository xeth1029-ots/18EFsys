Partial Class TC_01_004_InsertPlan
    Inherits AuthBasePage

    '(TIMS專用非產投) 'TC_01_004_InsertPlan.aspx
    '(產投)TIMS.Utl_Redirect1(Me, "TC_01_004_BusAdd.aspx?ID=" & Request("ID") & "&STDate=" & vsSTDate)
    'strScript1 += "location.href='TC_01_004_add.aspx?ProcessType=PlanUpdate&ID='+document.getElementById('Re_ID').value;" + vbCrLf
    Const cst_temp_classinfo As String = "temp_classinfo" 'Session(cst_temp_classinfo)

    Dim rqRID As String = ""
    Dim Re_Planid As String = ""
    Dim Re_ComIDNO As String = ""
    Dim Re_SeqNO As String = ""

    Dim iPlanCnt As Integer = 0
    Dim drPlaninfo As DataRow = Nothing 'DataRow 
    'drPlaninfo = Get_PlanInfoDataRow(Re_Planid, Re_ComIDNO, Re_SeqNO, rqRID, iPlanCnt)
    'Dim oTest_flag As Boolean = False '(正式)
    'If TIMS.sUtl_ChkTest() Then oTest_flag = True '測試
    'If Not oTest_flag Then '(正式)
    Dim flag_oTestEnv As Boolean = False '測試

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        If TIMS.sUtl_ChkTest() Then flag_oTestEnv = True '測試
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then LabTMID.Text = "訓練業別"

#Region "(No Use)"

        'If sm.UserInfo.FunDt Is Nothing Then
        '    Common.RespWrite(Me, "<script>alert('Session過期');</script>")
        '    Common.RespWrite(Me, "<script>top.location.href='../../logout.aspx';</script>")
        'Else
        '   'Dim FunDr As DataRow
        '    Dim FunDt As DataTable = sm.UserInfo.FunDt
        '    Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & Request("ID") & "'")
        '    Re_ID.Value = Request("ID")
        '    If FunDrArray.Length = 0 Then
        '        Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
        '        Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
        '    Else
        '        FunDr = FunDrArray(0)
        '        change.Disabled = True
        '        If FunDr("Adds") = "1" Then
        '            change.Disabled = False
        '        End If
        '    End If
        'End If

#End Region

        Re_Planid = TIMS.ClearSQM(Request("planid"))
        Re_ComIDNO = TIMS.ClearSQM(Request("ComIDNO"))
        Re_SeqNO = TIMS.ClearSQM(Request("SeqNO"))
        rqRID = TIMS.ClearSQM(Request("RID")) '選擇訓練單位
        If rqRID = "" Then rqRID = sm.UserInfo.RID '選擇訓練單位

        If flag_oTestEnv Then '--測試
            Dim drPCS As DataRow = TIMS.GetPCSDate(Re_Planid, Re_ComIDNO, Re_SeqNO, objconn)
            If Not drPCS Is Nothing Then rqRID = drPCS("RID").ToString() '選擇訓練單位
        End If

        'out: iPlanCnt
        drPlaninfo = TIMS.Get_PlanInfoDataRow(objconn, Re_Planid, Re_ComIDNO, Re_SeqNO, rqRID, "Y", iPlanCnt)
        If drPlaninfo Is Nothing Then
            Common.MessageBox(Me, "計畫資料有誤，請重新選擇!!")
            Return 'Exit Sub
        ElseIf iPlanCnt <> 1 Then
            drPlaninfo = Nothing
            Common.MessageBox(Me, String.Concat("計畫資料有誤，請重新選擇!!s." & iPlanCnt))
            Return 'Exit Sub
        End If
        'If Not oTest_flag Then '(正式) 'End If

        If Not IsPostBack Then
            cCreate1()
        End If

        cCreateET2()

        'If Not Page.IsPostBack Then If clsid.Value = "" Then change.Disabled = True
        If clsid.Value = "" Then change.Disabled = True

        back.Attributes("onclick") = "history.go(-1);"

        '確認機構是否為黑名單
        Dim vsMsg2 As String = ""
        If Chk_OrgBlackList(vsMsg2) Then
            change.Disabled = True
            TIMS.Tooltip(change, vsMsg2)
            Dim vsStrScript As String = $"<script>alert('{vsMsg2}');</script>"
            Page.RegisterStartupScript("", vsStrScript)
        End If
    End Sub

    Sub cCreate1()
        If drPlaninfo Is Nothing Then Return

        '登入者檢查
        Hid_ComIDNO.Value = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, objconn)
        Dim iBlackType As Integer = TIMS.Chk_OrgBlackType(Me, objconn)
        If TIMS.Check_OrgBlackList2(Me, Hid_ComIDNO.Value, iBlackType, objconn) Then
            Select Case iBlackType
                Case 1, 2, 3
                    'Errmsg &= "於處分日期起的期間，已審核通過的班級不可進行轉班作業。"
                    Common.MessageBox(Me, "於處分日期起的期間，已審核通過的班級不可進行轉班作業。")
                    Return 'Exit Sub '有錯誤訊息 'Return False '不可儲存
            End Select
        End If
        'Hid_ComIDNO.Value = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, objconn)
        'Dim iBlackType As Integer = TIMS.Chk_OrgBlackType(Me, objconn)
        '轉入班級者檢查
        If TIMS.Check_OrgBlackList2(Me, Convert.ToString(drPlaninfo("COMIDNO")), iBlackType, objconn) Then
            Select Case iBlackType
                Case 1, 2, 3
                    'Errmsg &= "於處分日期起的期間，已審核通過的班級不可進行轉班作業。"
                    Common.MessageBox(Me, "於處分日期起的期間，已審核通過的班級不可進行轉班作業。")
                    Return 'Exit Sub '有錯誤訊息 'Return False '不可儲存
            End Select
        End If
        'If oTest_flag Then '測試'(正式)
        '    Common.MessageBox(Me, "於處分日期起的期間，已審核通過的班級不可進行轉班作業。")
        '    return 'Exit Sub '有錯誤訊息 'Return False '不可儲存
        'End If
    End Sub

    Sub cCreateET2()
        If drPlaninfo Is Nothing Then Return

        Name.Text = Convert.ToString(drPlaninfo("OrgName"))
        Plan_Name.Text = Convert.ToString(drPlaninfo("PlanName"))
        TrainID.Text = Convert.ToString(drPlaninfo("TrainID"))
        Train_Name.Text = Convert.ToString(drPlaninfo("TrainName"))
        cjobValue.Text = Convert.ToString(drPlaninfo("CJOB_NO"))
        txtCJOB_NAME.Text = Convert.ToString(drPlaninfo("CJOB_NAME"))

        Dim parms As New Hashtable From {{"PLANID", Re_Planid}, {"COMIDNO", Re_ComIDNO}, {"SEQNO", Re_SeqNO}}
        Dim sql As String = ""
        sql &= " SELECT ClassName,TNum,THours,STDate,FDDate,CyclType,LevelType" & vbCrLf
        sql &= " FROM dbo.PLAN_PLANINFO WITH(NOLOCK)" & vbCrLf
        sql &= " WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO" & vbCrLf
        '(正式) 'sql += "   AND TransFlag = 'N'" & vbCrLf
        If Not flag_oTestEnv Then sql &= " AND TransFlag='N'" & vbCrLf
        sql &= " AND AppliedResult='Y'" & vbCrLf
        Dim dtPlan1 As DataTable = DbAccess.GetDataTable(sql, objconn, parms)

        Dim aryrow() As String = {"班別名稱", "開結訓起迄日", "訓練人數", "訓練時數", "期別", "班別代碼"}
        Dim cell As New HtmlTableCell
        Dim row As New HtmlTableRow
        'Dim i, j As Integer

        For i As Integer = 0 To aryrow.Length - 1
            cell = New HtmlTableCell
            cell.InnerText = aryrow(i)
            row.Cells.Add(cell)
            'row.Style("Color") = "#000000"  'edit，by:20181024
            row.Style("Color") = "#FFFFFF"   'edit，by:20181024
        Next
        row.Align = "center"
        'row.BgColor = "#ffcccc"  'edit，by:20181024
        row.Attributes.Add("class", "head_navy")  'edit，by:20181024
        search_tbl.Rows.Add(row)

        If dtPlan1.Rows.Count = 0 Then
            row = New HtmlTableRow
            cell = New HtmlTableCell
            cell.InnerText = "目前沒有計畫可以轉入!!"
            cell.ColSpan = search_tbl.Rows(0).Cells.Count
            row.Cells.Add(cell)
            row.Align = "center"
            search_tbl.Rows.Add(row)
            Return 'Exit Sub
        End If
        If dtPlan1.Rows.Count <> 1 Then
            row = New HtmlTableRow
            cell = New HtmlTableCell
            cell.InnerText = "計畫資料有誤，請重新選擇!!"
            cell.ColSpan = search_tbl.Rows(0).Cells.Count
            row.Cells.Add(cell)
            row.Align = "center"
            search_tbl.Rows.Add(row)
            Return 'Exit Sub
        End If

        Dim j As Integer = 1
        For Each dr As DataRow In dtPlan1.Rows
            'Dim classid As String
            '若開班轉入日期為 小於、等於 開訓前三天  則不能進行開班轉入動作
            If chk_period(dr("STDate")) = False Then
                change.Attributes.Add("title", "轉入日期未大於開訓前三天無法執行開班轉入!")
                change.Disabled = True
            End If

            row = New HtmlTableRow
            cell = New HtmlTableCell
            cell.InnerText = Convert.ToString(dr("ClassName"))
            row.Cells.Add(cell)

            cell = New HtmlTableCell
            cell.InnerHtml = DbAccess.GetRocDateValue(dr("STDate")) + "~" + DbAccess.GetRocDateValue(dr("FDDate"))
            row.Cells.Add(cell)

            cell = New HtmlTableCell
            cell.InnerHtml = Convert.ToString(dr("TNum"))
            row.Cells.Add(cell)

            cell = New HtmlTableCell
            cell.InnerHtml = Convert.ToString(dr("THours"))
            row.Cells.Add(cell)

            Dim v_CyclType As String = TIMS.ClearSQM(dr("CyclType"))
            If v_CyclType = "" Then v_CyclType = TIMS.cst_Default_CyclType
            v_CyclType = TIMS.FmtCyclType(v_CyclType)
            cell = New HtmlTableCell
            cell.InnerHtml = v_CyclType
            row.Cells.Add(cell)

            'cell = New HtmlTableCell
            'cell.InnerHtml = dr("LevelType")
            'row.Cells.Add(cell)

            Dim TrainIDTMID As String = ""
            If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                TrainIDTMID = drPlaninfo("TMID") '產業人才投資方案DATAROW '97年度後改成 JOBID 的 TMID 而97年度前為 TrainID 的 TMID
            Else
                TrainIDTMID = drPlaninfo("TrainID") '產業人才投資方案DATAROW '97年度後改成 JOBID 的 TMID 而97年度前為 TrainID 的 TMID
            End If

            cell = New HtmlTableCell
            'cell.InnerHtml = "<input type=button value=挑選代碼 onclick=""wopen('TC_01_004_classid.aspx','班別代碼',300,300,1);"" >"
            Dim str_myValue As String = ""
            str_myValue &= String.Format("&TMID={0}", TrainIDTMID)
            If sm.UserInfo.LID = 0 Then str_myValue &= String.Format("&PlanID={0}", Convert.ToString(drPlaninfo("PlanID")))

            cell.InnerHtml = "<input type=button value=挑選代碼 onclick=""wopen('TC_01_004_classid.aspx?pp=cc" & str_myValue & "' ,'班別代碼',560,560,1);"" class=""asp_Export_M"" >"
            row.Cells.Add(cell)

            row.Align = "center"
            search_tbl.Rows.Add(row)
            If j Mod 2 = 0 Then row.BgColor = "#FFFDC7"
            j = j + 1
        Next

    End Sub


    '機構黑名單內容(訓練單位處分功能)
    Function Chk_OrgBlackList(ByRef Errmsg As String) As Boolean
        Dim rst As Boolean = False
        Errmsg = ""
        Dim vsComIDNO As String = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, objconn)
        If TIMS.Check_OrgBlackList(Me, vsComIDNO, objconn) Then
            rst = True
            Errmsg = sm.UserInfo.OrgName & "，已列入處分名單!!"
            isBlack.Value = "Y"
            orgname.Value = sm.UserInfo.OrgName
            'btnAdd.Visible = False
            'Button8.Visible = False
        End If
        Return rst
    End Function

    '若開班轉入日期為 小於、等於 開訓前三天  則不能進行開班轉入動作
    Function chk_period(ByVal sdate As String) As Boolean
        Dim bln As Boolean = True
        'Dim d1 As DateTime = CDate(sdate)
        If DateDiff(DateInterval.Day, Date.Today, CDate(sdate)) <= 3 Then bln = False
        Return bln
    End Function

    ''' <summary>
    ''' 轉入資料(SAVE) PLAN_PLANINFO CLASS_CLASSINFO
    ''' </summary>
    Sub SAVE_CHANGEDATA1()
        If drPlaninfo Is Nothing Then
            Common.MessageBox(Me, "計畫資料有誤，請重新選擇!!")
            Return 'Exit Sub
        End If

        'Dim parms As Hashtable = New Hashtable()
        'Dim Errmsg As String = ""
        '登入者檢查
        Hid_ComIDNO.Value = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, objconn)
        Dim iBlackType As Integer = TIMS.Chk_OrgBlackType(Me, objconn)
        If TIMS.Check_OrgBlackList2(Me, Hid_ComIDNO.Value, iBlackType, objconn) Then
            Select Case iBlackType
                Case 1, 2, 3
                    'Errmsg &= "於處分日期起的期間，已審核通過的班級不可進行轉班作業。"
                    Common.MessageBox(Me, "於處分日期起的期間，已審核通過的班級不可進行轉班作業。")
                    Return 'Exit Sub '有錯誤訊息 'Return False '不可儲存
            End Select
        End If

        'Hid_ComIDNO.Value = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, objconn)
        'Dim iBlackType As Integer = TIMS.Chk_OrgBlackType(Me, objconn)

        '轉入班級者檢查
        If TIMS.Check_OrgBlackList2(Me, Convert.ToString(drPlaninfo("COMIDNO")), iBlackType, objconn) Then
            Select Case iBlackType
                Case 1, 2, 3
                    'Errmsg &= "於處分日期起的期間，已審核通過的班級不可進行轉班作業。"
                    Common.MessageBox(Me, "於處分日期起的期間，已審核通過的班級不可進行轉班作業。")
                    Return 'Exit Sub '有錯誤訊息 'Return False '不可儲存
            End Select
        End If

        'Dim sql9 As String = ""
        'Dim dr9 As DataRow = Nothing
        'Dim strScript2 As String = ""
        'PLAN_PLANINFO
        'Dim ilevel As Integer = 1 '預設為1班
        'ilevel = 1
        'If Convert.ToString(drPlaninfo("ClassCount")) <> "" Then ilevel = CInt(drPlaninfo("ClassCount"))

        Hid_RID1.Value = Convert.ToString(drPlaninfo("RID")).Substring(0, 1)

        clsid.Value = TIMS.ClearSQM(clsid.Value)
        If clsid.Value = "" Then
            Dim strScript1 As String = "<script language=""javascript"">alert('請選擇班級代碼!!');</script>"
            Page.RegisterStartupScript("", strScript1)
            change.Disabled = True
            Return 'Exit Sub '(未選擇離開)
        End If

        'If clsid.Value <> "" Then
        Dim PMSck As New Hashtable() From {{"CLSID", clsid.Value}, {"PlanID", drPlaninfo("PlanID")}, {"RID", drPlaninfo("RID")}}
        Dim check_sql As String = ""
        check_sql &= " SELECT concat('(',dbo.FN_CLASSID2(cc.CLSID),')',dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE)) CLASSCNAME" & vbCrLf
        check_sql &= " FROM dbo.CLASS_CLASSINFO cc" & vbCrLf
        check_sql &= " WHERE cc.CLSID=@CLSID" & vbCrLf '依班別代碼(重複)
        check_sql &= " AND cc.PlanID=@PlanID AND cc.RID=@RID" & vbCrLf 'PlanID,機構
        If $"{drPlaninfo("CyclType")}" <> "" Then
            check_sql &= " AND cc.CyclType=@CyclType" & vbCrLf '期別(重複)
            PMSck.Add("CyclType", drPlaninfo("CyclType"))
        Else
            check_sql &= " AND cc.CyclType IS NULL" & vbCrLf '期別(重複)
        End If
        Dim dr9ck As DataRow = DbAccess.GetOneRow(check_sql, objconn, PMSck)

        Dim blnChkIsDouble As Boolean = If(dr9ck IsNot Nothing, True, False) '重複

        If blnChkIsDouble Then '重複
            Dim strScript2 As String = String.Concat("轉入班級資料 班別代碼與期別重複!!", vbCrLf, dr9ck("classcname"))
            Common.MessageBox(Me, strScript2)
            'Me.change.Disabled = True 'Exit Sub '(有重複離開)
            If Not flag_oTestEnv Then '(正式)
                change.Disabled = True
                Return 'Exit Sub '(有重複離開)
            End If
        End If

        '沒有重複,新增資料
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 AndAlso sm.UserInfo.Years >= 2013 Then
            '2013年產投啟用 
            Dim sPMS9 As New Hashtable From {{"PlanID", drPlaninfo("PlanID")}, {"ComIDNO", drPlaninfo("ComIDNO")}, {"SeqNo", drPlaninfo("SeqNo")}}
            Dim sql9 As String = " SELECT * FROM PLAN_DEPOT WHERE PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNo=@SeqNo"
            Dim dr9 As DataRow = DbAccess.GetOneRow(sql9, objconn, sPMS9)
            If dr9 Is Nothing Then '沒資料
                Dim ERRMSG2 As String = "轉入失敗,請聯絡承辦人執行重點產業審核確認,才可進行開班轉入動作!!"
                Dim strScript2 As String = ""
                strScript2 += "<script language=""javascript"">" + vbCrLf
                strScript2 += "alert('轉入失敗,請聯絡承辦人執行重點產業審核確認,才可進行開班轉入動作!!');" + vbCrLf
                strScript2 += "</script>"
                Page.RegisterStartupScript("", strScript2)
                Return 'Exit Sub
            End If
        End If

        Dim sPMS9c As New Hashtable From {{"CLSID", clsid.Value}}
        Dim sql9c As String = " SELECT * FROM ID_CLASS WHERE CLSID =@CLSID"
        Dim dr9c As DataRow = DbAccess.GetOneRow(sql9c, objconn, sPMS9c)
        If dr9c Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Return 'Exit Sub
        End If

        If IsDBNull(dr9c("CJOB_UNKEY")) = True Then '如果有CJOB_UNKEY是NULL的
            Dim strScript2 As String = ""
            strScript2 += "<script language=""javascript"">" + vbCrLf
            strScript2 += "alert(' 轉入失敗,請聯絡承辦人設定此班別代碼的通俗職類資料,才可進行開班轉入動作!!');" + vbCrLf
            strScript2 += "</script>"
            Page.RegisterStartupScript("", strScript2)
            Return 'Exit Sub
        Else '如果有CJOB_UNKEY的值
            Dim uPMS9 As New Hashtable From {{"CJOB_UNKEY", dr9c("CJOB_UNKEY")}, {"PLANID", Re_Planid}, {"COMIDNO", Re_ComIDNO}, {"SEQNO", Re_SeqNO}}
            Dim sql9u As String = ""
            sql9u &= " UPDATE PLAN_PLANINFO "
            sql9u &= " SET CJOB_UNKEY =@CJOB_UNKEY"
            sql9u &= " WHERE PLANID = @PLANID AND COMIDNO = @COMIDNO AND SEQNO = @SEQNO"
            DbAccess.ExecuteNonQuery(sql9u, objconn, uPMS9)
        End If

        Session(cst_temp_classinfo) = Nothing
        Dim sql As String = ""
        Dim dtX As DataTable = Nothing
        Dim relship As String = "" 'Auth_Relship.relship
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Dim parms As New Hashtable()
            '企訓專用(產投)
            sql = " SELECT RELSHIP FROM AUTH_RELSHIP WHERE RID=@RID"
            parms.Clear()
            parms.Add("RID", Convert.ToString(drPlaninfo("RID")))
            dtX = DbAccess.GetDataTable(sql, objconn, parms)
            If dtX.Rows.Count <> 1 Then
                Common.MessageBox(Me, "業務權限異常，不可轉入!!")
                change.Disabled = True
                Return 'Exit Sub
            End If

            sql = "SELECT CLASSENAME FROM ID_CLASS WHERE CLSID=@CLSID "
            parms.Clear()
            parms.Add("CLSID", clsid.Value)
            dtX = DbAccess.GetDataTable(sql, objconn, parms)
            If dtX.Rows.Count <> 1 Then
                Common.MessageBox(Me, "轉入用班級代碼異常，不可轉入!!")
                change.Disabled = True
                Return 'Exit Sub
            End If

            sql = " SELECT 'x' FROM PLAN_TRAINDESC WHERE PLANID = @PLANID AND COMIDNO = @COMIDNO AND SEQNO = @SEQNO ORDER BY PTDID"
            parms.Clear()
            parms.Add("PLANID", Convert.ToString(drPlaninfo("PlanID")))
            parms.Add("COMIDNO", Convert.ToString(drPlaninfo("comidno")))
            parms.Add("SEQNO", Convert.ToString(drPlaninfo("SeqNO")))
            dtX = DbAccess.GetDataTable(sql, objconn, parms)
            If dtX.Rows.Count = 0 Then
                Common.MessageBox(Me, "計畫訓練內容簡介無資料，不可轉入!!")
                change.Disabled = True
                Return 'Exit Sub
            End If

        Else
            Dim parms As New Hashtable()
            '一般計畫檢核。
            sql = " SELECT RELSHIP FROM AUTH_RELSHIP WHERE RID=@RID"
            parms.Clear()
            parms.Add("RID", Convert.ToString(drPlaninfo("RID")))
            dtX = DbAccess.GetDataTable(sql, objconn, parms)
            If dtX.Rows.Count <> 1 Then
                Common.MessageBox(Me, "業務權限異常，不可轉入!!")
                change.Disabled = True
                Return 'Exit Sub
            End If

            sql = " SELECT CLASSENAME FROM ID_CLASS WHERE CLSID = @CLSID "
            parms.Clear()
            parms.Add("CLSID", clsid.Value)
            dtX = DbAccess.GetDataTable(sql, objconn, parms)
            If dtX.Rows.Count <> 1 Then
                Common.MessageBox(Me, "轉入用班級代碼異常，不可轉入!!")
                change.Disabled = True
                Return 'Exit Sub
            End If

            sql = "SELECT 'x' FROM PLAN_TRAINDESC WHERE PLANID = @PLANID AND COMIDNO = @COMIDNO AND SEQNO = @SEQNO ORDER BY PTDID "
            parms.Clear()
            parms.Add("PlanID", Convert.ToString(drPlaninfo("PlanID")))
            parms.Add("COMIDNO", Convert.ToString(drPlaninfo("comidno")))
            parms.Add("SEQNO", Convert.ToString(drPlaninfo("SeqNO")))
            dtX = DbAccess.GetDataTable(sql, objconn, parms)
            If dtX.Rows.Count = 0 Then
                Common.MessageBox(Me, "計畫訓練內容簡介無資料，不可轉入!!")
                change.Disabled = True
                Return 'Exit Sub
            End If
        End If

        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '產投轉班專用。'企訓專用(產投)
            Dim dt As DataTable 'CLASS_CLASSINFO
            Dim dr1 As DataRow 'PLAN_PLANINFO
            Dim TPeriod As String = ""
            Dim ClassEngName As String = "" 'ID_Class.ClassEName

            Dim pms_cd As New Hashtable() From {{"CLSID", clsid.Value}}
            Dim sql_cd As String = " SELECT CLASSENAME FROM ID_CLASS WHERE CLSID =@CLSID"
            ClassEngName = Convert.ToString(DbAccess.ExecuteScalar(sql_cd, objconn, pms_cd))

            Dim pms_p1 As New Hashtable() From {{"PlanID", drPlaninfo("PlanID")}, {"ComIDNO", drPlaninfo("ComIDNO")}, {"SeqNo", drPlaninfo("SeqNo")}}
            Dim sql_p1 As String = " SELECT * FROM PLAN_PLANINFO WHERE PlanID =@PlanID AND ComIDNO =@ComIDNO AND SeqNo =@SeqNo"
            dr1 = DbAccess.GetOneRow(sql_p1, objconn, pms_p1)

            'dr1應該有資料
            Dim pms_r1 As New Hashtable() From {{"RID", dr1("RID")}}
            Dim sql_r1 As String = " SELECT RELSHIP FROM AUTH_RELSHIP WHERE RID=@RID"
            relship = Convert.ToString(DbAccess.ExecuteScalar(sql_r1, objconn, pms_r1))

            sql = " SELECT * FROM CLASS_CLASSINFO WHERE 1<>1 "
            dt = DbAccess.GetDataTable(sql, objconn)
            Dim dr As DataRow = Nothing
            dr = dt.NewRow 'CLASS_CLASSINFO
            dt.Rows.Add(dr)
            dr("CLSID") = clsid.Value
            dr("PlanID") = dr1("PlanID")
            dr("Years") = Right(dr1("PlanYear"), 2)

            Dim vCyclType As String = TIMS.ClearSQM(dr1("CyclType"))
            If vCyclType = "" Then vCyclType = TIMS.cst_Default_CyclType
            vCyclType = TIMS.FmtCyclType(vCyclType)
            dr("CyclType") = If(vCyclType <> "", vCyclType, Convert.DBNull)
            dr("ClassNum") = "01" 'vCyclType

            dr("RID") = dr1("RID")
            dr("ClassCName") = dr1("ClassName")
            'dr("CJOB_UNKEY") = dr9("CJOB_UNKEY")  '通俗職類
            dr("ClassEngName") = If(ClassEngName <> "", ClassEngName, Convert.DBNull)
            dr("Content") = dr1("Content")
            'dr("Purpose") = "一、學科：" & dr1("PurScience") & vbCrLf & "二、術科：" & dr1("PurTech")
            '2007/9/26 修改成將訓練目標帶入即可--Charles
            dr("Purpose") = dr1("PurScience")
            dr("TMID") = dr1("TMID")
            dr("STDate") = dr1("STDate")
            dr("FTDate") = dr1("FDDate")
            TPeriod = TIMS.Get_Plan_VerReport(dr1("PlanID"), dr1("ComIDNO"), dr1("SeqNo"), "TPeriod", objconn)
            If TPeriod = "" Then dr("TPeriod") = Convert.DBNull Else dr("TPeriod") = TPeriod

            dr("TaddressZip") = dr1("TaddressZip")
            dr("TaddressZIP6W") = dr1("TaddressZIP6W")
            dr("TAddress") = dr1("TAddress")
            dr("THours") = dr1("THours")
            dr("TNum") = dr1("TNum")
            dr("ADVANCE") = dr1("ADVANCE") '訓練課程類型 ADVANCE
            dr("Relship") = relship 'DbAccess.ExecuteScalar("select relship from Auth_Relship where rid='" & dr1("RID") & "'", objconn)
            dr("ComIDNO") = dr1("ComIDNO")
            dr("SeqNO") = dr1("SeqNO")

            Dim vsSTDate As String = Common.FormatDate(dr("STDate"))
            'CLASS_CLASSINFO
            Session(cst_temp_classinfo) = dt
            Call TIMS.CloseDbConn(objconn)
            'Response.Redirect("TC_01_004_BusAdd.aspx?ID=" & Request("ID"))
            TIMS.Utl_Redirect1(Me, "TC_01_004_BusAdd.aspx?ID=" & Request("ID") & "&STDate=" & vsSTDate)

        Else
            'TIMS專用
            'Dim sqldr_class As DataRow
            Dim sqldr As DataRow = Nothing
            'Dim sqldr_new As DataRow = Nothing
            Dim daPlanInfo As SqlDataAdapter = Nothing
            Dim dtPlanInfo As DataTable = Nothing
            Dim dtTRAINDESC As DataTable = Nothing
            Dim ClassEngName As String = "" 'ID_Class.ClassEName

            Dim pms_r1 As New Hashtable() From {{"RID", drPlaninfo("RID")}}
            Dim sql_r1 As String = " SELECT RELSHIP FROM AUTH_RELSHIP WHERE RID=@RID"
            relship = Convert.ToString(DbAccess.ExecuteScalar(sql_r1, objconn, pms_r1))

            Dim pms_cd As New Hashtable() From {{"CLSID", clsid.Value}}
            Dim sql_cd As String = " SELECT CLASSENAME FROM ID_CLASS WHERE CLSID =@CLSID"
            ClassEngName = Convert.ToString(DbAccess.ExecuteScalar(sql_cd, objconn, pms_cd))

            Dim pms_tc As New Hashtable() From {{"PlanID", drPlaninfo("PlanID")}, {"ComIDNO", drPlaninfo("ComIDNO")}, {"SeqNo", drPlaninfo("SeqNo")}}
            Dim sql_tc As String = " SELECT PCONT,PNAME FROM PLAN_TRAINDESC WHERE PLANID = @PLANID AND COMIDNO = @COMIDNO AND SEQNO = @SEQNO ORDER BY PTDID"
            dtTRAINDESC = DbAccess.GetDataTable(sql_tc, objconn, pms_tc)

            Dim class_PName As String = ""
            For Each drTra As DataRow In dtTRAINDESC.Rows
                Dim sPName As String = TIMS.ClearSQM(drTra("PName"))
                class_PName &= String.Concat(If(class_PName <> "", ",", ""), sPName)
            Next

            Dim pp_years As String = CInt(drPlaninfo("PlanYear"))
            Try
                'Dim objTrans As SqlTransaction
                Dim sqlAdapter As SqlDataAdapter = Nothing
                Dim sqlTable As New DataTable
                sql = " SELECT * FROM CLASS_CLASSINFO WHERE 1<>1 "
                sqlTable = DbAccess.GetDataTable(sql, sqlAdapter, objconn)

                sqldr = sqlTable.NewRow 'CLASS_CLASSINFO
                sqldr("Relship") = relship
                'sqldr("Content") = class_cont
                sqldr("Content") = class_PName   '2007/11/19 修改成將訓練內容簡介的課程單元帶入--Charles
                sqldr("ClassEngName") = ClassEngName
                sqldr("Years") = pp_years.Substring(2) '012
                sqldr("PlanID") = drPlaninfo("PlanID")
                sqldr("ComIDNO") = drPlaninfo("ComIDNO")
                sqldr("SeqNO") = drPlaninfo("SeqNO")
                sqldr("RID") = drPlaninfo("RID")
                sqldr("TMID") = drPlaninfo("TMID")
                'CJOB_UNKEY
                sqldr("CJOB_UNKEY") = If(Convert.ToString(drPlaninfo("CJOB_UNKEY")) <> "", drPlaninfo("CJOB_UNKEY"), dr9c("CJOB_UNKEY")) '通俗職類
                sqldr("CLSID") = clsid.Value
                sqldr("ClassCName") = drPlaninfo("ClassName")

                Dim vCyclType As String = TIMS.ClearSQM(drPlaninfo("CyclType"))
                If vCyclType = "" Then vCyclType = TIMS.cst_Default_CyclType
                vCyclType = TIMS.FmtCyclType(vCyclType)
                sqldr("CyclType") = If(vCyclType <> "", vCyclType, Convert.DBNull)

                sqldr("ClassNum") = 1 'vCyclType '班數
                sqldr("ADVANCE") = drPlaninfo("ADVANCE") '訓練課程類型 ADVANCE
                sqldr("TNum") = drPlaninfo("TNum")
                sqldr("THours") = drPlaninfo("THours")
                sqldr("STDate") = drPlaninfo("STDate")
                sqldr("FTDate") = drPlaninfo("FDDate")
                'SELECT SENTERDATE,FENTERDATE,EXAMDATE,ExamPeriod FROM PLAN_PLANINFO WHERE ROWNUM  <=10
                'SELECT SENTERDATE,FENTERDATE,EXAMDATE,FENTERDATE2,ExamPeriod FROM CLASS_CLASSINFO  WHERE ROWNUM  <=10
                sqldr("SENTERDATE") = drPlaninfo("SENTERDATE")
                sqldr("FENTERDATE") = drPlaninfo("FENTERDATE")
                sqldr("EXAMDATE") = drPlaninfo("EXAMDATE")
                sqldr("ExamPeriod") = drPlaninfo("ExamPeriod")
                Dim sFENTERDATE As String = Convert.ToString(drPlaninfo("FENTERDATE"))
                Dim sEXAMDATE As String = Convert.ToString(drPlaninfo("EXAMDATE"))
                Dim SS1 As String = ""
                TIMS.SetMyValue(SS1, "RID1", Hid_RID1.Value) : TIMS.SetMyValue(SS1, "TPlanID", sm.UserInfo.TPlanID)
                Dim sFENTERDATE2 As String = TIMS.GET_FENTERDATE2(SS1, sFENTERDATE, sEXAMDATE, objconn)
                If sFENTERDATE2 <> "" Then sqldr("FENTERDATE2") = CDate(sFENTERDATE2) 'TIMS.GET_FENTERDATE2()
                sqldr("CheckInDate") = drPlaninfo("CheckInDate")
                '2005/8/11新增轉入訓練地點--Melody
                sqldr("TaddressZip") = drPlaninfo("TaddressZip")
                sqldr("TaddressZIP6W") = drPlaninfo("TaddressZIP6W")
                sqldr("TAddress") = drPlaninfo("TAddress")
                '2005/8/12新增轉入課程目標--Melody，2007/9/26 修改成將訓練目標帶入即可--Charles
                'sqldr("Purpose") = "一、學科：" & drPlaninfo("PurScience") & "二、術科：" & drPlaninfo("PurTech")
                sqldr("Purpose") = drPlaninfo("PurScience")
                sqldr("NotOpen") = "N"
                sqldr("IsCalculate") = "N"
                sqldr("IsApplic") = "N"
                '班級英文名稱
                sqldr("CLASSENGNAME") = drPlaninfo("CLASSENGNAME")
                '訓練時段'取得鍵值-訓練時段
                sqldr("TPERIOD") = drPlaninfo("TPERIOD")
                sqldr("NOTE3") = drPlaninfo("NOTE3")
                '「訓練期限」
                sqldr("TDEADLINE") = drPlaninfo("TDEADLINE")
                '導師名稱
                sqldr("CTName") = drPlaninfo("CTName")
                'EADDRESSZIP,EADDRESS,EADDRESSZIP6W
                sqldr("EADDRESSZIP") = drPlaninfo("EADDRESSZIP")
                sqldr("EADDRESSZIP6W") = drPlaninfo("EADDRESSZIP6W")
                sqldr("EADDRESS") = drPlaninfo("EADDRESS")

                sqldr("ModifyAcct") = sm.UserInfo.UserID
                sqldr("ModifyDate") = Now()
                '2005/5/31-新增班級單元-Melody
                'If TPlan_str = "15" Then sqldr("Class_Unit") = drPlaninfo("Class_Unit") '學習券
                If sm.UserInfo.TPlanID = "15" Then sqldr("Class_Unit") = drPlaninfo("Class_Unit") '學習券
                sqlTable.Rows.Add(sqldr)
                'sqldr_new("TransFlag") = "Y"
                'sqldr_new("ModifyAcct") = sm.UserInfo.UserID
                'sqldr_new("ModifyDate") = Now()
                'CLASS_CLASSINFO
                Session(cst_temp_classinfo) = sqlTable

                Dim strScript1 As String
                strScript1 = "<script language=""javascript"">" + vbCrLf
                'strScript1 += "alert('班級轉入成功!!');" + vbCrLf
                strScript1 += "location.href='TC_01_004_add.aspx?ProcessType=PlanUpdate&ID='+document.getElementById('Re_ID').value;" + vbCrLf
                strScript1 += "</script>"
                Page.RegisterStartupScript("", strScript1)

            Catch ex As Exception
                'objTrans.Rollback()
                Common.MessageBox(Page, "班級轉入失敗!!")
                Throw ex
                'Finally ' objconn.Close()
            End Try
        End If
    End Sub

    '轉入鈕
    Private Sub Change_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles change.ServerClick
        'SELECT ClassCount ,COUNT(1) CNT,MIN(MODIFYDATE),MAX(MODIFYDATE) FROM PLAN_PLANINFO GROUP  BY ClassCount  ORDER BY 1
        If drPlaninfo Is Nothing Then
            Common.MessageBox(Me, "計畫資料有誤，請重新選擇!!")
            Return 'Exit Sub
        End If
        Call SAVE_CHANGEDATA1()
    End Sub

    '回開班設定
    Private Sub Return_class_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles return_class.ServerClick
        Call TIMS.CloseDbConn(objconn)
        TIMS.Utl_Redirect1(Me, "TC_01_004.aspx?ID=" & Request("ID") & "")
    End Sub

#Region "(No Use)"

    'Function Get_Plan_VerReport(ByVal PlanID As String, ByVal ComIDNO As String, ByVal SeqNo As String, _
    '    ByVal FieldName As String, ByVal tConn As SqlConnection) As String
    '    Dim Rst As String = ""
    '    Dim dt As DataTable
    '    Dim sql As String = ""
    '    sql = ""
    '    sql &= " select * FROM Plan_VerReport   "
    '    sql &= " WHERE PlanID=" & PlanID & " "
    '    sql &= " AND ComIDNO='" & ComIDNO & "' "
    '    sql &= " AND SeqNo=" & SeqNo & " "
    '    dt = DbAccess.GetDataTable(sql, tConn)
    '    If dt.Rows.Count > 0 Then
    '        Rst = Convert.ToString(dt.Rows(0)(FieldName))
    '    End If
    '    Return Rst
    'End Function

#End Region
End Class