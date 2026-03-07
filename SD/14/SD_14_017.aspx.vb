Partial Class SD_14_017
    Inherits AuthBasePage

    Const cst_printFN1 As String = "SD_14_017b" 'SD_14_017b (SD_14_017b.jrxml)
    Const cst_printFN2 As String = "SD_14_017c" '2019版本

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

        If Not IsPostBack Then
            cCreate1()
        End If

        '署(局) 或 分署(中心)
        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

    End Sub

    Sub cCreate1()
        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID
        'Years.Value = sm.UserInfo.Years - 1911

        Dim vsOrgKind2 As String = TIMS.Get_OrgKind2(sm.UserInfo.OrgID, TIMS.c_ORGID, objconn)
        Common.SetListItem(rblOrgKind2, vsOrgKind2)

        rblOrgKind2 = TIMS.Get_RblOrgPlanKind(rblOrgKind2, objconn)
        Common.SetListItem(rblOrgKind2, "G")
        '依申請階段 '表示 (1：上半年、2：下半年、3：政策性產業)'AppStage = TIMS.Get_AppStage(AppStage)
        If tr_AppStage_TP28.Visible Then
            AppStage = If(sm.UserInfo.Years >= 2018, TIMS.Get_APPSTAGE2(AppStage), TIMS.Get_AppStage(AppStage))
            TIMS.SET_MY_APPSTAGE_LIST_VAL(Me, AppStage)
        End If

        Dim s_javascript_btn2 As String = ""
        Dim s_LevOrg As String = If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1")
        s_javascript_btn2 = String.Format("javascript:openOrg('../../Common/LevOrg{0}.aspx');", s_LevOrg)
        Button2.Attributes("onclick") = s_javascript_btn2

        '(V_DEPOT12)
        '課程分類 'KID12s 'SELECT * FROM KEY_BUSINESS A WHERE 1=1 AND A.DEPID='12'
        HidcblDepot12.Value = "0"
        cblDepot12 = TIMS.Get_KeyBusiness(cblDepot12, "12", objconn) '課程分類
        cblDepot12.Attributes("onclick") = "SelectAll('cblDepot12','HidcblDepot12');"
    End Sub

    Function checkData1(ByRef errMsg As String) As Boolean
        Dim rst As Boolean = True 'False
        STDate1.Text = TIMS.ClearSQM(STDate1.Text)
        STDate2.Text = TIMS.ClearSQM(STDate2.Text)
        AppliedDate1.Text = TIMS.ClearSQM(AppliedDate1.Text)
        AppliedDate2.Text = TIMS.ClearSQM(AppliedDate2.Text)
        If STDate1.Text <> "" Then
            If Not TIMS.IsDate1(STDate1.Text) Then errMsg &= "開訓期間 起始日期有誤" & vbCrLf
        End If
        If STDate2.Text <> "" Then
            If Not TIMS.IsDate1(STDate2.Text) Then errMsg &= "開訓期間 結束日期有誤" & vbCrLf
        End If
        If AppliedDate1.Text <> "" Then
            If Not TIMS.IsDate1(AppliedDate1.Text) Then errMsg &= "申請期間 起始日期有誤" & vbCrLf
        End If
        If AppliedDate2.Text <> "" Then
            If Not TIMS.IsDate1(AppliedDate2.Text) Then errMsg &= "申請期間 結束日期有誤" & vbCrLf
        End If
        If errMsg <> "" Then rst = False
        Return rst
    End Function

    ''' <summary>取得查詢參數</summary>
    ''' <returns></returns>
    Function GET_VSMyValue1() As String
        Dim vsDistID As String = ""
        Dim vsMyValue As String = ""
        vsMyValue &= "id=" & TIMS.ClearSQM(Request("ID"))
        vsMyValue &= "&Years=" & sm.UserInfo.Years 'USE(依年度)
        vsMyValue &= "&CYears=" & sm.UserInfo.Years - 1911 '顯示
        '28:產業人才投資方案
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Dim v_rblOrgKind2 As String = TIMS.GetListValue(rblOrgKind2)
            vsMyValue &= "&OrgKind2=" & v_rblOrgKind2 'rblOrgKind2.SelectedValue '(G.W.)
            Select Case v_rblOrgKind2 'rblOrgKind2.SelectedValue '產業人才投資計畫 種類
                Case "G"
                    vsMyValue &= "&OrgKind2Name=" & TIMS.Get_PName28(Me, v_rblOrgKind2, objconn) '顯示
                Case "W"
                    vsMyValue &= "&OrgKind2Name=" & TIMS.Get_PName28(Me, v_rblOrgKind2, objconn) '"提升勞工自主學習計畫" '顯示
            End Select
        Else
            If Hid_ORGKINDGW.Value <> "" Then
                vsMyValue &= "&OrgKind2=" & Hid_ORGKINDGW.Value 'rblOrgKind2.SelectedValue
            End If
            vsMyValue &= "&OrgKind2Name=" & TIMS.GetTPlanName(sm.UserInfo.TPlanID, objconn) '顯示
            'KindValue.Value = TIMS.GetTPlanName(sm.UserInfo.TPlanID)
        End If

        'Select Case rblReviewstage.SelectedValue
        '    Case 1
        '        vsMyValue &= "&Reviewstage=" & "上半年"
        '    Case 2
        '        vsMyValue &= "&Reviewstage=" & "下半年"
        'End Select
        'If AppStage.SelectedValue <> "" _
        '    AndAlso AppStage.SelectedValue <> "0" Then
        '    '資料為0不傳送
        '    vsMyValue &= "&AppStage=" & AppStage.SelectedValue
        'End If

        If tr_AppStage_TP28.Visible Then
            '依申請階段 
            Dim v_AppStage As String = TIMS.GetListValue(AppStage)
            If (v_AppStage <> "") Then Session(TIMS.SESS_DDL_APPSTAGE_VAL) = v_AppStage
            If v_AppStage <> "" AndAlso v_AppStage > "0" Then vsMyValue &= "&AppStage=" & v_AppStage
        End If

        If sm.UserInfo.RID = "A" Then
            If RIDValue.Value <> "A" Then
                'vsMyValue &= "&TPlanID=" & sm.UserInfo.TPlanID 'USE(依計畫)
                vsDistID = TIMS.Get_DistID_RID(RIDValue.Value, objconn)
                vsMyValue &= "&DistID=" & vsDistID 'USE(依轄區)
                vsMyValue &= "&DistName=" & TIMS.GET_DISTNAME(objconn, vsDistID)
                'vsMyValue &= "&RID=" & RIDValue.Value

                '===== 20180919 依承辦人需求,增加篩選[訓練機構]條件
                If center.Text.Trim.IndexOf("分署") = -1 And center.Text.Trim.IndexOf("署") = -1 Then vsMyValue &= "&RID=" & RIDValue.Value
                '======================================== END
            End If
        Else
            vsMyValue &= "&PlanID=" & sm.UserInfo.PlanID 'USE(依計畫)
            'vsMyValue &= "&TPlanID=" & sm.UserInfo.TPlanID 'USE(依計畫)
            vsMyValue &= "&DistID=" & sm.UserInfo.DistID 'USE(依轄區)
            vsMyValue &= "&DistName=" & TIMS.GET_DISTNAME(objconn, sm.UserInfo.DistID)
            'vsMyValue &= "&RID=" & RIDValue.Value

            '===== 20180919 依承辦人需求,增加篩選[訓練機構]條件
            If center.Text.Trim.IndexOf("分署") = -1 And center.Text.Trim.IndexOf("署") = -1 Then vsMyValue &= "&RID=" & RIDValue.Value
            '======================================== END
        End If

        vsMyValue &= "&STDate1=" & STDate1.Text 'USE(開訓日期1)
        vsMyValue &= "&STDate2=" & STDate2.Text 'USE(開訓日期2)
        vsMyValue &= "&AppliedDate1=" & AppliedDate1.Text 'USE(申請日期1)
        vsMyValue &= "&AppliedDate2=" & AppliedDate2.Text 'USE(申請日期2)
        Dim vsAppliedResult As String = "" '審核(未審、通過)
        Dim vsTransFlag As String = "" '轉班(已轉、未轉)
        '審核中
        '審核通過(未轉班)
        '審核通過(已轉班)
        For i As Integer = 0 To Me.cblClassStaus.Items.Count - 1
            If Me.cblClassStaus.Items(i).Selected Then
                Select Case Me.cblClassStaus.Items(i).Value
                    Case "1"
                        If vsAppliedResult.IndexOf("X") = -1 Then vsAppliedResult &= String.Concat(If(vsAppliedResult <> "", ",", ""), "X")
                        If vsTransFlag.IndexOf("N") = -1 Then vsTransFlag &= String.Concat(If(vsTransFlag <> "", ",", ""), "N")

                    Case "2"
                        If vsAppliedResult.IndexOf("Y") = -1 Then vsAppliedResult &= String.Concat(If(vsAppliedResult <> "", ",", ""), "Y")
                        If vsTransFlag.IndexOf("N") = -1 Then vsTransFlag &= String.Concat(If(vsTransFlag <> "", ",", ""), "N")
                    Case "3"
                        If vsAppliedResult.IndexOf("Y") = -1 Then vsAppliedResult &= String.Concat(If(vsAppliedResult <> "", ",", ""), "Y")
                        If vsTransFlag.IndexOf("Y") = -1 Then vsTransFlag &= String.Concat(If(vsTransFlag <> "", ",", ""), "Y")
                End Select
            End If
        Next

        vsMyValue &= "&AppliedResult=" & vsAppliedResult
        vsMyValue &= "&TransFlag=" & vsTransFlag

        '課程分類 '訓練課程分類 KID12
        Dim cblDepot12_Vals As String = TIMS.GetCblValue(cblDepot12)
        vsMyValue &= "&KID12s=" & cblDepot12_Vals

        'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "TPlanID28_2011", "SD_14_017b", vsMyValue)
        Return vsMyValue
    End Function

    ''' <summary>'列印</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btn_prt1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_prt1.Click
        Hid_ORGKINDGW.Value = ""
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
        If Len(RIDValue.Value) > 1 Then
            Dim drRR As DataRow = TIMS.Get_RID_DR(RIDValue.Value, objconn)
            If drRR IsNot Nothing Then
                Hid_ORGKINDGW.Value = Convert.ToString(drRR("ORGKINDGW"))
            End If
        End If

        Dim sErrMsg As String = ""
        Call checkData1(sErrMsg)
        If sErrMsg <> "" Then
            Common.MessageBox(Me, sErrMsg)
            Exit Sub
        End If

        '取得查詢參數
        Dim vsMyValue As String = GET_VSMyValue1()

        Dim prtFileName1 As String = cst_printFN1
        If sm.UserInfo.Years >= 2019 Then prtFileName1 = cst_printFN2
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, prtFileName1, vsMyValue)
    End Sub

    Function Search_SQL(ByRef o_parms As Hashtable) As String
        '取得查詢參數
        Dim vsMyValue As String = GET_VSMyValue1()

        Dim v_OrgKind2 As String = TIMS.GetMyValue(vsMyValue, "OrgKind2")
        Dim v_Years As String = TIMS.GetMyValue(vsMyValue, "Years")
        Dim v_DistID As String = TIMS.GetMyValue(vsMyValue, "DistID")
        Dim v_RID As String = TIMS.GetMyValue(vsMyValue, "RID")
        Dim v_PlanID As String = TIMS.GetMyValue(vsMyValue, "PlanID")
        Dim v_AppStage As String = TIMS.GetMyValue(vsMyValue, "AppStage")

        Dim v_STDate1 As String = TIMS.GetMyValue(vsMyValue, "STDate1")
        Dim v_STDate2 As String = TIMS.GetMyValue(vsMyValue, "STDate2")
        Dim v_AppliedDate1 As String = TIMS.GetMyValue(vsMyValue, "AppliedDate1")
        Dim v_AppliedDate2 As String = TIMS.GetMyValue(vsMyValue, "AppliedDate2")
        Dim v_AppliedResult As String = TIMS.GetMyValue(vsMyValue, "AppliedResult")
        Dim v_TransFlag As String = TIMS.GetMyValue(vsMyValue, "TransFlag")
        Dim v_KID12s As String = TIMS.GetMyValue(vsMyValue, "KID12s")

        o_parms.Clear()
        If v_OrgKind2 <> "" Then o_parms.Add("OrgKind2", v_OrgKind2)
        If v_Years <> "" Then o_parms.Add("Years", v_Years) 'sql &= " AND ip.Years = @Years" & vbCrLf
        If v_DistID <> "" Then o_parms.Add("DistID", v_DistID) 'sql &= " AND ip.DistID = #{DistID} " & vbCrLf
        If v_RID <> "" Then o_parms.Add("RID", v_RID) 'sql &= " AND pp.RID = #{RID}" & vbCrLf
        If v_PlanID <> "" Then o_parms.Add("PlanID", v_PlanID) 'sql &= " AND ip.PlanID = #{PlanID}" & vbCrLf
        If v_AppStage <> "" Then o_parms.Add("AppStage", v_AppStage) 'sql &= " AND pp.AppStage = #{AppStage}" & vbCrLf
        If v_STDate1 <> "" Then o_parms.Add("STDate1", v_STDate1) 'sql &= " AND pp.STDate >= CONVERT(DATETIME, #{STDate1})" & vbCrLf
        If v_STDate2 <> "" Then o_parms.Add("STDate2", v_STDate2) 'sql &= " AND pp.STDate <= CONVERT(DATETIME, #{STDate2})" & vbCrLf
        If v_AppliedDate1 <> "" Then o_parms.Add("AppliedDate1", v_AppliedDate1) 'sql &= " AND pp.AppliedDate >= CONVERT(DATETIME, #{AppliedDate1})" & vbCrLf
        If v_AppliedDate2 <> "" Then o_parms.Add("AppliedDate2", v_AppliedDate2) 'sql &= " AND pp.AppliedDate >= CONVERT(DATETIME, #{AppliedDate2})" & vbCrLf

        Dim sSql As String = ""
        sSql &= " SELECT ROW_NUMBER() OVER(ORDER BY vr.ORGNAME,pp.FIRSTSORT, pp.STDate) SEQNUM" & vbCrLf
        sSql &= " ,CONCAT(dbo.FN_CYEAR2(ip.YEARS),'年度',dbo.FN_GET_PLANKIND2(vr.ORGKIND2,ip.TPLANID),'審查彙整總表','(',ip.DistName,')') PLANKIND" & vbCrLf
        sSql &= " ,pp.PLANID ,pp.COMIDNO ,pp.SEQNO ,vr.ORGNAME" & vbCrLf
        sSql &= " ,dbo.FN_GET_CLASSCNAME(pp.CLASSNAME,pp.CYCLTYPE)" & vbCrLf
        sSql &= " +(CASE WHEN pp.RESULTBUTTON IN ('Y','R') THEN '(未送出)' ELSE '' END) CLASSCNAME" & vbCrLf
        sSql &= " ,pp.FIRSTSORT" & vbCrLf
        sSql &= " ,pp.PSNO28" & vbCrLf
        sSql &= " ,pp.THOURS ,pp.TNUM" & vbCrLf
        sSql &= " ,CASE WHEN pp.TNum IS NULL then FLOOR(pp.TotalCost/1) WHEN pp.TNum = 0 THEN FLOOR(pp.TotalCost/1) ELSE FLOOR(pp.TotalCost/(CASE WHEN ISNULL(pp.TNum, 1) <> 0 THEN ISNULL(pp.TNum, 1) ELSE 1 END)) END ONECOST" & vbCrLf
        sSql &= " ,pp.TOTALCOST ,pp.DEFGOVCOST" & vbCrLf
        sSql &= " ,FORMAT(pp.APPLIEDDATE,'yyyy/MM/dd') APPLIEDDATE" & vbCrLf
        sSql &= " ,FORMAT(pp.STDate,'yyyy/MM/dd') STDATE" & vbCrLf
        sSql &= " ,FORMAT(pp.FDDate,'yyyy/MM/dd') FTDATE" & vbCrLf
        'sql &= " ,ISNULL(pp.GCID3,pp.GCID2) GCID" & vbCrLf
        sSql &= " ,ISNULL(ig3.GCODE2,ig2.GCODE2) GOVCLASS" & vbCrLf
        sSql &= " ,ISNULL(ig3.CNAME,ig2.CNAME) GOVCLASSNAME" & vbCrLf
        sSql &= " ,kc.CodeID CCID" & vbCrLf
        sSql &= " ,kc.CCNAME" & vbCrLf
        sSql &= " ,pp.POINTYN" & vbCrLf
        sSql &= " ,vtt.TMID" & vbCrLf
        sSql &= " ,ISNULL(vtt.TrainName,vtt.JobName) TRAINNAME" & vbCrLf
        sSql &= " ,vr.ORGKIND2,vr.ORGKINDNAME" & vbCrLf ' --TRADENAME
        sSql &= " ,vp.CTNAME" & vbCrLf
        sSql &= " ,vr.CONTACTNAME" & vbCrLf
        sSql &= " ,vr.PHONE" & vbCrLf
        sSql &= " ,pp.APPSTAGE" & vbCrLf
        sSql &= " ,ISNULL(dd.KID12,ig3.GCODE31) KID12" & vbCrLf
        sSql &= " ,ISNULL(dd.KNAME12,ig3.PNAME) KNAME12" & vbCrLf
        sSql &= " ,(SELECT z2.CTNAME FROM VIEW_ZIPNAME z2 where z2.ZIPCODE=vr.ZIPCODE) CTNAME2" & vbCrLf
        sSql &= " ,(SELECT z2.CTNAME FROM VIEW_ZIPNAME z2 where z2.ZIPCODE=vr.ORGZIPCODE) CTNAME3" & vbCrLf
        sSql &= " ,(SELECT kt1.ORGTYPE FROM VIEW_ORGTYPE1 kt1 where kt1.OrgTypeID1=vr.ORGKIND1) ORGTYPE" & vbCrLf
        '線上送件。倘該班級係透過線上申辦送件並有送出至分署(【申辦狀態】非暫存)，則於此欄位顯示Y
        sSql &= " ,dbo.FN_GET_BIDCASEPI(pp.PlanID,pp.ComIDNO,pp.SeqNo,'Y') BIDCASEPI" & vbCrLf
        sSql &= " ,' ' MEMO1" & vbCrLf
        sSql &= " FROM dbo.PLAN_PLANINFO pp" & vbCrLf
        sSql &= " JOIN dbo.VIEW_RIDNAME vr ON vr.RID = pp.RID" & vbCrLf
        sSql &= " JOIN dbo.VIEW_TRAINTYPE vtt ON vtt.TMID = pp.TMID" & vbCrLf
        sSql &= " JOIN dbo.KEY_CLASSCATELOG kc ON kc.CCID = pp.ClassCate" & vbCrLf
        sSql &= " JOIN dbo.VIEW_PLAN ip ON ip.PlanID = pp.PlanID" & vbCrLf
        sSql &= " LEFT JOIN dbo.V_GOVCLASSCAST2 ig2 ON pp.GCID2 = ig2.GCID2" & vbCrLf
        sSql &= " LEFT JOIN dbo.V_GOVCLASSCAST3 ig3 ON pp.GCID3 = ig3.GCID3" & vbCrLf
        sSql &= " LEFT JOIN dbo.V_PLAN_DEPOT dd ON dd.PlanID=pp.PlanID AND dd.ComIDNO=pp.ComIDNO AND dd.SeqNo=pp.SeqNo" & vbCrLf
        sSql &= " LEFT JOIN dbo.VIEW_PLAN_CTNAME vp ON vp.PlanID=pp.PlanID AND vp.ComIDNO=pp.ComIDNO AND vp.SeqNo=pp.SeqNo" & vbCrLf
        sSql &= " WHERE pp.IsApprPaper='Y'" & vbCrLf

        If v_OrgKind2 <> "" Then sSql &= " AND vr.ORGKIND2 = @OrgKind2" & vbCrLf
        If v_Years <> "" Then sSql &= " AND ip.Years = @Years" & vbCrLf
        If v_DistID <> "" Then sSql &= " AND ip.DistID = @DistID " & vbCrLf
        If v_RID <> "" Then sSql &= " AND pp.RID = @RID" & vbCrLf
        If v_PlanID <> "" Then sSql &= " AND ip.PlanID = @PlanID" & vbCrLf
        If v_AppStage <> "" Then sSql &= " AND pp.AppStage = @AppStage" & vbCrLf
        If v_STDate1 <> "" Then sSql &= " AND pp.STDate >= CONVERT(DATETIME, @STDate1)" & vbCrLf
        If v_STDate2 <> "" Then sSql &= " AND pp.STDate <= CONVERT(DATETIME, @STDate2)" & vbCrLf
        If v_AppliedDate1 <> "" Then sSql &= " AND pp.AppliedDate >= CONVERT(DATETIME, @AppliedDate1)" & vbCrLf
        If v_AppliedDate2 <> "" Then sSql &= " AND pp.AppliedDate <= CONVERT(DATETIME, @AppliedDate2)" & vbCrLf

        Dim v_AppliedResult_in As String = TIMS.CombiSQM2IN(v_AppliedResult)
        If v_AppliedResult <> "" AndAlso v_AppliedResult_in <> "" Then sSql &= String.Concat(" AND ISNULL(CONVERT(VARCHAR,pp.AppliedResult),'X') IN (", v_AppliedResult_in, ") ", vbCrLf)

        Dim v_TransFlag_in As String = TIMS.CombiSQM2IN(v_TransFlag)
        If v_TransFlag <> "" AndAlso v_TransFlag_in <> "" Then sSql &= String.Concat(" AND pp.TransFlag IN (", v_TransFlag_in, ") ", vbCrLf)

        '課程分類 '訓練課程分類 KID12
        Dim v_KID12s_in As String = TIMS.CombiSQM2IN(v_KID12s)
        If v_KID12s <> "" AndAlso v_KID12s_in <> "" Then sSql &= String.Concat(" AND ISNULL(dd.KID12,ig3.GCODE31) IN (", v_KID12s_in, ")", vbCrLf)

        'test 測試環境測試
        'Dim flag_chktest As Boolean = If(TIMS.sUtl_ChkTest(), True, False) '(測試環境中)
        'If (flag_chktest) Then
        '    TIMS.writeLog(Me, String.Concat("##SD_14_017.aspx,", vbCrLf, ",Search_SQL sSql:", vbCrLf, sSql))
        '    TIMS.writeLog(Me, String.Concat("##SD_14_017.aspx,", vbCrLf, ",GetMyValue4:", vbCrLf, TIMS.GetMyValue4(o_parms)))
        'End If

        'sql &= " and ip.TPLANID='28'" & vbCrLf
        'sql &= " and ip.YEARS='2019'" & vbCrLf
        'sql &= " and ip.DISTID='001'" & vbCrLf
        'sql &= " and pp.STDATE <= '2019-03-01'" & vbCrLf
        sSql &= " ORDER BY vr.ORGNAME,pp.FIRSTSORT, pp.STDate" & vbCrLf
        Return sSql
    End Function

    ''' <summary> 匯出鈕 </summary>
    Sub Export1()
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) = -1 Then
            Common.MessageBox(Me, TIMS.cst_ErrorMsg17)
            Return 'Exit Sub
        End If

        Dim sErrMsg As String = ""
        Call checkData1(sErrMsg)
        If sErrMsg <> "" Then
            Common.MessageBox(Me, sErrMsg)
            Exit Sub
        End If

        'Dim vsMyValue As String = GET_VSMyValue1()
        Dim o_parms As Hashtable = New Hashtable()
        Dim sql As String = Search_SQL(o_parms)
        If sql = "" Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, o_parms)
        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        Dim iRows1 As Integer = 0

        'Review of aggregated tables
        Dim vrblOrgKind2 As String = TIMS.GetListValue(rblOrgKind2)
        Const cst_TitleS1 As String = "RVaggre"
        Dim strFilename1 As String = String.Concat(cst_TitleS1, vrblOrgKind2, TIMS.GetToday(objconn))
        Dim sPattern As String = ""
        Dim sColumn As String = ""
        Dim sTitle1 As String = ""
        'Select Case rblOrgKind2.SelectedValue
        '    Case "G"
        '        '產業人才投資計畫
        '        sTitle1 = CStr(sm.UserInfo.Years - 1911) & "年度產業人才投資方案(產業人才投資計畫)審查彙整總表"
        '    Case Else
        '        '提升勞工自主學習計畫
        '        sTitle1 = CStr(sm.UserInfo.Years - 1911) & "年度產業人才投資方案(提升勞工自主)審查彙整總表"
        'End Select
        Dim dr1 As DataRow = dt.Rows(0)
        sTitle1 = Convert.ToString(dr1("PLANKIND"))

        sPattern = "序號,訓練單位名稱,班別名稱,提案意願順序,班級課程流水號,訓練時數,訓練人數,每人訓練費用(元),每班總訓練費(元),每班總補助費(元),開訓日期,結訓日期"
        sPattern &= ",訓練業別編碼,訓練業別,訓練職能編碼,訓練職能,是否為學分班(Y/N),單位屬性,縣市別(辦訓地),聯絡人,聯絡電話,立案縣市,課程分類,統一編號,線上送件,備註"

        sColumn = "SEQNUM,ORGNAME,CLASSCNAME,FIRSTSORT,PSNO28,THOURS,TNUM,ONECOST,TOTALCOST,DEFGOVCOST,STDATE,FTDATE"
        sColumn &= ",GOVCLASS,GOVCLASSNAME,CCID,CCNAME,POINTYN,ORGTYPE,CTNAME,CONTACTNAME,PHONE,CTNAME3,KNAME12,COMIDNO,BIDCASEPI,MEMO1"

        Dim sPatternA() As String = Split(sPattern, ",")
        Dim sColumnA() As String = Split(sColumn, ",")
        Dim iColSpanCount As Integer = sColumnA.Length

        Dim parms As New Hashtable
        parms.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parms.Add("FileName", strFilename1)
        parms.Add("TitleName", TIMS.ClearSQM(sTitle1))
        parms.Add("TitleColSpanCnt", iColSpanCount)
        parms.Add("sPatternA", sPatternA)
        parms.Add("sColumnA", sColumnA)
        TIMS.Utl_Export(Me, dt, parms)
    End Sub

#Region "NO USE"
    'Sub exp1_old()
    '    Dim strFilename1 As String = ""
    '    Dim sTitle1 As String = ""
    '    Dim iColSpanCount As Integer = 0
    '    Dim sPatternA As String() = Nothing
    '    Dim sColumnA As String() = Nothing
    '    Dim dt As DataTable = Nothing

    '    Response.Clear()
    '    Response.ClearHeaders()
    '    Response.Buffer = True
    '    Response.Charset = "UTF-8" '"BIG5"
    '    'Response.ContentType = "Application/octet-stream"
    '    'Response.ContentType = "application/vnd.ms-excel"
    '    Response.ContentType = "application/ms-excel;charset=utf-8"
    '    Response.AddHeader("content-disposition", "attachment; filename=" & HttpUtility.UrlEncode(strFilename1, System.Text.Encoding.UTF8) & ".xls")
    '    'Response.ContentEncoding = System.Text.Encoding.GetEncoding("Big5")
    '    Response.ContentEncoding = System.Text.Encoding.GetEncoding("UTF-8")

    '    Common.RespWrite(Me, "<html>")
    '    Common.RespWrite(Me, "<head>")
    '    'Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=BIG5>")
    '    Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=UTF-8>")
    '    '<head><meta http-equiv='Content-Type' content='text/html; charset=utf-8'></head>
    '    '套CSS值
    '    Common.RespWrite(Me, "<style>")
    '    Common.RespWrite(Me, "td{mso-number-format:""\@"";}")
    '    Common.RespWrite(Me, ".noDecFormat{mso-number-format:""0"";}")
    '    'mso-number-format:"0" 
    '    Common.RespWrite(Me, "</style>")
    '    Common.RespWrite(Me, "</head>")

    '    Common.RespWrite(Me, "<body>")
    '    Common.RespWrite(Me, "<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")
    '    'Common.RespWrite(Me, "<tr>")

    '    '標題抬頭
    '    Dim ExportStr As String = "" '建立輸出文字
    '    ExportStr = "<tr>"
    '    ExportStr &= "<td colspan='" & iColSpanCount & "' align='center'>" & sTitle1 & "</td>" '& vbTab
    '    ExportStr &= "</tr>" & vbCrLf
    '    Common.RespWrite(Me, ExportStr)

    '    ExportStr = "<tr>"
    '    For i As Integer = 0 To sPatternA.Length - 1
    '        ExportStr &= "<td>" & sPatternA(i) & "</td>" '& vbTab
    '    Next
    '    ExportStr &= "</tr>" & vbCrLf

    '    Common.RespWrite(Me, ExportStr)
    '    '建立資料面
    '    Dim iNum As Integer = 0
    '    For Each dr As DataRow In dt.DefaultView.Table.Rows
    '        iNum += 1
    '        ExportStr = "<tr>"
    '        For i As Integer = 0 To sColumnA.Length - 1
    '            'Select Case CStr(sColumnA(i))
    '            '    Case "Phone1", "Phone2", "CellPhone"
    '            '        ExportStr &= "<td class=""noDecFormat"">" & Convert.ToString(dr(sColumnA(i))) & "</td>" '& vbTab
    '            '    Case Else
    '            '        ExportStr &= "<td>" & Convert.ToString(dr(sColumnA(i))) & "</td>" '& vbTab
    '            'End Select
    '            ExportStr &= "<td>" & Convert.ToString(dr(sColumnA(i))) & "</td>" '& vbTab
    '        Next
    '        ExportStr &= "</tr>" & vbCrLf
    '        Common.RespWrite(Me, ExportStr)
    '    Next
    '    Common.RespWrite(Me, "</table>")
    '    Common.RespWrite(Me, "</body>")
    '    Response.End()
    'End Sub
#End Region

    '匯出
    Protected Sub BtnExp1_Click(sender As Object, e As EventArgs) Handles BtnExp1.Click
        Call Export1()
    End Sub
End Class