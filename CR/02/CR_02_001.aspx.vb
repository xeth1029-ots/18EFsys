Public Class CR_02_001
    Inherits AuthBasePage 'System.Web.UI.Page

    'OJT-22063001
    'iType 1:查詢用 11:匯出(1) 12:匯出(2)
    'Const cst_iType_查詢1 As Integer = 1
    'Const cst_iType_匯出1 As Integer = 11
    'Const cst_iType_匯出2 As Integer = 12
    'Const cst_SCORELEVEL_A As String = "A"
    'Const cst_SCORELEVEL_B As String = "B"
    'Const cst_SCORELEVEL_C As String = "C"
    'Const cst_SCORELEVEL_D As String = "D"

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)

        If Not IsPostBack Then
            CCreate1()
        End If

        '委訓
        'Select Case sm.UserInfo.LID
        '    Case 2
        '        Button2.Visible = False
        '    Case Else
        '        'Button2.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg1.aspx');"
        '        If sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1 Then
        '            '署(局) 或 分署(中心)
        '            TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        '            If HistoryRID.Rows.Count <> 0 Then
        '                center.Attributes("onclick") = "showObj('HistoryList2');"
        '                center.Style("CURSOR") = "hand"
        '            End If
        '            Button2.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg.aspx');"
        '        End If
        'End Select
    End Sub

    Sub CCreate1()
        PanelSch1.Visible = True

        msg1.Text = ""

        'ddlYEARS_SCH = TIMS.GetSyear(ddlYEARS_SCH)
        'Common.SetListItem(ddlYEARS_SCH, sm.UserInfo.Years)

        ddlAPPSTAGE_SCH = TIMS.Get_APPSTAGE2(ddlAPPSTAGE_SCH)
        Common.SetListItem(ddlAPPSTAGE_SCH, "1")

        '訓練機構
        'center.Text = sm.UserInfo.OrgName
        'RIDValue.Value = sm.UserInfo.RID

        '計畫  產業人才投資計畫/提升勞工自主學習計畫
        Dim vsOrgKind2 As String = TIMS.Get_OrgKind2(sm.UserInfo.OrgID, TIMS.c_ORGID, objconn)
        If (vsOrgKind2 = "") Then vsOrgKind2 = "G"
        rblOrgKind2 = TIMS.Get_RblSearchPlan(rblOrgKind2, objconn, False)
        'Common.SetListItem(rblOrgKind2, "G")
        Common.SetListItem(rblOrgKind2, vsOrgKind2)

        '開訓日期～ 

        '跨區/ 轄區提案 不區分跨區提案單位轄區提案單位 '跨區/轄區提案 'D>不區分 C>跨區提案單位 J>轄區提案單位

        '初審建議結論 --Y 通過、N 不通過、P 調整後通過
        'ddlST1RESULT = TIMS.Get_ST1RESULT(ddlST1RESULT)
        'Result 初審建議結論 / 審查結果 - -Y 通過、N 不通過、P 調整後通過
        'ddlRESULT = TIMS.Get_ST1RESULT(ddlRESULT)
    End Sub

    ''' <summary> 查詢SQL DataTable </summary>
    ''' <returns></returns>
    Function SEARCH_DATA1_dt() As DataTable
        Dim dt As DataTable = Nothing

        Dim v_rblOrgKind2 As String = TIMS.GetListValue(rblOrgKind2)
        Dim v_APPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH) '申請階段
        If v_APPSTAGE_SCH = "" Then
            msg1.Text = TIMS.cst_NODATAMsg2
            Return dt
        End If

        Dim parms As New Hashtable From {{"TPLANID", sm.UserInfo.TPlanID}, {"YEARS", sm.UserInfo.Years}, {"APPSTAGE", v_APPSTAGE_SCH}, {"ORGKIND2", v_rblOrgKind2}}
        Dim sql As String = ""
        '先 機構統編範圍 確認 WC1
        sql &= " WITH WC1 AS ( SELECT cc.COMIDNO" & vbCrLf
        sql &= " FROM dbo.VIEW2B cc" & vbCrLf
        sql &= " JOIN dbo.PLAN_STAFFOPIN pf WITH(NOLOCK) ON pf.PSNO28=cc.PSNO28 AND pf.CURESULT='Y'" & vbCrLf '首頁>>課程審查>>二階審查>>核班結果裡【核班結果】：通過的班級
        sql &= " WHERE (cc.RESULTBUTTON IS NULL OR cc.APPLIEDRESULT='Y')" & vbCrLf '審核送出(已送審)
        sql &= " AND cc.PVR_ISAPPRPAPER='Y'" & vbCrLf '正式
        sql &= " AND cc.DATANOTSENT IS NULL" & vbCrLf '未檢送資料註記(排除有勾選)
        sql &= " AND cc.TPLANID=@TPLANID AND cc.YEARS=@YEARS AND cc.APPSTAGE=@APPSTAGE AND cc.ORGKIND2=@ORGKIND2 )" & vbCrLf

        '跨4區確認-機構統編 WORG1
        sql &= " ,WORG1 AS ( SELECT CASE WHEN dbo.FN_GET_CROSSDIST4(@YEARS,oo.COMIDNO,@APPSTAGE)>3 THEN 'Y' END CROSSDIST4" & vbCrLf
        sql &= " ,dbo.FN_SCORING2_UPLIMIT(oo.COMIDNO,@TPLANID,@YEARS,@APPSTAGE,@ORGKIND2) UPLIMIT" & vbCrLf '可核配上限,等級額度核配上限
        sql &= " ,dbo.FN_SCORING2_GRADE(oo.COMIDNO,@TPLANID,@YEARS,@APPSTAGE) GRADE" & vbCrLf '跨4區確認等級 跨區等級
        sql &= " ,oo.COMIDNO,oo.ORGID" & vbCrLf
        sql &= " FROM ORG_ORGINFO oo WITH(NOLOCK)" & vbCrLf
        sql &= " WHERE oo.COMIDNO IN (SELECT COMIDNO FROM WC1) )" & vbCrLf

        '班級過濾 WC2/WORG1
        sql &= " ,WC2 AS ( SELECT cc.OCID,cc.ORGPLANNAME" & vbCrLf
        sql &= " ,cc.TPLANID,cc.YEARS,cc.ORGKIND2" & vbCrLf
        sql &= " ,cc.PLANID,cc.COMIDNO,cc.SEQNO" & vbCrLf
        sql &= " ,cc.ORGNAME,cc.CTNAME,cc.ORGZIPCODE" & vbCrLf
        sql &= " ,cc.PSNO28 ,cc.RID,cc.APPSTAGE" & vbCrLf
        sql &= " ,cc.STDATE,cc.FTDATE" & vbCrLf
        sql &= " ,wo.UPLIMIT ,wo.GRADE" & vbCrLf
        sql &= " ,CC.DISTID,CC.DISTNAME" & vbCrLf
        sql &= " ,CC.CLASSCNAME2" & vbCrLf
        sql &= " ,CC.DEFGOVCOST,CC.DEFSTDCOST,CC.TOTALCOST" & vbCrLf
        sql &= " ,cc.FIRSTSORT,cc.ICAPNUM" & vbCrLf
        sql &= " ,o2.IMPLEVEL_1 LEVEL1" & vbCrLf
        sql &= " ,o2.RLEVEL_2 RLEVEL2" & vbCrLf '審查計分表等級(複審)
        sql &= " ,pf.CURESULT" & vbCrLf

        sql &= " FROM dbo.VIEW2B cc" & vbCrLf
        sql &= " JOIN dbo.PLAN_STAFFOPIN pf WITH(NOLOCK) ON pf.PSNO28=cc.PSNO28 AND pf.CURESULT='Y'" & vbCrLf '首頁>>課程審查>>二階審查>>核班結果裡【核班結果】：通過的班級
        sql &= " LEFT JOIN dbo.ORG_SCORING2 o2 WITH(NOLOCK) ON o2.OSID2=cc.OSID2" & vbCrLf
        '跨4區確認-機構統編 WORG1
        sql &= " JOIN WORG1 wo ON wo.COMIDNO=cc.COMIDNO AND wo.CROSSDIST4='Y' AND LEN(wo.GRADE)=3" & vbCrLf
        sql &= " WHERE (cc.RESULTBUTTON IS NULL OR cc.APPLIEDRESULT='Y')" & vbCrLf '審核送出(已送審)
        sql &= " AND cc.PVR_ISAPPRPAPER='Y'" & vbCrLf '正式
        sql &= " AND cc.DATANOTSENT IS NULL" & vbCrLf '未檢送資料註記(排除有勾選)

        sql &= " AND cc.TPLANID=@TPLANID AND cc.YEARS=@YEARS AND cc.APPSTAGE=@APPSTAGE AND cc.ORGKIND2=@ORGKIND2 )" & vbCrLf

        '5分署初核累計 WC3/WC2
        sql &= " ,WC3 AS ( SELECT COMIDNO, SUM(DEFGOVCOST) ST1GRANDTOTAL" & vbCrLf '五分署初核累計
        sql &= " ,COUNT(1) ST1CNT" & vbCrLf '五分署初核累計
        sql &= " FROM WC2" & vbCrLf
        sql &= " GROUP BY COMIDNO )" & vbCrLf

        '最後查詢 WC2/WC3
        sql &= " SELECT cc.ORGPLANNAME" & vbCrLf '計畫別
        sql &= " ,cc.PLANID,cc.COMIDNO,cc.SEQNO" & vbCrLf
        sql &= " ,cc.ORGNAME,cc.CTNAME" & vbCrLf '訓練單位名稱 '(班級)縣市
        sql &= " ,(SELECT x.CTNAME FROM dbo.VIEW_ZIPNAME x WHERE x.ZIPCODE=cc.ORGZIPCODE) ORGCTNAME" & vbCrLf '訓練單位名稱 '立案縣市
        sql &= " ,cc.PSNO28 ,cc.RID,cc.APPSTAGE" & vbCrLf
        sql &= " ,FORMAT(cc.STDATE,'yyyy/MM/dd') STDATE" & vbCrLf
        sql &= " ,FORMAT(cc.FTDATE,'yyyy/MM/dd') FTDATE" & vbCrLf
        sql &= " ,cc.UPLIMIT ,cc.GRADE" & vbCrLf '可核配上限 '等級
        sql &= " ,CC.DISTID,CC.DISTNAME" & vbCrLf '分署別
        sql &= " ,CC.CLASSCNAME2" & vbCrLf '課程名稱(含期別)

        ',總補助費(元)(以訓練費用之80%估算)
        sql &= " ,CC.DEFGOVCOST,CC.DEFSTDCOST,CC.TOTALCOST" & vbCrLf

        sql &= " ,wc3.ST1CNT" & vbCrLf '五分署初核累計數
        sql &= " ,wc3.ST1GRANDTOTAL" & vbCrLf ',五分署初核累計" & vbCrLf
        sql &= " ,(wc3.ST1GRANDTOTAL-cc.UPLIMIT) ST1DIFFAMOUNT" & vbCrLf ',差額(初核-上限)" & vbCrLf
        sql &= " ,case when (wc3.ST1GRANDTOTAL-cc.UPLIMIT)>=0 then 'Y' else 'N' end ST1OVERUPLIMIT" & vbCrLf ',是否超過上限" & vbCrLf

        sql &= " ,NULL ADJGRANDTOTAL" & vbCrLf ',調整後累計" & vbCrLf
        sql &= " ,NULL ADJDIFFAMOUNT" & vbCrLf ',差額(調整後-上限)" & vbCrLf
        sql &= " ,NULL ADJOVERUPLIMIT" & vbCrLf ',是否超過上限" & vbCrLf
        ''5+2產業創新計畫--,5+2產業,'新南向政策,'台灣AI行動計畫,'數位國家創新經濟發展方案,'國家資通安全發展方案,前瞻基礎建設計畫
        sql &= " ,dd.D20KNAME1,dd.D20KNAME6,dd.D20KNAME2,dd.D20KNAME3,dd.D20KNAME4,dd.D20KNAME5" & vbCrLf
        '亞洲矽谷,'重點產業,'台灣AI行動計畫,'智慧國家方案,'國家人才競爭力躍升方案,'新南向政策,
        sql &= " ,dd.D25KNAME1,dd.D25KNAME2,dd.D25KNAME3,dd.D25KNAME4,dd.D25KNAME5,dd.D25KNAME6" & vbCrLf
        sql &= " ,dd.KNAME21" & vbCrLf ',轄區重點產業" & vbCrLf
        sql &= " ,cc.LEVEL1" & vbCrLf '審查計分表等級
        sql &= " ,cc.RLEVEL2" & vbCrLf '審查計分表等級(複審)
        sql &= " ,ISNULL(ISNULL(cc.RLEVEL2,cc.LEVEL1),'C') RLEVEL2X" & vbCrLf '審查計分表等級(複審) 複審若沒分數用初審分數
        sql &= " ,cc.CURESULT" & vbCrLf '首頁>>課程審查>>二階審查>>核班結果裡【核班結果】：通過的班級
        sql &= " ,cc.FIRSTSORT" & vbCrLf '優先序
        sql &= " ,case when cc.ICAPNUM is not null then '是' END IS_ICAPNUM" & vbCrLf '是否為iCap課程
        sql &= " FROM WC2 cc" & vbCrLf
        'sql &= " JOIN dbo.VIEW_RIDNAME rr on rr.RID=cc.RID" & vbCrLf
        sql &= " JOIN dbo.V_PLAN_DEPOT dd ON dd.PLANID=cc.PLANID and dd.COMIDNO=cc.COMIDNO and dd.SEQNO=cc.SEQNO" & vbCrLf
        sql &= " JOIN WC3 ON wc3.COMIDNO=cc.COMIDNO" & vbCrLf
        sql &= " ORDER BY cc.ORGNAME,cc.DISTID,cc.COMIDNO,cc.FIRSTSORT,cc.STDATE" & vbCrLf

        dt = DbAccess.GetDataTable(sql, objconn, parms)

        If dt.Rows.Count = 0 Then Return dt

        Dim s_COMIDNO As String = dt.Rows(0)("COMIDNO").ToString()
        Dim iST1GRANDTOTAL As Integer = 0
        For Each dr As DataRow In dt.Rows
            If s_COMIDNO = Convert.ToString(dr("COMIDNO")) Then
                iST1GRANDTOTAL += Val(dr("DEFGOVCOST"))
            Else
                s_COMIDNO = Convert.ToString(dr("COMIDNO"))
                iST1GRANDTOTAL = Val(dr("DEFGOVCOST"))
            End If
        Next
        Return dt
    End Function

    Function CHK_ColumnA(ByRef s_COLNAME As String) As Boolean
        Dim rst As Boolean = False
        If (s_COLNAME = "ORGNAME") Then Return True
        If (s_COLNAME = "ORGPLANNAME") Then Return True
        If (s_COLNAME = "ST1GRANDTOTAL") Then Return True
        If (s_COLNAME = "ST1DIFFAMOUNT") Then Return True
        If (s_COLNAME = "ST1OVERUPLIMIT") Then Return True
        If (s_COLNAME = "ADJGRANDTOTAL") Then Return True
        If (s_COLNAME = "ADJDIFFAMOUNT") Then Return True
        If (s_COLNAME = "ADJOVERUPLIMIT") Then Return True
        Return rst
    End Function

    Sub EXPORT_6()
        Dim dtX1 As DataTable = SEARCH_DATA1_dt()
        If dtX1 Is Nothing Then
            Common.MessageBox(Me, "查無匯出資料!!!")
            Exit Sub
        End If
        If dtX1.Rows.Count = 0 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg4)
            Return
        End If

        Dim sPattern As String = "" '序號,
        Dim sColumn As String = ""
        If sm.UserInfo.Years >= 2025 Then
            sPattern &= "計畫別,訓練單位名稱,立案縣市,可核配上限,等級,分署別,課程名稱(含期別),總補助費(元)(以訓練費用之80%估算),五分署初核累計,差額(初核-上限),是否超過上限,調整後累計,差額(調整後-上限),是否超過上限"
            sPattern &= ",亞洲矽谷,重點產業,台灣AI行動計畫,智慧國家方案,國家人才競爭力躍升方案,新南向政策,轄區重點產業,審查計分表等級,優先序,是否為iCap課程"
            sColumn &= "ORGPLANNAME,ORGNAME,ORGCTNAME,UPLIMIT,GRADE,DISTNAME,CLASSCNAME2,DEFGOVCOST,ST1GRANDTOTAL,ST1DIFFAMOUNT,ST1OVERUPLIMIT,ADJGRANDTOTAL,ADJDIFFAMOUNT,ADJOVERUPLIMIT"
            sColumn &= ",D25KNAME1,D25KNAME2,D25KNAME3,D25KNAME4,D25KNAME5,D25KNAME6,KNAME21,RLEVEL2X,FIRSTSORT,IS_ICAPNUM"
        Else
            sPattern &= "計畫別,訓練單位名稱,立案縣市,可核配上限,等級,分署別,課程名稱(含期別),總補助費(元)(以訓練費用之80%估算),五分署初核累計,差額(初核-上限),是否超過上限,調整後累計,差額(調整後-上限),是否超過上限"
            sPattern &= ",5+2產業,新南向政策,台灣AI行動計畫,數位國家創新經濟發展方案,國家資訊安全發展方案,前瞻基礎建設計畫,轄區重點產業,審查計分表等級,優先序,是否為iCap課程"
            sColumn &= "ORGPLANNAME,ORGNAME,ORGCTNAME,UPLIMIT,GRADE,DISTNAME,CLASSCNAME2,DEFGOVCOST,ST1GRANDTOTAL,ST1DIFFAMOUNT,ST1OVERUPLIMIT,ADJGRANDTOTAL,ADJDIFFAMOUNT,ADJOVERUPLIMIT"
            sColumn &= ",D20KNAME1,D20KNAME6,D20KNAME2,D20KNAME3,D20KNAME4,D20KNAME5,KNAME21,RLEVEL2X,FIRSTSORT,IS_ICAPNUM"
        End If
        Dim sPatternA() As String = Split(sPattern, ",")
        Dim sColumnA() As String = Split(sColumn, ",")

        Dim s_FILENAME1 As String = String.Concat("核配額度上限控管彙整表_", TIMS.GetDateNo2(3))
        '套CSS值
        Dim strSTYLE As String = ""
        strSTYLE &= "<style>"
        strSTYLE &= "td{mso-number-format:""\@"";}"
        strSTYLE &= ".noDecFormat{mso-number-format:""0"";}"
        strSTYLE &= "</style>"

        Dim strHTML As String = ""
        strHTML &= "<div>"
        strHTML &= "<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">"
        'Common.RespWrite(Me, "<tr>")

        '標題抬頭
        Dim ExportStr As String = "" '建立輸出文字
        ExportStr &= "<tr>"
        ExportStr &= String.Format("<td>{0}</td>", "序號") '& vbTab
        For i As Integer = 0 To sPatternA.Length - 1
            ExportStr &= String.Format("<td>{0}</td>", sPatternA(i)) '& vbTab
        Next
        ExportStr &= "</tr>" & vbCrLf
        strHTML &= ExportStr

        '建立資料面
        Dim iNum As Integer = 0
        Dim s_ORGPLANNAME_comidno As String = ""
        Dim s_COMIDNO As String = ""
        For Each dr As DataRow In dtX1.Rows
            iNum += 1
            ExportStr = "<tr>"
            ExportStr &= String.Format("<td>{0}</td>", iNum) '& vbTab
            If s_COMIDNO <> dr("COMIDNO") Then
                s_COMIDNO = dr("COMIDNO")
                For i As Integer = 0 To sColumnA.Length - 1
                    Dim flag_USE_ROWSPAN As Boolean = CHK_ColumnA(sColumnA(i))
                    If flag_USE_ROWSPAN Then
                        ExportStr &= String.Format("<td rowspan={1}>{0}</td>", Convert.ToString(dr(sColumnA(i))), Val(dr("ST1CNT")))
                    Else
                        ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr(sColumnA(i))))
                    End If
                Next
            Else
                For i As Integer = 0 To sColumnA.Length - 1
                    Dim flag_USE_ROWSPAN As Boolean = CHK_ColumnA(sColumnA(i))
                    If Not flag_USE_ROWSPAN Then
                        ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr(sColumnA(i))))
                    End If
                    'If flag_USE_ROWSPAN Then
                    '    'ExportStr &= String.Format("<td rowspan={1}>{0}</td>", Convert.ToString(dr(sColumnA(i))), Val(dr("ST1CNT")))
                    'Else
                    '    ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr(sColumnA(i))))
                    'End If
                Next
            End If

            ExportStr &= "</tr>" & vbCrLf
            strHTML &= ExportStr
        Next
        strHTML &= "</table>"
        strHTML &= "</div>"

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType)) 'EXCEL/PDF/ODS
        parmsExp.Add("FileName", s_FILENAME1)
        parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", strHTML)
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)
        TIMS.Utl_RespWriteEnd(Me, objconn, "")
        'TIMS.CloseDbConn(objconn) 'Response.End()
    End Sub

    '匯出'表單06_核配額度上限控管彙整表.xlsx
    Protected Sub BtnExport1_Click(sender As Object, e As EventArgs) Handles BtnExport1.Click
        Call EXPORT_6()
    End Sub

End Class
