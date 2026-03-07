Public Class SD_15_025
    Inherits AuthBasePage

    'Const cst_DepID06 As String = "50"  'SELECT * FROM V_DEPOT05N ORDER BY KID  --[5＋N產業]
    Dim flag_ROC As Boolean = TIMS.CHK_REPLACE2ROC_YEARS()

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        If Not IsPostBack Then
            Call CreateItem()
            DistID.Attributes("onclick") = "SelectAll('DistID','DistHidden');"     '選擇全部轄區
            TPlanID.Attributes("onclick") = "SelectAll('TPlanID','TPlanHidden');"  '選擇全部訓練計畫
            'KID_5.Attributes("onclick") = "SelectAll('KID_5','KID_5_hid');"        '產業別鍵詞
            Me.Print.Attributes("onclick") = "javascript:return CheckPrint();"     '列印檢查
            Export1.Attributes("onclick") = "javascript:return CheckPrint();"      '匯出名細檢查
        End If
    End Sub

    Sub CreateItem()
        Syear = TIMS.GetSyear(Syear)                     '年度
        Common.SetListItem(Syear, sm.UserInfo.Years)     '預設值
        '=================
        DistID = TIMS.Get_DistID(DistID)                 '轄區
        DistID.Items.Insert(0, New ListItem("全部", ""))
        'DistID.Items.Remove(DistID.Items.FindByText("新北市政府職業訓練中心"))
        'DistID.Items.Remove(DistID.Items.FindByText("臺北市職能發展學院"))
        'DistID.Items.Remove(DistID.Items.FindByText("泰山職業訓練中心"))
        'DistID.Items.Remove(DistID.Items.FindByText("勞動力發展署"))
        '=================
        TPlanID = TIMS.Get_TPlan(TPlanID, , 1, "Y")       '計畫
        'Get_KeyBusiness(KID_5, cst_DepID06)               '[5＋N產業]關鍵字查詢項目，by:20181002
        ''2019年啟用 work2019x01:2019 政府政策性產業
        'Dim KID20 As String = Convert.ToString(dr2("KID20"))
        ''2019年啟用 work2019x01:2019 政府政策性產業
        'Call TIMS.SetCblValue(CBLKID20_1, KID20)
        'Call TIMS.SetCblValue(CBLKID20_2, KID20)
        'Call TIMS.SetCblValue(CBLKID20_3, KID20)
        'Call TIMS.SetCblValue(CBLKID20_4, KID20)
        'Call TIMS.SetCblValue(CBLKID20_5, KID20)
        'Call TIMS.SetCblValue(CBLKID20_6, KID20)

        '2019年啟用 work2019x01:2019 政府政策性產業
        trKID20.Visible = False
        Dim flag_SHOW_2019_1 As Boolean = TIMS.SHOW_2019_1(sm)
        If flag_SHOW_2019_1 Then trKID20.Visible = True

        'trKID20.Visible
        '2018 (政府政策性產業)
        'sql = " SELECT KID, KNAME FROM KEY_BUSINESS WHERE DEPID = '20' ORDER BY KID"
        'Dim dtKID_N20 As DataTable = DbAccess.GetDataTable(sql, objconn)
        Dim dtKID_N20 As DataTable = TIMS.Get_BUSINESS_KID_dt(objconn, "20")
        Dim dtKID_N22 As DataTable = TIMS.Get_BUSINESS_KID_dt(objconn, "22")
        Call TIMS.GET_CBL_KID20(CBLKID20_1, dtKID_N20, 1)
        Call TIMS.GET_CBL_KID20(CBLKID20_2, dtKID_N20, 2)
        Call TIMS.GET_CBL_KID20(CBLKID20_3, dtKID_N20, 3)
        Call TIMS.GET_CBL_KID20(CBLKID20_4, dtKID_N20, 4)
        Call TIMS.GET_CBL_KID20(CBLKID20_5, dtKID_N20, 5)
        Call TIMS.GET_CBL_KID20(CBLKID20_6, dtKID_N20, 6)
        Call TIMS.GET_CBL_KID22(CBLKID22, dtKID_N22)
    End Sub

    ''' <summary> 取值-2019年啟用 work2019x01:2019 政府政策性產業-取值 </summary>
    ''' <returns></returns>
    Function GET_KID20_VAL() As String
        Dim rst As String = ""
        Dim tmp01 As String = ""
        tmp01 = TIMS.GetCblValue(CBLKID20_1)
        If tmp01 <> "" Then
            If rst <> "" Then rst &= ","
            rst &= tmp01
        End If
        tmp01 = TIMS.GetCblValue(CBLKID20_2)
        If tmp01 <> "" Then
            If rst <> "" Then rst &= ","
            rst &= tmp01
        End If
        tmp01 = TIMS.GetCblValue(CBLKID20_3)
        If tmp01 <> "" Then
            If rst <> "" Then rst &= ","
            rst &= tmp01
        End If
        tmp01 = TIMS.GetCblValue(CBLKID20_4)
        If tmp01 <> "" Then
            If rst <> "" Then rst &= ","
            rst &= tmp01
        End If
        tmp01 = TIMS.GetCblValue(CBLKID20_5)
        If tmp01 <> "" Then
            If rst <> "" Then rst &= ","
            rst &= tmp01
        End If
        tmp01 = TIMS.GetCblValue(CBLKID20_6)
        If tmp01 <> "" Then
            If rst <> "" Then rst &= ","
            rst &= tmp01
        End If
        Return rst
    End Function

    '產業別鍵詞
    'Function Get_KeyBusiness(ByVal obj As ListControl, ByVal DepID As String) As ListControl
    '    Dim dt As DataTable = Nothing
    '    Dim da As SqlDataAdapter = TIMS.GetOneDA(objconn)
    '    Dim sql As String = ""
    '    sql = "" & vbCrLf
    '    sql &= " SELECT REPLACE(a.KNAME, '【5＋N產業】', '') KNAME ,a.KID " & vbCrLf
    '    sql &= " ,a.SeqNo ,a.DepID ,a.Status " & vbCrLf
    '    sql &= " FROM KEY_BUSINESS a " & vbCrLf
    '    sql &= " JOIN KEY_DEPOT b ON b.Depid = a.DepID " & vbCrLf
    '    sql &= " WHERE 1=1 AND a.Status IS NULL " & vbCrLf
    '    da.SelectCommand.Parameters.Clear()
    '    If DepID <> "" Then
    '        sql &= " AND a.DepID = @DepID "
    '        da.SelectCommand.Parameters.Add("DepID", SqlDbType.VarChar).Value = DepID
    '    End If
    '    TIMS.Fill(sql, da, dt)
    '    With obj
    '        .Items.Clear()
    '        .DataSource = dt
    '        .DataTextField = "KNAME"
    '        .DataValueField = "KID"
    '        .DataBind()
    '        If TypeOf obj Is DropDownList Then .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
    '        If TypeOf obj Is CheckBoxList Then .Items.Insert(0, New ListItem("全部", ""))
    '    End With
    '    Return obj
    'End Function

    '匯出 Response
    Sub ExpReport1(ByRef dt As DataTable)
        ', ByVal sKeyName As String
        'Const cst_title1 As String = "5＋N產業課程明細表.xls"

        'Response.AddHeader("content-disposition", "attachment; filename=" & HttpUtility.UrlEncode(cst_title1, System.Text.Encoding.UTF8))
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("Big5")

        ''文件內容指定為Excel
        'Response.ContentType = "application/ms-excel"
        'strHTML &= ("<html>")
        'strHTML &= ("<head>")
        'strHTML &= ("<meta http-equiv=Content-Type content=text/html;charset=BIG5>")
        'strHTML &= ("</head>")

        'strHTML &= ("<body>")
        'strHTML &= ("<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")


        Dim sFileName1 As String = "5＋N產業課程明細表"
        'Dim strSTYLE As String = ""
        'strSTYLE &= ("<style>")

        Dim strHTML As String = ""
        strHTML &= ("<div>")
        strHTML &= ("<table>")

        Dim ExportStr As String = ""
        '建立抬頭
        ExportStr = ""
        ExportStr += "<tr>" & vbCrLf
        ExportStr &= "<td>序號</td>" & vbTab
        ExportStr &= "<td>轄區</td>" & vbTab
        ExportStr &= "<td>訓練計畫</td>" & vbTab
        ExportStr &= "<td>訓練機構名稱</td>" & vbTab
        ExportStr &= "<td>班別名稱</td>" & vbTab
        ExportStr &= "<td>訓練職類</td>" & vbTab
        'ExportStr &= "<td>訓練性質</td>" & vbTab
        'ExportStr &= "<td>5+N產業類別</td>" & vbTab  '5+N產業類別，by:20181026
        ExportStr &= "<td>5+2產業創新計畫</td>" & vbTab
        ExportStr &= "<td>台灣AI行動計畫</td>" & vbTab
        ExportStr &= "<td>數位國家創新經濟發展方案</td>" & vbTab
        ExportStr &= "<td>國家資通安全發展方案</td>" & vbTab
        ExportStr &= "<td>前瞻基礎建設計畫</td>" & vbTab
        ExportStr &= "<td>新南向政策</td>" & vbTab

        ExportStr &= "<td>訓練時段</td>" & vbTab
        ExportStr &= "<td>開訓日期</td>" & vbTab
        ExportStr &= "<td>結訓日期</td>" & vbTab
        ExportStr &= "<td>招生人數</td>" & vbTab
        ExportStr &= "<td>時數</td>" & vbTab
        ExportStr &= "<td>開訓人數</td>" & vbTab
        ExportStr &= "<td>結訓人數</td>" & vbTab
        ExportStr += "</tr>" & vbCrLf
        strHTML &= (TIMS.sUtl_AntiXss(ExportStr))

        Dim iALL_TNum As Integer = 0 '招生人數
        Dim iALL_THours As Integer = 0 '時數
        Dim iALL_openNum As Integer = 0 '開訓人數
        Dim iALL_closeNum As Integer = 0 '結訓人數

        '建立資料面
        Dim i As Integer = 0
        For Each dr As DataRow In dt.Rows
            i += 1
            iALL_TNum += CInt(Val(dr("TNum")))
            iALL_THours += CInt(Val(dr("THours")))
            iALL_openNum += CInt(Val(dr("openNum")))
            iALL_closeNum += CInt(Val(dr("closeNum")))

            ExportStr = ""
            ExportStr &= "<tr>" & vbCrLf
            ExportStr &= "<td>" & Convert.ToString(i) & "</td>" & vbTab                    '序號
            ExportStr &= "<td>" & Convert.ToString(dr("distname")) & "</td>" & vbTab       '轄區
            ExportStr &= "<td>" & Convert.ToString(dr("planname")) & "</td>" & vbTab       '訓練計畫
            ExportStr &= "<td>" & Convert.ToString(dr("orgname")) & "</td>" & vbTab        '訓練機構名稱
            ExportStr &= "<td>" & Convert.ToString(dr("classcname")) & "</td>" & vbTab     '班別名稱
            ExportStr &= "<td>" & Convert.ToString(dr("trainName")) & "</td>" & vbTab      '訓練職類
            'ExportStr &= "<td>" & Convert.ToString(dr("PropertyID")) & "</td>" & vbTab    '訓練性質
            'ExportStr &= "<td>" & Convert.ToString(dr("INDUSTRY")) & "</td>" & vbTab       '5+N產業類別，by:20181026
            ExportStr &= "<td>" & Convert.ToString(dr("D20KNAME1")) & "</td>" & vbTab '5+2產業創新計畫
            ExportStr &= "<td>" & Convert.ToString(dr("D20KNAME2")) & "</td>" & vbTab '台灣AI行動計畫
            ExportStr &= "<td>" & Convert.ToString(dr("D20KNAME3")) & "</td>" & vbTab '"數位國家創新經濟發展方案</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("D20KNAME4")) & "</td>" & vbTab '"國家資通安全發展方案</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("D20KNAME5")) & "</td>" & vbTab '"前瞻基礎建設計畫</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("D20KNAME6")) & "</td>" & vbTab '"新南向政策</td>" & vbTab

            ExportStr &= "<td>" & Convert.ToString(dr("hourRanName")) & "</td>" & vbTab    '訓練時段
            ExportStr &= "<td>" & Convert.ToString(dr("STDate")) & "</td>" & vbTab         '開訓日期
            ExportStr &= "<td>" & Convert.ToString(dr("FTDate")) & "</td>" & vbTab         '結訓日期
            ExportStr &= "<td>" & Convert.ToString(Val(dr("TNum"))) & "</td>" & vbTab      '招生人數
            ExportStr &= "<td>" & Convert.ToString(Val(dr("THours"))) & "</td>" & vbTab    '時數
            ExportStr &= "<td>" & Convert.ToString(Val(dr("openNum"))) & "</td>" & vbTab   '開訓人數
            ExportStr &= "<td>" & Convert.ToString(Val(dr("closeNum"))) & "</td>" & vbTab  '結訓人數
            ExportStr += "</tr>" & vbCrLf
            strHTML &= (TIMS.sUtl_AntiXss(ExportStr))
        Next

        '建立尾1
        ExportStr = ""
        ExportStr &= "<tr>" & vbCrLf
        ExportStr &= "<td>合計</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(dt.Rows.Count) & "班</td>" & vbTab
        ExportStr &= "<td></td>" & vbTab
        ExportStr &= "<td></td>" & vbTab
        ExportStr &= "<td></td>" & vbTab
        ExportStr &= "<td></td>" & vbTab
        ExportStr &= "<td></td>" & vbTab
        ExportStr &= "<td></td>" & vbTab
        ExportStr &= "<td></td>" & vbTab
        ExportStr &= "<td></td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(iALL_TNum) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(iALL_THours) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(iALL_openNum) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(iALL_closeNum) & "</td>" & vbTab
        ExportStr += "</tr>" & vbCrLf
        strHTML &= (TIMS.sUtl_AntiXss(ExportStr))

        '建立尾2
        'If sKeyName <> "" Then
        '    ExportStr = ""
        '    ExportStr += "<tr>" & vbCrLf
        '    ExportStr &= "<td>查詢關鍵字：</td>" & vbTab
        '    ExportStr &= "<td colspan=""9"">" & Convert.ToString(sKeyName) & "</td>" & vbTab
        '    ExportStr &= "<td></td>" & vbTab
        '    ExportStr &= "<td></td>" & vbTab
        '    ExportStr &= "<td></td>" & vbTab
        '    ExportStr &= "<td></td>" & vbTab
        '    ExportStr += "</tr>" & vbCrLf
        '    strHTML &= (TIMS.sUtl_AntiXss(ExportStr))
        'End If
        strHTML &= ("</table>")
        strHTML &= ("</div>")

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        'parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", strHTML)
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)
        TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
    End Sub

    Function GetExportDt() As DataTable
        'Dim dt As New DataTable
        'cmdArg &= "&KID20=" & GET_KID20_VAL() 'TIMS.EncryptAes(vKID20) 

        Dim s_KID20 As String = GET_KID20_VAL()
        'If s_KID20 <> "" Then s_KID20 = TIMS.CombiSQM2IN(s_KID20) 

        '報表要用的轄區參數
        Dim DistID1 As String = ""
        DistID1 = TIMS.GetCheckBoxListRptVal(DistID, 1)

        '報表要用的訓練計畫參數
        Dim TPlanID1 As String = ""
        TPlanID1 = TIMS.GetCheckBoxListRptVal(TPlanID, 1)

        '報表要用的新興產業
        'Dim sKID_5 As String = ""
        'sKID_5 = TIMS.GetCheckBoxListRptVal(KID_5, 1)

        'Dim KeyNames As String = ""
        'KeyNames = ""
        'If sKID_5 <> "" Then
        '    If KeyNames <> "" Then KeyNames += ","
        '    KeyNames += GetKeyName(cst_DepID06, sKID_5.Replace("\'", "'")) '取出關鍵字
        'End If
        'Dim sqlKeyNames As String = ""
        'sqlKeyNames = ChgSqlWhere(KeyNames)

        'Dim MyValue As String = ""
        'MyValue = ""
        'MyValue += "&Years=" & Syear.SelectedValue '年度
        'If flag_ROC Then
        '    MyValue += "&STDate1=" & TIMS.cdate18(Me.STDate1.Text)
        '    MyValue += "&STDate2=" & TIMS.cdate18(Me.STDate2.Text)
        '    MyValue += "&FTDate1=" & TIMS.cdate18(Me.FTDate1.Text)
        '    MyValue += "&FTDate2=" & TIMS.cdate18(Me.FTDate2.Text)
        'Else
        '    MyValue += "&STDate1=" & Me.STDate1.Text
        '    MyValue += "&STDate2=" & Me.STDate2.Text
        '    MyValue += "&FTDate1=" & Me.FTDate1.Text
        '    MyValue += "&FTDate2=" & Me.FTDate2.Text
        'End If
        'MyValue += "&DistID=" & DistID1
        'MyValue += "&TPlanID=" & TPlanID1


        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT ip.distid" & vbCrLf
        sql &= " ,ip.planid" & vbCrLf
        sql &= " ,ip.distname" & vbCrLf
        sql &= " ,ip.planname" & vbCrLf
        sql &= " ,oo.orgname" & vbCrLf
        sql &= " ,cc.OCID" & vbCrLf
        sql &= " ,cc.classcname" & vbCrLf
        sql &= " ,vt.trainName" & vbCrLf
        sql &= " ,CASE WHEN ISNULL(ip.PropertyID, 0) = 0 THEN '職前' WHEN ISNULL(ip.PropertyID, 0)=1 THEN '在職' END PropertyID" & vbCrLf
        sql &= " ,k1.hourRanName hourRanName" & vbCrLf
        If flag_ROC Then
            '開訓日期/結訓日期(民國年)
            sql &= " ,dbo.RT_DataFormat(cc.STDate) STDate" & vbCrLf
            sql &= " ,dbo.RT_DataFormat(cc.FTDate) FTDate" & vbCrLf
        Else
            sql &= " ,CONVERT(VARCHAR, cc.stdate, 111) STDate " & vbCrLf  '開訓日期(西元年)
            sql &= " ,CONVERT(VARCHAR, cc.ftdate, 111) FTDate " & vbCrLf  '結訓日期(西元年)
        End If
        sql &= " ,cc.TNum" & vbCrLf
        sql &= " ,cc.THours" & vbCrLf
        sql &= " ,ISNULL(dbo.FN_GET_STDCNT(CC.OCID,1), 0) openNum" & vbCrLf
        sql &= " ,ISNULL(dbo.FN_GET_STDCNT(CC.OCID,52), 0) closeNum" & vbCrLf
        sql &= " ,dbo.DECODE2(dd.AppResult,'Y',dd.D20KNAME1 ,'無') D20KNAME1" & vbCrLf
        sql &= " ,dbo.DECODE2(dd.AppResult,'Y',dd.D20KNAME2 ,'無') D20KNAME2" & vbCrLf
        sql &= " ,dbo.DECODE2(dd.AppResult,'Y',dd.D20KNAME3 ,'無') D20KNAME3" & vbCrLf
        sql &= " ,dbo.DECODE2(dd.AppResult,'Y',dd.D20KNAME4 ,'無') D20KNAME4" & vbCrLf
        sql &= " ,dbo.DECODE2(dd.AppResult,'Y',dd.D20KNAME5 ,'無') D20KNAME5" & vbCrLf
        sql &= " ,dbo.DECODE2(dd.AppResult,'Y',dd.D20KNAME6 ,'無') D20KNAME6" & vbCrLf
        sql &= " FROM CLASS_CLASSINFO cc" & vbCrLf
        sql &= " JOIN PLAN_PLANINFO pp ON pp.planid = cc.planid AND pp.comidno = cc.comidno AND pp.seqno = cc.seqno" & vbCrLf
        sql &= " JOIN VIEW_PLAN ip ON ip.planid = cc.planid" & vbCrLf
        sql &= " JOIN VIEW_TRAINTYPE vt ON vt.TMID = cc.TMID" & vbCrLf
        sql &= " JOIN ORG_ORGINFO oo ON oo.comidno = cc.comidno" & vbCrLf
        sql &= " LEFT JOIN KEY_HOURRAN k1 ON k1.HRID = cc.TPeriod" & vbCrLf
        sql &= " LEFT JOIN dbo.V_PLAN_DEPOT dd WITH(NOLOCK) on dd.planid=pp.planid and dd.comidno=pp.comidno and dd.seqno=pp.seqno" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND cc.IsSuccess = 'Y'" & vbCrLf
        sql &= " AND cc.NotOpen = 'N'" & vbCrLf
        'sql &= " AND ip.Years = '2019'" & vbCrLf
        'sql &= " AND ip.DistID IN ('001')" & vbCrLf
        'sql &= " AND ip.TPlanID IN ('06')" & vbCrLf
        If Syear.SelectedValue <> "" Then sql &= " AND ip.Years = '" & Syear.SelectedValue & "' " & vbCrLf
        If flag_ROC Then
            If Me.STDate1.Text <> "" Then sql &= " AND cc.STDate >= CONVERT(DATETIME, '" & TIMS.Cdate18(Me.STDate1.Text) & "', 111) " & vbCrLf
            If Me.STDate2.Text <> "" Then sql &= " AND cc.STDate <= CONVERT(DATETIME, '" & TIMS.Cdate18(Me.STDate2.Text) & "', 111) " & vbCrLf
            If Me.FTDate1.Text <> "" Then sql &= " AND cc.FTDate >= CONVERT(DATETIME, '" & TIMS.Cdate18(Me.FTDate1.Text) & "', 111) " & vbCrLf
            If Me.FTDate2.Text <> "" Then sql &= " AND cc.FTDate <= CONVERT(DATETIME, '" & TIMS.Cdate18(Me.FTDate2.Text) & "', 111) " & vbCrLf
        Else
            If Me.STDate1.Text <> "" Then sql &= " AND cc.STDate >= CONVERT(DATETIME, '" & Me.STDate1.Text & "', 111) " & vbCrLf
            If Me.STDate2.Text <> "" Then sql &= " AND cc.STDate <= CONVERT(DATETIME, '" & Me.STDate2.Text & "', 111) " & vbCrLf
            If Me.FTDate1.Text <> "" Then sql &= " AND cc.FTDate >= CONVERT(DATETIME, '" & Me.FTDate1.Text & "', 111) " & vbCrLf
            If Me.FTDate2.Text <> "" Then sql &= " AND cc.FTDate <= CONVERT(DATETIME, '" & Me.FTDate2.Text & "', 111) " & vbCrLf
        End If
        If DistID1 <> "" Then sql &= " AND ip.DistID IN (" & DistID1.Replace("\'", "'") & ") " & vbCrLf
        If TPlanID1 <> "" Then sql &= " AND ip.TPlanID IN (" & TPlanID1.Replace("\'", "'") & ") " & vbCrLf
        If s_KID20 <> "" Then sql &= " AND dd.KID20 IN (" & TIMS.CombiSQM2IN(s_KID20) & ")" & vbCrLf

        'If sqlKeyNames <> "" Then
        '    sql &= " AND EXISTS ( " & vbCrLf
        '    sql &= "  SELECT 'x' FROM class_classinfo x WHERE 1=1 " & vbCrLf
        '    sql &= "  AND x.ocid = cc.ocid " & vbCrLf
        '    sql &= "  AND (1!=1 " & vbCrLf
        '    sql &= sqlKeyNames
        '    sql &= " ) " & vbCrLf
        '    sql &= " ) " & vbCrLf
        'End If

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)
        dt.DefaultView.Sort = "distid ,planid ,classcname"
        dt = TIMS.dv2dt(dt.DefaultView)

        Return dt
    End Function

    '取出關鍵字
    'Function GetKeyName(ByVal DepID As String, ByVal KID As String) As String
    '    Dim rst As String = ""
    '    If KID <> "" Then
    '        Dim sql As String = ""
    '        sql = "" & vbCrLf
    '        sql &= " SELECT k.DepID ,k.KID ,d.DNAME ,d.Years ,b.KName ,k.KeyName " & vbCrLf
    '        sql &= " FROM KEY_BUSINESS b " & vbCrLf
    '        sql &= " JOIN KEY_DEPOT d ON d.DepID = b.DepID " & vbCrLf
    '        sql &= " JOIN KEY_BUSINESSKEYS k ON k.DepID = b.DepID AND k.KID = b.KID " & vbCrLf
    '        sql &= " WHERE 1=1 " & vbCrLf
    '        sql &= " AND k.DepID = '" & DepID & "'" & vbCrLf
    '        sql &= " AND k.KID IN (" & KID & ")" & vbCrLf
    '        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)
    '        rst = ""
    '        For Each dr As DataRow In dt.Rows
    '            If rst.IndexOf(Convert.ToString(dr("KeyName")).Trim) = -1 Then  'edit，by:20181026
    '                If rst <> "" Then rst += ","
    '                rst += Convert.ToString(dr("KeyName"))
    '            End If
    '        Next
    '    End If
    '    Return rst
    'End Function

    '組合SQL語法
    'Function ChgSqlWhere(ByVal KeyNames As String) As String
    '    Dim rst As String = ""
    '    Dim cst_dtNickName As String = "x"
    '    If KeyNames <> "" Then
    '        rst = ""
    '        If KeyNames.IndexOf(",") > -1 Then
    '            Dim aryStr() As String = Split(KeyNames, ",")
    '            For i As Integer = 0 To aryStr.Length - 1
    '                rst += " OR " & cst_dtNickName & ".classcname LIKE N'%" & aryStr(i) & "%'" & vbCrLf '換行
    '            Next
    '        Else
    '            rst += " OR " & cst_dtNickName & ".classcname LIKE N'%" & KeyNames & "%'" & vbCrLf '換行
    '        End If
    '    End If
    '    Return rst
    'End Function

    '匯出 SQL
    Private Sub Export1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Export1.Click
        'Call ExpReport1(dt, KeyNames)
        Dim dt As DataTable = GetExportDt()
        Call ExpReport1(dt)
    End Sub

    '列印
    Protected Sub Print_Click(sender As Object, e As EventArgs) Handles Print.Click
        '(目前尚未提供列印報表功能,未來要製作的話,可參考程式編號"TR_05_018_R"，by:20181002)
    End Sub
End Class