Public Class SD_15_008_R_Prt2
    Inherits AuthBasePage

    '2017
    Const cst_rpt_性別 As String = "SD_15_008_R11_2012"
    Const cst_rpt_年齡 As String = "SD_15_008_R12_2012"
    Const cst_rpt_教育程度 As String = "SD_15_008_R13_2012"
    Const cst_rpt_身分別 As String = "SD_15_008_R14_2012"
    Const cst_rpt_工作年資 As String = "SD_15_008_R15_2012"
    Const cst_rpt_地理分佈 As String = "SD_15_008_R16_2012"
    Const cst_rpt_公司行業別 As String = "SD_15_008_R17_2012"
    Const cst_rpt_公司規模 As String = "SD_15_008_R18_2012"
    Const cst_rpt_參訓動機 As String = "SD_15_008_R19_2012"
    Const cst_rpt_訓後動向 As String = "SD_15_008_R20_2012"
    Const cst_rpt_參訓單位類別 As String = "SD_15_008_R21_2012"
    Const cst_rpt_參加課程職能別 As String = "SD_15_008_R22_2012"
    Const cst_rpt_參加課程型態別 As String = "SD_15_008_R23_2012"
    Const cst_rpt_訓練業別 As String = "SD_15_008_R24_2012"
    Const cst_rpt_不設定 As String = "SD_15_008_R25_2012"

    Dim arrItem(137, 2) As String '題目與答案
    Dim arrSpanRow(137) As String '處理合併表格
    Dim gResDt As DataTable = Nothing '計算行數共用表格

    Dim PName As String = ""
    Dim plankind As String = ""
    Dim TPlanID As String = ""
    Dim Years As String = ""
    Dim OCID As String = ""
    Dim RID As String = ""
    Dim SearchPlan As String = "" '"",G,W
    Dim PackageType As String = "" '"",2,3
    Dim filename As String = ""

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load 'Me.Load
        If TIMS.ChkSession(sm) Then
            Common.RespWrite(Me, "<script>alert('您的登入資訊已經遺失，請重新登入');window.close();</script>")
            TIMS.Utl_RespWriteEnd(Me, objconn, "") ' Response.End()
        End If
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload

        If Not IsPostBack Then
            Call Create11()
        End If

    End Sub

    Sub Create11()
        TPlanID = TIMS.ClearSQM(Request("TPlanID"))
        Years = TIMS.ClearSQM(Request("Years"))
        OCID = TIMS.ClearSQM(Request("OCID"))
        RID = TIMS.ClearSQM(Request("RID"))
        SearchPlan = TIMS.ClearSQM(Request("SearchPlan"))
        PackageType = TIMS.ClearSQM(Request("PackageType"))
        filename = TIMS.ClearSQM(Request("filename"))
        If RID = "A" Then RID = "" '署(局)屬

        Dim sValue As String = ""
        TIMS.SetMyValue(sValue, "RID", RID)
        TIMS.SetMyValue(sValue, "OCID", OCID)
        TIMS.SetMyValue(sValue, "TPlanID", TPlanID)
        TIMS.SetMyValue(sValue, "Years", Years)
        TIMS.SetMyValue(sValue, "PackageType", PackageType)
        TIMS.SetMyValue(sValue, "SearchPlan", SearchPlan)
        TIMS.SetMyValue(sValue, "filename", filename)

        Call setItemArray() '題目內容
        gResDt = sUtl_RsDt(sValue)
        If gResDt.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無資料!!")
            Exit Sub
        End If

        If TIMS.dtNODATA(gResDt) Then Exit Sub
        'If gResDt.Rows.Count = 0 Then Exit Sub

        btnCancel.Attributes.Add("onclick", "window.close();")

        Dim tDt As DataTable = RotationDT(gResDt)
        Const cst_pgCnt1 As Integer = 38
        Select Case filename
            Case cst_rpt_性別 '"SD_15_008_R11_2012"
                Call PrintDiv(tDt, filename, "200", "150", cst_pgCnt1, "10", "H")
            Case cst_rpt_年齡 '"SD_15_008_R12_2012"
                Call PrintDiv(tDt, filename, "120", "120", cst_pgCnt1, "7", "L")
            Case cst_rpt_教育程度 '"SD_15_008_R13_2012"
                Call PrintDiv(tDt, filename, "100", "110", cst_pgCnt1, "7", "L")
            Case cst_rpt_身分別 '"SD_15_008_R14_2012"
                Call PrintDiv(tDt, filename, "90", "65", cst_pgCnt1, "5", "L")
            Case cst_rpt_工作年資 '"SD_15_008_R15_2012"
                Call PrintDiv(tDt, filename, "150", "150", cst_pgCnt1, "10", "H")
            Case cst_rpt_地理分佈 '"SD_15_008_R16_2012"
                Call PrintDiv(tDt, filename, "95", "70", cst_pgCnt1, "6", "L")
            Case cst_rpt_公司行業別 '"SD_15_008_R17_2012"
                Call PrintDiv(tDt, filename, "95", "70", cst_pgCnt1, "6", "L")
            Case cst_rpt_公司規模 '"SD_15_008_R18_2012"
                Call PrintDiv(tDt, filename, "150", "150", cst_pgCnt1, "10", "H")
            Case cst_rpt_參訓動機 '"SD_15_008_R19_2012"
                Call PrintDiv(tDt, filename, "130", "140", cst_pgCnt1, "10", "H")
            Case cst_rpt_訓後動向 '"SD_15_008_R20_2012"
                Call PrintDiv(tDt, filename, "150", "150", cst_pgCnt1, "10", "H")
            Case cst_rpt_參訓單位類別 '"SD_15_008_R21_2012"
                Call PrintDiv(tDt, filename, "90", "100", cst_pgCnt1, "7", "L")
            Case cst_rpt_參加課程職能別 '"SD_15_008_R22_2012"
                Call PrintDiv(tDt, filename, "130", "120", cst_pgCnt1, "10", "H")
            Case cst_rpt_參加課程型態別 '"SD_15_008_R23_2012"
                Call PrintDiv(tDt, filename, "150", "150", cst_pgCnt1, "10", "H")
            Case cst_rpt_訓練業別 '"SD_15_008_R24_2012"
                Call PrintDiv(tDt, filename, "55", "70", cst_pgCnt1, "6", "L")
            Case cst_rpt_不設定 '"SD_15_008_R25_2012"
                Call PrintDiv(tDt, filename, "200", "150", cst_pgCnt1, "10", "H")
        End Select
    End Sub

    '組合資料
    Function sUtl_RsDt(ByVal sValue As String) As DataTable
        Dim RID As String = TIMS.GetMyValue(sValue, "RID")
        Dim OCID As String = TIMS.GetMyValue(sValue, "OCID")
        Dim TPlanID As String = TIMS.GetMyValue(sValue, "TPlanID")
        Dim Years As String = TIMS.GetMyValue(sValue, "Years")
        Dim PackageType As String = TIMS.GetMyValue(sValue, "PackageType")
        Dim SearchPlan As String = TIMS.GetMyValue(sValue, "SearchPlan") 'G/W
        Dim filename As String = TIMS.GetMyValue(sValue, "filename")

        Dim iType As Integer = 1
        Select Case SearchPlan
            Case "G", "W"
                'iType = 1
            Case Else
                iType = 2 '異常
        End Select

        Dim sTitle2 As String = "未設定"
        Select Case filename
            Case cst_rpt_性別 '"SD_15_008_R11_2012"
                sTitle2 = "性別"
            Case cst_rpt_年齡 '"SD_15_008_R12_2012"
                sTitle2 = "年齡"
            Case cst_rpt_教育程度 '"SD_15_008_R13_2012"
                sTitle2 = "教育程度"
            Case cst_rpt_身分別 '"SD_15_008_R14_2012"
                sTitle2 = "身分別"
            Case cst_rpt_工作年資 '"SD_15_008_R15_2012"
                sTitle2 = "工作年資"
            Case cst_rpt_地理分佈 '"SD_15_008_R16_2012"
                sTitle2 = "地理分佈"
            Case cst_rpt_公司行業別 '"SD_15_008_R17_2012"
                sTitle2 = "公司行業別"
            Case cst_rpt_公司規模 '"SD_15_008_R18_2012"
                sTitle2 = "公司規模"
            Case cst_rpt_參訓動機 '"SD_15_008_R19_2012"
                sTitle2 = "參訓動機"
            Case cst_rpt_訓後動向 '"SD_15_008_R20_2012"
                sTitle2 = "訓後動向"
            Case cst_rpt_參訓單位類別 '"SD_15_008_R21_2012"
                sTitle2 = "參訓單位類別"
            Case cst_rpt_參加課程職能別 '"SD_15_008_R22_2012"
                'sTitle2 = "參加課程職能別"
                sTitle2 = "參加課程職能"
            Case cst_rpt_參加課程型態別 '"SD_15_008_R23_2012"
                'sTitle2 = "參加課程型態別"
                sTitle2 = "參加課程型態"
            Case cst_rpt_訓練業別 '"SD_15_008_R24_2012"
                sTitle2 = "訓練業別"
            Case cst_rpt_不設定
                sTitle2 = "不設定"
        End Select

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT G.PName,G.MID NAME,G.PlanKind" & vbCrLf
        sql &= " ,ISNULL(COUNT(g.a11),0) a11,ISNULL(COUNT(g.a12),0) a12,ISNULL(COUNT(g.a13),0) a13" & vbCrLf
        sql &= " ,ISNULL(COUNT(g.a14),0) a14,ISNULL(COUNT(g.a15),0) a15,ISNULL(COUNT(g.a16),0) a16" & vbCrLf
        sql &= " ,ISNULL(COUNT(g.a21),0) a21,ISNULL(COUNT(g.a22),0) a22,ISNULL(COUNT(g.a23),0) a23,ISNULL(COUNT(g.a24),0) a24,ISNULL(COUNT(g.a25),0) a25" & vbCrLf
        sql &= " ,ISNULL(COUNT(g.a31),0) a31,ISNULL(COUNT(g.a32),0) a32,ISNULL(COUNT(g.a33),0) a33,ISNULL(COUNT(g.a34),0) a34" & vbCrLf
        sql &= " ,ISNULL(COUNT(g.a41),0) a41,ISNULL(COUNT(g.a42),0) a42,ISNULL(COUNT(g.a43),0) a43,ISNULL(COUNT(g.a44),0) a44,ISNULL(COUNT(g.a45),0) a45" & vbCrLf
        sql &= " ,ISNULL(COUNT(g.a51),0) a51,ISNULL(COUNT(g.a52),0) a52,ISNULL(COUNT(g.a53),0) a53,ISNULL(COUNT(g.a54),0) a54,ISNULL(COUNT(g.a55),0) a55" & vbCrLf
        sql &= " ,ISNULL(COUNT(g.a61),0) a61,ISNULL(COUNT(g.a62),0) a62,ISNULL(COUNT(g.a63),0) a63,ISNULL(COUNT(g.a64),0) a64,ISNULL(COUNT(g.a65),0) a65" & vbCrLf
        sql &= " ,ISNULL(COUNT(g.a71),0) a71,ISNULL(COUNT(g.a72),0) a7,ISNULL(COUNT(g.a73),0) a73,ISNULL(COUNT(g.a74),0) a74" & vbCrLf
        For ix8 As Integer = 1 To 8
            sql &= " ,ISNULL(COUNT(g.a21" & ix8 & "1),0) a21" & ix8 & "1" & vbCrLf
            sql &= " ,ISNULL(COUNT(g.a21" & ix8 & "2),0) a21" & ix8 & "2" & vbCrLf
            sql &= " ,ISNULL(COUNT(g.a21" & ix8 & "3),0) a21" & ix8 & "3" & vbCrLf
            sql &= " ,ISNULL(COUNT(g.a21" & ix8 & "4),0) a21" & ix8 & "4" & vbCrLf
            sql &= " ,ISNULL(COUNT(g.a21" & ix8 & "5),0) a21" & ix8 & "5" & vbCrLf
        Next
        For ix8 As Integer = 1 To 6
            sql &= " ,ISNULL(COUNT(g.a22" & ix8 & "1),0) a22" & ix8 & "1" & vbCrLf
            sql &= " ,ISNULL(COUNT(g.a22" & ix8 & "2),0) a22" & ix8 & "2" & vbCrLf
            sql &= " ,ISNULL(COUNT(g.a22" & ix8 & "3),0) a22" & ix8 & "3" & vbCrLf
            sql &= " ,ISNULL(COUNT(g.a22" & ix8 & "4),0) a22" & ix8 & "4" & vbCrLf
            sql &= " ,ISNULL(COUNT(g.a22" & ix8 & "5),0) a22" & ix8 & "5" & vbCrLf
        Next
        sql &= " FROM (" & vbCrLf

        Select Case iType
            Case 1
                sql &= " SELECT kp.PlanName+'('+cj1.name+')' COLLATE Chinese_Taiwan_Stroke_CS_AS PlanKind" & vbCrLf
            Case Else
                sql &= " SELECT kp.PlanName PlanKind" & vbCrLf '合併為1筆
        End Select
        sql &= " ,cast((ip.Years-1911) as varchar)+'年度　訓後動態調查統計表 " & sTitle2 & "' PName" & vbCrLf

        Select Case filename
            Case cst_rpt_性別 '"SD_15_008_R11_2012"
                'sTitle2 = "性別"
                sql &= " ,dbo.DECODE6(ss.Sex,'M','男','F','女','無') MID" & vbCrLf
                'sql += " ,dbo.DECODE(ss.Sex,'M',1,'F',2) SORT1" & vbCrLf
            Case cst_rpt_年齡 '"SD_15_008_R12_2012"
                'sTitle2 = "年齡"
                sql &= " ,case when 15> DATEPART(YEAR, cc.STDATE)-DATEPART(YEAR, ss.Birthday) then '15以下' " & vbCrLf
                sql &= " when 20> DATEPART(YEAR, cc.STDATE)-DATEPART(YEAR, ss.Birthday) and 15<=DATEPART(YEAR, cc.STDATE)-DATEPART(YEAR, ss.Birthday) then '15~19' " & vbCrLf
                sql &= " when 25> DATEPART(YEAR, cc.STDATE)-DATEPART(YEAR, ss.Birthday) and 20<=DATEPART(YEAR, cc.STDATE)-DATEPART(YEAR, ss.Birthday) then '20~24' " & vbCrLf
                sql &= " when 30> DATEPART(YEAR, cc.STDATE)-DATEPART(YEAR, ss.Birthday) and 25<=DATEPART(YEAR, cc.STDATE)-DATEPART(YEAR, ss.Birthday) then '25~29' " & vbCrLf
                sql &= " when 35> DATEPART(YEAR, cc.STDATE)-DATEPART(YEAR, ss.Birthday) and 30<=DATEPART(YEAR, cc.STDATE)-DATEPART(YEAR, ss.Birthday) then '30~34' " & vbCrLf
                sql &= " when 40> DATEPART(YEAR, cc.STDATE)-DATEPART(YEAR, ss.Birthday) and 35<=DATEPART(YEAR, cc.STDATE)-DATEPART(YEAR, ss.Birthday) then '35~39' " & vbCrLf
                sql &= " when 45> DATEPART(YEAR, cc.STDATE)-DATEPART(YEAR, ss.Birthday) and 40<=DATEPART(YEAR, cc.STDATE)-DATEPART(YEAR, ss.Birthday) then '40~44' " & vbCrLf
                sql &= " when 50> DATEPART(YEAR, cc.STDATE)-DATEPART(YEAR, ss.Birthday) and 45<=DATEPART(YEAR, cc.STDATE)-DATEPART(YEAR, ss.Birthday) then '45~49' " & vbCrLf
                sql &= " when 55> DATEPART(YEAR, cc.STDATE)-DATEPART(YEAR, ss.Birthday) and 50<=DATEPART(YEAR, cc.STDATE)-DATEPART(YEAR, ss.Birthday) then '50~54' " & vbCrLf
                sql &= " when 60> DATEPART(YEAR, cc.STDATE)-DATEPART(YEAR, ss.Birthday) and 55<=DATEPART(YEAR, cc.STDATE)-DATEPART(YEAR, ss.Birthday) then '55~59' " & vbCrLf
                sql &= " else '60' end MID" & vbCrLf

            Case cst_rpt_教育程度 '"SD_15_008_R13_2012"
                'sTitle2 = "教育程度"
                sql &= " ,kd.Name MID" & vbCrLf
            Case cst_rpt_身分別 '"SD_15_008_R14_2012"
                'sTitle2 = "身分別"
                sql &= " ,kd2.Name MID" & vbCrLf
            Case cst_rpt_工作年資 '"SD_15_008_R15_2012"
                'sTitle2 = "工作年資"
                sql &= " ,case when ISNULL(stb.q61,0) <3 then '3年以下' " & vbCrLf
                sql &= " when ISNULL(stb.q61,0) <5 then '3~5年' " & vbCrLf
                sql &= " when ISNULL(stb.q61,0)<10 then '5~10年' " & vbCrLf
                sql &= " else '10年以上' end MID " & vbCrLf
            Case cst_rpt_地理分佈 '"SD_15_008_R16_2012"
                'sTitle2 = "地理分佈"
                sql &= " ,iz.CTNAME MID " & vbCrLf
            Case cst_rpt_公司行業別 '"SD_15_008_R17_2012"
                'sTitle2 = "公司行業別"
                sql &= " ,kt1.TRADENAME MID " & vbCrLf
            Case cst_rpt_公司規模 '"SD_15_008_R18_2012"
                'sTitle2 = "公司規模"
                sql &= " ,case when ISNULL(stb.q5,0) =1 then '屬於中小企業' " & vbCrLf
                sql &= " when ISNULL(stb.q5,0) =0 then '非中小企業' " & vbCrLf
                sql &= " else '未選擇' end MID " & vbCrLf
            Case cst_rpt_參訓動機 '"SD_15_008_R19_2012"
                'sTitle2 = "參訓動機"
                sql &= " ,case when ISNULL(stb2.q2,0) =1 then '為補充與原專長相關之技能' " & vbCrLf
                sql &= " when ISNULL(stb2.q2,0) =2 then '轉換其他行職業所需技能' " & vbCrLf
                sql &= " when ISNULL(stb2.q2,0) =3 then '拓展工作領域及視野' " & vbCrLf
                sql &= " when ISNULL(stb2.q2,0) =4 then '其他' " & vbCrLf
                sql &= " else '未選擇' end MID " & vbCrLf
            Case cst_rpt_訓後動向 '"SD_15_008_R20_2012"
                'sTitle2 = "訓後動向"
                sql &= " ,case when ISNULL(stb.q3,0) =1 then '轉換工作' " & vbCrLf
                sql &= " when ISNULL(stb.q3,0) =2 then '留任' " & vbCrLf
                sql &= " when ISNULL(stb.q3,0) =3 then '其他' " & vbCrLf
                sql &= " else '未選擇' end MID " & vbCrLf
            Case cst_rpt_參訓單位類別 '"SD_15_008_R21_2012"
                'sTitle2 = "參訓單位類別"
                sql &= " ,ko.NAME MID " & vbCrLf
            Case cst_rpt_參加課程職能別 '"SD_15_008_R22_2012"
                'sTitle2 = "參加課程職能別"
                sql &= " ,kcc.CCName MID " & vbCrLf
            Case cst_rpt_參加課程型態別 '"SD_15_008_R23_2012"
                'sTitle2 = "參加課程型態別"'參加課程型態
                sql &= " ,case when pp.IsBusiness = 'Y' then '企業包班' " & vbCrLf
                sql &= " when pp.PointYN = 'Y' then '學分班' " & vbCrLf
                sql &= " when pp.PointYN = 'N'  then '非學分班' " & vbCrLf
                sql &= " else '未選擇' end MID " & vbCrLf
            Case cst_rpt_訓練業別 '"SD_15_008_R24_2012"
                'sTitle2 = "訓練業別"
                sql &= " ,ISNULL(ig.PGOVCLASS,ISNULL(ig2.GCODE2+':'+ig2.PCNAME+'-'+ig2.CNAME,ig3.GCODE2+':'+ig3.PNAME+'-'+ig3.CNAME)) MID " & vbCrLf
            Case cst_rpt_不設定
                sql &= " ,'無' MID" & vbCrLf
        End Select

        sql &= " ,case when a.Q1=1 then 1 end a11" & vbCrLf
        sql &= " ,case when a.Q1=2 then 1 end a12" & vbCrLf
        sql &= " ,case when a.Q1=3 then 1 end a13" & vbCrLf
        sql &= " ,case when a.Q1=4 then 1 end a14" & vbCrLf
        sql &= " ,case when a.Q1=5 then 1 end a15" & vbCrLf
        sql &= " ,case when a.Q1=6 then 1 end a16" & vbCrLf

        sql &= " ,case when a.Q2=1 then 1 end a21" & vbCrLf
        sql &= " ,case when a.Q2=2 then 1 end a22" & vbCrLf
        sql &= " ,case when a.Q2=3 then 1 end a23" & vbCrLf
        sql &= " ,case when a.Q2=4 then 1 end a24" & vbCrLf
        sql &= " ,case when a.Q2=5 then 1 end a25" & vbCrLf

        sql &= " ,case when a.Q3=1 then 1 end a31" & vbCrLf
        sql &= " ,case when a.Q3=2 then 1 end a32" & vbCrLf
        sql &= " ,case when a.Q3=3 then 1 end a33" & vbCrLf
        sql &= " ,case when a.Q3=4 then 1 end a34" & vbCrLf

        'sql += " ,case when a.Q3=5 then 1 end a35" & vbCrLf
        sql &= " ,case when a.Q4=1 then 1 end a41" & vbCrLf
        sql &= " ,case when a.Q4=2 then 1 end a42" & vbCrLf
        sql &= " ,case when a.Q4=3 then 1 end a43" & vbCrLf
        sql &= " ,case when a.Q4=4 then 1 end a44" & vbCrLf
        sql &= " ,case when a.Q4=5 then 1 end a45" & vbCrLf

        sql &= " ,case when a.Q5=1 then 1 end a51" & vbCrLf
        sql &= " ,case when a.Q5=2 then 1 end a52" & vbCrLf
        sql &= " ,case when a.Q5=3 then 1 end a53" & vbCrLf
        sql &= " ,case when a.Q5=4 then 1 end a54" & vbCrLf
        sql &= " ,case when a.Q5=5 then 1 end a55" & vbCrLf

        sql &= " ,case when a.Q8=1 then 1 end a61" & vbCrLf
        sql &= " ,case when a.Q8=2 then 1 end a62" & vbCrLf
        sql &= " ,case when a.Q8=3 then 1 end a63" & vbCrLf
        sql &= " ,case when a.Q8=4 then 1 end a64" & vbCrLf
        sql &= " ,case when a.Q8=5 then 1 end a65" & vbCrLf

        sql &= " ,case when a.Q7MR1='Y' then 1 end a71" & vbCrLf
        sql &= " ,case when a.Q7MR2='Y' then 1 end a72" & vbCrLf
        sql &= " ,case when a.Q7MR3='Y' then 1 end a73" & vbCrLf
        sql &= " ,case when a.Q7MR4='Y' then 1 end a74" & vbCrLf

        For ix8 As Integer = 1 To 8
            sql &= " ,case when a.Q21" & ix8 & "=1 then 1 end a21" & ix8 & "1" & vbCrLf
            sql &= " ,case when a.Q21" & ix8 & "=2 then 1 end a21" & ix8 & "2" & vbCrLf
            sql &= " ,case when a.Q21" & ix8 & "=3 then 1 end a21" & ix8 & "3" & vbCrLf
            sql &= " ,case when a.Q21" & ix8 & "=4 then 1 end a21" & ix8 & "4" & vbCrLf
            sql &= " ,case when a.Q21" & ix8 & "=5 then 1 end a21" & ix8 & "5" & vbCrLf
        Next
        For ix6 As Integer = 1 To 6
            sql &= " ,case when a.Q22" & ix6 & "=1 then 1 end a22" & ix6 & "1" & vbCrLf
            sql &= " ,case when a.Q22" & ix6 & "=2 then 1 end a22" & ix6 & "2" & vbCrLf
            sql &= " ,case when a.Q22" & ix6 & "=3 then 1 end a22" & ix6 & "3" & vbCrLf
            sql &= " ,case when a.Q22" & ix6 & "=4 then 1 end a22" & ix6 & "4" & vbCrLf
            sql &= " ,case when a.Q22" & ix6 & "=5 then 1 end a22" & ix6 & "5" & vbCrLf
        Next

        sql &= " FROM STUD_QUESTIONFIN A " & vbCrLf
        sql &= " JOIN CLASS_STUDENTSOFCLASS cs ON cs.SOCID=A.SOCID " & vbCrLf
        sql &= " JOIN STUD_STUDENTINFO ss ON ss.SID=cs.SID " & vbCrLf
        sql &= " JOIN STUD_SUBDATA ss2 ON ss2.SID=cs.SID " & vbCrLf
        sql &= " JOIN CLASS_CLASSINFO cc ON cc.OCID =cs.OCID" & vbCrLf
        sql &= " JOIN PLAN_PLANINFO PP ON PP.PLANID =CC.PLANID  AND PP.COMIDNO= CC.COMIDNO AND PP.SEQNO = CC.SEQNO " & vbCrLf
        sql &= " JOIN ID_PLAN ip ON ip.PlanID=cc.PlanID" & vbCrLf
        sql &= " JOIN KEY_PLAN kp ON kp.TPlanID =ip.TPlanID" & vbCrLf
        sql &= " JOIN ORG_ORGINFO oo ON oo.comidno=pp.comidno" & vbCrLf
        sql &= " JOIN VIEW_ZIPNAME iz ON iz.zipCode=ss2.zipCode1 " & vbCrLf
        'fix 無法解析 equal to 作業中 "Chinese_Taiwan_Stroke_CS_AS" 與 "Chinese_Taiwan_Stroke_CI_AS" 之間的定序衝突
        sql &= " JOIN V_ORGKIND1 cj1 ON cj1.VALUE COLLATE Chinese_Taiwan_Stroke_CS_AS=oo.ORGKIND2" & vbCrLf

        sql &= " LEFT JOIN VIEW_GOVCLASSCAST ig ON ig.GCID=pp.GCID" & vbCrLf
        sql &= " LEFT JOIN V_GOVCLASSCAST2 ig2 on ig2.GCID2=pp.GCID2" & vbCrLf
        sql &= " LEFT JOIN V_GOVCLASSCAST3 ig3 ON ig3.GCID3=pp.GCID3" & vbCrLf
        sql &= " LEFT JOIN STUD_TRAINBG stb ON stb.SOCID =cs.SOCID" & vbCrLf

        Select Case filename
            Case "SD_15_008_R11_2012"
                'sTitle2 = "性別"
            Case "SD_15_008_R12_2012"
                'sTitle2 = "年齡"
            Case "SD_15_008_R13_2012"
                'sTitle2 = "教育程度"
            Case "SD_15_008_R14_2012"
                'sTitle2 = "身分別"
            Case "SD_15_008_R15_2012"
                'sTitle2 = "工作年資"
            Case "SD_15_008_R16_2012"
                'sTitle2 = "地理分佈"
            Case "SD_15_008_R17_2012"
                'sTitle2 = "公司行業別"
            Case "SD_15_008_R18_2012"
                'sTitle2 = "公司規模"
            Case cst_rpt_參訓動機 '"SD_15_008_R19_2012"
                'sTitle2 = "參訓動機" '多筆(複選)
                sql &= " LEFT JOIN Stud_TrainBGQ2 stb2 ON stb2.SOCID =cs.SOCID" & vbCrLf
            Case "SD_15_008_R20_2012"
                'sTitle2 = "訓後動向"
            Case "SD_15_008_R21_2012"
                'sTitle2 = "參訓單位類別"
            Case "SD_15_008_R22_2012"
                'sTitle2 = "參加課程職能"
            Case "SD_15_008_R23_2012"
                'sTitle2 = "參加課程型態"
            Case "SD_15_008_R24_2012"
                'sTitle2 = "訓練業別"
            Case "SD_15_008_R25_2012"
                'sTitle2 = "不設定"
        End Select
        'Key_OrgType
        sql &= " LEFT JOIN KEY_ORGTYPE ko ON ko.OrgTypeID =oo.OrgKind" & vbCrLf
        'Key_ClassCatelog
        sql &= " LEFT JOIN KEY_CLASSCATELOG kcc ON kcc.ccid =pp.ClassCate " & vbCrLf
        sql &= " LEFT JOIN KEY_TRADE kt1 ON kt1.TRADEID =stb.Q4" & vbCrLf
        sql &= " LEFT JOIN KEY_DEGREE kd ON kd.DegreeID =ss.DegreeID" & vbCrLf
        sql &= " LEFT JOIN KEY_IDENTITY kd2 ON kd2.IdentityID =cs.MIdentityID" & vbCrLf
        sql &= " LEFT JOIN STUD_SUBSIDYCOST sc ON sc.SOCID =cs.SOCID" & vbCrLf
        'sql += " cross join (" & vbCrLf
        'sql += "   select v1.name from V_ORGKIND1 v1 WHERE v1.VALUE='" & SearchPlan & "'" & vbCrLf
        'sql += " ) cj1" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        'STUD_QUESTIONFIN
        '201608 加入排除條件 AMU
        'A.排除離退訓(離退訓作業功能)
        'B.排除有結訓未申請(補助申請功能)
        'C.排除審核不通過(補助審核功能)的學員
        sql &= " and cs.STUDSTATUS NOT IN (2,3)" & vbCrLf '非離退
        sql &= " and sc.SOCID IS NOT NULL" & vbCrLf '有申請資料
        sql &= " and ISNULL(sc.AppliedStatusM,'Y')='Y'" & vbCrLf '審核通過 或申請中的

        'G/W 產投/自主
        Select Case SearchPlan
            Case "G", "W"
                sql &= " AND oo.OrgKind2 ='" & SearchPlan & "'"
        End Select
        If RID <> "" Then
            sql &= " AND cc.RID LIKE '" & RID & "%'" & vbCrLf
        End If
        If OCID <> "" Then
            sql &= " AND cc.OCID ='" & OCID & "'" & vbCrLf
        End If
        If TPlanID <> "" Then
            sql &= " AND ip.TPlanID='" & TPlanID & "'" & vbCrLf
        End If
        If Years <> "" Then
            sql &= " AND ip.Years='" & Years & "'" & vbCrLf
        End If
        If PackageType <> "" Then
            sql &= " AND pp.PackageType ='" & PackageType & "'" & vbCrLf
        End If
        'sql += " and cc.RID LIKE 'D2479%'" & vbCrLf
        'sql += " and cc.OCID='74870'" & vbCrLf
        'sql += " AND ip.Years='2015'" & vbCrLf
        sql &= " ) G" & vbCrLf
        'sql += " GROUP BY G.PName,G.MID,G.PlanKind,G.SORT1" & vbCrLf
        sql &= " GROUP BY G.PName,G.MID,G.PlanKind " & vbCrLf
        'sql &= " ORDER BY NLSSORT(G.MID,'NLS_SORT=TCHINESE_STROKE_M')" & vbCrLf
        sql &= " order by G.MID "
        Dim sCmd As New SqlCommand(sql, objconn)

        TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            dt.Load(.ExecuteReader())
        End With

        If TIMS.dtHaveDATA(dt) Then
            Dim dr As DataRow = dt.Rows(0)
            PName = Convert.ToString(dr("PName"))
            plankind = Convert.ToString(dr("plankind"))
        End If

        Return dt
    End Function

    '組合 題目與答案1
    Sub SetRowColVal1(ByRef iRow As Integer, ByRef iCol As Integer, ByVal tmpVal1 As String, ByVal iType As Integer)
        Select Case iType
            Case 0
                If iRow <> 0 Then iRow += 1
                iCol = 0
            Case 1
                iCol += 1
            Case 2
                iRow += 1
        End Select
        arrItem(iRow, iCol) = tmpVal1
    End Sub

    '組合 題目與答案2
    Sub SetRowColVal1(ByRef iRow As Integer, ByRef iCol As Integer,
                      ByVal tmpVal1 As String, ByVal iType As Integer, ByVal iSpan As Integer)
        SetRowColVal1(iRow, iCol, tmpVal1, iType)
        '處理合併表格
        arrSpanRow(iRow) = iSpan
    End Sub

    '組合 題目與答案 
    Private Sub setItemArray()
        Dim aryX1 As String()
        '題目與答案
        '1
        'Dim tmpVal1 As String = "一、學員部分"
        Const cst_tmpVa11 As String = "(一)請問您目前的就業狀況為何？"
        Const cst_tmpVa12 As String = "(二)請問您的薪資於結訓後有提升嗎?"
        Const cst_tmpVa13 As String = "(三)請問您擔任的職位有變化嗎?"
        Const cst_tmpVa14 As String = "(四)請問您對目前工作的滿意度是否有變化?"
        Const cst_tmpVa15 As String = "(五)請問您目前工作內容是否與本次參訓課程有相關?"
        Const cst_tmpVa16 As String = "(六)請問您是否有繼續參與本方案的意願?"
        Const cst_tmpVa17 As String = "(七)結訓後是否有與下列人員保持聯絡?（可複選）"
        'Const cst_tmpVal1 As String = "二、訓練成效"
        'Const cst_tmpVal1 As String = "題目"
        'Const cst_tmpVal21 As String = "(一)訓練技能運用"
        Const cst_tmpVa21 As String = "1.參加訓練後，對工作能力更有信心"
        Const cst_tmpVa22 As String = "2.參加訓練後，有助於提升工作技能"
        Const cst_tmpVa23 As String = "3.參加訓練後，有助於提升工作效率"
        Const cst_tmpVa24 As String = "4.參加訓練後，能增進我的問題解決能力"
        Const cst_tmpVa25 As String = "5.參加訓練後，能將所學應用到工作上"
        Const cst_tmpVa26 As String = "6.參加訓練後，能將所學應用於日常生活中"
        Const cst_tmpVa27 As String = "7.是否同意參加訓練對第二專長有幫助"
        Const cst_tmpVa28 As String = "8.是否同意參加訓練對目前工作表現有幫助"
        'Const cst_tmpVal1 As String = "(二)訓練成果"
        Const cst_tmpVa31 As String = "1.參加訓練後，有助於提升我的績效考核"
        Const cst_tmpVa32 As String = "2.參加訓練後，有助於薪資的調升"
        Const cst_tmpVa33 As String = "3.參加訓練後，有助於職位的升遷"
        Const cst_tmpVa34 As String = "4.參加訓練後，有助於獲得證照"
        Const cst_tmpVa35 As String = "5.參加訓練後，有助於發展職涯"
        Const cst_tmpVa36 As String = "6.參加訓練後，有助於強化個人職場競爭力"

        Dim iRow As Integer = 0
        Dim iCol As Integer = 0
        SetRowColVal1(iRow, iCol, cst_tmpVa11, 0, 6)
        aryX1 = Split(",留任原公司,轉換至同產業的公司,轉換至不同產業的公司,創業（a.實體 b.網路c.兩者皆有）,已離職，待業中,其他", ",")
        SetRowColVal1(iRow, iCol, aryX1(1), 1)
        SetRowColVal1(iRow, iCol, aryX1(2), 2)
        SetRowColVal1(iRow, iCol, aryX1(3), 2)
        SetRowColVal1(iRow, iCol, aryX1(4), 2)
        SetRowColVal1(iRow, iCol, aryX1(5), 2)
        SetRowColVal1(iRow, iCol, aryX1(6), 2)

        SetRowColVal1(iRow, iCol, cst_tmpVa12, 0, 5)
        aryX1 = Split(",大幅提升,小幅提升,沒有變化,小幅減少,大幅減少", ",")
        SetRowColVal1(iRow, iCol, aryX1(1), 1)
        SetRowColVal1(iRow, iCol, aryX1(2), 2)
        SetRowColVal1(iRow, iCol, aryX1(3), 2)
        SetRowColVal1(iRow, iCol, aryX1(4), 2)
        SetRowColVal1(iRow, iCol, aryX1(5), 2)

        SetRowColVal1(iRow, iCol, cst_tmpVa13, 0, 4)
        aryX1 = Split(",升職,調職,沒有變化,降職", ",")
        SetRowColVal1(iRow, iCol, aryX1(1), 1)
        SetRowColVal1(iRow, iCol, aryX1(2), 2)
        SetRowColVal1(iRow, iCol, aryX1(3), 2)
        SetRowColVal1(iRow, iCol, aryX1(4), 2)

        SetRowColVal1(iRow, iCol, cst_tmpVa14, 0, 5)
        aryX1 = Split(",大幅提升,小幅提升,沒有變化,小幅減少,大幅減少", ",")
        SetRowColVal1(iRow, iCol, aryX1(1), 1)
        SetRowColVal1(iRow, iCol, aryX1(2), 2)
        SetRowColVal1(iRow, iCol, aryX1(3), 2)
        SetRowColVal1(iRow, iCol, aryX1(4), 2)
        SetRowColVal1(iRow, iCol, aryX1(5), 2)

        SetRowColVal1(iRow, iCol, cst_tmpVa15, 0, 5)
        aryX1 = Split(",非常相關,相關,尚可,不相關,非常不相關", ",")
        SetRowColVal1(iRow, iCol, aryX1(1), 1)
        SetRowColVal1(iRow, iCol, aryX1(2), 2)
        SetRowColVal1(iRow, iCol, aryX1(3), 2)
        SetRowColVal1(iRow, iCol, aryX1(4), 2)
        SetRowColVal1(iRow, iCol, aryX1(5), 2)

        SetRowColVal1(iRow, iCol, cst_tmpVa16, 0, 5)
        aryX1 = Split(",非常想參與,想參與,尚可,不想參與,非常不想參與", ",")
        SetRowColVal1(iRow, iCol, aryX1(1), 1)
        SetRowColVal1(iRow, iCol, aryX1(2), 2)
        SetRowColVal1(iRow, iCol, aryX1(3), 2)
        SetRowColVal1(iRow, iCol, aryX1(4), 2)
        SetRowColVal1(iRow, iCol, aryX1(5), 2)

        SetRowColVal1(iRow, iCol, cst_tmpVa17, 0, 4)
        aryX1 = Split(",講師,學員,工作人員,無", ",")
        SetRowColVal1(iRow, iCol, aryX1(1), 1)
        SetRowColVal1(iRow, iCol, aryX1(2), 2)
        SetRowColVal1(iRow, iCol, aryX1(3), 2)
        SetRowColVal1(iRow, iCol, aryX1(4), 2)

        aryX1 = Split(",非常同意,同意,普通,不同意,非常不同意", ",")
        SetRowColVal1(iRow, iCol, cst_tmpVa21, 0, 5)
        SetRowColVal1(iRow, iCol, aryX1(1), 1)
        SetRowColVal1(iRow, iCol, aryX1(2), 2)
        SetRowColVal1(iRow, iCol, aryX1(3), 2)
        SetRowColVal1(iRow, iCol, aryX1(4), 2)
        SetRowColVal1(iRow, iCol, aryX1(5), 2)

        SetRowColVal1(iRow, iCol, cst_tmpVa22, 0, 5)
        SetRowColVal1(iRow, iCol, aryX1(1), 1)
        SetRowColVal1(iRow, iCol, aryX1(2), 2)
        SetRowColVal1(iRow, iCol, aryX1(3), 2)
        SetRowColVal1(iRow, iCol, aryX1(4), 2)
        SetRowColVal1(iRow, iCol, aryX1(5), 2)

        SetRowColVal1(iRow, iCol, cst_tmpVa23, 0, 5)
        SetRowColVal1(iRow, iCol, aryX1(1), 1)
        SetRowColVal1(iRow, iCol, aryX1(2), 2)
        SetRowColVal1(iRow, iCol, aryX1(3), 2)
        SetRowColVal1(iRow, iCol, aryX1(4), 2)
        SetRowColVal1(iRow, iCol, aryX1(5), 2)

        SetRowColVal1(iRow, iCol, cst_tmpVa24, 0, 5)
        SetRowColVal1(iRow, iCol, aryX1(1), 1)
        SetRowColVal1(iRow, iCol, aryX1(2), 2)
        SetRowColVal1(iRow, iCol, aryX1(3), 2)
        SetRowColVal1(iRow, iCol, aryX1(4), 2)
        SetRowColVal1(iRow, iCol, aryX1(5), 2)

        SetRowColVal1(iRow, iCol, cst_tmpVa25, 0, 5)
        SetRowColVal1(iRow, iCol, aryX1(1), 1)
        SetRowColVal1(iRow, iCol, aryX1(2), 2)
        SetRowColVal1(iRow, iCol, aryX1(3), 2)
        SetRowColVal1(iRow, iCol, aryX1(4), 2)
        SetRowColVal1(iRow, iCol, aryX1(5), 2)

        SetRowColVal1(iRow, iCol, cst_tmpVa26, 0, 5)
        SetRowColVal1(iRow, iCol, aryX1(1), 1)
        SetRowColVal1(iRow, iCol, aryX1(2), 2)
        SetRowColVal1(iRow, iCol, aryX1(3), 2)
        SetRowColVal1(iRow, iCol, aryX1(4), 2)
        SetRowColVal1(iRow, iCol, aryX1(5), 2)

        SetRowColVal1(iRow, iCol, cst_tmpVa27, 0, 5)
        SetRowColVal1(iRow, iCol, aryX1(1), 1)
        SetRowColVal1(iRow, iCol, aryX1(2), 2)
        SetRowColVal1(iRow, iCol, aryX1(3), 2)
        SetRowColVal1(iRow, iCol, aryX1(4), 2)
        SetRowColVal1(iRow, iCol, aryX1(5), 2)

        SetRowColVal1(iRow, iCol, cst_tmpVa28, 0, 5)
        SetRowColVal1(iRow, iCol, aryX1(1), 1)
        SetRowColVal1(iRow, iCol, aryX1(2), 2)
        SetRowColVal1(iRow, iCol, aryX1(3), 2)
        SetRowColVal1(iRow, iCol, aryX1(4), 2)
        SetRowColVal1(iRow, iCol, aryX1(5), 2)

        aryX1 = Split(",非常同意,同意,普通,不同意,非常不同意", ",")
        SetRowColVal1(iRow, iCol, cst_tmpVa31, 0, 5)
        SetRowColVal1(iRow, iCol, aryX1(1), 1)
        SetRowColVal1(iRow, iCol, aryX1(2), 2)
        SetRowColVal1(iRow, iCol, aryX1(3), 2)
        SetRowColVal1(iRow, iCol, aryX1(4), 2)
        SetRowColVal1(iRow, iCol, aryX1(5), 2)

        SetRowColVal1(iRow, iCol, cst_tmpVa32, 0, 5)
        SetRowColVal1(iRow, iCol, aryX1(1), 1)
        SetRowColVal1(iRow, iCol, aryX1(2), 2)
        SetRowColVal1(iRow, iCol, aryX1(3), 2)
        SetRowColVal1(iRow, iCol, aryX1(4), 2)
        SetRowColVal1(iRow, iCol, aryX1(5), 2)

        SetRowColVal1(iRow, iCol, cst_tmpVa33, 0, 5)
        SetRowColVal1(iRow, iCol, aryX1(1), 1)
        SetRowColVal1(iRow, iCol, aryX1(2), 2)
        SetRowColVal1(iRow, iCol, aryX1(3), 2)
        SetRowColVal1(iRow, iCol, aryX1(4), 2)
        SetRowColVal1(iRow, iCol, aryX1(5), 2)

        SetRowColVal1(iRow, iCol, cst_tmpVa34, 0, 5)
        SetRowColVal1(iRow, iCol, aryX1(1), 1)
        SetRowColVal1(iRow, iCol, aryX1(2), 2)
        SetRowColVal1(iRow, iCol, aryX1(3), 2)
        SetRowColVal1(iRow, iCol, aryX1(4), 2)
        SetRowColVal1(iRow, iCol, aryX1(5), 2)

        SetRowColVal1(iRow, iCol, cst_tmpVa35, 0, 5)
        SetRowColVal1(iRow, iCol, aryX1(1), 1)
        SetRowColVal1(iRow, iCol, aryX1(2), 2)
        SetRowColVal1(iRow, iCol, aryX1(3), 2)
        SetRowColVal1(iRow, iCol, aryX1(4), 2)
        SetRowColVal1(iRow, iCol, aryX1(5), 2)

        SetRowColVal1(iRow, iCol, cst_tmpVa36, 0, 5)
        SetRowColVal1(iRow, iCol, aryX1(1), 1)
        SetRowColVal1(iRow, iCol, aryX1(2), 2)
        SetRowColVal1(iRow, iCol, aryX1(3), 2)
        SetRowColVal1(iRow, iCol, aryX1(4), 2)
        SetRowColVal1(iRow, iCol, aryX1(5), 2)

    End Sub

    '列印隱鈕
    Protected Sub btnPrt_Click(sender As Object, e As EventArgs) Handles btnPrt.Click
        Call Create11()
        trBtn.Attributes.Remove("style")
        trBtn.Attributes.Add("style", "display:none")

        Select Case filename
            Case cst_rpt_性別 '"SD_15_008_R11_2012"
                Call exc_Print("true")
            Case cst_rpt_年齡 '"SD_15_008_R12_2012"
                Call exc_Print("false")
            Case cst_rpt_教育程度 '"SD_15_008_R13_2012"
                Call exc_Print("false")
            Case cst_rpt_身分別 '"SD_15_008_R14_2012"
                Call exc_Print("false")
            Case cst_rpt_工作年資 '"SD_15_008_R15_2012"
                Call exc_Print("true")
            Case cst_rpt_地理分佈 '"SD_15_008_R16_2012"
                Call exc_Print("false")
            Case cst_rpt_公司行業別 '"SD_15_008_R17_2012"
                Call exc_Print("false")
            Case cst_rpt_公司規模 '"SD_15_008_R18_2012"
                Call exc_Print("true")
            Case cst_rpt_參訓動機 '"SD_15_008_R19_2012"
                Call exc_Print("true")
            Case cst_rpt_訓後動向 '"SD_15_008_R20_2012"
                Call exc_Print("true")
            Case cst_rpt_參訓單位類別 '"SD_15_008_R21_2012"
                Call exc_Print("false")
            Case cst_rpt_參加課程職能別 '"SD_15_008_R22_2012"
                Call exc_Print("true")
            Case cst_rpt_參加課程型態別 '"SD_15_008_R23_2012"
                Call exc_Print("true")
            Case cst_rpt_訓練業別 '"SD_15_008_R24_2012"
                Call exc_Print("false")
            Case cst_rpt_不設定 '"SD_15_008_R25_2012"
                Call exc_Print("true")
        End Select
    End Sub

    '列印程式
    Private Sub PrintDiv(ByVal dt As DataTable,
                         ByVal selRpt As String, ByVal Field1_width As String,
                         ByVal Field2_width As String, ByVal RCount As Integer,
                         ByVal font_size As String, ByVal portrait As String)
        'dt:要顯示的資料,selRpt:
        ',Field1_width:標題的寬度
        ',Field2_width:選項的寬度
        ',RCount:每頁筆數
        ',font_size:內容字型大小
        ',portrait:直式/橫式

        'Dim tmpDR As DataRow
        'Dim tmpObj As Object
        'Dim tmpDT As New DataTable
        'Dim sql As String = ""
        If TIMS.dtNODATA(gResDt) Then
            Common.MessageBox(Me, "查無資料!!")
            Exit Sub
        End If

        Dim PageCount As Int32 = 0  'Pages
        Dim ReportCount As Integer = RCount '每頁筆數
        'Dim ColCount As Integer = 0
        Dim intTmp As Integer = 0
        Dim rsCursor As Integer = 0   '報表內容列印的NO
        Dim intPageRecord As Integer = RCount '每頁列印幾筆

        Dim nt As HtmlTable
        Dim nr As HtmlTableRow
        Dim nc As HtmlTableCell
        Dim nl As HtmlGenericControl
        Dim strStyle As String = "font-size:" + font_size + "pt;font-family:DFKai-SB"
        Dim int_width As Integer
        Dim strRowHeight As String = ""

        'tmpDT = dt 'ColCount = dt.Columns.Count
        intTmp = dt.Rows.Count
        If (intTmp Mod ReportCount) = 0 Then
            PageCount = (intTmp / ReportCount) - 1
        Else
            PageCount = intTmp / ReportCount
        End If

        '表格寬度的設定
        If portrait = "H" Then
            int_width = Int((550 - Field1_width - Field2_width) / gResDt.Rows.Count)
        Else
            int_width = Int((820 - Field1_width - Field2_width) / gResDt.Rows.Count)
        End If

        '行高
        If portrait = "L" Then
            strRowHeight = "10"
        Else
            strRowHeight = "17"
        End If

        'Me.ViewState("xHtml") = "" .Rows.Count > 0 
        If TIMS.dtHaveDATA(dt) Then

            Dim iAColSpan As Integer = 0
            'iAColSpan = 2 + gResDt.Rows.Count + 1

            Select Case filename
                Case "SD_15_008_R25_2012"
                    'sTitle2 = "不設定" 'dt.Columns.Remove("MID")
                    iAColSpan = 2 + 0 + 1
                Case Else
                    iAColSpan = 2 + gResDt.Rows.Count + 1
            End Select

            For i As Integer = 0 To PageCount
                '表頭
                nt = New HtmlTable
                nt.Attributes.Add("style", "width:100%; BORDER-TOP-STYLE: none;FONT-FAMILY: 標楷體;BORDER-RIGHT-STYLE: none;BORDER-LEFT-STYLE: none;BORDER-COLLAPSE: collapse;BORDER-BOTTOM-STYLE: none")
                nt.Attributes.Add("align", "center")
                nt.Attributes.Add("border", "0")
                div_print.Controls.Add(nt)

                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "100%")
                nc.Attributes.Add("colspan", iAColSpan)
                nc.Attributes.Add("style", "font-size:14pt;font-family:DFKai-SB")
                nc.InnerHtml = ""

                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "100%")
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("colspan", iAColSpan)
                nc.Attributes.Add("style", "font-size:14pt;font-family:DFKai-SB")
                nc.InnerHtml = "勞動部勞動力發展署"

                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "80%")
                nc.Attributes.Add("align", "left")
                nc.Attributes.Add("colspan", 2)
                nc.Attributes.Add("style", "font-size:13pt;font-family:DFKai-SB")
                nc.InnerHtml = plankind

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "20%")
                nc.Attributes.Add("align", "right")
                nc.Attributes.Add("colspan", iAColSpan - 2)
                nc.Attributes.Add("style", "font-size:11pt;font-family:DFKai-SB")
                nc.InnerHtml = "列印日期：" + Now().ToShortDateString()

                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "80%")
                nc.Attributes.Add("align", "left")
                nc.Attributes.Add("colspan", 2)
                nc.Attributes.Add("style", "font-size:13pt;font-family:DFKai-SB")
                nc.InnerHtml = PName

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "20%")
                nc.Attributes.Add("align", "right")
                nc.Attributes.Add("colspan", iAColSpan - 2)
                nc.Attributes.Add("style", "font-size:11pt;font-family:DFKai-SB")
                nc.InnerHtml = "頁數：" + (i + 1).ToString + " / " + PageCount.ToString

                'Column Header
                nt = New HtmlTable
                nt.Attributes.Add("style", "width:100%; BORDER-TOP-STYLE: none;FONT-FAMILY: 標楷體;BORDER-RIGHT-STYLE: none;BORDER-LEFT-STYLE: none;BORDER-COLLAPSE: collapse;BORDER-BOTTOM-STYLE: none")
                nt.Attributes.Add("align", "center")
                nt.Attributes.Add("border", "2")
                div_print.Controls.Add(nt)

                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", strStyle)
                nc.Attributes.Add("colspan", "2") '題目，選項空白行
                nc.InnerHtml = "&nbsp;"

                Select Case filename
                    Case "SD_15_008_R25_2012"
                        'sTitle2 = "不設定" 'dt.Columns.Remove("MID")
                    Case Else
                End Select

                Select Case filename
                    Case "SD_15_008_R25_2012"
                        'sTitle2 = "不設定" 'dt.Columns.Remove("MID")
                    Case Else
                        For j As Integer = 0 To gResDt.Rows.Count - 1
                            nc = New HtmlTableCell
                            nr.Controls.Add(nc)
                            nc.Attributes.Add("align", "center")
                            nc.Attributes.Add("style", strStyle + ";word-break:break-all")
                            nc.Attributes.Add("width", int_width)
                            nc.InnerHtml = gResDt.Rows(j)("name").ToString
                        Next
                End Select

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", strStyle)
                nc.Attributes.Add("width", int_width)
                nc.InnerHtml = "總計"

                '報表內容
                Dim strTmp As String = ""
                Dim strAlign As String = ""
                Dim intPageRecord2 As Integer = intPageRecord

                Select Case i
                    Case 1
                        '第2頁 校正
                        intPageRecord2 = intPageRecord + 1
                End Select

                For j As Integer = 0 To intPageRecord2
                    If rsCursor >= dt.Rows.Count Then
                        GoTo [CONTINUE]
                    End If

                    nr = New HtmlTableRow
                    nt.Controls.Add(nr)

                    For m As Integer = 0 To dt.Columns.Count - 1

                        If m = 0 Or m = 1 Then
                            strAlign = "left"
                            If m = 0 Then

                                If UBound(arrSpanRow) >= rsCursor Then
                                    If Not IsNothing(arrSpanRow(rsCursor)) Then
                                        nc = New HtmlTableCell
                                        nr.Controls.Add(nc)
                                        '處理合併表格
                                        nc.Attributes.Add("rowspan", arrSpanRow(rsCursor))
                                        nc.Attributes.Add("width", Field1_width)
                                        strTmp = dt.Rows(rsCursor)(m).ToString
                                        nc.Attributes.Add("align", strAlign)
                                        nc.Attributes.Add("style", strStyle)
                                        nc.InnerHtml = strTmp
                                    End If
                                End If

                            Else
                                nc = New HtmlTableCell
                                nr.Controls.Add(nc)
                                nc.Attributes.Add("width", Field2_width)
                                strTmp = dt.Rows(rsCursor)(m).ToString
                                nc.Attributes.Add("align", strAlign)
                                nc.Attributes.Add("style", strStyle)
                                nc.Attributes.Add("height", strRowHeight)
                                nc.InnerHtml = strTmp
                            End If
                        Else
                            nc = New HtmlTableCell
                            nr.Controls.Add(nc)
                            strAlign = "right"
                            strTmp = dt.Rows(rsCursor)(m).ToString
                            nc.Attributes.Add("align", strAlign)
                            nc.Attributes.Add("style", strStyle)
                            nc.InnerHtml = strTmp
                        End If

                    Next
                    rsCursor += 1
                Next
                'Me.ViewState("xHtml") &= nt.InnerHtml

[CONTINUE]:
                '表尾

                If rsCursor + 1 > dt.Rows.Count Then
                    GoTo out
                End If
                '換頁列印
                nl = New HtmlGenericControl
                div_print.Controls.Add(nl)
                'nl.InnerHtml = "<p style='line-height:2px;margin:0cm;margin-bottom:0.0001pt;mso-pagination:widow-orphan;'><br clear=all style='mso-special-character:line-break;page-break-before:always'>"
                nl.InnerHtml = "<p style='line-height:2px;margin:0cm;margin-bottom:0.0001pt;mso-pagination:widow-orphan;'></p><br clear=all style='mso-special-character:line-break;page-break-before:always' />"
                'Me.ViewState("xHtml") &= nl.InnerHtml

            Next
out:
        End If

    End Sub

    '重組送入內容資訊。(陣列轉換)
    Function RotationDT(ByRef dt As DataTable) As DataTable
        'Threading.Thread.Sleep(1) '假設處理某段程序需花費1毫秒 (避免機器不同步)
        Const cst_iStart As Integer = 3 '答案起始位置
        'Dim tmpNames As String
        Dim tmpDT As New DataTable
        tmpDT.Columns.Add(New DataColumn("Item"))    '題目
        tmpDT.Columns.Add(New DataColumn("Ques"))    '答案

        Select Case filename
            Case "SD_15_008_R25_2012"
                'sTitle2 = "不設定" 'dt.Columns.Remove("MID")
            Case Else
                For j As Integer = 0 To dt.Rows.Count - 1
                    'Dim sName1 As String = Convert.ToString(dt.Rows(j)("name"))
                    'If tmpNames <> "" AndAlso tmpNames.IndexOf(sName1) > -1 Then
                    '    sName1 &= "2" '同名補2
                    'End If
                    'If tmpNames <> "" Then tmpNames &= ","
                    'tmpNames &= "'" & sName1 & "'"
                    'tmpDT.Columns.Add(New DataColumn(dt.Rows(j)("name")))
                    tmpDT.Columns.Add(New DataColumn(dt.Rows(j)("name")))
                Next
        End Select

        tmpDT.Columns.Add(New DataColumn("Total"))    '答案

        Dim iNo As Integer = 0
        For iA As Integer = cst_iStart To dt.Columns.Count - 1
            Dim tmpDR As DataRow = tmpDT.NewRow
            tmpDT.Rows.Add(tmpDR)

            Dim intTotal As Integer = 0
            '題目與答案
            If UBound(arrItem, 1) >= iA - cst_iStart Then
                If Not IsNothing(arrItem(iA - cst_iStart, 0)) Then tmpDR("item") = arrItem(iA - cst_iStart, 0).ToString
                If Not IsNothing(arrItem(iA - cst_iStart, 1)) Then tmpDR("Ques") = arrItem(iA - cst_iStart, 1).ToString
            End If

            Select Case filename
                Case "SD_15_008_R25_2012"
                    'sTitle2 = "不設定" 'dt.Columns.Remove("MID")
                    For j As Integer = 0 To dt.Rows.Count - 1
                        iNo = j + 2
                        'tmpDR(iNo) = dt.Rows(j)(iA)
                        intTotal += dt.Rows(j)(iA)
                    Next
                Case Else
                    For j As Integer = 0 To dt.Rows.Count - 1
                        iNo = j + 2
                        tmpDR(iNo) = dt.Rows(j)(iA)
                        intTotal += dt.Rows(j)(iA)
                    Next
            End Select
            tmpDR("Total") = intTotal
        Next

        Return tmpDT
    End Function

    '列印直橫~
    Private Sub exc_Print(ByVal portrait As String)
        Dim strScript As String = ""
        strScript = "<script language=""javascript"">window.print();</script>"
        Page.RegisterStartupScript("window_onload", strScript)
        Return

        'Dim strScript As String = ""
        'strScript = "<script language=""javascript"">" + vbCrLf
        'strScript += "if (!factory.object) {"
        'strScript += " window.print();"
        'strScript += "} else {"
        'strScript += "factory.printing.header = """";"
        'strScript += "factory.printing.footer = """";"
        'strScript += "factory.printing.leftMargin = 5; "
        'strScript += "factory.printing.topMargin = 10; "
        'strScript += "factory.printing.rightMargin = 5; "
        'strScript += "factory.printing.bottomMargin = 10; "
        'strScript += "factory.printing.portrait = " + portrait + ";"
        'strScript += "factory.printing.Print(true);"
        'strScript += "window.close();"
        'strScript += "}"
        'strScript += "</script>"
        'Page.RegisterStartupScript("window_onload", strScript)
    End Sub

    '覆寫，不執行 MyBase.VerifyRenderingInServerForm 方法，解決執行 RenderControl 產生的錯誤   
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)

    End Sub

    '匯山EXCEL
    Protected Sub BtnExcel_Click(sender As Object, e As EventArgs) Handles btnExcel.Click
        Call Create11()
        trBtn.Attributes.Remove("style")
        trBtn.Attributes.Add("style", "display:none")
        trBtn.Visible = False

        Dim fileName As String = "訓後動態調查統計表.xls"
        'If Request.Browser.Browser = "IE" Then
        '    fileName = Server.UrlPathEncode(fileName)
        'End If
        fileName = HttpUtility.UrlEncode(fileName, System.Text.Encoding.UTF8)

        Dim strContentDisposition As String = [String].Format("{0}; filename=""{1}""", "attachment", fileName)

        Response.Clear()
        Response.ClearHeaders()
        Response.Buffer = True
        Response.Charset = "UTF-8" '設定字集
        Response.AppendHeader("Content-Disposition", strContentDisposition)
        Response.ContentEncoding = System.Text.Encoding.GetEncoding("UTF-8")
        'Response.ContentType = "application/vnd.ms-excel " '內容型態設為Excel
        '文件內容指定為Excel
        Response.ContentType = "application/ms-excel;charset=utf-8"
        Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=UTF-8>")
        'Response.AddHeader("Content-Disposition", strContentDisposition)
        'Response.ContentType = "Application/vnd.ms-excel"
        Dim sw As New System.IO.StringWriter
        Dim htw As HtmlTextWriter = New HtmlTextWriter(sw)
        div_print.RenderControl(htw)
        Common.RespWrite(Me, sw.ToString().Replace("<div>", "").Replace("</div>", ""))
        TIMS.Utl_RespWriteEnd(Me, objconn, "") ' Response.End()
    End Sub

    Protected Sub btnExpOds1_Click(sender As Object, e As EventArgs) Handles btnExpOds1.Click
        Call Create11()
        trBtn.Attributes.Remove("style")
        trBtn.Attributes.Add("style", "display:none")
        trBtn.Visible = False

        Dim sw As New System.IO.StringWriter
        Dim htw As HtmlTextWriter = New HtmlTextWriter(sw)
        div_print.RenderControl(htw)

        Dim sFileName1 As String = "ExpFile" & TIMS.GetRnd6Eng()

        Dim strHTML As String = ""
        strHTML &= (sw.ToString().Replace("<div>", "").Replace("</div>", ""))

        'parmsExp.Add("xlsx_buf", buf)
        Dim parmsExp As New Hashtable From {
            {"ExpType", "ODS"}, 'EXCEL/PDF/ODS
            {"FileName", sFileName1},
            {"strHTML", strHTML},
            {"ResponseNoEnd", "Y"}
        }
        TIMS.Utl_ExportRp1(Me, parmsExp)
        TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
    End Sub
End Class