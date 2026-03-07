Public Class SD_15_008_R_Prt
    Inherits AuthBasePage

    '2016(Old)
    Dim arrItem(37, 2) As String '題目與答案
    Dim arrSpanRow(37) As String '處理合併表格
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
            'ClientScript.RegisterStartupScript(Me.GetType(), "", "<script>blockAlert('查無資料!!');</script>")
            Exit Sub
        End If

        If gResDt Is Nothing Then Exit Sub
        If gResDt.Rows.Count = 0 Then Exit Sub

        btnCancel.Attributes.Add("onclick", "window.close();")

        Dim tDt As DataTable = RotationDT(gResDt)
        Select Case filename
            Case "SD_15_008_R25_2012"
                'tDt.Columns.Remove("-")
                Call PrintDiv(tDt, filename, "200", "150", 45, "10", "H")
            Case "SD_15_008_R11_2012"
                Call PrintDiv(tDt, filename, "200", "150", 45, "10", "H")
                'exc_Print("true")
            Case "SD_15_008_R12_2012"
                'tDt = db_R12(TPlanID, Years, OCID, RID, SearchPlan, PackageType)
                Call PrintDiv(tDt, filename, "120", "120", 45, "7", "L")
                'exc_Print("false")
            Case "SD_15_008_R13_2012"
                'tDt = db_R13(TPlanID, Years, OCID, RID, SearchPlan, PackageType)
                Call PrintDiv(tDt, filename, "100", "110", 45, "7", "L")
                'exc_Print("false")
            Case "SD_15_008_R14_2012"
                'tDt = db_R14(TPlanID, Years, OCID, RID, SearchPlan, PackageType)
                Call PrintDiv(tDt, filename, "90", "65", 45, "5", "L")
                'exc_Print("false")
            Case "SD_15_008_R15_2012"
                'tDt = db_R15(TPlanID, Years, OCID, RID, SearchPlan, PackageType)
                Call PrintDiv(tDt, filename, "150", "150", 45, "10", "H")
                'exc_Print("true")
            Case "SD_15_008_R16_2012"
                'tDt = db_R16(TPlanID, Years, OCID, RID, SearchPlan, PackageType)
                Call PrintDiv(tDt, filename, "95", "70", 45, "6", "L")
                'exc_Print("false")
            Case "SD_15_008_R17_2012"
                'tDt = db_R17(TPlanID, Years, OCID, RID, SearchPlan, PackageType)
                Call PrintDiv(tDt, filename, "95", "70", 45, "6", "L")
                'exc_Print("false")
            Case "SD_15_008_R18_2012"
                'tDt = db_R18(TPlanID, Years, OCID, RID, SearchPlan, PackageType)
                Call PrintDiv(tDt, filename, "150", "150", 45, "10", "H")
                'exc_Print("true")
            Case "SD_15_008_R19_2012"
                'tDt = db_R19(TPlanID, Years, OCID, RID, SearchPlan, PackageType)
                Call PrintDiv(tDt, filename, "130", "140", 45, "10", "H")
                'exc_Print("true")
            Case "SD_15_008_R20_2012"
                'tDt = db_R20(TPlanID, Years, OCID, RID, SearchPlan, PackageType)
                Call PrintDiv(tDt, filename, "150", "150", 45, "10", "H")
                'exc_Print("true")
            Case "SD_15_008_R21_2012"
                'tDt = db_R21(TPlanID, Years, OCID, RID, SearchPlan, PackageType)
                Call PrintDiv(tDt, filename, "90", "100", 45, "7", "L")
                'exc_Print("false")
            Case "SD_15_008_R22_2012"
                'tDt = db_R22(TPlanID, Years, OCID, RID, SearchPlan, PackageType)
                Call PrintDiv(tDt, filename, "130", "120", 45, "10", "H")
                'exc_Print("true")
            Case "SD_15_008_R23_2012"
                'tDt = db_R23(TPlanID, Years, OCID, RID, SearchPlan, PackageType)
                Call PrintDiv(tDt, filename, "150", "150", 45, "10", "H")
                'exc_Print("true")
            Case "SD_15_008_R24_2012"
                'tDt = db_R24(TPlanID, Years, OCID, RID, SearchPlan, PackageType)
                Call PrintDiv(tDt, filename, "55", "70", 45, "6", "L")
                'exc_Print("false")
            Case Else
                Common.MessageBox(Me, "查無資料!!")
                Exit Sub
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
            Case "SD_15_008_R25_2012"
                sTitle2 = "不設定"
            Case "SD_15_008_R11_2012"
                sTitle2 = "性別"
            Case "SD_15_008_R12_2012"
                sTitle2 = "年齡"
            Case "SD_15_008_R13_2012"
                sTitle2 = "教育程度"
            Case "SD_15_008_R14_2012"
                sTitle2 = "身分別"
            Case "SD_15_008_R15_2012"
                sTitle2 = "工作年資"
            Case "SD_15_008_R16_2012"
                sTitle2 = "地理分佈"
            Case "SD_15_008_R17_2012"
                sTitle2 = "公司行業別"
            Case "SD_15_008_R18_2012"
                sTitle2 = "公司規模"
            Case "SD_15_008_R19_2012"
                sTitle2 = "參訓動機"
            Case "SD_15_008_R20_2012"
                sTitle2 = "訓後動向"
            Case "SD_15_008_R21_2012"
                sTitle2 = "參訓單位類別"
            Case "SD_15_008_R22_2012"
                'sTitle2 = "參加課程職能別"
                sTitle2 = "參加課程職能"
            Case "SD_15_008_R23_2012"
                'sTitle2 = "參加課程型態別"
                sTitle2 = "參加課程型態"
            Case "SD_15_008_R24_2012"
                sTitle2 = "訓練業別"
        End Select

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " SELECT G.PName,G.MID NAME,G.PlanKind" & vbCrLf
        '1.學員目前的近況為何
        sql += " ,ISNULL(COUNT(g.a11),0) a11,ISNULL(COUNT(g.a12),0) a12,ISNULL(COUNT(g.a13),0) a13,ISNULL(COUNT(g.a14),0) a14" & vbCrLf
        '2.學員於結訓後薪資有提升嗎
        sql += " ,ISNULL(COUNT(g.a21),0) a21,ISNULL(COUNT(g.a22),0) a22,ISNULL(COUNT(g.a23),0) a23,ISNULL(COUNT(g.a24),0) a24,ISNULL(COUNT(g.a25),0) a25" & vbCrLf
        '3.學員職位有變化嗎
        sql += " ,ISNULL(COUNT(g.a31),0) a31,ISNULL(COUNT(g.a32),0) a32,ISNULL(COUNT(g.a33),0) a33,ISNULL(COUNT(g.a34),0) a34" & vbCrLf
        '4.學員對目前工作的滿意度是否有變化
        sql += " ,ISNULL(COUNT(g.a41),0) a41,ISNULL(COUNT(g.a42),0) a42,ISNULL(COUNT(g.a43),0) a43,ISNULL(COUNT(g.a44),0) a44,ISNULL(COUNT(g.a45),0) a45" & vbCrLf
        '5.學員目前的工作內容是否與參訓課程內容相關
        sql += " ,ISNULL(COUNT(g.a51),0) a51,ISNULL(COUNT(g.a52),0) a52,ISNULL(COUNT(g.a53),0) a53,ISNULL(COUNT(g.a54),0) a54,ISNULL(COUNT(g.a55),0) a55" & vbCrLf
        'sql += " ,dbo.NVL(COUNT(g.a61),0) a61,dbo.NVL(COUNT(g.a62),0) a62,dbo.NVL(COUNT(g.a63),0) a63,dbo.NVL(COUNT(g.a64),0) a64,dbo.NVL(COUNT(g.a65),0) a65" & vbCrLf
        '6-1.學員是否同意參加訓練對目前工作表現有幫助
        sql += " ,ISNULL(COUNT(g.a6_71),0) a6_71,ISNULL(COUNT(g.a6_72),0) a6_72,ISNULL(COUNT(g.a6_73),0) a6_73,ISNULL(COUNT(g.a6_74),0) a6_74,ISNULL(COUNT(g.a6_75),0) a6_75" & vbCrLf
        '6-2.學員是否同意參加訓練對第二專長培育有幫助
        sql += " ,ISNULL(COUNT(g.a6_81),0) a6_81,ISNULL(COUNT(g.a6_82),0) a6_82,ISNULL(COUNT(g.a6_83),0) a6_83,ISNULL(COUNT(g.a6_84),0) a6_84,ISNULL(COUNT(g.a6_85),0) a6_85" & vbCrLf
        'sql += " ,dbo.NVL(COUNT(g.a71),0) a71,dbo.NVL(COUNT(g.a72),0) a72,dbo.NVL(COUNT(g.a73),0) a73" & vbCrLf
        '7.學員是否有繼續參與進修訓練的意願
        sql += " ,ISNULL(COUNT(g.a81),0) a81,ISNULL(COUNT(g.a82),0) a82,ISNULL(COUNT(g.a83),0) a83,ISNULL(COUNT(g.a84),0) a84,ISNULL(COUNT(g.a85),0) a85" & vbCrLf
        'sql += " ,G.SORT1" & vbCrLf
        'sql += " ,dbo.NVL(COUNT(1),0) CNT " & vbCrLf
        sql += " FROM (" & vbCrLf

        Select Case iType
            Case 1
                sql += " SELECT kp.PlanName+'('+cj1.name COLLATE Chinese_Taiwan_Stroke_CS_AS +')' PlanKind" & vbCrLf
            Case Else
                sql += " SELECT kp.PlanName PlanKind" & vbCrLf '合併為1筆
        End Select
        sql += " ,cast((ip.Years-1911) as varchar)+'年度　訓後動態調查統計表 " & sTitle2 & "' PName" & vbCrLf

        Select Case filename
            Case "SD_15_008_R25_2012"
                sql += " ,'-' MID" & vbCrLf
            Case "SD_15_008_R11_2012"
                'sTitle2 = "性別"
                sql += " ,dbo.DECODE6(ss.Sex,'M','男','F','女','無') MID" & vbCrLf
            Case "SD_15_008_R12_2012"
                'sTitle2 = "年齡"
                sql += " ,case when 15> DATEPART(YEAR, cc.STDATE)-DATEPART(YEAR, ss.Birthday) then '15以下' " & vbCrLf
                sql += " when 20> DATEPART(YEAR, cc.STDATE)-DATEPART(YEAR, ss.Birthday) and 15<=DATEPART(YEAR, cc.STDATE)-DATEPART(YEAR, ss.Birthday) then '15~19' " & vbCrLf
                sql += " when 25> DATEPART(YEAR, cc.STDATE)-DATEPART(YEAR, ss.Birthday) and 20<=DATEPART(YEAR, cc.STDATE)-DATEPART(YEAR, ss.Birthday) then '20~24' " & vbCrLf
                sql += " when 30> DATEPART(YEAR, cc.STDATE)-DATEPART(YEAR, ss.Birthday) and 25<=DATEPART(YEAR, cc.STDATE)-DATEPART(YEAR, ss.Birthday) then '25~29' " & vbCrLf
                sql += " when 35> DATEPART(YEAR, cc.STDATE)-DATEPART(YEAR, ss.Birthday) and 30<=DATEPART(YEAR, cc.STDATE)-DATEPART(YEAR, ss.Birthday) then '30~34' " & vbCrLf
                sql += " when 40> DATEPART(YEAR, cc.STDATE)-DATEPART(YEAR, ss.Birthday) and 35<=DATEPART(YEAR, cc.STDATE)-DATEPART(YEAR, ss.Birthday) then '35~39' " & vbCrLf
                sql += " when 45> DATEPART(YEAR, cc.STDATE)-DATEPART(YEAR, ss.Birthday) and 40<=DATEPART(YEAR, cc.STDATE)-DATEPART(YEAR, ss.Birthday) then '40~44' " & vbCrLf
                sql += " when 50> DATEPART(YEAR, cc.STDATE)-DATEPART(YEAR, ss.Birthday) and 45<=DATEPART(YEAR, cc.STDATE)-DATEPART(YEAR, ss.Birthday) then '45~49' " & vbCrLf
                sql += " when 55> DATEPART(YEAR, cc.STDATE)-DATEPART(YEAR, ss.Birthday) and 50<=DATEPART(YEAR, cc.STDATE)-DATEPART(YEAR, ss.Birthday) then '50~54' " & vbCrLf
                sql += " when 60> DATEPART(YEAR, cc.STDATE)-DATEPART(YEAR, ss.Birthday) and 55<=DATEPART(YEAR, cc.STDATE)-DATEPART(YEAR, ss.Birthday) then '55~59' " & vbCrLf
                sql += " else '60' end MID" & vbCrLf

            Case "SD_15_008_R13_2012"
                'sTitle2 = "教育程度"
                sql += " ,kd.Name MID" & vbCrLf
            Case "SD_15_008_R14_2012"
                'sTitle2 = "身分別"
                sql += " ,kd2.Name MID" & vbCrLf
            Case "SD_15_008_R15_2012"
                'sTitle2 = "工作年資"
                sql += " ,case when ISNULL(stb.q61,0) <3 then '3年以下' " & vbCrLf
                sql += " when ISNULL(stb.q61,0) <5 then '3~5年' " & vbCrLf
                sql += " when ISNULL(stb.q61,0)<10 then '5~10年' " & vbCrLf
                sql += " else '10年以上' end MID " & vbCrLf
            Case "SD_15_008_R16_2012"
                'sTitle2 = "地理分佈"
                sql += " ,iz.CTNAME MID " & vbCrLf
            Case "SD_15_008_R17_2012"
                'sTitle2 = "公司行業別"
                sql += " ,kt1.TRADENAME MID " & vbCrLf
            Case "SD_15_008_R18_2012"
                'sTitle2 = "公司規模"
                sql += " ,case when ISNULL(stb.q5,0) =1 then '屬於中小企業' " & vbCrLf
                sql += " when ISNULL(stb.q5,0) =0 then '非中小企業' " & vbCrLf
                sql += " else '未選擇' end MID " & vbCrLf
            Case "SD_15_008_R19_2012"
                'sTitle2 = "參訓動機"
                sql += " ,case when ISNULL(stb2.q2,0) =1 then '為補充與原專長相關之技能' " & vbCrLf
                sql += " when ISNULL(stb2.q2,0) =2 then '轉換其他行職業所需技能' " & vbCrLf
                sql += " when ISNULL(stb2.q2,0) =3 then '拓展工作領域及視野' " & vbCrLf
                sql += " when ISNULL(stb2.q2,0) =4 then '其他' " & vbCrLf
                sql += " else '未選擇' end MID " & vbCrLf
            Case "SD_15_008_R20_2012"
                'sTitle2 = "訓後動向"
                sql += " ,case when ISNULL(stb.q3,0) =1 then '轉換工作' " & vbCrLf
                sql += " when ISNULL(stb.q3,0) =2 then '留任' " & vbCrLf
                sql += " when ISNULL(stb.q3,0) =3 then '其他' " & vbCrLf
                sql += " else '未選擇' end MID " & vbCrLf
            Case "SD_15_008_R21_2012"
                'sTitle2 = "參訓單位類別"
                sql += " ,ko.NAME MID " & vbCrLf
            Case "SD_15_008_R22_2012"
                'sTitle2 = "參加課程職能別"
                sql += " ,kcc.CCName MID " & vbCrLf
            Case "SD_15_008_R23_2012"
                'sTitle2 = "參加課程型態別"'參加課程型態
                sql += " ,case when pp.IsBusiness = 'Y' then '企業包班' " & vbCrLf
                sql += " when pp.PointYN = 'Y' then '學分班' " & vbCrLf
                sql += " when pp.PointYN = 'N'  then '非學分班' " & vbCrLf
                sql += " else '未選擇' end MID " & vbCrLf
            Case "SD_15_008_R24_2012"
                'sTitle2 = "訓練業別"
                sql += " ,ISNULL(ig.PGOVCLASS,ig2.GCODE2+':'+ig2.PCNAME+'-'+ig2.CNAME) MID " & vbCrLf
        End Select
        sql += " ,case when a.Q1=1 then 1 end a11" & vbCrLf
        sql += " ,case when a.Q1=2 then 1 end a12" & vbCrLf
        sql += " ,case when a.Q1=3 then 1 end a13" & vbCrLf
        sql += " ,case when a.Q1=4 then 1 end a14" & vbCrLf
        'sql += " ,case when a.Q1=5 then 1 end a15" & vbCrLf
        sql += " ,case when a.Q2=1 then 1 end a21" & vbCrLf
        sql += " ,case when a.Q2=2 then 1 end a22" & vbCrLf
        sql += " ,case when a.Q2=3 then 1 end a23" & vbCrLf
        sql += " ,case when a.Q2=4 then 1 end a24" & vbCrLf
        sql += " ,case when a.Q2=5 then 1 end a25" & vbCrLf
        sql += " ,case when a.Q3=1 then 1 end a31" & vbCrLf
        sql += " ,case when a.Q3=2 then 1 end a32" & vbCrLf
        sql += " ,case when a.Q3=3 then 1 end a33" & vbCrLf
        sql += " ,case when a.Q3=4 then 1 end a34" & vbCrLf
        'sql += " ,case when a.Q3=5 then 1 end a35" & vbCrLf
        sql += " ,case when a.Q4=1 then 1 end a41" & vbCrLf
        sql += " ,case when a.Q4=2 then 1 end a42" & vbCrLf
        sql += " ,case when a.Q4=3 then 1 end a43" & vbCrLf
        sql += " ,case when a.Q4=4 then 1 end a44" & vbCrLf
        sql += " ,case when a.Q4=5 then 1 end a45" & vbCrLf

        sql += " ,case when a.Q5=1 then 1 end a51" & vbCrLf
        sql += " ,case when a.Q5=2 then 1 end a52" & vbCrLf
        sql += " ,case when a.Q5=3 then 1 end a53" & vbCrLf
        sql += " ,case when a.Q5=4 then 1 end a54" & vbCrLf
        sql += " ,case when a.Q5=5 then 1 end a55" & vbCrLf

        'sql += " ,case when a.Q6=1 then 1 end a61" & vbCrLf
        'sql += " ,case when a.Q6=2 then 1 end a62" & vbCrLf
        'sql += " ,case when a.Q6=3 then 1 end a63" & vbCrLf
        'sql += " ,case when a.Q6=4 then 1 end a64" & vbCrLf
        'sql += " ,case when a.Q6=5 then 1 end a65" & vbCrLf
        sql += " ,case when a.Q6_7=1 then 1 end a6_71" & vbCrLf
        sql += " ,case when a.Q6_7=2 then 1 end a6_72" & vbCrLf
        sql += " ,case when a.Q6_7=3 then 1 end a6_73" & vbCrLf
        sql += " ,case when a.Q6_7=4 then 1 end a6_74" & vbCrLf
        sql += " ,case when a.Q6_7=5 then 1 end a6_75" & vbCrLf

        sql += " ,case when a.Q6_8=1 then 1 end a6_81" & vbCrLf
        sql += " ,case when a.Q6_8=2 then 1 end a6_82" & vbCrLf
        sql += " ,case when a.Q6_8=3 then 1 end a6_83" & vbCrLf
        sql += " ,case when a.Q6_8=4 then 1 end a6_84" & vbCrLf
        sql += " ,case when a.Q6_8=5 then 1 end a6_85" & vbCrLf

        'sql += " ,case when a.Q7=1 then 1 end a71" & vbCrLf
        'sql += " ,case when a.Q7=2 then 1 end a72" & vbCrLf
        'sql += " ,case when a.Q7=3 then 1 end a73" & vbCrLf
        sql += " ,case when a.Q8='1' then 1 end a81" & vbCrLf
        sql += " ,case when a.Q8='2' then 1 end a82" & vbCrLf
        sql += " ,case when a.Q8='3' then 1 end a83" & vbCrLf
        sql += " ,case when a.Q8='4' then 1 end a84" & vbCrLf
        sql += " ,case when a.Q8='5' then 1 end a85" & vbCrLf
        sql += " FROM STUD_QUESTIONFIN A " & vbCrLf
        sql += " JOIN CLASS_STUDENTSOFCLASS cs ON cs.SOCID=A.SOCID " & vbCrLf
        sql += " JOIN STUD_STUDENTINFO ss ON ss.SID=cs.SID " & vbCrLf
        sql += " JOIN STUD_SUBDATA ss2 ON ss2.SID=cs.SID " & vbCrLf
        sql += " JOIN CLASS_CLASSINFO cc ON cc.OCID =cs.OCID" & vbCrLf
        sql += " JOIN PLAN_PLANINFO PP ON PP.PLANID =CC.PLANID  AND PP.COMIDNO= CC.COMIDNO AND PP.SEQNO = CC.SEQNO " & vbCrLf
        sql += " JOIN ID_PLAN ip ON ip.PlanID=cc.PlanID" & vbCrLf
        sql += " JOIN KEY_PLAN kp ON kp.TPlanID =ip.TPlanID" & vbCrLf
        sql += " JOIN ORG_ORGINFO oo ON oo.comidno=pp.comidno" & vbCrLf
        sql += " JOIN VIEW_ZIPNAME iz ON iz.zipCode=ss2.zipCode1 " & vbCrLf
        'fix 無法解析 equal to 作業中 "Chinese_Taiwan_Stroke_CS_AS" 與 "Chinese_Taiwan_Stroke_CI_AS" 之間的定序衝突
        sql += " JOIN V_ORGKIND1 cj1 ON cj1.VALUE COLLATE Chinese_Taiwan_Stroke_CS_AS=oo.ORGKIND2" & vbCrLf

        sql += " LEFT JOIN VIEW_GOVCLASSCAST ig ON ig.GCID=pp.GCID" & vbCrLf
        sql += " LEFT JOIN V_GOVCLASSCAST2 ig2 on ig2.GCID2=pp.GCID2" & vbCrLf
        sql += " LEFT JOIN STUD_TRAINBG stb ON stb.SOCID =cs.SOCID" & vbCrLf

        Select Case filename
            Case "SD_15_008_R25_2012"
                'sTitle2 = "不設定"
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
            Case "SD_15_008_R19_2012"
                'sTitle2 = "參訓動機" '多筆(複選)
                sql += " LEFT JOIN Stud_TrainBGQ2 stb2 ON stb2.SOCID =cs.SOCID" & vbCrLf
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
        End Select
        'Key_OrgType
        sql += " LEFT JOIN KEY_ORGTYPE ko ON ko.OrgTypeID =oo.OrgKind" & vbCrLf
        'Key_ClassCatelog
        sql += " LEFT JOIN KEY_CLASSCATELOG kcc ON kcc.ccid =pp.ClassCate " & vbCrLf
        sql += " LEFT JOIN KEY_TRADE kt1 ON kt1.TRADEID =stb.Q4" & vbCrLf
        sql += " LEFT JOIN KEY_DEGREE kd ON kd.DegreeID =ss.DegreeID" & vbCrLf
        sql += " LEFT JOIN KEY_IDENTITY kd2 ON kd2.IdentityID =cs.MIdentityID" & vbCrLf
        sql += " LEFT JOIN STUD_SUBSIDYCOST sc ON sc.SOCID =cs.SOCID" & vbCrLf
        'sql += " cross join (" & vbCrLf
        'sql += "   select v1.name from V_ORGKIND1 v1 WHERE v1.VALUE='" & SearchPlan & "'" & vbCrLf
        'sql += " ) cj1" & vbCrLf
        sql += " where 1=1" & vbCrLf
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
                sql += " AND oo.OrgKind2 ='" & SearchPlan & "'"
        End Select
        If RID <> "" Then
            sql += " AND cc.RID LIKE '" & RID & "%'" & vbCrLf
        End If
        If OCID <> "" Then
            sql += " AND cc.OCID ='" & OCID & "'" & vbCrLf
        End If
        If TPlanID <> "" Then
            sql += " AND ip.TPlanID='" & TPlanID & "'" & vbCrLf
        End If
        If Years <> "" Then
            sql += " AND ip.Years='" & Years & "'" & vbCrLf
        End If
        If PackageType <> "" Then
            sql += " AND pp.PackageType ='" & PackageType & "'" & vbCrLf
        End If
        'sql += " and cc.RID LIKE 'D2479%'" & vbCrLf
        'sql += " and cc.OCID='74870'" & vbCrLf
        'sql += " AND ip.Years='2015'" & vbCrLf
        sql += " ) G" & vbCrLf
        'sql += " GROUP BY G.PName,G.MID,G.PlanKind,G.SORT1" & vbCrLf
        sql += " GROUP BY G.PName,G.MID,G.PlanKind " & vbCrLf
        'sql += " ORDER BY NLSSORT(G.MID,'NLS_SORT=TCHINESE_STROKE_M')" & vbCrLf
        sql += "ORDER BY G.MID "
        Dim sCmd As New SqlCommand(sql, objconn)

        TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            dt.Load(.ExecuteReader())
        End With

        If dt.Rows.Count > 0 Then
            Dim dr As DataRow = dt.Rows(0)
            PName = Convert.ToString(dr("PName"))
            plankind = Convert.ToString(dr("plankind"))
        End If

        Return dt
    End Function

    '組合 題目與答案 
    Private Sub setItemArray()
        '題目與答案
        '1
        arrItem(0, 0) = "學員目前的近況為何"
        arrItem(0, 1) = "留任原公司"
        arrItem(1, 1) = "轉換至同產業公司"
        arrItem(2, 1) = "轉換至不同產業的公司"
        arrItem(3, 1) = "已離職，待業中"
        '2
        arrItem(4, 0) = "學員於結訓後薪資有提升嗎"
        arrItem(4, 1) = "大幅提升"
        arrItem(5, 1) = "小幅提升"
        arrItem(6, 1) = "沒有變化"
        arrItem(7, 1) = "小幅減少"
        arrItem(8, 1) = "大幅減少"
        '3
        arrItem(9, 0) = "學員職位有變化嗎"
        arrItem(9, 1) = "升遷"
        arrItem(10, 1) = "調職"
        arrItem(11, 1) = "沒有變化"
        arrItem(12, 1) = "降職"
        '4
        arrItem(13, 0) = "學員對目前工作的滿意度是否有變化"
        arrItem(13, 1) = "大幅提升"
        arrItem(14, 1) = "小幅提升"
        arrItem(15, 1) = "沒有變化"
        arrItem(16, 1) = "小幅減少"
        arrItem(17, 1) = "大幅減少"
        '5
        arrItem(18, 0) = "學員目前的工作內容是否與參訓課程內容相關"
        arrItem(18, 1) = "非常相關"
        arrItem(19, 1) = "相關"
        arrItem(20, 1) = "尚可"
        arrItem(21, 1) = "不相關"
        arrItem(22, 1) = "非常不相關"
        '6-1
        arrItem(23, 0) = "學員是否同意參加訓練對目前工作表現有幫助" '學員認為參加訓練對目前工作表現是否有幫助"
        arrItem(23, 1) = "幫助非常大"
        arrItem(24, 1) = "幫助頗多"
        arrItem(25, 1) = "有幫助"
        arrItem(26, 1) = "幫助有限"
        arrItem(27, 1) = "完全沒幫助"

        'arrItem(28, 0) = "學員認為參加訓練對未來工作表現是否有幫助"
        'arrItem(28, 1) = "幫助非常大"
        'arrItem(29, 1) = "幫助頗多"
        'arrItem(30, 1) = "有幫助"
        'arrItem(31, 1) = "幫助有限"
        'arrItem(32, 1) = "完全沒幫助"
        '6-2
        arrItem(28, 0) = "學員是否同意參加訓練對第二專長培育有幫助"
        arrItem(28, 1) = "幫助非常大"
        arrItem(29, 1) = "幫助頗多"
        arrItem(30, 1) = "有幫助"
        arrItem(31, 1) = "幫助有限"
        arrItem(32, 1) = "完全沒幫助"

        'arrItem(38, 0) = "參加本項訓練對學員的幫助是哪方面"
        'arrItem(38, 1) = "對適應工作環境有幫助"
        'arrItem(39, 1) = "對目前工作績效有幫助"
        'arrItem(40, 1) = "對轉換工作跑道有幫助"

        '7
        arrItem(33, 0) = "學員是否有繼續參與進修訓練的意願"
        arrItem(33, 1) = "非常想參與"
        arrItem(34, 1) = "想參與"
        arrItem(35, 1) = "尚無想法"
        arrItem(36, 1) = "不想參與"
        arrItem(37, 1) = "非常不想參與"

        '處理合併表格
        arrSpanRow(0) = "4"
        arrSpanRow(4) = "5"
        arrSpanRow(9) = "4"
        arrSpanRow(13) = "5"
        arrSpanRow(18) = "5"
        arrSpanRow(23) = "5"
        arrSpanRow(28) = "5"
        arrSpanRow(33) = "5"

        'arrSpanRow(38) = "3"
        'arrSpanRow(41) = "5"
    End Sub

    '列印隱鈕
    Protected Sub btnPrt_Click(sender As Object, e As EventArgs) Handles btnPrt.Click
        Call Create11()
        trBtn.Attributes.Remove("style")
        trBtn.Attributes.Add("style", "display:none")

        Select Case filename
            Case "SD_15_008_R25_2012" '不設定
                Call exc_Print("true")
            Case "SD_15_008_R11_2012"
                Call exc_Print("true")
            Case "SD_15_008_R12_2012"
                Call exc_Print("false")
            Case "SD_15_008_R13_2012"
                Call exc_Print("false")
            Case "SD_15_008_R14_2012"
                Call exc_Print("false")
            Case "SD_15_008_R15_2012"
                Call exc_Print("true")
            Case "SD_15_008_R16_2012"
                Call exc_Print("false")
            Case "SD_15_008_R17_2012"
                Call exc_Print("false")
            Case "SD_15_008_R18_2012"
                Call exc_Print("true")
            Case "SD_15_008_R19_2012"
                Call exc_Print("true")
            Case "SD_15_008_R20_2012"
                Call exc_Print("true")
            Case "SD_15_008_R21_2012"
                Call exc_Print("false")
            Case "SD_15_008_R22_2012"
                Call exc_Print("true")
            Case "SD_15_008_R23_2012"
                Call exc_Print("true")
            Case "SD_15_008_R24_2012"
                Call exc_Print("false")
        End Select
    End Sub

    '列印程式
    Private Sub PrintDiv(ByVal dt As DataTable,
                         ByVal selRptfileName As String, ByVal Field1_width As String, ByVal Field2_width As String, ByVal RCount As Integer, ByVal font_size As String, ByVal portrait As String)
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
        If gResDt Is Nothing Then
            Common.MessageBox(Me, "查無資料!!")
            Exit Sub
        End If
        If gResDt.Rows.Count = 0 Then
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

        'tmpDT = dt
        'ColCount = dt.Columns.Count
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

        'Me.ViewState("xHtml") = ""
        If dt.Rows.Count > 0 Then

            Dim iAColSpan As Integer = 0
            iAColSpan = 2 + gResDt.Rows.Count + 1

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

                Select Case selRptfileName
                    Case "SD_15_008_R25_2012" '不設定
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
                For j As Integer = 0 To intPageRecord
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
                nl.InnerHtml = "<p style='line-height:2px;margin:0cm;margin-bottom:0.0001pt;mso-pagination:widow-orphan;'><br clear=all style='mso-special-character:line-break;page-break-before:always'>"
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
                    tmpDT.Columns.Add(New DataColumn(Convert.ToString(dt.Rows(j)("NAME"))))
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

            For j As Integer = 0 To dt.Rows.Count - 1
                iNo = j + 2
                tmpDR(iNo) = dt.Rows(j)(iA)
                intTotal += dt.Rows(j)(iA)
            Next
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
        TIMS.Utl_RespWriteEnd(Me, objconn, "") ' Response.End()
    End Sub

#Region "NO USE"
    'Private Function db_R12(ByVal TPlanID As String, ByVal Years As String, ByVal OCID As String, ByVal RID As String, ByVal SearchPlan As String, ByVal PackageType As String) As DataTable
    '    Dim tmpDT As DataTable
    '    Dim resDt As DataTable = Nothing
    '    Dim sql As String = ""

    '    sql = ""
    '    sql = "select convert(varchar, CONVERT(numeric,  '" + Years + "') -1911 ) + '年度　訓後動態調查統計表 年齡' AS PName, k.Name, "
    '    sql += "( Select a.PlanName +dbo.NVL(case when a.TPlanID='28' then (select '('+name+')' from v_orgkind1 WHERE VALUE='" & SearchPlan & "') end,'') from key_plan a where a.TPlanID='" & TPlanID & "') as plankind, "
    '    sql += "dbo.NVL( q.a11,0) as a11,"
    '    sql += "dbo.NVL( q.a12,0) as a12,"
    '    sql += "dbo.NVL( q.a13,0) as a13,"
    '    sql += "dbo.NVL( q.a14,0) as a14,"
    '    sql += "dbo.NVL( q.a21,0) as a21,"
    '    sql += "dbo.NVL( q.a22,0) as a22,"
    '    sql += "dbo.NVL( q.a23,0) as a23,"
    '    sql += "dbo.NVL( q.a24,0) as a24,"
    '    sql += "dbo.NVL( q.a25,0) as a25,"
    '    sql += "dbo.NVL( q.a31,0) as a31,"
    '    sql += "dbo.NVL( q.a32,0) as a32,"
    '    sql += "dbo.NVL( q.a33,0) as a33,"
    '    sql += "dbo.NVL( q.a34,0) as a34,"
    '    sql += "dbo.NVL( q.a41,0) as a41,"
    '    sql += "dbo.NVL( q.a42,0) as a42,"
    '    sql += "dbo.NVL( q.a43,0) as a43,"
    '    sql += "dbo.NVL( q.a44,0) as a44,"
    '    sql += "dbo.NVL( q.a45,0) as a45,"
    '    sql += "dbo.NVL( q.a51,0) as a51,"
    '    sql += "dbo.NVL( q.a52,0) as a52,"
    '    sql += "dbo.NVL( q.a53,0) as a53,"
    '    sql += "dbo.NVL( q.a54,0) as a54,"
    '    sql += "dbo.NVL( q.a55,0) as a55,"
    '    sql += "dbo.NVL( q.a61,0) as a61,"
    '    sql += "dbo.NVL( q.a62,0) as a62,"
    '    sql += "dbo.NVL( q.a63,0) as a63,"
    '    sql += "dbo.NVL( q.a64,0) as a64,"
    '    sql += "dbo.NVL( q.a65,0) as a65,"
    '    sql += "dbo.NVL( q.a6_71,0) as a6_71,"
    '    sql += "dbo.NVL( q.a6_72,0) as a6_72,"
    '    sql += "dbo.NVL( q.a6_73,0) as a6_73,"
    '    sql += "dbo.NVL( q.a6_74,0) as a6_74,"
    '    sql += "dbo.NVL( q.a6_75,0) as a6_75,"
    '    sql += "dbo.NVL( q.a6_81,0) as a6_81,"
    '    sql += "dbo.NVL( q.a6_82,0) as a6_82,"
    '    sql += "dbo.NVL( q.a6_83,0) as a6_83,"
    '    sql += "dbo.NVL( q.a6_84,0) as a6_84,"
    '    sql += "dbo.NVL( q.a6_85,0) as a6_85,"
    '    sql += "dbo.NVL( q.a71,0) as a71,"
    '    sql += "dbo.NVL( q.a72,0) as a72,"
    '    sql += "dbo.NVL( q.a73,0) as a73,"
    '    sql += "dbo.NVL( q.a81,0) as a81,"
    '    sql += "dbo.NVL( q.a82,0) as a82,"
    '    sql += "dbo.NVL( q.a83,0) as a83,"
    '    sql += "dbo.NVL( q.a84,0) as a84,"
    '    sql += "dbo.NVL( q.a85,0) as a85 "

    '    sql += "from ( "
    '    sql += "select '01' as mid, '15以下' as Name  union "
    '    sql += "select '02' as mid, '15~19' as Name  union "
    '    sql += "select '03' as mid, '20~24' as Name  union "
    '    sql += "select '04' as mid, '25~29' as Name  union "
    '    sql += "select '05' as mid, '30~34' as Name  union "
    '    sql += "select '06' as mid, '35~39' as Name  union "
    '    sql += "select '07' as mid, '40~44' as Name  union "
    '    sql += "select '08' as mid, '45~49' as Name  union "
    '    sql += "select '09' as mid, '50~54' as Name  union "
    '    sql += "select '10' as mid, '55~59' as Name  union "
    '    sql += "select '11' as mid, '60以上' as Name  "
    '    sql += ") k "
    '    sql += "left outer join ( SELECT c.Mid,  "
    '    sql += "sum(case a.Q1 when 1 then 1 else 0 end) as  a11,"
    '    sql += "sum(case a.Q1 when 2 then 1 else 0 end) as  a12,"
    '    sql += "sum(case a.Q1 when 3 then 1 else 0 end) as  a13,"
    '    sql += "sum(case a.Q1 when 4 then 1 else 0 end) as  a14,"
    '    sql += "sum(case a.Q2 when 1 then 1 else 0 end) as  a21,"
    '    sql += "sum(case a.Q2 when 2 then 1 else 0 end) as  a22,"
    '    sql += "sum(case a.Q2 when 3 then 1 else 0 end) as  a23,"
    '    sql += "sum(case a.Q2 when 4 then 1 else 0 end) as  a24,"
    '    sql += "sum(case a.Q2 when 5 then 1 else 0 end) as  a25,"
    '    sql += "sum(case a.Q3 when 1 then 1 else 0 end) as  a31,"
    '    sql += "sum(case a.Q3 when 2 then 1 else 0 end) as  a32,"
    '    sql += "sum(case a.Q3 when 3 then 1 else 0 end) as  a33,"
    '    sql += "sum(case a.Q3 when 4 then 1 else 0 end) as  a34,"
    '    sql += "sum(case a.Q4 when 1 then 1 else 0 end) as  a41,"
    '    sql += "sum(case a.Q4 when 2 then 1 else 0 end) as  a42,"
    '    sql += "sum(case a.Q4 when 3 then 1 else 0 end) as  a43,"
    '    sql += "sum(case a.Q4 when 4 then 1 else 0 end) as  a44,"
    '    sql += "sum(case a.Q4 when 5 then 1 else 0 end) as  a45,"
    '    sql += "sum(case a.Q5 when 1 then 1 else 0 end) as  a51,"
    '    sql += "sum(case a.Q5 when 2 then 1 else 0 end) as  a52,"
    '    sql += "sum(case a.Q5 when 3 then 1 else 0 end) as  a53,"
    '    sql += "sum(case a.Q5 when 4 then 1 else 0 end) as  a54,"
    '    sql += "sum(case a.Q5 when 5 then 1 else 0 end) as  a55,"
    '    sql += "sum(case a.Q6 when 1 then 1 else 0 end) as  a61,"
    '    sql += "sum(case a.Q6 when 2 then 1 else 0 end) as  a62,"
    '    sql += "sum(case a.Q6 when 3 then 1 else 0 end) as  a63,"
    '    sql += "sum(case a.Q6 when 4 then 1 else 0 end) as  a64,"
    '    sql += "sum(case a.Q6 when 5 then 1 else 0 end) as  a65,"
    '    sql += "sum(case a.Q6_7 when 1 then 1 else 0 end) as a6_71,"
    '    sql += "sum(case a.Q6_7 when 2 then 1 else 0 end) as a6_72,"
    '    sql += "sum(case a.Q6_7 when 3 then 1 else 0 end) as a6_73,"
    '    sql += "sum(case a.Q6_7 when 4 then 1 else 0 end) as a6_74,"
    '    sql += "sum(case a.Q6_7 when 5 then 1 else 0 end) as a6_75,"
    '    sql += "sum(case a.Q6_8 when 1 then 1 else 0 end) as a6_81,"
    '    sql += "sum(case a.Q6_8 when 2 then 1 else 0 end) as a6_82,"
    '    sql += "sum(case a.Q6_8 when 3 then 1 else 0 end) as a6_83,"
    '    sql += "sum(case a.Q6_8 when 4 then 1 else 0 end) as a6_84,"
    '    sql += "sum(case a.Q6_8 when 5 then 1 else 0 end) as a6_85,  "
    '    sql += "sum(case a.Q7 when 1 then 1 else 0 end) as  a71,"
    '    sql += "sum(case a.Q7 when 2 then 1 else 0 end) as  a72,"
    '    sql += "sum(case a.Q7 when 3 then 1 else 0 end) as  a73,"
    '    sql += "sum(case a.Q8 when '1' then 1 else 0 end) as a81,"
    '    sql += "sum(case a.Q8 when '2' then 1 else 0 end) as a82,"
    '    sql += "sum(case a.Q8 when '3' then 1 else 0 end) as a83,"
    '    sql += "sum(case a.Q8 when '4' then 1 else 0 end) as a84,"
    '    sql += "sum(case a.Q8 when '5' then 1 else 0 end) as a85 "
    '    sql += "From Stud_QuestionFin a left outer join Class_StudentsOfClass b on a.SOCID=b.socid "
    '    sql += "left outer join (SELECT SID,case "
    '    sql += "when 15>DATEDIFF(YEAR, Birthday, getdate() ) then '01' "
    '    sql += "when DATEDIFF(YEAR, Birthday, getdate() )>=15 and  20>DATEDIFF(YEAR, Birthday, getdate() ) then '02' "
    '    sql += "when DATEDIFF(YEAR, Birthday, getdate() )>=20 and  25>DATEDIFF(YEAR, Birthday, getdate() ) then '03' "
    '    sql += "when DATEDIFF(YEAR, Birthday, getdate() )>=25 and  30>DATEDIFF(YEAR, Birthday, getdate() ) then '04' "
    '    sql += "when DATEDIFF(YEAR, Birthday, getdate() )>=30 and  35>DATEDIFF(YEAR, Birthday, getdate() ) then '05' "
    '    sql += "when DATEDIFF(YEAR, Birthday, getdate() )>=35 and  40>DATEDIFF(YEAR, Birthday, getdate() ) then '06' "
    '    sql += "when DATEDIFF(YEAR, Birthday, getdate() )>=40 and  45>DATEDIFF(YEAR, Birthday, getdate() ) then '07' "
    '    sql += "when DATEDIFF(YEAR, Birthday, getdate() )>=45 and  50>DATEDIFF(YEAR, Birthday, getdate() ) then '08' "
    '    sql += "when DATEDIFF(YEAR, Birthday, getdate() )>=50 and  55>DATEDIFF(YEAR, Birthday, getdate() ) then '09' "
    '    sql += "when DATEDIFF(YEAR, Birthday, getdate() )>=55 and  60>DATEDIFF(YEAR, Birthday, getdate() ) then '10' "
    '    sql += "else '11' end AS mid FROM Stud_StudentInfo ) c on b.SID=c.sid "
    '    sql += "INNER JOIN (select rid, PlanID, Years, OCID, SEQNO  from Class_ClassInfo where 1=1 "
    '    If RID <> "" Then
    '        sql += "AND RID LIKE '" + RID + "%'"
    '    End If
    '    If OCID <> "" Then
    '        sql += "AND OCID='" + OCID + "'"
    '    End If
    '    sql += ") cc ON b.OCID=cc.OCID JOIN (select PlanID, Years from ID_Plan where TPlanID='" + TPlanID + "' "
    '    If Years <> "" Then
    '        sql += "AND Years='" + Years + "'"
    '    End If
    '    sql += ") ip ON cc.PlanID = ip.PlanID AND cc.Years = dbo.SUBSTR(ip.Years,3,2) "
    '    sql += "JOIN Plan_Planinfo pp ON pp.Planid =cc.Planid  AND pp.RID = cc.RID AND pp.SEQNO = cc.SEQNO "
    '    If PackageType <> "" Then
    '        sql += "AND pp.PackageType ='" + PackageType + "'"
    '    End If
    '    sql += "JOIN (select * from org_orginfo  where 1=1 "
    '    If SearchPlan <> "" Then
    '        sql += "and OrgKind2 ='" + SearchPlan + "' "
    '    End If
    '    sql += ") oo ON pp.COMIDNO =oo.COMIDNO group by c.Mid) q on q.mid=k.mid "
    '    sql += "order by nlssort(NAME,'NLS_SORT=TCHINESE_STROKE_M') "

    '    resDt = DbAccess.GetDataTable(sql, objconn)
    '    'sourceDT = resDt
    '    'If sourceDT.Rows.Count > 0 Then
    '    '    PName = sourceDT.Rows(0)("PName")
    '    '    plankind = sourceDT.Rows(0)("plankind")
    '    'End If

    '    setItemArray() '題目內容

    '    tmpDT = RotationDT(resDt)
    '    Return tmpDT

    'End Function

    'Private Function db_R13(ByVal TPlanID As String, ByVal Years As String, ByVal OCID As String, ByVal RID As String, ByVal SearchPlan As String, ByVal PackageType As String) As DataTable
    '    Dim tmpDT As DataTable
    '    Dim resDt As DataTable = Nothing
    '    Dim sql As String = ""

    '    sql = ""
    '    sql = "select convert(varchar, CONVERT(numeric,  '" + Years + "') -1911 ) + '年度　訓後動態調查統計表 教育程度' AS PName, k.Name, "
    '    sql += "( Select a.PlanName +dbo.NVL(case when a.TPlanID='28' then (select '('+name+')' from v_orgkind1 WHERE VALUE='" & SearchPlan & "') end,'') from key_plan a where a.TPlanID='" & TPlanID & "') as plankind, "
    '    sql += "dbo.NVL( q.a11,0) as a11,"
    '    sql += "dbo.NVL( q.a12,0) as a12,"
    '    sql += "dbo.NVL( q.a13,0) as a13,"
    '    sql += "dbo.NVL( q.a14,0) as a14,"
    '    sql += "dbo.NVL( q.a21,0) as a21,"
    '    sql += "dbo.NVL( q.a22,0) as a22,"
    '    sql += "dbo.NVL( q.a23,0) as a23,"
    '    sql += "dbo.NVL( q.a24,0) as a24,"
    '    sql += "dbo.NVL( q.a25,0) as a25,"
    '    sql += "dbo.NVL( q.a31,0) as a31,"
    '    sql += "dbo.NVL( q.a32,0) as a32,"
    '    sql += "dbo.NVL( q.a33,0) as a33,"
    '    sql += "dbo.NVL( q.a34,0) as a34,"
    '    sql += "dbo.NVL( q.a41,0) as a41,"
    '    sql += "dbo.NVL( q.a42,0) as a42,"
    '    sql += "dbo.NVL( q.a43,0) as a43,"
    '    sql += "dbo.NVL( q.a44,0) as a44,"
    '    sql += "dbo.NVL( q.a45,0) as a45,"
    '    sql += "dbo.NVL( q.a51,0) as a51,"
    '    sql += "dbo.NVL( q.a52,0) as a52,"
    '    sql += "dbo.NVL( q.a53,0) as a53,"
    '    sql += "dbo.NVL( q.a54,0) as a54,"
    '    sql += "dbo.NVL( q.a55,0) as a55,"
    '    sql += "dbo.NVL( q.a61,0) as a61,"
    '    sql += "dbo.NVL( q.a62,0) as a62,"
    '    sql += "dbo.NVL( q.a63,0) as a63,"
    '    sql += "dbo.NVL( q.a64,0) as a64,"
    '    sql += "dbo.NVL( q.a65,0) as a65,"
    '    sql += "dbo.NVL( q.a6_71,0) as a6_71,"
    '    sql += "dbo.NVL( q.a6_72,0) as a6_72,"
    '    sql += "dbo.NVL( q.a6_73,0) as a6_73,"
    '    sql += "dbo.NVL( q.a6_74,0) as a6_74,"
    '    sql += "dbo.NVL( q.a6_75,0) as a6_75,"
    '    sql += "dbo.NVL( q.a6_81,0) as a6_81,"
    '    sql += "dbo.NVL( q.a6_82,0) as a6_82,"
    '    sql += "dbo.NVL( q.a6_83,0) as a6_83,"
    '    sql += "dbo.NVL( q.a6_84,0) as a6_84,"
    '    sql += "dbo.NVL( q.a6_85,0) as a6_85,"
    '    sql += "dbo.NVL( q.a71,0) as a71,"
    '    sql += "dbo.NVL( q.a72,0) as a72,"
    '    sql += "dbo.NVL( q.a73,0) as a73,"
    '    sql += "dbo.NVL( q.a81,0) as a81,"
    '    sql += "dbo.NVL( q.a82,0) as a82,"
    '    sql += "dbo.NVL( q.a83,0) as a83,"
    '    sql += "dbo.NVL( q.a84,0) as a84,"
    '    sql += "dbo.NVL( q.a85,0) as a85 "

    '    sql += "from (Select DegreeID as mid, Name from Key_Degree where DegreeID<>'00') k "
    '    sql += "left outer join ( SELECT c.Mid,  "
    '    sql += "sum(case a.Q1 when 1 then 1 else 0 end) as  a11,"
    '    sql += "sum(case a.Q1 when 2 then 1 else 0 end) as  a12,"
    '    sql += "sum(case a.Q1 when 3 then 1 else 0 end) as  a13,"
    '    sql += "sum(case a.Q1 when 4 then 1 else 0 end) as  a14,"
    '    sql += "sum(case a.Q2 when 1 then 1 else 0 end) as  a21,"
    '    sql += "sum(case a.Q2 when 2 then 1 else 0 end) as  a22,"
    '    sql += "sum(case a.Q2 when 3 then 1 else 0 end) as  a23,"
    '    sql += "sum(case a.Q2 when 4 then 1 else 0 end) as  a24,"
    '    sql += "sum(case a.Q2 when 5 then 1 else 0 end) as  a25,"
    '    sql += "sum(case a.Q3 when 1 then 1 else 0 end) as  a31,"
    '    sql += "sum(case a.Q3 when 2 then 1 else 0 end) as  a32,"
    '    sql += "sum(case a.Q3 when 3 then 1 else 0 end) as  a33,"
    '    sql += "sum(case a.Q3 when 4 then 1 else 0 end) as  a34,"
    '    sql += "sum(case a.Q4 when 1 then 1 else 0 end) as  a41,"
    '    sql += "sum(case a.Q4 when 2 then 1 else 0 end) as  a42,"
    '    sql += "sum(case a.Q4 when 3 then 1 else 0 end) as  a43,"
    '    sql += "sum(case a.Q4 when 4 then 1 else 0 end) as  a44,"
    '    sql += "sum(case a.Q4 when 5 then 1 else 0 end) as  a45,"
    '    sql += "sum(case a.Q5 when 1 then 1 else 0 end) as  a51,"
    '    sql += "sum(case a.Q5 when 2 then 1 else 0 end) as  a52,"
    '    sql += "sum(case a.Q5 when 3 then 1 else 0 end) as  a53,"
    '    sql += "sum(case a.Q5 when 4 then 1 else 0 end) as  a54,"
    '    sql += "sum(case a.Q5 when 5 then 1 else 0 end) as  a55,"
    '    sql += "sum(case a.Q6 when 1 then 1 else 0 end) as  a61,"
    '    sql += "sum(case a.Q6 when 2 then 1 else 0 end) as  a62,"
    '    sql += "sum(case a.Q6 when 3 then 1 else 0 end) as  a63,"
    '    sql += "sum(case a.Q6 when 4 then 1 else 0 end) as  a64,"
    '    sql += "sum(case a.Q6 when 5 then 1 else 0 end) as  a65,"
    '    sql += "sum(case a.Q6_7 when 1 then 1 else 0 end) as a6_71,"
    '    sql += "sum(case a.Q6_7 when 2 then 1 else 0 end) as a6_72,"
    '    sql += "sum(case a.Q6_7 when 3 then 1 else 0 end) as a6_73,"
    '    sql += "sum(case a.Q6_7 when 4 then 1 else 0 end) as a6_74,"
    '    sql += "sum(case a.Q6_7 when 5 then 1 else 0 end) as a6_75,"
    '    sql += "sum(case a.Q6_8 when 1 then 1 else 0 end) as a6_81,"
    '    sql += "sum(case a.Q6_8 when 2 then 1 else 0 end) as a6_82,"
    '    sql += "sum(case a.Q6_8 when 3 then 1 else 0 end) as a6_83,"
    '    sql += "sum(case a.Q6_8 when 4 then 1 else 0 end) as a6_84,"
    '    sql += "sum(case a.Q6_8 when 5 then 1 else 0 end) as a6_85,  "
    '    sql += "sum(case a.Q7 when 1 then 1 else 0 end) as  a71,"
    '    sql += "sum(case a.Q7 when 2 then 1 else 0 end) as  a72,"
    '    sql += "sum(case a.Q7 when 3 then 1 else 0 end) as  a73,"
    '    sql += "sum(case a.Q8 when '1' then 1 else 0 end) as a81,"
    '    sql += "sum(case a.Q8 when '2' then 1 else 0 end) as a82,"
    '    sql += "sum(case a.Q8 when '3' then 1 else 0 end) as a83,"
    '    sql += "sum(case a.Q8 when '4' then 1 else 0 end) as a84,"
    '    sql += "sum(case a.Q8 when '5' then 1 else 0 end) as a85 "
    '    sql += "From Stud_QuestionFin a left outer join Class_StudentsOfClass b on a.SOCID=b.socid "
    '    sql += "left outer join (SELECT SID, DegreeID AS mid  FROM Stud_StudentInfo ) c on b.SID=c.sid "
    '    sql += "INNER JOIN (select rid, PlanID, Years, OCID, SEQNO  from Class_ClassInfo where 1=1 "
    '    If RID <> "" Then
    '        sql += "AND RID LIKE '" + RID + "%'"
    '    End If
    '    If OCID <> "" Then
    '        sql += "AND OCID='" + OCID + "'"
    '    End If
    '    sql += ") cc ON b.OCID=cc.OCID JOIN (select PlanID, Years from ID_Plan where TPlanID='" + TPlanID + "' "
    '    If Years <> "" Then
    '        sql += "AND Years='" + Years + "'"
    '    End If
    '    sql += ") ip ON cc.PlanID = ip.PlanID AND cc.Years = dbo.SUBSTR(ip.Years,3,2) "
    '    sql += "JOIN Plan_Planinfo pp ON pp.Planid =cc.Planid  AND pp.RID = cc.RID AND pp.SEQNO = cc.SEQNO "
    '    If PackageType <> "" Then
    '        sql += "AND pp.PackageType ='" + PackageType + "'"
    '    End If
    '    sql += "JOIN (select * from org_orginfo  where 1=1 "
    '    If SearchPlan <> "" Then
    '        sql += "and OrgKind2 ='" + SearchPlan + "' "
    '    End If
    '    sql += ") oo ON pp.COMIDNO =oo.COMIDNO group by c.Mid) q on q.mid=k.mid "
    '    sql += "order by nlssort(NAME,'NLS_SORT=TCHINESE_STROKE_M') "

    '    resDt = DbAccess.GetDataTable(sql, objconn)
    '    'sourceDT = resDt
    '    'If sourceDT.Rows.Count > 0 Then
    '    '    PName = sourceDT.Rows(0)("PName")
    '    '    plankind = sourceDT.Rows(0)("plankind")
    '    'End If

    '    setItemArray() '題目內容

    '    tmpDT = RotationDT(resDt)
    '    Return tmpDT

    'End Function

    'Private Function db_R14(ByVal TPlanID As String, ByVal Years As String, ByVal OCID As String, ByVal RID As String, ByVal SearchPlan As String, ByVal PackageType As String) As DataTable
    '    Dim tmpDT As DataTable
    '    Dim resDt As DataTable = Nothing
    '    Dim sql As String = ""

    '    sql = ""
    '    sql = "select convert(varchar, CONVERT(numeric,  '" + Years + "') -1911 ) + '年度　訓後動態調查統計表 身分別' AS PName, k.Name, "
    '    sql += "( Select a.PlanName +dbo.NVL(case when a.TPlanID='28' then (select '('+name+')' from v_orgkind1 WHERE VALUE='" & SearchPlan & "') end,'') from key_plan a where a.TPlanID='" & TPlanID & "') as plankind, "
    '    sql += "dbo.NVL( q.a11,0) as a11,"
    '    sql += "dbo.NVL( q.a12,0) as a12,"
    '    sql += "dbo.NVL( q.a13,0) as a13,"
    '    sql += "dbo.NVL( q.a14,0) as a14,"
    '    sql += "dbo.NVL( q.a21,0) as a21,"
    '    sql += "dbo.NVL( q.a22,0) as a22,"
    '    sql += "dbo.NVL( q.a23,0) as a23,"
    '    sql += "dbo.NVL( q.a24,0) as a24,"
    '    sql += "dbo.NVL( q.a25,0) as a25,"
    '    sql += "dbo.NVL( q.a31,0) as a31,"
    '    sql += "dbo.NVL( q.a32,0) as a32,"
    '    sql += "dbo.NVL( q.a33,0) as a33,"
    '    sql += "dbo.NVL( q.a34,0) as a34,"
    '    sql += "dbo.NVL( q.a41,0) as a41,"
    '    sql += "dbo.NVL( q.a42,0) as a42,"
    '    sql += "dbo.NVL( q.a43,0) as a43,"
    '    sql += "dbo.NVL( q.a44,0) as a44,"
    '    sql += "dbo.NVL( q.a45,0) as a45,"
    '    sql += "dbo.NVL( q.a51,0) as a51,"
    '    sql += "dbo.NVL( q.a52,0) as a52,"
    '    sql += "dbo.NVL( q.a53,0) as a53,"
    '    sql += "dbo.NVL( q.a54,0) as a54,"
    '    sql += "dbo.NVL( q.a55,0) as a55,"
    '    sql += "dbo.NVL( q.a61,0) as a61,"
    '    sql += "dbo.NVL( q.a62,0) as a62,"
    '    sql += "dbo.NVL( q.a63,0) as a63,"
    '    sql += "dbo.NVL( q.a64,0) as a64,"
    '    sql += "dbo.NVL( q.a65,0) as a65,"
    '    sql += "dbo.NVL( q.a6_71,0) as a6_71,"
    '    sql += "dbo.NVL( q.a6_72,0) as a6_72,"
    '    sql += "dbo.NVL( q.a6_73,0) as a6_73,"
    '    sql += "dbo.NVL( q.a6_74,0) as a6_74,"
    '    sql += "dbo.NVL( q.a6_75,0) as a6_75,"
    '    sql += "dbo.NVL( q.a6_81,0) as a6_81,"
    '    sql += "dbo.NVL( q.a6_82,0) as a6_82,"
    '    sql += "dbo.NVL( q.a6_83,0) as a6_83,"
    '    sql += "dbo.NVL( q.a6_84,0) as a6_84,"
    '    sql += "dbo.NVL( q.a6_85,0) as a6_85,"
    '    sql += "dbo.NVL( q.a71,0) as a71,"
    '    sql += "dbo.NVL( q.a72,0) as a72,"
    '    sql += "dbo.NVL( q.a73,0) as a73,"
    '    sql += "dbo.NVL( q.a81,0) as a81,"
    '    sql += "dbo.NVL( q.a82,0) as a82,"
    '    sql += "dbo.NVL( q.a83,0) as a83,"
    '    sql += "dbo.NVL( q.a84,0) as a84,"
    '    sql += "dbo.NVL( q.a85,0) as a85 "

    '    sql += "from (select IdentityID mid, Name from key_Identity)  k "
    '    sql += "left outer join (SELECT b.Midentityid mid, "
    '    sql += "sum(case a.Q1 when 1 then 1 else 0 end) as  a11,"
    '    sql += "sum(case a.Q1 when 2 then 1 else 0 end) as  a12,"
    '    sql += "sum(case a.Q1 when 3 then 1 else 0 end) as  a13,"
    '    sql += "sum(case a.Q1 when 4 then 1 else 0 end) as  a14,"
    '    sql += "sum(case a.Q2 when 1 then 1 else 0 end) as  a21,"
    '    sql += "sum(case a.Q2 when 2 then 1 else 0 end) as  a22,"
    '    sql += "sum(case a.Q2 when 3 then 1 else 0 end) as  a23,"
    '    sql += "sum(case a.Q2 when 4 then 1 else 0 end) as  a24,"
    '    sql += "sum(case a.Q2 when 5 then 1 else 0 end) as  a25,"
    '    sql += "sum(case a.Q3 when 1 then 1 else 0 end) as  a31,"
    '    sql += "sum(case a.Q3 when 2 then 1 else 0 end) as  a32,"
    '    sql += "sum(case a.Q3 when 3 then 1 else 0 end) as  a33,"
    '    sql += "sum(case a.Q3 when 4 then 1 else 0 end) as  a34,"
    '    sql += "sum(case a.Q4 when 1 then 1 else 0 end) as  a41,"
    '    sql += "sum(case a.Q4 when 2 then 1 else 0 end) as  a42,"
    '    sql += "sum(case a.Q4 when 3 then 1 else 0 end) as  a43,"
    '    sql += "sum(case a.Q4 when 4 then 1 else 0 end) as  a44,"
    '    sql += "sum(case a.Q4 when 5 then 1 else 0 end) as  a45,"
    '    sql += "sum(case a.Q5 when 1 then 1 else 0 end) as  a51,"
    '    sql += "sum(case a.Q5 when 2 then 1 else 0 end) as  a52,"
    '    sql += "sum(case a.Q5 when 3 then 1 else 0 end) as  a53,"
    '    sql += "sum(case a.Q5 when 4 then 1 else 0 end) as  a54,"
    '    sql += "sum(case a.Q5 when 5 then 1 else 0 end) as  a55,"
    '    sql += "sum(case a.Q6 when 1 then 1 else 0 end) as  a61,"
    '    sql += "sum(case a.Q6 when 2 then 1 else 0 end) as  a62,"
    '    sql += "sum(case a.Q6 when 3 then 1 else 0 end) as  a63,"
    '    sql += "sum(case a.Q6 when 4 then 1 else 0 end) as  a64,"
    '    sql += "sum(case a.Q6 when 5 then 1 else 0 end) as  a65,"
    '    sql += "sum(case a.Q6_7 when 1 then 1 else 0 end) as a6_71,"
    '    sql += "sum(case a.Q6_7 when 2 then 1 else 0 end) as a6_72,"
    '    sql += "sum(case a.Q6_7 when 3 then 1 else 0 end) as a6_73,"
    '    sql += "sum(case a.Q6_7 when 4 then 1 else 0 end) as a6_74,"
    '    sql += "sum(case a.Q6_7 when 5 then 1 else 0 end) as a6_75,"
    '    sql += "sum(case a.Q6_8 when 1 then 1 else 0 end) as a6_81,"
    '    sql += "sum(case a.Q6_8 when 2 then 1 else 0 end) as a6_82,"
    '    sql += "sum(case a.Q6_8 when 3 then 1 else 0 end) as a6_83,"
    '    sql += "sum(case a.Q6_8 when 4 then 1 else 0 end) as a6_84,"
    '    sql += "sum(case a.Q6_8 when 5 then 1 else 0 end) as a6_85,  "
    '    sql += "sum(case a.Q7 when 1 then 1 else 0 end) as  a71,"
    '    sql += "sum(case a.Q7 when 2 then 1 else 0 end) as  a72,"
    '    sql += "sum(case a.Q7 when 3 then 1 else 0 end) as  a73,"
    '    sql += "sum(case a.Q8 when '1' then 1 else 0 end) as a81,"
    '    sql += "sum(case a.Q8 when '2' then 1 else 0 end) as a82,"
    '    sql += "sum(case a.Q8 when '3' then 1 else 0 end) as a83,"
    '    sql += "sum(case a.Q8 when '4' then 1 else 0 end) as a84,"
    '    sql += "sum(case a.Q8 when '5' then 1 else 0 end) as a85 "
    '    sql += "From Stud_QuestionFin a left outer join Class_StudentsOfClass b on a.SOCID=b.socid "
    '    sql += "INNER JOIN (select rid, PlanID, Years, OCID, SEQNO  from Class_ClassInfo where 1=1 "
    '    If RID <> "" Then
    '        sql += "AND RID LIKE '" + RID + "%'"
    '    End If
    '    If OCID <> "" Then
    '        sql += "AND OCID='" + OCID + "'"
    '    End If
    '    sql += ") cc ON b.OCID=cc.OCID JOIN (select PlanID, Years from ID_Plan where TPlanID='" + TPlanID + "' "
    '    If Years <> "" Then
    '        sql += "AND Years='" + Years + "'"
    '    End If
    '    sql += ") ip ON cc.PlanID = ip.PlanID AND cc.Years = dbo.SUBSTR(ip.Years,3,2) "
    '    sql += "JOIN Plan_Planinfo pp ON pp.Planid =cc.Planid  AND pp.RID = cc.RID AND pp.SEQNO = cc.SEQNO "
    '    If PackageType <> "" Then
    '        sql += "AND pp.PackageType ='" + PackageType + "'"
    '    End If
    '    sql += "JOIN (select * from org_orginfo  where 1=1 "
    '    If SearchPlan <> "" Then
    '        sql += "and OrgKind2 ='" + SearchPlan + "' "
    '    End If
    '    sql += ") oo ON pp.COMIDNO =oo.COMIDNO group by b.Midentityid ) q on q.mid=k.mid "
    '    sql += "where 1=1 order by k.mid "

    '    resDt = DbAccess.GetDataTable(sql, objconn)
    '    'sourceDT = resDt
    '    'If sourceDT.Rows.Count > 0 Then
    '    '    PName = sourceDT.Rows(0)("PName")
    '    '    plankind = sourceDT.Rows(0)("plankind")
    '    'End If

    '    setItemArray() '題目內容

    '    tmpDT = RotationDT(resDt)
    '    Return tmpDT

    'End Function

    'Private Function db_R15(ByVal TPlanID As String, ByVal Years As String, ByVal OCID As String, ByVal RID As String, ByVal SearchPlan As String, ByVal PackageType As String) As DataTable
    '    Dim tmpDT As DataTable
    '    Dim resDt As DataTable = Nothing
    '    Dim sql As String = ""

    '    sql = ""
    '    sql = "select convert(varchar, CONVERT(numeric,  '" + Years + "') -1911 ) + '年度　訓後動態調查統計表 工作年資' AS PName, k.Name, "
    '    sql += "( Select a.PlanName +dbo.NVL(case when a.TPlanID='28' then (select '('+name+')' from v_orgkind1 WHERE VALUE='" & SearchPlan & "') end,'') from key_plan a where a.TPlanID='" & TPlanID & "') as plankind, "
    '    sql += "dbo.NVL( q.a11,0) as a11,"
    '    sql += "dbo.NVL( q.a12,0) as a12,"
    '    sql += "dbo.NVL( q.a13,0) as a13,"
    '    sql += "dbo.NVL( q.a14,0) as a14,"
    '    sql += "dbo.NVL( q.a21,0) as a21,"
    '    sql += "dbo.NVL( q.a22,0) as a22,"
    '    sql += "dbo.NVL( q.a23,0) as a23,"
    '    sql += "dbo.NVL( q.a24,0) as a24,"
    '    sql += "dbo.NVL( q.a25,0) as a25,"
    '    sql += "dbo.NVL( q.a31,0) as a31,"
    '    sql += "dbo.NVL( q.a32,0) as a32,"
    '    sql += "dbo.NVL( q.a33,0) as a33,"
    '    sql += "dbo.NVL( q.a34,0) as a34,"
    '    sql += "dbo.NVL( q.a41,0) as a41,"
    '    sql += "dbo.NVL( q.a42,0) as a42,"
    '    sql += "dbo.NVL( q.a43,0) as a43,"
    '    sql += "dbo.NVL( q.a44,0) as a44,"
    '    sql += "dbo.NVL( q.a45,0) as a45,"
    '    sql += "dbo.NVL( q.a51,0) as a51,"
    '    sql += "dbo.NVL( q.a52,0) as a52,"
    '    sql += "dbo.NVL( q.a53,0) as a53,"
    '    sql += "dbo.NVL( q.a54,0) as a54,"
    '    sql += "dbo.NVL( q.a55,0) as a55,"
    '    sql += "dbo.NVL( q.a61,0) as a61,"
    '    sql += "dbo.NVL( q.a62,0) as a62,"
    '    sql += "dbo.NVL( q.a63,0) as a63,"
    '    sql += "dbo.NVL( q.a64,0) as a64,"
    '    sql += "dbo.NVL( q.a65,0) as a65,"
    '    sql += "dbo.NVL( q.a6_71,0) as a6_71,"
    '    sql += "dbo.NVL( q.a6_72,0) as a6_72,"
    '    sql += "dbo.NVL( q.a6_73,0) as a6_73,"
    '    sql += "dbo.NVL( q.a6_74,0) as a6_74,"
    '    sql += "dbo.NVL( q.a6_75,0) as a6_75,"
    '    sql += "dbo.NVL( q.a6_81,0) as a6_81,"
    '    sql += "dbo.NVL( q.a6_82,0) as a6_82,"
    '    sql += "dbo.NVL( q.a6_83,0) as a6_83,"
    '    sql += "dbo.NVL( q.a6_84,0) as a6_84,"
    '    sql += "dbo.NVL( q.a6_85,0) as a6_85,"
    '    sql += "dbo.NVL( q.a71,0) as a71,"
    '    sql += "dbo.NVL( q.a72,0) as a72,"
    '    sql += "dbo.NVL( q.a73,0) as a73,"
    '    sql += "dbo.NVL( q.a81,0) as a81,"
    '    sql += "dbo.NVL( q.a82,0) as a82,"
    '    sql += "dbo.NVL( q.a83,0) as a83,"
    '    sql += "dbo.NVL( q.a84,0) as a84,"
    '    sql += "dbo.NVL( q.a85,0) as a85 "

    '    sql += "from (Select '01' as mid, '3年以下' Name  union "
    '    sql += "Select '02' as mid, '3~5年' Name  union "
    '    sql += "Select '03' as mid, '5~10年' Name  union "
    '    sql += "Select '04' as mid, '10年以上' Name  ) k "
    '    sql += "left outer join ( SELECT c.Mid,  "
    '    sql += "sum(case a.Q1 when 1 then 1 else 0 end) as  a11,"
    '    sql += "sum(case a.Q1 when 2 then 1 else 0 end) as  a12,"
    '    sql += "sum(case a.Q1 when 3 then 1 else 0 end) as  a13,"
    '    sql += "sum(case a.Q1 when 4 then 1 else 0 end) as  a14,"
    '    sql += "sum(case a.Q2 when 1 then 1 else 0 end) as  a21,"
    '    sql += "sum(case a.Q2 when 2 then 1 else 0 end) as  a22,"
    '    sql += "sum(case a.Q2 when 3 then 1 else 0 end) as  a23,"
    '    sql += "sum(case a.Q2 when 4 then 1 else 0 end) as  a24,"
    '    sql += "sum(case a.Q2 when 5 then 1 else 0 end) as  a25,"
    '    sql += "sum(case a.Q3 when 1 then 1 else 0 end) as  a31,"
    '    sql += "sum(case a.Q3 when 2 then 1 else 0 end) as  a32,"
    '    sql += "sum(case a.Q3 when 3 then 1 else 0 end) as  a33,"
    '    sql += "sum(case a.Q3 when 4 then 1 else 0 end) as  a34,"
    '    sql += "sum(case a.Q4 when 1 then 1 else 0 end) as  a41,"
    '    sql += "sum(case a.Q4 when 2 then 1 else 0 end) as  a42,"
    '    sql += "sum(case a.Q4 when 3 then 1 else 0 end) as  a43,"
    '    sql += "sum(case a.Q4 when 4 then 1 else 0 end) as  a44,"
    '    sql += "sum(case a.Q4 when 5 then 1 else 0 end) as  a45,"
    '    sql += "sum(case a.Q5 when 1 then 1 else 0 end) as  a51,"
    '    sql += "sum(case a.Q5 when 2 then 1 else 0 end) as  a52,"
    '    sql += "sum(case a.Q5 when 3 then 1 else 0 end) as  a53,"
    '    sql += "sum(case a.Q5 when 4 then 1 else 0 end) as  a54,"
    '    sql += "sum(case a.Q5 when 5 then 1 else 0 end) as  a55,"
    '    sql += "sum(case a.Q6 when 1 then 1 else 0 end) as  a61,"
    '    sql += "sum(case a.Q6 when 2 then 1 else 0 end) as  a62,"
    '    sql += "sum(case a.Q6 when 3 then 1 else 0 end) as  a63,"
    '    sql += "sum(case a.Q6 when 4 then 1 else 0 end) as  a64,"
    '    sql += "sum(case a.Q6 when 5 then 1 else 0 end) as  a65,"
    '    sql += "sum(case a.Q6_7 when 1 then 1 else 0 end) as a6_71,"
    '    sql += "sum(case a.Q6_7 when 2 then 1 else 0 end) as a6_72,"
    '    sql += "sum(case a.Q6_7 when 3 then 1 else 0 end) as a6_73,"
    '    sql += "sum(case a.Q6_7 when 4 then 1 else 0 end) as a6_74,"
    '    sql += "sum(case a.Q6_7 when 5 then 1 else 0 end) as a6_75,"
    '    sql += "sum(case a.Q6_8 when 1 then 1 else 0 end) as a6_81,"
    '    sql += "sum(case a.Q6_8 when 2 then 1 else 0 end) as a6_82,"
    '    sql += "sum(case a.Q6_8 when 3 then 1 else 0 end) as a6_83,"
    '    sql += "sum(case a.Q6_8 when 4 then 1 else 0 end) as a6_84,"
    '    sql += "sum(case a.Q6_8 when 5 then 1 else 0 end) as a6_85,  "
    '    sql += "sum(case a.Q7 when 1 then 1 else 0 end) as  a71,"
    '    sql += "sum(case a.Q7 when 2 then 1 else 0 end) as  a72,"
    '    sql += "sum(case a.Q7 when 3 then 1 else 0 end) as  a73,"
    '    sql += "sum(case a.Q8 when '1' then 1 else 0 end) as a81,"
    '    sql += "sum(case a.Q8 when '2' then 1 else 0 end) as a82,"
    '    sql += "sum(case a.Q8 when '3' then 1 else 0 end) as a83,"
    '    sql += "sum(case a.Q8 when '4' then 1 else 0 end) as a84,"
    '    sql += "sum(case a.Q8 when '5' then 1 else 0 end) as a85 "
    '    sql += "From Stud_QuestionFin a left outer join Class_StudentsOfClass b on a.SOCID=b.socid "
    '    sql += "left outer join (select socid,case "
    '    sql += "when dbo.NVL(q61,0)<3 then '01' "
    '    sql += "when dbo.NVL(q61,0)<5 then '02' "
    '    sql += "when dbo.NVL(q61,0)<10 then '03' "
    '    sql += "else '04' end as mid from Stud_TrainBG  ) c on a.SOCID=c.socid  "
    '    sql += "INNER JOIN (select rid, PlanID, Years, OCID, SEQNO  from Class_ClassInfo where 1=1 "
    '    If RID <> "" Then
    '        sql += "AND RID LIKE '" + RID + "%'"
    '    End If
    '    If OCID <> "" Then
    '        sql += "AND OCID='" + OCID + "'"
    '    End If
    '    sql += ") cc ON b.OCID=cc.OCID JOIN (select PlanID, Years from ID_Plan where TPlanID='" + TPlanID + "' "
    '    If Years <> "" Then
    '        sql += "AND Years='" + Years + "'"
    '    End If
    '    sql += ") ip ON cc.PlanID = ip.PlanID AND cc.Years = dbo.SUBSTR(ip.Years,3,2) "
    '    sql += "JOIN Plan_Planinfo pp ON pp.Planid =cc.Planid  AND pp.RID = cc.RID AND pp.SEQNO = cc.SEQNO "
    '    If PackageType <> "" Then
    '        sql += "AND pp.PackageType ='" + PackageType + "'"
    '    End If
    '    sql += "JOIN (select * from org_orginfo  where 1=1 "
    '    If SearchPlan <> "" Then
    '        sql += "and OrgKind2 ='" + SearchPlan + "' "
    '    End If
    '    sql += ") oo ON pp.COMIDNO =oo.COMIDNO group by c.Mid ) q on q.mid=k.mid "
    '    sql += "order by nlssort(NAME,'NLS_SORT=TCHINESE_STROKE_M') "

    '    resDt = DbAccess.GetDataTable(sql, objconn)
    '    'sourceDT = resDt
    '    'If sourceDT.Rows.Count > 0 Then
    '    '    PName = sourceDT.Rows(0)("PName")
    '    '    plankind = sourceDT.Rows(0)("plankind")
    '    'End If

    '    setItemArray() '題目內容

    '    tmpDT = RotationDT(resDt)
    '    Return tmpDT

    'End Function

    'Private Function db_R16(ByVal TPlanID As String, ByVal Years As String, ByVal OCID As String, ByVal RID As String, ByVal SearchPlan As String, ByVal PackageType As String) As DataTable
    '    Dim tmpDT As DataTable
    '    Dim resDt As DataTable = Nothing
    '    Dim sql As String = ""

    '    sql = ""
    '    sql = "select convert(varchar, CONVERT(numeric,  '" + Years + "') -1911 ) + '年度　訓後動態調查統計表 地理分佈' AS PName, k.Name, "
    '    sql += "( Select a.PlanName +dbo.NVL(case when a.TPlanID='28' then (select '('+name+')' from v_orgkind1 WHERE VALUE='" & SearchPlan & "') end,'') from key_plan a where a.TPlanID='" & TPlanID & "') as plankind, "
    '    sql += "dbo.NVL( q.a11,0) as a11,"
    '    sql += "dbo.NVL( q.a12,0) as a12,"
    '    sql += "dbo.NVL( q.a13,0) as a13,"
    '    sql += "dbo.NVL( q.a14,0) as a14,"
    '    sql += "dbo.NVL( q.a21,0) as a21,"
    '    sql += "dbo.NVL( q.a22,0) as a22,"
    '    sql += "dbo.NVL( q.a23,0) as a23,"
    '    sql += "dbo.NVL( q.a24,0) as a24,"
    '    sql += "dbo.NVL( q.a25,0) as a25,"
    '    sql += "dbo.NVL( q.a31,0) as a31,"
    '    sql += "dbo.NVL( q.a32,0) as a32,"
    '    sql += "dbo.NVL( q.a33,0) as a33,"
    '    sql += "dbo.NVL( q.a34,0) as a34,"
    '    sql += "dbo.NVL( q.a41,0) as a41,"
    '    sql += "dbo.NVL( q.a42,0) as a42,"
    '    sql += "dbo.NVL( q.a43,0) as a43,"
    '    sql += "dbo.NVL( q.a44,0) as a44,"
    '    sql += "dbo.NVL( q.a45,0) as a45,"
    '    sql += "dbo.NVL( q.a51,0) as a51,"
    '    sql += "dbo.NVL( q.a52,0) as a52,"
    '    sql += "dbo.NVL( q.a53,0) as a53,"
    '    sql += "dbo.NVL( q.a54,0) as a54,"
    '    sql += "dbo.NVL( q.a55,0) as a55,"
    '    sql += "dbo.NVL( q.a61,0) as a61,"
    '    sql += "dbo.NVL( q.a62,0) as a62,"
    '    sql += "dbo.NVL( q.a63,0) as a63,"
    '    sql += "dbo.NVL( q.a64,0) as a64,"
    '    sql += "dbo.NVL( q.a65,0) as a65,"
    '    sql += "dbo.NVL( q.a6_71,0) as a6_71,"
    '    sql += "dbo.NVL( q.a6_72,0) as a6_72,"
    '    sql += "dbo.NVL( q.a6_73,0) as a6_73,"
    '    sql += "dbo.NVL( q.a6_74,0) as a6_74,"
    '    sql += "dbo.NVL( q.a6_75,0) as a6_75,"
    '    sql += "dbo.NVL( q.a6_81,0) as a6_81,"
    '    sql += "dbo.NVL( q.a6_82,0) as a6_82,"
    '    sql += "dbo.NVL( q.a6_83,0) as a6_83,"
    '    sql += "dbo.NVL( q.a6_84,0) as a6_84,"
    '    sql += "dbo.NVL( q.a6_85,0) as a6_85,"
    '    sql += "dbo.NVL( q.a71,0) as a71,"
    '    sql += "dbo.NVL( q.a72,0) as a72,"
    '    sql += "dbo.NVL( q.a73,0) as a73,"
    '    sql += "dbo.NVL( q.a81,0) as a81,"
    '    sql += "dbo.NVL( q.a82,0) as a82,"
    '    sql += "dbo.NVL( q.a83,0) as a83,"
    '    sql += "dbo.NVL( q.a84,0) as a84,"
    '    sql += "dbo.NVL( q.a85,0) as a85 "

    '    sql += "from (Select ctid as mid, ctname as Name from id_city)  k "
    '    sql += "left outer join ( SELECT c.Mid,  "
    '    sql += "sum(case a.Q1 when 1 then 1 else 0 end) as  a11,"
    '    sql += "sum(case a.Q1 when 2 then 1 else 0 end) as  a12,"
    '    sql += "sum(case a.Q1 when 3 then 1 else 0 end) as  a13,"
    '    sql += "sum(case a.Q1 when 4 then 1 else 0 end) as  a14,"
    '    sql += "sum(case a.Q2 when 1 then 1 else 0 end) as  a21,"
    '    sql += "sum(case a.Q2 when 2 then 1 else 0 end) as  a22,"
    '    sql += "sum(case a.Q2 when 3 then 1 else 0 end) as  a23,"
    '    sql += "sum(case a.Q2 when 4 then 1 else 0 end) as  a24,"
    '    sql += "sum(case a.Q2 when 5 then 1 else 0 end) as  a25,"
    '    sql += "sum(case a.Q3 when 1 then 1 else 0 end) as  a31,"
    '    sql += "sum(case a.Q3 when 2 then 1 else 0 end) as  a32,"
    '    sql += "sum(case a.Q3 when 3 then 1 else 0 end) as  a33,"
    '    sql += "sum(case a.Q3 when 4 then 1 else 0 end) as  a34,"
    '    sql += "sum(case a.Q4 when 1 then 1 else 0 end) as  a41,"
    '    sql += "sum(case a.Q4 when 2 then 1 else 0 end) as  a42,"
    '    sql += "sum(case a.Q4 when 3 then 1 else 0 end) as  a43,"
    '    sql += "sum(case a.Q4 when 4 then 1 else 0 end) as  a44,"
    '    sql += "sum(case a.Q4 when 5 then 1 else 0 end) as  a45,"
    '    sql += "sum(case a.Q5 when 1 then 1 else 0 end) as  a51,"
    '    sql += "sum(case a.Q5 when 2 then 1 else 0 end) as  a52,"
    '    sql += "sum(case a.Q5 when 3 then 1 else 0 end) as  a53,"
    '    sql += "sum(case a.Q5 when 4 then 1 else 0 end) as  a54,"
    '    sql += "sum(case a.Q5 when 5 then 1 else 0 end) as  a55,"
    '    sql += "sum(case a.Q6 when 1 then 1 else 0 end) as  a61,"
    '    sql += "sum(case a.Q6 when 2 then 1 else 0 end) as  a62,"
    '    sql += "sum(case a.Q6 when 3 then 1 else 0 end) as  a63,"
    '    sql += "sum(case a.Q6 when 4 then 1 else 0 end) as  a64,"
    '    sql += "sum(case a.Q6 when 5 then 1 else 0 end) as  a65,"
    '    sql += "sum(case a.Q6_7 when 1 then 1 else 0 end) as a6_71,"
    '    sql += "sum(case a.Q6_7 when 2 then 1 else 0 end) as a6_72,"
    '    sql += "sum(case a.Q6_7 when 3 then 1 else 0 end) as a6_73,"
    '    sql += "sum(case a.Q6_7 when 4 then 1 else 0 end) as a6_74,"
    '    sql += "sum(case a.Q6_7 when 5 then 1 else 0 end) as a6_75,"
    '    sql += "sum(case a.Q6_8 when 1 then 1 else 0 end) as a6_81,"
    '    sql += "sum(case a.Q6_8 when 2 then 1 else 0 end) as a6_82,"
    '    sql += "sum(case a.Q6_8 when 3 then 1 else 0 end) as a6_83,"
    '    sql += "sum(case a.Q6_8 when 4 then 1 else 0 end) as a6_84,"
    '    sql += "sum(case a.Q6_8 when 5 then 1 else 0 end) as a6_85,  "
    '    sql += "sum(case a.Q7 when 1 then 1 else 0 end) as  a71,"
    '    sql += "sum(case a.Q7 when 2 then 1 else 0 end) as  a72,"
    '    sql += "sum(case a.Q7 when 3 then 1 else 0 end) as  a73,"
    '    sql += "sum(case a.Q8 when '1' then 1 else 0 end) as a81,"
    '    sql += "sum(case a.Q8 when '2' then 1 else 0 end) as a82,"
    '    sql += "sum(case a.Q8 when '3' then 1 else 0 end) as a83,"
    '    sql += "sum(case a.Q8 when '4' then 1 else 0 end) as a84,"
    '    sql += "sum(case a.Q8 when '5' then 1 else 0 end) as a85 "
    '    sql += "From Stud_QuestionFin a left outer join Class_StudentsOfClass b on a.SOCID=b.socid "
    '    sql += "left outer join "
    '    sql += "(select  a.sid, b.ctid as mid from Stud_SubData a JOIN id_zip b on a.ZipCode1=b.zipcode ) c on b.SID=c.sid "
    '    sql += "INNER JOIN (select rid, PlanID, Years, OCID, SEQNO  from Class_ClassInfo where 1=1 "
    '    If RID <> "" Then
    '        sql += "AND RID LIKE '" + RID + "%'"
    '    End If
    '    If OCID <> "" Then
    '        sql += "AND OCID='" + OCID + "'"
    '    End If
    '    sql += ") cc ON b.OCID=cc.OCID JOIN (select PlanID, Years from ID_Plan where TPlanID='" + TPlanID + "' "
    '    If Years <> "" Then
    '        sql += "AND Years='" + Years + "'"
    '    End If
    '    sql += ") ip ON cc.PlanID = ip.PlanID AND cc.Years = dbo.SUBSTR(ip.Years,3,2) "
    '    sql += "JOIN Plan_Planinfo pp ON pp.Planid =cc.Planid  AND pp.RID = cc.RID AND pp.SEQNO = cc.SEQNO "
    '    If PackageType <> "" Then
    '        sql += "AND pp.PackageType ='" + PackageType + "'"
    '    End If
    '    sql += "JOIN (select * from org_orginfo  where 1=1 "
    '    If SearchPlan <> "" Then
    '        sql += "and OrgKind2 ='" + SearchPlan + "' "
    '    End If
    '    sql += ") oo ON pp.COMIDNO =oo.COMIDNO group by c.Mid ) q on q.mid=k.mid "
    '    sql += "order by nlssort(NAME,'NLS_SORT=TCHINESE_STROKE_M') "

    '    resDt = DbAccess.GetDataTable(sql, objconn)
    '    'sourceDT = resDt
    '    'If sourceDT.Rows.Count > 0 Then
    '    '    PName = sourceDT.Rows(0)("PName")
    '    '    plankind = sourceDT.Rows(0)("plankind")
    '    'End If

    '    setItemArray() '題目內容

    '    tmpDT = RotationDT(resDt)
    '    Return tmpDT

    'End Function

    'Private Function db_R17(ByVal TPlanID As String, ByVal Years As String, ByVal OCID As String, ByVal RID As String, ByVal SearchPlan As String, ByVal PackageType As String) As DataTable
    '    Dim tmpDT As DataTable
    '    Dim resDt As DataTable = Nothing
    '    Dim sql As String = ""

    '    sql = ""
    '    sql = "select convert(varchar, CONVERT(numeric,  '" + Years + "') -1911 ) + '年度　訓後動態調查統計表 公司行業別' AS PName, k.Name, "
    '    sql += "( Select a.PlanName +dbo.NVL(case when a.TPlanID='28' then (select '('+name+')' from v_orgkind1 WHERE VALUE='" & SearchPlan & "') end,'') from key_plan a where a.TPlanID='" & TPlanID & "') as plankind, "
    '    sql += "dbo.NVL( q.a11,0) as a11,"
    '    sql += "dbo.NVL( q.a12,0) as a12,"
    '    sql += "dbo.NVL( q.a13,0) as a13,"
    '    sql += "dbo.NVL( q.a14,0) as a14,"
    '    sql += "dbo.NVL( q.a21,0) as a21,"
    '    sql += "dbo.NVL( q.a22,0) as a22,"
    '    sql += "dbo.NVL( q.a23,0) as a23,"
    '    sql += "dbo.NVL( q.a24,0) as a24,"
    '    sql += "dbo.NVL( q.a25,0) as a25,"
    '    sql += "dbo.NVL( q.a31,0) as a31,"
    '    sql += "dbo.NVL( q.a32,0) as a32,"
    '    sql += "dbo.NVL( q.a33,0) as a33,"
    '    sql += "dbo.NVL( q.a34,0) as a34,"
    '    sql += "dbo.NVL( q.a41,0) as a41,"
    '    sql += "dbo.NVL( q.a42,0) as a42,"
    '    sql += "dbo.NVL( q.a43,0) as a43,"
    '    sql += "dbo.NVL( q.a44,0) as a44,"
    '    sql += "dbo.NVL( q.a45,0) as a45,"
    '    sql += "dbo.NVL( q.a51,0) as a51,"
    '    sql += "dbo.NVL( q.a52,0) as a52,"
    '    sql += "dbo.NVL( q.a53,0) as a53,"
    '    sql += "dbo.NVL( q.a54,0) as a54,"
    '    sql += "dbo.NVL( q.a55,0) as a55,"
    '    sql += "dbo.NVL( q.a61,0) as a61,"
    '    sql += "dbo.NVL( q.a62,0) as a62,"
    '    sql += "dbo.NVL( q.a63,0) as a63,"
    '    sql += "dbo.NVL( q.a64,0) as a64,"
    '    sql += "dbo.NVL( q.a65,0) as a65,"
    '    sql += "dbo.NVL( q.a6_71,0) as a6_71,"
    '    sql += "dbo.NVL( q.a6_72,0) as a6_72,"
    '    sql += "dbo.NVL( q.a6_73,0) as a6_73,"
    '    sql += "dbo.NVL( q.a6_74,0) as a6_74,"
    '    sql += "dbo.NVL( q.a6_75,0) as a6_75,"
    '    sql += "dbo.NVL( q.a6_81,0) as a6_81,"
    '    sql += "dbo.NVL( q.a6_82,0) as a6_82,"
    '    sql += "dbo.NVL( q.a6_83,0) as a6_83,"
    '    sql += "dbo.NVL( q.a6_84,0) as a6_84,"
    '    sql += "dbo.NVL( q.a6_85,0) as a6_85,"
    '    sql += "dbo.NVL( q.a71,0) as a71,"
    '    sql += "dbo.NVL( q.a72,0) as a72,"
    '    sql += "dbo.NVL( q.a73,0) as a73,"
    '    sql += "dbo.NVL( q.a81,0) as a81,"
    '    sql += "dbo.NVL( q.a82,0) as a82,"
    '    sql += "dbo.NVL( q.a83,0) as a83,"
    '    sql += "dbo.NVL( q.a84,0) as a84,"
    '    sql += "dbo.NVL( q.a85,0) as a85 "

    '    sql += "from (SELECT tradeid as mid, tradename as name FROM key_Trade) k "
    '    sql += "left outer join ( SELECT c.Mid,  "
    '    sql += "sum(case a.Q1 when 1 then 1 else 0 end) as  a11,"
    '    sql += "sum(case a.Q1 when 2 then 1 else 0 end) as  a12,"
    '    sql += "sum(case a.Q1 when 3 then 1 else 0 end) as  a13,"
    '    sql += "sum(case a.Q1 when 4 then 1 else 0 end) as  a14,"
    '    sql += "sum(case a.Q2 when 1 then 1 else 0 end) as  a21,"
    '    sql += "sum(case a.Q2 when 2 then 1 else 0 end) as  a22,"
    '    sql += "sum(case a.Q2 when 3 then 1 else 0 end) as  a23,"
    '    sql += "sum(case a.Q2 when 4 then 1 else 0 end) as  a24,"
    '    sql += "sum(case a.Q2 when 5 then 1 else 0 end) as  a25,"
    '    sql += "sum(case a.Q3 when 1 then 1 else 0 end) as  a31,"
    '    sql += "sum(case a.Q3 when 2 then 1 else 0 end) as  a32,"
    '    sql += "sum(case a.Q3 when 3 then 1 else 0 end) as  a33,"
    '    sql += "sum(case a.Q3 when 4 then 1 else 0 end) as  a34,"
    '    sql += "sum(case a.Q4 when 1 then 1 else 0 end) as  a41,"
    '    sql += "sum(case a.Q4 when 2 then 1 else 0 end) as  a42,"
    '    sql += "sum(case a.Q4 when 3 then 1 else 0 end) as  a43,"
    '    sql += "sum(case a.Q4 when 4 then 1 else 0 end) as  a44,"
    '    sql += "sum(case a.Q4 when 5 then 1 else 0 end) as  a45,"
    '    sql += "sum(case a.Q5 when 1 then 1 else 0 end) as  a51,"
    '    sql += "sum(case a.Q5 when 2 then 1 else 0 end) as  a52,"
    '    sql += "sum(case a.Q5 when 3 then 1 else 0 end) as  a53,"
    '    sql += "sum(case a.Q5 when 4 then 1 else 0 end) as  a54,"
    '    sql += "sum(case a.Q5 when 5 then 1 else 0 end) as  a55,"
    '    sql += "sum(case a.Q6 when 1 then 1 else 0 end) as  a61,"
    '    sql += "sum(case a.Q6 when 2 then 1 else 0 end) as  a62,"
    '    sql += "sum(case a.Q6 when 3 then 1 else 0 end) as  a63,"
    '    sql += "sum(case a.Q6 when 4 then 1 else 0 end) as  a64,"
    '    sql += "sum(case a.Q6 when 5 then 1 else 0 end) as  a65,"
    '    sql += "sum(case a.Q6_7 when 1 then 1 else 0 end) as a6_71,"
    '    sql += "sum(case a.Q6_7 when 2 then 1 else 0 end) as a6_72,"
    '    sql += "sum(case a.Q6_7 when 3 then 1 else 0 end) as a6_73,"
    '    sql += "sum(case a.Q6_7 when 4 then 1 else 0 end) as a6_74,"
    '    sql += "sum(case a.Q6_7 when 5 then 1 else 0 end) as a6_75,"
    '    sql += "sum(case a.Q6_8 when 1 then 1 else 0 end) as a6_81,"
    '    sql += "sum(case a.Q6_8 when 2 then 1 else 0 end) as a6_82,"
    '    sql += "sum(case a.Q6_8 when 3 then 1 else 0 end) as a6_83,"
    '    sql += "sum(case a.Q6_8 when 4 then 1 else 0 end) as a6_84,"
    '    sql += "sum(case a.Q6_8 when 5 then 1 else 0 end) as a6_85,  "
    '    sql += "sum(case a.Q7 when 1 then 1 else 0 end) as  a71,"
    '    sql += "sum(case a.Q7 when 2 then 1 else 0 end) as  a72,"
    '    sql += "sum(case a.Q7 when 3 then 1 else 0 end) as  a73,"
    '    sql += "sum(case a.Q8 when '1' then 1 else 0 end) as a81,"
    '    sql += "sum(case a.Q8 when '2' then 1 else 0 end) as a82,"
    '    sql += "sum(case a.Q8 when '3' then 1 else 0 end) as a83,"
    '    sql += "sum(case a.Q8 when '4' then 1 else 0 end) as a84,"
    '    sql += "sum(case a.Q8 when '5' then 1 else 0 end) as a85 "
    '    sql += "From Stud_QuestionFin a left outer join Class_StudentsOfClass b on a.SOCID=b.socid "
    '    sql += "left outer join "
    '    sql += "( select socid, q4 as mid from Stud_TrainBG ) c on a.SOCID=c.socid "
    '    sql += "INNER JOIN (select rid, PlanID, Years, OCID, SEQNO  from Class_ClassInfo where 1=1 "
    '    If RID <> "" Then
    '        sql += "AND RID LIKE '" + RID + "%'"
    '    End If
    '    If OCID <> "" Then
    '        sql += "AND OCID='" + OCID + "'"
    '    End If
    '    sql += ") cc ON b.OCID=cc.OCID JOIN (select PlanID, Years from ID_Plan where TPlanID='" + TPlanID + "' "
    '    If Years <> "" Then
    '        sql += "AND Years='" + Years + "'"
    '    End If
    '    sql += ") ip ON cc.PlanID = ip.PlanID AND cc.Years = dbo.SUBSTR(ip.Years,3,2) "
    '    sql += "JOIN Plan_Planinfo pp ON pp.Planid =cc.Planid  AND pp.RID = cc.RID AND pp.SEQNO = cc.SEQNO "
    '    If PackageType <> "" Then
    '        sql += "AND pp.PackageType ='" + PackageType + "'"
    '    End If
    '    sql += "JOIN (select * from org_orginfo  where 1=1 "
    '    If SearchPlan <> "" Then
    '        sql += "and OrgKind2 ='" + SearchPlan + "' "
    '    End If
    '    sql += ") oo ON pp.COMIDNO =oo.COMIDNO group by c.Mid ) q on q.mid=k.mid "
    '    sql += "order by nlssort(NAME,'NLS_SORT=TCHINESE_STROKE_M') "

    '    resDt = DbAccess.GetDataTable(sql, objconn)
    '    'sourceDT = resDt
    '    'If sourceDT.Rows.Count > 0 Then
    '    '    PName = sourceDT.Rows(0)("PName")
    '    '    plankind = sourceDT.Rows(0)("plankind")
    '    'End If

    '    setItemArray() '題目內容

    '    tmpDT = RotationDT(resDt)
    '    Return tmpDT

    'End Function

    'Private Function db_R18(ByVal TPlanID As String, ByVal Years As String, ByVal OCID As String, ByVal RID As String, ByVal SearchPlan As String, ByVal PackageType As String) As DataTable
    '    Dim tmpDT As DataTable
    '    Dim resDt As DataTable = Nothing
    '    Dim sql As String = ""

    '    sql = ""
    '    sql = "select convert(varchar, CONVERT(numeric,  '" + Years + "') -1911 ) + '年度　訓後動態調查統計表 公司規模' AS PName, k.Name, "
    '    sql += "( Select a.PlanName +dbo.NVL(case when a.TPlanID='28' then (select '('+name+')' from v_orgkind1 WHERE VALUE='" & SearchPlan & "') end,'') from key_plan a where a.TPlanID='" & TPlanID & "') as plankind, "
    '    sql += "dbo.NVL( q.a11,0) as a11,"
    '    sql += "dbo.NVL( q.a12,0) as a12,"
    '    sql += "dbo.NVL( q.a13,0) as a13,"
    '    sql += "dbo.NVL( q.a14,0) as a14,"
    '    sql += "dbo.NVL( q.a21,0) as a21,"
    '    sql += "dbo.NVL( q.a22,0) as a22,"
    '    sql += "dbo.NVL( q.a23,0) as a23,"
    '    sql += "dbo.NVL( q.a24,0) as a24,"
    '    sql += "dbo.NVL( q.a25,0) as a25,"
    '    sql += "dbo.NVL( q.a31,0) as a31,"
    '    sql += "dbo.NVL( q.a32,0) as a32,"
    '    sql += "dbo.NVL( q.a33,0) as a33,"
    '    sql += "dbo.NVL( q.a34,0) as a34,"
    '    sql += "dbo.NVL( q.a41,0) as a41,"
    '    sql += "dbo.NVL( q.a42,0) as a42,"
    '    sql += "dbo.NVL( q.a43,0) as a43,"
    '    sql += "dbo.NVL( q.a44,0) as a44,"
    '    sql += "dbo.NVL( q.a45,0) as a45,"
    '    sql += "dbo.NVL( q.a51,0) as a51,"
    '    sql += "dbo.NVL( q.a52,0) as a52,"
    '    sql += "dbo.NVL( q.a53,0) as a53,"
    '    sql += "dbo.NVL( q.a54,0) as a54,"
    '    sql += "dbo.NVL( q.a55,0) as a55,"
    '    sql += "dbo.NVL( q.a61,0) as a61,"
    '    sql += "dbo.NVL( q.a62,0) as a62,"
    '    sql += "dbo.NVL( q.a63,0) as a63,"
    '    sql += "dbo.NVL( q.a64,0) as a64,"
    '    sql += "dbo.NVL( q.a65,0) as a65,"
    '    sql += "dbo.NVL( q.a6_71,0) as a6_71,"
    '    sql += "dbo.NVL( q.a6_72,0) as a6_72,"
    '    sql += "dbo.NVL( q.a6_73,0) as a6_73,"
    '    sql += "dbo.NVL( q.a6_74,0) as a6_74,"
    '    sql += "dbo.NVL( q.a6_75,0) as a6_75,"
    '    sql += "dbo.NVL( q.a6_81,0) as a6_81,"
    '    sql += "dbo.NVL( q.a6_82,0) as a6_82,"
    '    sql += "dbo.NVL( q.a6_83,0) as a6_83,"
    '    sql += "dbo.NVL( q.a6_84,0) as a6_84,"
    '    sql += "dbo.NVL( q.a6_85,0) as a6_85,"
    '    sql += "dbo.NVL( q.a71,0) as a71,"
    '    sql += "dbo.NVL( q.a72,0) as a72,"
    '    sql += "dbo.NVL( q.a73,0) as a73,"
    '    sql += "dbo.NVL( q.a81,0) as a81,"
    '    sql += "dbo.NVL( q.a82,0) as a82,"
    '    sql += "dbo.NVL( q.a83,0) as a83,"
    '    sql += "dbo.NVL( q.a84,0) as a84,"
    '    sql += "dbo.NVL( q.a85,0) as a85 "

    '    sql += "from (Select 1 as mid, '屬於中小企業' Name  union "
    '    sql += "Select 0 as mid, '非中小企業' Name ) k "
    '    sql += "left outer join ( SELECT c.Mid,  "
    '    sql += "sum(case a.Q1 when 1 then 1 else 0 end) as  a11,"
    '    sql += "sum(case a.Q1 when 2 then 1 else 0 end) as  a12,"
    '    sql += "sum(case a.Q1 when 3 then 1 else 0 end) as  a13,"
    '    sql += "sum(case a.Q1 when 4 then 1 else 0 end) as  a14,"
    '    sql += "sum(case a.Q2 when 1 then 1 else 0 end) as  a21,"
    '    sql += "sum(case a.Q2 when 2 then 1 else 0 end) as  a22,"
    '    sql += "sum(case a.Q2 when 3 then 1 else 0 end) as  a23,"
    '    sql += "sum(case a.Q2 when 4 then 1 else 0 end) as  a24,"
    '    sql += "sum(case a.Q2 when 5 then 1 else 0 end) as  a25,"
    '    sql += "sum(case a.Q3 when 1 then 1 else 0 end) as  a31,"
    '    sql += "sum(case a.Q3 when 2 then 1 else 0 end) as  a32,"
    '    sql += "sum(case a.Q3 when 3 then 1 else 0 end) as  a33,"
    '    sql += "sum(case a.Q3 when 4 then 1 else 0 end) as  a34,"
    '    sql += "sum(case a.Q4 when 1 then 1 else 0 end) as  a41,"
    '    sql += "sum(case a.Q4 when 2 then 1 else 0 end) as  a42,"
    '    sql += "sum(case a.Q4 when 3 then 1 else 0 end) as  a43,"
    '    sql += "sum(case a.Q4 when 4 then 1 else 0 end) as  a44,"
    '    sql += "sum(case a.Q4 when 5 then 1 else 0 end) as  a45,"
    '    sql += "sum(case a.Q5 when 1 then 1 else 0 end) as  a51,"
    '    sql += "sum(case a.Q5 when 2 then 1 else 0 end) as  a52,"
    '    sql += "sum(case a.Q5 when 3 then 1 else 0 end) as  a53,"
    '    sql += "sum(case a.Q5 when 4 then 1 else 0 end) as  a54,"
    '    sql += "sum(case a.Q5 when 5 then 1 else 0 end) as  a55,"
    '    sql += "sum(case a.Q6 when 1 then 1 else 0 end) as  a61,"
    '    sql += "sum(case a.Q6 when 2 then 1 else 0 end) as  a62,"
    '    sql += "sum(case a.Q6 when 3 then 1 else 0 end) as  a63,"
    '    sql += "sum(case a.Q6 when 4 then 1 else 0 end) as  a64,"
    '    sql += "sum(case a.Q6 when 5 then 1 else 0 end) as  a65,"
    '    sql += "sum(case a.Q6_7 when 1 then 1 else 0 end) as a6_71,"
    '    sql += "sum(case a.Q6_7 when 2 then 1 else 0 end) as a6_72,"
    '    sql += "sum(case a.Q6_7 when 3 then 1 else 0 end) as a6_73,"
    '    sql += "sum(case a.Q6_7 when 4 then 1 else 0 end) as a6_74,"
    '    sql += "sum(case a.Q6_7 when 5 then 1 else 0 end) as a6_75,"
    '    sql += "sum(case a.Q6_8 when 1 then 1 else 0 end) as a6_81,"
    '    sql += "sum(case a.Q6_8 when 2 then 1 else 0 end) as a6_82,"
    '    sql += "sum(case a.Q6_8 when 3 then 1 else 0 end) as a6_83,"
    '    sql += "sum(case a.Q6_8 when 4 then 1 else 0 end) as a6_84,"
    '    sql += "sum(case a.Q6_8 when 5 then 1 else 0 end) as a6_85,  "
    '    sql += "sum(case a.Q7 when 1 then 1 else 0 end) as  a71,"
    '    sql += "sum(case a.Q7 when 2 then 1 else 0 end) as  a72,"
    '    sql += "sum(case a.Q7 when 3 then 1 else 0 end) as  a73,"
    '    sql += "sum(case a.Q8 when '1' then 1 else 0 end) as a81,"
    '    sql += "sum(case a.Q8 when '2' then 1 else 0 end) as a82,"
    '    sql += "sum(case a.Q8 when '3' then 1 else 0 end) as a83,"
    '    sql += "sum(case a.Q8 when '4' then 1 else 0 end) as a84,"
    '    sql += "sum(case a.Q8 when '5' then 1 else 0 end) as a85 "
    '    sql += "From Stud_QuestionFin a left outer join Class_StudentsOfClass b on a.SOCID=b.socid "
    '    sql += "left outer join "
    '    sql += "( select socid, q5 as mid from Stud_TrainBG ) c on a.SOCID=c.socid "
    '    sql += "INNER JOIN (select rid, PlanID, Years, OCID, SEQNO  from Class_ClassInfo where 1=1 "
    '    If RID <> "" Then
    '        sql += "AND RID LIKE '" + RID + "%'"
    '    End If
    '    If OCID <> "" Then
    '        sql += "AND OCID='" + OCID + "'"
    '    End If
    '    sql += ") cc ON b.OCID=cc.OCID JOIN (select PlanID, Years from ID_Plan where TPlanID='" + TPlanID + "' "
    '    If Years <> "" Then
    '        sql += "AND Years='" + Years + "'"
    '    End If
    '    sql += ") ip ON cc.PlanID = ip.PlanID AND cc.Years = dbo.SUBSTR(ip.Years,3,2) "
    '    sql += "JOIN Plan_Planinfo pp ON pp.Planid =cc.Planid  AND pp.RID = cc.RID AND pp.SEQNO = cc.SEQNO "
    '    If PackageType <> "" Then
    '        sql += "AND pp.PackageType ='" + PackageType + "'"
    '    End If
    '    sql += "JOIN (select * from org_orginfo  where 1=1 "
    '    If SearchPlan <> "" Then
    '        sql += "and OrgKind2 ='" + SearchPlan + "' "
    '    End If
    '    sql += ") oo ON pp.COMIDNO =oo.COMIDNO group by c.Mid ) q on q.mid=k.mid "
    '    sql += "order by nlssort(NAME,'NLS_SORT=TCHINESE_STROKE_M') "

    '    resDt = DbAccess.GetDataTable(sql, objconn)
    '    'sourceDT = resDt
    '    'If sourceDT.Rows.Count > 0 Then
    '    '    PName = sourceDT.Rows(0)("PName")
    '    '    plankind = sourceDT.Rows(0)("plankind")
    '    'End If

    '    setItemArray() '題目內容

    '    tmpDT = RotationDT(resDt)
    '    Return tmpDT

    'End Function

    'Private Function db_R19(ByVal TPlanID As String, ByVal Years As String, ByVal OCID As String, ByVal RID As String, ByVal SearchPlan As String, ByVal PackageType As String) As DataTable
    '    Dim tmpDT As DataTable
    '    Dim resDt As DataTable = Nothing
    '    Dim sql As String = ""

    '    sql = ""
    '    sql = "select convert(varchar, CONVERT(numeric,  '" + Years + "') -1911 ) + '年度　訓後動態調查統計表 參訓動機' AS PName, k.Name, "
    '    sql += "( Select a.PlanName +dbo.NVL(case when a.TPlanID='28' then (select '('+name+')' from v_orgkind1 WHERE VALUE='" & SearchPlan & "') end,'') from key_plan a where a.TPlanID='" & TPlanID & "') as plankind, "
    '    sql += "dbo.NVL( q.a11,0) as a11,"
    '    sql += "dbo.NVL( q.a12,0) as a12,"
    '    sql += "dbo.NVL( q.a13,0) as a13,"
    '    sql += "dbo.NVL( q.a14,0) as a14,"
    '    sql += "dbo.NVL( q.a21,0) as a21,"
    '    sql += "dbo.NVL( q.a22,0) as a22,"
    '    sql += "dbo.NVL( q.a23,0) as a23,"
    '    sql += "dbo.NVL( q.a24,0) as a24,"
    '    sql += "dbo.NVL( q.a25,0) as a25,"
    '    sql += "dbo.NVL( q.a31,0) as a31,"
    '    sql += "dbo.NVL( q.a32,0) as a32,"
    '    sql += "dbo.NVL( q.a33,0) as a33,"
    '    sql += "dbo.NVL( q.a34,0) as a34,"
    '    sql += "dbo.NVL( q.a41,0) as a41,"
    '    sql += "dbo.NVL( q.a42,0) as a42,"
    '    sql += "dbo.NVL( q.a43,0) as a43,"
    '    sql += "dbo.NVL( q.a44,0) as a44,"
    '    sql += "dbo.NVL( q.a45,0) as a45,"
    '    sql += "dbo.NVL( q.a51,0) as a51,"
    '    sql += "dbo.NVL( q.a52,0) as a52,"
    '    sql += "dbo.NVL( q.a53,0) as a53,"
    '    sql += "dbo.NVL( q.a54,0) as a54,"
    '    sql += "dbo.NVL( q.a55,0) as a55,"
    '    sql += "dbo.NVL( q.a61,0) as a61,"
    '    sql += "dbo.NVL( q.a62,0) as a62,"
    '    sql += "dbo.NVL( q.a63,0) as a63,"
    '    sql += "dbo.NVL( q.a64,0) as a64,"
    '    sql += "dbo.NVL( q.a65,0) as a65,"
    '    sql += "dbo.NVL( q.a6_71,0) as a6_71,"
    '    sql += "dbo.NVL( q.a6_72,0) as a6_72,"
    '    sql += "dbo.NVL( q.a6_73,0) as a6_73,"
    '    sql += "dbo.NVL( q.a6_74,0) as a6_74,"
    '    sql += "dbo.NVL( q.a6_75,0) as a6_75,"
    '    sql += "dbo.NVL( q.a6_81,0) as a6_81,"
    '    sql += "dbo.NVL( q.a6_82,0) as a6_82,"
    '    sql += "dbo.NVL( q.a6_83,0) as a6_83,"
    '    sql += "dbo.NVL( q.a6_84,0) as a6_84,"
    '    sql += "dbo.NVL( q.a6_85,0) as a6_85,"
    '    sql += "dbo.NVL( q.a71,0) as a71,"
    '    sql += "dbo.NVL( q.a72,0) as a72,"
    '    sql += "dbo.NVL( q.a73,0) as a73,"
    '    sql += "dbo.NVL( q.a81,0) as a81,"
    '    sql += "dbo.NVL( q.a82,0) as a82,"
    '    sql += "dbo.NVL( q.a83,0) as a83,"
    '    sql += "dbo.NVL( q.a84,0) as a84,"
    '    sql += "dbo.NVL( q.a85,0) as a85 "

    '    sql += "from (Select 1 as mid, '為補充與原專長相關之技能' Name  union "
    '    sql += "Select 2 as mid, '轉換其他行職業所需技能' Name  union "
    '    sql += "Select 3 as mid, '拓展工作領域及視野' Name  union "
    '    sql += "Select 4 as mid, '其他' Name  ) k "
    '    sql += "left outer join ( SELECT c.Mid,  "
    '    sql += "sum(case a.Q1 when 1 then 1 else 0 end) as  a11,"
    '    sql += "sum(case a.Q1 when 2 then 1 else 0 end) as  a12,"
    '    sql += "sum(case a.Q1 when 3 then 1 else 0 end) as  a13,"
    '    sql += "sum(case a.Q1 when 4 then 1 else 0 end) as  a14,"
    '    sql += "sum(case a.Q2 when 1 then 1 else 0 end) as  a21,"
    '    sql += "sum(case a.Q2 when 2 then 1 else 0 end) as  a22,"
    '    sql += "sum(case a.Q2 when 3 then 1 else 0 end) as  a23,"
    '    sql += "sum(case a.Q2 when 4 then 1 else 0 end) as  a24,"
    '    sql += "sum(case a.Q2 when 5 then 1 else 0 end) as  a25,"
    '    sql += "sum(case a.Q3 when 1 then 1 else 0 end) as  a31,"
    '    sql += "sum(case a.Q3 when 2 then 1 else 0 end) as  a32,"
    '    sql += "sum(case a.Q3 when 3 then 1 else 0 end) as  a33,"
    '    sql += "sum(case a.Q3 when 4 then 1 else 0 end) as  a34,"
    '    sql += "sum(case a.Q4 when 1 then 1 else 0 end) as  a41,"
    '    sql += "sum(case a.Q4 when 2 then 1 else 0 end) as  a42,"
    '    sql += "sum(case a.Q4 when 3 then 1 else 0 end) as  a43,"
    '    sql += "sum(case a.Q4 when 4 then 1 else 0 end) as  a44,"
    '    sql += "sum(case a.Q4 when 5 then 1 else 0 end) as  a45,"
    '    sql += "sum(case a.Q5 when 1 then 1 else 0 end) as  a51,"
    '    sql += "sum(case a.Q5 when 2 then 1 else 0 end) as  a52,"
    '    sql += "sum(case a.Q5 when 3 then 1 else 0 end) as  a53,"
    '    sql += "sum(case a.Q5 when 4 then 1 else 0 end) as  a54,"
    '    sql += "sum(case a.Q5 when 5 then 1 else 0 end) as  a55,"
    '    sql += "sum(case a.Q6 when 1 then 1 else 0 end) as  a61,"
    '    sql += "sum(case a.Q6 when 2 then 1 else 0 end) as  a62,"
    '    sql += "sum(case a.Q6 when 3 then 1 else 0 end) as  a63,"
    '    sql += "sum(case a.Q6 when 4 then 1 else 0 end) as  a64,"
    '    sql += "sum(case a.Q6 when 5 then 1 else 0 end) as  a65,"
    '    sql += "sum(case a.Q6_7 when 1 then 1 else 0 end) as a6_71,"
    '    sql += "sum(case a.Q6_7 when 2 then 1 else 0 end) as a6_72,"
    '    sql += "sum(case a.Q6_7 when 3 then 1 else 0 end) as a6_73,"
    '    sql += "sum(case a.Q6_7 when 4 then 1 else 0 end) as a6_74,"
    '    sql += "sum(case a.Q6_7 when 5 then 1 else 0 end) as a6_75,"
    '    sql += "sum(case a.Q6_8 when 1 then 1 else 0 end) as a6_81,"
    '    sql += "sum(case a.Q6_8 when 2 then 1 else 0 end) as a6_82,"
    '    sql += "sum(case a.Q6_8 when 3 then 1 else 0 end) as a6_83,"
    '    sql += "sum(case a.Q6_8 when 4 then 1 else 0 end) as a6_84,"
    '    sql += "sum(case a.Q6_8 when 5 then 1 else 0 end) as a6_85,  "
    '    sql += "sum(case a.Q7 when 1 then 1 else 0 end) as  a71,"
    '    sql += "sum(case a.Q7 when 2 then 1 else 0 end) as  a72,"
    '    sql += "sum(case a.Q7 when 3 then 1 else 0 end) as  a73,"
    '    sql += "sum(case a.Q8 when '1' then 1 else 0 end) as a81,"
    '    sql += "sum(case a.Q8 when '2' then 1 else 0 end) as a82,"
    '    sql += "sum(case a.Q8 when '3' then 1 else 0 end) as a83,"
    '    sql += "sum(case a.Q8 when '4' then 1 else 0 end) as a84,"
    '    sql += "sum(case a.Q8 when '5' then 1 else 0 end) as a85 "
    '    sql += "From Stud_QuestionFin a left outer join Class_StudentsOfClass b on a.SOCID=b.socid "
    '    sql += "left outer join "
    '    sql += "( select socid, q2 as mid from Stud_TrainBGQ2 	) c on a.SOCID=c.socid "
    '    sql += "INNER JOIN (select rid, PlanID, Years, OCID, SEQNO  from Class_ClassInfo where 1=1 "
    '    If RID <> "" Then
    '        sql += "AND RID LIKE '" + RID + "%'"
    '    End If
    '    If OCID <> "" Then
    '        sql += "AND OCID='" + OCID + "'"
    '    End If
    '    sql += ") cc ON b.OCID=cc.OCID JOIN (select PlanID, Years from ID_Plan where TPlanID='" + TPlanID + "' "
    '    If Years <> "" Then
    '        sql += "AND Years='" + Years + "'"
    '    End If
    '    sql += ") ip ON cc.PlanID = ip.PlanID AND cc.Years = dbo.SUBSTR(ip.Years,3,2) "
    '    sql += "JOIN Plan_Planinfo pp ON pp.Planid =cc.Planid  AND pp.RID = cc.RID AND pp.SEQNO = cc.SEQNO "
    '    If PackageType <> "" Then
    '        sql += "AND pp.PackageType ='" + PackageType + "'"
    '    End If
    '    sql += "JOIN (select * from org_orginfo  where 1=1 "
    '    If SearchPlan <> "" Then
    '        sql += "and OrgKind2 ='" + SearchPlan + "' "
    '    End If
    '    sql += ") oo ON pp.COMIDNO =oo.COMIDNO group by c.Mid ) q on q.mid=k.mid "
    '    sql += "order by nlssort(NAME,'NLS_SORT=TCHINESE_STROKE_M') "

    '    resDt = DbAccess.GetDataTable(sql, objconn)
    '    'sourceDT = resDt
    '    'If sourceDT.Rows.Count > 0 Then
    '    '    PName = sourceDT.Rows(0)("PName")
    '    '    plankind = sourceDT.Rows(0)("plankind")
    '    'End If

    '    setItemArray() '題目內容

    '    tmpDT = RotationDT(resDt)
    '    Return tmpDT

    'End Function

    'Private Function db_R20(ByVal TPlanID As String, ByVal Years As String, ByVal OCID As String, ByVal RID As String, ByVal SearchPlan As String, ByVal PackageType As String) As DataTable
    '    Dim tmpDT As DataTable
    '    Dim resDt As DataTable = Nothing
    '    Dim sql As String = ""

    '    sql = ""
    '    sql = "select convert(varchar, CONVERT(numeric,  '" + Years + "') -1911 ) + '年度　訓後動態調查統計表 訓後動向' AS PName, k.Name, "
    '    sql += "( Select a.PlanName +dbo.NVL(case when a.TPlanID='28' then (select '('+name+')' from v_orgkind1 WHERE VALUE='" & SearchPlan & "') end,'') from key_plan a where a.TPlanID='" & TPlanID & "') as plankind, "
    '    sql += "dbo.NVL( q.a11,0) as a11,"
    '    sql += "dbo.NVL( q.a12,0) as a12,"
    '    sql += "dbo.NVL( q.a13,0) as a13,"
    '    sql += "dbo.NVL( q.a14,0) as a14,"
    '    sql += "dbo.NVL( q.a21,0) as a21,"
    '    sql += "dbo.NVL( q.a22,0) as a22,"
    '    sql += "dbo.NVL( q.a23,0) as a23,"
    '    sql += "dbo.NVL( q.a24,0) as a24,"
    '    sql += "dbo.NVL( q.a25,0) as a25,"
    '    sql += "dbo.NVL( q.a31,0) as a31,"
    '    sql += "dbo.NVL( q.a32,0) as a32,"
    '    sql += "dbo.NVL( q.a33,0) as a33,"
    '    sql += "dbo.NVL( q.a34,0) as a34,"
    '    sql += "dbo.NVL( q.a41,0) as a41,"
    '    sql += "dbo.NVL( q.a42,0) as a42,"
    '    sql += "dbo.NVL( q.a43,0) as a43,"
    '    sql += "dbo.NVL( q.a44,0) as a44,"
    '    sql += "dbo.NVL( q.a45,0) as a45,"
    '    sql += "dbo.NVL( q.a51,0) as a51,"
    '    sql += "dbo.NVL( q.a52,0) as a52,"
    '    sql += "dbo.NVL( q.a53,0) as a53,"
    '    sql += "dbo.NVL( q.a54,0) as a54,"
    '    sql += "dbo.NVL( q.a55,0) as a55,"
    '    sql += "dbo.NVL( q.a61,0) as a61,"
    '    sql += "dbo.NVL( q.a62,0) as a62,"
    '    sql += "dbo.NVL( q.a63,0) as a63,"
    '    sql += "dbo.NVL( q.a64,0) as a64,"
    '    sql += "dbo.NVL( q.a65,0) as a65,"
    '    sql += "dbo.NVL( q.a6_71,0) as a6_71,"
    '    sql += "dbo.NVL( q.a6_72,0) as a6_72,"
    '    sql += "dbo.NVL( q.a6_73,0) as a6_73,"
    '    sql += "dbo.NVL( q.a6_74,0) as a6_74,"
    '    sql += "dbo.NVL( q.a6_75,0) as a6_75,"
    '    sql += "dbo.NVL( q.a6_81,0) as a6_81,"
    '    sql += "dbo.NVL( q.a6_82,0) as a6_82,"
    '    sql += "dbo.NVL( q.a6_83,0) as a6_83,"
    '    sql += "dbo.NVL( q.a6_84,0) as a6_84,"
    '    sql += "dbo.NVL( q.a6_85,0) as a6_85,"
    '    sql += "dbo.NVL( q.a71,0) as a71,"
    '    sql += "dbo.NVL( q.a72,0) as a72,"
    '    sql += "dbo.NVL( q.a73,0) as a73,"
    '    sql += "dbo.NVL( q.a81,0) as a81,"
    '    sql += "dbo.NVL( q.a82,0) as a82,"
    '    sql += "dbo.NVL( q.a83,0) as a83,"
    '    sql += "dbo.NVL( q.a84,0) as a84,"
    '    sql += "dbo.NVL( q.a85,0) as a85 "

    '    sql += "from (Select 1 as mid, '轉換工作' Name  union "
    '    sql += "Select 2 as mid, '留任' Name  union "
    '    sql += "Select 3 as mid, '其他' Name  ) k "
    '    sql += "left outer join ( SELECT c.Mid,  "
    '    sql += "sum(case a.Q1 when 1 then 1 else 0 end) as  a11,"
    '    sql += "sum(case a.Q1 when 2 then 1 else 0 end) as  a12,"
    '    sql += "sum(case a.Q1 when 3 then 1 else 0 end) as  a13,"
    '    sql += "sum(case a.Q1 when 4 then 1 else 0 end) as  a14,"
    '    sql += "sum(case a.Q2 when 1 then 1 else 0 end) as  a21,"
    '    sql += "sum(case a.Q2 when 2 then 1 else 0 end) as  a22,"
    '    sql += "sum(case a.Q2 when 3 then 1 else 0 end) as  a23,"
    '    sql += "sum(case a.Q2 when 4 then 1 else 0 end) as  a24,"
    '    sql += "sum(case a.Q2 when 5 then 1 else 0 end) as  a25,"
    '    sql += "sum(case a.Q3 when 1 then 1 else 0 end) as  a31,"
    '    sql += "sum(case a.Q3 when 2 then 1 else 0 end) as  a32,"
    '    sql += "sum(case a.Q3 when 3 then 1 else 0 end) as  a33,"
    '    sql += "sum(case a.Q3 when 4 then 1 else 0 end) as  a34,"
    '    sql += "sum(case a.Q4 when 1 then 1 else 0 end) as  a41,"
    '    sql += "sum(case a.Q4 when 2 then 1 else 0 end) as  a42,"
    '    sql += "sum(case a.Q4 when 3 then 1 else 0 end) as  a43,"
    '    sql += "sum(case a.Q4 when 4 then 1 else 0 end) as  a44,"
    '    sql += "sum(case a.Q4 when 5 then 1 else 0 end) as  a45,"
    '    sql += "sum(case a.Q5 when 1 then 1 else 0 end) as  a51,"
    '    sql += "sum(case a.Q5 when 2 then 1 else 0 end) as  a52,"
    '    sql += "sum(case a.Q5 when 3 then 1 else 0 end) as  a53,"
    '    sql += "sum(case a.Q5 when 4 then 1 else 0 end) as  a54,"
    '    sql += "sum(case a.Q5 when 5 then 1 else 0 end) as  a55,"
    '    sql += "sum(case a.Q6 when 1 then 1 else 0 end) as  a61,"
    '    sql += "sum(case a.Q6 when 2 then 1 else 0 end) as  a62,"
    '    sql += "sum(case a.Q6 when 3 then 1 else 0 end) as  a63,"
    '    sql += "sum(case a.Q6 when 4 then 1 else 0 end) as  a64,"
    '    sql += "sum(case a.Q6 when 5 then 1 else 0 end) as  a65,"
    '    sql += "sum(case a.Q6_7 when 1 then 1 else 0 end) as a6_71,"
    '    sql += "sum(case a.Q6_7 when 2 then 1 else 0 end) as a6_72,"
    '    sql += "sum(case a.Q6_7 when 3 then 1 else 0 end) as a6_73,"
    '    sql += "sum(case a.Q6_7 when 4 then 1 else 0 end) as a6_74,"
    '    sql += "sum(case a.Q6_7 when 5 then 1 else 0 end) as a6_75,"
    '    sql += "sum(case a.Q6_8 when 1 then 1 else 0 end) as a6_81,"
    '    sql += "sum(case a.Q6_8 when 2 then 1 else 0 end) as a6_82,"
    '    sql += "sum(case a.Q6_8 when 3 then 1 else 0 end) as a6_83,"
    '    sql += "sum(case a.Q6_8 when 4 then 1 else 0 end) as a6_84,"
    '    sql += "sum(case a.Q6_8 when 5 then 1 else 0 end) as a6_85,  "
    '    sql += "sum(case a.Q7 when 1 then 1 else 0 end) as  a71,"
    '    sql += "sum(case a.Q7 when 2 then 1 else 0 end) as  a72,"
    '    sql += "sum(case a.Q7 when 3 then 1 else 0 end) as  a73,"
    '    sql += "sum(case a.Q8 when '1' then 1 else 0 end) as a81,"
    '    sql += "sum(case a.Q8 when '2' then 1 else 0 end) as a82,"
    '    sql += "sum(case a.Q8 when '3' then 1 else 0 end) as a83,"
    '    sql += "sum(case a.Q8 when '4' then 1 else 0 end) as a84,"
    '    sql += "sum(case a.Q8 when '5' then 1 else 0 end) as a85 "
    '    sql += "From Stud_QuestionFin a left outer join Class_StudentsOfClass b on a.SOCID=b.socid "
    '    sql += "left outer join "
    '    sql += "( select socid, q3 as mid from Stud_TrainBG ) c on a.SOCID=c.socid "
    '    sql += "INNER JOIN (select rid, PlanID, Years, OCID, SEQNO  from Class_ClassInfo where 1=1 "
    '    If RID <> "" Then
    '        sql += "AND RID LIKE '" + RID + "%'"
    '    End If
    '    If OCID <> "" Then
    '        sql += "AND OCID='" + OCID + "'"
    '    End If
    '    sql += ") cc ON b.OCID=cc.OCID JOIN (select PlanID, Years from ID_Plan where TPlanID='" + TPlanID + "' "
    '    If Years <> "" Then
    '        sql += "AND Years='" + Years + "'"
    '    End If
    '    sql += ") ip ON cc.PlanID = ip.PlanID AND cc.Years = dbo.SUBSTR(ip.Years,3,2) "
    '    sql += "JOIN Plan_Planinfo pp ON pp.Planid =cc.Planid  AND pp.RID = cc.RID AND pp.SEQNO = cc.SEQNO "
    '    If PackageType <> "" Then
    '        sql += "AND pp.PackageType ='" + PackageType + "'"
    '    End If
    '    sql += "JOIN (select * from org_orginfo  where 1=1 "
    '    If SearchPlan <> "" Then
    '        sql += "and OrgKind2 ='" + SearchPlan + "' "
    '    End If
    '    sql += ") oo ON pp.COMIDNO =oo.COMIDNO group by c.Mid ) q on q.mid=k.mid "
    '    sql += "order by nlssort(NAME,'NLS_SORT=TCHINESE_STROKE_M') "

    '    resDt = DbAccess.GetDataTable(sql, objconn)
    '    'sourceDT = resDt
    '    'If sourceDT.Rows.Count > 0 Then
    '    '    PName = sourceDT.Rows(0)("PName")
    '    '    plankind = sourceDT.Rows(0)("plankind")
    '    'End If

    '    setItemArray() '題目內容

    '    tmpDT = RotationDT(resDt)
    '    Return tmpDT

    'End Function

    'Private Function db_R21(ByVal TPlanID As String, ByVal Years As String, ByVal OCID As String, ByVal RID As String, ByVal SearchPlan As String, ByVal PackageType As String) As DataTable
    '    Dim tmpDT As DataTable
    '    Dim resDt As DataTable = Nothing
    '    Dim sql As String = ""

    '    sql = ""
    '    sql = "select convert(varchar, CONVERT(numeric,  '" + Years + "') -1911 ) + '年度　訓後動態調查統計表 參訓單位類別' AS PName, k.Name, "
    '    sql += "( Select a.PlanName +dbo.NVL(case when a.TPlanID='28' then (select '('+name+')' from v_orgkind1 WHERE VALUE='" & SearchPlan & "') end,'') from key_plan a where a.TPlanID='" & TPlanID & "') as plankind, "
    '    sql += "dbo.NVL( q.a11,0) as a11,"
    '    sql += "dbo.NVL( q.a12,0) as a12,"
    '    sql += "dbo.NVL( q.a13,0) as a13,"
    '    sql += "dbo.NVL( q.a14,0) as a14,"
    '    sql += "dbo.NVL( q.a21,0) as a21,"
    '    sql += "dbo.NVL( q.a22,0) as a22,"
    '    sql += "dbo.NVL( q.a23,0) as a23,"
    '    sql += "dbo.NVL( q.a24,0) as a24,"
    '    sql += "dbo.NVL( q.a25,0) as a25,"
    '    sql += "dbo.NVL( q.a31,0) as a31,"
    '    sql += "dbo.NVL( q.a32,0) as a32,"
    '    sql += "dbo.NVL( q.a33,0) as a33,"
    '    sql += "dbo.NVL( q.a34,0) as a34,"
    '    sql += "dbo.NVL( q.a41,0) as a41,"
    '    sql += "dbo.NVL( q.a42,0) as a42,"
    '    sql += "dbo.NVL( q.a43,0) as a43,"
    '    sql += "dbo.NVL( q.a44,0) as a44,"
    '    sql += "dbo.NVL( q.a45,0) as a45,"
    '    sql += "dbo.NVL( q.a51,0) as a51,"
    '    sql += "dbo.NVL( q.a52,0) as a52,"
    '    sql += "dbo.NVL( q.a53,0) as a53,"
    '    sql += "dbo.NVL( q.a54,0) as a54,"
    '    sql += "dbo.NVL( q.a55,0) as a55,"
    '    sql += "dbo.NVL( q.a61,0) as a61,"
    '    sql += "dbo.NVL( q.a62,0) as a62,"
    '    sql += "dbo.NVL( q.a63,0) as a63,"
    '    sql += "dbo.NVL( q.a64,0) as a64,"
    '    sql += "dbo.NVL( q.a65,0) as a65,"
    '    sql += "dbo.NVL( q.a6_71,0) as a6_71,"
    '    sql += "dbo.NVL( q.a6_72,0) as a6_72,"
    '    sql += "dbo.NVL( q.a6_73,0) as a6_73,"
    '    sql += "dbo.NVL( q.a6_74,0) as a6_74,"
    '    sql += "dbo.NVL( q.a6_75,0) as a6_75,"
    '    sql += "dbo.NVL( q.a6_81,0) as a6_81,"
    '    sql += "dbo.NVL( q.a6_82,0) as a6_82,"
    '    sql += "dbo.NVL( q.a6_83,0) as a6_83,"
    '    sql += "dbo.NVL( q.a6_84,0) as a6_84,"
    '    sql += "dbo.NVL( q.a6_85,0) as a6_85,"
    '    sql += "dbo.NVL( q.a71,0) as a71,"
    '    sql += "dbo.NVL( q.a72,0) as a72,"
    '    sql += "dbo.NVL( q.a73,0) as a73,"
    '    sql += "dbo.NVL( q.a81,0) as a81,"
    '    sql += "dbo.NVL( q.a82,0) as a82,"
    '    sql += "dbo.NVL( q.a83,0) as a83,"
    '    sql += "dbo.NVL( q.a84,0) as a84,"
    '    sql += "dbo.NVL( q.a85,0) as a85 "

    '    sql += "from (Select OrgTypeID as mid, Name from Key_OrgType ) k "
    '    sql += "left outer join ( SELECT oo.Mid,  "
    '    sql += "sum(case a.Q1 when 1 then 1 else 0 end) as  a11,"
    '    sql += "sum(case a.Q1 when 2 then 1 else 0 end) as  a12,"
    '    sql += "sum(case a.Q1 when 3 then 1 else 0 end) as  a13,"
    '    sql += "sum(case a.Q1 when 4 then 1 else 0 end) as  a14,"
    '    sql += "sum(case a.Q2 when 1 then 1 else 0 end) as  a21,"
    '    sql += "sum(case a.Q2 when 2 then 1 else 0 end) as  a22,"
    '    sql += "sum(case a.Q2 when 3 then 1 else 0 end) as  a23,"
    '    sql += "sum(case a.Q2 when 4 then 1 else 0 end) as  a24,"
    '    sql += "sum(case a.Q2 when 5 then 1 else 0 end) as  a25,"
    '    sql += "sum(case a.Q3 when 1 then 1 else 0 end) as  a31,"
    '    sql += "sum(case a.Q3 when 2 then 1 else 0 end) as  a32,"
    '    sql += "sum(case a.Q3 when 3 then 1 else 0 end) as  a33,"
    '    sql += "sum(case a.Q3 when 4 then 1 else 0 end) as  a34,"
    '    sql += "sum(case a.Q4 when 1 then 1 else 0 end) as  a41,"
    '    sql += "sum(case a.Q4 when 2 then 1 else 0 end) as  a42,"
    '    sql += "sum(case a.Q4 when 3 then 1 else 0 end) as  a43,"
    '    sql += "sum(case a.Q4 when 4 then 1 else 0 end) as  a44,"
    '    sql += "sum(case a.Q4 when 5 then 1 else 0 end) as  a45,"
    '    sql += "sum(case a.Q5 when 1 then 1 else 0 end) as  a51,"
    '    sql += "sum(case a.Q5 when 2 then 1 else 0 end) as  a52,"
    '    sql += "sum(case a.Q5 when 3 then 1 else 0 end) as  a53,"
    '    sql += "sum(case a.Q5 when 4 then 1 else 0 end) as  a54,"
    '    sql += "sum(case a.Q5 when 5 then 1 else 0 end) as  a55,"
    '    sql += "sum(case a.Q6 when 1 then 1 else 0 end) as  a61,"
    '    sql += "sum(case a.Q6 when 2 then 1 else 0 end) as  a62,"
    '    sql += "sum(case a.Q6 when 3 then 1 else 0 end) as  a63,"
    '    sql += "sum(case a.Q6 when 4 then 1 else 0 end) as  a64,"
    '    sql += "sum(case a.Q6 when 5 then 1 else 0 end) as  a65,"
    '    sql += "sum(case a.Q6_7 when 1 then 1 else 0 end) as a6_71,"
    '    sql += "sum(case a.Q6_7 when 2 then 1 else 0 end) as a6_72,"
    '    sql += "sum(case a.Q6_7 when 3 then 1 else 0 end) as a6_73,"
    '    sql += "sum(case a.Q6_7 when 4 then 1 else 0 end) as a6_74,"
    '    sql += "sum(case a.Q6_7 when 5 then 1 else 0 end) as a6_75,"
    '    sql += "sum(case a.Q6_8 when 1 then 1 else 0 end) as a6_81,"
    '    sql += "sum(case a.Q6_8 when 2 then 1 else 0 end) as a6_82,"
    '    sql += "sum(case a.Q6_8 when 3 then 1 else 0 end) as a6_83,"
    '    sql += "sum(case a.Q6_8 when 4 then 1 else 0 end) as a6_84,"
    '    sql += "sum(case a.Q6_8 when 5 then 1 else 0 end) as a6_85,  "
    '    sql += "sum(case a.Q7 when 1 then 1 else 0 end) as  a71,"
    '    sql += "sum(case a.Q7 when 2 then 1 else 0 end) as  a72,"
    '    sql += "sum(case a.Q7 when 3 then 1 else 0 end) as  a73,"
    '    sql += "sum(case a.Q8 when '1' then 1 else 0 end) as a81,"
    '    sql += "sum(case a.Q8 when '2' then 1 else 0 end) as a82,"
    '    sql += "sum(case a.Q8 when '3' then 1 else 0 end) as a83,"
    '    sql += "sum(case a.Q8 when '4' then 1 else 0 end) as a84,"
    '    sql += "sum(case a.Q8 when '5' then 1 else 0 end) as a85 "
    '    sql += "From Stud_QuestionFin a left outer join Class_StudentsOfClass b on a.SOCID=b.socid "
    '    sql += "INNER JOIN (select rid, PlanID, Years, OCID, SEQNO  from Class_ClassInfo where 1=1 "
    '    If RID <> "" Then
    '        sql += "AND RID LIKE '" + RID + "%'"
    '    End If
    '    If OCID <> "" Then
    '        sql += "AND OCID='" + OCID + "'"
    '    End If
    '    sql += ") cc ON b.OCID=cc.OCID JOIN (select PlanID, Years from ID_Plan where TPlanID='" + TPlanID + "' "
    '    If Years <> "" Then
    '        sql += "AND Years='" + Years + "'"
    '    End If
    '    sql += ") ip ON cc.PlanID = ip.PlanID AND cc.Years = dbo.SUBSTR(ip.Years,3,2) "
    '    sql += "JOIN Plan_Planinfo pp ON pp.Planid =cc.Planid  AND pp.RID = cc.RID AND pp.SEQNO = cc.SEQNO "
    '    If PackageType <> "" Then
    '        sql += "AND pp.PackageType ='" + PackageType + "'"
    '    End If
    '    sql += "JOIN (select OrgKind As mid,comidno from org_orginfo  where 1=1 "
    '    If SearchPlan <> "" Then
    '        sql += "and OrgKind2 ='" + SearchPlan + "' "
    '    End If
    '    sql += ") oo ON pp.COMIDNO =oo.COMIDNO group by oo.Mid ) q on q.mid=k.mid "
    '    sql += "order by nlssort(NAME,'NLS_SORT=TCHINESE_STROKE_M') "

    '    resDt = DbAccess.GetDataTable(sql, objconn)
    '    'sourceDT = resDt
    '    'If sourceDT.Rows.Count > 0 Then
    '    '    PName = sourceDT.Rows(0)("PName")
    '    '    plankind = sourceDT.Rows(0)("plankind")
    '    'End If

    '    setItemArray() '題目內容

    '    tmpDT = RotationDT(resDt)
    '    Return tmpDT

    'End Function

    'Private Function db_R22(ByVal TPlanID As String, ByVal Years As String, ByVal OCID As String, ByVal RID As String, ByVal SearchPlan As String, ByVal PackageType As String) As DataTable
    '    Dim tmpDT As DataTable
    '    Dim resDt As DataTable = Nothing
    '    Dim sql As String = ""

    '    sql = ""
    '    sql = "select convert(varchar, CONVERT(numeric,  '" + Years + "') -1911 ) + '年度　訓後動態調查統計表 參加課程職能別' AS PName, k.CCName as Name, "
    '    sql += "( Select a.PlanName +dbo.NVL(case when a.TPlanID='28' then (select '('+name+')' from v_orgkind1 WHERE VALUE='" & SearchPlan & "') end,'') from key_plan a where a.TPlanID='" & TPlanID & "') as plankind, "
    '    sql += "dbo.NVL( q.a11,0) as a11,"
    '    sql += "dbo.NVL( q.a12,0) as a12,"
    '    sql += "dbo.NVL( q.a13,0) as a13,"
    '    sql += "dbo.NVL( q.a14,0) as a14,"
    '    sql += "dbo.NVL( q.a21,0) as a21,"
    '    sql += "dbo.NVL( q.a22,0) as a22,"
    '    sql += "dbo.NVL( q.a23,0) as a23,"
    '    sql += "dbo.NVL( q.a24,0) as a24,"
    '    sql += "dbo.NVL( q.a25,0) as a25,"
    '    sql += "dbo.NVL( q.a31,0) as a31,"
    '    sql += "dbo.NVL( q.a32,0) as a32,"
    '    sql += "dbo.NVL( q.a33,0) as a33,"
    '    sql += "dbo.NVL( q.a34,0) as a34,"
    '    sql += "dbo.NVL( q.a41,0) as a41,"
    '    sql += "dbo.NVL( q.a42,0) as a42,"
    '    sql += "dbo.NVL( q.a43,0) as a43,"
    '    sql += "dbo.NVL( q.a44,0) as a44,"
    '    sql += "dbo.NVL( q.a45,0) as a45,"
    '    sql += "dbo.NVL( q.a51,0) as a51,"
    '    sql += "dbo.NVL( q.a52,0) as a52,"
    '    sql += "dbo.NVL( q.a53,0) as a53,"
    '    sql += "dbo.NVL( q.a54,0) as a54,"
    '    sql += "dbo.NVL( q.a55,0) as a55,"
    '    sql += "dbo.NVL( q.a61,0) as a61,"
    '    sql += "dbo.NVL( q.a62,0) as a62,"
    '    sql += "dbo.NVL( q.a63,0) as a63,"
    '    sql += "dbo.NVL( q.a64,0) as a64,"
    '    sql += "dbo.NVL( q.a65,0) as a65,"
    '    sql += "dbo.NVL( q.a6_71,0) as a6_71,"
    '    sql += "dbo.NVL( q.a6_72,0) as a6_72,"
    '    sql += "dbo.NVL( q.a6_73,0) as a6_73,"
    '    sql += "dbo.NVL( q.a6_74,0) as a6_74,"
    '    sql += "dbo.NVL( q.a6_75,0) as a6_75,"
    '    sql += "dbo.NVL( q.a6_81,0) as a6_81,"
    '    sql += "dbo.NVL( q.a6_82,0) as a6_82,"
    '    sql += "dbo.NVL( q.a6_83,0) as a6_83,"
    '    sql += "dbo.NVL( q.a6_84,0) as a6_84,"
    '    sql += "dbo.NVL( q.a6_85,0) as a6_85,"
    '    sql += "dbo.NVL( q.a71,0) as a71,"
    '    sql += "dbo.NVL( q.a72,0) as a72,"
    '    sql += "dbo.NVL( q.a73,0) as a73,"
    '    sql += "dbo.NVL( q.a81,0) as a81,"
    '    sql += "dbo.NVL( q.a82,0) as a82,"
    '    sql += "dbo.NVL( q.a83,0) as a83,"
    '    sql += "dbo.NVL( q.a84,0) as a84,"
    '    sql += "dbo.NVL( q.a85,0) as a85 "

    '    sql += "from (Select ccid as mid, CCName from Key_ClassCatelog )  k "
    '    sql += "left outer join ( SELECT pp.Mid,  "
    '    sql += "sum(case a.Q1 when 1 then 1 else 0 end) as  a11,"
    '    sql += "sum(case a.Q1 when 2 then 1 else 0 end) as  a12,"
    '    sql += "sum(case a.Q1 when 3 then 1 else 0 end) as  a13,"
    '    sql += "sum(case a.Q1 when 4 then 1 else 0 end) as  a14,"
    '    sql += "sum(case a.Q2 when 1 then 1 else 0 end) as  a21,"
    '    sql += "sum(case a.Q2 when 2 then 1 else 0 end) as  a22,"
    '    sql += "sum(case a.Q2 when 3 then 1 else 0 end) as  a23,"
    '    sql += "sum(case a.Q2 when 4 then 1 else 0 end) as  a24,"
    '    sql += "sum(case a.Q2 when 5 then 1 else 0 end) as  a25,"
    '    sql += "sum(case a.Q3 when 1 then 1 else 0 end) as  a31,"
    '    sql += "sum(case a.Q3 when 2 then 1 else 0 end) as  a32,"
    '    sql += "sum(case a.Q3 when 3 then 1 else 0 end) as  a33,"
    '    sql += "sum(case a.Q3 when 4 then 1 else 0 end) as  a34,"
    '    sql += "sum(case a.Q4 when 1 then 1 else 0 end) as  a41,"
    '    sql += "sum(case a.Q4 when 2 then 1 else 0 end) as  a42,"
    '    sql += "sum(case a.Q4 when 3 then 1 else 0 end) as  a43,"
    '    sql += "sum(case a.Q4 when 4 then 1 else 0 end) as  a44,"
    '    sql += "sum(case a.Q4 when 5 then 1 else 0 end) as  a45,"
    '    sql += "sum(case a.Q5 when 1 then 1 else 0 end) as  a51,"
    '    sql += "sum(case a.Q5 when 2 then 1 else 0 end) as  a52,"
    '    sql += "sum(case a.Q5 when 3 then 1 else 0 end) as  a53,"
    '    sql += "sum(case a.Q5 when 4 then 1 else 0 end) as  a54,"
    '    sql += "sum(case a.Q5 when 5 then 1 else 0 end) as  a55,"
    '    sql += "sum(case a.Q6 when 1 then 1 else 0 end) as  a61,"
    '    sql += "sum(case a.Q6 when 2 then 1 else 0 end) as  a62,"
    '    sql += "sum(case a.Q6 when 3 then 1 else 0 end) as  a63,"
    '    sql += "sum(case a.Q6 when 4 then 1 else 0 end) as  a64,"
    '    sql += "sum(case a.Q6 when 5 then 1 else 0 end) as  a65,"
    '    sql += "sum(case a.Q6_7 when 1 then 1 else 0 end) as a6_71,"
    '    sql += "sum(case a.Q6_7 when 2 then 1 else 0 end) as a6_72,"
    '    sql += "sum(case a.Q6_7 when 3 then 1 else 0 end) as a6_73,"
    '    sql += "sum(case a.Q6_7 when 4 then 1 else 0 end) as a6_74,"
    '    sql += "sum(case a.Q6_7 when 5 then 1 else 0 end) as a6_75,"
    '    sql += "sum(case a.Q6_8 when 1 then 1 else 0 end) as a6_81,"
    '    sql += "sum(case a.Q6_8 when 2 then 1 else 0 end) as a6_82,"
    '    sql += "sum(case a.Q6_8 when 3 then 1 else 0 end) as a6_83,"
    '    sql += "sum(case a.Q6_8 when 4 then 1 else 0 end) as a6_84,"
    '    sql += "sum(case a.Q6_8 when 5 then 1 else 0 end) as a6_85,  "
    '    sql += "sum(case a.Q7 when 1 then 1 else 0 end) as  a71,"
    '    sql += "sum(case a.Q7 when 2 then 1 else 0 end) as  a72,"
    '    sql += "sum(case a.Q7 when 3 then 1 else 0 end) as  a73,"
    '    sql += "sum(case a.Q8 when '1' then 1 else 0 end) as a81,"
    '    sql += "sum(case a.Q8 when '2' then 1 else 0 end) as a82,"
    '    sql += "sum(case a.Q8 when '3' then 1 else 0 end) as a83,"
    '    sql += "sum(case a.Q8 when '4' then 1 else 0 end) as a84,"
    '    sql += "sum(case a.Q8 when '5' then 1 else 0 end) as a85 "
    '    sql += "From Stud_QuestionFin a left outer join Class_StudentsOfClass b on a.SOCID=b.socid "
    '    sql += "INNER JOIN (select rid, PlanID, Years, OCID, SEQNO  from Class_ClassInfo where 1=1 "
    '    If RID <> "" Then
    '        sql += "AND RID LIKE '" + RID + "%'"
    '    End If
    '    If OCID <> "" Then
    '        sql += "AND OCID='" + OCID + "'"
    '    End If
    '    sql += ") cc ON b.OCID=cc.OCID JOIN (select PlanID, Years from ID_Plan where TPlanID='" + TPlanID + "' "
    '    If Years <> "" Then
    '        sql += "AND Years='" + Years + "'"
    '    End If
    '    sql += ") ip ON cc.PlanID = ip.PlanID AND cc.Years = dbo.SUBSTR(ip.Years,3,2) "
    '    sql += "JOIN (select planid, comidno, seqno, rid, PackageType, ClassCate as Mid from Plan_Planinfo) pp ON pp.Planid =cc.Planid  AND pp.RID = cc.RID AND pp.SEQNO = cc.SEQNO "
    '    If PackageType <> "" Then
    '        sql += "AND pp.PackageType ='" + PackageType + "'"
    '    End If
    '    sql += "JOIN (select OrgKind As mid,comidno from org_orginfo  where 1=1 "
    '    If SearchPlan <> "" Then
    '        sql += "and OrgKind2 ='" + SearchPlan + "' "
    '    End If
    '    sql += ") oo ON pp.COMIDNO =oo.COMIDNO group by pp.Mid ) q on q.mid=k.mid "
    '    sql += "order by nlssort(NAME,'NLS_SORT=TCHINESE_STROKE_M') "

    '    resDt = DbAccess.GetDataTable(sql, objconn)
    '    'sourceDT = resDt
    '    'If sourceDT.Rows.Count > 0 Then
    '    '    PName = sourceDT.Rows(0)("PName")
    '    '    plankind = sourceDT.Rows(0)("plankind")
    '    'End If

    '    setItemArray() '題目內容

    '    tmpDT = RotationDT(resDt)
    '    Return tmpDT

    'End Function

    'Private Function db_R23(ByVal TPlanID As String, ByVal Years As String, ByVal OCID As String, ByVal RID As String, ByVal SearchPlan As String, ByVal PackageType As String) As DataTable
    '    Dim tmpDT As DataTable
    '    Dim resDt As DataTable = Nothing
    '    Dim sql As String = ""

    '    sql = ""
    '    sql = "select convert(varchar, CONVERT(numeric,  '" + Years + "') -1911 ) + '年度　訓後動態調查統計表 參加課程型態' AS PName, k.Name, "
    '    sql += "( Select a.PlanName +dbo.NVL(case when a.TPlanID='28' then (select '('+name+')' from v_orgkind1 WHERE VALUE='" & SearchPlan & "') end,'') from key_plan a where a.TPlanID='" & TPlanID & "') as plankind, "
    '    sql += "dbo.NVL( q.a11,0) as a11,"
    '    sql += "dbo.NVL( q.a12,0) as a12,"
    '    sql += "dbo.NVL( q.a13,0) as a13,"
    '    sql += "dbo.NVL( q.a14,0) as a14,"
    '    sql += "dbo.NVL( q.a21,0) as a21,"
    '    sql += "dbo.NVL( q.a22,0) as a22,"
    '    sql += "dbo.NVL( q.a23,0) as a23,"
    '    sql += "dbo.NVL( q.a24,0) as a24,"
    '    sql += "dbo.NVL( q.a25,0) as a25,"
    '    sql += "dbo.NVL( q.a31,0) as a31,"
    '    sql += "dbo.NVL( q.a32,0) as a32,"
    '    sql += "dbo.NVL( q.a33,0) as a33,"
    '    sql += "dbo.NVL( q.a34,0) as a34,"
    '    sql += "dbo.NVL( q.a41,0) as a41,"
    '    sql += "dbo.NVL( q.a42,0) as a42,"
    '    sql += "dbo.NVL( q.a43,0) as a43,"
    '    sql += "dbo.NVL( q.a44,0) as a44,"
    '    sql += "dbo.NVL( q.a45,0) as a45,"
    '    sql += "dbo.NVL( q.a51,0) as a51,"
    '    sql += "dbo.NVL( q.a52,0) as a52,"
    '    sql += "dbo.NVL( q.a53,0) as a53,"
    '    sql += "dbo.NVL( q.a54,0) as a54,"
    '    sql += "dbo.NVL( q.a55,0) as a55,"
    '    sql += "dbo.NVL( q.a61,0) as a61,"
    '    sql += "dbo.NVL( q.a62,0) as a62,"
    '    sql += "dbo.NVL( q.a63,0) as a63,"
    '    sql += "dbo.NVL( q.a64,0) as a64,"
    '    sql += "dbo.NVL( q.a65,0) as a65,"
    '    sql += "dbo.NVL( q.a6_71,0) as a6_71,"
    '    sql += "dbo.NVL( q.a6_72,0) as a6_72,"
    '    sql += "dbo.NVL( q.a6_73,0) as a6_73,"
    '    sql += "dbo.NVL( q.a6_74,0) as a6_74,"
    '    sql += "dbo.NVL( q.a6_75,0) as a6_75,"
    '    sql += "dbo.NVL( q.a6_81,0) as a6_81,"
    '    sql += "dbo.NVL( q.a6_82,0) as a6_82,"
    '    sql += "dbo.NVL( q.a6_83,0) as a6_83,"
    '    sql += "dbo.NVL( q.a6_84,0) as a6_84,"
    '    sql += "dbo.NVL( q.a6_85,0) as a6_85,"
    '    sql += "dbo.NVL( q.a71,0) as a71,"
    '    sql += "dbo.NVL( q.a72,0) as a72,"
    '    sql += "dbo.NVL( q.a73,0) as a73,"
    '    sql += "dbo.NVL( q.a81,0) as a81,"
    '    sql += "dbo.NVL( q.a82,0) as a82,"
    '    sql += "dbo.NVL( q.a83,0) as a83,"
    '    sql += "dbo.NVL( q.a84,0) as a84,"
    '    sql += "dbo.NVL( q.a85,0) as a85 "

    '    sql += "from (Select 1 as mid, '學分班' Name  union "
    '    sql += "Select 2 as mid, '非學分班' Name  union "
    '    sql += "Select 3 as mid, '企業包班' Name  ) k "

    '    sql += "left outer join ( SELECT pp.Mid,  "
    '    sql += "sum(case a.Q1 when 1 then 1 else 0 end) as  a11,"
    '    sql += "sum(case a.Q1 when 2 then 1 else 0 end) as  a12,"
    '    sql += "sum(case a.Q1 when 3 then 1 else 0 end) as  a13,"
    '    sql += "sum(case a.Q1 when 4 then 1 else 0 end) as  a14,"
    '    sql += "sum(case a.Q2 when 1 then 1 else 0 end) as  a21,"
    '    sql += "sum(case a.Q2 when 2 then 1 else 0 end) as  a22,"
    '    sql += "sum(case a.Q2 when 3 then 1 else 0 end) as  a23,"
    '    sql += "sum(case a.Q2 when 4 then 1 else 0 end) as  a24,"
    '    sql += "sum(case a.Q2 when 5 then 1 else 0 end) as  a25,"
    '    sql += "sum(case a.Q3 when 1 then 1 else 0 end) as  a31,"
    '    sql += "sum(case a.Q3 when 2 then 1 else 0 end) as  a32,"
    '    sql += "sum(case a.Q3 when 3 then 1 else 0 end) as  a33,"
    '    sql += "sum(case a.Q3 when 4 then 1 else 0 end) as  a34,"
    '    sql += "sum(case a.Q4 when 1 then 1 else 0 end) as  a41,"
    '    sql += "sum(case a.Q4 when 2 then 1 else 0 end) as  a42,"
    '    sql += "sum(case a.Q4 when 3 then 1 else 0 end) as  a43,"
    '    sql += "sum(case a.Q4 when 4 then 1 else 0 end) as  a44,"
    '    sql += "sum(case a.Q4 when 5 then 1 else 0 end) as  a45,"
    '    sql += "sum(case a.Q5 when 1 then 1 else 0 end) as  a51,"
    '    sql += "sum(case a.Q5 when 2 then 1 else 0 end) as  a52,"
    '    sql += "sum(case a.Q5 when 3 then 1 else 0 end) as  a53,"
    '    sql += "sum(case a.Q5 when 4 then 1 else 0 end) as  a54,"
    '    sql += "sum(case a.Q5 when 5 then 1 else 0 end) as  a55,"
    '    sql += "sum(case a.Q6 when 1 then 1 else 0 end) as  a61,"
    '    sql += "sum(case a.Q6 when 2 then 1 else 0 end) as  a62,"
    '    sql += "sum(case a.Q6 when 3 then 1 else 0 end) as  a63,"
    '    sql += "sum(case a.Q6 when 4 then 1 else 0 end) as  a64,"
    '    sql += "sum(case a.Q6 when 5 then 1 else 0 end) as  a65,"
    '    sql += "sum(case a.Q6_7 when 1 then 1 else 0 end) as a6_71,"
    '    sql += "sum(case a.Q6_7 when 2 then 1 else 0 end) as a6_72,"
    '    sql += "sum(case a.Q6_7 when 3 then 1 else 0 end) as a6_73,"
    '    sql += "sum(case a.Q6_7 when 4 then 1 else 0 end) as a6_74,"
    '    sql += "sum(case a.Q6_7 when 5 then 1 else 0 end) as a6_75,"
    '    sql += "sum(case a.Q6_8 when 1 then 1 else 0 end) as a6_81,"
    '    sql += "sum(case a.Q6_8 when 2 then 1 else 0 end) as a6_82,"
    '    sql += "sum(case a.Q6_8 when 3 then 1 else 0 end) as a6_83,"
    '    sql += "sum(case a.Q6_8 when 4 then 1 else 0 end) as a6_84,"
    '    sql += "sum(case a.Q6_8 when 5 then 1 else 0 end) as a6_85,  "
    '    sql += "sum(case a.Q7 when 1 then 1 else 0 end) as  a71,"
    '    sql += "sum(case a.Q7 when 2 then 1 else 0 end) as  a72,"
    '    sql += "sum(case a.Q7 when 3 then 1 else 0 end) as  a73,"
    '    sql += "sum(case a.Q8 when '1' then 1 else 0 end) as a81,"
    '    sql += "sum(case a.Q8 when '2' then 1 else 0 end) as a82,"
    '    sql += "sum(case a.Q8 when '3' then 1 else 0 end) as a83,"
    '    sql += "sum(case a.Q8 when '4' then 1 else 0 end) as a84,"
    '    sql += "sum(case a.Q8 when '5' then 1 else 0 end) as a85 "
    '    sql += "From Stud_QuestionFin a left outer join Class_StudentsOfClass b on a.SOCID=b.socid "
    '    sql += "left outer join ( select socid, q3 as mid from Stud_TrainBG  ) c on a.SOCID=c.socid "
    '    sql += "INNER JOIN (select rid, PlanID, Years, OCID, SEQNO  from Class_ClassInfo where 1=1 "
    '    If RID <> "" Then
    '        sql += "AND RID LIKE '" + RID + "%'"
    '    End If
    '    If OCID <> "" Then
    '        sql += "AND OCID='" + OCID + "'"
    '    End If
    '    sql += ") cc ON b.OCID=cc.OCID JOIN (select PlanID, Years from ID_Plan where TPlanID='" + TPlanID + "' "
    '    If Years <> "" Then
    '        sql += "AND Years='" + Years + "'"
    '    End If
    '    sql += ") ip ON cc.PlanID = ip.PlanID AND cc.Years = dbo.SUBSTR(ip.Years,3,2) "
    '    sql += "JOIN (select planid, comidno, seqno, rid, PackageType, "
    '    sql += "case when PointYN = 'Y' then 1 "
    '    sql += "when PointYN = 'N' then 2 "
    '    sql += "when IsBusiness = 'Y' then 3 end as Mid from Plan_Planinfo) pp "
    '    sql += "ON pp.Planid =cc.Planid  AND pp.RID = cc.RID AND pp.SEQNO = cc.SEQNO "
    '    If PackageType <> "" Then
    '        sql += "AND pp.PackageType ='" + PackageType + "'"
    '    End If
    '    sql += "JOIN (select * from org_orginfo  where 1=1 "
    '    If SearchPlan <> "" Then
    '        sql += "and OrgKind2 ='" + SearchPlan + "' "
    '    End If
    '    sql += ") oo ON pp.COMIDNO =oo.COMIDNO group by pp.Mid ) q on q.mid=k.mid "
    '    sql += "order by nlssort(NAME,'NLS_SORT=TCHINESE_STROKE_M') "

    '    resDt = DbAccess.GetDataTable(sql, objconn)
    '    'sourceDT = resDt
    '    'If sourceDT.Rows.Count > 0 Then
    '    '    PName = sourceDT.Rows(0)("PName")
    '    '    plankind = sourceDT.Rows(0)("plankind")
    '    'End If

    '    setItemArray() '題目內容

    '    tmpDT = RotationDT(resDt)
    '    Return tmpDT

    'End Function

    'Private Function db_R24(ByVal TPlanID As String, ByVal Years As String, ByVal OCID As String, ByVal RID As String, ByVal SearchPlan As String, ByVal PackageType As String) As DataTable
    '    Dim tmpDT As DataTable
    '    Dim resDt As DataTable = Nothing
    '    Dim sql As String = ""

    '    sql = ""
    '    sql = "select convert(varchar, CONVERT(numeric,  '" + Years + "') -1911 ) + '年度　訓後動態調查統計表 訓練業別' AS PName, k.mid as Name, "
    '    sql += "( Select a.PlanName +dbo.NVL(case when a.TPlanID='28' then (select '('+name+')' from v_orgkind1 WHERE VALUE='" & SearchPlan & "') end,'') from key_plan a where a.TPlanID='" & TPlanID & "') as plankind, "
    '    sql += "dbo.NVL( q.a11,0) as a11,"
    '    sql += "dbo.NVL( q.a12,0) as a12,"
    '    sql += "dbo.NVL( q.a13,0) as a13,"
    '    sql += "dbo.NVL( q.a14,0) as a14,"
    '    sql += "dbo.NVL( q.a21,0) as a21,"
    '    sql += "dbo.NVL( q.a22,0) as a22,"
    '    sql += "dbo.NVL( q.a23,0) as a23,"
    '    sql += "dbo.NVL( q.a24,0) as a24,"
    '    sql += "dbo.NVL( q.a25,0) as a25,"
    '    sql += "dbo.NVL( q.a31,0) as a31,"
    '    sql += "dbo.NVL( q.a32,0) as a32,"
    '    sql += "dbo.NVL( q.a33,0) as a33,"
    '    sql += "dbo.NVL( q.a34,0) as a34,"
    '    sql += "dbo.NVL( q.a41,0) as a41,"
    '    sql += "dbo.NVL( q.a42,0) as a42,"
    '    sql += "dbo.NVL( q.a43,0) as a43,"
    '    sql += "dbo.NVL( q.a44,0) as a44,"
    '    sql += "dbo.NVL( q.a45,0) as a45,"
    '    sql += "dbo.NVL( q.a51,0) as a51,"
    '    sql += "dbo.NVL( q.a52,0) as a52,"
    '    sql += "dbo.NVL( q.a53,0) as a53,"
    '    sql += "dbo.NVL( q.a54,0) as a54,"
    '    sql += "dbo.NVL( q.a55,0) as a55,"
    '    sql += "dbo.NVL( q.a61,0) as a61,"
    '    sql += "dbo.NVL( q.a62,0) as a62,"
    '    sql += "dbo.NVL( q.a63,0) as a63,"
    '    sql += "dbo.NVL( q.a64,0) as a64,"
    '    sql += "dbo.NVL( q.a65,0) as a65,"
    '    sql += "dbo.NVL( q.a6_71,0) as a6_71,"
    '    sql += "dbo.NVL( q.a6_72,0) as a6_72,"
    '    sql += "dbo.NVL( q.a6_73,0) as a6_73,"
    '    sql += "dbo.NVL( q.a6_74,0) as a6_74,"
    '    sql += "dbo.NVL( q.a6_75,0) as a6_75,"
    '    sql += "dbo.NVL( q.a6_81,0) as a6_81,"
    '    sql += "dbo.NVL( q.a6_82,0) as a6_82,"
    '    sql += "dbo.NVL( q.a6_83,0) as a6_83,"
    '    sql += "dbo.NVL( q.a6_84,0) as a6_84,"
    '    sql += "dbo.NVL( q.a6_85,0) as a6_85,"
    '    sql += "dbo.NVL( q.a71,0) as a71,"
    '    sql += "dbo.NVL( q.a72,0) as a72,"
    '    sql += "dbo.NVL( q.a73,0) as a73,"
    '    sql += "dbo.NVL( q.a81,0) as a81,"
    '    sql += "dbo.NVL( q.a82,0) as a82,"
    '    sql += "dbo.NVL( q.a83,0) as a83,"
    '    sql += "dbo.NVL( q.a84,0) as a84,"
    '    sql += "dbo.NVL( q.a85,0) as a85 "

    '    sql += "from  (select case GovClass when 1 Then '院' + '-' + dbo.SUBSTR('0'+GCode1,-2,2) "
    '    sql += "when 2 Then '局' + '-' + dbo.SUBSTR('0'+GCode1,-2,2) end as mid  "
    '    sql += "from ID_GovClassCast where 1=1 and GCode2 is NULL and GovClass in (1,2) )  k "
    '    sql += "left outer join ( SELECT pp.Mid,  "
    '    sql += "sum(case a.Q1 when 1 then 1 else 0 end) as  a11,"
    '    sql += "sum(case a.Q1 when 2 then 1 else 0 end) as  a12,"
    '    sql += "sum(case a.Q1 when 3 then 1 else 0 end) as  a13,"
    '    sql += "sum(case a.Q1 when 4 then 1 else 0 end) as  a14,"
    '    sql += "sum(case a.Q2 when 1 then 1 else 0 end) as  a21,"
    '    sql += "sum(case a.Q2 when 2 then 1 else 0 end) as  a22,"
    '    sql += "sum(case a.Q2 when 3 then 1 else 0 end) as  a23,"
    '    sql += "sum(case a.Q2 when 4 then 1 else 0 end) as  a24,"
    '    sql += "sum(case a.Q2 when 5 then 1 else 0 end) as  a25,"
    '    sql += "sum(case a.Q3 when 1 then 1 else 0 end) as  a31,"
    '    sql += "sum(case a.Q3 when 2 then 1 else 0 end) as  a32,"
    '    sql += "sum(case a.Q3 when 3 then 1 else 0 end) as  a33,"
    '    sql += "sum(case a.Q3 when 4 then 1 else 0 end) as  a34,"
    '    sql += "sum(case a.Q4 when 1 then 1 else 0 end) as  a41,"
    '    sql += "sum(case a.Q4 when 2 then 1 else 0 end) as  a42,"
    '    sql += "sum(case a.Q4 when 3 then 1 else 0 end) as  a43,"
    '    sql += "sum(case a.Q4 when 4 then 1 else 0 end) as  a44,"
    '    sql += "sum(case a.Q4 when 5 then 1 else 0 end) as  a45,"
    '    sql += "sum(case a.Q5 when 1 then 1 else 0 end) as  a51,"
    '    sql += "sum(case a.Q5 when 2 then 1 else 0 end) as  a52,"
    '    sql += "sum(case a.Q5 when 3 then 1 else 0 end) as  a53,"
    '    sql += "sum(case a.Q5 when 4 then 1 else 0 end) as  a54,"
    '    sql += "sum(case a.Q5 when 5 then 1 else 0 end) as  a55,"
    '    sql += "sum(case a.Q6 when 1 then 1 else 0 end) as  a61,"
    '    sql += "sum(case a.Q6 when 2 then 1 else 0 end) as  a62,"
    '    sql += "sum(case a.Q6 when 3 then 1 else 0 end) as  a63,"
    '    sql += "sum(case a.Q6 when 4 then 1 else 0 end) as  a64,"
    '    sql += "sum(case a.Q6 when 5 then 1 else 0 end) as  a65,"
    '    sql += "sum(case a.Q6_7 when 1 then 1 else 0 end) as a6_71,"
    '    sql += "sum(case a.Q6_7 when 2 then 1 else 0 end) as a6_72,"
    '    sql += "sum(case a.Q6_7 when 3 then 1 else 0 end) as a6_73,"
    '    sql += "sum(case a.Q6_7 when 4 then 1 else 0 end) as a6_74,"
    '    sql += "sum(case a.Q6_7 when 5 then 1 else 0 end) as a6_75,"
    '    sql += "sum(case a.Q6_8 when 1 then 1 else 0 end) as a6_81,"
    '    sql += "sum(case a.Q6_8 when 2 then 1 else 0 end) as a6_82,"
    '    sql += "sum(case a.Q6_8 when 3 then 1 else 0 end) as a6_83,"
    '    sql += "sum(case a.Q6_8 when 4 then 1 else 0 end) as a6_84,"
    '    sql += "sum(case a.Q6_8 when 5 then 1 else 0 end) as a6_85,  "
    '    sql += "sum(case a.Q7 when 1 then 1 else 0 end) as  a71,"
    '    sql += "sum(case a.Q7 when 2 then 1 else 0 end) as  a72,"
    '    sql += "sum(case a.Q7 when 3 then 1 else 0 end) as  a73,"
    '    sql += "sum(case a.Q8 when '1' then 1 else 0 end) as a81,"
    '    sql += "sum(case a.Q8 when '2' then 1 else 0 end) as a82,"
    '    sql += "sum(case a.Q8 when '3' then 1 else 0 end) as a83,"
    '    sql += "sum(case a.Q8 when '4' then 1 else 0 end) as a84,"
    '    sql += "sum(case a.Q8 when '5' then 1 else 0 end) as a85 "
    '    sql += "From Stud_QuestionFin a left outer join Class_StudentsOfClass b on a.SOCID=b.socid "
    '    sql += "left outer join ( select socid, q3 as mid from Stud_TrainBG  )  c on a.SOCID=c.socid "
    '    sql += "INNER JOIN (select rid, PlanID, Years, OCID, SEQNO  from Class_ClassInfo where 1=1 "
    '    If RID <> "" Then
    '        sql += "AND RID LIKE '" + RID + "%'"
    '    End If
    '    If OCID <> "" Then
    '        sql += "AND OCID='" + OCID + "'"
    '    End If
    '    sql += ") cc ON b.OCID=cc.OCID JOIN (select PlanID, Years from ID_Plan where TPlanID='" + TPlanID + "' "
    '    If Years <> "" Then
    '        sql += "AND Years='" + Years + "'"
    '    End If
    '    sql += ") ip ON cc.PlanID = ip.PlanID AND cc.Years = dbo.SUBSTR(ip.Years,3,2) "
    '    sql += "JOIN (select p.planid, p.comidno, p.seqno, p.rid, p.PackageType, case ig.GovClass "
    '    sql += "when 1 Then '院' + '-' + dbo.SUBSTR('0' +ig.GCode1,-2,2) "
    '    sql += "when 2 Then '局' + '-' + dbo.SUBSTR('0' +ig.GCode1,-2,2) end  as Mid "
    '    sql += "from Plan_Planinfo p "
    '    sql += "join ID_GovClassCast ig on p.GCID = ig.GCID) pp "
    '    sql += "ON pp.Planid =cc.Planid  AND pp.RID = cc.RID AND pp.SEQNO = cc.SEQNO "
    '    If PackageType <> "" Then
    '        sql += "AND pp.PackageType ='" + PackageType + "'"
    '    End If
    '    sql += "JOIN (select * from org_orginfo  where 1=1 "
    '    If SearchPlan <> "" Then
    '        sql += "and OrgKind2 ='" + SearchPlan + "' "
    '    End If
    '    sql += ") oo ON pp.COMIDNO =oo.COMIDNO group by pp.Mid ) q on q.mid=k.mid "
    '    sql += "order by nlssort(NAME,'NLS_SORT=TCHINESE_STROKE_M') "

    '    resDt = DbAccess.GetDataTable(sql, objconn)
    '    'sourceDT = resDt
    '    'If sourceDT.Rows.Count > 0 Then
    '    '    PName = sourceDT.Rows(0)("PName")
    '    '    plankind = sourceDT.Rows(0)("plankind")
    '    'End If

    '    setItemArray() '題目內容

    '    tmpDT = RotationDT(resDt)
    '    Return tmpDT

    'End Function

#End Region

End Class