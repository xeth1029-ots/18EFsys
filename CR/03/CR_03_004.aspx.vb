Imports System.IO
Imports OfficeOpenXml
Imports OfficeOpenXml.Style

Public Class CR_03_004
    Inherits AuthBasePage 'System.Web.UI.Page

    '<報表4> 管制類課程核定比例彙總表
    '系統：在職系統    '計畫：產業人才投資方案(產投+自主)    '使用者：只有署可以使用
    '功能：首頁>>課程審查>>【陳核版】課程核定報表>>管制類課程核定比例彙總表    '需求：    '<查詢介面>    '申請階段、計畫、匯出檔案格式
    '<匯出>    '1.匯出報表請依照附件格式產出。    '2.報表要區分產投、自主    '3.報表名稱：年度 + 申請階段 + 計畫 + "管制課程比例五分署彙總表"
    '如：113年度下半年產業人才投資計畫管制課程比例五分署彙總表    '4.管制類定義：
    '管制類指的是僅篩選【課程職類】為這三類的班級：美容【05-01】、餐飲【06】、手工藝品【07】
    '細分說明： '美容【05-01】：【03-02】、【03-03】、【03-04】、【03-05】 (不含【03-01】)    '餐飲【06】：【21-01】、【21-02】、【21-03】、【30-01】、【30-02】   
    '手工藝品【07】：【02-04】、【07-01】、【07-03】    '5.匯出欄位：參考附件
    '(1)上下半年不一樣，下半年會包括上半年跟全年度    '(2)【管制類課程類別】：%數是固定，但產投、自主不一樣    '(3)產投：後面會有104、106為基底 (數字固定是死的)    
    '(4)自主：後面會有104為基底 (數字固定是死的)    '(5)各欄位數值計算(可參考Excel公式)：
    '【管制類課程占比總補助費】：管制類課程核定總補助費 / 核定總補助費，四捨五入至小數第2位 > 取百分比    '【班次】：管制類課程核定總班次 / 核定總班次，取百分比
    '其他的Excel上面有寫......    '4.所謂核定：是指課程審查>>二階審查>>核班結果  ，【核班結果】= 通過之班級。
    '5.可核配上限：就是該單位之等級對應首頁>>課程審查>>一階審查>>等級核配額度設定內所設定之可核定額度    '6.排序：分署依 北>桃>中>南>高 排序
    Dim print_lock As New Object '(); //lock
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
    End Sub

    Sub CCreate1()
        PanelSch1.Visible = True

        msg1.Text = ""

        'ddlYEARS_SCH = TIMS.GetSyear(ddlYEARS_SCH)
        'Common.SetListItem(ddlYEARS_SCH, sm.UserInfo.Years)

        ddlAPPSTAGE_SCH = TIMS.GET_APPSTAGE2_N34(ddlAPPSTAGE_SCH)
        'Dim objItem3 As Web.UI.WebControls.ListItem = ddlAPPSTAGE_SCH.Items.FindByValue("3")
        'Dim objItem4 As Web.UI.WebControls.ListItem = ddlAPPSTAGE_SCH.Items.FindByValue("4")
        'If objItem3 IsNot Nothing Then ddlAPPSTAGE_SCH.Items.Remove(objItem3)
        'If objItem4 IsNot Nothing Then ddlAPPSTAGE_SCH.Items.Remove(objItem4)
        Common.SetListItem(ddlAPPSTAGE_SCH, "1")

        '計畫  產業人才投資計畫/提升勞工自主學習計畫
        Dim vsOrgKind2 As String = TIMS.Get_OrgKind2(sm.UserInfo.OrgID, TIMS.c_ORGID, objconn)
        If (vsOrgKind2 = "") Then vsOrgKind2 = "G"
        rblOrgKind2 = TIMS.Get_RblSearchPlan(rblOrgKind2, objconn, False)
        Common.SetListItem(rblOrgKind2, vsOrgKind2)

        If TIMS.sUtl_ChkTest() Then
            Common.SetListItem(rblOrgKind2, "G")
            Common.SetListItem(ddlAPPSTAGE_SCH, "2")
        End If
    End Sub

    ''' <summary> 查詢SQL DataTable </summary>
    ''' <returns></returns>
    Public Function SEARCH_DATA1_dt(hPMS As Hashtable) As DataTable
        Dim v_YEARS As String = sm.UserInfo.Years
        Dim v_APPSTAGE As String = TIMS.GetMyValue2(hPMS, "APPSTAGE")
        Dim i_APPSTAGE As Integer = TIMS.GetValue2(v_APPSTAGE)
        Dim v_APPLIEDRESULT As String = TIMS.GetMyValue2(hPMS, "APPLIEDRESULT")
        Dim v_CURESULT As String = TIMS.GetMyValue2(hPMS, "CURESULT")
        'sm As SessionModel, hPMS As Hashtable
        '班級審核已通過:v_APPLIEDRESULT:Y
        Dim dt As DataTable = Nothing
        '訓練機構 '計畫'TRPlanPoint28
        Dim v_rblOrgKind2 As String = TIMS.GetListValue(rblOrgKind2)
        If v_rblOrgKind2 = "" Then
            msg1.Text = String.Concat(TIMS.cst_NODATAMsg2, "請選擇計畫")
            Return dt
        End If
        'Dim v_APPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH) '申請階段
        'If v_APPSTAGE_SCH = "" Then
        '    msg1.Text = String.Concat(TIMS.cst_NODATAMsg2, "請選擇申請階段")
        '    Return dt
        'End If
        'sSql &= " DECLARE @TPLANID VARCHAR(4) ='28' DECLARE @YEARS VARCHAR(4) ='2024' DECLARE @APPSTAGE INT =1; DECLARE @ORGKIND2 VARCHAR(4) ='W';" & vbCrLf
        'DECLARE @TPLANID NVARCHAR(2)='28';/*2*/DECLARE @YEARS SMALLINT=CONVERT(SMALLINT,'2024');/*3*/DECLARE @APPSTAGE NVARCHAR(1)='2';/*4*/DECLARE @ORGKIND2 NVARCHAR(1)='G';/*1*/
        ' CURESULT 核班結果,核班結果'Y 通過、N 不通過
        Dim parms As New Hashtable From {{"TPLANID", sm.UserInfo.TPlanID}, {"YEARS", v_YEARS}, {"ORGKIND2", v_rblOrgKind2}}
        Dim sSql As String = ""
        sSql = "
WITH WC1 AS ( SELECT pp.PSNO28,pp.PCS,pp.PLANID,pp.COMIDNO,pp.SEQNO,pp.DISTID,pp.ORGKIND2 
,PP.ISAPPRPAPER,PP.TRANSFLAG,PP.OCID,pp.ISSUCCESS,pp.TNUM ,PP.TOTALCOST,PP.DEFGOVCOST,pp.GCID3
,pf.CURESULT,IG3.GCODE33,IG3.PNAME,IG3.PNAME2,IG3.PTMID
,case when pp.GCID3 IN (2068,2069,2070,2071) then 1 end G051
,case when pp.GCID3 IN (2073,2074,2075,2076,2077) then 1 end G06
,case when pp.GCID3 IN (2078,2079,2080) then 1 end G07
FROM dbo.VIEW2B PP
JOIN dbo.V_GOVCLASSCAST3 IG3 ON IG3.GCID3=PP.GCID3
JOIN dbo.PLAN_STAFFOPIN pf on pf.PSNO28=pp.PSNO28 
WHERE (pp.RESULTBUTTON IS NULL OR pp.APPLIEDRESULT='Y') 
AND pp.PVR_ISAPPRPAPER='Y' AND pp.DATANOTSENT IS NULL" '/*'審核送出(已送審)/正式/未檢送資料註記(排除有勾選)*/
        sSql &= " AND pp.TPLANID=@TPLANID AND pp.YEARS=@YEARS AND PP.ORGKIND2=@ORGKIND2"
        If v_CURESULT = "Y" Then sSql &= " AND PF.CURESULT='Y'"
        If v_APPLIEDRESULT = "Y" Then sSql &= " AND PP.APPLIEDRESULT ='Y'"
        If (i_APPSTAGE = 1 OrElse i_APPSTAGE = 2) Then
            parms.Add("APPSTAGE", i_APPSTAGE)
            sSql &= " AND pp.APPSTAGE=@APPSTAGE"
        ElseIf (i_APPSTAGE = 3) Then
            '全年度：113上半年班級審核通過+113下半年於二階審查核班結果已通過
            sSql &= " AND ((PP.APPSTAGE=1 AND PP.APPLIEDRESULT ='Y') OR (PP.APPSTAGE=2 AND PF.CURESULT='Y'))"
        Else
            '(異常)
            sSql &= " AND 1!=1"
        End If
        sSql &= " )"
        sSql &= "
,WG1 AS ( SELECT PP.DISTID,SUM(PP.DEFGOVCOST) DEFGOVCOSTG1, COUNT(1) CLASSCNTG1 FROM WC1 PP GROUP BY PP.DISTID )
,WG2 AS ( SELECT PP.DISTID,'99' GCODE33,SUM(PP.DEFGOVCOST) DEFGOVCOST, COUNT(1) CLASSCNT FROM WC1 PP WHERE (pp.G051=1 OR pp.G06=1 OR pp.G07=1) GROUP BY PP.DISTID )
,WG3 AS ( SELECT PP.DISTID,PP.GCODE33,SUM(PP.DEFGOVCOST) DEFGOVCOST, COUNT(1) CLASSCNT FROM WC1 PP WHERE (pp.G051=1 OR pp.G06=1 OR pp.G07=1) GROUP BY PP.DISTID,PP.GCODE33 )
,WG4 AS ( SELECT '999' DISTID,PP.GCODE33,SUM(PP.DEFGOVCOST) DEFGOVCOST, COUNT(1) CLASSCNT FROM WC1 PP WHERE (pp.G051=1 OR pp.G06=1 OR pp.G07=1) GROUP BY PP.GCODE33 )
,WG234 AS (
SELECT PP.DISTID,PP.GCODE33,DEFGOVCOST,CLASSCNT FROM WG3 PP
UNION ALL SELECT PP.DISTID,PP.GCODE33,DEFGOVCOST,CLASSCNT FROM WG2 PP
UNION ALL SELECT PP.DISTID,PP.GCODE33,DEFGOVCOST,CLASSCNT FROM WG4 PP
)
,WMG1 AS (
SELECT ROWSEQNO,DISTID,GCODE33 FROM (VALUES (1,'001','07'),(2,'001','051'),(3,'001','06'),(4,'001','99') 
,(11,'003','07'),(12,'003','051'),(13,'003','06'),(14,'003','99') 
,(21,'004','07'),(22,'004','051'),(23,'004','06'),(24,'004','99') 
,(31,'005','07'),(32,'005','051'),(33,'005','06'),(34,'005','99') 
,(41,'006','07'),(42,'006','051'),(43,'006','06'),(44,'006','99')
,(51,'999','07'),(52,'999','051'),(53,'999','06'),(54,'999','99') 
) AS MG(ROWSEQNO,DISTID,GCODE33)
)
SELECT W.ROWSEQNO,W.DISTID,W.GCODE33,A.DEFGOVCOST,A.CLASSCNT, B.DEFGOVCOSTG1, B.CLASSCNTG1
,FORMAT(ROUND((1.0*A.DEFGOVCOST/B.DEFGOVCOSTG1), 4), 'P2') RATE13
,FORMAT(ROUND((1.0*A.CLASSCNT/B.CLASSCNTG1), 4), 'P2') RATE23
,FORMAT(ROUND((1.0*A.CLASSCNT/B.CLASSCNTG1), 4), 'P2') RATE24
/* ,1.0*A.DEFGOVCOST/B.DEFGOVCOSTG1 RATE11
,ROUND(1.0*A.DEFGOVCOST/B.DEFGOVCOSTG1,2) RATE12
,FORMAT(ROUND((1.0*A.DEFGOVCOST/B.DEFGOVCOSTG1), 4), 'P2') RATE13
,FORMAT(ROUND((100.0*A.DEFGOVCOST/B.DEFGOVCOSTG1), 2)/100, 'P2') RATE14
,1.0*A.CLASSCNT/B.CLASSCNTG1 RATE21
,ROUND(1.0*A.CLASSCNT/B.CLASSCNTG1,2) RATE22
,FORMAT(ROUND((1.0*A.CLASSCNT/B.CLASSCNTG1), 4), 'P2') RATE23
,FORMAT(ROUND((100.0*A.CLASSCNT/B.CLASSCNTG1), 2)/100, 'P2') RATE24  */
FROM WMG1 W
LEFT JOIN WG234 A ON A.DISTID=W.DISTID AND A.GCODE33=W.GCODE33
LEFT JOIN WG1 B ON B.DISTID=A.DISTID  
ORDER BY W.ROWSEQNO
"
        If TIMS.sUtl_ChkTest() Then
            TIMS.WriteLog(Me, String.Concat("--", vbCrLf, TIMS.GetMyValue5(parms), vbCrLf, "--##CR_03_004-ORGKIND2:", v_rblOrgKind2, ",APPSTAGE:", v_APPSTAGE, ",sSql:", vbCrLf, sSql))
        End If
        dt = DbAccess.GetDataTable(sSql, objconn, parms)

        Return dt
    End Function

    '匯出 
    Protected Sub BtnExport1_Click(sender As Object, e As EventArgs) Handles BtnExport1.Click
        Dim v_rblOrgKind2 As String = TIMS.GetListValue(rblOrgKind2) '計畫 G/W
        If v_rblOrgKind2 = "" Then
            msg1.Text = String.Concat(TIMS.cst_NODATAMsg2, "請選擇計畫") '計畫 G/W
            Return 'dt
        End If
        Dim v_APPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH) '申請階段 1/2
        If v_APPSTAGE_SCH = "" Then
            msg1.Text = String.Concat(TIMS.cst_NODATAMsg2, "請選擇申請階段")
            Return 'dt
        End If

        If v_rblOrgKind2 = "G" AndAlso v_APPSTAGE_SCH = "1" Then
            Call ExportXlsG1()
        ElseIf v_rblOrgKind2 = "G" AndAlso v_APPSTAGE_SCH = "2" Then
            Call ExportXlsG2()
        ElseIf v_rblOrgKind2 = "W" AndAlso v_APPSTAGE_SCH = "1" Then
            Call ExportXlsW1()
        ElseIf v_rblOrgKind2 = "W" AndAlso v_APPSTAGE_SCH = "2" Then
            Call ExportXlsW2()
        End If
    End Sub

    Sub ExportXlsG1()
        Const cst_SampleXLS As String = "~\CR\03\Tmp\e1140407x1G1.xlsx"
        If Not IO.File.Exists(Server.MapPath(cst_SampleXLS)) Then
            Common.MessageBox(Me, "Sample檔案不存在")
            Return '  Exit Sub
        End If

        Const Cst_FileSavePath As String = "~/CR/03/Temp/" '"~\CO\01\Temp\"
        Call TIMS.MyCreateDir(Me, Cst_FileSavePath)

        Dim hPMS As New Hashtable From {{"APPSTAGE", 1}, {"APPLIEDRESULT", ""}, {"CURESULT", "Y"}}
        Dim dtXls1 As DataTable = SEARCH_DATA1_dt(hPMS) '上半年
        'Dim dtXls2 As DataTable = SEARCH_DATA1_dt(2) '下半年
        'Dim dtXls3 As DataTable = SEARCH_DATA1_dt(3) '全年度
        Dim v_APPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH) '申請階段 1/2
        If v_APPSTAGE_SCH = "1" AndAlso TIMS.dtNODATA(dtXls1) Then
            Common.MessageBox(Me, "查無(上半年)匯出資料。")
            Exit Sub
        End If
        'If v_APPSTAGE_SCH = "2" AndAlso TIMS.dtNODATA(dtXls2) Then
        '    Common.MessageBox(Me, "查無(下半年)匯出資料。")
        '    Exit Sub
        'End If
        'If TIMS.dtNODATA(dtXls3) Then
        '    Common.MessageBox(Me, "查無(全年度)匯出資料。")
        '    Exit Sub
        'End If

        Dim cellsCOLSPNumF As String = "A{0}:R{0}"
        Dim cellsCOLSPNumF2 As String = "R{0}"
        Dim cellsCOLSPNumF3 As String() = "D6:D8,D10:D12,D14:D16,D18:D20,D22:D24,D26:D28".Split(",")
        Dim cellsCOLSPNumF4 As String() = "E6:E8,E10:E12,E14:E16,E18:E20,E22:E24,E26:E28".Split(",")
        'Dim cellsCOLSPNumF5 As String() = "N6:N8,N10:N12,N14:N16,N18:N20,N22:N24,N26:N28".Split(",")

        Dim strErrmsg As String = ""

        Dim s_ROCYEAR1 As String = CStr(CInt(sm.UserInfo.Years) - 1911) '年度
        Dim v_ddlAPPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH)
        Dim s_APPSTAGE_NM2 As String = TIMS.GET_APPSTAGE2_NM2(v_ddlAPPSTAGE_SCH) '申請階段
        Dim v_rblOrgKind2 As String = TIMS.GetListValue(rblOrgKind2)
        Dim s_PLANNAME As String = TIMS.GetListText(rblOrgKind2) '計畫
        Dim s_TITLENAME1 As String = String.Concat(s_ROCYEAR1, "年度", s_APPSTAGE_NM2, s_PLANNAME, "管制類課程比例五分署彙總表")
        Dim s_FILENAME1 As String = String.Concat(s_ROCYEAR1, s_APPSTAGE_NM2, s_PLANNAME, "管制類課程比例_", TIMS.GetDateNo())
        Dim sFileName As String = String.Concat(Cst_FileSavePath, TIMS.GetDateNo(), ".xlsx") '複製一份(Sample)
        Dim filePath1 As String = Server.MapPath(sFileName) '複製一份(Sample)
        Try
            IO.File.Copy(Server.MapPath(cst_SampleXLS), filePath1, True)
        Catch ex As Exception
            strErrmsg = ""
            strErrmsg += "目錄名稱或磁碟區標籤語法錯誤!!!" & vbCrLf
            strErrmsg += " (若連續出現多次(3次以上)請連絡系統管理者協助，謝謝)(造成您的不便深感抱歉)" & vbCrLf
            strErrmsg += ex.ToString & vbCrLf
            Common.MessageBox(Me, strErrmsg)
            'Return 'Exit Sub
        End Try
        If strErrmsg <> "" Then
            Try
                strErrmsg += "Path/File: " & filePath1 & vbCrLf
                strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
                'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                Call TIMS.WriteTraceLog(strErrmsg)
            Catch ex As Exception
            End Try
            Return 'Exit Sub
        End If

        Dim fg_RespWriteEnd As Boolean = False
        SyncLock print_lock
            'ExcelPackage.LicenseContext = LicenseContext.Commercial' ExcelPackage.LicenseContext = LicenseContext.NonCommercial
            Dim ndt As DateTime = Now
            Dim file1 As New FileInfo(filePath1)
            Dim ep As New ExcelPackage(file1)
            'Dim ws As ExcelWorksheet = excel.Workbook.Worksheets.Add(rptName)
            Dim ws As ExcelWorksheet = ep.Workbook.Worksheets(0)

            ' 共用設定
            Dim fontName As String = "標楷體"
            Dim fontSize12s As Single = 12.0F
            ' 報表標題
            Using exlRow1 As ExcelRange = ws.Cells(String.Format(cellsCOLSPNumF, "1"))
                With exlRow1
                    .Merge = True
                    .Style.Font.Name = fontName
                    .Style.Font.Size = 16
                    .Value = s_TITLENAME1
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.VerticalAlignment = ExcelVerticalAlignment.Center
                End With
            End Using

            ' 列印日期
            Using exlRow2 As ExcelRange = ws.Cells(String.Format(cellsCOLSPNumF2, "2"))
                With exlRow2
                    .Merge = True
                    .Style.Font.Name = fontName
                    .Style.Font.Size = fontSize12s
                    .Value = String.Format("{0}.{1}.{2}", ndt.Year - 1911, ndt.Month, ndt.Day)
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    .Style.VerticalAlignment = ExcelVerticalAlignment.Center
                End With
            End Using

            ws.Cells("D3:I4").Value = String.Concat(s_ROCYEAR1, "上半年課程核定情形")
            'ws.Cells("J3:K4").Value = String.Concat(s_ROCYEAR1, "上半年")
            ws.Cells("L4:M4").Value = String.Concat(s_ROCYEAR1, "上半年管制類課程")
            ws.Cells("Q4").Value = String.Concat(s_ROCYEAR1, "上半年餐飲類課程")
            'ws.Cells("V4").Value = String.Concat(s_ROCYEAR1, "年餐飲類課程")

            Dim i_DEFGOVCOST_F29 As Integer = 0
            Dim i_CLASSCNT_H29 As Integer = 0
            Dim i_DEFGOVCOSTG1 As Integer = 0
            Dim i_CLASSCNTG1 As Integer = 0
            Dim idxStr As Integer = 6
            For Each dr1 As DataRow In dtXls1.Rows

                SetCellValue(ws, "F" & idxStr, dr1("DEFGOVCOST"))
                SetCellValue(ws, "G" & idxStr, dr1("RATE13"))
                SetCellValue(ws, "H" & idxStr, dr1("CLASSCNT"))
                SetCellValue(ws, "I" & idxStr, dr1("RATE24"))

                If Convert.ToString(dr1("GCODE33")) <> "99" AndAlso Convert.ToString(dr1("DISTID")) = "999" Then
                    i_DEFGOVCOST_F29 += TIMS.VAL1(dr1("DEFGOVCOST"))
                    i_CLASSCNT_H29 += TIMS.VAL1(dr1("CLASSCNT"))
                End If

                If (Convert.ToString(dr1("GCODE33")) = "99" AndAlso Convert.ToString(dr1("DISTID")) = "001") Then
                    i_DEFGOVCOSTG1 += TIMS.VAL1(dr1("DEFGOVCOSTG1"))
                    i_CLASSCNTG1 += TIMS.VAL1(dr1("CLASSCNTG1"))
                    SetCellValue(ws, cellsCOLSPNumF3(0), dr1("DEFGOVCOSTG1"))
                    SetCellValue(ws, cellsCOLSPNumF4(0), dr1("CLASSCNTG1"))
                ElseIf (Convert.ToString(dr1("GCODE33")) = "99" AndAlso Convert.ToString(dr1("DISTID")) = "003") Then
                    i_DEFGOVCOSTG1 += TIMS.VAL1(dr1("DEFGOVCOSTG1"))
                    i_CLASSCNTG1 += TIMS.VAL1(dr1("CLASSCNTG1"))
                    SetCellValue(ws, cellsCOLSPNumF3(1), dr1("DEFGOVCOSTG1"))
                    SetCellValue(ws, cellsCOLSPNumF4(1), dr1("CLASSCNTG1"))
                ElseIf (Convert.ToString(dr1("GCODE33")) = "99" AndAlso Convert.ToString(dr1("DISTID")) = "004") Then
                    i_DEFGOVCOSTG1 += TIMS.VAL1(dr1("DEFGOVCOSTG1"))
                    i_CLASSCNTG1 += TIMS.VAL1(dr1("CLASSCNTG1"))
                    SetCellValue(ws, cellsCOLSPNumF3(2), dr1("DEFGOVCOSTG1"))
                    SetCellValue(ws, cellsCOLSPNumF4(2), dr1("CLASSCNTG1"))
                ElseIf (Convert.ToString(dr1("GCODE33")) = "99" AndAlso Convert.ToString(dr1("DISTID")) = "005") Then
                    i_DEFGOVCOSTG1 += TIMS.VAL1(dr1("DEFGOVCOSTG1"))
                    i_CLASSCNTG1 += TIMS.VAL1(dr1("CLASSCNTG1"))
                    SetCellValue(ws, cellsCOLSPNumF3(3), dr1("DEFGOVCOSTG1"))
                    SetCellValue(ws, cellsCOLSPNumF4(3), dr1("CLASSCNTG1"))
                ElseIf (Convert.ToString(dr1("GCODE33")) = "99" AndAlso Convert.ToString(dr1("DISTID")) = "006") Then
                    i_DEFGOVCOSTG1 += TIMS.VAL1(dr1("DEFGOVCOSTG1"))
                    i_CLASSCNTG1 += TIMS.VAL1(dr1("CLASSCNTG1"))
                    SetCellValue(ws, cellsCOLSPNumF3(4), dr1("DEFGOVCOSTG1"))
                    SetCellValue(ws, cellsCOLSPNumF4(4), dr1("CLASSCNTG1"))
                ElseIf (Convert.ToString(dr1("GCODE33")) = "99" AndAlso Convert.ToString(dr1("DISTID")) = "999") Then
                    SetCellValue(ws, cellsCOLSPNumF3(5), i_DEFGOVCOSTG1)
                    SetCellValue(ws, cellsCOLSPNumF4(5), i_CLASSCNTG1)
                    SetCellValue(ws, "F29", i_DEFGOVCOST_F29)
                    SetCellValue(ws, "H29", i_CLASSCNT_H29)
                End If

                If (Convert.ToString(dr1("GCODE33")) = "99") Then
                    SetCellValue(ws, "L" & idxStr, "-")
                    SetCellValue(ws, "M" & idxStr, "-")
                Else
                    Dim i_B_COST As Double = TIMS.VAL1(ws.Cells("F" & idxStr).Value)
                    Dim i_A_COST As Double = TIMS.VAL1(ws.Cells("J" & idxStr).Value)
                    If i_A_COST > 0 AndAlso i_B_COST > 0 Then
                        SetCellValue(ws, "L" & idxStr, String.Concat(TIMS.ROUND(i_B_COST / i_A_COST * 100, 0), "%"))
                    End If
                    Dim i_D_CNT As Double = TIMS.VAL1(ws.Cells("H" & idxStr).Value)
                    Dim i_C_CNT As Double = TIMS.VAL1(ws.Cells("K" & idxStr).Value)
                    If i_C_CNT > 0 AndAlso i_D_CNT > 0 Then
                        SetCellValue(ws, "M" & idxStr, String.Concat(TIMS.ROUND(i_D_CNT / i_C_CNT * 100, 0), "%"))
                    End If
                End If

                If (Convert.ToString(dr1("GCODE33")) = "06") Then
                    Dim i_Q_RATE As Object = ws.Cells("I" & idxStr).Value
                    SetCellValue(ws, "Q" & idxStr, i_Q_RATE)
                End If

                idxStr += 1
            Next

            Dim iF26 As Double = If(i_DEFGOVCOSTG1 > 0, TIMS.ROUND(TIMS.VAL1(ws.Cells("F26").Value) / i_DEFGOVCOSTG1 * 100, 2), "0")
            Dim iF27 As Double = If(i_DEFGOVCOSTG1 > 0, TIMS.ROUND(TIMS.VAL1(ws.Cells("F27").Value) / i_DEFGOVCOSTG1 * 100, 2), "0")
            Dim iF28 As Double = If(i_DEFGOVCOSTG1 > 0, TIMS.ROUND(TIMS.VAL1(ws.Cells("F28").Value) / i_DEFGOVCOSTG1 * 100, 2), "0")
            Dim iF29 As Double = TIMS.ROUND(TIMS.VAL1(iF26 + iF27 + iF28), 2)
            SetCellValue(ws, "G26", String.Concat(iF26, "%"))
            SetCellValue(ws, "G27", String.Concat(iF27, "%"))
            SetCellValue(ws, "G28", String.Concat(iF28, "%"))
            SetCellValue(ws, "G29", String.Concat(iF29, "%"))

            Dim iH26 As Double = If(i_CLASSCNTG1 > 0, TIMS.ROUND(TIMS.VAL1(ws.Cells("H26").Value) / i_CLASSCNTG1 * 100, 2), "0")
            Dim iH27 As Double = If(i_CLASSCNTG1 > 0, TIMS.ROUND(TIMS.VAL1(ws.Cells("H27").Value) / i_CLASSCNTG1 * 100, 2), "0")
            Dim iH28 As Double = If(i_CLASSCNTG1 > 0, TIMS.ROUND(TIMS.VAL1(ws.Cells("H28").Value) / i_CLASSCNTG1 * 100, 2), "0")
            Dim iH29 As Double = TIMS.ROUND(TIMS.VAL1(iH26 + iH27 + iH28), 2)
            SetCellValue(ws, "I26", String.Concat(iH26, "%"))
            SetCellValue(ws, "I27", String.Concat(iH27, "%"))
            SetCellValue(ws, "I28", String.Concat(iH28, "%")) : SetCellValue(ws, "Q28", String.Concat(iH28, "%"))
            SetCellValue(ws, "I29", String.Concat(iH29, "%"))

            '(1) 114上半年管制類課程核配總補助費之規模比率(B/A)(綠色底)：南分署的114年上半年課程手工藝類及美容美髮類的核定總補助費加起來再除以104年核定總補助費。  (3,776,920+2,253,536)/11,672,608=52%
            '(2) 114上半年管制類課程核配班次之規模比率(D/C)(橘色底)：南分署的114年上半年管制類課程手工藝類及美容美髮類核定班次加起來再除以104年管制類課程核定總班次。  (15+8)/55=42%
            Dim iF18 As Double = TIMS.VAL1(ws.Cells("F18").Value)
            Dim iF19 As Double = TIMS.VAL1(ws.Cells("F19").Value)
            Dim iJ18 As Double = TIMS.VAL1(ws.Cells("J18").Value)
            Dim iH18 As Double = TIMS.VAL1(ws.Cells("H18").Value)
            Dim iH19 As Double = TIMS.VAL1(ws.Cells("H19").Value)
            Dim iK18 As Double = TIMS.VAL1(ws.Cells("K18").Value)
            'ws.Cells("H18:H19").Merge = True
            'ws.Cells("H18:H19").Value = TIMS.VAL1(iH18 + iH19)
            ws.Cells("K18:K19").Merge = True
            Dim iL18B As Double = If(iK18 > 0, TIMS.ROUND((iF18 + iF19) / iJ18 * 100, 0), "0")
            ws.Cells("L18:L19").Merge = True
            ws.Cells("L18:L19").Value = String.Concat(iL18B, "%")
            Dim iM18C As Double = If(iH18 > 0, TIMS.ROUND((iH18 + iH19) / iK18 * 100, 0), "0")
            ws.Cells("M18:M19").Merge = True
            ws.Cells("M18:M19").Value = String.Concat(iM18C, "%")

            Dim V_ExpType As String = TIMS.GetListValue(RBListExpType)
            Select Case V_ExpType
                Case "EXCEL"
                    TIMS.ExpExcel_1(Me, strErrmsg, ep, Cst_FileSavePath, s_FILENAME1)
                    fg_RespWriteEnd = True
                Case "ODS"
                    TIMS.ExpODSl_1(Me, strErrmsg, ep, Cst_FileSavePath, s_FILENAME1)
                    fg_RespWriteEnd = True
                Case Else
                    Dim s_log1 As String = String.Format("ExpType(參數有誤)!!{0}", V_ExpType)
                    Common.MessageBox(Me, s_log1)
                    Return ' Exit Sub
            End Select

            '刪除Temp中的資料 
            Call TIMS.MyFileDelete(filePath1)
            If fg_RespWriteEnd Then TIMS.Utl_RespWriteEnd(Me, objconn, "")
        End SyncLock

        '刪除Temp中的資料 'Call TIMS.MyFileDelete(myFileName1)
        If strErrmsg <> "" Then
            Common.MessageBox(Me, strErrmsg)
            Return
        End If
    End Sub

    Sub ExportXlsW1()
        Const cst_SampleXLS As String = "~\CR\03\Tmp\e1140407x1W1.xlsx"
        If Not IO.File.Exists(Server.MapPath(cst_SampleXLS)) Then
            Common.MessageBox(Me, "Sample檔案不存在")
            Return '  Exit Sub
        End If
        Const Cst_FileSavePath As String = "~/CR/03/Temp/" '"~\CO\01\Temp\"
        Call TIMS.MyCreateDir(Me, Cst_FileSavePath)

        Dim hPMS As New Hashtable From {{"APPSTAGE", 1}, {"APPLIEDRESULT", ""}, {"CURESULT", "Y"}}
        Dim dtXls1 As DataTable = SEARCH_DATA1_dt(hPMS) '上半年
        'Dim dtXls2 As DataTable = SEARCH_DATA1_dt(2) '下半年
        'Dim dtXls3 As DataTable = SEARCH_DATA1_dt(3) '全年度
        Dim v_APPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH) '申請階段 1/2
        If v_APPSTAGE_SCH = "1" AndAlso TIMS.dtNODATA(dtXls1) Then
            Common.MessageBox(Me, "查無(上半年)匯出資料。")
            Exit Sub
        End If
        'If v_APPSTAGE_SCH = "2" AndAlso TIMS.dtNODATA(dtXls2) Then
        '    Common.MessageBox(Me, "查無(下半年)匯出資料。")
        '    Exit Sub
        'End If
        'If TIMS.dtNODATA(dtXls3) Then
        '    Common.MessageBox(Me, "查無(全年度)匯出資料。")
        '    Exit Sub
        'End If

        Dim cellsCOLSPNumF As String = "A{0}:N{0}"
        Dim cellsCOLSPNumF2 As String = "N{0}"
        Dim cellsCOLSPNumF3 As String() = "D6:D8,D10:D12,D14:D16,D18:D20,D22:D24,D26:D28".Split(",")
        Dim cellsCOLSPNumF4 As String() = "E6:E8,E10:E12,E14:E16,E18:E20,E22:E24,E26:E28".Split(",")
        'Dim cellsCOLSPNumF5 As String() = "N6:N8,N10:N12,N14:N16,N18:N20,N22:N24,N26:N28".Split(",")

        Dim strErrmsg As String = ""

        Dim s_ROCYEAR1 As String = CStr(CInt(sm.UserInfo.Years) - 1911) '年度
        Dim v_ddlAPPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH)
        Dim s_APPSTAGE_NM2 As String = TIMS.GET_APPSTAGE2_NM2(v_ddlAPPSTAGE_SCH) '申請階段
        Dim v_rblOrgKind2 As String = TIMS.GetListValue(rblOrgKind2)
        Dim s_PLANNAME As String = TIMS.GetListText(rblOrgKind2) '計畫
        Dim s_TITLENAME1 As String = String.Concat(s_ROCYEAR1, "年度", s_APPSTAGE_NM2, s_PLANNAME, "管制類課程比例五分署彙總表")
        Dim s_FILENAME1 As String = String.Concat(s_ROCYEAR1, s_APPSTAGE_NM2, s_PLANNAME, "管制類課程比例_", TIMS.GetDateNo())
        Dim sFileName As String = String.Concat(Cst_FileSavePath, TIMS.GetDateNo(), ".xlsx") '複製一份(Sample)
        Dim filePath1 As String = Server.MapPath(sFileName) '複製一份(Sample)
        Try
            IO.File.Copy(Server.MapPath(cst_SampleXLS), filePath1, True)
        Catch ex As Exception
            strErrmsg = ""
            strErrmsg += "目錄名稱或磁碟區標籤語法錯誤!!!" & vbCrLf
            strErrmsg += " (若連續出現多次(3次以上)請連絡系統管理者協助，謝謝)(造成您的不便深感抱歉)" & vbCrLf
            strErrmsg += ex.ToString & vbCrLf
            Common.MessageBox(Me, strErrmsg)
            'Return 'Exit Sub
        End Try
        If strErrmsg <> "" Then
            Try
                strErrmsg += "Path/File: " & filePath1 & vbCrLf
                strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
                'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                Call TIMS.WriteTraceLog(strErrmsg)
            Catch ex As Exception
            End Try
            Return 'Exit Sub
        End If

        SyncLock print_lock
            'ExcelPackage.LicenseContext = LicenseContext.Commercial
            'ExcelPackage.LicenseContext = LicenseContext.NonCommercial

            Dim file1 As New FileInfo(filePath1)
            Dim ndt As DateTime = Now
            Dim ep As New ExcelPackage(file1)
            'Dim ws As ExcelWorksheet = excel.Workbook.Worksheets.Add(rptName)
            Dim ws As ExcelWorksheet = ep.Workbook.Worksheets(0)

            ' 共用設定
            Dim fontName As String = "標楷體"
            Dim fontSize12s As Single = 12.0F
            ' 報表標題
            Using exlRow1 As ExcelRange = ws.Cells(String.Format(cellsCOLSPNumF, "1"))
                With exlRow1
                    .Merge = True
                    .Style.Font.Name = fontName
                    .Style.Font.Size = 16
                    .Value = s_TITLENAME1
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.VerticalAlignment = ExcelVerticalAlignment.Center
                End With
            End Using

            ' 列印日期
            Using exlRow2 As ExcelRange = ws.Cells(String.Format(cellsCOLSPNumF2, "2"))
                With exlRow2
                    .Merge = True
                    .Style.Font.Name = fontName
                    .Style.Font.Size = fontSize12s
                    .Value = String.Format("{0}.{1}.{2}", ndt.Year - 1911, ndt.Month, ndt.Day)
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    .Style.VerticalAlignment = ExcelVerticalAlignment.Center
                End With
            End Using

            ws.Cells("D3:I3").Value = String.Concat(s_ROCYEAR1, "上半年課程核定情形")
            'ws.Cells("J3:K4").Value = String.Concat(s_ROCYEAR1, "上半年")
            ws.Cells("L4:M4").Value = String.Concat(s_ROCYEAR1, "上半年管制類課程")
            'ws.Cells("Q4").Value = String.Concat(s_ROCYEAR1, "上半年餐飲類課程")
            'ws.Cells("V4").Value = String.Concat(s_ROCYEAR1, "年餐飲類課程")

            Dim i_DEFGOVCOST_F29 As Integer = 0
            Dim i_CLASSCNT_H29 As Integer = 0
            Dim i_DEFGOVCOSTG1 As Integer = 0
            Dim i_CLASSCNTG1 As Integer = 0
            Dim idxStr As Integer = 6
            For Each dr1 As DataRow In dtXls1.Rows

                SetCellValue(ws, "F" & idxStr, dr1("DEFGOVCOST"))
                SetCellValue(ws, "G" & idxStr, dr1("RATE13"))
                SetCellValue(ws, "H" & idxStr, dr1("CLASSCNT"))
                SetCellValue(ws, "I" & idxStr, dr1("RATE24"))

                If Convert.ToString(dr1("GCODE33")) <> "99" AndAlso Convert.ToString(dr1("DISTID")) = "999" Then
                    i_DEFGOVCOST_F29 += TIMS.VAL1(dr1("DEFGOVCOST"))
                    i_CLASSCNT_H29 += TIMS.VAL1(dr1("CLASSCNT"))
                End If

                If (Convert.ToString(dr1("GCODE33")) = "99" AndAlso Convert.ToString(dr1("DISTID")) = "001") Then
                    i_DEFGOVCOSTG1 += TIMS.VAL1(dr1("DEFGOVCOSTG1"))
                    i_CLASSCNTG1 += TIMS.VAL1(dr1("CLASSCNTG1"))
                    SetCellValue(ws, cellsCOLSPNumF3(0), dr1("DEFGOVCOSTG1"))
                    SetCellValue(ws, cellsCOLSPNumF4(0), dr1("CLASSCNTG1"))
                ElseIf (Convert.ToString(dr1("GCODE33")) = "99" AndAlso Convert.ToString(dr1("DISTID")) = "003") Then
                    i_DEFGOVCOSTG1 += TIMS.VAL1(dr1("DEFGOVCOSTG1"))
                    i_CLASSCNTG1 += TIMS.VAL1(dr1("CLASSCNTG1"))
                    SetCellValue(ws, cellsCOLSPNumF3(1), dr1("DEFGOVCOSTG1"))
                    SetCellValue(ws, cellsCOLSPNumF4(1), dr1("CLASSCNTG1"))
                ElseIf (Convert.ToString(dr1("GCODE33")) = "99" AndAlso Convert.ToString(dr1("DISTID")) = "004") Then
                    i_DEFGOVCOSTG1 += TIMS.VAL1(dr1("DEFGOVCOSTG1"))
                    i_CLASSCNTG1 += TIMS.VAL1(dr1("CLASSCNTG1"))
                    SetCellValue(ws, cellsCOLSPNumF3(2), dr1("DEFGOVCOSTG1"))
                    SetCellValue(ws, cellsCOLSPNumF4(2), dr1("CLASSCNTG1"))
                ElseIf (Convert.ToString(dr1("GCODE33")) = "99" AndAlso Convert.ToString(dr1("DISTID")) = "005") Then
                    i_DEFGOVCOSTG1 += TIMS.VAL1(dr1("DEFGOVCOSTG1"))
                    i_CLASSCNTG1 += TIMS.VAL1(dr1("CLASSCNTG1"))
                    SetCellValue(ws, cellsCOLSPNumF3(3), dr1("DEFGOVCOSTG1"))
                    SetCellValue(ws, cellsCOLSPNumF4(3), dr1("CLASSCNTG1"))
                ElseIf (Convert.ToString(dr1("GCODE33")) = "99" AndAlso Convert.ToString(dr1("DISTID")) = "006") Then
                    i_DEFGOVCOSTG1 += TIMS.VAL1(dr1("DEFGOVCOSTG1"))
                    i_CLASSCNTG1 += TIMS.VAL1(dr1("CLASSCNTG1"))
                    SetCellValue(ws, cellsCOLSPNumF3(4), dr1("DEFGOVCOSTG1"))
                    SetCellValue(ws, cellsCOLSPNumF4(4), dr1("CLASSCNTG1"))
                ElseIf (Convert.ToString(dr1("GCODE33")) = "99" AndAlso Convert.ToString(dr1("DISTID")) = "999") Then
                    SetCellValue(ws, cellsCOLSPNumF3(5), i_DEFGOVCOSTG1)
                    SetCellValue(ws, cellsCOLSPNumF4(5), i_CLASSCNTG1)
                    SetCellValue(ws, "F29", i_DEFGOVCOST_F29)
                    SetCellValue(ws, "H29", i_CLASSCNT_H29)
                End If

                If Convert.ToString(dr1("GCODE33")) = "99" Then
                    SetCellValue(ws, "L" & idxStr, "-")
                    SetCellValue(ws, "M" & idxStr, "-")
                ElseIf Convert.ToString(dr1("GCODE33")) = "06" Then
                    SetCellValue(ws, "L" & idxStr, "-")
                    SetCellValue(ws, "M" & idxStr, "-")
                Else
                    Dim i_B_COST As Double = TIMS.VAL1(ws.Cells("F" & idxStr).Value)
                    Dim i_A_COST As Double = TIMS.VAL1(ws.Cells("J" & idxStr).Value)
                    If i_A_COST > 0 AndAlso i_B_COST > 0 Then
                        SetCellValue(ws, "L" & idxStr, String.Concat(TIMS.ROUND(i_B_COST / i_A_COST * 100, 0), "%"))
                    End If
                    Dim i_D_CNT As Double = TIMS.VAL1(ws.Cells("H" & idxStr).Value)
                    Dim i_C_CNT As Double = TIMS.VAL1(ws.Cells("K" & idxStr).Value)
                    If i_C_CNT > 0 AndAlso i_D_CNT > 0 Then
                        SetCellValue(ws, "M" & idxStr, String.Concat(TIMS.ROUND(i_D_CNT / i_C_CNT * 100, 0), "%"))
                    End If
                End If
                'If (Convert.ToString(dr1("GCODE33")) = "06") Then
                '    Dim i_Q_RATE As Object = ws.Cells("I" & idxStr).Value
                '    SetCellValue(ws, "Q" & idxStr, i_Q_RATE)
                'End If

                idxStr += 1
            Next

            Dim iF26 As Double = If(i_DEFGOVCOSTG1 > 0, TIMS.ROUND(TIMS.VAL1(ws.Cells("F26").Value) / i_DEFGOVCOSTG1 * 100, 2), "0")
            Dim iF27 As Double = If(i_DEFGOVCOSTG1 > 0, TIMS.ROUND(TIMS.VAL1(ws.Cells("F27").Value) / i_DEFGOVCOSTG1 * 100, 2), "0")
            Dim iF28 As Double = If(i_DEFGOVCOSTG1 > 0, TIMS.ROUND(TIMS.VAL1(ws.Cells("F28").Value) / i_DEFGOVCOSTG1 * 100, 2), "0")
            Dim iF29 As Double = TIMS.ROUND(TIMS.VAL1(iF26 + iF27 + iF28), 2)
            SetCellValue(ws, "G26", String.Concat(iF26, "%"))
            SetCellValue(ws, "G27", String.Concat(iF27, "%"))
            SetCellValue(ws, "G28", String.Concat(iF28, "%"))
            SetCellValue(ws, "G29", String.Concat(iF29, "%"))

            Dim iH26 As Double = If(i_CLASSCNTG1 > 0, TIMS.ROUND(TIMS.VAL1(ws.Cells("H26").Value) / i_CLASSCNTG1 * 100, 2), "0")
            Dim iH27 As Double = If(i_CLASSCNTG1 > 0, TIMS.ROUND(TIMS.VAL1(ws.Cells("H27").Value) / i_CLASSCNTG1 * 100, 2), "0")
            Dim iH28 As Double = If(i_CLASSCNTG1 > 0, TIMS.ROUND(TIMS.VAL1(ws.Cells("H28").Value) / i_CLASSCNTG1 * 100, 2), "0")
            Dim iH29 As Double = TIMS.ROUND(TIMS.VAL1(iH26 + iH27 + iH28), 2)
            SetCellValue(ws, "I26", String.Concat(iH26, "%"))
            SetCellValue(ws, "I27", String.Concat(iH27, "%"))
            SetCellValue(ws, "I28", String.Concat(iH28, "%"))
            SetCellValue(ws, "I29", String.Concat(iH29, "%"))

            '刪除Temp中的資料 
            Call TIMS.MyFileDelete(filePath1)
            Dim V_ExpType As String = TIMS.GetListValue(RBListExpType)
            Select Case V_ExpType
                Case "EXCEL"
                    TIMS.ExpExcel_1(Me, strErrmsg, ep, Cst_FileSavePath, s_FILENAME1)
                    TIMS.Utl_RespWriteEnd(Me, objconn, "")
                Case "ODS"
                    TIMS.ExpODSl_1(Me, strErrmsg, ep, Cst_FileSavePath, s_FILENAME1)
                    TIMS.Utl_RespWriteEnd(Me, objconn, "")
                Case Else
                    Dim s_log1 As String = String.Format("ExpType(參數有誤)!!{0}", V_ExpType)
                    Common.MessageBox(Me, s_log1)
                    Return ' Exit Sub
            End Select
        End SyncLock

        '刪除Temp中的資料 'Call TIMS.MyFileDelete(myFileName1)
        If strErrmsg <> "" Then
            Common.MessageBox(Me, strErrmsg)
            Return
        End If
    End Sub

    Sub ExportXlsG2()
        Const cst_SampleXLS As String = "~\CR\03\Tmp\e1140407x1G2.xlsx"
        If Not IO.File.Exists(Server.MapPath(cst_SampleXLS)) Then
            Common.MessageBox(Me, "Sample檔案不存在")
            Return '  Exit Sub
        End If
        Const Cst_FileSavePath As String = "~/CR/03/Temp/" '"~\CO\01\Temp\"
        Call TIMS.MyCreateDir(Me, Cst_FileSavePath)

        Dim hPMS As New Hashtable From {{"APPSTAGE", 1}, {"APPLIEDRESULT", "Y"}, {"CURESULT", ""}}
        Dim dtXls1 As DataTable = SEARCH_DATA1_dt(hPMS) '上半年
        Dim hPMS2 As New Hashtable From {{"APPSTAGE", 2}, {"APPLIEDRESULT", ""}, {"CURESULT", "Y"}}
        Dim dtXls2 As DataTable = SEARCH_DATA1_dt(hPMS2) '下半年
        Dim hPMS3 As New Hashtable From {{"APPSTAGE", 3}}
        Dim dtXls3 As DataTable = SEARCH_DATA1_dt(hPMS3) '全年度
        Dim v_APPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH) '申請階段 1/2
        If v_APPSTAGE_SCH = "1" AndAlso TIMS.dtNODATA(dtXls1) Then
            Common.MessageBox(Me, "查無(上半年)匯出資料。")
            Exit Sub
        End If
        If v_APPSTAGE_SCH = "2" AndAlso TIMS.dtNODATA(dtXls2) Then
            Common.MessageBox(Me, "查無(下半年)匯出資料。")
            Exit Sub
        End If
        If TIMS.dtNODATA(dtXls3) Then
            Common.MessageBox(Me, "查無(全年度)匯出資料。")
            Exit Sub
        End If

        Dim cellsCOLSPNumF As String = "A{0}:V{0}"
        Dim cellsCOLSPNumF2 As String = "V{0}"
        Dim cellsCOLSPNumF3 As String() = "D6:D8,D10:D12,D14:D16,D18:D20,D22:D24,D26:D28".Split(",")
        Dim cellsCOLSPNumF4 As String() = "E6:E8,E10:E12,E14:E16,E18:E20,E22:E24,E26:E28".Split(",")
        Dim cellsCOLSPNumF5 As String() = "N6:N8,N10:N12,N14:N16,N18:N20,N22:N24,N26:N28".Split(",")

        Dim strErrmsg As String = ""

        Dim s_ROCYEAR1 As String = CStr(CInt(sm.UserInfo.Years) - 1911) '年度
        Dim v_ddlAPPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH)
        Dim s_APPSTAGE_NM2 As String = TIMS.GET_APPSTAGE2_NM2(v_ddlAPPSTAGE_SCH) '申請階段
        Dim v_rblOrgKind2 As String = TIMS.GetListValue(rblOrgKind2)
        Dim s_PLANNAME As String = TIMS.GetListText(rblOrgKind2) '計畫
        Dim s_TITLENAME1 As String = String.Concat(s_ROCYEAR1, "年度", s_APPSTAGE_NM2, s_PLANNAME, "管制課程比例五分署彙總表")
        Dim s_FILENAME1 As String = String.Concat(s_ROCYEAR1, s_APPSTAGE_NM2, s_PLANNAME, "管制類課程比例_", TIMS.GetDateNo())
        Dim sFileName As String = String.Concat(Cst_FileSavePath, TIMS.GetDateNo(), ".xlsx") '複製一份(Sample)
        Dim filePath1 As String = Server.MapPath(sFileName) '複製一份(Sample)
        Try
            IO.File.Copy(Server.MapPath(cst_SampleXLS), filePath1, True)
        Catch ex As Exception
            strErrmsg = ""
            strErrmsg += "目錄名稱或磁碟區標籤語法錯誤!!!" & vbCrLf
            strErrmsg += " (若連續出現多次(3次以上)請連絡系統管理者協助，謝謝)(造成您的不便深感抱歉)" & vbCrLf
            strErrmsg += ex.ToString & vbCrLf
            Common.MessageBox(Me, strErrmsg)
            'Return 'Exit Sub
        End Try
        If strErrmsg <> "" Then
            Try
                strErrmsg += "Path/File: " & filePath1 & vbCrLf
                strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
                'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                Call TIMS.WriteTraceLog(strErrmsg)
            Catch ex As Exception
            End Try
            Return 'Exit Sub
        End If

        SyncLock print_lock
            'ExcelPackage.LicenseContext = LicenseContext.Commercial
            'ExcelPackage.LicenseContext = LicenseContext.NonCommercial

            Dim file1 As New FileInfo(filePath1)
            Dim ndt As DateTime = Now
            Dim ep As New ExcelPackage(file1)
            'Dim ws As ExcelWorksheet = excel.Workbook.Worksheets.Add(rptName)
            Dim ws As ExcelWorksheet = ep.Workbook.Worksheets(0)

            ' 共用設定
            Dim fontName As String = "標楷體"
            Dim fontSize12s As Single = 12.0F
            ' 報表標題
            Using exlRow1 As ExcelRange = ws.Cells(String.Format(cellsCOLSPNumF, "1"))
                With exlRow1
                    .Merge = True
                    .Style.Font.Name = fontName
                    .Style.Font.Size = 16
                    .Value = s_TITLENAME1
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.VerticalAlignment = ExcelVerticalAlignment.Center
                End With
            End Using

            ' 列印日期
            Using exlRow2 As ExcelRange = ws.Cells(String.Format(cellsCOLSPNumF2, "2"))
                With exlRow2
                    .Merge = True
                    .Style.Font.Name = fontName
                    .Style.Font.Size = fontSize12s
                    .Value = String.Format("{0}.{1}.{2}", ndt.Year - 1911, ndt.Month, ndt.Day)
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    .Style.VerticalAlignment = ExcelVerticalAlignment.Center
                End With
            End Using

            ws.Cells("D3:I4").Value = String.Concat(s_ROCYEAR1, "下半年課程核定情形")
            ws.Cells("J3:K4").Value = String.Concat(s_ROCYEAR1, "上半年")
            ws.Cells("L3:N4").Value = String.Concat(s_ROCYEAR1, "全年度")
            ws.Cells("Q4:R4").Value = String.Concat(s_ROCYEAR1, "全年度管制類課程")
            ws.Cells("V4").Value = String.Concat(s_ROCYEAR1, "年餐飲類課程")

            Dim i_DEFGOVCOST_F29 As Integer = 0
            Dim i_CLASSCNT_H29 As Integer = 0
            Dim i_DEFGOVCOSTG1 As Integer = 0
            Dim i_CLASSCNTG1 As Integer = 0
            Dim idxStr As Integer = 6
            For Each dr2 As DataRow In dtXls2.Rows
                SetCellValue(ws, "F" & idxStr, dr2("DEFGOVCOST"))
                SetCellValue(ws, "G" & idxStr, dr2("RATE13"))
                SetCellValue(ws, "H" & idxStr, dr2("CLASSCNT"))
                SetCellValue(ws, "I" & idxStr, dr2("RATE23"))

                If Convert.ToString(dr2("GCODE33")) <> "99" AndAlso Convert.ToString(dr2("DISTID")) = "999" Then
                    i_DEFGOVCOST_F29 += TIMS.VAL1(dr2("DEFGOVCOST"))
                    i_CLASSCNT_H29 += TIMS.VAL1(dr2("CLASSCNT"))
                End If

                If (Convert.ToString(dr2("GCODE33")) = "99" AndAlso Convert.ToString(dr2("DISTID")) = "001") Then
                    i_DEFGOVCOSTG1 += TIMS.VAL1(dr2("DEFGOVCOSTG1"))
                    i_CLASSCNTG1 += TIMS.VAL1(dr2("CLASSCNTG1"))
                    SetCellValue(ws, cellsCOLSPNumF3(0), dr2("DEFGOVCOSTG1"))
                    SetCellValue(ws, cellsCOLSPNumF4(0), dr2("CLASSCNTG1"))
                ElseIf (Convert.ToString(dr2("GCODE33")) = "99" AndAlso Convert.ToString(dr2("DISTID")) = "003") Then
                    i_DEFGOVCOSTG1 += TIMS.VAL1(dr2("DEFGOVCOSTG1"))
                    i_CLASSCNTG1 += TIMS.VAL1(dr2("CLASSCNTG1"))
                    SetCellValue(ws, cellsCOLSPNumF3(1), dr2("DEFGOVCOSTG1"))
                    SetCellValue(ws, cellsCOLSPNumF4(1), dr2("CLASSCNTG1"))
                ElseIf (Convert.ToString(dr2("GCODE33")) = "99" AndAlso Convert.ToString(dr2("DISTID")) = "004") Then
                    i_DEFGOVCOSTG1 += TIMS.VAL1(dr2("DEFGOVCOSTG1"))
                    i_CLASSCNTG1 += TIMS.VAL1(dr2("CLASSCNTG1"))
                    SetCellValue(ws, cellsCOLSPNumF3(2), dr2("DEFGOVCOSTG1"))
                    SetCellValue(ws, cellsCOLSPNumF4(2), dr2("CLASSCNTG1"))
                ElseIf (Convert.ToString(dr2("GCODE33")) = "99" AndAlso Convert.ToString(dr2("DISTID")) = "005") Then
                    i_DEFGOVCOSTG1 += TIMS.VAL1(dr2("DEFGOVCOSTG1"))
                    i_CLASSCNTG1 += TIMS.VAL1(dr2("CLASSCNTG1"))
                    SetCellValue(ws, cellsCOLSPNumF3(3), dr2("DEFGOVCOSTG1"))
                    SetCellValue(ws, cellsCOLSPNumF4(3), dr2("CLASSCNTG1"))
                ElseIf (Convert.ToString(dr2("GCODE33")) = "99" AndAlso Convert.ToString(dr2("DISTID")) = "006") Then
                    i_DEFGOVCOSTG1 += TIMS.VAL1(dr2("DEFGOVCOSTG1"))
                    i_CLASSCNTG1 += TIMS.VAL1(dr2("CLASSCNTG1"))
                    SetCellValue(ws, cellsCOLSPNumF3(4), dr2("DEFGOVCOSTG1"))
                    SetCellValue(ws, cellsCOLSPNumF4(4), dr2("CLASSCNTG1"))
                ElseIf (Convert.ToString(dr2("GCODE33")) = "99" AndAlso Convert.ToString(dr2("DISTID")) = "999") Then
                    SetCellValue(ws, cellsCOLSPNumF3(5), i_DEFGOVCOSTG1)
                    SetCellValue(ws, cellsCOLSPNumF4(5), i_CLASSCNTG1)
                    SetCellValue(ws, "F29", i_DEFGOVCOST_F29)
                    SetCellValue(ws, "H29", i_CLASSCNT_H29)
                End If

                idxStr += 1
            Next

            Dim iF26 As Double = If(i_DEFGOVCOSTG1 > 0, TIMS.ROUND(TIMS.VAL1(ws.Cells("F26").Value) / i_DEFGOVCOSTG1 * 100, 2), "0")
            Dim iF27 As Double = If(i_DEFGOVCOSTG1 > 0, TIMS.ROUND(TIMS.VAL1(ws.Cells("F27").Value) / i_DEFGOVCOSTG1 * 100, 2), "0")
            Dim iF28 As Double = If(i_DEFGOVCOSTG1 > 0, TIMS.ROUND(TIMS.VAL1(ws.Cells("F28").Value) / i_DEFGOVCOSTG1 * 100, 2), "0")
            Dim iF29 As Double = TIMS.ROUND(TIMS.VAL1(iF26 + iF27 + iF28), 2)
            SetCellValue(ws, "G26", String.Concat(iF26, "%"))
            SetCellValue(ws, "G27", String.Concat(iF27, "%"))
            SetCellValue(ws, "G28", String.Concat(iF28, "%"))
            SetCellValue(ws, "G29", String.Concat(iF29, "%"))

            Dim iH26 As Double = If(i_CLASSCNTG1 > 0, TIMS.ROUND(TIMS.VAL1(ws.Cells("H26").Value) / i_CLASSCNTG1 * 100, 2), "0")
            Dim iH27 As Double = If(i_CLASSCNTG1 > 0, TIMS.ROUND(TIMS.VAL1(ws.Cells("H27").Value) / i_CLASSCNTG1 * 100, 2), "0")
            Dim iH28 As Double = If(i_CLASSCNTG1 > 0, TIMS.ROUND(TIMS.VAL1(ws.Cells("H28").Value) / i_CLASSCNTG1 * 100, 2), "0")
            Dim iH29 As Double = TIMS.ROUND(TIMS.VAL1(iH26 + iH27 + iH28), 2)
            SetCellValue(ws, "I26", String.Concat(iH26, "%"))
            SetCellValue(ws, "I27", String.Concat(iH27, "%"))
            SetCellValue(ws, "I28", String.Concat(iH28, "%"))
            SetCellValue(ws, "I29", String.Concat(iH29, "%"))

            Dim i_DEFGOVCOST_J29 As Integer = 0
            Dim i_CLASSCNT_K29 As Integer = 0
            idxStr = 6
            For Each dr1 As DataRow In dtXls1.Rows
                SetCellValue(ws, "J" & idxStr, dr1("DEFGOVCOST"))
                SetCellValue(ws, "K" & idxStr, dr1("CLASSCNT"))

                If Convert.ToString(dr1("GCODE33")) <> "99" AndAlso Convert.ToString(dr1("DISTID")) = "999" Then
                    i_DEFGOVCOST_J29 += TIMS.VAL1(dr1("DEFGOVCOST"))
                    i_CLASSCNT_K29 += TIMS.VAL1(dr1("CLASSCNT"))
                ElseIf (Convert.ToString(dr1("GCODE33")) = "99" AndAlso Convert.ToString(dr1("DISTID")) = "999") Then
                    SetCellValue(ws, "J29", i_DEFGOVCOST_J29)
                    SetCellValue(ws, "K29", i_CLASSCNT_K29)
                End If

                Dim iDEFGOVCOST As Double = TIMS.VAL1(ws.Cells("J" & idxStr).Value) + TIMS.VAL1(ws.Cells("F" & idxStr).Value)
                Dim iCLASSCNT As Double = TIMS.VAL1(ws.Cells("K" & idxStr).Value) + TIMS.VAL1(ws.Cells("H" & idxStr).Value)
                SetCellValue(ws, "L" & idxStr, iDEFGOVCOST)
                SetCellValue(ws, "M" & idxStr, iCLASSCNT)

                Dim i_A_COST As Double = TIMS.VAL1(ws.Cells("L" & idxStr).Value)
                Dim i_B_COST As Double = TIMS.VAL1(ws.Cells("O" & idxStr).Value)
                If i_A_COST > 0 AndAlso i_B_COST > 0 Then SetCellValue(ws, "Q" & idxStr, String.Concat(TIMS.ROUND(i_A_COST / i_B_COST * 100, 0), "%"))
                Dim i_C_CNT As Double = TIMS.VAL1(ws.Cells("M" & idxStr).Value)
                Dim i_D_CNT As Double = TIMS.VAL1(ws.Cells("P" & idxStr).Value)
                If i_C_CNT > 0 AndAlso i_D_CNT > 0 Then SetCellValue(ws, "R" & idxStr, String.Concat(TIMS.ROUND(i_C_CNT / i_D_CNT * 100, 0), "%"))

                idxStr += 1
            Next

            i_CLASSCNTG1 = 0
            idxStr = 6
            For Each dr3 As DataRow In dtXls3.Rows
                If (Convert.ToString(dr3("GCODE33")) = "06") Then
                    Dim i_C_CNT As Double = TIMS.VAL1(ws.Cells("M" & idxStr).Value)
                    Dim i_E_CNT As Double = TIMS.VAL1(dr3("CLASSCNTG1"))
                    If i_C_CNT > 0 AndAlso i_E_CNT > 0 Then
                        SetCellValue(ws, "V" & idxStr, String.Concat(TIMS.ROUND(i_C_CNT / i_E_CNT * 100, 2), "%"))
                    End If
                End If

                If (Convert.ToString(dr3("GCODE33")) = "99" AndAlso Convert.ToString(dr3("DISTID")) = "001") Then
                    i_CLASSCNTG1 += TIMS.VAL1(dr3("CLASSCNTG1"))
                    SetCellValue(ws, cellsCOLSPNumF5(0), dr3("CLASSCNTG1"))
                ElseIf (Convert.ToString(dr3("GCODE33")) = "99" AndAlso Convert.ToString(dr3("DISTID")) = "003") Then
                    i_CLASSCNTG1 += TIMS.VAL1(dr3("CLASSCNTG1"))
                    SetCellValue(ws, cellsCOLSPNumF5(1), dr3("CLASSCNTG1"))
                ElseIf (Convert.ToString(dr3("GCODE33")) = "99" AndAlso Convert.ToString(dr3("DISTID")) = "004") Then
                    i_CLASSCNTG1 += TIMS.VAL1(dr3("CLASSCNTG1"))
                    SetCellValue(ws, cellsCOLSPNumF5(2), dr3("CLASSCNTG1"))
                ElseIf (Convert.ToString(dr3("GCODE33")) = "99" AndAlso Convert.ToString(dr3("DISTID")) = "005") Then
                    i_CLASSCNTG1 += TIMS.VAL1(dr3("CLASSCNTG1"))
                    SetCellValue(ws, cellsCOLSPNumF5(3), dr3("CLASSCNTG1"))
                ElseIf (Convert.ToString(dr3("GCODE33")) = "99" AndAlso Convert.ToString(dr3("DISTID")) = "006") Then
                    i_CLASSCNTG1 += TIMS.VAL1(dr3("CLASSCNTG1"))
                    SetCellValue(ws, cellsCOLSPNumF5(4), dr3("CLASSCNTG1"))
                ElseIf (Convert.ToString(dr3("GCODE33")) = "99" AndAlso Convert.ToString(dr3("DISTID")) = "999") Then
                    SetCellValue(ws, cellsCOLSPNumF5(5), i_CLASSCNTG1)
                End If

                idxStr += 1
            Next

            Dim iM28 As Double = TIMS.VAL1(ws.Cells("M28").Value)
            If iM28 > 0 AndAlso i_CLASSCNTG1 > 0 Then SetCellValue(ws, "V28", String.Concat(TIMS.ROUND(iM28 / i_CLASSCNTG1 * 100, 2), "%"))

            Dim i_L26_COST As Double = TIMS.VAL1(ws.Cells("L26").Value)
            Dim i_L27_COST As Double = TIMS.VAL1(ws.Cells("L27").Value)
            Dim i_L18_COST As Double = TIMS.VAL1(ws.Cells("L18").Value)
            Dim i_L19_COST As Double = TIMS.VAL1(ws.Cells("L19").Value)
            Dim i_M26_VAL As Double = TIMS.VAL1(ws.Cells("M26").Value)
            Dim i_M27_VAL As Double = TIMS.VAL1(ws.Cells("M27").Value)
            Dim i_M18_VAL As Double = TIMS.VAL1(ws.Cells("M18").Value)
            Dim i_M19_VAL As Double = TIMS.VAL1(ws.Cells("M19").Value)
            ws.Cells("L26").Value = (i_L26_COST + i_L19_COST)
            ws.Cells("L27").Value = (i_L27_COST - i_L19_COST)
            ws.Cells("M26").Value = (i_M26_VAL + i_M19_VAL)
            ws.Cells("M27").Value = (i_M27_VAL - i_M19_VAL)
            'SetCellValue(ws, "L27", (i_L27_COST - i_L19_COST))
            'SetCellValue(ws, "L26", (i_L26_COST + i_L19_COST))
            ws.Cells("L18:L19").Merge = True
            ws.Cells("L18").Value = (i_L18_COST + i_L19_COST)
            ws.Cells("M18:M19").Merge = True
            ws.Cells("M18").Value = (i_M18_VAL + i_M19_VAL)
            ws.Cells("Q18:Q19").Merge = True
            ws.Cells("R18:R19").Merge = True

            Dim i_A_L18_COST As Double = TIMS.VAL1(ws.Cells("L18").Value)
            Dim i_B_O18_COST As Double = TIMS.VAL1(ws.Cells("O18").Value)
            If i_A_L18_COST > 0 AndAlso i_B_O18_COST > 0 Then ws.Cells("Q18").Value = String.Concat(TIMS.ROUND(i_A_L18_COST / i_B_O18_COST * 100, 0), "%")
            Dim i_C_M18_CNT As Double = TIMS.VAL1(ws.Cells("M18").Value)
            Dim i_D_P18_CNT As Double = TIMS.VAL1(ws.Cells("P18").Value)
            If i_C_M18_CNT > 0 AndAlso i_D_P18_CNT > 0 Then ws.Cells("R18").Value = String.Concat(TIMS.ROUND(i_C_M18_CNT / i_D_P18_CNT * 100, 0), "%")
            For idxS As Integer = 26 To 27
                Dim i_A_COST As Double = TIMS.VAL1(ws.Cells("L" & idxS).Value)
                Dim i_B_COST As Double = TIMS.VAL1(ws.Cells("O" & idxS).Value)
                If i_A_COST > 0 AndAlso i_B_COST > 0 Then SetCellValue(ws, "Q" & idxS, String.Concat(TIMS.ROUND(i_A_COST / i_B_COST * 100, 0), "%"))
                Dim i_C_CNT As Double = TIMS.VAL1(ws.Cells("M" & idxS).Value)
                Dim i_D_CNT As Double = TIMS.VAL1(ws.Cells("P" & idxS).Value)
                If i_C_CNT > 0 AndAlso i_D_CNT > 0 Then SetCellValue(ws, "R" & idxS, String.Concat(TIMS.ROUND(i_C_CNT / i_D_CNT * 100, 0), "%"))
            Next
            Dim idxSU As Integer() = {9, 13, 17, 21, 25, 29}
            For Each idxU As Integer In idxSU
                ws.Cells("Q" & idxU).Value = "-"
                ws.Cells("R" & idxU).Value = "-"
            Next

            '刪除Temp中的資料 
            Call TIMS.MyFileDelete(filePath1)
            Dim V_ExpType As String = TIMS.GetListValue(RBListExpType)
            Select Case V_ExpType
                Case "EXCEL"
                    TIMS.ExpExcel_1(Me, strErrmsg, ep, Cst_FileSavePath, s_FILENAME1)
                    TIMS.Utl_RespWriteEnd(Me, objconn, "")
                Case "ODS"
                    TIMS.ExpODSl_1(Me, strErrmsg, ep, Cst_FileSavePath, s_FILENAME1)
                    TIMS.Utl_RespWriteEnd(Me, objconn, "")
                Case Else
                    Dim s_log1 As String = String.Format("ExpType(參數有誤)!!{0}", V_ExpType)
                    Common.MessageBox(Me, s_log1)
                    Return ' Exit Sub
            End Select
        End SyncLock

        '刪除Temp中的資料 'Call TIMS.MyFileDelete(myFileName1)
        If strErrmsg <> "" Then
            Common.MessageBox(Me, strErrmsg)
            Return
        End If
    End Sub

    Sub ExportXlsW2()
        Const cst_SampleXLS As String = "~\CR\03\Tmp\e1140407x1W2.xlsx"
        If Not IO.File.Exists(Server.MapPath(cst_SampleXLS)) Then
            Common.MessageBox(Me, "Sample檔案不存在")
            Return '  Exit Sub
        End If
        Const Cst_FileSavePath As String = "~/CR/03/Temp/" '"~\CO\01\Temp\"
        Call TIMS.MyCreateDir(Me, Cst_FileSavePath)

        Dim hPMS As New Hashtable From {{"APPSTAGE", 1}, {"APPLIEDRESULT", "Y"}, {"CURESULT", ""}}
        Dim dtXls1 As DataTable = SEARCH_DATA1_dt(hPMS) '上半年
        Dim hPMS2 As New Hashtable From {{"APPSTAGE", 2}, {"APPLIEDRESULT", ""}, {"CURESULT", "Y"}}
        Dim dtXls2 As DataTable = SEARCH_DATA1_dt(hPMS2) '下半年
        Dim hPMS3 As New Hashtable From {{"APPSTAGE", 3}}
        Dim dtXls3 As DataTable = SEARCH_DATA1_dt(hPMS3) '全年度

        Dim v_APPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH) '申請階段 1/2
        If v_APPSTAGE_SCH = "1" AndAlso TIMS.dtNODATA(dtXls1) Then
            Common.MessageBox(Me, "查無(上半年)匯出資料。")
            Exit Sub
        End If
        If v_APPSTAGE_SCH = "2" AndAlso TIMS.dtNODATA(dtXls2) Then
            Common.MessageBox(Me, "查無(下半年)匯出資料。")
            Exit Sub
        End If
        If TIMS.dtNODATA(dtXls3) Then
            Common.MessageBox(Me, "查無(全年度)匯出資料。")
            Exit Sub
        End If

        Dim cellsCOLSPNumF As String = "A{0}:R{0}"
        Dim cellsCOLSPNumF2 As String = "R{0}"
        Dim cellsCOLSPNumF3 As String() = "D6:D8,D10:D12,D14:D16,D18:D20,D22:D24,D26:D28".Split(",")
        Dim cellsCOLSPNumF4 As String() = "E6:E8,E10:E12,E14:E16,E18:E20,E22:E24,E26:E28".Split(",")
        Dim cellsCOLSPNumF5 As String() = "N6:N8,N10:N12,N14:N16,N18:N20,N22:N24,N26:N28".Split(",")

        Dim strErrmsg As String = ""

        Dim s_ROCYEAR1 As String = CStr(CInt(sm.UserInfo.Years) - 1911) '年度
        Dim v_ddlAPPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH)
        Dim s_APPSTAGE_NM2 As String = TIMS.GET_APPSTAGE2_NM2(v_ddlAPPSTAGE_SCH) '申請階段
        Dim v_rblOrgKind2 As String = TIMS.GetListValue(rblOrgKind2)
        Dim s_PLANNAME As String = TIMS.GetListText(rblOrgKind2) '計畫
        Dim s_TITLENAME1 As String = String.Concat(s_ROCYEAR1, "年度", s_APPSTAGE_NM2, s_PLANNAME, "管制課程比例五分署彙總表")
        Dim s_FILENAME1 As String = String.Concat(s_ROCYEAR1, s_APPSTAGE_NM2, s_PLANNAME, "管制類課程比例_", TIMS.GetDateNo())
        Dim sFileName As String = String.Concat(Cst_FileSavePath, TIMS.GetDateNo(), ".xlsx") '複製一份(Sample)
        Dim filePath1 As String = Server.MapPath(sFileName) '複製一份(Sample)
        Try
            IO.File.Copy(Server.MapPath(cst_SampleXLS), filePath1, True)
        Catch ex As Exception
            strErrmsg = ""
            strErrmsg += "目錄名稱或磁碟區標籤語法錯誤!!!" & vbCrLf
            strErrmsg += " (若連續出現多次(3次以上)請連絡系統管理者協助，謝謝)(造成您的不便深感抱歉)" & vbCrLf
            strErrmsg += ex.ToString & vbCrLf
            Common.MessageBox(Me, strErrmsg)
            'Return 'Exit Sub
        End Try
        If strErrmsg <> "" Then
            Try
                strErrmsg += "Path/File: " & filePath1 & vbCrLf
                strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
                'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                Call TIMS.WriteTraceLog(strErrmsg)
            Catch ex As Exception
            End Try
            Return 'Exit Sub
        End If

        SyncLock print_lock
            'ExcelPackage.LicenseContext = LicenseContext.Commercial
            'ExcelPackage.LicenseContext = LicenseContext.NonCommercial

            Dim file1 As New FileInfo(filePath1)
            Dim ndt As DateTime = Now
            Dim ep As New ExcelPackage(file1)
            'Dim ws As ExcelWorksheet = excel.Workbook.Worksheets.Add(rptName)
            Dim ws As ExcelWorksheet = ep.Workbook.Worksheets(0)

            ' 共用設定
            Dim fontName As String = "標楷體"
            Dim fontSize12s As Single = 12.0F
            ' 報表標題
            Using exlRow1 As ExcelRange = ws.Cells(String.Format(cellsCOLSPNumF, "1"))
                With exlRow1
                    .Merge = True
                    .Style.Font.Name = fontName
                    .Style.Font.Size = 16
                    .Value = s_TITLENAME1
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.VerticalAlignment = ExcelVerticalAlignment.Center
                End With
            End Using

            ' 列印日期
            Using exlRow2 As ExcelRange = ws.Cells(String.Format(cellsCOLSPNumF2, "2"))
                With exlRow2
                    .Merge = True
                    .Style.Font.Name = fontName
                    .Style.Font.Size = fontSize12s
                    .Value = String.Format("{0}.{1}.{2}", ndt.Year - 1911, ndt.Month, ndt.Day)
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    .Style.VerticalAlignment = ExcelVerticalAlignment.Center
                End With
            End Using

            ws.Cells("D3:I3").Value = String.Concat(s_ROCYEAR1, "下半年課程核定情形")
            ws.Cells("J3:K3").Value = String.Concat(s_ROCYEAR1, "上半年")
            ws.Cells("L3:N3").Value = String.Concat(s_ROCYEAR1, "全年度")
            ws.Cells("Q4:R4").Value = String.Concat(s_ROCYEAR1, "全年度管制類課程")

            'F/G/H/I/
            Dim i_DEFGOVCOST_F29 As Integer = 0
            Dim i_CLASSCNT_H29 As Integer = 0
            Dim i_DEFGOVCOSTG1 As Integer = 0
            Dim i_CLASSCNTG1 As Integer = 0
            Dim idxStr As Integer = 6
            For Each dr2 As DataRow In dtXls2.Rows
                SetCellValue(ws, "F" & idxStr, dr2("DEFGOVCOST"))
                SetCellValue(ws, "G" & idxStr, dr2("RATE13"))
                SetCellValue(ws, "H" & idxStr, dr2("CLASSCNT"))
                SetCellValue(ws, "I" & idxStr, dr2("RATE23"))

                If Convert.ToString(dr2("GCODE33")) <> "99" AndAlso Convert.ToString(dr2("DISTID")) = "999" Then
                    i_DEFGOVCOST_F29 += TIMS.VAL1(dr2("DEFGOVCOST"))
                    i_CLASSCNT_H29 += TIMS.VAL1(dr2("CLASSCNT"))
                End If

                If (Convert.ToString(dr2("GCODE33")) = "99" AndAlso Convert.ToString(dr2("DISTID")) = "001") Then
                    i_DEFGOVCOSTG1 += TIMS.VAL1(dr2("DEFGOVCOSTG1"))
                    i_CLASSCNTG1 += TIMS.VAL1(dr2("CLASSCNTG1"))
                    SetCellValue(ws, cellsCOLSPNumF3(0), dr2("DEFGOVCOSTG1"))
                    SetCellValue(ws, cellsCOLSPNumF4(0), dr2("CLASSCNTG1"))
                ElseIf (Convert.ToString(dr2("GCODE33")) = "99" AndAlso Convert.ToString(dr2("DISTID")) = "003") Then
                    i_DEFGOVCOSTG1 += TIMS.VAL1(dr2("DEFGOVCOSTG1"))
                    i_CLASSCNTG1 += TIMS.VAL1(dr2("CLASSCNTG1"))
                    SetCellValue(ws, cellsCOLSPNumF3(1), dr2("DEFGOVCOSTG1"))
                    SetCellValue(ws, cellsCOLSPNumF4(1), dr2("CLASSCNTG1"))
                ElseIf (Convert.ToString(dr2("GCODE33")) = "99" AndAlso Convert.ToString(dr2("DISTID")) = "004") Then
                    i_DEFGOVCOSTG1 += TIMS.VAL1(dr2("DEFGOVCOSTG1"))
                    i_CLASSCNTG1 += TIMS.VAL1(dr2("CLASSCNTG1"))
                    SetCellValue(ws, cellsCOLSPNumF3(2), dr2("DEFGOVCOSTG1"))
                    SetCellValue(ws, cellsCOLSPNumF4(2), dr2("CLASSCNTG1"))
                ElseIf (Convert.ToString(dr2("GCODE33")) = "99" AndAlso Convert.ToString(dr2("DISTID")) = "005") Then
                    i_DEFGOVCOSTG1 += TIMS.VAL1(dr2("DEFGOVCOSTG1"))
                    i_CLASSCNTG1 += TIMS.VAL1(dr2("CLASSCNTG1"))
                    SetCellValue(ws, cellsCOLSPNumF3(3), dr2("DEFGOVCOSTG1"))
                    SetCellValue(ws, cellsCOLSPNumF4(3), dr2("CLASSCNTG1"))
                ElseIf (Convert.ToString(dr2("GCODE33")) = "99" AndAlso Convert.ToString(dr2("DISTID")) = "006") Then
                    i_DEFGOVCOSTG1 += TIMS.VAL1(dr2("DEFGOVCOSTG1"))
                    i_CLASSCNTG1 += TIMS.VAL1(dr2("CLASSCNTG1"))
                    SetCellValue(ws, cellsCOLSPNumF3(4), dr2("DEFGOVCOSTG1"))
                    SetCellValue(ws, cellsCOLSPNumF4(4), dr2("CLASSCNTG1"))
                ElseIf (Convert.ToString(dr2("GCODE33")) = "99" AndAlso Convert.ToString(dr2("DISTID")) = "999") Then
                    SetCellValue(ws, cellsCOLSPNumF3(5), i_DEFGOVCOSTG1)
                    SetCellValue(ws, cellsCOLSPNumF4(5), i_CLASSCNTG1)
                    SetCellValue(ws, "F29", i_DEFGOVCOST_F29)
                    SetCellValue(ws, "H29", i_CLASSCNT_H29)
                End If

                idxStr += 1
            Next

            '管制類核定總補助費占比
            Dim iF26 As Double = If(i_DEFGOVCOSTG1 > 0, TIMS.ROUND(TIMS.VAL1(ws.Cells("F26").Value) / i_DEFGOVCOSTG1 * 100, 2), "0")
            Dim iF27 As Double = If(i_DEFGOVCOSTG1 > 0, TIMS.ROUND(TIMS.VAL1(ws.Cells("F27").Value) / i_DEFGOVCOSTG1 * 100, 2), "0")
            Dim iF28 As Double = If(i_DEFGOVCOSTG1 > 0, TIMS.ROUND(TIMS.VAL1(ws.Cells("F28").Value) / i_DEFGOVCOSTG1 * 100, 2), "0")
            Dim iF29 As Double = TIMS.ROUND(TIMS.VAL1(iF26 + iF27 + iF28), 2)
            SetCellValue(ws, "G26", String.Concat(iF26, "%"))
            SetCellValue(ws, "G27", String.Concat(iF27, "%"))
            SetCellValue(ws, "G28", String.Concat(iF28, "%"))
            SetCellValue(ws, "G29", String.Concat(iF29, "%"))

            '管制類核定班次占比
            Dim iH26 As Double = If(i_CLASSCNTG1 > 0, TIMS.ROUND(TIMS.VAL1(ws.Cells("H26").Value) / i_CLASSCNTG1 * 100, 2), "0")
            Dim iH27 As Double = If(i_CLASSCNTG1 > 0, TIMS.ROUND(TIMS.VAL1(ws.Cells("H27").Value) / i_CLASSCNTG1 * 100, 2), "0")
            Dim iH28 As Double = If(i_CLASSCNTG1 > 0, TIMS.ROUND(TIMS.VAL1(ws.Cells("H28").Value) / i_CLASSCNTG1 * 100, 2), "0")
            Dim iH29 As Double = TIMS.ROUND(TIMS.VAL1(iH26 + iH27 + iH28), 2)
            SetCellValue(ws, "I26", String.Concat(iH26, "%"))
            SetCellValue(ws, "I27", String.Concat(iH27, "%"))
            SetCellValue(ws, "I28", String.Concat(iH28, "%"))
            SetCellValue(ws, "I29", String.Concat(iH29, "%"))

            'J/K/
            Dim i_DEFGOVCOST_J29 As Integer = 0
            Dim i_CLASSCNT_K29 As Integer = 0
            idxStr = 6
            For Each dr1 As DataRow In dtXls1.Rows
                SetCellValue(ws, "J" & idxStr, dr1("DEFGOVCOST"))
                SetCellValue(ws, "K" & idxStr, dr1("CLASSCNT"))

                If Convert.ToString(dr1("GCODE33")) <> "99" AndAlso Convert.ToString(dr1("DISTID")) = "999" Then
                    i_DEFGOVCOST_J29 += TIMS.VAL1(dr1("DEFGOVCOST"))
                    i_CLASSCNT_K29 += TIMS.VAL1(dr1("CLASSCNT"))
                ElseIf (Convert.ToString(dr1("GCODE33")) = "99" AndAlso Convert.ToString(dr1("DISTID")) = "999") Then
                    SetCellValue(ws, "J29", i_DEFGOVCOST_J29)
                    SetCellValue(ws, "K29", i_CLASSCNT_K29)
                End If

                Dim iDEFGOVCOST As Double = TIMS.VAL1(ws.Cells("J" & idxStr).Value) + TIMS.VAL1(ws.Cells("F" & idxStr).Value)
                Dim iCLASSCNT As Double = TIMS.VAL1(ws.Cells("K" & idxStr).Value) + TIMS.VAL1(ws.Cells("H" & idxStr).Value)
                SetCellValue(ws, "L" & idxStr, iDEFGOVCOST)
                SetCellValue(ws, "M" & idxStr, iCLASSCNT)

                If Convert.ToString(dr1("GCODE33")) = "99" Then
                    SetCellValue(ws, "Q" & idxStr, "-")
                    SetCellValue(ws, "R" & idxStr, "-")
                ElseIf Convert.ToString(dr1("GCODE33")) = "06" Then
                    SetCellValue(ws, "Q" & idxStr, "-")
                    SetCellValue(ws, "R" & idxStr, "-")
                Else
                    'L/O/Q
                    Dim i_A_COST As Double = TIMS.VAL1(ws.Cells("L" & idxStr).Value)
                    Dim i_B_COST As Double = TIMS.VAL1(ws.Cells("O" & idxStr).Value)
                    If i_A_COST > 0 AndAlso i_B_COST > 0 Then
                        SetCellValue(ws, "Q" & idxStr, String.Concat(TIMS.ROUND(i_A_COST / i_B_COST * 100, 0), "%"))
                    End If
                    'M/P/R
                    Dim i_C_CNT As Double = TIMS.VAL1(ws.Cells("M" & idxStr).Value)
                    Dim i_D_CNT As Double = TIMS.VAL1(ws.Cells("P" & idxStr).Value)
                    If i_C_CNT > 0 AndAlso i_D_CNT > 0 Then
                        SetCellValue(ws, "R" & idxStr, String.Concat(TIMS.ROUND(i_C_CNT / i_D_CNT * 100, 0), "%"))
                    End If
                End If

                idxStr += 1
            Next

            i_CLASSCNTG1 = 0
            idxStr = 6
            For Each dr3 As DataRow In dtXls3.Rows
                If (Convert.ToString(dr3("GCODE33")) = "99" AndAlso Convert.ToString(dr3("DISTID")) = "001") Then
                    i_CLASSCNTG1 += TIMS.VAL1(dr3("CLASSCNTG1"))
                    SetCellValue(ws, cellsCOLSPNumF5(0), dr3("CLASSCNTG1"))
                ElseIf (Convert.ToString(dr3("GCODE33")) = "99" AndAlso Convert.ToString(dr3("DISTID")) = "003") Then
                    i_CLASSCNTG1 += TIMS.VAL1(dr3("CLASSCNTG1"))
                    SetCellValue(ws, cellsCOLSPNumF5(1), dr3("CLASSCNTG1"))
                ElseIf (Convert.ToString(dr3("GCODE33")) = "99" AndAlso Convert.ToString(dr3("DISTID")) = "004") Then
                    i_CLASSCNTG1 += TIMS.VAL1(dr3("CLASSCNTG1"))
                    SetCellValue(ws, cellsCOLSPNumF5(2), dr3("CLASSCNTG1"))
                ElseIf (Convert.ToString(dr3("GCODE33")) = "99" AndAlso Convert.ToString(dr3("DISTID")) = "005") Then
                    i_CLASSCNTG1 += TIMS.VAL1(dr3("CLASSCNTG1"))
                    SetCellValue(ws, cellsCOLSPNumF5(3), dr3("CLASSCNTG1"))
                ElseIf (Convert.ToString(dr3("GCODE33")) = "99" AndAlso Convert.ToString(dr3("DISTID")) = "006") Then
                    i_CLASSCNTG1 += TIMS.VAL1(dr3("CLASSCNTG1"))
                    SetCellValue(ws, cellsCOLSPNumF5(4), dr3("CLASSCNTG1"))
                ElseIf (Convert.ToString(dr3("GCODE33")) = "99" AndAlso Convert.ToString(dr3("DISTID")) = "999") Then
                    SetCellValue(ws, cellsCOLSPNumF5(5), i_CLASSCNTG1)
                    Dim i_C_CNT As Double = TIMS.VAL1(ws.Cells("M" & idxStr).Value)
                    Dim i_E_CNT As Double = i_CLASSCNTG1
                    If i_C_CNT > 0 AndAlso i_E_CNT > 0 Then
                        SetCellValue(ws, "N" & idxStr, String.Concat(TIMS.ROUND(i_C_CNT / i_E_CNT * 100, 2), "%"))
                    End If
                End If

                If (Convert.ToString(dr3("GCODE33")) = "99" AndAlso Convert.ToString(dr3("DISTID")) <> "999") Then
                    Dim i_C_CNT As Double = TIMS.VAL1(ws.Cells("M" & idxStr).Value)
                    Dim i_E_CNT As Double = TIMS.VAL1(dr3("CLASSCNTG1"))
                    If i_C_CNT > 0 AndAlso i_E_CNT > 0 Then
                        SetCellValue(ws, "N" & idxStr, String.Concat(TIMS.ROUND(i_C_CNT / i_E_CNT * 100, 2), "%"))
                    End If
                End If

                idxStr += 1
            Next

            '刪除Temp中的資料 
            Call TIMS.MyFileDelete(filePath1)
            Dim V_ExpType As String = TIMS.GetListValue(RBListExpType)
            Select Case V_ExpType
                Case "EXCEL"
                    TIMS.ExpExcel_1(Me, strErrmsg, ep, Cst_FileSavePath, s_FILENAME1)
                    TIMS.Utl_RespWriteEnd(Me, objconn, "")
                Case "ODS"
                    TIMS.ExpODSl_1(Me, strErrmsg, ep, Cst_FileSavePath, s_FILENAME1)
                    TIMS.Utl_RespWriteEnd(Me, objconn, "")
                Case Else
                    Dim s_log1 As String = String.Format("ExpType(參數有誤)!!{0}", V_ExpType)
                    Common.MessageBox(Me, s_log1)
                    Return ' Exit Sub
            End Select
        End SyncLock

        '刪除Temp中的資料 'Call TIMS.MyFileDelete(myFileName1)
        If strErrmsg <> "" Then
            Common.MessageBox(Me, strErrmsg)
            Return
        End If
    End Sub

    ''' <summary>設定 Cell 儲存格值</summary>
    ''' <param name="sheet">Excel 工作表</param>
    ''' <param name="cellAddress">Cell 儲存格位址 (例如 A4、A1:L5)</param>
    ''' <param name="V_OBJ">Cell 儲存格值</param>
    ''' <param name="alignH">水平對齊方式</param>
    ''' <param name="alignV">垂直對齊方式</param>
    Private Sub SetCellValue(ByVal sheet As ExcelWorksheet, ByVal cellAddress As String, ByVal V_OBJ As Object, Optional ByVal alignH As ExcelHorizontalAlignment = ExcelHorizontalAlignment.Center, Optional ByVal alignV As ExcelVerticalAlignment = ExcelVerticalAlignment.Center)
        If sheet Is Nothing OrElse V_OBJ Is Nothing OrElse IsDBNull(V_OBJ) Then Return
        Dim nCells As ExcelRange = sheet.Cells(cellAddress)
        If nCells.Merge AndAlso cellAddress.IndexOf(":") > -1 Then
            sheet.Cells(cellAddress.Split(":")(0)).Value = V_OBJ
        Else
            nCells.Value = V_OBJ
        End If
        nCells.Style.HorizontalAlignment = alignH
        nCells.Style.VerticalAlignment = alignV
    End Sub

End Class
