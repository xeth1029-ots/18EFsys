Partial Class SD_11_006_R
    Inherits AuthBasePage

    Dim blnPrint2016 As Boolean = False
    'Const Cst_defQA2 As String = "A2" '預設問卷A2 'A'B 'select * from ID_Questionary 'select * from Plan_Questionary
    Const Cst_defQA16 As String = "A16" '20160501問卷A16 
    Const cst_prtFN1 As String = "SD_11_006_R_1" 'old(原)'班級
    Const cst_prtFNr3 As String = "SD_11_006_R3" '統計全轄區 RID2 (old)
    Const cst_prtFNr0 As String = "SD_11_006_R" '不統計全轄區 RIDValue.Value (old)

    Const cst_prtFN2 As String = "SD_11_006_R_2" '2016 'A16'班級
    Const cst_prtFNr4 As String = "SD_11_006_R4" '統計全轄區 RID2 (2016)
    Const cst_prtFNr1 As String = "SD_11_006_R1" '不統計全轄區 RIDValue.Value (2016)

#Region "(No Use)"

    'SELECT * FROM KEY_SURVEYKIND
    'SELECT * FROM ID_SurveyQuestion
    'WITH WC1 AS (select *　FROM VIEW_PLAN IP WHERE 1=1 and ip.Years='2016'  and ip.DistID IN ('001') and ip.TPlanID IN ('02'))
    ',WC2 AS (SELECT * FROM VIEW2 WHERE PLANID IN (SELECT PLANID FROM WC1))
    ',WC3 AS (SELECT * FROM class_studentsofclass  WHERE OCID IN (SELECT OCID FROM WC2))
    'SELECT * FROM Stud_Survey  WHERE SOCID IN (SELECT SOCID FROM WC3)
    'WITH WC1 AS (select *　FROM VIEW_PLAN IP WHERE 1=1 and ip.Years='2016'  and ip.DistID IN ('001') and ip.TPlanID IN ('02'))
    ',WC2 AS (SELECT * FROM VIEW2 WHERE PLANID IN (SELECT PLANID FROM WC1))
    ',WC3 AS (SELECT * FROM class_studentsofclass  WHERE OCID IN (SELECT OCID FROM WC2))
    'SELECT SKID,COUNT(1) FROM Stud_Survey  WHERE SOCID IN (SELECT SOCID FROM WC3) GROUP BY SKID
    'WITH WC1 AS (select *　FROM VIEW_PLAN IP WHERE 1=1 and ip.Years='2016'  and ip.DistID IN ('001') and ip.TPlanID IN ('02'))
    ',WC2 AS (SELECT * FROM VIEW2 WHERE PLANID IN (SELECT PLANID FROM WC1))
    ',WC3 AS (SELECT * FROM class_studentsofclass  WHERE OCID IN (SELECT OCID FROM WC2))
    'SELECT * FROM Stud_Survey  WHERE SOCID IN (SELECT SOCID FROM WC3) AND SKID='13'

    'Stud_Questionary
    '(VIEW_QUESTIONARY1 VIEW_QUESTIONARY2 VIEW_QUESTIONARY3)
    'ReportQuery : Member
    'SD_11_006_R_1 DataGrid1_ItemDataBound Member
    'SD_11_006_R3 '(列印) '統計全轄區 Member
    'SD_11_006_R '(列印) '不統計全轄區 Member

    '20160501
    'SD_11_006_R_2 DataGrid1_ItemDataBound Member
    'SD_11_006_R4 '(列印) '統計全轄區 Member
    'SD_11_006_R1 '(列印) '不統計全轄區 Member

    'SD_11_006_R_2'測試機mark  
    'SD_11_006_R2'測試機mark
    'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "SD_11_006_R3", MyValue)

#End Region

    'DataGrid1
    'Columns
    'Cells
    Const Cst_序號 As Integer = 0
    Const Cst_縣市別 As Integer = 1
    Const Cst_訓練單位 As Integer = 2
    Const Cst_班別名稱 As Integer = 3 'ClassCName/CLASSCNAME2
    Const Cst_期別 As Integer = 4 'CYCLTYPE

    Const Cst_開訓日期 As Integer = 5
    Const Cst_結訓日期 As Integer = 6
    Const Cst_結訓人數 As Integer = 7
    Const Cst_填寫人數 As Integer = 8

    Const Cst_第1部分平均滿意度 As Integer = 9
    Const Cst_第2部分平均滿意度 As Integer = 10
    Const Cst_第3部分平均滿意度 As Integer = 11
    Const Cst_第4部分平均滿意度 As Integer = 12
    Const Cst_第5部分平均滿意度 As Integer = 13
    'Const Cst_第6部分平均滿意度 As Integer = 13
    Const Cst_平均滿意度 As Integer = 14
    Const Cst_功能欄位 As Integer = 15

#Region "(No Use)"

    'Const Cst_CyclType As Integer = 16
    'Const Cst_OCID As Integer = 17
    'Const Cst_QID As Integer = 18

    'Dim count1 As Integer = 0
    'Dim count2 As Integer = 0
    'Dim count3 As Integer = 0
    'Dim count4 As Integer = 0
    'Dim count5 As Integer = 0
    'Dim count6 As Integer = 0

#End Region
#Region "Old Function"

    '取得各班平均滿意度 測試機mark
    'Function Get_SumAvg(ByVal sOCID As String, ByVal sSVID As String) As String
    '    Dim Sql As String = ""
    '    Sql = "" & vbCrLf
    '    sql &= " SELECT *  " & vbCrLf
    '    sql &= " FROM view_SurveyAvgScore " & vbCrLf
    '    sql &= " WHERE 1=1" & vbCrLf
    '    sql &= " AND OCID=" & sOCID & vbCrLf
    '    sql &= " AND SVID=" & sSVID & vbCrLf
    '    Dim SumAvg1 As String = "0"
    '    Dim dr As DataRow
    '    'Dim dt As DataTable
    '    dr = DbAccess.GetOneRow(Sql, objconn)
    '    If Not dr Is Nothing Then
    '        SumAvg1 = dr("AvgScore")
    '    End If
    '    Return SumAvg1
    'End Function

    ''期末學員滿意度調查檔 OCID 某班 (第1部份)。
    'Sub CreateItem_1(ByVal OCID As String, ByVal QID As String)
    '    'QID 問卷類別 0:不區分 1:原版 2:新版
    '    Dim dt As DataTable
    '    Dim sqlstr As String = ""
    '    sqlstr = "" & vbCrLf
    '    sqlstr += " select a.Q111 Q1_1_A" & vbCrLf
    '    sqlstr += " ,a.Q112 Q1_1_B" & vbCrLf
    '    sqlstr += " ,a.Q113 Q1_1_C" & vbCrLf
    '    sqlstr += " ,a.Q114 Q1_1_D" & vbCrLf
    '    sqlstr += " ,a.Q115 Q1_1_E" & vbCrLf
    '    sqlstr += " ,(a.Q111*5+a.Q112*4+a.Q113*3+a.Q114*2+a.Q115*1) Q1_1_SubTotal" & vbCrLf
    '    sqlstr += " ,TO_CHAR(CASE WHEN (a.Q111+a.Q112+a.Q113+a.Q114+a.Q115) >0 THEN ROUND((a.Q111*5+a.Q112*4+a.Q113*3+a.Q114*2+a.Q115*1) / (a.Q111+a.Q112+a.Q113+a.Q114+a.Q115) * 20,2) " & vbCrLf
    '    sqlstr += "  ELSE 0 END,990.99) SubAverage" & vbCrLf
    '    sqlstr += " ,'1' title" & vbCrLf
    '    sqlstr += " from VIEW_QUESTIONARY1 a" & vbCrLf
    '    sqlstr += " where 1=1" & vbCrLf
    '    sqlstr += " and a.OCID='" & OCID & "'" & vbCrLf
    '    sqlstr += " UNION ALL" & vbCrLf
    '    sqlstr += " select a.Q121 Q1_1_A" & vbCrLf
    '    sqlstr += " ,a.Q122 Q1_1_B" & vbCrLf
    '    sqlstr += " ,a.Q123 Q1_1_C" & vbCrLf
    '    sqlstr += " ,a.Q124 Q1_1_D" & vbCrLf
    '    sqlstr += " ,a.Q125 Q1_1_E" & vbCrLf
    '    sqlstr += " ,(a.Q121*5+a.Q122*4+a.Q123*3+a.Q124*2+a.Q125*1) Q1_1_SubTotal" & vbCrLf
    '    sqlstr += " ,TO_CHAR(CASE WHEN (a.Q121+a.Q122+a.Q123+a.Q124+a.Q125) >0 THEN ROUND((a.Q121*5+a.Q122*4+a.Q123*3+a.Q124*2+a.Q125*1) / (a.Q121+a.Q122+a.Q123+a.Q124+a.Q125) * 20,2) ELSE 0 END,990.99) SubAverage" & vbCrLf
    '    sqlstr += " ,'2' title" & vbCrLf
    '    sqlstr += " from VIEW_QUESTIONARY1 a" & vbCrLf
    '    sqlstr += " where 1=1" & vbCrLf
    '    sqlstr += " and a.OCID='" & OCID & "'" & vbCrLf
    '    sqlstr += " UNION ALL" & vbCrLf
    '    sqlstr += " select a.Q131 Q1_1_A" & vbCrLf
    '    sqlstr += " ,a.Q132 Q1_1_B" & vbCrLf
    '    sqlstr += " ,a.Q133 Q1_1_C" & vbCrLf
    '    sqlstr += " ,a.Q134 Q1_1_D" & vbCrLf
    '    sqlstr += " ,a.Q135 Q1_1_E" & vbCrLf
    '    sqlstr += " ,(a.Q131*5+a.Q132*4+a.Q133*3+a.Q134*2+a.Q135*1) Q1_1_SubTotal" & vbCrLf
    '    sqlstr += " ,TO_CHAR(CASE WHEN (a.Q131+a.Q132+a.Q133+a.Q134+a.Q135) >0 THEN ROUND((a.Q131*5+a.Q132*4+a.Q133*3+a.Q134*2+a.Q135*1) / (a.Q131+a.Q132+a.Q133+a.Q134+a.Q135) * 20,2) ELSE 0 END,990.99) SubAverage" & vbCrLf
    '    sqlstr += " ,'3' title" & vbCrLf
    '    sqlstr += " from VIEW_QUESTIONARY1 a" & vbCrLf
    '    sqlstr += " where 1=1" & vbCrLf
    '    sqlstr += " and a.OCID='" & OCID & "'" & vbCrLf

    '    Try
    '        dt = DbAccess.GetDataTable(sqlstr, objconn)
    '        If dt.Rows.Count > 0 Then
    '            DataGrid1_Detail_1.Visible = False
    '            DataGrid1_Detail_1.DataSource = dt
    '            DataGrid1_Detail_1.DataBind()
    '        End If
    '    Catch ex As Exception
    '        Common.MessageBox(Me, ex.ToString)
    '        Exit Sub
    '        'Throw ex
    '        'Common.RespWrite(Me, sqlstr)
    '    End Try
    'End Sub

    ''期末學員滿意度調查檔 OCID 某班 (第2部份)。
    'Sub CreateItem_2(ByVal OCID As String, ByVal QID As String)
    '    '" & OCID & "'
    '    'QID 問卷類別 0:不區分 1:原版 2:新版
    '    Dim dt As DataTable
    '    Dim sqlstr As String = ""
    '    sqlstr = "" & vbCrLf
    '    sqlstr += " select a.Q211 Q2_1_A" & vbCrLf
    '    sqlstr += " ,a.Q212 Q2_1_B" & vbCrLf
    '    sqlstr += " ,a.Q213 Q2_1_C" & vbCrLf
    '    sqlstr += " ,a.Q214 Q2_1_D" & vbCrLf
    '    sqlstr += " ,a.Q215 Q2_1_E" & vbCrLf
    '    sqlstr += " ,(a.Q211*5+a.Q212*4+a.Q213*3+a.Q214*2+a.Q215*1) Q2_1_SubTotal" & vbCrLf
    '    sqlstr += " ,TO_CHAR(CASE WHEN (a.Q211+a.Q212+a.Q213+a.Q214+a.Q215) >0 THEN ROUND((a.Q211*5+a.Q212*4+a.Q213*3+a.Q214*2+a.Q215*1) / (a.Q211+a.Q212+a.Q213+a.Q214+a.Q215) * 20,2) ELSE 0 END,990.99) SubAverage" & vbCrLf
    '    sqlstr += " ,'1' title" & vbCrLf
    '    sqlstr += " from VIEW_QUESTIONARY1 a" & vbCrLf
    '    sqlstr += " where 1=1" & vbCrLf
    '    sqlstr += " and a.OCID='" & OCID & "'" & vbCrLf
    '    sqlstr += " UNION ALL" & vbCrLf
    '    sqlstr += " select a.Q221 Q2_1_A" & vbCrLf
    '    sqlstr += " ,a.Q222 Q2_1_B" & vbCrLf
    '    sqlstr += " ,a.Q223 Q2_1_C" & vbCrLf
    '    sqlstr += " ,a.Q224 Q2_1_D" & vbCrLf
    '    sqlstr += " ,a.Q225 Q2_1_E" & vbCrLf
    '    sqlstr += " ,(a.Q221*5+a.Q222*4+a.Q223*3+a.Q224*2+a.Q225*1) Q2_1_SubTotal" & vbCrLf
    '    sqlstr += " ,TO_CHAR(CASE WHEN (a.Q221+a.Q222+a.Q223+a.Q224+a.Q225) >0 THEN ROUND((a.Q221*5+a.Q222*4+a.Q223*3+a.Q224*2+a.Q225*1) / (a.Q221+a.Q222+a.Q223+a.Q224+a.Q225) * 20,2) ELSE 0 END,990.99) SubAverage" & vbCrLf
    '    sqlstr += " ,'2' title" & vbCrLf
    '    sqlstr += " from VIEW_QUESTIONARY1 a" & vbCrLf
    '    sqlstr += " where 1=1" & vbCrLf
    '    sqlstr += " and a.OCID='" & OCID & "'" & vbCrLf
    '    sqlstr += " UNION ALL" & vbCrLf
    '    sqlstr += " select a.Q231 Q2_1_A" & vbCrLf
    '    sqlstr += " ,a.Q232 Q2_1_B" & vbCrLf
    '    sqlstr += " ,a.Q233 Q2_1_C" & vbCrLf
    '    sqlstr += " ,a.Q234 Q2_1_D" & vbCrLf
    '    sqlstr += " ,a.Q235 Q2_1_E" & vbCrLf
    '    sqlstr += " ,(a.Q231*5+a.Q232*4+a.Q233*3+a.Q234*2+a.Q235*1) Q2_1_SubTotal" & vbCrLf
    '    sqlstr += " ,TO_CHAR(CASE WHEN (a.Q231+a.Q232+a.Q233+a.Q234+a.Q235) >0 THEN ROUND((a.Q231*5+a.Q232*4+a.Q233*3+a.Q234*2+a.Q235*1) / (a.Q231+a.Q232+a.Q233+a.Q234+a.Q235) * 20,2) ELSE 0 END,990.99) SubAverage" & vbCrLf
    '    sqlstr += " ,'3' title" & vbCrLf
    '    sqlstr += " from VIEW_QUESTIONARY1 a" & vbCrLf
    '    sqlstr += " where 1=1" & vbCrLf
    '    sqlstr += " and a.OCID='" & OCID & "'" & vbCrLf
    '    sqlstr += " UNION ALL" & vbCrLf
    '    sqlstr += " select a.Q241 Q2_1_A" & vbCrLf
    '    sqlstr += " ,a.Q242 Q2_1_B" & vbCrLf
    '    sqlstr += " ,a.Q243 Q2_1_C" & vbCrLf
    '    sqlstr += " ,a.Q244 Q2_1_D" & vbCrLf
    '    sqlstr += " ,a.Q245 Q2_1_E" & vbCrLf
    '    sqlstr += " ,(a.Q241*5+a.Q242*4+a.Q243*3+a.Q244*2+a.Q245*1) Q2_1_SubTotal" & vbCrLf
    '    sqlstr += " ,TO_CHAR(CASE WHEN (a.Q241+a.Q242+a.Q243+a.Q244+a.Q245) >0 THEN ROUND((a.Q241*5+a.Q242*4+a.Q243*3+a.Q244*2+a.Q245*1) / (a.Q241+a.Q242+a.Q243+a.Q244+a.Q245) * 20,2) ELSE 0 END,990.99) SubAverage" & vbCrLf
    '    sqlstr += " ,'4' title" & vbCrLf
    '    sqlstr += " from VIEW_QUESTIONARY1 a" & vbCrLf
    '    sqlstr += " where 1=1" & vbCrLf
    '    sqlstr += " and a.OCID='" & OCID & "'" & vbCrLf
    '    sqlstr += " UNION ALL" & vbCrLf
    '    sqlstr += " select a.Q251 Q2_1_A" & vbCrLf
    '    sqlstr += " ,a.Q252 Q2_1_B" & vbCrLf
    '    sqlstr += " ,a.Q253 Q2_1_C" & vbCrLf
    '    sqlstr += " ,a.Q254 Q2_1_D" & vbCrLf
    '    sqlstr += " ,a.Q255 Q2_1_E" & vbCrLf
    '    sqlstr += " ,(a.Q251*5+a.Q252*4+a.Q253*3+a.Q254*2+a.Q255*1) Q2_1_SubTotal" & vbCrLf
    '    sqlstr += " ,TO_CHAR(CASE WHEN (a.Q251+a.Q252+a.Q253+a.Q254+a.Q255) >0 THEN ROUND((a.Q251*5+a.Q252*4+a.Q253*3+a.Q254*2+a.Q255*1) / (a.Q251+a.Q252+a.Q253+a.Q254+a.Q255) * 20,2) ELSE 0 END,990.99) SubAverage" & vbCrLf
    '    sqlstr += " ,'5' title" & vbCrLf
    '    sqlstr += " from VIEW_QUESTIONARY1 a" & vbCrLf
    '    sqlstr += " where 1=1" & vbCrLf
    '    sqlstr += " and a.OCID='" & OCID & "'" & vbCrLf

    '    Try
    '        dt = DbAccess.GetDataTable(sqlstr, objconn)
    '        If dt.Rows.Count > 0 Then
    '            DataGrid1_Detail_2.Visible = False
    '            DataGrid1_Detail_2.DataSource = dt
    '            DataGrid1_Detail_2.DataBind()
    '        End If
    '    Catch ex As Exception
    '        Common.MessageBox(Me, ex.ToString)
    '        Exit Sub
    '        'Throw ex
    '        'Common.RespWrite(Me, sqlstr)
    '    End Try
    'End Sub

    ''期末學員滿意度調查檔 OCID 某班 (第3部份)。
    'Sub CreateItem_3(ByVal OCID As String, ByVal QID As String)
    '    'Dim sqlstr As String
    '    Dim dt As DataTable
    '    Dim sqlstr As String = ""
    '    sqlstr = "" & vbCrLf
    '    sqlstr += " select a.Q311 Q3_1_A" & vbCrLf
    '    sqlstr += " ,a.Q312 Q3_1_B" & vbCrLf
    '    sqlstr += " ,a.Q313 Q3_1_C" & vbCrLf
    '    sqlstr += " ,a.Q314 Q3_1_D" & vbCrLf
    '    sqlstr += " ,a.Q315 Q3_1_E" & vbCrLf
    '    sqlstr += " ,(a.Q311*5+a.Q312*4+a.Q313*3+a.Q314*2+a.Q315*1) Q3_1_SubTotal" & vbCrLf
    '    sqlstr += " ,TO_CHAR(CASE WHEN (a.Q311+a.Q312+a.Q313+a.Q314+a.Q315) >0 THEN ROUND((a.Q311*5+a.Q312*4+a.Q313*3+a.Q314*2+a.Q315*1) / (a.Q311+a.Q312+a.Q313+a.Q314+a.Q315) * 20,2) ELSE 0 END,990.99) SubAverage" & vbCrLf
    '    sqlstr += " ,'1' title" & vbCrLf
    '    sqlstr += " from VIEW_QUESTIONARY1 a" & vbCrLf
    '    sqlstr += " where 1=1" & vbCrLf
    '    sqlstr += " and a.OCID='" & OCID & "'" & vbCrLf
    '    sqlstr += " UNION ALL" & vbCrLf
    '    sqlstr += " select a.Q321 Q3_1_A" & vbCrLf
    '    sqlstr += " ,a.Q322 Q3_1_B" & vbCrLf
    '    sqlstr += " ,a.Q323 Q3_1_C" & vbCrLf
    '    sqlstr += " ,a.Q324 Q3_1_D" & vbCrLf
    '    sqlstr += " ,a.Q325 Q3_1_E" & vbCrLf
    '    sqlstr += " ,(a.Q321*5+a.Q322*4+a.Q323*3+a.Q324*2+a.Q325*1) Q3_1_SubTotal" & vbCrLf
    '    sqlstr += " ,TO_CHAR(CASE WHEN (a.Q321+a.Q322+a.Q323+a.Q324+a.Q325) >0 THEN ROUND((a.Q321*5+a.Q322*4+a.Q323*3+a.Q324*2+a.Q325*1) / (a.Q321+a.Q322+a.Q323+a.Q324+a.Q325) * 20,2) ELSE 0 END,990.99) SubAverage" & vbCrLf
    '    sqlstr += " ,'2' title" & vbCrLf
    '    sqlstr += " from VIEW_QUESTIONARY1 a" & vbCrLf
    '    sqlstr += " where 1=1" & vbCrLf
    '    sqlstr += " and a.OCID='" & OCID & "'" & vbCrLf
    '    sqlstr += " UNION ALL" & vbCrLf
    '    sqlstr += " select a.Q331 Q3_1_A" & vbCrLf
    '    sqlstr += " ,a.Q332 Q3_1_B" & vbCrLf
    '    sqlstr += " ,a.Q333 Q3_1_C" & vbCrLf
    '    sqlstr += " ,a.Q334 Q3_1_D" & vbCrLf
    '    sqlstr += " ,a.Q335 Q3_1_E" & vbCrLf
    '    sqlstr += " ,(a.Q331*5+a.Q332*4+a.Q333*3+a.Q334*2+a.Q335*1) Q3_1_SubTotal" & vbCrLf
    '    sqlstr += " ,TO_CHAR(CASE WHEN (a.Q331+a.Q332+a.Q333+a.Q334+a.Q335) >0 THEN ROUND((a.Q331*5+a.Q332*4+a.Q333*3+a.Q334*2+a.Q335*1) / (a.Q331+a.Q332+a.Q333+a.Q334+a.Q335) * 20,2) ELSE 0 END,990.99) SubAverage" & vbCrLf
    '    sqlstr += " ,'3' title" & vbCrLf
    '    sqlstr += " from VIEW_QUESTIONARY1 a" & vbCrLf
    '    sqlstr += " where 1=1" & vbCrLf
    '    sqlstr += " and a.OCID='" & OCID & "'" & vbCrLf
    '    sqlstr += " UNION ALL" & vbCrLf
    '    sqlstr += " select a.Q341 Q3_1_A" & vbCrLf
    '    sqlstr += " ,a.Q342 Q3_1_B" & vbCrLf
    '    sqlstr += " ,a.Q343 Q3_1_C" & vbCrLf
    '    sqlstr += " ,a.Q344 Q3_1_D" & vbCrLf
    '    sqlstr += " ,a.Q345 Q3_1_E" & vbCrLf
    '    sqlstr += " ,(a.Q341*5+a.Q342*4+a.Q343*3+a.Q344*2+a.Q345*1) Q3_1_SubTotal" & vbCrLf
    '    sqlstr += " ,TO_CHAR(CASE WHEN (a.Q341+a.Q342+a.Q343+a.Q344+a.Q345) >0 THEN ROUND((a.Q341*5+a.Q342*4+a.Q343*3+a.Q344*2+a.Q345*1) / (a.Q341+a.Q342+a.Q343+a.Q344+a.Q345) * 20,2) ELSE 0 END,990.99) SubAverage" & vbCrLf
    '    sqlstr += " ,'4' title" & vbCrLf
    '    sqlstr += " from VIEW_QUESTIONARY1 a" & vbCrLf
    '    sqlstr += " where 1=1" & vbCrLf
    '    sqlstr += " and a.OCID='" & OCID & "'" & vbCrLf
    '    sqlstr += " UNION ALL" & vbCrLf
    '    sqlstr += " select a.Q351 Q3_1_A" & vbCrLf
    '    sqlstr += " ,a.Q352 Q3_1_B" & vbCrLf
    '    sqlstr += " ,a.Q353 Q3_1_C" & vbCrLf
    '    sqlstr += " ,a.Q354 Q3_1_D" & vbCrLf
    '    sqlstr += " ,a.Q355 Q3_1_E" & vbCrLf
    '    sqlstr += " ,(a.Q351*5+a.Q352*4+a.Q353*3+a.Q354*2+a.Q355*1) Q3_1_SubTotal" & vbCrLf
    '    sqlstr += " ,TO_CHAR(CASE WHEN (a.Q351+a.Q352+a.Q353+a.Q354+a.Q355) >0 THEN ROUND((a.Q351*5+a.Q352*4+a.Q353*3+a.Q354*2+a.Q355*1) / (a.Q351+a.Q352+a.Q353+a.Q354+a.Q355) * 20,2) ELSE 0 END,990.99) SubAverage" & vbCrLf
    '    sqlstr += " ,'5' title" & vbCrLf
    '    sqlstr += " from VIEW_QUESTIONARY1 a" & vbCrLf
    '    sqlstr += " where 1=1" & vbCrLf
    '    sqlstr += " and a.OCID='" & OCID & "'" & vbCrLf
    '    sqlstr += " UNION ALL" & vbCrLf
    '    sqlstr += " select a.Q361 Q3_1_A" & vbCrLf
    '    sqlstr += " ,a.Q362 Q3_1_B" & vbCrLf
    '    sqlstr += " ,a.Q363 Q3_1_C" & vbCrLf
    '    sqlstr += " ,a.Q364 Q3_1_D" & vbCrLf
    '    sqlstr += " ,a.Q365 Q3_1_E" & vbCrLf
    '    sqlstr += " ,(a.Q361*5+a.Q362*4+a.Q363*3+a.Q364*2+a.Q365*1) Q3_1_SubTotal" & vbCrLf
    '    sqlstr += " ,TO_CHAR(CASE WHEN (a.Q361+a.Q362+a.Q363+a.Q364+a.Q365) >0 THEN ROUND((a.Q361*5+a.Q362*4+a.Q363*3+a.Q364*2+a.Q365*1) / (a.Q361+a.Q362+a.Q363+a.Q364+a.Q365) * 20,2) ELSE 0 END,990.99) SubAverage" & vbCrLf
    '    sqlstr += " ,'6' title" & vbCrLf
    '    sqlstr += " from VIEW_QUESTIONARY1 a" & vbCrLf
    '    sqlstr += " where 1=1" & vbCrLf
    '    sqlstr += " and a.OCID='" & OCID & "'" & vbCrLf
    '    sqlstr += " UNION ALL" & vbCrLf
    '    sqlstr += " select a.Q371 Q3_1_A" & vbCrLf
    '    sqlstr += " ,a.Q372 Q3_1_B" & vbCrLf
    '    sqlstr += " ,a.Q373 Q3_1_C" & vbCrLf
    '    sqlstr += " ,a.Q374 Q3_1_D" & vbCrLf
    '    sqlstr += " ,a.Q375 Q3_1_E" & vbCrLf
    '    sqlstr += " ,(a.Q371*5+a.Q372*4+a.Q373*3+a.Q374*2+a.Q375*1) Q3_1_SubTotal" & vbCrLf
    '    sqlstr += " ,TO_CHAR(CASE WHEN (a.Q371+a.Q372+a.Q373+a.Q374+a.Q375) >0 THEN ROUND((a.Q371*5+a.Q372*4+a.Q373*3+a.Q374*2+a.Q375*1) / (a.Q371+a.Q372+a.Q373+a.Q374+a.Q375) * 20,2) ELSE 0 END,990.99) SubAverage" & vbCrLf
    '    sqlstr += " ,'7' title" & vbCrLf
    '    sqlstr += " from VIEW_QUESTIONARY1 a" & vbCrLf
    '    sqlstr += " where 1=1" & vbCrLf
    '    sqlstr += " and a.OCID='" & OCID & "'" & vbCrLf

    '    Try
    '        dt = DbAccess.GetDataTable(sqlstr, objconn)
    '        If dt.Rows.Count > 0 Then
    '            DataGrid1_Detail_3.Visible = False
    '            DataGrid1_Detail_3.DataSource = dt
    '            DataGrid1_Detail_3.DataBind()
    '        End If
    '    Catch ex As Exception
    '        Common.MessageBox(Me, ex.ToString)
    '        Exit Sub
    '        'Throw ex
    '        'Common.RespWrite(Me, sqlstr)
    '    End Try
    'End Sub

    ''期末學員滿意度調查檔 OCID 某班 (第4部份)。
    'Sub CreateItem_4(ByVal OCID As String, ByVal QID As String)
    '    'Dim sqlstr As String
    '    Dim dt As DataTable
    '    Dim sqlstr As String = ""
    '    sqlstr = "" & vbCrLf
    '    sqlstr += " select a.Q411 Q4_1_A" & vbCrLf
    '    sqlstr += " ,a.Q412 Q4_1_B" & vbCrLf
    '    sqlstr += " ,a.Q413 Q4_1_C" & vbCrLf
    '    sqlstr += " ,a.Q414 Q4_1_D" & vbCrLf
    '    sqlstr += " ,a.Q415 Q4_1_E" & vbCrLf
    '    sqlstr += " ,(a.Q411*5+a.Q412*4+a.Q413*3+a.Q414*2+a.Q415*1) Q4_1_SubTotal" & vbCrLf
    '    sqlstr += " ,TO_CHAR(CASE WHEN (a.Q411+a.Q412+a.Q413+a.Q414+a.Q415) >0 THEN ROUND((a.Q411*5+a.Q412*4+a.Q413*3+a.Q414*2+a.Q415*1) / (a.Q411+a.Q412+a.Q413+a.Q414+a.Q415) * 20,2) ELSE 0 END,990.99) SubAverage" & vbCrLf
    '    sqlstr += " ,'1' title" & vbCrLf
    '    sqlstr += " from VIEW_QUESTIONARY1 a" & vbCrLf
    '    sqlstr += " where 1=1" & vbCrLf
    '    sqlstr += " and a.OCID='" & OCID & "'" & vbCrLf
    '    sqlstr += " UNION ALL" & vbCrLf
    '    sqlstr += " select a.Q421 Q4_1_A" & vbCrLf
    '    sqlstr += " ,a.Q422 Q4_1_B" & vbCrLf
    '    sqlstr += " ,a.Q423 Q4_1_C" & vbCrLf
    '    sqlstr += " ,a.Q424 Q4_1_D" & vbCrLf
    '    sqlstr += " ,a.Q425 Q4_1_E" & vbCrLf
    '    sqlstr += " ,(a.Q421*5+a.Q422*4+a.Q423*3+a.Q424*2+a.Q425*1) Q4_1_SubTotal" & vbCrLf
    '    sqlstr += " ,TO_CHAR(CASE WHEN (a.Q421+a.Q422+a.Q423+a.Q424+a.Q425) >0 THEN ROUND((a.Q421*5+a.Q422*4+a.Q423*3+a.Q424*2+a.Q425*1) / (a.Q421+a.Q422+a.Q423+a.Q424+a.Q425) * 20,2) ELSE 0 END,990.99) SubAverage" & vbCrLf
    '    sqlstr += " ,'2' title" & vbCrLf
    '    sqlstr += " from VIEW_QUESTIONARY1 a" & vbCrLf
    '    sqlstr += " where 1=1" & vbCrLf
    '    sqlstr += " and a.OCID='" & OCID & "'" & vbCrLf
    '    sqlstr += " UNION ALL" & vbCrLf
    '    sqlstr += " select a.Q431 Q4_1_A" & vbCrLf
    '    sqlstr += " ,a.Q432 Q4_1_B" & vbCrLf
    '    sqlstr += " ,a.Q433 Q4_1_C" & vbCrLf
    '    sqlstr += " ,a.Q434 Q4_1_D" & vbCrLf
    '    sqlstr += " ,a.Q435 Q4_1_E" & vbCrLf
    '    sqlstr += " ,(a.Q431*5+a.Q432*4+a.Q433*3+a.Q434*2+a.Q435*1) Q4_1_SubTotal" & vbCrLf
    '    sqlstr += " ,TO_CHAR(CASE WHEN (a.Q431+a.Q432+a.Q433+a.Q434+a.Q435) >0 THEN ROUND((a.Q431*5+a.Q432*4+a.Q433*3+a.Q434*2+a.Q435*1) / (a.Q431+a.Q432+a.Q433+a.Q434+a.Q435) * 20,2) ELSE 0 END,990.99) SubAverage" & vbCrLf
    '    sqlstr += " ,'3' title" & vbCrLf
    '    sqlstr += " from VIEW_QUESTIONARY1 a" & vbCrLf
    '    sqlstr += " where 1=1" & vbCrLf
    '    sqlstr += " and a.OCID='" & OCID & "'" & vbCrLf
    '    sqlstr += " UNION ALL" & vbCrLf
    '    sqlstr += " select a.Q441 Q4_1_A" & vbCrLf
    '    sqlstr += " ,a.Q442 Q4_1_B" & vbCrLf
    '    sqlstr += " ,a.Q443 Q4_1_C" & vbCrLf
    '    sqlstr += " ,a.Q444 Q4_1_D" & vbCrLf
    '    sqlstr += " ,a.Q445 Q4_1_E" & vbCrLf
    '    sqlstr += " ,(a.Q441*5+a.Q442*4+a.Q443*3+a.Q444*2+a.Q445*1) Q4_1_SubTotal" & vbCrLf
    '    sqlstr += " ,TO_CHAR(CASE WHEN (a.Q441+a.Q442+a.Q443+a.Q444+a.Q445) >0 THEN ROUND((a.Q441*5+a.Q442*4+a.Q443*3+a.Q444*2+a.Q445*1) / (a.Q441+a.Q442+a.Q443+a.Q444+a.Q445) * 20,2) ELSE 0 END,990.99) SubAverage" & vbCrLf
    '    sqlstr += " ,'4' title" & vbCrLf
    '    sqlstr += " from VIEW_QUESTIONARY1 a" & vbCrLf
    '    sqlstr += " where 1=1" & vbCrLf
    '    sqlstr += " and a.OCID='" & OCID & "'" & vbCrLf
    '    sqlstr += " UNION ALL" & vbCrLf
    '    sqlstr += " select a.Q451 Q4_1_A" & vbCrLf
    '    sqlstr += " ,a.Q452 Q4_1_B" & vbCrLf
    '    sqlstr += " ,a.Q453 Q4_1_C" & vbCrLf
    '    sqlstr += " ,a.Q454 Q4_1_D" & vbCrLf
    '    sqlstr += " ,a.Q455 Q4_1_E" & vbCrLf
    '    sqlstr += " ,(a.Q451*5+a.Q452*4+a.Q453*3+a.Q454*2+a.Q455*1) Q4_1_SubTotal" & vbCrLf
    '    sqlstr += " ,TO_CHAR(CASE WHEN (a.Q451+a.Q452+a.Q453+a.Q454+a.Q455) >0 THEN ROUND((a.Q451*5+a.Q452*4+a.Q453*3+a.Q454*2+a.Q455*1) / (a.Q451+a.Q452+a.Q453+a.Q454+a.Q455) * 20,2) ELSE 0 END,990.99) SubAverage" & vbCrLf
    '    sqlstr += " ,'5' title" & vbCrLf
    '    sqlstr += " from VIEW_QUESTIONARY1 a" & vbCrLf
    '    sqlstr += " where 1=1" & vbCrLf
    '    sqlstr += " and a.OCID='" & OCID & "'" & vbCrLf
    '    sqlstr += " UNION ALL" & vbCrLf
    '    sqlstr += " select a.Q461 Q4_1_A" & vbCrLf
    '    sqlstr += " ,a.Q462 Q4_1_B" & vbCrLf
    '    sqlstr += " ,a.Q463 Q4_1_C" & vbCrLf
    '    sqlstr += " ,a.Q464 Q4_1_D" & vbCrLf
    '    sqlstr += " ,a.Q465 Q4_1_E" & vbCrLf
    '    sqlstr += " ,(a.Q461*5+a.Q462*4+a.Q463*3+a.Q464*2+a.Q465*1) Q4_1_SubTotal" & vbCrLf
    '    sqlstr += " ,TO_CHAR(CASE WHEN (a.Q461+a.Q462+a.Q463+a.Q464+a.Q465) >0 THEN ROUND((a.Q461*5+a.Q462*4+a.Q463*3+a.Q464*2+a.Q465*1) / (a.Q461+a.Q462+a.Q463+a.Q464+a.Q465) * 20,2) ELSE 0 END,990.99) SubAverage" & vbCrLf
    '    sqlstr += " ,'6' title" & vbCrLf
    '    sqlstr += " from VIEW_QUESTIONARY1 a" & vbCrLf
    '    sqlstr += " where 1=1" & vbCrLf
    '    sqlstr += " and a.OCID='" & OCID & "'" & vbCrLf

    '    Try
    '        dt = DbAccess.GetDataTable(sqlstr, objconn)
    '        If dt.Rows.Count > 0 Then
    '            DataGrid1_Detail_4.Visible = False
    '            DataGrid1_Detail_4.DataSource = dt
    '            DataGrid1_Detail_4.DataBind()
    '        End If
    '    Catch ex As Exception
    '        Common.MessageBox(Me, ex.ToString)
    '        Exit Sub
    '        'Throw ex
    '        'Common.RespWrite(Me, sqlstr)
    '    End Try
    'End Sub

    ''期末學員滿意度調查檔 OCID 某班 (第5部份)。
    'Sub CreateItem_5(ByVal OCID As String, ByVal QID As String)
    '    'Dim sqlstr As String = ""
    '    Dim dt As DataTable
    '    Dim sqlstr As String = ""
    '    sqlstr = "" & vbCrLf
    '    sqlstr += " select a.Q521 Q5_2_A" & vbCrLf
    '    sqlstr += " ,a.Q522 Q5_2_B" & vbCrLf
    '    sqlstr += " ,a.Q523 Q5_2_C" & vbCrLf
    '    sqlstr += " ,a.Q524 Q5_2_D" & vbCrLf
    '    sqlstr += " ,a.Q525 Q5_2_E" & vbCrLf
    '    sqlstr += " ,(a.Q521*5+a.Q522*4+a.Q523*3+a.Q524*2+a.Q525*1) Q5_2_SubTotal" & vbCrLf
    '    sqlstr += " ,TO_CHAR(CASE WHEN (a.Q521+a.Q522+a.Q523+a.Q524+a.Q525) >0 THEN ROUND((a.Q521*5+a.Q522*4+a.Q523*3+a.Q524*2+a.Q525*1) / (a.Q521+a.Q522+a.Q523+a.Q524+a.Q525) * 20,2) ELSE 0 END,990.99) SubAverage" & vbCrLf
    '    sqlstr += " ,'2' title" & vbCrLf
    '    sqlstr += " from VIEW_QUESTIONARY1 a" & vbCrLf
    '    sqlstr += " where 1=1" & vbCrLf
    '    sqlstr += " and a.OCID='" & OCID & "'" & vbCrLf
    '    sqlstr += " UNION ALL" & vbCrLf
    '    sqlstr += " select a.Q541 Q5_2_A" & vbCrLf
    '    sqlstr += " ,a.Q542 Q5_2_B" & vbCrLf
    '    sqlstr += " ,a.Q543 Q5_2_C" & vbCrLf
    '    sqlstr += " ,a.Q544 Q5_2_D" & vbCrLf
    '    sqlstr += " ,a.Q545 Q5_2_E" & vbCrLf
    '    sqlstr += " ,(a.Q541*5+a.Q542*4+a.Q543*3+a.Q544*2+a.Q545*1) Q5_2_SubTotal" & vbCrLf
    '    sqlstr += " ,TO_CHAR(CASE WHEN (a.Q541+a.Q542+a.Q543+a.Q544+a.Q545) >0 THEN ROUND((a.Q541*5+a.Q542*4+a.Q543*3+a.Q544*2+a.Q545*1) / (a.Q541+a.Q542+a.Q543+a.Q544+a.Q545) * 20,2) ELSE 0 END,990.99) SubAverage" & vbCrLf
    '    sqlstr += " ,'4' title" & vbCrLf
    '    sqlstr += " from VIEW_QUESTIONARY1 a" & vbCrLf
    '    sqlstr += " where 1=1" & vbCrLf
    '    sqlstr += " and a.OCID='" & OCID & "'" & vbCrLf

    '    Try
    '        dt = DbAccess.GetDataTable(sqlstr, objconn)
    '        If dt.Rows.Count > 0 Then
    '            DataGrid1_Detail_5.Visible = False
    '            DataGrid1_Detail_5.DataSource = dt
    '            DataGrid1_Detail_5.DataBind()
    '        End If
    '    Catch ex As Exception
    '        Common.MessageBox(Me, ex.ToString)
    '        Exit Sub
    '        'Throw ex
    '        'Common.RespWrite(Me, sqlstr)
    '    End Try
    'End Sub

    ''期末學員滿意度調查檔 OCID 某班 (第6部份)。
    'Sub CreateItem_6(ByVal OCID As String, ByVal QID As String)
    '    Dim dt As DataTable
    '    Dim sqlstr As String = ""
    '    sqlstr = "" & vbCrLf
    '    sqlstr += " select a.Q621 Q6_2_A" & vbCrLf
    '    sqlstr += " ,a.Q622 Q6_2_B" & vbCrLf
    '    sqlstr += " ,a.Q623 Q6_2_C" & vbCrLf
    '    sqlstr += " ,a.Q624 Q6_2_D" & vbCrLf
    '    sqlstr += " ,a.Q625 Q6_2_E" & vbCrLf
    '    sqlstr += " ,(a.Q621*5+a.Q622*4+a.Q623*3+a.Q624*2+a.Q625*1) Q6_2_SubTotal" & vbCrLf
    '    sqlstr += " ,TO_CHAR(CASE WHEN (a.Q621+a.Q622+a.Q623+a.Q624+a.Q625) >0 THEN ROUND((a.Q621*5+a.Q622*4+a.Q623*3+a.Q624*2+a.Q625*1) / (a.Q621+a.Q622+a.Q623+a.Q624+a.Q625) * 20,2) ELSE 0 END,990.99) SubAverage" & vbCrLf
    '    sqlstr += " ,'2' title" & vbCrLf
    '    sqlstr += " from VIEW_QUESTIONARY1 a" & vbCrLf
    '    sqlstr += " where 1=1" & vbCrLf
    '    sqlstr += " and a.OCID='" & OCID & "'" & vbCrLf
    '    sqlstr += " UNION ALL" & vbCrLf
    '    sqlstr += " select a.Q631 Q6_2_A" & vbCrLf
    '    sqlstr += " ,a.Q632 Q6_2_B" & vbCrLf
    '    sqlstr += " ,a.Q633 Q6_2_C" & vbCrLf
    '    sqlstr += " ,a.Q634 Q6_2_D" & vbCrLf
    '    sqlstr += " ,a.Q635 Q6_2_E" & vbCrLf
    '    sqlstr += " ,(a.Q631*5+a.Q632*4+a.Q633*3+a.Q634*2+a.Q635*1) Q6_2_SubTotal" & vbCrLf
    '    sqlstr += " ,TO_CHAR(CASE WHEN (a.Q631+a.Q632+a.Q633+a.Q634+a.Q635) >0 THEN ROUND((a.Q631*5+a.Q632*4+a.Q633*3+a.Q634*2+a.Q635*1) / (a.Q631+a.Q632+a.Q633+a.Q634+a.Q635) * 20,2) ELSE 0 END,990.99) SubAverage" & vbCrLf
    '    sqlstr += " ,'3' title" & vbCrLf
    '    sqlstr += " from VIEW_QUESTIONARY1 a" & vbCrLf
    '    sqlstr += " where 1=1" & vbCrLf
    '    sqlstr += " and a.OCID='" & OCID & "'" & vbCrLf
    '    sqlstr += " UNION ALL" & vbCrLf
    '    sqlstr += " select a.Q641 Q6_2_A" & vbCrLf
    '    sqlstr += " ,a.Q642 Q6_2_B" & vbCrLf
    '    sqlstr += " ,a.Q643 Q6_2_C" & vbCrLf
    '    sqlstr += " ,a.Q644 Q6_2_D" & vbCrLf
    '    sqlstr += " ,a.Q645 Q6_2_E" & vbCrLf
    '    sqlstr += " ,(a.Q641*5+a.Q642*4+a.Q643*3+a.Q644*2+a.Q645*1) Q6_2_SubTotal" & vbCrLf
    '    sqlstr += " ,TO_CHAR(CASE WHEN (a.Q641+a.Q642+a.Q643+a.Q644+a.Q645) >0 THEN ROUND((a.Q641*5+a.Q642*4+a.Q643*3+a.Q644*2+a.Q645*1) / (a.Q641+a.Q642+a.Q643+a.Q644+a.Q645) * 20,2) ELSE 0 END,990.99) SubAverage" & vbCrLf
    '    sqlstr += " ,'4' title" & vbCrLf
    '    sqlstr += " from VIEW_QUESTIONARY1 a" & vbCrLf
    '    sqlstr += " where 1=1" & vbCrLf
    '    sqlstr += " and a.OCID='" & OCID & "'" & vbCrLf
    '    sqlstr += " UNION ALL" & vbCrLf
    '    sqlstr += " select a.Q651 Q6_2_A" & vbCrLf
    '    sqlstr += " ,a.Q652 Q6_2_B" & vbCrLf
    '    sqlstr += " ,a.Q653 Q6_2_C" & vbCrLf
    '    sqlstr += " ,a.Q654 Q6_2_D" & vbCrLf
    '    sqlstr += " ,a.Q655 Q6_2_E" & vbCrLf
    '    sqlstr += " ,(a.Q651*5+a.Q652*4+a.Q653*3+a.Q654*2+a.Q655*1) Q6_2_SubTotal" & vbCrLf
    '    sqlstr += " ,TO_CHAR(CASE WHEN (a.Q651+a.Q652+a.Q653+a.Q654+a.Q655) >0 THEN ROUND((a.Q651*5+a.Q652*4+a.Q653*3+a.Q654*2+a.Q655*1) / (a.Q651+a.Q652+a.Q653+a.Q654+a.Q655) * 20,2) ELSE 0 END,990.99) SubAverage" & vbCrLf
    '    sqlstr += " ,'5' title" & vbCrLf
    '    sqlstr += " from VIEW_QUESTIONARY1 a" & vbCrLf
    '    sqlstr += " where 1=1" & vbCrLf
    '    sqlstr += " and a.OCID='" & OCID & "'" & vbCrLf

    '    Try
    '        dt = DbAccess.GetDataTable(sqlstr, objconn)
    '        If dt.Rows.Count > 0 Then
    '            DataGrid1_Detail_6.Visible = False
    '            DataGrid1_Detail_6.DataSource = dt
    '            DataGrid1_Detail_6.DataBind()
    '        End If
    '    Catch ex As Exception
    '        Common.MessageBox(Me, ex.ToString)
    '        Exit Sub
    '        'Throw ex
    '        'Common.RespWrite(Me, sqlstr)
    '    End Try
    'End Sub

#End Region
#Region "(No Use)"

    'Function Get_SumAvg(ByVal sOCID As String, ByVal sSVID As String) As String
    '    Dim Sql As String = ""
    '    Sql = "" & vbCrLf
    '    sql &= " SELECT " & vbCrLf
    '    sql &= " 	KSeq" & vbCrLf
    '    sql &= " 	,QSeq" & vbCrLf
    '    sql &= " 	,SUM (CASE WHEN ASeq=1 AND SubCnt=1 THEN 1 ELSE 0 END) as Q_1	" & vbCrLf
    '    sql &= " 	,SUM (CASE WHEN ASeq=2 AND SubCnt=1 THEN 1 ELSE 0 END) as Q_2	" & vbCrLf
    '    sql &= " 	,SUM (CASE WHEN ASeq=3 AND SubCnt=1 THEN 1 ELSE 0 END) as Q_3	" & vbCrLf
    '    sql &= " 	,SUM (CASE WHEN ASeq=4 AND SubCnt=1 THEN 1 ELSE 0 END) as Q_4	" & vbCrLf
    '    sql &= " 	,SUM (CASE WHEN ASeq=5 AND SubCnt=1 THEN 1 ELSE 0 END) as Q_5	" & vbCrLf
    '    sql &= " 	,SUM (SubScore) as SubScore_Total	" & vbCrLf
    '    sql &= " 	, to_char(CASE " & vbCrLf
    '    sql &= " 			WHEN SUM (CASE WHEN SubCnt=1 THEN 1 ELSE 0 END)>0 " & vbCrLf
    '    sql &= " 			THEN SUM (SubScore)/SUM (CASE WHEN SubCnt=1 THEN 1 ELSE 0 END) * 20" & vbCrLf
    '    sql &= " 			ELSE 0 END,9999990.99) as SubAvgScore" & vbCrLf
    '    sql &= " from (" & vbCrLf
    '    sql &= " " & vbCrLf
    '    sql &= " 	select " & vbCrLf
    '    sql &= " 		ss.OCID" & vbCrLf
    '    sql &= " 		,isq.SVID" & vbCrLf
    '    sql &= " 		,ksk.Serial as KSeq" & vbCrLf
    '    sql &= " 	  ,isq.Serial as QSeq" & vbCrLf
    '    sql &= " 		,isa.Serial as ASeq" & vbCrLf
    '    sql &= " 		,ss.SAID" & vbCrLf
    '    sql &= " 		,case " & vbCrLf
    '    sql &= " 			when ss.SAID is not null then 6-isa.Serial" & vbCrLf
    '    sql &= " 			else 0" & vbCrLf
    '    sql &= " 		end as SubScore " & vbCrLf
    '    sql &= " 		,case " & vbCrLf
    '    sql &= " 			when ss.SAID is not null then 1" & vbCrLf
    '    sql &= " 			else 0" & vbCrLf
    '    sql &= " 		end as SubCnt " & vbCrLf
    '    sql &= " 	from  " & vbCrLf
    '    sql &= " 	ID_SurveyQuestion isq " & vbCrLf
    '    sql &= " 	join Key_SurveyKind ksk on isq.SKID=ksk.SKID" & vbCrLf
    '    sql &= " 	join ID_SurveyAnswer isa on isq.SQID=isa.SQID" & vbCrLf
    '    sql &= " 	LEFT JOIN (" & vbCrLf
    '    sql &= " 		select a.SAID , vs.OCID" & vbCrLf
    '    sql &= " 		from Stud_Survey a" & vbCrLf
    '    sql &= " 		join view_StudSurvey vs on vs.SOCID =a.SOCID AND vs.SVID=a.SVID" & vbCrLf
    '    sql &= " 	) ss on ss.SAID=isa.SAID" & vbCrLf
    '    sql &= " 	where 1=1" & vbCrLf
    '    sql &= " 	AND ss.OCID=" & sOCID & vbCrLf
    '    sql &= " 	AND isq.SVID=" & sSVID & vbCrLf
    '    sql &= " " & vbCrLf
    '    sql &= " ) g" & vbCrLf
    '    sql &= " group by" & vbCrLf
    '    sql &= " 	KSeq" & vbCrLf
    '    sql &= "   ,QSeq" & vbCrLf
    '    sql &= " Order by" & vbCrLf
    '    sql &= " 	KSeq" & vbCrLf
    '    sql &= "   ,QSeq" & vbCrLf
    '    sql &= " " & vbCrLf
    '    Dim SumTotal As Integer = 0
    '    Dim SumAvg1 As String = "0"
    '    Dim dr As DataRow
    '    Dim dt As DataTable
    '    dt = DbAccess.GetDataTable(Sql)
    '    SumTotal = 0
    '    If dt.Rows.Count > 0 Then
    '        For i As Integer = 0 To dt.Rows.Count - 1
    '            dr = dt.Rows(i)
    '            SumTotal += dr("SubAvgScore")
    '        Next
    '        SumAvg1 = TIMS.Round(SumTotal / dt.Rows.Count, 2)
    '    End If
    '    Return SumAvg1
    'End Function

    'Private Sub CheckData_ServerChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckData.ServerChange
    '    If CheckData.Checked = True Then
    '        Class_TR.Visible = True
    '    Else
    '        center.Text = sm.UserInfo.OrgName
    '        RIDValue.Value = sm.UserInfo.RID
    '        PlanID.Value = sm.UserInfo.PlanID
    '        Class_TR.Visible = False
    '    End If
    'End Sub

#End Region

    Dim vsQName As String = ""
    Dim vsQID As String = ""
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

        '分頁設定 Start
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End

        If Not TIMS.GetQType(Me, vsQName, vsQID) Then
            '計畫未設定問卷類型, 請先設定後, 再使用匯入功能
            Common.MessageBox(Me, "計畫未設定問卷類型，請先設定後，再使用此功能")
            '離開此功能()
            Exit Sub
        End If

        If Not IsPostBack Then
            cCreate1()
        End If

        DistID.Attributes("onclick") = "ClearData();"

        Query.Attributes("OnClick") = "javascript:return chk()"

        '選擇全部轄區
        DistID.Attributes("onclick") = "SelectAll('DistID','DistHidden');"
        Tcitycode.Attributes("onclick") = "SelectAll('Tcitycode','TcityHidden');"

        '選擇全部訓練計畫
        'TPlanID.Attributes("onclick") = "SelectAll('TPlanID','TPlanHidden');"
        'chkTPlanID0.Attributes("onclick") = "SelectAll('chkTPlanID0','TPlanID0HID');"
        'chkTPlanID1.Attributes("onclick") = "SelectAll('chkTPlanID1','TPlanID1HID');"
        'chkTPlanIDX.Attributes("onclick") = "SelectAll('chkTPlanIDX','TPlanIDXHID');"

        If CheckData.Checked = True Then   '增加是否為單一機構查詢
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            PlanID.Value = sm.UserInfo.PlanID
            Button2.Disabled = True
            OCID.Items.Clear()
            Class_TR.Style("display") = "none"
            'Table4.Style("display") = "none"
            Org_TR.Style("display") = If(sm.UserInfo.LID = 0, "none", "")
        Else
            Class_TR.Style("display") = ""
            Org_TR.Style("display") = ""
            Button2.Disabled = False
            'Table4.Style("display") = "none"
        End If

        CheckData.Attributes("OnClick") = "Enabled_OCID('" & sm.UserInfo.OrgName & "','" & sm.UserInfo.RID & "','" & sm.UserInfo.PlanID & "');"

        If sm.UserInfo.DistID = "000" Then
            DistID.Enabled = True
        Else
            DistID.SelectedValue = sm.UserInfo.DistID
            DistID.Enabled = False
        End If

        DataGrid1_Detail_1.Visible = False
        DataGrid1_Detail_2.Visible = False
        DataGrid1_Detail_3.Visible = False
        DataGrid1_Detail_4.Visible = False
        DataGrid1_Detail_5.Visible = False
        'DataGrid1_Detail_6.Visible = False
        Button3.Style("display") = "none"

#Region "(No Use)"

        'Me.ViewState("SVID") = ""
        'If TIMS.Server_Path() = "DEMO" Then
        '    If sm.UserInfo.Years >= "2009" Then '測試機mark
        '        Me.ViewState("SVID") = TIMS.GetSVID(sm.UserInfo.TPlanID)
        '    End If
        'End If

#End Region

        '========== (依照承辦人需求,將部份欄位隱藏，by20180914)
        'If Not Page.IsPostBack Then
        '    If sm.UserInfo.TPlanID = "06" Then
        '        TPlanID0_TR.Style("display") = "none"  '隱藏[訓練計畫(職前)]區塊
        '        chkTPlanID1.Items.RemoveAt(chkTPlanID1.Items.IndexOf(chkTPlanID1.Items.FindByValue("28")))  '隱藏[訓練計畫(在職)]區塊-"產業人才投資方案"選項
        '        chkTPlanID1.Items.RemoveAt(chkTPlanID1.Items.IndexOf(chkTPlanID1.Items.FindByValue("54")))  '隱藏[訓練計畫(在職)]區塊-"充電起飛計畫(補助在職勞工及自營作業者參訓)"選項
        '        TPlanIDX_TR.Style("display") = "none"  '隱藏[訓練計畫(其他)]區塊
        '        TPlanID1_TR.Style("display") = "none"  '隱藏[訓練計畫(在職)]區塊，by:20180917
        '    End If
        'End If
        '=================================================
    End Sub
    Sub cCreate1()
        'If TIMS.sUtl_ChkTest() Then Common.SetListItem(rblprtType1, Cst_defQA16) '20160501統計表
        yearlist = TIMS.GetSyear(yearlist)
        Common.SetListItem(yearlist, sm.UserInfo.Years)
        yearlist.Items.Remove(yearlist.Items.FindByValue(""))
        DistID = TIMS.Get_DistID(DistID)
        If DistID.Items.FindByValue("") Is Nothing Then DistID.Items.Insert(0, New ListItem("全部", ""))
        Tcitycode = TIMS.Get_CityName(Tcitycode, TIMS.dtNothing)
        'Call TIMS.Get_TPlan2(chkTPlanID0, chkTPlanID1, chkTPlanIDX, objconn)
        'TPlanID = TIMS.Get_TPlan(TPlanID)
        'TPlanID.Items.Insert(0, New ListItem("全部", ""))
        'TPlanID.SelectedValue = sm.UserInfo.TPlanID
        'CheckData.Checked = True
        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID
        PlanID.Value = sm.UserInfo.PlanID
        OCID.Style("display") = "none"
        Print.Visible = False
        btnExport1.Visible = False
        PageControler1.Visible = False
        msg.Text = TIMS.cst_NODATAMsg11
        'Button3_Click(sender, e)
        Call sSearch3()

        If sm.UserInfo.LID = "2" Then   '2010/05/24 改成若是委訓單位登入下列欄位就不顯示
            Year_TR.Style("display") = "none"
            DistID_TR.Style("display") = "none"

            Check_TR.Style("display") = "none"
            Button2.Style("display") = "none"
        Else
            'LID: 0.1.
            Year_TR.Style("display") = ""
            DistID_TR.Style("display") = ""

            Check_TR.Style("display") = ""
            Button2.Style("display") = ""
        End If

    End Sub

    '配合SQL 語法的WHERE條件 
    Function GET_SQLWHERE2_C() As String
        'Dim whereSql As String
        'Dim rst As Boolean = True
        Dim rst As String = "" '
        'Const cst_errMsg1 As String = "只能選擇一個計畫!"

        'Dim TPlanID1 As String = ""
        'Dim N1 As Integer = 0 '只能選擇一個計畫
        'N1 = 0 '預設 N1 =0 表示沒有勾選計畫選項
        'N1 = TIMS.Get_TPlan2Val(chkTPlanID0, chkTPlanID1, chkTPlanIDX, 1)
        'If N1 >= 2 Then
        '    'msg2 += "只能選擇一個計畫!" & vbCrLf
        '    TPlanID1 = ""
        '    Common.MessageBox(Me, cst_errMsg1)
        '    Return "" 'False
        '    'Exit Function
        'End If

        'If N1 = 0 Then '如果計畫選項沒有選
        '    'msg2 += "請選擇計畫!" & vbCrLf
        '    'TPlanID1 = ""
        '    TPlanID1 = sm.UserInfo.TPlanID
        '    Common.SetListItem(chkTPlanID0, TPlanID1)
        '    Common.SetListItem(chkTPlanID1, TPlanID1)
        '    Common.SetListItem(chkTPlanIDX, TPlanID1)
        'End If
        'If N1 = 1 Then TPlanID1 = TIMS.Get_TPlan2Val(chkTPlanID0, chkTPlanID1, chkTPlanIDX, 2)

        'If msg2 <> "" Then
        '    Common.MessageBox(Me, msg2)
        '    Exit Sub
        'End If

        '辦訓地縣市
        Dim TCityCode2 As String = ""
        TCityCode2 = ""
        For i As Integer = 1 To Tcitycode.Items.Count - 1
            If Tcitycode.Items.Item(i).Selected AndAlso Tcitycode.Items.Item(i).Value <> "" Then
                'If Tcitycode.Items.Item(i).Text <> "全部" Then
                'End If
                If TCityCode2 <> "" Then TCityCode2 += ","
                TCityCode2 += Tcitycode.Items.Item(i).Value
            End If
        Next

        '選擇轄區
        Dim wDistID As String = ""
        wDistID = ""
        For Each objitem As ListItem In Me.DistID.Items
            If objitem.Selected AndAlso objitem.Value <> "" Then
                If wDistID <> "" Then wDistID &= ","
                wDistID &= "'" & objitem.Value.ToString & "'"
            End If
        Next

        '大計畫
        'Dim wTPlanID As String = TIMS.Get_TPlan2Val(chkTPlanID0, chkTPlanID1, chkTPlanIDX)

        '班級選擇
        Dim OCIDStr As String = ""
        For Each item As ListItem In OCID.Items
            If item.Selected AndAlso item.Value <> "" Then
                If item.Value = "%" Then
                    OCIDStr = ""
                    Exit For
                Else
                    If OCIDStr <> "" Then OCIDStr &= ","
                    OCIDStr &= item.Value
                End If
            End If
        Next

        Dim sql As String = ""
        sql = ""
        If Me.yearlist.SelectedValue <> "" Then '年度選擇
            sql &= " AND ip.Years = '" & Me.yearlist.SelectedValue & "' " & vbCrLf
        Else
            sql &= " AND ip.Years = '" & sm.UserInfo.Years & "' " & vbCrLf '限定登入計畫年度
        End If

        If wDistID <> "" Then sql &= " AND ip.DistID IN (" & wDistID & ") " & vbCrLf '轄區選擇

        'If TCityCode2 <> "" Then
        '    TCityCode2 = TIMS.ChgSQLINOR("iz.CTID ", TCityCode2)
        'End If
        If TCityCode2 <> "" Then sql &= " AND iz.CTID IN (" & TCityCode2 & ") " & vbCrLf '縣市

        'If wTPlanID <> "" Then sql &= " AND ip.TPlanID IN (" & wTPlanID & ") " & vbCrLf '大計畫
        sql &= String.Format(" AND ip.TPlanID='{0}'", sm.UserInfo.TPlanID) & vbCrLf '大計畫

        If OCIDStr <> "" Then
            sql &= " AND a.OCID IN (" & OCIDStr & ") " & vbCrLf
        Else
            '開訓區間
            If STDate1.Text <> "" Then
                sql &= " AND a.STDate >= " & TIMS.To_date(STDate1.Text) & vbCrLf
            End If
            If STDate2.Text <> "" Then
                sql &= " AND a.STDate <= " & TIMS.To_date(STDate2.Text) & vbCrLf 'convert(datetime, '" & STDate2.Text & "', 111)" & vbCrLf
            End If

            '結訓區間
            If FTDate1.Text <> "" Then
                sql &= " AND a.FTDate >= " & TIMS.To_date(FTDate1.Text) & vbCrLf 'convert(datetime, '" & FTDate1.Text & "', 111)" & vbCrLf
            End If
            If FTDate2.Text <> "" Then
                sql &= " AND a.FTDate <= " & TIMS.To_date(FTDate2.Text) & vbCrLf 'convert(datetime, '" & FTDate2.Text & "', 111)" & vbCrLf
            End If
        End If

        If CheckData.Checked = True Then '分署(中心)底下的全部機構
            If RIDValue.Value <> "" Then '機構選擇
                If RIDValue.Value <> "A" Then sql &= " AND a.RID LIKE '" & RIDValue.Value & "%' " & vbCrLf '非署(局)
            End If
        Else
            '單一機構查詢
            Select Case sm.UserInfo.TPlanID
                Case "17"  '補助地方政府訓練
                    If RIDValue.Value <> "" Then sql &= " AND a.RID LIKE '" & RIDValue.Value & "%'" & vbCrLf '機構選擇
                Case Else
                    '檢查RID長度為1的話就是用LIKE 
                    '機構選擇
                    If TIMS.Chk_RIDLEN(RIDValue.Value) = 1 Then
                        If RIDValue.Value <> "A" Then sql &= " AND a.RID LIKE '" & RIDValue.Value & "%'" & vbCrLf '非署(局)
                    Else
                        If RIDValue.Value <> "" Then sql &= " AND a.RID = '" & RIDValue.Value & "'" & vbCrLf
                    End If
            End Select
        End If
        rst = sql
        Return rst
    End Function

    '查詢 SQL
    Sub gSearch0()
        'Dim wheresql As String = "" '取得where條件
        'If Not GET_SQLWHERE2(wheresql) Then Exit Sub
        blnPrint2016 = False
        Select Case rblprtType1.SelectedValue
            Case Cst_defQA16
                blnPrint2016 = True
        End Select
        If blnPrint2016 Then
            '20160501統計表
            Call sSearch2()
            Exit Sub
        End If
        'If TIMS.sUtl_ChkTest Then
        '    '20160501統計表
        '    Call sSearch2(wheresql)
        '    Exit Sub
        'End If
        Call sSearch1() '原統計表
    End Sub

    '查詢 SQL (原)
    Sub sSearch1()
        Dim c_whereSql As String = GET_SQLWHERE2_C()
        If c_whereSql = "" Then Return
        'select * from ID_Questionary 
        'select * from Plan_Questionary
        If vsQID = "" Then vsQID = "3" '沒有填的話先使用3

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " WITH WC1 AS (" & vbCrLf
        sql &= " SELECT a.OCID ,a.CyclType ,a.ClassCName" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(a.ClassCName,a.CyclType) CLASSCNAME2" & vbCrLf
        sql &= " ,a.STDate ,a.FTDate ,a.PlanID ,e.OrgName" & vbCrLf
        sql &= " ,d.RID ,d.DistID" & vbCrLf
        sql &= " ,iz.CTName ,iz.CTID" & vbCrLf
        sql &= " FROM dbo.CLASS_CLASSINFO a" & vbCrLf
        sql &= " JOIN dbo.AUTH_RELSHIP d ON d.RID=a.RID" & vbCrLf
        sql &= " JOIN dbo.ORG_ORGINFO e ON e.Orgid=d.OrgID" & vbCrLf
        sql &= " JOIN dbo.VIEW_PLAN ip ON ip.PlanID=a.PlanID" & vbCrLf
        sql &= " JOIN dbo.VIEW_ZIPNAME iz ON iz.ZipCode=a.TaddressZip" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= c_whereSql
        'sql &= " AND ip.Years = '2018'" & vbCrLf
        'sql &= " AND ip.DistID IN ('001')" & vbCrLf
        'sql &= " AND iz.CTID IN (1,2)" & vbCrLf
        'sql &= " AND ip.TPlanID IN ('06')" & vbCrLf
        sql &= " )" & vbCrLf

        sql &= " SELECT a.OCID" & vbCrLf
        sql &= " ,a.CyclType" & vbCrLf
        sql &= " ,a.ClassCName" & vbCrLf
        sql &= " ,a.CLASSCNAME2" & vbCrLf
        sql &= " ,format(a.STDate,'yyyy/MM/dd') STDate" & vbCrLf
        sql &= " ,format(a.FTDate,'yyyy/MM/dd') FTDate" & vbCrLf
        sql &= " ,a.PlanID" & vbCrLf
        sql &= " ,a.OrgName" & vbCrLf
        sql &= " ,a.RID" & vbCrLf
        sql &= " ,a.DistID" & vbCrLf
        sql &= " ,a.CTName" & vbCrLf
        sql &= " ,ISNULL(b.QID,2) QID" & vbCrLf
        sql &= " ,b.total" & vbCrLf
        sql &= " ,ISNULL(b.num,0) num1" & vbCrLf
        sql &= " ,q4.Q1_AVERAGE Q1_AVERAGE" & vbCrLf
        sql &= " ,q4.Q2_AVERAGE Q2_AVERAGE" & vbCrLf
        sql &= " ,q4.Q3_AVERAGE Q3_AVERAGE" & vbCrLf
        sql &= " ,q4.Q4_AVERAGE Q4_AVERAGE" & vbCrLf
        sql &= " ,q4.Q5_AVERAGE Q5_AVERAGE" & vbCrLf
        'sql &= " ,q4.Q6_AVERAGE Q6_AVERAGE" & vbCrLf
        sql &= " ,q4.AVERAGE AVERAGE" & vbCrLf
        sql &= " FROM WC1 a" & vbCrLf　'CLASS_CLASSINFO

        sql &= " JOIN (" & vbCrLf
        sql &= " SELECT a.OCID ,ISNULL(q1.QID,2) QID" & vbCrLf
        sql &= " ,COUNT(1) TOTAL" & vbCrLf '班級人數
        sql &= " ,COUNT(CASE WHEN q1.STUDID IS NOT NULL THEN 1 END) NUM" & vbCrLf '填寫人數
        sql &= " FROM dbo.CLASS_STUDENTSOFCLASS cs" & vbCrLf
        sql &= " JOIN WC1 a ON a.OCID = cs.OCID" & vbCrLf
        sql &= " LEFT JOIN dbo.STUD_QUESTIONARY q1 ON q1.STUDID = CS.STUDENTID AND q1.OCID = CS.OCID" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND cs.StudStatus = 5" & vbCrLf
        sql &= " GROUP BY a.OCID,ISNULL(q1.QID,2)" & vbCrLf
        sql &= " ) b ON a.ocid = b.ocid" & vbCrLf

        sql &= " JOIN dbo.VIEW_QUESTIONARY4 q4 ON q4.ocid = a.ocid" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf

        sql &= " ORDER BY a.DistID, a.CTID, a.PlanID, a.RID, a.OCID, a.CyclType" & vbCrLf

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)

        'Table4.Visible = True
        Table4.Style("display") = ""
        DataGrid1.Visible = False
        Print.Visible = False
        btnExport1.Visible = False
        PageControler1.Visible = False

        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無資料")
            Exit Sub
        End If

        'Table4.Visible = True
        'Table4.Style("display") = "inline"
        DataGrid1.Visible = True
        Print.Visible = True
        btnExport1.Visible = True
        PageControler1.Visible = True

        'PageControler1.SqlString = sqlstr_class
        'PageControler1.SqlPrimaryKeyDataCreate(sqlstr_class, "OCID", "DistID,PlanID,OCID,CyclType")
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    '查詢 SQL (2016/05/01)
    Sub sSearch2()
        Dim c_whereSql As String = GET_SQLWHERE2_C()
        If c_whereSql = "" Then Return

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " WITH WC1 AS (" & vbCrLf
        sql &= " SELECT a.OCID ,a.CyclType ,a.ClassCName" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(a.ClassCName,a.CyclType) CLASSCNAME2" & vbCrLf
        sql &= " ,a.STDate ,a.FTDate ,a.PlanID ,e.OrgName" & vbCrLf
        sql &= " ,d.RID ,d.DistID" & vbCrLf
        sql &= " ,iz.CTName ,iz.CTID" & vbCrLf
        sql &= " FROM dbo.CLASS_CLASSINFO a" & vbCrLf
        sql &= " JOIN dbo.AUTH_RELSHIP d ON a.RID = d.RID" & vbCrLf
        sql &= " JOIN dbo.ORG_ORGINFO e ON d.OrgID = e.Orgid" & vbCrLf
        sql &= " JOIN dbo.VIEW_PLAN ip ON ip.PlanID = a.PlanID" & vbCrLf
        sql &= " JOIN dbo.VIEW_ZIPNAME iz ON a.TaddressZip = iz.ZipCode" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= c_whereSql
        'sql &= " AND ip.Years = '2018'" & vbCrLf
        'sql &= " AND ip.DistID IN ('001')" & vbCrLf
        'sql &= " AND iz.CTID IN (2,3)" & vbCrLf
        'sql &= " AND ip.TPlanID IN ('06')" & vbCrLf
        sql &= " )" & vbCrLf

        sql &= " SELECT a.OCID ,a.CyclType ,a.ClassCName" & vbCrLf
        sql &= " ,a.CLASSCNAME2" & vbCrLf
        sql &= " ,format(a.STDate,'yyyy/MM/dd') STDate" & vbCrLf
        sql &= " ,format(a.FTDate,'yyyy/MM/dd') FTDate" & vbCrLf
        sql &= " ,a.PlanID ,a.OrgName" & vbCrLf
        sql &= " ,a.RID ,a.DistID ,a.CTName ,NULL QID" & vbCrLf
        sql &= " ,ISNULL(s3.total,0) total" & vbCrLf
        sql &= " ,ISNULL(s3.num1,0) num1" & vbCrLf
        sql &= " ,q2.Q1_AVERAGE Q1_AVERAGE" & vbCrLf
        sql &= " ,q2.Q2_AVERAGE Q2_AVERAGE" & vbCrLf
        sql &= " ,q2.Q3_AVERAGE Q3_AVERAGE" & vbCrLf
        sql &= " ,q2.Q4_AVERAGE Q4_AVERAGE" & vbCrLf
        sql &= " ,q2.Q5_AVERAGE Q5_AVERAGE" & vbCrLf
        'sql &= " ,ISNULL(q2.Q6_AVERAGE,0) Q6_AVERAGE" & vbCrLf
        sql &= " ,ISNULL(q2.AVERAGE,0) AVERAGE" & vbCrLf
        sql &= " FROM WC1 a" & vbCrLf
        sql &= " JOIN (" & vbCrLf
        sql &= " 	SELECT cc.OCID" & vbCrLf
        sql &= " 	,COUNT(1) total" & vbCrLf
        sql &= " 	,COUNT(CASE WHEN dbo.FN_GET_SVID2(CS.SOCID,4) IS NOT NULL THEN 1 END) num1" & vbCrLf
        sql &= " 	FROM WC1 cc" & vbCrLf
        sql &= " 	JOIN dbo.CLASS_STUDENTSOFCLASS cs ON cs.OCID=cc.OCID" & vbCrLf
        sql &= " 	WHERE 1=1" & vbCrLf
        sql &= " 	AND cs.StudStatus = 5" & vbCrLf
        sql &= " 	GROUP BY cc.OCID" & vbCrLf
        sql &= " ) s3 ON s3.ocid = a.ocid" & vbCrLf
        sql &= " JOIN dbo.V_STUDQUESTION2 q2 ON q2.ocid = a.ocid" & vbCrLf
        sql &= " ORDER BY a.DistID, a.CTID, a.PlanID, a.RID, a.OCID, a.CyclType" & vbCrLf

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)

        'Table4.Visible = True
        Table4.Style("display") = ""
        DataGrid1.Visible = False
        Print.Visible = False
        btnExport1.Visible = False
        PageControler1.Visible = False

        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無資料")
            Exit Sub
        End If

        'Table4.Visible = True
        'Table4.Style("display") = "inline"
        DataGrid1.Visible = True
        Print.Visible = True
        btnExport1.Visible = True
        PageControler1.Visible = True

        'PageControler1.SqlString = sqlstr_class
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
        'PageControler1.SqlPrimaryKeyDataCreate(sqlstr_class, "OCID", "DistID,PlanID,OCID,CyclType")
    End Sub

    '查詢鈕
    Private Sub Query_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Query.Click
        Call gSearch0()
    End Sub

    Private Sub DataGrid1_ItemCommand(source As Object, e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Dim sCmdArg As String = e.CommandArgument
        If sCmdArg = "" Then Exit Sub

        Select Case e.CommandName
            Case "Detail"
                Select Case rblprtType1.SelectedValue
                    Case Cst_defQA16
                        'blnPrint2016 = True
                        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_prtFN2, sCmdArg)
                        Exit Sub
                End Select
                TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_prtFN1, sCmdArg)
                'Detail_Button.Attributes("onclick") =     ReportQuery.ReportScript(Me, "SD_11_006_R_1", MyValue)
        End Select
    End Sub

    'list 各各班級
    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                e.Item.Cells(Cst_序號).Text = TIMS.Get_DGSeqNo(sender, e) '序號
                Dim drv As DataRowView = e.Item.DataItem
                Dim Detail_Btn As Button = e.Item.FindControl("Detail") '列印明細
                'Dim OCID_Namestr As String
                'OCID_Namestr = drv("ClassCName").ToString
                'If CInt(drv("CyclType").ToString) <> 0 Then OCID_Namestr += "第" & TIMS.GetChtNum(CInt(drv("CyclType").ToString)) & "期"
                'e.Item.Cells(Cst_班別名稱).Text = Convert.ToString(drv("ClassCName2"))
                'e.Item.Cells(Cst_開訓日期).Text = Common.FormatDate(drv("STDate")) '開訓日期
                'e.Item.Cells(Cst_結訓日期).Text = Common.FormatDate(drv("FTDate")) '結訓日期
                'e.Item.Cells(Cst_平均滿意度).Text = "0"

                Dim sCmdArg As String = ""
                sCmdArg = ""
                sCmdArg &= "&Years=" & yearlist.SelectedValue
                sCmdArg &= "&OCID=" & Convert.ToString(drv("OCID"))
                sCmdArg &= "&WriteNum=" & Convert.ToString(drv("NUM1"))
                sCmdArg &= "&SumAvg=" & If(Convert.ToString(drv("AVERAGE")) <> "", Convert.ToString(drv("AVERAGE")), "0") 'e.Item.Cells(Cst_平均滿意度).Text
                '判斷有無填寫人數
                Detail_Btn.Enabled = If(Val(drv("num1")) > 0, True, False)
                '判斷有無填寫人數
                Detail_Btn.CommandArgument = If(Val(drv("num1")) > 0, sCmdArg, "")

                'select top 10  * from VIEW_QUESTIONARY4 where ocid =132776
        End Select

    End Sub

    '訓練機構的查詢
    Private Sub Button2_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.ServerClick

        'Dim i As Integer = 0
        'Dim N As Integer = 0 '勾選轄區選項 '只能選擇一個轄區!
        'Dim N1 As Integer = 0 '勾選計畫選項 '只能選擇一個計畫!
        Dim msg As String = ""
        Dim DistID1 As String = ""
        Dim iNN As Integer = 0   '預設 N =0 表示沒有勾選轄區選項
        For i As Integer = 1 To Me.DistID.Items.Count - 1
            If Me.DistID.Items(i).Selected Then '假如有勾選
                iNN = iNN + 1  '計算轄區勾選選項的數目
                If iNN = 1 Then '如果是勾選一個選項
                    DistID1 = Convert.ToString(Me.DistID.Items(i).Value) '取得選項的值
                End If
                If iNN = 2 Then '如果轄區勾選選項的數目=2
                    msg += "只能選擇一個轄區!" & vbCrLf
                    'Common.MessageBox(Me, "只能選擇一個轄區")
                    DistID1 = ""
                    Exit For
                End If
            End If
        Next
        If iNN = 0 Then '如果轄區選項沒有選
            msg += "請選擇轄區!" & vbCrLf
            'Common.MessageBox(Me, "請選擇轄區")
        End If
        If msg <> "" Then
            Common.MessageBox(Me, msg)
            Exit Sub
        End If

        'Dim TPlanID1 As String = ""
        'TPlanID1 = ""
        'N1 = 0 '預設 N1 =0 表示沒有勾選計畫選項
        'N1 = TIMS.Get_TPlan2Val(chkTPlanID0, chkTPlanID1, chkTPlanIDX, 1)
        'If N1 >= 2 Then
        '    msg += "只能選擇一個計畫!" & vbCrLf
        '    TPlanID1 = ""
        'End If
        'If N1 = 0 Then '如果計畫選項沒有選
        '    msg += "請選擇計畫!" & vbCrLf
        '    TPlanID1 = ""
        '    TPlanID1 = sm.UserInfo.TPlanID
        '    Common.SetListItem(chkTPlanID0, TPlanID1)
        '    Common.SetListItem(chkTPlanID1, TPlanID1)
        '    Common.SetListItem(chkTPlanIDX, TPlanID1)
        'End If
        'If N1 = 1 Then TPlanID1 = TIMS.Get_TPlan2Val(chkTPlanID0, chkTPlanID1, chkTPlanIDX, 2)

        Dim TPlanID1 As String = sm.UserInfo.TPlanID
        DistID1 = TIMS.ClearSQM(DistID1)
        TPlanID1 = TIMS.ClearSQM(TPlanID1)
        If DistID1 <> "" AndAlso TPlanID1 <> "" Then
            center.Text = ""
            OCID.Items.Clear()
            Dim strScript1 As String
            strScript1 = "<script language=""javascript"">" + vbCrLf
            strScript1 &= String.Format("wopen('../../Common/MainOrg.aspx?DistID={0}&TPlanID={1}&BtnName=Button3','查詢機構',400,400,1);", DistID1, TPlanID1)
            strScript1 &= "</script>"
            Page.RegisterStartupScript("", strScript1)
        End If
    End Sub

    '查詢班級
    Sub sSearch3()
        'Dim sql As String
        'Dim dt As DataTable
        'Dim dr As DataRow
        'Dim strSelected As String = ""

        PlanID.Value = TIMS.ClearSQM(PlanID.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        'SELECT * FROM AUTH_RELSHIP WHERE RID ='E1571'
        Dim Relship As String = TIMS.GET_RelshipforRID(RIDValue.Value, objconn)

        Dim parms As New Hashtable From {{"PlanID", PlanID.Value}, {"RID", RIDValue.Value}}
        Dim dt As DataTable = Nothing
        Dim sql As String = ""
        sql &= " SELECT cc.OCID " & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.ClassCName,cc.CyclType) CLASSCNAME2" & vbCrLf
        sql &= " FROM CLASS_CLASSINFO cc " & vbCrLf
        sql &= " JOIN ID_PLAN ip ON ip.planid = cc.planid " & vbCrLf
        sql &= " WHERE cc.NotOpen = 'N' " & vbCrLf
        sql &= " AND cc.IsSuccess = 'Y' " & vbCrLf
        sql &= " AND cc.PlanID = @PlanID" & vbCrLf '" & PlanID.Value & "' " & vbCrLf 
        sql &= " AND cc.RID =@RID" & vbCrLf '" & RIDValue.Value & "' " & vbCrLf
        sql &= " ORDER BY cc.OCID" & vbCrLf

        Try
            dt = DbAccess.GetDataTable(sql, objconn, parms)
        Catch ex As Exception
            Common.MessageBox(Me, ex.ToString)
            Exit Sub
            'Common.RespWrite(Me, sqlstr_class)
        End Try

        msg.Text = "查無此機構底下的班級"
        OCID.Style("display") = "none"
        If dt.Rows.Count = 0 Then Return

        Dim strSelected As String = ""
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            OCID.Style("display") = ""
            TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
            OCID.Items.Clear()
            OCID.Items.Add(New ListItem("全選", "%"))
            For Each dr As DataRow In dt.Rows
                OCID.Items.Add(New ListItem(dr("CLASSCNAME2"), dr("OCID")))
                If Convert.ToString(dr("OCID")) = OCIDValue1.Value Then strSelected = Convert.ToString(dr("OCID"))
            Next
            If strSelected.ToString <> "" Then OCID.SelectedValue = strSelected
        End If
    End Sub

    '查詢班級
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Call sSearch3()
    End Sub

#Region "hidden1"

    'Private Sub DataGrid1_Detail_1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1_Detail_1.ItemDataBound
    '    Dim dr As DataRowView
    '    Dim a As Integer = 0
    '    Dim a1, a2, a3, a4, a5, a6 As String
    '    dr = e.Item.DataItem

    '    If e.Item.ItemType = ListItemType.Footer Then
    '        For i As Integer = 1 To DataGrid1_Detail_1.Columns.Count - 1
    '            e.Item.Cells(i).Text = 0
    '            For Each Item As DataGridItem In DataGrid1_Detail_1.Items
    '                Select Case i
    '                    Case 1, 2, 3, 4, 5, 6
    '                        e.Item.Cells(i).Text = Int(e.Item.Cells(i).Text) + Int(Item.Cells(i).Text)
    '                End Select
    '            Next

    '            If i = 1 Then a1 = e.Item.Cells(1).Text
    '            If i = 2 Then a2 = e.Item.Cells(2).Text
    '            If i = 3 Then a3 = e.Item.Cells(3).Text
    '            If i = 4 Then a4 = e.Item.Cells(4).Text
    '            If i = 5 Then a5 = e.Item.Cells(5).Text
    '            If i = 6 Then a6 = e.Item.Cells(6).Text

    '        Next

    '        For j As Integer = 1 To 5
    '            a += Int(e.Item.Cells(j).Text)
    '        Next

    '        If a > 0 Then
    '            e.Item.Cells(7).Text = Math.Round(Convert.ToDouble(e.Item.Cells(6).Text) / a * 20, 2).ToString
    '            count1 = 1
    '        Else
    '            e.Item.Cells(7).Text = "0"
    '        End If
    '        Me.ViewState("Q1") = CDbl(e.Item.Cells(7).Text)
    '    End If
    'End Sub

    'Private Sub DataGrid1_Detail_2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1_Detail_2.ItemDataBound
    '    Dim dr As DataRowView
    '    Dim a As Integer = 0
    '    Dim a1, a2, a3, a4, a5, a6 As String
    '    'Dim count2 As Integer = 0
    '    dr = e.Item.DataItem

    '    If e.Item.ItemType = ListItemType.Footer Then
    '        For i As Integer = 1 To DataGrid1_Detail_2.Columns.Count - 1
    '            e.Item.Cells(i).Text = 0
    '            For Each Item As DataGridItem In DataGrid1_Detail_2.Items
    '                If (i = 1) Or (i = 2) Or (i = 3) Or (i = 4) Or (i = 5) Or (i = 6) Then
    '                    e.Item.Cells(i).Text = Int(e.Item.Cells(i).Text) + Int(Item.Cells(i).Text)
    '                End If
    '            Next

    '            If i = 1 Then a1 = e.Item.Cells(1).Text
    '            If i = 2 Then a2 = e.Item.Cells(2).Text
    '            If i = 3 Then a3 = e.Item.Cells(3).Text
    '            If i = 4 Then a4 = e.Item.Cells(4).Text
    '            If i = 5 Then a5 = e.Item.Cells(5).Text
    '            If i = 6 Then a6 = e.Item.Cells(6).Text

    '        Next

    '        For j As Integer = 1 To 5
    '            a += Int(e.Item.Cells(j).Text)
    '        Next

    '        If a > 0 Then
    '            e.Item.Cells(7).Text = Math.Round(Convert.ToDouble(e.Item.Cells(6).Text) / a * 20, 2).ToString
    '            count2 = 1
    '        Else
    '            e.Item.Cells(7).Text = "0"
    '        End If
    '        Me.ViewState("Q2") = CDbl(e.Item.Cells(7).Text)
    '    End If
    'End Sub

    'Private Sub DataGrid1_Detail_3_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1_Detail_3.ItemDataBound
    '    Dim dr As DataRowView
    '    Dim a As Integer = 0
    '    Dim a1, a2, a3, a4, a5, a6 As String
    '    'Dim count3 As Integer = 0
    '    dr = e.Item.DataItem

    '    If e.Item.ItemType = ListItemType.Footer Then
    '        For i As Integer = 1 To DataGrid1_Detail_3.Columns.Count - 1
    '            e.Item.Cells(i).Text = 0
    '            For Each Item As DataGridItem In DataGrid1_Detail_3.Items
    '                If (i = 1) Or (i = 2) Or (i = 3) Or (i = 4) Or (i = 5) Or (i = 6) Then
    '                    e.Item.Cells(i).Text = Int(e.Item.Cells(i).Text) + Int(Item.Cells(i).Text)
    '                End If
    '            Next

    '            If i = 1 Then a1 = e.Item.Cells(1).Text
    '            If i = 2 Then a2 = e.Item.Cells(2).Text
    '            If i = 3 Then a3 = e.Item.Cells(3).Text
    '            If i = 4 Then a4 = e.Item.Cells(4).Text
    '            If i = 5 Then a5 = e.Item.Cells(5).Text
    '            If i = 6 Then a6 = e.Item.Cells(6).Text

    '        Next

    '        For j As Integer = 1 To 5
    '            a += Int(e.Item.Cells(j).Text)
    '        Next

    '        If a > 0 Then
    '            e.Item.Cells(7).Text = Math.Round(Convert.ToDouble(e.Item.Cells(6).Text) / a * 20, 2).ToString
    '            count3 = 1
    '        Else
    '            e.Item.Cells(7).Text = "0"
    '        End If
    '        Me.ViewState("Q3") = CDbl(e.Item.Cells(7).Text)
    '    End If
    'End Sub

    'Private Sub DataGrid1_Detail_4_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1_Detail_4.ItemDataBound
    '    Dim dr As DataRowView
    '    Dim a As Integer = 0
    '    Dim a1, a2, a3, a4, a5, a6 As String
    '    'Dim count4 As Integer = 0
    '    dr = e.Item.DataItem

    '    If e.Item.ItemType = ListItemType.Footer Then
    '        For i As Integer = 1 To DataGrid1_Detail_4.Columns.Count - 1
    '            e.Item.Cells(i).Text = 0
    '            For Each Item As DataGridItem In DataGrid1_Detail_4.Items
    '                If (i = 1) Or (i = 2) Or (i = 3) Or (i = 4) Or (i = 5) Or (i = 6) Then
    '                    e.Item.Cells(i).Text = Int(e.Item.Cells(i).Text) + Int(Item.Cells(i).Text)
    '                End If
    '            Next

    '            If i = 1 Then a1 = e.Item.Cells(1).Text
    '            If i = 2 Then a2 = e.Item.Cells(2).Text
    '            If i = 3 Then a3 = e.Item.Cells(3).Text
    '            If i = 4 Then a4 = e.Item.Cells(4).Text
    '            If i = 5 Then a5 = e.Item.Cells(5).Text
    '            If i = 6 Then a6 = e.Item.Cells(6).Text

    '        Next

    '        For j As Integer = 1 To 5
    '            a += Int(e.Item.Cells(j).Text)
    '        Next

    '        If a > 0 Then
    '            e.Item.Cells(7).Text = Math.Round(Convert.ToDouble(e.Item.Cells(6).Text) / a * 20, 2).ToString
    '            count4 = 1
    '        Else
    '            e.Item.Cells(7).Text = "0"
    '        End If
    '        Me.ViewState("Q4") = CDbl(e.Item.Cells(7).Text)
    '    End If
    'End Sub

    'Private Sub DataGrid1_Detail_5_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1_Detail_5.ItemDataBound
    '    Dim dr As DataRowView
    '    Dim a As Integer = 0
    '    Dim a1, a2, a3, a4, a5, a6 As String
    '    'Dim count5 As Integer = 0
    '    dr = e.Item.DataItem

    '    If e.Item.ItemType = ListItemType.Footer Then
    '        For i As Integer = 1 To DataGrid1_Detail_5.Columns.Count - 1
    '            e.Item.Cells(i).Text = 0
    '            For Each Item As DataGridItem In DataGrid1_Detail_5.Items
    '                If (i = 1) Or (i = 2) Or (i = 3) Or (i = 4) Or (i = 5) Or (i = 6) Then
    '                    e.Item.Cells(i).Text = Int(e.Item.Cells(i).Text) + Int(Item.Cells(i).Text)
    '                End If
    '            Next

    '            If i = 1 Then a1 = e.Item.Cells(1).Text
    '            If i = 2 Then a2 = e.Item.Cells(2).Text
    '            If i = 3 Then a3 = e.Item.Cells(3).Text
    '            If i = 4 Then a4 = e.Item.Cells(4).Text
    '            If i = 5 Then a5 = e.Item.Cells(5).Text
    '            If i = 6 Then a6 = e.Item.Cells(6).Text

    '        Next

    '        For j As Integer = 1 To 5
    '            a += Int(e.Item.Cells(j).Text)
    '        Next

    '        If a > 0 Then
    '            e.Item.Cells(7).Text = Math.Round(Convert.ToDouble(e.Item.Cells(6).Text) / a * 20, 2).ToString
    '            count5 = 1
    '        Else
    '            e.Item.Cells(7).Text = "0"
    '        End If
    '        Me.ViewState("Q5") = CDbl(e.Item.Cells(7).Text)
    '    End If
    'End Sub

    'Private Sub DataGrid1_Detail_6_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1_Detail_6.ItemDataBound
    '    Dim dr As DataRowView
    '    Dim a As Integer = 0
    '    Dim a1, a2, a3, a4, a5, a6 As String
    '    'Dim count6 As Integer = 0
    '    dr = e.Item.DataItem

    '    If e.Item.ItemType = ListItemType.Footer Then
    '        For i As Integer = 1 To DataGrid1_Detail_6.Columns.Count - 1
    '            e.Item.Cells(i).Text = 0
    '            For Each Item As DataGridItem In DataGrid1_Detail_6.Items
    '                If (i = 1) Or (i = 2) Or (i = 3) Or (i = 4) Or (i = 5) Or (i = 6) Then
    '                    e.Item.Cells(i).Text = Int(e.Item.Cells(i).Text) + Int(Item.Cells(i).Text)
    '                End If
    '            Next

    '            If i = 1 Then a1 = e.Item.Cells(1).Text
    '            If i = 2 Then a2 = e.Item.Cells(2).Text
    '            If i = 3 Then a3 = e.Item.Cells(3).Text
    '            If i = 4 Then a4 = e.Item.Cells(4).Text
    '            If i = 5 Then a5 = e.Item.Cells(5).Text
    '            If i = 6 Then a6 = e.Item.Cells(6).Text
    '        Next

    '        For j As Integer = 1 To 5
    '            a += Int(e.Item.Cells(j).Text)
    '        Next

    '        If a > 0 Then
    '            e.Item.Cells(7).Text = Math.Round(Convert.ToDouble(e.Item.Cells(6).Text) / a * 20, 2).ToString
    '            count6 = 1
    '        Else
    '            e.Item.Cells(7).Text = "0"
    '        End If
    '        Me.ViewState("Q6") = CDbl(e.Item.Cells(7).Text)
    '    End If
    'End Sub

#End Region

    Private Sub Print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Print.Click
        '選擇轄區
        '報表要用的 轄區參數
        Dim DistID1 As String = ""
        For i As Integer = 1 To Me.DistID.Items.Count - 1
            If Me.DistID.Items(i).Selected Then
                If DistID1 <> "" Then DistID1 &= ","
                DistID1 &= Convert.ToString("\'" & Me.DistID.Items(i).Value & "\'")
            End If
        Next

        '報表要用的 辦訓地縣市
        Dim TCityCode2 As String = ""
        For i As Integer = 1 To Tcitycode.Items.Count - 1
            If Tcitycode.Items.Item(i).Selected = True Then
                If TCityCode2 <> "" Then TCityCode2 += ","
                TCityCode2 += Convert.ToString("\'" & Tcitycode.Items.Item(i).Value & "\'")
            End If
        Next

        '報表要用的訓練計畫參數
        'Dim TPlanID1 As String = TIMS.Get_TPlan2Val(chkTPlanID0, chkTPlanID1, chkTPlanIDX, 3)

        '報表要用的班級名稱參數
        Dim OCIDStr As String = ""
        'OCIDName = "部份"
        For Each item As ListItem In OCID.Items
            If item.Selected = True Then
                If item.Value = "%" Then
                    OCIDStr = ""
                    'OCIDName = "全部"
                    Exit For
                Else
                    If OCIDStr <> "" Then OCIDStr &= ","
                    OCIDStr &= item.Value
                    'If OCIDName <> "" Then OCIDName &= ","
                    'OCIDName &= item.Text
                End If
            End If
        Next

        'Me.ViewState("SVID") = ""
        'If TIMS.Server_Path() = "DEMO" Then '測試機
        '    If Me.yearlist.SelectedValue >= "2009" Then Me.ViewState("SVID") = TIMS.GetSVID(sm.UserInfo.TPlanID) '測試機mark
        'End If

        Dim RID2 As String = RIDValue.Value
        Dim PlanID2 As String = PlanID.Value
        '查詢範圍-統計全轄區
        If CheckData.Checked Then
            Select Case sm.UserInfo.LID
                Case 0
                    RID2 = ""
                    PlanID2 = ""
            End Select
        End If

        'If TPlanID1 = "" Then
        '    Common.MessageBox(Me, "請選擇訓練計畫!!")
        '    Exit Sub
        'End If

        Dim MyValue As String = ""
        MyValue = "k=r"
        MyValue += "&Years=" & Mid(Me.yearlist.SelectedValue, 3, 2)
        MyValue += "&SYears=" & Me.yearlist.SelectedValue
        MyValue += "&DistID=" & DistID1
        MyValue += "&CTID=" & TCityCode2
        MyValue += "&TPlanID=" & sm.UserInfo.TPlanID 'TPlanID1
        MyValue += "&OCID=" & OCIDStr
        MyValue += "&RID=" & RID2
        MyValue += "&STTDate=" & Me.STDate1.Text
        MyValue += "&FTTDate=" & Me.STDate2.Text
        MyValue += "&SFTDate=" & Me.FTDate1.Text
        MyValue += "&FFTDate=" & Me.FTDate2.Text
        MyValue += "&OCID1=" & OCIDStr
        MyValue += "&Planid=" & PlanID2
        'MyValue += "&DistName=" & Convert.ToString(DistName)
        'MyValue += "&TPlanName=" & Convert.ToString(TPlanName)
        'MyValue += "&OCIDName=" & Convert.ToString(OCIDName)
        MyValue += "&OrgName=" & Convert.ToString(center.Text)
        'If TCityCode2 <> "" Then MyValue += "&SORT1=iz.CTID"

        Select Case rblprtType1.SelectedValue
            Case Cst_defQA16
                'blnPrint2016 = True
                If CheckData.Checked = True Then
                    '統計全轄區 RID2
                    TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_prtFNr4, MyValue)
                Else
                    '不統計全轄區 RIDValue.Value
                    TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_prtFNr1, MyValue)
                End If
                Exit Sub
        End Select

        If CheckData.Checked = True Then
            '統計全轄區 RID2
            TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_prtFNr3, MyValue)
        Else
            '不統計全轄區 RIDValue.Value
            TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_prtFNr0, MyValue)
        End If
    End Sub

    '覆寫，不執行 MyBase.VerifyRenderingInServerForm 方法，解決執行 RenderControl 產生的錯誤   
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)
    End Sub

    Sub Expore1()
        DataGrid1.AllowPaging = False
        DataGrid1.EnableViewState = False  '把ViewState給關了
        Call gSearch0()

        'sFileName = HttpUtility.UrlEncode("滿意度調查統計表.xls", System.Text.Encoding.UTF8)

        'Response.Clear()
        'Response.Buffer = True
        'Response.Charset = "UTF-8" '設定字集
        'Response.AppendHeader("Content-Disposition", "attachment;filename=" & sFileName)
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("UTF-8")
        ''Response.ContentType = "application/vnd.ms-excel " '內容型態設為Excel
        ''文件內容指定為Excel
        'Response.ContentType = "application/ms-excel;charset=utf-8"
        'Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=UTF-8>")

        DataGrid1.AllowPaging = False
        DataGrid1.Columns(Cst_功能欄位).Visible = False
        DataGrid1.EnableViewState = False  '把ViewState給關了
        Dim objStringWriter As New System.IO.StringWriter
        Dim objHtmlTextWriter As New System.Web.UI.HtmlTextWriter(objStringWriter)
        Div1.RenderControl(objHtmlTextWriter)
        DataGrid1.AllowPaging = True
        DataGrid1.Columns(Cst_功能欄位).Visible = True

        Dim sFileName1 As String = "滿意度調查統計表"
        Dim strSTYLE As String = ""
        '套CSS值
        strSTYLE &= ("<style>")
        strSTYLE &= ("td{mso-number-format:""\@"";}")
        strSTYLE &= ("</style>")

        Dim strHTML As String = ""
        strHTML &= (TIMS.sUtl_AntiXss(Convert.ToString(objStringWriter)))

        Dim v_ExpType As String = TIMS.GetListValue(RBListExpType)
        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", v_ExpType) 'EXCEL/PDF/ODS
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", strHTML)
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)

        'Call TIMS.CloseDbConn(objconn)
        TIMS.Utl_RespWriteEnd(Me, objconn, "") '  Response.End()
    End Sub

    '匯出excel
    Private Sub btnExport1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExport1.Click
        Expore1()
    End Sub

End Class