Partial Class SD_05_019_R
    Inherits AuthBasePage

    'Dim oCmd As SqlCommand = Nothing
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        PageControler1.PageDataGrid = DataGrid1

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button10.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

        If Not IsPostBack Then
            msg.Text = ""
            'conn = DbAccess.GetConnection
            Button1.Attributes("onclick") = "return search();"
            Button2.Attributes.Add("onclick", "return CheckPrint();")
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            DistID.Value = sm.UserInfo.DistID
            TPlanID.Value = sm.UserInfo.TPlanID
            DataGridTable.Style.Item("display") = "none"
            If sm.UserInfo.LID <> "2" Then
                TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
            Else
                Button3_Click(sender, e)
            End If
        End If
    End Sub

    Sub sSearch1()
        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        Dim parms As New Hashtable From {{"OCID", OCIDValue1.Value}}
        Dim sql As String = ""
        sql &= " SELECT a.SOCID ,a.STUDENTID ,a.SID ,a.Name SNAME,a.IDNO ,a.SEX" & vbCrLf
        sql &= " ,FORMAT(a.BIRTHDAY,'yyyy/MM/dd') BIRTHDAY" & vbCrLf
        sql &= " ,a.SEX2 SEX2_N,a.STUDID2" & vbCrLf
        sql &= " FROM V_STUDENTINFO a" & vbCrLf
        sql &= " WHERE a.OCID =@OCID" & vbCrLf '132735--
        sql &= " ORDER BY a.StudentID" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)

        Me.ViewState("_SearchSqlStr") = sql
        msg.Text = "查無學生資料!"
        DataGridTable.Style.Item("display") = "none"
        If dt.Rows.Count = 0 Then Return

        Call getSubjectHr()

        msg.Text = ""
        DataGridTable.Style.Item("display") = ""
        'PageControler1.SqlString = sql
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()

    End Sub

    '查詢按鈕
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Call sSearch1()

    End Sub

    Private Sub getSubjectHr()  '取得學、術科總時數
        Dim thours As Int16 = 0     '各科時數加總
        Dim Hour1 As Int16 = 0      '學科總時數
        Dim Hour2 As Int16 = 0      '術科總時數
        Dim HrInSchedule As Int16 = 0       '所有按排在Class_Schedule當中的科目總時數  

        If (OCIDValue1.Value = "") Then Return

        'Dim dt3 As New DataTable
        Call TIMS.OpenDbConn(objconn)

        Dim sql As String = ""
        sql &= " SELECT DISTINCT b.CourseName ,a.CourID ,b.classification1 "
        sql &= " FROM STUD_TRAININGRESULTS a "
        sql &= " JOIN CLASS_STUDENTSOFCLASS cs ON cs.SOCID = a.SOCID AND cs.OCID =@OCID" '" & OCIDValue1.Value & "' "
        sql &= " LEFT JOIN COURSE_COURSEINFO b ON a.CourID = b.CourID "
        sql &= " ORDER BY b.courseName "
        Dim oCmd As New SqlCommand(sql, objconn)
        Dim dt As New DataTable
        With oCmd
            .Parameters.Clear()
            .Parameters.Add("OCID", SqlDbType.BigInt).Value = Val(OCIDValue1.Value)
            dt.Load(.ExecuteReader())
        End With

        'Call TIMS.OpenDbConn(objconn)
        Dim sql_2 As String = ""
        sql_2 &= " SELECT Class1, Class2, Class3, Class4, Class5, Class6, Class7, Class8, Class9, Class10, Class11, Class12 "
        sql_2 &= " FROM CLASS_SCHEDULE "
        sql_2 &= " WHERE OCID =@OCID" ' '" & OCIDValue1.Value & "' "
        Dim oCmd2 As New SqlCommand(sql_2, objconn)
        Dim dt2 As New DataTable
        With oCmd2
            .Parameters.Clear()
            .Parameters.Add("OCID", SqlDbType.Int).Value = Val(OCIDValue1.Value)
            dt2.Load(.ExecuteReader())
        End With

        'Call TIMS.OpenDbConn(objconn)
        Dim sql_3 As String = " SELECT THOURS FROM CLASS_CLASSINFO WHERE OCID =@OCID" '" & OCIDValue1.Value & "' "
        Dim oCmd3 As New SqlCommand(sql_3, objconn)
        Dim dt3 As New DataTable
        With oCmd3
            .Parameters.Clear()
            .Parameters.Add("OCID", SqlDbType.Int).Value = Val(OCIDValue1.Value)
            dt3.Load(.ExecuteReader())
            'thours = .ExecuteScalar()
        End With
        thours = If(dt3.Rows.Count > 0, dt3.Rows(0)("THOURS"), 0)


        If dt.Rows.Count > 0 Then
            For Each dr As DataRow In dt.Rows
                For Each dr2 As DataRow In dt2.Rows
                    Dim i As Int16 = 0
                    For i = 1 To 12
                        If Convert.ToString(dr("CourID")) = Convert.ToString(dr2("Class" & i.ToString)) Then
                            ' Case "1"  '學科 '    H1 = H1 + 1
                            Select Case Convert.ToString(dr("classification1"))
                                Case "2"  '術科
                                    Hour2 = Hour2 + 1
                            End Select
                            HrInSchedule = HrInSchedule + 1
                        End If
                    Next
                Next
            Next
        End If

        Hour1 = thours - Hour2  '學科總時數=全部時數加總-術科總時數(未歸類為學科或術科之科目，依原「受訓學員成績單」的報表邏輯未設定之科目就歸在學科)
        H1.Value = Hour1
        H2.Value = Hour2
        'type 學術科總時數計算方式
        type.Value = If(HrInSchedule > 0, 1, 2)
        'If HrInSchedule > 0 Then
        '    type.Value = 1  'type 學術科總時數計算方式
        'Else
        '    type.Value = 2
        'End If
    End Sub

#Region "(No Use)"
    'Private Function GetPara(ByVal GVID As String, ByVal DistID As String, ByVal TPlanID As String, Optional ByVal ItemVar As Int16 = 1) As Double  '取得參數設定值
    '    Dim rst As Double = 0
    '    Dim sql As String = ""
    '    Dim dt As DataTable = Nothing
    '    sql = " SELECT * FROM Sys_GlobalVar WHERE DistID = '" & DistID & "' AND TPlanID = '" & TPlanID & "' "
    '    dt = DbAccess.GetDataTable(sql, objconn)
    '    Select Case GVID
    '        Case "3"  '操行底分
    '            If dt.Select("GVID='3'").Length <> 0 Then rst = Val(dt.Select("GVID='3'")(0)("ItemVar1"))
    '        Case "13" '成績計算方式( 1, 各科平均法  2, 訓練時數權重法 )
    '            rst = 1
    '            If Me.ViewState("ResultTyp") = "" Then rst = Val(dt.Select("GVID='13'")(0)("ItemVar1"))    '預設取各科平均法
    '        Case "17" '學術科百分比 (學科ItemVar1 ex:0.4,術科ItemVar2  ex:0.4) 
    '            'rst = 0
    '            Select Case ItemVar
    '                Case 1 '學科百分比
    '                    rst = 0
    '                Case 2 '術科百分比
    '                    rst = 0
    '            End Select
    '            If dt.Select("GVID='17'").Length <> 0 Then         '預設各為50%
    '                Select Case ItemVar
    '                    Case 1 '學科百分比
    '                        rst = Math.Round(Convert.ToDouble(dt.Select("GVID='17'")(0)("ItemVar1")), 2)
    '                    Case 2 '術科百分比
    '                        rst = Math.Round(Convert.ToDouble(dt.Select("GVID='17'")(0)("ItemVar2")), 2)
    '                End Select
    '            End If
    '    End Select
    '    Return rst
    'End Function

    ''建立成績檔案
    'Function CreateGradeTable(ByVal TPlanID As String, ByVal DistID As String, ByVal OCIDValue1 As String, ByVal SOCID As String) As DataTable
    '    Dim dtReult As New DataTable
    '    dtReult.Columns.Add(New DataColumn("SOCID"))
    '    dtReult.Columns.Add(New DataColumn("pClassCount"))
    '    dtReult.Columns.Add(New DataColumn("sClassCount"))
    '    dtReult.Columns.Add(New DataColumn("pTotal_1"))
    '    dtReult.Columns.Add(New DataColumn("sTotal_1"))
    '    dtReult.Columns.Add(New DataColumn("pTotal_2"))
    '    dtReult.Columns.Add(New DataColumn("sTotal_2"))
    '    dtReult.Columns.Add(New DataColumn("pHours"))
    '    dtReult.Columns.Add(New DataColumn("sHours"))
    '    dtReult.Columns.Add(New DataColumn("totalResults"))
    '    dtReult.Columns.Add(New DataColumn("totalResults2"))
    '    dtReult.Columns.Add(New DataColumn("totalAvg"))
    '    dtReult.Columns.Add(New DataColumn("pAvg"))
    '    dtReult.Columns.Add(New DataColumn("sAvg"))
    '    dtReult.Columns.Add(New DataColumn("ResultTyp"))
    '    dtReult.Columns.Add(New DataColumn("percent1"))
    '    dtReult.Columns.Add(New DataColumn("percent2"))

    '    Dim Errmsg As String = ""
    '    'Dim i As Integer
    '    Dim sql As String = ""
    '    Dim sql3 As String = ""
    '    'Dim dt As New DataTable
    '    'Dim dt2 As New DataTable
    '    Dim CheckMode As String = ""
    '    'Dim conn As New SqlConnection
    '    'Dim da As New SqlDataAdapter

    '    Dim pClassCount As Integer = 0              '學科科目總數
    '    Dim sClassCount As Integer = 0              '術科科目總數

    '    sql = "SELECT * FROM Sys_GlobalVar WHERE DistID='" & DistID & "' and TPlanID='" & TPlanID & "'"
    '    oCmd = New SqlCommand(sql, objconn)
    '    'dt = DbAccess.GetDataTable(sql)

    '    sql3 = "  SELECT  CourID,CourseName " & vbCrLf                                                             '先取出目前所選取的科目  
    '    sql3 += " ,case  Classification1 when  1  then '學科'  when  2  then  '術科' end  as  classType ,Classification1 " & vbCrLf
    '    sql3 += " FROM  Course_CourseInfo  WHERE CourID in" & vbCrLf
    '    sql3 += " (" & Request("ChooseClass") & " ) " & vbCrLf
    '    sql3 += "  Order By CourseName "
    '    oCmd = New SqlCommand(sql3, objconn)

    '    Call TIMS.OpenDbConn(objconn)
    '    Dim dt As New DataTable
    '    With oCmd
    '        dt.Load(.ExecuteReader())
    '    End With

    '    Call TIMS.OpenDbConn(objconn)
    '    Dim dt3 As New DataTable
    '    With oCmd
    '        dt3.Load(.ExecuteReader())
    '    End With

    '    If dt.Select("GVID='13'").Length = 0 Then
    '        Errmsg += "尚未設定成績計算模式,請聯絡中心系統管理者" & vbCrLf
    '        'Exit Sub
    '    Else
    '        CheckMode = IIf(IsDBNull(dt.Select("GVID='13'")(0)("ItemVar1")), "1", dt.Select("GVID='13'")(0)("ItemVar1"))  '學科
    '    End If
    '    If dt3.Rows.Count > 0 Then
    '        Dim n As Int16 = 0
    '        For n = 0 To dt3.Rows.Count - 1
    '            If Convert.ToString(dt3.Rows(n)("Classification1")) = "1" Then
    '                pClassCount = pClassCount + 1
    '            ElseIf Convert.ToString(dt3.Rows(n)("Classification1")) = "2" Then
    '                sClassCount = sClassCount + 1
    '            End If
    '        Next
    '    End If

    '    Dim dt2 As New DataTable
    '    Dim ChooseClass As String = Request("ChooseClass")  'Get_ChooseClass(OCIDValue1)
    '    If ChooseClass <> "" Then
    '        'conn = DbAccess.GetConnection
    '        'conn.Open()
    '        '  dt2 = getStudResult(2, SOCID)
    '        '----------------------------------------------------------------  
    '        '***************************************
    '        ' 【計算成績方式】
    '        '        <各科平均法>
    '        '=======================================
    '        '  (1)學、術科百分比(無)  (2) 學、術科百分比(有)
    '        '    ex: 總平均是由(學科全部成績加總/學科總數 4  科) * 學科百分比 20 %      
    '        '        加上      (術科全部成績加總/術科總數 11 科) * 術科百分比 80 %      
    '        '        計算產生之結果。
    '        '---------------------------------------
    '        '        <訓練時數權重法> 
    '        '======================================= 
    '        '  (1)學、術科百分比(無)  
    '        '   總平均=(科)成績*(科)總時數 /(各科加總)總時數  
    '        '   總平均是由(各科成績與訓練時數加權/各科時數加總 __ 小時)計算產生之結果。 
    '        'ex:
    '        '總平均=
    '        '(學科成績與時數加權/時數加總 36 小時)* 學科百分比 40 %   
    '        '     +
    '        '(術科成績與時數加權/時數加總 88 小時)* 術科百分比 60 %        
    '        '/2 計算之結果。 

    '        '---------------------------------------
    '        '  (2) 學、術科百分比(有)
    '        '    ex: 學科平均=(科)成績*(科)總時數 /(各科加總)總時數  
    '        '        總平均是由(學科成績與訓練時數加權/學科時數 __ 小時) * 學科百分比 20 %      
    '        '        加上      (術科成績與訓練時數加權/術科時數 __ 小時) * 術科百分比 20 %        
    '        '        計算產生之結果。 
    '        '***************************************
    '        Dim j As Integer = 0
    '        'Dim pClassCount As Integer = 0              '學科科目總數
    '        'Dim sClassCount As Integer = 0              '術科科目總數
    '        Dim pTotal_1 As Decimal = 0        '學科總成績---各科平均法
    '        Dim sTotal_1 As Decimal = 0        '術科總成績 
    '        Dim pTotal_2 As Decimal = 0        '學科總成績---訓練時數權重法
    '        Dim sTotal_2 As Decimal = 0        '術科總成績
    '        Dim pHours As Decimal = 0          '學科總時數
    '        Dim sHours As Decimal = 0          '術科總時數
    '        Dim totalResults As Integer = 0    '總成績 
    '        Dim totalResults2 As Integer = 0   '總成績(時數加權) 
    '        Dim totalAvg As Decimal = 0        '總平均 
    '        Dim pAvg As Decimal = 0            '學科總平均 
    '        Dim sAvg As Decimal = 0            '術科總平均 
    '        Dim ResultTyp As Integer = GetPara(13, DistID, TPlanID)      '成績計算方式
    '        Dim percent1 As Double = GetPara(17, DistID, TPlanID, 1)     '學科百分比
    '        Dim percent2 As Double = GetPara(17, DistID, TPlanID, 2)     '術科百分比

    '        If dt2.Rows.Count > 0 Then
    '            For j = 0 To dt2.Rows.Count - 1
    '                If Convert.ToString(dt2.Rows(j)("Classification1")) = "1" Then        '學科
    '                    'pClassCount = pClassCount + 1
    '                    pHours = pHours + Convert.ToInt16(dt2.Rows(j)("hours"))
    '                    pTotal_2 = pTotal_2 + Convert.ToDecimal(dt2.Rows(j)("Results2"))  '訓練時數權重法： 2
    '                    pTotal_1 = pTotal_1 + Convert.ToDecimal(dt2.Rows(j)("Results"))   '各科平均法： 1 '預設採用各科平均法(當未設定時)
    '                ElseIf Convert.ToString(dt2.Rows(j)("Classification1")) = "2" Then    '術科
    '                    'sClassCount = sClassCount + 1
    '                    sHours = sHours + Convert.ToInt16(dt2.Rows(j)("hours"))
    '                    sTotal_2 = sTotal_2 + Convert.ToDecimal(dt2.Rows(j)("Results2"))  '訓練時數權重法： 2
    '                    sTotal_1 = sTotal_1 + Convert.ToDecimal(dt2.Rows(j)("Results"))
    '                End If
    '            Next

    '            If pClassCount = 0 Then
    '                pAvg = 0
    '            Else
    '                pAvg = pTotal_1 / pClassCount
    '            End If

    '            If sClassCount = 0 Then
    '                sAvg = 0
    '            Else
    '                sAvg = sTotal_1 / sClassCount
    '            End If

    '            totalResults = sTotal_1 + pTotal_1  '原始成績加總
    '            totalResults2 = sTotal_2 + pTotal_2 '成績加總(訓練時數權重法)

    '            'sql = " SELECT CourID ,Classification1   FROM  Course_CourseInfo  WHERE CourID in ( " & Request("ChooseClass") & ") "
    '            'conn = DbAccess.GetConnection
    '            'conn.Open()
    '            'With da
    '            '    .SelectCommand = New SqlCommand(sql, conn)
    '            '    .Fill(dt)
    '            'End With
    '            'If dt.Rows.Count > 0 Then
    '            '    Dim n As Int16 = 0
    '            '    Dim drClass As DataRow
    '            '    For n = 0 To dt.Rows.Count - 1
    '            '        drClass = dt.Rows(n)
    '            '        If Convert.ToString(drClass("Classification1")) = "1" Then
    '            '            pClassCount = pClassCount + 1
    '            '        ElseIf Convert.ToString(drClass("Classification1")) = "2" Then
    '            '            sClassCount = sClassCount + 1
    '            '        End If
    '            '    Next
    '            'End If

    '            Select Case ResultTyp
    '                Case 2        '訓練時數權重法： 2
    '                    '總平均= (學科成績與時數加權/時數加總 36 小時)* 學科百分比 40% + (術科成績與時數加權/時數加總 88 小時)* 術科百分比 60 %        
    '                    '/2 計算之結果。 
    '                    If percent1 = 0 And percent2 = 0 Then  '百分比未設定時
    '                        If pClassCount = 0 And sClassCount <> 0 And sHours <> 0 Then      '只有術科
    '                            totalAvg = (sTotal_2 / sHours)
    '                        ElseIf pClassCount <> 0 And sClassCount = 0 And pHours <> 0 Then  '只有學科
    '                            totalAvg = (pTotal_2 / pHours)
    '                        ElseIf pClassCount <> 0 And sClassCount <> 0 Then
    '                            If sHours = 0 And pHours <> 0 Then
    '                                totalAvg = (pTotal_2 / pHours)
    '                            ElseIf sHours <> 0 And pHours = 0 Then
    '                                totalAvg = (sTotal_2 / sHours)
    '                            ElseIf sHours <> 0 And pHours <> 0 Then
    '                                totalAvg = ((sTotal_2 / sHours) + (pTotal_2 / pHours)) / 2
    '                            End If
    '                        Else
    '                            totalAvg = 0
    '                        End If
    '                    Else
    '                        If pClassCount = 0 And sClassCount <> 0 And sHours <> 0 Then      '只有術科
    '                            totalAvg = (sTotal_2 / sHours) * percent2
    '                        ElseIf pClassCount <> 0 And sClassCount = 0 And pHours <> 0 Then  '只有學科
    '                            totalAvg = (pTotal_2 / pHours) * percent1
    '                        ElseIf pClassCount <> 0 And sClassCount <> 0 Then
    '                            If sHours = 0 And pHours <> 0 Then
    '                                totalAvg = (pTotal_2 / pHours) * percent1
    '                            ElseIf sHours <> 0 And pHours = 0 Then
    '                                totalAvg = (sTotal_2 / sHours) * percent2
    '                            ElseIf sHours <> 0 And pHours <> 0 Then
    '                                totalAvg = (pTotal_2 / pHours) * percent1 + (sTotal_2 / sHours) * percent2
    '                            Else
    '                                totalAvg = 0
    '                            End If
    '                        Else
    '                            totalAvg = 0
    '                        End If
    '                    End If
    '                Case Else     '各科平均法： 1 '預設採用各科平均法(當未設定時)
    '                    If percent1 = 0 And percent2 = 0 Then  '百分比未設定時
    '                        totalAvg = totalResults / (pClassCount + sClassCount)
    '                    Else
    '                        If pClassCount = 0 And sClassCount <> 0 Then      '只有術科
    '                            totalAvg = (sTotal_1 / sClassCount) * percent2
    '                        ElseIf pClassCount <> 0 And sClassCount = 0 Then  '只有學科
    '                            totalAvg = (pTotal_1 / pClassCount) * percent1
    '                        ElseIf pClassCount <> 0 And sClassCount <> 0 Then
    '                            totalAvg = (pTotal_1 / pClassCount) * percent1 + (sTotal_1 / sClassCount) * percent2
    '                        Else
    '                            totalAvg = 0
    '                        End If
    '                    End If
    '            End Select
    '            ' lbmsg.Text = ""

    '            Dim dr As DataRow = dtReult.NewRow
    '            dtReult.Rows.Add(dr)
    '            dr("SOCID") = SOCID
    '            dr("pClassCount") = pClassCount      '學科科目總數
    '            dr("sClassCount") = sClassCount      '術科科目總數
    '            dr("pTotal_1") = pTotal_1            '學科總成績---各科平均法
    '            dr("sTotal_1") = sTotal_1            '術科總成績 
    '            dr("pTotal_2") = pTotal_2            '學科總成績---訓練時數權重法 
    '            dr("sTotal_2") = sTotal_2            '術科總成績 
    '            dr("pHours") = pHours                '學科總時數
    '            dr("sHours") = sHours                '術科總時數
    '            dr("totalResults") = totalResults    '總成績 
    '            dr("totalResults2") = totalResults2  '總成績(時數加權) 
    '            dr("totalAvg") = TIMS.Round(totalAvg, 1)     '總平均 
    '            dr("pAvg") = TIMS.Round(pAvg, 1)             '學科總平均 
    '            dr("sAvg") = TIMS.Round(sAvg, 1)             '術科總平均 
    '            dr("ResultTyp") = ResultTyp   '成績計算方式 
    '            dr("percent1") = percent1     '學科百分比
    '            dr("percent2") = percent2     '術科百分比
    '            'Return dtReult
    '        End If
    '        'Else
    '        '    '  lbmsg.Text = "查無資料"
    '    End If
    '    '建立資料---------------   End
    '    Return dtReult
    'End Function

#End Region

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                Dim Checkbox3 As HtmlInputCheckBox = e.Item.FindControl("Checkbox3")
                'Checkbox3.Attributes("onclick") = "ChangeAll(this);"
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem

                Dim SOCID As HtmlInputHidden = e.Item.FindControl("SOCID")
                SOCID.Value = "'" & drv("SOCID") & "'"
        End Select
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn) '判斷機構是否只有一個班級
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        DataGridTable.Style.Item("display") = "none"
        '如果只有一個班級
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
        DataGridTable.Style.Item("display") = "none"

    End Sub
End Class