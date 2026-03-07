Partial Class TR_04_008_R
    Inherits AuthBasePage

    'ReportQuery
    'TR_04_008_R_3 @TR
    Const cst_printF1 As String = "TR_04_008_R_3"

    Dim vMSG As String = ""
    Dim vsYears2 As String = ""
    Const Cst_All As String = "%"
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        If Not IsPostBack Then
            Call Create1()
        End If

        Button2.Style("display") = "none" '列印
    End Sub

    Sub Create1()
        Dim sql As String
        '城市只要台灣的，不要大陸的
        sql = ""
        sql &= " SELECT CTID,CTName FROM ID_City WHERE CTID < 100"
        Select Case CStr(sm.UserInfo.LID)
            Case "0" '除署外其餘都要限定轄區
            Case Else
                sql &= " AND DISTID='" & sm.UserInfo.DistID & "'"
        End Select
        sql += " ORDER BY CTID" & vbCrLf
        Dim dtCity As DataTable = DbAccess.GetDataTable(sql, objconn)

        SCTID1 = TIMS.Get_CityName(SCTID1, dtCity)
        SCTID1.Items.Remove(SCTID1.Items.FindByValue(""))
        SCTID1.Items.Insert(0, New ListItem("全選", ""))
        SCTID1.Attributes("onclick") = "SelectAll('SCTID1','HidSCTID1');"

        SCTID2 = TIMS.Get_CityName(SCTID2, dtCity)
        SCTID2.Items.Remove(SCTID2.Items.FindByValue(""))
        SCTID2.Items.Insert(0, New ListItem("全選", ""))
        SCTID2.Attributes("onclick") = "SelectAll('SCTID2','HidSCTID2');"

        TIMS.Tooltip(SCTID1, "委託單位辦訓的縣市")
        TIMS.Tooltip(SCTID2, "訓練單位辦訓地址內的縣市")

        '署(局)選擇
        If sm.UserInfo.DistID = "000" Then
            DistID = TIMS.Get_DistID(DistID)
            TPlanID = TIMS.Get_TPlan(TPlanID, , 1)
        Else
            DistID = TIMS.Get_DistID(DistID)
            TPlanID = TIMS.Get_TPlan(TPlanID, , 1)
            DistID.Enabled = False '分署鎖定轄區
        End If

        'Select Case CStr(sm.UserInfo.LID)
        '    Case "2" '委訓單位限定機構
        '        center.Text = sm.UserInfo.OrgName
        '        RIDValue.Value = sm.UserInfo.RID
        'End Select

        'Syear = TIMS.GetSyear(Syear)
        'Common.SetListItem(Syear, Now.Year)
        DistID.SelectedValue = sm.UserInfo.DistID
        TPlanID.SelectedValue = sm.UserInfo.TPlanID

        OCIDList.Items.Clear()
        OCIDList.Items.Add("請選擇機構")
        Page.RegisterStartupScript("", "<script>GetMode();</script>")

        'OCID.Attributes("onchange") = "if(this.selectedIndex!=0){document.form1.OCIDValue.value=this.value;}else{document.form1.OCIDValue.value='';}"
        DistID.Attributes("onchange") = "return GetMode();"
        TPlanID.Attributes("onchange") = "return GetMode();"
        Button1.Attributes("onclick") = "javascript:return print();" '查詢
        btnExport1.Attributes("onclick") = "javascript:return print();" '匯出Excel
    End Sub

    '查詢 (選擇機構後)
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        OCIDList.Items.Clear()
        If RIDValue.Value = "" Then Exit Sub

        Dim dt As DataTable = Nothing
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " select a.ClassCName" & vbCrLf
        sql += " ,a.CyclType" & vbCrLf
        sql += " ,a.LevelType" & vbCrLf
        sql += " ,a.OCID" & vbCrLf
        sql += " from Class_ClassInfo a" & vbCrLf
        sql += " join ID_Plan b ON a.PlanID=b.PlanID" & vbCrLf
        sql += " LEFT join MVIEW_RELSHIP23 r3 ON r3.RID3 =a.RID" & vbCrLf
        sql += " WHERE 1=1" & vbCrLf
        '不列入就業率統計 join cc
        sql += " AND NOT EXISTS (SELECT 'x' FROM Class_NoWorkRate cnw WHERE cnw.OCID=a.OCID)" & vbCrLf
        If Len(RIDValue.Value) <> 1 Then
            '補助地方政府搜尋。
            sql += " AND (1!=1 " & vbCrLf
            sql += " OR a.RID='" & RIDValue.Value & "'" & vbCrLf
            sql += " OR r3.RID2='" & RIDValue.Value & "'" & vbCrLf
            sql += " )" & vbCrLf
        Else
            sql += " AND a.RID='" & sm.UserInfo.RID & "'" & vbCrLf
        End If
        sql += " and b.TPlanID='" & TPlanID.SelectedValue & "'" & vbCrLf '計畫
        sql += " and b.DistID='" & DistID.SelectedValue & "'" & vbCrLf '轄區
        sql += " and b.PlanID='" & PlanID.Value & "'" & vbCrLf '某個計畫序號

        dt = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count = 0 Then
            OCIDList.Items.Insert(0, New ListItem("此計畫、機構底下沒有任何班級", ""))
        Else
            For Each dr As DataRow In dt.Rows
                Dim ClassName As String = dr("ClassCName").ToString
                If Int(dr("CyclType")) <> 0 Then
                    ClassName += "第" & Int(dr("CyclType")) & "期"
                End If
                If Not IsDBNull(dr("LevelType")) Then
                    If Int(dr("LevelType")) <> 0 Then
                        ClassName += "第" & Int(dr("LevelType")) & "階段"
                    End If
                End If
                OCIDList.Items.Add(New ListItem(ClassName, dr("OCID")))
            Next
            OCIDList.Items.Insert(0, New ListItem("全選", Cst_All))
        End If
    End Sub

    '勾選的班級，組合成查詢參數
    Function Get_OCIDlistValue() As String
        'OCID.Items.Add(New ListItem("全選", "%"))
        'Const Cst_All As String = "%"
        Dim rstOCIDStr As String = ""

        For Each item As ListItem In OCIDList.Items
            If item.Selected = True Then
                If item.Value = Cst_All Then
                    'rstOCIDStr = Cst_All '全選
                    rstOCIDStr = ""
                    For Each item2 As ListItem In OCIDList.Items
                        If Not item2.Value = Cst_All Then
                            'item2.Selected = True
                            If IsNumeric(item2.Value) Then
                                If TIMS.IndexOf1(rstOCIDStr, item2.Value) = -1 Then
                                    If rstOCIDStr <> "" Then rstOCIDStr &= ","
                                    rstOCIDStr &= item2.Value
                                End If
                            End If
                        End If
                    Next
                    Exit For
                Else
                    If IsNumeric(item.Value) Then
                        If TIMS.IndexOf1(rstOCIDStr, item.Value) = -1 Then
                            If rstOCIDStr <> "" Then rstOCIDStr &= ","
                            rstOCIDStr &= item.Value
                        End If
                    End If
                End If
            End If
        Next
        Return rstOCIDStr
    End Function

    '檢查1
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        If CPoint.SelectedValue = "" Then
            Errmsg += "請選擇 就業查核點" & vbCrLf
        End If

        STDate1.Text = TIMS.ClearSQM(STDate1.Text)
        STDate2.Text = TIMS.ClearSQM(STDate2.Text)
        FTDate1.Text = TIMS.ClearSQM(FTDate1.Text)
        FTDate2.Text = TIMS.ClearSQM(FTDate2.Text)

        If STDate1.Text <> "" Then
            'STDate1.Text = Trim(STDate1.Text)
            If Not TIMS.IsDate1(STDate1.Text) Then
                Errmsg += "開訓期間 起始日期格式有誤" & vbCrLf
            End If
            If Errmsg = "" Then
                STDate1.Text = CDate(STDate1.Text).ToString("yyyy/MM/dd")
            End If
        End If

        If STDate2.Text <> "" Then
            'STDate2.Text = Trim(STDate2.Text)
            If Not TIMS.IsDate1(STDate2.Text) Then
                Errmsg += "開訓期間 迄止日期格式有誤" & vbCrLf
            End If
            If Errmsg = "" Then
                STDate2.Text = CDate(STDate2.Text).ToString("yyyy/MM/dd")
            End If
        End If

        If FTDate1.Text <> "" Then
            'FTDate1.Text = Trim(FTDate1.Text)
            If Not TIMS.IsDate1(FTDate1.Text) Then
                Errmsg += "結訓期間 起始日期格式有誤" & vbCrLf
            End If
            If Errmsg = "" Then
                FTDate1.Text = CDate(FTDate1.Text).ToString("yyyy/MM/dd")
            End If
        End If

        If FTDate2.Text <> "" Then
            'FTDate2.Text = Trim(FTDate2.Text)
            If Not TIMS.IsDate1(FTDate2.Text) Then
                Errmsg += "結訓期間 迄止日期格式有誤" & vbCrLf
            End If
            If Errmsg = "" Then
                FTDate2.Text = CDate(FTDate2.Text).ToString("yyyy/MM/dd")
            End If
        End If

        'Dim Years06 As Integer = 0
        'Dim Years07 As Integer = 0
        Dim STYear1 As Integer = 0
        Dim STYear2 As Integer = 0
        Dim FTYear1 As Integer = 0
        Dim FTYear2 As Integer = 0

        If Errmsg = "" Then
            If (Me.STDate1.Text <> "") Then STYear1 = Year(Me.STDate1.Text)
            If (Me.STDate2.Text <> "") Then STYear2 = Year(Me.STDate2.Text)
            If (Me.FTDate1.Text <> "") Then FTYear1 = Year(Me.FTDate1.Text)
            If (Me.FTDate2.Text <> "") Then FTYear2 = Year(Me.FTDate2.Text)

            '開結訓起日要同年度
            If (Me.STDate1.Text <> "") AndAlso (Me.FTDate1.Text <> "") Then
                If STYear1 <> FTYear1 Then Errmsg += "開結訓起日要同年度!" & vbCrLf
            End If

            'Select Case PrintStyle.SelectedValue
            '    Case "2006"
            '        Years06 += 1
            '        If (Me.STDate1.Text <> "") AndAlso (STYear1 > 2006) Then Errmsg += "開訓起日年度請設在西元2006年之前!" & vbCrLf
            '        If (Me.STDate2.Text <> "") AndAlso (STYear2 > 2006) Then Errmsg += "開訓迄日年度請設在西元2006年之前!" & vbCrLf
            '        If (Me.FTDate1.Text <> "") AndAlso (FTYear1 > 2006) Then Errmsg += "結訓起日年度請設在西元2006年之前!" & vbCrLf
            '        If (Me.FTDate2.Text <> "") AndAlso (FTYear2 > 2006) Then Errmsg += "結訓迄日年度請設在西元2006年之前!" & vbCrLf
            '    Case Else '2007
            '        Years07 += 1
            '        If (Me.STDate1.Text <> "") AndAlso (STYear1 < 2007) Then Errmsg += "開訓起日年度請設在西元2007年之後!" & vbCrLf
            '        If (Me.STDate2.Text <> "") AndAlso (STYear2 < 2007) Then Errmsg += "開訓迄日年度請設在西元2007年之後!" & vbCrLf
            '        If (Me.FTDate1.Text <> "") AndAlso (FTYear1 < 2007) Then Errmsg += "結訓起日年度請設在西元2007年之後!" & vbCrLf
            '        If (Me.FTDate2.Text <> "") AndAlso (FTYear2 < 2007) Then Errmsg += "結訓迄日年度請設在西元2007年之後!" & vbCrLf
            'End Select

            'If (Me.STDate1.Text <> "") AndAlso (STYear1 < 2007) Then Errmsg += "開訓起日年度請設在西元2007年之後!" & vbCrLf
            'If (Me.STDate2.Text <> "") AndAlso (STYear2 < 2007) Then Errmsg += "開訓迄日年度請設在西元2007年之後!" & vbCrLf
            'If (Me.FTDate1.Text <> "") AndAlso (FTYear1 < 2007) Then Errmsg += "結訓起日年度請設在西元2007年之後!" & vbCrLf
            'If (Me.FTDate2.Text <> "") AndAlso (FTYear2 < 2007) Then Errmsg += "結訓迄日年度請設在西元2007年之後!" & vbCrLf

            Const cst_ilastYear As Integer = 1999
            If (Me.STDate1.Text <> "") AndAlso (STYear1 < cst_ilastYear) Then Errmsg += "開訓起日年度請設在西元" & cst_ilastYear & "年之後!" & vbCrLf
            If (Me.STDate2.Text <> "") AndAlso (STYear2 < cst_ilastYear) Then Errmsg += "開訓迄日年度請設在西元" & cst_ilastYear & "年之後!" & vbCrLf
            If (Me.FTDate1.Text <> "") AndAlso (FTYear1 < cst_ilastYear) Then Errmsg += "結訓起日年度請設在西元" & cst_ilastYear & "年之後!" & vbCrLf
            If (Me.FTDate2.Text <> "") AndAlso (FTYear2 < cst_ilastYear) Then Errmsg += "結訓迄日年度請設在西元" & cst_ilastYear & "年之後!" & vbCrLf

        End If

        Dim sql As String = ""
        OCIDValue.Value = Get_OCIDlistValue()

        If OCIDValue.Value <> "" Then
            '青菜挑一個年度。
            sql = "" & vbCrLf
            sql += " select DISTINCT ip.Years " & vbCrLf
            sql += " from Class_ClassInfo cc" & vbCrLf
            sql += " join id_plan ip on ip.planid=cc.planid" & vbCrLf
            sql += " where 1=1 and cc.OCID in (" & OCIDValue.Value & ")" & vbCrLf
            sql += " order by 1 "
            Dim dtYears As DataTable
            dtYears = DbAccess.GetDataTable(sql, objconn)

            If dtYears.Rows.Count = 0 Then
                Errmsg += "班級查詢有誤!" & vbCrLf
                Return False 'Exit Function
            End If
            For i As Integer = 0 To dtYears.Rows.Count - 1
                If vsYears2 <> "" Then
                    If CInt(vsYears2) < CInt(dtYears.Rows(i)("Years")) Then
                        vsYears2 = Convert.ToString(dtYears.Rows(i)("Years"))
                    End If
                Else
                    vsYears2 = Convert.ToString(dtYears.Rows(i)("Years"))
                End If
                'If CInt(dtYears.Rows(i)("Years")) > 2006 Then
                '    If Years06 = 0 Then
                '        Years07 += 1
                '    Else
                '        Errmsg += "2007年之後的班別 含有2006年之前的班別(或格式)，報表格式請選擇新格式!" & vbCrLf
                '        Exit For
                '    End If
                'End If
                'If CInt(dtYears.Rows(i)("Years")) < 2007 Then
                '    If Years07 = 0 Then
                '        Years06 += 1
                '    Else
                '        Errmsg += "2006年之前的班別 含有2007年之後的班別(或格式)，報表格式請選擇舊格式!" & vbCrLf
                '        Exit For
                '    End If
                'End If
            Next
            'If dtYears.Rows.Count > 0 Then
            'Else
            '    Errmsg += "班級查詢有誤!" & vbCrLf
            'End If
            'Try
            'Catch ex As Exception
            '    Errmsg += "班級查詢有誤!" & vbCrLf
            'End Try
        Else
            '或不選
            Call TIMS.InputYears2(vsYears2, STYear1)
            Call TIMS.InputYears2(vsYears2, STYear2)
            Call TIMS.InputYears2(vsYears2, FTYear1)
            Call TIMS.InputYears2(vsYears2, FTYear2)
        End If

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    '列印
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Dim rtnPath As String = Request.FilePath
        Dim InputVal As String = ""
        'Me.txtSearchClassName.Text = "" & Trim(Me.txtSearchClassName.Text)
        txtSearchClassName.Text = TIMS.ClearSQM(txtSearchClassName.Text)

        'SQL Injection
        InputVal = Me.txtSearchClassName.Text
        If InputVal <> "" Then
            If TIMS.CheckInput(InputVal) Then
                Common.MessageBox(Me, TIMS.cst_ErrorMsg2, rtnPath)
                Exit Sub
            End If
        End If

        Dim stitle As String = ""
        Dim etitle As String = ""

        If OCIDValue.Value = "" Then
            If STDate1.Text <> "" OrElse STDate2.Text <> "" Then
                stitle = STDate1.Text + " ~ " + STDate2.Text
            End If
            If FTDate1.Text <> "" OrElse FTDate2.Text <> "" Then
                etitle = FTDate1.Text + " ~ " + FTDate2.Text
            End If
        Else
            STDate1.Text = ""
            STDate2.Text = ""
            FTDate1.Text = ""
            FTDate2.Text = ""
        End If


        Dim vSCTID1 As String = ""
        For Each item As ListItem In SCTID1.Items
            If item.Selected = True AndAlso item.Value <> "" Then
                If vSCTID1 <> "" Then vSCTID1 += ","
                vSCTID1 += item.Value
            End If
        Next
        Dim vSCTID2 As String = ""
        For Each item As ListItem In SCTID2.Items
            If item.Selected = True AndAlso item.Value <> "" Then
                If vSCTID2 <> "" Then vSCTID2 += ","
                vSCTID2 += item.Value
            End If
        Next
        'If vSCTID1 <> "" Then
        '    Sql &= " and iz1.CTID IN (" & vSCTID1 & ")" & vbCrLf
        'End If
        'If vSCTID2 <> "" Then
        '    Sql &= " and iz2.CTID IN (" & vSCTID2 & ")" & vbCrLf
        'End If

        Dim myValue As String = ""
        'myValue = "prg=TR_04_008_R"
        myValue = "&tid=xx"
        myValue &= "&DistID=" & Me.DistID.SelectedValue
        myValue &= "&TPlanID=" & Me.TPlanID.SelectedValue
        myValue &= "&RID=" & Me.RIDValue.Value
        myValue &= "&OCID=" & OCIDValue.Value
        myValue &= "&CPoint=" & CPoint.SelectedValue
        myValue &= "&STDate1=" & Me.STDate1.Text
        myValue &= "&STDate2=" & Me.STDate2.Text
        myValue &= "&FTDate1=" & Me.FTDate1.Text
        myValue &= "&FTDate2=" & Me.FTDate2.Text
        myValue &= "&SCTID1=" & vSCTID1
        myValue &= "&SCTID2=" & vSCTID2 '上課縣市
        myValue &= "&stitle=" & stitle
        myValue &= "&etitle=" & etitle
        myValue &= "&Year2=" & vsYears2
        'myValue &= "&SchClass=" & Convert.ToString(txtSearchClassName.Text).Replace("'", "\'\'")
        myValue &= "&SchClass=" & txtSearchClassName.Text

        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printF1, myValue)

        'ReportQuery
        'Select Case PrintStyle.SelectedValue
        '    Case "2006"
        '        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "TR", "TR_04_008_R", myValue)
        '    Case Else '2011
        '        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "TR", "TR_04_008_R_3", myValue)
        'End Select
        'Case Else '2007
        'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "TR", "TR_04_008_R_2", myValue)

    End Sub

    '查詢 [SQL] 匯出Excel檔
    Private Function Search1() As Boolean
        Dim rst As Boolean = False '是否有資料 ，預設沒有資料
        'Dim dt As DataTable = Nothing
        'Dim sda As SqlDataAdapter = Nothing
        txtSearchClassName.Text = TIMS.ClearSQM(txtSearchClassName.Text)

        Dim vSCTID1 As String = ""
        For Each item As ListItem In SCTID1.Items
            If item.Selected = True AndAlso item.Value <> "" Then
                If vSCTID1 <> "" Then vSCTID1 += ","
                vSCTID1 += item.Value
            End If
        Next
        Dim vSCTID2 As String = ""
        For Each item As ListItem In SCTID2.Items
            If item.Selected = True AndAlso item.Value <> "" Then
                If vSCTID2 <> "" Then vSCTID2 += ","
                vSCTID2 += item.Value
            End If
        Next


        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " WITH WC1 AS (" & vbCrLf
        sql &= " select oo.orgname" & vbCrLf
        sql &= " ,cc.classcname" & vbCrLf
        sql &= " ,cc.cycltype" & vbCrLf
        sql &= " ,cc.stdate" & vbCrLf
        sql &= " ,cc.ftdate" & vbCrLf
        sql &= " ,ip.distid" & vbCrLf
        sql &= " ,cc.comidno" & vbCrLf
        sql &= " ,cc.ocid" & vbCrLf
        'https://jira.turbotech.com.tw/browse/TIMSC-158
        sql &= " ,iz1.ctname ctnameN1" & vbCrLf '轄區縣市
        sql &= " ,iz2.ctname ctnameN2" & vbCrLf '上課縣市
        sql &= " from class_classinfo cc" & vbCrLf
        sql &= " join org_orginfo oo on oo.comidno=cc.comidno" & vbCrLf
        sql &= " join view_plan ip on ip.planid=cc.planid" & vbCrLf
        sql &= " join view_ridname vr on vr.RID =cc.RID" & vbCrLf
        sql &= " left join view_zipname iz1 on iz1.zipcode=vr.zipcode" & vbCrLf
        sql &= " left join view_zipname iz2 on iz2.zipcode=cc.taddresszip" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND cc.NotOpen = 'N'" & vbCrLf
        sql &= " and cc.FTDate < getdate()" & vbCrLf
        '不列入就業率統計 join cc
        sql &= " and NOT EXISTS (SELECT 'x' FROM Class_NoWorkRate cnw WHERE cnw.OCID=cc.OCID)" & vbCrLf
        If vSCTID1 <> "" Then
            sql &= " and iz1.CTID IN (" & vSCTID1 & ")" & vbCrLf
        End If
        If vSCTID2 <> "" Then
            sql &= " and iz2.CTID IN (" & vSCTID2 & ")" & vbCrLf
        End If
        If Me.DistID.SelectedValue <> "" Then
            sql &= "  and ip.DistID= @DistID" & vbCrLf
        End If
        If Me.RIDValue.Value <> "" Then
            sql &= " and cc.RID= @RID" & vbCrLf
        End If
        If Me.TPlanID.SelectedValue <> "" Then
            sql &= " and ip.TPlanID= @TPlanID" & vbCrLf
        End If
        If OCIDValue.Value <> "" Then
            If OCIDValue.Value.IndexOf(",") = -1 Then
                sql &= " and cc.OCID IN (@OCID)" & vbCrLf
            Else
                sql &= " and cc.OCID IN (" & OCIDValue.Value & ")" & vbCrLf
            End If
        Else
            If Me.STDate1.Text <> "" Then
                sql &= " and cc.STDate>= @STDate1" & vbCrLf
            End If
            If Me.STDate2.Text <> "" Then
                sql &= " and cc.STDate<= @STDate2" & vbCrLf
            End If
            If Me.FTDate1.Text <> "" Then
                sql &= " and cc.FTDate>= @FTDate1" & vbCrLf
            End If
            If Me.FTDate2.Text <> "" Then
                sql &= " and cc.FTDate<= @FTDate2" & vbCrLf
            End If
        End If
        If Convert.ToString(txtSearchClassName.Text) <> "" Then
            sql &= " and cc.classcname like '%'+@classcname+'%'" & vbCrLf
        End If
        'sql &= " and ip.DistID= '001'" & vbCrLf
        'sql &= " and ip.TPlanID= '02'" & vbCrLf
        'sql &= " and cc.STDate>= convert(datetime, '2016/03/01', 111)" & vbCrLf
        'sql &= " and cc.STDate<= convert(datetime, '2016/05/01', 111)" & vbCrLf
        sql &= " )" & vbCrLf

        sql &= " ,WS1 AS (" & vbCrLf
        sql &= " select cs.ocid" & vbCrLf
        sql &= " /*結訓人數*/" & vbCrLf
        sql &= " ,COUNT(CASE WHEN cs.StudStatus not in (2,3) THEN 1 END ) sum_ENum" & vbCrLf
        sql &= " /*提前就業人數*/" & vbCrLf
        sql &= " ,COUNT(CASE WHEN cs.WkAheadOfSch='Y' THEN 1 END ) sum_WINum" & vbCrLf
        sql &= " ,COUNT(CASE WHEN cs.WkAheadOfSch='Y' and cs.StudStatus in (2,3) AND dbo.NVL(sg9.SureItem,'3')='3' THEN 1 END ) sum_WINumS3" & vbCrLf
        sql &= " ,COUNT(CASE WHEN cs.WkAheadOfSch='Y' and cs.StudStatus in (2,3) AND dbo.NVL(sg9.SureItem,'3')='2' THEN 1 END ) sum_WINumS2" & vbCrLf
        sql &= " ,COUNT(CASE WHEN cs.WkAheadOfSch='Y' and cs.StudStatus in (2,3) AND dbo.NVL(sg9.SureItem,'3')='1' THEN 1 END ) sum_WINumS1" & vbCrLf
        sql &= " /*在職人數*/" & vbCrLf
        sql &= " ,COUNT(CASE WHEN cs.StudStatus not in (2,3) AND dbo.NVL(cs.WorkSuppIdent,'N')= 'Y' THEN 1 END ) sum_ISWork" & vbCrLf
        sql &= " /*就業人數*/" & vbCrLf
        sql &= " ,COUNT(CASE WHEN cs.StudStatus not in (2,3) AND sg3.socid IS NOT NULL and dbo.NVL(sg3.IsGetJob,0)=1 THEN 1 END ) sum_INum" & vbCrLf
        sql &= " /*1.就業人數再細分為1.系統判定 勞保勾稽及2.人工判定*/" & vbCrLf
        sql &= " ,COUNT(CASE WHEN cs.StudStatus not in (2,3) and sg3.IsGetJob = 1 AND sg3.mode_=1 THEN 1 END) sum_INumM1" & vbCrLf
        sql &= " ,COUNT(CASE WHEN cs.StudStatus not in (2,3) and sg3.IsGetJob = 1 AND sg3.mode_=2 THEN 1 END) sum_INumM2" & vbCrLf
        sql &= " /*就業人數-系統判定人數 3 */" & vbCrLf
        sql &= " ,COUNT(CASE WHEN cs.StudStatus not in (2,3) AND sg3.socid IS NOT NULL and dbo.NVL(sg3.IsGetJob,0)=1 AND dbo.NVL(sg3.SureItem,'3')='3' THEN 1 END ) sum_INum3" & vbCrLf
        sql &= " /*就業人數-雇用證明人數 1*/" & vbCrLf
        sql &= " ,COUNT(CASE WHEN cs.StudStatus not in (2,3) AND sg3.socid IS NOT NULL and dbo.NVL(sg3.IsGetJob,0)=1 AND dbo.NVL(sg3.SureItem,'3')='1' THEN 1 END ) sum_INum1" & vbCrLf
        sql &= " /*就業人數-就業切結人數 2*/" & vbCrLf
        sql &= " ,COUNT(CASE WHEN cs.StudStatus not in (2,3) AND sg3.socid IS NOT NULL and dbo.NVL(sg3.IsGetJob,0)=1 AND dbo.NVL(sg3.SureItem,'3')='2'THEN 1 END ) sum_INum2" & vbCrLf
        sql &= " /*未就業人數1*/" & vbCrLf
        sql &= " ,COUNT(CASE WHEN cs.StudStatus not in (2,3) and sg3.getjobcode='11' AND sg3.SOCID IS NOT NULL THEN 1 END ) sum_NJob11" & vbCrLf
        sql &= " /*未就業人數2*/" & vbCrLf
        sql &= " ,COUNT(CASE WHEN cs.StudStatus not in (2,3) and sg3.getjobcode='12' AND sg3.SOCID IS NOT NULL THEN 1 END ) sum_NJob12" & vbCrLf
        sql &= " /*未就業人數3*/" & vbCrLf
        sql &= " ,COUNT(CASE WHEN cs.StudStatus not in (2,3) and sg3.getjobcode='13' AND sg3.SOCID IS NOT NULL THEN 1 END ) sum_NJob13" & vbCrLf
        sql &= " /*未就業人數4*/" & vbCrLf
        sql &= " ,COUNT(CASE WHEN cs.StudStatus not in (2,3) and sg3.getjobcode='14' AND sg3.SOCID IS NOT NULL THEN 1 END ) sum_NJob14" & vbCrLf
        sql &= " /*未就業人數99*/" & vbCrLf
        sql &= " ,COUNT(CASE WHEN cs.StudStatus not in (2,3)" & vbCrLf
        sql &= " AND (dbo.NVL(sg3.getjobcode,'99')='99' AND dbo.NVL(sg3.IsGetJob,0)=0)" & vbCrLf
        sql &= " AND (sg3.SOCID IS NOT NULL or cs.StudStatus not in (2,3)) THEN 1 END ) sum_NJob99" & vbCrLf
        sql &= " /*不就業人數1*/" & vbCrLf
        sql &= " ,COUNT(CASE WHEN cs.StudStatus not in (2,3) and sg3.IsGetJob=2 and sg3.getjobcode='01' AND sg3.SOCID IS NOT NULL THEN 1 END ) sum_NJob01" & vbCrLf
        sql &= " /*不就業人數2*/" & vbCrLf
        sql &= " ,COUNT(CASE WHEN cs.StudStatus not in (2,3) and sg3.IsGetJob=2 and sg3.getjobcode='02' AND sg3.SOCID IS NOT NULL THEN 1 END ) sum_NJob02" & vbCrLf
        sql &= " /*不就業人數3*/" & vbCrLf
        sql &= " ,COUNT(CASE WHEN cs.StudStatus not in (2,3) and sg3.IsGetJob=2 and sg3.getjobcode='03' AND sg3.SOCID IS NOT NULL THEN 1 END ) sum_NJob03" & vbCrLf
        sql &= " /*不就業人數4*/" & vbCrLf
        sql &= " ,COUNT(CASE WHEN cs.StudStatus not in (2,3) and sg3.IsGetJob=2 and sg3.getjobcode in ('04','06') AND sg3.SOCID IS NOT NULL THEN 1 END ) sum_NJob04" & vbCrLf
        sql &= " /*不就業人數5*/" & vbCrLf
        sql &= " ,COUNT(CASE WHEN cs.StudStatus not in (2,3) and sg3.IsGetJob=2 and sg3.getjobcode='05' AND sg3.SOCID IS NOT NULL THEN 1 END ) sum_NJob05" & vbCrLf
        sql &= " from Class_StudentsOfClass cs" & vbCrLf
        sql &= " JOIN WC1 cc on cc.ocid =cs.ocid" & vbCrLf
        sql &= " left join Stud_GetJobState3 sg3 on sg3.socid =cs.socid and sg3.CPoint= @CPoint" & vbCrLf
        sql &= " left join Stud_GetJobState3 sg9 on sg9.socid =cs.socid and sg9.CPoint= 9" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " group by cs.ocid" & vbCrLf
        sql &= " )" & vbCrLf

        sql &= " select cc.ctnameN1 ""轄區縣市""" & vbCrLf
        sql &= " ,cc.ctnameN2 ""上課縣市""" & vbCrLf
        sql &= " ,cc.orgname ""訓練機構""" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE) ""班級名稱""" & vbCrLf
        sql &= " ,CONVERT(varchar, cc.stdate, 111)+'~'+CONVERT(varchar, cc.ftdate, 111) ""訓練期間""" & vbCrLf
        sql &= " ,dbo.NVL(cs2.sum_ENum,0) ""結訓人數""" & vbCrLf
        sql &= " ,cc.distid" & vbCrLf
        sql &= " ,cc.comidno" & vbCrLf
        sql &= " ,cc.ocid" & vbCrLf
        sql &= " ,dbo.NVL(cs2.sum_WINum,0) ""提前就業人數""" & vbCrLf
        sql &= " ,dbo.NVL(cs2.sum_WINumS3,0) sum_WINumS3" & vbCrLf
        sql &= " ,dbo.NVL(cs2.sum_WINumS2,0) sum_WINumS2" & vbCrLf
        sql &= " ,dbo.NVL(cs2.sum_WINumS1,0) sum_WINumS1" & vbCrLf
        sql &= " ,dbo.NVL(cs2.sum_ISWork,0) sum_ISWork" & vbCrLf
        sql &= " ,dbo.NVL(cs2.sum_INum,0) ""就業人數""" & vbCrLf
        sql &= " ,dbo.NVL(cs2.sum_INum3,0) sum_INum3" & vbCrLf
        sql &= " ,dbo.NVL(cs2.sum_INum1,0) sum_INum1" & vbCrLf
        sql &= " ,dbo.NVL(cs2.sum_INum2,0) sum_INum2" & vbCrLf
        'sql &= " ,dbo.NVL(cs2.sum_INumM1,0) sum_INumM1" & vbCrLf
        'sql &= " ,dbo.NVL(cs2.sum_INumM2,0) sum_INumM2" & vbCrLf
        sql &= " ,dbo.NVL(cs2.sum_NJob11,0) AS ""曾經找工作但不順利""" & vbCrLf
        sql &= " ,dbo.NVL(cs2.sum_NJob12,0) AS ""曾經找到工作但已離職""" & vbCrLf
        sql &= " ,dbo.NVL(cs2.sum_NJob13,0) AS ""找不到技能相符的工作""" & vbCrLf
        sql &= " ,dbo.NVL(cs2.sum_NJob14,0) AS ""找不到滿意的工作""" & vbCrLf
        sql &= " ,dbo.NVL(cs2.sum_NJob99,0) AS ""其他""" & vbCrLf
        sql &= " ,dbo.NVL(cs2.sum_NJob01,0) AS ""升學""" & vbCrLf
        sql &= " ,dbo.NVL(cs2.sum_NJob02,0) AS ""就醫、就養、待產""" & vbCrLf
        sql &= " ,dbo.NVL(cs2.sum_NJob03,0) AS ""出國""" & vbCrLf
        sql &= " ,dbo.NVL(cs2.sum_NJob04,0) AS ""服役""" & vbCrLf
        sql &= " ,dbo.NVL(cs2.sum_NJob05,0) AS ""再訓""" & vbCrLf
        sql &= " FROM WC1 cc" & vbCrLf
        sql &= " LEFT JOIN WS1 cs2 on cs2.ocid =cc.ocid" & vbCrLf
        sql &= " ORDER BY cc.distid,cc.comidno,cc.ocid" & vbCrLf

        Dim sCmd As New SqlCommand(sql, objconn)
        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("CPoint", SqlDbType.VarChar).Value = CPoint.SelectedValue
            If Me.DistID.SelectedValue <> "" Then
                .Parameters.Add("DistID", SqlDbType.VarChar).Value = Me.DistID.SelectedValue
            End If
            If Me.RIDValue.Value <> "" Then
                .Parameters.Add("RID", SqlDbType.VarChar).Value = Me.RIDValue.Value
            End If
            If Me.TPlanID.SelectedValue <> "" Then
                .Parameters.Add("TPlanID", SqlDbType.VarChar).Value = Me.TPlanID.SelectedValue
            End If
            If OCIDValue.Value <> "" Then
                If OCIDValue.Value.IndexOf(",") = -1 Then
                    .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCIDValue.Value
                Else
                    vMSG = "班級為多選!!"
                End If
            Else
                If Me.STDate1.Text <> "" Then
                    .Parameters.Add("STDate1", SqlDbType.DateTime).Value = CDate(Me.STDate1.Text)
                End If
                If Me.STDate2.Text <> "" Then
                    .Parameters.Add("STDate2", SqlDbType.DateTime).Value = CDate(Me.STDate2.Text)
                End If
                If Me.FTDate1.Text <> "" Then
                    .Parameters.Add("FTDate1", SqlDbType.DateTime).Value = CDate(Me.FTDate1.Text)
                End If
                If Me.FTDate2.Text <> "" Then
                    .Parameters.Add("FTDate2", SqlDbType.DateTime).Value = CDate(Me.FTDate2.Text)
                End If
            End If

            If txtSearchClassName.Text <> "" Then
                .Parameters.Add("classcname", SqlDbType.VarChar).Value = txtSearchClassName.Text '.Replace("'", "''")
            End If
            dt.Load(.ExecuteReader())
        End With

        If dt.Rows.Count > 0 Then
            Const cst_ccf1 As String = "DISTID"
            Const cst_ccf2 As String = "COMIDNO"
            Const cst_ccf3 As String = "OCID"
            Dim del_f1 As Boolean = False '刪除記號
            Dim del_f2 As Boolean = False
            Dim del_f3 As Boolean = False
            For i As Integer = 0 To dt.Columns.Count - 1
                Select Case UCase(dt.Columns(i).ColumnName)
                    Case "SUM_WINUMS3"
                        dt.Columns(i).ColumnName = "提前就業勞保勾稽"
                    Case "SUM_WINUMS2"
                        dt.Columns(i).ColumnName = "提前就業學員切結"
                    Case "SUM_WINUMS1"
                        dt.Columns(i).ColumnName = "提前就業雇主切結"
                    Case "SUM_ISWORK"
                        dt.Columns(i).ColumnName = "在職者人數"
                    Case "SUM_INUMM1"
                        dt.Columns(i).ColumnName = "就業人數勞保勾稽"
                    Case "SUM_INUMM2"
                        dt.Columns(i).ColumnName = "就業人數人工判定"
                    Case "SUM_INUM3"
                        dt.Columns(i).ColumnName = "系統判定人數"
                    Case "SUM_INUM1"
                        dt.Columns(i).ColumnName = "雇用證明人數"
                    Case "SUM_INUM2"
                        dt.Columns(i).ColumnName = "就業切結人數"
                    Case cst_ccf1
                        del_f1 = True
                    Case cst_ccf2
                        del_f2 = True
                    Case cst_ccf3
                        del_f3 = True
                End Select
            Next
            If del_f1 Then dt.Columns.Remove(cst_ccf1)
            If del_f2 Then dt.Columns.Remove(cst_ccf2)
            If del_f3 Then dt.Columns.Remove(cst_ccf3)
        End If

        msg.Text = "查無資料"
        DataGrid1.Visible = False
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            rst = True
            With DataGrid1
                .Visible = True
                .DataSource = dt
                .DataBind()
            End With
        End If

        Return rst
    End Function

    '覆寫，不執行 MyBase.VerifyRenderingInServerForm 方法，解決執行 RenderControl 產生的錯誤   
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)

    End Sub

    '匯出Excel檔
    Sub sUtl_Export1()
        '列印以及匯出Excel檔，這兩個功能，"提前就業人數"欄位的右邊，插入"提前就業-勞保勾稽"
        '、"提前就業-學員切結"、"提前就業-雇主切結"欄位，另外，
        '把目前的"公法就助人數"欄位改名為"訓後就業公法就助人數"，
        '並在右邊增加"提前就業公法救助人數"。
        Dim sFileName As String = "就業追蹤統計表_依班別.xls"
        sFileName = HttpUtility.UrlEncode(sFileName, System.Text.Encoding.UTF8)
        Response.Clear()
        Response.Buffer = True
        Response.Charset = "UTF-8" '設定字集
        Response.AppendHeader("Content-Disposition", "attachment;filename=" & sFileName)
        Response.ContentEncoding = System.Text.Encoding.GetEncoding("UTF-8")
        'Response.ContentType = "application/vnd.ms-excel " '內容型態設為Excel
        '文件內容指定為Excel
        Response.ContentType = "application/ms-excel;charset=utf-8"
        Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=UTF-8>")
        ''套CSS值
        Common.RespWrite(Me, "<style>")
        Common.RespWrite(Me, "td{mso-number-format:""\@"";}")
        Common.RespWrite(Me, "</style>")
        DataGrid1.AllowPaging = False '關閉分頁功能
        'DataGrid1.Columns(8).Visible = False
        DataGrid1.EnableViewState = False  '把ViewState給關了
        Dim objStringWriter As New System.IO.StringWriter
        Dim objHtmlTextWriter As New System.Web.UI.HtmlTextWriter(objStringWriter)
        Div1.RenderControl(objHtmlTextWriter)
        Common.RespWrite(Me, TIMS.sUtl_AntiXss(Convert.ToString(objStringWriter)))
        Response.End()
        DataGrid1.Visible = False
        'DataGrid1.AllowPaging = True
        'DataGrid1.Columns(8).Visible = True
    End Sub

    '匯出Excel檔
    Private Sub btnExport1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExport1.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        'Dim rtnPath As String = Request.FilePath
        Dim InputVal As String = ""
        'Me.txtSearchClassName.Text = "" & Trim(Me.txtSearchClassName.Text)
        txtSearchClassName.Text = TIMS.ClearSQM(txtSearchClassName.Text)

        'SQL Injection
        InputVal = Me.txtSearchClassName.Text
        If InputVal <> "" Then
            If TIMS.CheckInput(InputVal) Then
                Common.MessageBox(Me, TIMS.cst_ErrorMsg2)
                Exit Sub
            End If
        End If

        If OCIDValue.Value <> "" Then
            '有輸入班別依班別為主不使用日期區間
            STDate1.Text = ""
            STDate2.Text = ""
            FTDate1.Text = ""
            FTDate2.Text = ""
        End If

        DataGrid1.AllowPaging = False '關閉分頁功能
        'DataGrid1.Columns(8).Visible = False
        DataGrid1.EnableViewState = False  '把ViewState給關了

        '查詢 [SQL]
        If Not Search1() Then
            Common.MessageBox(Page, "查無資料!!")
            Exit Sub
        End If
        If msg.Text <> "" Then
            Common.MessageBox(Page, msg.Text)
            Exit Sub
        End If

        '匯出EXCEL
        Call sUtl_Export1()
    End Sub
End Class
