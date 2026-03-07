Partial Class TR_04_009_R
    Inherits AuthBasePage

    'ReportQuery
    'TR_04_009_R_4.jrxml @TR
    Const cst_printFN1 As String = "TR_04_009_R_4"

    'Dim vsYears2 As String = ""
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
        '檢查Session是否存在 End

        If Not IsPostBack Then
            'Syear = TIMS.GetSyear(Syear)
            Call CreateItem()
        End If
    End Sub

    Sub CreateItem()
        '計畫複選
        TPlanID = TIMS.Get_TPlan(TPlanID, , 1, "Y")

        '身分別 (報表)   '就業追蹤統計表_依轄區 (TR_04_009_R_4) :7
        Identity = TIMS.Get_Identity(Identity, 7, objconn)
        Identity.Items.Insert(0, New ListItem("全部", ""))

        '年齡區間 (報表)
        ddlyearsOld = TIMS.Get_YearsOld(ddlyearsOld)
        ddlyearsOld.Items.Insert(0, New ListItem("全部", ""))

        '學歷 (報表)
        ddlDegreeID = TIMS.Get_Degree(ddlDegreeID, 1, objconn)
        ddlDegreeID.Items.Insert(0, New ListItem("全部", ""))

        '性別 (報表)
        ddlSex = TIMS.Get_vSex(ddlSex)
        ddlSex.Items.Insert(0, New ListItem("全部", ""))

        '選擇全部訓練計畫
        TPlanID.Attributes("onclick") = "SelectAll('TPlanID','TPlanHidden');"
        Identity.Attributes("onclick") = "SelectAll('Identity','hidIdentity');"
        ddlyearsOld.Attributes("onclick") = "SelectAll('ddlyearsOld','hidddlyearsOld');"
        ddlDegreeID.Attributes("onclick") = "SelectAll('ddlDegreeID','hidddlDegreeID');"
        ddlSex.Attributes("onclick") = "SelectAll('ddlSex','hidddlSex');"

        Button1.Attributes("onclick") = "return search();"

    End Sub

    '檢查1
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        '報表要用的訓練計畫參數
        Dim TPlanID1 As String = ""
        'Dim TPlanName As String = ""
        For i As Integer = 1 To Me.TPlanID.Items.Count - 1
            If Me.TPlanID.Items(i).Selected AndAlso TPlanID.Items(i).Value <> "" Then
                If TPlanID1 = "" Then
                    TPlanID1 = Convert.ToString("\'" & Me.TPlanID.Items(i).Value & "\'")
                    Exit For
                End If
            End If
        Next
        If TPlanID1 = "" Then
            Errmsg += "請選擇 訓練計畫" & vbCrLf
        End If

        Dim CPointValue As String = ""
        For i As Integer = 0 To Me.CPoint.Items.Count - 1
            If CPoint.Items(i).Selected AndAlso CPoint.Items(i).Value <> "" Then
                CPointValue = Convert.ToString("\'" & CPoint.Items(i).Value & "\'")
                Exit For
            End If
        Next
        If CPointValue = "" Then
            Errmsg += "請選擇 就業查核點" & vbCrLf
        End If

        STDate1.Text = TIMS.ClearSQM(STDate1.Text) '清空/整理
        If STDate1.Text <> "" Then
            If Not TIMS.IsDate1(STDate1.Text) Then
                Errmsg += "開訓期間 起始日期格式有誤" & vbCrLf
            End If
            If Errmsg = "" Then
                STDate1.Text = CDate(STDate1.Text).ToString("yyyy/MM/dd")
            End If
        End If

        STDate2.Text = TIMS.ClearSQM(STDate2.Text) '清空/整理
        If STDate2.Text <> "" Then
            'STDate2.Text = Trim(STDate2.Text)
            If Not TIMS.IsDate1(STDate2.Text) Then
                Errmsg += "開訓期間 迄止日期格式有誤" & vbCrLf
            End If
            If Errmsg = "" Then
                STDate2.Text = CDate(STDate2.Text).ToString("yyyy/MM/dd")
            End If
        End If

        FTDate1.Text = TIMS.ClearSQM(FTDate1.Text) '清空/整理
        If FTDate1.Text <> "" Then
            'FTDate1.Text = Trim(FTDate1.Text)
            If Not TIMS.IsDate1(FTDate1.Text) Then
                Errmsg += "結訓期間 起始日期格式有誤" & vbCrLf
            End If
            If Errmsg = "" Then
                FTDate1.Text = CDate(FTDate1.Text).ToString("yyyy/MM/dd")
            End If
        End If

        FTDate2.Text = TIMS.ClearSQM(FTDate2.Text) '清空/整理
        If FTDate2.Text <> "" Then
            'FTDate2.Text = Trim(FTDate2.Text)
            If Not TIMS.IsDate1(FTDate2.Text) Then
                Errmsg += "結訓期間 迄止日期格式有誤" & vbCrLf
            End If
            If Errmsg = "" Then
                FTDate2.Text = CDate(FTDate2.Text).ToString("yyyy/MM/dd")
            End If
        End If

        Dim Years06 As Integer = 0
        Dim Years07 As Integer = 0
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
        End If

        '或不選
        Dim vsYears2 As String = ""
        Call TIMS.InputYears2(vsYears2, STYear1)
        Call TIMS.InputYears2(vsYears2, STYear2)
        Call TIMS.InputYears2(vsYears2, FTYear1)
        Call TIMS.InputYears2(vsYears2, FTYear2)

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

        Dim stitle As String = ""
        Dim etitle As String = ""
        If STDate1.Text <> "" OrElse STDate2.Text <> "" Then
            stitle = STDate1.Text & " ~ " & STDate2.Text
        End If
        If FTDate1.Text <> "" OrElse FTDate2.Text <> "" Then
            etitle = FTDate1.Text & " ~ " & FTDate2.Text
        End If

        '就業查核點
        Dim CPoint_list As String = ""
        CPoint_list = CPoint.SelectedItem.Text & "就業率"

        '報表要用的訓練計畫參數
        Dim TPlanID1 As String = ""
        'Dim TPlanName As String = ""
        For i As Integer = 1 To Me.TPlanID.Items.Count - 1
            If Me.TPlanID.Items(i).Selected AndAlso Me.TPlanID.Items(i).Value <> "" Then
                If TPlanID1 <> "" Then TPlanID1 &= ","
                TPlanID1 &= "\'" & Me.TPlanID.Items(i).Value & "\'"

                'If TPlanName <> "" Then TPlanName &= Convert.ToString(",")
                'TPlanName &= Convert.ToString(Me.TPlanID.Items(i).Text)
            End If
        Next

        Dim identityid As String = ""
        Dim YID As String = ""
        Dim DegreeID As String = ""
        Dim SexID As String = ""
        For i As Integer = 1 To Me.Identity.Items.Count - 1
            If Me.Identity.Items(i).Selected AndAlso Me.Identity.Items(i).Value <> "" Then
                If identityid <> "" Then identityid += ","
                identityid += "\'" & Me.Identity.Items(i).Value & "\'"
            End If
        Next
        For i As Integer = 1 To Me.ddlyearsOld.Items.Count - 1
            If Me.ddlyearsOld.Items(i).Selected AndAlso Me.ddlyearsOld.Items(i).Value <> "" Then
                If YID <> "" Then YID += ","
                YID += "\'" & Me.ddlyearsOld.Items(i).Value & "\'"
            End If
        Next
        For i As Integer = 1 To Me.ddlDegreeID.Items.Count - 1
            If Me.ddlDegreeID.Items(i).Selected AndAlso Me.ddlDegreeID.Items(i).Value <> "" Then
                If DegreeID <> "" Then DegreeID += ","
                DegreeID += "\'" & Me.ddlDegreeID.Items(i).Value & "\'"
            End If
        Next
        For i As Integer = 1 To Me.ddlSex.Items.Count - 1
            If Me.ddlSex.Items(i).Selected AndAlso Me.ddlSex.Items(i).Value <> "" Then
                If SexID <> "" Then SexID += ","
                SexID += "\'" & Me.ddlSex.Items(i).Value & "\'"
            End If
        Next

        Dim myValue As String = ""
        'myValue = "prg=TR_04_009_R"
        myValue = "k=1"
        myValue += "&TPlanID=" & TPlanID1
        'myValue += "&PlanName=" & Server.UrlEncode(TPlanName)
        myValue += "&CPoint=" & CPoint.SelectedValue
        'myValue += "&CPoint_list=" & CPoint_list
        myValue += "&CPoint_list=" & Server.UrlEncode(CPoint_list)
        'myValue += "&CPoint_list3=" & Server.UrlDecode(CPoint_list)
        myValue += "&STDate1=" & Me.STDate1.Text
        myValue += "&STDate2=" & Me.STDate2.Text
        myValue += "&FTDate1=" & Me.FTDate1.Text
        myValue += "&FTDate2=" & Me.FTDate2.Text
        myValue += "&identityid=" & identityid
        myValue += "&YID=" & YID
        myValue += "&DegreeID=" & DegreeID
        myValue += "&SexID=" & SexID
        myValue += "&stitle=" & stitle
        myValue += "&etitle=" & etitle
        'ReportQuery
        'Select Case PrintStyle.SelectedValue
        '    Case "2006"
        '        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "TR", "TR_04_009R", myValue)
        '    Case Else '2011
        '        'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "TR", "TR_04_009_R_2", myValue)
        '        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "TR", "TR_04_009_R_3", myValue)
        'End Select
        'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "TR", "TR_04_009_R_3", myValue)
        '2012增加條件
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, myValue)

    End Sub

    '匯出班級明細  '匯出Excel檔
    Protected Sub btnExport1_Click(sender As Object, e As EventArgs) Handles btnExport1.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
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

        '匯出班級明細  '匯出Excel檔
        Call sUtl_Export1()
    End Sub

    '查詢 [SQL]
    Private Function Search1() As Boolean
        Dim rst As Boolean = False '是否有資料 ，預設沒有資料

        Dim stitle As String = ""
        Dim etitle As String = ""
        If STDate1.Text <> "" OrElse STDate2.Text <> "" Then
            stitle = STDate1.Text & " ~ " & STDate2.Text
        End If
        If FTDate1.Text <> "" OrElse FTDate2.Text <> "" Then
            etitle = FTDate1.Text & " ~ " & FTDate2.Text
        End If

        '就業查核點
        Dim CPoint_list As String = ""
        CPoint_list = CPoint.SelectedItem.Text & "就業率"

        '報表要用的訓練計畫參數
        Dim TPlanID1 As String = ""
        For i As Integer = 1 To Me.TPlanID.Items.Count - 1
            If Me.TPlanID.Items(i).Selected AndAlso Me.TPlanID.Items(i).Value <> "" Then
                If TPlanID1 <> "" Then TPlanID1 &= ","
                TPlanID1 &= "'" & Me.TPlanID.Items(i).Value & "'"
            End If
        Next

        Dim identityid As String = ""
        Dim YID As String = ""
        Dim DegreeID As String = ""
        Dim SexID As String = ""
        For i As Integer = 1 To Me.Identity.Items.Count - 1
            If Me.Identity.Items(i).Selected AndAlso Me.Identity.Items(i).Value <> "" Then
                If identityid <> "" Then identityid += ","
                identityid += "'" & Me.Identity.Items(i).Value & "'"
            End If
        Next
        For i As Integer = 1 To Me.ddlyearsOld.Items.Count - 1
            If Me.ddlyearsOld.Items(i).Selected AndAlso Me.ddlyearsOld.Items(i).Value <> "" Then
                If YID <> "" Then YID += ","
                YID += "'" & Me.ddlyearsOld.Items(i).Value & "'"
            End If
        Next
        For i As Integer = 1 To Me.ddlDegreeID.Items.Count - 1
            If Me.ddlDegreeID.Items(i).Selected AndAlso Me.ddlDegreeID.Items(i).Value <> "" Then
                If DegreeID <> "" Then DegreeID += ","
                DegreeID += "'" & Me.ddlDegreeID.Items(i).Value & "'"
            End If
        Next
        For i As Integer = 1 To Me.ddlSex.Items.Count - 1
            If Me.ddlSex.Items(i).Selected AndAlso Me.ddlSex.Items(i).Value <> "" Then
                If SexID <> "" Then SexID &= ","
                SexID &= "'" & Me.ddlSex.Items(i).Value & "'"
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
        sql &= " ,ip.planname" & vbCrLf
        sql &= " ,cc.comidno" & vbCrLf
        sql &= " ,cc.ocid" & vbCrLf
        sql &= " from class_classinfo cc" & vbCrLf
        sql &= " join org_orginfo oo on oo.comidno=cc.comidno" & vbCrLf
        sql &= " join view_plan ip on ip.planid=cc.planid" & vbCrLf
        sql &= " join view_ridname vr on vr.RID =cc.RID" & vbCrLf
        sql &= " left join id_zip iz1 on iz1.zipcode=cc.taddresszip" & vbCrLf
        sql &= " left join id_zip iz2 on iz2.zipcode=vr.zipcode" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND cc.NotOpen = 'N'" & vbCrLf
        sql &= " and cc.FTDate < getdate()" & vbCrLf
        '不列入就業率統計 join cc
        sql &= " and NOT EXISTS (SELECT 'x' FROM Class_NoWorkRate cnw WHERE cnw.OCID=cc.OCID)" & vbCrLf
        'sql &= " and ip.TPlanID IN ('02')" & vbCrLf
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
        If TPlanID1 <> "" Then
            sql &= " and ip.TPlanID IN (" & TPlanID1 & ")" & vbCrLf
        End If
        sql &= " )" & vbCrLf

        sql &= " ,WS1 AS (" & vbCrLf
        sql &= " select cc.OCID" & vbCrLf
        'sql &= " 	/*結訓人數*/" & vbCrLf
        sql &= "    ,COUNT (CASE WHEN cs.StudStatus not in (2,3) THEN 1 END ) sum_ENum" & vbCrLf
        'sql &= " 	/*提前就業人數*/" & vbCrLf
        sql &= "    ,COUNT(CASE WHEN cs.WkAheadOfSch='Y' and cs.StudStatus in (2,3) THEN 1 END ) sum_WINum" & vbCrLf
        'sql &= " 	/*SureItem: 1:雇主切結 2:學員切結 3:勞保勾稽(null)" & vbCrLf
        'sql &= " 	提前就業-勞保勾稽3 提前就業-學員切結2 提前就業-雇主切結1*/" & vbCrLf
        sql &= " 	,COUNT(CASE WHEN cs.WkAheadOfSch='Y' and cs.StudStatus in (2,3) AND sg9.mode_=1 and dbo.NVL(sg9.SureItem,'3')='3' THEN 1 END ) sum_WINumS3" & vbCrLf
        sql &= " 	,COUNT(CASE WHEN cs.WkAheadOfSch='Y' and cs.StudStatus in (2,3) AND sg9.mode_=2 and dbo.NVL(sg9.SureItem,'3')='2' THEN 1 END ) sum_WINumS2" & vbCrLf
        sql &= " 	,COUNT(CASE WHEN cs.WkAheadOfSch='Y' and cs.StudStatus in (2,3) AND sg9.mode_=2 and dbo.NVL(sg9.SureItem,'3')='1' THEN 1 END ) sum_WINumS1" & vbCrLf
        'sql &= " 	/*在職人數*/" & vbCrLf
        sql &= " 	,COUNT (CASE WHEN cs.StudStatus not in (2,3) AND dbo.NVL(cs.WorkSuppIdent,'N')= 'Y' THEN 1 END ) sum_ISWork" & vbCrLf
        'sql &= " 	/*就業人數  dbo.NVL(sg3.IsGetJob,0)=1 */" & vbCrLf
        sql &= " 	,COUNT (CASE WHEN cs.StudStatus not in (2,3) and sg3.IsGetJob = 1 THEN 1 END ) sum_INum" & vbCrLf
        'sql &= " 	/*1.就業人數再細分為1.系統判定 勞保勾稽及 2.人工判定*/" & vbCrLf
        sql &= " 	,COUNT (CASE WHEN cs.StudStatus not in (2,3) and sg3.IsGetJob = 1 AND sg3.mode_=1 THEN 1 END ) sum_INumM1" & vbCrLf
        sql &= " 	,COUNT (CASE WHEN cs.StudStatus not in (2,3) and sg3.IsGetJob = 1 AND sg3.mode_=2 THEN 1 END ) sum_INumM2" & vbCrLf
        'sql &= " 	/*就業人數-系統判定人數 3 就業人數-雇用證明人數 1 就業人數-就業切結人數 2*/" & vbCrLf
        sql &= " 	,COUNT(CASE WHEN cs.StudStatus not in (2,3) AND sg3.socid IS NOT NULL and dbo.NVL(sg3.IsGetJob,0)=1 AND dbo.NVL(sg3.SureItem,'3')='3' THEN 1 END ) sum_INum3" & vbCrLf
        sql &= " 	,COUNT(CASE WHEN cs.StudStatus not in (2,3) AND sg3.socid IS NOT NULL and dbo.NVL(sg3.IsGetJob,0)=1 AND dbo.NVL(sg3.SureItem,'3')='1' THEN 1 END ) sum_INum1" & vbCrLf
        sql &= " 	,COUNT(CASE WHEN cs.StudStatus not in (2,3) AND sg3.socid IS NOT NULL and dbo.NVL(sg3.IsGetJob,0)=1 AND dbo.NVL(sg3.SureItem,'3')='2' THEN 1 END ) sum_INum2" & vbCrLf

        'sql &= " 	/*未就業人數 dbo.NVL(sg3.IsGetJob,0)=0 */" & vbCrLf
        sql &= " 	,COUNT (CASE WHEN cs.StudStatus not in (2,3) and sg3.getjobcode='11' AND sg3.SOCID IS NOT NULL THEN 1  END ) sum_NJob11" & vbCrLf
        sql &= " 	,COUNT (CASE WHEN cs.StudStatus not in (2,3) and sg3.getjobcode='12' AND sg3.SOCID IS NOT NULL THEN 1  END ) sum_NJob12" & vbCrLf
        sql &= " 	,COUNT (CASE WHEN cs.StudStatus not in (2,3) and sg3.getjobcode='13' AND sg3.SOCID IS NOT NULL THEN 1  END ) sum_NJob13" & vbCrLf
        sql &= " 	,COUNT (CASE WHEN cs.StudStatus not in (2,3) and sg3.getjobcode='14' AND sg3.SOCID IS NOT NULL THEN 1  END ) sum_NJob14" & vbCrLf
        sql &= " 	,COUNT (CASE WHEN cs.StudStatus not in (2,3) AND (dbo.NVL(sg3.getjobcode,'99')='99' AND dbo.NVL(sg3.IsGetJob,0)=0)" & vbCrLf
        sql &= " 	AND (sg3.SOCID IS NOT NULL or cs.StudStatus not in (2,3)) THEN 1  END ) sum_NJob99" & vbCrLf

        'sql &= " 	/*不就業 sg3.IsGetJob=2 */" & vbCrLf
        sql &= " 	,COUNT (CASE WHEN cs.StudStatus not in (2,3) and sg3.IsGetJob=2 and sg3.getjobcode='01' AND sg3.SOCID IS NOT NULL THEN 1  END ) sum_NJob01" & vbCrLf
        sql &= " 	,COUNT (CASE WHEN cs.StudStatus not in (2,3) and sg3.IsGetJob=2 and sg3.getjobcode='02' AND sg3.SOCID IS NOT NULL THEN 1  END ) sum_NJob02" & vbCrLf
        sql &= " 	,COUNT (CASE WHEN cs.StudStatus not in (2,3) and sg3.IsGetJob=2 and sg3.getjobcode='03' AND sg3.SOCID IS NOT NULL THEN 1  END ) sum_NJob03" & vbCrLf
        sql &= " 	,COUNT (CASE WHEN cs.StudStatus not in (2,3) and sg3.IsGetJob=2 and sg3.getjobcode in ('04','06') AND sg3.SOCID IS NOT NULL THEN 1  END ) sum_NJob04" & vbCrLf
        sql &= " 	,COUNT (CASE WHEN cs.StudStatus not in (2,3) and sg3.IsGetJob=2 and sg3.getjobcode='05' AND sg3.SOCID IS NOT NULL THEN 1  END ) sum_NJob05" & vbCrLf
        'sql &= " 	/*「公法救助」*/" & vbCrLf
        sql &= "    ,COUNT(CASE WHEN cs.StudStatus not in (2,3) and sg3.PUBLICRESCUE='Y' AND sg3.SOCID IS NOT NULL THEN 1 END ) sum_PUBLICRESCUE" & vbCrLf
        'sql &= " 	/*提前就業「公法救助」*/" & vbCrLf
        sql &= "    ,COUNT(CASE WHEN cs.StudStatus in (2,3) and sg9.PUBLICRESCUE='Y' AND sg9.SOCID IS NOT NULL THEN 1 END ) sum_PUBLICRESCUE9" & vbCrLf
        sql &= "    ,COUNT(case when cs.StudStatus NOT IN (2,3) and sg3.JOBRELATE='Y' then 1 end) sum_jobrelx /*就業關聯性*/" & vbCrLf
        sql &= "    from WC1 cc" & vbCrLf
        sql &= "    JOIN Class_StudentsOfClass cs on cs.ocid =cc.ocid" & vbCrLf
        sql &= "    join stud_studentinfo ss on ss.SID=cs.SID" & vbCrLf
        sql &= "    left join id_yearRange iy on iy.yearsOld=DATEPART(YEAR, cc.STDate)-DATEPART(YEAR, ss.Birthday)" & vbCrLf
        sql &= "    left join Stud_GetJobState3 sg3 on sg3.socid =cs.socid and sg3.CPoint= @CPoint" & vbCrLf
        sql &= "    left join Stud_GetJobState3 sg9 on sg9.socid =cs.socid and sg9.CPoint= 9" & vbCrLf
        sql &= "    WHERE 1=1" & vbCrLf
        If identityid <> "" Then
            sql &= " AND cs.MIdentityID in (" & identityid & ")" & vbCrLf
        End If
        If YID <> "" Then
            sql &= " AND iy.YID in (" & YID & ")" & vbCrLf
        End If
        If DegreeID <> "" Then
            sql &= " AND ss.DegreeID in (" & DegreeID & ")" & vbCrLf
        End If
        If SexID <> "" Then
            sql &= " AND ss.Sex in (" & SexID & ")" & vbCrLf
        End If
        sql &= " GROUP BY cc.OCID" & vbCrLf
        sql &= " )" & vbCrLf

        sql &= " select cc.planname ""訓練計畫""" & vbCrLf
        sql &= " ,cc.orgname ""訓練機構""" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE) ""班級名稱""" & vbCrLf
        sql &= " ,CONVERT(varchar, cc.stdate, 111)+'~'+CONVERT(varchar, cc.ftdate, 111) ""訓練期間""" & vbCrLf
        sql &= " ,dbo.NVL(cs2.sum_ENum,0) ""結訓人數""" & vbCrLf

        sql &= " ,dbo.NVL(cs2.sum_WINum,0) ""提前就業人數""" & vbCrLf
        sql &= " ,dbo.NVL(cs2.sum_WINumS3,0) sum_WINumS3" & vbCrLf '勞保勾稽人數
        sql &= " ,dbo.NVL(cs2.sum_WINumS2,0) sum_WINumS2" & vbCrLf '學員切結人數
        sql &= " ,dbo.NVL(cs2.sum_WINumS1,0) sum_WINumS1" & vbCrLf '雇主切結人數

        sql &= " ,dbo.NVL(cs2.sum_INum,0) ""就業人數""" & vbCrLf
        sql &= " ,dbo.NVL(cs2.sum_INum3,0) sum_INum3" & vbCrLf '勞保勾稽人數
        sql &= " ,dbo.NVL(cs2.sum_INum2,0) sum_INum2" & vbCrLf '學員切結人數
        sql &= " ,dbo.NVL(cs2.sum_INum1,0) sum_INum1" & vbCrLf '雇主切結人數

        sql &= " ,dbo.NVL(cs2.sum_NJob11,0)+dbo.NVL(cs2.sum_NJob12,0)+dbo.NVL(cs2.sum_NJob13,0)+dbo.NVL(cs2.sum_NJob14,0)+dbo.NVL(cs2.sum_NJob99,0) ""未就業人數""" & vbCrLf
        sql &= " ,dbo.NVL(cs2.SUM_NJOB11,0) SUM_NJOB11" & vbCrLf '未就業曾經找工作但不順利
        sql &= " ,dbo.NVL(cs2.SUM_NJOB12,0) SUM_NJOB12" & vbCrLf '未就業曾經找到工作但已離職
        sql &= " ,dbo.NVL(cs2.SUM_NJOB13,0) SUM_NJOB13" & vbCrLf '未就業找不到技能相符的工作
        sql &= " ,dbo.NVL(cs2.SUM_NJOB14,0) SUM_NJOB14" & vbCrLf '未就業找不到滿意的工作
        sql &= " ,dbo.NVL(cs2.SUM_NJOB99,0) SUM_NJOB99" & vbCrLf '未就業其他

        sql &= " ,dbo.NVL(cs2.sum_NJob01,0)+dbo.NVL(cs2.sum_NJob02,0)+dbo.NVL(cs2.sum_NJob03,0)+dbo.NVL(cs2.sum_NJob04,0)+dbo.NVL(cs2.sum_NJob05,0) ""不就業人數""" & vbCrLf
        sql &= " ,dbo.NVL(cs2.SUM_NJOB01,0) SUM_NJOB01" & vbCrLf '不就業升學
        sql &= " ,dbo.NVL(cs2.SUM_NJOB02,0) SUM_NJOB02" & vbCrLf '不就業就醫、就養、待產
        sql &= " ,dbo.NVL(cs2.SUM_NJOB03,0) SUM_NJOB03" & vbCrLf '不就業出國
        sql &= " ,dbo.NVL(cs2.SUM_NJOB04,0) SUM_NJOB04" & vbCrLf '不就業服役
        sql &= " ,dbo.NVL(cs2.SUM_NJOB05,0) SUM_NJOB05" & vbCrLf '不就業再訓

        sql &= " ,dbo.NVL(cs2.sum_PUBLICRESCUE,0) sum_PUBLICRESCUE" & vbCrLf '訓後就業公法救助人數
        sql &= " ,dbo.NVL(cs2.sum_PUBLICRESCUE9,0) sum_PUBLICRESCUE9" & vbCrLf '提前就業公法救助人數
        sql &= " ,dbo.NVL(cs2.SUM_JOBRELX,0) SUM_JOBRELX" & vbCrLf '就業關聯人數
        sql &= " ,dbo.NVL(cs2.sum_ISWork,0) ""在職者""" & vbCrLf '在職者

        sql &= " FROM WC1 cc" & vbCrLf
        sql &= " LEFT JOIN WS1 cs2 on cs2.ocid =cc.ocid" & vbCrLf
        sql &= " ORDER BY cc.distid,cc.comidno,cc.ocid" & vbCrLf
        Dim sCmd As New SqlCommand(sql, objconn)

        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("CPoint", SqlDbType.VarChar).Value = CPoint.SelectedValue
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

            dt.Load(.ExecuteReader())
        End With

        If dt.Rows.Count > 0 Then
            'Const cst_ccf1 As String = "DISTID"
            'Const cst_ccf2 As String = "COMIDNO"
            'Const cst_ccf3 As String = "OCID"
            'Dim del_f1 As Boolean = False '刪除記號
            'Dim del_f2 As Boolean = False
            'Dim del_f3 As Boolean = False
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
                        dt.Columns(i).ColumnName = "就業人數勞保勾稽"
                    Case "SUM_INUM2"
                        dt.Columns(i).ColumnName = "就業人數學員切結"
                    Case "SUM_INUM1"
                        dt.Columns(i).ColumnName = "就業人數雇主切結"

                    Case "SUM_NJOB11"
                        dt.Columns(i).ColumnName = "未就業曾經找工作但不順利"
                    Case "SUM_NJOB12"
                        dt.Columns(i).ColumnName = "未就業曾經找到工作但已離職"
                    Case "SUM_NJOB13"
                        dt.Columns(i).ColumnName = "未就業找不到技能相符的工作"
                    Case "SUM_NJOB14"
                        dt.Columns(i).ColumnName = "未就業找不到滿意的工作"
                    Case "SUM_NJOB99"
                        dt.Columns(i).ColumnName = "未就業其他"

                    Case "SUM_NJOB01"
                        dt.Columns(i).ColumnName = "不就業升學"
                    Case "SUM_NJOB02"
                        dt.Columns(i).ColumnName = "不就業就醫、就養、待產"
                    Case "SUM_NJOB03"
                        dt.Columns(i).ColumnName = "不就業出國"
                    Case "SUM_NJOB04"
                        dt.Columns(i).ColumnName = "不就業服役"
                    Case "SUM_NJOB05"
                        dt.Columns(i).ColumnName = "不就業再訓"

                    Case "SUM_PUBLICRESCUE"
                        dt.Columns(i).ColumnName = "訓後就業公法救助人數"
                    Case "SUM_PUBLICRESCUE9"
                        dt.Columns(i).ColumnName = "提前就業公法救助人數"
                    Case "SUM_JOBRELX"
                        dt.Columns(i).ColumnName = "就業關聯人數"

                End Select
            Next
            'If del_f1 Then dt.Columns.Remove(cst_ccf1)
            'If del_f2 Then dt.Columns.Remove(cst_ccf2)
            'If del_f3 Then dt.Columns.Remove(cst_ccf3)
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

    '匯出班級明細  '匯出Excel檔
    Sub sUtl_Export1()
        '列印以及匯出Excel檔，這兩個功能，"提前就業人數"欄位的右邊，插入"提前就業-勞保勾稽"
        '、"提前就業-學員切結"、"提前就業-雇主切結"欄位，另外，
        '把目前的"公法就助人數"欄位改名為"訓後就業公法就助人數"，
        '並在右邊增加"提前就業公法救助人數"。
        Dim sFileName As String = "就業追蹤統計表_依轄區.xls"
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

End Class

