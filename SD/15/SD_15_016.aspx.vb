Public Class SD_15_016
    Inherits AuthBasePage

    Dim gDistIDVal As String = ""
    Dim gTCityCode2 As String = ""
    Dim gOCityCode2 As String = ""
    Dim gdt As DataTable
    Dim gdt2 As DataTable

    Dim objConn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objConn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objConn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End
        '分頁設定 Start
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End

        If Not IsPostBack Then
            msg.Text = ""
            PageControler1.Visible = False
            DataGrid1.Visible = False

            yearlist = TIMS.GetSyear(yearlist)
            Common.SetListItem(yearlist, sm.UserInfo.Years)

            Distid = TIMS.Get_DistID(Distid)
            Distid.Items.Insert(0, New ListItem("全部", 0))

            Tcitycode = TIMS.Get_CityName(Tcitycode, TIMS.dtNothing)
            Ocitycode = TIMS.Get_CityName(Ocitycode, TIMS.dtNothing)

            Distid.Attributes("onclick") = "SelectAll('Distid','DistHidden');"

            Tcitycode.Attributes("onclick") = "SelectAll('Tcitycode','TcityHidden');"
            Ocitycode.Attributes("onclick") = "SelectAll('Ocitycode','OcityHidden');"

            Distid.Enabled = True
            If sm.UserInfo.DistID <> "000" Then
                'Distid.Attributes("onclick") = "alert('x');"
                ''Distid.SelectedValue = sm.UserInfo.DistID
                Common.SetListItem(Distid, sm.UserInfo.DistID)
                Distid.Enabled = False
            End If

        End If
    End Sub

    '檢查輸入資料是否正確
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        If Me.yearlist.SelectedValue = "" Then
            Errmsg += "請選擇計畫年度" & vbCrLf
        End If

        Dim j As Integer = 0
        Dim CBLobj As CheckBoxList
        j = 0
        CBLobj = Distid
        For i As Integer = 1 To CBLobj.Items.Count - 1
            Dim objitem As ListItem = CBLobj.Items(i)
            If objitem.Selected = True Then
                j += 1
                Exit For
            End If
        Next
        If j = 0 Then Errmsg += "請選擇轄區" & vbCrLf

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    '求遞補人數
    Function Get_RecNum(ByVal TNum As Integer, ByVal OCID As String) As Integer
        Dim rst As Integer = 0

        Dim i As Integer = 0
        Dim idnoVals As String = ""
        i = 0
        idnoVals = ""
        For Each dr As DataRow In gdt.Select("OCID1='" & OCID & "'", "RelEnterDate")
            i += 1
            If idnoVals <> "" Then idnoVals &= ","
            idnoVals &= "'" & Convert.ToString(dr("idno")) & "'"
            If i >= TNum Then Exit For '前N筆報名 離開
        Next

        '1.報名人數大於核定人數(且依報名順序排定)
        If gdt2.Select("OCID='" & OCID & "'").Length > 0 Then
            '有班級學員人數 (班級學員確認)
            For Each dr2 As DataRow In gdt2.Select("OCID='" & OCID & "'")
                '找不到該學員的報名資料(前 N 名錄取人員)，即為遞補人數
                If idnoVals.IndexOf(dr2("idno")) = -1 Then
                    rst += 1
                End If
            Next
        End If
        'dt.Dispose()
        'dt2.Dispose()
        Return rst
    End Function

    Function sUtl_ChangeDt(ByRef dt As DataTable) As DataTable
        Dim parms As Hashtable = New Hashtable()
        Dim sql As String = ""
        '前 N 名錄取人員
        sql = "" & vbCrLf
        'signUpStatus 0:收件完成 1:報名成功 2:報名失敗 3:正取(Key_SelResult) 4:備取 5:未錄取
        sql += " select a.OCID1,a.idno,a.RelEnterDate" & vbCrLf
        sql += " from v_EnterType2a a " & vbCrLf
        sql += " JOIN view2 cc on cc.OCID= a.OCID1" & vbCrLf
        sql += " JOIN auth_relship ar on ar.RID =cc.RID " & vbCrLf
        sql += " JOIN Org_OrgPlanInfo op on op.RSID = ar.RSID" & vbCrLf
        sql += " LEFT JOIN view_ZipName iz2 on op.ZipCode = iz2.ZipCode" & vbCrLf
        sql += " LEFT JOIN view_ZipName iz3 on cc.TaddressZip = iz3.ZipCode" & vbCrLf
        sql += " where 1=1" & vbCrLf
        sql += " and cc.TPlanID=@TPlanID " & vbCrLf '計畫
        sql += " and cc.years=@years " & vbCrLf '年度

        parms.Add("TPlanID", sm.UserInfo.TPlanID)
        parms.Add("years", yearlist.SelectedValue)

        '轄區
        If gDistIDVal <> "" Then
            sql += " and cc.Distid in (" & gDistIDVal & ")" & vbCrLf
        End If
        '辦訓地縣市 
        If gTCityCode2 <> "" Then
            sql += " and iz3.CTID in (" & gTCityCode2 & ")" & vbCrLf
        End If
        '立案地縣市
        If gOCityCode2 <> "" Then
            sql += " and iz2.CTID in (" & gOCityCode2 & ")" & vbCrLf
        End If
        If SDate1.Text <> "" Then
            'sql += " and cc.STDate >= " & TIMS.to_date(SDate1.Text) & vbCrLf
            sql += " and cc.STDate >= @STDate1 " & vbCrLf
            parms.Add("STDate1", SDate1.Text)
        End If
        If SDate2.Text <> "" Then
            'sql += " and cc.STDate <= " & TIMS.to_date(SDate2.Text) & vbCrLf '" & SDate2.Text & "'" & vbCrLf
            sql += " and cc.STDate <= @STDate2 " & vbCrLf
            parms.Add("STDate2", SDate2.Text)
        End If
        If EDate1.Text <> "" Then
            'sql += " and cc.FTDate >= " & TIMS.to_date(EDate1.Text) & vbCrLf '" & EDate1.Text & "'" & vbCrLf
            sql += " and cc.FTDate >= @FTDate1 " & vbCrLf '" & EDate1.Text & "'" & vbCrLf
            parms.Add("FTDate1", EDate1.Text)
        End If
        If EDate2.Text <> "" Then
            'sql += " and cc.FTDate <= " & TIMS.to_date(EDate2.Text) & vbCrLf '" & EDate2.Text & "'" & vbCrLf
            sql += " and cc.FTDate <= @FTDate2 " & vbCrLf '" & EDate1.Text & "'" & vbCrLf
            parms.Add("FTDate2", EDate2.Text)
        End If
        'sql += " order by RelEnterDate" & vbCrLf
        gdt = DbAccess.GetDataTable(sql, objConn, parms)

        parms.Clear()
        '學員人數
        sql = "" & vbCrLf
        sql += " select a.OCID,a.idno" & vbCrLf
        sql += " from v_StudentInfo a " & vbCrLf
        sql += " JOIN view2 cc on cc.OCID= a.OCID" & vbCrLf
        sql += " JOIN auth_relship ar on ar.RID =cc.RID " & vbCrLf
        sql += " JOIN Org_OrgPlanInfo op on op.RSID = ar.RSID" & vbCrLf
        sql += " LEFT JOIN view_ZipName iz2 on op.ZipCode = iz2.ZipCode" & vbCrLf
        sql += " LEFT JOIN view_ZipName iz3 on cc.TaddressZip = iz3.ZipCode" & vbCrLf
        sql += " where 1=1" & vbCrLf
        sql += " and cc.TPlanID=@TPlanID " & vbCrLf '計畫
        sql += " and cc.years=@years " & vbCrLf '年度

        parms.Add("TPlanID", sm.UserInfo.TPlanID)
        parms.Add("years", yearlist.SelectedValue)

        '轄區
        If gDistIDVal <> "" Then
            sql += " and cc.Distid in (" & gDistIDVal & ")" & vbCrLf
        End If
        '辦訓地縣市 
        If gTCityCode2 <> "" Then
            sql += " and iz3.CTID in (" & gTCityCode2 & ")" & vbCrLf
        End If
        '立案地縣市
        If gOCityCode2 <> "" Then
            sql += " and iz2.CTID in (" & gOCityCode2 & ")" & vbCrLf
        End If
        If SDate1.Text <> "" Then
            'sql += " and cc.STDate >= " & TIMS.to_date(SDate1.Text) & vbCrLf
            sql += " and cc.STDate >= @STDate1 " & vbCrLf
            parms.Add("STDate1", SDate1.Text)
        End If
        If SDate2.Text <> "" Then
            'sql += " and cc.STDate <= " & TIMS.to_date(SDate2.Text) & vbCrLf '" & SDate2.Text & "'" & vbCrLf
            sql += " and cc.STDate <= @STDate2 " & vbCrLf '" & SDate2.Text & "'" & vbCrLf
            parms.Add("STDate2", SDate2.Text)
        End If
        If EDate1.Text <> "" Then
            'sql += " and cc.FTDate >= " & TIMS.to_date(EDate1.Text) & vbCrLf '" & EDate1.Text & "'" & vbCrLf
            sql += " and cc.FTDate >= @FTDate1 " & vbCrLf '" & EDate1.Text & "'" & vbCrLf
            parms.Add("FTDate1", EDate1.Text)
        End If
        If EDate2.Text <> "" Then
            'sql += " and cc.FTDate <= " & TIMS.to_date(EDate2.Text) & vbCrLf '" & EDate2.Text & "'" & vbCrLf
            sql += " and cc.FTDate <= @FTDate2 " & vbCrLf '" & EDate2.Text & "'" & vbCrLf
            parms.Add("FTDate2", EDate2.Text)
        End If

        gdt2 = DbAccess.GetDataTable(sql, objConn, parms)

        For Each dr As DataRow In dt.Rows
            '報名人數 大於 核定人數 才處理
            If dr("報名人數") > dr("核定人數") Then
                '找不到該學員的報名資料(前 N 名錄取人員)，即為遞補人數
                dr("遞補人數") = Get_RecNum(dr("核定人數"), dr("OCID"))
            End If
        Next
        Return dt
    End Function

    '統計 SQL 查詢
    Sub Search1()
        '轄區
        'Dim DistIDVal As String = ""
        For i As Integer = 0 To Distid.Items.Count - 1
            If Distid.Items.Item(i).Selected = True Then
                If Distid.Items.Item(i).Text <> "全部" Then
                    If gDistIDVal <> "" Then gDistIDVal += ","
                    gDistIDVal += "'" & Distid.Items.Item(i).Value & "'"
                End If
            End If
        Next
        '辦訓地縣市
        'Dim TCityCode2 As String = ""
        For i As Integer = 0 To Tcitycode.Items.Count - 1
            If Tcitycode.Items.Item(i).Selected = True Then
                If Tcitycode.Items.Item(i).Text <> "全部" Then
                    If gTCityCode2 <> "" Then gTCityCode2 += ","
                    gTCityCode2 += "'" & Tcitycode.Items.Item(i).Value & "'"
                End If
            End If
        Next
        '立案地縣市
        'Dim OCityCode2 As String = ""
        For i As Integer = 0 To Ocitycode.Items.Count - 1
            If Ocitycode.Items.Item(i).Selected = True Then
                If Ocitycode.Items.Item(i).Text <> "全部" Then
                    If gOCityCode2 <> "" Then gOCityCode2 += ","
                    gOCityCode2 += "'" & Ocitycode.Items.Item(i).Value & "'"
                End If
            End If
        Next

        Dim parms As Hashtable = New Hashtable()
        'Dim sql As String = ""
        'sql = "" & vbCrLf
        'sql += " select cc.ocid" & vbCrLf
        'sql += " ,cc.distname 轄區" & vbCrLf
        'sql += " ,cc.orgname 單位名稱" & vbCrLf
        'sql += " ,cc.classcname2 班級名稱" & vbCrLf
        'sql += " ,CONVERT(varchar, cc.stdate, 111) 開訓日期" & vbCrLf
        'sql += " ,CONVERT(varchar, cc.ftdate, 111) 結訓日期" & vbCrLf
        'sql += " ,cc.tnum 核定人數" & vbCrLf
        'sql += " ,'0' 遞補人數" & vbCrLf
        'sql += " ,isnull(e.enterNum,0)  報名人數" & vbCrLf

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " select cc.OCID" & vbCrLf '班級名稱
        sql += " ,cc.DISTNAME " & vbCrLf '轄區
        sql += " ,cc.ORGNAME " & vbCrLf '單位名稱
        sql += " ,cc.CLASSCNAME2 " & vbCrLf '班級名稱
        sql += " ,CONVERT(varchar, cc.STDATE, 111) STDATE" & vbCrLf '開訓日期
        sql += " ,CONVERT(varchar, cc.FTDATE, 111) FTDATE" & vbCrLf '結訓日期
        sql += " ,cc.TNUM" & vbCrLf ' 核定人數
        sql += " ,0 NOENTNUM" & vbCrLf '遞補人數
        sql += " ,0 ENTERNUM" & vbCrLf '報名人數
        sql += " FROM dbo.VIEW2 cc " & vbCrLf
        'sql += " LEFT JOIN VIEW_ZIPNAME iz3 on cc.TaddressZip = iz3.ZipCode" & vbCrLf
        sql += " JOIN dbo.VIEW_RIDNAME ar on ar.RID =cc.RID " & vbCrLf
        'sql += " JOIN Org_OrgPlanInfo op on op.RSID = ar.RSID" & vbCrLf
        sql += " LEFT JOIN dbo.VIEW_ZIPNAME iz2 on iz2.ZipCode=ar.ZIPCODE" & vbCrLf
        'sql += " left join (" & vbCrLf
        'sql += " 	select a.OCID1 " & vbCrLf
        'sql += " 	,count(1) enterNum " & vbCrLf
        'sql += " 	from v_EnterType2a a " & vbCrLf
        'sql += " 	group by a.OCID1" & vbCrLf
        'sql += " ) e on e.OCID1 =cc.ocid " & vbCrLf
        sql += " where 1=1" & vbCrLf
        sql += " and cc.TPlanID=@TPlanID " & vbCrLf '計畫
        sql += " and cc.years=@years " & vbCrLf '年度

        parms.Add("TPlanID", sm.UserInfo.TPlanID) '計畫
        parms.Add("years", yearlist.SelectedValue) '年度

        '轄區
        If gDistIDVal <> "" Then
            sql += " and cc.Distid in (" & gDistIDVal & ")" & vbCrLf
        End If
        '辦訓地縣市 
        If gTCityCode2 <> "" Then
            sql += " and CC.CTID in (" & gTCityCode2 & ")" & vbCrLf
        End If
        '立案地縣市
        If gOCityCode2 <> "" Then
            sql += " and iz2.CTID in (" & gOCityCode2 & ")" & vbCrLf
        End If
        If SDate1.Text <> "" Then
            'sql += " and cc.STDate >= " & TIMS.to_date(SDate1.Text) & vbCrLf
            sql += " and cc.STDate >= @STDate1" & vbCrLf
            parms.Add("STDate1", SDate1.Text)
        End If
        If SDate2.Text <> "" Then
            'sql += " and cc.STDate <= " & TIMS.to_date(SDate2.Text) & vbCrLf '" & SDate2.Text & "'" & vbCrLf
            sql += " and cc.STDate <= @STDate2" & vbCrLf '" & SDate2.Text & "'" & vbCrLf
            parms.Add("STDate2", SDate2.Text)
        End If
        If EDate1.Text <> "" Then
            'sql += " and cc.FTDate >= " & TIMS.to_date(EDate1.Text) & vbCrLf '" & EDate1.Text & "'" & vbCrLf
            sql += " and cc.FTDate >= @FTDate1 " & vbCrLf '" & EDate1.Text & "'" & vbCrLf
            parms.Add("FTDate1", EDate1.Text)
        End If
        If EDate2.Text <> "" Then
            'sql += " and cc.FTDate <= " & TIMS.to_date(EDate2.Text) & vbCrLf '" & EDate2.Text & "'" & vbCrLf
            sql += " and cc.FTDate <= @FTDate2 " & vbCrLf '" & EDate2.Text & "'" & vbCrLf
            parms.Add("FTDate2", EDate2.Text)
        End If

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objConn, parms)
        'dt = sUtl_ChangeDt(dt)
        'dt.Columns.Remove("ocid")

        msg.Text = "查無資料"
        PageControler1.Visible = False
        DataGrid1.Visible = False
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            PageControler1.Visible = True
            DataGrid1.Visible = True

            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
        End If

        'Try
        'Catch ex As Exception
        '    Common.MessageBox(Me, ex.ToString)
        '    Exit Sub
        'End Try

    End Sub

    '匯出 明細資料
    Sub sExport1()
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        DataGrid1.AllowPaging = False '關閉分頁
        DataGrid1.EnableViewState = False  '把ViewState給關了

        Call Search1()

        If DataGrid1.Visible = False OrElse msg.Text <> "" Then
            Common.MessageBox(Page, msg.Text)
            Exit Sub
        End If

        Dim sFileName1 As String = "遞補人數統計表"

        '勞保勾稽查詢
        'sFileName = HttpUtility.UrlEncode(Cst_xlsFileName, System.Text.Encoding.UTF8)

        ''套CSS值
        Dim strSTYLE As String = ""
        strSTYLE &= ("<style>")
        strSTYLE &= ("td{mso-number-format:""\@"";}")
        strSTYLE &= ("</style>")

        DataGrid1.AllowPaging = False '關閉分頁
        DataGrid1.EnableViewState = False  '把ViewState給關了

        Dim objStringWriter As New System.IO.StringWriter
        Dim objHtmlTextWriter As New System.Web.UI.HtmlTextWriter(objStringWriter)
        Div1.RenderControl(objHtmlTextWriter)

        Dim strHTML As String = ""
        strHTML &= (TIMS.sUtl_AntiXss(Convert.ToString(objStringWriter)))

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", strHTML)
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)
        DataGrid1.AllowPaging = True '開啟分頁
        TIMS.Utl_RespWriteEnd(Me, objConn, "") 'Response.End()
    End Sub

    '查詢 明細資料
    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        '查詢
        Call Search1()
    End Sub

    '覆寫，不執行 MyBase.VerifyRenderingInServerForm 方法，解決執行 RenderControl 產生的錯誤   
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)

    End Sub

    Private Sub btnExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExport.Click
        Call sExport1()
    End Sub


    Private Sub DataGrid1_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                'Sql += " ,'0' 遞補人數" & vbCrLf
                'sql += " ,isnull(e.enterNum,0)  報名人數" & vbCrLf
                Dim sql As String = ""
                sql = "SELECT dbo.FN_GET_ENTERCNT(" & drv("OCID") & ") ENTERNUM"
                Dim dr1 As DataRow = DbAccess.GetOneRow(sql, objConn)
                If dr1 IsNot Nothing Then
                    e.Item.Cells(7).Text = If(Convert.ToString(dr1("ENTERNUM")) <> "", dr1("ENTERNUM"), 0)
                    Dim iENTERNUM As Integer = If(Convert.ToString(dr1("ENTERNUM")) <> "", dr1("ENTERNUM"), 0)
                    '報名人數 大於 核定人數 才處理
                    If iENTERNUM > Val(drv("TNUM")) Then
                        '找不到該學員的報名資料(前 N 名錄取人員)，即為遞補人數 '求遞補人數
                        sql = "SELECT dbo.FN_GET_NOENTNUM(" & drv("OCID") & "," & drv("TNUM") & ") NOENTNUM"
                        Dim dr2 As DataRow = DbAccess.GetOneRow(sql, objConn)
                        e.Item.Cells(6).Text = If(Convert.ToString(dr2("NOENTNUM")) <> "", dr2("NOENTNUM"), 0)
                    End If
                End If
                'e.Item.Cells(6).Text=NOENTNUM
                'e.Item.Cells(7).Text ENTERNUM
        End Select
    End Sub
End Class
