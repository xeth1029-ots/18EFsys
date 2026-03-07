Partial Class CP_04_001
    Inherits AuthBasePage

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

        If Not Page.IsPostBack Then
            Dim dt As DataTable
            Dim dr As DataRow
            Dim sqlstr As String

            yearlist = TIMS.GetSyear(yearlist)

            sqlstr = "select PlanID,Years,tplanid from ID_Plan where PlanID= '" & sm.UserInfo.PlanID & "' order by years,tplanid"
            dr = DbAccess.GetOneRow(sqlstr, objconn)
            Common.SetListItem(yearlist, dr("Years"))
            'Me.yearlist.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))

            sqlstr = ""
            sqlstr &= " SELECT NAME,DISTID FROM ID_DISTRICT"
            sqlstr &= " WHERE DISTID NOT IN ('002','007','008')"
            sqlstr &= " ORDER BY DISTID"
            dt = DbAccess.GetDataTable(sqlstr, objconn)

            Me.DistrictList.DataSource = dt
            Me.DistrictList.DataTextField = "Name"
            Me.DistrictList.DataValueField = "DistID"
            Me.DistrictList.DataBind()
            Me.DistrictList.Items.Insert(0, New ListItem("全部", ""))

            '計畫
            PlanList = TIMS.Get_TPlan(PlanList, , 1, "Y")

            '縣市
            CityList = TIMS.Get_CityName(CityList, TIMS.dtNothing)

        End If

        '選擇全部轄區
        DistrictList.Attributes("onclick") = "SelectAll('DistrictList','DistHidden');"

        '選擇全部縣市
        Me.CityList.Attributes("onclick") = "SelectAll('CityList','CityHidden');"

        '選擇全部訓練計畫
        PlanList.Attributes("onclick") = "SelectAll('PlanList','TPlanHidden');"

    End Sub

    '明細查詢
    Private Sub bt_search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_search.Click


        '選擇轄區
        Dim itemstr As String
        itemstr = ""
        For Each objitem As ListItem In Me.DistrictList.Items
            If objitem.Selected = True Then
                If itemstr <> "" Then itemstr += ","
                itemstr += "'" & objitem.Value.ToString & "'"
            End If
        Next

        '選擇縣市
        Dim itemcity As String
        itemcity = ""
        For Each objitem As ListItem In Me.CityList.Items
            If objitem.Selected = True Then
                If itemcity <> "" Then itemcity += ","
                itemcity += "'" & objitem.Value.ToString & "'"
            End If
        Next

        '選擇訓練計畫
        Dim itemplan As String
        itemplan = ""
        For Each objitem As ListItem In Me.PlanList.Items
            If objitem.Selected = True Then
                If itemplan <> "" Then itemplan += ","
                itemplan += "'" & objitem.Value.ToString & "'"
            End If
        Next

        'SQL查詢用
        Session("itemstr") = itemstr
        Session("itemplan") = itemplan
        Session("itemcity") = itemcity

        'Response.Redirect("CP_04_001_01.aspx?yearlist=" & Me.yearlist.SelectedValue)
        Dim url1 As String = "CP_04_001_01.aspx?ID=" & Request("ID") & "&yearlist=" & Me.yearlist.SelectedValue
        Call TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    '重新設定
    Private Sub bt_reset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_reset.Click
        'Reset
        Me.yearlist.SelectedIndex = 1
        'Dim i As Short
        For i As Short = 0 To Me.DistrictList.Items.Count - 1
            Me.DistrictList.Items(i).Selected = False
        Next
        For i As Short = 0 To Me.CityList.Items.Count - 1
            Me.CityList.Items(i).Selected = False
        Next
        For i As Short = 0 To Me.PlanList.Items.Count - 1
            Me.PlanList.Items(i).Selected = False
        Next

    End Sub

    '統計查詢
    Private Sub bt_search1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_search1.Click
        Dim DistID, ICity, TPlanID As String
        Dim newDistID, newICity, newTPlanID As String


        '選擇轄區
        Dim itemstr As String
        itemstr = ""
        For Each objitem As ListItem In Me.DistrictList.Items
            If objitem.Selected = True Then
                If itemstr <> "" Then itemstr += ","
                itemstr += "'" & objitem.Value.ToString & "'"
            End If
        Next
        '報表要用的轄區參數
        DistID = ""
        For i As Integer = 1 To Me.DistrictList.Items.Count - 1
            If Me.DistrictList.Items(i).Selected Then
                If DistID <> "" Then DistID += ","
                DistID += "\'" & Me.DistrictList.Items(i).Value & "\'"
            End If
        Next

        '選擇縣市
        Dim itemcity As String
        itemcity = ""
        For Each objitem As ListItem In Me.CityList.Items
            If objitem.Selected = True Then
                If itemcity <> "" Then itemcity += ","
                itemcity += "'" & objitem.Value.ToString & "'"
            End If
        Next
        '報表要用的縣市參數
        ICity = ""
        For i As Integer = 1 To Me.CityList.Items.Count - 1
            If Me.CityList.Items(i).Selected Then
                If ICity <> "" Then ICity += ","
                ICity += "\'" & Me.CityList.Items(i).Value & "\'"
            End If
        Next

        '選擇訓練計畫
        Dim itemplan As String
        itemplan = ""
        For Each objitem As ListItem In Me.PlanList.Items
            If objitem.Selected = True Then
                If itemplan <> "" Then itemplan += ","
                itemplan += "'" & objitem.Value.ToString & "'"
            End If
        Next
        '報表要用的訓練計畫參數
        TPlanID = ""
        For i As Integer = 1 To Me.PlanList.Items.Count - 1
            If Me.PlanList.Items(i).Selected Then
                If TPlanID <> "" Then TPlanID += ","
                TPlanID += "\'" & Me.PlanList.Items(i).Value & "\'"
            End If
        Next
        newDistID = DistID
        newICity = ICity
        newTPlanID = TPlanID

        'SQL查詢用
        Session("itemstr") = itemstr
        Session("itemplan") = itemplan
        Session("itemcity") = itemcity
        '報表列印用
        Session("newDistID") = newDistID
        Session("newICity") = newICity
        Session("newTPlanID") = newTPlanID

        'Response.Redirect("CP_04_001_02.aspx?yearlist=" & Me.yearlist.SelectedValue)
        Dim url1 As String = "CP_04_001_02.aspx?ID=" & Request("ID") & "&yearlist=" & Me.yearlist.SelectedValue
        Call TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    '資料匯出
    Private Sub bt_export_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_export.Click
        Dim dt As DataTable
        Dim num As Integer = 0

        Try
            'conn.ConnectionString = connString
            'conn.Open()
            'TIMS.OpenDbConn(objconn)
            '選擇轄區
            'Dim objitem As ListItem
            Dim itemstr As String
            itemstr = ""
            For Each objitem As ListItem In Me.DistrictList.Items
                If objitem.Selected = True Then
                    If itemstr <> "" Then itemstr += ","
                    itemstr += "'" & objitem.Value.ToString & "'"
                End If
            Next

            '選擇縣市
            Dim itemcity As String
            itemcity = ""
            For Each objitem As ListItem In Me.CityList.Items
                If objitem.Selected = True Then
                    If itemcity <> "" Then itemcity += ","
                    itemcity += "'" & objitem.Value.ToString & "'"
                End If
            Next

            '選擇訓練計畫
            Dim itemplan As String
            itemplan = ""
            For Each objitem As ListItem In Me.PlanList.Items
                If objitem.Selected = True Then
                    If itemplan <> "" Then itemplan += ","
                    itemplan += "'" & objitem.Value.ToString & "'"
                End If
            Next

            Dim DistStr As String = ""
            Dim CityStr As String = ""
            Dim TPlanStr As String = ""
            Dim YearsStr As String = ""

            '選擇轄區
            If itemstr <> "" Then
                DistStr = " and DistID IN (" & itemstr & ") "
            End If
            '選擇縣市
            If itemcity <> "" Then
                CityStr = " and g.CTID IN (" & itemcity & ") "
            End If
            '選擇訓練計畫
            If itemplan <> "" Then
                TPlanStr = " and TPlanID IN (" & itemplan & ") "
            End If
            '選擇年度
            'Dim yearlist As String = Me.yearlist.SelectedValue
            If Me.yearlist.SelectedValue <> "" Then
                YearsStr = " and Years='" & Me.yearlist.SelectedValue & "' "
            End If


            '以轄區、訓練計畫做排序
            Dim sqlstr As String = ""
            sqlstr = "" & vbCrLf
            sqlstr += " SELECT a.RID,b.OrgName,b.ComIDNO" & vbCrLf
            sqlstr += " ,c.Name as DistName,c.DistID" & vbCrLf
            sqlstr += " ,e.PlanName" & vbCrLf
            sqlstr += " ,(g.CTName+f.ZipName+ replace(replace(h.Address,g.CTName,''),f.ZipName,'')) as Address" & vbCrLf
            sqlstr += " ,h.ContactName,h.Phone,h.ContactEmail" & vbCrLf
            sqlstr += " FROM " & vbCrLf 'Auth_Relship a
            sqlstr += " (SELECT * FROM Auth_Relship WHERE (PlanID IN (SELECT PlanID FROM ID_Plan WHERE 1=1" & TPlanStr & DistStr & YearsStr & ")) or (PlanID=0 " & DistStr & ")) a "
            sqlstr += " JOIN Org_OrgInfo b ON a.OrgID = b.OrgID" & vbCrLf
            sqlstr += " JOIN Org_OrgPlanInfo h ON a.RSID = h.RSID" & vbCrLf
            sqlstr += " JOIN ID_ZIP f ON f.ZipCode = h.ZipCode" & vbCrLf
            sqlstr += " JOIN ID_City g ON f.CTID = g.CTID " & vbCrLf
            If CityStr <> "" Then
                sqlstr += CityStr & vbCrLf
            End If
            sqlstr += " JOIN ID_District c ON a.DistID = c.DistID" & vbCrLf
            sqlstr += " LEFT JOIN ID_Plan d ON a.PlanID = d.PlanID" & vbCrLf
            sqlstr += " LEFT JOIN Key_Plan e ON d.TPlanID = e.TPlanID" & vbCrLf
            sqlstr += " Order By a.DistID,e.PlanName,a.RID"
            '以轄區、訓練計畫做排序
            dt = DbAccess.GetDataTable(sqlstr, objconn)

            Response.AddHeader("content-disposition", "attachment; filename=" & HttpUtility.UrlEncode("訓練機構資料", System.Text.Encoding.UTF8) & ".xls")
            Response.ContentType = "Application/octet-stream"
            Response.ContentEncoding = System.Text.Encoding.GetEncoding("Big5")
            Dim ExportStr As String             '建立輸出文字
            ' ExportStr = "序號" & vbTab & 
            ExportStr = "計畫名稱" & vbTab & "機構名稱" & vbTab & "地址" & vbTab & "聯絡人" & vbTab & "電話" & vbTab & "E-Mail" & vbTab

            ExportStr += vbCrLf
            Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))
            '建立資料面
            For Each dr As DataRow In dt.Rows
                ExportStr = ""
                num = num + 1
                'ExportStr = ExportStr & num & vbTab                               '序號
                ExportStr = ExportStr & Convert.ToString(dr("PlanName")) & vbTab  '計畫名稱
                ExportStr = ExportStr & Convert.ToString(dr("OrgName")) & vbTab   '機構名稱
                ExportStr = ExportStr & Convert.ToString(dr("Address")) & vbTab   '地址
                ExportStr = ExportStr & Convert.ToString(dr("ContactName")) & vbTab    '聯絡人
                ExportStr = ExportStr & Convert.ToString(dr("Phone")) & vbTab          '電話
                ExportStr = ExportStr & Convert.ToString(dr("ContactEmail")) & vbTab   'E-Mail
                ExportStr += vbCrLf
                Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))
            Next
            Response.End()

            'If Not da Is Nothing Then da.Dispose()
            If Not dt Is Nothing Then dt.Dispose()
        Catch ex As Exception
            TIMS.RegisterStartupScript(Me, "errMsg", "<script>alert('【發生錯誤】:\n" & ex.ToString.Replace("'", "\'").Replace(Convert.ToChar(10), "\n").Replace(Convert.ToChar(13), "") & "');</script>")

            Throw ex
            'Finally
        End Try

    End Sub
End Class
