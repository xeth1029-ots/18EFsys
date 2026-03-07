Partial Class CP_04_005
    Inherits AuthBasePage

    'CP_04_005.jrxml '縣市
    'CP_04_005_1.jrxml  '鄉鎮市區
    'CP_04_005*.jrxml
    Const cst_printFN1 As String = "CP_04_005"
    Const cst_printFN2 As String = "CP_04_005_1"

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
            '檢查日期格式
            Me.SSTDate.Attributes("onchange") = "check_date();"
            Me.ESTDate.Attributes("onchange") = "check_date();"
            Me.SFTDate.Attributes("onchange") = "check_date();"
            Me.EFTDate.Attributes("onchange") = "check_date();"

            print.Attributes("onclick") = "return search();"

            Call CreateItem()

            'RID = sm.UserInfo.RID
            Dim NewRID As String = Left(Convert.ToString(sm.UserInfo.RID), 1)
            Dist.Style("display") = "inline"
            If NewRID <> "A" Then         '判斷是否是署(局)還是分署(中心),若是署(局)就顯示轄區選項
                Dist.Style("display") = "none"
            End If

            '選擇全部轄區
            DistID.Attributes("onclick") = "SelectAll('DistID','DistHidden');"

            '選擇全部縣市
            CityID.Attributes("onclick") = "SelectAll('CityID','CityHidden');"

            '選擇全部訓練計畫
            TPlanID.Attributes("onclick") = "SelectAll('TPlanID','TPlanHidden');"
        End If
    End Sub

    Sub CreateItem()
        'Dim dt As DataTable
        'Dim dr As DataRow
        'Dim sqlstr As String
        DistID = TIMS.Get_DistID(DistID) '轄區
        DistID.Items.Insert(0, New ListItem("全部", ""))
        TPlanID = TIMS.Get_TPlan(TPlanID, , 1, "Y")
        '縣市
        CityID = TIMS.Get_CityName(CityID, TIMS.dtNothing)
        'CityList.Items(0).Selected = True
    End Sub

    '取得外部資訊
    Sub Get_SearchObjValue(ByRef MyPage As Page, ByRef outValue As String)
        'Dim i, j As Integer
        outValue = ""
        Dim RID1 As String = ""
        Dim NewRID1 As String = ""
        '選擇轄區
        RID1 = sm.UserInfo.RID
        NewRID1 = Left(RID1, 1)
        'Dim MeDistID As System.Web.UI.WebControls.CheckBoxList
        'Dim MeCityID As System.Web.UI.WebControls.CheckBoxList
        'Dim MeTPlanID As System.Web.UI.WebControls.CheckBoxList
        'MeDistID = CType(MyPage.FindControl("DistID"), System.Web.UI.WebControls.CheckBoxList)
        'MeCityID = CType(MyPage.FindControl("CityID"), System.Web.UI.WebControls.CheckBoxList)
        'MeTPlanID = CType(MyPage.FindControl("TPlanID"), System.Web.UI.WebControls.CheckBoxList)


        '報表要用的轄區參數
        Dim DistID1 As String = ""
        'Dim DistName As String = ""
        Select Case NewRID1
            Case "A"
                For i As Integer = 1 To DistID.Items.Count - 1
                    If DistID.Items(i).Selected AndAlso DistID.Items(i).Value <> "" Then
                        If DistID1 <> "" Then DistID1 &= ","
                        DistID1 &= Convert.ToString("\'" & DistID.Items(i).Value & "\'")
                        'If DistName <> "" Then DistName &= ","
                        'DistName &= Convert.ToString(MeDistID.Items(i).Text)
                    End If
                Next

            Case "B"
                DistID1 = Convert.ToString("\'" & DistID.Items(2).Value & "\'")
                'DistName = Convert.ToString(MeDistID.Items(2).Text)

            Case "C"
                DistID1 = Convert.ToString("\'" & DistID.Items(3).Value & "\'")
                'DistName = Convert.ToString(MeDistID.Items(3).Text)

            Case "D"
                DistID1 = Convert.ToString("\'" & DistID.Items(4).Value & "\'")
                'DistName = Convert.ToString(MeDistID.Items(4).Text)

            Case "E"
                DistID1 = Convert.ToString("\'" & DistID.Items(5).Value & "\'")
                'DistName = Convert.ToString(MeDistID.Items(5).Text)

            Case "F"
                DistID1 = Convert.ToString("\'" & DistID.Items(6).Value & "\'")
                'DistName = Convert.ToString(MeDistID.Items(6).Text)

            Case "G"
                DistID1 = Convert.ToString("\'" & DistID.Items(7).Value & "\'")
                'DistName = Convert.ToString(MeDistID.Items(7).Text)

        End Select

        '報表要用的縣市
        Dim ICityID1 As String = ""
        'Dim ICityName As String = ""
        For i As Integer = 1 To CityID.Items.Count - 1
            If CityID.Items(i).Selected Then
                If ICityID1 <> "" Then ICityID1 &= ","
                ICityID1 &= Convert.ToString("\'" & CityID.Items(i).Value & "\'")
                'If ICityName <> "" Then ICityName &= ","
                'ICityName &= Convert.ToString(MeCityID.Items(i).Text)
            End If
        Next


        '報表要用的訓練計畫參數
        Dim TPlanID1 As String = ""
        'Dim TPlanName As String = ""
        For i As Integer = 1 To TPlanID.Items.Count - 1
            If TPlanID.Items(i).Selected Then
                If TPlanID1 <> "" Then TPlanID1 &= ","
                TPlanID1 &= Convert.ToString("\'" & TPlanID.Items(i).Value & "\'")
                'If TPlanName <> "" Then TPlanName &= ","
                'TPlanName &= Convert.ToString(MeTPlanID.Items(i).Text)
            End If
        Next

        outValue = ""
        outValue &= "&DistID1=" & DistID1
        outValue &= "&ICityID1=" & ICityID1
        outValue &= "&TPlanID1=" & TPlanID1
        'outValue &= "&DistName=" & DistName
        'outValue &= "&ICityName=" & ICityName
        'outValue &= "&TPlanName=" & TPlanName
    End Sub

    Private Sub print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles print.Click
        Dim outValue As String = ""
        Call Get_SearchObjValue(Me, outValue) 'out@outValue

        Dim DistID1 As String = TIMS.GetMyValue(outValue, "DistID1")
        Dim ICityID1 As String = TIMS.GetMyValue(outValue, "ICityID1")
        Dim TPlanID1 As String = TIMS.GetMyValue(outValue, "TPlanID1")

        Dim str_sc As String = ""
        str_sc &= "SSTDate=" & SSTDate.Text
        str_sc &= "&ESTDate=" & ESTDate.Text
        str_sc &= "&SFTDate=" & SFTDate.Text
        str_sc &= "&EFTDate=" & EFTDate.Text
        If DistID1 <> "" Then str_sc &= "&DistID=" & DistID1
        If ICityID1 <> "" Then str_sc &= "&ICityID=" & ICityID1
        If TPlanID1 <> "" Then str_sc &= "&TPlanID=" & TPlanID1

        Select Case PrintStaus.SelectedValue
            Case "1" '縣市
                TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, str_sc)

            Case Else '鄉鎮市區
                TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN2, str_sc)

        End Select

    End Sub

    Private Sub bt_reset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_reset.Click
        'Me.yearlist.SelectedIndex = 1
        'Dim i As Short
        For i As Short = 0 To Me.DistID.Items.Count - 1
            Me.DistID.Items(i).Selected = False
        Next
        For i As Short = 0 To Me.CityID.Items.Count - 1
            Me.CityID.Items(i).Selected = False
        Next
        For i As Short = 0 To Me.TPlanID.Items.Count - 1
            Me.TPlanID.Items(i).Selected = False
        Next

        Me.SSTDate.Text = ""
        Me.ESTDate.Text = ""
        Me.SFTDate.Text = ""
        Me.EFTDate.Text = ""
        'Me.OrgName.Text = ""
        'Me.ClassCName.Text = ""

        Me.PrintStaus.SelectedIndex = 0
        'Me.TMID.SelectedIndex = 0

        'For i = 0 To Me.BudgetList.Items.Count - 1
        '    Me.BudgetList.Items(i).Selected = False
        'Next
    End Sub

    '匯出SUB (SQL)
    Private Sub ExpRpt(ByVal da As SqlDataAdapter)
        Dim outValue As String = ""
        Call Get_SearchObjValue(Me, outValue) 'out@outValue

        Dim DistID1 As String = TIMS.GetMyValue(outValue, "DistID1")
        Dim ICityID1 As String = TIMS.GetMyValue(outValue, "ICityID1")
        Dim TPlanID1 As String = TIMS.GetMyValue(outValue, "TPlanID1")

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " select ip.planname" & vbCrLf
        sql &= " ,oo.orgname" & vbCrLf
        sql &= " ,cc.classcname" & vbCrLf
        sql &= " ,cc.cycltype" & vbCrLf
        sql &= " ,vz.zipcode" & vbCrLf
        sql &= " ,vz.zipname" & vbCrLf
        sql &= " ,cc.Taddress" & vbCrLf
        sql &= " FROM Class_ClassInfo cc " & vbCrLf
        sql &= " JOIN Plan_PlanInfo pp ON pp.PlanID = cc.PlanID and cc.ComIDNO = pp.ComIDNO and cc.SeqNO = pp.SeqNO " & vbCrLf
        sql &= " JOIN view_plan ip on ip.planid =pp.planid" & vbCrLf
        sql &= " JOIN org_orginfo oo on oo.comidno=cc.comidno" & vbCrLf
        sql &= " LEFT JOIN VIEW_ZIPNAME vz on vz.zipcode = cc.TaddressZip" & vbCrLf
        sql &= " where 1=1" & vbCrLf
        sql &= " AND cc.notopen ='N'" & vbCrLf
        sql &= " AND cc.issuccess = 'Y'" & vbCrLf

        If Me.SSTDate.Text <> "" Then
            sql &= " and cc.STDate>= " & TIMS.To_date(Me.SSTDate.Text) & vbCrLf
        End If
        If Me.ESTDate.Text <> "" Then
            sql &= " and cc.STDate<= " & TIMS.To_date(Me.ESTDate.Text) & vbCrLf '" & Me.ESTDate.Text & "'" & vbCrLf
        End If
        If Me.SFTDate.Text <> "" Then
            sql &= " and cc.FTDate>= " & TIMS.To_date(Me.SFTDate.Text) & vbCrLf '" & Me.SFTDate.Text & "'" & vbCrLf
        End If
        If Me.EFTDate.Text <> "" Then
            sql &= " and cc.FTDate<= " & TIMS.To_date(Me.EFTDate.Text) & vbCrLf '" & Me.EFTDate.Text & "'" & vbCrLf
        End If
        If DistID1 <> "" Then
            sql &= " and ip.DistID IN (" & DistID1.Replace("\", "") & ") " & vbCrLf
        End If
        If ICityID1 <> "" Then
            sql &= " and vz.CTID IN (" & ICityID1.Replace("\", "") & ") " & vbCrLf
        End If
        If TPlanID1 <> "" Then
            sql &= " and ip.TPlanID IN (" & TPlanID1.Replace("\", "") & ") " & vbCrLf
        End If

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)
        'da.SelectCommand.CommandText = Sql
        'da.SelectCommand.Parameters.Clear()
        'dt = New DataTable
        'da.Fill(dt)

        Dim ExportStr As String = ""
        Response.AddHeader("content-disposition", "attachment; filename=" & HttpUtility.UrlEncode("上課地點匯出表", System.Text.Encoding.UTF8) & ".xls")
        'Response.ContentType = "Application/octet-stream"
        Response.ContentEncoding = System.Text.Encoding.GetEncoding("Big5")
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("UTF-8")
        'Response.ContentType = "application/vnd.ms-excel " '內容型態設為Excel
        '文件內容指定為Excel
        'Response.ContentType = "application/ms-excel;charset=utf-8"
        Response.ContentType = "application/ms-excel"
        'Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=UTF-8>")
        Common.RespWrite(Me, "<html>")
        Common.RespWrite(Me, "<head>")
        Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=BIG5>")
        '<head><meta http-equiv='Content-Type' content='text/html; charset=utf-8'></head>
        ''套CSS值
        Common.RespWrite(Me, "<style>")
        Common.RespWrite(Me, "td{mso-number-format:""\@"";}")
        Common.RespWrite(Me, ".noDecFormat{mso-number-format:""0"";}")
        'mso-number-format:"0" 
        Common.RespWrite(Me, "</style>")
        Common.RespWrite(Me, "</head>")

        Common.RespWrite(Me, "<body>")
        Common.RespWrite(Me, "<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")

        ExportStr = ""
        ExportStr &= "<tr>"
        ExportStr &= "<td>訓練計畫</td>" & vbTab
        ExportStr &= "<td>訓練單位</td>" & vbTab
        ExportStr &= "<td>班別名稱</td>" & vbTab
        ExportStr &= "<td>期別</td>" & vbTab
        ExportStr &= "<td>上課郵遞區號</td>" & vbTab
        ExportStr &= "<td>上課縣市</td>" & vbTab
        ExportStr &= "<td>上課地址</td>" & vbTab
        ExportStr &= "</tr>" & vbCrLf
        Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))

        '建立資料面
        For Each dr As DataRow In dt.Rows
            ExportStr = ""
            ExportStr &= "<tr>"
            ExportStr &= "<td>" & Convert.ToString(dr("planname")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("orgname")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("classcname")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("cycltype")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("zipcode")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("zipname")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("Taddress")) & "</td>" & vbTab
            ExportStr &= "</tr>" & vbCrLf
            Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))
        Next
        Common.RespWrite(Me, "</table>")
        Common.RespWrite(Me, "</body>")

    End Sub

    '覆寫，不執行 MyBase.VerifyRenderingInServerForm 方法，解決執行 RenderControl 產生的錯誤   
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)

    End Sub

    '匯出
    Private Sub bt_export_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_export.Click
        Dim tConn As SqlConnection = DbAccess.GetConnection()
        Dim okFlag As Boolean = False
        'Dim conn As SqlConnection
        Try
            'Call TIMS.TestDbConn(Me, tConn)
            Call TIMS.OpenDbConn(tConn)

            Dim da As New SqlDataAdapter
            da.SelectCommand = New SqlCommand
            da.SelectCommand.Connection = tConn
            da.SelectCommand.CommandTimeout = 100

            ExpRpt(da) '匯出SUB'SQL

            okFlag = True

            da.Dispose()
        Catch ex As Exception
            'If conn.State = ConnectionState.Open Then conn.Close()
            Common.MessageBox(Me.Page, "發生錯誤:" & vbCrLf & ex.ToString)
            'Me.Page.RegisterStartupScript("Errmsg", "<script>alert('發生錯誤:\n" & ex.ToString.Replace("'", "\'").Replace(Convert.ToChar(10), "\n").Replace(Convert.ToChar(13), "") & "');</script>")
        End Try
        Call TIMS.CloseDbConn(tConn)

        If okFlag Then
            Response.End()
        End If
    End Sub
End Class
