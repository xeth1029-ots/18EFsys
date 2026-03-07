Partial Class TR_04_007_R
    Inherits AuthBasePage

    'TR_04_007_R
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
            Syear = TIMS.GetSyear(Syear)
            Common.SetListItem(Syear, Now.Year)
            DistID = TIMS.Get_DistID(DistID)

            TPlanID = TIMS.Get_TPlan(TPlanID, , 1)

            DistID.SelectedValue = sm.UserInfo.DistID
            TPlanID.SelectedValue = sm.UserInfo.TPlanID
            OCID.Items.Add("請選擇機構")
            Page.RegisterStartupScript("", "<script>GetMode();</script>")
        End If

        OCID.Attributes("onchange") = "if(this.selectedIndex!=0){document.form1.OCIDValue.value=this.value;}else{document.form1.OCIDValue.value='';}"
        DistID.Attributes("onchange") = "GetMode();"
        TPlanID.Attributes("onchange") = "GetMode();"
        Button1.Attributes("onclick") = "return search();"
        Button2.Style("display") = "none"
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If RIDValue.Value <> "" Then
            OCID.Items.Clear()
            'Dim sql As String
            'Dim dt As DataTable
            'Dim dr As DataRow

            'sql = "SELECT * FROM "
            'sql += "(SELECT * FROM Class_ClassInfo WHERE RID='" & RIDValue.Value & "') a "
            'sql += "JOIN (SELECT * FROM ID_Plan WHERE TPlanID='" & TPlanID.SelectedValue & "' and distid='" & DistID.SelectedValue & "') b ON a.PlanID=b.PlanID "
            'dt = DbAccess.GetDataTable(sql, objconn)

            Dim sql As String = ""
            sql = "" & vbCrLf
            sql += " select cc.ClassCName" & vbCrLf
            sql += " ,cc.CyclType" & vbCrLf
            sql += " ,cc.LevelType" & vbCrLf
            sql += " ,cc.OCID" & vbCrLf
            sql += " ,cc.STDate" & vbCrLf
            sql += " ,cc.FTDate" & vbCrLf
            sql += " from Class_ClassInfo cc" & vbCrLf
            sql += " join id_plan ip on ip.planid =cc.planid " & vbCrLf
            sql += " where 1=1" & vbCrLf
            sql += " and cc.RID= @RID" & vbCrLf
            sql += " and ip.TPlanID= @TPlanID" & vbCrLf
            sql += " and ip.DistID= @DistID" & vbCrLf
            sql += " ORDER BY cc.STDate "
            Call TIMS.OpenDbConn(objconn)
            Dim dt As New DataTable
            Dim oCmd As New SqlCommand(sql, objconn)
            With oCmd
                .Parameters.Clear()
                .Parameters.Add("RID", SqlDbType.VarChar).Value = RIDValue.Value
                .Parameters.Add("TPlanID", SqlDbType.VarChar).Value = TPlanID.SelectedValue
                .Parameters.Add("DistID", SqlDbType.VarChar).Value = DistID.SelectedValue
                dt.Load(.ExecuteReader())
            End With
            If dt.Rows.Count = 0 Then
                OCID.Items.Insert(0, New ListItem("此計畫、機構底下沒有任何班級", ""))
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

                    OCID.Items.Add(New ListItem(ClassName, dr("OCID")))
                Next
                OCID.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
            End If
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim stitle As String = ""
        Dim etitle As String = ""
        'Dim cGuid As String =   ReportQuery.GetGuid(Page)
        'Dim Url As String =   ReportQuery.GetUrl(Page)
        Dim strScript As String = ""
        If STDate1.Text <> "" Or STDate2.Text <> "" Then
            stitle = STDate1.Text + " ~ " + STDate2.Text
        End If
        If FTDate1.Text <> "" Or FTDate2.Text <> "" Then
            etitle = FTDate1.Text + " ~ " + FTDate2.Text
        End If
        'strScript = "<script language=""javascript"">" + vbCrLf
        'strScript += "window.open('" & Url & "GUID=" + cGuid + "&SQ_AutoLogout=true&sys=TR&filename=TR_04_007_R&path=TIMS&TPlanID=" & TPlanID.SelectedValue & "&Years=" & Syear.SelectedValue & "&DistID=" & DistID.SelectedValue & "&RID=" & RIDValue.Value & "&OCID=" & OCIDValue.Value & "&STDate1=" & STDate1.Text & "&STDate2=" & STDate2.Text & "&FTDate1=" & FTDate1.Text & "&FTDate2=" & FTDate2.Text & "&stitle=" & stitle & "&etitle=" & etitle
        'If CPoint.SelectedItem.Value = 2 Then
        '    strScript += "&Kind1=1"
        'End If
        'If CPoint.SelectedItem.Value = 3 Then
        '    strScript += "&Kind2=1"
        'End If
        'strScript += "');" + vbCrLf
        'strScript += "</script>"
        'Page.RegisterStartupScript("window_onload", strScript)

        Dim sMyValue As String = ""
        TIMS.SetMyValue(sMyValue, "TPlanID", TPlanID.SelectedValue)
        TIMS.SetMyValue(sMyValue, "Years", Syear.SelectedValue)
        TIMS.SetMyValue(sMyValue, "DistID", DistID.SelectedValue)
        TIMS.SetMyValue(sMyValue, "RID", RIDValue.Value)
        TIMS.SetMyValue(sMyValue, "OCID", OCIDValue.Value)

        TIMS.SetMyValue(sMyValue, "STDate1", STDate1.Text)
        TIMS.SetMyValue(sMyValue, "STDate2", STDate2.Text)
        TIMS.SetMyValue(sMyValue, "FTDate1", FTDate1.Text)
        TIMS.SetMyValue(sMyValue, "FTDate2", FTDate2.Text)
        TIMS.SetMyValue(sMyValue, "stitle", stitle)
        TIMS.SetMyValue(sMyValue, "etitle", etitle)

        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "TR_04_007_R", sMyValue)

    End Sub
End Class
