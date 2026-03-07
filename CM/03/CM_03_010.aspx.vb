Partial Class CM_03_010
    Inherits AuthBasePage

    'Dim au As New cAUTH
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        If Not IsPostBack Then
            ddlYear = TIMS.GetSyear(ddlYear)
            ddlYear.SelectedValue = Now.ToString("yyyy")

            rblCity = TIMS.Get_CityName(rblCity, TIMS.dtNothing())
            rblCity.Items(0).Selected = True

            btnPrt.Attributes.Add("onclick", "return chkPrt();")
        End If
    End Sub

    Private Sub btnPrt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrt.Click
        Dim strCityID As String = ""
        Dim strOCID As String = ""
        Dim intCnt As Int16 = 0

        '找ctid
        For i As Integer = 0 To rblCity.Items.Count - 1
            If rblCity.Items(i).Selected Then
                strCityID = rblCity.Items(i).Value
                Exit For
            End If
        Next

        Try
            '查ocid
            Dim sql As String = ""
            sql = ""
            sql += " select a.ocid"
            sql += " from class_classinfo a "
            sql += " join id_plan b on b.planid=a.planid and b.tplanid NOT IN (" & TIMS.Cst_NotTPlanID2 & ") "
            sql += " join id_zip c on c.zipcode=a.taddresszip "
            sql += " join id_city d on d.ctid=c.ctid "
            sql += " where d.ctid='" & strCityID & "' "
            If ddlYear.SelectedValue <> "" Then
                sql += " and b.years='" & ddlYear.SelectedValue & "' "
            End If
            If txtSTSDate.Text <> "" Then
                sql += " and a.stdate>= " & TIMS.To_date(txtSTSDate.Text) & vbCrLf
            End If
            If txtSTEDate.Text <> "" Then
                sql += " and a.stdate<= " & TIMS.To_date(txtSTEDate.Text) & vbCrLf
            End If
            If txtETSDate.Text <> "" Then
                sql += " and a.ftdate>= " & TIMS.To_date(txtETSDate.Text) & vbCrLf
            End If
            If txtETEDate.Text <> "" Then
                sql += " and a.ftdate<= " & TIMS.To_date(txtETEDate.Text) & vbCrLf
            End If
            Dim oCMD As New SqlCommand(sql, objconn)
            'Call TIMS.OpenDbConn(conn)
            Dim odt As New DataTable
            With oCMD
                .Parameters.Clear()
                odt.Load(.ExecuteReader())
            End With
            'With sda
            '    .SelectCommand = New SqlCommand(sql, conn)
            '    .Fill(ds, "data")
            'End With
            strOCID = ""
            For Each dr As DataRow In odt.Rows
                If strOCID <> "" Then strOCID &= ","
                strOCID &= "'" & dr("ocid") & "'"
            Next

            '查詢有無學員
            sql = ""
            sql += " select ocid "
            sql += " from class_studentsofclass "
            sql += " where ocid in (" & strOCID & ")"
            Dim oCMD2 As New SqlCommand(sql, objconn)
            'Call TIMS.OpenDbConn(conn)
            Dim odt2 As New DataTable
            With oCMD2
                .Parameters.Clear()
                odt2.Load(.ExecuteReader())
            End With

            intCnt = odt2.Rows.Count
        Catch ex As Exception
            Common.MessageBox(Me, ex.ToString)
            Call TIMS.CloseDbConn(objconn) ' conn.Close()
            Exit Sub
        End Try
        Call TIMS.CloseDbConn(objconn)

        If intCnt = 0 Then
            Common.MessageBox(Me, "查無資料!")
            Exit Sub
        End If

        Common.RespWrite(Me, "<script>window.open('CM_03_010_R.aspx?year=" & ddlYear.SelectedValue & "&stsdate=" & txtSTSDate.Text & "&stedate=" & txtSTEDate.Text & "&etsdate=" & txtETSDate.Text & "&etedate=" & txtETEDate.Text & "&city=" & strCityID & "');</script>")

    End Sub
End Class
