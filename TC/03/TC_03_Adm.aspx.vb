Partial Class TC_03_Adm
    Inherits AuthBasePage


    Const cst_CostItemTable As String = "CostItemTable" 'KEY_COSTITEM2 產投用經費項目
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        TIMS.OpenDbConn(objconn)

        If Not IsPostBack Then Call Create1()
    End Sub

    Sub Create1()
        Hid_CostItem_GUID1.Value = TIMS.sUtl_GetRqValue(Me, "CIGD", "")
        If Hid_CostItem_GUID1.Value <> "" Then Session(cst_CostItemTable) = Session(Hid_CostItem_GUID1.Value)

        Button1.Attributes("onclick") = "return check_date();"

        If Session(cst_CostItemTable) Is Nothing Then Return
        Dim dt As DataTable = Session(cst_CostItemTable)
        If dt Is Nothing Then Return
        If dt.Rows.Count = 0 Then
            Common.RespWrite(Me, "<script>alert('您尚未選擇任何項目');window.close();</script>")
            Return
        End If

        Dim dtKeyCostItem As DataTable = TIMS.GET_KEY_COSTITEMdt1(objconn)
        For Each dr As DataRow In dt.Select("CostMode IN (1,4)")
            If dr("CostID").ToString <> "99" Then
                AdmFlag.Items.Add(New ListItem(dtKeyCostItem.Select("CostID='" & dr("CostID") & "'")(0)("CostName"), dr("PCID")))
            Else
                AdmFlag.Items.Add(New ListItem("其他-" & dr("ItemOther"), dr("PCID")))
            End If
        Next

        For Each dr As DataRow In dt.Rows
            If dr.RowState <> DataRowState.Deleted Then
                For Each item As ListItem In AdmFlag.Items
                    If item.Value = dr("PCID") Then
                        If dr("AdmFlag").ToString = "Y" Then item.Selected = True
                    End If
                Next
            End If
        Next
        If Not Session("AdmGrant") Is Nothing Then AdmGrant.Text = Session("AdmGrant")
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        If Session(cst_CostItemTable) Is Nothing Then Exit Sub
        Dim dt As DataTable = Session(cst_CostItemTable)
        If dt Is Nothing Then Exit Sub
        If dt.Rows.Count = 0 Then Common.RespWrite(Me, "<script>alert('您尚未選擇任何項目');window.close();</script>")
        'Dim dr As DataRow = Nothing
        'dt = Session(cst_CostItemTable)
        For Each item As ListItem In AdmFlag.Items
            If item.Value <> "" AndAlso dt.Select("PCID='" & item.Value & "'").Length > 0 Then
                Dim dr As DataRow = dt.Select("PCID='" & item.Value & "'")(0)
                If dr IsNot Nothing Then
                    dr("AdmFlag") = If(item.Selected = True, "Y", "N")
                End If
            End If
        Next
        Session(cst_CostItemTable) = dt
        If Hid_CostItem_GUID1.Value <> "" Then Session(Hid_CostItem_GUID1.Value) = dt 'Session(cst_CostItemTable)
        Session("AdmGrant") = TIMS.ClearSQM(AdmGrant.Text)
        Page.RegisterStartupScript("000", "<script>opener.document.form1.submit();window.close();</script>")
    End Sub
End Class