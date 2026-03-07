Partial Class TC_03_Adm2
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        TIMS.OpenDbConn(objconn)

        If Not IsPostBack Then Create1()

    End Sub

    Sub Create1()
        Hid_CostItem_GUID1.Value = TIMS.sUtl_GetRqValue(Me, "CIGD", "")
        If Hid_CostItem_GUID1.Value <> "" Then Session("CostItemTable") = Session(Hid_CostItem_GUID1.Value)

        Button1.Attributes("onclick") = "return check_date();"

        If Session("CostItemTable") Is Nothing Then Return
        Dim dt As DataTable = Session("CostItemTable")
        If dt Is Nothing Then Return

        If dt.Rows.Count = 0 Then
            Common.RespWrite(Me, "<script>alert('您尚未選擇任何項目');window.close();</script>")
            Return
        End If

        Dim dtKeyCostItem As DataTable = TIMS.GET_KEY_COSTITEMdt1(objconn)
        For Each dr As DataRow In dt.Select("CostMode IN (1,4)")
            If dr("CostID").ToString <> "99" Then
                TaxFlag.Items.Add(New ListItem(dtKeyCostItem.Select("CostID='" & dr("CostID") & "'")(0)("CostName"), dr("PCID")))
            Else
                TaxFlag.Items.Add(New ListItem("其他-" & dr("ItemOther"), dr("PCID")))
            End If
        Next

        For Each dr As DataRow In dt.Rows
            If dr.RowState <> DataRowState.Deleted Then
                For Each item As ListItem In TaxFlag.Items
                    If item.Value = dr("PCID") Then
                        If dr("TaxFlag").ToString = "" Then item.Selected = True '設定預設值為選擇
                        If dr("TaxFlag").ToString = "Y" Then item.Selected = True
                    End If
                Next
            End If
        Next
        If Not Session("TaxGrant") Is Nothing Then TaxGrant.Text = Session("TaxGrant")
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim dt As DataTable
        Dim dr As DataRow
        dt = Session("CostItemTable")
        For Each item As ListItem In TaxFlag.Items
            If dt.Select("PCID='" & item.Value & "'").Length <> 0 Then
                dr = dt.Select("PCID='" & item.Value & "'")(0)
                If dr IsNot Nothing Then
                    dr("TaxFlag") = If(item.Selected = True, "Y", "N")
                End If
            End If
        Next
        Session("CostItemTable") = dt
        If Hid_CostItem_GUID1.Value <> "" Then Session(Hid_CostItem_GUID1.Value) = dt 'Session("CostItemTable")
        Session("TaxGrant") = TIMS.ClearSQM(TaxGrant.Text)
        Page.RegisterStartupScript("000", "<script>opener.document.form1.submit();window.close();</script>")
    End Sub
End Class