Partial Class SD_09_002_R
    Inherits AuthBasePage

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        '檢查Session是否存在 End

        Button1.Attributes("onclick") = "search('SQ_AutoLogout=true&sys=list&filename=officer_list&path=TIMS&RID=" & sm.UserInfo.RID & "');return false;"
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "", "", "TMIDValue1", "TMID1")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If
        If Not IsPostBack Then
            Button2_Click(sender, e)
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Dim cGuid As String =   ReportQuery.GetGuid(Page)
        'Dim Url As String =   ReportQuery.GetUrl(Page)
        'Dim strScript As String
        'If RadioButtonList1.SelectedValue = 0 Then '不區分
        '    strScript = "<script language=""javascript"">" + vbCrLf
        '    strScript += "window.open('" & Url & "GUID=" + cGuid + "&SQ_AutoLogout=true&sys=list&filename=officer_list&path=TIMS&OCID=" & Me.OCIDValue1.Value & "&RID=" & sm.UserInfo.RID & "');" + vbCrLf
        '    strScript += "</script>"
        '    Page.RegisterStartupScript("window_onload", strScript)
        '    'Response.Redirect("" & Url & "GUID=" + cGuid + "&SQ_AutoLogout=true&sys=list&filename=officer_list&path=TIMS&OCID=" & Me.OCIDValue1.Value & "&RID=" & sm.UserInfo.RID & "")
        'Else
        '    strScript = "<script language=""javascript"">" + vbCrLf
        '    strScript += "window.open('" & Url & "GUID=" + cGuid + "&SQ_AutoLogout=true&sys=list&filename=officer_list&path=TIMS&OCID=" & Me.OCIDValue1.Value & "&StudStatus=" & RadioButtonList1.SelectedValue & "&RID=" & sm.UserInfo.RID & "');" + vbCrLf
        '    strScript += "</script>"
        '    Page.RegisterStartupScript("window_onload", strScript)
        '    'Response.Redirect("" & Url & "GUID=" + cGuid + "&SQ_AutoLogout=true&sys=list&filename=officer_list&path=TIMS&OCID=" & Me.OCIDValue1.Value & "&StudStatus=" & RadioButtonList1.SelectedValue & "&RID=" & sm.UserInfo.RID & "")
        'End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub
End Class
