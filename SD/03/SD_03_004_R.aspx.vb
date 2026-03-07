Partial Class SD_03_004_R
    Inherits System.Web.UI.Page

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在---------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        '檢查Session是否存在---------------------------End

        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)

        RadioButton1.Attributes("onclick") = "change()"
        RadioButton2.Attributes("onclick") = "change()"
        RadioButton3.Attributes("onclick") = "change()"
        Button1.Attributes("onclick") = "javascript@return print()"
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "", "", "TMIDValue1", "TMID1", True, "Button1")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim cGuid As String =   ReportQuery.GetGuid(Page)
        Dim Url As String =   ReportQuery.GetUrl(Page)
        Dim strScript As String
        If RadioButton1.Checked = True Then '身分
            strScript = "<script language=""javascript"">" + vbCrLf
            strScript += "window.open('" & Url & "GUID=" + cGuid + "&SQ_AutoLogout=true&sys=Member&filename=sign_up_list_by1&path=TIMS&OCID1=" & OCIDValue1.Value & "&TMID1=" & TMIDValue1.Value & "&RID=" & sm.UserInfo.RID & "&TPlanID=" & sm.UserInfo.TPlanID & "');" + vbCrLf
            strScript += "</script>"
            Page.RegisterStartupScript("window_onload", strScript)
        End If
        If RadioButton2.Checked = True Then '縣市
            strScript = "<script language=""javascript"">" + vbCrLf
            strScript += "window.open('" & Url & "GUID=" + cGuid + "&SQ_AutoLogout=true&sys=Member&filename=sign_up_list2&path=TIMS&OCID1=" & OCIDValue1.Value & "&TMID1=" & TMIDValue1.Value & "&RID=" & sm.UserInfo.RID & "&TPlanID=" & sm.UserInfo.TPlanID & "');" + vbCrLf
            strScript += "</script>"
            Page.RegisterStartupScript("window_onload", strScript)
        End If
        If RadioButton3.Checked = True Then '動態
            Dim strList As String = ""
            For i As Integer = 0 To Me.Sort1.Items.Count - 1
                If Me.Sort1.Items(i).Selected Then
                    strList &= Me.Sort1.Items(i).Value & ","
                End If
            Next

            strScript = "<script language=""javascript"">" + vbCrLf
            strScript += "window.open('" & Url & "GUID=" + cGuid + "&SQ_AutoLogout=true&sys=Member&filename=sing_up_list3&path=TIMS&OCID=" & OCIDValue1.Value & "&TMID=" & TMIDValue1.Value & "&RID=" & sm.UserInfo.RID & "&TPlanID=" & sm.UserInfo.TPlanID & "&Parameter='+escape('" & strList & "'));" + vbCrLf
            strScript += "</script>"
            Page.RegisterStartupScript("window_onload", strScript)

            'Response.Redirect("" & Url & "GUID=" + cGuid + "&SQ_AutoLogout=true&sys=list&filename=test&path=TIMS&Parameter=" & Server.UrlEncode(Sort1.SelectedValue) & "")
            '&OCID=" & Me.OCIDValue.Value & "&TMID=" & trainValue.Value & "&Years=" & Years.SelectedValue & "
        End If
    End Sub
End Class
