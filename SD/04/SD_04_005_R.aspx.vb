Partial Class SD_04_005_R
    Inherits AuthBasePage

    'iReport (會一次列印下列2張報表。)
    'Time_Schedule
    'Time_Schedule_Title

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
#Region "在這裡放置使用者程式碼以初始化網頁"

        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        '檢查Session是否存在 End

        Button1.Attributes("onclick") = "if(ReportPrint()){"
        'Button1.Attributes("onclick") +=     ReportQuery.ReportScript(Me, "list", "Time_Schedule_Title", "RID='+document.getElementById('RIDValue').value+'&OCID='+document.getElementById('OCIDValue1').value+'&TMID='+document.getElementById('TMIDValue1').value+'", , , False)
        Button1.Attributes("onclick") += ReportQuery.ReportScript(Me, "Member", "Time_Schedule", "RID='+document.getElementById('RIDValue').value+'&OCID='+document.getElementById('OCIDValue1').value+'&TMID='+document.getElementById('TMIDValue1').value+'")
        Button1.Attributes("onclick") += "}return false;"
        Button4.Attributes("onclick") = "if(ReportPrint()){"
        Button4.Attributes("onclick") += ReportQuery.ReportScript(Me, "Member", "Time_Schedule_Title", "RID='+document.getElementById('RIDValue').value+'&OCID='+document.getElementById('OCIDValue1').value+'&TMID='+document.getElementById('TMIDValue1').value+'", False)
        'Button4.Attributes("onclick") +=     ReportQuery.ReportScript(Me, "list", "Time_Schedule", "RID='+document.getElementById('RIDValue').value+'&OCID='+document.getElementById('OCIDValue1').value+'&TMID='+document.getElementById('TMIDValue1').value+'")
        Button4.Attributes("onclick") += "}return false;"

        If Not IsPostBack Then
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            If sm.UserInfo.LID <> "2" Then
                TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
            Else
                Button3_Click(sender, e)
            End If
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button2.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

#End Region
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
#Region "Button1_Click"

        'Dim cGuid1 As String =   ReportQuery.GetGuid(Page)
        'Dim cGuid2 As String =   ReportQuery.GetGuid(Page)
        'Dim Url As String =   ReportQuery.GetUrl(Page)
        'Dim strScript As String
        'strScript = "<script language=""javascript"">" + vbCrLf
        'strScript += "window.open('" & Url & "GUID=" + cGuid2 + "&SQ_AutoLogout=true&sys=list&filename=Time_Schedule_Title&path=TIMS&RID=" & Me.RIDValue.Value & "&OCID=" & Me.OCIDValue1.Value & "&TMID=" & TMIDValue1.Value & "');" + vbCrLf
        'strScript += "window.open('" & Url & "GUID=" + cGuid1 + "&SQ_AutoLogout=true&sys=list&filename=Time_Schedule&path=TIMS&RID=" & Me.RIDValue.Value & "&OCID=" & Me.OCIDValue1.Value & "&TMID=" & TMIDValue1.Value & "');" + vbCrLf
        'strScript += "</script>"
        'Page.RegisterStartupScript("window_onload", strScript)

#End Region
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '判斷機構是否只有一個班級
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value)
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        '如果只有一個班級
        If dr Is Nothing OrElse Convert.ToString(dr("total")) <> "1" Then Return
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
    End Sub
End Class