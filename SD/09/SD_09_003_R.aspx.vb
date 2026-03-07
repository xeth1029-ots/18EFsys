'Imports Turbo
'Imports System.Data.SqlClient

Partial Class SD_09_003_R
    Inherits System.Web.UI.Page

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在--------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        '檢查Session是否存在--------------------------End


        If Not IsPostBack Then
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            'course.Items.Add(New ListItem("--請選擇班別--", 0))
        End If
        Button1.Attributes("onclick") = "if(ReportPrint()){"
        Button1.Attributes("onclick") =     ReportQuery.ReportScript(Me, "MultiBlock", "MOROI_Report", "OCID='+document.getElementById('OCIDValue1').value+'TMID='+document.getElementById('TMIDValue1').value+'&SDate='+document.getElementById('SDate').value+'&RID='+document.getElementById('RIDValue').value+'&TPlanID=" & sm.UserInfo.TPlanID & "")
        Button1.Attributes("onclick") = "}return false;"
        If sm.UserInfo.RID = "A" Or sm.UserInfo.RoleID <= 1 Then
            Button3.Attributes("onclick") = "javascript@openOrg('../../Common/LevOrg.aspx');"
        Else
            Button3.Attributes("onclick") = "javascript@openOrg('../../Common/LevOrg1.aspx');"
        End If
        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

        'Button2.Attributes("onclick") = "javascript@return search()"
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Dim cGuid As String =   ReportQuery.GetGuid(Page)
        'Dim Url As String =   ReportQuery.GetUrl(Page)
        'Dim strScript As String
        'strScript = "<script language=""javascript"">" + vbCrLf
        'strScript += "window.open('" & Url & "GUID=" + cGuid + "&SQ_AutoLogout=true&sys=MultiBlock&path=TIMS&filename=MOROI_Report&OCID=" & OCIDValue1.Value & "&TMID=" & TMIDValue1.Value & "&SDate=" & SDate.Text & "&RID=" & RIDValue.Value & "&TPlanID=" & sm.UserInfo.TPlanID & "');" + vbCrLf
        'strScript += "</script>"

        'Page.RegisterStartupScript("window_onload", strScript)
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim sql As String
        sql = "SELECT b.CourID,b.CourseName FROM (SELECT CourID FROM Plan_Schedule WHERE OCID='" & OCIDValue1.Value & "') a join (SELECT CourID,CourseName From Course_CourseInfo) b on a.CourID=b.CourID"
        Dim dt As DataTable = DbAccess.GetDataTable(sql)
        'If dt.Rows.Count = 0 Then
        '    course.Items.Clear()
        '    course.Items.Add(New ListItem("--請選擇班別--", 0))
        'Else
        '    With course
        '        .DataSource = dt
        '        .DataTextField = "CourseName"
        '        .DataValueField = "CourID"
        '        .DataBind()
        '    End With
        'End If
    End Sub
End Class
