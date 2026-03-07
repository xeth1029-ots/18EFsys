Partial Class SD_05_017_R
    Inherits AuthBasePage

    'xx:  SD_05_017_R  'SD_05_017_R2
    'SD_05_017_R3
    'Dim conn As SqlConnection
    'Dim sql As String
    'Dim dr As DataRow
    'Dim item1, item2, item3, item4 As Integer
    'Dim table As DataTable
    'Dim da As SqlDataAdapter = nothing
    'SELECT * FROM KEY_LEAVE

    Const cst_AlertMsg1 As String = "請系統管理者至首頁>>系統管理>>系統參數管理>>參數設定,設定此計畫出缺勤警示(V1)!"
    Const cst_AlertMsg2 As String = "請系統管理者至首頁>>系統管理>>系統參數管理>>參數設定,設定此計畫出缺勤警示(V2)!!"

    Const cst_printFN1 As String = "SD_05_017_R3" '在職 
    'Const cst_printFN2 As String = "SD_05_017_R4" '2017加生理假-職前 
    'Const cst_printFN3 As String = "SD_05_017_R5" '2019-區域產業據點
    Const cst_printFN3 As String = "SD_05_017_R5v" '2021-區域產業據點 學員出缺勤統計表

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        'TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        If Not IsPostBack Then
            Call CreateShowhis1()
            Call Create1() '設定此計畫出缺勤警示
            Button1.Attributes("onclick") = "return search();"
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
        End If
    End Sub

    Sub CreateShowhis1()
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

    End Sub

    '設定此計畫出缺勤警示
    Sub Create1()
        Call SHOW_ITEMVAL_ALL_1234()
    End Sub

    Sub SHOW_ITEMVAL_ALL_1234()
        Dim item1 As Integer = 0
        Dim item2 As Integer = 0 '1
        Dim item3 As Integer = 0
        Dim item4 As Integer = 0 '1

        Dim s_AltERRMSG1 As String = ""
        Dim flagIV1 As Boolean = True
        Dim flagIV2 As Boolean = True
        Dim sItemVar1 As String = TIMS.GetGlobalVar(Me, "4", "1", objconn)
        Dim sItemVar2 As String = TIMS.GetGlobalVar(Me, "4", "2", objconn)
        If (sItemVar1 = "0") OrElse (sItemVar1 = "") OrElse sItemVar2.IndexOf("/") = -1 Then
            flagIV1 = False
            s_AltERRMSG1 &= cst_AlertMsg1
            'Page.RegisterStartupScript("Startup ", "<script language=""javascript"">alert('" & cst_AlertMsg1 & "')</Script>")
            'Return
        End If
        If (sItemVar2 = "0") OrElse (sItemVar2 = "") OrElse sItemVar2.IndexOf("/") = -1 Then
            flagIV2 = False
            s_AltERRMSG1 &= cst_AlertMsg2
            'Page.RegisterStartupScript("Startup ", "<script language=""javascript"">alert('" & cst_AlertMsg2 & "')</Script>")
            'Return
        End If
        If flagIV1 AndAlso sItemVar1 <> "" AndAlso sItemVar1.IndexOf("/") > -1 Then
            item1 = If(Split(sItemVar1, "/").Length > 0, Int(Split(sItemVar1, "/")(0)), 0)
            item2 = If(Split(sItemVar1, "/").Length > 1 AndAlso Int(Split(sItemVar1, "/")(1)) <> 0, Int(Split(sItemVar1, "/")(1)), 1)
        End If
        If flagIV2 AndAlso sItemVar2 <> "" AndAlso sItemVar2.IndexOf("/") > -1 Then
            item3 = If(Split(sItemVar2, "/").Length > 0, Int(Split(sItemVar2, "/")(0)), 0)
            item4 = If(Split(sItemVar2, "/").Length > 1 AndAlso Int(Split(sItemVar2, "/")(1)) <> 0, Int(Split(sItemVar2, "/")(1)), 1)
        End If
        If s_AltERRMSG1 <> "" Then
            Page.RegisterStartupScript("Startup ", "<script language=""javascript"">alert('" & s_AltERRMSG1 & "')</Script>")
        End If

        Hitem1.Value = CStr(item1)
        Hitem2.Value = CStr(item2)
        Hitem3.Value = CStr(item3)
        Hitem4.Value = CStr(item4)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Dim RID As String
        Dim s_RID As String = sm.UserInfo.RID
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value <> "" Then s_RID = RIDValue.Value
        s_RID = TIMS.ClearSQM(s_RID)

        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        TMIDValue1.Value = TIMS.ClearSQM(TMIDValue1.Value)
        Hitem1.Value = TIMS.ClearSQM(Hitem1.Value)
        Hitem2.Value = TIMS.ClearSQM(Hitem2.Value)
        Hitem3.Value = TIMS.ClearSQM(Hitem3.Value)
        Hitem4.Value = TIMS.ClearSQM(Hitem4.Value)
        TDate.Text = TIMS.ClearSQM(TDate.Text)

        Dim strMyValue As String = ""
        strMyValue += "OCID=" & OCIDValue1.Value
        strMyValue += "&TMIDValue1=" & TMIDValue1.Value
        strMyValue += "&TPlanID=" & sm.UserInfo.TPlanID
        strMyValue += "&RID=" & s_RID
        strMyValue += "&item1=" & Hitem1.Value
        strMyValue += "&item2=" & Hitem2.Value
        strMyValue += "&item3=" & Hitem3.Value
        strMyValue += "&item4=" & Hitem4.Value
        strMyValue += "&Enddate=" & TDate.Text '統計截止日 
        'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "list", "SD_05_017_R", "OCID='+document.getElementById('OCIDValue1').value+'&TMID='+document.getElementById('TMIDValue1').value+'&RID='+document.getElementById('RIDValue').value+'&TPlanID=" & sm.UserInfo.TPlanID & "&item1=" & item1 & "&item2=" & item2 & "&item3=" & item3 & "&item4=" & item4)
        'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "Member", "SD_05_017_R2", strMyValue)
        'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "SD_05_017_R3", strMyValue)

        'Dim flagYear2017 As Boolean = False
        'flagYear2017 = TIMS.Get_UseLEAVE_2017(Me)
        Dim sPrintName1 As String = cst_printFN1
        'If flagYear2017 Then sPrintName1 = cst_printFN2
        If TIMS.Cst_TPlanID70.IndexOf(sm.UserInfo.TPlanID) > -1 Then sPrintName1 = cst_printFN3

        '請假、缺曠課累計時數統計表
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, sPrintName1, strMyValue)
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '判斷機構是否只有一個班級
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn)
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        TDate.Text = ""
        '如果只有一個班級
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
        TDate.Text = dr("FTDate")
    End Sub

    Private Sub Button4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button4.Click
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC IsNot Nothing Then
            TDate.Text = TIMS.Cdate3(drCC("FTDate"))
        End If
    End Sub
End Class