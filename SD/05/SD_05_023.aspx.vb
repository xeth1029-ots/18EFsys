Partial Class SD_05_023
    Inherits AuthBasePage

   'Dim FunDr As DataRow
    Const Cst_StudentID As Integer = 1
    Const Cst_Sex As Integer = 4
    Const Cst_StudStatus As Integer = 6

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在--------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在--------------------------End

        If Not IsPostBack Then
            bt_search.Attributes("onclick") = "return CheckData();"

            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            TableShowData.Visible = False

            If sm.UserInfo.LID <> "2" Then
                TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
            Else
                Button3_Click(sender, e)
            End If

        End If
        If sm.UserInfo.RID = "A" Or sm.UserInfo.RoleID <= 1 Then
            BtnOrg.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg.aspx');"
        Else
            BtnOrg.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg1.aspx');"
        End If

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1", True, "bt_search")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

        'If sm.UserInfo.RoleID <> 0 Then
        '    Dim FunDt As DataTable = sm.UserInfo.FunDt
        '    Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & Request("ID") & "'")

        '    If FunDrArray.Length = 0 Then
        '        Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
        '        Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
        '    Else
        '        FunDr = FunDrArray(0)
        '        If FunDr("Adds") = 1 Or FunDr("Mod") = 1 Then
        '            btn_Save.Enabled = True
        '        Else
        '            btn_Save.Enabled = False
        '        End If
        '        If FunDr("Sech") = 1 Then
        '            bt_search.Enabled = True
        '        Else
        '            bt_search.Enabled = False
        '        End If
        '    End If
        'End If

    End Sub

    Private Sub bt_search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_search.Click
        Call GetShowData()
    End Sub

    Sub GetShowData()

        'Dim dt As DataTable
        'Dim dr As DataRow
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT c.Years" & vbCrLf
        sql &= " ,c.OCID" & vbCrLf ' --班別編號" & vbCrLf
        sql &= " ,c.PlanID" & vbCrLf
        sql &= " ,c.ComIDNO" & vbCrLf
        sql &= " ,c.SeqNo" & vbCrLf
        sql &= " ,c.ClassCName" & vbCrLf ' --培訓班別" & vbCrLf
        sql &= " ,c.CyclType" & vbCrLf ' --期別" & vbCrLf
        sql &= " ,c.TNum" & vbCrLf ' --計畫人數" & vbCrLf
        sql &= " ,c.THours" & vbCrLf '--訓練時數" & vbCrLf
        sql &= " ,dbo.NVL(s1.FinCount,0) FinCount1" & vbCrLf '--實際開訓人數
        sql &= " ,dbo.NVL(s1.FinCount2,0) FinCount2" & vbCrLf ' --結訓人數" & vbCrLf
        sql &= " ,j.JDID" & vbCrLf '  /*PK*/" & vbCrLf
        sql &= " ,j.TIMES" & vbCrLf
        sql &= " ,j.DEFAULTDATE" & vbCrLf
        sql &= " ,j.CONTRACTCOST" & vbCrLf
        sql &= " ,j.REMEDCOST" & vbCrLf
        sql &= " ,j.JOBENUM" & vbCrLf
        sql &= " ,j.JOBCNUM" & vbCrLf
        sql &= " ,j.JOBANUM" & vbCrLf
        'sql &= " ,j.MODIFYACCT,j.MODIFYDATE" & vbCrLf
        sql &= " FROM (" & vbCrLf
        sql &= " SELECT *" & vbCrLf
        sql &= " FROM Class_ClassInfo" & vbCrLf
        sql &= " WHERE NotOpen='N' and IsSuccess='Y'" & vbCrLf
        sql &= " AND OCID='" & Me.OCIDValue1.Value & "'" & vbCrLf
        sql &= " ) c" & vbCrLf
        sql &= " LEFT JOIN (" & vbCrLf
        sql &= " SELECT cs.OCID" & vbCrLf
        sql &= " ,Count(case when cs.StudStatus Not IN (2,3) then 1 end ) FinCount" & vbCrLf
        sql &= " ,Count(case when cs.StudStatus =5 then 1 end ) FinCount2" & vbCrLf
        sql &= " FROM Class_StudentsOfClass cs" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND cs.OCID='" & Me.OCIDValue1.Value & "'" & vbCrLf
        sql &= " Group By cs.OCID" & vbCrLf
        sql &= " ) s1 ON c.OCID=s1.OCID" & vbCrLf
        sql &= " LEFT JOIN Class_JobDefCost j on j.OCID =c.OCID" & vbCrLf
        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)

        If dt.Rows.Count > 0 Then
            TableSearch.Style.Item("display") = "none"
            Me.TableShowData.Style.Item("display") = "inline"
            TableShowData.Visible = True
            OCIDValue2.Disabled = True
            ContractCost.Enabled = False
            JobENum.Enabled = False
            JobANum.Enabled = False
            RemedCost.Enabled = False
            JobCNum.Enabled = False
            btn_Save.Visible = False
            btn_back.Visible = True
            btn_cancel.Visible = False
            btn_print.Visible = True
            btn_edit.Visible = True

            Dim dr As DataRow = dt.Rows(0)
            'If dr("Times").ToString <> "" Then
            '    labTimes.Text = dr("Times").ToString '申請第次數
            'Else
            labTimes.Text = TIMS.Get_StudTrainCostTimes(Me.OCIDValue1.Value, "T", objconn) '申請第次數
            'End If
            times2.Value = dr("Times").ToString

            Me.ClassCName.Text = dr("ClassCName").ToString
            Me.OCIDValue2.Value = dr("OCID").ToString 'OCID

            Me.CyclType.Text = dr("CyclType").ToString
            Me.TNum.Text = dr("TNum").ToString
            Me.THours.Text = dr("THours").ToString

            ContractCost.Text = dr("ContractCost").ToString '簽約訓練經費
            RemedCost.Text = dr("RemedCost").ToString '個人就業輔導費單價
            JobENum.Text = dr("JobENum").ToString '就業人數(檢附就業證明)
            JobCNum.Text = dr("JobCNum").ToString '就業人數(個案切結證明)
            JobANum.Text = dr("JobANum").ToString '就業輔導證明人數
        End If
    End Sub

    Private Sub btn_back_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_back.Click
        TableSearch.Style.Item("display") = "inline"
        Me.TableShowData.Style.Item("display") = "none"
    End Sub

    Private Sub btn_Save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Save.Click
        If Check_SaveDate() Then
            Insert_Class_JobDefCost()
            bt_search_Click(sender, e)
        End If
    End Sub

    Function Check_SaveDate() As Boolean
        Dim Rst As Boolean = True
        Dim Errmsg As String = ""
        'Check_SaveDate = False

        If Not IsNumeric(ContractCost.Text.Trim) Then
            Errmsg += "簽約訓練經費必須為數字" & vbCrLf
        End If

        If Not IsNumeric(RemedCost.Text.Trim) Then
            Errmsg += "個人就業輔導費單價必須為數字" & vbCrLf
        End If

        If Not IsNumeric(JobENum.Text.Trim) Then
            Errmsg += "就業人數(檢附就業證明)必須為數字" & vbCrLf
        End If

        If Not IsNumeric(JobCNum.Text.Trim) Then
            Errmsg += "就業人數(個案切結證明)必須為數字" & vbCrLf
        End If

        If Not IsNumeric(JobANum.Text.Trim) Then
            Errmsg += "就業輔導證明人數必須為數字" & vbCrLf
        End If

        If Errmsg <> "" Then
            Rst = False
            Common.MessageBox(Me, Errmsg)
        End If

        Return Rst
    End Function

    Sub Insert_Class_JobDefCost(Optional ByVal sType As Integer = 1)
        Dim dt As DataTable = Nothing
        Dim da As SqlDataAdapter = Nothing
        Dim dr As DataRow = Nothing

        Dim trans As SqlTransaction = Nothing
        'Dim conn As SqlConnection = Nothing
        Dim sql As String = ""

        Call TIMS.OpenDbConn(objconn)
        trans = DbAccess.BeginTrans(objconn)
        Try
            'conn = DbAccess.GetConnection
            'Dim labTimes As Label = Me.DataGrid1.FindControl("labTimes")
            'Dim txtRate As TextBox = Me.DataGrid1.FindControl("txtRate")
            'UPDATE Class_JobDefCost
            'sql = "SELECT * FROM Class_JobDefCost WHERE OCID='" & Me.OCIDValue2.Value & "' AND Times='" & CInt(labTimes.Text) & "'"
            sql = "SELECT * FROM Class_JobDefCost WHERE OCID='" & Me.OCIDValue2.Value & "'"
            dt = DbAccess.GetDataTable(sql, da, trans)
            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
            Else
                dr = dt.NewRow
                dt.Rows.Add(dr)
                dr("OCID") = Me.OCIDValue2.Value
                dr("Times") = CInt(labTimes.Text)
            End If
            dr("DefaultDate") = FormatDateTime(Now(), DateFormat.ShortDate) '申請日期
            dr("ContractCost") = IIf(ContractCost.Text.Trim = "", Convert.DBNull, ContractCost.Text.Trim) '簽約訓練經費
            dr("RemedCost") = IIf(RemedCost.Text.Trim = "", Convert.DBNull, RemedCost.Text.Trim)  '個人就業輔導費單價
            dr("JobENum") = IIf(JobENum.Text.Trim = "", Convert.DBNull, JobENum.Text.Trim)  '就業人數(檢附就業證明)
            dr("JobCNum") = IIf(JobCNum.Text.Trim = "", Convert.DBNull, JobCNum.Text.Trim)  '就業人數(個案切結證明)
            dr("JobANum") = IIf(JobANum.Text.Trim = "", Convert.DBNull, JobANum.Text.Trim)  '就業輔導證明人數
            dr("ModifyAcct") = sm.UserInfo.UserID
            dr("ModifyDate") = Now
            DbAccess.UpdateDataTable(dt, da, trans)

            DbAccess.CommitTrans(trans)
            Common.MessageBox(Me, "儲存成功")
        Catch ex As Exception
            DbAccess.RollbackTrans(trans)
            Throw ex
        End Try
    End Sub

    Private Sub btn_edit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_edit.Click

        OCIDValue2.Disabled = False
        ContractCost.Enabled = True
        JobENum.Enabled = True
        JobANum.Enabled = True
        RemedCost.Enabled = True
        JobCNum.Enabled = True
        btn_Save.Visible = True
        btn_back.Visible = False
        btn_cancel.Visible = True
        btn_print.Visible = False
        btn_edit.Visible = False

    End Sub

    Private Sub btn_cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_cancel.Click
        GetShowData()
    End Sub

    Private Sub btn_print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_print.Click
        Dim Years As Integer
        Years = sm.UserInfo.Years - 1911

        ReportQuery.PrintReport(Me, "Member", "SD_05_023", "OCID=" & OCIDValue1.Value & "&Years=" & Years & "&Times=" & times2.Value)
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '判斷機構是否只有一個班級
        Dim dr As DataRow
        dr = TIMS.GET_OnlyOne_OCID(RIDValue.Value)
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        Me.TableShowData.Style.Item("display") = "none"
        If Not dr Is Nothing Then
            If dr("total") = "1" Then '如果只有一個班級
                TMID1.Text = dr("trainname")
                OCID1.Text = dr("classname")
                TMIDValue1.Value = dr("trainid")
                OCIDValue1.Value = dr("ocid")
                Me.TableShowData.Style.Item("display") = "none"
            End If
        End If
    End Sub

End Class
