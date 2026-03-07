'Imports System.Data.SqlClient
'Imports System.Data
'Imports Turbo
Public Class TC_04_Trans
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DistID As System.Web.UI.WebControls.DropDownList
    Protected WithEvents choice_button As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents TBplan As System.Web.UI.WebControls.TextBox
    Protected WithEvents RIDValue As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents PlanIDValue As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents btnAdd As System.Web.UI.WebControls.Button
    Protected WithEvents OrgLevel As System.Web.UI.WebControls.DropDownList

    '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
    '請勿刪除或移動它。
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
        '請勿使用程式碼編輯器進行修改。
        InitializeComponent()
    End Sub

#End Region

    Dim SqlCmd As String = ""
    Dim UserID As String = ""
    Dim PlanID As String = ""
    Dim IDNO As String = ""
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在--------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在--------------------------End

        UserID = sm.UserInfo.UserID
        PlanID = Request("PlanID")
        IDNO = Request("ComIDNO")

        If Not Me.IsPostBack Then
            '轄區中心
            SqlCmd = "select * from ID_District where DistID <>000  order by DistID"
            Me.DistID.Items.Clear()
            DbAccess.MakeListItem(Me.DistID, SqlCmd, objconn)
            Me.DistID.Items.Insert(0, New ListItem("===請選擇===", ""))
        End If
    End Sub

    Private Sub OrgLevel_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OrgLevel.SelectedIndexChanged
        Me.choice_button.Disabled = False
        Me.TBplan.ReadOnly = False
        Me.TBplan.Text = ""
        Me.PlanIDValue.Value = ""
        Me.RIDValue.Value = ""
        Select Case Me.OrgLevel.SelectedValue
            Case "1"
                Me.choice_button.Disabled = True
                Me.TBplan.ReadOnly = True
                Me.TBplan.Text = "所有計畫"
                Me.PlanIDValue.Value = "0"
                Me.RIDValue.Value = "A"
        End Select
    End Sub

    '儲存
    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        '判斷訓練機構是否有值
        Dim drA As DataRow
        Dim drB As DataRow
        Dim sqlAdapter As SqlClient.SqlDataAdapter
        Dim sqlTable As DataTable

        'Dim objConn As SqlConnection
        'TIMS.TestDbConn(Me, objConn)

        Dim PlanID2 As String = ""
        Dim maxrid As String = ""
        Dim RID As String
        RID = Me.RIDValue.Value

        SqlCmd = "Select * From Org_OrgInfo where ComIDNO='" & IDNO & "'"
        If DbAccess.GetCount(SqlCmd, objconn) = 0 Then
            'Insert訓練機構檔
            SqlCmd = "Select * From Com_Tender where ComIDNO='" & IDNO & "'"
            drA = DbAccess.GetOneRow(SqlCmd, objconn)

            drB = DbAccess.GetInsertRow("Org_OrgInfo", sqlTable, sqlAdapter, objconn)
            drB("OrgID") = Me.OrgLevel.SelectedValue
            drB("OrgKind") = Me.DistID.SelectedValue
            drB("OrgName") = drA("ComName")
            drB("ComIDNO") = IDNO
            drB("ZipCode") = drA("ZipCode")
            drB("Address") = drA("Address")
            drB("Phone") = drA("Phone")
            drB("ContactName") = drA("ContactName")
            drB("ContactEmail") = drA("ContactEmail")
            drB("ContactCellPhone") = drA("ContactCellPhone")
            drB("ModifyAcct") = UserID
            drB("ModifyDate") = DateTime.Now
            sqlTable.Rows.Add(drB)

            'insert業務關係檔
            drA = DbAccess.GetInsertRow("Auth_Relship", sqlTable, sqlAdapter, objconn)

            If Me.OrgLevel.SelectedValue = 1 Then
                '取目前選取計畫的relship
                Dim sqlstr_A As String = "select relship from  Auth_Relship where rid='" & RID & "'"
                PlanID = Convert.ToString(DbAccess.ExecuteScalar(sqlstr_A, objconn))
                '選取新增項目的最大項
                Dim sqlstr_next As String = "select max(rid) from Auth_Relship where orglevel='" & Me.OrgLevel.SelectedValue & "' and  relship like '" & PlanID & "%'"
                maxrid = Convert.ToString(DbAccess.ExecuteScalar(sqlstr_next, objconn))
                If maxrid = "" Then
                    drA("RID") = Chr(Asc(RID) + 1)
                    drA("Relship") = PlanID + Chr(Asc(RID) + 1) + "/"
                    drA("OrgLevel") = Me.OrgLevel.SelectedValue
                Else
                    drA("RID") = Chr(Asc(maxrid) + 1)
                    drA("Relship") = PlanID + Chr(Asc(maxrid) + 1) + "/"
                    drA("OrgLevel") = Me.OrgLevel.SelectedValue
                End If
            ElseIf Me.OrgLevel.SelectedValue = 2 Then
                '取目前選取計畫的relship
                Dim sqlstr_A As String = "select relship from  Auth_Relship where rid='" & RID & "'"
                PlanID = Convert.ToString(DbAccess.ExecuteScalar(sqlstr_A, objconn))
                '取目前選取計畫的orglevel
                Dim sqlstr_orglevel As String = "select a.orglevel  from  Auth_Relship a where rid='" & Me.RIDValue.Value & "'"
                PlanID2 = Convert.ToString(DbAccess.ExecuteScalar(sqlstr_orglevel, objconn))
                '選取新增項目的最大項
                Dim sqlstr_next As String = "select max(rid)  from  Auth_Relship where orglevel=" & PlanID2 & "+'1' and  relship like '" & PlanID & "%'"
                maxrid = Convert.ToString(DbAccess.ExecuteScalar(sqlstr_next, objconn))
                '欲新增的計畫是否存在
                If maxrid = "" Then
                    If PlanID2 = "1" Then '沒有子單位
                        drA("RID") = RID & "01"
                        drA("Relship") = PlanID + RID & "01" + "/"
                        drA("OrgLevel") = CInt(PlanID2) + 1
                    Else '沒有子單位,orglevel>1
                        drA("RID") = RID.Substring(0, 1) & (CInt(RID.Substring(1))).ToString(New String("0", CInt(PlanID2) + 1))
                        drA("Relship") = PlanID + RID.Substring(0, 1) & (CInt(RID.Substring(1))).ToString(New String("0", CInt(PlanID2) + 1)) + "/"
                        drA("OrgLevel") = CInt(PlanID2) + 1
                    End If
                Else
                    '有子單位,加1
                    drA("RID") = maxrid.Substring(0, 1) & (CInt(maxrid.Substring(1)) + 1).ToString(New String("0", CInt(PlanID2) + 1))
                    drA("Relship") = PlanID + maxrid.Substring(0, 1) & (CInt(maxrid.Substring(1)) + 1).ToString(New String("0", CInt(PlanID2) + 1)) + "/"
                    drA("OrgLevel") = CInt(PlanID2) + 1
                End If

            End If

            drA("PlanID") = Me.PlanIDValue.Value
            drA("OrgID") = Me.OrgLevel.SelectedValue
            drA("DistID") = Me.DistID.SelectedValue
            drA("ModifyAcct") = UserID
            drA("ModifyDate") = DateTime.Now
            sqlTable.Rows.Add(drA)

            Dim objTrans As SqlTransaction
            Try
                objTrans = DbAccess.BeginTrans(objconn)

                DbAccess.UpdateDataTable(sqlTable, sqlAdapter, objTrans)

                DbAccess.CommitTrans(objTrans)

            Catch ex As Exception
                DbAccess.RollbackTrans(objTrans)
                Throw ex
            End Try

        Else
            drB = DbAccess.GetInsertRow("Auth_Relship", sqlTable, sqlAdapter, objconn)

            If Me.OrgLevel.SelectedValue = 1 Then
                '取目前選取計畫的relship
                Dim sqlstr_A As String = "select relship from  Auth_Relship where rid='" & RID & "'"
                PlanID = Convert.ToString(DbAccess.ExecuteScalar(sqlstr_A, objconn))
                '選取新增項目的最大項
                Dim sqlstr_next As String = "select max(rid) from Auth_Relship where orglevel='" & Me.OrgLevel.SelectedValue & "' and  relship like '" & PlanID & "%'"
                maxrid = Convert.ToString(DbAccess.ExecuteScalar(sqlstr_next, objconn))
                If maxrid = "" Then
                    drB("RID") = Chr(Asc(RID) + 1)
                    drB("Relship") = PlanID + Chr(Asc(RID) + 1) + "/"
                    drB("OrgLevel") = Me.OrgLevel.SelectedValue
                Else
                    drB("RID") = Chr(Asc(maxrid) + 1)
                    drB("Relship") = PlanID + Chr(Asc(maxrid) + 1) + "/"
                    drB("OrgLevel") = Me.OrgLevel.SelectedValue
                End If
            ElseIf Me.OrgLevel.SelectedValue = 2 Then
                '取目前選取計畫的relship
                Dim sqlstr_A As String = "select relship from  Auth_Relship where rid='" & RID & "'"
                PlanID = Convert.ToString(DbAccess.ExecuteScalar(sqlstr_A, objconn))
                '取目前選取計畫的orglevel
                Dim sqlstr_orglevel As String = "select a.orglevel  from  Auth_Relship a where rid='" & Me.RIDValue.Value & "'"
                PlanID2 = Convert.ToString(DbAccess.ExecuteScalar(sqlstr_orglevel, objconn))
                '選取新增項目的最大項
                Dim sqlstr_next As String = "select max(rid)  from  Auth_Relship where orglevel=" & PlanID2 & "+'1' and  relship like '" & PlanID & "%'"
                maxrid = Convert.ToString(DbAccess.ExecuteScalar(sqlstr_next, objconn))
                '欲新增的計畫是否存在
                If maxrid = "" Then
                    If PlanID2 = "1" Then '沒有子單位
                        drB("RID") = RID & "01"
                        drB("Relship") = PlanID + RID & "01" + "/"
                        drB("OrgLevel") = CInt(PlanID2) + 1
                    Else '沒有子單位,orglevel>1
                        drB("RID") = RID.Substring(0, 1) & (CInt(RID.Substring(1))).ToString(New String("0", CInt(PlanID2) + 1))
                        drB("Relship") = PlanID + RID.Substring(0, 1) & (CInt(RID.Substring(1))).ToString(New String("0", CInt(PlanID2) + 1)) + "/"
                        drB("OrgLevel") = CInt(PlanID2) + 1
                    End If
                Else
                    '有子單位,加1
                    drB("RID") = maxrid.Substring(0, 1) & (CInt(maxrid.Substring(1)) + 1).ToString(New String("0", CInt(PlanID2) + 1))
                    drB("Relship") = PlanID + maxrid.Substring(0, 1) & (CInt(maxrid.Substring(1)) + 1).ToString(New String("0", CInt(PlanID2) + 1)) + "/"
                    drB("OrgLevel") = CInt(PlanID2) + 1
                End If

            End If

            drB("PlanID") = Me.PlanIDValue.Value
            drB("OrgID") = Me.OrgLevel.SelectedValue
            drB("DistID") = Me.DistID.SelectedValue
            drB("ModifyAcct") = UserID
            drB("ModifyDate") = DateTime.Now
            sqlTable.Rows.Add(drB)
            DbAccess.UpdateDataTable(sqlTable, sqlAdapter)
            objconn.Close()
        End If
        Dim strScript As String
        strScript = "<script language=""javascript"">" + vbCrLf
        strScript += "window.close();" + vbCrLf
        strScript += "</script>"
        Page.RegisterStartupScript("", strScript)

    End Sub
End Class
