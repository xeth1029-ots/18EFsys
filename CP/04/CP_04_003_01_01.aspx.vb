Partial Class CP_04_003_01_01
    Inherits AuthBasePage

    '單筆學員歷程查詢 (依身分證號)
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁

        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End
        'PageControler1.PageDataGrid = Stud_DG

        If Not IsPostBack Then
            btnClose2.Attributes("onclick") = "window.close();"
            Call Search1()
        End If

    End Sub

    Sub Search1()
        Dim rqStudent_History As String = TIMS.ChangeIDNO(TIMS.ClearSQM(Request("Student_History")))
        If rqStudent_History = "" Then
            Common.MessageBox(Me, "查詢資料有誤，請重新查詢!!")
            Exit Sub
        End If
        If Convert.ToString(Session("Student_History")) = "" Then
            Common.MessageBox(Me, "查詢資料有誤，請重新查詢!!")
            Exit Sub
        End If
        If rqStudent_History <> Session("Student_History") Then
            Common.MessageBox(Me, "查詢資料有誤，請重新查詢!!")
            Exit Sub
        End If

        Dim dt As DataTable = Nothing
        Dim dt1 As DataTable = Nothing
        Call sUtl_CreateDt1(dt)

        Dim parms1 As New Hashtable
        parms1.Add("IDNO", rqStudent_History)
        Dim sqlstr As String = ""
        sqlstr = "" & vbCrLf
        sqlstr &= " SELECT Serial StdID ,DistName ,PlanName ,TrinUnit ,ClassName " & vbCrLf
        sqlstr &= " ,CONVERT(VARCHAR, Name) Name ,CONVERT(VARCHAR, Sex) Sex " & vbCrLf
        sqlstr &= " ,CONVERT(VARCHAR, IDNO) IDNO ,CONVERT(VARCHAR, Ident) Ident " & vbCrLf
        sqlstr &= " ,TPlanID ,CONVERT(VARCHAR, DistID) DistID ,'History_StudentInfo93' source " & vbCrLf
        sqlstr &= " FROM History_StudentInfo93 " & vbCrLf
        sqlstr &= " WHERE IDNO =@IDNO" & vbCrLf
        dt1 = DbAccess.GetDataTable(sqlstr, objconn, parms1)
        Call sUtl_dt1Data2dtData(dt1, dt)

        Dim parms2 As New Hashtable
        parms2.Add("SID", rqStudent_History)
        sqlstr = ""
        sqlstr &= " SELECT StdID ,CONVERT(VARCHAR, DistName) DistName ,PlanName ,TrinUnit ,ClassName " & vbCrLf
        sqlstr &= " ,CONVERT(VARCHAR, Name) Name ,CONVERT(VARCHAR, Sex) Sex ,CONVERT(VARCHAR, SID) IDNO" & vbCrLf
        sqlstr &= " ,CONVERT(VARCHAR, Ident) Ident ,TPlanID ,CONVERT(VARCHAR, DistID) DistID ,'StdAll' source " & vbCrLf
        sqlstr &= " FROM StdAll " & vbCrLf
        sqlstr &= " WHERE SID =@SID " & vbCrLf
        dt1 = DbAccess.GetDataTable(sqlstr, objconn, parms2)
        Call sUtl_dt1Data2dtData(dt1, dt)

        Dim parms3 As New Hashtable
        parms3.Add("IDNO", rqStudent_History)
        sqlstr = ""
        sqlstr &= " SELECT CONVERT(VARCHAR, a.SOCID) StdID ,CONVERT(VARCHAR, f.Name) DistName ,j.PlanName " & vbCrLf
        sqlstr &= " ,e.OrgName TrinUnit ,c.ClassCName ClassName ,CONVERT(VARCHAR, b.Name) Name " & vbCrLf
        sqlstr &= " ,CONVERT(VARCHAR, b.Sex) Sex ,CONVERT(VARCHAR, b.IDNO) IDNO ,a.IdentityID Ident ,i.TPlanID " & vbCrLf
        sqlstr &= " ,CONVERT(VARCHAR, d.DistID) DistID ,'Class_StudentsOfClass' source " & vbCrLf
        sqlstr &= " FROM Class_StudentsOfClass a " & vbCrLf
        sqlstr &= " JOIN Stud_StudentInfo b ON a.SID = b.SID " & vbCrLf
        sqlstr &= " JOIN Class_ClassInfo c ON a.OCID = c.OCID " & vbCrLf
        sqlstr &= " JOIN Auth_Relship d ON c.RID = d.RID " & vbCrLf
        sqlstr &= " JOIN Org_OrgInfo e ON d.OrgID = e.OrgID " & vbCrLf
        sqlstr &= " LEFT JOIN ID_District f ON d.DistID = f.DistID " & vbCrLf
        sqlstr &= " LEFT JOIN Key_TrainType g ON c.TMID = g.TMID " & vbCrLf
        sqlstr &= " JOIN ID_Plan i ON i.PlanID = c.PlanID " & vbCrLf
        sqlstr &= " JOIN Key_Plan j ON j.TPlanID = i.TPlanID " & vbCrLf
        sqlstr &= " WHERE 1=1 " & vbCrLf
        sqlstr &= " AND a.MAKESOCID IS NULL " & vbCrLf
        sqlstr &= " AND b.IDNO =@IDNO" & vbCrLf
        dt1 = DbAccess.GetDataTable(sqlstr, objconn, parms3)
        Call sUtl_dt1Data2dtData(dt1, dt)

        Stud_DG.Visible = False
        'Me.PageControler1.Visible = False
        msg.Text = "查無資料!!"
        If dt.Rows.Count > 0 Then
            Stud_DG.Visible = True
            msg.Text = ""

            dt.DefaultView.Sort = "DistID,TPlanID"
            Stud_DG.AllowPaging = False
            Stud_DG.DataSource = dt.DefaultView
            Stud_DG.DataBind()

            'Me.PageControler1.Visible = True
            'PageControler1.SqlDataCreate(sqlstr, "DistID,TPlanID")
            'PageControler1.PageDataTable = dt
            'PageControler1.Sort = "DistID,TPlanID"
            'PageControler1.ControlerLoad()
        End If
    End Sub

    Sub sUtl_CreateDt1(ByRef dt As DataTable)
        dt = New DataTable
        dt.Columns.Add(New DataColumn("StdID")) '序號
        dt.Columns.Add(New DataColumn("DistName")) '轄區分署(轄區中心)
        dt.Columns.Add(New DataColumn("PlanName")) '計畫名稱
        dt.Columns.Add(New DataColumn("TrinUnit")) '訓練機構
        dt.Columns.Add(New DataColumn("ClassName"))  '班別
        dt.Columns.Add(New DataColumn("Name"))  '姓名
        dt.Columns.Add(New DataColumn("Sex"))  '性別
        dt.Columns.Add(New DataColumn("IDNO"))  '身分證號
        dt.Columns.Add(New DataColumn("Ident"))  '身分別
        dt.Columns.Add(New DataColumn("TPlanID"))  '計畫序號
        dt.Columns.Add(New DataColumn("DistID"))  '轄區序號
        dt.Columns.Add(New DataColumn("source"))  '來源
    End Sub

    Sub sUtl_dt1Data2dtData(ByRef dt1 As DataTable, ByRef dt As DataTable)
        For Each dr1 As DataRow In dt1.Rows
            Dim dr As DataRow = dt.NewRow
            dt.Rows.Add(dr)
            dr("StdID") = Convert.ToString(dr1("StdID"))
            dr("DistName") = Convert.ToString(dr1("DistName"))
            dr("PlanName") = Convert.ToString(dr1("PlanName"))
            dr("TrinUnit") = Convert.ToString(dr1("TrinUnit"))
            dr("ClassName") = Convert.ToString(dr1("ClassName"))
            dr("Name") = Convert.ToString(dr1("Name"))
            dr("Sex") = Convert.ToString(dr1("Sex"))
            dr("IDNO") = TIMS.ChangeIDNO(Convert.ToString(dr1("IDNO")))
            dr("Ident") = Convert.ToString(dr1("Ident"))
            dr("TPlanID") = Convert.ToString(dr1("TPlanID"))
            dr("DistID") = Convert.ToString(dr1("DistID"))
            dr("source") = Convert.ToString(dr1("source"))
        Next
    End Sub
End Class