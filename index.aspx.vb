Imports System.Web.Mvc

Public Class index
    Inherits AuthBasePage
    Private logger As ILog = LogManager.GetLogger(GetType(index))

    Const cst_nullplanmsg1 As String = "請選擇計畫!!"
    Const cst_errmsg1 As String = "請重新選擇計畫!!"
    Const cst_errmsg2 As String = "該使用者/計畫無此功能，請重新選擇計畫!!"
    'Const cst_errmsg3 As String = "該使用者暫時停用，請稍後再試!!"

    Dim intKind As Integer = -1
    Dim tnCnt As Integer = 0

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)

        '防止駭客攻擊(紀錄) 啟動-Login / true:攻擊異常達標 false:未達標
        Dim flag_ChkHISTORY1 As Boolean = TIMS.sUtl_ChkHISTORY1(objconn)
        If flag_ChkHISTORY1 Then
            Dim strErrmsg1 As String = ""
            strErrmsg1 = "防止駭客攻擊(紀錄) 啟動-index.Page_Load" & vbCrLf
            strErrmsg1 &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            strErrmsg1 &= String.Concat("GetHTTP_HOST:", vbCrLf, TIMS.GetHTTP_HOST(Me), vbCrLf)
            'strErrmsg1 = Replace(strErrmsg1, vbCrLf, "<br>" & vbCrLf)
            logger.Warn(strErrmsg1)

            sm.LastErrorMessage = TIMS.cst_ErrorMsg1
            'AuthUtil.LoginLog(txtIDNO.Text, False)

            Const cst_iMaxCanMailCount As Integer = 30 '(寄狀況信數量)
            Dim iGlobalMailCount As Integer = TIMS.GlobalMailCount '目前寄信總數量
            '(狀況信)超過n:30
            If iGlobalMailCount >= cst_iMaxCanMailCount Then
                '(取得使用者正確ip)
                Dim v_IpAddress As String = Common.GetIpAddress() 'MyPage.Request.UserHostAddress
                Dim i_ChkHISTORY1_cnt As Integer = TIMS.SUtl_ChkHISTORY1_CNT(v_IpAddress)
                If i_ChkHISTORY1_cnt > 2 Then
                    TIMS.sUtl_404NOTFOUND(Me, objconn, i_ChkHISTORY1_cnt)
                Else
                    TIMS.sUtl_404NOTFOUND(Me, objconn)
                End If
            End If
            '登出/ 重登
            AuthUtil.LogoutLog()
            '清除登入狀態
            sm.ClearSession()
            Return 'redirectUrl
        End If

        Labmsg1.Text = TIMS.Get_LOCALADDR(Me, 2)
        'If sm.IsLogin AndAlso (Hid_idx99_tplanid.Value = "" OrElse Convert.ToString(ViewState("Hid_idx99_tplanid")) = "") Then
        '    ViewState("Hid_idx99_tplanid") = sm.UserInfo.TPlanID
        '    Hid_idx99_tplanid.Value = ViewState("Hid_idx99_tplanid")
        'End If
        'If Convert.ToString(ViewState("Hid_idx99_tplanid")) <> "" AndAlso Hid_idx99_tplanid.Value = "" Then
        '    Hid_idx99_tplanid.Value = ViewState("Hid_idx99_tplanid")
        'End If

        If Not IsPostBack Then
            ShowLoginUserInfo()
            ShowMenu()
        End If
    End Sub

    ' 抓取登入者的資訊
    Private Sub ShowLoginUserInfo()

        Me.labID.Text = ""   '如: 分署(職業訓練中心) > 系統測試員 > 測試員
        Me.labPlan.Text = cst_nullplanmsg1 '"" '如: 2011北區職業訓練中心自辦職前訓練001

        If sm.UserInfo.UserID Is Nothing Then Exit Sub
        If sm.UserInfo.PlanID Is Nothing Then Exit Sub

        Dim sql As String = ""
        If Convert.ToString(sm.UserInfo.PlanID) <> "0" Then
            sql &= " select b.name UserName" & vbCrLf
            sql &= " ,d.name UserRole" & vbCrLf
            sql &= " ,c.Years+e.Name+f.PlanName+c.seq+ ISNULL(c.SubTitle,'') UserPlan" & vbCrLf
            sql &= " FROM AUTH_ACCRWPLAN a" & vbCrLf
            sql &= " join Auth_Account b on a.Account=b.Account" & vbCrLf
            sql &= " join ID_Plan c on a.PlanID=c.PlanID" & vbCrLf
            sql &= " join ID_Role d on b.RoleID=d.RoleID" & vbCrLf
            sql &= " join ID_District e on c.DistID=e.DistID" & vbCrLf
            sql &= " join Key_Plan f on c.TPlanID=f.TPlanID" & vbCrLf
            sql &= " where a.Account=@Account" & vbCrLf
            sql &= " and a.PlanID =@PlanID" & vbCrLf

        Else
            sql &= " select b.name UserName" & vbCrLf
            sql &= " ,c.name UserRole " & vbCrLf
            sql &= " ,null UserPlan " & vbCrLf
            sql &= " FROM AUTH_ACCRWPLAN a" & vbCrLf
            sql &= " join Auth_Account b on a.Account=b.Account" & vbCrLf
            sql &= " join ID_Role c on b.RoleID=c.RoleID" & vbCrLf
            sql &= " where a.Account=@Account" & vbCrLf
            sql &= " and a.PlanID =@PlanID" & vbCrLf

        End If

        Dim parms As New Hashtable From {{"Account", sm.UserInfo.UserID}, {"PlanID", sm.UserInfo.PlanID}}
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)

        If TIMS.dtHaveDATA(dt) Then
            Dim dr As DataRow = dt.Rows(0)
            Dim sOrgName As String = sm.UserInfo.OrgName
            Dim sUserRole As String = Convert.ToString(dr("UserRole"))
            Dim sUserName As String = Convert.ToString(dr("UserName"))
            labID.Text = String.Concat(sOrgName, " > ", sUserRole, " > ", sUserName)
            labPlan.Text = Convert.ToString(dr("UserPlan"))
        End If
    End Sub

#Region "功能選單"

    '建置功能列
    Private Sub ShowMenu()

        '群組查詢
        Dim dt_group As DataTable = TIMS.sGet_CanUseSchDt(objconn) ' FROM 

        Dim blnGroupF As Boolean = True '查詢有資料。
        If dt_group Is Nothing Then
            blnGroupF = False '查詢無資料。
        Else
            dt_group.DefaultView.RowFilter = "Valid='Y'"
            dt_group = TIMS.dv2dt(dt_group.DefaultView)
            If dt_group.Rows.Count = 0 Then blnGroupF = False '查詢無資料。
        End If
        If Not blnGroupF Then
            sm.LastErrorMessage = cst_errmsg2
            Exit Sub
        End If
        If blnGroupF Then
            '有資料執行顯示菜單功能。
            Dim menuTree As TagBuilder = AddTreeView(dt_group)
            Me.ulMenu.InnerHtml = menuTree.InnerHtml
        End If
    End Sub

    '檢查是否有功能子集。
    Function ChkKindDataTable(ByRef dt As DataTable, ByVal Kind As String) As Boolean
        Dim rst As Boolean = True
        If dt.Select("Kind='" & Kind & "' and Levels=0", "newSort").Length = 0 Then
            rst = False
        End If
        Return rst
    End Function

    '主menu顯示順序
    Private Function AddTreeView(ByVal dt As DataTable) As TagBuilder
        ' 功能類別碼定義
        'Dim fKinds As String() = {"TC", "SD", "CP", "TR", "CM", "FM", "OB", "SE", "EXAM", "SV", "SYS", "FAQ", "OO"}
        Dim fKinds As String() = TIMS.c_FUNSORT.Split(",")
        Dim fKind As String

        Dim tree As TagBuilder = New TagBuilder("ul")

        For Each fKind In fKinds
            ''檢查是否有功能子集
            If ChkKindDataTable(dt, fKind) Then
                tree.InnerHtml &= AddTitleNode(dt, fKind)
            End If
        Next

        Return tree
    End Function

    ''' <summary>
    ''' 產生第一層的功能及其子功能
    ''' </summary>
    ''' <param name="dt"></param>
    ''' <param name="Kind"></param>
    ''' <returns></returns>
    Private Function AddTitleNode(ByVal dt As DataTable, ByVal Kind As String) As String
        Dim result As String = String.Empty

        If dt.Select("Kind='" & Kind & "' and Levels=0").Length <> 0 Or Kind = "FAQ" Then
            Dim MyNode As TagBuilder = New TagBuilder("li")
            MyNode.AddCssClass("has-sub")

            Dim nodeLink As TagBuilder = New TagBuilder("a")
            Dim childHtml As String = ""

            If (Kind = "SE") Then
                Dim vsStudExamUrl As String = ConfigurationSettings.AppSettings("StudExamUrl")
                nodeLink.Attributes.Add("href", vsStudExamUrl)
                nodeLink.Attributes.Add("target", "_new")
            Else
                nodeLink.Attributes.Add("href", "javascript:void(0);")
                ' 建子節點
                childHtml = AddNode(dt, Kind)
            End If

            '功能類別對照 取得中文名稱
            Dim rst As String = TIMS.Get_MainMenuName(UCase(Kind))
            If rst <> "" Then
                nodeLink.InnerHtml = rst
            End If

            'logger.Debug("AddTitleNode: " & vbCrLf & nodeLink.InnerHtml & " " & nodeLink.Attributes("href"))

            MyNode.InnerHtml = nodeLink.ToString()

            If childHtml <> "" Then
                ' 存在子節點內容(li)
                Dim childNode As TagBuilder = New TagBuilder("ul")
                childNode.InnerHtml = childHtml
                childNode.Attributes.Add("class", "hide")

                MyNode.InnerHtml &= childNode.ToString()
            End If

            result = MyNode.ToString() & vbCrLf
        End If

        'If (result.IndexOf("&") > -1) Then result = Replace(result, "&", "＆")

        Return result
    End Function

    '子節點樣式
    Private Function AddNode(ByVal dt As DataTable, ByVal Kind As String, Optional ByVal drChild As DataRow = Nothing) As String
        Dim vsUserID As String = "" & Convert.ToString(sm.UserInfo.UserID)
        Dim vsDistID As String = "" & Convert.ToString(sm.UserInfo.DistID)
        Dim vsLID As String = "" & Convert.ToString(sm.UserInfo.LID)
        If vsLID <> "" Then vsLID = CInt(vsLID)

        Dim dr As DataRow
        Dim MyNode As TagBuilder = Nothing
        Dim result As String = ""

        'Const cst_FunLength_set As Integer = 7 '10
        If drChild Is Nothing Then
            For Each dr In dt.Select("Kind='" & Kind & "' and Levels=0", "newSort")
                Dim FunName As String = "" '上層功能名稱  Levels=0 
                FunName = dr("Name").ToString
                If FunName <> "" Then FunName = Trim(FunName)
                If (FunName.IndexOf("&") > -1) Then FunName = Replace(FunName, "&", "＆")
                FunName = TIMS.ClearSQM(FunName)

                MyNode = New TagBuilder("li")
                Dim nodeLink As TagBuilder = New TagBuilder("a")
                Dim childHtml As String = ""

                nodeLink.InnerHtml = FunName

                Dim vsIDNO As String = ""
                Dim vsuser_ID As String = ""
                Dim vsWOassPrd As String = ""

                If IsDBNull(dr("SPage")) Then
                    '父層無子頁
                    nodeLink.Attributes.Add("href", "javascript:void(0);")
                Else
                    Select Case Kind
                        Case "OO"
                            '其他系統 另外開視窗
                            nodeLink.Attributes.Add("href", dr("SPage").ToString)
                            nodeLink.Attributes.Add("target", "_new")
                        Case Else
                            '連線系統本身功能
                            Dim sPage As String = Convert.ToString(dr("SPage"))
                            If sPage.EndsWith(".aspx", StringComparison.OrdinalIgnoreCase) Then
                                sPage = sPage.Substring(0, sPage.Length - 5)
                            End If
                            'Dim url As String = Convert.ToString(dr("SPage")) & "?ID=" & dr("FunID").ToString
                            Dim url As String = "menuClick(""" & sPage & """,""" & dr("FunID").ToString & """);"
                            nodeLink.Attributes.Add("href", "javascript:void(0);")
                            nodeLink.Attributes.Add("onclick", url)
                            nodeLink.Attributes.Add("class", "func")
                            nodeLink.Attributes.Add("target", "mainFrame")
                    End Select
                End If

                'logger.Debug("AddNode-1 (" & dr("Kind") & "): " & nodeLink.InnerHtml & " " & nodeLink.Attributes("href"))

                If dt.Select("Kind='" & Kind & "' and [Parent]='" & dr("FunID") & "'").Length <> 0 Then
                    ' 建子節點
                    childHtml = AddNode(dt, Kind, dr)
                End If

                MyNode.InnerHtml = nodeLink.ToString()

                If childHtml <> "" Then
                    ' 存在子節點內容(li)
                    Dim childNode As TagBuilder = New TagBuilder("ul")
                    childNode.InnerHtml = childHtml
                    childNode.Attributes.Add("class", "hide")

                    MyNode.InnerHtml &= childNode.ToString()
                    MyNode.Attributes.Add("class", "has-sub")
                End If

                ' 當前節點, 加到結果中
                result &= MyNode.ToString() & vbCrLf

                tnCnt += 1
            Next
        Else
            For Each dr In dt.Select("Kind='" & Kind & "' and [Parent]='" & drChild("FunID") & "'", "newSort")

                Dim FunName As String = "" '下層功能名稱  Levels=0 
                FunName = dr("Name").ToString
                If FunName <> "" Then FunName = Trim(FunName)
                If (FunName.IndexOf("&") > -1) Then FunName = Replace(FunName, "&", "＆")
                FunName = TIMS.ClearSQM(FunName)

                MyNode = New TagBuilder("li")
                Dim nodeLink As TagBuilder = New TagBuilder("a")
                Dim childHtml As String = ""

                nodeLink.InnerHtml = FunName

                '連線系統本身功能
                Dim sPage As String = ""
                If Not IsDBNull(dr("SPage")) Then sPage = Convert.ToString(dr("SPage"))
                If sPage = "" Then
                    '無子頁
                    nodeLink.Attributes.Add("href", "javascript:void(0);")
                Else
                    '含有http字眼另開視窗
                    Dim flag_new_windows As Boolean = TIMS.CHK_FUNCSPAGE2(sm, sPage)

                    If flag_new_windows Then
                        nodeLink.Attributes.Add("href", sPage)
                        nodeLink.Attributes.Add("class", "func")
                        nodeLink.Attributes.Add("target", "_blank")
                    Else
                        If sPage.EndsWith(".aspx", StringComparison.OrdinalIgnoreCase) Then sPage = sPage.Substring(0, sPage.Length - 5)
                        Dim url As String = "menuClick(""" & sPage & """,""" & dr("FunID").ToString & """);"
                        nodeLink.Attributes.Add("href", "javascript:void(0);")
                        nodeLink.Attributes.Add("onclick", url)
                        nodeLink.Attributes.Add("class", "func")
                        nodeLink.Attributes.Add("target", "mainFrame")
                    End If
                End If

                'logger.Debug("AddNode-2 (" & dr("Kind") & "): " & nodeLink.InnerHtml & " " & nodeLink.Attributes("href"))

                If dt.Select("Kind='" & Kind & "' and [Parent]='" & dr("FunID") & "'").Length <> 0 Then
                    ' 建子節點
                    childHtml = AddNode(dt, Kind, dr)
                End If

                MyNode.InnerHtml = nodeLink.ToString()

                If childHtml <> "" Then
                    ' 存在子節點內容(li)
                    Dim childNode As TagBuilder = New TagBuilder("ul")
                    childNode.InnerHtml = childHtml
                    childNode.Attributes.Add("class", "hide")

                    MyNode.InnerHtml &= childNode.ToString()
                    MyNode.Attributes.Add("class", "has-sub")
                End If

                ' 當前節點, 加到結果中
                result &= MyNode.ToString() & vbCrLf

                tnCnt += 1
            Next
        End If

        Return result
    End Function

#End Region


End Class