Public Class WUC1
    Inherits System.Web.UI.UserControl

    Const cst_default_page_1 As String = "SD_15_023"
    Const cst_wuc1_page_1 As String = "wuc1_page_1"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            cCreate1()
        End If
    End Sub

    Sub cCreate1()
        '產投
        With rblobj1
            .Items.Clear()
            .Items.Add(New ListItem("訓練人數統計表", "SD_15_001"))
            .Items.Add(New ListItem("交叉分析統計表", "SD_15_023")) 'base:SD_15_003(移除)-產投
            .Items.Add(New ListItem("綜合查詢統計表", "SD_15_012"))
            .Items.Add(New ListItem("開訓統計週報表", "SD_15_011"))
        End With

        'Page.AppRelativeTemplateSourceDirectory "~/SD/15/"	String
        'AppRelativeVirtualPath  "~/SD/15/SD_15_024.aspx"	String
        Dim strARTSD As String = Page.AppRelativeTemplateSourceDirectory
        Dim strARVP As String = Page.AppRelativeVirtualPath
        Dim strValue As String = Replace(Replace(UCase(strARVP), UCase(strARTSD), ""), ".ASPX", "")
        If strValue <> "" Then
            Session(cst_wuc1_page_1) = strValue
            Common.SetListItem(rblobj1, strValue)
        Else
            If Session(cst_wuc1_page_1) IsNot Nothing Then
                Common.SetListItem(rblobj1, Session(cst_wuc1_page_1))
            Else
                Session(cst_wuc1_page_1) = cst_default_page_1
                Common.SetListItem(rblobj1, cst_default_page_1)
            End If
        End If
    End Sub

    Protected Sub rblobj1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles rblobj1.SelectedIndexChanged
        If rblobj1 Is Nothing Then Exit Sub
        Dim vUrbl1 As String = TIMS.ClearSQM(rblobj1.SelectedValue)
        If vUrbl1 <> "" AndAlso vUrbl1.Length > 5 Then
            Session(cst_wuc1_page_1) = vUrbl1
            Dim vU1 As String = "~/" & Left(vUrbl1, 2) & "/" & Right(Left(vUrbl1, 5), 2) & "/"
            TIMS.Utl_Redirect1(Page, vU1 & vUrbl1)
            'TIMS.Utl_Redirect1(Page, vUrbl1)
        End If

    End Sub
End Class
