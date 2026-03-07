Public Class WUC2
    Inherits System.Web.UI.UserControl

    Const cst_default_page_1 As String = "SD_15_024"
    Const cst_wuc2_page_1 As String = "wuc2_page_1"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            cCreate1()
        End If
    End Sub

    Sub cCreate1()
        '非產投
        With rblobj1
            .Items.Clear()
            '.Items.Add(New ListItem("交叉分析統計表", "CM_03_011")) '(這個base程式) '311:
            .Items.Add(New ListItem("交叉分析統計表", "SD_15_024")) '(這個base程式) '311:非產投

            .Items.Add(New ListItem("年度職業訓練行業別_性別分佈", "TR_05_002_R"))  '502:
            .Items.Add(New ListItem("年度訓練人數統計_依行業別", "TR_05_001_R"))    '501:
            .Items.Add(New ListItem("年度訓練計畫特定對象人數分佈", "TR_05_007_R")) '507:
            .Items.Add(New ListItem("訓練計畫特定對象人數統計表", "TR_05_008_R"))   '508:
            .Items.Add(New ListItem("訓練時數統計分析", "TR_05_010_R")) '510
            .Items.Add(New ListItem("訓練職類統計分析", "TR_05_011_R")) '511

            .Items.Add(New ListItem("訓練人數綜合查詢", "CM_03_003")) '303 '無法整併
            .Items.Add(New ListItem("結訓人數綜合查詢", "CM_03_004")) '304 '無法整併
            .Items.Add(New ListItem("主要特定對象統計表", "CM_03_007")) '307
            .Items.Add(New ListItem("離退訓人數統計表", "CM_03_008")) '308
            '.Items.Add(New ListItem("志願役人數統計表", "CM_03_012"))
        End With
        'Page.AppRelativeTemplateSourceDirectory "~/SD/15/"	String
        'AppRelativeVirtualPath  "~/SD/15/SD_15_024.aspx"	String
        Dim strARTSD As String = Page.AppRelativeTemplateSourceDirectory
        Dim strARVP As String = Page.AppRelativeVirtualPath
        Dim strValue As String = Replace(Replace(UCase(strARVP), UCase(strARTSD), ""), ".ASPX", "")
        If strValue <> "" Then
            Session(cst_wuc2_page_1) = strValue
            Common.SetListItem(rblobj1, strValue)
        Else
            If Session(cst_wuc2_page_1) IsNot Nothing Then
                Common.SetListItem(rblobj1, Session(cst_wuc2_page_1))
            Else
                Session(cst_wuc2_page_1) = cst_default_page_1
                Common.SetListItem(rblobj1, cst_default_page_1)
            End If
        End If
    End Sub

    Protected Sub rblobj1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles rblobj1.SelectedIndexChanged
        If rblobj1 Is Nothing Then Exit Sub
        Dim vUrbl1 As String = TIMS.ClearSQM(rblobj1.SelectedValue)
        If vUrbl1 <> "" AndAlso vUrbl1.Length > 5 Then
            Session(cst_wuc2_page_1) = vUrbl1
            Dim vU1 As String = "~/" & Left(vUrbl1, 2) & "/" & Right(Left(vUrbl1, 5), 2) & "/"
            TIMS.Utl_Redirect1(Page, vU1 & vUrbl1)
            'TIMS.Utl_Redirect1(Page, vUrbl1)
        End If
    End Sub
End Class
