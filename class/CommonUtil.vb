'Imports Turbo
Imports System.Xml

Public Class CommonUtil

    ' 設定 DataGrid 滑鼠 scroll 顏色
    Public Shared Sub set_row_color(ByRef DG As DataGrid)
        Dim MarkColor As String = "#FFE1A4"
        If DG.Items.Count <= 0 Then
            Exit Sub
        End If
        Dim bgcolor0 As String = "#" & DG.ItemStyle.BackColor.R.ToString("X") & DG.ItemStyle.BackColor.G.ToString("X") & DG.ItemStyle.BackColor.B.ToString("X")
        Dim bgcolor1 As String = "#" & DG.AlternatingItemStyle.BackColor.R.ToString("X") & DG.AlternatingItemStyle.BackColor.G.ToString("X") & DG.AlternatingItemStyle.BackColor.B.ToString("X")
        Dim j As Integer
        For j = 0 To DG.Items.Count - 1
            DG.Items(j).Attributes.Add("onmouseover", "this.style.backgroundColor='" & MarkColor & "';")
            If (j Mod 2) = 0 Then
                DG.Items(j).Attributes.Add("onmouseout", "this.style.backgroundColor='" & bgcolor0 & "';")
            Else
                DG.Items(j).Attributes.Add("onmouseout", "this.style.backgroundColor='" & bgcolor1 & "';")
            End If
        Next
    End Sub

End Class
