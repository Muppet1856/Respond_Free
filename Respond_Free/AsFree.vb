Imports Microsoft.Office.Tools.Ribbon
Public Class AsFree
    Private Sub AsFree_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub
    Private Sub AcceptAsFree_Click(sender As Object, e As RibbonControlEventArgs) Handles AcceptAsFree.Click, Button1.Click
        Call ThisAddIn.AcceptFree()
    End Sub
    Private Sub TentativeAsFree_Click(sender As Object, e As RibbonControlEventArgs) Handles TentativeAsFree.Click, Button2.Click
        Call ThisAddIn.TentAcceptFree()
    End Sub
    Private Sub DeclineAsFree_Click(sender As Object, e As RibbonControlEventArgs) Handles DeclineAsFree.Click, Button3.Click
        Call ThisAddIn.DeclineFree()
    End Sub
End Class
