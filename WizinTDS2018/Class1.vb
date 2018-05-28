
Imports System.Drawing.Text
    Public Class Form1
        Dim pfc As New PrivateFontCollection()
    Private Function FindALLControlRecursive(ByVal list As List(Of Control), ByVal parent As Control) As List(Of Control)
        ' function that returns all control in a form, parent or child regardless of control's type
        If parent Is Nothing Then
            Return list
        Else
            list.Add(parent)
        End If
        For Each child As Control In parent.Controls
            FindALLControlRecursive(list, child)
        Next
        Return list
    End Function
    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
            ' On Form1 shown, start applying font 
            Dim CFontPath As String = Application.StartupPath
        'pfc.AddFontFile(CFontPath & "\Resources\Fonts\Roboto.ttf")
        Dim allCtrl As New List(Of Control)
            For Each ctrl As Control In FindALLControlRecursive(allCtrl, Me)
            ' You need to define which control type to change it's font family; not recommendd to just change all controls' fonts, it will create a missy shape
            If TypeOf ctrl Is Label Or TypeOf ctrl Is TextBox Or TypeOf ctrl Is Button Or TypeOf ctrl Is CheckBox Or TypeOf ctrl Is RadioButton Or TypeOf ctrl Is ProgressBar Or TypeOf ctrl Is GroupBox Then
                Dim CurrentCtrlFontSize = ctrl.Font.Size ' get current object's font size before applying new font family
                ctrl.Font = New Font("Microsoft Sans Serif", CurrentCtrlFontSize, FontStyle.Regular)
            End If
        Next
            allCtrl.Clear()
        End Sub
    End Class

