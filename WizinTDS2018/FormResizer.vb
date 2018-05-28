Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Windows.Forms
Imports System.Drawing
Imports System.ComponentModel

Public Class FormResizer

    'Considerations:
    'Change the Form AutoSize Mode to None.
    Private f_HeightRatio As New Single()
    Private f_WidthRatio As New Single()

    Public Sub ResizeForm(ByVal ObjForm As Form, ByVal DesignerWidth As Integer, ByVal DesignerHeight As Integer)
        '#Region "Code for Resizing and Font Change According to Resolution"
        'Specify Here the Resolution Y component in which this form is designed
        'For Example if the Form is Designed at 800 * 600 Resolution then DesignerHeight=600
        Dim i_StandardHeight As Integer = DesignerHeight

        'Specify Here the Resolution X component in which this form is designed
        'For Example if the Form is Designed at 800 * 600 Resolution then DesignerWidth=800
        Dim i_StandardWidth As Integer = DesignerWidth
        Dim i_PresentHeight As Integer = Screen.PrimaryScreen.Bounds.Height

        'Present Resolution Height
        Dim i_PresentWidth As Integer = Screen.PrimaryScreen.Bounds.Width

        'Presnet Resolution Width
        f_HeightRatio = CSng(CSng(i_PresentHeight) / CSng(i_StandardHeight))
        f_WidthRatio = CSng(CSng(i_PresentWidth) / CSng(i_StandardWidth))
        ObjForm.AutoScaleMode = AutoScaleMode.None

        'Make the Autoscale Mode=None
        ObjForm.Scale(New SizeF(f_WidthRatio, f_HeightRatio))
        For Each c As Control In ObjForm.Controls
            If c.HasChildren Then
                ResizeControlStore(c)
            Else
                c.Font = New Font(c.Font.FontFamily, c.Font.Size * f_HeightRatio, c.Font.Style, c.Font.Unit, CByte(0))
            End If
        Next
        ObjForm.Font = New Font(ObjForm.Font.FontFamily, ObjForm.Font.Size * f_HeightRatio, ObjForm.Font.Style, ObjForm.Font.Unit, CByte(0))
    End Sub

    ''' <summary>
    ''' This Function is Used to Change the Font of Controls that are Nested in Other Controls.
    ''' </summary>
    ''' <param name="objCtl"></param>
    Public Sub ResizeControlStore(ByVal objCtl As Control)
        If objCtl.HasChildren Then
            For Each cChildren As Control In objCtl.Controls
                If cChildren.HasChildren Then
                    ResizeControlStore(cChildren)
                Else
                    cChildren.Font = New Font(cChildren.Font.FontFamily, cChildren.Font.Size * f_HeightRatio, cChildren.Font.Style, cChildren.Font.Unit, CByte(0))
                End If
            Next
            objCtl.Font = New Font(objCtl.Font.FontFamily, objCtl.Font.Size * f_HeightRatio, objCtl.Font.Style, objCtl.Font.Unit, CByte(0))
        Else
            objCtl.Font = New Font(objCtl.Font.FontFamily, objCtl.Font.Size * f_HeightRatio, objCtl.Font.Style, objCtl.Font.Unit, CByte(0))
        End If
    End Sub
End Class
