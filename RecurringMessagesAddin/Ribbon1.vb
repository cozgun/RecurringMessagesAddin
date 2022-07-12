Imports System.Diagnostics
Imports System.Windows.Forms
Imports Microsoft.Office.Tools.Ribbon


Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button3_Click(sender As Object, e As RibbonControlEventArgs)
        MsgBox("You can add some macros into this tab")
    End Sub

    Private Sub Button6_Click_1(sender As Object, e As RibbonControlEventArgs) Handles Button6.Click
        MsgBox("You can add some macros into this tab")
    End Sub

    Private Sub CheckBox1_Click(sender As Object, e As RibbonControlEventArgs) Handles Mode_debug.Click

    End Sub

    Private Sub Button8_Click(sender As Object, e As RibbonControlEventArgs) Handles Button8.Click
        MsgBox("You can add some macros into this tab")
    End Sub

    Private Sub Button9_Click(sender As Object, e As RibbonControlEventArgs) Handles Button9.Click
        MsgBox("You can add some macros into this tab")
    End Sub

    Private Sub Button10_Click(sender As Object, e As RibbonControlEventArgs) Handles Button10.Click
        MsgBox("You can add some macros into this tab")
    End Sub

    Private Sub CheckBox1_Click_1(sender As Object, e As RibbonControlEventArgs) Handles ModeOnOff.Click

    End Sub

    Private Sub btnUpdateParams_Click(sender As Object, e As RibbonControlEventArgs) Handles btnUpdateParams.Click
        Call openParameters()
    End Sub
    Function GetAppPath() As String
        Dim i As Integer
        Dim strAppPath As String
        strAppPath = System.Reflection.Assembly.GetExecutingAssembly.Location()
        i = strAppPath.Length - 1
        Do Until strAppPath.Substring(i, 1) = "\"
            i = i - 1
        Loop
        strAppPath = strAppPath.Substring(0, i)
        Return strAppPath
    End Function
    Sub openParameters()

        Dim objShell = CreateObject("Wscript.Shell")
        'Dim appPath As String = GetAppPath()
        'MsgBox("appPath:" & GetAppPath())
        Dim intMessage = MsgBox("Update holidays:" & vbCr _
            & vbCr _
            & "On the first days of every new year, you should update holidays in order to return correct reporting dates to be used in functions." & vbCr _
            & vbCr _
            & vbCr _
            & "Config file to be updated is located @ c:\temp\appSettings.xml.  Click yes to open and update it.",
            vbYesNo, "Update holidays...")
        If intMessage = vbYes Then
            objShell.Run("c:\temp\appSettings.xml")
        Else
        End If
    End Sub
    Private Sub Button7_Click(sender As Object, e As RibbonControlEventArgs) Handles Button7.Click
        Dim fileName = AppDomain.CurrentDomain.BaseDirectory & "SendRecurringEmail.pdf"
        Process.Start(fileName)
    End Sub

    Private Sub Button11_Click(sender As Object, e As RibbonControlEventArgs) Handles Button11.Click
        Dim fileName = AppDomain.CurrentDomain.BaseDirectory & "SendRecurringEmailWithAttachment.pdf"
        Process.Start(fileName)
    End Sub

    Private Sub Button12_Click(sender As Object, e As RibbonControlEventArgs) Handles Button12.Click
        Dim fileName = AppDomain.CurrentDomain.BaseDirectory & "CheckProofFile.pdf"
        Process.Start(fileName)
    End Sub

    Private Sub Button13_Click(sender As Object, e As RibbonControlEventArgs) Handles Button13.Click
        Dim fileName = AppDomain.CurrentDomain.BaseDirectory & "CheckProofFileAndInformOnlyAbsence.pdf"
        Process.Start(fileName)
    End Sub

    Private Sub Button14_Click(sender As Object, e As RibbonControlEventArgs) Handles Button14.Click
        Dim fileName = AppDomain.CurrentDomain.BaseDirectory & "Variables.pdf"
        Process.Start(fileName)
    End Sub

    Private Sub Button15_Click(sender As Object, e As RibbonControlEventArgs) Handles Button15.Click
        MsgBox("There is one more structure ready to use with the addin:" &
            vbNewLine & vbNewLine & "Addin can check incoming e-mail messages." &
            vbNewLine & vbNewLine & "You can add your filters in VS code and trigger some other macros to give start to different processes or perform controls, save files and integrate with other functions.",
            MsgBoxStyle.OkOnly, "Incoming e-mail checks")
    End Sub

End Class
