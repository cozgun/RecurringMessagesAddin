Partial Class Ribbon1
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.ModeOnOff = Me.Factory.CreateRibbonCheckBox
        Me.Mode_debug = Me.Factory.CreateRibbonCheckBox
        Me.Group4 = Me.Factory.CreateRibbonGroup
        Me.Button7 = Me.Factory.CreateRibbonButton
        Me.Button11 = Me.Factory.CreateRibbonButton
        Me.Button12 = Me.Factory.CreateRibbonButton
        Me.Button13 = Me.Factory.CreateRibbonButton
        Me.Button14 = Me.Factory.CreateRibbonButton
        Me.Button15 = Me.Factory.CreateRibbonButton
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.btnUpdateParams = Me.Factory.CreateRibbonButton
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.ButtonGroup1 = Me.Factory.CreateRibbonButtonGroup
        Me.Button8 = Me.Factory.CreateRibbonButton
        Me.Button9 = Me.Factory.CreateRibbonButton
        Me.Button10 = Me.Factory.CreateRibbonButton
        Me.Button6 = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.Group4.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.ButtonGroup1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.Group2)
        Me.Tab1.Groups.Add(Me.Group3)
        Me.Tab1.Groups.Add(Me.Group4)
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Label = "Recurring Messages Addin"
        Me.Tab1.Name = "Tab1"
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.ModeOnOff)
        Me.Group2.Items.Add(Me.Mode_debug)
        Me.Group2.Label = "Main switch"
        Me.Group2.Name = "Group2"
        '
        'ModeOnOff
        '
        Me.ModeOnOff.Label = "Addin On/Off"
        Me.ModeOnOff.Name = "ModeOnOff"
        '
        'Mode_debug
        '
        Me.Mode_debug.Label = "Debug mode"
        Me.Mode_debug.Name = "Mode_debug"
        Me.Mode_debug.Visible = False
        '
        'Group4
        '
        Me.Group4.Items.Add(Me.Button7)
        Me.Group4.Items.Add(Me.Button11)
        Me.Group4.Items.Add(Me.Button12)
        Me.Group4.Items.Add(Me.Button13)
        Me.Group4.Items.Add(Me.Button14)
        Me.Group4.Items.Add(Me.Button15)
        Me.Group4.Label = "Get help on functions"
        Me.Group4.Name = "Group4"
        '
        'Button7
        '
        Me.Button7.Label = "Send Recurring Email"
        Me.Button7.Name = "Button7"
        '
        'Button11
        '
        Me.Button11.Label = "Send Email With Attachment"
        Me.Button11.Name = "Button11"
        '
        'Button12
        '
        Me.Button12.Label = "Check Proof File"
        Me.Button12.Name = "Button12"
        '
        'Button13
        '
        Me.Button13.Label = "Check Proof And Inform Only Absence"
        Me.Button13.Name = "Button13"
        '
        'Button14
        '
        Me.Button14.Label = "Dynamic Variables"
        Me.Button14.Name = "Button14"
        '
        'Button15
        '
        Me.Button15.Label = "Incoming Mail Checker"
        Me.Button15.Name = "Button15"
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.btnUpdateParams)
        Me.Group3.Label = "Manage parameters"
        Me.Group3.Name = "Group3"
        '
        'btnUpdateParams
        '
        Me.btnUpdateParams.Label = "Update holidays"
        Me.btnUpdateParams.Name = "btnUpdateParams"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.ButtonGroup1)
        Me.Group1.Items.Add(Me.Button6)
        Me.Group1.Label = "Run macros manually"
        Me.Group1.Name = "Group1"
        '
        'ButtonGroup1
        '
        Me.ButtonGroup1.Items.Add(Me.Button8)
        Me.ButtonGroup1.Items.Add(Me.Button9)
        Me.ButtonGroup1.Items.Add(Me.Button10)
        Me.ButtonGroup1.Name = "ButtonGroup1"
        '
        'Button8
        '
        Me.Button8.Label = "Macro #1"
        Me.Button8.Name = "Button8"
        '
        'Button9
        '
        Me.Button9.Label = "Macro #2"
        Me.Button9.Name = "Button9"
        '
        'Button10
        '
        Me.Button10.Label = "Macro #3"
        Me.Button10.Name = "Button10"
        '
        'Button6
        '
        Me.Button6.Label = "You can add your frequent items here"
        Me.Button6.Name = "Button6"
        Me.Button6.ScreenTip = "You can add here a shortcut for your most sent e-mails"
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.Outlook.Appointment, Microsoft.Outlook.Explorer"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.Group4.ResumeLayout(False)
        Me.Group4.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.ButtonGroup1.ResumeLayout(False)
        Me.ButtonGroup1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button6 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Mode_debug As Microsoft.Office.Tools.Ribbon.RibbonCheckBox
    Friend WithEvents ButtonGroup1 As Microsoft.Office.Tools.Ribbon.RibbonButtonGroup
    Friend WithEvents Button8 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button9 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button10 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ModeOnOff As Microsoft.Office.Tools.Ribbon.RibbonCheckBox
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnUpdateParams As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group4 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button12 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button7 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button11 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button13 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button14 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button15 As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
