Partial Class AsFree
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.TabAppointment = Me.Factory.CreateRibbonTab
        Me.AsFreeGrp = Me.Factory.CreateRibbonGroup
        Me.AcceptAsFree = Me.Factory.CreateRibbonButton
        Me.TentativeAsFree = Me.Factory.CreateRibbonButton
        Me.DeclineAsFree = Me.Factory.CreateRibbonButton
        Me.TabReadMessage = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.Button1 = Me.Factory.CreateRibbonButton
        Me.Button2 = Me.Factory.CreateRibbonButton
        Me.Button3 = Me.Factory.CreateRibbonButton
        Me.TabAppointment.SuspendLayout()
        Me.AsFreeGrp.SuspendLayout()
        Me.TabReadMessage.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabAppointment
        '
        Me.TabAppointment.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.TabAppointment.ControlId.OfficeId = "TabAppointment"
        Me.TabAppointment.Groups.Add(Me.AsFreeGrp)
        Me.TabAppointment.Label = "TabAppointment"
        Me.TabAppointment.Name = "TabAppointment"
        '
        'AsFreeGrp
        '
        Me.AsFreeGrp.Items.Add(Me.AcceptAsFree)
        Me.AsFreeGrp.Items.Add(Me.TentativeAsFree)
        Me.AsFreeGrp.Items.Add(Me.DeclineAsFree)
        Me.AsFreeGrp.Label = "Free Response"
        Me.AsFreeGrp.Name = "AsFreeGrp"
        Me.AsFreeGrp.Position = Me.Factory.RibbonPosition.AfterOfficeId("GroupRespond")
        '
        'AcceptAsFree
        '
        Me.AcceptAsFree.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.AcceptAsFree.Label = "Accept As Free"
        Me.AcceptAsFree.Name = "AcceptAsFree"
        Me.AcceptAsFree.OfficeImageId = "AcceptInvitation"
        Me.AcceptAsFree.ShowImage = True
        '
        'TentativeAsFree
        '
        Me.TentativeAsFree.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.TentativeAsFree.Label = "Tentative As Free"
        Me.TentativeAsFree.Name = "TentativeAsFree"
        Me.TentativeAsFree.OfficeImageId = "TentativeAcceptInvitation"
        Me.TentativeAsFree.ShowImage = True
        '
        'DeclineAsFree
        '
        Me.DeclineAsFree.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.DeclineAsFree.Label = "Decline As Free"
        Me.DeclineAsFree.Name = "DeclineAsFree"
        Me.DeclineAsFree.OfficeImageId = "DeclineInvitation"
        Me.DeclineAsFree.ShowImage = True
        '
        'TabReadMessage
        '
        Me.TabReadMessage.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.TabReadMessage.ControlId.OfficeId = "TabReadMessage"
        Me.TabReadMessage.Groups.Add(Me.Group1)
        Me.TabReadMessage.Label = "TabReadMessage"
        Me.TabReadMessage.Name = "TabReadMessage"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.Button1)
        Me.Group1.Items.Add(Me.Button2)
        Me.Group1.Items.Add(Me.Button3)
        Me.Group1.Label = "Free Response"
        Me.Group1.Name = "Group1"
        Me.Group1.Position = Me.Factory.RibbonPosition.AfterOfficeId("GroupRespond")
        '
        'Button1
        '
        Me.Button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button1.Label = "Accept As Free"
        Me.Button1.Name = "Button1"
        Me.Button1.OfficeImageId = "AcceptInvitation"
        Me.Button1.ShowImage = True
        '
        'Button2
        '
        Me.Button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button2.Label = "Tentative As Free"
        Me.Button2.Name = "Button2"
        Me.Button2.OfficeImageId = "TentativeAcceptInvitation"
        Me.Button2.ShowImage = True
        '
        'Button3
        '
        Me.Button3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button3.Label = "Decline As Free"
        Me.Button3.Name = "Button3"
        Me.Button3.OfficeImageId = "DeclineInvitation"
        Me.Button3.ShowImage = True
        '
        'AsFree
        '
        Me.Name = "AsFree"
        Me.RibbonType = "Microsoft.Outlook.Appointment, Microsoft.Outlook.Explorer, Microsoft.Outlook.Meet" &
    "ingRequest.Read"
        Me.Tabs.Add(Me.TabAppointment)
        Me.Tabs.Add(Me.TabReadMessage)
        Me.TabAppointment.ResumeLayout(False)
        Me.TabAppointment.PerformLayout()
        Me.AsFreeGrp.ResumeLayout(False)
        Me.AsFreeGrp.PerformLayout()
        Me.TabReadMessage.ResumeLayout(False)
        Me.TabReadMessage.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TabAppointment As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents AsFreeGrp As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents AcceptAsFree As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents TentativeAsFree As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents TabReadMessage As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents DeclineAsFree As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button3 As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property AsFree() As AsFree
        Get
            Return Me.GetRibbon(Of AsFree)()
        End Get
    End Property
End Class
