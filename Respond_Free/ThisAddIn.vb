Public Class ThisAddIn
    Private Sub ThisAddIn_Startup(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Startup
    End Sub
    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
    End Sub
    Public Shared Sub AcceptFree()

        Dim objItem As Object
        Dim strID As String
        Dim olNS As Outlook.NameSpace
        Dim oMeetingItem As Outlook.MeetingItem
        Dim oResponse As Outlook.AppointmentItem
        Dim oAppointment As Outlook.AppointmentItem

        objItem = GetCurrentItem()
        strID = objItem.EntryID

        olNS = Globals.ThisAddIn.Application.GetNamespace("MAPI")

        Select Case True
            Case TypeOf objItem Is Outlook.MeetingItem
                oMeetingItem = olNS.GetItemFromID(strID)
                oAppointment = oMeetingItem.GetAssociatedAppointment(True)
            Case TypeOf objItem Is Outlook.AppointmentItem
                oAppointment = objItem
            Case Else
                MsgBox("Not an appointment")
                Exit Sub
        End Select

        oResponse = oAppointment.Respond(Outlook.OlMeetingResponse.olMeetingAccepted, False, True)
        'oResponse.Send()
        oAppointment.BusyStatus = Outlook.OlBusyStatus.olFree

        Select Case True
            Case TypeOf objItem Is Outlook.MeetingItem
                oAppointment.Save()
                oMeetingItem.Save()
                oMeetingItem = Nothing
            Case TypeOf objItem Is Outlook.AppointmentItem
                oAppointment.Save()
            Case Else
                MsgBox("Not an appointment")
                Exit Sub
        End Select
        'Globals.ThisAddIn.Application.ActiveInspector().CurrentItem.Close(Outlook.OlInspectorClose.olSave)
        oAppointment = Nothing
        olNS = Nothing
    End Sub
    Public Shared Sub TentAcceptFree()

        Dim objItem As Object
        Dim strID As String
        Dim olNS As Outlook.NameSpace
        Dim oMeetingItem As Outlook.MeetingItem
        Dim oResponse As Outlook.AppointmentItem
        Dim oAppointment As Outlook.AppointmentItem

        objItem = GetCurrentItem()
        strID = objItem.EntryID

        olNS = Globals.ThisAddIn.Application.GetNamespace("MAPI")

        Select Case True
            Case TypeOf objItem Is Outlook.MeetingItem
                oMeetingItem = olNS.GetItemFromID(strID)
                oAppointment = oMeetingItem.GetAssociatedAppointment(True)
            Case TypeOf objItem Is Outlook.AppointmentItem
                oAppointment = objItem
            Case Else
                MsgBox("Not an appointment")
                Exit Sub
        End Select

        oResponse = oAppointment.Respond(Outlook.OlMeetingResponse.olMeetingTentative, False, True)
        'oResponse.Send()
        oAppointment.BusyStatus = Outlook.OlBusyStatus.olFree

        Select Case True
            Case TypeOf objItem Is Outlook.MeetingItem
                oAppointment.Save()
                oMeetingItem.Save()
                oMeetingItem = Nothing
            Case TypeOf objItem Is Outlook.AppointmentItem
                oAppointment.Save()
            Case Else
                MsgBox("Not an appointment")
                Exit Sub
        End Select
        'Globals.ThisAddIn.Application.ActiveInspector().CurrentItem.Close(Outlook.OlInspectorClose.olSave)
        oAppointment = Nothing
        olNS = Nothing
    End Sub
    Public Shared Sub DeclineFree()

        Dim objItem As Object
        Dim strID As String
        Dim olNS As Outlook.NameSpace
        Dim oMeetingItem As Outlook.MeetingItem
        Dim oResponse As Outlook.AppointmentItem
        Dim oAppointment As Outlook.AppointmentItem
        Dim oAppointmentCopy As Outlook.AppointmentItem

        Check_Category("Declined")

        objItem = GetCurrentItem()
        strID = objItem.EntryID

        olNS = Globals.ThisAddIn.Application.GetNamespace("MAPI")

        Select Case True
            Case TypeOf objItem Is Outlook.MeetingItem
                oMeetingItem = olNS.GetItemFromID(strID)
                oAppointment = oMeetingItem.GetAssociatedAppointment(True)
            Case TypeOf objItem Is Outlook.AppointmentItem
                oAppointment = objItem
            Case Else
                MsgBox("Not an appointment")
                Exit Sub
        End Select

        oAppointmentCopy = oAppointment.Copy()
        oAppointmentCopy.BusyStatus = Outlook.OlBusyStatus.olFree
        oAppointmentCopy.Subject = "DECLINED: " & oAppointment.Subject
        oAppointmentCopy.Save()

        oResponse = oAppointment.Respond(Outlook.OlMeetingResponse.olMeetingDeclined, False, True)
        'oResponse.Send()

        Select Case True
            Case TypeOf objItem Is Outlook.MeetingItem
                oAppointment.Save()
                oMeetingItem.Save()
                oMeetingItem = Nothing
            Case TypeOf objItem Is Outlook.AppointmentItem
                oAppointment.Save()
            Case Else
                MsgBox("Not an appointment")
                Exit Sub
        End Select
        'Globals.ThisAddIn.Application.ActiveInspector().CurrentItem.Close(Outlook.OlInspectorClose.olSave)
        oAppointment = Nothing
        oAppointmentCopy = Nothing
        olNS = Nothing

    End Sub
    Private Shared Function GetCurrentItem() As Object
        Dim objApp As Outlook.Application
        objApp = Globals.ThisAddIn.Application
        Dim CurrentItem As Object
        On Error Resume Next
        Select Case TypeName(objApp.ActiveWindow)
            Case "Explorer"
                CurrentItem = objApp.ActiveExplorer.Selection.Item(1)
            Case "Inspector"
                CurrentItem = objApp.ActiveInspector.CurrentItem
            Case Else
        End Select
        Return CurrentItem
    End Function
    Private Shared Sub Check_Category(strCategory As String)
        ' Check if category 'declined' exists. If not, create it and add a suitable color
        Dim objNameSpace
        Dim objCategory
        Dim itexists As Boolean = False
        objNameSpace = Globals.ThisAddIn.Application.GetNamespace("MAPI")

        objCategory = objNameSpace.Categories.Item(strCategory)
        If IsNothing(objCategory) Then
            objNameSpace.Categories.Add(strCategory, Outlook.OlCategoryColor.olCategoryColorTeal)
        End If
        objCategory = Nothing
        objNameSpace = Nothing
    End Sub
End Class
