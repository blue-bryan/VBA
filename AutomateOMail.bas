Attribute VB_Name = "AutomateOmail"
Public Sub CreateEmails(strSubject As String, _
                        strRecipients As String, _
                        strBody1 As String, _
                        Optional strBody2 As String, _
                        Optional strBody3 As String, _
                        Optional strCc As String, _
                        Optional strBcc As String, _
                        Optional strAttachment As String, _
                        Optional blnSend As Boolean)
'
' Subroutine to automate creating Outlook emails with user inputs
'

'
    Const olMailItem As Integer = 0, _
            olTo As Integer = 1, _
            olCC As Integer = 2, _
            olBCC As Integer = 3
    
    Dim objOutlook As Object
    Dim objOutlookMsg As Object
    Dim objOutlookInspector As Object
    Dim objOutlookTo As Object
    Dim objOutlookCc As Object
    Dim objOutlookBcc As Object
    Dim objOutlookAttach As Object
    
    Dim varRecipientsList As Variant, i As Integer
    Dim strGreeting As String
    
    On Error GoTo ErrorHandler
    
        Set objOutlook = CreateObject("Outlook.Application")
        Set objOutlookMsg = objOutlook.CreateItem(olMailItem)
        Set objOutlookIns = objOutlookMsg.GetInspector
        
        With objOutlookMsg
            varRecipientsList = Split(strRecipients, ";")
            For i = 0 To UBound(varRecipientsList)
                Set objOutlookTo = .Recipients.Add(Trim(varRecipientsList(i)))
                If Not objOutlookTo Is Nothing Then objOutlookTo.Type = olTo
                objOutlookTo.Resolve
            Next i
            RecipientsList = Split(strCc, ";")
            For i = 0 To UBound(RecipientsList)
                Set objOutlookCc = .Recipients.Add(Trim(varRecipientsList(i)))
                If Not objOutlookCc Is Nothing Then objOutlookCc.Type = olCC
                objOutlookCc.Resolve
            Next i
            varRecipientsList = Split(strBcc, ";")
            For i = 0 To UBound(varRecipientsList)
                Set objOutlookBcc = .Recipients.Add(Trim(varRecipientsList(i)))
                If Not objOutlookBcc Is Nothing Then objOutlookBcc.Type = olBCC
                objOutlookBcc.Resolve
            Next i
            
            .Subject = strSubject
            Select Case Time
            Case Is < TimeValue("12:00")
                strGreeting = "Good morning &mdash;"
            Case Is < TimeValue("16:00")
                strGreeting = "Good afternoon &mdash;"
            Case Else
                strGreeting = "Good evening &mdash;"
            End Select
            If Len(strBody2) > 0 Then
                .HTMLBody = strGreeting & _
                        "<br>" & _
                        strBody1 & _
                        "<br>" & _
                        strBody2 & _
                        "<br>" & _
                    .HTMLBody '<--- retain default signature
            ElseIf Len(strBody3) > 0 Then
                .HTMLBody = strGreeting & _
                        "<br>" & _
                        strBody1 & _
                        "<br>" & _
                        strBody2 & _
                        "<br>" & _
                        strBody3 & _
                        "<br>" & _
                    .HTMLBody '<--- retain default signature
            Else
                .HTMLBody = strGreeting & _
                        "<br>" & _
                        strBody1 & _
                        "<br>" & _
                    .HTMLBody '<--- retain default signature
            End If
            If strAttachment > "" Then
                Set objOutlookAttach = .Attachments.Add(objAttachment)
            End If
            If Not objOutlookTo.Resolve Then
                .Display
                ElseIf Not objOutlookCc Is Nothing Then
                    If Not objOutlookCc.Resolve Then .Display
                ElseIf Not objOutlookBcc Is Nothing Then
                    If Not objOutlookBcc.Resolve Then .Display
            End If
            If blnSend Then
                .Send
            Else
                .Display
            End If
        End With
        
        Set varRecipientsList = Nothing
        Set objOutlookInspector = Nothing
        Set objOutlookMsg = Nothing
        Set objOutlookAttach = Nothing
        Set objOutlookBcc = Nothing
        Set objOutlookCc = Nothing
        Set objOutlookTo = Nothing
        Set objOutlook = Nothing
        
        Exit Sub

ErrorHandler:
    If Err.Number = "287" Then
        MsgBox "Error: The Outlook security warning was declined. " & _
            vbNewLine & "Rerun the procedure and click Yes when " & _
            "prompted for the Outlook security warning." & _
            "For more information, see " & _
            vbNewLine & "http://www.microsoft.com/office/previous/" & _
            "outlook/downloads/security.asp"
    Else
        MsgBox Err.Number, Err.Description
        Resume Next
    End If
End Sub
