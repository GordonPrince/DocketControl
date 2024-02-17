Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports System.IO
Imports System.Net
Imports System.Net.Mail

Friend Class Form1
    Inherits System.Windows.Forms.Form

    Const strCourierOn As String = "<font face=""Courier New"" color=""blue"">"
    Const strInfo As String = "For further information about this notice<BR>" & "contact Melanie Simmonds or<BR>" &
                                "Gordon Prince * gordon@tekhelps.com * (901) 761-3393."

    ' these are used for sending emails
    Const strDocketControlEmail As String = "DocketControl@EvansPetree.com"
    Const strAdminEmail As String = "DockClerk@EvansPetree.com"
    Const strDocketIPEmail As String = "DocketIP@EvansPetree.com"
    Const strGordonPrince As String = "gprince@evanspetree.com"
    Const strHTMLspace As String = "&nbsp;"
    '12/3/2015 changed from "EPExchange"
    '2/21/2018 Const strHost As String = "EXCH2013" 
    Const strHost As String = "smtp.office365.com"
    Dim isDev As Boolean
    Dim strHTML As String
    Dim strScratch As String
    Dim cnn As New ADODB.Connection(), strConnection As String
    Dim dbServer As String = "EPFILE16"
    Dim rst As ADODB.Recordset
    Dim strDeadlines As String = vbNullString

    Private Sub Form1_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Try
            If Environ("UserDomain").StartsWith("TEKHELPS") Then
                isDev = True
                dbServer = "Tekhelps17\SQL2016B"
            End If
            strConnection = "Provider=MSDataShape.1;Persist Security Info=False;Data Source=" & dbServer & ";Integrated Security=SSPI;Initial Catalog=DocketControl;Data Provider=SQLOLEDB.1"
            With Me
                .Text = .Text & " -- " & Application.ProductVersion
                If isDev Then
                    .chkSendToGordon.Checked = True
                    .chkDontUpdateDatabase.Checked = True
                    .chkShowMessages.Checked = True
                ElseIf UCase(VB.Command()) = "/DONTSENDMAIL" Then
                    .chkDontSendMail.Checked = True
                ElseIf UCase(VB.Command()) = "/SHOWMESSAGES" Then
                    .chkShowMessages.Checked = True
                End If
                .txtCalendar.Text = CStr(Today)
                cmdSetDates_Click(cmdSetDates, New System.EventArgs())

                If Not .chkShowMessages.Checked Then
                    SendBothAndQuit()
                End If
            End With
        Catch ex As Exception
            If chkShowMessages.CheckState Then MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Form_Load")
        End Try
    End Sub

    Private Sub SendBothAndQuit()
        cmdDueDate_Click(cmdDueDate, New System.EventArgs())
        cmdEmailNotices_Click(cmdEmailNotices, New System.EventArgs())
        Close()
    End Sub

    Private Sub cmdEmailTest_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdEmailTest.Click
        Cursor = Cursors.WaitCursor
        Dim strTo As String = "AllieRaines@evanspetree.com"
        ' Build HTML for message body.
        strHTML = "<HTML><BODY><P>" & "This is the test HTML message body from the DocketControl Emailer.exe.</P>" &
                  "<P>Please discard it.</P></BODY></HTML>"

        Using Email As New MailMessage
            Try
                ' Apply the settings to the message.
                With Email
                    .From = New MailAddress(strDocketControlEmail, "DocketControl")

                    strTo = InputBox("Send test email to:", cmdEmailTest.Text, strTo)
                    If Not strTo.Contains("@") Then Exit Sub

                    .To.Add(New MailAddress(strTo))
                    .Subject = "Docket Control test message"
                    .IsBodyHtml = True
                    .Body = strHTML
                End With

                If SendEmail(Email) Then
                    MsgBox("Mail sent to " & strTo & ".", MsgBoxStyle.Information, cmdEmailTest.Text)
                End If

            Catch ex As SmtpFailedRecipientsException
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, cmdEmailTest.Text)
            Finally
                Cursor = Cursors.Default
            End Try
        End Using
    End Sub

    Private Function SendEmail(eM As MailMessage) As Boolean
        Dim attempt As Integer = 1
        Dim client As New SmtpClient

        Try
            With client
                .UseDefaultCredentials = False
                .EnableSsl = True
                'If chkUseGmail.Checked Then
                '    .Host = "smtp.gmail.com"
                '    .Credentials = New NetworkCredential("ridgeway17gordon@gmail.com", "OJLyWoqqP##UWJH")
                '    .Port = 587
                'Else
                .Host = strHost
                    .Port = 587
                    .Credentials = New NetworkCredential(strDocketControlEmail, "friday15")
                    .DeliveryMethod = SmtpDeliveryMethod.Network '7/19/2022 added this
                '.Credentials = CredentialCache.DefaultNetworkCredentials
                '.UseDefaultCredentials = True
                'End If
            End With

SendEmail:
            txtStatus.Text = "attempt=" & attempt.ToString
            client.Send(eM)
            txtStatus.Text = "Sent on attempt " & attempt & "."
            Return True

        Catch ex As SmtpFailedRecipientsException
            If attempt < 9 Then
                If chkShowMessages.Checked Then
                    If MsgBox("attempt=" & attempt & "." & vbNewLine & vbNewLine & ex.InnerException.ToString & vbNewLine & vbNewLine & ex.Message.ToString & vbNewLine & vbNewLine & "Try again?", vbQuestion + vbYesNo, cmdEmailTest.Text) = vbNo Then Return False
                Else
                    System.Threading.Thread.Sleep(500)
                End If
                attempt += 1
                GoTo SendEmail
            End If
        End Try
    End Function

    Private Sub cmdSetDates_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSetDates.Click
        ' Notices are sent based on their actual date, DueDate emails are sent the business day prior to the DueDate
        Dim datUse As Date, datBus As Date

        ' datUse is the day prior to the next business day (on Friday get notices through Sunday, unless Monday's a holiday then get Notices through Monday, etc.)
        ' if datUse = Friday then add days to get to Sunday
        ' if datUse + 1 is a holiday, add days until it's not a holiday
        ' repeat after each change to make sure the new date is not a Friday or a holiday

        datUse = CDate(txtCalendar.Text)
        datBus = datUse
        Do Until datUse = DateAdd(DateInterval.Day, -1, datBus)
            Select Case Weekday(datUse)
                Case FirstDayOfWeek.Friday
                    datUse = DateAdd(Microsoft.VisualBasic.DateInterval.Day, 2, datUse)
                Case FirstDayOfWeek.Saturday
                    datUse = DateAdd(Microsoft.VisualBasic.DateInterval.Day, 1, datUse)
            End Select
            ' if the next day will be a holiday use that day and check the next day
            datBus = DateAdd(DateInterval.Day, 1, datUse)
            If IsHoliday(datBus) Then datUse = datBus
        Loop
        txtNotice.Text = CStr(datUse)
        txtDueDate.Text = CStr(datBus)
    End Sub

    Private Sub cmdDueDate_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDueDate.Click
        Const strTitle As String = "E-mail from DueDate"
        Dim datCriteria As Date
        Dim strSQL As String ', strNotify As String
        Dim rstNotify As New ADODB.Recordset
        Dim strTo1 As String, strTo As String
        Dim intCounter As Short
        Dim rstMark As New ADODB.Recordset
        Dim strSubject As String

        If Not Date.TryParse(txtDueDate.Text, datCriteria) Then
            MsgBox("Invalid DueDate.", vbExclamation + vbOKOnly, strTitle)
            Exit Sub
        End If

        Try
            cnn.Open(strConnection)
            rst = New ADODB.Recordset
            ' 10/17/2002 is when the system went into operation
            ' 4/24/2014 only send notices for Docket items that came from IPMark rows
            'strSQL = "SELECT * FROM Docket WHERE Canceled = 0 AND DueDate BETWEEN '10/23/2021' AND '" & datCriteria & "' AND (DueDateEmailed IS NULL OR DueDateEmailed > '2/25/2022') " &
            strSQL = "SELECT * FROM Docket WHERE Canceled = 0 AND DueDate BETWEEN '10/23/2021' AND '" & datCriteria & "' AND DueDateEmailed IS NULL " &
                     "AND (Trademark = 0 OR (Trademark = 1 AND MarkID > 0)) ORDER BY DueDate, DocketID"
            With rst
                .Open(strSQL, cnn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
                Do Until .EOF
                    'added 1/11/2015 for special IP notice rules
                    If SkipNotice(.Fields("MarkID").Value) Then GoTo NextNotice

                    Using Email As New MailMessage
                        strSubject = "Docket Control Item Due " & CStr(.Fields("DueDate").Value) & " (ID: " & CStr(.Fields("DocketID").Value) & ")"
                        If IsDBNull(.Fields("MarkID").Value) Then
                        ElseIf .Fields("MarkID").Value > 0 Then
                            strScratch = "select ResponsibleAtty from IPmark where MarkID = " & .Fields("MarkID").Value
                            rstMark.Open(strScratch, cnn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                            If Not rstMark.EOF Then
                                If IsDBNull(rstMark.Fields("ResponsibleAtty").Value) Then
                                    strSubject = strSubject & " ResponsibleAtty: UNDEFINED"
                                Else
                                    If Len(rstMark.Fields("ResponsibleAtty").Value) > 0 Then
                                        strSubject = strSubject & " ResponsibleAtty: " & rstMark.Fields("ResponsibleAtty").Value
                                    End If
                                End If
                            End If
                            rstMark.Close()
                        End If
                        Email.Subject = strSubject

                        With Email
                            .From = New MailAddress(strDocketControlEmail, "DocketControl")
                            .IsBodyHtml = True
                        End With

                        'Email notices to the appropriate parties
                        If chkSendToGordon.Checked Then
                            strTo = strGordonPrince
                            Email.To.Add(strGordonPrince)
                        Else
                            If .Fields("Trademark").Value = 0 Then
                                rstNotify.Open("select Email from v_NotifyEmail where Email <> 'Admin' and DocketID = " & .Fields("DocketID").Value & " ORDER BY Email", cnn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                                With rstNotify
                                    If .EOF Then
                                        strHTML = "<HTML><BODY><font color=""red""<P>" &
                                                "<strong>NO ONE WAS NOTIFIED OF THE FOLLOWING DOCKET CONTROL ITEM.</strong></font>" & " DocketID=" & rst.Fields("DocketID").Value & "</P>"
                                        strTo = strAdminEmail
                                        Email.To.Add(New MailAddress(strTo))
                                    Else
                                        strHTML = "<HTML><BODY>"
                                        strTo1 = .Fields("Email").Value & "@evanspetree.com"
                                        Email.To.Add(New MailAddress(strTo1))
                                        strTo = strTo1
                                        .MoveNext()
                                        Do Until .EOF
                                            strTo1 = .Fields("Email").Value & "@evanspetree.com"
                                            'If chkUseGmail.CheckState = CheckState.Unchecked Then
                                            Email.To.Add(New MailAddress(strTo1))
                                            'End If
                                            strTo &= ", " & strTo1
                                            .MoveNext()
                                        Loop
                                    End If
                                    .Close()
                                End With
                            Else
                                strHTML = "<HTML><BODY>"
                                strTo = strDocketIPEmail
                                Email.To.Add(New MailAddress(strTo))
                            End If
                            ' 1/9/2013 added this (strTo is used in the body of the email below)
                            strTo = Replace(strTo, "@evanspetree.com", "")
                        End If

                        ' 1/16/2013 added this
                        If .Fields("Completed").Value <> 0 Then
                            strHTML = strHTML & "<font color=""red""><strong>This item was " & .Fields("CompletedBy").Value & "</strong></font></P><P>"
                            'Debug.Print(strHTML)
                        End If

                        Email.To.Add(New MailAddress(strAdminEmail))

                        If chkDontSendMail.CheckState = CheckState.Unchecked Then
                            Email.Body = strHTML & HTMLbody(rst, strTo) & "</P></BODY></HTML>"
                            If SendEmail(Email) AndAlso chkShowMessages.CheckState Then
                                MsgBox("Mail sent to " & strTo & ".", MsgBoxStyle.Information, cmdEmailTest.Text)
                            End If
                        End If
                    End Using

                    ' update the database that the email was sent
                    If chkDontUpdateDatabase.CheckState = CheckState.Checked Then
                        .CancelUpdate()
                    Else
                        .Fields("DueDateEmailed").Value = Now
                        .Update()
                    End If
                    intCounter += 1
                    If chkShowMessages.CheckState Then
                        strScratch = "Email sent to: " & strTo & vbNewLine &
                                    .Fields("Event").Value & vbNewLine &
                                    .Fields("MatterID").Value & vbNewLine &
                                    "DueDate = " & .Fields("DueDate").Value & vbNewLine &
                                    "Trademark = " & rst.Fields("Trademark").Value & vbNewLine & vbNewLine &
                                    "Process the next item?"
                        If MsgBox(strScratch, MsgBoxStyle.YesNo + MsgBoxStyle.Question, strTitle) = MsgBoxResult.No Then GoTo FinishedLoop
                    End If
NextNotice:
                    .MoveNext()
                Loop
FinishedLoop:
                .Close()
            End With

            If Not chkDontSendMail.Checked Then
                ' notify the administrator what was done
                '12/3/2015 changed this so only one email goes out daily
                strDeadlines = "<P>" & IIf(intCounter = 1, "One deadline E-mail was", intCounter & " deadline E-mails were") &
                                " sent for items with DueDates on or before " & datCriteria & ".</P>"
            End If
            If chkShowMessages.CheckState = CheckState.Checked Then MsgBox("Finished sending " & intCounter & " Email(s)", MsgBoxStyle.Information, strTitle)
        Catch ex As Exception
            MsgBox(ex.Message, vbExclamation, strTitle)
        Finally
            rstNotify = Nothing
            If rst.State <> 0 Then rst.Close()
            rst = Nothing
            cnn.Close()
        End Try

    End Sub

    Private Sub cmdEmailNotices_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdEmailNotices.Click
        Const strTitle As String = "E-mail Notices"
        Dim datCriteria As Date
        Dim strSQL As String
        Dim rstMark As New ADODB.Recordset, rstNotify As New ADODB.Recordset
        Dim bMark As Boolean, strFolder As String, strFile As String
        Dim strTo1 As String, strTo As String = "No one"
        Dim intCounter As Short
        Dim objStreamWriter As StreamWriter

        If isDev Then
            strFolder = "C:\tmp"
        Else
            strFolder = "\\EPFile16\Progs\EP Docket"
        End If
        strFolder &= "\ShowIP\"

        Try
            If Not Date.TryParse(txtNotice.Text, datCriteria) Then
                MsgBox("Notice E-mail date is not a valid date.", vbInformation + vbOKOnly, strTitle)
                Exit Sub
            End If
            ' even if there are no items, an email will be sent to the strAdmin to that effect. So SMTP needs to be initialized
            '12/22/2021
            Dim SMTP As New SmtpClient
            With SMTP
                .UseDefaultCredentials = False
                .Host = strHost
                .Credentials = New NetworkCredential(strDocketControlEmail, "friday15")
                .Port = 587
                .EnableSsl = True
            End With

            cnn.Open(strConnection)
            rst = New ADODB.Recordset
            strSQL = "select * from Docket WHERE DueDate > '10/23/2021' AND DueDateEmailed IS NULL AND Completed = 0 AND Canceled = 0 " &
                                "AND (Trademark = 0 OR (Trademark = 1 AND MarkID > 0)) " &
                                "AND ((NoticeFinal <= '" & txtNotice.Text & "' AND NoticeFinalEmailed IS NULL)" &
                                            " OR (TmNotice7 <= '" & txtNotice.Text & "' AND TmNotice7Emailed IS NULL)" &
                                            " OR (TmNotice30 <= '" & txtNotice.Text & "' AND TmNotice30Emailed IS NULL)" &
                                            " OR (Notice2 <= '" & txtNotice.Text & "' AND Notice2Emailed IS NULL)" &
                                            " OR (Notice1 <= '" & txtNotice.Text & "' AND Notice1Emailed IS NULL))" &
                            " ORDER BY DueDate, DocketID"
            'Console.WriteLine(strSQL)
            With rst
                .Open(strSQL, cnn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
                Do Until .EOF
                    'added 1/11/2015 for special IP notice rules
                    If SkipNotice(.Fields("MarkID").Value) Then GoTo NextNotice
                    Using Email As New MailMessage
                        With Email
                            .From = New MailAddress(strDocketControlEmail, "DocketControl")
                            .IsBodyHtml = True
                            If Not IsDBNull(rst.Fields("NoticeFinal").Value) Then
                                If rst.Fields("NoticeFinal").Value <= datCriteria Then
                                    If rst.Fields("Trademark").Value Then
                                        .Subject = NoticeLabel(DateDiff(DateInterval.Day, rst.Fields("NoticeFinal").Value, rst.Fields("DueDate").Value), False)
                                    Else
                                        .Subject = "Final Notice"
                                    End If
                                    GoTo HaveSubject
                                End If
                            End If
                            If Not IsDBNull(rst.Fields("TmNotice7").Value) Then
                                If rst.Fields("TmNotice7").Value <= datCriteria Then
                                    .Subject = NoticeLabel(DateDiff(DateInterval.Day, rst.Fields("TmNotice7").Value, rst.Fields("DueDate").Value), False)
                                    GoTo HaveSubject
                                End If
                            End If
                            If Not IsDBNull(rst.Fields("TmNotice30").Value) Then
                                If rst.Fields("TmNotice30").Value <= datCriteria Then
                                    .Subject = NoticeLabel(DateDiff(DateInterval.Day, rst.Fields("TmNotice30").Value, rst.Fields("DueDate").Value), False)
                                    GoTo HaveSubject
                                End If
                            End If
                            If Not IsDBNull(rst.Fields("Notice2").Value) Then
                                If rst.Fields("Notice2").Value <= datCriteria Then
                                    If rst.Fields("Trademark").Value Then
                                        .Subject = NoticeLabel(DateDiff(DateInterval.Day, rst.Fields("Notice2").Value, rst.Fields("DueDate").Value), False)
                                    Else
                                        .Subject = "Second Notice"
                                    End If
                                    GoTo HaveSubject
                                End If
                            End If
                            If Not IsDBNull(rst.Fields("Notice1").Value) Then
                                If rst.Fields("Notice1").Value <= datCriteria Then
                                    If rst.Fields("Trademark").Value Then
                                        .Subject = NoticeLabel(DateDiff(DateInterval.Day, rst.Fields("Notice1").Value, rst.Fields("DueDate").Value), False)
                                    Else
                                        .Subject = "First Notice"
                                    End If
                                    GoTo HaveSubject
                                End If
                            Else
                                .Subject = "Docket Control Notification"
                            End If
                        End With
HaveSubject:
                        Email.Subject = Email.Subject & " (ID: " & CStr(.Fields("DocketID").Value) & ")"
                        If .Fields("MarkID").Value > 0 Then
                            strScratch = "select * from IPmark where MarkID = " & .Fields("MarkID").Value
                            rstMark.Open(strScratch, cnn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                            If rstMark.EOF Then
                                bMark = False
                                rstMark.Close()
                                Email.Subject &= " ResponsibleAtty: UNKNOWN"
                            Else
                                bMark = True
                                Email.Subject &= " ResponsibleAtty: " & rstMark.Fields("ResponsibleAtty").Value
                            End If
                        End If

                        ' send email notices to the appropriate parties
                        If chkSendToGordon.Checked Then
                            strTo = strGordonPrince
                            Email.To.Add(strGordonPrince)
                        Else
                            rstNotify.Open("select Email from v_NotifyEmail where Email <> 'Admin' and DocketID = " & .Fields("DocketID").Value & " ORDER BY Email", cnn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                            With rstNotify
                                If .EOF Then
                                    strHTML = "<HTML><BODY><P><font color=""red""" &
                                              "<strong>NO ONE WAS NOTIFIED OF THE FOLLOWING DOCKET CONTROL ITEM.</strong></font>" & " DocketID=" & rst.Fields("DocketID").Value & "</P>"
                                    strTo = strAdminEmail
                                    Email.To.Add(New MailAddress(strTo))
                                Else
                                    strHTML = "<HTML><BODY>"
                                    strTo1 = .Fields("Email").Value & "@evanspetree.com"
                                    Email.To.Add(New MailAddress(strTo1))
                                    strTo = strTo1
                                    .MoveNext()
                                    Do Until .EOF
                                        strTo1 = .Fields("Email").Value & "@evanspetree.com"
                                        Email.To.Add(New MailAddress(strTo1))
                                        strTo &= ", " & strTo1
                                        .MoveNext()
                                    Loop
                                End If
                                .Close()
                            End With
                            strTo = Replace(strTo, "@evanspetree.com", "")
                        End If

                        If bMark = True Then
                            strFile = strFolder & "Mark" & CStr(rstMark.Fields("MarkID").Value) & ".bat"
                            'Pass the file path and the file name to the StreamWriter constructor.
                            objStreamWriter = New StreamWriter(strFile)
                            'Write a line of text.
                            If isDev Then
                                strScratch = "start ""C:\Program Files (x86)\Microsoft Office\Office14\MSACCESS.EXE"" ""C:\Access\Access2010\DocketControl\EPdocket2010.adp"" /cmd " & CStr(rstMark.Fields("MarkID").Value)
                            Else
                                strScratch = "start ""C:\Program Files (x86)\Microsoft Office\Office14\MSACCESS.EXE"" ""C:\Tekhelps\EPdocket.ade"" /cmd " & CStr(rstMark.Fields("MarkID").Value)
                            End If
                            objStreamWriter.WriteLine(strScratch)
                            'Close the file.
                            objStreamWriter.Close()

                            strScratch = Replace(strCourierOn, "blue", "red") & "<strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Deadline: </strong></font>" & CStr(.Fields("DueDate").Value) & " -- " & .Fields("Memo").Value & "<P>"
                            If Not IsDBNull(rstMark.Fields("Trademark").Value) Then strScratch = strScratch & strCourierOn & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Trademark: </font><strong>" & rstMark.Fields("Trademark").Value & "</strong><BR>"
                            If Not IsDBNull(rstMark.Fields("SerialNo").Value) Then strScratch = strScratch & strCourierOn & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Serial No: </font>" & rstMark.Fields("SerialNo").Value & "<BR>"
                            If Not IsDBNull(rstMark.Fields("RegistrationNo").Value) Then strScratch = strScratch & strCourierOn & "RegistrationNo: </font>" & rstMark.Fields("RegistrationNo").Value & "<BR>"
                            If Not IsDBNull(rstMark.Fields("Jurisdiction").Value) Then strScratch = strScratch & strCourierOn & "&nbsp;&nbsp;Jurisdiction: </font>" & rstMark.Fields("Jurisdiction").Value & "<BR>"
                            If Not IsDBNull(rstMark.Fields("ApplicantName").Value) Then strScratch = strScratch & strCourierOn & "&nbsp;ApplicantName: </font>" & rstMark.Fields("ApplicantName").Value & "<BR>"
                            If Not IsDBNull(rstMark.Fields("GoodsServices").Value) Then strScratch = strScratch & strCourierOn & "Goods/Services: </font>" & rstMark.Fields("GoodsServices").Value
                            strScratch = strScratch & "<P>" & strCourierOn & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Mark ID: </font>" & CStr(rstMark.Fields("MarkID").Value) &
                                                " <a href=""file://" & strFile & """>Click here to open Mark in IP Dashboard.</a>"
                            strHTML = strHTML & strScratch & "</P><P><font color=""blue"" size=""2""><I>The old format of this information is at the bottom of the page.</I></font><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR>"
                            rstMark.Close()
                            bMark = False
                        End If

                        If Not chkSendToGordon.Checked AndAlso rst.Fields("Trademark").Value <> 0 Then
                            Email.To.Add(New MailAddress(strDocketIPEmail))
                        End If

                        If chkDontSendMail.CheckState = CheckState.Unchecked Then
                            Email.Body = strHTML & HTMLbody(rst, strTo) & "</P></BODY></HTML>"
                            If SendEmail(Email) AndAlso chkShowMessages.CheckState Then
                                MsgBox("Mail sent to " & strTo & ".", MsgBoxStyle.Information, cmdEmailTest.Text)
                            End If
                        End If
                    End Using

                    If chkDontUpdateDatabase.CheckState = CheckState.Checked Then
                        .CancelUpdate()
                    Else
                        ' update the database that the email was sent
                        ' if >1 of the dates = datCriteria then >1 of the Emailed columns needs to be updated
                        If Not IsDBNull(rst.Fields("NoticeFinal").Value) Then
                            If .Fields("NoticeFinal").Value <= datCriteria And IsDBNull(.Fields("NoticeFinalEmailed").Value) Then .Fields("NoticeFinalEmailed").Value = Now
                        End If
                        If Not IsDBNull(rst.Fields("TmNotice7").Value) Then
                            If .Fields("TmNotice7").Value <= datCriteria And IsDBNull(.Fields("TmNotice7Emailed").Value) Then .Fields("TmNotice7Emailed").Value = Now
                        End If
                        If Not IsDBNull(rst.Fields("TmNotice30").Value) Then
                            If .Fields("TmNotice30").Value <= datCriteria And IsDBNull(.Fields("TmNotice30Emailed").Value) Then .Fields("TmNotice30Emailed").Value = Now
                        End If
                        If Not IsDBNull(rst.Fields("Notice2").Value) Then
                            If .Fields("Notice2").Value <= datCriteria And IsDBNull(.Fields("Notice2Emailed").Value) Then .Fields("Notice2Emailed").Value = Now
                        End If
                        If Not IsDBNull(rst.Fields("Notice1").Value) Then
                            If .Fields("Notice1").Value <= datCriteria And IsDBNull(.Fields("Notice1Emailed").Value) Then .Fields("Notice1Emailed").Value = Now
                        End If
                        .Update()
                    End If
                    intCounter += 1
                    If chkShowMessages.Checked Then
                        strScratch = "Email(s) sent to: " & strTo & vbNewLine &
                                    .Fields("Event").Value & vbNewLine &
                                    .Fields("MatterID").Value & vbNewLine &
                                    "DueDate = " & .Fields("DueDate").Value & vbNewLine &
                                    "Trademark = " & rst.Fields("Trademark").Value & vbNewLine & vbNewLine &
                                    "Process the next item?"
                        If MsgBox(strScratch, MsgBoxStyle.YesNo + MsgBoxStyle.Question, strTitle) = MsgBoxResult.No Then GoTo FinishedLoop
                    End If
NextNotice:
                    .MoveNext()
                Loop
FinishedLoop:
                .Close()
            End With

            If Not chkDontSendMail.Checked Then
                'notify the administrator what was done
                Using Email As New MailMessage
                    With Email
                        .IsBodyHtml = True
                        .From = New MailAddress(strDocketControlEmail, "DocketControl")
                        .Subject = "Docket Control E-mail Summary"
                        .To.Add(New MailAddress("DocketSummary@evanspetree.com"))
                        If Len(strDeadlines) > 0 Then
                            .Body = strDeadlines
                        Else
                            .Body = "<P><font color=""red""<strong>* * * THE DEADLINES DID NOT PROCESS PROPERLY * * *</strong></font></P>"
                        End If
                        .Body = .Body & IIf(intCounter = 1, "One reminder E-mail was", intCounter & " reminder E-mails were") &
                                " sent for items with Notice dates through " & txtNotice.Text & ".</P>"
                    End With
                    'wait 2 seconds to make sure the summary email is the last one sent
                    System.Threading.Thread.Sleep(1000)
                    If SendEmail(Email) AndAlso chkShowMessages.Checked Then
                        MsgBox("Mail (including Summary) sent.", MsgBoxStyle.Information, cmdEmailTest.Text)
                    End If
                End Using
            End If
            If chkShowMessages.Checked Then MsgBox("Finished sending " & intCounter & " Email(s)", MsgBoxStyle.Information, strTitle)
        Catch ex As Exception
            MsgBox(Err.Description, vbExclamation + vbOKOnly, strTitle)
        Finally
            If rst.State <> 0 Then rst.Close()
            rst = Nothing
            rstNotify = Nothing
            cnn.Close()
        End Try
    End Sub

    Private Function SkipNotice(id As Long) As Boolean
        'created 1/11/2016 for IP items, deployed finally 2/14/2016
        Dim strSQL As String = "SELECT NULL FROM IPmark " &
                                "WHERE Suspended IS NULL AND ApplicationAbandoned IS NULL AND MarkID = " & id
        Dim rst As New ADODB.Recordset
        With rst
            .Open(strSQL, cnn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            If .EOF Then
                .Close()
                Return True
            Else
                .Close()
                Return False
            End If
        End With
    End Function

    Private Function IsHoliday(ByVal datB As Date) As Boolean
        Dim rstH As ADODB.Recordset, strSQL As String
        Dim isHol As Boolean
        ' January 1st is always a holiday, often it won't get entered so don't let things go wrong on account of that
        If Month(datB) = 1 And Microsoft.VisualBasic.DateAndTime.Day(datB) = 1 Then
            Return True
        End If
        cnn.Open(strConnection)
        rstH = New ADODB.Recordset
        With rstH
CheckIfHoliday:
            strSQL = "select Holidate from Holiday where Holidate = '" & datB & "'"
            .Open(strSQL, cnn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
            If .EOF Then
                isHol = False
            Else
                isHol = True
            End If
            .Close()
        End With
        rstH = Nothing
        cnn.Close()
        Return isHol
    End Function

    Private Function HTMLbody(ByRef rst As ADODB.Recordset, ByRef strTo As String) As String
        Dim strHTML As String
        Dim date1 As Date
        Try
            With rst
                strHTML = strCourierOn & "MATTER ID:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font>" & .Fields("MatterID").Value & "<BR>" &
                                  strCourierOn & "CLIENT:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font>" & .Fields("ClientName").Value & "<BR>" &
                                  strCourierOn & "MATTER:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font>" & .Fields("MatterName").Value & "</P><P>"
                If .Fields("Trademark").Value Then
                    If Not IsDBNull(.Fields("MarkID").Value) Then strHTML = strHTML & strCourierOn & "MARK ID:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font>" & .Fields("MarkID").Value & "<BR>"
                End If
                If IsDBNull(.Fields("Event").Value) Then
                    strScratch = vbNullString
                Else
                    strScratch = Replace(.Fields("Event").Value, vbNewLine, "<BR>")
                End If
                If .Fields("Trademark").Value Then
                    strHTML = strHTML & strCourierOn & "COUNTRY:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font>" & .Fields("Court").Value & "<BR>" &
                                                         Replace(strCourierOn, "blue", "red") & "TRADEMARK:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font>" & strScratch & "<BR>" &
                                                         strCourierOn & "TM NUMBER:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font>" & .Fields("DocketNo").Value & "<BR>" &
                                                         strCourierOn & "CLASS:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font>" & .Fields("TmClass").Value & "<BR>"
                Else
                    strHTML = strHTML & Replace(strCourierOn, "blue", "red") & "EVENT:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font>" & strScratch & "<BR>"
                End If

                strHTML = strHTML & Replace(strCourierOn, "blue", "red") & "DUE DATE:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font><font color=""red""><strong>" & CDate(.Fields("DueDate").Value).ToString("MMM d, yyyy (ddd)") & "</strong></font><BR>" &
                                                     strCourierOn & "NOTIFY:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font>" & strTo & "<BR>"
                If Not IsDBNull(.Fields("Memo").Value) Then strHTML = strHTML & strCourierOn & "MEMO:</font><BR>" & Replace(.Fields("Memo").Value, vbNewLine, "<BR>")
                If .Fields("Trademark").Value Then
                    If Not IsDBNull(.Fields("Notice1").Value) Or
                        Not IsDBNull(.Fields("Notice2").Value) Or
                        Not IsDBNull(.Fields("TmNotice30").Value) Or
                        Not IsDBNull(.Fields("TmNotice7").Value) Or
                        Not IsDBNull(.Fields("NoticeFinal").Value) Then
                        strHTML = strHTML & "<HR>"
                    End If
                    If Not IsDBNull(.Fields("Notice1").Value) Then strHTML = strHTML & strCourierOn & NoticeLabel(DateDiff(DateInterval.Day, .Fields("Notice1").Value, .Fields("DueDate").Value), True) & "</font>" & CDate(.Fields("Notice1").Value).ToString("MMM d, yyyy") & "<BR>"
                    If Not IsDBNull(.Fields("Notice2").Value) Then strHTML = strHTML & strCourierOn & NoticeLabel(DateDiff(DateInterval.Day, .Fields("Notice2").Value, .Fields("DueDate").Value), True) & "</font>" & CDate(.Fields("Notice2").Value).ToString("MMM d, yyyy") & "<BR>"
                    If Not IsDBNull(.Fields("TmNotice30").Value) Then strHTML = strHTML & strCourierOn & NoticeLabel(DateDiff(DateInterval.Day, .Fields("TmNotice30").Value, .Fields("DueDate").Value), True) & "</font>" & CDate(.Fields("TmNotice30").Value).ToString("MMM d, yyyy") & "<BR>"
                    If Not IsDBNull(.Fields("TmNotice7").Value) Then strHTML = strHTML & strCourierOn & NoticeLabel(DateDiff(DateInterval.Day, .Fields("TmNotice7").Value, .Fields("DueDate").Value), True) & "</font>" & CDate(.Fields("TmNotice7").Value).ToString("MMM d, yyyy") & "<BR>"
                    If Not IsDBNull(.Fields("NoticeFinal").Value) Then strHTML = strHTML & strCourierOn & NoticeLabel(DateDiff(DateInterval.Day, .Fields("NoticeFinal").Value, .Fields("DueDate").Value), True) & "</font>" & CDate(.Fields("NoticeFinal").Value).ToString("MMM d, yyyy")
                Else
                    If Not IsDBNull(.Fields("NoticeFinal").Value) Or
                        Not IsDBNull(.Fields("Notice2").Value) Or
                        Not IsDBNull(.Fields("Notice1").Value) Or
                        Not IsDBNull(.Fields("Court").Value) Or
                        Not IsDBNull(.Fields("DocketNo").Value) Then
                        strHTML = strHTML & "<HR>"
                    End If
                    If Not IsDBNull(.Fields("NoticeFinal").Value) Then strHTML = strHTML & strCourierOn & "FINAL&nbsp;&nbsp;NOTICE: </font>" & CDate(.Fields("NoticeFinal").Value).ToString("MMM d, yyyy") & "<BR>"
                    If Not IsDBNull(.Fields("Notice2").Value) Then strHTML = strHTML & strCourierOn & "SECOND NOTICE: </font>" & CDate(.Fields("Notice2").Value).ToString("MMM d, yyyy") & "<BR>"
                    If Not IsDBNull(.Fields("Notice1").Value) Then strHTML = strHTML & strCourierOn & "FIRST&nbsp;&nbsp;NOTICE: </font>" & CDate(.Fields("Notice1").Value).ToString("MMM d, yyyy") & "<BR>"
                    If Not IsDBNull(.Fields("Court").Value) Then strHTML = strHTML & strCourierOn & "COURT:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font>" & .Fields("Court").Value & "<BR>"
                    If Not IsDBNull(.Fields("DocketNo").Value) Then strHTML = strHTML & strCourierOn & "DOCKET NO:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font>" & .Fields("DocketNo").Value
                End If
                strHTML = strHTML & "<HR>"
                If Not IsDBNull(.Fields("Created").Value) Then
                    date1 = .Fields("Created").Value
                    strHTML = strHTML & strCourierOn & "Created: </font>" &
                                      date1.ToString("ddd MMM dd yyyy hh:mm tt") & " by " & .Fields("CreatedBy").Value & "<BR>"
                End If
                If Not IsDBNull(.Fields("Updated").Value) Then strHTML = strHTML & strCourierOn & "Updated: </font>" & Replace(Replace(.Fields("Updated").Value, "Updated", vbNullString), vbNewLine, "<BR>") & "<BR>"
                strHTML = strHTML & "<HR><font color=""blue""><I>" & strInfo & "</I></font></P>"
            End With
            HTMLbody = strHTML
            Exit Function
        Catch ex As Exception
            MsgBox(ex.Message, vbExclamation + vbOKOnly, "HTMLbody")
            HTMLbody = ex.Message.ToString
        End Try
    End Function

    Private Function NoticeLabel(ByVal intDays As Integer, ByVal bHTMLspaces As Boolean) As String
        Dim intS As Integer, strLabel As String
        Dim intX As Integer
        Select Case intDays
            Case 364 To 367
                NoticeLabel = "One Year Notice"
            Case 178 To 185
                NoticeLabel = "Six Month Notice"
            Case 86 To 94
                NoticeLabel = "Three Month Notice"
            Case 57 To 63
                NoticeLabel = "Two Month Notice"
            Case 28 To 32
                NoticeLabel = "One Month Notice"
            Case Else
                NoticeLabel = intDays & " Day Notice"
        End Select
        If bHTMLspaces Then
            strLabel = NoticeLabel
            intS = 18 - Len(strLabel)
            For intX = 1 To intS
                NoticeLabel = strHTMLspace & NoticeLabel
            Next
            NoticeLabel = NoticeLabel & ":" & strHTMLspace
        End If
    End Function

    '    Private Sub SendIP_Click(sender As Object, e As EventArgs) Handles SendIP.Click
    '        'created 1/11/2015 for special IP notice rules
    '        Const strTitle As String = "Email Special IP"
    '        Dim datCriteria As Date, datCalendar As Date
    '        Dim strSQL As String
    '        Dim strFolder As String, strFile As String
    '        Dim SMTP As New SmtpClient
    '        Dim strTo As String = "No one"
    '        Dim intCounter As Short
    '        Dim objStreamWriter As StreamWriter

    '        If isDev Then
    '            strFolder = "D:\temp"
    '        Else
    '            strFolder = "\\EPFile16\PROGS\EP Docket"
    '        End If
    '        strFolder &= "\ShowIP\"

    '        Try
    '            datCalendar = CDate(txtCalendar.Text)
    '            datCriteria = CDate(txtNotice.Text)
    '            ' 4 months in the future
    '            datCriteria = DateAdd(DateInterval.Month, -4, datCriteria)
    '        Catch ex As Exception
    '            MsgBox("Invalid date entered on form.")
    '        End Try

    '        ' even if there are no items, an email will be sent to the strAdmin to that effect. So SMTP needs to be initialized
    '        With SMTP
    '            .UseDefaultCredentials = False
    '            If chkUseGmail.Checked Then
    '                .Host = "smtp.gmail.com"
    '                .Credentials = New NetworkCredential("ridgeway17gordon@gmail.com", "OJLyWoqqP##UWJH")
    '            Else
    '                .Host = strHost
    '                .Credentials = New NetworkCredential(strDocketControlEmail, "friday15")
    '            End If
    '            .Port = 587
    '            .EnableSsl = True
    '        End With

    '        cnn.Open("Provider=MSDataShape.1;Persist Security Info=False;Data Source=" & dbServer & ";Integrated Security=SSPI;Initial Catalog=DocketControl;Data Provider=SQLOLEDB.1")
    '        rst = New ADODB.Recordset
    '        strSQL = "select * from IPmark " &
    '                 "WHERE (Suspended IS NOT NULL) AND (ApplicationAbandoned IS NULL) " &
    '                    "AND (LastEmailSuspended IS NULL OR LastEmailSuspended < '" & datCriteria & "')" &
    '                 " ORDER BY Suspended"
    '        'Debug.Print(strSQL)
    '        With rst
    '            .Open(strSQL, cnn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
    '            Do Until .EOF
    '                Using Email As New MailMessage
    '                    With Email

    '                        If chkUseGmail.Checked Then
    '                            .From = New MailAddress("ridgeway17gordon@gmail.com", "Docket Control 25")
    '                            .Bcc.Add(New MailAddress(strDocketControlEmail))
    '                        Else
    '                            .From = New MailAddress(strDocketControlEmail, "DocketControl")
    '                        End If

    '                        .IsBodyHtml = True
    '                        If IsDBNull(rst.Fields("LastEmailSuspended").Value) Then
    '                            .Subject = "Mark " & rst.Fields("MarkID").Value & " was"
    '                        Else
    '                            .Subject = "Four month reminder Mark " & rst.Fields("MarkID").Value
    '                        End If
    '                        .Subject = .Subject & " Suspended " & rst.Fields("Suspended").Value &
    '                                   " ResponsibleAtty: " & rst.Fields("ResponsibleAtty").Value

    '                        'If chkUseGmail.CheckState = CheckState.Checked Then
    '                        '    strTo = strGordonPrince
    '                        'Else
    '                        strTo = rst.Fields("ResponsibleAtty").Value & "@evanspetree.com"
    '                        'End If
    '                        .To.Add(New MailAddress(strTo))
    '                        ' this is here because I want to confirm it's working in live operation once or twice
    '                        .Bcc.Add(New MailAddress(strGordonPrince))
    '                    End With
    '                    strFile = strFolder & "Mark" & CStr(rst.Fields("MarkID").Value) & ".bat"
    '                    'Pass the file path and the file name to the StreamWriter constructor.
    '                    objStreamWriter = New StreamWriter(strFile)
    '                    'Write a line of text.
    '                    If isDev Then
    '                        strScratch = "start ""C:\Program Files (x86)\Microsoft Office\Office14\MSACCESS.EXE"" ""C:\Access\Access2010\DocketControl\EPdocket2010.adp"" /cmd " & CStr(rst.Fields("MarkID").Value)
    '                    Else
    '                        strScratch = "start ""C:\Program Files (x86)\Microsoft Office\Office14\MSACCESS.EXE"" ""C:\Tekhelps\EPdocket.ade"" /cmd " & CStr(rst.Fields("MarkID").Value)
    '                    End If
    '                    objStreamWriter.WriteLine(strScratch)
    '                    'Close the file.
    '                    objStreamWriter.Close()

    '                    strHTML = "<HTML><BODY><P>This mark was Suspended " & rst.Fields("Suspended").Value & ".</P><P>" & _
    '                              "You will be reminded of this every four months until the Suspended field is cleared or the mark's application is abandoned.</P>"
    '                    strScratch = "<P>"
    '                    If Not IsDBNull(rst.Fields("Trademark").Value) Then strScratch = strScratch & strCourierOn & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Trademark: </font><strong>" & rst.Fields("Trademark").Value & "</strong><BR>"
    '                    If Not IsDBNull(rst.Fields("SerialNo").Value) Then strScratch = strScratch & strCourierOn & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Serial No: </font>" & rst.Fields("SerialNo").Value & "<BR>"
    '                    If Not IsDBNull(rst.Fields("RegistrationNo").Value) Then strScratch = strScratch & strCourierOn & "RegistrationNo: </font>" & rst.Fields("RegistrationNo").Value & "<BR>"
    '                    If Not IsDBNull(rst.Fields("Jurisdiction").Value) Then strScratch = strScratch & strCourierOn & "&nbsp;&nbsp;Jurisdiction: </font>" & rst.Fields("Jurisdiction").Value & "<BR>"
    '                    If Not IsDBNull(rst.Fields("ApplicantName").Value) Then strScratch = strScratch & strCourierOn & "&nbsp;ApplicantName: </font>" & rst.Fields("ApplicantName").Value & "<BR>"
    '                    If Not IsDBNull(rst.Fields("GoodsServices").Value) Then strScratch = strScratch & strCourierOn & "Goods/Services: </font>" & rst.Fields("GoodsServices").Value
    '                    strScratch = strScratch & "<P>" & strCourierOn & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Mark ID: </font>" & CStr(rst.Fields("MarkID").Value) & _
    '                                        " <a href=""file://" & strFile & """>Click here to open Mark in IP Dashboard.</a>"
    '                    strHTML = strHTML & strScratch & "</P></BODY></HTML>"
    '                    Email.Body = strHTML

    '                    If chkDontSendMail.CheckState = CheckState.Unchecked Then
    '                        If SendEmail(Email) AndAlso chkShowMessages.CheckState Then
    '                            MsgBox("Mail sent to " & strTo & ".", MsgBoxStyle.Information, cmdEmailTest.Text)
    '                        End If
    '                    End If
    '                End Using

    '                If chkDontUpdateDatabase.CheckState = CheckState.Checked Then
    '                    .CancelUpdate()
    '                Else
    '                    ' update the database that the email was sent
    '                    .Fields("LastEmailSuspended").Value = Now
    '                    .Update()
    '                End If
    '                intCounter = intCounter + 1
    'NextNotice:
    '                If chkShowMessages.CheckState Then
    '                    strScratch = "Email(s) sent to: " & strTo & vbNewLine & vbNewLine & "Process the next item?"
    '                    If MsgBox(strScratch, MsgBoxStyle.YesNo + MsgBoxStyle.Question, strTitle) = MsgBoxResult.No Then GoTo FinishedLoop
    '                End If
    '                .MoveNext()
    '            Loop
    'FinishedLoop:
    '            .Close()
    '        End With
    '        cnn.Close()
    '        strDeadlines = strDeadlines & "</P><P>" & _
    '                       IIf(intCounter = 1, "One reminder E-mail was", intCounter & " reminder E-mails were") & _
    '                       " sent for " & IIf(intCounter = 1, "a Suspended IP mark", "Suspended IP marks") & ".</P>"

    '        If chkShowMessages.CheckState Or isDev Then MsgBox("Finished sending " & intCounter & " Email(s)", MsgBoxStyle.Information, strTitle)
    '        Exit Sub
    '    End Sub

    Private Sub cmdBothQuit_Click(sender As Object, e As EventArgs) Handles cmdBothQuit.Click
        SendBothAndQuit()
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        '12/14/2021 added this
        cnn = Nothing
    End Sub
End Class