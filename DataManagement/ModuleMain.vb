

Imports System.Net.Mail

Imports System.Data.SqlClient

Module ModuleMain


    Public MyDocumentsPath As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
    Public LogFileName As String = MyDocumentsPath & "\ImpalaDataIntegrationServicesUserLog_" & Format(Now, "yyyyMMdd") & ".txt"
    Public LogFileNameErrors As String = MyDocumentsPath & "\ImpalaIntegrationServicesErrorLog_" & Format(Now, "yyyyMMdd") & ".txt"


    'this holds external data
    Public strconnectSQLBizCoachExternalData As String = "UID=" & "BizCoachDataLoading" & ";PWD=" & "JaMyt45!ord" & ";DATABASE=" & "BizCoachExternalData" &
                         ";SERVER=" & "ec2-13-229-212-138.ap-southeast-1.compute.amazonaws.com,1433" & ";Provider=SQLOLEDB"


    'this holds the customer names
    Public strconnectSQLBizCoachCustomerInformationDatabase As String = "UID=" & "BizCoachDataLoading" & ";PWD=" & "JaMyt45!ord" & ";DATABASE=" & "BizCoachCustomers" &
                         ";SERVER=" & "ec2-13-229-212-138.ap-southeast-1.compute.amazonaws.com,1433" & ";Provider=SQLOLEDB"


    'this holds the settings
    'new ec2 sql express - non RDS
    Public strconnectSQLBizCoachSettings As String = "UID=" & "BizCoachDataLoading" & ";PWD=" & "JaMyt45!ord" & ";DATABASE=" & "BizCoachSettings" &
                           ";SERVER=" & "ec2-13-229-212-138.ap-southeast-1.compute.amazonaws.com,1433" & ";Provider=SQLOLEDB"

    'old RDS server sql professional
    '    Public strconnectSQLBizCoachSettings As String = "UID=" & "BizCoachSQLStd" & ";PWD=" & "PaLm78jxY" & ";DATABASE=" & "BizCoachSettings" &
    '                           ";SERVER=" & "awsrdsbizcoachsqlstandard.c5sz1nxmndnm.ap-southeast-1.rds.amazonaws.com,1433" & ";Provider=SQLOLEDB"



    Public strconnectPostGresSQLUsers As String = "Driver={PostgreSQL ANSI};Server=api-bizcoach-production-server-db.c5sz1nxmndnm.ap-southeast-1.rds.amazonaws.com;Port=5432;Database=bizcoach_server_db;Uid=onboarding_app;Pwd=BnrA178@?;"


    Public strconnectSQLCommandTypeControl As SqlConnection = New SqlConnection("")

    Public strconnectSQLBizCoachCustomerDataSQL As String = ""

    Public strconnectSQLBizCoachCustomerDataSQLForInsert As String = ""

    Public VendCustomerUpdateRunning As Boolean = False
    Public KPIUpdateRunning As Boolean = False

    Public CustomerDatabaseName As String = ""

    'this holds the data for analysis
    '   Public DataForAnalysis(10000, 5)

    'this for the push notifications
    Public PushNotificationsRunning As Boolean = False

    Public PushNotificationsLoginRunning As Boolean = False
    Public PushNotificationsLoginRunningDate As Date = Now

    Public PushNotificationsNotifyRunning As Boolean = False
    Public PushNotificationsNotifyRunningDate As Date = Now

    'holds details of logins not needed to be recorded
    Public LoginsNotNeeded(300, 2)

    'this details the users and phone numbers
    Public CustomersDetails(1000, 2)

    'this details the specific KPI names
    Public CustomersKPIDetails(1000, 2)

    Public DateG1LastRan As Date = Now
    Public DateKPIsLastRan As Date = Now
    Public NotificationsRunning As Boolean = False


    'used to see if selection of notification has changed
    Public NotificationTypeNameOld As String = ""
    Public NotificationCompanyOld As String = ""

    Public WeekStartsOn As String = ""


    Sub Delay(ByVal dblSecs As Double)
        Const OneSec As Double = 1.0# / (1440.0# * 60.0#)
        Dim dblWaitTil As Date
        Now.AddSeconds(OneSec)
        dblWaitTil = Now.AddSeconds(OneSec).AddSeconds(dblSecs)
        Do Until Now > dblWaitTil
            ' Application.DoEvents() ' Allow windows messages to be processed
        Loop
    End Sub

    Public Sub AutoEmailMessage365(ByVal Emailmessage As String, ByVal FromString As String, ByVal MailSubject As String,
                             ByVal IsHTMLTemp As Boolean, ByVal Recipients As String,
                             ByVal File01ToSendLocation As String, ByVal File02ToSendLocation As String,
                             ByVal TypeOfMessage As String, ByVal RecipientsCC As String, ByVal RecipientsBCC As String, ByVal BizCoachSender As String)
        'change to ms exchange
        Dim client As New SmtpClient()
        client.UseDefaultCredentials = False
        If BizCoachSender = "applicationserver@impalacloud.com" Then
            client.Credentials = New System.Net.NetworkCredential("applicationserver@impalacloud.com", "XmR5617@")
        ElseIf BizCoachSender = "will.odwyer@impalacloud.com" Then
            client.Credentials = New System.Net.NetworkCredential("will.odwyer@impalacloud.com", "JasperPass16!")

        Else
            'default
            BizCoachSender = "applicationserver@impalacloud.com"
            client.Credentials = New System.Net.NetworkCredential("applicationserver@impalacloud.com", "XmR5617@")
        End If

        '"reports@secbattery.com", "Wobo4960"
        '"will.odwyer@impalacloud.com", "JamesPwd16!"
        client.Port = 587
        'client.Port = 25
        ' You can use Port 25 if 587 is blocked

        client.Host = "smtp.outlook.office365.com"
        client.Host = "smtp.office365.com"
        client.DeliveryMethod = SmtpDeliveryMethod.Network
        client.EnableSsl = True
        Dim mail As New MailMessage()
        mail = New MailMessage()

        'Dim addr() As String = Labelemailaddresses.Text.Split(",")
        '     Dim addr() As String = "Wod@impalacloud.com"
        Try
            '            mail.From = New MailAddress("will.odwyer@impalacloud.com")

            mail.From = New MailAddress(BizCoachSender.Trim)

            If IsHTMLTemp = True Then
                mail.IsBodyHtml = True
            Else
                mail.IsBodyHtml = False
            End If

            If Recipients = "" Then Recipients = "will.odwyer@impalacloud.com"

            mail.To.Add(Recipients)
            If RecipientsCC <> "" Then
                mail.CC.Add(RecipientsCC)
            End If

            If RecipientsBCC <> "" Then
                mail.Bcc.Add(RecipientsBCC)
            End If
            If RecipientsBCC = "" And Recipients <> "will.odwyer@impalacloud.com" Then
                RecipientsBCC = "will.odwyer@impalacloud.com"
            End If

            mail.Subject = MailSubject
            mail.Body = Emailmessage & ControlChars.CrLf & ControlChars.CrLf &
                     ControlChars.CrLf & ControlChars.CrLf
            If File01ToSendLocation <> "" Then
                mail.Attachments.Add(New Attachment(File01ToSendLocation))
            End If
            mail.DeliveryNotificationOptions =
                    DeliveryNotificationOptions.OnFailure
            client.Send(mail)
        Catch ex As Exception
            WriteToSystemErrorLog(ex.ToString() & ControlChars.CrLf & BizCoachSender)
        End Try

    End Sub

    Public Sub WriteToSystemErrorLog(ByVal MessageText)

        LogFileNameErrors = MyDocumentsPath & "\BizCoachDataErrorLog_" & Format(Now, "yyyyMMdd") & ".txt"
        Try
            Dim objWriter As New System.IO.StreamWriter(LogFileNameErrors, True)
            objWriter.WriteLine(MessageText & "  :  " & Format(Now, "dd/MMM/yyyy hh:mm tt"))
            objWriter.Close()
        Catch ex As Exception
            'if cant write to error log just write a single line to a new file
            Try
                Dim objWriter2 As New System.IO.StreamWriter(MyDocumentsPath & "\BizCoachErrorLog_" & Format(Now, "yyyyMMddsss") & ".txt", True)
                objWriter2.WriteLine(MessageText & "  :  " & Format(Now, "dd/MMM/yyyy hh:mm tt"))
            Catch
            End Try
        End Try
    End Sub


    Public Sub UpdateCustomerLog(ByVal CustomerName, ByVal StartDateTime, ByVal EndDateTime, ByRef UpdateStatus,
                                 ByVal EventLongText, ByVal EventStart, ByVal EventEnd, ByVal UpdateType, ByVal OutletName)
        'This is the standard update to the system log
        Dim rs As ADODB.Recordset
        rs = New ADODB.Recordset
        rs.CursorType = ADODB.CursorTypeEnum.adOpenDynamic
        rs.LockType = ADODB.LockTypeEnum.adLockOptimistic
        'writes to the databse the details of the run
        Dim UniqueIDTemt = System.Guid.NewGuid

        Dim UpdateSQLStringTemp As String = "Insert into CustomerStatusLog (CompanyName, [Date Record], [Unique Identity]," &
                        " EventType, DateProcessedFrom, DateProcessedTo, " &
                        "[EventLongText], [DateTimeStarted], [DateTimeEnded], " &
                        "[Outlet Name]) Values ('" &
                        CustomerName & "','" & Format(Now, "dd-MMM-yyyy") & "','" & UniqueIDTemt.ToString & "','" &
                        Mid(Replace(UpdateType, "'", "''"), 1, 150) & "','" & Format(StartDateTime, "dd-MMM-yyyy HH:mm:ss") & "','" & Format(EndDateTime, "dd-MMM-yyyy HH:mm:ss") &
                        "','" & Replace(EventLongText, "'", "''") & "','" & Format(EventStart, "dd-MMM-yyyy HH:mm:ss") & "','" & Format(EventEnd, "dd-MMM-yyyy HH:mm:ss") & "','" &
                        OutletName & "')"
        Try
            rs.Open(UpdateSQLStringTemp, strconnectSQLBizCoachCustomerInformationDatabase)
            UpdateStatus = "Updated to system Log " & Now
        Catch ex As Exception
            UpdateStatus = "Cant update customer log!  " & Err.Description
        End Try
    End Sub


    Public Sub GetDateCalculatedFields(ByVal DateGiven As Date, ByRef YearCode As Int16, ByRef MonthCode As String, ByRef DayOfWeekNumber As Int16,
                                        ByRef DayOfYear As Int16, ByRef MonthCodeShort As String,
                                        ByRef DayOfWeekName As String, ByRef DayOfWeekNameShort As String, ByRef WeekNumber As Int16,
                                        ByRef Weekend As Boolean, ByRef Weekday As Boolean, ByRef Workday As Boolean,
                                        ByRef NumberWorkDaysInMonth As Int16, ByRef NumberNonWorkDaysInMonth As Int16,
                                        ByRef PublicHoliday As Boolean, ByRef NumberDaysInMonth As Int16, ByRef DayOfMonth As Int16,
                                        ByRef MonthNumber As String, ByRef MonthNumber2Digits As String)

        'need day of month
        YearCode = DateGiven.Year
        MonthCode = Format(DateGiven, "MMMM")
        DayOfWeekNumber = DateGiven.DayOfWeek
        DayOfYear = DateGiven.DayOfYear
        MonthCodeShort = Format(DateGiven, "MMM")
        DayOfWeekName = Format(DateGiven, "dddd")
        DayOfWeekNameShort = Format(DateGiven, "ddd")
        DayOfMonth = DateGiven.Day
        ' ISO 8601 definition for week 01 is the week with the year's first Thursday in it and starts on Monday
        MonthNumber = DateGiven.Month
        MonthNumber2Digits = DateGiven.Month
        If Len(MonthNumber2Digits) = 1 Then
            MonthNumber2Digits = "0" & MonthNumber2Digits
        End If

        'needs work for other weekend types (middle east countries)
        '    MsgBox(DateGiven.DayOfWeek)
        If DateGiven.DayOfWeek = DayOfWeek.Saturday Or DateGiven.DayOfWeek = DayOfWeek.Sunday Then
            Weekend = True
            Weekday = False
            Workday = False
        Else
            Weekend = False
            Weekday = True
            Workday = True
        End If
        Dim WekkNumTemp As Int16 = 0

        'week numbers
        'NEEDS Work
        '2016 week 1 starts Jan 4
        '2017 week 1 starts January 2, 2017
        '2018 week 1 starts January 1, 2018
        '2019 week 1 starts December 31, 2018
        '2020 week 1 starts December 30, 2019
        '2021 week 1 starts January 4, 2021
        '2022 week 1 starts January 3, 2022

        Dim IsLeapYear As Boolean = False
        WeekNumber = 0

        Dim StartOfMonth As Date = DateGiven.AddDays(-DateGiven.Day + 1)
        '30 days have sept apr june and nov
        'others 31 except feb
        If DateGiven.Month = 9 Or DateGiven.Month = 4 Or DateGiven.Month = 6 Or DateGiven.Month = 11 Then
            NumberDaysInMonth = 30
        ElseIf DateGiven.Month = 2 Then
            Try
                IsLeapYear = IsDate("29/Feb/" & DateGiven.Year)
                '  IsLeapYear = IsDate("29/Feb/2015")
            Catch ex As Exception
            End Try
            If IsLeapYear Then
                NumberDaysInMonth = 29
            Else
                NumberDaysInMonth = 28
            End If
        Else
            NumberDaysInMonth = 31
        End If

        Dim EndOfMonth As Date = DateGiven.AddDays(-DateGiven.Day + NumberDaysInMonth)
        Dim Days As Int16 = DateDiff(DateInterval.Day, StartOfMonth, EndOfMonth) + 1
        Dim Weeks = Days \ 7
        Days = Days Mod 7
        If Days > 0 Then
            If StartOfMonth.DayOfWeek = DayOfWeek.Sunday Or EndOfMonth.DayOfWeek = DayOfWeek.Saturday Then
                Days = Days - 1
            ElseIf EndOfMonth.DayOfWeek < StartOfMonth.DayOfWeek Then
                Days = Days - 2
            End If
        End If
        Dim Weekdays = Weeks * 5 + Days

        NumberWorkDaysInMonth = Weekdays
        NumberNonWorkDaysInMonth = NumberDaysInMonth - Weekdays
        PublicHoliday = False
        WeekNumber = DatePart("ww", DateGiven, vbSunday)
    End Sub


    Public Sub UpdateTextFile(ByVal FileName As String, ByVal KPINameTemp As String, ByVal StartDateFrom01 As String,
                             ByVal StartDateTo01 As String, ByVal StartDateFromtext02 As String, ByVal StartDateTotext02 As String)

        Dim path As String = MyDocumentsPath & "\" & FileName & "_" & Format(Now, "yyyyMMdd") & ".csv"

        ' This text is added only once to the file. 
        If Not System.IO.File.Exists(path) Then
            ' Create a file to write to.
            Using sw As System.IO.StreamWriter = System.IO.File.CreateText(path)
                sw.WriteLine(Format(Now, "d/MMM/yyyy"))
            End Using
        End If

        ' This text is always added, making the file longer over time 
        ' if it is not deleted.
        Using sw As System.IO.StreamWriter = System.IO.File.AppendText(path)
            sw.WriteLine(KPINameTemp & ", " & StartDateFrom01 & ", " & StartDateTo01 & ", " & StartDateFromtext02 & ", " & StartDateTotext02)

        End Using



    End Sub


End Module
