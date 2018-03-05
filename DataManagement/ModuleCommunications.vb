Imports System.IO
Imports System.Net
Imports System.Security.Cryptography
Imports Amazon.Polly
Imports Amazon.Polly.Model
Imports Microsoft.Office.Interop



'*********************************************************************************
'*********
'********* this is a standard module that deals with communications in BizCoach
'********* Do not edit outside of 'DataManagement'
'*********
'*********************************************************************************


Module ModuleCommunications


    Public Sub CreateFileName(ByVal CustomerName As String, ByVal DateOfFile As Date, ByRef FileNameToUse As String)

        'this creates a pseudo encrypted file name
        CustomerName = Replace(CustomerName, " ", "")
        CustomerName = Replace(CustomerName, " ", "")
        CustomerName = CustomerName.ToLower

        FileNameToUse = Format(Now, "ddMMMyy") + CustomerName

        'this does a daily substitution
        If Now.DayOfWeek = DayOfWeek.Monday Then

            FileNameToUse = Replace(FileNameToUse, "0", "V")
            FileNameToUse = Replace(FileNameToUse, "1", "j")
            FileNameToUse = Replace(FileNameToUse, "2", "r")
            FileNameToUse = Replace(FileNameToUse, "3", "Q")
            FileNameToUse = Replace(FileNameToUse, "4", "1")
            FileNameToUse = Replace(FileNameToUse, "5", "S")
            FileNameToUse = Replace(FileNameToUse, "6", "k")
            FileNameToUse = Replace(FileNameToUse, "7", "E")
            FileNameToUse = Replace(FileNameToUse, "8", "M")
            FileNameToUse = Replace(FileNameToUse, "9", "x")

            FileNameToUse = Replace(FileNameToUse, "a", "1")
            FileNameToUse = Replace(FileNameToUse, "b", "M")
            FileNameToUse = Replace(FileNameToUse, "c", "3")
            FileNameToUse = Replace(FileNameToUse, "d", "L")
            FileNameToUse = Replace(FileNameToUse, "e", "R")
            FileNameToUse = Replace(FileNameToUse, "f", "5")
            FileNameToUse = Replace(FileNameToUse, "g", "N")
            FileNameToUse = Replace(FileNameToUse, "h", "m")
            FileNameToUse = Replace(FileNameToUse, "i", "6")
            FileNameToUse = Replace(FileNameToUse, "j", "m")
            FileNameToUse = Replace(FileNameToUse, "k", "M")
            '     FileNameToUse = Replace(FileNameToUse, "l", "m")
            '    FileNameToUse = Replace(FileNameToUse, "m", "m")
            '            FileNameToUse = Replace(FileNameToUse, "n", "m")
            '           FileNameToUse = Replace(FileNameToUse, "o", "m")
            FileNameToUse = Replace(FileNameToUse, "p", "q")
            '         FileNameToUse = Replace(FileNameToUse, "q", "m")
            '        FileNameToUse = Replace(FileNameToUse, "r", "m")
            '           FileNameToUse = Replace(FileNameToUse, "s", "m")
            FileNameToUse = Replace(FileNameToUse, "t", "a")
            '         FileNameToUse = Replace(FileNameToUse, "u", "m")
            '        FileNameToUse = Replace(FileNameToUse, "v", "m")
            '       FileNameToUse = Replace(FileNameToUse, "w", "m")
            '      FileNameToUse = Replace(FileNameToUse, "x", "m")
            '     FileNameToUse = Replace(FileNameToUse, "y", "m")
            '    FileNameToUse = Replace(FileNameToUse, "z", "m")

        End If

        FileNameToUse = FileNameToUse & ".htm"

    End Sub





    Public Sub DailySalesReportPOSYesterday(ByRef DailyReportHTMLEmailString As String, ByRef DailyReportDate As Date, ByVal DemoCompany As Boolean)

        'this is a daily report that is based upon yesterdays sales
        'it compares this to the monthly and last month figures

        'dont use KPIs now as may be a timing issue


        'as is yesterday it can be preset
        DailyReportDate = Now.AddDays(-1)


        Dim erpdb As ADODB.Connection
        Dim records As ADODB._Recordset
        Dim recordsLookup As ADODB._Recordset
        erpdb = New ADODB.Connection
        erpdb.Open(strconnectSQLBizCoachCustomerDataSQL)
        Dim SalesYesterday As Integer = 0
        Dim SalesMTDYesterday As Integer = 0
        Dim OutletName As String = ""
        Dim SalesYesterdayLastWeek As Integer = 0
        Dim SalesLastMonthVariance As Decimal = 0
        Dim SalesYesterdayLastWeekVariance As Decimal = 0
        Dim SalesLastMonthToDate As Integer = 0
        Dim StartDateFrom01 As Date = Now.ToShortDateString
        Dim StartDateTo01 As Date = Now.ToShortDateString
        Dim SQLStatement As String = ""

        'used for demo companies to go back 1 year

        Dim DemoParameter As Int16 = 0
        'uses UTC time so need to take away 8 hours or use HongKongLoginMoment

        If DemoCompany = True Then
            DemoParameter = 365
            DailyReportDate = DailyReportDate.AddYears(-1)
        End If
        DailyReportDate = DailyReportDate.ToShortDateString

        '     records = erpdb.Execute("Select * from [KPI Values] where [KPI Name] ='POS Daily Sales (yesterday) - Total'")

        SQLStatement = "Select sum([Total Net Sales]) as Sumsales from [POSSales] where [Date Record] = '" & Format(DailyReportDate, "dd-MMM-yyyy") & "'"
        records = erpdb.Execute("Select sum([Total Net Sales]) as Sumsales from [POSSales] where cast([Date Record] as date) = '" & Format(DailyReportDate, "dd-MMM-yyyy") & "'")

        SalesYesterday = 0
        While Not records.EOF
            Try
                SalesYesterday = records.Fields("Sumsales").Value
            Catch ex As Exception
                SalesYesterday = 0
            End Try
            records.MoveNext()
        End While
        records.Close()


        StartDateFrom01 = Now.AddDays((DemoParameter) * -1)
        StartDateFrom01 = StartDateFrom01.AddDays((Now.Day * -1) + 1)
        StartDateTo01 = Now.AddDays((DemoParameter) * -1)
        StartDateTo01 = StartDateTo01.AddDays(-1)
        'if the first then last month is whole month
        If Now.Day = 1 Then
            StartDateFrom01 = Now.AddDays((DemoParameter) * -1)
            StartDateFrom01 = StartDateFrom01.AddMonths(-1)
            StartDateTo01 = Now.AddDays((DemoParameter) * -1)
            StartDateTo01 = StartDateTo01.AddDays(-1)
        End If

        records = erpdb.Execute(" select sum([Total Net Sales]) as Sumsales from [POSSales] where cast([Date Record] as date) >= '" &
                                Format(StartDateFrom01, "dd-MMM-yyyy") &
                                "' and cast([Date Record] as date) <= '" & Format(StartDateTo01, "dd-MMM-yyyy") & "'")
        SalesMTDYesterday = 0
        While Not records.EOF
            Try
                SalesMTDYesterday = records.Fields("Sumsales").Value
            Catch ex As Exception
                SalesMTDYesterday = 0
            End Try
            records.MoveNext()
        End While
        records.Close()
        '      records = erpdb.Execute("Select * from [KPI Values] where [KPI Name] ='POS Daily Sales (yesterday same day last week) - Total'")

        records = erpdb.Execute("select sum([Total Net Sales]) as Sumsales from [POSSales] where cast([Date Record] as date) = '" & Format(DailyReportDate.AddDays(-7), "dd-MMM-yyyy") & "'")

        SalesYesterdayLastWeek = 0
        While Not records.EOF
            Try
                SalesYesterdayLastWeek = records.Fields("Sumsales").Value
            Catch ex As Exception
                SalesYesterdayLastWeek = 0
            End Try
            records.MoveNext()
        End While
        records.Close()


        Try
            SalesYesterdayLastWeekVariance = ((SalesYesterday - SalesYesterdayLastWeek)) / SalesYesterdayLastWeek * 100
        Catch ex As Exception
            SalesYesterdayLastWeekVariance = 0
        End Try

        '     SalesYesterdayLastWeekVariance = SalesYesterday / SalesYesterdayLastWeek
        '      records = erpdb.Execute("Select * from [KPI Values] where [KPI Name] ='POS Sales LMTD - Total'")

        StartDateFrom01 = Now.AddDays((DemoParameter) * -1)
        StartDateFrom01 = StartDateFrom01.AddMonths(-1)
        StartDateFrom01 = StartDateFrom01.AddDays((Now.Day * -1) + 1)

        StartDateTo01 = Now.AddDays((DemoParameter) * -1)
        StartDateTo01 = StartDateTo01.AddDays(-1)
        StartDateTo01 = StartDateTo01.AddMonths(-1)
        'if the first then last month is whole month
        'lmtd is month -2
        If Now.Day = 1 Then
            StartDateFrom01 = Now.AddDays((DemoParameter) * -1)
            StartDateFrom01 = StartDateFrom01.AddMonths(-2)
            StartDateTo01 = Now.AddDays((DemoParameter) * -1)
            StartDateTo01 = StartDateTo01.AddMonths(-1)
            StartDateTo01 = StartDateTo01.AddDays(-1)

        End If

        '        records = erpdb.Execute("Select * from [KPI Values] where [KPI Name] ='POS Sales LMTD - Total'")

        records = erpdb.Execute("Select sum([Total Net Sales]) As Sumsales from [POSSales] where cast([Date Record] as date) >= '" & Format(StartDateFrom01, "dd-MMM-yyyy") &
                                "' and cast([Date Record] as date) <= '" & Format(StartDateTo01, "dd-MMM-yyyy") & "'")
        SQLStatement = "Select sum([Total Net Sales]) As Sumsales from [POSSales] where cast([Date Record] as date) >= '" & Format(StartDateFrom01, "dd-MMM-yyyy") &
                                "' and cast([Date Record] as date) <= '" & Format(StartDateTo01, "dd-MMM-yyyy") & "'"

        SalesLastMonthToDate = 0
        While Not records.EOF
            Try
                SalesLastMonthToDate = records.Fields("Sumsales").Value
            Catch ex As Exception
                SalesLastMonthToDate = 0
            End Try
            records.MoveNext()
        End While
        records.Close()



        '    Format((ResultsArray(countertemp, 1) - ResultsArray(countertemp, 3)) / ResultsArray(countertemp, 3) * 100, "0.##")
        Try
            SalesLastMonthVariance = (SalesMTDYesterday - SalesLastMonthToDate) / SalesLastMonthToDate * 100
        Catch ex As Exception
            SalesLastMonthVariance = 0
        End Try



        'Yesterday last week
        '      DailyReportHTMLEmailString = "      <div style = ""width:400px;"">
        '         <h1 style=""text-align:center""> <background-color: blue;>Testing </h1>
        '        <p  style =""text-align:right""><a href=""#"">sample link</a> </p>
        '       </div>"

        'puts the bizcoach logo
        DailyReportHTMLEmailString = "    <img src = ""http://aagilitycom.ipage.com/wp-content/uploads/2017/12/BizCoach_small.png"" alt=""BizCoach"" ;  ><br /><br />"

        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "Group Daily Sales for : " & Format(DailyReportDate, "ddd dd MMM yyyy") & "<br /><br />" &
            "Yesterday: " & Format(SalesYesterday, "HK$ #,###,##0") & ", same day last week : " & Format(SalesYesterdayLastWeek, "HK$ #,###,##0") & ", variance to last week " & Format(SalesYesterdayLastWeekVariance, "0.##") & "%" & "<br />" &
            "MTD      : " & Format(SalesMTDYesterday, "HK$ #,###,##0") & ", last month to date " & "" & Format(SalesLastMonthToDate, "HK$ #,###,##0") & ", variance to last month " & Format(SalesLastMonthVariance, "0.##") & "%" & "<br />" &
                                                     "<br />" & "<br />" &
                                                     "<table border=1><col width=""300""><col width=""150""><col width=""150""><tr>" &
                                                     "<td><b>Outlet Sales</td>" &
                                                     "<td><b><p  style =""text-align:center"">Yesterday HK$</td>" &
                                                     "<td><b><p  style =""text-align:center"">MTD   HK$</b></td>" &
                                                     "</b></tr>"
        'now gets the individual outlets
        'now sums up
        '     records = erpdb.Execute("Select sum([KPI Value Number]) as Sumsales, [Outlet Name] from [KPI Values] where [KPI Name] ='POS Daily Sales (today)' group by [Outlet Name] order by [Outlet Name]")

        records = erpdb.Execute("select sum([Total Net Sales]) as Sumsales , [Outlet Name] from [POSSales] where cast([Date Record] as date) = '" & _
                                Format(DailyReportDate, "dd-MMM-yyyy") & "' group by [Outlet Name] order by [Outlet Name]")

        'MTD
        StartDateFrom01 = Now.AddDays((DemoParameter) * -1)
        StartDateFrom01 = StartDateFrom01.AddDays((Now.Day * -1) + 1)
        StartDateTo01 = Now.AddDays((DemoParameter) * -1)
        StartDateTo01 = StartDateTo01.AddDays(-1)
        'if the first then last month is whole month
        If Now.Day = 1 Then
            StartDateFrom01 = Now.AddDays((DemoParameter) * -1)
            StartDateFrom01 = StartDateFrom01.AddMonths(-1)
            StartDateTo01 = Now.AddDays((DemoParameter) * -1)
            StartDateTo01 = StartDateTo01.AddDays(-1)
        End If

        SalesYesterday = 0
        While Not records.EOF

            Try
                OutletName = records.Fields("Outlet Name").Value
                SalesYesterday = records.Fields("Sumsales").Value
            Catch ex As Exception
                MsgBox("error " & Err.Description)
            End Try
            'now needs to get the MTD for the outlet

            '         recordsLookup = erpdb.Execute("Select * from [KPI Values] where [KPI Name] ='POS Sales MTD by Dimensions' and [Outlet Name] ='" & Replace(OutletName, "'", "''") & "'")
            SQLStatement = "Select sum([Total Net Sales]) As Sumsales, [Outlet Name] from [POSSales] where cast([Date Record] As Date) >= '" & Format(StartDateFrom01, "dd-MMM-yyyy") &
                                          "' and cast([Date Record] as date) <= '" & Format(StartDateTo01, "dd-MMM-yyyy") & "' and [Outlet Name] ='" & Replace(OutletName, "'", "''") & "' group by [Outlet Name] "

            recordsLookup = erpdb.Execute("select sum([Total Net Sales]) as Sumsales, [Outlet Name] from [POSSales] where cast([Date Record] as date) >= '" & Format(StartDateFrom01, "dd-MMM-yyyy") &
                                          "' and cast([Date Record] as date) <= '" & Format(StartDateTo01, "dd-MMM-yyyy") & "' and [Outlet Name] ='" & Replace(OutletName, "'", "''") & "' group by [Outlet Name] ")

            SalesMTDYesterday = 0
            While Not recordsLookup.EOF
                Try
                    SalesMTDYesterday = recordsLookup.Fields("Sumsales").Value
                Catch ex As Exception
                    SalesMTDYesterday = 0
                End Try

                recordsLookup.MoveNext()
            End While
            recordsLookup.Close()
            'can build the line
            DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<tr><td>" & OutletName & "</td>" &
                                         "<td>" & "<p  style =""text-align:right"">" & Format(SalesYesterday, "#,###,###") & "</td>" &
                                        "<td>" & "<p  style =""text-align:right"">" & Format(SalesMTDYesterday, "#,###,###") & "</tr>"


            records.MoveNext()
        End While
        records.Close()
        erpdb.Close()

        'closes the table
        DailyReportHTMLEmailString = "" & DailyReportHTMLEmailString & "</table></span><br /><hr /><br />"
        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<br />"


    End Sub


    Public Sub DailySalesReportPOSYesterdayExcel(ByRef DailyReportHTMLEmailString As String, ByRef DailyReportDate As Date, ByVal DemoCompany As Boolean)

        'this is a daily report that is based upon yesterdays sales
        'it compares this to the monthly and last month figures

        'dont use KPIs now as may be a timing issue


        'as is yesterday it can be preset
        DailyReportDate = Now.AddDays(-1)


        Dim erpdb As ADODB.Connection
        Dim records As ADODB._Recordset
        Dim recordsLookup As ADODB._Recordset
        erpdb = New ADODB.Connection
        erpdb.Open(strconnectSQLBizCoachCustomerDataSQL)
        Dim SalesYesterday As Integer = 0
        Dim SalesMTDYesterday As Integer = 0
        Dim OutletName As String = ""
        Dim SalesYesterdayLastWeek As Integer = 0
        Dim SalesLastMonthVariance As Decimal = 0
        Dim SalesYesterdayLastWeekVariance As Decimal = 0
        Dim SalesLastMonthToDate As Integer = 0
        Dim StartDateFrom01 As Date = Now.ToShortDateString
        Dim StartDateTo01 As Date = Now.ToShortDateString
        Dim SQLStatement As String = ""

        'used for demo companies to go back 1 year

        Dim DemoParameter As Int16 = 0
        'uses UTC time so need to take away 8 hours or use HongKongLoginMoment

        If DemoCompany = True Then
            DemoParameter = 365
            DailyReportDate = DailyReportDate.AddYears(-1)
        End If
        DailyReportDate = DailyReportDate.ToShortDateString

        '     records = erpdb.Execute("Select * from [KPI Values] where [KPI Name] ='POS Daily Sales (yesterday) - Total'")

        SQLStatement = "Select sum([Total Net Sales]) as Sumsales from [POSSales] where [Date Record] = '" & Format(DailyReportDate, "dd-MMM-yyyy") & "'"
        records = erpdb.Execute("Select sum([Total Net Sales]) as Sumsales from [POSSales] where cast([Date Record] as date) = '" & Format(DailyReportDate, "dd-MMM-yyyy") & "'")

        SalesYesterday = 0
        While Not records.EOF
            Try
                SalesYesterday = records.Fields("Sumsales").Value
            Catch ex As Exception
                SalesYesterday = 0
            End Try
            records.MoveNext()
        End While
        records.Close()


        StartDateFrom01 = Now.AddDays((DemoParameter) * -1)
        StartDateFrom01 = StartDateFrom01.AddDays((Now.Day * -1) + 1)
        StartDateTo01 = Now.AddDays((DemoParameter) * -1)
        StartDateTo01 = StartDateTo01.AddDays(-1)
        'if the first then last month is whole month
        If Now.Day = 1 Then
            StartDateFrom01 = Now.AddDays((DemoParameter) * -1)
            StartDateFrom01 = StartDateFrom01.AddMonths(-1)
            StartDateTo01 = Now.AddDays((DemoParameter) * -1)
            StartDateTo01 = StartDateTo01.AddDays(-1)
        End If

        records = erpdb.Execute(" select sum([Total Net Sales]) as Sumsales from [POSSales] where cast([Date Record] as date) >= '" &
                                Format(StartDateFrom01, "dd-MMM-yyyy") &
                                "' and cast([Date Record] as date) <= '" & Format(StartDateTo01, "dd-MMM-yyyy") & "'")
        SalesMTDYesterday = 0
        While Not records.EOF
            Try
                SalesMTDYesterday = records.Fields("Sumsales").Value
            Catch ex As Exception
                SalesMTDYesterday = 0
            End Try
            records.MoveNext()
        End While
        records.Close()
        '      records = erpdb.Execute("Select * from [KPI Values] where [KPI Name] ='POS Daily Sales (yesterday same day last week) - Total'")

        records = erpdb.Execute("select sum([Total Net Sales]) as Sumsales from [POSSales] where cast([Date Record] as date) = '" & Format(DailyReportDate.AddDays(-7), "dd-MMM-yyyy") & "'")

        SalesYesterdayLastWeek = 0
        While Not records.EOF
            Try
                SalesYesterdayLastWeek = records.Fields("Sumsales").Value
            Catch ex As Exception
                SalesYesterdayLastWeek = 0
            End Try
            records.MoveNext()
        End While
        records.Close()


        Try
            SalesYesterdayLastWeekVariance = ((SalesYesterday - SalesYesterdayLastWeek)) / SalesYesterdayLastWeek * 100
        Catch ex As Exception
            SalesYesterdayLastWeekVariance = 0
        End Try

        '     SalesYesterdayLastWeekVariance = SalesYesterday / SalesYesterdayLastWeek
        '      records = erpdb.Execute("Select * from [KPI Values] where [KPI Name] ='POS Sales LMTD - Total'")

        StartDateFrom01 = Now.AddDays((DemoParameter) * -1)
        StartDateFrom01 = StartDateFrom01.AddMonths(-1)
        StartDateFrom01 = StartDateFrom01.AddDays((Now.Day * -1) + 1)

        StartDateTo01 = Now.AddDays((DemoParameter) * -1)
        StartDateTo01 = StartDateTo01.AddDays(-1)
        StartDateTo01 = StartDateTo01.AddMonths(-1)
        'if the first then last month is whole month
        'lmtd is month -2
        If Now.Day = 1 Then
            StartDateFrom01 = Now.AddDays((DemoParameter) * -1)
            StartDateFrom01 = StartDateFrom01.AddMonths(-2)
            StartDateTo01 = Now.AddDays((DemoParameter) * -1)
            StartDateTo01 = StartDateTo01.AddMonths(-1)
            StartDateTo01 = StartDateTo01.AddDays(-1)

        End If

        '        records = erpdb.Execute("Select * from [KPI Values] where [KPI Name] ='POS Sales LMTD - Total'")

        records = erpdb.Execute("Select sum([Total Net Sales]) As Sumsales from [POSSales] where cast([Date Record] as date) >= '" & Format(StartDateFrom01, "dd-MMM-yyyy") &
                                "' and cast([Date Record] as date) <= '" & Format(StartDateTo01, "dd-MMM-yyyy") & "'")
        SQLStatement = "Select sum([Total Net Sales]) As Sumsales from [POSSales] where cast([Date Record] as date) >= '" & Format(StartDateFrom01, "dd-MMM-yyyy") &
                                "' and cast([Date Record] as date) <= '" & Format(StartDateTo01, "dd-MMM-yyyy") & "'"

        SalesLastMonthToDate = 0
        While Not records.EOF
            Try
                SalesLastMonthToDate = records.Fields("Sumsales").Value
            Catch ex As Exception
                SalesLastMonthToDate = 0
            End Try
            records.MoveNext()
        End While
        records.Close()



        '    Format((ResultsArray(countertemp, 1) - ResultsArray(countertemp, 3)) / ResultsArray(countertemp, 3) * 100, "0.##")
        Try
            SalesLastMonthVariance = (SalesMTDYesterday - SalesLastMonthToDate) / SalesLastMonthToDate * 100
        Catch ex As Exception
            SalesLastMonthVariance = 0
        End Try



        'Yesterday last week
        '      DailyReportHTMLEmailString = "      <div style = ""width:400px;"">
        '         <h1 style=""text-align:center""> <background-color: blue;>Testing </h1>
        '        <p  style =""text-align:right""><a href=""#"">sample link</a> </p>
        '       </div>"

        'puts the bizcoach logo
        DailyReportHTMLEmailString = "    <img src = ""http://aagilitycom.ipage.com/wp-content/uploads/2017/12/BizCoach_small.png"" alt=""BizCoach"" ;  ><br /><br />"

        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "Group Daily Sales for : " & Format(DailyReportDate, "ddd dd MMM yyyy") & "<br /><br />" &
            "Yesterday: " & Format(SalesYesterday, "HK$ #,###,##0") & ", same day last week : " & Format(SalesYesterdayLastWeek, "HK$ #,###,##0") & ", variance to last week " & Format(SalesYesterdayLastWeekVariance, "0.##") & "%" & "<br />" &
            "MTD      : " & Format(SalesMTDYesterday, "HK$ #,###,##0") & ", last month to date " & "" & Format(SalesLastMonthToDate, "HK$ #,###,##0") & ", variance to last month " & Format(SalesLastMonthVariance, "0.##") & "%" & "<br />" &
                                                     "<br />" & "<br />" &
                                                     "<table border=1><col width=""300""><col width=""150""><col width=""150""><tr>" &
                                                     "<td><b>Outlet Sales</td>" &
                                                     "<td><b><p  style =""text-align:center"">Yesterday HK$</td>" &
                                                     "<td><b><p  style =""text-align:center"">MTD   HK$</b></td>" &
                                                     "</b></tr>"
        'now gets the individual outlets
        'now sums up
        '     records = erpdb.Execute("Select sum([KPI Value Number]) as Sumsales, [Outlet Name] from [KPI Values] where [KPI Name] ='POS Daily Sales (today)' group by [Outlet Name] order by [Outlet Name]")

        records = erpdb.Execute("select sum([Total Net Sales]) as Sumsales , [Outlet Name] from [POSSales] where cast([Date Record] as date) = '" &
                                Format(DailyReportDate, "dd-MMM-yyyy") & "' group by [Outlet Name] order by [Outlet Name]")

        'MTD
        StartDateFrom01 = Now.AddDays((DemoParameter) * -1)
        StartDateFrom01 = StartDateFrom01.AddDays((Now.Day * -1) + 1)
        StartDateTo01 = Now.AddDays((DemoParameter) * -1)
        StartDateTo01 = StartDateTo01.AddDays(-1)
        'if the first then last month is whole month
        If Now.Day = 1 Then
            StartDateFrom01 = Now.AddDays((DemoParameter) * -1)
            StartDateFrom01 = StartDateFrom01.AddMonths(-1)
            StartDateTo01 = Now.AddDays((DemoParameter) * -1)
            StartDateTo01 = StartDateTo01.AddDays(-1)
        End If

        SalesYesterday = 0
        While Not records.EOF
            SalesYesterday = 0
            Try
                OutletName = records.Fields("Outlet Name").Value
                SalesYesterday = records.Fields("Sumsales").Value
            Catch ex As Exception
                MsgBox("error " & Err.Description)
            End Try
            'now needs to get the MTD for the outlet

            '         recordsLookup = erpdb.Execute("Select * from [KPI Values] where [KPI Name] ='POS Sales MTD by Dimensions' and [Outlet Name] ='" & Replace(OutletName, "'", "''") & "'")
            SQLStatement = "Select sum([Total Net Sales]) As Sumsales, [Outlet Name] from [POSSales] where cast([Date Record] As Date) >= '" & Format(StartDateFrom01, "dd-MMM-yyyy") &
                                          "' and cast([Date Record] as date) <= '" & Format(StartDateTo01, "dd-MMM-yyyy") & "' and [Outlet Name] ='" & Replace(OutletName, "'", "''") & "' group by [Outlet Name] "

            recordsLookup = erpdb.Execute("select sum([Total Net Sales]) as Sumsales, [Outlet Name] from [POSSales] where cast([Date Record] as date) >= '" & Format(StartDateFrom01, "dd-MMM-yyyy") &
                                          "' and cast([Date Record] as date) <= '" & Format(StartDateTo01, "dd-MMM-yyyy") & "' and [Outlet Name] ='" & Replace(OutletName, "'", "''") & "' group by [Outlet Name] ")

            SalesMTDYesterday = 0
            While Not recordsLookup.EOF
                Try
                    SalesMTDYesterday = recordsLookup.Fields("Sumsales").Value
                Catch ex As Exception
                    SalesMTDYesterday = 0
                End Try

                recordsLookup.MoveNext()
            End While
            recordsLookup.Close()
            'can build the line
            DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<tr><td>" & OutletName & "</td>" &
                                         "<td>" & "<p  style =""text-align:right"">" & Format(SalesYesterday, "#,###,###") & "</td>" &
                                        "<td>" & "<p  style =""text-align:right"">" & Format(SalesMTDYesterday, "#,###,###") & "</tr>"


            records.MoveNext()
        End While
        records.Close()
        erpdb.Close()

        'closes the table
        DailyReportHTMLEmailString = "" & DailyReportHTMLEmailString & "</table></span><br /><hr /><br />"
        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<br />"


    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub


    Public Sub DailySalesReportPOSYesterdayDesktop(ByRef DailyReportHTMLEmailString As String, ByRef DailyReportDate As Date,
                                                   ByVal DemoCompany As Boolean, ByRef UrlOfReport As String,
                                                   ByRef DesktopText As String, ByRef Companyname As String,
                                                   ByRef SMSText As String, ByRef SpeechString As String)

        'this is a daily report that is based upon yesterdays sales
        'it compares this to the monthly and last month figures

        'as fixed date report then can take off the date here
        DailyReportDate = Now.AddDays(-1)



        'dont use KPIs now as may be a timing issue

        Dim erpdb As ADODB.Connection
        Dim records As ADODB._Recordset
        Dim recordsLookup As ADODB._Recordset
        Dim RecordsOutletLookup As ADODB._Recordset
        erpdb = New ADODB.Connection
        erpdb.Open(strconnectSQLBizCoachCustomerDataSQL)
        Dim SalesYesterday As Integer = 0
        Dim SalesMTDYesterday As Integer = 0
        Dim OutletName As String = ""
        Dim SalesYesterdayLastWeek As Integer = 0
        Dim SalesLastMonthVariance As Decimal = 0
        Dim SalesYesterdayLastWeekVariance As Decimal = 0
        Dim SalesLastMonthToDate As Integer = 0
        Dim StartDateFrom01 As Date = Now.ToShortDateString
        Dim StartDateTo01 As Date = Now.ToShortDateString
        Dim SQLStatement As String = ""
        Dim OutletFullNameTemp As String = ""
        Dim counter01 As Int16 = 0
        'this stores the excel data
        Dim ExcelTableData(10000, 10)

        'used for demo companies to go back 1 year
        SMSText = ""
        SpeechString = "Yesterdays revenue is "

        Dim DemoParameter As Int16 = 0
        'uses UTC time so need to take away 8 hours or use HongKongLoginMoment

        If DemoCompany = True Then
            DemoParameter = 365
            DailyReportDate = DailyReportDate.AddYears(-1)
        End If
        DailyReportDate = DailyReportDate.ToShortDateString

        '     records = erpdb.Execute("Select * from [KPI Values] where [KPI Name] ='POS Daily Sales (yesterday) - Total'")

        SQLStatement = "Select sum([Total Net Sales]) as Sumsales from [POSSales] where cast([Date Record] as date) = '" & Format(DailyReportDate, "dd-MMM-yyyy") & "'"
        records = erpdb.Execute(SQLStatement)

        SalesYesterday = 0
        While Not records.EOF
            SalesYesterday = 0
            Try
                SalesYesterday = records.Fields("Sumsales").Value
            Catch ex As Exception
                SalesYesterday = 0
            End Try
            records.MoveNext()
        End While
        records.Close()

        SpeechString = SpeechString & Format(SalesYesterday, "#,###") & "$" & " for your business. "

        StartDateFrom01 = Now.AddDays((DemoParameter) * -1)
        StartDateFrom01 = StartDateFrom01.AddDays((Now.Day * -1) + 1)
        StartDateTo01 = Now.AddDays((DemoParameter) * -1)
        StartDateTo01 = StartDateTo01.AddDays(-1)
        'if the first then last month is whole month
        If Now.Day = 1 Then
            StartDateFrom01 = Now.AddDays((DemoParameter) * -1)
            StartDateFrom01 = StartDateFrom01.AddMonths(-1)
            StartDateTo01 = Now.AddDays((DemoParameter) * -1)
            StartDateTo01 = StartDateTo01.AddDays(-1)
        End If

        '        records = erpdb.Execute("Select * from [KPI Values] where [KPI Name] ='POS Sales MTD - Total'")

        records = erpdb.Execute(" select sum([Total Net Sales]) as Sumsales from [POSSales] where cast([Date Record] as date) >= '" &
                                Format(StartDateFrom01, "dd-MMM-yyyy") &
                                "' and cast([Date Record] as date) <= '" & Format(StartDateTo01, "dd-MMM-yyyy") & "'")
        SalesMTDYesterday = 0
        While Not records.EOF
            Try
                SalesMTDYesterday = records.Fields("Sumsales").Value
            Catch ex As Exception
                SalesMTDYesterday = 0
            End Try
            records.MoveNext()
        End While
        records.Close()

        '      records = erpdb.Execute("Select * from [KPI Values] where [KPI Name] ='POS Daily Sales (yesterday same day last week) - Total'")

        records = erpdb.Execute("select sum([Total Net Sales]) as Sumsales from [POSSales] where cast([Date Record] as date) = '" & Format(DailyReportDate.AddDays(-7), "dd-MMM-yyyy") & "'")

        SalesYesterdayLastWeek = 0
        While Not records.EOF
            Try
                SalesYesterdayLastWeek = records.Fields("Sumsales").Value
            Catch ex As Exception
                SalesYesterdayLastWeek = 0
            End Try
            records.MoveNext()
        End While
        records.Close()

        Try
            SalesYesterdayLastWeekVariance = ((SalesYesterday - SalesYesterdayLastWeek)) / SalesYesterdayLastWeek * 100
        Catch ex As Exception
            SalesYesterdayLastWeekVariance = 0
        End Try

        '     SalesYesterdayLastWeekVariance = SalesYesterday / SalesYesterdayLastWeek
        '      records = erpdb.Execute("Select * from [KPI Values] where [KPI Name] ='POS Sales LMTD - Total'")

        StartDateFrom01 = Now.AddDays((DemoParameter) * -1)
        StartDateFrom01 = StartDateFrom01.AddMonths(-1)
        StartDateFrom01 = StartDateFrom01.AddDays((Now.Day * -1) + 1)

        StartDateTo01 = Now.AddDays((DemoParameter) * -1)
        StartDateTo01 = StartDateTo01.AddDays(-1)
        StartDateTo01 = StartDateTo01.AddMonths(-1)
        'if the first then last month is whole month
        'lmtd is month -2
        If Now.Day = 1 Then
            StartDateFrom01 = Now.AddDays((DemoParameter) * -1)
            StartDateFrom01 = StartDateFrom01.AddMonths(-2)
            StartDateTo01 = Now.AddDays((DemoParameter) * -1)
            StartDateTo01 = StartDateTo01.AddMonths(-1)
            StartDateTo01 = StartDateTo01.AddDays(-1)

        End If
        '        records = erpdb.Execute("Select * from [KPI Values] where [KPI Name] ='POS Sales LMTD - Total'")
        SQLStatement = "Select sum([Total Net Sales]) As Sumsales from [POSSales] where cast([Date Record] as date) >= '" & Format(StartDateFrom01, "dd-MMM-yyyy") &
                                "' and cast([Date Record] as date) <= '" & Format(StartDateTo01, "dd-MMM-yyyy") & "'"
        records = erpdb.Execute("Select sum([Total Net Sales]) As Sumsales from [POSSales] where cast([Date Record] as date) >= '" & Format(StartDateFrom01, "dd-MMM-yyyy") &
                                "' and cast([Date Record] as date) <= '" & Format(StartDateTo01, "dd-MMM-yyyy") & "'")


        SalesLastMonthToDate = 0
        While Not records.EOF
            Try
                SalesLastMonthToDate = records.Fields("Sumsales").Value
            Catch ex As Exception
                SalesLastMonthToDate = 0
            End Try
            records.MoveNext()
        End While
        records.Close()

        '    Format((ResultsArray(countertemp, 1) - ResultsArray(countertemp, 3)) / ResultsArray(countertemp, 3) * 100, "0.##")
        Try
            SalesLastMonthVariance = (SalesMTDYesterday - SalesLastMonthToDate) / SalesLastMonthToDate * 100
        Catch ex As Exception
            SalesLastMonthVariance = 0
        End Try


        'Yesterday last week
        '      DailyReportHTMLEmailString = "      <div style = ""width:400px;"">
        '         <h1 style=""text-align:center""> <background-color: blue;>Testing </h1>
        '        <p  style =""text-align:right""><a href=""#"">sample link</a> </p>
        '       </div>"
        DailyReportHTMLEmailString = "<!DOCTYPE html><html>" & "          <img src = ""http://aagilitycom.ipage.com/wp-content/uploads/2017/12/BizCoach_small.png""alt=""BizCoach"";><br /><br />"

        '        DailyReportHTMLEmailString = "<!DOCTYPE html><html>" & "          <br />"



        '        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<font face=""calibri""><b><u> TEST EMAIL - NOT FOR CUSTOMERS - " & FrmMain.ListBoxCustomerInfo.SelectedItem &
        '          " - " & FrmMain.ListBoxNotificationType.SelectedItem & "</b></u><br /><br />"

        Dim ExcelHeaderLine0401 As String = "Yesterday: "
        Dim ExcelHeaderLine0402 As String = Format(SalesYesterday, "HK$ #,###,##0")
        Dim ExcelHeaderLine0403 As String = "Same day last week : "
        Dim ExcelHeaderLine0404 As String = Format(SalesYesterdayLastWeek, "HK$ #,###,##0")
        Dim ExcelHeaderLine0405 As String = "Variance to last week "
        Dim ExcelHeaderLine0406 As String = Format(SalesYesterdayLastWeekVariance, "0.##") & "%"

        Dim ExcelHeaderLine0501 As String = "MTD      : "
        Dim ExcelHeaderLine0502 As String = Format(SalesMTDYesterday, "HK$ #,###,##0")
        Dim ExcelHeaderLine0503 As String = "Last month to date "
        Dim ExcelHeaderLine0504 As String = Format(SalesLastMonthToDate, "HK$ #,###,##0")
        Dim ExcelHeaderLine0505 As String = "Variance to last month "
        Dim ExcelHeaderLine0506 As String =  Format(SalesLastMonthVariance, "0.##") & "%"


        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "Your daily revenue snapshot as of " & Format(DailyReportDate, "ddd d MMM yyyy") & " in HK$<br /><br />" &
                                        "Yesterday: " & Format(SalesYesterday, "$ #,###,##0") & ", same day last week : " & Format(SalesYesterdayLastWeek, "$ #,###,##0") &
                                        ", variance to last week " & Format(SalesYesterdayLastWeekVariance, "0.##") & "%" & "<br />" &
                                        "MTD      : " & Format(SalesMTDYesterday, "$ #,###,##0") & ", last month to date " & "" & Format(SalesLastMonthToDate, "$ #,###,##0") &
                                        ", variance to last month " & Format(SalesLastMonthVariance, "0.##") & "%" & "<br />" &
                                        "<br />" & "<br />" &
                                        "<table border=1><col width=""400""><col width=""400""><col width=""400""><tr>" &
                                        "<td><b>Outlet Revenue</td>" &
                                        "<td><b><p style =""text-align:center"">Yesterday</td>" &
                                        "<td><b><p style =""text-align:center"">MTD   </b></td>" &
                                        "</b></tr>"

        SMSText = "BizCoach daily revenue snapshot as of " & Format(DailyReportDate, "ddd d MMM yyyy") & ControlChars.CrLf

        'now gets the individual outlets
        'now sums up
        '     records = erpdb.Execute("Select sum([KPI Value Number]) As Sumsales, [Outlet Name] from [KPI Values] where [KPI Name] ='POS Daily Sales (today)' group by [Outlet Name] order by [Outlet Name]")


        'MTD
        StartDateFrom01 = Now.AddDays((DemoParameter) * -1)
        StartDateFrom01 = StartDateFrom01.AddDays((Now.Day * -1) + 1)
        StartDateTo01 = Now.AddDays((DemoParameter) * -1)
        StartDateTo01 = StartDateTo01.AddDays(-1)
        'if the first then last month is whole month
        If Now.Day = 1 Then
            StartDateFrom01 = Now.AddDays((DemoParameter) * -1)
            StartDateFrom01 = StartDateFrom01.AddMonths(-1)
            StartDateTo01 = Now.AddDays((DemoParameter) * -1)
            StartDateTo01 = StartDateTo01.AddDays(-1)
        End If

        Dim SalesYesterdayAll As Decimal = 0


        'does all outlets
        SQLStatement = "select * from [Outlet Details] where [Date Closed] is Null order by [Outlet Name]"



        records = erpdb.Execute(SQLStatement)
        SalesYesterday = 0
        While Not records.EOF

            OutletName = records.Fields("Outlet Name").Value

            SQLStatement = "select sum([Total Net Sales]) as Sumsales , [Outlet Name] from [POSSales] where cast([Date Record] as date) = '" &
                                Format(DailyReportDate, "dd-MMM-yyyy") & "' and [Outlet Name] ='" & Replace(OutletName, "'", "''") & "' group by [Outlet Name] order by [Outlet Name]"
            SalesYesterday = 0
          
            recordsLookup = erpdb.Execute(SQLStatement)
            While Not recordsLookup.EOF
                Try
                    OutletName = recordsLookup.Fields("Outlet Name").Value
                    SalesYesterday = recordsLookup.Fields("Sumsales").Value
                    SalesYesterdayAll = SalesYesterdayAll + SalesYesterday
                Catch ex As Exception
                    MsgBox("error " & Err.Description)
                End Try

                recordsLookup.MoveNext()
            End While
            recordsLookup.Close()

            'now needs to get the MTD for the outlet

            '         recordsLookup = erpdb.Execute("Select * from [KPI Values] where [KPI Name] ='POS Sales MTD by Dimensions' and [Outlet Name] ='" & Replace(OutletName, "'", "''") & "'")
            SQLStatement = "Select sum([Total Net Sales]) As Sumsales, [Outlet Name] from [POSSales] where cast([Date Record] As Date) >= '" & Format(StartDateFrom01, "dd-MMM-yyyy") &
                                          "' and cast([Date Record] as date) <= '" & Format(StartDateTo01, "dd-MMM-yyyy") &
                                          "' and [Outlet Name] ='" & Replace(OutletName, "'", "''") & "' group by [Outlet Name] "

            recordsLookup = erpdb.Execute("select sum([Total Net Sales]) as Sumsales, [Outlet Name] from [POSSales] where cast([Date Record] as date) >= '" & Format(StartDateFrom01, "dd-MMM-yyyy") &
                                          "' and cast([Date Record] as date) <= '" & Format(StartDateTo01, "dd-MMM-yyyy") &
                                          "' and [Outlet Name] ='" & Replace(OutletName, "'", "''") & "' group by [Outlet Name] ")

            SalesMTDYesterday = 0
            While Not recordsLookup.EOF
                Try
                    SalesMTDYesterday = recordsLookup.Fields("Sumsales").Value
                Catch ex As Exception
                    SalesMTDYesterday = 0
                End Try

                recordsLookup.MoveNext()
            End While
            recordsLookup.Close()
            'can build the line
            DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<tr><td><p style =""text-align:left"">" & OutletName & "</td>" &
                                         "<td>" & "<p style =""text-align:right"">" & Format(SalesYesterday, "#,###,##0") & "</td>" &
                                        "<td>" & "<p style =""text-align:right"">" & Format(SalesMTDYesterday, "#,###,##0") & "</tr>"




            SMSText = SMSText & OutletName & " $ " & Format(SalesYesterday, "#,###,##0") & " " & ControlChars.CrLf

            'looks up full outlet name if available
            RecordsOutletLookup = erpdb.Execute("select * from [Outlet Details] where [Outlet Name]='" & Replace(OutletName, "'", "''") & "'")
            OutletFullNameTemp = ""
            While Not RecordsOutletLookup.EOF
                Try
                    OutletFullNameTemp = RecordsOutletLookup.Fields("Description").Value
                Catch ex As Exception
                    OutletFullNameTemp = ""
                End Try

                RecordsOutletLookup.MoveNext()
            End While
            RecordsOutletLookup.Close()
            If OutletFullNameTemp = "" Then
                OutletFullNameTemp = OutletName
            End If

            If counter01 = 0 Then
                SpeechString = SpeechString & " " & OutletFullNameTemp & "'s revenue was " & Format(SalesYesterday, "#,##0") & "$. "
            Else
                SpeechString = SpeechString & " " & OutletFullNameTemp & " " & Format(SalesYesterday, "#,##0") & "$. "
            End If

            counter01 += 1

            records.MoveNext()
        End While
        records.Close()
        erpdb.Close()


        'puts the bizcoach logo
        'this is desktop format so just a text string
        '"Yesterday: HK$ 1,014, last wk: HK$ 10,793"
        'http://aagilitycom.ipage.com/custview/20Jan1802glorystores.htm
        Dim CompanynameTemp As String = Companyname.Trim
        CompanynameTemp = Replace(CompanynameTemp, ".", "")
        CompanynameTemp = Replace(CompanynameTemp, " ", "")
        CompanynameTemp = Replace(CompanynameTemp, " ", "")
        CompanynameTemp = Replace(CompanynameTemp, "'", "")
        CompanynameTemp = CompanynameTemp.ToLower

        SMSText = SMSText & ControlChars.CrLf & "Total group $ " & Format(SalesYesterdayAll, "#,###,##0")
        SMSText = SMSText & ControlChars.CrLf & "Check email for more details and revenue predictions or login to BizCoach"

        UrlOfReport = "http://aagilitycom.ipage.com/custview/" & Format(DailyReportDate, "ddMMMyyyy") & "02" & CompanynameTemp & ".htm"
        DesktopText = "Yesterday: " & Format(SalesYesterdayAll, "HK$ #,###,##0") & ", last wk:" & Format(SalesYesterdayLastWeek, "HK$ #,###,##0")

        FrmMain.TextBoxSpeechText.Text = SpeechString.ToString.Trim

        'this creates the speech file
        AWSPollySpeak(SpeechString, "DailySalesPOS")
        'this publishes the speech to the internet 
        PublishSpeechtoImpala(SpeechString, "DailySalesPOS")
        'the file is bublished to    
        'http://aagilitycom.ipage.com/custview/    TextBoxPublishLocationName.text


        'closes the table
        DailyReportHTMLEmailString = "" & DailyReportHTMLEmailString & "</table><br /><br />"

        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "To see your dashboards login to BizCoach <a href=""https://app.impalabizcoach.com/"">here</a><br /><br />"

        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "To listen to your revenue snapshot click <a href=""http://aagilitycom.ipage.com/custview/" &
            FrmMain.TextBoxPublishLocationName.Text.Trim & """>here</a><br />"
        DailyReportHTMLEmailString = "" & DailyReportHTMLEmailString & "</table></span><br /><hr /><br />"
        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "Let us know how we are doing :<br /><br />"
        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "2 minute  <a href=""https://www.surveymonkey.com/r/M2Q557H"">survey</a><br />"
        DailyReportHTMLEmailString = DailyReportHTMLEmailString & " or<br />"
        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "email : <a href=""mailto:support@impalacloud.com"" target=""_top"">support@impalacloud.com</a><br /><br />"

        'survey is https://www.surveymonkey.com/r/M2Q557H 


        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "</html>"

        '      DailyReportHTMLEmailString = DailyReportHTMLEmailString & "Please give feedback To <a href=""mailto:support@impalacloud.com"" target=""_top"">support@impalacloud.com</a><br /><br />"

        '       DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<br />"

        'this creates the excel
        If FrmMain.CheckBoxCreateExcel.Checked = True Then
            CreateExcelSpreadsheet(DailyReportDate, ExcelHeaderLine0401,
                                      ExcelHeaderLine0402,
        ExcelHeaderLine0403,
        ExcelHeaderLine0404,
        ExcelHeaderLine0405,
        ExcelHeaderLine0406,
        ExcelHeaderLine0501,
        ExcelHeaderLine0502,
        ExcelHeaderLine0503,
        ExcelHeaderLine0504,
        ExcelHeaderLine0505,
        ExcelHeaderLine0506)
        End If


        'publish to the web browser on the notifications page
        Dim document As System.Windows.Forms.HtmlDocument = FrmMain.WebBrowserNotification.Document
        FrmMain.WebBrowserNotification.DocumentText = DailyReportHTMLEmailString


    End Sub



    Public Sub PredictionsPOSSystemMonthly(ByRef DailyReportHTMLEmailString As String, ByRef DailyReportDate As Date,
                                                   ByVal DemoCompany As Boolean, ByRef UrlOfReport As String,
                                                   ByRef DesktopText As String, ByRef Companyname As String,
                                                   ByRef SMSText As String, ByRef SpeechString As String)

        'this is a daily report that is based upon yesterdays sales
        'it compares this to the monthly and last month figures

        'as fixed date report then can take off the date here
        DailyReportDate = Now.AddDays(-1)



        'dont use KPIs now as may be a timing issue

        Dim erpdb As ADODB.Connection
        Dim records As ADODB._Recordset
        Dim recordsLookup As ADODB._Recordset
        Dim RecordsOutletLookup As ADODB._Recordset
        erpdb = New ADODB.Connection
        erpdb.Open(strconnectSQLBizCoachCustomerDataSQL)
        Dim SalesYesterday As Integer = 0
        Dim SalesMTDYesterday As Integer = 0
        Dim OutletName As String = ""
        Dim SalesYesterdayLastWeek As Integer = 0
        Dim SalesLastMonthVariance As Decimal = 0
        Dim SalesYesterdayLastWeekVariance As Decimal = 0
        Dim SalesLastMonthToDate As Integer = 0
        Dim StartDateFrom01 As Date = Now.ToShortDateString
        Dim StartDateTo01 As Date = Now.ToShortDateString
        Dim SQLStatement As String = ""
        Dim OutletFullNameTemp As String = ""
        Dim counter01 As Int16 = 0
        'this stores the excel data
        Dim ExcelTableData(10000, 10)

        'used for demo companies to go back 1 year
        SMSText = ""
        SpeechString = "Actual sales yesterday are "

        Dim DemoParameter As Int16 = 0
        'uses UTC time so need to take away 8 hours or use HongKongLoginMoment

        If DemoCompany = True Then
            DemoParameter = 365
            DailyReportDate = DailyReportDate.AddYears(-1)
        End If
        DailyReportDate = DailyReportDate.ToShortDateString

        '     records = erpdb.Execute("Select * from [KPI Values] where [KPI Name] ='POS Daily Sales (yesterday) - Total'")

        SQLStatement = "Select sum([Total Net Sales]) as Sumsales from [POSSales] where cast([Date Record] as date) = '" & Format(DailyReportDate, "dd-MMM-yyyy") & _
            "'"
        records = erpdb.Execute(SQLStatement)

        SalesYesterday = 0
        While Not records.EOF
            Try
                SalesYesterday = records.Fields("Sumsales").Value
            Catch ex As Exception
                SalesYesterday = 0
            End Try
            records.MoveNext()
        End While
        records.Close()

        SpeechString = SpeechString & Format(SalesYesterday, "#,###") & "$" & " for your business. "

        StartDateFrom01 = Now.AddDays((DemoParameter) * -1)
        StartDateFrom01 = StartDateFrom01.AddDays((Now.Day * -1) + 1)
        StartDateTo01 = Now.AddDays((DemoParameter) * -1)
        StartDateTo01 = StartDateTo01.AddDays(-1)
        'if the first then last month is whole month
        If Now.Day = 1 Then
            StartDateFrom01 = Now.AddDays((DemoParameter) * -1)
            StartDateFrom01 = StartDateFrom01.AddMonths(-1)
            StartDateTo01 = Now.AddDays((DemoParameter) * -1)
            StartDateTo01 = StartDateTo01.AddDays(-1)
        End If

        '        records = erpdb.Execute("Select * from [KPI Values] where [KPI Name] ='POS Sales MTD - Total'")

        records = erpdb.Execute(" select sum([Total Net Sales]) as Sumsales from [POSSales] where cast([Date Record] as date) >= '" &
                                Format(StartDateFrom01, "dd-MMM-yyyy") &
                                "' and cast([Date Record] as date) <= '" & Format(StartDateTo01, "dd-MMM-yyyy") & "'")
        SalesMTDYesterday = 0
        While Not records.EOF
            Try
                SalesMTDYesterday = records.Fields("Sumsales").Value
            Catch ex As Exception
                SalesMTDYesterday = 0
            End Try
            records.MoveNext()
        End While
        records.Close()

        '      records = erpdb.Execute("Select * from [KPI Values] where [KPI Name] ='POS Daily Sales (yesterday same day last week) - Total'")

        records = erpdb.Execute("select sum([Total Net Sales]) as Sumsales from [POSSales] where cast([Date Record] as date) = '" & Format(DailyReportDate.AddDays(-7), "dd-MMM-yyyy") & "'")

        SalesYesterdayLastWeek = 0
        While Not records.EOF
            Try
                SalesYesterdayLastWeek = records.Fields("Sumsales").Value
            Catch ex As Exception
                SalesYesterdayLastWeek = 0
            End Try
            records.MoveNext()
        End While
        records.Close()

        Try
            SalesYesterdayLastWeekVariance = ((SalesYesterday - SalesYesterdayLastWeek)) / SalesYesterdayLastWeek * 100
        Catch ex As Exception
            SalesYesterdayLastWeekVariance = 0
        End Try

        '     SalesYesterdayLastWeekVariance = SalesYesterday / SalesYesterdayLastWeek
        '      records = erpdb.Execute("Select * from [KPI Values] where [KPI Name] ='POS Sales LMTD - Total'")

        StartDateFrom01 = Now.AddDays((DemoParameter) * -1)
        StartDateFrom01 = StartDateFrom01.AddMonths(-1)
        StartDateFrom01 = StartDateFrom01.AddDays((Now.Day * -1) + 1)

        StartDateTo01 = Now.AddDays((DemoParameter) * -1)
        StartDateTo01 = StartDateTo01.AddDays(-1)
        StartDateTo01 = StartDateTo01.AddMonths(-1)
        'if the first then last month is whole month
        'lmtd is month -2
        If Now.Day = 1 Then
            StartDateFrom01 = Now.AddDays((DemoParameter) * -1)
            StartDateFrom01 = StartDateFrom01.AddMonths(-2)
            StartDateTo01 = Now.AddDays((DemoParameter) * -1)
            StartDateTo01 = StartDateTo01.AddMonths(-1)
            StartDateTo01 = StartDateTo01.AddDays(-1)

        End If
        '        records = erpdb.Execute("Select * from [KPI Values] where [KPI Name] ='POS Sales LMTD - Total'")

        records = erpdb.Execute("Select sum([Total Net Sales]) As Sumsales from [POSSales] where cast([Date Record] as date) >= '" & Format(StartDateFrom01, "dd-MMM-yyyy") &
                                "' and cast([Date Record] as date) <= '" & Format(StartDateTo01, "dd-MMM-yyyy") & "'")
        SQLStatement = "Select sum([Total Net Sales]) As Sumsales from [POSSales] where cast([Date Record] as date) >= '" & Format(StartDateFrom01, "dd-MMM-yyyy") &
                                "' and cast([Date Record] as date) <= '" & Format(StartDateTo01, "dd-MMM-yyyy") & "'"

        SalesLastMonthToDate = 0
        While Not records.EOF
            Try
                SalesLastMonthToDate = records.Fields("Sumsales").Value
            Catch ex As Exception
                SalesLastMonthToDate = 0
            End Try
            records.MoveNext()
        End While
        records.Close()

        '    Format((ResultsArray(countertemp, 1) - ResultsArray(countertemp, 3)) / ResultsArray(countertemp, 3) * 100, "0.##")
        Try
            SalesLastMonthVariance = (SalesMTDYesterday - SalesLastMonthToDate) / SalesLastMonthToDate * 100
        Catch ex As Exception
            SalesLastMonthVariance = 0
        End Try


        'Yesterday last week
        '      DailyReportHTMLEmailString = "      <div style = ""width:400px;"">
        '         <h1 style=""text-align:center""> <background-color: blue;>Testing </h1>
        '        <p  style =""text-align:right""><a href=""#"">sample link</a> </p>
        '       </div>"http://aagilitycom.ipage.com/wp-content/uploads/2017/12/BizCoach_small.png
        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "    <img src = ""http://aagilitycom.ipage.com/wp-content/uploads/2017/12/BizCoach_small.png"" alt=""BizCoach"" ;  ><br /><br />"


        '        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<font face=""calibri""><b><u> TEST EMAIL - NOT FOR CUSTOMERS - " & FrmMain.ListBoxCustomerInfo.SelectedItem &
        '          " - " & FrmMain.ListBoxNotificationType.SelectedItem & "</b></u><br /><br />"


        Dim ExcelHeaderLine0401 As String = "Yesterday: "
        Dim ExcelHeaderLine0402 As String = Format(SalesYesterday, "HK$ #,###,##0")
        Dim ExcelHeaderLine0403 As String = "Same day last week : "
        Dim ExcelHeaderLine0404 As String = Format(SalesYesterdayLastWeek, "HK$ #,###,##0")
        Dim ExcelHeaderLine0405 As String = "Variance to last week "
        Dim ExcelHeaderLine0406 As String = Format(SalesYesterdayLastWeekVariance, "0.##") & "%"

        Dim ExcelHeaderLine0501 As String = "MTD      : "
        Dim ExcelHeaderLine0502 As String = Format(SalesMTDYesterday, "HK$ #,###,##0")
        Dim ExcelHeaderLine0503 As String = "Last month to date "
        Dim ExcelHeaderLine0504 As String = Format(SalesLastMonthToDate, "HK$ #,###,##0")
        Dim ExcelHeaderLine0505 As String = "Variance to last month "
        Dim ExcelHeaderLine0506 As String = Format(SalesLastMonthVariance, "0.##") & "%"




        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "Daily Sales Prediction For : " & Format(DailyReportDate, "ddd d MMM yyyy") & "<br /><br />" &
                                        "Actual sales yesterday: " & Format(SalesYesterday, "HK$ #,###,##0") & ", same day last week : " & Format(SalesYesterdayLastWeek, "HK$ #,###,##0") &
                                        ", variance to last week " & Format(SalesYesterdayLastWeekVariance, "0.##") & "%" & "<br />" &
                                        "MTD      : " & Format(SalesMTDYesterday, "HK$ #,###,##0") & ", last month to date " & "" & Format(SalesLastMonthToDate, "HK$ #,###,##0") &
                                        ", variance to last month " & Format(SalesLastMonthVariance, "0.##") & "%" & "<br />" &
                                        "<br />" & "<br />" &
                                        "<table border=1><col width=""300""><col width=""150""><col width=""100""><col width=""100""><col width=""100""><col width=""100""><col width=""100""><col width=""100""><col width=""100""><tr>" &
                                        "<td><b>Outlet Sales</td>" &
                                        "<td><b><p style =""text-align:center"">Yesterday HK$</td>" &
                                        "<td><b><p style =""text-align:center"">Prediction " & Format(DailyReportDate.AddDays(1), "ddd") & "</b></td>" &
                                        "<td><b><p style =""text-align:center"">" & Format(DailyReportDate.AddDays(2), "ddd") & "</b></td>" &
                                        "<td><b><p style =""text-align:center"">" & Format(DailyReportDate.AddDays(3), "ddd") & "</b></td>" &
                                        "<td><b><p style =""text-align:center"">" & Format(DailyReportDate.AddDays(4), "ddd") & "</b></td>" &
                                        "<td><b><p style =""text-align:center"">" & Format(DailyReportDate.AddDays(5), "ddd") & "</b></td>" &
                                        "<td><b><p style =""text-align:center"">" & Format(DailyReportDate.AddDays(6), "ddd") & "</b></td>" &
                                        "<td><b><p style =""text-align:center"">" & Format(DailyReportDate.AddDays(7), "ddd") & "</b></td>" &
                                        "</b></tr>"

        SMSText = "Prediction for " & Format(DailyReportDate.AddDays(1), "ddd d MMM") & ControlChars.CrLf

        'now gets the individual outlets
        'now sums up
        '     records = erpdb.Execute("Select sum([KPI Value Number]) As Sumsales, [Outlet Name] from [KPI Values] where [KPI Name] ='POS Daily Sales (today)' group by [Outlet Name] order by [Outlet Name]")

        Dim LatestPredictionsDate As Date = Now.AddYears(-10)
        'gets the latest prediction date
        records = erpdb.Execute("select * from [Predictions] order by [prediction_created_on] Desc")
        While Not records.EOF
            Try
                LatestPredictionsDate = records.Fields("prediction_created_on").Value
            Catch ex As Exception

            End Try

            Exit While
            records.MoveNext()
        End While
        records.Close()
        LatestPredictionsDate = LatestPredictionsDate.ToShortDateString




        'MTD
        StartDateFrom01 = Now.AddDays((DemoParameter) * -1)
        StartDateFrom01 = StartDateFrom01.AddDays((Now.Day * -1) + 1)
        StartDateTo01 = Now.AddDays((DemoParameter) * -1)
        StartDateTo01 = StartDateTo01.AddDays(-1)
        'if the first then last month is whole month
        If Now.Day = 1 Then
            StartDateFrom01 = Now.AddDays((DemoParameter) * -1)
            StartDateFrom01 = StartDateFrom01.AddMonths(-1)
            StartDateTo01 = Now.AddDays((DemoParameter) * -1)
            StartDateTo01 = StartDateTo01.AddDays(-1)
        End If


        Dim SalesYesterdayAll As Decimal = 0
        Dim SalesPredictionPlus1 As Decimal = 0
        Dim SalesPredictionPlus2 As Decimal = 0
        Dim SalesPredictionPlus3 As Decimal = 0
        Dim SalesPredictionPlus4 As Decimal = 0
        Dim SalesPredictionPlus5 As Decimal = 0
        Dim SalesPredictionPlus6 As Decimal = 0
        Dim SalesPredictionPlus7 As Decimal = 0
        Dim TotalPrediction7Days As Decimal = 0

        'this gets all the outlets even if they have no sales
        SQLStatement = "select * from [Outlet Details]  where [Date Closed] is Null order by [Outlet Name]"

        records = erpdb.Execute(SQLStatement)
        SalesYesterday = 0
        While Not records.EOF

            Try
                OutletName = records.Fields("Outlet Name").Value
                '    SalesYesterday = records.Fields("Sumsales").Value
                '    SalesYesterdayAll = SalesYesterdayAll + SalesYesterday
            Catch ex As Exception
                MsgBox("error " & Err.Description)
            End Try


            SQLStatement = "select sum([Total Net Sales]) as Sumsales , [Outlet Name] from [POSSales] where cast([Date Record] as date) = '" &
                                Format(DailyReportDate, "dd-MMM-yyyy") & "' and [Outlet Name] ='" & Replace(OutletName, "'", "''") & "' group by [Outlet Name] order by [Outlet Name]"
            recordsLookup = erpdb.Execute(SQLStatement)
            SalesYesterday = 0
            While Not recordsLookup.EOF
                SalesYesterday = 0
                Try
                    OutletName = recordsLookup.Fields("Outlet Name").Value
                    SalesYesterday = recordsLookup.Fields("Sumsales").Value
                    SalesYesterdayAll = SalesYesterdayAll + SalesYesterday
                Catch ex As Exception
                    MsgBox("error " & Err.Description)
                End Try
                recordsLookup.MoveNext()
            End While
            recordsLookup.Close()


            'day plus 1
            SQLStatement = "Select sum([predicted_value_numeric]) As Sumsales, [Outlet Name] from [Predictions] where cast([predicted_date] As Date) = '" & Format(DailyReportDate.AddDays(1), "dd-MMM-yyyy") &
                "' and cast([prediction_created_on] as date) = '" & Format(LatestPredictionsDate, "dd-MMM-yyyy") & "' 
                and [Outlet Name] ='" & Replace(OutletName, "'", "''") & "' and [predictions_name] ='Total Net Sales - best fit' group by [Outlet Name]"
            recordsLookup = erpdb.Execute(SQLStatement)
            SalesPredictionPlus1 = 0
            While Not recordsLookup.EOF
                Try
                    SalesPredictionPlus1 = recordsLookup.Fields("Sumsales").Value
                    TotalPrediction7Days = TotalPrediction7Days + SalesPredictionPlus1
                Catch ex As Exception
                    SalesPredictionPlus1 = 0
                End Try
                recordsLookup.MoveNext()
            End While
            recordsLookup.Close()
            SQLStatement = "Select sum([predicted_value_numeric]) As Sumsales, [Outlet Name] from [Predictions] where cast([predicted_date] As Date) = '" & Format(DailyReportDate.AddDays(2), "dd-MMM-yyyy") &
                "' and cast([prediction_created_on] as date) = '" & Format(LatestPredictionsDate, "dd-MMM-yyyy") & "' 
                and [Outlet Name] ='" & Replace(OutletName, "'", "''") & "' and [predictions_name] ='Total Net Sales - best fit' group by [Outlet Name]"
            recordsLookup = erpdb.Execute(SQLStatement)
            SalesPredictionPlus2 = 0
            While Not recordsLookup.EOF
                Try
                    SalesPredictionPlus2 = recordsLookup.Fields("Sumsales").Value
                    TotalPrediction7Days = TotalPrediction7Days + SalesPredictionPlus2
                Catch ex As Exception
                    SalesPredictionPlus2 = 0
                End Try
                recordsLookup.MoveNext()
            End While
            recordsLookup.Close()
            SQLStatement = "Select sum([predicted_value_numeric]) As Sumsales, [Outlet Name] from [Predictions] where cast([predicted_date] As Date) = '" & Format(DailyReportDate.AddDays(3), "dd-MMM-yyyy") &
                "' and cast([prediction_created_on] as date) = '" & Format(LatestPredictionsDate, "dd-MMM-yyyy") & "' 
                and [Outlet Name] ='" & Replace(OutletName, "'", "''") & "' and [predictions_name] ='Total Net Sales - best fit' group by [Outlet Name]"
            recordsLookup = erpdb.Execute(SQLStatement)
            SalesPredictionPlus3 = 0
            While Not recordsLookup.EOF
                Try
                    SalesPredictionPlus3 = recordsLookup.Fields("Sumsales").Value
                    TotalPrediction7Days = TotalPrediction7Days + SalesPredictionPlus3
                Catch ex As Exception
                    SalesPredictionPlus3 = 0
                End Try
                recordsLookup.MoveNext()
            End While
            recordsLookup.Close()
            SQLStatement = "Select sum([predicted_value_numeric]) As Sumsales, [Outlet Name] from [Predictions] where cast([predicted_date] As Date) = '" & Format(DailyReportDate.AddDays(4), "dd-MMM-yyyy") &
                "' and cast([prediction_created_on] as date) = '" & Format(LatestPredictionsDate, "dd-MMM-yyyy") & "' 
                and [Outlet Name] ='" & Replace(OutletName, "'", "''") & "' and [predictions_name] ='Total Net Sales - best fit' group by [Outlet Name]"
            recordsLookup = erpdb.Execute(SQLStatement)
            SalesPredictionPlus4 = 0
            While Not recordsLookup.EOF
                Try
                    SalesPredictionPlus4 = recordsLookup.Fields("Sumsales").Value
                    TotalPrediction7Days = TotalPrediction7Days + SalesPredictionPlus4
                Catch ex As Exception
                    SalesPredictionPlus4 = 0
                End Try
                recordsLookup.MoveNext()
            End While
            recordsLookup.Close()
            SQLStatement = "Select sum([predicted_value_numeric]) As Sumsales, [Outlet Name] from [Predictions] where cast([predicted_date] As Date) = '" & Format(DailyReportDate.AddDays(5), "dd-MMM-yyyy") &
                "' and cast([prediction_created_on] as date) = '" & Format(LatestPredictionsDate, "dd-MMM-yyyy") & "' 
                and [Outlet Name] ='" & Replace(OutletName, "'", "''") & "' and [predictions_name] ='Total Net Sales - best fit' group by [Outlet Name]"
            recordsLookup = erpdb.Execute(SQLStatement)
            SalesPredictionPlus5 = 0
            While Not recordsLookup.EOF
                Try
                    SalesPredictionPlus5 = recordsLookup.Fields("Sumsales").Value
                    TotalPrediction7Days = TotalPrediction7Days + SalesPredictionPlus5
                Catch ex As Exception
                    SalesPredictionPlus5 = 0
                End Try
                recordsLookup.MoveNext()
            End While
            recordsLookup.Close()
            SQLStatement = "Select sum([predicted_value_numeric]) As Sumsales, [Outlet Name] from [Predictions] where cast([predicted_date] As Date) = '" & Format(DailyReportDate.AddDays(6), "dd-MMM-yyyy") &
                "' and cast([prediction_created_on] as date) = '" & Format(LatestPredictionsDate, "dd-MMM-yyyy") & "' 
                and [Outlet Name] ='" & Replace(OutletName, "'", "''") & "' and [predictions_name] ='Total Net Sales - best fit' group by [Outlet Name]"
            recordsLookup = erpdb.Execute(SQLStatement)
            SalesPredictionPlus6 = 0
            While Not recordsLookup.EOF
                Try
                    SalesPredictionPlus6 = recordsLookup.Fields("Sumsales").Value
                    TotalPrediction7Days = TotalPrediction7Days + SalesPredictionPlus6
                Catch ex As Exception
                    SalesPredictionPlus6 = 0
                End Try
                recordsLookup.MoveNext()
            End While
            recordsLookup.Close()
            SQLStatement = "Select sum([predicted_value_numeric]) As Sumsales, [Outlet Name] from [Predictions] where cast([predicted_date] As Date) = '" & Format(DailyReportDate.AddDays(7), "dd-MMM-yyyy") &
                "' and cast([prediction_created_on] as date) = '" & Format(LatestPredictionsDate, "dd-MMM-yyyy") & "' 
                and [Outlet Name] ='" &  Replace(OutletName, "'", "''")  & "' and [predictions_name] ='Total Net Sales - best fit' group by [Outlet Name]"
            recordsLookup = erpdb.Execute(SQLStatement)
            SalesPredictionPlus7 = 0
            While Not recordsLookup.EOF
                Try
                    SalesPredictionPlus7 = recordsLookup.Fields("Sumsales").Value
                    TotalPrediction7Days = TotalPrediction7Days + SalesPredictionPlus7
                Catch ex As Exception
                    SalesPredictionPlus7 = 0
                End Try
                recordsLookup.MoveNext()
            End While
            recordsLookup.Close()

            'can now build the line
            DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<tr><td>" & OutletName & "</td>" &
                                          "<td>" & "<p style =""text-align:right"">" & Format(SalesYesterday, "#,###,##0") & "</td>" &
                                          "<td>" & "<p style =""text-align:right"">" & Format(SalesPredictionPlus1, "#,###,###") & "</td>" &
                                         "<td>" & "<p style =""text-align:right"">" & Format(SalesPredictionPlus2, "#,###,###") & "</td>" &
                                         "<td>" & "<p style =""text-align:right"">" & Format(SalesPredictionPlus3, "#,###,###") & "</td>" &
                                         "<td>" & "<p style =""text-align:right"">" & Format(SalesPredictionPlus4, "#,###,###") & "</td>" &
                                         "<td>" & "<p style =""text-align:right"">" & Format(SalesPredictionPlus5, "#,###,###") & "</td>" &
                                         "<td>" & "<p style =""text-align:right"">" & Format(SalesPredictionPlus6, "#,###,###") & "</td>" &
                                         "<td>" & "<p style =""text-align:right"">" & Format(SalesPredictionPlus7, "#,###,###") & "</tr>"


            SMSText = SMSText & OutletName & " " & Format(DailyReportDate.AddDays(1), "ddd") & " $ " & Format(SalesPredictionPlus1, "#,###,###") & " " & ControlChars.CrLf

            'looks up full outlet name if available
            RecordsOutletLookup = erpdb.Execute("select * from [Outlet Details] where [Outlet Name]='" & Replace(OutletName, "'", "''") & "'")
            OutletFullNameTemp = ""
            While Not RecordsOutletLookup.EOF
                Try
                    OutletFullNameTemp = RecordsOutletLookup.Fields("Description").Value
                Catch ex As Exception
                    OutletFullNameTemp = ""
                End Try

                RecordsOutletLookup.MoveNext()
            End While
            RecordsOutletLookup.Close()
            If OutletFullNameTemp = "" Then
                OutletFullNameTemp = OutletName
            End If

            If counter01 = 0 Then
                SpeechString = SpeechString & " " & OutletFullNameTemp & " sold " & Format(SalesYesterday, "#,###") & "$ " &
                    " yesterday and is predicted to sell " & Format(SalesPredictionPlus1, "#,###,###") & "$ on " & Format(DailyReportDate.AddDays(1), "dddd") &
                    " and " & Format(SalesPredictionPlus2, "#,###,###") & "$ on " & Format(DailyReportDate.AddDays(2), "dddd") &
                    " a total of " & Format(SalesPredictionPlus1 + SalesPredictionPlus2 + SalesPredictionPlus3 + SalesPredictionPlus4 + SalesPredictionPlus5 +
                    SalesPredictionPlus6 + SalesPredictionPlus7, "#,###,###") & "$ by " & Format(DailyReportDate.AddDays(7), "dddd") & ". "

            Else
                SpeechString = SpeechString & " " & OutletFullNameTemp & " sold " & Format(SalesYesterday, "#,###") & "$ " &
                    " yesterday and is predicted to sell " & Format(SalesPredictionPlus1, "#,###,###") & "$ on " & Format(DailyReportDate.AddDays(1), "dddd") &
                    " and " & Format(SalesPredictionPlus2, "#,###,###") & "$ on " & Format(DailyReportDate.AddDays(2), "dddd") &
                    " a total of " & Format(SalesPredictionPlus1 + SalesPredictionPlus2 + SalesPredictionPlus3 + SalesPredictionPlus4 + SalesPredictionPlus5 +
                    SalesPredictionPlus6 + SalesPredictionPlus7, "#,###,###") & "$ by " & Format(DailyReportDate.AddDays(7), "dddd") & ". "

            End If

            counter01 += 1

            records.MoveNext()
        End While
        records.Close()
        erpdb.Close()


        'puts the bizcoach logo
        'this is desktop format so just a text string
        '"Yesterday: HK$ 1,014, last wk: HK$ 10,793"
        'http://aagilitycom.ipage.com/custview/20Jan1802glorystores.htm
        Dim CompanynameTemp As String = Companyname.Trim
        CompanynameTemp = Replace(CompanynameTemp, ".", "")
        CompanynameTemp = Replace(CompanynameTemp, " ", "")
        CompanynameTemp = Replace(CompanynameTemp, " ", "")
        CompanynameTemp = Replace(CompanynameTemp, "'", "")
        CompanynameTemp = CompanynameTemp.ToLower

        SMSText = SMSText & ControlChars.CrLf & "Prediction $ " & Format(TotalPrediction7Days, "#,###,##0") & " to " & Format(DailyReportDate.AddDays(7), "ddd d MMM")

        UrlOfReport = "http://aagilitycom.ipage.com/custview/" & Format(DailyReportDate, "ddMMMyyyy") & "02" & CompanynameTemp & ".htm"
        DesktopText = "Yesterday: " & Format(SalesYesterdayAll, "HK$ #,###,##0") & ", last wk:" & Format(SalesYesterdayLastWeek, "HK$ #,###,##0")


        SpeechString = SpeechString & ". Total prediction for the next 7 days is " & Format(TotalPrediction7Days, "#,###") & "$ . "

        FrmMain.TextBoxSpeechText.Text = SpeechString.ToString.Trim

        'this creates the speech file

        AWSPollySpeak(SpeechString, "PredictionsPOSMonthly")
        'this publishes the speech to the internet 
        PublishSpeechtoImpala(SpeechString, "PredictionsPOSMonthly")
        'the file is bublished to    
        'http://aagilitycom.ipage.com/custview/    TextBoxPublishLocationName.text


        'closes the table
        DailyReportHTMLEmailString = "" & DailyReportHTMLEmailString & "</table></span><br /><hr /><br />"
        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<br />"

        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "If you would like BizCoach to read your email click <a href=""http://aagilitycom.ipage.com/custview/" &
            FrmMain.TextBoxPublishLocationName.Text.Trim & """>here</a><br /><br />"
        '


        If FrmMain.ListBoxNotificationType.SelectedItem = "Daily Sales Report (POS System) - MTD" Then

            DailyReportHTMLEmailString = DailyReportHTMLEmailString & "See your dashboard <a href=""http://aagilitycom.ipage.com/custview/" &
            "izpSv2OHRcCL-w9jcRWyVQ.html" & """>here</a><br /><br />"

            'https://litmus.com/blog/a-guide-to-bulletproof-buttons-in-email-design#supporttable

            DailyReportHTMLEmailString = DailyReportHTMLEmailString & "</font> <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0""> " &
                        " <tr>" &
                        "    <td> " &
                        " <table border=""0"" cellspacing=""0"" cellpadding=""0"">" &
                        "        <tr>" &
                        "          <td align=""center"" style=""border-radius 3px;"" bgcolor=""#e9703e""><a href=""https://www.surveymonkey.com/r/M2Q557H"" target=""_blank"" style=""font-size: 16px; font-family: Helvetica, Arial, sans-serif; color: #ffffff; text-decoration:none; text-decoration: none;border-radius: 3px; padding: 12px 18px; border: 1px solid #e9703e; display: inline-block;"">Take our survey &rarr;</a></td>" &
                        "        </tr>" &
                        "      </table>" &
                        "    </td>" &
                        "  </tr>" &
                        "</table><br />"


            '               DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<div id = ""your-view-container"" ></div>
            '               <script type = ""text/javascript"" >
            '               var reflect = New ReflectUI();
            '               var element = document.getElementById('your-view-container');
            '
            '               // Customize your view using the reflect object
            '               reflect.view('izpSv2OHRcCL-w9jcRWyVQ')
            '               .render(element);
            '               </script>"

        End If


        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "Please give feedback To <a href=""mailto:support@impalacloud.com"" target=""_top"">support@impalacloud.com</a><br /><br />"

        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<br />"

        'this creates the excel
        If FrmMain.CheckBoxCreateExcel.Checked = True Then
            CreateExcelSpreadsheet(DailyReportDate, ExcelHeaderLine0401,
                                      ExcelHeaderLine0402,
        ExcelHeaderLine0403,
        ExcelHeaderLine0404,
        ExcelHeaderLine0405,
        ExcelHeaderLine0406,
        ExcelHeaderLine0501,
        ExcelHeaderLine0502,
        ExcelHeaderLine0503,
        ExcelHeaderLine0504,
        ExcelHeaderLine0505,
        ExcelHeaderLine0506)
        End If


        'publish to the web browser on the notifications page
        Dim document As System.Windows.Forms.HtmlDocument = FrmMain.WebBrowserNotification.Document
        FrmMain.WebBrowserNotification.DocumentText = DailyReportHTMLEmailString


    End Sub



    Public Sub PredictionsPOSSystemDaily(ByRef DailyReportHTMLEmailString As String, ByRef DailyReportDate As Date,
                                                   ByVal DemoCompany As Boolean, ByRef UrlOfReport As String,
                                                   ByRef DesktopText As String, ByRef Companyname As String,
                                                   ByRef SMSText As String, ByRef SpeechString As String)

        'this is a daily report that is based upon yesterdays sales
        'it compares this to the monthly and last month figures

        'as fixed date report then can take off the date here
        DailyReportDate = Now.AddDays(-1)



        'dont use KPIs now as may be a timing issue

        Dim erpdb As ADODB.Connection
        Dim records As ADODB._Recordset
        Dim recordsLookup As ADODB._Recordset
        Dim RecordsOutletLookup As ADODB._Recordset
        erpdb = New ADODB.Connection
        erpdb.Open(strconnectSQLBizCoachCustomerDataSQL)
        Dim SalesYesterday As Integer = 0
        Dim SalesMTDYesterday As Integer = 0
        Dim OutletName As String = ""
        Dim SalesYesterdayLastWeek As Integer = 0
        Dim SalesLastMonthVariance As Decimal = 0
        Dim SalesYesterdayLastWeekVariance As Decimal = 0
        Dim SalesLastMonthToDate As Integer = 0
        Dim StartDateFrom01 As Date = Now.ToShortDateString
        Dim StartDateTo01 As Date = Now.ToShortDateString
        Dim SQLStatement As String = ""
        Dim OutletFullNameTemp As String = ""
        Dim counter01 As Int16 = 0
        'this stores the excel data
        Dim ExcelTableData(10000, 10)




        Dim DemoParameter As Int16 = 0
        'uses UTC time so need to take away 8 hours or use HongKongLoginMoment


        DailyReportDate = DailyReportDate.ToShortDateString

        '     records = erpdb.Execute("Select * from [KPI Values] where [KPI Name] ='POS Daily Sales (yesterday) - Total'")

        SQLStatement = "Select sum([Total Net Sales]) as Sumsales from [POSSales] where cast([Date Record] as date) = '" & Format(DailyReportDate, "dd-MMM-yyyy") &
            "'"
        records = erpdb.Execute(SQLStatement)

        SalesYesterday = 0
        While Not records.EOF
            Try
                SalesYesterday = records.Fields("Sumsales").Value
            Catch ex As Exception
                SalesYesterday = 0
            End Try
            records.MoveNext()
        End While
        records.Close()



        SMSText = ""
        SpeechString = ""
        SpeechString = SpeechString & Format(SalesYesterday, "#,###") & "$" & " for your business. "

        StartDateFrom01 = Now.AddDays((DemoParameter) * -1)
        StartDateFrom01 = StartDateFrom01.AddDays((Now.Day * -1) + 1)
        StartDateTo01 = Now.AddDays((DemoParameter) * -1)
        StartDateTo01 = StartDateTo01.AddDays(-1)
        'if the first then last month is whole month
        If Now.Day = 1 Then
            StartDateFrom01 = Now.AddDays((DemoParameter) * -1)
            StartDateFrom01 = StartDateFrom01.AddMonths(-1)
            StartDateTo01 = Now.AddDays((DemoParameter) * -1)
            StartDateTo01 = StartDateTo01.AddDays(-1)
        End If

        '        records = erpdb.Execute("Select * from [KPI Values] where [KPI Name] ='POS Sales MTD - Total'")

        records = erpdb.Execute(" select sum([Total Net Sales]) as Sumsales from [POSSales] where cast([Date Record] as date) >= '" &
                                Format(StartDateFrom01, "dd-MMM-yyyy") &
                                "' and cast([Date Record] as date) <= '" & Format(StartDateTo01, "dd-MMM-yyyy") & "'")
        SalesMTDYesterday = 0
        While Not records.EOF
            Try
                SalesMTDYesterday = records.Fields("Sumsales").Value
            Catch ex As Exception
                SalesMTDYesterday = 0
            End Try
            records.MoveNext()
        End While
        records.Close()

        '      records = erpdb.Execute("Select * from [KPI Values] where [KPI Name] ='POS Daily Sales (yesterday same day last week) - Total'")

        records = erpdb.Execute("select sum([Total Net Sales]) as Sumsales from [POSSales] where cast([Date Record] as date) = '" & Format(DailyReportDate.AddDays(-7), "dd-MMM-yyyy") & "'")

        SalesYesterdayLastWeek = 0
        While Not records.EOF
            Try
                SalesYesterdayLastWeek = records.Fields("Sumsales").Value
            Catch ex As Exception
                SalesYesterdayLastWeek = 0
            End Try
            records.MoveNext()
        End While
        records.Close()

        Try
            SalesYesterdayLastWeekVariance = ((SalesYesterday - SalesYesterdayLastWeek)) / SalesYesterdayLastWeek * 100
        Catch ex As Exception
            SalesYesterdayLastWeekVariance = 0
        End Try

        '     SalesYesterdayLastWeekVariance = SalesYesterday / SalesYesterdayLastWeek
        '      records = erpdb.Execute("Select * from [KPI Values] where [KPI Name] ='POS Sales LMTD - Total'")

        StartDateFrom01 = Now.AddDays((DemoParameter) * -1)
        StartDateFrom01 = StartDateFrom01.AddMonths(-1)
        StartDateFrom01 = StartDateFrom01.AddDays((Now.Day * -1) + 1)

        StartDateTo01 = Now.AddDays((DemoParameter) * -1)
        StartDateTo01 = StartDateTo01.AddDays(-1)
        StartDateTo01 = StartDateTo01.AddMonths(-1)
        'if the first then last month is whole month
        'lmtd is month -2
        If Now.Day = 1 Then
            StartDateFrom01 = Now.AddDays((DemoParameter) * -1)
            StartDateFrom01 = StartDateFrom01.AddMonths(-2)
            StartDateTo01 = Now.AddDays((DemoParameter) * -1)
            StartDateTo01 = StartDateTo01.AddMonths(-1)
            StartDateTo01 = StartDateTo01.AddDays(-1)

        End If
        '        records = erpdb.Execute("Select * from [KPI Values] where [KPI Name] ='POS Sales LMTD - Total'")

        records = erpdb.Execute("Select sum([Total Net Sales]) As Sumsales from [POSSales] where cast([Date Record] as date) >= '" & Format(StartDateFrom01, "dd-MMM-yyyy") &
                                "' and cast([Date Record] as date) <= '" & Format(StartDateTo01, "dd-MMM-yyyy") & "'")
        SQLStatement = "Select sum([Total Net Sales]) As Sumsales from [POSSales] where cast([Date Record] as date) >= '" & Format(StartDateFrom01, "dd-MMM-yyyy") &
                                "' and cast([Date Record] as date) <= '" & Format(StartDateTo01, "dd-MMM-yyyy") & "'"

        SalesLastMonthToDate = 0
        While Not records.EOF
            Try
                SalesLastMonthToDate = records.Fields("Sumsales").Value
            Catch ex As Exception
                SalesLastMonthToDate = 0
            End Try
            records.MoveNext()
        End While
        records.Close()

        '    Format((ResultsArray(countertemp, 1) - ResultsArray(countertemp, 3)) / ResultsArray(countertemp, 3) * 100, "0.##")
        Try
            SalesLastMonthVariance = (SalesMTDYesterday - SalesLastMonthToDate) / SalesLastMonthToDate * 100
        Catch ex As Exception
            SalesLastMonthVariance = 0
        End Try


        'Yesterday last week
        '      DailyReportHTMLEmailString = "      <div style = ""width:400px;"">
        '         <h1 style=""text-align:center""> <background-color: blue;>Testing </h1>
        '        <p  style =""text-align:right""><a href=""#"">sample link</a> </p>
        '       </div>"http://aagilitycom.ipage.com/wp-content/uploads/2017/12/BizCoach_small.png


        DailyReportHTMLEmailString = "<!DOCTYPE html><html>" & "    <img src = ""http://aagilitycom.ipage.com/wp-content/uploads/2017/12/BizCoach_small.png"" alt=""BizCoach"" ;  ><br /><br />"

        '       DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<font face=""calibri""><b><u> TEST EMAIL - NOT FOR CUSTOMERS - " & FrmMain.ListBoxCustomerInfo.SelectedItem &
        '           " - " & FrmMain.ListBoxNotificationType.SelectedItem & "</b></u><br /><br />"


        Dim ExcelHeaderLine0401 As String = "Yesterday: "
        Dim ExcelHeaderLine0402 As String = Format(SalesYesterday, "HK$ #,###,##0")
        Dim ExcelHeaderLine0403 As String = "Same day last week : "
        Dim ExcelHeaderLine0404 As String = Format(SalesYesterdayLastWeek, "HK$ #,###,##0")
        Dim ExcelHeaderLine0405 As String = "Variance to last week "
        Dim ExcelHeaderLine0406 As String = Format(SalesYesterdayLastWeekVariance, "0.##") & "%"

        Dim ExcelHeaderLine0501 As String = "MTD      : "
        Dim ExcelHeaderLine0502 As String = Format(SalesMTDYesterday, "HK$ #,###,##0")
        Dim ExcelHeaderLine0503 As String = "Last month to date "
        Dim ExcelHeaderLine0504 As String = Format(SalesLastMonthToDate, "HK$ #,###,##0")
        Dim ExcelHeaderLine0505 As String = "Variance to last month "
        Dim ExcelHeaderLine0506 As String = Format(SalesLastMonthVariance, "0.##") & "%"


        'DailyReportHTMLEmailString = DailyReportHTMLEmailString & "Inventory For : " & Format(DailyReportDate, "ddd d MMM yyyy") & "<br /><br />" &
        '                               "Inventory Value: " & Format(InventoryValue, "HK$ #,###,##0") & ", same day last week : " & Format(InventoryValue * 0.92, "HK$ #,###,##0") &
        '                              ", variance to last week " & Format(InventoryValue / (InventoryValue * 0.92), "0.##") & "%" & "<br />" &
        '                              "<br />" & "<br />" &
        '                            "<table border=1><col width=""300""><col width=""150""><col width=""100""><col width=""100""><col width=""100""><col width=""100""><col width=""100""><col width=""100""><col width=""100""><tr>" &
        '                           "<td><b>Item Group</td>" &
        '                          "<td><b><p style =""text-align:center"">Inventory HK$</td>" &
        '                          "</b></tr>"


        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "Daily Revenue Prediction For : " & Format(DailyReportDate, "ddd d MMM yyyy") & "<br /><br />" &
                                        "<br />" & "<br />" &
                                        "<table border=1><col width=""200""><col width=""150""><col width=""150""><col width=""150""><col width=""150"">" &
                                        "<tr>" &
                                            "<td>Revenue/Predictions</td>" &
                                            "<th colspan=""2""><p style =""text-align:center"">Actual</th>" &
                                            "<th colspan=""2""><p style =""text-align:center"">Prediction" & "</th>" &
                                        "</tr>"

        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<tr>
                    <td>" & "$" & "</td>
                    <td><p style =""text-align:center"">Week</td>
                    <td><p style =""text-align:center"">Month</td>
                    <td><p style =""text-align:center"">Week</td>
                    <td><p style =""text-align:center"">Month</td>
                    </tr>"


        '        DailyReportHTMLEmailString = "<!DOCTYPE html>
        '<html>
        '<head>
        '<style>
        'table, th, td {
        '    border: 1px solid black;
        '}
        '</style>
        '</head>
        '<body>
        '
        '<table>
        '    <tr>
        '      <td>Sales</td>
        '      <th colspan=""2"">Actual</th>
        '      <th colspan = ""2"">Prediction</th>
        '  </tr>
        '  <tr>
        '    <td>ITSBOS</td>
        '    <td>$100</td>
        '    <td>$130</td>
        '    <td>$200</td>
        '     <td>$30</td>
        '  </tr>
        '  <tr>
        '  <td>ITSMIA</td>
        '    <td>$100</td>
        '    <td>$130</td>
        '    <td>$200</td>
        '     <td>$30</td>
        '  </tr>
        '</table>
        '
        '</body>
        '</html>"



        SMSText = "Prediction for " & Format(Now, "ddd d MMM") & ControlChars.CrLf

        'now gets the individual outlets
        'now sums up
        '     records = erpdb.Execute("Select sum([KPI Value Number]) As Sumsales, [Outlet Name] from [KPI Values] where [KPI Name] ='POS Daily Sales (today)' group by [Outlet Name] order by [Outlet Name]")

        Dim LatestPredictionsDate As Date = Now.AddYears(-10)
        'gets the latest prediction date
        records = erpdb.Execute("select * from [Predictions] order by [prediction_created_on] Desc")
        While Not records.EOF
            Try
                LatestPredictionsDate = records.Fields("prediction_created_on").Value
            Catch ex As Exception

            End Try
            Exit While
            records.MoveNext()
        End While
        records.Close()
        LatestPredictionsDate = LatestPredictionsDate.ToShortDateString


        Dim SalesYesterdayAll As Decimal = 0
        Dim SalesPredictionMonth As Decimal = 0
        Dim SalesPredictionWeek As Decimal = 0
        ' Dim SalesPredictionPlus3 As Decimal = 0
        '       Dim SalesPredictionPlus4 As Decimal = 0
        '      Dim SalesPredictionPlus5 As Decimal = 0
        '     Dim SalesPredictionPlus6 As Decimal = 0
        '    Dim SalesPredictionPlus7 As Decimal = 0
        '   Dim TotalPrediction7Days As Decimal = 0

        'this gets all the outlets even if they have no sales
        SQLStatement = "select * from [Outlet Details]  where [Date Closed] is Null order by [Outlet Name]"


        Dim DaysintheMonth = DateTime.DaysInMonth(Now.Year, Now.Month)

        records = erpdb.Execute(SQLStatement)
        SalesYesterday = 0
        Dim DaysBackTemp As Int16 = 0
        Dim DaysForwardTemp As Int16 = 0
        Dim SalesMonthToDate As Decimal = 0
        Dim SalesMonthToDateAll As Decimal = 0
        Dim SalesPredictionMonthAll As Decimal = 0
        Dim SalesPredictionWeekAll As Decimal = 0

        While Not records.EOF

            Try
                OutletName = records.Fields("Outlet Name").Value
                '    SalesYesterday = records.Fields("Sumsales").Value
                '    SalesYesterdayAll = SalesYesterdayAll + SalesYesterday
            Catch ex As Exception
                MsgBox("error " & Err.Description)
            End Try

            'this report currently does a actual sales from monday to the current day
            'and a actual for the month

            If Now.DayOfWeek = DayOfWeek.Monday Then
                DaysBackTemp = 7
                DaysForwardTemp = 0
            ElseIf Now.DayOfWeek = DayOfWeek.Tuesday Then
                DaysBackTemp = 1
                DaysForwardTemp = 6
            ElseIf Now.DayOfWeek = DayOfWeek.Wednesday Then
                DaysBackTemp = 2
                DaysForwardTemp = 5
            ElseIf Now.DayOfWeek = DayOfWeek.Thursday Then
                DaysBackTemp = 3
                DaysForwardTemp = 4
            ElseIf Now.DayOfWeek = DayOfWeek.Friday Then
                DaysBackTemp = 4
                DaysForwardTemp = 3
            ElseIf Now.DayOfWeek = DayOfWeek.Saturday Then
                DaysBackTemp = 5
                DaysForwardTemp = 2
            ElseIf Now.DayOfWeek = DayOfWeek.Sunday Then
                DaysBackTemp = 6
                DaysForwardTemp = 1
            End If

            SQLStatement = "select sum([Total Net Sales]) as Sumsales , [Outlet Name] from [POSSales] where cast([Date Record] as date) >= '" &
                                Format(Now.AddDays(DaysBackTemp * -1), "dd-MMM-yyyy") & "' and cast([Date Record] as date) <= '" &
                                Format(Now.AddDays(-1), "dd-MMM-yyyy") & "' and [Outlet Name] ='" & Replace(OutletName, "'", "''") & "' group by [Outlet Name] order by [Outlet Name]"
            recordsLookup = erpdb.Execute(SQLStatement)
            SalesYesterday = 0
            While Not recordsLookup.EOF
                SalesYesterday = 0
                Try
                    OutletName = recordsLookup.Fields("Outlet Name").Value
                    SalesYesterday = recordsLookup.Fields("Sumsales").Value
                    SalesYesterdayAll = SalesYesterdayAll + SalesYesterday
                Catch ex As Exception
                    MsgBox("error " & Err.Description)
                End Try
                recordsLookup.MoveNext()
            End While
            recordsLookup.Close()


            'actual for the month
            SQLStatement = "select sum([Total Net Sales]) as Sumsales , [Outlet Name] from [POSSales] where cast([Date Record] as date) >= '" & "01" & Format(Now, "-MMM-yyyy") & "' and cast([Date Record] as date) <= '" &
                                Format(Now.AddDays(-1), "dd-MMM-yyyy") & "' and [Outlet Name] ='" & Replace(OutletName, "'", "''") & "' group by [Outlet Name] order by [Outlet Name]"
            recordsLookup = erpdb.Execute(SQLStatement)
            SalesMonthToDate = 0
            While Not recordsLookup.EOF
                Try
                    SalesMonthToDate = recordsLookup.Fields("Sumsales").Value
                    SalesMonthToDateAll = SalesMonthToDateAll + SalesMonthToDate
                Catch ex As Exception
                    MsgBox("error " & Err.Description)
                End Try
                recordsLookup.MoveNext()
            End While
            recordsLookup.Close()


            SQLStatement = "Select sum([predicted_value_numeric]) As Sumsales, [Outlet Name] from [Predictions] where cast([predicted_date] As Date) >= '" & Format(Now, "dd-MMM-yyyy") &
                "' and  cast([predicted_date] As Date) <= '" & Format(Now.AddDays(DaysForwardTemp - 1), "dd-MMM-yyyy") &
                "' and prediction_active ='1' and [Outlet Name] ='" & Replace(OutletName, "'", "''") & "' and [predictions_name] ='Total Net Sales - best fit' group by [Outlet Name]"
            recordsLookup = erpdb.Execute(SQLStatement)


            SalesPredictionWeek = 0
            While Not recordsLookup.EOF
                Try
                    SalesPredictionWeek = recordsLookup.Fields("Sumsales").Value
                    SalesPredictionWeekAll = SalesPredictionWeekall + SalesPredictionWeek
                    '       TotalPrediction7Days = TotalPrediction7Days + SalesPredictionPlus1
                Catch ex As Exception
                    SalesPredictionWeek = 0
                End Try
                recordsLookup.MoveNext()
            End While
            recordsLookup.Close()


            SQLStatement = "Select sum([predicted_value_numeric]) As Sumsales, [Outlet Name] from [Predictions] where cast([predicted_date] As Date) >= '" & Format(Now, "dd-MMM-yyyy") &
                "' and  cast([predicted_date] As Date) <= '" & DaysintheMonth & Format(Now, "-MMM-yyyy") &
                "' and prediction_active ='1' and [Outlet Name] ='" & Replace(OutletName, "'", "''") & "' and [predictions_name] ='Total Net Sales - best fit' group by [Outlet Name]"
            recordsLookup = erpdb.Execute(SQLStatement)


            SalesPredictionMonth = 0
            While Not recordsLookup.EOF
                Try
                    SalesPredictionMonth = recordsLookup.Fields("Sumsales").Value
                    SalesPredictionMonthAll = SalesPredictionMonthall + SalesPredictionMonth
                    '       TotalPrediction7Days = TotalPrediction7Days + SalesPredictionPlus1
                Catch ex As Exception
                    SalesPredictionMonth = 0
                End Try
                recordsLookup.MoveNext()
            End While
            recordsLookup.Close()



            'can now build the line
            '   DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<tr><td>" & OutletName & "</td>" &
            '                                "<td>" & "<p style =""text-align:right"">" & Format(SalesYesterday, "#,###,##0") & "</td>" &
            '                               "<td>" & "<p style =""text-align:right"">" & Format(SalesPredictionPlus1, "#,###,###") & "</td>" &
            '                             "<td>" & "<p style =""text-align:right"">" & Format(SalesPredictionPlus2, "#,###,###") & "</td>" &
            '                            "<td>" & "<p style =""text-align:right"">" & Format(SalesPredictionPlus3, "#,###,###") & "</td>" &
            '                           "<td>" & "<p style =""text-align:right"">" & Format(SalesPredictionPlus4, "#,###,###") & "</td>" &
            '                          "<td>" & "<p style =""text-align:right"">" & Format(SalesPredictionPlus5, "#,###,###") & "</td>" &
            '                         "<td>" & "<p style =""text-align:right"">" & Format(SalesPredictionPlus6, "#,###,###") & "</td>" &
            '                        "<td>" & "<p style =""text-align:right"">" & Format(SalesPredictionPlus7, "#,###,###") & "</tr>"


            DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<tr>
                    <td><p style =""text-align:left"">" & OutletName & "</td>
                    <td><p style =""text-align:right"">" & Format(SalesYesterday, "#,###,##0") & "</td>
                    <td><p style =""text-align:right"">" & Format(SalesMonthToDate, "#,###,##0") & "</td>
                    <td><p style =""text-align:right"">" & Format(SalesYesterday + SalesPredictionWeek, "#,###,##0") & "</td>
                    <td><p style =""text-align:right"">" & Format(SalesMonthToDate + SalesPredictionMonth, "#,###,##0") & "</td>
                    </tr>"

            SMSText = SMSText & OutletName & " full week " & " $ " & Format(SalesYesterday + SalesPredictionWeek, "#,###,###") & " " & ControlChars.CrLf

            'looks up full outlet name if available
            RecordsOutletLookup = erpdb.Execute("select * from [Outlet Details] where [Outlet Name]='" & Replace(OutletName, "'", "''") & "'")
            OutletFullNameTemp = ""
            While Not RecordsOutletLookup.EOF
                Try
                    OutletFullNameTemp = RecordsOutletLookup.Fields("Description").Value
                Catch ex As Exception
                    OutletFullNameTemp = ""
                End Try

                RecordsOutletLookup.MoveNext()
            End While
            RecordsOutletLookup.Close()
            If OutletFullNameTemp = "" Then
                OutletFullNameTemp = OutletName
            End If

            If counter01 = 0 Then
                '               SpeechString = SpeechString & " " & OutletFullNameTemp & " sold " & Format(SalesYesterday, "#,###") & "$ " &
                '                  " yesterday and is predicted to sell " & Format(SalesPredictionPlus1, "#,###,###") & "$ on " & Format(DailyReportDate.AddDays(1), "dddd") &
                '                 " and " & Format(SalesPredictionPlus2, "#,###,###") & "$ on " & Format(DailyReportDate.AddDays(2), "dddd") &
                '                " a total of " & Format(SalesPredictionPlus1 + SalesPredictionPlus2 + SalesPredictionPlus3 + SalesPredictionPlus4 + SalesPredictionPlus5 +
                '           SalesPredictionPlus6 + SalesPredictionPlus7, "#,###,###") & "$ by " & Format(DailyReportDate.AddDays(7), "dddd") & ". "

            Else
                '             SpeechString = SpeechString & " " & OutletFullNameTemp & " sold " & Format(SalesYesterday, "#,###") & "$ " &
                '                " yesterday and is predicted to sell " & Format(SalesPredictionPlus1, "#,###,###") & "$ on " & Format(DailyReportDate.AddDays(1), "dddd") &
                '               " and " & Format(SalesPredictionPlus2, "#,###,###") & "$ on " & Format(DailyReportDate.AddDays(2), "dddd") &
                '              " a total of " & Format(SalesPredictionPlus1 + SalesPredictionPlus2 + SalesPredictionPlus3 + SalesPredictionPlus4 + SalesPredictionPlus5 +
                '         SalesPredictionPlus6 + SalesPredictionPlus7, "#,###,###") & "$ by " & Format(DailyReportDate.AddDays(7), "dddd") & ". "

            End If

            counter01 += 1

            records.MoveNext()
        End While
        records.Close()
        erpdb.Close()



        'group
        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<tr>
                    <td><p style =""text-align:left"">" & "All Outlets" & "</td>
                    <td><p style =""text-align:right"">" & Format(SalesYesterdayAll, "#,###,##0") & "</td>
                    <td><p style =""text-align:right"">" & Format(SalesMonthToDateAll, "#,###,##0") & "</td>
                    <td><p style =""text-align:right"">" & Format(SalesPredictionWeekAll + SalesYesterdayAll, "#,###,##0") & "</td>
                    <td><p style =""text-align:right"">" & Format(SalesMonthToDateAll + SalesPredictionMonth, "#,###,##0") & "</td>
                    </tr>"






        'puts the bizcoach logo
        'this is desktop format so just a text string
        '"Yesterday: HK$ 1,014, last wk: HK$ 10,793"
        'http://aagilitycom.ipage.com/custview/20Jan1802glorystores.htm
        Dim CompanynameTemp As String = Companyname.Trim
        CompanynameTemp = Replace(CompanynameTemp, ".", "")
        CompanynameTemp = Replace(CompanynameTemp, " ", "")
        CompanynameTemp = Replace(CompanynameTemp, " ", "")
        CompanynameTemp = Replace(CompanynameTemp, "'", "")
        CompanynameTemp = CompanynameTemp.ToLower

        '      SMSText = SMSText & ControlChars.CrLf & "Prediction $ " & Format(TotalPrediction7Days, "#,###,##0") & " to " & Format(DailyReportDate.AddDays(7), "ddd d MMM")

        UrlOfReport = "http://aagilitycom.ipage.com/custview/" & Format(DailyReportDate, "ddMMMyyyy") & "02" & CompanynameTemp & ".htm"
        DesktopText = "Yesterday: " & Format(SalesYesterdayAll, "HK$ #,###,##0") & ", last wk:" & Format(SalesYesterdayLastWeek, "HK$ #,###,##0")


 '       SpeechString = SpeechString & ". Total prediction for the next 7 days is " & Format(TotalPrediction7Days, "#,###") & "$ . "

        FrmMain.TextBoxSpeechText.Text = SpeechString.ToString.Trim

        'this creates the speech file

        AWSPollySpeak(SpeechString, "PredictionsPOSMonthly")
        'this publishes the speech to the internet 
        PublishSpeechtoImpala(SpeechString, "PredictionsPOSMonthly")
        'the file is bublished to    
        'http://aagilitycom.ipage.com/custview/    TextBoxPublishLocationName.text


        'closes the table
        DailyReportHTMLEmailString = "" & DailyReportHTMLEmailString & "</table><br /><hr /><br />"
        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<br />"



        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "To see your dashboards login to BizCoach <a href=""https://app.impalabizcoach.com/"">here</a><br /><br />"

        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "To listen to your revenue snapshot click <a href=""http://aagilitycom.ipage.com/custview/" &
            FrmMain.TextBoxPublishLocationName.Text.Trim & """>here</a><br />"
        DailyReportHTMLEmailString = "" & DailyReportHTMLEmailString & "</table></span><br /><hr /><br />"
        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "Let us know how we are doing :<br /><br />"
        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "2 minute  <a href=""https://www.surveymonkey.com/r/M2Q557H"">survey</a><br />"
        DailyReportHTMLEmailString = DailyReportHTMLEmailString & " or<br />"
        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "email : <a href=""mailto:support@impalacloud.com"" target=""_top"">support@impalacloud.com</a><br /><br />"


    
        'this creates the excel
        If FrmMain.CheckBoxCreateExcel.Checked = True Then
            CreateExcelSpreadsheet(DailyReportDate, ExcelHeaderLine0401,
                                      ExcelHeaderLine0402,
        ExcelHeaderLine0403,
        ExcelHeaderLine0404,
        ExcelHeaderLine0405,
        ExcelHeaderLine0406,
        ExcelHeaderLine0501,
        ExcelHeaderLine0502,
        ExcelHeaderLine0503,
        ExcelHeaderLine0504,
        ExcelHeaderLine0505,
        ExcelHeaderLine0506)
        End If


        'publish to the web browser on the notifications page
        Dim document As System.Windows.Forms.HtmlDocument = FrmMain.WebBrowserNotification.Document
        FrmMain.WebBrowserNotification.DocumentText = DailyReportHTMLEmailString


    End Sub




    Public Sub InventoryPOSSystemWeekly(ByRef DailyReportHTMLEmailString As String, ByRef DailyReportDate As Date,
                                                   ByVal DemoCompany As Boolean, ByRef UrlOfReport As String,
                                                   ByRef DesktopText As String, ByRef Companyname As String,
                                                   ByRef SMSText As String, ByRef SpeechString As String)

        'this is a daily report that is based upon yesterdays sales
        'it compares this to the monthly and last month figures

        'as fixed date report then can take off the date here
        DailyReportDate = Now.AddDays(-1)



        'dont use KPIs now as may be a timing issue

        Dim erpdb As ADODB.Connection
        Dim records As ADODB._Recordset
        Dim recordsLookup As ADODB._Recordset
        Dim RecordsOutletLookup As ADODB._Recordset
        erpdb = New ADODB.Connection
        erpdb.Open(strconnectSQLBizCoachCustomerDataSQL)
        Dim InventoryValue As Decimal = 0
        Dim SalesMTDYesterday As Integer = 0
        Dim OutletName As String = ""
        Dim SalesYesterdayLastWeek As Integer = 0
        Dim SalesLastMonthVariance As Decimal = 0
        Dim SalesYesterdayLastWeekVariance As Decimal = 0
        Dim SalesLastMonthToDate As Integer = 0
        Dim StartDateFrom01 As Date = Now.ToShortDateString
        Dim StartDateTo01 As Date = Now.ToShortDateString
        Dim SQLStatement As String = ""
        Dim OutletFullNameTemp As String = ""
        Dim counter01 As Int16 = 0
        'this stores the excel data
        Dim ExcelTableData(10000, 10)

        'used for demo companies to go back 1 year
        SMSText = ""
        SpeechString = "Inventory holdings are "

        Dim DemoParameter As Int16 = 0
        'uses UTC time so need to take away 8 hours or use HongKongLoginMoment

        If DemoCompany = True Then
            DemoParameter = 365
            DailyReportDate = DailyReportDate.AddYears(-1)
        End If
        DailyReportDate = DailyReportDate.ToShortDateString

        '     records = erpdb.Execute("Select * from [KPI Values] where [KPI Name] ='POS Daily Sales (yesterday) - Total'")

        Dim qtyOnhand As Decimal = 0
        Dim CostOfInv As Decimal = 0

        SQLStatement = "Select * from [InventoryView]"
        records = erpdb.Execute(SQLStatement)

        InventoryValue = 0
        While Not records.EOF
            Try
                qtyOnhand = 0
                CostOfInv = 0

                qtyOnhand = records.Fields("Quantity On Hand").Value
                CostOfInv = records.Fields("Cost Price").Value

                InventoryValue = InventoryValue + (qtyOnhand * CostOfInv)
            Catch ex As Exception
           '     InventoryValue = 0
            End Try
            records.MoveNext()
        End While
        records.Close()

        SpeechString = SpeechString & Format(InventoryValue, "#,###") & "$" & " for your business, same day last week " & Format(InventoryValue * 0.92, "$ #,###,##0") & ". "



        'Yesterday last week
        '      DailyReportHTMLEmailString = "      <div style = ""width:400px;"">
        '         <h1 style=""text-align:center""> <background-color: blue;>Testing </h1>
        '        <p  style =""text-align:right""><a href=""#"">sample link</a> </p>
        '       </div>"
        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "    <img src = ""http://aagilitycom.ipage.com/wp-content/uploads/2017/12/BizCoach_small.png"" alt=""BizCoach"" ;  ><br /><br />"

        '        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<font face=""calibri""><b><u> TEST EMAIL - NOT FOR CUSTOMERS - " & FrmMain.ListBoxCustomerInfo.SelectedItem &
        '                " - " & FrmMain.ListBoxNotificationType.SelectedItem & "</b></u><br /><br />"



        Dim ExcelHeaderLine0401 As String = "Yesterday: "
        Dim ExcelHeaderLine0402 As String = Format(InventoryValue, "HK$ #,###,##0")
        Dim ExcelHeaderLine0403 As String = ""
        Dim ExcelHeaderLine0404 As String = ""
        Dim ExcelHeaderLine0405 As String = ""
        Dim ExcelHeaderLine0406 As String = ""

        Dim ExcelHeaderLine0501 As String = ""
        Dim ExcelHeaderLine0502 As String = ""
        Dim ExcelHeaderLine0503 As String = ""
        Dim ExcelHeaderLine0504 As String = ""
        Dim ExcelHeaderLine0505 As String = ""
        Dim ExcelHeaderLine0506 As String = ""

        Dim ItemGroup(500, 3)
        Dim NumGroups As Int16 = 0
        'gets the item groups
        SQLStatement = "select top(500) [Item Group] from [InventoryView] group by [Item Group]  order by [Item Group]"
        records = erpdb.Execute(SQLStatement)

        While Not records.EOF
            Try
                ItemGroup(NumGroups, 0) = records.Fields("Item Group").Value
                NumGroups += 1
            Catch ex As Exception

            End Try
            records.MoveNext()
        End While
        records.Close()

        'this sums up the inventory by group
        Dim qtyOnhandGroup As Decimal = 0
        Dim CostOfInvGroup As Decimal = 0
        Dim InventoryValueGroup As Decimal = 0

        For counter = 0 To NumGroups - 1
            SQLStatement = "Select * from [InventoryView] where [Item Group] ='" & ItemGroup(counter, 0) & "'"
            records = erpdb.Execute(SQLStatement)

            InventoryValueGroup = 0
            While Not records.EOF
                Try
                    qtyOnhandGroup = 0
                    CostOfInvGroup = 0

                    qtyOnhandGroup = records.Fields("Quantity On Hand").Value
                    CostOfInvGroup = records.Fields("Cost Price").Value

                    InventoryValueGroup = InventoryValueGroup + (qtyOnhandGroup * CostOfInvGroup)

                    ItemGroup(counter, 1) = InventoryValueGroup
                Catch ex As Exception
                    ' InventoryValue = 0
                End Try
                records.MoveNext()
            End While
            records.Close()
        Next


        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "Inventory For : " & Format(DailyReportDate, "ddd d MMM yyyy") & "<br /><br />" &
                                        "Inventory Value: " & Format(InventoryValue, "HK$ #,###,##0") & ", same day last week : " & Format(InventoryValue * 0.92, "HK$ #,###,##0") &
                                        ", variance to last week " & Format(InventoryValue/(InventoryValue *.92), "0.##") & "%" & "<br />" &
                                         "<br />" & "<br />" &
                                        "<table border=1><col width=""300""><col width=""150""><col width=""100""><col width=""100""><col width=""100""><col width=""100""><col width=""100""><col width=""100""><col width=""100""><tr>" &
                                        "<td><b>Item Group</td>" &
                                        "<td><b><p style =""text-align:center"">Inventory HK$</td>" &
                                         "</b></tr>"

        SMSText = "Inventory for : " & Format(DailyReportDate.AddDays(1), "ddd d MMM") & ControlChars.CrLf
        SMSText = SMSText & Format(InventoryValue, "#,###,###") & " $. " & ControlChars.CrLf


        Dim GroupNameText As String = ""

        For counter = 0 To NumGroups - 1

            Try
                If ItemGroup(counter, 0).ToString.Trim = "" Then
                    GroupNameText = "Not classified"
                Else
                    GroupNameText = ItemGroup(counter, 0).ToString.Trim
                End If
            Catch ex As Exception
                GroupNameText = "Not classified"
            End Try


            'can now build the line
            DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<tr><td>" & "<p style =""text-align:right"">" & GroupNameText & "</td>" &
                                          "<td>" & "<p style =""text-align:right"">" & Format(ItemGroup(counter, 1), "#,###,###") & "</tr>"

            If counter = 0 Then
                SpeechString = SpeechString & " " & GroupNameText & " has inventory value of " & Format(ItemGroup(counter, 1), "#,###") & "$ . "
            Else
                SpeechString = SpeechString & " " & GroupNameText & " has " & Format(ItemGroup(counter, 1), "#,###") & "$ . "
            End If
        Next


        'puts the bizcoach logo
        'this is desktop format so just a text string
        '"Yesterday: HK$ 1,014, last wk: HK$ 10,793"
        'http://aagilitycom.ipage.com/custview/20Jan1802glorystores.htm
        Dim CompanynameTemp As String = Companyname.Trim
        CompanynameTemp = Replace(CompanynameTemp, ".", "")
        CompanynameTemp = Replace(CompanynameTemp, " ", "")
        CompanynameTemp = Replace(CompanynameTemp, " ", "")
        CompanynameTemp = Replace(CompanynameTemp, "'", "")
        CompanynameTemp = CompanynameTemp.ToLower

        '      SMSText = SMSText & ControlChars.CrLf & "Total: $ " & Format(SalesYesterdayAll, "#,###,##0")

        UrlOfReport = "http://aagilitycom.ipage.com/custview/" & Format(DailyReportDate, "ddMMMyyyy") & "02" & CompanynameTemp & ".htm"
        DesktopText = "Yesterday: " & Format(InventoryValue, "HK$ #,###,##0") & ", last wk:" & Format(SalesYesterdayLastWeek, "HK$ #,###,##0")


  '      SpeechString = SpeechString & ". Total prediction for the next 7 days is " & Format(TotalPrediction7Days, "#,###") & "$ . "

        FrmMain.TextBoxSpeechText.Text = SpeechString.ToString.Trim

        'this creates the speech file

        AWSPollySpeak(SpeechString, "InventoryPOSWeekly")
        'this publishes the speech to the internet 
        PublishSpeechtoImpala(SpeechString, "InventoryPOSWeekly")
        'the file is bublished to    
        'http://aagilitycom.ipage.com/custview/    TextBoxPublishLocationName.text


        'closes the table
        DailyReportHTMLEmailString = "" & DailyReportHTMLEmailString & "</table></span><br /><hr /><br />"
        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<br />"

        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "If you would like BizCoach to read your email click <a href=""http://aagilitycom.ipage.com/custview/" &
            FrmMain.TextBoxPublishLocationName.Text.Trim & """>here</a><br /><br />"





        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "Please give feedback To <a href=""mailto:support@impalacloud.com"" target=""_top"">support@impalacloud.com</a><br /><br />"

        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<br />"

        'this creates the excel
        If FrmMain.CheckBoxCreateExcel.Checked = True Then
            CreateExcelSpreadsheet(DailyReportDate, ExcelHeaderLine0401,
                                      ExcelHeaderLine0402,
        ExcelHeaderLine0403,
        ExcelHeaderLine0404,
        ExcelHeaderLine0405,
        ExcelHeaderLine0406,
        ExcelHeaderLine0501,
        ExcelHeaderLine0502,
        ExcelHeaderLine0503,
        ExcelHeaderLine0504,
        ExcelHeaderLine0505,
        ExcelHeaderLine0506)
        End If


        'publish to the web browser on the notifications page
        Dim document As System.Windows.Forms.HtmlDocument = FrmMain.WebBrowserNotification.Document
        FrmMain.WebBrowserNotification.DocumentText = DailyReportHTMLEmailString

    End Sub

    Public Sub SalesByItemPOS(ByRef DailyReportHTMLEmailString As String, ByRef DailyReportDate As Date,
                                                   ByVal DemoCompany As Boolean, ByRef UrlOfReport As String,
                                                   ByRef DesktopText As String, ByRef Companyname As String,
                                                   ByRef SMSText As String, ByRef SpeechString As String)

        'this is a daily report that is based upon yesterdays sales
        'it compares this to the monthly and last month figures

        'as fixed date report then can take off the date here
        DailyReportDate = Now.AddDays(-1)

        'dont use KPIs now as may be a timing issue

        Dim erpdb As ADODB.Connection
        Dim records As ADODB._Recordset
        Dim recordsLookup As ADODB._Recordset
        Dim RecordsOutletLookup As ADODB._Recordset
        erpdb = New ADODB.Connection
        erpdb.Open(strconnectSQLBizCoachCustomerDataSQL)
        Dim InventoryValue As Decimal = 0
        Dim SalesMTDYesterday As Integer = 0
        Dim OutletName As String = ""
        Dim SalesYesterdayLastWeek As Integer = 0
        Dim SalesLastMonthVariance As Decimal = 0
        Dim SalesYesterdayLastWeekVariance As Decimal = 0
        Dim SalesLastMonthToDate As Integer = 0
        Dim StartDateFrom01 As Date = Now.ToShortDateString
        Dim StartDateTo01 As Date = Now.ToShortDateString
        Dim SQLStatement As String = ""
        Dim OutletFullNameTemp As String = ""
        Dim counter01 As Int16 = 0
        'this stores the excel data
        Dim ExcelTableData(10000, 10)

        'used for demo companies to go back 1 year
        SMSText = ""
        SpeechString = "Sales are "

        Dim DemoParameter As Int16 = 0
        'uses UTC time so need to take away 8 hours or use HongKongLoginMoment

        If DemoCompany = True Then
            DemoParameter = 365
            DailyReportDate = DailyReportDate.AddYears(-1)
        End If
        DailyReportDate = DailyReportDate.ToShortDateString

        '     records = erpdb.Execute("Select * from [KPI Values] where [KPI Name] ='POS Daily Sales (yesterday) - Total'")

        '   Dim qtyOnhand As Decimal = 0
        '    Dim CostOfInv As Decimal = 0

        SQLStatement = " select ItemName, sum([Total Net Sales]) as SumTotalNetSales, sum([Quantity Sold]) as sumqty from [POSSalesView] where cast([Date Record] as date) >='" & "01" & Format(Now, "-MMM-yyyy") & "' group by ItemName order by SumTotalNetSales desc"
        records = erpdb.Execute(SQLStatement)

        Dim TotalMonthlySalesToDate As Decimal = 0


        While Not records.EOF
            Try
                TotalMonthlySalesToDate = TotalMonthlySalesToDate + records.Fields("SumTotalNetSales").Value
            Catch ex As Exception
                '     InventoryValue = 0
            End Try
            records.MoveNext()
        End While
        records.Close()

        SQLStatement = " select ItemName, sum([Total Net Sales]) as SumTotalNetSales, sum([Quantity Sold]) as sumqty from [POSSalesView] where  cast([Date Record] as date) >='" &
            "01" & Format(Now.AddMonths(-1), "-MMM-yyyy") & "' and   cast([Date Record] as date) <='" & Format(Now.AddMonths(-1), "dd-MMM-yyyy") & "' group by ItemName order by SumTotalNetSales desc"
        records = erpdb.Execute(SQLStatement)

        Dim TotalMonthlySalesLastMonth As Decimal = 0
        While Not records.EOF
            Try
                TotalMonthlySalesLastMonth = TotalMonthlySalesLastMonth + records.Fields("SumTotalNetSales").Value
            Catch ex As Exception
                '     InventoryValue = 0
            End Try
            records.MoveNext()
        End While
        records.Close()


        SpeechString = SpeechString & Format(TotalMonthlySalesToDate, "#,###") & "$" & " for your business this month, last month was " & Format(TotalMonthlySalesLastMonth, "$ #,###,##0") & ". "



        'Yesterday last week
        '      DailyReportHTMLEmailString = "      <div style = ""width:400px;"">
        '         <h1 style=""text-align:center""> <background-color: blue;>Testing </h1>
        '        <p  style =""text-align:right""><a href=""#"">sample link</a> </p>
        '       </div>"
        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "    <img src = ""http://aagilitycom.ipage.com/wp-content/uploads/2017/12/BizCoach_small.png"" alt=""BizCoach"" ;  ><br /><br />"

        '       DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<font face=""calibri""><b><u> TEST EMAIL - NOT FOR CUSTOMERS - " & FrmMain.ListBoxCustomerInfo.SelectedItem &
        '                " - " & FrmMain.ListBoxNotificationType.SelectedItem & "</b></u><br /><br />"


        Dim ExcelHeaderLine0401 As String = "Yesterday: "
        Dim ExcelHeaderLine0402 As String = Format(InventoryValue, "HK$ #,###,##0")
        Dim ExcelHeaderLine0403 As String = ""
        Dim ExcelHeaderLine0404 As String = ""
        Dim ExcelHeaderLine0405 As String = ""
        Dim ExcelHeaderLine0406 As String = ""

        Dim ExcelHeaderLine0501 As String = ""
        Dim ExcelHeaderLine0502 As String = ""
        Dim ExcelHeaderLine0503 As String = ""
        Dim ExcelHeaderLine0504 As String = ""
        Dim ExcelHeaderLine0505 As String = ""
        Dim ExcelHeaderLine0506 As String = ""

        Dim ItemGroup(500, 3)
        Dim NumGroups As Int16 = 0

        Dim VarianceLastMonth As Decimal = 0
        'can be divide by zero error
        Try
            VarianceLastMonth = TotalMonthlySalesToDate / TotalMonthlySalesLastMonth
        Catch ex As Exception
            VarianceLastMonth = 0
        End Try

        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "Sales by item For : " & Format(DailyReportDate, "ddd d MMM yyyy") & "<br /><br />" &
                             "Monthly sales : " & Format(TotalMonthlySalesToDate, "HK$ #,###,##0") & ", last month : " & Format(TotalMonthlySalesLastMonth, "HK$ #,###,##0") &
                             ", variance to last month " & Format(VarianceLastMonth, "0.##") & "%" & "<br />" &
                              "<br />" & "<br />" &
                             "<table border=1><col width=""350""><col width=""150""><col width=""150""><col width=""150""><tr>" &
                             "<td><b><p style =""text-align:left"">Item Name</td>" &
                             "<td><b><p style =""text-align:center"">Total Sales HK$</td>" &
                            "<td><b><p style =""text-align:center"">Quantity Sold</td>" &
                              "</b></tr>"




        SMSText = "Inventory for : " & Format(DailyReportDate.AddDays(1), "ddd d MMM") & ControlChars.CrLf
        SMSText = SMSText & Format(InventoryValue, "#,###,###") & " $. " & ControlChars.CrLf


        'gets the item groups
        SQLStatement = " select top(10)  sum([Total Net Sales]) as SumTotalNetSales, sum([Quantity Sold]) as sumqty, ItemName from [POSSalesView] where  cast([Date Record] as date) >='" &
            "01" & Format(Now, "-MMM-yyyy") & "' group by  ItemName order by SumTotalNetSales desc"
        records = erpdb.Execute(SQLStatement)
        '   cast([Date Record] as date)

        counter01 = 0
        Dim ItemNameTemp As String = ""
        Dim SumTotalNetSalesTemp As Decimal = 0
        Dim sumqtyTemp As Decimal = 0

        While Not records.EOF
            Try
                ItemNameTemp = ""
                ItemNameTemp = records.Fields("ItemName").Value

            Catch ex As Exception

            End Try
            SumTotalNetSalesTemp = 0
            SumTotalNetSalesTemp = records.Fields("SumTotalNetSales").Value
            sumqtyTemp = 0
            sumqtyTemp = records.Fields("sumqty").Value

            'can now build the line
            DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<tr><td>" & "<p style =""text-align:left"">" & ItemNameTemp & "</td>" &
                                            "<td>" & "<p style =""text-align:right"">" & Format(SumTotalNetSalesTemp, "#,###,###") & "</td>" &
                                            "<td>" & "<p style =""text-align:right"">" & Format(sumqtyTemp, "#,###,###") & "</tr>"

            If counter01 = 0 Then
                SpeechString = SpeechString & " " & ItemNameTemp & " has sales of " & Format(SumTotalNetSalesTemp, "#,###") & "$ with " & Format(sumqtyTemp, "#,###,###") & " sold. "
            Else
                SpeechString = SpeechString & " " & ItemNameTemp & " sales " & Format(SumTotalNetSalesTemp, "#,###") & "$ , " & Format(sumqtyTemp, "#,###,###") & " sold. "
            End If
            counter01 += 1

            records.MoveNext()
        End While
        records.Close()


        DailyReportHTMLEmailString = "" & DailyReportHTMLEmailString & "</table></span><br /><hr /><br />"
        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<br />"


        'sales by item
        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "Sales by Item by Outlet : " &
                                         "<br />" & "<br />" &
                                        "<table border=1><col width=""150""><col width=""350""><col width=""150""><col width=""150""><tr>" &
                                        "<td><b><p style =""text-align:center"">Outlet</td>" &
                                        "<td><b><p style =""text-align:left"">Item Name </td>" &
                                        "<td><b><p style =""text-align:center"">Total Sales HK$</td>" &
                                       "<td><b><p style =""text-align:center"">Quantity Sold</td>" &
                                         "</b></tr>"



        Dim OutletNameTemp As String = ""

        'gets the item groups
        SQLStatement = " Select top(20) [Outlet Name], ItemName, sum([Total Net Sales]) As SumTotalNetSales, sum([Quantity Sold]) As sumqty  from [POSSalesView] where cast([Date Record] as date) >='" &
            "01" & Format(Now, "-MMM-yyyy") & "' group by [Outlet Name], ItemName order by SumTotalNetSales desc"

        records = erpdb.Execute(SQLStatement)
            '   cast([Date Record] as date)

            counter01 = 0
         
            While Not records.EOF
            Try
                ItemNameTemp = ""
                ItemNameTemp = records.Fields("ItemName").Value

            Catch ex As Exception

            End Try

            Try
                OutletNameTemp = ""
                OutletNameTemp = records.Fields("Outlet Name").Value

            Catch ex As Exception

            End Try


            SumTotalNetSalesTemp = 0
                SumTotalNetSalesTemp = records.Fields("SumTotalNetSales").Value
                sumqtyTemp = 0
                sumqtyTemp = records.Fields("sumqty").Value

            'can now build the line
            DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<tr><td>" & "<p style =""text-align:left"">" & OutletNameTemp & "</td>" &
                                            "<td>" & "<p style =""text-align:left"">" & ItemNameTemp & "</td>" &
                                              "<td>" & "<p style =""text-align:right"">" & Format(SumTotalNetSalesTemp, "#,###,###") & "</td>" &
                                          "<td>" & "<p style =""text-align:right"">" & Format(sumqtyTemp, "#,###,###") & "</tr>"

            If counter01 = 0 Then
                '       SpeechString = SpeechString & " " & ItemNameTemp & " has sales of " & Format(SumTotalNetSalesTemp, "#,###") & "$ with " & Format(sumqtyTemp, "#,###,###") & " sold. "
            Else
                '       SpeechString = SpeechString & " " & ItemNameTemp & " sales " & Format(SumTotalNetSalesTemp, "#,###") & "$ , " & Format(sumqtyTemp, "#,###,###") & " sold. "
            End If
            counter01 += 1

                records.MoveNext()
            End While
            records.Close()


        'puts the bizcoach logo
        'this is desktop format so just a text string
        '"Yesterday: HK$ 1,014, last wk: HK$ 10,793"
        'http://aagilitycom.ipage.com/custview/20Jan1802glorystores.htm
        Dim CompanynameTemp As String = Companyname.Trim
        CompanynameTemp = Replace(CompanynameTemp, ".", "")
        CompanynameTemp = Replace(CompanynameTemp, " ", "")
        CompanynameTemp = Replace(CompanynameTemp, " ", "")
        CompanynameTemp = Replace(CompanynameTemp, "'", "")
        CompanynameTemp = CompanynameTemp.ToLower

        '      SMSText = SMSText & ControlChars.CrLf & "Total: $ " & Format(SalesYesterdayAll, "#,###,##0")

        UrlOfReport = "http://aagilitycom.ipage.com/custview/" & Format(DailyReportDate, "ddMMMyyyy") & "02" & CompanynameTemp & ".htm"
        DesktopText = "Yesterday: " & Format(InventoryValue, "HK$ #,###,##0") & ", last wk:" & Format(SalesYesterdayLastWeek, "HK$ #,###,##0")


        '      SpeechString = SpeechString & ". Total prediction for the next 7 days is " & Format(TotalPrediction7Days, "#,###") & "$ . "

        FrmMain.TextBoxSpeechText.Text = SpeechString.ToString.Trim

        'this creates the speech file

        AWSPollySpeak(SpeechString, "SalesByItemPOS")
        'this publishes the speech to the internet 
        PublishSpeechtoImpala(SpeechString, "SalesByItemPOS")
        'the file is bublished to    
        'http://aagilitycom.ipage.com/custview/    TextBoxPublishLocationName.text


        'closes the table
        DailyReportHTMLEmailString = "" & DailyReportHTMLEmailString & "</table></span><br /><hr /><br />"
        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<br />"

        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "If you would like BizCoach to read your email click <a href=""http://aagilitycom.ipage.com/custview/" &
            FrmMain.TextBoxPublishLocationName.Text.Trim & """>here</a><br /><br />"


        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "Please give feedback To <a href=""mailto:support@impalacloud.com"" target=""_top"">support@impalacloud.com</a><br /><br />"

        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<br />"

        'this creates the excel
        If FrmMain.CheckBoxCreateExcel.Checked = True Then
            CreateExcelSpreadsheet(DailyReportDate, ExcelHeaderLine0401,
                                      ExcelHeaderLine0402,
        ExcelHeaderLine0403,
        ExcelHeaderLine0404,
        ExcelHeaderLine0405,
        ExcelHeaderLine0406,
        ExcelHeaderLine0501,
        ExcelHeaderLine0502,
        ExcelHeaderLine0503,
        ExcelHeaderLine0504,
        ExcelHeaderLine0505,
        ExcelHeaderLine0506)
        End If


        'publish to the web browser on the notifications page
        Dim document As System.Windows.Forms.HtmlDocument = FrmMain.WebBrowserNotification.Document
        FrmMain.WebBrowserNotification.DocumentText = DailyReportHTMLEmailString

    End Sub


    Public Sub SalesByHourPOS(ByRef DailyReportHTMLEmailString As String, ByRef DailyReportDate As Date,
                                                   ByVal DemoCompany As Boolean, ByRef UrlOfReport As String,
                                                   ByRef DesktopText As String, ByRef Companyname As String,
                                                   ByRef SMSText As String, ByRef SpeechString As String)

        'this is a daily report that is based upon yesterdays sales
        'it compares this to the monthly and last month figures

        'as fixed date report then can take off the date here
        DailyReportDate = Now.AddDays(-1)

        'dont use KPIs now as may be a timing issue

        Dim erpdb As ADODB.Connection
        Dim records As ADODB._Recordset
        Dim recordsLookup As ADODB._Recordset
        Dim RecordsOutletLookup As ADODB._Recordset
        erpdb = New ADODB.Connection
        erpdb.Open(strconnectSQLBizCoachCustomerDataSQL)
        Dim InventoryValue As Decimal = 0
        Dim SalesMTDYesterday As Integer = 0
        Dim OutletName As String = ""
        Dim SalesYesterdayLastWeek As Integer = 0
        Dim SalesLastMonthVariance As Decimal = 0
        Dim SalesYesterdayLastWeekVariance As Decimal = 0
        Dim SalesLastMonthToDate As Integer = 0
        Dim StartDateFrom01 As Date = Now.ToShortDateString
        Dim StartDateTo01 As Date = Now.ToShortDateString
        Dim SQLStatement As String = ""
        Dim OutletFullNameTemp As String = ""
        Dim counter01 As Int16 = 0
        'this stores the excel data
        Dim ExcelTableData(10000, 10)





        Dim DemoParameter As Int16 = 0
        'uses UTC time so need to take away 8 hours or use HongKongLoginMoment

        If DemoCompany = True Then
            DemoParameter = 365
            DailyReportDate = DailyReportDate.AddYears(-1)
        End If
        DailyReportDate = DailyReportDate.ToShortDateString

        '     records = erpdb.Execute("Select * from [KPI Values] where [KPI Name] ='POS Daily Sales (yesterday) - Total'")

        '   Dim qtyOnhand As Decimal = 0
        '    Dim CostOfInv As Decimal = 0

        SQLStatement = " select ItemName, sum([Total Net Sales]) as SumTotalNetSales, sum([Quantity Sold]) as sumqty from [POSSalesView] where cast([Date Record] 
            as date) >='" & "01" & Format(Now, "-MMM-yyyy") & "' group by ItemName order by SumTotalNetSales desc"
        records = erpdb.Execute(SQLStatement)
        Dim TotalMonthlySalesToDate As Decimal = 0
        While Not records.EOF
            Try
                TotalMonthlySalesToDate = TotalMonthlySalesToDate + records.Fields("SumTotalNetSales").Value
            Catch ex As Exception
                '     InventoryValue = 0
            End Try
            records.MoveNext()
        End While
        records.Close()

        'this is loading all the sales into an array for the current month
        Dim POSSalesArray(100000, 5)
        Dim Past30Days As Date = Now.AddDays(-30)


        SQLStatement = "Select * from [POSSalesView] where cast([Date Record] as date) >='" & Format(Past30Days, "dd-MMM-yyyy") & "' order by [Date Record] desc"
        records = erpdb.Execute(SQLStatement)
        Dim NumRecordsTemp As Integer = 0

        While Not records.EOF
            Try
                POSSalesArray(NumRecordsTemp, 0) = records.Fields("Total Net Sales").Value
                POSSalesArray(NumRecordsTemp, 1) = records.Fields("Total Gross Sales").Value
                POSSalesArray(NumRecordsTemp, 2) = records.Fields("Outlet Name").Value
                POSSalesArray(NumRecordsTemp, 3) = records.Fields("Date Record").Value
                POSSalesArray(NumRecordsTemp, 4) = records.Fields("Quantity Sold").Value
                POSSalesArray(NumRecordsTemp, 5) = Format(records.Fields("Date Record").Value, "HH")

            Catch ex As Exception
                '     InventoryValue = 0
            End Try
            NumRecordsTemp += 1
            records.MoveNext()
        End While
        records.Close()

        'gets the number of outlets
        SQLStatement = "Select distinct [Outlet Name] from [POSSalesView] where cast([Date Record] as date) >='" & Format(Past30Days, "dd-MMM-yyyy") & "'"
        records = erpdb.Execute(SQLStatement)
        Dim NumOutletsTemp As Integer = 0
        Dim OutletsArray(200, 50)
        While Not records.EOF
            OutletsArray(NumOutletsTemp, 0) = records.Fields("Outlet Name").Value
            OutletsArray(NumOutletsTemp, 49) = records.Fields("Outlet Name").Value
            NumOutletsTemp += 1
            records.MoveNext()
        End While
        records.Close()


        SQLStatement = " select ItemName, sum([Total Net Sales]) as SumTotalNetSales, sum([Quantity Sold]) as sumqty from [POSSalesView] where  cast([Date Record] as date) >='" &
            "01" & Format(Now.AddMonths(-1), "-MMM-yyyy") & "' and   cast([Date Record] as date) <='" & Format(Now.AddMonths(-1), "dd-MMM-yyyy") & "' group by ItemName order by SumTotalNetSales desc"
        records = erpdb.Execute(SQLStatement)

        Dim TotalMonthlySalesLastMonth As Decimal = 0
        While Not records.EOF
            Try
                TotalMonthlySalesLastMonth = TotalMonthlySalesLastMonth + records.Fields("SumTotalNetSales").Value
            Catch ex As Exception
                '     InventoryValue = 0
            End Try
            records.MoveNext()
        End While
        records.Close()

        Dim GroupHours(24, 2)
        'now does sales by hour
        For counter01 = 0 To NumOutletsTemp - 1

            For counter02 = 0 To NumRecordsTemp - 1
                If POSSalesArray(counter02, 2) = OutletsArray(counter01, 0) Then
                    'here is outlet match
                    If POSSalesArray(counter02, 5) = "01" Then
                        OutletsArray(counter01, 1) = OutletsArray(counter01, 1) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "02" Then
                        OutletsArray(counter01, 2) = OutletsArray(counter01, 2) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "03" Then
                        OutletsArray(counter01, 3) = OutletsArray(counter01, 3) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "04" Then
                        OutletsArray(counter01, 4) = OutletsArray(counter01, 4) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "05" Then
                        OutletsArray(counter01, 5) = OutletsArray(counter01, 5) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "06" Then
                        OutletsArray(counter01, 6) = OutletsArray(counter01, 6) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "07" Then
                        OutletsArray(counter01, 7) = OutletsArray(counter01, 7) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "08" Then
                        OutletsArray(counter01, 8) = OutletsArray(counter01, 8) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "09" Then
                        OutletsArray(counter01, 9) = OutletsArray(counter01, 9) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "10" Then
                        OutletsArray(counter01, 10) = OutletsArray(counter01, 10) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "11" Then
                        OutletsArray(counter01, 11) = OutletsArray(counter01, 11) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "12" Then
                        OutletsArray(counter01, 12) = OutletsArray(counter01, 12) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "13" Then
                        OutletsArray(counter01, 13) = OutletsArray(counter01, 13) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "14" Then
                        OutletsArray(counter01, 14) = OutletsArray(counter01, 14) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "15" Then
                        OutletsArray(counter01, 15) = OutletsArray(counter01, 15) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "16" Then
                        OutletsArray(counter01, 16) = OutletsArray(counter01, 16) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "17" Then
                        OutletsArray(counter01, 17) = OutletsArray(counter01, 17) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "18" Then
                        OutletsArray(counter01, 18) = OutletsArray(counter01, 18) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "19" Then
                        OutletsArray(counter01, 19) = OutletsArray(counter01, 19) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "20" Then
                        OutletsArray(counter01, 20) = OutletsArray(counter01, 20) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "21" Then
                        OutletsArray(counter01, 21) = OutletsArray(counter01, 21) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "22" Then
                        OutletsArray(counter01, 22) = OutletsArray(counter01, 22) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "23" Then
                        OutletsArray(counter01, 23) = OutletsArray(counter01, 23) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "00" Then
                        OutletsArray(counter01, 24) = OutletsArray(counter01, 24) + POSSalesArray(counter02, 0)
                    Else
                        MsgBox("no home")
                    End If
                End If

                'this does group
                'only do it once
                If counter01 = 0 Then
                    If POSSalesArray(counter02, 5) = "01" Then
                        GroupHours(1, 0) = GroupHours(1, 0) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "02" Then
                        GroupHours(2, 0) = GroupHours(2, 0) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "03" Then
                        GroupHours(3, 0) = GroupHours(3, 0) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "04" Then
                        GroupHours(4, 0) = GroupHours(4, 0) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "05" Then
                        GroupHours(5, 0) = GroupHours(5, 0) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "06" Then
                        GroupHours(6, 0) = GroupHours(6, 0) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "07" Then
                        GroupHours(7, 0) = GroupHours(7, 0) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "08" Then
                        GroupHours(8, 0) = GroupHours(8, 0) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "09" Then
                        GroupHours(9, 0) = GroupHours(9, 0) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "10" Then
                        GroupHours(10, 0) = GroupHours(10, 0) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "11" Then
                        GroupHours(11, 0) = GroupHours(11, 0) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "12" Then
                        GroupHours(12, 0) = GroupHours(12, 0) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "13" Then
                        GroupHours(13, 0) = GroupHours(13, 0) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "14" Then
                        GroupHours(14, 0) = GroupHours(14, 0) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "15" Then
                        GroupHours(15, 0) = GroupHours(15, 0) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "16" Then
                        GroupHours(16, 0) = GroupHours(16, 0) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "17" Then
                        GroupHours(17, 0) = GroupHours(17, 0) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "18" Then
                        GroupHours(18, 0) = GroupHours(18, 0) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "19" Then
                        GroupHours(19, 0) = GroupHours(19, 0) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "20" Then
                        GroupHours(20, 0) = GroupHours(20, 0) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "21" Then
                        GroupHours(21, 0) = GroupHours(21, 0) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "22" Then
                        GroupHours(22, 0) = GroupHours(22, 0) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "23" Then
                        GroupHours(23, 0) = GroupHours(23, 0) + POSSalesArray(counter02, 0)
                    ElseIf POSSalesArray(counter02, 5) = "00" Then
                        GroupHours(24, 0) = GroupHours(24, 0) + POSSalesArray(counter02, 0)
                    Else
                        MsgBox("no home")
                    End If
                End If

            Next
        Next


        'which is the busiest hour
        Dim HourWithMostSales As Decimal = 0
        Dim HourSales As Decimal = 0

        For counter = 0 To 24
            If Val(GroupHours(counter, 0)) > HourSales Then
                HourSales = Val(GroupHours(counter, 0))
                HourWithMostSales = counter
            End If
        Next
        Dim TimeofdayTemp As String = "morning"
        If HourWithMostSales > 12 And HourWithMostSales < 24 Then
            TimeofdayTemp = "pm"
            HourWithMostSales = HourWithMostSales - 12
        ElseIf HourWithMostSales = 24 Then
            TimeofdayTemp = "midnight"
            HourWithMostSales = HourWithMostSales - 12
        Else
            TimeofdayTemp = "am"
        End If




        'start of speech string
        SpeechString = "Your busiest sales hour is " & HourWithMostSales & " " & TimeofdayTemp & " with " & Format(HourSales, "$ #,###,##0") & " group sales in the past 30 days. "

        'start of sms
        SMSText = "Busiest sales hour is " & HourWithMostSales & " " & TimeofdayTemp & " with " & Format(HourSales, "$ #,###,##0") & " group sales for the past 30 days. "

        'gets the outlet names
        SQLStatement = "select * from [Outlet Details] where [Date Closed] is Null order by [Outlet Name]"
        records = erpdb.Execute(SQLStatement)
        Dim Outletdesc As String = ""
        While Not records.EOF
            OutletName = records.Fields("Outlet Name").Value
            Try
                Outletdesc = records.Fields("Description").Value
            Catch ex As Exception
                Outletdesc = OutletName
            End Try

            For counter = 0 To NumOutletsTemp - 1
                If OutletsArray(counter, 0) = OutletName Then
                    OutletsArray(counter, 49) = Outletdesc
                End If
            Next
            records.MoveNext()
        End While
        records.Close()


        'now classifies accoring to sales by color
        '5 colors
        'green  69EF6D
        'green olive  A4DA38
        'light orange  FFC9A3
        'light red   FBA18D
        'light blue is DDFDFC
        'light green  A3FFAA


        'work out the outlet maximums
        For counter01 = 0 To NumOutletsTemp - 1
            HourSales = 0
            For counter02 = 1 To 24
                If Val(OutletsArray(counter01, counter02)) > HourSales Then
                    HourSales = Val(OutletsArray(counter01, counter02))
                    HourWithMostSales = counter02
                End If
            Next

            If HourWithMostSales > 12 And HourWithMostSales < 24 Then
                TimeofdayTemp = "pm"
                HourWithMostSales = HourWithMostSales - 12
            ElseIf HourWithMostSales = 24 Then
                TimeofdayTemp = "midnight"
                HourWithMostSales = HourWithMostSales - 12
            Else
                TimeofdayTemp = "am"
            End If


            'puts the maximum sales per outlet
            OutletsArray(counter01, 50) = HourSales
            'now has the times for outlets
            If counter01 = 0 Then
                SpeechString = SpeechString & " Busiest hour for  " & OutletsArray(counter01, 49) & " is " & HourWithMostSales & " " & TimeofdayTemp & " with " & Format(HourSales, "$ #,###,##0") & ". "
                SMSText = SMSText & " For  " & OutletsArray(counter01, 0) & " is " & HourWithMostSales & " " & TimeofdayTemp & " with " & Format(HourSales, "$ #,###,##0") & ". "
            Else
                SpeechString = SpeechString & " " & OutletsArray(counter01, 49) & " is " & HourWithMostSales & " " & TimeofdayTemp & " with " & Format(HourSales, "$ #,###,##0") & ". "
                SMSText = SMSText & " For  " & OutletsArray(counter01, 49) & " is " & HourWithMostSales & " " & TimeofdayTemp & " with " & Format(HourSales, "$ #,###,##0") & ". "
            End If
        Next


        For counter = 0 To 24
            GroupHours(counter, 1) = "DDFDFC"
            'light green
            If Val(GroupHours(counter, 0)) > 0 Then GroupHours(counter, 1) = "A3FFAA"

            If Val(GroupHours(counter, 0)) > HourSales * 0.2 And Val(GroupHours(counter, 0)) <= HourSales * 0.4 Then
                GroupHours(counter, 1) = "69EF6D"
            End If
            If Val(GroupHours(counter, 0)) > HourSales * 0.4 And Val(GroupHours(counter, 0)) <= HourSales * 0.6 Then
                GroupHours(counter, 1) = "A4DA38"
            End If
            If Val(GroupHours(counter, 0)) > HourSales * 0.6 And Val(GroupHours(counter, 0)) <= HourSales * 0.8 Then
                GroupHours(counter, 1) = "FFC9A3"
            End If
            If Val(GroupHours(counter, 0)) > HourSales * 0.8 Then
                GroupHours(counter, 1) = "FBA18D"
            End If
        Next

        'now does outlets
        'maximum for the outlet is in array element 50
        For counter01 = 0 To NumOutletsTemp - 1
            HourSales = OutletsArray(counter01, 50)
            For counter02 = 1 To 24
                OutletsArray(counter01, counter02 + 24) = "DDFDFC"
                'light green
                If Val(OutletsArray(counter01, counter02)) > 0 Then OutletsArray(counter01, counter02 + 24) = "A3FFAA"

                If Val(OutletsArray(counter01, counter02)) > HourSales * 0.2 And Val(OutletsArray(counter01, counter02)) <= HourSales * 0.4 Then
                    OutletsArray(counter01, counter02 + 24) = "69EF6D"
                End If
                If Val(OutletsArray(counter01, counter02)) > HourSales * 0.4 And Val(OutletsArray(counter01, counter02)) <= HourSales * 0.6 Then
                    OutletsArray(counter01, counter02 + 24) = "A4DA38"
                End If
                If Val(OutletsArray(counter01, counter02)) > HourSales * 0.6 And Val(OutletsArray(counter01, counter02)) <= HourSales * 0.8 Then
                    OutletsArray(counter01, counter02 + 24) = "FFC9A3"
                End If
                If Val(OutletsArray(counter01, counter02)) > HourSales * 0.8 Then
                    OutletsArray(counter01, counter02 + 24) = "FBA18D"
                End If
            Next


        Next





        'Yesterday last week
        '      DailyReportHTMLEmailString = "      <div style = ""width:400px;"">
        '         <h1 style=""text-align:center""> <background-color: blue;>Testing </h1>
        '        <p  style =""text-align:right""><a href=""#"">sample link</a> </p>
        '       </div>"
        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "    <img src = ""http://aagilitycom.ipage.com/wp-content/uploads/2017/12/BizCoach_small.png"" alt=""BizCoach"" ;  ><br /><br />"

        '       DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<font face=""calibri""><b><u> TEST EMAIL - NOT FOR CUSTOMERS - " & FrmMain.ListBoxCustomerInfo.SelectedItem &
        '               " - " & FrmMain.ListBoxNotificationType.SelectedItem & "</b></u><br /><br />"


        Dim ExcelHeaderLine0401 As String = "Yesterday: "
        Dim ExcelHeaderLine0402 As String = Format(InventoryValue, "HK$ #,###,##0")
        Dim ExcelHeaderLine0403 As String = ""
        Dim ExcelHeaderLine0404 As String = ""
        Dim ExcelHeaderLine0405 As String = ""
        Dim ExcelHeaderLine0406 As String = ""

        Dim ExcelHeaderLine0501 As String = ""
        Dim ExcelHeaderLine0502 As String = ""
        Dim ExcelHeaderLine0503 As String = ""
        Dim ExcelHeaderLine0504 As String = ""
        Dim ExcelHeaderLine0505 As String = ""
        Dim ExcelHeaderLine0506 As String = ""

        Dim ItemGroup(500, 3)
        Dim NumGroups As Int16 = 0

        Dim VarianceLastMonth As Decimal = 0
        'can be divide by zero error
        Try
            VarianceLastMonth = TotalMonthlySalesToDate / TotalMonthlySalesLastMonth
        Catch ex As Exception
            VarianceLastMonth = 0
        End Try

        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "Sales since : " & Format(Past30Days, "ddd d MMM yyyy") & "<br /><br />" &
                            "Sales : " & Format(TotalMonthlySalesToDate, "HK$ #,###,##0") & ", last month : " & Format(TotalMonthlySalesLastMonth, "HK$ #,###,##0") &
                            ", variance to last month " & Format(VarianceLastMonth, "0.##") & "%" & "<br />" &
                            "<br />" & "<br />" &
                            "<table border=1><col width=""350""><col width=""150""><col width=""150""><col width=""150""><tr>" &
                            "<td><b><p style =""text-align:left"">Outlet</td>" &
                            "<td><b><p style =""text-align:center"">0am to 1am HK$</td>" &
                            "<td><b><p style =""text-align:center"">1 to 2am</td>" &
                            "<td><b><p style =""text-align:center"">2 to 3am</td>" &
                            "<td><b><p style =""text-align:center"">3 to 4am</td>" &
                            "<td><b><p style =""text-align:center"">4 to 5am</td>" &
                            "<td><b><p style =""text-align:center"">5 to 6am</td>" &
                            "<td><b><p style =""text-align:center"">6 to 7am</td>" &
                            "<td><b><p style =""text-align:center"">7 to 8am</td>" &
                            "<td><b><p style =""text-align:center"">8 to 9am</td>" &
                            "<td><b><p style =""text-align:center"">9 to 10am</td>" &
                            "<td><b><p style =""text-align:center"">10 to 11am</td>" &
                            "<td><b><p style =""text-align:center"">11 to 12pm</td>" &
                            "<td><b><p style =""text-align:center"">12am to 1pm</td>" &
                            "<td><b><p style =""text-align:center"">1 to 2pm</td>" &
                            "<td><b><p style =""text-align:center"">2 to 3pm</td>" &
                            "<td><b><p style =""text-align:center"">3 to 4pm</td>" &
                            "<td><b><p style =""text-align:center"">4 to 5pm</td>" &
                            "<td><b><p style =""text-align:center"">5 to 6pm</td>" &
                            "<td><b><p style =""text-align:center"">6 to 6pm</td>" &
                            "<td><b><p style =""text-align:center"">7 to 8pm</td>" &
                            "<td><b><p style =""text-align:center"">8 to 9pm</td>" &
                            "<td><b><p style =""text-align:center"">9 to 10pm</td>" &
                            "<td><b><p style =""text-align:center"">10 to 11pm</td>" &
                            "<td><b><p style =""text-align:center"">11 to 12am</td>" &
                            "</b></tr>"






        For counter01 = 0 To NumOutletsTemp - 1

            'can now build the line
            DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<tr><td>" & "<p style =""text-align:left"">" & OutletsArray(counter01, 0) & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 23) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 24)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 1)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 1) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 2)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 2) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 3)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 3) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 4)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 4) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 5)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 5) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 6)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 6) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 7)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 7) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 8)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 8) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 9)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 9) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 10)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 10) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 11)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 11) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 12)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 12) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 13)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 13) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 14)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 14) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 15)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 15) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 16)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 16) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 17)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 17) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 18)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 18) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 19)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 19) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 20)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 20) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 21)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 21) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 22)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 22) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 23)), "#,###,##0") & "</tr>"

        Next


        '  <td align=""center"" style=""border-radius 3px;"" bgcolor=""#e9703e""><a href=""https://www.surveymonkey.com/r/M2Q557H"" target=""_blank"" style=""font-size: 16px; font-family: Helvetica, Arial, sans-serif; color: #ffffff; text-decoration:none; text-decoration: none;border-radius: 3px; padding: 12px 18px; border: 1px solid #e9703e; display: inline-block;"">Take our survey &rarr;</a></td>" &

        'https://htmlcolorcodes.com/
        'light blue is DDFDFC

        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<tr><td>" & "<p style =""text-align:left"">" & "Group" & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(24, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(24, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(1, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(1, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(2, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(2, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(3, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(3, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(4, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(4, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(5, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(5, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(6, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(6, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(7, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(7, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(8, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(8, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(9, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(9, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(10, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(10, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(11, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(11, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(12, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(12, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(13, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(13, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(14, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(14, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(15, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(15, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(16, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(16, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(17, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(17, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(18, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(18, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(19, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(19, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(20, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(20, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(21, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(21, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(22, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(22, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(23, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(23, 0)), "#,###,##0") & "</tr>"


        DailyReportHTMLEmailString = "" & DailyReportHTMLEmailString & "</table></span><br /><hr /><br />"
        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<br />"




        'puts the bizcoach logo
        'this is desktop format so just a text string
        '"Yesterday: HK$ 1,014, last wk: HK$ 10,793"
        'http://aagilitycom.ipage.com/custview/20Jan1802glorystores.htm
        Dim CompanynameTemp As String = Companyname.Trim
        CompanynameTemp = Replace(CompanynameTemp, ".", "")
        CompanynameTemp = Replace(CompanynameTemp, " ", "")
        CompanynameTemp = Replace(CompanynameTemp, " ", "")
        CompanynameTemp = Replace(CompanynameTemp, "'", "")
        CompanynameTemp = CompanynameTemp.ToLower

        '      SMSText = SMSText & ControlChars.CrLf & "Total: $ " & Format(SalesYesterdayAll, "#,###,##0")

        UrlOfReport = "http://aagilitycom.ipage.com/custview/" & Format(DailyReportDate, "ddMMMyyyy") & "02" & CompanynameTemp & ".htm"
        DesktopText = "Yesterday: " & Format(InventoryValue, "HK$ #,###,##0") & ", last wk:" & Format(SalesYesterdayLastWeek, "HK$ #,###,##0")


        '      SpeechString = SpeechString & ". Total prediction for the next 7 days is " & Format(TotalPrediction7Days, "#,###") & "$ . "

        FrmMain.TextBoxSpeechText.Text = SpeechString.ToString.Trim

        'this creates the speech file

        AWSPollySpeak(SpeechString, "SalesByItemPOS")
        'this publishes the speech to the internet 
        PublishSpeechtoImpala(SpeechString, "SalesByItemPOS")
        'the file is bublished to    
        'http://aagilitycom.ipage.com/custview/    TextBoxPublishLocationName.text


        'closes the table
        '        DailyReportHTMLEmailString = "" & DailyReportHTMLEmailString & "</table></span><br /><hr /><br />"
        '      DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<br />"

        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "If you would like BizCoach to read your email click <a href=""http://aagilitycom.ipage.com/custview/" &
            FrmMain.TextBoxPublishLocationName.Text.Trim & """>here</a><br /><br />"


        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "Please give feedback To <a href=""mailto:support@impalacloud.com"" target=""_top"">support@impalacloud.com</a><br /><br />"

        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<br />"

        'this creates the excel
        If FrmMain.CheckBoxCreateExcel.Checked = True Then
            CreateExcelSpreadsheet(DailyReportDate, ExcelHeaderLine0401,
                                      ExcelHeaderLine0402,
        ExcelHeaderLine0403,
        ExcelHeaderLine0404,
        ExcelHeaderLine0405,
        ExcelHeaderLine0406,
        ExcelHeaderLine0501,
        ExcelHeaderLine0502,
        ExcelHeaderLine0503,
        ExcelHeaderLine0504,
        ExcelHeaderLine0505,
        ExcelHeaderLine0506)
        End If


        'publish to the web browser on the notifications page
        Dim document As System.Windows.Forms.HtmlDocument = FrmMain.WebBrowserNotification.Document
        FrmMain.WebBrowserNotification.DocumentText = DailyReportHTMLEmailString

    End Sub

    Public Sub InvoicesByHourPOS(ByRef DailyReportHTMLEmailString As String, ByRef DailyReportDate As Date,
                                                   ByVal DemoCompany As Boolean, ByRef UrlOfReport As String,
                                                   ByRef DesktopText As String, ByRef Companyname As String,
                                                   ByRef SMSText As String, ByRef SpeechString As String)

        'this is a daily report that is based upon yesterdays sales
        'it compares this to the monthly and last month figures

        'as fixed date report then can take off the date here
        DailyReportDate = Now.AddDays(-1)

        'dont use KPIs now as may be a timing issue

        Dim erpdb As ADODB.Connection
        Dim records As ADODB._Recordset
        Dim recordsLookup As ADODB._Recordset
        Dim RecordsOutletLookup As ADODB._Recordset
        erpdb = New ADODB.Connection
        erpdb.Open(strconnectSQLBizCoachCustomerDataSQL)
        Dim InventoryValue As Decimal = 0
        Dim SalesMTDYesterday As Integer = 0
        Dim OutletName As String = ""
        Dim SalesYesterdayLastWeek As Integer = 0
        Dim SalesLastMonthVariance As Decimal = 0
        Dim SalesYesterdayLastWeekVariance As Decimal = 0
        Dim SalesLastMonthToDate As Integer = 0
        Dim StartDateFrom01 As Date = Now.ToShortDateString
        Dim StartDateTo01 As Date = Now.ToShortDateString
        Dim SQLStatement As String = ""
        Dim OutletFullNameTemp As String = ""
        Dim counter01 As Int16 = 0
        'this stores the excel data
        Dim ExcelTableData(10000, 10)





        Dim DemoParameter As Int16 = 0
        'uses UTC time so need to take away 8 hours or use HongKongLoginMoment

        If DemoCompany = True Then
            DemoParameter = 365
            DailyReportDate = DailyReportDate.AddYears(-1)
        End If
        DailyReportDate = DailyReportDate.ToShortDateString

        '     records = erpdb.Execute("Select * from [KPI Values] where [KPI Name] ='POS Daily Sales (yesterday) - Total'")

        '   Dim qtyOnhand As Decimal = 0
        '    Dim CostOfInv As Decimal = 0

        SQLStatement = " select ItemName, sum([Total Net Sales]) as SumTotalNetSales, sum([Quantity Sold]) as sumqty from [POSSalesView] where cast([Date Record] 
            as date) >='" & "01" & Format(Now, "-MMM-yyyy") & "' group by ItemName order by SumTotalNetSales desc"
        records = erpdb.Execute(SQLStatement)
        Dim TotalMonthlySalesToDate As Decimal = 0
        While Not records.EOF
            Try
                TotalMonthlySalesToDate = TotalMonthlySalesToDate + records.Fields("SumTotalNetSales").Value
            Catch ex As Exception
                '     InventoryValue = 0
            End Try
            records.MoveNext()
        End While
        records.Close()

        'this is loading all the sales into an array for the current month
        Dim POSSalesArray(100000, 7)
        Dim Past30Days As Date = Now.AddDays(-30)


        SQLStatement = "Select * from [POSSalesView] where cast([Date Record] as date) >='" & Format(Past30Days, "dd-MMM-yyyy") & "' order by [Reference Number] desc"


        '        SQLStatement = "   Select Case distinct([Reference Number]) As NumRecs, [Outlet Name] from [POSSalesView] where cast([Date Record] As Date) >='" & Format(Past30Days, "dd-MMM-yyyy") & "' 
        '        group by [Reference Number] , [Outlet Name] order by [Reference Number]  desc"

        'loads all the data first
        'does it this way as the date may be different for each line
        records = erpdb.Execute(SQLStatement)
        Dim NumRecordsTemp As Integer = 0

        While Not records.EOF
            Try
                POSSalesArray(NumRecordsTemp, 0) = records.Fields("Total Net Sales").Value
                POSSalesArray(NumRecordsTemp, 1) = records.Fields("Total Gross Sales").Value
                POSSalesArray(NumRecordsTemp, 2) = records.Fields("Outlet Name").Value
                POSSalesArray(NumRecordsTemp, 3) = records.Fields("Date Record").Value
                POSSalesArray(NumRecordsTemp, 4) = records.Fields("Quantity Sold").Value
                POSSalesArray(NumRecordsTemp, 5) = Format(records.Fields("Date Record").Value, "HH")
                POSSalesArray(NumRecordsTemp, 6) = records.Fields("Reference Number").Value

            Catch ex As Exception
                '     InventoryValue = 0
            End Try
            NumRecordsTemp += 1
            records.MoveNext()
        End While
        records.Close()


        'now sorts through the data and counts the distinct invoices
        Dim refNumTemp As String = ""
        Dim refNumOldTemp As String = ""

        For counter = 0 To NumRecordsTemp
            refNumTemp = POSSalesArray(counter, 6)
            If refNumTemp = refNumOldTemp Then
                POSSalesArray(counter, 7) = 0
            Else
                POSSalesArray(counter, 7) = 1
            End If


            refNumOldTemp = refNumTemp
        Next





        'gets the number of outlets
        SQLStatement = "Select distinct [Outlet Name] from [POSSalesView] where cast([Date Record] As Date) >='" & Format(Past30Days, "dd-MMM-yyyy") & "'"
        records = erpdb.Execute(SQLStatement)
        Dim NumOutletsTemp As Integer = 0
        Dim OutletsArray(200, 50)
        While Not records.EOF
            OutletsArray(NumOutletsTemp, 0) = records.Fields("Outlet Name").Value
            OutletsArray(NumOutletsTemp, 49) = records.Fields("Outlet Name").Value
            NumOutletsTemp += 1
            records.MoveNext()
        End While
        records.Close()


        SQLStatement = " select ItemName, sum([Total Net Sales]) as SumTotalNetSales, sum([Quantity Sold]) as sumqty from [POSSalesView] where  cast([Date Record] as date) >='" &
            "01" & Format(Now.AddMonths(-1), "-MMM-yyyy") & "' and   cast([Date Record] as date) <='" & Format(Now.AddMonths(-1), "dd-MMM-yyyy") & "' group by ItemName order by SumTotalNetSales desc"
        records = erpdb.Execute(SQLStatement)

        Dim TotalMonthlySalesLastMonth As Decimal = 0
        While Not records.EOF
            Try
                TotalMonthlySalesLastMonth = TotalMonthlySalesLastMonth + records.Fields("SumTotalNetSales").Value
            Catch ex As Exception
                '     InventoryValue = 0
            End Try
            records.MoveNext()
        End While
        records.Close()

        Dim GroupHours(24, 2)
        'now does sales by hour
        For counter01 = 0 To NumOutletsTemp - 1

            For counter02 = 0 To NumRecordsTemp - 1
                If POSSalesArray(counter02, 2) = OutletsArray(counter01, 0) Then
                    'here is outlet match
                    If POSSalesArray(counter02, 5) = "01" Then
                        OutletsArray(counter01, 1) = OutletsArray(counter01, 1) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "02" Then
                        OutletsArray(counter01, 2) = OutletsArray(counter01, 2) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "03" Then
                        OutletsArray(counter01, 3) = OutletsArray(counter01, 3) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "04" Then
                        OutletsArray(counter01, 4) = OutletsArray(counter01, 4) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "05" Then
                        OutletsArray(counter01, 5) = OutletsArray(counter01, 5) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "06" Then
                        OutletsArray(counter01, 6) = OutletsArray(counter01, 6) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "07" Then
                        OutletsArray(counter01, 7) = OutletsArray(counter01, 7) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "08" Then
                        OutletsArray(counter01, 8) = OutletsArray(counter01, 8) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "09" Then
                        OutletsArray(counter01, 9) = OutletsArray(counter01, 9) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "10" Then
                        OutletsArray(counter01, 10) = OutletsArray(counter01, 10) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "11" Then
                        OutletsArray(counter01, 11) = OutletsArray(counter01, 11) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "12" Then
                        OutletsArray(counter01, 12) = OutletsArray(counter01, 12) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "13" Then
                        OutletsArray(counter01, 13) = OutletsArray(counter01, 13) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "14" Then
                        OutletsArray(counter01, 14) = OutletsArray(counter01, 14) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "15" Then
                        OutletsArray(counter01, 15) = OutletsArray(counter01, 15) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "16" Then
                        OutletsArray(counter01, 16) = OutletsArray(counter01, 16) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "17" Then
                        OutletsArray(counter01, 17) = OutletsArray(counter01, 17) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "18" Then
                        OutletsArray(counter01, 18) = OutletsArray(counter01, 18) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "19" Then
                        OutletsArray(counter01, 19) = OutletsArray(counter01, 19) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "20" Then
                        OutletsArray(counter01, 20) = OutletsArray(counter01, 20) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "21" Then
                        OutletsArray(counter01, 21) = OutletsArray(counter01, 21) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "22" Then
                        OutletsArray(counter01, 22) = OutletsArray(counter01, 22) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "23" Then
                        OutletsArray(counter01, 23) = OutletsArray(counter01, 23) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "00" Then
                        OutletsArray(counter01, 24) = OutletsArray(counter01, 24) + POSSalesArray(counter02, 7)
                    Else
                        MsgBox("no home")
                    End If
                End If

                'this does group
                'only do it once
                If counter01 = 0 Then
                    If POSSalesArray(counter02, 5) = "01" Then
                        GroupHours(1, 0) = GroupHours(1, 0) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "02" Then
                        GroupHours(2, 0) = GroupHours(2, 0) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "03" Then
                        GroupHours(3, 0) = GroupHours(3, 0) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "04" Then
                        GroupHours(4, 0) = GroupHours(4, 0) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "05" Then
                        GroupHours(5, 0) = GroupHours(5, 0) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "06" Then
                        GroupHours(6, 0) = GroupHours(6, 0) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "07" Then
                        GroupHours(7, 0) = GroupHours(7, 0) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "08" Then
                        GroupHours(8, 0) = GroupHours(8, 0) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "09" Then
                        GroupHours(9, 0) = GroupHours(9, 0) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "10" Then
                        GroupHours(10, 0) = GroupHours(10, 0) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "11" Then
                        GroupHours(11, 0) = GroupHours(11, 0) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "12" Then
                        GroupHours(12, 0) = GroupHours(12, 0) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "13" Then
                        GroupHours(13, 0) = GroupHours(13, 0) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "14" Then
                        GroupHours(14, 0) = GroupHours(14, 0) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "15" Then
                        GroupHours(15, 0) = GroupHours(15, 0) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "16" Then
                        GroupHours(16, 0) = GroupHours(16, 0) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "17" Then
                        GroupHours(17, 0) = GroupHours(17, 0) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "18" Then
                        GroupHours(18, 0) = GroupHours(18, 0) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "19" Then
                        GroupHours(19, 0) = GroupHours(19, 0) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "20" Then
                        GroupHours(20, 0) = GroupHours(20, 0) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "21" Then
                        GroupHours(21, 0) = GroupHours(21, 0) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "22" Then
                        GroupHours(22, 0) = GroupHours(22, 0) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "23" Then
                        GroupHours(23, 0) = GroupHours(23, 0) + POSSalesArray(counter02, 7)
                    ElseIf POSSalesArray(counter02, 5) = "00" Then
                        GroupHours(24, 0) = GroupHours(24, 0) + POSSalesArray(counter02, 7)
                    Else
                        MsgBox("no home")
                    End If
                End If

            Next
        Next


        'which is the busiest hour
        Dim HourWithMostSales As Decimal = 0
        Dim HourSales As Decimal = 0

        For counter = 0 To 24
            If Val(GroupHours(counter, 0)) > HourSales Then
                HourSales = Val(GroupHours(counter, 0))
                HourWithMostSales = counter
            End If
        Next
        Dim TimeofdayTemp As String = "morning"
        If HourWithMostSales > 12 And HourWithMostSales < 24 Then
            TimeofdayTemp = "pm"
            HourWithMostSales = HourWithMostSales - 12
        ElseIf HourWithMostSales = 24 Then
            TimeofdayTemp = "midnight"
            HourWithMostSales = HourWithMostSales - 12
        Else
            TimeofdayTemp = "am"
        End If




        'start of speech string
        SpeechString = "Your busiest sales hour is " & HourWithMostSales & " " & TimeofdayTemp & " with " & Format(HourSales, "#,###,##0") & " customers in the past 30 days. "

        'start of sms
        SMSText = "Busiest sales hour is " & HourWithMostSales & " " & TimeofdayTemp & " with " & Format(HourSales, "#,###,##0") & " customers in the past 30 days. "

        'gets the outlet names
        SQLStatement = "select * from [Outlet Details] where [Date Closed] is Null order by [Outlet Name]"
        records = erpdb.Execute(SQLStatement)
        Dim Outletdesc As String = ""
        While Not records.EOF
            OutletName = records.Fields("Outlet Name").Value
            Try
                Outletdesc = records.Fields("Description").Value
            Catch ex As Exception
                Outletdesc = OutletName
            End Try

            For counter = 0 To NumOutletsTemp - 1
                If OutletsArray(counter, 0) = OutletName Then
                    OutletsArray(counter, 49) = Outletdesc
                End If
            Next
            records.MoveNext()
        End While
        records.Close()


        'now classifies accoring to sales by color
        '5 colors
        'green  69EF6D
        'green olive  A4DA38
        'light orange  FFC9A3
        'light red   FBA18D
        'light blue is DDFDFC
        'light green  A3FFAA


        'work out the outlet maximums
        For counter01 = 0 To NumOutletsTemp - 1
            HourSales = 0
            For counter02 = 1 To 24
                If Val(OutletsArray(counter01, counter02)) > HourSales Then
                    HourSales = Val(OutletsArray(counter01, counter02))
                    HourWithMostSales = counter02
                End If
            Next

            If HourWithMostSales > 12 And HourWithMostSales < 24 Then
                TimeofdayTemp = "pm"
                HourWithMostSales = HourWithMostSales - 12
            ElseIf HourWithMostSales = 24 Then
                TimeofdayTemp = "midnight"
                HourWithMostSales = HourWithMostSales - 12
            Else
                TimeofdayTemp = "am"
            End If


            'puts the maximum sales per outlet
            OutletsArray(counter01, 50) = HourSales
            'now has the times for outlets
            If counter01 = 0 Then
                SpeechString = SpeechString & " Busiest hour for  " & OutletsArray(counter01, 49) & " is " & HourWithMostSales & " " & TimeofdayTemp & " with " & Format(HourSales, "#,###,##0") & " customers. "
                SMSText = SMSText & " For  " & OutletsArray(counter01, 0) & " is " & HourWithMostSales & " " & TimeofdayTemp & " with " & Format(HourSales, "#,###,##0") & " customers. "
            Else
                SpeechString = SpeechString & " " & OutletsArray(counter01, 49) & " is " & HourWithMostSales & " " & TimeofdayTemp & " with " & Format(HourSales, "#,###,##0") & ". "
                SMSText = SMSText & " For  " & OutletsArray(counter01, 49) & " is " & HourWithMostSales & " " & TimeofdayTemp & " with " & Format(HourSales, "#,###,##0") & ". "
            End If
        Next


        For counter = 0 To 24
            GroupHours(counter, 1) = "DDFDFC"
            'light green
            If Val(GroupHours(counter, 0)) > 0 Then GroupHours(counter, 1) = "A3FFAA"

            If Val(GroupHours(counter, 0)) > HourSales * 0.2 And Val(GroupHours(counter, 0)) <= HourSales * 0.4 Then
                GroupHours(counter, 1) = "69EF6D"
            End If
            If Val(GroupHours(counter, 0)) > HourSales * 0.4 And Val(GroupHours(counter, 0)) <= HourSales * 0.6 Then
                GroupHours(counter, 1) = "A4DA38"
            End If
            If Val(GroupHours(counter, 0)) > HourSales * 0.6 And Val(GroupHours(counter, 0)) <= HourSales * 0.8 Then
                GroupHours(counter, 1) = "FFC9A3"
            End If
            If Val(GroupHours(counter, 0)) > HourSales * 0.8 Then
                GroupHours(counter, 1) = "FBA18D"
            End If
        Next

        'now does outlets
        'maximum for the outlet is in array element 50
        For counter01 = 0 To NumOutletsTemp - 1
            HourSales = OutletsArray(counter01, 50)
            For counter02 = 1 To 24
                OutletsArray(counter01, counter02 + 24) = "DDFDFC"
                'light green
                If Val(OutletsArray(counter01, counter02)) > 0 Then OutletsArray(counter01, counter02 + 24) = "A3FFAA"

                If Val(OutletsArray(counter01, counter02)) > HourSales * 0.2 And Val(OutletsArray(counter01, counter02)) <= HourSales * 0.4 Then
                    OutletsArray(counter01, counter02 + 24) = "69EF6D"
                End If
                If Val(OutletsArray(counter01, counter02)) > HourSales * 0.4 And Val(OutletsArray(counter01, counter02)) <= HourSales * 0.6 Then
                    OutletsArray(counter01, counter02 + 24) = "A4DA38"
                End If
                If Val(OutletsArray(counter01, counter02)) > HourSales * 0.6 And Val(OutletsArray(counter01, counter02)) <= HourSales * 0.8 Then
                    OutletsArray(counter01, counter02 + 24) = "FFC9A3"
                End If
                If Val(OutletsArray(counter01, counter02)) > HourSales * 0.8 Then
                    OutletsArray(counter01, counter02 + 24) = "FBA18D"
                End If
            Next


        Next





        'Yesterday last week
        '      DailyReportHTMLEmailString = "      <div style = ""width:400px;"">
        '         <h1 style=""text-align:center""> <background-color: blue;>Testing </h1>
        '        <p  style =""text-align:right""><a href=""#"">sample link</a> </p>
        '       </div>"
        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "    <img src = ""http://aagilitycom.ipage.com/wp-content/uploads/2017/12/BizCoach_small.png"" alt=""BizCoach"" ;  ><br /><br />"

        '   DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<font face=""calibri""><b><u> TEST EMAIL - NOT FOR CUSTOMERS - " & FrmMain.ListBoxCustomerInfo.SelectedItem &
        '              " - " & FrmMain.ListBoxNotificationType.SelectedItem & "</b></u><br /><br />"


        Dim ExcelHeaderLine0401 As String = "Yesterday: "
        Dim ExcelHeaderLine0402 As String = Format(InventoryValue, "HK$ #,###,##0")
        Dim ExcelHeaderLine0403 As String = ""
        Dim ExcelHeaderLine0404 As String = ""
        Dim ExcelHeaderLine0405 As String = ""
        Dim ExcelHeaderLine0406 As String = ""

        Dim ExcelHeaderLine0501 As String = ""
        Dim ExcelHeaderLine0502 As String = ""
        Dim ExcelHeaderLine0503 As String = ""
        Dim ExcelHeaderLine0504 As String = ""
        Dim ExcelHeaderLine0505 As String = ""
        Dim ExcelHeaderLine0506 As String = ""

        Dim ItemGroup(500, 3)
        Dim NumGroups As Int16 = 0

        Dim VarianceLastMonth As Decimal = 0
        'can be divide by zero error
        Try
            VarianceLastMonth = TotalMonthlySalesToDate / TotalMonthlySalesLastMonth
        Catch ex As Exception
            VarianceLastMonth = 0
        End Try

        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "Customers since : " & Format(Past30Days, "ddd d MMM yyyy") & "<br /><br />" &
                            "Sales : " & Format(TotalMonthlySalesToDate, "HK$ #,###,##0") & ", last month : " & Format(TotalMonthlySalesLastMonth, "HK$ #,###,##0") &
                            ", variance to last month " & Format(VarianceLastMonth, "0.##") & "%" & "<br />" &
                            "<br />" & "<br />" &
                            "<table border=1><col width=""350""><col width=""150""><col width=""150""><col width=""150""><tr>" &
                            "<td><b><p style =""text-align:left"">Outlet</td>" &
                            "<td><b><p style =""text-align:center"">0am to 1am</td>" &
                            "<td><b><p style =""text-align:center"">1 to 2am</td>" &
                            "<td><b><p style =""text-align:center"">2 to 3am</td>" &
                            "<td><b><p style =""text-align:center"">3 to 4am</td>" &
                            "<td><b><p style =""text-align:center"">4 to 5am</td>" &
                            "<td><b><p style =""text-align:center"">5 to 6am</td>" &
                            "<td><b><p style =""text-align:center"">6 to 7am</td>" &
                            "<td><b><p style =""text-align:center"">7 to 8am</td>" &
                            "<td><b><p style =""text-align:center"">8 to 9am</td>" &
                            "<td><b><p style =""text-align:center"">9 to 10am</td>" &
                            "<td><b><p style =""text-align:center"">10 to 11am</td>" &
                            "<td><b><p style =""text-align:center"">11 to 12pm</td>" &
                            "<td><b><p style =""text-align:center"">12am to 1pm</td>" &
                            "<td><b><p style =""text-align:center"">1 to 2pm</td>" &
                            "<td><b><p style =""text-align:center"">2 to 3pm</td>" &
                            "<td><b><p style =""text-align:center"">3 to 4pm</td>" &
                            "<td><b><p style =""text-align:center"">4 to 5pm</td>" &
                            "<td><b><p style =""text-align:center"">5 to 6pm</td>" &
                            "<td><b><p style =""text-align:center"">6 to 6pm</td>" &
                            "<td><b><p style =""text-align:center"">7 to 8pm</td>" &
                            "<td><b><p style =""text-align:center"">8 to 9pm</td>" &
                            "<td><b><p style =""text-align:center"">9 to 10pm</td>" &
                            "<td><b><p style =""text-align:center"">10 to 11pm</td>" &
                            "<td><b><p style =""text-align:center"">11 to 12am</td>" &
                            "</b></tr>"



        For counter01 = 0 To NumOutletsTemp - 1

            'can now build the line
            DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<tr><td>" & "<p style =""text-align:left"">" & OutletsArray(counter01, 0) & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 23) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 24)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 1)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 1) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 2)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 2) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 3)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 3) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 4)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 4) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 5)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 5) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 6)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 6) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 7)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 7) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 8)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 8) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 9)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 9) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 10)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 10) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 11)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 11) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 12)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 12) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 13)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 13) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 14)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 14) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 15)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 15) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 16)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 16) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 17)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 17) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 18)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 18) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 19)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 19) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 20)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 20) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 21)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 21) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 22)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & OutletsArray(counter01, 25 + 22) & """>" & "<p style =""text-align:right"">" & Format(Val(OutletsArray(counter01, 23)), "#,###,##0") & "</tr>"

        Next


        '  <td align=""center"" style=""border-radius 3px;"" bgcolor=""#e9703e""><a href=""https://www.surveymonkey.com/r/M2Q557H"" target=""_blank"" style=""font-size: 16px; font-family: Helvetica, Arial, sans-serif; color: #ffffff; text-decoration:none; text-decoration: none;border-radius: 3px; padding: 12px 18px; border: 1px solid #e9703e; display: inline-block;"">Take our survey &rarr;</a></td>" &

        'https://htmlcolorcodes.com/
        'light blue is DDFDFC

        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<tr><td>" & "<p style =""text-align:left"">" & "Group" & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(24, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(24, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(1, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(1, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(2, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(2, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(3, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(3, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(4, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(4, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(5, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(5, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(6, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(6, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(7, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(7, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(8, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(8, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(9, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(9, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(10, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(10, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(11, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(11, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(12, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(12, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(13, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(13, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(14, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(14, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(15, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(15, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(16, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(16, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(17, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(17, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(18, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(18, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(19, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(19, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(20, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(20, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(21, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(21, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(22, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(22, 0)), "#,###,##0") & "</td>" &
                                            "<td bgcolor=""#" & GroupHours(23, 1) & """>" & "<p style =""text-align:right"">" & Format(Val(GroupHours(23, 0)), "#,###,##0") & "</tr>"


        DailyReportHTMLEmailString = "" & DailyReportHTMLEmailString & "</table></span><br /><hr /><br />"
        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<br />"




        'puts the bizcoach logo
        'this is desktop format so just a text string
        '"Yesterday: HK$ 1,014, last wk: HK$ 10,793"
        'http://aagilitycom.ipage.com/custview/20Jan1802glorystores.htm
        Dim CompanynameTemp As String = Companyname.Trim
        CompanynameTemp = Replace(CompanynameTemp, ".", "")
        CompanynameTemp = Replace(CompanynameTemp, " ", "")
        CompanynameTemp = Replace(CompanynameTemp, " ", "")
        CompanynameTemp = Replace(CompanynameTemp, "'", "")
        CompanynameTemp = CompanynameTemp.ToLower

        '      SMSText = SMSText & ControlChars.CrLf & "Total: $ " & Format(SalesYesterdayAll, "#,###,##0")

        UrlOfReport = "http://aagilitycom.ipage.com/custview/" & Format(DailyReportDate, "ddMMMyyyy") & "02" & CompanynameTemp & ".htm"
        DesktopText = "Yesterday: " & Format(InventoryValue, "HK$ #,###,##0") & ", last wk:" & Format(SalesYesterdayLastWeek, "HK$ #,###,##0")


        '      SpeechString = SpeechString & ". Total prediction for the next 7 days is " & Format(TotalPrediction7Days, "#,###") & "$ . "

        FrmMain.TextBoxSpeechText.Text = SpeechString.ToString.Trim

        'this creates the speech file

        AWSPollySpeak(SpeechString, "SalesByItemPOS")
        'this publishes the speech to the internet 
        PublishSpeechtoImpala(SpeechString, "SalesByItemPOS")
        'the file is bublished to    
        'http://aagilitycom.ipage.com/custview/    TextBoxPublishLocationName.text


        'closes the table
        '    DailyReportHTMLEmailString = "" & DailyReportHTMLEmailString & "</table></span><br /><hr /><br />"
        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<br />"

        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "If you would like BizCoach to read your email click <a href=""http://aagilitycom.ipage.com/custview/" &
            FrmMain.TextBoxPublishLocationName.Text.Trim & """>here</a><br /><br />"


        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "Please give feedback To <a href=""mailto:support@impalacloud.com"" target=""_top"">support@impalacloud.com</a><br /><br />"

        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<br />"

        'this creates the excel
        If FrmMain.CheckBoxCreateExcel.Checked = True Then
            CreateExcelSpreadsheet(DailyReportDate, ExcelHeaderLine0401,
                                      ExcelHeaderLine0402,
        ExcelHeaderLine0403,
        ExcelHeaderLine0404,
        ExcelHeaderLine0405,
        ExcelHeaderLine0406,
        ExcelHeaderLine0501,
        ExcelHeaderLine0502,
        ExcelHeaderLine0503,
        ExcelHeaderLine0504,
        ExcelHeaderLine0505,
        ExcelHeaderLine0506)
        End If


        'publish to the web browser on the notifications page
        Dim document As System.Windows.Forms.HtmlDocument = FrmMain.WebBrowserNotification.Document
        FrmMain.WebBrowserNotification.DocumentText = DailyReportHTMLEmailString

    End Sub





    Public Sub InventoryTurnoverPOS(ByRef DailyReportHTMLEmailString As String, ByRef DailyReportDate As Date,
                                                   ByVal DemoCompany As Boolean, ByRef UrlOfReport As String,
                                                   ByRef DesktopText As String, ByRef Companyname As String,
                                                   ByRef SMSText As String, ByRef SpeechString As String)

        'this is a daily report that is based upon yesterdays sales
        'it compares this to the monthly and last month figures

        'as fixed date report then can take off the date here
        DailyReportDate = Now.AddDays(-1)

        'dont use KPIs now as may be a timing issue

        Dim erpdb As ADODB.Connection
        Dim records As ADODB._Recordset
        Dim recordsLookup As ADODB._Recordset
        Dim RecordsOutletLookup As ADODB._Recordset
        erpdb = New ADODB.Connection
        erpdb.Open(strconnectSQLBizCoachCustomerDataSQL)
        Dim InventoryValue As Decimal = 0
        Dim SalesMTDYesterday As Integer = 0
        Dim OutletName As String = ""
        Dim SalesYesterdayLastWeek As Integer = 0
        Dim SalesLastMonthVariance As Decimal = 0
        Dim SalesYesterdayLastWeekVariance As Decimal = 0
        Dim SalesLastMonthToDate As Integer = 0
        Dim StartDateFrom01 As Date = Now.ToShortDateString
        Dim StartDateTo01 As Date = Now.ToShortDateString
        Dim SQLStatement As String = ""
        Dim OutletFullNameTemp As String = ""
        Dim counter01 As Int16 = 0
        'this stores the excel data
        Dim ExcelTableData(10000, 10)

        'some definitions
        'https://www.investopedia.com/terms/d/dsi.asp

        SMSText = ""
        SpeechString = "Inventory turnover is "

        Dim DemoParameter As Int16 = 0
        'uses UTC time so need to take away 8 hours or use HongKongLoginMoment

        DailyReportDate = DailyReportDate.ToShortDateString

        SQLStatement = " select ItemName, sum([Total Net Sales]) as SumTotalNetSales, sum([Quantity Sold]) as sumqty from [POSSalesView] where cast([Date Record] as date) >='" & "01" & Format(Now, "-MMM-yyyy") & "' group by ItemName order by SumTotalNetSales desc"
        records = erpdb.Execute(SQLStatement)

        Dim TotalMonthlySalesToDate As Decimal = 0

        While Not records.EOF
            Try
                TotalMonthlySalesToDate = TotalMonthlySalesToDate + records.Fields("SumTotalNetSales").Value
            Catch ex As Exception
                '     InventoryValue = 0
            End Try
            records.MoveNext()
        End While
        records.Close()

        SQLStatement = " select ItemName, sum([Total Net Sales]) as SumTotalNetSales, sum([Quantity Sold]) as sumqty from [POSSalesView] where  cast([Date Record] as date) >='" &
            "01" & Format(Now.AddMonths(-1), "-MMM-yyyy") & "' and   cast([Date Record] as date) <='" & Format(Now.AddMonths(-1), "dd-MMM-yyyy") & "' group by ItemName order by SumTotalNetSales desc"
        records = erpdb.Execute(SQLStatement)

        Dim TotalMonthlySalesLastMonth As Decimal = 0
        While Not records.EOF
            Try
                TotalMonthlySalesLastMonth = TotalMonthlySalesLastMonth + records.Fields("SumTotalNetSales").Value
            Catch ex As Exception
                '     InventoryValue = 0
            End Try
            records.MoveNext()
        End While
        records.Close()


        SpeechString = SpeechString & Format(TotalMonthlySalesToDate, "#,###") & "$" & " for your business this month, last month was " & Format(TotalMonthlySalesLastMonth, "$ #,###,##0") & ". "



        'Yesterday last week
        '      DailyReportHTMLEmailString = "      <div style = ""width:400px;"">
        '         <h1 style=""text-align:center""> <background-color: blue;>Testing </h1>
        '        <p  style =""text-align:right""><a href=""#"">sample link</a> </p>
        '       </div>"
        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "    <img src = ""http://aagilitycom.ipage.com/wp-content/uploads/2017/12/BizCoach_small.png"" alt=""BizCoach"" ;  ><br /><br />"

        '       DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<font face=""calibri""><b><u> TEST EMAIL - NOT FOR CUSTOMERS - " & FrmMain.ListBoxCustomerInfo.SelectedItem &
        '               " - " & FrmMain.ListBoxNotificationType.SelectedItem & "</b></u><br /><br />"


        Dim ExcelHeaderLine0401 As String = "Yesterday: "
        Dim ExcelHeaderLine0402 As String = Format(InventoryValue, "HK$ #,###,##0")
        Dim ExcelHeaderLine0403 As String = ""
        Dim ExcelHeaderLine0404 As String = ""
        Dim ExcelHeaderLine0405 As String = ""
        Dim ExcelHeaderLine0406 As String = ""

        Dim ExcelHeaderLine0501 As String = ""
        Dim ExcelHeaderLine0502 As String = ""
        Dim ExcelHeaderLine0503 As String = ""
        Dim ExcelHeaderLine0504 As String = ""
        Dim ExcelHeaderLine0505 As String = ""
        Dim ExcelHeaderLine0506 As String = ""

        Dim ItemGroup(500, 3)
        Dim NumGroups As Int16 = 0

        Dim VarianceLastMonth As Decimal = 0
        'can be divide by zero error
        Try
            VarianceLastMonth = TotalMonthlySalesToDate / TotalMonthlySalesLastMonth
        Catch ex As Exception
            VarianceLastMonth = 0
        End Try

        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "Sales by item For : " & Format(DailyReportDate, "ddd d MMM yyyy") & "<br /><br />" &
                             "Monthly sales : " & Format(TotalMonthlySalesToDate, "HK$ #,###,##0") & ", last month : " & Format(TotalMonthlySalesLastMonth, "HK$ #,###,##0") &
                             ", variance to last month " & Format(VarianceLastMonth, "0.##") & "%" & "<br />" &
                              "<br />" & "<br />" &
                             "<table border=1><col width=""350""><col width=""150""><col width=""150""><col width=""150""><tr>" &
                             "<td><b><p style =""text-align:left"">Item Name</td>" &
                             "<td><b><p style =""text-align:center"">Total Sales HK$</td>" &
                            "<td><b><p style =""text-align:center"">Quantity Sold</td>" &
                              "</b></tr>"




        SMSText = "Inventory for : " & Format(DailyReportDate.AddDays(1), "ddd d MMM") & ControlChars.CrLf
        SMSText = SMSText & Format(InventoryValue, "#,###,###") & " $. " & ControlChars.CrLf


        'gets the item groups
        SQLStatement = " select top(10)  sum([Total Net Sales]) as SumTotalNetSales, sum([Quantity Sold]) as sumqty, ItemName from [POSSalesView] where  cast([Date Record] as date) >='" &
            "01" & Format(Now, "-MMM-yyyy") & "' group by  ItemName order by SumTotalNetSales desc"
        records = erpdb.Execute(SQLStatement)
        '   cast([Date Record] as date)

        counter01 = 0
        Dim ItemNameTemp As String = ""
        Dim SumTotalNetSalesTemp As Decimal = 0
        Dim sumqtyTemp As Decimal = 0

        While Not records.EOF
            Try
                ItemNameTemp = ""
                ItemNameTemp = records.Fields("ItemName").Value

            Catch ex As Exception

            End Try
            SumTotalNetSalesTemp = 0
            SumTotalNetSalesTemp = records.Fields("SumTotalNetSales").Value
            sumqtyTemp = 0
            sumqtyTemp = records.Fields("sumqty").Value

            'can now build the line
            DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<tr><td>" & "<p style =""text-align:left"">" & ItemNameTemp & "</td>" &
                                            "<td>" & "<p style =""text-align:right"">" & Format(SumTotalNetSalesTemp, "#,###,###") & "</td>" &
                                            "<td>" & "<p style =""text-align:right"">" & Format(sumqtyTemp, "#,###,###") & "</tr>"

            If counter01 = 0 Then
                SpeechString = SpeechString & " " & ItemNameTemp & " has sales of " & Format(SumTotalNetSalesTemp, "#,###") & "$ with " & Format(sumqtyTemp, "#,###,###") & " sold. "
            Else
                SpeechString = SpeechString & " " & ItemNameTemp & " sales " & Format(SumTotalNetSalesTemp, "#,###") & "$ , " & Format(sumqtyTemp, "#,###,###") & " sold. "
            End If
            counter01 += 1

            records.MoveNext()
        End While
        records.Close()


        DailyReportHTMLEmailString = "" & DailyReportHTMLEmailString & "</table></span><br /><hr /><br />"
        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<br />"


        'sales by item
        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "Sales by Item by Outlet : " &
                                         "<br />" & "<br />" &
                                        "<table border=1><col width=""150""><col width=""350""><col width=""150""><col width=""150""><tr>" &
                                        "<td><b><p style =""text-align:center"">Outlet</td>" &
                                        "<td><b><p style =""text-align:left"">Item Name </td>" &
                                        "<td><b><p style =""text-align:center"">Total Sales HK$</td>" &
                                       "<td><b><p style =""text-align:center"">Quantity Sold</td>" &
                                         "</b></tr>"



        Dim OutletNameTemp As String = ""

        'gets the item groups
        SQLStatement = " Select top(20) [Outlet Name], ItemName, sum([Total Net Sales]) As SumTotalNetSales, sum([Quantity Sold]) As sumqty  from [POSSalesView] where cast([Date Record] as date) >='" &
            "01" & Format(Now, "-MMM-yyyy") & "' group by [Outlet Name], ItemName order by SumTotalNetSales desc"

        records = erpdb.Execute(SQLStatement)
        '   cast([Date Record] as date)

        counter01 = 0

        While Not records.EOF
            Try
                ItemNameTemp = ""
                ItemNameTemp = records.Fields("ItemName").Value

            Catch ex As Exception

            End Try

            Try
                OutletNameTemp = ""
                OutletNameTemp = records.Fields("Outlet Name").Value

            Catch ex As Exception

            End Try


            SumTotalNetSalesTemp = 0
            SumTotalNetSalesTemp = records.Fields("SumTotalNetSales").Value
            sumqtyTemp = 0
            sumqtyTemp = records.Fields("sumqty").Value

            'can now build the line
            DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<tr><td>" & "<p style =""text-align:left"">" & OutletNameTemp & "</td>" &
                                            "<td>" & "<p style =""text-align:left"">" & ItemNameTemp & "</td>" &
                                              "<td>" & "<p style =""text-align:right"">" & Format(SumTotalNetSalesTemp, "#,###,###") & "</td>" &
                                          "<td>" & "<p style =""text-align:right"">" & Format(sumqtyTemp, "#,###,###") & "</tr>"

            If counter01 = 0 Then
                '       SpeechString = SpeechString & " " & ItemNameTemp & " has sales of " & Format(SumTotalNetSalesTemp, "#,###") & "$ with " & Format(sumqtyTemp, "#,###,###") & " sold. "
            Else
                '       SpeechString = SpeechString & " " & ItemNameTemp & " sales " & Format(SumTotalNetSalesTemp, "#,###") & "$ , " & Format(sumqtyTemp, "#,###,###") & " sold. "
            End If
            counter01 += 1

            records.MoveNext()
        End While
        records.Close()


        'puts the bizcoach logo
        'this is desktop format so just a text string
        '"Yesterday: HK$ 1,014, last wk: HK$ 10,793"
        'http://aagilitycom.ipage.com/custview/20Jan1802glorystores.htm
        Dim CompanynameTemp As String = Companyname.Trim
        CompanynameTemp = Replace(CompanynameTemp, ".", "")
        CompanynameTemp = Replace(CompanynameTemp, " ", "")
        CompanynameTemp = Replace(CompanynameTemp, " ", "")
        CompanynameTemp = Replace(CompanynameTemp, "'", "")
        CompanynameTemp = CompanynameTemp.ToLower

        '      SMSText = SMSText & ControlChars.CrLf & "Total: $ " & Format(SalesYesterdayAll, "#,###,##0")

        UrlOfReport = "http://aagilitycom.ipage.com/custview/" & Format(DailyReportDate, "ddMMMyyyy") & "02" & CompanynameTemp & ".htm"
        DesktopText = "Yesterday: " & Format(InventoryValue, "HK$ #,###,##0") & ", last wk:" & Format(SalesYesterdayLastWeek, "HK$ #,###,##0")


        '      SpeechString = SpeechString & ". Total prediction for the next 7 days is " & Format(TotalPrediction7Days, "#,###") & "$ . "

        FrmMain.TextBoxSpeechText.Text = SpeechString.ToString.Trim

        'this creates the speech file

        AWSPollySpeak(SpeechString, "SalesByItemPOS")
        'this publishes the speech to the internet 
        PublishSpeechtoImpala(SpeechString, "SalesByItemPOS")
        'the file is bublished to    
        'http://aagilitycom.ipage.com/custview/    TextBoxPublishLocationName.text


        'closes the table
        DailyReportHTMLEmailString = "" & DailyReportHTMLEmailString & "</table></span><br /><hr /><br />"
        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<br />"

        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "If you would like BizCoach to read your email click <a href=""http://aagilitycom.ipage.com/custview/" &
            FrmMain.TextBoxPublishLocationName.Text.Trim & """>here</a><br /><br />"


        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "Please give feedback To <a href=""mailto:support@impalacloud.com"" target=""_top"">support@impalacloud.com</a><br /><br />"

        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<br />"

        'this creates the excel
        If FrmMain.CheckBoxCreateExcel.Checked = True Then
            CreateExcelSpreadsheet(DailyReportDate, ExcelHeaderLine0401,
                                      ExcelHeaderLine0402,
        ExcelHeaderLine0403,
        ExcelHeaderLine0404,
        ExcelHeaderLine0405,
        ExcelHeaderLine0406,
        ExcelHeaderLine0501,
        ExcelHeaderLine0502,
        ExcelHeaderLine0503,
        ExcelHeaderLine0504,
        ExcelHeaderLine0505,
        ExcelHeaderLine0506)
        End If

        'publish to the web browser on the notifications page
        Dim document As System.Windows.Forms.HtmlDocument = FrmMain.WebBrowserNotification.Document
        FrmMain.WebBrowserNotification.DocumentText = DailyReportHTMLEmailString

    End Sub


    Public Sub CreateExcelSpreadsheet(DailyReportDate, ExcelHeaderLine0401,
                                      ExcelHeaderLine0402,
        ExcelHeaderLine0403,
        ExcelHeaderLine0404,
        ExcelHeaderLine0405,
        ExcelHeaderLine0406,
        ExcelHeaderLine0501,
        ExcelHeaderLine0502,
        ExcelHeaderLine0503,
        ExcelHeaderLine0504,
        ExcelHeaderLine0505,
        ExcelHeaderLine0506)

        'this creates a copy locally
        Dim CustomerName As String = FrmMain.ListBoxCustomerInfo.SelectedItem
        CustomerName = Replace(CustomerName, " ", "")
        CustomerName = Replace(CustomerName, " ", "")
        CustomerName = CustomerName.ToLower
        Dim ExcelFileName As String = Format(Now, "ddMMMyy") + "ExcelXLSFile" + CustomerName

        Dim pathrootlocal As String = "c:\impalacustomers\"
        If Directory.Exists(pathrootlocal) = False Then
            Dim di As DirectoryInfo = Directory.CreateDirectory(pathrootlocal)
        End If
        pathrootlocal = "c:\impalacustomers\reportexcel\"
        If Directory.Exists(pathrootlocal) = False Then
            Dim di As DirectoryInfo = Directory.CreateDirectory(pathrootlocal)
        End If

        Dim OutputExcel As String = Replace(ExcelFileName, " ", "")

        'removes illegal characters
        Dim illegalChars As Char() = "~`|@#$^*{}'[]""_<>\/+.%?".ToCharArray()
        Try
            For Each ch As Char In illegalChars
                OutputExcel = OutputExcel.Replace(ch, "")
            Next
        Catch ex As Exception
            Dim strErrMsg = "Oops! Something is wrong with verify special characters at IsThereAnySpecialCharacter"
            MessageBox.Show(strErrMsg & vbCrLf & "Err: " & ex.Message)
        End Try

        'deletes any there already
        Try
            File.Delete("C:\ImpalaCustomers\reportexcel\" & OutputExcel + ".xls")
        Catch ex As Exception

        End Try

        'puts the name of the file into the main form which is then used later to craete the encrypted file
        FrmMain.TextBoxExcelLocation.Text = "C:\ImpalaCustomers\reportexcel\" & OutputExcel + ".xls"

        'This is the excel export
        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()

        If xlApp Is Nothing Then
            FrmMain.ListBoxNotificationStatus.Items.Add("Excel is not properly installed!!")
            FrmMain.ButtonDoAllNotifyTypes.BackColor = Color.Red
            FrmMain.ButtonDoAllNotifyTypes.Refresh()
            Exit Sub
        End If

        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim xlWorkSheet1 As Excel.Worksheet

        Dim misValue As Object = System.Reflection.Missing.Value

        xlWorkBook = xlApp.Workbooks.Add(misValue)

        Dim worksheets As Excel.Sheets = xlWorkBook.Worksheets
        Dim xlNewSheet = DirectCast(worksheets.Add(worksheets(1), Type.Missing, Type.Missing, Type.Missing), Excel.Worksheet)

        xlWorkSheet = xlWorkBook.Sheets("Sheet1")
        xlWorkSheet1 = xlWorkBook.Sheets("Sheet2")
        With xlWorkBook
            .Sheets("Sheet1").Select()
            .Sheets("Sheet1").Name = "Main"
            .Sheets("Sheet2").Name = "Data"
            '   .Sheets("Sheet3").Name = "Taroko #2"
            .Worksheets.Add(After:= .Worksheets(2)) 'INSERT AFTER WORKSHEET 3 ( "Taroko #2" )
            '    .Worksheets("Sheet4").Select()
            '     xlApp.Visible = True
        End With

        xlWorkBook.Sheets("Main").select

        'this writes the data
        'row, column

        xlWorkSheet.Cells(1, 1) = "Group Daily Sales for "
        xlWorkSheet.Cells(2, 1) = Format(DailyReportDate, "ddd d MMM yyyy")

        xlWorkSheet.Cells(4, 1) = ExcelHeaderLine0401
        xlWorkSheet.Cells(5, 1) = ExcelHeaderLine0501

        xlWorkSheet.Cells(4, 2) = ExcelHeaderLine0402
        xlWorkSheet.Cells(5, 2) = ExcelHeaderLine0502

        xlWorkSheet.Cells(4, 3) = ExcelHeaderLine0403
        xlWorkSheet.Cells(5, 3) = ExcelHeaderLine0503

        xlWorkSheet.Cells(4, 4) = ExcelHeaderLine0404
        xlWorkSheet.Cells(5, 4) = ExcelHeaderLine0504

        xlWorkSheet.Cells(4, 5) = ExcelHeaderLine0405
        xlWorkSheet.Cells(5, 5) = ExcelHeaderLine0505

        xlWorkSheet.Cells(4, 6) = ExcelHeaderLine0406
        xlWorkSheet.Cells(5, 6) = ExcelHeaderLine0506

        xlWorkSheet.Range("A1:X1").EntireColumn.AutoFit()

        'this adds the data
        '     xlWorkBook.Sheets("Data").select
        '    xlWorkBook.Sheets("Data").select
        '      xlWorkBook.Sheets(1).Select
        xlWorkBook.Sheets("Data").Activate

        xlWorkSheet.Cells(1, 1) = "bbs in Data"
        xlWorkSheet.Cells(2, 1) = "ooiiue"

        xlApp.DisplayAlerts = False
        xlWorkBook.SaveAs("c:\impalacustomers\reportexcel\" + ExcelFileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
         Excel.XlSaveAsAccessMode.xlExclusive)
        xlWorkBook.Close(True, misValue)
        xlApp.Quit()

        releaseObject(xlWorkSheet)
        releaseObject(xlWorkBook)
        releaseObject(xlApp)

        Dim HtmlString As String = FrmMain.ListBoxCustomerInfo.SelectedItem + Format(Now, "ddMMMyyyy") + "ExcelXLSFile"

        FrmMain.TextBoxToEncypt.Text = HtmlString
        EncryptText(HtmlString)

        Dim ExcelFileNameEncrypt = HtmlString & ".xls"

        FrmMain.TextBoxPublishLocationNameExcel.Text = ExcelFileNameEncrypt

        ' Create a web request that will be used to talk with the server and set the request method to upload a file by ftp.
        Dim ftpRequest As FtpWebRequest = CType(WebRequest.Create("ftp://ftp.impalacloud.com/" & ExcelFileNameEncrypt), FtpWebRequest)

        Try
            ftpRequest.Method = WebRequestMethods.Ftp.UploadFile

            ' Confirm the Network credentials based on the user name and password passed in.
            ftpRequest.Credentials = New NetworkCredential("BizLoading65xF", "JimboPass12!")
            'the file name is here

            Dim UploadSpeechFile = FrmMain.TextBoxExcelLocation.Text.Trim

            ' Read into a Byte array the contents of the file to be uploaded 
            Dim bytes() As Byte = System.IO.File.ReadAllBytes(UploadSpeechFile)

            ' Transfer the byte array contents into the request stream, write and then close when done.
            ftpRequest.ContentLength = bytes.Length
            Using UploadStream As Stream = ftpRequest.GetRequestStream()
                UploadStream.Write(bytes, 0, bytes.Length)
                UploadStream.Close()
            End Using
        Catch ex As Exception
            '   MessageBox.Show(ex.Message)
            FrmMain.ListBoxNotificationStatus.Items.Add("Publish excel to Impala error " & Err.Description)
            '   Exit Sub
        End Try

    End Sub

    Public Sub AWSPollySpeak(ByRef Speechtext, ByRef TypeOfSpeech)

        'limit is 1500 characters (actually it is 1500 characters excluding spaces so NEEDS WORK)
        If Len(Speechtext) > 1500 Then
            Speechtext = Mid(Speechtext, 1, 1430)
            Speechtext = Speechtext & ". Sorry there is too much for me to say here! I will cut it short!"
        End If

        Dim client As AmazonPollyClient = New AmazonPollyClient()
        Dim describeVoicesRequest As DescribeVoicesRequest = New DescribeVoicesRequest()
        Dim describeVoicesResult As DescribeVoicesResponse = client.DescribeVoices(describeVoicesRequest)
        Dim voices As List(Of Voice) = describeVoicesResult.Voices
        Dim synthesizeSpeechPresignRequest As SynthesizeSpeechRequest = New SynthesizeSpeechRequest()
        synthesizeSpeechPresignRequest.Text = Speechtext
        synthesizeSpeechPresignRequest.VoiceId = voices(0).Id
        synthesizeSpeechPresignRequest.OutputFormat = OutputFormat.Mp3
        Dim presignedSynthesizeSpeechUrl = client.SynthesizeSpeechAsync(synthesizeSpeechPresignRequest).GetAwaiter().GetResult()

        Dim OutputPolly As String = FrmMain.ListBoxCustomerInfo.SelectedItem & TypeOfSpeech & Format(Now, "ddMMMyyyy")
        '   OutputPolly.Name = "C:\ImpalaCustomers\Speech"

        OutputPolly = Replace(OutputPolly, " ", "")

        'removes illegal characters
        Dim illegalChars As Char() = "~`|@#$^*{}'[]""_<>\/+.%?".ToCharArray()
        Try
            For Each ch As Char In illegalChars
                OutputPolly = OutputPolly.Replace(ch, "")
            Next
        Catch ex As Exception
            Dim strErrMsg = "Oops! Something is wrong with verify special characters at IsThereAnySpecialCharacter"
            MessageBox.Show(strErrMsg & vbCrLf & "Err: " & ex.Message)
        End Try

        'deletes any there already
        Try
            File.Delete("C:\ImpalaCustomers\Speech\" & OutputPolly & ".mp3")
        Catch ex As Exception

        End Try

        'creates the directory
        Dim pathrootlocal As String = "C:\ImpalaCustomers\"
        If Directory.Exists(pathrootlocal) = False Then
            Dim di As DirectoryInfo = Directory.CreateDirectory(pathrootlocal)
        End If
        pathrootlocal = "C:\ImpalaCustomers\Speech\"
        If Directory.Exists(pathrootlocal) = False Then
            Dim di As DirectoryInfo = Directory.CreateDirectory(pathrootlocal)
        End If

        FrmMain.TextBoxSpeechLocation.Text = "C:\ImpalaCustomers\Speech\" & OutputPolly & ".mp3"

        Using output As FileStream = File.OpenWrite("C:\ImpalaCustomers\Speech\" & OutputPolly & ".mp3")
            presignedSynthesizeSpeechUrl.AudioStream.CopyTo(output)
        End Using
    End Sub


    Public Sub PublishSpeechtoImpala(ByVal SpeechText, ByRef TypeOfSpeech)

        'http://vbcity.com/blogs/xtab/archive/2016/04/13/how-to-upload-and-download-files-with-ftp-from-a-vb-net-application.aspx
        'how to access
        'http://www.impalacloud.com/custview/yyyutyt.htm

        'need to get the file names
        If FrmMain.ListBoxCustomerInfo.SelectedIndex = -1 Then
            MsgBox("select customer")
            Exit Sub
        End If
        Dim CustomerName As String = FrmMain.ListBoxCustomerInfo.SelectedItem
        CustomerName = Replace(CustomerName, " ", "")
        CustomerName = Replace(CustomerName, " ", "")
        CustomerName = CustomerName.ToLower
        Dim SpeechFileName As String = Format(Now, "ddMMMyy") + TypeOfSpeech + CustomerName & ".speech"

        Dim pathrootlocal As String = "c:\impalacustomers\"
        If Directory.Exists(pathrootlocal) = False Then
            Dim di As DirectoryInfo = Directory.CreateDirectory(pathrootlocal)
        End If
        pathrootlocal = "c:\impalacustomers\reportshtml\"
        If Directory.Exists(pathrootlocal) = False Then
            Dim di As DirectoryInfo = Directory.CreateDirectory(pathrootlocal)
        End If

        'writes a copy locally
        Dim path As String = "c:\impalacustomers\reportshtml\" & SpeechFileName

        ' This text is added only once to the file.
        If File.Exists(path) = False Then
            ' Create a file to write to.
            Dim createText As String = SpeechText + Environment.NewLine
            File.WriteAllText(path, createText)
        End If

        'BizLoading65xF     and   JimboPass12!

        Dim HtmlString As String = FrmMain.ListBoxCustomerInfo.SelectedItem + TypeOfSpeech + Format(Now, "ddMMMyyyy")

        FrmMain.TextBoxToEncypt.Text = HtmlString
        EncryptText(HtmlString)

        Dim SpeechFileNameEncrypt = HtmlString & ".mp3"

        FrmMain.TextBoxPublishLocationName.Text = SpeechFileNameEncrypt

        ' Create a web request that will be used to talk with the server and set the request method to upload a file by ftp.
        Dim ftpRequest As FtpWebRequest = CType(WebRequest.Create("ftp://ftp.impalacloud.com/" & SpeechFileNameEncrypt), FtpWebRequest)
        Try
            ftpRequest.Method = WebRequestMethods.Ftp.UploadFile
            ' Confirm the Network credentials based on the user name and password passed in.
            ftpRequest.Credentials = New NetworkCredential("BizLoading65xF", "JimboPass12!")
            'the file name of the mp3 file is here
            Dim UploadSpeechFile = FrmMain.TextBoxSpeechLocation.Text.Trim
            ' Read into a Byte array the contents of the file to be uploaded 
            Dim bytes() As Byte = System.IO.File.ReadAllBytes(UploadSpeechFile)
            ' Transfer the byte array contents into the request stream, write and then close when done.
            ftpRequest.ContentLength = bytes.Length
            Using UploadStream As Stream = ftpRequest.GetRequestStream()
                UploadStream.Write(bytes, 0, bytes.Length)
                UploadStream.Close()
            End Using
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            FrmMain.ListBoxNotificationStatus.Items.Add("Publish speech to Impala error " & Err.Description)
            '   Exit Sub
        End Try
    End Sub


    Public Sub EncryptText(ByRef TextToEncrypt)

        'https://docs.microsoft.com/en-us/dotnet/visual-basic/programming-guide/language-features/strings/walkthrough-encrypting-and-decrypting-strings

        Dim plainText As String = TextToEncrypt
        Dim password As String = FrmMain.TextBoxCypher.Text.Trim

        Dim wrapper As New Simple3Des(password)
        Dim cipherText As String = wrapper.EncryptData(plainText)

        '     MsgBox("The cipher text is: " & cipherText)

        FrmMain.TextBoxEncypted.Text = cipherText

        My.Computer.FileSystem.WriteAllText(
            My.Computer.FileSystem.SpecialDirectories.MyDocuments &
            "\cipherText.txt", cipherText, False)

        'also removes the forbidden characters so that we publish the page
        Dim cipherTextstripped = Replace(cipherText, "/", "")
        cipherTextstripped = Replace(cipherTextstripped, "\", "")
        cipherTextstripped = Replace(cipherTextstripped, "+", "")
        cipherTextstripped = Replace(cipherTextstripped, " ", "")

        'removes illegal characters
        Dim illegalChars As Char() = "~`|@#$^*{}'[]""_<>\/+.%?".ToCharArray()
        Try
            For Each ch As Char In illegalChars
                cipherTextstripped = cipherTextstripped.Replace(ch, "")
            Next
        Catch ex As Exception
            Dim strErrMsg = "Oops! Something is wrong with verify special characters at IsThereAnySpecialCharacter"
            MessageBox.Show(strErrMsg & vbCrLf & "Err: " & ex.Message)
        End Try

        FrmMain.TextBoxStrippedOfIllegalCharacters.Text = cipherTextstripped
        TextToEncrypt = cipherTextstripped
    End Sub



    Public Sub DailySalesReportTest(ByRef DailyReportHTMLEmailString As String, ByRef DailyReportDate As Date)

        'gets the middel of the report
        Dim erpdb As ADODB.Connection
        Dim records As ADODB._Recordset
        Dim recordsLookup As ADODB._Recordset
        erpdb = New ADODB.Connection
        erpdb.Open(strconnectSQLBizCoachCustomerDataSQL)

        Dim SalesYesterday As Integer = 0
        Dim SalesMTDYesterday As Integer = 0
        Dim OutletName As String = ""

        Dim SalesYesterdayLastWeek As Integer = 0

        Dim SalesLastMonthVariance As Decimal = 0
        Dim SalesYesterdayLastWeekVariance As Decimal = 0
        Dim SalesLastMonthToDate As Integer = 0

        'uses UTC time so need to take away 8 hours or use HongKongLoginMoment

        records = erpdb.Execute("Select * from [KPI Values] where [KPI Name] ='POS Daily Sales (today) - Total'")
        SalesYesterday = 0
        While Not records.EOF
            SalesYesterday = 0
            Try
                SalesYesterday = records.Fields("KPI Value Number").Value
            Catch ex As Exception
                '        usernameTemp = ""
            End Try
            records.MoveNext()
        End While
        records.Close()


        records = erpdb.Execute("Select * from [KPI Values] where [KPI Name] ='POS Sales MTD - Total'")
        While Not records.EOF
            Try
                SalesMTDYesterday = records.Fields("KPI Value Number").Value
            Catch ex As Exception
                '        usernameTemp = ""
            End Try
            records.MoveNext()
        End While
        records.Close()

        records = erpdb.Execute("Select * from [KPI Values] where [KPI Name] ='POS Daily Sales (yesterday same day last week) - Total'")
        While Not records.EOF
            Try
                SalesYesterdayLastWeek = records.Fields("KPI Value Number").Value
            Catch ex As Exception
                '        usernameTemp = ""
            End Try
            records.MoveNext()
        End While
        records.Close()


        records = erpdb.Execute("Select * from [KPI Values] where [KPI Name] ='POS Daily Sales Yesterday to LW Same Day Variance (%) - Total'")
        While Not records.EOF
            Try
                SalesYesterdayLastWeekVariance = records.Fields("KPI Value Number").Value
            Catch ex As Exception
                '        usernameTemp = ""
            End Try
            records.MoveNext()
        End While
        records.Close()

        records = erpdb.Execute("Select * from [KPI Values] where [KPI Name] ='POS Sales MTD Yesterday Variance LMTD Sales % - Total'")

        While Not records.EOF
            Try
                SalesLastMonthVariance = records.Fields("KPI Value Number").Value
            Catch ex As Exception
                '        usernameTemp = ""
            End Try
            records.MoveNext()
        End While
        records.Close()

        records = erpdb.Execute("Select * from [KPI Values] where [KPI Name] ='POS Sales LMTD - Total'")

        While Not records.EOF
            Try
                SalesLastMonthToDate = records.Fields("KPI Value Number").Value
            Catch ex As Exception
                '        usernameTemp = ""
            End Try
            records.MoveNext()
        End While
        records.Close()



        'Yesterday last week


        '      DailyReportHTMLEmailString = "      <div style = ""width:400px;"">
        '         <h1 style=""text-align:center""> <background-color: blue;>Testing </h1>
        '        <p  style =""text-align:right""><a href=""#"">sample link</a> </p>
        '       </div>"


        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "Group Daily Sales for : " & Format(DailyReportDate, "ddd dd MMM yyyy") & "<br /><br />" &
            "Yesterday: " & Format(SalesYesterday, "HK$ #,###,###") & ", same day last week : " & Format(SalesYesterdayLastWeek, "HK$ #,###,###") & ", variance to last week " & Format(SalesYesterdayLastWeekVariance, "#.##") & "%" & "<br />" &
            "MTD      : " & Format(SalesMTDYesterday, "HK$ #,###,###") & ", last month to date " & "" & Format(SalesLastMonthToDate, "HK$ #,###,###") & ", variance to last month " & Format(SalesLastMonthVariance, "#.##") & "%" & "<br />" &
                                                     "<br />" & "<br />" &
                                                     "<table border=1><col width=""300""><col width=""150""><col width=""150""><tr>" &
                                                     "<td><b>Outlet Sales</td>" &
                                                     "<td><b><p  style =""text-align:center"">Yesterday</td>" &
                                                     "<td><b><p  style =""text-align:center"">MTD</b></td>" &
                                                     "</b></tr>"


        'now gets the individual outlets
        'needs to sum as may be multiple lines

        records = erpdb.Execute("Select sum([KPI Value Number]) as SumSales, [Outlet Name] from [KPI Values] where [KPI Name] ='POS Daily Sales (today)' group by [Outlet Name] order by [Outlet Name]")
        SalesYesterday = 0
        While Not records.EOF
            SalesYesterday = 0
            Try
                OutletName = records.Fields("Outlet Name").Value
                SalesYesterday = records.Fields("SumSales").Value
            Catch ex As Exception
                MsgBox("error " & Err.Description)
            End Try


            'now needs to get the MTD for the outlet
            recordsLookup = erpdb.Execute("Select * from [KPI Values] where [KPI Name] ='POS Sales MTD by Dimensions' and [Outlet Name] ='" & Replace(OutletName, "'", "''") & "'")
            SalesMTDYesterday = 0
            While Not recordsLookup.EOF
                SalesMTDYesterday = recordsLookup.Fields("KPI Value Number").Value
                recordsLookup.MoveNext()
            End While
            recordsLookup.Close()

            'can build the line

            DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<tr><td>" & OutletName & "</td>" &
                                         "<td><p  style =""text-align:right"">" & Format(SalesYesterday, "HK$ #,###,###") & "</td>" &
                                        "<td><p  style =""text-align:right"">" & Format(SalesMTDYesterday, "HK$ #,###,###") & "</tr>"

            records.MoveNext()
        End While
        records.Close()



        'closes the table
        DailyReportHTMLEmailString = "" & DailyReportHTMLEmailString & "</table></span><br /><hr /><br />"
        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<br />"
        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "    <img src = ""http://aagilitycom.ipage.com/bizcoachimages/BizCoach_small.png"" alt=""BizCoach"" ;  >"

        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<br />    <img src = ""http://aagilitycom.ipage.com/bizcoachimages/reddotsculpted.jpg"" alt=""BizCoach"" ;  >"
        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<br />    <img src = ""http://aagilitycom.ipage.com/bizcoachimages/orangedotsculpted.jpg"" alt=""BizCoach"" ;  >"
        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<br />    <img src = ""http://aagilitycom.ipage.com/bizcoachimages/greendotsculpted.jpg"" alt=""BizCoach"" ;  >"
        DailyReportHTMLEmailString = DailyReportHTMLEmailString & "<br />    <img src = ""http://aagilitycom.ipage.com/bizcoachimages/greendotplain.png"" alt=""BizCoach"" ;  >"


        'Old one         DailyReportHTMLEmailString = DailyReportHTMLEmailString & "    <img src = ""http://aagilitycom.ipage.com/bizcoachimages/greendotplain.png"" alt=""BizCoach"" ;  >"


        '    Dim document As System.Windows.Forms.HtmlDocument =
        '   WebBrowserNotifications.Document


        ' WebBrowserNotifications.DocumentText = DailyReportHTMLEmailString


        '    "<html><body>Please enter your name:<br/>" &
        '    "<input type='text' name='userName'/><br/>" &
        '    "<a href='http://www.microsoft.com'>continue</a>" &
        '    "</body></html>"
        'If document IsNot Nothing And
        'document.All("userName") IsNot Nothing And
        'String.IsNullOrEmpty(
        'document.All("userName").GetAttribute("value")) Then

        '        e.Cancel = True
        '        MsgBox("You must enter your name before you can navigate to " &
        '        e.Url.ToString())
        '       End If



        erpdb.Close()

    End Sub



End Module
