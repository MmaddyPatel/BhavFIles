Imports Excels = Microsoft.Office.Interop.Excel
Imports System.Linq
Imports System.Security
Imports System.Net.Mail
Imports Npgsql
Imports System.Net
Imports System.IO
Imports Ionic.Zip
Imports ZEE.DAL
Imports Newtonsoft.Json
Imports System.Reflection
Imports Microsoft.Office.Interop
Imports HtmlAgilityPack
Imports System.Text

Module Module1

#Region "DECLERATIONS"
    Dim Dt_Fno_Mapping As DataTable
    Dim MyBook As Excels.Workbook
    Dim MyApp As Excels.Application
    Dim MySheet As Excels.Worksheet
    Dim NIFTY_TOTALPUT As Long = 0
    Dim NIFTY_TOTALCALL As Long = 0
    Dim Dates As Date = Now.Date
    'Dim Dates As Date = Now.Date.AddDays(-1)
    Dim indices As String
    Dim STOCK_ITEMS_updated As Integer
    Dim localDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) & "\bhavfiles\"
    Dim localfile As String = Environment.CurrentDirectory & "\bhavfiles\" & "\bhavfiles\" & Format(Dates, "ddMMyy") & ".csv"
    Dim url2download As String = "https://www1.nseindia.com/content/equities/sec_list.csv"
    Dim startTime As DateTime
    Dim clsWrite As New WriteToLogs
    Dim NSEcounterInsert As Integer = 1
    Dim NSEcounterUpdate As Integer = 1
    Dim BSEcounterUpdate As Integer = 1
    Dim BSEcounterInsert As Integer = 1
    Dim mailContent As String = "<tr bgcolor=#FF0000><td>VARIANT</td><td>START TIME</td><td>END TIME</td></tr>"
    Dim strddMMyyyy As String = Format(Dates, "ddMMyyyy")
    Dim strddMMyy As String = Format(Dates, "ddMMyy")
    Dim strddMMMyyyy As String = Format(Dates, "ddMMMyyyy")
    Dim stryyyy As String = Format(Dates, "yyyy")
    Dim strMMM As String = Format(Dates, "MMM")
    Dim strdd As String = Format(Dates, "dd")
    Public Class WebClientWithTimeout
        Inherits WebClient

        Protected Overrides Function GetWebRequest(ByVal address As Uri) As WebRequest
            Dim wr As WebRequest = MyBase.GetWebRequest(address)
            wr.Timeout = 5000
            Return wr
        End Function
    End Class

    Public Class RootObject

        Public Property data As Dictionary(Of String, Data)
    End Class

    Public Class Data

        Public Property EXPECTED_REPORT_PERIOD_END_DATE As String
        Public Property VOLUME_AVG_30D As String
        Public Property EXPECTED_REPORT_DT As String
        Public Property EQY_TRR_PCT_1YR As String
        Public Property PX_TO_SALES_RATIO As String
        Public Property TRAIL_12M_EPS As String
        Public Property EQY_SH_OUT As String
        Public Property LATEST_ANNOUNCEMENT_PERIOD As String
        Public Property EQY_DVD_YLD_IND As String
        Public Property PE_RATIO As String
        Public Property PX_TO_BOOK_RATIO As String
        Public Property DVD_SH_LAST As String
        Public Property BEST_EEPS_CUR_YR As String
        Public Property BEST_PEG_RATIO As String
        Public Property BEST_PE_RATIO As String
    End Class

#End Region

#Region "FORM EVENTS"


    Sub Main()


        DeleteOldFiles(Directory.GetCurrentDirectory & "\bhavfiles\", 5)


        If (My.Settings.day <> 0) Then
            Dates = Now.Date.AddDays("-" & My.Settings.day)


        End If





        Dim myCulture As System.Globalization.CultureInfo = Globalization.CultureInfo.CurrentCulture
        Dim dayOfWeek As DayOfWeek = myCulture.Calendar.GetDayOfWeek(Dates)
        ' dayOfWeek.ToString() would return "Sunday" but it's an enum value,
        ' the correct dayname can be retrieved via DateTimeFormat.
        ' Following returns "Sonntag" for me since i'm in germany '
        Dim dayName As String = myCulture.DateTimeFormat.GetDayName(dayOfWeek)
        Console.WriteLine(dayName.ToString())


        If (My.Settings.day <> 0) Then
            Console.WriteLine("Are you sure you want to download data for " & dayName & "? [yes/no]")
            Dim stroption As String = Console.ReadLine()

            If (stroption.ToLower <> "yes") Then
                End
            End If
        End If
        strddMMyyyy = Format(Dates, "ddMMyyyy")
        strddMMyy = Format(Dates, "ddMMyy")
        strddMMyyyy = Format(Dates, "ddMMMyyyy")
        stryyyy = Format(Dates, "yyyy")
        strMMM = Format(Dates, "MMM")
        strdd = Format(Dates, "dd")



        'Return
        'readPESyndication()

        'Return

        'Dim dt_excel = New DataTable
        'dt_excel = ExcelToDataTable("G:\BhavFiles\121017\FNO\fo12OCT2017bhav.csv")
        'Dim dtFilter As String = "instrument = 'OPTIDX' and SYMBOL ='NIFTY'"

        'Dim dt_sql = New DataTable
        'Dim Sql As String = "select 'OPTIDX' as INSTRUMENT,'NIFTY' as SYMBOL,expiry_date as EXPIRY_DT,STRIKE_PRICE  from OPTIONS where exchange_symbol ='NIFTY 50' order by expiry_date,strike_price"
        'dt_sql = SqlHelper.ExecuteDataset(My.Settings.conn_str, CommandType.Text, Sql).Tables(0)


        'Dim dtFilter_sql As String = "instrument = 'OPTIDX' and SYMBOL ='NIFTY'"
        'For I As Integer = 0 To dt_sql.rows.count  - 1

        'Next'


        '
        Try




            If Directory.Exists(localDirectory & strddMMyy & "\BSE") = False Then
                Directory.CreateDirectory(localDirectory & strddMMyy & "\BSE")
            End If

            If Directory.Exists(localDirectory & strddMMyy & "\NSE") = False Then
                Directory.CreateDirectory(localDirectory & strddMMyy & "\NSE")
            End If

            If Directory.Exists(localDirectory & strddMMyy & "\NSE_ISIN_FV") = False Then
                Directory.CreateDirectory(localDirectory & strddMMyy & "\NSE_ISIN_FV")
            End If


            If Directory.Exists(localDirectory & strddMMyy & "\FNO") = False Then
                Directory.CreateDirectory(localDirectory & strddMMyy & "\FNO")
            End If

            If Directory.Exists(localDirectory & strddMMyy & "\CIRCUIT_NSE") = False Then
                Directory.CreateDirectory(localDirectory & strddMMyy & "\CIRCUIT_NSE")
            End If

            If Directory.Exists(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\BSEDELSTATS") = False Then
                Directory.CreateDirectory(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\BSEDELSTATS")
            End If

            If Directory.Exists(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\NSEDELSTATS") = False Then
                Directory.CreateDirectory(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\NSEDELSTATS")
            End If

            If Directory.Exists(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\Insert_BandChange") = False Then
                Directory.CreateDirectory(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\Insert_BandChange")
            End If

            If Directory.Exists(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\NSE_HISTORICAL") = False Then
                Directory.CreateDirectory(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\NSE_HISTORICAL")
            End If

            If Directory.Exists(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\FO_FII_DERIVATIVES_STATS") = False Then
                Directory.CreateDirectory(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\FO_FII_DERIVATIVES_STATS")
            End If

            If Directory.Exists(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "fo_participant_oi") = False Then
                Directory.CreateDirectory(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\fo_participant_oi")
            End If


            If Directory.Exists(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "BSE") = False Then
                Directory.CreateDirectory(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\BSE")
            End If

        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try

        ' Call Download_Circuit_NSE()
        ''call FNO()
        'ParseEOD()
        'Call download_fo_participant_oi()
        'Call downloadFIIFPIDIIActivity()
        'Call downloadFO_FII_DERIVATIVES_STATS()
        'Call BSE_ISIN()
        'Return
        Try
            Dim strsql As String = "exec SP_HOLIDAYCHECK"
            SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strsql)



            mailContent += "<tr><th>DAY</th><td>" & dayName & "</td>"
            Try

                mailContent += "<tr><td>NSE " & "https://archives.nseindia.com/archives/equities/bhavcopy/pr/PR" & strddMMyy & ".zip" & "</td><td>" & DateTime.Now & "</td>"
                Call NSE()
                mailContent += "<td>" & DateTime.Now & "</td></tr>"

            Catch ex As Exception
                mailContent += "<td bgcolor=#FF0000>" & ex.Message & "</td><" & DateTime.Now & "></td>"
                Console.WriteLine(Now.ToString() + " Error In  NSE " & ex.Message)
            End Try

            Try
                mailContent += "<tr><td>BSE " & "http://www.bseindia.com/download/BhavCopy/Equity/" & "EQ" & strddMMyy & "_CSV.ZIP" & "</td><td>" & DateTime.Now & "</td>"
                Call BSE()
                mailContent += "<td >" & DateTime.Now & "</td></tr>"
            Catch ex As Exception
                mailContent += "<td bgcolor=#FF0000>" & ex.Message & "</td><" & DateTime.Now & "></td>"
                Console.WriteLine(Now.ToString() + " Error In  BSE " & ex.Message)
            End Try

            Try

                mailContent += "<tr><td>FNO " & "https://archives.nseindia.com/content/historical/DERIVATIVES/" & stryyyy & "/" & strMMM.ToUpper() & "/fo" & strddMMyyyy.ToUpper() & "bhav.csv.zip" & "</td><td><" & DateTime.Now & "></td>"
                Call FNO()
                mailContent += "<td>" & DateTime.Now & "</td></tr>"


                Try
                    mailContent += "<tr><td>FNO EOD INSERTED </td><td><" & DateTime.Now & "></td>"
                    SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, "SP_INSERT_STOCKS_OPT_EOD " + (My.Settings.day).ToString())
                    'SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, "delete from  option_items where expiry_date < convert(date,getdate()-" + (My.Settings.day).ToString() + ",101)");
                Catch ex As Exception

                    mailContent += "<td bgcolor=#FF0000>" & ex.Message & "</td><" & DateTime.Now & "></td>"
                    '  SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, "SP_INSERT_STOCKS_OPT_EOD")
                End Try


            Catch ex As Exception
                mailContent += "<td>" & ex.Message & "</td></tr>"
                Console.WriteLine(Now.ToString() + " Error In  FNO " & ex.Message)
            End Try

            Try

                mailContent += "<tr><td>BSE ISIN " & "http://www.bseindia.com/download/bhavcopy/equity/" & "EQ_ISINCODE_" & strddMMyy & ".zip" & "</td><td>" & DateTime.Now & "</td>"
                Call BSE_ISIN()
                mailContent += "<td>" & DateTime.Now & "</td></tr>"
            Catch ex As Exception
                mailContent += "<td bgcolor=#FF0000>" & ex.Message & "</td><" & DateTime.Now & "></td>"
                Console.WriteLine(Now.ToString() + " Error In  BSE " & ex.Message)
            End Try

            Try
                mailContent += "<tr><td>NSE ISIN_FV " & "https://www1.nseindia.com/content/equities/EQUITY_L.csv" & "</td><td>" & DateTime.Now & "</td>"
                Call NSE_ISIN_FV()
                mailContent += "<td>" & DateTime.Now & "</td></tr>"

            Catch ex As Exception
                mailContent += "<td bgcolor=#FF0000>" & ex.Message & "</td><" & DateTime.Now & "></td>"
                Console.WriteLine(Now.ToString() + " Error In  BSE " & ex.Message)
            End Try

            Try

                mailContent += "<tr><td>CIRCUIT " & "https://archives.nseindia.com/content/equities/sec_list.csv" & "</td><td><" & DateTime.Now & "></td>"
                Call Download_Circuit_NSE()
                mailContent += "<td>" & DateTime.Now & "</td></tr>"
            Catch ex As Exception
                mailContent += "<td bgcolor=#FF0000>" & ex.Message & "</td><" & DateTime.Now & "></td>"
                Console.WriteLine(Now.ToString() + " Error In  CKT " & ex.Message)
            End Try

            'Try
            '    mailContent += "<tr><td>NSE</td><td>" & DateTime.Now & "</td>"
            '    Call DELIVERY()
            '    mailContent += "<td>" & DateTime.Now & "</td></tr>"
            'Catch ex As Exception
            '    mailContent += "<td>" & ex.Message & "</td></tr>"
            '    Console.WriteLine(Now.ToString() + " Error In  NSE " & ex.Message)
            'End Try


            Try
                mailContent += "<tr><td>MARGIN</td><td>" & DateTime.Now & "</td>"
                Call DownloadVarMarginFiles()
                mailContent += "<td>" & DateTime.Now & "</td></tr>"
            Catch ex As Exception
                mailContent += "<td bgcolor=#FF0000>" & ex.Message & "</td><" & DateTime.Now & "></td>"
                Console.WriteLine(Now.ToString() + " Error In  MARGIN " & ex.Message)
            End Try

            Try
                mailContent += "<tr><td>BSEDELIVERY" & url2download_BSEDelivery & "</td><td>" & DateTime.Now & "</td>"
                BSEDELIVERY()
                mailContent += "<td>" & DateTime.Now & "</td></tr>"
            Catch ex As Exception
                mailContent += "<td bgcolor=#FF0000>" & ex.Message & "</td><" & DateTime.Now & "></td>"
                Console.WriteLine(Now.ToString() + " Error In  BSEDELIVERY " & ex.Message)
            End Try

            Try
                mailContent += "<tr><td>NSEDELIVERY " & url2download_NSEDELIVERY & "</td><td>" & DateTime.Now & "</td>"
                Call NSEDELIVERY()
                mailContent += "<td>" & DateTime.Now & "</td></tr>"
            Catch ex As Exception
                mailContent += "<td bgcolor=#FF0000>" & ex.Message & "</td><" & DateTime.Now & "></td>"
                Console.WriteLine(Now.ToString() + " Error In  NSEDELIVERY " & ex.Message)
            End Try
            Try
                mailContent += "<tr><td>NSE_WKHIGH_LOW " & "https://archives.nseindia.com/content/CM_52_wk_High_low_" & strddMMyyyy & ".csv" & "</td><td>" & DateTime.Now & "</td>"
                Call NSE_WK52HI_LOW()
                mailContent += "<td>" & DateTime.Now & "</td></tr>"
            Catch ex As Exception
                mailContent += "<td>" & ex.Message & "</td></tr>"
                Console.WriteLine(Now.ToString() + " Error In  NNSE_WKHIGH_LOW " & ex.Message)
            End Try

            Try
                mailContent += "<tr><td>FII FPI DII " & "https://www.bseindia.com/markets/equity/EQReports/categorywise_turnover.aspx" & "</td><td>" & DateTime.Now & "</td>"
                Call downloadFIIFPIDIIActivity()
                mailContent += "<td>" & DateTime.Now & "</td></tr>"
            Catch ex As Exception
                mailContent += "<td bgcolor=#FF0000>" & ex.Message & "</td><" & DateTime.Now & "></td>"
                Console.WriteLine(Now.ToString() + " Error In  downloadFIIFPIDIIActivity() " & ex.Message)
            End Try

            Try
                mailContent += "<tr><td>downloadFO_FII_DERIVATIVES_STATS " & "https://archives.nseindia.com/content/fo/fii_stats_" & strdd & "-" & strMMM & "-" & stryyyy & ".xls</td><td>" & DateTime.Now & "</td>"
                Call downloadFO_FII_DERIVATIVES_STATS()
                mailContent += "<td>" & DateTime.Now & "</td></tr>"
            Catch ex As Exception
                mailContent += "<td bgcolor=#FF0000>" & ex.Message & "</td><" & DateTime.Now & "></td>"
                Console.WriteLine(Now.ToString() + " Error In  downloadFO_FII_DERIVATIVES_STATS " & ex.Message)
            End Try

            Try
                mailContent += "<tr><td>download_fo_participant_oi " & "https://archives.nseindia.com/content/nsccl/fao_participant_oi_" & Format(Dates, "ddMMyyyy") & ".csv" & "</td><td>" & DateTime.Now & "</td>"
                Call download_fo_participant_oi()
                mailContent += "<td>" & DateTime.Now & "</td></tr>"
            Catch ex As Exception
                mailContent += "<td bgcolor=#FF0000>" & ex.Message & "</td><" & DateTime.Now & "></td>"
                Console.WriteLine(Now.ToString() + " Error In  download_fo_participant_oi " & ex.Message)
            End Try



            Try
                mailContent += "</table>"
                Mail("developers@zeebiz.com", My.Settings.Mailto, "Bhav Files for " & dayName & " - " & Dates, "BHAV FILES FOR " & dayName & " - " & Dates, mailContent)
            Catch ex As Exception

            End Try


        Catch ex As Exception
            Mail("developers@zeebiz.com", My.Settings.Mailto, "Error in Bhav Files for " & dayName & " - " & Dates & "==> " & ex.Message, "Error in BHAV FILES FOR " & dayName & " - " & Dates & "==>" & ex.Message, mailContent)
            Console.WriteLine(ex.Message)
        End Try


        End


    End Sub
#End Region


    Private Sub downloadFIIFPIDIIActivity()
        Try
            Dim lstData As List(Of String) = New List(Of String)()
            Dim dataDic As Dictionary(Of String, String) = New Dictionary(Of String, String)()
            ServicePointManager.Expect100Continue = True
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
            Dim Url As String = "https://www.bseindia.com/markets/equity/EQReports/categorywise_turnover.aspx"

            Dim inStream As StreamReader
            Dim webRequest As WebRequest
            Dim webresponse As WebResponse
            webRequest = WebRequest.Create(Url)
            webresponse = webRequest.GetResponse()
            inStream = New StreamReader(webresponse.GetResponseStream())

            Dim filename As String = localDirectory & strddMMyy & "\FO_FII_DERIVATIVES_STATS\fii.html"

            Dim file As System.IO.StreamWriter
            file = My.Computer.FileSystem.OpenTextFileWriter(filename, False)
            file.WriteLine(inStream.ReadToEnd())

            file.Close()
            file = Nothing


            Dim htmlDoc As HtmlAgilityPack.HtmlDocument = New HtmlAgilityPack.HtmlDocument()
            htmlDoc.OptionFixNestedTags = True
            htmlDoc.Load(filename)
            lstData.Clear()

            For Each table As HtmlNode In htmlDoc.DocumentNode.SelectNodes("//table[@id='ContentPlaceHolder1_offTblBdyFII']")

                For Each row As HtmlNode In table.SelectNodes(".//tr")

                    For Each cell As HtmlNode In row.SelectNodes("th|td")
                        lstData.Add(cell.InnerText.ToString().Trim().Replace("&nbsp;", " "))
                    Next
                Next
            Next

            InsertIntoDB(lstData, "FO_FII_DERIVATIVES_STATS_" & strddMMyy & ".html")
            lstData.Clear()

            For Each table As HtmlNode In htmlDoc.DocumentNode.SelectNodes("//table[@id='ContentPlaceHolder1_offTblBdyDII']")

                For Each row As HtmlNode In table.SelectNodes(".//tr")

                    For Each cell As HtmlNode In row.SelectNodes("th|td")
                        lstData.Add(cell.InnerText.ToString().Trim().Replace("&nbsp;", " "))
                    Next
                Next
            Next

            InsertIntoDB(lstData, "FO_FII_DERIVATIVES_STATS_" & strddMMyy & ".html")
            lstData.Clear()
            'End Using


        Catch ex As Exception
            clsWrite.CaptureLogs(localDirectory & strddMMyy & "\FO_FII_DERIVATIVES_STATS\", ex.Message & " ===> ", "err")
        End Try
    End Sub

    Public Sub InsertIntoDB(ByVal lstFii As List(Of String), filename As String)
        Dim contractName As String = "", buyContracts As Double = 0, sellContracts As Double = 0, buyAmount As Double = 0, sellAmount As Double = 0, oiContracts As Integer = 0, xAmount As Double
        Dim fileDate As DateTime
        Dim fileDateStr As String
        Dim strSql As String
        If lstFii.Count > 0 Then

            For loopVar As Integer = 0 To lstFii.Count

                Try
                    contractName = lstFii(loopVar + 5).ToString()
                    fileDate = DateTime.ParseExact((lstFii(loopVar + 6).ToString()), "d/M/yyyy", Nothing, System.Globalization.DateTimeStyles.None)
                    'fileDate = DateTime.ParseExact((lstFii(loopVar + 6).ToString()), "yyyy-M-d", System.Globalization.CultureInfo.InvariantCulture)
                    'fileDate = Convert.ToDateTime((lstFii(loopVar + 6).ToString()))
                    'fileDateStr = fileDate.ToString("yyyy-MM-dd")

                    Dim dtN As DateTime = DateTime.Now()

                    'Dim cultInfo As System.Globalization.CultureInfo = New System.Globalization.CultureInfo("en-GB", True)
                    'Dim formatInfo As System.Globalization.DateTimeFormatInfo = cultInfo.DateTimeFormat
                    'formatInfo.ShortDatePattern = "dd/MM/yyyy"
                    'formatInfo.LongDatePattern = "dd MMMM yyyy"
                    'formatInfo.FullDateTimePattern = "dd MMMM yyyy HH:mm:ss"
                    'fileDate = DateTime.Parse(lstFii(loopVar + 6).ToString(), formatInfo)


                    buyAmount = lstFii(loopVar + 7).ToString()
                    sellAmount = lstFii(loopVar + 8).ToString()
                    xAmount = lstFii(loopVar + 9).ToString()

                    If loopVar < 12 Then

                        strSql = " Select * from FO_FII_DERIVATIVES_STATS where contractname ='" & contractName & "' and FILEDATE ='" & fileDate & "'"
                        Dim dt As DataTable = New DataTable
                        dt = SqlHelper.ExecuteDataset(My.Settings.conn_str, CommandType.Text, strSql).Tables(0)

                        If dt.Rows.Count > 0 Then
                            strSql = "update FO_FII_DERIVATIVES_STATS set buycontracts =" & buyContracts & ", sellcontracts=" & sellContracts & "" &
                                    ", buyamount =" & buyAmount & ", sellamount=" & sellAmount & ",oicontracts =" & oiContracts & ", xamount='" & xAmount & "'" &
                                    ", updatedatetime = getdate(), filename = '" & filename & "'" &
                                    " where contractname ='" & contractName & "' and filedate='" & fileDate & "'"
                            SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)





                        Else
                            strSql = "insert into FO_FII_DERIVATIVES_STATS "
                            strSql += "(contractName,buyContracts,buyAmount,sellContracts,sellAmount,oiContracts,fileDate,xAmount,UpdateDateTime,filename) values("
                            strSql += "'" & contractName & "'," & buyContracts & "," & buyAmount & "," & sellContracts & "," & sellAmount & "," & oiContracts & ",'" & fileDate & "','" & xAmount & "',getdate(),'" & filename & "')"
                            Console.WriteLine("Importing Record to FPI/DII Derivative :: " & strSql)

                            SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)
                            '' SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, Sql)
                        End If
                    End If
                Catch ex As Exception
                    clsWrite.CaptureLogs(localDirectory & strddMMyy & "\FO_FII_DERIVATIVES_STATS\", ex.Message & " ===> " & strSql, "err")
                End Try

                loopVar = loopVar + 10
            Next
        End If
    End Sub



#Region "download_fo_participant_oi" '' https://archives.nseindia.com/content/nsccl/fao_participant_oi_12052023.csv // https://www.nseindia.com/all-reports-derivatives
    Sub download_fo_participant_oi()

        Dim file2download As String = ""
        Dim strSql As String

        Try

            Dim DirectoryName = Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\fo_participant_oi"
            If Directory.Exists(DirectoryName) = False Then
                Directory.CreateDirectory(DirectoryName)
            End If

            ''  file2download = "CM_52_wk_High_low_" & strddMMyyyy & ".csv"
            Dim localfile As String = DirectoryName & "\fo_participant_oi" & strddMMyy & ".csv"

            Dim url2download As String = "https://archives.nseindia.com/content/nsccl/fao_participant_oi_" & Format(Dates, "ddMMyyyy") & ".csv"


            If (DownloadFile(url2download, localfile) = False) Then
                GoTo nextloop
            End If

            DownloadFile(url2download, localfile)




            Dim fileName As String = localfile ' Environment.CurrentDirectory & "\bhavfiles\" & localfile & "\NSEDELSTATS\" & file2download
            Dim fs As New FileStream(Trim(fileName), FileMode.Open, FileAccess.Read)
            Dim sr As New StreamReader(fs)

            Dim str As String = sr.ReadLine

            Dim i As Integer = 0
            Do Until str Is Nothing
                If i >= 2 Then


                    Try
                        Dim splt() As String = str.Replace(""",""", ",").Split(",")



                        Dim CLIENT_TYPE As String = splt(0)
                        Dim FI_LONG As String = splt(1).ToString().Trim().Replace("""", "")
                        Dim FI_SHORT As String = splt(2).ToString().Trim().Replace("""", "")
                        Dim FS_LONG As String = splt(3).ToString().Trim().Replace("""", "")
                        Dim FS_SHORT As String = splt(4).ToString().Trim().Replace("""", "")
                        Dim OI_CALL_LONG As String = splt(5).ToString().Trim().Replace("""", "")
                        Dim OI_PUT_LONG As String = splt(6).ToString().Trim().Replace("""", "")
                        Dim OI_CALL_SHORT As String = splt(7).ToString().Trim().Replace("""", "")
                        Dim OI_PUT_SHORT As String = splt(8).ToString().Trim().Replace("""", "")
                        Dim OS_CALL_LONG As String = splt(9).ToString().Trim().Replace("""", "")
                        Dim OS_PUT_LONG As String = splt(10).ToString().Trim().Replace("""", "")
                        Dim OS_CALL_SHORT As String = splt(11).ToString().Trim().Replace("""", "")
                        Dim OS_PUT_SHORT As String = splt(12).ToString().Trim().Replace("""", "")
                        Dim TOTAL_LONG_CON As String = splt(13).ToString().Trim().Replace("""", "")
                        Dim TOTAL_SHORT_CON As String = splt(14).ToString().Trim().Replace("""", "")

                        Dim FILE_NAME As String = "fo_participant_oi" & Format(Dates, "ddMMyyyy") & ".csv"




                        strSql = " select * from FO_OI where CLIENT_TYPE ='" & CLIENT_TYPE & "' and file_name ='" & FILE_NAME & "'"
                        Dim dt As DataTable = New DataTable
                        dt = SqlHelper.ExecuteDataset(My.Settings.conn_str, CommandType.Text, strSql).Tables(0)

                        If dt.Rows.Count > 0 Then
                            strSql = "update FO_OI set CLIENT_TYPE ='" & CLIENT_TYPE & "'" &
                                " ,FI_SHORT='" & FI_SHORT & "' , FI_LONG ='" & FI_LONG & "'" &
                                    ", FS_SHORT='" & FS_SHORT & "', FS_LONG ='" & FS_LONG & "'" &
                                    ", OI_CALL_LONG='" & OI_CALL_LONG & "', OI_CALL_SHORT='" & OI_CALL_SHORT & "'" &
                                    ", OS_PUT_LONG ='" & OS_PUT_LONG & "', OS_PUT_SHORT='" & OS_PUT_SHORT & "'" &
                                    ", TOTAL_LONG_CON ='" & TOTAL_LONG_CON & "', TOTAL_SHORT_CON='" & TOTAL_SHORT_CON & "'" &
                                     ", updatedatetime = getdate(), FILE_NAME= '" & FILE_NAME & "'" &
                                    " where CLIENT_TYPE ='" & CLIENT_TYPE & "' and file_date='" & Format(Dates, "MM-dd-yy") & "'"
                            SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)





                        Else
                            strSql = "INSERT INTO FO_OI (CLIENT_TYPE,FI_SHORT,FI_LONG,FS_SHORT,FS_LONG,OI_CALL_LONG,OI_CALL_SHORT,OI_PUT_LONG,OI_PUT_SHORT,OS_CALL_LONG,OS_CALL_SHORT,OS_PUT_LONG,OS_PUT_SHORT,TOTAL_LONG_CON,TOTAL_SHORT_CON, updatedatetime, FILE_NAME, FILE_DATE)"
                            strSql &= "values('" & CLIENT_TYPE & "','" & FI_SHORT & "','" & FI_LONG & "','" & FS_SHORT & "','" & FS_LONG & "','" & OI_CALL_LONG & "','" & OI_CALL_SHORT & "','" & OI_PUT_LONG & "','" & OI_PUT_SHORT & "','" & OS_CALL_LONG & "'"
                            strSql &= ",'" & OS_CALL_SHORT & "','" & OS_PUT_LONG & "','" & OS_PUT_SHORT & "','" & TOTAL_LONG_CON & "','" & TOTAL_SHORT_CON & "',getdate(),'" & FILE_NAME & "','" & Format(Dates, "MM-dd-yy") & "')"
                            Console.WriteLine("Importing Record to FO PARTICIPANT OI :: " & strSql)

                            SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)
                            '' SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, Sql)
                        End If
                    Catch ex As Exception
                        Console.WriteLine("Update Main DB" & ex.Message)
                    Finally

                    End Try

                End If
nextloop:
                str = sr.ReadLine
                i = i + 1
            Loop
        Catch ex As Exception
            Console.WriteLine("Update Main DB" & ex.Message)
        End Try
        '  Next
    End Sub
#End Region
#Region "GetEODData From NSE"




    Sub ParseEOD()
        For i As Integer = 0 To 730
            ' Dates = Now.Date.AddDays(-5)
            If Directory.Exists(localDirectory & Format(Dates, "ddMMyy") & "\NSE_HISTORICAL") = False Then
                Directory.CreateDirectory(localDirectory & Format(Dates, "ddMMyy") & "\NSE_HISTORICAL")
            End If

            If Directory.Exists(localDirectory & Format(Dates, "ddMMyy") & "\NSE_HISTORICAL") = False Then
                Directory.CreateDirectory(localDirectory & Format(Dates, "ddMMyy") & "\NSE_HISTORICAL")
            End If

            If Directory.Exists(localDirectory & Format(Dates, "ddMMyy") & "\BSE_HISTORICAL") = False Then
                Directory.CreateDirectory(localDirectory & Format(Dates, "ddMMyy") & "\BSE_HISTORICAL")
            End If

            If Directory.Exists(localDirectory & Format(Dates, "ddMMyy") & "\BSE_HISTORICAL") = False Then
                Directory.CreateDirectory(localDirectory & Format(Dates, "ddMMyy") & "\BSE_HISTORICAL")
            End If

            Try

                Call NSE_EOD()
                Call BSE_EOD()
            Catch ex As Exception
                Console.WriteLine(Now.ToString() + " Error In  NSE " & ex.Message)
            End Try


        Next
    End Sub


#End Region

#Region "VAR MARGIN"

    Sub DownloadVarMarginFiles()



        If Directory.Exists(localDirectory & strddMMyy & "\var") = False Then
            Directory.CreateDirectory(localDirectory & strddMMyy & "\var")
        End If

        If Directory.Exists(localDirectory & Format(DateAdd("d", -1, Dates), "ddMMyy") & "\var") = False Then
            Directory.CreateDirectory(localDirectory & Format(DateAdd("d", -1, Dates), "ddMMyy") & "\var")
        End If



        If Directory.Exists(localDirectory & Format(DateAdd("d", -2, Dates), "ddMMyy") & "\var") = False Then
            Directory.CreateDirectory(localDirectory & Format(DateAdd("d", -2, Dates), "ddMMyy") & "\var")
        End If





        Call VarNumber("1", Dates)
        Call VarNumber("6", Dates)

        Call VarNumber("1", DateAdd("d", -1, Dates))
        Call VarNumber("6", DateAdd("d", -1, Dates))

        Call VarNumber("1", DateAdd("d", -2, Dates))
        Call VarNumber("6", DateAdd("d", -2, Dates))

    End Sub

    Sub VarNumber(fileNumber As String, localDates As Date)
        'https://www1.nseindia.com/archives/nsccl/var/C_VAR1_13012020_1.DAT
        Dim file2download As String = "C_VAR1_" & Format(localDates, "dd") & UCase(Format(localDates, "MM")) & Now.Year
        Dim url2download As String = "https://www1.nseindia.com/archives/nsccl/var/" & file2download & "_" & fileNumber & ".DAT"


        Dim localfile As String = Environment.CurrentDirectory & "\bhavfiles\" & Format(localDates, "ddMMyy") & "\var\" & fileNumber & ".dat"

        If (DownloadFile(url2download, localfile) = False) Then
            Exit Sub
        End If
        Console.WriteLine("VAR File Downloaded..." & url2download)
        'Dim fileName As String = Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\" & fileNumber & ".dat"
        FnoDataTable = ExcelToDataTable_NoFilter(localfile)
        Dim fs As New FileStream(Trim(localfile), FileMode.Open, FileAccess.Read)
        Dim sr As New StreamReader(fs)


        Dim str As String
        str = sr.ReadLine


        Dim i As Integer
        Dim dt As DataTable
        i = 0
        Dim sqlstr As String
        Dim security_var As Double
        Dim index_var As Double
        Dim var_margin As Double
        Dim extreme_loss_rate As Double
        Dim adhoc_margin As Double
        Dim applicable_margin_rate As Double
        Dim var_file_number As String
        Dim varfilename As String
        Dim STOCK_ITEMS_id As String
        Dim total As Double
        Try

            Do Until str Is Nothing
                If i <> 0 Then
                    Dim splt() As String = str.Split(",")



                    If (splt(2) = "EQ" Or splt(2) = "BE") Then
                        STOCK_ITEMS_id = splt(1).Trim()
                        security_var = splt(4)
                        index_var = splt(5)
                        var_margin = splt(6)
                        extreme_loss_rate = splt(7)
                        adhoc_margin = splt(8)
                        applicable_margin_rate = splt(9)
                        var_file_number = fileNumber
                        varfilename = Format(localDates, "ddMMyyyy") & "_" & var_file_number

                        total = security_var + index_var + var_margin + extreme_loss_rate + adhoc_margin + applicable_margin_rate
                        str = "select STOCK_ID, EXCHANGE_SYMBOL, ABBRS from STOCK_ITEMS where exchange_symbol ='" & STOCK_ITEMS_id & "'"
                        dt = SqlHelper.ExecuteDataset(My.Settings.conn_str, CommandType.Text, str).Tables(0)

                        If (dt.Rows.Count > 0) Then


                            Try



                                sqlstr = "IF EXISTS (SELECT NULL FROM VAR_MARGIN WHERE stock_id='" & dt.Rows(0)("stock_id") & "' and var_file_number ='" & var_file_number & "' and var_file_name ='" & varfilename & "')"
                                sqlstr = sqlstr & " BEGIN"
                                sqlstr = sqlstr & " Update var_margin Set var_file_date ='" & Format(localDates, "MM-dd-yy") & "', total='" & total & "' ,security_var='" & security_var & "' ,index_var='" & index_var & "',var_margin='" & var_margin & "',extreme_loss_rate='" & extreme_loss_rate & "',adhoc_margin='" & adhoc_margin & "',applicable_margin_rate='" & applicable_margin_rate & "'  where  stock_id ='" & dt.Rows(0)("stock_id") & "' and var_file_number ='" & var_file_number & "' and var_file_name ='" & varfilename & "'"
                                sqlstr = sqlstr & " END"
                                sqlstr = sqlstr & " ELSE"
                                sqlstr = sqlstr & " BEGIN"
                                sqlstr = sqlstr & " insert into var_margin(var_file_date,total,stock_id,security_var,index_var,var_margin, extreme_loss_rate, adhoc_margin, applicable_margin_rate, var_file_number, var_file_name) values ('" & Format(localDates, "MM-dd-yy") & "','" & total & "','" & dt.Rows(0)("stock_id") & "','" & security_var & "','" & index_var & "','" & var_margin & "','" & extreme_loss_rate & "','" & adhoc_margin & "','" & applicable_margin_rate & "','" & var_file_number & "','" & varfilename & "') "
                                sqlstr = sqlstr & " END"
                                SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, sqlstr)

                                Console.WriteLine("UPDATE VAR  for " & dt.Rows(0)("ABBRS") & ". File Name: " & varfilename & ". File Number: " & var_file_number)

                                clsWrite.CaptureLogs(Environment.CurrentDirectory & "\bhavfiles\" & Format(localDates, "ddMMyy") & "\var\", sqlstr)
                            Catch ex As Exception
                                clsWrite.CaptureLogs(Environment.CurrentDirectory & "\bhavfiles\" & Format(localDates, "ddMMyy") & "\var\", MethodBase.GetCurrentMethod().ToString() & " ==> " & sqlstr, "ERR")
                            End Try
                        End If
                    End If





                End If
                str = sr.ReadLine
                i = i + 1



            Loop



        Catch ex As Exception
            clsWrite.CaptureLogs(Environment.CurrentDirectory & "\bhavfiles\" & Format(localDates, "MM-dd-yy") & "\var\", sqlstr, "ERR")
        End Try
    End Sub
#End Region

#Region "DOWNLOAD CIRCUIT BHAV"
    Sub Download_Circuit_NSE()
        Try
            url2download = "https://archives.nseindia.com/content/equities/sec_list.csv"


            Dim localfile As String = Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\CIRCUIT_NSE\sec_list.csv"
            If (DownloadFile(url2download, localfile)) Then
                startTime = Now
                Console.WriteLine("Circuit DB Updated..." & startTime)
                Call UpdateStartTime("START_DATE_TIME")
                Call ReadCSVFile_CIRCUIT(localfile)
                'Call UpdateCircuit()
                Call UpdateStartTime("END_DATE_TIME")
                Console.WriteLine("CIRCUIT UPDATED In DB ..." & Now)
                'Console.ReadLine()
            End If
        Catch ex As Exception
            Console.WriteLine("CIRCUITS " & ex.Message)
        End Try
    End Sub




    Sub UpdateCircuit(ByVal arr As Array)
        Try
            Dim i As Integer
            Dim sql As String
            Dim upper_circuit, lower_circuit, close, isin As String
            Dim exch As String = arr(0).ToString().Trim()
            Dim series As String = arr(1).ToString().Trim()
            Dim circuitBand As String = arr(3).ToString().Trim()

            If exch = "ZEELEARN" Then
                Dim aa As String
                aa = ""
            End If

            If (series.Trim.ToUpper = "EQ" Or series.Trim.ToUpper = "BE") Then
                series = "EQ"
            End If

            If (exch = "ARE&M") Then
                Dim ss As String = ""
            End If

            sql = "Select EXCHANGE_SYMBOL,PREVDAY_CLOSE as prev_close,isnull(isin,'') as isin from STOCK_ITEMS where exchange_symbol ='" & exch & "' and series = '" & series & "' and exchange_id in (1,2)"
            Dim dt As DataTable = New DataTable
            Dim dr() As DataRow
            dt = SqlHelper.ExecuteDataset(My.Settings.conn_str, CommandType.Text, sql).Tables(0)
            dr = dt.Select
            For i = 0 To dr.Length - 1
                If circuitBand.ToUpper.Trim() = "NO BAND" Then
                    Continue For
                End If
                Try
                    close = dr(i)("prev_close")
                    isin = dr(i)("isin")

                    If Trim(circuitBand) <> "0" Then
                        upper_circuit = close + (CInt(Trim(circuitBand)) / 100) * close

                    Else
                        upper_circuit = "0"
                    End If

                    If Trim(circuitBand) <> "0" Then
                        lower_circuit = close - (CInt(circuitBand) / 100) * close


                    Else
                        lower_circuit = "0"
                    End If
                    'upper_circuit = 0.2 * upper_circuit
                    'lower_circuit = 0.2 * upper_circuit
                    If isin = "" Then
                        sql = "UPDATE STOCK_ITEMS Set UPPER_CIRCUIT='" & upper_circuit & "',LOWER_CIRCUIT='" & lower_circuit & "' where exchange_symbol ='" & Trim(exch) & "' and exchange_id in(1,8)"
                    Else
                        sql = "UPDATE STOCK_ITEMS Set UPPER_CIRCUIT='" & upper_circuit & "',LOWER_CIRCUIT='" & lower_circuit & "' where isin='" & Trim(isin) & "' and series ='" & series & "'"
                    End If


                    Try
                        SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, sql)
                        clsWrite.CaptureUpdateLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\CIRCUIT_NSE\", sql)
                    Catch ex As Exception
                        Console.WriteLine("CIRCUIT UPDATETRANSACTION 1 " & ex.Message)
                        clsWrite.CaptureLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\CIRCUIT_NSE\", sql, "err")
                    End Try
                    Try
                        SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, sql)
                    Catch ex As Exception
                        Console.WriteLine("CIRCUIT UPDATETRANSACTION 2 " & ex.Message)
                    End Try
                    Console.WriteLine("Updating circuit for " + exch.ToString().Trim())
                Catch ex As Exception
                    Console.WriteLine("CIRCUIT UPDATETRANSACTION 3 " & ex.Message)
                End Try
            Next


        Catch ex As Exception

        End Try
    End Sub

    Sub ReadIndices()
        Dim sqlstr As String = "SELECT exchangesymbol FROM equity_transaction_data WHERE exchangeid='1732257' AND Series='INX'"
        Dim dt As New DataTable
        dt = SqlHelper.ExecuteDataset(My.Settings.conn_str, CommandType.Text, sqlstr).Tables(0)
        For i As Integer = 0 To dt.Rows.Count - 1
            indices &= Trim(dt.Rows(i)("Exchangesymbol")) & ","
        Next
    End Sub

    Sub ReadCSVFile_CIRCUIT(ByVal fileName As String)
        Try
            Dim fs As New FileStream(Trim(fileName), FileMode.Open, FileAccess.Read)
            Dim sr As New StreamReader(fs)

            Dim str As String
            str = sr.ReadLine

            'Loop till the last line...
            Dim i As Integer
            i = 0
            Do Until str Is Nothing
                If i <> 0 Then
                    Dim arr1 As Array
                    arr1 = Split(str, ",")
                    UpdateCircuit(arr1)
                    Console.WriteLine("CIRCUIT-" & str)
                End If
                str = sr.ReadLine
                i = 1
            Loop
        Catch ex As Exception
            Console.WriteLine(" ERR IN ReadCSVFile_CIRCUIT " & fileName & " ==> " & ex.Message)
        End Try
    End Sub

    'Sub Insert_BandChange(ByVal arr As Array)
    '    Dim sqlstr As String
    '    sqlstr = "IF EXISTS (SELECT NULL FROM CIRCUIT WHERE EXCHANGESYMBOL='" & arr(0).ToString().Trim() & "' and SERIES ='" & arr(1).ToString().Trim() & "')"
    '    sqlstr = sqlstr & " BEGIN"
    '    sqlstr = sqlstr & " Update CIRCUIT Set CIRCUIT_PERC='" & Replace(arr(3).ToString().Trim(), "No Band", "0") & "',INSERT_DATE_TIME='" & Trim(Now) & "' where EXCHANGESYMBOL='" & arr(0).ToString().Trim() & "' and SERIES ='" & arr(1).ToString().Trim() & "'"
    '    sqlstr = sqlstr & " END"
    '    sqlstr = sqlstr & " ELSE"
    '    sqlstr = sqlstr & " BEGIN"
    '    sqlstr = sqlstr & " Insert Into Circuit (ExchangeSymbol,CIRCUIT_PERC,INSERT_DATE_TIME,SERIES) Values ('" & arr(0).ToString().Trim() & "','" & Replace(arr(3).ToString().Trim(), "No Band", "0") & "','" & Trim(Now) & "','" & arr(1).ToString().Trim() & "')"
    '    sqlstr = sqlstr & " END"
    '    Try
    '        SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, sqlstr)
    '    Catch ex As Exception
    '        clsWrite.CaptureLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\Insert_BandChange\", MethodBase.GetCurrentMethod().ToString() & " ==> " & sqlstr, "ERR")
    '    End Try
    '    Try
    '        SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, sqlstr)
    '    Catch ex As Exception
    '        clsWrite.CaptureLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\Insert_BandChange\", MethodBase.GetCurrentMethod().ToString() & " ==> " & sqlstr, "ERR")
    '    End Try
    'End Sub
#End Region

#Region "DOWNLOAD FREE FLOAT NOT USED"
    Sub ReadCSVFile_freefloat(ByVal fileName As String)

        Dim fs As New FileStream(Trim(fileName), FileMode.Open, FileAccess.Read)
        Dim sr As New StreamReader(fs)

        Dim str As String
        str = sr.ReadLine
        'Loop till the last line...
        Dim i As Integer
        i = 0
        Do Until str Is Nothing

            If i <> 0 Then
                Dim arr1 As Array
                arr1 = Split(str, ",")
                UpdateDB_freefloat(arr1)
                Console.WriteLine("NSE-" & str)
            End If
            str = sr.ReadLine
            i = 1
        Loop
        sr.Close()
        fs.Close()
    End Sub

    Sub UpdateDB_freefloat(ByVal arr As Array)
        If arr(0) <> " " And arr(0) <> "" Then
            'symbol,series,open,high,low,close,last,prevclose,ttrdqty,ttrdval,timestamp

            Dim symbol As String

            symbol = SqlHelper.ExecuteScalar(My.Settings.conn_str, CommandType.Text, "SELECT LEFT(symbol,6) FROM dbo.VEHICLE_EXCHANGE WHERE ExchangeSymbol='" & Trim(arr(1)) & "' AND Series='EQUITY'")

            If symbol = "" Then
                Exit Sub
            End If
            Dim sql As String = ""
            sql &= "   Update company set company_name='" & Trim(arr(3)) & "',freefloat='" & Trim(CDbl(arr(5))) & "'"
            sql &= " where symbol like '" & symbol & "%'"

            Try
                SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, sql)

            Catch ex As Exception
                Console.WriteLine("Update Main DB" & ex.Message)
            End Try
            Try
                SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, sql)
            Catch ex As Exception
                Console.WriteLine("Update Bkup DB" & ex.Message)
            End Try
            'Console.WriteLine("NSE-" & Trim(arr(3)))
            STOCK_ITEMS_updated += 1
        End If
    End Sub
#End Region

#Region "USER DEFINED EVENTS AND SUB"
    Sub UpdateStartTime(var As String)
        STOCK_ITEMS_updated = 0
        Dim sql As String
        sql = "update  JOB_REPORTS set " & var & " = getdate() where  NAME = 'NSE_BHAVCOPY_CIRCUIT'"
        Try
            SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, sql)
        Catch ex As Exception

        End Try
        Try
            SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, sql)
        Catch ex As Exception

        End Try

    End Sub




    Public Function DownloadFile(ByVal url As String, ByVal localfile As String) As Boolean
        Try


            ' file2download = "eq_bands_24082016.csv"

            ''https://archives.nseindia.com/content/equities/sec_list.csv


            ServicePointManager.Expect100Continue = True
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12


            Dim _WebClient As New WebClientWithTimeout()
            _WebClient.UseDefaultCredentials = True
            _WebClient.Headers("User-Agent") = "Mozilla/5.0 (compatible; MSIE 9.0; windows NT 6.1; WOW64; Trident/5.0)"
            _WebClient.Headers("Method") = "GET"
            _WebClient.Headers("AllowAutoRedirect") = True
            _WebClient.Headers("KeepAlive") = True

            '_WebClient.Headers("Accept") = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
            clsWrite.CaptureLogs(localDirectory & strddMMyy & "\DownloadFile\", " ===> " & url, "info")
            Try
                ''    Console.WriteLine("DownloadFile 1  " & url)
                _WebClient.DownloadFile(New Uri(url), localfile)
                ''     Console.WriteLine("DownloadFile 2")
                _WebClient.Dispose()
                ''     Console.WriteLine("DownloadFile 3")
                _WebClient = Nothing
                ''     Console.WriteLine("DownloadFile 4")
                ''    Console.ReadLine()


                Return True
            Catch ex As Exception

                Console.WriteLine(" ERR IN DOWNLOADING " & url & " ==> " & ex.Message)
                clsWrite.CaptureLogs(localDirectory & strddMMyy & "\DownloadFile\", ex.Message & " ===> " & url, "err")
                '' Console.ReadLine()
                Return False
            End Try
        Catch ex As Exception
            Console.WriteLine("XXXXXXX")
            Return False
        End Try
    End Function




    Sub UnzipFile(ByVal ZipToUnpack As String, ByVal file2extract As String, ByVal targetdir As String)
        Using zip1 As ZipFile = ZipFile.Read(ZipToUnpack)
            Dim ZP As ZipEntry
            For Each ZP In zip1
                If UCase(ZP.FileName) = UCase(file2extract) Then
                    ZP.Extract(targetdir, ExtractExistingFileAction.OverwriteSilently)
                    Exit For
                End If

            Next
        End Using
    End Sub
#End Region

#Region "TEST NSE"


    Private Sub readPESyndication()

        Dim req As WebRequest = WebRequest.Create("https://syndication.bloomberg.com/finance/v2/reference?securities=MDF:IB,HDFC:IB&fields=PE_RATIO,BEST_PE_RATIO,BEST_PEG_RATIO,EQY_SH_OUT,PX_TO_BOOK_RATIO,PX_TO_SALES_RATIO,EQY_TRR_PCT_1YR,VOLUME_AVG_30D,TRAIL_12M_EPS,BEST_EEPS_CUR_YR,EQY_DVD_YLD_IND,DVD_SH_LAST,EXPECTED_REPORT_DT,LATEST_ANNOUNCEMENT_PERIOD,EXPECTED_REPORT_PERIOD_END_DATE")
        req.Method = "GET"





        Dim MyCon As New Odbc.OdbcConnection
        MyCon.ConnectionString = "Driver={PostgreSQL ANSI};database=database_name;server=127.0.0.1;port=5432;uid=postgres;sslmode=disable;readonly=0;protocol=7.4;User ID=postgres;password=password;"

        MyCon.Open()
        If MyCon.State = ConnectionState.Open Then
            MsgBox("Connected To PostGres", MsgBoxStyle.MsgBoxSetForeground)
        End If


        Dim username As String = "syn_d7d4cd5e"
        Dim password As String = "FSbc6Z]IWw?r"
        req.Credentials = New NetworkCredential(username, password)
        Dim resp As HttpWebResponse = TryCast(req.GetResponse(), HttpWebResponse)

        Dim responseJson As String = New StreamReader(resp.GetResponseStream()).ReadToEnd()
        Dim root As RootObject = JsonConvert.DeserializeObject(Of RootObject)(responseJson)

        For Each key As String In root.data.Keys
            Dim da As Data = root.data(key)

        Next

    End Sub


    ''''    Sub DownloadEodFromNSE()
    ''''        Dim m As Integer = 0
    ''''        Dim dates As Date
    ''''        For m = 365 To 0 Step -1
    ''''            dates = DateAdd(DateInterval.Day, -m, DateTime.Now.Date)

    ''''            If Directory.Exists(localDirectory & Format(dates, "ddMMyy") & "\NSE") = False Then
    ''''                Directory.CreateDirectory(localDirectory & Format(dates, "ddMMyy") & "\NSE")
    ''''            End If

    ''''            Dim file2download As String = "PR" & Format(dates, "ddMMyy") & ".zip"
    ''''            Dim url2download As String = "https://www.nseindia.com/archives/equities/bhavcopy/pr/" & file2download
    ''''            Dim localfile As String = Environment.CurrentDirectory & "\bhavfiles\" & Format(dates, "ddMMyy") & "\NSE\" & file2download
    ''''            If (DownloadFile(url2download, localfile)) Then


    ''''                UnzipFile(localfile, "Pd" & Format(dates, "ddMMyy") & ".csv", Environment.CurrentDirectory & "\bhavfiles\" & Format(dates, "ddMMyy") & "\NSE")





    ''''                Dim fs As New FileStream(Trim(Environment.CurrentDirectory & "\bhavfiles\" & Format(dates, "ddMMyy") & "\NSE\" & "Pd" & Format(dates, "ddMMyy") & ".csv"), FileMode.Open, FileAccess.Read)
    ''''                Dim sr As New StreamReader(fs)

    ''''                Dim str As String = sr.ReadLine

    ''''                Dim i As Integer = 0

    ''''                Do Until str Is Nothing
    ''''                    If i <> 0 Then
    ''''                        Try
    ''''                            Dim splt() As String = str.Split(",")
    ''''                            Dim exch As String
    ''''                            Dim Fhi_52_wk As Double
    ''''                            Dim Flo_52_wk As Double
    ''''                            Dim Open As Double
    ''''                            Dim High As Double
    ''''                            Dim Low As Double
    ''''                            Dim PrevClose As Double
    ''''                            Dim lastprice As Double
    ''''                            Dim Volume As Long
    ''''                            Dim Name As String
    ''''                            Dim Series As String
    ''''                            Dim isIndex As String

    ''''                            isIndex = splt(0).ToString().Trim()


    ''''                            If (isIndex = "Y") Then
    ''''                                exch = splt(3).ToString().Trim()

    ''''                                'If (exch.ToUpper() = "NIFTY 50") Then
    ''''                                '  exch = "S&P CNX NIFTY"
    ''''                                ' End If
    ''''                            ElseIf (isIndex = "N") Then

    ''''                                exch = splt(2).ToString().Trim()


    ''''                            Else
    ''''                                GoTo nextloop
    ''''                            End If

    ''''                            Series = splt(1).ToString().Trim()
    ''''                            Name = splt(3).ToString().Trim()
    ''''                            PrevClose = splt(4).ToString().Trim()
    ''''                            Open = splt(5).ToString().Trim()
    ''''                            High = splt(6).ToString().Trim()
    ''''                            Low = splt(7).ToString().Trim()
    ''''                            lastprice = splt(8).ToString().Trim()
    ''''                            Volume = splt(10).ToString().Trim()
    ''''                            Fhi_52_wk = splt(14).ToString().Trim()
    ''''                            Flo_52_wk = splt(15).ToString().Trim()

    ''''                            Dim Sql As String = " select id  from STOCK_ITEMS where exchange_symbol ='" & exch & "'"
    ''''                            Dim dt As DataTable = New DataTable
    ''''                            dt = SqlHelper.ExecuteDataset(My.Settings.conn_str, CommandType.Text, Sql).Tables(0)

    ''''                            Dim STOCK_ITEMS_id As Integer

    ''''                            If (dt.Rows.Count > 0) Then

    ''''                                STOCK_ITEMS_id = dt.Rows(0)("id")
    ''''                                Sql = " select *  from STOCK_ITEMS_EOD where exchange_symbol ='" & exch & "' and trade_date = '" & dates.ToShortDateString() & "'"

    ''''                                Try
    ''''                                    dt = SqlHelper.ExecuteDataset(My.Settings.conn_str, CommandType.Text, Sql).Tables(0)
    ''''                                    If dt.Rows.Count > 0 Then
    ''''                                        'If (Convert.ToDouble(dt.Rows(0)("LAST_PRICE")).ToString("#.00") <> Convert.ToDouble(lastprice).ToString("#.00")) Then
    ''''                                        '    Dim strSql As String = ""
    ''''                                        '    Try

    ''''                                        '        If isIndex.Trim() = "Y" Then '' INDEX

    ''''                                        '            If (exch.Trim() <> "") Then
    ''''                                        '                strSql = "update STOCK_ITEMS set LAST_PRICE ='" & lastprice & "',ACC_VOLUME='" & Volume & "',  DAY_OPEN='" & Open & "',DAY_HIGH='" & High & "'" &
    ''''                                        '                " ,DAY_LOW='" & Low & "', PREV_CLOSE ='" & PrevClose & "',[hi_52_wk] ='" & Fhi_52_wk & "',[lo_52_wk] ='" & Flo_52_wk & "'  where exchange_symbol ='" & exch & "' and exchange_id =1 and instrument_id =1"
    ''''                                        '                SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)
    ''''                                        '                'Debug.WriteLine(splt(2).ToString().Trim() & " " & " DB VALUE " & dt.Rows(0)("LAST_PRICE") & ". EXCHANGE VALUE " & splt(8).ToString().Trim() & " " & splt(1).ToString().Trim())
    ''''                                        '                clsWrite.CaptureUpdateLogs(Environment.CurrentDirectory & "\bhavfiles\" & Format(dates, "ddMMyy") & "\NSE\", strSql)
    ''''                                        '                NSEcounterUpdate = NSEcounterUpdate + 1
    ''''                                        '            End If
    ''''                                        '        ElseIf isIndex.Trim() = "N" Then
    ''''                                        '            If (Series.Trim().ToUpper() = "EQ" Or Series.Trim().ToUpper() = "BE") Then '' STOCK_ITEMS
    ''''                                        '                If (exch.Trim() <> "") Then
    ''''                                        '                    strSql = "update STOCK_ITEMS set LAST_PRICE ='" & lastprice & "',ACC_VOLUME='" & Volume & "',  DAY_OPEN='" & Open & "',DAY_HIGH='" & High & "'" &
    ''''                                        '                    " ,DAY_LOW='" & Low & "', PREV_CLOSE ='" & PrevClose & "',[hi_52_wk] ='" & Fhi_52_wk & "',[lo_52_wk] ='" & Flo_52_wk & "'  where exchange_symbol ='" & exch & "' and series ='" & Series & "' and exchange_id =1 and instrument_id =2"
    ''''                                        '                    SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)
    ''''                                        '                    'Debug.WriteLine(splt(2).ToString().Trim() & " " & " DB VALUE " & dt.Rows(0)("LAST_PRICE") & ". EXCHANGE VALUE " & splt(8).ToString().Trim() & " " & splt(1).ToString().Trim())
    ''''                                        '                    clsWrite.CaptureUpdateLogs(Environment.CurrentDirectory & "\bhavfiles\" & Format(dates, "ddMMyy") & "\NSE\", strSql)
    ''''                                        '                    NSEcounterUpdate = NSEcounterUpdate + 1
    ''''                                        '                End If
    ''''                                        '            End If
    ''''                                        '        End If

    ''''                                        '    Catch ex As Exception
    ''''                                        '        clsWrite.CaptureLogs(Environment.CurrentDirectory & "\bhavfiles\" & Format(dates, "ddMMyy") & "\NSE\", ex.Message & " ===> " & strSql, "err")
    ''''                                        '    End Try

    ''''                                        'Else

    ''''                                        'End If
    ''''                                    Else

    ''''                                        Dim strSql As String = ""
    ''''                                        Try

    ''''                                            If (isIndex.Trim() = "Y") Then
    ''''                                                If (exch.Trim() <> "") Then
    ''''                                                    strSql = "insert into STOCK_ITEMS_EOD (" &
    ''''                                                    "  STOCK_ITEMS_id, day_open, day_high, day_low " &
    ''''                                                    " ,day_close, prev_close" &
    ''''                                                    " ,acc_volume, exchange_symbol,trade_date, update_date_time" &
    ''''                                                    " ,insert_date_time" &
    ''''                                                    ") values" &
    ''''                                                    "('" & STOCK_ITEMS_id & "','" & Open & "','" & High & "'," &
    ''''                                                    " '" & Low & "','" & lastprice & "','" & PrevClose & "','" & Volume & "','" & exch & "','" & dates.ToShortDateString() & "'," &
    ''''                                                    " getdate(),getDate())"

    ''''                                                    SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)

    ''''                                                    clsWrite.CaptureInsertsLogs(Environment.CurrentDirectory & "\bhavfiles\" & Format(dates, "ddMMyy") & "\NSE\", strSql)
    ''''                                                    NSEcounterInsert = NSEcounterInsert + 1
    ''''                                                End If
    ''''                                            ElseIf isIndex = "N" Then

    ''''                                                If (Series.Trim().ToUpper() = "EQ" Or Series.Trim().ToUpper() = "BE") Then '' STOCK_ITEMS
    ''''                                                    If (exch.Trim() <> "") Then
    ''''                                                        strSql = "insert into STOCK_ITEMS_EOD (" &
    ''''                                                        "  STOCK_ITEMS_id, day_open, day_high, day_low " &
    ''''                                                        " ,day_close, prev_close" &
    ''''                                                        " ,acc_volume, exchange_symbol,trade_date, update_date_time" &
    ''''                                                        " ,insert_date_time" &
    ''''                                                        ") values" &
    ''''                                                        "('" & STOCK_ITEMS_id & "','" & Open & "','" & High & "'," &
    ''''                                                        " '" & Low & "','" & lastprice & "','" & PrevClose & "','" & Volume & "','" & exch & "','" & dates.ToShortDateString() & "'," &
    ''''                                                        " getdate(),getDate())"

    ''''                                                        SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)

    ''''                                                        clsWrite.CaptureInsertsLogs(Environment.CurrentDirectory & "\bhavfiles\" & Format(dates, "ddMMyy") & "\NSE\", strSql)
    ''''                                                        NSEcounterInsert = NSEcounterInsert + 1
    ''''                                                    End If
    ''''                                                Else


    ''''                                                End If
    ''''                                            End If


    ''''                                        Catch ex As Exception
    ''''                                            clsWrite.CaptureLogs(Environment.CurrentDirectory & "\bhavfiles\" & Format(dates, "ddMMyy") & "\NSE\", ex.Message & " ===> " & strSql, "err")
    ''''                                        End Try


    ''''                                    End If

    ''''                                Catch ex As Exception
    ''''                                    Console.WriteLine("Update Main DB" & ex.Message)
    ''''                                End Try
    ''''                            Else
    ''''                                clsWrite.CaptureLogs(Environment.CurrentDirectory & "\bhavfiles\" & Format(dates, "ddMMyy") & "\NSE\", "no data forund in STOCK_ITEMS table for " & exch)
    ''''                            End If
    ''''                        Catch ex As Exception

    ''''                        End Try
    ''''                    End If
    ''''nextloop:
    ''''                    str = sr.ReadLine
    ''''                    i = 1
    ''''                Loop



    ''''                'File.Delete(localfile)
    ''''            End If

    ''''        Next
    ''''    End Sub




    Sub NSE()
        'https://www.nseindia.com/archives/equities/bhavcopy/pr/PR040613.zip
        Dim file2download As String = "Pd" & strddMMyy & ".csv"
        Dim url2download As String = "https://archives.nseindia.com/archives/equities/bhavcopy/pr/PR" & strddMMyy & ".zip"
        ''https://archives.nseindia.com/archives/equities/bhavcopy/pr/PR120523.zip
        Dim localfile As String = localDirectory & strddMMyy & "\NSE\PR.zip"
        DownloadFile(url2download, localfile)
        UnzipFile(localfile, file2download, localDirectory & strddMMyy & "\NSE")


        File.Delete(localfile)

        Console.WriteLine("NSE File Downloaded..." & file2download)

        Dim fileName As String = Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\NSE\" & file2download
        Dim fs As New FileStream(Trim(fileName), FileMode.Open, FileAccess.Read)
        Dim sr As New StreamReader(fs)

        Dim str As String = sr.ReadLine

        Dim i As Integer = 0

        Do Until str Is Nothing
            If i <> 0 Then
                Try
                    Dim splt() As String = str.Split(",")
                    Dim exch As String
                    ' Dim Fhi_52_wk As Double '' NOT TAKING THIS SINCE BHAV COPY HAS UNADJUSTED RATES.
                    ' Dim Flo_52_wk As Double '' NOT TAKING THIS SINCE BHAV COPY HAS UNADJUSTED RATES.
                    Dim Open As Double
                    Dim High As Double
                    Dim Low As Double
                    Dim PrevClose As Double
                    Dim lastprice As Double
                    Dim Volume As Long
                    Dim Name As String
                    Dim Series As String
                    Dim isIndex As String
                    Dim exchange_id As String
                    isIndex = splt(0).ToString().Trim()

                    Series = splt(1).ToString().Trim()
                    If (isIndex = "Y") Then
                        exch = splt(3).ToString().Trim()
                        Series = "INX"
                        exchange_id = 8
                    ElseIf (isIndex = "N") Then
                        exch = splt(2).ToString().Trim()
                        If (Series.Trim.ToUpper = "EQ" Or Series.Trim.ToUpper = "BE") Then
                            Series = "EQ"
                        End If
                        exchange_id = 1
                    Else
                        GoTo nextloop
                    End If




                    Name = splt(3).ToString().Trim()
                    PrevClose = splt(4).ToString().Trim()
                    Open = splt(5).ToString().Trim()
                    High = splt(6).ToString().Trim()
                    Low = splt(7).ToString().Trim()
                    lastprice = splt(8).ToString().Trim()
                    Volume = splt(10).ToString().Trim()
                    ' Fhi_52_wk = splt(13).ToString().Trim()
                    '  Flo_52_wk = splt(14).ToString().Trim()

                    Dim Sql As String
                    If (Series.Trim().ToUpper() = "EQ" Or Series.Trim().ToUpper() = "BE" Or Series.Trim().ToUpper() = "ST" Or Series.Trim().ToUpper() = "SM" Or Series.Trim().ToUpper() = "BZ") Then
                        Sql = " select isnull(LAST_PRICE,0) as LAST_PRICE from STOCK_ITEMS where EXCHANGE_SYMBOL ='" & exch & "' and exchange_id ='" & exchange_id & "' and series in('EQ','BE','BZ','ST','SM')"
                    Else

                        Sql = " select isnull(LAST_PRICE,0) as LAST_PRICE from STOCK_ITEMS where EXCHANGE_SYMBOL ='" & exch & "' and exchange_id ='" & exchange_id & "' and series ='" & Series & "'"
                    End If


                    If (exch = "INNOVANA") Then
                        Dim strdd As String = "INNOVANA"
                    End If
                    Dim dt As DataTable = New DataTable
                    Try
                        dt = SqlHelper.ExecuteDataset(My.Settings.conn_str, CommandType.Text, Sql).Tables(0)
                        If dt.Rows.Count > 0 Then
                            If (Convert.ToDouble(dt.Rows(0)("LAST_PRICE")).ToString("#.00") <> Convert.ToDouble(lastprice).ToString("#.00")) Then
                                Dim strSql As String = ""
                                Try

                                    If isIndex.Trim() = "Y" Then '' INDEX 

                                        If (exch.Trim() <> "") Then
                                            strSql = "update STOCK_ITEMS set UPDATE_DATE_TIME_BHAV=getDate(),LAST_PRICE ='" & lastprice & "',ACC_VOLUME='" & Volume & "',  DAY_OPEN='" & Open & "',DAY_HIGH='" & High & "'" &
                                                " ,DAY_LOW='" & Low & "', PREVDAY_CLOSE ='" & PrevClose & "', group_name ='" & Series & "'  where exchange_symbol ='" & exch & "' and exchange_id in (1,2,8) and series ='" & Series & "'"


                                            SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)
                                            ' SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)
                                            'Debug.WriteLine(splt(2).ToString().Trim() & " " & " DB VALUE " & dt.Rows(0)("LAST_PRICE") & ". EXCHANGE VALUE " & splt(8).ToString().Trim() & " " & splt(1).ToString().Trim())
                                            clsWrite.CaptureUpdateLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\NSE\", strSql)
                                            Console.WriteLine("Update NSE STOCK " & exch & " " & Series)
                                            NSEcounterUpdate = NSEcounterUpdate + 1
                                        End If
                                    ElseIf isIndex.Trim() = "N" Then
                                        If (Series.Trim().ToUpper() = "EQ" Or Series.Trim().ToUpper() = "BE" Or Series.Trim().ToUpper() = "ST" Or Series.Trim().ToUpper() = "SM" Or Series.Trim().ToUpper() = "BZ") Then '' STOCK_ITEMS
                                            If (exch.Trim() <> "") Then
                                                strSql = "update STOCK_ITEMS set UPDATE_DATE_TIME_BHAV=getDate(),LAST_PRICE ='" & lastprice & "',ACC_VOLUME='" & Volume & "',  DAY_OPEN='" & Open & "',DAY_HIGH='" & High & "'" &
                                                    " ,DAY_LOW='" & Low & "', PREVDAY_CLOSE ='" & PrevClose & "', group_name ='" & Series & "',series='" & Series & "'  where exchange_symbol ='" & exch & "' and exchange_id in (1,2,8) and series in('EQ','BE','BZ','ST','SM')"
                                                SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)
                                                '  SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)
                                                'Debug.WriteLine(splt(2).ToString().Trim() & " " & " DB VALUE " & dt.Rows(0)("LAST_PRICE") & ". EXCHANGE VALUE " & splt(8).ToString().Trim() & " " & splt(1).ToString().Trim())


                                                Console.WriteLine("Update NSE STOCKS " & exch & " " & Series)
                                                clsWrite.CaptureUpdateLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\NSE\", strSql)
                                                NSEcounterUpdate = NSEcounterUpdate + 1
                                            End If
                                        Else
                                            If (exch.Trim() <> "") Then
                                                strSql = "update STOCK_ITEMS set UPDATE_DATE_TIME_BHAV=getDate(),LAST_PRICE ='" & lastprice & "',ACC_VOLUME='" & Volume & "',  DAY_OPEN='" & Open & "',DAY_HIGH='" & High & "'" &
                                                    " ,DAY_LOW='" & Low & "', PREVDAY_CLOSE ='" & PrevClose & "', group_name ='" & Series & "',  series ='" & Series & "'  where exchange_symbol ='" & exch & "' and exchange_id in (1,2,8) and   series ='" & Series & "'"
                                                SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)
                                                '  SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)
                                                'Debug.WriteLine(splt(2).ToString().Trim() & " " & " DB VALUE " & dt.Rows(0)("LAST_PRICE") & ". EXCHANGE VALUE " & splt(8).ToString().Trim() & " " & splt(1).ToString().Trim())


                                                Console.WriteLine("Update NSE INDEX " & exch & " " & Series)
                                                clsWrite.CaptureUpdateLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\NSE\", strSql)
                                                NSEcounterUpdate = NSEcounterUpdate + 1
                                            End If
                                        End If
                                    End If

                                Catch ex As Exception
                                    clsWrite.CaptureLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\NSE\", ex.Message & " ===> " & strSql, "err")
                                End Try

                            Else

                            End If
                        Else
                            'Debug.WriteLine(splt(2).ToString().Trim() & " " & splt(1).ToString().Trim())
                            Dim strSql As String = ""
                            Try
                                ' If isIndex.Trim() = "N" And (Series.Trim().ToUpper() = "EQ" Or Series.Trim().ToUpper() = "BE") Then '' STOCK_ITEMS
                                If (exch.Trim() <> "") Then
                                    strSql = "insert into STOCK_ITEMS (" &
                                    "  exchange_id, Series, NAME, ABBRS, ABBRL," &
                                    " exchange_symbol,last_price,day_open,day_high,day_low, PREVDAY_CLOSE," &
                                    " acc_volume,trade_volume,face_value,update_date_time," &
                                    " upper_circuit,lower_circuit, group_name,update_date_time_bhav" &
                                    ") values" &
                                    "('" & exchange_id & "','" & Series & "','" & Name & "','" & Name & "', '" & Name & "'," &
                                    " '" & exch & "','" & lastprice & "','" & Open & "','" & High & "','" & Low & "','" & PrevClose & "'," &
                                    " '" & Volume & "',0,0,getDate()," &
                                    " 0,0, '" & Series & "',getdate()" &
                                    ")"
                                    SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)
                                    ' SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)
                                    Console.WriteLine("Insert NSE " & exch & " " & Series)
                                    clsWrite.CaptureInsertsLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\NSE\", strSql)
                                    NSEcounterInsert = NSEcounterInsert + 1
                                End If
                                ' End If
                            Catch ex As Exception
                                clsWrite.CaptureLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\NSE\", ex.Message & " ===> " & strSql, "err")
                            End Try


                        End If

                    Catch ex As Exception
                        Console.WriteLine("Update Main DB" & ex.Message)
                    End Try
                Catch ex As Exception

                End Try
            End If
nextloop:
            str = sr.ReadLine
            i = 1
        Loop
    End Sub
    Sub BSE_EOD()
        'https://archives.nseindia.com/content/historical/EQUITIES/2021/JAN/cm19JAN2021bhav.csv.zip


        strddMMyy = Format(Dates, "ddMMyy")

        ''https://www.bseindia.com/download/BhavCopy/Equity/EQ021222_CSV.ZIP
        Dim file2download As String = "EQ" & strddMMyy & ".csv"
        Dim url2download As String = "https://www.bseindia.com/download/BhavCopy/Equity/EQ" & strddMMyy.ToUpper() & "_CSV.ZIP"

        Dim localfile As String = localDirectory & strddMMyy & "\BSE_HISTORICAL\eq" & strddMMyy.ToUpper() & "_CSV.ZIP"
        DownloadFile(url2download, localfile)
        UnzipFile(localfile, file2download, localDirectory & strddMMyy & "\BSE_HISTORICAL")


        File.Delete(localfile)

        Console.WriteLine("BSE File Downloaded..." & file2download)

        Dim fileName As String = Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\BSE_HISTORICAL\" & file2download
        Dim fs As New FileStream(Trim(fileName), FileMode.Open, FileAccess.Read)
        Dim sr As New StreamReader(fs)

        Dim str As String = sr.ReadLine
        ' SC_CODE SC_NAME	    SC_GROUP(2)	SC_TYPE(3)	    OPEN(4)	    HIGH-5	    LOW-6	    CLOSE-7    LAST-8	    PREVCLOSE-9 	NO_TRADES-10	NO_OF_SHRS-11	NET_TURNOV
        '540005  LTI         	A 	             Q	        5025	    5124.75	    4970	    5068	    5068	    5015.45	         6175	            33532	    169767163

        Dim i As Integer = 0

        Do Until str Is Nothing
            If i <> 0 Then
                Try
                    Dim splt() As String = str.Split(",")
                    Dim exch As String

                    Dim Open As Double
                    Dim High As Double
                    Dim Low As Double
                    Dim PrevClose As Double
                    Dim lastprice As Double
                    Dim Volume As Long
                    Dim series As String
                    Dim name As String



                    exch = splt(0).ToString().Trim()
                    name = splt(1).ToString().Trim()
                    series = splt(2).ToString().Trim()
                    Open = splt(4).ToString().Trim()
                    High = splt(5).ToString().Trim()
                    Low = splt(6).ToString().Trim()
                    lastprice = splt(7).ToString().Trim()
                    PrevClose = splt(9).ToString().Trim()
                    Volume = splt(11).ToString().Trim()




                    Dim Sql As String = " select stock_id from STOCK_ITEMS where EXCHANGE_SYMBOL ='" & exch & "'"
                    Dim dt As DataTable = New DataTable
                    Try
                        dt = SqlHelper.ExecuteDataset(My.Settings.conn_str, CommandType.Text, Sql).Tables(0)
                        If dt.Rows.Count = 0 Then

                        Else

                            Dim strSql As String = ""
                            Try


                                Sql = " select * from STOCK_EOD where stock_id ='" & dt.Rows(0)("stock_id") & "'  and trade_date ='" & Format(Dates, "MM-dd-yy") & "' "
                                Dim dt_dates As DataTable = New DataTable
                                dt_dates = SqlHelper.ExecuteDataset(My.Settings.conn_str, CommandType.Text, Sql).Tables(0)

                                If dt_dates.Rows.Count = 0 Then

                                    If (exch.Trim() <> "") Then
                                        strSql = "insert into STOCK_EOD (" &
                                            "  stock_id, series," &
                                            " last_price,day_open,day_high,day_low,PREVDAY_CLOSE," &
                                            " acc_volume,TRADE_DATE" &
                                            ") values" &
                                            "('" & dt.Rows(0)("stock_id") & "','EQ'," &
                                            " '" & lastprice & "','" & Open & "','" & High & "','" & Low & "','" & PrevClose & "'," &
                                            " '" & Volume & "','" & Format(Dates, "MM-dd-yy") & "' " &
                                            ")"
                                        SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)
                                        ' MUKESH SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)
                                        clsWrite.CaptureInsertsLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\BSE_HISTORICAL\", strSql)
                                        Console.WriteLine(strddMMyy & " " & exch)
                                        NSEcounterInsert = NSEcounterInsert + 1
                                    End If
                                Else
                                    Sql = "update STOCK_EOD set exchange_symbol ='" & exch & "',series ='EQ' where stock_id ='" & dt.Rows(0)("stock_id") & "'"
                                    SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, Sql)
                                    Console.WriteLine("Updating " & exch)
                                End If
                            Catch ex As Exception
                                clsWrite.CaptureLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\BSE_HISTORICAL\", ex.Message & " ===> " & strSql, "err")
                            End Try


                        End If

                    Catch ex As Exception
                        clsWrite.CaptureLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\BSE_HISTORICAL\", ex.Message & " ===> ", "err")
                        Console.WriteLine("XXXXXXXXXXXXXXXXX " & ex.Message)
                    End Try
                Catch ex As Exception
                    clsWrite.CaptureLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\BSE_HISTORICAL\", ex.Message & " ===> ", "err")
                    Console.WriteLine("XXXXXXXXXXXXXXXXX " & ex.Message)
                End Try
            End If
nextloop:
            str = sr.ReadLine
            i = 1
        Loop
    End Sub

    Sub NSE_EOD()

        'https://archives.nseindia.com/content/historical/EQUITIES/2021/JAN/cm19JAN2021bhav.csv.zip
        'https://archives.nseindia.com/content/historical/EQUITIES/2021/Jan/cm19Jan2021bhav.csv.zip



        Console.WriteLine(Dates)



        Dim file2download As String = "cm" & Format(Dates, "ddMMMyyyy") & "bhav.csv"
        Dim url2download As String = "https://archives.nseindia.com/content/historical/EQUITIES/" & Format(Dates, "yyyy") & "/" & Format(Dates, "MMM").ToUpper() & "/cm" & Format(Dates, "ddMMMyyyy").ToUpper() & "bhav.csv.zip"

        Dim localfile As String = localDirectory & Format(Dates, "ddMMyy") & "\NSE_HISTORICAL\cm" & Format(Dates, "ddMMMyyyy").ToUpper() & "bhav.csv.zip"
        DownloadFile(url2download, localfile)
        UnzipFile(localfile, file2download, localDirectory & Format(Dates, "ddMMyy") & "\NSE_HISTORICAL")


        File.Delete(localfile)

        Console.WriteLine("NSE File Downloaded..." & file2download)

        Dim fileName As String = Environment.CurrentDirectory & "\bhavfiles\" & Format(Dates, "ddMMyy") & "\NSE_HISTORICAL\" & file2download
        Dim fs As New FileStream(Trim(fileName), FileMode.Open, FileAccess.Read)
        Dim sr As New StreamReader(fs)

        Dim str As String = sr.ReadLine

        Dim i As Integer = 0

        Do Until str Is Nothing
            If i <> 0 Then
                Try
                    Dim splt() As String = str.Split(",")
                    Dim exch As String

                    Dim Open As Double
                    Dim High As Double
                    Dim Low As Double
                    Dim PrevClose As Double
                    Dim lastprice As Double
                    Dim Volume As Long

                    Dim Series As String



                    exch = splt(0).ToString().Trim()
                    Series = splt(1).ToString().Trim()
                    Open = splt(2).ToString().Trim()
                    High = splt(3).ToString().Trim()
                    Low = splt(4).ToString().Trim()
                    lastprice = splt(5).ToString().Trim()
                    PrevClose = splt(7).ToString().Trim()
                    Volume = splt(8).ToString().Trim()

                    ''If (Series.Trim.ToUpper = "EQ" Or Series.Trim.ToUpper = "BE") Then
                    Series = "EQ"
                    '' End If


                    Dim Sql As String = " select STOCK_ID,exchange_symbol,series from STOCK_ITEMS where EXCHANGE_SYMBOL ='" & exch & "' and series in ('INX','EQ','BE','BZ','ST','SM')"
                    Dim dt As DataTable = New DataTable
                    Try
                        dt = SqlHelper.ExecuteDataset(My.Settings.conn_str, CommandType.Text, Sql).Tables(0)
                        If dt.Rows.Count <> 0 Then

                            Try


                                Sql = " select * from STOCK_EOD where STOCK_ID ='" & dt.Rows(0)("stock_id") & "' and SERIES ='" & Series & "' and trade_date ='" & Format(Dates, "MM-dd-yy") & "' "
                                Dim dt_dates As DataTable = New DataTable
                                dt_dates = SqlHelper.ExecuteDataset(My.Settings.conn_str, CommandType.Text, Sql).Tables(0)

                                If dt_dates.Rows.Count = 0 Then
                                    Series = dt.Rows(0)("series")
                                    'If Series.Trim().ToUpper() = "EQ" Or Series.Trim().ToUpper() = "BE" Then '' STOCK_ITEMS
                                    If (exch.Trim() <> "") Then
                                        Sql = "insert into STOCK_EOD (" &
                                        "  exchange_symbol,stock_id, series," &
                                        " last_price,day_open,day_high,day_low,PREVDAY_CLOSE," &
                                        " acc_volume,TRADE_DATE" &
                                        ") values" &
                                        "('" & dt.Rows(0)("exchange_symbol") & "','" & dt.Rows(0)("stock_id") & "','" & Series & "'," &
                                        " '" & lastprice & "','" & Open & "','" & High & "','" & Low & "','" & PrevClose & "'," &
                                        " '" & Volume & "','" & Format(Dates, "MM-dd-yy") & "' " &
                                        ")"
                                        SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, Sql)
                                        ' SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)
                                        clsWrite.CaptureInsertsLogs(Environment.CurrentDirectory & "\bhavfiles\" & Format(Dates, "ddMMyy") & "\NSE_HISTORICAL\", Sql)
                                        Console.WriteLine(Format(Dates, "ddMMyy") & " " & exch)
                                        NSEcounterInsert = NSEcounterInsert + 1
                                    End If
                                    'End If
                                Else
                                    Sql = "update STOCK_EOD set exchange_symbol ='" & exch & "' where stock_id ='" & dt.Rows(0)("stock_id") & "' and series ='" & Series & "'"
                                    SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, Sql)
                                End If
                            Catch ex As Exception
                                clsWrite.CaptureLogs(Environment.CurrentDirectory & "\bhavfiles\" & Format(Dates, "ddMMyy") & "\NSE_HISTORICAL\", ex.Message & " ===> " & Sql, "err")
                            End Try

                        Else

                            'Sql = "insert into STOCK_ITEMS (" &
                            '                "  exchange_id,exchange_symbol, series," &
                            '                " last_price,day_open,day_high,day_low,PREVDAY_CLOSE," &
                            '                " acc_volume,UPDATE_DATE_TIME" &
                            '                ") values" &
                            '                "(1,'" & exch & "','" & Series & "'," &
                            '                " '" & lastprice & "','" & Open & "','" & High & "','" & Low & "','" & PrevClose & "'," &
                            '                " '" & Volume & "','" & Format(Dates, "MM-dd-yy") & "' " &
                            '                ")"
                            'SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, Sql)
                            '' SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)
                            'clsWrite.CaptureInsertsLogs(Environment.CurrentDirectory & "\bhavfiles\" & Format(Dates, "ddMMyy") & "\NSE_HISTORICAL\", Sql)
                            'Console.WriteLine(Format(Dates, "ddMMyy") & " " & exch)
                            'NSEcounterInsert = NSEcounterInsert + 1
                        End If

                    Catch ex As Exception
                        clsWrite.CaptureLogs(Environment.CurrentDirectory & "\bhavfiles\" & Format(Dates, "ddMMyy") & "\NSE_HISTORICAL\", ex.Message & " ===> ", "err")
                        Console.WriteLine("XXXXXXXXXXXXXXXXX " & ex.Message)
                    End Try
                Catch ex As Exception
                    clsWrite.CaptureLogs(Environment.CurrentDirectory & "\bhavfiles\" & Format(Dates, "ddMMyy") & "\NSE_HISTORICAL\", ex.Message & " ===> ", "err")
                    Console.WriteLine("XXXXXXXXXXXXXXXXX " & ex.Message)
                End Try
            End If
nextloop:
            str = sr.ReadLine
            i = 1
        Loop
    End Sub

    Sub downloadFO_FII_DERIVATIVES_STATS()
        ''  Console.WriteLine("IN")
        Try
            ''  Console.ReadLine()
            ''https://archives.nseindia.com/content/fo/fii_stats_12-May-2023.xls
            Dim file2download As String = "fii_stats_" & strdd & "-" & strMMM & "-" & stryyyy & ".xls"
            Dim url2download As String = "https://archives.nseindia.com/content/fo/" & file2download

            Dim localfile As String = localDirectory & strddMMyy & "\FO_FII_DERIVATIVES_STATS\" & file2download
            ''      Console.WriteLine("before DownloadFile try catch ")
            Try
                ''     Console.WriteLine("before DownloadFile ")
                Dim bool As Boolean
                bool = DownloadFile(url2download, localfile)
                Console.WriteLine("bool =" & bool)
                ''   Console.WriteLine("after DownloadFile ")
            Catch ex As Exception
                clsWrite.CaptureLogs(localDirectory & strddMMyy & "\FO_FII_DERIVATIVES_STATS\", ex.Message & " 11 ===> ", "err")
                Console.WriteLine("XXXXXXXXXXXXXXXXX " & ex.Message)
                ''     Console.ReadLine()
            End Try
            ''   Console.ReadLine()
            ''  Console.WriteLine("after DownloadFile try catch ")
            'UnzipFile(localfile, file2download, localDirectory & Format(Dates, "ddMMyy") & "\NSE_HISTORICAL")


            Try

                Dim APP As New Excel.Application

                Dim WorkBook As Excel.Workbook

                Dim WorkSheet As Excel.Worksheet

                Dim xRange As Excel.Range
                Dim strSql As String



                WorkBook = APP.Workbooks.Open(localfile)

                WorkSheet = WorkBook.Worksheets("sheet1")


                For row As Integer = 4 To 24


                    If (WorkSheet.Cells(row, 1).Value() <> "") Then
                        ''   If WorkSheet.Cells(row, 1).Value() <> "MIDCPNIFTY OPTIONS" And WorkSheet.Cells(row, 2).Value() <> 0 Then
                        If WorkSheet.Cells(row, 1).Value() <> "" Then

                            strSql = " select * from FO_FII_DERIVATIVES_STATS where contractname ='" & WorkSheet.Cells(row, 1).Value() & "' and filename ='" & file2download & "'"
                            Dim dt As DataTable = New DataTable
                            dt = SqlHelper.ExecuteDataset(My.Settings.conn_str, CommandType.Text, strSql).Tables(0)

                            If dt.Rows.Count > 0 Then
                                strSql = "update FO_FII_DERIVATIVES_STATS set buycontracts ='" & WorkSheet.Cells(row, 2).Value() & "', sellcontracts='" & WorkSheet.Cells(row, 4).Value() & "'" &
                                    ", buyamount ='" & WorkSheet.Cells(row, 3).Value() & "', sellamount='" & WorkSheet.Cells(row, 5).Value() & "',oicontracts ='" & WorkSheet.Cells(row, 6).Value() & "', xamount='" & WorkSheet.Cells(row, 7).Value() & "'" &
                                    ", updatedatetime = getdate()" &
                                    " where contractname ='" & WorkSheet.Cells(row, 1).Value() & "' and filedate='" & Format(Dates, "MM-dd-yy") & "'"
                                SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)
                            Else
                                strSql = "insert into  FO_FII_DERIVATIVES_STATS (contractname, buycontracts, sellcontracts, buyamount, sellamount, oicontracts, xamount, filedate, filename, updatedatetime) values " &
                            " ( '" & WorkSheet.Cells(row, 1).Value() & "', '" & WorkSheet.Cells(row, 2).Value() & "', '" & WorkSheet.Cells(row, 4).Value() & "'" &
                             "  ,'" & WorkSheet.Cells(row, 3).Value() & "', '" & WorkSheet.Cells(row, 5).Value() & "', '" & WorkSheet.Cells(row, 6).Value() & "', '" & WorkSheet.Cells(row, 7).Value() & "'" &
                                   ",'" & Format(Dates, "MM-dd-yyyy") & "','" & file2download & "',getdate())"

                                SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)
                            End If
                        End If
                    Else '
                        Dim str As String = "0"
                    End If
                Next


                WorkBook.Close()

                WorkSheet = Nothing

                WorkBook = Nothing
            Catch ex As Exception

                Console.WriteLine(ex.Message)
            End Try



        Catch ex As Exception

            Try
                clsWrite.CaptureLogs(localDirectory & strddMMyy & "\FO_FII_DERIVATIVES_STATS\", ex.Message & " ===> ", "err")
            Catch exx As Exception
                Console.WriteLine("ERRR " & exx.Message)
            End Try


            Console.WriteLine("ERRR " & ex.Message)

        End Try


    End Sub

    Dim url2download_BSEDelivery As String
    Sub BSEDELIVERY()



        Dim today = DateTime.Now
        Dim file2download As String = ""
        Dim strSql As String
        For del_var As Integer = 1 To 1
            Try


                Dim answer As Date = Dates ' DateTime.Today.AddDays(-del_var).ToString("D")
                If Directory.Exists(Environment.CurrentDirectory & "\bhavfiles\" & Format(answer, "ddMMyy") & "\BSEDELSTATS") = False Then
                    Directory.CreateDirectory(Environment.CurrentDirectory & "\bhavfiles\" & Format(answer, "ddMMyy") & "\BSEDELSTATS")
                End If
                Console.WriteLine("Update Main DB for " & answer.ToString() & " ")

                Dim concatdate = ""
                Dim strdate = Day(answer)
                Dim strmth = Month(answer)
                If Len(Day(answer).ToString()) = "1" Then
                    strdate = "0" & Day(answer)
                End If

                If Len(Month(answer).ToString()) = "1" Then
                    strmth = "0" & Month(answer)
                End If
                concatdate = strdate & strmth
                file2download = "SCBSEALL" & concatdate & ".txt"

                Dim localfile As String = Environment.CurrentDirectory & "\bhavfiles\" & Format(answer, "ddMMyy") & "\BSEDELSTATS\SCBSEALL" & concatdate & ".zip"
                url2download_BSEDelivery = "https://www.bseindia.com/BSEDATA/gross/" & Year(answer) & "/SCBSEALL" & concatdate & ".zip"
                DownloadFile(url2download_BSEDelivery, localfile)
                UnzipFile(localfile, file2download, Environment.CurrentDirectory & "\bhavfiles\" & Format(answer, "ddMMyy") & "\BSEDELSTATS")
                File.Delete(localfile)





                Dim fileName As String = Environment.CurrentDirectory & "\bhavfiles\" & Format(answer, "ddMMyy") & "\BSEDELSTATS\" & file2download
                Dim fs As New FileStream(Trim(fileName), FileMode.Open, FileAccess.Read)
                Dim sr As New StreamReader(fs)

                Dim str As String = sr.ReadLine

                Dim i As Integer = 0
                BSEcounterUpdate = 0
                Do Until str Is Nothing
                    If i <> 0 Then


                        Try
                            Dim splt() As String = str.Split("|")
                            Dim exch As String
                            Dim delQty As Long
                            Dim delperc As Double


                            exch = splt(1).ToString().Trim()
                            delQty = splt(2).ToString().Trim()
                            delperc = splt(6).ToString().Trim()


                            strSql = " select stock_id from STOCK_ITEMS where exchange_symbol ='" & exch.Trim() & "'"
                            Dim dt_dates As DataTable = New DataTable
                            dt_dates = SqlHelper.ExecuteDataset(My.Settings.conn_str, CommandType.Text, strSql).Tables(0)


                            If (dt_dates.Rows.Count > 0) Then
                                strSql = "update stock_eod set DEL_QTY ='" & delQty & "',DEL_PERC='" & delperc & "'" &
                                                        "  where STOCK_ID ='" & dt_dates.Rows(0)("STOCK_ID") & "' and trade_date='" & Format(Dates, "MM-dd-yy") & "'"
                                SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)
                                '  SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)

                                clsWrite.CaptureUpdateLogs(Environment.CurrentDirectory & "\bhavfiles\" & Format(Dates, "ddMMyy") & "\BSEDELSTATS\", strSql)
                                BSEcounterUpdate = BSEcounterUpdate + 1
                            End If

                            Console.WriteLine("Updating BSE DELIVERABLES ==> " & exch & " " & Format(answer, "ddMMyy"))
                        Catch ex As Exception
                            Console.WriteLine("Update Main DB" & ex.Message)
                        End Try

                    End If
nextloop:
                    str = sr.ReadLine
                    i = i + 1
                Loop
            Catch ex As Exception
                Console.WriteLine("Update Main DB" & ex.Message)
            End Try
        Next
    End Sub

    Dim url2download_NSEDELIVERY As String
    Sub NSEDELIVERY()



        Dim today = DateTime.Now
        Dim file2download As String = ""
        Dim strSql As String
        '  For k As Integer = 1 To 15
        Try
            '  Dates = Now.Date.AddDays(-k)

            If Directory.Exists(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\NSEDELSTATS") = False Then
                Directory.CreateDirectory(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\NSEDELSTATS")
            End If


            file2download = "MTO_" & Format(Dates, "ddMMyyyy") & ".dat"

            Dim localfile As String = Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\NSEDELSTATS\MTO_" & Format(Dates, "ddMMyyyy") & ".DAT"
            'https://www1.nseindia.com/archives/equities/mto/MTO_20012021.DAT
            url2download_NSEDELIVERY = "https://www1.nseindia.com/archives/equities/mto/MTO_" & Format(Dates, "ddMMyyyy") & ".DAT"


            If (DownloadFile(url2download_NSEDELIVERY, localfile) = False) Then
                GoTo nextloop
            End If

            DownloadFile(url2download_NSEDELIVERY, localfile)




            Dim fileName As String = localfile ' Environment.CurrentDirectory & "\bhavfiles\" & localfile & "\NSEDELSTATS\" & file2download
            Dim fs As New FileStream(Trim(fileName), FileMode.Open, FileAccess.Read)
            Dim sr As New StreamReader(fs)

            Dim str As String = sr.ReadLine

            Dim i As Integer = 0
            Do Until str Is Nothing
                If i <> 0 Then


                    Try
                        Dim splt() As String = str.Split(",")
                        Dim exch As String
                        Dim delQty As Long
                        Dim delperc As Double
                        Dim series As String

                        exch = splt(2).ToString().Trim()
                        series = splt(3).ToString().Trim()
                        delQty = splt(5).ToString().Trim()
                        delperc = splt(6).ToString().Trim()
                        Console.WriteLine(exch & " ==> " & Dates)

                        ''If (series.Trim.ToUpper = "EQ" Or series.Trim.ToUpper = "BE") Then
                        series = "EQ"
                        '' End If


                        strSql = " select stock_id from STOCK_ITEMS where exchange_symbol ='" & exch.Trim() & "' and SERIES ='" & series & "'"
                        Dim dt_dates As DataTable = New DataTable
                        dt_dates = SqlHelper.ExecuteDataset(My.Settings.conn_str, CommandType.Text, strSql).Tables(0)


                        If (dt_dates.Rows.Count > 0) Then
                            strSql = "update STOCK_EOD set DEL_QTY ='" & delQty & "',DEL_PERC='" & delperc & "'" &
                                                        "  where STOCK_ID ='" & dt_dates.Rows(0)("STOCK_ID") & "' and trade_date='" & Format(Dates, "MM-dd-yy") & "'"
                            SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)
                            clsWrite.CaptureUpdateLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\NSEDELSTATS\", strSql)
                            NSEcounterUpdate = NSEcounterUpdate + 1
                        End If
                        dt_dates = Nothing
                    Catch ex As Exception
                        Console.WriteLine("Update Main DB" & ex.Message)
                    Finally

                    End Try

                End If
nextloop:
                str = sr.ReadLine
                i = i + 1
            Loop
        Catch ex As Exception
            Console.WriteLine("Update Main DB" & ex.Message)
        End Try
        '  Next
    End Sub

    Sub NSE_WK52HI_LOW()
        'https://archives.nseindia.com/content/CM_52_wk_High_low_05122022.csv
        ''Dim today = DateTime.Now


        Dim file2download As String = ""
        Dim strSql As String
        '  For k As Integer = 1 To 15
        Try

            Dim DirectoryName = Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\NSE_WK52HI_LOW"
            If Directory.Exists(DirectoryName) = False Then
                Directory.CreateDirectory(DirectoryName)
            End If

            ''  file2download = "CM_52_wk_High_low_" & strddMMyyyy & ".csv"
            Dim localfile As String = DirectoryName & "\CM_52_wk_High_low_" & strddMMyyyy & ".csv"

            Dim url2download As String = "https://archives.nseindia.com/content/CM_52_wk_High_low_" & strddMMyyyy & ".csv"


            If (DownloadFile(url2download, localfile) = False) Then
                GoTo nextloop
            End If

            DownloadFile(url2download, localfile)




            Dim fileName As String = localfile ' Environment.CurrentDirectory & "\bhavfiles\" & localfile & "\NSEDELSTATS\" & file2download
            Dim fs As New FileStream(Trim(fileName), FileMode.Open, FileAccess.Read)
            Dim sr As New StreamReader(fs)

            Dim str As String = sr.ReadLine

            Dim i As Integer = 0
            Do Until str Is Nothing
                If i > 2 Then


                    Try
                        Dim splt() As String = str.Replace(""",""", ",").Split(",")
                        Dim exch As String
                        Dim Adjusted_52_Week_High As String
                        Dim str52_Week_High_Date As String
                        Dim Adjusted_52_Week_Low As String
                        Dim str52_Week_Low_Date As String
                        Dim series As String

                        exch = splt(0).ToString().Replace("""", "")
                        series = splt(1).ToString().Trim().Replace(""",""", "").Trim()
                        Adjusted_52_Week_High = splt(2).ToString().Trim()
                        str52_Week_High_Date = splt(3).ToString().Trim()
                        Adjusted_52_Week_Low = splt(4).ToString().Trim()
                        str52_Week_Low_Date = splt(5).ToString().Replace("""", "")


                        If (series.Trim.ToUpper = "EQ" Or series.Trim.ToUpper = "BE") Then
                            series = "'EQ','BE'"
                        Else

                            series = "'" & series & "'"
                        End If


                        strSql = " select stock_id,series from STOCK_ITEMS where exchange_symbol ='" & exch.Trim() & "' and SERIES in (" & series & ")"
                        Dim dt_dates As DataTable = New DataTable
                        dt_dates = SqlHelper.ExecuteDataset(My.Settings.conn_str, CommandType.Text, strSql).Tables(0)

                        If (dt_dates.Rows.Count > 0) Then
                            If (dt_dates.Rows(0)("SERIES").Trim.ToUpper = "EQ" Or dt_dates.Rows(0)("SERIES").Trim.ToUpper = "BE") Then
                                series = "'EQ','BE'"
                            Else
                                series = "'" & dt_dates.Rows(0)("SERIES").Trim.ToUpper() & "'"

                            End If


                            strSql = "update STOCK_ITEMS set WK52_HIGH ='" & Adjusted_52_Week_High & "', WK52_HIGH_DATE ='" & str52_Week_High_Date & "'," &
                                                        " WK52_LOW ='" & Adjusted_52_Week_Low & "', WK52_LOW_DATE ='" & str52_Week_Low_Date & "'" &
                                                        "  where STOCK_ID ='" & dt_dates.Rows(0)("STOCK_ID") & "' and series in (" & series & ")"
                            SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)
                            clsWrite.CaptureUpdateLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\NSE_WK52HI_LOW\", strSql)
                            NSEcounterUpdate = NSEcounterUpdate + 1
                        End If
                        dt_dates = Nothing
                    Catch ex As Exception
                        Console.WriteLine("Update Main DB" & ex.Message)
                    Finally

                    End Try

                End If
nextloop:
                str = sr.ReadLine
                i = i + 1
            Loop
        Catch ex As Exception
            Console.WriteLine("Update Main DB" & ex.Message)
        End Try
        '  Next
    End Sub

    Sub ReadNSEDb(str As String)
        Try
            Dim splt() As String = str.Split(",")
            Dim exch As String
            Dim OptionType As String
            Dim Fhi_52_wk As Double
            Dim Flo_52_wk As Double
            Dim Open As Double
            Dim High As Double
            Dim Low As Double
            Dim PrevClose As Double
            Dim lastprice As Double
            Dim Volume As Long
            Dim Name As String
            Dim Series As String
            Dim isIndex As String

            isIndex = splt(0).ToString().Trim()

            If (isIndex = "Y") Then
                exch = splt(3).ToString().Trim()

                If (exch.ToUpper() = "NIFTY 50") Then
                    exch = "S&P CNX NIFTY"
                End If
            ElseIf (isIndex = "N") Then
                exch = splt(2).ToString().Trim()
            Else
                Exit Sub
            End If

            Series = splt(1).ToString().Trim()
            Name = splt(3).ToString().Trim()
            PrevClose = splt(4).ToString().Trim()
            Open = splt(5).ToString().Trim()
            High = splt(6).ToString().Trim()
            Low = splt(7).ToString().Trim()
            lastprice = splt(8).ToString().Trim()
            Volume = splt(10).ToString().Trim()
            Fhi_52_wk = splt(13).ToString().Trim()
            Flo_52_wk = splt(14).ToString().Trim()


            If (Series.Trim.ToUpper = "EQ" Or Series.Trim.ToUpper = "BE") Then
                Series = "EQ"
            End If

            Dim counterInsert As Integer = 1
            Dim counterUpdate As Integer = 1
            Dim Sql As String = " select LAST_PRICE from STOCK_ITEMS_TRANSACTION where exchangesymbol ='" & exch & "'"
            Dim dt As DataTable = New DataTable
            Try
                dt = SqlHelper.ExecuteDataset(My.Settings.conn_str, CommandType.Text, Sql).Tables(0)
                If dt.Rows.Count > 0 Then
                    If (Convert.ToDouble(dt.Rows(0)("LAST_PRICE")).ToString("#.00") <> Convert.ToDouble(lastprice).ToString("#.00")) Then
                        Try
                            Dim strSql As String = "update STOCK_ITEMS_TRANSACTION set LAST_PRICE ='" & lastprice & "',ACC_VOLUME='" & Volume & "',  DAY_OPEN='" & Open & "',DAY_HIGH='" & High & "'" &
                                                    " ,DAY_LOW='" & Low & "', PREVDAY_CLOSE ='" & PrevClose & "',hi_52_wk ='" & Fhi_52_wk & "',lo_52_wk ='" & Flo_52_wk & "'  where exchangesymbol ='" & exch & "' and series ='" & Series & "' and exchange_id =1 and instrument_id =2"
                            SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)
                            'Debug.WriteLine(splt(2).ToString().Trim() & " " & " DB VALUE " & dt.Rows(0)("LAST_PRICE") & ". EXCHANGE VALUE " & splt(8).ToString().Trim() & " " & splt(1).ToString().Trim())
                            clsWrite.CaptureLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\NSE\", " UPDATE Count ==> " + counterUpdate.ToString() + " .." + strSql)
                            counterUpdate = counterUpdate + 1

                        Catch ex As Exception

                        End Try

                    Else

                    End If
                Else
                    'Debug.WriteLine(splt(2).ToString().Trim() & " " & splt(1).ToString().Trim())
                    Try
                        If isIndex.Trim() = "N" Then '' STOCK_ITEMS
                            If (exch.Trim() <> "") Then
                                Dim strSql As String = "insert into STOCK_ITEMS_TRANSACTION (" &
                                "  exchange_id,instrument_id,group_id,name,graphic_name, " &
                                " exchangesymbol,last_price,day_open,day_high,day_low, PREVDAY_CLOSE," &
                                " acc_volume,trade_volume,series,free_float,face_value,update_date_time," &
                                " market_type,upper_circuit,lower_circuit,open_interest," &
                                " open_int_change,open_int_close,yield,yld_netchange,isin,bbop,bboq,bsop,bsoq,lot_size,bridgesymbol,average_trade_price_fut) values" &
                                "(1,2,5,'" & Name & "','" & Name & "'," &
                                " '" & exch & "','" & lastprice & "','" & Open & "','" & High & "','" & Low & "','" & PrevClose & "'," &
                                " '" & Volume & "',0,'" & Series & "',0,0,getDate()," &
                                " 'N',0,0,0" &
                                " ,0,0,0,0,0,0,0,0,0,0,'',0)"
                                SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)

                                clsWrite.CaptureLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\NSE\", "", " INSERT Count ==> " + counterInsert.ToString() + " .." + strSql)
                                counterInsert = counterInsert + 1
                            End If
                        End If
                    Catch ex As Exception
                        clsWrite.CaptureLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\NSE\", "", ex.Message)
                    End Try


                End If

            Catch ex As Exception
                Console.WriteLine("Update Main DB" & ex.Message)
            End Try
        Catch ex As Exception

        End Try

    End Sub

#End Region

#Region "TEST BSE"
    Sub BSE()
        'http://www.bseindia.com/download/BhavCopy/Equity/EQ190417_CSV.ZIP
        Dim file2download As String = "EQ" & strddMMyy & ".csv"
        Dim url2download As String = "http://www.bseindia.com/download/BhavCopy/Equity/" & "EQ" & strddMMyy & "_CSV.ZIP"
        'http://www.bseindia.com/download/BhavCopy/Equity/EQ060417_CSV.ZIP
        Dim localfile As String = Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\BSE\" & "EQ" & strddMMyy & "_csv.zip"
        DownloadFile(url2download, localfile)
        'Exit Sub
        UnzipFile(localfile, file2download, Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\BSE\")

        File.Delete(localfile)
        Console.WriteLine("BSE File Downloaded..." & file2download)


        ''''ReadCSVFile_BSE(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\" & file2download)





        Dim fileName As String = Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\BSE\" & file2download
        Dim fs As New FileStream(Trim(fileName), FileMode.Open, FileAccess.Read)
        Dim sr As New StreamReader(fs)

        Dim str As String
        str = sr.ReadLine

        Dim i As Integer
        i = 0
        Do Until str Is Nothing

            If i <> 0 Then
                Dim splt() As String = str.Split(",")



                Dim exch As String
                ' Dim OptionType As String
                Dim Open As Double
                Dim High As Double
                Dim Low As Double
                Dim PrevClose As Double
                Dim lastprice As Double
                Dim Volume As Long
                Dim Name As String
                Dim sc_type As String
                Dim group_name As String

                exch = splt(0).ToString().Trim()
                group_name = splt(2).ToString().Trim()
                sc_type = splt(3).ToString().Trim()
                lastprice = splt(7).ToString().Trim()
                Name = splt(1).ToString().Trim()
                PrevClose = splt(9).ToString().Trim()
                Volume = splt(11).ToString().Trim()
                Open = splt(4).ToString().Trim()
                High = splt(5).ToString().Trim()
                Low = splt(6).ToString().Trim()

                If (exch = "501144") Then
                    Dim s As String
                    s = ""
                End If


                If sc_type = "Q" And group_name <> "F" And group_name <> "G" Then
                    Dim Sql As String = " select LAST_PRICE from STOCK_ITEMS where exchange_symbol ='" & exch & "'"
                    Dim dt As DataTable = New DataTable
                    Try
                        dt = SqlHelper.ExecuteDataset(My.Settings.conn_str, CommandType.Text, Sql).Tables(0)
                        Dim strSql As String = ""
                        If dt.Rows.Count > 0 Then
                            If (Convert.ToDouble(dt.Rows(0)("LAST_PRICE")).ToString("#.00") <> Convert.ToDouble(splt(7).Trim()).ToString("#.00")) Then

                                Try
                                    'LAST_PRICE='" & Trim(arr(7)) & "',DAY_OPEN='" & Trim(arr(4)) & "',DAY_HIGH='" & Trim(arr(5)) & "',DAY_LOW='" & Trim(arr(6)) & "' where EXCHANGESYMBOL='" & Trim(arr(0)) & "'"

                                    strSql = "update STOCK_ITEMS SET series ='EQ', UPDATE_DATE_TIME_BHAV=getDate(),group_name ='" & group_name & "',LAST_PRICE ='" & lastprice & "',ACC_VOLUME='" & Volume & "',DAY_OPEN='" & Open & "',DAY_HIGH='" & High & "',DAY_LOW='" & Low & "' where exchange_symbol ='" & exch & "' and exchange_id in (2,8)"

                                    SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)
                                    Console.WriteLine("Update BSE  " & exch)
                                    'Debug.WriteLine(exch & " " & " DB VALUE " & dt.Rows(0)("LAST_PRICE") & ". EXCHANGE VALUE " & lastprice & " " & exch)
                                    clsWrite.CaptureUpdateLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\BSE\", strSql)
                                    BSEcounterUpdate = BSEcounterUpdate + 1
                                Catch ex As Exception
                                    clsWrite.CaptureLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\BSE\", ex.Message & " ===> " & strSql, "err")

                                End Try

                                ' Else
                            Else
                                strSql = "update STOCK_ITEMS SET series ='EQ',group_name ='" & group_name & "' where exchange_symbol ='" & exch & "' and exchange_id in (2,8)"

                                SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)
                                Console.WriteLine("Update BSE SERIES ONLY  " & exch)



                            End If

                        Else
                            'Debug.WriteLine(splt(2).ToString().Trim() & " " & splt(1).ToString().Trim())
                            Try
                                'If splt(0).ToString().Trim() = "N" Then '' STOCK_ITEMS

                                Name = Name.Replace("'", "''")
                                If (splt(0).ToString().Trim() <> "") Then
                                    strSql = "insert into STOCK_ITEMS (" &
                            "  exchange_id,name,ABBRS, ABBRL, " &
                            " exchange_symbol, last_price, PREVDAY_CLOSE, day_open, day_high, day_low," &
                            " acc_volume,trade_volume,face_value,update_date_time," &
                            " upper_circuit,lower_circuit,UPDATE_DATE_TIME_BHAV" &
                            " ) values" &
                            "(2,'" & Name & "','" & Name & "','" & Name & "'," &
                            " '" & exch & "','" & lastprice & "','" & PrevClose & "','" & Open & "','" & High & "','" & Low & "'," &
                            " '" & Volume & "',0,0,getDate()," &
                            " 0,0,getdate()" &
                            ")"
                                    SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)
                                    Console.WriteLine("Insert BSE  " & exch)
                                    clsWrite.CaptureInsertsLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\BSE\", strSql)
                                    BSEcounterInsert = BSEcounterInsert + 1
                                End If
                                ' End If
                            Catch ex As Exception
                                clsWrite.CaptureLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\BSE\", ex.Message & " ===> " & strSql, "err")
                            End Try


                        End If

                    Catch ex As Exception
                        Console.WriteLine("Update Main DB" & ex.Message)
                    End Try
                Else
                    Console.WriteLine(str)
                End If


            End If
            str = sr.ReadLine
                i = 1

        Loop

    End Sub
    Sub DeleteOldLogFilesFile()
        Dim files As String() = Directory.GetFiles(Directory.GetCurrentDirectory & "\bhavfiles\")

        For Each file As String In files
            Dim fi As FileInfo = New FileInfo(file)
            If fi.LastAccessTime < DateTime.Now.AddDays(-5) Then fi.Delete()
        Next


    End Sub

    '//This will delete files with specified pattern which are older than specified days
    Public Function DeleteOldFiles(ByVal sPath As String, Optional ByVal sPattern As String = "*.*", Optional ByVal OlderThanDays As Integer = 5) As Integer



        Try
            Dim dtCreated As DateTime
            Dim dtToday As DateTime = Today.Date
            Dim diObj As DirectoryInfo
            Dim ts As TimeSpan
            Dim lstDirsToDelete As New List(Of String)

            For Each sSubDir As String In Directory.GetDirectories(sPath)
                diObj = New DirectoryInfo(sSubDir)
                dtCreated = diObj.CreationTime

                ts = dtToday - dtCreated

                'Add whatever storing you want here for all folders...

                If ts.Days > OlderThanDays Then
                    lstDirsToDelete.Add(sSubDir)
                    'Store whatever values you want here... like how old the folder is
                    diObj.Delete(True) 'True for recursive deleting
                End If
            Next
        Catch ex As Exception

        End Try
    End Function

    Sub BSE_ISIN()

        '' http://www.bseindia.com/download/BhavCopy/Equity/EQ_ISINCODE_030417.zip
        Dim file2download As String = "EQ_ISINCODE_" & strddMMyy & ".csv"
        '' file2download = "EQ_ISINCODE_301122.csv"
        Dim url2download As String = "http://www.bseindia.com/download/bhavcopy/equity/" & "EQ_ISINCODE_" & strddMMyy & ".zip"
        '  url2download = "http://www.bseindia.com/download/bhavcopy/equity/EQ_ISINCODE_301122.zip"
        Dim localfile As String = Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\BSE\" & "EQ_ISINCODE_" & strddMMyy & "_csv.zip"
        ''  localfile = "D:\test\Download_BhavFiles\Download_BhavFiles\bin\Debug\bhavfiles\301122\BSE\EQ_ISINCODE_301122_csv.zip"
        DownloadFile(url2download, localfile)
        'Exit Sub
        UnzipFile(localfile, file2download, Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\BSE\")

        File.Delete(localfile)
        Console.WriteLine("BSE File Downloaded..." & file2download)



        Dim fileName As String = Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\BSE\" & file2download

        Dim fs As New FileStream(Trim(fileName), FileMode.Open, FileAccess.Read)
        Dim sr As New StreamReader(fs)

        Dim str As String
        str = sr.ReadLine

        Dim i As Integer
        i = 0
        Do Until str Is Nothing

            If i <> 0 Then
                Dim splt() As String = str.Split(",")



                Dim exch As String
                Dim ISIN As String
                Dim fv As String
                Dim Open As Double
                Dim High As Double
                Dim Low As Double
                Dim PrevClose As Double
                Dim lastprice As Double
                Dim Volume As Long
                Dim Name As String
                Dim series As String


                exch = splt(0).ToString().Trim()
                series = splt(2).ToString().Trim()
                lastprice = splt(7).ToString().Trim()
                Name = splt(1).ToString().Trim().Replace("'", "''")
                PrevClose = splt(9).ToString().Trim()
                Volume = splt(11).ToString().Trim()
                Open = splt(4).ToString().Trim()
                High = splt(5).ToString().Trim()
                Low = splt(6).ToString().Trim()
                exch = splt(0).ToString().Trim()
                ISIN = splt(14).ToString().Trim()

                ' series = "EQ"
                Dim Sql As String = " select LAST_PRICE from STOCK_ITEMS where exchange_symbol ='" & exch & "'"
                Dim dt As DataTable = New DataTable
                Try
                    dt = SqlHelper.ExecuteDataset(My.Settings.conn_str, CommandType.Text, Sql).Tables(0)
                    Dim strSql As String = ""
                    ''' If dt.Rows.Count > 0 Then
                    ' If (Convert.ToDouble(dt.Rows(0)("LAST_PRICE")).ToString("#.00") <> Convert.ToDouble(splt(7).Trim()).ToString("#.00")) Then

                    Try
                                'LAST_PRICE='" & Trim(arr(7)) & "',DAY_OPEN='" & Trim(arr(4)) & "',DAY_HIGH='" & Trim(arr(5)) & "',DAY_LOW='" & Trim(arr(6)) & "' where EXCHANGESYMBOL='" & Trim(arr(0)) & "'"

                                ''MUKESH strSql = "update STOCK_ITEMS SET ISIN ='" & ISIN & "', LAST_PRICE ='" & lastprice & "',ACC_VOLUME='" & Volume & "',DAY_OPEN='" & Open & "',DAY_HIGH='" & High & "',DAY_LOW='" & Low & "' where exchange_symbol ='" & exch & "'and exchange_id =2"
                                strSql = "update STOCK_ITEMS SET series ='EQ',group_name ='" & series & "', ISIN ='" & ISIN & "' where exchange_symbol ='" & exch & "'and exchange_id =2"

                                SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)
                                ' SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)
                                'Debug.WriteLine(exch & " " & " DB VALUE " & dt.Rows(0)("LAST_PRICE") & ". EXCHANGE VALUE " & lastprice & " " & exch)
                                clsWrite.CaptureUpdateLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\BSE\", strSql)
                                BSEcounterUpdate = BSEcounterUpdate + 1
                            Catch ex As Exception
                                clsWrite.CaptureLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\BSE\", ex.Message & " ===> " & strSql, "err")

                            End Try

                    '  Else

                    'End If

                    '''Else
                    'Debug.WriteLine(splt(2).ToString().Trim() & " " & splt(1).ToString().Trim())
                    '''Try
                    '''    'If splt(0).ToString().Trim() = "N" Then '' STOCK_ITEMS


                    '''    If (splt(0).ToString().Trim() <> "") Then
                    '''        strSql = "insert into STOCK_ITEMS (" &
                    '''    "  exchange_id,series,group_name,name,abbrs,abbrl, " &
                    '''    " isin,exchange_symbol, last_price, PREVDAY_CLOSE, day_open, day_high, day_low," &
                    '''    " acc_volume,trade_volume,face_value,update_date_time" &
                    '''    " ) values" &
                    '''    "(2,'EQ','" & series & "','" & Name & "','" & Name & "','" & Name & "'," &
                    '''    " '" & ISIN & "','" & exch & "','" & lastprice & "','" & PrevClose & "','" & Open & "','" & High & "','" & Low & "'," &
                    '''    " '" & Volume & "',0,0,getDate()" &
                    '''    ")"
                    '''        SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)
                    '''        '''MUKESH SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)

                    '''        clsWrite.CaptureInsertsLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\BSE\", strSql)
                    '''        BSEcounterInsert = BSEcounterInsert + 1
                    '''    End If
                    '''    ' End If
                    '''Catch ex As Exception
                    '''    clsWrite.CaptureLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\BSE\", ex.Message & " ===> " & strSql, "err")
                    '''End Try


                    ''' End If

                Catch ex As Exception
                    Console.WriteLine("Update Main DB" & ex.Message)
                End Try











                'Try
                '    Dim strSql As String = "update STOCK_ITEMS SET   ISIN ='" & ISIN & "' where exchange_symbol ='" & exch & "' and exchange_id =2 and instrument_id =2"

                '    SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)

                '    clsWrite.CaptureUpdateLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\BSE\", strSql)
                '    BSEcounterUpdate = BSEcounterUpdate + 1
                'Catch ex As Exception

                'End Try





            End If
            str = sr.ReadLine
            i = 1
        Loop

    End Sub


    Sub NSE_ISIN_FV()
        Dim url2download As String = "https://www1.nseindia.com/content/equities/EQUITY_L.csv"
        If (DownloadFile(url2download, localDirectory & strddMMyy & "\NSE_ISIN_FV\" & strddMMyy & ".csv")) Then
            ''SYMBOL        NAME OF COMPANY	 SERIES	 DATE OF LISTING	 PAIDUPVALUE	 MARKET LOT	 ISIN NUMBER	    FACE VALUE
            ''20MICRONS	    20 Microns Limited	EQ	    06-Oct-08	        5	            1	         INE144J01027	5


            Dim fileName As String = Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\NSE_ISIN_FV\" & strddMMyy & ".csv"
            Dim fs As New FileStream(Trim(fileName), FileMode.Open, FileAccess.Read)
            Dim sr As New StreamReader(fs)
            Dim exch As String
            Dim ISIN As String
            Dim fv As String
            Dim series As String
            Dim name As String
            Dim str As String
            str = sr.ReadLine

            Dim i As Integer
            i = 0
            Do Until str Is Nothing

                If i <> 0 Then
                    Dim splt() As String = str.Split(",")


                    exch = splt(0).ToString().Trim()
                    name = splt(1).ToString().Trim().Replace("'", "''")
                    series = splt(2).ToString().Trim()
                    ISIN = splt(6).ToString().Trim()
                    fv = splt(7).ToString().Trim()

                    Dim Sql As String
                    If (series.Trim.ToUpper = "EQ" Or series.Trim.ToUpper = "BE") Then
                        series = "EQ"
                        Sql = "select * from STOCK_ITEMS where exchange_symbol ='" & exch & "' and series in ('EQ','BE')"
                    Else
                        Sql = "select * from STOCK_ITEMS where exchange_symbol ='" & exch & "' and series ='" & series.Trim() & "'"

                    End If

                    Dim strSql As String = ""


                    Dim dt As DataTable = New DataTable

                    dt = SqlHelper.ExecuteDataset(My.Settings.conn_str, CommandType.Text, Sql).Tables(0)

                    If (dt.Rows.Count > 0) Then
                        Try
                            If (series.Trim.ToUpper = "EQ" Or series.Trim.ToUpper = "BE") Then
                                strSql = "update STOCK_ITEMS SET SERIES ='EQ',group_name ='" & series & "',ISIN ='" & ISIN & "',name ='" & name & "',abbrs ='" & name & "',abbrl ='" & name & "',face_value='" & fv & "' where exchange_symbol ='" & exch & "' and exchange_id =1 and series in ('EQ','BE') "
                            Else
                                strSql = "update STOCK_ITEMS SET SERIES ='" & series & "', group_name ='" & series & "', ISIN ='" & ISIN & "',name ='" & name & "',abbrs ='" & name & "',abbrl ='" & name & "',face_value='" & fv & "' where exchange_symbol ='" & exch & "' and exchange_id =1 and series = '" & series & "'"
                            End If
                            SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)
                            '' MUKESH SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)
                            clsWrite.CaptureUpdateLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\NSE_ISIN_FV\", strSql)
                            NSEcounterUpdate = NSEcounterUpdate + 1
                        Catch ex As Exception
                            clsWrite.CaptureLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\NSE_ISIN_FV\", ex.Message & " ===> " & strSql, "err")

                        End Try
                    Else
                        Try

                            If (series.Trim.ToUpper = "EQ" Or series.Trim.ToUpper = "BE") Then
                                series = "EQ"
                            End If
                            If (splt(0).ToString().Trim() <> "") Then
                                strSql = "insert into STOCK_ITEMS (" &
                            "  exchange_id,series,group_name,name,abbrs,abbrl, " &
                            " isin,exchange_symbol," &
                            " acc_volume,trade_volume,face_value,update_date_time" &
                            " ) values" &
                            "(1,'" & series & "','" & series & "','" & name & "','" & name & "','" & name & "'," &
                            " '" & ISIN & "','" & exch & "'," &
                            " 0,0,'" & fv & "',getDate()" &
                            ")"
                                SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)
                                '''MUKESH SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)

                                clsWrite.CaptureInsertsLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\NSE_ISIN_FV\", strSql)
                                NSEcounterInsert = NSEcounterInsert + 1
                            End If
                            ' End If
                        Catch ex As Exception
                            clsWrite.CaptureLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\NSE_ISIN_FV\", ex.Message & " ===> " & strSql, "err")
                        End Try

                    End If

                End If


                str = sr.ReadLine
                i = 1
            Loop
        End If

    End Sub





#End Region

#Region "Test FNO"


    Function ExcelToDataTable(strFilename As String) As DataTable



        Dim data As String = My.Computer.FileSystem.ReadAllText(strFilename)
        Dim dt As New DataTable
        Dim dtFilter As String = "instrument = 'OPTIDX' and SYMBOL ='NIFTY'"
        Dim FinalDt As New DataTable
        Dim dv As DataView
        Using sr As New StringReader(data)
            ' The true indicates it has header values which can be used to access fields by their name, switch to
            ' false if the CSV doesn't have them
            Using csv As New LumenWorks.Framework.IO.Csv.CsvReader(sr, True)
                dt.Load(csv)

                '  FinalDt = dt.Select(dtFilter, "trade_dates desc").CopyToDataTable()
                dv = New DataView(dt)
                dv.RowFilter = dtFilter

            End Using

            sr.Close()
        End Using

        Return dv.ToTable()



    End Function


    Function ExcelToDataTable_NoFilter(strFilename As String) As DataTable



        Dim data As String = My.Computer.FileSystem.ReadAllText(strFilename)
        Dim dt As New DataTable
        Dim dtFilter As String = "instrument = 'OPTIDX' and SYMBOL ='NIFTY'"
        Dim FinalDt As New DataTable
        Dim dv As DataView
        Using sr As New StringReader(data)
            ' The true indicates it has header values which can be used to access fields by their name, switch to
            ' false if the CSV doesn't have them
            Using csv As New LumenWorks.Framework.IO.Csv.CsvReader(sr, True)
                dt.Load(csv)

                '  FinalDt = dt.Select(dtFilter, "trade_dates desc").CopyToDataTable()
                dv = New DataView(dt)
                'dv.RowFilter = dtFilter

            End Using

            sr.Close()
        End Using

        Return dv.ToTable()



    End Function


    Dim FnoDataTable
    Sub FNO()

        Dim file2download As String = "fo" & Format(Dates, "dd") & UCase(Format(Dates, "MMM")) & Now.Year & "bhav.csv"
        Dim url2download As String = "https://archives.nseindia.com/content/historical/DERIVATIVES/" & stryyyy & "/" & strMMM.ToUpper() & "/fo" & strddMMyyyy.ToUpper() & "bhav.csv.zip"
        'url2download = "https://archives.nseindia.com/content/historical/DERIVATIVES/2023/MAY/fo12MAY2023bhav.csv.zip"

        Dim localfile As String = Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\FNO\" & file2download & ".zip"
        'localfile = "C:\BhavFiles\291015\fo29OCT2015bhav.csv.zip"
        ''https://archives.nseindia.com/content/historical/DERIVATIVES/2023/MAY/fo15MAY2023bhav.csv.zip
        DownloadFile(url2download, localfile)
        UnzipFile(localfile, file2download, Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\FNO\")

        File.Delete(localfile)
        Console.WriteLine("FNO File Downloaded..." & file2download)


        Dim fileName As String = Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\FNO\" & file2download
        FnoDataTable = ExcelToDataTable(fileName)
        Dim fs As New FileStream(Trim(fileName), FileMode.Open, FileAccess.Read)
        Dim sr As New StreamReader(fs)


        Dt_Fno_Mapping = New DataTable()


        'Dim Dt_Fno_ExpryDates As New DataTable()
        'Dt_Fno_ExpryDates = New DataTable()
        Dim Sql As String
        'Sql = = "SELECT DBEXCH,BHAVEXCH FROM BHAV_EXCH_MAPPING where exchange_id = 1 and instrument_id =3"
        '  Dt_Fno_Mapping = SqlHelper.ExecuteDataset(My.Settings.conn_str, CommandType.Text, Sql).Tables(0)


        'Sql = "select month_expiry_date,future_symbol_suffix  from future_expiry_date"
        'Dt_Fno_ExpryDates = SqlHelper.ExecuteDataset(My.Settings.conn_str, CommandType.Text, Sql).Tables(0)



        Sql = "update OPTION_ITEMS set IS_TRADED = 0"
        SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, Sql)
        ' SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, Sql)
        Dim str As String
        str = sr.ReadLine

        'Loop till the last line...
        Dim i As Integer
        i = 0
        Try
            Do Until str Is Nothing

                If i <> 0 Then

                    ReadFNODb(str, Dt_Fno_Mapping, file2download)

                End If
                str = sr.ReadLine
                i = 1


                'strSql = "update APPS_STATUS set update_date_time = getda where app_name ='BHAV_FNO' and )"
            Loop

        Catch ex As Exception

        End Try

    End Sub



    Sub ReadFNODb(str As String, Dt_Fno_Mapping As DataTable, file2download As String)
        Dim splt() As String = str.Split(",")
        Dim expdate As Date
        Dim strike_Price As Double
        Dim exch As String
        Dim OptionType As String
        Dim Open As Double
        Dim High As Double
        Dim Low As Double
        Dim PrevClose As Double
        Dim lastprice As Double
        '' Dim Volume As Long
        Dim OI As Long
        Dim PrevOI As Long
        Dim PrevOIChange As Long

        Dim counterFUTInsert As Integer = 1
        Dim counterFUTUpdate As Integer = 1
        Dim counterOPTInsert As Integer = 1
        Dim counterOPTUpdate As Integer = 1
        Dim acc_vol_incontracts As Long

        expdate = splt(2)
        exch = splt(1).ToString().Trim()
        strike_Price = splt(3).Trim()
        OptionType = splt(4).Trim()
        Open = splt(5).Trim()
        High = splt(6).Trim()
        Low = splt(7).Trim()
        High = splt(6).Trim()
        PrevClose = splt(8).Trim()
        lastprice = splt(9).Trim()
        acc_vol_incontracts = splt(10).Trim()
        OI = splt(12).Trim()
        PrevOIChange = splt(13).Trim()

        '''  Not SURE WICH VALUE TO TAKE FROM exCEL
      ''  Volume = splt(10).Trim()
        '''''
        PrevOI = OI - PrevOIChange
        'Sql = "IF EXISTS(SELECT NULL FROM options_transaction_data where month(expiry_date)='" & Month(Trim(arr(2))) & "' and YEAR(expiry_date)='" & Year(Trim(arr(2))) & "' and ExchangeSymbol='" & Trim(arr(1)) & "' and strike_price='" & Trim(arr(3)) & "'  and option_type='" & Trim(arr(4)) & "') BEGIN "
        'Sql &= " Update options_transaction_data Set open_interest='" & Trim(arr(12)) & "',lastprice='" & Trim(arr(9)) & "',Settlement_price='" & Trim(arr(9)) & "' where month(expiry_date)='" & Month(Trim(arr(2))) & "' and YEAR(expiry_date)='" & Year(Trim(arr(2))) & "' and ExchangeSymbol='" & Trim(arr(1)) & "' and strike_price='" & Trim(arr(3)) & "'  and option_type='" & Trim(arr(4)) & "'"
        'Sql &= " END ELSE BEGIN "
        'Sql &= " INSERT INTO OPTIONS_TRANSACTION_DATA (EXCHANGESYMBOL,EXPIRY_DATE,strike_price,option_type,lastprice,settlement_price,open_interest,prevday_open_interest,market_type)"
        'Sql &= " VALUES ('" & Trim(arr(1)) & "','" & Trim(arr(2)) & "','" & Trim(arr(3)) & "','" & Trim(arr(4)) & "','" & Trim(arr(8)) & "','" & Trim(arr(9)) & "','" & Trim(arr(12)) & "','" & Trim(arr(12)) & "','N') END"


        'Dim drFNOExpiryDates As DataRow()
        'drFNOExpiryDates = Dt_Fno_ExpryDates.Select("month_expiry_date = '" & expdate & "'")

        'Dim exchExtension As String
        'If (drFNOExpiryDates.Length > 0) Then
        '    exchExtension = drFNOExpiryDates(0)(1).ToString()
        'End If
        Dim stock_items_exch As String
        If (exch = "BANKNIFTY") Then
            stock_items_exch = "Nifty Bank"
        ElseIf (exch = "FINNIFTY") Then
            stock_items_exch = "Nifty Fin Service"
        ElseIf (exch = "MIDCPNIFTY") Then
            stock_items_exch = "Nifty Midcap 50"
        ElseIf (exch = "NIFTY") Then
            stock_items_exch = "Nifty 50"
        Else
            stock_items_exch = exch
        End If
        'Nifty Bank	BANKNIFTY
        'Nifty Fin Service	FINNIFTY
        'Nifty Midcap 50	MIDCPNIFTY
        'Nifty 50	NIFTY




        If (splt(0).ToString().Trim() = "FUTIDX" Or splt(0).ToString().Trim() = "FUTSTK" Or splt(0).ToString().Trim() = "FUTIVX") Then ''' FUTURES

            'If (splt(0).ToString().Trim() = "FUTIDX") Then  ''INDEX FUTURES
            '    Dim drResult As DataRow()
            '    drResult = Dt_Fno_Mapping.Select("bhavexch = '" & exch & "'")

            '    If (drResult.Length > 0) Then
            '        exch = drResult(0)(0).ToString()
            '    End If
            'End If

            'If (exchExtension <> "") Then
            '    '     exch = exch & exchExtension
            'End If

            Dim stock_id As String
            Dim Sql As String
            Dim strSql As String
            'If splt(0).ToString().Trim() = "FUTSTK") Then ''' FUT STOCKS
            '    Sql = " select stock_id, LAST_PRICE,LOT_SIZE from stock_items where  exchange_id in( 1,8) and  exchange_symbol ='" & stock_items_exch.Trim() & "'"
            'Else
            '    Sql = " select stock_id, LAST_PRICE,LOT_SIZE from stock_items where  exchange_id = 8 and  exchange_symbol ='" & stock_items_exch.Trim() & "'"
            'End If

            Sql = " select stock_id, LAST_PRICE,LOT_SIZE from stock_items where  exchange_id  in(1,8) and  exchange_symbol ='" & stock_items_exch.Trim() & "'"

            Dim dt As DataTable = New DataTable
            Dim dt_STOCK_ITEMS As DataTable = New DataTable
            Try
                dt_STOCK_ITEMS = SqlHelper.ExecuteDataset(My.Settings.conn_str, CommandType.Text, Sql).Tables(0)


                Sql = " select * from stock_items where  exchange_id in (3,4,5) and  exchange_symbol ='" & exch.Trim() & "' and expiry_date = '" & expdate.ToString("MM/dd/yyyy").Trim() & "'"
                dt = SqlHelper.ExecuteDataset(My.Settings.conn_str, CommandType.Text, Sql).Tables(0)
                If dt.Rows.Count > 0 Then

                    If (Convert.ToDouble(dt.Rows(0)("LAST_PRICE")).ToString("#.00") <> Convert.ToDouble(lastprice).ToString("#.00")) Then
                        Try
                            Dim acc_vol As Long
                            ' acc_vol = acc_vol_incontracts * dt_STOCK_ITEMS.Rows(0)("lot_size")
                            If (dt_STOCK_ITEMS.Rows.Count > 0) Then
                                acc_vol = acc_vol_incontracts * dt.Rows(0)("lot_size")
                                stock_id = dt_STOCK_ITEMS.Rows(0)("stock_id")
                            Else
                                acc_vol = acc_vol_incontracts
                                stock_id = "-1"
                            End If
                            ' prevday_close ='" & PrevClose & "', commented as bhav copy doesnty give prevday close
                            strSql = "update stock_items set LAST_PRICE ='" & lastprice & "', PREVDAY_OPEN_INTEREST = '" & PrevOI & "' ,acc_volume ='" & acc_vol & "', OPEn_INTEREST ='" & OI & "', DAY_OPEN='" & Open & "',DAY_HIGH='" & High & "',DAY_LOW='" & Low & "',UPDATE_DATE_TIME_BHAV=getDate()" &
                        "   where exchange_symbol ='" & exch & "' and exchange_id in (3,4,5) and EXPIRY_DATE ='" & expdate.ToString("MM/dd/yyyy").Trim() & "'"
                            SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)



                            Debug.WriteLine("UPDATE FUT " & exch & " " & splt(2).ToString().Trim() & " " & " DB VALUE " & dt_STOCK_ITEMS.Rows(0)("LAST_PRICE") & ". EXCHANGE VALUE " & lastprice & " " & expdate)
                            clsWrite.CaptureLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\FNO\", " FNO UPDATE Count ==> " + counterFUTInsert.ToString() + " .." + strSql, "info")
                            counterFUTInsert = counterFUTInsert + 1
                        Catch ex As Exception
                            clsWrite.CaptureLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\FNO\", " " + strSql, "err")
                        End Try
                    End If

                Else

                    'If (exch.Trim = "NIFTY BANK") Then


                    ' Sql = " select stock_id from STOCK_ITEMS where exchange_symbol ='" & stock_items_exch.Trim & "' and exchange_id = 1"
                    ' Dim dt_STOCK_ITEMS As DataTable = New DataTable
                    Try
                        'dt_STOCK_ITEMS = SqlHelper.ExecuteDataset(My.Settings.conn_str, CommandType.Text, Sql).Tables(0)

                        Dim acc_vol As Long

                        If (dt_STOCK_ITEMS.Rows.Count > 0) Then
                            acc_vol = acc_vol_incontracts * dt_STOCK_ITEMS.Rows(0)("lot_size")
                            stock_id = dt_STOCK_ITEMS.Rows(0)("stock_id")
                        Else
                            acc_vol = acc_vol_incontracts
                            stock_id = "-1"
                        End If

                        strSql = "insert into stock_items(series,PREVDAY_OPEN_INTEREST, acc_volume, LAST_PRICE,prevday_close, OPEn_INTEREST,DAY_OPEN, DAY_HIGH, day_low, exchange_symbol, exchange_id,EXPIRY_DATE, primary_stock_id, UPDATE_DATE_TIME_BHAV )" &
                                          " values ('FUT','" & PrevOI & "', '" & acc_vol & "', '" & lastprice & "','" & PrevClose & "', '" & OI & "',  '" & Open & "',  '" & High & "',  '" & Low & "',  '" & exch & "', 3,'" & expdate.ToString("MM/dd/yyyy").Trim() & "','" & stock_id & "',getdate() )"
                        SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)
                        Debug.WriteLine("INSERT FUT " & exch & " " & splt(2).ToString().Trim() & " " & " DB VALUE " & dt_STOCK_ITEMS.Rows(0)("LAST_PRICE") & ". EXCHANGE VALUE " & lastprice & " " & expdate)
                    Catch ex As Exception
                        clsWrite.CaptureLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\FNO\", " FUT INSERT ERR ==> " + strSql, "err")
                    End Try



                    'End If

                End If

            Catch ex As Exception
                Console.WriteLine("Update Main DB" & ex.Message)
            End Try
        ElseIf (splt(0).ToString().Trim() = "OPTIDX" Or splt(0).ToString().Trim() = "OPTSTK") Then ''' OPTIONS
            ' Dim instrument_id As String
            If (splt(0).ToString().Trim() = "OPTIDX") Then ' 'INDEX OPTIONS

                If (exch = "BANKNIFTY") Then
                    stock_items_exch = "Nifty Bank"
                ElseIf (exch = "FINNIFTY") Then
                    stock_items_exch = "Nifty Fin Service"
                ElseIf (exch = "MIDCPNIFTY") Then
                    stock_items_exch = "Nifty Midcap 50"
                ElseIf (exch = "NIFTY") Then
                    stock_items_exch = "Nifty 50"
                Else
                    stock_items_exch = exch
                End If



                If (exch = "NIFTY") And strike_Price = "19500" Then
                    Dim stssr As String = ""
                End If


                '  instrument_id = "1"

                'Dim drResult As DataRow()
                'drResult = Dt_Fno_Mapping.Select("bhavexch = '" & exch & "'")

                'If (drResult.Length > 0) Then

                '    exch = drResult(0)(0).ToString()
                'End If
            Else
                '  instrument_id = "2"
            End If

            If ((exch.ToUpper().Trim() = "NIFTY") Or (exch.ToUpper().Trim() = "NIFTY 50")) Then

                If (OptionType).Trim.ToUpper() = "PE" And strike_Price = "10500" Then
                    Dim s As String = ""
                End If
            End If
            Dim types As String = ""

            If (OptionType).Trim.ToUpper() = "CE" Then
                types = "_C"

                If exch = "NIFTY" Or exch = "NIFTY 50" Then
                    NIFTY_TOTALCALL += OI
                End If
            ElseIf (OptionType).Trim.ToUpper() = "PE" Then
                types = "_p"

                If exch = "NIFTY" Or exch = "NIFTY 50" Then
                    NIFTY_TOTALPUT += OI
                End If
            End If
            types = ""


            If strike_Price = "24000" And exch = "NIFTY" And expdate.ToString("MM/dd/yyyy").Trim() = "06/27/2024" Then
                Dim strddd As String = ""
            End If

            If strike_Price = "24000" And exch = "NIFTY" Then
                Dim strddd As String = ""
            End If

            ' Dim fn_id As Integer
            ' Dim isfuturesFound As Boolean = False
            Dim lotSize As Double
            'Dim isStockFound As Boolean = False
            Dim acc_vol As Long

            Dim Sql As String = " select stock_id,lot_size from stock_items where exchange_id ='3' and  exchange_symbol ='" & exch & "'"
            Dim dt_fno As DataTable = New DataTable
            Try
                dt_fno = SqlHelper.ExecuteDataset(My.Settings.conn_str, CommandType.Text, Sql).Tables(0)

                If dt_fno.Rows.Count > 0 Then

                    lotSize = dt_fno.Rows(0)("lot_size")
                    'fn_id = dt_fno.Rows(0)("id")
                    'isfuturesFound = True
                    'Else
                    '    Try
                    '        If exch = "NIFTY BANK" Or exch = "NIFTY 50" Then
                    '            Sql = "select top 1 id,lot_size from stock_items where  exchange_symbol ='" & exch & "'  order by expiry_date"
                    '            Dim dt_fno_bankNifty As DataTable = New DataTable
                    '            dt_fno_bankNifty = SqlHelper.ExecuteDataset(My.Settings.conn_str, CommandType.Text, Sql).Tables(0)
                    '            lotSize = dt_fno_bankNifty.Rows(0)("lot_size")
                    '            fn_id = dt_fno_bankNifty.Rows(0)("id") ''''''''' JHOL..Added FNO id for 1st expiry date for BANK NIFTY as there is no correcponding entry for BANK NIFTY (weekly expiry entry)
                    '            isfuturesFound = True
                    '        Else
                    '            isfuturesFound = False
                    '        End If
                    '    Catch ex As Exception
                    '        isfuturesFound = 0
                    '    End Try

                Else
                    lotSize = 0

                End If
            Catch
                ''    isfuturesFound = False
            End Try

            ''Sql = " select stock_id from STOCK_ITEMS where exchange_symbol ='" & stock_items_exch.Trim & "' and exchange_id = 1"
            ''Dim dt_STOCK_ITEMS As DataTable = New DataTable
            ''Try
            ''    dt_STOCK_ITEMS = SqlHelper.ExecuteDataset(My.Settings.conn_str, CommandType.Text, Sql).Tables(0)
            ''    If dt_STOCK_ITEMS.Rows.Count > 0 Then
            ''        isStockFound = True
            ''    Else
            ''        isStockFound = False
            ''    End If

            ''Catch ex As Exception
            ''    isStockFound = False
            ''End Try



            Sql = " select * from option_items where option_type ='" & OptionType & "' and  exchange_symbol ='" & exch.Trim & "' And EXPIRY_DATE ='" & expdate.ToString("MM/dd/yyyy").Trim() & "' and strike_price ='" & strike_Price & "'"
            Dim dt_options As DataTable = New DataTable
            Try
                dt_options = SqlHelper.ExecuteDataset(My.Settings.conn_str, CommandType.Text, Sql).Tables(0)
                acc_vol = acc_vol_incontracts * lotSize


                If dt_options.Rows.Count > 0 Then
                    If (Convert.ToDouble(dt_options.Rows(0)("acc_volume" & types)) <> Convert.ToDouble(acc_vol)) Then
                        Dim strSql As String
                        Try

                            acc_vol = acc_vol_incontracts * lotSize
                            strSql = "update OPTION_ITEMS set IS_TRADED=1, PREVDAY_OPEN_INTEREST" & types & "='" & PrevOI & "',  PREVDAY_CLOSE" & types & " = '" & PrevClose & "', acc_volume" & types & " = '" & acc_vol & "', LAST_PRICE" & types & " ='" & lastprice & "',  Open_Interest" & types & "='" & OI & "', filedate ='" & file2download & "' , UPDATE_DATE_TIME_BHAV=getDate()" &
                            "   where exchange_symbol ='" & exch & "'  and strike_price ='" & strike_Price & "' and EXPIRY_DATE ='" & expdate.ToString("MM/dd/yyyy").Trim() & "' and option_type ='" & OptionType & "' "
                            Console.WriteLine("Update OPT  " & exch & " " & OptionType & " " & expdate.ToString("MM/dd/yyyy").Trim() & " " & strike_Price)
                            SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)
                            ' SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)
                            'Debug.WriteLine(splt(2).ToString().Trim() & " " & " DB VALUE " & dt_STOCK_ITEMS.Rows(0)("LAST_PRICE" & types) & ". EXCHANGE VALUE " & lastprice & types & " " & expdate)
                            clsWrite.CaptureUpdateLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\FNO\", " OPT UPDATE Count ==> " + counterOPTUpdate.ToString() + " .." + strSql)

                            counterOPTUpdate = counterOPTUpdate + 1
                        Catch ex As Exception
                            clsWrite.CaptureLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\FNO\", " OPT UPDATE ERR ==> " + strSql, "err")
                        End Try
                    Else
                        Console.WriteLine("  OPT  already exists : " & exch & " " & OptionType & " " & expdate.ToString("MM/dd/yyyy").Trim() & " " & strike_Price)
                    End If

                Else '' NO CORR
                        If exch = "NIFTY 50" Or exch = "NIFTY" Or exch = "BANKNIFTY" Then '' INSERT ALL YEARS ENTRY FOR NIFTY and BKX ONLY
                        ''If isfuturesFound And isStockFound Then
                        Debug.WriteLine(splt(2).ToString().Trim() & " " & splt(1).ToString().Trim())
                        Dim strSql As String
                        Try
                            acc_vol = acc_vol_incontracts * lotSize

                            strSql = "insert into OPTION_ITEMS (OPTION_TYPE, IS_TRADED,instrument, Open_Interest" & types & ", PREVDAY_OPEN_INTEREST" & types & ", LAST_PRICE" & types & " ,  exchange_symbol, PREVDAY_CLOSE" & types & "," &
                                    " acc_volume" & types & ", strike_price, EXPIRY_DATE, filedate, UPDATE_DATE_TIME_BHAV) values" &
                                    "('" & OptionType & "',1,'" & splt(0).ToString().Trim() & "','" & OI & "','" & PrevOI & "','" & lastprice & "'," &
                                    "'" & exch & "','" & PrevClose & "','" & acc_vol & "'" &
                                    ",'" & strike_Price & "','" & expdate & "', '" & file2download & "', getDate())"
                            SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)
                            ' SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)
                            Console.WriteLine("INSERT OPT INDEX " & exch & " " & OptionType & " " & expdate.ToString("MM/dd/yyyy").Trim() & " " & strike_Price)
                            clsWrite.CaptureInsertsLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\FNO\", " OPT INSERT Count ==> " + counterOPTInsert.ToString() + " .." + strSql)
                            counterOPTInsert = counterOPTInsert + 1


                        Catch ex As Exception
                            clsWrite.CaptureLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\FNO\", " OPT INSERT ERR ==> " + strSql, "err")
                        End Try

                        ''Else '' SINCE THERE IS NO CORRESPONDING FUTURES AVAILABLE FUTURES_ID DONT ADD
                        ''  ''' fn_id = 0
                        ''End If '' SINCE THERE IS NO CORRESPONDING FUTURES AVAILABLE FUTURES_ID DONT ADD
                    Else
                        '' If isfuturesFound And isStockFound Then '' INSERT ONLY IF CORRESPONDING FUTURES IS AVAILABLE
                        Debug.WriteLine(splt(2).ToString().Trim() & " " & splt(1).ToString().Trim())
                        Dim strSql As String
                        Try
                            acc_vol = acc_vol_incontracts * lotSize

                            strSql = "insert into OPTION_ITEMS (OPTION_TYPE, IS_TRADED, instrument,  Open_Interest" & types & ", PREVDAY_OPEN_INTEREST" & types & ", LAST_PRICE" & types & " ,  exchange_symbol, PREVDAY_CLOSE" & types & "," &
                                    " acc_volume" & types & ", strike_price, EXPIRY_DATE, filedate, UPDATE_DATE_TIME_BHAV) values" &
                                    "('" & OptionType & "',1, '" & splt(0).ToString().Trim() & "','" & OI & "','" & PrevOI & "','" & lastprice & "'," &
                                    "'" & exch & "','" & PrevClose & "','" & acc_vol & "'" &
                                    ",'" & strike_Price & "','" & expdate & "','" & file2download & "',getdate())"
                            SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)
                            '  SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)
                            Console.WriteLine("INSERT OPT INDEX " & exch & " " & OptionType & " " & expdate.ToString("MM/dd/yyyy").Trim() & " " & strike_Price)
                            clsWrite.CaptureInsertsLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\FNO\", " OPT INSERT Count ==> " + counterOPTInsert.ToString() + " .." + strSql)
                            counterOPTInsert = counterOPTInsert + 1


                        Catch ex As Exception
                            clsWrite.CaptureLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\FNO\", " OPT INSERT ERR ==> " + strSql, "err")
                        End Try
                        '' Else
                        '' '  clsWrite.CaptureLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\FNO\", " OPT INSERT ERR ==> ", "err")
                        ''End If
                    End If


                End If

            Catch ex As Exception
                Console.WriteLine("Update Main DB" & ex.Message)
            End Try


        End If


    End Sub



#End Region

#Region "SendMail"

    Sub Mail(FromEmailAddress As String, ToEmailAddress As String, mailSubject As String, MailHeader As String, MailDesc As String)


        Dim strMailHeader As String = MailHeader
        Dim Htmlbody As String = "<html>"
        Htmlbody += "<table border =1>"
        Htmlbody += MailDesc
        Htmlbody += "<tr>"
        Htmlbody += "<td> Mail send at " + DateTime.Now
        Htmlbody += "</td>"
        Htmlbody += "</tr>"
        Htmlbody += "</table>"
        Htmlbody += "</html>"



        Dim mailPassword As String = "Biz@20222"
        Dim secureMailpassword As SecureString = New SecureString()

        For Each c As Char In mailPassword
            secureMailpassword.AppendChar(c)
        Next


        Dim message As MailMessage = New MailMessage() '' = New System.Net.Mail.MailMessage(FromEmailAddress, ToEmailAddress, mailSubject, Htmlbody)


        message.From = New MailAddress(FromEmailAddress)


        Dim emailGroup() As String = ToEmailAddress.Split(";")
        For Each Multiemail As String In emailGroup
            message.To.Add(New MailAddress(Multiemail))
        Next
        message.Subject = mailSubject
        message.Body = Htmlbody
        message.BodyEncoding = Encoding.UTF8
        message.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure
        message.IsBodyHtml = True

        Dim smtp As SmtpClient = New SmtpClient()
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
        smtp.Port = 587
        smtp.Host = "smtp-mail.outlook.com"
        smtp.EnableSsl = True
        smtp.Timeout = 100000
        smtp.DeliveryMethod = SmtpDeliveryMethod.Network
        smtp.UseDefaultCredentials = True
        smtp.Credentials = New NetworkCredential(FromEmailAddress, secureMailpassword)

        smtp.Send(message)
        message = Nothing
        smtp = Nothing

        '' Console.WriteLine("ALERT " + Subject + " SENT to " + ToEmailAddress + " at " + DateTime.Now.ToString())
    End Sub

    Sub Mailss(FromEmailAddress As String, ToEmailAddress As String, mailSubject As String, MailHeader As String, MailDesc As String)


        '   Dim ToMailAddress As String = ReceiverMailId
        Dim Subject As String = mailSubject
        Dim strMailHeader As String = MailHeader
        Dim Htmlbody As String = "<html>"
        Htmlbody += "<table border =1>"
        Htmlbody += MailDesc
        Htmlbody += "<tr>"
        Htmlbody += "<td> Mail send at " + DateTime.Now
        Htmlbody += "</td>"
        Htmlbody += "</tr>"
        Htmlbody += "</table>"
        Htmlbody += "</html>"

        Dim client As SmtpClient = New SmtpClient()

        Dim mailPassword As String = "Biz@20222"
        Dim secureMailpassword As SecureString = New SecureString()

        For Each c As Char In mailPassword
            secureMailpassword.AppendChar(c)
        Next


        System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12
        client.Port = 587
        client.Host = "smtp-mail.outlook.com"
        'client.Host = "smtp-mail.outlook.com"
        client.EnableSsl = True
        client.Timeout = 100000
        client.DeliveryMethod = SmtpDeliveryMethod.Network
        client.UseDefaultCredentials = True
        client.Credentials = New System.Net.NetworkCredential(FromEmailAddress, secureMailpassword)
        Dim reportEmail As System.Net.Mail.MailMessage = New System.Net.Mail.MailMessage(FromEmailAddress, ToEmailAddress, Subject, Htmlbody)

        reportEmail.BodyEncoding = System.Text.UTF8Encoding.UTF8
        reportEmail.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure
        reportEmail.IsBodyHtml = True


        client.Send(reportEmail)
        reportEmail = Nothing
        client = Nothing

        Console.WriteLine("ALERT " + Subject + " SENT to " + ToEmailAddress + " at " + DateTime.Now.ToString())
    End Sub
#End Region

#Region "Not in use"
    Sub ReadBSEDb(str As String)
        Dim splt() As String = str.Split(",")



        Dim exch As String
        Dim OptionType As String
        Dim Open As Double
        Dim High As Double
        Dim Low As Double
        Dim PrevClose As Double
        Dim lastprice As Double
        Dim Volume As Long
        Dim Name As String
        Dim counterUpdate As Integer = 1
        Dim counterInsert As Integer = 1

        exch = splt(0).ToString().Trim()
        lastprice = splt(8).ToString().Trim()
        Name = splt(1).ToString().Trim()
        PrevClose = splt(7).ToString().Trim()
        Volume = splt(11).ToString().Trim()
        Open = splt(4).ToString().Trim()
        High = splt(5).ToString().Trim()
        Low = splt(6).ToString().Trim()


        Dim Sql As String = " select LAST_PRICE from STOCK_ITEMS_TRANSACTION where exchangesymbol ='" & exch & "'"
        Dim dt As DataTable = New DataTable
        Try
            dt = SqlHelper.ExecuteDataset(My.Settings.conn_str, CommandType.Text, Sql).Tables(0)

            If dt.Rows.Count > 0 Then
                If (Convert.ToDouble(dt.Rows(0)("LAST_PRICE")).ToString("#.00") <> Convert.ToDouble(splt(8).Trim()).ToString("#.00")) Then

                    Try
                        'LAST_PRICE='" & Trim(arr(7)) & "',DAY_OPEN='" & Trim(arr(4)) & "',DAY_HIGH='" & Trim(arr(5)) & "',DAY_LOW='" & Trim(arr(6)) & "' where EXCHANGESYMBOL='" & Trim(arr(0)) & "'"

                        Dim strSql As String = "update STOCK_ITEMS_TRANSACTION SET LAST_PRICE ='" & lastprice & "',ACC_VOLUME='" & Volume & "',DAY_OPEN='" & Open & "',DAY_HIGH='" & High & "',DAY_LOW='" & Low & "' where exchangesymbol ='" & exch & "' and series ='EQ' and exchange_id =2 and instrument_id =2"

                        SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)
                        'Debug.WriteLine(exch & " " & " DB VALUE " & dt.Rows(0)("LAST_PRICE") & ". EXCHANGE VALUE " & lastprice & " " & exch)
                        clsWrite.CaptureLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\BSE\", "", " INSERT Count ==> " + counterUpdate.ToString() + " .." + strSql)
                        counterUpdate = counterUpdate + 1
                    Catch ex As Exception

                    End Try

                Else

                End If

            Else
                'Debug.WriteLine(splt(2).ToString().Trim() & " " & splt(1).ToString().Trim())
                Try
                    'If splt(0).ToString().Trim() = "N" Then '' STOCK_ITEMS


                    If (splt(0).ToString().Trim() <> "") Then
                        Dim strSql As String = "insert into STOCK_ITEMS_TRANSACTION (" &
                            "  exchange_id,instrument_id,group_id,name,graphic_name, " &
                            " exchangesymbol, last_price, PREVDAY_CLOSE, day_open, day_high, day_low," &
                            " acc_volume,trade_volume,series,free_float,face_value,update_date_time," &
                            " market_type,upper_circuit,lower_circuit,open_interest," &
                            " open_int_change,open_int_close,yield,yld_netchange,isin,bbop,bboq,bsop,bsoq,lot_size,bridgesymbol,average_trade_price_fut) values" &
                            "(2,2,19,'" & Name & "','" & Name & "'," &
                            " '" & exch & "','" & lastprice & "','" & PrevClose & "','" & Open & "','" & High & "','" & Low & "'," &
                            " '" & Volume & "',0,'EQ',0,0,getDate()," &
                            " 'N',0,0,0" &
                            " ,0,0,0,0,0,0,0,0,0,0,'',0)"
                        SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)

                        clsWrite.CaptureLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\BSE\", "", " INSERT Count ==> " + counterInsert.ToString() + " .." + strSql)
                        counterInsert = counterInsert + 1
                    End If
                    ' End If
                Catch ex As Exception

                End Try


            End If

        Catch ex As Exception
            Console.WriteLine("Update Main DB" & ex.Message)
        End Try
    End Sub


    '''    Sub NSE_MM()
    '''        Dim file2download As String = "Pd" & strddMMyy & ".csv"
    '''        Dim url2download As String = "http://www.nseindia.com/content/equities/PR.zip"

    '''        Dim localfile As String = Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\NSE\PR.zip"
    '''        DownloadFile(url2download, localfile)
    '''        UnzipFile(localfile, file2download, Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\NSE")


    '''        File.Delete(localfile)


    '''        Dim fileName As String = Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\NSE\" & file2download
    '''        Dim fs As New FileStream(Trim(fileName), FileMode.Open, FileAccess.Read)
    '''        Dim sr As New StreamReader(fs)

    '''        Dim str As String = sr.ReadLine

    '''        Dim i As Integer = 0

    '''        Do Until str Is Nothing
    '''            If i <> 0 Then
    '''                Try
    '''                    Dim splt() As String = str.Split(",")
    '''                    Dim exch As String
    '''                    Dim Fhi_52_wk As Double
    '''                    Dim Flo_52_wk As Double
    '''                    Dim Open As Double
    '''                    Dim High As Double
    '''                    Dim Low As Double
    '''                    Dim PrevClose As Double
    '''                    Dim lastprice As Double
    '''                    Dim Volume As Long
    '''                    Dim Name As String
    '''                    Dim Series As String
    '''                    Dim isIndex As String

    '''                    isIndex = splt(0).ToString().Trim()

    '''                    If (isIndex = "Y") Then
    '''                        exch = splt(3).ToString().Trim()

    '''                        If (exch.ToUpper() = "NIFTY 50") Then
    '''                            exch = "S&P CNX NIFTY"
    '''                        End If
    '''                    ElseIf (isIndex = "N") Then
    '''                        exch = splt(2).ToString().Trim()
    '''                    Else
    '''                        GoTo nextloop
    '''                    End If

    '''                    Series = splt(1).ToString().Trim()





    '''                    Name = splt(3).ToString().Trim()
    '''                    PrevClose = splt(4).ToString().Trim()
    '''                    Open = splt(5).ToString().Trim()
    '''                    High = splt(6).ToString().Trim()
    '''                    Low = splt(7).ToString().Trim()
    '''                    lastprice = splt(8).ToString().Trim()

    '''                    Volume = splt(10).ToString().Trim()
    '''                    Fhi_52_wk = splt(13).ToString().Trim()
    '''                    Flo_52_wk = splt(14).ToString().Trim()


    '''                    Dim counterInsert As Integer = 1
    '''                    Dim counterUpdate As Integer = 1
    '''                    Dim Sql As String = " select LAST_PRICE from mj_live_prices where exchangesymbol ='" & exch & "'"
    '''                    Dim dt As DataTable = New DataTable
    '''                    Try
    '''                        dt = SqlHelper.ExecuteDataset(My.Settings.conn_str_bkup_MM, CommandType.Text, Sql).Tables(0)
    '''                        If dt.Rows.Count > 0 Then
    '''                            If (Convert.ToDouble(dt.Rows(0)("LAST_PRICE")).ToString("#.00") <> Convert.ToDouble(lastprice).ToString("#.00")) Then
    '''                                Try
    '''                                    Dim strSql As String = "update mj_live_prices set LAST_PRICE ='" & lastprice & "',ACC_VOLUME='" & Volume & "',  DAY_OPEN='" & Open & "',DAY_HIGH='" & High & "'" &
    '''                                                    " ,DAY_LOW='" & Low & "' where exchangesymbol ='" & exch & "' and exchange_id =1 and instrument_id =2"
    '''                                    SqlHelper.ExecuteNonQuery(My.Settings.conn_str_bkup_MM, CommandType.Text, strSql)
    '''                                    'Debug.WriteLine(splt(2).ToString().Trim() & " " & " DB VALUE " & dt.Rows(0)("LAST_PRICE") & ". EXCHANGE VALUE " & splt(8).ToString().Trim() & " " & splt(1).ToString().Trim())
    '''                                    clsWrite.CaptureLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\NSE\", " UPDATE Count ==> " + counterUpdate.ToString() + " .." + strSql)
    '''                                    counterUpdate = counterUpdate + 1

    '''                                Catch ex As Exception

    '''                                End Try

    '''                            Else

    '''                            End If
    '''                        Else

    '''                            Try
    '''                                If isIndex.Trim() = "N" Then '' STOCK_ITEMS
    '''                                    If (exch.Trim() <> "") Then
    '''                                        Dim strSql As String = "insert into mj_live_prices (" &
    '''                                "  exchange_id,instrument_id, " &
    '''                                " exchangesymbol,last_price, day_open, day_high, day_low, " &
    '''                                " acc_volume,value_TRADED, BROADCAST_DATETIME," &
    '''                               " bbop,bboq,bsop,bsoq) values" &
    '''                                "(1,2," &
    '''                                " '" & exch & "','" & lastprice & "','" & Open & "','" & High & "','" & Low & "'," &
    '''                                " '" & Volume & "',0,getDate()," &
    '''                                " 0,0,0,0)"

    '''                                        SqlHelper.ExecuteNonQuery(My.Settings.conn_str_bkup_MM, CommandType.Text, strSql)
    '''                                        clsWrite.CaptureLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\NSE\", "", " INSERT Count ==> " + counterInsert.ToString() + " .." + strSql)
    '''                                        counterInsert = counterInsert + 1
    '''                                    End If
    '''                                End If
    '''                            Catch ex As Exception
    '''                                clsWrite.CaptureLogs(Environment.CurrentDirectory & "\bhavfiles\" & strddMMyy & "\NSE\", "", ex.Message)
    '''                            End Try


    '''                        End If

    '''                    Catch ex As Exception
    '''                        Console.WriteLine("Update Main DB" & ex.Message)
    '''                    End Try
    '''                Catch ex As Exception

    '''                End Try
    '''            End If
    '''nextloop:
    '''            str = sr.ReadLine
    '''            i = 1
    '''        Loop


    '''    End Sub

    '''    Sub BSE_MM()


    '''        Dim file2download As String = "EQ" & strddMMyy & ".csv"
    '''        Dim url2download As String = "http://www.bseindia.com/bhavcopy/" & "EQ" & strddMMyy & "_csv.zip"

    '''        Dim localfile As String = Environment.CurrentDirectory & "\bhavfiles\" & Format(Dates, "ddMMyy") & "\BSE\" & "EQ" & Format(Dates, "ddMMyy") & "_csv.zip"
    '''        DownloadFile(url2download, localfile)
    '''        'Exit Sub
    '''        UnzipFile(localfile, file2download, Environment.CurrentDirectory & "\bhavfiles\" & Format(Dates, "ddMMyy") & "\BSE\")

    '''        File.Delete(localfile)
    '''        Console.WriteLine("BSE File Downloaded..." & file2download)


    '''        ''''ReadCSVFile_BSE(Environment.CurrentDirectory & "\bhavfiles\" & Format(Dates, "ddMMyy") & "\" & file2download)





    '''        Dim fileName As String = Environment.CurrentDirectory & "\bhavfiles\" & Format(Dates, "ddMMyy") & "\BSE\" & file2download
    '''        Dim fs As New FileStream(Trim(fileName), FileMode.Open, FileAccess.Read)
    '''        Dim sr As New StreamReader(fs)

    '''        Dim str As String
    '''        str = sr.ReadLine

    '''        Dim i As Integer
    '''        i = 0
    '''        Do Until str Is Nothing

    '''            If i <> 0 Then
    '''                Dim splt() As String = str.Split(",")



    '''                Dim exch As String
    '''                Dim OptionType As String
    '''                Dim Open As Double
    '''                Dim High As Double
    '''                Dim Low As Double
    '''                Dim PrevClose As Double
    '''                Dim lastprice As Double
    '''                Dim Volume As Long
    '''                Dim Name As String
    '''                Dim counterUpdate As Integer = 1
    '''                Dim counterInsert As Integer = 1

    '''                exch = splt(0).ToString().Trim()
    '''                lastprice = splt(8).ToString().Trim()
    '''                Name = splt(1).ToString().Trim()
    '''                PrevClose = splt(7).ToString().Trim()
    '''                Volume = splt(11).ToString().Trim()
    '''                Open = splt(4).ToString().Trim()
    '''                High = splt(5).ToString().Trim()
    '''                Low = splt(6).ToString().Trim()


    '''                Dim Sql As String = " select LAST_PRICE from mj_live_prices where exchangesymbol ='" & exch & "'"
    '''                Dim dt As DataTable = New DataTable
    '''                Try
    '''                    dt = SqlHelper.ExecuteDataset(My.Settings.conn_str_bkup_MM, CommandType.Text, Sql).Tables(0)

    '''                    If dt.Rows.Count > 0 Then
    '''                        If (Convert.ToDouble(dt.Rows(0)("LAST_PRICE")).ToString("#.00") <> Convert.ToDouble(splt(8).Trim()).ToString("#.00")) Then

    '''                            Try
    '''                                'LAST_PRICE='" & Trim(arr(7)) & "',DAY_OPEN='" & Trim(arr(4)) & "',DAY_HIGH='" & Trim(arr(5)) & "',DAY_LOW='" & Trim(arr(6)) & "' where EXCHANGESYMBOL='" & Trim(arr(0)) & "'"

    '''                                Dim strSql As String = "update mj_live_prices SET LAST_PRICE ='" & lastprice & "',ACC_VOLUME='" & Volume & "',DAY_OPEN='" & Open & "',DAY_HIGH='" & High & "',DAY_LOW='" & Low & "' where exchangesymbol ='" & exch & "' and exchange_id =2 and instrument_id =2"

    '''                                SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strSql)
    '''                                'Debug.WriteLine(exch & " " & " DB VALUE " & dt.Rows(0)("LAST_PRICE") & ". EXCHANGE VALUE " & lastprice & " " & exch)
    '''                                clsWrite.CaptureLogs(Environment.CurrentDirectory & "\bhavfiles\" & Format(Dates, "ddMMyy") & "\BSE\", "", " INSERT Count ==> " + counterUpdate.ToString() + " .." + strSql)
    '''                                counterUpdate = counterUpdate + 1
    '''                            Catch ex As Exception

    '''                            End Try

    '''                        Else

    '''                        End If

    '''                    Else
    '''                        'Debug.WriteLine(splt(2).ToString().Trim() & " " & splt(1).ToString().Trim())
    '''                        Try
    '''                            'If splt(0).ToString().Trim() = "N" Then '' STOCK_ITEMS


    '''                            If (splt(0).ToString().Trim() <> "") Then
    '''                                Dim STRSQL As String = "INSERT INTO mj_live_prices (" &
    '''                            "  EXCHANGE_ID,INSTRUMENT_ID, " &
    '''                            " EXCHANGESYMBOL, LAST_PRICE,  DAY_OPEN, DAY_HIGH, DAY_LOW," &
    '''                            " ACC_VOLUME, VALUE_TRADED, BROADCAST_DATETIME," &
    '''                            " BBOP,BBOQ,BSOP,BSOQ) VALUES" &
    '''                            "(2,2," &
    '''                            " '" & exch & "','" & lastprice & "','" & Open & "','" & High & "','" & Low & "'," &
    '''                            " '" & Volume & "',0,'EQ',0,0,getDate()" &
    '''                            " ,0,0,0,0)"
    '''                                SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, STRSQL)

    '''                                clsWrite.CaptureLogs(Environment.CurrentDirectory & "\bhavfiles\" & Format(Dates, "ddMMyy") & "\BSE\", "", " INSERT Count ==> " + counterInsert.ToString() + " .." + STRSQL)
    '''                                counterInsert = counterInsert + 1
    '''                            End If
    '''                            ' End If
    '''                        Catch ex As Exception

    '''                        End Try


    '''                    End If

    '''                Catch ex As Exception
    '''                    Console.WriteLine("Update Main DB" & ex.Message)
    '''                End Try

    '''            End If
    '''            str = sr.ReadLine
    '''            i = 1
    '''        Loop




    '''    End Sub
    '''    

    'Public Static void DeleteOldLogFiles(String FolderNameToCheck, int daysAgo = 7, int maxToDelete = 10000)
    '    {


    '        String tempDir = Application.StartupPath + "\\" + FolderNameToCheck;
    '        If (!Directory.Exists(tempDir))
    '        {
    '            Directory.CreateDirectory(tempDir);
    '        }
    '        String[] files = Directory.GetFiles(tempDir, "*.log", SearchOption.TopDirectoryOnly);
    '        If (files.Length > 0)
    '        {
    '            String[] filesToDelete = files.Where(c =>
    '            {
    '                TimeSpan ts = DateTime.Now - File.GetLastAccessTime(c);
    '                Return (ts.Days > daysAgo);
    '            }).ToArray();
    '            For (int i = 0; i < Math.Min(filesToDelete.Length, maxToDelete); i++)
    '            {
    '                File.Delete(filesToDelete[i]);
    '            }
    '        }
    '    }
#End Region

End Module
