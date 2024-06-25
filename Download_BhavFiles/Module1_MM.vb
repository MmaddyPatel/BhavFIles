Imports System.Net
Imports System.IO
Imports Ionic.Zip
Imports ZEE.DAL

Module Module1_MM

#Region "DECLERATIONS"

    Dim Dates As Date = Now.Date
    Dim indices As String
    Dim stocks_updated As Integer
    Dim localfile As String = System.Reflection.Assembly.GetExecutingAssembly().Location & Format(Dates, "ddMMyy") & ".csv"
    Dim url2download As String = "https://www.nseindia.com/content/equities/sec_list.csv"
    Dim startTime As DateTime
    Dim clsWrite As New WriteToLogs
    Dim strStartUpPath = System.Reflection.Assembly.GetExecutingAssembly().Location

#End Region
#Region "FORM EVENTS"
    Sub Main1()



        If Directory.Exists(System.Reflection.Assembly.GetExecutingAssembly().Location & "\" & Format(Dates, "ddMMyy") & "\BSE") = False Then
            Directory.CreateDirectory(System.Reflection.Assembly.GetExecutingAssembly().Location & "\" & Format(Dates, "ddMMyy") & "\BSE")
        End If

        If Directory.Exists(System.Reflection.Assembly.GetExecutingAssembly().Location & "\" & Format(Dates, "ddMMyy") & "\NSE") = False Then
            Directory.CreateDirectory(System.Reflection.Assembly.GetExecutingAssembly().Location & "\" & Format(Dates, "ddMMyy") & "\NSE")
        End If


        Try
            Dim strsql As String = "exec SP_HOLIDAYCHECK"
            SqlHelper.ExecuteNonQuery(My.Settings.conn_str, CommandType.Text, strsql)


            Try
                Call FNO()
            Catch ex As Exception

            End Try



        Catch ex As Exception

        End Try



        End


    End Sub
#End Region

    Public Function DownloadFile(ByVal url As String, ByVal localfile As String) As Boolean
        ' file2download = "eq_bands_24082016.csv"

        'url = "http://www.nseindia.com/content/equities/eq_bands_12092016.csv"
        Dim _WebClient As New System.Net.WebClient
        _WebClient.UseDefaultCredentials = True
        _WebClient.Headers("User-Agent") = "Mozilla/5.0 (compatible; MSIE 9.0; windows NT 6.1; WOW64; Trident/5.0)"
        _WebClient.Headers("Method") = "GET"
        _WebClient.Headers("AllowAutoRedirect") = True
        _WebClient.Headers("KeepAlive") = True
        _WebClient.Headers("Accept") = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
        Try
            _WebClient.DownloadFile(New Uri(url), localfile)
            Return True
        Catch ex As Exception
            Console.WriteLine(" ERR IN DOWNLOADING " & url & " ==> " & ex.Message)
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

#Region "TEST NSE"


#End Region

#Region "TEST BSE"
    '''    Sub BSE_MM()


    '''        Dim file2download As String = "EQ" & Format(Dates, "ddMMyy") & ".csv"
    '''        Dim url2download As String = "http://www.bseindia.com/DOWNLOAD/bhavcopy/EQUITY" & "EQ" & Format(Dates, "ddMMyy") & "_csv.zip"

    '''        'http://www.bseindia.com/download/BhavCopy/Equity/EQ240317_CSV.ZIP

    '''        Dim localfile As String = System.Reflection.Assembly.GetExecutingAssembly().Location & Format(Dates, "ddMMyy") & "\BSE\" & "EQ" & Format(Dates, "ddMMyy") & "_csv.zip"
    '''        DownloadFile(url2download, localfile)
    '''        'Exit Sub
    '''        UnzipFile(localfile, file2download, System.Reflection.Assembly.GetExecutingAssembly().Location & Format(Dates, "ddMMyy") & "\BSE\")

    '''        File.Delete(localfile)
    '''        Console.WriteLine("BSE File Downloaded..." & file2download)


    '''        ''''ReadCSVFile_BSE(System.Reflection.Assembly.GetExecutingAssembly().Location & Format(Dates, "ddMMyy") & "\" & file2download)





    '''        Dim fileName As String = System.Reflection.Assembly.GetExecutingAssembly().Location & Format(Dates, "ddMMyy") & "\BSE\" & file2download
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
    '''                    dt = SqlHelper.ExecuteDataset(My.Settings.conn_str_bkup, CommandType.Text, Sql).Tables(0)

    '''                    If dt.Rows.Count > 0 Then
    '''                        If (Convert.ToDouble(dt.Rows(0)("LAST_PRICE")).ToString("#.00") <> Convert.ToDouble(splt(8).Trim()).ToString("#.00")) Then

    '''                            Try
    '''                                'LAST_PRICE='" & Trim(arr(7)) & "',DAY_OPEN='" & Trim(arr(4)) & "',DAY_HIGH='" & Trim(arr(5)) & "',DAY_LOW='" & Trim(arr(6)) & "' where EXCHANGESYMBOL='" & Trim(arr(0)) & "'"

    '''                                Dim strSql As String = "update mj_live_prices SET LAST_PRICE ='" & lastprice & "',ACC_VOLUME='" & Volume & "',DAY_OPEN='" & Open & "',DAY_HIGH='" & High & "',DAY_LOW='" & Low & "' where exchangesymbol ='" & exch & "' and exchange_id =2 and instrument_id =2"

    '''                                ' SqlHelper.ExecuteNonQuery(My.Settings.conn_str_bkup, CommandType.Text, strSql)
    '''                                'Debug.WriteLine(exch & " " & " DB VALUE " & dt.Rows(0)("LAST_PRICE") & ". EXCHANGE VALUE " & lastprice & " " & exch)
    '''                                clsWrite.CaptureLogs(System.Reflection.Assembly.GetExecutingAssembly().Location & Format(Dates, "ddMMyy") & "\BSE\", "", " INSERT Count ==> " + counterUpdate.ToString() + " .." + strSql)
    '''                                counterUpdate = counterUpdate + 1
    '''                            Catch ex As Exception

    '''                            End Try

    '''                        Else

    '''                        End If

    '''                    Else
    '''                        'Debug.WriteLine(splt(2).ToString().Trim() & " " & splt(1).ToString().Trim())
    '''                        Try
    '''                            'If splt(0).ToString().Trim() = "N" Then '' STOCKS


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
    '''                                '    SqlHelper.ExecuteNonQuery(My.Settings.conn_str_bkup, CommandType.Text, strSql)

    '''                                clsWrite.CaptureLogs(System.Reflection.Assembly.GetExecutingAssembly().Location & Format(Dates, "ddMMyy") & "\BSE\", "", " INSERT Count ==> " + counterInsert.ToString() + " .." + STRSQL)
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

    '''    Sub NSE_MM()
    '''        Dim file2download As String = "Pd" & Format(Dates, "ddMMyy") & ".csv"
    '''        Dim url2download As String = "http://www.nseindia.com/content/equities/PR.zip"

    '''        Dim localfile As String = System.Reflection.Assembly.GetExecutingAssembly().Location & Format(Dates, "ddMMyy") & "\NSE\PR.zip"
    '''        DownloadFile(url2download, localfile)
    '''        UnzipFile(localfile, file2download, System.Reflection.Assembly.GetExecutingAssembly().Location & Format(Dates, "ddMMyy") & "\NSE")


    '''        File.Delete(localfile)


    '''        Dim fileName As String = System.Reflection.Assembly.GetExecutingAssembly().Location & Format(Dates, "ddMMyy") & "\NSE\" & file2download
    '''        Dim fs As New FileStream(Trim(fileName), FileMode.Open, FileAccess.Read)
    '''        Dim sr As New StreamReader(fs)

    '''        Dim str As String = sr.ReadLine

    '''        Dim i As Integer = 0

    '''        Do Until str Is Nothing
    '''            If i <> 0 Then
    '''                Try
    '''                    Dim splt() As String = str.Split(",")
    '''                    Dim exch As String
    '''                    Dim F52Week_High As Double
    '''                    Dim F52Week_Low As Double
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
    '''                    F52Week_High = splt(13).ToString().Trim()
    '''                    F52Week_Low = splt(14).ToString().Trim()


    '''                    Dim counterInsert As Integer = 1
    '''                    Dim counterUpdate As Integer = 1
    '''                    Dim Sql As String = " select LAST_PRICE from mj_live_prices where exchangesymbol ='" & exch & "'"
    '''                    Dim dt As DataTable = New DataTable
    '''                    Try
    '''                        dt = SqlHelper.ExecuteDataset(My.Settings.conn_str_bkup, CommandType.Text, Sql).Tables(0)
    '''                        If dt.Rows.Count > 0 Then
    '''                            If (Convert.ToDouble(dt.Rows(0)("LAST_PRICE")).ToString("#.00") <> Convert.ToDouble(lastprice).ToString("#.00")) Then
    '''                                Try
    '''                                    Dim strSql As String = "update mj_live_prices set LAST_PRICE ='" & lastprice & "',ACC_VOLUME='" & Volume & "',  DAY_OPEN='" & Open & "',DAY_HIGH='" & High & "'" &
    '''                                                    " ,DAY_LOW='" & Low & "' where exchangesymbol ='" & exch & "' and exchange_id =1 and instrument_id =2"
    '''                                    ' SqlHelper.ExecuteNonQuery(My.Settings.conn_str_bkup, CommandType.Text, strSql)
    '''                                    'Debug.WriteLine(splt(2).ToString().Trim() & " " & " DB VALUE " & dt.Rows(0)("LAST_PRICE") & ". EXCHANGE VALUE " & splt(8).ToString().Trim() & " " & splt(1).ToString().Trim())
    '''                                    clsWrite.CaptureLogs(System.Reflection.Assembly.GetExecutingAssembly().Location & Format(Dates, "ddMMyy") & "\NSE\", " UPDATE Count ==> " + counterUpdate.ToString() + " .." + strSql)
    '''                                    counterUpdate = counterUpdate + 1

    '''                                Catch ex As Exception

    '''                                End Try

    '''                            Else

    '''                            End If
    '''                        Else

    '''                            Try
    '''                                If isIndex.Trim() = "N" Then '' STOCKS
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


    '''                                        clsWrite.CaptureLogs(System.Reflection.Assembly.GetExecutingAssembly().Location & Format(Dates, "ddMMyy") & "\NSE\", "", " INSERT Count ==> " + counterInsert.ToString() + " .." + strSql)
    '''                                        counterInsert = counterInsert + 1
    '''                                    End If
    '''                                End If
    '''                            Catch ex As Exception
    '''                                clsWrite.CaptureLogs(System.Reflection.Assembly.GetExecutingAssembly().Location & Format(Dates, "ddMMyy") & "\NSE\", "", ex.Message)
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
#End Region

End Module
