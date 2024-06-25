Imports BQ.DAL
Public Class Class1

    Public Sub calculate()
        Dim Sql As String = " select STOCK_ID from SECTOR_STOCKS_DETAILS"
        Dim dt As DataTable = New DataTable
        Try


            dt = SqlHelper.ExecuteDataset(My.Settings.conn_str_bkup, CommandType.Text, Sql).Tables(0)

            If dt.Rows.Count > 0 Then
                If (Convert.ToDouble(dt.Rows(0)("LAST_PRICE")).ToString("#.00") <> Convert.ToDouble(lastprice).ToString("#.00")) Then
                    Try
                        Dim strSql As String = "update STOCKS_TRANSACTION set LAST_PRICE ='" & lastprice & "', OPEN_INTEREST='" & OI & "', DAY_OPEN='" & Open & "',DAY_HIGH='" & High & "',DAY_LOW='" & Low & "'" &
                        "   where exchangesymbol ='" & exch & "' and exchange_id =1 and instrument_id =3 and FUT_EXPIRY_DATE ='" & expdate.ToString("MM/dd/yyyy").Trim() & "'"

                        SqlHelper.ExecuteNonQuery(My.Settings.conn_str_bkup, CommandType.Text, strSql)
                        Debug.WriteLine(splt(2).ToString().Trim() & " " & " DB VALUE " & dt.Rows(0)("LAST_PRICE") & ". EXCHANGE VALUE " & lastprice & " " & expdate)
                    Catch ex As Exception

                    End Try
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

End Class
