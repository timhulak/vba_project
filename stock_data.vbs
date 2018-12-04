Sub stock_script()
    Dim WS_Count As Integer
         Dim x As Integer

         WS_Count = ActiveWorkbook.Worksheets.Count

         For x = 1 To WS_Count

            Range("I1") = "ticker"
            Range("J1") = "total_stock_volume"

            Dim ticker As String

            Dim volume As Double
            volume = 0

            Dim summary_table As Integer
            summary_table = 2

            Dim lr As Integer
            Dim last_row As Long
            last_row = Cells(Rows.Count, 1).End(xlUp).Row

            For I = 2 To last_row
                If Cells(I + 1, 1).Value <> Cells(I, 1) Then
                    ticker = Cells(I, 1).Value
                    volume = volume + Cells(I, 7).Value
                    Range("I" & summary_table).Value = ticker
                    Range("J" & summary_table).Value = volume
                    summary_table = summary_table + 1
                    volume = 0
                    Else
                    volume = volume + Cells(I, 7).Value
                End If
            Next
         Next

End Sub
