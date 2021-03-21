Attribute VB_Name = "Module1"
Sub Verilerin_Topsis_Hesaplanmasý()
    
    satýr = Range("c3").End(xlDown).Row
    sütun = Range("c3").End(xlToRight).Column
    Cells(1, sütun + 2) = "Toplam:"
    Cells(1, sütun + 2).Font.Color = vbRed
    Cells(1, sütun + 2).Font.Bold = True
    
    toplam = 0
    For i = 3 To sütun
        
            toplam = toplam + Cells(2, i)
            Cells(2, sütun + 2) = toplam
           
    Next i
    
    For i = 3 To sütun
    
        toplam = 0
        For j = 3 To satýr
            toplam = toplam + Cells(j, i) ^ 2
        Next j
            Cells(satýr + 1, i) = toplam ^ 0.5
            Cells(satýr + 1, 2) = "Kare Toplamlarý:"
    Next i
  
   Range("b" & satýr + 1).Font.Color = vbRed
   Range("b" & satýr + 1).Font.Bold = True
   
   Cells(satýr + 3, "a") = "R MATRÝSÝ"
   
   For y = 3 To sütun
        x = satýr + 2
        For i = 3 To satýr
                x = x + 1
                Cells(x, y) = Cells(i, y) / Cells(satýr + 1, y)
                Next i
   Next y

   satýýr = Range("c" & satýr + 3).End(xlDown).Row
   Cells(satýýr + 2, "b") = "Normalize Aðýrlýk:"
   Range("b" & satýýr + 2).Font.Color = vbRed
   Range("b" & satýýr + 2).Font.Bold = True
   Range("b" & satýr + 3 & ":" & "b" & satýýr) = Range("b" & "3" & ":" & "b" & satýr).Value
   
   For c = 3 To sütun
       Cells(satýýr + 2, c) = Cells(2, c) / Cells(2, sütun + 2)
   Next c
   
   Range("a" & satýýr + 4) = "V Matrisi"
   
   For d = 3 To sütun
        v = satýýr + 3
        For t = satýr + 3 To satýýr
                v = v + 1
                Cells(v, d) = Cells(t, d) * Cells(satýýr + 2, d)
        Next t
   Next d
   
   satýýýr = Range("c" & satýýr + 4).End(xlDown).Row
   Range("b" & satýýr + 4 & ":" & "b" & satýýýr) = Range("b" & "3" & ":" & "b" & satýr).Value
   
   Range("b" & satýýýr + 2) = "A+"
   Range("b" & satýýýr + 2).Font.Color = vbRed
   Range("b" & satýýýr + 2).Font.Bold = True
   Range("b" & satýýýr + 3) = "A-"
   Range("b" & satýýýr + 3).Font.Color = vbRed
   Range("b" & satýýýr + 3).Font.Bold = True
   
    For j = 3 To sütun
        x = Columns(j).Address
        x = Mid(x, 2, 1)
        Cells(satýýýr + 2, j) = WorksheetFunction.Max(Range(x & satýýr + 4 & ":" & x & satýýýr))
    Next j

    For j = 3 To sütun
        x = Columns(j).Address
        x = Mid(x, 2, 1)
        Cells(satýýýr + 3, j) = WorksheetFunction.Min(Range(x & satýýr + 4 & ":" & x & satýýýr))
    Next j
                   
    For j = 3 To sütun
            k = satýýýr + 4
            For i = satýýr + 4 To satýýýr
            k = k + 1
            sonuc = Cells(i, j) - Cells(satýýýr + 2, j)
            Cells(k, j) = sonuc ^ 2
            Next i
    Next j
   
    satýýýýr = Range("c" & satýýýr + 5).End(xlDown).Row
    Range("b" & satýýýr + 5 & ":" & "b" & satýýýýr) = Range("b" & "3" & ":" & "b" & satýr).Value
  
    For j = 3 To sütun
            k = satýýýýr + 1
            For i = satýýr + 4 To satýýýr
            k = k + 1
            sonuc = Cells(i, j) - Cells(satýýýr + 3, j)
            Cells(k, j) = sonuc ^ 2
            Next i
    Next j

   satýýýýýr = Range("c" & satýýýýr + 2).End(xlDown).Row
   Range("b" & satýýýýr + 2 & ":" & "b" & satýýýýýr) = Range("b" & "3" & ":" & "b" & satýr).Value
   
   Range("a" & satýýýýýr + 3) = "S+ Matrisi"
   Range("a" & satýýýýýr + 3).Font.Color = vbRed
   Range("a" & satýýýýýr + 3).Font.Bold = True
   
   k = satýýýýýr + 2
   For i = satýýýr + 5 To satýýýýr
        k = k + 1
        toplam = 0
        For j = 3 To sütun
            toplam = toplam + Cells(i, j)
        Next j
            Cells(k, 3) = toplam ^ 0.5
   Next i
   
   Range("e" & satýýýýýr + 3) = "S- Matrisi"
   Range("e" & satýýýýýr + 3).Font.Color = vbRed
   Range("e" & satýýýýýr + 3).Font.Bold = True
   
   k = satýýýýýr + 2
   For i = satýýýýr + 2 To satýýýýýr
        k = k + 1
        toplam = 0
        For j = 3 To sütun
            toplam = toplam + Cells(i, j)
        Next j
            Cells(k, 7) = toplam ^ 0.5
    Next i
      
    satýýýýýýr = Range("c" & satýýýýýr + 3).End(xlDown).Row
    Range("b" & satýýýýýr + 3 & ":" & "b" & satýýýýýýr) = Range("b" & "3" & ":" & "b" & satýr).Value
    Range("f" & satýýýýýr + 3 & ":" & "f" & satýýýýýýr) = Range("b" & "3" & ":" & "b" & satýr).Value
    
    j = satýýýýýýr + 2
    For i = satýýýýýr + 3 To satýýýýýýr
        j = j + 1
        Cells(j, "e") = Cells(i, "g") / (Cells(i, "c") + Cells(i, "g"))
    Next i
    
    satýýýýýýýr = Range("e" & satýýýýýýr + 2).End(xlDown).Row
    satýýýýýýýýr = Range("e" & satýýýýýýýr).End(xlDown).Row
    Range("d" & satýýýýýýýr & ":" & "d" & satýýýýýýýýr) = Range("b" & "3" & ":" & "b" & satýr).Value
    
    Cells(satýýýýýýýýr + 2, "d") = "SONUÇ:"
    Range("d" & satýýýýýýýýr + 2).Font.Italic = True
    Range("d" & satýýýýýýýýr + 2).Font.Bold = True
    Cells(satýýýýýýýýr + 2, 5) = WorksheetFunction.Max(Range("e" & satýýýýýýýr & ":" & "e" & satýýýýýýýýr))
    
End Sub

Sub calistir()
  
    Verilerin_Topsis_Hesaplanmasý
    satýr = Range("c3").End(xlDown).Row
    satýýr = Range("c" & satýr + 3).End(xlDown).Row
    satýýýr = Range("c" & satýýr + 4).End(xlDown).Row
    satýýýýr = Range("c" & satýýýr + 5).End(xlDown).Row
    satýýýýýr = Range("c" & satýýýýr + 2).End(xlDown).Row
    satýýýýýýr = Range("c" & satýýýýýr + 3).End(xlDown).Row
    satýýýýýýýr = Range("e" & satýýýýýýr + 2).End(xlDown).Row
    satýýýýýýýýr = Range("e" & satýýýýýýýr).End(xlDown).Row
    
    cevap = WorksheetFunction.Max(Range("e" & satýýýýýýýr & ":" & "e" & satýýýýýýýýr))
    MsgBox (cevap & " " & "deðeri ile hangi projeyi tercih etmeniz gerektiðine bakabilirsiniz. :)")

End Sub




