Attribute VB_Name = "Module1"
Sub Verilerin_Topsis_Hesaplanmas�()
    
    sat�r = Range("c3").End(xlDown).Row
    s�tun = Range("c3").End(xlToRight).Column
    Cells(1, s�tun + 2) = "Toplam:"
    Cells(1, s�tun + 2).Font.Color = vbRed
    Cells(1, s�tun + 2).Font.Bold = True
    
    toplam = 0
    For i = 3 To s�tun
        
            toplam = toplam + Cells(2, i)
            Cells(2, s�tun + 2) = toplam
           
    Next i
    
    For i = 3 To s�tun
    
        toplam = 0
        For j = 3 To sat�r
            toplam = toplam + Cells(j, i) ^ 2
        Next j
            Cells(sat�r + 1, i) = toplam ^ 0.5
            Cells(sat�r + 1, 2) = "Kare Toplamlar�:"
    Next i
  
   Range("b" & sat�r + 1).Font.Color = vbRed
   Range("b" & sat�r + 1).Font.Bold = True
   
   Cells(sat�r + 3, "a") = "R MATR�S�"
   
   For y = 3 To s�tun
        x = sat�r + 2
        For i = 3 To sat�r
                x = x + 1
                Cells(x, y) = Cells(i, y) / Cells(sat�r + 1, y)
                Next i
   Next y

   sat��r = Range("c" & sat�r + 3).End(xlDown).Row
   Cells(sat��r + 2, "b") = "Normalize A��rl�k:"
   Range("b" & sat��r + 2).Font.Color = vbRed
   Range("b" & sat��r + 2).Font.Bold = True
   Range("b" & sat�r + 3 & ":" & "b" & sat��r) = Range("b" & "3" & ":" & "b" & sat�r).Value
   
   For c = 3 To s�tun
       Cells(sat��r + 2, c) = Cells(2, c) / Cells(2, s�tun + 2)
   Next c
   
   Range("a" & sat��r + 4) = "V Matrisi"
   
   For d = 3 To s�tun
        v = sat��r + 3
        For t = sat�r + 3 To sat��r
                v = v + 1
                Cells(v, d) = Cells(t, d) * Cells(sat��r + 2, d)
        Next t
   Next d
   
   sat���r = Range("c" & sat��r + 4).End(xlDown).Row
   Range("b" & sat��r + 4 & ":" & "b" & sat���r) = Range("b" & "3" & ":" & "b" & sat�r).Value
   
   Range("b" & sat���r + 2) = "A+"
   Range("b" & sat���r + 2).Font.Color = vbRed
   Range("b" & sat���r + 2).Font.Bold = True
   Range("b" & sat���r + 3) = "A-"
   Range("b" & sat���r + 3).Font.Color = vbRed
   Range("b" & sat���r + 3).Font.Bold = True
   
    For j = 3 To s�tun
        x = Columns(j).Address
        x = Mid(x, 2, 1)
        Cells(sat���r + 2, j) = WorksheetFunction.Max(Range(x & sat��r + 4 & ":" & x & sat���r))
    Next j

    For j = 3 To s�tun
        x = Columns(j).Address
        x = Mid(x, 2, 1)
        Cells(sat���r + 3, j) = WorksheetFunction.Min(Range(x & sat��r + 4 & ":" & x & sat���r))
    Next j
                   
    For j = 3 To s�tun
            k = sat���r + 4
            For i = sat��r + 4 To sat���r
            k = k + 1
            sonuc = Cells(i, j) - Cells(sat���r + 2, j)
            Cells(k, j) = sonuc ^ 2
            Next i
    Next j
   
    sat����r = Range("c" & sat���r + 5).End(xlDown).Row
    Range("b" & sat���r + 5 & ":" & "b" & sat����r) = Range("b" & "3" & ":" & "b" & sat�r).Value
  
    For j = 3 To s�tun
            k = sat����r + 1
            For i = sat��r + 4 To sat���r
            k = k + 1
            sonuc = Cells(i, j) - Cells(sat���r + 3, j)
            Cells(k, j) = sonuc ^ 2
            Next i
    Next j

   sat�����r = Range("c" & sat����r + 2).End(xlDown).Row
   Range("b" & sat����r + 2 & ":" & "b" & sat�����r) = Range("b" & "3" & ":" & "b" & sat�r).Value
   
   Range("a" & sat�����r + 3) = "S+ Matrisi"
   Range("a" & sat�����r + 3).Font.Color = vbRed
   Range("a" & sat�����r + 3).Font.Bold = True
   
   k = sat�����r + 2
   For i = sat���r + 5 To sat����r
        k = k + 1
        toplam = 0
        For j = 3 To s�tun
            toplam = toplam + Cells(i, j)
        Next j
            Cells(k, 3) = toplam ^ 0.5
   Next i
   
   Range("e" & sat�����r + 3) = "S- Matrisi"
   Range("e" & sat�����r + 3).Font.Color = vbRed
   Range("e" & sat�����r + 3).Font.Bold = True
   
   k = sat�����r + 2
   For i = sat����r + 2 To sat�����r
        k = k + 1
        toplam = 0
        For j = 3 To s�tun
            toplam = toplam + Cells(i, j)
        Next j
            Cells(k, 7) = toplam ^ 0.5
    Next i
      
    sat������r = Range("c" & sat�����r + 3).End(xlDown).Row
    Range("b" & sat�����r + 3 & ":" & "b" & sat������r) = Range("b" & "3" & ":" & "b" & sat�r).Value
    Range("f" & sat�����r + 3 & ":" & "f" & sat������r) = Range("b" & "3" & ":" & "b" & sat�r).Value
    
    j = sat������r + 2
    For i = sat�����r + 3 To sat������r
        j = j + 1
        Cells(j, "e") = Cells(i, "g") / (Cells(i, "c") + Cells(i, "g"))
    Next i
    
    sat�������r = Range("e" & sat������r + 2).End(xlDown).Row
    sat��������r = Range("e" & sat�������r).End(xlDown).Row
    Range("d" & sat�������r & ":" & "d" & sat��������r) = Range("b" & "3" & ":" & "b" & sat�r).Value
    
    Cells(sat��������r + 2, "d") = "SONU�:"
    Range("d" & sat��������r + 2).Font.Italic = True
    Range("d" & sat��������r + 2).Font.Bold = True
    Cells(sat��������r + 2, 5) = WorksheetFunction.Max(Range("e" & sat�������r & ":" & "e" & sat��������r))
    
End Sub

Sub calistir()
  
    Verilerin_Topsis_Hesaplanmas�
    sat�r = Range("c3").End(xlDown).Row
    sat��r = Range("c" & sat�r + 3).End(xlDown).Row
    sat���r = Range("c" & sat��r + 4).End(xlDown).Row
    sat����r = Range("c" & sat���r + 5).End(xlDown).Row
    sat�����r = Range("c" & sat����r + 2).End(xlDown).Row
    sat������r = Range("c" & sat�����r + 3).End(xlDown).Row
    sat�������r = Range("e" & sat������r + 2).End(xlDown).Row
    sat��������r = Range("e" & sat�������r).End(xlDown).Row
    
    cevap = WorksheetFunction.Max(Range("e" & sat�������r & ":" & "e" & sat��������r))
    MsgBox (cevap & " " & "de�eri ile hangi projeyi tercih etmeniz gerekti�ine bakabilirsiniz. :)")

End Sub




