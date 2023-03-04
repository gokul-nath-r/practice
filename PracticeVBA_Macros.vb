'The notes while learning this VBA is updated in the below document
' https://docs.google.com/document/d/1dqBAp0T96oXiJRXkK5DMzpo8EsTm8tFi7pKBopq7FR8/edit?usp=sharing

'------------------------------------------------------------------------------------------------
'Refernces for learning is from tutorials point youtube
'
'''
' accessing cells in excel
''' 
Sub First_piece()
   ' MsgBox "Gokul summa va "
    ActiveCell = "poda"
    [C5] = 10
    [c14:C20] = Len([c13])
    Cells(2, 3) = "haha"
    
    Range([A1], [A10]) = "Delhi"
    Range("B2") = "Bnaglore"
End Sub

Sub CopyCells()
    'to copy cells from one set of cols to other
   'easy method for copy paste
   Range("D1:d10") = Range("a1:a10").Value
    
    'normal excel func method
    Range("a1:a10").Copy
    Range("f1:f10").PasteSpecial
    'if the below piece is not written the selection of that wavy box while copy will remain
    Application.CutCopyMode = False
    
End Sub

'''
'Usage of font properties
'font style, name , size, color
'''
Sub font1()
  teststr = Range("a1:a10").font.Name
    Range("a1:a10").font.Name = "Times New roman"
    Range("a1:a10").font.size = 12
    Range([a1], [a5]).font.Bold = True
    Range([a1], [a5]).font.ColorIndex = 10
    Range([a1], [a5]).Borders.Color = RGB(128, 0, 128)
    Range([a1], [a5]).Interior.Color = VbGray
End Sub

'''
'The goal of writing excel vba is to do repetitive tast / automation,
'But when we look at the code above we keep on writing the same piece range many times
' Thus With Block is used in vb

'Also discusses on
' Border properties, alignemnt properties
'''
Sub withBlock()
    With Range("a1:a10")
        With .font
            .Bold = False
            .Color = vbRed                      'names given for 8 std colors like red,blue green, black, gray, etc..
            .Italic = True
            .Strikethrough = True
        End With
        
        With .Borders
            .ColorIndex = 1                     ' 1 to 56s
            .LineStyle = xlDot
            .Weight = 3                         ' 1 to 4
            .LineStyle = xlNone                 'removes all the borders
        End With

        'alignment
        .HorizontalAlignment = xlRight          ' right center left
        .VerticalAlignment = xlTop              ' top center bottom

    End With
End Sub

'''
'usage of copy and paste special commands
'''
Sub Pastespl()
    Range("a2:a10").Copy
    Range("b3:b10").PasteSpecial xlPasteFormats
    Range("b3:b10").PasteSpecial xlPasteColumnWidths
    Range("a15:a25").PasteSpecial xlPasteValues
    
    Application.CutCopyMode = False              'to exit from cut copy mode
End Sub

'Note:
'   Disable Screen Updating is used to stop screen flickering
'   Disable Events is used to avoid interrupted dialog boxes or popups.
Application.ScreenUpdating = False

'   Display Alerts is used to stop pop-ups while deleting Worksheet
 Application.DisplayAlerts = False