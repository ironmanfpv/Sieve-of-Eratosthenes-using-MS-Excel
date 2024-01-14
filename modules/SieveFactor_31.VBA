Sub SeiveFactor31()

    Dim CellsDown As Long, CellsAcross As Integer
    Dim CurrRow As Long, CurrCol As Integer, RowScan As Integer
    Dim Factor As Integer
    Dim CountMultiple As Integer
    Dim StartTime As Double
    Dim Response As String
    Dim Confirmation As Boolean
    Dim GothruSeed As Integer
   
     Response = Application.InputBox(prompt:="Do you wish to identify the multiples now ?(Reply: Y for Yes and N for No)", Type:=2)
    
    'MsgBox Response
     
     CellsDown = Sheet1.Range("m1").Value
     CellsAcross = Sheet1.Range("o1").Value
     Factor = Sheet1.Range("m14").Value

     
'   Record starting time
    StartTime = Timer


             
    If Response <> "Y" And Response <> "y" And Response <> "N" And Response <> "n" Then
       MsgBox "Invalid Response !!"
       GothruSeed = 0
    Else
        If Response = "Y" Or Response = "y" Then
        Confirmation = True
        GothruSeed = 1
    Else
        If Response = "N" Or Response = "n" Then
        Confirmation = False
        GothruSeed = 1
    End If
    End If
    End If
    
    Application.ScreenUpdating = Confirmation
         
    If GothruSeed = 1 Then
         
    For CurrRow = 1 To CellsDown
        For CurrCol = 1 To CellsAcross

 

            If Sheet1.Range("a1").Offset(CurrRow - 1, CurrCol - 1).Value Mod Factor = 0 Then
                
                Sheet1.Range("a1").Offset(CurrRow - 1, CurrCol - 1).Interior.ColorIndex = 13
                CountMultiple = CountMultiple + 1
                
           
            End If
            
        Next CurrCol
     Next CurrRow

    
    Sheet1.Range("u14") = CountMultiple
    
    End If
    
    MsgBox Format(Timer - StartTime, "00.00") & " seconds"
    

End Sub


'This Algorithm is written by Mr. Jason Ng Yee Boon







