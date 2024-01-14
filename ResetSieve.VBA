' This sub routine clears the coloured status of the Sieve grids

Sub ResetSieve()


    Dim CellsDown As Long, CellsAcross As Integer
    Dim CurrRow As Long, CurrCol As Integer, RowScan As Integer
    Dim Factor As Integer
    Dim CountMultiple As Integer
    Dim StartTime As Double
  
     Response = Application.InputBox(prompt:="Do you wish to reset the WHOLE Sieve ?(Reply: Y for Yes and N for No)", Type:=2)
    
    'MsgBox Response
     
     CellsDown = 100
     CellsAcross = 10
     Factor = Sheet1.Range("m4").Value

     
'   Record starting time
    StartTime = Timer

             
    If Response <> "Y" And Response <> "y" And Response <> "N" And Response <> "n" Then
       MsgBox "Invalid Response !!"
        'GothruSeed = 0
    Else
        If Response = "Y" Or Response = "y" Then
        Confirmation = True
       ' GothruSeed = 1
    Else
        If Response = "N" Or Response = "n" Then
        Confirmation = False
       ' GothruSeed = 1
    End If
    End If
    End If
    
    'Application.ScreenUpdating = Confirmation
         
    If Confirmation Then
              
    For CurrRow = 1 To CellsDown
        For CurrCol = 1 To CellsAcross
            
                             
                Sheet1.Range("a1").Offset(CurrRow - 1, CurrCol - 1).Interior.ColorIndex = 0
               
            
        Next CurrCol
     Next CurrRow

    
    End If
    
    MsgBox Format(Timer - StartTime, "00.00") & " seconds"
    

End Sub

'This Algorithm is written by Mr. Jason Ng Yee Boon