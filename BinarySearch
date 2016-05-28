Option Explicit

Sub BinarySearch()

'Specify some array of numbers
    Dim intThousand(1000) As Integer
'Need variables to demarcate what the function will search against

    Dim i As Integer
    Dim intTop As Integer
    Dim intMiddle As Integer
    Dim intBottom As Integer
    Dim Usernumb As Variant

    For i = 1 To 1000
      intThousand(i) = i
    Next i
      
      Usernumb = InputBox("Specify a whole number between 1 and 1000: ")
      intTop = UBound(intThousand)
      intBottom = LBound(intThousand)

'Several ways to do this besides while loop
    Do
      intMiddle = (intTop + intBottom) / 2
        If Usernumbe > intThousand(intMiddle) Then
          intBottom = intMiddle + 1
        Else
          intTop = intMiddle - 1
        End If
    Loop Until (Usernumb = intThousand(intMiddle)) _
      Or (intBottom > intTop)

    If Usernumb = intThousand(intMiddle) Then
      Debug.Print Usernumb & ", at position " & intMiddle 
    Else
      Debug.Print "Number suggested was not in Array"
    End If
  End Sub
