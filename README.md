<div align="center">

## Determine Prime and return Factors


</div>

### Description

Determine if a number is Prime if it is not prime return Factors of the number
 
### More Info
 
Long number

If not prime returns factors otherwise returns that the number is Prime

None Known


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Thomas D\. Tomlins](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/thomas-d-tomlins.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/thomas-d-tomlins-determine-prime-and-return-factors__1-7720/archive/master.zip)





### Source Code

```
Option Explicit
'Add two textboxes 1- txtFactors(Returns Factors) and 1- Text1(Input number
'Two commandbuttons 1-cmdPrimeand 1-cmdPrime2
'One Label 1-Label2
Private Sub cmdPrime_Click()
  Dim I As Long, J As Long, Num As Long
  Num = Val(Text1.Text)
  If Num <= 3 Then
    Label2.Caption = "Entry is Prime"
    Exit Sub
  End If
  If Num Mod 2 = 0 Then
    Label2.Caption = "Entry is Not Prime"
    Exit Sub
  End If
    I = Int(Sqr(Num))  ' Should be Sqrt(Num)
    For J = 3 To I Step 2
     If Num Mod J = 0 Then
        Label2.Caption = "Entry is Not Prime"
        Exit Sub
     End If
    Next J
  Label2.Caption = "Entry is Prime"
End Sub
Private Sub cmdPrime2_Click()
  Dim Factors As New Collection
  Dim I As Long, J As Long, K As Long, L As Long, Num As Long
  Num = Val(Text1.Text)
    I = Int(Sqr(Num))  ' Should be Sqrt(Num)
    For J = 2 To I
     If Num Mod J = 0 Then
        L = Factors.Count \ 2
        K = Num \ J
        If Factors.Count > 0 Then
        Factors.Add J, , , L
        If (K <> J) Then Factors.Add K, , , L + 1
        Else
        Factors.Add J
        If (K <> J) Then Factors.Add K
        End If
     End If
    Next J
    If Factors.Count = 0 Then
     txtFactors.Text = Text1.Text & " is prime."
    Else
     txtFactors.Text = Text1.Text & " is not prime." & vbCrLf
     txtFactors.Text = txtFactors.Text & "It is divisible by "
    For I = 1 To Factors.Count
    txtFactors.Text = txtFactors.Text & Factors.Item(I) & " ,"
    Next I
    End If
End Sub
```

