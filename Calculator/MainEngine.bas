Attribute VB_Name = "MainEngine"

Function Get_Single_Vaule(Vaule As String) As String  ' «·Õ’Ê· ⁄·Ï «·ﬁÌ„… «·⁄œœÌ… ·Õﬁ· ÊÕÌœ À«‰ÊÌ
On Error Resume Next
Dim InStrPos As Integer
Dim LastVaule As String
Dim LastSub As String ' ¬Œ— ⁄„·Ì… Õ”«»Ì…
Dim LastSub2 As String ' ¬Œ— √„— Õ”«»Ì ( Ã–— , .. )
Dim ThereStr As Boolean ' Â·  „ «œŒ«· ﬁÌ„… ‰’Ì… ø «–« ﬂ«‰ ‰⁄„ ÌÃ» «÷«›… «·√⁄œ«œ «·ﬁ«œ„… ﬂ‰’Ê’
Dim FramePos() As String

FramePos() = Split(Vaule, " ")
LastSub = "+"
LastSub2 = ""
LastVaule = ""
ThereStr = False

'MsgBox "-" + Vaule + "-"

For InStrPos = 0 To Num_Of_Space(Vaule)

If FramePos(InStrPos) = "+" Then
LastSub = "+"
GoTo 1
End If

If FramePos(InStrPos) = "-" Then
LastSub = "-"
GoTo 1
End If

If FramePos(InStrPos) = "*" Then
LastSub = "*"
GoTo 1
End If

If FramePos(InStrPos) = "\" Then
LastSub = "\"
GoTo 1
End If

If FramePos(InStrPos) = "/" Then
LastSub = "/"
GoTo 1
End If

If FramePos(InStrPos) = "^" Then
LastSub = "^"
GoTo 1
End If


If LastSub2 = "" Then  ' ·« ÌÊÃœ √„— Õ”«»Ì ”«»ﬁ«

'  Êﬁ⁄ ÕœÊÀ √„— Õ”«»Ì

If FramePos(InStrPos) = TxTi("#Sqr#") Then
LastSub2 = "Ã–—"
GoTo 1
End If

If FramePos(InStrPos) = TxTi("#Sin#") Then
LastSub2 = "Ã»"
GoTo 1
End If

If FramePos(InStrPos) = TxTi("#Cos#") Then
LastSub2 = " Ã»"
GoTo 1
End If

If FramePos(InStrPos) = TxTi("#Tan#") Then
LastSub2 = "Ÿ·"
GoTo 1
End If

If FramePos(InStrPos) = TxTi("#Ctg#") Then
LastSub2 = " Ÿ·"
GoTo 1
End If

If FramePos(InStrPos) = TxTi("#Rad#") Then
LastSub2 = "—«œ"
GoTo 1
End If

If FramePos(InStrPos) = TxTi("#!#") Then
LastSub2 = "!"
GoTo 1
End If

If FramePos(InStrPos) = TxTi("#|#") Then
LastSub2 = "„ÿ·ﬁ…"
GoTo 1
End If

' ·Õœ «·¬‰ , ·« ÌÊÃœ √„— Õ”«»Ì , ·⁄· «·√„— ÌﬂÊ‰ ÂÊ »Ì / «·À«»  »Ì
If FramePos(InStrPos) = TxTi("#Pi#") Then
LastSub2 = "»Ì"
InStrPos = InStrPos - 1
GoTo 1
End If

' ·Õœ «·¬‰ , ·« ÌÊÃœ √„— Õ”«»Ì , ·⁄· «·√„— ÌﬂÊ‰ ÂÊ »Ì / «·À«»  »Ì
If FramePos(InStrPos) = TxTi("#Mg#") Then
LastSub2 = "Mg"
InStrPos = InStrPos - 1
GoTo 1
End If

' ·Õœ «·¬‰ , ·« ÌÊÃœ √„— Õ”«»Ì , ·⁄· «·√„— ÌﬂÊ‰ ÂÊ »·«‰ﬂ / À«»  »·«‰ﬂ
If FramePos(InStrPos) = TxTi("#Plank#") Then
LastSub2 = "»·«‰ﬂ"
InStrPos = InStrPos - 1
GoTo 1
End If

' ·Õœ «·¬‰ , ·« ÌÊÃœ √„— Õ”«»Ì , Ê·‰ ÌÊÃœ , ›«·√„— ÂÊ ﬁÌ„… ⁄œœÌ… ⁄«œÌ…

If ThereStr = False Then  ' «·Õﬁ· ·« ÌÕ„· ﬁÌ„ ‰’Ì…

If LastVaule = "" Then  ' Ã⁄· «·ﬁÌ„… «·»œ«∆Ì… ’›— ( · ÂÌÀ…  ﬁ»· «·ﬁÌ„ «·⁄œœÌ… )
LastVaule = "0"
End If
' «Ã—«¡ «ÕœÏ «·⁄„·Ì«  «·»”Ìÿ… »«⁄ »«œÂ« ﬁÌ„ ⁄œœÌ…
Select Case LastSub
Case Is = "+": LastVaule = LastVaule + CSng(FramePos(InStrPos))
Case Is = "-": LastVaule = LastVaule - CSng(FramePos(InStrPos))
Case Is = "*": LastVaule = LastVaule * CSng(FramePos(InStrPos))
Case Is = "/": LastVaule = LastVaule / CSng(FramePos(InStrPos))
Case Is = "\": LastVaule = LastVaule \ CSng(FramePos(InStrPos))
Case Is = "^": LastVaule = LastVaule ^ CSng(FramePos(InStrPos))
End Select




End If






' «‰ ÂÏ «Ã—«¡ «·⁄„·Ì«  «·Õ”«»Ì… «·»”Ìÿ… , ·Õœ «·¬‰ ﬂ«‰ Ì› —÷ ⁄œ„ ÊÃÊœ √„— Õ”«»Ì


Else ' ÌÊÃœ √„— Õ”«»Ì ( Ã–— , ’ÕÌÕ , .. )


If LastSub2 = "Ã–—" Then
If ThereStr = False Then ' «·–«ﬂ—… «·¬‰ ⁄»«—… ⁄‰ ﬁÌ„… ⁄œœÌ…
If LastVaule = "" Then '  ÂÌÀ… «·–«ﬂ—… · ﬁ»· ﬁÌ„ ⁄œœÌ…
LastVaule = "0"
End If
Select Case LastSub
Case Is = "+": LastVaule = LastVaule + Sqr(CSng(FramePos(InStrPos)))
Case Is = "-": LastVaule = LastVaule - Sqr(CSng(FramePos(InStrPos)))
Case Is = "*": LastVaule = LastVaule * Sqr(CSng(FramePos(InStrPos)))
Case Is = "/": LastVaule = LastVaule / Sqr(CSng(FramePos(InStrPos)))
Case Is = "\": LastVaule = LastVaule \ Sqr(CSng(FramePos(InStrPos)))
Case Is = "^": LastVaule = LastVaule ^ Sqr(CSng(FramePos(InStrPos)))
End Select
Else ' «·–«ﬂ—… ‰’Ì… , ÌÃ» «÷«›… «·ﬁÌ„ «·⁄œœÌ… ﬂ‰’
Select Case LastSub
Case Is = "+": LastVaule = LastVaule + Format(Sqr(CSng(FramePos(InStrPos))))
End Select
End If
LastSub2 = ""
End If

If LastSub2 = "Ã»" Then
If ThereStr = False Then ' «·–«ﬂ—… «·¬‰ ⁄»«—… ⁄‰ ﬁÌ„… ⁄œœÌ…
If LastVaule = "" Then '  ÂÌÀ… «·–«ﬂ—… · ﬁ»· ﬁÌ„ ⁄œœÌ…
LastVaule = "0"
End If
Select Case LastSub
Case Is = "+": LastVaule = LastVaule + Sin(CSng(FramePos(InStrPos)))
Case Is = "-": LastVaule = LastVaule - Sin(CSng(FramePos(InStrPos)))
Case Is = "*": LastVaule = LastVaule * Sin(CSng(FramePos(InStrPos)))
Case Is = "/": LastVaule = LastVaule / Sin(CSng(FramePos(InStrPos)))
Case Is = "\": LastVaule = LastVaule \ Sin(CSng(FramePos(InStrPos)))
Case Is = "^": LastVaule = LastVaule ^ Sin(CSng(FramePos(InStrPos)))
End Select
Else ' «·–«ﬂ—… ‰’Ì… , ÌÃ» «÷«›… «·ﬁÌ„ «·⁄œœÌ… ﬂ‰’
Select Case LastSub
Case Is = "+": LastVaule = LastVaule + Format(Sin(CSng(FramePos(InStrPos))))
End Select
End If
LastSub2 = ""
End If

If LastSub2 = " Ã»" Then
If ThereStr = False Then ' «·–«ﬂ—… «·¬‰ ⁄»«—… ⁄‰ ﬁÌ„… ⁄œœÌ…
If LastVaule = "" Then '  ÂÌÀ… «·–«ﬂ—… · ﬁ»· ﬁÌ„ ⁄œœÌ…
LastVaule = "0"
End If
Select Case LastSub
Case Is = "+": LastVaule = LastVaule + Cos(CSng(FramePos(InStrPos)))
Case Is = "-": LastVaule = LastVaule - Cos(CSng(FramePos(InStrPos)))
Case Is = "*": LastVaule = LastVaule * Cos(CSng(FramePos(InStrPos)))
Case Is = "/": LastVaule = LastVaule / Cos(CSng(FramePos(InStrPos)))
Case Is = "\": LastVaule = LastVaule \ Cos(CSng(FramePos(InStrPos)))
Case Is = "^": LastVaule = LastVaule ^ Cos(CSng(FramePos(InStrPos)))
End Select
Else ' «·–«ﬂ—… ‰’Ì… , ÌÃ» «÷«›… «·ﬁÌ„ «·⁄œœÌ… ﬂ‰’
Select Case LastSub
Case Is = "+": LastVaule = LastVaule + Format(Cos(CSng(FramePos(InStrPos))))
End Select
End If
LastSub2 = ""
End If

If LastSub2 = "Ÿ·" Then
If ThereStr = False Then ' «·–«ﬂ—… «·¬‰ ⁄»«—… ⁄‰ ﬁÌ„… ⁄œœÌ…
If LastVaule = "" Then '  ÂÌÀ… «·–«ﬂ—… · ﬁ»· ﬁÌ„ ⁄œœÌ…
LastVaule = "0"
End If
Select Case LastSub
Case Is = "+": LastVaule = LastVaule + Tan(CSng(FramePos(InStrPos)))
Case Is = "-": LastVaule = LastVaule - Tan(CSng(FramePos(InStrPos)))
Case Is = "*": LastVaule = LastVaule * Tan(CSng(FramePos(InStrPos)))
Case Is = "/": LastVaule = LastVaule / Tan(CSng(FramePos(InStrPos)))
Case Is = "\": LastVaule = LastVaule \ Tan(CSng(FramePos(InStrPos)))
Case Is = "^": LastVaule = LastVaule ^ Tan(CSng(FramePos(InStrPos)))
End Select
Else ' «·–«ﬂ—… ‰’Ì… , ÌÃ» «÷«›… «·ﬁÌ„ «·⁄œœÌ… ﬂ‰’
Select Case LastSub
Case Is = "+": LastVaule = LastVaule + Format(Tan(CSng(FramePos(InStrPos))))
End Select
End If
LastSub2 = ""
End If

If LastSub2 = " Ÿ·" Then
If ThereStr = False Then ' «·–«ﬂ—… «·¬‰ ⁄»«—… ⁄‰ ﬁÌ„… ⁄œœÌ…
If LastVaule = "" Then '  ÂÌÀ… «·–«ﬂ—… · ﬁ»· ﬁÌ„ ⁄œœÌ…
LastVaule = "0"
End If
Select Case LastSub
Case Is = "+": LastVaule = LastVaule + (1 / Tan(CSng(FramePos(InStrPos))))
Case Is = "-": LastVaule = LastVaule - (1 / Tan(CSng(FramePos(InStrPos))))
Case Is = "*": LastVaule = LastVaule * (1 / Tan(CSng(FramePos(InStrPos))))
Case Is = "/": LastVaule = LastVaule / (1 / Tan(CSng(FramePos(InStrPos))))
Case Is = "\": LastVaule = LastVaule \ (1 / Tan(CSng(FramePos(InStrPos))))
Case Is = "^": LastVaule = LastVaule ^ (1 / Tan(CSng(FramePos(InStrPos))))
End Select
Else ' «·–«ﬂ—… ‰’Ì… , ÌÃ» «÷«›… «·ﬁÌ„ «·⁄œœÌ… ﬂ‰’
Select Case LastSub
Case Is = "+": LastVaule = LastVaule + Format((1 / Tan(CSng(FramePos(InStrPos)))))
End Select
End If
LastSub2 = ""
End If

If LastSub2 = "—«œ" Then
If ThereStr = False Then ' «·–«ﬂ—… «·¬‰ ⁄»«—… ⁄‰ ﬁÌ„… ⁄œœÌ…
If LastVaule = "" Then '  ÂÌÀ… «·–«ﬂ—… · ﬁ»· ﬁÌ„ ⁄œœÌ…
LastVaule = "0"
End If
Select Case LastSub
Case Is = "+": LastVaule = LastVaule + (CSng(FramePos(InStrPos)) * 3.141593 / 180)
Case Is = "-": LastVaule = LastVaule - (CSng(FramePos(InStrPos)) * 3.141593 / 180)
Case Is = "*": LastVaule = LastVaule * (CSng(FramePos(InStrPos)) * 3.141593 / 180)
Case Is = "/": LastVaule = LastVaule / (CSng(FramePos(InStrPos)) * 3.141593 / 180)
Case Is = "\": LastVaule = LastVaule \ (CSng(FramePos(InStrPos)) * 3.141593 / 180)
Case Is = "^": LastVaule = LastVaule ^ (CSng(FramePos(InStrPos)) * 3.141593 / 180)
End Select
Else ' «·–«ﬂ—… ‰’Ì… , ÌÃ» «÷«›… «·ﬁÌ„ «·⁄œœÌ… ﬂ‰’
Select Case LastSub
Case Is = "+": LastVaule = LastVaule + Format(CSng(FramePos(InStrPos)) * 3.141593 / 180)
End Select
End If
LastSub2 = ""
End If

If LastSub2 = "!" Then
If ThereStr = False Then ' «·–«ﬂ—… «·¬‰ ⁄»«—… ⁄‰ ﬁÌ„… ⁄œœÌ…
If LastVaule = "" Then '  ÂÌÀ… «·–«ﬂ—… · ﬁ»· ﬁÌ„ ⁄œœÌ…
LastVaule = "0"
End If
Select Case LastSub
Case Is = "+": LastVaule = LastVaule + Get_Amly_Vaule(CSng(FramePos(InStrPos)))
Case Is = "-": LastVaule = LastVaule - Get_Amly_Vaule(CSng(FramePos(InStrPos)))
Case Is = "*": LastVaule = LastVaule * Get_Amly_Vaule(CSng(FramePos(InStrPos)))
Case Is = "/": LastVaule = LastVaule / Get_Amly_Vaule(CSng(FramePos(InStrPos)))
Case Is = "\": LastVaule = LastVaule \ Get_Amly_Vaule(CSng(FramePos(InStrPos)))
Case Is = "^": LastVaule = LastVaule ^ Get_Amly_Vaule(CSng(FramePos(InStrPos)))
End Select
Else ' «·–«ﬂ—… ‰’Ì… , ÌÃ» «÷«›… «·ﬁÌ„ «·⁄œœÌ… ﬂ‰’
Select Case LastSub
Case Is = "+": LastVaule = LastVaule + Format(Get_Amly_Vaule(CSng(FramePos(InStrPos))))
End Select
End If
LastSub2 = ""
End If

If LastSub2 = "»Ì" Then
If ThereStr = False Then ' «·–«ﬂ—… «·¬‰ ⁄»«—… ⁄‰ ﬁÌ„… ⁄œœÌ…
If LastVaule = "" Then ' Õﬁ· ›«—€
LastVaule = "0"
End If
Select Case LastSub
Case Is = "+": LastVaule = LastVaule + 3.141593
Case Is = "-": LastVaule = LastVaule - 3.141593
Case Is = "*": LastVaule = LastVaule * 3.141593
Case Is = "/": LastVaule = LastVaule / 3.141593
Case Is = "\": LastVaule = LastVaule \ 3.141593
Case Is = "^": LastVaule = LastVaule ^ 3.141593
End Select
Else
Select Case LastSub
Case Is = "+": LastVaule = LastVaule + Format(3.141593)
End Select
End If
LastSub2 = ""
End If

If LastSub2 = "Mg" Then
If ThereStr = False Then ' «·–«ﬂ—… «·¬‰ ⁄»«—… ⁄‰ ﬁÌ„… ⁄œœÌ…
If LastVaule = "" Then ' Õﬁ· ›«—€
LastVaule = "0"
End If
Select Case LastSub
Case Is = "+": LastVaule = LastVaule + 9.81
Case Is = "-": LastVaule = LastVaule - 9.81
Case Is = "*": LastVaule = LastVaule * 9.81
Case Is = "/": LastVaule = LastVaule / 9.81
Case Is = "\": LastVaule = LastVaule \ 9.81
Case Is = "^": LastVaule = LastVaule ^ 9.81
End Select
Else
Select Case LastSub
Case Is = "+": LastVaule = LastVaule + Format(9.81)
End Select
End If
LastSub2 = ""
End If

If LastSub2 = "»·«‰ﬂ" Then
If ThereStr = False Then ' «·–«ﬂ—… «·¬‰ ⁄»«—… ⁄‰ ﬁÌ„… ⁄œœÌ…
If LastVaule = "" Then ' Õﬁ· ›«—€
LastVaule = "0"
End If
Select Case LastSub
Case Is = "+": LastVaule = LastVaule + (6.625 * (10 ^ (-34)))
Case Is = "-": LastVaule = LastVaule - (6.625 * (10 ^ (-34)))
Case Is = "*": LastVaule = LastVaule * (6.625 * (10 ^ (-34)))
Case Is = "/": LastVaule = LastVaule / (6.625 * (10 ^ (-34)))
Case Is = "\": LastVaule = LastVaule \ (6.625 * (10 ^ (-34)))
Case Is = "^": LastVaule = LastVaule ^ (6.625 * (10 ^ (-34)))
End Select
Else
Select Case LastSub
Case Is = "+": LastVaule = LastVaule + Format((6.625 * (10 ^ (-34))))
End Select
End If
LastSub2 = ""
End If

If LastSub2 = "„ÿ·ﬁ…" Then
Select Case LastSub
Case Is = "+"
For I4 = InStrPos To Num_Of_Space(Vaule)  ' ”Õ» «·ﬁÌ„ „‰ »œ«Ì… —„“ «·„ÿ·ﬁ… Õ Ï ‰Ãœ —„“ «·ﬁÌ„… «·„ÿ·ﬁ… „‰ ÃœÌœ
If Not FramePos(I4) = TxTi("#|#") Then
X2 = X2 + FramePos(I4) + " "
Else
X2 = Left(X2, Len(X2) - 1)
X2 = Get_Vaule("( " + X2 + " )")
If ThereStr = False Then ' «·–«ﬂ—… «·¬‰ ⁄»«—… ⁄‰ ﬁÌ„… ⁄œœÌ…
If LastVaule = "" Then LastVaule = "0"
If CSng(X2) > 0 Then
LastVaule = LastVaule + CSng(Get_Vaule("( " + X2 + " )"))
Else
LastVaule = LastVaule - CSng(Get_Vaule("( " + X2 + " )"))
End If
Else
If CSng(X2) > 0 Then
LastVaule = LastVaule + Get_Vaule("( " + X2 + " )")
Else
LastVaule = LastVaule + Format(0 - CSng(Get_Vaule("( " + X2 + " )")))
End If
End If
InStrPos = I4
LastSub2 = ""
GoTo 1
End If
Next

Case Is = "-"
For I4 = InStrPos To Num_Of_Space(Vaule)  ' ”Õ» «·ﬁÌ„ „‰ »œ«Ì… —„“ «·„ÿ·ﬁ… Õ Ï ‰Ãœ —„“ «·ﬁÌ„… «·„ÿ·ﬁ… „‰ ÃœÌœ
If Not FramePos(I4) = TxTi("#|#") Then
X2 = X2 + FramePos(I4) + " "
Else
X2 = Left(X2, Len(X2) - 1)
X2 = Get_Vaule("( " + X2 + " )")
If ThereStr = False Then ' «·–«ﬂ—… «·¬‰ ⁄»«—… ⁄‰ ﬁÌ„… ⁄œœÌ…
If LastVaule = "" Then LastVaule = "0"
If CSng(X2) < 0 Then
LastVaule = LastVaule + CSng(Get_Vaule("( " + X2 + " )"))
Else
LastVaule = LastVaule - CSng(Get_Vaule("( " + X2 + " )"))
End If
Else
If CSng(X2) < 0 Then
LastVaule = LastVaule + Get_Vaule("( " + X2 + " )")
Else
LastVaule = LastVaule + Format(0 - CSng(Get_Vaule("( " + X2 + " )")))
End If
End If
InStrPos = I4
LastSub2 = ""
GoTo 1
End If
Next
End Select
End If


End If

1:
Next

If Len(LastVaule) > 1 Then
If Left(LastVaule, 2) = "0-" Then
LastVaule = Right(LastVaule, Len(LastVaule) - 1)
End If
End If

Get_Single_Vaule = Format(LastVaule)
End Function


Function Get_Vaule(Vaule As String) As String ' «·Õ’Ê· ⁄·Ï «·ﬁÌ„„… «·⁄œœÌ… ·Õﬁ· —∆Ì”Ì „«
Dim InStrPos As Integer
Dim StartPos As Integer
Dim EndPos As Integer

Vaule = CmdAddSpaces(Vaule)
Vaule = CmdRemoveSpaces(Vaule)
Vaule = "( " + Vaule + " )"

Do While InStr(1, Vaule, "(", vbTextCompare) > 0
DoEvents
InStrPos = InStr(1, Vaule, "(", vbTextCompare)

StartPos = InStrPos

For EndPos = StartPos To Len(Vaule)

X = Mid(Vaule, EndPos, 1)
If X = "(" Then
StartPos = EndPos
End If

If X = ")" Then
Dim OO As String

OO = Get_Single_Vaule(Trim(Mid(Vaule, StartPos + 1, EndPos - StartPos - 1)))

Vaule = Left(Vaule, StartPos - 1) + Format(OO) + Right(Vaule, Len(Vaule) - EndPos)
Get_Vaule = Vaule
GoTo 1
End If

Next

1:
Loop
End Function



Function Num_Of_Space(Vaule As String) As Integer ' «Ã—«¡ „”«⁄œ ||| „⁄—›… ⁄œœ «·›—«€«  ›Ì ”ÿ— „«
Dim NumOfSpace As Integer

Vaule = Trim(Vaule)

For I = 1 To Len(Vaule)

X = Mid(Vaule, I, 1)

If X = " " Then
NumOfSpace = NumOfSpace + 1
End If

Next

Num_Of_Space = NumOfSpace
End Function

Function Get_Amly_Vaule(Vaule As Integer) As Integer    ' «·Õ’Ê· ⁄·Ï ﬁÌ„… «·⁄«„·Ì ·⁄œœ „«
Dim Xamly As Integer
Dim Yamly As Integer
Yamly = 1
Xamly = Vaule
Do While Xamly > 1
DoEvents
Yamly = Yamly * Xamly
Xamly = Xamly - 1
Loop
Get_Amly_Vaule = Yamly
End Function

Function TxTi(Text As String) As String

TxTi = Mid(Text, 2, Len(Text) - 2)

End Function

Function CmdRemoveSpaces(MainCommand As String) As String '  ‰”Ìﬁ «·ﬂÊœ - „”Õ ›—«€«  “«∆œ…
Dim NewCommand As String
Dim NextSpace As String
Dim BackSpace As String
Dim PosText As String
Dim InStrTxt As Boolean

MainCommand = Trim(MainCommand)

BackSpace = " "
InStrTxt = False

For I = 1 To Len(MainCommand)
PosText = Mid(MainCommand + " ", I, 1)
NextSpace = Mid(MainCommand + "  ", I + 1, 1)

If PosText = TxTi("#|#") Then InStrTxt = Not InStrTxt

If InStrTxt = False And PosText = " " And BackSpace = " " Then GoTo 1

NewCommand = NewCommand + PosText
1:

BackSpace = PosText
Next

CmdRemoveSpaces = NewCommand
End Function


Function CmdAddSpaces(MainCommand As String) As String '  ‰”Ìﬁ «·ﬂÊœ - «÷«›… ›—«€«  ‰«ﬁ’…
Dim NewCommand As String
Dim NextSpace As String
Dim BackSpace As String
Dim PosText As String
Dim InStrTxt As Boolean

BackSpace = ""
InStrTxt = False

For I = 1 To Len(MainCommand)
PosText = Mid(MainCommand + " ", I, 1)
NextSpace = Mid(MainCommand + "  ", I + 1, 1)

If PosText = TxTi("#|#") And Not BackSpace = " " And Not NextSpace = " " Then
NewCommand = NewCommand + " " + PosText + " "
GoTo 1
End If
If PosText = TxTi("#|#") And Not BackSpace = " " Then
NewCommand = NewCommand + " " + PosText
GoTo 1
End If
If PosText = TxTi("#|#") And Not NextSpace = " " Then
NewCommand = NewCommand + PosText + " "
GoTo 1
End If

If PosText = "+" And Not BackSpace = " " And Not NextSpace = " " Then
NewCommand = NewCommand + " " + PosText + " "
GoTo 1
End If
If PosText = "+" And Not BackSpace = " " Then
NewCommand = NewCommand + " " + PosText
GoTo 1
End If
If PosText = "+" And Not NextSpace = " " Then
NewCommand = NewCommand + PosText + " "
GoTo 1
End If

If PosText = "-" And Not BackSpace = " " And Not NextSpace = " " Then
NewCommand = NewCommand + " " + PosText + " "
GoTo 1
End If
If PosText = "-" And Not BackSpace = " " Then
NewCommand = NewCommand + " " + PosText
GoTo 1
End If
If PosText = "-" And Not NextSpace = " " Then
NewCommand = NewCommand + PosText + " "
GoTo 1
End If

If PosText = "*" And Not BackSpace = " " And Not NextSpace = " " Then
NewCommand = NewCommand + " " + PosText + " "
GoTo 1
End If
If PosText = "*" And Not BackSpace = " " Then
NewCommand = NewCommand + " " + PosText
GoTo 1
End If
If PosText = "*" And Not NextSpace = " " Then
NewCommand = NewCommand + PosText + " "
GoTo 1
End If

If PosText = "\" And Not BackSpace = " " And Not NextSpace = " " Then
NewCommand = NewCommand + " " + PosText + " "
GoTo 1
End If
If PosText = "\" And Not BackSpace = " " Then
NewCommand = NewCommand + " " + PosText
GoTo 1
End If
If PosText = "\" And Not NextSpace = " " Then
NewCommand = NewCommand + PosText + " "
GoTo 1
End If

If PosText = "/" And Not BackSpace = " " And Not NextSpace = " " Then
NewCommand = NewCommand + " " + PosText + " "
GoTo 1
End If
If PosText = "/" And Not BackSpace = " " Then
NewCommand = NewCommand + " " + PosText
GoTo 1
End If
If PosText = "/" And Not NextSpace = " " Then
NewCommand = NewCommand + PosText + " "
GoTo 1
End If

If PosText = "^" And Not BackSpace = " " And Not NextSpace = " " Then
NewCommand = NewCommand + " " + PosText + " "
GoTo 1
End If
If PosText = "^" And Not BackSpace = " " Then
NewCommand = NewCommand + " " + PosText
GoTo 1
End If
If PosText = "^" And Not NextSpace = " " Then
NewCommand = NewCommand + PosText + " "
GoTo 1
End If

If PosText = "=" And Not BackSpace = " " And Not NextSpace = " " Then
NewCommand = NewCommand + " " + PosText + " "
GoTo 1
End If
If PosText = "=" And Not BackSpace = " " Then
NewCommand = NewCommand + " " + PosText
GoTo 1
End If
If PosText = "=" And Not NextSpace = " " Then
NewCommand = NewCommand + PosText + " "
GoTo 1
End If

If PosText = ">" And Not BackSpace = " " And Not NextSpace = " " Then
NewCommand = NewCommand + " " + PosText + " "
GoTo 1
End If
If PosText = ">" And Not BackSpace = " " Then
NewCommand = NewCommand + " " + PosText
GoTo 1
End If
If PosText = ">" And Not NextSpace = " " Then
NewCommand = NewCommand + PosText + " "
GoTo 1
End If

If PosText = "<" And Not BackSpace = " " And Not NextSpace = " " Then
NewCommand = NewCommand + " " + PosText + " "
GoTo 1
End If
If PosText = "<" And Not BackSpace = " " Then
NewCommand = NewCommand + " " + PosText
GoTo 1
End If
If PosText = "<" And Not NextSpace = " " Then
NewCommand = NewCommand + PosText + " "
GoTo 1
End If

If PosText = "(" And Not BackSpace = " " And Not NextSpace = " " Then
NewCommand = NewCommand + " " + PosText + " "
GoTo 1
End If
If PosText = "(" And Not BackSpace = " " Then
NewCommand = NewCommand + " " + PosText
GoTo 1
End If
If PosText = "(" And Not NextSpace = " " Then
NewCommand = NewCommand + PosText + " "
GoTo 1
End If

If PosText = ")" And Not BackSpace = " " And Not NextSpace = " " Then
NewCommand = NewCommand + " " + PosText + " "
GoTo 1
End If
If PosText = ")" And Not BackSpace = " " Then
NewCommand = NewCommand + " " + PosText
GoTo 1
End If
If PosText = ")" And Not NextSpace = " " Then
NewCommand = NewCommand + PosText + " "
GoTo 1
End If

If PosText = "|" And Not BackSpace = " " And Not NextSpace = " " Then
NewCommand = NewCommand + " " + PosText + " "
GoTo 1
End If
If PosText = "|" And Not BackSpace = " " Then
NewCommand = NewCommand + " " + PosText
GoTo 1
End If
If PosText = "|" And Not NextSpace = " " Then
NewCommand = NewCommand + PosText + " "
GoTo 1
End If


NewCommand = NewCommand + PosText
1:

BackSpace = PosText
Next

CmdAddSpaces = NewCommand
End Function

Function ChangeA2B(MainValue As String, WantToChange As String, ChangeTo As String, LoopNum As Integer) As String ' «” »œ«· «Õ —«›Ì
On Error GoTo 1
Dim S1 As String ' «·‰’ «·–Ì ‰—Ìœ «” »œ«·Â
Dim S2 As String ' «·‰’ «·ÃœÌœ
Dim S3 As String ' «·‰’ «·–Ì ‰” »œ· ›ÌÂ
Dim N As Integer ' ‰ﬁÿ…  ÕœÌœ

S1 = WantToChange
S2 = ChangeTo
S3 = MainValue

For I = 0 To LoopNum - 1

N = 0
Do While InStr(N + 1, S3, S1) > 0
DoEvents
N = InStr(S3, S1)
S3 = Left(S3, N - 1) + S2 + Right(S3, Len(S3) - (N - 1) - Len(S1))
N = N + Len(S2) - Len(S1)
Loop
1:
Next

ChangeA2B = S3

End Function
