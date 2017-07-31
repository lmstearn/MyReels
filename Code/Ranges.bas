Attribute VB_Name = "Ranges"
Private S1 As Long, S2 As Long
Public Function sumtype(prizecount As Long, Optional reelend As Long, Optional singl1 As Long, Optional singl2 As Long, Optional BLOCK111 As Boolean = False)
If intRLloop = 2 Then

'Swap direction
singl1 = 6 - singl1
singl2 = 6 - singl2

If multctstat = True Then
VOGstepdir = -1
loopLRorany1 = 6 - loopLRorany1
loopLRorany2 = 6 - loopLRorany2
loopLRorany3 = 6 - loopLRorany3
loopLRorany4 = 6 - loopLRorany4
loopLRorany5 = 6 - loopLRorany5
loopLRorany6 = 6 - loopLRorany6
loopLRorany7 = 6 - loopLRorany7
loopLRorany8 = 6 - loopLRorany8

Else    'vogstepdir 1
temp = loopLRorany1
loopLRorany1 = 6 - loopLRorany2
loopLRorany2 = 6 - temp
temp = loopLRorany3
loopLRorany3 = 6 - loopLRorany4
loopLRorany4 = 6 - temp
temp = loopLRorany5
loopLRorany5 = 6 - loopLRorany6
loopLRorany6 = 6 - temp
temp = loopLRorany7
loopLRorany7 = 6 - loopLRorany8
loopLRorany8 = 6 - temp
End If

reelend = 1
intRLloop = 1   'Not good here but ....
Exit Function
End If
'Set defaults
loopLRorany1 = 0
loopLRorany2 = 0
loopLRorany3 = 0
loopLRorany4 = 0
loopLRorany5 = 0
loopLRorany6 = 0
loopLRorany7 = 0
loopLRorany8 = 0

reelend = 5
VOGstepdir = 1 'No confusion with partdir in partsums
multctstat = False  'No confusion with multiptstat in partsums
sumtype = False


Select Case prizecount


Case 0 '4 0 loopLRorany1-2 map the no-prize X
If sst(mainp, 5) = 1 Then
loopLRorany1 = 1
loopLRorany2 = 5
Else
loopLRorany1 = 5
loopLRorany2 = 5
If sst(mainp, 3) = 1 Then intRLloop = 2
End If

Case 1 '4 1 loopLRorany1-2 map single prize
If sst(mainp, 5) = 1 And sst(testp, 5) = 1 Then
loopLRorany1 = 1
loopLRorany2 = 5
sumtype = True
Exit Function
ElseIf sst(mainp, 1) = 1 And sst(testp, 1) = 1 And sst(mainp, 3) = 0 And sst(testp, 3) = 0 Then Exit Function
ElseIf sst(testp, 1) = 1 Then
    If sst(mainp, 3) = 0 And sst(mainp, 5) = 0 Then
    loopLRorany1 = 5
    loopLRorany2 = 5
    Else
    loopLRorany1 = 1
    loopLRorany2 = 1
    End If
Else 'sst(mainp, 1) = 1 Then
    loopLRorany1 = 5
    loopLRorany2 = 5
End If

If sst(mainp, 3) + sst(mainp, 5) + sst(testp, 3) + sst(testp, 5) >= 2 Then intRLloop = 2 '1 pass




Case 2 '3 2 loopLRorany1-4 maps the double prize
If sst(mainp, 5) = 1 And sst(testp, 5) = 1 Then
loopLRorany1 = 1
loopLRorany2 = 4
loopLRorany3 = 0 'loopfix later
loopLRorany4 = 0
multctstat = True
sumtype = True
Exit Function
ElseIf sst(mainp, 1) = 1 And sst(testp, 1) = 1 And sst(mainp, 3) = 0 And sst(testp, 3) = 0 Then Exit Function
ElseIf sst(testp, 1) = 1 Then
    If sst(mainp, 3) = 0 And sst(mainp, 5) = 0 Then
    loopLRorany1 = 4
    loopLRorany2 = 4
    loopLRorany3 = 5
    loopLRorany4 = 5
    Else
    loopLRorany1 = 2
    loopLRorany2 = 2
    loopLRorany3 = 1
    loopLRorany4 = 1
    End If
Else 'sst(testp, 1) = 0 Then
    loopLRorany1 = 4
    loopLRorany2 = 4
    loopLRorany3 = 5
    loopLRorany4 = 5
End If

If sst(mainp, 3) + sst(mainp, 5) + sst(testp, 3) + sst(testp, 5) >= 2 Then intRLloop = 2


Case 3 '3 1 1 loopLRorany1-4 map single prizes
If sst(mainp, 1) = 1 And sst(testp, 1) = 1 And sst(testp1, 1) = 1 Then Exit Function
If sst(mainp, 5) = 1 And sst(testp, 5) = 1 And sst(testp1, 5) = 1 Then
loopLRorany1 = 1
loopLRorany2 = 5
loopLRorany3 = 1
loopLRorany4 = 5
sumtype = True
Exit Function
ElseIf sst(mainp, 1) = 1 And sst(testp, 1) = 1 And sst(mainp, 3) = 0 And sst(testp, 3) = 0 Then Exit Function
ElseIf sst(mainp, 1) = 1 And sst(testp1, 1) = 1 And sst(mainp, 3) = 0 And sst(testp1, 3) = 0 Then Exit Function
ElseIf sst(testp, 1) = 1 And sst(testp1, 1) = 1 And sst(testp, 3) = 0 And sst(testp1, 3) = 0 Then Exit Function
ElseIf sst(mainp, 1) = 1 Then
If sst(testp, 5) = 1 And sst(testp1, 5) = 1 Then
loopLRorany1 = 4
loopLRorany2 = 5
loopLRorany3 = 4
loopLRorany4 = 5
ElseIf sst(testp1, 3) = 1 Then 'testp is "any" by deduction
loopLRorany1 = 4
loopLRorany2 = 4
loopLRorany3 = 5
loopLRorany4 = 5
ElseIf sst(testp1, 1) = 1 Then
loopLRorany1 = 2
loopLRorany2 = 2
loopLRorany3 = 1
loopLRorany4 = 1
ElseIf sst(testp, 3) = 1 Then
loopLRorany1 = 5
loopLRorany2 = 5
loopLRorany3 = 4
loopLRorany4 = 4
ElseIf sst(testp, 1) = 1 Then
loopLRorany1 = 1
loopLRorany2 = 1
loopLRorany3 = 2
loopLRorany4 = 2
End If

ElseIf sst(testp, 1) = 1 Then 'sst(mainp, 5) = 1
If sst(testp1, 5) = 1 Then
loopLRorany1 = 1
loopLRorany2 = 1
loopLRorany3 = 2
loopLRorany4 = 5
ElseIf sst(testp1, 3) = 1 Then
loopLRorany1 = 1
loopLRorany2 = 1
loopLRorany3 = 5
loopLRorany4 = 5
Else    'test1rl = 0
loopLRorany1 = 5
loopLRorany2 = 5
loopLRorany3 = 1
loopLRorany4 = 1
End If

ElseIf sst(testp1, 1) = 1 Then
loopLRorany1 = 2
loopLRorany2 = 5
loopLRorany3 = 1
loopLRorany4 = 1
End If
If sst(mainp, 3) + sst(mainp, 5) + sst(testp, 3) + sst(testp, 5) + sst(testp1, 3) + sst(testp1, 5) >= 3 Then intRLloop = 2


Case 4 '3 1 X loopLRorany1-2 maps single prize, loopLRorany3-4 maps spare
If sst(mainp, 5) = 1 And sst(testp, 5) = 1 Then
loopLRorany1 = 1
loopLRorany2 = 5
loopLRorany3 = 1
loopLRorany4 = 5
sumtype = True
Exit Function
ElseIf sst(mainp, 1) = 1 And sst(testp, 1) = 1 And sst(mainp, 3) = 0 And sst(testp, 3) = 0 Then Exit Function
ElseIf sst(mainp, 1) = 1 Then
    If sst(testp, 5) = 1 Then
    loopLRorany1 = 4
    loopLRorany2 = 5
    loopLRorany3 = 4
    loopLRorany4 = 5
    Rsgl 4
    ElseIf sst(testp, 3) = 1 Then
    loopLRorany1 = 5
    loopLRorany2 = 5
    loopLRorany3 = 4
    loopLRorany4 = 4
    Rsgl 4, 4
    Else 'testp, 3 = 0 .... mainp,3 = 1 but can't RLswaploop
    loopLRorany1 = 1
    loopLRorany2 = 1
    loopLRorany3 = 2
    loopLRorany4 = 2
    Rsgl 2, 2
    End If
Else 'sst(mainp, 5) = 1
'sst(testp, 1) must be 1
loopLRorany1 = 1
loopLRorany2 = 1
loopLRorany3 = 2
loopLRorany4 = 5
Rsgl 2
End If

If sst(mainp, 3) + sst(mainp, 5) + sst(testp, 3) + sst(testp, 5) >= 2 Then intRLloop = 2


Case 5 '3 X X loopLRorany1-4 maps the spare XX nonprize reels
If sst(mainp, 5) = 1 Then
loopLRorany1 = 1
loopLRorany2 = 4
loopLRorany3 = 0 'loopfix later
loopLRorany4 = 0
multctstat = True
sumtype = True
Exit Function
Else
loopLRorany1 = 4
loopLRorany2 = 4
loopLRorany3 = 5
loopLRorany4 = 5
'All other combos reversed
If sst(mainp, 3) = 1 Then intRLloop = 2
End If


Case 6 '2 2 1 loopLRorany1-4 map out double (testp) prize, loopLRorany5-6 map single
If sst(mainp, 1) = 1 And sst(testp, 1) = 1 And sst(testp1, 1) = 1 Then Exit Function
If sst(mainp, 5) = 1 And sst(testp, 5) = 1 And sst(testp1, 5) = 1 Then
loopLRorany1 = 1
loopLRorany2 = 4
loopLRorany3 = 0
loopLRorany4 = 0
loopLRorany5 = 1
loopLRorany6 = 5
multctstat = True
sumtype = True
Exit Function
ElseIf sst(mainp, 1) = 1 And sst(testp, 1) = 1 And sst(mainp, 3) = 0 And sst(testp, 3) = 0 Then Exit Function
ElseIf sst(mainp, 1) = 1 And sst(testp1, 1) = 1 And sst(mainp, 3) = 0 And sst(testp1, 3) = 0 Then Exit Function
ElseIf sst(testp, 1) = 1 And sst(testp1, 1) = 1 And sst(testp, 3) = 0 And sst(testp1, 3) = 0 Then Exit Function
ElseIf sst(mainp, 1) = 1 Then
If sst(testp, 5) = 1 And sst(testp1, 5) = 1 Then
loopLRorany1 = 3
loopLRorany2 = 4
loopLRorany3 = 0
loopLRorany4 = 0
loopLRorany5 = 3
loopLRorany6 = 5
multctstat = True
ElseIf sst(testp1, 3) = 1 Then
loopLRorany1 = 3
loopLRorany2 = 3
loopLRorany3 = 4
loopLRorany4 = 4
loopLRorany5 = 5
loopLRorany6 = 5
ElseIf sst(testp1, 1) = 1 Then
loopLRorany1 = 2
loopLRorany2 = 2
loopLRorany3 = 3
loopLRorany4 = 3
loopLRorany5 = 1
loopLRorany6 = 1
ElseIf sst(testp, 3) = 1 Then
loopLRorany1 = 4
loopLRorany2 = 4
loopLRorany3 = 5
loopLRorany4 = 5
loopLRorany5 = 3
loopLRorany6 = 3
ElseIf sst(testp, 1) = 1 Then
loopLRorany1 = 1
loopLRorany2 = 1
loopLRorany3 = 2
loopLRorany4 = 2
loopLRorany5 = 3
loopLRorany6 = 3
End If

ElseIf sst(testp, 1) = 1 Then 'sst(mainp, 5) = 1
If sst(testp1, 5) = 1 Then
loopLRorany1 = 1
loopLRorany2 = 1
loopLRorany3 = 2
loopLRorany4 = 2
loopLRorany5 = 3
loopLRorany6 = 5
ElseIf sst(testp1, 3) = 1 Then
loopLRorany1 = 1
loopLRorany2 = 1
loopLRorany3 = 2
loopLRorany4 = 2
loopLRorany5 = 5
loopLRorany6 = 5
Else
loopLRorany1 = 4
loopLRorany2 = 4
loopLRorany3 = 5
loopLRorany4 = 5
loopLRorany5 = 1
loopLRorany6 = 1
End If

ElseIf sst(testp1, 1) = 1 Then
'here sst(mainp, 5) and sst(testp, 5) must be 1
loopLRorany1 = 2
loopLRorany2 = 4
loopLRorany3 = 0
loopLRorany4 = 0
loopLRorany5 = 1
loopLRorany6 = 1
multctstat = True
End If
If sst(mainp, 3) + sst(mainp, 5) + sst(testp, 3) + sst(testp, 5) + sst(testp1, 3) + sst(testp1, 5) >= 3 Then intRLloop = 2


Case 7 '2 2 X loopLRorany1-4 map double prize (testp), loopLRorany5-6 map spare
If sst(mainp, 5) = 1 And sst(testp, 5) = 1 Then
loopLRorany1 = 1
loopLRorany2 = 4
loopLRorany3 = 0
loopLRorany4 = 0
loopLRorany5 = 1
loopLRorany6 = 5
multctstat = True
sumtype = True
Exit Function
ElseIf sst(mainp, 1) = 1 And sst(testp, 1) = 1 And sst(mainp, 3) = 0 And sst(testp, 3) = 0 Then Exit Function
ElseIf sst(mainp, 1) = 1 Then
    If sst(testp, 5) = 1 Then
    loopLRorany1 = 3
    loopLRorany2 = 4
    loopLRorany3 = 0
    loopLRorany4 = 0
    loopLRorany5 = 3
    loopLRorany6 = 5
    multctstat = True
    ElseIf sst(testp, 3) = 1 Then
    loopLRorany1 = 4
    loopLRorany2 = 4
    loopLRorany3 = 5
    loopLRorany4 = 5
    loopLRorany5 = 3
    loopLRorany6 = 3
    Else
    loopLRorany1 = 1
    loopLRorany2 = 1
    loopLRorany3 = 2
    loopLRorany4 = 2
    loopLRorany5 = 3
    loopLRorany6 = 3
    End If
Else    'mainlr = 0
loopLRorany1 = 1
loopLRorany2 = 1
loopLRorany3 = 2
loopLRorany4 = 2
loopLRorany5 = 3
loopLRorany6 = 5
End If
'singls not required
If sst(mainp, 3) + sst(mainp, 5) + sst(testp, 3) + sst(testp, 5) >= 2 Then intRLloop = 2


Case 8 '2 1 1 1 loopLRorany1-6 map out singles
If sst(mainp, 1) = 1 And sst(testp, 1) = 1 And sst(testp1, 1) = 1 Then Exit Function
If sst(mainp, 1) = 1 And sst(testp, 1) = 1 And sst(testp2, 1) = 1 Then Exit Function
If sst(mainp, 1) = 1 And sst(testp1, 1) = 1 And sst(testp2, 1) = 1 Then Exit Function
If sst(testp, 1) = 1 And sst(testp1, 1) = 1 And sst(testp2, 1) = 1 Then Exit Function

If sst(mainp, 5) = 1 And sst(testp, 5) = 1 And sst(testp1, 5) = 1 And sst(testp2, 5) = 1 Then
loopLRorany1 = 1
loopLRorany2 = 5
loopLRorany3 = 1
loopLRorany4 = 5
loopLRorany5 = 1
loopLRorany6 = 5
sumtype = True
Exit Function
ElseIf sst(mainp, 1) = 1 And sst(testp, 1) = 1 And sst(mainp, 3) = 0 And sst(testp, 3) = 0 Then Exit Function
ElseIf sst(mainp, 1) = 1 And sst(testp1, 1) = 1 And sst(mainp, 3) = 0 And sst(testp1, 3) = 0 Then Exit Function
ElseIf sst(mainp, 1) = 1 And sst(testp2, 1) = 1 And sst(mainp, 3) = 0 And sst(testp2, 3) = 0 Then Exit Function
ElseIf sst(testp, 1) = 1 And sst(testp1, 1) = 1 And sst(testp, 3) = 0 And sst(testp1, 3) = 0 Then Exit Function
ElseIf sst(testp, 1) = 1 And sst(testp2, 1) = 1 And sst(testp, 3) = 0 And sst(testp2, 3) = 0 Then Exit Function
ElseIf sst(testp1, 1) = 1 And sst(testp2, 1) = 1 And sst(testp1, 3) = 0 And sst(testp2, 3) = 0 Then Exit Function

ElseIf sst(mainp, 1) = 1 Then

If sst(testp, 5) = 1 And sst(testp1, 5) = 1 Then
    If sst(testp2, 5) = 1 Then
    loopLRorany1 = 3
    loopLRorany2 = 5
    loopLRorany3 = 3
    loopLRorany4 = 5
    loopLRorany5 = 3
    loopLRorany6 = 5
    ElseIf sst(testp2, 3) = 1 Then
    loopLRorany1 = 3
    loopLRorany2 = 4
    loopLRorany3 = 3
    loopLRorany4 = 4
    loopLRorany5 = 5
    loopLRorany6 = 5
    Else
    loopLRorany1 = 2
    loopLRorany2 = 3
    loopLRorany3 = 2
    loopLRorany4 = 3
    loopLRorany5 = 1
    loopLRorany6 = 1
    End If

ElseIf sst(testp1, 3) = 1 Then 'testp 0 & 2 (5) must = 1
loopLRorany1 = 3
loopLRorany2 = 4
loopLRorany3 = 5
loopLRorany4 = 5
loopLRorany5 = 3
loopLRorany6 = 4
ElseIf sst(testp1, 1) = 1 Then
loopLRorany1 = 2
loopLRorany2 = 3
loopLRorany3 = 1
loopLRorany4 = 1
loopLRorany5 = 2
loopLRorany6 = 3
ElseIf sst(testp, 3) = 1 Then 'again testp 1;2 (5) = 1
loopLRorany1 = 5
loopLRorany2 = 5
loopLRorany3 = 3
loopLRorany4 = 4
loopLRorany5 = 3
loopLRorany6 = 4
Else
loopLRorany1 = 1
loopLRorany2 = 1
loopLRorany3 = 2
loopLRorany4 = 3
loopLRorany5 = 2
loopLRorany6 = 3
End If

ElseIf sst(testp, 1) = 1 Then 'sst(mainp, 5) = 1 see (first if)
If sst(testp1, 5) = 1 And sst(testp2, 5) = 1 Then
loopLRorany1 = 1
loopLRorany2 = 1
loopLRorany3 = 2
loopLRorany4 = 5
loopLRorany5 = 2
loopLRorany6 = 5
ElseIf sst(testp2, 3) = 1 Then
loopLRorany1 = 1
loopLRorany2 = 1
loopLRorany3 = 2
loopLRorany4 = 4
loopLRorany5 = 5
loopLRorany6 = 5
ElseIf sst(testp2, 1) = 1 Then
loopLRorany1 = 5
loopLRorany2 = 5
loopLRorany3 = 2
loopLRorany4 = 4
loopLRorany5 = 1
loopLRorany6 = 1
ElseIf sst(testp1, 3) = 1 Then
loopLRorany1 = 1
loopLRorany2 = 1
loopLRorany3 = 5
loopLRorany4 = 5
loopLRorany5 = 2
loopLRorany6 = 4
Else
loopLRorany1 = 5
loopLRorany2 = 5
loopLRorany3 = 1
loopLRorany4 = 1
loopLRorany5 = 2
loopLRorany6 = 4
End If

ElseIf sst(testp1, 1) = 1 Then
'here sst(mainp, 5) and sst(testp, 5) must be 1
If sst(testp2, 5) = 1 Then
loopLRorany1 = 2
loopLRorany2 = 5
loopLRorany3 = 1
loopLRorany4 = 1
loopLRorany5 = 2
loopLRorany6 = 5
ElseIf sst(testp2, 3) = 1 Then
loopLRorany1 = 2
loopLRorany2 = 4
loopLRorany3 = 1
loopLRorany4 = 1
loopLRorany5 = 5
loopLRorany6 = 5
Else
loopLRorany1 = 2
loopLRorany2 = 4
loopLRorany3 = 5
loopLRorany4 = 5
loopLRorany5 = 1
loopLRorany6 = 1
End If

ElseIf sst(testp2, 1) = 1 Then
'All previous three are "anys"
loopLRorany1 = 2
loopLRorany2 = 5
loopLRorany3 = 2
loopLRorany4 = 5
loopLRorany5 = 1
loopLRorany6 = 1
End If
If sst(mainp, 3) + sst(mainp, 5) + sst(testp, 3) + sst(testp, 5) + sst(testp1, 3) + sst(testp1, 5) + sst(testp2, 3) + sst(testp2, 5) >= 4 Then intRLloop = 2


Case 9

'9 - 2 1 1 X loopLRorany1-4 map out single prizes, loopLRorany5-6 maps the no-prize
If sst(mainp, 1) = 1 And sst(testp, 1) = 1 And sst(testp1, 1) = 1 Then Exit Function
If sst(mainp, 5) = 1 And sst(testp, 5) = 1 And sst(testp1, 5) = 1 Then
loopLRorany1 = 1
loopLRorany2 = 5
loopLRorany3 = 1
loopLRorany4 = 5
loopLRorany5 = 1
loopLRorany6 = 5
sumtype = True
Exit Function
ElseIf sst(mainp, 1) = 1 And sst(testp, 1) = 1 And sst(mainp, 3) = 0 And sst(testp, 3) = 0 Then Exit Function
ElseIf sst(mainp, 1) = 1 And sst(testp1, 1) = 1 And sst(mainp, 3) = 0 And sst(testp1, 3) = 0 Then Exit Function
ElseIf sst(testp, 1) = 1 And sst(testp1, 1) = 1 And sst(testp, 3) = 0 And sst(testp1, 3) = 0 Then Exit Function
ElseIf sst(mainp, 1) = 1 Then
    If sst(testp, 5) = 1 And sst(testp1, 5) = 1 Then
    loopLRorany1 = 3
    loopLRorany2 = 5
    loopLRorany3 = 3
    loopLRorany4 = 5
    loopLRorany5 = 3
    loopLRorany6 = 5
    Rsgl 3
    ElseIf sst(testp1, 3) = 1 Then
    loopLRorany1 = 3
    loopLRorany2 = 4
    loopLRorany3 = 5
    loopLRorany4 = 5
    loopLRorany5 = 3
    loopLRorany6 = 4
    Rsgl 3, 4
    ElseIf sst(testp1, 1) = 1 Then
    loopLRorany1 = 2
    loopLRorany2 = 3
    loopLRorany3 = 1
    loopLRorany4 = 1
    loopLRorany5 = 2
    loopLRorany6 = 3
    Rsgl 3, 2
    ElseIf sst(testp, 3) = 1 Then
    loopLRorany1 = 5
    loopLRorany2 = 5
    loopLRorany3 = 3
    loopLRorany4 = 4
    loopLRorany5 = 3
    loopLRorany6 = 4
    Rsgl 3, 4
    ElseIf sst(testp, 1) = 1 Then
    loopLRorany1 = 1
    loopLRorany2 = 1
    loopLRorany3 = 2
    loopLRorany4 = 3
    loopLRorany5 = 2
    loopLRorany6 = 3
    Rsgl 3, 2
    End If

ElseIf sst(testp, 1) = 1 Then
If sst(testp1, 5) = 1 Then 'sst(mainp, 1)=0
loopLRorany1 = 1
loopLRorany2 = 1
loopLRorany3 = 2
loopLRorany4 = 5
loopLRorany5 = 2
loopLRorany6 = 5
Rsgl 2
ElseIf sst(testp1, 3) = 1 Then
loopLRorany1 = 1
loopLRorany2 = 1
loopLRorany3 = 5
loopLRorany4 = 5
loopLRorany5 = 2
loopLRorany6 = 4
Rsgl 2, 4
Else
loopLRorany1 = 5
loopLRorany2 = 5
loopLRorany3 = 1
loopLRorany4 = 1
loopLRorany5 = 2
loopLRorany6 = 4
Rsgl 4, 2
End If

ElseIf sst(testp1, 1) = 1 Then
'here sst(mainp, 5) and sst(testp, 5) must be 1
loopLRorany1 = 2
loopLRorany2 = 5
loopLRorany3 = 1
loopLRorany4 = 1
loopLRorany5 = 2
loopLRorany6 = 5
Rsgl 2
End If
If sst(mainp, 3) + sst(mainp, 5) + sst(testp, 3) + sst(testp, 5) + sst(testp1, 3) + sst(testp1, 5) >= 3 Then intRLloop = 2


Case 10

'2 1 X X loopLRorany1-4 map no-prize, loopLRorany5-6 map out single,singl2 not required
If sst(mainp, 5) = 1 And sst(testp, 5) = 1 Then
loopLRorany1 = 1
loopLRorany2 = 4
loopLRorany3 = 0
loopLRorany4 = 0
loopLRorany5 = 1
loopLRorany6 = 5
multctstat = True
sumtype = True
Exit Function
ElseIf sst(mainp, 1) = 1 And sst(testp, 1) = 1 And sst(mainp, 3) = 0 And sst(testp, 3) = 0 Then Exit Function
ElseIf sst(mainp, 1) = 1 Then
    If sst(testp, 5) = 1 Then
    loopLRorany1 = 3
    loopLRorany2 = 4
    loopLRorany3 = 0
    loopLRorany4 = 0
    loopLRorany5 = 3
    loopLRorany6 = 5
    multctstat = True
    ElseIf sst(testp, 3) = 1 Then
    loopLRorany1 = 3
    loopLRorany2 = 3
    loopLRorany3 = 4
    loopLRorany4 = 4
    loopLRorany5 = 5
    loopLRorany6 = 5
    Else
    loopLRorany1 = 2
    loopLRorany2 = 2
    loopLRorany3 = 3
    loopLRorany4 = 3
    loopLRorany5 = 1
    loopLRorany6 = 1
    End If
    Rsgl 3
Else    'mainlr=0
loopLRorany1 = 2
loopLRorany2 = 4
loopLRorany3 = 0
loopLRorany4 = 0
loopLRorany5 = 1
loopLRorany6 = 1
Rsgl 2
multctstat = True
End If

If sst(mainp, 3) + sst(mainp, 5) + sst(testp, 3) + sst(testp, 5) >= 2 Then intRLloop = 2


Case 11

'2 X X X  counter1 - 6 no-prize
If sst(mainp, 5) = 1 Then
loopLRorany1 = 1
loopLRorany2 = 3
loopLRorany3 = 0
loopLRorany4 = 0
loopLRorany5 = 0
loopLRorany6 = 0
multctstat = True
sumtype = True
Exit Function
Else
loopLRorany1 = 3
loopLRorany2 = 3
loopLRorany3 = 4
loopLRorany4 = 4
loopLRorany5 = 5
loopLRorany6 = 5
End If

If sst(mainp, 3) = 1 Then intRLloop = 2


Case 12 '1 1 1 1 1 loopLRorany1-8 map out singles
If sst(mainp, 1) = 1 And sst(testp, 1) = 1 And sst(testp1, 1) = 1 Then Exit Function
If sst(mainp, 1) = 1 And sst(testp, 1) = 1 And sst(testp2, 1) = 1 Then Exit Function
If sst(mainp, 1) = 1 And sst(testp, 1) = 1 And sst(testp3, 1) = 1 Then Exit Function
If sst(mainp, 1) = 1 And sst(testp1, 1) = 1 And sst(testp2, 1) = 1 Then Exit Function
If sst(mainp, 1) = 1 And sst(testp1, 1) = 1 And sst(testp3, 1) = 1 Then Exit Function
If sst(mainp, 1) = 1 And sst(testp2, 1) = 1 And sst(testp3, 1) = 1 Then Exit Function
If sst(testp, 1) = 1 And sst(testp1, 1) = 1 And sst(testp2, 1) = 1 Then Exit Function
If sst(testp, 1) = 1 And sst(testp1, 1) = 1 And sst(testp3, 1) = 1 Then Exit Function
If sst(testp, 1) = 1 And sst(testp2, 1) = 1 And sst(testp3, 1) = 1 Then Exit Function
If sst(testp1, 1) = 1 And sst(testp2, 1) = 1 And sst(testp3, 1) = 1 Then Exit Function


If sst(mainp, 5) = 1 And sst(testp, 5) = 1 And sst(testp1, 5) = 1 And sst(testp2, 5) = 1 And sst(testp3, 5) = 1 Then
loopLRorany1 = 1
loopLRorany2 = 5
loopLRorany3 = 1
loopLRorany4 = 5
loopLRorany5 = 1
loopLRorany6 = 5
loopLRorany7 = 1
loopLRorany8 = 5
sumtype = True
Exit Function
ElseIf sst(mainp, 1) = 1 And sst(testp, 1) = 1 And sst(mainp, 3) = 0 And sst(testp, 3) = 0 Then Exit Function
ElseIf sst(mainp, 1) = 1 And sst(testp1, 1) = 1 And sst(mainp, 3) = 0 And sst(testp1, 3) = 0 Then Exit Function
ElseIf sst(mainp, 1) = 1 And sst(testp2, 1) = 1 And sst(mainp, 3) = 0 And sst(testp2, 3) = 0 Then Exit Function
ElseIf sst(mainp, 1) = 1 And sst(testp3, 1) = 1 And sst(mainp, 3) = 0 And sst(testp3, 3) = 0 Then Exit Function
ElseIf sst(testp, 1) = 1 And sst(testp1, 1) = 1 And sst(testp, 3) = 0 And sst(testp1, 3) = 0 Then Exit Function
ElseIf sst(testp, 1) = 1 And sst(testp2, 1) = 1 And sst(testp, 3) = 0 And sst(testp2, 3) = 0 Then Exit Function
ElseIf sst(testp, 1) = 1 And sst(testp3, 1) = 1 And sst(testp, 3) = 0 And sst(testp3, 3) = 0 Then Exit Function
ElseIf sst(testp1, 1) = 1 And sst(testp2, 1) = 1 And sst(testp1, 3) = 0 And sst(testp2, 3) = 0 Then Exit Function
ElseIf sst(testp1, 1) = 1 And sst(testp3, 1) = 1 And sst(testp1, 3) = 0 And sst(testp3, 3) = 0 Then Exit Function
ElseIf sst(testp2, 1) = 1 And sst(testp3, 1) = 1 And sst(testp2, 3) = 0 And sst(testp3, 3) = 0 Then Exit Function
ElseIf sst(mainp, 1) = 1 Then

    If sst(testp, 5) = 1 And sst(testp1, 5) = 1 And sst(testp2, 5) = 1 Then
    If sst(testp3, 5) = 1 Then
    loopLRorany1 = 2
    loopLRorany2 = 5
    loopLRorany3 = 2
    loopLRorany4 = 5
    loopLRorany5 = 2
    loopLRorany6 = 5
    loopLRorany7 = 2
    loopLRorany8 = 5
    ElseIf sst(testp3, 3) = 1 Then
    loopLRorany1 = 2
    loopLRorany2 = 4
    loopLRorany3 = 2
    loopLRorany4 = 4
    loopLRorany5 = 2
    loopLRorany6 = 4
    loopLRorany7 = 5
    loopLRorany8 = 5
    Else
    loopLRorany1 = 2
    loopLRorany2 = 4
    loopLRorany3 = 2
    loopLRorany4 = 4
    loopLRorany5 = 2
    loopLRorany6 = 4
    loopLRorany7 = 1
    loopLRorany8 = 1
    End If

    ElseIf sst(testp, 5) = 1 And sst(testp1, 5) = 1 And sst(testp3, 5) = 1 Then
    If sst(testp2, 3) = 1 Then
    loopLRorany1 = 2
    loopLRorany2 = 4
    loopLRorany3 = 2
    loopLRorany4 = 4
    loopLRorany5 = 5
    loopLRorany6 = 5
    loopLRorany7 = 2
    loopLRorany8 = 4
    Else
    loopLRorany1 = 2
    loopLRorany2 = 4
    loopLRorany3 = 2
    loopLRorany4 = 4
    loopLRorany5 = 1
    loopLRorany6 = 1
    loopLRorany7 = 2
    loopLRorany8 = 4
    End If
    ElseIf sst(testp, 5) = 1 And sst(testp2, 5) = 1 And sst(testp3, 5) = 1 Then
    If sst(testp1, 3) = 1 Then
    loopLRorany1 = 2
    loopLRorany2 = 4
    loopLRorany3 = 5
    loopLRorany4 = 5
    loopLRorany5 = 2
    loopLRorany6 = 4
    loopLRorany7 = 2
    loopLRorany8 = 4
    Else
    loopLRorany1 = 2
    loopLRorany2 = 4
    loopLRorany3 = 1
    loopLRorany4 = 1
    loopLRorany5 = 2
    loopLRorany6 = 4
    loopLRorany7 = 2
    loopLRorany8 = 4
    End If
    Else    'testp1, testp2,testp3(5) = 1
    If sst(testp, 3) = 1 Then
    loopLRorany1 = 5
    loopLRorany2 = 5
    loopLRorany3 = 2
    loopLRorany4 = 4
    loopLRorany5 = 2
    loopLRorany6 = 4
    loopLRorany7 = 2
    loopLRorany8 = 4
    Else
    loopLRorany1 = 1
    loopLRorany2 = 1
    loopLRorany3 = 2
    loopLRorany4 = 4
    loopLRorany5 = 2
    loopLRorany6 = 4
    loopLRorany7 = 2
    loopLRorany8 = 4
    End If
    End If
ElseIf sst(testp, 1) = 1 Then 'mainp5  is 1
    If sst(testp1, 5) = 1 And sst(testp2, 5) = 1 Then
    If sst(testp3, 5) = 1 Then
    loopLRorany1 = 1
    loopLRorany2 = 1
    loopLRorany3 = 2
    loopLRorany4 = 5
    loopLRorany5 = 2
    loopLRorany6 = 5
    loopLRorany7 = 2
    loopLRorany8 = 5
    ElseIf sst(testp3, 3) = 1 Then
    loopLRorany1 = 1
    loopLRorany2 = 1
    loopLRorany3 = 2
    loopLRorany4 = 4
    loopLRorany5 = 2
    loopLRorany6 = 4
    loopLRorany7 = 5
    loopLRorany8 = 5
    Else
    loopLRorany1 = 5
    loopLRorany2 = 5
    loopLRorany3 = 2
    loopLRorany4 = 4
    loopLRorany5 = 2
    loopLRorany6 = 4
    loopLRorany7 = 1
    loopLRorany8 = 1
    End If
    ElseIf sst(testp1, 5) = 1 And sst(testp3, 5) = 1 Then
    If sst(testp2, 3) = 1 Then
    loopLRorany1 = 1
    loopLRorany2 = 1
    loopLRorany3 = 2
    loopLRorany4 = 4
    loopLRorany5 = 5
    loopLRorany6 = 5
    loopLRorany7 = 2
    loopLRorany8 = 4
    Else
    loopLRorany1 = 5
    loopLRorany2 = 5
    loopLRorany3 = 2
    loopLRorany4 = 4
    loopLRorany5 = 1
    loopLRorany6 = 1
    loopLRorany7 = 2
    loopLRorany8 = 4
    End If
    ElseIf sst(testp2, 5) = 1 And sst(testp3, 5) = 1 Then
    If sst(testp1, 3) = 1 Then
    loopLRorany1 = 1
    loopLRorany2 = 1
    loopLRorany3 = 5
    loopLRorany4 = 5
    loopLRorany5 = 2
    loopLRorany6 = 4
    loopLRorany7 = 2
    loopLRorany8 = 4
    Else
    loopLRorany1 = 5
    loopLRorany2 = 5
    loopLRorany3 = 1
    loopLRorany4 = 1
    loopLRorany5 = 2
    loopLRorany6 = 4
    loopLRorany7 = 2
    loopLRorany8 = 4
    End If
    End If
ElseIf sst(testp1, 1) = 1 Then 'mainp5 , testp5 is 1
    If sst(testp2, 5) = 1 And sst(testp3, 5) = 1 Then
    loopLRorany1 = 2
    loopLRorany2 = 5
    loopLRorany3 = 1
    loopLRorany4 = 1
    loopLRorany5 = 2
    loopLRorany6 = 5
    loopLRorany7 = 2
    loopLRorany8 = 5
    ElseIf sst(testp3, 3) = 1 Then
    loopLRorany1 = 2
    loopLRorany2 = 4
    loopLRorany3 = 1
    loopLRorany4 = 1
    loopLRorany5 = 2
    loopLRorany6 = 4
    loopLRorany7 = 5
    loopLRorany8 = 5
    ElseIf sst(testp3, 1) = 1 Then
    loopLRorany1 = 2
    loopLRorany2 = 4
    loopLRorany3 = 5
    loopLRorany4 = 5
    loopLRorany5 = 2
    loopLRorany6 = 4
    loopLRorany7 = 1
    loopLRorany8 = 1
    ElseIf sst(testp2, 3) = 1 Then
    loopLRorany1 = 2
    loopLRorany2 = 4
    loopLRorany3 = 1
    loopLRorany4 = 1
    loopLRorany5 = 5
    loopLRorany6 = 5
    loopLRorany7 = 2
    loopLRorany8 = 4
    Else
    loopLRorany1 = 2
    loopLRorany2 = 4
    loopLRorany3 = 5
    loopLRorany4 = 5
    loopLRorany5 = 1
    loopLRorany6 = 1
    loopLRorany7 = 2
    loopLRorany8 = 4
    End If
ElseIf sst(testp2, 1) = 1 Then 'mainp5 , testp5, testp15 is 1
    If sst(testp3, 5) = 1 Then
    loopLRorany1 = 2
    loopLRorany2 = 5
    loopLRorany3 = 2
    loopLRorany4 = 5
    loopLRorany5 = 1
    loopLRorany6 = 1
    loopLRorany7 = 2
    loopLRorany8 = 5
    ElseIf sst(testp3, 3) = 1 Then
    loopLRorany1 = 2
    loopLRorany2 = 4
    loopLRorany3 = 2
    loopLRorany4 = 4
    loopLRorany5 = 1
    loopLRorany6 = 1
    loopLRorany7 = 5
    loopLRorany8 = 5
    Else
    loopLRorany1 = 2
    loopLRorany2 = 4
    loopLRorany3 = 2
    loopLRorany4 = 4
    loopLRorany5 = 5
    loopLRorany6 = 5
    loopLRorany7 = 1
    loopLRorany8 = 1
    End If
Else
loopLRorany1 = 2
loopLRorany2 = 5
loopLRorany3 = 2
loopLRorany4 = 5
loopLRorany5 = 2
loopLRorany6 = 5
loopLRorany7 = 1
loopLRorany8 = 1
End If

If sst(mainp, 3) + sst(mainp, 5) + sst(testp, 3) + sst(testp, 5) + sst(testp1, 3) + sst(testp1, 5) + sst(testp2, 3) + sst(testp2, 5) + sst(testp3, 3) + sst(testp3, 5) >= 5 Then intRLloop = 2


Case 13 '1 1 1 1 X 'loopLRorany1-6 map out single prizes, loopLRorany5-6 maps the no-prize
If sst(mainp, 1) = 1 And sst(testp, 1) = 1 And sst(testp1, 1) = 1 Then Exit Function
If sst(mainp, 1) = 1 And sst(testp, 1) = 1 And sst(testp2, 1) = 1 Then Exit Function
If sst(mainp, 1) = 1 And sst(testp1, 1) = 1 And sst(testp2, 1) = 1 Then Exit Function
If sst(testp, 1) = 1 And sst(testp1, 1) = 1 And sst(testp2, 1) = 1 Then Exit Function

If sst(mainp, 5) = 1 And sst(testp, 5) = 1 And sst(testp1, 5) = 1 And sst(testp2, 5) = 1 Then
loopLRorany1 = 1
loopLRorany2 = 5
loopLRorany3 = 1
loopLRorany4 = 5
loopLRorany5 = 1
loopLRorany6 = 5
loopLRorany7 = 1
loopLRorany8 = 5
sumtype = True
Exit Function
ElseIf sst(mainp, 1) = 1 And sst(testp, 1) = 1 And sst(mainp, 3) = 0 And sst(testp, 3) = 0 Then Exit Function
ElseIf sst(mainp, 1) = 1 And sst(testp1, 1) = 1 And sst(mainp, 3) = 0 And sst(testp1, 3) = 0 Then Exit Function
ElseIf sst(mainp, 1) = 1 And sst(testp2, 1) = 1 And sst(mainp, 3) = 0 And sst(testp2, 3) = 0 Then Exit Function
ElseIf sst(testp, 1) = 1 And sst(testp1, 1) = 1 And sst(testp, 3) = 0 And sst(testp1, 3) = 0 Then Exit Function
ElseIf sst(testp, 1) = 1 And sst(testp2, 1) = 1 And sst(testp, 3) = 0 And sst(testp2, 3) = 0 Then Exit Function
ElseIf sst(testp1, 1) = 1 And sst(testp2, 1) = 1 And sst(testp1, 3) = 0 And sst(testp2, 3) = 0 Then Exit Function
ElseIf sst(mainp, 1) = 1 Then
If sst(testp, 5) = 1 And sst(testp1, 5) = 1 Then
    If sst(testp2, 5) = 1 Then
    loopLRorany1 = 2
    loopLRorany2 = 5
    loopLRorany3 = 2
    loopLRorany4 = 5
    loopLRorany5 = 2
    loopLRorany6 = 5
    loopLRorany7 = 2
    loopLRorany8 = 5
    Rsgl 2
    ElseIf sst(testp2, 3) = 1 Then
    loopLRorany1 = 2
    loopLRorany2 = 4
    loopLRorany3 = 2
    loopLRorany4 = 4
    loopLRorany5 = 5
    loopLRorany6 = 5
    loopLRorany7 = 2
    loopLRorany8 = 4
    Rsgl 2, 4
    Else
    loopLRorany1 = 2
    loopLRorany2 = 4
    loopLRorany3 = 2
    loopLRorany4 = 4
    loopLRorany5 = 1
    loopLRorany6 = 1
    loopLRorany7 = 2
    loopLRorany8 = 4
    Rsgl 4, 2
    End If

ElseIf sst(testp1, 3) = 1 Then 'testp 0 & 2 (5) must = 1
loopLRorany1 = 2
loopLRorany2 = 4
loopLRorany3 = 5
loopLRorany4 = 5
loopLRorany5 = 2
loopLRorany6 = 4
loopLRorany7 = 2
loopLRorany8 = 4
Rsgl 2, 4
ElseIf sst(testp1, 1) = 1 Then
loopLRorany1 = 2
loopLRorany2 = 4
loopLRorany3 = 1
loopLRorany4 = 1
loopLRorany5 = 2
loopLRorany6 = 4
loopLRorany7 = 2
loopLRorany8 = 4
Rsgl 4, 2
ElseIf sst(testp, 3) = 1 Then 'again testp 1;2 (5) = 1
loopLRorany1 = 5
loopLRorany2 = 5
loopLRorany3 = 2
loopLRorany4 = 4
loopLRorany5 = 2
loopLRorany6 = 4
loopLRorany7 = 2
loopLRorany8 = 4
Rsgl 2, 4
Else
loopLRorany1 = 1
loopLRorany2 = 1
loopLRorany3 = 2
loopLRorany4 = 4
loopLRorany5 = 2
loopLRorany6 = 4
loopLRorany7 = 2
loopLRorany8 = 4
Rsgl 4, 2
End If

ElseIf sst(testp, 1) = 1 Then 'sst(mainp, 5) = 1 see (first if)
If sst(testp1, 5) = 1 And sst(testp2, 5) = 1 Then
loopLRorany1 = 1
loopLRorany2 = 1
loopLRorany3 = 2
loopLRorany4 = 5
loopLRorany5 = 2
loopLRorany6 = 5
loopLRorany7 = 2
loopLRorany8 = 5
Rsgl 2
ElseIf sst(testp2, 3) = 1 Then
loopLRorany1 = 1
loopLRorany2 = 1
loopLRorany3 = 2
loopLRorany4 = 4
loopLRorany5 = 5
loopLRorany6 = 5
loopLRorany7 = 2
loopLRorany8 = 4
Rsgl 2, 4
ElseIf sst(testp2, 1) = 1 Then
loopLRorany1 = 5
loopLRorany2 = 5
loopLRorany3 = 2
loopLRorany4 = 4
loopLRorany5 = 1
loopLRorany6 = 1
loopLRorany7 = 2
loopLRorany8 = 4
Rsgl 4, 2
ElseIf sst(testp1, 3) = 1 Then
loopLRorany1 = 1
loopLRorany2 = 1
loopLRorany3 = 5
loopLRorany4 = 5
loopLRorany5 = 2
loopLRorany6 = 4
loopLRorany7 = 2
loopLRorany8 = 4
Rsgl 2, 4
Else
loopLRorany1 = 5
loopLRorany2 = 5
loopLRorany3 = 1
loopLRorany4 = 1
loopLRorany5 = 2
loopLRorany6 = 4
loopLRorany7 = 2
loopLRorany8 = 4
Rsgl 4, 2
End If

ElseIf sst(testp1, 1) = 1 Then
'here sst(mainp, 5) and sst(testp, 5) must be 1
If sst(testp2, 5) = 1 Then
loopLRorany1 = 2
loopLRorany2 = 5
loopLRorany3 = 1
loopLRorany4 = 1
loopLRorany5 = 2
loopLRorany6 = 5
loopLRorany7 = 2
loopLRorany8 = 5
Rsgl 2
ElseIf sst(testp2, 3) = 1 Then
loopLRorany1 = 2
loopLRorany2 = 4
loopLRorany3 = 1
loopLRorany4 = 1
loopLRorany5 = 5
loopLRorany6 = 5
loopLRorany7 = 2
loopLRorany8 = 4
Rsgl 2, 4
Else
loopLRorany1 = 2
loopLRorany2 = 4
loopLRorany3 = 5
loopLRorany4 = 5
loopLRorany5 = 1
loopLRorany6 = 1
loopLRorany7 = 2
loopLRorany8 = 4
Rsgl 4, 2
End If

ElseIf sst(testp2, 1) = 1 Then
'All previous three are "anys"
loopLRorany1 = 2
loopLRorany2 = 5
loopLRorany3 = 2
loopLRorany4 = 5
loopLRorany5 = 1
loopLRorany6 = 1
loopLRorany7 = 2
loopLRorany8 = 5
Rsgl 2
End If
If sst(mainp, 3) + sst(mainp, 5) + sst(testp, 3) + sst(testp, 5) + sst(testp1, 3) + sst(testp1, 5) + sst(testp2, 3) + sst(testp2, 5) >= 4 Then intRLloop = 2


Case 14 '1 1 1 X X 'loopLRorany1-4 map non -scoring double, loopLRorany5-8 maps singles testp, testp1
If sst(mainp, 1) = 1 And sst(testp, 1) = 1 And sst(testp1, 1) = 1 Then Exit Function
If sst(mainp, 5) = 1 And sst(testp, 5) = 1 And sst(testp1, 5) = 1 Then
loopLRorany1 = 1
loopLRorany2 = 4
loopLRorany3 = 0
loopLRorany4 = 0
loopLRorany5 = 1
loopLRorany6 = 5
loopLRorany7 = 1
loopLRorany8 = 5
sumtype = True
multctstat = True
Exit Function
ElseIf sst(mainp, 1) = 1 And sst(testp, 1) = 1 And sst(mainp, 3) = 0 And sst(testp, 3) = 0 Then Exit Function
ElseIf sst(mainp, 1) = 1 And sst(testp1, 1) = 1 And sst(mainp, 3) = 0 And sst(testp1, 3) = 0 Then Exit Function
ElseIf sst(testp, 1) = 1 And sst(testp1, 1) = 1 And sst(testp, 3) = 0 And sst(testp1, 3) = 0 Then Exit Function
ElseIf sst(mainp, 1) = 1 Then

    If sst(testp, 5) = 1 And sst(testp1, 5) = 1 Then
    loopLRorany1 = 2
    loopLRorany2 = 4
    loopLRorany3 = 0
    loopLRorany4 = 0
    loopLRorany5 = 2
    loopLRorany6 = 5
    loopLRorany7 = 2
    loopLRorany8 = 5
    Rsgl 2
    ElseIf sst(testp1, 3) = 1 Then
    loopLRorany1 = 2
    loopLRorany2 = 3
    loopLRorany3 = 0
    loopLRorany4 = 0
    loopLRorany5 = 2
    loopLRorany6 = 4
    loopLRorany7 = 5
    loopLRorany8 = 5
    Rsgl 2, 4
    BLOCK111 = True
    ElseIf sst(testp1, 1) = 1 Then
    loopLRorany1 = 2
    loopLRorany2 = 3
    loopLRorany3 = 0
    loopLRorany4 = 0
    loopLRorany5 = 2
    loopLRorany6 = 4
    loopLRorany7 = 1
    loopLRorany8 = 1
    Rsgl 4, 2
    BLOCK111 = True
    ElseIf sst(testp, 3) = 1 Then
    loopLRorany1 = 2
    loopLRorany2 = 3
    loopLRorany3 = 0
    loopLRorany4 = 0
    loopLRorany5 = 5
    loopLRorany6 = 5
    loopLRorany7 = 2
    loopLRorany8 = 4
    Rsgl 2, 4
    BLOCK111 = True
    Else
    loopLRorany1 = 2
    loopLRorany2 = 3
    loopLRorany3 = 0
    loopLRorany4 = 0
    loopLRorany5 = 1
    loopLRorany6 = 1
    loopLRorany7 = 2
    loopLRorany8 = 4
    BLOCK111 = True
    Rsgl 4, 2
    End If
'mainlr = 0
ElseIf sst(testp, 1) = 1 Then
    If sst(testp1, 5) = 1 Then
    loopLRorany1 = 2
    loopLRorany2 = 4
    loopLRorany3 = 0
    loopLRorany4 = 0
    loopLRorany5 = 1
    loopLRorany6 = 1
    loopLRorany7 = 2
    loopLRorany8 = 5
    Rsgl 2
    ElseIf sst(testp1, 3) = 1 Then
    loopLRorany1 = 2
    loopLRorany2 = 3
    loopLRorany3 = 0
    loopLRorany4 = 0
    loopLRorany5 = 1
    loopLRorany6 = 1
    loopLRorany7 = 5
    loopLRorany8 = 5
    Rsgl 2, 4
    BLOCK111 = True
    Else
    loopLRorany1 = 2
    loopLRorany2 = 3
    loopLRorany3 = 0
    loopLRorany4 = 0
    loopLRorany5 = 5
    loopLRorany6 = 5
    loopLRorany7 = 1
    loopLRorany8 = 1
    BLOCK111 = True
    Rsgl 4, 2
    End If
ElseIf sst(testp1, 1) = 1 Then
loopLRorany1 = 2
loopLRorany2 = 4
loopLRorany3 = 0
loopLRorany4 = 0
loopLRorany5 = 2
loopLRorany6 = 5
loopLRorany7 = 1
loopLRorany8 = 1
Rsgl 2
End If
multctstat = True       'All above combinations require this
If sst(mainp, 3) + sst(mainp, 5) + sst(testp, 3) + sst(testp, 5) + sst(testp1, 3) + sst(testp1, 5) >= 3 Then intRLloop = 2


Case 15 '1 1 X X X 'loopLRorany1-6 map non - scoring triple, loopLRorany7-8 maps testp,singl2 not reqired


If sst(mainp, 5) = 1 And sst(testp, 5) = 1 Then
loopLRorany1 = 1
loopLRorany2 = 3
loopLRorany3 = 0
loopLRorany4 = 0
loopLRorany5 = 0
loopLRorany6 = 0
loopLRorany7 = 1
loopLRorany8 = 5
multctstat = True
sumtype = True
Exit Function
ElseIf sst(mainp, 1) = 1 And sst(testp, 1) = 1 And sst(mainp, 3) = 0 And sst(testp, 3) = 0 Then Exit Function
ElseIf sst(mainp, 1) = 1 Then
If sst(testp, 5) = 1 Then
loopLRorany1 = 2
loopLRorany2 = 3
loopLRorany3 = 0
loopLRorany4 = 0
loopLRorany5 = 0
loopLRorany6 = 0
loopLRorany7 = 2
loopLRorany8 = 5
Rsgl 2
multctstat = True
ElseIf sst(testp, 3) = 1 Then
loopLRorany1 = 2
loopLRorany2 = 2
loopLRorany3 = 3
loopLRorany4 = 3
loopLRorany5 = 4
loopLRorany6 = 4
loopLRorany7 = 5
loopLRorany8 = 5
Rsgl 2
Else
loopLRorany1 = 2
loopLRorany2 = 2
loopLRorany3 = 3
loopLRorany4 = 3
loopLRorany5 = 4
loopLRorany6 = 4
loopLRorany7 = 1
loopLRorany8 = 1
Rsgl 4
End If
Else 'sstab (mainp 5) = 1
loopLRorany1 = 2
loopLRorany2 = 3
loopLRorany3 = 0
loopLRorany4 = 0
loopLRorany5 = 0
loopLRorany6 = 0
loopLRorany7 = 1
loopLRorany8 = 1
Rsgl 2
multctstat = True
End If

If sst(mainp, 3) + sst(mainp, 5) + sst(testp, 3) + sst(testp, 5) >= 2 Then intRLloop = 2


Case 16 '1 X X X X 'loopLRorany1-8 map non -scoring quad
If sst(mainp, 5) = 1 Then
'NO loopfix required here as ct1 <> ct2 <> ct3 etc
loopLRorany1 = 1
loopLRorany2 = 2
loopLRorany3 = 2
loopLRorany4 = 3
loopLRorany5 = 3
loopLRorany6 = 4
loopLRorany7 = 4
loopLRorany8 = 5
Else
loopLRorany1 = 2
loopLRorany2 = 2
loopLRorany3 = 3
loopLRorany4 = 3
loopLRorany5 = 4
loopLRorany6 = 4
loopLRorany7 = 5
loopLRorany8 = 5
End If

If sst(mainp, 3) = 1 Then intRLloop = 2

End Select

singl1 = S1
singl2 = S2
sumtype = True
End Function
Private Sub Rsgl(zsingl1 As Long, Optional zsingl2 As Long = 5)
'singl1, singl2 only of interest if a spare slot
'Reelend not relevant when both LR 1
S1 = zsingl1
S2 = zsingl2
End Sub

