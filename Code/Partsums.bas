Attribute VB_Name = "Partsums"
Option Explicit
Private intcount As Long, ct As Long, thumbsize As Long
Private singl1 As Long, singl2 As Long, reelend As Long
Private tmp As Long, tmp1 As Long, tmp2 As Long, zsum As Long
Private Xrst As Long, Xrst1 As Long, Xmn As Long
Private workvec(5, 14) As Long, substotalvec(5) As Long, exv(5) As Long, qrst(5) As Long
Private multptstat As Boolean, partdir As Long, nopics As Long, reelduc(5) As Long
Private RLloop As Long, mp As Long, tp As Long, tp1 As Long, tp2 As Long, tp3 As Long
Private tryRL As Long, visitextra As Boolean
Private loopLRany1 As Long, loopLRany2 As Long, loopLRany3 As Long
Private loopLRany4 As Long, loopLRany5 As Long, loopLRany6 As Long
Private mdpz As Long, mspz As Long, aspz As Long, aspz1 As Long
Private MLR As Long, MRL As Long, TLR As Long, TRL As Long, T1LR As Long, T1RL As Long
Private mainRL As Long, mainLR As Long, mainmid As Long
Private testLR As Long, test1LR As Long, test2LR As Long
Private testRL As Long, test1RL As Long, test2RL As Long
Public Sub partials(substotalarr() As Long, wheelvec() As Long, zthumbsize As Long)
'`NOTE THAT IF DEALING WITH SUBSTITUTES, ALL COMBINATIONS HERE WILL NEED TO
'BE EXTENDED TO COUNT SUBSTITUTES, ie singlsum, dsum, tsum
'will all include substitutes and substotalvec can be dispensed with

thumbsize = zthumbsize


'Initialise workvec
For ct = 1 To 5
singlsum(ct) = 0
For mp = 1 To thumbsize
workvec(ct, mp) = (wheelvec(ct, mp))
Next
Next
'Need to sum ALL substituted combos here for the sumnation of substitutes to work

For ct = 1 To 5
substotalvec(ct) = substotalarr(ct)
singlsum(ct) = ssm(ct)
Next

End Sub
Private Function ssm(cc1 As Long, Optional pct1 As Long = 0, Optional pct2 As Long = 0, Optional pct3 As Long = 0) As Long

ssm = 0

'safe using tp3
For tp3 = 1 To thumbsize  '1

If sst(tp3, 10) > 0 Then
If tp3 <> pct1 And tp3 <> pct2 And tp3 <> pct3 Then
If sst(tp3, 5) = 1 Then
ssm = ssm + workvec(cc1, tp3)
Else
If cc1 = 1 Then ssm = ssm + workvec(1, tp3)
If cc1 = 5 And sst(tp3, 3) = 1 Then ssm = ssm + workvec(5, tp3)
End If
End If
End If
Next
End Function
Public Function dsm(ct1 As Long, ct2 As Long, pct1 As Long, Optional pct2 As Long = 0, Optional pct3 As Long = 0) As Long
Dim c1 As Long, c2 As Long, dexv(5) As Long, compsize As Long

If ct1 < ct2 Then
c1 = ct1
c2 = ct2
Else
c1 = ct2
c2 = ct1
End If

For intcount = 1 To 5
dexv(intcount) = 0
Next

compsize = 3

If pct3 = 0 Then compsize = compsize - 1
If pct2 = 0 Then compsize = compsize - 1

For pct = 1 To thumbsize
For intcount = 1 To 5
If pct = pct1 Or pct = pct2 Or pct = pct3 Then dexv(intcount) = workvec(intcount, pct) + dexv(intcount)
Next
Next

dsm = 0



For pct = 1 To thumbsize - 1    '1 1
'Safe using tp3
For tp3 = pct + 1 To thumbsize
If sst(pct, 10) > 0 And sst(tp3, 10) > 0 And comparingpics(compsize + 1, thumbsize, pct, tp3, pct1, pct2, pct3) = True Then

MLR = sst(pct, 1)
MRL = sst(pct, 3)
TLR = sst(tp3, 1)
TRL = sst(tp3, 3)

    If MLR = 0 And TLR = 0 Then
    dsm = dsm + workvec(c1, pct) * workvec(c2, tp3) + workvec(c1, tp3) * workvec(c2, pct)
    Else
        If MLR = 1 And TLR = 0 Then
            If c1 = 1 Then
            dsm = dsm + workvec(1, pct) * workvec(c2, tp3)
            If MRL = 1 And c2 = 5 Then dsm = dsm + workvec(1, tp3) * workvec(5, pct)
            ElseIf MRL = 1 And c2 = 5 Then
            dsm = dsm + workvec(c1, tp3) * workvec(5, pct)
            End If
        ElseIf MLR = 0 And TLR = 1 Then
            If c1 = 1 Then
            dsm = dsm + workvec(1, tp3) * workvec(c2, pct)
            If TRL = 1 And c2 = 5 Then dsm = dsm + workvec(1, pct) * workvec(5, tp3)
            ElseIf TRL = 1 And c2 = 5 Then
            dsm = dsm + workvec(c1, pct) * workvec(5, tp3)
            End If
        ElseIf c1 = 1 And c2 = 5 Then
            If MRL = 1 And TRL = 1 Then
            dsm = dsm + workvec(1, pct) * workvec(5, tp3) + workvec(1, tp3) * workvec(5, pct)
            ElseIf TRL = 1 Then
            dsm = dsm + workvec(1, pct) * workvec(5, tp3)
            ElseIf MRL = 1 Then
            dsm = dsm + workvec(1, tp3) * workvec(5, pct)
            End If
    End If
    End If
End If
Next
Next



For pct = 1 To thumbsize     '1 X
If sst(pct, 10) > 0 And comparingpics(compsize, thumbsize, pct, pct1, pct2, pct3) = True Then
MRL = sst(pct, 3)


If sst(pct, 5) = 1 Then
dsm = dsm + workvec(c1, pct) * (24 - substotalvec(c2) - ssm(c2, pct1, pct2, pct3) - dexv(c2)) + (24 - dexv(c1) - substotalvec(c1) - ssm(c1, pct1, pct2, pct3)) * workvec(c2, pct)
Else
    If c1 = 1 Then
    
        If c2 = 5 And MRL = 1 Then
        tmp = 24 - dexv(5) - substotalvec(5) - ssm(5, pct1, pct2, pct3)
    
        'X P
        dsm = dsm + (24 - dexv(1) - substotalvec(1) - ssm(1, pct1, pct2, pct3)) * workvec(5, pct)
        Else
        tmp = 24 - dexv(c2) - workvec(c2, pct) - substotalvec(c2) - ssm(c2, pct1, pct2, pct3)
        End If
    
    'P X
    dsm = dsm + workvec(1, pct) * tmp
    
    'P P
    If c2 > 2 Then dsm = dsm + workvec(1, pct) * workvec(c2, pct)
    
    ElseIf c2 = 5 And MRL = 1 Then
    
    tmp = 24 - dexv(c1) - workvec(c1, pct) - substotalvec(c1) - ssm(c1, pct1, pct2, pct3)
    
    'X P
    dsm = dsm + tmp * workvec(5, pct)
    
    'P P
    If c1 < 4 Then dsm = dsm + workvec(c1, pct) * workvec(5, pct)
    
    End If

End If
End If
Next


For pct = 1 To thumbsize      '2
If sst(pct, 9) > 0 And comparingpics(compsize, thumbsize, pct, pct1, pct2, pct3) = True Then
If sst(pct, 1) = 0 Then
dsm = dsm + workvec(c1, pct) * workvec(c2, pct)
Else
If c1 = 1 And c2 = 2 Then dsm = dsm + workvec(1, pct) * workvec(2, pct)
If c1 = 4 And c2 = 5 And sst(pct, 3) = 1 Then dsm = dsm + workvec(4, pct) * workvec(5, pct)
End If
End If
Next


End Function
Public Function tsm(ct1 As Long, ct2 As Long, ct3 As Long, pct1 As Long, Optional pct2 As Long = 0) As Long
Dim c1 As Long, c2 As Long, c3 As Long, MMLR As Long, MMRL As Long, TTLR As Long, TTRL As Long, TT1LR As Long, TT1RL As Long
Dim Trst(5) As Long, Texv(5) As Long, compsize As Long

tsm = 0

If ct1 < ct2 Then
c1 = ct1
c2 = ct2
c3 = ct3
Else
c1 = ct3
c2 = ct2
c3 = ct1
End If


For intcount = 1 To 5
Texv(intcount) = 0
Next

For pct = 1 To thumbsize
For intcount = 1 To 5
If pct = pct1 Or pct = pct2 Then Texv(intcount) = workvec(intcount, pct) + Texv(intcount)
Next
Next

For intcount = 1 To 5
Trst(intcount) = 24 - ssm(intcount, pct1, pct2) - substotalvec(intcount) - Texv(intcount)
Next


compsize = 2

If pct2 = 0 Then compsize = compsize - 1


For mp = 1 To thumbsize      '3 - prize always > 0
If comparingpics(compsize, thumbsize, mp, pct1, pct2) = True Then

If sst(mp, 1) = 0 Then
tsm = tsm + workvec(c1, mp) * workvec(c2, mp) * workvec(c3, mp)
Else

'c1,c2,c3 for middles sorted
If c1 = 2 And c2 = 3 And c3 = 4 And sst(mp, 4) = 1 Then tsm = tsm + workvec(2, mp) * workvec(3, mp) * workvec(4, mp)
If c1 = 1 And c2 = 2 And c3 = 3 Then tsm = tsm + workvec(1, mp) * workvec(2, mp) * workvec(3, mp)
If c1 = 3 And c2 = 4 And c3 = 5 And sst(mp, 3) = 1 Then tsm = tsm + workvec(3, mp) * workvec(4, mp) * workvec(5, mp)
End If

End If
Next



For mp = 1 To thumbsize      '2 1

For tp = 1 To thumbsize
If sst(mp, 9) > 0 And sst(tp, 10) > 0 And comparingpics(compsize + 1, thumbsize, mp, tp, pct1, pct2) = True Then

MMLR = sst(mp, 1)
MMRL = sst(mp, 3)
TTLR = sst(tp, 1)
TTRL = sst(tp, 3)

If MMLR = 0 And TTLR = 0 Then

tsm = tsm + workvec(c1, tp) * workvec(c2, mp) * workvec(c3, mp)
tsm = tsm + workvec(c1, mp) * workvec(c2, tp) * workvec(c3, mp)
tsm = tsm + workvec(c1, mp) * workvec(c2, mp) * workvec(c3, tp)

Else

If MMLR = 1 And TTLR = 0 Then
If c1 = 1 And c2 = 2 Then tsm = tsm + workvec(1, mp) * workvec(2, mp) * workvec(c3, tp)
If c2 = 4 And c3 = 5 And MMRL = 1 Then tsm = tsm + workvec(c1, tp) * workvec(4, mp) * workvec(5, mp)

ElseIf MMLR = 0 And TTLR = 1 Then
If c1 = 1 Then tsm = tsm + workvec(c2, mp) * workvec(c3, mp) * workvec(1, tp)
If c3 = 5 And TTRL = 1 Then tsm = tsm + workvec(c1, mp) * workvec(c2, mp) * workvec(5, tp)
Else

If TTRL = 1 And c1 = 1 And c2 = 2 And c3 = 5 Then tsm = tsm + workvec(1, mp) * workvec(2, mp) * workvec(5, tp)
If MMRL = 1 And c1 = 1 And c2 = 4 And c3 = 5 Then tsm = tsm + workvec(1, tp) * workvec(4, mp) * workvec(5, mp)

End If

End If
End If
Next
Next


For mp = 1 To thumbsize      '2 X
mdpz = sst(mp, 9) 'md = main double
If mdpz > 0 And comparingpics(compsize, thumbsize, mp, pct1, pct2) = True Then
mspz = sst(mp, 10) 'ms = main single
MMLR = sst(mp, 1)
MMRL = sst(mp, 3)


If MMLR = 0 Then

If mspz = 0 Then
tsm = tsm + (Trst(c1) - workvec(c1, mp)) * workvec(c2, mp) * workvec(c3, mp)
tsm = tsm + workvec(c1, mp) * (Trst(c2) - workvec(c2, mp)) * workvec(c3, mp)
tsm = tsm + workvec(c1, mp) * workvec(c2, mp) * (Trst(c3) - workvec(c3, mp))
Else
tsm = tsm + Trst(c1) * workvec(c2, mp) * workvec(c3, mp)
tsm = tsm + workvec(c1, mp) * Trst(c2) * workvec(c3, mp)
tsm = tsm + workvec(c1, mp) * workvec(c2, mp) * Trst(c3)
End If


    
Else
    'LR or RL
        
    If c1 = 1 And c2 = 2 Then
        If mspz > 0 And MMRL = 1 And c3 = 5 Then
        tmp1 = Trst(5)
        Else
        tmp1 = Trst(c3) - workvec(c3, mp)
        End If
    'P P X
    tsm = tsm + workvec(1, mp) * workvec(2, mp) * tmp1
    'P P p
    If c3 > 3 Then tsm = tsm + workvec(1, mp) * workvec(2, mp) * workvec(c3, mp)
        
    ElseIf c2 = 4 And c3 = 5 And MMRL = 1 Then
        
        If mspz > 0 And c1 = 1 Then
        tmp1 = Trst(1)
        Else
        tmp1 = Trst(c1) - workvec(c1, mp)
        End If
        
    'X P P
    tsm = tsm + tmp1 * workvec(4, mp) * workvec(5, mp)
        
    'p P P
    If c1 < 3 Then tsm = tsm + workvec(c1, mp) * workvec(4, mp) * workvec(5, mp)
    End If
    
End If

End If
Next


For mp = 1 To thumbsize     '1 X X
If sst(mp, 10) > 0 And comparingpics(compsize, thumbsize, mp, pct1, pct2) = True Then
MMLR = sst(mp, 1)
MMRL = sst(mp, 3)

'Trst CANNOT be used

If MMLR = 0 Then
tsm = tsm + workvec(c1, mp) * ((24 - substotalvec(c2) - Texv(c2) - workvec(c2, mp)) * (24 - substotalvec(c3) - Texv(c3) - workvec(c3, mp)) - dsm(c2, c3, mp, pct1, pct2))
tsm = tsm + workvec(c2, mp) * ((24 - substotalvec(c1) - Texv(c1) - workvec(c1, mp)) * (24 - substotalvec(c3) - Texv(c3) - workvec(c3, mp)) - dsm(c1, c3, mp, pct1, pct2))
tsm = tsm + ((24 - substotalvec(c1) - Texv(c1) - workvec(c1, mp)) * (24 - substotalvec(c2) - Texv(c2) - workvec(c2, mp)) - dsm(c1, c2, mp, pct1, pct2)) * workvec(c3, mp)
Else

If c1 = 1 Then

'P X X
tsm = tsm + workvec(1, mp) * ((24 - substotalvec(c2) - Texv(c2) - workvec(c2, mp)) * (24 - substotalvec(c3) - Texv(c3) - workvec(c3, mp)) - dsm(c2, c3, mp, pct1, pct2))

'X X P
If c3 = 5 And MMRL = 1 Then tsm = tsm + ((24 - substotalvec(1) - Texv(1) - workvec(1, mp)) * (24 - substotalvec(c2) - Texv(c2) - workvec(c2, mp)) - dsm(1, c2, mp, pct1, pct2)) * workvec(5, mp)

    If c2 = 2 Then
    'P X P extras
    tsm = tsm + workvec(1, mp) * (Trst(2) - workvec(2, mp)) * workvec(c3, mp)
    
    'X P P
    If c3 = 5 And MMRL = 1 Then tsm = tsm + Trst(1) * workvec(2, mp) * workvec(5, mp)
    
    ElseIf (c2 = 4 And MMRL = 1) Then  'exclude pair at other end
    
    'P P X
    tsm = tsm + workvec(1, mp) * workvec(4, mp) * Trst(5)
    'P X P
    tsm = tsm + workvec(1, mp) * (Trst(4) - workvec(4, mp)) * workvec(5, mp)
    
    Else    'c2<>2,1 3 4, 1 3 5, 1 4 5 (mmrl=0)
    'P X P includes P P P
    tsm = tsm + workvec(1, mp) * Trst(c2) * workvec(c3, mp)
    
        If c3 = 5 And MMRL = 1 Then 'c2 : 3
        'P P X
        tsm = tsm + workvec(1, mp) * workvec(3, mp) * Trst(5)
        
        'X P P
        tsm = tsm + Trst(1) * workvec(3, mp) * workvec(5, mp)
        
        Else
        'P P X
        tsm = tsm + workvec(1, mp) * workvec(c2, mp) * (Trst(c3) - workvec(c3, mp))
        End If

    End If

ElseIf c3 = 5 And MMRL = 1 Then 'c1 <> 1
        
'X X P
tsm = tsm + ((24 - substotalvec(c1) - Texv(c1) - workvec(c1, mp)) * (24 - substotalvec(c2) - Texv(c2) - workvec(c2, mp)) - dsm(c1, c2, mp, pct1, pct2)) * workvec(5, mp)

    'P X P extras
    If c2 = 4 Then
    tsm = tsm + workvec(c1, mp) * (Trst(4) - workvec(4, mp)) * workvec(5, mp)
    Else
    'P X P includes P P P
    tsm = tsm + workvec(2, mp) * Trst(3) * workvec(5, mp)
    
    'X P P c1 <> 1, covered above
    tsm = tsm + (Trst(2) - workvec(2, mp)) * workvec(3, mp) * workvec(5, mp)
    
    End If


End If

End If
End If
Next


For mp = 1 To thumbsize - 1 '1 1 X   - can reverse prizes
For tp = mp + 1 To thumbsize
mspz = sst(mp, 10) 'ms = main single
aspz = sst(tp, 10) 'as = aux single
MMLR = sst(mp, 1)
MMRL = sst(mp, 3)
TTLR = sst(tp, 1)
TTRL = sst(tp, 3)
If mspz > 0 And aspz > 0 And comparingpics(compsize + 1, thumbsize, mp, tp, pct1, pct2) = True Then

'First get spslot moving, then 2 combos for each sum
If MMLR = 0 And TTLR = 0 Then
tsm = tsm + Trst(c1) * (workvec(c2, mp) * workvec(c3, tp) + workvec(c2, tp) * workvec(c3, mp))
tsm = tsm + Trst(c2) * (workvec(c1, mp) * workvec(c3, tp) + workvec(c1, tp) * workvec(c3, mp))
tsm = tsm + (workvec(c1, mp) * workvec(c2, tp) + workvec(c1, tp) * workvec(c2, mp)) * Trst(c3)

ElseIf MMLR = 1 And TTLR = 1 Then
'only count 2 since (at least 1) is fixed

If c1 = 1 And c3 = 5 Then
    If TTRL = 1 Then
        Select Case c2
        Case 2
        tsm = tsm + workvec(1, mp) * (Trst(2) - workvec(2, mp)) * workvec(5, tp)
        Case 4
        tsm = tsm + workvec(1, mp) * (Trst(4) - workvec(4, tp)) * workvec(5, tp)
        Case 3
        tsm = tsm + workvec(1, mp) * Trst(3) * workvec(5, tp)
        End Select
    End If
    If MMRL = 1 Then
        Select Case c2
        Case 2
        tsm = tsm + workvec(1, tp) * (Trst(2) - workvec(2, tp)) * workvec(5, mp)
        Case 4
        tsm = tsm + workvec(1, tp) * (Trst(4) - workvec(4, mp)) * workvec(5, mp)
        Case 3
        tsm = tsm + workvec(1, tp) * Trst(3) * workvec(5, mp)
        End Select
    End If
End If

ElseIf MMLR = 1 And TTLR = 0 Then

    If c1 = 1 Then
    
    'First calculate tmp1, tmp2
    tmp1 = Trst(c2) - workvec(c2, mp)
    If MMRL = 1 And c3 = 5 Then
    
    'X T P
    tsm = tsm + Trst(1) * workvec(c2, tp) * workvec(5, mp)
    
    'T X P
    tsm = tsm + workvec(1, tp) * tmp1 * workvec(5, mp)
    
    'T P P
    If c2 < 4 Then tsm = tsm + workvec(1, tp) * workvec(c2, mp) * workvec(5, mp)

    tmp2 = Trst(5)
    Else
    tmp2 = Trst(c3) - workvec(c3, mp)
    End If
    
    
    'P X T
    tsm = tsm + workvec(1, mp) * tmp1 * workvec(c3, tp)
    
    'P P T
    If c2 > 2 Then tsm = tsm + workvec(1, mp) * workvec(c2, mp) * workvec(c3, tp)
    
    'P T X
    tsm = tsm + workvec(1, mp) * workvec(c2, tp) * tmp2
    
    'P T P
    tsm = tsm + workvec(1, mp) * workvec(c2, tp) * workvec(c3, mp)
    
    ElseIf c3 = 5 And MMRL = 1 Then
    
    'First calculate tmp1, tmp2
    tmp1 = Trst(c1) - workvec(c1, mp)
    tmp2 = Trst(c2) - workvec(c2, mp)
    
    
    'T X P
    tsm = tsm + workvec(c1, tp) * tmp2 * workvec(5, mp)
    
    
    
    'T P P
    If c2 < 4 Then tsm = tsm + workvec(c1, tp) * workvec(c2, mp) * workvec(5, mp)
    
    'X T P
    tsm = tsm + tmp1 * workvec(c2, tp) * workvec(5, mp)
    
    'P T P
    tsm = tsm + workvec(c1, mp) * workvec(c2, tp) * workvec(5, mp)

    
    End If

Else    'mmLR = 0 And ttLR = 1
    
    If c1 = 1 Then
    
    'First calculate tmp1, tmp2
    tmp1 = Trst(c2) - workvec(c2, tp)
    If TTRL = 1 And c3 = 5 Then
    
    'X P T
    tsm = tsm + Trst(1) * workvec(c2, mp) * workvec(5, tp)
    
    'P X T
    tsm = tsm + workvec(1, mp) * tmp1 * workvec(5, tp)
    
    'P T T
    If c2 < 4 Then tsm = tsm + workvec(1, mp) * workvec(c2, tp) * workvec(5, tp)

    tmp2 = Trst(5)
    Else
    tmp2 = Trst(c3) - workvec(c3, tp)
    End If
    
    
    'T X P
    tsm = tsm + workvec(1, tp) * tmp1 * workvec(c3, mp)
    
    'T T P
    If c2 > 2 Then tsm = tsm + workvec(1, tp) * workvec(c2, tp) * workvec(c3, mp)
    
    'T P X
    tsm = tsm + workvec(1, tp) * workvec(c2, mp) * tmp2
    
    'T P T
    tsm = tsm + workvec(1, tp) * workvec(c2, mp) * workvec(c3, tp)
    
    ElseIf c3 = 5 And TTRL = 1 Then
    
    'First calculate tmp1, tmp2
    tmp1 = Trst(c1) - workvec(c1, tp)
    tmp2 = Trst(c2) - workvec(c2, tp)
    
    
    'P X T
    tsm = tsm + workvec(c1, mp) * tmp2 * workvec(5, tp)
    
    'P T T
    If c2 < 4 Then tsm = tsm + workvec(c1, mp) * workvec(c2, tp) * workvec(5, tp)
    
    'X P T
    tsm = tsm + tmp1 * workvec(c2, mp) * workvec(5, tp)
    
    'T P T
    tsm = tsm + workvec(c1, tp) * workvec(c2, mp) * workvec(5, tp)

    
    End If
    
End If  'mmlr,ttlr = 0


End If
Next
Next




For mp = 1 To thumbsize - 2 '1 1 1   - can reverse prizes
For tp = mp + 1 To thumbsize - 1
For tp1 = tp + 1 To thumbsize
mspz = sst(mp, 10) 'ms = main single
aspz = sst(tp, 10) 'as = aux single
aspz1 = sst(tp1, 10)
MMLR = sst(mp, 1)
MMRL = sst(mp, 3)
TTLR = sst(tp, 1)
TTRL = sst(tp, 3)
TT1LR = sst(tp1, 1)
TT1RL = sst(tp1, 3)
If MMLR = 1 And TTLR = 1 And TT1LR = 1 Then
'do nothing
Else
If mspz > 0 And aspz > 0 And aspz1 > 0 And comparingpics(compsize + 2, thumbsize, mp, tp, tp1, pct1, pct2) = True Then

If MMLR = 0 And TTLR = 0 And TT1LR = 0 Then

'3! cross product combinations
tsm = tsm + workvec(c1, mp) * (workvec(c2, tp) * workvec(c3, tp1) + workvec(c2, tp1) * workvec(c3, tp)) + workvec(c1, tp) * (workvec(c2, mp) * workvec(c3, tp1) + workvec(c2, tp1) * workvec(c3, mp)) + workvec(c1, tp1) * (workvec(c2, mp) * workvec(c3, tp) + workvec(c2, tp) * workvec(c3, mp))


Else
If (MMLR = 1 And TTLR = 1 And MMRL = 0 And TTRL = 0) Or (MMLR = 1 And TT1LR = 1 And MMRL = 0 And TT1RL = 0) Or (TTLR = 1 And TT1LR = 1 And TTRL = 0 And TT1RL = 0) Then
'do nothing!
Else

    If MMLR = 1 And TTLR = 0 And TT1LR = 0 Then
        If c1 = 1 Then tsm = tsm + workvec(1, mp) * (workvec(c2, tp) * workvec(c3, tp1) + workvec(c2, tp1) * workvec(c3, tp))
        If c3 = 5 And MMRL = 1 Then tsm = tsm + (workvec(c1, tp) * workvec(c2, tp1) + workvec(c1, tp1) * workvec(c2, tp)) * workvec(5, mp)

    ElseIf MMLR = 0 And TTLR = 1 And TT1LR = 0 Then
        If c1 = 1 Then tsm = tsm + workvec(1, tp) * (workvec(c2, mp) * workvec(c3, tp1) + workvec(c2, tp1) * workvec(c3, mp))
        If c3 = 5 And TTRL = 1 Then tsm = tsm + (workvec(c1, mp) * workvec(c2, tp1) + workvec(c1, tp1) * workvec(c2, mp)) * workvec(5, tp)
        
    ElseIf MMLR = 0 And TTLR = 0 And TT1LR = 1 Then
        If c1 = 1 Then tsm = tsm + workvec(1, tp1) * (workvec(c2, mp) * workvec(c3, tp) + workvec(c2, tp) * workvec(c3, mp))
        If c3 = 5 And TT1RL = 1 Then tsm = tsm + (workvec(c1, mp) * workvec(c2, tp) + workvec(c1, tp) * workvec(c2, mp)) * workvec(5, tp1)
        
    ElseIf MMLR = 1 And TTLR = 1 Then
        If c1 = 1 And c3 = 5 Then
        If TTRL = 1 Then tsm = tsm + workvec(1, mp) * workvec(c2, tp1) * workvec(5, tp)
        If MMRL = 1 Then tsm = tsm + workvec(1, tp) * workvec(c2, tp1) * workvec(5, mp)
        End If

    ElseIf MMLR = 1 And TT1LR = 1 Then
        If c1 = 1 And c3 = 5 Then
        If TT1RL = 1 Then tsm = tsm + workvec(1, mp) * workvec(c2, tp) * workvec(5, tp1)
        If MMRL = 1 Then tsm = tsm + workvec(1, tp1) * workvec(c2, tp) * workvec(5, mp)
        End If

    ElseIf TTLR = 1 And TT1LR = 1 Then
        If c1 = 1 And c3 = 5 Then
        If TT1RL = 1 Then tsm = tsm + workvec(1, tp) * workvec(c2, mp) * workvec(5, tp1)
        If TTRL = 1 Then tsm = tsm + workvec(1, tp1) * workvec(c2, mp) * workvec(5, tp)
        End If
    End If
End If
End If
End If
End If

Next
Next
Next

End Function
Public Function qsm(znopics As Long, pct1 As Long) As Long
Dim c1 As Long, c2 As Long, c3 As Long, c4 As Long, qmp As Long
Dim qtp As Long, qtp1 As Long, qtp2 As Long, rstns(5) As Long
Dim qaxpz As Long, qaxpz1 As Long, qaxpz2 As Long, qmspz As Long


testLR = 0
test1LR = 0
test2LR = 0
testRL = 0
test1RL = 0
test2RL = 0



qsm = 0
nopics = znopics

For c1 = 1 To 5
exv(c1) = 0
If c1 > nopics Then
reelduc(c1 - 1) = c1
ElseIf c1 < nopics Then
reelduc(c1) = c1
End If
Next


For pct = 1 To thumbsize
For c1 = 1 To 5
If pct = pct1 Then exv(c1) = workvec(c1, pct) + exv(c1)
Next
Next

For c1 = 1 To 5
rstns(c1) = 24 - substotalvec(c1) - exv(c1)
qrst(c1) = rstns(c1) - ssm(c1, pct1)
Next



'Summing FOURS

'Begin Main loop
For qmp = 1 To thumbsize

reelend = 5
Xmn = 1
qmspz = sst(qmp, 10) 'ms = main single
mainLR = sst(qmp, 1)
mainRL = sst(qmp, 3)
mainmid = sst(qmp, 4)


'First, 4s
If comparingpics(1, thumbsize, qmp, pct1) = True Then

If mainLR = 0 Or (mainLR = 1 And nopics = 5) Or (mainRL = 1 And nopics = 1) Then
For c1 = 1 To 5
If c1 <> nopics Then Xmn = workvec(c1, qmp) * Xmn
Next
qsm = qsm + Xmn
End If
End If

'next 3's with 1's

For qtp = 1 To thumbsize
qaxpz = sst(qtp, 10)
If qaxpz > 0 And comparingpics(2, thumbsize, qmp, qtp, pct1) = True Then
testLR = sst(qtp, 1)
testRL = sst(qtp, 3)

If sumbnds(1) = True Then

For ct = tryRL To RLloop

If ct = 2 Then
If sumbnds(1) = False Then Exit For
End If

For c1 = loopLRany1 To loopLRany2

If c1 <> nopics Then
    Xmn = 1
    For c2 = 1 To 5
        If c2 <> c1 And c2 <> nopics Then
        'main prize on this reel
        Xmn = workvec(c2, qmp) * Xmn
        ElseIf c2 <> nopics Then
        'aux prize on this reel
        Xmn = workvec(c2, qtp) * Xmn
        End If
    Next
qsm = qsm + Xmn
End If

Next
Next
End If

'middle threes
If mainmid = 1 Then
If nopics = 1 Then
If testRL = 1 Or testLR = 0 Then qsm = qsm + workvec(2, qmp) * workvec(3, qmp) * workvec(4, qmp) * workvec(5, qtp)
ElseIf nopics = 5 Then
qsm = qsm + workvec(2, qmp) * workvec(3, qmp) * workvec(4, qmp) * workvec(1, qtp)
End If
End If


End If
Next

'now for the 3's without the aux prizes

If comparingpics(1, thumbsize, qmp, pct1) = True Then
If sumbnds(0) = True Then

For ct = tryRL To RLloop

If ct = 2 Then
If sumbnds(0) = False Then Exit For
End If

For c1 = loopLRany1 To loopLRany2

Xmn = 1
Xrst = 1
If c1 <> nopics Then

    For c2 = 1 To 5
    If c2 <> c1 And c2 <> nopics Then
    'prize on this reel
    Xmn = workvec(c2, qmp) * Xmn
    ElseIf c2 <> nopics Then
    'nothing on this reel but need to discount single prizes
        If mainLR = 1 Then
        If qmspz > 0 Then
            If c2 = 1 Or (c2 = 5 And mainRL > 0) Then
            Xrst = qrst(c2)
            Else
            Xrst = qrst(c2) - workvec(c2, qmp)
            End If
        Else
        Xrst = qrst(c2) - workvec(c2, qmp)
        End If
        Else    'mainlr=0
        If qmspz > 0 Then
        Xrst = qrst(c2)
        Else
        Xrst = qrst(c2) - workvec(c2, qmp)
        End If
        End If
    End If
    Next
qsm = qsm + Xmn * Xrst

'P P P p extra nopics must be 4 (2)
If (nopics = 4 Or nopics = 2) And mainLR = 1 Then qsm = qsm + Xmn * workvec(c1, qmp)

End If

Next
Next
End If

'middles
If mainmid = 1 Then
    If nopics = 1 Then
        Xmn = workvec(2, qmp) * workvec(3, qmp) * workvec(4, qmp)
        
        'PPPp extra
        If mainRL = 0 Then qsm = qsm + Xmn * workvec(5, qmp)
        
        If qmspz > 0 And mainRL > 0 Then
        Xmn = Xmn * qrst(5)
        Else
        Xmn = Xmn * (qrst(5) - workvec(5, qmp))
        End If
        qsm = qsm + Xmn
    ElseIf nopics = 5 Then
        Xmn = workvec(2, qmp) * workvec(3, qmp) * workvec(4, qmp)
        If qmspz > 0 Then
        Xmn = Xmn * qrst(1)
        Else
        Xmn = Xmn * (qrst(1) - workvec(1, qmp))
        End If
        qsm = qsm + Xmn
    End If
End If

End If  'end 3X

'next 2's - first with 2's
'sst(qmp, 9) ALWAYS > 0 here
If sst(qmp, 9) > 0 Then

If qmp < thumbsize Then
For qtp = qmp + 1 To thumbsize
'don't want 3's or 4's again
qaxpz = sst(qtp, 9)
If qaxpz > 0 And comparingpics(2, thumbsize, qmp, qtp, pct1) = True Then
testLR = sst(qtp, 1)
testRL = sst(qtp, 3)

If sumbnds(2) = True Then

For ct = tryRL To RLloop

If ct = 2 Then
If sumbnds(2) = False Then Exit For
End If


For c1 = loopLRany1 To loopLRany2 Step partdir
If multptstat = True Then loopfix c1, loopLRany2, loopLRany3, loopLRany4
For c2 = loopLRany3 To loopLRany4 Step partdir
If c2 <> c1 And c2 <> nopics And c1 <> nopics Then
Xmn = 1
Xrst = 1

    For c3 = 1 To 5
        If c3 = c1 Or c3 = c2 Then
        'aux prize on this reel
        Xrst = workvec(c3, qtp) * Xrst
        ElseIf c3 <> nopics Then
        'main prize on this reel
        Xmn = workvec(c3, qmp) * Xmn
        End If
    Next
    qsm = qsm + Xmn * Xrst

End If
Next
Next
Next

End If
End If
Next
End If


'Now for 2's 1's 1's combinations (two loops to test)
For qtp = 1 To thumbsize - 1
For qtp1 = qtp + 1 To thumbsize
qaxpz = sst(qtp, 10)
qaxpz1 = sst(qtp1, 10)
If qaxpz > 0 And qaxpz1 > 0 And comparingpics(3, thumbsize, qmp, qtp, qtp1, pct1) = True Then
testLR = sst(qtp, 1)
testRL = sst(qtp, 3)
test1LR = sst(qtp1, 1)
test1RL = sst(qtp1, 3)

If sumbnds(3) = True Then

For ct = tryRL To RLloop

If ct = 2 Then
If sumbnds(3) = False Then Exit For
End If


For c1 = loopLRany1 To loopLRany2
For c2 = loopLRany3 To loopLRany4
If c2 <> c1 And c2 <> nopics And c1 <> nopics Then
Xmn = 1
Xrst = 1

'Tricky, the idea is to test the 2 possible locations for single prizes

    For c3 = 1 To 5
        If c3 = c1 Then
        'aux prize on this reel
        Xrst = workvec(c3, qtp) * Xrst
        ElseIf c3 = c2 Then
        Xrst = workvec(c3, qtp1) * Xrst
        ElseIf c3 <> nopics Then
        'main prize on this reel
        Xmn = workvec(c3, qmp) * Xmn
        End If
    Next
    qsm = qsm + Xmn * Xrst
End If
Next
Next
Next

End If
End If
Next
Next


'now 2's - with 1 X and X 1
For qtp = 1 To thumbsize
'don't want 4's or 5's again
qaxpz = sst(qtp, 10)
If qaxpz > 0 And comparingpics(2, thumbsize, qmp, qtp, pct1) = True Then
testLR = sst(qtp, 1)
testRL = sst(qtp, 3)

'In pairs, as before
If sumbnds(4) = True Then

For ct = tryRL To RLloop

If ct = 2 Then
If sumbnds(4) = False Then Exit For
End If

For c1 = loopLRany1 To loopLRany2
For c2 = loopLRany3 To loopLRany4
If c2 <> c1 And c2 <> nopics And c1 <> nopics Then
Xmn = 1

    For c3 = 1 To 5
        If c3 = c1 Then
        'aux prize on this reel
        Xrst = workvec(c3, qtp)
        ElseIf c3 <> c2 And c3 <> nopics Then
        'main prize on this reel
        Xmn = workvec(c3, qmp) * Xmn
        End If
    Next

    fp1 c2, qrst(c2), mainLR, testLR, mainRL, testRL, qmp, qtp, qmspz, qaxpz
    
    qsm = qsm + Xmn * Xrst * tmp
    
    
    If mainLR = 1 And testLR = 0 Then
    'P P T p, P P (N) p T
    If c2 <> singl1 Then qsm = qsm + Xmn * Xrst * workvec(c2, qmp)
    ElseIf mainLR = 0 And testLR = 1 Then
        'T P P t
        If c2 = reelend Then    'symmetry
        If visitextra = False Then qsm = qsm + Xmn * Xrst * workvec(c2, qtp)
        'T P t P
        ElseIf c2 <> singl1 Then
        qsm = qsm + Xmn * Xrst * workvec(c2, qtp)
        End If
    ElseIf mainLR = 1 And testLR = 1 Then
        'P P p T
        If nopics = 3 Then
        qsm = qsm + Xmn * Xrst * workvec(c2, qmp)
        'P P t T
        ElseIf c2 = 3 Then
        qsm = qsm + Xmn * Xrst * workvec(3, qtp)
        End If
    End If
 
End If
Next
Next

Next

End If
End If
Next


'Now for 2's - with X X

If comparingpics(1, thumbsize, qmp, pct1) = True Then
If sumbnds(5) = True Then

For ct = tryRL To RLloop


If ct = 2 Then
If sumbnds(5) = False Then Exit For
End If
For c1 = loopLRany1 To loopLRany2 Step partdir
If multptstat = True Then loopfix c1, loopLRany2, loopLRany3, loopLRany4
For c2 = loopLRany3 To loopLRany4 Step partdir
If c2 <> c1 And c2 <> nopics And c1 <> nopics Then
Xmn = 1
For c3 = 1 To 5
If c3 <> c1 And c3 <> c2 And c3 <> nopics Then Xmn = workvec(c3, qmp) * Xmn
Next
'dsum is through all combinations WITHOUT qmp
zsum = (rstns(c1) - workvec(c1, qmp)) * (rstns(c2) - workvec(c2, qmp))

qsm = qsm + Xmn * (zsum - dsm(c1, c2, qmp, pct1))

'Extras

If mainLR = 1 Then
    If c2 = reelend Then
        If nopics = 3 Then
        'PP pp
        If visitextra = False Then qsm = qsm + Xmn * workvec(c1, qmp) * workvec(c2, qmp)
        
        'PP x p
        qsm = qsm + Xmn * (qrst(c1) - workvec(c1, qmp)) * workvec(c2, qmp)
        
        'PP p x
        If qmspz > 0 And (mainRL = 1 Or c2 = 1) Then
        qsm = qsm + Xmn * workvec(c1, qmp) * qrst(c2)
        Else
        qsm = qsm + Xmn * workvec(c1, qmp) * (qrst(c2) - workvec(c2, qmp))
        End If
        
        Else    'nopics=4 (2)
        
        'PP x p
        qsm = qsm + Xmn * (qrst(c1) - workvec(c1, qmp)) * workvec(c2, qmp)
        End If
    Else    'c2 <> reelend
        'PP x p
        qsm = qsm + Xmn * (qrst(c1) - workvec(c1, qmp)) * workvec(c2, qmp)
     End If
End If



End If
Next
Next


Next
End If
End If

End If '2's prize condition


'Now for 1 1 1 1
If qmspz > 0 Then

If thumbsize > 4 And qmp < thumbsize - 2 Then
For qtp = qmp + 1 To thumbsize - 2
For qtp1 = qtp + 1 To thumbsize - 1
For qtp2 = qtp1 + 1 To thumbsize
qaxpz = sst(qtp, 10)
qaxpz1 = sst(qtp1, 10)
qaxpz2 = sst(qtp2, 10)

If qaxpz > 0 And qaxpz1 > 0 And qaxpz2 > 0 And comparingpics(4, thumbsize, qmp, qtp, qtp1, qtp2, pct1) = True Then
testLR = sst(qtp, 1)
testRL = sst(qtp, 3)
test1LR = sst(qtp1, 1)
test1RL = sst(qtp1, 3)
test2LR = sst(qtp2, 1)
test2RL = sst(qtp2, 3)

If sumbnds(6) = True Then

For ct = tryRL To RLloop

If ct = 2 Then
If sumbnds(6) = False Then Exit For
End If

For c1 = loopLRany1 To loopLRany2 Step partdir
For c2 = loopLRany3 To loopLRany4 Step partdir
For c3 = loopLRany5 To loopLRany6 Step partdir
If c2 <> c1 And c3 <> c1 And c3 <> c2 And c3 <> nopics And c2 <> nopics And c1 <> nopics Then
Xrst = 1

'Tricky, the idea is to test the 3 possible locations for single prizes

    For c4 = 1 To 5
    If c4 <> nopics Then
        'Sum through all 3 single aux prizes
        If c4 = c1 Then
        Xrst = workvec(c4, qtp) * Xrst
        ElseIf c4 = c2 Then
        Xrst = workvec(c4, qtp1) * Xrst
        ElseIf c4 = c3 Then
        Xrst = workvec(c4, qtp2) * Xrst
        Else
        'main prize on this reel
        Xmn = workvec(c4, qmp)
        End If
    End If
    Next
    qsm = qsm + Xmn * Xrst

End If
Next
Next
Next
Next

End If
End If
Next
Next
Next
End If 'dups


'Now for 1 1 1 X
If thumbsize > 3 And qmp < thumbsize - 1 Then
For qtp = qmp + 1 To thumbsize - 1
For qtp1 = qtp + 1 To thumbsize
qaxpz = sst(qtp, 10)
qaxpz1 = sst(qtp1, 10)
If qaxpz > 0 And qaxpz1 > 0 And comparingpics(3, thumbsize, qmp, qtp, qtp1, pct1) = True Then
testLR = sst(qtp, 1)
test1LR = sst(qtp1, 1)
testRL = sst(qtp, 3)
test1RL = sst(qtp1, 3)

If sumbnds(7) = True Then

For ct = tryRL To RLloop

If ct = 2 Then
If sumbnds(7) = False Then Exit For
End If

For c1 = loopLRany1 To loopLRany2 Step partdir
For c2 = loopLRany3 To loopLRany4 Step partdir
For c3 = loopLRany5 To loopLRany6 Step partdir
If c2 <> c1 And c3 <> c1 And c3 <> c2 And c3 <> nopics And c2 <> nopics And c1 <> nopics Then
Xrst = 1

    For c4 = 1 To 5
    If c4 <> nopics Then
        'Sum through all 3 single aux prizes
        If c4 = c1 Then
        Xrst = workvec(c4, qtp) * Xrst
        ElseIf c4 = c2 Then
        Xrst = workvec(c4, qtp1) * Xrst
        ElseIf c4 <> c3 Then
        'main prize on this reel
        Xmn = workvec(c4, qmp)
        End If
    End If
    Next
    
    
    If mainLR = 1 And testLR = 1 Then
    fp1 c3, qrst(c3), mainLR, testLR, mainRL, testRL, qmp, qtp, qmspz, qaxpz, qtp1, qaxpz1
    qsm = qsm + zsum
    
    ElseIf mainLR = 1 And test1LR = 1 Then
    fp1 c3, qrst(c3), mainLR, test1LR, mainRL, test1RL, qmp, qtp1, qmspz, qaxpz1, qtp, qaxpz
    qsm = qsm + zsum
    
    ElseIf testLR = 1 And test1LR = 1 Then
    fp1 c3, qrst(c3), testLR, test1LR, testRL, test1RL, qtp, qtp1, qaxpz, qaxpz1, qmp, qmspz
    qsm = qsm + zsum
    
    ElseIf mainLR = 1 Then
    fp1 c3, qrst(c3), mainLR, testLR, mainRL, testRL, qmp, qtp, qmspz, qaxpz
    'P T T1 p
    If c3 = reelend Then 'symmetry
    If visitextra = False Then qsm = qsm + Xmn * Xrst * workvec(c3, qmp)
    ElseIf c3 <> singl1 Then
    qsm = qsm + Xmn * Xrst * workvec(c3, qmp)
    End If
    
    ElseIf testLR = 1 Then
    fp1 c3, qrst(c3), testLR, test1LR, testRL, test1RL, qtp, qtp1, qaxpz, qaxpz1
    'T P T1 t
    If c3 = reelend Then
    If visitextra = False Then qsm = qsm + Xmn * Xrst * workvec(c3, qtp)
    ElseIf c3 <> singl1 Then
    qsm = qsm + Xmn * Xrst * workvec(c3, qtp)
    End If
    
    ElseIf test1LR = 1 Then
    fp1 c3, qrst(c3), test1LR, mainLR, test1RL, mainRL, qtp1, qmp, qaxpz1, qmspz
    'T1 P T t1
    If c3 = reelend Then
    If visitextra = False Then qsm = qsm + Xmn * Xrst * workvec(c3, qtp1)
    ElseIf c3 <> singl1 Then
    qsm = qsm + Xmn * Xrst * workvec(c3, qtp1)
    End If
    
    Else    'easy
    tmp = qrst(c3)
    End If
    

    qsm = qsm + Xmn * Xrst * tmp
    

End If
Next
Next
Next


Next
End If
End If 'dups
Next
Next
End If


'Now for 1 1 X X
If qmp < thumbsize Then
For qtp = qmp + 1 To thumbsize
qaxpz = sst(qtp, 10)
If qaxpz > 0 And comparingpics(2, thumbsize, qmp, qtp, pct1) = True Then
testLR = sst(qtp, 1)
testRL = sst(qtp, 3)


If sumbnds(8) = True Then

For ct = tryRL To RLloop


If ct = 2 Then
If sumbnds(8) = False Then Exit For
End If


For c1 = loopLRany1 To loopLRany2 Step partdir
If multptstat = True Then loopfix c1, loopLRany2, loopLRany3, loopLRany4
For c2 = loopLRany3 To loopLRany4 Step partdir
For c3 = loopLRany5 To loopLRany6 Step partdir
If c2 <> c1 And c3 <> c1 And c3 <> c2 And c3 <> nopics And c2 <> nopics And c1 <> nopics Then
Xrst = 1

For c4 = 1 To 5
    If c4 <> nopics Then
    'Don't include qmp, qtp in the X X portion either, but this is in dsum anyway
    If c4 = c3 Then
    'Single prize here
    Xrst = workvec(c4, qtp)
    ElseIf c4 <> c1 And c4 <> c2 Then
    'Main prize here
    Xmn = workvec(c4, qmp)
    End If
    End If
Next


zsum = (rstns(c1) - workvec(c1, qmp) - workvec(c1, qtp)) * (rstns(c2) - workvec(c2, qmp) - workvec(c2, qtp))


qsm = qsm + Xmn * Xrst * (zsum - dsm(c1, c2, qmp, qtp, pct1))


If mainLR = 1 And testLR = 0 Then

qsm = qsm + fp2(c1, c2, mainRL, 0, qmp, 0)

ElseIf mainLR = 0 And testLR = 1 Then

qsm = qsm + fp2(c1, c2, testRL, 0, qtp, 0)

ElseIf mainLR = 1 And testLR = 1 Then

    If mainRL = 1 And testRL = 0 Then
    qsm = qsm + fp2(c1, c2, testRL, mainRL, qtp, qmp, 6 - singl1, 6 - singl2)
    Else
    qsm = qsm + fp2(c1, c2, mainRL, testRL, qmp, qtp, singl1, singl2)
    End If

Else
'easy no extras
End If


End If

Next
Next
Next


Next

End If
End If
Next
End If 'dups


'Now for 1 X X X

If comparingpics(1, thumbsize, qmp, pct1) = True Then

If sumbnds(9) = True Then


For ct = tryRL To RLloop

If ct = 2 Then
If sumbnds(9) = False Then Exit For
End If


For c1 = loopLRany1 To loopLRany2 Step partdir
For c2 = loopLRany3 To loopLRany4 Step partdir
For c3 = loopLRany5 To loopLRany6 Step partdir
If c2 <> c1 And c3 <> c1 And c3 <> c2 And c3 <> nopics And c2 <> nopics And c1 <> nopics Then
Xmn = 1
For c4 = 1 To 5
    If c4 <> nopics Then
    If c4 <> c1 And c4 <> c2 And c4 <> c3 Then
    Xmn = workvec(c4, qmp)
    End If
    End If
Next

zsum = (rstns(c1) - workvec(c1, qmp)) * (rstns(c2) - workvec(c2, qmp)) * (rstns(c3) - workvec(c3, qmp))

'tsm is through all combinations WITHOUT qmp
qsm = qsm + Xmn * (zsum - tsm(c1, c2, c3, qmp, pct1))


If mainLR = 1 Then

If c3 = reelend Then

If nopics = singl1 Then
'All combos possible

'P x x p symmetrical
If visitextra = False Then qsm = qsm + Xmn * ((rstns(c1) - workvec(c1, qmp)) * (rstns(c2) - workvec(c2, qmp)) - dsm(c1, c2, qmp, pct1)) * workvec(c3, qmp)

'P x p x
qsm = qsm + Xmn * ((rstns(c1) - workvec(c1, qmp)) * (rstns(c3) - workvec(c3, qmp)) - dsm(c1, c3, qmp, pct1)) * workvec(c2, qmp)

'P p x x
qsm = qsm + Xmn * ((rstns(c2) - workvec(c2, qmp)) * (rstns(c3) - workvec(c3, qmp)) - dsm(c2, c3, qmp, pct1)) * workvec(c1, qmp)


'P x p p
If mainRL = 0 Then qsm = qsm + Xmn * (qrst(c1) - workvec(c1, qmp)) * workvec(c2, qmp) * workvec(c3, qmp)

'P p x p
If visitextra = False Then qsm = qsm + Xmn * (qrst(c2) - workvec(c2, qmp)) * workvec(c1, qmp) * workvec(c3, qmp)

'P p p x
If mainRL = 1 Or c3 = 1 Then
qsm = qsm + Xmn * qrst(c3) * workvec(c1, qmp) * workvec(c2, qmp)
Else
qsm = qsm + Xmn * (qrst(c3) - workvec(c3, qmp)) * workvec(c1, qmp) * workvec(c2, qmp)
End If

'P p p p
If mainRL = 0 Then qsm = qsm + Xmn * workvec(c1, qmp) * workvec(c2, qmp) * workvec(c3, qmp)

ElseIf nopics = 3 Then
'P x x p "symmetrical"
If visitextra = False Then qsm = qsm + Xmn * ((rstns(c1) - workvec(c1, qmp)) * (rstns(c2) - workvec(c2, qmp)) - dsm(c1, c2, qmp, pct1)) * workvec(c3, qmp)


'P x p x
qsm = qsm + Xmn * ((rstns(c1) - workvec(c1, qmp)) * (rstns(c3) - workvec(c3, qmp)) - dsm(c1, c3, qmp, pct1)) * workvec(c2, qmp)

'P x p p
If mainRL = 0 Then qsm = qsm + Xmn * (qrst(c1) - workvec(c1, qmp)) * workvec(c2, qmp) * workvec(c3, qmp)

Else    'nopics <> singl1 or <> 3, MUST be 4 (2 if ct = 2)

'P x x p
If visitextra = False Then qsm = qsm + Xmn * ((rstns(c1) - workvec(c1, qmp)) * (rstns(c2) - workvec(c2, qmp)) - dsm(c1, c2, qmp, pct1)) * workvec(c3, qmp)

'P x p x
qsm = qsm + Xmn * ((rstns(c1) - workvec(c1, qmp)) * (rstns(c3) - workvec(c3, qmp)) - dsm(c1, c3, qmp, pct1)) * workvec(c2, qmp)

'P x p p
If visitextra = False Then qsm = qsm + Xmn * (qrst(c1) - workvec(c1, qmp)) * workvec(c2, qmp) * workvec(c3, qmp)


End If


Else    'c3 <> reelend, nopics MUST be reelend

'P x x p
qsm = qsm + Xmn * ((rstns(c1) - workvec(c1, qmp)) * (rstns(c2) - workvec(c2, qmp)) - dsm(c1, c2, qmp, pct1)) * workvec(c3, qmp)

'P x p x
qsm = qsm + Xmn * ((rstns(c1) - workvec(c1, qmp)) * (rstns(c3) - workvec(c3, qmp)) - dsm(c1, c3, qmp, pct1)) * workvec(c2, qmp)

'P x p p
qsm = qsm + Xmn * (qrst(c1) - workvec(c1, qmp)) * workvec(c2, qmp) * workvec(c3, qmp)

End If
End If


End If
Next
Next
Next


Next
End If
End If
End If 'single prize


Next

End Function
Private Sub loopfix(basect As Long, upperct As Long, newctlow As Long, newctup As Long)
'Used for counting no-prize pairs, triples, quads AND prize pair doubles
'e.g. P P T T (both "any"). Other prize doubles of type Xmn are "missed".
'bc =1, upper=3
If partdir = 1 Then
If nopics = basect + 1 Then
newctlow = basect + 2
Else
newctlow = basect + 1
End If

If nopics = upperct + 1 Then
newctup = upperct + 2
Else
newctup = upperct + 1
End If

Else

'Reverse direction of loop
If nopics = basect - 1 Then
newctlow = basect - 2
Else
newctlow = basect - 1
End If
If nopics = upperct - 1 Then
newctup = upperct - 2
Else
newctup = upperct - 1
End If
End If

End Sub
Private Sub orddoublet(addme As Long)
'reminiscing
'If ct1 < ct2 Then
'dsum(pct, ct1, ct2) = dsum(pct, ct1, ct2) + addme
'Else
'dsum(pct, ct2, ct1) = dsum(pct, ct2, ct1) + addme
'End If
End Sub
Private Sub ordtriplet(addme As Long)
'reminiscing
'If ct1 < ct2 Then
'If ct1 < ct3 Then
'If ct2 < ct3 Then
'tsum(pct, ct1, ct2, ct3) = tsum(pct, ct1, ct2, ct3) + addme
'Else
'tsum(pct, ct1, ct3, ct2) = tsum(pct, ct1, ct3, ct2) + addme
'End If
'Else    '3 < 1
'tsum(pct, ct3, ct1, ct2) = tsum(pct, ct3, ct1, ct2) + addme
'End If
'Else
'If ct3 < ct1 Then
'If ct2 < ct3 Then
'tsum(pct, ct2, ct3, ct1) = tsum(pct, ct2, ct3, ct1) + addme
'Else
'tsum(pct, ct3, ct2, ct1) = tsum(pct, ct3, ct2, ct1) + addme
'End If
'Else    '1 < 3
'tsum(pct, ct2, ct1, ct3) = tsum(pct, ct2, ct1, ct3) + addme
'End If
'End If
End Sub
Private Sub fp1(spslot As Long, qval As Long, TT1LR As Long, TT2LR As Long, TT1RL As Long, TT2RL As Long, pic1 As Long, pic2 As Long, singT1pz As Long, singT2pz As Long, Optional pic3 As Long = 0, Optional singT3pz As Long = 0)

tmp = qval

If TT1LR = 1 And TT2LR = 0 Then
    
    If spslot = reelend Then
    If singT1pz = 0 Or TT1RL = 0 Then tmp = tmp - workvec(spslot, pic1)
    Else
    tmp = tmp - workvec(spslot, pic1)
    End If

    If singT2pz = 0 Then tmp = tmp - workvec(spslot, pic2)
ElseIf TT1LR = 0 And TT2LR = 1 Then
    If spslot = reelend Then
    If singT2pz = 0 Or TT2RL = 0 Then tmp = tmp - workvec(spslot, pic2)
    Else
    tmp = tmp - workvec(spslot, pic2)
    End If
    If singT1pz = 0 Then tmp = tmp - workvec(spslot, pic1)
ElseIf TT1LR = 1 And TT2LR = 1 Then
tmp = tmp - workvec(spslot, pic1) - workvec(spslot, pic2)


If pic3 > 0 Then    'process extras here
zsum = 0

If spslot <> singl1 Then zsum = Xmn * Xrst * workvec(spslot, pic1)
If spslot <> singl2 Then zsum = zsum + Xmn * Xrst * workvec(spslot, pic2)
End If

Else    'tt1LR & tt2LR = 0
    If singT1pz = 0 And singT2pz = 0 Then
    tmp = tmp - workvec(spslot, pic1) - workvec(spslot, pic2)
    ElseIf singT1pz = 0 Then
    tmp = tmp - workvec(spslot, pic1)
    ElseIf singT2pz = 0 Then
    tmp = tmp - workvec(spslot, pic2)
    End If
End If

End Sub
Public Function fp2(c1 As Long, c2 As Long, MMRL As Long, TTRL As Long, mnp As Long, tsp As Long, Optional S1 As Long, Optional S2 As Long) As Long

fp2 = 0

'All Valid for ct = 2

If tsp = 0 Then 'ttlr = 0


tmp = qrst(c1) - workvec(c1, mnp)
If (c2 = 5 And MMRL = 1) Or c2 = 1 Then
zsum = qrst(c2)
Else
zsum = qrst(c2) - workvec(c2, mnp)
End If


If c2 = reelend Then

'Nopics = singl1 Then 'P p T x, P p T p allowed
    
    If c1 = singl1 Then
    'P x T p
    If visitextra = False Then fp2 = fp2 + Xmn * Xrst * tmp * workvec(c2, mnp)
    
    Else    'c1 <> singl1
    
    'P N T x p, P N x T p, P T x p
    If visitextra = False Then fp2 = fp2 + Xmn * Xrst * tmp * workvec(c2, mnp)

    
    If c1 = 3 Then    'P N p T p, P T p N p
    If visitextra = False Then fp2 = fp2 + Xmn * Xrst * workvec(c1, mnp) * workvec(c2, mnp)
    Else  'P N T p p, P T N p p
    If MMRL = 0 Then fp2 = fp2 + Xmn * Xrst * workvec(c1, mnp) * workvec(c2, mnp)
    End If
    
    'P N T p x, P N p T x, P T N p x, P T p N x
    fp2 = fp2 + Xmn * Xrst * workvec(c1, mnp) * zsum
    
    End If

ElseIf nopics = reelend Then    'c2 <> reelend

    If c1 = singl1 Then
    'P x p T, P x T p
    fp2 = fp2 + Xmn * Xrst * tmp * workvec(c2, mnp)
    
    Else

    'NO RL condition required!
    'P T x p
    fp2 = fp2 + Xmn * Xrst * tmp * workvec(c2, mnp)

    'P T p p
    fp2 = fp2 + Xmn * Xrst * workvec(c1, mnp) * workvec(c2, mnp)

    'P T p x
    fp2 = fp2 + Xmn * Xrst * workvec(c1, mnp) * zsum

    End If

Else    'c2 <> reelend

    If nopics = singl1 Then
    'P N p x T
    fp2 = fp2 + Xmn * Xrst * workvec(c1, mnp) * zsum
    'P N p p T
    fp2 = fp2 + Xmn * Xrst * workvec(c1, mnp) * workvec(c2, mnp)
    End If

'P X p T
fp2 = fp2 + Xmn * Xrst * tmp * workvec(c2, mnp)

End If

Else    'mmlr, ttlr = 1


tmp = qrst(c1) - workvec(c1, mnp) - workvec(c1, tsp)
zsum = qrst(c2) - workvec(c2, mnp) - workvec(c2, tsp)

'P x p T
fp2 = fp2 + Xmn * Xrst * tmp * workvec(c2, mnp)

'P t x T
fp2 = fp2 + Xmn * Xrst * workvec(c1, tsp) * zsum

'P t p T
fp2 = fp2 + Xmn * Xrst * workvec(c1, tsp) * workvec(c2, mnp)

If nopics = S1 Then
'P p x T
fp2 = fp2 + Xmn * Xrst * workvec(c1, mnp) * zsum
'P p p T
fp2 = fp2 + Xmn * Xrst * workvec(c1, mnp) * workvec(c2, mnp)
End If

If nopics = S2 Then
'P x t T
fp2 = fp2 + Xmn * Xrst * tmp * workvec(c2, tsp)
'P t t T
fp2 = fp2 + Xmn * Xrst * workvec(c1, tsp) * workvec(c2, tsp)
End If


End If

End Function
Private Function sumbnds(prizecount As Long)


If RLloop = 2 Then

If tryRL = 1 Then
visitextra = True
singl1 = 6 - singl1
singl2 = 6 - singl2
End If

If multptstat = True Then
partdir = -1
loopLRany1 = 6 - loopLRany1
loopLRany2 = 6 - loopLRany2
loopLRany3 = 6 - loopLRany3
loopLRany4 = 6 - loopLRany4
loopLRany5 = 6 - loopLRany5
loopLRany6 = 6 - loopLRany6
Else
Dim temp As Long

temp = loopLRany1
loopLRany1 = 6 - loopLRany2
loopLRany2 = 6 - temp
temp = loopLRany3
loopLRany3 = 6 - loopLRany4
loopLRany4 = 6 - temp
temp = loopLRany5
loopLRany5 = 6 - loopLRany6
loopLRany6 = 6 - temp
End If

'Net effect is to reverse direction
If nopics > 3 Then
If loopLRany1 > 6 - nopics And loopLRany1 <= nopics Then loopLRany1 = loopLRany1 - 1
If loopLRany2 > 6 - nopics And loopLRany2 <= nopics Then loopLRany2 = loopLRany2 - 1
If loopLRany3 > 6 - nopics And loopLRany3 <= nopics Then loopLRany3 = loopLRany3 - 1
If loopLRany4 > 6 - nopics And loopLRany4 <= nopics Then loopLRany4 = loopLRany4 - 1
If loopLRany5 > 6 - nopics And loopLRany5 <= nopics Then loopLRany5 = loopLRany5 - 1
If loopLRany6 > 6 - nopics And loopLRany6 <= nopics Then loopLRany6 = loopLRany6 - 1
ElseIf nopics < 3 Then
If loopLRany1 < 6 - nopics And loopLRany1 >= nopics Then loopLRany1 = loopLRany1 + 1
If loopLRany2 < 6 - nopics And loopLRany2 >= nopics Then loopLRany2 = loopLRany2 + 1
If loopLRany3 < 6 - nopics And loopLRany3 >= nopics Then loopLRany3 = loopLRany3 + 1
If loopLRany4 < 6 - nopics And loopLRany4 >= nopics Then loopLRany4 = loopLRany4 + 1
If loopLRany5 < 6 - nopics And loopLRany5 >= nopics Then loopLRany5 = loopLRany5 + 1
If loopLRany6 < 6 - nopics And loopLRany6 >= nopics Then loopLRany6 = loopLRany6 + 1
End If

reelend = 1
RLloop = 1
sumbnds = True
Exit Function

Else
'Set defaults
loopLRany1 = 0
loopLRany2 = 0
loopLRany3 = 0
loopLRany4 = 0
loopLRany5 = 0
loopLRany6 = 0

RLloop = 1
tryRL = 1
partdir = 1
multptstat = False
sumbnds = False
visitextra = False
reelend = 5



Select Case prizecount

Case 0 '3 0 loopLRany1-2 map the no-prize X
If mainLR = 0 Then
loopLRany1 = reelduc(1)
loopLRany2 = reelduc(4)
Else
loopLRany1 = reelduc(4)
loopLRany2 = reelduc(4)
Select Case nopics
Case 1, 2
tryRL = 2
Rsgl 2
Case 3
tryRL = 10
Case 4, 5
tryRL = 0
Rsgl 4
End Select
If mainRL = 1 Then RLloop = 2
End If


Case 1 '3 1 loopLRany1-2 map single prize
If mainLR = 0 And testLR = 0 Then
loopLRany1 = reelduc(1)
loopLRany2 = reelduc(4)
sumbnds = True
Exit Function
ElseIf mainLR = 1 And testLR = 1 And mainRL = 0 And testRL = 0 Then Exit Function
ElseIf mainLR = 1 Then
    If testLR = 0 Then
    loopLRany1 = reelduc(4)
    loopLRany2 = reelduc(4)
    Select Case nopics
    Case 1, 2
    tryRL = 2
    Case 4, 5
    tryRL = 0
    Case 3
    tryRL = 10
    End Select
    ElseIf testRL = 1 Then
    loopLRany1 = reelduc(4)
    loopLRany2 = reelduc(4)
    Select Case nopics
    Case 4
    tryRL = 0
    Case 2
    tryRL = 2
    Case Else
    tryRL = 10
    End Select
    Else    'testlr=1
    loopLRany1 = reelduc(1)
    loopLRany2 = reelduc(1)
    'no RL as 0
    Select Case nopics
    Case 2
    tryRL = 0
    Case Else
    tryRL = 10
    End Select
    End If
Else    'mainlr=0
loopLRany1 = reelduc(1)
loopLRany2 = reelduc(1)
Select Case nopics
Case 1
tryRL = 2
Case 5
tryRL = 0
End Select
End If

If mainRL + (1 - mainLR) + testRL + (1 - testLR) >= 2 Then RLloop = 2




Case 2 '2 2 loopLRany1-4 maps the double aux prize
If mainLR = 0 And testLR = 0 Then
loopLRany1 = reelduc(1)
loopLRany2 = reelduc(3)
loopLRany3 = 0 'adjusted by loopfix later
loopLRany4 = 0 'adjusted by loopfix later
multptstat = True
sumbnds = True
Exit Function
ElseIf mainLR = 1 And testLR = 1 And mainRL = 0 And testRL = 0 Then Exit Function
ElseIf testLR = 1 Then
    If mainLR = 0 Then
    loopLRany1 = reelduc(1)
    loopLRany2 = reelduc(1)
    loopLRany3 = reelduc(2)
    loopLRany4 = reelduc(2)
    Select Case nopics
    Case 1, 2
    tryRL = 2
    Case 4, 5
    tryRL = 0
    End Select
    ElseIf mainRL = 1 Then
    loopLRany1 = reelduc(1)
    loopLRany2 = reelduc(1)
    loopLRany3 = reelduc(2)
    loopLRany4 = reelduc(2)
    If nopics <> 3 Then tryRL = 10
    Else
    loopLRany1 = reelduc(3)
    loopLRany2 = reelduc(3)
    loopLRany3 = reelduc(4)
    loopLRany4 = reelduc(4)
    If nopics <> 3 Then tryRL = 10
    End If
Else 'testlr=0
    loopLRany1 = reelduc(3)
    loopLRany2 = reelduc(3)
    loopLRany3 = reelduc(4)
    loopLRany4 = reelduc(4)
    Select Case nopics
    Case 1, 2
    tryRL = 2
    Case 4, 5
    tryRL = 0
    End Select
End If

If mainRL + (1 - mainLR) + testRL + (1 - testLR) >= 2 Then RLloop = 2



Case 3 '2 1 1 loopLRany1-4 map single prizes
If mainLR = 1 And testLR = 1 And test1LR = 1 Then Exit Function
If mainLR = 0 And testLR = 0 And test1LR = 0 Then
loopLRany1 = reelduc(1)
loopLRany2 = reelduc(4)
loopLRany3 = reelduc(1)
loopLRany4 = reelduc(4)
sumbnds = True
Exit Function
ElseIf mainLR = 1 And testLR = 1 And mainRL = 0 And testRL = 0 Then Exit Function
ElseIf mainLR = 1 And test1LR = 1 And mainRL = 0 And test1RL = 0 Then Exit Function
ElseIf testLR = 1 And test1LR = 1 And testRL = 0 And test1RL = 0 Then Exit Function
ElseIf mainLR = 1 Then
If testLR = 0 And test1LR = 0 Then
loopLRany1 = reelduc(3)
loopLRany2 = reelduc(4)
loopLRany3 = reelduc(3)
loopLRany4 = reelduc(4)
Select Case nopics
Case 1, 2
tryRL = 2
Case 4, 5
tryRL = 0
End Select
ElseIf test1RL = 1 Then 'tp is "any" by deduction
loopLRany1 = reelduc(3)
loopLRany2 = reelduc(3)
loopLRany3 = reelduc(4)
loopLRany4 = reelduc(4)
Select Case nopics
Case 2
tryRL = 2
Case 4
tryRL = 0
Case 1, 5
tryRL = 10
End Select
ElseIf test1LR = 1 Then
loopLRany1 = reelduc(2)
loopLRany2 = reelduc(2)
loopLRany3 = reelduc(1)
loopLRany4 = reelduc(1)
'no RL case as rL = 0!
Select Case nopics
Case 2
tryRL = 0
Case 1, 4, 5
tryRL = 10
End Select
ElseIf testRL = 1 Then
loopLRany1 = reelduc(4)
loopLRany2 = reelduc(4)
loopLRany3 = reelduc(3)
loopLRany4 = reelduc(3)
Select Case nopics
Case 2
tryRL = 2
Case 4
tryRL = 0
Case 1, 5
tryRL = 10
End Select
ElseIf testLR = 1 Then
loopLRany1 = reelduc(1)
loopLRany2 = reelduc(1)
loopLRany3 = reelduc(2)
loopLRany4 = reelduc(2)
Select Case nopics
Case 2
tryRL = 0
Case 1, 4, 5
tryRL = 10
End Select
End If

ElseIf testLR = 1 Then 'mainlr=0
    If test1LR = 0 Then
    loopLRany1 = reelduc(1)
    loopLRany2 = reelduc(1)
    loopLRany3 = reelduc(2)
    loopLRany4 = reelduc(4)
    Select Case nopics
    Case 1
    tryRL = 2
    Case 5
    tryRL = 0
    End Select
    ElseIf test1RL = 1 Then
    loopLRany1 = reelduc(1)
    loopLRany2 = reelduc(1)
    loopLRany3 = reelduc(4)
    loopLRany4 = reelduc(4)
    If nopics = 1 Or nopics = 5 Then tryRL = 10
    Else
    loopLRany1 = reelduc(4)
    loopLRany2 = reelduc(4)
    loopLRany3 = reelduc(1)
    loopLRany4 = reelduc(1)
    If nopics = 1 Or nopics = 5 Then tryRL = 10
    End If
ElseIf test1LR = 1 Then
loopLRany1 = reelduc(2)
loopLRany2 = reelduc(4)
loopLRany3 = reelduc(1)
loopLRany4 = reelduc(1)
Select Case nopics
Case 1
tryRL = 2
Case 5
tryRL = 0
End Select
End If
If mainRL + (1 - mainLR) + testRL + (1 - testLR) + test1RL + (1 - test1LR) >= 3 Then RLloop = 2


Case 4 '2 1 X loopLRany1-2 maps single prize, loopLRany3-4 maps spare
If mainLR = 0 And testLR = 0 Then
loopLRany1 = reelduc(1)
loopLRany2 = reelduc(4)
loopLRany3 = reelduc(1)
loopLRany4 = reelduc(4)
sumbnds = True
Exit Function
ElseIf mainLR = 1 And testLR = 1 And mainRL = 0 And testRL = 0 Then Exit Function
ElseIf mainLR = 1 Then
    If testLR = 0 Then
    loopLRany1 = reelduc(3)
    loopLRany2 = reelduc(4)
    loopLRany3 = reelduc(3)
    loopLRany4 = reelduc(4)
    Select Case nopics
    Case 1, 2
    tryRL = 2
    Case 4, 5
    tryRL = 0
    End Select
    Rsgl 3
    ElseIf testRL = 1 Then
    loopLRany1 = reelduc(4)
    loopLRany2 = reelduc(4)
    loopLRany3 = reelduc(3)
    loopLRany4 = reelduc(3)
    Select Case nopics
    Case 2
    tryRL = 2
    Rsgl 3, 2
    Case 4
    tryRL = 0
    Rsgl 3, 4
    Case 1, 5
    tryRL = 10
    Case 3
    Rsgl 3, 4
    End Select
    Else 'testrl = 0 .... mainrl = 1 but can't RLswploop
    loopLRany1 = reelduc(1)
    loopLRany2 = reelduc(1)
    loopLRany3 = reelduc(2)
    loopLRany4 = reelduc(2)
    Select Case nopics
    Case 2
    tryRL = 0
    Rsgl 3, 2
    Case 1, 4, 5
    tryRL = 10
    End Select
    End If
Else 'mainlr 0
loopLRany1 = reelduc(1)
loopLRany2 = reelduc(1)
loopLRany3 = reelduc(2)
loopLRany4 = reelduc(4)
Select Case nopics
Case 1
tryRL = 2
Rsgl 4
Case 5
tryRL = 0
Rsgl 2
Case Else
Rsgl 2
End Select
End If

If mainRL + (1 - mainLR) + testRL + (1 - testLR) >= 2 Then RLloop = 2


Case 5 '2 X X loopLRany1-4 maps the spare XX nonprize reels
If mainLR = 0 Then
loopLRany1 = reelduc(1)
loopLRany2 = reelduc(3)
loopLRany3 = 0        'see loopfix later
loopLRany4 = 0
multptstat = True
sumbnds = True
Exit Function
Else
loopLRany1 = reelduc(3)
loopLRany2 = reelduc(3)
loopLRany3 = reelduc(4)
loopLRany4 = reelduc(4)
Select Case nopics
Case 1, 2
tryRL = 2
Case 4, 5
tryRL = 0
End Select
Rsgl 3
'All other combos reversed
If mainRL = 1 Then RLloop = 2
End If


Case 6 '1 1 1 1 'loopLRany1-6 map out single prize
If mainLR = 1 And testLR = 1 And test1LR = 1 Then Exit Function
If mainLR = 1 And testLR = 1 And test2LR = 1 Then Exit Function
If testLR = 1 And test1LR = 1 And test2LR = 1 Then Exit Function
If mainLR = 0 And testLR = 0 And test1LR = 0 And test2LR = 0 Then
loopLRany1 = reelduc(1)
loopLRany2 = reelduc(4)
loopLRany3 = reelduc(1)
loopLRany4 = reelduc(4)
loopLRany5 = reelduc(1)
loopLRany6 = reelduc(4)
sumbnds = True
Exit Function
ElseIf mainLR = 1 And testLR = 1 And mainRL = 0 And testRL = 0 Then Exit Function
ElseIf mainLR = 1 And test1LR = 1 And mainRL = 0 And test1RL = 0 Then Exit Function
ElseIf mainLR = 1 And test2LR = 1 And mainRL = 0 And test2RL = 0 Then Exit Function
ElseIf testLR = 1 And test1LR = 1 And testRL = 0 And test1RL = 0 Then Exit Function
ElseIf testLR = 1 And test2LR = 1 And testRL = 0 And test2RL = 0 Then Exit Function
ElseIf test1LR = 1 And test2LR = 1 And test1RL = 0 And test2RL = 0 Then Exit Function
ElseIf mainLR = 1 Then
If testLR = 0 And test1LR = 0 Then
    If test2LR = 0 Then
    loopLRany1 = reelduc(2)
    loopLRany2 = reelduc(4)
    loopLRany3 = reelduc(2)
    loopLRany4 = reelduc(4)
    loopLRany5 = reelduc(2)
    loopLRany6 = reelduc(4)
    Select Case nopics
    Case 1
    tryRL = 2
    Case 5
    tryRL = 0
    End Select
    ElseIf test2RL = 1 Then
    loopLRany1 = reelduc(2)
    loopLRany2 = reelduc(3)
    loopLRany3 = reelduc(2)
    loopLRany4 = reelduc(3)
    loopLRany5 = reelduc(4)
    loopLRany6 = reelduc(4)
    If nopics = 1 Or nopics = 5 Then tryRL = 10
    Else
    loopLRany1 = reelduc(2)
    loopLRany2 = reelduc(3)
    loopLRany3 = reelduc(2)
    loopLRany4 = reelduc(3)
    loopLRany5 = reelduc(1)
    loopLRany6 = reelduc(1)
    If nopics = 1 Or nopics = 5 Then tryRL = 10

    End If

ElseIf test1RL = 1 Then 'testlr & test2lr  must = 0
loopLRany1 = reelduc(2)
loopLRany2 = reelduc(3)
loopLRany3 = reelduc(4)
loopLRany4 = reelduc(4)
loopLRany5 = reelduc(2)
loopLRany6 = reelduc(3)
If nopics = 1 Or nopics = 5 Then tryRL = 10
ElseIf test1LR = 1 Then
loopLRany1 = reelduc(2)
loopLRany2 = reelduc(3)
loopLRany3 = reelduc(1)
loopLRany4 = reelduc(1)
loopLRany5 = reelduc(2)
loopLRany6 = reelduc(3)
If nopics = 1 Or nopics = 5 Then tryRL = 10
ElseIf testRL = 1 Then 'again test1lr,test2lr = 0
loopLRany1 = reelduc(4)
loopLRany2 = reelduc(4)
loopLRany3 = reelduc(2)
loopLRany4 = reelduc(3)
loopLRany5 = reelduc(2)
loopLRany6 = reelduc(3)
If nopics = 1 Or nopics = 5 Then tryRL = 10
Else    'testLR = 1, testrl= 0 test2 cases covered above
loopLRany1 = reelduc(1)
loopLRany2 = reelduc(1)
loopLRany3 = reelduc(2)
loopLRany4 = reelduc(3)
loopLRany5 = reelduc(2)
loopLRany6 = reelduc(3)
If nopics = 1 Or nopics = 5 Then tryRL = 10
End If

ElseIf testLR = 1 Then 'mainlr=0 see (first if)
If test1LR = 0 And test2LR = 0 Then
loopLRany1 = reelduc(1)
loopLRany2 = reelduc(1)
loopLRany3 = reelduc(2)
loopLRany4 = reelduc(4)
loopLRany5 = reelduc(2)
loopLRany6 = reelduc(4)
Select Case nopics
Case 1
tryRL = 2
Case 5
tryRL = 0
End Select
ElseIf test2RL = 1 Then
loopLRany1 = reelduc(1)
loopLRany2 = reelduc(1)
loopLRany3 = reelduc(2)
loopLRany4 = reelduc(3)
loopLRany5 = reelduc(4)
loopLRany6 = reelduc(4)
If nopics = 1 Or nopics = 5 Then tryRL = 10
ElseIf test2LR = 1 Then
loopLRany1 = reelduc(4)
loopLRany2 = reelduc(4)
loopLRany3 = reelduc(2)
loopLRany4 = reelduc(3)
loopLRany5 = reelduc(1)
loopLRany6 = reelduc(1)
If nopics = 1 Or nopics = 5 Then tryRL = 10
ElseIf test1RL = 1 Then
loopLRany1 = reelduc(1)
loopLRany2 = reelduc(1)
loopLRany3 = reelduc(4)
loopLRany4 = reelduc(4)
loopLRany5 = reelduc(2)
loopLRany6 = reelduc(3)
If nopics = 1 Or nopics = 5 Then tryRL = 10
Else
loopLRany1 = reelduc(4)
loopLRany2 = reelduc(4)
loopLRany3 = reelduc(1)
loopLRany4 = reelduc(1)
loopLRany5 = reelduc(2)
loopLRany6 = reelduc(3)
If nopics = 1 Or nopics = 5 Then tryRL = 10
End If

ElseIf test1LR = 1 Then
'mainlr,testlr=0
If test2LR = 0 Then
loopLRany1 = reelduc(2)
loopLRany2 = reelduc(4)
loopLRany3 = reelduc(1)
loopLRany4 = reelduc(1)
loopLRany5 = reelduc(2)
loopLRany6 = reelduc(4)
Select Case nopics
Case 1
tryRL = 2
Case 5
tryRL = 0
End Select
ElseIf test2RL = 1 Then
loopLRany1 = reelduc(2)
loopLRany2 = reelduc(3)
loopLRany3 = reelduc(1)
loopLRany4 = reelduc(1)
loopLRany5 = reelduc(4)
loopLRany6 = reelduc(4)
If nopics = 1 Or nopics = 5 Then tryRL = 10
Else
loopLRany1 = reelduc(2)
loopLRany2 = reelduc(3)
loopLRany3 = reelduc(4)
loopLRany4 = reelduc(4)
loopLRany5 = reelduc(1)
loopLRany6 = reelduc(1)
If nopics = 1 Or nopics = 5 Then tryRL = 10
End If

ElseIf test2LR = 1 Then
'All previous three are "anys"
loopLRany1 = reelduc(2)
loopLRany2 = reelduc(4)
loopLRany3 = reelduc(2)
loopLRany4 = reelduc(4)
loopLRany5 = reelduc(1)
loopLRany6 = reelduc(1)
Select Case nopics
Case 1
tryRL = 2
Case 5
tryRL = 0
End Select
End If
If mainRL + (1 - mainLR) + testRL + (1 - testLR) + test1RL + (1 - test1LR) + test2RL + (1 - test2LR) >= 4 Then RLloop = 2


Case 7 '1 1 1 X  'loopLRany1-4 maps singles tp, tp1, loopLRany5-6 map non - scoring single

If mainLR = 1 And testLR = 1 And test1LR = 1 Then Exit Function
If mainLR = 0 And testLR = 0 And test1LR = 0 Then
loopLRany1 = reelduc(1)
loopLRany2 = reelduc(4)
loopLRany3 = reelduc(1)
loopLRany4 = reelduc(4)
loopLRany5 = reelduc(1)
loopLRany6 = reelduc(4)
sumbnds = True
Exit Function
ElseIf mainLR = 1 And testLR = 1 And mainRL = 0 And testRL = 0 Then Exit Function
ElseIf mainLR = 1 And test1LR = 1 And mainRL = 0 And test1RL = 0 Then Exit Function
ElseIf testLR = 1 And test1LR = 1 And testRL = 0 And test1RL = 0 Then Exit Function
ElseIf mainLR = 1 Then
If testLR = 0 Then
    If test1LR = 0 Then
    loopLRany1 = reelduc(2)
    loopLRany2 = reelduc(4)
    loopLRany3 = reelduc(2)
    loopLRany4 = reelduc(4)
    loopLRany5 = reelduc(2)
    loopLRany6 = reelduc(4)
    Select Case nopics
    Case 1
    tryRL = 2
    Rsgl 4
    Case 5
    tryRL = 0
    Rsgl 2
    Case Else
    Rsgl 2
    End Select
    ElseIf test1RL = 1 Then
    loopLRany1 = reelduc(2)
    loopLRany2 = reelduc(3)
    loopLRany3 = reelduc(4)
    loopLRany4 = reelduc(4)
    loopLRany5 = reelduc(2)
    loopLRany6 = reelduc(3)
    If nopics = 1 Or nopics = 5 Then tryRL = 10
    Rsgl 2, 4
    Else
    loopLRany1 = reelduc(2)
    loopLRany2 = reelduc(3)
    loopLRany3 = reelduc(1)
    loopLRany4 = reelduc(1)
    loopLRany5 = reelduc(2)
    loopLRany6 = reelduc(3)
    If nopics = 1 Or nopics = 5 Then tryRL = 10
    Rsgl 4, 2
    End If

ElseIf testRL = 1 Then
loopLRany1 = reelduc(4)
loopLRany2 = reelduc(4)
loopLRany3 = reelduc(2)
loopLRany4 = reelduc(3)
loopLRany5 = reelduc(2)
loopLRany6 = reelduc(3)
If nopics = 1 Or nopics = 5 Then tryRL = 10
Rsgl 2, 4
Else
loopLRany1 = reelduc(1)
loopLRany2 = reelduc(1)
loopLRany3 = reelduc(2)
loopLRany4 = reelduc(3)
loopLRany5 = reelduc(2)
loopLRany6 = reelduc(3)
If nopics = 1 Or nopics = 5 Then tryRL = 10
Rsgl 4, 2
End If
ElseIf testLR = 1 Then  'mainlr=0
    If test1LR = 0 Then
    loopLRany1 = reelduc(1)
    loopLRany2 = reelduc(1)
    loopLRany3 = reelduc(2)
    loopLRany4 = reelduc(4)
    loopLRany5 = reelduc(2)
    loopLRany6 = reelduc(4)
    Select Case nopics
    Case 1
    tryRL = 2
    Rsgl 4
    Case 5
    tryRL = 0
    Rsgl 2
    Case Else
    Rsgl 2
    End Select
    ElseIf test1RL = 1 Then
    loopLRany1 = reelduc(1)
    loopLRany2 = reelduc(1)
    loopLRany3 = reelduc(4)
    loopLRany4 = reelduc(4)
    loopLRany5 = reelduc(2)
    loopLRany6 = reelduc(3)
    If nopics = 1 Or nopics = 5 Then tryRL = 10
    Rsgl 2, 4
    Else
    loopLRany1 = reelduc(4)
    loopLRany2 = reelduc(4)
    loopLRany3 = reelduc(1)
    loopLRany4 = reelduc(1)
    loopLRany5 = reelduc(2)
    loopLRany6 = reelduc(3)
    If nopics = 1 Or nopics = 5 Then tryRL = 10
    Rsgl 4, 2
    End If
ElseIf test1LR = 1 Then
loopLRany1 = reelduc(2)
loopLRany2 = reelduc(4)
loopLRany3 = reelduc(1)
loopLRany4 = reelduc(1)
loopLRany5 = reelduc(2)
loopLRany6 = reelduc(4)
Select Case nopics
Case 1
tryRL = 2
Rsgl 4
Case 5
tryRL = 0
Rsgl 2
Case Else
Rsgl 2
End Select
End If

If mainRL + (1 - mainLR) + testRL + (1 - testLR) + test1RL + (1 - test1LR) >= 3 Then RLloop = 2


Case 8 '1 1 X X 'loopLRany1-4 map non - scoring double, loopLRany5-6 maps tp

If mainLR = 0 And testLR = 0 Then
loopLRany1 = reelduc(1)
loopLRany2 = reelduc(3)
loopLRany3 = 0
loopLRany4 = 0
loopLRany5 = reelduc(1)
loopLRany6 = reelduc(4)
multptstat = True
sumbnds = True
Exit Function
ElseIf mainLR = 1 And testLR = 1 And mainRL = 0 And testRL = 0 Then Exit Function
ElseIf mainLR = 1 Then
If testLR = 0 Then
loopLRany1 = reelduc(2)
loopLRany2 = reelduc(3)
loopLRany3 = 0
loopLRany4 = 0
loopLRany5 = reelduc(2)
loopLRany6 = reelduc(4)
multptstat = True
Select Case nopics
Case 1
tryRL = 2
Rsgl 4
Case 5
tryRL = 0
Rsgl 2
Case Else
Rsgl 2
End Select
ElseIf testRL = 1 Then
loopLRany1 = reelduc(2)
loopLRany2 = reelduc(2)
loopLRany3 = reelduc(3)
loopLRany4 = reelduc(3)
loopLRany5 = reelduc(4)
loopLRany6 = reelduc(4)
If nopics = 1 Or nopics = 5 Then tryRL = 10
Rsgl 2, 4
Else
loopLRany1 = reelduc(2)
loopLRany2 = reelduc(2)
loopLRany3 = reelduc(3)
loopLRany4 = reelduc(3)
loopLRany5 = reelduc(1)
loopLRany6 = reelduc(1)
If nopics = 1 Or nopics = 5 Then tryRL = 10
Rsgl 4, 2
End If
Else 'mainlr 0
loopLRany1 = reelduc(2)
loopLRany2 = reelduc(3)
loopLRany3 = 0
loopLRany4 = 0
loopLRany5 = reelduc(1)
loopLRany6 = reelduc(1)
multptstat = True
Select Case nopics
Case 1
tryRL = 2
Rsgl 4
Case 5
tryRL = 0
Rsgl 2
Case Else
Rsgl 2
End Select
End If

If mainRL + (1 - mainLR) + testRL + (1 - testLR) >= 2 Then RLloop = 2


Case 9 '1 X X X  'loopLRany1-6 map non -scoring triple

If mainLR = 0 Then
loopLRany1 = reelduc(1)
loopLRany2 = reelduc(2)
loopLRany3 = reelduc(2)
loopLRany4 = reelduc(3)
loopLRany5 = reelduc(3)
loopLRany6 = reelduc(4)
Else
loopLRany1 = reelduc(2)
loopLRany2 = reelduc(2)
loopLRany3 = reelduc(3)
loopLRany4 = reelduc(3)
loopLRany5 = reelduc(4)
loopLRany6 = reelduc(4)
Select Case nopics
Case 1
tryRL = 2
Rsgl 4
Case 5
tryRL = 0
Rsgl 2
Case Else
Rsgl 2
End Select
End If

If mainRL = 1 Then RLloop = 2


End Select

Select Case tryRL
Case 0
RLloop = 1    'overriding
tryRL = 1
Case 10
RLloop = 1
End Select

sumbnds = True
End If
End Function
Private Sub Rsgl(zsingl1 As Long, Optional zsingl2 As Long = 0)
'singl1, singl2 only of interest if a spare slot
'Reelend not relevant when both LR 1
singl1 = zsingl1
singl2 = zsingl2
End Sub
