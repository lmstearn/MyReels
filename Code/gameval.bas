Attribute VB_Name = "gameval"
Public singlsum(5) As Long, subsvec(14) As Long, probgtequalthan(10) As Long
Public multctstat As Boolean, VOGstepdir As Long
Public loopLRorany1 As Long, loopLRorany2 As Long, loopLRorany3 As Long
Public loopLRorany4 As Long, loopLRorany5 As Long, loopLRorany6 As Long
Public loopLRorany7 As Long, loopLRorany8 As Long
Public intRLloop As Long, mainp As Long, testp As Long, testp1 As Long, testp2 As Long, testp3 As Long
Private singl1 As Long, singl2 As Long, reelend As Long, Btotal As Long, betkeeptrakker As Long
Private ct As Long, ct1 As Long, ct2 As Long, ct3 As Long, ct4 As Long, ct5 As Long
Private Xrst As Long, Xrst1 As Long, Xmn As Long
Private tmp As Long, tmp1 As Long, tmp2 As Long, tmp3 As Long
Private VOG As Single, p0 As Single, moneybackVOG As Single, bestmonbackval As Single
Private bmax(5) As Long, substotalvec(5) As Long, pieN(15) As Single, betkeep(126, 15) As Long, betvec(15) As Long
Private thumbsize As Long, workvec(5, 14) As Long, VOGerror As Boolean
Public Sub calculatgamepercent(wheelvec() As Long, piccount As Long, Optional VOGerr As Boolean = False)
Dim nopics As Long, pnum As Long, qnum As Long
Dim axpz1 As Long, axpz2 As Long, axpz3 As Long
Dim mnpz As Long, mdpz As Long, mspz As Long
Dim axpz As Long, aspz As Long
Dim mainRL As Long, mainLR As Long, mainmid As Long
Dim testLR As Long, test1LR As Long, test2LR As Long
Dim testRL As Long, test1RL As Long, test2RL As Long
Dim sumtyperesult As Boolean, substwarn As Boolean, BLOCK111 As Boolean
Dim bettest(126, 15) As Long, Y As Single, ymax As Single, probRJ As Single, p0RJ As Single, scatcomp As Long, scatpz As Long

VOG = 0
intRLloop = 1
thumbsize = piccount
scatcomp = 0
scatpz = 0

zeroscatter ' zero scatter vars

For ct = 0 To 10
probgtequalthan(ct) = 0
Next

For ct = 1 To 5
substotalvec(ct) = 0
Next

For mainp = 1 To thumbsize
subsvec(mainp) = 0

For ct = 1 To thumbsize
If substitute(mainp, ct) = True Or substitute(ct, mainp) = True Then
subsvec(mainp) = mainp
Exit For
End If
Next

For ct = 0 To 4
If subsvec(mainp) = mainp Then substotalvec(ct) = substotalvec(ct) + wheelvec(ct, mainp)
Next


'Initialise workvec
For ct = 1 To 5
For ct1 = 1 To thumbsize
workvec(ct, ct1) = wheelvec(ct, ct1)
Next
Next

'Initialize intscatternumber and intscattervec, gametotalvec, substitutes & totals
Scatterinit wheelvec(1, mainp), mainp, False
Next


'Get partial sums
partials substotalvec, wheelvec, thumbsize


'Begin Main loop
For mainp = 1 To thumbsize

substwarn = False
Xmn = 1
Xrst = 1
mspz = sst(mainp, 10) 'ms = main single
mdpz = sst(mainp, 9) 'md = main double
mainLR = sst(mainp, 1)
mainRL = sst(mainp, 3)
mainmid = sst(mainp, 4)


For ct = 1 To thumbsize
'Mainp is substituter?
If mainp = subsvec(ct) Then substwarn = True
Next


If mainp <> intscattervec(1, 2) And mainp <> intscattervec(2, 2) Then  'no scatters


'First, Sum the 5s

If comparingpics(0, thumbsize, mainp) = True Then
For ct = 1 To 5
Xmn = workvec(ct, mainp) * Xmn
Next
Addvaluegame Xmn, sst(mainp, 6)
End If

'next 4's with 1's
mnpz = sst(mainp, 7)

For testp = 1 To thumbsize
axpz = sst(testp, 10)
If axpz > 0 And comparingpics(1, thumbsize, mainp, testp) = True Then
If sumtype(1) = True Then

For ct = 1 To intRLloop
If ct = 2 Then sumtyperesult = sumtype(1)

    For ct1 = loopLRorany1 To loopLRorany2
    Xmn = 1
    
    For ct2 = 1 To 5
        If ct2 <> ct1 Then
        'main prize on this reel
        Xmn = workvec(ct2, mainp) * Xmn
        Else
        'aux prize on this reel
        Xrst = workvec(ct2, testp)
        End If
    Next
    Addvaluegame Xmn * Xrst, axpz + mnpz
    Next

Next
End If
End If
Next


'now for the 4's without the aux prizes

If sumtype(0) = True And comparingpics(0, thumbsize, mainp) = True Then
For ct = 1 To intRLloop
If ct = 2 Then sumtyperesult = sumtype(0)

For ct1 = loopLRorany1 To loopLRorany2


Xmn = 1
Xrst = 1
    For ct2 = 1 To 5
    If ct2 <> ct1 Then
    'prize on this reel
    Xmn = workvec(ct2, mainp) * Xmn
    Else
        tmp = 24 - substotalvec(ct1) - singlsum(ct1)
        'nothing on this reel but need to discount single prizes
        If mainLR > 0 Then
            If mspz > 0 And (mainRL = 1 Or ct1 = 1) Then
            Xrst = tmp
            Else
            Xrst = tmp - workvec(ct1, mainp)
            End If
        Else
            If mspz > 0 Then
            Xrst = tmp
            Else
            Xrst = tmp - workvec(ct1, mainp)
            End If
        End If
    End If
    Next
    Addvaluegame Xmn * Xrst, mnpz

Next
Next
End If


doingtrips = True


'next 3's - first with 2's
mnpz = sst(mainp, 8) 'Note mnpz ALWAYS > 0 here

For testp = 1 To thumbsize
'don't want 4's or 5's again
axpz = sst(testp, 9)
If axpz > 0 And comparingpics(1, thumbsize, mainp, testp) = True Then
If sumtype(2) = True Then

For ct = 1 To intRLloop
If ct = 2 Then sumtyperesult = sumtype(2)
For ct1 = loopLRorany1 To loopLRorany2 Step VOGstepdir
If multctstat = True Then loopfix ct1, loopLRorany2, loopLRorany3, loopLRorany4
For ct2 = loopLRorany3 To loopLRorany4 Step VOGstepdir
If ct1 <> ct2 Then
Xmn = 1
Xrst = 1

    For ct3 = 1 To 5
        If ct3 = ct1 Or ct3 = ct2 Then
        'aux prize on this reel
        Xrst = workvec(ct3, testp) * Xrst
        Else
        'main prize on this reel
        Xmn = workvec(ct3, mainp) * Xmn
        End If
    Next
    Addvaluegame Xmn * Xrst, axpz + mnpz

End If
Next
Next
Next
End If
    
    'Middle threes
    If sst(testp, 5) = 1 And mainmid = 1 Then
    Xmn = workvec(2, mainp) * workvec(3, mainp) * workvec(4, mainp) * workvec(1, testp) * workvec(5, testp)
    Addvaluegame Xmn, axpz + mnpz
    End If

End If
Next


'Now for 3's 1's 1's combinations (two loops to test)
For testp = 1 To thumbsize - 1
For testp1 = testp + 1 To thumbsize
axpz = sst(testp, 10)
axpz1 = sst(testp1, 10)
If axpz > 0 And axpz1 > 0 And comparingpics(2, thumbsize, mainp, testp, testp1) = True Then
If sumtype(3) = True Then

For ct = 1 To intRLloop
If ct = 2 Then sumtyperesult = sumtype(3)
For ct1 = loopLRorany1 To loopLRorany2
For ct2 = loopLRorany3 To loopLRorany4
If ct1 <> ct2 Then
Xmn = 1
Xrst = 1
'Tricky, the idea is to test the 2 possible locations for single prizes

    For ct3 = 1 To 5
        If ct3 = ct1 Then
        'aux prize on this reel
        Xrst = workvec(ct3, testp) * Xrst
        ElseIf ct3 = ct2 Then
        Xrst = workvec(ct3, testp1) * Xrst
        Else
        'main prize on this reel
        Xmn = workvec(ct3, mainp) * Xmn
        End If
    Next
    Addvaluegame Xmn * Xrst, axpz + axpz1 + mnpz
End If
Next
Next
Next
End If

    'middle threes
    If mainmid = 1 Then
    testLR = sst(testp, 1)
    testRL = sst(testp, 3)
    test1LR = sst(testp1, 1)
    test1RL = sst(testp1, 3)
    Xmn = workvec(2, mainp) * workvec(3, mainp) * workvec(4, mainp)
    
    If test1RL = 0 And test1LR = 1 And testRL = 0 And testLR = 1 Then
    'do nothing
    ElseIf test1RL = 0 And test1LR = 1 Then
    Addvaluegame Xmn * workvec(1, testp1) * workvec(5, testp), axpz + axpz1 + mnpz
    ElseIf testRL = 0 And testLR = 1 Then
    Addvaluegame Xmn * workvec(1, testp) * workvec(5, testp1), axpz + axpz1 + mnpz
    Else
    '2 combos
    Addvaluegame Xmn * workvec(1, testp1) * workvec(5, testp), axpz + axpz1 + mnpz
    Addvaluegame Xmn * workvec(1, testp) * workvec(5, testp1), axpz + axpz1 + mnpz
    End If
    
    End If  'middles

End If
Next
Next


'now 3's - with 1 X and X 1
For testp = 1 To thumbsize
'don't want 4's or 5's again
axpz = sst(testp, 10)
If axpz > 0 And comparingpics(1, thumbsize, mainp, testp) = True Then
testLR = sst(testp, 1)
testRL = sst(testp, 3)
If sumtype(4, reelend, singl1, singl2) = True Then
'In pairs, as before


For ct = 1 To intRLloop
If ct = 2 Then sumtyperesult = sumtype(4, reelend, singl1, singl2)
For ct1 = loopLRorany1 To loopLRorany2
For ct2 = loopLRorany3 To loopLRorany4
If ct1 <> ct2 Then
Xmn = 1
Xrst = 1
    For ct3 = 1 To 5
        If ct3 = ct1 Then
        'aux prize on this reel
        Xrst = workvec(ct3, testp)
        ElseIf ct3 <> ct2 Then
        'main prize on this reel
        Xmn = workvec(ct3, mainp) * Xmn
        End If
    Next

    fixmult1 ct2, mainLR, testLR, mainRL, testRL, mainp, testp, mspz, axpz
    
    Addvaluegame Xmn * Xrst * tmp, axpz + mnpz
    
    'The following sums are for the EXTRA combinations e.g P P P T p where p can be a scoring mainp
    If mainLR = 1 And testLR = 0 Then
    If ct2 = reelend Then Addvaluegame Xmn * workvec(ct1, testp) * workvec(ct2, mainp), axpz + mnpz + mainRL * mspz
    
    'T P t P P, T P P t P, T P P P t; middle 3's PPP considered separately
    ElseIf mainLR = 0 And testLR = 1 Then
    If ct2 <> singl1 Then
    If ct2 = reelend Then
    If ct = 1 Then Addvaluegame Xmn * Xrst * workvec(ct2, testp), (1 + testRL) * axpz + mnpz
    Else
    Addvaluegame Xmn * Xrst * workvec(ct2, testp), axpz + mnpz
    End If
    End If
    Else
    'Easy no combos
    End If


End If
Next
Next
Next
End If

    'middle threes
    If mainmid = 1 Then
    Xmn = workvec(2, mainp) * workvec(3, mainp) * workvec(4, mainp)
    
    'T P P P P
    If mainRL = 0 Then Addvaluegame Xmn * workvec(1, testp) * workvec(5, mainp), axpz + mnpz
    
    
    If testLR = 1 Then
        If mainRL = 1 And mspz > 0 Then
        tmp = 24 - substotalvec(5) - singlsum(5)
        Else
        tmp = 24 - substotalvec(5) - workvec(5, mainp) - singlsum(5)
        End If
        If testRL = 0 Then tmp = tmp - workvec(5, testp)
        
        Addvaluegame workvec(1, testp) * Xmn * tmp, axpz + mnpz
        
        'Extra combo T P P P T
        Addvaluegame Xmn * workvec(1, testp) * workvec(5, testp), (1 + testRL) * axpz + mnpz
        
        
        If testRL = 1 Then
            If mspz > 0 Then
            tmp = 24 - substotalvec(1) - singlsum(1)
            Else
            tmp = 24 - substotalvec(1) - workvec(1, mainp) - singlsum(1)
            End If
            Addvaluegame Xmn * tmp * workvec(5, testp), axpz + mnpz
        End If
        
    
    Else 'testlr = 0
        If mainRL = 1 And mspz > 0 Then
        tmp = 24 - substotalvec(5) - singlsum(5)
        Else
        tmp = 24 - substotalvec(5) - workvec(5, mainp) - singlsum(5)
        End If
        Addvaluegame workvec(1, testp) * Xmn * tmp, axpz + mnpz
        If mspz > 0 Then
        tmp = 24 - substotalvec(1) - singlsum(1)
        Else
        tmp = 24 - substotalvec(1) - workvec(1, mainp) - singlsum(1)
        End If
        Addvaluegame Xmn * tmp * workvec(5, testp), axpz + mnpz
    End If
    
    End If  'middles


End If
Next


'Now for 3's - with X X
If sumtype(5, reelend) = True And comparingpics(0, thumbsize, mainp) = True Then
For ct = 1 To intRLloop
If ct = 2 Then sumtyperesult = sumtype(5, reelend)
For ct1 = loopLRorany1 To loopLRorany2 Step VOGstepdir
If multctstat = True Then loopfix ct1, loopLRorany2, loopLRorany3, loopLRorany4
For ct2 = loopLRorany3 To loopLRorany4 Step VOGstepdir

Xmn = 1

If ct1 <> ct2 Then
For ct3 = 1 To 5
    If ct3 <> ct1 And ct3 <> ct2 Then
    Xmn = workvec(ct3, mainp) * Xmn
    End If
Next

'dsum is through all combinations WITHOUT mainp

tmp1 = 24 - substotalvec(ct1) - workvec(ct1, mainp)
tmp2 = 24 - substotalvec(ct2) - workvec(ct2, mainp)


Addvaluegame Xmn * (tmp1 * tmp2 - dsm(ct1, ct2, mainp)), mnpz


'Extra combos
If mainLR = 1 Then
'Note workvec(mainp) is NOT in the appropriate singlsum below so subtract
Xrst = (tmp1 - singlsum(ct1)) * workvec(ct2, mainp)
Addvaluegame Xmn * Xrst, mnpz + mainRL * mspz
End If

End If
Next
Next
Next
End If

'middle threes
If mainmid = 1 Then
Xmn = workvec(2, mainp) * workvec(3, mainp) * workvec(4, mainp) * ((24 - substotalvec(1) - workvec(1, mainp)) * (24 - substotalvec(5) - workvec(5, mainp)) - dsm(1, 5, mainp))
Addvaluegame Xmn, mnpz
End If

doingtrips = False


'Summing through 2, 2, 1
mnpz = sst(mainp, 9)

If sst(mainp, 9) > 0 Then

If mainp < thumbsize Then
For testp = mainp + 1 To thumbsize  'no dups
For testp1 = 1 To thumbsize
axpz = sst(testp, 9)
axpz1 = sst(testp1, 10)
If axpz > 0 And axpz1 > 0 And comparingpics(2, thumbsize, mainp, testp, testp1) = True Then
If sumtype(6) = True Then

For ct = 1 To intRLloop
If ct = 2 Then sumtyperesult = sumtype(6)

For ct1 = loopLRorany1 To loopLRorany2 Step VOGstepdir
If multctstat = True Then loopfix ct1, loopLRorany2, loopLRorany3, loopLRorany4
For ct2 = loopLRorany3 To loopLRorany4 Step VOGstepdir
For ct3 = loopLRorany5 To loopLRorany6 Step VOGstepdir
If ct1 <> ct2 And ct3 <> ct1 And ct3 <> ct2 Then
Xmn = 1
Xrst = 1
    For ct4 = 1 To 5
    If ct4 = ct1 Or ct4 = ct2 Then
    Xrst = Xrst * workvec(ct4, testp)
    ElseIf ct4 = ct3 Then
    Xrst = Xrst * workvec(ct4, testp1)
    Else
    Xmn = workvec(ct4, mainp) * Xmn
    End If
    Next
    Addvaluegame Xmn * Xrst, axpz + axpz1 + mnpz

End If
Next
Next
Next
Next

End If
End If
Next
Next


'Summing through 2, 2, x (x cannot be a scoring triple component)
If mainp < thumbsize Then
For testp = mainp + 1 To thumbsize ' to avoid dups.
axpz = sst(testp, 9)
aspz = sst(testp, 10)
If axpz > 0 And comparingpics(1, thumbsize, mainp, testp) = True Then
testLR = sst(testp, 1)
testRL = sst(testp, 3)
If sumtype(7, reelend) = True Then

For ct = 1 To intRLloop
If ct = 2 Then sumtyperesult = sumtype(7, reelend)

For ct1 = loopLRorany1 To loopLRorany2 Step VOGstepdir
If multctstat = True Then loopfix ct1, loopLRorany2, loopLRorany3, loopLRorany4
For ct2 = loopLRorany3 To loopLRorany4 Step VOGstepdir
For ct3 = loopLRorany5 To loopLRorany6 Step VOGstepdir 'spslot
If ct1 <> ct2 And ct3 <> ct1 And ct3 <> ct2 Then

Xmn = 1
Xrst = 1
    For ct4 = 1 To 5
    If ct4 = ct1 Or ct4 = ct2 Then
    Xrst = workvec(ct4, testp) * Xrst
    ElseIf ct4 <> ct3 Then
    Xmn = workvec(ct4, mainp) * Xmn
    End If
    Next

    fixmult1 ct3, mainLR, testLR, mainRL, testRL, mainp, testp, mspz, aspz
    
    Addvaluegame Xmn * Xrst * tmp, axpz + mnpz

    If mainLR = 1 And testLR = 0 Then
        'P P T T p, p T T P P
        If ct3 = reelend Then
        Addvaluegame Xmn * Xrst * workvec(ct3, mainp), axpz + mnpz + mainRL * mspz
        ElseIf ct3 <> 3 Then
        'P P T p T, T p T P P
        Addvaluegame Xmn * Xrst * workvec(ct3, mainp), axpz + mnpz
        End If
    ElseIf mainLR = 0 And testLR = 1 Then
        'T T P P t, t P P T T
        If ct3 = reelend Then
        Addvaluegame Xmn * Xrst * workvec(ct3, testp), axpz + mnpz + testRL * aspz
        ElseIf ct3 <> 3 Then
        'T T P t P, P t P T T
        Addvaluegame Xmn * Xrst * workvec(ct3, testp), axpz + mnpz
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
End If

End If 'mainp dups limit


'Summing through 2 1 1 1
For testp = 1 To thumbsize - 2
For testp1 = testp + 1 To thumbsize - 1
For testp2 = testp1 + 1 To thumbsize
axpz = sst(testp, 10)
axpz1 = sst(testp1, 10)
axpz2 = sst(testp2, 10)
If axpz > 0 And axpz1 > 0 And axpz2 > 0 And comparingpics(3, thumbsize, mainp, testp, testp1, testp2) = True Then
If sumtype(8) = True Then

For ct = 1 To intRLloop
If ct = 2 Then sumtyperesult = sumtype(8)

For ct1 = loopLRorany1 To loopLRorany2
For ct2 = loopLRorany3 To loopLRorany4
For ct3 = loopLRorany5 To loopLRorany6
If ct1 <> ct2 And ct3 <> ct1 And ct3 <> ct2 Then
Xmn = 1
Xrst = 1
'Tricky, the idea is to test the 3 possible locations for single prizes

    For ct4 = 1 To 5
        'Sum through all 3 single aux prizes
        If ct4 = ct1 Then
        Xrst = workvec(ct4, testp) * Xrst
        ElseIf ct4 = ct2 Then
        Xrst = workvec(ct4, testp1) * Xrst
        ElseIf ct4 = ct3 Then
        Xrst = workvec(ct4, testp2) * Xrst
        Else
        'main prize on this reel
        Xmn = workvec(ct4, mainp) * Xmn
        End If
    Next
    Addvaluegame Xmn * Xrst, axpz + axpz1 + axpz2 + mnpz

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


'2 1 1 X
For testp = 1 To thumbsize - 1
For testp1 = testp + 1 To thumbsize
axpz = sst(testp, 10)
axpz1 = sst(testp1, 10)
If axpz > 0 And axpz1 > 0 And comparingpics(2, thumbsize, mainp, testp, testp1) = True Then
testLR = sst(testp, 1)
testRL = sst(testp, 3)
test1LR = sst(testp1, 1)
test1RL = sst(testp1, 3)
If sumtype(9, reelend, singl1, singl2) = True Then

For ct = 1 To intRLloop
If ct = 2 Then sumtyperesult = sumtype(9, reelend, singl1, singl2)

For ct1 = loopLRorany1 To loopLRorany2
For ct2 = loopLRorany3 To loopLRorany4
For ct3 = loopLRorany5 To loopLRorany6
If ct3 <> ct2 And ct3 <> ct1 And ct1 <> ct2 Then
Xmn = 1
Xrst = 1
    For ct4 = 1 To 5
        If ct4 = ct1 Then
        'aux prize on this reel
        Xrst = workvec(ct4, testp) * Xrst
        ElseIf ct4 = ct2 Then
        'aux prize on this reel
        Xrst = workvec(ct4, testp1) * Xrst
        ElseIf ct4 <> ct3 Then
        'main prize on this reel
        Xmn = workvec(ct4, mainp) * Xmn
        End If
    Next
        
    'extras included ct1:testp,ct2:testp1
    If mainLR = 1 And testLR = 1 Then
    fixmult1 ct3, mainLR, testLR, mainRL, testRL, mainp, testp, mspz, axpz, testp1, axpz1

    ElseIf mainLR = 1 And test1LR = 1 Then
    fixmult1 ct3, mainLR, test1LR, mainRL, test1RL, mainp, testp1, mspz, axpz1, testp, axpz

    ElseIf testLR = 1 And test1LR = 1 Then
    fixmult1 ct3, testLR, test1LR, testRL, test1RL, testp, testp1, axpz, axpz1, mainp, mspz
    
    ElseIf mainLR = 1 Then
    fixmult1 ct3, mainLR, testLR, mainRL, testRL, mainp, testp, mspz, axpz
    If ct3 = reelend Then
    'e.g P P T1 T2 X
    Addvaluegame Xmn * Xrst * workvec(ct3, mainp), axpz + axpz1 + mnpz + mainRL * mspz
    ElseIf ct3 <> singl1 Then
    'e.g P P T1 X T2
    Addvaluegame Xmn * Xrst * workvec(ct3, mainp), axpz + axpz1 + mnpz
    End If
    
    ElseIf testLR = 1 Then
    fixmult1 ct3, testLR, test1LR, testRL, test1RL, testp, testp1, axpz, axpz1
    If ct3 = reelend Then
    'e.g T1 P P T2 X
    If ct = 1 Then Addvaluegame Xmn * Xrst * workvec(5, testp), axpz1 + mnpz + (testRL + 1) * axpz
    ElseIf ct3 <> singl1 Then
    'e.g T1 P P X T2
    Addvaluegame Xmn * Xrst * workvec(ct3, testp), axpz + axpz1 + mnpz
    End If
    
    ElseIf test1LR = 1 Then
    fixmult1 ct3, test1LR, mainLR, test1RL, mainRL, testp1, mainp, axpz1, mspz
    If ct3 = reelend Then
    If ct = 1 Then Addvaluegame Xmn * Xrst * workvec(5, testp1), axpz + mnpz + (test1RL + 1) * axpz1
    ElseIf ct3 <> singl1 Then
    Addvaluegame Xmn * Xrst * workvec(ct3, testp1), axpz + axpz1 + mnpz
    End If
    
    Else
    'arbitrary for anys
    tmp = 24 - substotalvec(ct3) - singlsum(ct3)
    End If

    Addvaluegame Xmn * Xrst * tmp, axpz + axpz1 + mnpz

End If
Next
Next
Next
Next

End If
End If
Next
Next


'2 1 X X

For testp = 1 To thumbsize
axpz = sst(testp, 10)
If axpz > 0 And comparingpics(1, thumbsize, mainp, testp) = True Then
testLR = sst(testp, 1)
testRL = sst(testp, 3)
If sumtype(10, reelend, singl1) = True Then

For ct = 1 To intRLloop



If ct = 2 Then sumtyperesult = sumtype(10, reelend, singl1)
For ct1 = loopLRorany1 To loopLRorany2 Step VOGstepdir
If multctstat = True Then loopfix ct1, loopLRorany2, loopLRorany3, loopLRorany4
For ct2 = loopLRorany3 To loopLRorany4 Step VOGstepdir
For ct3 = loopLRorany5 To loopLRorany6 Step VOGstepdir
If ct3 <> ct2 And ct3 <> ct1 And ct1 <> ct2 Then
Xmn = 1
Xrst = 1

For ct4 = 1 To 5
    If ct4 = ct3 Then
    'Single prize here
    Xrst = workvec(ct4, testp)
    ElseIf ct4 <> ct1 And ct4 <> ct2 Then
    'Main prize here
    Xmn = workvec(ct4, mainp) * Xmn
    End If
Next
    
    tmp1 = 24 - substotalvec(ct1) - workvec(ct1, mainp) - workvec(ct1, testp)
    tmp2 = 24 - substotalvec(ct2) - workvec(ct2, mainp) - workvec(ct2, testp)
    
    
    
    Addvaluegame Xmn * (tmp1 * tmp2 - dsm(ct1, ct2, mainp, testp)) * Xrst, axpz + mnpz

If mainLR = 1 And testLR = 0 Then
    
    
    tmp1 = 24 - substotalvec(ct1) - singlsum(ct1) - workvec(ct1, mainp)
    
    If ct2 = reelend Then
    
        If mspz > 0 And (mainRL = 1 Or ct2 = 1) Then
        tmp2 = 24 - substotalvec(ct2) - singlsum(ct2)
        Else
        tmp2 = 24 - substotalvec(ct2) - singlsum(ct2) - workvec(ct2, mainp)
        End If
    
        
        If ct1 <> 3 Then
        'PPTpx
        Addvaluegame Xmn * Xrst * workvec(ct1, mainp) * tmp2, axpz + mnpz
        'Extra combo : PPTpp : only need these once
        If ct = 1 Then Addvaluegame Xmn * Xrst * workvec(4, mainp) * workvec(5, mainp), axpz + (1 + mainRL) * mnpz
        End If
        
        
        'Extra combos : 'PPTxp (ct3 = singl1) , PPxTp, (ct1 "precedes" ct2)
        Addvaluegame Xmn * Xrst * tmp1 * workvec(ct2, mainp), axpz + mnpz + mainRL * mspz
        
        
    Else    'ct2<>reelend
    
    'PPxpT
    Addvaluegame Xmn * Xrst * tmp1 * workvec(ct2, mainp), axpz + mnpz
    End If

    
ElseIf mainLR = 0 And testLR = 1 Then
 
 
    If mspz = 0 Then
    tmp1 = 24 - substotalvec(ct1) - singlsum(ct1) - workvec(ct1, mainp) - workvec(ct1, testp)
    tmp2 = 24 - substotalvec(ct2) - singlsum(ct2) - workvec(ct2, mainp)
    Else
    tmp1 = 24 - substotalvec(ct1) - singlsum(ct1) - workvec(ct1, testp)
    tmp2 = 24 - substotalvec(ct2) - singlsum(ct2)
    End If
        
        
    If ct2 <> reelend Then
        tmp2 = tmp2 - workvec(ct2, testp)

        'Note if testRL = 1 then ct can be 2; ok
            If ct1 <> singl1 Then
            'TPttP
            Addvaluegame Xmn * Xrst * workvec(ct1, testp) * workvec(ct2, testp), axpz + mnpz
            'TPtxP
            Addvaluegame Xmn * Xrst * workvec(ct1, testp) * tmp2, axpz + mnpz
            End If
            
            'TxtPP, TxPtP, TPxtP (ct1 <> singl1)
            Addvaluegame Xmn * Xrst * tmp1 * workvec(ct2, testp), axpz + mnpz
            
        Else    'ct2 = reelend
        
        If testRL = 0 And ct2 = 5 Then tmp2 = tmp2 - workvec(5, testp)
        
        'TPtPx, TPPtx
        If ct1 <> singl1 Then Addvaluegame Xmn * Xrst * workvec(ct1, testp) * tmp2, axpz + mnpz
        
        'Get the pair together TPPtt
        If testRL = 0 And ct1 = 4 Then Addvaluegame Xmn * Xrst * workvec(4, testp) * workvec(5, testp), axpz + mnpz
        
        
        If ct = 1 Then
        'testLR extra combos : TxPPt, TPxPt, TPPxt SYMMETRICAL  ;ct2=reelend
        Addvaluegame Xmn * Xrst * tmp1 * workvec(5, testp), (1 + testRL) * axpz + mnpz
        
        'TPtPt
        If ct1 = 3 Then Addvaluegame Xmn * Xrst * workvec(3, testp) * workvec(5, testp), (1 + testRL) * axpz + mnpz
        
        End If

        
    End If
ElseIf mainLR = 1 And testLR = 1 Then
    
'Not concerned with singl1


    tmp1 = tmp1 - singlsum(ct1)
    tmp2 = tmp2 - singlsum(ct2)

    
    If mainRL = 1 Then
    If testRL = 0 Or ct = 2 Then
    'TpxPP
    Addvaluegame Xmn * Xrst * workvec(2, mainp) * tmp2, axpz + mnpz
    'TptPP "cross" product
    Addvaluegame Xmn * Xrst * workvec(2, mainp) * workvec(3, testp), axpz + mnpz
    'TxtPP
    Addvaluegame Xmn * Xrst * tmp1 * workvec(3, testp), axpz + mnpz
    End If
    End If
    
    If testRL = 1 Then
    If mainRL = 0 Or ct = 1 Then
    'PPxpT
    Addvaluegame Xmn * Xrst * tmp1 * workvec(4, mainp), axpz + mnpz
    'PPtpT
    Addvaluegame Xmn * Xrst * workvec(3, testp) * workvec(4, mainp), axpz + mnpz
    'PPtxT
    Addvaluegame Xmn * Xrst * workvec(3, testp) * tmp2, axpz + mnpz
    End If
    End If

Else


'Anys ;easy ... no extra combos
End If

  
End If
Next
Next
Next
Next

End If
End If
Next


'2 X X X
If sumtype(11, reelend) = True And comparingpics(0, thumbsize, mainp) = True Then
For ct = 1 To intRLloop
If ct = 2 Then sumtyperesult = sumtype(11, reelend)
For ct1 = loopLRorany1 To loopLRorany2 Step VOGstepdir
If multctstat = True Then loopfix ct1, loopLRorany2, loopLRorany3, loopLRorany4
For ct2 = loopLRorany3 To loopLRorany4 Step VOGstepdir
If multctstat = True Then loopfix ct2, loopLRorany4, loopLRorany5, loopLRorany6
For ct3 = loopLRorany5 To loopLRorany6 Step VOGstepdir
If ct1 <> ct2 Then
Xmn = 1

For ct4 = 1 To 5
    Select Case ct4
    Case ct1, ct2, ct3
    Case Else
    Xmn = workvec(ct4, mainp) * Xmn
    End Select
Next

tmp1 = 24 - substotalvec(ct1) - workvec(ct1, mainp)
tmp2 = 24 - substotalvec(ct2) - workvec(ct2, mainp)
tmp3 = 24 - substotalvec(ct3) - workvec(ct3, mainp)

Addvaluegame Xmn * (tmp1 * tmp2 * tmp3 - tsm(ct1, ct2, ct3, mainp)), mnpz


If mainLR = 1 Then


'PPxpx
Addvaluegame Xmn * workvec(ct2, mainp) * (tmp1 * tmp3 - dsm(3, ct3, mainp)), mnpz

'PPxxp
Addvaluegame Xmn * workvec(ct3, mainp) * (tmp1 * tmp2 - dsm(3, ct2, mainp)), mnpz + mainRL * mspz

'PPxpp symmetrical
If ct = 1 Then Addvaluegame Xmn * workvec(4, mainp) * workvec(5, mainp) * (tmp1 - singlsum(3)), (1 + mainRL) * mnpz

End If


End If
Next
Next
Next
Next

End If


End If '2's prize condition



'Now for 1 1 1 1 1
If mspz > 0 Then

If thumbsize > 4 And mainp < thumbsize - 3 Then
For testp = mainp + 1 To thumbsize - 3
For testp1 = testp + 1 To thumbsize - 2
For testp2 = testp1 + 1 To thumbsize - 1
For testp3 = testp2 + 1 To thumbsize
axpz = sst(testp, 10)
axpz1 = sst(testp1, 10)
axpz2 = sst(testp2, 10)
axpz3 = sst(testp3, 10)
If axpz > 0 And axpz1 > 0 And axpz2 > 0 And axpz3 > 0 And comparingpics(4, thumbsize, mainp, testp, testp1, testp2, testp3) = True Then
If sumtype(12) Then

For ct = 1 To intRLloop
If ct = 2 Then sumtyperesult = sumtype(12)

For ct1 = loopLRorany1 To loopLRorany2
For ct2 = loopLRorany3 To loopLRorany4
For ct3 = loopLRorany5 To loopLRorany6
For ct4 = loopLRorany7 To loopLRorany8
If ct1 <> ct2 And ct3 <> ct1 And ct3 <> ct2 And ct4 <> ct1 And ct4 <> ct2 And ct4 <> ct3 Then
Xrst = 1

'Tricky, the idea is to test the 3 possible locations for single prizes

    For ct5 = 1 To 5
        'Sum through all 3 single aux prizes
        If ct5 = ct1 Then
        Xrst = workvec(ct5, testp) * Xrst
        ElseIf ct5 = ct2 Then
        Xrst = workvec(ct5, testp1) * Xrst
        ElseIf ct5 = ct3 Then
        Xrst = workvec(ct5, testp2) * Xrst
        ElseIf ct5 = ct4 Then
        Xrst = workvec(ct5, testp3) * Xrst
        Else
        'main prize on this reel
        Xmn = workvec(ct5, mainp)
        End If
    Next
    Addvaluegame Xmn * Xrst, axpz + axpz1 + axpz2 + axpz3 + mspz

End If
Next
Next
Next
Next
Next

End If
End If 'dups
Next
Next
Next
Next
End If


'Now for 1 1 1 1 X
If thumbsize > 3 And mainp < thumbsize - 2 Then
For testp = mainp + 1 To thumbsize - 2
For testp1 = testp + 1 To thumbsize - 1
For testp2 = testp1 + 1 To thumbsize
axpz = sst(testp, 10)
axpz1 = sst(testp1, 10)
axpz2 = sst(testp2, 10)
If axpz > 0 And axpz1 > 0 And axpz2 > 0 And comparingpics(3, thumbsize, mainp, testp, testp1, testp2) = True Then
testLR = sst(testp, 1)
test1LR = sst(testp1, 1)
test2LR = sst(testp2, 1)
testRL = sst(testp, 3)
test1RL = sst(testp1, 3)
test2RL = sst(testp2, 3)
If sumtype(13, reelend, singl1, singl2) = True Then

For ct = 1 To intRLloop
If ct = 2 Then sumtyperesult = sumtype(13, reelend, singl1, singl2)

For ct1 = loopLRorany1 To loopLRorany2
For ct2 = loopLRorany3 To loopLRorany4
For ct3 = loopLRorany5 To loopLRorany6
For ct4 = loopLRorany7 To loopLRorany8
If ct1 <> ct2 And ct3 <> ct1 And ct3 <> ct2 And ct4 <> ct1 And ct4 <> ct2 And ct4 <> ct3 Then
Xrst = 1

    For ct5 = 1 To 5
        'Sum through all 3 single aux prizes
        If ct5 = ct1 Then
        Xrst = workvec(ct5, testp) * Xrst
        ElseIf ct5 = ct2 Then
        Xrst = workvec(ct5, testp1) * Xrst
        ElseIf ct5 = ct3 Then
        Xrst = workvec(ct5, testp2) * Xrst
        ElseIf ct5 <> ct4 Then
        'main prize on this reel
        Xmn = workvec(ct5, mainp)
        End If
    Next
    
    
    'Extra combos included
    If mainLR = 1 And testLR = 1 Then
    fixmult1 ct4, mainLR, testLR, mainRL, testRL, mainp, testp, mspz, axpz, testp1, axpz1, testp2, axpz2
    
    ElseIf mainLR = 1 And test1LR = 1 Then
    fixmult1 ct4, mainLR, test1LR, mainRL, test1RL, mainp, testp1, mspz, axpz1, testp, axpz, testp2, axpz2
    
    ElseIf mainLR = 1 And test2LR = 1 Then
    fixmult1 ct4, mainLR, test2LR, mainRL, test2RL, mainp, testp2, mspz, axpz2, testp, axpz, testp1, axpz2
    
    ElseIf testLR = 1 And test1LR = 1 Then
    fixmult1 ct4, testLR, test1LR, testRL, test1RL, testp, testp1, axpz, axpz1, mainp, mspz, testp2, axpz2
    
    ElseIf testLR = 1 And test2LR = 1 Then
    fixmult1 ct4, testLR, test2LR, testRL, test2RL, testp, testp2, axpz, axpz2, mainp, mspz, testp1, axpz1
    
    ElseIf test1LR = 1 And test2LR = 1 Then
    fixmult1 ct4, test1LR, test2LR, test1RL, test2RL, testp1, testp2, axpz1, axpz2, mainp, mspz, testp, axpz
    
    ElseIf mainLR = 1 Then
    fixmult1 ct4, mainLR, testLR, mainRL, testRL, mainp, testp, mspz, axpz
    If axpz2 = 0 Then tmp = tmp - workvec(ct4, testp2)
    If ct4 = reelend Then
    If ct = 1 Then Addvaluegame Xmn * Xrst * workvec(ct4, mainp), axpz + axpz1 + axpz2 + (mainRL + 1) * mspz
    ElseIf ct4 <> singl1 Then
    Addvaluegame Xmn * Xrst * workvec(ct4, mainp), axpz + axpz1 + axpz2 + mspz
    End If
    
    ElseIf testLR = 1 Then
    fixmult1 ct4, testLR, test1LR, testRL, test1RL, testp, testp1, axpz, axpz1
    If axpz2 = 0 Then tmp = tmp - workvec(ct4, testp2)
    If ct4 = reelend Then
    If ct = 1 Then Addvaluegame Xmn * Xrst * workvec(ct4, testp), axpz1 + axpz2 + mspz + (testRL + 1) * axpz
    ElseIf ct4 <> singl1 Then
    Addvaluegame Xmn * Xrst * workvec(ct4, testp), axpz + axpz1 + axpz2 + mspz
    End If
    
    ElseIf test1LR = 1 Then
    fixmult1 ct4, test1LR, test2LR, test1RL, test2RL, testp1, testp2, axpz1, axpz2
    If axpz = 0 Then tmp = tmp - workvec(ct4, testp)
    If ct4 = reelend Then
    If ct = 1 Then Addvaluegame Xmn * Xrst * workvec(ct4, testp1), axpz + axpz2 + mspz + (test1RL + 1) * axpz1
    ElseIf ct4 <> singl1 Then
    Addvaluegame Xmn * Xrst * workvec(ct4, testp1), axpz + axpz1 + axpz2 + mspz
    End If
    
    ElseIf test2LR = 1 Then
    fixmult1 ct4, test2LR, test1LR, test2RL, test1RL, testp2, testp1, axpz2, axpz1
    If axpz = 0 Then tmp = tmp - workvec(ct4, testp)
    If ct4 = reelend Then
    If ct = 1 Then Addvaluegame Xmn * Xrst * workvec(ct4, testp2), axpz + axpz1 + mspz + (test2RL + 1) * axpz2
    ElseIf ct4 <> singl1 Then
    Addvaluegame Xmn * Xrst * workvec(ct4, testp2), axpz + axpz1 + axpz2 + mspz
    End If
        
    Else    'easy
    tmp = 24 - substotalvec(ct4) - singlsum(ct4)
    End If

    Addvaluegame Xmn * Xrst * tmp, axpz + axpz1 + axpz2 + mspz
    


End If
Next
Next
Next
Next
Next

End If
End If 'dups
Next
Next
Next
End If


'Now for 1 1 1 X X
If mainp < thumbsize - 1 Then
For testp = mainp + 1 To thumbsize - 1
For testp1 = testp + 1 To thumbsize
axpz = sst(testp, 10)
axpz1 = sst(testp1, 10)
If axpz > 0 And axpz1 > 0 And comparingpics(2, thumbsize, mainp, testp, testp1) = True Then
testLR = sst(testp, 1)
test1LR = sst(testp1, 1)
testRL = sst(testp, 3)
test1RL = sst(testp1, 3)

BLOCK111 = False    'blocks double n/s pair from moving into mainp LR prize
If sumtype(14, reelend, singl1, singl2, BLOCK111) = True Then
For ct = 1 To intRLloop

If ct = 2 Then sumtyperesult = sumtype(14, reelend, singl1, singl2, BLOCK111)

For ct1 = loopLRorany1 To loopLRorany2 Step VOGstepdir
If multctstat = True Then loopfix ct1, loopLRorany2, loopLRorany3, loopLRorany4
If BLOCK111 = True Then
If loopLRorany4 = 5 Then loopLRorany4 = 4
If loopLRorany4 = 1 Then loopLRorany4 = 2
End If

For ct2 = loopLRorany3 To loopLRorany4 Step VOGstepdir
For ct3 = loopLRorany5 To loopLRorany6 Step VOGstepdir
For ct4 = loopLRorany7 To loopLRorany8 Step VOGstepdir
If ct1 <> ct2 And ct3 <> ct1 And ct3 <> ct2 And ct4 <> ct1 And ct4 <> ct2 And ct4 <> ct3 Then
Xrst = 1
For ct5 = 1 To 5

    'Don't include mainp, testp, testp1 in the X X portion either
    If ct5 = ct3 Then
    'Single prize here
    Xrst = workvec(ct5, testp) * Xrst
    ElseIf ct5 = ct4 Then
   'Single prize here
    Xrst = workvec(ct5, testp1) * Xrst
    ElseIf ct5 <> ct1 And ct5 <> ct2 Then
    'Main prize here
    Xmn = workvec(ct5, mainp)
    End If
Next


tmp = (24 - workvec(ct1, mainp) - workvec(ct1, testp) - workvec(ct1, testp1)) * (24 - workvec(ct2, mainp) - workvec(ct2, testp) - workvec(ct2, testp1))


Addvaluegame Xmn * (tmp - dsm(ct1, ct2, mainp, testp, testp1)) * Xrst, axpz + axpz1 + mspz
  

If mainLR = 1 And testLR = 0 And test1LR = 0 Then
fixmult2 mainp, 0, mainRL, testRL, mspz, axpz, axpz1
ElseIf mainLR = 0 And testLR = 1 And test1LR = 0 Then
fixmult2 testp, 0, testRL, mainRL, axpz, mspz, axpz1
ElseIf mainLR = 0 And testLR = 0 And test1LR = 1 Then
fixmult2 testp1, 0, test1RL, mainRL, axpz1, mspz, axpz
ElseIf mainLR = 1 And testLR = 1 Then

If testRL = 0 Then
fixmult2 testp, mainp, testRL, mainRL, axpz, mspz, axpz1, 6 - singl1, 6 - singl2
Else
fixmult2 mainp, testp, mainRL, testRL, mspz, axpz, axpz1, singl1, singl2
End If

ElseIf mainLR = 1 And test1LR = 1 Then


If test1RL = 0 Then
fixmult2 testp1, mainp, test1RL, mainRL, axpz1, mspz, axpz, 6 - singl1, 6 - singl2
Else
fixmult2 mainp, testp1, mainRL, test1RL, mspz, axpz1, axpz, singl1, singl2
End If

ElseIf testLR = 1 And test1LR = 1 Then


If test1RL = 0 Then
fixmult2 testp1, testp, test1RL, testRL, axpz1, axpz, mspz, 6 - singl1, 6 - singl2
Else
fixmult2 testp, testp1, testRL, test1RL, axpz, axpz1, mspz, singl1, singl2
End If

Else
'Easy; there are no extra combos
End If
    


End If

Next
Next
Next
Next
Next

End If
End If
Next
Next
End If 'dups


'Now for 1 1 X X X
If mainp < thumbsize Then
For testp = mainp + 1 To thumbsize
axpz = sst(testp, 10)
If axpz > 0 And comparingpics(1, thumbsize, mainp, testp) = True Then
testLR = sst(testp, 1)
testRL = sst(testp, 3)

If sumtype(15, reelend, singl1) = True Then
For ct = 1 To intRLloop


If ct = 2 Then sumtyperesult = sumtype(15, reelend, singl1)

For ct1 = loopLRorany1 To loopLRorany2 Step VOGstepdir
If multctstat = True Then loopfix ct1, loopLRorany2, loopLRorany3, loopLRorany4
For ct2 = loopLRorany3 To loopLRorany4 Step VOGstepdir
If multctstat = True Then loopfix ct2, loopLRorany4, loopLRorany5, loopLRorany6
For ct3 = loopLRorany5 To loopLRorany6 Step VOGstepdir
For ct4 = loopLRorany7 To loopLRorany8 Step VOGstepdir
If ct1 <> ct2 And ct3 <> ct1 And ct3 <> ct2 And ct4 <> ct1 And ct4 <> ct2 And ct4 <> ct3 Then
For ct5 = 1 To 5
    If ct5 = ct4 Then
   'Single prize here
    Xrst = workvec(ct5, testp)
    ElseIf ct5 <> ct1 And ct5 <> ct2 And ct5 <> ct3 Then
    Xmn = workvec(ct5, mainp)
    End If
Next

tmp = 24 - substotalvec(ct1) - workvec(ct1, mainp) - workvec(ct1, testp)
tmp1 = 24 - substotalvec(ct2) - workvec(ct2, mainp) - workvec(ct2, testp)
tmp2 = 24 - substotalvec(ct3) - workvec(ct3, mainp) - workvec(ct3, testp)


Addvaluegame Xmn * Xrst * (tmp * tmp1 * tmp2 - tsm(ct1, ct2, ct3, mainp, testp)), axpz + mspz

If mainLR = 1 And testLR = 0 Then
fixmult3 False, mainp, testp, mainRL, 0, mspz, axpz
ElseIf mainLR = 0 And testLR = 1 Then
fixmult3 False, testp, mainp, testRL, 0, axpz, mspz
ElseIf mainLR = 1 And testLR = 1 Then


'reversible don't need singl1 etc
If testRL = 0 Then
fixmult3 True, testp, mainp, testRL, mainRL, axpz, mspz
Else
fixmult3 True, mainp, testp, mainRL, testRL, mspz, axpz
End If

Else
'Easy no extra combos
End If




End If
Next
Next
Next
Next
Next

End If
End If 'dups
Next
End If



'Now for 1 X X X X

If sumtype(16) = True And comparingpics(0, thumbsize, mainp) = True Then
For ct = 1 To intRLloop
If ct = 2 Then sumtyperesult = sumtype(16)
For ct1 = loopLRorany1 To loopLRorany2 Step VOGstepdir
For ct2 = loopLRorany3 To loopLRorany4 Step VOGstepdir
For ct3 = loopLRorany5 To loopLRorany6 Step VOGstepdir
For ct4 = loopLRorany7 To loopLRorany8 Step VOGstepdir
If ct1 <> ct2 And ct3 <> ct1 And ct3 <> ct2 And ct4 <> ct1 And ct4 <> ct2 And ct4 <> ct3 Then
For ct5 = 1 To 5
    If ct5 <> ct1 And ct5 <> ct2 And ct5 <> ct3 And ct5 <> ct4 Then
    'Single prize here
    nopics = ct5
    Xmn = workvec(ct5, mainp)
    End If
Next

tmp = 24 - substotalvec(ct1) - workvec(ct1, mainp)
tmp1 = 24 - substotalvec(ct2) - workvec(ct2, mainp)
tmp2 = 24 - substotalvec(ct3) - workvec(ct3, mainp)
tmp3 = 24 - substotalvec(ct4) - workvec(ct4, mainp)


'tsum is through all combinations WITHOUT mainp
Xrst = tmp * tmp1 * tmp2 * tmp3 - qsm(nopics, mainp)

Addvaluegame Xmn * Xrst, mspz


If mainLR = 1 Then
If ct = 1 Then
'P x x x p
Addvaluegame Xmn * workvec(5, mainp) * (tmp * tmp1 * tmp2 - tsm(2, 3, 4, mainp)), (1 + mainRL) * mspz
'P x p x p
Addvaluegame Xmn * workvec(3, mainp) * workvec(5, mainp) * (tmp * tmp2 - dsm(2, 4, mainp)), (1 + mainRL) * mspz
End If


'P x x p x
Addvaluegame Xmn * workvec(ct3, mainp) * (tmp * tmp1 * tmp3 - tsm(ct1, ct2, ct4, mainp)), mspz

'P x p x x
Addvaluegame Xmn * workvec(ct2, mainp) * (tmp * tmp2 * tmp3 - tsm(ct1, ct3, ct4, mainp)), mspz

'P x p p x
Addvaluegame Xmn * workvec(ct2, mainp) * workvec(ct3, mainp) * (tmp * tmp3 - dsm(ct1, ct4, mainp)), mspz

If mainRL = 0 Then
'P x x p p
Addvaluegame Xmn * workvec(4, mainp) * workvec(5, mainp) * (tmp * tmp1 - dsm(2, 3, mainp)), mspz

'P x p p p
Addvaluegame Xmn * (tmp - singlsum(2)) * workvec(3, mainp) * workvec(4, mainp) * workvec(5, mainp), mspz

End If
End If


End If
Next
Next
Next
Next
Next

End If

End If 'mspz > 0



ElseIf substwarn = False Then ' now process scatters


pnum = 3 * workvec(1, mainp)
qnum = 24 - pnum
If sst(mainp, 5) = 1 Then 'anys
For ct = 2 To 5
tmp = 0

Select Case ct
Case 2
For ct1 = 1 To 3
For ct2 = ct1 + 1 To 4
For ct3 = ct2 + 1 To 5
tmp = tmp + tsm(ct1, ct2, ct3, mainp)
Next
Next
Next
Case 3
For ct1 = 1 To 4
For ct2 = ct1 + 1 To 5
tmp = tmp + dsm(ct1, ct2, mainp)
Next
Next
Case 4
For ct1 = 1 To 5
tmp = tmp + singlsum(ct1)
Next
End Select
If sst(mainp, 11 - ct) > 0 Then
Xmn = pnum ^ ct * qnum ^ (5 - ct) * XdownI(5, ct) / XdownI(ct, ct)
Addvaluegame Xmn - tmp, sst(mainp, 11 - ct)
scatcomp = scatcomp + Xmn - tmp
scatpz = scatpz + (Xmn - tmp) * sst(mainp, 11 - ct)
End If
Next

Else
For ct = 2 To 5
Select Case ct
Case 2
'Exclude PPXPP
tmp = ((1 + sst(mainp, 3)) * qnum * ((qnum + pnum) ^ 2) * (pnum ^ 2) - mainRL * pnum ^ 4 * qnum) - tsm(3, 4, 5, mainp) - mainRL * tsm(1, 2, 3, mainp)
Case 3
tmp = (1 + sst(mainp, 3)) * qnum * (qnum + pnum) * (pnum ^ 3) - dsm(4, 5, mainp) - mainRL * dsm(1, 2, mainp)
Case 4
tmp = (1 + sst(mainp, 3)) * (pnum ^ 4) * qnum - singlsum(5) - mainRL * singlsum(1)
Case 5
tmp = pnum ^ 5
End Select


If sst(mainp, 11 - ct) > 0 Then
Addvaluegame tmp, sst(mainp, 11 - ct)
scatcomp = scatcomp + tmp
scatpz = scatpz + tmp * sst(mainp, 11 - ct)
End If


Next
End If
End If
'End main loop
Next


'now process substitutes
'substituting substotalvec, wheelvec, thumbsize


'Now for the free game combos
'If gamespinsymbol(1) > 0 Then   'don't ever calculate VOG when more than 1 fgame

'mainp = gamespinsymbol(1)

'If freegamesettings(1, 1) = 1 Then  'LR RL
'loopLRorany1 = 4
'loopLRorany2 = 4
'loopLRorany3 = 5
'loopLRorany4 = 5
'If freegamesettings(1, 2) = 1 Then intRLloop = 2

'Else    'any
'loopLRorany1 = 1
'loopLRorany2 = 4
'loopLRorany3 = 0
'loopLRorany4 = 0
'End If
'
'For ct = 1 To intRLloop
'If intRLloop = 2 Then
'loopLRorany1 = 1
'loopLRorany2 = 1
'loopLRorany3 = 2
'loopLRorany4 = 2
'End If
'For ct1 = loopLRorany1 To loopLRorany2
'If freegamesettings(1, 1) = 2 Then loopfix ct1, loopLRorany2, loopLRorany3, loopLRorany4
'For ct2 = loopLRorany3 To loopLRorany4
'
'Xmn = 1
'
'If ct1 <> ct2 Then
'For ct3 = 1 To 5
    'If ct3 <> ct1 And ct3 <> ct2 Then
    'Xmn = workvec(ct3, mainp) * Xmn
    'End If
'Next
'
'dsum is through all combinations WITHOUT mainp

'If freegamesettings(1, 1) = 1 Then
'tmp = (24 - substotalvec(ct1) - workvec(ct1, mainp)) * (24 - substotalvec(ct2))
'Else
'tmp = (24 - substotalvec(ct1) - workvec(ct1, mainp)) * (24 - substotalvec(ct2) - workvec(ct2, mainp))
'End If
'
'probfgame = tmp + probfgame
'
'End If
'Next
'Next
'Next
'
'
'
''middle threes
'If freegamesettings(1, 3) = 1 Then
'Xmn = workvec(2, mainp) * workvec(3, mainp) * workvec(4, mainp) * ((24 - substotalvec(1) - workvec(1, mainp)) * (24 - substotalvec(5) - workvec(5, mainp)))
'probfgame = Xmn + probfgame
'End If
'
'
''freegamesettings(1, 4) = 1 if only naturals,break the sequence freegamesettings(1, 5) = 0,freegamesettings(1, 8)=no of spins
''freegamesettings(1, 7)=noprize bonus,freegamesettings(1, 6)= prize bonus,probfgame p(FG)
''VOG = no free game * its prob + no of spins * prob event * (p(no prize)*reward+VOG) difference equation - please check
''freegamesettings(1, 5) = 0 SHOULD be taken into account but is ignored - Please comment
''now if freegamesettings(1, 7) > 0, a free game ALWAYS wins a prize
'
''Adjust VOG with expected fgame prize
''PZfg = N * ( P0 * B1 + (1 - P0)* VOG * B2) + Pfg * PZfg
'expzfgame = freegamesettings(1, 8) * ((1 - probgtequalthan(1) / 24 ^ 5) * freegamesettings(1, 7) + (probgtequalthan(1) / 24 ^ 5) * (VOG / 24 ^ 5) * freegamesettings(1, 6)) / (1 - probfgame / 24 ^ 5)
'VOG = VOG + probfgame * expzfgame
'
''Adjust P0 with expectation, really need to calculate trips above
'For ct = 1 To 10
'If expzfgame > ct Then probgtequalthan(ct) = probgtequalthan(ct) + probfgame
'Next
'
'End If
''Shall leave processing gamespinsymbol(2) for now


'VOG = (((VOG - scatpz) / 24 ^ 5) * (24 ^ 5 - scatcomp) / (24 ^ 5)) + ((scatpz / 24 ^ 5) * (1 - (probgtequalthan(1) - scatcomp) / 24 ^ 5)) + (VOG + scatpz) / (24 ^ 5) * (1 - (probgtequalthan(1) - scatcomp) / 24 ^ 5) * scatcomp / (24 ^ 5)

VOG = 100 * VOG / (24 ^ 5)


'Moneyback :More than 1 bet choice amount allowed


getBmax bmax

p0 = 100 * (1 - CSng((probgtequalthan(1) / 24 ^ 5)))

If p0 < 0 Then
VOGerror = True
p0 = 0
End If


'Random Jackpot
If gt(14) = 1 Then
VOG = VOG + 5
'Also require Random Jackpot contribution to P0
probRJ = 1 / (20 * CSng(gt(16) & decsep & gt(17)))
p0RJ = 100 * (1 - ((1 - CSng(probgtequalthan(1) / 24 ^ 5)) * probRJ + CSng(probgtequalthan(1) / 24 ^ 5) * probRJ + CSng(probgtequalthan(1) / 24 ^ 5) * (1 - probRJ)))
Else
p0RJ = p0
End If


If gt(10) > 0 Then

If VOG < 100 Then

'Find Btotal here

For Btotal = gt(10) - 1 + bmax(5) To bmax(5) * gt(10)
If moneybackvalue(VOG, p0RJ, Y, bettest, Btotal) = True Then
If Y > ymax Then ymax = Y
End If
Next

Else
'if VOG > 100, always bet MAX
For ct = 1 To gt(10)
bettest(0, ct) = bmax(5)
Next
If moneybackvalue(VOG, p0RJ, Y, bettest, gt(10) * bmax(5)) = True Then ymax = Y
End If

End If



If VOG > 0 Then
outputforformat VOG, 29, 137
outputforformat calcmonbackvog(ymax), 31, 138
outputforformat p0RJ, 33, 139
End If

VOGerr = VOGerror
VOGerror = False

End Sub
Public Function calcmonbackvog(ymax As Single) As Single
Dim probgt12 As Long
'gt(12)
If gt(12) > 0 And gt(10) > 0 Then
'Find mean time before getting prize is 1/probg12
    If gt(152) = 0 Then 'No Monte Carlo
    probgt12 = probgtequalthan(gt(12))
    If probgt12 = 0 Then probgt12 = gt(154)
    gt(154) = probgt12
    calcmonbackvog = VOG * ((1 - 2 * (p0 / 100) ^ (24 ^ 5 / probgt12)) + ymax * 2 * ((p0 / 100) ^ (24 ^ 5 / probgt12)))
    Else
    probgt12 = gt(154)
    calcmonbackvog = VOG * ((1 - 2 * (p0 / 100) ^ (gt(152) / probgt12)) + ymax * 2 * ((p0 / 100) ^ (gt(152) / probgt12)))
    End If
Else
calcmonbackvog = ymax * VOG
End If
End Function
Public Sub Addvaluegame(X As Long, intprizeno As Long)
VOG = VOG + X * intprizeno


If X < 0 Then VOGerror = True


If intprizeno >= 1 Then probgtequalthan(1) = probgtequalthan(1) + X
For loopval = 2 To gt(12)
If intprizeno >= loopval Then
probgtequalthan(loopval) = probgtequalthan(loopval) + X
End If
Next
End Sub
Public Function comparingpics(auxprizecount As Long, piccount As Long, mp As Long, Optional tp As Long = 0, Optional tp1 As Long = 0, Optional tp2 As Long = 0, Optional tp3 As Long = 0)
Dim intcount As Long, scatty As Long, substest As Long
'Slightly inefficient as comparingpics checks mainp each call with an extra expected  100 unnecessary comparisons on mainp

For intcount = 1 To piccount
comparingpics = False
substest = subsvec(intcount)
Select Case auxprizecount
Case 0
If mp <> substest Then comparingpics = True
Case 1
If mp <> substest And tp <> substest Then comparingpics = True
Case 2
If mp <> substest And tp <> substest And tp1 <> substest Then comparingpics = True
Case 3
If mp <> substest And tp <> substest And tp1 <> substest And tp2 <> substest Then comparingpics = True
Case 4
If mp <> substest And tp <> substest And tp1 <> substest And tp2 <> substest And tp3 <> substest Then comparingpics = True
End Select
If comparingpics = False Then Exit Function
Next



comparingpics = False
Select Case auxprizecount
Case 0
comparingpics = True    'mainp successfully compares with itself
Case 1
If tp <> mp Then comparingpics = True
Case 2
If tp <> mp And tp1 <> mp And tp1 <> tp Then comparingpics = True
Case 3
If tp <> mp And tp1 <> mp And tp2 <> mp And tp1 <> tp And tp2 <> tp And tp2 <> tp1 Then comparingpics = True
Case 4
If tp <> mp And tp1 <> mp And tp2 <> mp And tp3 <> mp And tp1 <> tp And tp2 <> tp And tp3 <> tp And tp2 <> tp1 And tp3 <> tp1 And tp3 <> tp2 Then comparingpics = True
End Select
If comparingpics = False Then Exit Function


For intcount = 1 To intscatternumber
scatty = intscattervec(intscatternumber, 2)
comparingpics = False
Select Case auxprizecount
Case 0
If mp <> scatty Then comparingpics = True
Case 1
If mp <> scatty And tp <> scatty Then comparingpics = True
Case 2
If mp <> scatty And tp <> scatty And tp1 <> scatty Then comparingpics = True
Case 3
If mp <> scatty And tp <> scatty And tp1 <> scatty And tp2 <> scatty Then comparingpics = True
Case 4
If mp <> scatty And tp <> scatty And tp1 <> scatty And tp2 <> scatty And tp3 <> scatty Then comparingpics = True
End Select
If comparingpics = False Then Exit Function
Next
End Function
Public Sub fixmult1(spslot As Long, T1LR As Long, T2LR As Long, T1RL As Long, T2RL As Long, pic1 As Long, pic2 As Long, singT1pz As Long, singT2pz As Long, Optional pic3 As Long = 0, Optional singT3pz As Long = 0, Optional pic4 As Long = 0, Optional singT4pz As Long = 0)
'The other prizes other than pic1 & pic2 in sum singles except in 2 1 1 X and pic1,pic2 are singles
'Note that singl2 is for 211X and 31X benefit only



tmp = 24 - substotalvec(spslot) - singlsum(spslot)


If T1LR = 1 And T2LR = 0 Then
    If spslot = reelend Then
    If singT1pz = 0 Or T1RL = 0 Then tmp = tmp - workvec(spslot, pic1)
    Else
    tmp = tmp - workvec(spslot, pic1)
    End If
    If singT2pz = 0 Then tmp = tmp - workvec(spslot, pic2)
ElseIf T1LR = 0 And T2LR = 1 Then
    If spslot = reelend Then
    If singT2pz = 0 Or T2RL = 0 Then tmp = tmp - workvec(spslot, pic2)
    Else
    tmp = tmp - workvec(spslot, pic2)
    End If
    If singT1pz = 0 Then tmp = tmp - workvec(spslot, pic1)
ElseIf T1LR = 1 And T2LR = 1 Then
tmp = tmp - workvec(spslot, pic1) - workvec(spslot, pic2)
If pic3 > 0 Then    'process extras here

'In a 211X, mainp  as spare (mainLR=0) may not be in singlsum
If singT3pz = 0 Then tmp = tmp - workvec(spslot, pic3)


If spslot <> singl1 Then Addvaluegame Xmn * Xrst * workvec(spslot, pic1), singT1pz + singT2pz + singT3pz + singT4pz
If spslot <> singl2 Then Addvaluegame Xmn * Xrst * workvec(spslot, pic2), singT1pz + singT1pz + singT3pz + singT4pz

End If

Else    'T1LR & T2LR = 0
    If singT1pz = 0 And singT2pz = 0 Then
    tmp = tmp - workvec(spslot, pic1) - workvec(spslot, pic2)
    ElseIf singT1pz = 0 Then
    tmp = tmp - workvec(spslot, pic1)
    ElseIf singT2pz = 0 Then
    tmp = tmp - workvec(spslot, pic2)
    End If
End If

End Sub
Private Sub fixmult2(mp As Long, ts As Long, mainRL As Long, testRL As Long, mspz As Long, axpz As Long, axpz1 As Long, Optional S1 As Long, Optional S2 As Long)


'S1,S2 now refer to opposing ends (different from 21XX)


If ts = 0 Then  '"mainLR" always 1 , testLR = 0 And test1LR = 0


tmp1 = 24 - substotalvec(ct1) - singlsum(ct1) - workvec(ct1, mp)
        
If ct2 <> reelend Then
tmp2 = 24 - substotalvec(ct2) - singlsum(ct2) - workvec(ct2, mp)

        'Note if mainRL = 1 then ct can be 2; ok
        If ct1 = singl1 Then
            
        'P x T1 p T2, P x p T1 T2
            
        Addvaluegame Xmn * Xrst * tmp1 * workvec(ct2, mp), axpz + axpz1 + mspz
            
        Else    'ct1<>singl1, gets here once
        
        'P T1 p p T2
        Addvaluegame Xmn * Xrst * workvec(ct1, mp) * workvec(ct2, mp), axpz + axpz1 + mspz
            
            
        'P T1 x p T2
        Addvaluegame Xmn * Xrst * tmp1 * workvec(ct2, mp), axpz + axpz1 + mspz
            
        'P T1 p x T2
        Addvaluegame Xmn * Xrst * workvec(ct1, mp) * tmp2, axpz + axpz1 + mspz
            
        End If
            
Else    'ct2=reelend

If mainRL = 1 Or ct2 = 1 Then
tmp2 = 24 - substotalvec(ct2) - singlsum(ct2)
Else
tmp2 = 24 - substotalvec(ct2) - singlsum(ct2) - workvec(ct2, mp)
End If
        
'P T1 p T2 x, P T1 T2 p x
If ct1 <> singl1 Then Addvaluegame Xmn * Xrst * workvec(ct1, mp) * tmp2, axpz + axpz1 + mspz


If ct = 1 Then

'mainLR extra combos : P x T1 T2 p, P T1 x T2 p ,P T1 T2 x p 'SYMMETRICAL
Addvaluegame Xmn * Xrst * tmp1 * workvec(ct2, mp), axpz + axpz1 + (1 + mainRL) * mspz
    
If ct1 = 3 Then
'P T1 p T2 p 'SYMMETRICAL
Addvaluegame Xmn * Xrst * workvec(3, mp) * workvec(ct2, mp), axpz + axpz1 + mspz
ElseIf ct1 <> singl1 Then
'P T1 T2 p p
If mainRL = 0 Then Addvaluegame Xmn * Xrst * workvec(ct1, mp) * workvec(ct2, mp), axpz + axpz1 + mspz
End If

End If


End If
Else    '"mainLR" always 1 , testLR = 1 or test1LR = 1


'Symmetrical and valid if ct = 2
    
tmp1 = 24 - substotalvec(ct1) - singlsum(ct1) - workvec(ct1, mp) - workvec(ct1, ts)
tmp2 = 24 - substotalvec(ct2) - singlsum(ct2) - workvec(ct2, mp) - workvec(ct2, ts)
    
    

'P T1 x p T, P x T1 p T, P x p T1 T
Addvaluegame Xmn * Xrst * tmp1 * workvec(ct2, mp), axpz + axpz1 + mspz

If ct1 <> S1 Then
'P T1 p x T
Addvaluegame Xmn * Xrst * workvec(ct1, mp) * tmp2, axpz + axpz1 + mspz
'P T1 p p T
Addvaluegame Xmn * Xrst * workvec(ct1, mp) * workvec(ct2, mp), axpz + axpz1 + mspz
End If
    

'P t x T1 T, P T1 t x T, P t T1 x T
Addvaluegame Xmn * Xrst * workvec(ct1, ts) * tmp2, axpz + axpz1 + mspz

If ct2 <> S2 Then
'P x t T1 T
Addvaluegame Xmn * Xrst * tmp1 * workvec(ct2, ts), axpz + axpz1 + mspz
'P t t T1 T
Addvaluegame Xmn * Xrst * workvec(ct1, ts) * workvec(ct2, ts), axpz + axpz1 + mspz
End If

'Finally "cross" products
'P T1 t p T, P t T1 p T, P t p T1 T
Addvaluegame Xmn * Xrst * workvec(ct1, ts) * workvec(ct2, mp), axpz + axpz1 + mspz

End If
End Sub
Private Sub fixmult3(LRLR As Boolean, mp As Long, ts As Long, mainRL As Long, testRL As Long, mspz As Long, axpz As Long)

'reversible

If LRLR = True Then


'P x p x T, P x t x T
Addvaluegame Xmn * Xrst * (workvec(ct2, mp) + workvec(ct2, ts)) * (tmp * tmp2 - dsm(ct1, ct3, mp, ts)), axpz + mspz

'P x x p T
Addvaluegame Xmn * Xrst * workvec(ct3, mp) * (tmp * tmp1 - dsm(ct1, ct2, mp, ts)), axpz + mspz

'P t x x T
Addvaluegame Xmn * Xrst * workvec(ct1, ts) * (tmp1 * tmp2 - dsm(ct2, ct3, mp, ts)), axpz + mspz

tmp = tmp - singlsum(ct1)
tmp1 = tmp1 - singlsum(ct2)
tmp2 = tmp2 - singlsum(ct3)

'P x p p T, P x t p T
Addvaluegame Xmn * Xrst * tmp * (workvec(ct2, mp) + workvec(ct2, ts)) * workvec(ct3, mp), axpz + mspz

'P t p x T, P t t x T
Addvaluegame Xmn * Xrst * workvec(ct1, ts) * (workvec(ct2, mp) + workvec(ct2, ts)) * tmp2, axpz + mspz

'P t x p T
Addvaluegame Xmn * Xrst * workvec(ct1, ts) * tmp1 * workvec(ct3, mp), axpz + mspz

'P t t p T, P t p p T
Addvaluegame Xmn * Xrst * workvec(ct1, ts) * (workvec(ct2, mp) + workvec(ct2, ts)) * workvec(ct3, mp), axpz + mspz


Else    'lrlr = false

If ct3 = reelend Then

If ct1 <> singl1 Then   'only gets here once

'P T p x x
Addvaluegame Xmn * Xrst * workvec(ct1, mp) * (tmp1 * tmp2 - dsm(ct2, ct3, mp, ts)), axpz + mspz

'P T x p x
Addvaluegame Xmn * Xrst * workvec(ct2, mp) * (tmp * tmp2 - dsm(ct1, ct3, mp, ts)), axpz + mspz

'P T x x p
If ct = 1 Then Addvaluegame Xmn * Xrst * workvec(5, mp) * (tmp * tmp1 - dsm(3, 4, mp, ts)), axpz + (1 + mainRL) * mspz

tmp = 24 - substotalvec(ct1) - workvec(ct1, mp) - singlsum(ct1)
tmp1 = 24 - substotalvec(ct2) - workvec(ct2, mp) - singlsum(ct2)
If mainRL = 1 Or ct3 = 1 Then
tmp2 = 24 - substotalvec(ct3) - singlsum(ct3)
Else
tmp2 = 24 - substotalvec(ct3) - singlsum(ct3) - workvec(ct3, mp)
End If

'P T p p x
Addvaluegame Xmn * Xrst * workvec(ct1, mp) * workvec(ct2, mp) * tmp2, axpz + mspz

'P T p x p
If ct = 1 Then Addvaluegame Xmn * Xrst * workvec(3, mp) * workvec(5, mp) * tmp1, axpz + (1 + mainRL) * mspz
    
    If mainRL = 0 Then
    'P T x p p
    Addvaluegame Xmn * Xrst * workvec(4, mp) * workvec(5, mp) * tmp, axpz + mspz
    'P T p p p
    Addvaluegame Xmn * Xrst * workvec(3, mp) * workvec(4, mp) * workvec(5, mp), axpz + mspz
    End If

Else    'ct1=singl1, here twice

'P x p T x, P x T p x
Addvaluegame Xmn * Xrst * workvec(ct2, mp) * (tmp * tmp2 - dsm(ct1, ct3, mp, ts)), axpz + mspz

'P x x T p, P x T x p
If ct = 1 Then Addvaluegame Xmn * Xrst * workvec(5, mp) * (tmp * tmp1 - dsm(2, ct2, mp, ts)), axpz + (1 + mainRL) * mspz


tmp = 24 - substotalvec(ct1) - singlsum(ct1) - workvec(ct1, mp)

    
    If ct2 = 3 Then
    'P x p T p
    If ct = 1 Then Addvaluegame Xmn * Xrst * workvec(3, mp) * workvec(5, mp) * tmp, axpz + (1 + mainRL) * mspz
    Else
    'P x T p p
    If mainRL = 0 Then Addvaluegame Xmn * Xrst * workvec(4, mp) * workvec(5, mp) * tmp, axpz + mspz
    End If

End If

Else    'ct3 <> reelend   ;only gets here once

'P x p x T
Addvaluegame Xmn * Xrst * workvec(ct2, mp) * (tmp * tmp2 - dsm(ct1, ct3, mp, ts)), axpz + mspz

'P x x p T
Addvaluegame Xmn * Xrst * workvec(ct3, mp) * (tmp * tmp1 - dsm(ct1, ct2, mp, ts)), axpz + mspz

tmp = 24 - substotalvec(ct1) - singlsum(ct1) - workvec(ct1, mp)

'P x p p T
Addvaluegame Xmn * Xrst * workvec(ct2, mp) * workvec(ct3, mp) * tmp, axpz + mspz


End If
End If

End Sub
Private Sub loopfix(basect As Long, upperct As Long, newctlow As Long, newctup As Long)
'Used for counting no-prize pairs, triples, quads AND prize pair doubles
'e.g. P P T T (both "any"). Other prize doubles of type Xmn are "missed" in parent loops.
If VOGstepdir = 1 Then
If basect = 5 Then
newctlow = 5
Else
newctlow = basect + 1
End If

If upperct = 5 Then
newctup = 5
Else
newctup = upperct + 1
End If
Else

'Reverse direction of loop
If basect = 1 Then
newctlow = 1
Else
newctlow = basect - 1
End If

If upperct = 1 Then
newctup = 1
Else
newctup = upperct - 1
End If
End If
End Sub
Public Sub getBmax(Bmaxget() As Long)


For ct = 1 To 5
bmax(ct) = 0
Next

'Doesn't hurt to init other vars here
For ct = 1 To 15
For ct1 = 0 To 126
betkeep(ct1, ct) = 0
Next
Next

If gt(25) > 0 Then
For ct = 1 To 5
bmax(ct) = gt(20 + ct)
Next
ElseIf gt(24) > 0 Then
For ct = 2 To 5
bmax(ct) = gt(19 + ct)
Next
ElseIf gt(23) > 0 Then
For ct = 3 To 5
bmax(ct) = gt(18 + ct)
Next
ElseIf gt(22) > 0 Then
bmax(4) = gt(21)
bmax(5) = gt(22)
Else
'Option of only 1 bet
bmax(5) = gt(21)
End If

For ct = 1 To 5
Bmaxget(ct) = bmax(ct)
Next

End Sub
Public Function moneybackvalue(zVOG As Single, zp0 As Single, newratio As Single, bettest() As Long, zBtotal As Long)
Dim testsum As Single
Btotal = zBtotal
VOG = zVOG
p0 = zp0
invprobgt12comp = 0

betkeeptrakker = 0
bestmonbackval = 0 'comparison value

For ct = 1 To gt(10)
'pieN = pieN * P gives steady state probs.
pieN(ct) = (1 - (p0 / 100)) * ((p0 / 100) ^ (ct - 1)) / (1 - (p0 / 100) ^ gt(10))
Next

If VOG < 100 Then

For ct = 0 To 126
For ct1 = 1 To gt(10)
betkeep(ct, ct1) = 0
bettest(ct, ct1) = 0
Next
Next


'Begin process

For ct = 1 To 5
If bmax(ct) > 0 Then

'Reinit betvec each pass
For ct1 = gt(10) - 1 To 1 Step -1
betvec(ct1) = bmax(ct)
Next
betvec(gt(10)) = bmax(5)

testbet 1

End If
Next

Else
For ct = 1 To gt(10)
betkeep(0, ct) = bettest(0, ct)
Next
End If

    'betkeep (0,X) > 0 when max is found
    If betkeep(0, gt(10)) > 0 Then
    
    'Mow need to find expected money on hand is sum(1,N)(P(Mn) * Mn)/N
    
    'First consider pieN * sum(1,N - 1)(sum(n + 1 ,N)(bn) + VOG * bn + p0 * M(n + 1))
    
    testsum = 0
    
    
    For ct = 1 To gt(10) - 1
    
    'testsum is evaluating expected return from bets 1 1 1 .. 1 with no money back for comparison
    testsum = testsum + pieN(ct) * (betkeep(0, ct) * (VOG / 100))
    Next
    
    newratio = (pieN(gt(10)) * ((1 - p0 / 100) * betkeep(0, gt(10)) * (VOG / 100) + Btotal * (p0 / 100)) + testsum) / ((1 - p0 / 100) * pieN(gt(10)) * betkeep(0, gt(10)) * (VOG / 100) + testsum)
   
   
    For ct = 0 To 126
    For ct1 = 1 To gt(10)
    bettest(ct, ct1) = betkeep(ct, ct1)
    Next
    Next
    
    moneybackvalue = True   'bets add to btotal to meet this condition
    Else
    moneybackvalue = False
    End If


End Function
Private Sub testbet(n As Long)
Dim loopcount1 As Long, loopcount2 As Long
For loopcount1 = 1 To 5
If betvec(n) = bmax(loopcount1) Then

        For loopcount2 = loopcount1 To 5
        If Testbetsum(n + 1) = False Then Exit For
        If n + 1 < gt(10) Then betvec(n + 1) = bmax(loopcount2)
        testbet n + 1
        Next

If n + 1 < gt(10) Then betvec(n + 1) = bmax(loopcount1) 'restore future betvec

End If
Next

End Sub
Private Function Testbetsum(Mx As Long)
Dim tempsum As Long, betfut As Long, testmonbackval As Single
tempsum = 0
Testbetsum = False

If Mx > gt(10) Then Exit Function

For ct1 = 1 To Mx
tempsum = tempsum + betvec(ct1)
Next

If tempsum > Btotal Then Exit Function

If Mx = gt(10) And tempsum = Btotal Then

'MONEYn = cash bet&won now + cash for future bets

testmonbackval = 0
For ct1 = 1 To gt(10) - 1
betfut = 0
For ct2 = ct1 + 1 To gt(10)
betfut = betfut + betvec(ct2)
Next
testmonbackval = pieN(ct1) * (betfut + betvec(ct1) * (VOG / 100)) + testmonbackval
Next

testmonbackval = pieN(gt(10)) * (betfut + betvec(gt(10)) * (VOG / 100) + Btotal * (p0 / 100)) + testmonbackval

betkeeptrakker = betkeeptrakker + 1
For ct1 = 1 To gt(10)
betkeep(betkeeptrakker, ct1) = betvec(ct1)
Next


If testmonbackval > bestmonbackval Then
bestmonbackval = testmonbackval
For ct1 = 1 To gt(10)
betkeep(0, ct1) = betvec(ct1)
Next
End If


Exit Function
End If

Testbetsum = True
End Function
Public Sub CalcScatterprize(scatterpicture As Long, scattersperreel As Long, baseprizeno As Long)
Dim pnum As Long, qnum As Long, i As Long, prizefactor As Long
pnum = 1    'both p,q have factor of 24 taken out
qnum = (24 - (3 * scattersperreel)) / (3 * scattersperreel)
If sst(scatterpicture, 1) = 1 Then 'L-R, R-L etc
'Disabled middle threes for these arrangements - same prize with R-L
'The following refers to combos like xxxyx or xxxyy where x is scatter
'sequenceof combos based on pq(p+q)**3, p**2*q(p+q)**2,p**3q(p+q),p**5 - expand these
'If baseno is 2 very complicated as we have to subtract p**2*q**3 & p**4*q
'from first two terms respectively, (in RL case of course) BUT if baseno is 3,
'RL prizefactor becomes 1,2*p,P*q
'More analysis
'xxxxx = 1
'xxxxy = q
'xxxyx =q
'xxxyy = q * q

'xxyxx = q
'xxyyx = q * q
'xxyxy = q * q
'xxyyy = q *q * q
'Factors are q, 1+q, 1+q

For i = baseprizeno To 5
Select Case i

Case 2
If baseprizeno = 2 Then prizefactor = 1
sst(scatterpicture, 9) = prizefactor
Case 3
If baseprizeno = 2 Then
prizefactor = 1 + qnum
Else
prizefactor = 1
End If
sst(scatterpicture, 8) = prizefactor
Case 4
prizefactor = (1 + qnum) * prizefactor
sst(scatterpicture, 7) = prizefactor
Case 5
prizefactor = (1 + sst(scatterpicture, 3)) * qnum * prizefactor
sst(scatterpicture, 6) = prizefactor

End Select
Next

Else 'Anys
'initial baseprizeno is 2
For i = baseprizeno To 5
Select Case i
Case 2
If baseprizeno = 2 Then prizefactor = 1
sst(scatterpicture, 9) = prizefactor
Case 3
If baseprizeno = 2 Then
prizefactor = qnum
Else
prizefactor = 1
End If
sst(scatterpicture, 8) = prizefactor
Case 4
prizefactor = 2 * qnum * prizefactor
sst(scatterpicture, 7) = prizefactor
Case 5
prizefactor = 5 * qnum * prizefactor
sst(scatterpicture, 6) = prizefactor
End Select
Next
End If
End Sub
