Attribute VB_Name = "Subtutes"
Option Explicit
'Private singl1 As Long, singl2 As Long, reelend As Long
'Private ct As Long, ct1 As Long, ct2 As Long, ct3 As Long, ct4 As Long, ct5 As Long
'Private tmp As Long, tmp1 As Long, tmp2 As Long
'Private Xrst As Long, Xrst1 As Long, Xmn As Long
'Private substotalvec(5) As Long, workvec(5, 14) As Long
'Private substatus As Long
'Private spairpz As Long, spairpz1 As Long, spairpz2 As Long
'Private ssingpz As Long, ssingpz1 As Long, ssingpz2 As Long
'Private striplpz As Long, striplpz1 As Long
'Private axpz1 As Long, axpz2 As Long, axpz3 As Long
'Private mnpz As Long, mdpz As Long, mspz As Long
'Private axpz As Long, aspz As Long
'Dim intRLloop As Long, mainp As Long, testp As Long, testp1 As Long, testp2 As Long, testp3 As Long
'Private mainRL As Long, mainLR As Long, mainmid As Long
'Private testLR As Long, test1LR As Long, test2LR As Long
'Private testRL As Long, test1RL As Long, test2RL As Long
'Private subsLR As Long, subs1LR As Long, subs2LR As Long
'Private subsRL As Long, subs1RL As Long, subs2RL As Long
'Private subsmid As Long, subs1mid As Long, subs2mid As Long
'Private subs As Long, subs1 As Long, subs2 As Long
'Private reelchk1 As Long, reelchk2 As Long, reelchk3 As Long
'Private subs1forsubs As Boolean, subsforsubs1 As Boolean, subs1forsubs2 As Long
'Private subs2forsubs1 As Long, subs2forsubs As Long, subsforsubs2
'Private piccount As Long, tmplong As Long
'Public Sub substituting(passsubstotalvec() As Long, wheelvec() As Long, thumbsize As Long)
'Dim counttemp1 As Long, pnum As Long, qnum As Long, intcount As Long
'Dim comparingtemp As Boolean, sumtyperesult As Boolean, subtst As Boolean


'piccount = thumbsize
'For ct = 1 To 5
'substotalvec(ct) = passsubstotalvec(ct)
'Next

'Mainp, testp, testp1 etc are NOT substitutes or substituters

'P T T1
'For mainp = 1 To thumbsize - 2
'mspz = sst(mainp, 10)
'For testp = mainp + 1 To thumbsize - 1
'axpz = sst(testp, 10)
'For testp1 = testp + 1 To thumbsize
'axpz1 = sst(testp1, 10)
'If mspz > 0 And axpz > 0 And axpz1 > 0 And comparingpics(2, thumbsize, mainp, testp, testp1, testp2, testp3) = True Then
'Sum to allow reversing the substitutes
'For subs = 1 To piccount - 1
'For subs1 = subs + 1 To piccount

'spairpz = sst(subs, 9)
'spairpz1 = sst(subs1, 9)
'ssingpz = sst(subs, 10)
'ssingpz1 = sst(subs1, 10)
'mainLR = sst(mainp, 1)
'testLR = sst(testp, 1)
'subsLR = sst(subs, 1)
'subs1LR = sst(subs1, 1)
'mainRL = sst(mainp, 3)
'testRL = sst(testp, 3)
'subsRL = sst(subs, 3)
'subs1RL = sst(subs1, 3)

'First mixed prize pairs S D in any category
'If comparingsubs(3, 2) = True Then 'The ssingpz > 0 done later
'If sumtype(1) = True Then


'For ct = 1 To intRLloop
'If ct = 2 Then sumtyperesult = sumtype(1)


'For ct1 = loopLRorany1 To loopLRorany2
'For ct2 = loopLRorany3 To loopLRorany4
'For ct3 = loopLRorany5 To loopLRorany6
'For ct4 = loopLRorany7 To loopLRorany8

'If ct1 <> ct2 And ct1 <> ct3 And ct1 <> ct4 And ct2 <> ct3 And ct2 <> ct4 And ct3 <> ct4 Then

'Xmn = 1
'Xrst = 1
'reelchk1 = 1
'reelchk2 = 1

    
'    For ct5 = 1 To 5
'        If ct5 = ct1 Then
'        'aux prize on this reel
'        Xrst = workvec(ct5, testp) * Xrst
'        ElseIf ct5 = ct2 Then
'        Xrst = workvec(ct5, testp1) * Xrst
'        ElseIf ct5 = ct3 Then
'        If reelcheck(subs, ct5) = False Then reelchk1 = 0
'        Xrst = workvec(ct5, subs) * Xrst
'        ElseIf ct5 = ct4 Then
'        If reelcheck(subs1, ct5) = True Then reelchk2 = 0
'        Xrst = workvec(ct5, subs1) * Xrst
'        Else
'        'main prize on this reel
'        Xmn = workvec(ct5, mainp) * Xmn
'        End If
'    Next
'    'Note spairpz's CAN be 0
'    If substitute(subs, subs1) = True And reelcheck(subs, ct3) = True Then
'    If subs1LR = 1 Then   'remember sumtype specifies range 1 - 5 of subs, subs1
'    If ct4 = 1 And ct3 = 2 Then Addvaluegame Xmn * Xrst, mspz + axpz + axpz1 + spairpz1
'    If subs1RL = 1 And ct4 = 5 And ct3 = 4 Then Addvaluegame Xmn * Xrst, mspz + axpz + axpz1 + spairpz1
'    Else 'subs1 is 'any'
'    Addvaluegame Xmn * Xrst, mspz + axpz + axpz1 + spairpz1
'    End If
'    ElseIf substitute(subs1, subs) = True And reelcheck(subs1, ct3) = True Then
'    If subsLR = 1 Then 'remember sumtype specifies range 1 - 5 of subs, subs1
'    If ct4 = 1 And ct3 = 2 Then Addvaluegame Xmn * Xrst, mspz + axpz + axpz1 + spairpz
'    If subsRL = 1 And ct4 = 5 And ct3 = 4 Then Addvaluegame Xmn * Xrst, mspz + axpz + axpz1 + spairpz
'    Else 'subs1 is 'any'
'    Addvaluegame Xmn * Xrst, mspz + axpz + axpz1 + spairpz
'    End If
'    Else 'Non matching subs ssingpz's  CAN be 0
'    Addvaluegame Xmn * Xrst, mspz + axpz + axpz1 + reelchk1 * ssingpz + reelchk2 * ssingpz1
'    End If

'End If
'Next
'Next
'Next
'Next
'Next
'End If
'Next
'Next



'The following 2 sums necessary since subs <> subs1 above
'For subs = 1 To piccount
'spairpz = sst(subs, 9)

'D D(same symbol) or S S (same symbol) doubles
'Note spairpz's, ssingpz's CAN be 0
'If comparingsubs(3, 1) = True Then


'If sumtype(2) = True Then


'For ct = 1 To intRLloop
'If ct = 2 Then sumtyperesult = sumtype(2)

'For ct1 = loopLRorany1 To loopLRorany2
'For ct2 = loopLRorany3 To loopLRorany4
'For ct3 = loopLRorany5 To loopLRorany6
'For ct4 = loopLRorany7 To loopLRorany8

'If ct1 <> ct2 And ct1 <> ct3 And ct1 <> ct4 And ct2 <> ct3 And ct2 <> ct4 And ct3 <> ct4 Then

'Xmn = 1
'Xrst = 1
'reelchk1 = 1
'reelchk2 = 1


'Note Xsub will be used for a different purpose here

'    For ct5 = 1 To 5
'        If ct5 = ct1 Then
        'aux prize on this reel
'        Xrst = workvec(ct5, testp) * Xrst
'        ElseIf ct5 = ct2 Then
'        Xrst = workvec(ct5, testp1) * Xrst
'        ElseIf ct5 = ct3 Or ct5 = ct4 Then
'        If reelcheck(subs, ct3) = False Then reelchk1 = 0
'        If reelcheck(subs, ct4) = False Then reelchk2 = 0
'        Xrst = workvec(ct5, subs) * Xrst
'        Else
'        'main prize on this reel
'        Xmn = workvec(ct5, mainp) * Xmn
'        End If
'    Next
     
'    singl1 = 2
'    singl2 = 4
'    fixreelends ct, reelend, singl1, singl2



'    If subsLR = 0 Then
'        If reelchk1 = 1 And reelchk2 = 1 Then
'        Addvaluegame Xmn * Xrst, mspz + axpz + axpz1 + spairpz
'                Else
'        Addvaluegame Xmn * Xrst, mspz + axpz + axpz1 + (reelchk1 + reelchk2) * ssingpz
'        End If
'        ElseIf (ct4 + ct3 = 3) Or (ct4 + ct3 = 9) Then
'        If reelchk1 = 1 And reelchk2 = 1 Then
'        Addvaluegame Xmn * Xrst, mspz + axpz + axpz1 + spairpz
'                Else
'        Addvaluegame Xmn * Xrst, mspz + axpz + axpz1 + reelchk1 * ssingpz
'        End If
'        ElseIf ct3 = 1 And ct4 = 5 Then
'        Addvaluegame Xmn * Xrst, mspz + axpz + axpz1 + reelchk1 * ssingpz + reelchk2 * ssingpz1
'        End If
        

'End If
'Next
'Next
'Next
'Next
'Next


'End If
'Next
'End If
'Next
'Next
'Next

'P T T1 S



'P P T
'For mainp = 1 To thumbsize
'mnpz = sst(mainp, 9)
'For testp = 1 To thumbsize
'axpz = sst(testp, 10)
'If mnpz > 0 And axpz > 0 And comparingpics(1, thumbsize, mainp, testp, testp1, testp2, testp3) = True Then
'End If
'Next
'Next

'P P P
'For mainp = 1 To thumbsize
'mnpz = sst(mainp, 9)
'If mnpz > 0 Then
'mainmid = sst(mainp, 4)
'End If
'Next



'P T X
'For mainp = 1 To thumbsize - 2
'mspz = sst(mainp, 10)
'For testp = mainp + 1 To thumbsize - 1
'axpz = sst(testp, 10)
'If mspz > 0 And axpz > 0 And comparingpics(1, thumbsize, mainp, testp, testp1, testp2, testp3) = True Then

'subsforsubs1 = False
'subs1forsubs = False

'Sum to allow reversing the substitutes
'For subs = 1 To piccount - 1
'For subs1 = subs + 1 To piccount

'subs1forsubs = substitute(subs1, subs)
'subsforsubs1 = substitute(subs, subs1)
'striplpz = sst(subs, 8)
'striplpz1 = sst(subs1, 8)
'spairpz = sst(subs, 9)
'spairpz1 = sst(subs1, 9)
'ssingpz = sst(subs, 10)
'ssingpz1 = sst(subs1, 10)
'mainLR = sst(mainp, 1)
'testLR = sst(testp, 1)
'subsLR = sst(subs, 1)
'subs1LR = sst(subs1, 1)
'mainRL = sst(mainp, 3)
'testRL = sst(testp, 3)
'subsRL = sst(subs, 3)
'subs1RL = sst(subs1, 3)



'Sum mixed single prizes S D
'If (comparingsubs(2, 2)) = True Then 'prizes done later
'If sumtype(4) = True Then


'For ct = 1 To intRLloop
'If ct = 2 Then sumtyperesult = sumtype(4)

'For ct1 = loopLRorany1 To loopLRorany2
'For ct2 = loopLRorany3 To loopLRorany4
'For ct3 = loopLRorany5 To loopLRorany6
'For ct4 = loopLRorany7 To loopLRorany8

'If ct1 <> ct2 And ct1 <> ct3 And ct1 <> ct4 And ct2 <> ct3 And ct2 <> ct4 And ct3 <> ct4 Then


'Xmn = 1
'Xrst = 1
'reelchk1 = 1
'reelchk2 = 1
'reelchk3 = 1
'If reelcheck(subs, ct3) = False Then reelchk1 = 0
'If reelcheck(subs1, ct4) = False Then reelchk2 = 0



    
'    For ct5 = 1 To 5
'        If ct5 = ct1 Then
        'aux prize on this reel
'        Xrst = workvec(ct5, testp) * Xrst
'        ElseIf ct5 = ct3 Then
'        Xrst = workvec(ct5, subs) * Xrst
'        ElseIf ct5 = ct4 Then
'        Xrst = workvec(ct5, subs1) * Xrst
'        ElseIf ct5 <> ct2 Then
        'main prize on this reel
'        Xmn = workvec(ct5, mainp) * Xmn
'        End If
'    Next
    

    'Note spairpz's, ssingpz's CAN be 0
        

'    If (ct4 + ct3 = 3) Or (ct4 + ct3 = 9) And subs1LR = 1 And reelcheck(subs, ct3) = True Then ' sumtype allows subs* in either reelend or reelend - 1 when LR = 1

'    If subsforsubs1 = True Then

    'subsLR for the moment is under "umbrella" of subs1LR
                
        'special fixmult on Y Y X S S , S S X Y Y
'        If mainLR = 0 Then
'        fixmult1 ct2, testLR, subs1LR, testRL, subs1RL, testp, subs1, 0, axpz, ssingpz1
'        ElseIf testLR = 0 Then
'        fixmult1 ct2, mainLR, subs1LR, mainRL, subs1RL, mainp, subs1, 0, mspz, ssingpz1
'        End If
'    Xrst = Xrst * tmp
'    Addvaluegame Xmn * Xrst, mspz + axpz + spairpz1
'        End If

'        ElseIf (ct4 + ct3 = 3) Or (ct4 + ct3 = 9) And subsLR = 1 And reelcheck(subs1, ct3) = True Then ' sumtype allows subs* in either reelend or reelend - 1 when LR = 1
    
'    If subs1forsubs = True Then

    'subs1LR for the moment is under "umbrella" of subsLR
                'special fixmult on Y Y X S S , S S X Y Y
'        If mainLR = 0 Then
'        fixmult1 ct2, testLR, subsLR, testRL, subsRL, testp, subs, 0, axpz, ssingpz
'        ElseIf testLR = 0 Then
'        fixmult1 ct2, mainLR, subsLR, mainRL, subsRL, mainp, subs, 0, mspz, ssingpz
'        End If
'    Xrst = Xrst * tmp
'    Addvaluegame Xmn * Xrst, mspz + axpz + spairpz
'        End If


'        ElseIf subsLR = 1 Or subs1LR = 1 Then
        
'        Addvaluegame Xmn * Xrst, mspz + axpz + axpz1 + reelchk1 * ssingpz + reelchk2 * ssingpz1

'        ElseIf subsforsubs1 = True And subs1LR = 0 Then

'    fixmult1 ct2, mainLR, testLR, mainRL, testRL, mainp, testp, 0, mspz, axpz

'    If reelchk1 = 1 And reelchk2 = 1 Then
'    Addvaluegame Xmn * Xrst, mspz + axpz + axpz1 + spairpz1
'        Else
'    Addvaluegame Xmn * Xrst, mspz + axpz + axpz1 + reelchk1 * ssingpz + reelchk2 * ssingpz1
'    End If

'        ElseIf subs1forsubs = True And subsLR = 0 Then

'    fixmult1 ct2, mainLR, testLR, mainRL, testRL, mainp, testp, 0, mspz, axpz

'    If reelchk1 = 1 And reelchk2 = 1 Then
'    Addvaluegame Xmn * Xrst, mspz + axpz + axpz1 + spairpz
'        Else
'    Addvaluegame Xmn * Xrst, mspz + axpz + axpz1 + reelchk1 * ssingpz + reelchk2 * ssingpz1
'    End If

'    Else 'singles
    

'    If mainLR = 0 And testLR = 0 Then
'    fixmult1 ct2, subsLR, subs1LR, subsRL, subs1RL, subs, subs1, 0, ssingpz, ssingpz1
'    ElseIf mainLR = 0 And subsLR = 0 Then
'    fixmult1 ct2, testLR, subs1LR, testRL, subs1RL, testp, subs1, 0, axpz, ssingpz1
'    ElseIf mainLR = 0 And subs1LR = 0 Then
'    fixmult1 ct2, testLR, subsLR, testRL, subsRL, testp, subs, 0, ssingpz, ssingpz
'    ElseIf testLR = 0 And subsLR = 0 Then
'    fixmult1 ct2, mainLR, subs1LR, mainRL, subs1RL, mainp, subs1, 0, mspz, ssingpz1
'    ElseIf testLR = 0 And subs1LR = 0 Then
'    fixmult1 ct2, mainLR, subsLR, mainRL, subsRL, mainp, subs, 0, mspz, ssingpz
'    ElseIf subsLR = 0 And subs1LR = 0 Then
'    fixmult1 ct2, mainLR, testLR, mainRL, testRL, mainp, testp, 0, mspz, axpz
'    End If
    
'    Xrst = Xrst * tmp
'    Addvaluegame Xmn * Xrst, mspz + axpz + reelchk1 * ssingpz + reelchk2 * ssingpz1

'    End If


    'Now for extra combos
    
 '   singl1 = 4
 '   singl2 = 2
 '   fixreelends ct, reelend, singl1, singl2

    
'    If ct2 = reelend And intRLloop = 1 And mainRL = 1 And testLR = 0 And subsLR = 0 And subs1LR = 0 Then
'        calc1111X mainp
'        If subsforsubs1 = True Then
'                If reelchk1 = 1 Then
'                Addvaluegame Xmn * Xrst, 2 * mspz + axpz + spairpz1
'                Else
'                Addvaluegame Xmn * Xrst, 2 * mspz + axpz + ssingpz + ssingpz1
'                End If
'        ElseIf subs1forsubs = True Then
'                If reelchk2 = 1 Then
'                Addvaluegame Xmn * Xrst, 2 * mspz + axpz + spairpz
'                Else
'                Addvaluegame Xmn * Xrst, 2 * mspz + axpz + ssingpz + ssingpz1
'                End If
'    Else
'        Addvaluegame Xmn * Xrst, 2 * mspz + axpz + ssingpz + ssingpz1
'        End If
'        End If

'    If ct2 = reelend And intRLloop = 1 And testRL = 1 And mainLR = 0 And subsLR = 0 And subs1LR = 0 Then
'    calc1111X testp
'        If subsforsubs1 = True Then
'                If reelchk2 = 1 Then
'                Addvaluegame Xmn * Xrst, mspz + 2 * axpz + spairpz1
'                Else
'                Addvaluegame Xmn * Xrst, mspz + 2 * axpz + ssingpz + ssingpz1
'                End If

'        ElseIf subs1forsubs = True Then
'                If reelchk2 = 1 Then
'                Addvaluegame Xmn * Xrst, mspz + 2 * axpz + spairpz
'                Else
'                Addvaluegame Xmn * Xrst, mspz + 2 * axpz + ssingpz + ssingpz1
'                End If
'    Else
'    Addvaluegame Xmn * Xrst, mspz + 2 * axpz + ssingpz + ssingpz1
'    End If
'        End If

        'SubsRL = 0, subs1RL = 0 and with extra subs or subs2 covered in separate loops below

        'S P T S1 S, S P T S S1
        'S P T S S1 only included by virtue of S, S1 being mutual substitutes
'    If subsRL = 1 And mainLR = 0 And testLR = 0 And subs1LR = 0 Then
    'subsforsubs1 = True CANNOT occur as subsRL = 1 and subs1LR = 0 above
'    calc1111X subs
'    If subs1forsubs = True And ct2 = reelend And ct4 = singl1 Then
'        calc1111X subs
'        If reelchk2 = 1 Then
'        Addvaluegame Xmn * Xrst, mspz + axpz + ssingpz + spairpz
'        Else
'        Addvaluegame Xmn * Xrst, mspz + axpz + 2 * ssingpz + ssingpz1
'        End If
'    ElseIf subs1forsubs = True And ct2 = singl1 And ct4 = reelend Then
'        If reelchk2 = 1 Then
'        Addvaluegame Xmn * Xrst, mspz + axpz + ssingpz + spairpz
'        Else
'        Addvaluegame Xmn * Xrst, mspz + axpz + ssingpz + ssingpz1
'        End If
'        ElseIf ct2 = reelend And intRLloop = 1 Then 'need to count only once
'        Addvaluegame Xmn * Xrst, mspz + axpz + 2 * ssingpz + ssingpz1
'    End If
'    End If

'    If subs1RL = 1 And mainLR = 0 And testLR = 0 And subsLR = 0 Then
'    calc1111X subs1
    'subs1forsubs = True CANNOT occur as subs1RL = 1 and subsLR = 0 above
 
'    If subsforsubs1 = True And ct2 = reelend And ct3 = singl1 Then
'        If reelchk1 = 1 Then
'        Addvaluegame Xmn * Xrst, mspz + axpz + ssingpz1 + spairpz1
'        Else
'        Addvaluegame Xmn * Xrst, mspz + axpz + ssingpz + 2 * ssingpz1
'        End If
'    ElseIf subsforsubs1 = True And ct2 = singl1 And ct3 = reelend Then
'        If reelchk1 = 1 Then
'        Addvaluegame Xmn * Xrst, mspz + axpz + ssingpz1 + spairpz1
'        Else
'        Addvaluegame Xmn * Xrst, mspz + axpz + ssingpz + ssingpz1
'        End If
'        ElseIf ct2 = reelend And intRLloop = 1 Then 'need to count only once
'        Addvaluegame Xmn * Xrst, mspz + axpz + ssingpz + 2 * ssingpz1
'    End If
'    End If

        
        'S P T S S1, S1 P T S1 S  with subsLR = 1 and subs1LR = 1
        'Of course S1 P T S S etc treated later
'        If subsLR = 1 And subs1LR = 1 And mainLR = 0 And testLR = 0 And ct2 = singl1 Then
'        If subsforsubs1 = True Then
'        calc1111X subs

'        If reelcheck(subs, singl1) = True Then
'        Addvaluegame Xmn * Xrst, mspz + axpz + ssingpz1 + spairpz1
'        Else
'        Addvaluegame Xmn * Xrst, mspz + axpz + ssingpz + ssingpz1
'        End If


'        ElseIf subs1forsubs = True Then 'S1 P T S1 S
'        calc1111X subs1

'        If reelcheck(subs1, singl1) = True Then
'        Addvaluegame Xmn * Xrst, mspz + axpz + ssingpz + spairpz
'        Else
'        Addvaluegame Xmn * Xrst, mspz + axpz + ssingpz + ssingpz1
'        End If

'        End If
'        End If
                
'End If
'Next
'Next
'Next
'Next
'Next
'End If
'Next
'Next






'P T S S1 S2  - no repeated Ds or Ss
'For subs = 1 To piccount - 2
'For subs1 = subs + 1 To piccount - 1
'For subs2 = subs1 + 1 To piccount

'subs1forsubs = substitute(subs1, subs)
'subsforsubs1 = substitute(subs, subs1)
'subs1forsubs2 = substitute(subs1, subs2)
'subs2forsubs1 = substitute(subs2, subs1)
'subs2forsubs = substitute(subs2, subs)
'subsforsubs2 = substitute(subs, subs2)
'ssingpz = sst(subs, 10)
'ssingpz1 = sst(subs1, 10)
'ssingpz2 = sst(subs2, 10)
'spairpz = sst(subs, 9)
'spairpz1 = sst(subs1, 9)
'spairpz2 = sst(subs2, 9)
'mainLR = sst(mainp, 1)
'testLR = sst(testp, 1)
'subsLR = sst(subs, 1)
'subs1LR = sst(subs1, 1)
'subs2LR = sst(subs2, 1)
'mainRL = sst(mainp, 3)
'testRL = sst(testp, 3)
'subsRL = sst(subs, 3)
'subs1RL = sst(subs1, 3)
'subs2RL = sst(subs2, 3)



'Sum mixed single prizes S D note subsLR = 0, subs1LR = 0, subs2LR = 0
'SubsLR cases above, note subs2 CANNOT form a triple with others
'If (comparingsubs(2, 3)) = True And subsLR = 0 And subs1LR = 0 And subs2LR = 0 Then 'prizes done later
'If sumtype(4) = True Then


'For ct = 1 To intRLloop
'If ct = 2 Then sumtyperesult = sumtype(4)

'For ct1 = loopLRorany1 To loopLRorany2
'For ct2 = loopLRorany3 To loopLRorany4
'For ct3 = loopLRorany5 To loopLRorany6
'For ct4 = loopLRorany7 To loopLRorany8

'If ct1 <> ct2 And ct1 <> ct3 And ct1 <> ct4 And ct2 <> ct3 And ct2 <> ct4 And ct3 <> ct4 Then

'Xmn = 1
'Xrst = 1
'reelchk1 = 1
'reelchk2 = 1
'reelchk3 = 1
'If reelcheck(subs, ct2) = False Then reelchk1 = 0
'If reelcheck(subs1, ct3) = False Then reelchk2 = 0
'If reelcheck(subs2, ct4) = False Then reelchk3 = 0



'    For ct5 = 1 To 5
'        If ct5 = ct1 Then
'        'aux prize on this reel
'        Xrst = workvec(ct5, testp) * Xrst
'        ElseIf ct5 = ct2 Then
'        Xrst = workvec(ct5, subs) * Xrst
'        ElseIf ct5 = ct3 Then
'        Xrst = workvec(ct5, subs) * Xrst
'        ElseIf ct5 = ct4 Then
'        Xrst = workvec(ct5, subs2) * Xrst
'        Else
'        'main prize on this reel
'        Xmn = workvec(ct5, mainp) * Xmn
'        End If
'    Next


        'Now decide by rank which prize to take
        
'        If subsforsubs1 = True Or subsforsubs2 = True Then
'                If reelchk1 = 1 Then
'                    If subs1 > subs2 Then
'                    Addvaluegame Xmn * Xrst, mspz + axpz + ssingpz1 + spairpz2
'                    Else
'                    Addvaluegame Xmn * Xrst, mspz + axpz + spairpz1 + ssingpz2
'                    End If
'                Else
'                Addvaluegame Xmn * Xrst, mspz + axpz + ssingpz + ssingpz1 + ssingpz2
'                End If
                        
'        ElseIf subs1forsubs = True Or subs1forsubs2 = True Then
'                If reelchk2 = 1 Then
'                    If subs > subs2 Then
'                    Addvaluegame Xmn * Xrst, mspz + axpz + ssingpz + spairpz2
'                    Else
'                    Addvaluegame Xmn * Xrst, mspz + axpz + spairpz + ssingpz2
'                    End If
'                Else
'                Addvaluegame Xmn * Xrst, mspz + axpz + ssingpz + ssingpz1 + ssingpz2
'                End If

'        ElseIf subs2forsubs = True Or subs2forsubs1 = True Then
'                If reelchk3 = 1 Then
'                    If subs > subs1 Then
'                    Addvaluegame Xmn * Xrst, mspz + axpz + ssingpz + spairpz1
'                    Else
'                    Addvaluegame Xmn * Xrst, mspz + axpz + spairpz + ssingpz1
'                    End If
'                Else
'                Addvaluegame Xmn * Xrst, mspz + axpz + ssingpz + ssingpz1 + ssingpz2
'                End If

'        Else    'No substituters so all singles
'        Addvaluegame Xmn * Xrst, mspz + axpz + ssingpz + ssingpz1 + ssingpz2
        
'        End If

'End If
'Next
'Next
'Next
'Next
'Next
'End If
'Next
'Next
'Next



'P T S D1 D1, P T S1 D D , D D D - repeated Ss, subs pairs only
'Can repeat as subs is a double S S
'For subs = 1 To piccount
'For subs1 = 1 To piccount

'subs1forsubs = substitute(subs1, subs)
'subsforsubs1 = substitute(subs, subs1)
'striplpz = sst(subs, 8)
'striplpz1 = sst(subs1, 8)
'spairpz = sst(subs, 9)
'spairpz1 = sst(subs1, 9)
'ssingpz = sst(subs, 10)
'ssingpz1 = sst(subs1, 10)
'mainLR = sst(mainp, 1)
'testLR = sst(testp, 1)
'subsLR = sst(subs, 1)
'subs1LR = sst(subs1, 1)
'mainRL = sst(mainp, 3)
'testRL = sst(testp, 3)


'If (comparingsubs(2, 3)) = True And subsLR = 0 And subs1LR = 0 Then 'prizes done later
'If sumtype(4) = True Then

'For ct = 1 To intRLloop
'If ct = 2 Then sumtyperesult = sumtype(4)

'For ct1 = loopLRorany1 To loopLRorany2
'For ct2 = loopLRorany3 To loopLRorany4
'For ct3 = loopLRorany5 To loopLRorany6
'For ct4 = loopLRorany7 To loopLRorany8

'If ct1 <> ct2 And ct1 <> ct3 And ct1 <> ct4 And ct2 <> ct3 And ct2 <> ct4 And ct3 <> ct4 Then

'Xmn = 1
'Xrst = 1
'reelchk1 = 1
'reelchk2 = 1
'reelchk3 = 1
'If reelcheck(subs, ct2) = False Then reelchk1 = 0
'If reelcheck(subs, ct3) = False Then reelchk2 = 0
'If reelcheck(subs1, ct4) = False Then reelchk2 = 0
    
    
'    For ct5 = 1 To 5
'        If ct5 = ct1 Then
'        'aux prize on this reel
'        Xrst = workvec(ct5, testp) * Xrst
'        ElseIf ct5 = ct2 Or ct5 = ct3 Then
'        Xrst = workvec(ct5, subs) * Xrst
'        ElseIf ct5 = ct4 Then
'        Xrst = workvec(ct5, subs) * Xrst
'        Else
'        'main prize on this reel
'        Xmn = workvec(ct5, mainp) * Xmn
'        End If
'    Next

'        singl1 = 2
'        singl2 = 3
'        fixreelends ct, reelend, singl1, singl2


        'Calculate S S S triple
        'Note sumtype expands ct2, ct3, ct4 from 1 to 3 here on combinations with SUBSTITUTES! only
        'Sumtype will have to accomodate each combination of reelchk1, reelchk2
'        If subsforsubs1 = True Then
        'subsLR for the moment is under "umbrella" of subs1LR
'        If reelchk1 = 1 And reelchk2 = 1 Then
'        Addvaluegame Xmn * Xrst, mspz + axpz + striplpz1
'        ElseIf reelchk1 = 1 Then
                'S S S1 combos note sumtype restricts ct2, ct3, ct4 from 1 to 3
'                If subs1LR = 0 Then
'                Addvaluegame Xmn * Xrst, mspz + axpz + spairpz + ssingpz1
'                Else     'subsLR = 1 here
'                If subsLR = 0 Then
'                        If ct4 = reelend Then
'                        Addvaluegame Xmn * Xrst, mspz + axpz + spairpz + ssingpz1
'                        Else
'                        Addvaluegame Xmn * Xrst, mspz + axpz + spairpz
'                        End If
'                Else    'both subsLR, subs1LR = 1
'                        If ct4 = 3 Then 'note sumtype expands ct2, ct3, ct4 from 1 to 3 here
'                        Addvaluegame Xmn * Xrst, mspz + axpz + spairpz
'                        Else
'                        Addvaluegame Xmn * Xrst, mspz + axpz + spairpz1
'                        End If
'                End If
'                End If
'        ElseIf reelchk2 = 1 Then
'                'S S S1 combos note sumtype restricts ct2, ct3, ct4 from 1 to 3
'                If subs1LR = 0 Then
'                Addvaluegame Xmn * Xrst, mspz + axpz + spairpz + ssingpz1
'                Else     'subsLR = 1 here - note sumtype expands ct2, ct3, ct4 from 1 to 3 here
'                If subsLR = 0 Then
'                        If ct4 = reelend Then
'                        Addvaluegame Xmn * Xrst, mspz + axpz + spairpz + ssingpz1
'                        Else
'                        Addvaluegame Xmn * Xrst, mspz + axpz + spairpz
'                        End If
'                Else    'both subsLR, subs1LR = 1
'                        If ct4 = 3 Then   'note sumtype expands ct2, ct3, ct4 from 1 to 3 here
'                        Addvaluegame Xmn * Xrst, mspz + axpz + spairpz
'                        ElseIf ct4 = singl2 Then
'                        Addvaluegame Xmn * Xrst, mspz + axpz + ssingpz
'                        Else    'ct4 = reelend
'                        Addvaluegame Xmn * Xrst, mspz + axpz + ssingpz1
'                        End If
'                End If
'                End If
'        Else    'reelchk1 & 2 = 0       'This is treated in sumtype as if there are no substitutes - more code
'        Addvaluegame Xmn * Xrst, mspz + axpz + spairpz + ssingpz1
'        End If
'        ElseIf subs1forsubs = True Then
'        'subs1LR for the moment is under "umbrella" of subsLR
'        If reelchk3 = 1 Then
'        Addvaluegame Xmn * Xrst, mspz + axpz + striplpz
'        Else    'reelchk3 = 0
'                'S S S1 combos note sumtype restricts ct2, ct3, ct4 from 1 to 3
'                If subs1LR = 0 Then
'                Addvaluegame Xmn * Xrst, mspz + axpz + spairpz + ssingpz1
'                Else    'subs1LR = 1
'                If subsLR = 0 Then
'                        If ct4 = reelend Then
'                        Addvaluegame Xmn * Xrst, mspz + axpz + spairpz + ssingpz1
'                        Else
'                        Addvaluegame Xmn * Xrst, mspz + axpz + spairpz
'                        End If
'                Else    'both subsLR, subs1LR = 1
'                        If ct4 = 3 Then   'note sumtype expands ct2, ct3, ct4 from 1 to 3 here
'                        Addvaluegame Xmn * Xrst, mspz + axpz + spairpz
'                        ElseIf ct4 = singl2 Then
'                        Addvaluegame Xmn * Xrst, mspz + axpz + ssingpz
'                        Else    'ct4 = reelend
'                        Addvaluegame Xmn * Xrst, mspz + axpz + ssingpz1
'                        End If
'                End If
'                End If
'        End If
'        Else    'no substitutes 'Different range in sumtype
'        Addvaluegame Xmn * Xrst, mspz + axpz + spairpz
'        End If

'End If
'Next
'Next
'Next
'Next
'Next
'End If


'Middles
'subsLR = sst(subs, 1)
'subs1LR = sst(subs1, 1)
'subsmid = sst(subs, 4)
'subs1mid = sst(subs1, 4)
'reelchk1 = 1
'reelchk2 = 1
'tmp = 1
'Xmn = 1
'Xrst = 1
'For intcount = 2 To 4
'If wheelvec(intcount, subs) = True Then
'        Xmn = wheelvec(intcount, subs) * Xmn
'        If reelcheck(subs, intcount) = False Then
'        If tmp = 1 Then
'        reelchk1 = 0
'        tmp = 2
'        ElseIf tmp = 2 Then
'        reelchk2 = 0
'        End If
'        End If
'If wheelvec(intcount, subs1) = True Then
'Xrst = wheelvec(intcount, subs1) * Xrst
'If reelcheck(subs1, intcount) = False Then reelchk3 = 0
'End If
'End If
'Next

'For intcount = 1 To 2
'If intcount = 1 Then
'Xmn = wheelvec(1, mainp) * Xmn
'Xrst = wheelvec(5, testp) * Xmn
'Else
'Xmn = wheelvec(5, mainp) * Xmn
'Xrst = wheelvec(1, testp) * Xmn
'End If

'If subs1mid = 1 And subsLR = 0 Then     'subs1forsubs always false

'If subsforsubs1 = True Then
'        If reelchk1 = 1 And reelchk2 = 1 Then
'        Addvaluegame Xmn * Xrst, mspz + axpz + striplpz
'        Else
'        Addvaluegame Xmn * Xrst, mspz + axpz + spairpz + ssingpz1
'        End If
'Else
'Addvaluegame Xmn * Xrst, mspz + axpz + spairpz + ssingpz1
'End If

'ElseIf subsmid = 1 And subs1LR = 0 Then     'subsforsubs1 always false

'If subs1forsubs = True Then
'        If reelchk1 = 1 And reelchk2 = 1 Then
'        Addvaluegame Xmn * Xrst, mspz + axpz + striplpz1
'        Else
'        Addvaluegame Xmn * Xrst, mspz + axpz + spairpz + ssingpz1
'        End If
'Else
'Addvaluegame Xmn * Xrst, mspz + axpz + spairpz + ssingpz1
'End If

'ElseIf subs1mid = 1 Then

'If subs1forsubs = True Then
'        If reelchk1 = 1 And reelchk2 = 1 Then
'        Addvaluegame Xmn * Xrst, mspz + axpz + striplpz
'        Else
'        Addvaluegame Xmn * Xrst, mspz + axpz
'        End If
'Else
'Addvaluegame Xmn * Xrst, mspz + axpz
'End If

'Else
'If subsforsubs1 = True Then
'        If reelchk1 = 1 And reelchk2 = 1 Then
'        Addvaluegame Xmn * Xrst, mspz + axpz + striplpz1
'        Else
'        Addvaluegame Xmn * Xrst, mspz + axpz
'        End If
'Else
'Addvaluegame Xmn * Xrst, mspz + axpz
'End If
'End If
'Next






'Next
'Next
'End If
'Next
'Next




'P P X
'For mainp = 1 To thumbsize
'mnpz = sst(mainp, 9)
'If mnpz > 0 Then

'End If
'Next

'P X X
'For mainp = 1 To thumbsize
'mnpz = sst(mainp, 10)
'If mnpz > 0 Then

'End If
'Next


'X X X


'End Sub
'Private Function comparingsubs(prizecount1 As Long, prizecount2 As Long)
'prizecount1 is mainp to testp*, prizecount2 is sub to sub*
'Dim intcount As Long
'Dim intcount1 As Long
'Dim scatty As Long
'Dim substest As Long

'comparingsubs = False
'For intcount = 1 To piccount
'If intcount <> subs Then
'If substitute(intcount, subs) = True Then comparingsubs = True
'If substitute(subs, intcount) = True Then comparingsubs = True
'End If
'Next
'If comparingsubs = False Then Exit Function

'If prizecount2 = 2 Then
'comparingsubs = False
'For intcount = 1 To piccount
'If intcount <> subs1 Then
'If substitute(intcount, subs1) = True Then comparingsubs = True
'If substitute(subs1, intcount) = True Then comparingsubs = True
'End If
'Next
'If comparingsubs = False Then Exit Function
'End If


'Select Case prizecount1
'Case 1

'For intcount = 1 To piccount
'comparingsubs = False
'Select Case prizecount2
'Case 1
'If subs <> mainp Then comparingsubs = True
'Case 2
'If subs <> mainp Then comparingsubs = True
'If subs1 <> mainp Then comparingsubs = True
'End Select
'If comparingsubs = False Then Exit Function
'Next

'Case 2

'For intcount = 1 To piccount
'comparingsubs = False
'Select Case prizecount2
'Case 1
'If subs <> mainp And subs <> testp Then comparingsubs = True
'Case 2
'If subs <> mainp And subs <> testp Then comparingsubs = True
'If subs1 <> mainp And subs1 <> testp Then comparingsubs = True
'End Select
'If comparingsubs = False Then Exit Function
'Next

'Case 3

'For intcount = 1 To piccount
'comparingsubs = False
'Select Case prizecount2
'Case 1
'If subs <> mainp And subs <> testp And subs <> testp1 Then comparingsubs = True
'Case 2
'If subs <> mainp And subs <> testp And subs <> testp1 Then comparingsubs = True
'If subs1 <> mainp And subs1 <> testp And subs1 <> testp1 Then comparingsubs = True
'End Select
'If comparingsubs = False Then Exit Function
'Next

'End Select


'For intcount = 1 To intscatternumber
'scatty = intscattervec(intscatternumber, 2)
'comparingsubs = False
'Select Case prizecount2
'Case 1
'If subs <> scatty Then comparingsubs = True
'Case 2
'If subs <> scatty And subs1 <> scatty Then comparingsubs = True
'End Select
'If comparingsubs = False Then Exit Function
'Next

'End Function
'Private Sub calc2X11(testpik As Long)
'    Xrst = 1
'    For ct5 = 1 To 5
'        If ct5 = ct1 Then
'        Xrst = workvec(ct4, testp) * Xrst
'        ElseIf ct5 = ct2 Then
'        Xrst = workvec(ct2, testpik) * Xrst
'        ElseIf ct5 = ct3 Then
'        Xrst = workvec(ct4, subsRL) * Xrst
'        ElseIf ct5 = ct4 Then
'        Xrst = workvec(ct3, subs1RL) * Xrst
'        End If
'    Next
'End Sub
'Private Sub calc21XX(testpik As Long)
'Xrst = 1
'Xrst1 = 1
'For ct5 = 1 To 5
'    If ct5 = ct1 Then
'    tmp = workvec(ct5, testpik)
'
'    If ct5 = singl1 Or ct5 = singl2 Then
'    tmp1 = 24 - substotalvec(ct5) - singlsum(ct5) - workvec(ct5, testpik)
'    Else
'    tmp1 = 24 - substotalvec(ct5) - singlsum(ct5)
'    End If
'
'    Xrst = tmp1 * Xrst
'    Xrst1 = tmp * Xrst1
'    ElseIf ct5 = ct2 And ct5 = reelend Then 'Prize at the end
'    tmp = workvec(ct5, testpik)
'    Xrst = tmp * Xrst
'    Xrst1 = tmp * Xrst1
'    ElseIf ct5 = ct3 Then
'    'Single prize here
'    tmp = workvec(ct5, testp)
'    Xrst = tmp * Xrst
'    Xrst1 = tmp * Xrst1
'    End If
'Next
'End Sub
'Private Sub calc211X(testpik As Long)
'    Xrst = 1
'    For ct5 = 1 To 5
'        If ct5 = ct1 Then
'        Xrst = workvec(ct5, testp) * Xrst
'
'        ElseIf ct5 = ct2 Then
'        Xrst = workvec(ct5, testp1) * Xrst
'
'        ElseIf ct5 = ct3 Then
'        Xrst = workvec(ct3, testpik) * Xrst
'        End If
'    Next
'End Sub
'Private Sub calc1111X(testpik As Long)
'    Xrst = 1
'    For ct5 = 1 To 5
'        If ct5 = ct1 Then
'        Xrst = workvec(ct5, testp) * Xrst
'
'        ElseIf ct5 = ct3 Then
'        Xrst = workvec(ct5, subs) * Xrst
'
'        ElseIf ct5 = ct4 Then
'        Xrst = workvec(ct5, subs1) * Xrst
'
'        ElseIf ct5 = ct2 Then
'        Xrst = workvec(ct5, testpik) * Xrst
'        End If
'    Next
'End Sub


