Attribute VB_Name = "Pokrouts"
Option Explicit
Private thumbsize As Long
Private wheelorder(4, 24) As Long, thumbstrue(14) As StdPicture, wheelvec(5, 14) As Long
Private savevec(5, 5, 5) As Long, PZs(2, 9, 3) As Long, subcount(1, 6, 1) As Long, lonesubz(3, 1) As Long, bonuz(14) As Long, subshere(4) As Boolean
Private win1(4) As Long, win2(4) As Long, win3(4) As Long, win2old(2, 4) As Long, subsused(4) As Boolean, allocsub(4) As Boolean
Private spinvec(4) As Long, gamevec(4) As Long, spinztemp(4) As Long, holdq(2, 3, 10, 6) As Long, PYLN(4) As Long
Private ohqgs(2, 4) As Long, heldnow(4) As Long, holdqnew(4) As Long, qtrakker(2, 3) As Long, havereset As Boolean, winsub(2, 4) As Long, lonz(3, 14) As Long
Private sstabgen(14, 10) As Long, GTp(200) As Long
Private showprize As Boolean, prizecount(2) As Long, prizetotal As Long
Private intreel As Long, temp As Long, response As Long, currline As Long, subsline(2) As Boolean
Private pztk1 As Long, pztk2 As Long, pztk3 As Long, lz1 As Long, lz2 As Long, ct As Long, ct1 As Long, ct2 As Long, ct3 As Long
Private gamenotspin As Boolean, gamesaved As Boolean, spinsaved As Boolean, juststarted As Boolean
Private gamespintot As Long, kept1or2 As Long, fgct As Long, spct As Long
Private Defaultgt(65) As Long, Defaultstr(100) As String
Public Sub FinishPictureConfig(thumb() As StdPicture, piccount As Long)
thumbsize = piccount
For pct = 1 To piccount
Set thumbstrue(pct) = thumb(pct)
Next
setinauxrouts thumbsize
End Sub
Public Sub getthumbspiccount(thumb() As StdPicture, piccount As Long, weelvec() As Long, weelorder() As Long)
piccount = thumbsize
For ct = 1 To piccount
Set thumb(ct) = thumbstrue(ct)
For intreel = 1 To 5
weelvec(intreel, ct) = wheelvec(intreel, ct)
Next
Next
For intreel = 0 To 4
For ct = 1 To 24
weelorder(intreel, ct) = wheelorder(intreel, ct)
Next
Next
End Sub
Public Sub setthumbspiccount(thumb() As StdPicture, piccount As Long, weelvec() As Long, weelorder() As Long)
thumbsize = piccount
For ct = 1 To piccount
Set thumbstrue(ct) = thumb(ct)
For intreel = 1 To 5
wheelvec(intreel, ct) = weelvec(intreel, ct)
Next
Next
For intreel = 0 To 4
For ct = 1 To 24
wheelorder(intreel, ct) = weelorder(intreel, ct)
Next
Next
End Sub
Public Function Inputvars()
Dim c As New cRegistry
runAdmin = False
Inputvars = False
temp = 0        '1 if an error


sDatabaseName = loaddirectory & "Slotdata.s$t"


If OpenDb(sDatabaseName, 0) = False Then Exit Function

Set dbsCurrent = gdbCurrentDB

REORGSLOT



Set rectemp = dbsCurrent.OpenRecordset("SELECT Bullean FROM Inpoot")
With rectemp

On Error GoTo ErrHandler

.MoveFirst

For ct = 0 To 1
disablegamespintabs(ct) = ![Bullean]
.MoveNext
Next

For ct = 1 To 14
For ct1 = 1 To 14
substitute(ct, ct1) = ![Bullean]
.MoveNext
Next

For ct1 = 0 To 5
reelcheck(ct, ct1) = ![Bullean]
.MoveNext
Next
Next
End With

Set rectemp = dbsCurrent.OpenRecordset("SELECT Lorng FROM Inpoot")
With rectemp
.MoveFirst

thumbsize = ![Lorng]
.MoveNext


For ct = 0 To 4
For ct1 = 1 To 24
wheelorder(ct, ct1) = ![Lorng]
.MoveNext
If wheelorder(ct, ct1) < 1 Or wheelorder(ct, ct1) > 24 Then IE (1)
Next
Next


For ct = 1 To 5
For ct1 = 1 To 14
wheelvec(ct, ct1) = ![Lorng]
.MoveNext
If wheelvec(ct, ct1) < 0 Or wheelvec(ct, ct1) > 22 Then IE (2)
Next
Next



For ct = 1 To 14
For ct1 = 0 To 10
sstabgen(ct, ct1) = ![Lorng]
.MoveNext
Select Case ct1
Case 0
If sst(ct, 0) < 0 Or sstabgen(ct, 0) > 17 Then IE (5)
Case 1 To 5
If sst(ct, ct1) < 0 Or sstabgen(ct, ct1) > 1 Then IE (6)
Case Else
If sst(ct, ct1) < 0 Or sstabgen(ct, ct1) > 10200 Then IE (7)
End Select
Next
Next


For ct = 1 To 200
GTp(ct) = ![Lorng]
.MoveNext

If GTp(ct) < 0 Then
    Select Case ct
    Case 29
    GTp(ct) = -GTp(ct)
    VOGchg = 1
    Case 188
    gt(188) = 0    'Only for V 2.14
    Case 55, 56, 152, 161, 162, 185, 187, 188
    'Ok
    Case Else
    IE 8, 0
    End Select
End If

Select Case ct
Case 1, 9, 14, 37, 44 To 46, 156, 183, 184, 189, 190, 193, 195
If GTp(ct) > 1 Then IE 9, 1
Case 2, 26 To 29, 31, 33, 16, 47 To 88, 104 To 136, 149 To 150, 154, 161 To 179, 181, 185, 188, 191, 195
'No limit here
Case 3, 192
If GTp(ct) > 5 Then IE 10, 5
Case 187
If Abs(GTp(187)) > 8 Then IE 11, 7
Case 153, 157, 158
If GTp(ct) > 2 Then IE 12, 2
Case 4, 5, 137 To 148, 160, 186
If GTp(ct) > 3 Then IE 13, 3
Case 19, 30, 32, 34, 43, 106
If GTp(ct) > 9999 Then IE 14, 9999
Case 6, 7, 8
If GTp(ct) > 4 Then IE 15, 4
Case 10, 11, 151
If GTp(ct) > 15 Then IE 16, 15
Case 12, 15, 20 To 25
If GTp(ct) > 10 Then IE 17, 10
Case 13
If GTp(13) > 300 Then IE 18, 300
Case 180, 182, 17
If GTp(ct) > 9 Then IE 19, 9
Case 18
If GTp(18) > 20000 Then IE 20, 20000
Case 89 To 103
If GTp(ct) > 30030 Then IE 21, 30030
Case 35, 152
If GTp(ct) > 50000000 Then IE 22, 50000000
Case 36
If GTp(ct) > 6 Then IE 23, 6
Case 38, 39, 194
If GTp(ct) > 72 Then IE 24, 59
Case 40
If GTp(ct) > 23 Then IE 25, 23
Case 41
If GTp(41) > 31 Then IE 26, 28
Case 42
If GTp(42) > 12 Then IE 27, 12
Case 155
If GTp(155) > 500500 Then IE 28, 500500
End Select
Next
For ct = 0 To 1
For ct1 = 1 To 9
freegamesettings(ct, ct1) = ![Lorng]
.MoveNext
If freegamesettings(ct, ct1) < 0 Then IE (29)
Select Case ct1
Case 1 To 3
If freegamesettings(ct, ct1) > 2 Then IE (30)
Case 4, 5
If freegamesettings(ct, ct1) > 1 Then IE (31)
Case 9
If freegamesettings(ct, 9) > 10 Then IE (32)
Case Else
If freegamesettings(ct, 1) > 0 And freegamesettings(ct, 6) = 0 Then IE (33)
If freegamesettings(ct, ct1) > 10 Then IE (34)

End Select
Next
For ct1 = 1 To 15
spinsettings(ct, ct1) = ![Lorng]
.MoveNext
If spinsettings(ct, ct1) < 0 Then IE (35)
Select Case ct1
Case 1 To 8
If spinsettings(ct, ct1) > 2 Then IE (36)
Case 9 To 11
If spinsettings(ct, ct1) > 1 Then IE (37)
Case 14
If (spinsettings(ct, 1) > 0 Or spinsettings(ct, 4) > 0 Or spinsettings(ct, 7) > 0) And spinsettings(ct, 14) = 0 Then IE (38)
Case 15
If spinsettings(ct, 15) > 10 Then IE (39)
Case Else
If spinsettings(ct, ct1) > 5 Then IE (40)
End Select
Next
Next
For ct = 0 To 3
gamespinsymbol(ct) = ![Lorng]
.MoveNext
If gamespinsymbol(ct) < 0 Or gamespinsymbol(ct) > 14 Then IE (41)
gamespinkeep(ct) = ![Lorng]
.MoveNext
If gamespinkeep(ct) < 0 Or gamespinkeep(ct) > 14 Then IE (42)
Next

For ct = 1 To 65
Defaultgt(ct) = ![Lorng]
.MoveNext
Next

End With


Set rectemp = dbsCurrent.OpenRecordset("SELECT Streeng FROM Inpoot")
With rectemp
.MoveFirst

Dim tempstr As Integer
For ct = 1 To 100
Stringvars(ct) = ![Streeng]
.MoveNext
Select Case ct
Case 5 To 8
If Len(Stringvars(ct)) > 20 Then IE (44)
If Namevalid(2, Stringvars(ct)) = False Then IE (45)
Case 1 To 4, 9 To 12, 26 To 50
If Len(Stringvars(ct)) > 100 Then IE (46)
If Namevalid(1, Stringvars(ct)) = False Then IE (47)
Case 13 To 25
If Namevalid(2, Stringvars(ct)) = False Then IE (48)
If Len(Stringvars(ct)) > 30 Then IE (49)
End Select
Next
For ct = 1 To 100
Defaultstr(ct) = CStr(![Streeng])
If Not (.EOF) Then .MoveNext
Select Case ct
Case 5 To 8
If Len(Defaultstr(ct)) > 20 Then IE (50)
If Namevalid(2, Defaultstr(ct)) = False Then IE (51)
Case 1 To 4, 9 To 12, 26 To 50
If Len(Defaultstr(ct)) > 100 Then IE (52)
If Namevalid(1, Defaultstr(ct)) = False Then IE (53)
Case 13 To 25
If Namevalid(2, Stringvars(ct)) = False Then IE (54)
If Len(Stringvars(ct)) > 30 Then IE (55)
End Select
Next

End With


For ct = 1 To 200
gt(ct) = GTp(ct)
Next

For pct = 1 To 14
For ct = 0 To 10
sst(pct, ct) = sstabgen(pct, ct)
Next
Next


If gt(200) = 1 Or gt(192) < 4 Then 'Install trigger or old ver
'Associate a File of type .s$t
If gt(200) = 1 Then
If Not c.CreateEXEAssociation(App.Path & "\MyReels.exe", "MyReels", "MyReels", "s$t", , , , , , , , 0) Then
runAdmin = True
GoTo ErrHandler
End If
End If
gt(158) = 2

If Dir(Stringvars(1)) <> "" Then
Restoredefaults False, (gt(192) < 4)
Else
Restoredefaults False, False
End If
gt(158) = 1
If gt(194) = 0 Then gt(194) = c.MHZ
End If

If gt(191) > 32767 Then gt(191) = 32767 'Quote limits


killdb sDatabaseName

Set rectemp = Nothing
Set dbsCurrent = Nothing


If temp = 0 Then Inputvars = True

Exit Function
ErrHandler:
Set rectemp = Nothing
Set dbsCurrent = Nothing
ShowError
killdb sDatabaseName
End Function
Private Sub IE(errorcode As Long, Optional correctval As Long = -1)
response = MsgBox("The game input values have been corrupted - error code " & CStr(errorcode) & " . Click OK to continue attempting to load the game, max default values assumed.", vbOKCancel)
If response = vbOK Then
If correctval > -1 Then GTp(ct) = correctval
temp = 0
Else
temp = 1
End If
End Sub
Public Sub gamspintot(zgamespintot As Long)
'Get a value for gamespintot
gamespintot = 0
For ct = 0 To 3
If gamespinkeep(ct) > 0 Then gamespintot = ct + 1 'initialise gamespintot
Next
zgamespintot = gamespintot
End Sub
Public Function outputvars()

outputvars = False

'Record vogchanges
If VOGchg = 0 Then
For ct = 29 To 34
GTp(ct) = gt(ct)
Next
VOGchg = -1
Else
For ct = 29 To 34
If gt(ct) <> GTp(ct) Then VOGchg = 1
Next
If VOGchg = 1 And gt(29) > 0 Then gt(29) = -gt(29)
End If

'Prevent error 21
If gt(10) > 0 And gt(11) > 0 And (fgct > 0 Or spct > 0) Then
ct = 88 + gt(11)
If Left(CStr(gt(ct)), 2) = "10" Then
gt(ct) = 10
Else
gt(ct) = CLng(Left(CStr(gt(ct)), 1))
End If
End If

On Error GoTo DBerror

sDatabaseName = loaddirectory & "Slotdata.s$t"

If OpenDb(sDatabaseName, 0) = False Then Exit Function
Set dbsCurrent = gdbCurrentDB

Set rectemp = dbsCurrent.OpenRecordset("SELECT Bullean FROM Inpoot")

With rectemp

For ct = 0 To 1

.Edit
![Bullean] = disablegamespintabs(ct)
.Update
.MoveNext
Next
For ct = 1 To 14
For ct1 = 1 To 14
.Edit
![Bullean] = substitute(ct, ct1)
.Update
.MoveNext
Next
For ct1 = 0 To 5
.Edit
![Bullean] = reelcheck(ct, ct1)
.Update
.MoveNext
Next
Next

End With

Set rectemp = dbsCurrent.OpenRecordset("SELECT Lorng FROM Inpoot")
With rectemp
.MoveFirst

.Edit
![Lorng] = thumbsize
.Update
.MoveNext
For ct = 0 To 4
For ct1 = 1 To 24
.Edit
![Lorng] = wheelorder(ct, ct1)
.Update
.MoveNext
Next
Next
For ct = 1 To 5
For ct1 = 1 To 14
.Edit
![Lorng] = wheelvec(ct, ct1)
.Update
.MoveNext
Next
Next
For ct = 1 To 14
For ct1 = 0 To 10
.Edit
![Lorng] = sst(ct, ct1)
.Update
.MoveNext
Next
Next
For ct = 1 To 200
.Edit
![Lorng] = gt(ct)
.Update
.MoveNext
Next
For ct = 0 To 1
For ct1 = 1 To 9
.Edit
![Lorng] = freegamesettings(ct, ct1)
.Update
.MoveNext
Next
For ct1 = 1 To 15
.Edit
![Lorng] = spinsettings(ct, ct1)
.Update
.MoveNext
Next
Next
For ct = 0 To 3
.Edit
![Lorng] = gamespinsymbol(ct)
.Update
.MoveNext
.Edit
![Lorng] = gamespinkeep(ct)
.Update
.MoveNext
Next


For ct = 1 To 65
.Edit
![Lorng] = Defaultgt(ct)
.Update
.MoveNext
Next

End With


Set rectemp = dbsCurrent.OpenRecordset("SELECT Streeng FROM Inpoot")
With rectemp
.MoveFirst


For ct = 1 To 100
.Edit
![Streeng] = Stringvars(ct)
.Update
.MoveNext
Next

For ct = 1 To 100
.Edit
![Streeng] = Defaultstr(ct)
.Update
If Not .EOF Then .MoveNext
Next

.Close

End With
Set gdbCurrentDB = dbsCurrent

killdb sDatabaseName

Set rectemp = Nothing
Set dbsCurrent = Nothing

If gt(29) < 0 Then gt(29) = -gt(29)
If compactdb(3) = False Then GoTo DBerror
outputvars = True
Exit Function
DBerror:
Set rectemp = Nothing
Set dbsCurrent = Nothing
MsgBox "Slotdata corrupt or inaccessible, cannot continue ..", vbOKOnly
killdb sDatabaseName
End Function
Public Sub savedefault()

ct1 = 0
For ct = 1 To 200
Select Case ct 'max 64 used in default
Case 1, 4 To 10, 12, 14, 15, 19 To 25, 36, 37, 44 To 47, 152, 153, 155, 157 To 187, 189, 190, 193 To 196
ct1 = ct1 + 1
Defaultgt(ct1) = gt(ct)
End Select
Next

For ct = 1 To 50
Defaultstr(ct) = Stringvars(ct)
Next
End Sub
Public Function IsDefaultSaved() As Boolean
If Defaultstr(5) = "" Then
IsDefaultSaved = False
Else
IsDefaultSaved = True
End If
End Function
Public Sub Restoresaved(fromscratch As Boolean)
ct1 = 0
For ct = 1 To 200
    Select Case ct
    Case 1, 4 To 10, 12, 14, 15, 19 To 25, 36, 37, 44 To 47, 153, 155, 157 To 187, 189, 190, 193 To 196
    ct1 = ct1 + 1
    gt(ct) = Defaultgt(ct1)
    Case 152
    ct1 = ct1 + 1
    If fromscratch = True Then gt(152) = Defaultgt(ct1)
    Case 156
    If gt(192) = 1 Then ct1 = ct1 + 1
    End Select
Next



For ct = 1 To 2
Stringvars(ct) = Defaultstr(ct)
Next

'gt(191)= 0 if stvr(3) = 0
If Stringvars(3) <> "" Then
Stringvars(3) = Defaultstr(3)
Else
'Quotes DB should be curdir if satisfied
If fromscratch = False Then Stringvars(3) = Defaultstr(3)
End If

For ct = 4 To 50
Stringvars(ct) = Defaultstr(ct)
Next

'Clear quotes if big pics
If gt(159) = 1 Then Stringvars(3) = ""

If gt(192) = 1 Then 'discard gt(156) for good
gt(192) = 5
savedefault
End If

End Sub
Private Sub initwinner()

For ct1 = 1 To thumbsize
bonuz(ct1) = 1
Next


For intreel = 0 To 4
If allocsub(intreel) = False Then

subshere(intreel) = False


Select Case currline
Case 0
spinztemp(intreel) = spinz(intreel)

If hreel(intreel) = True Then
Select Case currheld
Case 0
win1(intreel) = wheelorder(intreel, Advanz(spinztemp(intreel), 0))
win2(intreel) = wheelorder(intreel, Advanz(spinztemp(intreel), -dirspin(intreel)))
win3(intreel) = wheelorder(intreel, Advanz(spinztemp(intreel), -dirspin(intreel)))
Case 1
win2(intreel) = wheelorder(intreel, Advanz(spinztemp(intreel), 0))
win1(intreel) = wheelorder(intreel, Advanz(spinztemp(intreel), dirspin(intreel)))
win3(intreel) = wheelorder(intreel, Advanz(spinztemp(intreel), -2 * dirspin(intreel)))
Case 2
win3(intreel) = wheelorder(intreel, Advanz(spinztemp(intreel), -3 * dirspin(intreel)))
win2(intreel) = wheelorder(intreel, Advanz(spinztemp(intreel), dirspin(intreel)))
win1(intreel) = wheelorder(intreel, Advanz(spinztemp(intreel), dirspin(intreel)))
End Select
Else
win1(intreel) = wheelorder(intreel, Advanz(spinztemp(intreel), 0))
win2(intreel) = wheelorder(intreel, Advanz(spinztemp(intreel), -dirofspin))
win3(intreel) = wheelorder(intreel, Advanz(spinztemp(intreel), -dirofspin))
End If

PYLN(intreel) = win2(intreel)

Case 1
If dirspin(intreel) <> 1 Then
PYLN(intreel) = win3(intreel)
Else
PYLN(intreel) = win1(intreel)
End If
Case 2
If dirspin(intreel) <> 1 Then
PYLN(intreel) = win1(intreel)
Else
PYLN(intreel) = win3(intreel)
End If
End Select

End If
Next

End Sub
Public Sub sortsubstitutes(zprizecount() As Long, zPZs() As Long)
Dim subcountspare(6, 1) As Long, substno(1) As Long, competeno(1) As Long, tempmultprize(9) As Long
Dim oldcompg1(3) As Long, oldcompgG2 As Long, zuBz(14) As Long, nc(14, 14) As Boolean
Dim c1 As Long, c2 As Long, c3 As Long, c4 As Long, runtot As Long

For c1 = 0 To 2
prizecount(c1) = 0
Next

For currline = 0 To gt(153)
'PZs (currline,prizetype,c2) : c2 = subs or not,6-10,picno,bonus


For c1 = 0 To thumbsize
For c2 = 0 To thumbsize
nc(c1, c2) = True
Next
zuBz(c1) = 0
bonuz(c1) = 1
Next

For c1 = 0 To 3
oldcompg1(c1) = -1
Next

For c1 = 0 To 4
allocsub(c1) = False
Next


For c1 = 0 To 3
For c2 = 0 To 1
lonesubz(c1, c2) = 0
Next
Next


For c1 = 1 To 9
'max is 9 possible prizecounts
tempmultprize(c1) = 0
For c2 = 0 To 3
PZs(currline, c1, c2) = 0
Next
Next

For c1 = 0 To 1
competeno(c1) = 0
substno(c1) = 0
For c3 = 0 To 1
For c2 = 0 To 6
subcount(c1, c2, c3) = 0
Next
Next
Next

For c1 = 0 To 5
For c2 = 0 To 4
For c3 = 0 To 4
savevec(c1, c2, c3) = 0
Next
Next
Next

lz1 = 0
lz2 = 0
pztk1 = 0
pztk2 = 0
pztk3 = 0
c3 = 1
ct = 1
runtot = 0
showprize = False


initwinner

'Preserve original pay combo without touching held symbols, keep record of subs for FSFG
For c1 = 0 To 4
winsub(currline, c1) = PYLN(c1)
If heldnow(c1) <= 0 Then
win2old(currline, c1) = PYLN(c1)
ElseIf heldnow(c1) <> PYLN(c1) Then
'Check allocated subs
If substitute(PYLN(c1), heldnow(c1)) = True And currline = currheld And subsline(currline) = True Then
PYLN(c1) = heldnow(c1)
allocsub(c1) = True
End If
End If
Next



For pct = 1 To thumbsize


'first check for substitutes
For c1 = 0 To 4

'Here at least one substituter with at least one substitute required
    If substitute(PYLN(c1), pct) = True And reelcheck(PYLN(c1), c1 + 1) = True Then

    zuBz(PYLN(c1)) = 1 + zuBz(PYLN(c1))
    For c2 = 0 To 4
        If c1 <> c2 And pct = PYLN(c2) Then

        nc(PYLN(c1), PYLN(c2)) = False

        'Subcount(x,y,z) = subcount(Substituter max 4(group ID),competitor id, substituter/substituted)
        'Competitor range is 5
        Select Case c3
        Case 1

        subcount(0, 1, 0) = PYLN(c1)
        subcount(0, 1, 1) = PYLN(c2)
        c3 = 2
        oldcompg1(0) = c2 'used to track whether substitute is the same
        Case 2  'S1 C1 already (C = competitor)
        
            If subcount(0, 1, 0) = PYLN(c1) Then    'first group

            checksubs c1, c2, oldcompg1

            Else    'second group gets here if S1 C1 or S1 S1 C1 or S1 C1 C2
            subcount(1, 1, 0) = PYLN(c1)
            subcount(1, 1, 1) = PYLN(c2)
            c3 = 3
            oldcompgG2 = c2
            'if S1 S1 C1 or S1 C11 C1 previously, exit both loops anyway
            End If

        Case 3 'S1 C1 S2 C2 already
        If subcount(0, 1, 0) = PYLN(c1) Then
        
        subcount(0, 2, 0) = PYLN(c1)    'getting another S1
        
        'Any extra competitor for S1 would have been dealt with previously
        
        ElseIf subcount(1, 1, 0) = PYLN(c1) Then
        
            If c2 <> oldcompgG2 Then
            'still another competitor to count
            subcount(1, 2, 0) = PYLN(c1)
            subcount(1, 2, 1) = PYLN(c2)
            Else     'getting another S2
            subcount(1, 2, 0) = PYLN(c1)
            End If

        End If

        End Select
        End If
    Next
    End If
Next
Next


c3 = 0
For pct = 1 To thumbsize
If zuBz(pct) > 1 Then
For c1 = 1 To thumbsize
'no comps
If nc(pct, c1) = False Then zuBz(pct) = 0
Next
End If

If zuBz(pct) > 1 Then
lonesubz(c3, 0) = pct
c3 = c3 + 1
End If

Next



If subcount(0, 1, 0) > 0 Then   'process substitutes


'must be at least 2 nats
For c1 = 0 To 4
For c2 = 0 To 4
For temp = 1 To 6
If c1 <> c2 And PYLN(c1) = PYLN(c2) And PYLN(c1) = subcount(0, temp, 1) Then setbonus 0, PYLN(c1)
Next
Next
Next

competeno(0) = 0
    
    
    For c1 = 0 To 3
    If oldcompg1(c1) > -1 Then competeno(0) = competeno(0) + 1
    'competeno is number of competitors - alike or unlike
    Next


    ct = 1
    For c1 = 2 To 5
    If subcount(0, c1, 0) > 0 And subcount(0, c1, 1) = 0 Then ct = ct + 1
    Next
    If ct = 0 Then ct = 1
    substno(0) = ct 'substno(0) is the number of same substitutes in group 1


If subcount(1, 1, 0) > 0 Then  'for group 2
    If subcount(1, 2, 0) > 0 Then
    If subcount(1, 1, 1) > 0 Then
    competeno(1) = 2
    Else
    substno(1) = 2
    End If
    Else
    competeno(1) = 1
    substno(1) = 1
    End If

End If



    'Now assign substitutes their values

If competeno(1) = 0 Then
        
        Select Case substno(0) 'substno is no of substitutes in group 1
        
        
        Case 1  'only one substitute symbol
        For c1 = 0 To competeno(0) 'sum through different competitors
        countsingle c1
        Next
        
        Compareprizes
        countsingle pztk1

        Case 2  'two (same) substitute symbols  'max 3 competitors
        
        For c1 = 0 To 2 * competeno(0) - 1
        For c2 = 0 To 2 * competeno(0) - 1
        countpair c1, c2
        Next
        Next
        
        Compareprizes
        countpair pztk1, pztk2


        Case 3  '3 (same) substitute symbols
        
        For c1 = 0 To 2 * competeno(0) - 1
        For c2 = 0 To 2 * competeno(0) - 1
        For c3 = 0 To 2 * competeno(0) - 1
        counttriple c1, c2, c3
        Next
        Next
        Next
        
        Compareprizes
        counttriple pztk1, pztk2, pztk3

        Case 4  '4 (same) substitute symbols

        initwinner    'do all substitutions
        For c1 = 0 To 4
        If subcount(0, 1, 0) = PYLN(c1) Then
        PYLN(c1) = subcount(0, 1, 1)
        subshere(c1) = True
        Else
        setbonus 1, PYLN(c1)
        End If
        Next
        'Nothing to compare
        showprize = True
        Payprize

        End Select

Else 'competeno(1) > 0 2 groups

        If competeno(0) = 2 Or competeno(1) = 2 Then 'Only room for 1 substituter

            'if competeno(0) = 2, fine but ....

            If competeno(1) = 2 Then      '
            For c2 = 1 To 6
            For c3 = 0 To 1
            subcountspare(c2, c3) = subcount(0, c2, c3)
            subcount(0, c2, c3) = subcount(1, c2, c3)
            subcount(1, c2, c3) = subcountspare(c2, c3)
            Next
            Next

            temp = substno(0)
            substno(0) = substno(1)
            substno(1) = temp
            
            temp = competeno(0)
            competeno(0) = competeno(1)
            competeno(1) = temp
            End If
            
        
            '2 competitors in group one, dispense with substno (1)
            For c1 = 0 To 2
            For c2 = 0 To 2
            For c3 = 0 To 1
            countpair2Gs c1, c2, c3
            Next
            Next
            Next


        ElseIf substno(0) = 2 Or substno(1) = 2 Then 'Only room for 1 substitute

            If substno(1) = 2 Then     '
            For c2 = 1 To 6
            For c3 = 0 To 1
            subcountspare(c2, c3) = subcount(0, c2, c3)
            subcount(0, c2, c3) = subcount(1, c2, c3)
            subcount(1, c2, c3) = subcountspare(c2, c3)
            Next
            Next

            temp = substno(0)
            substno(0) = substno(1)
            substno(1) = temp
            
            temp = competeno(0)
            competeno(0) = competeno(1)
            competeno(1) = temp
            End If
            
            For c1 = 0 To 1
            For c2 = 0 To 1
            For c3 = 0 To 1
            countpair2Gs c1, c2, c3
            Next
            Next
            Next

        Else 'the two groups have 1 substitute & 1 competitor each
            
            For c1 = 0 To 1
            For c2 = 0 To 1
            countpair2Gs c1, 0, c2
            Next
            Next
        
        End If
        Compareprizes
        countpair2Gs pztk1, pztk2, pztk3
    
    
End If  'competeno condition
    
   
ElseIf lonesubz(0, 0) > 0 Then 'lone subs

Payprize    'No subs here as yet

Compareprizes

initwinner

For lz1 = 0 To 3
If lonesubz(lz1, 1) > 0 Then DOsubz1 lz1
Next


Payprize



Else    '0 subs
showprize = True
Payprize
End If


If gt(5) > 0 Then
'Naturals bonus

For ct2 = 1 To prizecount(currline)
    If PZs(currline, ct2, 3) = 0 Then 'if substituter detected in prize; exit
    For ct1 = 1 To thumbsize
    ' symbol on reel, not a single and a substituted symbol with no substitute on current payline!
    For ct = 1 To 5
    If PYLN(ct - 1) = PZs(currline, ct2, 2) And PZs(currline, ct2, 1) <> 10 And substitute(ct1, PZs(currline, ct2, 2)) = True And reelcheck(ct1, ct) = True And wheelvec(ct, ct1) > 0 Then PZs(currline, ct2, 3) = 1  'activate bonus
    Next
    Next
    Else
    PZs(currline, ct2, 3) = 0
    End If
Next

End If


If gt(9) = 0 Then  'highest win pays

For c1 = 1 To thumbsize
bonuz(c1) = 0
Next

    For ct = 1 To prizecount(currline)
        
        pct = PZs(currline, ct, 2)
        ct1 = sst(pct, PZs(currline, ct, 1))
        If PZs(currline, ct, 3) = 1 Then
        bonuz(pct) = gt(4) + gt(5)
        tempmultprize(ct) = ct1 * bonuz(pct)
        Else    'not in a substitute family
        tempmultprize(ct) = ct1
        End If
    Next
        If prizecount(currline) > 0 Then
        
        ShellSort tempmultprize, prizecount(currline)
        c1 = tempmultprize(prizecount(currline))
        'Now identify the prize
         For ct = 1 To prizecount(currline)
         pct = PZs(currline, ct, 2)
         ct1 = sst(pct, PZs(currline, ct, 1))
                If bonuz(pct) > 0 Then
                PZs(currline, runtot + 1, 3) = 1
                Else
                PZs(currline, runtot + 1, 3) = 0
                End If

                If (bonuz(pct) > 0 And c1 = bonuz(pct) * ct1) Or (bonuz(pct) = 0 And c1 = ct1) Then
                runtot = runtot + 1 'case of equal prizes
                PZs(currline, runtot, 1) = PZs(currline, ct, 1)
                PZs(currline, runtot, 2) = pct
                End If
        Next
        prizecount(currline) = runtot
        End If
End If


'Return values to Pokemach
For ct1 = 1 To 9
For ct2 = 0 To 3
zPZs(currline, ct1, ct2) = PZs(currline, ct1, ct2)
Next
Next

zprizecount(currline) = prizecount(currline)

Next 'currline

Select Case gt(153)
Case 0
zprizecount(1) = 0
zprizecount(2) = 0
Case 1
zprizecount(2) = 0
End Select

End Sub
Private Sub checksubs(cc1 As Long, cc2 As Long, oldcompg1() As Long)
If cc2 = oldcompg1(0) Then
'cc1,cc2 through 0 to 4, win(cc1) subs for win(cc2)
ct = ct + 1 'ct is "static" here
subcount(0, ct, 0) = PYLN(cc1)
ElseIf cc2 <> oldcompg1(0) And oldcompg1(1) = -1 Then
'first , second or third substitute, add another competitor
ct = ct + 1 'ct is "static" here
subcount(0, ct, 1) = PYLN(cc2)
oldcompg1(1) = cc2
subcount(0, ct, 0) = PYLN(cc1)
ElseIf cc2 <> oldcompg1(0) And cc2 <> oldcompg1(1) And oldcompg1(2) = -1 Then
ct = ct + 1
subcount(0, ct, 1) = PYLN(cc2)
oldcompg1(2) = cc2
subcount(0, ct, 0) = PYLN(cc1)
ElseIf cc2 <> oldcompg1(0) And cc2 <> oldcompg1(1) And cc2 <> oldcompg1(2) And oldcompg1(3) = -1 Then
ct = ct + 1
subcount(0, ct, 1) = PYLN(cc2)
oldcompg1(3) = cc2
subcount(0, ct, 0) = PYLN(cc1)
End If
End Sub
Private Sub Payprize()
Dim Stot As Long, betterthansingle As Boolean, Singleonly As Boolean, bonus As Boolean
'note subshere is for GROUP 1 only

If showprize = True Then prizecount(currline) = 0


'loop for prizes

For pct = 1 To thumbsize
bonus = False

'scatter pays
If sst(pct, 2) = 1 Then
    If currline = 0 Then

    Stot = 0

        If sst(pct, 5) = 1 Then 'Anys
    
        For intreel = 0 To 4
        If win1(intreel) = pct Then Stot = Stot + 1
        If win2(intreel) = pct Then Stot = Stot + 1
        If win3(intreel) = pct Then Stot = Stot + 1
        Next
        If Stot > 5 Then Exit Sub
        If Stot > 0 Then addprize 11 - Stot, pct, False

        Else    'left to right- can't have a scatter single prize
        Stot = 0
        For intreel = 0 To 4
                If win1(intreel) = pct Or win2(intreel) = pct Or win3(intreel) = pct Then
                Stot = Stot + 1
                If Stot > 5 Then Exit Sub
                If intreel = 4 Then addprize 6, pct, False
        
                Else
                If intreel > 1 Then addprize 11 - Stot, pct, False
                Exit For
                End If
        Next

        If sst(pct, 3) = 1 And Stot < 3 Then 'Right to left
            Stot = 0
            For intreel = 4 To 0 Step -1
            If win1(intreel) = pct Or win2(intreel) = pct Or win3(intreel) = pct Then
            Stot = Stot + 1
            If Stot > 5 Then Exit Sub
            Else
            If Stot > 1 Then addprize 11 - Stot, pct, False
            Exit For
            End If
            Next
        End If
        End If 'End Scatter Anys if

    End If  'currline
Else 'Line pays

temp = 1
betterthansingle = False
Singleonly = False
If PYLN(4) = pct Then Singleonly = True
For intreel = 0 To 3
If betterthansingle = True Then
'Exit here
Exit For
End If
If PYLN(intreel) = pct Then
Singleonly = True
For ct = intreel + 1 To 4
If PYLN(ct) = PYLN(intreel) Then
temp = temp + 1
betterthansingle = True
Singleonly = False
End If
Next
End If
Next

If Singleonly = True And betterthansingle = False Then
'process single prize
If sst(pct, 5) = 1 Then
'pay 'any'
addprize 10, pct, False
ElseIf PYLN(0) = pct Then
addprize 10, pct, False
ElseIf sst(pct, 3) = 1 And PYLN(4) = pct Then
addprize 10, pct, False
End If

ElseIf Singleonly = False And betterthansingle = True Then
'Now for doubles


Select Case temp

Case 2
If sst(pct, 5) = 1 Then
'pay 'any'

For intreel = 0 To 4
If PYLN(intreel) = pct And subshere(intreel) = True Then bonus = True
Next

addprize 9, pct, bonus
ElseIf PYLN(0) = pct Then
'pair
If PYLN(1) = pct Then
If subshere(0) = True Or subshere(1) = True Then bonus = True
addprize 9, pct, bonus
Else
'single
If PYLN(4) = pct And sst(pct, 3) = 1 Then addprize 10, pct, False
addprize 10, pct, False
End If
ElseIf sst(pct, 3) = 1 Then
'Right to left
If PYLN(4) = pct Then
If PYLN(3) = pct Then
'pair
If subshere(3) = True Or subshere(4) = True Then bonus = True
addprize 9, pct, bonus
Else
'single
addprize 10, pct, False
End If
End If
End If

Case 3
If sst(pct, 5) = 1 Then
'pay 'any' prize

For intreel = 0 To 4
If PYLN(intreel) = pct And subshere(intreel) = True Then bonus = True
Next

addprize 8, pct, bonus
ElseIf PYLN(0) = pct Then
'Left to right
If PYLN(1) = pct Then
    If PYLN(2) = pct Then
    'triple
    If subshere(0) = True Or subshere(1) = True Or subshere(2) = True Then bonus = True
    addprize 8, pct, bonus
    Else
    'double
    If PYLN(4) = pct And sst(pct, 3) = 1 Then addprize 10, pct, False
    If subshere(0) = True Or subshere(1) = True Then bonus = True
    addprize 9, pct, bonus
    End If
Else
'single
addprize 10, pct, False
    If sst(pct, 3) = 1 Then
    If PYLN(4) = pct Then
    If PYLN(3) = pct Then
    If subshere(3) = True Or subshere(4) = True Then bonus = True
    addprize 9, pct, bonus
    Else
    addprize 10, pct, False
    End If
    End If
    End If
End If

'Right to left
ElseIf sst(pct, 3) = 1 Then
If PYLN(4) = pct Then
If PYLN(3) = pct Then
    If PYLN(2) = pct Then
    If subshere(2) = True Or subshere(3) = True Or subshere(4) = True Then bonus = True
    addprize 8, pct, bonus
    Else
    'double
    If subshere(3) = True Or subshere(4) = True Then bonus = True
    addprize 9, pct, bonus
    End If
Else
'single
addprize 10, pct, False
End If
End If
End If

If sst(pct, 4) = 1 Then
'middle threes
If PYLN(0) <> pct And PYLN(4) <> pct Then
If subshere(1) = True Or subshere(2) = True Or subshere(3) = True Then bonus = True
addprize 8, pct, bonus
End If
End If


Case 4
If sst(pct, 5) = 1 Then
'pay 'any' prize

For intreel = 0 To 4
If PYLN(intreel) = pct And subshere(intreel) = True Then bonus = True
Next

addprize 7, pct, bonus
ElseIf PYLN(0) = pct And PYLN(1) = pct Then
If PYLN(2) = pct Then
    If PYLN(3) = pct Then
    'This combo always pays
    If subshere(0) = True Or subshere(1) = True Or subshere(2) = True Or subshere(3) = True Then bonus = True
    addprize 7, pct, bonus
    Else
    'triple + single
    If sst(pct, 3) = 1 Then addprize 10, pct, False
    If subshere(0) = True Or subshere(1) = True Or subshere(2) = True Then bonus = True
    addprize 8, pct, bonus
    End If
Else
    ' 2 pair
    If sst(pct, 3) = 1 Then
    If subshere(3) = True Or subshere(4) = True Then
    bonus = True
    Else
    bonus = False
    End If
    addprize 9, pct, bonus
    End If
bonus = False
If subshere(0) = True Or subshere(1) = True Then bonus = True
addprize 9, pct, bonus
End If
ElseIf PYLN(2) = pct And PYLN(3) = pct And PYLN(4) = pct Then
    If sst(pct, 3) = 1 Then
        If PYLN(1) = pct Then
        '4's r-l
        If subshere(1) = True Or subshere(2) = True Or subshere(3) = True Or subshere(4) = True Then bonus = True
        addprize 7, pct, bonus
        Else    'PYLN(1) <> pct
        If subshere(2) = True Or subshere(3) = True Or subshere(4) = True Then bonus = True
        addprize 8, pct, bonus
        addprize 10, pct, False
        End If
    ElseIf PYLN(0) = pct Then
    'n/s trip + sing
    addprize 10, pct, False
    End If
End If

Case 5
If subshere(0) = True Or subshere(1) = True Or subshere(2) = True Or subshere(3) = True Or subshere(4) = True Then bonus = True
addprize 6, pct, bonus
End Select

End If 'Singleonly condition

End If 'Scatter condition
'End loop
Next

End Sub
Private Sub addprize(prizetype As Long, pictureno As Long, bonus As Boolean)
'Definitely non zero prizes
If sst(pictureno, prizetype) = 0 Then Exit Sub

If showprize = True Then

prizecount(currline) = prizecount(currline) + 1

PZs(currline, prizecount(currline), 1) = prizetype
PZs(currline, prizecount(currline), 2) = pictureno
If bonus = True Then
PZs(currline, prizecount(currline), 0) = 1
PZs(currline, prizecount(currline), 3) = 1
Else
PZs(currline, prizecount(currline), 3) = 0
End If

Else

If lz1 = 0 Then
savevec(pztk1, pztk2, pztk3) = savevec(pztk1, pztk2, pztk3) + bonuz(pictureno) * sst(pictureno, prizetype)
Else
lonz(lz1 - 1, lz2) = bonuz(pictureno) * sst(pictureno, prizetype) + lonz(lz1 - 1, lz2)
End If

End If
End Sub
Private Sub Compareprizes()
Dim maxx As Long

maxx = 0


If lonesubz(0, 0) > 0 Then lzz


'Substitute prizes have preference so sum backwards
For ct1 = 4 To 0 Step -1
For ct2 = 4 To 0 Step -1
For ct3 = 4 To 0 Step -1
If savevec(ct1, ct2, ct3) > maxx Then
maxx = savevec(ct1, ct2, ct3)
pztk1 = ct1
pztk2 = ct2
pztk3 = ct3
End If
Next
Next
Next


For ct1 = 0 To 3
lz2 = 0


For ct2 = 0 To thumbsize
If lonz(ct1, ct2) > maxx Then
maxx = lonz(ct1, ct2)
lz2 = ct2
End If
lonz(ct1, ct2) = 0  'convenient to zero here
Next

lonesubz(ct1, 1) = lz2

Next
showprize = True
End Sub
Private Sub lzz()

For lz1 = 1 To 4
If lonesubz(lz1 - 1, 0) > 0 Then
    
    
    For lz2 = 0 To thumbsize

    If substitute(lonesubz(lz1 - 1, 0), lz2) = True Then

    lonesubz(lz1 - 1, 1) = lz2
    
    initwinner
    
    If lz2 > 0 Then DOsubz1 lz1 - 1

    Payprize
    End If

    Next


End If
Next


End Sub
Private Sub DOsubz1(subzno As Long)
For ct2 = 0 To 4
If allocsub(ct2) = True Then
setbonus 1, PYLN(ct2)
subshere(ct2) = True
ElseIf PYLN(ct2) = lonesubz(subzno, 0) And reelcheck(PYLN(ct2), ct2 + 1) = True Then
PYLN(ct2) = lonesubz(subzno, 1)
setbonus 1, PYLN(ct2)
subshere(ct2) = True
End If
Next
End Sub
Private Function DOsubz2(X As Long, Optional startct As Long = -1, Optional nextct As Long = 4)
For ct2 = startct + 1 To nextct
If allocsub(ct2) = True Then
setbonus 1, PYLN(ct2)
Else
DOsubz2 = (subcount(0, X, 0) = PYLN(ct2) And reelcheck(PYLN(ct2), ct2 + 1))
    If DOsubz2 = True Then
        If subcount(0, X, 1) > 0 Then
        PYLN(ct2) = subcount(0, X, 1)
        setbonus 1, PYLN(ct2)
        Else
        DOsubz2 = False
        End If

    nextct = ct2
    subshere(ct2) = True
    Exit For
    End If
End If
Next
End Function
Private Sub countsingle(X As Long)
'just 1 substituter, X is competitor id, 0 for no substitution


initwinner

If showprize = True Then
For lz1 = 0 To 3
If lonesubz(lz1, 1) > 0 Then DOsubz1 lz1
Next
End If

If X > 0 Then DOsubz2 X

pztk1 = X

Payprize


End Sub
Private Sub countpair(X As Long, Y As Long)
Dim firstsub As Long
'2 substituters, X or Y,3 comps


firstsub = 3
initwinner

If showprize = True Then
For lz1 = 0 To 3
If lonesubz(lz1, 1) > 0 Then DOsubz1 lz1
Next
End If



If X > 0 Then
DOsubz2 X, , firstsub
Else
firstsub = 0
End If


If Y > 0 Then DOsubz2 Y, firstsub

pztk1 = X
pztk2 = Y

Payprize

End Sub
Private Sub counttriple(X As Long, Y As Long, Z As Long)
'3 substituters, max of 2 competitors
Dim firstsub As Long, secsub As Long


firstsub = 2
secsub = 3
initwinner



If X > 0 Then
DOsubz2 X, , firstsub
Else
firstsub = 0
End If

If Y > 0 Then
DOsubz2 Y, firstsub, secsub
Else
    If X = 0 Then
    secsub = 0
    Else
    secsub = firstsub + 1
    End If
End If


If Z > 0 Then DOsubz2 Z, secsub


pztk1 = X
pztk2 = Y
pztk3 = Z


Payprize

End Sub
Private Sub countpair2Gs(X As Long, Y As Long, Z As Long)
Dim firstsub As Long
'Now count substitute pair combos   -note both X, Y can be Zero or 1 or 2


firstsub = 3
initwinner



If X > 0 Then
DOsubz2 X, , firstsub
Else
firstsub = 0
End If

If Y > 0 Then
DOsubz2 Y, firstsub
Else

End If

'now  for second group

If Z > 0 Then
For ct1 = 0 To 4
If subcount(1, 1, 0) = PYLN(ct1) And reelcheck(PYLN(ct1), ct1 + 1) = True Then
PYLN(ct1) = subcount(1, 1, 1)
setbonus 1, subcount(1, 1, 1)
subshere(ct1) = True
Exit For
End If
Next
End If

pztk1 = X
pztk2 = Y
pztk3 = Z
Payprize


End Sub
Private Sub setbonus(subz As Long, Pik As Long)
If Pik = 0 Then Exit Sub
If subz = 0 Then
If gt(5) > 0 Then bonuz(Pik) = gt(5)
Else
If gt(4) > 0 Then bonuz(Pik) = gt(4)
End If
End Sub
Public Sub eligiblespingame(zfgct As Long, zspct As Long, zkept1or2 As Long, zprizetotal As Long, zgamenotspin As Boolean, zgamesaved As Boolean, zspinsaved As Boolean, zhavereset As Boolean)
Dim anotherpz As Boolean, Kurrloop As Long, heldd(4) As Long
gamesaved = False
spinsaved = False
juststarted = False


'Now determine free game, spin eligibility
'2nd last element of hq is pic-id
If gamespintot > 0 Or chtr <= 0 Then

fgct = zfgct
spct = zspct
prizetotal = zprizetotal

For intreel = 0 To 4
heldd(intreel) = heldnow(intreel)
Next

For currline = 0 To gt(153)

havereset = False
For intreel = 0 To 4
ohqgs(currline, intreel) = hq(currline, intreel)

If heldnow(intreel) > 0 Then
    If currline = currheld Then
    heldnow(intreel) = heldd(intreel)
    Else
    heldnow(intreel) = winsub(currline, intreel)
    End If
End If
Next


For ct = gamespintot To 1 Step -1   'according to specified priority

pct = gamespinkeep(ct - 1)

If currline = 0 Or sst(pct, 2) = 0 Then 'scatters only considered on 0 payline


'Now check & assign subs according to specified priority

For intreel = 0 To 4
anotherpz = False

'win2 can change for each pct
If heldnow(intreel) <= 0 Then
PYLN(intreel) = win2old(currline, intreel)
Else
'Must replace PYLN with substituter
PYLN(intreel) = winsub(currline, intreel)
End If
Next

For intreel = 0 To 4
If substitute(pct, PYLN(intreel)) = True And reelcheck(pct, intreel + 1) = True Then
For ct1 = 1 To prizecount(currline)
If PYLN(intreel) = PZs(currline, ct1, 2) And PZs(currline, ct1, 0) = 1 Then
For ct2 = 0 To 4
If PYLN(ct2) = pct Then
PYLN(ct2) = PYLN(intreel)
subsline(currline) = True
anotherpz = True
End If
Next
End If
If anotherpz = True Then Exit For
Next
If anotherpz = True Then Exit For
'reassign substitutes where apt
ElseIf substitute(PYLN(intreel), pct) = True And reelcheck(PYLN(intreel), intreel + 1) = True And heldnow(intreel) <= 0 And subsused(intreel) = False Then

'If PYLN(intreel) used in prize NOT containg pct then exit
For ct2 = 1 To prizecount(currline)
If PZs(currline, ct2, 0) = 1 Then
For ct1 = 0 To 4
'if PYLN(ct1) is substituted by PYLN(intreel) and has won a prize
If PYLN(ct1) <> pct And PZs(currline, ct2, 2) = PYLN(ct1) And substitute(PYLN(intreel), PYLN(ct1)) = True And reelcheck(PYLN(intreel), ct1 + 1) = True Then anotherpz = True
Next
End If
Next

If anotherpz = True Then Exit For

PYLN(intreel) = pct
subsline(currline) = True

End If
Next

'At each pass accumulate spins/games and check according to priority
'poss of other spin/game combintion(s)
If gamespinsymbol(0) = pct Then  'free games
Freegamecheck 0, pct
ElseIf gamespinsymbol(1) = pct Then
Freegamecheck 1, pct
ElseIf gamespinsymbol(2) = pct Then  'free spins
spincheck 0, pct
ElseIf gamespinsymbol(3) = pct Then
spincheck 1, pct
End If

End If  'currline, scatter

Next

If havereset = True Then Exit For
If currline = currheld Then
For intreel = 0 To 4
heldd(intreel) = heldnow(intreel)
Next
End If

Next


zhavereset = havereset

For intreel = 0 To 4
If havereset = False And heldd(intreel) <> 0 Then heldnow(intreel) = heldd(intreel)
subsused(intreel) = False
Next


zgamesaved = gamesaved
zspinsaved = spinsaved
zfgct = fgct
zspct = spct
zgamenotspin = gamenotspin
zkept1or2 = kept1or2


End If  'gamespintot > 0

If gamenotspin = True And fgct > 0 Then
For ct = 0 To 6
hq(currheld, ct) = holdq(currheld, kept1or2, 0, ct)
Next
ElseIf spct > 0 And gamenotspin = False Then
For ct = 0 To 6
hq(currheld, ct) = holdq(currheld, kept1or2 + 2, 0, ct)
Next
Else
For ct = 0 To 6
hq(currheld, ct) = 0
Next
End If



End Sub
Private Sub Freegamecheck(test1or2 As Long, pictemp As Long)
Dim subster As Boolean, centrewin As Long
subster = False

'we are already in free game reset, don't want to resave old free game symbols
If havereset = True Then Exit Sub

If freegamegen(test1or2, pictemp) = False Then Exit Sub

For ct1 = 0 To 4
If substitute(winsub(currline, ct1), pictemp) = True And reelcheck(winsub(currline, ct1), ct1 + 1) = True Then subster = True
Next

'quit if substitutes allowed
If freegamesettings(test1or2, 4) = 1 And subster = True Then Exit Sub



'preserve higher ranking held symbols
If juststarted = True And freegamesettings(test1or2, 5) = 0 Then Exit Sub


If fgct = 0 Then
        If spct > 0 Then 'suppose we are spinning
                If freegamesettings(test1or2, 5) = 0 Then        'break the sequence
                
                'Reset free spin counters
                spct = 0
                'Reduce the spin queue by one
                qtrakker(currheld, kept1or2 + 2) = qtrakker(currheld, kept1or2 + 2) - 1
                For ct1 = 0 To 6
                holdq(currheld, kept1or2 + 2, 0, ct1) = 0
                Next


                'Set game counters
                currheld = currline
                kept1or2 = test1or2
                fgct = 1
                havereset = True
                gamenotspin = True
                holdqstore kept1or2, 0
                For ct1 = 0 To 4
                heldnow(ct1) = holdqnew(ct1)
                Next
                
                Else    'preserve cumulative sequence total
                    If qtrakker(currline, test1or2) < freegamesettings(test1or2, 9) Then
                    gamesaved = True
                    qtrakker(currline, test1or2) = qtrakker(currline, test1or2) + 1
                    holdqstore test1or2, qtrakker(currline, test1or2)
                    Else
                    'give next pic a chance
                    For ct1 = 0 To 4
                    subsused(ct1) = False
                    Next
                    End If
                End If
        Else    'no current gamespin sequence qtrakker = 1 for first game
        juststarted = True
        kept1or2 = test1or2
        fgct = 1
        currheld = currline
        gamenotspin = True
        holdqstore kept1or2, 0
        qtrakker(currline, kept1or2) = 1
        For ct1 = 0 To 4
        heldnow(ct1) = holdqnew(ct1)
        Next
        End If
Else    'Already having a free game
        If freegamesettings(test1or2, 5) = 0 Then ' break the old game

        'gamenotspin already True
        qtrakker(currheld, kept1or2) = qtrakker(currheld, kept1or2) - 1
        fgct = 1
        currheld = currline
        kept1or2 = test1or2
        holdqstore kept1or2, 0
        havereset = True
        For ct1 = 0 To 4
        heldnow(ct1) = holdqnew(ct1)
        Next
        
        qtrakker(currline, kept1or2) = 1
        Else    'keep old game, save queue
            If qtrakker(currline, test1or2) < freegamesettings(test1or2, 9) Then
            gamesaved = True
            If test1or2 <> kept1or2 Or currline <> currheld Then qtrakker(currline, test1or2) = qtrakker(currline, test1or2) + 1
            holdqstore test1or2, qtrakker(currline, test1or2)
            If test1or2 = kept1or2 And currline = currheld Then qtrakker(currline, test1or2) = qtrakker(currline, test1or2) + 1
            Else
            'give next pic a chance
            For ct1 = 0 To 4
            subsused(ct1) = False
            Next
            End If
        End If
End If
End Sub
Private Function freegamegen(test1or2 As Long, pic As Long)
Dim gamechoice As Long, piccase(4) As Long, ct4 As Long
freegamegen = False
For ct1 = 0 To 4
piccase(ct1) = 0
Next


If freegamesettings(test1or2, 1) = 1 Then '1 if L - R, 2 if any

piccase(0) = pic
piccase(1) = pic
piccase(2) = pic


    If Winxcompare(piccase, -pic, 0) = True Then
    freegamegen = True
    Exit Function
    End If


ElseIf freegamesettings(test1or2, 1) = 2 Then    'any

For ct1 = 0 To 2
piccase(ct1) = pic

For ct2 = ct1 + 1 To 3
piccase(ct2) = pic
For ct3 = ct2 + 1 To 4
piccase(ct3) = pic

For ct4 = 0 To 4
If ct4 <> ct1 And ct4 <> ct2 And ct4 <> ct3 Then piccase(ct4) = 0
Next

    If Winxcompare(piccase, -pic, 0) = True Then
    freegamegen = True
    Exit Function
    End If

Next
Next
Next

End If


For gamechoice = 2 To 3
If freegamesettings(test1or2, gamechoice) > 0 Then

For ct1 = 0 To 4
piccase(ct1) = 0
Next

Select Case gamechoice
Case 3  'middle 3s
piccase(1) = pic
piccase(2) = pic
piccase(3) = pic

    If Winxcompare(piccase, -pic, 0) = True Then
    freegamegen = True
    Exit Function
    End If
Case 2  'R - L
piccase(2) = pic
piccase(3) = pic
piccase(4) = pic

    If Winxcompare(piccase, -pic, 0) = True Then
    freegamegen = True
    Exit Function
    End If

End Select
End If
Next


End Function
Private Sub spincheck(test1or2 As Long, pictemp As Long)
Dim subster As Boolean, substed As Boolean, centrewin As Long
subster = False

'we are already in free game reset, don't want to resave old free spin symbols
If havereset = True Then Exit Sub


If spinsettings(test1or2, 9) = 1 Then 'don't bother if prize won
For ct1 = 1 To prizecount(currline)
If PZs(currline, ct1, 2) = pictemp Then Exit Sub
Next
End If

If spingen(test1or2, pictemp) = False Then Exit Sub



For ct1 = 0 To 4
If substitute(winsub(currline, ct1), pictemp) = True And reelcheck(winsub(currline, ct1), ct1 + 1) = True And holdqnew(ct1) > 0 Then subster = True
Next

'substitutes mixed with substituters
If subster = True And spinsettings(test1or2, 10) = 1 Then Exit Sub






'preserve higher ranking held symbols on first try
If juststarted = True And spinsettings(test1or2, 11) = 0 Then Exit Sub


If spct = 0 Then
        
        If fgct > 0 Then   'Suppose we are gaming
            If spinsettings(test1or2, 11) = 0 Then       'new spin, scratch old game
            
            
            'finish game sequence
            qtrakker(currheld, kept1or2) = qtrakker(currheld, kept1or2) - 1
            
            For ct1 = 0 To 6
            holdq(currheld, kept1or2, 0, ct1) = 0
            Next
            havereset = True
            
            fgct = 0
            currheld = currline
            'Now set spin
            spct = 1
            gamenotspin = False
            kept1or2 = test1or2
            holdqstore kept1or2 + 2, 0
            qtrakker(currline, kept1or2 + 2) = 1
            For ct1 = 0 To 4
            heldnow(ct1) = holdqnew(ct1)
            Next
            
            Else    'preserve cumulative sequence total
            temp = qtrakker(currline, test1or2 + 2) + 1
                If temp <= spinsettings(test1or2, 15) Then
                qtrakker(currline, test1or2 + 2) = temp
                spinsaved = True
                holdqstore test1or2 + 2, temp
                Else
                'give next pic a chance
                For ct1 = 0 To 4
                subsused(ct1) = False
                Next
                End If
        End If
        Else    'new spin from scratch
        juststarted = True
        spct = 1
        currheld = currline
        gamenotspin = False
        kept1or2 = test1or2
        holdqstore kept1or2 + 2, 0
        qtrakker(currline, kept1or2 + 2) = 1
        For ct1 = 0 To 4
        heldnow(ct1) = holdqnew(ct1)
        Next
        End If
Else    'Now if we are spinning
    If spinsettings(test1or2, 11) = 0 Then 'spct > 0 and can have new spin while spinning
        
        
        'gamenotspin already False
        qtrakker(currheld, kept1or2 + 2) = qtrakker(currheld, kept1or2 + 2) - 1
        For ct1 = 0 To 6
        holdq(currheld, kept1or2 + 2, 0, ct1) = 0
        Next
        currheld = currline
        havereset = True
        spct = 1
        kept1or2 = test1or2
        holdqstore kept1or2 + 2, 0
        qtrakker(currline, kept1or2 + 2) = 1
        For ct1 = 0 To 4
        heldnow(ct1) = holdqnew(ct1)
        Next
 
    Else    'Put spin in queue
        If qtrakker(currline, test1or2 + 2) < spinsettings(test1or2, 15) Then
        If test1or2 <> kept1or2 Or currline <> currheld Then qtrakker(currline, test1or2 + 2) = qtrakker(currline, test1or2 + 2) + 1
        spinsaved = True
        holdqstore test1or2 + 2, qtrakker(currline, test1or2 + 2)
        If test1or2 = kept1or2 And currline = currheld Then qtrakker(currline, test1or2 + 2) = qtrakker(currline, test1or2 + 2) + 1
        Else
        'give next pic a chance
        For ct1 = 0 To 4
        subsused(ct1) = False
        Next
        End If
    End If
End If
End Sub
Private Function spingen(test1or2 As Long, pic As Long)
Dim spingrade As Long, piccase(4) As Long, ct4 As Long, testcombany(5) As Long

spingen = False


For spingrade = 8 To 1 Step -1
If spinsettings(test1or2, spingrade) > 0 Then

For ct1 = 0 To 4
piccase(ct1) = 0
Next

Select Case spingrade
Case 8
piccase(1) = pic
piccase(2) = pic
piccase(3) = pic
piccase(4) = pic

If Winxcompare(piccase, pic, 0) = True Then
spingen = True
Exit Function
End If
Case 7
'spinsettings(test1or2, 7) = 0 ,1, 2 depending if
'0 required, 1 L - R or 2, any 4
If spinsettings(test1or2, 7) = 1 Then
piccase(0) = pic
piccase(1) = pic
piccase(2) = pic
piccase(3) = pic

If Winxcompare(piccase, pic, 0) = True Then
spingen = True
Exit Function
End If
ElseIf spinsettings(test1or2, 7) = 2 Then

For ct1 = 0 To 4

For ct2 = 0 To 4
piccase(ct2) = pic
Next

piccase(ct1) = 0
If Winxcompare(piccase, pic, 0) = True Then
spingen = True
Exit Function
End If
Next
End If
Case 6
piccase(1) = pic
piccase(2) = pic
piccase(3) = pic

If Winxcompare(piccase, pic, 0) = True Then
spingen = True
Exit Function
End If
Case 5
piccase(0) = -pic

piccase(2) = pic
piccase(3) = pic
piccase(4) = pic

If Winxcompare(piccase, pic, 1) = True Then
spingen = True
Exit Function
End If

piccase(0) = 0

If Winxcompare(piccase, pic, 0) = True Then
spingen = True
Exit Function
End If

Case 4
If spinsettings(test1or2, 4) = 1 Then
piccase(0) = pic
piccase(1) = pic
piccase(2) = pic
piccase(4) = -pic

    If Winxcompare(piccase, pic, 1) = True Then
    spingen = True
    Exit Function
    End If

piccase(4) = 0

    If Winxcompare(piccase, pic, 0) = True Then
    spingen = True
    Exit Function
    End If

ElseIf spinsettings(test1or2, 4) = 2 Then

For ct1 = 0 To 2
piccase(ct1) = pic

For ct2 = ct1 + 1 To 3
piccase(ct2) = pic
For ct3 = ct2 + 1 To 4
piccase(ct3) = pic

For ct4 = 0 To 4
If ct4 <> ct1 And ct4 <> ct2 And ct4 <> ct3 Then piccase(ct4) = 0
Next

If Winxcompare(piccase, pic, 0) = True Then
spingen = True
Exit Function
End If

Next
Next
Next

End If
Case 3
piccase(1) = pic
piccase(2) = pic


piccase(4) = -pic
If Winxcompare(piccase, pic, 1) = True Then
spingen = True
Exit Function
End If

piccase(4) = 0

If Winxcompare(piccase, pic, 0) = True Then
spingen = True
Exit Function
End If

piccase(0) = -pic

piccase(1) = 0
piccase(3) = pic
If Winxcompare(piccase, pic, 1) = True Then
spingen = True
Exit Function
End If

piccase(0) = 0

If Winxcompare(piccase, pic, 0) = True Then
spingen = True
Exit Function
End If

Case 2
piccase(0) = -pic
piccase(1) = -pic

piccase(3) = pic
piccase(4) = pic

For ct1 = 2 To 1 Step -1
piccase(ct1) = 0
If Winxcompare(piccase, pic, ct1) = True Then
spingen = True
Exit Function
End If
Next

piccase(0) = 0

For ct1 = 1 To 0 Step -1
piccase(1) = -ct1 * pic
If Winxcompare(piccase, pic, ct1) = True Then
spingen = True
Exit Function
End If
Next

Case 1
If spinsettings(test1or2, 1) = 1 Then

piccase(0) = pic
piccase(1) = pic

piccase(3) = -pic
piccase(4) = -pic

For ct1 = 2 To 3
piccase(ct1) = 0
If Winxcompare(piccase, pic, 4 - ct1) = True Then
spingen = True
Exit Function
End If
Next

piccase(4) = 0

For ct1 = 3 To 4
piccase(3) = (ct1 - 4) * pic
If Winxcompare(piccase, pic, 4 - ct1) = True Then
spingen = True
Exit Function
End If
Next

ElseIf spinsettings(test1or2, 1) = 2 Then
For ct1 = 0 To 3
piccase(ct1) = pic

For ct2 = ct1 + 1 To 4
piccase(ct2) = pic
    
    For ct3 = 0 To 4
    If ct3 <> ct1 And ct3 <> ct2 Then piccase(ct3) = 0
    Next
    
    If Winxcompare(piccase, pic, 0) = True Then
    spingen = True
    Exit Function
    End If


piccase(ct2) = 0
Next
piccase(ct1) = 0
Next

End If

End Select
End If
Next
End Function
Private Function Winxcompare(piccase() As Long, pictest As Long, possodds As Long)
Dim ptest As Long, testodds As Long, rltotal As Long, testtotal As Long, c1 As Long
'Piccase has prize symbol values slotted for creating test on perfect match
rltotal = 0
testtotal = 0
testodds = 0
Winxcompare = False

If pictest < 0 Then
ptest = -pictest    'testing for free game
Else
ptest = pictest
End If

'clear old test values
For c1 = 0 To 4
If piccase(c1) > 0 Then
If spct > 0 And heldnow(c1) > 0 Then Exit Function
rltotal = rltotal + 1
End If
holdqnew(c1) = 0
Next


If sst(ptest, 2) = 1 Then    'For scatters

For c1 = 0 To 4
'if we are spinning,we DO NOT wish to include as a saved spin any of the same symbols in HELD combinations in piccase

If win1(c1) = ptest Or win2(c1) = ptest Or win3(c1) = ptest Then
    If piccase(c1) > 0 Then
    holdqnew(c1) = piccase(c1)
    testtotal = testtotal + 1
    ElseIf piccase(c1) < 0 Then
        If heldnow(c1) <= 0 Then testodds = testodds + 1
        holdqnew(c1) = piccase(c1)
    Else
        If pictest > 0 Then
        If spct = 0 Then
        Exit Function
        ElseIf heldnow(c1) <= 0 And pictest > 0 Then
        Exit Function
        End If
        End If
    End If
End If
Next

Else    'Line symbols

For c1 = 0 To 4
If PYLN(c1) = ptest Then
    If piccase(c1) > 0 Then
    testtotal = testtotal + 1
    holdqnew(c1) = piccase(c1)
    ElseIf piccase(c1) < 0 Then
        If heldnow(c1) <= 0 Then testodds = testodds + 1
        holdqnew(c1) = piccase(c1)
    Else
        If pictest > 0 Then
        If spct = 0 Then
        Exit Function
        ElseIf heldnow(c1) <= 0 Then
        Exit Function
        End If
        End If
    End If
End If

Next
End If

If testodds = possodds And rltotal > 1 Then

If (pictest < 0 And testtotal > 2) Or (pictest > 0 And testtotal = rltotal) Then 'free spin or game

Winxcompare = True

For c1 = 0 To 4
If PYLN(c1) <> win2old(currline, c1) And heldnow(c1) = 0 Then subsused(c1) = True
Next

End If
End If

End Function
Public Sub zerogamspnvars()

For ct1 = 0 To 3
heldnow(ct1) = 0
For temp = 0 To 2
qtrakker(temp, ct1) = 0
For ct2 = 0 To 10
For ct3 = 0 To 6
holdq(temp, ct1, ct2, ct3) = 0
Next
Next
Next
Next
For temp = 0 To 2
subsline(temp) = False
For ct1 = 0 To 6
hq(temp, ct1) = 0
Next
Next
heldnow(4) = 0
currheld = 0
End Sub
Public Function clearq(zkept1or2 As Long, zfgct As Long, zspct As Long, zgamenotspin As Boolean)
'Called at the end of every freegamespin sequence only from the timer
'This function returns false if there are queued games or spins
'and shifts the counters accordingly
'else clears all counters, sets gamespinwait=false
clearq = False

'First clear CURRENT gamespin vars

If gamenotspin = True Then
qtrakker(currheld, kept1or2) = qtrakker(currheld, kept1or2) - 1
For ct1 = 0 To 4
holdq(currheld, kept1or2, 0, ct1) = 0
Next
Else
qtrakker(currheld, kept1or2 + 2) = qtrakker(currheld, kept1or2 + 2) - 1
For ct1 = 0 To 4
holdq(currheld, kept1or2 + 2, 0, ct1) = 0
Next
End If

fgct = 0
spct = 0
zfgct = 0
zspct = 0
subsline(currheld) = False

'check for outstanding qs

For currline = 0 To gt(153)
For ct = gamespintot To 1 Step -1   'according to specified priority

pct = gamespinkeep(ct - 1)

If gamespinsymbol(0) = pct Then  'free games
If gameclear(0, zkept1or2, zfgct, zspct, zgamenotspin) = True Then Exit Function
ElseIf gamespinsymbol(1) = pct Then
If gameclear(1, zkept1or2, zfgct, zspct, zgamenotspin) = True Then Exit Function
ElseIf gamespinsymbol(2) = pct Then  'free spins
If spinclear(2, zkept1or2, zfgct, zspct, zgamenotspin) = True Then Exit Function
ElseIf gamespinsymbol(3) = pct Then
If spinclear(3, zkept1or2, zfgct, zspct, zgamenotspin) = True Then Exit Function
End If
Next
Next

zerogamspnvars


clearq = True

End Function
Private Function gameclear(oneortwo As Long, zkept1or2 As Long, zfgct As Long, zspct As Long, zgamenotspin As Boolean)
If qtrakker(currline, oneortwo) > 0 Then
fgct = 1
gamenotspin = True
kept1or2 = oneortwo
For ct1 = 1 To qtrakker(currline, oneortwo)
For ct2 = 0 To 6
holdq(currline, oneortwo, ct1 - 1, ct2) = holdq(currline, oneortwo, ct1, ct2)
Next
Next
'The reels are not held. holdq is to provide the historical data initially stored in holdqforgame
For ct1 = 0 To 6
hq(currline, ct1) = holdq(currline, oneortwo, 0, ct1)
'cleanup top of pile
holdq(currline, oneortwo, qtrakker(currline, oneortwo), ct1) = 0
Next
For ct1 = 0 To 4
heldnow(ct1) = 0    'don't care
Next


currheld = currline
zfgct = fgct
zgamenotspin = True
zkept1or2 = kept1or2
gameclear = True
Else
gameclear = False
End If

End Function
Private Function spinclear(oneortwo As Long, zkept1or2 As Long, zfgct As Long, zspct As Long, zgamenotspin As Boolean)
Dim centrewin As Long, currmov As Long
If qtrakker(currline, oneortwo) > 0 Then
spct = 1
gamenotspin = False
kept1or2 = oneortwo - 2



For ct1 = 1 To qtrakker(currline, oneortwo)
For ct2 = 0 To 6
holdq(currline, oneortwo, ct1 - 1, ct2) = holdq(currline, oneortwo, ct1, ct2)
Next
Next
'Get next in queue
For ct1 = 0 To 6
hq(currline, ct1) = holdq(currline, oneortwo, 0, ct1)
holdq(currline, oneortwo, qtrakker(currline, oneortwo), ct1) = 0
Next

'Now initialize heldnow
For ct1 = 0 To 4

Select Case currline
Case 0
currmov = 0
Case 1
currmov = -1
Case 2
currmov = 1
End Select

heldnow(ct1) = 0

spinztemp(ct1) = hq(currline, ct1)
If spinztemp(ct1) > 0 Then
If sst(hq(currline, 5), 2) = 1 Then
If currline = 0 And wheelorder(ct1, spinztemp(ct1)) = hq(currline, 5) Or wheelorder(ct1, Advanz(spinztemp(ct1), -hq(currline, 6))) = hq(currline, 5) Or wheelorder(ct1, Advanz(spinztemp(ct1), -hq(currline, 6))) = hq(currline, 5) Then heldnow(ct1) = hq(currline, 5)
Else
centrewin = Advanz(spinztemp(ct1), -hq(currline, 6) - currmov)
If wheelorder(ct1, centrewin) = hq(currline, 5) Or substitute(wheelorder(ct1, centrewin), hq(currline, 5)) = True Then heldnow(ct1) = hq(currline, 5)
End If
End If
Next

currheld = currline
zspct = spct
zgamenotspin = False
zkept1or2 = kept1or2
spinclear = True
Else
spinclear = False
End If
End Function
Private Sub holdqstore(test1or2 As Long, wheretostore As Long)
Dim toholdin As Long

For ct1 = 0 To 4
If holdqnew(ct1) > 0 Then
toholdin = hq(currline, ct1)
If toholdin < 0 Then toholdin = -toholdin
Else
toholdin = -hq(currline, ct1)
If toholdin > 0 Then toholdin = -toholdin
End If

holdq(currline, test1or2, wheretostore, ct1) = toholdin
Next
holdq(currline, test1or2, wheretostore, 5) = gamespinsymbol(test1or2)
holdq(currline, test1or2, wheretostore, 6) = dirofspin
End Sub
Public Sub arrangespins(gendirchange As Boolean)
Dim winz As Long, currmov As Long

Select Case currheld
Case 0
currmov = 0
Case 1
currmov = -1
Case 2
currmov = 1
End Select


        For intreel = 0 To 4
                
                dirspin(intreel) = dirofspin
                'adjust restore position for dirofspin
                If hq(currheld, 6) <> dirofspin Then
                If hq(currheld, intreel) < 0 Then
                ct = -1
                Else
                ct = 1
                End If
                hq(currheld, intreel) = ct * Advanz(ct * hq(currheld, intreel), 2 * dirofspin)
                End If
            
                
        spinztemp(intreel) = hq(currheld, intreel) 'top line going down
        
        'For odd spin dups
        If spinztemp(intreel) >= 0 Then
        
        winz = wheelorder(intreel, Advanz(spinztemp(intreel), -dirofspin - currmov))
        
            If winz = hq(currheld, 5) Or (substitute(winz, hq(currheld, 5)) = True And reelcheck(winz, intreel + 1) = True) Then
            hreel(intreel) = True
            ElseIf sst(hq(currheld, 5), 2) = 1 Then
            'Do extra test for scatter
        
                If wheelorder(intreel, Advanz(spinztemp(intreel), dirofspin)) = hq(currheld, 5) Then
                hreel(intreel) = True
                ElseIf wheelorder(intreel, Advanz(spinztemp(intreel), -2 * dirofspin)) = hq(currheld, 5) Then
                'Advanz to row 3
                hreel(intreel) = True
                Else
                hreel(intreel) = False
                End If
        
            Else
            hreel(intreel) = False
            End If
        Else
        hreel(intreel) = False
        End If
        Next
        
        
         'Now figure out what goes on the held reels
        For intreel = 0 To 4
        If hreel(intreel) = True Then
        spinz(intreel) = hq(currheld, intreel)
        spinz(intreel) = Advanz(spinz(intreel), dirofspin * currmov)
        For ct = 0 To 3
        spinztemp(intreel) = hq(currheld, intreel) 'necessary for advanz
        'necessary for advanz - top down arrangement 1,2,3,0
            If dirofspin = 1 Then
            If gendirchange = False Then
                If ct = 0 Then
                ct1 = -3
                Else
                ct1 = 1 - ct
                End If
            Else
                If ct = 0 Then
                ct1 = 1
                Else
                ct1 = ct - 3
                End If
            End If
            Else    '0,3,2,1
            If gendirchange = False Then
                If ct = 0 Then
                ct1 = 3
                Else
                ct1 = ct - 1
                End If
            Else
                If ct = 0 Then
                ct1 = -1
                Else
                ct1 = 3 - ct
                End If
            End If
            End If
        Set Pokemach.M(picnum(intreel, ct)).Picture = Pokemach.Thumbslist(intreel).ListImages(Advanz(spinztemp(intreel), ct1)).Picture
        Next
        Else
        
        'Now restore current non held symbols
        hq(currheld, intreel) = ohqgs(currheld, intreel)
        End If
        
        Next
End Sub
Public Function Anychanges()

Anychanges = False

'Changes in wheelvec?
If gt(26) <> GTp(26) Then
GTp(26) = gt(26)
Anychanges = True
End If


For pct = 1 To thumbsize
For ct = 1 To 10
If sst(pct, ct) <> sstabgen(pct, ct) Then
sstabgen(pct, ct) = sst(pct, ct)
gt(27) = gt(27) + 1
Anychanges = True
End If
Next
Next


For ct = 0 To 200
Select Case ct
Case 0, 1, 2, 3, 6, 7, 8, 11, 13, 16, 17, 18, 20, 26 To 200


Case Else

If gt(ct) <> GTp(ct) Then
gt(28) = gt(28) + 1
GTp(ct) = gt(ct)
Anychanges = True
End If
End Select
Next

End Function
Public Sub sstgtsav()
'Do every new game
For ct = 1 To 200
GTp(ct) = gt(ct)
Next
For pct = 1 To 14
For ct = 0 To 10
sstabgen(pct, ct) = sst(pct, ct)
Next
Next

End Sub
