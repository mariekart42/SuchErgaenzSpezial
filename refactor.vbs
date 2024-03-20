
Function IsLetter(char As String) As Boolean
    ' Convert the character to its ASCII code
    Dim charCode As Integer
    charCode = Asc(UCase(char))

    ' Check if the character code is within the range of uppercase letters
    If charCode >= Asc("A") And charCode <= Asc("Z") Then
        IsLetter = True
    Else
        IsLetter = False
    End If
End Function

Function GetFindText(suchstring) As String
    Dim lowercase As String
    Dim uppercase As String
    Dim cutFirstLetter As String

    lowercase = LCase(left(suchstring, 1))
    uppercase = UCase(lowercase)
    cutFirstLetter = Mid(suchstring, 2)

    GetFindText = "[" & lowercase & "," & uppercase & "]" & cutFirstLetter
End Function

Function getLetterBefore(rangeObj As Range) As String
    Dim copyRange As Range
    Set copyRange = rangeObj.Duplicate
    Dim letterBefore As String

    ' Check if the range is at the beginning of the document
    If copyRange.Start = 0 Then
        getLetterBefore = "BOF"
        Set copyRange = Nothing
        Exit Function
    End If
    copyRange.MoveStart unit:=wdCharacter, Count:=-1
    letterBefore = copyRange.Characters(1).text
    Set copyRange = Nothing
    getLetterBefore = letterBefore
End Function

Function getLetterAfter(rangeObj As Range) As String
    Dim copyRange As Range
    Dim letterAfter As String

    Set copyRange = rangeObj.Duplicate

    copyRange.MoveEnd unit:=wdCharacter, Count:=1
    If copyRange.End = copyRange.Document.Content.End Then
        getLetterAfter = "EOF"
        Set copyRange = Nothing
        Exit Function
    End If
    copyRange.MoveEnd unit:=wdCharacter, Count:=-1
    copyRange.MoveEnd wdCharacter
    letterAfter = copyRange.Characters(copyRange.Characters.Count).text
    Set copyRange = Nothing
    getLetterAfter = letterAfter
End Function

Function getSecondLetterAfter(rangeObj As Range) As String
    ' Create a copy of the original range
    Dim copyRange As Range
    Set copyRange = rangeObj.Duplicate
    Dim letterAfter As String

    copyRange.MoveEnd unit:=wdCharacter, Count:=2
    ' Check if the range is at the beginning of the document
    If copyRange.End = copyRange.Document.Content.End Then
        getSecondLetterAfter = "EOF"
        Set copyRange = Nothing
        Exit Function
    End If

    ' Move the range one character backward
    copyRange.MoveEnd unit:=wdCharacter, Count:=-2
    copyRange.MoveEnd wdCharacter, 2
    letterAfter = copyRange.Characters(copyRange.Characters.Count).text
    Set copyRange = Nothing
    getSecondLetterAfter = letterAfter
End Function

Function FoundDuplicate(splitar() As String) As Boolean
    Dim i As Integer
    Dim m As Integer
    Dim n As Integer

    i = UBound(splitar)

    For m = 1 To i
        For n = 1 To i
            If Not (m = n) Then ' Avoid comparing an element with itself
                If InStr(splitar(m), splitar(n)) > 0 Then
                    Dim msg As String
                    MsgBox "The search term " & Chr(34) & splitar(n) & Chr(34) & " is reused in the search term " & Chr(34) & splitar(m) & Chr(34) & "!", vbOKOnly
                    FoundDuplicate = True
                    Exit Function
                ElseIf InStr(splitar(n), splitar(m)) > 0 Then
                    MsgBox "The search term " & Chr(34) & splitar(m) & Chr(34) & " is reused in the search term " & Chr(34) & splitar(n) & Chr(34) & "!", vbOKOnly
                    FoundDuplicate = True
                    Exit Function
                End If
            End If
        Next n
    Next m
    ' If no duplicates are found
    FoundDuplicate = False
End Function

Function GetBezugArray(myPath As String) As Variant
    Dim result(1 To 2) As Variant
    Open myPath For Input As #1
    Dim suchstring, strVariable As String
    Dim bezugArray() As String
    Dim splitar() As String
    ReDim bezugArray(0), splitar(0)
    Dim i, l As Integer

    i = 0

    Do While Not EOF(1)
        'Ãƒâ€žnderung 19.01.2017: Befehl "Line" hinzugefÃƒÂ¼t (Zeile samt Komma als String einlesen), damit mehrere kommagetrennte Bezugszeichen benutzt werden kÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¶nnen
        Line Input #1, strVariable
        i = i + 1: ReDim Preserve bezugArray(i)
        bezugArray(i) = strVariable
        'Prüfung, ob Trennzeichen vorhanden
        If InStr(strVariable, "@") = 0 Then
            result(1) = "ENDE"
            Dim lol As String
            lol = MsgBox("Trennzeichen (@) in Datei bezug.txt fehlt! Vorgang wird abgebrochen!", vbCritical, "TrennzeichenprÃƒÂ¼fung")
            Close #1
            GetBezugArray = result
            Exit Function
        End If

        suchstring = left(bezugArray(i), InStr(bezugArray(i), "@") - 1)
        l = l + 1: ReDim Preserve splitar(l)
        splitar(l) = suchstring
        Selection.HomeKey unit:=wdStory
        Selection.Find.ClearFormatting

        With Selection.Find
            .text = suchstring
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = True
            .MatchAllWordForms = False 'proplem with sentences if True!
            .MatchSoundsLike = False
            .MatchWildcards = False
        End With
        Selection.Find.Execute
        If Selection.Find.Found = False Then
            If MsgBox("Der Suchbegriff " & Chr(34) & suchstring & Chr(34) & " konnte nicht gefunden werden." & vbCrLf & vbCrLf & "Trotzdem fortfahren?", vbYesNo + vbCritical, "Suchen & ErgÃƒÆ’Ã‚Â¤nzen fehlgeschlagen!") = vbNo Then
                'fix later: collect not founds in array and display yesNo Box
                'Close #1
                'result(1) = "ENDE"
                'GetBezugArray = result
                'Exit Function
            End If
        End If
    Loop
    Close #1

    'If FoundDuplicate(splitar) Then
    '    MsgBox ("BZEE")
    '    result(1) = "DUPLICATE"
    'Else
    If strVariable <> "" Then
        result(1) = "OK"
        result(2) = bezugArray
    Else
        If MsgBox("Datei bezug.txt ist leer. Fortfahren mit manueller Eingabe?", vbYesNo, "Inhaltspüfung") = vbYes Then
            result(1) = "DOCUMENT EMPTY"
        Else
            result(1) = "ENDE"
        End If
    End If
    GetBezugArray = result
End Function
Function getEnvironmentPath() As Variant
    Dim result(1 To 2) As Variant
    Dim path As String
    Dim debugState As Boolean
    debugState = True

    If debugState = False Then
        Dim user As String
        Dim profil As String

        user = Environ("Username")
        profil = Environ("AppData")

        'TXT-Datei auf SH-User-Desktop
        If InStr(profil, "\sh\") <> 0 Then
            path = "\\brefile01\profile$\" & user & "\sh\Desktop\" & "bezug.txt"
        Else
            path = "\\brefile11\userhomes$\" & user & "\Desktop\" & "bezug.txt"
        End If
        If Dir(path) <> "" Then
            If MsgBox("Datei bezug.txt nicht vorhanden. Fortfahren mit manueller Eingabe?", vbYesNo, "Dateiprüfung") = vbYes Then
                result(1) = "SUCHEINGABE"
            Else
                result(1) = "ENDE"
            End If
        Else
            If MsgBox("Datei bezug.txt vorhanden. Fortfahren?", vbYesNo, "Dateiprüfung") = vbYes Then
                result(1) = "OK"
                result(2) = path
            Else
                result(1) = "ENDE"
            End If
        End If
    Else 'delete this section later, only for me testing
        'my thing to make it work for citrix:
        path = "\\brefile11.esp.dom\citrixprofiles$\msg\Desktop\bezug.txt"
        If Dir(path) <> "" Then
            If MsgBox("Datei bezug.txt vorhanden. Fortfahren?", vbYesNo, "Dateiprüfung") = vbYes Then
                result(1) = "OK"
                result(2) = path
            Else
                result(1) = "ENDE"
            End If
        Else
            If MsgBox("Datei bezug.txt nicht vorhanden. Fortfahren mit manueller Eingabe?", vbYesNo, "Dateiprüfung") = vbYes Then
                result(1) = "SUCHEINGABE"
            Else
                result(1) = "ENDE"
            End If
        End If
    End If
    getEnvironmentPath = result
End Function
Function InvalidDocument() As Boolean
    Dim response As String
    Dim effC As Variant
    Dim effS, effI As Integer

    'no documents open
    If Documents.Count < 1 Then
        response = MsgBox("Es ist kein Dokument geÃƒÂ¶ffnet.", vbOKOnly + vbCritical, "Suchen & Ergänzen fehlgeschlagen!")
        InvalidDocument = True
        Exit Function
    End If

    'Anzahl der Zeichen, Shapes oder InlineShapes feststellen
    effC = ActiveDocument.BuiltInDocumentProperties(wdPropertyCharsWSpaces)
    effS = ActiveDocument.Shapes.Count
    effI = ActiveDocument.InlineShapes.Count

    'keine Zeichen, Shapes oder InlineShapes
    If effC < 1 And effS < 1 And effI < 1 Then
        response = MsgBox("Suchen & Ergänzen im leeren Dokument nicht möglich.", vbOKOnly + vbCritical, "Suchen & Ergänzen fehlgeschlagen!")
        InvalidDocument = True
        Exit Function
    End If
    InvalidDocument = False
End Function

Sub SetTrackingSettings()
    Dim o, p As Integer
    'Ãƒâ€žnderung 19.01.2017: prÃƒÂ¼fen, ob nicht angenommene Ãƒâ€žnderungen eines anderen Benutzers vorhanden sind
    o = ActiveDocument.Revisions.Count
    For p = 1 To o
        If ActiveDocument.Revisions.Count <> 0 And ActiveDocument.Revisions(p).Author <> Application.UserName Then
            MsgBox "ACHTUNG:" & vbCrLf & vbCrLf & "Nicht angenommene ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾nderungen eines anderen Benutzers (" & ActiveDocument.Revisions(p).Author & ") vorhanden - nachtrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¤gliche ErgÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¤nzungen beeinflussen diese ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾nderungen!" & vbCrLf & vbCrLf & "Bitte anschlieÃƒÆ’Ã†â€™Ãƒâ€¦Ã‚Â¸end prÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¼fen!", vbOKOnly + vbExclamation, "Suchen & ErgÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¤nzen"
            Exit Sub
        End If
    Next

    'wenn 'ÃƒÆ’Ã¢â‚¬Å¾nderungen verfolgen' deaktiviert ist -> aktivieren
    If ActiveDocument.TrackRevisions = False Then
        ActiveDocument.TrackRevisions = True
    End If

    'Sprechblasen-Einstellung auf balloon umstellen
    ActiveWindow.View.MarkupMode = wdBalloonRevisions

    'ErklÃƒÂ¤rung
    Dim response As String
    response = MsgBox("Die Datei muss auf dem Desktop mit dem Dateinamen bezug.txt angelegt werden und muss folgende zeilenweise Syntax aufweisen:" & vbCrLf & vbCrLf & _
    "Suchbegriff1@Bezugsbezeichnung1" & vbCrLf & "Suchbegriff2@Bezugsbezeichnung2" & vbCrLf & "..." & vbCrLf & vbCrLf & _
    "Es wird nur nach ganzen WÃƒÂ¶rtern gesucht." & vbCrLf & "Die Bezugsbezeichnung (nur Zahl, ohne Suchbegriff) wird beim EinfÃƒÂ¼gen automatisch in Klammern gesetzt.", vbInformation, "ErklÃƒÂ¤rung")
End Sub

Sub SortArrayByStringLength(inputArray As Variant)
    Dim temp As String
    Dim i As Integer, j As Integer

    ' Perform sorting using bubble sort algorithm
    For i = LBound(inputArray) To UBound(inputArray)
        ' Skip empty lines or lines without "@" symbol
        If Len(inputArray(i)) > 0 And InStr(inputArray(i), "@") > 0 Then
            For j = i + 1 To UBound(inputArray)
                ' Skip empty lines or lines without "@" symbol
                If Len(inputArray(j)) > 0 And InStr(inputArray(j), "@") > 0 Then
                    ' Extract the string before the "@" symbol for comparison
                    Dim string1 As String
                    Dim string2 As String
                    string1 = left(inputArray(i), InStr(inputArray(i), "@") - 1)
                    string2 = left(inputArray(j), InStr(inputArray(j), "@") - 1)

                    ' Compare lengths and swap if necessary
                    If Len(string1) < Len(string2) Then
                        temp = inputArray(i)
                        inputArray(i) = inputArray(j)
                        inputArray(j) = temp
                    End If
                End If
            Next j
        End If
    Next i
End Sub
Function CheckForRangeConflict(newStart As Integer, newEnd As Integer, rangesArray As Variant) As Boolean
    ' Iterate through the existing ranges array
    'MsgBox ("lBound: " & LBound(rangesArray, 2))
    'MsgBox ("UBound: " & UBound(rangesArray, 2))
    Dim i As Integer
    For i = LBound(rangesArray, 2) To UBound(rangesArray, 2)
        ' Extract start and end positions of the existing range
        Dim existingStart As Integer
        Dim existingEnd As Integer
        existingStart = rangesArray(1, i)
        existingEnd = rangesArray(2, i)

        ' Check for overlap
        If newStart <= existingEnd And newEnd >= existingStart Then
            ' Conflict detected
            CheckForRangeConflict = True
            Exit Function
        End If
    Next i

    ' No conflicts found
    CheckForRangeConflict = False
End Function
Function GetRangesArray(bezugArray() As String) As Variant()
    Dim i, k, numCol As Integer
    Dim suchstring, ergaenzstring As String
    Dim rangesArray() As Variant

    numCol = 0
    'ReDim rangesArray(1 To 1, 1 To 3)
    i = (UBound(bezugArray) - LBound(bezugArray) + 1) - 1
    For k = 1 To i
        suchstring = left(bezugArray(k), InStr(bezugArray(k), "@") - 1)
        ergaenzstring = " (" & right(bezugArray(k), Len(bezugArray(k)) - InStr(bezugArray(k), "@")) & ")"

        'MsgBox ("HERE: suchstring: " & suchstring & "| number: " & ergaenzstring)

        Dim selectionRange As Range
        Set selectionRange = ActiveDocument.Range
        Dim text As String
        text = GetFindText(suchstring)
        Do While selectionRange.Find.Execute(FindText:=text, MatchAllWordForms:=False, MatchSoundsLike:=False, MatchWildcards:=True, Forward:=True) = True
            Dim letterAfter As String
            Dim secondLetterAfter As String
            Dim letterBefore As String

            letterAfter = getLetterAfter(selectionRange)
            secondLetterAfter = getSecondLetterAfter(selectionRange)
            letterBefore = getLetterBefore(selectionRange)

            If (Not IsLetter(letterBefore) And Not IsNumeric(letterBefore)) Or letterBefore = "BOF" Then
                If letterAfter = "s" Then
                    'check if next character is a letter or number, if yes dont consider as suchstring
                    If Not (IsLetter(secondLetterAfter) And Not IsNumeric(secondLetterAfter)) Or secondLetterAfter = "EOF" Then
                        'move selection to one character before word ends ('s' as extra letter)
                        'MsgBox ("FOUND RANGE: word: " & text & " Range.start: " & selectionRange.Start & " End: " & selectionRange.End + 1)

                        'MsgBox ("numCol: " & numCol)

                        Dim lol() As Variant
                        lol = rangesArray
                        If numCol > 0 Then
                            MsgBox ("lBound: " & LBound(rangesArray, 2))
                            If CheckForRangeConflict(selectionRange.Start, selectionRange.End + 1, rangesArray) = False Then
                                ' Resize the array to accommodate one more row
                                ReDim Preserve rangesArray(3, numCol + 1)
                                ' Add the new item to the last row
                                rangesArray(1, numCol) = selectionRange.Start
                                rangesArray(2, numCol) = selectionRange.End + 1
                                rangesArray(3, numCol) = ergaenzstring
                                numCol = numCol + 1
                                selectionRange.MoveEnd wdCharacter
                                'selectionRange.InsertAfter ergaenzstring
                            Else


                            End If
                        Else
                        ' Add the new item to the last row
                            rangesArray(1, numCol) = selectionRange.Start
                            rangesArray(2, numCol) = selectionRange.End + 1
                            rangesArray(3, numCol) = ergaenzstring
                            numCol = numCol + 1
                            selectionRange.MoveEnd wdCharacter
                        End If
                        'MsgBox ("lBound: " & LBound(rangesArray, 2))

                    End If
                ElseIf (Not IsLetter(letterAfter) And Not IsNumeric(letterAfter)) Or letterAfter = "EOF" Then

                    Dim lol2() As Variant
                        lol2 = rangesArray
                        If numCol > 0 Then
                            'MsgBox ("lBound: " & LBound(rangesArray, 2))
                            If CheckForRangeConflict(selectionRange.Start, selectionRange.End, rangesArray) = False Then
                                ' Resize the array to accommodate one more row
                                ReDim Preserve rangesArray(3, numCol + 1)
                                ' Add the new item to the last row
                                rangesArray(1, numCol) = selectionRange.Start
                                rangesArray(2, numCol) = selectionRange.End
                                rangesArray(3, numCol) = ergaenzstring
                                numCol = numCol + 1
                                selectionRange.MoveEnd wdCharacter
                                'selectionRange.InsertAfter ergaenzstring
                            Else
                                MsgBox ("FOUND OVERLAP")
                            End If
                        Else
                        ' Resize the array to accommodate one more row
                            ReDim Preserve rangesArray(3, numCol + 1)
                            ' Add the new item to the last row
                            rangesArray(1, numCol) = selectionRange.Start
                            rangesArray(2, numCol) = selectionRange.End
                            rangesArray(3, numCol) = ergaenzstring
                            numCol = numCol + 1
                            selectionRange.MoveEnd wdCharacter
                        End If




                    'move selection to two character before word ends
                    'ReDim Preserve rangesArray(3, numCol + 1)
                    'rangesArray(1, numCol) = selectionRange.Start
                    'rangesArray(2, numCol) = selectionRange.End
                    'rangesArray(3, numCol) = ergaenzstring
                    'numCol = numCol + 1
                End If
            End If
            selectionRange.Collapse wdCollapseEnd
        Loop
    Next
    GetRangesArray = rangesArray
End Function

Sub InsertNumbers(rangesArray() As Variant)
    Dim i As Integer
    For i = LBound(rangesArray, 2) To UBound(rangesArray, 2)
        MsgBox (rangesArray(1, i) & "  " & rangesArray(2, i) & "  " & rangesArray(3, i) & "  ")
    Next i

End Sub

Sub SortArrayBySecondColumnDescending(ByRef rangesArray() As Variant)
    Dim numRows As Long
    numRows = UBound(rangesArray, 2)

    Dim i, j, k As Long
    Dim temp As Variant

    ' Perform sorting using bubble sort algorithm
    For i = 1 To numRows
        For j = 1 To numRows - 1 ' Adjusted loop bound to avoid going out of bounds
            If rangesArray(2, j) < rangesArray(2, j + 1) Then
                ' Swap elements if needed
                For k = 1 To UBound(rangesArray, 1)
                    temp = rangesArray(k, j)
                    rangesArray(k, j) = rangesArray(k, j + 1)
                    rangesArray(k, j + 1) = temp
                Next k
            End If
        Next j
    Next i
End Sub
Function NewSuchErsetz(bezugArray() As String)
    MsgBox ("coming soon")

    '- sort array of bezugArray (biggest first)
    SortArrayByStringLength bezugArray


    '- iterate through array & store found texts ranges in 3 times X array, third position is bezugszeichen (range.start | range.end | number to insert)
    Dim rangesArray() As Variant
    rangesArray = GetRangesArray(bezugArray)

    MsgBox ("lbound outside: " & LBound(rangesArray))
    SortArrayBySecondColumnDescending rangesArray


    ' Move the insertion point to the desired position
    'docContent.MoveStart unit:=wdCharacter, Count:=insertPosition

    ' Insert the text at the specified position
    'docContent.InsertAfter textToInsert

    Dim insertPosition As Integer
    Dim textToInsert As String



    ' Get a reference to the document's content
    Dim docContent As Range
    Set docContent = ActiveDocument.Content




    Dim q As Integer
    For q = LBound(rangesArray, 2) To UBound(rangesArray, 2)
        MsgBox ("Range.Start: " & rangesArray(1, q) & " Range.End: " & rangesArray(2, q) & " Nbr: " & rangesArray(3, q))

        insertPosition = rangesArray(2, q)  ' For example, insert at position 10
        textToInsert = " " & rangesArray(3, q)

         ' Set the insertion point to the desired position
        'docContent.End = rangesArray(2, q)  ' Move the start of the range to the desired position
        docContent.SetRange Start:=rangesArray(1, q), End:=rangesArray(2, q)
        ' Insert the text at the specified position
        'docContent.text = rangesArray(3, q)

        docContent.InsertAfter rangesArray(3, q)
    Next q


    'InsertNumbers rangesArray


'- for every following string, check if found string is in the 3X array if yes, skip it
'- iterate backwards through text and insert specific numbers after range end


End Function

Sub SuchErgaenzSpezial()

'
'Makro vom 30.09.2016 von Jacek Manka
'bearbeitet 10.03.2024 von Marie Mensing
'

'Variablen setzen
Dim markup, trackrev
Dim Titel5, Mldg5, Stil5, Antwort5
Dim suchstring, ergaenzstring, Pfad, ar() As String
Dim i, k, l, m, n, o, p As Integer

ReDim ar(0)
i = 0

    Mldg5 = "Suchen & Ergänzen durch Abbruch beendet."
    Stil5 = vbInformation
    Titel5 = "Suchen & Ergänzen abgebrochen!"

    If InvalidDocument Then
        GoTo Ende
    End If

    'Nachverfolgungseinstellungen sichern
    trackrev = ActiveDocument.TrackRevisions
    markup = ActiveWindow.View.MarkupMode
    SetTrackingSettings

    Dim vals As Variant
    vals = getEnvironmentPath()
    If vals(1) = "ENDE" Then
        GoTo Ende
    ElseIf vals(1) = "SUCHEINGABE" Then
        GoTo Sucheingabe
    Else
        Dim path As String
        path = vals(2)
        Dim bezugArray() As String
        Dim values As Variant
        values = GetBezugArray(path)
        Select Case values(1)
            Case "OK"
                bezugArray = values(2)
                NewSuchErsetz bezugArray
                GoTo Ende
                'GoTo SuchErsetz
            Case "ENDE"
                GoTo Ende
            Case "DUPLICATE"
                GoTo Ende
            Case "DOCUMENT EMPTY"
                GoTo Sucheingabe
        End Select
    End If
    GoTo Ende

Sucheingabe:
    'Suchstring
    suchstring = InputBox("Bitte geben Sie den " & (i + 1) & ". Suchbegriff ein:", "Eingabe des Suchbegriffs")

    'Suchstring Cancel?
    If StrPtr(suchstring) = 0 Then
        Antwort5 = MsgBox(Mldg5, Stil5, Titel5)
        GoTo Ende:
    Else
        'OK und kein Suchstring?
        If Len(suchstring) = 0 Then
            If MsgBox("Suchen & Ergänzen kann nicht stattfinden, weil kein Suchbegriff / keine Ergänzung eingegeben wurde.", vbRetryCancel, "Suchen & ErgÃƒÂ¤nzen fehlgeschlagen!") = vbRetry Then
                GoTo Sucheingabe:
            Else
                Antwort5 = MsgBox(Mldg5, Stil5, Titel5)
                GoTo Ende:
            End If
        End If
    End If

    'Suchen
    Selection.HomeKey unit:=wdStory
    Selection.Find.ClearFormatting
    With Selection.Find
        .text = suchstring
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = True
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = False
    End With
    Selection.Find.Execute

    If Selection.Find.Found = False Then
        Dim response As String
        response = MsgBox("Suchen & Ergänzen kann nicht stattfinden, weil der Suchbegriff " & Chr(34) & suchstring & Chr(34) & " nicht gefunden werden konnte.", vbOKOnly + vbCritical, "Suchen & ErgÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¤nzen fehlgeschlagen!")
        GoTo Ende
    End If

Ergaenzeingabe:
    'Ergaenzstring
    ergaenzstring = InputBox("Bitte geben Sie die " & (i + 1) & ". Bezugsbezeichnung ein, um die Sie den " & (i + 1) & ". Suchbegriff ergÃƒÂ¤nzen mÃƒÂ¶chten:", "Eingabe der Bezugsbezeichnung")

    'Ergaenzstring Cancel?
    If StrPtr(ergaenzstring) = 0 Then
        Antwort5 = MsgBox(Mldg5, Stil5, Titel5)
        GoTo Ende:
    Else

        'OK und kein Suchstring?
        If Len(ergaenzstring) = 0 Then
            If MsgBox("Suchen & Ergänzen kann nicht stattfinden, weil kein Suchbegriff / keine ErgÃƒÂ¤nzung eingegeben wurde.", vbRetryCancel, "Suchen & ErgÃƒÂ¤nzen fehlgeschlagen!") = vbRetry Then
                GoTo Ergaenzeingabe:
            Else
                Antwort5 = MsgBox(Mldg5, Stil5, Titel5)
                GoTo Ende:
            End If
        End If
    End If

    'Array aus manueller Eingabe erstellen
    i = i + 1: ReDim Preserve ar(i)
    ar(i) = suchstring & "@" & ergaenzstring

    If MsgBox("MÃƒÂ¶chten Sie weitere Bezugsbezeichnungen einfÃƒÂ¼gen?", vbYesNo, "Wiederholung") = vbYes Then
        GoTo Sucheingabe
    End If

SuchErsetz:
    i = (UBound(bezugArray) - LBound(bezugArray) + 1) - 1
    For k = 1 To i
        suchstring = left(bezugArray(k), InStr(bezugArray(k), "@") - 1)
        ergaenzstring = " (" & right(bezugArray(k), Len(bezugArray(k)) - InStr(bezugArray(k), "@")) & ")"
        Dim selectionRange As Range
        Set selectionRange = ActiveDocument.Range
        Dim text As String
        text = GetFindText(suchstring)
        Do While selectionRange.Find.Execute(FindText:=text, MatchAllWordForms:=False, MatchSoundsLike:=False, MatchWildcards:=True, Forward:=True) = True
            Dim letterAfter As String
            Dim secondLetterAfter As String
            Dim letterBefore As String

            letterAfter = getLetterAfter(selectionRange)
            secondLetterAfter = getSecondLetterAfter(selectionRange)
            letterBefore = getLetterBefore(selectionRange)

            If (Not IsLetter(letterBefore) And Not IsNumeric(letterBefore)) Or letterBefore = "BOF" Then
                If letterAfter = "s" Then
                    'check if next character is a letter or number, if yes dont consider as suchstring
                    If Not (IsLetter(secondLetterAfter) And Not IsNumeric(secondLetterAfter)) Or secondLetterAfter = "EOF" Then
                        'move selection to one character before word ends ('s' as extra letter)
                        selectionRange.MoveEnd wdCharacter
                        selectionRange.InsertAfter ergaenzstring
                    End If
                ElseIf (Not IsLetter(letterAfter) And Not IsNumeric(letterAfter)) Or letterAfter = "EOF" Then
                    'move selection to two character before word ends
                    selectionRange.InsertAfter ergaenzstring
                End If
            End If
            selectionRange.Collapse wdCollapseEnd
        Loop
    Next
Ende:
    MsgBox ("ENDE")
    'Nachverfolgungseinstellungen wiederherstellen
     ActiveWindow.View.MarkupMode = markup
     ActiveDocument.TrackRevisions = trackrev

    Selection.HomeKey unit:=wdStory

    'Suchparameter zurÃƒÂ¼cksetzen
    With Selection.Find
       .ClearFormatting
       .Replacement.ClearFormatting
       .text = ""
       .Replacement.text = ""
       .Forward = True
       .Wrap = wdFindContinue
       .Format = False
       .MatchCase = False
       .MatchWholeWord = False
       .MatchWildcards = False
       .MatchSoundsLike = False
       .MatchAllWordForms = False
    End With
End Sub