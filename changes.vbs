Function IsLetter(char As String) As Boolean
    Dim charCode As Integer
    charCode = Asc(UCase(char))
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

    lowercase = LCase(Left(suchstring, 1))
    uppercase = UCase(lowercase)
    cutFirstLetter = Mid(suchstring, 2)

    GetFindText = "[" & lowercase & "," & uppercase & "]" & cutFirstLetter
End Function

Function getLetterBefore(rangeObj As range) As String
    Dim copyRange As range
    Set copyRange = rangeObj.Duplicate
    Dim letterBefore As String

    ' Check if the range is at the beginning of the document
    If copyRange.start = 0 Then
        getLetterBefore = "BOF"
        Set copyRange = Nothing
        Exit Function
    End If
    copyRange.MoveStart unit:=wdCharacter, Count:=-1
    letterBefore = copyRange.Characters(1).text
    Set copyRange = Nothing
    getLetterBefore = letterBefore
End Function

Function getLetterAfter(rangeObj As range) As String
    Dim copyRange As range
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

Function getSecondLetterAfter(rangeObj As range) As String
    ' Create a copy of the original range
    Dim copyRange As range
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
        'ÃƒÆ’Ã¢â‚¬Å¾nderung 19.01.2017: Befehl "Line" hinzugefÃƒÆ’Ã‚Â¼t (Zeile samt Komma als String einlesen), damit mehrere kommagetrennte Bezugszeichen benutzt werden kÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¶nnen
        Line Input #1, strVariable
        i = i + 1: ReDim Preserve bezugArray(i)
        bezugArray(i) = strVariable
        'PrÃ¼fung, ob Trennzeichen vorhanden
        If InStr(strVariable, "@") = 0 Then
            result(1) = "ENDE"
            Dim lol As String
            lol = MsgBox("Trennzeichen (@) in Datei bezug.txt fehlt! Vorgang wird abgebrochen!", vbCritical, "TrennzeichenprÃƒÆ’Ã‚Â¼fung")
            Close #1
            GetBezugArray = result
            Exit Function
        End If

        suchstring = Left(bezugArray(i), InStr(bezugArray(i), "@") - 1)
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
            If MsgBox("Der Suchbegriff " & Chr(34) & suchstring & Chr(34) & " konnte nicht gefunden werden." & vbCrLf & vbCrLf & "Trotzdem fortfahren?", vbYesNo + vbCritical, "Suchen & ErgÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¤nzen fehlgeschlagen!") = vbNo Then
                'fix later: collect not founds in array and display yesNo Box
                'Close #1
                'result(1) = "ENDE"
                'GetBezugArray = result
                'Exit Function
            End If
        End If
    Loop
    Close #1

    If strVariable <> "" Then
        result(1) = "OK"
        result(2) = bezugArray
    Else
        If MsgBox("Datei bezug.txt ist leer. Fortfahren mit manueller Eingabe?", vbYesNo, "InhaltspÃ¼fung") = vbYes Then
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
    Dim user As String
    Dim profil As String

    Dim debugState As Boolean
    debugState = True

    If debugState = False Then


        user = Environ("Username")
        profil = Environ("AppData")

        'TXT-Datei auf SH-User-Desktop
        If InStr(profil, "\sh\") <> 0 Then
            path = "\\brefile01\profile$\" & user & "\sh\Desktop\" & "bezug.txt"
        Else
            path = "\\brefile11\userhomes$\" & user & "\Desktop\" & "bezug.txt"
        End If
        If Dir(path) <> "" Then
            If MsgBox("Datei bezug.txt nicht vorhanden. Fortfahren mit manueller Eingabe?", vbYesNo, "DateiprÃ¼fung") = vbYes Then
                result(1) = "SUCHEINGABE"
            Else
                result(1) = "ENDE"
            End If
        Else
            If MsgBox("Datei bezug.txt vorhanden. Fortfahren?", vbYesNo, "DateiprÃ¼fung") = vbYes Then
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
            If MsgBox("Datei bezug.txt vorhanden. Fortfahren?", vbYesNo, "DateiprÃ¼fung") = vbYes Then
                result(1) = "OK"
                result(2) = path
            Else
                result(1) = "ENDE"
            End If
        Else
            If MsgBox("Datei bezug.txt nicht vorhanden. Fortfahren mit manueller Eingabe?", vbYesNo, "DateiprÃ¼fung") = vbYes Then
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
        response = MsgBox("Es ist kein Dokument geÃƒÆ’Ã‚Â¶ffnet.", vbOKOnly + vbCritical, "Suchen & ErgÃ¤nzen fehlgeschlagen!")
        InvalidDocument = True
        Exit Function
    End If

    'Anzahl der Zeichen, Shapes oder InlineShapes feststellen
    effC = ActiveDocument.BuiltInDocumentProperties(wdPropertyCharsWSpaces)
    effS = ActiveDocument.Shapes.Count
    effI = ActiveDocument.InlineShapes.Count

    'keine Zeichen, Shapes oder InlineShapes
    If effC < 1 And effS < 1 And effI < 1 Then
        response = MsgBox("Suchen & ErgÃ¤nzen im leeren Dokument nicht mÃ¶glich.", vbOKOnly + vbCritical, "Suchen & ErgÃ¤nzen fehlgeschlagen!")
        InvalidDocument = True
        Exit Function
    End If
    InvalidDocument = False
End Function
Sub SetTrackingSettings()
    Dim o, p As Integer
    'Änderung 19.01.2017: prüfen, ob nicht angenommene Änderungenen eines anderen Benutzers vorhanden sind
    o = ActiveDocument.Revisions.Count
    For p = 1 To o
        If ActiveDocument.Revisions.Count <> 0 And ActiveDocument.Revisions(p).Author <> Application.UserName Then
            MsgBox "ACHTUNG:" & vbCrLf & vbCrLf & "Nicht angenommene Änderungenen eines anderen Benutzers (" & ActiveDocument.Revisions(p).Author & ") vorhanden - nachträgliche ErgÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¤nzungen beeinflussen diese ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¾nderungen!" & vbCrLf & vbCrLf & "Bitte anschlieÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¸end prÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¼fen!", vbOKOnly + vbExclamation, "Suchen & ErgÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¤nzen"
            Exit Sub
        End If
    Next

    'wenn 'Änderungen verfolgen' deaktiviert ist -> aktivieren
    If ActiveDocument.TrackRevisions = False Then
        ActiveDocument.TrackRevisions = True
    End If

    'Sprechblasen-Einstellung auf balloon umstellen
    ActiveWindow.View.MarkupMode = wdBalloonRevisions

    'Erklärung
    Dim response As String
    response = MsgBox("Die Datei muss auf dem Desktop mit dem Dateinamen bezug.txt angelegt werden und muss folgende zeilenweise Syntax aufweisen:" & vbCrLf & vbCrLf & _
    "Suchbegriff1@Bezugsbezeichnung1" & vbCrLf & "Suchbegriff2@Bezugsbezeichnung2" & vbCrLf & "..." & vbCrLf & vbCrLf & _
    "Es wird nur nach ganzen Wörtern gesucht." & vbCrLf & "Die Bezugsbezeichnung (nur Zahl, ohne Suchbegriff) wird beim Einfügen automatisch in Klammern gesetzt.", vbInformation, "ErklÃƒÆ’Ã‚Â¤rung")
End Sub

Sub SortArrayByStringLength(inputArray As Variant)
    Dim temp As String
    Dim i As Integer, j As Integer
    Dim string1 As String
    Dim string2 As String

    ' Perform sorting using bubble sort algorithm
    For i = LBound(inputArray) To UBound(inputArray)
        If Len(inputArray(i)) > 0 And InStr(inputArray(i), "@") > 0 Then
            For j = i + 1 To UBound(inputArray)
                If Len(inputArray(j)) > 0 And InStr(inputArray(j), "@") > 0 Then
                    string1 = Left(inputArray(i), InStr(inputArray(i), "@") - 1)
                    string2 = Left(inputArray(j), InStr(inputArray(j), "@") - 1)
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
Function FoundRangeConflict(newStart As Integer, newEnd As Integer, rangesArray As Variant, numCol As Integer) As Boolean
    Dim i As Integer
    Dim existingStart As Integer
    Dim existingEnd As Integer

    If numCol = 0 Then
        FoundRangeConflict = False
        Exit Function
    End If
    For i = LBound(rangesArray, 2) To UBound(rangesArray, 2)
        existingStart = rangesArray(0, i)
        existingEnd = rangesArray(1, i)
        ' Check for conflicting overlap
        If newStart <= existingEnd And newEnd >= existingStart Then
            FoundRangeConflict = True
            Exit Function
        End If
    Next i
    FoundRangeConflict = False
End Function
Function GetRangesArray(bezugArray() As String) As Variant()
    Dim i, k, numCol As Integer
    Dim suchstring, ergaenzstring, text As String
    Dim rangesArray() As Variant
    Dim letterAfter As String
    Dim secondLetterAfter As String
    Dim letterBefore As String
    Dim selectionRange As range

    numCol = 0
    i = (UBound(bezugArray) - LBound(bezugArray) + 1) - 1
    For k = 1 To i
        suchstring = Left(bezugArray(k), InStr(bezugArray(k), "@") - 1)
        ergaenzstring = " (" & Right(bezugArray(k), Len(bezugArray(k)) - InStr(bezugArray(k), "@")) & ")"
        Set selectionRange = ActiveDocument.range
        text = GetFindText(suchstring)
        Do While selectionRange.Find.Execute(FindText:=text, MatchAllWordForms:=False, MatchSoundsLike:=False, MatchWildcards:=True, Forward:=True) = True
            letterAfter = getLetterAfter(selectionRange)
            secondLetterAfter = getSecondLetterAfter(selectionRange)
            letterBefore = getLetterBefore(selectionRange)

            If (Not IsLetter(letterBefore) And Not IsNumeric(letterBefore)) Or letterBefore = "BOF" Then
                If letterAfter = "s" Then
                    'check if next character is a letter or number, if yes dont consider as suchstring
                    If Not (IsLetter(secondLetterAfter) And Not IsNumeric(secondLetterAfter)) Or secondLetterAfter = "EOF" Then
                        If FoundRangeConflict(selectionRange.start, selectionRange.End, rangesArray, numCol) = False Then
                            ReDim Preserve rangesArray(2, numCol)
                            rangesArray(0, numCol) = selectionRange.start
                            rangesArray(1, numCol) = selectionRange.End + 1
                            rangesArray(2, numCol) = ergaenzstring
                            numCol = numCol + 1
                            selectionRange.MoveEnd wdCharacter
                        End If
                    End If
                ElseIf (Not IsLetter(letterAfter) And Not IsNumeric(letterAfter)) Or letterAfter = "EOF" Then
                    If FoundRangeConflict(selectionRange.start, selectionRange.End, rangesArray, numCol) = False Then
                        ReDim Preserve rangesArray(2, numCol)
                        rangesArray(0, numCol) = selectionRange.start
                        rangesArray(1, numCol) = selectionRange.End
                        rangesArray(2, numCol) = ergaenzstring
                        numCol = numCol + 1
                        selectionRange.MoveEnd wdCharacter
                    End If
                End If
            End If
            selectionRange.Collapse wdCollapseEnd
        Loop
    Next
    GetRangesArray = rangesArray
End Function
Sub InsertNumbers(rangesArray() As Variant)
    Dim docContent As range
    Dim k As Integer

    Set docContent = ActiveDocument.Content
    For k = LBound(rangesArray, 2) To UBound(rangesArray, 2)
        docContent.SetRange start:=rangesArray(0, k), End:=rangesArray(1, k)
        docContent.InsertAfter rangesArray(2, k)
    Next k
End Sub
Sub SortArrayBySecondColumnDescending(ByRef rangesArray() As Variant)
    Dim numRows As Long
    Dim numCols As Long
    numRows = UBound(rangesArray, 2)
    numCols = UBound(rangesArray, 1)

    Dim i, j, k As Long
    Dim temp As Variant

    For i = 0 To numRows
        For j = 0 To numRows - 1
            If rangesArray(1, j) < rangesArray(1, j + 1) Then
                For k = 0 To numCols
                    temp = rangesArray(k, j)
                    rangesArray(k, j) = rangesArray(k, j + 1)
                    rangesArray(k, j + 1) = temp
                Next k
            End If
        Next j
    Next i
End Sub
Function NewSuchErsetz(bezugArray() As String)
    Dim rangesArray() As Variant

    SortArrayByStringLength bezugArray
    rangesArray = GetRangesArray(bezugArray)
    SortArrayBySecondColumnDescending rangesArray
    InsertNumbers rangesArray
End Function

Sub SuchErgaenzSpezial()

'
'Makro vom 30.09.2016 von Jacek Manka
'bearbeitet 21.03.2024 von Marie Mensing
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
        GoTo ende
    End If

    'Nachverfolgungseinstellungen sichern
    trackrev = ActiveDocument.TrackRevisions
    markup = ActiveWindow.View.MarkupMode
    SetTrackingSettings

    Dim vals As Variant
    vals = getEnvironmentPath()
    If vals(1) = "ENDE" Then
        GoTo ende
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
                GoTo ende
            Case "ENDE"
                GoTo ende
            Case "DUPLICATE"
                GoTo ende
            Case "DOCUMENT EMPTY"
                GoTo Sucheingabe
        End Select
    End If
    GoTo ende

Sucheingabe:
    suchstring = InputBox("Bitte geben Sie den " & (i + 1) & ". Suchbegriff ein:", "Eingabe des Suchbegriffs")

    'Suchstring Cancel?
    If StrPtr(suchstring) = 0 Then
        Antwort5 = MsgBox(Mldg5, Stil5, Titel5)
        GoTo ende:
    Else
        'OK und kein Suchstring?
        If Len(suchstring) = 0 Then
            If MsgBox("Suchen & Ergänzen kann nicht stattfinden, weil kein Suchbegriff / keine ErgÃ¤nzung eingegeben wurde.", vbRetryCancel, "Suchen & ErgÃƒÆ’Ã‚Â¤nzen fehlgeschlagen!") = vbRetry Then
                GoTo Sucheingabe:
            Else
                Antwort5 = MsgBox(Mldg5, Stil5, Titel5)
                GoTo ende:
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
        response = MsgBox("Suchen & Ergänzen kann nicht stattfinden, weil der Suchbegriff " & Chr(34) & suchstring & Chr(34) & " nicht gefunden werden konnte.", vbOKOnly + vbCritical, "Suchen & ErgÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¤nzen fehlgeschlagen!")
        GoTo ende
    End If

Ergaenzeingabe:
    'Ergaenzstring
    ergaenzstring = InputBox("Bitte geben Sie die " & (i + 1) & ". Bezugsbezeichnung ein, um die Sie den " & (i + 1) & ". Suchbegriff ergÃƒÆ’Ã‚Â¤nzen mÃƒÆ’Ã‚Â¶chten:", "Eingabe der Bezugsbezeichnung")

    'Ergaenzstring Cancel?
    If StrPtr(ergaenzstring) = 0 Then
        Antwort5 = MsgBox(Mldg5, Stil5, Titel5)
        GoTo ende:
    Else
        'OK und kein Suchstring?
        If Len(ergaenzstring) = 0 Then
            If MsgBox("Suchen & Ergänzen kann nicht stattfinden, weil kein Suchbegriff / keine Ergänzung eingegeben wurde.", vbRetryCancel, "Suchen & ErgÃƒÆ’Ã‚Â¤nzen fehlgeschlagen!") = vbRetry Then
                GoTo Ergaenzeingabe:
            Else
                Antwort5 = MsgBox(Mldg5, Stil5, Titel5)
                GoTo ende:
            End If
        End If
    End If

    'Array aus manueller Eingabe erstellen
    i = i + 1: ReDim Preserve ar(i)
    ar(i) = suchstring & "@" & ergaenzstring

    If MsgBox("Möchten Sie weitere Bezugsbezeichnungen einfügen?", vbYesNo, "Wiederholung") = vbYes Then
        GoTo Sucheingabe
    End If

ende:
    MsgBox ("ENDE")
    'Nachverfolgungseinstellungen wiederherstellen
     ActiveWindow.View.MarkupMode = markup
     ActiveDocument.TrackRevisions = trackrev

    Selection.HomeKey unit:=wdStory

    'Suchparameter zurücksetzen
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