'Function checks if passed character is a letter [a-zA-Z]
Function IsLetter(char As String) As Boolean
    Dim charCode As Integer
    charCode = Asc(UCase(char))
    If charCode >= Asc("A") And charCode <= Asc("Z") Then
        IsLetter = True
    Else
        IsLetter = False
    End If
End Function

'Function that returns a string with wildcards, enabling the search for both upper and lowercase strings
Function GetCaseInsensitiveSearchString(suchstring) As String
    Dim lowercase As String
    Dim uppercase As String
    Dim cutFirstLetter As String

    lowercase = LCase(Left(suchstring, 1))
    uppercase = UCase(lowercase)
    cutFirstLetter = Mid(suchstring, 2)

    GetCaseInsensitiveSearchString = "[" & lowercase & "," & uppercase & "]" & cutFirstLetter
End Function

'Function that returns the letter before range
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

'Function that returns the first letter after range
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

'Function that returns the second letter after range
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

'Function extracts suchbegriffe from bezug.txt and checks for syntactical errors
Function GetBezugArray(myPath As String) As Variant
    Dim result(1 To 2) As Variant
    Open myPath For Input As #1
    Dim suchstring, strVariable As String
    Dim bezugArray() As String
    Dim splitar() As String
    ReDim bezugArray(0), splitar(0)
    Dim i, l As Integer

    'decide if all not founds in one MsgBox or all in single ones
    Dim everyNotFoundInOneMsgBox As Boolean
    Dim notFounds As String
    Dim appendString As String
    notFounds = "Die Suchbegriffe: " & vbCrLf
    everyNotFoundInOneMsgBox = False

    i = 0

    Do While Not EOF(1)
        'Änderung 19.01.2017: Befehl "Line" hinzugefüt (Zeile samt Komma als String einlesen), damit mehrere kommagetrennte Bezugszeichen benutzt werden kÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¶nnen
        Line Input #1, strVariable
        i = i + 1: ReDim Preserve bezugArray(i)
        bezugArray(i) = strVariable
        'PrÃƒÂ¼fung, ob Trennzeichen vorhanden
        If InStr(strVariable, "@") = 0 Then
            result(1) = "ENDE"
            Dim lol As String
            lol = MsgBox("Trennzeichen (@) in Datei bezug.txt fehlt! Vorgang wird abgebrochen!", vbCritical, "TrennzeichenprÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¼fung")
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

        'decide if all not founds in one MsgBox or all in single ones
        If everyNotFoundInOneMsgBox = False Then
            If Selection.Find.Found = False Then
                If MsgBox("Der Suchbegriff " & Chr(34) & suchstring & Chr(34) & " konnte nicht gefunden werden." & vbCrLf & vbCrLf & "Trotzdem fortfahren?", vbYesNo + vbCritical, "Suchen & ErgÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¤nzen fehlgeschlagen!") = vbNo Then
                    'fix later: collect not founds in array and display yesNo Box
                    Close #1
                    result(1) = "ENDE"
                    GetBezugArray = result
                    Exit Function
                End If
            End If
        Else
            appendString = " - " & suchstring & vbCrLf
            notFounds = notFounds + appendString
        End If
    Loop
    Close #1

    'decide if all not founds in one MsgBox or all in single ones
    If everyNotFoundInOneMsgBox Then
        appendString = " konnte nicht gefunden werden." & vbCrLf & vbCrLf & "Trotzdem fortfahren?"
        notFounds = notFounds + appendString
        If MsgBox(notFounds, vbYesNo + vbCritical, "Suchen & Ergänzen fehlgeschlagen!") = vbNo Then
            Close #1
            result(1) = "ENDE"
            GetBezugArray = result
            Exit Function
        End If
    End If

    If strVariable <> "" Then
        result(1) = "OK"
        result(2) = bezugArray
    Else
        If MsgBox("Datei bezug.txt ist leer. Fortfahren mit manueller Eingabe?", vbYesNo, "Inhaltsprüfung") = vbYes Then
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
            If MsgBox("Datei bezug.txt nicht vorhanden. Fortfahren mit manueller Eingabe?", vbYesNo, "DateiprÃƒÂ¼fung") = vbYes Then
                result(1) = "SUCHEINGABE"
            Else
                result(1) = "ENDE"
            End If
        Else
            If MsgBox("Datei bezug.txt vorhanden. Fortfahren?", vbYesNo, "DateiprÃƒÂ¼fung") = vbYes Then
                result(1) = "OK"
                result(2) = path
            Else
                result(1) = "ENDE"
            End If
        End If
    Else 'delete this section later
        'my thing to make it work for citrix:
        path = "\\brefile11.esp.dom\citrixprofiles$\msg\Desktop\bezug.txt"
        If Dir(path) <> "" Then
            If MsgBox("Datei bezug.txt vorhanden. Fortfahren?", vbYesNo, "DateiprÃƒÂ¼fung") = vbYes Then
                result(1) = "OK"
                result(2) = path
            Else
                result(1) = "ENDE"
            End If
        Else
            If MsgBox("Datei bezug.txt nicht vorhanden. Fortfahren mit manueller Eingabe?", vbYesNo, "DateiprÃƒÂ¼fung") = vbYes Then
                result(1) = "SUCHEINGABE"
            Else
                result(1) = "ENDE"
            End If
        End If
    End If
    getEnvironmentPath = result
End Function

'basic checks for word document
Function InvalidDocument() As Boolean
    Dim response As String
    Dim effC As Variant
    Dim effS, effI As Integer

    'no documents open
    If Documents.Count < 1 Then
        response = MsgBox("Es ist kein Dokument geÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¶ffnet.", vbOKOnly + vbCritical, "Suchen & ErgÃƒÂ¤nzen fehlgeschlagen!")
        InvalidDocument = True
        Exit Function
    End If

    'Anzahl der Zeichen, Shapes oder InlineShapes feststellen
    effC = ActiveDocument.BuiltInDocumentProperties(wdPropertyCharsWSpaces)
    effS = ActiveDocument.Shapes.Count
    effI = ActiveDocument.InlineShapes.Count

    'keine Zeichen, Shapes oder InlineShapes
    If effC < 1 And effS < 1 And effI < 1 Then
        response = MsgBox("Suchen & ErgÃƒÂ¤nzen im leeren Dokument nicht mÃƒÂ¶glich.", vbOKOnly + vbCritical, "Suchen & ErgÃƒÂ¤nzen fehlgeschlagen!")
        InvalidDocument = True
        Exit Function
    End If
    InvalidDocument = False
End Function

Sub SetTrackingSettings()
    Dim o, p As Integer
    'Ã„nderung 19.01.2017: prÃ¼fen, ob nicht angenommene Ã„nderungenen eines anderen Benutzers vorhanden sind
    o = ActiveDocument.Revisions.Count
    For p = 1 To o
        If ActiveDocument.Revisions.Count <> 0 And ActiveDocument.Revisions(p).Author <> Application.UserName Then
            MsgBox "ACHTUNG:" & vbCrLf & vbCrLf & "Nicht angenommene Ã„nderungenen eines anderen Benutzers (" & ActiveDocument.Revisions(p).Author & ") vorhanden - nachtrÃ¤gliche ErgÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¤nzungen beeinflussen diese ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¾nderungen!" & vbCrLf & vbCrLf & "Bitte anschlieÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â¦ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¸end prÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¼fen!", vbOKOnly + vbExclamation, "Suchen & ErgÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¤nzen"
            Exit Sub
        End If
    Next

    'wenn 'Ã„nderungen verfolgen' deaktiviert ist -> aktivieren
    If ActiveDocument.TrackRevisions = False Then
        ActiveDocument.TrackRevisions = True
    End If

    'Sprechblasen-Einstellung auf balloon umstellen
    ActiveWindow.View.MarkupMode = wdBalloonRevisions

    'ErklÃ¤rung
    Dim response As String
    response = MsgBox("Die Datei muss auf dem Desktop mit dem Dateinamen bezug.txt angelegt werden und muss folgende zeilenweise Syntax aufweisen:" & vbCrLf & vbCrLf & _
    "Suchbegriff1@Bezugsbezeichnung1" & vbCrLf & "Suchbegriff2@Bezugsbezeichnung2" & vbCrLf & "..." & vbCrLf & vbCrLf & _
    "Es wird nur nach ganzen WÃ¶rtern gesucht." & vbCrLf & "Die Bezugsbezeichnung (nur Zahl, ohne Suchbegriff) wird beim EinfÃ¼gen automatisch in Klammern gesetzt.", vbInformation, "ErklÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¤rung")
End Sub

'Function sorts based on lenght the Suchebegriffe from the bezug.txt file
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

'Function that checks if newly found search term conflicts with previously identified search term ranges
'It skips strings that are contained within other strings found in the document
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

'creates Array with start and end Range of every found Suchbegriff in the word document
'strings that are Plural [ending with 's'] or capitalized in the beginning are also valid now
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
        text = GetCaseInsensitiveSearchString(suchstring)
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

'Insert Bezugszeichen in word document after found Suchbegriff
Sub InsertNumbers(rangesArray() As Variant)
    Dim docContent As range
    Dim k As Integer

    Set docContent = ActiveDocument.Content
    For k = LBound(rangesArray, 2) To UBound(rangesArray, 2)
        docContent.SetRange start:=rangesArray(0, k), End:=rangesArray(1, k)
        docContent.InsertAfter rangesArray(2, k)
    Next k
End Sub

'Bubble sort to sort rangesArray based on Range.End Value [rangesArray(1,x)]
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

    Mldg5 = "Suchen & ErgÃ¤nzen durch Abbruch beendet."
    Stil5 = vbInformation
    Titel5 = "Suchen & ErgÃ¤nzen abgebrochen!"

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
            If MsgBox("Suchen & ErgÃ¤nzen kann nicht stattfinden, weil kein Suchbegriff / keine ErgÃƒÂ¤nzung eingegeben wurde.", vbRetryCancel, "Suchen & ErgÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¤nzen fehlgeschlagen!") = vbRetry Then
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
        response = MsgBox("Suchen & ErgÃ¤nzen kann nicht stattfinden, weil der Suchbegriff " & Chr(34) & suchstring & Chr(34) & " nicht gefunden werden konnte.", vbOKOnly + vbCritical, "Suchen & ErgÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¤nzen fehlgeschlagen!")
        GoTo ende
    End If

Ergaenzeingabe:
    'Ergaenzstring
    ergaenzstring = InputBox("Bitte geben Sie die " & (i + 1) & ". Bezugsbezeichnung ein, um die Sie den " & (i + 1) & ". Suchbegriff ergÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¤nzen mÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¶chten:", "Eingabe der Bezugsbezeichnung")

    'Ergaenzstring Cancel?
    If StrPtr(ergaenzstring) = 0 Then
        Antwort5 = MsgBox(Mldg5, Stil5, Titel5)
        GoTo ende:
    Else
        'OK und kein Suchstring?
        If Len(ergaenzstring) = 0 Then
            If MsgBox("Suchen & ErgÃ¤nzen kann nicht stattfinden, weil kein Suchbegriff / keine ErgÃ¤nzung eingegeben wurde.", vbRetryCancel, "Suchen & ErgÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¤nzen fehlgeschlagen!") = vbRetry Then
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

    If MsgBox("MÃ¶chten Sie weitere Bezugsbezeichnungen einfÃ¼gen?", vbYesNo, "Wiederholung") = vbYes Then
        GoTo Sucheingabe
    End If

ende:
    MsgBox ("ENDE")
    'Nachverfolgungseinstellungen wiederherstellen
     ActiveWindow.View.MarkupMode = markup
     ActiveDocument.TrackRevisions = trackrev

    Selection.HomeKey unit:=wdStory

    'Suchparameter zurÃ¼cksetzen
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