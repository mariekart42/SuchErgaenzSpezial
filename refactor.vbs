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
    lowercase = LCase(Left(suchstring, 1))
    Dim uppercase As String
    uppercase = UCase(lowercase)
    Dim cutFirstLetter As String
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
    copyRange.MoveStart Unit:=wdCharacter, Count:=-1
    letterBefore = copyRange.Characters(1).text
    Set copyRange = Nothing
    getLetterBefore = letterBefore
End Function

Function getLetterAfter(rangeObj As Range) As String
    Dim copyRange As Range
    Dim letterAfter As String

    Set copyRange = rangeObj.Duplicate

    copyRange.MoveEnd Unit:=wdCharacter, Count:=1
    If copyRange.End = copyRange.Document.Content.End Then
        getLetterAfter = "EOF"
        Set copyRange = Nothing
        Exit Function
    End If
    copyRange.MoveEnd Unit:=wdCharacter, Count:=-1
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

    copyRange.MoveEnd Unit:=wdCharacter, Count:=2
    ' Check if the range is at the beginning of the document
    If copyRange.End = copyRange.Document.Content.End Then
        getSecondLetterAfter = "EOF"
        Set copyRange = Nothing
        Exit Function
    End If

    ' Move the range one character backward
    copyRange.MoveEnd Unit:=wdCharacter, Count:=-2
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
        'Änderung 19.01.2017: Befehl "Line" hinzugefüt (Zeile samt Komma als String einlesen), damit mehrere kommagetrennte Bezugszeichen benutzt werden kÃƒÆ’Ã‚Â¶nnen
        Line Input #1, strVariable
        i = i + 1: ReDim Preserve bezugArray(i)
        bezugArray(i) = strVariable
        'Prüfung, ob Trennzeichen vorhanden
        If InStr(strVariable, "@") = 0 Then
            result(1) = "ENDE"
            Dim lol As String
            lol = MsgBox("Trennzeichen (@) in Datei bezug.txt fehlt! Vorgang wird abgebrochen!", vbCritical, "Trennzeichenprüfung")
            Close #1
            GetBezugArray = result
            Exit Function
        End If

        suchstring = Left(bezugArray(i), InStr(bezugArray(i), "@") - 1)
        l = l + 1: ReDim Preserve splitar(l)
        splitar(l) = suchstring
        Selection.HomeKey Unit:=wdStory
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
            If MsgBox("Der Suchbegriff " & Chr(34) & suchstring & Chr(34) & " konnte nicht gefunden werden." & vbCrLf & vbCrLf & "Trotzdem fortfahren?", vbYesNo + vbCritical, "Suchen & ErgÃ¤nzen fehlgeschlagen!") = vbNo Then
                'fix later: collect not founds in array and display yesNo Box
                Close #1
                result(1) = "ENDE"
                GetBezugArray = result
                Exit Function
            End If
        End If
    Loop
    Close #1

    If FoundDuplicate(splitar) Then
        MsgBox ("BZEE")
        result(1) = "DUPLICATE"
    ElseIf strVariable <> "" Then
        result(1) = "OK"
        result(2) = bezugArray
    Else
        If MsgBox("Datei bezug.txt ist leer. Fortfahren mit manueller Eingabe?", vbYesNo, "InhaltsprÃ¼fung") = vbYes Then
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
            If MsgBox("Datei bezug.txt nicht vorhanden. Fortfahren mit manueller Eingabe?", vbYesNo, "DateiprÃ¼fung") = vbYes Then
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
    Else 'delete this section later, only for me testing
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
Function InvalidDocument() As Boolean
    Dim response As String
    Dim effC As Variant
    Dim effS, effI As Integer

    'no documents open
    If Documents.Count < 1 Then
        response = MsgBox("Es ist kein Dokument geöffnet.", vbOKOnly + vbCritical, "Suchen & ErgÃ¤nzen fehlgeschlagen!")
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
    'Änderung 19.01.2017: prüfen, ob nicht angenommene Änderungen eines anderen Benutzers vorhanden sind
    o = ActiveDocument.Revisions.Count
    For p = 1 To o
        If ActiveDocument.Revisions.Count <> 0 And ActiveDocument.Revisions(p).Author <> Application.UserName Then
            MsgBox "ACHTUNG:" & vbCrLf & vbCrLf & "Nicht angenommene Ãƒâ€žnderungen eines anderen Benutzers (" & ActiveDocument.Revisions(p).Author & ") vorhanden - nachtrÃƒÂ¤gliche ErgÃƒÂ¤nzungen beeinflussen diese Ãƒâ€žnderungen!" & vbCrLf & vbCrLf & "Bitte anschlieÃƒÅ¸end prÃƒÂ¼fen!", vbOKOnly + vbExclamation, "Suchen & ErgÃƒÂ¤nzen"
            Exit Sub
        End If
    Next

    'wenn 'Ã„nderungen verfolgen' deaktiviert ist -> aktivieren
    If ActiveDocument.TrackRevisions = False Then
        ActiveDocument.TrackRevisions = True
    End If

    'Sprechblasen-Einstellung auf balloon umstellen
    ActiveWindow.View.MarkupMode = wdBalloonRevisions

    'Erklärung
    Antwort9 = MsgBox("Die Datei muss auf dem Desktop mit dem Dateinamen bezug.txt angelegt werden und muss folgende zeilenweise Syntax aufweisen:" & vbCrLf & vbCrLf & _
    "Suchbegriff1@Bezugsbezeichnung1" & vbCrLf & "Suchbegriff2@Bezugsbezeichnung2" & vbCrLf & "..." & vbCrLf & vbCrLf & _
    "Es wird nur nach ganzen Wörtern gesucht." & vbCrLf & "Die Bezugsbezeichnung (nur Zahl, ohne Suchbegriff) wird beim Einfügen automatisch in Klammern gesetzt.", vbInformation, "Erklärung")
End Sub



Sub SuchErgaenzSpezial()

'
'Makro vom 30.09.2016 von Jacek Manka
'bearbeitet 10.03.2024 von Marie Mensing
'



'Variablen setzen
Dim markup, trackrev
Dim Titel1, Titel2, Titel3, Titel4, Titel5, Titel6, Titel7, Titel8, Titel9, Titel10, Titel11, Titel12, Titel13, _
Mldg1, Mldg2, Mldg3, Mldg4, Mldg5, Mldg6, Mldg7, Mldg8, Mldg9, Mldg10, Mldg11, Mldg12, Mldg13, _
Stil1, Stil2, Stil3, Stil4, Stil5, Stil6, Stil7, Stil8, Stil9, Stil10, Stil11, Stil12, Stil13, _
Antwort1, Antwort2, Antwort3, Antwort4, Antwort5, Antwort6, Antwort7, Antwort8, Antwort9, Antwort10, Antwort11, Antwort12, Antwort13
Dim strVariable, suchstring, ergaenzstring, user, profil, Pfad, server, ar(), splitar() As String
Dim i, k, l, m, n, o, p As Integer

ReDim ar(0), splitar(0)
i = 0
user = Environ("Username")
profil = Environ("AppData")

    'Meldungen erstellen
    Mldg1 = "Suchen & ErgÃƒÂ¤nzen kann nicht stattfinden, weil der Suchbegriff nicht gefunden werden konnte."
    Stil1 = vbOKOnly + vbCritical
    Titel1 = "Suchen & ErgÃƒÂ¤nzen fehlgeschlagen!"

    'Mldg2 = "Bitte geben Sie den " & (i + 1) & ". Suchbegriff ein:"
    'Stil2 = vbOKOnly
    Titel2 = "Eingabe des Suchbegriffs"

    'Mldg3 = "Bitte geben Sie den " & (i + 1) & ". Begriff ein, um den Sie den " & (i + 1) & ". Suchbegriff ergÃƒÂ¤nzen mÃƒÂ¶chten:"
    'Stil3 = vbOKOnly
    Titel3 = "Eingabe der Bezugsbezeichnung"

    Mldg4 = "Suchen & ErgÃƒÂ¤nzen kann nicht stattfinden, weil kein Suchbegriff / keine ErgÃƒÂ¤nzung eingegeben wurde."
    Stil4 = vbRetryCancel
    Titel4 = "Suchen & ErgÃƒÂ¤nzen fehlgeschlagen!"

    Mldg5 = "Suchen & ErgÃƒÂ¤nzen durch Abbruch beendet."
    Stil5 = vbInformation
    Titel5 = "Suchen & ErgÃƒÂ¤nzen abgebrochen!"

    Mldg6 = "Datei bezug.txt vorhanden. Fortfahren?"
    Stil6 = vbYesNo
    Titel6 = "DateiprÃƒÂ¼fung"

    Mldg7 = "Datei bezug.txt nicht vorhanden. Fortfahren mit manueller Eingabe?"
    Stil7 = vbYesNo
    Titel7 = "DateiprÃƒÂ¼fung"

    Mldg8 = "Trennzeichen (@) in Datei bezug.txt fehlt! Vorgang wird abgebrochen!"
    Stil8 = vbCritical
    Titel8 = "TrennzeichenprÃƒÂ¼fung"

    Mldg9 = "Sie kÃ¶nnen die Bezugsbezeichnungen per Datei oder manuell einfÃ¼gen." & vbCrLf & vbCrLf & _
    "Die Datei muss auf dem Desktop mit dem Dateinamen bezug.txt angelegt werden und muss folgende zeilenweise Syntax aufweisen:" & vbCrLf & vbCrLf & _
    "Suchbegriff1@Bezugsbezeichnung1" & vbCrLf & "Suchbegriff2@Bezugsbezeichnung2" & vbCrLf & "..." & vbCrLf & vbCrLf & _
    "Es wird nur nach ganzen WÃ¶rtern gesucht." & vbCrLf & "Die Bezugsbezeichnung (nur Zahl, ohne Suchbegriff) wird beim EinfÃ¼gen automatisch in Klammern gesetzt."
    Stil9 = vbInformation
    Titel9 = "ErklÃƒÂ¤rung"

    Mldg10 = "Datei bezug.txt ist leer. Fortfahren mit manueller Eingabe?"
    Stil10 = vbYesNo
    Titel10 = "InhaltsprÃƒÂ¼fung"

    Mldg11 = "MÃƒÂ¶chten Sie weitere Bezugsbezeichnungen einfÃƒÂ¼gen?"
    Stil11 = vbYesNo
    Titel11 = "Wiederholung"

    Mldg12 = "Es ist kein Dokument geÃƒÂ¶ffnet."
    Stil12 = vbOKOnly + vbCritical
    Titel12 = "Suchen & ErgÃƒÂ¤nzen fehlgeschlagen!"

    Mldg13 = "Suchen & ErgÃƒÂ¤nzen im leeren Dokument nicht mÃƒÂ¶glich."
    Stil13 = vbOKOnly + vbCritical
    Titel13 = "Suchen & ErgÃƒÂ¤nzen fehlgeschlagen!"



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
                GoTo SuchErsetz
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
    suchstring = InputBox("Bitte geben Sie den " & (i + 1) & ". Suchbegriff ein:", Titel2)

    'Suchstring Cancel?
    If StrPtr(suchstring) = 0 Then
        Antwort5 = MsgBox(Mldg5, Stil5, Titel5)
        GoTo Ende:
    Else
        'OK und kein Suchstring?
        If Len(suchstring) = 0 Then
            Antwort4 = MsgBox(Mldg4, Stil4, Titel4)
            If Antwort4 = vbRetry Then
                GoTo Sucheingabe:
            Else
                Antwort5 = MsgBox(Mldg5, Stil5, Titel5)
                GoTo Ende:
            End If
        End If
    End If

    'Suchen
    Selection.HomeKey Unit:=wdStory
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
        Antwort1 = MsgBox("Suchen & Ergänzen kann nicht stattfinden, weil der Suchbegriff " & Chr(34) & suchstring & Chr(34) & " nicht gefunden werden konnte.", vbOKOnly + vbCritical, "Suchen & ErgÃƒÂ¤nzen fehlgeschlagen!")
        GoTo Ende
    End If

Ergaenzeingabe:
    'Ergaenzstring
    ergaenzstring = InputBox("Bitte geben Sie die " & (i + 1) & ". Bezugsbezeichnung ein, um die Sie den " & (i + 1) & ". Suchbegriff ergÃƒÂ¤nzen mÃƒÂ¶chten:", Titel3)

    'Ergaenzstring Cancel?
    If StrPtr(ergaenzstring) = 0 Then
        Antwort5 = MsgBox(Mldg5, Stil5, Titel5)
        GoTo Ende:
    Else

        'OK und kein Suchstring?
        If Len(ergaenzstring) = 0 Then
            Antwort4 = MsgBox(Mldg4, Stil4, Titel4)
            If Antwort4 = vbRetry Then
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

    If MsgBox("Möchten Sie weitere Bezugsbezeichnungen einfügen?", vbYesNo, "Wiederholung") = vbYes Then
        GoTo Sucheingabe
    End If

SuchErsetz:
    i = (UBound(bezugArray) - LBound(bezugArray) + 1) - 1
    For k = 1 To i
        suchstring = Left(bezugArray(k), InStr(bezugArray(k), "@") - 1)
        ergaenzstring = " (" & Right(bezugArray(k), Len(bezugArray(k)) - InStr(bezugArray(k), "@")) & ")"
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

    Selection.HomeKey Unit:=wdStory

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