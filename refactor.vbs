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
    ' Create a copy of the original range
    Dim copyRange As Range
    Set copyRange = rangeObj.Duplicate
    Dim letterBefore As String

    ' Check if the range is at the beginning of the document
    If copyRange.Start = 0 Then
        'MsgBox "The range is at the beginning of the document."
        getLetterBefore = "BOF"
        Set copyRange = Nothing
        Exit Function
    End If

    ' Move the range one character backward
    copyRange.MoveStart Unit:=wdCharacter, Count:=-1

    ' Get the text of the character before the range
    letterBefore = copyRange.Characters(1).text

    ' Display the character before the range
    'MsgBox "Character before the range: [" & letterBefore & "]"
    Set copyRange = Nothing
    getLetterBefore = letterBefore
End Function

Function getLetterAfter(rangeObj As Range) As String
'MsgBox ("HEHE: |" & rangeObj.text & "|")
    ' Create a copy of the original range
    Dim copyRange As Range
    Set copyRange = rangeObj.Duplicate
    Dim letterAfter As String
    'MsgBox ("before: |" & copyRange.text & "|")
    copyRange.MoveEnd Unit:=wdCharacter, Count:=1
    'MsgBox ("after: |" & copyRange.text & "|")
    ' Check if the range is at the end of the document
    If copyRange.End = copyRange.Document.Content.End Then
        'MsgBox "The range is at the end of the document."
        getLetterAfter = "EOF"
        Set copyRange = Nothing
        Exit Function
    End If


     ' Check if the range is at the beginning of the document
    'If copyRange.Start = 0 Then
        ' Extend the range by one character
        'copyRange.Next wdCharacter, Count:=1
        'copyRange.MoveEnd unit:=wdCharacter, Count:=-1
        'letterAfter = copyRange.Characters(copyRange.Characters.Count).text
    '    MsgBox ("before: |" & copyRange.text & "|")
    '    copyRange.Next wdCharacter, rangeObj.Characters.Count
    '    MsgBox ("after: |" & copyRange.text & "|")
    '    letterAfter = copyRange.Characters(1).text

     '   MsgBox ("SPECIAL CASE beginn: letter after: |" & letterAfter & "|")
     '   getLetterAfter = letterAfter
     '   Exit Function

    'End If


    copyRange.MoveEnd Unit:=wdCharacter, Count:=-1
    'MsgBox ("after after: |" & copyRange.text & "|")

    'copyRange.MoveStart unit:=wdCharacter, Count:=1

    ' Move the range one character backward
    'copyRange.MoveEnd unit:=wdCharacter, Count:=-1

    'Dim last As String
    'last = copyRange.Characters(copyRange.End).text
    copyRange.MoveEnd wdCharacter
    letterAfter = copyRange.Characters(copyRange.Characters.Count).text
    'MsgBox "LAST LETTER AFTER: |" & letterAfter & "|"

    ' Get the text of the character after the range
    'letterAfter = copyRange.Characters(1).text
    'MsgBox "LAST LETTER AFTER: |" & letterAfter & "|"

    ' Display the character before the range
    'MsgBox "Character after the range: [" & letterAfter & "]"
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
        'MsgBox "The range is at the end of the document."
        getSecondLetterAfter = "EOF"
        Set copyRange = Nothing
        Exit Function
    End If

    ' Move the range one character backward
    copyRange.MoveEnd Unit:=wdCharacter, Count:=-2

    copyRange.MoveEnd wdCharacter, 2
    letterAfter = copyRange.Characters(copyRange.Characters.Count).text
    ' Get the text of the character after the range
    'letterAfter = copyRange.Characters(1).text

    ' Display the character before the range
    'MsgBox "Character after the range: [" & letterAfter & "]"
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
        'Ãƒâ€žnderung 19.01.2017: Befehl "Line" hinzugefÃƒÂ¼gt (Zeile samt Komma als String einlesen), damit mehrere kommagetrennte Bezugszeichen benutzt werden kÃƒÂ¶nnen
        Line Input #1, strVariable
        i = i + 1: ReDim Preserve bezugArray(i)
        bezugArray(i) = strVariable

        'PrÃƒÂ¼fung, ob Trennzeichen vorhanden
        If InStr(strVariable, "@") = 0 Then
            result(1) = "ENDE"
            Dim lol As String
            lol = MsgBox("Trennzeichen (@) in Datei bezug.txt fehlt! Vorgang wird abge-brochen!", vbCritical, "TrennzeichenprÃ¼fung")
            Close #1
            GetBezugArray = result
            Exit Function
            'GoTo Ende
        End If

        'PrÃƒÂ¼fung des Suchstrings
        suchstring = Left(bezugArray(i), InStr(bezugArray(i), "@") - 1)
        'MsgBox ("suchstring before ergeanz: " & suchstring)
        l = l + 1: ReDim Preserve splitar(l)
        splitar(l) = suchstring
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
            .MatchAllWordForms = False 'proplem with sentences if True!
            .MatchSoundsLike = False
            .MatchWildcards = False
        End With
        Selection.Find.Execute
        If Selection.Find.Found = False Then
            Dim msg As String
            msg = MsgBox("Der Suchbegriff " & Chr(34) & suchstring & Chr(34) & " konnte nicht gefunden werden." & vbCrLf & vbCrLf & "Trotzdem fortfahren?", vbYesNo + vbCritical, "Suchen & ErgÃƒÂ¤nzen fehlgeschlagen!")
            If msg = vbNo Then
                'fix later: collect not founds in array and display yesNo Box
                Close #1
                result(1) = "ENDE" 'go to Ende
                GetBezugArray = result
                Exit Function
                'GoTo Ende
            End If
        End If

    Loop
    Close #1

    If FoundDuplicate(splitar) Then
        MsgBox ("BZEE")
        result(1) = "DUPLICATE"
        GetBezugArray = result
        Exit Function
    End If

    'Datei leer?
    If strVariable <> "" Then
        result(1) = "OK" 'go to suchErsetz
        result(2) = bezugArray
        GetBezugArray = result
        Exit Function
        'GoTo SuchErsetz
    Else
        Dim msg2 As String
        msg2 = MsgBox("Datei bezug.txt ist leer. Fortfahren mit manueller Eingabe?", vbYesNo, "InhaltsprÃ¼fung")
        If msg2 = vbYes Then
            result(1) = "DOCUMENT EMPTY" 'go to Sucheingabe
            GetBezugArray = result
            Exit Function
            'GoTo Sucheingabe
        Else
            result(1) = "ENDE" 'go to Ende
            GetBezugArray = result
            'GoTo Ende
        End If
    End If
    result(1) = "OK" 'go to suchErsetz
    result(2) = bezugArray
    GetBezugArray = result
End Function


Function getEnvironmentPath() As Variant
    Dim result(1 To 2) As Variant
    Dim path As String

    'my thing to make it work for citrix:
    path = "\\brefile11.esp.dom\citrixprofiles$\msg\Desktop\bezug.txt"
    If Dir(path) <> "" Then
        Dim msg1 As String
        msg1 = MsgBox("Datei bezug.txt vorhanden. Fortfahren?", vbYesNo, "Dateiprüfung")
        If msg1 = vbYes Then
            result (1)
            getEnvironmentPath = "WEITER" 'change
        Else
            getEnvironmentPath = "ENDE"
        End If
    Else
        Dim msg2 As String
        msg2 = MsgBox("Datei bezug.txt nicht vorhanden. Fortfahren mit manueller Eingabe?", vbYesNo, "Dateiprüfung")
        If msg2 = vbYes Then
            getEnvironmentPath = "SUCHEINGABE" 'change
        Else
            getEnvironmentPath = "ENDE"
        End If
    End If

End Function



Sub SuchErgaenzSpezial()

'
'Makro vom 30.09.2016 von Jacek Manka
'bearbeitet 10.03.2024 von Marie Mensing
'

'Public Type bzg
'    Dim rangeArray() As Variant 'stores [range.start|range.end|bzgzeichen]
'    Dim bezugArray() As String
'End Type




'Variablen setzen
Dim markup, trackrev
Dim Titel1, Titel2, Titel3, Titel4, Titel5, Titel6, Titel7, Titel8, Titel9, Titel10, Titel11, Titel12, Titel13, _
Mldg1, Mldg2, Mldg3, Mldg4, Mldg5, Mldg6, Mldg7, Mldg8, Mldg9, Mldg10, Mldg11, Mldg12, Mldg13, _
Stil1, Stil2, Stil3, Stil4, Stil5, Stil6, Stil7, Stil8, Stil9, Stil10, Stil11, Stil12, Stil13, _
Antwort1, Antwort2, Antwort3, Antwort4, Antwort5, Antwort6, Antwort7, Antwort8, Antwort9, Antwort10, Antwort11, Antwort12, Antwort13
Dim strVariable, suchstring, ergaenzstring, user, profil, Pfad, Server, ar(), splitar() As String
Dim i, k, l, m, n, o, p As Integer
Dim effC As Variant
Dim effS, effI As Integer
ReDim ar(0), splitar(0)
i = 0
user = Environ("Username")
profil = Environ("AppData")

    'Meldungen erstellen
    Mldg1 = "Suchen & ErgÃ¤nzen kann nicht stattfinden, weil der Suchbegriff nicht gefunden werden konnte."
    Stil1 = vbOKOnly + vbCritical
    Titel1 = "Suchen & ErgÃ¤nzen fehlgeschlagen!"

    'Mldg2 = "Bitte geben Sie den " & (i + 1) & ". Suchbegriff ein:"
    'Stil2 = vbOKOnly
    Titel2 = "Eingabe des Suchbegriffs"

    'Mldg3 = "Bitte geben Sie den " & (i + 1) & ". Begriff ein, um den Sie den " & (i + 1) & ". Suchbegriff ergÃ¤nzen mÃ¶chten:"
    'Stil3 = vbOKOnly
    Titel3 = "Eingabe der Bezugsbezeichnung"

    Mldg4 = "Suchen & ErgÃ¤nzen kann nicht stattfinden, weil kein Suchbegriff / keine ErgÃ¤nzung eingegeben wurde."
    Stil4 = vbRetryCancel
    Titel4 = "Suchen & ErgÃ¤nzen fehlgeschlagen!"

    Mldg5 = "Suchen & ErgÃ¤nzen durch Abbruch beendet."
    Stil5 = vbInformation
    Titel5 = "Suchen & ErgÃ¤nzen abgebrochen!"

    Mldg6 = "Datei bezug.txt vorhanden. Fortfahren?"
    Stil6 = vbYesNo
    Titel6 = "DateiprÃ¼fung"

    Mldg7 = "Datei bezug.txt nicht vorhanden. Fortfahren mit manueller Eingabe?"
    Stil7 = vbYesNo
    Titel7 = "DateiprÃ¼fung"

    Mldg8 = "Trennzeichen (@) in Datei bezug.txt fehlt! Vorgang wird abgebrochen!"
    Stil8 = vbCritical
    Titel8 = "TrennzeichenprÃ¼fung"

    Mldg9 = "Sie kÃ¶nnen die Bezugsbezeichnungen per Datei oder manuell einfÃ¼gen." & vbCrLf & vbCrLf & _
    "Die Datei muss auf dem Desktop mit dem Dateinamen bezug.txt angelegt werden und muss folgende zeilenweise Syntax aufweisen:" & vbCrLf & vbCrLf & _
    "Suchbegriff1@Bezugsbezeichnung1" & vbCrLf & "Suchbe-griff2@Bezugsbezeichnung2" & vbCrLf & "..." & vbCrLf & vbCrLf & _
    "Bitte achten Sie bei den Suchbegriffen auf GroÃŸ-/Kleinschreibung sowie Ein-/Mehrzahl. Es wird nur nach ganzen WÃ¶rtern gesucht." & vbCrLf & "Die Bezugsbez-eichnung (nur Zahl, ohne Suchbegriff) wird beim EinfÃ¼gen automatisch in Klammern gesetzt." & vbCrLf & vbCrLf & "Die Datei bezug.txt wird abschlieÃŸend geleert."
    Stil9 = vbInformation
    Titel9 = "ErklÃ¤rung"

    Mldg10 = "Datei bezug.txt ist leer. Fortfahren mit manueller Eingabe?"
    Stil10 = vbYesNo
    Titel10 = "InhaltsprÃ¼fung"

    Mldg11 = "MÃ¶chten Sie weitere Bezugsbezeichnungen einfÃ¼gen?"
    Stil11 = vbYesNo
    Titel11 = "Wiederholung"

    Mldg12 = "Es ist kein Dokument geÃ¶ffnet."
    Stil12 = vbOKOnly + vbCritical
    Titel12 = "Suchen & ErgÃ¤nzen fehlgeschlagen!"

    Mldg13 = "Suchen & ErgÃ¤nzen im leeren Dokument nicht mÃ¶glich."
    Stil13 = vbOKOnly + vbCritical
    Titel13 = "Suchen & ErgÃ¤nzen fehlgeschlagen!"


    'Abfrage, ob ein Dokument geÃ¶ffnet ist
    If Documents.Count >= 1 Then

        'Anzahl der Zeichen, Shapes oder InlineShapes feststellen
        effC = ActiveDocument.BuiltInDocumentProperties(wdPropertyCharsWSpaces)
        effS = ActiveDocument.Shapes.Count
        effI = ActiveDocument.InlineShapes.Count

        'Abfrage, ob Dokument Zeichen, Shapes oder InlineShapes enthÃ¤lt (auÃŸer Kopf- und FuÃŸzeile)
        If effC >= 1 Or effS >= 1 Or effI >= 1 Then

            'Nachverfolgungseinstellungen sichern
            trackrev = ActiveDocument.TrackRevisions
            markup = ActiveWindow.View.MarkupMode

            'Ã„nderung 19.01.2017: prÃ¼fen, ob nicht angenommene Ã„nderungen eines anderen Benutzers vorhanden sind
            o = ActiveDocument.Revisions.Count
            For p = 1 To o
                If ActiveDocument.Revisions.Count <> 0 And ActiveDocu-ment.Revisions(p).Author <> Application.UserName Then
                    MsgBox "ACHTUNG:" & vbCrLf & vbCrLf & "Nicht angenommene Ã„nderungen eines anderen Benutzers (" & ActiveDocument.Revisions(p).Author & ") vorhanden - nachtrÃ¤gliche ErgÃ¤nzungen beeinflussen diese Ã„nderungen!" & vbCrLf & vbCrLf & "Bitte anschlieÃŸend prÃ¼fen!", vbOKOnly + vbExclamation, "Suchen & ErgÃ¤nzen"
                    Exit For
                End If
            Next

            'wenn 'Ã„nderungen verfolgen' deaktiviert ist -> aktivieren
            If ActiveDocument.TrackRevisions = False Then
                ActiveDocument.TrackRevisions = True
            End If

            'Sprechblasen-Einstellung auf balloon umstellen
            ActiveWindow.View.MarkupMode = wdBalloonRevisions

            'ErklÃ¤rung
            Antwort9 = MsgBox(Mldg9, Stil9, Titel9)

            Dim vals As Variant
            vals = getEnvironmentPath()
            If vals(1) = "ENDE" Then
                GoTo Ende
            ElseIf vals(1) = "SUCHEINGABE" Then
                GoTo Sucheingabe
            Else
                'InsertBezuegszeichen (vals(2))
                'GoTo Ende
                Dim path As String
                path = vals(2)
                Dim bezugArray() As String
                Dim values As Variant
                values = GetBezugArray(path)
                'MsgBox ("value(1): " & values(1))
                If values(1) = "OK" Then
                    bezugArray = values(2)
                    GoTo SuchErsetz
                ElseIf values(1) = "ENDE" Then
                    GoTo Ende
                ElseIf values(1) = "DUPLICATE" Then
                    GoTo Ende
                ElseIf values(1) = "DOCUMENT EMPTY" Then
                    MsgBox ("document is empty lol, sucheingabe:")
                    GoTo Sucheingabe
                End If
            End If

        'sonst keine Zeichen, Shapes oder InlineShapes
        Else
            'Wenn leeres Dokument geÃ¶ffnet ist -> Fehlermeldung
            Antwort13 = MsgBox(Mldg13, Stil13, Titel13)
        End If
    Else
        'Wenn kein Dokument geÃ¶ffnet ist, Makro nicht ausfÃ¼hrbar -> Fehlermeldung
        Antwort12 = MsgBox(Mldg12, Stil12, Titel12)
    End If



Weiter:












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

                Antwort1 = MsgBox("Suchen & ErgÃ¤nzen kann nicht stattfinden, weil der Suchbegriff " & Chr(34) & suchstring & Chr(34) & " nicht gefunden werden konnte.", vbOKOnly + vbCritical, "Suchen & ErgÃ¤nzen fehlgeschlagen!")
                GoTo Ende

            End If

Ergaenzeingabe:

            'Ergaenzstring
            ergaenzstring = InputBox("Bitte geben Sie die " & (i + 1) & ". Bezugsbezeichnung ein, um die Sie den " & (i + 1) & ". Suchbegriff ergÃ¤nzen mÃ¶chten:", Titel3)

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
            Antwort11 = MsgBox(Mldg11, Stil11, Titel11)

            'Eingabe fÃ¼r weitere Strings wiederholen?
            If Antwort11 = vbYes Then
                GoTo Sucheingabe
            End If

SuchErsetz:
            i = (UBound(bezugArray) - LBound(bezugArray) + 1) - 1
            For k = 1 To i
                suchstring = Left(bezugArray(k), InStr(bezugArray(k), "@") - 1)
                ergaenzstring = " (" & Right(bezugArray(k), Len(bezugArray(k)) - In-Str(bezugArray(k), "@")) & ")"

                Dim selectionRange As Range
                Set selectionRange = ActiveDocument.Range

                Dim text As String
                text = GetFindText(suchstring)

                Do While selectionRange.Find.Execute(FindText:=text, MatchAllWord-Forms:=False, MatchSoundsLike:=False, MatchWildcards:=True, Forward:=True) = True
                    Dim letterAfter As String
                    letterAfter = getLetterAfter(selectionRange)
                    Dim secondLetterAfter As String
                    secondLetterAfter = getSecondLetterAfter(selectionRange)
                    'MsgBox (selectionRange.text & "[" & letterAfter & "][" & secondLetterAfter & "]")

                    Dim letterBefore As String
                    letterBefore = getLetterBefore(selectionRange)


                    If (Not IsLetter(letterBefore) And Not IsNumeric(letterBefore)) Or letterBe-fore = "BOF" Then

                        'MsgBox (selectionRange.text & "[" & letterAfter & "][" & secondLetter-After & "]")
                        If letterAfter = "s" Then
                            'check if next character is a letter or number, if yes dont consider as suchstring
                            If Not (IsLetter(secondLetterAfter) And Not IsNumer-ic(secondLetterAfter)) Or secondLetterAfter = "EOF" Then
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

Leeren:

            'If Antwort6 = vbYes Then
            '    'MsgBox "Datei bezug.txt wird geleert."
            '    'Datei leeren
            '    Open Server & user & Pfad & "bezug.txt" For Output As #1
            '    Close #1
            'End If

Ende:
            MsgBox ("ENDE")
            'Nachverfolgungseinstellungen wiederherstellen
             ActiveWindow.View.MarkupMode = markup
             ActiveDocument.TrackRevisions = trackrev

            Selection.HomeKey Unit:=wdStory

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
