Sub SuchErgaenzSpezial()
'
'Makro vom 30.09.2016 von Jacek Manka
'

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
    Mldg1 = "Suchen & Ergänzen kann nicht stattfinden, weil der Suchbegriff nicht gefunden werden konnte."
    Stil1 = vbOKOnly + vbCritical
    Titel1 = "Suchen & Ergänzen fehlgeschlagen!"

    'Mldg2 = "Bitte geben Sie den " & (i + 1) & ". Suchbegriff ein:"
    'Stil2 = vbOKOnly
    Titel2 = "Eingabe des Suchbegriffs"

    'Mldg3 = "Bitte geben Sie den " & (i + 1) & ". Begriff ein, um den Sie den " & (i + 1) & ". Suchbegriff ergänzen möchten:"
    'Stil3 = vbOKOnly
    Titel3 = "Eingabe der Bezugsbezeichnung"

    Mldg4 = "Suchen & Ergänzen kann nicht stattfinden, weil kein Suchbegriff / keine Ergänzung eingegeben wurde."
    Stil4 = vbRetryCancel
    Titel4 = "Suchen & Ergänzen fehlgeschlagen!"

    Mldg5 = "Suchen & Ergänzen durch Abbruch beendet."
    Stil5 = vbInformation
    Titel5 = "Suchen & Ergänzen abgebrochen!"

    Mldg6 = "Datei bezug.txt vorhanden. Fortfahren?"
    Stil6 = vbYesNo
    Titel6 = "Dateiprüfung"

    Mldg7 = "Datei bezug.txt nicht vorhanden. Fortfahren mit manueller Eingabe?"
    Stil7 = vbYesNo
    Titel7 = "Dateiprüfung"

    Mldg8 = "Trennzeichen (@) in Datei bezug.txt fehlt! Vorgang wird abgebrochen!"
    Stil8 = vbCritical
    Titel8 = "Trennzeichenprüfung"

    Mldg9 = "Sie können die Bezugsbezeichnungen per Datei oder manuell einfügen." & vbCrLf & vbCrLf & _
    "Die Datei muss auf dem Desktop mit dem Dateinamen bezug.txt angelegt werden und muss folgende zeilenweise Syntax aufweisen:" & vbCrLf & vbCrLf & _
    "Suchbegriff1@Bezugsbezeichnung1" & vbCrLf & "Suchbegriff2@Bezugsbezeichnung2" & vbCrLf & "..." & vbCrLf & vbCrLf & _
    "Bitte achten Sie bei den Suchbegriffen auf Groß-/Kleinschreibung sowie Ein-/Mehrzahl. Es wird nur nach ganzen Wörtern gesucht." & vbCrLf & "Die Bezugsbezeichnung (nur Zahl, ohne Suchbegriff) wird beim Einfügen automatisch in Klammern gesetzt." & vbCrLf & vbCrLf & "Die Datei bezug.txt wird abschließend geleert."
    Stil9 = vbInformation
    Titel9 = "Erklärung"

    Mldg10 = "Datei bezug.txt ist leer. Fortfahren mit manueller Eingabe?"
    Stil10 = vbYesNo
    Titel10 = "Inhaltsprüfung"

    Mldg11 = "Möchten Sie weitere Bezugsbezeichnungen einfügen?"
    Stil11 = vbYesNo
    Titel11 = "Wiederholung"

    Mldg12 = "Es ist kein Dokument geöffnet."
    Stil12 = vbOKOnly + vbCritical
    Titel12 = "Suchen & Ergänzen fehlgeschlagen!"

    Mldg13 = "Suchen & Ergänzen im leeren Dokument nicht möglich."
    Stil13 = vbOKOnly + vbCritical
    Titel13 = "Suchen & Ergänzen fehlgeschlagen!"

    'Abfrage, ob ein Dokument geöffnet ist
    If Documents.Count >= 1 Then

        'Anzahl der Zeichen, Shapes oder InlineShapes feststellen
        effC = ActiveDocument.BuiltInDocumentProperties(wdPropertyCharsWSpaces)
        effS = ActiveDocument.Shapes.Count
        effI = ActiveDocument.InlineShapes.Count

        'Abfrage, ob Dokument Zeichen, Shapes oder InlineShapes enthält (außer Kopf- und Fußzeile)
        If effC >= 1 Or effS >= 1 Or effI >= 1 Then

            'Nachverfolgungseinstellungen sichern
            trackrev = ActiveDocument.TrackRevisions
            markup = ActiveWindow.View.MarkupMode

            'Änderung 19.01.2017: prüfen, ob nicht angenommene Änderungen eines anderen Benutzers vorhanden sind
            o = ActiveDocument.Revisions.Count
            For p = 1 To o
                If ActiveDocument.Revisions.Count <> 0 And ActiveDocument.Revisions(p).Author <> Application.UserName Then

                    MsgBox "ACHTUNG:" & vbCrLf & vbCrLf & "Nicht angenommene Änderungen eines anderen Benutzers (" & ActiveDocument.Revisions(p).Author & ") vorhanden - nachträgliche Ergänzungen beeinflussen diese Änderungen!" & vbCrLf & vbCrLf & "Bitte anschließend prüfen!", vbOKOnly + vbExclamation, "Suchen & Ergänzen"
                    GoTo Weiter

                End If
            Next

Weiter:
            'wenn 'Änderungen verfolgen' deaktiviert ist -> aktivieren
            If ActiveDocument.TrackRevisions = False Then
                ActiveDocument.TrackRevisions = True
            End If

            'Sprechblasen-Einstellung auf balloon umstellen
            ActiveWindow.View.MarkupMode = wdBalloonRevisions

            'Erklärung
            Antwort9 = MsgBox(Mldg9, Stil9, Titel9)

Dim myPath As String
    myPath = "\\brefile11.esp.dom\citrixprofiles$\msg\Desktop\bezug.txt"
    If Dir(myPath) <> "" Then
        Antwort6 = MsgBox(Mldg6, Stil6, Titel6)
    Else
        Antwort7 = MsgBox(Mldg7, Stil7, Titel7)
    End If

Open myPath For Input As #1

                    'ReDim ar(0)
                    Do While Not EOF(1)
                        'Änderung 19.01.2017: Befehl "Line" hinzugefügt (Zeile samt Komma als String einlesen), damit mehrere kommagetrennte Bezugszeichen benutzt werden können
                        Line Input #1, strVariable
                        i = i + 1: ReDim Preserve ar(i)
                        ar(i) = strVariable

                        'Prüfung, ob Trennzeichen vorhanden
                        If InStr(strVariable, "@") = 0 Then
                            Antwort8 = MsgBox(Mldg8, Stil8, Titel8)
                            Close #1
                            GoTo Ende
                        End If

                        'Prüfung des Suchstrings
                        suchstring = Left(ar(i), InStr(ar(i), "@") - 1)
                        l = l + 1: ReDim Preserve splitar(l)
                        splitar(l) = suchstring
                        'Suchen
                        Selection.HomeKey Unit:=wdStory
                        Selection.Find.ClearFormatting
                        With Selection.Find
                            .Text = suchstring
                            .Forward = True
                            .Wrap = wdFindContinue
                            .Format = False
                            .MatchCase = True
                            .MatchWholeWord = True
                            .MatchAllWordForms = True
                            .MatchSoundsLike = False
                            .MatchWildcards = False
                        End With
                        Selection.Find.Execute

                        If Selection.Find.Found = False Then

                            Antwort1 = MsgBox("Suchen & Ergänzen kann nicht stattfinden, weil der Suchbegriff " & Chr(34) & suchstring & Chr(34) & " nicht gefunden werden konnte.", vbOKOnly + vbCritical, "Suchen & Ergänzen fehlgeschlagen!")
                            Close #1
                            GoTo Ende

                        End If
                    Loop
                    Close #1

                    'Strings im Array miteinander vergleichen
                    For m = 1 To i
                        For n = 1 To i
                            If Not (m = n) Then 'kein Objekt mit sich selbst vergleichen
                                If InStr(splitar(m), splitar(n)) > 0 Then
                                      MsgBox "Der Suchbegriff " & Chr(34) & splitar(n) & Chr(34) & " wird im Suchbegriff " & Chr(34) & splitar(m) & Chr(34) & " wiederverwendet!", vbOKOnly + vbCritical, "Suchen & Ergänzen fehlgeschlagen!"
                                      GoTo Ende
                                Else
                                    If InStr(splitar(n), splitar(m)) > 0 Then
                                      MsgBox "Der Suchbegriff " & Chr(34) & splitar(m) & Chr(34) & " wird im Suchbegriff " & Chr(34) & splitar(n) & Chr(34) & " wiederverwendet!", vbOKOnly + vbCritical, "Suchen & Ergänzen fehlgeschlagen!"
                                      GoTo Ende
                                    End If
                                End If
                            End If
                        Next
                    Next

                    'Datei leer?
                    If strVariable <> "" Then
                        GoTo SuchErsetz
                    Else
                    Antwort10 = MsgBox(Mldg10, Stil10, Titel10)
                        If Antwort10 = vbYes Then
                            GoTo Sucheingabe
                        Else
                            GoTo Ende
                        End If
                    End If




            Select Case Antwort7
                Case vbYes
                    GoTo Sucheingabe

                Case vbNo
                    GoTo Ende

            End Select

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
                .Text = suchstring
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

                Antwort1 = MsgBox("Suchen & Ergänzen kann nicht stattfinden, weil der Suchbegriff " & Chr(34) & suchstring & Chr(34) & " nicht gefunden werden konnte.", vbOKOnly + vbCritical, "Suchen & Ergänzen fehlgeschlagen!")
                GoTo Ende

            End If

Ergaenzeingabe:

            'Ergaenzstring
            ergaenzstring = InputBox("Bitte geben Sie die " & (i + 1) & ". Bezugsbezeichnung ein, um die Sie den " & (i + 1) & ". Suchbegriff ergänzen möchten:", Titel3)

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

            'Eingabe für weitere Strings wiederholen?
            If Antwort11 = vbYes Then
                GoTo Sucheingabe
            End If

SuchErsetz:

            'Array feldweise auswerten
            For k = 1 To i
            suchstring = Left(ar(k), InStr(ar(k), "@") - 1)
            ergaenzstring = Right(ar(k), Len(ar(k)) - InStr(ar(k), "@"))

                'Suchen und Ersetzen
                Selection.HomeKey Unit:=wdStory
                Selection.Find.ClearFormatting
                With Selection.Find
                    .Text = suchstring
                    .replacement.Text = suchstring & " (" & ergaenzstring & ")"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .MatchCase = True
                    .MatchWholeWord = True
                    .MatchAllWordForms = True
                    .MatchSoundsLike = False
                    .MatchWildcards = False
                End With
                Selection.Find.Execute

                If Selection.Find.Found = True Then

                    Selection.Find.Execute Replace:=wdReplaceAll

                Else

                    Antwort1 = MsgBox(Mldg1, Stil1, Titel1)
                    GoTo Ende

                End If
            Next

            'Array feldweise auswerten
            For k = 1 To i
            suchstring = Left(ar(k), InStr(ar(k), "@") - 1)
            ergaenzstring = Right(ar(k), Len(ar(k)) - InStr(ar(k), "@"))

                Selection.HomeKey Unit:=wdStory

                'Sprechblasen-Einstellung auf inline umstellen
                ActiveWindow.View.MarkupMode = wdInLineRevisions

                'Änderungen verfolgen deaktivieren
                ActiveDocument.TrackRevisions = False

                'Suchen und Ersetzen
                Selection.HomeKey Unit:=wdStory
                Selection.Find.ClearFormatting
                With Selection.Find
                    .Text = suchstring & suchstring
                    .replacement.Text = suchstring
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .MatchCase = True
                    .MatchWholeWord = True
                    .MatchAllWordForms = True
                    .MatchSoundsLike = False
                    .MatchWildcards = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
            Next

Leeren:

            If Antwort6 = vbYes Then
            'MsgBox "Datei bezug.txt wird geleert."
            'Datei leeren
            Open Server & user & Pfad & "bezug.txt" For Output As #1
            Close #1
            End If

Ende:
            'Nachverfolgungseinstellungen wiederherstellen
            ActiveWindow.View.MarkupMode = markup
            ActiveDocument.TrackRevisions = trackrev

            Selection.HomeKey Unit:=wdStory

            'Suchparameter zurücksetzen
            With Selection.Find
               .ClearFormatting
               .replacement.ClearFormatting
               .Text = ""
               .replacement.Text = ""
               .Forward = True
               .Wrap = wdFindContinue
               .Format = False
               .MatchCase = False
               .MatchWholeWord = False
               .MatchWildcards = False
               .MatchSoundsLike = False
               .MatchAllWordForms = False
            End With
        'sonst keine Zeichen, Shapes oder InlineShapes
        Else
            'Wenn leeres Dokument geöffnet ist -> Fehlermeldung
            Antwort13 = MsgBox(Mldg13, Stil13, Titel13)
        End If
    Else
        'Wenn kein Dokument geöffnet ist, Makro nicht ausführbar -> Fehlermeldung
        Antwort12 = MsgBox(Mldg12, Stil12, Titel12)
    End If
End Sub