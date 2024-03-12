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
                    .Replacement.Text = suchstring & " (" & ergaenzstring & ")"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .MatchCase = True
                    .MatchWholeWord = True

                    'setting this to true means, that the program considers now all word related matches
                    'like plurals (does not REPLACE the whole thing lol)
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

            'This loop checks for multiple occurrence of suchstring right after each other and replaces it with a single instance of suchstring
            For k = 1 To i
            suchstring = Left(ar(k), InStr(ar(k), "@") - 1)
            ergaenzstring = Right(ar(k), Len(ar(k)) - InStr(ar(k), "@"))
                Selection.HomeKey Unit:=wdStory
                'Sprechblasen-Einstellung auf inline umstellen
                ActiveWindow.View.MarkupMode = wdInLineRevisions
                'Änderungen verfolgen deaktivieren
                ActiveDocument.TrackRevisions = False
                'moves cursor for reading document to the very beginning again
                Selection.HomeKey Unit:=wdStory
                Selection.Find.ClearFormatting
                With Selection.Find
                    .Text = suchstring & suchstring
                    .Replacement.Text = suchstring
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .MatchCase = True
                    .MatchWholeWord = True
                    .MatchAllWordForms = False
                    .MatchSoundsLike = False
                    .MatchWildcards = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
            Next