Imports System.Management 'ajouter référence system.management
Imports System.Threading

Module modMain
    Public tabThread() As Thread
    Public nbThreadActif As Integer
    Public nbTaches As Integer
    Public maxThreads As Integer

    Public tabFichiers() As String

    Public nbFichiers As Integer
    Public nbFichiersTraites As Integer

    Public nbParties() As Long
    Public nbCompteurs(,) As Long
    Public nbVictoires(,) As Single
    Public tabCoups(20) As String

    Public depart As Date
    Public ecouleMem As Long

    Sub Main()
        Dim chaine As String, indexFichier As Integer, indexThread As Integer, reponse As String, victoire As Single
        Dim ecoule As Long, tabEcoule() As String, message As String, i As Integer, tabVisites(20) As Long, tabChaines(20) As String, tabVictoires(20) As Single

        tabCoups(1) = "a3 "
        tabCoups(2) = "a4 "
        tabCoups(3) = "b3 "
        tabCoups(4) = "b4 "
        tabCoups(5) = "c3 "
        tabCoups(6) = "c4 "
        tabCoups(7) = "d3 "
        tabCoups(8) = "d4 "
        tabCoups(9) = "e3 "
        tabCoups(10) = "e4 "
        tabCoups(11) = "f3 "
        tabCoups(12) = "f4 "
        tabCoups(13) = "g3 "
        tabCoups(14) = "g4 "
        tabCoups(15) = "h3 "
        tabCoups(16) = "h4 "
        tabCoups(17) = "Na3"
        tabCoups(18) = "Nc3"
        tabCoups(19) = "Nf3"
        tabCoups(20) = "Nh3"

        maxThreads = cpu()
        nbTaches = maxThreads
        nbThreadActif = 0
        nbFichiers = 0
        chaine = ""
        ecouleMem = 0

        nbFichiersTraites = 0

        For Each pgn In My.Computer.FileSystem.GetFiles(My.Application.Info.DirectoryPath, FileIO.SearchOption.SearchAllSubDirectories, "*.pgn")
            nbFichiers = nbFichiers + 1
            chaine = chaine & pgn & vbCrLf
        Next

        If nbFichiers = 0 Then
            End
        End If

        reponse = InputBox("How many files to read at the same time ?", nbFichiers & " files to read", Format(nbTaches))
        If reponse <> "" Then
            Try
                nbTaches = CInt(reponse)
                's'il y a moins de fichiers que de cores/threads, on corrige la demande de l'utilisateur
                'demande 15 taches, max 12 cores/threads, 1 fichier => max 12 cores/threads pour ce fichier
                'demande 15 taches, max 12 cores/threads, 6 fichiers => max 2 cores/threads par fichier
                'demande 15 taches, max 12 cores/threads, 12 fichiers => max 1 cores/threads par fichier
                If maxThreads < nbTaches And nbFichiers <= maxThreads Then
                    nbTaches = maxThreads
                End If
            Catch ex As Exception
                End
            End Try
            If nbTaches < 1 Then
                End
            End If
        Else
            End
        End If

        ReDim tabThread(nbTaches)

        tabFichiers = Split(vbCrLf & chaine, vbCrLf)
        ReDim nbParties(nbFichiers)
        ReDim nbCompteurs(nbFichiers, 20)
        ReDim nbVictoires(nbFichiers, 20)

        indexThread = 1
        For indexFichier = 1 To nbFichiers
            'si on est déjà au taquet
            While nbThreadActif >= nbTaches
                'on attend qu'une tache se libère
                Threading.Thread.Sleep(1000)
                affichage()
            End While

            If indexFichier > nbTaches Then
                While tabThread(indexThread).IsAlive
                    indexThread = indexThread + 1
                    If indexThread > UBound(tabThread) Then
                        indexThread = 1
                    End If
                    Threading.Thread.Sleep(1000)
                End While
            End If

            'exécution
            nbThreadActif = nbThreadActif + 1
            tabThread(indexThread) = New Thread(AddressOf pourcentage)
            tabThread(indexThread).Start(Format(indexFichier & ":" & indexThread))
            indexThread = indexThread + 1
            If indexThread > UBound(tabThread) Then
                indexThread = 1
            End If
        Next

        Do
            Threading.Thread.Sleep(1000)
            affichage()
        Loop While nbThreadActif > 0

        ecoule = DateDiff(DateInterval.Second, depart, Now)

        sommeParties()
        sommeCompteurs()

        Console.Clear()
        tabEcoule = Split(secJHMS(ecoule), ";")
        Console.Title = My.Computer.Name & " (" & Format(nbFichiersTraites / nbFichiers, "0.00%") & ") : " & tabEcoule(0) & "days " & tabEcoule(1) & "hrs " & tabEcoule(2) & "min " & tabEcoule(3) & "sec"
        Console.WriteLine(Trim(Format(nbParties(0) / ecoule, "# ### ##0 games/sec")) & " (" & Trim(Format(nbParties(0), "# ### ##0")) & ") :")

        For i = 1 To 20
            tabVisites(i) = nbCompteurs(0, i)
            tabVictoires(i) = nbVictoires(0, i)
            tabChaines(i) = tabCoups(i)
        Next

        For i = 1 To 20
            For j = 1 To 20
                If tabVisites(i) > tabVisites(j) Then
                    ecoule = tabVisites(j)
                    tabVisites(j) = tabVisites(i)
                    tabVisites(i) = ecoule

                    victoire = tabVictoires(j)
                    tabVictoires(j) = tabVictoires(i)
                    tabVictoires(i) = victoire

                    message = tabChaines(j)
                    tabChaines(j) = tabChaines(i)
                    tabChaines(i) = message
                End If
            Next
        Next

        Console.WriteLine()

        message = StrDup(5, "-") & "-|-" & StrDup(4, "-") & "-|--" & StrDup(9, "-") & "-|-" & StrDup(6, "-") & "-|-" & StrDup(7, "-") & vbCrLf
        message = message & " Rank | Move |   Played   |  Rate  |   Win" & vbCrLf
        message = message & StrDup(5, "-") & "-|-" & StrDup(4, "-") & "-|--" & StrDup(9, "-") & "-|-" & StrDup(6, "-") & "-|-" & StrDup(7, "-") & vbCrLf
        For i = 1 To 20
            If tabVisites(i) > 0 Then
                message = message & "  " & Replace("#" & Format(i, "00"), "#0", " #") & " |  " & tabChaines(i) & " | " & StrDup(10 - Len(Format(tabVisites(i), "## ### ##0")), " ") & Format(tabVisites(i), "## ### ##0") & " | " & StrDup(6 - Len(Format(tabVisites(i) / nbParties(0), "0.0%")), " ") & Format(tabVisites(i) / nbParties(0), "0.0%") & " | " & StrDup(6 - Len(Format(tabVictoires(i) / tabVisites(i), "0.0%")), " ") & Format(tabVictoires(i) / tabVisites(i), "0.0%") & vbCrLf
                message = message & StrDup(5, "-") & "-|-" & StrDup(4, "-") & "-|--" & StrDup(9, "-") & "-|-" & StrDup(6, "-") & "-|-" & StrDup(7, "-") & vbCrLf
            End If
        Next

        Console.WriteLine(message)

        Console.WriteLine("Press ENTER to close this window.")
        Console.ReadLine()

    End Sub

    Public Sub affichage()
        Dim ecoule As Long, tabEcoule() As String, message As String, i As Integer, tabVisites(20) As Long, tabChaines(20) As String, tabVictoires(20) As Single, victoire As Single

        ecoule = 0
        If Not depart = Nothing Then
            ecoule = DateDiff(DateInterval.Second, depart, Now)
        End If

        If (ecoule - ecouleMem) < 5 And ecoule > ecouleMem Then
            Exit Sub
        End If
        ecouleMem = ecoule

        sommeParties()
        sommeCompteurs()

        Console.Clear()
        tabEcoule = Split(secJHMS(ecoule), ";")
        Console.Title = My.Computer.Name & " (" & Format(nbFichiersTraites / nbFichiers, "0.00%") & ") : " & tabEcoule(0) & "days " & tabEcoule(1) & "hrs " & tabEcoule(2) & "min " & tabEcoule(3) & "sec"
        Console.WriteLine(Trim(Format(nbParties(0) / ecoule, "# ### ##0 games/sec")) & " (" & Trim(Format(nbParties(0), "# ### ##0")) & ") :")

        For i = 1 To 20
            tabVisites(i) = nbCompteurs(0, i)
            tabVictoires(i) = nbVictoires(0, i)
            tabChaines(i) = tabCoups(i)
        Next

        For i = 1 To 20
            For j = 1 To 20
                If tabVisites(i) > tabVisites(j) Then
                    ecoule = tabVisites(j)
                    tabVisites(j) = tabVisites(i)
                    tabVisites(i) = ecoule

                    victoire = tabVictoires(j)
                    tabVictoires(j) = tabVictoires(i)
                    tabVictoires(i) = victoire

                    message = tabChaines(j)
                    tabChaines(j) = tabChaines(i)
                    tabChaines(i) = message
                End If
            Next
        Next

        Console.WriteLine()

        message = StrDup(5, "-") & "-|-" & StrDup(4, "-") & "-|--" & StrDup(9, "-") & "-|-" & StrDup(6, "-") & "-|-" & StrDup(7, "-") & vbCrLf
        message = message & " Rank | Move |   Played   |  Rate  |   Win" & vbCrLf
        message = message & StrDup(5, "-") & "-|-" & StrDup(4, "-") & "-|--" & StrDup(9, "-") & "-|-" & StrDup(6, "-") & "-|-" & StrDup(7, "-") & vbCrLf
        For i = 1 To 20
            If tabVisites(i) > 0 Then
                message = message & "  " & Replace("#" & Format(i, "00"), "#0", " #") & " |  " & tabChaines(i) & " | " & StrDup(10 - Len(Format(tabVisites(i), "## ### ##0")), " ") & Format(tabVisites(i), "## ### ##0") & " | " & StrDup(6 - Len(Format(tabVisites(i) / nbParties(0), "0.0%")), " ") & Format(tabVisites(i) / nbParties(0), "0.0%") & " | " & StrDup(6 - Len(Format(tabVictoires(i) / tabVisites(i), "0.0%")), " ") & Format(tabVictoires(i) / tabVisites(i), "0.0%") & vbCrLf
                message = message & StrDup(5, "-") & "-|-" & StrDup(4, "-") & "-|--" & StrDup(9, "-") & "-|-" & StrDup(6, "-") & "-|-" & StrDup(7, "-") & vbCrLf
            End If
        Next

        Console.WriteLine(message)
    End Sub

    Public Function cpu(Optional reel As Boolean = False) As Integer
        Dim collection As New ManagementObjectSearcher("select * from Win32_Processor"), taches As Integer
        taches = 0

        For Each element As ManagementObject In collection.Get
            If reel Then
                taches = taches + element.Properties("NumberOfCores").Value 'cores
            Else
                taches = taches + element.Properties("NumberOfLogicalProcessors").Value 'threads
            End If
        Next

        Return taches
    End Function

    Public Sub pourcentage(indexes As String)
        Dim tabChaine() As String, indexfichier As Integer, indexThread As Integer
        Dim fichierPGN As String, ligne As String, lecture As System.IO.TextReader, tailleFichier As Long, cumul As Long
        Dim trouve As Boolean, victoire As Single

        tabChaine = Split(indexes, ":")
        indexfichier = Int(tabChaine(0))
        indexThread = Int(tabChaine(1))

        fichierPGN = tabFichiers(indexfichier)
        If My.Computer.FileSystem.FileExists(fichierPGN) Then
            If depart = Nothing Then
                depart = Now()
            End If

            tailleFichier = FileLen(fichierPGN)
            lecture = My.Computer.FileSystem.OpenTextFileReader(fichierPGN)

            nbParties(indexfichier) = 0
            trouve = False
            victoire = 0.5
            cumul = 0
            'ligne par ligne
            Do
                ligne = lecture.ReadLine()
                cumul = cumul + Len(ligne) + 2
                If InStr(ligne, "[Fen ", CompareMethod.Text) > 0 Then
                    trouve = True
                ElseIf InStr(ligne, "[Event ", CompareMethod.Text) > 0 Then
                    trouve = False
                    victoire = 0.5
                ElseIf InStr(ligne, "[Result ", CompareMethod.Text) > 0 Then
                    If InStr(ligne, "1-0") > 0 Then
                        victoire = 1
                    ElseIf InStr(ligne, "0-1") > 0 Then
                        victoire = 0
                    End If
                ElseIf Not trouve Then
                    If ligne <> "" Then
                        If InStr(ligne, "[", CompareMethod.Text) = 0 And InStr(ligne, "]", CompareMethod.Text) = 0 And InStr(ligne, """", CompareMethod.Text) = 0 Then
                            ligne = Replace(ligne, ". ", ".")
                            ligne = ligne.Substring(0, ligne.IndexOf(" "))
                            If InStr(ligne, ".") > 0 Then
                                ligne = ligne.Substring(ligne.IndexOf(".") + 1)
                            End If
                            Select Case LCase(ligne)
                                Case "a3", "a2a3", "a2-a3"
                                    nbCompteurs(indexfichier, 1) = nbCompteurs(indexfichier, 1) + 1
                                    nbVictoires(indexfichier, 1) = nbVictoires(indexfichier, 1) + victoire

                                Case "a4", "a2a4", "a2-a4"
                                    nbCompteurs(indexfichier, 2) = nbCompteurs(indexfichier, 2) + 1
                                    nbVictoires(indexfichier, 2) = nbVictoires(indexfichier, 2) + victoire

                                Case "b3", "b2b3", "b2-b3"
                                    nbCompteurs(indexfichier, 3) = nbCompteurs(indexfichier, 3) + 1
                                    nbVictoires(indexfichier, 3) = nbVictoires(indexfichier, 3) + victoire

                                Case "b4", "b2b4", "b2-b4"
                                    nbCompteurs(indexfichier, 4) = nbCompteurs(indexfichier, 4) + 1
                                    nbVictoires(indexfichier, 4) = nbVictoires(indexfichier, 4) + victoire

                                Case "c3", "c2c3", "c2-c3"
                                    nbCompteurs(indexfichier, 5) = nbCompteurs(indexfichier, 5) + 1
                                    nbVictoires(indexfichier, 5) = nbVictoires(indexfichier, 5) + victoire

                                Case "c4", "c2c4", "c2-c4"
                                    nbCompteurs(indexfichier, 6) = nbCompteurs(indexfichier, 6) + 1
                                    nbVictoires(indexfichier, 6) = nbVictoires(indexfichier, 6) + victoire

                                Case "d3", "d2d3", "d2-d3"
                                    nbCompteurs(indexfichier, 7) = nbCompteurs(indexfichier, 7) + 1
                                    nbVictoires(indexfichier, 7) = nbVictoires(indexfichier, 7) + victoire

                                Case "d4", "d2d4", "d2-d4"
                                    nbCompteurs(indexfichier, 8) = nbCompteurs(indexfichier, 8) + 1
                                    nbVictoires(indexfichier, 8) = nbVictoires(indexfichier, 8) + victoire

                                Case "e3", "e2e3", "e2-e3"
                                    nbCompteurs(indexfichier, 9) = nbCompteurs(indexfichier, 9) + 1
                                    nbVictoires(indexfichier, 9) = nbVictoires(indexfichier, 9) + victoire

                                Case "e4", "e2e4", "e2-e4"
                                    nbCompteurs(indexfichier, 10) = nbCompteurs(indexfichier, 10) + 1
                                    nbVictoires(indexfichier, 10) = nbVictoires(indexfichier, 10) + victoire

                                Case "f3", "f2f3", "f2-f3"
                                    nbCompteurs(indexfichier, 11) = nbCompteurs(indexfichier, 11) + 1
                                    nbVictoires(indexfichier, 11) = nbVictoires(indexfichier, 11) + victoire

                                Case "f4", "f2f4", "f2-f4"
                                    nbCompteurs(indexfichier, 12) = nbCompteurs(indexfichier, 12) + 1
                                    nbVictoires(indexfichier, 12) = nbVictoires(indexfichier, 12) + victoire

                                Case "g3", "g2g3", "g2-g3"
                                    nbCompteurs(indexfichier, 13) = nbCompteurs(indexfichier, 13) + 1
                                    nbVictoires(indexfichier, 13) = nbVictoires(indexfichier, 13) + victoire

                                Case "g4", "g2g4", "g2-g4"
                                    nbCompteurs(indexfichier, 14) = nbCompteurs(indexfichier, 14) + 1
                                    nbVictoires(indexfichier, 14) = nbVictoires(indexfichier, 14) + victoire

                                Case "h3", "h2h3", "h2-h3"
                                    nbCompteurs(indexfichier, 15) = nbCompteurs(indexfichier, 15) + 1
                                    nbVictoires(indexfichier, 15) = nbVictoires(indexfichier, 15) + victoire

                                Case "h4", "h2h4", "h2-h4"
                                    nbCompteurs(indexfichier, 16) = nbCompteurs(indexfichier, 16) + 1
                                    nbVictoires(indexfichier, 16) = nbVictoires(indexfichier, 16) + victoire

                                Case "na3", "b1a3", "b1-a3", "nb1a3", "nb1-a3"
                                    nbCompteurs(indexfichier, 17) = nbCompteurs(indexfichier, 17) + 1
                                    nbVictoires(indexfichier, 17) = nbVictoires(indexfichier, 17) + victoire

                                Case "nc3", "b1c3", "b1-c3", "nb1c3", "nb1-c3"
                                    nbCompteurs(indexfichier, 18) = nbCompteurs(indexfichier, 18) + 1
                                    nbVictoires(indexfichier, 18) = nbVictoires(indexfichier, 18) + victoire

                                Case "nf3", "g1f3", "g1-f3", "ng1f3", "ng1-f3"
                                    nbCompteurs(indexfichier, 19) = nbCompteurs(indexfichier, 19) + 1
                                    nbVictoires(indexfichier, 19) = nbVictoires(indexfichier, 19) + victoire

                                Case "nh3", "g1h3", "g1-h3", "ng1h3", "ng1-h3"
                                    nbCompteurs(indexfichier, 20) = nbCompteurs(indexfichier, 20) + 1
                                    nbVictoires(indexfichier, 20) = nbVictoires(indexfichier, 20) + victoire

                                Case Else
                                    MsgBox("en travaux")

                            End Select
                            trouve = True
                            nbParties(indexfichier) = nbParties(indexfichier) + 1
                        End If
                    End If
                End If
            Loop Until ligne Is Nothing
            lecture.Close()

            nbFichiersTraites = nbFichiersTraites + 1
        End If

        lecture = Nothing

        If nbThreadActif > 0 Then
            nbThreadActif = nbThreadActif - 1
        End If

        tabThread(indexThread).Abort()

    End Sub

    Public Function secJHMS(ByVal secondes As Long) As String
        Dim restant As Integer, valeur As Integer, chaine As String

        restant = secondes

        valeur = Fix(restant / 60 / 60 / 24) 'jours
        chaine = valeur
        restant = restant - valeur * 60 * 60 * 24

        valeur = Fix(restant / 60 / 60) 'heures
        chaine = chaine & ";" & valeur
        restant = restant - valeur * 60 * 60

        valeur = Fix(restant / 60) 'minutes
        chaine = chaine & ";" & valeur
        restant = restant - valeur * 60

        valeur = Fix(restant) 'secondes
        chaine = chaine & ";" & valeur

        'ex : secondes = 34849 => secJHMS = "0;9;40;49" ("Jours;Heures;Minutes;Secondes")

        Return chaine
    End Function

    Public Sub sommeCompteurs()
        Dim i As Integer, j As Integer

        For j = 1 To 20
            nbCompteurs(0, j) = 0
            nbVictoires(0, j) = 0
            For i = 1 To nbFichiers
                nbCompteurs(0, j) = nbCompteurs(0, j) + nbCompteurs(i, j)
                nbVictoires(0, j) = nbVictoires(0, j) + nbVictoires(i, j)
            Next
        Next

    End Sub

    Public Sub sommeParties()
        Dim i As Integer

        nbParties(0) = 0
        For i = 1 To nbFichiers
            nbParties(0) = nbParties(0) + nbParties(i)
        Next

    End Sub

End Module
