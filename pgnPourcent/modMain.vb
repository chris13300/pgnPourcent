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
    Public nbIgnorees() As Long
    Public nbCompteurs(,) As Long
    Public nbVictoires(,) As Single
    Public tabCoups(20) As String

    Public depart As Date
    Public ecouleMem As Long

    Sub Main()
        Dim chaine As String, indexFichier As Integer, indexThread As Integer, reponse As String, victoire As Single
        Dim ecoule As Long, tabEcoule() As String, message As String, i As Integer, tabVisites(20) As Long, tabChaines(20) As String, tabVictoires(20) As Single

        If My.Computer.FileSystem.GetFileInfo(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) & "\Documents\Visual Studio 2013\Projects\pgnPourcent\pgnPourcent\bin\x64\Debug\pgnPourcent.exe").LastWriteTime > My.Computer.FileSystem.GetFileInfo(My.Application.Info.AssemblyName & ".exe").LastWriteTime Then
            MsgBox("Il existe une version plus récente de ce programme !", MsgBoxStyle.Information)
            End
        End If

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
        ReDim nbIgnorees(nbFichiers)
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
        Console.WriteLine(Trim(Format(nbParties(0) / ecoule, "# ### ##0 games/sec")) & " (" & Trim(Format(nbParties(0) + nbIgnorees(0), "# ### ##0")) & " games with " & Trim(Format(nbIgnorees(0), "# ### ##0")) & " ignored) :")

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
        Console.WriteLine(Trim(Format(nbParties(0) / ecoule, "# ### ##0 games/sec")) & " (" & Trim(Format(nbParties(0) + nbIgnorees(0), "# ### ##0")) & " games with " & Trim(Format(nbIgnorees(0), "# ### ##0")) & " ignored) :")

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

End Module
