Imports System.Management 'ajouter référence system.management

Module modFonctions

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
            nbIgnorees(indexfichier) = 0
            trouve = False
            victoire = 0.5
            cumul = 0
            'ligne par ligne
            Do
                ligne = lecture.ReadLine()
                cumul = cumul + Len(ligne) + 2
                If InStr(ligne, "[Fen ", CompareMethod.Text) > 0 Then
                    trouve = True
                    nbIgnorees(indexfichier) = nbIgnorees(indexfichier) + 1
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
        nbIgnorees(0) = 0
        For i = 1 To nbFichiers
            nbParties(0) = nbParties(0) + nbParties(i)
            nbIgnorees(0) = nbIgnorees(0) + nbIgnorees(i)
        Next

    End Sub

End Module
