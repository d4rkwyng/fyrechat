Attribute VB_Name = "modStatstring"
Option Explicit

Public Sub sprintf(ByRef Source As String, ByVal nText As String, _
    Optional ByVal A As Variant, _
    Optional ByVal B As Variant, _
    Optional ByVal c As Variant, _
    Optional ByVal d As Variant, _
    Optional ByVal E As Variant, _
    Optional ByVal F As Variant, _
    Optional ByVal G As Variant, _
    Optional ByVal H As Variant)
On Local Error Resume Next
    
    nText = Replace(nText, "%S", "%s")
    Dim i As Byte
    i = 0
    Do While (InStr(1, nText, "%s") <> 0)
        Select Case i
            Case 0
                If IsEmpty(A) Then GoTo TheEnd
                nText = Replace(nText, "%s", A, 1, 1)
            Case 1
                If IsEmpty(B) Then GoTo TheEnd
                nText = Replace(nText, "%s", B, 1, 1)
            Case 2
                If IsEmpty(c) Then GoTo TheEnd
                nText = Replace(nText, "%s", c, 1, 1)
            Case 3
                If IsEmpty(d) Then GoTo TheEnd
                nText = Replace(nText, "%s", d, 1, 1)
            Case 4
                If IsEmpty(E) Then GoTo TheEnd
                nText = Replace(nText, "%s", E, 1, 1)
            Case 5
                If IsEmpty(F) Then GoTo TheEnd
                nText = Replace(nText, "%s", F, 1, 1)
            Case 6
                If IsEmpty(G) Then GoTo TheEnd
                nText = Replace(nText, "%s", G, 1, 1)
            Case 7
                If IsEmpty(H) Then GoTo TheEnd
                nText = Replace(nText, "%s", H, 1, 1)
        End Select
        i = i + 1
    Loop
TheEnd:
    Source = Source & nText
End Sub

Public Sub ParseStatString(ByVal statstring As String, ByRef outbuf As String)
On Local Error Resume Next
    Dim Values() As String
    Dim cType As String
    
    Select Case Left$(statstring, 4)
        Case "3RAW"
            sprintf outbuf, "WarCraft III: "
            If Len(statstring) > 4 Then
                Values = Split(statstring, " ")
                    strcpy outbuf, "Level " & Values(2)
                    Select Case Mid$(Values(1), 2, 1)
                        Case "H": strcpy outbuf, " Human, "
                        Case "O": strcpy outbuf, " Orc, "
                        Case "N": strcpy outbuf, " Night Elves, "
                        Case "U": strcpy outbuf, " Undead, "
                        Case "R": strcpy outbuf, " Random, "
                        Case "K": strcpy outbuf, " The Lost Viking, "
                        Case Else: strcpy outbuf, " Unknown, "
                    End Select
                    Select Case Mid$(Values(1), 1, 1)
                        Case 1: strcpy outbuf, "icon: " & GetIconWAR3(Mid$(Values(1), 1, 1), Mid$(Values(1), 2, 1))
                        Case 2: strcpy outbuf, "icon: " & GetIconWAR3(Mid$(Values(1), 1, 1), Mid$(Values(1), 2, 1))
                        Case 3: strcpy outbuf, "icon: " & GetIconWAR3(Mid$(Values(1), 1, 1), Mid$(Values(1), 2, 1))
                        Case 4: strcpy outbuf, "icon: " & GetIconWAR3(Mid$(Values(1), 1, 1), Mid$(Values(1), 2, 1))
                        Case 5: strcpy outbuf, "icon: " & GetIconWAR3(Mid$(Values(1), 1, 1), Mid$(Values(1), 2, 1))
                        Case Else: strcpy outbuf, Mid$(Values(1), 1, 1)
                    End Select

                Exit Sub
            ElseIf Len(statstring) = 4 Then
                strcpy outbuf, "No information available"
                Exit Sub
            Else
                strcpy outbuf, "Error: " & statstring
                Exit Sub
            End If
        Case "PX3W"
            sprintf outbuf, "WarCraft III The Frozen Throne: "
            If Len(statstring) > 4 Then
                Values = Split(statstring, " ")
                    strcpy outbuf, "Level " & Values(2)
                    Select Case Mid$(Values(1), 2, 1)
                        Case "H": strcpy outbuf, " Human, "
                        Case "O": strcpy outbuf, " Orc, "
                        Case "N": strcpy outbuf, " Night Elves, "
                        Case "U": strcpy outbuf, " Undead, "
                        Case "R": strcpy outbuf, " Random, "
                        Case Else: strcpy outbuf, " Unknown, "
                    End Select
                    Select Case Mid$(Values(1), 1, 1)
                        Case 1: strcpy outbuf, "icon: " & GetIconW3XP(Mid$(Values(1), 1, 1), Mid$(Values(1), 2, 1))
                        Case 2: strcpy outbuf, "icon: " & GetIconW3XP(Mid$(Values(1), 1, 1), Mid$(Values(1), 2, 1))
                        Case 3: strcpy outbuf, "icon: " & GetIconW3XP(Mid$(Values(1), 1, 1), Mid$(Values(1), 2, 1))
                        Case 4: strcpy outbuf, "icon: " & GetIconW3XP(Mid$(Values(1), 1, 1), Mid$(Values(1), 2, 1))
                        Case Else: strcpy outbuf, Mid$(Values(1), 1, 1)
                    End Select

                Exit Sub
            ElseIf Len(statstring) = 4 Then
                strcpy outbuf, "No information available"
                Exit Sub
            Else
                strcpy outbuf, "Error: " & statstring
                Exit Sub
            End If
        Case "RHSS"
            Call strcpy(outbuf, "StarCraft Shareware")
        Case "RATS"
            Values() = Split(Mid$(statstring, 6), " ")
            If UBound(Values) <> 8 Then
                Call sprintf(outbuf, "StarCraft%s Bot", IIf((Values(3) = 1), " Spawn ", ""))
                Exit Sub
            End If
            If Values(0) > 0 Then
                Call sprintf(outbuf, "StarCraft%s: %s wins and a ladder rating of %s", IIf((Values(3) = 1), " Spawn ", ""), Values(2), Values(0))
            Else
                Call sprintf(outbuf, "StarCraft%s: %s wins", IIf((Values(3) = 1), " Spawn", ""), Values(2))
            End If
        Case "PXES"
            Values() = Split(Mid(statstring, 6), " ")
            If UBound(Values) <> 8 Then
                Call sprintf(outbuf, "StarCraft Brood War%s Bot", IIf((Values(3) = 1), " Spawn", ""))
                Exit Sub
            End If
            If Values(0) > 0 Then
                Call sprintf(outbuf, "StarCraft Brood War%s: %s wins and a ladder rating of %s", IIf((Values(3) = 1), " Spawn", ""), Values(2), Values(0))
            Else
                Call sprintf(outbuf, "StarCraft Brood War%s: %s wins", IIf((Values(3) = 1), " Spawn", ""), Values(2))
            End If
        Case "RTSJ"
            Values() = Split(Mid(statstring, 6), " ")
            If UBound(Values) <> 8 Then
                Call sprintf(outbuf, "StarCraft Japanese%s Bot", IIf((Values(3) = 1), " Spawn", ""))
                Exit Sub
            End If
            If Values(0) > 0 Then
                Call sprintf(outbuf, "StarCraft Japanese%s: %s wins and a ladder rating of %s", IIf((Values(3) = 1), " Spawn", ""), Values(2), Values(0))
            Else
                Call sprintf(outbuf, "StarCraft Japanese%s: %s wins", IIf((Values(3) = 1), " Spawn", ""), Values(2))
            End If
        Case "NB2W"
            Values() = Split(Mid$(statstring, 6), " ")
            If UBound(Values) <> 8 Then
                Call sprintf(outbuf, "WarCraft II%s Bot", IIf((Values(3) = 1), " Spawn", ""))
                Exit Sub
            End If
            If Values(0) > 0 Then
                Call sprintf(outbuf, "WarCraft II%s: %s wins and a ladder rating of %s", IIf((Values(3) = 1), " Spawn", ""), Values(2), Values(0))
            Else
                Call sprintf(outbuf, "WarCraft II%s: %s wins", IIf((Values(3) = 1), " Spawn", ""), Values(2))
            End If
        Case "RHSD"
            Values() = Split(Mid$(statstring, 6), " ")
            If UBound(Values) <> 8 Then
                Call strcpy(outbuf, "Diablo Shareware Bot")
                Exit Sub
            End If
            Select Case Values(2)
                Case 0: cType = "warrior"
                Case 1: cType = "rogue"
                Case 2: cType = "sorceror"
            End Select
            Call sprintf(outbuf, "Diablo Shareware: a level %s %s with %s dots, %s strength, %s magic, %s dexterity, %s vitality, and %s gold", Values(0), cType, Values(1), Values(3), Values(4), Values(5), Values(6), Values(7))
        Case "LTRD"
            Values() = Split(Mid$(statstring, 6), " ")
            If UBound(Values) <> 8 Then
                Call strcpy(outbuf, "Diablo Bot")
                Exit Sub
            End If
            Select Case Values(2)
                Case 0: cType = "warrior"
                Case 1: cType = "rogue"
                Case 2: cType = "sorceror"
            End Select
            'Diablo: a level 7 with warrior dots, 2 strength, 34 magic, 35 dexterity, 20 vitality, and 30 gold with a ping of 3344.
            Call sprintf(outbuf, "Diablo: a level %s %s with %s dots, %s strength, %s magic, %s dexterity, %s vitality, and %s gold", Values(0), cType, Values(1), Values(3), Values(4), Values(5), Values(6), Values(7))
        Case "VD2D", "PX2D"
            Call strcpy(outbuf, ParseD2Stats(statstring))
        Case "TAHC"
            Call strcpy(outbuf, "Chat Client")
    End Select
End Sub

Public Function ParseD2Stats(ByVal stats As String)
On Local Error Resume Next
    Dim d2classes(0 To 7) As String
        d2classes(0) = "amazon"
        d2classes(1) = "sorceress"
        d2classes(2) = "necromancer"
        d2classes(3) = "paladin"
        d2classes(4) = "barbarian"
        d2classes(5) = "druid"
        d2classes(6) = "assassin"
        d2classes(7) = "unknown"
    Dim statbuf As String, p() As String, server As String, name As String
    If Len(stats) > 4 Then
        Dim sLen As Byte
        sLen = GetServer(stats, server)
        sLen = GetCharacterName(stats, sLen, name)
        Call MakeArray(Mid$(stats, sLen), p())
    End If
    If Left$(stats, 4) = "VD2D" Then
        Call strcpy(statbuf, "Diablo II: ")
    Else
        Call strcpy(statbuf, "Diablo II Lord of Destruction: ")
    End If
    If (Len(stats) = 4) Then
        Call strcpy(statbuf, "Open Character")
    Else
        Dim Version As Byte
        Version = Asc(p(0)) - &H80
        Dim charclass As Byte
        charclass = Asc(p(13)) - 1
        If (charclass < 0) Or (charclass > 6) Then
            charclass = 7
        End If
        Dim female As Boolean
        female = False
        If (charclass = 0) Or (charclass = 1) Or (charclass = 6) Then
            female = True
        End If
        Dim charlevel As Byte
        charlevel = Asc(p(25))
        Dim hardcore As Byte
        hardcore = Asc(p(26)) And 4
        Dim expansion As Boolean
        expansion = False
        If Left$(stats, 4) = "PX2D" Then
            If (Asc(p(26)) And &H20) Then
                Select Case RShift((Asc(p(27)) And &H18), 3)
                    Case 1
                        If hardcore Then
                            Call strcpy(statbuf, "Destroyer ")
                        Else
                            Call strcpy(statbuf, "Slayer ")
                        End If
                    Case 2
                        If hardcore Then
                            Call strcpy(statbuf, "Conquerer ")
                        Else
                            Call strcpy(statbuf, "Champion ")
                        End If
                    Case 3
                        If hardcore Then
                            Call strcpy(statbuf, "Guardian ")
                        Else
                            If Not female Then
                                Call strcpy(statbuf, "Patriarch ")
                            Else
                                Call strcpy(statbuf, "Matriarch ")
                            End If
                        End If
                End Select
                expansion = True
            End If
        End If
        If Not expansion Then
            Select Case RShift((Asc(p(27)) And &H18), 3)
                Case 1
                    If female = False Then
                        If hardcore Then
                            Call strcpy(statbuf, "Count ")
                        Else
                            Call strcpy(statbuf, "Sir ")
                        End If
                    Else
                        If hardcore Then
                            Call strcpy(statbuf, "Countess ")
                        Else
                            Call strcpy(statbuf, "Dame ")
                        End If
                    End If
                Case 2
                    If female = False Then
                        If hardcore Then
                            Call strcpy(statbuf, "Duke ")
                        Else
                            Call strcpy(statbuf, "Lord ")
                        End If
                    Else
                        If hardcore Then
                            Call strcpy(statbuf, "Duchess ")
                        Else
                            Call strcpy(statbuf, "Lady ")
                        End If
                    End If
                Case 3
                    If female = False Then
                        If hardcore Then
                            Call strcpy(statbuf, "King ")
                        Else
                            Call strcpy(statbuf, "Baron ")
                        End If
                    Else
                        If hardcore Then
                            Call strcpy(statbuf, "Queen ")
                        Else
                            Call strcpy(statbuf, "Baroness ")
                        End If
                    End If
            End Select
        End If
        Call sprintf(statbuf, "%s, ", name)
        If hardcore Then
            If (Asc(p(26)) And &H8) Then
                Call strcpy(statbuf, "a dead ")
            Else
                Call strcpy(statbuf, "a ")
            End If
            Call sprintf(statbuf, "hardcore level %s ", charlevel)
        Else
            Call sprintf(statbuf, "a level %s ", charlevel)
        End If
        Call sprintf(statbuf, "%s", d2classes(charclass))
    End If
    ParseD2Stats = statbuf
End Function

Public Function GetIconWAR3(ByVal IconNum As Long, ByVal Race As String) As String
On Local Error Resume Next
    Select Case Race
        Case "H"
            Select Case IconNum
                Case 1: GetIconWAR3 = "Peon"
                Case 2: GetIconWAR3 = "Footman"
                Case 3: GetIconWAR3 = "Knight"
                Case 4: GetIconWAR3 = "Archmage"
                Case 5: GetIconWAR3 = "Medivh"
                Case Else: GetIconWAR3 = "Unknown Human"
            End Select
        Case "O"
            Select Case IconNum
                Case 1: GetIconWAR3 = "Peon"
                Case 2: GetIconWAR3 = "Grunt"
                Case 3: GetIconWAR3 = "Tauren"
                Case 4: GetIconWAR3 = "Far Seer"
                Case 5: GetIconWAR3 = "Thrall"
                Case Else: GetIconWAR3 = "Unknown Orc"
            End Select
        Case "N"
            Select Case IconNum
                Case 1: GetIconWAR3 = "Peon"
                Case 2: GetIconWAR3 = "Archer"
                Case 3: GetIconWAR3 = "Druid of the Claw"
                Case 4: GetIconWAR3 = "Priestess of the Moon"
                Case 5: GetIconWAR3 = "Furion Stomrage"
                Case Else: GetIconWAR3 = "Unknown Night Elf"
            End Select
        Case "U"
            Select Case IconNum
                Case 1: GetIconWAR3 = "Peon"
                Case 2: GetIconWAR3 = "Ghoul"
                Case 3: GetIconWAR3 = "Abomination"
                Case 4: GetIconWAR3 = "Lich"
                Case 5: GetIconWAR3 = "Tichondrius"
                Case Else: GetIconWAR3 = "Unknown Undead"
            End Select
        Case "R"
            Select Case IconNum
                Case 1: GetIconWAR3 = "Peon"
                Case 2: GetIconWAR3 = "Green Dragon Whelp"
                Case 3: GetIconWAR3 = "Blue Dragon"
                Case 4: GetIconWAR3 = "Red Dragon"
                Case 5: GetIconWAR3 = "Deathwing"
                Case Else: GetIconWAR3 = "Unknown Random"
            End Select
        Case "K"
            Select Case IconNum
                Case 1: GetIconWAR3 = "Game Banner"
                Case 2: GetIconWAR3 = "Erik the Swift"
                Case 3: GetIconWAR3 = "Olaf the Stout"
                Case 4: GetIconWAR3 = "Baleog the Fierce"
                Case Else: GetIconWAR3 = "Unknown Viking"
            End Select
        Case Else
            GetIconWAR3 = "Unknown"
    End Select
    RaceIcon = GetIconWAR3
End Function

Public Function GetIconW3XP(ByVal IconNum As Long, ByVal Race As String) As String
On Local Error Resume Next
    Select Case Race
        Case "H"
            Select Case IconNum
                Case 1: GetIconW3XP = "Peon"
                Case 2: GetIconW3XP = "Rifleman"
                Case 3: GetIconW3XP = "Sorceress"
                Case 4: GetIconW3XP = "Spellbreaker"
                Case 5: GetIconW3XP = "Blood Mage"
                Case 5: GetIconW3XP = "Jaina"
                Case Else: GetIconW3XP = "Unknown Human"
            End Select
        Case "O"
            Select Case IconNum
                Case 1: GetIconW3XP = "Peon"
                Case 2: GetIconW3XP = "Troll Headhunter"
                Case 3: GetIconW3XP = "Shaman"
                Case 4: GetIconW3XP = "Spirit Walker"
                Case 5: GetIconW3XP = "Shadow Hunter"
                Case 6: GetIconW3XP = "Rexxar"
                Case Else: GetIconW3XP = "Unknown Orc"
            End Select
        Case "N"
            Select Case IconNum
                Case 1: GetIconW3XP = "Peon"
                Case 2: GetIconW3XP = "Huntress"
                Case 3: GetIconW3XP = "Druid of the Talon"
                Case 4: GetIconW3XP = "Dryad"
                Case 5: GetIconW3XP = "Keeper of the Grove"
                Case 6: GetIconW3XP = "Maiev"
                Case Else: GetIconW3XP = "Unknown Night Elf"
            End Select
        Case "U"
            Select Case IconNum
                Case 1: GetIconW3XP = "Peon"
                Case 2: GetIconW3XP = "Crypt Fiend"
                Case 3: GetIconW3XP = "Banshee"
                Case 4: GetIconW3XP = "Destroyer"
                Case 5: GetIconW3XP = "Crypt Lord"
                Case 5: GetIconW3XP = "Sylvanas"
                Case Else: GetIconW3XP = "Unknown Undead"
            End Select
        Case "R"
            Select Case IconNum
                Case 1: GetIconW3XP = "Peon"
                Case 2: GetIconW3XP = "Myrmidon"
                Case 3: GetIconW3XP = "Siren"
                Case 4: GetIconW3XP = "Dragon Turtle"
                Case 5: GetIconW3XP = "Sea Witch"
                Case 6: GetIconW3XP = "Illidan"
                Case Else: GetIconW3XP = "Unknown Random"
            End Select
        Case "T"
            Select Case IconNum
                Case 1: GetIconW3XP = "Peon"
                Case 2: GetIconW3XP = "Felguard"
                Case 3: GetIconW3XP = "Infernal"
                Case 4: GetIconW3XP = "Doomguard"
                Case 5: GetIconW3XP = "Pit Lord"
                Case 6: GetIconW3XP = "Archimonde"
                Case Else: GetIconW3XP = "Unknown Tournament"
            End Select
        Case Else
            GetIconW3XP = "Unknown"
    End Select
    RaceIcon = GetIconW3XP
End Function

Public Sub RealmCat(ByVal stats As String)
On Local Error Resume Next
    Dim d2classes(0 To 7) As String
        d2classes(0) = "Amazon"
        d2classes(1) = "Sorceress"
        d2classes(2) = "Necromancer"
        d2classes(3) = "Paladin"
        d2classes(4) = "Barbarian"
        d2classes(5) = "Druid"
        d2classes(6) = "Assassin"
        d2classes(7) = "Unknown"
    Dim statbuf As String, p() As String, server As String, name As String, icon As Integer
    Dim Product As String
    If Len(stats) > 4 Then
        Dim sLen As Byte
        sLen = GetServer(stats, server)
        sLen = GetCharacterName(stats, sLen, name)
        Call MakeArray(Mid$(stats, sLen), p())
    End If
    If Left$(stats, 4) = "VD2D" Then
        Product = ""
    Else
        Product = "X"
    End If
    Dim Version As Byte
    Version = Asc(p(0)) - &H80
    Dim charclass As Byte
    charclass = Asc(p(13)) - 1
    If (charclass < 0) Or (charclass > 6) Then
        charclass = 7
    End If
    Dim charlevel As Byte
    charlevel = Asc(p(25))
    Dim hardcore As Byte
    hardcore = Asc(p(26)) And 4
    Dim expansion As Boolean
    expansion = False
    frmRealm.lstCharacter.ListItems(frmRealm.lstCharacter.ListItems.Count).ListSubItems.Add 1, , d2classes(charclass), icon
    frmRealm.lstCharacter.ListItems(frmRealm.lstCharacter.ListItems.Count).ListSubItems.Add 2, , charlevel, icon
    If hardcore Then
        frmRealm.lstCharacter.ListItems(frmRealm.lstCharacter.ListItems.Count).ListSubItems.Add 3, , "X", icon
        If (Asc(p(26)) And &H8) Then frmRealm.lstCharacter.ListItems(frmRealm.lstCharacter.ListItems.Count).ForeColor = &H808080
    Else
        frmRealm.lstCharacter.ListItems(frmRealm.lstCharacter.ListItems.Count).ListSubItems.Add 3, , "", icon
    End If
    frmRealm.lstCharacter.ListItems(frmRealm.lstCharacter.ListItems.Count).ListSubItems.Add 4, , Product, icon
End Sub

Private Function GetServer(ByVal statstring As String, ByRef server As String) As Byte
    server = Mid$(statstring, 5, InStr(5, statstring, ",") - 5)
    GetServer = InStr(5, statstring, ",") + 1
End Function

Private Function GetCharacterName(ByVal statstring As String, ByVal start As Byte, ByRef cName As String) As Byte
    cName = Mid$(statstring, start, InStr(start, statstring, ",") - start)
    GetCharacterName = InStr(start, statstring, ",") + 1
End Function

Private Sub MakeArray(ByVal text As String, ByRef nArray() As String)
On Local Error Resume Next
    Dim i As Long
    ReDim nArray(0)
    For i = 0 To Len(text)
        nArray(i) = Mid$(text, i + 1, 1)
        If i <> Len(text) Then
            ReDim Preserve nArray(0 To UBound(nArray) + 1)
        End If
    Next i
End Sub

Public Sub strcpy(ByRef Source As String, ByVal nText As String): Source = Source & nText: End Sub

