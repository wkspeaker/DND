Attribute VB_Name = "Functions"
Option Explicit
' Constants for property types
Private Const PROP_TYPE_COLLECTION As String = "Collection"

Public Function ConvertListObjectToMarkdown(ByVal TableName As String) As String
    Dim tbl As ListObject
    Dim Result As String
    Dim headerRow As String
    Dim separatorRow As String
    Dim dataRow As Variant
    Dim i As Long, j As Long
    Dim outputRow As Long
    
    ' Clear output worksheet
    shOutput.Cells.Clear
    
    ' Iterate through all ListObjects
    For Each tbl In shTableSchema.ListObjects
        ' Check if ListObject name matches
        If tbl.Name = TableName Then
            ' Build header row
            headerRow = "|"
            For i = 1 To tbl.HeaderRowRange.Columns.Count
                headerRow = headerRow & tbl.HeaderRowRange.Cells(1, i).Value & "|"
            Next i
            
            ' Build separator row
            separatorRow = "|"
            For i = 1 To tbl.HeaderRowRange.Columns.Count
                separatorRow = separatorRow & "---|"
            Next i
            
            ' Initialize result string
            Result = headerRow & vbCrLf & separatorRow
            
            ' Build data rows
            If Not tbl.DataBodyRange Is Nothing Then
                For i = 1 To tbl.DataBodyRange.Rows.Count
                    dataRow = "|"
                    For j = 1 To tbl.DataBodyRange.Columns.Count
                        dataRow = dataRow & tbl.DataBodyRange.Cells(i, j).Value & "|"
                    Next j
                    Result = Result & vbCrLf & dataRow
                Next i
            End If
            
            ' Output to worksheet
            outputRow = 1
            For Each dataRow In Split(Result, vbCrLf)
                shOutput.Cells(outputRow, 1).Value = dataRow
                outputRow = outputRow + 1
            Next dataRow
            
            ' Return result
            ConvertListObjectToMarkdown = Result
            Exit Function
        End If
    Next tbl
    
    ' If no matching table is found, return error message
    ConvertListObjectToMarkdown = "Table '" & TableName & "' not found in shTableSchema"
    shOutput.Cells(1, 1).Value = ConvertListObjectToMarkdown
End Function

' Test function to function ConvertListObjectToMarkdown
Public Sub TestConvertToMarkdown()
    ' Test converting table named "Test"
    ConvertListObjectToMarkdown "CharacterEquipmentSchema"
End Sub

Private Function GetFieldColumn(ByVal TableName As String, ByVal FieldName As String) As Integer
' Get column number from Excel sheets based on table/field name
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim headerRange As Range
    Dim Cell As Range
    
    ' Get the worksheet based on table name
    Select Case TableName
        Case "CharacterMaster"
            Set ws = shCharacterMaster
        Case "CharacterMemo"
            Set ws = shCharacterMemo
        Case "CharacterAttackSpell"
            Set ws = shCharacterAttackSpell
        Case "CharacterEquipment"
            Set ws = shCharacterEquipment
        Case Else
            GetFieldColumn = 0
            Exit Function
    End Select
    
    ' Get the ListObject
    Set lo = ws.ListObjects(TableName)
    Set headerRange = lo.HeaderRowRange
    
    ' Find the column
    For Each Cell In headerRange
        If Cell.Value = FieldName Then
            GetFieldColumn = Cell.Column - headerRange.Column + 1
            Exit Function
        End If
    Next Cell
    
    GetFieldColumn = 0
End Function

Public Sub ReadCharacters()
    Dim i As Long
    Dim j As Long
    Dim charMaster As CharacterMaster
    Dim charMemo As CharacterMemo
    Dim charAttackSpell As CharacterAttackSpell
    Dim charEquipment As CharacterEquipment
    Dim dirtyDataCount As Long
    
    ' First pass: Create CharacterMaster objects and add to dictionary
    If Not IsEmpty(CharacterMasterList) Then
        For i = 1 To UBound(CharacterMasterList, 1)
            Set charMaster = New CharacterMaster
            
            ' Set CharacterMaster properties from CharacterMasterList
            With charMaster
                .CharacterId = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "CharacterID"))
                .CharacterType = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "CharacterType"))
                .CharacterStatus = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "CharacterStatus"))
                .Player = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "Player"))
                .Character = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "Character"))
                .Background = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "Background"))
                .Class = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "Class"))
                .ClassLv = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "ClassLv"))
                .Race = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "Race"))
                .Alignment = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "Alignment"))
                .Exp = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "Exp"))
                .Strength = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "Strength"))
                .StrengthAdd = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "StrengthAdd"))
                .Dexterity = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "Dexterity"))
                .DexterityAdd = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "DexterityAdd"))
                .Constitution = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "Constitution"))
                .ConstitutionAdd = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "ConstitutionAdd"))
                .Intelligence = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "Intelligence"))
                .IntelligenceAdd = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "IntelligenceAdd"))
                .Wisdom = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "Wisdom"))
                .WisdomAdd = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "WisdomAdd"))
                .Charisma = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "Charisma"))
                .CharismaAdd = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "CharismaAdd"))
                .ArmorClass = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "ArmorClass"))
                .Initiative = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "Initiative"))
                .Speed = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "Speed"))
                .Inspiration = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "Inspiration"))
                .ProficiencyBonus = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "ProficiencyBonus"))
                .SavingThrowStr = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SavingThrowStr"))
                .SavingThrowDex = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SavingThrowDex"))
                .SavingThrowCon = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SavingThrowCon"))
                .SavingThrowInt = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SavingThrowInt"))
                .SavingThrowWis = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SavingThrowWis"))
                .SavingThrowCha = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SavingThrowCha"))
                .SavingThrowStrP = ReadBoolean(CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SavingThrowStrP")))
                .SavingThrowDexP = ReadBoolean(CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SavingThrowDexP")))
                .SavingThrowConP = ReadBoolean(CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SavingThrowConP")))
                .SavingThrowIntP = ReadBoolean(CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SavingThrowIntP")))
                .SavingThrowWisP = ReadBoolean(CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SavingThrowWisP")))
                .SavingThrowChaP = ReadBoolean(CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SavingThrowChaP")))
                .SkillAcrobatics = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SkillAcrobatics"))
                .SkillAnimalHandling = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SkillAnimalHandling"))
                .SkillArcana = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SkillArcana"))
                .SkillAthletics = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SkillAthletics"))
                .SkillDeception = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SkillDeception"))
                .SkillHistory = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SkillHistory"))
                .SkillInsight = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SkillInsight"))
                .SkillIntimidation = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SkillIntimidation"))
                .SkillInvestigation = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SkillInvestigation"))
                .SkillMedicine = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SkillMedicine"))
                .SkillNature = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SkillNature"))
                .SkillPerception = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SkillPerception"))
                .SkillPerformance = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SkillPerformance"))
                .SkillPersuasion = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SkillPersuasion"))
                .SkillReligion = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SkillReligion"))
                .SkillSleightOfHand = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SkillSleightOfHand"))
                .SkillStealth = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SkillStealth"))
                .SkillSurvival = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SkillSurvival"))
                .SkillAcrobaticsP = ReadBoolean(CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SkillAcrobaticsP")))
                .SkillAnimalHandlingP = ReadBoolean(CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SkillAnimalHandlingP")))
                .SkillArcanaP = ReadBoolean(CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SkillArcanaP")))
                .SkillAthleticsP = ReadBoolean(CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SkillAthleticsP")))
                .SkillDeceptionP = ReadBoolean(CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SkillDeceptionP")))
                .SkillHistoryP = ReadBoolean(CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SkillHistoryP")))
                .SkillInsightP = ReadBoolean(CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SkillInsightP")))
                .SkillIntimidationP = ReadBoolean(CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SkillIntimidationP")))
                .SkillInvestigationP = ReadBoolean(CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SkillInvestigationP")))
                .SkillMedicineP = ReadBoolean(CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SkillMedicineP")))
                .SkillNatureP = ReadBoolean(CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SkillNatureP")))
                .SkillPerceptionP = ReadBoolean(CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SkillPerceptionP")))
                .SkillPerformanceP = ReadBoolean(CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SkillPerformanceP")))
                .SkillPersuasionP = ReadBoolean(CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SkillPersuasionP")))
                .SkillReligionP = ReadBoolean(CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SkillReligionP")))
                .SkillSleightOfHandP = ReadBoolean(CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SkillSleightOfHandP")))
                .SkillStealthP = ReadBoolean(CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SkillStealthP")))
                .SkillSurvivalP = ReadBoolean(CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SkillSurvivalP")))
                .PassiveWisdom = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "PassiveWisdom"))
                .MaxHP = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "MaxHP"))
                .CurHP = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "CurHP"))
                .TmpHP = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "TmpHP"))
                .HD = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "HD"))
                .MaxHD = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "MaxHD"))
                .MoneyCP = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "MoneyCP"))
                .MoneySP = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "MoneySP"))
                .MoneyEP = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "MoneyEP"))
                .MoneyGP = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "MoneyGP"))
                .MoneyPP = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "MoneyPP"))
                .Age = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "Age"))
                .Height = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "Height"))
                .Weight = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "Weight"))
                .Eyes = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "Eyes"))
                .Skin = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "Skin"))
                .Hair = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "Hair"))
            End With
            
            ' Add CharacterMaster to dictionary
            Set Characters(charMaster.CharacterId) = charMaster
        Next i
    End If
    
    ' Second pass: Process CharacterMemoList
    If Not IsEmpty(CharacterMemoList) Then
        For j = 1 To UBound(CharacterMemoList, 1)
            If Characters.Exists(CharacterMemoList(j, GetFieldColumn("CharacterMemo", "CharacterID"))) Then
                Set charMemo = New CharacterMemo
                With charMemo
                    .CharacterId = CharacterMemoList(j, GetFieldColumn("CharacterMemo", "CharacterID"))
                    .MemoType = CharacterMemoList(j, GetFieldColumn("CharacterMemo", "MemoType"))
                    .Contents = CharacterMemoList(j, GetFieldColumn("CharacterMemo", "Contents"))
                End With
                Characters(charMemo.CharacterId).CharacterMemoList.Add charMemo
            Else
                dirtyDataCount = dirtyDataCount + 1
            End If
        Next j
    End If
    
    ' Third pass: Process CharacterAttackSpellList
    If Not IsEmpty(CharacterAttackSpellList) Then
        For j = 1 To UBound(CharacterAttackSpellList, 1)
            If Characters.Exists(CharacterAttackSpellList(j, GetFieldColumn("CharacterAttackSpell", "CharacterID"))) Then
                Set charAttackSpell = New CharacterAttackSpell
                With charAttackSpell
                    .CharacterId = CharacterAttackSpellList(j, GetFieldColumn("CharacterAttackSpell", "CharacterID"))
                    .ItemType = CharacterAttackSpellList(j, GetFieldColumn("CharacterAttackSpell", "Type"))
                    .Name = CharacterAttackSpellList(j, GetFieldColumn("CharacterAttackSpell", "Name"))
                    .AtkBonus = CharacterAttackSpellList(j, GetFieldColumn("CharacterAttackSpell", "AtkBonus"))
                    .Damage_Type = CharacterAttackSpellList(j, GetFieldColumn("CharacterAttackSpell", "Damage_Type"))
                    .SpellMemo = CharacterAttackSpellList(j, GetFieldColumn("CharacterAttackSpell", "SpellMemo"))
                    .Attuned = ReadBoolean(CharacterAttackSpellList(j, GetFieldColumn("CharacterAttackSpell", "Attuned")))
                    .Equiped = ReadBoolean(CharacterAttackSpellList(j, GetFieldColumn("CharacterAttackSpell", "Equiped")))
                End With
                Characters(charAttackSpell.CharacterId).CharacterAttackSpellList.Add charAttackSpell
            Else
                dirtyDataCount = dirtyDataCount + 1
            End If
        Next j
    End If
    
    ' Fourth pass: Process CharacterEquipmentList
    If Not IsEmpty(CharacterEquipmentList) Then
        For j = 1 To UBound(CharacterEquipmentList, 1)
            If Characters.Exists(CharacterEquipmentList(j, GetFieldColumn("CharacterEquipment", "CharacterID"))) Then
                Set charEquipment = New CharacterEquipment
                With charEquipment
                    .CharacterId = CharacterEquipmentList(j, GetFieldColumn("CharacterEquipment", "CharacterID"))
                    .ItemType = CharacterEquipmentList(j, GetFieldColumn("CharacterEquipment", "Type"))
                    .Name = CharacterEquipmentList(j, GetFieldColumn("CharacterEquipment", "Name"))
                    .Quantity = CharacterEquipmentList(j, GetFieldColumn("CharacterEquipment", "Quantity"))
                    .Attuned = ReadBoolean(CharacterEquipmentList(j, GetFieldColumn("CharacterEquipment", "Attuned")))
                    .Equiped = ReadBoolean(CharacterEquipmentList(j, GetFieldColumn("CharacterEquipment", "Equiped")))
                End With
                Characters(charEquipment.CharacterId).pCharacterEquipmentList.Add charEquipment
            Else
                dirtyDataCount = dirtyDataCount + 1
            End If
        Next j
    End If
    
    ' Report dirty data if any
    If dirtyDataCount > 0 Then
        MsgBox "Found " & dirtyDataCount & " records with invalid CharacterID references.", vbInformation, "Data Validation"
    End If
End Sub

Public Sub WriteCharacters()
    Dim charMaster As CharacterMaster
    Dim charMemo As CharacterMemo
    Dim charAttackSpell As CharacterAttackSpell
    Dim charEquipment As CharacterEquipment
    Dim i As Long
    Dim j As Long
    Dim masterRowIndex As Long
    Dim tmpRowIndex As Long
    
    ' Clear all related tables
    shCharacterMaster.ListObjects(1).DataBodyRange.Clear
    shCharacterMemo.ListObjects(1).DataBodyRange.Clear
    shCharacterAttackSpell.ListObjects(1).DataBodyRange.Clear
    shCharacterEquipment.ListObjects(1).DataBodyRange.Clear
    
    ' Initialize master row index
    masterRowIndex = 1
    
    ' Write CharacterMaster data and its related data
    For Each charMaster In Characters.Items
        ' Write CharacterMaster data
        With shCharacterMaster.ListObjects(1)
            .ListRows.Add
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "CharacterID")) = charMaster.CharacterId
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "CharacterType")) = charMaster.CharacterType
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "CharacterStatus")) = charMaster.CharacterStatus
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "Player")) = charMaster.Player
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "Character")) = charMaster.Character
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "Background")) = charMaster.Background
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "Class")) = charMaster.Class
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "ClassLv")) = charMaster.ClassLv
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "Race")) = charMaster.Race
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "Alignment")) = charMaster.Alignment
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "Exp")) = charMaster.Exp
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "Strength")) = charMaster.Strength
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "StrengthAdd")) = charMaster.StrengthAdd
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "Dexterity")) = charMaster.Dexterity
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "DexterityAdd")) = charMaster.DexterityAdd
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "Constitution")) = charMaster.Constitution
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "ConstitutionAdd")) = charMaster.ConstitutionAdd
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "Intelligence")) = charMaster.Intelligence
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "IntelligenceAdd")) = charMaster.IntelligenceAdd
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "Wisdom")) = charMaster.Wisdom
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "WisdomAdd")) = charMaster.WisdomAdd
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "Charisma")) = charMaster.Charisma
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "CharismaAdd")) = charMaster.CharismaAdd
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "ArmorClass")) = charMaster.ArmorClass
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "Initiative")) = charMaster.Initiative
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "Speed")) = charMaster.Speed
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "Inspiration")) = charMaster.Inspiration
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "ProficiencyBonus")) = charMaster.ProficiencyBonus
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SavingThrowStr")) = charMaster.SavingThrowStr
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SavingThrowDex")) = charMaster.SavingThrowDex
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SavingThrowCon")) = charMaster.SavingThrowCon
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SavingThrowInt")) = charMaster.SavingThrowInt
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SavingThrowWis")) = charMaster.SavingThrowWis
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SavingThrowCha")) = charMaster.SavingThrowCha
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SavingThrowStrP")) = charMaster.SavingThrowStrP
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SavingThrowDexP")) = charMaster.SavingThrowDexP
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SavingThrowConP")) = charMaster.SavingThrowConP
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SavingThrowIntP")) = charMaster.SavingThrowIntP
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SavingThrowWisP")) = charMaster.SavingThrowWisP
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SavingThrowChaP")) = charMaster.SavingThrowChaP
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SkillAcrobatics")) = charMaster.SkillAcrobatics
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SkillAnimalHandling")) = charMaster.SkillAnimalHandling
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SkillArcana")) = charMaster.SkillArcana
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SkillAthletics")) = charMaster.SkillAthletics
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SkillDeception")) = charMaster.SkillDeception
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SkillHistory")) = charMaster.SkillHistory
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SkillInsight")) = charMaster.SkillInsight
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SkillIntimidation")) = charMaster.SkillIntimidation
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SkillInvestigation")) = charMaster.SkillInvestigation
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SkillMedicine")) = charMaster.SkillMedicine
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SkillNature")) = charMaster.SkillNature
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SkillPerception")) = charMaster.SkillPerception
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SkillPerformance")) = charMaster.SkillPerformance
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SkillPersuasion")) = charMaster.SkillPersuasion
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SkillReligion")) = charMaster.SkillReligion
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SkillSleightOfHand")) = charMaster.SkillSleightOfHand
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SkillStealth")) = charMaster.SkillStealth
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SkillSurvival")) = charMaster.SkillSurvival
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SkillAcrobaticsP")) = charMaster.SkillAcrobaticsP
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SkillAnimalHandlingP")) = charMaster.SkillAnimalHandlingP
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SkillArcanaP")) = charMaster.SkillArcanaP
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SkillAthleticsP")) = charMaster.SkillAthleticsP
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SkillDeceptionP")) = charMaster.SkillDeceptionP
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SkillHistoryP")) = charMaster.SkillHistoryP
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SkillInsightP")) = charMaster.SkillInsightP
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SkillIntimidationP")) = charMaster.SkillIntimidationP
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SkillInvestigationP")) = charMaster.SkillInvestigationP
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SkillMedicineP")) = charMaster.SkillMedicineP
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SkillNatureP")) = charMaster.SkillNatureP
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SkillPerceptionP")) = charMaster.SkillPerceptionP
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SkillPerformanceP")) = charMaster.SkillPerformanceP
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SkillPersuasionP")) = charMaster.SkillPersuasionP
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SkillReligionP")) = charMaster.SkillReligionP
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SkillSleightOfHandP")) = charMaster.SkillSleightOfHandP
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SkillStealthP")) = charMaster.SkillStealthP
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "SkillSurvivalP")) = charMaster.SkillSurvivalP
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "PassiveWisdom")) = charMaster.PassiveWisdom
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "MaxHP")) = charMaster.MaxHP
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "CurHP")) = charMaster.CurHP
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "TmpHP")) = charMaster.TmpHP
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "HD")) = charMaster.HD
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "MaxHD")) = charMaster.MaxHD
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "MoneyCP")) = charMaster.MoneyCP
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "MoneySP")) = charMaster.MoneySP
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "MoneyEP")) = charMaster.MoneyEP
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "MoneyGP")) = charMaster.MoneyGP
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "MoneyPP")) = charMaster.MoneyPP
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "Age")) = charMaster.Age
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "Height")) = charMaster.Height
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "Weight")) = charMaster.Weight
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "Eyes")) = charMaster.Eyes
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "Skin")) = charMaster.Skin
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "Hair")) = charMaster.Hair
        End With
        
        ' Write CharacterMemo data for current CharacterMaster
        For Each charMemo In charMaster.CharacterMemoList
            With shCharacterMemo.ListObjects(1)
                .ListRows.Add
                tmpRowIndex = .DataBodyRange.Rows.Count
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterMemo", "CharacterID")) = charMemo.CharacterId
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterMemo", "MemoType")) = charMemo.MemoType
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterMemo", "Contents")) = charMemo.Contents
            End With
        Next charMemo
        
        ' Write CharacterAttackSpell data for current CharacterMaster
        For Each charAttackSpell In charMaster.CharacterAttackSpellList
            With shCharacterAttackSpell.ListObjects(1)
                .ListRows.Add
                tmpRowIndex = .DataBodyRange.Rows.Count
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterAttackSpell", "CharacterID")) = charAttackSpell.CharacterId
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterAttackSpell", "Type")) = charAttackSpell.ItemType
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterAttackSpell", "Name")) = charAttackSpell.Name
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterAttackSpell", "AtkBonus")) = charAttackSpell.AtkBonus
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterAttackSpell", "Damage_Type")) = charAttackSpell.Damage_Type
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterAttackSpell", "SpellMemo")) = charAttackSpell.SpellMemo
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterAttackSpell", "Attuned")) = charAttackSpell.Attuned
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterAttackSpell", "Equiped")) = charAttackSpell.Equiped
            End With
        Next charAttackSpell
        
        ' Write CharacterEquipment data for current CharacterMaster
        For Each charEquipment In charMaster.pCharacterEquipmentList
            With shCharacterEquipment.ListObjects(1)
                .ListRows.Add
                tmpRowIndex = .DataBodyRange.Rows.Count
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterEquipment", "CharacterID")) = charEquipment.CharacterId
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterEquipment", "Type")) = charEquipment.ItemType
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterEquipment", "Name")) = charEquipment.Name
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterEquipment", "Quantity")) = charEquipment.Quantity
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterEquipment", "Attuned")) = charEquipment.Attuned
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterEquipment", "Equiped")) = charEquipment.Equiped
            End With
        Next charEquipment
        
        masterRowIndex = masterRowIndex + 1
    Next charMaster
End Sub

Public Sub CharacterToUI(ByVal CharacterId As Long)
    ' Check if character exists in dictionary
    If Not Characters.Exists(CharacterId) Then
        MsgBox "Character ID " & CharacterId & " not found in dictionary", vbExclamation, "Error"
        Exit Sub
    End If
    
    ' Get character from dictionary
    Dim Character As CharacterMaster
    Set Character = Characters(CharacterId)
    
    ' Get all properties from CharacterMaster class
    Dim Prop As Variant
    Dim PropValue As Variant
    Dim RangeExists As Boolean
    Dim Response As VbMsgBoxResult
    
    ' Get properties from CharacterMasterSchema table
    Dim Properties As Variant
    Properties = GetPropertiesFromSchema("CharacterMasterSchema")
    
    For Each Prop In Properties
        ' Skip Collection type properties
        PropValue = CallByName(Character, Prop, VbGet)
        If TypeName(PropValue) = PROP_TYPE_COLLECTION Then
            ' Skip this property as it's a collection
            ' Collections will be handled in future updates
        Else
            ' Check if named range exists in shGeneral
            On Error Resume Next
            RangeExists = Not shGeneral.Range(Prop) Is Nothing
            On Error GoTo 0
            
            If Not RangeExists Then
                ' Ask user if should continue
                Response = MsgBox("Member " & Prop & " not found in UI. Continue writing other members?", _
                                vbQuestion + vbYesNo, "Missing Member")
                
                If Response = vbNo Then
                    Exit Sub
                End If
            Else
                ' Write property value to UI
                If VarType(PropValue) = vbBoolean Then
                    shGeneral.Range(Prop) = WriteBoolean(PropValue)
                Else
                    shGeneral.Range(Prop) = PropValue
                End If
            End If
        End If
    Next Prop
End Sub

' Helper function to get properties from schema table
Private Function GetPropertiesFromSchema(ByVal SchemaName As String) As Variant
    Dim Properties As Collection
    Set Properties = New Collection
    
    ' Get the schema table
    Dim SchemaTable As ListObject
    Set SchemaTable = shTableSchema.ListObjects(SchemaName)
    
    ' Get the field names from the schema
    Dim DataRange As Range
    Set DataRange = SchemaTable.ListColumns("�ֶ�").DataBodyRange
    
    Dim Cell As Range
    For Each Cell In DataRange
        If Not IsEmpty(Cell) Then
            Properties.Add Cell.Value
        End If
    Next Cell
    
    ' Convert collection to array
    Dim Result() As Variant
    ReDim Result(1 To Properties.Count)
    Dim i As Long
    For i = 1 To Properties.Count
        Result(i) = Properties(i)
    Next i
    
    GetPropertiesFromSchema = Result
End Function

' Helper function to get maximum ID from dictionary
Public Function GetMaxCharacterId() As Long
    Dim MaxId As Long
    MaxId = 0
    
    Dim Key As Variant
    For Each Key In Characters.Keys
        If CLng(Key) > MaxId Then
            MaxId = CLng(Key)
        End If
    Next Key
    
    GetMaxCharacterId = MaxId
End Function

' ��Excel��Ԫ����ַ���ת��ΪBoolean
Public Function ReadBoolean(ByVal Value As Variant) As Boolean
    Dim s As String
    s = UCase(Trim(CStr(Value)))
    Select Case s
        Case "Y", "YES"
            ReadBoolean = True
        Case "N", "NO", ""
            ReadBoolean = False
        Case Else
            ReadBoolean = False ' �����׳��쳣/����
    End Select
End Function

' ��Booleanֵת��ΪExcel�õ��ַ���
Public Function WriteBoolean(ByVal Value As Boolean) As String
    If Value Then
        WriteBoolean = "Y"
    Else
        WriteBoolean = "N"
    End If
End Function
