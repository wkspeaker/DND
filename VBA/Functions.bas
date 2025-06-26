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
                headerRow = headerRow & tbl.HeaderRowRange.Cells(1, i).value & "|"
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
                        dataRow = dataRow & tbl.DataBodyRange.Cells(i, j).value & "|"
                    Next j
                    Result = Result & vbCrLf & dataRow
                Next i
            End If
            
            ' Output to worksheet
            outputRow = 1
            For Each dataRow In Split(Result, vbCrLf)
                shOutput.Cells(outputRow, 1).value = dataRow
                outputRow = outputRow + 1
            Next dataRow
            
            ' Return result
            ConvertListObjectToMarkdown = Result
            Exit Function
        End If
    Next tbl
    
    ' If no matching table is found, return error message
    ConvertListObjectToMarkdown = "Table '" & TableName & "' not found in shTableSchema"
    shOutput.Cells(1, 1).value = ConvertListObjectToMarkdown
End Function

' Test function to function ConvertListObjectToMarkdown
Public Sub TestConvertToMarkdown()
    ' Test converting table named "Test"
    ConvertListObjectToMarkdown "CharacterSpellSlotSchema"
End Sub

Private Function GetFieldColumn(ByVal TableName As String, ByVal FieldName As String) As Integer
' Get column number from Excel sheets based on table/field name
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim headerRange As Range
    Dim cell As Range
    
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
        Case "CharacterSpell"
            Set ws = shCharacterSpell
        Case "CharacterSpellSlot"
            Set ws = shCharacterSpellSlot
        Case Else
            GetFieldColumn = 0
            Exit Function
    End Select
    
    ' Get the ListObject
    Set lo = ws.ListObjects(TableName)
    Set headerRange = lo.HeaderRowRange
    
    ' Find the column
    For Each cell In headerRange
        If cell.value = FieldName Then
            GetFieldColumn = cell.Column - headerRange.Column + 1
            Exit Function
        End If
    Next cell
    
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
                .CharacterID = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "CharacterID"))
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
                .SpellCastingClass = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SpellCastingClass"))
                .SpellCastingAbility = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SpellCastingAbility"))
                .SpellSaveDC = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SpellSaveDC"))
                .SpellAttackBonus = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SpellAttackBonus"))
                
            End With
            
            ' Add CharacterMaster to dictionary
            Set Characters(charMaster.CharacterID) = charMaster
        Next i
    End If
    
    ' Second pass: Process CharacterMemoList
    If Not IsEmpty(CharacterMemoList) Then
        For j = 1 To UBound(CharacterMemoList, 1)
            If Characters.Exists(CharacterMemoList(j, GetFieldColumn("CharacterMemo", "CharacterID"))) Then
                Set charMemo = New CharacterMemo
                With charMemo
                    .CharacterID = CharacterMemoList(j, GetFieldColumn("CharacterMemo", "CharacterID"))
                    .MemoType = CharacterMemoList(j, GetFieldColumn("CharacterMemo", "MemoType"))
                    .Contents = CharacterMemoList(j, GetFieldColumn("CharacterMemo", "Contents"))
                End With
                Characters(charMemo.CharacterID).CharacterMemoList.Add charMemo
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
                    .CharacterID = CharacterAttackSpellList(j, GetFieldColumn("CharacterAttackSpell", "CharacterID"))
                    .ItemType = CharacterAttackSpellList(j, GetFieldColumn("CharacterAttackSpell", "Type"))
                    .Name = CharacterAttackSpellList(j, GetFieldColumn("CharacterAttackSpell", "Name"))
                    .AtkBonus = CharacterAttackSpellList(j, GetFieldColumn("CharacterAttackSpell", "AtkBonus"))
                    .Damage_Type = CharacterAttackSpellList(j, GetFieldColumn("CharacterAttackSpell", "Damage_Type"))
                    .SpellMemo = CharacterAttackSpellList(j, GetFieldColumn("CharacterAttackSpell", "SpellMemo"))
                    .Attuned = ReadBoolean(CharacterAttackSpellList(j, GetFieldColumn("CharacterAttackSpell", "Attuned")))
                    .Equiped = ReadBoolean(CharacterAttackSpellList(j, GetFieldColumn("CharacterAttackSpell", "Equiped")))
                End With
                Characters(charAttackSpell.CharacterID).CharacterAttackSpellList.Add charAttackSpell
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
                    .CharacterID = CharacterEquipmentList(j, GetFieldColumn("CharacterEquipment", "CharacterID"))
                    .ItemType = CharacterEquipmentList(j, GetFieldColumn("CharacterEquipment", "Type"))
                    .Name = CharacterEquipmentList(j, GetFieldColumn("CharacterEquipment", "Name"))
                    .Quantity = CharacterEquipmentList(j, GetFieldColumn("CharacterEquipment", "Quantity"))
                    .Attuned = ReadBoolean(CharacterEquipmentList(j, GetFieldColumn("CharacterEquipment", "Attuned")))
                    .Equiped = ReadBoolean(CharacterEquipmentList(j, GetFieldColumn("CharacterEquipment", "Equiped")))
                End With
                Characters(charEquipment.CharacterID).CharacterEquipmentList.Add charEquipment
            Else
                dirtyDataCount = dirtyDataCount + 1
            End If
        Next j
    End If
    
    ' Fifth pass: Process CharacterSpellList
    If Not IsEmpty(CharacterSpellList) Then
        For j = 1 To UBound(CharacterSpellList, 1)
            If Characters.Exists(CharacterSpellList(j, GetFieldColumn("CharacterSpell", "CharacterID"))) Then
                Dim charSpell As CharacterSpell
                Set charSpell = New CharacterSpell
                With charSpell
                    .CharacterID = CharacterSpellList(j, GetFieldColumn("CharacterSpell", "CharacterID"))
                    .SpellLevel = CharacterSpellList(j, GetFieldColumn("CharacterSpell", "SpellLevel"))
                    .Name = CharacterSpellList(j, GetFieldColumn("CharacterSpell", "Name"))
                    .Description = CharacterSpellList(j, GetFieldColumn("CharacterSpell", "Description"))
                    .Prepared = ReadBoolean(CharacterSpellList(j, GetFieldColumn("CharacterSpell", "Prepared")))
                End With
                Characters(charSpell.CharacterID).CharacterSpellList.Add charSpell
            Else
                dirtyDataCount = dirtyDataCount + 1
            End If
        Next j
    End If

    ' Sixth pass: Process CharacterSpellSlotList
    If Not IsEmpty(CharacterSpellSlotList) Then
        For j = 1 To UBound(CharacterSpellSlotList, 1)
            If Characters.Exists(CharacterSpellSlotList(j, GetFieldColumn("CharacterSpellSlot", "CharacterID"))) Then
                Dim charSpellSlot As CharacterSpellSlot
                Set charSpellSlot = New CharacterSpellSlot
                With charSpellSlot
                    .CharacterID = CharacterSpellSlotList(j, GetFieldColumn("CharacterSpellSlot", "CharacterID"))
                    .SpellLevel = CharacterSpellSlotList(j, GetFieldColumn("CharacterSpellSlot", "SpellLevel"))
                    .SlotsTotal = CharacterSpellSlotList(j, GetFieldColumn("CharacterSpellSlot", "SlotsTotal"))
                    .SlotsExpended = CharacterSpellSlotList(j, GetFieldColumn("CharacterSpellSlot", "SlotsExpended"))
                End With
                Characters(charSpellSlot.CharacterID).CharacterSpellSlots.Add charSpellSlot
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
    Dim charSpell As CharacterSpell
    Dim charSpellSlot As CharacterSpellSlot
    Dim i As Long
    Dim j As Long
    Dim masterRowIndex As Long
    Dim tmpRowIndex As Long
    
    ' Clear all related tables
    shCharacterMaster.ListObjects(1).DataBodyRange.Clear
    shCharacterMemo.ListObjects(1).DataBodyRange.Clear
    shCharacterAttackSpell.ListObjects(1).DataBodyRange.Clear
    shCharacterEquipment.ListObjects(1).DataBodyRange.Clear
    shCharacterSpell.ListObjects(1).DataBodyRange.Clear
    shCharacterSpellSlot.ListObjects(1).DataBodyRange.Clear
    
    ' Initialize master row index
    masterRowIndex = 1
    
    ' Write CharacterMaster data and its related data
    For Each charMaster In Characters.Items
        ' Write CharacterMaster data
        With shCharacterMaster.ListObjects(1)
            .ListRows.Add
            .ListRows(masterRowIndex).Range(GetFieldColumn("CharacterMaster", "CharacterID")) = charMaster.CharacterID
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
        End With
        
        ' Write CharacterMemo data for current CharacterMaster
        For Each charMemo In charMaster.CharacterMemoList
            With shCharacterMemo.ListObjects(1)
                .ListRows.Add
                tmpRowIndex = .DataBodyRange.Rows.Count
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterMemo", "CharacterID")) = charMemo.CharacterID
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterMemo", "MemoType")) = charMemo.MemoType
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterMemo", "Contents")) = charMemo.Contents
            End With
        Next charMemo
        
        ' Write CharacterAttackSpell data for current CharacterMaster
        For Each charAttackSpell In charMaster.CharacterAttackSpellList
            With shCharacterAttackSpell.ListObjects(1)
                .ListRows.Add
                tmpRowIndex = .DataBodyRange.Rows.Count
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterAttackSpell", "CharacterID")) = charAttackSpell.CharacterID
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
        For Each charEquipment In charMaster.CharacterEquipmentList
            With shCharacterEquipment.ListObjects(1)
                .ListRows.Add
                tmpRowIndex = .DataBodyRange.Rows.Count
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterEquipment", "CharacterID")) = charEquipment.CharacterID
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterEquipment", "Type")) = charEquipment.ItemType
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterEquipment", "Name")) = charEquipment.Name
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterEquipment", "Quantity")) = charEquipment.Quantity
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterEquipment", "Attuned")) = charEquipment.Attuned
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterEquipment", "Equiped")) = charEquipment.Equiped
            End With
        Next charEquipment

        ' Write CharacterSpell data for current CharacterMaster
        For Each charSpell In charMaster.CharacterSpellList
            With shCharacterSpell.ListObjects(1)
                .ListRows.Add
                tmpRowIndex = .DataBodyRange.Rows.Count
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterSpell", "CharacterID")) = charSpell.CharacterID
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterSpell", "SpellLevel")) = charSpell.SpellLevel
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterSpell", "Name")) = charSpell.Name
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterSpell", "Description")) = charSpell.Description
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterSpell", "Prepared")) = charSpell.Prepared
            End With
        Next charSpell

        ' Write CharacterSpellSlot data for current CharacterMaster
        For Each charSpellSlot In charMaster.CharacterSpellSlots
            With shCharacterSpellSlot.ListObjects(1)
                .ListRows.Add
                tmpRowIndex = .DataBodyRange.Rows.Count
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterSpellSlot", "CharacterID")) = charSpellSlot.CharacterID
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterSpellSlot", "SpellLevel")) = charSpellSlot.SpellLevel
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterSpellSlot", "SlotsTotal")) = charSpellSlot.SlotsTotal
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterSpellSlot", "SlotsExpended")) = charSpellSlot.SlotsExpended
            End With
        Next charSpellSlot

        masterRowIndex = masterRowIndex + 1
    Next charMaster
End Sub

Public Sub CharacterToUI(ByVal CharacterID As Long)
    ' Check if character exists in dictionary
    If Not Characters.Exists(CharacterID) Then
        MsgBox "Character ID " & CharacterID & " not found in dictionary", vbExclamation, "Error"
        Exit Sub
    End If
    
    ' Get character from dictionary
    Dim Character As CharacterMaster
    Set Character = Characters(CharacterID)
    
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
        'If TypeName(PropValue) = PROP_TYPE_COLLECTION Then
            '�����˲����߼�
        'Else
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
        'End If
    Next Prop
    
    
    '���CharacterMemoList�Ĵ���
    If Not Character.CharacterMemoList Is Nothing Then
        If Character.CharacterMemoList.Count > 0 Then
            Dim charMemo As CharacterMemo
            For Each charMemo In Character.CharacterMemoList
                Call WriteDataBlockByRange(charMemo.MemoType, charMemo.Contents)
            Next charMemo
        End If
    End If
    
    '����CharacterAttackSpellList
    If Not Character.CharacterAttackSpellList Is Nothing Then
        Dim charAttackSpell As CharacterAttackSpell
        For Each charAttackSpell In Character.CharacterAttackSpellList
            If charAttackSpell.ItemType = "Attack" Then
                Call WriteAttacksByRange(charAttackSpell)
            ElseIf charAttackSpell.ItemType = "Spell" Then
                Call WriteSpellsByRange(charAttackSpell)
            End If
        Next charAttackSpell
    End If
    
    '����CharacterEquipmentList
    If Not Character.CharacterEquipmentList Is Nothing Then
        Dim charEquipment As CharacterEquipment
        For Each charEquipment In Character.CharacterEquipmentList
            Call WriteEquipmentsByRange(charEquipment)
        Next charEquipment
    End If
    
    '����CharacterSpellList
    If Not Character.CharacterSpellList Is Nothing Then
        Dim charSpell As CharacterSpell
        For Each charSpell In Character.CharacterSpellList
            Call WriteSpellListByRange(charSpell)
        Next charSpell
    End If
    
    '����CharacterSpellSlots
    If Not Character.CharacterSpellSlots Is Nothing Then
        Dim charSpellSlot As CharacterSpellSlot
        For Each charSpellSlot In Character.CharacterSpellSlots
            Call WriteSpellSlotsByRange(charSpellSlot)
        Next charSpellSlot
    End If
End Sub

' Helper function to get properties and types from schema table
Private Function GetPropertiesFromSchema(ByVal SchemaName As String) As Variant
    Dim Properties As Collection
    Set Properties = New Collection
    
    ' Get the schema table
    Dim SchemaTable As ListObject
    Set SchemaTable = shTableSchema.ListObjects(SchemaName)
    
    ' Get the field names and types from the schema
    Dim FieldRange As Range, TypeRange As Range
    Set FieldRange = SchemaTable.ListColumns("�ֶ�").DataBodyRange
    Set TypeRange = SchemaTable.ListColumns("����").DataBodyRange
    
    Dim i As Long
    For i = 1 To FieldRange.Rows.Count
        If Not IsEmpty(FieldRange.Cells(i, 1)) Then
            Dim arr(1 To 2) As Variant
            arr(1) = FieldRange.Cells(i, 1).value ' ???
            arr(2) = TypeRange.Cells(i, 1).value  ' ??
            Properties.Add arr
        End If
    Next i
    
    ' Convert collection to 2D array
    Dim Result() As Variant
    ReDim Result(1 To Properties.Count, 1 To 2)
    For i = 1 To Properties.Count
        Result(i, 1) = Properties(i)(1)
        Result(i, 2) = Properties(i)(2)
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
Public Function ReadBoolean(ByVal value As Variant) As Boolean
    Dim s As String
    s = UCase(Trim(CStr(value)))
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
Public Function WriteBoolean(ByVal value As Boolean) As String
    If value Then
        WriteBoolean = "Y"
    Else
        WriteBoolean = "N"
    End If
End Function

'��CharacterIDName�ж�ȡCharacterID
Public Function GetCharacterIDFromCharacterIDName(ByVal CharacterIDName As String) As Long
    Dim pos As Long
    pos = InStr(CharacterIDName, "|")
    If pos > 0 Then
        GetCharacterIDFromCharacterIDName = CLng(Trim(Left(CharacterIDName, pos - 1)))
    Else
        GetCharacterIDFromCharacterIDName = 0
    End If
End Function

'�����µ�Character����,CharacterIDȡ���ֵ��1, ������CharacterIDֵ
Public Function AddNewCharacter() As Long
    Dim NewCharacterId As Long
    
    ' New character - get max ID and add 1

    NewCharacterId = GetMaxCharacterId() + 1
        
    ' Create new character and add to dictionary
    Dim newCharacter As CharacterMaster
    Set newCharacter = New CharacterMaster
    newCharacter.CharacterID = NewCharacterId
        
    ' Add to dictionary
    Characters.Add NewCharacterId, newCharacter
    
    AddNewCharacter = NewCharacterId
End Function

'��������:��ָ���й��������ᴰ���·�
Public Sub ScrollToRow(ByVal TargetRegion As String)
    Dim freezeRow As Long
    Dim targetRow As Long
    
    targetRow = shGeneral.Range(TargetRegion).Row
    freezeRow = ActiveWindow.SplitRow
    If freezeRow < 1 Then freezeRow = 0
    'ѡ��Ŀ�굥Ԫ��
    shGeneral.Cells(targetRow, 1).Select
    '����,ʹĿ������ʾ�ڶ������·�
    ActiveWindow.ScrollRow = targetRow
End Sub

'����shGeneralҳ��ָ����������ĵڶ��м��Ժ�����
Public Sub PrepDataBlockByRange(ByVal TargetName As String, Optional HasHeadLine As Boolean = False)
    Dim rng As Range
    Dim colCount As Long
    
    '��λ�����������CurrentRegion
    On Error Resume Next
    Set rng = shGeneral.Range(TargetName).CurrentRegion
    On Error GoTo 0
    
    If rng Is Nothing Then
        MsgBox "Named range '" & TargetName & "' not found.", vbExclamation, "Error"
        Exit Sub
    End If
    
    colCount = rng.Columns.Count
    If colCount > 1 Then
        '����HasHeadLine�����������,��մӵڶ��п�ʼ������
        If HasHeadLine Then
            rng.Offset(1, 1).Resize(rng.Rows.Count - 1, colCount - 1).ClearContents
        Else
            rng.Offset(0, 1).Resize(rng.Rows.Count, colCount - 1).ClearContents
        End If
    End If
End Sub

'��shGeneralҳ��ָ�����������Ҳ���������д������
Public Sub WriteDataBlockByRange(ByVal TargetName As String, ByVal content As String)
    Dim targetCell As Range
    Dim writeCell As Range
    Dim i As Long
    
    On Error Resume Next
    Set targetCell = shGeneral.Range(TargetName)
    On Error GoTo 0
    
    If targetCell Is Nothing Then
        MsgBox "Named range '" & TargetName & "' not found.", vbExclamation, "Error"
        Exit Sub
    End If
    
    '���Ҳ����ڵ�Ԫ��ʼ,���²��ҿյ�Ԫ��
    i = 0
    Do
        Set writeCell = targetCell.Offset(i, 1)
        If IsEmpty(writeCell.value) Then
            writeCell.value = content
            Exit Sub
        End If
        i = i + 1
        '��ֹ��ѭ��,������50��
        If i > 50 Then Exit Do
    Loop
    '���û�п�λ,��д��
End Sub

'д��������Ϣ��shGeneralҳ��
Public Sub WriteAttacksByRange(ByRef Attack As CharacterAttackSpell)
    Const TargetName As String = "Attacks"
    Dim targetCell As Range, writeCell As Range
    Dim i As Long
    On Error Resume Next
    Set targetCell = shGeneral.Range(TargetName)
    On Error GoTo 0
    If targetCell Is Nothing Then
        MsgBox "Named range '" & TargetName & "' not found.", vbExclamation, "Error"
        Exit Sub
    End If
    i = 0
    Do
        Set writeCell = targetCell.Offset(i, 1)
        If IsEmpty(writeCell.value) Then
            writeCell.value = Attack.Name
            writeCell.Offset(0, 1).value = Attack.AtkBonus
            writeCell.Offset(0, 2).value = Attack.Damage_Type
            writeCell.Offset(0, 3).value = WriteBoolean(Attack.Equiped)
            writeCell.Offset(0, 4).value = WriteBoolean(Attack.Attuned)
            Exit Sub
        End If
        i = i + 1
        If i > 50 Then Exit Do
    Loop
End Sub

'д�뷨����Ϣ��shGeneralҳ��
Public Sub WriteSpellsByRange(ByRef Spell As CharacterAttackSpell)
    Const TargetName As String = "Spells"
    Dim targetCell As Range, writeCell As Range
    Dim i As Long
    On Error Resume Next
    Set targetCell = shGeneral.Range(TargetName)
    On Error GoTo 0
    If targetCell Is Nothing Then
        MsgBox "Named range '" & TargetName & "' not found.", vbExclamation, "Error"
        Exit Sub
    End If
    i = 0
    Do
        Set writeCell = targetCell.Offset(i, 1)
        If IsEmpty(writeCell.value) Then
            writeCell.value = Spell.Name
            writeCell.Offset(0, 1).value = Spell.AtkBonus
            writeCell.Offset(0, 2).value = Spell.Damage_Type
            writeCell.Offset(0, 3).value = Spell.SpellMemo
            Exit Sub
        End If
        i = i + 1
        If i > 50 Then Exit Do
    Loop
End Sub

'д��װ����Ϣ��shGeneralҳ��
Public Sub WriteEquipmentsByRange(ByRef Equipment As CharacterEquipment)
    Const TargetName As String = "Equipments"
    Dim targetCell As Range, writeCell As Range
    Dim i As Long
    On Error Resume Next
    Set targetCell = shGeneral.Range(TargetName)
    On Error GoTo 0
    If targetCell Is Nothing Then
        MsgBox "Named range '" & TargetName & "' not found.", vbExclamation, "Error"
        Exit Sub
    End If
    i = 0
    Do
        Set writeCell = targetCell.Offset(i, 1)
        If IsEmpty(writeCell.value) Then
            writeCell.value = Equipment.Name
            writeCell.Offset(0, 1).value = Equipment.Quantity
            writeCell.Offset(0, 2).value = WriteBoolean(Equipment.Attuned)
            writeCell.Offset(0, 3).value = WriteBoolean(Equipment.Equiped)
            Exit Sub
        End If
        i = i + 1
        If i > 50 Then Exit Do
    Loop
End Sub

'д�뷨���б���shGeneralҳ��
Public Sub WriteSpellListByRange(ByRef Spell As CharacterSpell)
    Const TargetName As String = "SpellList"
    Dim targetCell As Range, writeCell As Range
    Dim i As Long
    On Error Resume Next
    Set targetCell = shGeneral.Range(TargetName)
    On Error GoTo 0
    If targetCell Is Nothing Then
        MsgBox "Named range '" & TargetName & "' not found.", vbExclamation, "Error"
        Exit Sub
    End If
    i = 0
    Do
        Set writeCell = targetCell.Offset(i, 1)
        If IsEmpty(writeCell.value) Then
            writeCell.value = Spell.SpellLevel
            writeCell.Offset(0, 1).value = Spell.Name
            writeCell.Offset(0, 2).value = Spell.Description
            writeCell.Offset(0, 3).value = WriteBoolean(Spell.Prepared)
            Exit Sub
        End If
        i = i + 1
        If i > 50 Then Exit Do
    Loop
End Sub

'д�뷨��λ��shGeneralҳ��
Public Sub WriteSpellSlotsByRange(ByRef SpellSlot As CharacterSpellSlot)
    Const TargetName As String = "SpellSlots"
    Dim targetCell As Range, writeCell As Range
    Dim i As Long
    On Error Resume Next
    Set targetCell = shGeneral.Range(TargetName)
    On Error GoTo 0
    If targetCell Is Nothing Then
        MsgBox "Named range '" & TargetName & "' not found.", vbExclamation, "Error"
        Exit Sub
    End If
    i = 0
    Do
        Set writeCell = targetCell.Offset(i, 1)
        If IsEmpty(writeCell.value) Then
            writeCell.value = SpellSlot.SpellLevel
            writeCell.Offset(0, 1).value = SpellSlot.SlotsTotal
            writeCell.Offset(0, 2).value = SpellSlot.SlotsExpended
            Exit Sub
        End If
        i = i + 1
        If i > 10 Then Exit Do
    Loop
End Sub


'�����������������ݿ�
Public Sub PrepDataBlocksBetweenNames(ByVal StartName As String, Optional HasHeadLine As Boolean = False)
    Dim startCell As Range
    ' 1.������ʵ��Ԫ��
    On Error Resume Next
    Set startCell = shGeneral.Range(StartName)
    On Error GoTo 0
    If startCell Is Nothing Then
        MsgBox "Named range '" & StartName & "' not found.", vbExclamation, "Error"
        Exit Sub
    End If

    Dim startRow As Long
    startRow = startCell.Row
    Dim col As Long
    col = startCell.Column

    ' 2.���²�����һ�������Ƶĵ�Ԫ��,������100��
    Dim endRow As Long
    Dim i As Long
    endRow = startRow + 100
    For i = startRow + 1 To startRow + 100
        If i > shGeneral.Rows.Count Then Exit For
        If HasCellName(shGeneral.Cells(i, col)) Then
            If shGeneral.Cells(i, col).Name.NameLocal <> shGeneral.Cells(i, col).Address(False, False, xlA1, True) Then
                endRow = i
                Exit For
            End If
        End If
    Next i
    If endRow > shGeneral.Rows.Count Then endRow = shGeneral.Rows.Count

    ' 3.��ȡ����(��ʼ��+1 �� ������-1)
    Dim regionStart As Long, regionEnd As Long
    regionStart = startRow + 1
    regionEnd = endRow - 1
    If regionStart > regionEnd Then Exit Sub '������Ч

    'ֻ���������������
    Dim usedColCount As Long
    usedColCount = shGeneral.UsedRange.Columns.Count
    Dim r As Long, c As Long
    For r = regionStart To regionEnd
        For c = 1 To usedColCount
            Dim cell As Range
            Set cell = shGeneral.Cells(r, c)
            If HasCellName(cell) Then
                If cell.Name.NameLocal <> cell.Address(False, False, xlA1, True) Then
                    Call PrepDataBlockByRange(cell.Name.NameLocal, HasHeadLine)
                End If
            End If
        Next c
    Next r
End Sub

Function HasCellName(cell As Range) As Boolean
    On Error Resume Next
    Dim n As String
    n = cell.Name.NameLocal
    HasCellName = (Err.Number = 0)
    On Error GoTo 0
End Function


'��ȡָ������������������Ԫ��������б�
Public Function GetNamesByRegionName(ByVal TargetName As String) As Collection
    Dim startCell As Range
    Dim startRow As Long, endRow As Long, col As Long, i As Long
    Dim usedColCount As Long
    Dim cell As Range
    Dim namesList As New Collection
    
    On Error Resume Next
    Set startCell = shGeneral.Range(TargetName)
    On Error GoTo 0
    If startCell Is Nothing Then Exit Function
    
    startRow = startCell.Row
    col = startCell.Column
    endRow = startRow + 100
    For i = startRow + 1 To startRow + 100
        If i > shGeneral.Rows.Count Then Exit For
        If HasCellName(shGeneral.Cells(i, col)) Then
            If shGeneral.Cells(i, col).Name.NameLocal <> shGeneral.Cells(i, col).Address(False, False, xlA1, True) Then
                endRow = i
                Exit For
            End If
        End If
    Next i
    If endRow > shGeneral.Rows.Count Then endRow = shGeneral.Rows.Count
    
    Dim regionStart As Long, regionEnd As Long
    regionStart = startRow + 1
    regionEnd = endRow - 1
    If regionStart > regionEnd Then Set GetNamesByRegionName = namesList: Exit Function
    
    usedColCount = shGeneral.UsedRange.Columns.Count
    Dim r As Long, c As Long
    For r = regionStart To regionEnd
        For c = 1 To usedColCount
            Set cell = shGeneral.Cells(r, c)
            If HasCellName(cell) Then
                If cell.Name.NameLocal <> cell.Address(False, False, xlA1, True) Then
                    namesList.Add cell.Name.NameLocal
                End If
            End If
        Next c
    Next r
    Set GetNamesByRegionName = namesList
End Function



'��shGeneral��ȡ���ݵ�Character����
Public Function UIToCharacter(ByVal CharacterID As Long) As CharacterMaster
    Dim Character As New CharacterMaster
    Dim Properties As Variant
    Dim v As Variant
    Dim i As Long
    Dim Prop As String, PropType As String
    
    ' 1.��ȡ������
    Properties = GetPropertiesFromSchema("CharacterMasterSchema")
    For i = 1 To UBound(Properties, 1)
        Prop = Properties(i, 1)
        PropType = LCase(Trim(Properties(i, 2)))
        On Error Resume Next
        v = shGeneral.Range(Prop).value
        On Error GoTo 0
        If PropType = "bool" Or PropType = "boolean" Then
            CallByName Character, Prop, VbLet, ReadBoolean(v)
        Else
            CallByName Character, Prop, VbLet, v
        End If
    Next i
    Character.CharacterID = CharacterID
    
    ' 2.��ȡCharacterMemoList(�ޱ�ͷ,��������һ��)
    Dim memoNames As Collection, memoName As Variant
    Set memoNames = GetNamesByRegionName("RegionMemo")
    For Each memoName In memoNames
        Dim memoCell As Range, memoRegion As Range
        Set memoCell = shGeneral.Range(memoName)
        Set memoRegion = memoCell.CurrentRegion
        Dim rowIdx As Long, colIdx As Long, content As String, hasData As Boolean
        For rowIdx = 1 To memoRegion.Rows.Count
            content = ""
            hasData = False
            For colIdx = 2 To memoRegion.Columns.Count
                v = memoRegion.Cells(rowIdx, colIdx).value
                If Not IsEmpty(v) And v <> "" Then
                    If content <> "" Then content = content & ";"
                    content = content & v
                    hasData = True
                End If
            Next colIdx
            If hasData Then
                Dim memoObj As CharacterMemo
                Set memoObj = New CharacterMemo
                memoObj.CharacterID = CharacterID
                memoObj.MemoType = memoName
                memoObj.Contents = content
                Character.CharacterMemoList.Add memoObj
            End If
        Next rowIdx
    Next memoName
    
    ' 3.��ȡAttacks/Spells/Equipments(�б�ͷ,������һ�к͵�һ��)
    Dim listNames As Collection, listName As Variant
    Set listNames = GetNamesByRegionName("RegionList")
    For Each listName In listNames
        Dim listCell As Range, listRegion As Range
        Set listCell = shGeneral.Range(listName)
        Set listRegion = listCell.CurrentRegion
        For rowIdx = 2 To listRegion.Rows.Count '������ͷ
            hasData = False
            For colIdx = 2 To listRegion.Columns.Count
                v = listRegion.Cells(rowIdx, colIdx).value
                If Not IsEmpty(v) And v <> "" Then hasData = True
            Next colIdx
            If hasData Then
                Select Case listName
                    Case "Attacks"
                        Dim atk As CharacterAttackSpell
                        Set atk = New CharacterAttackSpell
                        atk.CharacterID = CharacterID
                        atk.ItemType = "Attack"
                        atk.Name = listRegion.Cells(rowIdx, 2).value
                        atk.AtkBonus = listRegion.Cells(rowIdx, 3).value
                        atk.Damage_Type = listRegion.Cells(rowIdx, 4).value
                        atk.Equiped = ReadBoolean(listRegion.Cells(rowIdx, 5).value)
                        atk.Attuned = ReadBoolean(listRegion.Cells(rowIdx, 6).value)
                        Character.CharacterAttackSpellList.Add atk
                    Case "Spells"
                        Dim spl As CharacterAttackSpell
                        Set spl = New CharacterAttackSpell
                        spl.CharacterID = CharacterID
                        spl.ItemType = "Spell"
                        spl.Name = listRegion.Cells(rowIdx, 2).value
                        spl.AtkBonus = listRegion.Cells(rowIdx, 3).value
                        spl.Damage_Type = listRegion.Cells(rowIdx, 4).value
                        spl.SpellMemo = listRegion.Cells(rowIdx, 5).value
                        Character.CharacterAttackSpellList.Add spl
                    Case "Equipments"
                        Dim eq As CharacterEquipment
                        Set eq = New CharacterEquipment
                        eq.CharacterID = CharacterID
                        eq.ItemType = "Equipment"
                        eq.Name = listRegion.Cells(rowIdx, 2).value
                        eq.Quantity = listRegion.Cells(rowIdx, 3).value
                        eq.Attuned = ReadBoolean(listRegion.Cells(rowIdx, 4).value)
                        eq.Equiped = ReadBoolean(listRegion.Cells(rowIdx, 5).value)
                        Character.CharacterEquipmentList.Add eq
                End Select
            End If
        Next rowIdx
    Next listName
    
    ' 4.��ȡSpellList/SpellSlots(�б�ͷ,������һ�к͵�һ��)
    Dim spellNames As Collection, spellName As Variant
    Set spellNames = GetNamesByRegionName("RegionSpellList")
    For Each spellName In spellNames
        Dim spellCell As Range, spellRegion As Range
        Set spellCell = shGeneral.Range(spellName)
        Set spellRegion = spellCell.CurrentRegion
        For rowIdx = 2 To spellRegion.Rows.Count '������ͷ
            hasData = False
            For colIdx = 2 To spellRegion.Columns.Count
                v = spellRegion.Cells(rowIdx, colIdx).value
                If Not IsEmpty(v) And v <> "" Then hasData = True
            Next colIdx
            If hasData Then
                Select Case spellName
                    Case "SpellList"
                        Dim sp As CharacterSpell
                        Set sp = New CharacterSpell
                        sp.CharacterID = CharacterID
                        sp.SpellLevel = spellRegion.Cells(rowIdx, 2).value
                        sp.Name = spellRegion.Cells(rowIdx, 3).value
                        sp.Description = spellRegion.Cells(rowIdx, 4).value
                        sp.Prepared = ReadBoolean(spellRegion.Cells(rowIdx, 5).value)
                        Character.CharacterSpellList.Add sp
                    Case "SpellSlots"
                        Dim ss As CharacterSpellSlot
                        Set ss = New CharacterSpellSlot
                        ss.CharacterID = CharacterID
                        ss.SpellLevel = spellRegion.Cells(rowIdx, 2).value
                        ss.SlotsTotal = spellRegion.Cells(rowIdx, 3).value
                        ss.SlotsExpended = spellRegion.Cells(rowIdx, 4).value
                        Character.CharacterSpellSlots.Add ss
                End Select
            End If
        Next rowIdx
    Next spellName
    
    Set UIToCharacter = Character
End Function
