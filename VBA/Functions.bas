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

' 辅助函数：删除表格中的所有数据行
Private Sub ClearTableRows(TableObj As ListObject)
    Do While TableObj.ListRows.Count > 0
        TableObj.ListRows(1).Delete
    Loop
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
    
    ' 清空所有相关表格
    Call ClearTableRows(shCharacterMaster.ListObjects(1))
    Call ClearTableRows(shCharacterMemo.ListObjects(1))
    Call ClearTableRows(shCharacterAttackSpell.ListObjects(1))
    Call ClearTableRows(shCharacterEquipment.ListObjects(1))
    Call ClearTableRows(shCharacterSpell.ListObjects(1))
    Call ClearTableRows(shCharacterSpellSlot.ListObjects(1))

    ' Initialize master row index
    masterRowIndex = 1
    
    ' Write CharacterMaster data and its related data
    Dim key As Variant
    For Each key In Characters.Keys
        Set charMaster = Characters(key)
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
    Next key
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
    Dim Prop As String
    Dim PropType As String
    Dim PropValue As Variant
    Dim RangeExists As Boolean
    Dim Response As VbMsgBoxResult
    
    ' Get properties from CharacterMasterSchema table
    Dim Properties As Variant
    Properties = GetPropertiesFromSchema("CharacterMasterSchema")
    Dim i As Long
    For i = 1 To UBound(Properties, 1)
        Prop = Properties(i, 1)      ' 字段名
        PropType = LCase(Trim(Properties(i, 2)))  ' 类型
        On Error Resume Next
        PropValue = CallByName(Character, Prop, VbGet)
        On Error GoTo 0
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
            If PropType = "bool" Or PropType = "boolean" Then
                shGeneral.Range(Prop) = WriteBoolean(PropValue)
            Else
                shGeneral.Range(Prop) = PropValue
            End If
        End If
    Next i
    
    ' 处理CharacterMemoList
    If Not Character.CharacterMemoList Is Nothing Then
        If Character.CharacterMemoList.Count > 0 Then
            Dim charMemo As CharacterMemo
            For Each charMemo In Character.CharacterMemoList
                Call WriteDataBlockByRange(charMemo.MemoType, charMemo.Contents)
            Next charMemo
        End If
    End If
    ' 处理CharacterAttackSpellList
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
    ' 处理CharacterEquipmentList
    If Not Character.CharacterEquipmentList Is Nothing Then
        Dim charEquipment As CharacterEquipment
        For Each charEquipment In Character.CharacterEquipmentList
            Call WriteEquipmentsByRange(charEquipment)
        Next charEquipment
    End If
    ' 处理CharacterSpellList
    If Not Character.CharacterSpellList Is Nothing Then
        Dim charSpell As CharacterSpell
        For Each charSpell In Character.CharacterSpellList
            Call WriteSpellListByRange(charSpell)
        Next charSpell
    End If
    ' 处理CharacterSpellSlots
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
    Set FieldRange = SchemaTable.ListColumns("字段").DataBodyRange
    Set TypeRange = SchemaTable.ListColumns("类型").DataBodyRange
    
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
    
    Dim key As Variant
    For Each key In Characters.Keys
        If CLng(key) > MaxId Then
            MaxId = CLng(key)
        End If
    Next key
    
    GetMaxCharacterId = MaxId
End Function

' 将Excel单元格地址转换为Boolean
Public Function ReadBoolean(ByVal value As Variant) As Boolean
    Dim s As String
    s = UCase(Trim(CStr(value)))
    Select Case s
        Case "Y", "YES"
            ReadBoolean = True
        Case "N", "NO", ""
            ReadBoolean = False
        Case Else
            ReadBoolean = False ' 返回异常/错误
    End Select
End Function

' 将Boolean值转换为Excel可用的字符串
Public Function WriteBoolean(ByVal value As Boolean) As String
    If value Then
        WriteBoolean = "Y"
    Else
        WriteBoolean = "N"
    End If
End Function

'根据CharacterIDName获取CharacterID
Public Function GetCharacterIDFromCharacterIDName(ByVal CharacterIDName As String) As Long
    Dim pos As Long
    pos = InStr(CharacterIDName, "|")
    If pos > 0 Then
        GetCharacterIDFromCharacterIDName = CLng(Trim(Left(CharacterIDName, pos - 1)))
    Else
        GetCharacterIDFromCharacterIDName = 0
    End If
End Function

'新增Character，CharacterID取最大值加1，返回新增的CharacterID
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

'滚动到shGeneral页指定区域的第二行中间
Public Sub ScrollToRow(ByVal TargetRegion As String)
    Dim freezeRow As Long
    Dim targetRow As Long
    
    targetRow = shGeneral.Range(TargetRegion).Row
    freezeRow = ActiveWindow.SplitRow
    If freezeRow < 1 Then freezeRow = 0
    '选择目标单元格
    shGeneral.Cells(targetRow, 1).Select
    '滚动，使目标单元格显示在第二行中间
    ActiveWindow.ScrollRow = targetRow
End Sub

'在shGeneral页指定区域预处理数据块
Public Sub PrepDataBlockByRange(ByVal TargetName As String, Optional HasHeadLine As Boolean = False)
    Dim rng As Range
    Dim colCount As Long
    
    '定位当前区域CurrentRegion
    On Error Resume Next
    Set rng = shGeneral.Range(TargetName).CurrentRegion
    On Error GoTo 0
    
    If rng Is Nothing Then
        MsgBox "Named range '" & TargetName & "' not found.", vbExclamation, "Error"
        Exit Sub
    End If
    
    colCount = rng.Columns.Count
    If colCount > 1 Then
        '如果HasHeadLine为True，则清除从第二行开始的数据
        If HasHeadLine Then
            rng.Offset(1, 1).Resize(rng.Rows.Count - 1, colCount - 1).ClearContents
        Else
            rng.Offset(0, 1).Resize(rng.Rows.Count, colCount - 1).ClearContents
        End If
    End If
End Sub

'在shGeneral页指定区域写入数据块
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
    
    '从末尾空单元格开始,写入下空单元格
    i = 0
    Do
        Set writeCell = targetCell.Offset(i, 1)
        If IsEmpty(writeCell.value) Then
            writeCell.value = content
            Exit Sub
        End If
        i = i + 1
        '终止循环,最多写入50行
        If i > 50 Then Exit Do
    Loop
    '如果没有空单元格,则写入
End Sub

'写入攻击信息到shGeneral页
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

'写入法术信息到shGeneral页
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

'写入装备信息到shGeneral页
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

'写入法术列表到shGeneral页
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

'写入法术位置到shGeneral页
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


'预处理数据块
Public Sub PrepDataBlocksBetweenNames(ByVal StartName As String, Optional HasHeadLine As Boolean = False)
    Dim startCell As Range
    ' 1.定位实际单元格
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

    ' 2.查找下一行连续的单元格,最多100行
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

    ' 3.获取区域(开始行+1 到 结束行-1)
    Dim regionStart As Long, regionEnd As Long
    regionStart = startRow + 1
    regionEnd = endRow - 1
    If regionStart > regionEnd Then Exit Sub '无效

    '只处理连续的单元格
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


'获取指定区域的所有单元格名称
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



'从shGeneral获取数据到Character
Public Function UIToCharacter(ByVal CharacterID As Long) As CharacterMaster
    Dim Character As New CharacterMaster
    Dim Properties As Variant
    Dim v As Variant
    Dim i As Long
    Dim Prop As String, PropType As String
    
    ' 1.获取属性
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
    
    ' 2.获取CharacterMemoList(备注头,可以多行)
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
    
    ' 3.获取Attacks/Spells/Equipments(备注头,第一行和第一列)
    Dim listNames As Collection, listName As Variant
    Set listNames = GetNamesByRegionName("RegionList")
    For Each listName In listNames
        Dim listCell As Range, listRegion As Range
        Set listCell = shGeneral.Range(listName)
        Set listRegion = listCell.CurrentRegion
        For rowIdx = 2 To listRegion.Rows.Count '去掉头
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
    
    ' 4.获取SpellList/SpellSlots(备注头,第一行和第一列)
    Dim spellNames As Collection, spellName As Variant
    Set spellNames = GetNamesByRegionName("RegionSpellList")
    For Each spellName In spellNames
        Dim spellCell As Range, spellRegion As Range
        Set spellCell = shGeneral.Range(spellName)
        Set spellRegion = spellCell.CurrentRegion
        For rowIdx = 2 To spellRegion.Rows.Count '去掉头
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

Public Sub TestFillWordContentControl()
    Dim wordApp As Word.Application
    Dim wordDoc As Word.Document
    Dim fso As Object
    Dim srcPath As String, newPath As String, fileName As String, fileExt As String
    Dim dateStr As String
    Dim docFolder As String
    Dim cc As Word.ContentControl
    
    ' 1. 读取Word文件名
    fileName = shGeneral.Range("CharacterSheet_SpellList").Value
    If fileName = "" Then
        MsgBox "未指定Word文件名！", vbExclamation
        Exit Sub
    End If
    
    ' 2. 构造完整路径
    docFolder = ThisWorkbook.Path & "\Documents\"
    If Right(docFolder, 1) <> "\" Then docFolder = docFolder & "\"
    srcPath = docFolder & fileName
    
    If Dir(srcPath) = "" Then
        MsgBox "找不到模板文件：" & srcPath, vbExclamation
        Exit Sub
    End If
    
    ' 3. 复制为新文件（加日期）
    fileExt = Mid(fileName, InStrRev(fileName, "."))
    dateStr = Format(Now, "yyyymmdd_HHMMSS")
    newPath = docFolder & Left(fileName, InStrRev(fileName, ".") - 1) & "_" & dateStr & fileExt
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile srcPath, newPath, True
    
    ' 4. 打开新文件
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = True ' 调试时可见
    Set wordDoc = wordApp.Documents.Open(newPath)
    
    ' 5. 填充内容控件
    For Each cc In wordDoc.ContentControls
        If cc.Tag = "CharacterName" Then
            cc.Range.Text = "测试内容"
        End If
    Next
    
    ' 6. 保存并关闭
    wordDoc.Save
    wordDoc.Close
    wordApp.Quit
    
    Set wordDoc = Nothing
    Set wordApp = Nothing
    Set fso = Nothing
    
    MsgBox "测试完成，新文件已生成：" & newPath, vbInformation
End Sub

Public Sub TestFillWordContentControlFromTemplate()
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim templatePath As String, fileName As String
    Dim docFolder As String
    Dim cc As Object
    Dim shp As Object

    ' 1. 读取Word文件名
    fileName = shGeneral.Range("CharacterSheet_SpellList").Value
    If fileName = "" Then
        MsgBox "未指定Word文件名！", vbExclamation
        Exit Sub
    End If

    ' 2. 构造完整路径
    docFolder = ThisWorkbook.Path & "\Documents\"
    If Right(docFolder, 1) <> "\" Then docFolder = docFolder & "\"
    templatePath = docFolder & fileName

    If Dir(templatePath) = "" Then
        MsgBox "找不到模板文件：" & templatePath, vbExclamation
        Exit Sub
    End If

    ' 3. 用模板新建文档
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = True
    Set wordDoc = wordApp.Documents.Add(Template:=templatePath, NewTemplate:=False)

    ' 4. 填充内容控件（正文）
    For Each cc In wordDoc.ContentControls
        If cc.Tag = "CharacterName" Then
            cc.Range.Text = "测试内容"
        End If
    Next
    ' 4b. 填充内容控件（文本框等Shape中）
    For Each shp In wordDoc.Shapes
        If shp.Type = 17 Then ' msoTextBox = 17
            For Each cc In shp.TextFrame.TextRange.ContentControls
                If cc.Tag = "CharacterName" Then
                    cc.Range.Text = "测试内容"
                End If
            Next
        End If
    Next

    ' 5. 不保存也不关闭文档，保持Word窗口打开
    ' 6. 释放对象
    Set wordDoc = Nothing
    Set wordApp = Nothing
End Sub

Public Function CreateWordFromTemplate(ByVal TemplateName As String) As Object
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim templatePath As String, fileName As String
    Dim docFolder As String
    
    ' 1. 读取Word文件名（从shNote或Worksheets("Note")）
    On Error Resume Next
    fileName = shNote.Range(TemplateName).Value
    If fileName = "" Then
        fileName = Worksheets("Note").Range(TemplateName).Value
    End If
    On Error GoTo 0
    If fileName = "" Then
        MsgBox "未指定Word文件名！", vbExclamation
        Set CreateWordFromTemplate = Nothing
        Exit Function
    End If
    
    ' 2. 构造完整路径
    docFolder = ThisWorkbook.Path & "\Documents\"
    If Right(docFolder, 1) <> "\" Then docFolder = docFolder & "\"
    templatePath = docFolder & fileName
    
    If Dir(templatePath) = "" Then
        MsgBox "找不到模板文件：" & templatePath, vbExclamation
        Set CreateWordFromTemplate = Nothing
        Exit Function
    End If
    
    ' 3. 用模板新建文档
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = True
    Set wordDoc = wordApp.Documents.Add(Template:=templatePath, NewTemplate:=False)
    
    ' 4. 返回文档对象
    Set CreateWordFromTemplate = wordDoc
End Function

Public Sub PrintToWord(ByRef TargetDoc As Object, ByVal TargetTag As String, ByVal ContentText As String)
    Dim cc As Object
    Dim shp As Object
    ' 1. 正文中的ContentControl
    For Each cc In TargetDoc.ContentControls
        If cc.Tag = TargetTag Then
            cc.Range.Text = ContentText
        End If
    Next
    ' 2. Shapes（如文本框）中的ContentControl
    For Each shp In TargetDoc.Shapes
        If shp.Type = 17 Then ' msoTextBox = 17
            For Each cc In shp.TextFrame.TextRange.ContentControls
                If cc.Tag = TargetTag Then
                    cc.Range.Text = ContentText
                End If
            Next
        End If
    Next
End Sub

Public Sub ExportCharacter()
    Dim CharacterID As Variant
    Dim Character As CharacterMaster
    Dim wordDoc As Object
    
    ' 1. 获取当前角色ID
    CharacterID = shGeneral.Range("CharacterID").Value
    If Not IsNumeric(CharacterID) Or IsEmpty(CharacterID) Then
        MsgBox "无效的角色ID！", vbExclamation
        Exit Sub
    End If
    
    ' 2. 获取当前角色对象
    Set Character = UIToCharacter(CLng(CharacterID))
    If Character Is Nothing Then
        MsgBox "未能获取角色对象！", vbExclamation
        Exit Sub
    End If
    
    ' 3. 创建Word文档（模板名可根据实际情况调整）
    Set wordDoc = CreateWordFromTemplate("CharacterSheet")
    If wordDoc Is Nothing Then
        MsgBox "Word文档创建失败！", vbExclamation
        Exit Sub
    End If
    
    ' 4. 写入角色名到Word文档Tag为"Character"的ContentControl
    Call PrintToWord(wordDoc, "Character", Character.Character)
    
    ' TODO: 后续可补充写入更多内容
    MsgBox "角色导出完成，后续内容请补充PrintToWord调用。", vbInformation
End Sub

Public Function PrintSignedNumber(ByVal num As Integer) As String
    If num > 0 Then
        PrintSignedNumber = "+" & CStr(num)
    ElseIf num < 0 Then
        PrintSignedNumber = CStr(num)
    Else
        PrintSignedNumber = ""
    End If
End Function

Public Function PrintBoolean(ByVal val As Boolean) As String
    If val = True Then
        PrintBoolean = ChrW(&H2022)
    Else
        PrintBoolean = ""
    End If
End Function
