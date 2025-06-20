Option Explicit

' Global variables as defined in FRD
Public CharacterMasterList As Variant
Public CharacterMemoList As Variant
Public CharacterAttackSpellList As Variant
Public CharacterEquipmentList As Variant
Public Characters As Scripting.Dictionary

Private Function GetFieldColumn(ByVal TableName As String, ByVal FieldName As String) As Integer
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
        Case Else
            GetFieldColumn = 0
            Exit Function
    End Select
    
    ' Get the ListObject
    Set lo = ws.ListObjects(TableName)
    Set headerRange = lo.HeaderRange
    
    ' Find the column
    For Each cell In headerRange
        If cell.Value = FieldName Then
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
                .SavingThrowStrP = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SavingThrowStrP"))
                .SavingThrowDexP = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SavingThrowDexP"))
                .SavingThrowConP = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SavingThrowConP"))
                .SavingThrowIntP = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SavingThrowIntP"))
                .SavingThrowWisP = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SavingThrowWisP"))
                .SavingThrowChaP = CharacterMasterList(i, GetFieldColumn("CharacterMaster", "SavingThrowChaP"))
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
                    .Attuned = CharacterAttackSpellList(j, GetFieldColumn("CharacterAttackSpell", "Attuned"))
                    .Equiped = CharacterAttackSpellList(j, GetFieldColumn("CharacterAttackSpell", "Equiped"))
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
                    .Attuned = CharacterEquipmentList(j, GetFieldColumn("CharacterEquipment", "Attuned"))
                    .Equiped = CharacterEquipmentList(j, GetFieldColumn("CharacterEquipment", "Equiped"))
                End With
                Characters(charEquipment.CharacterID).pCharacterEquipmentList.Add charEquipment
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
    Dim foundRow As Range
    Dim searchRange As Range
    
    ' Write CharacterMaster data and its related data
    For Each charMaster In Characters.Items
        ' Write CharacterMaster data
        With shCharacterMaster.ListObject(1)
            ' Try to find existing record
            Set searchRange = .DataBodyRange
            Set foundRow = searchRange.Find(What:=charMaster.CharacterID, _
                                          LookIn:=xlValues, _
                                          LookAt:=xlWhole, _
                                          SearchOrder:=xlByRows, _
                                          SearchDirection:=xlNext, _
                                          MatchCase:=False)
            
            If foundRow Is Nothing Then
                ' Record doesn't exist, add new row
                .ListRows.Add
                masterRowIndex = .DataBodyRange.Rows.Count
            Else
                ' Record exists, use existing row
                masterRowIndex = foundRow.Row - .HeaderRowRange.Row
            End If
            
            ' Update all fields
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
            With shCharacterMemo.ListObject(1)
                ' Try to find existing record
                Set searchRange = .DataBodyRange
                Set foundRow = searchRange.Find(What:=charMemo.CharacterID, _
                                              LookIn:=xlValues, _
                                              LookAt:=xlWhole, _
                                              SearchOrder:=xlByRows, _
                                              SearchDirection:=xlNext, _
                                              MatchCase:=False)
                
                If foundRow Is Nothing Then
                    ' Record doesn't exist, add new row
                    .ListRows.Add
                    tmpRowIndex = .DataBodyRange.Rows.Count
                Else
                    ' Record exists, use existing row
                    tmpRowIndex = foundRow.Row - .HeaderRowRange.Row
                End If
                
                ' Update fields
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterMemo", "CharacterID")) = charMemo.CharacterID
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterMemo", "MemoType")) = charMemo.MemoType
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterMemo", "Contents")) = charMemo.Contents
            End With
        Next charMemo
        
        ' Write CharacterAttackSpell data for current CharacterMaster
        For Each charAttackSpell In charMaster.CharacterAttackSpellList
            With shCharacterAttackSpell.ListObject(1)
                ' Try to find existing record
                Set searchRange = .DataBodyRange
                Set foundRow = searchRange.Find(What:=charAttackSpell.CharacterID, _
                                              LookIn:=xlValues, _
                                              LookAt:=xlWhole, _
                                              SearchOrder:=xlByRows, _
                                              SearchDirection:=xlNext, _
                                              MatchCase:=False)
                
                If foundRow Is Nothing Then
                    ' Record doesn't exist, add new row
                    .ListRows.Add
                    tmpRowIndex = .DataBodyRange.Rows.Count
                Else
                    ' Record exists, use existing row
                    tmpRowIndex = foundRow.Row - .HeaderRowRange.Row
                End If
                
                ' Update fields
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
        For Each charEquipment In charMaster.pCharacterEquipmentList
            With shCharacterEquipment.ListObject(1)
                ' Try to find existing record
                Set searchRange = .DataBodyRange
                Set foundRow = searchRange.Find(What:=charEquipment.CharacterID, _
                                              LookIn:=xlValues, _
                                              LookAt:=xlWhole, _
                                              SearchOrder:=xlByRows, _
                                              SearchDirection:=xlNext, _
                                              MatchCase:=False)
                
                If foundRow Is Nothing Then
                    ' Record doesn't exist, add new row
                    .ListRows.Add
                    tmpRowIndex = .DataBodyRange.Rows.Count
                Else
                    ' Record exists, use existing row
                    tmpRowIndex = foundRow.Row - .HeaderRowRange.Row
                End If
                
                ' Update fields
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterEquipment", "CharacterID")) = charEquipment.CharacterID
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterEquipment", "Type")) = charEquipment.ItemType
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterEquipment", "Name")) = charEquipment.Name
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterEquipment", "Quantity")) = charEquipment.Quantity
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterEquipment", "Attuned")) = charEquipment.Attuned
                .ListRows(tmpRowIndex).Range(GetFieldColumn("CharacterEquipment", "Equiped")) = charEquipment.Equiped
            End With
        Next charEquipment
    Next charMaster
End Sub 