Public Sub CharacterToUI(ByVal CharacterID As Long)
    ' Check if CharacterID exists in dictionary
    If Not Characters.Exists(CharacterID) Then
        MsgBox "CharacterID " & CharacterID & " does not exist!", vbExclamation
        Exit Sub
    End If
    
    ' Get the character object
    Dim character As CharacterMaster
    Set character = Characters(CharacterID)
    
    ' Write basic properties to UI
    shGeneral.Range("Player") = character.Player
    shGeneral.Range("Character") = character.Character
    shGeneral.Range("Background") = character.Background
    shGeneral.Range("Class") = character.Class
    shGeneral.Range("ClassLv") = character.ClassLv
    shGeneral.Range("Race") = character.Race
    shGeneral.Range("Alignment") = character.Alignment
    
    ' Write ability scores
    shGeneral.Range("Strength") = character.Strength
    shGeneral.Range("StrengthAdd") = character.StrengthAdd
    shGeneral.Range("Dexterity") = character.Dexterity
    shGeneral.Range("DexterityAdd") = character.DexterityAdd
    shGeneral.Range("Constitution") = character.Constitution
    shGeneral.Range("ConstitutionAdd") = character.ConstitutionAdd
    shGeneral.Range("Intelligence") = character.Intelligence
    shGeneral.Range("IntelligenceAdd") = character.IntelligenceAdd
    shGeneral.Range("Wisdom") = character.Wisdom
    shGeneral.Range("WisdomAdd") = character.WisdomAdd
    shGeneral.Range("Charisma") = character.Charisma
    shGeneral.Range("CharismaAdd") = character.CharismaAdd
    
    ' Write combat stats
    shGeneral.Range("ArmorClass") = character.ArmorClass
    shGeneral.Range("Initiative") = character.Initiative
    shGeneral.Range("Speed") = character.Speed
    shGeneral.Range("ProficiencyBonus") = character.ProficiencyBonus
    
    ' Write saving throws
    shGeneral.Range("SavingThrowStr") = character.SavingThrowStr
    shGeneral.Range("SavingThrowDex") = character.SavingThrowDex
    shGeneral.Range("SavingThrowCon") = character.SavingThrowCon
    shGeneral.Range("SavingThrowInt") = character.SavingThrowInt
    shGeneral.Range("SavingThrowWis") = character.SavingThrowWis
    shGeneral.Range("SavingThrowCha") = character.SavingThrowCha
    
    ' Write saving throw proficiencies
    shGeneral.Range("SavingThrowStrP") = character.SavingThrowStrP
    shGeneral.Range("SavingThrowDexP") = character.SavingThrowDexP
    shGeneral.Range("SavingThrowConP") = character.SavingThrowConP
    shGeneral.Range("SavingThrowIntP") = character.SavingThrowIntP
    shGeneral.Range("SavingThrowWisP") = character.SavingThrowWisP
    shGeneral.Range("SavingThrowChaP") = character.SavingThrowChaP
    
    ' Write skills
    shGeneral.Range("SkillAcrobatics") = character.SkillAcrobatics
    shGeneral.Range("SkillAnimalHandling") = character.SkillAnimalHandling
    shGeneral.Range("SkillArcana") = character.SkillArcana
    shGeneral.Range("SkillAthletics") = character.SkillAthletics
    shGeneral.Range("SkillDeception") = character.SkillDeception
    shGeneral.Range("SkillHistory") = character.SkillHistory
    shGeneral.Range("SkillInsight") = character.SkillInsight
    shGeneral.Range("SkillIntimidation") = character.SkillIntimidation
    shGeneral.Range("SkillInvestigation") = character.SkillInvestigation
    shGeneral.Range("SkillMedicine") = character.SkillMedicine
    shGeneral.Range("SkillNature") = character.SkillNature
    shGeneral.Range("SkillPerception") = character.SkillPerception
    shGeneral.Range("SkillPerformance") = character.SkillPerformance
    shGeneral.Range("SkillPersuasion") = character.SkillPersuasion
    shGeneral.Range("SkillReligion") = character.SkillReligion
    shGeneral.Range("SkillSleightOfHand") = character.SkillSleightOfHand
    shGeneral.Range("SkillStealth") = character.SkillStealth
    shGeneral.Range("SkillSurvival") = character.SkillSurvival
    
    ' Write skill proficiencies
    shGeneral.Range("SkillAcrobaticsP") = character.SkillAcrobaticsP
    shGeneral.Range("SkillAnimalHandlingP") = character.SkillAnimalHandlingP
    shGeneral.Range("SkillArcanaP") = character.SkillArcanaP
    shGeneral.Range("SkillAthleticsP") = character.SkillAthleticsP
    shGeneral.Range("SkillDeceptionP") = character.SkillDeceptionP
    shGeneral.Range("SkillHistoryP") = character.SkillHistoryP
    shGeneral.Range("SkillInsightP") = character.SkillInsightP
    shGeneral.Range("SkillIntimidationP") = character.SkillIntimidationP
    shGeneral.Range("SkillInvestigationP") = character.SkillInvestigationP
    shGeneral.Range("SkillMedicineP") = character.SkillMedicineP
    shGeneral.Range("SkillNatureP") = character.SkillNatureP
    shGeneral.Range("SkillPerceptionP") = character.SkillPerceptionP
    shGeneral.Range("SkillPerformanceP") = character.SkillPerformanceP
    shGeneral.Range("SkillPersuasionP") = character.SkillPersuasionP
    shGeneral.Range("SkillReligionP") = character.SkillReligionP
    shGeneral.Range("SkillSleightOfHandP") = character.SkillSleightOfHandP
    shGeneral.Range("SkillStealthP") = character.SkillStealthP
    shGeneral.Range("SkillSurvivalP") = character.SkillSurvivalP
    
    ' Write other stats
    shGeneral.Range("PassiveWisdom") = character.PassiveWisdom
    shGeneral.Range("MaxHP") = character.MaxHP
    shGeneral.Range("CurHP") = character.CurHP
    shGeneral.Range("TmpHP") = character.TmpHP
    shGeneral.Range("HD") = character.HD
    shGeneral.Range("MaxHD") = character.MaxHD
    
    ' Write money
    shGeneral.Range("MoneyCP") = character.MoneyCP
    shGeneral.Range("MoneySP") = character.MoneySP
    shGeneral.Range("MoneyEP") = character.MoneyEP
    shGeneral.Range("MoneyGP") = character.MoneyGP
    shGeneral.Range("MoneyPP") = character.MoneyPP
    
    ' Write physical characteristics
    shGeneral.Range("Age") = character.Age
    shGeneral.Range("Height") = character.Height
    shGeneral.Range("Weight") = character.Weight
    shGeneral.Range("Eyes") = character.Eyes
    shGeneral.Range("Skin") = character.Skin
    shGeneral.Range("Hair") = character.Hair
End Sub 