VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CharacterMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Private member variables
Private mCharacterID As Long
Private mCharacterType As String
Private mCharacterStatus As String
Private mPlayer As String
Private mCharacter As String
Private mBackground As String
Private mClass As String
Private mClassLv As Integer
Private mRace As String
Private mAlignment As String
Private mExp As Long
Private mStrength As Integer
Private mStrengthAdd As Integer
Private mDexterity As Integer
Private mDexterityAdd As Integer
Private mConstitution As Integer
Private mConstitutionAdd As Integer
Private mIntelligence As Integer
Private mIntelligenceAdd As Integer
Private mWisdom As Integer
Private mWisdomAdd As Integer
Private mCharisma As Integer
Private mCharismaAdd As Integer
Private mArmorClass As Integer
Private mInitiative As Integer
Private mSpeed As Integer
Private mInspiration As Integer
Private mProficiencyBonus As Integer
Private mSavingThrowStr As Integer
Private mSavingThrowDex As Integer
Private mSavingThrowCon As Integer
Private mSavingThrowInt As Integer
Private mSavingThrowWis As Integer
Private mSavingThrowCha As Integer
Private mSavingThrowStrP As Boolean
Private mSavingThrowDexP As Boolean
Private mSavingThrowConP As Boolean
Private mSavingThrowIntP As Boolean
Private mSavingThrowWisP As Boolean
Private mSavingThrowChaP As Boolean
Private mSkillAcrobatics As Integer
Private mSkillAnimalHandling As Integer
Private mSkillArcana As Integer
Private mSkillAthletics As Integer
Private mSkillDeception As Integer
Private mSkillHistory As Integer
Private mSkillInsight As Integer
Private mSkillIntimidation As Integer
Private mSkillInvestigation As Integer
Private mSkillMedicine As Integer
Private mSkillNature As Integer
Private mSkillPerception As Integer
Private mSkillPerformance As Integer
Private mSkillPersuasion As Integer
Private mSkillReligion As Integer
Private mSkillSleightOfHand As Integer
Private mSkillStealth As Integer
Private mSkillSurvival As Integer
Private mSkillAcrobaticsP As Boolean
Private mSkillAnimalHandlingP As Boolean
Private mSkillArcanaP As Boolean
Private mSkillAthleticsP As Boolean
Private mSkillDeceptionP As Boolean
Private mSkillHistoryP As Boolean
Private mSkillInsightP As Boolean
Private mSkillIntimidationP As Boolean
Private mSkillInvestigationP As Boolean
Private mSkillMedicineP As Boolean
Private mSkillNatureP As Boolean
Private mSkillPerceptionP As Boolean
Private mSkillPerformanceP As Boolean
Private mSkillPersuasionP As Boolean
Private mSkillReligionP As Boolean
Private mSkillSleightOfHandP As Boolean
Private mSkillStealthP As Boolean
Private mSkillSurvivalP As Boolean
Private mPassiveWisdom As Integer
Private mMaxHP As Integer
Private mCurHP As Integer
Private mTmpHP As Integer
Private mHD As Integer
Private mMaxHD As Integer
Private mMoneyCP As Integer
Private mMoneySP As Integer
Private mMoneyEP As Integer
Private mMoneyGP As Integer
Private mMoneyPP As Integer
Private mAge As String
Private mHeight As String
Private mWeight As String
Private mEyes As String
Private mSkin As String
Private mHair As String
Private mSpellCastingClass As String
Private mSpellCastingAbility As String
Private mSpellSaveDC As Integer
Private mSpellAttackBonus As Integer

' Collections for related objects
Private mCharacterMemoList As Collection
Private mCharacterAttackSpellList As Collection
Private mCharacterEquipmentList As Collection
Private mCharacterSpellList As Collection
Private mCharacterSpellSlots As Collection

' Properties
Public Property Get CharacterID() As Long
    CharacterID = mCharacterID
End Property

Public Property Let CharacterID(ByVal value As Long)
    mCharacterID = value
End Property

Public Property Get CharacterType() As String
    CharacterType = mCharacterType
End Property

Public Property Let CharacterType(ByVal value As String)
    mCharacterType = value
End Property

Public Property Get CharacterStatus() As String
    CharacterStatus = mCharacterStatus
End Property

Public Property Let CharacterStatus(ByVal value As String)
    mCharacterStatus = value
End Property

Public Property Get Player() As String
    Player = mPlayer
End Property

Public Property Let Player(ByVal value As String)
    mPlayer = value
End Property

Public Property Get Character() As String
    Character = mCharacter
End Property

Public Property Let Character(ByVal value As String)
    mCharacter = value
End Property

Public Property Get Background() As String
    Background = mBackground
End Property

Public Property Let Background(ByVal value As String)
    mBackground = value
End Property

Public Property Get Class() As String
    Class = mClass
End Property

Public Property Let Class(ByVal value As String)
    mClass = value
End Property

Public Property Get ClassLv() As Integer
    ClassLv = mClassLv
End Property

Public Property Let ClassLv(ByVal value As Integer)
    mClassLv = value
End Property

Public Property Get Race() As String
    Race = mRace
End Property

Public Property Let Race(ByVal value As String)
    mRace = value
End Property

Public Property Get Alignment() As String
    Alignment = mAlignment
End Property

Public Property Let Alignment(ByVal value As String)
    mAlignment = value
End Property

Public Property Get Exp() As Long
    Exp = mExp
End Property

Public Property Let Exp(ByVal value As Long)
    mExp = value
End Property

Public Property Get Strength() As Integer
    Strength = mStrength
End Property

Public Property Let Strength(ByVal value As Integer)
    mStrength = value
End Property

Public Property Get StrengthAdd() As Integer
    StrengthAdd = mStrengthAdd
End Property

Public Property Let StrengthAdd(ByVal value As Integer)
    mStrengthAdd = value
End Property

Public Property Get Dexterity() As Integer
    Dexterity = mDexterity
End Property

Public Property Let Dexterity(ByVal value As Integer)
    mDexterity = value
End Property

Public Property Get DexterityAdd() As Integer
    DexterityAdd = mDexterityAdd
End Property

Public Property Let DexterityAdd(ByVal value As Integer)
    mDexterityAdd = value
End Property

Public Property Get Constitution() As Integer
    Constitution = mConstitution
End Property

Public Property Let Constitution(ByVal value As Integer)
    mConstitution = value
End Property

Public Property Get ConstitutionAdd() As Integer
    ConstitutionAdd = mConstitutionAdd
End Property

Public Property Let ConstitutionAdd(ByVal value As Integer)
    mConstitutionAdd = value
End Property

Public Property Get Intelligence() As Integer
    Intelligence = mIntelligence
End Property

Public Property Let Intelligence(ByVal value As Integer)
    mIntelligence = value
End Property

Public Property Get IntelligenceAdd() As Integer
    IntelligenceAdd = mIntelligenceAdd
End Property

Public Property Let IntelligenceAdd(ByVal value As Integer)
    mIntelligenceAdd = value
End Property

Public Property Get Wisdom() As Integer
    Wisdom = mWisdom
End Property

Public Property Let Wisdom(ByVal value As Integer)
    mWisdom = value
End Property

Public Property Get WisdomAdd() As Integer
    WisdomAdd = mWisdomAdd
End Property

Public Property Let WisdomAdd(ByVal value As Integer)
    mWisdomAdd = value
End Property

Public Property Get Charisma() As Integer
    Charisma = mCharisma
End Property

Public Property Let Charisma(ByVal value As Integer)
    mCharisma = value
End Property

Public Property Get CharismaAdd() As Integer
    CharismaAdd = mCharismaAdd
End Property

Public Property Let CharismaAdd(ByVal value As Integer)
    mCharismaAdd = value
End Property

Public Property Get ArmorClass() As Integer
    ArmorClass = mArmorClass
End Property

Public Property Let ArmorClass(ByVal value As Integer)
    mArmorClass = value
End Property

Public Property Get Initiative() As Integer
    Initiative = mInitiative
End Property

Public Property Let Initiative(ByVal value As Integer)
    mInitiative = value
End Property

Public Property Get Speed() As Integer
    Speed = mSpeed
End Property

Public Property Let Speed(ByVal value As Integer)
    mSpeed = value
End Property

Public Property Get Inspiration() As Integer
    Inspiration = mInspiration
End Property

Public Property Let Inspiration(ByVal value As Integer)
    mInspiration = value
End Property

Public Property Get ProficiencyBonus() As Integer
    ProficiencyBonus = mProficiencyBonus
End Property

Public Property Let ProficiencyBonus(ByVal value As Integer)
    mProficiencyBonus = value
End Property

Public Property Get SavingThrowStr() As Integer
    SavingThrowStr = mSavingThrowStr
End Property

Public Property Let SavingThrowStr(ByVal value As Integer)
    mSavingThrowStr = value
End Property

Public Property Get SavingThrowDex() As Integer
    SavingThrowDex = mSavingThrowDex
End Property

Public Property Let SavingThrowDex(ByVal value As Integer)
    mSavingThrowDex = value
End Property

Public Property Get SavingThrowCon() As Integer
    SavingThrowCon = mSavingThrowCon
End Property

Public Property Let SavingThrowCon(ByVal value As Integer)
    mSavingThrowCon = value
End Property

Public Property Get SavingThrowInt() As Integer
    SavingThrowInt = mSavingThrowInt
End Property

Public Property Let SavingThrowInt(ByVal value As Integer)
    mSavingThrowInt = value
End Property

Public Property Get SavingThrowWis() As Integer
    SavingThrowWis = mSavingThrowWis
End Property

Public Property Let SavingThrowWis(ByVal value As Integer)
    mSavingThrowWis = value
End Property

Public Property Get SavingThrowCha() As Integer
    SavingThrowCha = mSavingThrowCha
End Property

Public Property Let SavingThrowCha(ByVal value As Integer)
    mSavingThrowCha = value
End Property

Public Property Get SavingThrowStrP() As Boolean
    SavingThrowStrP = mSavingThrowStrP
End Property

Public Property Let SavingThrowStrP(ByVal value As Boolean)
    mSavingThrowStrP = value
End Property

Public Property Get SavingThrowDexP() As Boolean
    SavingThrowDexP = mSavingThrowDexP
End Property

Public Property Let SavingThrowDexP(ByVal value As Boolean)
    mSavingThrowDexP = value
End Property

Public Property Get SavingThrowConP() As Boolean
    SavingThrowConP = mSavingThrowConP
End Property

Public Property Let SavingThrowConP(ByVal value As Boolean)
    mSavingThrowConP = value
End Property

Public Property Get SavingThrowIntP() As Boolean
    SavingThrowIntP = mSavingThrowIntP
End Property

Public Property Let SavingThrowIntP(ByVal value As Boolean)
    mSavingThrowIntP = value
End Property

Public Property Get SavingThrowWisP() As Boolean
    SavingThrowWisP = mSavingThrowWisP
End Property

Public Property Let SavingThrowWisP(ByVal value As Boolean)
    mSavingThrowWisP = value
End Property

Public Property Get SavingThrowChaP() As Boolean
    SavingThrowChaP = mSavingThrowChaP
End Property

Public Property Let SavingThrowChaP(ByVal value As Boolean)
    mSavingThrowChaP = value
End Property

Public Property Get SkillAcrobatics() As Integer
    SkillAcrobatics = mSkillAcrobatics
End Property

Public Property Let SkillAcrobatics(ByVal value As Integer)
    mSkillAcrobatics = value
End Property

Public Property Get SkillAnimalHandling() As Integer
    SkillAnimalHandling = mSkillAnimalHandling
End Property

Public Property Let SkillAnimalHandling(ByVal value As Integer)
    mSkillAnimalHandling = value
End Property

Public Property Get SkillArcana() As Integer
    SkillArcana = mSkillArcana
End Property

Public Property Let SkillArcana(ByVal value As Integer)
    mSkillArcana = value
End Property

Public Property Get SkillAthletics() As Integer
    SkillAthletics = mSkillAthletics
End Property

Public Property Let SkillAthletics(ByVal value As Integer)
    mSkillAthletics = value
End Property

Public Property Get SkillDeception() As Integer
    SkillDeception = mSkillDeception
End Property

Public Property Let SkillDeception(ByVal value As Integer)
    mSkillDeception = value
End Property

Public Property Get SkillHistory() As Integer
    SkillHistory = mSkillHistory
End Property

Public Property Let SkillHistory(ByVal value As Integer)
    mSkillHistory = value
End Property

Public Property Get SkillInsight() As Integer
    SkillInsight = mSkillInsight
End Property

Public Property Let SkillInsight(ByVal value As Integer)
    mSkillInsight = value
End Property

Public Property Get SkillIntimidation() As Integer
    SkillIntimidation = mSkillIntimidation
End Property

Public Property Let SkillIntimidation(ByVal value As Integer)
    mSkillIntimidation = value
End Property

Public Property Get SkillInvestigation() As Integer
    SkillInvestigation = mSkillInvestigation
End Property

Public Property Let SkillInvestigation(ByVal value As Integer)
    mSkillInvestigation = value
End Property

Public Property Get SkillMedicine() As Integer
    SkillMedicine = mSkillMedicine
End Property

Public Property Let SkillMedicine(ByVal value As Integer)
    mSkillMedicine = value
End Property

Public Property Get SkillNature() As Integer
    SkillNature = mSkillNature
End Property

Public Property Let SkillNature(ByVal value As Integer)
    mSkillNature = value
End Property

Public Property Get SkillPerception() As Integer
    SkillPerception = mSkillPerception
End Property

Public Property Let SkillPerception(ByVal value As Integer)
    mSkillPerception = value
End Property

Public Property Get SkillPerformance() As Integer
    SkillPerformance = mSkillPerformance
End Property

Public Property Let SkillPerformance(ByVal value As Integer)
    mSkillPerformance = value
End Property

Public Property Get SkillPersuasion() As Integer
    SkillPersuasion = mSkillPersuasion
End Property

Public Property Let SkillPersuasion(ByVal value As Integer)
    mSkillPersuasion = value
End Property

Public Property Get SkillReligion() As Integer
    SkillReligion = mSkillReligion
End Property

Public Property Let SkillReligion(ByVal value As Integer)
    mSkillReligion = value
End Property

Public Property Get SkillSleightOfHand() As Integer
    SkillSleightOfHand = mSkillSleightOfHand
End Property

Public Property Let SkillSleightOfHand(ByVal value As Integer)
    mSkillSleightOfHand = value
End Property

Public Property Get SkillStealth() As Integer
    SkillStealth = mSkillStealth
End Property

Public Property Let SkillStealth(ByVal value As Integer)
    mSkillStealth = value
End Property

Public Property Get SkillSurvival() As Integer
    SkillSurvival = mSkillSurvival
End Property

Public Property Let SkillSurvival(ByVal value As Integer)
    mSkillSurvival = value
End Property

Public Property Get SkillAcrobaticsP() As Boolean
    SkillAcrobaticsP = mSkillAcrobaticsP
End Property

Public Property Let SkillAcrobaticsP(ByVal value As Boolean)
    mSkillAcrobaticsP = value
End Property

Public Property Get SkillAnimalHandlingP() As Boolean
    SkillAnimalHandlingP = mSkillAnimalHandlingP
End Property

Public Property Let SkillAnimalHandlingP(ByVal value As Boolean)
    mSkillAnimalHandlingP = value
End Property

Public Property Get SkillArcanaP() As Boolean
    SkillArcanaP = mSkillArcanaP
End Property

Public Property Let SkillArcanaP(ByVal value As Boolean)
    mSkillArcanaP = value
End Property

Public Property Get SkillAthleticsP() As Boolean
    SkillAthleticsP = mSkillAthleticsP
End Property

Public Property Let SkillAthleticsP(ByVal value As Boolean)
    mSkillAthleticsP = value
End Property

Public Property Get SkillDeceptionP() As Boolean
    SkillDeceptionP = mSkillDeceptionP
End Property

Public Property Let SkillDeceptionP(ByVal value As Boolean)
    mSkillDeceptionP = value
End Property

Public Property Get SkillHistoryP() As Boolean
    SkillHistoryP = mSkillHistoryP
End Property

Public Property Let SkillHistoryP(ByVal value As Boolean)
    mSkillHistoryP = value
End Property

Public Property Get SkillInsightP() As Boolean
    SkillInsightP = mSkillInsightP
End Property

Public Property Let SkillInsightP(ByVal value As Boolean)
    mSkillInsightP = value
End Property

Public Property Get SkillIntimidationP() As Boolean
    SkillIntimidationP = mSkillIntimidationP
End Property

Public Property Let SkillIntimidationP(ByVal value As Boolean)
    mSkillIntimidationP = value
End Property

Public Property Get SkillInvestigationP() As Boolean
    SkillInvestigationP = mSkillInvestigationP
End Property

Public Property Let SkillInvestigationP(ByVal value As Boolean)
    mSkillInvestigationP = value
End Property

Public Property Get SkillMedicineP() As Boolean
    SkillMedicineP = mSkillMedicineP
End Property

Public Property Let SkillMedicineP(ByVal value As Boolean)
    mSkillMedicineP = value
End Property

Public Property Get SkillNatureP() As Boolean
    SkillNatureP = mSkillNatureP
End Property

Public Property Let SkillNatureP(ByVal value As Boolean)
    mSkillNatureP = value
End Property

Public Property Get SkillPerceptionP() As Boolean
    SkillPerceptionP = mSkillPerceptionP
End Property

Public Property Let SkillPerceptionP(ByVal value As Boolean)
    mSkillPerceptionP = value
End Property

Public Property Get SkillPerformanceP() As Boolean
    SkillPerformanceP = mSkillPerformanceP
End Property

Public Property Let SkillPerformanceP(ByVal value As Boolean)
    mSkillPerformanceP = value
End Property

Public Property Get SkillPersuasionP() As Boolean
    SkillPersuasionP = mSkillPersuasionP
End Property

Public Property Let SkillPersuasionP(ByVal value As Boolean)
    mSkillPersuasionP = value
End Property

Public Property Get SkillReligionP() As Boolean
    SkillReligionP = mSkillReligionP
End Property

Public Property Let SkillReligionP(ByVal value As Boolean)
    mSkillReligionP = value
End Property

Public Property Get SkillSleightOfHandP() As Boolean
    SkillSleightOfHandP = mSkillSleightOfHandP
End Property

Public Property Let SkillSleightOfHandP(ByVal value As Boolean)
    mSkillSleightOfHandP = value
End Property

Public Property Get SkillStealthP() As Boolean
    SkillStealthP = mSkillStealthP
End Property

Public Property Let SkillStealthP(ByVal value As Boolean)
    mSkillStealthP = value
End Property

Public Property Get SkillSurvivalP() As Boolean
    SkillSurvivalP = mSkillSurvivalP
End Property

Public Property Let SkillSurvivalP(ByVal value As Boolean)
    mSkillSurvivalP = value
End Property

Public Property Get PassiveWisdom() As Integer
    PassiveWisdom = mPassiveWisdom
End Property

Public Property Let PassiveWisdom(ByVal value As Integer)
    mPassiveWisdom = value
End Property

Public Property Get MaxHP() As Integer
    MaxHP = mMaxHP
End Property

Public Property Let MaxHP(ByVal value As Integer)
    mMaxHP = value
End Property

Public Property Get CurHP() As Integer
    CurHP = mCurHP
End Property

Public Property Let CurHP(ByVal value As Integer)
    mCurHP = value
End Property

Public Property Get TmpHP() As Integer
    TmpHP = mTmpHP
End Property

Public Property Let TmpHP(ByVal value As Integer)
    mTmpHP = value
End Property

Public Property Get HD() As Integer
    HD = mHD
End Property

Public Property Let HD(ByVal value As Integer)
    mHD = value
End Property

Public Property Get MaxHD() As Integer
    MaxHD = mMaxHD
End Property

Public Property Let MaxHD(ByVal value As Integer)
    mMaxHD = value
End Property

Public Property Get MoneyCP() As Integer
    MoneyCP = mMoneyCP
End Property

Public Property Let MoneyCP(ByVal value As Integer)
    mMoneyCP = value
End Property

Public Property Get MoneySP() As Integer
    MoneySP = mMoneySP
End Property

Public Property Let MoneySP(ByVal value As Integer)
    mMoneySP = value
End Property

Public Property Get MoneyEP() As Integer
    MoneyEP = mMoneyEP
End Property

Public Property Let MoneyEP(ByVal value As Integer)
    mMoneyEP = value
End Property

Public Property Get MoneyGP() As Integer
    MoneyGP = mMoneyGP
End Property

Public Property Let MoneyGP(ByVal value As Integer)
    mMoneyGP = value
End Property

Public Property Get MoneyPP() As Integer
    MoneyPP = mMoneyPP
End Property

Public Property Let MoneyPP(ByVal value As Integer)
    mMoneyPP = value
End Property

Public Property Get Age() As String
    Age = mAge
End Property

Public Property Let Age(ByVal value As String)
    mAge = value
End Property

Public Property Get Height() As String
    Height = mHeight
End Property

Public Property Let Height(ByVal value As String)
    mHeight = value
End Property

Public Property Get Weight() As String
    Weight = mWeight
End Property

Public Property Let Weight(ByVal value As String)
    mWeight = value
End Property

Public Property Get Eyes() As String
    Eyes = mEyes
End Property

Public Property Let Eyes(ByVal value As String)
    mEyes = value
End Property

Public Property Get Skin() As String
    Skin = mSkin
End Property

Public Property Let Skin(ByVal value As String)
    mSkin = value
End Property

Public Property Get Hair() As String
    Hair = mHair
End Property

Public Property Let Hair(ByVal value As String)
    mHair = value
End Property

Public Property Get SpellCastingClass() As String
    SpellCastingClass = mSpellCastingClass
End Property

Public Property Let SpellCastingClass(ByVal value As String)
    mSpellCastingClass = value
End Property

Public Property Get SpellCastingAbility() As String
    SpellCastingAbility = mSpellCastingAbility
End Property

Public Property Let SpellCastingAbility(ByVal value As String)
    mSpellCastingAbility = value
End Property

Public Property Get SpellSaveDC() As Integer
    SpellSaveDC = mSpellSaveDC
End Property

Public Property Let SpellSaveDC(ByVal value As Integer)
    mSpellSaveDC = value
End Property

Public Property Get SpellAttackBonus() As Integer
    SpellAttackBonus = mSpellAttackBonus
End Property

Public Property Let SpellAttackBonus(ByVal value As Integer)
    mSpellAttackBonus = value
End Property

' Collection Properties
Public Property Get CharacterMemoList() As Collection
    Set CharacterMemoList = mCharacterMemoList
End Property

Public Property Set CharacterMemoList(ByVal value As Collection)
    Set mCharacterMemoList = value
End Property

Public Property Get CharacterAttackSpellList() As Collection
    Set CharacterAttackSpellList = mCharacterAttackSpellList
End Property

Public Property Set CharacterAttackSpellList(ByVal value As Collection)
    Set mCharacterAttackSpellList = value
End Property

Public Property Get CharacterEquipmentList() As Collection
    Set CharacterEquipmentList = mCharacterEquipmentList
End Property

Public Property Set CharacterEquipmentList(ByVal value As Collection)
    Set mCharacterEquipmentList = value
End Property

Public Property Get CharacterSpellList() As Collection
    Set CharacterSpellList = mCharacterSpellList
End Property

Public Property Set CharacterSpellList(ByVal value As Collection)
    Set mCharacterSpellList = value
End Property

Public Property Get CharacterSpellSlots() As Collection
    Set CharacterSpellSlots = mCharacterSpellSlots
End Property

Public Property Set CharacterSpellSlots(ByVal value As Collection)
    Set mCharacterSpellSlots = value
End Property

Private Sub Class_Initialize()
    ' Initialize collections
    Set mCharacterMemoList = New Collection
    Set mCharacterAttackSpellList = New Collection
    Set mCharacterEquipmentList = New Collection
    Set mCharacterSpellList = New Collection
    Set mCharacterSpellSlots = New Collection
End Sub

Private Sub Class_Terminate()
    ' Clean up collections
    Set mCharacterMemoList = Nothing
    Set mCharacterAttackSpellList = Nothing
    Set mCharacterEquipmentList = Nothing
    Set mCharacterSpellList = Nothing
    Set mCharacterSpellSlots = Nothing
End Sub


'����ָ������ֵ,��SplitString�ָ�
Public Function ExportValues(ByVal SplitString As String, ParamArray PropNames() As Variant) As String
    Dim result As String
    Dim i As Integer
    For i = LBound(PropNames) To UBound(PropNames)
        If i > LBound(PropNames) Then result = result & SplitString
        result = result & CStr(CallByName(Me, PropNames(i), VbGet))
    Next i
    ExportValues = result
End Function

'����ָ�����͵�Memo����,�����зָ�
Public Function ExportMemoLists(ParamArray MemoTypes() As Variant) As String
    Dim result As String
    Dim memo As Object
    Dim i As Integer
    For Each memo In Me.CharacterMemoList
        For i = LBound(MemoTypes) To UBound(MemoTypes)
            If memo.MemoType = MemoTypes(i) Then
                If result <> "" Then result = result & vbCrLf
                result = result & memo.Contents
            End If
        Next i
    Next memo
    ExportMemoLists = result
End Function

'����װ��״̬ΪEquiped�Ĺ����б�,��Pattern��ʽ��
Public Function ExportAttackLists(ByVal Pattern As String, ByVal Equiped As Boolean) As String
    Dim result As String, line As String
    Dim atk As Object
    Dim arr() As String, widths() As Integer
    Dim i As Integer
    Dim valStr As String
    arr = Split(Pattern, "|")
    ReDim widths(UBound(arr))
    For i = 0 To UBound(arr)
        widths(i) = val(arr(i))
    Next i
    For Each atk In Me.CharacterAttackSpellList
        If atk.ItemType = "Attack" And atk.Equiped = Equiped Then
            Dim fields(4) As String
            fields(0) = Functions.PrintBoolean(atk.Equiped)
            fields(1) = atk.Name
            If atk.Attuned Then
                fields(2) = "(*)"
            Else
                fields(2) = ""
            End If
            fields(3) = Functions.PrintSignedNumber(atk.AtkBonus)
            fields(4) = atk.Damage_Type
            line = ""
            For i = 0 To 4
                valStr = fields(i)
                If widths(i) > 0 And Len(valStr) < widths(i) Then
                    valStr = valStr & Space(widths(i) - Len(valStr))
                End If
                line = line & valStr
            Next i
            If result <> "" Then result = result & vbCrLf
            result = result & line
        End If
    Next atk
    ExportAttackLists = result
End Function

'����������CharacterAtkSpellList�е�Spell,��Pattern��ʽ��
Public Function ExportAtkSpellLists(ByVal Pattern As String) As String
    Dim result As String, line As String
    Dim spl As Object
    Dim arr() As String, widths() As Integer
    Dim i As Integer
    Dim valStr As String
    arr = Split(Pattern, "|")
    ReDim widths(UBound(arr))
    For i = 0 To UBound(arr)
        widths(i) = val(arr(i))
    Next i
    For Each spl In Me.CharacterAttackSpellList
        If spl.ItemType = "Spell" Then
            Dim fields(3) As String
            fields(0) = Functions.PrintBoolean(spl.Equiped)
            fields(1) = spl.Name
            fields(2) = spl.Damage_Type
            fields(3) = spl.SpellMemo
            line = ""
            For i = 0 To 3
                valStr = fields(i)
                If widths(i) > 0 And Len(valStr) < widths(i) Then
                    valStr = valStr & Space(widths(i) - Len(valStr))
                End If
                line = line & valStr
            Next i
            If result <> "" Then result = result & vbCrLf
            result = result & line
        End If
    Next spl
    ExportAtkSpellLists = result
End Function

'����װ���б�,��Pattern��ʽ��
Public Function ExportEquipmntList(ByVal Pattern As String) As String
    Dim result As String, line As String
    Dim eq As Object
    Dim arr() As String, widths() As Integer
    Dim i As Integer
    Dim valStr As String
    arr = Split(Pattern, "|")
    ReDim widths(UBound(arr))
    For i = 0 To UBound(arr)
        widths(i) = val(arr(i))
    Next i
    For Each eq In Me.CharacterEquipmentList
        Dim fields(2) As String
        fields(0) = eq.Name
        If eq.Attuned Then
            fields(1) = "(*)"
        Else
            fields(1) = ""
        End If
        fields(2) = eq.Quantity
        line = ""
        For i = 0 To 2
            valStr = fields(i)
            If widths(i) > 0 And Len(valStr) < widths(i) Then
                valStr = valStr & Space(widths(i) - Len(valStr))
            End If
            line = line & valStr
        Next i
        If result <> "" Then result = result & vbCrLf
        result = result & line
    Next eq
    ExportEquipmntList = result
End Function

' 根据SpellLv返回SlotsTotal
Public Function SpellSlotByLevel(ByVal SpellLv As Integer) As String
    Dim slot As Object
    If Me.CharacterSpellSlots Is Nothing Or Me.CharacterSpellSlots.Count = 0 Then
        SpellSlotByLevel = ""
        Exit Function
    End If
    For Each slot In Me.CharacterSpellSlots
        If slot.SpellLevel = SpellLv Then
            SpellSlotByLevel = CStr(slot.SlotsTotal)
            Exit Function
        End If
    Next slot
    SpellSlotByLevel = ""
End Function

' 根据SpellLv将法术位以图形方式返回字符串
Public Function SpellSlotsInString(ByVal SpellLv As Integer) As String
    Dim slot As Object
    Dim i As Integer
    If Me.CharacterSpellSlots Is Nothing Or Me.CharacterSpellSlots.Count = 0 Then
        SpellSlotsInString = ""
        Exit Function
    End If
    For Each slot In Me.CharacterSpellSlots
        If slot.SpellLevel = SpellLv Then
            For i = 1 To slot.SlotsTotal
                SpellSlotsInString = SpellSlotsInString & ChrW(&H25EF)
            Next i
            Exit Function
        End If
    Next slot
    SpellSlotsInString = ""
End Function

' 根据输入的SpellLv导出法术列表
Public Function ExportSpellList(ByVal Pattern As String, ByVal SpellLv As Integer) As String
    Dim result As String, line As String
    Dim spl As Object
    Dim arr() As String, widths() As Integer
    Dim i As Integer
    Dim valStr As String
    arr = Split(Pattern, "|")
    ReDim widths(UBound(arr))
    For i = 0 To UBound(arr)
        widths(i) = Val(arr(i))
    Next i
    For Each spl In Me.CharacterSpellList
        If spl.SpellLevel = SpellLv Then
            Dim fields(2) As String
            ' 第一个字段: NeedPrep/Prepared 状态符号
            If Not spl.NeedPrep Then
                fields(0) = ChrW(&H2605) ' 实心五角星
            Else
                If spl.Prepared Then
                    fields(0) = ChrW(&H2022) ' 实心圆点
                Else
                    fields(0) = ChrW(&H25EF) ' 空心圆
                End If
            End If
            fields(1) = spl.Name
            fields(2) = spl.Description
            line = ""
            For i = 0 To 2
                valStr = fields(i)
                If widths(i) > 0 And Len(valStr) < widths(i) Then
                    valStr = valStr & Space(widths(i) - Len(valStr))
                End If
                line = line & valStr
            Next i
            If result <> "" Then result = result & vbCrLf
            result = result & line
        End If
    Next spl
    ExportSpellList = result
End Function

' End of class


