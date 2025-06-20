VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CharacterAttackSpell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Character ID and Item Information
Private pCharacterID As Long
Private pItemType As String
Private pName As String
Private pAtkBonus As Integer
Private pDamage_Type As String
Private pSpellMemo As String
Private pAttuned As Boolean
Private pEquiped As Boolean

' Property Let/Get for CharacterID
Public Property Let CharacterID(value As Long)
    pCharacterID = value
End Property
Public Property Get CharacterID() As Long
    CharacterID = pCharacterID
End Property

' Property Let/Get for ItemType
Public Property Let ItemType(value As String)
    pItemType = value
End Property
Public Property Get ItemType() As String
    ItemType = pItemType
End Property

' Property Let/Get for Name
Public Property Let Name(value As String)
    pName = value
End Property
Public Property Get Name() As String
    Name = pName
End Property

' Property Let/Get for AtkBonus
Public Property Let AtkBonus(value As Integer)
    pAtkBonus = value
End Property
Public Property Get AtkBonus() As Integer
    AtkBonus = pAtkBonus
End Property

' Property Let/Get for Damage_Type
Public Property Let Damage_Type(value As String)
    pDamage_Type = value
End Property
Public Property Get Damage_Type() As String
    Damage_Type = pDamage_Type
End Property

' Property Let/Get for SpellMemo
Public Property Let SpellMemo(value As String)
    pSpellMemo = value
End Property
Public Property Get SpellMemo() As String
    SpellMemo = pSpellMemo
End Property

' Property Let/Get for Attuned
Public Property Let Attuned(value As Boolean)
    pAttuned = value
End Property
Public Property Get Attuned() As Boolean
    Attuned = pAttuned
End Property

' Property Let/Get for Equiped
Public Property Let Equiped(value As Boolean)
    pEquiped = value
End Property
Public Property Get Equiped() As Boolean
    Equiped = pEquiped
End Property
