VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CharacterSpellSlot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Private member variables
Private pCharacterID As Long
Private pSpellLevel As Integer
Private pSlotsTotal As Integer
Private pSlotsExpended As Integer

' Property Let/Get for CharacterID
Public Property Let CharacterID(value As Long)
    pCharacterID = value
End Property
Public Property Get CharacterID() As Long
    CharacterID = pCharacterID
End Property

' Property Let/Get for SpellLevel
Public Property Let SpellLevel(value As Integer)
    pSpellLevel = value
End Property
Public Property Get SpellLevel() As Integer
    SpellLevel = pSpellLevel
End Property

' Property Let/Get for SlotsTotal
Public Property Let SlotsTotal(value As Integer)
    pSlotsTotal = value
End Property
Public Property Get SlotsTotal() As Integer
    SlotsTotal = pSlotsTotal
End Property

' Property Let/Get for SlotsExpended
Public Property Let SlotsExpended(value As Integer)
    pSlotsExpended = value
End Property
Public Property Get SlotsExpended() As Integer
    SlotsExpended = pSlotsExpended
End Property 