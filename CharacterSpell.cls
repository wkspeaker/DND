VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CharacterSpell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Private member variables
Private pCharacterID As Long
Private pSpellLevel As Integer
Private pName As String
Private pDescription As String
Private pPrepared As Boolean

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

' Property Let/Get for Name
Public Property Let Name(value As String)
    pName = value
End Property
Public Property Get Name() As String
    Name = pName
End Property

' Property Let/Get for Description
Public Property Let Description(value As String)
    pDescription = value
End Property
Public Property Get Description() As String
    Description = pDescription
End Property

' Property Let/Get for Prepared
Public Property Let Prepared(value As Boolean)
    pPrepared = value
End Property
Public Property Get Prepared() As Boolean
    Prepared = pPrepared
End Property 