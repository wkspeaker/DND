VERSION 1.0 CLASS
Attribute VB_Name = "CharacterMasterSchema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Collections for related objects
Private pCharacterMemoList As Collection
Private pCharacterAttackSpellList As Collection
Private pCharacterEquipmentList As Collection

' Character ID and Type
Private pCharacterID As Long

' Property Let/Get for CharacterMemoList
Public Property Let CharacterMemoList(value As Collection)
    Set pCharacterMemoList = value
End Property
Public Property Get CharacterMemoList() As Collection
    If pCharacterMemoList Is Nothing Then
        Set pCharacterMemoList = New Collection
    End If
    Set CharacterMemoList = pCharacterMemoList
End Property

' Property Let/Get for CharacterAttackSpellList
Public Property Let CharacterAttackSpellList(value As Collection)
    Set pCharacterAttackSpellList = value
End Property
Public Property Get CharacterAttackSpellList() As Collection
    If pCharacterAttackSpellList Is Nothing Then
        Set pCharacterAttackSpellList = New Collection
    End If
    Set CharacterAttackSpellList = pCharacterAttackSpellList
End Property

' Property Let/Get for CharacterEquipmentList
Public Property Let CharacterEquipmentList(value As Collection)
    Set pCharacterEquipmentList = value
End Property
Public Property Get CharacterEquipmentList() As Collection
    If pCharacterEquipmentList Is Nothing Then
        Set pCharacterEquipmentList = New Collection
    End If
    Set CharacterEquipmentList = pCharacterEquipmentList
End Property

' Class Initialize event
Private Sub Class_Initialize()
    Set pCharacterMemoList = New Collection
    Set pCharacterAttackSpellList = New Collection
    Set pCharacterEquipmentList = New Collection
End Sub

' Class Terminate event
Private Sub Class_Terminate()
    Set pCharacterMemoList = Nothing
    Set pCharacterAttackSpellList = Nothing
    Set pCharacterEquipmentList = Nothing
End Sub 