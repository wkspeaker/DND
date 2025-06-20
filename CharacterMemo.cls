VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CharacterMemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Character ID and Memo Type
Private pCharacterID As Long
Private pMemoType As String
Private pContents As String

' Property Let/Get for CharacterID
Public Property Let CharacterID(value As Long)
    pCharacterID = value
End Property
Public Property Get CharacterID() As Long
    CharacterID = pCharacterID
End Property

' Property Let/Get for MemoType
Public Property Let MemoType(value As String)
    pMemoType = value
End Property
Public Property Get MemoType() As String
    MemoType = pMemoType
End Property

' Property Let/Get for Contents
Public Property Let Contents(value As String)
    pContents = value
End Property
Public Property Get Contents() As String
    Contents = pContents
End Property
