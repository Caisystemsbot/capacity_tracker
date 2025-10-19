VERSION 1.0 CLASS
Attribute VB_Name = "clsTeam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mTeamName As String
Private mHoursPerPoint As Double

Public Property Get TeamName() As String
    TeamName = mTeamName
End Property

Public Property Let TeamName(ByVal value As String)
    mTeamName = value
End Property

Public Property Get HoursPerPoint() As Double
    HoursPerPoint = mHoursPerPoint
End Property

Public Property Let HoursPerPoint(ByVal value As Double)
    mHoursPerPoint = value
End Property

