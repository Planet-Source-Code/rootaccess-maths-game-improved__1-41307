Attribute VB_Name = "Module1"
'Play sound api
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Wins% 'As integer (% = Integer)
Public Loses%
Public Times%
Public Result As Variant
Public Function PercentWins(Wins, TimesTried) As String
    Percent = Wins / TimesTried 'Function to Calculate the Percent of Wins
    PercentWins = Format(Percent, "0.0%")
End Function
