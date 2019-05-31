VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ECNLIB04_EFFECTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub LetrasEnCaidaDelluvias(ByVal sMsg As String, _
                                  ByRef xPctBox As Object, _
                                  Optional ByVal Rinicio As Integer = 80)
    Randomize Timer                         'Init Rnd

    'Declarations
    Dim StartTime(100)                      'Starttime of a up/down movement
    Dim DownMovement(100) As Boolean        'are we doing a up or down movement ???
    Dim MoveDistance As Double              'distance target has moved since the start of the movement
    Dim YPos(100) As Double                 'Holds the y position of a letter
    Dim MovementDone(100) As Boolean        'Is set to true when a up / down movement is completed
    Dim StartHeight(100) As Double               'From which hight will the letter fall down ?
    Dim UpMovementTime(100) As Double            'How long will it the letter take to move up
    Dim PowerLoss(100) As Double                 'losing xx% of power when touching the ground
    Dim Message As String                   'Message you want to display
    Dim Looop As Integer                    'Loop var
    Dim TextColor(100) As ColorConstants    'Color of one letter
    
    'Settings
    
    xPctBox.ScaleMode = 4
    'xPctBox.FontName = "Courier New"
    
    Message = sMsg
    
    For Looop = 1 To Len(Message)
    
        PowerLoss(Looop) = 0.2 + ((Rnd * 25) / 100)                  'losing xx% of power when touching the ground
        StartHeight(Looop) = 0
        TextColor(Looop) = RGB(Rinicio + Looop * 2, Rinicio + Looop * 2, 0)
    
    Next Looop
        
    For Looop = 1 To Len(Message)
        StartTime(Looop) = Timer                       'Setting up startime for a following movement, needed for calculation of position
    Next Looop
    
    Do
        
        xPctBox.Cls                             'Clear picturebox
        
        'Looping throung the textmessage
        For Looop = 1 To Len(Message)
        
        
            If DownMovement(Looop) = True Then
                
                MoveDistance = (StartHeight(Looop) + (0.5 * 9.81 * ((Timer - StartTime(Looop)) ^ 2))) 'Calculating falling distance
                
                If YPos(Looop) >= xPctBox.ScaleHeight - 1 Then MovementDone(Looop) = True     'The letter reached the bottom border. The Downmovement is complete
        
            Else
                MoveDistance = (StartHeight(Looop) + (0.5 * 9.81 * (UpMovementTime(Looop) - (Timer - StartTime(Looop))) ^ 2)) 'Calculating falling distance
                
                If YPos(Looop) <= StartHeight(Looop) + 0.1 Then MovementDone(Looop) = True      'The letter reached the max. height. The upmovement is complete
                
            End If
            
            YPos(Looop) = MoveDistance
            
            If YPos(Looop) > xPctBox.ScaleHeight - 1 Then                                   'If the letter fell xPctBox of our picturebox ;) we fix it
                YPos(Looop) = xPctBox.ScaleHeight - 1                                       'At the bottom position
            End If
            
            xPctBox.CurrentX = xPctBox.ScaleWidth / 2 - Int((Len(Message) / 2)) + Looop
            xPctBox.CurrentY = YPos(Looop)                                                  'Setting the letters y position
            xPctBox.ForeColor = TextColor(Looop)                                            'Setting the letters color
            xPctBox.Print Mid(Message, Looop, 1)                                            'Text xPctBoxput
        
        Next Looop
        
        DoEvents
    
        For Looop = 1 To Len(Message)
        
            If MovementDone(Looop) = True Then
                
                If DownMovement(Looop) = True Then     'Switch between up/downmovement
                    DownMovement(Looop) = False
                    StartHeight(Looop) = StartHeight(Looop) + ((xPctBox.ScaleHeight - StartHeight(Looop)) * PowerLoss(Looop))   'New Startheight, because of speed lost ?!?!
                    UpMovementTime(Looop) = Sqr((xPctBox.ScaleHeight - StartHeight(Looop)) / (0.5 * 9.81))        'How long will the NEXT upmovement last ???
                Else
                    DownMovement(Looop) = True
                End If
                
                StartTime(Looop) = Timer               'Set the StartTime of a new movement
                MovementDone(Looop) = False
            End If
            
         Next Looop
                
    Loop 'Until StartHeight = xPctBox.ScaleHeight
    
End Sub

