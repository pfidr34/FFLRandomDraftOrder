Attribute VB_Name = "modRandom"
Option Compare Database
Option Explicit

'Function to randomly pick the draft order
Public Function f_Draft_Order() As Boolean

    'Add Error Handler
    On Error GoTo f_Draft_Order_Err
    
    'Dim variables
    Dim rst         As ADODB.Recordset
    Dim i           As Integer
    Dim intPlayers  As Integer
    Dim intDrafted  As Integer
    Dim intCurPlyr  As Integer
    Dim intDraftOrd As Integer
    Dim strCurPlyr  As String
    Dim blnDraftOrd As Boolean
    
    'Create draft results file
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim Fileout As Object
    Set Fileout = fso.CreateTextFile(CurrentProject.Path & "\DraftResults.txt", True, False)
    Fileout.Close
    
    'Find out how many players we have
    f_Set_Stat ("Counting players...")
    intPlayers = f_ADO_Lookup("COUNT(1)", "tblPlayers", "1=1", CurrentProject.Connection)
    
    'Get player names in random order
    Set rst = f_ADO_Get_Recordset("SELECT PLAYER FROM TBLPLAYERS ORDER BY RND(ID)", CurrentProject.Connection)
    f_Debug ("Collected " & intPlayers & " players in random order")
    f_Debug ("")
    
    'Set counter
    i = 1
    
    'Loop through and assign player a player number
    Do Until rst.EOF
    
        'Set player number
        f_Set_Stat ("Setting player position number for " & rst.Fields("PLAYER"))
        strSQL = "UPDATE TBLPLAYERS SET PLAYERNUMBER = " & i & " WHERE PLAYER = '" & rst.Fields("PLAYER") & "'"
        f_ADO_Command strSQL, CurrentProject.Connection
        f_Debug (rst.Fields("PLAYER") & " was assigned player # " & i)
        f_Debug ("")
        
        'Increment counter
        i = i + 1
        
        'Next Player
        rst.MoveNext
    Loop
    
    'Loop through until all 12 players are drafter
    f_Set_Stat ("Setting draft postition for players")
    f_Debug ("")
    Do Until intDrafted = intPlayers
    
        'Select a random player number
        Randomize
        intCurPlyr = Int((intPlayers - 1 + 1) * Rnd + 1)
        f_Debug ("Randomly selected player # " & intCurPlyr & " - " & Nz(f_ADO_Lookup("PLAYER", "tblPlayers", "PLAYERNUMBER = " & intCurPlyr, CurrentProject.Connection), ""))
        
        'Check if player has a draft number
        intDraftOrd = Nz(f_ADO_Lookup("DRAFTNUMBER", "tblPlayers", "PLAYERNUMBER = " & intCurPlyr, CurrentProject.Connection), -1)
        If intDraftOrd = -1 Then
            f_Debug ("Player has no draft postition. Selecting a draft postition...")
        Else
            f_Debug ("Player already has draft postition #" & intDraftOrd)
            f_Debug ("")
        End If
            
        'Player doesn't have a draft number
        If intDraftOrd = -1 Then
        
            'Get player name
            strCurPlyr = Nz(f_ADO_Lookup("PLAYER", "tblPlayers", "PLAYERNUMBER = " & intCurPlyr, CurrentProject.Connection), "")
            
            'Reset draft order boolean
            blnDraftOrd = False
        
            'Run loop until player gets an open draft postion
            Do Until blnDraftOrd
            
                'Select a random draft postion
                Randomize
                intDraftOrd = Int((intPlayers - 1 + 1) * Rnd + 1)
                f_Debug ("Randomly selected draft positon # " & intDraftOrd & " - Checking to see if position is taken")
                
                'See if draft postion is availabe, set boolean accordingly
                blnDraftOrd = f_ADO_Lookup("COUNT(1)", "tblPlayers", "DRAFTNUMBER = " & intDraftOrd, CurrentProject.Connection)
                
                f_Debug ("Is draft postiton taken? " & blnDraftOrd)
        
                'Draft position was available, assign it to the player
                If Not blnDraftOrd Then
                    strSQL = "UPDATE TBLPLAYERS SET DRAFTNUMBER = " & intDraftOrd & " WHERE PLAYERNUMBER = " & intCurPlyr
                    f_ADO_Command strSQL, CurrentProject.Connection
                    f_Set_Stat ("Draft postition set for player " & strCurPlyr & " is # " & intDraftOrd)
                    blnDraftOrd = True
                End If
                
                f_Debug ("")
        
            Loop 'to find an open draft postion
        
        End If
        
        'Set number of players drafted
        intDrafted = f_ADO_Lookup("COUNT(1)", "tblPlayers", "DRAFTNUMBER IS NOT NULL", CurrentProject.Connection)
        
    Loop 'to fill all draft postions
    
    f_Debug ("Draft order complete")
    
    'Exit and return
    f_Draft_Order = True
    Exit Function
      
f_Draft_Order_Err:
    f_Draft_Order = False

End Function

'Function to easily set the status on the main form, and save it as a debug item
Public Function f_Set_Stat(ByVal strStat As String) As Boolean

    'Add error handler
    On Error GoTo f_Set_Stat_Err

    'Set caption
    Form_frmDraft.lblStat.Caption = strStat
    
    'Record as debug action
    f_Debug strStat

    'Sleep for 2 seconds to let everyone read
    p_Sleep (1)

    DoEvents

    'Return and Exit
    f_Set_Stat = True
    Exit Function
    
    'Error handler
f_Set_Stat_Err:
    f_Set_Stat = False

End Function
    
'Public function to spit out debug results
Public Function f_Debug(ByVal strDebug As String) As Boolean
    
    Debug.Print Time & " - " & strDebug
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim Fileout As Object
    Set Fileout = fso.OpenTextFile(CurrentProject.Path & "\DraftResults.txt", ForAppending, TristateFalse)
    Fileout.WriteLine Time & " - " & strDebug
    Fileout.Close
    
End Function

