VERSION 5.00
Begin VB.Form frmVBRecent 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Visual Basic 6.0 Recent File Manager"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstRecentFiles 
      Height          =   7665
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   8115
   End
   Begin VB.Menu mnuManage 
      Caption         =   "Manage Files"
      Begin VB.Menu mnuUp 
         Caption         =   "Move selected item up"
      End
      Begin VB.Menu mnuDown 
         Caption         =   "Move selected item down"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "Remove selected item"
      End
      Begin VB.Menu SEPERATOR1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save Changes"
         Shortcut        =   ^S
      End
      Begin VB.Menu SEPERATOR2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmVBRecent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

'---NOTE FROM THE DEVELOPER:
'---In order for the registry to be updated, you have to close all instances of the Visual
'---Basic IDE, including the current one.  Compile this project into an executable then
'---run the application.

'---Short-cut keys:
'---Shift-Up:  Move the selected project up in the list order.
'---Shift-Dn:  Move the selected project down in the list order.
'---Delete:    Deletes the selected project.
'---Ctrl-S:    Saves the current settings.

'#########################################################################################
'#########################################################################################
'#########################################################################################
'########                                                                        #########
'########                                 VARIABLES                              #########
'########                                                                        #########
'#########################################################################################
'#########################################################################################
'#########################################################################################

Dim lngHandleKey        As Long
Dim lngReturnVal        As Long
Dim strReturnValueName  As String
Dim intCounter          As Integer
Dim lngBuffer           As Long
Dim strReturnValue      As String
Dim blnChangesMade      As Boolean
Dim intListIndex        As Integer

'#########################################################################################
'#########################################################################################
'#########################################################################################
'########                                                                        #########
'########                                 EVENTS                                 #########
'########                                                                        #########
'#########################################################################################
'#########################################################################################
'#########################################################################################

Private Sub Form_Load()
   PopulateList
   If lstRecentFiles.ListCount = 0 Then Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If blnChangesMade Then
      If MsgBox("Would you like to save the changes made to the recent project " & _
                 "file list?", vbYesNo + vbQuestion, "Save Changes") = vbYes Then SaveChanges
   End If
End Sub

Private Sub lstRecentFiles_Click()
   mnuUp.Enabled = (lstRecentFiles.ListIndex > 0)
   mnuDown.Enabled = (lstRecentFiles.ListIndex < lstRecentFiles.ListCount - 1)
   
   If intListIndex > 0 Then lstRecentFiles.Selected(intListIndex) = True
   
End Sub

Private Sub lstRecentFiles_KeyDown(KeyCode As Integer, Shift As Integer)
   
   If KeyCode = vbKeyDelete Then RemoveSelectedItem
   
   If (KeyCode = vbKeyUp) Then
      If Shift > 0 Then
         RelocateItem (MOVE_ITEM_UP)
      Else
         intListIndex = 0
      End If
   ElseIf (KeyCode = vbKeyDown) Then
      If (Shift > 0) Then
         RelocateItem (MOVE_ITEM_DOWN)
      Else
         intListIndex = 0
      End If
   End If
   
End Sub

Private Sub lstRecentFiles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   intListIndex = 0
End Sub

Private Sub lstRecentFiles_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then
      Me.PopupMenu mnuManage
   End If
End Sub

Private Sub mnuUp_Click()
   RelocateItem (MOVE_ITEM_UP)
End Sub

Private Sub mnuDown_Click()
   RelocateItem (MOVE_ITEM_DOWN)
End Sub

Private Sub mnuRemove_Click()
   RemoveSelectedItem
End Sub

Private Sub mnuSave_Click()
   If Not SaveChanges Then MsgBox "An error occured while attempting to update the registry.", vbOKOnly + vbCritical, "Error"
End Sub

Private Sub mnuExit_Click()
   Unload Me
End Sub

'#########################################################################################
'#########################################################################################
'#########################################################################################
'########                                                                        #########
'########                                FUNCTIONS                               #########
'########                                                                        #########
'#########################################################################################
'#########################################################################################
'#########################################################################################

Private Sub PopulateList()
   
      '---Attempt to open the key.
   If Not OpenKey(KEY_QUERY_VALUE) Then Exit Sub
   
      '---Clear the listbox.
   lstRecentFiles.Clear
   
      '---Populate the listbox with the entries found from the registry for Visual Basic's
      '---recent project section.  The most projects that can be stored is 50.
   For intCounter = 1 To 50
   
      strReturnValue = Space(255)
      lngBuffer = 255
      
      strReturnValueName = Trim(Str(intCounter))
      lngReturnVal = RegQueryValueEx(lngHandleKey, strReturnValueName, 0&, 0, strReturnValue, lngBuffer)
      
         '---If we get an error, either the key doesn't exist, or something else happened,
         '---in either case, we're done populating.
      If lngReturnVal > 0 Then Exit For
         
      strReturnValue = Left(strReturnValue, lngBuffer - 1)
      
      lstRecentFiles.AddItem strReturnValue
   
   Next
      
      '--- Close the registry key.
   RegCloseKey lngHandleKey
   
   blnChangesMade = False
End Sub

Private Sub RemoveSelectedItem()

      '---Unremark this code if you like a nag screen before each delete attempt.
   'If MsgBox("Are you certain you want to remove " & lstRecentFiles.Text & " from your Visual " & _
             "Basic recent projects list?", vbYesNo + vbQuestion, "Confirm remove") = vbNo Then Exit Sub

      '---Make sure there is an item selected in the listbox.
   If lstRecentFiles.ListIndex < 0 Then Exit Sub
   
      '---Remove the selected item.
   lstRecentFiles.RemoveItem lstRecentFiles.ListIndex
   
   blnChangesMade = True
End Sub

Private Sub RelocateItem(intIndex As Integer)
   Dim strTemp             As String
   
         '---Make sure there is an item selected in the listbox.
   If lstRecentFiles.ListIndex < 0 Then Exit Sub
   
      '---Store the selected item's index to highlight it again when the sorting is finished.
   intListIndex = lstRecentFiles.ListIndex
   
      '---If intIndex = 0 then move the item up, otherwise move it down.
   If (intIndex = 0) And (intListIndex > 0) Then
   
         '---Add the selected item one index above where it's at now, then
         '---remove the selected item.
      lstRecentFiles.AddItem lstRecentFiles.Text, intListIndex - 1
      lstRecentFiles.RemoveItem intListIndex + 1
      intListIndex = intListIndex - 1
      
   ElseIf (intListIndex < lstRecentFiles.ListCount - 1) Then
   
         '---Add the selected item one index below where it's at now, then
         '---remove the selected item.
      lstRecentFiles.AddItem lstRecentFiles.Text, intListIndex + 2
      lstRecentFiles.RemoveItem intListIndex
      intListIndex = intListIndex + 1
      
   End If
   
      '---Select the newly added (relocated) item.
   lstRecentFiles.Selected(intListIndex) = True
   
   blnChangesMade = True
End Sub

Private Function SaveChanges() As Boolean
   Dim aryKeyValues(50)    As String
   Dim intItemCount        As Integer
   Dim intLooper           As Integer

   intItemCount = lstRecentFiles.ListCount - 1
   Erase aryKeyValues
   SaveChanges = False
   
   intListIndex = lstRecentFiles.ListIndex
   
      '--- Populate the array with the files in the listbox.
   For intCounter = 0 To intItemCount
         aryKeyValues(intCounter + 1) = lstRecentFiles.List(intCounter)
   Next
   
   If Not OpenKey(KEY_ALL_ACCESS) Then Exit Function
   
      '--- Save filenames to the registry.
   For intCounter = 1 To intItemCount + 1
      
      strReturnValueName = Trim(Str(intCounter))
      
      strReturnValue = aryKeyValues(intCounter) & Chr(0)
      lngBuffer = Len(strReturnValue)
      
      lngReturnVal = RegSetValueEx(lngHandleKey, strReturnValueName, 0, REG_SZ, strReturnValue, lngBuffer)
      
      If lngReturnVal > 0 Then
         RegCloseKey lngHandleKey
         Exit Function
      End If
      
   Next intCounter
   
      '--- Delete the remaining values from the registry.
   For intCounter = (intItemCount + 2) To 50
      
      strReturnValueName = Trim(Str(intCounter))
      
      lngReturnVal = RegDeleteValue(lngHandleKey, strReturnValueName)
      
      If lngReturnVal > 0 Then
         RegCloseKey lngHandleKey
         Exit For
      End If
      
   Next intCounter
   
      '--- Close the registry key.
   RegCloseKey lngHandleKey
   
   PopulateList
   
   SaveChanges = True
   
   If intListIndex > 0 Then lstRecentFiles.Selected(intListIndex) = True
   
End Function

Private Function OpenKey(vlngKeyType As Long) As Boolean
   
      '---Open the Current User key.
   lngReturnVal = RegOpenKeyEx(HKEY_CURRENT_USER, HKEY_CURRENT_USER_KEY, 0, vlngKeyType, lngHandleKey)
   
   If lngReturnVal > 0 Then MsgBox "An error occured while attempting to open the Recent Files key", vbOKOnly + vbCritical

   OpenKey = (lngReturnVal = 0)

End Function















