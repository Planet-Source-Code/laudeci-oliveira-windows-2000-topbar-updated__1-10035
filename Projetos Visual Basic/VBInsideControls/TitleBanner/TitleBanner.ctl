VERSION 5.00
Begin VB.UserControl TitleBanner 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   1110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7290
   ScaleHeight     =   1110
   ScaleWidth      =   7290
   Begin VB.Image pctImage 
      Height          =   825
      Left            =   5655
      Top             =   135
      Width           =   1275
   End
End
Attribute VB_Name = "TitleBanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const LeftDistance As Long = 60
'Event Declarations:

'^''^
' any comments to laudecioliveira@hotmail.com
Private mCaption As String
Private Const m_def_caption = "TopBar Caption"

Private mCaptionColor As OLE_COLOR
Private Const m_def_caption_color = vbBlack

Private mCaptionFont As New StdFont

Private mDescription As String
Private Const m_def_Description = "TopBar Description"

Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."



Private Sub Draw3dLine()
    Dim r As RECT   ' Used by DrawEdge to determine where to draw.
    Dim oldColor As OLE_COLOR
    Dim oldFont As Object
    Dim mLeft As Long, mTop As Long
    '-----------------------------------------------------------------
    ' Location of the etched box.
    '-----------------------------------------------------------------
    With r
        .Left = ScaleX(UserControl.ScaleLeft + 10, vbTwips, vbPixels)
        .Top = ScaleX(UserControl.ScaleHeight - 30, vbTwips, vbPixels)
        .Right = ScaleX(UserControl.ScaleWidth - 10, vbTwips, vbPixels)
        .Bottom = ScaleX(UserControl.ScaleHeight, vbTwips, vbPixels)
    End With
    '-----------------------------------------------------------------
    ' Draw it.
    '-----------------------------------------------------------------
    UserControl.Cls
    DrawEdge UserControl.hDC, r, EDGE_ETCHED, BF_RECT
    
    oldColor = UserControl.ForeColor
    Set oldFont = UserControl.Font

    
    'set the especified font for the title caption
    UserControl.ForeColor = mCaptionColor
    Set UserControl.Font = mCaptionFont
    
    ' Draw the big caption
    SetRect r, 20, 20, UserControl.ScaleWidth - pctImage.Width, UserControl.TextHeight("X")
    DrawTextEx UserControl.hDC, mCaption, Len(mCaption), r, DT_WORDBREAK, ByVal 0&
    
    'restore the old font and color
    UserControl.ForeColor = oldColor
    Set UserControl.Font = oldFont
    
    ' draw the description
    mLeft = r.Left + 20
    mTop = r.Top + ScaleY(UserControl.TextHeight("X"), vbTwips, vbPixels) + mCaptionFont.Size
    SetRect r, mLeft, mTop, UserControl.ScaleWidth - pctImage.Width, UserControl.TextHeight("X")
    DrawTextEx UserControl.hDC, mDescription, Len(mDescription), r, DT_WORDBREAK, ByVal 0&

End Sub


Private Sub UserControl_InitProperties()
    Set mCaptionFont = Ambient.Font
    mCaptionFont.Bold = True
    mCaptionFont.Size = 10
    
    UserControl.FontName = "Arial"
    UserControl.FontBold = False
    UserControl.FontSize = 8
    mCaption = m_def_caption
    mDescription = m_def_Description
    
End Sub

Private Sub UserControl_Resize()
    Draw3dLine
    CenterImage
End Sub

Private Sub CenterImage()
    Dim mTop As Single
    Dim mLeft As Single
    ' put the image in the correct place
    mTop = (UserControl.Height - pctImage.Height) / 2
    mLeft = (UserControl.Width - pctImage.Width) - LeftDistance
    pctImage.Move mLeft, mTop
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = mCaptionColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    mCaptionColor = New_ForeColor
    PropertyChanged "ForeColor"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = mCaptionFont
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set mCaptionFont = New_Font
    PropertyChanged "Font"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=pctImage,pctImage,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = pctImage.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set pctImage.Picture = New_Picture
    PropertyChanged "Picture"
    If UserControl.Height <= pctImage.Height Then
        UserControl.Height = (pctImage.Height + (LeftDistance * 2))
    Else
        UserControl_Resize
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,Caption
Public Property Get CaptionTitle() As String
Attribute CaptionTitle.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    CaptionTitle = mCaption
End Property

Public Property Let CaptionTitle(ByVal New_CaptionTitle As String)
    mCaption = New_CaptionTitle
    PropertyChanged "CaptionTitle"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label2,Label2,-1,Caption
Public Property Get CaptionDescription() As String
Attribute CaptionDescription.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    CaptionDescription = mDescription
End Property

Public Property Let CaptionDescription(ByVal New_CaptionDescription As String)
    mDescription = New_CaptionDescription
    PropertyChanged "CaptionDescription"
    UserControl_Resize
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    mCaptionColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set mCaptionFont = PropBag.ReadProperty("Font", Ambient.Font)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    mCaption = PropBag.ReadProperty("CaptionTitle", "Put your title here:")
    mDescription = PropBag.ReadProperty("CaptionDescription", "Put your description here:")
    Debug.Print "ReadProperty"
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, vbWhite)
    Call PropBag.WriteProperty("ForeColor", mCaptionColor, vbBlack)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", mCaptionFont, Ambient.Font)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("CaptionTitle", mCaption, "Put your title here:")
    Call PropBag.WriteProperty("CaptionDescription", mDescription, "Put your description here:")
End Sub

