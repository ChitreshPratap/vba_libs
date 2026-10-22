VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IHeaderStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' Interface : IHeaderStyle
' Purpose   : The contract between CCollapsibleFrame and any header style.
'             Contains no code - style classes use "Implements IHeaderStyle".
'             CCollapsibleFrame.HeaderStyle is declared As IHeaderStyle, so
'             any implementing class can be assigned to it.
'
' To add a new style: create a class, write "Implements IHeaderStyle", and
' implement the two members below. Name every control you create
' NamePrefix & "<something>" so the container never treats it as content
' and can clean up leftovers.
'==============================================================================

' Called by the container whenever the header geometry, heading text or
' font changes (and once after the style is assigned).
'
'   Host        Object that owns the header controls (Frame, form or Page);
'               add decoration with Host.Controls.Add.
'   NamePrefix  Prefix for every control the style creates.
'   Title       The heading label. A style may move/resize it horizontally
'               inside [TextLeft, TextRight]; it must not change its caption.
'   BarLeft/Top/Width/Height  Header bar rectangle in Host coordinates.
'   TextLeft/TextRight        Horizontal area reserved for the heading
'                             (already excludes the icon, wherever it is).
Public Sub Render(ByVal Host As Object, _
                  ByVal NamePrefix As String, _
                  ByVal Title As MSForms.Label, _
                  ByVal BarLeft As Single, _
                  ByVal BarTop As Single, _
                  ByVal BarWidth As Single, _
                  ByVal BarHeight As Single, _
                  ByVal TextLeft As Single, _
                  ByVal TextRight As Single)
End Sub

' Removes every control this style created and drops its references.
' The style keeps its own settings, so it can be rendered again later.
Public Sub Clear()
End Sub
