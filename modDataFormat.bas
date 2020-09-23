Attribute VB_Name = "modDataFormat"
Option Explicit

Type tResource
    sName As String
    lLength As Long
    lBuffer() As Byte
End Type

Type tResourceStart
    lResourceCount As Long
    vResources() As tResource
End Type

Type tVersion
    major As Byte
    minor As Byte
    special As Byte
    build As Byte
End Type

Type tHeader
    id(0 To 4) As Byte
    lNumOfElements As Long
    vVersion As tVersion
    sTitle As String
End Type

Type tClosingHeader
    id(0 To 4) As Byte
End Type

Type tProperty
    sName As String
    sValue As Variant
End Type

Type tElement
    sClassName As String
    lNumOfProperties As Long
'    properties() As tProperties
End Type

Function ReadBinaryFile(sFileName As String)
Dim header As tHeader
Dim closingheader As tClosingHeader
Dim elements() As tElement
Dim xProperty() As tProperty
Dim i As Long
Dim j As Long
PATH = StripPath(sFileName)

Open sFileName For Binary As #1
Get #1, , header
ReDim elements(header.lNumOfElements - 1)
If Not ValidateHeader(header) Then
    MsgBox "Corrupt file"
    Exit Function
End If

'MsgBox "XHTML v" & header.vVersion.major & "." & header.vVersion.minor & "." & header.vVersion.special & "." & header.vVersion.build
'MsgBox header.sTitle
Load frmHidden
frmHidden.Tag = sFileName
SetPageTitle header.sTitle

For i = 0 To header.lNumOfElements - 1
    Get #1, , elements(i)
    ReDim xProperty(elements(i).lNumOfProperties - 1)
    Get #1, , xProperty '(j)
    ProcessProperty elements(i).sClassName, xProperty
    Erase xProperty
Next i
Get #1, , closingheader

If Not ValidateClosingHeader(closingheader) Then
    MsgBox "end of file expected, but not found!" & vbCrLf & "file may be incomplete of damaged!"
    Exit Function
End If

Close #1

frmHidden.Show

End Function

Function WriteTestBinaryFile(sFileName As String)
Dim header As tHeader
Dim elements() As tElement
Dim xProperty() As tProperty
Dim i As Long
Dim j As Long
Open sFileName For Output As #1: Close #1
Open sFileName For Binary As #1
WriteHeader 4, "XHTML Demonstration"

    WriteElement "XHTML.Label", 9 '10
        WriteProperty "Width", "80%"
        WriteProperty "Height", "100%"
        
        WriteProperty "ForeColor", vbBlack
        WriteProperty "Caption", "This is an example of XHTML. It allows the execution of 'compiled html files'. This is a format I created. It is a binary file format and is a lot smaller and compact that HTML. This sample can be seen in the 'sample' folder along with this product. This form is resizeable, so try and resize it and watch as the controls are re-positioned. The format is designed to allow easy streaming over a HTTP connection, with will be implemented in the next version. It also features support for ANY type of controls (even custom ones), support for ANY event and much more. The web page is compiled into a small binary file. Future versions will include a full HTML to XHTML converter and editor." & vbCrLf & vbCrLf & "Please vote and support my efforts! If you have any comments, please email me."
        WriteProperty "HAlign", "Centered"
        WriteProperty "Top", "30%"
        WriteProperty "FontSize", 10
        WriteProperty "AutoSize", True
        WriteProperty "Visible", True

    WriteElement "XHTML.CommandButton", 8
        WriteProperty "Width", "50%"
        WriteProperty "Height", 30
        
        WriteProperty "HAlign", "centered"
        WriteProperty "Top", "91.5%"
        
        WriteProperty "Event Click", "Hello, Windows User!"
        WriteProperty "Caption", "Click for hello"
        WriteProperty "Enabled", True
        WriteProperty "Visible", False
        
    WriteElement "XHTML.Image", 9
        WriteProperty "Width", 230
        WriteProperty "Height", 92
        
        WriteProperty "BorderStyle", 1
        
        WriteProperty "HAlign", "centered"
        WriteProperty "Top", "5%"
        
        WriteProperty "AutoSize", False
        WriteProperty "Picture", "title.jpg"
        
        WriteProperty "Event Click", "Whay did you click the logo???"
        WriteProperty "Visible", True
    
    WriteElement "XHTML.Hyperlink", 10 '10
        WriteProperty "Width", "170"
        WriteProperty "Height", "50"
        WriteProperty "Event Click", "My Email address is jbay101@hotmail.com."
        WriteProperty "ForeColor", vbBlack
        WriteProperty "Caption", "Email the author"
        WriteProperty "FontSize", 15
        WriteProperty "AutoSize", True
        WriteProperty "HAlign", "Centered"
        WriteProperty "VAlign", "bottom"
        WriteProperty "Visible", True

WriteClosingHeader
Close #1
End Function

Private Function WriteElement(sClassName As String, lNumOfProperties As Long)
Dim tmp_element As tElement
tmp_element.lNumOfProperties = lNumOfProperties
tmp_element.sClassName = sClassName

Put #1, , tmp_element
End Function

Private Function WriteProperty(sName As String, Optional sValue As Variant = "")
Dim tmp_property As tProperty
tmp_property.sName = LCase(sName)
tmp_property.sValue = sValue

Put #1, , tmp_property
End Function

Private Function WriteHeader(lNumOfElements As Long, Optional sTitle As String = "", Optional sVersion As String = "1.0.0.0")
Dim header As tHeader
header.id(0) = Asc("X")
header.id(1) = Asc("H")
header.id(2) = Asc("T")
header.id(3) = Asc("M")
header.id(4) = Asc("L")
header.sTitle = sTitle
header.lNumOfElements = lNumOfElements

On Error Resume Next
Dim sv() As String
sv = Split(sVersion, ".")

header.vVersion.major = CByte((sv(0)))
header.vVersion.minor = CByte((sv(1)))
header.vVersion.special = CByte((sv(2)))
header.vVersion.build = CByte((sv(3)))

Put #1, , header
End Function

Private Function WriteClosingHeader()
Dim closingheader As tClosingHeader

closingheader.id(0) = 128
closingheader.id(1) = 64
closingheader.id(2) = 32
closingheader.id(3) = 16
closingheader.id(4) = 0

Put #1, , closingheader
End Function

Private Function ValidateClosingHeader(closingheader As tClosingHeader) As Boolean
ValidateClosingHeader = False
If closingheader.id(0) <> 128 Then Exit Function
If closingheader.id(1) <> 64 Then Exit Function
If closingheader.id(2) <> 32 Then Exit Function
If closingheader.id(3) <> 16 Then Exit Function
If closingheader.id(4) <> 0 Then Exit Function


ValidateClosingHeader = True
End Function
Private Function ValidateHeader(header As tHeader) As Boolean
ValidateHeader = False
If header.id(0) <> Asc("X") Then Exit Function
If header.id(1) <> Asc("H") Then Exit Function
If header.id(2) <> Asc("T") Then Exit Function
If header.id(3) <> Asc("M") Then Exit Function
If header.id(4) <> Asc("L") Then Exit Function


ValidateHeader = True
End Function
