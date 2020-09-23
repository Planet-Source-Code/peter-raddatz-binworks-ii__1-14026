VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BinDLL Demo - By P.Raddatz"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   5835
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option13 
      Caption         =   "Toggle Bit"
      Height          =   255
      Left            =   3960
      TabIndex        =   21
      Top             =   2040
      Width           =   1215
   End
   Begin VB.OptionButton Option12 
      Caption         =   "Right Shift"
      Height          =   255
      Left            =   3960
      TabIndex        =   20
      Top             =   1800
      Width           =   1335
   End
   Begin VB.OptionButton Option11 
      Caption         =   "Left Shift"
      Height          =   255
      Left            =   3960
      TabIndex        =   19
      Top             =   1560
      Width           =   1335
   End
   Begin VB.OptionButton Option10 
      Caption         =   "Right Rotate"
      Height          =   255
      Left            =   3960
      TabIndex        =   18
      Top             =   1320
      Width           =   1335
   End
   Begin VB.OptionButton Option9 
      Caption         =   "Left Rotate"
      Height          =   255
      Left            =   3960
      TabIndex        =   17
      Top             =   1080
      Width           =   1335
   End
   Begin VB.OptionButton Option8 
      Caption         =   "HexAr to Array"
      Height          =   195
      Left            =   2280
      TabIndex        =   15
      Top             =   1800
      Width           =   1335
   End
   Begin VB.OptionButton Option7 
      Caption         =   "Array to HexAr"
      Height          =   195
      Left            =   2280
      TabIndex        =   14
      Top             =   1560
      Width           =   1335
   End
   Begin VB.OptionButton Option6 
      Caption         =   "Hex to Var"
      Height          =   195
      Left            =   2280
      TabIndex        =   13
      Top             =   1320
      Width           =   1335
   End
   Begin VB.OptionButton Option5 
      Caption         =   "Var to Hex"
      Height          =   195
      Left            =   2280
      TabIndex        =   12
      Top             =   1080
      Width           =   1335
   End
   Begin VB.OptionButton Option4 
      Caption         =   "BinAr to Array"
      Height          =   195
      Left            =   600
      TabIndex        =   11
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Frame Frame4 
      Caption         =   " Misc. "
      Height          =   1575
      Left            =   3840
      TabIndex        =   10
      Top             =   840
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      Caption         =   " Hexadecimal "
      Height          =   1215
      Left            =   2160
      TabIndex        =   9
      Top             =   840
      Width           =   1575
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Array to BinAr"
      Height          =   195
      Left            =   600
      TabIndex        =   7
      Top             =   1560
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Bin to Var"
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   1320
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Var to Bin"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   1080
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   4440
      TabIndex        =   4
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   3240
      Width           =   1935
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   " Binary "
      Height          =   1215
      Left            =   480
      TabIndex        =   8
      Top             =   840
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   1935
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   5175
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "That's 256,000,000 conversions!"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1560
         Visible         =   0   'False
         Width           =   3255
      End
   End
   Begin TabDlg.SSTab Tab1 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   4471
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Bytes"
      TabPicture(0)   =   "Dll-Demo.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Integers"
      TabPicture(1)   =   "Dll-Demo.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Longs"
      TabPicture(2)   =   "Dll-Demo.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
   End
   Begin VB.Label Label2 
      Caption         =   "Seconds:"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pick the desired VARTYPE and FUNCTION to see 1,000,000 operations."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2760
      Width           =   5535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mostly this code arose from my need for a Bin$, rotL, rotR,
'shifts etc.(see, I'm a crypto nut!)

'There are some snippets posted on the net that show how to
'implement, at least, some of what I was looking for, but I
'still was not satisfied. A little while ago I wrote a couple
'of VB snippets that would solve millions of Bins / sec. and
'at the same time reworked the Hex$ to do the same, but even
'that was not fast enough. So I decided to try a DLL in C.
'Knowing absolutely nothing about C, the task was daunting,
'but I'm always game to take on a challenge. Let me at'em!

'The attached bindll.dll has 13 functions, each, for BYTEs,
'INTEGERs and LONGs to take care of the following...
'Variable to Bin        (byte2bin, int2bin, lng2bin)
'Bin to Variable        (bin2byte, bin2int, bin2lng)
'VarArray to BinArray   (byteA2binA, intA2binA, lngA2binA)
'BinArray to VarArray   (binA2byteA, binA2intA, binA2lngA)
'Variable to Hex        (byte2hex, int2hex, lng2hex)
'Hex to Variable        (hex2byte, hex2int, hex2lng)
'VarArray to HexArray   (byteA2hexA, intA2hexA, lngA2hexA)
'HexArray to VarArray   (hexA2byteA, hexA2intA, hexA2lngA)
'VarRotateLeft          (BrotL, IrotL, LrotL)
'VarRotateRight         (BrotR, IrotR, LrotR)
'VarShiftLeft           (BshL, IshL, LshL)
'VarShiftRight          (BshR, IshR, LshR)
'and...
'VarBittoggle           (Btoggle, Itoggle, Ltoggle)

'My aim with this DLL has been 1)SPEED, 2)SPEED and 3)SPEED.
'In that order.
'Because of that, there is absolutely NO ERROR CHECKING in
'the C code. It is YOUR responsibility to make sure that you
'dimension the arrays properly and that you do send a LONG
'when a BYTE is expected, etc.
'This Demo, not only shows off the speed of the DLL, but you
'can put in a FOR...NEXT loop after each function to check
'all the answers. (that's what I did and everything works as
'advertised!)

'Why write a Hex function when VB has HEX$? The answer is...
'1)SPEED, 2)SP... you get the picture.
'On my 300 h$=HEX$(243) & BVar=Val("&h"+h$) chuck along at
'about 323,000 conv./sec. as where this code does over 40,000,000
'/sec. (array version) The draw back? Unless you need a STRING
'like 7AF4 for 31476 instead of 7 10 15 4 that you get here you are,
'almost, 124 times better off using this code.

'As mentioned, this code DOES NOT produce BIN$ or HEX$ it pro-
'duces ARRAYS of BYTEs containing 1s & 0s or 15s and 6 etc. This
'method enables you to manipulate the individual digits much
'more easily. (i%=Bin(4) instead of i%=Val(mid$(B$,4,1)) ). In
'any event, if you have use for these Functions I'm sure you'll
'make the switch quite easily.

'Any questions? e-mail me! rabbit@bluecrow.com
'Peter Raddatz January 1, 2001

Option Explicit
Dim tabnum As Byte
Dim optnum As Byte

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdStart_Click()
    StartDemo
End Sub

Private Sub Form_Load()
Dim b As Byte
    
    b = 243
    byte2bin b, BinA(0) 'load the DLL, nothing else.
    
    optnum = 1
    tabnum = 0
End Sub

Private Sub Option1_Click()
    optnum = 1
    Label3.Visible = False
End Sub

Private Sub Option2_Click()
    optnum = 2
    Label3.Visible = False
End Sub

Private Sub Option3_Click()
    optnum = 3
    Label3.Visible = True
End Sub
Private Sub Option4_Click()
    optnum = 4
    Label3.Visible = True
End Sub
Private Sub Option5_Click()
    optnum = 5
    Label3.Visible = False
End Sub
Private Sub Option6_Click()
    optnum = 6
    Label3.Visible = False
End Sub
Private Sub Option7_Click()
    optnum = 7
    Label3.Visible = True
End Sub
Private Sub Option8_Click()
    optnum = 8
    Label3.Visible = True
End Sub
Private Sub Option9_Click()
    optnum = 9
    Label3.Visible = False
End Sub
Private Sub Option10_Click()
    optnum = 10
    Label3.Visible = False
End Sub
Private Sub Option11_Click()
    optnum = 11
    Label3.Visible = False
    Label3.Visible = False
End Sub
Private Sub Option12_Click()
    optnum = 12
    Label3.Visible = False
End Sub
Private Sub Option13_Click()
    optnum = 13
    Label3.Visible = False
End Sub

Private Sub Tab1_Click(PreviousTab As Integer)
    tabnum = Tab1.Tab
End Sub

Sub StartDemo()
Select Case tabnum
    Case 0      'Bytes
        Select Case optnum
            Case 1
                b = 243
                t = GetTickCount
                For x = 1 To 1000000
                    byte2bin b, BinA(0)
                Next
            Case 2
                t = GetTickCount
                For x = 1 To 1000000
                    b = bin2byte(BinA(0))
                Next
            Case 3
                For x = 0 To 255    'init ByteArray
                    ByteA(x) = x
                Next
                t = GetTickCount
                For x = 1 To 1000000
                    byteA2binA ByteA(0), BBinA(0), 256
                Next
            Case 4
                t = GetTickCount
                For x = 1 To 1000000
                    binA2byteA BBinA(0), ByteA(0), 256
                Next
            Case 5
                b = 254
                t = GetTickCount
                For x = 1 To 1000000
                    byte2hex b, BHex(0)
                Next
            Case 6
                t = GetTickCount
                For x = 1 To 1000000
                    b = hex2byte(BHex(0))
                Next
            Case 7
                For x = 0 To 255
                    ByteA(x) = x
                Next
                t = GetTickCount
                For x = 1 To 1000000
                    byteA2hexA ByteA(0), BHexA(0), 256
                Next
            Case 8
                t = GetTickCount
                For x = 1 To 1000000
                    hexA2byteA BHexA(0), ByteA(0), 256
                Next
            Case 9
                b = 168
                t = GetTickCount
                For x = 1 To 1000000
                    B1 = BrotL(b, 3)
                Next
            Case 10
                b = 69
                t = GetTickCount
                For x = 1 To 1000000
                    B1 = BrotR(b, 3)
                Next
            Case 11
                b = 63
                t = GetTickCount
                For x = 1 To 1000000
                    B1 = BshL(b, 2)
                Next
            Case 12
                b = 252
                t = GetTickCount
                For x = 1 To 1000000
                    B1 = BshR(b, 2)
                Next
            Case 13
                b = 255
                t = GetTickCount
                For x = 1 To 1000000
                    B1 = Btoggle(b, 3)
                Next
        End Select
    Case 1      'Integers
        Select Case optnum
            Case 1
                i = 24347
                t = GetTickCount
                For x = 1 To 1000000
                    int2bin i, IntA(0)
                Next
            Case 2
                t = GetTickCount
                For x = 1 To 1000000
                    i = bin2int(IntA(0))
                Next
            Case 3
                For x = 0 To 255    'init IntArray
                    IIntA(x) = x
                Next
                t = GetTickCount
                For x = 1 To 1000000
                    intA2binA IIntA(0), IBinA(0), 256
                Next
            Case 4
                t = GetTickCount
                For x = 1 To 1000000
                    binA2intA IBinA(0), IIntA(0), 256
                Next
            Case 5
                i = 25469
                t = GetTickCount
                For x = 1 To 1000000
                    int2hex i, IHex(0)
                Next
            Case 6
                t = GetTickCount
                For x = 1 To 1000000
                    i = hex2int(IHex(0))
                Next
            Case 7
                For x = 0 To 255
                    IIntA(x) = x
                Next
                t = GetTickCount
                For x = 1 To 1000000
                    intA2hexA IIntA(0), IHexA(0), 256
                Next
            Case 8
                t = GetTickCount
                For x = 1 To 1000000
                    hexA2intA IHexA(0), IIntA(0), 256
                Next
            Case 9
                i = 1684
                t = GetTickCount
                For x = 1 To 1000000
                    I1 = IrotL(i, 3)
                Next
            Case 10
                i = 13472
                t = GetTickCount
                For x = 1 To 1000000
                    I1 = IrotR(i, 3)
                Next
            Case 11
                i = 6303
                t = GetTickCount
                For x = 1 To 1000000
                    I1 = IshL(i, 2)
                Next
            Case 12
                i = 25212
                t = GetTickCount
                For x = 1 To 1000000
                    I1 = IshR(i, 2)
                Next
            Case 13
                i = 25578
                t = GetTickCount
                For x = 1 To 1000000
                    I1 = Itoggle(i, 12)
                Next
        End Select
    Case 2      'Longs
        Select Case optnum
            Case 1
                L = 243478591
                t = GetTickCount
                For x = 1 To 1000000
                    lng2bin L, LngA(0)
                Next
            Case 2
                t = GetTickCount
                For x = 1 To 1000000
                    L = bin2lng(LngA(0))
                Next
            Case 3
                For x = 0 To 255    'init LngArray
                    LLngA(x) = x
                Next
                t = GetTickCount
                For x = 1 To 1000000
                    lngA2binA LLngA(0), LBinA(0), 256
                Next
           Case 4
                t = GetTickCount
                For x = 1 To 1000000
                    binA2lngA LBinA(0), LLngA(0), 256
                Next
            Case 5
                L = 254694641
                t = GetTickCount
                For x = 1 To 1000000
                    lng2hex L, LHex(0)
                Next
            Case 6
                t = GetTickCount
                For x = 1 To 1000000
                    L = hex2lng(LHex(0))
                Next
            Case 7
                For x = 0 To 255
                    LLngA(x) = x
                Next
                t = GetTickCount
                For x = 1 To 1000000
                    lngA2hexA LLngA(0), LHexA(0), 256
                Next
            Case 8
                t = GetTickCount
                For x = 1 To 1000000
                    hexA2lngA LHexA(0), LLngA(0), 256
                Next
            Case 9
                L = 1684463
                t = GetTickCount
                For x = 1 To 1000000
                    L1 = LrotL(L, 3)
                Next
            Case 10
                L = 13475704
                t = GetTickCount
                For x = 1 To 1000000
                    L1 = LrotR(L, 3)
                Next
            Case 11
                L = 507633
                t = GetTickCount
                For x = 1 To 1000000
                    L1 = LshL(L, 2)
                Next
            Case 12
                L = 2030532
                t = GetTickCount
                For x = 1 To 1000000
                    L1 = LshR(L, 2)
                Next
            Case 13
                L = 2583578
                t = GetTickCount
                For x = 1 To 1000000
                    L1 = Ltoggle(L, 12)
                Next
        End Select
End Select
Text1 = CStr((GetTickCount - t) / 1000)
End Sub
        
            
