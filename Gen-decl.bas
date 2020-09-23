Attribute VB_Name = "Module2"
Option Explicit

'common DIMS for DLL-demo
Public x As Long
Public t As Long

'---Bytes
Public B As Byte, B1 As Byte, BinA(0 To 7) As Byte
Public ByteA(0 To 255) As Byte, BBinA(0 To 2047) As Byte
Public BHex(0 To 1) As Byte, BHexA(0 To 511) As Byte
'-
'---Ints
Public I As Integer, I1 As Integer, IntA(0 To 15) As Byte
Public IIntA(0 To 255) As Integer, IBinA(0 To 4095) As Byte
Public IHex(0 To 3) As Byte, IHexA(0 To 1023) As Byte
'-
'---Longs
Public L As Long, L1 As Long, LngA(0 To 31) As Byte
Public LLngA(0 To 255) As Long, LBinA(0 To 8191) As Byte
Public LHex(0 To 7) As Byte, LHexA(0 To 2047) As Byte
'-
