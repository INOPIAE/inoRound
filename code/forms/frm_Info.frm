VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Info 
   Caption         =   "inoRound"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "frm_Info.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Me.lblTitle.Caption = "inoRound Version " & strMakroVersion & " (" & dtVersionOf & ")"
    Me.lblCopyright.Caption = "Copyright 2020 - " & Year(Date)
    Me.lblInfo.Caption = strfrmInfo(0) & "https://github.com/INOPIAE/inoRound"
End Sub
