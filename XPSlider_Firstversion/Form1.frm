VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Slider"
   ClientHeight    =   2685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   ScaleHeight     =   2685
   ScaleWidth      =   4485
   StartUpPosition =   3  'Windows Default
   Begin Project1.XPSlider XPSlider5 
      Height          =   300
      Left            =   390
      TabIndex        =   4
      Top             =   1965
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   529
      Value           =   0
      SliderWid_Height=   400
      BaseColor       =   12583104
      MidColor        =   16777214
      ValueCol        =   8388736
   End
   Begin Project1.XPSlider XPSlider4 
      Height          =   300
      Left            =   390
      TabIndex        =   3
      Top             =   1590
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   529
      Min             =   -100
      Value           =   0
      SliderWid_Height=   400
      BaseColor       =   14737632
      MidColor        =   4210752
      ValueVis        =   0   'False
      ValueCol        =   16777215
   End
   Begin Project1.XPSlider XPSlider3 
      Height          =   300
      Left            =   390
      TabIndex        =   2
      Top             =   1215
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   529
      Max             =   1000
      Value           =   0
      SliderWid_Height=   400
      MidColor        =   16777214
   End
   Begin Project1.XPSlider XPSlider2 
      Height          =   300
      Left            =   390
      TabIndex        =   1
      Top             =   585
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   529
      Value           =   0
      SliderWid_Height=   400
      BaseColor       =   32768
      MidColor        =   16777214
      ValueCol        =   255
   End
   Begin Project1.XPSlider XPSlider1 
      Height          =   300
      Left            =   390
      TabIndex        =   0
      Top             =   255
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   529
      Max             =   255
      Value           =   0
      SliderWid_Height=   400
      BaseColor       =   255
      MidColor        =   16777214
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   3285
      TabIndex        =   5
      Top             =   1620
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub XPSlider4_Changed()
Label1.Caption = XPSlider4.Value
End Sub
