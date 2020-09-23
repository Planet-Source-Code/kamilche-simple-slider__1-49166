VERSION 5.00
Object = "{A769E682-6C6B-48D2-86D9-F4DC7A1AB1F7}#19.0#0"; "MySliderProj.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000016&
   Caption         =   "Form1"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   ScaleHeight     =   312
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   308
   StartUpPosition =   3  'Windows Default
   Begin MySliderProj.MySlider MySlider1 
      Height          =   390
      Index           =   0
      Left            =   345
      TabIndex        =   1
      Top             =   90
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   688
      Value           =   100
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   4305
      Width           =   4620
      _ExtentX        =   8149
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MySliderProj.MySlider MySlider1 
      Height          =   2685
      Index           =   1
      Left            =   4005
      TabIndex        =   2
      Top             =   990
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   4736
      Value           =   100
      Orientation     =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MySlider1_Change(Index As Integer, ByVal NewValue As Long)
    StatusBar1.SimpleText = "Index " & Index & ", NewValue " & NewValue
End Sub
