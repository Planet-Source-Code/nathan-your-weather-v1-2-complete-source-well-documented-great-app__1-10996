VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "Msinet.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Begin VB.Form Mainfrm 
   BackColor       =   &H80000009&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Your Weather v1.2"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4350
   BeginProperty Font 
      Name            =   "Small Fonts"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   Picture         =   "Form1.frx":030A
   ScaleHeight     =   5160
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   53
      Top             =   4965
      Width           =   4350
      _ExtentX        =   7673
      _ExtentY        =   344
      SimpleText      =   ""
      ShowTips        =   0   'False
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   4939
            MinWidth        =   4939
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2681
            MinWidth        =   2681
            TextSave        =   "8/24/2000"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Ziptxt 
      BackColor       =   &H80000004&
      Height          =   240
      Left            =   1560
      MaxLength       =   5
      MousePointer    =   3  'I-Beam
      TabIndex        =   50
      Top             =   60
      Width           =   495
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   4800
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":802E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Zip Code:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   656
      TabIndex        =   52
      Top             =   90
      Width           =   855
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   3049
      X2              =   3649
      Y1              =   60
      Y2              =   60
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Update"
      Height          =   255
      Left            =   3049
      TabIndex        =   51
      Top             =   90
      Width           =   615
   End
   Begin VB.Line Line6 
      X1              =   3649
      X2              =   3649
      Y1              =   60
      Y2              =   280
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000009&
      X1              =   3049
      X2              =   3049
      Y1              =   260
      Y2              =   60
   End
   Begin VB.Line Line3 
      X1              =   3049
      X2              =   3649
      Y1              =   270
      Y2              =   270
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000006&
      X1              =   1215
      X2              =   3255
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label Report 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   49
      Top             =   360
      Width           =   4335
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Forecast"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2535
      TabIndex        =   48
      Top             =   3045
      Width           =   1335
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Low"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1935
      TabIndex        =   47
      Top             =   3045
      Width           =   495
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "High"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1335
      TabIndex        =   46
      Top             =   3045
      Width           =   495
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000007&
      X1              =   1215
      X2              =   1215
      Y1              =   3000
      Y2              =   4800
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000007&
      X1              =   2175
      X2              =   2175
      Y1              =   720
      Y2              =   3000
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Temperature:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   975
      TabIndex        =   45
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Wind:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   975
      TabIndex        =   44
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Humidity:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   975
      TabIndex        =   43
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Dewpoint:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1215
      TabIndex        =   42
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Visibility:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1335
      TabIndex        =   41
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Barometer:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   975
      TabIndex        =   40
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Sunrise:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1455
      TabIndex        =   39
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Sunset:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1455
      TabIndex        =   38
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Conditions:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1095
      TabIndex        =   37
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Day7_hi 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   1455
      TabIndex        =   36
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Day7_lo 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   2055
      TabIndex        =   35
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Day7_Weather 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   2535
      TabIndex        =   34
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label Day6_hi 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   1455
      TabIndex        =   33
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Day6_lo 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   2055
      TabIndex        =   32
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Day6_Weather 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   2535
      TabIndex        =   31
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Day5_hi 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   1455
      TabIndex        =   30
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Day5_lo 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   2055
      TabIndex        =   29
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Day5_Weather 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   2535
      TabIndex        =   28
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Day4_hi 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   1455
      TabIndex        =   27
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Day4_lo 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   2055
      TabIndex        =   26
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Day4_Weather 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   2535
      TabIndex        =   25
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Day3_hi 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   1455
      TabIndex        =   24
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Day3_lo 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   2055
      TabIndex        =   23
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Day3_Weather 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   2535
      TabIndex        =   22
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Day2_hi 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   1455
      TabIndex        =   21
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Day2_lo 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   2055
      TabIndex        =   20
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Day2_Weather 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   2535
      TabIndex        =   19
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Day1_Weather 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   2535
      TabIndex        =   18
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Day1_lo 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   2055
      TabIndex        =   17
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Day1_hi 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   1455
      TabIndex        =   16
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Day7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   735
      TabIndex        =   15
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label Day6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   735
      TabIndex        =   14
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Day5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   735
      TabIndex        =   13
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Day4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   735
      TabIndex        =   12
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Day3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   735
      TabIndex        =   11
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Day2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   735
      TabIndex        =   10
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Day1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   735
      TabIndex        =   9
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Temperature 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   2295
      TabIndex        =   8
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Wind 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   2295
      TabIndex        =   7
      Top             =   1200
      Width           =   2040
   End
   Begin VB.Label Dewpoint 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   2295
      TabIndex        =   6
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Humidity 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   2295
      TabIndex        =   5
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Visibility 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   2295
      TabIndex        =   4
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Barometer 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   2295
      TabIndex        =   3
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Sunrise 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   2295
      TabIndex        =   2
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Sunset 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   2295
      TabIndex        =   1
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Conditions 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   2295
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
End
Attribute VB_Name = "Mainfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub load_data()
    ' Loads data from weather.dat, this data is automatically saved each '
    ' time you update.                                                   '
    iniPath$ = App.Path + "\weather.dat"
    Report.Caption = GetFromINI("Report", "Area", iniPath$)
    Ziptxt.Text = GetFromINI("Report", "Zip", iniPath$)
    Conditions.Caption = GetFromINI("Current", "Conditions", iniPath$)
    Temperature.Caption = GetFromINI("Current", "Temperature", iniPath$)
    Wind.Caption = GetFromINI("Current", "Wind", iniPath$)
    Humidity.Caption = GetFromINI("Current", "Humidity", iniPath$)
    Barometer.Caption = GetFromINI("Current", "Barometer", iniPath$)
    Dewpoint.Caption = GetFromINI("Current", "Dewpoint", iniPath$)
    Visibility.Caption = GetFromINI("Current", "Visibility", iniPath$)
    Sunrise.Caption = GetFromINI("Current", "Sunrise", iniPath$)
    Sunset.Caption = GetFromINI("Current", "Sunset", iniPath$)
    Day1.Caption = GetFromINI("Weekday", "Day1", iniPath$)
    Day2.Caption = GetFromINI("Weekday", "Day2", iniPath$)
    Day3.Caption = GetFromINI("Weekday", "Day3", iniPath$)
    Day4.Caption = GetFromINI("Weekday", "Day4", iniPath$)
    Day5.Caption = GetFromINI("Weekday", "Day5", iniPath$)
    Day6.Caption = GetFromINI("Weekday", "Day6", iniPath$)
    Day7.Caption = GetFromINI("Weekday", "Day7", iniPath$)
    Day1_hi.Caption = GetFromINI("High Temp", "Day1", iniPath$)
    Day2_hi.Caption = GetFromINI("High Temp", "Day2", iniPath$)
    Day3_hi.Caption = GetFromINI("High Temp", "Day3", iniPath$)
    Day4_hi.Caption = GetFromINI("High Temp", "Day4", iniPath$)
    Day5_hi.Caption = GetFromINI("High Temp", "Day5", iniPath$)
    Day6_hi.Caption = GetFromINI("High Temp", "Day6", iniPath$)
    Day7_hi.Caption = GetFromINI("High Temp", "Day7", iniPath$)
    Day1_lo.Caption = GetFromINI("Low Temp", "Day1", iniPath$)
    Day2_lo.Caption = GetFromINI("Low Temp", "Day2", iniPath$)
    Day3_lo.Caption = GetFromINI("Low Temp", "Day3", iniPath$)
    Day4_lo.Caption = GetFromINI("Low Temp", "Day4", iniPath$)
    Day5_lo.Caption = GetFromINI("Low Temp", "Day5", iniPath$)
    Day6_lo.Caption = GetFromINI("Low Temp", "Day6", iniPath$)
    Day7_lo.Caption = GetFromINI("Low Temp", "Day7", iniPath$)
    Day1_Weather.Caption = GetFromINI("Weather", "Day1", iniPath$)
    Day2_Weather.Caption = GetFromINI("Weather", "Day2", iniPath$)
    Day3_Weather.Caption = GetFromINI("Weather", "Day3", iniPath$)
    Day4_Weather.Caption = GetFromINI("Weather", "Day4", iniPath$)
    Day5_Weather.Caption = GetFromINI("Weather", "Day5", iniPath$)
    Day6_Weather.Caption = GetFromINI("Weather", "Day6", iniPath$)
    Day7_Weather.Caption = GetFromINI("Weather", "Day7", iniPath$)
    StatusBar1.Panels.Item(1).Text = "Updated on " & GetFromINI("Report", "Time", iniPath$)
End Sub
Sub save_data()
    ' Saves all information to weather.dat to be loaded later.  This is   '
    ' useful so that you don't have to update each time just to view the  '
    ' same data.  Although you must udpate at least once a day to get the '
    ' most recent and accuarate data.                                     '
    iniPath$ = App.Path + "\weather.dat"
    entry$ = Date & " at " & Time
    r% = WritePrivateProfileString("Report", "Time", entry$, iniPath$)
    entry$ = Report.Caption
    r% = WritePrivateProfileString("Report", "Area", entry$, iniPath$)
    entry$ = Ziptxt.Text
    r% = WritePrivateProfileString("Report", "Zip", entry$, iniPath$)
    entry$ = Conditions.Caption
    r% = WritePrivateProfileString("Current", "Conditions", entry$, iniPath$)
    entry$ = Temperature.Caption
    r% = WritePrivateProfileString("Current", "Temperature", entry$, iniPath$)
    entry$ = Wind.Caption
    r% = WritePrivateProfileString("Current", "Wind", entry$, iniPath$)
    entry$ = Humidity.Caption
    r% = WritePrivateProfileString("Current", "Humidity", entry$, iniPath$)
    entry$ = Barometer.Caption
    r% = WritePrivateProfileString("Current", "Barometer", entry$, iniPath$)
    entry$ = Dewpoint.Caption
    r% = WritePrivateProfileString("Current", "Dewpoint", entry$, iniPath$)
    entry$ = Visibility.Caption
    r% = WritePrivateProfileString("Current", "Visibility", entry$, iniPath$)
    entry$ = Sunrise.Caption
    r% = WritePrivateProfileString("Current", "Sunrise", entry$, iniPath$)
    entry$ = Sunset.Caption
    r% = WritePrivateProfileString("Current", "Sunset", entry$, iniPath$)
    entry$ = Day1.Caption
    r% = WritePrivateProfileString("Weekday", "Day1", entry$, iniPath$)
    entry$ = Day2.Caption
    r% = WritePrivateProfileString("Weekday", "Day2", entry$, iniPath$)
    entry$ = Day3.Caption
    r% = WritePrivateProfileString("Weekday", "Day3", entry$, iniPath$)
    entry$ = Day4.Caption
    r% = WritePrivateProfileString("Weekday", "Day4", entry$, iniPath$)
    entry$ = Day5.Caption
    r% = WritePrivateProfileString("Weekday", "Day5", entry$, iniPath$)
    entry$ = Day6.Caption
    r% = WritePrivateProfileString("Weekday", "Day6", entry$, iniPath$)
    entry$ = Day7.Caption
    r% = WritePrivateProfileString("Weekday", "Day7", entry$, iniPath$)
    entry$ = Day1_hi.Caption
    r% = WritePrivateProfileString("High Temp", "Day1", entry$, iniPath$)
    entry$ = Day2_hi.Caption
    r% = WritePrivateProfileString("High Temp", "Day2", entry$, iniPath$)
    entry$ = Day3_hi.Caption
    r% = WritePrivateProfileString("High Temp", "Day3", entry$, iniPath$)
    entry$ = Day4_hi.Caption
    r% = WritePrivateProfileString("High Temp", "Day4", entry$, iniPath$)
    entry$ = Day5_hi.Caption
    r% = WritePrivateProfileString("High Temp", "Day5", entry$, iniPath$)
    entry$ = Day6_hi.Caption
    r% = WritePrivateProfileString("High Temp", "Day6", entry$, iniPath$)
    entry$ = Day7_hi.Caption
    r% = WritePrivateProfileString("High Temp", "Day7", entry$, iniPath$)
    entry$ = Day1_lo.Caption
    r% = WritePrivateProfileString("Low Temp", "Day1", entry$, iniPath$)
    entry$ = Day2_lo.Caption
    r% = WritePrivateProfileString("Low Temp", "Day2", entry$, iniPath$)
    entry$ = Day3_lo.Caption
    r% = WritePrivateProfileString("Low Temp", "Day3", entry$, iniPath$)
    entry$ = Day4_lo.Caption
    r% = WritePrivateProfileString("Low Temp", "Day4", entry$, iniPath$)
    entry$ = Day5_lo.Caption
    r% = WritePrivateProfileString("Low Temp", "Day5", entry$, iniPath$)
    entry$ = Day6_lo.Caption
    r% = WritePrivateProfileString("Low Temp", "Day6", entry$, iniPath$)
    entry$ = Day7_lo.Caption
    r% = WritePrivateProfileString("Low Temp", "Day7", entry$, iniPath$)
    entry$ = Day1_Weather.Caption
    r% = WritePrivateProfileString("Weather", "Day1", entry$, iniPath$)
    entry$ = Day2_Weather.Caption
    r% = WritePrivateProfileString("Weather", "Day2", entry$, iniPath$)
    entry$ = Day3_Weather.Caption
    r% = WritePrivateProfileString("Weather", "Day3", entry$, iniPath$)
    entry$ = Day4_Weather.Caption
    r% = WritePrivateProfileString("Weather", "Day4", entry$, iniPath$)
    entry$ = Day5_Weather.Caption
    r% = WritePrivateProfileString("Weather", "Day5", entry$, iniPath$)
    entry$ = Day6_Weather.Caption
    r% = WritePrivateProfileString("Weather", "Day6", entry$, iniPath$)
    entry$ = Day7_Weather.Caption
    r% = WritePrivateProfileString("Weather", "Day7", entry$, iniPath$)
    StatusBar1.Panels.Item(1).Text = "Updated on " & GetFromINI("Report", "Time", iniPath$)
    Label11.Enabled = True
End Sub
Sub LoadWeather()
    ' ************************************************************************    '
    ' Example : we have the string "abcdefghijklmnopqrstuvwxyz"                   '
    '                                                                             '
    ' We want to get "hijkl"                                                      '
    '                                                                             '
    ' First thing that we do is search for "abcdefg"                              '
    '                                                                             '
    ' Next we say goto the position of occurence of "abcdefg" (which = 1)         '
    ' and add length of our string we searched for "abcdefg" which = 7.           '
    ' (7 + 1 = 8)                                                                 '
    '                                                                             '
    ' So now we have a position of "8" which happens to be the 1st character      '
    ' we need.                                                                    '
    '                                                                             '
    ' All we need now is the character(s) that follow the string that we want,    '
    ' which is "m".                                                               '
    '                                                                             '
    ' Now we do a search for "m" in the string that we have have left which is    '
    ' "hijklmnopqrstuvwxyz".  "m" is found at position 6 and we now say           '
    ' (6 - 1 = 5) so get the first 5 characters in the string.                    '
    ' you get "hijkl"                                                             '
    ' *************************************************************************** '
    '                                                                             '
    ' Goes out to "http://www.weather.com" and gets the source information        '
    ' from the website for the zipcode you entered. You can also view the         '
    ' source by right clicking on the web page and going to "View Source"         '
    ' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''  '
    Label11.Enabled = False
    On Error GoTo Weather_Error
    StatusBar1.Panels.Item(1).Text = "Contacting www.weather.com"
    DoEvents
    weather_text = Inet.OpenURL("http://www.weather.com/weather/us/zips/" & Ziptxt.Text & ".html")   '
    If weather_text = "" Then GoTo Weather_Error
    ' Get Location '
    StatusBar1.Panels.Item(1).Text = "Getting Weather Data"
    DoEvents
    tempo = "<TD HEIGHT=20 ALIGN=" & Chr(34) & "center" & Chr(34) & " VALIGN=" & Chr(34) & "top" & Chr(34) & "><FONT FACE=" & Chr(34) & "Arial, Helvetica, Chicago, Sans Serif" & Chr(34) & " SIZE=1>"
    pos1 = InStr(1, weather_text, tempo)
    report_info = Mid(weather_text, pos1 + Len(tempo))
    report_info = Mid$(report_info, 1, InStr(report_info, "<BR>") - 1)
    Mainfrm.Report = "Conditions " & report_info
    ' Get Todays Forecast '
    If InStr(weather_text, "<B>TODAY</B>") > 0 Then
        temp_day = "<B>TODAY</B>"
        pos1 = InStr(weather_text, temp_day)
        tempo_string = "<TD ALIGN=" & Chr(34) & "center" & Chr(34) & " VALIGN=" & Chr(34) & "middle" & Chr(34) & " WIDTH=" & Chr(34) & "65" & Chr(34) & " BGCOLOR=" & Chr(34) & "#E4ECF4" & Chr(34) & "><FONT FACE=" & Chr(34) & "Arial, Helvetica, Chicago, Sans Serif" & Chr(34) & " SIZE=" & Chr(34) & "2" & Chr(34) & ">"
        pos2 = InStr(pos1, weather_text, tempo_string)
        today_weather = Mid(weather_text, pos2 + Len(tempo_string))
        tempo_string = InStr(today_weather, "<")
        today_weather = Mid(today_weather, 1, tempo_string - 1)
    ' Get Tonights Forecast ' (Special thanks to Don Borchert for finding this bug) '
    End If
    If InStr(weather_text, "<B>TONIGHT</B>") > 0 Then
        temp_day = "<B>TONIGHT</B>"                                                                                                                                                                                                                                                                                             '
        pos1 = InStr(weather_text, temp_day)                                                                                                                                                                                                                                                                                    '
        tempo_string = "<TD ALIGN=" & Chr(34) & "center" & Chr(34) & " VALIGN=" & Chr(34) & "middle" & Chr(34) & " WIDTH=" & Chr(34) & "65" & Chr(34) & " BGCOLOR=" & Chr(34) & "#E4ECF4" & Chr(34) & "><FONT FACE=" & Chr(34) & "Arial, Helvetica, Chicago, Sans Serif" & Chr(34) & " SIZE=" & Chr(34) & "2" & Chr(34) & ">"   '
        pos2 = InStr(pos1, weather_text, tempo_string)                                                                                                                                                                                                                                                                          '
        today_weather = Mid(weather_text, pos2 + Len(tempo_string))                                                                                                                                                                                                                                                             '
        tempo_string = InStr(today_weather, "<")                                                                                                                                                                                                                                                                                '
        today_weather = Mid(today_weather, 1, tempo_string - 1)                                                                                                                                                                                                                                                                 '
    End If                                                                                                                                                                                                                                                                                                                      '
    ' Get Today High Temperature '
    temp_hi = "<NOBR>hi&nbsp;"
    pos3 = InStr(pos1, weather_text, temp_hi)
    Today_hi = Mid(weather_text, pos3 + Len(temp_hi))
    temp_hi = InStr(Today_hi, "&")
    Today_hi = Mid(Today_hi, 1, temp_hi - 1)
    ' Get Today Low Temperature '
    temp_lo = "<NOBR>lo&nbsp;"
    pos3 = InStr(pos1, weather_text, temp_lo)
    today_lo = Mid(weather_text, pos3 + Len(temp_lo))
    temp_lo = InStr(today_lo, "&")
    today_lo = Mid(today_lo, 1, temp_lo - 1)
    ' Current Conditions '
    tempo = "<FONT FACE=" & Chr(34) & "Arial, Helvetica, Chicago, Sans Serif" & Chr(34) & " SIZE=3><B>"
    pos1 = InStr(1, weather_text, tempo)
    current_condition = Mid(weather_text, pos1 + Len(tempo))
    current_condition = Mid$(current_condition, 1, InStr(current_condition, "</B>") - 1)
    Mainfrm.Conditions = current_condition
    ' Current Temp '
    tempo = "Temp:</B></FONT></TD>" & Chr(10) & "<TD WIDTH=5><IMG WIDTH=5 HEIGHT=1 SRC=" & Chr(34) & "http://image.weather.com/pics/blank.gif" & Chr(34) & " ALT=" & Chr(34) & Chr(34) & "></TD>" & Chr(10) & "          <TD WIDTH=90><FONT FACE=" & Chr(34) & "Arial, Helvetica, Chicago, Sans Serif" & Chr(34) & " SIZE=2>"
    pos1 = InStr(weather_text, tempo)
    current_temp = Mid(weather_text, pos1 + Len(tempo))
    current_temp = Mid(current_temp, 1, InStr(current_temp, "&") - 1)
    Mainfrm.Temperature.Caption = current_temp
    ' Current Wind '
    tempo = "Wind:</B></FONT></TD>" & Chr(10) & "<TD WIDTH=5><IMG WIDTH=5 HEIGHT=1 SRC=" & Chr(34) & "http://image.weather.com/pics/blank.gif" & Chr(34) & " ALT=" & Chr(34) & Chr(34) & "></TD>" & Chr(10) & "          <TD WIDTH=90><FONT FACE=" & Chr(34) & "Arial, Helvetica, Chicago, Sans Serif" & Chr(34) & " SIZE=2>"
    pos1 = InStr(weather_text, tempo)
    current_wind = Mid(weather_text, pos1 + Len(tempo))
    tempo = InStr(current_wind, "<")
    current_wind = Mid(current_wind, 1, tempo - 1)
    Mainfrm.Wind = current_wind
    ' Current Dewpoint '
    tempo = "Dewpoint:</B></FONT></TD>" & Chr(10) & "<TD WIDTH=5><IMG WIDTH=5 HEIGHT=1 SRC=" & Chr(34) & "http://image.weather.com/pics/blank.gif" & Chr(34) & " ALT=" & Chr(34) & Chr(34) & "></TD>" & Chr(10) & "          <TD WIDTH=90><FONT FACE=" & Chr(34) & "Arial, Helvetica, Chicago, Sans Serif" & Chr(34) & " SIZE=2>"
    pos1 = InStr(weather_text, tempo)
    current_dewpoint = Mid(weather_text, pos1 + Len(tempo))
    tempo = InStr(current_dewpoint, "&")
    current_dewpoint = Mid(current_dewpoint, 1, tempo - 1)
    Mainfrm.Dewpoint = current_dewpoint
    ' Current Relative Humidity '
    tempo = "Rel. Humidity:</B></FONT></TD>" & Chr(10) & "<TD WIDTH=5><IMG WIDTH=5 HEIGHT=1 SRC=" & Chr(34) & "http://image.weather.com/pics/blank.gif" & Chr(34) & " ALT=" & Chr(34) & Chr(34) & "></TD>" & Chr(10) & "          <TD WIDTH=90><FONT FACE=" & Chr(34) & "Arial, Helvetica, Chicago, Sans Serif" & Chr(34) & " SIZE=2>"
    pos1 = InStr(weather_text, tempo)
    current_humidity = Mid(weather_text, pos1 + Len(tempo))
    tempo = InStr(current_humidity, "<")
    current_humidity = Mid(current_humidity, 1, tempo - 1)
    Mainfrm.Humidity = current_humidity
    ' Current Visibility '
    tempo = "Visibility:</B></FONT></TD>" & Chr(10) & "<TD WIDTH=5><IMG WIDTH=5 HEIGHT=1 SRC=" & Chr(34) & "http://image.weather.com/pics/blank.gif" & Chr(34) & " ALT=" & Chr(34) & Chr(34) & "></TD>" & Chr(10) & "          <TD WIDTH=90><FONT FACE=" & Chr(34) & "Arial, Helvetica, Chicago, Sans Serif" & Chr(34) & " SIZE=2>"
    pos1 = InStr(weather_text, tempo)
    current_visibility = Mid(weather_text, pos1 + Len(tempo))
    tempo = InStr(current_visibility, "<")
    current_visibility = Mid(current_visibility, 1, tempo - 1)
    Mainfrm.Visibility = current_visibility
    ' Current Barometric Pressure '
    tempo = "Barometer:</B></FONT></TD>" & Chr(10) & "<TD WIDTH=5><IMG WIDTH=5 HEIGHT=1 SRC=" & Chr(34) & "http://image.weather.com/pics/blank.gif" & Chr(34) & " ALT=" & Chr(34) & Chr(34) & "></TD>" & Chr(10) & "          <TD WIDTH=90><FONT FACE=" & Chr(34) & "Arial, Helvetica, Chicago, Sans Serif" & Chr(34) & " SIZE=2>"
    pos1 = InStr(weather_text, tempo)
    current_barometer = Mid(weather_text, pos1 + Len(tempo))
    tempo = InStr(current_barometer, "<")
    current_barometer = Mid(current_barometer, 1, tempo - 1)
    Mainfrm.Barometer = current_barometer
    ' Sunrise For Current Day '
    tempo = "Sunrise:</B></FONT></TD>" & Chr(10) & "<TD WIDTH=5><IMG WIDTH=5 HEIGHT=1 SRC=" & Chr(34) & "http://image.weather.com/pics/blank.gif" & Chr(34) & " ALT=" & Chr(34) & Chr(34) & "></TD>" & Chr(10) & "          <TD WIDTH=90><FONT FACE=" & Chr(34) & "Arial, Helvetica, Chicago, Sans Serif" & Chr(34) & " SIZE=2>"
    pos1 = InStr(weather_text, tempo)
    current_sunrise = Mid(weather_text, pos1 + Len(tempo))
    tempo = InStr(current_sunrise, "<")
    current_sunrise = Mid(current_sunrise, 1, tempo - 1)
    Mainfrm.Sunrise = current_sunrise
    ' Sunset For Current Day '
    tempo = "Sunset:</B></FONT></TD>" & Chr(10) & "<TD WIDTH=5><IMG WIDTH=5 HEIGHT=1 SRC=" & Chr(34) & "http://image.weather.com/pics/blank.gif" & Chr(34) & " ALT=" & Chr(34) & Chr(34) & "></TD>" & Chr(10) & "          <TD WIDTH=90><FONT FACE=" & Chr(34) & "Arial, Helvetica, Chicago, Sans Serif" & Chr(34) & " SIZE=2>"
    pos1 = InStr(weather_text, tempo)
    current_sunset = Mid(weather_text, pos1 + Len(tempo))
    tempo = InStr(current_sunset, "<")
    current_sunset = Mid(current_sunset, 1, tempo - 1)
    Mainfrm.Sunset = current_sunset
    ' Five Day Forecast                 '
    ' Get Forecasted Weather For Sunday '
    If WeekDay(Now) <> 1 Then
        temp_day = "<B>SUN</B>"
        pos1 = InStr(weather_text, temp_day)
        tempo_string = "<TD ALIGN=" & Chr(34) & "center" & Chr(34) & " VALIGN=" & Chr(34) & "middle" & Chr(34) & " WIDTH=" & Chr(34) & "65" & Chr(34) & " BGCOLOR=" & Chr(34) & "#E4ECF4" & Chr(34) & "><FONT FACE=" & Chr(34) & "Arial, Helvetica, Chicago, Sans Serif" & Chr(34) & " SIZE=" & Chr(34) & "2" & Chr(34) & ">"
        pos2 = InStr(pos1, weather_text, tempo_string)
        sun_weather = Mid(weather_text, pos2 + Len(tempo_string))
        tempo_string = InStr(sun_weather, "<")
        sun_weather = Mid(sun_weather, 1, tempo_string - 1)
        ' Get Sundays High Temperature '
        temp_hi = "<NOBR>hi&nbsp;"
        pos3 = InStr(pos1, weather_text, temp_hi)
        sun_hi = Mid(weather_text, pos3 + Len(temp_hi))
        temp_hi = InStr(sun_hi, "&")
        sun_hi = Mid(sun_hi, 1, temp_hi - 1)
        ' Get Sundays Low Temperature '
        temp_lo = "<NOBR>lo&nbsp;"
        pos3 = InStr(pos1, weather_text, temp_lo)
        sun_lo = Mid(weather_text, pos3 + Len(temp_lo))
        temp_lo = InStr(sun_lo, "&")
        sun_lo = Mid(sun_lo, 1, temp_lo - 1)
    End If
    ' Get Forecasted Weather For Monday '
    If WeekDay(Now) <> 2 Then
        temp_day = "<B>MON</B>"
        pos1 = InStr(weather_text, temp_day)
        tempo_string = "<TD ALIGN=" & Chr(34) & "center" & Chr(34) & " VALIGN=" & Chr(34) & "middle" & Chr(34) & " WIDTH=" & Chr(34) & "65" & Chr(34) & " BGCOLOR=" & Chr(34) & "#E4ECF4" & Chr(34) & "><FONT FACE=" & Chr(34) & "Arial, Helvetica, Chicago, Sans Serif" & Chr(34) & " SIZE=" & Chr(34) & "2" & Chr(34) & ">"
        pos2 = InStr(pos1, weather_text, tempo_string)
        mon_weather = Mid(weather_text, pos2 + Len(tempo_string))
        tempo_string = InStr(mon_weather, "<")
        mon_weather = Mid(mon_weather, 1, tempo_string - 1)
        ' Get Monday High Temperature '
        temp_hi = "<NOBR>hi&nbsp;"
        pos3 = InStr(pos1, weather_text, temp_hi)
        mon_hi = Mid(weather_text, pos3 + Len(temp_hi))
        temp_hi = InStr(mon_hi, "&")
        mon_hi = Mid(mon_hi, 1, temp_hi - 1)
        ' Get Monday Low Temperature '
        temp_lo = "<NOBR>lo&nbsp;"
        pos3 = InStr(pos1, weather_text, temp_lo)
        mon_lo = Mid(weather_text, pos3 + Len(temp_lo))
        temp_lo = InStr(mon_lo, "&")
        mon_lo = Mid(mon_lo, 1, temp_lo - 1)
    End If
    ' Get Forecasted Weather For Tuesday '
    If WeekDay(Now) <> 3 Then
        temp_day = "<B>TUE</B>"
        pos1 = InStr(weather_text, temp_day)
        tempo_string = "<TD ALIGN=" & Chr(34) & "center" & Chr(34) & " VALIGN=" & Chr(34) & "middle" & Chr(34) & " WIDTH=" & Chr(34) & "65" & Chr(34) & " BGCOLOR=" & Chr(34) & "#E4ECF4" & Chr(34) & "><FONT FACE=" & Chr(34) & "Arial, Helvetica, Chicago, Sans Serif" & Chr(34) & " SIZE=" & Chr(34) & "2" & Chr(34) & ">"
        pos2 = InStr(pos1, weather_text, tempo_string)
        tue_weather = Mid(weather_text, pos2 + Len(tempo_string))
        tempo_string = InStr(tue_weather, "<")
        tue_weather = Mid(tue_weather, 1, tempo_string - 1)
        ' Get Tuesday High Temperature '
        temp_hi = "<NOBR>hi&nbsp;"
        pos3 = InStr(pos1, weather_text, temp_hi)
        tue_hi = Mid(weather_text, pos3 + Len(temp_hi))
        temp_hi = InStr(tue_hi, "&")
        tue_hi = Mid(tue_hi, 1, temp_hi - 1)
        ' Get Tuesday Low Temperature '
        temp_lo = "<NOBR>lo&nbsp;"
        pos3 = InStr(pos1, weather_text, temp_lo)
        tue_lo = Mid(weather_text, pos3 + Len(temp_lo))
        temp_lo = InStr(tue_lo, "&")
        tue_lo = Mid(tue_lo, 1, temp_lo - 1)
    End If
    ' Get Forecasted Weather For Wednesday '
    If WeekDay(Now) <> 4 Then
        temp_day = "<B>WED</B>"
        pos1 = InStr(weather_text, temp_day)
        tempo_string = "<TD ALIGN=" & Chr(34) & "center" & Chr(34) & " VALIGN=" & Chr(34) & "middle" & Chr(34) & " WIDTH=" & Chr(34) & "65" & Chr(34) & " BGCOLOR=" & Chr(34) & "#E4ECF4" & Chr(34) & "><FONT FACE=" & Chr(34) & "Arial, Helvetica, Chicago, Sans Serif" & Chr(34) & " SIZE=" & Chr(34) & "2" & Chr(34) & ">"
        pos2 = InStr(pos1, weather_text, tempo_string)
        wed_weather = Mid(weather_text, pos2 + Len(tempo_string))
        tempo_string = InStr(wed_weather, "<")
        wed_weather = Mid(wed_weather, 1, tempo_string - 1)
        ' Get Wednesday High Temperature '
        temp_hi = "<NOBR>hi&nbsp;"
        pos3 = InStr(pos1, weather_text, temp_hi)
        wed_hi = Mid(weather_text, pos3 + Len(temp_hi))
        temp_hi = InStr(wed_hi, "&")
        wed_hi = Mid(wed_hi, 1, temp_hi - 1)
        ' Get Wednesday Low Temperature '
        temp_lo = "<NOBR>lo&nbsp;"
        pos3 = InStr(pos1, weather_text, temp_lo)
        wed_lo = Mid(weather_text, pos3 + Len(temp_lo))
        temp_lo = InStr(wed_lo, "&")
        wed_lo = Mid(wed_lo, 1, temp_lo - 1)
    End If
    ' Get Forecasted Weather For Thursday '
    If WeekDay(Now) <> 5 Then
        temp_day = "<B>THU</B>"
        pos1 = InStr(weather_text, temp_day)
        tempo_string = "<TD ALIGN=" & Chr(34) & "center" & Chr(34) & " VALIGN=" & Chr(34) & "middle" & Chr(34) & " WIDTH=" & Chr(34) & "65" & Chr(34) & " BGCOLOR=" & Chr(34) & "#E4ECF4" & Chr(34) & "><FONT FACE=" & Chr(34) & "Arial, Helvetica, Chicago, Sans Serif" & Chr(34) & " SIZE=" & Chr(34) & "2" & Chr(34) & ">"
        pos2 = InStr(pos1, weather_text, tempo_string)
        thu_weather = Mid(weather_text, pos2 + Len(tempo_string))
        tempo_string = InStr(thu_weather, "<")
        thu_weather = Mid(thu_weather, 1, tempo_string - 1)
        ' Get Thursday High Temperature '
        temp_hi = "<NOBR>hi&nbsp;"
        pos3 = InStr(pos1, weather_text, temp_hi)
        thu_hi = Mid(weather_text, pos3 + Len(temp_hi))
        temp_hi = InStr(thu_hi, "&")
        thu_hi = Mid(thu_hi, 1, temp_hi - 1)
        ' Get Thursday Low Temperature '
        temp_lo = "<NOBR>lo&nbsp;"
        pos3 = InStr(pos1, weather_text, temp_lo)
        thu_lo = Mid(weather_text, pos3 + Len(temp_lo))
        temp_lo = InStr(thu_lo, "&")
        thu_lo = Mid(thu_lo, 1, temp_lo - 1)
    End If
    ' Get Forecasted Weather For Friday '
    If WeekDay(Now) <> 6 Then
        temp_day = "<B>FRI</B>"
        pos1 = InStr(weather_text, temp_day)
        tempo_string = "<TD ALIGN=" & Chr(34) & "center" & Chr(34) & " VALIGN=" & Chr(34) & "middle" & Chr(34) & " WIDTH=" & Chr(34) & "65" & Chr(34) & " BGCOLOR=" & Chr(34) & "#E4ECF4" & Chr(34) & "><FONT FACE=" & Chr(34) & "Arial, Helvetica, Chicago, Sans Serif" & Chr(34) & " SIZE=" & Chr(34) & "2" & Chr(34) & ">"
        pos2 = InStr(pos1, weather_text, tempo_string)
        fri_weather = Mid(weather_text, pos2 + Len(tempo_string))
        tempo_string = InStr(fri_weather, "<")
        fri_weather = Mid(fri_weather, 1, tempo_string - 1)
        ' Get Friday High Temperature '
        temp_hi = "<NOBR>hi&nbsp;"
        pos3 = InStr(pos1, weather_text, temp_hi)
        fri_hi = Mid(weather_text, pos3 + Len(temp_hi))
        temp_hi = InStr(fri_hi, "&")
        fri_hi = Mid(fri_hi, 1, temp_hi - 1)
        ' Get Friday Low Temperature '
        temp_lo = "<NOBR>lo&nbsp;"
        pos3 = InStr(pos1, weather_text, temp_lo)
        fri_lo = Mid(weather_text, pos3 + Len(temp_lo))
        temp_lo = InStr(fri_lo, "&")
        fri_lo = Mid(fri_lo, 1, temp_lo - 1)
    End If
    ' Get Forecasted Weather For Saturday '
    If WeekDay(Now) <> 7 Then
        temp_day = "<B>SAT</B>"
        pos1 = InStr(weather_text, temp_day)
        tempo_string = "<TD ALIGN=" & Chr(34) & "center" & Chr(34) & " VALIGN=" & Chr(34) & "middle" & Chr(34) & " WIDTH=" & Chr(34) & "65" & Chr(34) & " BGCOLOR=" & Chr(34) & "#E4ECF4" & Chr(34) & "><FONT FACE=" & Chr(34) & "Arial, Helvetica, Chicago, Sans Serif" & Chr(34) & " SIZE=" & Chr(34) & "2" & Chr(34) & ">"
        pos2 = InStr(pos1, weather_text, tempo_string)
        sat_weather = Mid(weather_text, pos2 + Len(tempo_string))
        tempo_string = InStr(sat_weather, "<")
        sat_weather = Mid(sat_weather, 1, tempo_string - 1)
        ' Get Saturday High Temperature '
        temp_hi = "<NOBR>hi&nbsp;"
        pos3 = InStr(pos1, weather_text, temp_hi)
        sat_hi = Mid(weather_text, pos3 + Len(temp_hi))
        temp_hi = InStr(sat_hi, "&")
        sat_hi = Mid(sat_hi, 1, temp_hi - 1)
        ' Get Saturday Low Temperature '
        temp_lo = "<NOBR>lo&nbsp;"
        pos3 = InStr(pos1, weather_text, temp_lo)
        sat_lo = Mid(weather_text, pos3 + Len(temp_lo))
        temp_lo = InStr(sat_lo, "&")
        sat_lo = Mid(sat_lo, 1, temp_lo - 1)
    End If
    ' Determines which day it is and puts days in order '
    today_weekday = WeekDay(Now)
    If today_weekday = 1 Then
        With Mainfrm
            .Day1.Caption = "Sun"
            .Day1_hi = Today_hi
            .Day1_lo = today_lo
            .Day1_Weather = today_weather
            .Day2.Caption = "Mon"
            .Day2_hi = mon_hi
            .Day2_lo = mon_lo
            .Day2_Weather = mon_weather
            .Day3.Caption = "Tue"
            .Day3_hi = tue_hi
            .Day3_lo = tue_lo
            .Day3_Weather = tue_weather
            .Day4.Caption = "Wed"
            .Day4_hi = wed_hi
            .Day4_lo = wed_lo
            .Day4_Weather = wed_weather
            .Day5.Caption = "Thu"
            .Day5_hi = thu_hi
            .Day5_lo = thu_lo
            .Day5_Weather = thu_weather
            .Day6.Caption = "Fri"
            .Day6_hi = fri_hi
            .Day6_lo = fri_lo
            .Day6_Weather = fri_weather
            .Day7.Caption = "Sat"
            .Day7_hi = sat_hi
            .Day7_lo = sat_lo
            .Day7_Weather = sat_weather
        End With
    ElseIf today_weekday = 2 Then
        With Mainfrm
            .Day1.Caption = "Mon"
            .Day1_hi = Today_hi
            .Day1_lo = today_lo
            .Day1_Weather = today_weather
            .Day2.Caption = "Tue"
            .Day2_hi = tue_hi
            .Day2_lo = tue_lo
            .Day2_Weather = tue_weather
            .Day3.Caption = "Wed"
            .Day3_hi = wed_hi
            .Day3_lo = wed_lo
            .Day3_Weather = wed_weather
            .Day4.Caption = "Thu"
            .Day4_hi = thu_hi
            .Day4_lo = thu_lo
            .Day4_Weather = thu_weather
            .Day5.Caption = "Fri"
            .Day5_hi = fri_hi
            .Day5_lo = fri_lo
            .Day5_Weather = fri_weather
            .Day6.Caption = "Sat"
            .Day6_hi = sat_hi
            .Day6_lo = sat_lo
            .Day6_Weather = sat_weather
            .Day7.Caption = "Sun"
            .Day7_hi = sun_hi
            .Day7_lo = sun_lo
            .Day7_Weather = sun_weather
        End With
    ElseIf today_weekday = 3 Then
        With Mainfrm
            .Day1.Caption = "Tue"
            .Day1_hi = Today_hi
            .Day1_lo = today_lo
            .Day1_Weather = today_weather
            .Day2.Caption = "Wed"
            .Day2_hi = wed_hi
            .Day2_lo = wed_lo
            .Day2_Weather = wed_weather
            .Day3.Caption = "Thu"
            .Day3_hi = thu_hi
            .Day3_lo = thu_lo
            .Day3_Weather = thu_weather
            .Day4.Caption = "Fri"
            .Day4_hi = fri_hi
            .Day4_lo = fri_lo
            .Day4_Weather = fri_weather
            .Day5.Caption = "Sat"
            .Day5_hi = sat_hi
            .Day5_lo = sat_lo
            .Day5_Weather = sat_weather
            .Day6.Caption = "Sun"
            .Day6_hi = sun_hi
            .Day6_lo = sun_lo
            .Day6_Weather = sun_weather
            .Day7.Caption = "Mon"
            .Day7_hi = mon_hi
            .Day7_lo = mon_lo
            .Day7_Weather = mon_weather
        End With
    ElseIf today_weekday = 4 Then
        With Mainfrm
            .Day1.Caption = "Wed"
            .Day1_hi = Today_hi
            .Day1_lo = today_lo
            .Day1_Weather = today_weather
            .Day2.Caption = "Thu"
            .Day2_hi = thu_hi
            .Day2_lo = thu_lo
            .Day2_Weather = thu_weather
            .Day3.Caption = "Fri"
            .Day3_hi = fri_hi
            .Day3_lo = fri_lo
            .Day3_Weather = fri_weather
            .Day4.Caption = "Sat"
            .Day4_hi = sat_hi
            .Day4_lo = sat_lo
            .Day4_Weather = sat_weather
            .Day5.Caption = "Sun"
            .Day5_hi = sun_hi
            .Day5_lo = sun_lo
            .Day5_Weather = sun_weather
            .Day6.Caption = "Mon"
            .Day6_hi = mon_hi
            .Day6_lo = mon_lo
            .Day6_Weather = mon_weather
            .Day7.Caption = "Tue"
            .Day7_hi = tue_hi
            .Day7_lo = tue_lo
            .Day7_Weather = tue_weather
        End With
    ElseIf today_weekday = 5 Then
        With Mainfrm
            .Day1.Caption = "Thu"
            .Day1_hi = Today_hi
            .Day1_lo = today_lo
            .Day1_Weather = today_weather
            .Day2.Caption = "Fri"
            .Day2_hi = fri_hi
            .Day2_lo = fri_lo
            .Day2_Weather = fri_weather
            .Day3.Caption = "Sat"
            .Day3_hi = sat_hi
            .Day3_lo = sat_lo
            .Day3_Weather = sat_weather
            .Day4.Caption = "Sun"
            .Day4_hi = sun_hi
            .Day4_lo = sun_lo
            .Day4_Weather = sun_weather
            .Day5.Caption = "Mon"
            .Day5_hi = mon_hi
            .Day5_lo = mon_lo
            .Day5_Weather = mon_weather
            .Day6.Caption = "Tue"
            .Day6_hi = tue_hi
            .Day6_lo = tue_lo
            .Day6_Weather = tue_weather
            .Day7.Caption = "Wed"
            .Day7_hi = wed_hi
            .Day7_lo = wed_lo
            .Day7_Weather = wed_weather
        End With
    ElseIf today_weekday = 6 Then
        With Mainfrm
            .Day1.Caption = "Fri"
            .Day1_hi = Today_hi
            .Day1_lo = today_lo
            .Day1_Weather = today_weather
            .Day2.Caption = "Sat"
            .Day2_hi = sat_hi
            .Day2_lo = sat_lo
            .Day2_Weather = sat_weather
            .Day3.Caption = "Sun"
            .Day3_hi = sun_hi
            .Day3_lo = sun_lo
            .Day3_Weather = sun_weather
            .Day4.Caption = "Mon"
            .Day4_hi = mon_hi
            .Day4_lo = mon_lo
            .Day4_Weather = mon_weather
            .Day5.Caption = "Tue"
            .Day5_hi = tue_hi
            .Day5_lo = tue_lo
            .Day5_Weather = tue_weather
            .Day6.Caption = "Wed"
            .Day6_hi = wed_hi
            .Day6_lo = wed_lo
            .Day6_Weather = wed_weather
            .Day7.Caption = "Thu"
            .Day7_hi = thu_hi
            .Day7_lo = thu_lo
            .Day7_Weather = thu_weather
        End With
    ElseIf today_weekday = 7 Then
        With Mainfrm
            .Day1.Caption = "Sat"
            .Day1_hi = Today_hi
            .Day1_lo = today_lo
            .Day1_Weather = today_weather
            .Day2.Caption = "Sun"
            .Day2_hi = sun_hi
            .Day2_lo = sun_lo
            .Day2_Weather = sun_weather
            .Day3.Caption = "Mon"
            .Day3_hi = mon_hi
            .Day3_lo = mon_lo
            .Day3_Weather = mon_weather
            .Day4.Caption = "Tue"
            .Day4_hi = tue_hi
            .Day4_lo = tue_lo
            .Day4_Weather = tue_weather
            .Day5.Caption = "Wed"
            .Day5_hi = wed_hi
            .Day5_lo = wed_lo
            .Day5_Weather = wed_weather
            .Day6.Caption = "Thu"
            .Day6_hi = thu_hi
            .Day6_lo = thu_lo
            .Day6_Weather = thu_weather
            .Day7.Caption = "Fri"
            .Day7_hi = fri_hi
            .Day7_lo = fri_lo
            .Day7_Weather = fri_weather
        End With
    End If
    Call save_data
    Exit Sub
Weather_Error:
    MsgBox "Make Sure You Are Connected To The Internet And You Have Entered A Valid Zip Code", vbOKOnly + vbInformation, "Weather Error"
    StatusBar1.Panels.Item(1).Text = "Updated on " & GetFromINI("Report", "Time", iniPath$)
    Label11.Enabled = True
    Exit Sub
End Sub
Private Sub Form_Load()
    ' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' '
    ' Your Weather v1.2                                                           '
    ' By: Nathan Snyder                                                           '
    ' Programmed in Visual Basic 5.0 Enterprise                                   '
    '                                                                             '
    ' I ripped this program apart from top to bottom and realized there           '
    ' was a lot of code and functions that didn't need to be. I was just          '
    ' going to use this as a starting point, but instead I totally overhauled     '
    ' the entire thing.  I give credit for the idea to the original author,who    '
    ' I would mention, but the person who uploaded the .zip file said he/she      '
    ' forgot who the original author was.  To my knowledge this is the only VB    '
    ' example that gets an extended forecast as well as current information.      '
    '                                                                             '
    ' VOTE FOR ME IF YOU ENJOY THIS <~~~~~~~~~~////                               '
    '                                                                             '
    ' Thanks,                                                                     '
    ' Nathan Snyder                                                               '
    '                                                                             '
    ' If you use this with great frequency, and www.weather.com changes format.   '
    ' Let me know and I will edit this program to keep up with www.weather.com.   '
    '                                                                             '
    ' "Those you dream by day are cognizant of many things which escape those     '
    ' who only dream by night"                                                    '
    ' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' '
    Mainfrm.MouseIcon = ImageList1.ListImages(1).Picture
    Call load_data
End Sub
Private Sub Label11_Click()
    If Ziptxt.Text <> "" And Len(Ziptxt.Text) = 5 Then
        Call LoadWeather
    Else
        MsgBox "Zip Code is invalid", vbOKOnly + vbInformation, "Zipcode Error"
    End If
End Sub
Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' makes for a nifty button effect (down) '
    Line2.BorderColor = vbBlack
    Line7.BorderColor = vbBlack
    Line6.BorderColor = vbWhite
    Line3.BorderColor = vbWhite
    Label11.ForeColor = vbWhite
    Label11.Left = Label11.Left + 15
    Label11.Top = Label11.Top + 15
    Label11.Top = 105
    Label11.Left = 3064
    DoEvents
End Sub
Private Sub Label11_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' makes for a nifty button effect (up) '
    Line6.BorderColor = vbBlack
    Line3.BorderColor = vbBlack
    Line2.BorderColor = vbWhite
    Line7.BorderColor = vbWhite
    Label11.ForeColor = vbBlack
    Label11.Left = Label11.Left - 25
    Label11.Top = Label11.Top - 15
    Label11.Top = 90
    Label11.Left = 3049
    DoEvents
End Sub


