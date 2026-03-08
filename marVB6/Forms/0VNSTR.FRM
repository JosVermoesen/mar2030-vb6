VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form Venster 
   AutoRedraw      =   -1  'True
   Caption         =   "DashBoard Klanten"
   ClientHeight    =   9570
   ClientLeft      =   645
   ClientTop       =   1260
   ClientWidth     =   13125
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9570
   ScaleWidth      =   13125
   Begin TabDlg.SSTab SSTab1 
      Height          =   3255
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   5741
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Klanten"
      TabPicture(0)   =   "0VNSTR.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblCommissieWJ"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblCommissieBJ"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(5)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(4)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(3)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(2)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Line1(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label3"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "TxtInfo(22)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "TxtInfo(4)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "TxtInfo(8)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "TxtInfo(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "TxtInfo(1)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "TxtInfo(2)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "TxtInfo(3)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "TxtInfo(5)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "TxtInfo(6)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "TxtInfo(7)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "TxtInfo(9)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "TxtInfo(10)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "TxtInfo(11)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "TxtInfo(12)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "TxtInfo(13)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "TxtInfo(14)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "TxtInfo(15)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "TxtInfo(16)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "TxtInfo(17)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "TxtInfo(18)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "TxtInfo(19)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "TxtInfo(20)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "TxtInfo(21)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "CmbBestandsKeuze"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "cmdMijnDocumentenServer"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "cmdMijnDocumentenOnline"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).ControlCount=   35
      TabCaption(1)   =   "Leveranciers"
      TabPicture(1)   =   "0VNSTR.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2(0)"
      Tab(1).ControlCount=   1
      Begin VB.CommandButton cmdMijnDocumentenOnline 
         Caption         =   "Online"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         TabIndex        =   59
         Top             =   2340
         Width           =   975
      End
      Begin VB.CommandButton cmdMijnDocumentenServer 
         Caption         =   "Server"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         TabIndex        =   58
         Top             =   2340
         Width           =   855
      End
      Begin VB.ComboBox CmbBestandsKeuze 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   2400
         Width           =   2295
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   285
         Index           =   21
         Left            =   2580
         TabIndex        =   26
         TabStop         =   0   'False
         Tag             =   "v227"
         Top             =   2100
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   285
         Index           =   20
         Left            =   6960
         TabIndex        =   27
         TabStop         =   0   'False
         Tag             =   "v224"
         Top             =   2280
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   285
         Index           =   19
         Left            =   6960
         TabIndex        =   28
         TabStop         =   0   'False
         Tag             =   "v226"
         Top             =   1920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   285
         Index           =   18
         Left            =   9060
         TabIndex        =   29
         TabStop         =   0   'False
         Tag             =   "A110"
         Top             =   1620
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   285
         Index           =   17
         Left            =   9060
         TabIndex        =   30
         TabStop         =   0   'False
         Tag             =   "vs03"
         Top             =   1320
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   285
         Index           =   16
         Left            =   9060
         TabIndex        =   31
         TabStop         =   0   'False
         Tag             =   "A161"
         Top             =   1020
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   285
         Index           =   15
         Left            =   9060
         TabIndex        =   32
         TabStop         =   0   'False
         Tag             =   "A170"
         Top             =   720
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   285
         Index           =   14
         Left            =   6960
         TabIndex        =   33
         TabStop         =   0   'False
         Tag             =   "vs02"
         Top             =   1620
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   285
         Index           =   13
         Left            =   6960
         TabIndex        =   34
         TabStop         =   0   'False
         Tag             =   "A10A"
         Top             =   1320
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   285
         Index           =   12
         Left            =   6960
         TabIndex        =   35
         TabStop         =   0   'False
         Tag             =   "v201"
         Top             =   1020
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   285
         Index           =   11
         Left            =   6960
         TabIndex        =   36
         TabStop         =   0   'False
         Tag             =   "A120"
         Top             =   780
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   285
         Index           =   10
         Left            =   1620
         TabIndex        =   37
         TabStop         =   0   'False
         Tag             =   "A108"
         Top             =   1800
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   285
         Index           =   9
         Left            =   720
         TabIndex        =   38
         TabStop         =   0   'False
         Tag             =   "A107"
         Top             =   1800
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   285
         Index           =   7
         Left            =   3810
         TabIndex        =   39
         TabStop         =   0   'False
         Tag             =   "A106"
         Top             =   1470
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   285
         Index           =   6
         Left            =   2910
         TabIndex        =   40
         TabStop         =   0   'False
         Tag             =   "A105"
         Top             =   1470
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   255
         Index           =   5
         Left            =   180
         TabIndex        =   41
         TabStop         =   0   'False
         Tag             =   "A104"
         Top             =   1470
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   285
         Index           =   3
         Left            =   180
         TabIndex        =   42
         TabStop         =   0   'False
         Tag             =   "A125"
         Top             =   1140
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   285
         Index           =   2
         Left            =   2880
         TabIndex        =   43
         TabStop         =   0   'False
         Tag             =   "A101"
         Top             =   780
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   285
         Index           =   1
         Left            =   180
         TabIndex        =   44
         TabStop         =   0   'False
         Tag             =   "A100"
         Top             =   780
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   285
         Index           =   0
         Left            =   3000
         TabIndex        =   45
         Top             =   2460
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   285
         Index           =   8
         Left            =   180
         TabIndex        =   46
         TabStop         =   0   'False
         Tag             =   "A109"
         Top             =   1800
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   285
         Index           =   4
         Left            =   2880
         TabIndex        =   54
         TabStop         =   0   'False
         Tag             =   "A127"
         Top             =   1140
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   285
         Index           =   22
         Left            =   6960
         TabIndex        =   57
         TabStop         =   0   'False
         Tag             =   "v002"
         Top             =   2580
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Klanten map Documenten"
         Height          =   495
         Left            =   4800
         TabIndex        =   60
         Top             =   1740
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Nog niets voorzien !"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   0
         Left            =   -73920
         TabIndex        =   55
         Top             =   780
         Width           =   9135
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Adresgegevens"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   53
         Top             =   420
         Width           =   1725
      End
      Begin VB.Line Line1 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   2
         Index           =   0
         X1              =   6840
         X2              =   6840
         Y1              =   780
         Y2              =   2520
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Geboortedata,tel.,fax"
         Height          =   255
         Index           =   2
         Left            =   6960
         TabIndex        =   52
         Top             =   420
         Width           =   1965
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bank, btw, munt, o.ref."
         Height          =   255
         Index           =   3
         Left            =   9060
         TabIndex        =   51
         Top             =   420
         Width           =   2025
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "GSM, E-mail adres"
         Height          =   255
         Index           =   4
         Left            =   9060
         TabIndex        =   50
         Top             =   1980
         Width           =   1965
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Datum laatste contact"
         Height          =   255
         Index           =   5
         Left            =   180
         TabIndex        =   49
         Top             =   2100
         Width           =   1965
      End
      Begin VB.Label lblCommissieBJ 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3360
         TabIndex        =   48
         ToolTipText     =   "Boekjaar"
         Top             =   420
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lblCommissieWJ 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2040
         TabIndex        =   47
         ToolTipText     =   "Werkelijk Jaar"
         Top             =   420
         Visible         =   0   'False
         Width           =   1305
      End
   End
   Begin VB.CommandButton CmdWijzigen 
      Caption         =   "&Wijzigen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   11400
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1450
   End
   Begin VB.CommandButton CmdSluiten 
      Cancel          =   -1  'True
      Caption         =   "Sluiten"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   11400
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   1450
   End
   Begin VB.CommandButton CmdBoekhouding 
      Caption         =   "Boekhoudin&g"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   11400
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1450
   End
   Begin VB.CommandButton Alfa 
      Caption         =   "Klant &zoeken "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   11400
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1450
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   5895
      Left            =   240
      TabIndex        =   5
      Top             =   3480
      Width           =   12765
      _ExtentX        =   22516
      _ExtentY        =   10398
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Port&Folio"
      TabPicture(0)   =   "0VNSTR.frx":0038
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "mebBoekDatum"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "X"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "CmdBeheerKwijtingen"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "CmdProduktStatus"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdKennisBeheer"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "CmdTelebib"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cbBoekDatum"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cbTB2query"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cbIPTool"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdGroeneKaart"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "CmdAs1Log"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "CmdPrinterAfdruk"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "CmdSchadedossiers(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "CmdAgenda"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Retourblokken"
      TabPicture(1)   =   "0VNSTR.frx":0054
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cbPolisHistoriek"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "rtbIndentTB2"
      Tab(1).Control(2)=   "cbInboeken"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "NoIndent"
      TabPicture(2)   =   "0VNSTR.frx":0070
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label2(1)"
      Tab(2).Control(1)=   "rtbNoIndent"
      Tab(2).Control(2)=   "cbIndent"
      Tab(2).ControlCount=   3
      Begin VB.CommandButton cbIndent 
         Caption         =   "Indent"
         Height          =   375
         Left            =   -75000
         TabIndex        =   62
         Top             =   5040
         Width           =   1215
      End
      Begin VB.CommandButton cbInboeken 
         Caption         =   "(her)Inboeking"
         Enabled         =   0   'False
         Height          =   255
         Left            =   -71040
         TabIndex        =   23
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton CmdAgenda 
         Caption         =   "Ag&enda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   11160
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   120
         Width           =   1450
      End
      Begin VB.CommandButton CmdSchadedossiers 
         Caption         =   "&Schades"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   9600
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   120
         Width           =   1450
      End
      Begin VB.CommandButton CmdPrinterAfdruk 
         Caption         =   "Printer Af&druk"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8160
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton CmdAs1Log 
         Caption         =   "&AS/x Logboek"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4920
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   120
         Width           =   1450
      End
      Begin VB.CommandButton cmdGroeneKaart 
         Caption         =   "Gr&oene Kaart"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4920
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton cbIPTool 
         Caption         =   "Informatieplicht"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9600
         TabIndex        =   16
         Top             =   480
         Width           =   1455
      End
      Begin VB.CheckBox cbTB2query 
         Caption         =   "TB2 Controle"
         Height          =   255
         Left            =   6480
         TabIndex        =   15
         Top             =   480
         Width           =   1455
      End
      Begin VB.CheckBox cbBoekDatum 
         Caption         =   "Met Boekdatum"
         Height          =   255
         Left            =   6480
         TabIndex        =   14
         Top             =   120
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton CmdTelebib 
         Caption         =   "&TelebibInfo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3240
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   120
         Width           =   1450
      End
      Begin VB.CommandButton cmdKennisBeheer 
         Caption         =   "KennisBe&heer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3240
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton CmdProduktStatus 
         Caption         =   "Overzicht Waar&borgen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   1680
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   120
         Width           =   1450
      End
      Begin VB.CommandButton CmdBeheerKwijtingen 
         Caption         =   "(AS1) Beheer &Kwijtingen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   120
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   120
         Width           =   1450
      End
      Begin RichTextLib.RichTextBox rtbIndentTB2 
         Height          =   4815
         Left            =   -74880
         TabIndex        =   9
         Top             =   480
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   8493
         _Version        =   393217
         TextRTF         =   $"0VNSTR.frx":008C
      End
      Begin VB.ComboBox cbPolisHistoriek 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74880
         Style           =   2  'Dropdown List
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   120
         Width           =   3735
      End
      Begin MSFlexGridLib.MSFlexGrid X 
         Height          =   4035
         Left            =   120
         TabIndex        =   0
         Top             =   960
         Width           =   12585
         _ExtentX        =   22199
         _ExtentY        =   7117
         _Version        =   393216
         FixedCols       =   0
         BackColor       =   -2147483624
         BackColorBkg    =   16777215
         FocusRect       =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox mebBoekDatum 
         Height          =   255
         Left            =   8400
         TabIndex        =   21
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   7
         Format          =   "MM/YYYY"
         Mask            =   "##/####"
         PromptChar      =   "_"
      End
      Begin RichTextLib.RichTextBox rtbNoIndent 
         Height          =   4695
         Left            =   -75000
         TabIndex        =   61
         Top             =   240
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   8281
         _Version        =   393217
         TextRTF         =   $"0VNSTR.frx":0119
      End
      Begin VB.Label Label2 
         Caption         =   "Nog niets voorzien !"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   1
         Left            =   -71640
         TabIndex        =   56
         Top             =   1800
         Width           =   5055
      End
      Begin VB.Label Label1 
         Caption         =   "Vervaldag"
         Height          =   240
         Index           =   6
         Left            =   -74910
         TabIndex        =   7
         Top             =   3330
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Koers"
         Height          =   240
         Index           =   9
         Left            =   -71040
         TabIndex        =   6
         Top             =   4080
         Width           =   525
      End
   End
   Begin VB.Image imgPapierArm 
      Height          =   615
      Left            =   11880
      Picture         =   "0VNSTR.frx":019D
      Stretch         =   -1  'True
      ToolTipText     =   "Voorkeur Papierarm Beheer"
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "Venster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute _
                            Lib "shell32.dll" _
                            Alias "ShellExecuteA" ( _
                            ByVal hwnd As Long, _
                            ByVal lpOperation As String, _
                            ByVal lpFile As String, _
                            ByVal lpParameters As String, _
                            ByVal lpDirectory As String, _
                            ByVal nShowCmd As Long) _
                            As Long

Dim rsHier As ADODB.Recordset
Dim rsTB2 As ADODB.Recordset
Dim rsTB2Qualifiers As ADODB.Recordset
Dim rsTB2Lists As ADODB.Recordset

Sub GroeneKaart()

Dim DummyText As String
Dim TempoX As Integer
Dim Ktrl99 As Integer

bGet TABLE_VARIOUS, 1, vSet("25" + vBibTekst(TABLE_CONTRACTS, "#A010 #") + Venster!TxtInfo(0).text, 20)
If Ktrl Then
    MsgBox "Detailfiche voertuig opmaken a.u.b.", vbInformation
Else
    RecordToVeld TABLE_VARIOUS
    'hier moet het gebeuren...
    Load frmGroeneKaart
    frmGroeneKaart.mebTekstinfo(0).text = TxtInfo(1).text & " " & TxtInfo(2).text
    frmGroeneKaart.mebTekstinfo(1).text = TxtInfo(5).text & " " & TxtInfo(6).text & " " & TxtInfo(7).text
    frmGroeneKaart.mebTekstinfo(2).text = TxtInfo(9).text & " " & TxtInfo(10).text
    frmGroeneKaart.mebTekstinfo(3).text = vBibTekst(TABLE_CONTRACTS, "#A000 #")
    frmGroeneKaart.mebTekstinfo(4).text = vBibTekst(TABLE_VARIOUS, "#5020 #")
    frmGroeneKaart.mebTekstinfo(5).text = Trim$(vBibTekst(TABLE_VARIOUS, "#500A #")) & " " & Trim$(vBibTekst(TABLE_VARIOUS, "#500B #"))
    For COUNT_TO = 1 To frmGroeneKaart.cmbMaatschappij.ListCount
        If InStr(frmGroeneKaart.cmbMaatschappij.List(COUNT_TO), vBibTekst(TABLE_CONTRACTS, "#A010 #")) Then
            frmGroeneKaart.cmbMaatschappij.ListIndex = COUNT_TO
            Exit For
        End If
    Next
    Err = 0
    On Local Error Resume Next
    frmGroeneKaart.Show
    If Err Then MsgBox Error
End If

End Sub

Sub PolisLogBoek()
Dim DummyText As String
Dim TempoX As Integer
Dim Ktrl99 As Integer

bGet TABLE_VARIOUS, 1, vSet("25" + vBibTekst(TABLE_CONTRACTS, "#A010 #") + Venster!TxtInfo(0).text, 20)
If Ktrl Then
    Ktrl99 = Ktrl
    TLB_RECORD(TABLE_VARIOUS) = ""
    vBib TABLE_VARIOUS, Venster!TxtInfo(0).text, "A000"
    vBib TABLE_VARIOUS, Venster!TxtInfo(18).text, "A110"
    bGet TABLE_CONTRACTS, 0, Venster!TxtInfo(0).text
    If Ktrl Then
        MsgBox "logicastop"
    Else
        RecordToVeld TABLE_CONTRACTS
        bGet TABLE_SUPPLIERS, 0, "CO" & vBibTekst(TABLE_CONTRACTS, "#A010 #")
        If Ktrl Then
        Else
            RecordToVeld TABLE_SUPPLIERS
            vBib TABLE_VARIOUS, vBibTekst(TABLE_SUPPLIERS, "#A400 #"), "A400"
        End If
        vBib TABLE_VARIOUS, vBibTekst(TABLE_CONTRACTS, "#A010 #"), "A010"
        If TeleBibClick(4000 + Val(Left(vBibTekst(TABLE_CONTRACTS, "#vs96 #"), 1))) Then
            vBib TABLE_VARIOUS, "K" + Venster!TxtInfo(18).text, "v004"
            vBib TABLE_VARIOUS, "25" + vBibTekst(TABLE_VARIOUS, "#A010 #") + vBibTekst(TABLE_VARIOUS, "#A000 #"), "v005"
            bInsert TABLE_VARIOUS, 1
        End If
    End If
Else
    RecordToVeld TABLE_VARIOUS
    bGet TABLE_SUPPLIERS, 0, "CO" & vBibTekst(TABLE_CONTRACTS, "#A010 #")
    If Ktrl Then
    Else
        RecordToVeld TABLE_SUPPLIERS
        vBib TABLE_VARIOUS, vBibTekst(TABLE_SUPPLIERS, "#A400 #"), "A400"
    End If
    vBib TABLE_VARIOUS, vBibTekst(TABLE_CONTRACTS, "#A010 #"), "A010"
    
    If TeleBibClick(4000 + Val(Left(vBibTekst(TABLE_CONTRACTS, "#vs96 #"), 1))) Then
        vBib TABLE_VARIOUS, "K" + Venster!TxtInfo(18).text, "v004"
        vBib TABLE_VARIOUS, "25" + vBibTekst(TABLE_VARIOUS, "#A010 #") + vBibTekst(TABLE_VARIOUS, "#A000 #"), "v005"
        bUpdate TABLE_VARIOUS, 1
    End If
End If
Exit Sub

End Sub

Private Sub RefreshKlanten(Klant As String)

    Dim Teller As Integer
    
    bGet TABLE_CUSTOMERS, 0, vSet(Klant, 12)
    If Ktrl Then
        bFirst TABLE_CUSTOMERS, 0
        RecordToVeld TABLE_CUSTOMERS
    Else
        RecordToVeld TABLE_CUSTOMERS
    End If
    Venster.Caption = "Fiche : " + RTrim$(vBibTekst(TABLE_CUSTOMERS, "#A100 #")) + " " + RTrim$(vBibTekst(TABLE_CUSTOMERS, "#A125 #"))

    For Teller = 1 To 22
        TxtInfo(Teller).text = vBibTekst(TABLE_CUSTOMERS, "#" + TxtInfo(Teller).Tag + " #")
    Next
    If rsMAR(TABLE_CUSTOMERS)("g101") = "0" Then
        Me.imgPapierArm.Visible = False
    Else
        Me.imgPapierArm.Visible = True
    End If
    If Trim(rsMAR(TABLE_CUSTOMERS)("v247")) <> "" Then
        Me.cmdMijnDocumentenServer.Enabled = True
    Else
        Me.cmdMijnDocumentenServer.Enabled = False
    End If
    If Trim(rsMAR(TABLE_CUSTOMERS)("v254")) <> "" Then
        Me.cmdMijnDocumentenOnline.Enabled = True
    Else
        Me.cmdMijnDocumentenOnline.Enabled = False
    End If
    

End Sub

Sub SchadeLogBoek()
Dim DummyText As String
Dim TempoX As Integer

BeginSchadeLog:
Unload Xlog
Load Xlog
Xlog.WijzigenLijn.Visible = False
Xlog.Caption = "Schadedossiers voor klant : " + TxtInfo(18).text
Xlog.X.Cols = 4
Xlog.X.Rows = 2
Xlog.X.Col = 0
Xlog.X.Row = 0
Xlog.X.text = "Schadenummer"
Xlog.X.Col = 1
Xlog.X.text = "Datum schade"
Xlog.X.Col = 2
Xlog.X.text = "Datum afsluiting"
Xlog.X.Col = 3
Xlog.X.text = "Tegenpartij"

bGetOrGreater TABLE_VARIOUS, 0, vSet("K" + vBibTekst(TABLE_CUSTOMERS, "#A110 #"), 13)
If Ktrl Then
    Beep
    Unload Xlog
    Exit Sub
Else
    RecordToVeld TABLE_VARIOUS
    GoSub SchadeLijnErBij
End If
    
Screen.MousePointer = vbHourglass
Do
    bNext TABLE_VARIOUS
    If Ktrl Or vSet(KEY_BUF(TABLE_VARIOUS), 13) <> vSet("K" + vBibTekst(TABLE_CUSTOMERS, "#A110 #"), 13) Then
        Exit Do
    Else
        RecordToVeld TABLE_VARIOUS
        GoSub SchadeLijnErBij
    End If
Loop
Screen.MousePointer = vbNormal
Xlog.X.Row = 1
Xlog.X.Col = 0

Xlog.X.ColWidth(0) = 1995
Xlog.X.ColWidth(1) = 1110
Xlog.X.ColWidth(2) = 1110
Xlog.X.ColWidth(3) = 3255

KiesSchadeUitLijst:
Xlog.SSTab1.TabVisible(1) = False
Xlog.Show 1
If RTrim$(XLogKey) <> "" Then
    If XLogKey = "Nieuw" Then
        Msg = "Nieuw schadedossier invoegen voor polis " + Venster!TxtInfo(0).text + ".  Bent U zeker ?"
        KtrlBox = MsgBox(Msg, vbYesNo + vbDefaultButton2)
        If KtrlBox = vbYes Then
            TLB_RECORD(TABLE_VARIOUS) = ""
            vBib TABLE_VARIOUS, Venster!TxtInfo(0).text, "A000"
            bGet TABLE_CONTRACTS, 0, Venster!TxtInfo(0).text
            If Ktrl Then
                MsgBox "logikastop"
            Else
                RecordToVeld TABLE_CONTRACTS
                vBib TABLE_VARIOUS, vBibTekst(TABLE_CONTRACTS, "#A010 #"), "A010"
            End If
            
            If TeleBibClick(3001) Then
                vBib TABLE_VARIOUS, "K" + Venster!TxtInfo(18).text, "v004"
                vBib TABLE_VARIOUS, "24" + vBibTekst(TABLE_VARIOUS, "#A010 #") + vBibTekst(TABLE_VARIOUS, "#C000 #"), "v005"
                bInsert TABLE_VARIOUS, 1
            End If
        Else
            GoTo BeginSchadeLog
        End If
    Else
        Screen.MousePointer = vbNormal
        bGet TABLE_VARIOUS, 1, Mid(XLogKey, 1, InStr(XLogKey, Chr(13)) - 1)
        If Ktrl Then
            MsgBox "logikafout"
        Else
            RecordToVeld TABLE_VARIOUS
            If TeleBibClick(3001) Then
                Msg = "Bestaand schadedossier wijzigen.  Bent U zeker ?"
                KtrlBox = MsgBox(Msg, vbYesNo + vbDefaultButton2)
                If KtrlBox = vbYes Then
                    bUpdate TABLE_VARIOUS, 1
                End If
            End If
        End If
    End If
End If
Unload Xlog
Exit Sub

SchadeLijnErBij:
If Left(vBibTekst(TABLE_VARIOUS, "#v005 #"), 2) <> "24" Then
Else
    DummyText = vBibTekst(TABLE_VARIOUS, "#v005 #") + vbTab
    DummyText = DummyText + vBibTekst(TABLE_VARIOUS, "#C200 #") + vbTab
    DummyText = DummyText + vBibTekst(TABLE_VARIOUS, "#C029 #") + vbTab
    DummyText = DummyText + vBibTekst(TABLE_VARIOUS, "#C100 #")
    Xlog.X.AddItem DummyText, Xlog.X.Rows - 1
End If
Return

End Sub


Sub XSchoon()

X.Rows = 2
With X
    .Col = 0
    .Row = 1
    .text = ""
    
    .Col = 1
    .text = ""
    
    .Col = 2
    .text = ""

    .Col = 3
    .text = ""

    .Col = 4
    .text = ""
End With
lblCommissieWJ.Caption = Format(TotaalCommissie(vBibTekst(TABLE_CUSTOMERS, "#A110 #"), Left(BOOKYEAR_FROMTO, 4) + "0101", Left(BOOKYEAR_FROMTO, 4) + "1231"), "#,##0.00")
lblCommissieBJ.Caption = Format(TotaalCommissie(vBibTekst(TABLE_CUSTOMERS, "#A110 #"), Left(BOOKYEAR_FROMTO, 8), Right(BOOKYEAR_FROMTO, 8)), "#,##0.00")

End Sub

Private Sub Alfa_Click()

SharedFl = TABLE_CUSTOMERS

aIndex = 1
GridText = TxtInfo(1).text
SqlSearch.Show 1
If Ktrl Then
Else
    RefreshKlanten Trim(rsMAR(TABLE_CUSTOMERS)("A110"))
    If Mim.SnelHelp.Panels(1).text = "Verzekeringsbemiddelaar" Then
        ContractVensterVullen Trim(rsMAR(TABLE_CUSTOMERS)("A110"))
    End If
End If

End Sub


Private Sub cbBoekDatum_Click()

    Me.mebBoekDatum.text = Format(Now, "mm/yyyy")
    If Me.cbBoekDatum.Value = vbChecked Then
        Me.mebBoekDatum.Enabled = True
    Else
        Me.mebBoekDatum.Enabled = False
    End If
    
End Sub

Private Sub cbInboeken_Click()

    Dim MAPIString As String
        
    rsTB2.AbsolutePosition = Me.cbPolisHistoriek.ListIndex + 1
    MAPIString = rsTB2("memoTB2").Value
    EdiFactCONTANT MAPIString

End Sub

Private Sub cbIndent_Click()

    Me.rtbNoIndent.text = tb2Indent(Me.rtbNoIndent.text)

End Sub

Private Sub cbIPTool_Click()
    
    Static MyLongID As Long
    
    Dim flHier As Integer
    Dim strLocKlantGegevens As String
    Dim strIPTool As String
    
    strIPTool = GetSetting("IP-Tool", "General", "ProductCode")
    If strIPTool = "" Then
        MsgBox "IP-Tool toepassing niet geïnstalleerd op dit toestel", vbExclamation
    Else
        strLocKlantGegevens = GetSetting("IP-Tool", "General", "Klantgegevens")
        If strLocKlantGegevens = "" Then
            strLocKlantGegevens = App.path & "\ipTool.txt"
            SaveSetting "IP-Tool", "General", "Klantgegevens", strLocKlantGegevens
        End If
        
        Dim path As String * 260
        Dim ret As Integer

        'Application Data lokatie ?
        ret = SHGetFolderPath(0, CSIDL_PROGRAM_FILES, 0, 0, path)
        Dim ProgrammaBestanden As String
        ProgrammaBestanden = Left(path, InStr(path, Chr(0)) - 1)
        
        'IPTool poging programmalocatie !
        strIPTool = ProgrammaBestanden & "\IPTool" & "\iptool.exe"
    
        flHier = FreeFile
        Open strLocKlantGegevens For Output As flHier
            Print #flHier, Trim$(Trim$(Me.TxtInfo(1)) & " " & Trim$(Me.TxtInfo(3))); vbCrLf;
            Print #flHier, Trim$(Trim$(Me.TxtInfo(5)) & " " & Trim$(Me.TxtInfo(6)) & " " & Trim$(Me.TxtInfo(7))); vbCrLf;
            Print #flHier, Trim$(Me.TxtInfo(8)) & " " & Trim$(Me.TxtInfo(9)); " " & Trim$(Me.TxtInfo(10));
        Close flHier

        If MyLongID <> 0 Then
            DoEvents
            On Error Resume Next
            AppActivate MyLongID
            'voorlopig niets meer
            'SendKeys "{ENTER}", True
            'SendKeys "%{F4}", True
            If Err = 0 Then Exit Sub
        End If
        On Error Resume Next
        Err = 0
        MyLongID = Shell(strIPTool, 1)
        If Err Then
            MsgBox Error & vbCr & vbCr & strIPTool
        Else
            DoEvents
            AppActivate MyLongID
        End If
    End If
    
End Sub


Private Sub cbPolisHistoriek_Click()

    Dim MAPIString As String
        
    rsTB2.AbsolutePosition = Me.cbPolisHistoriek.ListIndex + 1
    MAPIString = rsTB2("memoTB2").Value
    Me.cbInboeken.Enabled = False
    
    Select Case Right(cbPolisHistoriek.text, 2)
        Case "03"
            Me.rtbIndentTB2.text = SnippetXEH(MAPIString, Trim(Me.TxtInfo(0).text), True)
            Me.rtbNoIndent.text = MAPIString
        
        Case "01"
            Me.rtbIndentTB2.text = tb2Indent(MAPIString)
            Me.rtbNoIndent.text = MAPIString
            'controleren of er een (nog te verwerken) boekingsstuk is
            Me.cbInboeken.Enabled = False
            If InStr(MAPIString, "XEH+03") Then 'ER IS EEN CONTANTE!
                Me.cbInboeken.Enabled = True
            End If
            Me.rtbIndentTB2.SetFocus
        
        Case Else
            MsgBox "nog niets voorzien voor: " & Right(cbPolisHistoriek.text, 2)
            Me.rtbIndentTB2.text = tb2Indent(MAPIString)
            Me.rtbNoIndent.text = MAPIString
    End Select
   
End Sub


Private Sub cbTB2query_Click()

    If Me.cbTB2query.Value = vbChecked Then
        Me.cbBoekDatum.Visible = True
        Me.mebBoekDatum.Visible = True
        Me.mebBoekDatum.Enabled = False
        Me.cbBoekDatum.Value = vbUnchecked
    Else
        Me.cbBoekDatum.Visible = False
        Me.mebBoekDatum.Visible = False
        Me.mebBoekDatum.Enabled = False
        Me.cbBoekDatum.Value = vbUnchecked
    End If
    
End Sub



Private Sub CmdAs1Log_Click()

ASLogBoek

End Sub

Private Sub CmdBeheerKwijtingen_Click()

On Error Resume Next
BeheerKwijtingen.Show 1

End Sub

Private Sub CmdBoekhouding_Click()

bGet TABLE_CUSTOMERS, 0, vSet(TxtInfo(18).text, 12)
If Ktrl Then
Else
    RecordToVeld TABLE_CUSTOMERS
    'BalansKontrole TABLE_CUSTOMERS
    'Stop
    BalansKontroleWithRecordSet TABLE_CUSTOMERS
End If

End Sub



Private Sub cmdGroeneKaart_Click()

If Trim(Venster!TxtInfo(0).text) = "" Then
    Exit Sub
Else
    bGet TABLE_CONTRACTS, 0, Venster!TxtInfo(0).text
    If Ktrl Then
    Else
        RecordToVeld TABLE_CONTRACTS
    End If
    If Left(vBibTekst(TABLE_CONTRACTS, "#v223 #"), 1) <> "5" Then
        MsgBox "Polis: " & Trim$(vBibTekst(TABLE_CONTRACTS, "#A000 #")) & " betreft géén B.A. motorvoertuig.  Controleer en/of verbeter eventueel", vbInformation
    Else
        GroeneKaart
    End If
End If

End Sub

Private Sub CmdKennisBeheer_Click()

bGet TABLE_CONTRACTS, 0, Venster!TxtInfo(0).text
If Ktrl Then
Else
    RecordToVeld TABLE_CONTRACTS
    FrmKennisBeheer.Show 1
End If

End Sub

Private Sub cmdMijnDocumentenOnline_Click()

    'Dim frmB As frmBrowser
    Dim r As Long

    'Set frmB = New frmBrowser
    If Err Then MsgBox Error: Exit Sub
    Dim strAdress As String
    strAdress = Trim(rsMAR(TABLE_CUSTOMERS)("v254"))
    r = ShellExecute(0, "open", strAdress, 0, 0, 1)
        
End Sub

Private Sub cmdMijnDocumentenServer_Click()

    On Local Error Resume Next
    
    Dim strPath As String
    Dim r As String
        
    strPath = LaadTekst("dnnInstellingen", "Archief") & "\" & Trim(rsMAR(TABLE_CUSTOMERS)("v247"))
    If Err Then
        MsgBox Err.Description & vbCrLf & vbCrLf & LaadTekst("dnnInstellingen", "Archief") & "\" & Trim(rsMAR(TABLE_CUSTOMERS)("v257"))
    Else
        r = ShellExecute(0, "open", strPath, 0, 0, 1)
    End If
    
End Sub

Private Sub CmdPrinterAfdruk_Click()

    'PrintAnywhere Venster, Printer

ProbeerTeDrukken:
Err = 0
On Error Resume Next
Venster.PrintForm
If Err Then
    MsgBox Error
    Set Printer = Printers(0)
    GoTo ProbeerTeDrukken
End If

End Sub

Private Sub CmdProduktStatus_Click()
Dim fKtrl As Integer
Dim OptieTxt As String

bGet TABLE_CONTRACTS, 0, TxtInfo(0).text
If Ktrl Then
    MsgBox "stop"
Else
    RecordToVeld TABLE_CONTRACTS
    bGet TABLE_SUPPLIERS, 0, "CO" + vBibTekst(TABLE_CONTRACTS, "#A010 #")
    If Ktrl Then
        MsgBox "stop bij zoeken maatschappij"
        Exit Sub
    Else
        RecordToVeld TABLE_SUPPLIERS
    End If
    bGet TABLE_VARIOUS, 1, "25" + vBibTekst(TABLE_CONTRACTS, "#A010 #") + TxtInfo(0).text
    fKtrl = Ktrl
    If fKtrl Then
        MsgBox "Polis nog niet in bestand.  Wordt nu aangemaakt..."
        TLB_RECORD(TABLE_VARIOUS) = ""
        vBib TABLE_VARIOUS, "K" + vBibTekst(TABLE_CONTRACTS, "#A110 #"), "v004"
        vBib TABLE_VARIOUS, vBibTekst(TABLE_CONTRACTS, "#A110 #"), "A110"
        vBib TABLE_VARIOUS, vBibTekst(TABLE_CONTRACTS, "#A010 #"), "A010"
        vBib TABLE_VARIOUS, vBibTekst(TABLE_CONTRACTS, "#A000 #"), "A000"
        vBib TABLE_VARIOUS, vSet("25" + vSet(vBibTekst(TABLE_CONTRACTS, "#A010 #"), 4) + vBibTekst(TABLE_CONTRACTS, "#A000 #"), 20), "v005"
    Else
        RecordToVeld TABLE_VARIOUS
    End If
End If

If TeleBibClick(2000 + Val(vBibTekst(TABLE_CONTRACTS, "#vs96 #"))) Then
    If fKtrl Then
        bInsert TABLE_VARIOUS, 1
    Else
        bUpdate TABLE_VARIOUS, 1
    End If
    If Ktrl Then MsgBox "Stopkode " + Str$(Ktrl)
End If

Msg = "Gegevens afdrukken.  Bent U zeker ?"
KtrlBox = MsgBox(Msg, vbYesNo + vbDefaultButton2)
If KtrlBox = vbYes Then
    Printer.FontName = "Courier New"
    Printer.FontSize = 12
    Printer.FontBold = True
        Printer.Print vbCrLf; vbCrLf;
        Printer.Print UCase$(Mid(Mim.Caption, InStr(Mim.Caption, "[")));
    Printer.FontBold = False
        Printer.Print vbCrLf; vbCrLf;
        Printer.Print "marINTEGRAAL NT : ";
    Printer.FontBold = True
        Printer.Print "Gekontroleerde toestand risico's via polis :" + TxtInfo(0).text;
    Printer.FontBold = False
        Printer.Print vbCrLf;
        Printer.Print Left(FULL_LINE, 80); vbCrLf;
    Printer.FontBold = False
        Printer.Print Tab(1); "Maatschappij  : ";
    Printer.FontBold = True
        Printer.Print vBibTekst(TABLE_SUPPLIERS, "#A010 #"); " "; vBibTekst(TABLE_SUPPLIERS, "#A100 #");
    Printer.FontBold = False
        Printer.Print vbCrLf;
        Printer.Print Left(FULL_LINE, 80); vbCrLf;
        Printer.Print Tab(1); "Klanten       : ";
    Printer.FontBold = True
        Printer.Print RTrim$(TxtInfo(1).text + " " + TxtInfo(2).text); " "; vbCrLf;
        Printer.Print Tab(1); "                ";
        Printer.Print ; RTrim$(TxtInfo(3).text + " " + TxtInfo(4).text);
    Printer.FontBold = False
        Printer.Print vbCrLf;
        Printer.Print Tab(1); "Adres         : ";
    Printer.FontBold = True
        Printer.Print RTrim$(TxtInfo(5).text + " " + TxtInfo(6).text + " " + TxtInfo(7).text); vbCrLf;
        Printer.Print Tab(1); "                "; RTrim$(TxtInfo(8).text + " " + TxtInfo(9).text + " " + TxtInfo(10).text);
    Printer.FontBold = False
        Printer.Print vbCrLf;
        Printer.Print Tab(1); "Telefoon      : ";
    Printer.FontBold = True
        Printer.Print TxtInfo(13).text;
    Printer.FontBold = False
        Printer.Print vbCrLf;
        Printer.Print Tab(1); "Fax           : ";
    Printer.FontBold = True
        Printer.Print TxtInfo(14).text;
    Printer.FontBold = False
        Printer.Print vbCrLf; vbCrLf;
        Printer.Print Tab(1); "Korte omschrijving voor polis bovenvermeld :"; vbCrLf;
        Printer.Print Left(FULL_LINE, 80); vbCrLf;
    Printer.FontBold = True
        Printer.Print vBibTekst(TABLE_CONTRACTS, "#vs99 #"); " "; vBibTekst(TABLE_CONTRACTS, "#vs98 #");
    Printer.FontBold = False
        Printer.Print vbCrLf; vbCrLf;
        Printer.Print Tab(1); "'GEDEKTE?' waarborgen via maatschappijen onder kontrole van ons kantoor :"; vbCrLf;
        Printer.Print Left(FULL_LINE, 80); vbCrLf;

    Dim T As Integer
    T = 0
    Do While TELEBIB_CODE(T) <> Space$(10)
        If RTrim$(vBibTekst(TABLE_VARIOUS, "#" + Mid(TELEBIB_CODE(T), 5, 5) + "#")) = "" Then
        Else
            OptieTxt = vBibTekst(TABLE_VARIOUS, "#" + Mid(TELEBIB_CODE(T), 5, 5) + "#")
            Printer.FontBold = False
            Printer.Print Tab(1); TELEBIB_TEXT(T);
            Printer.FontBold = True
            If Mid(TELEBIB_CODE(T), 1, 3) <> "   " Then
                Printer.Print Tab(40); fmarBoxText(Mid(TELEBIB_CODE(T), 1, 3), "2", OptieTxt);
            Else
                If Len(OptieTxt) > 35 Then
                    Print vbCrLf;
                    Printer.Print Tab(40); OptieTxt;
                Else
                    Printer.Print Tab(40); OptieTxt;
                End If
            End If
            Printer.Print vbCrLf;
        End If
        T = T + 1
    Loop
    
    Printer.FontBold = False
        Printer.Print vbCrLf; vbCrLf; "Datum en uur vandaag             : ";
    Printer.FontBold = True
        Printer.Print Format(Now, "dd/mm/yy") + " " + Format(Now, "hh:mm:ss");
    Printer.FontBold = False
        Printer.Print vbCrLf;
        Printer.Print vbCrLf; vbCrLf; "Dit dokument is slechts rechtsgeldig voor zover het hieronder voorzien is"; vbCrLf;
        Printer.Print "van een GEDAGTEKENDE handtekening door HEDWIG of JOS !"
    Printer.NewPage
    Printer.EndDoc
End If

End Sub

Private Sub CmdSchadedossiers_Click(Index As Integer)

SchadeLogBoek
'KTRLSchadeDossiers.Show 1

End Sub

Private Sub cmdSluiten_Click()

Unload Me

End Sub

Private Sub CmdTelebib_Click()

    PolisLogBoek

End Sub


Private Sub cmdWijzigen_Click()
Dim TempoFL As Integer

TempoFL = Val(Left(CmbBestandsKeuze.text, 2))
Select Case TempoFL
    Case TABLE_CONTRACTS
        bGet TABLE_CONTRACTS, 0, vSet(TxtInfo(0).text, 12)
    Case TABLE_CUSTOMERS
        bGet TABLE_CUSTOMERS, 0, vSet(TxtInfo(18).text, 12)
End Select
If Ktrl = 0 Then
    RecordToVeld TempoFL
    INSERT_FLAG(TempoFL) = 0
Else
    If TempoFL = TABLE_CUSTOMERS Then MsgBox "Nieuwe klanten moeten via klantenfiches aangemaakt worden": Exit Sub
    INSERT_FLAG(TempoFL) = 1
    TLB_RECORD(TempoFL) = ""
    vBib TempoFL, (TxtInfo(0).text), "A000"   'Kontraktnummer
End If
           
If TeleBibClick(TempoFL) = False Then
Else
    If INSERT_FLAG(TempoFL) = 0 Then
        'Msg = "Gegevens bestaande '" + BstNaam(TempoFL) + "'-fiche wijzigen.  Bent U zeker ?"
        'KtrlBox = MsgBox(Msg, 292)
        'If KtrlBox = 6 Then
            bUpdate TempoFL, 0
        'Else
        '    Exit Sub
        'End If
    Else
        bInsert TempoFL, 0
    End If
    cmdWijzigen.Default = True
    RefreshKlanten vBibTekst(TABLE_CONTRACTS, "#A110 #")
    ContractVensterVullen vBibTekst(TABLE_CONTRACTS, "#A110 #")
End If

End Sub



Private Sub CmdWijzigenInlichtingen_Click()

    'WijzigenInlichtingen.Show 1 : eventueel later aanpassen aan AS2 WEBVERSIE !!!
    MsgBox "vanaf nu gebruik voor opvragen telebib2 ontvangen retourblokken"


End Sub


Private Sub Form_Load()

LaadInstellingen Me
Me.SSTab1.TabVisible(1) = False

    Set rsTB2 = New ADODB.Recordset
    rsTB2.CursorLocation = adUseClient
    
'rsMAR(TABLE_INVOICES).Close
Ktrl = ktrlAllerleiContract
If Ktrl = False Then
    AutoUnloadBedrijf
    Exit Sub
End If

X.Cols = 6
With X
    .Col = 0
    .Row = 0
    .text = "Jaarvdag"
    .ColWidth(0) = 750
    .ColAlignment(0) = flexAlignLeftTop
    
    .Col = 1
    .text = "Betaalbaar, status"
    .ColWidth(1) = 3000
    .ColAlignment(1) = flexAlignLeftTop
        
    .Col = 2
    .text = "Contract"
    .ColWidth(2) = 1365
    .ColAlignment(2) = flexAlignLeftTop
        
    .Col = 3
    .text = "Maatschappij"
    .ColWidth(3) = 2145
    .ColAlignment(3) = flexAlignLeftTop
        
    .Col = 4
    .text = "Soort Verzekering en [Detail]"
    .ColWidth(4) = 4485
    .ColAlignment(4) = flexAlignLeftTop
    
    .Col = 5
    .text = "%Fee"
    .ColWidth(5) = 495
    .ColAlignment(5) = flexAlignRightTop

End With

CmbBestandsKeuze.AddItem "01: Klantenfiche"
SSTab2.TabVisible(1) = False 'retourblokken

If Mim.SnelHelp.Panels(1).text = "Verzekeringsbemiddelaar" Then
    TxtInfo(0).Visible = True
    CmbBestandsKeuze.AddItem "05: Polisfiche"
    CmbBestandsKeuze.ListIndex = 1

    If vBibTekst(TABLE_CONTRACTS, "#A000 #") <> "" Then
        bGet TABLE_CONTRACTS, 0, vBibTekst(TABLE_CONTRACTS, "#A000 #")
        If Ktrl Then
            RefreshKlanten ("180087")
            ContractVensterVullen "180087"
        Else
            RecordToVeld TABLE_CONTRACTS
            RefreshKlanten vBibTekst(TABLE_CONTRACTS, "#A110 #")
            ContractVensterVullen vBibTekst(TABLE_CONTRACTS, "#A110 #")
        End If
    Else
        RefreshKlanten ("180087")
        ContractVensterVullen "180087"
    End If
Else
    TxtInfo(0).Visible = False
    SSTab2.TabVisible(0) = False
    SSTab2.TabVisible(1) = False
    CmbBestandsKeuze.ListIndex = 0
    RefreshKlanten ("")
End If

End Sub

Sub ASLogBoek()
Dim DummyText As String

Unload Xlog
Load Xlog
Xlog.WijzigenLijn.Visible = False
Xlog.Caption = "Logboek voor polis : " + TxtInfo(0).text
Xlog.X.Cols = 3
Xlog.X.Rows = 2
Xlog.X.Col = 0
Xlog.X.Row = 0
Xlog.X.text = "Assurnet IDA"
Xlog.X.Col = 1
Xlog.X.text = "Datum en uur"
Xlog.X.Col = 2
Xlog.X.text = "Omschrijving"

bGetOrGreater TABLE_VARIOUS, 1, vSet("22" + TxtInfo(0).text, 20)
If Ktrl Then
    Beep
    Unload Xlog
    Exit Sub
Else
    RecordToVeld TABLE_VARIOUS
End If
If vSet(KEY_BUF(TABLE_VARIOUS), 20) <> vSet("22" + TxtInfo(0).text, 20) Then
    Unload Xlog
    MsgBox "Geen Assurnetberichten te vinden voor polis : " + TxtInfo(0).text
    Exit Sub
Else
    Screen.MousePointer = vbHourglass
    GoSub LijnErBij
    Do
        bNext TABLE_VARIOUS
        If Ktrl Or vSet(KEY_BUF(TABLE_VARIOUS), 20) <> vSet("22" + TxtInfo(0).text, 20) Then
            Exit Do
        Else
            RecordToVeld TABLE_VARIOUS
            GoSub LijnErBij
        End If
    Loop
End If
Screen.MousePointer = vbNormal
Xlog.X.Row = 1
Xlog.X.Col = 0

Xlog.X.ColWidth(0) = 1560
Xlog.X.ColWidth(1) = 1425
Xlog.X.ColWidth(2) = 3360

KiesUitLijst:
Xlog.SSTab1.TabVisible(1) = False
Xlog.Show 1
If RTrim$(XLogKey) <> "" Then
    Screen.MousePointer = vbHourglass
    LogBoekKontrole.Show 1
    GoTo KiesUitLijst
End If
Unload Xlog
Exit Sub

LijnErBij:
DummyText = Mid(vBibTekst(TABLE_VARIOUS, "#v219 #"), 1, 3) + "/" + Mid(vBibTekst(TABLE_VARIOUS, "#v219 #"), 4, 5) + "/" + Mid(vBibTekst(TABLE_VARIOUS, "#v219 #"), 9, 4) + vbTab
DummyText = DummyText + Mid(vBibTekst(TABLE_VARIOUS, "#v128 #"), 7, 2) + "/" + Mid(vBibTekst(TABLE_VARIOUS, "#v128 #"), 5, 2) + "/" + Mid(vBibTekst(TABLE_VARIOUS, "#v128 #"), 1, 4) + " " + Mid(vBibTekst(TABLE_VARIOUS, "#v128 #"), 9, 2) + "u." + vbTab
DummyText = DummyText + vBibTekst(TABLE_VARIOUS, "#v105 #")
Xlog.X.AddItem DummyText, Xlog.X.Rows - 1
Return

End Sub


Private Sub ContractVensterVullen(Klant As String)
Dim MaatschappijNaam As String
Dim TakString As String
Dim Teller As Integer

XSchoon
bGet TABLE_CONTRACTS, 1, vSet(Klant, 12)
If Ktrl Then
    Beep
    GoTo KlantVerversen
Else
    RecordToVeld TABLE_CONTRACTS
End If
Screen.MousePointer = vbHourglass
If vSet(KEY_BUF(TABLE_CONTRACTS), 12) <> vSet(Klant, 12) Then
    Beep
Else
    Do
        GridText = Left(DATE_TEXT(vBibTekst(TABLE_CONTRACTS, "#AW_2 #")), 5) + vbTab
        GridText = GridText + RTrim$(fmarBoxText("914", "2", vBibTekst(TABLE_CONTRACTS, "#A325 #"))) + ", " + RTrim$(fmarBoxText("001", "2", vBibTekst(TABLE_CONTRACTS, "#vs97 #"))) + vbTab
        GridText = GridText + vBibTekst(TABLE_CONTRACTS, "#A000 #") + vbTab
        bGet TABLE_SUPPLIERS, 0, vSet("CO" + vBibTekst(TABLE_CONTRACTS, "#A010 #"), 12)
        If Ktrl Then
            MaatschappijNaam = vBibTekst(TABLE_CONTRACTS, "#A010 #") + ": Niet aanwezig"
        Else
            RecordToVeld TABLE_SUPPLIERS
            MaatschappijNaam = vBibTekst(TABLE_CONTRACTS, "#A010 #") + ": " + vBibTekst(TABLE_SUPPLIERS, "#A100 #")
        End If
        GridText = GridText + MaatschappijNaam + vbTab
        GridText = GridText + RTrim$(vBibTekst(TABLE_CONTRACTS, "#vs99 #")) + " [" + RTrim$(vBibTekst(TABLE_CONTRACTS, "#vs98 #")) + "]" + vbTab
        GridText = GridText + vBibTekst(TABLE_CONTRACTS, "#e071 #")
        
        X.AddItem GridText, X.Rows - 1
        If RTrim$(vBibTekst(TABLE_CONTRACTS, "#v223 #")) = "" Then
            TakString = Left(vBibTekst(TABLE_CONTRACTS, "#vs96 #"), 1)
            If TakString = "6" Then
                TakString = "611"
            ElseIf TakString = "7" Then
                TakString = "710"
            ElseIf TakString = "8" Then
                TakString = "810"
            ElseIf TakString = "9" Then
                If vBibTekst(TABLE_CONTRACTS, "#vs96 #") = "96" Then
                    TakString = "961"
                ElseIf vBibTekst(TABLE_CONTRACTS, "#vs96 #") = "93" Then
                    TakString = "930"
                ElseIf vBibTekst(TABLE_CONTRACTS, "#vs96 #") = "94" Then
                    TakString = "940"
                Else
                    TakString = "961"
                End If
            Else
                TakString = TakString + "00"
            End If
            vBib TABLE_CONTRACTS, TakString, "v223"
            bUpdate TABLE_CONTRACTS, 1
        End If
        bNext TABLE_CONTRACTS
        If Ktrl Or vSet(KEY_BUF(TABLE_CONTRACTS), 12) <> vSet(Klant, 12) Then
            Exit Do
        Else
            RecordToVeld TABLE_CONTRACTS
        End If
    Loop
End If
Screen.MousePointer = vbNormal

KlantVerversen:
On Local Error Resume Next

EindeOpzoeking:
If X.Rows = 2 Then
    TxtInfo(0).text = ""
Else
    X.Row = 1
    X.Col = 2
    'X.SelStartCol = 0
    'X.SelEndCol = 4
    'X.SelStartRow = X.Row
    'X.SelEndRow = X.Row
    TxtInfo(0).text = X.text
    On Local Error Resume Next
End If

End Sub

Private Sub Form_Resize()


On Local Error Resume Next
SSTab2.Width = Venster.Width - 655
SSTab2.Height = Venster.Height - 4210

X.Width = SSTab2.Width - 250
X.Height = SSTab2.Height - 1400

rtbIndentTB2.Width = SSTab2.Width - 300
rtbIndentTB2.Height = SSTab2.Height - 500 - 400
rtbNoIndent.Width = SSTab2.Width - 300
rtbNoIndent.Height = SSTab2.Height - 500 - 400
'Me.cbInboeken.Left = Me.SSTab2.Width - 2085

End Sub



Private Sub Form_Unload(Cancel As Integer)

Dim X As Boolean

On Local Error Resume Next
bGet TABLE_CONTRACTS, 0, Venster!TxtInfo(0).text
If Ktrl Then
Else
    RecordToVeld TABLE_CONTRACTS
End If
X = InstellingenBewaard(Me)

End Sub










Private Sub TxtInfo_Change(Index As Integer)

    Dim strTB2 As String
    Dim iDTM008 As Integer
    Dim strDatum008 As String
            
If Index = 0 Then
    If Me.cbTB2query.Value = vbChecked And Trim(Me.TxtInfo(0).text) <> "" Then
        Me.cbPolisHistoriek.Clear
        Me.rtbIndentTB2.text = ""
        Me.rtbNoIndent.text = ""
            
        bGet TABLE_CONTRACTS, 0, Trim(TxtInfo(0).text)
        If Ktrl Then
            TLB_RECORD(TABLE_CONTRACTS) = ""
        Else
            RecordToVeld TABLE_CONTRACTS
        End If
        
        If rsTB2.State = adStateOpen Then
            rsTB2.Close
        End If
        Msg = "SELECT * FROM TB2 WHERE Polis = '" & TxtInfo(0).text & "'"
        'Msg = "SELECT * FROM TB2 WHERE memoTB2 Like '%RFF+001:" & Trim(Me.TxtInfo(0).Text) & "%'"
        'If Me.cbBoekDatum.Value = vbChecked Then
        '    Msg = Msg & " AND memoTB2 Like '%DTM+005:" & Me.mebBoekDatum.ClipText & ":005" & "%'"
        'End If
        
        Screen.MousePointer = vbHourglass
        rsTB2.Open Msg, adntDB, adOpenDynamic, adLockOptimistic, adCmdText
        Screen.MousePointer = vbNormal
        If rsTB2.RecordCount = 0 Then
            SSTab2.TabCaption(1) = "&TB2= (0)"
            SSTab2.TabVisible(1) = False
            Me.rtbIndentTB2.text = ""
            Me.rtbNoIndent.text = ""
        Else
            SSTab2.TabCaption(1) = "&TB2= (" & Trim$(Str$(rsTB2.RecordCount)) & ")"
            SSTab2.TabVisible(1) = True
            Do While Not rsTB2.EOF
                strTB2 = rsTB2("memoTB2").Value
                Me.cbPolisHistoriek.AddItem rsTB2("DatumAanvang").Value & " - " & rsTB2("DocType").Value
                rsTB2.MoveNext
            Loop
        End If
        'rsAllerlei.Close
    End If
End If

If RTrim$(TxtInfo(0).text) = "" Then
    TxtInfo(0).TabStop = True
Else
    TxtInfo(0).TabStop = False
End If

End Sub

Private Sub TxtInfo_GotFocus(Index As Integer)

If Index = 0 Then cmdWijzigen.Default = True

End Sub

Private Sub TxtInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

If Index <> 0 Then Exit Sub
Select Case KeyCode
    Case 38 'naar omhoog
        X.Col = 2
        If X.Row < 2 Then
        Else
            X.Row = X.Row - 1
        End If
        TxtInfo(0).text = X.text
    Case 40 'naar beneden
        X.Col = 2
        If X.Row = X.Rows - 1 Then
        Else
            X.Row = X.Row + 1
        End If
        TxtInfo(0).text = X.text
End Select

End Sub

Private Sub TxtInfo_LostFocus(Index As Integer)

If Index = 0 Then Alfa.Default = True

End Sub

Private Sub X_Click()

On Local Error Resume Next
With X
    .Col = 2
    TxtInfo(0).text = .text
End With

End Sub

Private Sub X_DblClick()

cmdWijzigen_Click

End Sub


Private Sub X_GotFocus()

cmdWijzigen.Default = True

End Sub


Private Sub X_KeyUp(KeyCode As Integer, Shift As Integer)

X_Click

End Sub


Private Sub X_RowColChange()

If X.Rows <> 2 Then
    X_Click
End If

End Sub


Private Function TotaalCommissie(KlantKode As String, PeriodeVan As String, PeriodeTot As String) As Currency

Dim datPrimairRS  As ADODB.Recordset
Set datPrimairRS = New ADODB.Recordset

TotaalCommissie = 0
On Error Resume Next
Screen.MousePointer = vbHourglass
datPrimairRS.CursorLocation = adUseClient
Msg = "SELECT Sum(val(B014)) FROM Klanten, dokumenten WHERE NOT IsNull(B014) AND 'K'+A110 = v034 AND v035 >= '" + PeriodeVan + "' AND v035 <= '" + PeriodeTot + "' AND A110 = '" + KlantKode + "'"
datPrimairRS.Open Msg, adntDB, adOpenStatic, adLockReadOnly
If datPrimairRS.RecordCount Then
    TotaalCommissie = datPrimairRS.fields(0) / 40.3399
End If
datPrimairRS.Close
Msg = "SELECT Sum(val(B094)) FROM Klanten, dokumenten WHERE NOT IsNull(B094) AND 'K'+A110 = v034 AND v035 >= '" + PeriodeVan + "' AND v035 <= '" + PeriodeTot + "' AND A110 = '" + KlantKode + "'"
datPrimairRS.Open Msg, adntDB, adOpenStatic, adLockReadOnly
If datPrimairRS.RecordCount Then
    TotaalCommissie = (TotaalCommissie - datPrimairRS.fields(0)) / 40.3399
End If
datPrimairRS.Close
Set datPrimairRS = Nothing
Screen.MousePointer = vbNormal

End Function


Function ktrlAllerleiContract() As Boolean

    Dim ktrlString As String
    Dim MAPIString As String
        
    Dim StringXGH As String
    Dim StringXEH As String
    Dim StringXET As String
    Dim StringXGT As String
    Dim MijNummer As String
    Dim PolisNummer As String

    Dim Xarray() As String
    Dim XGHarray() As String
    Dim XEHarray() As String
    Dim XEHcode As String
        
    Dim posSTR As Integer
    
        
Err = 0
On Error Resume Next

'poging 03/2007

ktrlAllerleiContract = True
ktrlString = rsMAR(TABLE_VARIOUS)("A000").Name

If Err = 3265 Then
    rsMAR(TABLE_VARIOUS).Close
    Msg = "ALTER TABLE Allerlei ADD COLUMN A000 TEXT(12);"
    Err = 0
    adntDB.Execute Msg
    If Err Then
        MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
    Else
        MsgBox Msg, vbInformation, "Met succes"
    End If
    MsgBox "Nieuwe velden werden toegevoegd aan tabel 'Allerlei'. Gelieve het bedrijf opnieuw te openen a.u.b."
    
    Set rsHier = New ADODB.Recordset
    rsHier.CursorLocation = adUseClient
        Msg = "SELECT * FROM Allerlei WHERE v005 Like '30%'"
        rsHier.Open Msg, adntDB, adOpenDynamic, adLockOptimistic, adCmdText
    
    Set rsTB2 = New ADODB.Recordset
    rsTB2.CursorLocation = adUseClient
        Msg = "SELECT * FROM TB2"
        rsTB2.Open Msg, adntDB, adOpenDynamic, adLockOptimistic, adCmdText
    
    While Not rsHier.EOF
        TLB_RECORD(TABLE_VARIOUS) = rsHier.fields("MEMO")
        MAPIString = vBibTekst(TABLE_VARIOUS, "#v132 #")
        'tb2Indent (MAPIString)
        'If InStr$(MAPIString, "XEH+01") Then
            Xarray = Split(MAPIString, "'")
            XGHarray = Split(Xarray(0) + "+", "+")
            rsTB2.AddNew
            rsTB2("Mij").Value = Format(XGHarray(2), "0000")
            rsTB2("memoTB2").Value = MAPIString
            XEHarray = Split(Xarray(1) + "+", "+")
            XEHcode = Format(XEHarray(1), "00")
            rsTB2("DocType").Value = XEHcode
            
            'op zoek naar een polisnummer
            posSTR = InStr(MAPIString, "RFF+001:")
            If posSTR = 0 Then
                'er is geen polisnummer !?!
                If XEHcode = "07" Then
                    'logisch mogelijk
                Else
                    'Niks want het zal oude AS1 zijn!
                End If
            Else
                posSTR = InStr(MAPIString, "RFF+001:")
                Msg = Mid(MAPIString, posSTR + 8)
                Msg = Left(Msg, InStr(Msg, "'") - 1)
                PolisNummer = Msg
                If Left(PolisNummer, Len(XEHarray(4))) = XEHarray(4) Then
                    'ok polisnummer 100%
                    rsTB2("Polis").Value = XEHarray(4)
                Else
                    'mogelijk is het een schadenummer (eerst xeharray(1) checken !!
                    If XEHarray(1) = "02" Then 'schade !
                        'inderdaad schadegeval..
                        rsTB2("Polis").Value = PolisNummer
                    Else
                        If rsTB2("Mij").Value = "0497" Then
                            rsTB2("Polis").Value = PolisNummer
                        Else
                            rsTB2("Polis").Value = PolisNummer
                        End If
                    End If
                End If
                'nu nog de datum op basis van doctype
                If XEHcode = "01" Then 'contract!
                    posSTR = InStr(MAPIString, "DTM+008:")
                    If posSTR = 0 Then
                        'mogelijk nieuwe zaak dus ktrl op DTM+041:
                        posSTR = InStr(MAPIString, "DTM+041:")
                        If posSTR = 0 Then
                            'verder zoeken!
                            posSTR = InStr(MAPIString, "DTM+060:")
                            If posSTR = 0 Then
                                'verder zoeken!
                                Stop
                            End If
                           
                        End If
                    End If
                    Msg = Mid(MAPIString, posSTR + 8)
                    Msg = Left(Msg, 8)
                    rsTB2("DatumAanvang").Value = CVDate(Mid(Msg, 1, 2) & "/" & Mid(Msg, 3, 2) & "/" & Mid(Msg, 5, 4))
                ElseIf XEHcode = "02" Then 'contract
                    posSTR = InStr(MAPIString, "DTM+048:")
                    If posSTR = 0 Then
                        Stop
                    Else
                        Msg = Mid(MAPIString, posSTR + 8)
                        Msg = Left(Msg, 8)
                    End If
                    rsTB2("DatumAanvang").Value = CVDate(Mid(Msg, 1, 2) & "/" & Mid(Msg, 3, 2) & "/" & Mid(Msg, 5, 4))
                ElseIf XEHcode = "07" Then 'rekeninguittreksel?
                    rsTB2("Polis").Value = "" 'zeker geen polisnummer dan!
                    posSTR = InStr(MAPIString, "DTM+069:")
                    If posSTR = 0 Then
                        Stop
                    Else
                        Msg = Mid(MAPIString, posSTR + 8)
                        Msg = Left(Msg, 8)
                    End If
                    rsTB2("DatumAanvang").Value = CVDate(Mid(Msg, 1, 2) & "/" & Mid(Msg, 3, 2) & "/" & Mid(Msg, 5, 4))
                ElseIf XEHcode = "03" Then 'hernieuwingen ?
                    rsTB2("Polis").Value = "" 'zeker geen polisnummer dan!
                    posSTR = InStr(MAPIString, "DTM+005:")
                    If posSTR = 0 Then
                        'Stop
                    Else
                        Msg = Mid(MAPIString, posSTR + 8)
                        Msg = Left(Msg, 6)
                    End If
                    rsTB2("DatumAanvang").Value = CVDate("01" & "/" & Mid(Msg, 1, 2) & "/" & Mid(Msg, 3, 4))
                Else
                    Stop
                End If
                SnelHelpPrint Msg, False
                'laatste controle, testen of de maatschappij en de polis (nog) aanwezig is..
                If XEHcode = "01" Or XEHcode = "02" Then
                    bGet TABLE_CONTRACTS, 0, rsTB2("Polis").Value
                    If Ktrl Then
                        'Stop
                    End If
                End If
                bGet TABLE_SUPPLIERS, 0, "CO" & rsTB2("Mij").Value
                If Ktrl Then
                    If IsNull(rsTB2("Mij").Value) Then
                        posSTR = InStr(MAPIString, "PTY+006+")
                        If posSTR = 0 Then
                            Stop
                        Else
                            Msg = Mid(MAPIString, posSTR + 8)
                            Msg = Left(Msg, 4)
                            rsTB2("Mij").Value = Msg
                            If Msg <> "0039" Then
                                Stop
                            End If
                        End If
                    Else
                        Stop
                    End If
                End If
                
                rsTB2.Update
   
            End If
        rsHier.MoveNext
    Wend
    rsHier.Close
    
End If

End Function


