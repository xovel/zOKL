VERSION 5.00
Begin VB.UserControl zOKL 
   BackColor       =   &H80000004&
   ClientHeight    =   3495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8955
   LockControls    =   -1  'True
   ScaleHeight     =   3112.595
   ScaleMode       =   0  'User
   ScaleWidth      =   8382.215
   Begin VB.Frame fraXHP 
      Caption         =   "快捷栏中的消耗品位置"
      Height          =   735
      Left            =   5760
      TabIndex        =   112
      Top             =   480
      Visible         =   0   'False
      Width           =   3015
      Begin VB.CommandButton btnXHPOK 
         Caption         =   "确认"
         Height          =   375
         Left            =   2400
         TabIndex        =   119
         Top             =   240
         Width           =   495
      End
      Begin VB.CheckBox chkXHP 
         Height          =   375
         Index           =   5
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   113
         Top             =   240
         Width           =   375
      End
      Begin VB.CheckBox chkXHP 
         Height          =   375
         Index           =   4
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   114
         Top             =   240
         Width           =   375
      End
      Begin VB.CheckBox chkXHP 
         Height          =   375
         Index           =   3
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   115
         Top             =   240
         Width           =   375
      End
      Begin VB.CheckBox chkXHP 
         Height          =   375
         Index           =   2
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   116
         Top             =   240
         Width           =   375
      End
      Begin VB.CheckBox chkXHP 
         Height          =   375
         Index           =   1
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   117
         Top             =   240
         Width           =   375
      End
      Begin VB.CheckBox chkXHP 
         Height          =   375
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   118
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame fraAdvanced 
      Caption         =   "高级设置"
      Height          =   3495
      Left            =   5640
      TabIndex        =   78
      Top             =   0
      Width           =   3255
      Begin VB.TextBox txtPetSkillHotKey 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   270
         Left            =   2640
         TabIndex        =   121
         Text            =   "V"
         Top             =   480
         Width           =   495
      End
      Begin VB.CheckBox chkPetSkillActive 
         Caption         =   "启动后施放宠物技能 快捷键"
         Height          =   180
         Left            =   120
         TabIndex        =   111
         Top             =   480
         Width           =   2775
      End
      Begin VB.CheckBox chkXHPActive 
         Caption         =   "自动嗑药"
         Height          =   255
         Left            =   2040
         TabIndex        =   120
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox chkTitleActive 
         Caption         =   "称号"
         Height          =   255
         Left            =   1320
         TabIndex        =   80
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox chkActiveDIY 
         Caption         =   "激活自定义功能"
         Height          =   255
         Left            =   1560
         TabIndex        =   79
         Top             =   720
         Width           =   1575
      End
      Begin VB.Frame fraDIY 
         Caption         =   "自定义参数"
         Enabled         =   0   'False
         Height          =   2775
         Left            =   0
         TabIndex        =   82
         Top             =   720
         Width           =   3255
         Begin VB.CheckBox chkActiveFastSet 
            Caption         =   "开启"
            Height          =   255
            Left            =   2280
            TabIndex        =   110
            Top             =   2040
            Width           =   735
         End
         Begin VB.Frame fraFastSet 
            Caption         =   "固定设置"
            Enabled         =   0   'False
            Height          =   615
            Left            =   1200
            TabIndex        =   83
            Top             =   2040
            Width           =   1935
            Begin VB.ComboBox cboFastSet 
               Enabled         =   0   'False
               Height          =   300
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   84
               Top             =   240
               Width           =   1695
            End
         End
         Begin VB.Frame fraFirstConvItem 
            Caption         =   "快捷栏首格"
            Height          =   615
            Left            =   120
            TabIndex        =   88
            Top             =   2040
            Width           =   975
            Begin VB.TextBox txtFirstConvItemX 
               Alignment       =   2  'Center
               Height          =   270
               Left            =   0
               TabIndex        =   89
               Top             =   240
               Width           =   495
            End
            Begin VB.TextBox txtFirstConvItemY 
               Alignment       =   2  'Center
               Height          =   270
               Left            =   480
               TabIndex        =   90
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.Frame fraMouseCenter 
            Caption         =   "换装后鼠标移至的位置"
            Height          =   615
            Left            =   1200
            TabIndex        =   85
            Top             =   1440
            Width           =   1935
            Begin VB.TextBox txtMouseCenterX 
               Alignment       =   2  'Center
               Height          =   270
               Left            =   480
               TabIndex        =   86
               Top             =   240
               Width           =   495
            End
            Begin VB.TextBox txtMouseCenterY 
               Alignment       =   2  'Center
               Height          =   270
               Left            =   960
               TabIndex        =   87
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.Frame fraFirstItem 
            Caption         =   "首格坐标"
            Height          =   615
            Left            =   120
            TabIndex        =   91
            Top             =   1440
            Width           =   975
            Begin VB.TextBox txtFirstItemX 
               Alignment       =   2  'Center
               Height          =   270
               Left            =   0
               TabIndex        =   92
               Top             =   240
               Width           =   495
            End
            Begin VB.TextBox txtFirstItemY 
               Alignment       =   2  'Center
               Height          =   270
               Left            =   480
               TabIndex        =   93
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.Frame fraExEquip 
            Caption         =   "装备选项卡检测范围"
            Height          =   615
            Left            =   1200
            TabIndex        =   94
            Top             =   840
            Width           =   1935
            Begin VB.TextBox txtEquipPos 
               Alignment       =   2  'Center
               Height          =   270
               Index           =   0
               Left            =   0
               TabIndex        =   95
               Top             =   240
               Width           =   495
            End
            Begin VB.TextBox txtEquipPos 
               Alignment       =   2  'Center
               Height          =   270
               Index           =   1
               Left            =   480
               TabIndex        =   96
               Top             =   240
               Width           =   495
            End
            Begin VB.TextBox txtEquipPos 
               Alignment       =   2  'Center
               Height          =   270
               Index           =   2
               Left            =   960
               TabIndex        =   97
               Top             =   240
               Width           =   495
            End
            Begin VB.TextBox txtEquipPos 
               Alignment       =   2  'Center
               Height          =   270
               Index           =   3
               Left            =   1440
               TabIndex        =   98
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.Frame fraEquipPos 
            Caption         =   "装备选项卡"
            Height          =   615
            Left            =   120
            TabIndex        =   99
            Top             =   840
            Width           =   975
            Begin VB.TextBox txtEquipX 
               Alignment       =   2  'Center
               Height          =   270
               Left            =   0
               TabIndex        =   100
               Top             =   240
               Width           =   495
            End
            Begin VB.TextBox txtEquipY 
               Alignment       =   2  'Center
               Height          =   270
               Left            =   480
               TabIndex        =   101
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.Frame fraExBag 
            Caption         =   "物品栏检测范围"
            Height          =   615
            Left            =   1200
            TabIndex        =   102
            Top             =   240
            Width           =   1935
            Begin VB.TextBox txtExBagPos 
               Alignment       =   2  'Center
               Height          =   270
               Index           =   0
               Left            =   0
               TabIndex        =   103
               Top             =   240
               Width           =   495
            End
            Begin VB.TextBox txtExBagPos 
               Alignment       =   2  'Center
               Height          =   270
               Index           =   1
               Left            =   480
               TabIndex        =   104
               Top             =   240
               Width           =   495
            End
            Begin VB.TextBox txtExBagPos 
               Alignment       =   2  'Center
               Height          =   270
               Index           =   2
               Left            =   960
               TabIndex        =   105
               Top             =   240
               Width           =   495
            End
            Begin VB.TextBox txtExBagPos 
               Alignment       =   2  'Center
               Height          =   270
               Index           =   3
               Left            =   1440
               TabIndex        =   106
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.Frame fraEachItem 
            Caption         =   "小格长宽"
            Height          =   615
            Left            =   120
            TabIndex        =   107
            Top             =   240
            Width           =   975
            Begin VB.TextBox txtEachLength 
               Alignment       =   2  'Center
               Height          =   270
               Left            =   0
               TabIndex        =   108
               Top             =   240
               Width           =   495
            End
            Begin VB.TextBox txtEachHeight 
               Alignment       =   2  'Center
               Height          =   270
               Left            =   480
               TabIndex        =   109
               Top             =   240
               Width           =   495
            End
         End
      End
      Begin VB.CheckBox chkConvActive 
         Caption         =   "快捷栏换装"
         Height          =   255
         Left            =   120
         TabIndex        =   81
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox txtBagHotKey 
      Alignment       =   2  'Center
      Height          =   270
      Left            =   2280
      TabIndex        =   77
      Text            =   "I"
      Top             =   0
      Width           =   735
   End
   Begin VB.Frame fraTitle 
      Caption         =   "称号"
      Height          =   615
      Left            =   2520
      TabIndex        =   75
      Top             =   2880
      Visible         =   0   'False
      Width           =   615
      Begin VB.CheckBox chkTitle 
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   180
         Width           =   375
      End
   End
   Begin VB.Frame fraConv 
      Caption         =   "快捷栏"
      Height          =   615
      Left            =   0
      TabIndex        =   68
      Top             =   2880
      Visible         =   0   'False
      Width           =   2415
      Begin VB.CheckBox chkConvPosition 
         Height          =   375
         Index           =   5
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   180
         Width           =   375
      End
      Begin VB.CheckBox chkConvPosition 
         Height          =   375
         Index           =   4
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   180
         Width           =   375
      End
      Begin VB.CheckBox chkConvPosition 
         Height          =   375
         Index           =   3
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   180
         Width           =   375
      End
      Begin VB.CheckBox chkConvPosition 
         Height          =   375
         Index           =   2
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   180
         Width           =   375
      End
      Begin VB.CheckBox chkConvPosition 
         Height          =   375
         Index           =   1
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   180
         Width           =   375
      End
      Begin VB.CheckBox chkConvPosition 
         Height          =   375
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   180
         Width           =   375
      End
   End
   Begin VB.Frame fraEquip 
      Caption         =   "装备栏"
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   3135
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   47
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   2040
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   46
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   2040
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   45
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   2040
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   44
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   2040
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   43
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   2040
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   42
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   2040
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   41
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   2040
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   40
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   2040
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   39
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   1680
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   38
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   1680
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   37
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   1680
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   36
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   1680
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   35
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   1680
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   34
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   1680
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   33
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   1680
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   32
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   1680
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   31
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   1320
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   30
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1320
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   29
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1320
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   28
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   1320
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   27
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   1320
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   26
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1320
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   25
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1320
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   24
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1320
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   23
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   960
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   22
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   960
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   21
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   960
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   20
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   960
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   19
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   960
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   18
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   960
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   17
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   960
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   16
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   960
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   15
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   600
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   14
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   600
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   13
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   600
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   12
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   600
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   11
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   600
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   10
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   600
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   9
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   600
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   8
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   600
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   7
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   6
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   5
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   4
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   3
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   2
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   1
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   375
      End
      Begin VB.CheckBox chkEquipPosition 
         Height          =   375
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   375
      End
      Begin VB.CheckBox chkShowNumber 
         Caption         =   "显示序号"
         Height          =   255
         Left            =   1920
         TabIndex        =   1
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Frame fraOKL 
      Caption         =   "换装设置"
      Height          =   3495
      Left            =   3240
      TabIndex        =   50
      Top             =   0
      Width           =   2295
      Begin VB.Frame fraLeftHand 
         Caption         =   "特殊项"
         Height          =   615
         Left            =   120
         TabIndex        =   61
         Top             =   2760
         Width           =   2055
         Begin VB.CheckBox chkLeftHand 
            Caption         =   "鼠标主次按钮切换"
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame fraExtraDelay 
         Caption         =   "单次额外延迟(毫秒)"
         Height          =   615
         Left            =   120
         TabIndex        =   54
         Top             =   2160
         Width           =   2055
         Begin VB.TextBox txtExtraDelay 
            Alignment       =   2  'Center
            Height          =   270
            Left            =   120
            TabIndex        =   62
            Text            =   "0"
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame fraMouseRCM 
         Caption         =   "换装速度与模式"
         Height          =   615
         Left            =   120
         TabIndex        =   53
         Top             =   1560
         Width           =   2055
         Begin VB.OptionButton optMouseRCM 
            Caption         =   "一般"
            Height          =   255
            Index           =   1
            Left            =   1080
            TabIndex        =   60
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optMouseRCM 
            Caption         =   "超速"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   59
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.Frame fraMode 
         Caption         =   "换装检测方案"
         Height          =   615
         Left            =   120
         TabIndex        =   52
         Top             =   840
         Width           =   2055
         Begin VB.OptionButton optMode 
            Caption         =   "方案二"
            Height          =   300
            Index           =   1
            Left            =   1080
            TabIndex        =   58
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optMode 
            Caption         =   "方案一"
            Height          =   300
            Index           =   0
            Left            =   120
            TabIndex        =   57
            Top             =   240
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.Frame fraAutoBack 
         Caption         =   "自动换回"
         Height          =   615
         Left            =   120
         TabIndex        =   51
         Top             =   240
         Width           =   2055
         Begin VB.TextBox txtAutoBackDelay 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   270
            Left            =   1560
            TabIndex        =   56
            Text            =   "10"
            Top             =   240
            Width           =   375
         End
         Begin VB.CheckBox chkAutoBack 
            Caption         =   "自动换回(秒)"
            Height          =   255
            Left            =   120
            TabIndex        =   55
            Top             =   240
            Width           =   1455
         End
      End
   End
   Begin VB.Frame fraBagEx 
      BackColor       =   &H80000004&
      Caption         =   "背包扩充情况与快捷键"
      Height          =   615
      Left            =   0
      TabIndex        =   64
      Top             =   0
      Width           =   3135
      Begin VB.OptionButton optBagEx 
         Caption         =   "扩充二阶"
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   67
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optBagEx 
         Caption         =   "扩充一阶"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   66
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optBagEx 
         Caption         =   "未扩充"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   65
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "zOKL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'------------------------------------------------------------------------------
'以下的内容，一般情况下您不需要看懂，也不需要进行修改
Type ControlData
    Name As String
    Value As String
End Type

Public FormSizeCX As Long, FormSizeCY As Long

Public Function GetControlData() As Collection
    Set GetControlData = ControlDataCollection
End Function



Private Sub UserControl_Initialize()
    FormSizeCX = Width
    FormSizeCY = Height
    OnInitialize
End Sub
'以上的内容，一般情况下您不需要看懂，也不需要进行修改
'------------------------------------------------------------------------------

Private Sub btnXHPOK_Click()
    Dim i%, n%
    n = 0
    For i = 0 To chkXHP.Count - 1
       If chkXHP(i).Value = Checked Then n = n + 1
    Next
    fraXHP.Visible = False
    If n = 0 Then chkXHPActive.Value = Unchecked
End Sub

Private Sub cboFastSet_Change()
    If chkActiveFastSet.Value = Unchecked Then Exit Sub
    Select Case cboFastSet.ListIndex
        Case Is = 0:
            fraBagEx.Enabled = True
        Case Is = 1:
            zSetParaForTextBoxes
            fraBagEx.Enabled = False
            optBagEx(0).Value = True
        Case Is = 2:
            zSetParaForTextBoxes , 1
            fraBagEx.Enabled = False
            optBagEx(1).Value = True
        Case Is = 3: zSetParaForTextBoxes , 2
            fraBagEx.Enabled = False
            optBagEx(2).Value = True
        Case Is = 4:
            zSetParaForTextBoxes "1024*768"
            fraBagEx.Enabled = False
            optBagEx(0).Value = True
        Case Is = 5:
            zSetParaForTextBoxes "1024*768", 1
            fraBagEx.Enabled = False
            optBagEx(1).Value = True
        Case Is = 6:
            zSetParaForTextBoxes "1024*768", 2
            fraBagEx.Enabled = False
            optBagEx(2).Value = True
        Case Is = 7:
            zSetParaForTextBoxes "1280*960"
            fraBagEx.Enabled = False
            optBagEx(0).Value = True
        Case Is = 8:
            zSetParaForTextBoxes "1280*960", 1
            fraBagEx.Enabled = False
            optBagEx(1).Value = True
        Case Is = 9:
            zSetParaForTextBoxes "1280*960", 2
            fraBagEx.Enabled = False
            optBagEx(2).Value = True
        Case Else
    End Select
End Sub

Private Sub cboFastSet_Click()
    Call cboFastSet_Change
End Sub

Private Sub chkActiveDIY_Click()
    fraDIY.Enabled = chkActiveDIY.Value
    If chkActiveDIY.Value = Checked Then
        fraMode.Enabled = False
        optMode(1).Value = True
    Else
        fraMode.Enabled = True
        chkActiveFastSet.Value = Unchecked
    End If
End Sub

Private Sub chkActiveFastSet_Click()
    If chkActiveFastSet.Value = Checked Then
        fraFastSet.Enabled = True
        cboFastSet.Enabled = True
        fraEachItem.Enabled = False
        fraEachItem.Enabled = False
        fraExBag.Enabled = False
        fraEquipPos.Enabled = False
        fraExEquip.Enabled = False
        fraFirstItem.Enabled = False
        fraMouseCenter.Enabled = False
        fraFirstConvItem.Enabled = False
        Call cboFastSet_Change
    Else
        fraFastSet.Enabled = False
        cboFastSet.Enabled = False
        fraEachItem.Enabled = True
        fraEachItem.Enabled = True
        fraExBag.Enabled = True
        fraEquipPos.Enabled = True
        fraExEquip.Enabled = True
        fraFirstItem.Enabled = True
        fraMouseCenter.Enabled = True
        fraFirstConvItem.Enabled = True
        fraBagEx.Enabled = True
    End If
End Sub

Private Sub chkActivePet_Click()

End Sub

Private Sub chkAutoBack_Click()
    txtAutoBackDelay.Enabled = chkAutoBack.Value
End Sub

Private Sub chkConvActive_Click()
    Dim i As Integer
    fraConv.Visible = chkConvActive.Value
'    fraFirstConvItem.Visible = chkConvActive.Value
    If chkConvActive.Value = Unchecked Then
        For i = 0 To chkConvPosition.Count - 1
            chkConvPosition(i).Value = Unchecked
        Next
        If chkXHPActive.Value = Checked Then
            fraConv.Visible = True
        End If
    End If
End Sub
'
'Private Sub chkEquipPosition_Click(Index As Integer)
'    If chkTitleActive.Value = Checked And _
'        chkTitle.Value = Checked And _
'        chkTitle.Caption = CStr(Index + 1) And _
'        chkEquipPosition(Index).Value = Checked Then
'            MsgBox "该位置已选择了称号。不能继续选择！", vbInformation, "提示"
'            chkEquipPosition(Index).Value = Unchecked
'    End If
'End Sub

Private Sub chkPetSkillActive_Click()
    txtPetSkillHotKey.Enabled = chkPetSkillActive.Value
End Sub

'Private Sub chkLeftHand_Click()
'    If chkLeftHand.Value = Checked Then
'        If MsgBox("请确认是否开启左撇子模式。本模式下鼠标左右键将会对调功能，请注意选择。", vbYesNo + vbQuestion, "提示") = vbNo Then
'            chkLeftHand.Value = Unchecked
'        End If
'    End If
'End Sub

Private Sub chkShowNumber_Click()
    Dim i As Integer
    If chkShowNumber.Value = Checked Then
        For i = 0 To chkEquipPosition.Count - 1
            chkEquipPosition(i).Caption = CStr(i + 1)
        Next
        For i = 0 To chkConvPosition.Count - 1
            chkConvPosition(i).Caption = CStr(i + 1)
        Next
        For i = 0 To chkXHP.Count - 1
            chkXHP(i).Caption = CStr(i + 1)
        Next
    Else
        For i = 0 To chkEquipPosition.Count - 1
            chkEquipPosition(i).Caption = ""
        Next
        For i = 0 To chkConvPosition.Count - 1
            chkConvPosition(i).Caption = ""
        Next
        For i = 0 To chkXHP.Count - 1
            chkXHP(i).Caption = ""
        Next
    End If
End Sub

Private Sub chkTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
        Case Is = 1
            Dim s As String
            If chkTitle.Value = Unchecked Then
                s = InputBox("警告：称号由于不检测冷却，所以不能确保能否正确换上。要换装的称号位置请不要设置在快捷栏中，那样相信手动换上更加直观与快速！" & vbCrLf & vbCrLf & "请输入称号摆放的位置：", "称号位置输入", 0)
                If Trim(s) = Empty Then Exit Sub
                If IsNumeric(s) = False Then
                    MsgBox "请输入一个数值！", vbExclamation, "警告"
                    Exit Sub
                End If
                If CInt(s) > 48 Or CInt(s) < 0 Then
                    MsgBox "不在基本范围之内，请重新设置！", vbExclamation, "警告"
                    Exit Sub
                End If
                If CInt(s) = 0 Then Exit Sub
            End If
            chkTitle.Caption = CStr(s)
            chkTitle.Value = Checked
        Case Is = 2
        
    End Select
End Sub

Private Sub chkTitleActive_Click()
    fraTitle.Visible = chkTitleActive.Value
    chkTitle.Value = Unchecked
    chkTitle.Caption = ""
End Sub

Private Sub chkXHPActive_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
        Case Is = 1
            If chkXHPActive.Value = Unchecked Then
                fraXHP.Visible = True
            Else
                fraXHP.Visible = False
            End If
        Case Is = 2
    End Select
End Sub

Private Sub optBagEx_Click(Index As Integer)
    Dim i As Integer
    
    fraEquip.Visible = True
    fraOKL.Visible = True
    fraAdvanced.Visible = True
    
    Select Case Index
        Case Is = 0
            '''未扩充
            fraEquip.Height = 1616
            For i = 32 To 47
                chkEquipPosition(i).Visible = False
            Next
        Case Is = 1
            fraEquip.Height = 1937
            For i = 32 To 39
                chkEquipPosition(i).Visible = True
            Next
            For i = 40 To 47
                chkEquipPosition(i).Visible = False
            Next
        Case Is = 2
            fraEquip.Height = 2257
            For i = 32 To 47
                chkEquipPosition(i).Visible = True
            Next
        Case Else
            fraEquip.Visible = False
            fraOKL.Visible = False
            fraAdvanced.Visible = False
    End Select
    If chkActiveDIY.Value = Unchecked Then zSetParaForTextBoxes , Index
    If chkActiveFastSet.Value = Unchecked Then
        cboFastSet.ListIndex = Index + 1
    End If

    For i = 0 To chkEquipPosition.Count - 1
        If chkEquipPosition(i).Visible = False Then
            chkEquipPosition(i).Value = Unchecked
        End If
    Next
End Sub

'您可以在OnInitiallize过程里面，写一些关于界面初始化的程序代码
'如果没什么好初始化的，那就不用修改这个函数了
Private Sub OnInitialize()
        
    '''控件安排
    fraEquip.Visible = False
    fraOKL.Visible = False
    fraAdvanced.Visible = False
'    fraFirstConvItem.Visible = False
    cboFastSet.Enabled = False
    fraFastSet.Enabled = False
    fraXHP.Visible = False
    
    '''默认设置800*600窗口模式下的参数
    zSetParaForTextBoxes
    
    zSetToolTipText
    
    With cboFastSet
        .Clear
        .AddItem "未使用快速设置"
        .AddItem "800*600未扩充"
        .AddItem "800*600一阶扩充"
        .AddItem "800*600二阶扩充"
        .AddItem "1024*768未扩充"
        .AddItem "1024*768一阶扩充"
        .AddItem "1024*768二阶扩充"
        .AddItem "1280*960未扩充"
        .AddItem "1280*960一阶扩充"
        .AddItem "1280*960二阶扩充"
    End With
End Sub

'OnSave过程是在用户按下了“保存”按钮的时候调用的
'调用SaveControlData
Public Sub OnSave()
    ClearControlData    '固定位置，不需要进行修改

    Dim i%, n1%, n2%, n3%
    Dim zTemp As Variant
    
    '''勾选式复选框保存
    SaveControlData "zShowNumber", chkShowNumber.Value
    SaveControlData "zAutoBack", chkAutoBack.Value
    SaveControlData "zLeftHand", chkLeftHand.Value
    SaveControlData "zConvActive", chkConvActive.Value
    SaveControlData "zTitleActive", chkTitleActive.Value
    SaveControlData "zActiveDIY", chkActiveDIY.Value
    
    '''固定设置项
    SaveControlData "zActiveFastSet", chkActiveFastSet.Value
    SaveControlData "zFastSetValue", cboFastSet.ListIndex
    
    '''选择项保存
    If optMode(0).Value = True Then
        SaveControlData "zMode", "1"
    Else
        If optMode(1).Value = True Then
            SaveControlData "zMode", "2"
        End If
    End If
    If optMouseRCM(0).Value = True Then
        SaveControlData "zMouseRCM", "1"
    Else
        If optMouseRCM(1).Value = True Then
            SaveControlData "zMouseRCM", "2"
        End If
    End If
'    SaveControlData "zMode1", optMode(0).Value
'    SaveControlData "zMode2", optMode(1).Value
'    SaveControlData "zMouseRCM1", optMouseRCM(0).Value
'    SaveControlData "zMouseRCM2", optMouseRCM(1).Value
'
    '''装备扩充情况保存
    If optBagEx(0).Value = True Then
        SaveControlData "zBagEx", "1"
    Else
        If optBagEx(1).Value = True Then
            SaveControlData "zBagEx", "2"
        Else
            If optBagEx(2).Value = True Then
                SaveControlData "zBagEx", "3"
            Else
                SaveControlData "zBagEx", "0"
            End If
        End If
    End If
    
    '''装备选择信息保存
    n1 = 0
    zTemp = ""
    For i = 0 To chkEquipPosition.Count - 1
        If chkEquipPosition(i).Value = Checked Then
            n1 = n1 + 1
            If n1 > 1 Then
                zTemp = zTemp & "|"
            End If
            zTemp = zTemp & CStr(i + 1)
            SaveControlData "zPos" & CStr(n1), CStr(i + 1)
        End If
    Next
    SaveControlData "zSelectNumber1", CStr(n1)
    SaveControlData "zPosDetail", zTemp
    n2 = 0
    zTemp = ""
    For i = 0 To chkConvPosition.Count - 1
        If chkConvPosition(i).Value = Checked Then
            n2 = n2 + 1
            If n2 > 1 Then
                zTemp = zTemp & "|"
            End If
            zTemp = zTemp & CStr(i + 1)
            SaveControlData "zConvPos" & CStr(n2), CStr(i + 1)
        End If
    Next
    SaveControlData "zSelectNumber2", CStr(n2)
    SaveControlData "zConvPosDetail", zTemp
    If n1 + n2 > 11 Then
        MsgBox "选择的装备数量已经超过11个！请确认装备选择情况并重新保存设置。", vbExclamation, "警告"
    End If
    
    n3 = 0
    zTemp = ""
    For i = 0 To chkXHP.Count - 1
        If chkXHP(i).Value = Checked Then
            n3 = n3 + 1
            If n3 > 1 Then
                zTemp = zTemp & "|"
            End If
            zTemp = zTemp & CStr(i + 1)
            SaveControlData "zXHPPos" & CStr(n3), CStr(i + 1)
        End If
    Next
    SaveControlData "zSelectNumber3", CStr(n3)
    SaveControlData "zXHPPosDetail", zTemp
    
    '''所有文本框内容保存
    SaveControlData "zBagHotKey", txtBagHotKey.Text
    SaveControlData "zAutoBackDelay", txtAutoBackDelay.Text
    SaveControlData "zExtraDelay", txtExtraDelay.Text
    SaveControlData "zEachLength", txtEachLength.Text
    SaveControlData "zEachHeight", txtEachHeight.Text
    SaveControlData "zEquipX", txtEquipX.Text
    SaveControlData "zEquipY", txtEquipY.Text
    SaveControlData "zMouseCenterX", txtMouseCenterX.Text
    SaveControlData "zMouseCenterY", txtMouseCenterY.Text
    SaveControlData "zFirstItemX", txtFirstItemX.Text
    SaveControlData "zFirstItemY", txtFirstItemY.Text
    SaveControlData "zFirstConvItemX", txtFirstConvItemX.Text
    SaveControlData "zFirstConvItemY", txtFirstConvItemY.Text
    SaveControlData "zExBagPos1", txtExBagPos(0).Text
    SaveControlData "zExBagPos2", txtExBagPos(1).Text
    SaveControlData "zExBagPos3", txtExBagPos(2).Text
    SaveControlData "zExBagPos4", txtExBagPos(3).Text
    SaveControlData "zEquipPos1", txtEquipPos(0).Text
    SaveControlData "zEquipPos2", txtEquipPos(1).Text
    SaveControlData "zEquipPos3", txtEquipPos(2).Text
    SaveControlData "zEquipPos4", txtEquipPos(3).Text
    
    '''称号
    SaveControlData "zTitleValue", chkTitle.Value
    SaveControlData "zTitlePos", chkTitle.Caption
    
    '''宠物与消耗品
    SaveControlData "zXHPActive", chkXHPActive.Value
    SaveControlData "zPetSkillActive", chkPetSkillActive.Value
    SaveControlData "zPetSkillHotKey", txtPetSkillHotKey.Text
    
End Sub

'OnLoad过程是在您的界面初始化完成以后被调用的，它的作用是把每个控件的值设为上次保存的值
'调用 LoadControlData

Public Sub OnLoad()
    ''''
    
    Dim i%, k%, zStr$
    Dim zTemp As Variant
    
    txtBagHotKey.Text = LoadControlData("zBagHotKey")
    txtAutoBackDelay.Text = LoadControlData("zAutoBackDelay")
    txtExtraDelay.Text = LoadControlData("zExtraDelay")
    txtEachLength.Text = LoadControlData("zEachLength")
    txtEachHeight.Text = LoadControlData("zEachHeight")
    txtEquipX.Text = LoadControlData("zEquipX")
    txtEquipY.Text = LoadControlData("zEquipY")
    txtMouseCenterX.Text = LoadControlData("zMouseCenterX")
    txtMouseCenterY.Text = LoadControlData("zMouseCenterY")
    txtFirstItemX.Text = LoadControlData("zFirstItemX")
    txtFirstItemY.Text = LoadControlData("zFirstItemY")
    txtFirstConvItemX.Text = LoadControlData("zFirstConvItemX")
    txtFirstConvItemY.Text = LoadControlData("zFirstConvItemY")
    txtExBagPos(0).Text = LoadControlData("zExBagPos1")
    txtExBagPos(1).Text = LoadControlData("zExBagPos2")
    txtExBagPos(2).Text = LoadControlData("zExBagPos3")
    txtExBagPos(3).Text = LoadControlData("zExBagPos4")
    txtEquipPos(0).Text = LoadControlData("zEquipPos1")
    txtEquipPos(1).Text = LoadControlData("zEquipPos2")
    txtEquipPos(2).Text = LoadControlData("zEquipPos3")
    txtEquipPos(3).Text = LoadControlData("zEquipPos4")
    chkShowNumber.Value = LoadControlData("zShowNumber")
    chkAutoBack.Value = LoadControlData("zAutoBack")
    chkLeftHand.Value = LoadControlData("zLeftHand")
    chkConvActive.Value = LoadControlData("zConvActive")
    chkTitleActive.Value = LoadControlData("zTitleActive")
    chkActiveDIY.Value = LoadControlData("zActiveDIY")
    
    zStr = LoadControlData("zMode")
    Select Case zStr
        Case Is = "1"
            optMode(0).Value = True
        Case Is = "2"
            optMode(1).Value = True
    End Select
    zStr = LoadControlData("zMouseRCM")
    Select Case zStr
        Case Is = "1"
            optMouseRCM(0).Value = True
        Case Is = "2"
            optMouseRCM(1).Value = True
    End Select
'    optMode(0).Value = LoadControlData("zMode1")
'    optMode(1).Value = LoadControlData("zMode2")
'    optMouseRCM(0).Value = LoadControlData("zMouseRCM1")
'    optMouseRCM(1).Value = LoadControlData("zMouseRCM2")
    zStr = LoadControlData("zBagEx")
    Select Case zStr
        Case Is = "0"
            fraEquip.Visible = False
            fraOKL.Visible = False
            fraAdvanced.Visible = False
        Case Is = "1"
            optBagEx(0).Value = True
        Case Is = "2"
            optBagEx(1).Value = True
        Case Is = "3"
            optBagEx(2).Value = True
        Case Else
    End Select
    chkTitle.Value = LoadControlData("zTitleValue")
    chkTitle.Caption = LoadControlData("zTitlePos")
    
    '''读取装备选择情况
    zStr = LoadControlData("zSelectNumber1")
    k = CInt(zStr)
    zStr = LoadControlData("zPosDetail")
    zTemp = Split(zStr, "|")
    If k <> (UBound(zTemp) + 1) Then
        MsgBox "读取出现校验错误！", vbCritical, "提示"
    End If
    If k > 0 Then
        For i = 0 To k - 1
            chkEquipPosition(CInt(zTemp(i)) - 1).Value = Checked
        Next
    End If
    
    zStr = LoadControlData("zSelectNumber2")
    k = CInt(zStr)
    zStr = LoadControlData("zConvPosDetail")
    zTemp = Split(zStr, "|")
    If k <> (UBound(zTemp) + 1) Then
        MsgBox "读取出现校验错误！", vbCritical, "提示"
    End If
    If k > 0 Then
        For i = 0 To k - 1
            chkConvPosition(CInt(zTemp(i)) - 1).Value = Checked
        Next
    End If
    
    zStr = LoadControlData("zSelectNumber3")
    k = CInt(zStr)
    zStr = LoadControlData("zXHPPosDetail")
    zTemp = Split(zStr, "|")
    If k <> (UBound(zTemp) + 1) Then
        MsgBox "读取出现校验错误！", vbCritical, "提示"
    End If
    If k > 0 Then
        For i = 0 To k - 1
            chkXHP(CInt(zTemp(i)) - 1).Value = Checked
        Next
    End If
    
    chkActiveFastSet.Value = LoadControlData("zActiveFastSet")
    cboFastSet.ListIndex = LoadControlData("zFastSetValue")
    
    chkXHPActive.Value = LoadControlData("zXHPActive")
    chkPetSkillActive.Value = LoadControlData("zPetSkillActive")
    txtPetSkillHotKey.Text = LoadControlData("zPetSkillHotKey")
    
End Sub

Private Function zSetParaForTextBoxes(Optional ByVal zBasePixel As String = "800*600", Optional ByVal zBagExLevel As Integer = 0)
'    txtEachLength.Text = "30"
'    txtEachHeight.Text = "30"
'    txtExBagPos(0).Text = "486"
'    txtExBagPos(1).Text = "120"
'    txtExBagPos(2).Text = "528"
'    txtExBagPos(3).Text = "134"
'    txtEquipX.Text = "482"
'    txtEquipY.Text = "283"
'    txtEquipPos(0).Text = "474"
'    txtEquipPos(1).Text = "281"
'    txtEquipPos(2).Text = "516"
'    txtEquipPos(3).Text = "297"
'    txtFirstItemX.Text = "487"
'    txtFirstItemY.Text = "307"
'    txtFirstConvItemX.Text = "97"
'    txtFirstConvItemY.Text = "557"
'    txtMouseCenterX.Text = "593"
'    txtMouseCenterY.Text = "194"

    '''zBasePixel标记分辨率。
    '''zBagExLevel标记背包扩充。0表示未扩充，1表示扩充一阶，2表示扩充二阶。
    Select Case zBasePixel
        Case Is = "800*600"
            txtEachLength.Text = "30"
            txtEachHeight.Text = "30"
            txtExBagPos(0).Text = "477"
            txtExBagPos(2).Text = "538"
            txtEquipX.Text = "482"
            txtEquipPos(0).Text = "474"
            txtEquipPos(2).Text = "516"
            txtFirstItemX.Text = "487"
            txtFirstConvItemX.Text = "97"
            txtFirstConvItemY.Text = "558"
            txtMouseCenterX.Text = "593"
            Select Case zBagExLevel
                Case Is = 0
                    txtEquipY.Text = "283"
                    txtExBagPos(1).Text = "120"
                    txtExBagPos(3).Text = "134"
                    txtFirstItemY.Text = "307"
                    txtEquipPos(1).Text = "281"
                    txtEquipPos(3).Text = "297"
                    txtMouseCenterY.Text = "241"
                Case Is = 1
                    txtEquipY.Text = "270"
                    txtExBagPos(1).Text = "103"
                    txtExBagPos(3).Text = "116"
                    txtFirstItemY.Text = "288"
                    txtEquipPos(1).Text = "263"
                    txtEquipPos(3).Text = "280"
                    txtMouseCenterY.Text = "216"
                Case Is = 2
                    txtEquipY.Text = "257"
                    txtExBagPos(1).Text = "87"
                    txtExBagPos(3).Text = "101"
                    txtFirstItemY.Text = "275"
                    txtEquipPos(1).Text = "252"
                    txtEquipPos(3).Text = "267"
                    txtMouseCenterY.Text = "191"
                Case Else
            End Select
        Case Is = "1024*768"
            txtEachLength.Text = "39"
            txtEachHeight.Text = "39"
            txtExBagPos(0).Text = "609"
            txtExBagPos(2).Text = "690"
            txtEquipX.Text = "624"
            txtEquipPos(0).Text = "608"
            txtEquipPos(2).Text = "664"
            txtFirstItemX.Text = "627"
            txtFirstConvItemX.Text = "128"
            txtFirstConvItemY.Text = "713"
            txtMouseCenterX.Text = "730"
            Select Case zBagExLevel
                Case Is = 0
                    txtEquipY.Text = "372"
                    txtExBagPos(1).Text = "151"
                    txtExBagPos(3).Text = "174"
                    txtFirstItemY.Text = "393"
                    txtEquipPos(1).Text = "361"
                    txtEquipPos(3).Text = "384"
                    txtMouseCenterY.Text = "284"
                Case Is = 1
                    txtEquipY.Text = "348"
                    txtExBagPos(1).Text = "127"
                    txtExBagPos(3).Text = "151"
                    txtFirstItemY.Text = "369"
                    txtEquipPos(1).Text = "336"
                    txtEquipPos(3).Text = "361"
                    txtMouseCenterY.Text = "263"
                Case Is = 2
                    txtEquipY.Text = "330"
                    txtExBagPos(1).Text = "112"
                    txtExBagPos(3).Text = "130"
                    txtFirstItemY.Text = "351"
                    txtEquipPos(1).Text = "322"
                    txtEquipPos(3).Text = "340"
                    txtMouseCenterY.Text = "237"
                Case Else
            End Select
        Case Is = "1280*960"
            txtEachLength.Text = "48"
            txtEachHeight.Text = "48"
            txtExBagPos(0).Text = "761"
            txtExBagPos(2).Text = "862"
            txtEquipX.Text = "781"
            txtEquipPos(0).Text = "760"
            txtEquipPos(2).Text = "830"
            txtFirstItemX.Text = "784"
            txtFirstConvItemX.Text = "160"
            txtFirstConvItemY.Text = "893"
            txtMouseCenterX.Text = "921"
            Select Case zBagExLevel
                Case Is = 0
                    txtEquipY.Text = "463"
                    txtExBagPos(1).Text = "189"
                    txtExBagPos(3).Text = "218"
                    txtFirstItemY.Text = "492"
                    txtEquipPos(1).Text = "455"
                    txtEquipPos(3).Text = "480"
                    txtMouseCenterY.Text = "376"
                Case Is = 1
                    txtEquipY.Text = "436"
                    txtExBagPos(1).Text = "159"
                    txtExBagPos(3).Text = "187"
                    txtFirstItemY.Text = "462"
                    txtEquipPos(1).Text = "421"
                    txtEquipPos(3).Text = "449"
                    txtMouseCenterY.Text = "353"
                Case Is = 2
                    txtEquipY.Text = "412"
                    txtExBagPos(1).Text = "139"
                    txtExBagPos(3).Text = "160"
                    txtFirstItemY.Text = "440"
                    txtEquipPos(1).Text = "402"
                    txtEquipPos(3).Text = "433"
                    txtMouseCenterY.Text = "321"
                Case Else
            End Select
        Case Else
    End Select
End Function

'''设置控件说明
Private Function zSetToolTipText()
    cboFastSet.ToolTipText = "使用作者预设的参数。背包二阶扩充的参数由于国服暂未开放，未经过调试，请慎选。"
    chkXHPActive.ToolTipText = "在换装之前使用快捷栏中的消耗品。该功能为附加功能，不对使用结果进行判断。请注意选择！"
    chkPetSkillActive.ToolTipText = "在换装之前施放宠物技能。该功能为附加功能，不确定对应技能是否施放成功。请注意选择！"
    chkTitleActive.ToolTipText = "称号由于自身不检测冷却的问题，不能确保检测换装情况。请慎重选择！"
    chkConvActive.ToolTipText = "作者建议：快捷栏内装备自行手动更换！"
    chkAutoBack.ToolTipText = "经过指定时间后自动将之前换过的装备再换回来。"
    chkLeftHand.ToolTipText = "针对特定用户(比如左撇子)切换了鼠标主次按钮功能而设立的一个参数。"
    chkActiveDIY.ToolTipText = "开启后可以自行设定各项参数坐标值，不过同时会将换装检测方案强制更换为第二种。"
    chkActiveFastSet.ToolTipText = "使用作者提供的预设参数。仅供参考，且开启后各项参数不可更换，请慎重选择！"
    fraEquipPos.ToolTipText = "点击该位置之后将激活装备栏选项卡。"
    fraEachItem.ToolTipText = "每个格子的长宽。zOKL核心参数！实现批量坐标的重要参数。"
    fraFirstItem.ToolTipText = "背包中的第一个格子。zOKL核心参数！自行设置请特别注意。"
    fraFirstConvItem.ToolTipText = "消耗品栏中的第一个格子。zOKL重要参数。在开启消耗品栏换装或自动嗑药后的核心参数。"
    fraMouseCenter.ToolTipText = "换装后鼠标的位置。鼠标移动至该位置后，只要不遮盖工具判断的核心位置即可。"
    optMode(0).ToolTipText = "综合检测方法。对于未冷却完毕的装备也能进行判别并等待其冷却完毕之后进行换装。"
    optMode(1).ToolTipText = "强制向换装。不检测冷却情况。开启高级模式之后必须选择该项！"
    optMouseRCM(0).ToolTipText = "超速模拟方式。飞一般的速度，让你体验超快速换装。"
    optMouseRCM(1).ToolTipText = "正常模拟方式。正常方式换装，网络压力较大的用户不妨选择该项。"
    optBagEx(0).ToolTipText = "背包未进行扩充。此时背包为四排格子。"
    optBagEx(1).ToolTipText = "背包扩充第一阶段。此时背包为五排格子。"
    optBagEx(2).ToolTipText = "背包扩充第二阶段。此时背包为六排格子。"
    txtAutoBackDelay.ToolTipText = "自动换回等待的时间，单位为秒。建议在10-15之间。"
    txtExtraDelay.ToolTipText = "该参数用于设置每一次换装后额外的等待时间。网络延迟问题较高的用户可以适当设置。建议在0-100之间。"
    txtBagHotKey.ToolTipText = "游戏中打开物品栏的快捷键。一般游戏中怎么描述的这里就怎么填写。什么？没有设？开什么玩笑！"
    txtPetSkillHotKey.ToolTipText = "游戏中宠物技能的快捷键。"
    txtEachLength.ToolTipText = "每一个背包小格子的横向长度"
    txtEachHeight.ToolTipText = "每一个背包小格子的纵向长度"
    txtExBagPos(0).ToolTipText = "起始横坐标"
    txtExBagPos(1).ToolTipText = "起始纵坐标"
    txtExBagPos(2).ToolTipText = "终止横坐标"
    txtExBagPos(3).ToolTipText = "终止纵坐标"
    txtEquipPos(0).ToolTipText = "起始横坐标"
    txtEquipPos(1).ToolTipText = "起始纵坐标"
    txtEquipPos(2).ToolTipText = "终止横坐标"
    txtEquipPos(3).ToolTipText = "终止纵坐标"
    txtEquipX.ToolTipText = "横坐标"
    txtEquipY.ToolTipText = "纵坐标"
    txtMouseCenterX.ToolTipText = "横坐标"
    txtMouseCenterY.ToolTipText = "纵坐标"
    txtFirstItemX.ToolTipText = "中央横坐标"
    txtFirstItemY.ToolTipText = "起始纵坐标(放置不可交易物品时最上面金色边框横条所在的纵坐标)"
    txtFirstConvItemX.ToolTipText = "中央横坐标"
    txtFirstConvItemY.ToolTipText = "起始纵坐标(放置不可交易物品时最上面金色边框横条所在的纵坐标)"
End Function

