VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "MOS 2016 Outlook 课件演示系统, Designed by Kukmoon谷月 (QQ/微信：752927706), Ver. 1.0.20200202"
   ClientHeight    =   3120
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11040
   LinkTopic       =   "Form1"
   ScaleHeight     =   208
   ScaleMode       =   0  'User
   ScaleWidth      =   736
   Begin VB.CommandButton Command1 
      Caption         =   "跳过任务"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   6
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdRestartProject 
      Caption         =   "重新开始任务"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8880
      TabIndex        =   4
      Top             =   90
      Width           =   1215
   End
   Begin VB.CommandButton cmdNextTask 
      Caption         =   "下一个任务"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   1
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txtTaskContent 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   720
      Width           =   8415
   End
   Begin VB.Shape cmdRestore 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00808000&
      FillStyle       =   0  'Solid
      Height          =   180
      Index           =   1
      Left            =   10350
      Top             =   180
      Width           =   180
   End
   Begin VB.Shape cmdRestore 
      BorderColor     =   &H00FFFFFF&
      Height          =   180
      Index           =   0
      Left            =   10440
      Top             =   120
      Width           =   180
   End
   Begin VB.Label cmdHelp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "？"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10680
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblMarkFeedback 
      BackColor       =   &H00FFFFFF&
      Caption         =   "☆ 完成考试后提供反馈"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4200
      TabIndex        =   3
      Top             =   2325
      Width           =   2055
   End
   Begin VB.Label lblProject 
      BackStyle       =   0  'Transparent
      Caption         =   "任务 1/35"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.Shape rectTaskButtonBottLine 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   15
      Index           =   0
      Left            =   1350
      Top             =   885
      Width           =   1050
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   2295
      Left            =   180
      Top             =   480
      Width           =   10740
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808000&
      FillStyle       =   0  'Solid
      Height          =   3135
      Left            =   0
      Top             =   0
      Width           =   11055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ap() As cp

Dim TaskIndex As Integer

Const TaskButtonLeft = 90, TaskButtonTop = 35, TaskButtonHeight = 25, TaskButtonWidth = 70
Const TaskButtonBottLineLeft = 90, TaskButtonBottLineTop = 59, TaskButtonBottLineHeight = 1, TaskButtonBottLineWidth = 70

Sub ai()
    Dim i As Integer
    For i = 0 To Controls.Count - 1
        With ap(i)
            .wp = Controls(i).Width / Form1.ScaleWidth
            .hp = Controls(i).Height / Form1.ScaleHeight
            .tp = Controls(i).Top / Form1.ScaleHeight
            .lp = Controls(i).Left / Form1.ScaleWidth
        End With
    Next i
End Sub


Private Sub Form_Load()
    TaskIndex = 0
    TaskNumber = 5
    ReDim Preserve Tasks(TaskNumber)
    ReDim ap(0 To Controls.Count - 1)
    ai
End Sub

Private Sub Form_Resize()
    Dim i As Integer
    For i = 0 To Controls.Count - 1
        Controls(i).Move ap(i).lp * Form1.ScaleWidth, ap(i).tp * Form1.ScaleHeight, ap(i).wp * Form1.ScaleWidth, ap(i).hp * Form1.ScaleHeight
    Next i
End Sub


