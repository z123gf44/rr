VERSION 5.00
Begin VB.Form setCytimes 
   BackColor       =   &H00FFFF80&
   Caption         =   "������������"
   ClientHeight    =   4995
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4185
   FillColor       =   &H00C0FFC0&
   ForeColor       =   &H00C0FFC0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   4185
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   2100
      TabIndex        =   7
      Top             =   3660
      Width           =   1395
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�ָ�Ĭ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   360
      TabIndex        =   6
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox RecordText 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   435
      Left            =   1860
      TabIndex        =   4
      Text            =   "500"
      Top             =   2220
      Width           =   1275
   End
   Begin VB.TextBox cyText 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   435
      Left            =   1860
      TabIndex        =   1
      Text            =   "500"
      Top             =   600
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "ms"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3300
      TabIndex        =   5
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label recordLabel1 
      Caption         =   "��¼����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "ms"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3300
      TabIndex        =   2
      Top             =   660
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label cyLabel1 
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   660
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "setCytimes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' �ָ�Ĭ������ֵ
Private Sub Command1_Click()

  If setCytimes.cyLabel1.Visible Then
        cyInfoTime = 5
       cyText.Text = cyInfoTime * 100
  Else
      RecordTime = 5
     RecordText = RecordTime * 100
  End If
End Sub

Private Sub Command2_Click()
  If setCytimes.cyLabel1.Visible Then
     cyInfoTime = cyText.Text / 100
  Else
     RecordTime = RecordText / 100
  End If
End Sub

