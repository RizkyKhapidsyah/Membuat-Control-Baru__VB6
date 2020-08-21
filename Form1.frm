VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Membuat Control Baru"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    On Error GoTo Pesan  'Jika error, lompat ke Pesan
   Selisih = Text1(0).Height * 1.33  'Batas antar
             'control atas dengan bawah
   For i% = 1 To 9     'Tambahkan 9 control lagi
       Load Label1(i%) 'Load label ke memori
       Label1(i%).Caption = "Label" & i% + 1 'Buat nama
                                             'label
       Label1(i%).Top = Label1(i% - 1).Top + Selisih
       'Definisikan batas atas label
       Label1(i%).Visible = True
       'Munculkan label ke form
       Load Text1(i%)  'Load text ke memori
       Text1(i%).Text = "Text" & i% + 1
       'Buat nama textbox
       Text1(i%).Top = Text1(i% - 1).Top + Selisih
       'Definisikan batas atas textbox
       Text1(i%).Visible = True  'Munculkan text ke
       'form
   Next i%   'Maju ke control berikutnya
   Exit Sub  'Jika sudah selesai, keluar dari prosedur
Pesan:  'Jika terjadi kesalahan, karena sudah ada
   MsgBox "Control sudah ada!", vbCritical, "Uuups"
End Sub
