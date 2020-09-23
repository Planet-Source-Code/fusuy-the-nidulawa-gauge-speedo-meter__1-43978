VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Performance Transparant Gauge Meter"
   ClientHeight    =   2655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4560
   Icon            =   "transparent.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   177
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   304
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Khijau 
      Height          =   345
      Left            =   6780
      Picture         =   "transparent.frx":0CCA
      ScaleHeight     =   285
      ScaleWidth      =   225
      TabIndex        =   9
      Top             =   1665
      Width           =   285
   End
   Begin VB.PictureBox Kkuning 
      Height          =   345
      Left            =   6465
      Picture         =   "transparent.frx":0F06
      ScaleHeight     =   285
      ScaleWidth      =   225
      TabIndex        =   8
      Top             =   1665
      Width           =   285
   End
   Begin VB.PictureBox Kmerah 
      Height          =   345
      Left            =   6165
      Picture         =   "transparent.frx":1142
      ScaleHeight     =   285
      ScaleWidth      =   225
      TabIndex        =   7
      Top             =   1665
      Width           =   285
   End
   Begin VB.PictureBox KHitam 
      Height          =   345
      Left            =   5805
      Picture         =   "transparent.frx":137E
      ScaleHeight     =   285
      ScaleWidth      =   225
      TabIndex        =   6
      Top             =   1695
      Width           =   285
   End
   Begin VB.PictureBox skala2 
      Height          =   2655
      Left            =   3585
      Picture         =   "transparent.frx":14D3
      ScaleHeight     =   2595
      ScaleWidth      =   7095
      TabIndex        =   5
      Top             =   5445
      Visible         =   0   'False
      Width           =   7155
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      Height          =   2625
      Left            =   150
      Picture         =   "transparent.frx":37AED
      ScaleHeight     =   171
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   171
      TabIndex        =   4
      Top             =   2775
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Proses"
      Height          =   390
      Left            =   3255
      TabIndex        =   3
      Top             =   2235
      Width           =   1275
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   2700
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "0"
      Top             =   1665
      Width           =   1830
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2580
      Left            =   60
      ScaleHeight     =   172
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   172
      TabIndex        =   1
      Top             =   60
      Width           =   2580
   End
   Begin VB.PictureBox Skala1 
      Height          =   2655
      Left            =   3585
      Picture         =   "transparent.frx":37E2B
      ScaleHeight     =   2595
      ScaleWidth      =   7095
      TabIndex        =   0
      Top             =   2745
      Visible         =   0   'False
      Width           =   7155
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Val(Text1.Text) >= 0 And Val(Text1.Text) <= 200 Then
    x = Int(Text1.Text / 5)
    y = Val(Text1.Text / 5)
    For i = 0 To x
        With Picture4
            .Cls
            .PaintPicture Picture5.Picture, 0, 0, , , , , , , vbSrcPaint
            If y > x Then
                .PaintPicture skala2.Picture, 0, 0, 172, 172, i * 172, 0, 172, 172, vbSrcAnd
            Else
                .PaintPicture Skala1.Picture, 0, 0, 172, 172, i * 172, 0, 172, 172, vbSrcAnd
            End If
            
            If Val(Text1.Text) >= 90 Then
                'Hijau Meriah euy
                .PaintPicture KHitam.Picture, 55, 130, , , , , , , vbSrcAnd
                .PaintPicture KHitam.Picture, 77, 130, , , , , , , vbSrcAnd
                .PaintPicture Khijau.Picture, 99, 130, , , , , , , vbSrcAnd
            ElseIf Val(Text1.Text) >= 80 And Val(Text1.Text) < 90 Then
                'Kuning Meriah euy
                .PaintPicture KHitam.Picture, 55, 130, , , , , , , vbSrcAnd
                .PaintPicture Kkuning.Picture, 77, 130, , , , , , , vbSrcAnd
                .PaintPicture KHitam.Picture, 99, 130, , , , , , , vbSrcAnd
            Else
                'Merah Meriah euy
                .PaintPicture Kmerah.Picture, 55, 130, , , , , , , vbSrcAnd
                .PaintPicture KHitam.Picture, 77, 130, , , , , , , vbSrcAnd
                .PaintPicture KHitam.Picture, 99, 130, , , , , , , vbSrcAnd
            End If
        End With
        DoEvents
    Next
    

    
    
End If
End Sub

Private Sub Form_Load()
    With Picture4
        .Cls
        .PaintPicture Picture5.Picture, 0, 0, , , , , , , vbSrcPaint
        .PaintPicture Skala1.Picture, 0, 0, 172, 172, 0 * 172, 0, 172, 172, vbSrcAnd
        'Light Panel
        .PaintPicture KHitam.Picture, 55, 130, , , , , , , vbSrcAnd
        .PaintPicture KHitam.Picture, 77, 130, , , , , , , vbSrcAnd
        .PaintPicture KHitam.Picture, 99, 130, , , , , , , vbSrcAnd
    End With
End Sub


