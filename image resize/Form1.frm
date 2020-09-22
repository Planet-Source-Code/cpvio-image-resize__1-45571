VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   8715
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      Height          =   5655
      Left            =   2880
      ScaleHeight     =   5595
      ScaleWidth      =   5475
      TabIndex        =   2
      Top             =   120
      Width           =   5535
      Begin VB.Image Image1 
         Height          =   5655
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   5535
      End
   End
   Begin VB.DirListBox Dir1 
      Height          =   2565
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
   Begin VB.FileListBox File1 
      Height          =   2625
      Left            =   240
      Pattern         =   "*.bmp;*.jpg;*.gif;*.ico*"
      TabIndex        =   0
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "File Name"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   6000
      Width           =   6855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
File1 = Dir1
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1 = Drive1
End Sub

Private Sub File1_Click()
Dim Mypic As String
Call Reset_Image
If Right(File1.Path, 1) = "\" Then
   Mypic = File1.Path & File1.filename
   Else
   Mypic = File1.Path & "\" & File1.filename
End If

Image1.Picture = LoadPicture(Mypic)
Label1.Caption = Mypic

End Sub



Private Sub Reset_Image()
    Image1.Visible = False


    If Image1.Picture Then
        '~ this is used in case the image change
        '     s
        '~ if it's not used, the image control i
        '     s
        '~ still the same size as the previous p
        '     ic
        Image1.Height = Image1.Picture.Height
        Image1.Width = Image1.Picture.Width


        If Image1.Picture.Height > Image1.Picture.Width Then
            '~ the Pic is taller than wide
            Image1.Height = Picture1.Height
            Image1.Width = Image1.Width / (Image1.Picture.Height / Image1.Height)


            If Image1.Width > Picture1.Width Then
                '~ If the PictureBox isn't square, the p
                '     ic still may be larger than it
                Image1.Width = Picture1.Width
                Image1.Height = Image1.Picture.Height / (Image1.Picture.Width / Image1.Width)
            End If
        End If


        If Image1.Picture.Width > Image1.Picture.Height Then
            '~ Image is wider than tall
            Image1.Width = Picture1.Width
            Image1.Height = Image1.Height / (Image1.Picture.Width / Image1.Width)


            If Image1.Height > Picture1.Height Then
                Image1.Height = Picture1.Height
                Image1.Width = Image1.Picture.Width / (Image1.Picture.Height / Image1.Height)
            End If
        End If
        '~ Center Image1 within Picture1
        Image1.Left = (Picture1.Width / 2) - (Image1.Width / 2)
        Image1.Top = (Picture1.Height / 2) - (Image1.Height / 2)
        Image1.Visible = True
    End If
End Sub




