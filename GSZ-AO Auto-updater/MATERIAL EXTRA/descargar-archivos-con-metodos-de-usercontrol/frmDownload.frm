VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmTest 
   Caption         =   "Multiple async downloads without external components/classes or APIs"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleWidth      =   10380
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9780
      Top             =   4290
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":0000
            Key             =   "failed"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":08DA
            Key             =   "downloaded"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":11B4
            Key             =   "default"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownload.frx":192E
            Key             =   "downloading"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView myDownloads 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   7435
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Url"
         Object.Width           =   7585
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Bytes"
         Object.Width           =   4480
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Status"
         Object.Width           =   4410
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   0
      TabIndex        =   1
      Top             =   4260
      Width           =   10335
      Begin VB.CommandButton Command1 
         Caption         =   "Start the demonstration!"
         Height          =   405
         Left            =   2400
         TabIndex        =   3
         Top             =   990
         Width           =   2385
      End
      Begin AsyncDownloader.ctlDownload x 
         Height          =   960
         Left            =   8130
         TabIndex        =   2
         Top             =   150
         Width           =   960
         _extentx        =   1693
         _extenty        =   1693
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF8080&
         X1              =   6780
         X2              =   6780
         Y1              =   1680
         Y2              =   90
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Filipe Lage - fclage@ezlinkng.com"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   6
         Top             =   1440
         Width           =   3195
      End
      Begin VB.Label Label1 
         Caption         =   $"frmDownload.frx":20A8
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   90
         TabIndex        =   5
         Top             =   180
         Width           =   6645
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "This control is only visible if the project is running in VB6 (debug mode)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6840
         TabIndex        =   4
         Top             =   1140
         Width           =   3435
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
For Each n In Me.myDownloads.ListItems
x.Download n.Text
Next
End Sub

Private Sub Form_Activate()
Label2.Visible = x.Visible

End Sub

Private Sub Form_Load()

xurl = "http://www.google.com/"
Set n = myDownloads.ListItems.Add(, xurl, xurl, , "default")

xurl = "http://www.microsoft.com/"
Set n = myDownloads.ListItems.Add(, xurl, xurl, , "default")

xurl = "http://www.theinquirer.net/inquirer.rss"
Set n = myDownloads.ListItems.Add(, xurl, xurl, , "default")

xurl = "http://www.yahoo.com/"
Set n = myDownloads.ListItems.Add(, xurl, xurl, , "default")

xurl = "ftp://ftp.unb.br/pub/capes/Coleta63/coleta63.exe"
Set n = myDownloads.ListItems.Add(, xurl, xurl, , "default")

End Sub

Private Sub Form_Resize()
On Error Resume Next
Frame1.Move 0, ScaleHeight - Frame1.Height, ScaleWidth, Frame1.Height
Me.myDownloads.Move 0, 0, ScaleWidth, ScaleHeight - Frame1.Height
End Sub

Private Sub myDownloads_DblClick()
If myDownloads.SelectedItem Is Nothing Then Exit Sub
MsgBox myDownloads.SelectedItem.Tag, vbInformation, myDownloads.SelectedItem.Text
End Sub

Private Sub x_Finished(x As AsyncProperty)
Dim n As ListItem
Set n = myDownloads.ListItems(x.PropertyName)
If x.StatusCode = vbAsyncStatusCodeEndDownloadData Then
    n.SmallIcon = "downloaded"
    n.SubItems(1) = x.BytesRead & " bytes"
    n.SubItems(2) = "Downloaded"
    n.Tag = StrConv(x.Value, vbUnicode)
    Else
    n.SmallIcon = "failed"
    n.SubItems(2) = "Failed"
    n.Tag = ""
    End If
End Sub

Private Sub x_Progress(x As AsyncProperty, percent As Single)
Dim n As ListItem
Set n = myDownloads.ListItems(x.PropertyName)
n.SubItems(2) = "Downloading " & Format(percent, "0.0") & "%"
n.SmallIcon = "downloading"
n.SubItems(1) = x.BytesRead & " / " & x.BytesMax
End Sub
