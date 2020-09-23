VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H0014CCFA&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shop Bot - Coded by Machiavellian"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timTimeout2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   720
      Top             =   120
   End
   Begin VB.Timer timTimeout 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   120
   End
   Begin VB.Timer timRefresh 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7080
      Top             =   1800
   End
   Begin VB.Timer timCheck 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7080
      Top             =   2400
   End
   Begin VB.PictureBox picCode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   2520
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   79
      TabIndex        =   38
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin TabDlg.SSTab tabMain 
      Height          =   4215
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   7435
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   1363194
      TabCaption(0)   =   "Main"
      TabPicture(0)   =   "Shop Bot Main.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblItemInformation"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblAvailableItems"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblDescription"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblPrice"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblName"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblStock"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblStatus"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtName"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtDescription"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtPrice"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdAutoRun"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lstItems"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtStock"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdStop"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Options"
      TabPicture(1)   =   "Shop Bot Main.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblCheckTime"
      Tab(1).Control(1)=   "lblRefreshTime"
      Tab(1).Control(2)=   "lblWaitTime"
      Tab(1).Control(3)=   "lblAutorunSettings"
      Tab(1).Control(4)=   "lblShopSettings"
      Tab(1).Control(5)=   "lblHaggle"
      Tab(1).Control(6)=   "lblRefreshDelay"
      Tab(1).Control(7)=   "lblShopId"
      Tab(1).Control(8)=   "lblReference"
      Tab(1).Control(9)=   "lblFilter"
      Tab(1).Control(10)=   "txtCheckTime"
      Tab(1).Control(11)=   "txtRefreshTime"
      Tab(1).Control(12)=   "txtWaitTime"
      Tab(1).Control(13)=   "txtHaggle"
      Tab(1).Control(14)=   "txtRefreshDelay"
      Tab(1).Control(15)=   "txtShopId"
      Tab(1).Control(16)=   "txtReference"
      Tab(1).Control(17)=   "lstFilter"
      Tab(1).ControlCount=   18
      TabCaption(2)   =   "Logs"
      TabPicture(2)   =   "Shop Bot Main.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblAttempted"
      Tab(2).Control(1)=   "lblSuccess"
      Tab(2).Control(2)=   "lstAttempted"
      Tab(2).Control(3)=   "lstSuccess"
      Tab(2).Control(4)=   "cmdHaggle"
      Tab(2).ControlCount=   5
      Begin VB.CommandButton cmdHaggle 
         Caption         =   "Haggle Browser"
         Height          =   495
         Left            =   -72000
         TabIndex        =   41
         Top             =   3360
         Width           =   2535
      End
      Begin VB.ListBox lstFilter 
         Appearance      =   0  'Flat
         Height          =   2955
         ItemData        =   "Shop Bot Main.frx":0054
         Left            =   -72000
         List            =   "Shop Bot Main.frx":0056
         Sorted          =   -1  'True
         TabIndex        =   31
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox txtReference 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73440
         TabIndex        =   23
         Text            =   "1"
         Top             =   2400
         Width           =   1215
      End
      Begin VB.ListBox lstSuccess 
         Appearance      =   0  'Flat
         Height          =   2370
         ItemData        =   "Shop Bot Main.frx":0058
         Left            =   -72000
         List            =   "Shop Bot Main.frx":005A
         TabIndex        =   35
         Top             =   840
         Width           =   2535
      End
      Begin VB.ListBox lstAttempted 
         Appearance      =   0  'Flat
         Height          =   2955
         ItemData        =   "Shop Bot Main.frx":005C
         Left            =   -74760
         List            =   "Shop Bot Main.frx":005E
         TabIndex        =   33
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox txtShopId 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73440
         TabIndex        =   16
         Text            =   "58"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtRefreshDelay 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73440
         TabIndex        =   18
         Text            =   "2"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtHaggle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73440
         TabIndex        =   20
         Text            =   "97"
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox txtWaitTime 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73440
         TabIndex        =   29
         Text            =   "160"
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox txtRefreshTime 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73440
         TabIndex        =   27
         Text            =   "90"
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox txtCheckTime 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73440
         TabIndex        =   25
         Text            =   "60"
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   375
         Left            =   2820
         TabIndex        =   13
         Top             =   3600
         Width           =   2715
      End
      Begin VB.TextBox txtStock 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1680
         Width           =   1215
      End
      Begin VB.ListBox lstItems 
         Appearance      =   0  'Flat
         Height          =   2370
         ItemData        =   "Shop Bot Main.frx":0060
         Left            =   240
         List            =   "Shop Bot Main.frx":0062
         TabIndex        =   2
         Top             =   720
         Width           =   2535
      End
      Begin VB.CommandButton cmdAutoRun 
         Caption         =   "Run"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   3600
         Width           =   2595
      End
      Begin VB.TextBox txtPrice 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtDescription 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   1125
         Left            =   3000
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   2280
         Width           =   2535
      End
      Begin VB.TextBox txtName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   3120
         Width           =   2535
      End
      Begin VB.Label lblFilter 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Filter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72000
         TabIndex        =   30
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label lblReference 
         BackStyle       =   0  'Transparent
         Caption         =   "Reference Shop"
         Height          =   255
         Left            =   -74760
         TabIndex        =   22
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label lblSuccess 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Successful Items"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72000
         TabIndex        =   34
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label lblAttempted 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Attempted Items"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   32
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label lblShopId 
         BackStyle       =   0  'Transparent
         Caption         =   "Shop Id"
         Height          =   255
         Left            =   -74760
         TabIndex        =   15
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblRefreshDelay 
         BackStyle       =   0  'Transparent
         Caption         =   "Refresh Delay"
         Height          =   255
         Left            =   -74760
         TabIndex        =   17
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblHaggle 
         BackStyle       =   0  'Transparent
         Caption         =   "Haggle %"
         Height          =   255
         Left            =   -74760
         TabIndex        =   19
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lblShopSettings 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Shop"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   14
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label lblAutorunSettings 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Run"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   21
         Top             =   2040
         Width           =   2535
      End
      Begin VB.Label lblWaitTime 
         BackStyle       =   0  'Transparent
         Caption         =   "Wait Interval"
         Height          =   255
         Left            =   -74760
         TabIndex        =   28
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label lblRefreshTime 
         BackStyle       =   0  'Transparent
         Caption         =   "Refresh Interval"
         Height          =   255
         Left            =   -74760
         TabIndex        =   26
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label lblCheckTime 
         BackStyle       =   0  'Transparent
         Caption         =   "Check Interval"
         Height          =   255
         Left            =   -74760
         TabIndex        =   24
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label lblStock 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock"
         Height          =   255
         Left            =   4320
         TabIndex        =   8
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   255
         Left            =   3000
         TabIndex        =   4
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label lblPrice 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         Height          =   255
         Left            =   3000
         TabIndex        =   6
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   255
         Left            =   3000
         TabIndex        =   10
         Top             =   2040
         Width           =   2535
      End
      Begin VB.Label lblAvailableItems 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Available Items"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label lblItemInformation 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Item Information"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   3
         Top             =   480
         Width           =   2535
      End
   End
   Begin SHDocVwCtl.WebBrowser webShop 
      Height          =   615
      Left            =   6360
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   1560
      Width           =   615
      ExtentX         =   1085
      ExtentY         =   1085
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   6360
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin SHDocVwCtl.WebBrowser webHaggle 
      Height          =   3855
      Left            =   240
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   5280
      Width           =   5775
      ExtentX         =   10186
      ExtentY         =   6800
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin SHDocVwCtl.WebBrowser webCheck 
      Height          =   615
      Left            =   6360
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   2280
      Width           =   615
      ExtentX         =   1085
      ExtentY         =   1085
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Image imgLogo 
      Height          =   375
      Left            =   2400
      Picture         =   "Shop Bot Main.frx":0064
      Top             =   120
      Width           =   1500
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Neopets Shop Bot (Or Autobuyer)
'Coded by Machiavellian

'This is an autobuyer for the online game Neopets. If you're looking for this, you
'probably know what an autobuyer does. Everything was coded by me for me, but since I
'don't play Neopets anymore I decided to make it available to the public. Enjoy.

Private Type Item
    Name As String
    Description As String
    Url As String
    Stock As Integer
    Price As Long
End Type

Dim udtItem() As Item
Dim udtBuy As Item

Const strShopUrl As String = "http://www.neopets.com/objects.phtml?type=shop&obj_type="
Dim lonOffered As Long
Dim bytPostData() As Byte
Dim bolWait As Boolean
Dim bolStop As Boolean
Dim intCount As Integer
Dim intCheck As Integer
Dim bytState As Byte
Dim intTimeout As Integer
Dim intTimeout2 As Integer

Private Sub Run()
    bolStop = False
    webShop.Navigate strShopUrl & txtShopId.Text
    frmMain.Caption = "Shop Bot - Refreshing..."
End Sub

Private Sub EndRun()
    timTimeout2.Enabled = False
    bolStop = True
    timRefresh.Enabled = False
    frmMain.Caption = "Shop Bot - Inactive"
    intCount = 0
End Sub

Private Sub cmdAutoRun_Click()
    bytState = 0
    intCheck = 0
    webCheck.Navigate "http://www.neopets.com/objects.phtml?type=shop&obj_type=" & txtReference.Text

    txtReference.Enabled = False
    txtCheckTime.Enabled = False
    txtRefreshTime.Enabled = False
    txtWaitTime.Enabled = False
End Sub

Private Sub cmdHaggle_Click()
    If frmMain.Height = 9870 Then
        frmMain.Height = 5550
    Else
        frmMain.Height = 9870
    End If
End Sub

Private Sub cmdStop_Click()
    EndRun
    timCheck.Enabled = False
    lblStatus.Caption = ""

    txtReference.Enabled = True
    txtCheckTime.Enabled = True
    txtRefreshTime.Enabled = True
    txtWaitTime.Enabled = True
End Sub

Private Sub timTimeout_Timer()
    intTimeout = intTimeout + 1

    If intTimeout = 10 Then
        Debug.Print "Timeout"
        intTimeout = 0
        timTimeout.Enabled = False
        webCheck.Navigate "http://www.neopets.com/objects.phtml?type=shop&obj_type=" & txtReference.Text
    End If
End Sub

Private Sub timCheck_Timer()
    Select Case bytState

    Case Is = 0
        lblStatus.Caption = "Checking in: " & txtCheckTime.Text - intCheck
        If intCheck = txtCheckTime.Text Then
            intCheck = 0
            webCheck.Navigate "http://www.neopets.com/objects.phtml?type=shop&obj_type=" & txtReference.Text
            timCheck.Enabled = False
        End If
    Case Is = 1
        lblStatus.Caption = "Refreshing for: " & txtRefreshTime.Text - intCheck
        If intCheck = txtRefreshTime.Text Then
            intCheck = 0
            bytState = 2
            EndRun
        End If
    Case Is = 2
        lblStatus.Caption = "Waiting for: " & txtWaitTime.Text - intCheck
        If intCheck = txtWaitTime.Text Then
            bolWait = False
            intCheck = 0
            bytState = 0
        End If
    End Select

    intCheck = intCheck + 1
End Sub

Private Sub timTimeout2_Timer()
    intTimeout2 = intTimeout2 + 1

    If intTimeout2 = 4 Then
        Debug.Print "Timeout Refreshing"
        intTimeout2 = 0
        timTimeout2.Enabled = False
        timRefresh.Enabled = True
    End If
End Sub

Private Sub webCheck_DownloadBegin()
    intTimeout = 0
    timTimeout.Enabled = True
End Sub

Private Sub webCheck_DownloadComplete()
    Dim strSource As String
    
    timTimeout.Enabled = False

    If InStr(1, webCheck.LocationURL, "http://www.neopets.com/objects.phtml") <> 0 Then
        strSource = webCheck.Document.body.innerhtml

        If InStr(1, strSource, "in stock") <> 0 Then
            Run
            bytState = 1
            timCheck.Enabled = True
        Else
            bytState = 0
            timCheck.Enabled = True
        End If
    End If
End Sub

Private Sub Form_Load()
    LoadFilters
End Sub

Private Sub lstItems_Click()
    txtName.Text = udtItem(lstItems.ListIndex).Name
    txtPrice.Text = udtItem(lstItems.ListIndex).Price
    txtStock.Text = udtItem(lstItems.ListIndex).Stock
    txtDescription.Text = udtItem(lstItems.ListIndex).Description
End Sub

Private Sub lstItems_DblClick()
    BuyItem (lstItems.ListIndex)
End Sub

Private Sub timRefresh_Timer()
    If intCount = txtRefreshDelay.Text Then
        intCount = 0
        Run
        timRefresh.Enabled = False
    Else
        frmMain.Caption = "Shop Bot - Refreshing In " & txtRefreshDelay.Text - intCount & " Seconds"
        intCount = intCount + 1
    End If
End Sub

Private Sub txtShopId_Change()
    LoadFilters
End Sub

Private Sub LoadFilters()
    lstFilter.Clear
    Dim strInput As String
    On Error GoTo Error:
    Open App.Path & "\Data\" & txtShopId.Text & ".txt" For Input As #1
        Do Until EOF(1)
            Input #1, strInput
            lstFilter.AddItem strInput
        Loop
    Close #1
Error:
End Sub

Private Sub webShop_DownloadBegin()
    intTimeout2 = 0
    timTimeout2.Enabled = True
End Sub

Private Sub webShop_DownloadComplete()
    Dim strSource As String
    Dim strItem As String
    Dim intStart As Integer
    Dim intEnd As Integer
    Dim intItemCount As Integer

    timTimeout2.Enabled = False
    
    If InStr(1, webShop.LocationURL, "http://www.neopets.com/objects.phtml") <> 0 Then

    strSource = webShop.Document.body.innerhtml

    lstItems.Clear
    intStart = 1
    intEnd = 1

    Do
        intStart = InStr(intEnd, strSource, "obj_info_id")
        If intStart = 0 Then Exit Do
        intEnd = InStr(intStart, strSource, " <BR><BR></TD>")

        strItem = Mid(strSource, intStart, intEnd - intStart)
        strItem = Replace(strItem, "amp;", "")
        strItem = Replace(strItem, "&g=3""><IMG height=80 alt=""", vbCrLf)
        strItem = Replace(strItem, """ src=""", vbCrLf)
        strItem = Replace(strItem, "&g=3""><IMG height=80 alt='", vbCrLf)
        strItem = Replace(strItem, "' src=""", vbCrLf)
        strItem = Replace(strItem, "&g=3""><IMG height=80 alt=", vbCrLf)
        strItem = Replace(strItem, " src=""", vbCrLf)
        strItem = Replace(strItem, """ width=80 border=1></A><BR><B>", vbCrLf)
        strItem = Replace(strItem, "</B><BR>", vbCrLf)
        strItem = Replace(strItem, " in stock<BR>Cost : ", vbCrLf)
        strItem = Replace(strItem, " NP", "")
        ReDim Preserve udtItem(intItemCount)

        udtItem(intItemCount).Url = Split(strItem, vbCrLf)(0)
        udtItem(intItemCount).Description = Split(strItem, vbCrLf)(1)
        udtItem(intItemCount).Name = Split(strItem, vbCrLf)(3)
        udtItem(intItemCount).Stock = Split(strItem, vbCrLf)(4)
        udtItem(intItemCount).Price = Replace(Split(strItem, vbCrLf)(5), ",", "")
        
        lstItems.AddItem udtItem(intItemCount).Name
        
        For x% = 1 To lstFilter.ListCount
            If InStr(1, udtItem(intItemCount).Name, lstFilter.List(x% - 1)) <> 0 Then
                BuyItem (intItemCount)
                Exit For
            End If
        Next x%

        intItemCount = intItemCount + 1
        DoEvents
    Loop
        If bolStop = False Then
            timRefresh.Enabled = True
        End If
    End If
End Sub

Private Sub webHaggle_DownloadComplete()
    Dim strSource As String
    Dim strPostData As String
    Dim strHeader As String
    Dim varPostData As Variant
    
    If InStr(1, webHaggle.LocationURL, "http://www.neopets.com/") <> 0 Then

    strSource = webHaggle.Document.body.innerhtml

    If InStr(1, strSource, "Try and Haggle!") <> 0 Then
        lonOffered = Int((txtHaggle.Text / 100 * udtBuy.Price))
        strPostData = "brr=1366&current_offer=" & lonOffered
        BuildPostData bytPostData(), strPostData
        varPostData = bytPostData

        strHeader = "Referer: " & webShop.LocationURL + Chr(10) + Chr(13) + "Content-Type: application/x-www-form-urlencoded" + Chr(10) + Chr(13)
        webHaggle.Navigate2 "http://www.neopets.com/haggle.phtml", 0, "", varPostData, strHeader
    ElseIf InStr(1, strSource, "rscheck.phtml") <> 0 Then
        strPostData = "current_offer=" & lonOffered & "&" & udtBuy.Url & "&brr=1366&rscheck=" & strCode
        BuildPostData bytPostData(), strPostData
        varPostData = bytPostData

        strHeader = "Referer: " & webHaggle.LocationURL + Chr(10) + Chr(13) + "Content-Type: application/x-www-form-urlencoded" + Chr(10) + Chr(13)
        webHaggle.Navigate2 "http://www.neopets.com/haggle.phtml", 0, "", varPostData, strHeader
    ElseIf InStr(1, strSource, "added to your inventory") <> 0 Then
        bolWait = False
        lstSuccess.AddItem udtBuy.Name
        Debug.Print udtBuy.Name & " : " & udtBuy.Price
    ElseIf InStr(1, strSource, "is SOLD OUT") <> 0 Then
        bolWait = False
        Debug.Print udtBuy.Name & " : Sold out"
        Debug.Print webHaggle.LocationURL
    ElseIf InStr(1, strSource, "to massive demand on the Neopian Shops") <> 0 Then
        bolWait = False
        Debug.Print udtBuy.Name & " : Massive demand"
    ElseIf InStr(1, strSource, "taking stock away from the shop") <> 0 Then
        bolWait = False
        Debug.Print udtBuy.Name & " : Error taking stock"
    ElseIf InStr(1, strSource, "The shop doesnt have any of this item left!") <> 0 Then
        bolWait = False
        Debug.Print udtBuy.Name & " : No stock left"
    ElseIf Len(strSource) = 0 Then
        bolWait = False
    Else
        bolWait = False
        Debug.Print "Unknown error"
    End If

    End If
End Sub

Private Sub BuyItem(intItemIndex As Integer)
    If bolWait = False Then
        bolWait = True
        udtBuy = udtItem(intItemIndex)
        lstAttempted.AddItem udtBuy.Name
        webHaggle.Navigate "http://www.neopets.com/haggle.phtml?" & udtBuy.Url & "&brr1366", , , , "Referer: " & webShop.LocationURL
    End If
End Sub

Private Function strCode() As String
    Dim strLetter As String
    Dim strBuffer As String
    Dim strInput As String
    Dim strData As String
    Dim strCharacter As String

    Dim intScore As Integer
    Dim intHighScore As Integer
    Dim intCount As Integer
    
    Dim ax As Byte
    Dim ay As Byte
    Dim by As Byte

    Dim Image() As Byte

    Image() = Inet.OpenURL("http://www.neopets.com/rscheck.phtml", icByteArray)

    Do
    DoEvents
    Loop While Inet.StillExecuting

    Open App.Path & "\" & "Code.gif" For Binary As #1
    Put #1, , Image()
    Close #1

    picCode.Picture = LoadPicture("Code.gif")

    For z% = 0 To 2
    strLetter = ""
    strBuffer = ""
    intHighScore = 0
    ax = 255
    ay = 255

    For y% = 0 To 30
    For x% = 0 To 22
        If picCode.Point(x% + z% * 24, y%) = 8421504 And picCode.Point(x% + z% * 24 + 1, y%) = 8421504 And picCode.Point(x% + z% * 24 + 2, y%) = 8421504 Then
            If ax > x% Then ax = x%
            If ay > y% Then ay = y%
            If by < y% Then by = y%

            strBuffer = strBuffer & "000"
            x% = x% + 2
        Else
            strBuffer = strBuffer & " "
        End If
    Next x%
    Next y%

    For y% = ay To by
        strLetter = strLetter & Mid(strBuffer, ax + 1 + y% * 23, 18)
    Next y%

    Open App.Path & "\OCR.dat" For Input As #1
        Do While Not EOF(1)
            Line Input #1, strInput

            strData = Split(strInput, ":")(1)
            intScore = 0
            intCount = 0

            For x% = 1 To Len(strLetter)
                If Mid(strLetter, x%, 1) = "0" Then
                    intCount = intCount + 1
                End If
                
                If Mid(strData, x%, 1) = "0" Then
                    intCount = intCount - 1
                End If
            
                If Mid(strLetter, x%, 1) = "0" And Mid(strData, x%, 1) = "0" Then
                    intScore = intScore + 1
                End If
            Next x%

            intScore = intScore - Abs(intCount) / 2
            
            If intScore > intHighScore Then
                intHighScore = intScore
                strCharacter = Split(strInput, ":")(0)
            End If
        Loop
    Close #1

    strCode = strCode & strCharacter
    Next z%
End Function

Private Sub BuildPostData(ByteArray() As Byte, ByVal strPostData As String)
    Dim intNewBytes As Integer
    Dim strCH As String
    Dim i As Integer

    intNewBytes = Len(strPostData) - 1

    If intNewBytes < 0 Then
        Exit Sub
    End If
    
    ReDim ByteArray(intNewBytes)
    
    For i = 0 To intNewBytes
        strCH = Mid$(strPostData, i + 1, 1)

        If strCH = Space(1) Then
            strCH = "+"
        End If
     
        ByteArray(i) = Asc(strCH)
    Next
End Sub
