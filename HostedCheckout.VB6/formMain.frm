VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form formMain 
   Caption         =   "Mercury HostedCheckout POS iFrame"
   ClientHeight    =   11220
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17835
   LinkTopic       =   "Form1"
   ScaleHeight     =   11220
   ScaleWidth      =   17835
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   7095
      Left            =   4440
      TabIndex        =   15
      Top             =   240
      Width           =   13095
      ExtentX         =   23098
      ExtentY         =   12515
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
   Begin VB.CommandButton cmdReset 
      Caption         =   "RESET"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   14
      Top             =   6360
      Width           =   1695
   End
   Begin VB.Frame frmStep2 
      Caption         =   "STEP 2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   240
      TabIndex        =   10
      Top             =   2520
      Width           =   3735
      Begin VB.Label lblStep2c 
         Caption         =   """PROCESS"""
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label lblStep2b 
         Caption         =   "Swipe or enter card data"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   3255
      End
      Begin VB.Label lblStep2a 
         Caption         =   "Load iFrame"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   2535
      End
   End
   Begin VB.CommandButton cmdVerify 
      Caption         =   "STEP 3:  Verify Payment"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   9
      Top             =   5040
      Width           =   3735
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "CLOSE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      TabIndex        =   8
      Top             =   6360
      Width           =   1695
   End
   Begin VB.TextBox txtPaymentid 
      Alignment       =   2  'Center
      Height          =   405
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1800
      Width           =   3255
   End
   Begin VB.TextBox txtCmdStatus 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   7800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdInitPayment 
      Caption         =   "STEP 1: Initialize Payment"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   2655
      Left            =   9480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   8400
      Width           =   8175
   End
   Begin VB.TextBox Text2 
      Height          =   2655
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   8400
      Width           =   8175
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   17760
      Y1              =   7560
      Y2              =   7560
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Payment ID"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Response"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11400
      TabIndex        =   3
      Top             =   7920
      Width           =   4095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Request"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   7920
      Width           =   4095
   End
End
Attribute VB_Name = "formMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
    Unload formMain
End Sub
Private Sub cmdInitPayment_Click()
    Dim sURL As String
    Dim sEnv As String
    Dim sResp As String
    Dim xmlHtp As New MSXML2.XMLHTTP60
    
    sURL = "https://hc.mercurydev.net/hcws/HCService.asmx"
    
    sEnv = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:xsd=""http://www.mercurypay.com/"">"
    sEnv = sEnv & "<soapenv:Header/>"
    sEnv = sEnv & "<soapenv:Body>"
    sEnv = sEnv & "<xsd:InitializePayment>"
    sEnv = sEnv & "<xsd:request>"
        sEnv = sEnv & "<xsd:MerchantID>494691720</xsd:MerchantID>"
        sEnv = sEnv & "<xsd:Password>KRD%8rw#+p9C13,T</xsd:Password>"
        sEnv = sEnv & "<xsd:Invoice>3472</xsd:Invoice>"
        sEnv = sEnv & "<xsd:TotalAmount>7.50</xsd:TotalAmount>"
        sEnv = sEnv & "<xsd:TaxAmount>0</xsd:TaxAmount>"
        sEnv = sEnv & "<AVSAddress />"
        sEnv = sEnv & "<AVSZip />"
        sEnv = sEnv & "<xsd:TranType>Sale</xsd:TranType>"
        sEnv = sEnv & "<xsd:CardHolderName>Christine Jennings</xsd:CardHolderName>"
        sEnv = sEnv & "<xsd:Frequency>OneTime</xsd:Frequency>"
        sEnv = sEnv & "<xsd:Memo>VB6 HC POS iFrame</xsd:Memo>"
        sEnv = sEnv & "<xsd:ProcessCompleteUrl>http://www.mercurypay.com/developer-solutions/</xsd:ProcessCompleteUrl>"
        sEnv = sEnv & "<xsd:ReturnUrl>http://www.mercurypay.com/developer-solutions/</xsd:ReturnUrl>"
        sEnv = sEnv & "<DisplayStyle>Custom</DisplayStyle>"
        sEnv = sEnv & "<PageTitle>Merchant Processing..</PageTitle>"
        sEnv = sEnv & "<SecurityLogo>On</SecurityLogo>"
        sEnv = sEnv & "<OrderTotal>On</OrderTotal>"
        sEnv = sEnv & "<DefaultSwipe>Swipe</DefaultSwipe>"
        sEnv = sEnv & "<PageTimeoutDuration>0</PageTimeoutDuration>"
    sEnv = sEnv & "</xsd:request>"
    sEnv = sEnv & "</xsd:InitializePayment>"
    sEnv = sEnv & "</soapenv:Body>"
    sEnv = sEnv & "</soapenv:Envelope>"

    With xmlHtp
        .Open "post", sURL, False
        .setRequestHeader "Host", "w1.mercurypay.com"
        .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
        .setRequestHeader "soapAction", "http://www.mercurypay.com/InitializePayment"
        Text2.Text = sEnv
        .send sEnv
        sResp = .responseText
        Text1.Text = .responseText
        found1 = InStr(1, sResp, "ResponseCode", vbTextCompare)
        If found1 > 0 Then 'Found
            found1 = found1 + 13
            found2 = InStr(found1, sResp, "/ResponseCode", vbTextCompare)
            If found2 > found1 Then 'Found
                found2 = found2 - 1
                flength = found2 - found1
                txtCmdStatus.Text = Mid$(sResp, found1, flength)
                If txtCmdStatus.Text = 0 Then 'Good - Display iFrame
                    found1 = InStr(1, sResp, "PaymentID", vbTextCompare)
                    If found1 > 0 Then 'Found
                        found1 = found1 + 10
                        found2 = InStr(found1, sResp, "/PaymentID", vbTextCompare)
                        If found2 > found1 Then 'Found
                           found2 = found2 - 1
                           flength = found2 - found1
                           txtPaymentid.Text = Mid$(sResp, found1, flength)
                        End If
                    End If
                    WebBrowser1.Navigate "https://hc.mercurydev.net/CheckoutPOSiFrame.aspx?pid=" & txtPaymentid.Text
                    cmdInitPayment.Enabled = False
                    frmStep2.Enabled = True
                    lblStep2a.Enabled = True
                    lblStep2b.Enabled = True
                    lblStep2c.Enabled = True
                End If
            End If
        End If
    End With
End Sub
Private Sub cmdReset_Click()
    frmStep2.Enabled = False
    lblStep2a.Enabled = False
    lblStep2b.Enabled = False
    lblStep2c.Enabled = False
    cmdVerify.Enabled = False
    cmdInitPayment.Enabled = True
    WebBrowser1.Navigate App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & "start.html"
    Text2.Text = "TextRequest"
    Text1.Text = "TextResponse"
    txtPaymentid.Text = ""
End Sub
Private Sub cmdVerify_Click()
    Dim sURL As String
    Dim sEnv As String
    Dim sResp As String
    Dim xmlHtp As New MSXML2.XMLHTTP60
    
    sURL = "https://hc.mercurydev.net/hcws/HCService.asmx"
    
    sEnv = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:xsd=""http://www.mercurypay.com/"">"
    sEnv = sEnv & "<soapenv:Header/>"
    sEnv = sEnv & "<soapenv:Body>"
    sEnv = sEnv & "<xsd:VerifyPayment>"
    sEnv = sEnv & "<xsd:request>"
        sEnv = sEnv & "<xsd:MerchantID>494691720</xsd:MerchantID>"
        sEnv = sEnv & "<xsd:Password>KRD%8rw#+p9C13,T</xsd:Password>"
        sEnv = sEnv & "<xsd:PaymentID>" & txtPaymentid.Text & "</xsd:PaymentID>"
    sEnv = sEnv & "</xsd:request>"
    sEnv = sEnv & "</xsd:VerifyPayment>"
    sEnv = sEnv & "</soapenv:Body>"
    sEnv = sEnv & "</soapenv:Envelope>"
    
    With xmlHtp
        .Open "post", sURL, False
        .setRequestHeader "Host", "w1.mercurypay.com"
        .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
        .setRequestHeader "soapAction", "http://www.mercurypay.com/VerifyPayment"
        Text2.Text = sEnv
        .send sEnv
        sResp = .responseText
        Text1.Text = .responseText
        found1 = InStr(1, sResp, "ResponseCode", vbTextCompare)
        If found1 > 0 Then 'Found
            found1 = found1 + 13
            found2 = InStr(found1, sResp, "/ResponseCode", vbTextCompare)
            If found2 > found1 Then 'Found
                found2 = found2 - 1
                flength = found2 - found1
                txtCmdStatus.Text = Mid$(sResp, found1, flength)
                If txtCmdStatus.Text = 0 Then 'Good
                    WebBrowser1.Navigate App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & "transactioncomplete.html"
                    cmdVerify.Enabled = False
                End If
            End If
        End If
    End With
End Sub
Private Sub Form_Load()
    frmStep2.Enabled = False
    lblStep2a.Enabled = False
    lblStep2b.Enabled = False
    lblStep2c.Enabled = False
    cmdVerify.Enabled = False
    WebBrowser1.Navigate App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & "start.html"
End Sub
Private Sub WebBrowser1_BeforeNavigate2(ByVal pDisp As Object, _
           URL As Variant, Flags As Variant, _
           TargetFrameName As Variant, PostData As Variant, _
           Headers As Variant, Cancel As Boolean)
           
    If URL = "http://www.mercurypay.com/developer-solutions/" Then
        cmdVerify.Enabled = True
        frmStep2.Enabled = False
        lblStep2a.Enabled = False
        lblStep2b.Enabled = False
        lblStep2c.Enabled = False
        Cancel = True
        WebBrowser1.Navigate App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & "verify.html"
    End If
           
End Sub
