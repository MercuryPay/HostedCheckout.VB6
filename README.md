HostedCheckout.VB6
====================

A VB6 application that demonstrates integrating to HostedCheckout using a webbrowser control.

Note: depending on which operating system you are using for your development environment you may need to experiment with the MSXML2 control and the webbrowser control.  This project uses Windows&trade; 7 64bit and requires the XMLHTTP60 vs. XMLHTTP20 interface.

>There are 3 steps to process a payment with Mercury's Hosted Checkout platform.

##Step 1: Initialize Payment


###Process: Initialize Payment Transaction

```
    Dim xmlHtp As New MSXML2.XMLHTTP60
    
    sURL = "https://hc.mercurydev.net/hcws/HCService.asmx"
    
    sEnv = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:xsd=""http://www.mercurypay.com/"">"
    sEnv = sEnv & "<soapenv:Header/>"
    sEnv = sEnv & "<soapenv:Body>"
    sEnv = sEnv & "<xsd:InitializePayment>"
    sEnv = sEnv & "<xsd:request>"
        sEnv = sEnv & "<xsd:MerchantID></xsd:MerchantID>"
        sEnv = sEnv & "<xsd:Password></xsd:Password>"
        sEnv = sEnv & "<xsd:Invoice>3472</xsd:Invoice>"
        sEnv = sEnv & "<xsd:TotalAmount>7.50</xsd:TotalAmount>"
        sEnv = sEnv & "<xsd:TaxAmount>0</xsd:TaxAmount>"
        sEnv = sEnv & "<AVSAddress />"
        sEnv = sEnv & "<AVSZip />"
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
    End With
```

###Parse: Response

```
sResp = .responseText
Text1.Text = .responseText
found1 = InStr(1, sResp, "ResponseCode", vbTextCompare)

```

##Step 2: Display HostedCheckout

>Display the HostedCheckout Web page

In this case the Navigate method of the webbrowser control is called to display the HostedCheckout page

```
WebBrowser1.Navigate "https://hc.mercurydev.net/CheckoutPOSiFrame.aspx?pid=" & txtPaymentid.Text

```

##Step 3: Verify Payment

###Process: Verify Transaction

```
    Dim xmlHtp As New MSXML2.XMLHTTP60
    
    sURL = "https://hc.mercurydev.net/hcws/HCService.asmx"
    
    sEnv = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:xsd=""http://www.mercurypay.com/"">"
    sEnv = sEnv & "<soapenv:Header/>"
    sEnv = sEnv & "<soapenv:Body>"
    sEnv = sEnv & "<xsd:VerifyPayment>"
    sEnv = sEnv & "<xsd:request>"
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
End With
    
```

###Parse: Response

>Approved transactions will have a CmdStatus equal to "Approved".

```
sResp = .responseText
Text1.Text = .responseText
found1 = InStr(1, sResp, "ResponseCode", vbTextCompare)
```

###©2015 Mercury Payment Systems, LLC - all rights reserved.

Disclaimer:
This software and all specifications and documentation contained herein or provided to you hereunder (the "Software") are provided free of charge strictly on an "AS IS" basis. No representations or warranties are expressed or implied, including, but not limited to, warranties of suitability, quality, merchantability, or fitness for a particular purpose (irrespective of any course of dealing, custom or usage of trade), and all such warranties are expressly and specifically disclaimed. Mercury Payment Systems shall have no liability or responsibility to you nor any other person or entity with respect to any liability, loss, or damage, including lost profits whether foreseeable or not, or other obligation for any cause whatsoever, caused or alleged to be caused directly or indirectly by the Software. Use of the Software signifies agreement with this disclaimer notice.


