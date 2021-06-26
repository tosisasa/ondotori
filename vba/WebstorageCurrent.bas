Attribute VB_Name = "WebstorageCurrent"
Option Explicit


Sub ���ݒl�擾_Click()
    '�������i�`�撆�~�j
    Application.ScreenUpdating = False
    Application.StatusBar = "������..."
    
    
    Dim iRow As Integer
    iRow = 10
    
    
    
    ' ����ǂƂ� WebStorage API�d�l
    ' https://ondotori.webstorage.jp/docs/api/
    '
    ' ���ݒl�̎擾
    ' https://ondotori.webstorage.jp/docs/api/reference/devices_device.html
    
    
    Dim httpReq As New XMLHTTP60   '�uMicrosoft XML, v6.0�v���Q�Ɛݒ�
    
    httpReq.Open "POST", "https://api.webstorage.jp/v1/devices/current"
    httpReq.setRequestHeader "Host", "api.webstrage.js:443"
    httpReq.setRequestHeader "Content-Type", "application/json"
    httpReq.setRequestHeader "X-HTTP-Method-Override", "GET"
    
    
    '���N�G�X�g�{�f�B
    'API KEY�͊Ǘ���ID��Web Storage�Ƀ��O�C�����A�u�J���Ҍ���API�Ǘ��v���甭�s����B
    '���O�C��ID�ƃp�X���[�h�́A�{���p�A�J�E���g���w��B
    
    
    Dim apikey As String
    Dim loginid As String
    Dim loginpass As String
    
    apikey = "API�L�["
    loginid = "�{���pID"
    loginpass = "�{���p�A�J�E���g�̃p�X���[�h"

    
    Dim sRequestBody As String
    sRequestBody = "{"
    sRequestBody = sRequestBody + """api-key"":""" & apikey & """"
    sRequestBody = sRequestBody + ",""login-id"":""" & loginid & """"
    sRequestBody = sRequestBody + ",""login-pass"":""" & loginpass & """"
    sRequestBody = sRequestBody + "}"
    
    httpReq.send sRequestBody
      

    Do While httpReq.readyState < 4
        DoEvents
    Loop

    

    'VBA-JSON
    ' https://github.com/VBA-tools/VBA-JSON/releases/tag/v2.3.1

    Dim jsonObj As Object
    Set jsonObj = JsonConverter.ParseJson(httpReq.responseText)
    
    
    Dim dicDevices As Dictionary
    
    For Each dicDevices In jsonObj("devices")
    
        Dim sGroup As String
        
        Dim dicGroup As Dictionary
        Set dicGroup = dicDevices.Item("group")
        sGroup = dicGroup("name")
    
    
        Dim sUnixtime As String
        sUnixtime = dicDevices("unixtime")
        
        
        Dim sName As String
        sName = dicDevices("name")
        
        
        Dim sCh1Name As String
        Dim sCh1 As String
        
        Dim sCh2Name As String
        Dim sCh2 As String
        
        Dim iChannel As Integer
        iChannel = 0
        
        Dim dicChannel As Dictionary
        For Each dicChannel In dicDevices.Item("channel")
            If iChannel = 0 Then
                sCh1Name = dicChannel("name")
                sCh1 = dicChannel("value")
                iChannel = 1
            Else
                sCh2Name = dicChannel("name")
                sCh2 = dicChannel("value")
            End If
        Next
    
        ActiveSheet.Cells(iRow, 1).Value = sGroup
   
    
        ActiveSheet.Cells(iRow, 2).Value = (sUnixtime + 32400) / 86400 + 25569
        '=(sunixtime + 32400) / 86400 + 25569
    
        ActiveSheet.Cells(iRow, 3).Value = sName
        ActiveSheet.Cells(iRow, 4).Value = sCh1Name
        ActiveSheet.Cells(iRow, 5).Value = sCh1
        ActiveSheet.Cells(iRow, 6).Value = sCh2Name
        ActiveSheet.Cells(iRow, 7).Value = sCh2
    
        iRow = iRow + 1
    
    Next


    Application.ScreenUpdating = True
    Application.StatusBar = ""
End Sub

