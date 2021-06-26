Attribute VB_Name = "webstoragedata"
Option Explicit

Sub 指定期間件数取得_Click()
    '高速化（描画中止）
    Application.ScreenUpdating = False
    Application.StatusBar = "検索中..."


    Dim sStartDt, sEndDt As String
    sStartDt = ActiveSheet.Cells(1, 2).Value
    sEndDt = ActiveSheet.Cells(2, 2).Value
    
    If Trim(sStartDt) = "" Or Trim(sEndDt) = "" Then
        MsgBox "日付範囲を入力してください。", vbOKOnly + vbExclamation
        Application.ScreenUpdating = True
        Application.StatusBar = ""
        Exit Sub
    End If
    
    'クリア処理
    Dim i As Long
    For i = ActiveSheet.ListObjects(1).ListRows.Count To 1 Step -1
      ActiveSheet.ListObjects(1).ListRows.Item(i).Delete
    Next i

    
    Dim iRow As Integer
    iRow = 10
        
    Dim sFromIn As String
    Dim sToIn As String
    
    sFromIn = ActiveSheet.Cells(1, 3).Value
    sToIn = ActiveSheet.Cells(2, 3).Value
    
    
    ' おんどとり WebStorage API仕様
    ' https://ondotori.webstorage.jp/docs/api/
    
    ' 指定期間・件数によるデータの取得
    ' https://ondotori.webstorage.jp/docs/api/reference/devices_data.html
    
    Dim httpReq As New XMLHTTP60   '「Microsoft XML, v6.0」を参照設定
    
    httpReq.Open "POST", "https://api.webstorage.jp/v1/devices/data"
    httpReq.setRequestHeader "Host", "api.webstrage.js:443"
    httpReq.setRequestHeader "Content-Type", "application/json"
    httpReq.setRequestHeader "X-HTTP-Method-Override", "GET"
    
    
    'リクエストボディ
    'API KEYは管理者IDでWeb Storageにログインし、「開発者向けAPI管理」から発行する。
    'ログインIDとパスワードは、閲覧用アカウントを指定。
    
    'Web Storage
    Dim apikey As String
    Dim loginid As String
    Dim loginpass As String
    Dim remoteserial As String
    
    apikey = "APIキー"
    loginid = "閲覧用ID"
    loginpass = "閲覧用アカウントのパスワード"
    remoteserial = "機器のシリアル番号"
    
    
    Dim sRequestBody As String
    sRequestBody = "{"
    sRequestBody = sRequestBody + """api-key"":""" & apikey & """"
    sRequestBody = sRequestBody + ",""login-id"":""" & loginid & """"
    sRequestBody = sRequestBody + ",""login-pass"":""" & loginpass & """"
    sRequestBody = sRequestBody + ",""remote-serial"": """ & remoteserial & """"
    sRequestBody = sRequestBody + ",""unixtime-from"":" & sFromIn & ""
    sRequestBody = sRequestBody + ",""unixtime-to"":" & sToIn & ""
    sRequestBody = sRequestBody + ",""type"":""json"""
    sRequestBody = sRequestBody + "}"
    
    httpReq.send sRequestBody
      

    Do While httpReq.readyState < 4
        DoEvents
    Loop

    

    'VBA-JSON
    ' https://github.com/VBA-tools/VBA-JSON/releases/tag/v2.3.1
    Dim jsonObj As Object
    Set jsonObj = JsonConverter.ParseJson(httpReq.responseText)
    

    Dim dic As Dictionary
    For Each dic In jsonObj("data")
    
        Dim sUnixtime As String
        Dim sCh1 As String
        Dim sCh2 As String
    
        sUnixtime = dic("unixtime")
        sCh1 = dic("ch1")
        sCh2 = dic("ch2")
    
    
        ActiveSheet.Cells(iRow, 1).Value = (sUnixtime + 32400) / 86400 + 25569
        '=(sunixtime + 32400) / 86400 + 25569
    
        ActiveSheet.Cells(iRow, 2).Value = sCh1
        ActiveSheet.Cells(iRow, 3).Value = sCh2
        
        
        
        iRow = iRow + 1
    
    Next
    
    
    Application.ScreenUpdating = True
    Application.StatusBar = ""
End Sub

