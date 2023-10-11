'Imports NPOI.HSSF.UserModel
'Imports NPOI.SS.UserModel
'Imports ClsConfigTool
'Imports System.IO
'Imports NDde
'Imports NDde.Client
'Imports System.Net.Mail

'Public Class FrmPrint
'  Shared ClsConfig As ClsConfigTool.ConfigTool = New ClsConfigTool.ConfigTool(Application.StartupPath & "\PrintFormatSetting.xml")

'  '---for 只把excel整理到sample資料夾並copy 到特定資料夾
'  Public Shadows Sub MoveFile(ByVal _list As List(Of String), ByVal _queue As Queue(Of List(Of List(Of String))), ByVal _NAME As String, ByVal Attachefilename As String)
'    Try
'      Dim Original_samplePath = Application.StartupPath & "\OriginalSAMPLE" '---參考範本來源(必存在)
'      Dim _samplePath = Application.StartupPath & "\SAMPLE" '---print來源 (必存在)

'      '1.檢查folder,original ,sample 必要存在 
'      If CheckFolder(Original_samplePath, _samplePath) Then

'        '2.檢查OriginalSAMPLE資料夾內 是否有要列印的格式
'        For Each fname As String In System.IO.Directory.GetFileSystemEntries(Original_samplePath)
'          If Path.GetFileNameWithoutExtension(fname) = _NAME Then
'            '3.file 複製到sample資料夾內
'            Dim _destFile = _samplePath & "\" & Attachefilename & ".xls"
'            File.Copy(fname, _destFile)
'            '---main          
'            If ReadExcel(_destFile, _NAME, _list, _queue) Then
'              Dim name = Path.GetFileName(_destFile)
'              Dim _cpname = gMain.SaveFolderPath & "\" & name
'              File.Copy(_destFile, _cpname, True)
'              File.Delete(_destFile)
'            End If

'            Exit For
'          End If
'        Next
'      End If
'    Catch ex As Exception
'      MessageBox.Show(ex.Message)
'    End Try
'  End Sub

'  '---for 請款單
'  Public Shared Sub StartPrintBill(ByVal _path As String)
'    Try
'      printExcel(_path, ClsConfig.ReadStringValueKey("Public", "EXCEL執行檔"))
'    Catch ex As Exception
'      MessageBox.Show(ex.Message)
'    End Try
'  End Sub

'  '---for excel inoutput
'  Public Shared Sub StartPrint(ByVal _list As List(Of String), ByVal _queue As Queue(Of List(Of List(Of String))), ByVal _NAME As String)
'    Try
'      Dim Original_samplePath = Application.StartupPath & "\OriginalSAMPLE" '---參考範本來源(必存在)
'      Dim _samplePath = Application.StartupPath & "\SAMPLE" '---print來源 (必存在)

'      '1.檢查folder,original ,sample 必要存在 
'      If CheckFolder(Original_samplePath, _samplePath) Then

'        '2.檢查OriginalSAMPLE資料夾內 是否有要列印的格式
'        For Each fname As String In System.IO.Directory.GetFileSystemEntries(Original_samplePath)
'          If Path.GetFileNameWithoutExtension(fname) = _NAME Then
'            '3.file 複製到sample資料夾內
'            Dim _destFile = _samplePath & "\" & _NAME & ".xls"
'            File.Copy(fname, _destFile)
'            '---main          
'            If ReadExcel(_destFile, _NAME, _list, _queue) Then printExcel(_destFile, ClsConfig.ReadStringValueKey("Public", "EXCEL執行檔"))

'            Exit For
'          End If
'        Next
'      End If
'    Catch ex As Exception
'      MessageBox.Show(ex.Message)
'    End Try
'  End Sub
'  Private Shared Function CheckFolder(ByVal Original_samplePath As String, ByVal _samplePath As String) As Boolean
'    Try
'      '1檢查folder 是否存在
'      If Not Directory.Exists(Original_samplePath) Then Return False
'      If Not Directory.Exists(_samplePath) Then Directory.CreateDirectory(_samplePath) ' Return False

'      '2.清空sample Folder 內的的檔案
'      For Each fname As String In System.IO.Directory.GetFileSystemEntries(_samplePath)
'        File.Delete(fname)
'      Next

'      Return True
'    Catch ex As Exception
'      MessageBox.Show(ex.ToString)
'      Return False
'    End Try
'  End Function
'  Private Shared Function ReadExcel(ByVal sExcelPath As String, ByVal strName As String, ByVal _list As List(Of String), ByVal _queue As Queue(Of List(Of List(Of String)))) As Boolean
'    Try
'      '1.read
'      Dim wk As HSSFWorkbook
'      Using fs As New FileStream(sExcelPath, FileMode.Open, FileAccess.ReadWrite)
'        wk = New HSSFWorkbook(fs) '-抓整個excel
'        Dim hst = DirectCast(wk.GetSheetAt(0), HSSFSheet) '---抓一個sheet 理論上應該只有一個Sheet

'        '2.塞資料
'        Dim Public_dic As Dictionary(Of String, String) = ClsConfig.ReadKEYDictionary("Public")
'        Dim _strSection = Public_dic(strName)
'        Dim _dic As Dictionary(Of String, String) = ClsConfig.ReadKEYDictionary(_strSection)
'        Dim _count = 0
'        Dim _Tablecount = 0
'        For Each item In _dic
'          Dim _info = ClsConfig.ReadStringValueDictionary(_strSection, item.Key)
'          If _info("TABLE") = "0" Then
'            Dim hr = DirectCast(hst.GetRow(_info("Y") - 1), HSSFRow)
'            hr.Cells(_info("X") - 1).SetCellValue(_list(_count))
'            _count = _count + 1
'          Else '---TABLE表          
'            Dim Allist = _queue.Dequeue '---取得資料

'            For i = 0 To Allist.Count - 1 '---Table 生成            
'              hst.CopyRow(_info("Y") + _Tablecount, _info("Y") + 1 + _Tablecount)
'            Next

'            '---取得塞Table的位置 new         
'            Dim _locat As New List(Of Seat)
'            Dim _dicseatitem As Dictionary(Of String, String) = ClsConfig.ReadKEYDictionary(_info("TABLE"))
'            For Each _item In _dicseatitem
'              Dim seat_dic = ClsConfig.ReadStringValueDictionary(_info("TABLE"), _item.Key)
'              Dim _point As Seat = New Seat(seat_dic("X"), seat_dic("Y"), seat_dic("TYPE"))

'              _locat.Add(_point)
'            Next

'            '---塞table new
'            For i = 0 To Allist.Count - 1
'              'Dim hr = DirectCast(hst.GetRow(_info("Y") - 1 + i), HSSFRow)
'              For j = 0 To Allist(i).Count - 1
'                Dim hr = DirectCast(hst.GetRow(_locat(j).Y - 1 + i + _Tablecount), HSSFRow)
'                If _locat(j).TYPE = "STRING" Then
'                  hr.Cells(_locat(j).X - 1).SetCellValue(Allist(i)(j))
'                Else
'                  hr.Cells(_locat(j).X - 1).SetCellValue(Convert.ToInt32(Allist(i)(j)))
'                End If
'              Next
'            Next

'            _Tablecount = _Tablecount + Allist.Count

'          End If

'        Next
'      End Using


'      '3.寫回檔案
'      Dim file As New FileStream(sExcelPath, FileMode.Create)
'      '產生檔案
'      wk.Write(file)
'      file.Close()
'      Return True
'    Catch ex As Exception
'      MessageBox.Show(ex.Message)
'      Return False
'    End Try
'  End Function

'  '---for excel print
'  Private Shared Sub printExcel(ByVal sFilePath As String, ByVal evFilePath As String)
'    'For demo
'    'If DisablePrinterFlag = 1 Then Return


'    Dim sFolderPath As String = Application.StartupPath & "\temp"


'    ' 1. 判斷檔案是否存在
'    Dim fi As New FileInfo(sFilePath)
'    If fi.Exists = False Then
'      Return
'    End If
'    'For Each Proc In Process.GetProcessesByName("Excel")
'    '  Proc.Kill()
'    'Next
'    '資料夾存在
'    'If Directory.Exists(sFolderPath) Then
'    '  If CleanFolder(sFolderPath) = False Then 'YES -CLEAN
'    '    Return
'    '  End If
'    'Else
'    '  Directory.CreateDirectory(sFolderPath) 'NO-新增資料夾
'    'End If

'    ''2.將目標excel 拆到指定folder 並依sheet 拆成不同的excel
'    'If FolderInfo(sFilePath, sFolderPath) = False Then Return


'    ' 3. 要列印 EXCEL 檔案的程式位置
'    '    Excel Viewer 在 C:\Program Files\Microsoft Office\Office12\XLVIEW.exe
'    'Dim evFilePath As String = "C:\Program Files (x86)\Microsoft Office\Office12\EXCEL.EXE"


'    ' 4. 初始化 DdeClint 類別物件 ddeClient
'    '    DdeClint(Server 名稱,string topic 名稱)
'    Dim ddeClient As New DdeClient("excel", "system")


'    Dim process__1 As Process = Nothing

'    Do
'      Try
'        ' 5. DDE Client 進行連結
'        If ddeClient.IsConnected = False Then

'          ddeClient.Connect()

'        End If

'      Catch generatedExceptionName As DdeException

'        ' 6. 開啟 Excel Viewer
'        Dim info As New ProcessStartInfo(evFilePath)

'        info.WindowStyle = ProcessWindowStyle.Minimized

'        info.UseShellExecute = True

'        'Excel Viewer used --- 
'        info.Arguments = sFilePath
'        info.Arguments = String.Format("""{0}""", sFilePath)
'        '---

'        process__1 = Process.Start(info)


'        process__1.WaitForInputIdle()

'      End Try
'    Loop While ddeClient.IsConnected = False AndAlso process__1.HasExited = False

'    ' 7. DDE 處理
'    Try


'      'For Each fname As String In System.IO.Directory.GetFileSystemEntries(sFolderPath)
'      ddeClient.Execute(String.Format("[Open(""{0}"")]", sFilePath), 60000)

'      ' 開啟 EXCEL 檔案           
'      ddeClient.Execute("[Print()]", 60000)

'      ' 列印 EXCEL 檔案
'      ddeClient.Execute("[Close()]", 60000)
'      'Next

'      process__1.Kill()

'    Catch ex As Exception
'      ' process__1.Kill()
'      ' MessageBox.Show(ex.Message)
'    End Try

'  End Sub

'  '--for excel mail (會夾帶附件檔,會將寄出去的附件檔改成特定檔名)(mail 前須將所有SMTP設定好)
'  Public Shared Sub StartMail(ByVal _list As List(Of String), ByVal _queue As Queue(Of List(Of List(Of String))), ByVal _NAME As String, ByVal Attachefilename As String)
'    Try
'      Dim Original_samplePath = Application.StartupPath & "\OriginalSAMPLE" '---參考範本來源(必存在)
'      Dim _samplePath = Application.StartupPath & "\SAMPLE" '---print來源 (必存在)

'      '1.檢查folder,original ,sample 必要存在 
'      If CheckFolder(Original_samplePath, _samplePath) Then

'        '2.檢查OriginalSAMPLE資料夾內 是否有要列印的格式
'        For Each fname As String In System.IO.Directory.GetFileSystemEntries(Original_samplePath)
'          If Path.GetFileNameWithoutExtension(fname) = _NAME Then
'            '3.file 複製到sample資料夾內
'            'Dim _destFile = _samplePath & "\" & _NAME & ".xls" '-old
'            Dim _destFile = _samplePath & "\" & Attachefilename & ".xls"
'            File.Copy(fname, _destFile)
'            '---main          
'            If ReadExcel(_destFile, _NAME, _list, _queue) Then

'              If CheckSMTP() = 0 Then '-夾帶mail附件
'                CMailInfo.lstAttachment.Clear()
'                CMailInfo.lstAttachment.Add(_destFile)
'                CMailInfo.SendMail()
'              End If
'            End If
'            Exit For
'          End If
'        Next
'      End If
'    Catch ex As Exception
'      MessageBox.Show(ex.Message)
'    End Try
'  End Sub
'  Private Shared Function CheckSMTP() As Integer
'    Try
'      If CMailInfo.SMTPAddress = "" Or CMailInfo.SMTPPort = 0 Or CMailInfo.SMTPAccount = "" Or CMailInfo.SMTPPassword = "" Or CMailInfo.fromAccount = "" Then
'        Return -1
'      End If
'      Return 0
'    Catch ex As Exception
'      MessageBox.Show("SMTP setting error")
'      Return -1
'    End Try

'  End Function


'  Private Class Seat
'    Public X As String
'    Public Y As String
'    Public TYPE As String

'    Sub New(ByVal _x As String, ByVal _y As String, ByVal _type As String)
'      X = _x
'      Y = _y
'      TYPE = _type
'    End Sub
'  End Class
'  Public Class CMailInfo
'    Private Shared _SMTPAddress As String = ""   '-smtp.gmail.com
'    Private Shared _SMTPPort As Integer = 0      '-587
'    Private Shared _SMTPAccount As String = ""   '-zhumake0208@gmail.com
'    Private Shared _SMTPPassword As String = ""  '-password
'    Private Shared _UseSSL As Boolean = True
'    Public Shared lstRecipient As New List(Of String) '-所有收件者
'    Public Shared lstAttachment As New List(Of String) '-附件檔路徑
'    Private Shared _Message As String = ""
'    Private Shared _MailSubject As String = "" '-主旨

'    Private Shared _fromAccount As String = "" '-zhumake0208@gmail.com

'    Private Shared _fromAlias As String = "" '-zhu


'    '-先不用這個function
'    Public Sub SendMailAsync()
'      Try
'        Dim MailInfo As New MailMessage()
'        'Dim MailInfo As New NET.Mail.MailMessage() '-VITO
'        '發送者
'        MailInfo.From = New MailAddress(fromAccount, fromAlias)
'        'MailInfo.From = New NET.Mail.MailAddress(fromAccount, fromAlias) 'VITO
'        '收件者
'        For Each Item As String In lstRecipient
'          MailInfo.To.Add(Item)
'        Next
'        'myMail.Bcc.Add("main@mail.com")
'        'myMail.CC.Add("main@mail.com")
'        MailInfo.SubjectEncoding = System.Text.Encoding.UTF8
'        MailInfo.Subject = MailSubject '-主旨
'        MailInfo.IsBodyHtml = True
'        MailInfo.BodyEncoding = System.Text.Encoding.UTF8
'        MailInfo.Body = Message '-信件內容
'        'attach file
'        'Dim attachFile As New MemoryStream()
'        'Dim Writer As StreamWriter = New StreamWriter(attachFile)
'        'Writer.WriteLine(Message)
'        'Writer.Flush()
'        Dim attachment As Attachment = Attachment.CreateAttachmentFromString(Message, "Message.txt") '-附件

'        MailInfo.Attachments.Add(attachment)
'        'MailInfo.Attachments.Add(New Net.Mail.Attachment(attachFile, "Message.txt")
'        'SMTP Setting
'        Dim SMTP As New SmtpClient()
'        'Credential
'        SMTP.Credentials = New System.Net.NetworkCredential(SMTPAccount, SMTPPassword)
'        'mySmtp.Port = 587 'gmail
'        SMTP.Port = SMTPPort
'        SMTP.Host = SMTPAddress
'        SMTP.EnableSsl = UseSSL

'        'SMTP.Send(MailInfo)
'        SMTP.SendAsync(MailInfo, MailInfo)
'        AddHandler SMTP.SendCompleted, AddressOf SMTPSendCompleted
'        'Writer.Dispose()
'        'Writer.Close()
'      Catch ex As Exception
'        Console.WriteLine(ex.ToString())
'      End Try
'    End Sub
'    Private Sub SMTPSendCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.AsyncCompletedEventArgs)
'      If e.[Error] IsNot Nothing Then
'        Console.WriteLine(e.[Error].ToString())
'      Else
'        Console.WriteLine("Message sent.")
'      End If
'    End Sub


'    '-直接發送 mail
'    Public Shared Sub SendMail()
'      Try
'        Dim MailInfo As New MailMessage()

'        '發送者
'        MailInfo.From = New MailAddress(fromAccount, fromAlias)

'        '收件者
'        For Each Item As String In lstRecipient
'          MailInfo.To.Add(Item)
'        Next

'        MailInfo.SubjectEncoding = System.Text.Encoding.UTF8
'        MailInfo.Subject = MailSubject
'        MailInfo.IsBodyHtml = True
'        MailInfo.BodyEncoding = System.Text.Encoding.UTF8
'        MailInfo.Body = Message

'        'attach file
'        For Each item In lstAttachment
'          Dim attachment As New Attachment(item)
'          MailInfo.Attachments.Add(attachment)
'        Next
'        'Dim attachment As Attachment = attachment.CreateAttachmentFromString(Message, "Message.txt")
'        'MailInfo.Attachments.Add(attachment)


'        'SMTP Setting
'        Dim SMTP As New SmtpClient()
'        'Credential
'        SMTP.Credentials = New System.Net.NetworkCredential(SMTPAccount, SMTPPassword)
'        'mySmtp.Port = 587 'gmail
'        SMTP.Port = SMTPPort
'        SMTP.Host = SMTPAddress
'        'SMTP.EnableSsl = UseSSL
'        SMTP.EnableSsl = False
'        SMTP.Send(MailInfo)

'      Catch ex As Exception
'        Throw New Exception(ex.ToString())

'      End Try
'    End Sub

'    Public Shared Property fromAlias() As String
'      Get
'        Return _fromAlias
'      End Get
'      Set(ByVal value As String)
'        _fromAlias = value
'      End Set
'    End Property

'    Public Shared Property fromAccount() As String
'      Get
'        Return _fromAccount
'      End Get
'      Set(ByVal value As String)
'        _fromAccount = value
'      End Set
'    End Property


'    Public Shared Property MailSubject() As String
'      Get
'        Return _MailSubject
'      End Get
'      Set(ByVal value As String)
'        _MailSubject = value
'      End Set
'    End Property
'    Public Shared Property Message() As String
'      Get
'        Return _Message
'      End Get
'      Set(ByVal value As String)
'        _Message = value
'      End Set
'    End Property


'    Public Shared Property UseSSL() As Boolean
'      Get
'        Return _UseSSL
'      End Get
'      Set(ByVal value As Boolean)
'        _UseSSL = value
'      End Set
'    End Property


'    Public Shared Property SMTPAddress() As String
'      Get
'        Return _SMTPAddress
'      End Get
'      Set(ByVal value As String)
'        _SMTPAddress = value
'      End Set
'    End Property
'    Public Shared Property SMTPPort() As Integer
'      Get
'        Return _SMTPPort
'      End Get
'      Set(ByVal value As Integer)
'        _SMTPPort = value
'      End Set
'    End Property
'    Public Shared Property SMTPAccount() As String
'      Get
'        Return _SMTPAccount
'      End Get
'      Set(ByVal value As String)
'        _SMTPAccount = value
'      End Set
'    End Property
'    Public Shared Property SMTPPassword() As String
'      Get
'        Return _SMTPPassword
'      End Get
'      Set(ByVal value As String)
'        _SMTPPassword = value
'      End Set
'    End Property

'  End Class



'End Class
