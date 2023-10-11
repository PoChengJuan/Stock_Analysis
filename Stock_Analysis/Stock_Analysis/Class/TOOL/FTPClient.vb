' Description:
'   Created by Vito 
'
'   Purpose:
'     The class is aim to connect to FTP Server with some operations like getfilelist, getfile or putfile.
'     It uses .Net class called WebRequest to implement the Ftp connection.
'     The functions of this class is that you can do either single thread or multithread to operate the ftp connection.
'     By Single thread mode: 
'       You can call the function with a keyword "do" at the beginning (for example: doGetFileList, doPutFile, doGetFile).
'                           
'     By Multithread mode: 
'       You need to register a callback function. 
'       After executing your request, the class will reply the result to you via Callback function
'     
'   Usage:
'     1. Declare the class
'        ex. Dim mFTPClient As CFTPClient.CFTPClient
'     2. Create an instance of class by using "new" with three parameters: server address, username, password
'        ex. mFTPClient = New CFTPClient.CFTPClient("ftp://my.Server.com/abc/ddd/", "Username", "Password")
'     3. Single thread mode: 
'          call doFunction ( doGetFileList, doGetFile, doPutFile) and wait for the reply.
'        Multithread mode:  
'          3.1 register callback function
'             mFTPClient.mfuncGetFileList = New CFTPClient.CFTPClient.ResultGetFileList(AddressOf CBGetFileList)
'             mFTPClient.mfuncPutFile = New CFTPClient.CFTPClient.ResultPutFile(AddressOf CBPutFile)
'             mFTPClient.mfuncGetFile = New CFTPClient.CFTPClient.ResultGetFile(AddressOf CBGetFile)
'             mFTPClient.mfuncDeleteFile = New CFTPClient.CFTPClient.ResultDeleteFile(AddressOf CBDeleteFile)
'          3.2 call operational function ()
'             mFTPClient.PutFile("./LocalFilename.txt", "FTPFilename")
'             mFTPClient.GetFile("FTPFilename", returnFileString)
'             mFTPClient.GetFileList(returnFileListString)
'          3.3 Callback function will be called by the time it completes the request.
'   Enjoy it.   

Imports System
Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Runtime.Remoting.Messaging
'Created by Vito for return the result
Public Class CResult
  Public success As Boolean = False
  Public message As String = ""
  Public functionName As String = ""
  Public tag As String = "" 'can be used to identify which command you asked when using multithreading mode
End Class
'Created by Vito. Class for connecting to FTP Server
Public Class FTPClient

  Public mfuncGetFileList As ResultGetFileList
  Public mfuncPutFile As ResultPutFile
  Public mfuncGetFile As ResultGetFile
  Public mfuncDeleteFile As ResultDeleteFile
  Public Delegate Sub ResultGetFileList(ByVal result As CResult)
  Public Delegate Sub ResultPutFile(ByVal result As CResult)
  Public Delegate Sub ResultGetFile(ByVal result As CResult)
  Public Delegate Sub ResultDeleteFile(ByVal result As CResult)



	Private _mszIP As String = ""
  Private _mszServerURL As String = ""
  Private _mszUserName As String = ""
  Private _mszPassword As String = ""

  Private Delegate Function DelegateGetFileList(ByRef returnString As String, ByVal tag As String) As CResult
  Private Delegate Function DelegatePutFile(ByVal szLocalFilename As String, ByVal szRemoteFilename As String, ByVal tag As String) As CResult
  Private Delegate Function DelegateGetFile(ByVal szFilename As String, ByRef returnString As String, ByVal tag As String) As CResult
  Private Delegate Function DelegateDeleteFile(ByVal szFilename As String, ByVal tag As String) As CResult


  Private Function CallBackGetFileList(ByVal result As IAsyncResult) As Boolean
    Try
      'Console.WriteLine(Now & " CallBackGetFileList")
      Dim returnMessage As String = ""
      Dim resultClass As AsyncResult = CType(result, AsyncResult)
      Dim d As DelegateGetFileList = CType(resultClass.AsyncDelegate, DelegateGetFileList)
      Dim returnStatus As New CResult
      returnStatus = d.EndInvoke(returnMessage, result)
      If mfuncGetFileList Is Nothing Then
        Return True
      Else
        mfuncGetFileList(returnStatus)
      End If
      Return True
    Catch ex As Exception
      'Console.WriteLine(Now & " CallBackGetFileList" & ex.ToString)
      Return False
    End Try

  End Function
  Private Function CallBackPutFile(ByVal result As IAsyncResult) As Boolean
    Try
      'Console.WriteLine(Now & " CallBackPutFile")
      Dim returnMessage As String = ""
      Dim resultClass As AsyncResult = CType(result, AsyncResult)
      Dim d As DelegatePutFile = CType(resultClass.AsyncDelegate, DelegatePutFile)
      Dim returnStatus As New CResult
      returnStatus = d.EndInvoke(result)
      If mfuncPutFile Is Nothing Then
        Return True
      Else
        mfuncPutFile(returnStatus)
      End If
      Return True
    Catch ex As Exception
      'Console.WriteLine(Now & " CallBackPutFile" & ex.ToString)
      Return False
    End Try
  End Function
  Private Function CallBackGetFile(ByVal result As IAsyncResult) As Boolean
    Try
      'Console.WriteLine(Now & " CallBackGetFile")
      Dim returnMessage As String = ""
      Dim resultClass As AsyncResult = CType(result, AsyncResult)
      Dim d As DelegateGetFile = CType(resultClass.AsyncDelegate, DelegateGetFile)
      Dim returnStatus As New CResult
      returnStatus = d.EndInvoke(returnMessage, result)
      If mfuncGetFile Is Nothing Then
        Return True
      Else
        mfuncGetFile(returnStatus)
      End If
      Return True
    Catch ex As Exception
      'Console.WriteLine(Now & " CallBackGetFile" & ex.ToString)
      Return False
    End Try
  End Function
  Private Function CallBackDeleteFile(ByVal result As IAsyncResult) As Boolean
    Try
      'Console.WriteLine(Now & " CallBackDeleteFile")
      Dim returnMessage As String = ""
      Dim resultClass As AsyncResult = CType(result, AsyncResult)
      Dim d As DelegateDeleteFile = CType(resultClass.AsyncDelegate, DelegateDeleteFile)
      Dim returnStatus As New CResult
      returnStatus = d.EndInvoke(result)
      If mfuncDeleteFile Is Nothing Then
        Return True
      Else
        mfuncDeleteFile(returnStatus)
      End If
      Return True
    Catch ex As Exception
      'Console.WriteLine(Now & " CallBackDeleteFile" & ex.ToString)
      Return False
    End Try

  End Function


	Public Property mszIP() As String
		Get
			Return _mszIP
		End Get
		Set(ByVal value As String)
			_mszIP = value
		End Set
	End Property
  Public Property mszServerURL() As String
    Get
      Return _mszServerURL
    End Get
    Set(ByVal value As String)
      If Not value Is Nothing And value.Length > 0 Then
        If value.Trim.EndsWith("/") = False Then
          value = value + "/"
        End If
      End If
      _mszServerURL = value
    End Set
  End Property
  Public Property mszUserName() As String
    Get
      Return _mszUserName
    End Get
    Set(ByVal value As String)
      _mszUserName = value
    End Set
  End Property
  Public Property mszPassword() As String
    Get
      Return _mszPassword
    End Get
    Set(ByVal value As String)
      _mszPassword = value
    End Set
  End Property

  Public Sub New()
    mszServerURL = ""
    mszUserName = ""
    mszPassword = ""

  End Sub
  Public Sub New(ByVal serverURL As String, ByVal userName As String, ByVal password As String)
    mszServerURL = serverURL
    mszUserName = userName
    mszPassword = password
  End Sub
  Public Sub Main()
    mszServerURL = ""
    mszUserName = ""
    mszPassword = ""


  End Sub
  Public Function GetFileList(ByRef returnString As String, Optional ByVal tag As String = "") As Boolean
    Try
      Dim worker As New DelegateGetFileList(AddressOf doGetFileList)
      worker.BeginInvoke(returnString, tag, AddressOf CallBackGetFileList, Nothing)

      Return True
    Catch ex As Exception
      'Console.WriteLine(ex.ToString)
      Return False
    End Try
  End Function
  Public Function PutFile(ByVal szLocalFilename As String, ByVal szRemoteFilename As String, Optional ByVal tag As String = "") As Boolean
    Try
      Dim worker As New DelegatePutFile(AddressOf doPutFile)
      worker.BeginInvoke(szLocalFilename, szRemoteFilename, tag, AddressOf CallBackPutFile, Nothing)
      Return True
    Catch ex As Exception
      'Console.WriteLine(ex.ToString)
      Return False
    End Try

  End Function
  Public Function GetFile(ByVal szFilename As String, ByRef returnString As String, Optional ByVal tag As String = "") As Boolean
    Try
      Dim worker As New DelegateGetFile(AddressOf doGetFile)
      worker.BeginInvoke(szFilename, returnString, tag, AddressOf CallBackGetFile, Nothing)

      Return True

    Catch ex As Exception
      'Console.WriteLine(ex.ToString)
      Return False
    End Try

  End Function
  Public Function DeleteFile(ByVal szFilename As String, Optional ByVal tag As String = "") As Boolean
    Try
      Dim worker As New DelegateDeleteFile(AddressOf doDeleteFile)
      worker.BeginInvoke(szFilename, tag, AddressOf CallBackDeleteFile, Nothing)
      Return True
    Catch ex As Exception
      'Console.WriteLine(ex.ToString)
      Return False
    End Try

  End Function

  Public Function doGetFileList(ByRef returnString As String, Optional ByVal tag As String = "") As CResult
    Dim rtnMessage As New CResult
    rtnMessage.success = False
    rtnMessage.functionName = "getfilelist"
    rtnMessage.tag = tag
    rtnMessage.message = ""
    Dim returnStr As String = ""
    If mszServerURL.Length > 0 And mszPassword.Length > 0 And mszUserName.Length > 0 Then
      Try
        ' create ftp obj.
        Dim request As FtpWebRequest = WebRequest.Create(mszServerURL)
        request.Method = WebRequestMethods.Ftp.ListDirectoryDetails
        ' put the FTP site login information.
        request.Credentials = New NetworkCredential(mszUserName, mszPassword)

        Dim response As FtpWebResponse = request.GetResponse()

        Dim responseStream = response.GetResponseStream()
        Dim reader As StreamReader = New StreamReader(responseStream)
        returnStr = reader.ReadToEnd()
        'Console.WriteLine(returnStr)

        'Console.WriteLine("Directory List Complete, status {0}", response.StatusDescription)

        reader.Close()
        response.Close()
        returnString = returnStr
        rtnMessage.success = True
        rtnMessage.message = returnStr
        Return rtnMessage
      Catch ex As Exception
        'Console.WriteLine("Failed to get file list!!")
        rtnMessage.message = "Failed to get file list!!"
        rtnMessage.success = False
        Return rtnMessage
      End Try
    Else
      'Console.WriteLine("Please set the ServerURL, Username and Password first!!")
      rtnMessage.success = False
      rtnMessage.message = "Please set the ServerURL, Username and Password first!!"
      Return rtnMessage
    End If


  End Function
  Public Function doPutFile(ByVal szLocalFilename As String, ByVal szRemoteFilename As String, Optional ByVal tag As String = "") As CResult
    Dim rtnMessage As New CResult
    rtnMessage.success = False
    rtnMessage.functionName = "putfile"
    rtnMessage.tag = tag
    rtnMessage.message = ""
    Try
      ' create ftp obj.
      Dim request As FtpWebRequest = WebRequest.Create(mszServerURL & szRemoteFilename)
      request.Method = WebRequestMethods.Ftp.UploadFile
      ' put the FTP site login information.
      request.Credentials = New NetworkCredential(mszUserName, mszPassword)
      ' Copy the contents of the file to the request stream.
      Dim sourceStream As StreamReader = New StreamReader(szLocalFilename)
            Dim fileContents As Byte() = Encoding.UTF8.GetBytes(sourceStream.ReadToEnd())
      sourceStream.Close()
      request.ContentLength = fileContents.Length

            Dim requestStream As Stream = request.GetRequestStream()          
			requestStream.Write(fileContents, 0, fileContents.Length)
			requestStream.Close()

			'Dim wr As StreamWriter = New StreamWriter(requestStream)
			'wr.Write(fileContents)
			'wr.Close()

      Dim response As FtpWebResponse = request.GetResponse()



      response.Close()

      rtnMessage.success = True
      rtnMessage.message = "Upload File Complete"
      Return rtnMessage
    Catch ex As Exception
      rtnMessage.success = False
      rtnMessage.message = ex.ToString
      'Console.WriteLine(ex.ToString)
      Return rtnMessage
    End Try
  End Function
  Public Function doGetFile(ByVal szFilename As String, ByRef returnString As String, Optional ByVal tag As String = "") As CResult
    Dim rtnMessage As New CResult
    rtnMessage.success = False
    rtnMessage.functionName = "getfile"
    rtnMessage.tag = tag
    rtnMessage.message = ""
    Dim returnStr As String = ""

    Try
      ' create ftp obj.
      Dim request As FtpWebRequest = WebRequest.Create(mszServerURL & szFilename)
      request.Method = WebRequestMethods.Ftp.DownloadFile

      ' set username and password.
      request.Credentials = New NetworkCredential(mszUserName, mszPassword)

      Dim response As FtpWebResponse = request.GetResponse()
      Dim responseStream As Stream = response.GetResponseStream()
      Dim reader As StreamReader = New StreamReader(responseStream)
      returnStr = reader.ReadToEnd
      'Console.WriteLine(returnStr)
      'Console.WriteLine("Download Complete, status {0}", response.StatusDescription)

      reader.Close()
      response.Close()
      returnString = returnStr
      rtnMessage.success = True
      rtnMessage.message = returnString
      Return rtnMessage
    Catch ex As Exception
      rtnMessage.success = False
      rtnMessage.message = ex.ToString
      'Console.WriteLine(ex.ToString)
      Return rtnMessage
    End Try
  End Function
  Public Function doDeleteFile(ByVal szFilename As String, Optional ByVal tag As String = "") As CResult
    Dim rtnMessage As New CResult
    rtnMessage.success = False
    rtnMessage.functionName = "deletefile"
    rtnMessage.tag = tag
    rtnMessage.message = ""
    Dim returnStr As String = ""

    Try
      ' Get the object used to communicate with the server.
      Dim request As FtpWebRequest = WebRequest.Create(mszServerURL & szFilename)
      request.Method = WebRequestMethods.Ftp.DeleteFile

      'set username and password
      request.Credentials = New NetworkCredential(mszUserName, mszPassword)

      Dim response As FtpWebResponse = request.GetResponse()
      Dim responseStream As Stream = response.GetResponseStream()
      Dim reader As StreamReader = New StreamReader(responseStream)
      returnStr = reader.ReadToEnd
      'Console.WriteLine(returnStr)
      'Console.WriteLine("DeleteFile Complete, status {0}", response.StatusDescription)

      reader.Close()
      response.Close()
      rtnMessage.success = True
      rtnMessage.message = "DeleteFile Complete"
      Return rtnMessage
    Catch ex As Exception
      rtnMessage.success = False
      rtnMessage.message = ex.ToString()
      'Console.WriteLine(ex.ToString)
      Return rtnMessage
    End Try
  End Function


  'Public Function Download(ByVal sourceFilename As String, ByVal localFilename As String, Optional ByVal PermitOverwrite As Boolean = False) As Boolean

  '    '2. determine target file   

  '    Dim fi As New FileInfo(localFilename)

  '    Return Me.Download(sourceFilename, fi, PermitOverwrite)

  'End Function



  ''Version taking an FtpFileInfo   

  'Public Function Download(ByVal file As FTPfileInfo, ByVal localFilename As String, Optional ByVal PermitOverwrite As Boolean = False) As Boolean

  '    Return Me.Download(file.FullName, localFilename, PermitOverwrite)

  'End Function


  ''Another version taking FtpFileInfo and FileInfo   

  'Public Function Download(ByVal file As FTPfileInfo, ByVal localFI As FileInfo, Optional ByVal PermitOverwrite As Boolean = False) As Boolean

  '    Return Me.Download(file.FullName, localFI, PermitOverwrite)

  'End Function



  ''Version taking string/FileInfo   

  'Public Function Download(ByVal sourceFilename As String, ByVal targetFI As FileInfo, Optional ByVal PermitOverwrite As Boolean = False) As Boolean

  '    '1. check target   

  '    If targetFI.Exists And Not (PermitOverwrite) Then Throw New ApplicationException("Target file already exists")

  '    '2. check source   
  '    Dim target As String

  '    If sourceFilename.Trim = "" Then

  '        Throw New ApplicationException("File not specified")

  '    ElseIf sourceFilename.Contains("/") Then

  '        'treat as a full path   

  '        target = AdjustDir(sourceFilename)

  '    Else

  '        'treat as filename only, use current directory   

  '        target = CurrentDirectory & sourceFilename

  '    End If


  '    Dim URI As String = Hostname & target

  '    '3. perform copy   

  '    Dim ftp As Net.FtpWebRequest = GetRequest(URI)

  '    'Set request to download a file in binary mode   

  '    ftp.Method = Net.WebRequestMethods.Ftp.DownloadFile
  '    ftp.UseBinary = True

  '    'open request and get response stream   

  '    Using response As FtpWebResponse = CType(ftp.GetResponse, FtpWebResponse)

  '        Using responseStream As Stream = response.GetResponseStream

  '            'loop to read & write to file   

  '            Using fs As FileStream = targetFI.OpenWrite

  '                Try

  '                    Dim buffer(2047) As Byte

  '                    Dim read As Integer = 0

  '                    Do

  '                        read = responseStream.Read(Buffer, 0, Buffer.Length)

  '                        fs.Write(Buffer, 0, read)

  '                    Loop Until read = 0

  '                    responseStream.Close()

  '                    fs.Flush()

  '                    fs.Close()

  '                Catch ex As Exception

  '                    'catch error and delete file only partially downloaded   

  '                    fs.Close()

  '                    'delete target file as it's incomplete   

  '                    targetFI.Delete()

  '                    Throw

  '                End Try

  '            End Using

  '            responseStream.Close()

  '        End Using

  '        response.Close()

  '    End Using



    'End Function
    Public Function Download(ByVal strFTPIP As String, ByVal strUserID As String, ByVal strPassword As String, ByVal filePath As String, ByVal fileName As String) As Boolean
        'FTPSettings.IP = strFTPIP
        'FTPSettings.UserID = strUserID
        'FTPSettings.Password = strPassword
        Dim reqFTP As FtpWebRequest = Nothing
        Dim ftpStream As Stream = Nothing
        Try
            Dim outputStream As New FileStream(filePath + "\" + fileName, FileMode.Create)
            reqFTP = DirectCast(FtpWebRequest.Create(New Uri(mszServerURL + "/" + fileName)), FtpWebRequest)
            reqFTP.Method = WebRequestMethods.Ftp.DownloadFile
            reqFTP.UseBinary = True
            reqFTP.Credentials = New NetworkCredential(mszUserName, mszPassword)
            Dim response As FtpWebResponse = DirectCast(reqFTP.GetResponse(), FtpWebResponse)
            ftpStream = response.GetResponseStream()
            Dim cl As Long = response.ContentLength
            Dim bufferSize As Integer = 2048
            Dim readCount As Integer
            Dim buffer As Byte() = New Byte(bufferSize - 1) {}

            readCount = ftpStream.Read(buffer, 0, bufferSize)
            While readCount > 0
                outputStream.Write(buffer, 0, readCount)
                readCount = ftpStream.Read(buffer, 0, bufferSize)
            End While

            ftpStream.Close()
            outputStream.Close()
            response.Close()
            Return True
        Catch ex As Exception
            If ftpStream IsNot Nothing Then
                ftpStream.Close()
                ftpStream.Dispose()
            End If
            Throw New Exception(ex.Message.ToString())
        End Try
        Return False
    End Function
    'Public Function Download(ByVal strFTPIP As String, ByVal strUserID As String, ByVal strPassword As String, ByVal filePath As String, ByVal fileName As String) As Boolean
    '  FTPSettings.IP = strFTPIP
    '  FTPSettings.UserID = strUserID
    '  FTPSettings.Password = strPassword
    '  Dim reqFTP As FtpWebRequest = Nothing
    '  Dim ftpStream As Stream = Nothing

    '  Try
    '    Dim outputStream As New FileStream(filePath + "\" + fileName, FileMode.Create)
    '    reqFTP = DirectCast(FtpWebRequest.Create(New Uri(FTPSettings.IP + "/" + fileName)), FtpWebRequest)
    '    reqFTP.Method = WebRequestMethods.Ftp.DownloadFile
    '    reqFTP.UseBinary = True
    '    reqFTP.Credentials = New NetworkCredential(FTPSettings.UserID, FTPSettings.Password)
    '    Dim response As FtpWebResponse = DirectCast(reqFTP.GetResponse(), FtpWebResponse)
    '    ftpStream = response.GetResponseStream()
    '    Dim cl As Long = response.ContentLength
    '    Dim bufferSize As Integer = 2048
    '    Dim readCount As Integer
    '    Dim buffer As Byte() = New Byte(bufferSize - 1) {}

    '    readCount = ftpStream.Read(buffer, 0, bufferSize)
    '    While readCount > 0
    '      outputStream.Write(buffer, 0, readCount)
    '      readCount = ftpStream.Read(buffer, 0, bufferSize)
    '    End While

    '    ftpStream.Close()
    '    outputStream.Close()
    '    response.Close()
    '    Return True
    '  Catch ex As Exception
    '    If ftpStream IsNot Nothing Then
    '      ftpStream.Close()
    '      ftpStream.Dispose()
    '    End If
    '    Throw New Exception(ex.Message.ToString())
    '  End Try
    '  Return False
    'End Function

  Public NotInheritable Class FTPSettings
    Private Sub New()
    End Sub
    Public Shared Property IP() As String
      Get
        Return m_IP
      End Get
      Set(ByVal value As String)
        m_IP = value

      End Set
    End Property
    Private Shared m_IP As String
    Public Shared Property UserID() As String
      Get
        Return m_UserID
      End Get
      Set(ByVal value As String)
        m_UserID = value
      End Set
    End Property
    Private Shared m_UserID As String
    Public Shared Property Password() As String
      Get
        Return m_Password
      End Get
      Set(ByVal value As String)
        m_Password = value
      End Set
    End Property
    Private Shared m_Password As String

  End Class


End Class

