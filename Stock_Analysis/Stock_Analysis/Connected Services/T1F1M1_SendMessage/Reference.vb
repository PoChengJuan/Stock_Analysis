﻿'------------------------------------------------------------------------------
' <auto-generated>
'     這段程式碼是由工具產生的。
'     執行階段版本:4.0.30319.42000
'
'     對這個檔案所做的變更可能會造成錯誤的行為，而且如果重新產生程式碼，
'     變更將會遺失。
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On


Namespace T1F1M1_SendMessage
    
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "4.0.0.0"),  _
     System.ServiceModel.ServiceContractAttribute(ConfigurationName:="T1F1M1_SendMessage.IWMS_WCFService")>  _
    Public Interface IWMS_WCFService
        
        <System.ServiceModel.OperationContractAttribute(Action:="http://tempuri.org/IWMS_WCFService/T1F1M1_SendMessage", ReplyAction:="http://tempuri.org/IWMS_WCFService/T1F1M1_SendMessageResponse"),  _
         System.ServiceModel.DataContractFormatAttribute(Style:=System.ServiceModel.OperationFormatStyle.Rpc)>  _
        Function T1F1M1_SendMessage(ByVal xml As String) As String
        
        <System.ServiceModel.OperationContractAttribute(Action:="http://tempuri.org/IWMS_WCFService/T1F1M1_SendMessage", ReplyAction:="http://tempuri.org/IWMS_WCFService/T1F1M1_SendMessageResponse")>  _
        Function T1F1M1_SendMessageAsync(ByVal xml As String) As System.Threading.Tasks.Task(Of String)
    End Interface
    
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "4.0.0.0")>  _
    Public Interface IWMS_WCFServiceChannel
        Inherits T1F1M1_SendMessage.IWMS_WCFService, System.ServiceModel.IClientChannel
    End Interface
    
    <System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "4.0.0.0")>  _
    Partial Public Class WMS_WCFServiceClient
        Inherits System.ServiceModel.ClientBase(Of T1F1M1_SendMessage.IWMS_WCFService)
        Implements T1F1M1_SendMessage.IWMS_WCFService
        
        Public Sub New()
            MyBase.New
        End Sub
        
        Public Sub New(ByVal endpointConfigurationName As String)
            MyBase.New(endpointConfigurationName)
        End Sub
        
        Public Sub New(ByVal endpointConfigurationName As String, ByVal remoteAddress As String)
            MyBase.New(endpointConfigurationName, remoteAddress)
        End Sub
        
        Public Sub New(ByVal endpointConfigurationName As String, ByVal remoteAddress As System.ServiceModel.EndpointAddress)
            MyBase.New(endpointConfigurationName, remoteAddress)
        End Sub
        
        Public Sub New(ByVal binding As System.ServiceModel.Channels.Binding, ByVal remoteAddress As System.ServiceModel.EndpointAddress)
            MyBase.New(binding, remoteAddress)
        End Sub
        
        Public Function T1F1M1_SendMessage(ByVal xml As String) As String Implements T1F1M1_SendMessage.IWMS_WCFService.T1F1M1_SendMessage
            Return MyBase.Channel.T1F1M1_SendMessage(xml)
        End Function
        
        Public Function T1F1M1_SendMessageAsync(ByVal xml As String) As System.Threading.Tasks.Task(Of String) Implements T1F1M1_SendMessage.IWMS_WCFService.T1F1M1_SendMessageAsync
            Return MyBase.Channel.T1F1M1_SendMessageAsync(xml)
        End Function
    End Class
End Namespace
