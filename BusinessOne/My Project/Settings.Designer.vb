﻿'------------------------------------------------------------------------------
' <auto-generated>
'     This code was generated by a tool.
'     Runtime Version:4.0.30319.42000
'
'     Changes to this file may cause incorrect behavior and will be lost if
'     the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On


Namespace My
    
    <Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute(),  _
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "12.0.0.0"),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
    Partial Friend NotInheritable Class MySettings
        Inherits Global.System.Configuration.ApplicationSettingsBase
        
        Private Shared defaultInstance As MySettings = CType(Global.System.Configuration.ApplicationSettingsBase.Synchronized(New MySettings()),MySettings)
        
#Region "My.Settings Auto-Save Functionality"
#If _MyType = "WindowsForms" Then
    Private Shared addedHandler As Boolean

    Private Shared addedHandlerLockObject As New Object

    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(), Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)> _
    Private Shared Sub AutoSaveSettings(ByVal sender As Global.System.Object, ByVal e As Global.System.EventArgs)
        If My.Application.SaveMySettingsOnExit Then
            My.Settings.Save()
        End If
    End Sub
#End If
#End Region
        
        Public Shared ReadOnly Property [Default]() As MySettings
            Get
                
#If _MyType = "WindowsForms" Then
               If Not addedHandler Then
                    SyncLock addedHandlerLockObject
                        If Not addedHandler Then
                            AddHandler My.Application.Shutdown, AddressOf AutoSaveSettings
                            addedHandler = True
                        End If
                    End SyncLock
                End If
#End If
                Return defaultInstance
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.SpecialSettingAttribute(Global.System.Configuration.SpecialSetting.ConnectionString),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Initial Catalog=GSATaiwan_test;server=LIL161NT;")>  _
        Public ReadOnly Property MyConnString() As String
            Get
                Return CType(Me("MyConnString"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("\\172.22.10.77\sharedfolder\BOne\TW\")>  _
        Public ReadOnly Property TWAutoReport() As String
            Get
                Return CType(Me("TWAutoReport"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("\\172.22.10.77\BOneHK\")>  _
        Public ReadOnly Property HKAutoReport() As String
            Get
                Return CType(Me("HKAutoReport"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("ODBC;DSN=PostgreSQLhon03nt;")>  _
        Public ReadOnly Property oExcon() As String
            Get
                Return CType(Me("oExcon"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("\\172.22.20.8\Users\SAP_B1_SALES_REPORT\")>  _
        Public ReadOnly Property TWAutoReportOld() As String
            Get
                Return CType(Me("TWAutoReportOld"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("\\172.22.10.77\sharedfolder\BOne\HK\")>  _
        Public ReadOnly Property HKAutoReportOld() As String
            Get
                Return CType(Me("HKAutoReportOld"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("host=localhost;port=5433;database=LogisticDb20200406;CommandTimeout=1000;TimeOut="& _ 
            "1000;")>  _
        Public ReadOnly Property PostgreSQLDev() As String
            Get
                Return CType(Me("PostgreSQLDev"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("host=hon14nt;port=5432;database=LogisticDb;CommandTimeout=1000;TimeOut=1000;")>  _
        Public ReadOnly Property PostgreSQLLive() As String
            Get
                Return CType(Me("PostgreSQLLive"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.SpecialSettingAttribute(Global.System.Configuration.SpecialSetting.ConnectionString),  _
         Global.System.Configuration.DefaultSettingValueAttribute("host=hon14nt;port=5432;database=LogisticDb;CommandTimeout=1000;TimeOut=1000;")>  _
        Public ReadOnly Property PostgreSQLCon() As String
            Get
                Return CType(Me("PostgreSQLCon"),String)
            End Get
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("SW04R149")>  _
        Public Property host() As String
            Get
                Return CType(Me("host"),String)
            End Get
            Set
                Me("host") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("sa")>  _
        Public Property user() As String
            Get
                Return CType(Me("user"),String)
            End Get
            Set
                Me("user") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("SB04eVer")>  _
        Public Property pwd() As String
            Get
                Return CType(Me("pwd"),String)
            End Get
            Set
                Me("pwd") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("b1ifce")>  _
        Public Property userinterface() As String
            Get
                Return CType(Me("userinterface"),String)
            End Get
            Set
                Me("userinterface") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("B104int")>  _
        Public Property pwdinterface() As String
            Get
                Return CType(Me("pwdinterface"),String)
            End Get
            Set
                Me("pwdinterface") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("HON19NT")>  _
        Public Property hostold() As String
            Get
                Return CType(Me("hostold"),String)
            End Get
            Set
                Me("hostold") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("sa")>  _
        Public Property userold() As String
            Get
                Return CType(Me("userold"),String)
            End Get
            Set
                Me("userold") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("SB19eVer")>  _
        Public Property pwdold() As String
            Get
                Return CType(Me("pwdold"),String)
            End Get
            Set
                Me("pwdold") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("SW04R149")>  _
        Public Property hostnew() As String
            Get
                Return CType(Me("hostnew"),String)
            End Get
            Set
                Me("hostnew") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("sa")>  _
        Public Property usernew() As String
            Get
                Return CType(Me("usernew"),String)
            End Get
            Set
                Me("usernew") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("SB04eVer")>  _
        Public Property pwdnew() As String
            Get
                Return CType(Me("pwdnew"),String)
            End Get
            Set
                Me("pwdnew") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("b1ifce")>  _
        Public Property userinterfacenew() As String
            Get
                Return CType(Me("userinterfacenew"),String)
            End Get
            Set
                Me("userinterfacenew") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("B104int")>  _
        Public Property pwdinterfacenew() As String
            Get
                Return CType(Me("pwdinterfacenew"),String)
            End Get
            Set
                Me("pwdinterfacenew") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("smtp.seb.com")>  _
        Public Property smtpclient() As String
            Get
                Return CType(Me("smtpclient"),String)
            End Get
            Set
                Me("smtpclient") = value
            End Set
        End Property
    End Class
End Namespace

Namespace My
    
    <Global.Microsoft.VisualBasic.HideModuleNameAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute()>  _
    Friend Module MySettingsProperty
        
        <Global.System.ComponentModel.Design.HelpKeywordAttribute("My.Settings")>  _
        Friend ReadOnly Property Settings() As Global.BusinessOne.My.MySettings
            Get
                Return Global.BusinessOne.My.MySettings.Default
            End Get
        End Property
    End Module
End Namespace
