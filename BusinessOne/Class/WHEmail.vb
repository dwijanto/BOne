Imports System.Text
Imports System.Net.Mail
Imports System.Net.Mime
Public Class WHEmail
    Inherits Email

    Private sendtoname As String
    Private drv As DataRowView
    Public _errorMessage As String = String.Empty
    Private statusname As String = String.Empty
    Private dtlbs As BindingSource
    Private Function GetBodyMessage() As String
        Dim sb As New StringBuilder

        sb.Append("<!DOCTYPE html><html><head><meta name='description' content='[ProductRequest]' /><meta http-equiv='Content-Type' content='text/html; charset=us-ascii'></head>" &
                  "<style>td,th {padding-left:5px;padding-right:10px;  }  th {background-color:red;    color:white;text-align:left;}  .defaultfont{    font-size:11.0pt; font-family:'Calibri','sans-serif';    }</style><body class='defaultfont'>")
        sb.Append(String.Format("<p>Dear {0},</p> <p>Attached please find the daily SEB vs DSV inventory.</p>", sendtoname))
       

        sb.Append(String.Format("<p>Thank you.</p></body>"))

        Return sb.ToString
    End Function

    Public ReadOnly Property ErrorMessage As String
        Get
            Return _errorMessage
        End Get
    End Property


    Public Function Execute(ByVal sendto As String, ByVal sendtoname As String, Attachmentlist As List(Of String), Optional ByVal cc As String = "") As Boolean
        Dim myret As Boolean = False
        Try
            'Prepare Email
            'Me.statusname = statusname
            Me.sendtoname = sendtoname
            'Me.drv = drv
            Me.sendto = Trim(sendto)
            Me.subject = String.Format("Daily SEB vs DSV Inventory. ({0:dd-MMM-yyyy}).", Today.Date)
            'Me.dtlbs = dtlbs

            If Not IsNothing(Me.sendto) Then

                Dim mycontent = GetBodyMessage()

                'Dim htmlView As AlternateView = AlternateView.CreateAlternateViewFromString(String.Format("{0} <br>Or click the Product Request Application icon on your desktop: <br><p> <img src=cid:myLogo> <br></p><p>Produt Request System Administrator</p></body></html>", mycontent), Nothing, MediaTypeNames.Text.Html)

                'Dim logo As New LinkedResource(Application.StartupPath & "\PR.png")
                'logo.ContentId = "myLogo"
                'htmlView.LinkedResources.Add(logo)

                'Me.htmlView = htmlView
                Me.isBodyHtml = True
                Me.sender = "no-reply@groupeseb.com"
                Me.body = mycontent
                Me.cc = String.Format("{0}", cc)
                Me.attachmentlist = Attachmentlist
                myret = Me.send(ErrorMessage)
            End If
        Catch ex As Exception
            Logger.log(ex.Message)
            MessageBox.Show(ex.Message)
        End Try

        Return myret
    End Function
End Class
