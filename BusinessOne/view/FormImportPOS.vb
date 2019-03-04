Imports System.Threading
Imports System.Text

Public Class FormImportPOS
    Dim dbAdapter1 = PostgreSQLDBAdapter.getInstance
    Dim FolderBrowserDialog1 As New System.Windows.Forms.FolderBrowserDialog
    Dim mythread As New Thread(AddressOf doWork)
    Dim openfiledialog1 As New OpenFileDialog
    Dim mySelectedPath As String
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Start Thread
        If Not mythread.IsAlive Then
            'Get file
            'If openfiledialog1.ShowDialog = DialogResult.OK Then
            '    mythread = New Thread(AddressOf doWork)
            '    mythread.Start()
            'End If
            With FolderBrowserDialog1
                .RootFolder = Environment.SpecialFolder.Desktop
                .SelectedPath = "c:\"
                .Description = "Select the source directory"
                If .ShowDialog = Windows.Forms.DialogResult.OK Then
                    mySelectedPath = .SelectedPath

                    Try
                        mythread = New Thread(AddressOf doWork)
                        mythread.Start()
                    Catch ex As Exception
                        MsgBox(ex.Message)
                    End Try
                End If
            End With
        Else
            MessageBox.Show("Process still running. Please Wait!")
        End If
    End Sub

    Private Sub doWork()
        Dim mystr As New StringBuilder
        Dim myInsert As New System.Text.StringBuilder
        Dim myrecord() As String
        Dim myTextFile As String = String.Empty


        Dim dir As New IO.DirectoryInfo(mySelectedPath)
        Dim arrFI As IO.FileInfo() = dir.GetFiles("*.txt")

        For Each fi As IO.FileInfo In arrFI
            myTextFile = fi.FullName
            ProgressReport(2, String.Format("Read Text File...{0}", fi.FullName))

            Using objTFParser = New FileIO.TextFieldParser(fi.FullName)
                With objTFParser
                    .TextFieldType = FileIO.FieldType.Delimited
                    .SetDelimiters(",")

                    .HasFieldsEnclosedInQuotes = True
                    Dim count As Long = 0

                    Do Until .EndOfData
                        myrecord = .ReadFields

                        myInsert.Append(fi.Name & vbTab &
                                        myrecord(0) & vbTab &
                                        myrecord(1) & vbTab &
                                        validdateyyyymmdd(myrecord(2)) & vbTab &
                                        myrecord(3) & vbTab &
                                        myrecord(4) & vbTab &
                                        myrecord(5) & vbTab &
                                        myrecord(6) & vbTab &
                                        myrecord(7) & vbTab &
                                        myrecord(8) & vbTab &
                                        myrecord(9) & vbTab &
                                        myrecord(10) & vbTab &
                                        myrecord(11) & vbTab &
                                        myrecord(12) & vbTab &
                                        myrecord(13) & vbCrLf 
                                        )
                        'If count > 0 Then
                        'myInsert.Append(myrecord(0) & vbTab &
                        '                myrecord(2) & vbTab &
                        '                validstr(myrecord(4)) & vbTab &
                        '                validstr(myrecord(5)) & vbTab &
                        '                validstr(myrecord(6)) & vbTab &
                        '                validdate(myrecord(7)) & vbTab &
                        '                validstr(myrecord(8)) & vbTab &
                        '                validdate(myrecord(9)) & vbCrLf)
                        'End If
                        ' count += 1
                    Loop
                End With
            End Using
        Next



        
        'update record
        If myInsert.Length > 0 Then
            ProgressReport(1, "Start Add New Records")
            mystr.Append("delete from bone.posdata;")
            mystr.Append("select setval('bone.posdata_id_seq',1,false);")
            Dim sqlstr As String = "copy bone.posdata(filename,counter,loc,txdate,lineno,productid,cmmf,data1,data2,data3,data4,data5,data6,data7,data8) from stdin with null as 'Null';"
            'mystr.Append(sqlstr)
            Dim ra As Long = 0
            Dim errmessage As String = String.Empty
            Dim myret As Boolean = False
            'If Not DbAdapter1.ExecuteNonQuery(sqlstr, ra, errmessage) Then
            '    MessageBox.Show(errmessage)
            'Else
            '    ProgressReport(1, "Update Done.")
            'End If
            Try
                ra = dbAdapter1.ExecuteNonQuery(mystr.ToString)
                errmessage = dbAdapter1.copy(sqlstr, myInsert.ToString, myret)
                If myret Then
                    ProgressReport(1, "Add Records Done.")
                Else
                    ProgressReport(1, errmessage)
                End If
            Catch ex As Exception
                ProgressReport(1, ex.Message)
            End Try
        End If
    End Sub
    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    Me.ToolStripStatusLabel1.Text = message
            End Select

        End If

    End Sub
    Private Function validstr(ByVal data As Object) As Object
        If IsDBNull(data) Then
            Return "Null"
        ElseIf data = "" Then
            Return "Null"
        End If
        Return data
    End Function

    Private Sub StatusStrip1_ItemClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles StatusStrip1.ItemClicked

    End Sub

    Private Function validdate(ByVal myrecord As String) As Object
        If myrecord = "" Then
            Return "Null"
        Else
            Return String.Format("'{0:yyyy-MM-dd}'", CDate(myrecord))
        End If
    End Function

    Private Function validdateyyyymmdd(ByVal myrecord As String) As Object
        If myrecord = "" Then
            Return "Null"
        Else

            Return String.Format("'{0}-{1}-{2}'", myrecord.Substring(0, 4), myrecord.Substring(4, 2), myrecord.Substring(6, 2))
        End If
    End Function
End Class