Public Class ItemPriceTMPController
    Implements IController
    Implements IToolbarAction

    Public Model As New ItemPriceTMPModel
    Public BS As BindingSource
    Public DS As DataSet
    Public errorMsg As String

    Public ReadOnly Property GetTable As DataTable Implements IController.GetTable
        Get
            Return DS.Tables(Model.tablename).Copy()
        End Get
    End Property

    Public Function loaddata() As Boolean Implements IController.loaddata
        Dim myret As Boolean = False
        Model = New ItemPriceTMPModel
        DS = New DataSet
        Try
            If Model.LoadData(DS) Then
                Dim pk(0) As DataColumn
                pk(0) = DS.Tables(0).Columns("itemcode")
                DS.Tables(0).PrimaryKey = pk
                BS = New BindingSource
                BS.DataSource = DS.Tables(0)
                myret = True
            End If
        Catch ex As Exception
            errorMsg = ex.Message
        End Try        
        Return myret
    End Function

    Public Function DeleteAll() As Boolean
        Return Model.DeleteAll()
    End Function

    Public Function DeleteItemPrice() As Boolean
        Return Model.DeleteItemPrice
    End Function

    Public Function save(ds As DataSet) As Boolean
        Dim myret As Boolean = False

        If Not IsNothing(ds) Then
            Dim mymessage As String = String.Empty
            Dim ra As Integer
            Dim mye As New ContentBaseEventArgs(ds, True, mymessage, ra, True)
            Try
                If save(mye) Then
                    ds.Merge(ds)
                    ds.AcceptChanges()                   
                    myret = True
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                ds.Merge(ds)
            End Try
        End If

        Return myret
    End Function

    Public Function saveItemPrice(myds As DataSet) As Boolean
        Dim myret As Boolean = False
        Dim ds2 As DataSet = myds.GetChanges
        If Not IsNothing(ds2) Then
            Dim mymessage As String = String.Empty
            Dim ra As Integer
            Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
            Try
                If saveItemPrice(mye) Then
                    DS.Merge(ds2)
                    'ds.AcceptChanges()
                    myret = True
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                DS.Merge(ds2)
            End Try
        End If

        Return myret
    End Function

    Public Function save() As Boolean Implements IController.save
        Dim myret As Boolean = False
        BS.EndEdit()

        Dim ds2 As DataSet = DS.GetChanges
        If Not IsNothing(ds2) Then
            Dim mymessage As String = String.Empty
            Dim ra As Integer
            Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
            Try
                If save(mye) Then
                    DS.Merge(ds2)
                    DS.AcceptChanges()
                    MessageBox.Show("Saved.")
                    myret = True
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                DS.Merge(ds2)
            End Try
        End If

        Return myret
    End Function
    Public Function Save(ByVal mye As ContentBaseEventArgs) As Boolean Implements IToolbarAction.Save
        Dim myret As Boolean = False
        If Model.save(Me, mye) Then
            myret = True
        End If
        Return myret
    End Function
    Public Property ApplyFilter As String Implements IToolbarAction.ApplyFilter
        Get
            Return BS.Filter
        End Get
        Set(ByVal value As String)
            BS.Filter = String.Format(Model.FilterField, value)
        End Set
    End Property

    Public Function GetCurrentRecord() As DataRowView Implements IToolbarAction.GetCurrentRecord
        Return Nothing
    End Function

    Public Function GetNewRecord() As DataRowView Implements IToolbarAction.GetNewRecord
        Return Nothing
    End Function

    Public Sub RemoveAt(value As Integer) Implements IToolbarAction.RemoveAt

    End Sub

    Private Function saveItemPrice(mye As ContentBaseEventArgs) As Boolean
        Dim myret As Boolean = False
        If Model.saveStaffItemPrice(Me, mye) Then
            myret = True
        End If
        Return myret
    End Function

    Public Function Rollback() As Boolean
        Dim myret As Boolean = False
        If Model.RollBack Then
            myret = True
        End If
        Return myret
    End Function

  
End Class
