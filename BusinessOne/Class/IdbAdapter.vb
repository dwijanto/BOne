
Public Interface IdbAdapter
    Property BS As BindingSource
    'Property DS As DataSet

    Function LoadData() As Boolean
    Function Save(ByVal mye As TxBaseEventArgs) As Boolean
    Function Save() As Boolean

End Interface