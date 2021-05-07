Public Class ResultBoolean

    Dim class_blnResult As Boolean
    Dim class_strMessage As String

    Public Sub New(blnResult As Boolean)
        Me.New(blnResult, "")
    End Sub

    Public Sub New(blnResult As Boolean, strMessage As String)

        class_blnResult = blnResult
        class_strMessage = strMessage

    End Sub

    Public ReadOnly Property Result As Boolean
        Get
            Return class_blnResult
        End Get
    End Property

    Public ReadOnly Property Message As String
        Get
            Return class_strMessage
        End Get
    End Property

End Class
