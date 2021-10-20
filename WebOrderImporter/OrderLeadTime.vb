
''' <summary>
''' This class is intended to hold the Product Lead Time data for the product
''' with the longest lead time in an order (thereby defining the lead time of
''' the entire order).
''' </summary>

Public Class OrderLeadTime

    Private class_tsCutoffTime As TimeSpan
    Private class_intLeadTimeDays As Integer

    Public Sub New(intLeadTimeDays, tsCutoffTime)
        class_intLeadTimeDays = intLeadTimeDays
        class_tsCutoffTime = tsCutoffTime
    End Sub

    ''' <summary>
    ''' Represents the cutoff time for a product's lead time days. Past this time,
    ''' the website will use tomorrow's date as a basis for calculating the delivery
    ''' date rather than today's.
    ''' </summary>
    ''' <returns></returns>
    Public Property CutoffTime As TimeSpan
        Get
            Return class_tsCutoffTime
        End Get
        Set(value As TimeSpan)
            class_tsCutoffTime = value
        End Set
    End Property

    ''' <summary>
    ''' Number of days to delivery
    ''' </summary>
    ''' <returns></returns>
    Public Property LeadTimeDays As Integer
        Get
            Return class_intLeadTimeDays
        End Get
        Set(value As Integer)
            class_intLeadTimeDays = value
        End Set
    End Property

End Class
