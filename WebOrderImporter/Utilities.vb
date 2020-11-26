Module Utilities

    ''' <summary>
    ''' Adds the given number of working days to the supplied date
    ''' </summary>
    ''' <param name="dateInput">Date to add working days to</param>
    ''' <param name="intWorkingDays">Number of working days to add</param>
    ''' <remarks>Author: Huw Day - 05/02/2018</remarks>
    ''' <returns>Date</returns>
    Function AddWorkingDays(ByVal dateInput As Date, intWorkingDays As Integer) As Date

        Dim dateOutput As Date
        Dim intCounter As Integer = 0

        If intWorkingDays < 0 Then
            ' Number of days to add is a negative number. Run the subtractor instead
            dateOutput = Utilities.SubtractWorkingDays(dateInput, intWorkingDays * -1)

        Else
            ' Number of days to add is a positive number. Perform the addition operation
            dateOutput = dateInput
            Do While intCounter < intWorkingDays
                dateOutput = dateOutput.AddDays(1)
                If dateOutput.DayOfWeek = DayOfWeek.Saturday Or dateOutput.DayOfWeek = DayOfWeek.Sunday Then
                    ' Don't increment the counter.
                Else
                    intCounter = intCounter + 1
                End If
            Loop
        End If


        Return dateOutput

    End Function

    ''' <summary>
    ''' Subtracts the given number of working days from the supplied date
    ''' </summary>
    ''' <param name="dateInput">Date to add working days to</param>
    ''' <param name="intWorkingDays">Number of working days to add</param>
    ''' <remarks>Author: Huw Day - 05/02/2018</remarks>
    ''' <returns>Date</returns>
    Function SubtractWorkingDays(dateInput As Date, intWorkingDays As Integer) As Date

        Dim dateOutput As Date = dateInput
        Dim intWorkingDaysRemoved As Integer = 0

        If intWorkingDays < 0 Then
            ' Someone has asked to subtract a negative number of days. Run the AddWorkingDays function instead
            dateOutput = Utilities.AddWorkingDays(dateInput, intWorkingDays * -1)

        Else
            ' Run the subtract working days algorithm
            Do While intWorkingDaysRemoved < intWorkingDays
                dateOutput = dateOutput.AddDays(-1)
                If dateOutput.DayOfWeek = DayOfWeek.Saturday Or dateOutput.DayOfWeek = DayOfWeek.Sunday Then
                    ' Don't increment the counter
                Else
                    intWorkingDaysRemoved = intWorkingDaysRemoved + 1
                End If
            Loop

        End If

        Return dateOutput

    End Function

End Module
