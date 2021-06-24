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


    Function WorkingDaysDifference(date1 As Date, date2 As Date) As Integer

        ' --- Defining objects and variables ---
        Dim blnDifferenceIsPositive As Boolean
        Dim dateTemp As Date
        Dim intOutput As Integer = 0
        Dim intSafetyCounter As Integer = 0


        ' --- Switching dates, if required ---
        If date1 > date2 Then
            dateTemp = date1
            date1 = date2
            date2 = dateTemp
            blnDifferenceIsPositive = False
        Else
            blnDifferenceIsPositive = True
        End If
        intSafetyCounter = DateDiff(DateInterval.Day, date1, date2)     ' Safety counter to help prevent infinite loops


        ' --- Counting loop ---
        Do While date1 < date2 And intSafetyCounter > 0
            If date1.DayOfWeek = DayOfWeek.Saturday Or date1.DayOfWeek = DayOfWeek.Sunday Then
                ' Don't increment counter
            Else
                intOutput += 1
            End If
            intSafetyCounter -= 1
        Loop


        ' --- Outputting data ---
        If blnDifferenceIsPositive = False Then intOutput = intOutput * -1
        Return intOutput

    End Function

End Module
