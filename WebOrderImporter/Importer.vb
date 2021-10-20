Imports System.Configuration
Imports System.Data.SqlClient
Imports System.Text

Module Importer

    Const SAMPLE_DISCOUNT_ON As Boolean = False
    Const CURRENCY_CONVERT_POUNDS_TO_EUROS As Decimal = 1.2
    Const PRODSYS_LOCK_USERNAME As String = "BLUK"

    ' --- Class level variables ---
    Dim lst_ProdSys_ProductNames As Dictionary(Of String, String)           ' Translates Website Product names into ProdSys product names
    Dim lst_ProdSys_ItemCodePrefix As Dictionary(Of String, String)         ' Translates Website Product names into ProdSys item code prefixes
    Dim lst_ProdSys_LineType As Dictionary(Of String, String)               ' Translates Website Product names into ProdSys order line types
    Dim lst_ProdSys_ManufacturerCodes As Dictionary(Of String, String)      ' Translates website manufacturer names into 2 character manufacturer codes used in item codes
    Dim lst_ProdSys_FittingTypes As Dictionary(Of String, String)           ' Translates website product fitting types into ProdSys fitting types
    Dim lst_SpareParts As List(Of String)                                   ' List of Spare Parts products
    Dim class_blnIsSparesOrder As Boolean                                   ' Boolean to indicate if an order is a spare-parts only order


    Public Enum enumWebsites
        UK
        IE
        TEST
        BPPE_UK
    End Enum


    ''' <summary>
    ''' This function adds the INSERT statement and associated parameters required to import an order header record to the supplied command 
    ''' </summary>
    ''' <param name="cmd">The full SQL transaction command to add an order header insert statement / parameters to</param>
    ''' <param name="drOrderHeader">The ORD_HEADER record from the website to create an insert statement from</param>
    ''' <param name="strWebsite">The location code of the website (e.g. UK)</param>
    ''' <param name="orderLeadTime">Information about the lead time associated with the order (derived from the order lines)</param>
    ''' <remarks>
    ''' Author: Huw Day - 05/02/2018
    ''' The command is passed in by reference. Any changes will be available to the calling function without
    ''' the command needing to be passed back.
    ''' </remarks>
    Sub Build_OrderHeader(ByRef cmd As SqlCommand, drOrderHeader As DataRow, strWebsite As String, orderLeadTime As OrderLeadTime)


        ' --- Defining objects and variables ---
        Dim strOrdHeaderSql As String
        Dim strCurrencyRate As String
        Dim dateOfOrder As Date
        Dim dateQuotedDelivery As Date
        Dim dateDispatchBy As Date
        Dim intFactoryBufferDays As Integer = 0
        Dim intWorkingDaysDifference As Integer = 0
        Dim blnUse24HourDelivery As Boolean = False
        Dim blnOnHold As Boolean = False
        Dim param_OH_Order_no As SqlParameter
        Dim param_OH_Order_date As SqlParameter
        Dim param_OH_Order_priority As SqlParameter
        Dim param_OH_Order_type As SqlParameter
        Dim param_OH_Client_name As SqlParameter
        Dim param_OH_Client_email As SqlParameter
        Dim param_OH_Client_tel As SqlParameter
        Dim param_OH_Client_street As SqlParameter
        Dim param_OH_Client_town As SqlParameter
        Dim param_OH_Client_country As SqlParameter
        Dim param_OH_Client_postcode As SqlParameter
        Dim param_OH_STATUS As SqlParameter
        Dim param_OH_INVOICEACCOUNT As SqlParameter
        Dim param_OH_DISPATCHBYDATE As SqlParameter
        Dim param_OH_OrderIs24HrDelivery As SqlParameter
        Dim param_OH_CurrencyRate As SqlParameter
        Dim param_OH_SiteLanguage As SqlParameter
        Dim param_OH_Website As SqlParameter
        Dim param_OH_WebsiteOrderID As SqlParameter
        Dim param_OH_OnHold As SqlParameter
        Dim param_OH_PreferredCourier As SqlParameter


        ' --- Constructing ORD_HEADER insert statement ---
        strOrdHeaderSql = "INSERT INTO ORD_HEADERS (order_no, order_date, order_priority, order_type, client_name, client_email, " _
                        & "client_tel, client_street, client_town, client_county, client_postcode, STATUS, INVOICEACCOUNT, DISPATCHBYDATE, " _
                        & "OrderIs24HrDelivery, CurrencyRate, SiteLanguage, Website, WebsiteOrderID, ONHOLD, PreferredCourier) " _
                        & "VALUES ( @order_no_OH, @order_date_OH, @order_priority_OH, @order_type_OH, @client_name_OH, @client_email_OH, " _
                        & "@client_tel_OH, @client_street_OH, @client_town_OH, @client_country_OH, @client_postcode_OH, @STATUS_OH, @INVOICEACCOUNT_OH, " _
                        & "@DISPATCHBYDATE_OH, @OrderIs24HrDelivery_OH, @CurrencyRate_OH, @SiteLanguage_OH, @Website_OH, @WebsiteOrderID_OH, @OnHold_OH, @PreferredCourier_OH); "


        ' --- Pre-calculating some parameter values ---
        strCurrencyRate = Currency_GetConversionRatio(strWebsite).ToString()
        If IsNumeric(strCurrencyRate) = False Then
            strCurrencyRate = "1"
        End If
        If strWebsite.ToUpper = "TEST" Then blnOnHold = True Else blnOnHold = False

        dateOfOrder = CDate(drOrderHeader("DateEntered"))
        dateQuotedDelivery = CDate(drOrderHeader("QuotedDeliveryDate"))
        If drOrderHeader.Table.Columns.Contains("FBDaTOO") Then Integer.TryParse(drOrderHeader("FBDaTOO"), intFactoryBufferDays)    ' Attempting to retrieve factory buffer days at time of order (default 0, as set above)
        dateDispatchBy = Utilities.SubtractWorkingDays(dateQuotedDelivery, 2 + intFactoryBufferDays)        ' Pre-calculating the dispatch-by date (2 days for DPD delivery, plus however many days were set as a buffer)

        If orderLeadTime.LeadTimeDays = 1 Then
            blnUse24HourDelivery = True
            If dateOfOrder.TimeOfDay < orderLeadTime.CutoffTime Then
                dateDispatchBy = Today
            Else
                dateDispatchBy = Utilities.AddWorkingDays(Today, 1)
            End If
        Else
            If dateDispatchBy < Today Then dateDispatchBy = Today
        End If


        ' --- Setting parameters ---
        param_OH_Order_no = New SqlParameter("@order_no_OH", "BL" & strWebsite & drOrderHeader("Order_No").ToString)
        param_OH_Order_date = New SqlParameter("@order_date_OH", drOrderHeader("DateEntered"))
        param_OH_Order_priority = New SqlParameter("@order_priority_OH", "1")
        param_OH_Order_type = New SqlParameter("@order_type_OH", "normal")
        param_OH_Client_name = New SqlParameter("@client_name_OH", drOrderHeader("Delivery_FirstName") & " " & drOrderHeader("Delivery_LastName"))
        param_OH_Client_email = New SqlParameter("@client_email_OH", drOrderHeader("Billing_Email"))
        param_OH_Client_tel = New SqlParameter("@client_tel_OH", drOrderHeader("Billing_PhoneNo"))
        param_OH_Client_street = New SqlParameter("@client_street_OH", drOrderHeader("Delivery_HouseNumber") & " " & drOrderHeader("Delivery_Address"))
        param_OH_Client_town = New SqlParameter("@client_town_OH", drOrderHeader("Delivery_City"))
        param_OH_Client_country = New SqlParameter("@client_country_OH", drOrderHeader("Delivery_Country"))
        param_OH_Client_postcode = New SqlParameter("@client_postcode_OH", drOrderHeader("Delivery_Postcode"))
        param_OH_STATUS = New SqlParameter("@STATUS_OH", "NEW")
        param_OH_INVOICEACCOUNT = New SqlParameter("@INVOICEACCOUNT_OH", "CAS001")
        param_OH_DISPATCHBYDATE = New SqlParameter("@DISPATCHBYDATE_OH", dateDispatchBy)
        If blnUse24HourDelivery Then
            param_OH_OrderIs24HrDelivery = New SqlParameter("@OrderIs24HrDelivery_OH", "YES")
        Else
            param_OH_OrderIs24HrDelivery = New SqlParameter("@OrderIs24HrDelivery_OH", DBNull.Value)
        End If
        param_OH_CurrencyRate = New SqlParameter("@CurrencyRate_OH", strCurrencyRate)
        param_OH_SiteLanguage = New SqlParameter("@SiteLanguage_OH", "eng-uk")
        param_OH_Website = New SqlParameter("@Website_OH", strWebsite)
        param_OH_WebsiteOrderID = New SqlParameter("@WebsiteOrderID_OH", drOrderHeader("Order_No").ToString)
        param_OH_OnHold = New SqlParameter("@OnHold_OH", blnOnHold)
        param_OH_PreferredCourier = New SqlParameter("@PreferredCourier_OH", DBNull.Value)
        cmd.Parameters.Add(param_OH_Order_no)
        cmd.Parameters.Add(param_OH_Order_date)
        cmd.Parameters.Add(param_OH_Order_priority)
        cmd.Parameters.Add(param_OH_Order_type)
        cmd.Parameters.Add(param_OH_Client_name)
        cmd.Parameters.Add(param_OH_Client_email)
        cmd.Parameters.Add(param_OH_Client_tel)
        cmd.Parameters.Add(param_OH_Client_street)
        cmd.Parameters.Add(param_OH_Client_town)
        cmd.Parameters.Add(param_OH_Client_country)
        cmd.Parameters.Add(param_OH_Client_postcode)
        cmd.Parameters.Add(param_OH_STATUS)
        cmd.Parameters.Add(param_OH_INVOICEACCOUNT)
        cmd.Parameters.Add(param_OH_DISPATCHBYDATE)
        cmd.Parameters.Add(param_OH_OrderIs24HrDelivery)
        cmd.Parameters.Add(param_OH_CurrencyRate)
        cmd.Parameters.Add(param_OH_SiteLanguage)
        cmd.Parameters.Add(param_OH_Website)
        cmd.Parameters.Add(param_OH_WebsiteOrderID)
        cmd.Parameters.Add(param_OH_OnHold)
        cmd.Parameters.Add(param_OH_PreferredCourier)


        ' --- Writing the insert header SQL into the command ---
        cmd.CommandText = cmd.CommandText.Replace("/*ORD_HEADER_BLOCK*/", strOrdHeaderSql)


        ' ## Command now has the necessary elements to insert the order header ##

    End Sub


    ''' <summary>
    ''' This function adds the INSERT statement and associated parameters required to import an order
    ''' header record from the Bloc FaceShield website to the supplied command.
    ''' </summary>
    ''' <param name="cmd"></param>
    ''' <param name="drOrderHeader"></param>
    ''' <param name="strWebsite"></param>
    ''' <remarks>
    ''' Author: Huw Day - 23/07/2020
    ''' The command is passed in by reference. Any changes will be available to the calling function without
    ''' the command needing to be passed back.
    ''' </remarks>
    Sub Build_OrderHeader_BFS(ByRef cmd As SqlCommand, drOrderHeader As DataRow, strWebsite As String)


        ' --- Defining objects and variables ---
        Dim strOrdHeaderSql As String
        Dim strCurrencyRate As String
        Dim dateDispatchBy As Date
        Dim blnOnHold As Boolean = False
        Dim param_OH_Order_no As SqlParameter
        Dim param_OH_Order_date As SqlParameter
        Dim param_OH_Order_priority As SqlParameter
        Dim param_OH_Order_type As SqlParameter
        Dim param_OH_Client_name As SqlParameter
        Dim param_OH_Client_email As SqlParameter
        Dim param_OH_Client_tel As SqlParameter
        Dim param_OH_Client_street As SqlParameter
        Dim param_OH_Client_town As SqlParameter
        Dim param_OH_Client_country As SqlParameter
        Dim param_OH_Client_postcode As SqlParameter
        Dim param_OH_STATUS As SqlParameter
        Dim param_OH_INVOICEACCOUNT As SqlParameter
        Dim param_OH_DISPATCHBYDATE As SqlParameter
        Dim param_OH_CurrencyRate As SqlParameter
        Dim param_OH_SiteLanguage As SqlParameter
        Dim param_OH_Website As SqlParameter
        Dim param_OH_WebsiteOrderID As SqlParameter
        Dim param_OH_OnHold As SqlParameter


        ' --- Constructing ORD_HEADER insert statement ---
        strOrdHeaderSql = "INSERT INTO ORD_HEADERS (order_no, order_date, order_priority, order_type, client_name, client_email, " _
                        & "client_tel, client_street, client_town, client_county, client_postcode, STATUS, INVOICEACCOUNT, DISPATCHBYDATE, " _
                        & "CurrencyRate, SiteLanguage, Website, WebsiteOrderID, ONHOLD) " _
                        & "VALUES ( @order_no_OH, @order_date_OH, @order_priority_OH, @order_type_OH, @client_name_OH, @client_email_OH, " _
                        & "@client_tel_OH, @client_street_OH, @client_town_OH, @client_country_OH, @client_postcode_OH, @STATUS_OH, @INVOICEACCOUNT_OH, " _
                        & "@DISPATCHBYDATE_OH, @CurrencyRate_OH, @SiteLanguage_OH, @Website_OH, @WebsiteOrderID_OH, @OnHold_OH); "


        ' --- Pre-calculating some parameter values ---
        strCurrencyRate = Currency_GetConversionRatio(strWebsite).ToString()
        If IsNumeric(strCurrencyRate) = False Then
            strCurrencyRate = "1"
        End If


        ' --- Setting parameters ---
        param_OH_Order_no = New SqlParameter("@order_no_OH", strWebsite & drOrderHeader("Order_No").ToString)
        param_OH_Order_date = New SqlParameter("@order_date_OH", drOrderHeader("DateEntered"))
        param_OH_Order_priority = New SqlParameter("@order_priority_OH", "1")
        param_OH_Order_type = New SqlParameter("@order_type_OH", "normal")
        param_OH_Client_name = New SqlParameter("@client_name_OH", drOrderHeader("Delivery_FirstName") & " " & drOrderHeader("Delivery_LastName"))
        param_OH_Client_email = New SqlParameter("@client_email_OH", drOrderHeader("Billing_Email"))
        param_OH_Client_tel = New SqlParameter("@client_tel_OH", drOrderHeader("Billing_PhoneNo"))
        param_OH_Client_street = New SqlParameter("@client_street_OH", drOrderHeader("Delivery_HouseNumber") & " " & drOrderHeader("Delivery_Address"))
        param_OH_Client_town = New SqlParameter("@client_town_OH", drOrderHeader("Delivery_City"))
        param_OH_Client_country = New SqlParameter("@client_country_OH", drOrderHeader("Delivery_Country"))
        param_OH_Client_postcode = New SqlParameter("@client_postcode_OH", drOrderHeader("Delivery_Postcode"))
        param_OH_STATUS = New SqlParameter("@STATUS_OH", "NEW")
        param_OH_INVOICEACCOUNT = New SqlParameter("@INVOICEACCOUNT_OH", "CAS001")
        dateDispatchBy = Utilities.AddWorkingDays(CDate(drOrderHeader("DateEntered")), 2)
        If dateDispatchBy < Today Then dateDispatchBy = Today                                               ' If the dispatch-by date is before today, then cap it at today
        param_OH_DISPATCHBYDATE = New SqlParameter("@DISPATCHBYDATE_OH", dateDispatchBy)
        param_OH_CurrencyRate = New SqlParameter("@CurrencyRate_OH", strCurrencyRate)
        param_OH_SiteLanguage = New SqlParameter("@SiteLanguage_OH", "eng-uk")
        param_OH_Website = New SqlParameter("@Website_OH", strWebsite)
        param_OH_WebsiteOrderID = New SqlParameter("@WebsiteOrderID_OH", drOrderHeader("Order_No").ToString)
        param_OH_OnHold = New SqlParameter("@OnHold_OH", blnOnHold)
        cmd.Parameters.Add(param_OH_Order_no)
        cmd.Parameters.Add(param_OH_Order_date)
        cmd.Parameters.Add(param_OH_Order_priority)
        cmd.Parameters.Add(param_OH_Order_type)
        cmd.Parameters.Add(param_OH_Client_name)
        cmd.Parameters.Add(param_OH_Client_email)
        cmd.Parameters.Add(param_OH_Client_tel)
        cmd.Parameters.Add(param_OH_Client_street)
        cmd.Parameters.Add(param_OH_Client_town)
        cmd.Parameters.Add(param_OH_Client_country)
        cmd.Parameters.Add(param_OH_Client_postcode)
        cmd.Parameters.Add(param_OH_STATUS)
        cmd.Parameters.Add(param_OH_INVOICEACCOUNT)
        cmd.Parameters.Add(param_OH_DISPATCHBYDATE)
        cmd.Parameters.Add(param_OH_CurrencyRate)
        cmd.Parameters.Add(param_OH_SiteLanguage)
        cmd.Parameters.Add(param_OH_Website)
        cmd.Parameters.Add(param_OH_WebsiteOrderID)
        cmd.Parameters.Add(param_OH_OnHold)


        ' --- Writing the insert header SQL into the command ---
        cmd.CommandText = cmd.CommandText.Replace("/*ORD_HEADER_BLOCK*/", strOrdHeaderSql)


        ' ## Command now has the necessary elements to insert the order header ##

    End Sub


    Sub Build_OrderLine_BFS_Delivery(ByRef cmd As SqlCommand, drOrderHeader As DataRow, strWebsite As String)

        ' --- Defining objects and variables ---
        Dim strOrderDeliveryLine As String
        Dim decDeliveryCost As Decimal              ' ### TEMPORARY!!! This should be removed once the website starts accurately modelling delivery costs and it's tax
        Dim param_DV_OrderNo As SqlParameter
        Dim param_DV_DeliveryCost As SqlParameter
        Dim param_DV_InvoiceUnitPrice As SqlParameter


        ' --- Calculating delivery cost ex tax ---
        decDeliveryCost = 7.95 / 1.2     ' ### TEMPORARY!!! This should be removed once the website starts accurately modelling delivery costs and it's tax


        ' --- Converting website currency to pounds ---
        decDeliveryCost = Currency_ConvertToPounds(decDeliveryCost, strWebsite)


        ' --- Writing Delivery insert statement ---
        strOrderDeliveryLine = "INSERT INTO ORD_LINES (order_no, item_code, item_description, item_price, item_qty, ORD_HEADERS_ID, " _
            & "ALUCOLOUR, FABCOLOUR, KeepPricingForInvoicing, LINETYPE, InvoiceUnitPrice) " _
            & "VALUES (@OrderNo_DV, 'DELUK', 'Delivery costs -', @DeliveryCost_DV, '1', @ProdSysID, 'NA', 'NA', 'True', 'DELIVERY', @InvoiceUnitPrice_DV); "


        ' --- Setting parameters ---
        param_DV_OrderNo = New SqlParameter("@OrderNo_DV", strWebsite & drOrderHeader("Order_No").ToString)
        param_DV_DeliveryCost = New SqlParameter("@DeliveryCost_DV", decDeliveryCost)
        param_DV_InvoiceUnitPrice = New SqlParameter("@InvoiceUnitPrice_DV", decDeliveryCost)


        ' --- Adding parameters to command ---
        cmd.Parameters.Add(param_DV_OrderNo)
        cmd.Parameters.Add(param_DV_DeliveryCost)
        cmd.Parameters.Add(param_DV_InvoiceUnitPrice)


        ' --- Writing the SQL insert statement into the command's text ---
        ' Ensure the text "/*ORD_LINE_BLOCKS*/" is persisted, as other functions will try to add more lines into
        ' the command text at this point
        cmd.CommandText = cmd.CommandText.Replace("/*ORD_LINE_BLOCKS*/", strOrderDeliveryLine & vbNewLine & "/*ORD_LINE_BLOCKS*/")


        ' ## Command now has the necessary elements to insert the Delivery line ##

    End Sub


    ''' <summary>
    ''' This function adds the INSERT statement and parameters to import a single order line from the 
    ''' BlocFaceShields website.
    ''' </summary>
    ''' <param name="cmd"></param>
    ''' <param name="drOrderLine"></param>
    ''' <param name="strWebsite"></param>
    ''' <remarks>
    ''' Author: Huw Day - 06/08/2020
    ''' </remarks>
    Sub Build_OrderLine_BFS_Product(ByRef cmd As SqlCommand, drOrderLine As DataRow, strWebsite As String)


        ' --- Defining objects and variables ---
        Dim strProductLineQuery As String
        Dim strSourceLineID As String = drOrderLine("Line_ID")


        ' --- Defining parameters ---
        Dim param_OL_OrderNo As SqlParameter
        Dim param_OL_ItemCode As SqlParameter
        Dim param_OL_ItemDescription As SqlParameter
        Dim param_OL_ItemPrice As SqlParameter
        Dim param_OL_ItemQty As SqlParameter
        Dim param_OL_A As SqlParameter
        Dim param_OL_B As SqlParameter
        Dim param_OL_C As SqlParameter
        Dim param_OL_D As SqlParameter
        Dim param_OL_E As SqlParameter
        Dim param_OL_F As SqlParameter
        Dim param_OL_G As SqlParameter
        Dim param_OL_AluColour As SqlParameter
        Dim param_OL_FabColour As SqlParameter
        Dim param_OL_FabColour2 As SqlParameter
        Dim param_OL_BracketColour As SqlParameter
        Dim param_OL_KeepPricingForInvoicing As SqlParameter
        Dim param_OL_CustomerReferenceNo As SqlParameter
        Dim param_OL_LineType As SqlParameter
        Dim param_OL_HandleType As SqlParameter
        Dim param_OL_ChainType As SqlParameter
        Dim param_OL_ChainSide As SqlParameter
        Dim param_OL_FabricRoleDir As SqlParameter
        Dim param_OL_FittingType As SqlParameter
        Dim param_OL_InstallationHeight As SqlParameter
        Dim param_OL_ProductRange As SqlParameter
        Dim param_OL_ProductName As SqlParameter
        Dim param_OL_InvoiceUnitPrice As SqlParameter
        Dim param_OL_Operation As SqlParameter
        Dim param_OL_RuleSet As SqlParameter


        ' --- Defining parameter content variables ---
        Dim strOrderNo_value As String
        Dim strItemCode_value As String
        Dim strItemDescription_value As String
        Dim decItemPrice_value As Decimal = 0               ' Defaults to 0 if no valid price could be found
        Dim decMeasureProtectValue As Decimal = 0
        Dim intItemQty_value As Integer = 1                 ' Defaults to 1 if no valid quantity could be found
        Dim intA_value As Integer = 0
        Dim intB_value As Integer = 0
        Dim intC_value As Integer = 0
        Dim intD_value As Integer = 0
        Dim intE_value As Integer = 0
        Dim intF_value As Integer = 0
        Dim intG_value As Integer = 0
        Dim strAluColour_value As String = ""
        Dim strFabColour_value As String = ""
        Dim strFabColour2_value As String = ""
        Dim strBracketColour_value As String = ""
        Dim strKeepPricingForInvoicing_value As String = "True"     ' Always True
        Dim strCustomerReferenceNo_value As String = ""
        Dim strLineType_value As String
        Dim strHandleType_value As String = ""
        Dim strChainType_value As String = ""
        Dim strChainSide_value As String = "Left"                   ' defaults to Left
        Dim strFabricRoleDir_value As String = "Standard"           ' defaults to Standard
        Dim strFittingType_value As String = ""
        Dim strInstallationHeight_value As String = ""              ' Always empty
        Dim strProductRange_value As String = "Bloc"
        Dim strProductName_value As String = ""
        Dim strOperation_value As String = ""
        Dim strRuleSet_value As String = ""


        ' --- Writing Product order line insert statement ---
        ' As the below SQL statement could be loaded multiple times into the transaction command, the parameter names
        ' have to be unique to the order line. The LineID from the website's order line table is appended to each
        ' parameter name to help ensure that each parameter in the command is unique.
        strProductLineQuery = "INSERT INTO ORD_LINES (order_no, item_code, item_description, item_price, item_qty, " _
            & "ORD_HEADERS_ID, A, B, C, D, E, F, G, ALUCOLOUR, FABCOLOUR, FABCOLOUR2, BracketColour, KeepPricingForInvoicing, " _
            & "CustomerReferenceNo, LINETYPE, HandleType, ChainType, ChainSide, FabricRoleDir, FittingType, " _
            & "InstallationHeight, ProductRange, ProductName, InvoiceUnitPrice, Operation, RuleSet) " _
            & "VALUES (" _
            & "@OrderNo_OL" & strSourceLineID _
            & ", @ItemCode_OL" & strSourceLineID _
            & ", @ItemDescription_OL" & strSourceLineID _
            & ", @ItemPrice_OL" & strSourceLineID _
            & ", @ItemQty_OL" & strSourceLineID _
            & ", @ProdSysID" _
            & ", @A_OL" & strSourceLineID _
            & ", @B_OL" & strSourceLineID _
            & ", @C_OL" & strSourceLineID _
            & ", @D_OL" & strSourceLineID _
            & ", @E_OL" & strSourceLineID _
            & ", @F_OL" & strSourceLineID _
            & ", @G_OL" & strSourceLineID _
            & ", @AluColour_OL" & strSourceLineID _
            & ", @FabColour_OL" & strSourceLineID _
            & ", @FabColour2_OL" & strSourceLineID _
            & ", @BracketColour_OL" & strSourceLineID _
            & ", @KeepPricingForInvoicing_OL" & strSourceLineID _
            & ", @CustomerReferenceNo_OL" & strSourceLineID _
            & ", @LineType_OL" & strSourceLineID _
            & ", @HandleType_OL" & strSourceLineID _
            & ", @ChainType_OL" & strSourceLineID _
            & ", @ChainSide_OL" & strSourceLineID _
            & ", @FabricRoleDir_OL" & strSourceLineID _
            & ", @FittingType_OL" & strSourceLineID _
            & ", @InstallationHeight_OL" & strSourceLineID _
            & ", @ProductRange_OL" & strSourceLineID _
            & ", @ProductName_OL" & strSourceLineID _
            & ", @InvoiceUnitPrice_OL" & strSourceLineID _
            & ", @Operation_OL" & strSourceLineID _
            & ", @RuleSet_OL" & strSourceLineID _
            & ")"
        ' Note: @ProdSysID is not a vb SqlParameter, it is set as a SQL variable in the main transaction block. See Import_Orders function


        ' --- Setting parameter content variables ---
        ' -- order_no --
        strOrderNo_value = strWebsite & drOrderLine("Order_No").ToString()

        ' -- item_code --
        strItemCode_value = drOrderLine("ProdSysCode").ToString()

        ' -- item description --
        strItemDescription_value = drOrderLine("ProdSysName").ToString()

        ' -- Quantity --
        If drOrderLine.Table.Columns.Contains("Quantity") Then
            If IsDBNull(drOrderLine("Quantity")) = False Then
                If IsNumeric(drOrderLine("Quantity")) Then
                    intItemQty_value = CInt(drOrderLine("Quantity"))
                End If
            End If
        End If

        ' -- ItemPrice --
        If drOrderLine.Table.Columns.Contains("Price_Final_ExTax") Then
            Decimal.TryParse(drOrderLine("Price_Final_ExTax"), decItemPrice_value)
        End If
        decItemPrice_value = Currency_ConvertToPounds(decItemPrice_value, strWebsite)
        decItemPrice_value = decItemPrice_value / intItemQty_value                                  ' Converting Line Total price to price per unit, for the benefit of the production system
        decItemPrice_value = Math.Round(decItemPrice_value, 2, MidpointRounding.AwayFromZero)       ' Rounding to 2 decimal places

        ' -- LineType --
        strLineType_value = drOrderLine("ProdSys_LineType").ToString()

        ' -- Product Name --
        strProductName_value = drOrderLine("ProdSys_LineType").ToString()


        ' --- Setting parameters ---
        param_OL_OrderNo = New SqlParameter("@OrderNo_OL" & strSourceLineID, strOrderNo_value)
        param_OL_ItemCode = New SqlParameter("@ItemCode_OL" & strSourceLineID, strItemCode_value)
        param_OL_ItemDescription = New SqlParameter("@ItemDescription_OL" & strSourceLineID, strItemDescription_value)
        param_OL_ItemPrice = New SqlParameter("@ItemPrice_OL" & strSourceLineID, decItemPrice_value)
        param_OL_ItemQty = New SqlParameter("@ItemQty_OL" & strSourceLineID, intItemQty_value)
        param_OL_A = New SqlParameter("@A_OL" & strSourceLineID, intA_value)
        param_OL_B = New SqlParameter("@B_OL" & strSourceLineID, intB_value)
        param_OL_C = New SqlParameter("@C_OL" & strSourceLineID, intC_value)
        param_OL_D = New SqlParameter("@D_OL" & strSourceLineID, intD_value)
        param_OL_E = New SqlParameter("@E_OL" & strSourceLineID, intE_value)
        param_OL_F = New SqlParameter("@F_OL" & strSourceLineID, intF_value)
        param_OL_G = New SqlParameter("@G_OL" & strSourceLineID, intG_value)
        param_OL_AluColour = New SqlParameter("@AluColour_OL" & strSourceLineID, strAluColour_value)
        param_OL_FabColour = New SqlParameter("@FabColour_OL" & strSourceLineID, strFabColour_value)
        param_OL_FabColour2 = New SqlParameter("@FabColour2_OL" & strSourceLineID, strFabColour2_value)
        param_OL_BracketColour = New SqlParameter("@BracketColour_OL" & strSourceLineID, strBracketColour_value)
        param_OL_KeepPricingForInvoicing = New SqlParameter("@KeepPricingForInvoicing_OL" & strSourceLineID, strKeepPricingForInvoicing_value)
        param_OL_CustomerReferenceNo = New SqlParameter("@CustomerReferenceNo_OL" & strSourceLineID, strCustomerReferenceNo_value)
        param_OL_LineType = New SqlParameter("@LineType_OL" & strSourceLineID, strLineType_value)
        param_OL_HandleType = New SqlParameter("@HandleType_OL" & strSourceLineID, strHandleType_value)
        param_OL_ChainType = New SqlParameter("@ChainType_OL" & strSourceLineID, strChainType_value)
        param_OL_ChainSide = New SqlParameter("@ChainSide_OL" & strSourceLineID, strChainSide_value)
        param_OL_FabricRoleDir = New SqlParameter("@FabricRoleDir_OL" & strSourceLineID, strFabricRoleDir_value)
        param_OL_FittingType = New SqlParameter("@FittingType_OL" & strSourceLineID, strFittingType_value)
        param_OL_InstallationHeight = New SqlParameter("@InstallationHeight_OL" & strSourceLineID, strInstallationHeight_value)
        param_OL_ProductRange = New SqlParameter("@ProductRange_OL" & strSourceLineID, strProductRange_value)
        param_OL_ProductName = New SqlParameter("@ProductName_OL" & strSourceLineID, strProductName_value)
        param_OL_InvoiceUnitPrice = New SqlParameter("@InvoiceUnitPrice_OL" & strSourceLineID, decItemPrice_value)
        param_OL_Operation = New SqlParameter("@Operation_OL" & strSourceLineID, strOperation_value)
        param_OL_RuleSet = New SqlParameter("@RuleSet_OL" & strSourceLineID, strRuleSet_value)


        ' --- Adding parameters ---
        cmd.Parameters.Add(param_OL_OrderNo)
        cmd.Parameters.Add(param_OL_ItemCode)
        cmd.Parameters.Add(param_OL_ItemDescription)
        cmd.Parameters.Add(param_OL_ItemPrice)
        cmd.Parameters.Add(param_OL_ItemQty)
        cmd.Parameters.Add(param_OL_A)
        cmd.Parameters.Add(param_OL_B)
        cmd.Parameters.Add(param_OL_C)
        cmd.Parameters.Add(param_OL_D)
        cmd.Parameters.Add(param_OL_E)
        cmd.Parameters.Add(param_OL_F)
        cmd.Parameters.Add(param_OL_G)
        cmd.Parameters.Add(param_OL_AluColour)
        cmd.Parameters.Add(param_OL_FabColour)
        cmd.Parameters.Add(param_OL_FabColour2)
        cmd.Parameters.Add(param_OL_BracketColour)
        cmd.Parameters.Add(param_OL_KeepPricingForInvoicing)
        cmd.Parameters.Add(param_OL_CustomerReferenceNo)
        cmd.Parameters.Add(param_OL_LineType)
        cmd.Parameters.Add(param_OL_HandleType)
        cmd.Parameters.Add(param_OL_ChainType)
        cmd.Parameters.Add(param_OL_ChainSide)
        cmd.Parameters.Add(param_OL_FabricRoleDir)
        cmd.Parameters.Add(param_OL_FittingType)
        cmd.Parameters.Add(param_OL_InstallationHeight)
        cmd.Parameters.Add(param_OL_ProductRange)
        cmd.Parameters.Add(param_OL_ProductName)
        cmd.Parameters.Add(param_OL_InvoiceUnitPrice)
        cmd.Parameters.Add(param_OL_Operation)
        cmd.Parameters.Add(param_OL_RuleSet)


        ' --- Adding SQL insert statement to command text ---
        ' Ensure the text "/*ORD_LINE_BLOCKS*/" is persisted, as other functions will try to add more lines into
        ' the command text at this point
        cmd.CommandText = cmd.CommandText.Replace("/*ORD_LINE_BLOCKS*/", strProductLineQuery & vbNewLine & "/*ORD_LINE_BLOCKS*/")


    End Sub


    ''' <summary>
    ''' This function adds the INSERT statement and parameters to add a Delivery order line to the production system
    ''' </summary>
    ''' <param name="cmd">The full SQL transaction command to add a order lines insert statements / parameters to</param>
    ''' <param name="drOrderHeader">The ORD_HEADER record to create a delivery line for</param>
    ''' <param name="strWebsite">The location code of the source website. E.g. UK</param>
    ''' <remarks>
    ''' Author: Huw Day - 06/02/2018
    ''' The command is passed in by reference. Any changes will be available to the calling function without
    ''' the command needing to be passed back.
    ''' </remarks>
    Sub Build_OrderLine_Delivery(ByRef cmd As SqlCommand, drOrderHeader As DataRow, strWebsite As String)


        ' --- Defining objects and variables ---
        Dim strOrderDeliveryLine As String
        Dim decDeliveryCost As Decimal              ' ### TEMPORARY!!! This should be removed once the website starts accurately modelling delivery costs and it's tax
        Dim param_DV_OrderNo As SqlParameter
        Dim param_DV_DeliveryCost As SqlParameter
        Dim param_DV_InvoiceUnitPrice As SqlParameter


        ' --- Calculating delivery cost ex tax ---
        decDeliveryCost = CDec(drOrderHeader("DeliveryCost"))
        decDeliveryCost = decDeliveryCost / 1.2     ' ### TEMPORARY!!! This should be removed once the website starts accurately modelling delivery costs and it's tax


        ' --- Converting website currency to pounds ---
        decDeliveryCost = Currency_ConvertToPounds(decDeliveryCost, strWebsite)


        ' --- Writing Delivery insert statement ---
        strOrderDeliveryLine = "INSERT INTO ORD_LINES (order_no, item_code, item_description, item_price, item_qty, ORD_HEADERS_ID, " _
            & "ALUCOLOUR, FABCOLOUR, KeepPricingForInvoicing, LINETYPE, InvoiceUnitPrice) " _
            & "VALUES (@OrderNo_DV, 'DELUK', 'Delivery costs -', @DeliveryCost_DV, '1', @ProdSysID, 'NA', 'NA', 'True', 'DELIVERY', @InvoiceUnitPrice_DV); "


        ' --- Setting parameters ---
        param_DV_OrderNo = New SqlParameter("@OrderNo_DV", "BL" & strWebsite & drOrderHeader("Order_No").ToString)
        param_DV_DeliveryCost = New SqlParameter("@DeliveryCost_DV", decDeliveryCost)
        param_DV_InvoiceUnitPrice = New SqlParameter("@InvoiceUnitPrice_DV", decDeliveryCost)


        ' --- Adding parameters to command ---
        cmd.Parameters.Add(param_DV_OrderNo)
        cmd.Parameters.Add(param_DV_DeliveryCost)
        cmd.Parameters.Add(param_DV_InvoiceUnitPrice)


        ' --- Writing the SQL insert statement into the command's text ---
        ' Ensure the text "/*ORD_LINE_BLOCKS*/" is persisted, as other functions will try to add more lines into
        ' the command text at this point
        cmd.CommandText = cmd.CommandText.Replace("/*ORD_LINE_BLOCKS*/", strOrderDeliveryLine & vbNewLine & "/*ORD_LINE_BLOCKS*/")


        ' ## Command now has the necessary elements to insert the Delivery line ##

    End Sub

    ''' <summary>
    ''' This function adds the INSERT statement and parameters to add a FRAME line to the production system. This is largely copied
    ''' from Build_OrderLine_Product, with a few updates
    ''' </summary>
    ''' <param name="cmd">The full SQL transaction command to add a order lines insert statements / parameters to</param>
    ''' <param name="drOrderLine">The ORD_LINE record to import</param>
    ''' <param name="strWebsite">The location code of the source website. E.g. UK</param>
    ''' <remarks>
    ''' Author: Huw Day - 06/02/2018
    ''' The command is passed in by reference. Any changes will be available to the calling function without
    ''' the command needing to be passed back.
    ''' </remarks>
    Sub Build_OrderLine_Frame(ByRef cmd As SqlCommand, drOrderLine As DataRow, strWebsite As String)


        ' --- Defining objects and variables ---
        Dim strProductLineQuery As String
        Dim strSourceLineID As String = drOrderLine("LineID")
        Dim intVAT As Integer = CInt(Importer.Get_SettingFromWebsite("VAT", strWebsite))
        Dim blnIsMotorised As Boolean = 0


        ' --- Defining parameters ---
        Dim param_OLFR_OrderNo As SqlParameter
        Dim param_OLFR_ItemCode As SqlParameter
        Dim param_OLFR_ItemDescription As SqlParameter
        Dim param_OLFR_ItemPrice As SqlParameter
        Dim param_OLFR_ItemQty As SqlParameter
        Dim param_OLFR_A As SqlParameter
        Dim param_OLFR_B As SqlParameter
        Dim param_OLFR_C As SqlParameter
        Dim param_OLFR_D As SqlParameter
        Dim param_OLFR_E As SqlParameter
        Dim param_OLFR_F As SqlParameter
        Dim param_OLFR_G As SqlParameter
        Dim param_OLFR_AluColour As SqlParameter
        Dim param_OLFR_FabColour As SqlParameter
        Dim param_OLFR_FabColour2 As SqlParameter
        Dim param_OLFR_BracketColour As SqlParameter
        Dim param_OLFR_KeepPricingForInvoicing As SqlParameter
        Dim param_OLFR_CustomerReferenceNo As SqlParameter
        Dim param_OLFR_LineType As SqlParameter
        Dim param_OLFR_HandleType As SqlParameter
        Dim param_OLFR_ChainType As SqlParameter
        Dim param_OLFR_ChainSide As SqlParameter
        Dim param_OLFR_FabricRoleDir As SqlParameter
        Dim param_OLFR_FittingType As SqlParameter
        Dim param_OLFR_InstallationHeight As SqlParameter
        Dim param_OLFR_ProductRange As SqlParameter
        Dim param_OLFR_ProductName As SqlParameter
        Dim param_OLFR_InvoiceUnitPrice As SqlParameter
        Dim param_OLFR_Operation As SqlParameter
        Dim param_OLFR_RuleSet As SqlParameter


        ' --- Defining parameter content variables ---
        Dim strOrderNo_value As String
        Dim strItemCode_value As String
        Dim strItemDescription_value As String
        Dim decItemPrice_value As Decimal = 0               ' Defaults to 0 if no valid price could be found
        Dim intItemQty_value As Integer = 1                 ' Defaults to 1 if no valid quantity could be found
        Dim intA_value As Integer = 0
        Dim intB_value As Integer = 0
        Dim intC_value As Integer = 0
        Dim intD_value As Integer = 0
        Dim intE_value As Integer = 0
        Dim intF_value As Integer = 0
        Dim intG_value As Integer = 0
        Dim strAluColour_value As String = ""
        Dim strFabColour_value As String = ""
        Dim strFabColour2_value As String = ""
        Dim strBracketColour_value As String = ""
        Dim strKeepPricingForInvoicing_value As String = "True"     ' Always True
        Dim strCustomerReferenceNo_value As String = ""
        Dim strLineType_value As String
        Dim strHandleType_value As String
        Dim strChainType_value As String
        Dim strChainSide_value As String = "Left"                   ' defaults to Left
        Dim strFabricRoleDir_value As String = "Standard"           ' defaults to Standard
        Dim strFittingType_value As String = ""
        Dim strInstallationHeight_value As String = ""              ' Always empty
        Dim strProductRange_value As String
        Dim strProductName_value As String = ""
        Dim strOperation_value As String = ""
        Dim strRuleSet_value As String = ""


        ' --- Checking if product is motorised ---
        If drOrderLine.Table.Columns.Contains("Motorised") Then
            If IsDBNull(drOrderLine("Motorised")) = False Then
                blnIsMotorised = CBool(drOrderLine("Motorised"))
            End If
        End If


        ' --- Writing Product order line insert statement ---
        ' As the below SQL statement could be loaded multiple times into the transaction command, the parameter names
        ' have to be unique to the order line. The LineID from the website's order line table is appended to each
        ' parameter name to help ensure that each parameter in the command is unique.
        strProductLineQuery = "INSERT INTO ORD_LINES (order_no, item_code, item_description, item_price, item_qty, " _
            & "ORD_HEADERS_ID, A, B, C, D, E, F, G, ALUCOLOUR, FABCOLOUR, FABCOLOUR2, BracketColour, KeepPricingForInvoicing, " _
            & "CustomerReferenceNo, LINETYPE, HandleType, ChainType, ChainSide, FabricRoleDir, FittingType, " _
            & "InstallationHeight, ProductRange, ProductName, InvoiceUnitPrice, Operation, RuleSet) " _
            & "VALUES (" _
            & "@OrderNo_OLFR" & strSourceLineID _
            & ", @ItemCode_OLFR" & strSourceLineID _
            & ", @ItemDescription_OLFR" & strSourceLineID _
            & ", @ItemPrice_OLFR" & strSourceLineID _
            & ", @ItemQty_OLFR" & strSourceLineID _
            & ", @ProdSysID" _
            & ", @A_OLFR" & strSourceLineID _
            & ", @B_OLFR" & strSourceLineID _
            & ", @C_OLFR" & strSourceLineID _
            & ", @D_OLFR" & strSourceLineID _
            & ", @E_OLFR" & strSourceLineID _
            & ", @F_OLFR" & strSourceLineID _
            & ", @G_OLFR" & strSourceLineID _
            & ", @AluColour_OLFR" & strSourceLineID _
            & ", @FabColour_OLFR" & strSourceLineID _
            & ", @FabColour2_OLFR" & strSourceLineID _
            & ", @BracketColour_OLFR" & strSourceLineID _
            & ", @KeepPricingForInvoicing_OLFR" & strSourceLineID _
            & ", @CustomerReferenceNo_OLFR" & strSourceLineID _
            & ", @LineType_OLFR" & strSourceLineID _
            & ", @HandleType_OLFR" & strSourceLineID _
            & ", @ChainType_OLFR" & strSourceLineID _
            & ", @ChainSide_OLFR" & strSourceLineID _
            & ", @FabricRoleDir_OLFR" & strSourceLineID _
            & ", @FittingType_OLFR" & strSourceLineID _
            & ", @InstallationHeight_OLFR" & strSourceLineID _
            & ", @ProductRange_OLFR" & strSourceLineID _
            & ", @ProductName_OLFR" & strSourceLineID _
            & ", @InvoiceUnitPrice_OLFR" & strSourceLineID _
            & ", @Operation_OLFR" & strSourceLineID _
            & ", @RuleSet_OLFR" & strSourceLineID _
            & ")"
        ' Note: @ProdSysID is not a vb SqlParameter, it is set as a SQL variable in the main transaction block. See Import_Orders function


        ' --- Setting parameter content variables ---
        ' -- FittingType - Must be calculated at the start as other parameters depend on it -
        If drOrderLine.Table.Columns.Contains("FittingType") Then
            If IsDBNull(drOrderLine("FittingType")) = False Then
                If lst_ProdSys_FittingTypes.ContainsKey(drOrderLine("FittingType").ToString().ToLower()) Then
                    strFittingType_value = lst_ProdSys_FittingTypes(drOrderLine("FittingType").ToString.ToLower)
                End If
            End If
        End If

        ' -- ALUCOLOUR - Must be calculated at the start as other parameters depend on it -
        If drOrderLine.Table.Columns.Contains("AluminiumColour") Then
            If IsDBNull(drOrderLine("AluminiumColour")) = False Then
                strAluColour_value = drOrderLine("AluminiumColour").ToString
            End If
        End If

        ' -- order_no --
        strOrderNo_value = "BL" & strWebsite & drOrderLine("Order_No").ToString()

        ' -- item_code --
        If strFittingType_value = "Inside the Window Recess" Then
            strItemCode_value = "FRAME-WOOD-" & strFittingType_value
        ElseIf strFittingType_value = "" Or strFittingType_value = "" Then
            strItemCode_value = "FRAME-" & strFittingType_value
        Else
            strItemCode_value = "FRAME-"
        End If
        strItemCode_value = strItemCode_value & "|" & strAluColour_value & "|NA"

        ' -- item_description --
        strItemDescription_value = ""

        ' -- ItemPrice --
        decItemPrice_value = 0
        decItemPrice_value = Currency_ConvertToPounds(decItemPrice_value, strWebsite)

        ' -- Quantity --
        If drOrderLine.Table.Columns.Contains("Quantity") Then
            If IsDBNull(drOrderLine("Quantity")) = False Then
                If IsNumeric(drOrderLine("Quantity")) Then
                    intItemQty_value = CInt(drOrderLine("Quantity"))
                End If
            End If
        End If

        ' -- A --
        If drOrderLine.Table.Columns.Contains("Width") Then
            If IsDBNull(drOrderLine("Width")) = False Then
                If IsNumeric(drOrderLine("Width")) Then
                    intA_value = CInt(drOrderLine("Width"))
                End If
            End If
        End If

        ' -- B --
        If drOrderLine.Table.Columns.Contains("WidthMiddle") Then
            If IsDBNull(drOrderLine("WidthMiddle")) = False Then
                If IsNumeric(drOrderLine("WidthMiddle")) Then
                    intB_value = CInt(drOrderLine("WidthMiddle"))
                End If
            End If
        End If

        ' -- C --
        If drOrderLine.Table.Columns.Contains("WidthBottom") Then
            If IsDBNull(drOrderLine("WidthBottom")) = False Then
                If IsNumeric(drOrderLine("WidthBottom")) Then
                    intC_value = CInt(drOrderLine("WidthBottom"))
                End If
            End If
        End If

        ' -- D, E --
        ' If the product just has a Height measurement (HeightLeft and HeightRight are null), Height should go here. 
        ' Otherwise, HeightLeft should go here and Height should go in E
        If drOrderLine.Table.Columns.Contains("Height") Then
            If IsDBNull(drOrderLine("Height")) = False Then
                If IsNumeric(drOrderLine("Height")) Then
                    intD_value = CInt(drOrderLine("Height"))
                End If
            End If
        End If
        If drOrderLine.Table.Columns.Contains("HeightLeft") Then
            If IsDBNull(drOrderLine("HeightLeft")) = False Then
                If IsNumeric(drOrderLine("HeightLeft")) Then
                    intE_value = intD_value
                    intD_value = CInt(drOrderLine("HeightLeft"))
                End If
            End If
        End If

        ' -- F --
        If drOrderLine.Table.Columns.Contains("HeightRight") Then
            If IsDBNull(drOrderLine("HeightRight")) = False Then
                If IsNumeric(drOrderLine("HeightRight")) Then
                    intF_value = CInt(drOrderLine("HeightRight"))
                End If
            End If
        End If

        ' -- G --
        If drOrderLine.Table.Columns.Contains("RecessDepth") Then
            If IsDBNull(drOrderLine("RecessDepth")) = False Then
                If IsNumeric(drOrderLine("RecessDepth")) Then
                    intG_value = CInt(drOrderLine("RecessDepth"))
                End If
            End If
        End If

        ' -- FABCOLOUR --
        strFabColour_value = "NA"

        ' -- FABCOLOUR2 --
        ' - Leave at default of ""

        ' -- BracketColour --
        If strFittingType_value = "Inside the Window Recess" Then
            strBracketColour_value = "FRAMEWOOD"
        Else
            strBracketColour_value = strAluColour_value
        End If

        ' -- CustomerReferenceNo --
        If drOrderLine.Table.Columns.Contains("CustomerReference") Then
            If IsDBNull(drOrderLine("CustomerReference")) = False Then
                strCustomerReferenceNo_value = drOrderLine("CustomerReference").ToString
            End If
        End If

        ' -- LINETYPE --
        If strFittingType_value = "Inside the Window Recess" Then
            strLineType_value = "FRAMEWOOD"
        Else
            strLineType_value = "FRAME"
        End If

        ' -- HandleType --
        Select Case drOrderLine("ProductName").ToString
            Case "PremiumRollerBlind", "RollerBlind", "RollerBlindDrillFree", "RollerBlindMotorised", "TwinRollerBlind"
                strHandleType_value = "WHITE PVC BOTTOM BAR"
            Case Else
                strHandleType_value = ""
        End Select

        ' -- ChainType --
        If drOrderLine("ProductName").ToString.ToLower = "blocout40" Then
            strChainType_value = ""
        ElseIf blnIsMotorised = True Then
            strChainType_value = "MOTORISED"
        Else
            strChainType_value = "METAL"
        End If

        ' -- ChainSide --
        If drOrderLine.Table.Columns.Contains("ChainSide") Then
            If IsDBNull(drOrderLine("ChainSide")) = False Then
                strChainSide_value = drOrderLine("ChainSide").ToString
            End If
        End If

        ' -- FabricRoleDirection --
        If drOrderLine.Table.Columns.Contains("RollDirection") Then
            If IsDBNull(drOrderLine("RollDirection")) = False Then
                strFabricRoleDir_value = drOrderLine("RollDirection").ToString
            End If
        End If

        ' -- ProductRange --
        If blnIsMotorised Then strProductRange_value = "BlocMotorised" Else strProductRange_value = "Bloc"

        ' -- ProductName --
        strProductName_value = "Frame"

        ' -- Operation --
        If blnIsMotorised = True Then strOperation_value = "MOTORISED" Else strOperation_value = ""

        ' -- RuleSet --
        ' check based on already calculated production system values of ProductName and Fitting Type
        If strProductName_value.ToLower = "blocout40" And strFittingType_value.ToLower = "Edge of Window Recess" Then
            strRuleSet_value = "BLOCOUT40EDGESPRING-AUTO"
        End If


        ' --- Setting parameters ---
        param_OLFR_OrderNo = New SqlParameter("@OrderNo_OLFR" & strSourceLineID, strOrderNo_value)
        param_OLFR_ItemCode = New SqlParameter("@ItemCode_OLFR" & strSourceLineID, strItemCode_value)
        param_OLFR_ItemDescription = New SqlParameter("@ItemDescription_OLFR" & strSourceLineID, strItemDescription_value)
        param_OLFR_ItemPrice = New SqlParameter("@ItemPrice_OLFR" & strSourceLineID, decItemPrice_value)
        param_OLFR_ItemQty = New SqlParameter("@ItemQty_OLFR" & strSourceLineID, intItemQty_value)
        param_OLFR_A = New SqlParameter("@A_OLFR" & strSourceLineID, intA_value)
        param_OLFR_B = New SqlParameter("@B_OLFR" & strSourceLineID, intB_value)
        param_OLFR_C = New SqlParameter("@C_OLFR" & strSourceLineID, intC_value)
        param_OLFR_D = New SqlParameter("@D_OLFR" & strSourceLineID, intD_value)
        param_OLFR_E = New SqlParameter("@E_OLFR" & strSourceLineID, intE_value)
        param_OLFR_F = New SqlParameter("@F_OLFR" & strSourceLineID, intF_value)
        param_OLFR_G = New SqlParameter("@G_OLFR" & strSourceLineID, intG_value)
        param_OLFR_AluColour = New SqlParameter("@AluColour_OLFR" & strSourceLineID, strAluColour_value)
        param_OLFR_FabColour = New SqlParameter("@FabColour_OLFR" & strSourceLineID, strFabColour_value)
        param_OLFR_FabColour2 = New SqlParameter("@FabColour2_OLFR" & strSourceLineID, strFabColour2_value)
        param_OLFR_BracketColour = New SqlParameter("@BracketColour_OLFR" & strSourceLineID, strBracketColour_value)
        param_OLFR_KeepPricingForInvoicing = New SqlParameter("@KeepPricingForInvoicing_OLFR" & strSourceLineID, strKeepPricingForInvoicing_value)
        param_OLFR_CustomerReferenceNo = New SqlParameter("@CustomerReferenceNo_OLFR" & strSourceLineID, strCustomerReferenceNo_value)
        param_OLFR_LineType = New SqlParameter("@LineType_OLFR" & strSourceLineID, strLineType_value)
        param_OLFR_HandleType = New SqlParameter("@HandleType_OLFR" & strSourceLineID, strHandleType_value)
        param_OLFR_ChainType = New SqlParameter("@ChainType_OLFR" & strSourceLineID, strChainType_value)
        param_OLFR_ChainSide = New SqlParameter("@ChainSide_OLFR" & strSourceLineID, strChainSide_value)
        param_OLFR_FabricRoleDir = New SqlParameter("@FabricRoleDir_OLFR" & strSourceLineID, strFabricRoleDir_value)
        param_OLFR_FittingType = New SqlParameter("@FittingType_OLFR" & strSourceLineID, strFittingType_value)
        param_OLFR_InstallationHeight = New SqlParameter("@InstallationHeight_OLFR" & strSourceLineID, strInstallationHeight_value)
        param_OLFR_ProductRange = New SqlParameter("@ProductRange_OLFR" & strSourceLineID, strProductRange_value)
        param_OLFR_ProductName = New SqlParameter("@ProductName_OLFR" & strSourceLineID, strProductName_value)
        param_OLFR_InvoiceUnitPrice = New SqlParameter("@InvoiceUnitPrice_OLFR" & strSourceLineID, decItemPrice_value)
        param_OLFR_Operation = New SqlParameter("@Operation_OLFR" & strSourceLineID, strOperation_value)
        param_OLFR_RuleSet = New SqlParameter("@RuleSet_OLFR" & strSourceLineID, strRuleSet_value)


        ' --- Adding parameters ---
        cmd.Parameters.Add(param_OLFR_OrderNo)
        cmd.Parameters.Add(param_OLFR_ItemCode)
        cmd.Parameters.Add(param_OLFR_ItemDescription)
        cmd.Parameters.Add(param_OLFR_ItemPrice)
        cmd.Parameters.Add(param_OLFR_ItemQty)
        cmd.Parameters.Add(param_OLFR_A)
        cmd.Parameters.Add(param_OLFR_B)
        cmd.Parameters.Add(param_OLFR_C)
        cmd.Parameters.Add(param_OLFR_D)
        cmd.Parameters.Add(param_OLFR_E)
        cmd.Parameters.Add(param_OLFR_F)
        cmd.Parameters.Add(param_OLFR_G)
        cmd.Parameters.Add(param_OLFR_AluColour)
        cmd.Parameters.Add(param_OLFR_FabColour)
        cmd.Parameters.Add(param_OLFR_FabColour2)
        cmd.Parameters.Add(param_OLFR_BracketColour)
        cmd.Parameters.Add(param_OLFR_KeepPricingForInvoicing)
        cmd.Parameters.Add(param_OLFR_CustomerReferenceNo)
        cmd.Parameters.Add(param_OLFR_LineType)
        cmd.Parameters.Add(param_OLFR_HandleType)
        cmd.Parameters.Add(param_OLFR_ChainType)
        cmd.Parameters.Add(param_OLFR_ChainSide)
        cmd.Parameters.Add(param_OLFR_FabricRoleDir)
        cmd.Parameters.Add(param_OLFR_FittingType)
        cmd.Parameters.Add(param_OLFR_InstallationHeight)
        cmd.Parameters.Add(param_OLFR_ProductRange)
        cmd.Parameters.Add(param_OLFR_ProductName)
        cmd.Parameters.Add(param_OLFR_InvoiceUnitPrice)
        cmd.Parameters.Add(param_OLFR_Operation)
        cmd.Parameters.Add(param_OLFR_RuleSet)


        ' --- Adding SQL insert statement to command text ---
        ' Ensure the text "/*ORD_LINE_BLOCKS*/" is persisted, as other functions will try to add more lines into
        ' the command text at this point
        cmd.CommandText = cmd.CommandText.Replace("/*ORD_LINE_BLOCKS*/", strProductLineQuery & vbNewLine & "/*ORD_LINE_BLOCKS*/")


    End Sub

    ''' <summary>
    ''' This function adds the INSERT statement and parameters to create a measure protect line in the Production System.
    ''' </summary>
    ''' <param name="cmd">The full SQL transaction command to add a order lines insert statements / parameters to</param>
    ''' <param name="drOrderHeader">The ORD_HEADER record to create a measure protect line for</param>
    ''' <param name="strWebsite">The location code of the source website. E.g. UK</param>
    ''' <remarks>
    ''' Author: Huw Day - 06/02/2018
    ''' The command is passed in by reference. Any changes will be available to the calling function without
    ''' the command needing to be passed back.
    ''' </remarks>
    Sub Build_OrderLine_MeasureProtect(ByRef cmd As SqlCommand, drOrderHeader As DataRow, strWebsite As String, decMeasureProtectValue As Decimal)


        ' --- Defining objects and variables ---
        Dim intVAT As Integer = CInt(Importer.Get_SettingFromWebsite("VAT", strWebsite))
        Dim strOrderMeasureProtectLine As String
        Dim param_MP_OrderNo As SqlParameter
        Dim param_MP_ItemPrice As SqlParameter


        ' --- Performing calculations on the Measure Protect value ---

        ' ## The value decMeasureProtectValue is now supplied without tax
        'decMeasureProtectValue = (decMeasureProtectValue / decCurrencyConvert) / (1 + (intVAT / 100))       ' Converting Measure Protect value back to £, then removing VAT

        decMeasureProtectValue = Currency_ConvertToPounds(decMeasureProtectValue, strWebsite)               ' Converting Measure Protect value back to £
        decMeasureProtectValue = Math.Round(decMeasureProtectValue, 2, MidpointRounding.AwayFromZero)       ' Rounding to 2 decimal places


        ' --- Writing Measure Protect insert statement ---
        strOrderMeasureProtectLine = "INSERT INTO ORD_LINES (order_no, item_code, item_description, item_price, item_qty, " _
            & "ORD_HEADERS_ID, A, B, C, D, E, F, G, ALUCOLOUR, FABCOLOUR, KeepPricingForInvoicing, CustomerReferenceNo, " _
            & "LINETYPE, HandleType, ChainType, ChainSide, FabricRoleDir, FittingType, InstallationHeight, ProductRange, ProductName) " _
            & "VALUES (@OrderNo_MP, 'MEASUREPROTECT|NA|NA', 'Measure Protect - ', @ItemPrice_MP, '0', @ProdSysID, '0', '0', '0', '0', " _
            & "'0', '0', '0', 'NA', 'NA', 'True', '', 'MEASUREPROTECT', '', 'METAL', '','', '', '', '', ''); "


        ' --- Setting parameters ---
        param_MP_OrderNo = New SqlParameter("@OrderNo_MP", "BL" & strWebsite & drOrderHeader("Order_No").ToString)
        param_MP_ItemPrice = New SqlParameter("@ItemPrice_MP", decMeasureProtectValue)


        ' --- Adding parameters to command ---
        cmd.Parameters.Add(param_MP_OrderNo)
        cmd.Parameters.Add(param_MP_ItemPrice)


        ' --- Writing the SQL insert statement into the command's text ---
        ' Ensure the text "/*ORD_LINE_BLOCKS*/" is persisted, as other functions will try to add more lines into
        ' the command text at this point
        cmd.CommandText = cmd.CommandText.Replace("/*ORD_LINE_BLOCKS*/", strOrderMeasureProtectLine & vbNewLine & "/*ORD_LINE_BLOCKS*/")


        ' ## Command now has the necessary elements to insert the Measure Protect line ##

    End Sub

    ''' <summary>
    ''' This function adds the INSERT statement and parameters to import a single order line.
    ''' </summary>
    ''' <param name="cmd">The full SQL transaction command to add a order lines insert statements / parameters to</param>
    ''' <param name="drOrderLine">The ORD_LINE record to import</param>
    ''' <param name="strWebsite">The location code of the source website. E.g. UK</param>
    ''' <remarks>
    ''' Author: Huw Day - 06/02/2018
    ''' The command is passed in by reference. Any changes will be available to the calling function without
    ''' the command needing to be passed back.
    ''' </remarks>
    Sub Build_OrderLine_Product(ByRef cmd As SqlCommand, drOrderLine As DataRow, strWebsite As String)


        ' --- Defining objects and variables ---
        Dim strProductLineQuery As String
        Dim strSourceLineID As String = drOrderLine("LineID")
        Dim intVAT As Integer = CInt(Importer.Get_SettingFromWebsite("VAT", strWebsite))
        Dim blnIsMotorised As Boolean = 0


        ' --- Defining parameters ---
        Dim param_OL_OrderNo As SqlParameter
        Dim param_OL_ItemCode As SqlParameter
        Dim param_OL_ItemDescription As SqlParameter
        Dim param_OL_ItemPrice As SqlParameter
        Dim param_OL_ItemQty As SqlParameter
        Dim param_OL_A As SqlParameter
        Dim param_OL_B As SqlParameter
        Dim param_OL_C As SqlParameter
        Dim param_OL_D As SqlParameter
        Dim param_OL_E As SqlParameter
        Dim param_OL_F As SqlParameter
        Dim param_OL_G As SqlParameter
        Dim param_OL_AluColour As SqlParameter
        Dim param_OL_FabColour As SqlParameter
        Dim param_OL_FabColour2 As SqlParameter
        Dim param_OL_BracketColour As SqlParameter
        Dim param_OL_KeepPricingForInvoicing As SqlParameter
        Dim param_OL_CustomerReferenceNo As SqlParameter
        Dim param_OL_LineType As SqlParameter
        Dim param_OL_HandleType As SqlParameter
        Dim param_OL_ChainType As SqlParameter
        Dim param_OL_ChainSide As SqlParameter
        Dim param_OL_FabricRoleDir As SqlParameter
        Dim param_OL_FittingType As SqlParameter
        Dim param_OL_InstallationHeight As SqlParameter
        Dim param_OL_ProductRange As SqlParameter
        Dim param_OL_ProductName As SqlParameter
        Dim param_OL_InvoiceUnitPrice As SqlParameter
        Dim param_OL_Operation As SqlParameter
        Dim param_OL_RuleSet As SqlParameter


        ' --- Defining parameter content variables ---
        Dim strOrderNo_value As String
        Dim strItemCode_value As String
        Dim strItemDescription_value As String
        Dim decItemPrice_value As Decimal = 0               ' Defaults to 0 if no valid price could be found
        Dim decMeasureProtectValue As Decimal = 0
        Dim intItemQty_value As Integer = 1                 ' Defaults to 1 if no valid quantity could be found
        Dim intA_value As Integer = 0
        Dim intB_value As Integer = 0
        Dim intC_value As Integer = 0
        Dim intD_value As Integer = 0
        Dim intE_value As Integer = 0
        Dim intF_value As Integer = 0
        Dim intG_value As Integer = 0
        Dim strAluColour_value As String = ""
        Dim strFabColour_value As String = ""
        Dim strFabColour2_value As String = ""
        Dim strBracketColour_value As String = ""
        Dim strKeepPricingForInvoicing_value As String = "True"     ' Always True
        Dim strCustomerReferenceNo_value As String = ""
        Dim strLineType_value As String
        Dim strHandleType_value As String
        Dim strChainType_value As String
        Dim strChainSide_value As String = "Left"                   ' defaults to Left
        Dim strFabricRoleDir_value As String = "Standard"           ' defaults to Standard
        Dim strFittingType_value As String = ""
        Dim strInstallationHeight_value As String = ""              ' Always empty
        Dim strProductRange_value As String
        Dim strProductName_value As String = ""
        Dim strOperation_value As String = ""
        Dim strRuleSet_value As String = ""


        ' --- Checking if product is motorised ---
        If drOrderLine.Table.Columns.Contains("Motorised") Then
            If IsDBNull(drOrderLine("Motorised")) = False Then
                blnIsMotorised = CBool(drOrderLine("Motorised"))
            End If
        End If


        ' --- Writing Product order line insert statement ---
        ' As the below SQL statement could be loaded multiple times into the transaction command, the parameter names
        ' have to be unique to the order line. The LineID from the website's order line table is appended to each
        ' parameter name to help ensure that each parameter in the command is unique.
        strProductLineQuery = "INSERT INTO ORD_LINES (order_no, item_code, item_description, item_price, item_qty, " _
            & "ORD_HEADERS_ID, A, B, C, D, E, F, G, ALUCOLOUR, FABCOLOUR, FABCOLOUR2, BracketColour, KeepPricingForInvoicing, " _
            & "CustomerReferenceNo, LINETYPE, HandleType, ChainType, ChainSide, FabricRoleDir, FittingType, " _
            & "InstallationHeight, ProductRange, ProductName, InvoiceUnitPrice, Operation, RuleSet) " _
            & "VALUES (" _
            & "@OrderNo_OL" & strSourceLineID _
            & ", @ItemCode_OL" & strSourceLineID _
            & ", @ItemDescription_OL" & strSourceLineID _
            & ", @ItemPrice_OL" & strSourceLineID _
            & ", @ItemQty_OL" & strSourceLineID _
            & ", @ProdSysID" _
            & ", @A_OL" & strSourceLineID _
            & ", @B_OL" & strSourceLineID _
            & ", @C_OL" & strSourceLineID _
            & ", @D_OL" & strSourceLineID _
            & ", @E_OL" & strSourceLineID _
            & ", @F_OL" & strSourceLineID _
            & ", @G_OL" & strSourceLineID _
            & ", @AluColour_OL" & strSourceLineID _
            & ", @FabColour_OL" & strSourceLineID _
            & ", @FabColour2_OL" & strSourceLineID _
            & ", @BracketColour_OL" & strSourceLineID _
            & ", @KeepPricingForInvoicing_OL" & strSourceLineID _
            & ", @CustomerReferenceNo_OL" & strSourceLineID _
            & ", @LineType_OL" & strSourceLineID _
            & ", @HandleType_OL" & strSourceLineID _
            & ", @ChainType_OL" & strSourceLineID _
            & ", @ChainSide_OL" & strSourceLineID _
            & ", @FabricRoleDir_OL" & strSourceLineID _
            & ", @FittingType_OL" & strSourceLineID _
            & ", @InstallationHeight_OL" & strSourceLineID _
            & ", @ProductRange_OL" & strSourceLineID _
            & ", @ProductName_OL" & strSourceLineID _
            & ", @InvoiceUnitPrice_OL" & strSourceLineID _
            & ", @Operation_OL" & strSourceLineID _
            & ", @RuleSet_OL" & strSourceLineID _
            & ")"
        ' Note: @ProdSysID is not a vb SqlParameter, it is set as a SQL variable in the main transaction block. See Import_Orders function


        ' --- Setting parameter content variables ---
        ' -- order_no --
        strOrderNo_value = "BL" & strWebsite & drOrderLine("Order_No").ToString()

        ' -- item_code --
        strItemCode_value = Importer.Get_ItemCode(drOrderLine, strWebsite)

        ' -- item_description --
        strItemDescription_value = Importer.Get_ItemDescription(drOrderLine)

        ' -- Quantity --
        If drOrderLine.Table.Columns.Contains("Quantity") Then
            If IsDBNull(drOrderLine("Quantity")) = False Then
                If IsNumeric(drOrderLine("Quantity")) Then
                    intItemQty_value = CInt(drOrderLine("Quantity"))
                End If
            End If
        End If
        If drOrderLine("ProductName").ToString.ToLower = "faceshieldkit_4b24" Then      ' Override for FaceShieldKit_4b24   (as a single of this product equates to 4 boxes)
            If IsDBNull(drOrderLine("Quantity")) = False Then
                If IsNumeric(drOrderLine("Quantity")) Then
                    intItemQty_value = CInt(drOrderLine("Quantity")) * 4
                End If
            End If
        End If

        ' -- ItemPrice --
        If drOrderLine.Table.Columns.Contains("Price_Final_ExTax") Then
            If IsNumeric(drOrderLine("Price_Final_ExTax")) Then
                decItemPrice_value = CDec(drOrderLine("Price_Final_ExTax"))
            End If
        End If
        If drOrderLine.Table.Columns.Contains("MeasureProtectValue") Then
            If IsNumeric(drOrderLine("MeasureProtectValue")) Then
                decMeasureProtectValue = CDec(drOrderLine("MeasureProtectValue"))
            End If
        End If
        decItemPrice_value -= decMeasureProtectValue                                                ' Removing the Measure Protect value from the overall cost of the item
        decItemPrice_value = Currency_ConvertToPounds(decItemPrice_value, strWebsite)               ' Converting item price back to £
        decItemPrice_value = decItemPrice_value / intItemQty_value                                  ' Converting Line Total price to price per unit, for the benefit of the production system
        decItemPrice_value = Math.Round(decItemPrice_value, 2, MidpointRounding.AwayFromZero)       ' Rounding to 2 decimal places

        ' -- A --
        If drOrderLine("ProductName") = "CustomSkylight" Then
            ' If the product is CustomSkylight, then A matches up with FrameWidth
            If drOrderLine.Table.Columns.Contains("FrameWidth") Then
                If IsDBNull(drOrderLine("FrameWidth")) = False Then
                    If IsNumeric(drOrderLine("FrameWidth")) Then
                        intA_value = CInt(drOrderLine("FrameWidth"))
                    End If
                End If
            End If
        Else
            ' For all other products, A matches up with Width
            If drOrderLine.Table.Columns.Contains("Width") Then
                If IsDBNull(drOrderLine("Width")) = False Then
                    If IsNumeric(drOrderLine("Width")) Then
                        intA_value = CInt(drOrderLine("Width"))
                    End If
                End If
            End If
        End If

        ' -- B --
        If drOrderLine("ProductName") = "CustomSkylight" Then
            ' If the product is CustomSkylight, then B matches up with GlassWidth
            If drOrderLine.Table.Columns.Contains("GlassWidth") Then
                If IsDBNull(drOrderLine("GlassWidth")) = False Then
                    If IsNumeric(drOrderLine("GlassWidth")) Then
                        intB_value = CInt(drOrderLine("GlassWidth"))
                    End If
                End If
            End If
        Else
            ' For all other products, B matches up with WidthMiddle
            If drOrderLine.Table.Columns.Contains("WidthMiddle") Then
                If IsDBNull(drOrderLine("WidthMiddle")) = False Then
                    If IsNumeric(drOrderLine("WidthMiddle")) Then
                        intB_value = CInt(drOrderLine("WidthMiddle"))
                    End If
                End If
            End If
        End If

        ' -- C --
        If drOrderLine("ProductName") = "CustomSkylight" Then
            ' If the product is CustomSkylight, then C matches up with FrameDrop
            If drOrderLine.Table.Columns.Contains("FrameDrop") Then
                If IsDBNull(drOrderLine("FrameDrop")) = False Then
                    If IsNumeric(drOrderLine("FrameDrop")) Then
                        intC_value = CInt(drOrderLine("FrameDrop"))
                    End If
                End If
            End If
        Else
            ' For all other products, C matches up with WidthBottom
            If drOrderLine.Table.Columns.Contains("WidthBottom") Then
                If IsDBNull(drOrderLine("WidthBottom")) = False Then
                    If IsNumeric(drOrderLine("WidthBottom")) Then
                        intC_value = CInt(drOrderLine("WidthBottom"))
                    End If
                End If
            End If
        End If

        ' -- D, E --
        If drOrderLine("ProductName") = "CustomSkylight" Then
            ' If the product is CustomSkylight, then D matches up with Glass Drop and E matches up with Depth of Frame
            If drOrderLine.Table.Columns.Contains("GlassDrop") Then
                If IsDBNull(drOrderLine("GlassDrop")) = False Then
                    If IsNumeric(drOrderLine("GlassDrop")) Then
                        intD_value = CInt(drOrderLine("GlassDrop"))
                    End If
                End If
            End If
            If drOrderLine.Table.Columns.Contains("FrameDepth") Then
                If IsDBNull(drOrderLine("FrameDepth")) = False Then
                    If IsNumeric(drOrderLine("FrameDepth")) Then
                        intE_value = CInt(drOrderLine("FrameDepth"))
                    End If
                End If
            End If
        Else
            ' If the product just has a Height measurement (HeightLeft and HeightRight are null), Height should go here. 
            ' Otherwise, HeightLeft should go here and Height should go in E
            If drOrderLine.Table.Columns.Contains("Height") Then
                If IsDBNull(drOrderLine("Height")) = False Then
                    If IsNumeric(drOrderLine("Height")) Then
                        intD_value = CInt(drOrderLine("Height"))
                    End If
                End If
            End If
            If drOrderLine.Table.Columns.Contains("HeightLeft") Then
                If IsDBNull(drOrderLine("HeightLeft")) = False Then
                    If IsNumeric(drOrderLine("HeightLeft")) Then
                        intE_value = intD_value
                        intD_value = CInt(drOrderLine("HeightLeft"))
                    End If
                End If
            End If

        End If

        ' -- F --
        If drOrderLine("ProductName") = "CustomSkylight" Then
            ' If the product is CustomSkylight, then F matches up with Depth Of Frame (again)
            If drOrderLine.Table.Columns.Contains("FrameDepth") Then
                If IsDBNull(drOrderLine("FrameDepth")) = False Then
                    If IsNumeric(drOrderLine("FrameDepth")) Then
                        intF_value = CInt(drOrderLine("FrameDepth"))
                    End If
                End If
            End If
        Else
            ' Otherwise, F matches up with HeightRight
            If drOrderLine.Table.Columns.Contains("HeightRight") Then
                If IsDBNull(drOrderLine("HeightRight")) = False Then
                    If IsNumeric(drOrderLine("HeightRight")) Then
                        intF_value = CInt(drOrderLine("HeightRight"))
                    End If
                End If
            End If
        End If

        ' -- G --
        If drOrderLine.Table.Columns.Contains("RecessDepth") Then
            If IsDBNull(drOrderLine("RecessDepth")) = False Then
                If IsNumeric(drOrderLine("RecessDepth")) Then
                    intG_value = CInt(drOrderLine("RecessDepth"))
                End If
            End If
        End If

        ' -- ALUCOLOUR --
        If drOrderLine.Table.Columns.Contains("AluminiumColour") Then
            If IsDBNull(drOrderLine("AluminiumColour")) = False Then
                strAluColour_value = drOrderLine("AluminiumColour").ToString
            End If
        End If
        If drOrderLine("ProductName").ToString().ToLower() = "rollerblindextend_black" _
        Or drOrderLine("ProductName").ToString().ToLower() = "rollerblindextend_blue" _
        Or drOrderLine("ProductName").ToString().ToLower() = "rollerblindextend_red" _
        Or drOrderLine("ProductName").ToString().ToLower() = "rollerblindextend_green" Then
            ' Override default aluminium colour for rollerblind extend
            strAluColour_value = "NA"
        End If

        ' -- FABCOLOUR --
        If drOrderLine.Table.Columns.Contains("FabricCode") Then
            If IsDBNull(drOrderLine("FabricCode")) = False Then
                strFabColour_value = drOrderLine("FabricCode").ToString

                ' Translate duplicated fabric colours back to original values
                strFabColour_value = Importer.Get_OriginalFabricCode(strFabColour_value, strWebsite)
            Else
                If drOrderLine("ProductName").ToString().ToLower() = "rollerblindextend_black" _
                Or drOrderLine("ProductName").ToString().ToLower() = "rollerblindextend_blue" _
                Or drOrderLine("ProductName").ToString().ToLower() = "rollerblindextend_red" _
                Or drOrderLine("ProductName").ToString().ToLower() = "rollerblindextend_green" Then
                    strFabColour_value = "KIT"
                End If
            End If
        End If

        ' -- FABCOLOUR2 --
        If drOrderLine.Table.Columns.Contains("InnerFabricCode") Then
            If IsDBNull(drOrderLine("InnerFabricCode")) = False Then
                strFabColour2_value = drOrderLine("InnerFabricCode").ToString

                ' Translate duplicated fabric colours back to original values
                strFabColour2_value = Importer.Get_OriginalFabricCode(strFabColour2_value, strWebsite)
            End If
        End If

        ' -- BracketColour --
        If drOrderLine.Table.Columns.Contains("BracketColour") Then
            If IsDBNull(drOrderLine("BracketColour")) = False Then
                strBracketColour_value = drOrderLine("BracketColour").ToString
            End If
        End If

        ' -- CustomerReferenceNo --
        If drOrderLine.Table.Columns.Contains("CustomerReference") Then
            If IsDBNull(drOrderLine("CustomerReference")) = False Then
                strCustomerReferenceNo_value = drOrderLine("CustomerReference").ToString
            End If
        End If

        ' -- LINETYPE --
        strLineType_value = lst_ProdSys_LineType(drOrderLine("ProductName").ToString.ToLower)
        If blnIsMotorised And drOrderLine("ProductName").ToString.Contains("SolarSkylight") = False Then strLineType_value = strLineType_value + "-MOTORISED"

        ' -- HandleType --
        Select Case drOrderLine("ProductName").ToString
            Case "PremiumRollerBlind", "RollerBlind", "RollerBlindDrillFree", "RollerBlindMotorised", "TwinRollerBlind"
                strHandleType_value = "WHITE PVC BOTTOM BAR"
            Case Else
                strHandleType_value = ""
        End Select

        ' -- ChainType --
        If drOrderLine("ProductName").ToString().ToLower() = "blocout40" _
        Or drOrderLine("ProductName").ToString().ToLower() = "venetian" _
        Or drOrderLine("ProductName").ToString().ToLower() = "smartmotorcharger" _
        Or drOrderLine("ProductName").ToString().ToLower() = "smartremote" _
        Or drOrderLine("ProductName").ToString().ToLower() = "smarthub" _
        Or drOrderLine("ProductName").ToString().ToLower() = "smartrepeater" _
        Or drOrderLine("ProductName").ToString().ToLower() = "faceshield_1b12" _
        Or drOrderLine("ProductName").ToString().ToLower() = "faceshield_5b12" _
        Or drOrderLine("ProductName").ToString().ToLower() = "faceshield_20b12" _
        Or drOrderLine("ProductName").ToString().ToLower() = "faceshieldkit_1b1" _
        Or drOrderLine("ProductName").ToString().ToLower() = "faceshieldkit_1b2" _
        Or drOrderLine("ProductName").ToString().ToLower() = "faceshieldkit_1b3" _
        Or drOrderLine("ProductName").ToString().ToLower() = "faceshieldkit_1b4" _
        Or drOrderLine("ProductName").ToString().ToLower() = "faceshieldkit_1b5" _
        Or drOrderLine("ProductName").ToString().ToLower() = "faceshieldkit_1b12" _
        Or drOrderLine("ProductName").ToString().ToLower() = "faceshieldkit_1b24" _
        Or drOrderLine("ProductName").ToString().ToLower() = "faceshieldkit_1b40" _
        Or drOrderLine("ProductName").ToString().ToLower() = "faceshieldkit_4b24" _
        Or drOrderLine("ProductName").ToString().ToLower() = "BO40_Brakes_Cream" _
        Or drOrderLine("ProductName").ToString().ToLower() = "BO40_Brakes_Grey" _
        Or drOrderLine("ProductName").ToString().ToLower() = "BO40_Brakes_White" _
        Or drOrderLine("ProductName").ToString().ToLower() = "BO40_Fittings" _
        Or drOrderLine("ProductName").ToString().ToLower() = "BO80_Fittings_Recess" _
        Or drOrderLine("ProductName").ToString().ToLower() = "BO80_Fittings_Surface" _
        Or drOrderLine("ProductName").ToString().ToLower() = "PRB_Fittings_Recess" _
        Or drOrderLine("ProductName").ToString().ToLower() = "PRB_Fittings_Surface" _
        Or drOrderLine("ProductName").ToString().ToLower() = "RB_Brackets_Black" _
        Or drOrderLine("ProductName").ToString().ToLower() = "RB_Brackets_Grey" _
        Or drOrderLine("ProductName").ToString().ToLower() = "RB_Brackets_White" _
        Or drOrderLine("ProductName").ToString().ToLower() = "RB_DF_Brackets_White" _
        Or drOrderLine("ProductName").ToString().ToLower() = "RB_M_Brackets_Black" _
        Or drOrderLine("ProductName").ToString().ToLower() = "RB_M_Brackets_White" _
        Or drOrderLine("ProductName").ToString().ToLower() = "SKY_Brakes_Cream" _
        Or drOrderLine("ProductName").ToString().ToLower() = "SKY_Brakes_Grey" _
        Or drOrderLine("ProductName").ToString().ToLower() = "SKY_Brakes_White" _
        Or drOrderLine("ProductName").ToString().ToLower() = "SS_Battery_Post_03_19" _
        Or drOrderLine("ProductName").ToString().ToLower() = "SS_Battery_Pre_03_19" _
        Or drOrderLine("ProductName").ToString().ToLower() = "ZRB_Fittings_Recess" _
        Or drOrderLine("ProductName").ToString().ToLower() = "ZRB_Fittings_Surface" Then
            strChainType_value = ""
        ElseIf blnIsMotorised = True Then
            strChainType_value = "MOTORISED"
        Else
            strChainType_value = "METAL"
        End If

        ' -- ChainSide --
        If drOrderLine.Table.Columns.Contains("ChainSide") Then
            If IsDBNull(drOrderLine("ChainSide")) = False Then
                strChainSide_value = drOrderLine("ChainSide").ToString
            End If
        End If

        ' -- FabricRoleDirection --
        If drOrderLine.Table.Columns.Contains("RollDirection") Then
            If IsDBNull(drOrderLine("RollDirection")) = False Then
                strFabricRoleDir_value = drOrderLine("RollDirection").ToString
            End If
        End If

        ' -- FittingType --
        If drOrderLine.Table.Columns.Contains("FittingType") Then
            If IsDBNull(drOrderLine("FittingType")) = False Then
                If lst_ProdSys_FittingTypes.ContainsKey(drOrderLine("FittingType").ToString().ToLower()) Then
                    strFittingType_value = lst_ProdSys_FittingTypes(drOrderLine("FittingType").ToString.ToLower)
                End If
            End If
        End If

        ' -- ProductRange --
        If blnIsMotorised Then strProductRange_value = "BlocMotorised" Else strProductRange_value = "Bloc"

        ' -- ProductName --
        If drOrderLine.Table.Columns.Contains("ProductName") Then
            If IsDBNull(drOrderLine("ProductName")) = False Then
                strProductName_value = lst_ProdSys_ProductNames(drOrderLine("ProductName").ToString.ToLower)

                ' Remember to replace the <manufacturer> tag, if it is present
                If drOrderLine.Table.Columns.Contains("Manufacturer") Then
                    If IsDBNull(drOrderLine("Manufacturer")) = False Then
                        strProductName_value = strProductName_value.Replace("<manufacturer>", drOrderLine("Manufacturer").ToString.ToLower)
                    End If
                End If

            End If
        End If


        ' -- Operation --
        If blnIsMotorised = True Then
            strOperation_value = "MOTORISED"
        Else
            strOperation_value = ""
            If drOrderLine.Table.Columns.Contains("DriveType") Then
                If IsDBNull(drOrderLine("DriveType")) = False Then
                    If drOrderLine("DriveType").ToString().ToUpper().Contains("WAND") Then
                        strOperation_value = drOrderLine("DriveType").ToString().ToUpper()
                    End If
                End If
            End If
        End If


        ' -- RuleSet --
        ' check based on already calculated production system values of ProductName and Fitting Type
        If strProductName_value.ToLower = "blocout40" And strFittingType_value.ToLower = "Edge of Window Recess" Then
            strRuleSet_value = "BLOCOUT40EDGESPRING-AUTO"
        End If


        ' --- Setting parameters ---
        param_OL_OrderNo = New SqlParameter("@OrderNo_OL" & strSourceLineID, strOrderNo_value)
        param_OL_ItemCode = New SqlParameter("@ItemCode_OL" & strSourceLineID, strItemCode_value)
        param_OL_ItemDescription = New SqlParameter("@ItemDescription_OL" & strSourceLineID, strItemDescription_value)
        param_OL_ItemPrice = New SqlParameter("@ItemPrice_OL" & strSourceLineID, decItemPrice_value)
        param_OL_ItemQty = New SqlParameter("@ItemQty_OL" & strSourceLineID, intItemQty_value)
        param_OL_A = New SqlParameter("@A_OL" & strSourceLineID, intA_value)
        param_OL_B = New SqlParameter("@B_OL" & strSourceLineID, intB_value)
        param_OL_C = New SqlParameter("@C_OL" & strSourceLineID, intC_value)
        param_OL_D = New SqlParameter("@D_OL" & strSourceLineID, intD_value)
        param_OL_E = New SqlParameter("@E_OL" & strSourceLineID, intE_value)
        param_OL_F = New SqlParameter("@F_OL" & strSourceLineID, intF_value)
        param_OL_G = New SqlParameter("@G_OL" & strSourceLineID, intG_value)
        param_OL_AluColour = New SqlParameter("@AluColour_OL" & strSourceLineID, strAluColour_value)
        param_OL_FabColour = New SqlParameter("@FabColour_OL" & strSourceLineID, strFabColour_value)
        param_OL_FabColour2 = New SqlParameter("@FabColour2_OL" & strSourceLineID, strFabColour2_value)
        param_OL_BracketColour = New SqlParameter("@BracketColour_OL" & strSourceLineID, strBracketColour_value)
        param_OL_KeepPricingForInvoicing = New SqlParameter("@KeepPricingForInvoicing_OL" & strSourceLineID, strKeepPricingForInvoicing_value)
        param_OL_CustomerReferenceNo = New SqlParameter("@CustomerReferenceNo_OL" & strSourceLineID, strCustomerReferenceNo_value)
        param_OL_LineType = New SqlParameter("@LineType_OL" & strSourceLineID, strLineType_value)
        param_OL_HandleType = New SqlParameter("@HandleType_OL" & strSourceLineID, strHandleType_value)
        param_OL_ChainType = New SqlParameter("@ChainType_OL" & strSourceLineID, strChainType_value)
        param_OL_ChainSide = New SqlParameter("@ChainSide_OL" & strSourceLineID, strChainSide_value)
        param_OL_FabricRoleDir = New SqlParameter("@FabricRoleDir_OL" & strSourceLineID, strFabricRoleDir_value)
        param_OL_FittingType = New SqlParameter("@FittingType_OL" & strSourceLineID, strFittingType_value)
        param_OL_InstallationHeight = New SqlParameter("@InstallationHeight_OL" & strSourceLineID, strInstallationHeight_value)
        param_OL_ProductRange = New SqlParameter("@ProductRange_OL" & strSourceLineID, strProductRange_value)
        param_OL_ProductName = New SqlParameter("@ProductName_OL" & strSourceLineID, strProductName_value)
        param_OL_InvoiceUnitPrice = New SqlParameter("@InvoiceUnitPrice_OL" & strSourceLineID, decItemPrice_value)
        param_OL_Operation = New SqlParameter("@Operation_OL" & strSourceLineID, strOperation_value)
        param_OL_RuleSet = New SqlParameter("@RuleSet_OL" & strSourceLineID, strRuleSet_value)


        ' --- Adding parameters ---
        cmd.Parameters.Add(param_OL_OrderNo)
        cmd.Parameters.Add(param_OL_ItemCode)
        cmd.Parameters.Add(param_OL_ItemDescription)
        cmd.Parameters.Add(param_OL_ItemPrice)
        cmd.Parameters.Add(param_OL_ItemQty)
        cmd.Parameters.Add(param_OL_A)
        cmd.Parameters.Add(param_OL_B)
        cmd.Parameters.Add(param_OL_C)
        cmd.Parameters.Add(param_OL_D)
        cmd.Parameters.Add(param_OL_E)
        cmd.Parameters.Add(param_OL_F)
        cmd.Parameters.Add(param_OL_G)
        cmd.Parameters.Add(param_OL_AluColour)
        cmd.Parameters.Add(param_OL_FabColour)
        cmd.Parameters.Add(param_OL_FabColour2)
        cmd.Parameters.Add(param_OL_BracketColour)
        cmd.Parameters.Add(param_OL_KeepPricingForInvoicing)
        cmd.Parameters.Add(param_OL_CustomerReferenceNo)
        cmd.Parameters.Add(param_OL_LineType)
        cmd.Parameters.Add(param_OL_HandleType)
        cmd.Parameters.Add(param_OL_ChainType)
        cmd.Parameters.Add(param_OL_ChainSide)
        cmd.Parameters.Add(param_OL_FabricRoleDir)
        cmd.Parameters.Add(param_OL_FittingType)
        cmd.Parameters.Add(param_OL_InstallationHeight)
        cmd.Parameters.Add(param_OL_ProductRange)
        cmd.Parameters.Add(param_OL_ProductName)
        cmd.Parameters.Add(param_OL_InvoiceUnitPrice)
        cmd.Parameters.Add(param_OL_Operation)
        cmd.Parameters.Add(param_OL_RuleSet)


        ' --- Adding SQL insert statement to command text ---
        ' Ensure the text "/*ORD_LINE_BLOCKS*/" is persisted, as other functions will try to add more lines into
        ' the command text at this point
        cmd.CommandText = cmd.CommandText.Replace("/*ORD_LINE_BLOCKS*/", strProductLineQuery & vbNewLine & "/*ORD_LINE_BLOCKS*/")

    End Sub


    ''' <summary>
    ''' Adds the INSERT statement and parameters to import the second barrel of a twin roller blind. A copy of Build_OrderLine_Product,
    ''' with a few adjustments
    ''' </summary>
    ''' <param name="cmd">The full SQL transaction command to add a order lines insert statements / parameters to</param>
    ''' <param name="drOrderLine">The ORD_LINE record to import</param>
    ''' <param name="strWebsite">The location code of the source website. E.g. UK</param>
    ''' <remarks>
    ''' Author: Huw Day - 06/02/2018
    ''' The command is passed in by reference. Any changes will be available to the calling function without
    ''' the command needing to be passed back.
    ''' </remarks>
    Sub Build_OrderLine_TwinSecondRow(ByRef cmd As SqlCommand, drOrderLine As DataRow, strWebsite As String)


        ' --- Defining objects and variables ---
        Dim strProductLineQuery As String
        Dim strSourceLineID As String = drOrderLine("LineID")
        Dim blnIsMotorised As Boolean = 0


        ' --- Defining parameters ---
        Dim param_OLT2_OrderNo As SqlParameter
        Dim param_OLT2_ItemCode As SqlParameter
        Dim param_OLT2_ItemDescription As SqlParameter
        Dim param_OLT2_ItemPrice As SqlParameter
        Dim param_OLT2_ItemQty As SqlParameter
        Dim param_OLT2_A As SqlParameter
        Dim param_OLT2_B As SqlParameter
        Dim param_OLT2_C As SqlParameter
        Dim param_OLT2_D As SqlParameter
        Dim param_OLT2_E As SqlParameter
        Dim param_OLT2_F As SqlParameter
        Dim param_OLT2_G As SqlParameter
        Dim param_OLT2_AluColour As SqlParameter
        Dim param_OLT2_FabColour As SqlParameter
        Dim param_OLT2_FabColour2 As SqlParameter
        Dim param_OLT2_BracketColour As SqlParameter
        Dim param_OLT2_KeepPricingForInvoicing As SqlParameter
        Dim param_OLT2_CustomerReferenceNo As SqlParameter
        Dim param_OLT2_LineType As SqlParameter
        Dim param_OLT2_HandleType As SqlParameter
        Dim param_OLT2_ChainType As SqlParameter
        Dim param_OLT2_ChainSide As SqlParameter
        Dim param_OLT2_FabricRoleDir As SqlParameter
        Dim param_OLT2_FittingType As SqlParameter
        Dim param_OLT2_InstallationHeight As SqlParameter
        Dim param_OLT2_ProductRange As SqlParameter
        Dim param_OLT2_ProductName As SqlParameter
        Dim param_OLT2_InvoiceUnitPrice As SqlParameter
        Dim param_OLT2_Operation As SqlParameter
        Dim param_OLT2_RuleSet As SqlParameter


        ' --- Defining parameter content variables ---
        Dim strOrderNo_value As String
        Dim strItemCode_value As String
        Dim strItemDescription_value As String
        Dim decItemPrice_value As Decimal = 0               ' Defaults to 0 if no valid price could be found
        Dim intItemQty_value As Integer = 1                 ' Defaults to 1 if no valid quantity could be found
        Dim intA_value As Integer = 0
        Dim intB_value As Integer = 0
        Dim intC_value As Integer = 0
        Dim intD_value As Integer = 0
        Dim intE_value As Integer = 0
        Dim intF_value As Integer = 0
        Dim intG_value As Integer = 0
        Dim strAluColour_value As String = ""
        Dim strFabColour_value As String = ""
        Dim strFabColour2_value As String = ""
        Dim strBracketColour_value As String = ""
        Dim strKeepPricingForInvoicing_value As String = "True"     ' Always True
        Dim strCustomerReferenceNo_value As String = ""
        Dim strLineType_value As String
        Dim strHandleType_value As String
        Dim strChainType_value As String
        Dim strChainSide_value As String = "Left"                   ' defaults to Left
        Dim strFabricRoleDir_value As String = "Standard"           ' defaults to Standard
        Dim strFittingType_value As String = ""
        Dim strInstallationHeight_value As String = ""              ' Always empty
        Dim strProductRange_value As String
        Dim strProductName_value As String = ""
        Dim strOperation_value As String = ""
        Dim strRuleSet_value As String = ""


        ' --- Checking if product is motorised ---
        If drOrderLine.Table.Columns.Contains("Motorised") Then
            If IsDBNull(drOrderLine("Motorised")) = False Then
                blnIsMotorised = CBool(drOrderLine("Motorised"))
            End If
        End If


        ' --- Writing Product order line insert statement ---
        ' As the below SQL statement could be loaded multiple times into the transaction command, the parameter names
        ' have to be unique to the order line. The LineID from the website's order line table is appended to each
        ' parameter name to help ensure that each parameter in the command is unique.
        strProductLineQuery = "INSERT INTO ORD_LINES (order_no, item_code, item_description, item_price, item_qty, " _
            & "ORD_HEADERS_ID, A, B, C, D, E, F, G, ALUCOLOUR, FABCOLOUR, FABCOLOUR2, BracketColour, KeepPricingForInvoicing, " _
            & "CustomerReferenceNo, LINETYPE, HandleType, ChainType, ChainSide, FabricRoleDir, FittingType, " _
            & "InstallationHeight, ProductRange, ProductName, InvoiceUnitPrice, Operation, RuleSet) " _
            & "VALUES (" _
            & "@OrderNo_OLT2" & strSourceLineID _
            & ", @ItemCode_OLT2" & strSourceLineID _
            & ", @ItemDescription_OLT2" & strSourceLineID _
            & ", @ItemPrice_OLT2" & strSourceLineID _
            & ", @ItemQty_OLT2" & strSourceLineID _
            & ", @ProdSysID" _
            & ", @A_OLT2" & strSourceLineID _
            & ", @B_OLT2" & strSourceLineID _
            & ", @C_OLT2" & strSourceLineID _
            & ", @D_OLT2" & strSourceLineID _
            & ", @E_OLT2" & strSourceLineID _
            & ", @F_OLT2" & strSourceLineID _
            & ", @G_OLT2" & strSourceLineID _
            & ", @AluColour_OLT2" & strSourceLineID _
            & ", @FabColour_OLT2" & strSourceLineID _
            & ", @FabColour2_OLT2" & strSourceLineID _
            & ", @BracketColour_OLT2" & strSourceLineID _
            & ", @KeepPricingForInvoicing_OLT2" & strSourceLineID _
            & ", @CustomerReferenceNo_OLT2" & strSourceLineID _
            & ", @LineType_OLT2" & strSourceLineID _
            & ", @HandleType_OLT2" & strSourceLineID _
            & ", @ChainType_OLT2" & strSourceLineID _
            & ", @ChainSide_OLT2" & strSourceLineID _
            & ", @FabricRoleDir_OLT2" & strSourceLineID _
            & ", @FittingType_OLT2" & strSourceLineID _
            & ", @InstallationHeight_OLT2" & strSourceLineID _
            & ", @ProductRange_OLT2" & strSourceLineID _
            & ", @ProductName_OLT2" & strSourceLineID _
            & ", @InvoiceUnitPrice_OLT2" & strSourceLineID _
            & ", @Operation_OLT2" & strSourceLineID _
            & ", @RuleSet_OLT2" & strSourceLineID _
            & ")"
        ' Note: @ProdSysID is not a vb SqlParameter, it is set as a SQL variable in the main transaction block. See Import_Orders function


        ' --- Setting parameter content variables ---
        ' -- order_no --
        strOrderNo_value = "BL" & strWebsite & drOrderLine("Order_No").ToString()

        ' -- item_code --
        strItemCode_value = Importer.Get_ItemCode(drOrderLine, strWebsite)

        ' -- item_description --
        strItemDescription_value = Importer.Get_ItemDescription(drOrderLine)

        ' -- ItemPrice --
        decItemPrice_value = 0      ' The second Twin Roller order line has 0 price. The full price of the unit is attached to the primary product line
        decItemPrice_value = Currency_ConvertToPounds(decItemPrice_value, strWebsite)

        ' -- Quantity --
        If drOrderLine.Table.Columns.Contains("Quantity") Then
            If IsDBNull(drOrderLine("Quantity")) = False Then
                If IsNumeric(drOrderLine("Quantity")) Then
                    intItemQty_value = CInt(drOrderLine("Quantity"))
                End If
            End If
        End If

        ' -- A --
        If drOrderLine.Table.Columns.Contains("Width") Then
            If IsDBNull(drOrderLine("Width")) = False Then
                If IsNumeric(drOrderLine("Width")) Then
                    intA_value = CInt(drOrderLine("Width"))
                End If
            End If
        End If

        ' -- B --
        If drOrderLine.Table.Columns.Contains("WidthMiddle") Then
            If IsDBNull(drOrderLine("WidthMiddle")) = False Then
                If IsNumeric(drOrderLine("WidthMiddle")) Then
                    intB_value = CInt(drOrderLine("WidthMiddle"))
                End If
            End If
        End If

        ' -- C --
        If drOrderLine.Table.Columns.Contains("WidthBottom") Then
            If IsDBNull(drOrderLine("WidthBottom")) = False Then
                If IsNumeric(drOrderLine("WidthBottom")) Then
                    intC_value = CInt(drOrderLine("WidthBottom"))
                End If
            End If
        End If

        ' -- D, E --
        ' If the product just has a Height measurement (HeightLeft and HeightRight are null), Height should go here. 
        ' Otherwise, HeightLeft should go here and Height should go in E
        If drOrderLine.Table.Columns.Contains("Height") Then
            If IsDBNull(drOrderLine("Height")) = False Then
                If IsNumeric(drOrderLine("Height")) Then
                    intD_value = CInt(drOrderLine("Height"))
                End If
            End If
        End If
        If drOrderLine.Table.Columns.Contains("HeightLeft") Then
            If IsDBNull(drOrderLine("HeightLeft")) = False Then
                If IsNumeric(drOrderLine("HeightLeft")) Then
                    intE_value = intD_value
                    intD_value = CInt(drOrderLine("HeightLeft"))
                End If
            End If
        End If

        ' -- F --
        If drOrderLine.Table.Columns.Contains("HeightRight") Then
            If IsDBNull(drOrderLine("HeightRight")) = False Then
                If IsNumeric(drOrderLine("HeightRight")) Then
                    intF_value = CInt(drOrderLine("HeightRight"))
                End If
            End If
        End If

        ' -- G --
        If drOrderLine.Table.Columns.Contains("RecessDepth") Then
            If IsDBNull(drOrderLine("RecessDepth")) = False Then
                If IsNumeric(drOrderLine("RecessDepth")) Then
                    intG_value = CInt(drOrderLine("RecessDepth"))
                End If
            End If
        End If

        ' -- ALUCOLOUR --
        If drOrderLine.Table.Columns.Contains("AluminiumColour") Then
            If IsDBNull(drOrderLine("AluminiumColour")) = False Then
                strAluColour_value = drOrderLine("AluminiumColour").ToString
            End If
        End If

        ' -- FABCOLOUR --
        ' Use the inner fabric code for the second twin roller order line.
        If drOrderLine.Table.Columns.Contains("InnerFabricCode") Then
            If IsDBNull(drOrderLine("InnerFabricCode")) = False Then
                strFabColour_value = drOrderLine("InnerFabricCode").ToString

                ' Translate duplicated fabric colours back to original values
                strFabColour_value = Importer.Get_OriginalFabricCode(strFabColour_value, strWebsite)
            End If
        End If

        ' -- FABCOLOUR2 --
        strFabColour2_value = ""        ' Empty for the second Twin roller order line

        ' -- BracketColour --
        strBracketColour_value = "NOBRACKETS AS TWIN"

        ' -- CustomerReferenceNo --
        If drOrderLine.Table.Columns.Contains("CustomerReference") Then
            If IsDBNull(drOrderLine("CustomerReference")) = False Then
                strCustomerReferenceNo_value = drOrderLine("CustomerReference").ToString
            End If
        End If

        ' -- LINETYPE --
        strLineType_value = lst_ProdSys_LineType(drOrderLine("ProductName").ToString.ToLower)
        If blnIsMotorised Then strLineType_value = strLineType_value + "-MOTORISED"

        ' -- HandleType --
        Select Case drOrderLine("ProductName").ToString
            Case "PremiumRollerBlind", "RollerBlind", "RollerBlindDrillFree", "RollerBlindMotorised", "TwinRollerBlind"
                strHandleType_value = "WHITE PVC BOTTOM BAR"
            Case Else
                strHandleType_value = ""
        End Select

        ' -- ChainType --
        If drOrderLine("ProductName").ToString.ToLower = "blocout40" Then
            strChainType_value = ""
        ElseIf blnIsMotorised = True Then
            strChainType_value = "MOTORISED"
        Else
            strChainType_value = "METAL"
        End If

        ' -- ChainSide --
        If drOrderLine.Table.Columns.Contains("ChainSide") Then
            If IsDBNull(drOrderLine("ChainSide")) = False Then
                strChainSide_value = drOrderLine("ChainSide").ToString
            End If
        End If

        ' -- FabricRoleDirection --
        If drOrderLine.Table.Columns.Contains("RollDirection") Then
            If IsDBNull(drOrderLine("RollDirection")) = False Then
                strFabricRoleDir_value = drOrderLine("RollDirection").ToString
            End If
        End If

        ' -- FittingType --
        If drOrderLine.Table.Columns.Contains("FittingType") Then
            If IsDBNull(drOrderLine("FittingType")) = False Then
                If lst_ProdSys_FittingTypes.ContainsKey(drOrderLine("FittingType").ToString().ToLower()) Then
                    strFittingType_value = lst_ProdSys_FittingTypes(drOrderLine("FittingType").ToString.ToLower)
                End If
            End If
        End If

        ' -- ProductRange --
        If blnIsMotorised Then strProductRange_value = "BlocMotorised" Else strProductRange_value = "Bloc"

        ' -- ProductName --
        If drOrderLine.Table.Columns.Contains("ProductName") Then
            If IsDBNull(drOrderLine("ProductName")) = False Then
                strProductName_value = lst_ProdSys_ProductNames(drOrderLine("ProductName").ToString.ToLower)

                ' Remember to replace the <manufacturer> tag, if it is present
                If drOrderLine.Table.Columns.Contains("Manufacturer") Then
                    If IsDBNull(drOrderLine("Manufacturer")) = False Then
                        strProductName_value = strProductName_value.Replace("<manufacturer>", drOrderLine("Manufacturer").ToString.ToLower)
                    End If
                End If

            End If
        End If

        ' -- Operation --
        If blnIsMotorised = True Then strOperation_value = "MOTORISED" Else strOperation_value = ""

        ' -- RuleSet --
        ' check based on already calculated production system values of ProductName and Fitting Type
        If strProductName_value.ToLower = "blocout40" And strFittingType_value.ToLower = "Edge of Window Recess" Then
            strRuleSet_value = "BLOCOUT40EDGESPRING-AUTO"
        End If


        ' --- Setting parameters ---
        param_OLT2_OrderNo = New SqlParameter("@OrderNo_OLT2" & strSourceLineID, strOrderNo_value)
        param_OLT2_ItemCode = New SqlParameter("@ItemCode_OLT2" & strSourceLineID, strItemCode_value)
        param_OLT2_ItemDescription = New SqlParameter("@ItemDescription_OLT2" & strSourceLineID, strItemDescription_value)
        param_OLT2_ItemPrice = New SqlParameter("@ItemPrice_OLT2" & strSourceLineID, decItemPrice_value)
        param_OLT2_ItemQty = New SqlParameter("@ItemQty_OLT2" & strSourceLineID, intItemQty_value)
        param_OLT2_A = New SqlParameter("@A_OLT2" & strSourceLineID, intA_value)
        param_OLT2_B = New SqlParameter("@B_OLT2" & strSourceLineID, intB_value)
        param_OLT2_C = New SqlParameter("@C_OLT2" & strSourceLineID, intC_value)
        param_OLT2_D = New SqlParameter("@D_OLT2" & strSourceLineID, intD_value)
        param_OLT2_E = New SqlParameter("@E_OLT2" & strSourceLineID, intE_value)
        param_OLT2_F = New SqlParameter("@F_OLT2" & strSourceLineID, intF_value)
        param_OLT2_G = New SqlParameter("@G_OLT2" & strSourceLineID, intG_value)
        param_OLT2_AluColour = New SqlParameter("@AluColour_OLT2" & strSourceLineID, strAluColour_value)
        param_OLT2_FabColour = New SqlParameter("@FabColour_OLT2" & strSourceLineID, strFabColour_value)
        param_OLT2_FabColour2 = New SqlParameter("@FabColour2_OLT2" & strSourceLineID, strFabColour2_value)
        param_OLT2_BracketColour = New SqlParameter("@BracketColour_OLT2" & strSourceLineID, strBracketColour_value)
        param_OLT2_KeepPricingForInvoicing = New SqlParameter("@KeepPricingForInvoicing_OLT2" & strSourceLineID, strKeepPricingForInvoicing_value)
        param_OLT2_CustomerReferenceNo = New SqlParameter("@CustomerReferenceNo_OLT2" & strSourceLineID, strCustomerReferenceNo_value)
        param_OLT2_LineType = New SqlParameter("@LineType_OLT2" & strSourceLineID, strLineType_value)
        param_OLT2_HandleType = New SqlParameter("@HandleType_OLT2" & strSourceLineID, strHandleType_value)
        param_OLT2_ChainType = New SqlParameter("@ChainType_OLT2" & strSourceLineID, strChainType_value)
        param_OLT2_ChainSide = New SqlParameter("@ChainSide_OLT2" & strSourceLineID, strChainSide_value)
        param_OLT2_FabricRoleDir = New SqlParameter("@FabricRoleDir_OLT2" & strSourceLineID, strFabricRoleDir_value)
        param_OLT2_FittingType = New SqlParameter("@FittingType_OLT2" & strSourceLineID, strFittingType_value)
        param_OLT2_InstallationHeight = New SqlParameter("@InstallationHeight_OLT2" & strSourceLineID, strInstallationHeight_value)
        param_OLT2_ProductRange = New SqlParameter("@ProductRange_OLT2" & strSourceLineID, strProductRange_value)
        param_OLT2_ProductName = New SqlParameter("@ProductName_OLT2" & strSourceLineID, strProductName_value)
        param_OLT2_InvoiceUnitPrice = New SqlParameter("@InvoiceUnitPrice_OLT2" & strSourceLineID, decItemPrice_value)
        param_OLT2_Operation = New SqlParameter("@Operation_OLT2" & strSourceLineID, strOperation_value)
        param_OLT2_RuleSet = New SqlParameter("@RuleSet_OLT2" & strSourceLineID, strRuleSet_value)


        ' --- Adding parameters ---
        cmd.Parameters.Add(param_OLT2_OrderNo)
        cmd.Parameters.Add(param_OLT2_ItemCode)
        cmd.Parameters.Add(param_OLT2_ItemDescription)
        cmd.Parameters.Add(param_OLT2_ItemPrice)
        cmd.Parameters.Add(param_OLT2_ItemQty)
        cmd.Parameters.Add(param_OLT2_A)
        cmd.Parameters.Add(param_OLT2_B)
        cmd.Parameters.Add(param_OLT2_C)
        cmd.Parameters.Add(param_OLT2_D)
        cmd.Parameters.Add(param_OLT2_E)
        cmd.Parameters.Add(param_OLT2_F)
        cmd.Parameters.Add(param_OLT2_G)
        cmd.Parameters.Add(param_OLT2_AluColour)
        cmd.Parameters.Add(param_OLT2_FabColour)
        cmd.Parameters.Add(param_OLT2_FabColour2)
        cmd.Parameters.Add(param_OLT2_BracketColour)
        cmd.Parameters.Add(param_OLT2_KeepPricingForInvoicing)
        cmd.Parameters.Add(param_OLT2_CustomerReferenceNo)
        cmd.Parameters.Add(param_OLT2_LineType)
        cmd.Parameters.Add(param_OLT2_HandleType)
        cmd.Parameters.Add(param_OLT2_ChainType)
        cmd.Parameters.Add(param_OLT2_ChainSide)
        cmd.Parameters.Add(param_OLT2_FabricRoleDir)
        cmd.Parameters.Add(param_OLT2_FittingType)
        cmd.Parameters.Add(param_OLT2_InstallationHeight)
        cmd.Parameters.Add(param_OLT2_ProductRange)
        cmd.Parameters.Add(param_OLT2_ProductName)
        cmd.Parameters.Add(param_OLT2_InvoiceUnitPrice)
        cmd.Parameters.Add(param_OLT2_Operation)
        cmd.Parameters.Add(param_OLT2_RuleSet)


        ' --- Adding SQL insert statement to command text ---
        ' Ensure the text "/*ORD_LINE_BLOCKS*/" is persisted, as other functions will try to add more lines into
        ' the command text at this point
        cmd.CommandText = cmd.CommandText.Replace("/*ORD_LINE_BLOCKS*/", strProductLineQuery & vbNewLine & "/*ORD_LINE_BLOCKS*/")

    End Sub


    ''' <summary>
    ''' This function adds the INSERT statements and associated parameters required to import an order's line records to the supplied command 
    ''' </summary>
    ''' <param name="cmd">The full SQL transaction command to add a order lines insert statements / parameters to</param>
    ''' <param name="drOrderHeader">The ORD_HEADER record from the website to create ORD_LINES insert statements from</param>
    ''' <param name="strWebsite">The location code of the source website. E.g. UK</param>
    ''' <remarks>
    ''' Author: Huw Day - 06/02/2018
    ''' The command is passed in by reference. Any changes will be available to the calling function without
    ''' the command needing to be passed back.
    ''' </remarks>
    Function Build_OrderLines(ByRef cmd As SqlCommand, drOrderHeader As DataRow, strWebsite As String) As OrderLeadTime

        ' - Use order header to load table of order lines
        ' - Loop through each order line:
        '   - Write a standard product line into ProdSys
        '   - If the product is a blocout40 and the fitting type contains "Recess" or "Surface", create a FRAME line
        '   - If the product is a twin roller blind, create a secondary twin roller line
        ' - If any product has Measure Protect, create a single Measure Protect line for the order
        ' - Create a final delivery line for the whole order


        ' --- Defining objects and variables ---
        Dim strGetOrderLinesQuery As String
        Dim cmdGetOrderLines As SqlCommand
        Dim paramOrderNo As SqlParameter
        Dim dtOrderLines As New DataTable
        Dim da As SqlDataAdapter
        Dim blnHasMeasureProtect As Boolean = False
        Dim decMeasureProtectValue As Decimal = 0
        Dim orderLeadTime As New OrderLeadTime(0, New TimeSpan(23, 59, 59))


        ' --- Getting Order's lines ---
        'strGetOrderLinesQuery = "SELECT ORD_LINES.*, Ref_ProductNames.ExportAsName " _
        '                      & "FROM ORD_LINES INNER JOIN Ref_ProductNames On ORD_LINES.ProductName = Ref_ProductNames.AppName " _
        '                      & "WHERE Order_No = @OrderNo;"

        strGetOrderLinesQuery = "SELECT ORD_LINES.*, Ref_ProductNames.ExportAsName, " _
                              & "(SELECT UKLeadMaxDays FROM Products_LeadTimes As PLT2 WHERE PLT2.ID = IsNull( " _
                              & "(SELECT ID FROM Products_LeadTimes As PLT1 WHERE PLT1.ProductName = ORD_LINES.ProductName AND PLT1.FabricsList LIKE '%#' + ORD_LINES.FabricCode + ';%'), " _
                              & "(SELECT ID FROM Products_LeadTimes As PLT1 WHERE PLT1.ProductName = ORD_LINES.ProductName AND PLT1.FabricsList = '#default;') " _
                              & ")) As LeadTimeDays, " _
                              & "(SELECT Cutoff_Time FROM Products_LeadTimes As PLT4 WHERE PLT4.ID = IsNull( " _
                              & "(SELECT ID FROM Products_LeadTimes As PLT3 WHERE PLT3.ProductName = ORD_LINES.ProductName AND PLT3.FabricsList LIKE '%#' + ORD_LINES.FabricCode + ';%'), " _
                              & "(SELECT ID FROM Products_LeadTimes As PLT3 WHERE PLT3.ProductName = ORD_LINES.ProductName AND PLT3.FabricsList = '#default;') " _
                              & ")) As Cutoff_Time " _
                              & "FROM ORD_LINES INNER JOIN Ref_ProductNames On ORD_LINES.ProductName = Ref_ProductNames.AppName " _
                              & "WHERE Order_No = @OrderNo;"


        Using conn As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings(strWebsite).ConnectionString)
            cmdGetOrderLines = New SqlCommand(strGetOrderLinesQuery, conn)
            paramOrderNo = New SqlParameter("@OrderNo", drOrderHeader("Order_No"))
            cmdGetOrderLines.Parameters.Add(paramOrderNo)
            da = New SqlDataAdapter
            da.SelectCommand = cmdGetOrderLines
            Try
                da.Fill(dtOrderLines)
            Catch ex As Exception
                Throw ex
            End Try
        End Using


        ' --- Looping through each line in the ORD_LINES table ---
        For Each drOrderLine As DataRow In dtOrderLines.Rows


            ' --- Processing product lead time data ---
            If IsDBNull(drOrderLine("LeadTimeDays")) = False Then
                Dim intTempLeadtimeDays As Integer
                If Integer.TryParse(drOrderLine("LeadTimeDays"), intTempLeadtimeDays) Then
                    Dim tsTempCutoffTime As TimeSpan
                    If IsDBNull(drOrderLine("Cutoff_Time")) = False Then
                        If TimeSpan.TryParse(drOrderLine("Cutoff_Time").ToString(), tsTempCutoffTime) = False Then
                            tsTempCutoffTime = New TimeSpan(23, 59, 59)
                        End If
                    Else
                        tsTempCutoffTime = New TimeSpan(23, 59, 59)
                    End If
                    If orderLeadTime.LeadTimeDays < intTempLeadtimeDays Then
                        ' Take the largest cutoff time
                        orderLeadTime.LeadTimeDays = intTempLeadtimeDays
                        orderLeadTime.CutoffTime = tsTempCutoffTime

                    ElseIf orderLeadTime.LeadTimeDays = intTempLeadtimeDays And orderLeadTime.CutoffTime > tsTempCutoffTime Then
                        ' If the lead times are the same, compare the cutoff times and keep the sooner cutoff time
                        orderLeadTime.CutoffTime = tsTempCutoffTime
                    End If
                End If
            End If


            ' --- Checking if product is or isn't a spare part ---
            If lst_SpareParts.Contains(drOrderLine("ProductName").ToString()) = False Then
                class_blnIsSparesOrder = False
            End If


            ' --- Excluding Fabric Samples from being imported ---
            If drOrderLine("ProductName").ToString().ToLower() = "fabricsample" Then
                ' Do nothing. This line should not be added to the command

            ElseIf drOrderLine("ProductName").ToString().ToLower() = "measureservice" Then
                ' Do nothing. This line should not be added to the command

            Else
                ' --- Writing product's order line into the command ---
                Build_OrderLine_Product(cmd, drOrderLine, strWebsite)

                ' --- Write extra line for Twin Rollers ---
                If drOrderLine("ProductName").ToString.ToLower = "twinrollerblind" Then
                    Build_OrderLine_TwinSecondRow(cmd, drOrderLine, strWebsite)
                End If

                ' --- If the product is a blocout40, write a frame line into the command ---
                If drOrderLine("ProductName").ToString.ToLower = "blocout40" Then
                    Select Case drOrderLine("FittingType").ToString.ToLower
                        Case "recessinside", "surface3", "surface4"                 ' Only create a frame for these three fitting types
                            Build_OrderLine_Frame(cmd, drOrderLine, strWebsite)
                    End Select
                End If

                ' --- If the product has measure protect, record it ---
                If CBool(drOrderLine("MeasureProtect")) Then
                    blnHasMeasureProtect = True
                    decMeasureProtectValue = CDec(drOrderLine("MeasureProtectValue"))       ' Keep a running total of the measure protect cost for all products
                End If

            End If
        Next


        ' --- If any product in the order has measure protect, write a measure protect line ---
        If blnHasMeasureProtect Then
            Build_OrderLine_MeasureProtect(cmd, drOrderHeader, strWebsite, decMeasureProtectValue)
        End If


        ' --- Writing delivery line ---
        Build_OrderLine_Delivery(cmd, drOrderHeader, strWebsite)


        ' --- Removing the /*ORD_LINE_BLOCKS*/ placeholder from the command text ---
        cmd.CommandText = cmd.CommandText.Replace("/*ORD_LINE_BLOCKS*/", "")


        ' --- Returning order lead time info ---
        Return orderLeadTime

    End Function


    ''' <summary>
    ''' This function adds the INSERT statements and associated parameters required to import an order's line records to the supplied command 
    ''' </summary>
    ''' <param name="cmd">The full SQL transaction command to add a order lines insert statements / parameters to</param>
    ''' <param name="drOrderHeader">The ORD_HEADER record from the website to create ORD_LINES insert statements from</param>
    ''' <param name="strWebsite">The location code of the source website. E.g. UK</param>
    ''' <remarks>
    ''' Author: Huw Day - /04/08/2020
    ''' The command is passed in by reference. Any changes will be available to the calling function without
    ''' the command needing to be passed back.
    ''' </remarks>
    Sub Build_OrderLines_BFS(ByRef cmd As SqlCommand, drOrderHeader As DataRow, strWebsite As String)

        ' - Use order header to load table of order lines
        ' - Loop through each order line and write a standard product line into ProdSys
        ' - Create a final delivery line for the whole order


        ' --- Defining objects and variables ---
        Dim strQuery_GetOrderLineData As String
        Dim cmdGetOrderLines As SqlCommand
        Dim paramOrderNo As SqlParameter
        Dim dtOrderLines As New DataTable
        Dim daOrderLines As SqlDataAdapter


        ' --- Writing out queries ---
        strQuery_GetOrderLineData = "SELECT " _
                                 & "IsNull((SELECT ParamValue FROM Product_AdditionalParameters WHERE Product_AdditionalParameters.ProductID = ORD_LINES.ProductID AND Product_AdditionalParameters.ParamName = 'ProdSysCode'), '') As ProdSysCode, " _
                                 & "IsNull((SELECT ParamValue FROM Product_AdditionalParameters WHERE Product_AdditionalParameters.ProductID = ORD_LINES.ProductID AND Product_AdditionalParameters.ParamName = 'ProdSysName'), '') As ProdSysName, " _
                                 & "IsNull((SELECT ParamValue FROM Product_AdditionalParameters WHERE Product_AdditionalParameters.ProductID = ORD_LINES.ProductID AND Product_AdditionalParameters.ParamName = 'ProdSys_LineType'), '') As ProdSys_LineType, " _
                                 & "* " _
                                 & "FROM ORD_LINES	" _
                                 & "INNER JOIN Product ON ORD_LINES.ProductID = Product.ProductID " _
                                 & "WHERE Order_No = @OrderNo; "


        ' --- Retrieving order line data ---
        Using conn As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings(strWebsite).ConnectionString)
            cmdGetOrderLines = New SqlCommand(strQuery_GetOrderLineData, conn)
            paramOrderNo = New SqlParameter("@OrderNo", drOrderHeader("Order_No"))
            cmdGetOrderLines.Parameters.Add(paramOrderNo)
            daOrderLines = New SqlDataAdapter
            daOrderLines.SelectCommand = cmdGetOrderLines
            Try
                daOrderLines.Fill(dtOrderLines)
            Catch ex As Exception
                Throw ex
            Finally
                daOrderLines.Dispose()
            End Try
        End Using


        ' --- Looping through each line in the ORD_LINES table ---
        For Each drOrderLine As DataRow In dtOrderLines.Rows
            Build_OrderLine_BFS_Product(cmd, drOrderLine, strWebsite)
        Next


        ' --- Writing delivery line ---
        Build_OrderLine_BFS_Delivery(cmd, drOrderHeader, strWebsite)


        ' --- Removing the /*ORD_LINE_BLOCKS*/ placeholder from the command text ---
        cmd.CommandText = cmd.CommandText.Replace("/*ORD_LINE_BLOCKS*/", "")


    End Sub


    ''' <summary>
    ''' This function takes a price value from the specified source website (e.g. Euros from the IE site)
    ''' and converts it to Pounds Sterling
    ''' </summary>
    ''' <param name="decPriceToConvert"></param>
    ''' <param name="strWebsite"></param>
    ''' <returns></returns>
    Function Currency_ConvertToPounds(decPriceToConvert As Decimal, strWebsite As String) As Decimal

        Return decPriceToConvert / Currency_GetConversionRatio(strWebsite)

    End Function


    ''' <summary>
    ''' This function returns the ratio required to convert a pounds sterling price to that of
    ''' the named website (e.g. returns 1.2 to convert pounds to euros)
    ''' </summary>
    ''' <param name="strWebsite"></param>
    ''' <returns></returns>
    Function Currency_GetConversionRatio(strWebsite As String) As Decimal

        Dim decOutput As Decimal = 1
        Select Case strWebsite.ToUpper()
            Case enumWebsites.IE.ToString().ToUpper()
                decOutput = CURRENCY_CONVERT_POUNDS_TO_EUROS
            Case enumWebsites.TEST.ToString().ToUpper()
                decOutput = 1
            Case enumWebsites.IE.ToString().ToUpper()
                decOutput = 1
            Case Else
                decOutput = 1
        End Select
        Return decOutput

    End Function


    ''' <summary>
    ''' This function gets the Fabric Sample Discount setting, as a percentage, from the appropriate website
    ''' </summary>
    ''' <param name="strWebsite">The location code of the source website. E.g. UK</param>
    ''' <remarks>
    ''' Author: Huw Day - 16/02/2018
    ''' </remarks>
    Function Get_FabricSamples_Discount(strWebsite As String) As Integer

        ' --- Defining objects and variables ---
        Dim intResult As Integer = 0   ' Default value
        Dim strResult As String = ""
        Dim cmd As SqlCommand
        Dim query As String = "Select SettingValue FROM Settings WHERE SettingName = 'FabricSampleDiscount';"

        ' --- Executing query ---
        Using conn As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings(strWebsite).ConnectionString)
            cmd = New SqlCommand(query, conn)
            Try
                conn.Open()
                strResult = cmd.ExecuteScalar
                If IsNumeric(strResult) Then
                    intResult = CInt(strResult)
                End If
            Catch ex As Exception
                ' Do nothing. The only error likely to occur would be a missing database connection or setting
            End Try
        End Using

        ' --- Returning result ---
        Return intResult

    End Function

    ''' <summary>
    ''' Returns the fabric sample discount code
    ''' </summary>
    ''' <param name="intOrderNo">The number of the order that contains the fabric samples</param>
    ''' <param name="strWebsite">The location code of the source website. E.g. UK</param>
    ''' <remarks>
    ''' Author: Huw Day - 16/02/2018
    ''' </remarks>
    Function Get_FabricSamples_DiscountCode(intOrderNo As Integer, strWebsite As String) As String
        ' strWebsite is currently unused. It has been added in anticipation of the coupon code being
        ' different between different instances of the website

        Return "SAMP" & intOrderNo.ToString

    End Function

    ''' <summary>
    ''' Returns the expiry date of a fabric sample discount code issued today
    ''' </summary>
    ''' <remarks>
    ''' Author: Huw Day - 16/02/2018
    ''' </remarks>
    Function Get_FabricSamples_DiscountExpiry() As Date

        Dim dateOutput As Date = DateAdd(DateInterval.Day, 7, Now)
        dateOutput = New Date(dateOutput.Year, dateOutput.Month, dateOutput.Day, 23, 59, 59)
        Return dateOutput

    End Function

    ''' <summary>
    ''' This function retrieves the number of orders that have fabric samples that have not yet been imported to the Production System
    ''' </summary>
    ''' <param name="strWebsite">The location code of the source website. E.g. UK</param>
    ''' <remarks>
    ''' Author: Huw Day - 16/02/2018
    ''' </remarks>
    Function Get_FabricSamplesToImport_Count(strWebsite As String) As Integer

        Select Case strWebsite
            Case "UK", "IE", "TEST"

                ' --- Defining objects and variables ---
                Dim cmd As SqlCommand
                Dim paramLastOrderImported As SqlParameter
                Dim strResult As String
                Dim intCount As Integer = 0
                Dim intLastOrderImported = Importer.Get_ProdSys_LastOrderImported_FabricSamples(strWebsite)
                Dim query As String = "SELECT Count(Distinct ORD_HEADER.Order_No) FROM ORD_HEADER INNER JOIN ORD_LINES ON ORD_HEADER.Order_No = ORD_LINES.Order_No " _
                    & "WHERE ORD_HEADER.OrderStatus = 'PAY_AUTH' And ProductName = 'FabricSample' And ORD_HEADER.Order_No > @LastOrderImported;"

                ' --- Executing query ---
                Using conn As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings(strWebsite).ConnectionString)
                    cmd = New SqlCommand(query, conn)
                    paramLastOrderImported = New SqlParameter("@LastOrderImported", intLastOrderImported)
                    cmd.Parameters.Add(paramLastOrderImported)

                    Try
                        conn.Open()
                        strResult = cmd.ExecuteScalar
                        intCount = CInt(strResult)
                    Catch ex As Exception
                        Throw ex
                    End Try
                End Using

                Return intCount

            Case Else
                Return 0

        End Select

    End Function

    ''' <summary>
    ''' This function retrieves a table of order headers that have fabric samples that have not yet been imported to the Production System
    ''' </summary>
    ''' <param name="strWebsite">The location code of the source website. E.g. UK</param>
    ''' <remarks>
    ''' Author: Huw Day - 16/02/2018
    ''' </remarks>
    Function Get_FabricSamplesToImport_OrderHeaders(strWebsite As String) As DataTable

        ' --- Defining objects and variables ---
        Dim cmd As SqlCommand
        Dim dtFabricSampleHeaders As New DataTable
        Dim da As SqlDataAdapter
        Dim paramLastOrderImported As SqlParameter
        Dim intLastOrderImported = Importer.Get_ProdSys_LastOrderImported_FabricSamples(strWebsite)
        Dim query As String = "SELECT Distinct ORD_HEADER.Order_No, ORD_HEADER.DateEntered, ORD_HEADER.Delivery_FirstName, " _
            & "ORD_HEADER.Delivery_LastName, ORD_HEADER.Delivery_HouseNumber, ORD_HEADER.Delivery_Address, ORD_HEADER.Delivery_City, " _
            & "ORD_HEADER.Delivery_Country, ORD_HEADER.Delivery_Postcode, ORD_HEADER.Billing_Email, ORD_HEADER.Billing_PhoneNo " _
            & "FROM ORD_HEADER INNER JOIN ORD_LINES On ORD_HEADER.Order_No = ORD_LINES.Order_No " _
            & "Where ORD_HEADER.OrderStatus = 'PAY_AUTH' And ProductName = 'FabricSample' And ORD_HEADER.Order_No > @LastOrderImported;"

        ' --- Connecting to database and executing query ---
        Using conn As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings(strWebsite).ConnectionString)
            cmd = New SqlCommand(query, conn)
            da = New SqlDataAdapter
            paramLastOrderImported = New SqlParameter("@LastOrderImported", intLastOrderImported)
            cmd.Parameters.Add(paramLastOrderImported)
            da.SelectCommand = cmd

            Try
                da.Fill(dtFabricSampleHeaders)
            Catch ex As Exception
                Throw ex
            Finally
                da.Dispose()
            End Try
        End Using

        ' --- Returning data table ---
        Return dtFabricSampleHeaders

    End Function

    ''' <summary>
    ''' This function retrieves a data table containing fabric samples for the specified order
    ''' </summary>
    ''' <param name="strWebsite">The location code of the source website. E.g. UK</param>
    ''' <param name="intOrderNo">The order number of the target order</param>
    ''' <remarks>
    ''' Author: Huw Day - 16/02/2018
    ''' </remarks>
    Function Get_FabricSamplesToImport_OrderLines(strWebsite As String, intOrderNo As Integer)

        ' --- Defining objects and variables ---
        Dim cmd As SqlCommand
        Dim dtFabricSamplesToImport As New DataTable
        Dim da As SqlDataAdapter
        Dim paramOrderNo As SqlParameter
        Dim query As String = "SELECT ORD_LINES.ProductName, ORD_LINES.FabricCode, Ref_FabricCodes.FabricName " _
            & "FROM ORD_LINES INNER JOIN Ref_FabricCodes ON ORD_LINES.FabricCode = Ref_FabricCodes.FabricCode " _
            & "WHERE ORD_LINES.Order_No = @OrderNo And ProductName = 'FabricSample';"

        ' --- Connecting to database and executing query ---
        Using conn As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings(strWebsite).ConnectionString)

            cmd = New SqlCommand(query, conn)
            da = New SqlDataAdapter
            paramOrderNo = New SqlParameter("@OrderNo", intOrderNo)
            cmd.Parameters.Add(paramOrderNo)
            da.SelectCommand = cmd

            Try
                da.Fill(dtFabricSamplesToImport)
            Catch ex As Exception
                Throw ex
            Finally
                da.Dispose()
            End Try
        End Using

        ' --- Returning data table ---
        Return dtFabricSamplesToImport

    End Function


    ''' <summary>
    ''' A function to create the item code to be entered into the Production System's ORD_LINES table for any ordered product
    ''' </summary>
    ''' <param name="drOrderLine">Data row from the website's ORD_LINE table</param>
    ''' <remarks>Author: Huw Day - 05/02/2018</remarks>
    ''' <returns>The item code to add to the production system's ORD_LINES table</returns>
    Function Get_ItemCode(drOrderLine As DataRow, strWebsite As String) As String


        ' --- Defining variables ---
        Dim strResult As String = ""
        Dim strAppName As String = ""
        Dim strItemCodePrefix As String = ""
        Dim strManufacturerCode As String = ""
        Dim strMotor As String = ""
        Dim strWindowCode As String = ""
        Dim strWidthHeight As String = ""
        Dim strAluminiumCode As String = ""
        Dim strFabricCode As String = ""


        ' --- Grabbing website's product name ---
        If drOrderLine.Table.Columns.Contains("ProductName") Then
            strAppName = drOrderLine("ProductName").ToString
        End If


        ' --- Establishing item code format based on product name ---
        Select Case strAppName.ToLower
            Case "additionalfabric"
                Dim strPrefix = ""
                If drOrderLine.Table.Columns.Contains("FittingType") Then
                    strPrefix = drOrderLine("FittingType")
                End If
                If strPrefix = "" Then strPrefix = lst_ProdSys_ItemCodePrefix("additionalfabric")      ' default, as set in the Initialise_ClassVariables function
                strResult = strPrefix & "-<widthxheight>|<aluminiumcode>|<fabriccode>"

            Case "blocout40", "blocout80", "customskylight", "pelmet", "premiumrollerblind",
                 "rollerblind", "rollerblinddrillfree", "rollerblindmotorised", "twinrollerblind", "virtualfabric", "zebrarollerblind"
                strResult = "<prefix><motor>-<widthxheight>|<aluminiumcode>|<fabriccode>"

            Case "premiumskylight", "solarskylight"
                strResult = "<prefix><manufacturercode><motor>-<windowcode>|<aluminiumcode>|<fabriccode>"

            Case "pole2m"
                strResult = "<prefix>|NA|2m Pole"

            Case "pole3m"
                strResult = "<prefix>|NA|3m Pole"

            Case "multichannelremote", "singlechannelremote"
                strResult = "<prefix>|NA|Remote"

            Case "smartmotorcharger",
                "smartremote",
                "smarthub",
                "smartrepeater",
                "faceshield_1b12",
                "faceshield_5b12",
                "faceshield_20b12",
                "faceshieldkit_1b1",
                "faceshieldkit_1b2",
                "faceshieldkit_1b3",
                "faceshieldkit_1b4",
                "faceshieldkit_1b5",
                "faceshieldkit_1b12",
                "faceshieldkit_1b24",
                "faceshieldkit_1b40",
                "faceshieldkit_4b24",
                "bo40_brakes_cream",
                "bo40_brakes_grey",
                "bo40_brakes_white",
                "bo40_fittings",
                "bo80_fittings_recess",
                "bo80_fittings_surface",
                "prb_fittings_recess",
                "prb_fittings_surface",
                "rb_brackets_black",
                "rb_brackets_grey",
                "rb_brackets_white",
                "rb_df_brackets_white",
                "rb_m_brackets_black",
                "rb_m_brackets_white",
                "sky_brakes_cream",
                "sky_brakes_grey",
                "sky_brakes_white",
                "ss_battery_post_03_19",
                "ss_battery_pre_03_19",
                "zrb_fittings_recess",
                "zrb_fittings_surface"
                strResult = drOrderLine("Product_ItemCode")

            Case "rollerblindextend_black"
                strResult = drOrderLine("Product_ItemCode") & "-869x1980|NA|kit"

            Case "rollerblindextend_blue"
                strResult = drOrderLine("Product_ItemCode") & "-1169x1980|NA|kit"

            Case "rollerblindextend_red"
                strResult = drOrderLine("Product_ItemCode") & "-1469x1980|NA|kit"

            Case "rollerblindextend_green"
                strResult = drOrderLine("Product_ItemCode") & "-1769x1980|NA|kit"

            Case "fabricsample"
                ' Note, fabric samples should be filtered out before they get here
                strResult = "<prefix>|NA|<fabriccode>"

            Case "measureprotect"
                strResult = "<prefix>|NA|NA"

            Case "venetian"
                strResult = "<prefix>-<widthxheight>|na|<fabriccode>"

            Case Else
                strResult = "<prefix>|NA|NA"

        End Select


        ' --- Looking up item code prefix ---
        If lst_ProdSys_ItemCodePrefix.ContainsKey(strAppName.ToLower) Then
            strItemCodePrefix = lst_ProdSys_ItemCodePrefix(strAppName.ToLower)
        End If
        strResult = strResult.Replace("<prefix>", strItemCodePrefix)


        ' --- Looking up manufacturer code ---
        ' Note: This will not be applied to the item code if the preset format does not possess a <manufacturercode> tag
        If drOrderLine.Table.Columns.Contains("Manufacturer") Then
            If IsDBNull(drOrderLine("Manufacturer")) = False Then
                strManufacturerCode = lst_ProdSys_ManufacturerCodes(drOrderLine("Manufacturer").ToString.ToLower)
            End If
        End If
        strResult = strResult.Replace("<manufacturercode>", strManufacturerCode)


        ' --- Applying motorised tag ---
        ' Note: This will not be applied to the item code if the preset format does not possess a <motor> tag
        If drOrderLine.Table.Columns.Contains("Motorised") Then
            If drOrderLine("Motorised") = 1 Then
                strMotor = "-M"
            End If
        End If
        strResult = strResult.Replace("<motor>", strMotor)


        ' --- Applying window code ---
        ' Note: This will not be applied to the item code if the preset format does not possess a <windowcode> tag
        If drOrderLine.Table.Columns.Contains("WindowCode") Then
            If IsDBNull(drOrderLine("WindowCode")) = False Then
                strWindowCode = drOrderLine("WindowCode").ToString
            End If
        End If
        strResult = strResult.Replace("<windowcode>", strWindowCode)


        ' --- Applying width and height ---
        ' Note: This will not be applied to the item code if the preset format does not possess a <widthxheight> tag
        If drOrderLine.Table.Columns.Contains("Width") And drOrderLine.Table.Columns.Contains("Height") Then
            If IsDBNull(drOrderLine("Width")) = False And IsDBNull(drOrderLine("Height")) = False Then
                strWidthHeight = drOrderLine("Width").ToString() & "x" & drOrderLine("Height").ToString()
            End If
        End If
        strResult = strResult.Replace("<widthxheight>", strWidthHeight)


        ' --- Applying aluminium code ---
        ' Note: This will not be applied to the item code if the preset format does not possess an <aluminiumcode> tag
        If drOrderLine.Table.Columns.Contains("AluminiumColour") Then
            If IsDBNull(drOrderLine("AluminiumColour")) = False Then
                strAluminiumCode = drOrderLine("AluminiumColour").ToString
            End If
        End If
        strResult = strResult.Replace("<aluminiumcode>", strAluminiumCode)


        ' --- Applying fabric code ---
        ' Note: This will not be applied to the item code if the preset format does not possess a <fabriccode> tag
        If drOrderLine.Table.Columns.Contains("FabricCode") Then
            If IsDBNull(drOrderLine("FabricCode")) = False Then
                strFabricCode = drOrderLine("FabricCode").ToString
                strFabricCode = Get_OriginalFabricCode(strFabricCode, strWebsite)
            End If
        End If
        strResult = strResult.Replace("<fabriccode>", strFabricCode)


        ' --- Returning the calculated item code ---
        Return strResult.ToUpper()

    End Function

    ''' <summary>
    ''' A function to create the item description to be entered into the Production System's ORD_LINES table for any ordered product
    ''' </summary>
    ''' <param name="drOrderLine">Data row from the website's ORD_LINE table</param>
    ''' <remarks>Author: Huw Day - 06/02/2018</remarks>
    ''' <returns>The item description to add to the production system's ORD_LINES table</returns>
    Function Get_ItemDescription(drOrderLine As DataRow) As String

        ' This function mirrors the old order importer, in that it cycles through each
        ' column from the supplied table and, if the content is not null, it writes
        ' the column name and it's contents into the description.
        ' The whole description is then preceded by a string calculated from the
        ' product range and name

        Dim lstColumnWhitelist As New List(Of String)
        Dim strResult As String = ""
        Dim strAppName As String = ""
        Dim strDescriptionLead As String = ""
        Dim strProductRange As String = ""
        Dim strProductName As String = ""
        Dim strManufacturer As String = ""
        Dim strBatterySolar As String = ""


        ' --- Compiling list of columns that should appear in the item description (if they have content) ---
        lstColumnWhitelist.Add("FabricCode".ToLower)
        lstColumnWhitelist.Add("Quantity".ToLower)
        lstColumnWhitelist.Add("UnitOfMeasure".ToLower)
        lstColumnWhitelist.Add("AluminiumColour".ToLower)
        lstColumnWhitelist.Add("BracketColour".ToLower)
        lstColumnWhitelist.Add("ChainSide".ToLower)
        lstColumnWhitelist.Add("CustomerReference".ToLower)
        lstColumnWhitelist.Add("InnerFabricCode".ToLower)
        lstColumnWhitelist.Add("FittingType".ToLower)
        lstColumnWhitelist.Add("FrameDepth".ToLower)
        lstColumnWhitelist.Add("FrameDrop".ToLower)
        lstColumnWhitelist.Add("FrameWidth".ToLower)
        lstColumnWhitelist.Add("GlassDrop".ToLower)
        lstColumnWhitelist.Add("GlassWidth".ToLower)
        lstColumnWhitelist.Add("Height".ToLower)
        lstColumnWhitelist.Add("HeightLeft".ToLower)
        lstColumnWhitelist.Add("HeightRight".ToLower)
        lstColumnWhitelist.Add("Manufacturer".ToLower)
        lstColumnWhitelist.Add("MeasureProtect".ToLower)
        lstColumnWhitelist.Add("Motorised".ToLower)
        lstColumnWhitelist.Add("RecessDepth".ToLower)
        lstColumnWhitelist.Add("RollDirection".ToLower)
        lstColumnWhitelist.Add("Width".ToLower)
        lstColumnWhitelist.Add("WidthBottom".ToLower)
        lstColumnWhitelist.Add("WidthMiddle".ToLower)
        lstColumnWhitelist.Add("WindowCode".ToLower)


        ' --- Grabbing website's product name ---
        If drOrderLine.Table.Columns.Contains("ProductName") Then
            strAppName = drOrderLine("ProductName").ToString
        End If


        ' --- Component items have different item_description rules. Detecting item type ---
        Select Case strAppName.ToLower
            Case "smartmotorcharger",
                "smartremote",
                "smarthub",
                "smartrepeater",
                "rollerblindextend_black",
                "rollerblindextend_blue",
                "rollerblindextend_red",
                "rollerblindextend_green",
                "faceshield_1b12",
                "faceshield_5b12",
                "faceshield_20b12",
                "faceshieldkit_1b1",
                "faceshieldkit_1b2",
                "faceshieldkit_1b3",
                "faceshieldkit_1b4",
                "faceshieldkit_1b5",
                "faceshieldkit_1b12",
                "faceshieldkit_1b24",
                "faceshieldkit_1b40",
                "faceshieldkit_4b24",
                "bo40_brakes_cream",
                "bo40_brakes_grey",
                "bo40_brakes_white",
                "bo40_fittings",
                "bo80_fittings_recess",
                "bo80_fittings_surface",
                "prb_fittings_recess",
                "prb_fittings_surface",
                "rb_brackets_black",
                "rb_brackets_grey",
                "rb_brackets_white",
                "rb_df_brackets_white",
                "rb_m_brackets_black",
                "rb_m_brackets_white",
                "sky_brakes_cream",
                "sky_brakes_grey",
                "sky_brakes_white",
                "ss_battery_post_03_19",
                "ss_battery_pre_03_19",
                "zrb_fittings_recess",
                "zrb_fittings_surface"

                ' ===== Using Component Item description rules =====
                strResult = drOrderLine("ExportAsName").ToString
                ' ===== End of Component Item description rules =====

            Case Else

                ' ===== Using standard item description rules =====
                ' --- Calculating item description lead ---
                Select Case strAppName.ToLower
                    Case "premiumskylight", "solarskylight"
                        strDescriptionLead = "<range>-<productname>-<manufacturer><BatterySolar>"
                    Case Else
                        strDescriptionLead = "<range>-<productname>"
                End Select


                ' --- Selecting ProductRange / BatterySolar tags based on Motor ---
                strProductRange = "Bloc"
                If drOrderLine.Table.Columns.Contains("Motorised") Then
                    If CBool(drOrderLine("Motorised")) = True Then
                        strProductRange = "BlocMotorised"
                        strBatterySolar = "-BatterySolar"
                    End If
                End If
                strDescriptionLead = strDescriptionLead.Replace("<range>", strProductRange)             ' Either "Bloc" or "BlocMotorised"
                strDescriptionLead = strDescriptionLead.Replace("<BatterySolar>", strBatterySolar)      ' Either "-BatterySolar" or ""


                ' --- Looking up ProdSys Product Name ---
                If lst_ProdSys_ProductNames.ContainsKey(strAppName.ToLower) Then
                    strProductName = lst_ProdSys_ProductNames(strAppName.ToLower)
                End If
                strDescriptionLead = strDescriptionLead.Replace("<productname>", strProductName)


                ' --- Applying manufacturer to Description Lead ---
                If drOrderLine.Table.Columns.Contains("Manufacturer") Then
                    If IsDBNull(drOrderLine("Manufacturer")) = False Then
                        strManufacturer = drOrderLine("Manufacturer").ToString
                    End If
                End If
                strDescriptionLead = strDescriptionLead.Replace("<manufacturer>", strManufacturer)

                ' ## The description lead should now be completed ##


                ' --- Adding description lead to full description string ---
                strResult = strDescriptionLead & ": "


                ' --- Adding each column name and it's content to the full description string ---
                For Each column As DataColumn In drOrderLine.Table.Columns
                    If lstColumnWhitelist.Contains(column.ColumnName.ToLower) Then

                        If strAppName.ToLower() = "additionalfabric" And column.ColumnName.ToLower() = "fittingtype" Then
                            ' Do nothing. Fitting type is being borrowed here to track the type of blind that the additional fabric is for
                        Else
                            strResult = strResult & column.ColumnName & ": " & drOrderLine(column).ToString & " "
                        End If

                    Else
                        ' Do nothing. The column is not in the whitelist and should not go in to the description
                    End If
                Next

                ' ===== End of standard item description rules =====

        End Select


        ' --- Return result ---
        Return strResult


    End Function

    ''' <summary>
    ''' A data layer function to retrieve the original fabric code that the input fabric was based on
    ''' </summary>
    ''' <param name="strFabricCode">The current, possibly duplicated fabric code</param>
    ''' <param name="strWebsite">Website to retrieve the setting from</param>
    ''' <remarks>Author: Huw Day - 05/02/2018</remarks>
    ''' <returns>The original fabric code</returns>
    Function Get_OriginalFabricCode(strFabricCode As String, strWebsite As String) As String

        Dim query As String = "SELECT OriginalCode FROM Ref_FabricCodes WHERE FabricCode = @FabricCode;"
        Dim strResult As String = strFabricCode         ' By default, the output defaults to the input. Only translate if a valid result is returned
        Dim cmd As SqlCommand
        Dim paramFabricCode As SqlParameter

        Using conn As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings(strWebsite).ConnectionString)

            cmd = New SqlCommand(query, conn)
            paramFabricCode = New SqlParameter("@FabricCode", strFabricCode)
            cmd.Parameters.Add(paramFabricCode)

            Try
                conn.Open()
                strResult = cmd.ExecuteScalar

            Catch ex As Exception
                ' Do nothing for now

            End Try

        End Using

        Return strResult

    End Function


    Function Get_ProdSys_LastOrderImported_FabricSamples(strWebsite As String) As Integer

        ' --- Defining objects and variables ---
        Dim intLastOrderImported As Integer
        Dim strResult As String
        Dim query As String
        Dim SiteName As String = ""
        Dim paramSiteName As SqlParameter
        Dim cmd As SqlCommand

        ' --- pre-calculating ProdSys site name ---
        If strWebsite = "UK" Then
            SiteName = "BLOCUKWebsite"
        ElseIf strWebsite = "IE" Then
            SiteName = "BlocBlindsIE"
        Else
            SiteName = "BlocBlinds" & strWebsite
        End If

        ' --- running query ---
        query = "Select isnull(MAX(cast(OriginatingOrderNumber as integer)),0) from FabricSampleOrders where OrderOrigination= @SiteName;"
        Using connProdSys As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("PRODSYS").ConnectionString)

            paramSiteName = New SqlParameter("@SiteName", SiteName)
            cmd = New SqlCommand(query, connProdSys)
            cmd.Parameters.Add(paramSiteName)

            Try
                connProdSys.Open()
                strResult = cmd.ExecuteScalar()
                intLastOrderImported = CInt(strResult)
            Catch ex As Exception
                Throw ex
            End Try
        End Using

        ' --- Applying limits to result ---
        If intLastOrderImported < 2356 And strWebsite = "UK" Then
            intLastOrderImported = 2356
        ElseIf intLastOrderImported < 1109 And strWebsite = "IE" Then
            intLastOrderImported = 1109
        End If

        Return intLastOrderImported

    End Function


    ''' <summary>
    ''' A data layer function to retrieve a named setting from the specified website's database
    ''' </summary>
    ''' <param name="strSettingName">Name of the setting to retrieve</param>
    ''' <param name="strWebsite">Website to retrieve the setting from</param>
    ''' <remarks>Author: Huw Day - 05/02/2018</remarks>
    ''' <returns>String</returns>
    Function Get_SettingFromWebsite(strSettingName As String, strWebsite As String) As String

        Dim query As String = "SELECT SettingValue FROM Settings WHERE SettingName = @SettingName;"
        Dim strResult As String = ""
        Dim cmd As SqlCommand
        Dim paramSettingName As SqlParameter

        Using conn As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings(strWebsite).ConnectionString)

            cmd = New SqlCommand(query, conn)
            paramSettingName = New SqlParameter("@SettingName", strSettingName)
            cmd.Parameters.Add(paramSettingName)

            Try
                conn.Open()
                strResult = cmd.ExecuteScalar

            Catch ex As Exception
                ' Do nothing for now

            End Try

        End Using

        Return strResult

    End Function

    ''' <summary>
    ''' Imports all outstanding fabric samples from the named website database into the production system
    ''' </summary>
    ''' <param name="strWebsite">Website to retrieve the setting from</param>
    ''' <remarks>Author: Huw Day - 05/02/2018</remarks>
    Function ImportFabricSamples(strWebsite As String) As String


        ' --- Defining objects and variables ---
        Dim dtFabricSampleHeaders As DataTable
        Dim dtFabricSampleLines As DataTable
        Dim strSiteName As String = ""
        Dim strCouponCode As String
        Dim intCouponDiscount As Integer
        Dim dateCouponExpiry As Date
        Dim strOutput As String = ""


        If strWebsite = enumWebsites.BPPE_UK.ToString() Then
            ' skip
        Else

            ' --- Defining query objects ---
            Dim query As String = "INSERT INTO FabricSampleOrders([FabricSampleOrderDate], [OrderOrigination], OriginatingOrderNumber, " _
                & "CustomerName, Add1, Add2, Add3, Add4, Add5, Email, Phone, CouponCode, CouponExpiry, CouponPercent, " _
                & "SampleCode1, SampleCode2, SampleCode3, SampleCode4, SampleCode5, SampleCode6, SampleCode7, SampleCode8, SampleCode9, SampleCode10, " _
                & "SampleDescription1, SampleDescription2, SampleDescription3, SampleDescription4, SampleDescription5, " _
                & "SampleDescription6, SampleDescription7, SampleDescription8, SampleDescription9, SampleDescription10) " _
                & "VALUES (@FabricSampleOrderDate, @OrderOrigination, @OriginatingOrderNumber, @CustomerName, @Add1, @Add2, @Add3, @Add4, @Add5, @Email, " _
                & "@Phone, @CouponCode, @CouponExpiry, @CouponPercent, " _
                & "@SampleCode1, @SampleCode2, @SampleCode3, @SampleCode4, @SampleCode5, @SampleCode6, @SampleCode7, @SampleCode8, @SampleCode9, @SampleCode10, " _
                & "@SampleDescription1, @SampleDescription2, @SampleDescription3, @SampleDescription4, @SampleDescription5, " _
                & "@SampleDescription6, @SampleDescription7, @SampleDescription8, @SampleDescription9, @SampleDescription10 " _
                & ");"
            Dim cmd As SqlCommand

            ' --- Defining sql parameters ---
            Dim paramOrderDate As SqlParameter
            Dim paramOrderOriginSite As SqlParameter
            Dim paramOrderOriginNumber As SqlParameter
            Dim paramCustomerName As SqlParameter
            Dim paramAddress1 As SqlParameter
            Dim paramAddress2 As SqlParameter
            Dim paramAddress3 As SqlParameter
            Dim paramAddress4 As SqlParameter
            Dim paramAddress5 As SqlParameter
            Dim paramEmail As SqlParameter
            Dim paramPhone As SqlParameter
            Dim paramCouponCode As SqlParameter
            Dim paramCouponExpiry As SqlParameter
            Dim paramCouponPercent As SqlParameter

            Dim lstParamsSampleCodes As List(Of SqlParameter)
            Dim lstParamsSampleNames As List(Of SqlParameter)


            ' --- pre-calculating ProdSys site name ---
            If strWebsite = "UK" Then
                strSiteName = "BLOCUKWebsite"
            ElseIf strWebsite = "IE" Then
                strSiteName = "BlocBlindsIE"
            Else
                strSiteName = "BlocBlinds" & strWebsite
            End If


            ' --- Looping through each order header that has fabric samples ---
            Try
                dtFabricSampleHeaders = Get_FabricSamplesToImport_OrderHeaders(strWebsite)
                For Each drHeaderRow As DataRow In dtFabricSampleHeaders.Rows

                    ' --- pre-calculating fabric sample coupon information ---
                    strCouponCode = Importer.Get_FabricSamples_DiscountCode(CInt(drHeaderRow("Order_No").ToString), strWebsite)
                    intCouponDiscount = Importer.Get_FabricSamples_Discount(strWebsite)
                    dateCouponExpiry = Importer.Get_FabricSamples_DiscountExpiry()

                    ' --- Getting all fabric sample order lines for this order ---
                    dtFabricSampleLines = Get_FabricSamplesToImport_OrderLines(strWebsite, CInt(drHeaderRow("Order_No")))
                    If dtFabricSampleLines.Rows.Count > 0 Then

                        ' --- Loop through each set of 10 fabric samples in the order ---
                        For i = 0 To dtFabricSampleLines.Rows.Count - 1 Step 10

                            ' --- Create command and header-level parameters ---
                            cmd = New SqlCommand
                            lstParamsSampleCodes = New List(Of SqlParameter)
                            lstParamsSampleNames = New List(Of SqlParameter)
                            cmd.CommandText = query
                            paramOrderDate = New SqlParameter("@FabricSampleOrderDate", drHeaderRow("DateEntered"))
                            paramOrderOriginSite = New SqlParameter("@OrderOrigination", strSiteName)
                            paramOrderOriginNumber = New SqlParameter("@OriginatingOrderNumber", drHeaderRow("Order_No").ToString)
                            paramCustomerName = New SqlParameter("@CustomerName", drHeaderRow("Delivery_FirstName") & " " & drHeaderRow("Delivery_LastName"))
                            paramAddress1 = New SqlParameter("@Add1", drHeaderRow("Delivery_HouseNumber"))
                            paramAddress2 = New SqlParameter("@Add2", drHeaderRow("Delivery_Address"))
                            paramAddress3 = New SqlParameter("@Add3", drHeaderRow("Delivery_City"))
                            paramAddress4 = New SqlParameter("@Add4", drHeaderRow("Delivery_Country"))
                            paramAddress5 = New SqlParameter("@Add5", drHeaderRow("Delivery_PostCode"))
                            paramEmail = New SqlParameter("@Email", drHeaderRow("Billing_Email"))
                            paramPhone = New SqlParameter("@Phone", drHeaderRow("Billing_PhoneNo"))
                            If SAMPLE_DISCOUNT_ON = True Then
                                paramCouponCode = New SqlParameter("@CouponCode", strCouponCode)
                                paramCouponExpiry = New SqlParameter("@CouponExpiry", DateSerial(dateCouponExpiry.Year, dateCouponExpiry.Month, dateCouponExpiry.Day))  ' Remove the time element from the supplied date
                                paramCouponPercent = New SqlParameter("@CouponPercent", intCouponDiscount)
                            Else
                                paramCouponCode = New SqlParameter("@CouponCode", DBNull.Value)
                                paramCouponExpiry = New SqlParameter("@CouponExpiry", DBNull.Value)
                                paramCouponPercent = New SqlParameter("@CouponPercent", DBNull.Value)
                            End If
                            cmd.Parameters.Add(paramOrderDate)
                            cmd.Parameters.Add(paramOrderOriginSite)
                            cmd.Parameters.Add(paramOrderOriginNumber)
                            cmd.Parameters.Add(paramCustomerName)
                            cmd.Parameters.Add(paramAddress1)
                            cmd.Parameters.Add(paramAddress2)
                            cmd.Parameters.Add(paramAddress3)
                            cmd.Parameters.Add(paramAddress4)
                            cmd.Parameters.Add(paramAddress5)
                            cmd.Parameters.Add(paramEmail)
                            cmd.Parameters.Add(paramPhone)
                            cmd.Parameters.Add(paramCouponCode)
                            cmd.Parameters.Add(paramCouponExpiry)
                            cmd.Parameters.Add(paramCouponPercent)

                            ' --- Create fabric sample parameters ---
                            For j = 1 To 10
                                If (i + j) <= dtFabricSampleLines.Rows.Count Then
                                    ' Write in parameter values here
                                    lstParamsSampleCodes.Add(New SqlParameter("@SampleCode" & j.ToString, dtFabricSampleLines.Rows(j - 1)("FabricCode")))
                                    lstParamsSampleNames.Add(New SqlParameter("@SampleDescription" & j.ToString, dtFabricSampleLines.Rows(j - 1)("FabricName")))
                                Else
                                    ' Write in null values for parameters here
                                    lstParamsSampleCodes.Add(New SqlParameter("@SampleCode" & j.ToString, DBNull.Value))
                                    lstParamsSampleNames.Add(New SqlParameter("@SampleDescription" & j.ToString, DBNull.Value))
                                End If
                                cmd.Parameters.Add(lstParamsSampleCodes(j - 1))
                                cmd.Parameters.Add(lstParamsSampleNames(j - 1))
                            Next

                            ' --- Execute query here. 1 line of 10 samples will be added. ---
                            Using conn As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("PRODSYS").ConnectionString)
                                cmd.Connection = conn

                                Try
                                    conn.Open()
                                    cmd.ExecuteNonQuery()
                                Catch ex As Exception
                                    Throw ex
                                End Try

                            End Using

                        Next    ' Loop to next set of 10 samples within this order


                        ' --- Write sample code into the website's Offers_PercentOff table ---
                        If SAMPLE_DISCOUNT_ON = True Then
                            Write_FabricSamples_DiscountToWebsiteDB(strWebsite, strCouponCode, intCouponDiscount, dateCouponExpiry)
                        End If

                    End If
                Next    ' Loop to next order

                If dtFabricSampleHeaders.Rows.Count <= 0 Then
                    strOutput = "No fabric samples to import"
                Else
                    strOutput = dtFabricSampleHeaders.Rows.Count & " fabric samples imported"
                End If
            Catch ex As Exception
                strOutput = "Error importing fabric samples: " & ex.Message
            End Try
        End If


        ' --- Returning result ---
        Return strOutput

    End Function


    Function ImportLock_Release() As ResultBoolean


        ' --- Defining objects and variables ---
        Dim blnResult As Boolean = False                                ' Fail-safe value
        Dim strOutputMessage As String = "Unknown error occurred"       ' Fail-safe value
        Dim sbQuery As New StringBuilder
        Dim dtOutput As New DataTable
        Dim da As SqlDataAdapter


        ' --- Writing out query ---
        sbQuery.AppendLine("UPDATE Settings SET SettingValue = '' WHERE SettingName = 'IMPORTINGLOCK' AND (SettingValue = '' OR SettingValue = Null OR SettingValue LIKE '" & PRODSYS_LOCK_USERNAME & "%'); ")
        sbQuery.AppendLine("IF @@ROWCOUNT >= 1 ")
        sbQuery.AppendLine("	SELECT 1 As Result, 'Database lock released' As [Message]; ")
        sbQuery.AppendLine("ELSE ")
        sbQuery.AppendLine("	SELECT 0 As Result, CONCAT('Could release database lock. Currently locked by ', (SELECT SettingValue FROM Settings WHERE SettingName =  'IMPORTINGLOCK')) As [Message]; ")


        ' --- Executing command ---
        Using connProdSys As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("PRODSYS").ConnectionString)
            da = New SqlDataAdapter(sbQuery.ToString(), connProdSys)
            da.SelectCommand.Parameters.Add(New SqlParameter("@LockValue", PRODSYS_LOCK_USERNAME & "-" & Now().ToString()))
            Try
                da.Fill(dtOutput)
                If IsNothing(dtOutput) = False Then
                    If dtOutput.Columns.Contains("Result") And dtOutput.Columns.Contains("Message") Then
                        If dtOutput.Rows.Count >= 1 Then
                            If dtOutput.Rows(0)("Result").ToString() = "1" Then
                                blnResult = True
                            Else
                                blnResult = False
                            End If
                            strOutputMessage = dtOutput.Rows(0)("Message").ToString()
                        Else
                            blnResult = False
                            strOutputMessage = "No data was received from database"
                        End If
                    Else
                        blnResult = False
                        strOutputMessage = "Invalid response received from database"
                    End If
                Else
                    blnResult = False
                    strOutputMessage = "No response received from database"
                End If
            Catch ex As Exception
                blnResult = False
                strOutputMessage = ex.Message & vbNewLine & ex.StackTrace
            Finally
                da.Dispose()
            End Try
        End Using


        ' --- Returning result ---
        Return New ResultBoolean(blnResult, strOutputMessage)

    End Function


    Function ImportLock_Set() As ResultBoolean


        ' --- Defining objects and variables ---
        Dim blnResult As Boolean = False                                ' Fail-safe value
        Dim strOutputMessage As String = "Unknown error occurred"       ' Fail-safe value
        Dim sbQuery As New StringBuilder
        Dim dtOutput As New DataTable
        Dim da As SqlDataAdapter


        ' --- Writing out query ---
        sbQuery.AppendLine("IF EXISTS(SELECT 1 FROM Settings WHERE SettingName = 'IMPORTINGLOCK' AND (SettingValue = '' OR SettingValue = Null OR SettingValue LIKE '" & PRODSYS_LOCK_USERNAME & "%')) ")
        sbQuery.AppendLine("BEGIN ")
        sbQuery.AppendLine("	UPDATE Settings SET SettingValue = @LockValue WHERE SettingName = 'IMPORTINGLOCK'; ")
        sbQuery.AppendLine("	IF (SELECT @@ROWCOUNT) >= 1 ")
        sbQuery.AppendLine("		SELECT 1 As Result, 'Database lock set' As [Message]; ")
        sbQuery.AppendLine("	ELSE ")
        sbQuery.AppendLine("		SELECT 0 As Result, 'Could not set database lock' As [Message]; ")
        sbQuery.AppendLine("END ")
        sbQuery.AppendLine("ELSE ")
        sbQuery.AppendLine("SELECT 0 As Result, CONCAT('Database is locked for importing: ',(SELECT SettingValue FROM Settings WHERE SettingName = 'IMPORTINGLOCK')) As [Message]; ")


        ' --- Executing command ---
        Using connProdSys As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("PRODSYS").ConnectionString)
            da = New SqlDataAdapter(sbQuery.ToString(), connProdSys)
            da.SelectCommand.Parameters.Add(New SqlParameter("@LockValue", PRODSYS_LOCK_USERNAME & "-" & Now().ToString()))
            Try
                da.Fill(dtOutput)
                If IsNothing(dtOutput) = False Then
                    If dtOutput.Columns.Contains("Result") And dtOutput.Columns.Contains("Message") Then
                        If dtOutput.Rows.Count >= 1 Then
                            If dtOutput.Rows(0)("Result").ToString() = "1" Then
                                blnResult = True
                            Else
                                blnResult = False
                            End If
                            strOutputMessage = dtOutput.Rows(0)("Message").ToString()
                        Else
                            blnResult = False
                            strOutputMessage = "No data was received from database"
                        End If
                    Else
                        blnResult = False
                        strOutputMessage = "Invalid response received from database"
                    End If
                Else
                    blnResult = False
                    strOutputMessage = "No response received from database"
                End If
            Catch ex As Exception
                blnResult = False
                strOutputMessage = ex.Message & vbNewLine & ex.StackTrace
            Finally
                da.Dispose()
            End Try
        End Using


        ' --- Returning result ---
        Return New ResultBoolean(blnResult, strOutputMessage)

    End Function


    ''' <summary>
    ''' This function acts as a hub for different websites' importer functions.
    ''' </summary>
    ''' <param name="strWebsite"></param>
    ''' <returns></returns>
    Function ImportOrders(strWebsite As String) As List(Of String)


        ' --- Defining objects and variables ---
        Dim lstOutput As New List(Of String)


        ' --- Selecting appropriate importer to run ---
        Select Case strWebsite.ToUpper()
            Case enumWebsites.UK.ToString().ToUpper(), enumWebsites.IE.ToString().ToUpper(), enumWebsites.TEST.ToString().ToUpper()
                lstOutput = ImportOrders_BlocBlinds(strWebsite)
            Case enumWebsites.BPPE_UK.ToString().ToUpper()
                lstOutput = ImportOrders_BlocFaceShields(strWebsite)
            Case Else
                ' Skip
        End Select


        ' --- Returning list of orders imported ---
        Return lstOutput


    End Function


    ''' <summary>
    ''' Imports all outstanding orders from the named database into the production system
    ''' </summary>
    ''' <remarks>Author: Huw Day - 05/02/2018</remarks>
    Function ImportOrders_BlocBlinds(strWebsite As String) As List(Of String)


        ' --- Initialising global objects and variables ---
        Initialise_ClassVariables()


        ' --- Defining objects and variables ---
        Dim dtOrderHeaders As New DataTable
        Dim da As SqlDataAdapter
        Dim paramProdSysID As SqlParameter
        Dim paramOrderNo As SqlParameter
        Dim cmdGetOrdHeaders As SqlCommand
        Dim cmdImportTransaction As SqlCommand
        Dim cmdUpdateWebDb As SqlCommand
        Dim strGetOrdHeadersQuery As String
        Dim strUpdateWebDBQuery As String
        Dim strResult As String
        Dim sbTransactionSql As New StringBuilder
        Dim lstOutput As New List(Of String)
        Dim blnAbort As Boolean = False
        Dim orderLeadTime As OrderLeadTime


        ' --- Writing out queries ---
        Dim sbQuery As New System.Text.StringBuilder
        sbQuery.Append("SELECT Order_No, DateEntered, QuotedDeliveryDate, Billing_Email, Billing_PhoneNo, Delivery_FirstName, ")
        sbQuery.Append("Delivery_LastName, Delivery_HouseNumber, Delivery_Address, Delivery_City, Delivery_Country, Delivery_PostCode, DeliveryCost, FBDaTOO ")
        sbQuery.Append("FROM ORD_HEADER ")
        sbQuery.Append("WHERE OrderStatus = 'PAY_AUTH' And ProductionSystemID IS NULL ")
        sbQuery.Append("AND ((PaymentType <> 'FREE') ")
        sbQuery.Append("AND Exists (SELECT * FROM ORD_LINES WHERE ORD_LINES.Order_No = ORD_HEADER.Order_No AND ProductName <> 'MeasureService') ")                              ' Qualifier to exclude those paid orders whose only order line is a Measure Service
        sbQuery.Append("OR (PaymentType = 'FREE' AND EXISTS(SELECT * FROM ORD_LINES WHERE ORD_LINES.Order_No = ORD_HEADER.Order_No AND ProductName = 'VirtualFabric')))")       ' Qualifier to include those orders that are free, but have a virtual fabric to be imported
        sbQuery.Append(";")
        strGetOrdHeadersQuery = sbQuery.ToString

        strUpdateWebDBQuery = "UPDATE ORD_HEADER SET ProductionSystemID = @ProdSysID WHERE Order_No = @OrderNo;"


        ' --- Constructing insert transaction template ---
        sbTransactionSql.AppendLine("BEGIN TRANSACTION [OrdImport1] ")
        sbTransactionSql.AppendLine("BEGIN TRY ")
        sbTransactionSql.AppendLine("DECLARE @ProdSysID int; ")
        sbTransactionSql.AppendLine("/*ORD_HEADER_BLOCK*/ ")
        sbTransactionSql.AppendLine("SET @ProdSysID = SCOPE_IDENTITY(); ")
        sbTransactionSql.AppendLine("/*ORD_LINE_BLOCKS*/ ")
        sbTransactionSql.AppendLine("SELECT @ProdSysID; ")
        sbTransactionSql.AppendLine("COMMIT TRANSACTION [OrdImport1] ")
        sbTransactionSql.AppendLine("END TRY")
        sbTransactionSql.AppendLine("BEGIN CATCH ")
        sbTransactionSql.AppendLine("ROLLBACK TRANSACTION [OrdImport1]; ")
        sbTransactionSql.AppendLine("SELECT ERROR_MESSAGE(); ")
        sbTransactionSql.AppendLine("END CATCH ")
        'sbTransactionSql.AppendLine("GO")


        ' --- Retrieving Order Headers to import ---
        Using conn As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings(strWebsite).ConnectionString)
            cmdGetOrdHeaders = New SqlCommand(strGetOrdHeadersQuery, conn)
            da = New SqlDataAdapter
            da.SelectCommand = cmdGetOrdHeaders
            Try
                da.Fill(dtOrderHeaders)
            Catch ex As Exception
                blnAbort = True
                lstOutput.Insert(0, "Could not retrieve orders to import. Database not available.")
            End Try
        End Using


        ' --- Constructing an Insert SQL Transaction for each order ---
        If blnAbort = False Then
            If dtOrderHeaders.Rows.Count <= 0 Then
                lstOutput.Insert(0, "No orders to import")
            Else
                For Each drOrderHeader In dtOrderHeaders.Rows

                    ' --- Setting default information ---
                    class_blnIsSparesOrder = True       ' True by default. If any product is found in the order that is not in the Spares list, this is set to False


                    ' --- Creating SQL command ---
                    cmdImportTransaction = New SqlCommand
                    cmdImportTransaction.CommandText = sbTransactionSql.ToString


                    ' --- Adding order lines SQL and parameters to command ---
                    orderLeadTime = Build_OrderLines(cmdImportTransaction, drOrderHeader, strWebsite)


                    ' --- Adding order header SQL and parameters to command ---
                    Build_OrderHeader(cmdImportTransaction, drOrderHeader, strWebsite, orderLeadTime)


                    ' --- Modifying ORD_HEADER parameters based on lines data ---
                    If class_blnIsSparesOrder = True Then
                        Dim intOnHoldIndex As Integer = cmdImportTransaction.Parameters.IndexOf("@OnHold_OH")
                        Dim intPreferredCourierIndex As Integer = cmdImportTransaction.Parameters.IndexOf("@PreferredCourier_OH")
                        Dim paramOnHold As SqlParameter = Nothing
                        Dim paramPreferredCourier As SqlParameter = Nothing
                        If intOnHoldIndex >= 0 And intPreferredCourierIndex >= 0 Then
                            paramOnHold = cmdImportTransaction.Parameters(intOnHoldIndex)
                            paramPreferredCourier = cmdImportTransaction.Parameters(intPreferredCourierIndex)
                        End If
                        If IsNothing(paramOnHold) Then
                            Console.WriteLine("Could not retrieve OnHold parameter object")
                        ElseIf IsNothing(paramPreferredCourier) Then
                            Console.WriteLine("Could not retrieve OnHold parameter object")
                        Else
                            paramOnHold.Value = True
                            paramPreferredCourier.Value = "POST"
                        End If
                    End If


                    ' --- Executing Insert Transaction ---
                    Using connProdSys As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("PRODSYS").ConnectionString)
                        Try
                            cmdImportTransaction.Connection = connProdSys
                            connProdSys.Open()
                            strResult = cmdImportTransaction.ExecuteScalar
                        Catch ex As Exception
                            'Throw ex
                            ' ## Write the Error To the console Or otherwise log it ##
                            strResult = ex.Message
                        End Try
                    End Using


                    ' --- Analysing result ---
                    If IsNumeric(strResult) Then
                        ' ## Add ProdSys ID to the website's database ##
                        Using conn As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings(strWebsite).ConnectionString)
                            cmdUpdateWebDb = New SqlCommand(strUpdateWebDBQuery, conn)
                            paramOrderNo = New SqlParameter("@OrderNo", drOrderHeader("Order_No"))
                            paramProdSysID = New SqlParameter("@ProdSysID", strResult)
                            cmdUpdateWebDb.Parameters.Add(paramOrderNo)
                            cmdUpdateWebDb.Parameters.Add(paramProdSysID)
                            Try
                                conn.Open()
                                cmdUpdateWebDb.ExecuteNonQuery()
                                lstOutput.Insert(0, "Order " & drOrderHeader("Order_No").ToString & " imported")
                                Importer.Log_OrderImport(strWebsite, CInt(drOrderHeader("Order_No")), "imported")
                            Catch ex As Exception
                                'Throw ex
                                lstOutput.Insert(0, "Order " & drOrderHeader("Order_No").ToString & ex.Message)
                                Importer.Log_OrderImport(strWebsite, CInt(drOrderHeader("Order_No")), ex.Message)
                            End Try
                        End Using

                    Else
                        lstOutput.Insert(0, "Order " & drOrderHeader("Order_No").ToString & "The SQL Server returned the following error: " & strResult)
                        Importer.Log_OrderImport(strWebsite, CInt(drOrderHeader("Order_No")), strResult)
                    End If

                Next
            End If
        End If


        ' --- Returning result ---
        Return lstOutput

    End Function


    ''' <summary>
    ''' Imports all outstanding orders from the named database into the production system
    ''' </summary>
    ''' <param name="strWebsite"></param>
    ''' <returns></returns>
    Function ImportOrders_BlocFaceShields(strWebsite As String) As List(Of String)


        ' --- Initialising global objects and variables ---
        Initialise_ClassVariables()


        ' --- Defining objects and variables ---
        Const DELIVERY_COST As String = "7.95"      ' Hard-coded in pro-tem
        Dim dtOrderHeaders As New DataTable
        Dim da As SqlDataAdapter
        Dim paramProdSysID As SqlParameter
        Dim paramOrderNo As SqlParameter
        Dim cmdGetOrdHeaders As SqlCommand
        Dim cmdImportTransaction As SqlCommand
        Dim cmdUpdateWebDb As SqlCommand
        Dim strGetOrdHeadersQuery As String
        Dim strUpdateWebDBQuery As String
        Dim strResult As String
        Dim sbTransactionSql As New StringBuilder
        Dim lstOutput As New List(Of String)
        Dim blnAbort As Boolean = False


        ' --- Writing out queries ---
        ' -- Get Order Header query
        Dim sbQuery As New System.Text.StringBuilder
        sbQuery.Append("SELECT Order_No, DateEntered, Billing_Email, Billing_PhoneNo, Delivery_FirstName, Delivery_LastName, ")
        sbQuery.Append("Delivery_HouseNumber, Delivery_Address, Delivery_City, Delivery_Country, Delivery_PostCode, CAST('" & DELIVERY_COST & "' As decimal(18,2)) As Delivery_Cost ")
        sbQuery.Append("FROM ORD_HEADER ")
        sbQuery.Append("WHERE OrderStatus = 'PAY_AUTH' And ProductionSystemID IS NULL ")
        sbQuery.Append("AND Exists (SELECT * FROM ORD_LINES WHERE ORD_LINES.Order_No = ORD_HEADER.Order_No) ")
        sbQuery.Append(";")
        strGetOrdHeadersQuery = sbQuery.ToString()

        ' -- Update Order Header with production system ID query
        strUpdateWebDBQuery = "UPDATE ORD_HEADER SET ProductionSystemID = @ProdSysID WHERE Order_No = @OrderNo;"


        ' --- Constructing insert transaction template ---
        sbTransactionSql.AppendLine("BEGIN TRANSACTION [OrdImport1] ")
        sbTransactionSql.AppendLine("BEGIN TRY ")
        sbTransactionSql.AppendLine("DECLARE @ProdSysID int; ")
        sbTransactionSql.AppendLine("/*ORD_HEADER_BLOCK*/ ")
        sbTransactionSql.AppendLine("SET @ProdSysID = SCOPE_IDENTITY(); ")
        sbTransactionSql.AppendLine("/*ORD_LINE_BLOCKS*/ ")
        sbTransactionSql.AppendLine("SELECT @ProdSysID; ")
        sbTransactionSql.AppendLine("COMMIT TRANSACTION [OrdImport1] ")
        sbTransactionSql.AppendLine("END TRY")
        sbTransactionSql.AppendLine("BEGIN CATCH ")
        sbTransactionSql.AppendLine("ROLLBACK TRANSACTION [OrdImport1]; ")
        sbTransactionSql.AppendLine("SELECT ERROR_MESSAGE(); ")
        sbTransactionSql.AppendLine("END CATCH ")


        ' --- Retrieving Order Headers to import ---
        Using conn As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings(strWebsite).ConnectionString)
            cmdGetOrdHeaders = New SqlCommand(strGetOrdHeadersQuery, conn)
            da = New SqlDataAdapter
            da.SelectCommand = cmdGetOrdHeaders
            Try
                da.Fill(dtOrderHeaders)
            Catch ex As Exception
                blnAbort = True
                lstOutput.Insert(0, "Could not retrieve orders to import. Database not available.")
            End Try
        End Using


        ' --- Constructing an Insert SQL transaction for each order ---
        If blnAbort = False Then
            If dtOrderHeaders.Rows.Count <= 0 Then
                lstOutput.Insert(0, "No orders to import")
            Else
                For Each drOrderHeader As DataRow In dtOrderHeaders.Rows

                    ' --- Creating SQL command ---
                    cmdImportTransaction = New SqlCommand
                    cmdImportTransaction.CommandText = sbTransactionSql.ToString()


                    ' --- Adding Order Header SQL and parameters to command ---
                    Build_OrderHeader_BFS(cmdImportTransaction, drOrderHeader, strWebsite)


                    ' --- Adding Order Lines SQL and parameters to command ---
                    Build_OrderLines_BFS(cmdImportTransaction, drOrderHeader, strWebsite)


                    ' --- Executing Insert Transaction ---
                    Using connProdSys As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("PRODSYS").ConnectionString)
                        Try
                            cmdImportTransaction.Connection = connProdSys
                            connProdSys.Open()
                            strResult = cmdImportTransaction.ExecuteScalar
                        Catch ex As Exception
                            strResult = ex.Message
                        End Try
                    End Using


                    ' --- Analysing result ---
                    If IsNumeric(strResult) Then
                        ' ## Add ProdSys ID to the website's database ##
                        Using conn As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings(strWebsite).ConnectionString)
                            cmdUpdateWebDb = New SqlCommand(strUpdateWebDBQuery, conn)
                            paramOrderNo = New SqlParameter("@OrderNo", drOrderHeader("Order_No"))
                            paramProdSysID = New SqlParameter("@ProdSysID", strResult)
                            cmdUpdateWebDb.Parameters.Add(paramOrderNo)
                            cmdUpdateWebDb.Parameters.Add(paramProdSysID)
                            Try
                                conn.Open()
                                cmdUpdateWebDb.ExecuteNonQuery()
                                lstOutput.Insert(0, "Order " & drOrderHeader("Order_No").ToString & " imported")
                                Importer.Log_OrderImport(strWebsite, CInt(drOrderHeader("Order_No")), "imported")
                            Catch ex As Exception
                                'Throw ex
                                lstOutput.Insert(0, "Order " & drOrderHeader("Order_No").ToString & ex.Message)
                                Importer.Log_OrderImport(strWebsite, CInt(drOrderHeader("Order_No")), ex.Message)
                            End Try
                        End Using

                    Else
                        lstOutput.Insert(0, "Order " & drOrderHeader("Order_No").ToString & "The SQL Server returned the following error: " & strResult)
                        Importer.Log_OrderImport(strWebsite, CInt(drOrderHeader("Order_No")), strResult)
                    End If

                Next
            End If
        End If


        ' --- Returning result ---
        Return lstOutput


    End Function


    ''' <summary>
    ''' This subroutine enters any required data into class-level objects. Run at the start of import
    ''' </summary>
    ''' <remarks>Author: Huw Day - 05/02/2018</remarks>
    Sub Initialise_ClassVariables()

        lst_ProdSys_ProductNames = New Dictionary(Of String, String)           ' Translates Website Product names into ProdSys product names
        lst_ProdSys_ItemCodePrefix = New Dictionary(Of String, String)         ' Translates Website Product names into ProdSys item code prefixes
        lst_ProdSys_LineType = New Dictionary(Of String, String)               ' Translates Website Product names into ProdSys order line types
        lst_ProdSys_ManufacturerCodes = New Dictionary(Of String, String)      ' Translates website manufacturer names into 2 character manufacturer codes used in item codes
        lst_ProdSys_FittingTypes = New Dictionary(Of String, String)           ' Translates website product fitting types into ProdSys fitting types
        lst_SpareParts = New List(Of String)                                   ' A list of product names that are Spare Parts


        ' --- Creating Item Code Prefix translation list ---
        lst_ProdSys_ItemCodePrefix.Add("additionalfabric".ToLower(), "BRB-AF")
        lst_ProdSys_ItemCodePrefix.Add("blocout40".ToLower(), "BLOCOUT40")
        lst_ProdSys_ItemCodePrefix.Add("blocout80".ToLower(), "BLOCOUT80")
        lst_ProdSys_ItemCodePrefix.Add("BO40_Brakes_Cream".ToLower(), "")         ' The prefix for this product is pulled from the Product_ItemCode field of the website's ORD_LINES table in the Build_OrderLine_Product function
        lst_ProdSys_ItemCodePrefix.Add("BO40_Brakes_Grey".ToLower(), "")          ' The prefix for this product is pulled from the Product_ItemCode field of the website's ORD_LINES table in the Build_OrderLine_Product function
        lst_ProdSys_ItemCodePrefix.Add("BO40_Brakes_White".ToLower(), "")         ' The prefix for this product is pulled from the Product_ItemCode field of the website's ORD_LINES table in the Build_OrderLine_Product function
        lst_ProdSys_ItemCodePrefix.Add("BO40_Fittings".ToLower(), "")             ' The prefix for this product is pulled from the Product_ItemCode field of the website's ORD_LINES table in the Build_OrderLine_Product function
        lst_ProdSys_ItemCodePrefix.Add("BO80_Fittings_Recess".ToLower(), "")      ' The prefix for this product is pulled from the Product_ItemCode field of the website's ORD_LINES table in the Build_OrderLine_Product function
        lst_ProdSys_ItemCodePrefix.Add("BO80_Fittings_Surface".ToLower(), "")     ' The prefix for this product is pulled from the Product_ItemCode field of the website's ORD_LINES table in the Build_OrderLine_Product function
        lst_ProdSys_ItemCodePrefix.Add("customskylight".ToLower(), "NON-STANDARD")
        lst_ProdSys_ItemCodePrefix.Add("fabricsample".ToLower(), "-")
        lst_ProdSys_ItemCodePrefix.Add("faceshield_1b12".ToLower(), "")           ' The prefix for this product is pulled from the Product_ItemCode field of the website's ORD_LINES table in the Build_OrderLine_Product function
        lst_ProdSys_ItemCodePrefix.Add("faceshield_5b12".ToLower(), "")           ' The prefix for this product is pulled from the Product_ItemCode field of the website's ORD_LINES table in the Build_OrderLine_Product function
        lst_ProdSys_ItemCodePrefix.Add("faceshield_20b12".ToLower(), "")          ' The prefix for this product is pulled from the Product_ItemCode field of the website's ORD_LINES table in the Build_OrderLine_Product function
        lst_ProdSys_ItemCodePrefix.Add("faceshieldkit_1b1".ToLower(), "")         ' The prefix for this product is pulled from the Product_ItemCode field of the website's ORD_LINES table in the Build_OrderLine_Product function
        lst_ProdSys_ItemCodePrefix.Add("faceshieldkit_1b2".ToLower(), "")         ' The prefix for this product is pulled from the Product_ItemCode field of the website's ORD_LINES table in the Build_OrderLine_Product function
        lst_ProdSys_ItemCodePrefix.Add("faceshieldkit_1b3".ToLower(), "")         ' The prefix for this product is pulled from the Product_ItemCode field of the website's ORD_LINES table in the Build_OrderLine_Product function
        lst_ProdSys_ItemCodePrefix.Add("faceshieldkit_1b4".ToLower(), "")         ' The prefix for this product is pulled from the Product_ItemCode field of the website's ORD_LINES table in the Build_OrderLine_Product function
        lst_ProdSys_ItemCodePrefix.Add("faceshieldkit_1b5".ToLower(), "")         ' The prefix for this product is pulled from the Product_ItemCode field of the website's ORD_LINES table in the Build_OrderLine_Product function
        lst_ProdSys_ItemCodePrefix.Add("faceshieldkit_1b12".ToLower(), "")        ' The prefix for this product is pulled from the Product_ItemCode field of the website's ORD_LINES table in the Build_OrderLine_Product function
        lst_ProdSys_ItemCodePrefix.Add("faceshieldkit_1b24".ToLower(), "")        ' The prefix for this product is pulled from the Product_ItemCode field of the website's ORD_LINES table in the Build_OrderLine_Product function
        lst_ProdSys_ItemCodePrefix.Add("faceshieldkit_1b40".ToLower(), "")        ' The prefix for this product is pulled from the Product_ItemCode field of the website's ORD_LINES table in the Build_OrderLine_Product function
        lst_ProdSys_ItemCodePrefix.Add("faceshieldkit_4b24".ToLower(), "")        ' The prefix for this product is pulled from the Product_ItemCode field of the website's ORD_LINES table in the Build_OrderLine_Product function
        lst_ProdSys_ItemCodePrefix.Add("measure protect".ToLower(), "MEASUREPROTECT")
        lst_ProdSys_ItemCodePrefix.Add("multichannelremote".ToLower(), "MISC-NA-NA")
        lst_ProdSys_ItemCodePrefix.Add("pelmet".ToLower(), "BPL")
        lst_ProdSys_ItemCodePrefix.Add("pole2m".ToLower(), "POLE")
        lst_ProdSys_ItemCodePrefix.Add("pole3m".ToLower(), "POLE")
        lst_ProdSys_ItemCodePrefix.Add("PRB_Fittings_Recess".ToLower(), "")       ' The prefix for this product is pulled from the Product_ItemCode field of the website's ORD_LINES table in the Build_OrderLine_Product function
        lst_ProdSys_ItemCodePrefix.Add("PRB_Fittings_Surface".ToLower(), "")      ' The prefix for this product is pulled from the Product_ItemCode field of the website's ORD_LINES table in the Build_OrderLine_Product function
        lst_ProdSys_ItemCodePrefix.Add("premiumrollerblind".ToLower(), "CRB")
        lst_ProdSys_ItemCodePrefix.Add("premiumskylight".ToLower(), "P")
        lst_ProdSys_ItemCodePrefix.Add("RB_Brackets_Black".ToLower(), "")         ' The prefix for this product is pulled from the Product_ItemCode field of the website's ORD_LINES table in the Build_OrderLine_Product function
        lst_ProdSys_ItemCodePrefix.Add("RB_Brackets_Grey".ToLower(), "")          ' The prefix for this product is pulled from the Product_ItemCode field of the website's ORD_LINES table in the Build_OrderLine_Product function
        lst_ProdSys_ItemCodePrefix.Add("RB_Brackets_White".ToLower(), "")         ' The prefix for this product is pulled from the Product_ItemCode field of the website's ORD_LINES table in the Build_OrderLine_Product function
        lst_ProdSys_ItemCodePrefix.Add("RB_DF_Brackets_White".ToLower(), "")      ' The prefix for this product is pulled from the Product_ItemCode field of the website's ORD_LINES table in the Build_OrderLine_Product function
        lst_ProdSys_ItemCodePrefix.Add("RB_M_Brackets_Black".ToLower(), "")       ' The prefix for this product is pulled from the Product_ItemCode field of the website's ORD_LINES table in the Build_OrderLine_Product function
        lst_ProdSys_ItemCodePrefix.Add("RB_M_Brackets_White".ToLower(), "")       ' The prefix for this product is pulled from the Product_ItemCode field of the website's ORD_LINES table in the Build_OrderLine_Product function
        lst_ProdSys_ItemCodePrefix.Add("rollerblind".ToLower(), "BRB")
        lst_ProdSys_ItemCodePrefix.Add("rollerblinddrillfree".ToLower(), "BRB-DF")
        lst_ProdSys_ItemCodePrefix.Add("rollerblindextend_black".ToLower(), "")   ' The prefix for this product is pulled from the Product_ItemCode field of the website's ORD_LINES table in the Build_OrderLine_Product function
        lst_ProdSys_ItemCodePrefix.Add("rollerblindextend_blue".ToLower(), "")    ' The prefix for this product is pulled from the Product_ItemCode field of the website's ORD_LINES table in the Build_OrderLine_Product function
        lst_ProdSys_ItemCodePrefix.Add("rollerblindextend_red".ToLower(), "")     ' The prefix for this product is pulled from the Product_ItemCode field of the website's ORD_LINES table in the Build_OrderLine_Product function
        lst_ProdSys_ItemCodePrefix.Add("rollerblindextend_green".ToLower(), "")   ' The prefix for this product is pulled from the Product_ItemCode field of the website's ORD_LINES table in the Build_OrderLine_Product function
        lst_ProdSys_ItemCodePrefix.Add("rollerblindmotorised".ToLower(), "BRB-M")
        lst_ProdSys_ItemCodePrefix.Add("singlechannelremote".ToLower(), "MISC-NA-NA")
        lst_ProdSys_ItemCodePrefix.Add("SKY_Brakes_Cream".ToLower(), "")          ' The prefix for this product is pulled from the Product_ItemCode field of the website's ORD_LINES table in the Build_OrderLine_Product function
        lst_ProdSys_ItemCodePrefix.Add("SKY_Brakes_Grey".ToLower(), "")           ' The prefix for this product is pulled from the Product_ItemCode field of the website's ORD_LINES table in the Build_OrderLine_Product function
        lst_ProdSys_ItemCodePrefix.Add("SKY_Brakes_White".ToLower(), "")          ' The prefix for this product is pulled from the Product_ItemCode field of the website's ORD_LINES table in the Build_OrderLine_Product function
        lst_ProdSys_ItemCodePrefix.Add("smarthub".ToLower(), "")                  ' The prefix for this product is pulled from the Product_ItemCode field of the website's ORD_LINES table in the Build_OrderLine_Product function
        lst_ProdSys_ItemCodePrefix.Add("smartmotorcharger".ToLower(), "")         ' The prefix for this product is pulled from the Product_ItemCode field of the website's ORD_LINES table in the Build_OrderLine_Product function
        lst_ProdSys_ItemCodePrefix.Add("smartremote".ToLower(), "")               ' The prefix for this product is pulled from the Product_ItemCode field of the website's ORD_LINES table in the Build_OrderLine_Product function
        lst_ProdSys_ItemCodePrefix.Add("smartrepeater".ToLower(), "")             ' The prefix for this product is pulled from the Product_ItemCode field of the website's ORD_LINES table in the Build_OrderLine_Product function
        lst_ProdSys_ItemCodePrefix.Add("solarskylight".ToLower(), "BS")
        lst_ProdSys_ItemCodePrefix.Add("SS_Battery_Post_03_19".ToLower(), "")     ' The prefix for this product is pulled from the Product_ItemCode field of the website's ORD_LINES table in the Build_OrderLine_Product function
        lst_ProdSys_ItemCodePrefix.Add("SS_Battery_Pre_03_19".ToLower(), "")      ' The prefix for this product is pulled from the Product_ItemCode field of the website's ORD_LINES table in the Build_OrderLine_Product function
        lst_ProdSys_ItemCodePrefix.Add("twinrollerblind".ToLower(), "BTRB")
        lst_ProdSys_ItemCodePrefix.Add("venetian".ToLower(), "BVN")
        lst_ProdSys_ItemCodePrefix.Add("virtualfabric".ToLower(), "BRBAF")
        lst_ProdSys_ItemCodePrefix.Add("zebrarollerblind".ToLower(), "BZRB")
        lst_ProdSys_ItemCodePrefix.Add("ZRB_Fittings_Recess".ToLower(), "")       ' The prefix for this product is pulled from the Product_ItemCode field of the website's ORD_LINES table in the Build_OrderLine_Product function
        lst_ProdSys_ItemCodePrefix.Add("ZRB_Fittings_Surface".ToLower(), "")      ' The prefix for this product is pulled from the Product_ItemCode field of the website's ORD_LINES table in the Build_OrderLine_Product function

        ' --- Creating Line Type translation list ---
        lst_ProdSys_LineType.Add("additionalfabric".ToLower(), "ADDITIONALFABRIC")
        lst_ProdSys_LineType.Add("blocout40".ToLower(), "BLOCOUT40")
        lst_ProdSys_LineType.Add("blocout80".ToLower(), "BLOCOUT80")
        lst_ProdSys_LineType.Add("BO40_Brakes_Cream".ToLower(), "COMPONENT")
        lst_ProdSys_LineType.Add("BO40_Brakes_Grey".ToLower(), "COMPONENT")
        lst_ProdSys_LineType.Add("BO40_Brakes_White".ToLower(), "COMPONENT")
        lst_ProdSys_LineType.Add("BO40_Fittings".ToLower(), "COMPONENT")
        lst_ProdSys_LineType.Add("BO80_Fittings_Recess".ToLower(), "COMPONENT")
        lst_ProdSys_LineType.Add("BO80_Fittings_Surface".ToLower(), "COMPONENT")
        lst_ProdSys_LineType.Add("customskylight".ToLower(), "CUSTOM-SKYLITE")
        lst_ProdSys_LineType.Add("fabricsample".ToLower(), "-")
        lst_ProdSys_LineType.Add("faceshield_1b12".ToLower(), "COMPONENT")
        lst_ProdSys_LineType.Add("faceshield_5b12".ToLower(), "COMPONENT")
        lst_ProdSys_LineType.Add("faceshield_20b12".ToLower(), "COMPONENT")
        lst_ProdSys_LineType.Add("faceshieldkit_1b1".ToLower(), "COMPONENT")
        lst_ProdSys_LineType.Add("faceshieldkit_1b2".ToLower(), "COMPONENT")
        lst_ProdSys_LineType.Add("faceshieldkit_1b3".ToLower(), "COMPONENT")
        lst_ProdSys_LineType.Add("faceshieldkit_1b4".ToLower(), "COMPONENT")
        lst_ProdSys_LineType.Add("faceshieldkit_1b5".ToLower(), "COMPONENT")
        lst_ProdSys_LineType.Add("faceshieldkit_1b12".ToLower(), "COMPONENT")
        lst_ProdSys_LineType.Add("faceshieldkit_1b24".ToLower(), "COMPONENT")
        lst_ProdSys_LineType.Add("faceshieldkit_1b40".ToLower(), "COMPONENT")
        lst_ProdSys_LineType.Add("faceshieldkit_4b24".ToLower(), "COMPONENT")
        lst_ProdSys_LineType.Add("measure protect".ToLower(), "MEASUREPROTECT")
        lst_ProdSys_LineType.Add("multichannelremote".ToLower(), "REMOTE")
        lst_ProdSys_LineType.Add("pelmet".ToLower(), "PELMET")
        lst_ProdSys_LineType.Add("pole2m".ToLower(), "ACCESSORIES")
        lst_ProdSys_LineType.Add("pole3m".ToLower(), "ACCESSORIES")
        lst_ProdSys_LineType.Add("PRB_Fittings_Recess".ToLower(), "COMPONENT")
        lst_ProdSys_LineType.Add("PRB_Fittings_Surface".ToLower(), "COMPONENT")
        lst_ProdSys_LineType.Add("premiumrollerblind".ToLower(), "PREMIUMROLLERBLIND")
        lst_ProdSys_LineType.Add("premiumskylight".ToLower(), "SKYLITE")
        lst_ProdSys_LineType.Add("RB_Brackets_Black".ToLower(), "COMPONENT")
        lst_ProdSys_LineType.Add("RB_Brackets_Grey".ToLower(), "COMPONENT")
        lst_ProdSys_LineType.Add("RB_Brackets_White".ToLower(), "COMPONENT")
        lst_ProdSys_LineType.Add("RB_DF_Brackets_White".ToLower(), "COMPONENT")
        lst_ProdSys_LineType.Add("RB_M_Brackets_Black".ToLower(), "COMPONENT")
        lst_ProdSys_LineType.Add("RB_M_Brackets_White".ToLower(), "COMPONENT")
        lst_ProdSys_LineType.Add("rollerblind".ToLower(), "ROLLERBLIND")
        lst_ProdSys_LineType.Add("rollerblinddrillfree".ToLower(), "ROLLERBLIND")
        lst_ProdSys_LineType.Add("rollerblindextend_black".ToLower(), "ROLLERBLIND")
        lst_ProdSys_LineType.Add("rollerblindextend_blue".ToLower(), "ROLLERBLIND")
        lst_ProdSys_LineType.Add("rollerblindextend_red".ToLower(), "ROLLERBLIND")
        lst_ProdSys_LineType.Add("rollerblindextend_green".ToLower(), "ROLLERBLIND")
        lst_ProdSys_LineType.Add("rollerblindmotorised".ToLower(), "ROLLERBLIND")
        lst_ProdSys_LineType.Add("smarthub".ToLower(), "COMPONENT")
        lst_ProdSys_LineType.Add("smartmotorcharger".ToLower(), "COMPONENT")
        lst_ProdSys_LineType.Add("smartremote".ToLower(), "COMPONENT")
        lst_ProdSys_LineType.Add("smartrepeater".ToLower(), "COMPONENT")
        lst_ProdSys_LineType.Add("singlechannelremote".ToLower(), "REMOTE")
        lst_ProdSys_LineType.Add("SKY_Brakes_Cream".ToLower(), "COMPONENT")
        lst_ProdSys_LineType.Add("SKY_Brakes_Grey".ToLower(), "COMPONENT")
        lst_ProdSys_LineType.Add("SKY_Brakes_White".ToLower(), "COMPONENT")
        lst_ProdSys_LineType.Add("solarskylight".ToLower(), "SOLARSKYLITE")
        lst_ProdSys_LineType.Add("SS_Battery_Post_03_19".ToLower(), "COMPONENT")
        lst_ProdSys_LineType.Add("SS_Battery_Pre_03_19".ToLower(), "COMPONENT")
        lst_ProdSys_LineType.Add("twinrollerblind".ToLower(), "ROLLERBLIND-TWIN")
        lst_ProdSys_LineType.Add("venetian".ToLower(), "VENETIAN")
        lst_ProdSys_LineType.Add("virtualfabric".ToLower(), "ADDITIONALFABRIC")
        lst_ProdSys_LineType.Add("zebrarollerblind".ToLower(), "ZEBRAROLLERBLIND")
        lst_ProdSys_LineType.Add("ZRB_Fittings_Recess".ToLower(), "COMPONENT")
        lst_ProdSys_LineType.Add("ZRB_Fittings_Surface".ToLower(), "COMPONENT")

        ' --- Creating ProdSys product name translation list ---
        lst_ProdSys_ProductNames.Add("additionalfabric".ToLower(), "rollerblind-additionalfabric")
        lst_ProdSys_ProductNames.Add("blocout40".ToLower(), "blocout40")
        lst_ProdSys_ProductNames.Add("blocout80".ToLower(), "blocout80")
        lst_ProdSys_ProductNames.Add("BO40_Brakes_Cream".ToLower(), "COMPONENT")
        lst_ProdSys_ProductNames.Add("BO40_Brakes_Grey".ToLower(), "COMPONENT")
        lst_ProdSys_ProductNames.Add("BO40_Brakes_White".ToLower(), "COMPONENT")
        lst_ProdSys_ProductNames.Add("BO40_Fittings".ToLower(), "COMPONENT")
        lst_ProdSys_ProductNames.Add("BO80_Fittings_Recess".ToLower(), "COMPONENT")
        lst_ProdSys_ProductNames.Add("BO80_Fittings_Surface".ToLower(), "COMPONENT")
        lst_ProdSys_ProductNames.Add("customskylight".ToLower(), "skylight-custom")
        lst_ProdSys_ProductNames.Add("fabricsample".ToLower(), "-")
        lst_ProdSys_ProductNames.Add("faceshield_1b12".ToLower(), "COMPONENT")
        lst_ProdSys_ProductNames.Add("faceshield_5b12".ToLower(), "COMPONENT")
        lst_ProdSys_ProductNames.Add("faceshield_20b12".ToLower(), "COMPONENT")
        lst_ProdSys_ProductNames.Add("faceshieldkit_1b1".ToLower(), "COMPONENT")
        lst_ProdSys_ProductNames.Add("faceshieldkit_1b2".ToLower(), "COMPONENT")
        lst_ProdSys_ProductNames.Add("faceshieldkit_1b3".ToLower(), "COMPONENT")
        lst_ProdSys_ProductNames.Add("faceshieldkit_1b4".ToLower(), "COMPONENT")
        lst_ProdSys_ProductNames.Add("faceshieldkit_1b5".ToLower(), "COMPONENT")
        lst_ProdSys_ProductNames.Add("faceshieldkit_1b12".ToLower(), "COMPONENT")
        lst_ProdSys_ProductNames.Add("faceshieldkit_1b24".ToLower(), "COMPONENT")
        lst_ProdSys_ProductNames.Add("faceshieldkit_1b40".ToLower(), "COMPONENT")
        lst_ProdSys_ProductNames.Add("faceshieldkit_4b24".ToLower(), "COMPONENT")
        lst_ProdSys_ProductNames.Add("measure protect".ToLower(), "measure protect")
        lst_ProdSys_ProductNames.Add("multichannelremote".ToLower(), "multi-channel remote")
        lst_ProdSys_ProductNames.Add("pelmet".ToLower(), "pelmet")
        lst_ProdSys_ProductNames.Add("pole2m".ToLower(), "2m telescopic pole")
        lst_ProdSys_ProductNames.Add("pole3m".ToLower(), "3m telescopic pole")
        lst_ProdSys_ProductNames.Add("PRB_Fittings_Recess".ToLower(), "COMPONENT")
        lst_ProdSys_ProductNames.Add("PRB_Fittings_Surface".ToLower(), "COMPONENT")
        lst_ProdSys_ProductNames.Add("premiumrollerblind".ToLower(), "premiumrollerblind")
        lst_ProdSys_ProductNames.Add("premiumskylight".ToLower(), "skylightplus")
        lst_ProdSys_ProductNames.Add("RB_Brackets_Black".ToLower(), "COMPONENT")
        lst_ProdSys_ProductNames.Add("RB_Brackets_Grey".ToLower(), "COMPONENT")
        lst_ProdSys_ProductNames.Add("RB_Brackets_White".ToLower(), "COMPONENT")
        lst_ProdSys_ProductNames.Add("RB_DF_Brackets_White".ToLower(), "COMPONENT")
        lst_ProdSys_ProductNames.Add("RB_M_Brackets_Black".ToLower(), "COMPONENT")
        lst_ProdSys_ProductNames.Add("RB_M_Brackets_White".ToLower(), "COMPONENT")
        lst_ProdSys_ProductNames.Add("rollerblind".ToLower(), "rollerblind")
        lst_ProdSys_ProductNames.Add("rollerblinddrillfree".ToLower(), "rollerblind-drillfree")
        lst_ProdSys_ProductNames.Add("rollerblindextend_black".ToLower(), "COMPONENT")
        lst_ProdSys_ProductNames.Add("rollerblindextend_blue".ToLower(), "COMPONENT")
        lst_ProdSys_ProductNames.Add("rollerblindextend_red".ToLower(), "COMPONENT")
        lst_ProdSys_ProductNames.Add("rollerblindextend_green".ToLower(), "COMPONENT")
        lst_ProdSys_ProductNames.Add("rollerblindmotorised".ToLower(), "rollerblind-motorised")
        lst_ProdSys_ProductNames.Add("smarthub".ToLower(), "COMPONENT")
        lst_ProdSys_ProductNames.Add("smartmotorcharger".ToLower(), "COMPONENT")
        lst_ProdSys_ProductNames.Add("smartremote".ToLower(), "COMPONENT")
        lst_ProdSys_ProductNames.Add("smartrepeater".ToLower(), "COMPONENT")
        lst_ProdSys_ProductNames.Add("singlechannelremote".ToLower(), "single-channel remote")
        lst_ProdSys_ProductNames.Add("SKY_Brakes_Cream".ToLower(), "COMPONENT")
        lst_ProdSys_ProductNames.Add("SKY_Brakes_Grey".ToLower(), "COMPONENT")
        lst_ProdSys_ProductNames.Add("SKY_Brakes_White".ToLower(), "COMPONENT")
        lst_ProdSys_ProductNames.Add("solarskylight".ToLower(), "skylightplus")
        lst_ProdSys_ProductNames.Add("SS_Battery_Post_03_19".ToLower(), "COMPONENT")
        lst_ProdSys_ProductNames.Add("SS_Battery_Pre_03_19".ToLower(), "COMPONENT")
        lst_ProdSys_ProductNames.Add("twinrollerblind".ToLower(), "rollerblind")                                  ' Twin roller blinds import as two rollerblind lines (with a couple of changes)
        lst_ProdSys_ProductNames.Add("venetian".ToLower(), "venetian")
        lst_ProdSys_ProductNames.Add("virtualfabric".ToLower(), "rollerblind-additionalfabric")
        lst_ProdSys_ProductNames.Add("zebrarollerblind".ToLower(), "zebrarollerblind")
        lst_ProdSys_ProductNames.Add("ZRB_Fittings_Recess".ToLower(), "COMPONENT")
        lst_ProdSys_ProductNames.Add("ZRB_Fittings_Surface".ToLower(), "COMPONENT")

        ' --- Creating ProdSys manufacturer name / code list ---
        lst_ProdSys_ManufacturerCodes.Add("velux".ToLower(), "VX")
        lst_ProdSys_ManufacturerCodes.Add("roto".ToLower(), "RO")
        lst_ProdSys_ManufacturerCodes.Add("fakro".ToLower(), "FO")
        lst_ProdSys_ManufacturerCodes.Add("rooflite".ToLower(), "RE")
        lst_ProdSys_ManufacturerCodes.Add("okpol".ToLower(), "OL")
        lst_ProdSys_ManufacturerCodes.Add("luctis".ToLower(), "LS")
        lst_ProdSys_ManufacturerCodes.Add("dakstra".ToLower(), "DA")

        ' --- Creating fitting type translation list ---
        lst_ProdSys_FittingTypes.Add("recessedge".ToLower(), "Edge of Window Recess")         ' BlocOut40 fitting type
        lst_ProdSys_FittingTypes.Add("recessinside".ToLower(), "Inside the Window Recess")    ' BlocOut40 fitting type
        lst_ProdSys_FittingTypes.Add("surface3".ToLower(), "Surface three sides")             ' BlocOut40 fitting type
        lst_ProdSys_FittingTypes.Add("surface4".ToLower(), "Surface four sides")              ' BlocOut40 fitting type
        lst_ProdSys_FittingTypes.Add("recess".ToLower(), "Recess")                            ' BlocOut80, Premium Roller, Zebra, Rollers fitting type
        lst_ProdSys_FittingTypes.Add("surface".ToLower(), "Surface")                          ' BlocOut80, Premium Roller, Zebra fitting type
        lst_ProdSys_FittingTypes.Add("topfix".ToLower(), "Topfix")                            ' Premium Roller, Zebra fitting type
        lst_ProdSys_FittingTypes.Add("exact".ToLower(), "Exact")                              ' rollerblind / motorised / twin fitting type

        ' --- Creating Spare Parts product names list ---
        lst_SpareParts.Add("BO40_Brakes_Grey")
        lst_SpareParts.Add("BO40_Brakes_White")
        lst_SpareParts.Add("BO40_Brakes_Cream")
        lst_SpareParts.Add("SKY_Brakes_Grey")
        lst_SpareParts.Add("SKY_Brakes_White")
        lst_SpareParts.Add("SKY_Brakes_Cream")
        lst_SpareParts.Add("BO40_Fittings")
        lst_SpareParts.Add("BO80_Fittings_Recess")
        lst_SpareParts.Add("BO80_Fittings_Surface")
        lst_SpareParts.Add("RB_Brackets_White")
        lst_SpareParts.Add("RB_Brackets_Black")
        lst_SpareParts.Add("RB_Brackets_Grey")
        lst_SpareParts.Add("RB_DF_Brackets_White")
        lst_SpareParts.Add("RB_M_Brackets_White")
        lst_SpareParts.Add("RB_M_Brackets_Black")
        lst_SpareParts.Add("PRB_Fittings_Recess")
        lst_SpareParts.Add("PRB_Fittings_Surface")
        lst_SpareParts.Add("ZRB_Fittings_Recess")
        lst_SpareParts.Add("ZRB_Fittings_Surface")
        'lst_SpareParts.Add("SS_Battery_Pre_03_19")
        'lst_SpareParts.Add("SS_Battery_Post_03_19")

    End Sub


    Sub Log_OrderImport(strWebsite As String, intOrderNo As Integer, strImportResult As String)


        ' --- Defining objects and variables ---
        Dim query As String = "Insert into Log_OrderImports(Order_No, ImportResult) VALUES (@OrderNo, @ImportResult);"
        Dim cmd As SqlCommand
        Dim paramOrderNo As SqlParameter
        Dim paramImportResult As SqlParameter


        ' --- Constructing command ---
        Using conn As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings(strWebsite).ConnectionString)
            cmd = New SqlCommand(query, conn)
            paramOrderNo = New SqlParameter("@OrderNo", intOrderNo)
            paramImportResult = New SqlParameter("@ImportResult", strImportResult)
            cmd.Parameters.Add(paramOrderNo)
            cmd.Parameters.Add(paramImportResult)


            ' --- Executing query ---
            Try
                conn.Open()
                cmd.ExecuteNonQuery()
            Catch ex As Exception
                ' Do nothing
            Finally
                conn.Close()
            End Try

        End Using

    End Sub


    Sub Write_FabricSamples_DiscountToWebsiteDB(strWebsite As String, strDiscountCode As String, intDiscount As Integer, dateExpiry As Date)

        Dim query As String = "INSERT INTO Offers_Percent(OfferName, PercentDiscount, DateStart, DateEnd, RequiredCouponCode, IsActive) " _
            & "VALUES (@OfferName, @PercentDiscount, @DateStart, @DateEnd, @RequiredCouponCode, 1);"
        Dim cmd As SqlCommand
        Dim paramOfferName As SqlParameter
        Dim paramPercentDiscount As SqlParameter
        Dim paramDateStart As SqlParameter
        Dim paramDateEnd As SqlParameter
        Dim paramCouponCode As SqlParameter

        Using conn As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings(strWebsite).ConnectionString)

            cmd = New SqlCommand(query, conn)
            paramOfferName = New SqlParameter("@OfferName", "FabricSampleDiscount")
            paramPercentDiscount = New SqlParameter("@PercentDiscount", intDiscount)
            paramDateStart = New SqlParameter("@DateStart", Now)
            paramDateEnd = New SqlParameter("@DateEnd", dateExpiry)
            paramCouponCode = New SqlParameter("@RequiredCouponCode", strDiscountCode)
            cmd.Parameters.Add(paramOfferName)
            cmd.Parameters.Add(paramPercentDiscount)
            cmd.Parameters.Add(paramDateStart)
            cmd.Parameters.Add(paramDateEnd)
            cmd.Parameters.Add(paramCouponCode)

            Try
                conn.Open()
                cmd.ExecuteNonQuery()
            Catch ex As Exception
                Throw ex
            End Try

        End Using

    End Sub

End Module


