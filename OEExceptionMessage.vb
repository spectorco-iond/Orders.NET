Public Class OEExceptionMessage

    Private ErrorNumber As OEError = OEError.No_Error

    Public Sub New()

    End Sub

    Public Sub New(ByVal piNumber As OEError)

        ErrorNumber = piNumber

    End Sub

    Public Function Message() As String

        Message = Message(ErrorNumber)

    End Function

    Public Function Message(ByVal pinumber As OEError) As String

        Message = ""

        Select Case pinumber
            Case OEError.Account_Not_On_File
                Message = "Account Not On File."
            Case OEError.Account_Not_Valid_For_OE
                Message = "Account Not Valid For OE."
            Case OEError.Auth_Qty_GT_Applied_Qty
                Message = "Authorized Quantity is greater than the applied quantity."
            Case OEError.Cannot_Use_Term_If_CC_Exist
                Message = "Cannot use this terms code if CC records exist"
            Case OEError.Cannot_Use_Term_If_OECash_Exist
                Message = "Cannot use this terms code if OE cash records exist"
            Case OEError.Charge_Not_Found
                Message = "Charge not found."
            Case OEError.Contact_First_Name_Needed
                Message = "Please enter a first name."
            Case OEError.Contact_Incorrect_Email
                Message = "Incorrect e-mail syntax"
            Case OEError.Contact_Incorrect_Fax_Syntax
                Message = "Incorrect fax number syntax. Make sure the area code is included"
            Case OEError.Contact_Incorrect_Phone_Syntax
                Message = "Incorrect telephone number syntax. Make sure the area code is included"
            Case OEError.Contact_Last_Name_Needed
                Message = "Please enter a last name."
            Case OEError.Contact_Method_Needed
                Message = "Please enter at least one contact method"
            Case OEError.Contact_Name_Pred_Needed
                Message = "Please select a name pred."
            Case OEError.Contact_Preferred_Language_Needed
                Message = "Please select a language."
            Case OEError.Copy_Not_Allowed_From_History_Load
                Message = "Copy not allowed, record entered through sales history load"
            Case OEError.Current_Period_Outsd_Valid_Per
                Message = "Current period outside valid periods"
            Case OEError.Cust_Curr_Cd_Outsd_Valid_Per
                Message = "Customer currency code does not match currency code"
            Case OEError.Cust_Do_Not_BO
                Message = "Customer does not accept backorders.  Is this OK?"
            Case OEError.Cust_Do_Not_Partial_Ship
                Message = "Customer does not accept partial shipments.  Is this OK?"
            Case OEError.Cust_Do_Not_Substitute
                Message = "Customer doesn't allow substitute items.  OK to use anyway?"
            Case OEError.Cust_Balance_Over_Terms
                Message = "Customer has balance over terms.  Do you want to place order on hold?"
            Case OEError.Cust_Is_Inactive
                Message = "Customer is inactive."
            Case OEError.Cust_Orders_Over_Terms
                Message = "Customer has orders over terms. Is this OK?"
            Case OEError.Cust_On_Credit_Hold
                Message = "Customer in on credit hold.  Is this This OK?"
            Case OEError.Cust_On_Credit_Hold_OE_Not_Allowed
                Message = "Customer is on credit hold.  Order entry not allowed."
            Case OEError.Cust_Item_Not_Found
                Message = "Customer item record not found"
            Case OEError.Cust_Not_Found
                Message = "Customer not found"
            Case OEError.Cust_Locked
                Message = "Customer locked at another station."
            Case OEError.Cust_Over_Limit_And_Terms
                Message = "Customer over credit limit and terms.  Is this OK?"
            Case OEError.Cust_Over_Limit_And_Terms_Place_On_Hold
                Message = "Customer over credit limit and terms.  Do you want to place order on hold?"
            Case OEError.Cust_Over_Credit_Limit
                Message = "Customer over credit limit.  Do you want to put order on hold?"
            Case OEError.Cust_PO_Exists
                Message = "Customer P/O No. already on file.  Is this OK?"
            Case OEError.Cust_PO_Exists_In_History
                Message = "Customer PO No. already in history file.  Is this OK?"
            Case OEError.Date_Outside_Valid_Trx_Dates
                Message = "Date is outside of valid trx dates. "
            Case OEError.Delete_Error
                Message = "Delete error"
            Case OEError.Delete_Successful
                Message = "Delete Successful"
            Case OEError.Disc_Pct_GT_100_Pct
                Message = "Discount percent cannot be greater than 100%"
            Case OEError.Disc_Pct_LT_0_Pct
                Message = "Discount percent cannot be lower than 0%"
            Case OEError.Duplic_SalesP_Not_Allowed
                Message = "Duplicate salesperson not allowed"
            Case OEError.Duplic_Tax_Code_Not_Allowed
                Message = "Duplicate tax code not allowed."
            Case OEError.Entry_Too_Long
                Message = "Entry too long."
            Case OEError.Exit_Function
                Message = "Exit Function."
            Case OEError.Exit_Property
                Message = "Exit Property."
            Case OEError.Exit_Sub
                Message = "Exit Sub."
            Case OEError.Insuf_Qty_Avail_In_Stock
                Message = "Insufficient quantity available in stock."
            Case OEError.Invalid_Date_Format
                Message = "Invalid Date Format."
            Case OEError.Invalid_Number_Format
                Message = "Invalid Number Format."
            Case OEError.Item_Not_Found
                Message = "Inventory item not on file."
            Case OEError.Item_Location_Not_Found
                Message = "Inventory location not on file."
            Case OEError.Item_Already_On_Order
                Message = "Item already on order.  Is this OK?"
            Case OEError.Item_Inactive
                Message = "Item inactive"
            Case OEError.Item_NA_At_Location
                Message = "Item not available at this location"
            Case OEError.Item_Not_BO_At_Location
                Message = "Item not backordered at this location"
            Case OEError.Item_No_Required
                Message = "Item number required"
            Case OEError.Item_Obsolete
                Message = "Item obsolete"
            Case OEError.Item_Locked
                Message = "Item record is locked at another station"
            Case OEError.Location_Must_Be_Numeric
                Message = "Location Must Be Numeric"
            Case OEError.Low_Stock_For_All_Star_Customer
                Message = "All Star Traveler Alert! " & vbCrLf & "Low Stock - Please verify inventory."
            Case OEError.Negative_Qty_Not_Allowed
                Message = "Negative quantity not allowed. Must enter credit memo."
            Case OEError.No_Accounting_Period_Record_Found
                Message = "No accounting period record found."
            Case OEError.No_Item_Notes_Found
                Message = "No Item Notes Found."
            Case OEError.No_Valid_Line_Items_Found
                Message = "No Valid Line Items Found."
            Case OEError.Non_Inv_Items_Not_Allowed
                Message = "Non-inventory items not allowed."
            Case OEError.Nothing_To_Delete
                Message = "Nothing To Delete."
            Case OEError.Inactive_Item_Not_Allowed
                Message = "Inactive Items Not Allowed."
            Case OEError.Order_Exported_To_Macola
                Message = "Order Exported To Macola. Any modification will not be saved."
            Case OEError.Order_No_Not_Found
                Message = "Order no not found."
            Case OEError.Order_No_Found_In_OEI
                Message = "The order no was found in OEI."
            Case OEError.Order_No_Found_In_OEI_Exported_Not_Processed
                Message = "The order no was found in OEI and is in process of being exported to Macola. Entry not allowed."
            Case OEError.Order_No_Found_In_OEI_Exported_And_Processed
                Message = "The order no was found in OEI and is exported to Macola. Entry not allowed."
            Case OEError.Order_No_Found_In_History_Files_Entry_Not_Allowed
                Message = "Order no. found in history files, entry not allowed."
            Case OEError.Order_No_From_Macola_Entry_Not_Allowed
                Message = "Order no. from Macola, entry not allowed."
            Case OEError.Order_No_In_Use
                Message = "Order no. in use, entry not allowed."
            Case OEError.Order_No_Is_New
                Message = "Order no. is new"
            Case OEError.Order_Non_Taxable_Must_Be_Zero
                Message = "Order is non-taxable - percent must be zero."
            Case OEError.Order_Not_Exported_No_Item_Lines
                Message = "Order not exported, no item lines."
            Case OEError.Order_On_Hold
                Message = "Order on hold."
            Case OEError.Order_On_Hold_Access_Denied
                Message = "Order on hold. Access denied."
            Case OEError.Order_Selected_For_Bill_Access_Denied
                Message = "Order selected for billing. Access denied."
            Case OEError.Order_Type_Different_From_Selected_Type
                Message = "Order type not equal to selected order type."
            Case OEError.Password_Required_To_Override_Price
                Message = "Password required to override price."
            Case OEError.Price_Code_Deleted
                Message = "Price code deleted"
            Case OEError.Price_Code_Not_Valid
                Message = "Price code not valid"
            Case OEError.Promise_Date_Before_System_Date
                Message = "Promise date is before system date."
            Case OEError.Qty_Ordered_Cannot_Be_Zero
                Message = "Qty ordered cannot be zero."
            Case OEError.Qty_Ordered_GT_Max_Qty_That_Can_Be_Produced
                Message = "Qty ordered is greater than the maximum qty that can be produced"
            Case OEError.Qty_To_Ship_GT_Qty_Available
                Message = "Qty to ship cannot be greater than qty available."
            Case OEError.Qty_To_Ship_GT_Qty_Ordered
                Message = "Qty to Ship greater than Qty Ordered"
            Case OEError.Qty_GT_Qty_On_Hand
                Message = "Quantity greater than quantity on hand"
            Case OEError.Qty_GT_Qty_Ordered
                Message = "Quantity greater than quantity ordered"
            Case OEError.Req_Date_Before_System_Date
                Message = "Request date is before system date."
            Case OEError.Req_Ship_Date_GT_Promise_Date
                Message = "Required ship date is greater than promise date."
            Case OEError.Salesperson_Not_Found
                Message = "Salesperson not on file."
            Case OEError.Ship_Via_Cd_Not_Found
                Message = "Ship via code not on file."
            Case OEError.Susbt_Items_Not_Found
                Message = "Substitute item not found"
            Case OEError.Subst_Items_Found_Do_You_Wish_To_Substitute
                Message = "Substitute items found.  Do you wish to substitute?"
            Case OEError.Tax_Cd_Cannot_Be_Blank
                Message = "Tax code cannot be blank."
            Case OEError.Tax_Code_Not_Found
                Message = "Tax code not on file."
            Case OEError.Tax_Pct_Zero_Change_Not_Allowed
                Message = "Tax percent 1 is zero and change is not allowed."
            Case OEError.Term_Cd_Not_Found
                Message = "Terms code not on file."
            Case OEError.Insuf_Qty_On_Hand_At_Location
                Message = "There is insufficient quantity on-hand at this location."
            Case OEError.No_Minimum_Sales_Volume_For_Eligibility
                Message = "There is no minimum sales volume for item eligibility."
            Case OEError.This_Is_Non_Inventory_Item
                Message = "This is a non inventory item.  Is this OK?"
            Case OEError.Too_Many_Setups_Entered
                Message = "Too Many Setups Entered."
            Case OEError.Total_Of_SalesP_GT_100_Pct
                Message = "Total of all salesperson percents cannot be greater than 100 percent."
            Case OEError.Total_Qty_Cannot_Be_Zero
                Message = "Total quantity cannot be zero."
            Case OEError.Total_Of_SalesP_Diff_100_Pct
                Message = "Total salespersons commission doesn't equal 100%."
            Case OEError.Unit_Price_Below_Unit_Cost
                Message = "Unit price below unit cost."
            Case OEError.Unit_Price_Below_Unit_Cost_Is_This_OK
                Message = "Unit price below unit cost.  Is this OK?"
            Case OEError.Unit_Price_Change_Not_Allowed
                Message = "Unit price change not allowed."
            Case OEError.Unit_Price_Has_Been_Changed
                Message = "Unit price has been changed."
            Case OEError.Unit_Price_Has_Been_Changed_Is_This_OK
                Message = "Unit price has been changed.  Is this OK?"
            Case OEError.User_Does_Not_Exist_In_Macola
                Message = "User does not exist in Macola."
            Case OEError.User_Is_Blocked_From_Macola
                Message = "User is blocked from Macola."
            Case OEError.Year_Cannot_Be_GT_Follow_Year
                Message = "Year cannot be greater than following year"
            Case OEError.Year_Cannot_Be_LT_Previous_Year
                Message = "Year cannot be less than previous year"
            Case OEError.Specific_Instructions
                Message = "Specific Customer Instructions"
            Case OEError.OrderTypeInvoiceIsActive
                Message = "You cannot pull exact repeat data from an invoice type order that hasn't been closed yet"
            Case Else
                Message = "Error occured in OEExceptionMessage"
        End Select

    End Function

End Class
