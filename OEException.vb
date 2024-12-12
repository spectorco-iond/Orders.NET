Imports System.Runtime.Serialization

Public Class OEException
    Inherits System.ApplicationException

    Private m_Number As OEError
    Private m_Cancel As Boolean = False
    Private m_ShowMessage As Boolean = True

    'Private m_NoFinally As Boolean = False

#Region "Contructors ######################################################"

    Public Sub New(ByVal piNumber As OEError, ByVal pstrMessage As String)

        MyBase.New(New OEExceptionMessage(piNumber).Message & " " & pstrMessage)
        m_Number = piNumber

    End Sub

    Public Sub New(ByVal piNumber As OEError)

        MyBase.New(New OEExceptionMessage(piNumber).Message)
        m_Number = piNumber

    End Sub

    Public Sub New(ByVal piNumber As OEError, ByVal pbCancel As Boolean)

        MyBase.New(New OEExceptionMessage(piNumber).Message)
        m_Number = piNumber
        m_Cancel = pbCancel

    End Sub

    Public Sub New(ByVal piNumber As OEError, ByVal pbCancel As Boolean, ByVal pbShowMessage As Boolean)

        MyBase.New(New OEExceptionMessage(piNumber).Message)
        m_Number = piNumber
        m_Cancel = pbCancel
        m_ShowMessage = pbShowMessage

    End Sub

    Public Sub New(ByVal piNumber As OEError, ByVal pstrMessage As String, ByVal pbCancel As Boolean, ByVal pbShowMessage As Boolean)

        MyBase.New(New OEExceptionMessage(piNumber).Message & " " & pstrMessage)
        m_Number = piNumber
        m_Cancel = pbCancel
        m_ShowMessage = pbShowMessage

    End Sub

    Public Sub New(ByVal pstrMessage As String)

        MyBase.New(pstrMessage)
        m_Number = OEError.No_Error

    End Sub

    Public Sub New(ByVal info As SerializationInfo, ByVal context As StreamingContext)

        MyBase.New(info, context)
        m_Number = OEError.No_Error

    End Sub

    Public Sub New(ByVal pstrMessage As String, ByVal innerException As System.Exception)

        MyBase.New(pstrMessage, innerException)
        m_Number = OEError.No_Error
        If TypeOf innerException Is OEException Then
            Dim myException As OEException = innerException
            m_Cancel = myException.Cancel
            m_ShowMessage = myException.ShowMessage
        End If

    End Sub

#End Region

#Region "Public properties ################################################"

    ' Cancel will be used in the try catch to set event cancellation
    Public Property Cancel() As Boolean
        Get
            Cancel = m_Cancel
        End Get
        Set(ByVal value As Boolean)
            m_Cancel = value
        End Set
    End Property

    ' Returns error number
    Public Property Number() As String
        Get
            Number = m_Number
        End Get
        Set(ByVal value As String)
            m_Number = Number
        End Set
    End Property

    ' ShowMessage will be used in the try catch to set event cancellation 
    ' without showing the message when set to false. Default will show message.
    Public Property ShowMessage() As Boolean
        Get
            ShowMessage = m_ShowMessage
        End Get
        Set(ByVal value As Boolean)
            m_ShowMessage = value
        End Set
    End Property

#End Region

    Public Function Header() As String

        Header = "Error"

        Select Case m_Number
            'Case OEError.Account_Not_On_File
            'Case OEError.Account_Not_Valid_For_OE
            'Case OEError.Auth_Qty_GT_Applied_Qty
            'Case OEError.Cannot_Use_Term_If_CC_Exist
            'Case OEError.Cannot_Use_Term_If_OECash_Exist
            'Case OEError.Contact_First_Name_Needed
            'Case OEError.Contact_Incorrect_Email
            'Case OEError.Contact_Incorrect_Fax_Syntax
            'Case OEError.Contact_Incorrect_Phone_Syntax
            'Case OEError.Contact_Last_Name_Needed
            'Case OEError.Contact_Method_Needed
            'Case OEError.Contact_Name_Pred_Needed
            'Case OEError.Contact_Preferred_Language_Needed
            'Case OEError.Copy_Not_Allowed_From_History_Load
            'Case OEError.Current_Period_Outsd_Valid_Per
            'Case OEError.Cust_Curr_Cd_Outsd_Valid_Per
            'Case OEError.Cust_Do_Not_BO
            'Case OEError.Cust_Do_Not_Partial_Ship
            'Case OEError.Cust_Do_Not_Substitute
            'Case OEError.Cust_Balance_Over_Terms
            'Case OEError.Cust_Orders_Over_Terms
            'Case OEError.Cust_On_Credit_Hold
            'Case OEError.Cust_On_Credit_Hold_OE_Not_Allowed
            'Case OEError.Cust_Item_Not_Found
            'Case OEError.Cust_Not_Found
            'Case OEError.Cust_Locked
            'Case OEError.Cust_Over_Limit_And_Terms
            'Case OEError.Cust_Over_Limit_And_Terms_Place_On_Hold
            'Case OEError.Cust_Over_Credit_Limit
            'Case OEError.Cust_PO_Exists
            'Case OEError.Cust_PO_Exists_In_History
            'Case OEError.Date_Outside_Valid_Trx_Dates
            'Case OEError.Delete_Error
            'Case OEError.Delete_Successful
            'Case OEError.Disc_Pct_GT_100_Pct
            'Case OEError.Disc_Pct_LT_0_Pct
            'Case OEError.Duplic_SalesP_Not_Allowed
            'Case OEError.Duplic_Tax_Code_Not_Allowed
            'Case OEError.Entry_Too_Long
            'Case OEError.Exit_Function
            'Case OEError.Exit_Property
            'Case OEError.Exit_Sub
            'Case OEError.Insuf_Qty_Avail_In_Stock
            'Case OEError.Invalid_Date_Format
            'Case OEError.Invalid_Number_Format
            'Case OEError.Item_Not_Found
            'Case OEError.Item_Location_Not_Found
            'Case OEError.Item_Already_On_Order
            'Case OEError.Item_Inactive
            'Case OEError.Item_NA_At_Location
            'Case OEError.Item_Not_BO_At_Location
            'Case OEError.Item_No_Required
            'Case OEError.Item_Obsolete
            'Case OEError.Item_Locked
            Case OEError.Low_Stock_For_All_Star_Customer
                Header = "Low Stock For 6/7-Star Customers"

                'Case OEError.Negative_Qty_Not_Allowed
                'Case OEError.No_Accounting_Period_Record_Found
                'Case OEError.No_Item_Notes_Found
                'Case OEError.No_Valid_Line_Items_Found
                'Case OEError.Non_Inv_Items_Not_Allowed
                'Case OEError.Nothing_To_Delete
                'Case OEError.Inactive_Item_Not_Allowed
                'Case OEError.Order_No_Not_Found
                'Case OEError.Order_No_Found_In_OEI
                'Case OEError.Order_No_Found_In_OEI_Was_Exported
                'Case OEError.Order_No_Found_In_History_Files_Entry_Not_Allowed
                'Case OEError.Order_No_From_Macola_Entry_Not_Allowed
                'Case OEError.Order_No_Is_New
                'Case OEError.Order_Non_Taxable_Must_Be_Zero
                'Case OEError.Order_On_Hold
                'Case OEError.Order_On_Hold_Access_Denied
                'Case OEError.Order_Selected_For_Bill_Access_Denied
                'Case OEError.Order_Type_Different_From_Selected_Type
                'Case OEError.Password_Required_To_Override_Price
                'Case OEError.Price_Code_Deleted
                'Case OEError.Price_Code_Not_Valid
                'Case OEError.Promise_Date_Before_System_Date
                'Case OEError.Qty_Ordered_Cannot_Be_Zero
                'Case OEError.Qty_Ordered_GT_Max_Qty_That_Can_Be_Produced
                'Case OEError.Qty_To_Ship_GT_Qty_Available
                'Case OEError.Qty_To_Ship_GT_Qty_Ordered
                'Case OEError.Qty_GT_Qty_On_Hand
                'Case OEError.Qty_GT_Qty_Ordered
                'Case OEError.Req_Date_Before_System_Date
                'Case OEError.Req_Ship_Date_GT_Promise_Date
                'Case OEError.Salesperson_Not_Found
                'Case OEError.Ship_Via_Cd_Not_Found
                'Case OEError.Susbt_Items_Not_Found
                'Case OEError.Subst_Items_Found_Do_You_Wish_To_Substitute
                'Case OEError.Tax_Cd_Cannot_Be_Blank
                'Case OEError.Tax_Code_Not_Found
                'Case OEError.Tax_Pct_Zero_Change_Not_Allowed
                'Case OEError.Term_Cd_Not_Found
                'Case OEError.Insuf_Qty_On_Hand_At_Location
                'Case OEError.No_Minimum_Sales_Volume_For_Eligibility
                'Case OEError.This_Is_Non_Inventory_Item
                'case OEError.Too_Many_Setups_Entered
                'Case OEError.Total_Of_SalesP_GT_100_Pct
                'Case OEError.Total_Qty_Cannot_Be_Zero
                'Case OEError.Total_Of_SalesP_Diff_100_Pct
                'Case OEError.Unit_Price_Below_Unit_Cost
                'Case OEError.Unit_Price_Below_Unit_Cost_Is_This_OK
                'Case OEError.Unit_Price_Change_Not_Allowed
                'Case OEError.Unit_Price_Has_Been_Changed
                'Case OEError.Unit_Price_Has_Been_Changed_Is_This_OK
                'Case OEError.User_Does_Not_Exist_In_Macola
                'Case OEError.User_Is_Blocked_From_Macola
                'Case OEError.Year_Cannot_Be_GT_Follow_Year
                'Case OEError.Year_Cannot_Be_LT_Previous_Year
                'Case Else

        End Select

    End Function

End Class

Public Enum OEError

    No_Error = 0
    Account_Not_On_File
    Account_Not_Valid_For_OE
    Auth_Qty_GT_Applied_Qty
    Cannot_Use_Term_If_CC_Exist
    Cannot_Use_Term_If_OECash_Exist
    Charge_Not_Found
    Contact_First_Name_Needed
    Contact_Incorrect_Email
    Contact_Incorrect_Fax_Syntax
    Contact_Incorrect_Phone_Syntax
    Contact_Last_Name_Needed
    Contact_Method_Needed
    Contact_Name_Pred_Needed
    Contact_Preferred_Language_Needed
    Copy_Not_Allowed_From_History_Load
    Current_Period_Outsd_Valid_Per
    Cust_Curr_Cd_Outsd_Valid_Per
    Cust_Do_Not_BO
    Cust_Do_Not_Partial_Ship
    Cust_Do_Not_Substitute
    Cust_Balance_Over_Terms
    Cust_Orders_Over_Terms
    Cust_On_Credit_Hold
    Cust_On_Credit_Hold_OE_Not_Allowed
    Cust_Is_Inactive
    Cust_Item_Not_Found
    Cust_Not_Found
    Cust_Locked
    Cust_Over_Limit_And_Terms
    Cust_Over_Limit_And_Terms_Place_On_Hold
    Cust_Over_Credit_Limit
    Cust_PO_Exists
    Cust_PO_Exists_In_History
    Date_Outside_Valid_Trx_Dates
    Delete_Error
    Delete_Successful
    Disc_Pct_GT_100_Pct
    Disc_Pct_LT_0_Pct
    Duplic_SalesP_Not_Allowed
    Duplic_Tax_Code_Not_Allowed
    Entry_Too_Long
    Exit_Function
    Exit_Property
    Exit_Sub
    Insuf_Qty_Avail_In_Stock
    Invalid_Date_Format
    Invalid_Number_Format
    Item_Not_Found
    Item_Location_Not_Found
    Item_Already_On_Order
    Item_Inactive
    Item_NA_At_Location
    Item_Not_BO_At_Location
    Item_No_Required
    Item_Obsolete
    Item_Locked
    Location_Must_Be_Numeric
    Low_Stock_For_All_Star_Customer
    Negative_Qty_Not_Allowed
    No_Accounting_Period_Record_Found
    No_Item_Notes_Found
    No_Valid_Line_Items_Found
    Non_Inv_Items_Not_Allowed
    Nothing_To_Delete
    Inactive_Item_Not_Allowed
    Order_Exported_To_Macola
    Order_No_Found_In_History_Files_Entry_Not_Allowed
    Order_No_Found_In_OEI
    Order_No_Found_In_OEI_Exported_Not_Processed
    Order_No_Found_In_OEI_Exported_And_Processed
    Order_No_From_Macola_Entry_Not_Allowed
    Order_No_Is_New
    Order_No_In_Use
    Order_No_Not_Found
    Order_Non_Taxable_Must_Be_Zero
    Order_Not_Exported_No_Item_Lines
    Order_On_Hold
    Order_On_Hold_Access_Denied
    Order_Selected_For_Bill_Access_Denied
    Order_Type_Different_From_Selected_Type
    Password_Required_To_Override_Price
    Price_Code_Deleted
    Price_Code_Not_Valid
    Promise_Date_Before_System_Date
    Qty_Ordered_Cannot_Be_Zero
    Qty_Ordered_GT_Max_Qty_That_Can_Be_Produced
    Qty_To_Ship_GT_Qty_Available
    Qty_To_Ship_GT_Qty_Ordered
    Qty_GT_Qty_On_Hand
    Qty_GT_Qty_Ordered
    Req_Date_Before_System_Date
    Req_Ship_Date_GT_Promise_Date
    Salesperson_Not_Found
    Ship_Via_Cd_Not_Found
    Susbt_Items_Not_Found
    Subst_Items_Found_Do_You_Wish_To_Substitute
    Tax_Cd_Cannot_Be_Blank
    Tax_Code_Not_Found
    Tax_Pct_Zero_Change_Not_Allowed
    Term_Cd_Not_Found
    Insuf_Qty_On_Hand_At_Location
    No_Minimum_Sales_Volume_For_Eligibility
    This_Is_Non_Inventory_Item
    Too_Many_Setups_Entered
    Total_Of_SalesP_GT_100_Pct
    Total_Qty_Cannot_Be_Zero
    Total_Of_SalesP_Diff_100_Pct
    Unit_Price_Below_Unit_Cost
    Unit_Price_Below_Unit_Cost_Is_This_OK
    Unit_Price_Change_Not_Allowed
    Unit_Price_Has_Been_Changed
    Unit_Price_Has_Been_Changed_Is_This_OK
    User_Does_Not_Exist_In_Macola
    User_Is_Blocked_From_Macola
    Year_Cannot_Be_GT_Follow_Year
    Year_Cannot_Be_LT_Previous_Year
    Specific_Instructions
    OrderTypeInvoiceIsActive
End Enum