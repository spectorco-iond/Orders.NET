Option Strict Off
Option Explicit On

Public Class cOrdHead

    ' FIELDS THAT HAVE NO DATA ON THE RIGHT ARE NOT TO BE INCLUDED IN THE EXPORT FILE
    ' BUT ARE INCLUDED HERE FOR RECUPERATING PURPOSES.

#Region "Class variables ##################################################"

    'Champs obligatoires de OEORDHDR_SQL
    '[ord_type] [char] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
    '[ord_no] [char] (8) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
    '[status] [char] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
    '[oe_po_no] [char] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
    '[cus_no] [char] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
    '[slspsn_no] [int] NOT NULL ,
    '[inv_no] [char] (8) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
    '[selection_cd] [char] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
    '[hold_fg] [char] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
    '[oe_cash_no] [char] (8) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
    '[curr_cd] [char] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
    '[rma_no] [char] (8) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
    '[ID] [numeric](9, 0) IDENTITY (1, 1) NOT NULL ,

    Public Header_Table As String = ""
    Public Lines_Table As String = ""

    Public Record_Type As String = "H" 'Char (1)       'H' for header ' NOT TO BE RELOADED

    Public Ord_Type As String 'Char (1)
    Public Ord_No As String 'Char (8)       Order No    Order No from Website for tracking
    Private m_OEI_Ord_No As String 'Char (8)       Order No    Order No from Website for tracking
    Public Status As String 'Char(1) NOT NULL
    Public Entered_Dt As Date 'DateTime NULL
    Public Ord_Dt As Date 'Datetime       Order date
    Public Apply_To_No As String 'Char(8) NULL
    Public Oe_Po_No As String 'Char(25) NOT NULL Customer PO No
    Public Cus_No As String 'Char(20) NOT NULL Customer No     Always '999999999999'
    Public Cus_Name As String ' used for information purposes
    Public Bal_Meth As String 'Char(1) NULL
    Public Bill_To_Name As String 'Char (40)      Customer Name
    Public Bill_To_Addr_1 As String 'Char (40)      1st address line
    Public Bill_To_Addr_2 As String 'Char (40)      2nd address line
    Public Bill_To_Addr_3 As String 'Char (40)      3rd address line
    Public Bill_To_Addr_4 As String 'Char(44) NULL
    Public Bill_To_Country As String 'Char(3) NULL
    Private m_Cus_Alt_Adr_Cd As String 'Char(15) NULL
    Public Ship_To_Name As String 'Char(40) NULL  Ship To name
    Public Ship_To_Addr_1 As String 'Char(40) NULL  1st Ship To address line
    Public Ship_To_Addr_2 As String 'Char(40) NULL  2nd Ship To address line
    Public Ship_To_Addr_3 As String 'Char(40) NULL
    Public Ship_To_Addr_4 As String 'Char(44) NULL
    Public Ship_To_Country As String 'Char(3) NULL   Ship To Country
    Public Shipping_Dt As Date 'DateTime NULL  Requested Ship Date
    Public Ship_Via_Cd As String 'Char(3) NULL   Ship Via Code
    Public Ar_Terms_Cd As String 'Char(2) NULL
    Public Frt_Pay_Cd As String 'Char(1) NULL
    Public Ship_Instruction_1 As String 'Char(40) NULL Shipping Instructions Line 1
    Public Ship_Instruction_2 As String 'Char(40) NULL Shipping Instructions Line 2
    Public Slspsn_No As Integer 'Int NOT NULL
    Public Slspsn_Pct_Comm As Double 'Numeric NULL
    Public Slspsn_Comm_Amt As Double 'Numeric NULL
    Public Slspsn_No_2 As Integer 'Int NULL
    Public Slspsn_Pct_Comm_2 As Double 'Numeric NULL
    Public Slspsn_Comm_Amt_2 As Double 'Numeric NULL
    Public Slspsn_No_3 As Integer 'Int NULL
    Public Slspsn_Pct_Comm_3 As Double 'Numeric NULL
    Public Slspsn_Comm_Amt_3 As Double 'Numeric NULL
    Public Tax_Cd As String 'Char(3) NULL   Tax Code    Applies to this order only
    Public Tax_Pct As Double 'Numeric NULL
    Public Tax_Cd_2 As String 'Char(3) NULL
    Public Tax_Pct_2 As Double 'Numeric NULL
    Public Tax_Cd_3 As String 'Char(3) NULL
    Public Tax_Pct_3 As Double 'Numeric NULL
    Public Discount_Pct As Double 'Numeric NULL
    Public Job_No As String 'Char(20) NULL
    Public Mfg_Loc As String 'Char(3) NULL
    Public Profit_Center As String 'Char(8) NULL
    Public Dept As String 'Char(8) NULL
    Public Ar_Reference As String 'Char(45) NULL
    Public Tot_Sls_Amt As Double 'Numeric NULL
    Public Tot_Sls_Disc As Double 'Numeric NULL
    Public Tot_Tax_Amt As Double 'Numeric NULL
    Public Tot_Cost As Double 'Numeric NULL
    Public Tot_Weight As Double 'Numeric NULL
    Public Misc_Amt As Double 'Numeric(16,2)  NULL   Miscellaneous Charge    Handling charge
    Public Misc_Mn_No As String 'Char(9) NULL
    Public Misc_Sb_No As String 'Char(8) NULL
    Public Misc_Dp_No As String 'Char(8) NULL
    Public Frt_Amt As Double 'Numeric(16,2)  NULL   Freight Charge
    Public Frt_Mn_No As String 'Char(9) NULL
    Public Frt_Sb_No As String 'Char(8) NULL
    Public Frt_Dp_No As String 'Char(8) NULL
    Public Sls_Tax_Amt_1 As Double 'Numeric NULL
    Public Sls_Tax_Amt_2 As Double 'Numeric NULL
    Public Sls_Tax_Amt_3 As Double 'Numeric NULL
    Public Comm_Pct As Double 'Numeric NULL
    Public Comm_Amt As Double 'Numeric NULL
    Public Cmt_1 As String 'Char(35) NULL
    Public Cmt_2 As String 'Char(35) NULL
    Public Cmt_3 As String 'Char(35) NULL
    Public Payment_Amt As Double 'Numeric NULL
    Public Payment_Disc_Amt As Double 'Numeric NULL
    Public Chk_No As String 'Char(8) NULL
    Public Chk_Dt As Date 'DateTime NULL
    Public Cash_Mn_No As String 'Char(9) NULL
    Public Cash_Sb_No As String 'Char(8) NULL
    Public Cash_Dp_No As String 'Char(8) NULL
    Public Ord_Dt_Picked As Date 'DateTime NULL
    Public Ord_Dt_Billed As Date 'DateTime NULL
    Public Inv_No As String 'Char(8) NOT NULL
    Public Inv_Dt As Date 'DateTime NULL
    Public Selection_Cd As String 'Char(1) NOT NULL
    Public Posted_Dt As Date 'DateTime NULL
    Public Part_Posted_Fg As String 'Char(1) NULL
    Public Ship_To_Freefrm_Fg As String 'Char(1) NULL
    Public Bill_To_Freefrm_Fg As String 'Char(1) NULL
    Public Copy_To_Bm_Fg As String 'Char(1) NULL
    Public Edi_Fg As String 'Char(1) NULL
    Public Closed_Fg As String 'Char(1) NULL
    Public Accum_Misc_Amt As Double 'Numeric NULL
    Public Accum_Frt_Amt As Double 'Numeric NULL
    Public Accum_Tot_Tax_Amt As Double 'Numeric NULL
    Public Accum_Sls_Tax_Amt As Double 'Numeric NULL
    Public Accum_Tot_Sls_Amt As Double 'Numeric NULL
    Public Hold_Fg As String = String.Empty  'Char(1) NOT NULL
    Public Prepayment_Fg As String 'Char(1) NULL
    Public Lost_Sale_Cd As String 'Char(3) NULL
    Public Orig_Ord_Type As String 'Char(1) NULL
    Public Orig_Ord_Dt As Date 'DateTime NULL
    Public Orig_Ord_No As String 'Char(8) NULL
    Public Award_Probability As Short 'TinyInt NULL
    Public Oe_Cash_No As String 'Char(8) NOT NULL
    Public Exch_Rt_Fg As String 'Char(1) NULL
    Public m_Curr_Cd As String 'Char(3) NOT NULL
    Public Orig_Trx_Rt As Double 'Numeric NULL
    Public Curr_Trx_Rt As Double 'Numeric NULL
    Public Tax_Sched As String 'Char(5) NULL
    Public User_Def_Fld_1 As String = String.Empty 'Char(30) NULL
    Public User_Def_Fld_2 As String = String.Empty 'Char(30) NULL
    Public User_Def_Fld_3 As String = String.Empty 'Char(30) NULL
    Public User_Def_Fld_4 As String = String.Empty 'Char(30) NULL
    Public User_Def_Fld_5 As String = String.Empty 'Char(30) NULL
    Public Deter_Rate_By As String 'Char(1) NULL
    Public Form_No As Short 'TinyInt NULL
    Public Tax_Fg As String 'Char(1) NULL  Taxable status  'Y'es or 'N'o - Applies to this order only
    Public Sls_Mn_No As String 'Char(9) NULL
    Public Sls_Sb_No As String 'Char(8) NULL
    Public Sls_Dp_No As String 'Char(8) NULL
    Public Ord_Dt_Shipped As Date 'DateTime NULL
    Public Tot_Dollars As Double 'Numeric NULL
    Public Mult_Loc_Fg As String 'Char(1) NULL
    Public Tot_Tax_Cost As Double 'Numeric NULL
    Public Hist_Load_Record As String 'Char(1) NULL
    Public Pre_Select_Status As String 'Char(1) NULL
    Public Packing_No As Integer 'Int NULL
    Public Deliv_Ar_Terms_Cd As String 'Char(2) NULL
    Public Inv_Batch_Id As String 'Char(8) NULL
    Public Bill_To_No As String 'Char(20) NULL
    Public Rma_No As String 'Char(8) NOT NULL
    Public Prog_Term_No As Integer 'Int NULL
    Public Extra_1 As String 'Char(1) NULL
    Public Extra_2 As String 'Char(1) NULL
    Public Extra_3 As String 'Char(1) NULL
    Public Extra_4 As String 'Char(1) NULL
    Public Extra_5 As String 'Char(1) NULL
    Public Extra_6 As String 'Char(8) NULL
    Public Extra_7 As String 'Char(8) NULL
    Public Extra_8 As String 'Char(12) NULL
    Public Extra_9 As String 'Char(12) NULL
    Public Extra_10 As Double 'Numeric NULL
    Public Extra_11 As Double 'Numeric NULL
    Public Extra_12 As Double 'Numeric NULL
    Public Extra_13 As Double 'Numeric NULL
    Public Extra_14 As Integer 'Int NULL
    Public Extra_15 As Integer 'Int NULL
    Public Edi_Doc_Seq As Short 'SmallInt NULL
    Public Contact_1 As String 'Char(100) NULL
    Public Phone_Number As String 'Char(25) NULL
    Public Fax_Number As String 'Char(25) NULL
    Public Email_Address As String 'Char(128) NULL
    Public Use_Email As String 'Char(10) NULL
    Public Ship_To_City As String 'Char(100) NULL     Ship To City
    Public Ship_To_State As String 'Char(3) NULL       Ship To State
    Public Ship_To_Zip As String 'Char(20) NULL      Ship To Zip Code
    Public Bill_To_City As String 'Char(100) NULL
    Public Bill_To_State As String 'Char(3) NULL
    Public Bill_To_Zip As String 'Char(20) NULL
    Public Filler_0001 As String 'Char(146) NULL
    Public Id As Double 'Numeric NOT NULL
    Public In_Hands_Date As Date 'DateTime NULL

    Public Ord_GUID As String
    Public Saved As Boolean = False

    ' Supplementary fields not included in the OEORDHDR_SQL table
    Public Curr_Cd_Desc As String
    Public Cus_Type_Cd As String

    ' These fields are to be included in the export file, have to find where to link.
    Public Ship_To_Contact As String 'Char(40)   Ship To Contact Name
    Public Ship_To_Phone As String 'Char(20)   Ship To Phone No
    Public Ship_To_Email As String 'Char(40)   Ship To email address
    Public Ship_To_Terr As String 'Char(2)    Ship To Territory
    Public Ship_To_UDF1 As String 'Char(30)   Ship To User Defined Field 1
    Public Ship_To_UDF2 As String 'Char(30)   Ship To User Defined Field 2

    Public m_Dirty As Boolean

    Public Customer As New cCustomer
    Public Contacts As New cOrderContacts

    Public Ship_Link As String
    Public Cus_Email_Address As String ' from ARCUSFIL_SQL View, non editable.

    Private m_SendOrderAck As Integer
    Public ContactView As Integer

    Public ExportTS As String = ""
    Public MacolaTS As String = ""
    Public OpenTS As String = ""
    Public OrderAckSaveOnly As Integer = 0
    Public AutoCompleteReship As Integer = 0

    Public InvalidShipDateEmail As Boolean = False
    Public InvalidEmail As Boolean = False

    Public RepeatOrd_No As String = String.Empty
    Public Cus_Prog_ID As Integer
    Private m_strCus_Spector_Cd As String
    Public White_Glove As Boolean = False
    Public End_User As String = String.Empty
    Public intSSP As Int32

#End Region

    'Public Property User_Def_Fld_3() As String
    '    Get
    '        User_Def_Fld_3 = m_User_Def_Fld_3
    '    End Get
    '    Set(ByVal value As String)
    '        m_User_Def_Fld_3 = value
    '    End Set
    'End Property

    Public Sub New()

        Try
            Dim vg As New VariableGuid
            Ord_GUID = vg.Guid(30) ' Guid.NewGuid.ToString

        Catch er As Exception
            MsgBox("Error in COrdhead." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub New(ByVal pstrOrd_No As String, ByRef poDescriptions As cOrdheadDesc, ByVal pSource As OrderSourceEnum)

        'Ord_No = pstrOrd_No

        Try
            Dim vg As New VariableGuid

                If pSource = OrderSourceEnum.Macola Then
                Ord_No = pstrOrd_No
                Ord_GUID = vg.Guid(30) ' Guid.NewGuid.ToString
                Call LoadMacolaOrder(poDescriptions)
            ElseIf pSource = OrderSourceEnum.OEInterface Then
                Header_Table = "OEI_OrdHdr"
                OEI_Ord_No = pstrOrd_No.ToUpper
                'Ord_GUID = pstrOrd_No
                Call LoadOEInterfaceOrder(poDescriptions)
                Contacts = New cOrderContacts(Ord_GUID, OEI_Ord_No)
            End If

        Catch er As Exception
            MsgBox("Error in COrdhead." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#Region "Public properties ################################################"

    Public Property Cus_Alt_Adr_Cd() As String
        Get
            Cus_Alt_Adr_Cd = m_Cus_Alt_Adr_Cd
        End Get
        Set(ByVal value As String)
            m_Cus_Alt_Adr_Cd = value
        End Set
    End Property

#End Region

    Public Sub Init()

        Try
            Header_Table = String.Empty

            Record_Type = "H"               'Char (1)       'H' for header ' NOT TO BE RELOADED

            Ord_Type = String.Empty         'Char (1)
            Ord_No = String.Empty           'Char (8)       Order No    Order No from Website for tracking
            OEI_Ord_No = String.Empty       'Char (8)       Order No    Order No from Website for tracking
            Status = String.Empty           'Char(1) NOT NULL
            Entered_Dt = Now()              'DateTime NULL
            Ord_Dt = Now()                  'Datetime       Order date
            Apply_To_No = String.Empty      'Char(8) NULL
            Oe_Po_No = String.Empty         'Char(25) NOT NULL Customer PO No
            Cus_No = String.Empty           'Char(20) NOT NULL Customer No     Always '999999999999'
            Cus_Name = String.Empty         ' used for information purposes
            Bal_Meth = String.Empty         'Char(1) NULL
            Bill_To_Name = String.Empty     'Char (40)      Customer Name
            Bill_To_Addr_1 = String.Empty   'Char (40)      1st address line
            Bill_To_Addr_2 = String.Empty   'Char (40)      2nd address line
            Bill_To_Addr_3 = String.Empty   'Char (40)      3rd address line
            Bill_To_Addr_4 = String.Empty   'Char(44) NULL
            Bill_To_Country = String.Empty  'Char(3) NULL
            m_Cus_Alt_Adr_Cd = String.Empty   'Char(15) NULL
            Ship_To_Name = String.Empty     'Char(40) NULL  Ship To name
            Ship_To_Addr_1 = String.Empty   'Char(40) NULL  1st Ship To address line
            Ship_To_Addr_2 = String.Empty   'Char(40) NULL  2nd Ship To address line
            Ship_To_Addr_3 = String.Empty   'Char(40) NULL
            Ship_To_Addr_4 = String.Empty   'Char(44) NULL
            Ship_To_Country = String.Empty  'Char(3) NULL   Ship To Country
            Shipping_Dt = Now()             'DateTime NULL  Requested Ship Date
            Ship_Via_Cd = String.Empty      'Char(3) NULL   Ship Via Code
            Ar_Terms_Cd = String.Empty      'Char(2) NULL
            Frt_Pay_Cd = String.Empty       'Char(1) NULL
            Ship_Instruction_1 = String.Empty 'Char(40) NULL Shipping Instructions Line 1
            Ship_Instruction_2 = String.Empty 'Char(40) NULL Shipping Instructions Line 2
            Slspsn_No = 0                   'Int NOT NULL
            Slspsn_Pct_Comm = 0             'Numeric NULL
            Slspsn_Comm_Amt = 0             'Numeric NULL
            Slspsn_No_2 = 0                 'Int NULL
            Slspsn_Pct_Comm_2 = 0           'Numeric NULL
            Slspsn_Comm_Amt_2 = 0           'Numeric NULL
            Slspsn_No_3 = 0                 'Int NULL
            Slspsn_Pct_Comm_3 = 0           'Numeric NULL
            Slspsn_Comm_Amt_3 = 0           'Numeric NULL
            Tax_Cd = String.Empty           'Char(3) NULL   Tax Code    Applies to this order only
            Tax_Pct = 0                     'Numeric NULL
            Tax_Cd_2 = String.Empty         'Char(3) NULL
            Tax_Pct_2 = 0                   'Numeric NULL
            Tax_Cd_3 = String.Empty         'Char(3) NULL
            Tax_Pct_3 = 0                   'Numeric NULL
            Discount_Pct = 0                'Numeric NULL
            Job_No = String.Empty           'Char(20) NULL
            Mfg_Loc = String.Empty          'Char(3) NULL
            Profit_Center = String.Empty    'Char(8) NULL
            Dept = String.Empty             'Char(8) NULL
            Ar_Reference = String.Empty     'Char(45) NULL
            Tot_Sls_Amt = 0                 'Numeric NULL
            Tot_Sls_Disc = 0                'Numeric NULL
            Tot_Tax_Amt = 0                 'Numeric NULL
            Tot_Cost = 0                    'Numeric NULL
            Tot_Weight = 0                  'Numeric NULL
            Misc_Amt = 0                    'Numeric(16,2)  NULL   Miscellaneous Charge    Handling charge
            Misc_Mn_No = String.Empty       'Char(9) NULL
            Misc_Sb_No = String.Empty       'Char(8) NULL
            Misc_Dp_No = String.Empty       'Char(8) NULL
            Frt_Amt = 0                     'Numeric(16,2)  NULL   Freight Charge
            Frt_Mn_No = String.Empty        'Char(9) NULL
            Frt_Sb_No = String.Empty        'Char(8) NULL
            Frt_Dp_No = String.Empty        'Char(8) NULL
            Sls_Tax_Amt_1 = 0               'Numeric NULL
            Sls_Tax_Amt_2 = 0               'Numeric NULL
            Sls_Tax_Amt_3 = 0               'Numeric NULL
            Comm_Pct = 0                    'Numeric NULL
            Comm_Amt = 0                    'Numeric NULL
            Cmt_1 = String.Empty            'Char(35) NULL
            Cmt_2 = String.Empty            'Char(35) NULL
            Cmt_3 = String.Empty            'Char(35) NULL
            Payment_Amt = 0                 'Numeric NULL
            Payment_Disc_Amt = 0            'Numeric NULL
            Chk_No = String.Empty           'Char(8) NULL
            Chk_Dt = Now()                  'DateTime NULL
            Cash_Mn_No = String.Empty       'Char(9) NULL
            Cash_Sb_No = String.Empty       'Char(8) NULL
            Cash_Dp_No = String.Empty       'Char(8) NULL
            Ord_Dt_Picked = Now()           'DateTime NULL
            Ord_Dt_Billed = Now()           'DateTime NULL
            Inv_No = String.Empty           'Char(8) NOT NULL
            Inv_Dt = Now()                  'DateTime NULL
            Selection_Cd = String.Empty     'Char(1) NOT NULL
            Posted_Dt = Now()               'DateTime NULL
            Part_Posted_Fg = String.Empty   'Char(1) NULL
            Ship_To_Freefrm_Fg = String.Empty 'Char(1) NULL
            Bill_To_Freefrm_Fg = String.Empty 'Char(1) NULL
            Copy_To_Bm_Fg = String.Empty    'Char(1) NULL
            Edi_Fg = String.Empty           'Char(1) NULL
            Closed_Fg = String.Empty        'Char(1) NULL
            Accum_Misc_Amt = 0              'Numeric NULL
            Accum_Frt_Amt = 0               'Numeric NULL
            Accum_Tot_Tax_Amt = 0           'Numeric NULL
            Accum_Sls_Tax_Amt = 0           'Numeric NULL
            Accum_Tot_Sls_Amt = 0           'Numeric NULL
            Hold_Fg = String.Empty          'Char(1) NOT NULL
            Prepayment_Fg = String.Empty    'Char(1) NULL
            Lost_Sale_Cd = String.Empty     'Char(3) NULL
            Orig_Ord_Type = String.Empty    'Char(1) NULL
            Orig_Ord_Dt = Now()             'DateTime NULL
            Orig_Ord_No = String.Empty      'Char(8) NULL
            Award_Probability = 0           'TinyInt NULL
            Oe_Cash_No = String.Empty       'Char(8) NOT NULL
            Exch_Rt_Fg = String.Empty       'Char(1) NULL
            m_Curr_Cd = String.Empty        'Char(3) NOT NULL
            Orig_Trx_Rt = 0                 'Numeric NULL
            Curr_Trx_Rt = 0                 'Numeric NULL
            Tax_Sched = String.Empty        'Char(5) NULL
            User_Def_Fld_1 = String.Empty   'Char(30) NULL
            User_Def_Fld_2 = String.Empty   'Char(30) NULL
            User_Def_Fld_3 = String.Empty   'Char(30) NULL
            User_Def_Fld_4 = String.Empty   'Char(30) NULL
            User_Def_Fld_5 = String.Empty   'Char(30) NULL
            Deter_Rate_By = String.Empty    'Char(1) NULL
            Form_No = 0                     'TinyInt NULL
            Tax_Fg = String.Empty           'Char(1) NULL  Taxable status  'Y'es or 'N'o - Applies to this order only
            Sls_Mn_No = String.Empty        'Char(9) NULL
            Sls_Sb_No = String.Empty        'Char(8) NULL
            Sls_Dp_No = String.Empty        'Char(8) NULL
            Ord_Dt_Shipped = Now()          'DateTime NULL
            Tot_Dollars = 0                 'Numeric NULL
            Mult_Loc_Fg = String.Empty      'Char(1) NULL
            Tot_Tax_Cost = 0                'Numeric NULL
            Hist_Load_Record = String.Empty 'Char(1) NULL
            Pre_Select_Status = String.Empty 'Char(1) NULL
            Packing_No = 0                  'Int NULL
            Deliv_Ar_Terms_Cd = String.Empty 'Char(2) NULL
            Inv_Batch_Id = String.Empty     'Char(8) NULL
            Bill_To_No = String.Empty       'Char(20) NULL
            Rma_No = String.Empty           'Char(8) NOT NULL
            Prog_Term_No = 0                'Int NULL
            Extra_1 = String.Empty          'Char(1) NULL
            Extra_2 = String.Empty          'Char(1) NULL
            Extra_3 = String.Empty          'Char(1) NULL
            Extra_4 = String.Empty          'Char(1) NULL
            Extra_5 = String.Empty          'Char(1) NULL
            Extra_6 = String.Empty          'Char(8) NULL
            Extra_7 = String.Empty          'Char(8) NULL
            Extra_8 = String.Empty          'Char(12) NULL
            Extra_9 = String.Empty          'Char(12) NULL
            Extra_10 = 0                    'Numeric NULL
            Extra_11 = 0                    'Numeric NULL
            Extra_12 = 0                    'Numeric NULL
            Extra_13 = 0                    'Numeric NULL
            Extra_14 = 0                    'Int NULL
            Extra_15 = 0                    'Int NULL
            Edi_Doc_Seq = 0                 'SmallInt NULL
            Contact_1 = String.Empty        'Char(100) NULL
            Phone_Number = String.Empty     'Char(25) NULL
            Fax_Number = String.Empty       'Char(25) NULL
            Email_Address = String.Empty    'Char(128) NULL
            Use_Email = String.Empty        'Char(10) NULL
            Ship_To_City = String.Empty     'Char(100) NULL     Ship To City
            Ship_To_State = String.Empty    'Char(3) NULL       Ship To State
            Ship_To_Zip = String.Empty      'Char(20) NULL      Ship To Zip Code
            Bill_To_City = String.Empty     'Char(100) NULL
            Bill_To_State = String.Empty    'Char(3) NULL
            Bill_To_Zip = String.Empty      'Char(20) NULL
            Filler_0001 = String.Empty      'Char(146) NULL
            In_Hands_Date = NoDate()        'DateTime NULL
            Id = 0                          'Numeric NOT NULL

            Ord_GUID = String.Empty
            Saved = False

            ' Supplementary fields not included in the OEORDHDR_SQL table
            Curr_Cd_Desc = String.Empty
            Cus_Type_Cd = String.Empty

            ' These fields are to be included in the export file, have to find where to link.
            Ship_To_Contact = String.Empty  'Char(40)   Ship To Contact Name
            Ship_To_Phone = String.Empty    'Char(20)   Ship To Phone No
            Ship_To_Email = String.Empty    'Char(40)   Ship To email address
            Ship_To_Terr = String.Empty     'Char(2)    Ship To Territory
            Ship_To_UDF1 = String.Empty     'Char(30)   Ship To User Defined Field 1
            Ship_To_UDF2 = String.Empty     'Char(30)   Ship To User Defined Field 2

            m_Dirty = False

            Customer = New cCustomer
            Contacts = New cOrderContacts

            Ship_Link = String.Empty
            SendOrderAck = 0
            ContactView = 0
            OrderAckSaveOnly = 0
            AutoCompleteReship = 0
            Cus_Prog_ID = 0
            Prog_Spector_Cd = String.Empty

            White_Glove = False

            intSSP = 0

        Catch er As Exception
            MsgBox("Error in COrdhead." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub ClearCustomerValues()

        Try
            Curr_Cd = String.Empty
            Cus_Type_Cd = String.Empty
            Cus_Name = String.Empty

            Bill_To_City = String.Empty
            Bill_To_State = String.Empty
            Bill_To_Zip = String.Empty

            Form_No = 0

            Tax_Sched = String.Empty
            Tax_Fg = String.Empty

            Tax_Cd = String.Empty
            Tax_Pct = 0

            Tax_Cd_2 = String.Empty
            Tax_Pct_2 = 0

            Tax_Cd_3 = String.Empty
            Tax_Pct_3 = 0

            m_Cus_Alt_Adr_Cd = String.Empty

            Oe_Po_No = String.Empty
            Slspsn_No = 0
            Job_No = String.Empty

            Mfg_Loc = String.Empty
            Ship_Via_Cd = String.Empty
            Ar_Terms_Cd = String.Empty
            Discount_Pct = 0
            Profit_Center = String.Empty

            Dept = String.Empty

            Ship_Instruction_1 = String.Empty
            Ship_Instruction_2 = String.Empty

            Bill_To_Name = String.Empty
            Bill_To_Addr_1 = String.Empty
            Bill_To_Addr_2 = String.Empty
            Bill_To_Addr_3 = String.Empty
            Bill_To_Addr_4 = String.Empty
            Bill_To_City = String.Empty
            Bill_To_State = String.Empty
            Bill_To_Zip = String.Empty
            Bill_To_Country = String.Empty
            'txtBill_To_CountryDesc.Text = rsc.Fields("").Value 'will populate on country change

            'm_oOrder.Ordhead.Curr_Cd = IIf(System.DBNull.Value.Equals(rsc.Fields("Curr_Cd").Value), "", rsc.Fields("Curr_Cd").Value)

            Ship_To_Name = String.Empty
            Ship_To_Addr_1 = String.Empty
            Ship_To_Addr_2 = String.Empty
            Ship_To_Addr_3 = String.Empty
            Ship_To_Addr_4 = String.Empty
            Ship_To_Country = String.Empty
            Ship_To_City = String.Empty
            Ship_To_State = String.Empty
            Ship_To_Zip = String.Empty

            'ShipToCountryDesc.Text = String.Empty

            Contact_1 = String.Empty
            Phone_Number = String.Empty
            Fax_Number = String.Empty
            Email_Address = String.Empty

            Slspsn_Pct_Comm = 100
            Slspsn_Pct_Comm_2 = 0
            Slspsn_Pct_Comm_3 = 0

            Deter_Rate_By = String.Empty

            Cus_Email_Address = String.Empty
            User_Def_Fld_2 = String.Empty
            User_Def_Fld_3 = String.Empty
            'Slspsn_Name 

            Customer = New cCustomer
            Contacts = New cOrderContacts

        Catch er As Exception
            MsgBox("Error in COrdhead." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub LoadMacolaOrder(ByRef poDescriptions As cOrdheadDesc)

        Dim dt As New DataTable
        Dim db As New cDBA

        Dim strSql As String

        Try
            'strSql = "" & _
            '"SELECT TOP 1 * " & _
            '"FROM   OEOrdHdr_Sql WITH (Nolock) " & _
            '"WHERE  Ord_No = '" & Ord_No & "' "

            Call Set_Header_Table_Name(Ord_No)

            strSql = "" & _
            "SELECT 	TOP 1 O.*, ISNULL(Loc.Loc_Desc, '') as Location_Desc, ISNULL(SV.Code_Desc, '') AS Ship_Via_Desc, " & _
            "			CASE ISNULL(O.Status, '') WHEN '1' THEN 'Booked Order' WHEN '4' THEN 'Pick Ticket printed' WHEN '8' THEN 'Selected for Billing' WHEN 'C' THEN 'Credit Hold' WHEN 'I' THEN 'Order Incomplete' WHEN 'L' THEN 'Closed Order' ELSE 'Unknown' END AS Status_Desc, " & _
            "			ISNULL(Term.Description, '') as Term_Description, ISNULL(TS.Tax_Sched_Desc, '') AS Tax_Sched_Desc, " & _
            "			ISNULL(Tax1.Tax_Cd_Description, '') as Tax_Cd_1_Desc, " & _
            "			ISNULL(Tax2.Tax_Cd_Description, '') as Tax_Cd_2_Desc, " & _
            "			ISNULL(Tax3.Tax_Cd_Description, '') as Tax_Cd_3_Desc, " & _
            "			ISNULL(Sls1.Slspsn_Name, '') AS Slspsn_1_Desc, " & _
            "			ISNULL(Sls2.Slspsn_Name, '') AS Slspsn_2_Desc, " & _
            "			ISNULL(Sls3.Slspsn_Name, '') AS Slspsn_3_Desc " & _
            "FROM   	" & Header_Table & " O    WITH (Nolock)  " & _
            "LEFT JOIN	IMLOCFIL_SQL Loc  WITH (Nolock) ON O.Mfg_Loc = Loc.Loc " & _
            "LEFT JOIN 	SYCDEFIL_SQL SV   WITH (Nolock) ON O.Ship_Via_Cd = SV.sy_code " & _
            "LEFT JOIN 	SYTRMFIL_SQL Term WITH (Nolock) ON O.AR_Terms_CD = Term.Term_Code " & _
            "LEFT JOIN 	TAXSCHED_SQL TS   WITH (Nolock) ON O.Tax_Sched = TS.Tax_Sched " & _
            "LEFT JOIN 	TAXDETL_SQL  Tax1 WITH (Nolock) ON O.Tax_Cd = Tax1.Tax_Cd " & _
            "LEFT JOIN 	TAXDETL_SQL  Tax2 WITH (Nolock) ON O.Tax_Cd_2 = Tax2.Tax_Cd " & _
            "LEFT JOIN 	TAXDETL_SQL  Tax3 WITH (Nolock) ON O.Tax_Cd_3 = Tax3.Tax_Cd " & _
            "LEFT JOIN 	ARSLMFIL_SQL Sls1 WITH (Nolock) ON O.SlsPsn_No = Sls1.HumRes_ID " & _
            "LEFT JOIN 	ARSLMFIL_SQL Sls2 WITH (Nolock) ON O.SlsPsn_No_2 = Sls2.HumRes_ID " & _
            "LEFT JOIN 	ARSLMFIL_SQL Sls3 WITH (Nolock) ON O.SlsPsn_No_3 = Sls3.HumRes_ID " & _
            "WHERE  Ord_No = '" & Ord_No.ToString.PadLeft(8) & "' "

            dt = db.DataTable(strSql)
            If dt.Rows.Count <> 0 Then

                If Not dt.Rows(0).Item("Ord_No").Equals(DBNull.Value) Then Ord_No = dt.Rows(0).Item("Ord_No") 'Char (8)       Order No    Order No from Website for tracking
                If Not dt.Rows(0).Item("Status").Equals(DBNull.Value) Then Status = dt.Rows(0).Item("Status") 'Char(1) NOT NULL
                If Not dt.Rows(0).Item("Entered_Dt").Equals(DBNull.Value) Then Entered_Dt = dt.Rows(0).Item("Entered_Dt") 'DateTime NULL
                If Not dt.Rows(0).Item("Ord_Dt").Equals(DBNull.Value) Then Ord_Dt = dt.Rows(0).Item("Ord_Dt") 'Datetime       Order date
                If Not dt.Rows(0).Item("Apply_To_No").Equals(DBNull.Value) Then Apply_To_No = dt.Rows(0).Item("Apply_To_No") 'Char(8) NULL
                If Not dt.Rows(0).Item("Oe_Po_No").Equals(DBNull.Value) Then Oe_Po_No = dt.Rows(0).Item("Oe_Po_No") 'Char(25) NOT NULL Customer PO No
                If Not dt.Rows(0).Item("Cus_No").Equals(DBNull.Value) Then Cus_No = dt.Rows(0).Item("Cus_No") 'Char(20) NOT NULL Customer No     Always '999999999999'
                If Not dt.Rows(0).Item("Bal_Meth").Equals(DBNull.Value) Then Bal_Meth = dt.Rows(0).Item("Bal_Meth") 'Char(1) NULL
                If Not dt.Rows(0).Item("Bill_To_Name").Equals(DBNull.Value) Then Bill_To_Name = dt.Rows(0).Item("Bill_To_Name") 'Char (40)      Customer Name
                If Not dt.Rows(0).Item("Bill_To_Addr_1").Equals(DBNull.Value) Then Bill_To_Addr_1 = dt.Rows(0).Item("Bill_To_Addr_1") 'Char (40)      1st address line
                If Not dt.Rows(0).Item("Bill_To_Addr_2").Equals(DBNull.Value) Then Bill_To_Addr_2 = dt.Rows(0).Item("Bill_To_Addr_2") 'Char (40)      2nd address line
                If Not dt.Rows(0).Item("Bill_To_Addr_3").Equals(DBNull.Value) Then Bill_To_Addr_3 = dt.Rows(0).Item("Bill_To_Addr_3") 'Char (40)      3rd address line
                If Not dt.Rows(0).Item("Bill_To_Addr_4").Equals(DBNull.Value) Then Bill_To_Addr_4 = dt.Rows(0).Item("Bill_To_Addr_4") 'Char(44) NULL
                If Not dt.Rows(0).Item("Bill_To_Country").Equals(DBNull.Value) Then Bill_To_Country = dt.Rows(0).Item("Bill_To_Country") 'Char(3) NULL
                If Not dt.Rows(0).Item("Cus_Alt_Adr_Cd").Equals(DBNull.Value) Then m_Cus_Alt_Adr_Cd = dt.Rows(0).Item("Cus_Alt_Adr_Cd") 'Char(15) NULL
                If Not dt.Rows(0).Item("Ship_To_Name").Equals(DBNull.Value) Then Ship_To_Name = dt.Rows(0).Item("Ship_To_Name") 'Char(40) NULL  Ship To name
                If Not dt.Rows(0).Item("Ship_To_Addr_1").Equals(DBNull.Value) Then Ship_To_Addr_1 = dt.Rows(0).Item("Ship_To_Addr_1") 'Char(40) NULL  1st Ship To address line
                If Not dt.Rows(0).Item("Ship_To_Addr_2").Equals(DBNull.Value) Then Ship_To_Addr_2 = dt.Rows(0).Item("Ship_To_Addr_2") 'Char(40) NULL  2nd Ship To address line
                If Not dt.Rows(0).Item("Ship_To_Addr_3").Equals(DBNull.Value) Then Ship_To_Addr_3 = dt.Rows(0).Item("Ship_To_Addr_3") 'Char(40) NULL
                If Not dt.Rows(0).Item("Ship_To_Addr_4").Equals(DBNull.Value) Then Ship_To_Addr_4 = dt.Rows(0).Item("Ship_To_Addr_4") 'Char(44) NULL
                If Not dt.Rows(0).Item("Ship_To_Country").Equals(DBNull.Value) Then Ship_To_Country = dt.Rows(0).Item("Ship_To_Country") 'Char(3) NULL   Ship To Country
                If Not dt.Rows(0).Item("Shipping_Dt").Equals(DBNull.Value) Then Shipping_Dt = dt.Rows(0).Item("Shipping_Dt") 'DateTime NULL  Requested Ship Date
                If Not dt.Rows(0).Item("Ship_Via_Cd").Equals(DBNull.Value) Then Ship_Via_Cd = dt.Rows(0).Item("Ship_Via_Cd") 'Char(3) NULL   Ship Via Code
                If Not dt.Rows(0).Item("Ar_Terms_Cd").Equals(DBNull.Value) Then Ar_Terms_Cd = dt.Rows(0).Item("Ar_Terms_Cd") 'Char(2) NULL
                If Not dt.Rows(0).Item("Frt_Pay_Cd").Equals(DBNull.Value) Then Frt_Pay_Cd = dt.Rows(0).Item("Frt_Pay_Cd") 'Char(1) NULL
                If Not dt.Rows(0).Item("Ship_Instruction_1").Equals(DBNull.Value) Then Ship_Instruction_1 = dt.Rows(0).Item("Ship_Instruction_1") 'Char(40) NULL Shipping Instructions Line 1
                If Not dt.Rows(0).Item("Ship_Instruction_2").Equals(DBNull.Value) Then Ship_Instruction_2 = dt.Rows(0).Item("Ship_Instruction_2") 'Char(40) NULL Shipping Instructions Line 2
                If Not dt.Rows(0).Item("Slspsn_No").Equals(DBNull.Value) Then Slspsn_No = dt.Rows(0).Item("Slspsn_No") 'Int NOT NULL
                If Not dt.Rows(0).Item("Slspsn_Pct_Comm").Equals(DBNull.Value) Then Slspsn_Pct_Comm = dt.Rows(0).Item("Slspsn_Pct_Comm") 'Numeric NULL
                If Not dt.Rows(0).Item("Slspsn_Comm_Amt").Equals(DBNull.Value) Then Slspsn_Comm_Amt = dt.Rows(0).Item("Slspsn_Comm_Amt") 'Numeric NULL
                If Not dt.Rows(0).Item("Slspsn_No_2").Equals(DBNull.Value) Then Slspsn_No_2 = dt.Rows(0).Item("Slspsn_No_2") 'Int NULL
                If Not dt.Rows(0).Item("Slspsn_Pct_Comm_2").Equals(DBNull.Value) Then Slspsn_Pct_Comm_2 = dt.Rows(0).Item("Slspsn_Pct_Comm_2") 'Numeric NULL
                If Not dt.Rows(0).Item("Slspsn_Comm_Amt_2").Equals(DBNull.Value) Then Slspsn_Comm_Amt_2 = dt.Rows(0).Item("Slspsn_Comm_Amt_2") 'Numeric NULL
                If Not dt.Rows(0).Item("Slspsn_No_3").Equals(DBNull.Value) Then Slspsn_No_3 = dt.Rows(0).Item("Slspsn_No_3") 'Int NULL
                If Not dt.Rows(0).Item("Slspsn_Pct_Comm_3").Equals(DBNull.Value) Then Slspsn_Pct_Comm_3 = dt.Rows(0).Item("Slspsn_Pct_Comm_3") 'Numeric NULL
                If Not dt.Rows(0).Item("Slspsn_Comm_Amt_3").Equals(DBNull.Value) Then Slspsn_Comm_Amt_3 = dt.Rows(0).Item("Slspsn_Comm_Amt_3") 'Numeric NULL
                If Not dt.Rows(0).Item("Tax_Cd").Equals(DBNull.Value) Then Tax_Cd = dt.Rows(0).Item("Tax_Cd") 'Char(3) NULL   Tax Code    Applies to this order only
                If Not dt.Rows(0).Item("Tax_Pct").Equals(DBNull.Value) Then Tax_Pct = dt.Rows(0).Item("Tax_Pct") 'Numeric NULL
                If Not dt.Rows(0).Item("Tax_Cd_2").Equals(DBNull.Value) Then Tax_Cd_2 = dt.Rows(0).Item("Tax_Cd_2") 'Char(3) NULL
                If Not dt.Rows(0).Item("Tax_Pct_2").Equals(DBNull.Value) Then Tax_Pct_2 = dt.Rows(0).Item("Tax_Pct_2") 'Numeric NULL
                If Not dt.Rows(0).Item("Tax_Cd_3").Equals(DBNull.Value) Then Tax_Cd_3 = dt.Rows(0).Item("Tax_Cd_3") 'Char(3) NULL
                If Not dt.Rows(0).Item("Tax_Pct_3").Equals(DBNull.Value) Then Tax_Pct_3 = dt.Rows(0).Item("Tax_Pct_3") 'Numeric NULL
                If Not dt.Rows(0).Item("Discount_Pct").Equals(DBNull.Value) Then Discount_Pct = dt.Rows(0).Item("Discount_Pct") 'Numeric NULL
                If Not dt.Rows(0).Item("Job_No").Equals(DBNull.Value) Then Job_No = dt.Rows(0).Item("Job_No") 'Char(20) NULL
                If Not dt.Rows(0).Item("Mfg_Loc").Equals(DBNull.Value) Then Mfg_Loc = dt.Rows(0).Item("Mfg_Loc") 'Char(3) NULL
                If Not dt.Rows(0).Item("Profit_Center").Equals(DBNull.Value) Then Profit_Center = dt.Rows(0).Item("Profit_Center") 'Char(8) NULL
                If Not dt.Rows(0).Item("Dept").Equals(DBNull.Value) Then Dept = dt.Rows(0).Item("Dept") 'Char(8) NULL
                If Not dt.Rows(0).Item("Ar_Reference").Equals(DBNull.Value) Then Ar_Reference = dt.Rows(0).Item("Ar_Reference") 'Char(45) NULL
                If Not dt.Rows(0).Item("Tot_Sls_Amt").Equals(DBNull.Value) Then Tot_Sls_Amt = dt.Rows(0).Item("Tot_Sls_Amt") 'Numeric NULL
                If Not dt.Rows(0).Item("Tot_Sls_Disc").Equals(DBNull.Value) Then Tot_Sls_Disc = dt.Rows(0).Item("Tot_Sls_Disc") 'Numeric NULL
                If Not dt.Rows(0).Item("Tot_Tax_Amt").Equals(DBNull.Value) Then Tot_Tax_Amt = dt.Rows(0).Item("Tot_Tax_Amt") 'Numeric NULL
                If Not dt.Rows(0).Item("Tot_Cost").Equals(DBNull.Value) Then Tot_Cost = dt.Rows(0).Item("Tot_Cost") 'Numeric NULL
                If Not dt.Rows(0).Item("Tot_Weight").Equals(DBNull.Value) Then Tot_Weight = dt.Rows(0).Item("Tot_Weight") 'Numeric NULL
                If Not dt.Rows(0).Item("Misc_Amt").Equals(DBNull.Value) Then Misc_Amt = dt.Rows(0).Item("Misc_Amt") 'Numeric(16,2)  NULL   Miscellaneous Charge    Handling charge
                If Not dt.Rows(0).Item("Misc_Mn_No").Equals(DBNull.Value) Then Misc_Mn_No = dt.Rows(0).Item("Misc_Mn_No") 'Char(9) NULL
                If Not dt.Rows(0).Item("Misc_Sb_No").Equals(DBNull.Value) Then Misc_Sb_No = dt.Rows(0).Item("Misc_Sb_No") 'Char(8) NULL
                If Not dt.Rows(0).Item("Misc_Dp_No").Equals(DBNull.Value) Then Misc_Dp_No = dt.Rows(0).Item("Misc_Dp_No") 'Char(8) NULL
                If Not dt.Rows(0).Item("Frt_Amt").Equals(DBNull.Value) Then Frt_Amt = dt.Rows(0).Item("Frt_Amt") 'Numeric(16,2)  NULL   Freight Charge
                If Not dt.Rows(0).Item("Frt_Mn_No").Equals(DBNull.Value) Then Frt_Mn_No = dt.Rows(0).Item("Frt_Mn_No") 'Char(9) NULL
                If Not dt.Rows(0).Item("Frt_Sb_No").Equals(DBNull.Value) Then Frt_Sb_No = dt.Rows(0).Item("Frt_Sb_No") 'Char(8) NULL
                If Not dt.Rows(0).Item("Frt_Dp_No").Equals(DBNull.Value) Then Frt_Dp_No = dt.Rows(0).Item("Frt_Dp_No") 'Char(8) NULL
                If Not dt.Rows(0).Item("Sls_Tax_Amt_1").Equals(DBNull.Value) Then Sls_Tax_Amt_1 = dt.Rows(0).Item("Sls_Tax_Amt_1") 'Numeric NULL
                If Not dt.Rows(0).Item("Sls_Tax_Amt_2").Equals(DBNull.Value) Then Sls_Tax_Amt_2 = dt.Rows(0).Item("Sls_Tax_Amt_2") 'Numeric NULL
                If Not dt.Rows(0).Item("Sls_Tax_Amt_3").Equals(DBNull.Value) Then Sls_Tax_Amt_3 = dt.Rows(0).Item("Sls_Tax_Amt_3") 'Numeric NULL
                If Not dt.Rows(0).Item("Comm_Pct").Equals(DBNull.Value) Then Comm_Pct = dt.Rows(0).Item("Comm_Pct") 'Numeric NULL
                If Not dt.Rows(0).Item("Comm_Amt").Equals(DBNull.Value) Then Comm_Amt = dt.Rows(0).Item("Comm_Amt") 'Numeric NULL
                If Not dt.Rows(0).Item("Cmt_1").Equals(DBNull.Value) Then Cmt_1 = dt.Rows(0).Item("Cmt_1") 'Char(35) NULL
                If Not dt.Rows(0).Item("Cmt_2").Equals(DBNull.Value) Then Cmt_2 = dt.Rows(0).Item("Cmt_2") 'Char(35) NULL
                If Not dt.Rows(0).Item("Cmt_3").Equals(DBNull.Value) Then Cmt_3 = dt.Rows(0).Item("Cmt_3") 'Char(35) NULL
                If Not dt.Rows(0).Item("Payment_Amt").Equals(DBNull.Value) Then Payment_Amt = dt.Rows(0).Item("Payment_Amt") 'Numeric NULL
                If Not dt.Rows(0).Item("Payment_Disc_Amt").Equals(DBNull.Value) Then Payment_Disc_Amt = dt.Rows(0).Item("Payment_Disc_Amt") 'Numeric NULL
                If Not dt.Rows(0).Item("Chk_No").Equals(DBNull.Value) Then Chk_No = dt.Rows(0).Item("Chk_No") 'Char(8) NULL
                If Not dt.Rows(0).Item("Chk_Dt").Equals(DBNull.Value) Then Chk_Dt = dt.Rows(0).Item("Chk_Dt") 'DateTime NULL
                If Not dt.Rows(0).Item("Cash_Mn_No").Equals(DBNull.Value) Then Cash_Mn_No = dt.Rows(0).Item("Cash_Mn_No") 'Char(9) NULL
                If Not dt.Rows(0).Item("Cash_Sb_No").Equals(DBNull.Value) Then Cash_Sb_No = dt.Rows(0).Item("Cash_Sb_No") 'Char(8) NULL
                If Not dt.Rows(0).Item("Cash_Dp_No").Equals(DBNull.Value) Then Cash_Dp_No = dt.Rows(0).Item("Cash_Dp_No") 'Char(8) NULL
                If Not dt.Rows(0).Item("Ord_Dt_Picked").Equals(DBNull.Value) Then Ord_Dt_Picked = dt.Rows(0).Item("Ord_Dt_Picked") 'DateTime NULL
                If Not dt.Rows(0).Item("Ord_Dt_Billed").Equals(DBNull.Value) Then Ord_Dt_Billed = dt.Rows(0).Item("Ord_Dt_Billed") 'DateTime NULL
                If Not dt.Rows(0).Item("Inv_No").Equals(DBNull.Value) Then Inv_No = dt.Rows(0).Item("Inv_No") 'Char(8) NOT NULL
                If Not dt.Rows(0).Item("Inv_Dt").Equals(DBNull.Value) Then Inv_Dt = dt.Rows(0).Item("Inv_Dt") 'DateTime NULL
                If Not dt.Rows(0).Item("Selection_Cd").Equals(DBNull.Value) Then Selection_Cd = dt.Rows(0).Item("Selection_Cd") 'Char(1) NOT NULL
                If Not dt.Rows(0).Item("Posted_Dt").Equals(DBNull.Value) Then Posted_Dt = dt.Rows(0).Item("Posted_Dt") 'DateTime NULL
                If Not dt.Rows(0).Item("Part_Posted_Fg").Equals(DBNull.Value) Then Part_Posted_Fg = dt.Rows(0).Item("Part_Posted_Fg") 'Char(1) NULL
                If Not dt.Rows(0).Item("Ship_To_Freefrm_Fg").Equals(DBNull.Value) Then Ship_To_Freefrm_Fg = dt.Rows(0).Item("Ship_To_Freefrm_Fg") 'Char(1) NULL
                If Not dt.Rows(0).Item("Bill_To_Freefrm_Fg").Equals(DBNull.Value) Then Bill_To_Freefrm_Fg = dt.Rows(0).Item("Bill_To_Freefrm_Fg") 'Char(1) NULL
                If Not dt.Rows(0).Item("Copy_To_Bm_Fg").Equals(DBNull.Value) Then Copy_To_Bm_Fg = dt.Rows(0).Item("Copy_To_Bm_Fg") 'Char(1) NULL
                If Not dt.Rows(0).Item("Edi_Fg").Equals(DBNull.Value) Then Edi_Fg = dt.Rows(0).Item("Edi_Fg") 'Char(1) NULL
                If Not dt.Rows(0).Item("Closed_Fg").Equals(DBNull.Value) Then Closed_Fg = dt.Rows(0).Item("Closed_Fg") 'Char(1) NULL
                If Not dt.Rows(0).Item("Accum_Misc_Amt").Equals(DBNull.Value) Then Accum_Misc_Amt = dt.Rows(0).Item("Accum_Misc_Amt") 'Numeric NULL
                If Not dt.Rows(0).Item("Accum_Frt_Amt").Equals(DBNull.Value) Then Accum_Frt_Amt = dt.Rows(0).Item("Accum_Frt_Amt") 'Numeric NULL
                If Not dt.Rows(0).Item("Accum_Tot_Tax_Amt").Equals(DBNull.Value) Then Accum_Tot_Tax_Amt = dt.Rows(0).Item("Accum_Tot_Tax_Amt") 'Numeric NULL
                If Not dt.Rows(0).Item("Accum_Sls_Tax_Amt").Equals(DBNull.Value) Then Accum_Sls_Tax_Amt = dt.Rows(0).Item("Accum_Sls_Tax_Amt") 'Numeric NULL
                If Not dt.Rows(0).Item("Accum_Tot_Sls_Amt").Equals(DBNull.Value) Then Accum_Tot_Sls_Amt = dt.Rows(0).Item("Accum_Tot_Sls_Amt") 'Numeric NULL
                If Not dt.Rows(0).Item("Hold_Fg").Equals(DBNull.Value) Then Hold_Fg = dt.Rows(0).Item("Hold_Fg") 'Char(1) NOT NULL
                If Not dt.Rows(0).Item("Prepayment_Fg").Equals(DBNull.Value) Then Prepayment_Fg = dt.Rows(0).Item("Prepayment_Fg") 'Char(1) NULL
                If Not dt.Rows(0).Item("Lost_Sale_Cd").Equals(DBNull.Value) Then Lost_Sale_Cd = dt.Rows(0).Item("Lost_Sale_Cd") 'Char(3) NULL
                If Not dt.Rows(0).Item("Orig_Ord_Type").Equals(DBNull.Value) Then Orig_Ord_Type = dt.Rows(0).Item("Orig_Ord_Type") 'Char(1) NULL
                If Not dt.Rows(0).Item("Orig_Ord_Dt").Equals(DBNull.Value) Then Orig_Ord_Dt = dt.Rows(0).Item("Orig_Ord_Dt") 'DateTime NULL
                If Not dt.Rows(0).Item("Orig_Ord_No").Equals(DBNull.Value) Then Orig_Ord_No = dt.Rows(0).Item("Orig_Ord_No") 'Char(8) NULL
                If Not dt.Rows(0).Item("Award_Probability").Equals(DBNull.Value) Then Award_Probability = dt.Rows(0).Item("Award_Probability") 'TinyInt NULL
                If Not dt.Rows(0).Item("Oe_Cash_No").Equals(DBNull.Value) Then Oe_Cash_No = dt.Rows(0).Item("Oe_Cash_No") 'Char(8) NOT NULL
                If Not dt.Rows(0).Item("Exch_Rt_Fg").Equals(DBNull.Value) Then Exch_Rt_Fg = dt.Rows(0).Item("Exch_Rt_Fg") 'Char(1) NULL
                If Not dt.Rows(0).Item("Curr_Cd").Equals(DBNull.Value) Then Curr_Cd = dt.Rows(0).Item("Curr_Cd") 'Char(3) NOT NULL
                If Not dt.Rows(0).Item("Orig_Trx_Rt").Equals(DBNull.Value) Then Orig_Trx_Rt = dt.Rows(0).Item("Orig_Trx_Rt") 'Numeric NULL
                If Not dt.Rows(0).Item("Curr_Trx_Rt").Equals(DBNull.Value) Then Curr_Trx_Rt = dt.Rows(0).Item("Curr_Trx_Rt") 'Numeric NULL
                If Not dt.Rows(0).Item("Tax_Sched").Equals(DBNull.Value) Then Tax_Sched = dt.Rows(0).Item("Tax_Sched") 'Char(5) NULL
                If Not dt.Rows(0).Item("User_Def_Fld_1").Equals(DBNull.Value) Then User_Def_Fld_1 = dt.Rows(0).Item("User_Def_Fld_1") 'Char(30) NULL
                If Not dt.Rows(0).Item("User_Def_Fld_2").Equals(DBNull.Value) Then User_Def_Fld_2 = dt.Rows(0).Item("User_Def_Fld_2") 'Char(30) NULL
                If Not dt.Rows(0).Item("User_Def_Fld_3").Equals(DBNull.Value) Then User_Def_Fld_3 = dt.Rows(0).Item("User_Def_Fld_3") 'Char(30) NULL
                If Not dt.Rows(0).Item("User_Def_Fld_4").Equals(DBNull.Value) Then User_Def_Fld_4 = dt.Rows(0).Item("User_Def_Fld_4") 'Char(30) NULL
                If Not dt.Rows(0).Item("User_Def_Fld_5").Equals(DBNull.Value) Then User_Def_Fld_5 = dt.Rows(0).Item("User_Def_Fld_5") 'Char(30) NULL
                If Not dt.Rows(0).Item("Deter_Rate_By").Equals(DBNull.Value) Then Deter_Rate_By = dt.Rows(0).Item("Deter_Rate_By") 'Char(1) NULL
                If Not dt.Rows(0).Item("Form_No").Equals(DBNull.Value) Then Form_No = dt.Rows(0).Item("Form_No") 'TinyInt NULL
                If Not dt.Rows(0).Item("Tax_Fg").Equals(DBNull.Value) Then Tax_Fg = dt.Rows(0).Item("Tax_Fg") 'Char(1) NULL  Taxable status  'Y'es or 'N'o - Applies to this order only
                If Not dt.Rows(0).Item("Sls_Mn_No").Equals(DBNull.Value) Then Sls_Mn_No = dt.Rows(0).Item("Sls_Mn_No") 'Char(9) NULL
                If Not dt.Rows(0).Item("Sls_Sb_No").Equals(DBNull.Value) Then Sls_Sb_No = dt.Rows(0).Item("Sls_Sb_No") 'Char(8) NULL
                If Not dt.Rows(0).Item("Sls_Dp_No").Equals(DBNull.Value) Then Sls_Dp_No = dt.Rows(0).Item("Sls_Dp_No") 'Char(8) NULL
                If Not dt.Rows(0).Item("Ord_Dt_Shipped").Equals(DBNull.Value) Then Ord_Dt_Shipped = dt.Rows(0).Item("Ord_Dt_Shipped") 'DateTime NULL
                If Not dt.Rows(0).Item("Tot_Dollars").Equals(DBNull.Value) Then Tot_Dollars = dt.Rows(0).Item("Tot_Dollars") 'Numeric NULL
                If Not dt.Rows(0).Item("Mult_Loc_Fg").Equals(DBNull.Value) Then Mult_Loc_Fg = dt.Rows(0).Item("Mult_Loc_Fg") 'Char(1) NULL
                If Not dt.Rows(0).Item("Tot_Tax_Cost").Equals(DBNull.Value) Then Tot_Tax_Cost = dt.Rows(0).Item("Tot_Tax_Cost") 'Numeric NULL
                If Not dt.Rows(0).Item("Hist_Load_Record").Equals(DBNull.Value) Then Hist_Load_Record = dt.Rows(0).Item("Hist_Load_Record") 'Char(1) NULL
                If Not dt.Rows(0).Item("Pre_Select_Status").Equals(DBNull.Value) Then Pre_Select_Status = dt.Rows(0).Item("Pre_Select_Status") 'Char(1) NULL
                If Not dt.Rows(0).Item("Packing_No").Equals(DBNull.Value) Then Packing_No = dt.Rows(0).Item("Packing_No") 'Int NULL
                If Not dt.Rows(0).Item("Deliv_Ar_Terms_Cd").Equals(DBNull.Value) Then Deliv_Ar_Terms_Cd = dt.Rows(0).Item("Deliv_Ar_Terms_Cd") 'Char(2) NULL
                If Not dt.Rows(0).Item("Inv_Batch_Id").Equals(DBNull.Value) Then Inv_Batch_Id = dt.Rows(0).Item("Inv_Batch_Id") 'Char(8) NULL
                If Not dt.Rows(0).Item("Bill_To_No").Equals(DBNull.Value) Then Bill_To_No = dt.Rows(0).Item("Bill_To_No") 'Char(20) NULL
                If Not dt.Rows(0).Item("Rma_No").Equals(DBNull.Value) Then Rma_No = dt.Rows(0).Item("Rma_No") 'Char(8) NOT NULL
                If Not dt.Rows(0).Item("Prog_Term_No").Equals(DBNull.Value) Then Prog_Term_No = dt.Rows(0).Item("Prog_Term_No") 'Int NULL
                If Not dt.Rows(0).Item("Extra_1").Equals(DBNull.Value) Then Extra_1 = dt.Rows(0).Item("Extra_1") 'Char(1) NULL
                If Not dt.Rows(0).Item("Extra_2").Equals(DBNull.Value) Then Extra_2 = dt.Rows(0).Item("Extra_2") 'Char(1) NULL
                If Not dt.Rows(0).Item("Extra_3").Equals(DBNull.Value) Then Extra_3 = dt.Rows(0).Item("Extra_3") 'Char(1) NULL
                If Not dt.Rows(0).Item("Extra_4").Equals(DBNull.Value) Then Extra_4 = dt.Rows(0).Item("Extra_4") 'Char(1) NULL
                If Not dt.Rows(0).Item("Extra_5").Equals(DBNull.Value) Then Extra_5 = dt.Rows(0).Item("Extra_5") 'Char(1) NULL
                'Extra_6 and extra_7 are set by the trigger so we leave it blank.
                Extra_6 = String.Empty
                Extra_7 = String.Empty
                'If Not dt.Rows(0).Item("Extra_6").Equals(DBNull.Value) Then Extra_6 = dt.Rows(0).Item("Extra_6") 'Char(8) NULL
                'If Not dt.Rows(0).Item("Extra_7").Equals(DBNull.Value) Then Extra_7 = dt.Rows(0).Item("Extra_7") 'Char(8) NULL
                If Not dt.Rows(0).Item("Extra_8").Equals(DBNull.Value) Then Extra_8 = dt.Rows(0).Item("Extra_8") 'Char(12) NULL
                If Not dt.Rows(0).Item("Extra_9").Equals(DBNull.Value) Then Extra_9 = dt.Rows(0).Item("Extra_9") 'Char(12) NULL
                If Not dt.Rows(0).Item("Extra_10").Equals(DBNull.Value) Then Extra_10 = dt.Rows(0).Item("Extra_10") 'Numeric NULL
                If Not dt.Rows(0).Item("Extra_11").Equals(DBNull.Value) Then Extra_11 = dt.Rows(0).Item("Extra_11") 'Numeric NULL
                If Not dt.Rows(0).Item("Extra_12").Equals(DBNull.Value) Then Extra_12 = dt.Rows(0).Item("Extra_12") 'Numeric NULL
                If Not dt.Rows(0).Item("Extra_13").Equals(DBNull.Value) Then Extra_13 = dt.Rows(0).Item("Extra_13") 'Numeric NULL
                If Not dt.Rows(0).Item("Extra_14").Equals(DBNull.Value) Then Extra_14 = dt.Rows(0).Item("Extra_14") 'Int NULL
                If Not dt.Rows(0).Item("Extra_15").Equals(DBNull.Value) Then Extra_15 = dt.Rows(0).Item("Extra_15") 'Int NULL
                If Not dt.Rows(0).Item("Edi_Doc_Seq").Equals(DBNull.Value) Then Edi_Doc_Seq = dt.Rows(0).Item("Edi_Doc_Seq") 'SmallInt NULL
                If Not dt.Rows(0).Item("Contact_1").Equals(DBNull.Value) Then Contact_1 = dt.Rows(0).Item("Contact_1") 'Char(100) NULL
                If Not dt.Rows(0).Item("Phone_Number").Equals(DBNull.Value) Then Phone_Number = dt.Rows(0).Item("Phone_Number") 'Char(25) NULL
                If Not dt.Rows(0).Item("Fax_Number").Equals(DBNull.Value) Then Fax_Number = dt.Rows(0).Item("Fax_Number") 'Char(25) NULL
                If Not dt.Rows(0).Item("Email_Address").Equals(DBNull.Value) Then Email_Address = dt.Rows(0).Item("Email_Address") 'Char(128) NULL
                If Not dt.Rows(0).Item("Use_Email").Equals(DBNull.Value) Then Use_Email = dt.Rows(0).Item("Use_Email") 'Char(10) NULL
                If Not dt.Rows(0).Item("Ship_To_City").Equals(DBNull.Value) Then Ship_To_City = dt.Rows(0).Item("Ship_To_City") 'Char(100) NULL     Ship To City
                If Not dt.Rows(0).Item("Ship_To_State").Equals(DBNull.Value) Then Ship_To_State = dt.Rows(0).Item("Ship_To_State") 'Char(3) NULL       Ship To State
                If Not dt.Rows(0).Item("Ship_To_Zip").Equals(DBNull.Value) Then Ship_To_Zip = dt.Rows(0).Item("Ship_To_Zip") 'Char(20) NULL      Ship To Zip Code
                If Not dt.Rows(0).Item("Bill_To_City").Equals(DBNull.Value) Then Bill_To_City = dt.Rows(0).Item("Bill_To_City") 'Char(100) NULL
                If Not dt.Rows(0).Item("Bill_To_State").Equals(DBNull.Value) Then Bill_To_State = dt.Rows(0).Item("Bill_To_State") 'Char(3) NULL
                If Not dt.Rows(0).Item("Bill_To_Zip").Equals(DBNull.Value) Then Bill_To_Zip = dt.Rows(0).Item("Bill_To_Zip") 'Char(20) NULL
                If Not dt.Rows(0).Item("Filler_0001").Equals(DBNull.Value) Then Filler_0001 = dt.Rows(0).Item("Filler_0001") 'Char(146) NULL
                If Not dt.Rows(0).Item("Id").Equals(DBNull.Value) Then Id = dt.Rows(0).Item("Id") 'Numeric NOT NULL
                If Not dt.Rows(0).Item("In_Hands_Date").Equals(DBNull.Value) Then in_hands_date = dt.Rows(0).Item("In_Hands_Date") 'DateTime NULL

                poDescriptions.Location = dt.Rows(0).Item("Location_Desc")
                poDescriptions.ShipVia = dt.Rows(0).Item("Ship_Via_Desc")
                poDescriptions.Status = dt.Rows(0).Item("Status_Desc")
                poDescriptions.Terms = dt.Rows(0).Item("Term_Description")
                poDescriptions.TaxSchedule = dt.Rows(0).Item("Tax_Sched_Desc")
                poDescriptions.TaxCode1 = dt.Rows(0).Item("Tax_Cd_1_Desc")
                poDescriptions.TaxCode2 = dt.Rows(0).Item("Tax_Cd_2_Desc")
                poDescriptions.TaxCode3 = dt.Rows(0).Item("Tax_Cd_3_Desc")
                poDescriptions.Salesperson1 = dt.Rows(0).Item("Slspsn_1_Desc")
                poDescriptions.Salesperson2 = dt.Rows(0).Item("Slspsn_2_Desc")
                poDescriptions.Salesperson3 = dt.Rows(0).Item("Slspsn_3_Desc")

                '' These fields are to be included in the export file, have to find where to link.
                'Public Ship_To_Contact As String 'Char(40)   Ship To Contact Name
                'Public Ship_To_Phone As String 'Char(20)   Ship To Phone No
                'Public Ship_To_Email As String 'Char(40)   Ship To email address
                'Public Ship_To_Terr As String 'Char(2)    Ship To Territory
                'Public Ship_To_UDF1 As String 'Char(30)   Ship To User Defined Field 1
                'Public Ship_To_UDF2 As String 'Char(30)   Ship To User Defined Field 2

            End If

        Catch er As Exception
            MsgBox("Error in COrdhead." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub LoadOEInterfaceOrder(ByRef poDescriptions As cOrdheadDesc)

        Dim dt As DataTable
        Dim db As New cDBA

        Dim strSql As String

        Try
            'strSql = "" & _
            '"SELECT TOP 1 * " & _
            '"FROM   OEI_ORDHDR WITH (Nolock) " & _
            '"WHERE  Ord_Guid = '" & Ord_GUID & "' "

            strSql = "" & _
            "SELECT TOP 1 * " & _
            "FROM   OEI_ORDHDR WITH (Nolock) " & _
            "WHERE  OEI_Ord_No = '" & OEI_Ord_No & "' "

            Header_Table = "OEI_ORDHDR"

            strSql = "" & _
            "SELECT 	TOP 1 O.*, ISNULL(Loc.Loc_Desc, '') as Location_Desc, ISNULL(SV.Code_Desc, '') AS Ship_Via_Desc, " & _
            "			CASE ISNULL(O.Status, '') WHEN '1' THEN 'Booked Order' WHEN '4' THEN 'Pick Ticket printed' WHEN '8' THEN 'Selected for Billing' WHEN 'C' THEN 'Credit Hold' WHEN 'I' THEN 'Order Incomplete' WHEN 'L' THEN 'Closed Order' ELSE 'Unknown' END AS Status_Desc, " & _
            "			ISNULL(Term.Description, '') as Term_Description, ISNULL(TS.Tax_Sched_Desc, '') AS Tax_Sched_Desc, " & _
            "			ISNULL(Tax1.Tax_Cd_Description, '') as Tax_Cd_1_Desc, " & _
            "			ISNULL(Tax2.Tax_Cd_Description, '') as Tax_Cd_2_Desc, " & _
            "			ISNULL(Tax3.Tax_Cd_Description, '') as Tax_Cd_3_Desc, " & _
            "			ISNULL(Sls1.Slspsn_Name, '') AS Slspsn_1_Desc, " & _
            "			ISNULL(Sls2.Slspsn_Name, '') AS Slspsn_2_Desc, " & _
            "			ISNULL(Sls3.Slspsn_Name, '') AS Slspsn_3_Desc, " & _
            "           ISNULL(P.Spector_Cd, '') AS Prog_Spector_Cd " & _
            "FROM   	OEI_ORDHDR O      WITH (Nolock)  " & _
            "LEFT JOIN	IMLOCFIL_SQL Loc  WITH (Nolock) ON O.Mfg_Loc = Loc.Loc " & _
            "LEFT JOIN 	SYCDEFIL_SQL SV   WITH (Nolock) ON O.Ship_Via_Cd = SV.sy_code " & _
            "LEFT JOIN 	SYTRMFIL_SQL Term WITH (Nolock) ON O.AR_Terms_CD = Term.Term_Code " & _
            "LEFT JOIN 	TAXSCHED_SQL TS   WITH (Nolock) ON O.Tax_Sched = TS.Tax_Sched " & _
            "LEFT JOIN 	TAXDETL_SQL  Tax1 WITH (Nolock) ON O.Tax_Cd = Tax1.Tax_Cd " & _
            "LEFT JOIN 	TAXDETL_SQL  Tax2 WITH (Nolock) ON O.Tax_Cd_2 = Tax2.Tax_Cd " & _
            "LEFT JOIN 	TAXDETL_SQL  Tax3 WITH (Nolock) ON O.Tax_Cd_3 = Tax3.Tax_Cd " & _
            "LEFT JOIN 	ARSLMFIL_SQL Sls1 WITH (Nolock) ON O.SlsPsn_No = Sls1.HumRes_ID " & _
            "LEFT JOIN 	ARSLMFIL_SQL Sls2 WITH (Nolock) ON O.SlsPsn_No_2 = Sls2.HumRes_ID " & _
            "LEFT JOIN 	ARSLMFIL_SQL Sls3 WITH (Nolock) ON O.SlsPsn_No_3 = Sls3.HumRes_ID " & _
            "LEFT JOIN 	MDB_CUS_PROG P    WITH (Nolock) ON O.Cus_Prog_ID = P.Cus_Prog_ID " & _
            "WHERE      LTRIM(RTRIM(O.OEI_Ord_No)) = '" & OEI_Ord_No & "' "

            '"WHERE  Ord_Guid = '" & Ord_GUID & "' "

            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then
                If Not dt.Rows(0).Item("Cus_No").Equals(DBNull.Value) Then Cus_No = dt.Rows(0).Item("Cus_No") 'Char(20) NOT NULL Customer No     Always '999999999999'

                ' We get customer default values, then overwrite with order values
                Dim blnGo As Boolean = GetCustomerDefaultValues()

                If Not dt.Rows(0).Item("Ord_Type").Equals(DBNull.Value) Then
                    Select Case dt.Rows(0).Item("Ord_Type")
                        Case "I"
                            Ord_Type = "I=Invoice"
                        Case "O"
                            Ord_Type = "O=Order"
                    End Select
                End If

                If Not dt.Rows(0).Item("Ord_Guid").Equals(DBNull.Value) Then Ord_GUID = dt.Rows(0).Item("Ord_Guid") 'Char (8)       Order No    Order No from Website for tracking
                If Not dt.Rows(0).Item("Ord_No").Equals(DBNull.Value) Then Ord_No = dt.Rows(0).Item("Ord_No") 'Char (8)       Order No    Order No from Website for tracking
                If Not dt.Rows(0).Item("Status").Equals(DBNull.Value) Then Status = dt.Rows(0).Item("Status") 'Char(1) NOT NULL
                If Not dt.Rows(0).Item("Entered_Dt").Equals(DBNull.Value) Then Entered_Dt = dt.Rows(0).Item("Entered_Dt") 'DateTime NULL
                If Not dt.Rows(0).Item("Ord_Dt").Equals(DBNull.Value) Then Ord_Dt = dt.Rows(0).Item("Ord_Dt") 'Datetime       Order date
                If Not dt.Rows(0).Item("Apply_To_No").Equals(DBNull.Value) Then Apply_To_No = dt.Rows(0).Item("Apply_To_No") 'Char(8) NULL
                If Not dt.Rows(0).Item("Oe_Po_No").Equals(DBNull.Value) Then Oe_Po_No = dt.Rows(0).Item("Oe_Po_No") 'Char(25) NOT NULL Customer PO No
                If Not dt.Rows(0).Item("Cus_No").Equals(DBNull.Value) Then Cus_No = dt.Rows(0).Item("Cus_No") 'Char(20) NOT NULL Customer No     Always '999999999999'
                If Not dt.Rows(0).Item("Bal_Meth").Equals(DBNull.Value) Then Bal_Meth = dt.Rows(0).Item("Bal_Meth") 'Char(1) NULL
                If Not dt.Rows(0).Item("Bill_To_Name").Equals(DBNull.Value) Then Bill_To_Name = dt.Rows(0).Item("Bill_To_Name") 'Char (40)      Customer Name
                If Not dt.Rows(0).Item("Bill_To_Addr_1").Equals(DBNull.Value) Then Bill_To_Addr_1 = dt.Rows(0).Item("Bill_To_Addr_1") 'Char (40)      1st address line
                If Not dt.Rows(0).Item("Bill_To_Addr_2").Equals(DBNull.Value) Then Bill_To_Addr_2 = dt.Rows(0).Item("Bill_To_Addr_2") 'Char (40)      2nd address line
                If Not dt.Rows(0).Item("Bill_To_Addr_3").Equals(DBNull.Value) Then Bill_To_Addr_3 = dt.Rows(0).Item("Bill_To_Addr_3") 'Char (40)      3rd address line
                If Not dt.Rows(0).Item("Bill_To_Addr_4").Equals(DBNull.Value) Then Bill_To_Addr_4 = dt.Rows(0).Item("Bill_To_Addr_4") 'Char(44) NULL
                If Not dt.Rows(0).Item("Bill_To_Country").Equals(DBNull.Value) Then Bill_To_Country = dt.Rows(0).Item("Bill_To_Country") 'Char(3) NULL
                If Not dt.Rows(0).Item("Cus_Alt_Adr_Cd").Equals(DBNull.Value) Then m_Cus_Alt_Adr_Cd = dt.Rows(0).Item("Cus_Alt_Adr_Cd") 'Char(15) NULL
                If Not dt.Rows(0).Item("Ship_To_Name").Equals(DBNull.Value) Then Ship_To_Name = dt.Rows(0).Item("Ship_To_Name") 'Char(40) NULL  Ship To name
                If Not dt.Rows(0).Item("Ship_To_Addr_1").Equals(DBNull.Value) Then Ship_To_Addr_1 = dt.Rows(0).Item("Ship_To_Addr_1") 'Char(40) NULL  1st Ship To address line
                If Not dt.Rows(0).Item("Ship_To_Addr_2").Equals(DBNull.Value) Then Ship_To_Addr_2 = dt.Rows(0).Item("Ship_To_Addr_2") 'Char(40) NULL  2nd Ship To address line
                If Not dt.Rows(0).Item("Ship_To_Addr_3").Equals(DBNull.Value) Then Ship_To_Addr_3 = dt.Rows(0).Item("Ship_To_Addr_3") 'Char(40) NULL
                If Not dt.Rows(0).Item("Ship_To_Addr_4").Equals(DBNull.Value) Then Ship_To_Addr_4 = dt.Rows(0).Item("Ship_To_Addr_4") 'Char(44) NULL
                If Not dt.Rows(0).Item("Ship_To_Country").Equals(DBNull.Value) Then Ship_To_Country = dt.Rows(0).Item("Ship_To_Country") 'Char(3) NULL   Ship To Country
                If Not dt.Rows(0).Item("Shipping_Dt").Equals(DBNull.Value) Then Shipping_Dt = dt.Rows(0).Item("Shipping_Dt") 'DateTime NULL  Requested Ship Date
                If Not dt.Rows(0).Item("Shipping_Dt").Equals(DBNull.Value) Then Ord_Dt_Shipped = dt.Rows(0).Item("Shipping_Dt") 'DateTime NULL  Requested Ship Date
                If Not dt.Rows(0).Item("Ship_Via_Cd").Equals(DBNull.Value) Then Ship_Via_Cd = dt.Rows(0).Item("Ship_Via_Cd") 'Char(3) NULL   Ship Via Code
                If Not dt.Rows(0).Item("Ar_Terms_Cd").Equals(DBNull.Value) Then Ar_Terms_Cd = dt.Rows(0).Item("Ar_Terms_Cd") 'Char(2) NULL
                If Not dt.Rows(0).Item("Frt_Pay_Cd").Equals(DBNull.Value) Then Frt_Pay_Cd = dt.Rows(0).Item("Frt_Pay_Cd") 'Char(1) NULL
                If Not dt.Rows(0).Item("Ship_Instruction_1").Equals(DBNull.Value) Then Ship_Instruction_1 = dt.Rows(0).Item("Ship_Instruction_1") 'Char(40) NULL Shipping Instructions Line 1
                If Not dt.Rows(0).Item("Ship_Instruction_2").Equals(DBNull.Value) Then Ship_Instruction_2 = dt.Rows(0).Item("Ship_Instruction_2") 'Char(40) NULL Shipping Instructions Line 2
                If Not dt.Rows(0).Item("Slspsn_No").Equals(DBNull.Value) Then Slspsn_No = dt.Rows(0).Item("Slspsn_No") 'Int NOT NULL
                If Not dt.Rows(0).Item("Slspsn_Pct_Comm").Equals(DBNull.Value) Then Slspsn_Pct_Comm = dt.Rows(0).Item("Slspsn_Pct_Comm") 'Numeric NULL
                If Not dt.Rows(0).Item("Slspsn_Comm_Amt").Equals(DBNull.Value) Then Slspsn_Comm_Amt = dt.Rows(0).Item("Slspsn_Comm_Amt") 'Numeric NULL
                If Not dt.Rows(0).Item("Slspsn_No_2").Equals(DBNull.Value) Then Slspsn_No_2 = dt.Rows(0).Item("Slspsn_No_2") 'Int NULL
                If Not dt.Rows(0).Item("Slspsn_Pct_Comm_2").Equals(DBNull.Value) Then Slspsn_Pct_Comm_2 = dt.Rows(0).Item("Slspsn_Pct_Comm_2") 'Numeric NULL
                If Not dt.Rows(0).Item("Slspsn_Comm_Amt_2").Equals(DBNull.Value) Then Slspsn_Comm_Amt_2 = dt.Rows(0).Item("Slspsn_Comm_Amt_2") 'Numeric NULL
                If Not dt.Rows(0).Item("Slspsn_No_3").Equals(DBNull.Value) Then Slspsn_No_3 = dt.Rows(0).Item("Slspsn_No_3") 'Int NULL
                If Not dt.Rows(0).Item("Slspsn_Pct_Comm_3").Equals(DBNull.Value) Then Slspsn_Pct_Comm_3 = dt.Rows(0).Item("Slspsn_Pct_Comm_3") 'Numeric NULL
                If Not dt.Rows(0).Item("Slspsn_Comm_Amt_3").Equals(DBNull.Value) Then Slspsn_Comm_Amt_3 = dt.Rows(0).Item("Slspsn_Comm_Amt_3") 'Numeric NULL
                If Not dt.Rows(0).Item("Tax_Cd").Equals(DBNull.Value) Then Tax_Cd = dt.Rows(0).Item("Tax_Cd") 'Char(3) NULL   Tax Code    Applies to this order only
                If Not dt.Rows(0).Item("Tax_Pct").Equals(DBNull.Value) Then Tax_Pct = dt.Rows(0).Item("Tax_Pct") 'Numeric NULL
                If Not dt.Rows(0).Item("Tax_Cd_2").Equals(DBNull.Value) Then Tax_Cd_2 = dt.Rows(0).Item("Tax_Cd_2") 'Char(3) NULL
                If Not dt.Rows(0).Item("Tax_Pct_2").Equals(DBNull.Value) Then Tax_Pct_2 = dt.Rows(0).Item("Tax_Pct_2") 'Numeric NULL
                If Not dt.Rows(0).Item("Tax_Cd_3").Equals(DBNull.Value) Then Tax_Cd_3 = dt.Rows(0).Item("Tax_Cd_3") 'Char(3) NULL
                If Not dt.Rows(0).Item("Tax_Pct_3").Equals(DBNull.Value) Then Tax_Pct_3 = dt.Rows(0).Item("Tax_Pct_3") 'Numeric NULL
                If Not dt.Rows(0).Item("Discount_Pct").Equals(DBNull.Value) Then Discount_Pct = dt.Rows(0).Item("Discount_Pct") 'Numeric NULL
                If Not dt.Rows(0).Item("Job_No").Equals(DBNull.Value) Then Job_No = dt.Rows(0).Item("Job_No") 'Char(20) NULL
                If Not dt.Rows(0).Item("Mfg_Loc").Equals(DBNull.Value) Then Mfg_Loc = dt.Rows(0).Item("Mfg_Loc") 'Char(3) NULL
                If Not dt.Rows(0).Item("Profit_Center").Equals(DBNull.Value) Then Profit_Center = dt.Rows(0).Item("Profit_Center") 'Char(8) NULL
                If Not dt.Rows(0).Item("Dept").Equals(DBNull.Value) Then Dept = dt.Rows(0).Item("Dept") 'Char(8) NULL
                If Not dt.Rows(0).Item("Ar_Reference").Equals(DBNull.Value) Then Ar_Reference = dt.Rows(0).Item("Ar_Reference") 'Char(45) NULL
                If Not dt.Rows(0).Item("Tot_Sls_Amt").Equals(DBNull.Value) Then Tot_Sls_Amt = dt.Rows(0).Item("Tot_Sls_Amt") 'Numeric NULL
                If Not dt.Rows(0).Item("Tot_Sls_Disc").Equals(DBNull.Value) Then Tot_Sls_Disc = dt.Rows(0).Item("Tot_Sls_Disc") 'Numeric NULL
                If Not dt.Rows(0).Item("Tot_Tax_Amt").Equals(DBNull.Value) Then Tot_Tax_Amt = dt.Rows(0).Item("Tot_Tax_Amt") 'Numeric NULL
                If Not dt.Rows(0).Item("Tot_Cost").Equals(DBNull.Value) Then Tot_Cost = dt.Rows(0).Item("Tot_Cost") 'Numeric NULL
                If Not dt.Rows(0).Item("Tot_Weight").Equals(DBNull.Value) Then Tot_Weight = dt.Rows(0).Item("Tot_Weight") 'Numeric NULL
                If Not dt.Rows(0).Item("Misc_Amt").Equals(DBNull.Value) Then Misc_Amt = dt.Rows(0).Item("Misc_Amt") 'Numeric(16,2)  NULL   Miscellaneous Charge    Handling charge
                If Not dt.Rows(0).Item("Misc_Mn_No").Equals(DBNull.Value) Then Misc_Mn_No = dt.Rows(0).Item("Misc_Mn_No") 'Char(9) NULL
                If Not dt.Rows(0).Item("Misc_Sb_No").Equals(DBNull.Value) Then Misc_Sb_No = dt.Rows(0).Item("Misc_Sb_No") 'Char(8) NULL
                If Not dt.Rows(0).Item("Misc_Dp_No").Equals(DBNull.Value) Then Misc_Dp_No = dt.Rows(0).Item("Misc_Dp_No") 'Char(8) NULL
                If Not dt.Rows(0).Item("Frt_Amt").Equals(DBNull.Value) Then Frt_Amt = dt.Rows(0).Item("Frt_Amt") 'Numeric(16,2)  NULL   Freight Charge
                If Not dt.Rows(0).Item("Frt_Mn_No").Equals(DBNull.Value) Then Frt_Mn_No = dt.Rows(0).Item("Frt_Mn_No") 'Char(9) NULL
                If Not dt.Rows(0).Item("Frt_Sb_No").Equals(DBNull.Value) Then Frt_Sb_No = dt.Rows(0).Item("Frt_Sb_No") 'Char(8) NULL
                If Not dt.Rows(0).Item("Frt_Dp_No").Equals(DBNull.Value) Then Frt_Dp_No = dt.Rows(0).Item("Frt_Dp_No") 'Char(8) NULL
                If Not dt.Rows(0).Item("Sls_Tax_Amt_1").Equals(DBNull.Value) Then Sls_Tax_Amt_1 = dt.Rows(0).Item("Sls_Tax_Amt_1") 'Numeric NULL
                If Not dt.Rows(0).Item("Sls_Tax_Amt_2").Equals(DBNull.Value) Then Sls_Tax_Amt_2 = dt.Rows(0).Item("Sls_Tax_Amt_2") 'Numeric NULL
                If Not dt.Rows(0).Item("Sls_Tax_Amt_3").Equals(DBNull.Value) Then Sls_Tax_Amt_3 = dt.Rows(0).Item("Sls_Tax_Amt_3") 'Numeric NULL
                If Not dt.Rows(0).Item("Comm_Pct").Equals(DBNull.Value) Then Comm_Pct = dt.Rows(0).Item("Comm_Pct") 'Numeric NULL
                If Not dt.Rows(0).Item("Comm_Amt").Equals(DBNull.Value) Then Comm_Amt = dt.Rows(0).Item("Comm_Amt") 'Numeric NULL
                If Not dt.Rows(0).Item("Cmt_1").Equals(DBNull.Value) Then Cmt_1 = dt.Rows(0).Item("Cmt_1") 'Char(35) NULL
                If Not dt.Rows(0).Item("Cmt_2").Equals(DBNull.Value) Then Cmt_2 = dt.Rows(0).Item("Cmt_2") 'Char(35) NULL
                If Not dt.Rows(0).Item("Cmt_3").Equals(DBNull.Value) Then Cmt_3 = dt.Rows(0).Item("Cmt_3") 'Char(35) NULL
                If Not dt.Rows(0).Item("Payment_Amt").Equals(DBNull.Value) Then Payment_Amt = dt.Rows(0).Item("Payment_Amt") 'Numeric NULL
                If Not dt.Rows(0).Item("Payment_Disc_Amt").Equals(DBNull.Value) Then Payment_Disc_Amt = dt.Rows(0).Item("Payment_Disc_Amt") 'Numeric NULL
                If Not dt.Rows(0).Item("Chk_No").Equals(DBNull.Value) Then Chk_No = dt.Rows(0).Item("Chk_No") 'Char(8) NULL
                If Not dt.Rows(0).Item("Chk_Dt").Equals(DBNull.Value) Then Chk_Dt = dt.Rows(0).Item("Chk_Dt") 'DateTime NULL
                If Not dt.Rows(0).Item("Cash_Mn_No").Equals(DBNull.Value) Then Cash_Mn_No = dt.Rows(0).Item("Cash_Mn_No") 'Char(9) NULL
                If Not dt.Rows(0).Item("Cash_Sb_No").Equals(DBNull.Value) Then Cash_Sb_No = dt.Rows(0).Item("Cash_Sb_No") 'Char(8) NULL
                If Not dt.Rows(0).Item("Cash_Dp_No").Equals(DBNull.Value) Then Cash_Dp_No = dt.Rows(0).Item("Cash_Dp_No") 'Char(8) NULL
                If Not dt.Rows(0).Item("Ord_Dt_Picked").Equals(DBNull.Value) Then Ord_Dt_Picked = dt.Rows(0).Item("Ord_Dt_Picked") 'DateTime NULL
                If Not dt.Rows(0).Item("Ord_Dt_Billed").Equals(DBNull.Value) Then Ord_Dt_Billed = dt.Rows(0).Item("Ord_Dt_Billed") 'DateTime NULL
                If Not dt.Rows(0).Item("Inv_No").Equals(DBNull.Value) Then Inv_No = dt.Rows(0).Item("Inv_No") 'Char(8) NOT NULL
                If Not dt.Rows(0).Item("Inv_Dt").Equals(DBNull.Value) Then Inv_Dt = dt.Rows(0).Item("Inv_Dt") 'DateTime NULL
                If Not dt.Rows(0).Item("Selection_Cd").Equals(DBNull.Value) Then Selection_Cd = dt.Rows(0).Item("Selection_Cd") 'Char(1) NOT NULL
                If Not dt.Rows(0).Item("Posted_Dt").Equals(DBNull.Value) Then Posted_Dt = dt.Rows(0).Item("Posted_Dt") 'DateTime NULL
                If Not dt.Rows(0).Item("Part_Posted_Fg").Equals(DBNull.Value) Then Part_Posted_Fg = dt.Rows(0).Item("Part_Posted_Fg") 'Char(1) NULL
                If Not dt.Rows(0).Item("Ship_To_Freefrm_Fg").Equals(DBNull.Value) Then Ship_To_Freefrm_Fg = dt.Rows(0).Item("Ship_To_Freefrm_Fg") 'Char(1) NULL
                If Not dt.Rows(0).Item("Bill_To_Freefrm_Fg").Equals(DBNull.Value) Then Bill_To_Freefrm_Fg = dt.Rows(0).Item("Bill_To_Freefrm_Fg") 'Char(1) NULL
                If Not dt.Rows(0).Item("Copy_To_Bm_Fg").Equals(DBNull.Value) Then Copy_To_Bm_Fg = dt.Rows(0).Item("Copy_To_Bm_Fg") 'Char(1) NULL
                If Not dt.Rows(0).Item("Edi_Fg").Equals(DBNull.Value) Then Edi_Fg = dt.Rows(0).Item("Edi_Fg") 'Char(1) NULL
                If Not dt.Rows(0).Item("Closed_Fg").Equals(DBNull.Value) Then Closed_Fg = dt.Rows(0).Item("Closed_Fg") 'Char(1) NULL
                If Not dt.Rows(0).Item("Accum_Misc_Amt").Equals(DBNull.Value) Then Accum_Misc_Amt = dt.Rows(0).Item("Accum_Misc_Amt") 'Numeric NULL
                If Not dt.Rows(0).Item("Accum_Frt_Amt").Equals(DBNull.Value) Then Accum_Frt_Amt = dt.Rows(0).Item("Accum_Frt_Amt") 'Numeric NULL
                If Not dt.Rows(0).Item("Accum_Tot_Tax_Amt").Equals(DBNull.Value) Then Accum_Tot_Tax_Amt = dt.Rows(0).Item("Accum_Tot_Tax_Amt") 'Numeric NULL
                If Not dt.Rows(0).Item("Accum_Sls_Tax_Amt").Equals(DBNull.Value) Then Accum_Sls_Tax_Amt = dt.Rows(0).Item("Accum_Sls_Tax_Amt") 'Numeric NULL
                If Not dt.Rows(0).Item("Accum_Tot_Sls_Amt").Equals(DBNull.Value) Then Accum_Tot_Sls_Amt = dt.Rows(0).Item("Accum_Tot_Sls_Amt") 'Numeric NULL
                If Not dt.Rows(0).Item("Hold_Fg").Equals(DBNull.Value) Then Hold_Fg = dt.Rows(0).Item("Hold_Fg") 'Char(1) NOT NULL
                If Not dt.Rows(0).Item("Prepayment_Fg").Equals(DBNull.Value) Then Prepayment_Fg = dt.Rows(0).Item("Prepayment_Fg") 'Char(1) NULL
                If Not dt.Rows(0).Item("Lost_Sale_Cd").Equals(DBNull.Value) Then Lost_Sale_Cd = dt.Rows(0).Item("Lost_Sale_Cd") 'Char(3) NULL
                If Not dt.Rows(0).Item("Orig_Ord_Type").Equals(DBNull.Value) Then Orig_Ord_Type = dt.Rows(0).Item("Orig_Ord_Type") 'Char(1) NULL
                If Not dt.Rows(0).Item("Orig_Ord_Dt").Equals(DBNull.Value) Then Orig_Ord_Dt = dt.Rows(0).Item("Orig_Ord_Dt") 'DateTime NULL
                If Not dt.Rows(0).Item("Orig_Ord_No").Equals(DBNull.Value) Then Orig_Ord_No = dt.Rows(0).Item("Orig_Ord_No") 'Char(8) NULL
                If Not dt.Rows(0).Item("Award_Probability").Equals(DBNull.Value) Then Award_Probability = dt.Rows(0).Item("Award_Probability") 'TinyInt NULL
                If Not dt.Rows(0).Item("Oe_Cash_No").Equals(DBNull.Value) Then Oe_Cash_No = dt.Rows(0).Item("Oe_Cash_No") 'Char(8) NOT NULL
                If Not dt.Rows(0).Item("Exch_Rt_Fg").Equals(DBNull.Value) Then Exch_Rt_Fg = dt.Rows(0).Item("Exch_Rt_Fg") 'Char(1) NULL
                If Not dt.Rows(0).Item("Curr_Cd").Equals(DBNull.Value) Then m_Curr_Cd = dt.Rows(0).Item("Curr_Cd") 'Char(3) NOT NULL
                If Not dt.Rows(0).Item("Orig_Trx_Rt").Equals(DBNull.Value) Then Orig_Trx_Rt = dt.Rows(0).Item("Orig_Trx_Rt") 'Numeric NULL
                If Not dt.Rows(0).Item("Curr_Trx_Rt").Equals(DBNull.Value) Then Curr_Trx_Rt = dt.Rows(0).Item("Curr_Trx_Rt") 'Numeric NULL
                If Not dt.Rows(0).Item("Tax_Sched").Equals(DBNull.Value) Then Tax_Sched = dt.Rows(0).Item("Tax_Sched") 'Char(5) NULL
                If Not dt.Rows(0).Item("User_Def_Fld_1").Equals(DBNull.Value) Then User_Def_Fld_1 = dt.Rows(0).Item("User_Def_Fld_1") 'Char(30) NULL
                If Not dt.Rows(0).Item("User_Def_Fld_2").Equals(DBNull.Value) Then User_Def_Fld_2 = dt.Rows(0).Item("User_Def_Fld_2") 'Char(30) NULL
                If Not dt.Rows(0).Item("User_Def_Fld_3").Equals(DBNull.Value) Then User_Def_Fld_3 = dt.Rows(0).Item("User_Def_Fld_3") 'Char(30) NULL
                If Not dt.Rows(0).Item("User_Def_Fld_4").Equals(DBNull.Value) Then User_Def_Fld_4 = dt.Rows(0).Item("User_Def_Fld_4") 'Char(30) NULL
                If Not dt.Rows(0).Item("User_Def_Fld_5").Equals(DBNull.Value) Then User_Def_Fld_5 = dt.Rows(0).Item("User_Def_Fld_5") 'Char(30) NULL
                If Not dt.Rows(0).Item("Deter_Rate_By").Equals(DBNull.Value) Then Deter_Rate_By = dt.Rows(0).Item("Deter_Rate_By") 'Char(1) NULL
                If Not dt.Rows(0).Item("Form_No").Equals(DBNull.Value) Then Form_No = dt.Rows(0).Item("Form_No") 'TinyInt NULL
                If Not dt.Rows(0).Item("Tax_Fg").Equals(DBNull.Value) Then Tax_Fg = dt.Rows(0).Item("Tax_Fg") 'Char(1) NULL  Taxable status  'Y'es or 'N'o - Applies to this order only
                If Not dt.Rows(0).Item("Sls_Mn_No").Equals(DBNull.Value) Then Sls_Mn_No = dt.Rows(0).Item("Sls_Mn_No") 'Char(9) NULL
                If Not dt.Rows(0).Item("Sls_Sb_No").Equals(DBNull.Value) Then Sls_Sb_No = dt.Rows(0).Item("Sls_Sb_No") 'Char(8) NULL
                If Not dt.Rows(0).Item("Sls_Dp_No").Equals(DBNull.Value) Then Sls_Dp_No = dt.Rows(0).Item("Sls_Dp_No") 'Char(8) NULL
                If Not dt.Rows(0).Item("Ord_Dt_Shipped").Equals(DBNull.Value) Then Ord_Dt_Shipped = dt.Rows(0).Item("Ord_Dt_Shipped") 'DateTime NULL
                If Not dt.Rows(0).Item("Tot_Dollars").Equals(DBNull.Value) Then Tot_Dollars = dt.Rows(0).Item("Tot_Dollars") 'Numeric NULL
                If Not dt.Rows(0).Item("Mult_Loc_Fg").Equals(DBNull.Value) Then Mult_Loc_Fg = dt.Rows(0).Item("Mult_Loc_Fg") 'Char(1) NULL
                If Not dt.Rows(0).Item("Tot_Tax_Cost").Equals(DBNull.Value) Then Tot_Tax_Cost = dt.Rows(0).Item("Tot_Tax_Cost") 'Numeric NULL
                If Not dt.Rows(0).Item("Hist_Load_Record").Equals(DBNull.Value) Then Hist_Load_Record = dt.Rows(0).Item("Hist_Load_Record") 'Char(1) NULL
                If Not dt.Rows(0).Item("Pre_Select_Status").Equals(DBNull.Value) Then Pre_Select_Status = dt.Rows(0).Item("Pre_Select_Status") 'Char(1) NULL
                If Not dt.Rows(0).Item("Packing_No").Equals(DBNull.Value) Then Packing_No = dt.Rows(0).Item("Packing_No") 'Int NULL
                If Not dt.Rows(0).Item("Deliv_Ar_Terms_Cd").Equals(DBNull.Value) Then Deliv_Ar_Terms_Cd = dt.Rows(0).Item("Deliv_Ar_Terms_Cd") 'Char(2) NULL
                If Not dt.Rows(0).Item("Inv_Batch_Id").Equals(DBNull.Value) Then Inv_Batch_Id = dt.Rows(0).Item("Inv_Batch_Id") 'Char(8) NULL
                If Not dt.Rows(0).Item("Bill_To_No").Equals(DBNull.Value) Then Bill_To_No = dt.Rows(0).Item("Bill_To_No") 'Char(20) NULL
                If Not dt.Rows(0).Item("Rma_No").Equals(DBNull.Value) Then Rma_No = dt.Rows(0).Item("Rma_No") 'Char(8) NOT NULL
                If Not dt.Rows(0).Item("Prog_Term_No").Equals(DBNull.Value) Then Prog_Term_No = dt.Rows(0).Item("Prog_Term_No") 'Int NULL
                If Not dt.Rows(0).Item("Extra_1").Equals(DBNull.Value) Then Extra_1 = dt.Rows(0).Item("Extra_1") 'Char(1) NULL
                If Not dt.Rows(0).Item("Extra_2").Equals(DBNull.Value) Then Extra_2 = dt.Rows(0).Item("Extra_2") 'Char(1) NULL
                If Not dt.Rows(0).Item("Extra_3").Equals(DBNull.Value) Then Extra_3 = dt.Rows(0).Item("Extra_3") 'Char(1) NULL
                If Not dt.Rows(0).Item("Extra_4").Equals(DBNull.Value) Then Extra_4 = dt.Rows(0).Item("Extra_4") 'Char(1) NULL
                If Not dt.Rows(0).Item("Extra_5").Equals(DBNull.Value) Then Extra_5 = dt.Rows(0).Item("Extra_5") 'Char(1) NULL
                If Not dt.Rows(0).Item("Extra_6").Equals(DBNull.Value) Then Extra_6 = dt.Rows(0).Item("Extra_6") 'Char(8) NULL
                If Not dt.Rows(0).Item("Extra_7").Equals(DBNull.Value) Then Extra_7 = dt.Rows(0).Item("Extra_7") 'Char(8) NULL
                If Not dt.Rows(0).Item("Extra_8").Equals(DBNull.Value) Then Extra_8 = dt.Rows(0).Item("Extra_8") 'Char(12) NULL
                If Not dt.Rows(0).Item("Extra_9").Equals(DBNull.Value) Then Extra_9 = dt.Rows(0).Item("Extra_9") 'Char(12) NULL
                If Not dt.Rows(0).Item("Extra_10").Equals(DBNull.Value) Then Extra_10 = dt.Rows(0).Item("Extra_10") 'Numeric NULL
                If Not dt.Rows(0).Item("Extra_11").Equals(DBNull.Value) Then Extra_11 = dt.Rows(0).Item("Extra_11") 'Numeric NULL
                If Not dt.Rows(0).Item("Extra_12").Equals(DBNull.Value) Then Extra_12 = dt.Rows(0).Item("Extra_12") 'Numeric NULL
                If Not dt.Rows(0).Item("Extra_13").Equals(DBNull.Value) Then Extra_13 = dt.Rows(0).Item("Extra_13") 'Numeric NULL
                If Not dt.Rows(0).Item("Extra_14").Equals(DBNull.Value) Then Extra_14 = dt.Rows(0).Item("Extra_14") 'Int NULL
                If Not dt.Rows(0).Item("Extra_15").Equals(DBNull.Value) Then Extra_15 = dt.Rows(0).Item("Extra_15") 'Int NULL
                If Not dt.Rows(0).Item("Edi_Doc_Seq").Equals(DBNull.Value) Then Edi_Doc_Seq = dt.Rows(0).Item("Edi_Doc_Seq") 'SmallInt NULL
                If Not dt.Rows(0).Item("Contact_1").Equals(DBNull.Value) Then Contact_1 = dt.Rows(0).Item("Contact_1") 'Char(100) NULL
                If Not dt.Rows(0).Item("Phone_Number").Equals(DBNull.Value) Then Phone_Number = dt.Rows(0).Item("Phone_Number") 'Char(25) NULL
                If Not dt.Rows(0).Item("Fax_Number").Equals(DBNull.Value) Then Fax_Number = dt.Rows(0).Item("Fax_Number") 'Char(25) NULL
                If Not dt.Rows(0).Item("Email_Address").Equals(DBNull.Value) Then Email_Address = dt.Rows(0).Item("Email_Address") 'Char(128) NULL
                If Not dt.Rows(0).Item("Use_Email").Equals(DBNull.Value) Then Use_Email = dt.Rows(0).Item("Use_Email") 'Char(10) NULL
                If Not dt.Rows(0).Item("Ship_To_City").Equals(DBNull.Value) Then Ship_To_City = dt.Rows(0).Item("Ship_To_City") 'Char(100) NULL     Ship To City
                If Not dt.Rows(0).Item("Ship_To_State").Equals(DBNull.Value) Then Ship_To_State = dt.Rows(0).Item("Ship_To_State") 'Char(3) NULL       Ship To State
                If Not dt.Rows(0).Item("Ship_To_Zip").Equals(DBNull.Value) Then Ship_To_Zip = dt.Rows(0).Item("Ship_To_Zip") 'Char(20) NULL      Ship To Zip Code
                If Not dt.Rows(0).Item("Bill_To_City").Equals(DBNull.Value) Then Bill_To_City = dt.Rows(0).Item("Bill_To_City") 'Char(100) NULL
                If Not dt.Rows(0).Item("Bill_To_State").Equals(DBNull.Value) Then Bill_To_State = dt.Rows(0).Item("Bill_To_State") 'Char(3) NULL
                If Not dt.Rows(0).Item("Bill_To_Zip").Equals(DBNull.Value) Then Bill_To_Zip = dt.Rows(0).Item("Bill_To_Zip") 'Char(20) NULL
                If Not dt.Rows(0).Item("Filler_0001").Equals(DBNull.Value) Then Filler_0001 = dt.Rows(0).Item("Filler_0001") 'Char(146) NULL
                If Not dt.Rows(0).Item("Id").Equals(DBNull.Value) Then Id = dt.Rows(0).Item("Id") 'Numeric NOT NULL
                If Not dt.Rows(0).Item("Ship_Link").Equals(DBNull.Value) Then Ship_Link = dt.Rows(0).Item("Ship_Link") 'Numeric NOT NULL
                If Not dt.Rows(0).Item("SendOrderAck").Equals(DBNull.Value) Then SendOrderAck = dt.Rows(0).Item("SendOrderAck") 'Numeric NOT NULL
                If Not dt.Rows(0).Item("ContactView").Equals(DBNull.Value) Then ContactView = dt.Rows(0).Item("ContactView") 'Numeric NOT NULL

                If Not dt.Rows(0).Item("ExportTS").Equals(DBNull.Value) Then ExportTS = dt.Rows(0).Item("ExportTS") 'Numeric NOT NULL
                If Not dt.Rows(0).Item("MacolaTS").Equals(DBNull.Value) Then MacolaTS = dt.Rows(0).Item("MacolaTS") 'Numeric NOT NULL
                If Not dt.Rows(0).Item("OpenTS").Equals(DBNull.Value) Then OpenTS = dt.Rows(0).Item("OpenTS") 'Numeric NOT NULL
                If Not dt.Rows(0).Item("OrderAckSaveOnly").Equals(DBNull.Value) Then OrderAckSaveOnly = IIf(dt.Rows(0).Item("OrderAckSaveOnly"), 1, 0) 'Numeric NOT NULL
                If Not dt.Rows(0).Item("AutoCompleteReship").Equals(DBNull.Value) Then AutoCompleteReship = IIf(dt.Rows(0).Item("AutoCompleteReship"), 1, 0) 'Numeric NOT NULL

                If Not dt.Rows(0).Item("RepeatOrd_No").Equals(DBNull.Value) Then RepeatOrd_No = dt.Rows(0).Item("RepeatOrd_No") 'Char(20) NULL

                If Not dt.Rows(0).Item("Cus_Prog_ID").Equals(DBNull.Value) Then Cus_Prog_ID = dt.Rows(0).Item("Cus_Prog_ID") 'Char(20) NULL
                If Not dt.Rows(0).Item("Prog_Spector_Cd").Equals(DBNull.Value) Then m_strCus_Spector_Cd = dt.Rows(0).Item("Prog_Spector_Cd") 'Char(20) NULL

                If Not dt.Rows(0).Item("White_Glove").Equals(DBNull.Value) Then White_Glove = dt.Rows(0).Item("White_Glove")
                '++ID 07042024 new field ticket# 40254
                If Not dt.Rows(0).Item("SSP").Equals(DBNull.Value) Then intSSP = dt.Rows(0).Item("SSP")

                If Not dt.Rows(0).Item("In_Hands_Date").Equals(DBNull.Value) Then in_hands_date = dt.Rows(0).Item("In_Hands_Date")

                poDescriptions.Location = dt.Rows(0).Item("Location_Desc")
                poDescriptions.ShipVia = dt.Rows(0).Item("Ship_Via_Desc")
                poDescriptions.Status = dt.Rows(0).Item("Status_Desc")
                poDescriptions.Terms = dt.Rows(0).Item("Term_Description")
                poDescriptions.TaxSchedule = dt.Rows(0).Item("Tax_Sched_Desc")
                poDescriptions.TaxCode1 = dt.Rows(0).Item("Tax_Cd_1_Desc")
                poDescriptions.TaxCode2 = dt.Rows(0).Item("Tax_Cd_2_Desc")
                poDescriptions.TaxCode3 = dt.Rows(0).Item("Tax_Cd_3_Desc")
                poDescriptions.Salesperson1 = dt.Rows(0).Item("Slspsn_1_Desc")
                poDescriptions.Salesperson2 = dt.Rows(0).Item("Slspsn_2_Desc")
                poDescriptions.Salesperson3 = dt.Rows(0).Item("Slspsn_3_Desc")

                '' These fields are to be included in the export file, have to find where to link.
                'Public Ship_To_Contact As String 'Char(40)   Ship To Contact Name
                'Public Ship_To_Phone As String 'Char(20)   Ship To Phone No
                'Public Ship_To_Email As String 'Char(40)   Ship To email address
                'Public Ship_To_Terr As String 'Char(2)    Ship To Territory
                'Public Ship_To_UDF1 As String 'Char(30)   Ship To User Defined Field 1
                'Public Ship_To_UDF2 As String 'Char(30)   Ship To User Defined Field 2

                m_oOrder.OrderLinesLoad(OrderSourceEnum.OEInterface)

                Call SetOpenTS()

            End If

        Catch er As Exception
            MsgBox("Error in COrdhead." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Function GetCustomerDefaultValues() As Boolean

        GetCustomerDefaultValues = False

        Dim strSql As String
        Dim dt As DataTable
        Dim db As New cDBA

        Try
            'strSql = "" & _
            '"SELECT     TOP 1 CU.*, CN.FullName, CN.Cnt_F_Tel, CN.Cnt_F_Fax, CN.Cnt_Email, " & _
            '"           ISNULL(Tax1.Tax_Cd_Percent, 0) AS Tax_Cd_Percent_1, " & _
            '"           ISNULL(Tax2.Tax_Cd_Percent, 0) AS Tax_Cd_Percent_2, " & _
            '"           ISNULL(Tax3.Tax_Cd_Percent, 0) AS Tax_Cd_Percent_3, " & _
            '"           CM.Cmp_Acc_Man, CM.Cmp_Code, CM.Cmp_FAdd1, CM.Cmp_FAdd2, " & _
            '"           CM.Cmp_FAdd3, CM.Cmp_FCity, CM.Cmp_FCounty, CM.Cmp_FCtry, " & _
            '"           CM.Cmp_FPC, CM.Cmp_Name, CM.Cmp_WWN, ISNULL(S.Slspsn_Name, '') AS Slspsn_Name, " & _
            '"           ClassificationID " & _
            '"FROM       ArCusFil_Sql CU   WITH (Nolock) " & _
            '"INNER JOIN CiCmpy       CM   WITH (Nolock) ON CU.Cus_No = CM.cmp_code " & _
            '"LEFT JOIN  CiCntp       CN   WITH (Nolock) ON CM.Cnt_ID = CN.Cnt_ID " & _
            '"LEFT JOIN 	TAXDETL_SQL  Tax1 WITH (Nolock) ON CU.Tax_Cd = Tax1.Tax_Cd " & _
            '"LEFT JOIN 	TAXDETL_SQL  Tax2 WITH (Nolock) ON CU.Tax_Cd_2 = Tax2.Tax_Cd " & _
            '"LEFT JOIN 	TAXDETL_SQL  Tax3 WITH (Nolock) ON CU.Tax_Cd_3 = Tax3.Tax_Cd " & _
            '"LEFT JOIN  ARSLMFIL_SQL S WITH (Nolock) ON CU.Slspsn_No = S.HumRes_ID " & _
            '"WHERE      RTRIM(ISNULL(CU.Cus_No, '')) = '" & RTrim(Cus_No) & "' " & _
            '"ORDER BY   CU.Cus_No ASC "

            strSql = "" _
          & "SELECT     TOP 1 CM.*, CASE CM.IsTaxable WHEN 1 THEN 'Y' ELSE 'N' END AS 'IsCusTaxable', CN.FullName, CN.Cnt_F_Tel, CN.Cnt_F_Fax, CN.Cnt_Email, " _
           & "           ISNULL(Tax1.Tax_Cd_Percent, 0) AS Tax_Cd_Percent_1, " _
           & "           ISNULL(Tax2.Tax_Cd_Percent, 0) AS Tax_Cd_Percent_2, " _
            & "           ISNULL(Tax3.Tax_Cd_Percent, 0) AS Tax_Cd_Percent_3, " _
           & "           ISNULL(S.Slspsn_Name, '') AS Slspsn_Name, ISNULL(osi.SPEC_INSTRUCTION, '') as Specific_Instructions, " _
          & "           CASE  " _
          & "               WHEN opc.extra_5 = 1 THEN 'ALWAYS INCLUDE PAPERPROOF' " _
          & "               WHEN opc.extra_5 = 2 THEN 'NEVER INCLUDE PAPERPROOF' " _
          & "               ELSE '' " _
           & "           END AS paperproof_status, " _
          & "           CASE  " _
          & "               WHEN opc2.extra_5 = 1 THEN 'ALWAYS EXACT QUANTITY' " _
          & "               WHEN opc2.extra_5 = 2 THEN 'NEVER EXACT QUANTITY' " _
          & "               ELSE '' " _
          & "           END AS exact_status ," _
          & "           CASE  " _
         & "               WHEN opc3.extra_5 = 1 THEN 'ALWAYS BLIND PAPERPROOF' " _
          & "               WHEN opc3.extra_5 = 2 THEN 'NEVER BLIND PAPERPROOF' " _
          & "               ELSE '' " _
          & "           END AS blind_paperproof_status " _
          & "FROM       CiCmpy       CM   WITH (Nolock) " _
           & "LEFT JOIN  CiCmpy       B    WITH (Nolock) ON CM.InvoiceDebtor = b.debnr " _
           & "LEFT JOIN  CiCntp       CN   WITH (Nolock) ON CM.Cnt_ID = CN.Cnt_ID and ISNULL(CN.active_y,0) = 1 " _
         & "LEFT JOIN 	TAXDETL_SQL  Tax1 WITH (Nolock) ON CM.TaxCode = Tax1.Tax_Cd " _
           & "LEFT JOIN 	TAXDETL_SQL  Tax2 WITH (Nolock) ON CM.TaxCode2 = Tax2.Tax_Cd " _
          & "LEFT JOIN 	TAXDETL_SQL  Tax3 WITH (Nolock) ON CM.TaxCode3 = Tax3.Tax_Cd " _
          & "LEFT JOIN  ARSLMFIL_SQL S WITH (Nolock) ON CM.AccountEmployeeId = S.HumRes_ID " _
         & "LEFT JOIN OEI_ITEM_SPEC_INSTRUCTION as osi " _
        & "   ON osi.CUS_NO = CM.cmp_code AND osi.ITEM_NO IS NULL AND osi.DATA_FIELD = 'CUSTOMER_SPEC' AND osi.IS_GLOBAL = 1 AND  osi.SHOW_MESSAGEBOX = 1 " _
          & "LEFT JOIN oeprcfil_sql as opc " _
          & "   ON opc.cd_tp_1_cust_no = CM.cmp_code AND opc.cd_tp = 1 AND opc.cd_tp_1_item_no = '44PAPR00PRF' AND opc.extra_5 IS NOT NULL " _
         & "       AND opc.cd_tp_3_cust_type = '' AND opc.cd_tp_2_prod_cat = '' " _
           & "       AND ISNULL(opc.start_dt, GETDATE()) <= CAST(GETDATE() AS DATE) AND ISNULL(opc.end_dt, GETDATE()) >= CAST(GETDATE() AS DATE) " _
          & "LEFT JOIN oeprcfil_sql as opc2 " _
         & "   ON opc2.cd_tp_1_cust_no = CM.cmp_code AND opc2.cd_tp = 1 AND opc2.cd_tp_1_item_no = '44EXCT00QNT' AND opc2.extra_5 IS NOT NULL " _
          & "       AND opc2.cd_tp_3_cust_type = '' AND opc2.cd_tp_2_prod_cat = '' " _
          & "       AND ISNULL(opc2.start_dt, GETDATE()) <= CAST(GETDATE() AS DATE) AND ISNULL(opc2.end_dt, GETDATE()) >= CAST(GETDATE() AS DATE) " _
          & "LEFT JOIN oeprcfil_sql as opc3 " _
         & "   ON opc3.cd_tp_1_cust_no = CM.cmp_code AND opc3.cd_tp = 1 AND opc3.cd_tp_1_item_no = '44BLINDPAP0' AND opc3.extra_5 IS NOT NULL " _
          & "       AND opc3.cd_tp_3_cust_type = '' AND opc3.cd_tp_2_prod_cat = '' " _
          & "       AND ISNULL(opc3.start_dt, GETDATE()) <= CAST(GETDATE() AS DATE) AND ISNULL(opc3.end_dt, GETDATE()) >= CAST(GETDATE() AS DATE) " _
          & "WHERE      RTRIM(ISNULL(CM.Cmp_code, '')) = '" & RTrim(Cus_No) & "' AND CM.debcode IS NOT NULL " _
          & "ORDER BY   CM.Cmp_code ASC "

            ' Verifier ceux qui ont des Cus_Alt_Adr_Cd
            '++ID 11.13.2019 added criteria for exclude inactive contacts active_y = 1 
                dt = db.DataTable(strSql)
            If dt.Rows.Count <> 0 Then

                Curr_Cd = IIf(dt.Rows(0).Item("Currency").Equals(DBNull.Value), "", dt.Rows(0).Item("Currency"))
                Cus_Type_Cd = IIf(dt.Rows(0).Item("AccountTypeCode").Equals(DBNull.Value), "", dt.Rows(0).Item("AccountTypeCode"))
                Cus_Name = IIf(dt.Rows(0).Item("Cmp_Name").Equals(DBNull.Value), "", dt.Rows(0).Item("Cmp_Name"))

                'm_oOrder.Ordhead.Bill_To_City = IIf(dt.Rows(0).item("Bill_To_City")), "", dt.Rows(0).item("Bill_To_City"))
                'm_oOrder.Ordhead.Bill_To_State = IIf(dt.Rows(0).item("Bill_To_State")), "", dt.Rows(0).item("Bill_To_State"))
                'm_oOrder.Ordhead.Bill_To_Zip = IIf(dt.Rows(0).item("Bill_To_Zip")), "", dt.Rows(0).item("Bill_To_Zip"))

                Bill_To_City = IIf(dt.Rows(0).Item("cmp_fcity").Equals(DBNull.Value), "", dt.Rows(0).Item("cmp_fcity"))
                Bill_To_State = IIf(dt.Rows(0).Item("StateCode").Equals(DBNull.Value), "", dt.Rows(0).Item("StateCode"))
                Bill_To_Zip = IIf(dt.Rows(0).Item("cmp_fpc").Equals(DBNull.Value), "", dt.Rows(0).Item("cmp_fpc"))

                Form_No = IIf(dt.Rows(0).Item("DefaultInvoiceForm").Equals(DBNull.Value), 0, dt.Rows(0).Item("DefaultInvoiceForm"))

                Tax_Sched = IIf(dt.Rows(0).Item("TaxSchedule").Equals(DBNull.Value), "", dt.Rows(0).Item("TaxSchedule"))
                'Tax_Fg = IIf(dt.Rows(0).Item("IsTaxable").Equals(DBNull.Value), "N", IIf(Trim(dt.Rows(0).Item("IsTaxable").ToString) = "1", "Y", "N"))
                Tax_Fg = IIf(dt.Rows(0).Item("IsCusTaxable").Equals(DBNull.Value), "N", dt.Rows(0).Item("IsCusTaxable").ToString)

                Tax_Cd = IIf(dt.Rows(0).Item("TaxCode").Equals(DBNull.Value), "", dt.Rows(0).Item("TaxCode"))
                Tax_Pct = dt.Rows(0).Item("Tax_Cd_Percent_1")

                Tax_Cd_2 = IIf(dt.Rows(0).Item("TaxCode2").Equals(DBNull.Value), "", dt.Rows(0).Item("TaxCode2"))
                Tax_Pct_2 = dt.Rows(0).Item("Tax_Cd_Percent_2")

                Tax_Cd_3 = IIf(dt.Rows(0).Item("TaxCode3").Equals(DBNull.Value), "", dt.Rows(0).Item("TaxCode3"))
                Tax_Pct_3 = dt.Rows(0).Item("Tax_Cd_Percent_3")

                m_Cus_Alt_Adr_Cd = String.Empty

                Oe_Po_No = String.Empty
                Slspsn_No = IIf(dt.Rows(0).Item("SalesPersonNumber").Equals(System.DBNull.Value), 0, dt.Rows(0).Item("SalesPersonNumber"))
                Job_No = String.Empty

                If g_User.Group_Defaults.Contains("MFG_LOC") Then
                    Mfg_Loc = Trim(g_User.Group_Defaults.Item("MFG_LOC").ToString)
                Else
                    Mfg_Loc = IIf(dt.Rows(0).Item("LocationShort").Equals(DBNull.Value), "", dt.Rows(0).Item("LocationShort"))
                End If

                Ship_Via_Cd = IIf(dt.Rows(0).Item("ShipVia").Equals(DBNull.Value), "", dt.Rows(0).Item("ShipVia"))
                Ar_Terms_Cd = IIf(dt.Rows(0).Item("PaymentCondition").Equals(DBNull.Value), "", dt.Rows(0).Item("PaymentCondition"))
                Discount_Pct = IIf(dt.Rows(0).Item("Discount").Equals(DBNull.Value), "", dt.Rows(0).Item("Discount"))
                Profit_Center = GetProfit_Center(Cus_Type_Cd)

                Dept = String.Empty

                Ship_Instruction_1 = String.Empty
                Ship_Instruction_2 = String.Empty

                Bill_To_Name = IIf(dt.Rows(0).Item("cmp_name").Equals(DBNull.Value), "", dt.Rows(0).Item("cmp_name"))
                Bill_To_Addr_1 = IIf(dt.Rows(0).Item("cmp_fadd1").Equals(DBNull.Value), "", dt.Rows(0).Item("cmp_fadd1"))
                Bill_To_Addr_2 = IIf(dt.Rows(0).Item("cmp_fadd2").Equals(DBNull.Value), "", dt.Rows(0).Item("cmp_fadd2"))
                Bill_To_Addr_3 = IIf(dt.Rows(0).Item("cmp_fadd3").Equals(DBNull.Value), "", dt.Rows(0).Item("cmp_fadd3"))
                Bill_To_Addr_4 = IIf(dt.Rows(0).Item("cmp_fcity").Equals(DBNull.Value), "", dt.Rows(0).Item("cmp_fcity")) & ", " & _
                    IIf(dt.Rows(0).Item("StateCode").Equals(DBNull.Value), "", dt.Rows(0).Item("StateCode")) & " " & _
                    IIf(dt.Rows(0).Item("cmp_fpc").Equals(DBNull.Value), "", dt.Rows(0).Item("cmp_fpc"))
                Bill_To_City = IIf(dt.Rows(0).Item("cmp_fcity").Equals(DBNull.Value), "", dt.Rows(0).Item("cmp_fcity"))
                Bill_To_State = IIf(dt.Rows(0).Item("StateCode").Equals(DBNull.Value), "", dt.Rows(0).Item("StateCode"))
                Bill_To_Zip = IIf(dt.Rows(0).Item("cmp_fpc").Equals(DBNull.Value), "", dt.Rows(0).Item("cmp_fpc"))
                Bill_To_Country = IIf(dt.Rows(0).Item("cmp_fctry").Equals(DBNull.Value), "", dt.Rows(0).Item("cmp_fctry"))
                'txtBill_To_CountryDesc.Text = dt.Rows(0).item("").Value 'will populate on country change

                'm_oOrder.Ordhead.Curr_Cd = IIf(dt.Rows(0).item("Curr_Cd")), "", dt.Rows(0).item("Curr_Cd"))

                Ship_To_Name = String.Empty
                Ship_To_Addr_1 = String.Empty
                Ship_To_Addr_2 = String.Empty
                Ship_To_Addr_3 = String.Empty
                Ship_To_Addr_4 = String.Empty
                Ship_To_Country = String.Empty
                Ship_To_City = String.Empty
                Ship_To_State = String.Empty
                Ship_To_Zip = String.Empty

                'ShipToCountryDesc.Text = String.Empty

                Contact_1 = IIf(dt.Rows(0).Item("FullName").Equals(DBNull.Value), "", dt.Rows(0).Item("FullName"))
                Phone_Number = IIf(dt.Rows(0).Item("Cnt_F_Tel").Equals(DBNull.Value), "", dt.Rows(0).Item("Cnt_F_Tel"))
                Fax_Number = IIf(dt.Rows(0).Item("Cnt_F_Fax").Equals(DBNull.Value), "", dt.Rows(0).Item("Cnt_F_Fax"))
                Email_Address = IIf(dt.Rows(0).Item("Cnt_Email").Equals(DBNull.Value), "", dt.Rows(0).Item("Cnt_Email"))

                Slspsn_Pct_Comm = 100
                Slspsn_Pct_Comm_2 = 0
                Slspsn_Pct_Comm_3 = 0

                Deter_Rate_By = m_oOrder.GetOE_Ctl_OE_Exchange_Rate_Flag

                Cus_Email_Address = IIf(dt.Rows(0).Item("cmp_e_mail").Equals(DBNull.Value), 0, dt.Rows(0).Item("cmp_e_mail"))

                If g_User.Group_Defaults.Contains("USER_DEF_FLD_2") Then
                    User_Def_Fld_2 = Trim(g_User.Group_Defaults.Item("USER_DEF_FLD_2").ToString)
                Else
                    User_Def_Fld_2 = "CAD" ' IIf(dt.Rows(0).item("AccountTypeCode")), "", dt.Rows(0).item("AccountTypeCode"))
                End If

                User_Def_Fld_3 = IIf(dt.Rows(0).Item("Slspsn_Name").Equals(DBNull.Value), 0, dt.Rows(0).Item("Slspsn_Name"))

                Frt_Pay_Cd = "N"
                Selection_Cd = "C"
                Status = "1"
                Bal_Meth = "O"

                'Slspsn_Name 

                'm_oOrder.Ordhead.Ship_To_City = IIf(dt.Rows(0).item("Cnt_Email")), "", dt.Rows(0).item("Cnt_Email"))
                'm_oOrder.Ordhead.Ship_To_State = IIf(dt.Rows(0).item("Cnt_Email")), "", dt.Rows(0).item("Cnt_Email"))
                'm_oOrder.Ordhead.Ship_To_Zip = IIf(dt.Rows(0).item("Cnt_Email")), "", dt.Rows(0).item("Cnt_Email"))

                'Customer.Cmp_Acc_Man = IIf(dt.Rows(0).item("Cmp_Acc_Man")), 0, dt.Rows(0).item("Cmp_Acc_Man"))
                'Customer.Cmp_Code = IIf(dt.Rows(0).item("Cmp_Code")), "", dt.Rows(0).item("Cmp_Code"))
                'Customer.Cmp_FAdd1 = IIf(dt.Rows(0).item("Cmp_FAdd1")), "", dt.Rows(0).item("Cmp_FAdd1"))
                'Customer.Cmp_FAdd2 = IIf(dt.Rows(0).item("Cmp_FAdd2")), "", dt.Rows(0).item("Cmp_FAdd2"))
                'Customer.Cmp_FAdd3 = IIf(dt.Rows(0).item("Cmp_FAdd3")), "", dt.Rows(0).item("Cmp_FAdd3"))
                'Customer.Cmp_FCity = IIf(dt.Rows(0).item("Cmp_FCity")), "", dt.Rows(0).item("Cmp_FCity"))
                'Customer.Cmp_FCounty = IIf(dt.Rows(0).item("Cmp_FCounty")), "", dt.Rows(0).item("Cmp_FCounty"))
                'Customer.Cmp_FCtry = IIf(dt.Rows(0).item("Cmp_FCtry")), "", dt.Rows(0).item("Cmp_FCtry"))
                'Customer.Cmp_FPC = IIf(dt.Rows(0).item("Cmp_FPC")), "", dt.Rows(0).item("Cmp_FPC"))
                'Customer.Cmp_Name = IIf(dt.Rows(0).item("Cmp_Name")), "", dt.Rows(0).item("Cmp_Name"))
                'Customer.Cmp_WWN = IIf(dt.Rows(0).item("Cmp_WWN")), "", dt.Rows(0).item("Cmp_WWN"))
                Customer.ClassificationId = IIf(dt.Rows(0).Item("ClassificationID").Equals(DBNull.Value), "", dt.Rows(0).Item("ClassificationID").ToString)




                Customer.ID = IIf(dt.Rows(0).Item("ID").Equals(DBNull.Value), 0, dt.Rows(0).Item("ID"))
                Customer.cmp_wwn = IIf(dt.Rows(0).Item("cmp_wwn").Equals(DBNull.Value), "", dt.Rows(0).Item("cmp_wwn").ToString)
                Customer.cmp_code = IIf(dt.Rows(0).Item("cmp_code").Equals(DBNull.Value), "", dt.Rows(0).Item("cmp_code"))
                Customer.cnt_id = IIf(dt.Rows(0).Item("cnt_id").Equals(DBNull.Value), "", dt.Rows(0).Item("cnt_id").ToString)
                Customer.cmp_parent = IIf(dt.Rows(0).Item("cmp_parent").Equals(DBNull.Value), "", dt.Rows(0).Item("cmp_parent").ToString)
                Customer.cmp_name = IIf(dt.Rows(0).Item("cmp_name").Equals(DBNull.Value), "", dt.Rows(0).Item("cmp_name"))
                Customer.cmp_fadd1 = IIf(dt.Rows(0).Item("cmp_fadd1").Equals(DBNull.Value), "", dt.Rows(0).Item("cmp_fadd1"))
                Customer.cmp_fadd2 = IIf(dt.Rows(0).Item("cmp_fadd2").Equals(DBNull.Value), "", dt.Rows(0).Item("cmp_fadd2"))
                Customer.cmp_fadd3 = IIf(dt.Rows(0).Item("cmp_fadd3").Equals(DBNull.Value), "", dt.Rows(0).Item("cmp_fadd3"))
                Customer.cmp_fpc = IIf(dt.Rows(0).Item("cmp_fpc").Equals(DBNull.Value), "", dt.Rows(0).Item("cmp_fpc"))
                Customer.cmp_fcity = IIf(dt.Rows(0).Item("cmp_fcity").Equals(DBNull.Value), "", dt.Rows(0).Item("cmp_fcity"))
                Customer.cmp_fcounty = IIf(dt.Rows(0).Item("cmp_fcounty").Equals(DBNull.Value), "", dt.Rows(0).Item("cmp_fcounty"))
                Customer.StateCode = IIf(dt.Rows(0).Item("StateCode").Equals(DBNull.Value), "", dt.Rows(0).Item("StateCode"))
                Customer.cmp_fctry = IIf(dt.Rows(0).Item("cmp_fctry").Equals(DBNull.Value), "", dt.Rows(0).Item("cmp_fctry"))
                Customer.cmp_e_mail = IIf(dt.Rows(0).Item("cmp_e_mail").Equals(DBNull.Value), "", dt.Rows(0).Item("cmp_e_mail"))
                Customer.cmp_web = IIf(dt.Rows(0).Item("cmp_web").Equals(DBNull.Value), "", dt.Rows(0).Item("cmp_web"))
                Customer.cmp_fax = IIf(dt.Rows(0).Item("cmp_fax").Equals(DBNull.Value), "", dt.Rows(0).Item("cmp_fax"))
                Customer.cmp_tel = IIf(dt.Rows(0).Item("cmp_tel").Equals(DBNull.Value), "", dt.Rows(0).Item("cmp_tel"))
                Customer.sct_code = IIf(dt.Rows(0).Item("sct_code").Equals(DBNull.Value), "", dt.Rows(0).Item("sct_code"))
                Customer.SubSector = IIf(dt.Rows(0).Item("SubSector").Equals(DBNull.Value), "", dt.Rows(0).Item("SubSector"))
                Customer.siz_code = IIf(dt.Rows(0).Item("siz_code").Equals(DBNull.Value), "", dt.Rows(0).Item("siz_code"))
                If Not (dt.Rows(0).Item("cmp_date_cust").Equals(DBNull.Value)) Then Customer.cmp_date_cust = dt.Rows(0).Item("cmp_date_cust")
                Customer.cmp_note = IIf(dt.Rows(0).Item("cmp_note").Equals(DBNull.Value), "", dt.Rows(0).Item("cmp_note"))
                Customer.cmp_acc_man = IIf(dt.Rows(0).Item("cmp_acc_man").Equals(DBNull.Value), 0, dt.Rows(0).Item("cmp_acc_man"))
                Customer.cmp_reseller = IIf(dt.Rows(0).Item("cmp_reseller").Equals(DBNull.Value), "", dt.Rows(0).Item("cmp_reseller"))
                Customer.Administration = IIf(dt.Rows(0).Item("Administration").Equals(DBNull.Value), 0, dt.Rows(0).Item("Administration"))
                Customer.cmp_type = IIf(dt.Rows(0).Item("cmp_type").Equals(DBNull.Value), "", dt.Rows(0).Item("cmp_type"))
                Customer.cmp_status = IIf(dt.Rows(0).Item("cmp_status").Equals(DBNull.Value), "", dt.Rows(0).Item("cmp_status"))
                Customer.DivisionDebtorID = IIf(dt.Rows(0).Item("DivisionDebtorID").Equals(DBNull.Value), "", dt.Rows(0).Item("DivisionDebtorID").ToString)
                Customer.DivisionCreditorID = IIf(dt.Rows(0).Item("DivisionCreditorID").Equals(DBNull.Value), "", dt.Rows(0).Item("DivisionCreditorID").ToString)
                If Not (dt.Rows(0).Item("datefield1").Equals(DBNull.Value)) Then Customer.datefield1 = dt.Rows(0).Item("datefield1")
                If Not (dt.Rows(0).Item("datefield2").Equals(DBNull.Value)) Then Customer.datefield2 = dt.Rows(0).Item("datefield2")
                If Not (dt.Rows(0).Item("datefield3").Equals(DBNull.Value)) Then Customer.datefield3 = dt.Rows(0).Item("datefield3")
                If Not (dt.Rows(0).Item("datefield4").Equals(DBNull.Value)) Then Customer.datefield4 = dt.Rows(0).Item("datefield4")
                If Not (dt.Rows(0).Item("datefield5").Equals(DBNull.Value)) Then Customer.datefield5 = dt.Rows(0).Item("datefield5")
                Customer.numberfield1 = IIf(dt.Rows(0).Item("numberfield1").Equals(DBNull.Value), 0, dt.Rows(0).Item("numberfield1"))
                Customer.numberfield2 = IIf(dt.Rows(0).Item("numberfield2").Equals(DBNull.Value), 0, dt.Rows(0).Item("numberfield2"))
                Customer.numberfield3 = IIf(dt.Rows(0).Item("numberfield3").Equals(DBNull.Value), 0, dt.Rows(0).Item("numberfield3"))
                Customer.numberfield4 = IIf(dt.Rows(0).Item("numberfield4").Equals(DBNull.Value), 0, dt.Rows(0).Item("numberfield4"))
                Customer.numberfield5 = IIf(dt.Rows(0).Item("numberfield5").Equals(DBNull.Value), 0, dt.Rows(0).Item("numberfield5"))
                Customer.YesNofield1 = IIf(dt.Rows(0).Item("YesNofield1").Equals(DBNull.Value), 0, dt.Rows(0).Item("YesNofield1"))
                Customer.YesNofield2 = IIf(dt.Rows(0).Item("YesNofield2").Equals(DBNull.Value), 0, dt.Rows(0).Item("YesNofield2"))
                Customer.YesNofield3 = IIf(dt.Rows(0).Item("YesNofield3").Equals(DBNull.Value), 0, dt.Rows(0).Item("YesNofield3"))
                Customer.YesNofield4 = IIf(dt.Rows(0).Item("YesNofield4").Equals(DBNull.Value), 0, dt.Rows(0).Item("YesNofield4"))
                Customer.YesNofield5 = IIf(dt.Rows(0).Item("YesNofield5").Equals(DBNull.Value), 0, dt.Rows(0).Item("YesNofield5"))
                Customer.textfield1 = IIf(dt.Rows(0).Item("textfield1").Equals(DBNull.Value), "", dt.Rows(0).Item("textfield1"))
                Customer.textfield2 = IIf(dt.Rows(0).Item("textfield2").Equals(DBNull.Value), "", dt.Rows(0).Item("textfield2"))
                Customer.textfield3 = IIf(dt.Rows(0).Item("textfield3").Equals(DBNull.Value), "", dt.Rows(0).Item("textfield3"))
                Customer.textfield4 = IIf(dt.Rows(0).Item("textfield4").Equals(DBNull.Value), "", dt.Rows(0).Item("textfield4"))
                Customer.textfield5 = IIf(dt.Rows(0).Item("textfield5").Equals(DBNull.Value), "", dt.Rows(0).Item("textfield5"))
                Customer.cmp_rating = IIf(dt.Rows(0).Item("cmp_rating").Equals(DBNull.Value), 0, dt.Rows(0).Item("cmp_rating"))
                Customer.cmp_origin = IIf(dt.Rows(0).Item("cmp_origin").Equals(DBNull.Value), "", dt.Rows(0).Item("cmp_origin"))
                ' Customer.Logo = IIF(dt.Rows(0).item("Logo")), "", dt.Rows(0).item("Logo"))
                Customer.LogoFileName = IIf(dt.Rows(0).Item("LogoFileName").Equals(DBNull.Value), "", dt.Rows(0).Item("LogoFileName"))
                Customer.document_id = IIf(dt.Rows(0).Item("document_id").Equals(DBNull.Value), "", dt.Rows(0).Item("document_id"))
                Customer.ClassificationId = IIf(dt.Rows(0).Item("ClassificationId").Equals(DBNull.Value), "", dt.Rows(0).Item("ClassificationId"))
                If Not (dt.Rows(0).Item("type_since").Equals(DBNull.Value)) Then Customer.type_since = dt.Rows(0).Item("type_since")
                If Not (dt.Rows(0).Item("status_since").Equals(DBNull.Value)) Then Customer.status_since = dt.Rows(0).Item("status_since")
                If Not (dt.Rows(0).Item("WebAccessSince").Equals(DBNull.Value)) Then Customer.WebAccessSince = dt.Rows(0).Item("WebAccessSince")
                Customer.ProcessedByBackgroundJob = IIf(dt.Rows(0).Item("ProcessedByBackgroundJob").Equals(DBNull.Value), 0, dt.Rows(0).Item("ProcessedByBackgroundJob"))
                Customer.debnr = IIf(dt.Rows(0).Item("debnr").Equals(DBNull.Value), "", dt.Rows(0).Item("debnr"))
                Customer.crdnr = IIf(dt.Rows(0).Item("crdnr").Equals(DBNull.Value), "", dt.Rows(0).Item("crdnr"))
                Customer.debcode = IIf(dt.Rows(0).Item("debcode").Equals(DBNull.Value), "", dt.Rows(0).Item("debcode"))
                Customer.crdcode = IIf(dt.Rows(0).Item("crdcode").Equals(DBNull.Value), "", dt.Rows(0).Item("crdcode"))
                Customer.ASPServer = IIf(dt.Rows(0).Item("ASPServer").Equals(DBNull.Value), "", dt.Rows(0).Item("ASPServer"))
                Customer.ASPDatabase = IIf(dt.Rows(0).Item("ASPDatabase").Equals(DBNull.Value), "", dt.Rows(0).Item("ASPDatabase"))
                Customer.ASPWebServer = IIf(dt.Rows(0).Item("ASPWebServer").Equals(DBNull.Value), "", dt.Rows(0).Item("ASPWebServer"))
                Customer.ASPWebSite = IIf(dt.Rows(0).Item("ASPWebSite").Equals(DBNull.Value), "", dt.Rows(0).Item("ASPWebSite"))
                If Not (dt.Rows(0).Item("syscreated").Equals(DBNull.Value)) Then Customer.syscreated = dt.Rows(0).Item("syscreated")
                Customer.syscreator = IIf(dt.Rows(0).Item("syscreator").Equals(DBNull.Value), 0, dt.Rows(0).Item("syscreator"))
                If Not (dt.Rows(0).Item("sysmodified").Equals(DBNull.Value)) Then Customer.sysmodified = dt.Rows(0).Item("sysmodified")
                Customer.sysmodifier = IIf(dt.Rows(0).Item("sysmodifier").Equals(DBNull.Value), 0, dt.Rows(0).Item("sysmodifier"))
                Customer.sysguid = IIf(dt.Rows(0).Item("sysguid").Equals(DBNull.Value), "", dt.Rows(0).Item("sysguid").ToString)
                Customer.timestamp = IIf(dt.Rows(0).Item("timestamp").Equals(DBNull.Value), "", dt.Rows(0).Item("timestamp").ToString)
                Customer.SearchCode = IIf(dt.Rows(0).Item("SearchCode").Equals(DBNull.Value), "", dt.Rows(0).Item("SearchCode"))
                Customer.Telex = IIf(dt.Rows(0).Item("Telex").Equals(DBNull.Value), "", dt.Rows(0).Item("Telex"))
                Customer.PostBankNumber = IIf(dt.Rows(0).Item("PostBankNumber").Equals(DBNull.Value), "", dt.Rows(0).Item("PostBankNumber"))
                Customer.NoteID = IIf(dt.Rows(0).Item("NoteID").Equals(DBNull.Value), 0, dt.Rows(0).Item("NoteID"))
                Customer.Blocked = IIf(dt.Rows(0).Item("Blocked").Equals(DBNull.Value), 0, dt.Rows(0).Item("Blocked"))
                Customer.LayoutCode = IIf(dt.Rows(0).Item("LayoutCode").Equals(DBNull.Value), "", dt.Rows(0).Item("LayoutCode"))
                Customer.BalanceDebit = IIf(dt.Rows(0).Item("BalanceDebit").Equals(DBNull.Value), 0, dt.Rows(0).Item("BalanceDebit"))
                Customer.BalanceCredit = IIf(dt.Rows(0).Item("BalanceCredit").Equals(DBNull.Value), 0, dt.Rows(0).Item("BalanceCredit"))
                Customer.SalesOrderAmount = IIf(dt.Rows(0).Item("SalesOrderAmount").Equals(DBNull.Value), 0, dt.Rows(0).Item("SalesOrderAmount"))
                Customer.PageNumber = IIf(dt.Rows(0).Item("PageNumber").Equals(DBNull.Value), 0, dt.Rows(0).Item("PageNumber"))
                Customer.AmountOpen = IIf(dt.Rows(0).Item("AmountOpen").Equals(DBNull.Value), 0, dt.Rows(0).Item("AmountOpen"))
                Customer.Factoring = IIf(dt.Rows(0).Item("Factoring").Equals(DBNull.Value), 0, dt.Rows(0).Item("Factoring"))
                Customer.ISOCountry = IIf(dt.Rows(0).Item("ISOCountry").Equals(DBNull.Value), "", dt.Rows(0).Item("ISOCountry"))
                Customer.LiableToPayVAT = IIf(dt.Rows(0).Item("LiableToPayVAT").Equals(DBNull.Value), 0, dt.Rows(0).Item("LiableToPayVAT"))
                Customer.BackOrders = IIf(dt.Rows(0).Item("BackOrders").Equals(DBNull.Value), 0, dt.Rows(0).Item("BackOrders"))
                Customer.CostCenter = IIf(dt.Rows(0).Item("CostCenter").Equals(DBNull.Value), "", dt.Rows(0).Item("CostCenter"))
                Customer.AddressNumber = IIf(dt.Rows(0).Item("AddressNumber").Equals(DBNull.Value), "", dt.Rows(0).Item("AddressNumber"))
                Customer.DeliveryAddress = IIf(dt.Rows(0).Item("DeliveryAddress").Equals(DBNull.Value), "", dt.Rows(0).Item("DeliveryAddress"))
                Customer.RouteCode = IIf(dt.Rows(0).Item("RouteCode").Equals(DBNull.Value), "", dt.Rows(0).Item("RouteCode"))
                Customer.InvoiceDiscount = IIf(dt.Rows(0).Item("InvoiceDiscount").Equals(DBNull.Value), 0, dt.Rows(0).Item("InvoiceDiscount"))
                Customer.PaymentConditionSearchCode = IIf(dt.Rows(0).Item("PaymentConditionSearchCode").Equals(DBNull.Value), "", dt.Rows(0).Item("PaymentConditionSearchCode"))
                Customer.SearchCodeGoods = IIf(dt.Rows(0).Item("SearchCodeGoods").Equals(DBNull.Value), "", dt.Rows(0).Item("SearchCodeGoods"))
                Customer.ExpenseCode = IIf(dt.Rows(0).Item("ExpenseCode").Equals(DBNull.Value), "", dt.Rows(0).Item("ExpenseCode"))
                Customer.ICONumber = IIf(dt.Rows(0).Item("ICONumber").Equals(DBNull.Value), "", dt.Rows(0).Item("ICONumber"))
                Customer.BankNumber2 = IIf(dt.Rows(0).Item("BankNumber2").Equals(DBNull.Value), "", dt.Rows(0).Item("BankNumber2"))
                Customer.Area = IIf(dt.Rows(0).Item("Area").Equals(DBNull.Value), "", dt.Rows(0).Item("Area"))
                Customer.InvoiceLayout = IIf(dt.Rows(0).Item("InvoiceLayout").Equals(DBNull.Value), "", dt.Rows(0).Item("InvoiceLayout"))
                Customer.InvoiceName = IIf(dt.Rows(0).Item("InvoiceName").Equals(DBNull.Value), "", dt.Rows(0).Item("InvoiceName"))
                Customer.Status = IIf(dt.Rows(0).Item("Status").Equals(DBNull.Value), "", dt.Rows(0).Item("Status"))
                Customer.InvoiceThreshold = IIf(dt.Rows(0).Item("InvoiceThreshold").Equals(DBNull.Value), 0, dt.Rows(0).Item("InvoiceThreshold"))
                Customer.Location = IIf(dt.Rows(0).Item("Location").Equals(DBNull.Value), "", dt.Rows(0).Item("Location"))
                Customer.VATSource = IIf(dt.Rows(0).Item("VATSource").Equals(DBNull.Value), "", dt.Rows(0).Item("VATSource"))
                Customer.CalculatePenaltyInterest = IIf(dt.Rows(0).Item("CalculatePenaltyInterest").Equals(DBNull.Value), 0, dt.Rows(0).Item("CalculatePenaltyInterest"))
                Customer.IntermediaryAssociate = IIf(dt.Rows(0).Item("IntermediaryAssociate").Equals(DBNull.Value), "", dt.Rows(0).Item("IntermediaryAssociate"))
                Customer.SendPenaltyInvoices = IIf(dt.Rows(0).Item("SendPenaltyInvoices").Equals(DBNull.Value), 0, dt.Rows(0).Item("SendPenaltyInvoices"))
                Customer.CentralizationAccount = IIf(dt.Rows(0).Item("CentralizationAccount").Equals(DBNull.Value), 0, dt.Rows(0).Item("CentralizationAccount"))
                Customer.Currency = IIf(dt.Rows(0).Item("Currency").Equals(DBNull.Value), "", dt.Rows(0).Item("Currency"))
                Customer.BankAccountNumber = IIf(dt.Rows(0).Item("BankAccountNumber").Equals(DBNull.Value), "", dt.Rows(0).Item("BankAccountNumber"))
                Customer.PaymentMethod = IIf(dt.Rows(0).Item("PaymentMethod").Equals(DBNull.Value), "", dt.Rows(0).Item("PaymentMethod"))
                Customer.InvoiceDebtor = IIf(dt.Rows(0).Item("InvoiceDebtor").Equals(DBNull.Value), "", dt.Rows(0).Item("InvoiceDebtor"))
                Customer.CreditLine = IIf(dt.Rows(0).Item("CreditLine").Equals(DBNull.Value), 0, dt.Rows(0).Item("CreditLine"))
                Customer.Discount = IIf(dt.Rows(0).Item("Discount").Equals(DBNull.Value), 0, dt.Rows(0).Item("Discount"))
                If Not (dt.Rows(0).Item("DateLastReminder").Equals(DBNull.Value)) Then Customer.DateLastReminder = dt.Rows(0).Item("DateLastReminder")
                Customer.VatNumber = IIf(dt.Rows(0).Item("VatNumber").Equals(DBNull.Value), "", dt.Rows(0).Item("VatNumber"))
                Customer.PurchaseReceipt = IIf(dt.Rows(0).Item("PurchaseReceipt").Equals(DBNull.Value), "", dt.Rows(0).Item("PurchaseReceipt"))
                Customer.OffSetAccount = IIf(dt.Rows(0).Item("OffSetAccount").Equals(DBNull.Value), "", dt.Rows(0).Item("OffSetAccount"))
                Customer.JournalCode = IIf(dt.Rows(0).Item("JournalCode").Equals(DBNull.Value), "", dt.Rows(0).Item("JournalCode"))
                Customer.Reminder = IIf(dt.Rows(0).Item("Reminder").Equals(DBNull.Value), 0, dt.Rows(0).Item("Reminder"))
                Customer.PaymentCondition = IIf(dt.Rows(0).Item("PaymentCondition").Equals(DBNull.Value), "", dt.Rows(0).Item("PaymentCondition"))
                Customer.PriceList = IIf(dt.Rows(0).Item("PriceList").Equals(DBNull.Value), "", dt.Rows(0).Item("PriceList"))
                Customer.ItemCode = IIf(dt.Rows(0).Item("ItemCode").Equals(DBNull.Value), "", dt.Rows(0).Item("ItemCode"))
                Customer.DeliveryMethod = IIf(dt.Rows(0).Item("DeliveryMethod").Equals(DBNull.Value), "", dt.Rows(0).Item("DeliveryMethod"))
                Customer.CheckDate = IIf(dt.Rows(0).Item("CheckDate").Equals(DBNull.Value), "", dt.Rows(0).Item("CheckDate"))
                Customer.StatFactor = IIf(dt.Rows(0).Item("StatFactor").Equals(DBNull.Value), "", dt.Rows(0).Item("StatFactor"))
                Customer.VatCode = IIf(dt.Rows(0).Item("VatCode").Equals(DBNull.Value), "", dt.Rows(0).Item("VatCode"))
                Customer.ChangeVatCode = IIf(dt.Rows(0).Item("ChangeVatCode").Equals(DBNull.Value), 0, dt.Rows(0).Item("ChangeVatCode"))
                Customer.IntrastatSystem = IIf(dt.Rows(0).Item("IntrastatSystem").Equals(DBNull.Value), "", dt.Rows(0).Item("IntrastatSystem"))
                Customer.ChangeIntraStatSystem = IIf(dt.Rows(0).Item("ChangeIntraStatSystem").Equals(DBNull.Value), 0, dt.Rows(0).Item("ChangeIntraStatSystem"))
                Customer.TransActionA = IIf(dt.Rows(0).Item("TransActionA").Equals(DBNull.Value), "", dt.Rows(0).Item("TransActionA"))
                Customer.ChangeTransActionA = IIf(dt.Rows(0).Item("ChangeTransActionA").Equals(DBNull.Value), 0, dt.Rows(0).Item("ChangeTransActionA"))
                Customer.TransActionB = IIf(dt.Rows(0).Item("TransActionB").Equals(DBNull.Value), "", dt.Rows(0).Item("TransActionB"))
                Customer.ChangeTransActionB = IIf(dt.Rows(0).Item("ChangeTransActionB").Equals(DBNull.Value), 0, dt.Rows(0).Item("ChangeTransActionB"))
                Customer.DestinationCountry = IIf(dt.Rows(0).Item("DestinationCountry").Equals(DBNull.Value), "", dt.Rows(0).Item("DestinationCountry"))
                Customer.ChangeDestinationCountry = IIf(dt.Rows(0).Item("ChangeDestinationCountry").Equals(DBNull.Value), 0, dt.Rows(0).Item("ChangeDestinationCountry"))
                Customer.Transport = IIf(dt.Rows(0).Item("Transport").Equals(DBNull.Value), "", dt.Rows(0).Item("Transport"))
                Customer.ChangeTransport = IIf(dt.Rows(0).Item("ChangeTransport").Equals(DBNull.Value), 0, dt.Rows(0).Item("ChangeTransport"))
                Customer.CityOfLoadUnload = IIf(dt.Rows(0).Item("CityOfLoadUnload").Equals(DBNull.Value), "", dt.Rows(0).Item("CityOfLoadUnload"))
                Customer.ChangeCityOfLoadUnload = IIf(dt.Rows(0).Item("ChangeCityOfLoadUnload").Equals(DBNull.Value), 0, dt.Rows(0).Item("ChangeCityOfLoadUnload"))
                Customer.DeliveryTerms = IIf(dt.Rows(0).Item("DeliveryTerms").Equals(DBNull.Value), "", dt.Rows(0).Item("DeliveryTerms"))
                Customer.ChangeDeliveryTerms = IIf(dt.Rows(0).Item("ChangeDeliveryTerms").Equals(DBNull.Value), 0, dt.Rows(0).Item("ChangeDeliveryTerms"))
                Customer.TransShipment = IIf(dt.Rows(0).Item("TransShipment").Equals(DBNull.Value), "", dt.Rows(0).Item("TransShipment"))
                Customer.ChangeTransShipment = IIf(dt.Rows(0).Item("ChangeTransShipment").Equals(DBNull.Value), 0, dt.Rows(0).Item("ChangeTransShipment"))
                Customer.TriangularCountry = IIf(dt.Rows(0).Item("TriangularCountry").Equals(DBNull.Value), "", dt.Rows(0).Item("TriangularCountry"))
                Customer.ChangeTriangularCountry = IIf(dt.Rows(0).Item("ChangeTriangularCountry").Equals(DBNull.Value), 0, dt.Rows(0).Item("ChangeTriangularCountry"))
                Customer.IntRegion = IIf(dt.Rows(0).Item("IntRegion").Equals(DBNull.Value), "", dt.Rows(0).Item("IntRegion"))
                Customer.ChangeIntRegion = IIf(dt.Rows(0).Item("ChangeIntRegion").Equals(DBNull.Value), 0, dt.Rows(0).Item("ChangeIntRegion"))
                Customer.Collect = IIf(dt.Rows(0).Item("Collect").Equals(DBNull.Value), "", dt.Rows(0).Item("Collect"))
                Customer.InvoiceCopies = IIf(dt.Rows(0).Item("InvoiceCopies").Equals(DBNull.Value), 0, dt.Rows(0).Item("InvoiceCopies"))
                Customer.PaymentDay1 = IIf(dt.Rows(0).Item("PaymentDay1").Equals(DBNull.Value), 0, dt.Rows(0).Item("PaymentDay1"))
                Customer.PaymentDay2 = IIf(dt.Rows(0).Item("PaymentDay2").Equals(DBNull.Value), 0, dt.Rows(0).Item("PaymentDay2"))
                Customer.PaymentDay3 = IIf(dt.Rows(0).Item("PaymentDay3").Equals(DBNull.Value), 0, dt.Rows(0).Item("PaymentDay3"))
                Customer.PaymentDay4 = IIf(dt.Rows(0).Item("PaymentDay4").Equals(DBNull.Value), 0, dt.Rows(0).Item("PaymentDay4"))
                Customer.PaymentDay5 = IIf(dt.Rows(0).Item("PaymentDay5").Equals(DBNull.Value), 0, dt.Rows(0).Item("PaymentDay5"))
                Customer.FiscalCode = IIf(dt.Rows(0).Item("FiscalCode").Equals(DBNull.Value), "", dt.Rows(0).Item("FiscalCode"))
                Customer.CreditCard = IIf(dt.Rows(0).Item("CreditCard").Equals(DBNull.Value), "", dt.Rows(0).Item("CreditCard"))
                If Not (dt.Rows(0).Item("ExpiryDate").Equals(DBNull.Value)) Then Customer.ExpiryDate = dt.Rows(0).Item("ExpiryDate")
                Customer.TextField6 = IIf(dt.Rows(0).Item("TextField6").Equals(DBNull.Value), "", dt.Rows(0).Item("TextField6"))
                Customer.TextField7 = IIf(dt.Rows(0).Item("TextField7").Equals(DBNull.Value), "", dt.Rows(0).Item("TextField7"))
                Customer.TextField8 = IIf(dt.Rows(0).Item("TextField8").Equals(DBNull.Value), "", dt.Rows(0).Item("TextField8"))
                Customer.TextField9 = IIf(dt.Rows(0).Item("TextField9").Equals(DBNull.Value), "", dt.Rows(0).Item("TextField9"))
                Customer.TextField10 = IIf(dt.Rows(0).Item("TextField10").Equals(DBNull.Value), "", dt.Rows(0).Item("TextField10"))
                Customer.AccountEmployeeId = IIf(dt.Rows(0).Item("AccountEmployeeId").Equals(DBNull.Value), 0, dt.Rows(0).Item("AccountEmployeeId"))
                Customer.CreditabilityScenario = IIf(dt.Rows(0).Item("CreditabilityScenario").Equals(DBNull.Value), "", dt.Rows(0).Item("CreditabilityScenario"))
                Customer.VatLiability = IIf(dt.Rows(0).Item("VatLiability").Equals(DBNull.Value), "", dt.Rows(0).Item("VatLiability"))
                Customer.Attention = IIf(dt.Rows(0).Item("Attention").Equals(DBNull.Value), 0, dt.Rows(0).Item("Attention"))
                Customer.Category = IIf(dt.Rows(0).Item("Category").Equals(DBNull.Value), "", dt.Rows(0).Item("Category"))
                Customer.StatementAddress = IIf(dt.Rows(0).Item("StatementAddress").Equals(DBNull.Value), "", dt.Rows(0).Item("StatementAddress"))
                Customer.StatementPrint = IIf(dt.Rows(0).Item("StatementPrint").Equals(DBNull.Value), 0, dt.Rows(0).Item("StatementPrint"))
                Customer.StatementNumber = IIf(dt.Rows(0).Item("StatementNumber").Equals(DBNull.Value), 0, dt.Rows(0).Item("StatementNumber"))
                If Not (dt.Rows(0).Item("StatementDate").Equals(DBNull.Value)) Then Customer.StatementDate = dt.Rows(0).Item("StatementDate")
                Customer.DefaultSelCode = IIf(dt.Rows(0).Item("DefaultSelCode").Equals(DBNull.Value), "", dt.Rows(0).Item("DefaultSelCode"))
                Customer.GroupPayments = IIf(dt.Rows(0).Item("GroupPayments").Equals(DBNull.Value), "", dt.Rows(0).Item("GroupPayments"))
                Customer.BOELimitAmount = IIf(dt.Rows(0).Item("BOELimitAmount").Equals(DBNull.Value), 0, dt.Rows(0).Item("BOELimitAmount"))
                Customer.BOEMaxAmount = IIf(dt.Rows(0).Item("BOEMaxAmount").Equals(DBNull.Value), 0, dt.Rows(0).Item("BOEMaxAmount"))
                Customer.ExtraDuty = IIf(dt.Rows(0).Item("ExtraDuty").Equals(DBNull.Value), 0, dt.Rows(0).Item("ExtraDuty"))
                Customer.RegionCD = IIf(dt.Rows(0).Item("RegionCD").Equals(DBNull.Value), "", dt.Rows(0).Item("RegionCD"))
                Customer.region = IIf(dt.Rows(0).Item("region").Equals(DBNull.Value), "", dt.Rows(0).Item("region"))
                Customer.IntermediaryAssociateID = IIf(dt.Rows(0).Item("IntermediaryAssociateID").Equals(DBNull.Value), "", dt.Rows(0).Item("IntermediaryAssociateID"))
                Customer.CompanyType = IIf(dt.Rows(0).Item("CompanyType").Equals(DBNull.Value), 0, dt.Rows(0).Item("CompanyType"))
                Customer.SalesPersonNumber = IIf(dt.Rows(0).Item("SalesPersonNumber").Equals(DBNull.Value), 0, dt.Rows(0).Item("SalesPersonNumber"))
                Customer.AccountTypeCode = IIf(dt.Rows(0).Item("AccountTypeCode").Equals(DBNull.Value), "", dt.Rows(0).Item("AccountTypeCode"))
                Customer.StatementFrequency = IIf(dt.Rows(0).Item("StatementFrequency").Equals(DBNull.Value), "", dt.Rows(0).Item("StatementFrequency"))
                Customer.AccountRating = IIf(dt.Rows(0).Item("AccountRating").Equals(DBNull.Value), "", dt.Rows(0).Item("AccountRating"))
                Customer.FinanceCharge = IIf(dt.Rows(0).Item("FinanceCharge").Equals(DBNull.Value), 0, dt.Rows(0).Item("FinanceCharge"))
                Customer.ParentAccount = IIf(dt.Rows(0).Item("ParentAccount").Equals(DBNull.Value), "", dt.Rows(0).Item("ParentAccount"))
                Customer.IsParentAccount = IIf(dt.Rows(0).Item("IsParentAccount").Equals(DBNull.Value), 0, dt.Rows(0).Item("IsParentAccount"))
                Customer.ShipVia = IIf(dt.Rows(0).Item("ShipVia").Equals(DBNull.Value), 0, dt.Rows(0).Item("ShipVia"))
                Customer.UPSZone = IIf(dt.Rows(0).Item("UPSZone").Equals(DBNull.Value), "", dt.Rows(0).Item("UPSZone"))
                Customer.Terms = IIf(dt.Rows(0).Item("Terms").Equals(DBNull.Value), "", dt.Rows(0).Item("Terms"))
                'Customer.IsTaxable = IIf(dt.Rows(0).Item("IsTaxable").Equals(DBNull.Value), 0, IIf(dt.Rows(0).Item("IsTaxable") = "Y", 1, 0))
                Customer.IsTaxable = IIf(dt.Rows(0).Item("IsTaxable").Equals(DBNull.Value), 0, dt.Rows(0).Item("IsTaxable"))
                Customer.TaxCode = IIf(dt.Rows(0).Item("TaxCode").Equals(DBNull.Value), "", dt.Rows(0).Item("TaxCode"))
                Customer.TaxCode2 = IIf(dt.Rows(0).Item("TaxCode2").Equals(DBNull.Value), "", dt.Rows(0).Item("TaxCode2"))
                Customer.TaxCode3 = IIf(dt.Rows(0).Item("TaxCode3").Equals(DBNull.Value), "", dt.Rows(0).Item("TaxCode3"))
                Customer.ExemptNumber = IIf(dt.Rows(0).Item("ExemptNumber").Equals(DBNull.Value), "", dt.Rows(0).Item("ExemptNumber"))
                Customer.AllowSubstituteItem = IIf(dt.Rows(0).Item("AllowSubstituteItem").Equals(DBNull.Value), 0, dt.Rows(0).Item("AllowSubstituteItem"))
                Customer.AllowBackOrders = IIf(dt.Rows(0).Item("AllowBackOrders").Equals(DBNull.Value), 0, dt.Rows(0).Item("AllowBackOrders"))
                Customer.AllowPartialShipment = IIf(dt.Rows(0).Item("AllowPartialShipment").Equals(DBNull.Value), 0, dt.Rows(0).Item("AllowPartialShipment"))
                Customer.DunningLetter = IIf(dt.Rows(0).Item("DunningLetter").Equals(DBNull.Value), 0, dt.Rows(0).Item("DunningLetter"))
                Customer.Comment1 = IIf(dt.Rows(0).Item("Comment1").Equals(DBNull.Value), "", dt.Rows(0).Item("Comment1"))
                Customer.Comment2 = IIf(dt.Rows(0).Item("Comment2").Equals(DBNull.Value), "", dt.Rows(0).Item("Comment2"))
                Customer.TaxSchedule = IIf(dt.Rows(0).Item("TaxSchedule").Equals(DBNull.Value), "", dt.Rows(0).Item("TaxSchedule"))
                Customer.CreditCardDescription = IIf(dt.Rows(0).Item("CreditCardDescription").Equals(DBNull.Value), "", dt.Rows(0).Item("CreditCardDescription"))
                Customer.CreditCardHolder = IIf(dt.Rows(0).Item("CreditCardHolder").Equals(DBNull.Value), "", dt.Rows(0).Item("CreditCardHolder"))
                Customer.DefaultInvoiceForm = IIf(dt.Rows(0).Item("DefaultInvoiceForm").Equals(DBNull.Value), 0, dt.Rows(0).Item("DefaultInvoiceForm"))
                Customer.LocationShort = IIf(dt.Rows(0).Item("LocationShort").Equals(DBNull.Value), 0, dt.Rows(0).Item("LocationShort"))
                Customer.TaxID = IIf(dt.Rows(0).Item("TaxID").Equals(DBNull.Value), "", dt.Rows(0).Item("TaxID"))
                Customer.BillParent = IIf(dt.Rows(0).Item("BillParent").Equals(DBNull.Value), 0, dt.Rows(0).Item("BillParent"))
                Customer.Balance = IIf(dt.Rows(0).Item("Balance").Equals(DBNull.Value), 0, dt.Rows(0).Item("Balance"))
                Customer.PhoneExtention = IIf(dt.Rows(0).Item("PhoneExtention").Equals(DBNull.Value), "", dt.Rows(0).Item("PhoneExtention"))
                Customer.AutomaticPayment = IIf(dt.Rows(0).Item("AutomaticPayment").Equals(DBNull.Value), 0, dt.Rows(0).Item("AutomaticPayment"))
                Customer.CustomerCode = IIf(dt.Rows(0).Item("CustomerCode").Equals(DBNull.Value), "", dt.Rows(0).Item("CustomerCode"))
                Customer.PurchaseOrderAmount = IIf(dt.Rows(0).Item("PurchaseOrderAmount").Equals(DBNull.Value), 0, dt.Rows(0).Item("PurchaseOrderAmount"))
                Customer.CountryOfOrigin = IIf(dt.Rows(0).Item("CountryOfOrigin").Equals(DBNull.Value), "", dt.Rows(0).Item("CountryOfOrigin"))
                Customer.ChangeCountryOfOrigin = IIf(dt.Rows(0).Item("ChangeCountryOfOrigin").Equals(DBNull.Value), 0, dt.Rows(0).Item("ChangeCountryOfOrigin"))
                Customer.CountryOfAssembly = IIf(dt.Rows(0).Item("CountryOfAssembly").Equals(DBNull.Value), "", dt.Rows(0).Item("CountryOfAssembly"))
                Customer.ChangeCountryOfAssembly = IIf(dt.Rows(0).Item("ChangeCountryOfAssembly").Equals(DBNull.Value), 0, dt.Rows(0).Item("ChangeCountryOfAssembly"))
                Customer.PurchaseOrderConfirmation = IIf(dt.Rows(0).Item("PurchaseOrderConfirmation").Equals(DBNull.Value), 0, dt.Rows(0).Item("PurchaseOrderConfirmation"))
                Customer.RecepientOfCommissions = IIf(dt.Rows(0).Item("RecepientOfCommissions").Equals(DBNull.Value), 0, dt.Rows(0).Item("RecepientOfCommissions"))
                Customer.CreditorType = IIf(dt.Rows(0).Item("CreditorType").Equals(DBNull.Value), "", dt.Rows(0).Item("CreditorType"))
                Customer.PercentageToBePaid = IIf(dt.Rows(0).Item("PercentageToBePaid").Equals(DBNull.Value), 0, dt.Rows(0).Item("PercentageToBePaid"))
                If Not (dt.Rows(0).Item("SignDate").Equals(DBNull.Value)) Then Customer.SignDate = dt.Rows(0).Item("SignDate")
                Customer.ImportOriginCode = IIf(dt.Rows(0).Item("ImportOriginCode").Equals(DBNull.Value), "", dt.Rows(0).Item("ImportOriginCode"))
                Customer.ParticipantNumber = IIf(dt.Rows(0).Item("ParticipantNumber").Equals(DBNull.Value), "", dt.Rows(0).Item("ParticipantNumber"))
                Customer.CertifiedSupplier = IIf(dt.Rows(0).Item("CertifiedSupplier").Equals(DBNull.Value), 0, dt.Rows(0).Item("CertifiedSupplier"))
                Customer.FederalIDNumber = IIf(dt.Rows(0).Item("FederalIDNumber").Equals(DBNull.Value), 0, dt.Rows(0).Item("FederalIDNumber"))
                Customer.FederalIDType = IIf(dt.Rows(0).Item("FederalIDType").Equals(DBNull.Value), "", dt.Rows(0).Item("FederalIDType"))
                Customer.FederalCategory = IIf(dt.Rows(0).Item("FederalCategory").Equals(DBNull.Value), "", dt.Rows(0).Item("FederalCategory"))
                Customer.AutoDistribute = IIf(dt.Rows(0).Item("AutoDistribute").Equals(DBNull.Value), 0, dt.Rows(0).Item("AutoDistribute"))
                Customer.SupplierStatus = IIf(dt.Rows(0).Item("SupplierStatus").Equals(DBNull.Value), "", dt.Rows(0).Item("SupplierStatus"))
                Customer.FOBCode = IIf(dt.Rows(0).Item("FOBCode").Equals(DBNull.Value), "", dt.Rows(0).Item("FOBCode"))
                Customer.PrintPrice = IIf(dt.Rows(0).Item("PrintPrice").Equals(DBNull.Value), 0, dt.Rows(0).Item("PrintPrice"))
                Customer.Acknowledge = IIf(dt.Rows(0).Item("Acknowledge").Equals(DBNull.Value), 0, dt.Rows(0).Item("Acknowledge"))
                Customer.Confirmed = IIf(dt.Rows(0).Item("Confirmed").Equals(DBNull.Value), 0, dt.Rows(0).Item("Confirmed"))
                Customer.LandedCostCode = IIf(dt.Rows(0).Item("LandedCostCode").Equals(DBNull.Value), "", dt.Rows(0).Item("LandedCostCode"))
                Customer.LandedCostCode2 = IIf(dt.Rows(0).Item("LandedCostCode2").Equals(DBNull.Value), "", dt.Rows(0).Item("LandedCostCode2"))
                Customer.LandedCostCode3 = IIf(dt.Rows(0).Item("LandedCostCode3").Equals(DBNull.Value), "", dt.Rows(0).Item("LandedCostCode3"))
                Customer.LandedCostCode4 = IIf(dt.Rows(0).Item("LandedCostCode4").Equals(DBNull.Value), "", dt.Rows(0).Item("LandedCostCode4"))
                Customer.LandedCostCode5 = IIf(dt.Rows(0).Item("LandedCostCode5").Equals(DBNull.Value), "", dt.Rows(0).Item("LandedCostCode5"))
                Customer.LandedCostCode6 = IIf(dt.Rows(0).Item("LandedCostCode6").Equals(DBNull.Value), "", dt.Rows(0).Item("LandedCostCode6"))
                Customer.LandedCostCode7 = IIf(dt.Rows(0).Item("LandedCostCode7").Equals(DBNull.Value), "", dt.Rows(0).Item("LandedCostCode7"))
                Customer.LandedCostCode8 = IIf(dt.Rows(0).Item("LandedCostCode8").Equals(DBNull.Value), "", dt.Rows(0).Item("LandedCostCode8"))
                Customer.LandedCostCode9 = IIf(dt.Rows(0).Item("LandedCostCode9").Equals(DBNull.Value), "", dt.Rows(0).Item("LandedCostCode9"))
                Customer.LandedCostCode10 = IIf(dt.Rows(0).Item("LandedCostCode10").Equals(DBNull.Value), "", dt.Rows(0).Item("LandedCostCode10"))
                Customer.DefaultPOForm = IIf(dt.Rows(0).Item("DefaultPOForm").Equals(DBNull.Value), 0, dt.Rows(0).Item("DefaultPOForm"))
                Customer.PayeeName = IIf(dt.Rows(0).Item("PayeeName").Equals(DBNull.Value), "", dt.Rows(0).Item("PayeeName"))
                Customer.CommodityCode1 = IIf(dt.Rows(0).Item("CommodityCode1").Equals(DBNull.Value), "", dt.Rows(0).Item("CommodityCode1"))
                Customer.CommodityCode2 = IIf(dt.Rows(0).Item("CommodityCode2").Equals(DBNull.Value), "", dt.Rows(0).Item("CommodityCode2"))
                Customer.CommodityCode3 = IIf(dt.Rows(0).Item("CommodityCode3").Equals(DBNull.Value), "", dt.Rows(0).Item("CommodityCode3"))
                Customer.CommodityCode4 = IIf(dt.Rows(0).Item("CommodityCode4").Equals(DBNull.Value), "", dt.Rows(0).Item("CommodityCode4"))
                Customer.CommodityCode5 = IIf(dt.Rows(0).Item("CommodityCode5").Equals(DBNull.Value), "", dt.Rows(0).Item("CommodityCode5"))
                Customer.SecurityLevel = IIf(dt.Rows(0).Item("SecurityLevel").Equals(DBNull.Value), 0, dt.Rows(0).Item("SecurityLevel"))
                Customer.ChamberOfCommerce = IIf(dt.Rows(0).Item("ChamberOfCommerce").Equals(DBNull.Value), "", dt.Rows(0).Item("ChamberOfCommerce"))
                Customer.DunsNumber = IIf(dt.Rows(0).Item("DunsNumber").Equals(DBNull.Value), "", dt.Rows(0).Item("DunsNumber"))
                Customer.TextField11 = IIf(dt.Rows(0).Item("TextField11").Equals(DBNull.Value), "", dt.Rows(0).Item("TextField11"))
                Customer.TextField12 = IIf(dt.Rows(0).Item("TextField12").Equals(DBNull.Value), "", dt.Rows(0).Item("TextField12"))
                Customer.TextField13 = IIf(dt.Rows(0).Item("TextField13").Equals(DBNull.Value), "", dt.Rows(0).Item("TextField13"))
                Customer.TextField14 = IIf(dt.Rows(0).Item("TextField14").Equals(DBNull.Value), "", dt.Rows(0).Item("TextField14"))
                Customer.TextField15 = IIf(dt.Rows(0).Item("TextField15").Equals(DBNull.Value), "", dt.Rows(0).Item("TextField15"))
                Customer.TextField16 = IIf(dt.Rows(0).Item("TextField16").Equals(DBNull.Value), "", dt.Rows(0).Item("TextField16"))
                Customer.TextField17 = IIf(dt.Rows(0).Item("TextField17").Equals(DBNull.Value), "", dt.Rows(0).Item("TextField17"))
                Customer.TextField18 = IIf(dt.Rows(0).Item("TextField18").Equals(DBNull.Value), "", dt.Rows(0).Item("TextField18"))
                Customer.TextField19 = IIf(dt.Rows(0).Item("TextField19").Equals(DBNull.Value), "", dt.Rows(0).Item("TextField19"))
                Customer.TextField20 = IIf(dt.Rows(0).Item("TextField20").Equals(DBNull.Value), "", dt.Rows(0).Item("TextField20"))
                Customer.TextField21 = IIf(dt.Rows(0).Item("TextField21").Equals(DBNull.Value), "", dt.Rows(0).Item("TextField21"))
                Customer.TextField22 = IIf(dt.Rows(0).Item("TextField22").Equals(DBNull.Value), "", dt.Rows(0).Item("TextField22"))
                Customer.TextField23 = IIf(dt.Rows(0).Item("TextField23").Equals(DBNull.Value), "", dt.Rows(0).Item("TextField23"))
                Customer.TextField24 = IIf(dt.Rows(0).Item("TextField24").Equals(DBNull.Value), "", dt.Rows(0).Item("TextField24"))
                Customer.TextField25 = IIf(dt.Rows(0).Item("TextField25").Equals(DBNull.Value), "", dt.Rows(0).Item("TextField25"))
                Customer.TextField26 = IIf(dt.Rows(0).Item("TextField26").Equals(DBNull.Value), "", dt.Rows(0).Item("TextField26"))
                Customer.TextField27 = IIf(dt.Rows(0).Item("TextField27").Equals(DBNull.Value), "", dt.Rows(0).Item("TextField27"))
                Customer.TextField28 = IIf(dt.Rows(0).Item("TextField28").Equals(DBNull.Value), "", dt.Rows(0).Item("TextField28"))
                Customer.TextField29 = IIf(dt.Rows(0).Item("TextField29").Equals(DBNull.Value), "", dt.Rows(0).Item("TextField29"))
                Customer.TextField30 = IIf(dt.Rows(0).Item("TextField30").Equals(DBNull.Value), "", dt.Rows(0).Item("TextField30"))
                Customer.PhoneQueue = IIf(dt.Rows(0).Item("PhoneQueue").Equals(DBNull.Value), "", dt.Rows(0).Item("PhoneQueue"))
                Customer.cmp_Directory = IIf(dt.Rows(0).Item("cmp_Directory").Equals(DBNull.Value), "", dt.Rows(0).Item("cmp_Directory"))
                Customer.GUIDField1 = IIf(dt.Rows(0).Item("GUIDField1").Equals(DBNull.Value), "", dt.Rows(0).Item("GUIDField1").ToString)
                Customer.GUIDField2 = IIf(dt.Rows(0).Item("GUIDField2").Equals(DBNull.Value), "", dt.Rows(0).Item("GUIDField2").ToString)
                Customer.GUIDField3 = IIf(dt.Rows(0).Item("GUIDField3").Equals(DBNull.Value), "", dt.Rows(0).Item("GUIDField3").ToString)
                Customer.GUIDField4 = IIf(dt.Rows(0).Item("GUIDField4").Equals(DBNull.Value), "", dt.Rows(0).Item("GUIDField4").ToString)
                Customer.GUIDField5 = IIf(dt.Rows(0).Item("GUIDField5").Equals(DBNull.Value), "", dt.Rows(0).Item("GUIDField5").ToString)
                Customer.NumIntField1 = IIf(dt.Rows(0).Item("NumIntField1").Equals(DBNull.Value), 0, dt.Rows(0).Item("NumIntField1"))
                Customer.NumIntField2 = IIf(dt.Rows(0).Item("NumIntField2").Equals(DBNull.Value), 0, dt.Rows(0).Item("NumIntField2"))
                Customer.NumIntField3 = IIf(dt.Rows(0).Item("NumIntField3").Equals(DBNull.Value), 0, dt.Rows(0).Item("NumIntField3"))
                Customer.NumIntField4 = IIf(dt.Rows(0).Item("NumIntField4").Equals(DBNull.Value), 0, dt.Rows(0).Item("NumIntField4"))
                Customer.NumIntField5 = IIf(dt.Rows(0).Item("NumIntField5").Equals(DBNull.Value), 0, dt.Rows(0).Item("NumIntField5"))
                ' Customer.EncryptedCreditCard = IIF(dt.Rows(0).item("EncryptedCreditCard")), "", dt.Rows(0).item("EncryptedCreditCard"))
                Customer.RuleItem = IIf(dt.Rows(0).Item("RuleItem").Equals(DBNull.Value), "", dt.Rows(0).Item("RuleItem"))
                Customer.Substitute = IIf(dt.Rows(0).Item("Substitute").Equals(DBNull.Value), 0, dt.Rows(0).Item("Substitute"))
                Customer.Division = IIf(dt.Rows(0).Item("Division").Equals(DBNull.Value), 0, dt.Rows(0).Item("Division"))
                Customer.AutoAllocate = IIf(dt.Rows(0).Item("AutoAllocate").Equals(DBNull.Value), 0, dt.Rows(0).Item("AutoAllocate"))
                Customer.FedIDType = IIf(dt.Rows(0).Item("FedIDType").Equals(DBNull.Value), "", dt.Rows(0).Item("FedIDType"))
                Customer.FedCategory = IIf(dt.Rows(0).Item("FedCategory").Equals(DBNull.Value), "", dt.Rows(0).Item("FedCategory"))
                Customer.Category_01 = IIf(dt.Rows(0).Item("Category_01").Equals(DBNull.Value), "", dt.Rows(0).Item("Category_01"))
                Customer.Category_02 = IIf(dt.Rows(0).Item("Category_02").Equals(DBNull.Value), "", dt.Rows(0).Item("Category_02"))
                Customer.Category_03 = IIf(dt.Rows(0).Item("Category_03").Equals(DBNull.Value), "", dt.Rows(0).Item("Category_03"))
                Customer.Category_04 = IIf(dt.Rows(0).Item("Category_04").Equals(DBNull.Value), "", dt.Rows(0).Item("Category_04"))
                Customer.Category_05 = IIf(dt.Rows(0).Item("Category_05").Equals(DBNull.Value), "", dt.Rows(0).Item("Category_05"))
                Customer.Category_06 = IIf(dt.Rows(0).Item("Category_06").Equals(DBNull.Value), "", dt.Rows(0).Item("Category_06"))
                Customer.Category_07 = IIf(dt.Rows(0).Item("Category_07").Equals(DBNull.Value), "", dt.Rows(0).Item("Category_07"))
                Customer.Category_08 = IIf(dt.Rows(0).Item("Category_08").Equals(DBNull.Value), "", dt.Rows(0).Item("Category_08"))
                Customer.Category_09 = IIf(dt.Rows(0).Item("Category_09").Equals(DBNull.Value), "", dt.Rows(0).Item("Category_09"))
                Customer.Category_10 = IIf(dt.Rows(0).Item("Category_10").Equals(DBNull.Value), "", dt.Rows(0).Item("Category_10"))
                Customer.Category_11 = IIf(dt.Rows(0).Item("Category_11").Equals(DBNull.Value), "", dt.Rows(0).Item("Category_11"))
                Customer.Category_12 = IIf(dt.Rows(0).Item("Category_12").Equals(DBNull.Value), "", dt.Rows(0).Item("Category_12"))
                Customer.Category_13 = IIf(dt.Rows(0).Item("Category_13").Equals(DBNull.Value), "", dt.Rows(0).Item("Category_13"))
                Customer.Category_14 = IIf(dt.Rows(0).Item("Category_14").Equals(DBNull.Value), "", dt.Rows(0).Item("Category_14"))
                Customer.Category_15 = IIf(dt.Rows(0).Item("Category_15").Equals(DBNull.Value), "", dt.Rows(0).Item("Category_15"))
                Customer.FedIDNumber = IIf(dt.Rows(0).Item("FedIDNumber").Equals(DBNull.Value), "", dt.Rows(0).Item("FedIDNumber"))
                Customer.DefaultDeliveryAddress = IIf(dt.Rows(0).Item("DefaultDeliveryAddress").Equals(DBNull.Value), "", dt.Rows(0).Item("DefaultDeliveryAddress"))

                'specific customer instructions (if not empty will be a popup)

                'messages from oei_item_spec_instruction table
                Customer.Specific_Instructions = IIf(dt.Rows(0).Item("Specific_Instructions").Equals(DBNull.Value), "", dt.Rows(0).Item("Specific_Instructions"))

                'Show paperproof inclusion status message (if applicable)
                If dt.Rows(0).Item("paperproof_status") <> "" Then
                    If Customer.Specific_Instructions <> "" Then
                        Customer.Specific_Instructions = Customer.Specific_Instructions & "\n"
                    End If
                    Customer.Specific_Instructions = Customer.Specific_Instructions & dt.Rows(0).Item("paperproof_status")
                End If

                '++ID 3.25.2020 message added for COVID2020
                Dim _mess As String = "1"
                cOrdLine.customerInstructions1(User_Def_Fld_2, _mess)

                If _mess.Length > 1 Then
                    If Customer.Specific_Instructions <> "" Then
                        Customer.Specific_Instructions = Customer.Specific_Instructions & "\n"
                    End If
                    Customer.Specific_Instructions = Customer.Specific_Instructions & _mess
                End If

                'moved to frmOrder.sstOrder.Click 2960
                '++ID 02.28.2020 3 party Message for enter free charge item  line ---------------

                '  Dim _mestext As String = ""

                '_mestext = spec_ship_charge(Trim(dt.Rows(0).Item("ShipVia").ToString), m_oOrder.Ordhead.Customer.ClassificationId)

                'If _mestext <> String.Empty Then
                '    If Customer.Specific_Instructions <> "" Then
                '        Customer.Specific_Instructions = Customer.Specific_Instructions & "\n"
                '    End If
                '    Customer.Specific_Instructions = Customer.Specific_Instructions & _mestext
                'End If

                'If Trim(dt.Rows(0).Item("ShipVia").ToString) = "72" Then
                '    If Customer.Specific_Instructions <> "" Then
                '        Customer.Specific_Instructions = Customer.Specific_Instructions & "\n"
                '    End If
                '    Customer.Specific_Instructions = Customer.Specific_Instructions & "Please enter free charge at line item for shipping 3rd party." 'dt.Rows(0).Item("paperproof_status")
                'End If
                '--------------------------------------------------------------------------------------

                'Show exact quantity inclusion status message (if applicable)
                If dt.Rows(0).Item("exact_status") <> "" Then
                    If Customer.Specific_Instructions <> "" Then
                        Customer.Specific_Instructions = Customer.Specific_Instructions & "\n"
                    End If
                    Customer.Specific_Instructions = Customer.Specific_Instructions & dt.Rows(0).Item("exact_status")
                End If

                '++ID 4.25.2018
                'Show blind quantity inclusion status message (if applicable)
                If dt.Rows(0).Item("blind_paperproof_status") <> "" Then
                    If Customer.Specific_Instructions <> "" Then
                        Customer.Specific_Instructions = Customer.Specific_Instructions & "\n"
                    End If
                    Customer.Specific_Instructions = Customer.Specific_Instructions & dt.Rows(0).Item("blind_paperproof_status")
                End If


                'replace unix linefeeds with vb equivalent
                Customer.Specific_Instructions = Replace(Customer.Specific_Instructions, "\n", vbCrLf)

                GetCustomerDefaultValues = True

            End If

            Contacts = New cOrderContacts(Ord_GUID)

        Catch er As Exception
            MsgBox("Error in COrdhead." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    Private Function GetProfit_Center(ByVal pstrCurr_Cd As String) As String

        GetProfit_Center = ""

        Try
            Dim strSql As String = _
            "SELECT ISNULL(Sls_Sb_No, '') AS Sls_Sb_No " & _
            "FROM   Artypfil_Sql WITH (NoLock) " & _
            "WHERE  CUS_TYPE_CD = '" & pstrCurr_Cd & "'"

            Dim dt As DataTable
            Dim db As New cDBA

            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then
                GetProfit_Center = dt.Rows(0).Item("sls_sb_no")
            End If

        Catch er As Exception
            MsgBox("Error in COrdhead." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    Public Sub GetAlternateAddress(Optional ByVal pstrCus_Alt_Adr_Cd As String = "")

        Try

            Dim strSql As String '= "" & _

            strSql = _
            "SELECT     TOP 1 WITH TIES * " & _
            "FROM       ARALTADR_SQL WITH (Nolock)  " & _
            "WHERE     (RTRIM(ARALTADR_SQL.cus_no) LIKE '" & RTrim(LTrim(Cus_No)) & "%' OR RTRIM(ARALTADR_SQL.cus_no) LIKE '% " & RTrim(LTrim(Cus_No)) & "%') AND " & _
            "		    CUS_ALT_ADR_CD = '" & Trim(SqlCompliantString(pstrCus_Alt_Adr_Cd)) & "' " & _
            "ORDER BY   ARALTADR_SQL.cus_alt_adr_cd "

            Dim db As New cDBA
            Dim dt As DataTable = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then
                'CUS_ALT_ADR_CD -> DS LE CODE 
                'ALTERNATE_ADDRESS -> DS LE NOM ADRESSE
                '            ADDR_1()
                '            ADDR_2()
                '            ADDR_3()
                'CITY & ", " & STATE & " " & ZIP DANS ADDR_4
                '            COUNTRY()
                m_Cus_Alt_Adr_Cd = pstrCus_Alt_Adr_Cd
                Ship_To_Name = IIf(dt.Rows(0).Item("user_def_fld_1").Equals(DBNull.Value), "", dt.Rows(0).Item("user_def_fld_1"))
                Ship_To_Addr_1 = IIf(dt.Rows(0).Item("ADDR_1").Equals(DBNull.Value), "", dt.Rows(0).Item("ADDR_1"))
                Ship_To_Addr_2 = IIf(dt.Rows(0).Item("ADDR_2").Equals(DBNull.Value), "", dt.Rows(0).Item("ADDR_2"))
                Ship_To_Addr_3 = IIf(dt.Rows(0).Item("ADDR_3").Equals(DBNull.Value), "", dt.Rows(0).Item("ADDR_3"))
                Ship_To_Addr_4 = IIf(dt.Rows(0).Item("CITY").Equals(DBNull.Value), "", dt.Rows(0).Item("CITY") & ", ")
                Ship_To_Addr_4 = Ship_To_Addr_4 & IIf(dt.Rows(0).Item("STATE").Equals(DBNull.Value), "", dt.Rows(0).Item("STATE") & " ")
                Ship_To_Addr_4 = Ship_To_Addr_4 & IIf(dt.Rows(0).Item("ZIP").Equals(DBNull.Value), "", dt.Rows(0).Item("ZIP"))
                Ship_To_Country = IIf(dt.Rows(0).Item("COUNTRY").Equals(DBNull.Value), "", dt.Rows(0).Item("COUNTRY"))

                Ship_To_City = IIf(dt.Rows(0).Item("CITY").Equals(DBNull.Value), "", dt.Rows(0).Item("CITY"))
                Ship_To_State = IIf(dt.Rows(0).Item("STATE").Equals(DBNull.Value), "", dt.Rows(0).Item("STATE"))
                Ship_To_Zip = IIf(dt.Rows(0).Item("ZIP").Equals(DBNull.Value), "", dt.Rows(0).Item("ZIP"))

            Else
                m_Cus_Alt_Adr_Cd = "SAME"
            End If

            Call Save()

        Catch er As Exception
            MsgBox("Error in COrdhead." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub Set_Header_Table_Name(ByVal pstrOrd_No As String)

        Try
            pstrOrd_No = pstrOrd_No.Trim

            Dim strSql As String
            Dim dt As DataTable
            Dim db As cDBA

            strSql = "" & _
            "SELECT TOP 1 * " & _
            "FROM   OEOrdHdr_Sql WITH (Nolock) " & _
            "WHERE  LTRIM(RTRIM(Ord_No)) = '" & Ord_No & "' "

            db = New cDBA
            dt = New DataTable

            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then
                Header_Table = "OEORDHDR_SQL"
                Lines_Table = "OEORDLIN_SQL"
            Else
                strSql = "" & _
                "SELECT TOP 1 * " & _
                "FROM   OEHdrHst_Sql WITH (Nolock) " & _
                "WHERE  LTRIM(RTRIM(Ord_No)) = '" & Ord_No & "' "

                db = New cDBA
                dt = New DataTable

                dt = db.DataTable(strSql)
                If dt.Rows.Count <> 0 Then
                    Header_Table = "OEHDRHST_SQL"
                    Lines_Table = "OELINHST_SQL"
                End If

            End If

            dt.Dispose()

        Catch er As Exception
            MsgBox("Error in COrdhead." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub SetOpenTS()

        Dim db As New cDBA

        Try

            If m_oOrder.Ordhead.ExportTS <> "" Then Throw New OEException(OEError.Order_Exported_To_Macola, False, False)

            Dim strSql As String

            strSql = _
            "UPDATE OEI_ORDHDR SET OpenTS = '" & Trim(SqlCompliantString(g_User.Usr_ID)) & "' + ' ' + CONVERT(CHAR(30), GETDATE(), 121) " & _
            "WHERE  Ord_Guid = '" & Ord_GUID & "' "

            db.Execute(strSql)

        Catch oe_er As OEException
            If oe_er.ShowMessage Then MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in COrdhead." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub UnsetOpenTS()

        Dim db As New cDBA

        Try

            If m_oOrder.Ordhead.ExportTS <> "" Then Throw New OEException(OEError.Order_Exported_To_Macola, False, False)

            Dim strSql As String

            strSql = _
            "UPDATE OEI_ORDHDR SET OpenTS = NULL " & _
            "WHERE  Ord_Guid = '" & Ord_GUID & "' "

            db.Execute(strSql)

        Catch oe_er As OEException
            If oe_er.ShowMessage Then MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in COrdhead." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Function Get_Descriptions() As cOrdheadDesc

        Get_Descriptions = New cOrdheadDesc

        Get_Descriptions.Location = Get_Loc_Description(Mfg_Loc)
        Get_Descriptions.ShipVia = Get_Ship_Via_Cd_Description(Ship_Via_Cd)
        Get_Descriptions.Terms = Get_Ar_Terms_Cd_Description(Ar_Terms_Cd)
        Get_Descriptions.TaxSchedule = Get_Tax_Sched_Description(Tax_Sched)
        Get_Descriptions.TaxCode1 = Get_Tax_Cd_Description(1, Tax_Cd)
        Get_Descriptions.TaxCode2 = Get_Tax_Cd_Description(2, Tax_Cd_2)
        Get_Descriptions.TaxCode3 = Get_Tax_Cd_Description(3, Tax_Cd_3)
        Get_Descriptions.Salesperson1 = Get_HumRes_ID_Description(Slspsn_No)

    End Function

    Public Function Get_Loc_Description(ByRef pstrLoc As String) As String

        Get_Loc_Description = ""

        Dim strSql As String = "SELECT * FROM IMLOCFIL_SQL WITH (Nolock) WHERE Loc = '" & pstrLoc & "'"

        Dim db As New cDBA
        Dim dt As DataTable = db.DataTable(strSql)

        If dt.Rows.Count <> 0 Then
            Get_Loc_Description = dt.Rows(0).Item("Loc_Desc")
        End If

    End Function

    Public Function Get_Ship_Via_Cd_Description(ByVal pstrShip_Via_Cd As String) As String

        Get_Ship_Via_Cd_Description = ""

        Dim strSql As String = "SELECT * FROM SYCDEFIL_SQL WITH (Nolock) WHERE sy_code = '" & pstrShip_Via_Cd & "'"

        Dim db As New cDBA
        Dim dt As DataTable = db.DataTable(strSql)

        If dt.Rows.Count <> 0 Then
            Get_Ship_Via_Cd_Description = dt.Rows(0).Item("Code_Desc")
        End If

    End Function

    ' Check what is entered in Macola on a new order.
    'strSql = "SELECT * FROM TABLE WITH (Nolock) WHERE Champ = '" & Val() & "'"
    'If rsdesc.RecordCount <> 0 Then
    '    Get_Descriptions.Status = rsDesc.Fields("Status_Desc").Value
    'End If
    'rsdesc.Close()

    Public Function Get_Ar_Terms_Cd_Description(ByVal pstrAr_Terms_Cd As String) As String

        Get_Ar_Terms_Cd_Description = ""

        Dim strSql As String = "SELECT * FROM SYTRMFIL_SQL WITH (Nolock) WHERE Term_Code = '" & pstrAr_Terms_Cd & "'"

        Dim db As New cDBA
        Dim dt As DataTable = db.DataTable(strSql)

        If dt.Rows.Count <> 0 Then
            Get_Ar_Terms_Cd_Description = dt.Rows(0).Item("Description")
        End If

    End Function

    Public Function Get_Tax_Sched_Description(ByVal pstrTax_Sched As String) As String

        Get_Tax_Sched_Description = ""

        Dim strSql As String = "SELECT * FROM TAXSCHED_SQL WITH (Nolock) WHERE Tax_Sched = '" & pstrTax_Sched & "'"

        Dim db As New cDBA
        Dim dt As DataTable = db.DataTable(strSql)

        If dt.Rows.Count <> 0 Then
            Get_Tax_Sched_Description = dt.Rows(0).Item("Tax_Sched_Desc")
        End If

    End Function

    Public Function Get_Tax_Cd_Description(ByVal piPos As Integer, ByVal pstrTax_Cd As String) As String

        Get_Tax_Cd_Description = ""

        Dim strSql As String = _
        "SELECT ISNULL(Tax_Cd_Description, '') AS Tax_Cd_Description " & _
        "FROM TAXDETL_SQL WITH (Nolock) WHERE Tax_Cd "
        strSql = strSql & " = '" & pstrTax_Cd & "'"

        Dim db As New cDBA
        Dim dt As DataTable = db.DataTable(strSql)

        If dt.Rows.Count <> 0 Then
            Get_Tax_Cd_Description = dt.Rows(0).Item("Tax_Cd_Description")
        End If

    End Function

    Public Function Get_HumRes_ID_Description(ByVal piSlspsn_No As Integer) As String

        Get_HumRes_ID_Description = ""
        Dim strSql As String = "SELECT * FROM ARSLMFIL_SQL WITH (Nolock) WHERE HumRes_ID = '" & piSlspsn_No & "'"

        Dim db As New cDBA
        Dim dt As DataTable = db.DataTable(strSql)

        If dt.Rows.Count <> 0 Then
            Get_HumRes_ID_Description = dt.Rows(0).Item("Slspsn_Name")
        End If

    End Function

    Private Sub Set_Curr_Trx_Rt()

        Try
            'Value is currency code or string code, will search for it else will put it as is.

            Dim strSql As String = "" & _
            "SELECT OMS30_0 FROM Valuta WITH (Nolock) WHERE Valcode = '" & Curr_Cd & "'"

            Dim db As New cDBA
            Dim dt As DataTable = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then
                Curr_Cd_Desc = dt.Rows(0).Item("OMS30_0")
            Else
                Curr_Cd_Desc = Curr_Cd
            End If

            If Ord_Dt.Year = 1 Then
                strSql = _
                "Select Rate_Exchange, Rate_Buy, Rate_Sell, Rate_Official " & _
                "FROM RATES WITH (Nolock) " & _
                "WHERE SOURCE_CURRENCY = '" & Curr_Cd & "' and TARGET_CURRENCY = 'CAD' AND " & _
                "DATE_L <= '" & String.Format("{0:D4}", Date.Now.Year) & String.Format("{0:D2}", Date.Now.Month) & String.Format("{0:D2}", Date.Now.Day) & "' order by DATE_L desc "
            Else
                strSql = _
                "Select Rate_Exchange, Rate_Buy, Rate_Sell, Rate_Official " & _
                "FROM RATES WITH (Nolock) " & _
                "WHERE SOURCE_CURRENCY = '" & Curr_Cd & "' and TARGET_CURRENCY = 'CAD' AND " & _
                "DATE_L <= '" & String.Format("{0:D4}", Ord_Dt.Year) & String.Format("{0:D2}", Ord_Dt.Month) & String.Format("{0:D2}", Ord_Dt.Day) & "' order by DATE_L desc "
            End If

            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then
                Curr_Trx_Rt = dt.Rows(0).Item("Rate_Exchange")
            Else
                Curr_Trx_Rt = 1
            End If

            If Orig_Trx_Rt = 0 Then Orig_Trx_Rt = Curr_Trx_Rt

        Catch er As Exception
            MsgBox("Error in cOrdHead." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Property Curr_Cd() As String
        Get
            Curr_Cd = m_Curr_Cd
        End Get
        Set(ByVal value As String)
            m_Curr_Cd = value
            Call Set_Curr_Trx_Rt()
        End Set
    End Property

    Public Property SendOrderAck() As Integer
        Get
            SendOrderAck = m_SendOrderAck
        End Get
        Set(ByVal value As Integer)
            m_SendOrderAck = value
        End Set
    End Property

    Public Sub Delete()

        Dim db As New cDBA

        Try

            If m_oOrder.Ordhead.ExportTS <> "" Then Throw New OEException(OEError.Order_Exported_To_Macola, False, False)

            Dim strSql As String

            strSql = _
            "DELETE FROM OEI_ORDHDR " & _
            "WHERE  Ord_Guid = '" & Ord_GUID & "' "

            db.Execute(strSql)

        Catch oe_er As OEException
            If oe_er.ShowMessage Then MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in COrdhead." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub Reset()

        ' Reset is the same as Init here, as we always create a NEW order or reloads one, so never keep anything.
        Call Init()

    End Sub

    Public Sub Save()

        Dim dt As DataTable
        Dim db As New cDBA
        Dim drRow As DataRow

        Try
            If Trim(OEI_Ord_No) = "" Then Exit Sub

            If Trim(Cus_No) = "" Then Exit Sub

            If ExportTS <> "" Then Exit Sub

            'Dim odbCommand As New OleDb.OleDbCommand()

            Dim strSql As String = "SELECT * FROM OEI_ORDHDR WHERE Ord_Guid = '" & Ord_GUID & "' "

            dt = db.DataTable(strSql)

            If dt.Rows.Count = 0 Then
                drRow = dt.NewRow()
            Else
                drRow = dt.Rows(0)
            End If

            Call SaveDataRow(drRow)

            If dt.Rows.Count = 0 Then
                drRow.Item("OpenTS") = Trim(g_User.Usr_ID) & " " & Date.Now.ToString
                db.DBDataTable.Rows.Add(drRow)
                Dim cmd As Odbc.OdbcCommandBuilder = New Odbc.OdbcCommandBuilder(db.DBDataAdapter)
                db.DBDataAdapter.InsertCommand = cmd.GetInsertCommand
            Else
                Dim cmd As Odbc.OdbcCommandBuilder = New Odbc.OdbcCommandBuilder(db.DBDataAdapter)
                db.DBDataAdapter.UpdateCommand = cmd.GetUpdateCommand
            End If

            db.DBDataAdapter.Update(db.DBDataTable)

            'Dim strSqlOrdHdr As String
            'strSqlOrdHdr = "SELECT * FROM OEI_ORDHDR WHERE Ord_Guid = '" & Ord_GUID & "' "

            'Dim conConnection As New OleDb.OleDbConnection("Provider=SQLOLEDB;" & gConn.ConnectionString)

            'odbCommand.Connection = conConnection

            'odbCommand.Connection.Open()

            'Dim dsDataSet As New DataSet

            'Dim daDataAdapter As New OleDb.OleDbDataAdapter(strSqlOrdHdr, conConnection)
            'daDataAdapter.Fill(dsDataSet, "OEI_ORDHDR")

            'Dim cbOEI_ORDHDR As OleDb.OleDbCommandBuilder = New OleDb.OleDbCommandBuilder(daDataAdapter)

            'Dim drNewRow As DataRow

            'If dsDataSet.Tables("OEI_ORDHDR").Rows.Count <> 0 Then

            '    drNewRow = dsDataSet.Tables("OEI_ORDHDR").Rows(0)
            '    daDataAdapter.UpdateCommand = cbOEI_ORDHDR.GetUpdateCommand()
            '    SaveDataRow(drNewRow)

            'Else

            '    drNewRow = dsDataSet.Tables("OEI_ORDHDR").NewRow()
            '    SaveDataRow(drNewRow)
            '    drNewRow.Item("OpenTS") = Trim(g_User.Usr_ID) & " " & Date.Now.ToString
            '    daDataAdapter.InsertCommand = cbOEI_ORDHDR.GetInsertCommand()
            '    dsDataSet.Tables("OEI_ORDHDR").Rows.Add(drNewRow)

            'End If

            'daDataAdapter.Update(dsDataSet, "OEI_ORDHDR")
            'dbtTransaction.Commit()

        Catch er As Exception
            MsgBox("Error in COrdead." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
            'Finally
            '    odbCommand.Connection.Close()
        End Try

    End Sub

    Public Sub SaveDataRow(ByRef pDataRow As DataRow)

        Try
            ' Nothing to change for this, should take the I
            pDataRow.Item("Ord_Type") = Mid(Ord_Type, 1, 1) ' "O" ' Ord_Type
            'pDataRow.Item("Ord_No") = "" ' Ord_No We don't fill Ord_No Now, will change with trigger.
            OEI_Ord_No = OEI_Ord_No.ToUpper
            pDataRow.Item("OEI_Ord_No") = Trim(OEI_Ord_No) ' Ord_No We don't fill Ord_No Now, will change with trigger.
            pDataRow.Item("Status") = Status ' "1" ' Status
            pDataRow.Item("Entered_Dt") = Date.Now.Date
            'If Entered_Dt.Year <> 1 Then pDataRow.Item("Entered_Dt") = Entered_Dt.Date
            If Ord_Dt.Year <> 1 Then pDataRow.Item("Ord_Dt") = Ord_Dt.Date

            If Shipping_Dt.Date.Equals(NoDate()) Then
                '                pDataRow.Item("Ord_Dt_Shipped") = DBNull.Value
                pDataRow.Item("Shipping_Dt") = DBNull.Value
            Else
                '                pDataRow.Item("Ord_Dt_Shipped") = Shipping_Dt.Date
                pDataRow.Item("Shipping_Dt") = Shipping_Dt.Date
            End If

            'pDataRow.Item("Ord_Dt") = Date.Now
            pDataRow.Item("Apply_To_No") = "" ' Apply_To_No
            pDataRow.Item("Oe_Po_No") = Trim(Oe_Po_No)
            pDataRow.Item("Cus_No") = Trim(Cus_No)
            pDataRow.Item("Bal_Meth") = Bal_Meth ' "O" ' Bal_Meth

            pDataRow.Item("Bill_To_Name") = IIf(Bill_To_Name = "", DBNull.Value, Left(Trim(Bill_To_Name).ToString, 40)) ' .Substring(0, 40))
            pDataRow.Item("Bill_To_Addr_1") = IIf(Bill_To_Addr_1 = "", DBNull.Value, Trim(Bill_To_Addr_1))
            pDataRow.Item("Bill_To_Addr_2") = IIf(Bill_To_Addr_2 = "", DBNull.Value, Trim(Bill_To_Addr_2))
            pDataRow.Item("Bill_To_Addr_3") = IIf(Bill_To_Addr_3 = "", DBNull.Value, Trim(Bill_To_Addr_3))

            pDataRow.Item("Bill_To_City") = IIf(Bill_To_City = "", DBNull.Value, Trim(Bill_To_City))
            pDataRow.Item("Bill_To_State") = IIf(Bill_To_State = "", DBNull.Value, Trim(Bill_To_State))
            pDataRow.Item("Bill_To_Zip") = IIf(Bill_To_Zip = "", DBNull.Value, Trim(Bill_To_Zip))

            Bill_To_Addr_4 = IIf(Bill_To_City.Equals(DBNull.Value), "", Trim(Bill_To_City) & ", ")
            Bill_To_Addr_4 = Bill_To_Addr_4 & IIf(Bill_To_State.Equals(DBNull.Value), "", Trim(Bill_To_State) & "  ")
            Bill_To_Addr_4 = Bill_To_Addr_4 & IIf(Bill_To_Zip.Equals(DBNull.Value), "", Trim(Bill_To_Zip) & "  ")

            pDataRow.Item("Bill_To_Addr_4") = IIf(Bill_To_Addr_4 = "", DBNull.Value, Mid(Bill_To_Addr_4, 1, 44))

            pDataRow.Item("Bill_To_Country") = IIf(Bill_To_Country = "", DBNull.Value, Trim(Bill_To_Country))

            If m_Cus_Alt_Adr_Cd <> "SAME" Then
                pDataRow.Item("Cus_Alt_Adr_Cd") = m_Cus_Alt_Adr_Cd
            Else
                pDataRow.Item("Cus_Alt_Adr_Cd") = DBNull.Value
            End If

            pDataRow.Item("Ship_To_Name") = IIf(Ship_To_Name = "", DBNull.Value, Trim(Ship_To_Name))
            pDataRow.Item("Ship_To_Addr_1") = IIf(Ship_To_Addr_1 = "", DBNull.Value, Trim(Ship_To_Addr_1))
            pDataRow.Item("Ship_To_Addr_2") = IIf(Ship_To_Addr_2 = "", DBNull.Value, Trim(Ship_To_Addr_2))
            pDataRow.Item("Ship_To_Addr_3") = IIf(Ship_To_Addr_3 = "", DBNull.Value, Trim(Ship_To_Addr_3))

            pDataRow.Item("Ship_To_City") = IIf(Ship_To_City = "", DBNull.Value, Trim(Ship_To_City)) ' ArCusFil_Sql.City
            pDataRow.Item("Ship_To_State") = IIf(Ship_To_State = "", DBNull.Value, Trim(Ship_To_State)) ' ArCusFil_Sql.State
            pDataRow.Item("Ship_To_Zip") = IIf(Ship_To_Zip = "", DBNull.Value, Trim(Ship_To_Zip)) ' ArCusFil_Sql.Zip

            Ship_To_Addr_4 = IIf(Ship_To_City.Equals(DBNull.Value), "", Trim(Ship_To_City) & ", ")
            Ship_To_Addr_4 = Ship_To_Addr_4 & IIf(Ship_To_State.Equals(DBNull.Value), "", Trim(Ship_To_State) & "  ")
            Ship_To_Addr_4 = Ship_To_Addr_4 & IIf(Ship_To_Zip.Equals(DBNull.Value), "", Trim(Ship_To_Zip) & "  ")

            pDataRow.Item("Ship_To_Addr_4") = IIf(Ship_To_Addr_4 = "", DBNull.Value, Mid(Ship_To_Addr_4, 1, 44))

            pDataRow.Item("Ship_To_Country") = IIf(Ship_To_Country = "", DBNull.Value, Trim(Ship_To_Country))

            If in_hands_date.Date.Equals(NoDate()) Then
                '                pDataRow.Item("Ord_Dt_Shipped") = DBNull.Value
                pDataRow.Item("In_Hands_Date") = DBNull.Value
            Else
                '                pDataRow.Item("Ord_Dt_Shipped") = Shipping_Dt.Date
                pDataRow.Item("In_Hands_Date") = In_Hands_Date.Date
            End If

            'If Shipping_Dt.Date.Equals(NoDate()) Then
            '    pDataRow.Item("Shipping_Dt") = DBNull.Value
            'Else
            '    If Entered_Dt.Year <> 1 Then pDataRow.Item("Shipping_Dt") = Shipping_Dt.Date
            'End If

            pDataRow.Item("Ship_Via_Cd") = Trim(Ship_Via_Cd)
            pDataRow.Item("Ar_Terms_Cd") = Trim(Ar_Terms_Cd)
            pDataRow.Item("Frt_Pay_Cd") = Frt_Pay_Cd ' "N" ' 

            pDataRow.Item("Ship_Instruction_1") = IIf(Ship_Instruction_1 = "", DBNull.Value, Trim(Ship_Instruction_1))
            pDataRow.Item("Ship_Instruction_2") = IIf(Ship_Instruction_2 = "", DBNull.Value, Trim(Ship_Instruction_2))

            pDataRow.Item("Slspsn_No") = Slspsn_No
            pDataRow.Item("Slspsn_Pct_Comm") = Slspsn_Pct_Comm
            pDataRow.Item("Slspsn_Comm_Amt") = Slspsn_Comm_Amt

            pDataRow.Item("Slspsn_No_2") = Slspsn_No_2 ' doit etre 0
            pDataRow.Item("Slspsn_Pct_Comm_2") = Slspsn_Pct_Comm_2 ' doit etre 0
            pDataRow.Item("Slspsn_Comm_Amt_2") = Slspsn_Comm_Amt_2 ' doit etre 0

            pDataRow.Item("Slspsn_No_3") = Slspsn_No_3 ' doit etre 0
            pDataRow.Item("Slspsn_Pct_Comm_3") = Slspsn_Pct_Comm_3 ' doit etre 0
            pDataRow.Item("Slspsn_Comm_Amt_3") = Slspsn_Comm_Amt_3 ' doit etre 0

            pDataRow.Item("Tax_Cd") = IIf(Tax_Cd = "", "", Trim(Tax_Cd))
            pDataRow.Item("Tax_Pct") = Tax_Pct ' 0

            pDataRow.Item("Tax_Cd_2") = Tax_Cd_2 ' .Trim ' null
            pDataRow.Item("Tax_Pct_2") = Tax_Pct_2 ' 0

            pDataRow.Item("Tax_Cd_3") = Tax_Cd_3 ' .Trim ' null
            pDataRow.Item("Tax_Pct_3") = Tax_Pct_3 ' 0

            pDataRow.Item("Discount_Pct") = Discount_Pct
            pDataRow.Item("Job_No") = IIf(Job_No = "", DBNull.Value, Trim(Job_No))
            pDataRow.Item("Mfg_Loc") = Trim(Mfg_Loc)
            pDataRow.Item("Profit_Center") = Trim(Profit_Center)
            pDataRow.Item("Dept") = IIf(Dept = "", DBNull.Value, Trim(Dept))

            'pDataRow.Item("Ar_Reference") = Ar_Reference
            'pDataRow.Item("Tot_Sls_Amt") = Tot_Sls_Amt             ' 20.08	
            'pDataRow.Item("Tot_Sls_Disc") = Tot_Sls_Disc           ' 0.62
            'pDataRow.Item("Tot_Tax_Amt") = Tot_Tax_Amt             ' 20.08
            'pDataRow.Item("Tot_Cost") = Tot_Cost                   ' 5.45
            'pDataRow.Item("Tot_Weight") = Tot_Weight               ' 0.75
            'pDataRow.Item("Misc_Amt") = Misc_Amt                   ' 0
            'pDataRow.Item("Misc_Mn_No") = Misc_Mn_No               ' null
            'pDataRow.Item("Misc_Sb_No") = Misc_Sb_No               ' null
            'pDataRow.Item("Misc_Dp_No") = Misc_Dp_No               ' null
            'pDataRow.Item("Frt_Amt") = Frt_Amt
            'pDataRow.Item("Frt_Mn_No") = Frt_Mn_No
            'pDataRow.Item("Frt_Sb_No") = Frt_Sb_No
            'pDataRow.Item("Frt_Dp_No") = Frt_Dp_No
            'pDataRow.Item("Sls_Tax_Amt_1") = Sls_Tax_Amt_1
            'pDataRow.Item("Sls_Tax_Amt_2") = Sls_Tax_Amt_2
            'pDataRow.Item("Sls_Tax_Amt_3") = Sls_Tax_Amt_3
            'pDataRow.Item("Comm_Pct") = Comm_Pct
            'pDataRow.Item("Comm_Amt") = Comm_Amt
            'pDataRow.Item("Cmt_1") = Cmt_1
            'pDataRow.Item("Cmt_2") = Cmt_2
            'pDataRow.Item("Cmt_3") = Cmt_3
            'pDataRow.Item("Payment_Amt") = Payment_Amt
            'pDataRow.Item("Payment_Disc_Amt") = Payment_Disc_Amt
            'pDataRow.Item("Chk_No") = Chk_No
            'pDataRow.Item("Chk_Dt") = Chk_Dt
            'pDataRow.Item("Cash_Mn_No") = Cash_Mn_No
            'pDataRow.Item("Cash_Sb_No") = Cash_Sb_No
            'pDataRow.Item("Cash_Dp_No") = Cash_Dp_No
            'pDataRow.Item("Ord_Dt_Picked") = Ord_Dt_Picked
            'pDataRow.Item("Ord_Dt_Billed") = Ord_Dt_Billed
            pDataRow.Item("Inv_No") = "" ' Inv_No
            'pDataRow.Item("Inv_Dt") = Inv_Dt
            pDataRow.Item("Selection_Cd") = "C" ' Selection_Cd      ' est I.
            'pDataRow.Item("Posted_Dt") = Posted_Dt
            'pDataRow.Item("Part_Posted_Fg") = Part_Posted_Fg
            'pDataRow.Item("Ship_To_Freefrm_Fg") = Ship_To_Freefrm_Fg
            'pDataRow.Item("Bill_To_Freefrm_Fg") = Bill_To_Freefrm_Fg
            'pDataRow.Item("Copy_To_Bm_Fg") = Copy_To_Bm_Fg
            pDataRow.Item("Edi_Fg") = Edi_Fg
            'pDataRow.Item("Closed_Fg") = Closed_Fg
            'pDataRow.Item("Accum_Misc_Amt") = Accum_Misc_Amt
            'pDataRow.Item("Accum_Frt_Amt") = Accum_Frt_Amt
            'pDataRow.Item("Accum_Tot_Tax_Amt") = Accum_Tot_Tax_Amt
            'pDataRow.Item("Accum_Sls_Tax_Amt") = Accum_Sls_Tax_Amt
            'pDataRow.Item("Accum_Tot_Sls_Amt") = Accum_Tot_Sls_Amt
            pDataRow.Item("Hold_Fg") = Hold_Fg
            pDataRow.Item("Prepayment_Fg") = "N" ' Prepayment_Fg
            'pDataRow.Item("Lost_Sale_Cd") = Lost_Sale_Cd
            'pDataRow.Item("Orig_Ord_Type") = Orig_Ord_Type
            'pDataRow.Item("Orig_Ord_Dt") = Orig_Ord_Dt
            'pDataRow.Item("Orig_Ord_No") = Orig_Ord_No
            pDataRow.Item("Award_Probability") = 0 ' Award_Probability
            pDataRow.Item("Oe_Cash_No") = "" ' Oe_Cash_No
            'pDataRow.Item("Exch_Rt_Fg") = Exch_Rt_Fg
            pDataRow.Item("Curr_Cd") = Trim(Curr_Cd)
            pDataRow.Item("Orig_Trx_Rt") = Curr_Trx_Rt ' Verifier sur un reload
            pDataRow.Item("Curr_Trx_Rt") = Curr_Trx_Rt

            pDataRow.Item("Tax_Sched") = IIf(Tax_Sched = "", DBNull.Value, Trim(Tax_Sched))

            pDataRow.Item("User_Def_Fld_1") = IIf(User_Def_Fld_1 = "", DBNull.Value, Left(User_Def_Fld_1.Trim, 30))
            pDataRow.Item("User_Def_Fld_2") = IIf(User_Def_Fld_2 = "", DBNull.Value, Left(User_Def_Fld_2.Trim, 30))
            pDataRow.Item("User_Def_Fld_3") = IIf(User_Def_Fld_3 = "", DBNull.Value, Left(User_Def_Fld_3.Trim, 30))
            pDataRow.Item("User_Def_Fld_4") = IIf(User_Def_Fld_4 = "", DBNull.Value, Left(User_Def_Fld_4.Trim, 30))
            pDataRow.Item("User_Def_Fld_5") = IIf(User_Def_Fld_5 = "", DBNull.Value, Left(User_Def_Fld_5.Trim, 30))

            pDataRow.Item("Deter_Rate_By") = Deter_Rate_By ' "I" ' Deter_Rate_By
            pDataRow.Item("Form_No") = Form_No                          ' ArCusFil_Sql.Dflt_Inv_Form

            pDataRow.Item("Tax_Fg") = IIf(Tax_Fg = "", DBNull.Value, Trim(Tax_Fg))  'ArCusFil.Txbl_Fg

            'pDataRow.Item("Sls_Mn_No") = Sls_Mn_No
            'pDataRow.Item("Sls_Sb_No") = Sls_Sb_No
            'pDataRow.Item("Sls_Dp_No") = Sls_Dp_No
            'pDataRow.Item("Ord_Dt_Shipped") = Ord_Dt_Shipped
            pDataRow.Item("Tot_Dollars") = Tot_Dollars                  ' 22.08 
            'pDataRow.Item("Mult_Loc_Fg") = Mult_Loc_Fg
            'pDataRow.Item("Tot_Tax_Cost") = Tot_Tax_Cost               ' 0
            'pDataRow.Item("Hist_Load_Record") = Hist_Load_Record
            'pDataRow.Item("Pre_Select_Status") = Pre_Select_Status     ' 1
            'pDataRow.Item("Packing_No") = Packing_No                   ' 0
            'pDataRow.Item("Deliv_Ar_Terms_Cd") = Deliv_Ar_Terms_Cd
            '++ID 07312024 uncoment 
            pDataRow.Item("Inv_Batch_Id") = Inv_Batch_Id

            'pDataRow.Item("Bill_To_No") = Bill_To_No
            pDataRow.Item("Rma_No") = "" ' Rma_No
            'pDataRow.Item("Prog_Term_No") = Prog_Term_No
            'pDataRow.Item("Extra_1") = Extra_1
            'pDataRow.Item("Extra_2") = Extra_2
            'pDataRow.Item("Extra_3") = Extra_3
            'pDataRow.Item("Extra_4") = Extra_4
            'pDataRow.Item("Extra_5") = Extra_5
            pDataRow.Item("Extra_6") = Extra_6
            'pDataRow.Item("Extra_7") = Extra_7
            'pDataRow.Item("Extra_8") = Extra_8

            '++ID 06.17.2020 uncommented field extra_9 for scribl project
            pDataRow.Item("Extra_9") = Extra_9
            '-------------------------------------
            'pDataRow.Item("Extra_10") = Extra_10
            'pDataRow.Item("Extra_11") = Extra_11
            'pDataRow.Item("Extra_12") = Extra_12
            'pDataRow.Item("Extra_13") = Extra_13
            'pDataRow.Item("Extra_14") = Extra_14
            'pDataRow.Item("Extra_15") = Extra_15
            'pDataRow.Item("Edi_Doc_Seq") = Edi_Doc_Seq
            pDataRow.Item("Contact_1") = IIf(Contact_1 = "", DBNull.Value, Mid(Contact_1.Trim, 30))
            pDataRow.Item("Phone_Number") = IIf(Phone_Number = "", DBNull.Value, Trim(Phone_Number))
            pDataRow.Item("Fax_Number") = IIf(Fax_Number = "", DBNull.Value, Trim(Fax_Number))
            pDataRow.Item("Email_Address") = IIf(Email_Address = "", DBNull.Value, Trim(Email_Address))

            'pDataRow.Item("Use_Email") = Use_Email
            'pDataRow.Item("Filler_0001") = Filler_0001
            'pDataRow.Item("Id") = Id

            pDataRow.Item("Ord_GUID") = Ord_GUID
            pDataRow.Item("UserID") = g_User.Usr_ID ' Trim(GetNTUserID())

            pDataRow.Item("Ship_Link") = Trim(Ship_Link)
            pDataRow.Item("SendOrderAck") = SendOrderAck
            pDataRow.Item("ContactView") = ContactView
            pDataRow.Item("Email_Sent") = 0
            pDataRow.Item("OrderAckSaveOnly") = OrderAckSaveOnly
            pDataRow.Item("AutoCompleteReship") = AutoCompleteReship

            pDataRow.Item("RepeatOrd_No") = IIf(RepeatOrd_No = "", DBNull.Value, Trim(RepeatOrd_No))
            pDataRow.Item("Cus_Prog_ID") = Cus_Prog_ID
            pDataRow.Item("Prog_Spector_Cd") = Prog_Spector_Cd
            pDataRow.Item("White_Glove") = White_Glove
            pDataRow.Item("SSP") = intSSP

            If in_hands_date.Date.Equals(NoDate()) Then
                pDataRow.Item("In_Hands_Date") = DBNull.Value
            Else
                pDataRow.Item("In_Hands_Date") = in_hands_date.Date
            End If

        Catch er As Exception
            MsgBox("Error in COrdhead." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Property OEI_Ord_No() As String
        Get
            OEI_Ord_No = m_OEI_Ord_No
        End Get
        Set(ByVal value As String)
            m_OEI_Ord_No = value
        End Set
    End Property

    Public Property Prog_Spector_Cd() As String
        Get
            Prog_Spector_Cd = m_strCus_Spector_Cd
        End Get
        Set(ByVal value As String)
            m_strCus_Spector_Cd = value
            Call SetCus_Prog_ID(m_strCus_Spector_Cd)
        End Set
    End Property

    Public Sub SetCus_Prog_ID(ByVal pstrCd As String)

        Dim db As New cDBA
        Dim dt As DataTable

        Cus_Prog_ID = 0

        Try
            Dim strsql As String = _
            "SELECT CUS_PROG_ID " & _
            "FROM MDB_CUS_PROG P WITH (NOLOCK) " & _
            "WHERE SPECTOR_CD = '" & pstrCd.Replace("'", "''").Trim & "' "

            dt = db.DataTable(strsql)
            If dt.Rows.Count <> 0 Then
                Cus_Prog_ID = dt.Rows(0).Item("Cus_Prog_ID")
            End If

        Catch er As Exception
            MsgBox("Error in COrdhead." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Function IsARepOrder() As Boolean

        IsARepOrder = False

        Try
            Dim db As New cDBA
            Dim dt As DataTable
            Dim strSql As String = "SELECT * FROM EXACT_TRAVELER_REP_CODES WITH (Nolock) WHERE REPCUST = '" & Cus_No & "' "

            dt = db.DataTable(strSql)
            IsARepOrder = (dt.Rows.Count <> 0)

        Catch er As Exception
            MsgBox("Error in COrdhead." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    Public Sub CheckExistingNumber(ByRef pOrd_No As xTextBox)

        Try
            Dim db As New cDBA
            Dim dt As DataTable
            Dim strSql As String =
            "SELECT * " &
            "FROM   OEI_ORDHDR WITH (Nolock) " &
            "WHERE  OEI_ORD_NO = '" & pOrd_No.Text.Trim & "' AND ORD_GUID <> '" & Ord_GUID & "' "

            dt = db.DataTable(strSql)
            If dt.Rows.Count <> 0 Then
                pOrd_No.Text = g_User.NextOrderNumber()
            End If

        Catch er As Exception
            MsgBox("Error in COrdhead." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    '============================================================
    'All Star Program - December 6th 2017 - Clayton Solomon
    '============================================================

    Private Sub AllStarProgram(ByVal companyCode As String)
        Try
            Dim db As New cDBA
            Dim dt As DataTable
            Dim _25percent As Double = 0.25
            Dim _50percent As Double = 0.5
            Dim starLevel As String = ""
            Dim orderType As String = m_oOrder.Ordhead.Ord_Type

            Dim sqlQuery As String = "SELECT ClassificationId, cmp_code, cmp_name, textfield4, region " &
                "FROM cicmpy " &
                "WHERE cmp_code = '" & companyCode.Trim & "' " &
                "AND region like '%USA%' OR region like '%CA%' " &
                "ORDER BY id"

            dt = db.DataTable(sqlQuery)

            If dt.Rows.Count > 0 Then

                starLevel = dt.Rows(0).Item("ClassificationId").ToString.ToUpper

                Select Case starLevel

                    Case "1ST", "2ST", "3ST"

                        If orderType = "RS" Then
                            'Customer must pay freight or be billed for it
                            'Item under $20 (EQP NET) are no charge
                            'Item OVER $20 (EQP NET) are charged at EQP-25%
                            'Key Accounts: EQP - 50%
                        ElseIf orderType = "SB" Or orderType = "SS" Then
                            'Customer must pay freight or be billed for it
                            'Pricing: EQP-25% (set-up, item cost and all other charges)
                            'Key Accounts: EQP - 50%
                        ElseIf orderType = "SP" Then
                            'Pricing: EQP-25% (set-up, item cost and all other charges)
                            'Key Accounts: EQP - 50%
                        End If

                    Case "4ST", "5ST", "6ST", "7ST"

                        If orderType = "RS" Then
                            'Customer must pay freight or be billed for it
                            'Item under $20 (EQP NET) are no charge
                            'Item OVER $20 (EQP NET) are charged at EQP-25%
                            'Key Accounts: EQP - 50%
                        ElseIf orderType = "SB" Or orderType = "SS" Then
                            'Customer must pay freight or be billed for it
                            'Pricing: EQP-25% (set-up, item cost and all other charges)
                            'Key Accounts: EQP - 50%
                        ElseIf orderType = "SP" Then
                            'Customer must pay freight or be billed for it
                            'Pricing: EQP-25% (set-up, item cost and all other charges)
                            'Key Accounts: EQP - 50%
                        End If

                    Case "CA", "N/A"

                End Select
            End If

        Catch ex As Exception
            MsgBox("Error in AllStarProgram. " & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try
    End Sub

End Class