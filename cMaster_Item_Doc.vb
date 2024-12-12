Public Class cMaster_Item_Doc
    Implements ICloneable


#Region "private variables #################################################"

    Private m_intItem_Doc_Id As Int32
        Private m_intItem_Id As Int32
        Private m_intColor_Id As Int32
        Private m_strItem_Doc_Desc As String
        Private m_strItem_Doc_Desc_Fr As String
        Private m_strDoc_Lang As String
        Private m_intItem_Doc_Type_Id As Int32
        Private m_bItem_Doc As Byte()
        Private m_strItem_Doc_Folder As String
        Private m_strItem_Doc_Name As String
        Private m_strItem_Doc_File_Ext As String
        Private m_intDec_Met_Id As Int32
        Private m_strUser_Login As String
        Private m_dtCreate_Ts As DateTime
        Private m_dtUpdate_Ts As DateTime

#End Region


#Region "Public properties ################################################"

        Public Property Item_Doc_Id() As Int32
            Get
                Item_Doc_Id = m_intItem_Doc_Id
            End Get
            Set(ByVal value As Int32)
                m_intItem_Doc_Id = value
            End Set
        End Property

        Public Property Item_Id() As Int32
            Get
                Item_Id = m_intItem_Id
            End Get
            Set(ByVal value As Int32)
                m_intItem_Id = value
            End Set
        End Property

        Public Property Color_Id() As Int32
            Get
                Color_Id = m_intColor_Id
            End Get
            Set(ByVal value As Int32)
                m_intColor_Id = value
            End Set
        End Property

        Public Property Item_Doc_Desc() As String
            Get
                Item_Doc_Desc = m_strItem_Doc_Desc
            End Get
            Set(ByVal value As String)
                m_strItem_Doc_Desc = value
            End Set
        End Property

        Public Property Item_Doc_Desc_Fr() As String
            Get
                Item_Doc_Desc_Fr = m_strItem_Doc_Desc_Fr
            End Get
            Set(ByVal value As String)
                m_strItem_Doc_Desc_Fr = value
            End Set
        End Property

        Public Property Doc_Lang() As String
            Get
                Doc_Lang = m_strDoc_Lang
            End Get
            Set(ByVal value As String)
                m_strDoc_Lang = value
            End Set
        End Property

        Public Property Item_Doc_Type_Id() As Int32
            Get
                Item_Doc_Type_Id = m_intItem_Doc_Type_Id
            End Get
            Set(ByVal value As Int32)
                m_intItem_Doc_Type_Id = value
            End Set
        End Property

        Public Property Item_Doc() As Byte()
            Get
                Item_Doc = m_bItem_Doc
            End Get
            Set(ByVal value As Byte())
                m_bItem_Doc = value
            End Set
        End Property

        Public Property Item_Doc_Folder() As String
            Get
                Item_Doc_Folder = m_strItem_Doc_Folder
            End Get
            Set(ByVal value As String)
                m_strItem_Doc_Folder = value
            End Set
        End Property

        Public Property Item_Doc_Name() As String
            Get
                Item_Doc_Name = m_strItem_Doc_Name
            End Get
            Set(ByVal value As String)
                m_strItem_Doc_Name = value
            End Set
        End Property

        Public Property Item_Doc_File_Ext() As String
            Get
                Item_Doc_File_Ext = m_strItem_Doc_File_Ext
            End Get
            Set(ByVal value As String)
                m_strItem_Doc_File_Ext = value
            End Set
        End Property

        Public Property Dec_Met_Id() As Int32
            Get
                Dec_Met_Id = m_intDec_Met_Id
            End Get
            Set(ByVal value As Int32)
                m_intDec_Met_Id = value
            End Set
        End Property

        Public Property User_Login() As String
            Get
                User_Login = m_strUser_Login
            End Get
            Set(ByVal value As String)
                m_strUser_Login = value
            End Set
        End Property

        Public Property Create_Ts() As DateTime
            Get
                Create_Ts = m_dtCreate_Ts
            End Get
            Set(ByVal value As DateTime)
                m_dtCreate_Ts = value
            End Set
        End Property

        Public Property Update_Ts() As DateTime
            Get
                Update_Ts = m_dtUpdate_Ts
            End Get
            Set(ByVal value As DateTime)
                m_dtUpdate_Ts = value
            End Set
        End Property

#End Region

    Public Function Clone() As Object Implements System.ICloneable.Clone
        Return Me.MemberwiseClone
    End Function


End Class
