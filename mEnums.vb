Public Module mEnums

    Public Enum ContactType

        MainContact = 0
        Paperproof
        PreShipping
        Shipping
        OrderAck
        Reminder90day

    End Enum

    Public Enum ContactMethod
        None
        Email
        Fax
    End Enum

    Public Enum ContactInfo
        None
        Primary
        Secondary
    End Enum

End Module
