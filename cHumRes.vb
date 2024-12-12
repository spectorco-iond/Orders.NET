Imports Microsoft.VisualBasic
Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.Data
Imports System.Diagnostics

' Insert Include here


Public Class cHumres


#Region "private variables #################################################"


    Private m_intId As Int32
    Private m_intRes_Id As Int32
    Private m_strFullname As String
    Private m_strSur_Name As String
    Private m_strFirst_Name As String
    Private m_strMiddle_Name As String
    Private m_strInitialen As String
    Private m_strMv1 As String
    Private m_strAdres1 As String
    Private m_strAdres2 As String
    Private m_strWoonpl As String
    Private m_strPostcode As String
    Private m_strLand_Iso As String
    Private m_strIso_Taalcd As String
    Private m_strMail As String
    Private m_strTelnr_Prv As String
    Private m_strTelnr_Prv2 As String
    Private m_strTelnr_Werk As String
    Private m_strTelnr_Werk2 As String
    Private m_strFaxnr As String
    Private m_strKamer As String
    Private m_strToestel As String
    Private m_strFuncnivo As String
    Private m_strSocsec_Nr As String
    Private m_dtGeb_Ldat As DateTime
    Private m_strGeb_Pl As String
    Private m_dtLdatindienst As DateTime
    Private m_dtLdatuitdienst As DateTime
    Private m_strUsr_Id As String
    Private m_bBlocked As Byte
    Private m_intRepto_Id As Int32
    Private m_strNat As String
    Private m_strMar_Stat As String
    Private m_dtMar_Date As DateTime
    Private m_strAddr_No As String
    Private m_strEmail As String
    Private m_strJob_Title As String
    Private m_strLoc As String
    Private m_strComp As String
    Private m_strTask As String
    Private m_strEmp_Type As String
    Private m_dtEmp_Statd As DateTime
    Private m_strProb_Per As String
    Private m_strProb_Pert As String
    Private m_dtProb_Endd As DateTime
    Private m_intNotes As Int32
    Private m_strReason_Resign As String
    Private m_strCrcard_Type As String
    Private m_strCrcard_No As String
    Private m_dtCrcard_Expd As DateTime
    Private m_strAddr_Xtra As String
    Private m_dtCont_End_Date As DateTime
    Private m_strWorkstat As String
    Private m_strWorkstat_Addr As String
    Private m_bBuyer As Byte
    Private m_bRepresentative As Byte
    Private m_bProjempl As Byte
    Private m_strMaiden_Name As String
    Private m_strPrafd_Code As String
    Private m_dblInttar As Double
    Private m_dblExttar As Double
    Private m_intJob_Level As Int32
    Private m_bPayempl As Byte
    Private m_strIdentiteit As String
    Private m_strKstdr_Code As String
    Private m_strUitk_Code As String
    Private m_strExtra_Text As String
    Private m_strExtra_Code As String
    Private m_strAdv_Code As String
    Private m_strSlip_Tekst As String
    Private m_strPredcode As String
    Private m_strCostcenter As String
    Private m_strCrdnr As String
    Private m_strBankac_0 As String
    Private m_strBankac_1 As String
    Private m_strValcode As String
    Private m_dblPur_Limit As Double
    Private m_dblFte As Double
    Private m_strPicturefilename As String
    Private m_strFreefield1 As String
    Private m_strFreefield2 As String
    Private m_strFreefield3 As String
    Private m_strFreefield4 As String
    Private m_strFreefield5 As String
    Private m_dtFreefield6 As DateTime
    Private m_dtFreefield7 As DateTime
    Private m_dtFreefield8 As DateTime
    Private m_dtFreefield9 As DateTime
    Private m_dtFreefield10 As DateTime
    Private m_dblFreefield11 As Double
    Private m_dblFreefield12 As Double
    Private m_dblFreefield13 As Double
    Private m_dblFreefield14 As Double
    Private m_dblFreefield15 As Double
    Private m_bFreefield16 As Byte
    Private m_bFreefield17 As Byte
    Private m_bFreefield18 As Byte
    Private m_bFreefield19 As Byte
    Private m_bFreefield20 As Byte
    Private m_strEmp_Stat As String
    Private m_intAbsent As Int32
    Private m_intAssistant As Int32
    Private m_guidCmp_Wwn As Guid
    Private m_strInternetac As String
    Private m_strMobile_Short As String
    Private m_strRating As String
    Private m_strUsergroup As String
    Private m_dtSyscreated As DateTime
    Private m_intSyscreator As Int32
    Private m_dtSysmodified As DateTime
    Private m_intSysmodifier As Int32
    Private m_guidSysguid As Guid
    'private m_bTimestamp As Byte[]
    'private m_bPicture As Byte[]
    Private m_strBtw_Nummer As String
    Private m_bIspersonalaccount As Byte
    Private m_intScale As Int32
    'private m_bSignature As Byte[]
    Private m_strSignaturefile As String
    Private m_dblBankinglimit As Double
    Private m_strOldntfullname As String
    Private m_intTimezone As Int32
    Private m_blnSynchronizeexchange As Boolean
    Private m_strStatecode As String
    Private m_strChildname1 As String
    Private m_dtChildbirthdate1 As DateTime
    Private m_strChildname2 As String
    Private m_dtChildbirthdate2 As DateTime
    Private m_strChildname3 As String
    Private m_dtChildbirthdate3 As DateTime
    Private m_strChildname4 As String
    Private m_dtChildbirthdate4 As DateTime
    Private m_strChildname5 As String
    Private m_dtChildbirthdate5 As DateTime
    Private m_strPartner As String
    Private m_dtPartnerbirthdate As DateTime
    Private m_strItemcode As String
    Private m_strUsr_Id2 As String
    Private m_strMainloc As String
    Private m_strPrefix As String
    Private m_strAffix As String
    Private m_strBirth_Prefix As String
    Private m_strBirth_Affix As String
    Private m_strPredcode2 As String
    Private m_strLand_Iso2 As String
    Private m_strRace As String
    Private m_dblSaleslimit As Double
    Private m_dblInvoicelimit As Double
    Private m_dblRmalimit As Double
    Private m_strSkypeid As String
    Private m_strMsnid As String
    Private m_bBackofficeblocked As Byte
    Private m_intProcessedbybackgroundjob As Int16
    Private m_dtVacancystartdate As DateTime
    Private m_strApplicationstage As String
    Private m_intDivision As Int16
    Private m_strVacancytype As String
    Private m_strClassification As String
    Private m_strJobcategory As String
    Private m_dtAdjustedhiredate As DateTime
    Private m_intPortalcreator As Int32


#End Region


#Region "Public constructors ##############################################"


    Public Sub New()

        Try

            Call Init()

        Catch ex As Exception
            MsgBox("Error in cHumres." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Sub


    Public Sub New(ByVal pstrUsr_Id As String)

        Try

            Call Init()
            Call Load(pstrUsr_Id)

        Catch ex As Exception
            MsgBox("Error in cHumres." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Sub

    Public Sub New(ByVal pId As Int32)

        Try

            Call Init()
            Call Load(pId)

        Catch ex As Exception
            MsgBox("Error in cHumres." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Sub


#End Region


#Region "Private maintenance procedures ###################################"


    Private Sub Init()

        Try

            m_intId = 0
            m_intRes_Id = 0
            m_strFullname = String.Empty
            m_strSur_Name = String.Empty
            m_strFirst_Name = String.Empty
            m_strMiddle_Name = String.Empty
            m_strInitialen = String.Empty
            m_strMv1 = String.Empty
            m_strAdres1 = String.Empty
            m_strAdres2 = String.Empty
            m_strWoonpl = String.Empty
            m_strPostcode = String.Empty
            m_strLand_Iso = String.Empty
            m_strIso_Taalcd = String.Empty
            m_strMail = String.Empty
            m_strTelnr_Prv = String.Empty
            m_strTelnr_Prv2 = String.Empty
            m_strTelnr_Werk = String.Empty
            m_strTelnr_Werk2 = String.Empty
            m_strFaxnr = String.Empty
            m_strKamer = String.Empty
            m_strToestel = String.Empty
            m_strFuncnivo = String.Empty
            m_strSocsec_Nr = String.Empty
            m_dtGeb_Ldat = NoDate()
            m_strGeb_Pl = String.Empty
            m_dtLdatindienst = NoDate()
            m_dtLdatuitdienst = NoDate()
            m_strUsr_Id = String.Empty

            m_intRepto_Id = 0
            m_strNat = String.Empty
            m_strMar_Stat = String.Empty
            m_dtMar_Date = NoDate()
            m_strAddr_No = String.Empty
            m_strEmail = String.Empty
            m_strJob_Title = String.Empty
            m_strLoc = String.Empty
            m_strComp = String.Empty
            m_strTask = String.Empty
            m_strEmp_Type = String.Empty
            m_dtEmp_Statd = NoDate()
            m_strProb_Per = String.Empty
            m_strProb_Pert = String.Empty
            m_dtProb_Endd = NoDate()
            m_intNotes = 0
            m_strReason_Resign = String.Empty
            m_strCrcard_Type = String.Empty
            m_strCrcard_No = String.Empty
            m_dtCrcard_Expd = NoDate()
            m_strAddr_Xtra = String.Empty
            m_dtCont_End_Date = NoDate()
            m_strWorkstat = String.Empty
            m_strWorkstat_Addr = String.Empty



            m_strMaiden_Name = String.Empty
            m_strPrafd_Code = String.Empty
            m_dblInttar = 0
            m_dblExttar = 0
            m_intJob_Level = 0

            m_strIdentiteit = String.Empty
            m_strKstdr_Code = String.Empty
            m_strUitk_Code = String.Empty
            m_strExtra_Text = String.Empty
            m_strExtra_Code = String.Empty
            m_strAdv_Code = String.Empty
            m_strSlip_Tekst = String.Empty
            m_strPredcode = String.Empty
            m_strCostcenter = String.Empty
            m_strCrdnr = String.Empty
            m_strBankac_0 = String.Empty
            m_strBankac_1 = String.Empty
            m_strValcode = String.Empty
            m_dblPur_Limit = 0
            m_dblFte = 0
            m_strPicturefilename = String.Empty
            m_strFreefield1 = String.Empty
            m_strFreefield2 = String.Empty
            m_strFreefield3 = String.Empty
            m_strFreefield4 = String.Empty
            m_strFreefield5 = String.Empty
            m_dtFreefield6 = NoDate()
            m_dtFreefield7 = NoDate()
            m_dtFreefield8 = NoDate()
            m_dtFreefield9 = NoDate()
            m_dtFreefield10 = NoDate()
            m_dblFreefield11 = 0
            m_dblFreefield12 = 0
            m_dblFreefield13 = 0
            m_dblFreefield14 = 0
            m_dblFreefield15 = 0





            m_strEmp_Stat = String.Empty
            m_intAbsent = 0
            m_intAssistant = 0
            m_guidCmp_Wwn = Nothing ' string.Empty
            m_strInternetac = String.Empty
            m_strMobile_Short = String.Empty
            m_strRating = String.Empty
            m_strUsergroup = String.Empty
            m_dtSyscreated = NoDate()
            m_intSyscreator = 0
            m_dtSysmodified = NoDate()
            m_intSysmodifier = 0
            m_guidSysguid = Nothing ' string.Empty


            m_strBtw_Nummer = String.Empty

            m_intScale = 0

            m_strSignaturefile = String.Empty
            m_dblBankinglimit = 0
            m_strOldntfullname = String.Empty
            m_intTimezone = 0
            m_blnSynchronizeexchange = False
            m_strStatecode = String.Empty
            m_strChildname1 = String.Empty
            m_dtChildbirthdate1 = NoDate()
            m_strChildname2 = String.Empty
            m_dtChildbirthdate2 = NoDate()
            m_strChildname3 = String.Empty
            m_dtChildbirthdate3 = NoDate()
            m_strChildname4 = String.Empty
            m_dtChildbirthdate4 = NoDate()
            m_strChildname5 = String.Empty
            m_dtChildbirthdate5 = NoDate()
            m_strPartner = String.Empty
            m_dtPartnerbirthdate = NoDate()
            m_strItemcode = String.Empty
            m_strUsr_Id2 = String.Empty
            m_strMainloc = String.Empty
            m_strPrefix = String.Empty
            m_strAffix = String.Empty
            m_strBirth_Prefix = String.Empty
            m_strBirth_Affix = String.Empty
            m_strPredcode2 = String.Empty
            m_strLand_Iso2 = String.Empty
            m_strRace = String.Empty
            m_dblSaleslimit = 0
            m_dblInvoicelimit = 0
            m_dblRmalimit = 0
            m_strSkypeid = String.Empty
            m_strMsnid = String.Empty

            m_intProcessedbybackgroundjob = 0
            m_dtVacancystartdate = NoDate()
            m_strApplicationstage = String.Empty
            m_intDivision = 0
            m_strVacancytype = String.Empty
            m_strClassification = String.Empty
            m_strJobcategory = String.Empty
            m_dtAdjustedhiredate = NoDate()
            m_intPortalcreator = 0

        Catch ex As Exception
            MsgBox("Error in cHumres." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Sub


    Private Sub LoadLine(ByRef pdrRow As DataRow)

        Try

            If Not (pdrRow.Item("Id").Equals(DBNull.Value)) Then m_intId = pdrRow.Item("Id")
            If Not (pdrRow.Item("Res_Id").Equals(DBNull.Value)) Then m_intRes_Id = pdrRow.Item("Res_Id")
            If Not (pdrRow.Item("Fullname").Equals(DBNull.Value)) Then m_strFullname = pdrRow.Item("Fullname").ToString
            If Not (pdrRow.Item("Sur_Name").Equals(DBNull.Value)) Then m_strSur_Name = pdrRow.Item("Sur_Name").ToString
            If Not (pdrRow.Item("First_Name").Equals(DBNull.Value)) Then m_strFirst_Name = pdrRow.Item("First_Name").ToString
            If Not (pdrRow.Item("Middle_Name").Equals(DBNull.Value)) Then m_strMiddle_Name = pdrRow.Item("Middle_Name").ToString
            If Not (pdrRow.Item("Initialen").Equals(DBNull.Value)) Then m_strInitialen = pdrRow.Item("Initialen").ToString
            If Not (pdrRow.Item("Mv1").Equals(DBNull.Value)) Then m_strMv1 = pdrRow.Item("Mv1").ToString
            If Not (pdrRow.Item("Adres1").Equals(DBNull.Value)) Then m_strAdres1 = pdrRow.Item("Adres1").ToString
            If Not (pdrRow.Item("Adres2").Equals(DBNull.Value)) Then m_strAdres2 = pdrRow.Item("Adres2").ToString
            If Not (pdrRow.Item("Woonpl").Equals(DBNull.Value)) Then m_strWoonpl = pdrRow.Item("Woonpl").ToString
            If Not (pdrRow.Item("Postcode").Equals(DBNull.Value)) Then m_strPostcode = pdrRow.Item("Postcode").ToString
            If Not (pdrRow.Item("Land_Iso").Equals(DBNull.Value)) Then m_strLand_Iso = pdrRow.Item("Land_Iso").ToString
            If Not (pdrRow.Item("Iso_Taalcd").Equals(DBNull.Value)) Then m_strIso_Taalcd = pdrRow.Item("Iso_Taalcd").ToString
            If Not (pdrRow.Item("Mail").Equals(DBNull.Value)) Then m_strMail = pdrRow.Item("Mail").ToString
            If Not (pdrRow.Item("Telnr_Prv").Equals(DBNull.Value)) Then m_strTelnr_Prv = pdrRow.Item("Telnr_Prv").ToString
            If Not (pdrRow.Item("Telnr_Prv2").Equals(DBNull.Value)) Then m_strTelnr_Prv2 = pdrRow.Item("Telnr_Prv2").ToString
            If Not (pdrRow.Item("Telnr_Werk").Equals(DBNull.Value)) Then m_strTelnr_Werk = pdrRow.Item("Telnr_Werk").ToString
            If Not (pdrRow.Item("Telnr_Werk2").Equals(DBNull.Value)) Then m_strTelnr_Werk2 = pdrRow.Item("Telnr_Werk2").ToString
            If Not (pdrRow.Item("Faxnr").Equals(DBNull.Value)) Then m_strFaxnr = pdrRow.Item("Faxnr").ToString
            If Not (pdrRow.Item("Kamer").Equals(DBNull.Value)) Then m_strKamer = pdrRow.Item("Kamer").ToString
            If Not (pdrRow.Item("Toestel").Equals(DBNull.Value)) Then m_strToestel = pdrRow.Item("Toestel").ToString
            If Not (pdrRow.Item("Funcnivo").Equals(DBNull.Value)) Then m_strFuncnivo = pdrRow.Item("Funcnivo").ToString
            If Not (pdrRow.Item("Socsec_Nr").Equals(DBNull.Value)) Then m_strSocsec_Nr = pdrRow.Item("Socsec_Nr").ToString
            If Not (pdrRow.Item("Geb_Ldat").Equals(DBNull.Value)) Then m_dtGeb_Ldat = pdrRow.Item("Geb_Ldat")
            If Not (pdrRow.Item("Geb_Pl").Equals(DBNull.Value)) Then m_strGeb_Pl = pdrRow.Item("Geb_Pl").ToString
            If Not (pdrRow.Item("Ldatindienst").Equals(DBNull.Value)) Then m_dtLdatindienst = pdrRow.Item("Ldatindienst")
            If Not (pdrRow.Item("Ldatuitdienst").Equals(DBNull.Value)) Then m_dtLdatuitdienst = pdrRow.Item("Ldatuitdienst")
            If Not (pdrRow.Item("Usr_Id").Equals(DBNull.Value)) Then m_strUsr_Id = pdrRow.Item("Usr_Id").ToString
            If Not (pdrRow.Item("Blocked").Equals(DBNull.Value)) Then m_bBlocked = pdrRow.Item("Blocked")
            If Not (pdrRow.Item("Repto_Id").Equals(DBNull.Value)) Then m_intRepto_Id = pdrRow.Item("Repto_Id")
            If Not (pdrRow.Item("Nat").Equals(DBNull.Value)) Then m_strNat = pdrRow.Item("Nat").ToString
            If Not (pdrRow.Item("Mar_Stat").Equals(DBNull.Value)) Then m_strMar_Stat = pdrRow.Item("Mar_Stat").ToString
            If Not (pdrRow.Item("Mar_Date").Equals(DBNull.Value)) Then m_dtMar_Date = pdrRow.Item("Mar_Date")
            If Not (pdrRow.Item("Addr_No").Equals(DBNull.Value)) Then m_strAddr_No = pdrRow.Item("Addr_No").ToString
            If Not (pdrRow.Item("Email").Equals(DBNull.Value)) Then m_strEmail = pdrRow.Item("Email").ToString
            If Not (pdrRow.Item("Job_Title").Equals(DBNull.Value)) Then m_strJob_Title = pdrRow.Item("Job_Title").ToString
            If Not (pdrRow.Item("Loc").Equals(DBNull.Value)) Then m_strLoc = pdrRow.Item("Loc").ToString
            If Not (pdrRow.Item("Comp").Equals(DBNull.Value)) Then m_strComp = pdrRow.Item("Comp").ToString
            If Not (pdrRow.Item("Task").Equals(DBNull.Value)) Then m_strTask = pdrRow.Item("Task").ToString
            If Not (pdrRow.Item("Emp_Type").Equals(DBNull.Value)) Then m_strEmp_Type = pdrRow.Item("Emp_Type").ToString
            If Not (pdrRow.Item("Emp_Statd").Equals(DBNull.Value)) Then m_dtEmp_Statd = pdrRow.Item("Emp_Statd")
            If Not (pdrRow.Item("Prob_Per").Equals(DBNull.Value)) Then m_strProb_Per = pdrRow.Item("Prob_Per").ToString
            If Not (pdrRow.Item("Prob_Pert").Equals(DBNull.Value)) Then m_strProb_Pert = pdrRow.Item("Prob_Pert").ToString
            If Not (pdrRow.Item("Prob_Endd").Equals(DBNull.Value)) Then m_dtProb_Endd = pdrRow.Item("Prob_Endd")
            If Not (pdrRow.Item("Notes").Equals(DBNull.Value)) Then m_intNotes = pdrRow.Item("Notes")
            If Not (pdrRow.Item("Reason_Resign").Equals(DBNull.Value)) Then m_strReason_Resign = pdrRow.Item("Reason_Resign").ToString
            If Not (pdrRow.Item("Crcard_Type").Equals(DBNull.Value)) Then m_strCrcard_Type = pdrRow.Item("Crcard_Type").ToString
            If Not (pdrRow.Item("Crcard_No").Equals(DBNull.Value)) Then m_strCrcard_No = pdrRow.Item("Crcard_No").ToString
            If Not (pdrRow.Item("Crcard_Expd").Equals(DBNull.Value)) Then m_dtCrcard_Expd = pdrRow.Item("Crcard_Expd")
            If Not (pdrRow.Item("Addr_Xtra").Equals(DBNull.Value)) Then m_strAddr_Xtra = pdrRow.Item("Addr_Xtra").ToString
            If Not (pdrRow.Item("Cont_End_Date").Equals(DBNull.Value)) Then m_dtCont_End_Date = pdrRow.Item("Cont_End_Date")
            If Not (pdrRow.Item("Workstat").Equals(DBNull.Value)) Then m_strWorkstat = pdrRow.Item("Workstat").ToString
            If Not (pdrRow.Item("Workstat_Addr").Equals(DBNull.Value)) Then m_strWorkstat_Addr = pdrRow.Item("Workstat_Addr").ToString
            If Not (pdrRow.Item("Buyer").Equals(DBNull.Value)) Then m_bBuyer = pdrRow.Item("Buyer")
            If Not (pdrRow.Item("Representative").Equals(DBNull.Value)) Then m_bRepresentative = pdrRow.Item("Representative")
            If Not (pdrRow.Item("Projempl").Equals(DBNull.Value)) Then m_bProjempl = pdrRow.Item("Projempl")
            If Not (pdrRow.Item("Maiden_Name").Equals(DBNull.Value)) Then m_strMaiden_Name = pdrRow.Item("Maiden_Name").ToString
            If Not (pdrRow.Item("Prafd_Code").Equals(DBNull.Value)) Then m_strPrafd_Code = pdrRow.Item("Prafd_Code").ToString
            If Not (pdrRow.Item("Inttar").Equals(DBNull.Value)) Then m_dblInttar = pdrRow.Item("Inttar")
            If Not (pdrRow.Item("Exttar").Equals(DBNull.Value)) Then m_dblExttar = pdrRow.Item("Exttar")
            If Not (pdrRow.Item("Job_Level").Equals(DBNull.Value)) Then m_intJob_Level = pdrRow.Item("Job_Level")
            If Not (pdrRow.Item("Payempl").Equals(DBNull.Value)) Then m_bPayempl = pdrRow.Item("Payempl")
            If Not (pdrRow.Item("Identiteit").Equals(DBNull.Value)) Then m_strIdentiteit = pdrRow.Item("Identiteit").ToString
            If Not (pdrRow.Item("Kstdr_Code").Equals(DBNull.Value)) Then m_strKstdr_Code = pdrRow.Item("Kstdr_Code").ToString
            If Not (pdrRow.Item("Uitk_Code").Equals(DBNull.Value)) Then m_strUitk_Code = pdrRow.Item("Uitk_Code").ToString
            If Not (pdrRow.Item("Extra_Text").Equals(DBNull.Value)) Then m_strExtra_Text = pdrRow.Item("Extra_Text").ToString
            If Not (pdrRow.Item("Extra_Code").Equals(DBNull.Value)) Then m_strExtra_Code = pdrRow.Item("Extra_Code").ToString
            If Not (pdrRow.Item("Adv_Code").Equals(DBNull.Value)) Then m_strAdv_Code = pdrRow.Item("Adv_Code").ToString
            If Not (pdrRow.Item("Slip_Tekst").Equals(DBNull.Value)) Then m_strSlip_Tekst = pdrRow.Item("Slip_Tekst").ToString
            If Not (pdrRow.Item("Predcode").Equals(DBNull.Value)) Then m_strPredcode = pdrRow.Item("Predcode").ToString
            If Not (pdrRow.Item("Costcenter").Equals(DBNull.Value)) Then m_strCostcenter = pdrRow.Item("Costcenter").ToString
            If Not (pdrRow.Item("Crdnr").Equals(DBNull.Value)) Then m_strCrdnr = pdrRow.Item("Crdnr").ToString
            If Not (pdrRow.Item("Bankac_0").Equals(DBNull.Value)) Then m_strBankac_0 = pdrRow.Item("Bankac_0").ToString
            If Not (pdrRow.Item("Bankac_1").Equals(DBNull.Value)) Then m_strBankac_1 = pdrRow.Item("Bankac_1").ToString
            If Not (pdrRow.Item("Valcode").Equals(DBNull.Value)) Then m_strValcode = pdrRow.Item("Valcode").ToString
            If Not (pdrRow.Item("Pur_Limit").Equals(DBNull.Value)) Then m_dblPur_Limit = pdrRow.Item("Pur_Limit")
            If Not (pdrRow.Item("Fte").Equals(DBNull.Value)) Then m_dblFte = pdrRow.Item("Fte")
            If Not (pdrRow.Item("Picturefilename").Equals(DBNull.Value)) Then m_strPicturefilename = pdrRow.Item("Picturefilename").ToString
            If Not (pdrRow.Item("Freefield1").Equals(DBNull.Value)) Then m_strFreefield1 = pdrRow.Item("Freefield1").ToString
            If Not (pdrRow.Item("Freefield2").Equals(DBNull.Value)) Then m_strFreefield2 = pdrRow.Item("Freefield2").ToString
            If Not (pdrRow.Item("Freefield3").Equals(DBNull.Value)) Then m_strFreefield3 = pdrRow.Item("Freefield3").ToString
            If Not (pdrRow.Item("Freefield4").Equals(DBNull.Value)) Then m_strFreefield4 = pdrRow.Item("Freefield4").ToString
            If Not (pdrRow.Item("Freefield5").Equals(DBNull.Value)) Then m_strFreefield5 = pdrRow.Item("Freefield5").ToString
            If Not (pdrRow.Item("Freefield6").Equals(DBNull.Value)) Then m_dtFreefield6 = pdrRow.Item("Freefield6")
            If Not (pdrRow.Item("Freefield7").Equals(DBNull.Value)) Then m_dtFreefield7 = pdrRow.Item("Freefield7")
            If Not (pdrRow.Item("Freefield8").Equals(DBNull.Value)) Then m_dtFreefield8 = pdrRow.Item("Freefield8")
            If Not (pdrRow.Item("Freefield9").Equals(DBNull.Value)) Then m_dtFreefield9 = pdrRow.Item("Freefield9")
            If Not (pdrRow.Item("Freefield10").Equals(DBNull.Value)) Then m_dtFreefield10 = pdrRow.Item("Freefield10")
            If Not (pdrRow.Item("Freefield11").Equals(DBNull.Value)) Then m_dblFreefield11 = pdrRow.Item("Freefield11")
            If Not (pdrRow.Item("Freefield12").Equals(DBNull.Value)) Then m_dblFreefield12 = pdrRow.Item("Freefield12")
            If Not (pdrRow.Item("Freefield13").Equals(DBNull.Value)) Then m_dblFreefield13 = pdrRow.Item("Freefield13")
            If Not (pdrRow.Item("Freefield14").Equals(DBNull.Value)) Then m_dblFreefield14 = pdrRow.Item("Freefield14")
            If Not (pdrRow.Item("Freefield15").Equals(DBNull.Value)) Then m_dblFreefield15 = pdrRow.Item("Freefield15")
            If Not (pdrRow.Item("Freefield16").Equals(DBNull.Value)) Then m_bFreefield16 = pdrRow.Item("Freefield16")
            If Not (pdrRow.Item("Freefield17").Equals(DBNull.Value)) Then m_bFreefield17 = pdrRow.Item("Freefield17")
            If Not (pdrRow.Item("Freefield18").Equals(DBNull.Value)) Then m_bFreefield18 = pdrRow.Item("Freefield18")
            If Not (pdrRow.Item("Freefield19").Equals(DBNull.Value)) Then m_bFreefield19 = pdrRow.Item("Freefield19")
            If Not (pdrRow.Item("Freefield20").Equals(DBNull.Value)) Then m_bFreefield20 = pdrRow.Item("Freefield20")
            If Not (pdrRow.Item("Emp_Stat").Equals(DBNull.Value)) Then m_strEmp_Stat = pdrRow.Item("Emp_Stat").ToString
            If Not (pdrRow.Item("Absent").Equals(DBNull.Value)) Then m_intAbsent = pdrRow.Item("Absent")
            If Not (pdrRow.Item("Assistant").Equals(DBNull.Value)) Then m_intAssistant = pdrRow.Item("Assistant")
            If Not (pdrRow.Item("Cmp_Wwn").Equals(DBNull.Value)) Then m_guidCmp_Wwn = pdrRow.Item("Cmp_Wwn")
            If Not (pdrRow.Item("Internetac").Equals(DBNull.Value)) Then m_strInternetac = pdrRow.Item("Internetac").ToString
            If Not (pdrRow.Item("Mobile_Short").Equals(DBNull.Value)) Then m_strMobile_Short = pdrRow.Item("Mobile_Short").ToString
            If Not (pdrRow.Item("Rating").Equals(DBNull.Value)) Then m_strRating = pdrRow.Item("Rating").ToString
            If Not (pdrRow.Item("Usergroup").Equals(DBNull.Value)) Then m_strUsergroup = pdrRow.Item("Usergroup").ToString
            If Not (pdrRow.Item("Syscreated").Equals(DBNull.Value)) Then m_dtSyscreated = pdrRow.Item("Syscreated")
            If Not (pdrRow.Item("Syscreator").Equals(DBNull.Value)) Then m_intSyscreator = pdrRow.Item("Syscreator")
            If Not (pdrRow.Item("Sysmodified").Equals(DBNull.Value)) Then m_dtSysmodified = pdrRow.Item("Sysmodified")
            If Not (pdrRow.Item("Sysmodifier").Equals(DBNull.Value)) Then m_intSysmodifier = pdrRow.Item("Sysmodifier")
            If Not (pdrRow.Item("Sysguid").Equals(DBNull.Value)) Then m_guidSysguid = pdrRow.Item("Sysguid")
            'If Not (pdrRow.Item("Timestamp").Equals(DBNull.Value)) Then m_bTimestamp = pdrRow.Item("Timestamp")
            'If Not (pdrRow.Item("Picture").Equals(DBNull.Value)) Then m_bPicture = pdrRow.Item("Picture")
            If Not (pdrRow.Item("Btw_Nummer").Equals(DBNull.Value)) Then m_strBtw_Nummer = pdrRow.Item("Btw_Nummer").ToString
            If Not (pdrRow.Item("Ispersonalaccount").Equals(DBNull.Value)) Then m_bIspersonalaccount = pdrRow.Item("Ispersonalaccount")
            If Not (pdrRow.Item("Scale").Equals(DBNull.Value)) Then m_intScale = pdrRow.Item("Scale")
            'If Not (pdrRow.Item("Signature").Equals(DBNull.Value)) Then m_bSignature = pdrRow.Item("Signature")
            If Not (pdrRow.Item("Signaturefile").Equals(DBNull.Value)) Then m_strSignaturefile = pdrRow.Item("Signaturefile").ToString
            If Not (pdrRow.Item("Bankinglimit").Equals(DBNull.Value)) Then m_dblBankinglimit = pdrRow.Item("Bankinglimit")
            If Not (pdrRow.Item("Oldntfullname").Equals(DBNull.Value)) Then m_strOldntfullname = pdrRow.Item("Oldntfullname").ToString
            If Not (pdrRow.Item("Timezone").Equals(DBNull.Value)) Then m_intTimezone = pdrRow.Item("Timezone")
            If Not (pdrRow.Item("Synchronizeexchange").Equals(DBNull.Value)) Then m_blnSynchronizeexchange = pdrRow.Item("Synchronizeexchange")
            If Not (pdrRow.Item("Statecode").Equals(DBNull.Value)) Then m_strStatecode = pdrRow.Item("Statecode").ToString
            If Not (pdrRow.Item("Childname1").Equals(DBNull.Value)) Then m_strChildname1 = pdrRow.Item("Childname1").ToString
            If Not (pdrRow.Item("Childbirthdate1").Equals(DBNull.Value)) Then m_dtChildbirthdate1 = pdrRow.Item("Childbirthdate1")
            If Not (pdrRow.Item("Childname2").Equals(DBNull.Value)) Then m_strChildname2 = pdrRow.Item("Childname2").ToString
            If Not (pdrRow.Item("Childbirthdate2").Equals(DBNull.Value)) Then m_dtChildbirthdate2 = pdrRow.Item("Childbirthdate2")
            If Not (pdrRow.Item("Childname3").Equals(DBNull.Value)) Then m_strChildname3 = pdrRow.Item("Childname3").ToString
            If Not (pdrRow.Item("Childbirthdate3").Equals(DBNull.Value)) Then m_dtChildbirthdate3 = pdrRow.Item("Childbirthdate3")
            If Not (pdrRow.Item("Childname4").Equals(DBNull.Value)) Then m_strChildname4 = pdrRow.Item("Childname4").ToString
            If Not (pdrRow.Item("Childbirthdate4").Equals(DBNull.Value)) Then m_dtChildbirthdate4 = pdrRow.Item("Childbirthdate4")
            If Not (pdrRow.Item("Childname5").Equals(DBNull.Value)) Then m_strChildname5 = pdrRow.Item("Childname5").ToString
            If Not (pdrRow.Item("Childbirthdate5").Equals(DBNull.Value)) Then m_dtChildbirthdate5 = pdrRow.Item("Childbirthdate5")
            If Not (pdrRow.Item("Partner").Equals(DBNull.Value)) Then m_strPartner = pdrRow.Item("Partner").ToString
            If Not (pdrRow.Item("Partnerbirthdate").Equals(DBNull.Value)) Then m_dtPartnerbirthdate = pdrRow.Item("Partnerbirthdate")
            If Not (pdrRow.Item("Itemcode").Equals(DBNull.Value)) Then m_strItemcode = pdrRow.Item("Itemcode").ToString
            If Not (pdrRow.Item("Usr_Id2").Equals(DBNull.Value)) Then m_strUsr_Id2 = pdrRow.Item("Usr_Id2").ToString
            If Not (pdrRow.Item("Mainloc").Equals(DBNull.Value)) Then m_strMainloc = pdrRow.Item("Mainloc").ToString
            If Not (pdrRow.Item("Prefix").Equals(DBNull.Value)) Then m_strPrefix = pdrRow.Item("Prefix").ToString
            If Not (pdrRow.Item("Affix").Equals(DBNull.Value)) Then m_strAffix = pdrRow.Item("Affix").ToString
            If Not (pdrRow.Item("Birth_Prefix").Equals(DBNull.Value)) Then m_strBirth_Prefix = pdrRow.Item("Birth_Prefix").ToString
            If Not (pdrRow.Item("Birth_Affix").Equals(DBNull.Value)) Then m_strBirth_Affix = pdrRow.Item("Birth_Affix").ToString
            If Not (pdrRow.Item("Predcode2").Equals(DBNull.Value)) Then m_strPredcode2 = pdrRow.Item("Predcode2").ToString
            If Not (pdrRow.Item("Land_Iso2").Equals(DBNull.Value)) Then m_strLand_Iso2 = pdrRow.Item("Land_Iso2").ToString
            If Not (pdrRow.Item("Race").Equals(DBNull.Value)) Then m_strRace = pdrRow.Item("Race").ToString
            If Not (pdrRow.Item("Saleslimit").Equals(DBNull.Value)) Then m_dblSaleslimit = pdrRow.Item("Saleslimit")
            If Not (pdrRow.Item("Invoicelimit").Equals(DBNull.Value)) Then m_dblInvoicelimit = pdrRow.Item("Invoicelimit")
            If Not (pdrRow.Item("Rmalimit").Equals(DBNull.Value)) Then m_dblRmalimit = pdrRow.Item("Rmalimit")
            If Not (pdrRow.Item("Skypeid").Equals(DBNull.Value)) Then m_strSkypeid = pdrRow.Item("Skypeid").ToString
            If Not (pdrRow.Item("Msnid").Equals(DBNull.Value)) Then m_strMsnid = pdrRow.Item("Msnid").ToString
            If Not (pdrRow.Item("Backofficeblocked").Equals(DBNull.Value)) Then m_bBackofficeblocked = pdrRow.Item("Backofficeblocked")
            If Not (pdrRow.Item("Processedbybackgroundjob").Equals(DBNull.Value)) Then m_intProcessedbybackgroundjob = pdrRow.Item("Processedbybackgroundjob")
            If Not (pdrRow.Item("Vacancystartdate").Equals(DBNull.Value)) Then m_dtVacancystartdate = pdrRow.Item("Vacancystartdate")
            If Not (pdrRow.Item("Applicationstage").Equals(DBNull.Value)) Then m_strApplicationstage = pdrRow.Item("Applicationstage").ToString
            If Not (pdrRow.Item("Division").Equals(DBNull.Value)) Then m_intDivision = pdrRow.Item("Division")
            If Not (pdrRow.Item("Vacancytype").Equals(DBNull.Value)) Then m_strVacancytype = pdrRow.Item("Vacancytype").ToString
            If Not (pdrRow.Item("Classification").Equals(DBNull.Value)) Then m_strClassification = pdrRow.Item("Classification").ToString
            If Not (pdrRow.Item("Jobcategory").Equals(DBNull.Value)) Then m_strJobcategory = pdrRow.Item("Jobcategory").ToString
            If Not (pdrRow.Item("Adjustedhiredate").Equals(DBNull.Value)) Then m_dtAdjustedhiredate = pdrRow.Item("Adjustedhiredate")
            If Not (pdrRow.Item("Portalcreator").Equals(DBNull.Value)) Then m_intPortalcreator = pdrRow.Item("Portalcreator")

        Catch ex As Exception
            MsgBox("Error in cHumres." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Sub


    'Private sub SaveLine(ByRef pdrRow As DataRow) 

    '    Try

    '         pdrRow.Item("Id") = m_intId 
    '         pdrRow.Item("Res_Id") = m_intRes_Id 
    '         pdrRow.Item("Fullname") = m_strFullname 
    '         pdrRow.Item("Sur_Name") = m_strSur_Name 
    '         pdrRow.Item("First_Name") = m_strFirst_Name 
    '         pdrRow.Item("Middle_Name") = m_strMiddle_Name 
    '         pdrRow.Item("Initialen") = m_strInitialen 
    '         pdrRow.Item("Mv1") = m_strMv1 
    '         pdrRow.Item("Adres1") = m_strAdres1 
    '         pdrRow.Item("Adres2") = m_strAdres2 
    '         pdrRow.Item("Woonpl") = m_strWoonpl 
    '         pdrRow.Item("Postcode") = m_strPostcode 
    '         pdrRow.Item("Land_Iso") = m_strLand_Iso 
    '         pdrRow.Item("Iso_Taalcd") = m_strIso_Taalcd 
    '         pdrRow.Item("Mail") = m_strMail 
    '         pdrRow.Item("Telnr_Prv") = m_strTelnr_Prv 
    '         pdrRow.Item("Telnr_Prv2") = m_strTelnr_Prv2 
    '         pdrRow.Item("Telnr_Werk") = m_strTelnr_Werk 
    '         pdrRow.Item("Telnr_Werk2") = m_strTelnr_Werk2 
    '         pdrRow.Item("Faxnr") = m_strFaxnr 
    '         pdrRow.Item("Kamer") = m_strKamer 
    '         pdrRow.Item("Toestel") = m_strToestel 
    '         pdrRow.Item("Funcnivo") = m_strFuncnivo 
    '         pdrRow.Item("Socsec_Nr") = m_strSocsec_Nr 
    '        if m_dtGeb_Ldat.Year <> 1 Then pdrRow.Item("Geb_Ldat") = m_dtGeb_Ldat  
    '         pdrRow.Item("Geb_Pl") = m_strGeb_Pl 
    '        if m_dtLdatindienst.Year <> 1 Then pdrRow.Item("Ldatindienst") = m_dtLdatindienst  
    '        if m_dtLdatuitdienst.Year <> 1 Then pdrRow.Item("Ldatuitdienst") = m_dtLdatuitdienst  
    '         pdrRow.Item("Usr_Id") = m_strUsr_Id 
    '         pdrRow.Item("Blocked") = m_bBlocked 
    '         pdrRow.Item("Repto_Id") = m_intRepto_Id 
    '         pdrRow.Item("Nat") = m_strNat 
    '         pdrRow.Item("Mar_Stat") = m_strMar_Stat 
    '        if m_dtMar_Date.Year <> 1 Then pdrRow.Item("Mar_Date") = m_dtMar_Date  
    '         pdrRow.Item("Addr_No") = m_strAddr_No 
    '         pdrRow.Item("Email") = m_strEmail 
    '         pdrRow.Item("Job_Title") = m_strJob_Title 
    '         pdrRow.Item("Loc") = m_strLoc 
    '         pdrRow.Item("Comp") = m_strComp 
    '         pdrRow.Item("Task") = m_strTask 
    '         pdrRow.Item("Emp_Type") = m_strEmp_Type 
    '        if m_dtEmp_Statd.Year <> 1 Then pdrRow.Item("Emp_Statd") = m_dtEmp_Statd  
    '         pdrRow.Item("Prob_Per") = m_strProb_Per 
    '         pdrRow.Item("Prob_Pert") = m_strProb_Pert 
    '        if m_dtProb_Endd.Year <> 1 Then pdrRow.Item("Prob_Endd") = m_dtProb_Endd  
    '         pdrRow.Item("Notes") = m_intNotes 
    '         pdrRow.Item("Reason_Resign") = m_strReason_Resign 
    '         pdrRow.Item("Crcard_Type") = m_strCrcard_Type 
    '         pdrRow.Item("Crcard_No") = m_strCrcard_No 
    '        if m_dtCrcard_Expd.Year <> 1 Then pdrRow.Item("Crcard_Expd") = m_dtCrcard_Expd  
    '         pdrRow.Item("Addr_Xtra") = m_strAddr_Xtra 
    '        if m_dtCont_End_Date.Year <> 1 Then pdrRow.Item("Cont_End_Date") = m_dtCont_End_Date  
    '         pdrRow.Item("Workstat") = m_strWorkstat 
    '         pdrRow.Item("Workstat_Addr") = m_strWorkstat_Addr 
    '         pdrRow.Item("Buyer") = m_bBuyer 
    '         pdrRow.Item("Representative") = m_bRepresentative 
    '         pdrRow.Item("Projempl") = m_bProjempl 
    '         pdrRow.Item("Maiden_Name") = m_strMaiden_Name 
    '         pdrRow.Item("Prafd_Code") = m_strPrafd_Code 
    '         pdrRow.Item("Inttar") = m_dblInttar 
    '         pdrRow.Item("Exttar") = m_dblExttar 
    '         pdrRow.Item("Job_Level") = m_intJob_Level 
    '         pdrRow.Item("Payempl") = m_bPayempl 
    '         pdrRow.Item("Identiteit") = m_strIdentiteit 
    '         pdrRow.Item("Kstdr_Code") = m_strKstdr_Code 
    '         pdrRow.Item("Uitk_Code") = m_strUitk_Code 
    '         pdrRow.Item("Extra_Text") = m_strExtra_Text 
    '         pdrRow.Item("Extra_Code") = m_strExtra_Code 
    '         pdrRow.Item("Adv_Code") = m_strAdv_Code 
    '         pdrRow.Item("Slip_Tekst") = m_strSlip_Tekst 
    '         pdrRow.Item("Predcode") = m_strPredcode 
    '         pdrRow.Item("Costcenter") = m_strCostcenter 
    '         pdrRow.Item("Crdnr") = m_strCrdnr 
    '         pdrRow.Item("Bankac_0") = m_strBankac_0 
    '         pdrRow.Item("Bankac_1") = m_strBankac_1 
    '         pdrRow.Item("Valcode") = m_strValcode 
    '         pdrRow.Item("Pur_Limit") = m_dblPur_Limit 
    '         pdrRow.Item("Fte") = m_dblFte 
    '         pdrRow.Item("Picturefilename") = m_strPicturefilename 
    '         pdrRow.Item("Freefield1") = m_strFreefield1 
    '         pdrRow.Item("Freefield2") = m_strFreefield2 
    '         pdrRow.Item("Freefield3") = m_strFreefield3 
    '         pdrRow.Item("Freefield4") = m_strFreefield4 
    '         pdrRow.Item("Freefield5") = m_strFreefield5 
    '        if m_dtFreefield6.Year <> 1 Then pdrRow.Item("Freefield6") = m_dtFreefield6  
    '        if m_dtFreefield7.Year <> 1 Then pdrRow.Item("Freefield7") = m_dtFreefield7  
    '        if m_dtFreefield8.Year <> 1 Then pdrRow.Item("Freefield8") = m_dtFreefield8  
    '        if m_dtFreefield9.Year <> 1 Then pdrRow.Item("Freefield9") = m_dtFreefield9  
    '        if m_dtFreefield10.Year <> 1 Then pdrRow.Item("Freefield10") = m_dtFreefield10  
    '         pdrRow.Item("Freefield11") = m_dblFreefield11 
    '         pdrRow.Item("Freefield12") = m_dblFreefield12 
    '         pdrRow.Item("Freefield13") = m_dblFreefield13 
    '         pdrRow.Item("Freefield14") = m_dblFreefield14 
    '         pdrRow.Item("Freefield15") = m_dblFreefield15 
    '         pdrRow.Item("Freefield16") = m_bFreefield16 
    '         pdrRow.Item("Freefield17") = m_bFreefield17 
    '         pdrRow.Item("Freefield18") = m_bFreefield18 
    '         pdrRow.Item("Freefield19") = m_bFreefield19 
    '         pdrRow.Item("Freefield20") = m_bFreefield20 
    '         pdrRow.Item("Emp_Stat") = m_strEmp_Stat 
    '         pdrRow.Item("Absent") = m_intAbsent 
    '         pdrRow.Item("Assistant") = m_intAssistant 
    '         pdrRow.Item("Cmp_Wwn") = m_guidCmp_Wwn 
    '         pdrRow.Item("Internetac") = m_strInternetac 
    '         pdrRow.Item("Mobile_Short") = m_strMobile_Short 
    '         pdrRow.Item("Rating") = m_strRating 
    '         pdrRow.Item("Usergroup") = m_strUsergroup 
    '        if m_dtSyscreated.Year <> 1 Then pdrRow.Item("Syscreated") = m_dtSyscreated  
    '         pdrRow.Item("Syscreator") = m_intSyscreator 
    '        if m_dtSysmodified.Year <> 1 Then pdrRow.Item("Sysmodified") = m_dtSysmodified  
    '         pdrRow.Item("Sysmodifier") = m_intSysmodifier 
    '         pdrRow.Item("Sysguid") = m_guidSysguid 
    '        'pdrRow.Item("Timestamp") = m_bTimestamp 
    '        'pdrRow.Item("Picture") = m_bPicture 
    '         pdrRow.Item("Btw_Nummer") = m_strBtw_Nummer 
    '         pdrRow.Item("Ispersonalaccount") = m_bIspersonalaccount 
    '         pdrRow.Item("Scale") = m_intScale 
    '        'pdrRow.Item("Signature") = m_bSignature 
    '         pdrRow.Item("Signaturefile") = m_strSignaturefile 
    '         pdrRow.Item("Bankinglimit") = m_dblBankinglimit 
    '         pdrRow.Item("Oldntfullname") = m_strOldntfullname 
    '         pdrRow.Item("Timezone") = m_intTimezone 
    '         pdrRow.Item("Synchronizeexchange") = m_blnSynchronizeexchange 
    '         pdrRow.Item("Statecode") = m_strStatecode 
    '         pdrRow.Item("Childname1") = m_strChildname1 
    '        if m_dtChildbirthdate1.Year <> 1 Then pdrRow.Item("Childbirthdate1") = m_dtChildbirthdate1  
    '         pdrRow.Item("Childname2") = m_strChildname2 
    '        if m_dtChildbirthdate2.Year <> 1 Then pdrRow.Item("Childbirthdate2") = m_dtChildbirthdate2  
    '         pdrRow.Item("Childname3") = m_strChildname3 
    '        if m_dtChildbirthdate3.Year <> 1 Then pdrRow.Item("Childbirthdate3") = m_dtChildbirthdate3  
    '         pdrRow.Item("Childname4") = m_strChildname4 
    '        if m_dtChildbirthdate4.Year <> 1 Then pdrRow.Item("Childbirthdate4") = m_dtChildbirthdate4  
    '         pdrRow.Item("Childname5") = m_strChildname5 
    '        if m_dtChildbirthdate5.Year <> 1 Then pdrRow.Item("Childbirthdate5") = m_dtChildbirthdate5  
    '         pdrRow.Item("Partner") = m_strPartner 
    '        if m_dtPartnerbirthdate.Year <> 1 Then pdrRow.Item("Partnerbirthdate") = m_dtPartnerbirthdate  
    '         pdrRow.Item("Itemcode") = m_strItemcode 
    '         pdrRow.Item("Usr_Id2") = m_strUsr_Id2 
    '         pdrRow.Item("Mainloc") = m_strMainloc 
    '         pdrRow.Item("Prefix") = m_strPrefix 
    '         pdrRow.Item("Affix") = m_strAffix 
    '         pdrRow.Item("Birth_Prefix") = m_strBirth_Prefix 
    '         pdrRow.Item("Birth_Affix") = m_strBirth_Affix 
    '         pdrRow.Item("Predcode2") = m_strPredcode2 
    '         pdrRow.Item("Land_Iso2") = m_strLand_Iso2 
    '         pdrRow.Item("Race") = m_strRace 
    '         pdrRow.Item("Saleslimit") = m_dblSaleslimit 
    '         pdrRow.Item("Invoicelimit") = m_dblInvoicelimit 
    '         pdrRow.Item("Rmalimit") = m_dblRmalimit 
    '         pdrRow.Item("Skypeid") = m_strSkypeid 
    '         pdrRow.Item("Msnid") = m_strMsnid 
    '         pdrRow.Item("Backofficeblocked") = m_bBackofficeblocked 
    '         pdrRow.Item("Processedbybackgroundjob") = m_intProcessedbybackgroundjob 
    '        if m_dtVacancystartdate.Year <> 1 Then pdrRow.Item("Vacancystartdate") = m_dtVacancystartdate  
    '         pdrRow.Item("Applicationstage") = m_strApplicationstage 
    '         pdrRow.Item("Division") = m_intDivision 
    '         pdrRow.Item("Vacancytype") = m_strVacancytype 
    '         pdrRow.Item("Classification") = m_strClassification 
    '         pdrRow.Item("Jobcategory") = m_strJobcategory 
    '        if m_dtAdjustedhiredate.Year <> 1 Then pdrRow.Item("Adjustedhiredate") = m_dtAdjustedhiredate  
    '         pdrRow.Item("Portalcreator") = m_intPortalcreator 

    '    Catch ex as Exception
    '        MsgBox("Error in cHumres." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
    '    End Try

    'End Sub


#End Region


#Region "Public maintenance procedures ####################################"


    ' WE DO NOT ALLOW DELETE ON HUMRES TABLE WITH THIS PROGRAM
    'Public Sub Delete()

    '    Try

    '        Dim db As New cDBA
    '        Dim dt As New DataTable

    '        Dim strSql As String

    '        strSql = _
    '        "DELETE FROM Humres " & _
    '        "WHERE  Id = " & m_intId & " "

    '        db.Execute(strSql)

    '    Catch ex as Exception
    '        MsgBox("Error in cHumres." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
    '    End Try

    'End Sub


    Private Sub Load(ByVal pintID As Integer)

        Try

            Dim db As New cDBA
            Dim dt As New DataTable

            Dim strSql As String
            strSql = _
            "SELECT * " & _
            "FROM   Humres WITH (Nolock) " & _
            "WHERE  Id = " & pintID & " "

            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then
                Call LoadLine(dt.Rows(0))
            End If

        Catch ex As Exception
            MsgBox("Error in cHumres." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Sub


    Private Sub Load(ByVal pstrUsr_Id As String)

        Try

            Dim db As New cDBA
            Dim dt As New DataTable

            Dim strSql As String
            strSql = _
            "SELECT * " & _
            "FROM   Humres WITH (Nolock) " & _
            "WHERE  Usr_Id = '" & pstrUsr_Id & "' "

            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then
                Call LoadLine(dt.Rows(0))
            End If

        Catch ex As Exception
            MsgBox("Error in cHumres." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Sub

    Public Sub Load_By_Res_ID(ByVal pintID As Integer)

        Try

            Dim db As New cDBA
            Dim dt As New DataTable

            Dim strSql As String
            strSql = _
            "SELECT * " & _
            "FROM   Humres WITH (Nolock) " & _
            "WHERE  Res_Id = " & pintID & " "

            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then
                Call LoadLine(dt.Rows(0))
            End If

        Catch ex As Exception
            MsgBox("Error in cHumres." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Sub


    ' WE DO NOT ALLOW SAVE ON HUMRES TABLE WITH THIS PROGRAM
    'Public Sub Save()

    '    Try

    '        Dim db As New cDBA
    '        Dim dt As New DataTable
    '        Dim drRow As DataRow

    '        Dim strSql As String

    '        strSql = _
    '        "SELECT * " & _
    '        "FROM   Humres " & _
    '        "WHERE  Id = " & m_intId & " "

    '        dt = db.DataTable(strSql)

    '        If dt.Rows.Count = 0 Then
    '            drRow = dt.NewRow()
    '        Else
    '            drRow = dt.Rows(0)
    '        End If

    '        Call SaveLine(drRow)

    '        If dt.Rows.Count = 0 Then

    '            db.DBDataTable.Rows.Add(drRow)
    '            Dim cmd As Odbc.OdbcCommandBuilder = New Odbc.OdbcCommandBuilder(db.DBDataAdapter)
    '            db.DBDataAdapter.InsertCommand = cmd.GetInsertCommand

    '        Else

    '            Dim cmd As Odbc.OdbcCommandBuilder = New Odbc.OdbcCommandBuilder(db.DBDataAdapter)
    '            db.DBDataAdapter.UpdateCommand = cmd.GetUpdateCommand

    '        End If

    '        db.DBDataAdapter.Update(db.DBDataTable)

    '    Catch ex as Exception
    '        MsgBox("Error in cHumres." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
    '    End Try

    'End Sub


#End Region


#Region "Public properties ################################################"

    Public Property Id() As Int32
        Get
            Id = m_intId
        End Get
        Set(ByVal value As Int32)
            m_intId = value
        End Set
    End Property

    Public Property Res_Id() As Int32
        Get
            Res_Id = m_intRes_Id
        End Get
        Set(ByVal value As Int32)
            m_intRes_Id = value
        End Set
    End Property

    Public Property Fullname() As String
        Get
            Fullname = m_strFullname
        End Get
        Set(ByVal value As String)
            m_strFullname = value
        End Set
    End Property

    Public Property Sur_Name() As String
        Get
            Sur_Name = m_strSur_Name
        End Get
        Set(ByVal value As String)
            m_strSur_Name = value
        End Set
    End Property

    Public Property First_Name() As String
        Get
            First_Name = m_strFirst_Name
        End Get
        Set(ByVal value As String)
            m_strFirst_Name = value
        End Set
    End Property

    Public Property Middle_Name() As String
        Get
            Middle_Name = m_strMiddle_Name
        End Get
        Set(ByVal value As String)
            m_strMiddle_Name = value
        End Set
    End Property

    Public Property Initialen() As String
        Get
            Initialen = m_strInitialen
        End Get
        Set(ByVal value As String)
            m_strInitialen = value
        End Set
    End Property

    Public Property Mv1() As String
        Get
            Mv1 = m_strMv1
        End Get
        Set(ByVal value As String)
            m_strMv1 = value
        End Set
    End Property

    Public Property Adres1() As String
        Get
            Adres1 = m_strAdres1
        End Get
        Set(ByVal value As String)
            m_strAdres1 = value
        End Set
    End Property

    Public Property Adres2() As String
        Get
            Adres2 = m_strAdres2
        End Get
        Set(ByVal value As String)
            m_strAdres2 = value
        End Set
    End Property

    Public Property Woonpl() As String
        Get
            Woonpl = m_strWoonpl
        End Get
        Set(ByVal value As String)
            m_strWoonpl = value
        End Set
    End Property

    Public Property Postcode() As String
        Get
            Postcode = m_strPostcode
        End Get
        Set(ByVal value As String)
            m_strPostcode = value
        End Set
    End Property

    Public Property Land_Iso() As String
        Get
            Land_Iso = m_strLand_Iso
        End Get
        Set(ByVal value As String)
            m_strLand_Iso = value
        End Set
    End Property

    Public Property Iso_Taalcd() As String
        Get
            Iso_Taalcd = m_strIso_Taalcd
        End Get
        Set(ByVal value As String)
            m_strIso_Taalcd = value
        End Set
    End Property

    Public Property Mail() As String
        Get
            Mail = m_strMail
        End Get
        Set(ByVal value As String)
            m_strMail = value
        End Set
    End Property

    Public Property Telnr_Prv() As String
        Get
            Telnr_Prv = m_strTelnr_Prv
        End Get
        Set(ByVal value As String)
            m_strTelnr_Prv = value
        End Set
    End Property

    Public Property Telnr_Prv2() As String
        Get
            Telnr_Prv2 = m_strTelnr_Prv2
        End Get
        Set(ByVal value As String)
            m_strTelnr_Prv2 = value
        End Set
    End Property

    Public Property Telnr_Werk() As String
        Get
            Telnr_Werk = m_strTelnr_Werk
        End Get
        Set(ByVal value As String)
            m_strTelnr_Werk = value
        End Set
    End Property

    Public Property Telnr_Werk2() As String
        Get
            Telnr_Werk2 = m_strTelnr_Werk2
        End Get
        Set(ByVal value As String)
            m_strTelnr_Werk2 = value
        End Set
    End Property

    Public Property Faxnr() As String
        Get
            Faxnr = m_strFaxnr
        End Get
        Set(ByVal value As String)
            m_strFaxnr = value
        End Set
    End Property

    Public Property Kamer() As String
        Get
            Kamer = m_strKamer
        End Get
        Set(ByVal value As String)
            m_strKamer = value
        End Set
    End Property

    Public Property Toestel() As String
        Get
            Toestel = m_strToestel
        End Get
        Set(ByVal value As String)
            m_strToestel = value
        End Set
    End Property

    Public Property Funcnivo() As String
        Get
            Funcnivo = m_strFuncnivo
        End Get
        Set(ByVal value As String)
            m_strFuncnivo = value
        End Set
    End Property

    Public Property Socsec_Nr() As String
        Get
            Socsec_Nr = m_strSocsec_Nr
        End Get
        Set(ByVal value As String)
            m_strSocsec_Nr = value
        End Set
    End Property

    Public Property Geb_Ldat() As DateTime
        Get
            Geb_Ldat = m_dtGeb_Ldat
        End Get
        Set(ByVal value As DateTime)
            m_dtGeb_Ldat = value
        End Set
    End Property

    Public Property Geb_Pl() As String
        Get
            Geb_Pl = m_strGeb_Pl
        End Get
        Set(ByVal value As String)
            m_strGeb_Pl = value
        End Set
    End Property

    Public Property Ldatindienst() As DateTime
        Get
            Ldatindienst = m_dtLdatindienst
        End Get
        Set(ByVal value As DateTime)
            m_dtLdatindienst = value
        End Set
    End Property

    Public Property Ldatuitdienst() As DateTime
        Get
            Ldatuitdienst = m_dtLdatuitdienst
        End Get
        Set(ByVal value As DateTime)
            m_dtLdatuitdienst = value
        End Set
    End Property

    Public Property Usr_Id() As String
        Get
            Usr_Id = m_strUsr_Id
        End Get
        Set(ByVal value As String)
            m_strUsr_Id = value
        End Set
    End Property

    Public Property Blocked() As Byte
        Get
            Blocked = m_bBlocked
        End Get
        Set(ByVal value As Byte)
            m_bBlocked = value
        End Set
    End Property

    Public Property Repto_Id() As Int32
        Get
            Repto_Id = m_intRepto_Id
        End Get
        Set(ByVal value As Int32)
            m_intRepto_Id = value
        End Set
    End Property

    Public Property Nat() As String
        Get
            Nat = m_strNat
        End Get
        Set(ByVal value As String)
            m_strNat = value
        End Set
    End Property

    Public Property Mar_Stat() As String
        Get
            Mar_Stat = m_strMar_Stat
        End Get
        Set(ByVal value As String)
            m_strMar_Stat = value
        End Set
    End Property

    Public Property Mar_Date() As DateTime
        Get
            Mar_Date = m_dtMar_Date
        End Get
        Set(ByVal value As DateTime)
            m_dtMar_Date = value
        End Set
    End Property

    Public Property Addr_No() As String
        Get
            Addr_No = m_strAddr_No
        End Get
        Set(ByVal value As String)
            m_strAddr_No = value
        End Set
    End Property

    Public Property Email() As String
        Get
            Email = m_strEmail
        End Get
        Set(ByVal value As String)
            m_strEmail = value
        End Set
    End Property

    Public Property Job_Title() As String
        Get
            Job_Title = m_strJob_Title
        End Get
        Set(ByVal value As String)
            m_strJob_Title = value
        End Set
    End Property

    Public Property Loc() As String
        Get
            Loc = m_strLoc
        End Get
        Set(ByVal value As String)
            m_strLoc = value
        End Set
    End Property

    Public Property Comp() As String
        Get
            Comp = m_strComp
        End Get
        Set(ByVal value As String)
            m_strComp = value
        End Set
    End Property

    Public Property Task() As String
        Get
            Task = m_strTask
        End Get
        Set(ByVal value As String)
            m_strTask = value
        End Set
    End Property

    Public Property Emp_Type() As String
        Get
            Emp_Type = m_strEmp_Type
        End Get
        Set(ByVal value As String)
            m_strEmp_Type = value
        End Set
    End Property

    Public Property Emp_Statd() As DateTime
        Get
            Emp_Statd = m_dtEmp_Statd
        End Get
        Set(ByVal value As DateTime)
            m_dtEmp_Statd = value
        End Set
    End Property

    Public Property Prob_Per() As String
        Get
            Prob_Per = m_strProb_Per
        End Get
        Set(ByVal value As String)
            m_strProb_Per = value
        End Set
    End Property

    Public Property Prob_Pert() As String
        Get
            Prob_Pert = m_strProb_Pert
        End Get
        Set(ByVal value As String)
            m_strProb_Pert = value
        End Set
    End Property

    Public Property Prob_Endd() As DateTime
        Get
            Prob_Endd = m_dtProb_Endd
        End Get
        Set(ByVal value As DateTime)
            m_dtProb_Endd = value
        End Set
    End Property

    Public Property Notes() As Int32
        Get
            Notes = m_intNotes
        End Get
        Set(ByVal value As Int32)
            m_intNotes = value
        End Set
    End Property

    Public Property Reason_Resign() As String
        Get
            Reason_Resign = m_strReason_Resign
        End Get
        Set(ByVal value As String)
            m_strReason_Resign = value
        End Set
    End Property

    Public Property Crcard_Type() As String
        Get
            Crcard_Type = m_strCrcard_Type
        End Get
        Set(ByVal value As String)
            m_strCrcard_Type = value
        End Set
    End Property

    Public Property Crcard_No() As String
        Get
            Crcard_No = m_strCrcard_No
        End Get
        Set(ByVal value As String)
            m_strCrcard_No = value
        End Set
    End Property

    Public Property Crcard_Expd() As DateTime
        Get
            Crcard_Expd = m_dtCrcard_Expd
        End Get
        Set(ByVal value As DateTime)
            m_dtCrcard_Expd = value
        End Set
    End Property

    Public Property Addr_Xtra() As String
        Get
            Addr_Xtra = m_strAddr_Xtra
        End Get
        Set(ByVal value As String)
            m_strAddr_Xtra = value
        End Set
    End Property

    Public Property Cont_End_Date() As DateTime
        Get
            Cont_End_Date = m_dtCont_End_Date
        End Get
        Set(ByVal value As DateTime)
            m_dtCont_End_Date = value
        End Set
    End Property

    Public Property Workstat() As String
        Get
            Workstat = m_strWorkstat
        End Get
        Set(ByVal value As String)
            m_strWorkstat = value
        End Set
    End Property

    Public Property Workstat_Addr() As String
        Get
            Workstat_Addr = m_strWorkstat_Addr
        End Get
        Set(ByVal value As String)
            m_strWorkstat_Addr = value
        End Set
    End Property

    Public Property Buyer() As Byte
        Get
            Buyer = m_bBuyer
        End Get
        Set(ByVal value As Byte)
            m_bBuyer = value
        End Set
    End Property

    Public Property Representative() As Byte
        Get
            Representative = m_bRepresentative
        End Get
        Set(ByVal value As Byte)
            m_bRepresentative = value
        End Set
    End Property

    Public Property Projempl() As Byte
        Get
            Projempl = m_bProjempl
        End Get
        Set(ByVal value As Byte)
            m_bProjempl = value
        End Set
    End Property

    Public Property Maiden_Name() As String
        Get
            Maiden_Name = m_strMaiden_Name
        End Get
        Set(ByVal value As String)
            m_strMaiden_Name = value
        End Set
    End Property

    Public Property Prafd_Code() As String
        Get
            Prafd_Code = m_strPrafd_Code
        End Get
        Set(ByVal value As String)
            m_strPrafd_Code = value
        End Set
    End Property

    Public Property Inttar() As Double
        Get
            Inttar = m_dblInttar
        End Get
        Set(ByVal value As Double)
            m_dblInttar = value
        End Set
    End Property

    Public Property Exttar() As Double
        Get
            Exttar = m_dblExttar
        End Get
        Set(ByVal value As Double)
            m_dblExttar = value
        End Set
    End Property

    Public Property Job_Level() As Int32
        Get
            Job_Level = m_intJob_Level
        End Get
        Set(ByVal value As Int32)
            m_intJob_Level = value
        End Set
    End Property

    Public Property Payempl() As Byte
        Get
            Payempl = m_bPayempl
        End Get
        Set(ByVal value As Byte)
            m_bPayempl = value
        End Set
    End Property

    Public Property Identiteit() As String
        Get
            Identiteit = m_strIdentiteit
        End Get
        Set(ByVal value As String)
            m_strIdentiteit = value
        End Set
    End Property

    Public Property Kstdr_Code() As String
        Get
            Kstdr_Code = m_strKstdr_Code
        End Get
        Set(ByVal value As String)
            m_strKstdr_Code = value
        End Set
    End Property

    Public Property Uitk_Code() As String
        Get
            Uitk_Code = m_strUitk_Code
        End Get
        Set(ByVal value As String)
            m_strUitk_Code = value
        End Set
    End Property

    Public Property Extra_Text() As String
        Get
            Extra_Text = m_strExtra_Text
        End Get
        Set(ByVal value As String)
            m_strExtra_Text = value
        End Set
    End Property

    Public Property Extra_Code() As String
        Get
            Extra_Code = m_strExtra_Code
        End Get
        Set(ByVal value As String)
            m_strExtra_Code = value
        End Set
    End Property

    Public Property Adv_Code() As String
        Get
            Adv_Code = m_strAdv_Code
        End Get
        Set(ByVal value As String)
            m_strAdv_Code = value
        End Set
    End Property

    Public Property Slip_Tekst() As String
        Get
            Slip_Tekst = m_strSlip_Tekst
        End Get
        Set(ByVal value As String)
            m_strSlip_Tekst = value
        End Set
    End Property

    Public Property Predcode() As String
        Get
            Predcode = m_strPredcode
        End Get
        Set(ByVal value As String)
            m_strPredcode = value
        End Set
    End Property

    Public Property Costcenter() As String
        Get
            Costcenter = m_strCostcenter
        End Get
        Set(ByVal value As String)
            m_strCostcenter = value
        End Set
    End Property

    Public Property Crdnr() As String
        Get
            Crdnr = m_strCrdnr
        End Get
        Set(ByVal value As String)
            m_strCrdnr = value
        End Set
    End Property

    Public Property Bankac_0() As String
        Get
            Bankac_0 = m_strBankac_0
        End Get
        Set(ByVal value As String)
            m_strBankac_0 = value
        End Set
    End Property

    Public Property Bankac_1() As String
        Get
            Bankac_1 = m_strBankac_1
        End Get
        Set(ByVal value As String)
            m_strBankac_1 = value
        End Set
    End Property

    Public Property Valcode() As String
        Get
            Valcode = m_strValcode
        End Get
        Set(ByVal value As String)
            m_strValcode = value
        End Set
    End Property

    Public Property Pur_Limit() As Double
        Get
            Pur_Limit = m_dblPur_Limit
        End Get
        Set(ByVal value As Double)
            m_dblPur_Limit = value
        End Set
    End Property

    Public Property Fte() As Double
        Get
            Fte = m_dblFte
        End Get
        Set(ByVal value As Double)
            m_dblFte = value
        End Set
    End Property

    Public Property Picturefilename() As String
        Get
            Picturefilename = m_strPicturefilename
        End Get
        Set(ByVal value As String)
            m_strPicturefilename = value
        End Set
    End Property

    Public Property Freefield1() As String
        Get
            Freefield1 = m_strFreefield1
        End Get
        Set(ByVal value As String)
            m_strFreefield1 = value
        End Set
    End Property

    Public Property Freefield2() As String
        Get
            Freefield2 = m_strFreefield2
        End Get
        Set(ByVal value As String)
            m_strFreefield2 = value
        End Set
    End Property

    Public Property Freefield3() As String
        Get
            Freefield3 = m_strFreefield3
        End Get
        Set(ByVal value As String)
            m_strFreefield3 = value
        End Set
    End Property

    Public Property Freefield4() As String
        Get
            Freefield4 = m_strFreefield4
        End Get
        Set(ByVal value As String)
            m_strFreefield4 = value
        End Set
    End Property

    Public Property Freefield5() As String
        Get
            Freefield5 = m_strFreefield5
        End Get
        Set(ByVal value As String)
            m_strFreefield5 = value
        End Set
    End Property

    Public Property Freefield6() As DateTime
        Get
            Freefield6 = m_dtFreefield6
        End Get
        Set(ByVal value As DateTime)
            m_dtFreefield6 = value
        End Set
    End Property

    Public Property Freefield7() As DateTime
        Get
            Freefield7 = m_dtFreefield7
        End Get
        Set(ByVal value As DateTime)
            m_dtFreefield7 = value
        End Set
    End Property

    Public Property Freefield8() As DateTime
        Get
            Freefield8 = m_dtFreefield8
        End Get
        Set(ByVal value As DateTime)
            m_dtFreefield8 = value
        End Set
    End Property

    Public Property Freefield9() As DateTime
        Get
            Freefield9 = m_dtFreefield9
        End Get
        Set(ByVal value As DateTime)
            m_dtFreefield9 = value
        End Set
    End Property

    Public Property Freefield10() As DateTime
        Get
            Freefield10 = m_dtFreefield10
        End Get
        Set(ByVal value As DateTime)
            m_dtFreefield10 = value
        End Set
    End Property

    Public Property Freefield11() As Double
        Get
            Freefield11 = m_dblFreefield11
        End Get
        Set(ByVal value As Double)
            m_dblFreefield11 = value
        End Set
    End Property

    Public Property Freefield12() As Double
        Get
            Freefield12 = m_dblFreefield12
        End Get
        Set(ByVal value As Double)
            m_dblFreefield12 = value
        End Set
    End Property

    Public Property Freefield13() As Double
        Get
            Freefield13 = m_dblFreefield13
        End Get
        Set(ByVal value As Double)
            m_dblFreefield13 = value
        End Set
    End Property

    Public Property Freefield14() As Double
        Get
            Freefield14 = m_dblFreefield14
        End Get
        Set(ByVal value As Double)
            m_dblFreefield14 = value
        End Set
    End Property

    Public Property Freefield15() As Double
        Get
            Freefield15 = m_dblFreefield15
        End Get
        Set(ByVal value As Double)
            m_dblFreefield15 = value
        End Set
    End Property

    Public Property Freefield16() As Byte
        Get
            Freefield16 = m_bFreefield16
        End Get
        Set(ByVal value As Byte)
            m_bFreefield16 = value
        End Set
    End Property

    Public Property Freefield17() As Byte
        Get
            Freefield17 = m_bFreefield17
        End Get
        Set(ByVal value As Byte)
            m_bFreefield17 = value
        End Set
    End Property

    Public Property Freefield18() As Byte
        Get
            Freefield18 = m_bFreefield18
        End Get
        Set(ByVal value As Byte)
            m_bFreefield18 = value
        End Set
    End Property

    Public Property Freefield19() As Byte
        Get
            Freefield19 = m_bFreefield19
        End Get
        Set(ByVal value As Byte)
            m_bFreefield19 = value
        End Set
    End Property

    Public Property Freefield20() As Byte
        Get
            Freefield20 = m_bFreefield20
        End Get
        Set(ByVal value As Byte)
            m_bFreefield20 = value
        End Set
    End Property

    Public Property Emp_Stat() As String
        Get
            Emp_Stat = m_strEmp_Stat
        End Get
        Set(ByVal value As String)
            m_strEmp_Stat = value
        End Set
    End Property

    Public Property Absent() As Int32
        Get
            Absent = m_intAbsent
        End Get
        Set(ByVal value As Int32)
            m_intAbsent = value
        End Set
    End Property

    Public Property Assistant() As Int32
        Get
            Assistant = m_intAssistant
        End Get
        Set(ByVal value As Int32)
            m_intAssistant = value
        End Set
    End Property

    Public Property Cmp_Wwn() As Guid
        Get
            Cmp_Wwn = m_guidCmp_Wwn
        End Get
        Set(ByVal value As Guid)
            m_guidCmp_Wwn = value
        End Set
    End Property

    Public Property Internetac() As String
        Get
            Internetac = m_strInternetac
        End Get
        Set(ByVal value As String)
            m_strInternetac = value
        End Set
    End Property

    Public Property Mobile_Short() As String
        Get
            Mobile_Short = m_strMobile_Short
        End Get
        Set(ByVal value As String)
            m_strMobile_Short = value
        End Set
    End Property

    Public Property Rating() As String
        Get
            Rating = m_strRating
        End Get
        Set(ByVal value As String)
            m_strRating = value
        End Set
    End Property

    Public Property Usergroup() As String
        Get
            Usergroup = m_strUsergroup
        End Get
        Set(ByVal value As String)
            m_strUsergroup = value
        End Set
    End Property

    Public Property Syscreated() As DateTime
        Get
            Syscreated = m_dtSyscreated
        End Get
        Set(ByVal value As DateTime)
            m_dtSyscreated = value
        End Set
    End Property

    Public Property Syscreator() As Int32
        Get
            Syscreator = m_intSyscreator
        End Get
        Set(ByVal value As Int32)
            m_intSyscreator = value
        End Set
    End Property

    Public Property Sysmodified() As DateTime
        Get
            Sysmodified = m_dtSysmodified
        End Get
        Set(ByVal value As DateTime)
            m_dtSysmodified = value
        End Set
    End Property

    Public Property Sysmodifier() As Int32
        Get
            Sysmodifier = m_intSysmodifier
        End Get
        Set(ByVal value As Int32)
            m_intSysmodifier = value
        End Set
    End Property

    Public Property Sysguid() As Guid
        Get
            Sysguid = m_guidSysguid
        End Get
        Set(ByVal value As Guid)
            m_guidSysguid = value
        End Set
    End Property

    'Public Property Timestamp() As Byte[]
    '    Get
    '        Timestamp = m_bTimestamp 
    '    End Get
    '    Set(ByVal value As Byte[])
    '        m_bTimestamp = value
    '    End Set
    'End Property

    'Public Property Picture() As Byte[]
    '    Get
    '        Picture = m_bPicture 
    '    End Get
    '    Set(ByVal value As Byte[])
    '        m_bPicture = value
    '    End Set
    'End Property

    Public Property Btw_Nummer() As String
        Get
            Btw_Nummer = m_strBtw_Nummer
        End Get
        Set(ByVal value As String)
            m_strBtw_Nummer = value
        End Set
    End Property

    Public Property Ispersonalaccount() As Byte
        Get
            Ispersonalaccount = m_bIspersonalaccount
        End Get
        Set(ByVal value As Byte)
            m_bIspersonalaccount = value
        End Set
    End Property

    Public Property Scale() As Int32
        Get
            Scale = m_intScale
        End Get
        Set(ByVal value As Int32)
            m_intScale = value
        End Set
    End Property

    'Public Property Signature() As Byte[]
    '    Get
    '        Signature = m_bSignature 
    '    End Get
    '    Set(ByVal value As Byte[])
    '        m_bSignature = value
    '    End Set
    'End Property

    Public Property Signaturefile() As String
        Get
            Signaturefile = m_strSignaturefile
        End Get
        Set(ByVal value As String)
            m_strSignaturefile = value
        End Set
    End Property

    Public Property Bankinglimit() As Double
        Get
            Bankinglimit = m_dblBankinglimit
        End Get
        Set(ByVal value As Double)
            m_dblBankinglimit = value
        End Set
    End Property

    Public Property Oldntfullname() As String
        Get
            Oldntfullname = m_strOldntfullname
        End Get
        Set(ByVal value As String)
            m_strOldntfullname = value
        End Set
    End Property

    Public Property Timezone() As Int32
        Get
            Timezone = m_intTimezone
        End Get
        Set(ByVal value As Int32)
            m_intTimezone = value
        End Set
    End Property

    Public Property Synchronizeexchange() As Boolean
        Get
            Synchronizeexchange = m_blnSynchronizeexchange
        End Get
        Set(ByVal value As Boolean)
            m_blnSynchronizeexchange = value
        End Set
    End Property

    Public Property Statecode() As String
        Get
            Statecode = m_strStatecode
        End Get
        Set(ByVal value As String)
            m_strStatecode = value
        End Set
    End Property

    Public Property Childname1() As String
        Get
            Childname1 = m_strChildname1
        End Get
        Set(ByVal value As String)
            m_strChildname1 = value
        End Set
    End Property

    Public Property Childbirthdate1() As DateTime
        Get
            Childbirthdate1 = m_dtChildbirthdate1
        End Get
        Set(ByVal value As DateTime)
            m_dtChildbirthdate1 = value
        End Set
    End Property

    Public Property Childname2() As String
        Get
            Childname2 = m_strChildname2
        End Get
        Set(ByVal value As String)
            m_strChildname2 = value
        End Set
    End Property

    Public Property Childbirthdate2() As DateTime
        Get
            Childbirthdate2 = m_dtChildbirthdate2
        End Get
        Set(ByVal value As DateTime)
            m_dtChildbirthdate2 = value
        End Set
    End Property

    Public Property Childname3() As String
        Get
            Childname3 = m_strChildname3
        End Get
        Set(ByVal value As String)
            m_strChildname3 = value
        End Set
    End Property

    Public Property Childbirthdate3() As DateTime
        Get
            Childbirthdate3 = m_dtChildbirthdate3
        End Get
        Set(ByVal value As DateTime)
            m_dtChildbirthdate3 = value
        End Set
    End Property

    Public Property Childname4() As String
        Get
            Childname4 = m_strChildname4
        End Get
        Set(ByVal value As String)
            m_strChildname4 = value
        End Set
    End Property

    Public Property Childbirthdate4() As DateTime
        Get
            Childbirthdate4 = m_dtChildbirthdate4
        End Get
        Set(ByVal value As DateTime)
            m_dtChildbirthdate4 = value
        End Set
    End Property

    Public Property Childname5() As String
        Get
            Childname5 = m_strChildname5
        End Get
        Set(ByVal value As String)
            m_strChildname5 = value
        End Set
    End Property

    Public Property Childbirthdate5() As DateTime
        Get
            Childbirthdate5 = m_dtChildbirthdate5
        End Get
        Set(ByVal value As DateTime)
            m_dtChildbirthdate5 = value
        End Set
    End Property

    Public Property Partner() As String
        Get
            Partner = m_strPartner
        End Get
        Set(ByVal value As String)
            m_strPartner = value
        End Set
    End Property

    Public Property Partnerbirthdate() As DateTime
        Get
            Partnerbirthdate = m_dtPartnerbirthdate
        End Get
        Set(ByVal value As DateTime)
            m_dtPartnerbirthdate = value
        End Set
    End Property

    Public Property Itemcode() As String
        Get
            Itemcode = m_strItemcode
        End Get
        Set(ByVal value As String)
            m_strItemcode = value
        End Set
    End Property

    Public Property Usr_Id2() As String
        Get
            Usr_Id2 = m_strUsr_Id2
        End Get
        Set(ByVal value As String)
            m_strUsr_Id2 = value
        End Set
    End Property

    Public Property Mainloc() As String
        Get
            Mainloc = m_strMainloc
        End Get
        Set(ByVal value As String)
            m_strMainloc = value
        End Set
    End Property

    Public Property Prefix() As String
        Get
            Prefix = m_strPrefix
        End Get
        Set(ByVal value As String)
            m_strPrefix = value
        End Set
    End Property

    Public Property Affix() As String
        Get
            Affix = m_strAffix
        End Get
        Set(ByVal value As String)
            m_strAffix = value
        End Set
    End Property

    Public Property Birth_Prefix() As String
        Get
            Birth_Prefix = m_strBirth_Prefix
        End Get
        Set(ByVal value As String)
            m_strBirth_Prefix = value
        End Set
    End Property

    Public Property Birth_Affix() As String
        Get
            Birth_Affix = m_strBirth_Affix
        End Get
        Set(ByVal value As String)
            m_strBirth_Affix = value
        End Set
    End Property

    Public Property Predcode2() As String
        Get
            Predcode2 = m_strPredcode2
        End Get
        Set(ByVal value As String)
            m_strPredcode2 = value
        End Set
    End Property

    Public Property Land_Iso2() As String
        Get
            Land_Iso2 = m_strLand_Iso2
        End Get
        Set(ByVal value As String)
            m_strLand_Iso2 = value
        End Set
    End Property

    Public Property Race() As String
        Get
            Race = m_strRace
        End Get
        Set(ByVal value As String)
            m_strRace = value
        End Set
    End Property

    Public Property Saleslimit() As Double
        Get
            Saleslimit = m_dblSaleslimit
        End Get
        Set(ByVal value As Double)
            m_dblSaleslimit = value
        End Set
    End Property

    Public Property Invoicelimit() As Double
        Get
            Invoicelimit = m_dblInvoicelimit
        End Get
        Set(ByVal value As Double)
            m_dblInvoicelimit = value
        End Set
    End Property

    Public Property Rmalimit() As Double
        Get
            Rmalimit = m_dblRmalimit
        End Get
        Set(ByVal value As Double)
            m_dblRmalimit = value
        End Set
    End Property

    Public Property Skypeid() As String
        Get
            Skypeid = m_strSkypeid
        End Get
        Set(ByVal value As String)
            m_strSkypeid = value
        End Set
    End Property

    Public Property Msnid() As String
        Get
            Msnid = m_strMsnid
        End Get
        Set(ByVal value As String)
            m_strMsnid = value
        End Set
    End Property

    Public Property Backofficeblocked() As Byte
        Get
            Backofficeblocked = m_bBackofficeblocked
        End Get
        Set(ByVal value As Byte)
            m_bBackofficeblocked = value
        End Set
    End Property

    Public Property Processedbybackgroundjob() As Int16
        Get
            Processedbybackgroundjob = m_intProcessedbybackgroundjob
        End Get
        Set(ByVal value As Int16)
            m_intProcessedbybackgroundjob = value
        End Set
    End Property

    Public Property Vacancystartdate() As DateTime
        Get
            Vacancystartdate = m_dtVacancystartdate
        End Get
        Set(ByVal value As DateTime)
            m_dtVacancystartdate = value
        End Set
    End Property

    Public Property Applicationstage() As String
        Get
            Applicationstage = m_strApplicationstage
        End Get
        Set(ByVal value As String)
            m_strApplicationstage = value
        End Set
    End Property

    Public Property Division() As Int16
        Get
            Division = m_intDivision
        End Get
        Set(ByVal value As Int16)
            m_intDivision = value
        End Set
    End Property

    Public Property Vacancytype() As String
        Get
            Vacancytype = m_strVacancytype
        End Get
        Set(ByVal value As String)
            m_strVacancytype = value
        End Set
    End Property

    Public Property Classification() As String
        Get
            Classification = m_strClassification
        End Get
        Set(ByVal value As String)
            m_strClassification = value
        End Set
    End Property

    Public Property Jobcategory() As String
        Get
            Jobcategory = m_strJobcategory
        End Get
        Set(ByVal value As String)
            m_strJobcategory = value
        End Set
    End Property

    Public Property Adjustedhiredate() As DateTime
        Get
            Adjustedhiredate = m_dtAdjustedhiredate
        End Get
        Set(ByVal value As DateTime)
            m_dtAdjustedhiredate = value
        End Set
    End Property

    Public Property Portalcreator() As Int32
        Get
            Portalcreator = m_intPortalcreator
        End Get
        Set(ByVal value As Int32)
            m_intPortalcreator = value
        End Set
    End Property

#End Region


End Class ' cHumres


