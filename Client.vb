'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Client Master Object
'
' Copyright 2016 and Beyond
' All Rights Reserved
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
' €  All  rights reserved. No part of this  software  €€  This Software is Owned by        €
' €  may be reproduced or transmitted in any form or  €€                                   €
' €  by   any   means,  electronic   or  mechanical,  €€    GUANZON MERCHANDISING CORP.    €
' €  including recording, or by information  storage  €€     Guanzon Bldg. Perez Blvd.     €
' €  and  retrieval  systems, without  prior written  €€           Dagupan City            €
' €  from the author.                                 €€  Tel No. 522-1085 ; 522-9275      €
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'
' ==========================================================================================
'  Kalyptus [ 04/07/2016 10:15 am ]
'      Started creating this object.
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€

'Please check the following URL to check the order in firing of events for the textbox control
' https://msdn.microsoft.com/en-us/library/system.windows.forms.control.lostfocus(v=vs.110).aspx
'
'Note: 
'  1. Usage of DataTable.Rows.IndexOf(DataRow)
'
'Usage:
'   Imports ggcClient
'
'   Private p_oClient As Client
'   p_oClient = New Client(p_oAppDriver)
'   p_oClient.SearchClient("Sayson, Josh Ramsej")
'   If p_oClient.ShowClient() Then
'       p_oClient.SaveClient()
'   End If
'

Imports MySql.Data.MySqlClient
Imports ADODB
Imports ggcAppDriver

Public Class Client
    Implements ICloneable

    Private p_oApp As GRider
    Private p_oDTMstr As DataTable
    Private p_oDTMstr_Old As DataTable
    Private p_nEditMode As xeEditMode
    Private p_oOthersx As New Others
    Private p_sParent As String

    Private Const p_sMasTable As String = "Client_Master"
    Private Const p_sMsgHeadr As String = "Client Info Maintenance"

    Public Event MasterRetrieved(ByVal Index As Integer, _
                                  ByVal Value As Object)

    Private Const xeMOBILE_LEN = 11

    Public Property Master(ByVal Index As Integer) As Object
        Get
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                Select Case Index
                    Case 80 ' xBirthPlc
                        If Trim(IFNull(p_oDTMstr(0).Item(10))) <> "" And Trim(p_oOthersx.xBirthPlc) = "" Then
                            getBirthPlace(10, 80, p_oDTMstr(0).Item(10), True, False)
                        End If
                        Return p_oOthersx.xBirthPlc
                    Case 81 ' sTownName
                        If Trim(IFNull(p_oDTMstr(0).Item(13))) <> "" And Trim(p_oOthersx.sTownName) = "" Then
                            getTown(13, 81, p_oDTMstr(0).Item(13), True, False)
                        End If
                        Return p_oOthersx.sTownName
                    Case 82 ' sBrgyName
                        If Trim(IFNull(p_oDTMstr(0).Item(14))) <> "" And Trim(p_oOthersx.sBrgyName) = "" Then
                            getBarangay(14, 82, p_oDTMstr(0).Item(14), True, False)
                        End If
                        Return p_oOthersx.sBrgyName
                    Case 83 ' sRelgnNme
                        If Trim(IFNull(p_oDTMstr(0).Item(19))) <> "" And Trim(p_oOthersx.sRelgnNme) = "" Then
                            getReligion(19, 83, p_oDTMstr(0).Item(19), True, False)
                        End If
                        Return p_oOthersx.sRelgnNme
                    Case 84 ' sOccptnNm
                        If Trim(IFNull(p_oDTMstr(0).Item(24))) <> "" And Trim(p_oOthersx.sOccptnNm) = "" Then
                            getOccupation(24, 84, p_oDTMstr(0).Item(24), True, False)
                        End If
                        Return p_oOthersx.sOccptnNm
                    Case 85 ' sNational
                        If Trim(IFNull(p_oDTMstr(0).Item(8))) <> "" And Trim(p_oOthersx.sNational) = "" Then
                            getNational(8, 85, p_oDTMstr(0).Item(8), True, False)
                        End If
                        Return p_oOthersx.sNational
                    Case 86 ' sSpouseNm
                        If Trim(IFNull(p_oDTMstr(0).Item(28))) <> "" And Trim(p_oOthersx.sSpouseNm) = "" Then
                            Call getSpouse()
                        End If
                        Return p_oOthersx.sSpouseNm
                    Case Else
                        Return p_oDTMstr(0).Item(Index)
                End Select
            Else
                Return vbEmpty
            End If
        End Get

        Set(ByVal value As Object)
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                Select Case Index
                    Case 80 ' xBirthPlc
                        getBirthPlace(10, 80, value, False, False)
                    Case 81 ' sTownName
                        getTown(13, 81, value, False, False)
                    Case 82 ' sBrgyName
                        getBarangay(14, 82, value, False, False)
                    Case 83 ' sRelgnNme
                        getReligion(19, 83, value, False, False)
                    Case 84 ' sOccptnNm
                        getOccupation(24, 84, value, False, False)
                    Case 85 ' sNational
                        getNational(8, 85, value, False, False)
                    Case 9 ' dBirthDte  
                        If IsDate(value) Then
                            p_oDTMstr(0).Item(Index) = value
                        End If
                        RaiseEvent MasterRetrieved(9, p_oDTMstr(0).Item(Index))
                    Case 15 ' sPhoneNox
                        p_oDTMstr(0).Item(Index) = getValidPhone(value)
                        RaiseEvent MasterRetrieved(15, p_oDTMstr(0).Item(Index))
                    Case 16 ' sMobileNo
                        p_oDTMstr(0).Item(Index) = getValidMobile(value)
                        RaiseEvent MasterRetrieved(16, p_oDTMstr(0).Item(Index))
                    Case 17 ' sEmailAdd
                        If isValidEmail(value) Then
                            p_oDTMstr(0).Item(Index) = value
                        End If
                        RaiseEvent MasterRetrieved(17, p_oDTMstr(0).Item(Index))
                    Case 26 'nGrssIncm
                        If IsNumeric(value) Then
                            p_oDTMstr(0).Item(Index) = value
                        End If
                        RaiseEvent MasterRetrieved(26, p_oDTMstr(0).Item(Index))
                    Case Else
                        p_oDTMstr(0).Item(Index) = value
                End Select
            End If
        End Set
    End Property

    'Property Master(String)
    Public Property Master(ByVal Index As String) As Object
        Get
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                Select Case LCase(Index)
                    Case "xbirthplc" '80 
                        If Trim(IFNull(p_oDTMstr(0).Item(10))) <> "" And Trim(p_oOthersx.xBirthPlc) = "" Then
                            getBirthPlace(10, 80, p_oDTMstr(0).Item(10), True, False)
                        End If
                        Return p_oOthersx.xBirthPlc
                    Case "stownname" ' 81
                        If Trim(IFNull(p_oDTMstr(0).Item(13))) <> "" And Trim(p_oOthersx.sTownName) = "" Then
                            getTown(13, 81, p_oDTMstr(0).Item(13), True, False)
                        End If
                        Return p_oOthersx.sTownName
                    Case "sbrgyname" ' 82 
                        If Trim(IFNull(p_oDTMstr(0).Item(14))) <> "" And Trim(p_oOthersx.sBrgyName) = "" Then
                            getBarangay(14, 82, p_oDTMstr(0).Item(14), True, False)
                        End If
                        Return p_oOthersx.sBrgyName
                    Case "srelgnnme" ' 83  
                        If Trim(IFNull(p_oDTMstr(0).Item(19))) <> "" And Trim(p_oOthersx.sRelgnNme) = "" Then
                            getReligion(19, 83, p_oDTMstr(0).Item(19), True, False)
                        End If
                        Return p_oOthersx.sRelgnNme
                    Case "soccptnnm" ' 84 
                        If Trim(IFNull(p_oDTMstr(0).Item(24))) <> "" And Trim(p_oOthersx.sOccptnNm) = "" Then
                            getOccupation(24, 84, p_oDTMstr(0).Item(24), True, False)
                        End If
                        Return p_oOthersx.sOccptnNm
                    Case "snational" ' 85 
                        If Trim(IFNull(p_oDTMstr(0).Item(8))) <> "" And Trim(p_oOthersx.sNational) = "" Then
                            getOccupation(8, 85, p_oDTMstr(0).Item(8), True, False)
                        End If
                        Return p_oOthersx.sNational
                    Case "sspousenm" ' 86  
                        If Trim(IFNull(p_oDTMstr(0).Item(28))) <> "" And Trim(p_oOthersx.sSpouseNm) = "" Then
                            Call getSpouse()
                        End If
                        Return p_oOthersx.sSpouseNm
                    Case Else
                        Return p_oDTMstr(0).Item(Index)
                End Select
            Else
                Return vbEmpty
            End If
        End Get

        Set(ByVal value As Object)
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                Select Case LCase(Index)
                    Case "xbirthplc" ' 80  
                        getBirthPlace(10, 80, value, False, False)
                    Case "stownname" ' 81  
                        getTown(13, 81, value, False, False)
                    Case "sbrgyname" ' 82  
                        getBarangay(14, 82, value, False, False)
                    Case "srelgnnme" ' 83  
                        getReligion(19, 83, value, False, False)
                    Case "soccptnnm" ' 84  
                        getOccupation(24, 84, value, False, False)
                    Case "snational" ' 85  
                        getNational(8, 85, value, False, False)
                    Case "dbirthdte" ' 09
                        If IsDate(value) Then
                            p_oDTMstr(0).Item(Index) = value
                        End If
                        RaiseEvent MasterRetrieved(9, p_oDTMstr(0).Item(Index))
                    Case "sphonenox" '15  
                        p_oDTMstr(0).Item(Index) = getValidPhone(value)
                        RaiseEvent MasterRetrieved(15, p_oDTMstr(0).Item(Index))
                    Case "smobileno" '16  
                        p_oDTMstr(0).Item(Index) = getValidMobile(value)
                        RaiseEvent MasterRetrieved(16, p_oDTMstr(0).Item(Index))
                    Case "semailadd" ' 17  
                        If isValidEmail(value) Then
                            p_oDTMstr(0).Item(Index) = value
                        End If
                        RaiseEvent MasterRetrieved(17, p_oDTMstr(0).Item(Index))
                    Case "ngrssincm" ' 26 
                        If IsNumeric(value) Then
                            p_oDTMstr(0).Item(Index) = value
                        End If
                        RaiseEvent MasterRetrieved(26, p_oDTMstr(0).Item(Index))
                    Case Else
                        p_oDTMstr(0).Item(Index) = value
                End Select
            End If
        End Set
    End Property

    Public ReadOnly Property OriginalMaster(ByVal Index As String) As Object
        Get
            If p_nEditMode = xeEditMode.MODE_READY Then
                Return p_oDTMstr(0).Item(Index)
            Else
                Return vbEmpty
            End If
        End Get
    End Property

    Public ReadOnly Property OriginalMaster(ByVal Index As Integer) As Object
        Get
            If p_nEditMode = xeEditMode.MODE_READY Then
                Return p_oDTMstr_Old(0).Item(Index)
            Else
                Return vbEmpty
            End If
        End Get
    End Property

    'Property EditMode()
    Public ReadOnly Property EditMode() As xeEditMode
        Get
            Return p_nEditMode
        End Get
    End Property

    Public Property Parent() As String
        Get
            Return p_sParent
        End Get
        Set(ByVal value As String)
            p_sParent = value
        End Set
    End Property

    'Public Function NewClient()
    Private Function NewClient(ByVal fsLastName As String, ByVal fsFirstName As String) As Boolean
        Dim lsSQL As String

        lsSQL = AddCondition(getSQ_Master, "0=1")
        p_oDTMstr = p_oApp.ExecuteQuery(lsSQL)
        p_oDTMstr.Rows.Add(p_oDTMstr.NewRow())

        p_nEditMode = xeEditMode.MODE_ADDNEW

        Call initMaster()

        p_oDTMstr(0).Item("sLastName") = fsLastName
        p_oDTMstr(0).Item("sFrstName") = fsFirstName

        Call InitOthers()

        Return True
    End Function

    'Public Function OpenClient(String)
    Public Function OpenClient(ByVal fsClientID As String) As Boolean
        Dim lsSQL As String

        lsSQL = AddCondition(getSQ_Master, "a.sClientID = " & strParm(fsClientID))
        p_oDTMstr = p_oApp.ExecuteQuery(lsSQL)

        If p_oDTMstr.Rows.Count <= 0 Then
            p_nEditMode = xeEditMode.MODE_UNKNOWN
            Return False
        End If

        'Copy the structure and data from the previous client record
        p_oDTMstr_Old = p_oDTMstr.Copy()

        Call InitOthers()

        p_nEditMode = xeEditMode.MODE_READY
        Return True
    End Function

    'Public Function SearchClient(String, Boolean, Boolean=False)
    Public Function SearchClient( _
                        ByVal fsValue As String _
                      , Optional ByVal fbByCode As Boolean = False) As Boolean

        Dim lsSQL As String

        'Check if already loaded base on edit mode
        If p_nEditMode = xeEditMode.MODE_READY Or p_nEditMode = xeEditMode.MODE_UPDATE Then
            If fbByCode Then
                If fsValue = p_oDTMstr(0).Item("sClientID") Then Return True
            Else
                If fsValue = p_oDTMstr(0).Item("sCompnyNm") Then Return True
            End If
        End If

        lsSQL = getSQ_Browse()

        'create Kwiksearch filter
        Dim lsFilter As String
        If fbByCode Then
            lsFilter = "a.sClientID LIKE " & strParm("%" & fsValue)
        Else
            If fsValue = "" Then
                Return False
            End If
            lsFilter = "a.sCompnyNm like " & strParm(fsValue & "%")
        End If

        Dim loDta As DataRow = KwikSearch(p_oApp _
                                        , lsSQL _
                                        , True _
                                        , fsValue _
                                        , "sClientID»sClientNm»xAddressx»dBirthDte" _
                                        , "ID»Client»Address»Birth Day", _
                                        , "a.sClientID»a.sCompnyNm»CONCAT(IF(IFNull(a.sHouseNox, '') = '', '', CONCAT(a.sHouseNox, ' ')), a.sAddressx, ', ', b.sTownName, ', ', c.sProvName, ' ', b.sZippCode)»a.dBirthDte" _
                                        , IIf(fbByCode, 0, 1))
        If IsNothing(loDta) Then
            Dim lsValue() As String = Split(fsValue, ", ")
            If UBound(lsValue) > 0 Then
                Return NewClient(lsValue(0), lsValue(1))
            ElseIf UBound(lsValue) = 0 Then
                Return NewClient(lsValue(0), "")
            Else
                Return NewClient("", "")
            End If
        Else
            Return OpenClient(loDta.Item("sClientID"))
        End If
    End Function

    Public Function ShowClient() As Boolean
        Dim loFrm As frmClient
        loFrm = New frmClient
        loFrm.iClient = Me

        loFrm.ShowDialog()

        If loFrm.Cancelled Then
            loFrm.Dispose()
            Return False
        Else
            loFrm.Dispose()
            Return True
        End If
    End Function

    Public Function SaveClient() As Boolean
        If p_nEditMode = xeEditMode.MODE_UNKNOWN Then Return False

        'Check if there are empty required info 
        If Not isEntryOk() Then Return False

        Dim lsSQL As String

        'Reconstruct the Full Name of Client
        p_oDTMstr(0).Item("sCompnyNm") = p_oDTMstr(0).Item("sLastName") & ", " & _
                                         p_oDTMstr(0).Item("sFrstName") & _
                                            IIf(p_oDTMstr(0).Item("sSuffixNm") = "", "", " " & p_oDTMstr(0).Item("sSuffixNm")) & " " & _
                                         p_oDTMstr(0).Item("sMiddName")

        If p_sParent = "" Then p_oApp.BeginTransaction()

        If p_nEditMode = xeEditMode.MODE_ADDNEW Then
            'kalyptus - 2024.09.17 11:41am
            'Include terminal number in sClientID
            p_oDTMstr(0).Item("sClientID") = GetNextCode(p_sMasTable, "sClientID", True, p_oApp.Connection, True, p_oApp.BranchCode + p_oApp.POSTerminal)
            lsSQL = ADO2SQL(p_oDTMstr, p_sMasTable, , p_oApp.UserID, p_oApp.SysDate)

            If p_oApp.Execute(lsSQL, p_sMasTable) = 0 Then
                If p_sParent = "" Then p_oApp.RollBackTransaction()
                Return False
            End If

            'Save Mobile
            If p_oDTMstr(0).Item("sMobileNo") <> "" Then
                lsSQL = "INSERT INTO Client_Mobile" &
                       " SET sClientID = " & strParm(p_oDTMstr(0).Item("sClientID")) &
                          ", nEntryNox = 1" &
                          ", sMobileNo = " & strParm(p_oDTMstr(0).Item("sMobileNo")) &
                          ", nPriority = 1" &
                          ", cIncdMktg = '1'" &
                          ", nNoRetryx = 0" &
                          ", cInvalidx = '0'" &
                          ", cNewMobil = '1'" &
                          ", cRecdStat = '1'"
                If p_oApp.Execute(lsSQL, "Client_Mobile") = 0 Then
                    If p_sParent = "" Then p_oApp.RollBackTransaction()
                    Return False
                End If
            End If

            'Save Telephone
            If p_oDTMstr(0).Item("sPhoneNox") <> "" Then
                lsSQL = "INSERT INTO Client_Telephone" &
                       " SET sClientID = " & strParm(p_oDTMstr(0).Item("sClientID")) &
                          ", nEntryNox = 1" &
                          ", sPhoneNox = " & strParm(p_oDTMstr(0).Item("sPhoneNox")) &
                          ", nPriority = 1" &
                          ", cInvalidx = '0'" &
                          ", cConfirmd = '0'" &
                          ", cRecdStat = '1'"
                If p_oApp.Execute(lsSQL, "Client_Telephone") = 0 Then
                    If p_sParent = "" Then p_oApp.RollBackTransaction()
                    Return False
                End If
            End If

            'Save Email Address
            If p_oDTMstr(0).Item("sEmailAdd") <> "" Then
                lsSQL = "INSERT INTO Client_eMail_Address" &
                       " SET sClientID = " & strParm(p_oDTMstr(0).Item("sClientID")) &
                          ", nEntryNox = 1" &
                          ", sEmailAdd = " & strParm(p_oDTMstr(0).Item("sEmailAdd")) &
                          ", nPriority = 1"
                If p_oApp.Execute(lsSQL, "Client_eMail_Address") = 0 Then
                    If p_sParent = "" Then p_oApp.RollBackTransaction()
                    Return False
                End If
            End If
        Else
            lsSQL = ADO2SQL(p_oDTMstr, p_sMasTable, "sClientID = " & strParm(p_oDTMstr(0).Item("sClientID")), p_oApp.UserID, Format(p_oApp.SysDate, "yyyy-MM-dd"))
            If lsSQL <> "" Then
                If p_oApp.Execute(lsSQL, p_sMasTable) = 0 Then
                    If p_sParent = "" Then p_oApp.RollBackTransaction()
                    Return False
                End If

                'Review the Client_Mobile if there is an entry only...
                If p_oDTMstr(0).Item("sMobileNo") <> p_oDTMstr_Old(0).Item("sMobileNo") Then
                    If Not saveMobile() Then
                        If p_sParent = "" Then p_oApp.RollBackTransaction()
                        Return False
                    End If
                End If

                If p_oDTMstr(0).Item("sPhoneNox") <> p_oDTMstr_Old(0).Item("sPhoneNox") Then
                    If Not saveTelephone() Then
                        If p_sParent = "" Then p_oApp.RollBackTransaction()
                        Return False
                    End If
                End If

                If p_oDTMstr(0).Item("sEmailAdd") <> p_oDTMstr_Old(0).Item("sEmailAdd") Then
                    If Not SaveEmailAdd() Then
                        If p_sParent = "" Then p_oApp.RollBackTransaction()
                        Return False
                    End If
                End If

            End If
        End If

        p_nEditMode = xeEditMode.MODE_READY
        If p_sParent = "" Then p_oApp.CommitTransaction()

        Return True
    End Function

    Private Function saveMobile() As Boolean
        Dim lsSQL As String

        'Load all Mobile No of this Employee
        lsSQL = "SELECT sClientID" & _
                     ", nEntryNox" & _
                     ", sMobileNo" & _
                     ", nPriority" & _
                     ", cIncdMktg" & _
                     ", nNoRetryx" & _
                     ", cInvalidx" & _
                     ", cNewMobil" & _
                     ", cRecdStat" & _
                     ", '0' xNewDatax" & _
               " FROM Client_Mobile" & _
               " WHERE sClientID = " & strParm(p_oDTMstr(0).Item("sClientID"))
        Dim loData = p_oApp.ExecuteQuery(lsSQL)

        'Search for the existence of this Mobile No
        Dim loRow() As DataRow = loData.Select("sMobileNo = " & strParm(p_oDTMstr(0).Item("sMobileNo")))

        Dim lnRow As Integer
        If loRow.Count = 0 Then
            'It seems the mobile no is not existing from this client's record
            lnRow = loData.Rows.Count
            loData.Rows.Add(loData.NewRow())
            loData(lnRow).Item("sClientID") = p_oDTMstr(0).Item("sClientID")
            loData(lnRow).Item("nEntryNox") = lnRow + 1
            loData(lnRow).Item("sMobileNo") = p_oDTMstr(0).Item("sMobileNo")
            loData(lnRow).Item("nPriority") = 0
            loData(lnRow).Item("cIncdMktg") = "1"
            loData(lnRow).Item("nNoRetryx") = 0
            loData(lnRow).Item("cInvalidx") = "0"
            loData(lnRow).Item("cNewMobil") = "1"
            loData(lnRow).Item("cRecdStat") = "1"
            loData(lnRow).Item("xNewDatax") = "1"
            loRow = loData.Select("sMobileNo = " & strParm(p_oDTMstr(0).Item("sMobileNo")))
        Else
            'Locate for the actual index of the found row using DataTable.Rows.IndexOf(DataRow)
            lnRow = loData.Rows.IndexOf(loRow(0))
            loData(lnRow).Item("nPriority") = 0
            loData(lnRow).Item("cRecdStat") = 1
            loData(lnRow).Item("cInvalidx") = "0"
            loData(lnRow).Item("cNewMobil") = "1"

        End If

        If loData.Rows.Count = 1 Then
            loData(lnRow).Item("nEntryNox") = 1
            loData(lnRow).Item("nPriority") = 1
        Else
            'Set the Entry No
            Call SortTable(loData, "nEntryNox")
            For lnRow = 0 To loData.Rows.Count - 1
                loData(lnRow).Item("nEntryNox") = lnRow + 1
            Next

            'Set the Priority No
            Call SortTable(loData, "nPriority")
            For lnRow = 0 To loData.Rows.Count - 1
                loData(lnRow).Item("nPriority") = lnRow + 1
            Next
        End If

        For lnRow = 0 To loData.Rows.Count - 1
            If loData(lnRow).Item("xNewDatax") = "1" Then
                lsSQL = ADO2SQL(loData, lnRow, "Client_Mobile", , , , "xNewDatax")
            Else
                lsSQL = ADO2SQL(loData, lnRow, "Client_Mobile", "sClientID = " & strParm(loData(lnRow).Item("sClientID")) & " AND sMobileNo = " & strParm(loData(lnRow).Item("sMobileNo")), , , "xNewDatax")
            End If

            If lsSQL <> "" Then
                If p_oApp.Execute(lsSQL, "Client_Mobile") = 0 Then
                    Return False
                End If
            End If
        Next

        Return True
    End Function

    Private Function saveTelephone() As Boolean
        Dim lsSQL As String

        'Load all Mobile No of this Employee
        lsSQL = "SELECT sClientID" & _
                     ", nEntryNox" & _
                     ", sPhoneNox" & _
                     ", nPriority" & _
                     ", cInvalidx" & _
                     ", cConfirmd" & _
                     ", cRecdStat" & _
                     ", '0' xNewDatax" & _
               " FROM Client_Telephone" & _
               " WHERE sClientID = " & strParm(p_oDTMstr(0).Item("sClientID"))
        Dim loData = p_oApp.ExecuteQuery(lsSQL)

        'Search for the existence of this Mobile No
        Dim loRow() As DataRow = loData.Select("sPhoneNox = " & strParm(p_oDTMstr(0).Item("sPhoneNox")))

        Dim lnRow As Integer
        If loRow.Count = 0 Then
            'It seems the mobile no is not existing from this client's record
            lnRow = loData.Rows.Count
            loData.Rows.Add(loData.NewRow())
            loData(lnRow).Item("sClientID") = p_oDTMstr(0).Item("sClientID")
            loData(lnRow).Item("nEntryNox") = lnRow + 1
            loData(lnRow).Item("sPhoneNox") = p_oDTMstr(0).Item("sPhoneNox")
            loData(lnRow).Item("nPriority") = 0
            loData(lnRow).Item("cInvalidx") = "0"
            loData(lnRow).Item("cConfirmd") = "0"
            loData(lnRow).Item("cRecdStat") = "1"
            loData(lnRow).Item("xNewDatax") = "1"
            loRow = loData.Select("sPhoneNox = " & strParm(p_oDTMstr(0).Item("sPhoneNox")))
        Else
            'Locate for the actual index of the found row using DataTable.Rows.IndexOf(DataRow)
            lnRow = loData.Rows.IndexOf(loRow(0))
            loData(lnRow).Item("nPriority") = 0
            loData(lnRow).Item("cRecdStat") = 1
            loData(lnRow).Item("cInvalidx") = "0"
        End If

        If loData.Rows.Count = 1 Then
            loData(lnRow).Item("nEntryNox") = 1
            loData(lnRow).Item("nPriority") = 1
        Else
            'Set the Entry No
            Call SortTable(loData, "nEntryNox")
            For lnRow = 0 To loData.Rows.Count - 1
                loData(lnRow).Item("nEntryNox") = lnRow + 1
            Next

            'Set the Priority No
            Call SortTable(loData, "nPriority")
            For lnRow = 0 To loData.Rows.Count - 1
                loData(lnRow).Item("nPriority") = lnRow + 1
            Next
        End If

        For lnRow = 0 To loData.Rows.Count - 1
            If loData(lnRow).Item("xNewDatax") = "1" Then
                lsSQL = ADO2SQL(loData, lnRow, "Client_Telephone", , , , "xNewDatax")
            Else
                lsSQL = ADO2SQL(loData, lnRow, "Client_Telephone", "sClientID = " & strParm(loData(lnRow).Item("sClientID")) & " AND sPhoneNox = " & strParm(loData(lnRow).Item("sPhoneNox")), , , "xNewDatax")
            End If

            If lsSQL <> "" Then
                If p_oApp.Execute(lsSQL, "Client_Telephone") = 0 Then
                    Return False
                End If
            End If
        Next
        Return True
    End Function

    Private Function SaveEmailAdd() As Boolean
        Dim lsSQL As String

        'Load all Mobile No of this Employee
        lsSQL = "SELECT sClientID" & _
                     ", nEntryNox" & _
                     ", sEmailAdd" & _
                     ", nPriority" & _
                     ", '0' xNewDatax" & _
               " FROM Client_eMail_Address" & _
               " WHERE sClientID = " & strParm(p_oDTMstr(0).Item("sClientID"))
        Dim loData = p_oApp.ExecuteQuery(lsSQL)

        'Search for the existence of this Mobile No
        Dim loRow() As DataRow = loData.Select("sEmailAdd = " & strParm(p_oDTMstr(0).Item("sEmailAdd")))

        Dim lnRow As Integer
        If loRow.Count = 0 Then
            'It seems the mobile no is not existing from this client's record
            lnRow = loData.Rows.Count
            loData.Rows.Add(loData.NewRow())
            loData(lnRow).Item("sClientID") = p_oDTMstr(0).Item("sClientID")
            loData(lnRow).Item("nEntryNox") = lnRow + 1
            loData(lnRow).Item("sEmailAdd") = p_oDTMstr(0).Item("sEmailAdd")
            loData(lnRow).Item("nPriority") = 0
            loData(lnRow).Item("xNewDatax") = "1"
            loRow = loData.Select("sEmailAdd = " & strParm(p_oDTMstr(0).Item("sEmailAdd")))
        Else
            'Locate for the actual index of the found row using DataTable.Rows.IndexOf(DataRow)
            lnRow = loData.Rows.IndexOf(loRow(0))
            loData(lnRow).Item("nPriority") = 0
        End If

        If loData.Rows.Count = 1 Then
            loData(lnRow).Item("nEntryNox") = 1
            loData(lnRow).Item("nPriority") = 1
        Else
            'Set the Entry No
            Call SortTable(loData, "nEntryNox")
            For lnRow = 0 To loData.Rows.Count - 1
                loData(lnRow).Item("nEntryNox") = lnRow + 1
            Next

            'Set the Priority No
            Call SortTable(loData, "nPriority")
            For lnRow = 0 To loData.Rows.Count - 1
                loData(lnRow).Item("nPriority") = lnRow + 1
            Next
        End If

        For lnRow = 0 To loData.Rows.Count - 1
            If loData(lnRow).Item("xNewDatax") = "1" Then
                lsSQL = ADO2SQL(loData, lnRow, "Client_eMail_Address", , , , "xNewDatax")
            Else
                lsSQL = ADO2SQL(loData, lnRow, "Client_eMail_Address", "sClientID = " & strParm(loData(lnRow).Item("sClientID")) & " AND sEmailAdd = " & strParm(loData(lnRow).Item("sEmailAdd")), , , "xNewDatax")
            End If

            If lsSQL <> "" Then
                If p_oApp.Execute(lsSQL, "Client_eMail_Address") = 0 Then
                    Return False
                End If
            End If
        Next
        Return True
    End Function

    Public Sub SearchMaster(ByVal fnIndex As Integer, ByVal fsValue As String)
        Select Case fnIndex
            Case 80 ' xBirthPlc
                getBirthPlace(10, 80, fsValue, False, True)
            Case 81 ' sTownName
                getTown(13, 81, fsValue, False, True)
            Case 82 ' sBrgyName
                getBarangay(14, 82, fsValue, False, True)
            Case 83 ' sRelgnNme
                getReligion(19, 83, fsValue, False, True)
            Case 84 ' sOccptnNm
                getOccupation(24, 84, fsValue, False, True)
            Case 85 ' sNational
                getNational(8, 85, fsValue, False, True)
        End Select
    End Sub

    Private Sub initMaster()
        Dim lnCtr As Integer
        For lnCtr = 0 To p_oDTMstr.Columns.Count - 1
            Select Case LCase(p_oDTMstr.Columns(lnCtr).ColumnName)
                Case "sclientid"
                    'kalyptus - 2024.09.17 11:42am
                    'Include terminal number in sClientID
                    p_oDTMstr(0).Item(lnCtr) = GetNextCode(p_sMasTable, "sClientID", True, p_oApp.Connection, True, p_oApp.BranchCode + p_oApp.POSTerminal)
                Case "dbirthdte"
                    p_oDTMstr(0).Item(lnCtr) = xsNULL_DATE
                Case "ngrssincm"
                    p_oDTMstr(0).Item(lnCtr) = 0
                Case "dmodified", "smodified"
                Case "crecdstat"
                    p_oDTMstr(0).Item(lnCtr) = "1"
                Case "cclienttp", "clrclient", "cmcclient", "cscclient", "cspclient", "ccpslient"
                    p_oDTMstr(0).Item(lnCtr) = "0"
                Case Else
                    p_oDTMstr(0).Item(lnCtr) = ""
            End Select
        Next
    End Sub

    Private Sub InitOthers()
        p_oOthersx.xBirthPlc = ""
        p_oOthersx.sTownName = ""
        p_oOthersx.sBrgyName = ""
        p_oOthersx.sRelgnNme = ""
        p_oOthersx.sOccptnNm = ""
        p_oOthersx.sNational = ""
    End Sub

    Private Function isEntryOk() As Boolean
        If p_oDTMstr(0).Item("sLastName") = "" Or _
           p_oDTMstr(0).Item("sFrstName") = "" Or _
           p_oDTMstr(0).Item("sMiddName") = "" Then
            MsgBox("Invalid Name Info Detected...", vbOKOnly, p_sMsgHeadr)
            Return False
        ElseIf p_oDTMstr(0).Item("sTownIDxx") = "" Then
            MsgBox("Invalid Town Info Detected...", vbOKOnly, p_sMsgHeadr)
            Return False
        ElseIf p_oDTMstr(0).Item("sMobileNo") = "" Then
            If MsgBox("Invalid Mobile No Info Detected..." & vbCrLf & _
                   "Do you want to continue?", vbYesNo, p_sMsgHeadr) <> vbNo Then
                Return False
            End If
        End If

        If p_oDTMstr(0).Item("cLRClient") = "1" Or p_oDTMstr(0).Item("cMCClient") = "1" Then
            If p_oDTMstr(0).Item("dBirthDte") = "1900-01-01" Then
                MsgBox("Invalid Birth Date Detected...", vbOKOnly, p_sMsgHeadr)
                Return False
            ElseIf p_oDTMstr(0).Item("sBirthPlc") = "" Then
                MsgBox("Invalid Birth Place Detected...", vbOKOnly, p_sMsgHeadr)
                Return False
            ElseIf p_oDTMstr(0).Item("sAddressx") = "" Then
                MsgBox("Invalid Address Detected...", vbOKOnly, p_sMsgHeadr)
                Return False
            End If
        End If

        Return True
    End Function

    'This method implements a search master where id and desc are not joined.
    Private Sub getBirthPlace( _
                      ByVal fnColIdx As Integer _
                    , ByVal fnColDsc As Integer _
                    , ByVal fsValue As String _
                    , ByVal fbIsCode As Boolean _
                    , ByVal fbIsSrch As Boolean)

        'Compare the value to be search against the value in our column
        If fbIsCode Then
            If fsValue = p_oDTMstr(0).Item(fnColIdx) And fsValue <> "" And p_oOthersx.xBirthPlc <> "" Then Exit Sub
        Else
            If fsValue = p_oOthersx.xBirthPlc And fsValue <> "" Then Exit Sub
        End If

        Dim lsSQL As String
        lsSQL = "SELECT" & _
                       "  a.sTownIDxx" & _
                       ", CONCAT(a.sTownName, ', ', b.sProvName, ' ', a.sZippCode) sTownName" & _
               " FROM TownCity a" & _
                  " LEFT JOIN Province b" & _
                     " ON a.sProvIDxx = b.sProvIDxx" & _
               " WHERE a.cRecdStat = '1'" & _
               " ORDER BY a.sTownName, b.sProvName"

        'Are we using like comparison or equality comparison
        If fbIsSrch Then
            Dim loRow As DataRow = KwikSearch(p_oApp _
                                             , lsSQL _
                                             , True _
                                             , fsValue _
                                             , "sTownIDxx»sTownName" _
                                             , "Code»Town/City", _
                                             , "a.sTownIDxx»CONCAT(a.sTownName, ', ', b.sProvName, ' ', a.sZippCode)" _
                                             , IIf(fbIsCode, 0, 1))
            If IsNothing(loRow) Then
                p_oDTMstr(0).Item(fnColIdx) = ""
                p_oOthersx.xBirthPlc = ""
            Else
                p_oDTMstr(0).Item(fnColIdx) = loRow.Item("sTownIDxx")
                p_oOthersx.xBirthPlc = loRow.Item("sTownName")
            End If

            RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.xBirthPlc)
            Exit Sub

        End If

        If fsValue = "" Then
            lsSQL = AddCondition(lsSQL, "0=1")
        Else
            If fbIsCode Then
                lsSQL = AddCondition(lsSQL, "a.sTownIDxx = " & strParm(fsValue))
            Else
                lsSQL = AddCondition(lsSQL, "CONCAT(a.sTownName, ', ', b.sProvName, ' ', a.sZippCode) = " & strParm(fsValue))
            End If
        End If

        Dim loDta As DataTable
        loDta = p_oApp.ExecuteQuery(lsSQL)

        If loDta.Rows.Count = 0 Then
            p_oDTMstr(0).Item(fnColIdx) = ""
            p_oOthersx.xBirthPlc = ""
        ElseIf loDta.Rows.Count = 1 Then
            p_oDTMstr(0).Item(fnColIdx) = loDta(0).Item("sTownIDxx")
            p_oOthersx.xBirthPlc = loDta(0).Item("sTownName")
        End If

        RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.xBirthPlc)

    End Sub

    'This method implements a search master where id and desc are not joined.
    Private Sub getTown( _
                      ByVal fnColIdx As Integer _
                    , ByVal fnColDsc As Integer _
                    , ByVal fsValue As String _
                    , ByVal fbIsCode As Boolean _
                    , ByVal fbIsSrch As Boolean)

        'Compare the value to be search against the value in our column
        If fbIsCode Then
            If fsValue = p_oDTMstr(0).Item(fnColIdx) And fsValue <> "" And p_oOthersx.sTownName <> "" Then Exit Sub
        Else
            If fsValue = p_oOthersx.sTownName And fsValue <> "" Then Exit Sub
        End If

        Dim lsSQL As String
        lsSQL = "SELECT" & _
                       "  a.sTownIDxx" & _
                       ", CONCAT(a.sTownName, ', ', b.sProvName, ' ', a.sZippCode) sTownName" & _
               " FROM TownCity a" & _
                  " LEFT JOIN Province b" & _
                     " ON a.sProvIDxx = b.sProvIDxx" & _
               " WHERE a.cRecdStat = '1'" & _
               " ORDER BY a.sTownName, b.sProvName"

        'Are we using like comparison or equality comparison
        If fbIsSrch Then
            Dim loRow As DataRow = KwikSearch(p_oApp _
                                             , lsSQL _
                                             , True _
                                             , fsValue _
                                             , "sTownIDxx»sTownName" _
                                             , "Code»Town/City", _
                                             , "a.sTownIDxx»CONCAT(a.sTownName, ', ', b.sProvName, ' ', a.sZippCode)" _
                                             , IIf(fbIsCode, 0, 1))
            If IsNothing(loRow) Then
                p_oDTMstr(0).Item(fnColIdx) = ""
                p_oOthersx.sTownName = ""
            Else
                p_oDTMstr(0).Item(fnColIdx) = loRow.Item("sTownIDxx")
                p_oOthersx.sTownName = loRow.Item("sTownName")
            End If

            RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sTownName)
            Exit Sub

        End If

        If fsValue = "" Then
            lsSQL = AddCondition(lsSQL, "0=1")
        Else
            If fbIsCode Then
                lsSQL = AddCondition(lsSQL, "a.sTownIDxx = " & strParm(fsValue))
            Else
                lsSQL = AddCondition(lsSQL, "CONCAT(a.sTownName, ', ', b.sProvName, ' ', a.sZippCode) = " & strParm(fsValue))
            End If
        End If

        Dim loDta As DataTable
        loDta = p_oApp.ExecuteQuery(lsSQL)

        If loDta.Rows.Count = 0 Then
            p_oDTMstr(0).Item(fnColIdx) = ""
            p_oOthersx.sTownName = ""
        ElseIf loDta.Rows.Count = 1 Then
            p_oDTMstr(0).Item(fnColIdx) = loDta(0).Item("sTownIDxx")
            p_oOthersx.sTownName = loDta(0).Item("sTownName")
        End If

        RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sTownName)

    End Sub

    'This method implements a search master where id and desc are not joined.
    Private Sub getBarangay( _
                      ByVal fnColIdx As Integer _
                    , ByVal fnColDsc As Integer _
                    , ByVal fsValue As String _
                    , ByVal fbIsCode As Boolean _
                    , ByVal fbIsSrch As Boolean)

        'Reset the value of barangay if Town is empty
        If p_oDTMstr(0).Item("sTownIDxx") = "" Then
            p_oOthersx.sBrgyName = ""
            p_oDTMstr(0).Item(fnColIdx) = ""

            RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sBrgyName)
            Exit Sub
        End If

        'Compare the value to be search against the value in our column
        If fbIsCode Then
            If fsValue = p_oDTMstr(0).Item(fnColIdx) And fsValue <> "" And p_oOthersx.sBrgyName <> "" Then Exit Sub
        Else
            If fsValue = p_oOthersx.sBrgyName And fsValue <> "" Then Exit Sub
        End If

        Dim lsSQL As String
        lsSQL = "SELECT" & _
                       "  a.sBrgyIDxx" & _
                       ", a.sBrgyName" & _
               " FROM Barangay a" & _
               " WHERE a.sTownIDxx = " & strParm(p_oDTMstr(0).Item("sTownIDxx")) & _
                 " AND a.cRecdStat = '1'"

        'Are we using like comparison or equality comparison
        If fbIsSrch Then
            Dim loRow As DataRow = KwikSearch(p_oApp _
                                             , lsSQL _
                                             , True _
                                             , fsValue _
                                             , "sBrgyIDxx»sBrgyName" _
                                             , "Code»Barangay", _
                                             , "a.sBrgyIDxx»a.sBrgyName" _
                                             , IIf(fbIsCode, 0, 1))
            If IsNothing(loRow) Then
                p_oDTMstr(0).Item(fnColIdx) = ""
                p_oOthersx.sBrgyName = ""
            Else
                p_oDTMstr(0).Item(fnColIdx) = loRow.Item("sBrgyIDxx")
                p_oOthersx.sBrgyName = loRow.Item("sBrgyName")
            End If

            RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sBrgyName)
            Exit Sub

        End If

        If fsValue = "" Then
            lsSQL = AddCondition(lsSQL, "0=1")
        Else
            If fbIsCode Then
                lsSQL = AddCondition(lsSQL, "a.sBrgyIDxx = " & strParm(fsValue))
            Else
                lsSQL = AddCondition(lsSQL, "a.sBrgyName = " & strParm(fsValue))
            End If
        End If

        Dim loDta As DataTable
        loDta = p_oApp.ExecuteQuery(lsSQL)

        If loDta.Rows.Count = 0 Then
            p_oDTMstr(0).Item(fnColIdx) = ""
            p_oOthersx.sBrgyName = ""
        ElseIf loDta.Rows.Count = 1 Then
            p_oDTMstr(0).Item(fnColIdx) = loDta(0).Item("sBrgyIDxx")
            p_oOthersx.sBrgyName = loDta(0).Item("sBrgyName")
        End If

        RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sBrgyName)

    End Sub

    'This method implements a search master where id and desc are not joined.
    Private Sub getReligion( _
                      ByVal fnColIdx As Integer _
                    , ByVal fnColDsc As Integer _
                    , ByVal fsValue As String _
                    , ByVal fbIsCode As Boolean _
                    , ByVal fbIsSrch As Boolean)

        'Compare the value to be search against the value in our column
        If fbIsCode Then
            If fsValue = p_oDTMstr(0).Item(fnColIdx) And fsValue <> "" And p_oOthersx.sRelgnNme <> "" Then Exit Sub
        Else
            If fsValue = p_oOthersx.sRelgnNme And fsValue <> "" Then Exit Sub
        End If

        Dim lsSQL As String
        lsSQL = "SELECT" & _
                       "  a.sRelgnIDx" & _
                       ", a.sRelgnNme" & _
               " FROM Religion a" & _
               " WHERE a.cRecdStat = '1'"

        'Are we using like comparison or equality comparison
        If fbIsSrch Then
            Dim loRow As DataRow = KwikSearch(p_oApp _
                                             , lsSQL _
                                             , True _
                                             , fsValue _
                                             , "sRelgnIDx»sRelgnNme" _
                                             , "Code»Religion", _
                                             , "a.sRelgnIDx»a.sRelgnNme" _
                                             , IIf(fbIsCode, 0, 1))
            If IsNothing(loRow) Then
                p_oDTMstr(0).Item(fnColIdx) = ""
                p_oOthersx.sRelgnNme = ""
            Else
                p_oDTMstr(0).Item(fnColIdx) = loRow.Item("sRelgnIDx")
                p_oOthersx.sRelgnNme = loRow.Item("sRelgnNme")
            End If

            RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sRelgnNme)
            Exit Sub

        End If

        If fsValue = "" Then
            lsSQL = AddCondition(lsSQL, "0=1")
        Else
            If fbIsCode Then
                lsSQL = AddCondition(lsSQL, "a.sRelgnIDx = " & strParm(fsValue))
            Else
                lsSQL = AddCondition(lsSQL, "a.sRelgnNme = " & strParm(fsValue))
            End If
        End If

        Dim loDta As DataTable
        loDta = p_oApp.ExecuteQuery(lsSQL)

        If loDta.Rows.Count = 0 Then
            p_oDTMstr(0).Item(fnColIdx) = ""
            p_oOthersx.sRelgnNme = ""
        ElseIf loDta.Rows.Count = 1 Then
            p_oDTMstr(0).Item(fnColIdx) = loDta(0).Item("sRelgnIDx")
            p_oOthersx.sRelgnNme = loDta(0).Item("sRelgnNme")
        End If

        RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sRelgnNme)

    End Sub


    'This method implements a search master where id and desc are not joined.
    Private Sub getOccupation( _
                      ByVal fnColIdx As Integer _
                    , ByVal fnColDsc As Integer _
                    , ByVal fsValue As String _
                    , ByVal fbIsCode As Boolean _
                    , ByVal fbIsSrch As Boolean)

        'Compare the value to be search against the value in our column
        If fbIsCode Then
            If fsValue = p_oDTMstr(0).Item(fnColIdx) And fsValue <> "" And p_oOthersx.sOccptnNm <> "" Then Exit Sub
        Else
            If fsValue = p_oOthersx.sOccptnNm And fsValue <> "" Then Exit Sub
        End If

        Dim lsSQL As String
        lsSQL = "SELECT" & _
                       "  a.sOccptnID" & _
                       ", a.sOccptnNm" & _
               " FROM Occupation a" & _
               " WHERE a.cRecdStat = '1'"

        'Are we using like comparison or equality comparison
        If fbIsSrch Then
            Dim loRow As DataRow = KwikSearch(p_oApp _
                                             , lsSQL _
                                             , True _
                                             , fsValue _
                                             , "sOccptnID»sOccptnNm" _
                                             , "Code»Occupation", _
                                             , "a.sOccptnID»a.sOccptnNm" _
                                             , IIf(fbIsCode, 0, 1))
            If IsNothing(loRow) Then
                p_oDTMstr(0).Item(fnColIdx) = ""
                p_oOthersx.sOccptnNm = ""
            Else
                p_oDTMstr(0).Item(fnColIdx) = loRow.Item("sOccptnID")
                p_oOthersx.sOccptnNm = loRow.Item("sOccptnNm")
            End If

            RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sOccptnNm)
            Exit Sub

        End If

        If fsValue = "" Then
            lsSQL = AddCondition(lsSQL, "0=1")
        Else
            If fbIsCode Then
                lsSQL = AddCondition(lsSQL, "a.sOccptnID = " & strParm(fsValue))
            Else
                lsSQL = AddCondition(lsSQL, "a.sOccptnNm = " & strParm(fsValue))
            End If
        End If

        Dim loDta As DataTable
        loDta = p_oApp.ExecuteQuery(lsSQL)

        If loDta.Rows.Count = 0 Then
            p_oDTMstr(0).Item(fnColIdx) = ""
            p_oOthersx.sOccptnNm = ""
        ElseIf loDta.Rows.Count = 1 Then
            p_oDTMstr(0).Item(fnColIdx) = loDta(0).Item("sOccptnID")
            p_oOthersx.sOccptnNm = loDta(0).Item("sOccptnNm")
        End If

        RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sOccptnNm)

    End Sub

    'This method implements a search master where id and desc are not joined.
    Private Sub getNational( _
                      ByVal fnColIdx As Integer _
                    , ByVal fnColDsc As Integer _
                    , ByVal fsValue As String _
                    , ByVal fbIsCode As Boolean _
                    , ByVal fbIsSrch As Boolean)

        'Compare the value to be search against the value in our column
        If fbIsCode Then
            If fsValue = p_oDTMstr(0).Item(fnColIdx) And fsValue <> "" And p_oOthersx.sNational <> "" Then Exit Sub
        Else
            If fsValue = p_oOthersx.sNational And fsValue <> "" Then Exit Sub
        End If

        Dim lsSQL As String
        lsSQL = "SELECT" & _
                       "  a.sCntryCde" & _
                       ", a.sNational" & _
               " FROM Country a" & _
               " WHERE a.cRecdStat = '1'"

        'Are we using like comparison or equality comparison
        If fbIsSrch Then
            Dim loRow As DataRow = KwikSearch(p_oApp _
                                             , lsSQL _
                                             , True _
                                             , fsValue _
                                             , "sCntryCde»sNational" _
                                             , "Code»Nationality", _
                                             , "a.sCntryCde»a.sNational" _
                                             , IIf(fbIsCode, 0, 1))
            If IsNothing(loRow) Then
                p_oDTMstr(0).Item(fnColIdx) = ""
                p_oOthersx.sNational = ""
            Else
                p_oDTMstr(0).Item(fnColIdx) = loRow.Item("sCntryCde")
                p_oOthersx.sNational = loRow.Item("sNational")
            End If

            RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sNational)
            Exit Sub

        End If

        If fsValue = "" Then
            lsSQL = AddCondition(lsSQL, "0=1")
        Else
            If fbIsCode Then
                lsSQL = AddCondition(lsSQL, "a.sCntryCde = " & strParm(fsValue))
            Else
                lsSQL = AddCondition(lsSQL, "a.sNational = " & strParm(fsValue))
            End If
        End If

        Dim loDta As DataTable
        loDta = p_oApp.ExecuteQuery(lsSQL)

        If loDta.Rows.Count = 0 Then
            p_oDTMstr(0).Item(fnColIdx) = ""
            p_oOthersx.sNational = ""
        ElseIf loDta.Rows.Count = 1 Then
            p_oDTMstr(0).Item(fnColIdx) = loDta(0).Item("sCntryCde")
            p_oOthersx.sNational = loDta(0).Item("sNational")
        End If

        RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sNational)

    End Sub

    'This method implements a search master where id and desc are not joined.
    Private Sub getSpouse()

        Dim lsSQL As String
        lsSQL = "SELECT DISTINCT" & _
                       " sCompnyNm sClientNm" & _
               " FROM Client_Master a" & _
               " WHERE a.sSpouseID = " & strParm(p_oDTMstr(0).Item("sSpouseID")) & _
                 " AND a.sClientID <> " & strParm(p_oDTMstr(0).Item("sClientID"))

        Dim loDta As DataTable
        loDta = p_oApp.ExecuteQuery(lsSQL)

        If loDta.Rows.Count = 0 Then
            p_oOthersx.sSpouseNm = ""
        Else
            p_oOthersx.sSpouseNm = loDta(0).Item("sClientNm")
        End If
    End Sub

    Private Function getSQ_Master() As String
        Return "SELECT a.sClientID" & _
                    ", a.sLastName" & _
                    ", a.sFrstName" & _
                    ", a.sMiddName" & _
                    ", a.sMaidenNm" & _
                    ", a.sSuffixNm" & _
                    ", a.cGenderCd" & _
                    ", a.cCvilStat" & _
                    ", a.sCitizenx" & _
                    ", a.dBirthDte" & _
                    ", a.sBirthPlc" & _
                    ", a.sHouseNox" & _
                    ", a.sAddressx" & _
                    ", a.sTownIDxx" & _
                    ", a.sBrgyIDxx" & _
                    ", a.sPhoneNox" & _
                    ", a.sMobileNo" & _
                    ", a.sEmailAdd" & _
                    ", a.cEducLevl" & _
                    ", a.sRelgnIDx" & _
                    ", a.sTaxIDNox" & _
                    ", a.sSSSNoxxx" & _
                    ", a.sAddlInfo" & _
                    ", a.sCompnyNm" & _
                    ", a.sOccptnID" & _
                    ", a.sOccptnOT" & _
                    ", a.nGrssIncm" & _
                    ", a.sClientNo" & _
                    ", a.sSpouseID" & _
                    ", a.sFatherID" & _
                    ", a.sMotherID" & _
                    ", a.sSiblngID" & _
                    ", a.cClientTp" & _
                    ", a.cLRClient" & _
                    ", a.cMCClient" & _
                    ", a.cSCClient" & _
                    ", a.cSPClient" & _
                    ", a.cCPClient" & _
                    ", a.cRecdStat" & _
                    ", a.sModified" & _
                    ", a.dModified" & _
                " FROM " & p_sMasTable & " a"
    End Function

    Private Function getSQ_Browse() As String
        Return "SELECT a.sClientID" & _
                       ", a.sCompnyNm sClientNm" & _
                       ", CONCAT(IF(IFNull(a.sHouseNox, '') = '', '', CONCAT(a.sHouseNox, ' ')), a.sAddressx, ', ', b.sTownName, ', ', c.sProvName, ' ', b.sZippCode) xAddressx" & _
                       ", a.dBirthDte" & _
              " FROM " & p_sMasTable & " a" & _
                    " LEFT JOIN TownCity b ON a.sTownIDxx = b.sTownIDxx" & _
                    " LEFT JOIN Province c ON b.sProvIDxx = c.sProvIDxx" & _
              " WHERE a.cRecdStat = '1'"
    End Function

    Public Function getValidMobile(ByVal fsMobileNo As String) As String

        fsMobileNo = getValidPhone(fsMobileNo)

        If fsMobileNo <> "" Then
            If IsNumeric(fsMobileNo) Then
                'Convert country code to 0
                If Left(fsMobileNo, 3) = "+63" Then
                    fsMobileNo = "0" & Mid(fsMobileNo, 4)
                ElseIf Left(fsMobileNo, 2) = "63" Then
                    fsMobileNo = "0" & Mid(fsMobileNo, 3)
                End If

                If Len(fsMobileNo) <> xeMOBILE_LEN Then fsMobileNo = ""
            End If
        Else
            fsMobileNo = ""
        End If

        Return fsMobileNo
    End Function

    Public Function getValidPhone(ByVal fsPhone As String) As String
        If fsPhone <> "" Then
            fsPhone = Replace(fsPhone, " ", "")
            fsPhone = Replace(fsPhone, "-", "")
            fsPhone = Replace(fsPhone, "(", "")
            fsPhone = Replace(fsPhone, ")", "")
        Else
            fsPhone = ""
        End If

        Return fsPhone
    End Function

    Private Class Others
        Public xBirthPlc As String  '10
        Public sTownName As String  '13
        Public sBrgyName As String  '14
        Public sRelgnNme As String  '19
        Public sOccptnNm As String  '24
        Public sNational As String  '08
        Public sSpouseNm As String  '28
    End Class

    Public Function Clone() As Object Implements System.ICloneable.Clone
        Return Me.MemberwiseClone
    End Function

    Public Sub New(ByVal foRider As GRider)
        p_oApp = foRider
        p_sParent = ""
        p_nEditMode = xeEditMode.MODE_UNKNOWN
    End Sub
End Class
