'Note:
'    Example and Usage of Overriding ProcessCmdKey
'       See the function [Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean]
Imports ggcAppDriver
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Diagnostics

Public Class frmClient
    Private WithEvents p_oClient As Client
    Private pnLoadx As Integer
    Private poControl As Control
    Private p_nButton As Integer
    Private p_bOnSeek As Boolean

    'Property ShowMessage()
    Public WriteOnly Property iClient() As Client
        Set(ByVal value As Client)
            p_oClient = value
        End Set
    End Property

    Public ReadOnly Property Cancelled() As Boolean
        Get
            Return p_nButton = 3
        End Get
    End Property

    Private Sub frmClient_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Debug.Print("frmClient_Activated")
        If pnLoadx = 1 Then

            Call loadMaster(Me)

            txtField01.Focus()
            pnLoadx = 2
        End If
    End Sub

    Private Sub frmClient_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.F3 Or e.KeyCode = Keys.Enter Then
            Dim loTxt As Control

            loTxt = Nothing
            If TypeOf poControl Is TextBox Then
                loTxt = CType(poControl, System.Windows.Forms.TextBox)
            ElseIf TypeOf poControl Is CheckBox Then
                loTxt = CType(poControl, System.Windows.Forms.CheckBox)
            ElseIf TypeOf poControl Is ComboBox Then
                loTxt = CType(poControl, System.Windows.Forms.ComboBox)
            End If

            '******************
            Dim loIndex As Integer
            loIndex = Val(Mid(loTxt.Name, 9))

            If Mid(loTxt.Name, 1, 8) = "txtField" Then
                p_bOnSeek = True
                Call p_oClient.SearchMaster(loIndex, loTxt.Text)
                p_bOnSeek = False
            End If

            If TypeOf poControl Is TextBox Or
               TypeOf poControl Is CheckBox Or
               TypeOf poControl Is ComboBox Then
                SelectNextControl(loTxt, True, True, True, True)
            End If
        End If
    End Sub

    Private Sub frmClient_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Debug.Print("frmClient_Load")
        If pnLoadx = 0 Then

            'Set event Handler for txtField
            Call grpEventHandler(Me, GetType(TextBox), "txtField", "GotFocus", AddressOf txtField_GotFocus)
            Call grpEventHandler(Me, GetType(TextBox), "txtField", "LostFocus", AddressOf txtField_LostFocus)
            Call grpCancelHandler(Me, GetType(TextBox), "txtField", "Validating", AddressOf txtField_Validating)

            Call grpEventHandler(Me, GetType(Button), "cmdButtn", "Click", AddressOf cmdButton_Click)

            pnLoadx = 1
        End If
    End Sub

    'Handles GotFocus Events for txtField & txtItems
    Private Sub txtField_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim loIndex As Integer
        loIndex = Val(Mid(sender.Name, 9))

        Console.WriteLine("»Got Focus: " & sender.Name)

        If Mid(sender.Name, 1, 8) = "txtField" Then
            Dim loTxt As TextBox
            loTxt = CType(sender, System.Windows.Forms.TextBox)

            If Not loTxt.ReadOnly Then
                Select Case loIndex
                    Case 9
                        loTxt.Text = Format(p_oClient.Master(loIndex), "yyyy/MM/dd")
                End Select

                loTxt.BackColor = Color.Azure
                loTxt.SelectAll()
            End If

            poControl = loTxt
        End If
    End Sub

    'Handles LostFocus Events for txtField & txtItems
    Private Sub txtField_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)

        Console.WriteLine("Lost Focus: " & sender.Name)

        Dim loIndex As Integer
        loIndex = Val(Mid(sender.Name, 9))

        If Mid(sender.Name, 1, 8) = "txtField" Then
            Dim loTxt As TextBox
            loTxt = CType(sender, System.Windows.Forms.TextBox)

            If Not loTxt.ReadOnly Then
                If Not p_bOnSeek Then p_oClient.Master(loIndex) = loTxt.Text
                Select Case loIndex
                    Case 9
                        loTxt.Text = Format(p_oClient.Master(loIndex), "MMMM dd, yyyy")
                    Case 11
                        TabCntrl01.SelectedTab = TabPages02
                    Case 26
                        If IsDBNull(p_oClient.Master(loIndex)) Then
                            loTxt.Text = ""
                        Else
                            loTxt.Text = Format(p_oClient.Master(loIndex), xsDECIMAL)
                        End If
                End Select

                loTxt.BackColor = SystemColors.Window
                'poControl = Nothing
            End If
        End If
    End Sub

    'Handles Validating Events for txtField & txtItems
    Private Sub txtField_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
        'Dim loIndex As Integer
        'loIndex = Val(Mid(sender.Name, 9))
        'If Mid(sender.Name, 1, 8) = "txtField" Then
        'Dim loTxt As TextBox
        'loTxt = CType(sender, System.Windows.Forms.TextBox)
        'p_oClient.Master(loIndex) = loTxt.Text
        'ElseIf Mid(sender.Name, 1, 8) = "cmbField" Then
        'Dim loCmb As ComboBox
        'loCmb = DirectCast(sender, ComboBox)
        'p_oClient.Master(loIndex) = loCmb.SelectedIndex
        'End If
    End Sub

    Private Sub cmdButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim loChk As Button
        loChk = CType(sender, System.Windows.Forms.Button)

        Dim lnIndex As Integer
        lnIndex = Val(Mid(loChk.Name, 10))

        Select Case lnIndex
            Case 1 ' Ok
                If isEntryOk Then
                    p_nButton = 1
                    Me.Hide()
                End If
            Case 2 ' Search
                'Test if last control is a textbox
                If (TypeOf poControl Is TextBox) Then
                    Dim loTxt As TextBox
                    Dim lnIndex1 As Integer
                    loTxt = CType(poControl, TextBox)
                    lnIndex1 = Val(Mid(loTxt.Name, 9))
                    If LCase(Mid(loTxt.Name, 1, 8)) = "txtfield" Then
                        p_oClient.SearchMaster(lnIndex1, loTxt.Text)
                    End If
                End If
            Case 3 ' Cancel Update
                p_nButton = 3
                Me.Hide()
        End Select
    End Sub

    Private Sub p_oClient_MasterRetrieved(ByVal Index As Integer, ByVal Value As Object) Handles p_oClient.MasterRetrieved
        Dim loTxt As TextBox

        'Find TextBox with specified name
        loTxt = CType(FindTextBox(Me, "txtField" & Format(Index, "00")), TextBox)

        Select Case Index
            Case 9
                loTxt.Text = Format(Value, "MMMM dd, yyyy")
            Case 26
                If IsDBNull(Value) Then
                    loTxt.Text = ""
                Else
                    loTxt.Text = Format(Value, xsDECIMAL)
                End If
            Case Else
                loTxt.Text = Value
        End Select
    End Sub

    Private Sub loadMaster(ByVal loControl As Control)
        Dim loTxt As Control

        For Each loTxt In loControl.Controls
            If loTxt.HasChildren Then
                Call loadMaster(loTxt)
            Else
                If (TypeOf loTxt Is TextBox) Then
                    Dim loIndex As Integer
                    loIndex = Val(Mid(loTxt.Name, 9))
                    Dim loBox As TextBox
                    loBox = CType(loTxt, TextBox)
                    If LCase(Mid(loBox.Name, 1, 8)) = "txtfield" Then
                        If p_oClient.EditMode <> xeEditMode.MODE_UNKNOWN Then
                            Select Case loIndex
                                'sLastName, FrstName, sMiddName, dBirthDte, xBirthPlc
                                Case 1, 2, 3, 9, 80

                                    If loIndex = 9 Then
                                        If p_oClient.EditMode = xeEditMode.MODE_READY Then
                                            If IFNull(p_oClient.OriginalMaster(loIndex), xsNULL_DATE) = xsNULL_DATE Then
                                                p_oClient.Master(loIndex) = xsNULL_DATE
                                                loBox.ReadOnly = False
                                            Else
                                                loBox.ReadOnly = True
                                            End If
                                        End If

                                        loBox.Text = Format(p_oClient.Master(loIndex), "MMMM dd, yyyy")
                                    Else
                                        If p_oClient.EditMode = xeEditMode.MODE_READY Then
                                            If loIndex = 80 Then
                                                If p_oClient.OriginalMaster(10) = "" Then
                                                    loBox.ReadOnly = False
                                                Else
                                                    loBox.ReadOnly = True
                                                End If
                                            Else
                                                If p_oClient.OriginalMaster(loIndex) = "" Then
                                                    loBox.ReadOnly = False
                                                Else
                                                    loBox.ReadOnly = True
                                                End If
                                            End If
                                        End If
                                        loBox.Text = p_oClient.Master(loIndex)
                                    End If
                                Case Else
                                    If IsDBNull(p_oClient.Master(loIndex)) Then
                                        loBox.Text = ""
                                    Else
                                        loBox.Text = p_oClient.Master(loIndex)
                                    End If
                            End Select
                        Else
                            loBox.Text = ""
                        End If
                    ElseIf LCase(Mid(loBox.Name, 1, 8)) = "txtother" Then
                        If p_oClient.EditMode <> xeEditMode.MODE_UNKNOWN Then
                            Select Case loIndex
                                Case 1  'xAddressx
                                    loBox.Text = IIf(p_oClient.Master("sHouseNox") = "", "", p_oClient.Master("sHouseNox") & " ") & _
                                                 p_oClient.Master("sAddressx") & ", " & _
                                                 p_oClient.Master("sTownName")
                            End Select
                        Else
                            loBox.Text = ""
                        End If
                    End If 'LCase(Mid(loTxt.Name, 1, 8)) = "txtfield"
                ElseIf LCase(Mid(loTxt.Name, 1, 8)) = "cmbfield" Then
                    Dim loIndex As Integer
                    loIndex = Val(Mid(loTxt.Name, 9))
                    Dim loCmb As ComboBox
                    loCmb = DirectCast(loTxt, ComboBox)
                    If p_oClient.EditMode <> xeEditMode.MODE_UNKNOWN Then
                        loCmb.SelectedIndex = IIf(IsNumeric(p_oClient.Master(loIndex)), p_oClient.Master(loIndex), -1)
                    Else
                        loCmb.SelectedIndex = -1
                    End If
                End If '(TypeOf loTxt Is TextBox)
            End If 'If loTxt.HasChildren
        Next 'loTxt In loControl.Controls
    End Sub

    Private Function isEntryOk() As Boolean
        If p_oClient.Master("sLastName") = "" Or _
           p_oClient.Master("sFrstName") = "" Or _
           p_oClient.Master("sMiddName") = "" Then
            MsgBox("Invalid Name Info Detected...")
            Return False
        ElseIf p_oClient.Master("sTownIDxx") = "" Then
            MsgBox("Invalid Town Info Detected...")
            Return False
        ElseIf p_oClient.Master("sMobileNo") = "" Then
            MsgBox("Invalid Mobile No Info Detected...")
            Return False
        End If

        If p_oClient.Master("cLRClient") = "1" Or p_oClient.Master("cMCClient") = "1" Then
            If p_oClient.Master("dBirthDte") = xsNULL_DATE Then
                MsgBox("Invalid Birth Date Detected...")
                Return False
            ElseIf p_oClient.Master("sBirthPlc") = "" Then
                MsgBox("Invalid Birth Place Detected...")
                Return False
            ElseIf p_oClient.Master("sAddressx") = "" Then
                MsgBox("Invalid Address Detected...")
                Return False
            End If
        End If

        Return True
    End Function

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabCntrl01.SelectedIndexChanged
        If TabCntrl01.SelectedTab.Name = "TabPage1" Then
            txtField01.Focus()
        ElseIf TabCntrl01.SelectedTab.Name = "TabPage2" Then
            txtField11.Focus()
        ElseIf TabCntrl01.SelectedTab.Name = "TabPage3" Then
            txtField84.Focus()
        End If
    End Sub

    Private Sub cmbField_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbField06.GotFocus, cmbField07.GotFocus, cmbField18.GotFocus
        Console.WriteLine("Got Focus: " & sender.Name)

        Dim loCmb As ComboBox
        loCmb = DirectCast(sender, ComboBox)

        loCmb.BackColor = Color.Azure
        loCmb.SelectAll()

        poControl = loCmb
    End Sub

    Private Sub cmbField_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbField06.LostFocus, cmbField07.LostFocus, cmbField18.LostFocus
        'Console.WriteLine("Lost Focus: " & sender.Name)

        Dim loIndex As Integer
        loIndex = Val(Mid(sender.Name, 9))

        Dim loCmb As ComboBox
        loCmb = DirectCast(sender, ComboBox)

        p_oClient.Master(loIndex) = IIf(loCmb.SelectedIndex = -1, "", loCmb.SelectedIndex)

        loCmb.BackColor = SystemColors.Window
        'poControl = Nothing
    End Sub

    'Returning true means original function of the key is override...
    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, ByVal keyData As Keys) As Boolean
        'Console.WriteLine("Keypressed: " & keyData)

        If keyData = 65545 Then     'Handle the events for (shift + tab)
            If txtField01.Focused Then
                TabCntrl01.SelectedTab = TabPages03
                Return True
            ElseIf txtField11.Focused Then
                TabCntrl01.SelectedTab = TabPages01
                Return True
            ElseIf txtField84.Focused Then
                TabCntrl01.SelectedTab = TabPages02
                Return True
            End If
        ElseIf keyData = Keys.Tab Then 'Handle the events for (tab)
            If txtField04.Focused Or txtField86.Focused Then
                TabCntrl01.SelectedTab = TabPages02
                Return True
            ElseIf txtField17.Focused Then
                TabCntrl01.SelectedTab = TabPages03
                Return True
            ElseIf cmdButtn03.Focused Then
                TabCntrl01.SelectedTab = TabPages01
                Return True
            End If
        End If

        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function
End Class