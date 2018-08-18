#################################### cmdclose
Private Sub cmd_close_Click()
  Me.Undo
DoCmd.Close acForm, Me.Name
End Sub

################################### cmdsave
Private Sub cmd_Save_Click() 
'Nur Adden wenn, beide Passwörter übereinstimmen
With Me
  If Not IsNull(.txtPass) And (.txtPass = .txtPass_Bestät) Then
    DoCmd.Close acForm, Me.Name, acSaveYes
  
 ElseIf IsNull(.txtPass) Or IsNull(.txtPass_Bestät) _
 Or (.txtPass <> .txtPass_Bestät) Then
      Me.Undo
    MsgBox "Das Passwort bitte richtig eingeben!"
 End If
End With
End Sub

################################### fuelleKombinationsfeld
Sub fuelleKombinationsfeld()
  Dim strSql As String
DoCmd.OpenForm "LoginForm"
strSql = "SELECT tblBenutzer.bzrLogin, tblBenutzer.bzrPass FROM tblBenutzer;"
With Forms("LoginForm").cboUsername
    .RowSource = strSql
    .ColumnCount = 1
  End With
End Sub

################################   #################################
Option Compare Database

Private Sub cboUsername_AfterUpdate()
  Me.txtPassword.SetFocus
End Sub



Sub versteckePasswort()
  DoCmd.OpenForm "LoginForm"
 
  With Forms("LoginForm")
    .txtPassword.InputMask = "Password"
  End With
End Sub

Private Sub cmd_close_Click()
 DoCmd.Close 'Schließt das Login-Formular
 DoCmd.Quit 'Schließt die komplette Access-Umgebung
End Sub

Private Sub cmd_login_Click()
Dim logID As Long
Dim strCboPass As String
Dim strPass As String
Dim username As String
 
On Error GoTo Error_Handler
 
If Len(Trim(Me.txtPassword)) > 0 Then
  strCboPass = Me.cboUsername.Column(1)
  strPass = Me.txtPassword.Value
  username = Me.cboUsername
 
  If strCboPass = strPass Then
   DoCmd.Close acForm, Me.Name
  Else
   Me.lblStatus.Visible = True
   With Me.txtPassword
    .Value = vbNullString
    .SetFocus
   End With
  End If
 ElseIf Len(Trim(Me.txtPassword)) = 0 Then
  MsgBox "Sie haben Ihr Passwort nicht eingegeben", vbInformation, "Passwort eingeben, bitte!"
 End If
 
Exit_Procedure:
 DoCmd.SetWarnings True
 Exit Sub
 
Error_Handler:
If IsNull(Me.cboUsername) Then
  MsgBox ("Sie haben Ihren Benutzernamen nicht eingegeben")
 Else
  MsgBox (Me.cboUsername.Value & " ist kein berechtigter Benutzer")
 End If
 Me.cboUsername.Value = vbNullString 'Null
 Me.cboUsername.SetFocus
 Me.txtPassword.Value = vbNullString
 Resume Exit_Procedure
End Sub

Private Sub Form_Load()
    Call MdlLogin.fuelleKombinationsfeld
End Sub

#################################################################################################################
SELECT tblMitarbeiter.StammNr, tblMitarbeiter.Nachname, tblMitarbeiter.Vorname, tblMitarbeiter.DKXKennung, tblMitarbeiter.Email, tblRechtseinheit.RE, tblMitKreis.MitKreis, tblKostenstelle.Kostenstelle, tblFzgGrp.FzgGrp, tblAbteilung.Abteilung, tblAbteilung.Abteilung
FROM tblAbteilung INNER JOIN ((tblKostenstelle INNER JOIN (tblMitKreis INNER JOIN (tblRechtseinheit INNER JOIN tblMitarbeiter ON tblRechtseinheit.REID = tblMitarbeiter.REID) ON tblMitKreis.MitKrID = tblMitarbeiter.MitKreisID) ON tblKostenstelle.KstID = tblMitarbeiter.KstID) INNER JOIN (tblFzgGrp INNER JOIN tblVerknuepftMit_FzgGrp ON tblFzgGrp.FzgGrpID = tblVerknuepftMit_FzgGrp.FzgGrpID) ON tblMitarbeiter.MitID = tblVerknuepftMit_FzgGrp.MitID) ON tblAbteilung.AbtID = tblKostenstelle.AbtID;



