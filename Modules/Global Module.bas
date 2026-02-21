Attribute VB_Name = "Global Module"
Option Compare Database
Public Const DUNS = "623351603"
Public Const BuyerMPID = "1acdaafa-addd-4269-9058-2523814333f0" 'New Boeing (Buyer) MPID for Production 08/26/2025
'Public Const BuyerMPID = "a1d8e6d8-7802-1000-bfb4-ac16042a0001"        'Boeing (Buyer) MPID  for Production ##OLD MPID Number as of 08/26/2025
'Public Const BuyerMPID = "e78ab758-78a0-1000-b1a4-0a1c0c090001"         'Test
Public Const SupplierMPID = "65bbf4ce-e9bb-4112-a577-7c140376e200"      'Exostar MPID

Public MyUsername
Public MyPassword
Option Explicit
Option Base 1

Public Function PosOfFirstDigit(strTest As String) As Long

Dim i As Long

PosOfFirstDigit = 0
For i = 1 To Len(strTest)
If Mid$(strTest, i, 1) Like "#" Then
PosOfFirstDigit = i
Exit For
End If
Next

End Function

Public Function PosOfLastDigit(strTest As String) As Long
Dim i As Long
For i = Len(strTest) To 1 - 1
If Mid$(strTest, i, 1) Like "#" Then
PosOfLastDigit = i
Exit For
End If
Next

End Function

Public Function KeyCode(Password As String) As Long
   ' This function will produce a unique key for the
   ' string that is passed in as the Password.
   Dim i As Integer
   Dim Hold As Long

   For i = 1 To Len(Password)
      Select Case (Asc(Left(Password, 1)) * i) Mod 4
      Case Is = 0
         Hold = Hold + (Asc(Mid(Password, i, 1)) * i)
      Case Is = 1
         Hold = Hold - (Asc(Mid(Password, i, 1)) * i)
      Case Is = 2
         Hold = Hold + (Asc(Mid(Password, i, 1)) * _
            (i - Asc(Mid(Password, i, 1))))
      Case Is = 3
         Hold = Hold - (Asc(Mid(Password, i, 1)) * _
            (i + Len(Password)))
   End Select
   Next i
   KeyCode = Hold
End Function

Public Function Dispatch(StrInput As String) As String
'Add prefix to Dispatchnote input
 Dim IntChkLen As Integer
 IntChkLen = Len(StrInput)  'Assign the length of input to an integer
  Select Case IntChkLen
    Case 1          'Input string length is one character
        Dispatch = "00000000000000" & StrInput
    Case 2          'Input string length is two characters
        Dispatch = "0000000000000" & StrInput
    Case 3
        Dispatch = "000000000000" & StrInput
    Case 4
        Dispatch = "00000000000" & StrInput
    Case 5
        Dispatch = "0000000000" & StrInput
    Case 6
        Dispatch = "000000000" & StrInput
    Case 7
        Dispatch = "00000000" & StrInput
    Case 8
        Dispatch = "0000000" & StrInput
    Case 9
        Dispatch = "000000" & StrInput
    Case 10
        Dispatch = "00000" & StrInput
    Case 11
        Dispatch = "0000" & StrInput
    Case 12
        Dispatch = "000" & StrInput
    Case 13
        Dispatch = "00" & StrInput
    Case 14
        Dispatch = "0" & StrInput
    Case Else
        Dispatch = StrInput
  End Select
End Function

Public Function Job(StrInput As String) As String
'Add prefix to Dispatchnote input
 Dim IntChkLen As Integer
 IntChkLen = Len(StrInput)  'Assign the length of input to an integer
  Select Case IntChkLen
    Case 1          'Input string length is one character
        Job = "00000000000000" & StrInput
    Case 2          'Input string length is two characters
        Job = "0000000000000" & StrInput
    Case 3
        Job = "000000000000" & StrInput
    Case 4
        Job = "00000000000" & StrInput
    Case 5
        Job = "0000000000" & StrInput
    Case 6
        Job = "000000000" & StrInput
    Case 7
        Job = "00000000" & StrInput
    Case 8
        Job = "0000000" & StrInput
    Case 9
        Job = "000000" & StrInput
    Case 10
        Job = "00000" & StrInput
    Case 11
        Job = "0000" & StrInput
    Case 12
        Job = "000" & StrInput
    Case 13
        Job = "00" & StrInput
    Case 14
        Job = "0" & StrInput
    Case Else
        Job = StrInput
  End Select
End Function

Public Function CheckDataExist(StrSQLf As String) As Boolean
   'Check Query data's existence

    Dim CurDB As Database, cnn As ADODB.Connection                             'Create object variables
    Dim rst As ADODB.Recordset, strCnn As String
    
   'Create Valid object reference
    Set CurDB = CurrentDb
   
    'Open connection.
    strCnn = "Provider=Microsoft.Jet.OLEDB.4.0; " & _
            "Data Source=" & CurDB.Name
    Set cnn = New ADODB.Connection
    'cnn.Open strCnn
    cnn.ConnectionString = strCnn
701833
cnn.Open CurrentProject.Connection

    'Open Records
    Set rst = New ADODB.Recordset
    rst.Open StrSQLf, cnn, adOpenKeyset, adLockOptimistic, adCmdText

    'Check recordset's row count
    If rst.RecordCount > 0 Then
        CheckDataExist = True           'Data exist, return true
    Else
        CheckDataExist = False          'Data non-exist, return false
    End If
    rst.Close
    cnn.Close
    Set rst = Nothing
    Set cnn = Nothing
    Set CurDB = Nothing
End Function

Public Function CheckDataExist2(StrSQLf As String) As Boolean
   'Check Query data's existence
    
    'Dim CurDB As Database, cnn As ADODB.Connection                             'Create object variables
    'Dim rst As ADODB.Recordset, strCnn As String
    
   'Create Valid object reference
    'Set CurDB = CurrentDb
   
    'Open connection.
    'strCnn = "Provider=Microsoft.Jet.OLEDB.4.0; " & _
            "Data Source=" & CurDB.Name
    'Set cnn = New ADODB.Connection
    'cnn.Open strCnn
    'cnn.ConnectionString = strCnn
    'cnn.Open CurrentProject.Connection

    'Open Records
    'Set rst = New ADODB.Recordset
    'rst.Open StrSQLf, cnn, adOpenKeyset, adLockOptimistic, adCmdText
   Dim rst As ADODB.Recordset
   Set rst = New ADODB.Recordset
   rst.ActiveConnection = CurrentProject.Connection
   rst.CursorType = adOpenKeyset
   rst.LockType = adLockOptimistic
   'rst.CommandType = adCmdText
   rst.Open StrSQLf
   
    'Check recordset's row count
    If rst.RecordCount > 0 Then
        CheckDataExist2 = True           'Data exist, return true
    Else
        CheckDataExist2 = False          'Data non-exist, return false
    End If
    rst.Close
    'cnn.Close
    Set rst = Nothing
    'Set cnn = Nothing
    'Set CurDB = Nothing
End Function

Public Function printsql()
Dim strSQL As String
strSQL = "SELECT dbo_MdnMaster.DispatchNote, dbo_MdnMaster.CustomerPoNumber, dbo_MdnMaster.Customer , " & _
                    "dbo_MdnMaster.DispatchName, dbo_MdnMaster.DispatchAddress1, dbo_MdnMaster.DispatchAddress2, dbo_MdnMaster.DispatchAddress3, " & _
                    "dbo_MdnMaster.DispatchAddress4, dbo_MdnMaster.DispatchPostalCode, dbo_ArCustomer.Name, " & _
                    "dbo_ArCustomer.SoldToAddr1, dbo_ArCustomer.SoldToAddr2, dbo_ArCustomer.SoldToAddr3, " & _
                    "dbo_ArCustomer.SoldPostalCode, Mid$(dbo_MdnMaster.ShippingInstrs,3,20) AS ShipVia, " & _
                    "dbo_MdnMaster.SalesOrder, dbo_TblArTerms.Description, " & _
                    "dbo_MdnDetail.MOrderQty, dbo_MdnDetail.MStockQtyToShp, dbo_MdnDetail.MStockCode, " & _
                    "dbo_MdnDetail.MStockDes, dbo_MdnDetail.MCustRequestDat, StockCode, dbo_MdnDetail.MOrderUom, dbo_MdnDetailLot.Lot, " & _
                    "dbo_InvMaster.UserField1 " & _
                    "FROM dbo_InvMaster INNER JOIN (dbo_MdnDetailLot RIGHT JOIN (((dbo_MdnMaster INNER JOIN dbo_ArCustomer ON dbo_MdnMaster.Customer = dbo_ArCustomer.Customer) INNER JOIN dbo_TblArTerms ON dbo_ArCustomer.TermsCode = dbo_TblArTerms.TermsCode) INNER JOIN dbo_MdnDetail ON dbo_MdnMaster.DispatchNote = dbo_MdnDetail.DispatchNote) ON dbo_MdnDetailLot.DispatchNote = dbo_MdnMaster.DispatchNote) ON dbo_InvMaster.StockCode = dbo_MdnDetail.MStockCode " & _
                    "WHERE dbo_MdnMaster.DispatchNote= '" & Forms!Frm7PackingSlip![txtDispatchNote] & "' " & "AND dbo_MdnDetail.LineType= '" & 1 & " '" & "AND dbo_MDNMASTER.DispatchNoteStatus between '" & 5 & "' AND '" & 9 & "'"
            Debug.Print strSQL
End Function



