Dim strip As String
Dim danhapdataip As Boolean
Dim truycapduocATE As Boolean
Dim strSN As String
Dim strSNATE As String
Dim dungIPcuatramATE As Boolean
Dim daFindpart As String
Dim chuatimxong As String
Dim dung As String

  Sub btncheckip_Click()
    strip = ""
    strip = txtip.Text
    danhapdataip = False
    truycapduocATE = False
    strSN = ""
    'DOI SN O DAY DE SU DUNG CHO TRAM KHAC
    strSNATE = "S11077"
    dungIPcuatramATE = False
    
    Dim objFolder As Object
    Dim objFile As Object
    
    If strip = "" Then
        MsgBox "IP khong duoc rong", vbInformation
    Else
        danhapdataip = True
    End If
    
    If danhapdataip Then
        On Error Resume Next
        If Dir("\\" & strip & "\Reports\TEST GO-NOGO", vbDirectory) = "" Then
            MsgBox "IP khong phai cua tram ATE !", vbCritical
        Else
            truycapduocATE = True
        End If
    End If
    
    If truycapduocATE Then
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Set objFolder = objFSO.GetFolder("\\" & strip & "\Reports\TEST GO-NOGO\HTML")
        
        For Each objFile In objFolder.files
            With objFile
                strSN = Mid(objFile.Name, 2, 6)
            End With
            Exit For
        Next objFile
        
        Debug.Print strSN
    
        If strSN = strSNATE Then
            MsgBox "Da truy cap thanh cong !", vbInformation
            dungIPcuatramATE = True
            txtip.Locked = True
            btncheckip.Locked = True
        Else
            MsgBox "Khong dung IP cua tram ATE !", vbCritical
        End If
    End If
End Sub

Sub btnfoundpart_Click()
    Dim folder As Object
    Dim i As Integer
    
    daFindpart = ThisWorkbook.Sheets(2).Range("B1").Text
    i = 0
    
    If dungIPcuatramATE = False Then MsgBox "Vui long thuc hien kiem tra IP truoc !", vbCritical
    
    If dungIPcuatramATE Then
    
        If daFindpart = "True" Then
            MsgBox "Da tim part roi !", vbInformation
            txtmountpart.Text = ThisWorkbook.Sheets(2).Range("B2")
            btnfoundpart.Caption = "Found"
            btnfoundpart.BackColor = &HFF00&
        Else
            btnfoundpart.Caption = "Finding"
            btnfoundpart.BackColor = &H80FF&
            Application.Wait Now + TimeValue("00:00:01")
            
            Set objFSO = CreateObject("Scripting.FileSystemObject")
            Set objFolder = objFSO.GetFolder("\\" & strip & "\Reports")
            
            For Each folder In objFolder.SubFolders
                With folder
                    i = i + 1
                    ThisWorkbook.Sheets(2).Range("D" & i) = i
                    ThisWorkbook.Sheets(2).Range("E" & i) = folder.Name
                    Debug.Print i & " -> " & folder.Name
                    
                    If i Mod 10 = 0 Then
                        txtmountpart.Text = i
'                        Application.Wait Now + TimeValue("00:00:01")
                        ThisWorkbook.Save
                    End If
                End With
            Next folder
            
            daFindpart = "True"
            ThisWorkbook.Sheets(2).Range("B1") = daFindpart
            txtmountpart.Text = ThisWorkbook.Sheets(2).Range("B2")
            ThisWorkbook.Save
            MsgBox "Find Part thanh cong", vbInformation
            
            btnfoundpart.Caption = "Found"
            btnfoundpart.BackColor = &HFF00&
        End If
    End If
End Sub

'CONTROL CHO WORKBOOK BACK DE COI SU THAY DOI CUA TEXTBOX
'MO CHON TASK MANAGER TRONG KHI CHAY
Sub btnfindpath_Click()
    Dim folder As Object
    Dim folderhtml As Object
    Dim file As Object
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim demsoluongfiletrongmotfoler As Long
    i = 0
    j = CLng(ThisWorkbook.Sheets("Quet Part").Range("H3").Text)

    Dim dungquet As String
    Dim daquetxongfile As String
    
    daquetxongfile = "False"
    ThisWorkbook.Sheets("Upload").Range("B1") = daquetfilexong
    
    Dim dakiemtramocpart As Boolean
    dakiemtramocpart = False
    Dim dakiemtramocfile As Boolean
    dakiemtramocfile = False
    'KIEM TRA DA FIND PART CHUA
    daFindpart = ThisWorkbook.Sheets(2).Range("B1").Text
    
    If dungIPcuatramATE <> True Then MsgBox "Thuc hien Find Part truoc khi Find Paths File !", vbCritical
    If dungIPcuatramATE = True Then
    
        btnfindpath.Caption = "Finding ..."
        btnfindpath.BackColor = &H80FF&
        txtfilefound.Text = ThisWorkbook.Sheets("Quet Part").Range("H3").Text
        Application.Wait Now + TimeValue("00:00:03")
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Set objFolder = objFSO.GetFolder("\\" & strip & "\Reports")
        
        For Each folder In objFolder.SubFolders
            With folder
                i = i + 1
                If i = CLng(ThisWorkbook.Sheets("Quet Part").Range("H1").Text) Or dakiemtramocpart Then
                        dakiemtramocpart = True
                        Debug.Print i & " -> " & folder.Name
                        'GHI NHAN LAI DA QUET THU MUC PART NAO
                        ThisWorkbook.Sheets("Quet Part").Range("A" & i) = i
                        ThisWorkbook.Sheets("Quet Part").Range("B" & i) = folder.Name
                        txtpartaccessed.Text = ""
                        txtpartaccessed.Text = i
                        Application.Wait Now + TimeValue("00:00:01")
                        'LUU SO LUONG PART DA QUET VAO SHEET QUET PART
                        ThisWorkbook.Sheets("Quet Part").Range("H1") = i
                        ThisWorkbook.Save
                        
                        k = 0
                        demsoluongfiletrongmotfoler = 0
                        For Each folderhtml In folder.SubFolders
                            With folderhtml
                                'DEM SO LUONG FILE TRONG 1 FOLDER
                                For Each file In folderhtml.files
                                    demsoluongfiletrongmotfoler = demsoluongfiletrongmotfoler + 1
                                Next file
                                For Each file In folderhtml.files
                                    With file
                                        'DINH VI CAC FILE DA QUET
                                        k = k + 1
                                        If k = CLng(ThisWorkbook.Sheets("Quet Part").Range("H2").Text) Or dakiemtramocfile Then
                                            dakiemtramocfile = True
                                            'SUA NAM O DAY DE DOI NAM KHAC
                                            'Debug.Print "Year: " & Year(CDate(file.DateLastModified)) & " Month: " & Month(CDate(file.DateLastModified))
                                            If (Year(CDate(file.DateLastModified)) = 2023) Then
                                                j = j + 1
                                                Debug.Print i & ", " & j & " " & folder.Name & " -> " & file.DateLastModified
                                                ThisWorkbook.Sheets("Quet Files").Range("A" & j) = i
                                                ThisWorkbook.Sheets("Quet Files").Range("B" & j) = j
                                                ThisWorkbook.Sheets("Quet Files").Range("C" & j) = folder.Name
                                                ThisWorkbook.Sheets("Quet Files").Range("D" & j) = file.DateLastModified
                                                ThisWorkbook.Sheets("Quet Files").Range("E" & j) = file.Path
                                                
                                                'LUU SO LUONG PART DA QUET VAO SHEET QUET PART
                                                ThisWorkbook.Sheets("Quet Part").Range("H1") = i
                                                'LUU SO LUONG FILE DA LAY PATH VAO SHEET QUET PART
                                                ThisWorkbook.Sheets("Quet Part").Range("H3") = j
                                                'LUU SO LUONG FILE DA QUET (CHI TRONG THU MUC PART TUONG UNG) VAO SHEET QUET PART - MATRIX 2 CHIEU i,k
                                                ThisWorkbook.Sheets("Quet Part").Range("H2") = k + 1
                                                
                                                If j Mod 100 = 0 Then
                                                    txtfilefound.Text = j
                                                    Application.Wait Now + TimeValue("00:00:01")
                                                    ThisWorkbook.Save
                                                End If
                                                
                                                If j Mod 100000 = 0 Then
                                                    txtfilefound.Text = j
'                                                    Application.Wait Now + TimeValue("00:00:01")
                                                    ThisWorkbook.Save
                                                    dungquet = InputBox("Type y/Y to stop finding:", "Stop box")
                                                    'THOAT KHOI VIEC QUET
                                                    If dungquet = "Y" Or dungquet = "y" Then
                                                        ThisWorkbook.Save
                                                        MsgBox "Da tim duoc " & j & " paths"
                                                        Exit For
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End With
                                Next file
                                'THOAT KHOI VIEC QUET
                                If dungquet = "Y" Or dungquet = "y" Then
                                    Exit For
                                End If
                            End With
                            
                            'THOAT KHOI VIEC QUET
                            If dungquet = "Y" Or dungquet = "y" Then
                                Exit For
                            End If
                            
                        Next folderhtml
                End If 'end if cua so sanh voi moc part
            End With
            
            'THOAT KHOI VIEC QUET
            If dungquet = "Y" Or dungquet = "y" Then
                Exit For
            End If
        Next folder
        
            
        'LUU SO LUONG PART DA QUET VAO SHEET QUET PART
        ThisWorkbook.Sheets("Quet Part").Range("H1") = i
        'LUU SO LUONG FILE DA LAY PATH VAO SHEET QUET PART
        ThisWorkbook.Sheets("Quet Part").Range("H3") = j
        'LUU SO LUONG FILE DA QUET VAO SHEET QUET PART
        ThisWorkbook.Sheets("Quet Part").Range("H2") = k
        
        'DA QUET FILE XONG
        'CHUA XAC DINH KHI NAO XOA XONG
        If i = CLng(ThisWorkbook.Sheets("Quet Part").Range("H1").Text) And k = demsoluongfiletrongmotfoler Then
            MsgBox "Da quet xong !"
            btnfindpath.Caption = "Find path completed"
            btnfindpath.Font.Size = 8
            btnfindpath.BackColor = &H80FF80
            daquetxongfile = "True"
            ThisWorkbook.Sheets("Upload").Range("B1") = daquetxongfile
        End If

        ThisWorkbook.Save
    End If
End Sub

Sub btnupload_Click()
    Dim i As Long
    Dim dungxoa As String
    Dim serverTheotram As String
    Dim dakiemtramocupload As Boolean
    
    serverTheotram = "\\192.168.78.226\TestDataStore\42.1 TEST DATA\4_ATE\09.ATE_4.3"
    dakiemtramocupload = False
    
    If ThisWorkbook.Sheets("Upload").Range("B1").Text = "False" Then MsgBox "Hoan thanh viec quet Files truoc khi Upload !", vbCritical
    If ThisWorkbook.Sheets("Upload").Range("B1").Text = "True" Then
        btnupload.Locked = True
        btnupload.Caption = "Uploading..."
        btnupload.BackColor = &H80FF&
        For i = 1 To CLng(ThisWorkbook.Sheets("Quet Part").Range("H3").Text) Step 1
            If i = CLng(ThisWorkbook.Sheets("Upload").Range("B2").Text) Or dakiemtramocupload Then
                dakiemtramocupload = True
                On Error Resume Next
                If Dir(ThisWorkbook.Sheets("Quet Files").Range("E" & i).Text) <> "" Then
                    ThisWorkbook.Sheets("Upload").Range("E" & i) = ThisWorkbook.Sheets("Quet Files").Range("E" & i)
                
                    'KIEM TRA NEU THU MUC PART CHUA CO THI TAO THU MUC TEN PART
                    If Dir(serverTheotram & "\" & ThisWorkbook.Sheets("Quet Files").Range("AH" & i).Text, vbDirectory) = "" Then
                        MkDir serverTheotram & "\" & ThisWorkbook.Sheets("Quet Files").Range("AH" & i).Text
                        MkDir serverTheotram & "\" & ThisWorkbook.Sheets("Quet Files").Range("AH" & i).Text & "\HTML"
                    End If
                    
                    'KIEM TRA FILE DA CO TREN SERVER CHUA, NEU CHUA THI COPY
                    Dim strDich As String
                    strDich = serverTheotram & "\" & ThisWorkbook.Sheets("Quet Files").Range("AH" & i).Text & "\HTML\" & ThisWorkbook.Sheets("Quet Files").Range("AJ" & i).Text
                    Debug.Print "Link:"; strDich
                    Debug.Print i & " Lenght: -> " & Len(strDich) & " ky tu"
                    
                    If (Len(strDich) <= 259) Then
                        If Dir(strDich, vbDirectory) = "" Then
                            FileCopy ThisWorkbook.Sheets("Upload").Range("E" & i).Text, strDich
                            If Dir(strDich, vbDirectory) <> "" Then Debug.Print i & " Uploaded -> okey"
                        Else
                            Debug.Print i & " Da co tren server"
                        End If
                    Else
                        MsgBox "Chieu dai link upload vuot qua so luong ky tu cho phep (259 ky tu) !!!"
                        Debug.Print "Chieu dai link upload vuot qua so luong ky tu cho phep (259 ky tu) !!!"
                        Exit Sub
                    End If
                    
                    'LUU SO LUONG PART DA UPLOAD VAO SHEET UPLOAD
                    ThisWorkbook.Sheets("Upload").Range("B2") = i + 1
                    
                    Application.Wait Now + TimeValue("00:00:02")
                    
                    If i Mod 50 = 0 Then
                        txtupload.Text = i
                        Application.Wait Now + TimeValue("00:00:01")
                        ThisWorkbook.Save
                    End If
                    
                    If i Mod 100000 = 0 Then
                        dungxoa = InputBox("Type y/Y to stop uploading:", "Stop box")
                        'THOAT KHOI VIEC XOA
                        If dungxoa = "Y" Or dungxoa = "y" Then
                            ThisWorkbook.Save
                            MsgBox "Da upload duoc " & i & " files"
                            Exit For
                        End If
                    End If
                End If
                
                If i = CLng(ThisWorkbook.Sheets("Quet Part").Range("H3").Text) Then
                    MsgBox "Da upload xong !", vbInformation
                    txtupload.Text = i
                    btnupload.Caption = "Upload completed"
                    btnupload.Font.Size = 8
                    btnupload.BackColor = &H80FF80
                End If
            End If
        Next i
    End If
End Sub
