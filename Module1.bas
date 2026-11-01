Attribute VB_Name = "Module1"
Option Explicit

Sub SAP_Final_Solution()
    Dim wsInput As Worksheet
    Dim SapGuiAuto As Object, SAPGUIApp As Object, Connection As Object, Session As Object
    Dim i As Integer
    
    ' --- Variables for Dynamic Loop ---
    Dim lastRow As Long
    Dim strMaterial As String, strMRPArea As String
    Dim statusMsg As String
    Dim tbl As Object, r As Long, visRow As Long, foundRow As Long
    Dim mrpText As String
    Dim forceEnter As Integer
    Dim md04_sub As String
    Dim cIcon As Integer
    Dim detailClicked As Boolean

    ' 1. INITIALIZE SAP CONNECTION
    On Error Resume Next
    Set SapGuiAuto = GetObject("SAPGUI")
    If SapGuiAuto Is Nothing Then
        MsgBox "SAP GUI is not running. Please log in to SAP and try again.", vbCritical, "SAP Connection Error"
        Exit Sub
    End If
    
    Set SAPGUIApp = SapGuiAuto.GetScriptingEngine
    Set Connection = SAPGUIApp.Children(0)
    Set Session = Connection.Children(0)
    On Error GoTo 0
    
    If Session Is Nothing Then
        MsgBox "Could not attach to an active SAP session.", vbCritical, "SAP Connection Error"
        Exit Sub
    End If

    ' 2. INITIALIZE EXCEL LOOP
    Set wsInput = ActiveSheet
    lastRow = wsInput.Cells(wsInput.Rows.Count, "A").End(xlUp).Row
    wsInput.Range("A2:C" & lastRow).Interior.ColorIndex = xlNone
    
    ' 3. MAIN PROCESSING LOOP
    For i = 2 To lastRow
        strMaterial = Trim(wsInput.Cells(i, 1).Value)
        strMRPArea = Trim(wsInput.Cells(i, 2).Value)
        
        If Len(strMaterial) = 0 Then GoTo NextRow
        
        On Error GoTo RowErrorHandler

        If Len(strMRPArea) = 0 Then Err.Raise vbObjectError + 10, "InputValidation", "MRP area is blank."

        ' --- START MD04 AUTOMATION ---
        Session.findById("wnd[0]").maximize
        Session.findById("wnd[0]/tbar[0]/okcd").Text = "/NMD04"
        Session.findById("wnd[0]/tbar[0]/btn[0]").press
        
        Session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-MATNR").Text = strMaterial
        Session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-BERID").Text = strMRPArea
        Session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-WERKS").Text = strMRPArea
        Session.findById("wnd[0]/tbar[0]/btn[0]").press
        
        ' --- CRITICAL FIX: SMART FILTER TOGGLE ---
        ' Only presses the filter button if the dropdown is currently hidden
        On Error Resume Next
        If TypeName(Session.findById("wnd[0]/usr/subINCLUDE12XX:SAPMM61R:1200/cmbRM61R-FILBZ", False)) = "Nothing" Then
            Session.findById("wnd[0]/tbar[1]/btn[29]").press
        End If
        Session.findById("wnd[0]/usr/subINCLUDE12XX:SAPMM61R:1200/cmbRM61R-FILBZ").SetFocus
        Session.findById("wnd[0]/usr/subINCLUDE12XX:SAPMM61R:1200/cmbRM61R-FILBZ").Key = "SAP00001"
        Err.Clear
        On Error GoTo RowErrorHandler
        
        ' --- CRITICAL FIX: DYNAMIC MD04 TABLE ID ---
        md04_sub = GetMD04Sub(Session)
        If Len(md04_sub) = 0 Then Err.Raise vbObjectError + 5, "SAP_Macro", "Could not identify the MD04 table container."
        
        ' --- STEP 1: DYNAMICALLY FIND AND CONVERT PLA.OR TO PR ---
        Set tbl = Session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:" & md04_sub & "/tblSAPMM61RTC_EZ")
        foundRow = -1
        
        For r = 0 To tbl.RowCount - 1
            If r >= tbl.verticalScrollbar.Position + tbl.visibleRowCount Then
                tbl.verticalScrollbar.Position = r
            End If
            visRow = r - tbl.verticalScrollbar.Position
            
            On Error Resume Next
            mrpText = Session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:" & md04_sub & "/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2," & visRow & "]").Text
            Err.Clear
            On Error GoTo RowErrorHandler
            
            If InStr(1, UCase(mrpText), "PLA.OR") > 0 Or InStr(1, UCase(mrpText), "PLDORD") > 0 Then
                foundRow = visRow
                Exit For
            End If
        Next r
        
        If foundRow = -1 Then Err.Raise vbObjectError + 1, "SAP_Macro", "Could not find 'PLA.OR' in the MD04 list."
        
' --- CRITICAL FIX: DYNAMIC ICON CLICK ---
        ' Scans columns 8 to 15 to find the magnifying glass if columns shift
        detailClicked = False
        On Error Resume Next
        For cIcon = 8 To 15
            If TypeName(Session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:" & md04_sub & "/tblSAPMM61RTC_EZ/btnMDEZ-DETAIL_ICON[" & cIcon & "," & foundRow & "]", False)) <> "Nothing" Then
                Session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:" & md04_sub & "/tblSAPMM61RTC_EZ/btnMDEZ-DETAIL_ICON[" & cIcon & "," & foundRow & "]").SetFocus
                Session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:" & md04_sub & "/tblSAPMM61RTC_EZ/btnMDEZ-DETAIL_ICON[" & cIcon & "," & foundRow & "]").press
                detailClicked = True
                Exit For
            End If
        Next cIcon
        On Error GoTo RowErrorHandler
        If Not detailClicked Then Err.Raise vbObjectError + 2, "SAP_Macro", "Could not find Detail Icon for PLA.OR"
        
        Session.findById("wnd[1]/tbar[0]/btn[27]").press ' Convert to PR
        
        ' --- ADDED: TOGGLE TO EDIT MODE ---
        ' Ensures the PR screen is in Change mode before trying to save
        Session.findById("wnd[0]/tbar[1]/btn[9]").press
        
        Session.findById("wnd[0]/tbar[0]/btn[11]").press ' Save PR
        
        ' Bypass Date/Firmed Warnings safely
        On Error Resume Next
        If TypeName(Session.findById("wnd[1]", False)) <> "Nothing" Then Session.findById("wnd[1]/tbar[0]/btn[0]").press
        If Session.findById("wnd[0]/sbar").MessageType = "W" Or Session.findById("wnd[0]/sbar").MessageType = "I" Then
            Session.findById("wnd[0]").sendVKey 0
            Session.findById("wnd[0]/tbar[0]/btn[11]").press
        End If
        Err.Clear
        On Error GoTo RowErrorHandler
        
        Session.findById("wnd[0]/tbar[1]/btn[6]").press ' Refresh MD04
        
        ' --- STEP 2: DYNAMICALLY RE-SCAN FOR THE NEW PURREQ ---
        md04_sub = GetMD04Sub(Session) ' Re-verify container ID after refresh
        If Len(md04_sub) = 0 Then Err.Raise vbObjectError + 5, "SAP_Macro", "Could not identify the MD04 table container after refresh."
        Set tbl = Session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:" & md04_sub & "/tblSAPMM61RTC_EZ")
        foundRow = -1
        
        For r = 0 To tbl.RowCount - 1
            If r >= tbl.verticalScrollbar.Position + tbl.visibleRowCount Then
                tbl.verticalScrollbar.Position = r
            End If
            visRow = r - tbl.verticalScrollbar.Position
            
            On Error Resume Next
            mrpText = Session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:" & md04_sub & "/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2," & visRow & "]").Text
            Err.Clear
            On Error GoTo RowErrorHandler
            
            If InStr(1, UCase(mrpText), "PURREQ") > 0 Then
                foundRow = visRow
                Exit For
            End If
        Next r
        
        If foundRow = -1 Then Err.Raise vbObjectError + 3, "SAP_Macro", "Could not find 'PURREQ' after conversion."
        
        detailClicked = False
        On Error Resume Next
        For cIcon = 8 To 15
            If TypeName(Session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:" & md04_sub & "/tblSAPMM61RTC_EZ/btnMDEZ-DETAIL_ICON[" & cIcon & "," & foundRow & "]", False)) <> "Nothing" Then
                Session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:" & md04_sub & "/tblSAPMM61RTC_EZ/btnMDEZ-DETAIL_ICON[" & cIcon & "," & foundRow & "]").SetFocus
                Session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:" & md04_sub & "/tblSAPMM61RTC_EZ/btnMDEZ-DETAIL_ICON[" & cIcon & "," & foundRow & "]").press
                detailClicked = True
                Exit For
            End If
        Next cIcon
        On Error GoTo RowErrorHandler
        If Not detailClicked Then Err.Raise vbObjectError + 4, "SAP_Macro", "Could not find Detail Icon for PURREQ"
        
        Session.findById("wnd[1]/tbar[0]/btn[28]").press ' Navigate to ME21N
        
        
        ' ====================================================================
        ' --- START ME21N LOGIC ---
        ' ====================================================================
        
        ' Document Overview Bypass
        On Error Resume Next
        Session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell[1]").selectItem "          3", "&Hierarchy"
        Session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell[1]").ensureVisibleHorizontalItem "          3", "&Hierarchy"
        Session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell[1]").topNode = "          1"
        Session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell[0]").pressButton "COPY"
        If TypeName(Session.findById("wnd[1]/tbar[0]/btn[0]", False)) <> "Nothing" Then
            Session.findById("wnd[1]/tbar[0]/btn[0]").press
        End If
        Err.Clear

        ' ====================================================================
        ' THE "5X ENTER" WARNING CLEAR
        ' ====================================================================
        On Error Resume Next
        For forceEnter = 1 To 5
            Session.findById("wnd[0]").sendVKey 0
            If TypeName(Session.findById("wnd[1]", False)) <> "Nothing" Then
                Session.findById("wnd[1]/tbar[0]/btn[0]").press
            End If
        Next forceEnter
        Err.Clear
        On Error GoTo ME21N_Warning_Handler

        Dim retryCount As Integer
        retryCount = 0

        ' ====================================================================
        ' YOUR EXACT RECORDED ME21N LOGIC (Protected by ExecuteSAP)
        ' ====================================================================
        
        ' 1. Delivery Schedule Tab
        ExecuteSAP Session, "wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT5", "SELECT"
        ExecuteSAP Session, "wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT5/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1320/tblSAPLMEGUITC_1320/ctxtMEPO1320-UZEIT[4,0]", "TEXT", "08:00:00"
        ExecuteSAP Session, "wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT5/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1320/tblSAPLMEGUITC_1320/ctxtMEPO1320-UZEIT[4,0]", "SETFOCUS"
        ExecuteSAP Session, "wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT5/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1320/tblSAPLMEGUITC_1320/ctxtMEPO1320-UZEIT[4,0]", "CARETPOSITION", "8"
        
        ' 2. Invoice Tab
        ExecuteSAP Session, "wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7", "SELECT"
        ExecuteSAP Session, "wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/chkMEPO1317-REPOS", "SELECTED", "True"
        ExecuteSAP Session, "wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/chkMEPO1317-REPOS", "SETFOCUS"
        Session.findById("wnd[0]").sendVKey 2
        ExecuteSAP Session, "wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/chkMEPO1317-WEBRE", "SELECTED", "True"
        ExecuteSAP Session, "wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/chkMEPO1317-WEBRE", "SETFOCUS"
        Session.findById("wnd[0]").sendVKey 2
        
        ' 3. Confirmations Tab
        ExecuteSAP Session, "wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT17", "SELECT"
        ExecuteSAP Session, "wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT17/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB1:SAPLMEGUI:1334/chkMEPO1334-KZABS", "SELECTED", "True"
        ExecuteSAP Session, "wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT17/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB1:SAPLMEGUI:1334/chkMEPO1334-KZABS", "SETFOCUS"
        Session.findById("wnd[0]").sendVKey 2
        
        ' --- SMART TOGGLE (Prevents the Header from collapsing if already open) ---
        Dim headerTabs As Object
        Set headerTabs = SafeGet(Session, "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL")
        If headerTabs Is Nothing Then
            ExecuteSAP Session, "wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB1:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4000/btnDYN_4000-BUTTON", "PRESS"
        End If
        ' -------------------------------------------------------------------------

        ' 4. Texts Tab
        ExecuteSAP Session, "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT3", "SELECT"
        ' Public-repository placeholder: replace with an approved internal recipient and message.
        ExecuteSAP Session, "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT3/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1230/subTEXTS:SAPLMMTE:0100/subEDITOR:SAPLMMTE:0101/cntlTEXT_EDITOR_0101/shellcont/shell", "TEXT", "Please confirm the order acknowledgement and order confirmation to procurement@example.invalid" + vbCr + "Order Acknowledgment: Confirms receipt of the purchase order" + vbCr + "Order Confirmation: Confirms the ETA and delivery date of the materials" + vbCr + ""
        ExecuteSAP Session, "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT3/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1230/subTEXTS:SAPLMMTE:0100/subEDITOR:SAPLMMTE:0101/cntlTEXT_EDITOR_0101/shellcont/shell", "SETSELECTIONINDEXES", "244,244"
        
        ' 5. Communication Tab
        ExecuteSAP Session, "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT5", "SELECT"
        ' Public-repository placeholders: replace with approved process values.
        ExecuteSAP Session, "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT5/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1225/txtMEPOCOMM-IHREZ", "TEXT", "DEMO-ORDER"
        ExecuteSAP Session, "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT5/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1225/txtMEPOCOMM-UNSEZ", "TEXT", "DEMO-PROC"
        ExecuteSAP Session, "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT5/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1225/txtMEPOCOMM-UNSEZ", "SETFOCUS"
        ExecuteSAP Session, "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT5/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1225/txtMEPOCOMM-UNSEZ", "CARETPOSITION", "3"
        
        GoTo ME21N_Success

ME21N_Warning_Handler:
        If retryCount < 4 Then
            retryCount = retryCount + 1
            On Error Resume Next
            Session.findById("wnd[0]").sendVKey 0
            If TypeName(Session.findById("wnd[1]", False)) <> "Nothing" Then
                Session.findById("wnd[1]/tbar[0]/btn[0]").press
            End If
            On Error GoTo 0
            Err.Clear
            Resume
        Else
            retryCount = 0
            On Error GoTo RowErrorHandler
            Err.Raise vbObjectError + 20, "ME21N", "ME21N warning could not be cleared after four attempts."
        End If

ME21N_Success:
        On Error GoTo RowErrorHandler

        Session.findById("wnd[0]/tbar[0]/btn[11]").press
        
        ' --- FINAL SAVE WARNING LOOP ---
        Dim saveLoop As Integer
        saveLoop = 0
        On Error Resume Next
        Do While saveLoop < 10
            If TypeName(Session.findById("wnd[1]", False)) <> "Nothing" Then
                Session.findById("wnd[1]/tbar[0]/btn[0]").press
            End If
            If Session.findById("wnd[0]/sbar").MessageType = "W" Or Session.findById("wnd[0]/sbar").MessageType = "I" Then
                Session.findById("wnd[0]").sendVKey 0
            Else
                Exit Do
            End If
            saveLoop = saveLoop + 1
        Loop
        Err.Clear
        On Error GoTo RowErrorHandler
        
        ' Clean focus from recording
        On Error Resume Next
        Session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:" & md04_sub & "/tblSAPMM61RTC_EZ/txtMDEZ-EXTRA[5,2]").SetFocus
        Session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:" & md04_sub & "/tblSAPMM61RTC_EZ/txtMDEZ-EXTRA[5,2]").caretPosition = 0
        Err.Clear

        statusMsg = Session.findById("wnd[0]/sbar").Text
        wsInput.Cells(i, 3).Value = ExtractPONumber(statusMsg)
        
        GoTo NextRow

RowErrorHandler:
        wsInput.Rows(i).Interior.Color = vbYellow
        wsInput.Cells(i, 3).Value = "ERROR: " & Err.Description
        Err.Clear
        
        On Error Resume Next
        If TypeName(Session.findById("wnd[1]", False)) <> "Nothing" Then
            Session.findById("wnd[1]/tbar[0]/btn[12]").press
            Session.findById("wnd[1]/tbar[0]/btn[0]").press
        End If
        Err.Clear

NextRow:
        On Error GoTo 0
    Next i
    
    MsgBox "Processing Complete.", vbInformation, "SAP Automation"

End Sub

' =================================================================================
' --- HELPER FUNCTION: GET DYNAMIC MD04 TABLE ID ---
' Safely detects if the MD04 container is 0750, 0740, etc., based on column layout
' =================================================================================
Function GetMD04Sub(Session As Object) As String
    Dim i As Integer
    Dim testStr As String
    On Error Resume Next
    For i = 700 To 790 Step 10
        testStr = "0" & i
        If TypeName(Session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:" & testStr & "/tblSAPMM61RTC_EZ", False)) <> "Nothing" Then
            GetMD04Sub = testStr
            Exit Function
        End If
    Next i
    For i = 700 To 799
        testStr = "0" & i
        If TypeName(Session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:" & testStr & "/tblSAPMM61RTC_EZ", False)) <> "Nothing" Then
            GetMD04Sub = testStr
            Exit Function
        End If
    Next i
    GetMD04Sub = vbNullString ' Do not guess a screen container when none was found
End Function

' =================================================================================
' --- HELPER FUNCTION: EXECUTE SAP WRAPPER ---
' =================================================================================
Sub ExecuteSAP(Session As Object, path As String, action As String, Optional val As String = "")
    Dim obj As Object
    Set obj = SafeGet(Session, path)
    
    If Not obj Is Nothing Then
        Select Case UCase(action)
            Case "SELECT": obj.Select
            Case "TEXT": obj.Text = val
            Case "SETFOCUS": obj.SetFocus
            Case "CARETPOSITION": obj.caretPosition = CInt(val)
            Case "SELECTED": obj.Selected = CBool(val)
            Case "PRESS": obj.press
            Case "SETSELECTIONINDEXES"
                Dim idxs() As String
                idxs = Split(val, ",")
                obj.setSelectionIndexes CInt(idxs(0)), CInt(idxs(1))
            Case Else
                Err.Raise vbObjectError + 2, "ExecuteSAP", "Unsupported SAP action: " & action
        End Select
    Else
        Err.Raise vbObjectError + 1, "ExecuteSAP", "Element not found or locked by warning"
    End If
End Sub

' =================================================================================
' --- HELPER FUNCTION: MEMORY CACHE & PERMUTATION MATRIX ---
' =================================================================================
Function SafeGet(Session As Object, fullRecordedPath As String) As Object
    Static PathCache As Object
    If PathCache Is Nothing Then Set PathCache = CreateObject("Scripting.Dictionary")

    Dim i As Integer, j As Integer, k As Integer
    Dim testPath As String, tempPath As String
    Dim result As Object

    On Error Resume Next

    If PathCache.Exists(fullRecordedPath) Then
        Set result = Session.findById(PathCache(fullRecordedPath), False)
        If Not result Is Nothing Then
            Set SafeGet = result
            Exit Function
        End If
        PathCache.Remove fullRecordedPath
    End If

    For i = 10 To 30
        testPath = Replace(fullRecordedPath, "SAPLMEGUI:0010", "SAPLMEGUI:00" & i)
        testPath = Replace(testPath, "SAPLMEGUI:0013", "SAPLMEGUI:00" & i)
        testPath = Replace(testPath, "SAPLMEGUI:0015", "SAPLMEGUI:00" & i)
        testPath = Replace(testPath, "SAPLMEGUI:0019", "SAPLMEGUI:00" & i)

        Set result = Session.findById(testPath, False)
        If Not result Is Nothing Then
            PathCache.Add fullRecordedPath, testPath
            Set SafeGet = result
            Exit Function
        End If
    Next i

    For i = 10 To 30
        testPath = Replace(fullRecordedPath, "SAPLMEGUI:0010", "SAPLMEGUI:00" & i)
        testPath = Replace(testPath, "SAPLMEGUI:0013", "SAPLMEGUI:00" & i)
        testPath = Replace(testPath, "SAPLMEGUI:0019", "SAPLMEGUI:00" & i)

        For j = 1 To 4
            For k = 1 To 4
                tempPath = testPath
                tempPath = Replace(tempPath, "subSUB3:SAPLMEVIEWS:1100", "subSUB" & j & ":SAPLMEVIEWS:1100")
                tempPath = Replace(tempPath, "subSUB2:SAPLMEVIEWS:1100", "subSUB" & j & ":SAPLMEVIEWS:1100")
                tempPath = Replace(tempPath, "subSUB1:SAPLMEVIEWS:1100", "subSUB" & j & ":SAPLMEVIEWS:1100")
                
                tempPath = Replace(tempPath, "subSUB2:SAPLMEVIEWS:1200", "subSUB" & k & ":SAPLMEVIEWS:1200")
                tempPath = Replace(tempPath, "subSUB1:SAPLMEVIEWS:1200", "subSUB" & k & ":SAPLMEVIEWS:1200")

                Set result = Session.findById(tempPath, False)
                If Not result Is Nothing Then
                    PathCache.Add fullRecordedPath, tempPath
                    Set SafeGet = result
                    Exit Function
                End If
            Next k
        Next j
    Next i

    Set SafeGet = Nothing
End Function

' --- HELPER FUNCTION: REGEX PO EXTRACTION ---
Function ExtractPONumber(statusText As String) As String
    Dim regEx As Object
    Dim matches As Object
    
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = "\b\d{10}\b"
    regEx.Global = False
    
    If regEx.Test(statusText) Then
        Set matches = regEx.Execute(statusText)
        ExtractPONumber = matches(0).Value
    Else
        ExtractPONumber = statusText
    End If
End Function
