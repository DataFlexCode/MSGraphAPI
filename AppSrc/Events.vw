Use Windows.pkg
Use DFClient.pkg
Use cCJGrid.pkg
Use cCJGridColumn.pkg
Use GraphStructs\stGraphAttachments.pkg
Use cLinkLabel.pkg
Use WinUuid.pkg
Use cCJGridColumnRowIndicator.pkg
Use cTextEdit.pkg
Use dfBitmap.pkg
Use seq_chnl.pkg
Use JsonConvFuncs.pkg

Register_Object oOAuth2
Register_Function psGrantScopes Returns String

Deferred_View Activate_oEvents for ;
Object oEvents is a dbView
    Property stGraphAttachments patAttachments
    Property Handle phoEvents

    Set Border_Style to Border_Thick
    Set Size to 285 843
    Set Location to 2 4
    Set Label to "MS Graph Events"
    
    Procedure SetLoginState
        Boolean bLoggedin
        
        Get IsLoggedIn                     to bLoggedIn
        Set Label       of oLoginStatus    to (If(bLoggedIn, "Logged into Graph", ;
                                                             "Not logged in"))
        Set TextColor   of oLoginStatus    to (If(bLoggedIn, clBlue, clRed))
        Set Label       of oLoginBtn       to (If(bLoggedIn, "Log out of Microsoft Graph", ;
                                                             "Log in to Microsoft Graph"))
    End_Procedure
    
    Procedure Activating
        Forward Send Activating
        Send SetLoginState
    End_Procedure
    
    Procedure OnOAuth2LoggedIn
        Send SetLoginState
    End_Procedure

    Procedure OnOAuth2NotLoggedIn
        Send SetLoginState
    End_Procedure
    
    Object oLoginBtn is a Button
        Set Size to 14 92
        Set Location to 3 3
        Set Label to "Login to Microsoft Graph"
    
        Procedure OnClick
            
            If (IsLoggedIn(Self)) ;
                Send Logout of (oOauth2(Self))
            Else Begin
                Set psGraphScopes of oGraph to (psGrantScopes(Self)) //"Calendars.ReadWrite"
                Set pbLoggingIn of oOAuthLoginDialog to True
                Send Login of (oOAuth2(Self))
            End
            
        End_Procedure
    
    End_Object
    
    Object oLoginStatus is a TextBox
        Set Size to 9 47
        Set Location to 6 104
        Set Label to "Not Logged In"
        Set TextColor to clRed
    End_Object
    
    Object oRefresh is a Button
        Set Location to 3 471
        Set Label to "Refresh"
        Set peAnchors to anTopRight
    
        Procedure OnClick
            Send LoadEvents of oEventsGrid
        End_Procedure
    
    End_Object
    
    Object oEventsGrid is a cCJGrid
        Set Size to 261 511
        Set Location to 20 9
        Set peAnchors to anAll
        Set pbReadOnly to True

        Object oCJGridColumnRowIndicator1 is a cCJGridColumnRowIndicator
            Set piWidth to 21
        End_Object

        Object oEventStart is a cCJGridColumn
            Set piWidth to 252
            Set psCaption to "Start"
        End_Object

        Object oEventEnd is a cCJGridColumn
            Set piWidth to 275
            Set psCaption to "End"
        End_Object

        Object oEventName is a cCJGridColumn
            Set piWidth to 938
            Set psCaption to "Event"
        End_Object
        
        Procedure LoadEvents
            tDataSourceRow[] atRows
            Integer iStartCol iEndCol iNameCol iMax i
            Handle  hoEvents
            String sTemp
            
            Get piColumnId of oEventStart to iStartCol
            Get piColumnId of oEventEnd   to iEndCol
            Get piColumnId of oEventName  to iNameCol
            
            Get ListEvents of oGraph "" "$orderby=start/dateTime desc&$top=10000" to hoEvents
            
            If not hoEvents ;
                Procedure_Return
            
            Set phoEvents to hoEvents

            Move (JsonCountAtPath(hoEvents, "value")) to iMax
            Decrement iMax
            
            For i from 0 to iMax
                Move (IsoDt2DfDt(JsonValueAtPath(hoEvents, ;
                     "value[" + String(i) + "].start.dateTime"))) to ;
                                                              atRows[i].sValue[iStartCol]
                Move (IsoDt2DfDt(JsonValueAtPath(hoEvents, ;
                     "value[" + String(i) + "].end.dateTime"))) to ;
                                                              atRows[i].sValue[iEndCol]
                Move (JsonValueAtPath(hoEvents, ;
                      "value[" + String(i) + "].subject")) to atRows[i].sValue[iNameCol]
            Loop
            
            Send InitializeData atRows
            Send MoveToFirstRow
        End_Procedure
        
        Procedure OnRowChanged Integer iOldRow Integer iNewSelectedRow
            Send SetEventValues iNewSelectedRow
        End_Procedure
        
    End_Object
    
    Procedure SetEventValues Integer iEvt
        Handle hoEvents
        String sType
        
        Get phoEvents to hoEvents
        
        Set Visible_State of oBody      to False
        Set Visible_State of oHtmlBody  to False
        Set Visible_State of oTextAtt   to False
        Set Visible_State of oBMPAtt    to False
        
        Set Value of oEvent    to ;
            (JsonValueAtPath(hoEvents, "value[" + String(iEvt) + "].bodyPreview"))
        Set Value of oReminder to ;
            (JsonValueAtPath(hoEvents, "value[" + String(iEvt) + "].reminderMinutesBeforeStart"))
        Set Value of oLocation to   ;
            (JsonValueAtPath(hoEvents, "value[" + String(iEvt) + "].location.displayName"))
        Set Checked_State of oAllday to ;
            (JsonValueAtPath(hoEvents, "value[" + String(iEvt) + "].isAllDay"))
        Set Checked_State of oCancelled to ;
            (JsonValueAtPath(hoEvents, "value[" + String(iEvt) + "].isCancelled"))
        Set Checked_State of oRespReq to ;
            (JsonValueAtPath(hoEvents, "value[" + String(iEvt) + "].responseRequested"))
        Set Value of oShowAs to ;
            (JsonValueAtPath(hoEvents, "value[" + String(iEvt) + "].showAs"))
        Set Label of oLink to ('<a href="' + ;
               JsonValueAtPath(hoEvents, "value[" + String(iEvt) + "].webLink") + '">View Event in browser</a>')
        
        Move (JsonValueAtPath(hoEvents, "value[" + String(iEvt) + "].body.contentType")) to sType
        If (sType = "text") Begin
            Set Value of oBody to ;
                (JsonValueAtPath(hoEvents, "value[" + String(iEvt) + "].body.content"))
            Set Visible_State of oBody to True
        End
        Else If (sType = "html") Begin
            Send NavigateToString of oHtmlBody ;
                (JsonValueAtPath(hoEvents, "value[" + String(iEvt) + "].body.content"))
            Set Visible_State of oHtmlBody to True
        End
        
        Send LoadAttendees of oAttendees ;
            (JsonObjectAtPath(hoEvents, "value[" + String(iEvt) + "].attendees"))
        Send LoadAttachments of oAttachment iEvt
    End_Procedure

    Object oEvent is a Form
        Set Size to 12 267
        Set Location to 27 574
        Set Label to "Body Preview:"
        Set Label_Col_Offset to 5
        Set Label_Justification_Mode to JMode_Right
        Set peAnchors to anTopRight
    End_Object
    
    Object oLocation is a Form
        Set Size to 12 131
        Set Location to 42 574
        Set Label to "Location:"
        Set Label_Col_Offset to 5
        Set Label_Justification_Mode to JMode_Right
        Set peAnchors to anTopRight
    End_Object
    
    Object oReminder is a Form
        Set Size to 12 24
        Set Location to 42 816
        Set Label to "Reminder before (minutes):"
        Set Label_Col_Offset to 5
        Set Label_Justification_Mode to JMode_Right
        Set peAnchors to anTopRight
    End_Object

    Object oAllday is a CheckBox
        Set Size to 10 50
        Set Location to 59 575
        Set Label to "All day"
        Set peAnchors to anTopRight
    End_Object

    Object oCancelled is a CheckBox
        Set Size to 10 50
        Set Location to 59 646
        Set Label to "Cancelled"
        Set peAnchors to anTopRight
    End_Object

    Object oRespReq is a CheckBox
        Set Size to 10 50
        Set Location to 59 736
        Set Label to "Response requested"
        Set peAnchors to anTopRight
    End_Object
    
    Object oShowAs is a Form
        Set Size to 12 61
        Set Location to 73 574
        Set Label to "Show time as:"
        Set Label_Col_Offset to 5
        Set Label_Justification_Mode to JMode_Right
        Set peAnchors to anTopRight
    End_Object

    Object oLink is a cLinkLabel
        Set Size to 8 81
        Set Location to 74 713
        Set Label to ""
        Set peAnchors to anTopRight
    End_Object

    Object oBody is a cTextEdit
        Set Size to 68 302
        Set Location to 98 535
        Set Label to "Body:"
        Set peAnchors to anTopRight
    End_Object

    Object oHtmlBody is a cWebView2Browser
        Set Size to 68 302
        Set Location to 98 535
        Set Label to "Body:"
        Set Label_Justification_Mode to JMode_Top
        Set Label_Col_Offset to 0
        Set peAnchors to anTopRight
        Set Visible_State to False
    End_Object

    Object oAttendees is a cCJGrid
        Set Size to 34 302
        Set Location to 168 535
        Set pbReadOnly to True
        Set peAnchors to anTopBottomRight

        Object oAttName is a cCJGridColumn
            Set piWidth to 305
            Set psCaption to "Attendee"
        End_Object
        
        Object oAttEMail is a cCJGridColumn
            Set piWidth to 573
            Set psCaption to "EMail"
        End_Object
        
        Procedure LoadAttendees Handle hoAtts
            tDataSourceRow[] atRows
            Integer i iMax iNameCol iEMailCol
            
            Get piColumnId of oAttName  to iNameCol
            Get piColumnId of oAttEMail to iEMailCol
            Move (JsonCountAtPath(hoAtts, "") - 1) to iMax
            
            For i from 0 to iMax
                Move (JsonValueAtPath(hoAtts, "[" + String(i) + "].emailAddress.name")) ;
                                                        to atRows[i].sValue[iNameCol]
                Move (JsonValueAtPath(hoAtts, "[" + String(i) + "].emailAddress.address")) ;
                                                        to atRows[i].sValue[iEMailCol]
            Loop
            
            Send InitializeData atRows
            Send MoveToFirstRow            
        End_Procedure

    End_Object

    Object oAttachments is a cCJGrid
        Set Size to 75 176
        Set Location to 206 534
        Set peAnchors to anBottomRight
        Set pbReadOnly to True

        Object oCJGridColumnRowIndicator2 is a cCJGridColumnRowIndicator
        End_Object
        
        Object oAttachment is a cCJGridColumn
            Set piWidth to 338
            Set psCaption to "Attachment"
        End_Object
        
        Object oAttSize is a cCJGridColumn
            Set piWidth to 173
            Set psCaption to "Size"
        End_Object

        Procedure LoadAttachments Integer iEvt
            tDataSourceRow[] atRows
            Handle  hoAtts hoEvents
            Integer i iMax iNameCol iSizeCol
            stGraphAttachments atAtts
            
            Get piColumnId of oAttachment to iNameCol
            Get piColumnId of oAttSize    to iSizeCol
            
            Get phoEvents to hoEvents
            
            If (JsonValueAtPath(hoEvents, "value[" + String(iEvt) + "].hasAttachments")) Begin
                Get ListEventAttachments of oGraph ;
                    (JsonValueAtPath(hoEvents, "value[" + String(iEvt) + "].id")) "" to hoAtts
                
                If hoAtts Begin
                    Get JsonToDataType of hoAtts to atAtts
                    Send Destroy of hoAtts
                End
                
                Move (SizeOfArray(atAtts.value) - 1) to iMax
                    
                For i from 0 to iMax
                    Move atAtts.value[i].name to atRows[i].sValue[iNameCol]
                    Move atAtts.value[i].size to atRows[i].sValue[iSizeCol]
                Loop
                
            End
            
            Set patAttachments to atAtts
            
            Send InitializeData atRows
            Send MoveToFirstRow            
        End_Procedure
        
        Procedure OnRowChanged Integer iOldRow Integer iNewRow
            stGraphAttachments atAtts
            String  sText sB64 sTemp
            Integer iLen iVoid iChn
            Pointer aBin
            UChar[] ucaBMP
            
            Set Visible_State of oTextAtt to False
            Set Visible_State of oBMPAtt  to False
            
            Get patAttachments to atAtts
            
            If (Left(atAtts.value[iNewRow].contentType, 4) = "text") Begin
                Move atAtts.value[iNewRow].contentBytes             to sB64
                Move (Base64DecodeFromStr(oCharTr, sB64, &iLen))    to aBin
                Move (PointerToString(aBin))                        to sText
                Move (Free(aBin))                                   to iVoid
                Set Value of oTextAtt                               to sText
                Set Visible_State of oTextAtt                       to True
            End
            
            If (atAtts.value[iNewRow].contentType = "image/bmp") Begin
                Move (StringToUCharArray(atAtts.value[iNewRow].contentBytes))   to ucaBmp
                Move (Base64DecodeUCharArray(oCharTr, ucaBMP))                  to ucaBmp
                Move (psBitmapPath(phoWorkspace(ghoApplication)) + "\" + atAtts.value[iNewRow].name) to sTemp
                Get Seq_New_Channel to iChn
                Direct_Output channel iChn ("binary:" * sTemp)
                Write channel iChn ucaBMP
                Close_Output channel iChn
                Send Seq_Release_Channel iChn
                Set Bitmap of oBMPAtt to sTemp
                Set Visible_State of oBMPAtt to True
                EraseFile sTemp
            End
            
        End_Procedure
        
        // WARNING: This is a nasty hack.
        //
        // This is the last COM object in the view, so once it is created it
        // becomes safe to do stuff with all the other COM objects, so this is
        // where we trigger stuff which relies on them.  Yuk!  :-(
        Procedure OnCreate
            Forward Send OnCreate
            
            Send LoadEvents of oEventsGrid
        End_Procedure
        
    End_Object

    Object oTextAtt is a cTextEdit
        Set Size to 74 122
        Set Location to 208 716
        Set peAnchors to anBottomRight
        Set Visible_State to False
        Set Read_Only_State to True
    End_Object

    Object oBMPAtt is a BitmapContainer
        Set Size to 74 109
        Set Location to 208 729
        Set peAnchors to anBottomRight
        Set Visible_State to False
        Set Bitmap_Style to Bitmap_Stretch
    End_Object
    
CD_End_Object
