Use Windows.pkg
Use DFClient.pkg
Use cCJGrid.pkg
Use cCJGridColumn.pkg
Use cTextEdit.pkg
Use dfLine.pkg
Use cWebView2Browser.pkg
Use cMSGraphAPI.pkg
Use JsonConvFuncs.pkg
Use ComposeMail.dg
Use cCJGridColumnRowIndicator.pkg
Use seq_chnl.pkg

Struct stMsgItem
    String sID
    String sSubj
    String sFrom
    DateTime dtRcvd
    Boolean bAtts
End_Struct

Deferred_View Activate_oEMail for ;
Object oEMail is a dbView
    
    Property String psMailFolder
    Property String psMsgID
    
    Property String psInboxID
    Property String psDraftsID
    Property String psSentID
    Property String psJunkID
    Property String psDeltdID
    Property String psArchID
    Property stMsgItem[] patMsgs

    Set Border_Style to Border_Thick
    Set Size to 269 568
    Set Location to 2 2
    Set Label to "EMail"
    
    Procedure SetLoginState
        Boolean bLoggedIn
        
        Get IsLoggedIn to bLoggedIn
        Set Label of  oLogIn          to (If(bLoggedIn, "Log out of Graph", "Log in to Graph"))
        Set Label of  oLoginStatus    to (If(bLoggedIn, "Logged in to Graph", "Not logged in"))
        Set TextColor of oLoginStatus to (If(bLoggedIn, clBlue, clRed))
        Set Enabled_State of oComposeBtn to bLoggedIn
    End_Procedure
    
    Procedure Activating
        Forward Send Activating
        Send SetLoginState
    End_Procedure
    
    Procedure OnOAuth2LoggedIn
        Send SetLoginState
        Send GetFolders
        Send OnClick of oInboxBtn
    End_Procedure
    
    Procedure GetFolders
        Handle  hoResp hoFldrs hoFldr
        Integer iFldrs i
        String  sFldr sID
        
        Send SetLoginState
        
        // Set up Mail Folder IDs:
        Get ListMailFolders of (oGraph(Self)) "$select=displayName" to hoResp
        
        If not hoResp ;
            Procedure_Return
        
        If not (HasMember(hoResp, "value")) Begin
            Send Destroy of hoResp
            Procedure_Return
        End
        
        Get Member of hoResp "value" to hoFldrs
        Get MemberCount of hoFldrs to iFldrs
        Decrement iFldrs
        
        For i from 0 to iFldrs
            Get MemberByIndex of hoFldrs i to hoFldr
            Get MemberValue of hoFldr "displayName" to sFldr
            Get MemberValue of hoFldr "id" to sID
            
            Case Begin
                
                Case (sFldr = "Archive")
                    Set psArchID to sID
                    Case Break
                
                Case (sFldr = "Drafts")
                    Set psDraftsID to sID
                    Case Break
                
                Case (sFldr = "Inbox")
                    Set psInboxID to sID
                    Case Break
                
                Case (sFldr = "Junk Email")
                    Set psJunkID to sID
                    Case Break
                
                Case (sFldr = "Sent Items")
                    Set psSentID to sID
                    Case Break
                
                Case (sFldr = "Deleted Items")
                    Set psDeltdID to sID
                    Case Break
                
            Case End
            
            Send Destroy of hoFldr
        Loop
        
        Send Destroy of hoFldrs
        Send Destroy of hoResp
    End_Procedure

    Procedure OnOAuth2NotLoggedIn
        Send SetLoginState
    End_Procedure
    
    Object oLogIn is a Button
        Set Size to 14 76
        Set Location to 3 4
        Set Label to "Log in to Graph"
    
        Procedure OnClick
            Boolean bLoggedIn
            
            Get IsLoggedIn to bLoggedIn
            
            If bLoggedIn ;
                Send Logout of oOAuth2
            Else ;
                Send Login of oOauth2
            
            Send SetLoginState
        End_Procedure
    
    End_Object

    Object oLoginStatus is a TextBox
        Set Size to 9 49
        Set Location to 5 89
        Set Label to "Not logged in"
        Set TextColor to clRed
    End_Object

    Object oComposeBtn is a Button
        Set Size to 14 64
        Set Location to 3 394
        Set Label to "Compose New"
        Set peAnchors to anTopRight
        Set Enabled_State to False
    
        Procedure OnClick
            Send Popup of oComposeMail
        End_Procedure
    
    End_Object

    Object oRefreshBtn is a Button
        Set Location to 3 513
        Set Label to "Refresh"
        Set peAnchors to anTopRight
    
        Procedure OnClick
            Send LoadMessages
        End_Procedure
    
    End_Object
    
    Procedure LoadMessages
        String      sFldr sID sName sAddr sRcvd sParams
        Handle      hoResp hoMsgs hoMsg hoAddrs
        Integer     iMsgs i
        stMsgItem[] atMsgs
        DateTime    dtRcvd
        Boolean     bShowTo
        
        Get psMailFolder to sFldr
        
        If (sFldr = "Inbox") ;
            Get psInboxID to sID
        If (sFldr = "Drafts") ;
            Get psDraftsID to sID
        If (sFldr = "Sent Items") ;
            Get psSentID to sID
        If (sFldr = "Junk") ;
            Get psJunkID to sID
        If (sFldr = "Deleted") ;
            Get psDeltdID to sID
        If (sFldr = "Archived") ;
            Get psArchID to sID
        
        Move ((sFldr  = "Drafts") or (sFldr = "Sent Items")) to bShowTo
        
        Move "$orderby=receivedDateTime desc&$top=10000&$select=subject,receivedDateTime,hasAttachments," to sParams
        Move (sParams + If(bShowTo, "toRecipients", "from")) to sParams
            
        Get ListMessages of oGraph sID sParams to hoResp
        
        If not hoResp ;
            Procedure_Return
        
        If not (HasMember(hoResp, "value")) ;
            Procedure_Return
        
        Get Member of hoResp "value" to hoMsgs
        Get MemberCount of hoMsgs to iMsgs
        Decrement iMsgs
        
        For i from 0 to iMsgs
            Get MemberByIndex of hoMsgs i               to hoMsg
            
            Get MemberValue of hoMsg "id"               to atMsgs[i].sID

            If (HasMember(hoMsg, "subject") and ;
               (MemberJsonType(hoMsg, "subject") <> jsonTypeNull)) ;
                Get MemberValue of hoMsg "subject"      to atMsgs[i].sSubj
            
            If bShowTo Begin

                If (HasMember(hoMsg, "toRecipients")) Begin
                    Get Member of hoMsg "toRecipients" to hoAddrs
                    Get JEmAddrs2Str hoAddrs to atMsgs[i].sFrom
                    Send Destroy of hoAddrs
                End

            End
            Else Begin

                If (HasMember(hoMsg, "from")) Begin
                    Get Member of hoMsg "from" to hoAddrs
                    Get JEmAddr2Str hoAddrs to atMsgs[i].sFrom
                    Send Destroy of hoAddrs
                End
                
            End
            
            If (HasMember(hoMsg, "receivedDateTime")) ;
                Move (IsoDt2DfDt(MemberValue(hoMsg, "receivedDateTime"))) to atMsgs[i].dtRcvd
            
            If (HasMember(hoMsg, "hasAttachments")) ;
                Move (MemberValue(hoMsg, "hasAttachments")) to atMsgs[i].bAtts
            
            Send Destroy of hoMsg
        Loop
        
        Send Destroy of hoMsgs
        Send Destroy of hoResp
        
        Set patMsgs to atMsgs
        Send LoadData of oMailGrid bShowTo
    End_Procedure

    Procedure LoadMessage Integer iRow
        Handle hoDS hoResp hoVal
        String sID sParams
        tDataSourceRow[] atRows
        
        // Blank the message fields
        Set Value of oFrom      to ""
        Set Value of oTo        to ""
        Set Value of oCC        to ""
        Set Value of oSubject   to ""
        Set Value of oMsgTxt    to ""
        Set Value of oRcvdDT    to ""
        Send ResetGrid of oAttachments
        Set psMsgID             to ""
        
        // Get the message ID from the grid datasource
        Get phoDataSource of oMailGrid to hoDS
        Get DataSource of hoDS to atRows
        If (iRow >= SizeOfArray(atRows)) ;
            Procedure_Return
        Move atRows[iRow].vTag to sID
        
        Move "$select=subject,from,toRecipients,ccRecipients,body,receivedDateTime,hasAttachments" to sParams
        
        Get GetMessage of (oGraph(Self)) sID SParams to hoResp
        
        If not hoResp ;
            Procedure_Return
        
        // See Dennis? THIS is the issue!
        String sJson
        Get Stringify of hoResp to sJson
        
        If (HasMember(hoResp, "from")) Begin
            Get Member of hoResp "from"  to hoVal
            Set Value of oFrom to (JEmAddr2Str(hoVal))
            Send Destroy of hoVal
        End

        If (HasMember(hoResp, "toRecipients")) Begin
            Get Member of hoResp "toRecipients" to hoVal
            Set Value of oTo to (JEmAddrs2Str(hoVal))
            Send Destroy of hoVal
        End
                
        If (HasMember(hoResp, "ccRecipients")) Begin
            Get Member of hoResp "ccRecipients" to hoVal
            Set Value of oCC to (JEmAddrs2Str(hoVal))
            Send Destroy of hoVal
        End
        
        If (HasMember(hoResp, "subject")) ;
            Set Value of oSubject to (MemberValue(hoResp, "subject"))
        
        If (HasMember(hoResp, "receivedDateTime")) ;
            Set Value of oRcvdDT to (IsoDt2DfDt(MemberValue(hoResp, "receivedDateTime")))
        
        If (HasMember(hoResp, "body")) Begin
            Get Member of hoResp "body" to hoVal
            
            If not (HasMember(hoVal, "content")) ;
                Break
            
            If (MemberValue(hoVal, "contentType") = "html") Begin
                Send NavigateToString of oMsgHtml (MemberValue(hoVal, "content"))
                Set Visible_State of oMsgHtml to True
                Set Visible_State of oMsgTxt  to False
            End
            Else Begin
                Set Value of oMsgTxt to (MemberValue(hoVal, "content"))
                Set Visible_State of oMsgHtml to False
                Set Visible_State of oMsgTxt  to True
            End
            
            Send Destroy of hoVal
        End
        
        If (HasMember(hoResp, "hasAttachments") and ;
            (MemberValue(hoResp, "hasAttachments"))) ;
            Send LoadAttachments of oAttachments sID
        
        Send Destroy of hoResp
        
        Set psMsgID to sID        
    End_Procedure
    
    Procedure DisplayAttachment String sAttID
        String  sMsgID sBytes sName sDwnld
        Handle  hoResp
        UChar[] ucaBytes
        Integer iChn
        Boolean bExist
        
        Get psMsgID to sMsgID
        
        If ((sMsgID = "") or (sAttID = "")) ;
            Procedure_Return
        
        Get Attachment of (oGraph(Self)) sMsgID sAttID "" to hoResp
        
        If not hoResp ;
            Procedure_Return
                
        If not (HasMember(hoResp, "name")) ;
            Procedure_Return
        If not (HasMember(hoResp, "contentBytes")) ;
            Procedure_Return
        
        Get MemberValue of hoResp "name"            to sName
        Get MemberValue of hoResp "contentBytes"    to sBytes
        
        // NOTE: This is nasty and will leave the file in your Downloads folder
        //       for later cleanup.  We can't just delete it now because the
        //       program that opens it needs it to be there.
        //
        //       Adding (see below):
        //           Sleep 1
        //           EraseFile sDwnld
        //       to the end of the procedure does work sometimes, but is a
        //       bit "brittle", so I have commented it out.
        Move (StringToUCharArray(sBytes))                   to ucaBytes
        Move (Base64DecodeUCharArray(oCharTr, ucaBytes))    to ucaBytes
        Move (psHome(phoWorkspace(ghoApplication)) + "\Downloads") to sDwnld
        
        // Make sure we have the folder "Downloads"
        File_Exist sDwnld bExist
        
        If not bExist ;
            Make_Directory sDwnld
        
        // Write the file:
        Move (sDwnld + "\" + sName) to sDwnld
        Get Seq_New_Channel to iChn
        Direct_Output channel iChn ("binary:" * sDwnld)
        Write channel iChn ucaBytes
        Close_Output channel iChn
        Send Seq_Release_Channel iChn
        
        // Get the OS to open it as appropriate
        Runprogram Shell Background sDwnld
        
        // Use at your own risk
//        Sleep 1
//        EraseFile sDwnld
    End_Procedure
    
    Object oBoxGrp is a Group
        Set Size to 120 56
        Set Location to 19 4
        Set Label to "Folders:"
        
        Procedure ResetBtns Handle hoCurr
            
            Set Form_FontWeight of oInboxBtn    to 400
            Set Form_FontWeight of oDraftsBtn   to 400
            Set Form_FontWeight of oSentBtn     to 400
            Set Form_FontWeight of oJunkBtn     to 400
            Set Form_FontWeight of oDeletedBtn  to 400
            Set Form_FontWeight of oArchBtn     to 400
            
            Set Form_FontWeight of hoCurr       to 800
            
            Set psMailFolder to (Label(hoCurr))
            Send LoadMessages
        End_Procedure

        Object oInboxBtn is a Button
            Set Location to 12 2
            Set Label to "Inbox"
        
            Procedure OnClick
                Send ResetBtns Self
            End_Procedure
        
        End_Object

        Object oDraftsBtn is a Button
            Set Location to 30 2
            Set Label to "Drafts"
        
            Procedure OnClick
                Send ResetBtns Self
                
            End_Procedure
        
        End_Object

        Object oSentBtn is a Button
            Set Location to 49 2
            Set Label to "Sent Items"
        
            Procedure OnClick
                Send ResetBtns Self
            End_Procedure
        
        End_Object

        Object oJunkBtn is a Button
            Set Location to 67 2
            Set Label to "Junk"
        
            Procedure OnClick
                Send ResetBtns Self
            End_Procedure
        
        End_Object

        Object oDeletedBtn is a Button
            Set Location to 86 2
            Set Label to "Deleted"
        
            Procedure OnClick
                Send ResetBtns Self
            End_Procedure
        
        End_Object

        Object oArchBtn is a Button
            Set Location to 104 2
            Set Label to "Archived"
        
            Procedure OnClick
                Send ResetBtns Self
                
            End_Procedure
        
        End_Object
        
    End_Object

    Object oMailGrid is a cCJGrid
        Set Size to 126 498
        Set Location to 21 68
        Set pbReadOnly to True
        Set peAnchors to anTopLeftRight

        Object oCJGridColumnRowIndicator1 is a cCJGridColumnRowIndicator
        End_Object

        Object oSubjCol is a cCJGridColumn
            Set piWidth to 110
            Set psCaption to "Subject"
        End_Object

        Object oFromCol is a cCJGridColumn
            Set piWidth to 200
            Set psCaption to "From"
        End_Object

        Object oRcvdCol is a cCJGridColumn
            Set piWidth to 40
            Set psCaption to "Received"
        End_Object

        Object oHasAtts is a cCJGridColumn
            Set piWidth to 6
            Set psCaption to (Character(128206))
            Set peTextAlignment to xtpAlignmentCenter
            Set peHeaderAlignment to xtpAlignmentCenter
            
            Procedure OnSetDisplayMetrics Handle hoGridItemMetrics Integer iRow String ByRef sValue
                
                If (sValue = "1") ;
                    Move (Character(128206)) to sValue
                Else ;
                    Move "" to sValue
            End_Procedure
            
        End_Object
        
        Procedure LoadData Boolean bShowTo
            stMsgItem[] atMsgs
            tDataSourceRow[] atRows
            Integer iSubjCol iFromCol iRcvdCol iAttsCol i iMax
            
            Set psCaption of oFromCol to (If(bShowTo, "To", "From"))
            
            Get piColumnId of oSubjCol to iSubjCol
            Get piColumnId of oFromCol to iFromCol
            Get piColumnId of oRcvdCol to iRcvdCol
            Get piColumnId of oHasAtts to iAttsCol
            
            Get patMsgs to atMsgs
            Move (SizeOfArray(atMsgs) - 1) to iMax
            
            For i from 0 to iMax
                Move atMsgs[i].sID    to atRows[i].vTag
                Move atMsgs[i].sSubj  to atRows[i].sValue[iSubjCol]
                Move atMsgs[i].sFrom  to atRows[i].sValue[iFromCol]
                Move atMsgs[i].dtRcvd to atRows[i].sValue[iRcvdCol]
                Move atMsgs[i].bAtts  to atRows[i].sValue[iAttsCol]
            Loop
            
            Send InitializeData atRows
            Send MovetoFirstRow
        End_Procedure
        
        Procedure OnRowChanged Integer iOldRow Integer iNewRow
            Send LoadMessage iNewRow
        End_Procedure
        
    End_Object
    
    Object oFrom is a Form
        Set Size to 12 352
        Set Location to 151 37
        Set Label to "From:"
        Set Label_Justification_Mode to JMode_Right
        Set Label_Col_Offset to 5
        Set peAnchors to anTopLeftRight
    End_Object

    Object oTo is a Form
        Set Size to 12 352
        Set Location to 165 37
        Set Label to "To:"
        Set Label_Justification_Mode to JMode_Right
        Set Label_Col_Offset to 5
        Set peAnchors to anTopLeftRight
    End_Object

    Object oCC is a Form
        Set Size to 12 352
        Set Location to 179 37
        Set Label to "CC:"
        Set Label_Justification_Mode to JMode_Right
        Set Label_Col_Offset to 5
        Set peAnchors to anTopLeftRight
    End_Object
    
    Object oSubject is a Form
        Set Size to 12 352
        Set Location to 193 37
        Set Label to "Subject:"
        Set Label_Justification_Mode to JMode_Right
        Set Label_Col_Offset to 5
        Set peAnchors to anTopLeftRight
    End_Object

    Object oAttachments is a cCJGrid
        Set Size to 39 170
        Set Location to 151 395
        Set peAnchors to anTopRight
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

        Procedure LoadAttachments String sMsgID
            tDataSourceRow[] atRows
            Handle  hoResp hoAtts hoAtt
            Integer i iMax iNameCol iSizeCol
            String  sParams
            
            Move "$select=name,size" to sParams
            Get ListMailAttachments of (oGraph(Self)) sMsgID sParams to hoResp
            
            If not hoResp ;
                Procedure_Return
            
            Get piColumnId of oAttachment to iNameCol
            Get piColumnId of oAttSize    to iSizeCol
            
            If not (HasMember(hoResp, "value"))  ;
                Procedure_Return
            
            Get Member of hoResp "value" to hoAtts
            Get MemberCount of hoAtts to iMax
            Decrement iMax
            
            For i from 0 to iMax
                Get MemberByIndex of hoAtts i to hoAtt
                If (HasMember(hoAtt, "id")) ;
                    Get MemberValue of hoAtt "id" to atRows[i].vTag
                If (HasMember(hoAtt, "name")) ;
                    Get MemberValue of hoAtt "name" to atRows[i].sValue[iNameCol]
                If (HasMember(hoAtt, "size")) ;
                    Get MemberValue of hoAtt "size" to atRows[i].sValue[iSizeCol]
                Send Destroy of hoAtt
            Loop
            
            Send Destroy of hoAtts
            Send Destroy of hoResp
            
            Send InitializeData atRows
            Send MoveToFirstRow
        End_Procedure
        
        Procedure OnRowDoubleClick Integer iRow Integer iCol
            Handle  hoDS hoResp
            tDataSourceRow[] atRows
            String  sID
            
            Get phoDataSource to hoDS
            Get DataSource of hoDS to atRows
            Move atRows[iRow].vTag to sID
            
            Send DisplayAttachment sID
        End_Procedure
        
    End_Object
        
    Object oBodyTextBox is a TextBox
        Set Size to 10 50
        Set Location to 207 3
        Set Label to "Message body:"
    End_Object

    Object oMsgTxt is a cTextEdit
        Set Size to 49 562
        Set Location to 217 3
        Set peAnchors to anAll
    End_Object

    Object oMsgHtml is a cWebView2Browser
        Set Size to 49 562
        Set Location to 217 3
        Set Label_Justification_Mode to JMode_Top
        Set Label_Col_Offset to 0
        Set peAnchors to anAll
        Set Enabled_State to True
        Set Visible_State to False
        Set pbBorderVisible to True
        
        // WARNING: This is a nasty hack.
        //
        // This is the last COM object in the view, so once it is created it
        // becomes safe to do stuff with all the other COM objects, so this is
        // where we trigger stuff which relies on them.  Yuk!  :-(
        Procedure OnCreate
            Forward Send OnCreate
            
            If (IsLoggedIn(Self)) Begin
                Send GetFolders
                Send OnClick of oInboxBtn
            End
            
        End_Procedure
        
    End_Object

    Object oRcvdDT is a Form
        Set Size to 12 69
        Set Location to 193 496
        Set Label to "Received:"
        Set Label_Justification_Mode to JMode_Right
        Set Label_Col_Offset to 5
        Set peAnchors to anTopRight
    End_Object

Cd_End_Object
