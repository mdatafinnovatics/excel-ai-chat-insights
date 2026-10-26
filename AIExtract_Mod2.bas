Attribute VB_Name = "AIExtract_Mod2"

Option Explicit
' ============================================================
' MODULE: modPrompt   (VERSION 2 - improved time/trailer/notes logic)
' PURPOSE: Construct the full prompt sent to Claude
' Replaces the previous modPrompt module entirely.
' ============================================================

Public Function BuildPrompt(ByVal chatText As String, ByVal driverName As String, _
                            ByVal mode As String) As String
    Dim glossary As String, examples As String, todayStr As String
    Dim p As String
    
    glossary = BuildGlossaryText()
    examples = BuildExamplesText()
    todayStr = Format(Date, "mm/dd/yyyy")
    
    p = "You are an expert dispatch data extraction assistant for Canadian trucking company. " & _
        "Read the iMessage conversation between dispatcher and driver " & driverName & " below. " & _
        "Extract every individual driver EVENT into a structured JSON row." & vbLf & vbLf
    
    p = p & "# CONTEXT" & vbLf
    p = p & "- Driver: " & driverName & vbLf
    p = p & "- Today: " & todayStr & vbLf & vbLf
    
    p = p & "# OUTPUT FORMAT" & vbLf
    p = p & "Return ONLY a valid JSON array. No commentary. Each element:" & vbLf
    p = p & "{""time"":""HH:MM"",""time_inferred"":bool,""trailer"":""..."",""address"":""..."",""city"":""UPPERCASE"",""leave_or_arrive"":""leave|arrived"",""notes"":""..."",""rtn_new_picked"":""new|rtn|picked up|SCRAP|"""",""del_time"":""HH:MM""}" & vbLf & vbLf
    
    p = p & "# ABSOLUTE RULES (NEVER VIOLATE)" & vbLf
    
    p = p & "RULE 1: leave_or_arrive must be 'leave', 'arrived', or 'Finish'. NEVER output 'Start'. Even if chat says 'Start 8am', do NOT create a Start row." & vbLf
    p = p & "RULE 1B: The VERY LAST row of output MUST be a Finish row. Apply this logic:" & vbLf
    p = p & "  - If the chat explicitly says 'Finish X:XX' or 'Done X:XX' or similar: use that time as the Finish time." & vbLf
    p = p & "  - If chat does NOT explicitly say a finish time: estimate it based on the LAST event before it:" & vbLf
    p = p & "      * If the last event's address contains 'yard' (any form: 'yard', 'mud st', 'mud yard', 'home yard'): Finish time = same time as the last event." & vbLf
    p = p & "      * If the last event's address is anywhere ELSE (a customer, a stop): Finish time = last event time + 15 minutes." & vbLf
    p = p & "  - For the Finish row, set: trailer='', address='', city='', notes='Finish', rtn_new_picked='', del_time=''. Set time_inferred=false if time came from chat explicitly, true if estimated." & vbLf
    p = p & "RULE 2: NEVER use the current clock time anywhere. Times must come from the chat content only (ETA messages, timestamps, or inferred from neighbouring events)." & vbLf
    p = p & "RULE 3: EVERY 'time' field must be filled with HH:MM (24-hour). NEVER blank. Use the inference logic in Section TIME LOGIC." & vbLf
    p = p & "RULE 4: If dispatcher says 'Take to X' then later 'Change take to Y' BEFORE driver arrives at X _ do NOT create an arrived row for X. Driver only arrives at Y." & vbLf
    p = p & "RULE 5: BOBTAIL tracking _ maintain mental state of 'current trailer'. When driver drops without picking up new _ current trailer becomes BOBTAIL. The NEXT leave row MUST show BOBTAIL, not the old trailer number." & vbLf & vbLf
    
    p = p & "# TIME LOGIC (CRITICAL - follow exactly)" & vbLf
    p = p & "STEP 1: Read entire chat top-to-bottom. Note every explicit time mentioned: 'ETA X:XX', '@ X:XX', 'at X:XX', timestamps like '12:00 PM'." & vbLf
    p = p & "STEP 2: For each driver event, determine its time using this priority order:" & vbLf
    p = p & "  (a) ARRIVED event time: Use the most recent ETA message mentioned BEFORE the 'Arrived' / 'At customer' / 'At yard' message. If driver says 'ETA 1:15pm' then 'At customer' _ arrived time = 13:15." & vbLf
    p = p & "  (b) LEAVE event time (after an arrival): If chat doesn't state, set = arrived_time + 15 minutes (typical unloading/coupling)." & vbLf
    p = p & "  (c) LEAVE event time (bobtail after delivery): If chat doesn't state, set = arrived_time + 30 minutes (typical drop and depart time)." & vbLf
    p = p & "  (d) LEAVE event time (after switch/PU): If chat doesn't state, set = arrived_time + 15 minutes." & vbLf
    p = p & "  (e) FIRST event time: Must come from chat _ first explicit time mentioned. NEVER current clock time. If chat starts with driver already at customer at 12:00, the first 'arrived' = 12:00." & vbLf
    p = p & "STEP 3: When you infer a time using (b), (c), or (d) above, set time_inferred=true. When time comes directly from chat (a), (e), set time_inferred=false." & vbLf & vbLf
    
    p = p & "# TRAILER LOGIC" & vbLf
    p = p & "- Track 'current trailer' as you read. Start = unknown (read context to find first trailer)." & vbLf
    p = p & "- When driver says 'Dropped trl' or 'Drop and PU XXXX' or 'Delivered/empty trailer' WITHOUT immediate pickup ? next leave is BOBTAIL." & vbLf
    p = p & "- When driver says 'Picked up Trl XXXXX' or 'PU XXXXX' ? current trailer becomes that number." & vbLf
    p = p & "- BOBTAIL must appear on BOTH the leave row AND the matching arrive row of an empty-truck trip." & vbLf
    p = p & "- When chat says 'bobtail to <place>' _ driver leaves with BOBTAIL, arrives with BOBTAIL." & vbLf & vbLf
    
    p = p & "# NOTES LOGIC (rich notes like a human dispatcher would write)" & vbLf
    p = p & "Notes column should include relevant context per event:" & vbLf
    p = p & "- For 'arrived' at customer: status info (e.g. 'At customer - no one here', 'Live load', 'NFI SDR live load then drop', 'Dropped trl 18 dock')." & vbLf
    p = p & "- For 'leave' events: reason if special (e.g. 'Trailer rejected - rust, roof, rims, odour', 'No loaded available - bobtail to mud', 'Drop and PU 859744 (switch)')." & vbLf
    p = p & "- Include dispatcher instructions when they explain WHY the event happened (e.g., 'Take to mud st after rejection')." & vbLf
    p = p & "- Include customer codes mentioned in chat (FGF, NFI SDR, Boston Blue S1, etc.)." & vbLf
    p = p & "- Be DESCRIPTIVE, not minimal. Match the depth of a dispatcher writing in a logbook." & vbLf & vbLf
    
    p = p & "# GLOSSARY" & vbLf & glossary & vbLf & vbLf
    
    p = p & "# WORKED EXAMPLES" & vbLf & examples & vbLf & vbLf
    
    p = p & "# CHAT TO EXTRACT" & vbLf & chatText & vbLf & vbLf
    
    p = p & "Return ONLY the JSON array. Apply all rules strictly."
    
    BuildPrompt = p
End Function

Private Function BuildGlossaryText() As String
    Dim ws As Worksheet, LastRow As Long, i As Long, Result As String
    Set ws = ThisWorkbook.Sheets("Glossary")
    LastRow = ws.Cells(ws.rows.Count, "C").End(xlUp).Row
    Result = ""
    For i = 4 To LastRow
        If ws.Cells(i, 3).Value <> "" Then
            Result = Result & "- [" & ws.Cells(i, 2).Value & "] '" & _
                     ws.Cells(i, 3).Value & "' = " & ws.Cells(i, 4).Value & vbLf
        End If
    Next i
    BuildGlossaryText = Result
End Function

Private Function BuildExamplesText() As String
    ' This example shows the exact reasoning style needed for Chat 1 (Khan)
    ' Key teaching points: redirected destination, BOBTAIL transitions, rich notes, time inference
    Dim r As String
    r = "## EXAMPLE A - Driver redirected mid-route + trailer rejection + switch" & vbLf
    r = r & "INPUT CHAT:" & vbLf
    r = r & "  Driver: Delivered/empty trailer" & vbLf
    r = r & "  Jacob: Take to 140 Ormont Dr, North York" & vbLf
    r = r & "  Driver: ETA 12:00" & vbLf
    r = r & "  Driver: No one here" & vbLf
    r = r & "  Jacob: It's a switch. Drop and pu 859744" & vbLf
    r = r & "  Driver: Rust and roof and rims rust and odour in the trailer" & vbLf
    r = r & "  Jacob: Okay take the trailer to mud st" & vbLf
    r = r & "  Driver: ETA 2:00 pm" & vbLf
    r = r & "  Jacob: Change take to 5900 Avebury Rd Mississauga" & vbLf
    r = r & "  Driver: ETA 1:15pm" & vbLf
    r = r & "  Driver: At customer" & vbLf
    r = r & "  Jacob: NFI SDR, live load, then take to mud" & vbLf
    r = r & "  Driver: Dropped trl 18 dock" & vbLf & vbLf
    r = r & "REASONING:" & vbLf
    r = r & "- First event: driver arrives at 140 Ormont at 12:00 (ETA 12:00 immediately before 'No one here')." & vbLf
    r = r & "- Driver was told 'take to mud st' AND THEN redirected to Avebury BEFORE arriving at mud. So NO arrival row for mud." & vbLf
    r = r & "- Driver leaves 140 Ormont with the rejected trailer (current trailer unknown at this point, so just trailer = previously implied or empty)." & vbLf
    r = r & "  Wait _ driver said 'Delivered/empty trailer' at start, then Jacob said 'switch, drop and PU 859744'. So driver actually has 859744 after switch. The rejected one stays at customer." & vbLf
    r = r & "  After rejection: driver bobtails (rejected trailer left there). So leave from 140 Ormont = BOBTAIL." & vbLf
    r = r & "- Driver arrives at Avebury at 13:15 (ETA 1:15pm right before 'At customer')." & vbLf
    r = r & "- After dropping at dock 18, driver bobtails to mud. Leave time = 13:15 + 30min = 13:45 (inferred)." & vbLf & vbLf
    r = r & "OUTPUT:" & vbLf
    r = r & "[" & vbLf
    r = r & "  {""time"":""12:00"",""time_inferred"":false,""trailer"":""BOBTAIL"",""address"":""140 Ormont Dr"",""city"":""NORTH YORK"",""leave_or_arrive"":""arrived"",""notes"":""At customer - no one here, switch instructed (drop and PU 859744)"",""rtn_new_picked"":"""",""del_time"":""12:00""}," & vbLf
    r = r & "  {""time"":""12:30"",""time_inferred"":true,""trailer"":""BOBTAIL"",""address"":"""",""city"":"""",""leave_or_arrive"":""leave"",""notes"":""Trailer rejected - rust, roof, rims, odour. Redirected to Avebury (originally mud)"",""rtn_new_picked"":"""",""del_time"":""""}," & vbLf
    r = r & "  {""time"":""13:15"",""time_inferred"":false,""trailer"":""BOBTAIL"",""address"":""5900 Avebury Rd"",""city"":""MISSISSAUGA"",""leave_or_arrive"":""arrived"",""notes"":""At customer - NFI SDR live load then drop at dock 18"",""rtn_new_picked"":"""",""del_time"":""13:15""}," & vbLf
    r = r & "  {""time"":""13:45"",""time_inferred"":true,""trailer"":""BOBTAIL"",""address"":"""",""city"":"""",""leave_or_arrive"":""leave"",""notes"":""Dropped trl at dock 18 - bobtail to mud"",""rtn_new_picked"":"""",""del_time"":""""}," & vbLf
    r = r & "  {""time"":""14:00"",""time_inferred"":true,""trailer"":"""",""address"":"""",""city"":"""",""leave_or_arrive"":""Finish"",""notes"":""Finish (estimated - last event not at yard, +15min)"",""rtn_new_picked"":"""",""del_time"":""""}" & vbLf
    r = r & "]" & vbLf & vbLf
    
    r = r & "## EXAMPLE B - Simple bobtail trip with switch" & vbLf
    r = r & "INPUT CHAT:" & vbLf
    r = r & "  Driver: Leaving yard with Trl 1013 empty. ETA 8am" & vbLf
    r = r & "  Driver: Arrived at location" & vbLf
    r = r & "  Driver: Picked up Trl 350215 loaded for yard. ETA 11am" & vbLf
    r = r & "  Driver: Arrived at yard" & vbLf & vbLf
    r = r & "REASONING:" & vbLf
    r = r & "- Driver leaves yard with 1013 at 7:30 (inferred: ETA 8am, ~30min trip)." & vbLf
    r = r & "- Arrived at customer at 8:00 (ETA 8am)." & vbLf
    r = r & "- Picked up new trailer 350215. Leave time = 8:00 + 15min = 8:15 (inferred)." & vbLf
    r = r & "- Arrived at yard at 11:00 (ETA 11am)." & vbLf & vbLf
    r = r & "OUTPUT:" & vbLf
    r = r & "[" & vbLf
    r = r & "  {""time"":""07:30"",""time_inferred"":true,""trailer"":""1013"",""address"":"""",""city"":"""",""leave_or_arrive"":""leave"",""notes"":""Leaving yard with empty Trl 1013 - ETA 8am"",""rtn_new_picked"":"""",""del_time"":""""}," & vbLf
    r = r & "  {""time"":""08:00"",""time_inferred"":false,""trailer"":""1013"",""address"":""(customer)"",""city"":"""",""leave_or_arrive"":""arrived"",""notes"":""Trailer being loaded - will drop trl I brought (switch)"",""rtn_new_picked"":"""",""del_time"":""08:00""}," & vbLf
    r = r & "  {""time"":""08:15"",""time_inferred"":true,""trailer"":""350215"",""address"":"""",""city"":"""",""leave_or_arrive"":""leave"",""notes"":""Picked up Trl 350215 loaded for yard"",""rtn_new_picked"":""picked up"",""del_time"":""""}," & vbLf
    r = r & "  {""time"":""11:00"",""time_inferred"":false,""trailer"":""350215"",""address"":""yard"",""city"":"""",""leave_or_arrive"":""arrived"",""notes"":""Arrived at yard"",""rtn_new_picked"":"""",""del_time"":""""}," & vbLf
    r = r & "  {""time"":""11:00"",""time_inferred"":true,""trailer"":"""",""address"":"""",""city"":"""",""leave_or_arrive"":""Finish"",""notes"":""Finish (estimated - last event at yard, same time)"",""rtn_new_picked"":"""",""del_time"":""""}" & vbLf
    r = r & "]"
    
    BuildExamplesText = r
End Function








