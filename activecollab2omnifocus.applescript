-- Copyright 2014 Digital Marauders.  All rights reserved.
using terms from application "Mail"
  -- Trims "foo <foo@bar.com>" down to "foo@bar.com"
  on trim_address(theAddress)
    try
      set AppleScript's text item delimiters to "<"
      set WithoutPrefix to item 2 of theAddress's text items
      set AppleScript's text item delimiters to ">"
      set MyResult to item 1 of WithoutPrefix's text items
    on error
      set MyResult to theAddress
    end try
    set AppleScript's text item delimiters to {""} --> restore delimiters to default value
    return MyResult
  end trim_address
  
  on process_message(theMessage)
    tell application "OmniFocus"
      log "OmniFocus calling process_message in MailAction script"
    end tell
    set theSubject to subject of theMessage
    if (theSubject contains "Task '" and theSubject contains "' has been Created") then
      set singleTask to true
      set offset1 to offset of "Task '" in theSubject
      set offset1 to offset1 + 6
      set offset2 to offset of "' has been Created" in theSubject
      set offset2 to offset2 - 1
      set taskTitle to ""
      copy characters 1 through offset2 of theSubject as string to taskTitle
      set theText to taskTitle & return & content of theMessage
      tell application "OmniFocus"
        -- starting project lookup
        -- set projectList to the flattened projects of default document
        -- set theProjectCount to count of projectList
        -- repeat with theProjectIndex from 1 to theProjectCount
        --   log the name of item theProjectIndex of projectList as string
        -- end repeat

        parse tasks into default document with transport text theText as single task singleTask
      end tell
    end if
  end process_message
  
  on perform mail action with messages theMessages
    try
      set theMessageCount to count of theMessages
      repeat with theMessageIndex from 1 to theMessageCount
        my process_message(item theMessageIndex of theMessages)
      end repeat
    on error m number n
      tell application "OmniFocus"
        log "Exception in Mail action: (" & n & ") " & m
      end tell
    end try
  end perform mail action with messages
  
end using terms from