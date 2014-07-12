--
--  AppDelegate.applescript
--  Hours Log
--
--  Created by Monty on 8/20/13.
--  Copyright (c) 2013 Monty. All rights reserved.
--

script AppDelegate
	property parent : class "NSObject"
    -- this property will be used to configure our notification
	property myNotification : missing value
    
    property theAppName : "Hours Log"
    property theSetMessage : "Time tracking spreadsheet opened."
    property theClearMessage : "Time tracking spreadsheet closed."
    property theResult : "false"    
	
	on applicationWillFinishLaunching_(aNotification)
        
        -- Launch a Excel and open the spreadsheet
        set theHours to "SSD:Users:monty:Dropbox:Echostar:Managerial:Hours Archive:hours_2013.xlsx"
        tell application "Microsoft Excel"
            activate
            set myFile to full name of active workbook
            if myFile is equal to theHours then
                close workbook (the name of active workbook)
                else
                open theHours
            end if
        end tell
        
        quit
        
	end applicationWillFinishLaunching_
	
	on applicationShouldTerminate_(sender)
		-- Insert code here to do any housekeeping before your application quits
		return current application's NSTerminateNow
	end applicationShouldTerminate_
        
	-- Method for sending a notification based on supplied title and text
	on sendNotificationWithTitle_AndMessage_(aTitle, aMessage)
		set myNotification to current application's NSUserNotification's alloc()'s init()
		set myNotification's title to aTitle
		set myNotification's informativeText to aMessage
		current application's NSUserNotificationCenter's defaultUserNotificationCenter's deliverNotification_(myNotification)
	end sendNotification
    
end script