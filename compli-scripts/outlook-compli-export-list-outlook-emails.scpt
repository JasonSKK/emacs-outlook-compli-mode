-- "How to print a list of selected emails in Outlook" original script by Moth Software, retrieved on May 1st, 2023, from the Moth Software Blog
-- Author: Beatrix Willius
-- https://www.mothsoftware.com/how-to-print-a-list-of-selected-emails-in-outlook

on run argv
	set theFolder to POSIX file (item 1 of argv) as alias
	tell application "Microsoft Outlook"
		activate         
		tell application "System Events"
			tell process "Microsoft Outlook"
				delay 1
				keystroke "a" using {command down}
			end tell
		end tell
	end tell

	tell application "Microsoft Outlook"
		set this_data to ""
		set SelectedMails to selection
		repeat with currentMail in SelectedMails
			set theSender to sender of currentMail
			set theName to ""
			try
				set theName to name of theSender
			on error
				--nothing to do
			end try

			set theAddress to address of theSender
			set theDate to time received of currentMail
			set theSubject to subject of currentMail
			set this_data to this_data & theName & " " & theAddress & ", " & theDate & ", " & theSubject & return
		end repeat

		set this_file to ((theFolder as text) & "emacs-outlook-compli-email-list.csv")
		my write_to_file(this_data, this_file, true)
	end tell
end run

on write_to_file(this_data, target_file, append_data)
	try
		set the target_file to the target_file as string
		set the open_target_file to open for access file target_file with write permission
		if append_data is false then set eof of the open_target_file to 0
		write this_data to the open_target_file as «class utf8» starting at eof
		close access the open_target_file
		return true

	on error
		try
			close access file target_file
		end try
		return false
	end try
end write_to_file
