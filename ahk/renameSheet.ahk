Rename() {
    xlApp := ComObjActive("Excel.Application")
    if (xlApp) {
        currentName := xlApp.ActiveSheet.Name
        newName := InputBox("Enter the new name for the sheet:", "Rename Sheet",, currentName).Value
        
        if (newName != "" && newName != currentName) {
            try {
                ; Attempt to rename the active sheet
                xlApp.ActiveSheet.Name := newName
            } catch as e {
                ; Handle any errors that occur
                MsgBox("An error occurred while renaming the sheet: " e.Message, "Error", 16)
            }
        } else if (newName == currentName) {
            ; Notify the user if the name hasn't changed
            MsgBox("The sheet name was not changed.", "Info", 64)
        }
    } else {
		Send("{F2}")
	}
}
