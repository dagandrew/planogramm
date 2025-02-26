LWin & Shift::^+L ;filter
F2 & Esc::SelectAllExcel()
F2 & Space::Rename()
F3 & Space::WrapTextAndAutoFitColumns()


#z::UniqueData()
#x::!F8 ;macro
#s::TextOneRow()
#c::Borders()

!s::SheetSeeker()

CapsLock & a::UsualStyle()
CapsLock & s::GreenStyle()
CapsLock & d::NeutralStyle()
CapsLock & f::BadStyle()

^q::JumpFirstSheet()
^w::CreateNewSheet()

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

CreateNewSheet(){
	xlApp := ComObjActive("Excel.Application")
	xlApp.Worksheets.Add()
}

JumpFirstSheet(){
	xlApp := ComObjActive("Excel.Application")
	try xlApp.Sheets(1).Activate
	return
}



GreenStyle() {
	xlApp := ComObjActive("Excel.Application")
    try xlApp.Selection.Style := "Хороший"
    return
}

UsualStyle() {
	xlApp := ComObjActive("Excel.Application")
	try xlApp.Selection.Style := "Обычный"
    return
}

NeutralStyle() {
	xlApp := ComObjActive("Excel.Application")
	try xlApp.Selection.Style := "Нейтральный"
    return
}

BadStyle() {
	xlApp := ComObjActive("Excel.Application")
	try xlApp.Selection.Style := "Плохой"
    return
}



#q::Send("^{PgUp}")
#w::Send("^{PgDn}")

TextOneRow(){
 Send("{Alt}")
 Sleep(300)
 Send("{z}")
 Sleep(300)
 Send("{t}")
 Sleep(300)
}

Borders(){
 Send("^+7")
}

UniqueData(){
 Send("{Alt}")
 Sleep(300)
 Send("{s}")
 Sleep(300)
 Send("{e}")
 Sleep(300)
}

SelectAllExcel(){
	Send("{Home}")
	Send("{Shift down}{End}{Shift up}")
}

WrapTextAndAutoFitColumns() {
	Send("^{v}")
    xlApp := ComObjActive("Excel.Application")
    
    if (xlApp) {
        xlApp.Selection.WrapText := true
        
        xlApp.Selection.Columns.AutoFit
        
        MsgBox("Text wrapping and column autofit applied successfully.", "Success", 64)
    } else {
        MsgBox("Excel is not active.", "Error", 16)
    }
}

#include sheetSeeker.ahk

