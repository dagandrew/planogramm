SheetSeeker() {
    try {
        XL := ComObjActive("Excel.Application")
    } catch {
        MsgBox("Excel is not active or no workbook open.", "Error", "16")
        return
    }
    
    Sheets := XL.ActiveWorkbook.Sheets
    ib := InputBox("Enter part of the sheet name:", "Search Sheet", "w300 h150")
    if ib.Result = "Cancel"
        return
    
    SearchText := Trim(ib.Value)
    Found := false
    
    ; Create normalized versions for comparison
    SearchLower := StrLower(SearchText)
    
    for Sheet in Sheets {
        SheetLower := StrLower(Sheet.Name)
        if InStr(SheetLower, SearchLower) {  ; Case-insensitive match
            Sheet.Activate
            Found := true
            break
        }
    }
    
    if !Found
        MsgBox('No sheet containing "' SearchText '" found.', "Not Found", "48")
}
