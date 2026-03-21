import uno

def convert_to_table(*args):
    desktop = XSCRIPTCONTEXT.getDesktop()
    model = desktop.getCurrentComponent()
    sheet = model.getCurrentController().getActiveSheet()

    cursor = sheet.createCursor()
    cursor.gotoEndOfUsedArea(False)
    end_pos = cursor.getRangeAddress()

    if end_pos.EndColumn == 0 and end_pos.EndRow == 0:
        return

    has_headers = show_message_box(model, "Table Headers", "Does your data have headers in the first row?")

    if not has_headers:
        sheet.getRows().insertByIndex(0, 1)
        for col in range(end_pos.EndColumn + 1):
            sheet.getCellByPosition(col, 0).String = f"Column {col + 1}"
        cursor.gotoEndOfUsedArea(False)
        end_pos = cursor.getRangeAddress()

    table_range = sheet.getCellRangeByPosition(0, 0, end_pos.EndColumn, end_pos.EndRow)
    range_addr = table_range.getRangeAddress()

    db_ranges = model.DatabaseRanges
    db_name = "MyFormattedTable"

    if not db_ranges.hasByName(db_name):
        db_ranges.addNewByName(db_name, range_addr)
    else:
        db_ranges.removeByName(db_name)
        db_ranges.addNewByName(db_name, range_addr)

    db_entry = db_ranges.getByName(db_name)
    db_entry.AutoFilter = True

    header_bg  = int("004586", 16)
    text_white = int("FFFFFF", 16)
    row_alt_bg = int("DCE6F1", 16)

    header_range = sheet.getCellRangeByPosition(0, 0, end_pos.EndColumn, 0)
    header_range.CellBackColor = header_bg
    header_range.CharColor = text_white
    bold = uno.getConstantByName("com.sun.star.awt.FontWeight.BOLD")
    header_range.CharWeight = bold

    for row_idx in range(1, end_pos.EndRow + 1):
        target_row = sheet.getCellRangeByPosition(0, row_idx, end_pos.EndColumn, row_idx)
        if row_idx % 2 == 0:
            target_row.CellBackColor = row_alt_bg
        else:
            target_row.CellBackColor = -1

    for col_idx in range(end_pos.EndColumn + 1):
        column = sheet.getColumns().getByIndex(col_idx)
        column.OptimalWidth = True


def show_message_box(doc, title, message):
    parent_window = doc.getCurrentController().getFrame().getContainerWindow()
    toolkit = parent_window.getToolkit()
    YES_NO = uno.getConstantByName("com.sun.star.awt.MessageBoxButtons.BUTTONS_YES_NO")
    IDYES  = uno.getConstantByName("com.sun.star.awt.MessageBoxResults.YES")

    msgbox = toolkit.createMessageBox(
        parent_window,
        uno.Enum("com.sun.star.awt.MessageBoxType", "QUERYBOX"),
        YES_NO,
        title,
        message
    )
    return msgbox.execute() == IDYES


g_exportedScripts = (convert_to_table,)