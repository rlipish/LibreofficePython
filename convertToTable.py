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

    # Clear existing row colors before re-banding
    data_range = sheet.getCellRangeByPosition(0, 1, end_pos.EndColumn, end_pos.EndRow)
    data_range.CellBackColor = -1

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
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    dp = smgr.createInstanceWithContext("com.sun.star.awt.DialogProvider", ctx)

    # Build dialog programmatically
    dialog_model = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)
    dialog_model.Width  = 220
    dialog_model.Height = 70
    dialog_model.Title  = title

    # Label
    label_model = dialog_model.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
    label_model.PositionX = 10
    label_model.PositionY = 10
    label_model.Width     = 200
    label_model.Height    = 25
    label_model.Label     = message
    label_model.MultiLine = True
    dialog_model.insertByName("lbl", label_model)

    # Yes button — PushButtonType 1 = OK, makes it the default/Enter button
    yes_model = dialog_model.createInstance("com.sun.star.awt.UnoControlButtonModel")
    yes_model.PositionX     = 60
    yes_model.PositionY     = 45
    yes_model.Width         = 40
    yes_model.Height        = 15
    yes_model.Label         = "Yes"
    yes_model.PushButtonType = 1   # OK — accepts on Enter
    yes_model.DefaultButton  = True
    dialog_model.insertByName("btnYes", yes_model)

    # No button — PushButtonType 2 = CANCEL
    no_model = dialog_model.createInstance("com.sun.star.awt.UnoControlButtonModel")
    no_model.PositionX      = 110
    no_model.PositionY      = 45
    no_model.Width          = 40
    no_model.Height         = 15
    no_model.Label          = "No"
    no_model.PushButtonType = 2   # CANCEL
    dialog_model.insertByName("btnNo", no_model)

    # Show it
    dialog_ctrl = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)
    dialog_ctrl.setModel(dialog_model)
    parent_window = doc.getCurrentController().getFrame().getContainerWindow()
    dialog_ctrl.createPeer(parent_window.getToolkit(), parent_window)
    result = dialog_ctrl.execute()
    dialog_ctrl.dispose()

    return result == 1  # 1 = OK/Yes, 0 = Cancel/No


g_exportedScripts = (convert_to_table,)