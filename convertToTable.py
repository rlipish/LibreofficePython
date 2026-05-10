import uno
import unohelper
from com.sun.star.awt import XKeyListener

def convert_to_table(*args):
    desktop = XSCRIPTCONTEXT.getDesktop()
    model = desktop.getCurrentComponent()
    sheet = model.getCurrentController().getActiveSheet()

    cursor = sheet.createCursor()
    cursor.gotoEndOfUsedArea(False)
    end_pos = cursor.getRangeAddress()

    if end_pos.EndColumn == 0 and end_pos.EndRow == 0:
        return

    result = show_message_box(model, "Table Headers", "Does your data have headers in the first row?")

    if result is None:  # User cancelled / hit X
        return

    has_headers = result

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

    dialog_model = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)
    dialog_model.Width  = 220
    dialog_model.Height = 90
    dialog_model.Title  = title

    # Message label
    label_model = dialog_model.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
    label_model.PositionX = 10
    label_model.PositionY = 8
    label_model.Width     = 200
    label_model.Height    = 25
    label_model.Label     = message
    label_model.MultiLine = True
    dialog_model.insertByName("lbl", label_model)

    # Instruction label
    hint_model = dialog_model.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
    hint_model.PositionX = 10
    hint_model.PositionY = 34
    hint_model.Width     = 200
    hint_model.Height    = 12
    hint_model.Label     = "Type Y or N, or click a button:"
    dialog_model.insertByName("hint", hint_model)

    # Hidden text field that captures typing — starts with focus
    edit_model = dialog_model.createInstance("com.sun.star.awt.UnoControlEditModel")
    edit_model.PositionX  = 80
    edit_model.PositionY  = 46
    edit_model.Width      = 1
    edit_model.Height     = 1
    edit_model.MaxTextLen = 1
    dialog_model.insertByName("capture", edit_model)

    # Yes button
    yes_model = dialog_model.createInstance("com.sun.star.awt.UnoControlButtonModel")
    yes_model.PositionX      = 40
    yes_model.PositionY      = 68
    yes_model.Width          = 50
    yes_model.Height         = 15
    yes_model.Label          = "Yes (Y)"
    yes_model.PushButtonType = 0
    yes_model.DefaultButton  = False
    dialog_model.insertByName("btnYes", yes_model)

    # No button
    no_model = dialog_model.createInstance("com.sun.star.awt.UnoControlButtonModel")
    no_model.PositionX      = 115
    no_model.PositionY      = 68
    no_model.Width          = 50
    no_model.Height         = 15
    no_model.Label          = "No (N)"
    no_model.PushButtonType = 0
    dialog_model.insertByName("btnNo", no_model)

    dialog_ctrl = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)
    dialog_ctrl.setModel(dialog_model)
    parent_window = doc.getCurrentController().getFrame().getContainerWindow()
    dialog_ctrl.createPeer(parent_window.getToolkit(), parent_window)

    state = {"result": None}

    from com.sun.star.awt import XActionListener, XTextListener

    # Button click handlers
    class ActionHandler(unohelper.Base, XActionListener):
        def __init__(self, dlg, answer):
            self.dlg = dlg
            self.answer = answer

        def actionPerformed(self, event):
            state["result"] = self.answer
            self.dlg.endExecute()

    dialog_ctrl.getControl("btnYes").addActionListener(ActionHandler(dialog_ctrl, True))
    dialog_ctrl.getControl("btnNo").addActionListener(ActionHandler(dialog_ctrl, False))

    # Text change handler on the hidden capture field
    class TextHandler(unohelper.Base, XTextListener):
        def __init__(self, dlg, edit_ctrl):
            self.dlg = dlg
            self.edit = edit_ctrl

        def textChanged(self, event):
            text = self.edit.getText().upper()
            if text == "Y":
                state["result"] = True
                self.dlg.endExecute()
            elif text == "N":
                state["result"] = False
                self.dlg.endExecute()
            else:
                self.edit.setText("")  # clear anything else

        def disposing(self, event):
            pass

    capture_ctrl = dialog_ctrl.getControl("capture")
    capture_ctrl.addTextListener(TextHandler(dialog_ctrl, capture_ctrl))

    # Give the capture field focus immediately so typing works right away
    capture_ctrl.setFocus()

    dialog_ctrl.execute()
    dialog_ctrl.dispose()

    return state["result"]


g_exportedScripts = (convert_to_table,)