# IMAGELOCAL for Excel

Display pictures from local file paths using a worksheet formula:

```excel
=IMAGELOCAL(A2)
```

IMAGELOCAL is a VBA add-in for Windows desktop Excel. Install it once, then use the function in existing and new workbooks, including ordinary `.xlsx` files.

The add-in embeds a picture over the formula cell and fits it inside the cell while preserving its proportions. The picture is a floating Excel object; the formula itself returns blank text when successful.

## Requirements

- Windows desktop Excel with VBA and Excel add-ins permitted.
- The `ImageLocal.xlam` add-in installed and enabled.
- Read access to the image files. Use full Windows drive paths or accessible UNC network paths, including the filename and extension.
- A worksheet that allows changes to drawing objects.

This implementation is for Windows. Excel for the web, mobile Excel, and Excel for Mac are not supported by this add-in.

## Installation

1. Download [ImageLocal.xlam](./ImageLocal.xlam). On the GitHub file page, use **Download raw file**.
2. Keep the file in a permanent folder on your computer. If you downloaded a ZIP, extract it first.
3. In Excel, open **File → Options → Add-ins**.
4. At the bottom, set **Manage** to **Excel Add-ins**, then click **Go**.
5. Click **Browse**, select `ImageLocal.xlam`, and click **OK**.
6. Check the box beside the add-in and click **OK**.

Excel will load the enabled add-in in future sessions. You do not need to paste VBA code into each workbook. See [Microsoft's instructions for reusable Excel functions](https://support.microsoft.com/en-us/excel/create-custom-functions-in-excel).

### If Excel blocks the downloaded add-in

If you trust the downloaded file, close Excel, right-click `ImageLocal.xlam` in File Explorer, and select **Properties**. On the **General** tab, select **Unblock** if it appears, then click **Apply** and **OK**. Reopen Excel and enable the add-in.

If your organization still blocks it, follow its process for approving VBA add-ins. See [Microsoft's guidance for downloaded Excel add-ins](https://learn.microsoft.com/en-us/microsoft-365-apps/security/internet-macros-blocked).

## Quick start

1. Put the full image path in **A2**, for example `C:\Photos\photo.jpg`.
2. Enter `=IMAGELOCAL(A2)` in **B2**.
3. Set column B's width and row 2's height to the picture size you want.
4. Press **F9** to recalculate and fit the picture.
5. Fill the formula down to display pictures for the remaining paths.

| Cell | Content |
| --- | --- |
| A2 | `C:\Photos\photo.jpg` |
| B2 | `=IMAGELOCAL(A2)` |
| A3 | `C:\Photos\another-photo.png` |
| B3 | `=IMAGELOCAL(A3)` |

Changing a source path updates its picture on recalculation. Clearing the source path removes the picture on recalculation. Deleting the formula removes the picture when Excel's events are enabled.

## Formula examples

Reference a cell containing the full path:

```excel
=IMAGELOCAL(A2)
```

Enter the path directly:

```excel
=IMAGELOCAL("C:\Photos\photo.jpg")
```

Build a path when A2 contains a filename without its extension, such as `photo`:

```excel
=IMAGELOCAL("C:\Photos\"&A2&".jpg")
```

Use an accessible network folder:

```excel
=IMAGELOCAL("\\server\shared\Photos\photo.jpg")
```

Use one direct `IMAGELOCAL` formula per destination cell. Wrapping it inside `IF`, `IFERROR`, or `LET`, or using it as an array/spill formula, is not supported. For conditional display, calculate the path or an empty string in a helper cell, then pass that cell to `IMAGELOCAL`.

## Refreshing pictures

Press **F9** after changing row heights or column widths, updating an image file on disk, or changing paths while Excel is in manual calculation mode. The add-in checks files during recalculation; it does not watch folders continuously.

For a refresh limited to the active workbook:

1. Press **Alt+F8**.
2. Enter the following in **Macro name**, even if it is not listed:

   ```text
   'ImageLocal.xlam'!ImageLocal_Refresh
   ```

3. Click **Run**.

Existing pictures are reused when the path, file modification time, and file size match. If you replace a file while preserving both its modification time and size, clear and re-enter the destination formula to reload it.

## Behavior and limitations

- **Pictures float over cells.** They are embedded shapes rather than native Picture in Cell values. Another formula cannot retrieve the displayed picture as the cell's value.
- **Pictures follow the cell layout.** They are set to move and size with cells. After sorting or filtering, recalculate and check that the pictures still align with the intended rows.
- **Paths must be absolute.** Relative paths such as `photo.jpg`, web URLs, and `file://` URLs are not supported.
- **Image support depends on Excel.** Start with JPG or PNG files. Files Excel cannot insert display a `cannot load picture` message.
- **Recalculation can take time with large lists.** Each `IMAGELOCAL` formula is volatile, so Excel evaluates it whenever it recalculates.
- **Drawing requires Excel events.** If another macro disables events, automatic picture updates stop until events are restored. The refresh macro can also draw queued pictures.
- **Pictures can intercept mouse clicks.** To edit the underlying formula, select its cell using the Name Box or keyboard, then use the formula bar.

## Sharing workbooks

The pictures are embedded in the workbook when it is saved. Sharing the workbook therefore includes those pictures and any file paths stored in the sheet.

To recalculate and update the pictures, another user needs the add-in installed and access to the source files at the referenced paths. Without the add-in, recalculating the formulas can produce `#NAME?`. If the files are in a different location, update the path column.

The VBA code stays in `ImageLocal.xlam`, so the workbook containing your data can remain an `.xlsx` file.

## Troubleshooting

| Message or symptom | What to check |
| --- | --- |
| `#NAME?` | Confirm that the add-in is installed, enabled, and allowed to run. Check the spelling of `IMAGELOCAL`. |
| `IMAGELOCAL: enable the ImageLocal add-in` | Enable the add-in. If it is already checked, uncheck it, click OK, then enable it again. |
| `IMAGELOCAL: file not found` | Check the full path, filename, extension, and access to the drive or network folder. |
| `IMAGELOCAL: use a full local file path` | Use a path such as `C:\Photos\photo.jpg` or `\\server\share\photo.jpg`. |
| `IMAGELOCAL: invalid or inaccessible path` | Pass one cell containing a path, or a text expression that produces a path. Check file access. |
| `IMAGELOCAL: use =IMAGELOCAL(path) directly` | Remove any surrounding formula. Put conditional path logic in a helper cell. |
| `IMAGELOCAL: use one formula per cell` | Enter the formula in one destination cell, then fill it down. |
| `IMAGELOCAL: sheet drawing objects are protected` | Allow drawing-object changes on the worksheet, then recalculate. |
| `IMAGELOCAL: cannot load picture` | Open the file outside Excel to check it. Try a valid JPG or PNG, then recalculate. |
| Picture is tiny, distorted, or out of date | Adjust the row height and column width, then press F9. |
| Formula cell is blank but no picture appears | Check that the source path is not blank, Excel events are enabled, and the destination row and column are visible. Try the refresh macro. |
| Add-in stopped loading after moving its file | Browse to its new location in the Excel Add-ins dialog and enable it again. |

## Build from the VBA source

If you are using the source files instead of the prepared `.xlam`:

| File | Purpose |
| --- | --- |
| `Import_ImageLocalCore.bas` | Formula, file checks, picture insertion, fitting, and refresh procedures. |
| `Import_CImageLocalEvents.cls` | Application events for drawing after calculation and removing obsolete pictures. |
| `Paste_into_ThisWorkbook.txt` | Startup and shutdown events for the add-in. |

1. Create a new blank Excel workbook and press **Alt+F11**.
2. Select that workbook's project. Use **File → Import File** to import the `.bas` and `.cls` files.
3. Open **ThisWorkbook** and paste the contents of `Paste_into_ThisWorkbook.txt`.
4. Choose **Debug → Compile VBAProject**.
5. Return to Excel and save as **Excel Add-in (*.xlam)**, named `ImageLocal.xlam`.
6. Close the workbook used to build the add-in, then follow the installation steps above.

## Uninstall

Open **File → Options → Add-ins → Manage: Excel Add-ins → Go**, clear the add-in's checkbox, and click **OK**.

Existing embedded pictures remain in saved workbooks. Updating them through `IMAGELOCAL` requires the add-in to be enabled again.

## Reporting a problem

Open an issue with your Windows and Excel versions, the formula you used, the image format, the exact error message, and steps to reproduce it. A small sample workbook and image can help; replace personal data with sample data before attaching them.
