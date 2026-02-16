# Installing XLOOKUP for Excel 2019

**This is the most compatible version.** It works on standard Windows computers without extra downloads.

1. **Locate the XLL File**:
   - Navigate to: `C:\Users\Asus\Desktop\XLOOKUP_2019\XLookupAddIn\bin\Release\net472`
   - You will see two files:
     - `XLookupAddIn-AddIn64.xll` (for 64-bit Excel - most likely)
     - `XLookupAddIn-AddIn.xll` (for 32-bit Excel)

2. **Open Excel Options**:
   - Open Excel 2019.
   - Go to **File > Options**.
   - Select **Add-ins** from the left menu.

3. **Load the Add-in**:
   - At the bottom, ensure "Excel Add-ins" is selected in the "Manage" dropdown and click **Go...**.
   - Click **Browse...**.
   - Navigate to the folder mentioned in Step 1.
   - Select **`XLookupAddIn-AddIn64.xll`** (try this first).
   - Click **OK**.
   - Ensure it is checked in the list.
   - Click **OK**.

4. **Verify**:
   - Type `=XLOOKUP(` in a cell.

## Troubleshooting

- **"Add-in is not valid"**: This means you picked the wrong bitness. Try the other file (`XLookupAddIn-AddIn.xll`) instead.
- **Security Warning**: Right-click the `.xll` file, select **Properties**, and check **Unblock** at the bottom.
