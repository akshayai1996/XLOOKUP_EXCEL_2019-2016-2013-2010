# XLOOKUP for Excel 2019, 2016, 2013, 2010

This repository provides a native **XLOOKUP** function for older versions of Excel that don't have it built-in. It is implemented as a high-performance `.xll` add-in using C# and Excel-DNA.

---

## 🚀 Quick Download
Go to the [Release](./Release) folder and download the version that matches your Excel (32-bit or 64-bit).

- **[XLookup_64bit.xll](./Release/XLookup_64bit.xll)** (Most Common)
- **[XLookup_32bit.xll](./Release/XLookup_32bit.xll)**

---

## 🔍 How to check if your Excel is 32-bit or 64-bit

Before installing, you must know which version of Excel you are running. 

1. Open **Excel**.
2. Click on the **File** tab in the top-left corner.
3. Click on **Account** (or **Help** in some older versions) in the left sidebar.
4. Click the **About Excel** button.
5. A window will pop up. At the very top, it will show the version number followed by either **32-bit** or **64-bit**.

> [!TIP]
> Most modern computers run **64-bit** Excel, but many older corporate environments still use **32-bit**.

---

## 🛠️ Detailed Installation Guide

Once you have downloaded the correct `.xll` file, follow these steps to install it permanently:

### Step 1: Open Excel Add-ins Menu
1. Open Excel.
2. Go to **File > Options**.
3. In the Excel Options window, click on **Add-ins** on the left side.
4. At the bottom of the window, you will see a **Manage** dropdown. Ensure it says **Excel Add-ins** and click **Go...**.

### Step 2: Load the Add-in
1. In the Add-ins window, click the **Browse...** button.
2. Find the `.xll` file you downloaded (`XLookup_64bit.xll` or `XLookup_32bit.xll`).
3. Select the file and click **OK**.
4. Excel might ask if you want to copy the add-in to your library folder. Click **Yes** so it stays installed even if you move the original download.
5. Make sure the add-in is **checked** in the list and click **OK**.

### Step 3: Start using XLOOKUP
In any cell, type:
`=XLOOKUP(lookup_value, lookup_array, return_array)`

It will work exactly like the native formula in Office 365!

---

## ⚠️ Troubleshooting

### "Security Warning: Application Add-ins have been disabled"
If you see a security bar at the top, you may need to "Unblock" the file:
1. Right-click the downloaded `.xll` file in your Downloads folder.
2. Select **Properties**.
3. At the bottom, look for **Security** and check the box that says **Unblock**.
4. Click **OK** and try loading it in Excel again.

### "This add-in is not a valid Office Add-in"
This error usually means you tried to load a **64-bit** add-in into **32-bit** Excel, or vice versa. Double-check your Excel bitness (see instructions above) and try the other file.

---

## 💻 Technical Details
- **Target Framework**: .NET Framework 4.7.2 (Built into Windows 10/11)
- **Library**: Excel-DNA 1.6.0
- **Language**: C#
