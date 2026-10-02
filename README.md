# Laptop Check

One report for your laptop: manufacturer support app, Windows version,
drivers, BIOS, touchscreen and screen. It tells you what needs doing and
how to do it.

It only reads information. Nothing is installed, removed or changed, and no
administrator rights are needed.

---

## Step 1 — Download

**Right-click** this link and choose **Save link as…**

https://raw.githubusercontent.com/screenologist/touch_windows/main/LaptopCheck.ps1

Save it to your **Desktop**.

> A normal left-click opens the file as text in your browser. Use right-click.
> The file name must stay exactly `LaptopCheck.ps1`.

---

## Step 2 — Open PowerShell

1. Press the **Windows key**
2. Type `powershell`
3. Click **Windows PowerShell**

A window opens. It may be black or dark blue — either is correct.

---

## Step 3 — Go to your Desktop

Type this and press **Enter**:

```
cd ([Environment]::GetFolderPath('Desktop'))
```

---

## Step 4 — Run the check

Type this and press **Enter**:

```
powershell -ExecutionPolicy Bypass -File .\LaptopCheck.ps1
```

It takes **2 to 5 minutes** — most of that is asking Windows Update what is
available. When it finishes, the report opens in Notepad and is saved to your
Desktop as `LaptopCheck_<date>.txt`.

---

## Step 5 — Read "WHAT TO DO"

The last section of the report lists, in order, what needs attention and how
to fix it. If anything is unclear, or the touchscreen still does not work,
email us the report file.

---

## What the report checks

| Section | What it tells you |
|---|---|
| Laptop | Make, model, SKU and serial number |
| Support app | Whether the manufacturer's update app is installed, and where to get it if not |
| Windows | Your version, whether it still gets security updates, and updates waiting |
| Drivers | Driver updates available, devices with problems, where to download drivers |
| BIOS | Your BIOS version, whether a newer one is offered, where to download it |
| Touchscreen | Whether the laptop had a touchscreen when Windows was first set up, whether it works now, and step-by-step fixes |
| Screen | PnP ID, part number (when the screen stores one), resolution, refresh rate and size, plus earlier screens this laptop has used |

Support apps covered: Lenovo Vantage, HP Support Assistant, Dell SupportAssist /
Command | Update, AcerSense, MyASUS / Armoury Crate, MSI Center, Samsung Update,
GIGABYTE Control Center, LG Update, HUAWEI / HONOR PC Manager, dynabook Service
Station, Fujitsu DeskUpdate, VAIO Update. Surface, Razer, Xiaomi, Panasonic,
Framework and Clevo/TongFang-based laptops get the right download page instead.

---

## Questions

**Is this safe?**
Yes. It reads information and writes one text file to your Desktop.

**"...running scripts is disabled on this system"**
Retype the Step 4 command exactly, including `-ExecutionPolicy Bypass`. If the
message still appears, the laptop may be managed by an employer or school that
blocks scripts. Ask their IT department.

**"Cannot find path ... Desktop because it does not exist"**
Use the Step 3 command exactly as written — it works with OneDrive Desktops.

**It cannot find `LaptopCheck.ps1`**
The file is not on your Desktop, or its name is not exactly `LaptopCheck.ps1`
(for example it ends in `.txt`). Download it again with right-click and
**Save link as…**, and save it to your Desktop.

**It is taking a long time.**
Windows Update can be slow. Wait up to 5 minutes. To skip that part, run:
`powershell -ExecutionPolicy Bypass -File .\LaptopCheck.ps1 -SkipWindowsUpdate`

**The window closed too fast.**
The report is saved on your Desktop. Open it from there.

---

## For technicians

| Option | Use |
|---|---|
| `-Label "RMA-4471"` | Adds a reference to the report |
| `-SkipWindowsUpdate` | Faster run, no online check |
| `-OutputPath C:\Reports` | Save the report somewhere other than the Desktop |
| `-NoPause` | Do not wait for Enter at the end (for scripted runs) |

**Short version:** `PanelID.ps1` prints only the laptop, the screen (PnP ID,
part number, max refresh) and the touch controller. Run it the same way; it
accepts `-Label`, `-OutputPath` and `-NoPause` too. About half of laptop
screens store their part number (roughly 4,200 of 8,300 laptop-size screens in
a public collection of real screen data); for the rest you still get the PnP ID
and refresh rate.

**Testing:** put `Test-LaptopCheck.ps1` next to `LaptopCheck.ps1` and run
`powershell -ExecutionPolicy Bypass -File .\Test-LaptopCheck.ps1`. It runs the
real script against simulated laptops and real screen data and must end with
`ALL TESTS PASSED`. Run it after every change.

**Yearly maintenance:** the Windows support dates at the top of the script
were reviewed on 2026-10-01. Update the table after each autumn Windows
release (source: Microsoft Lifecycle, Windows 11 Home and Pro / Enterprise and
Education).

Licensed under the MIT License.
