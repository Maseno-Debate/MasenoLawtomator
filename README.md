# 📘 **MasenoLawtomator**

*A starter collection of VBA macros for Maseno University LLB w IT students.*

---

## 🎓 **Overview**

**MasenoLawtomator** is a lightweight legal automation toolkit designed to help **Maseno University Law students** draft faster and produce cleaner, properly formatted documents.

This initial version focuses on **Microsoft Word VBA macros**, allowing students to automatically formattheir documents.

The goal: make legal drafting easier, without needing to know Git or advanced programming.

---

## ✨ **What’s Inside**

### ✔️ **VBA Macros for Microsoft Word**

* Auto-format assignments or clinic documents
* Fix spacing & paragraph alignment
* Legal-style numbered headings
* Quick citation formating helpers
* Cleanup tools (remove double spaces, fix broken numbering, etc.)

### ✔️ **Ready-to-Use Macros**

All macros are available as `.bas` files in the repo and can also be **copied directly from GitHub** into the Word VBA editor.

---

## 🚀 **How to Use `formatting.bas`** (Press each step to drop down instructions)

<details>
<summary><b>Step 1: Enable Developer Tab in Word</b></summary>

1. Open Microsoft Word.
2. Go to **File → Options → Customize Ribbon**.
3. On the right side, check **Developer**.
4. Click **OK**.

*The Developer tab will now appear in your Word ribbon.*

</details>

<details>
<summary><b>Step 2: Enable Macros</b></summary>

1. Go to **File → Options → Trust Center → Trust Center Settings → Macro Settings**.
2. Select **Enable all macros** (or **Disable all macros with notification** if you want to approve macros each time).
3. Click **OK**.

*This allows your VBA macros to run safely.*

</details>

<details>
<summary><b>Step 3: Import `formatting.bas`</b></summary>

**Option A – Download and Import**

1. Click [`formatting.bas`](https://github.com/Maseno-Debate/MasenoLawtomator/blob/main/vba/formatting.bas) or go to the `vba/` folder → **Raw** → Right-click → **Save As**.
2. Open Word → press **Alt + F11** to open the VBA editor.
3. Go to **File → Import File…** → select the downloaded `.bas`.
4. The module will now appear in your project.

**Option B – Copy-Paste**

1. Open [`formatting.bas`](https://github.com/Maseno-Debate/MasenoLawtomator/blob/main/vba/formatting.bas), → copy all code.
2. Open Word → **Alt + F11**.
3. Insert a new module → paste the code.
4. Save the document.

</details>

<details>
<summary><b>Step 4: Run the Macro / Use Shortcut</b></summary>

1. Run via **Developer → Macros → select `FormatWholeDocument` → Run**.
2. The first time you run it, it will **automatically assign CTRL + ALT + A** as a shortcut.
3. From then on, simply press **CTRL + ALT + A** to instantly format the document with:

   * 1.5 line spacing
   * Times New Roman, size 12, black font
   * Justified paragraphs (left-aligned if too short)
   * Standard legal margins

*No further setup required.*

</details>

---

## 📂 **Project Structure**

<details>
<summary>View Directory</summary>

```
MasenoLawtomator/
│
├── vba/
│   ├── formatting.bas     # Document formatting macro
│   ├── numbering.bas      # Automatic numbering tools
│   ├── cleanup.bas        # Document cleanup scripts
│   └── citations.bas      # Citation helpers
│
├── CONTRIBUTING.md
├── LICENSE
├── README.md
├── DIRECTORY.md
└── DIRECTORY.txt
```

*All `.bas` files can be downloaded or copied into Word VBA editor.*

</details>

---

## 🛣️ **Future Additions**

* More Word automation macros (opinions, submissions, affidavits)
* Google Docs automation (Apps Script)
* Case-law note organizers
* Law clinic intake forms and report templates

---

## 🤝 **Contributing**

**Contributors** who want to add new macros or update scripts can **clone the repo** or submit pull requests.
Git is not required for users.

---

## 📄 **License**

This project is licensed under the **Apache License 2.0**.
See the `LICENSE` file for full terms.

---

## 👤 **Author**

**James E. Limbe**
Maseno University School of Law
GitHub: **[B0mb37](https://github.com/B0mb37)**