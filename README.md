# 📘 **MasenoLawtomator**

*A starter collection of VBA macros for Maseno University LLB w/ IT students.*

---

## 🎓 **Overview**

**MasenoLawtomator** is a lightweight legal-automation toolkit designed for **Maseno University Law students**, focusing on faster drafting and cleaner formatting.

This first version provides **Microsoft Word VBA macros** that automate formatting tasks and reduce repetitive work—no Git or programming knowledge needed.

---

## ✨ **What’s Inside**

### ✔️ **VBA Macros for Microsoft Word**

* Auto-format assignments, memos & clinic documents
* Fix spacing & alignment
* Generate legal-style numbered headings
* Quick citation formatting helpers
* Cleanup tools (remove double spaces, fix numbering, etc.)

### ✔️ **Ready-to-Use `.bas` Files**

All macros are stored in the `vba/` folder and can be:

* **Copied directly** from GitHub, or
* **Imported** into Word through the VBA editor.

---

## 🚀 **Using `formatting.bas`**

Press each step to expand instructions.

<details>
<summary><b>Step 1: Enable Developer Tab</b></summary>

1. Word → **File → Options → Customize Ribbon**
2. Check **Developer**
3. Click **OK**

</details>

<details>
<summary><b>Step 2: Enable Macros</b></summary>

1. Word → **File → Options → Trust Center**
2. **Trust Center Settings → Macro Settings**
3. Choose **Enable all macros** or **Disable with notification**

</details>

<details>
<summary><b>Step 3: Import <code>formatting.bas</code></b></summary>

**Option A – Download & Import**

1. Open `formatting.bas` → **Raw → Save As**
2. Word → **Alt + F11**
3. **File → Import File…** → select `.bas`

**Option B – Copy-Paste**

1. Open `formatting.bas` → copy code
2. Word → **Alt + F11**
3. Insert new module → paste → save

</details>

<details>
<summary><b>Step 4: Run or Use Shortcut</b></summary>

1. **Developer → Macros → FormatWholeDocument → Run**
2. First run assigns **CTRL + ALT + A** automatically
3. Press **CTRL + ALT + A** anytime to format:

   * Times New Roman, 12
   * 1.5 spacing
   * Standard margins
   * Justified paragraphs

</details>

---

## 📂 **Project Structure**

<details>
<summary><b>View Directory</b></summary>

```
MasenoLawtomator/
│
├── vba/
│   ├── formatting.bas      # Document formatting macro
│   ├── numbering.bas       # Automatic numbering tools
│   ├── cleanup.bas         # Document cleanup scripts
│   └── citations.bas       # Citation helpers
│
├── CONTRIBUTING.md         # Contribution guidelines
├── LICENSE                 # Apache 2.0 license
├── README.md
├── DIRECTORY.md
└── DIRECTORY.txt
```

</details>

---

## 🛣️ **Future Additions**

* Opinion, submissions & pleadings automation
* Google Docs (Apps Script) version
* Case-law & notes organizers
* Clinic intake forms
* Template generators (letters, affidavits, annexures)

---

## 🤝 **Contributing**

Students wishing to contribute can **fork the repo**, create branches, and submit pull requests.

No Git knowledge is needed for users—only for contributors.

---

## 📄 **License**

This project is licensed under the **Apache License 2.0**.
See the `LICENSE` file for full terms.

---

## 👤 **Author**

**James E. Limbe**
Maseno University School of Law
GitHub: **[B0mb37](https://github.com/B0mb37)**