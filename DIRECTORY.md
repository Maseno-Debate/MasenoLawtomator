# 📁 **Project Directory Structure**

<details>
<summary><b>Root Directory</b></summary>

```
.
├── CONTRIBUTING.md
├── DIRECTORY.md
├── DIRECTORY.txt
├── LICENSE
├── README.md
├── docs/
└── vba/
```

**Descriptions**

* **CONTRIBUTING.md** — Guidelines for contributors and coding standards.
* **DIRECTORY.md** — Human-readable directory tree of the project.
* **DIRECTORY.txt** — Auto-generated ASCII directory tree (from `tree /F /A`).
* **LICENSE** — Project license file (Apache 2.0).
* **README.md** — Project overview, usage instructions, and setup guide.
* **docs/** — Documentation folder for user guides, screenshots, and tutorials.
* **vba/** — All VBA modules used in Word formatting tools.

</details>

---

<details>
<summary><b>/docs</b></summary>

```
docs/
```

**Descriptions**

* *(Currently empty)* — Reserved for manuals, examples, and extended documentation.

</details>

---

<details>
<summary><b>/vba</b></summary>

```
vba/
├── citations.bas
├── cleanup.bas
├── formatting.bas
└── numbering.bas
```

**Descriptions**

* **citations.bas** — Tools for legal citations, reference formatting, and academic style.
* **cleanup.bas** — Fixes spacing, removes extra breaks, normalizes text, and sanitizes formatting.
* **formatting.bas** — Main styling engine (margins, spacing, fonts, justification, shortcut keys).
* **numbering.bas** — Automatic numbering tools for headings, sections, paragraphs, and lists.

</details>