# PasteMarkdown 🚀

Lightweight VBA solution to **paste Markdown** into Microsoft Word or Outlook to avoid manual formatting! 🎉

## 🌟 Overview

PasteMarkdown is a VBA macro for Microsoft Word and Outlook that converts clipboard Markdown into formatted Office content—headings, lists, bold/italic, code blocks, links, and more.

Ideal for inserting Markdown (e.g., AI-generated text) into documents or emails without manual styling.

## 📦 Features

* **Headings**: `#` → Word Heading 1…6 styles
* **Lists**: Unordered (`-`, `*`, `+`) and ordered (`1.`) with nesting support
* **Bold & Italic**: `**bold**`, `*italic*`, `***both***`
* **Blockquotes**: `> quoted text` → Word Quote style
* **Fenced Code Blocks**: `…` → monospaced, shaded Code style
* **Inline Code**: `` `code` `` → Consolas + shading
* **Strikethrough**: `~~strike~~` → strike-through formatting
* **Links**: `[text](https://...)` → live hyperlinks

## ⚙️ Versions

1. **Word**: VBA macro to paste Markdown into MS Word documents.
2. **Outlook**: VBA macro to paste Markdown into Outlook email editor.

## 🚀 Installation & Usage

1. Open the VBA editor (Alt + F11) in Word or Outlook.
2. Insert a new module and paste the corresponding code from this repository.
3. Enable/Add the references listed in the comment section at the bginning of each macro (Tools → References).
4. Place your cursor where you want to paste Markdown.
5. Run the **`PasteMarkdown`** (Word) or **`PasteMarkdownInEmail`** (Outlook) macro.

> 💡 **Tip:** For faster access, add a button to the Ribbon or assign a keyboard shortcut to the macro!

Enjoy perfectly formatted Markdown in your documents and emails! 🥳

## ⚠️ Limitations

This macro covers many common Markdown features, but **does not support**:

* Images (`![alt](url)`)
* Tables (`| col1 | col2 |` rows)
* Footnotes and references (`[^1]`, `[1]: url`)
* Nested blockquotes deeper than one level
* Task lists (`- [ ]`, `- [x]`)
* Definition lists, math formulas, and HTML blocks

❗️ It’s unlikely these will be added in the future—feel free to fork, customize or extend if you need them!

## 📜 License

This project is released under the **MIT License**. See [LICENSE](LICENSE) for details.
