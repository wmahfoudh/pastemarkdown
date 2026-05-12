# PasteMarkdown

Lightweight VBA solution to paste Markdown into Microsoft Word or Outlook and avoid manual formatting.

## Overview

PasteMarkdown is a VBA macro for Microsoft Word and Outlook that converts clipboard Markdown into formatted Office content, including headings, lists, tables, bold and italic text, code blocks, links, task lists, and more.

It is useful for inserting Markdown, including AI-generated text, into Word documents or Outlook emails without manually reformatting the content.

## Features

* **Headings**: `#` to `######` converted to Word Heading 1 to Heading 6 styles
* **Unordered lists**: `-`, `*`, and `+` converted to bullet lists
* **Ordered lists**: `1.`, `2.`, `3.` converted to numbered lists
* **Nested lists**: Basic nesting support based on leading spaces
* **Task lists**: `- [ ]` and `- [x]` converted to checkbox-style task items
* **Bold and italic**:
  * `**bold**`
  * `__bold__`
  * `*italic*`
  * `_italic_`
  * `***bold italic***`
  * `___bold italic___`
* **Blockquotes**: `> quoted text` converted to Word Quote style
* **Fenced code blocks**: Triple backtick code blocks converted to a monospaced, shaded Code style
* **Inline code**: `` `code` `` converted to Consolas with shading
* **Strikethrough**: `~~strike~~` converted to strikethrough formatting
* **Links**: `https://example.com` converted to live hyperlinks
* **Markdown tables**: Standard pipe tables converted to native Word or Outlook tables
* **Table alignment**:
  * `:---` for left alignment
  * `:---:` for center alignment
  * `---:` for right alignment
* **Escaped pipes in tables**: `\|` supported inside table cells
* **Line breaks inside table cells**: `<br>`, `<br/>`, and `<br />` converted to line breaks inside the cell

## Versions

1. **Word version**
   * Macro name: `PasteMarkdownToWord`
   * Designed to paste Markdown directly into Microsoft Word documents.

2. **Outlook version**
   * Macro name: `PasteMarkdownToEmail`
   * Designed to paste Markdown directly into the Outlook email editor.
   * Uses Outlook's Word-based email editor.

## Installation and Usage

### Word

1. Open Microsoft Word.
2. Press `Alt + F11` to open the VBA editor.
3. Insert a new module.
4. Paste the Word version of the code into the module.
5. Place your cursor where you want to insert the formatted content.
6. Copy Markdown content to the clipboard.
7. Run the `PasteMarkdownToWord` macro.

### Outlook

1. Open Microsoft Outlook.
2. Press `Alt + F11` to open the VBA editor.
3. Insert a new module.
4. Paste the Outlook version of the code into the module.
5. Enable the required Word reference as described in the Dependencies section.
6. Open a new email, reply, or forward.
7. Place your cursor where you want to insert the formatted content.
8. Copy Markdown content to the clipboard.
9. Run the `PasteMarkdownToEmail` macro.

For faster access, add the macro to the Ribbon or Quick Access Toolbar, or assign a keyboard shortcut where supported.

## Dependencies

### Word version

No additional references are required when running the Word version inside Microsoft Word.

The Word object model is already available in Word VBA.

### Outlook version

The Outlook version uses Outlook's built-in Word email editor and early-bound Word objects such as:

```vba
Word.Document
Word.Selection
Word.Range
Word.Table
````

Therefore, in the Outlook VBA editor, enable:

```text
Microsoft Word 16.0 Object Library
```

Depending on your Office version, the number may differ, for example:

```text
Microsoft Word 15.0 Object Library
Microsoft Word 16.0 Object Library
```

The Outlook object model is already available when running the code inside Outlook VBA.

### Regular expressions

No explicit reference to `Microsoft VBScript Regular Expressions 5.5` is required.

The modules use late binding:

```vba
CreateObject("VBScript.RegExp")
```

Because of this, the regex library does not need to be manually enabled under Tools > References.

## Supported Markdown Table Format

PasteMarkdown supports standard Markdown pipe tables.
This is converted into a native Word or Outlook table with:

*   Header row formatting
*   Borders
*   Basic cell padding
*   Column alignment based on the separator row

## Configuration

Both modules include this setting near the top:

```vba
Private Const REMOVE_EMPTY_PARAGRAPHS_AFTER_PASTE As Boolean = False
```

Default:

```vba
False
```

This preserves paragraph spacing, which is usually better for business documents and emails. I found that in Word it is better to set it to `True`.

## Processing Logic

The macro follows this sequence:

1.  Paste clipboard content as plain text.
2.  Process only the newly pasted content.
3.  Normalize manual line breaks.
4.  Convert fenced code blocks.
5.  Convert Markdown tables.
6.  Apply headings, quotes, lists, task lists, inline code, bold, italic, links, and strikethrough.
7.  Optionally remove empty paragraphs.

This order is intentional. Code blocks are protected before inline formatting, and tables are converted before the remaining paragraph-level formatting is applied.

## Limitations

PasteMarkdown covers common Markdown used in business documents, emails, meeting notes, summaries, and AI-generated responses.

It does not currently support:

*   Images: `url`
*   Footnotes and references: `[^1]`, `url`
*   Nested blockquotes deeper than one level
*   Definition lists
*   Math formulas
*   Raw HTML blocks
*   Merged table cells
*   Multiline Markdown tables beyond simple `<br>` line breaks inside cells
*   Complex nested Markdown inside table cells
*   Links with complex URLs containing unmatched closing parentheses

## License

This project is released under the MIT License. See `LICENSE` for details.
