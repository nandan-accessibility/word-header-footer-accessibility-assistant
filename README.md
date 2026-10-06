# Word Header & Footer Accessibility Assistant

Word Header & Footer Accessibility Assistant is an NVDA add-on designed to improve access to Microsoft Word headers and footers for blind and visually impaired users.

The add-on provides keyboard commands to read and edit headers and footers in the current Microsoft Word document.

## Features

- Read the current Header.
- Read the current Footer.
- Read both Header and Footer.
- Edit the current Header.
- Edit the current Footer.
- Support for different first-page headers and footers when configured in the document.
- Support for different odd-page and even-page headers and footers when configured in the document.

## Keyboard Commands

| Shortcut | Action |
|---|---|
| `NVDA+Shift+H` | Read current Header |
| `NVDA+Shift+F` | Read current Footer |
| `NVDA+Shift+W` | Read both Header and Footer |
| `NVDA+Ctrl+Shift+H` | Edit current Header |
| `NVDA+Ctrl+Shift+F` | Edit current Footer |

## Usage

Open a Microsoft Word document and place the NVDA focus inside the document.

### Reading a Header

Press:

`NVDA+Shift+H`

The add-on reads the applicable Header for the current Word section.

### Reading a Footer

Press:

`NVDA+Shift+F`

The add-on reads the applicable Footer for the current Word section.

### Reading Both Header and Footer

Press:

`NVDA+Shift+W`

The add-on reads the applicable Header and Footer.

### Editing a Header

Press:

`NVDA+Ctrl+Shift+H`

The add-on selects the appropriate Header area and enters Word's Header editing context. The Header can then be edited using normal Microsoft Word commands.

### Editing a Footer

Press:

`NVDA+Ctrl+Shift+F`

The add-on selects the appropriate Footer area and enters Word's Footer editing context. The Footer can then be edited using normal Microsoft Word commands.

## First-Page, Odd-Page and Even-Page Headers and Footers

Microsoft Word can use different Headers and Footers for:

- First pages
- Odd pages
- Even pages

The add-on supports these Header and Footer variations when they are configured in the document.

The applicable Header or Footer is determined for the current Word section and page type.

## Sections

Microsoft Word documents can contain multiple sections, each with its own Header and Footer configuration.

The add-on works with the current Word section and accesses the applicable Header or Footer for that section.

## Installation

1. Download the latest `.nvda-addon` file from the GitHub Releases page.
2. Open the downloaded `.nvda-addon` file.
3. Confirm the NVDA installation prompt.
4. Restart NVDA if requested.
5. Open Microsoft Word and use the keyboard commands listed above.

## Requirements

- Windows
- NVDA
- Microsoft Word desktop application

### NVDA Compatibility

Minimum NVDA version: **2021.1**

Last tested NVDA version: **2026.2**

## Notes

- Microsoft Word must be open when using the add-on.
- The NVDA focus should be inside the active Microsoft Word document.
- Header and Footer content depends on the configuration of the Word document.
- Documents using different first-page, odd-page, or even-page Headers and Footers are supported.
- Documents containing multiple sections may have different Header and Footer content for each section.
- If the applicable Header or Footer does not contain content, there may be nothing to read.

## Troubleshooting

If the add-on does not provide Header or Footer information:

1. Make sure Microsoft Word is open.
2. Make sure a Word document is open.
3. Place the NVDA focus inside the Word document.
4. Make sure the Word window is active.
5. Try the appropriate keyboard command again.

If an issue occurs, please report it with the NVDA version, Microsoft Word version, Windows version, keyboard command used, and a description of the problem.

## Feedback and Bug Reports

Feedback, suggestions, and bug reports are welcome.

Please provide the following information when reporting an issue:

- NVDA version
- Microsoft Word version
- Windows version
- Keyboard command used
- What happened
- What you expected to happen
- Any error message announced by NVDA

### Contact

Email: **nandan@nandankumarsingh.in**

GitHub Issues:

https://github.com/nandan-accessibility/word-header-footer-accessibility-assistant/issues

## Project Information

**Project:** Word Header & Footer Accessibility Assistant

**Author:** Nandan Kumar Singh

**Repository:**

https://github.com/nandan-accessibility/word-header-footer-accessibility-assistant

**Releases:**

https://github.com/nandan-accessibility/word-header-footer-accessibility-assistant/releases

**Website:**

https://nandankumarsingh.in/

## License

This add-on is licensed under the **GNU General Public License, version 2 (GPL-2.0)**.

See the `LICENSE` file for the complete license text.

## Author

**Nandan Kumar Singh**

Accessibility Professional focused on digital accessibility, assistive technology, and inclusive technology.

## Current Version

**0.2.1  **  
