# Word Header Footer Reader for NVDA
# Author: Nandan Kumar Singh
# Version: 0.1.2
#
# Commands:
# NVDA+Shift+H = Read current header
# NVDA+Shift+F = Read current footer
# NVDA+Shift+W = Read header and footer
#
# Uses Word's native accessibility/COM object exposed by the active Word window.
# This avoids relying only on Word.Application being registered in the COM
# Running Object Table.

import ctypes
from ctypes import byref
import re

import globalPluginHandler
from scriptHandler import script
import api
import ui

try:
    from comtypes import POINTER
    from comtypes.automation import IDispatch
    import comtypes.client.dynamic as comDynamic
except Exception:
    POINTER = None
    IDispatch = None
    comDynamic = None


OBJID_NATIVEOM = -16

WD_HEADER_FOOTER_PRIMARY = 1
WD_HEADER_FOOTER_FIRST_PAGE = 2
WD_HEADER_FOOTER_EVEN_PAGES = 3


def _clean_text(value):
    if value is None:
        return ""
    text = str(value)
    text = text.replace("\r", " ").replace("\x07", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text


class GlobalPlugin(globalPluginHandler.GlobalPlugin):

    def _get_word_from_active_window(self):
        """Get Word automation object from the active Word window."""
        if POINTER is None or IDispatch is None or comDynamic is None:
            return None

        try:
            obj = api.getFocusObject()
            hwnd = obj.windowHandle
        except Exception:
            return None

        if not hwnd:
            return None

        try:
            p = POINTER(IDispatch)()
            hr = ctypes.oledll.oleacc.AccessibleObjectFromWindow(
                hwnd,
                OBJID_NATIVEOM,
                byref(IDispatch._iid_),
                byref(p),
            )
            if hr != 0 or not p:
                return None

            window = comDynamic.Dispatch(p)
            word = window.application
            return word
        except Exception:
            return None

    def _get_document(self):
        word = self._get_word_from_active_window()
        if word is None:
            return None, None

        try:
            documents = word.Documents
            if documents.Count == 0:
                return word, None
            return word, word.ActiveDocument
        except Exception:
            return word, None

    def _get_current_page(self, word):
        try:
            # Word WdInformation.wdActiveEndPageNumber = 3
            return int(word.Selection.Information(3))
        except Exception:
            return 1

    def _get_current_section(self, word, doc):
        try:
            return word.Selection.Sections(1)
        except Exception:
            try:
                return doc.Sections(1)
            except Exception:
                return None

    def _variant_order(self, section, page):
        variants = []

        try:
            first = bool(section.PageSetup.DifferentFirstPageHeaderFooter)
        except Exception:
            first = False

        try:
            odd_even = bool(section.PageSetup.OddAndEvenPagesHeaderFooter)
        except Exception:
            odd_even = False

        if page == 1 and first:
            variants.append(("First Page", WD_HEADER_FOOTER_FIRST_PAGE))

        if odd_even and page % 2 == 0:
            variants.append(("Even Pages", WD_HEADER_FOOTER_EVEN_PAGES))

        variants.append(("Primary", WD_HEADER_FOOTER_PRIMARY))

        # Remove duplicate indexes while preserving order.
        seen = set()
        result = []
        for item in variants:
            if item[1] not in seen:
                seen.add(item[1])
                result.append(item)
        return result

    def _collect(self, kind):
        word, doc = self._get_document()

        if word is None:
            return None, "Microsoft Word active window not found."

        if doc is None:
            return None, "No active Microsoft Word document found."

        section = self._get_current_section(word, doc)
        if section is None:
            return None, "Could not determine the current Word section."

        page = self._get_current_page(word)
        collection_name = "Headers" if kind == "header" else "Footers"
        results = []

        for label, index in self._variant_order(section, page):
            try:
                collection = getattr(section, collection_name)
                hf = collection(index)

                try:
                    if not bool(hf.Exists):
                        continue
                except Exception:
                    pass

                text = _clean_text(hf.Range.Text)
                if text:
                    results.append((label, text))
            except Exception:
                continue

        return results, None

    def _edit(self, kind):
        """Move the Word selection into the current section's header/footer."""
        word, doc = self._get_document()

        if word is None:
            ui.message("Microsoft Word active window not found.")
            return

        if doc is None:
            ui.message("No active Microsoft Word document found.")
            return

        section = self._get_current_section(word, doc)
        if section is None:
            ui.message("Could not determine the current Word section.")
            return

        page = self._get_current_page(word)
        collection_name = "Headers" if kind == "header" else "Footers"
        name = "Header" if kind == "header" else "Footer"

        # Select the most appropriate Header/Footer variant for the current page.
        selected = False
        for label, index in self._variant_order(section, page):
            try:
                collection = getattr(section, collection_name)
                hf = collection(index)

                try:
                    if not bool(hf.Exists):
                        continue
                except Exception:
                    pass

                # Select the Word range. This moves the insertion point into
                # Header/Footer editing mode in Microsoft Word.
                hf.Range.Select()
                selected = True
                break
            except Exception:
                continue

        if not selected:
            ui.message("No %s content found in this section." % name)
            return

        try:
            word.Activate()
        except Exception:
            pass

        ui.message("%s editing mode." % name)

    def _speak(self, kind):
        results, error = self._collect(kind)

        if error:
            ui.message(error)
            return

        name = "Header" if kind == "header" else "Footer"

        if not results:
            ui.message("No %s content found in this section." % name)
            return

        parts = []
        for label, text in results:
            parts.append("%s, %s: %s" % (name, label, text))

        ui.message(". ".join(parts))

    @script(gesture="kb:NVDA+control+shift+h")
    def script_editHeader(self, gesture):
        """Enter editing mode for the current Word section header."""
        self._edit("header")

    @script(gesture="kb:NVDA+control+shift+f")
    def script_editFooter(self, gesture):
        """Enter editing mode for the current Word section footer."""
        self._edit("footer")

    @script(gesture="kb:NVDA+shift+h")
    def script_readHeader(self, gesture):
        """Read the current Word section header."""
        self._speak("header")

    @script(gesture="kb:NVDA+shift+f")
    def script_readFooter(self, gesture):
        """Read the current Word section footer."""
        self._speak("footer")

    @script(gesture="kb:NVDA+shift+w")
    def script_readHeaderFooter(self, gesture):
        """Read the current Word section header and footer."""
        headers, error = self._collect("header")
        if error:
            ui.message(error)
            return

        footers, error = self._collect("footer")
        if error:
            ui.message(error)
            return

        parts = []

        if headers:
            parts.append("Header: " + ". ".join(
                "%s: %s" % (label, text) for label, text in headers))
        else:
            parts.append("Header not found")

        if footers:
            parts.append("Footer: " + ". ".join(
                "%s: %s" % (label, text) for label, text in footers))
        else:
            parts.append("Footer not found")

        ui.message(". ".join(parts))
