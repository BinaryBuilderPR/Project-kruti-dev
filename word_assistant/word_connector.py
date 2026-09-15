"""
Direct Microsoft Word COM Connector (Office LTSC 2024 / Office 2016-365).
Reads active document text, inserts Kruti Dev text, applies templates, and highlights errors.
"""

import sys
from typing import Optional, Dict, List, Tuple


class WordConnector:
    """Interacts with the running Microsoft Word application via Windows COM."""
    
    def __init__(self):
        self._word_app = None

    def _get_app(self):
        try:
            import win32com.client
            try:
                # Attach to existing running Word instance
                self._word_app = win32com.client.GetActiveObject("Word.Application")
            except Exception:
                # Start new Word instance if not running
                self._word_app = win32com.client.Dispatch("Word.Application")
                self._word_app.Visible = True
            return self._word_app
        except ImportError:
            print("win32com is not installed. Run 'pip install pywin32' for direct Word COM interaction.")
            return None
        except Exception as e:
            print(f"Could not connect to MS Word: {e}")
            return None

    def is_word_connected(self) -> bool:
        app = self._get_app()
        return app is not None and app.Documents.Count > 0

    def get_active_document_text(self) -> Optional[str]:
        """Returns the full text of the currently active document in MS Word."""
        app = self._get_app()
        if not app or app.Documents.Count == 0:
            return None
        doc = app.ActiveDocument
        return doc.Content.Text

    def get_selection_text(self) -> Optional[str]:
        """Returns the currently selected text in MS Word."""
        app = self._get_app()
        if not app or app.Documents.Count == 0:
            return None
        return app.Selection.Text

    def insert_kruti_text_at_cursor(self, kruti_text: str, font_name: str = "Kruti Dev 010"):
        """Inserts text at current Word cursor position and sets font to Kruti Dev 010."""
        app = self._get_app()
        if not app:
            return False
            
        if app.Documents.Count == 0:
            app.Documents.Add()
            
        selection = app.Selection
        selection.Font.Name = font_name
        selection.TypeText(kruti_text)
        return True

    def replace_word_in_active_doc(self, old_kruti: str, new_kruti: str, font_name: str = "Kruti Dev 010"):
        """Replaces all occurrences of an incorrect Kruti Dev word with the corrected one."""
        app = self._get_app()
        if not app or app.Documents.Count == 0:
            return False
            
        doc = app.ActiveDocument
        # Use Word's native Find & Replace
        find_obj = doc.Content.Find
        find_obj.ClearFormatting()
        find_obj.Replacement.ClearFormatting()
        find_obj.Replacement.Font.Name = font_name
        
        # wdReplaceAll = 2
        find_obj.Execute(
            FindText=old_kruti,
            ReplaceWith=new_kruti,
            Replace=2,
            Forward=True,
            MatchCase=True
        )
        return True

    def insert_template(self, kruti_template_content: str, font_name: str = "Kruti Dev 010"):
        """Inserts an entire official letter template into a new or active Word document."""
        app = self._get_app()
        if not app:
            return False
            
        if app.Documents.Count == 0:
            app.Documents.Add()
            
        selection = app.Selection
        selection.Font.Name = font_name
        selection.Font.Size = 14
        selection.ParagraphFormat.LineSpacingRule = 0 # Single spacing
        selection.TypeText(kruti_template_content)
        return True

