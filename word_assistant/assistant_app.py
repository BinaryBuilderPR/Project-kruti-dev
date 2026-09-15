"""
Smart Kruti Dev 010 Drafting & Spellcheck Assistant for Microsoft Word (Office LTSC 2024).
Designed specifically for the Department of School Education & Literacy, Jharkhand.
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox, font as tkfont
from typing import List, Dict

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from krutidev_engine.converter import kruti_to_unicode, unicode_to_kruti
from krutidev_engine.spellchecker import KrutiSpellChecker
from krutidev_engine.templates import get_all_templates
from word_assistant.word_connector import WordConnector


class KrutiDevAssistantGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("झारखण्ड शिक्षा विभाग - कृति देव 010 प्रारूपक एवं वर्तनी सहायक")
        self.root.geometry("980x720")
        self.root.minsize(800, 600)
        
        # Initialize Core Engines
        self.spellchecker = KrutiSpellChecker()
        self.word_connector = WordConnector()
        self.templates = get_all_templates()
        self.current_suggestions = []
        
        self._setup_theme()
        self._build_ui()

    def _setup_theme(self):
        style = ttk.Style()
        style.theme_use('clam')
        
        # Color palette
        self.bg_color = "#f4f6f9"
        self.primary_color = "#0056b3"
        self.accent_color = "#28a745"
        self.warn_color = "#dc3545"
        self.text_color = "#212529"
        
        self.root.configure(bg=self.bg_color)
        
        style.configure("TNotebook", background=self.bg_color)
        style.configure("TNotebook.Tab", font=("Segoe UI", 11, "bold"), padding=[14, 8])
        style.configure("TFrame", background=self.bg_color)
        style.configure("TLabel", background=self.bg_color, foreground=self.text_color, font=("Segoe UI", 10))
        style.configure("Header.TLabel", font=("Segoe UI", 14, "bold"), foreground="#0a3622")
        style.configure("Action.TButton", font=("Segoe UI", 10, "bold"), padding=6)

    def _build_ui(self):
        # Header Banner
        header_frame = tk.Frame(self.root, bg="#1b4d3e", height=65)
        header_frame.pack(fill=tk.X, side=tk.TOP)
        
        title_label = tk.Label(
            header_frame, 
            text="झारखण्ड सरकार | स्कूली शिक्षा एवं साक्षरता विभाग (कृति देव 010 प्रारूपक)", 
            fg="white", 
            bg="#1b4d3e", 
            font=("Segoe UI", 14, "bold")
        )
        title_label.pack(side=tk.LEFT, padx=20, pady=12)
        
        subtitle_label = tk.Label(
            header_frame,
            text=f"शब्दकोश: {len(self.spellchecker.unicode_words)} शब्द लोड | MS Word LTSC 2024",
            fg="#d1e7dd",
            bg="#1b4d3e",
            font=("Segoe UI", 9)
        )
        subtitle_label.pack(side=tk.RIGHT, padx=20, pady=12)

        # Tab Navigation
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=15, pady=12)

        # Tab 1: Live Typing & IntelliSense
        self.tab_typing = ttk.Frame(notebook)
        notebook.add(self.tab_typing, text="  1. लाइव टाइपिंग एवं स्वतः पूर्ण (IntelliSense)  ")
        self._build_typing_tab()

        # Tab 2: Scan & Fix Word Document
        self.tab_scan = ttk.Frame(notebook)
        notebook.add(self.tab_scan, text="  2. वर्ड दस्तावेज़ वर्तनी एवं व्याकरण जाँच  ")
        self._build_scan_tab()

        # Tab 3: Official Letter Templates
        self.tab_templates = ttk.Frame(notebook)
        notebook.add(self.tab_templates, text="  3. सरकारी प्रारूप एवं पत्र (Templates)  ")
        self._build_templates_tab()

        # Tab 4: Custom Dictionary
        self.tab_dict = ttk.Frame(notebook)
        notebook.add(self.tab_dict, text="  4. विभागीय शब्दावली एवं सेटिंग्स  ")
        self._build_dict_tab()

    # ================= TAB 1: LIVE TYPING & INTELLISENSE =================
    def _build_typing_tab(self):
        main_pane = tk.PanedWindow(self.tab_typing, orient=tk.HORIZONTAL, bg=self.bg_color)
        main_pane.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Left Column: Typing input & Live Hindi Preview
        left_frame = tk.Frame(main_pane, bg=self.bg_color)
        main_pane.add(left_frame, width=540)

        lbl_input = tk.Label(left_frame, text="कृति देव 010 टाइपिंग बॉक्स (Type in Remington Keyboard):", font=("Segoe UI", 10, "bold"), bg=self.bg_color)
        lbl_input.pack(anchor="w", pady=(0, 4))

        self.txt_input = tk.Text(left_frame, height=9, font=("Consolas", 12), wrap=tk.WORD, relief=tk.SOLID, bd=1)
        self.txt_input.pack(fill=tk.BOTH, expand=True, pady=(0, 8))
        self.txt_input.bind("<KeyRelease>", self._on_typing_change)
        self.txt_input.bind("<Tab>", self._on_tab_autocomplete)

        lbl_preview = tk.Label(left_frame, text="हिन्दी रूप (Live Devanagari Preview):", font=("Segoe UI", 10, "bold"), bg=self.bg_color)
        lbl_preview.pack(anchor="w", pady=(4, 4))

        self.txt_preview = tk.Text(left_frame, height=9, font=("Segoe UI", 12), wrap=tk.WORD, bg="#f8f9fa", relief=tk.SOLID, bd=1)
        self.txt_preview.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # Bottom Buttons
        btn_box = tk.Frame(left_frame, bg=self.bg_color)
        btn_box.pack(fill=tk.X)

        btn_insert = tk.Button(btn_box, text="वर्ड में भेजें (Insert to MS Word)", bg="#0056b3", fg="white", font=("Segoe UI", 10, "bold"), padx=12, pady=6, command=self._insert_typing_to_word)
        btn_insert.pack(side=tk.LEFT, padx=(0, 8))

        btn_copy = tk.Button(btn_box, text="कृति देव कॉपी करें (Copy Kruti Text)", bg="#6c757d", fg="white", font=("Segoe UI", 10), padx=12, pady=6, command=self._copy_kruti)
        btn_copy.pack(side=tk.LEFT, padx=(0, 8))

        btn_clear = tk.Button(btn_box, text="साफ करें (Clear)", bg="#f8d7da", fg="#721c24", font=("Segoe UI", 10), padx=10, pady=6, command=self._clear_typing)
        btn_clear.pack(side=tk.RIGHT)

        # Right Column: Autocomplete Suggestions
        right_frame = tk.Frame(main_pane, bg="#ffffff", relief=tk.SOLID, bd=1)
        main_pane.add(right_frame, width=380)

        sug_title = tk.Label(right_frame, text="स्वतः पूर्ण सुझाव (IntelliSense Suggestions)", bg="#e9ecef", font=("Segoe UI", 11, "bold"), pady=8)
        sug_title.pack(fill=tk.X)

        help_lbl = tk.Label(right_frame, text="सुझाव चुनने के लिए डबल क्लिक करें या Tab दबाएं", bg="#ffffff", fg="#6c757d", font=("Segoe UI", 8, "italic"))
        help_lbl.pack(pady=4)

        self.sug_listbox = tk.Listbox(right_frame, font=("Segoe UI", 12), height=14, relief=tk.FLAT, selectbackground="#0056b3", selectforeground="white")
        self.sug_listbox.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        self.sug_listbox.bind("<Double-Button-1>", self._on_suggestion_select)

    def _on_typing_change(self, event=None):
        content = self.txt_input.get("1.0", tk.END).strip()
        if not content:
            self.txt_preview.delete("1.0", tk.END)
            self.sug_listbox.delete(0, tk.END)
            return

        # Update preview in Unicode
        uni = kruti_to_unicode(content)
        self.txt_preview.delete("1.0", tk.END)
        self.txt_preview.insert("1.0", uni)

        # Get current active word prefix
        cursor_pos = self.txt_input.index(tk.INSERT)
        line_text = self.txt_input.get(f"{cursor_pos} linestart", cursor_pos)
        words = line_text.split()
        current_word = words[-1] if words else ""

        if current_word:
            sugs = self.spellchecker.autocomplete(current_word, max_results=8)
            self.current_suggestions = sugs
            self.sug_listbox.delete(0, tk.END)
            for i, s in enumerate(sugs, 1):
                self.sug_listbox.insert(tk.END, f"{i}. {s['unicode']} ({s['kruti']})")
        else:
            self.sug_listbox.delete(0, tk.END)

    def _on_tab_autocomplete(self, event):
        if self.current_suggestions:
            best = self.current_suggestions[0]["kruti"]
            self._replace_current_word_with(best)
            return "break"

    def _on_suggestion_select(self, event):
        sel = self.sug_listbox.curselection()
        if sel and sel[0] < len(self.current_suggestions):
            chosen = self.current_suggestions[sel[0]]["kruti"]
            self._replace_current_word_with(chosen)

    def _replace_current_word_with(self, new_word_kruti: str):
        cursor_pos = self.txt_input.index(tk.INSERT)
        line_start = f"{cursor_pos} linestart"
        line_text = self.txt_input.get(line_start, cursor_pos)
        
        words = line_text.split()
        if words:
            last_word = words[-1]
            start_word_pos = f"{cursor_pos} - {len(last_word)}c"
            self.txt_input.delete(start_word_pos, cursor_pos)
            self.txt_input.insert(start_word_pos, new_word_kruti + " ")
        self._on_typing_change()

    def _insert_typing_to_word(self):
        content = self.txt_input.get("1.0", tk.END).strip()
        if not content:
            messagebox.showinfo("सूचना", "कृपया पहले कुछ टेक्स्ट टाइप करें।")
            return
        success = self.word_connector.insert_kruti_text_at_cursor(content)
        if success:
            messagebox.showinfo("सफल", "टेक्स्ट एमएस वर्ड में सफलतापूर्वक भेज दिया गया!")
        else:
            messagebox.showwarning("चेतावनी", "एमएस वर्ड से कनेक्ट नहीं हो सका। कृपया जांचें कि एमएस वर्ड खुला है।")

    def _copy_kruti(self):
        content = self.txt_input.get("1.0", tk.END).strip()
        self.root.clipboard_clear()
        self.root.clipboard_append(content)
        messagebox.showinfo("कॉपी", "कृति देव टेक्स्ट क्लिपबोर्ड में कॉपी हो गया!")

    def _clear_typing(self):
        self.txt_input.delete("1.0", tk.END)
        self.txt_preview.delete("1.0", tk.END)
        self.sug_listbox.delete(0, tk.END)

    # ================= TAB 2: SCAN & FIX WORD DOC =================
    def _build_scan_tab(self):
        top_bar = tk.Frame(self.tab_scan, bg=self.bg_color)
        top_bar.pack(fill=tk.X, padx=10, pady=10)

        btn_scan = tk.Button(top_bar, text="🔍 सक्रिय वर्ड दस्तावेज़ स्कैन करें (Scan Active Word Document)", bg="#0056b3", fg="white", font=("Segoe UI", 11, "bold"), padx=15, pady=8, command=self._scan_active_word)
        btn_scan.pack(side=tk.LEFT)

        self.lbl_scan_status = tk.Label(top_bar, text="वर्ड फाइल की स्थिति: जांच के लिए बटन दबाएं", font=("Segoe UI", 10, "italic"), bg=self.bg_color, fg="#495057")
        self.lbl_scan_status.pack(side=tk.LEFT, padx=15)

        # Results tree / list
        tree_frame = tk.Frame(self.tab_scan, bg=self.bg_color)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        cols = ("type", "error_kruti", "error_uni", "suggestion", "context")
        self.tree_errors = ttk.Treeview(tree_frame, columns=cols, show="headings", height=14)
        self.tree_errors.heading("type", text="प्रकार (Type)")
        self.tree_errors.heading("error_kruti", text="गलत शब्द (Kruti)")
        self.tree_errors.heading("error_uni", text="हिन्दी रूप (Unicode)")
        self.tree_errors.heading("suggestion", text="सुझाव (Suggested Fix)")
        self.tree_errors.heading("context", text="संदर्भ / नियम (Context & Rule)")

        self.tree_errors.column("type", width=120)
        self.tree_errors.column("error_kruti", width=120)
        self.tree_errors.column("error_uni", width=130)
        self.tree_errors.column("suggestion", width=180)
        self.tree_errors.column("context", width=360)

        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree_errors.yview)
        self.tree_errors.configure(yscroll=scrollbar.set)
        
        self.tree_errors.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Bottom Action Bar
        bottom_bar = tk.Frame(self.tab_scan, bg=self.bg_color)
        bottom_bar.pack(fill=tk.X, padx=10, pady=(0, 10))

        btn_fix = tk.Button(bottom_bar, text="✓ चयनित त्रुटि वर्ड में ठीक करें (Auto-Fix Selected in Word)", bg="#28a745", fg="white", font=("Segoe UI", 10, "bold"), padx=12, pady=6, command=self._fix_selected_error)
        btn_fix.pack(side=tk.LEFT)

    def _scan_active_word(self):
        text = self.word_connector.get_active_document_text()
        if not text:
            messagebox.showwarning("वर्ड कनेक्ट त्रुटि", "कोई सक्रिय एमएस वर्ड दस्तावेज़ नहीं मिला। कृपया वर्ड में फाइल खोलें और पुनः प्रयास करें।")
            return

        report = self.spellchecker.verify_document_text(text)
        self.tree_errors.delete(*self.tree_errors.get_children())

        # Populate grammar issues
        for g in report["grammar_issues"]:
            self.tree_errors.insert("", tk.END, values=(
                "व्याकरण (Grammar)",
                unicode_to_kruti(g["original_word"]),
                g["original_word"],
                f"{g['suggestion']} ({g['suggestion_kruti']})",
                g["message"]
            ))

        # Populate spelling errors
        for s in report["spelling_errors"]:
            sug_text = ", ".join([f"{c['unicode']} ({c['kruti']})" for c in s["suggestions"][:2]])
            self.tree_errors.insert("", tk.END, values=(
                "वर्तनी (Spelling)",
                s["word_kruti"],
                s["word_unicode"],
                sug_text if sug_text else "कोई सुझाव नहीं",
                "शब्दकोश में नहीं मिला"
            ))

        total = report["total_errors"]
        self.lbl_scan_status.config(text=f"स्कैन पूर्ण: कुल {total} त्रुटियाँ / सुझाव मिले।", fg="#d9534f" if total > 0 else "#28a745")

    def _fix_selected_error(self):
        sel = self.tree_errors.selection()
        if not sel:
            messagebox.showinfo("चयन करें", "कृपया सूची से एक त्रुटि चुनें जिसे ठीक करना है।")
            return
            
        vals = self.tree_errors.item(sel[0], "values")
        error_kruti = vals[1]
        sug_raw = vals[3]
        
        # Extract kruti suggestion from parentheses
        if "(" in sug_raw and ")" in sug_raw:
            sug_kruti = sug_raw.split("(")[1].split(")")[0].strip()
            success = self.word_connector.replace_word_in_active_doc(error_kruti, sug_kruti)
            if success:
                messagebox.showinfo("सफल", f"वर्ड में '{error_kruti}' को बदलकर '{sug_kruti}' कर दिया गया!")
                self.tree_errors.delete(sel[0])
            else:
                messagebox.showerror("त्रुटि", "वर्ड में सुधार लागू नहीं हो सका।")

    # ================= TAB 3: OFFICIAL TEMPLATES =================
    def _build_templates_tab(self):
        paned = tk.PanedWindow(self.tab_templates, orient=tk.HORIZONTAL, bg=self.bg_color)
        paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Left list of templates
        left_box = tk.Frame(paned, bg="#ffffff", relief=tk.SOLID, bd=1)
        paned.add(left_box, width=320)

        t_lbl = tk.Label(left_box, text="सरकारी पत्र एवं प्रारूप सूची", bg="#e9ecef", font=("Segoe UI", 11, "bold"), pady=8)
        t_lbl.pack(fill=tk.X)

        self.tpl_listbox = tk.Listbox(left_box, font=("Segoe UI", 11), height=16, relief=tk.FLAT, selectbackground="#0056b3")
        self.tpl_listbox.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        for t in self.templates:
            self.tpl_listbox.insert(tk.END, t["title"])
        self.tpl_listbox.bind("<<ListboxSelect>>", self._on_template_select)

        # Right preview & insert
        right_box = tk.Frame(paned, bg=self.bg_color)
        paned.add(right_box, width=600)

        top_r = tk.Frame(right_box, bg=self.bg_color)
        top_r.pack(fill=tk.X, pady=(0, 6))

        self.lbl_tpl_name = tk.Label(top_r, text="प्रारूप पूर्वावलोकन (Preview):", font=("Segoe UI", 11, "bold"), bg=self.bg_color)
        self.lbl_tpl_name.pack(side=tk.LEFT)

        btn_insert_tpl = tk.Button(top_r, text="📄 वर्ड में प्रारूप खोलें (Insert Template in MS Word)", bg="#28a745", fg="white", font=("Segoe UI", 10, "bold"), padx=12, pady=5, command=self._insert_active_template)
        btn_insert_tpl.pack(side=tk.RIGHT)

        self.txt_tpl_preview = tk.Text(right_box, font=("Segoe UI", 11), wrap=tk.WORD, bg="#ffffff", relief=tk.SOLID, bd=1)
        self.txt_tpl_preview.pack(fill=tk.BOTH, expand=True)

    def _on_template_select(self, event):
        sel = self.tpl_listbox.curselection()
        if sel:
            tpl = self.templates[sel[0]]
            self.txt_tpl_preview.delete("1.0", tk.END)
            self.txt_tpl_preview.insert("1.0", tpl["content_unicode"])

    def _insert_active_template(self):
        sel = self.tpl_listbox.curselection()
        if not sel:
            messagebox.showinfo("चयन करें", "कृपया पहले एक प्रारूप चुनें।")
            return
        tpl = self.templates[sel[0]]
        success = self.word_connector.insert_template(tpl["content_kruti"])
        if success:
            messagebox.showinfo("सफल", f"प्रारूप '{tpl['title']}' एमएस वर्ड में सफलतापूर्वक सम्मिलित कर दिया गया!")
        else:
            messagebox.showwarning("चेतावनी", "एमएस वर्ड से कनेक्ट नहीं हो सका।")

    # ================= TAB 4: CUSTOM DICTIONARY =================
    def _build_dict_tab(self):
        frame = tk.Frame(self.tab_dict, bg=self.bg_color, padx=20, pady=20)
        frame.pack(fill=tk.BOTH, expand=True)

        info = tk.Label(frame, text="विभागीय नया शब्द अथवा पदाधिकारी का नाम जोड़ें (Add Custom Word):", font=("Segoe UI", 11, "bold"), bg=self.bg_color)
        info.pack(anchor="w", pady=(0, 10))

        add_box = tk.Frame(frame, bg=self.bg_color)
        add_box.pack(fill=tk.X, pady=(0, 15))

        self.entry_custom_word = ttk.Entry(add_box, font=("Segoe UI", 12), width=35)
        self.entry_custom_word.pack(side=tk.LEFT, padx=(0, 10))

        btn_add = tk.Button(add_box, text="+ शब्दकोश में जोड़ें (Add Word)", bg="#0056b3", fg="white", font=("Segoe UI", 10, "bold"), padx=12, pady=5, command=self._add_custom_word)
        btn_add.pack(side=tk.LEFT)

        stats_box = tk.LabelFrame(frame, text=" शब्दकोश सांख्यिकी (Dictionary Statistics) ", font=("Segoe UI", 10, "bold"), bg=self.bg_color, padx=15, pady=15)
        stats_box.pack(fill=tk.BOTH, expand=True)

        lbl1 = tk.Label(stats_box, text=f"• कुल लोड किए गए हिन्दी शब्द: {len(self.spellchecker.unicode_words)}", font=("Segoe UI", 10), bg=self.bg_color)
        lbl1.pack(anchor="w", pady=4)

        lbl2 = tk.Label(stats_box, text="• स्रोत: 2024 डेटा संचिका (सरायकेला-खरसावाँ शिक्षा विभाग) + केन्द्रीय राजभाषा शब्दावली", font=("Segoe UI", 10), bg=self.bg_color)
        lbl2.pack(anchor="w", pady=4)

        lbl3 = tk.Label(stats_box, text="• समर्थित फॉन्ट: Kruti Dev 010 (Remington Keyboard Layout)", font=("Segoe UI", 10), bg=self.bg_color)
        lbl3.pack(anchor="w", pady=4)

    def _add_custom_word(self):
        w = self.entry_custom_word.get().strip()
        if not w:
            return
        self.spellchecker.add_custom_word(w)
        self.entry_custom_word.delete(0, tk.END)
        messagebox.showinfo("सफल", f"शब्द '{w}' को सफलतापूर्वक शब्दकोश में जोड़ दिया गया है!")


def launch_app():
    root = tk.Tk()
    app = KrutiDevAssistantGUI(root)
    root.mainloop()


if __name__ == "__main__":
    launch_app()

