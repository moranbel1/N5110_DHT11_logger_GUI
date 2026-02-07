#!/usr/bin/env python3
"""
English Learning App for Kids - אפליקציה ללימוד אנגלית לילדים
Target audience: 4th grade students who know the English alphabet and a few words.
Features: Vocabulary learning (flashcards), translation quizzes, progress tracking.
"""

import tkinter as tk
from tkinter import messagebox
import json
import random
import os

# ──────────────────────────── Color Palette ────────────────────────────

COLORS = {
    "bg": "#E8F4FD",
    "card": "#FFFFFF",
    "primary": "#4A90D9",
    "success": "#27AE60",
    "warning": "#F39C12",
    "danger": "#E74C3C",
    "purple": "#8E44AD",
    "pink": "#E91E8A",
    "teal": "#1ABC9C",
    "text": "#2C3E50",
    "light_text": "#7F8C8D",
    "white": "#FFFFFF",
    "star": "#F1C40F",
}

CATEGORY_COLORS = [
    "#E67E22",  # Animals - orange
    "#9B59B6",  # Colors - purple
    "#E74C3C",  # Food - red
    "#3498DB",  # Family - blue
    "#2ECC71",  # School - green
    "#1ABC9C",  # Body - teal
    "#F39C12",  # Clothes - yellow
    "#27AE60",  # Nature - dark green
]

# ──────────────────────────── Vocabulary ────────────────────────────

VOCABULARY = {
    "🐾 חיות": {
        "dog": "כלב",
        "cat": "חתול",
        "bird": "ציפור",
        "fish": "דג",
        "horse": "סוס",
        "cow": "פרה",
        "rabbit": "ארנב",
        "lion": "אריה",
        "elephant": "פיל",
        "monkey": "קוף",
        "bear": "דוב",
        "snake": "נחש",
    },
    "🎨 צבעים": {
        "red": "אדום",
        "blue": "כחול",
        "green": "ירוק",
        "yellow": "צהוב",
        "orange": "כתום",
        "purple": "סגול",
        "pink": "ורוד",
        "white": "לבן",
        "black": "שחור",
        "brown": "חום",
    },
    "🍎 אוכל": {
        "apple": "תפוח",
        "banana": "בננה",
        "bread": "לחם",
        "water": "מים",
        "milk": "חלב",
        "egg": "ביצה",
        "rice": "אורז",
        "chicken": "עוף",
        "pizza": "פיצה",
        "cake": "עוגה",
        "cheese": "גבינה",
        "salad": "סלט",
    },
    "👨‍👩‍👧‍👦 משפחה": {
        "mother": "אמא",
        "father": "אבא",
        "brother": "אח",
        "sister": "אחות",
        "baby": "תינוק",
        "grandmother": "סבתא",
        "grandfather": "סבא",
        "uncle": "דוד",
        "aunt": "דודה",
        "friend": "חבר",
        "boy": "ילד",
        "girl": "ילדה",
    },
    "📖 בית ספר": {
        "book": "ספר",
        "pencil": "עיפרון",
        "teacher": "מורה",
        "student": "תלמיד",
        "desk": "שולחן",
        "chair": "כיסא",
        "board": "לוח",
        "bag": "תיק",
        "ruler": "סרגל",
        "eraser": "מחק",
        "pen": "עט",
        "school": "בית ספר",
    },
    "🏃 גוף": {
        "head": "ראש",
        "hand": "יד",
        "eye": "עין",
        "ear": "אוזן",
        "nose": "אף",
        "mouth": "פה",
        "leg": "רגל",
        "foot": "כף רגל",
        "hair": "שיער",
        "finger": "אצבע",
        "face": "פנים",
        "teeth": "שיניים",
    },
    "👕 בגדים": {
        "shirt": "חולצה",
        "pants": "מכנסיים",
        "shoes": "נעליים",
        "hat": "כובע",
        "dress": "שמלה",
        "socks": "גרביים",
        "jacket": "ז'קט",
        "skirt": "חצאית",
        "scarf": "צעיף",
        "gloves": "כפפות",
        "coat": "מעיל",
        "belt": "חגורה",
    },
    "🌳 טבע": {
        "sun": "שמש",
        "moon": "ירח",
        "tree": "עץ",
        "flower": "פרח",
        "rain": "גשם",
        "snow": "שלג",
        "cloud": "ענן",
        "star": "כוכב",
        "river": "נהר",
        "mountain": "הר",
        "sea": "ים",
        "wind": "רוח",
    },
}

ENCOURAGEMENT_CORRECT = [
    "!מצוין! כל הכבוד 🎉",
    "!נכון! אתה כוכב 🌟",
    "!יופי! המשך כך 💪",
    "!בול! מדהים 🏆",
    "!נהדר! אלוף 👑",
]

ENCOURAGEMENT_WRONG = [
    "לא נורא, ננסה שוב! 💪",
    "כמעט! בפעם הבאה 😊",
    "!קרוב! אל תוותר 🌟",
]

QUIZ_RESULTS_MESSAGES = {
    "perfect": "🏆 מושלם! אתה גאון אנגלית!",
    "great": "🌟 מעולה! עבודה נהדרת!",
    "good": "👏 יפה מאוד! ממשיכים להתקדם!",
    "ok": "💪 !לא רע! עוד קצת תרגול",
    "try_again": "😊 בוא ננסה שוב! תרגול עושה מושלם!",
}

# ──────────────────────────── Application ────────────────────────────


class EnglishLearningApp:
    """Main application class for the English learning GUI."""

    WINDOW_WIDTH = 900
    WINDOW_HEIGHT = 680

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Let's Learn English! - בואו נלמד אנגלית!")
        self.root.geometry(f"{self.WINDOW_WIDTH}x{self.WINDOW_HEIGHT}")
        self.root.configure(bg=COLORS["bg"])
        self.root.resizable(False, False)

        self.progress_file = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "progress.json"
        )
        self.progress = self._load_progress()

        # State for learning mode
        self._learn_index = 0
        self._learn_words = []
        self._learn_category = ""
        self._translation_visible = False

        # State for quiz mode
        self._quiz_words = []
        self._quiz_index = 0
        self._quiz_score = 0
        self._quiz_category = ""
        self._quiz_answered = False

        self._show_main_menu()
        self.root.mainloop()

    # ───────────────── Helpers ─────────────────

    def _clear(self):
        """Remove all widgets from the root window."""
        for w in self.root.winfo_children():
            w.destroy()

    def _make_button(self, parent, text, command, bg, fg="#FFFFFF",
                     font_size=16, width=None, height=None, padx=20, pady=10):
        """Create a styled button with hover effect."""
        font = ("Arial", font_size, "bold")
        btn = tk.Button(
            parent,
            text=text,
            command=command,
            bg=bg,
            fg=fg,
            font=font,
            relief="flat",
            cursor="hand2",
            activebackground=self._adjust_color(bg, -30),
            activeforeground=fg,
            padx=padx,
            pady=pady,
            bd=0,
            highlightthickness=0,
        )
        if width:
            btn.configure(width=width)
        if height:
            btn.configure(height=height)

        # Hover effects
        normal_bg = bg
        hover_bg = self._adjust_color(bg, -25)
        btn.bind("<Enter>", lambda e: btn.configure(bg=hover_bg))
        btn.bind("<Leave>", lambda e: btn.configure(bg=normal_bg))
        return btn

    @staticmethod
    def _adjust_color(hex_color, amount):
        """Lighten (positive) or darken (negative) a hex color."""
        hex_color = hex_color.lstrip("#")
        r = max(0, min(255, int(hex_color[0:2], 16) + amount))
        g = max(0, min(255, int(hex_color[2:4], 16) + amount))
        b = max(0, min(255, int(hex_color[4:6], 16) + amount))
        return f"#{r:02x}{g:02x}{b:02x}"

    def _create_header(self, title, subtitle=None):
        """Create a consistent header with a back button."""
        header = tk.Frame(self.root, bg=COLORS["bg"])
        header.pack(fill="x", padx=20, pady=(15, 5))

        back_btn = self._make_button(
            header, "⬅ חזרה", self._show_main_menu,
            COLORS["light_text"], font_size=11, padx=10, pady=5,
        )
        back_btn.pack(side="right")

        tk.Label(
            header, text=title, font=("Arial", 22, "bold"),
            bg=COLORS["bg"], fg=COLORS["text"],
        ).pack(side="left")

        if subtitle:
            tk.Label(
                header, text=subtitle, font=("Arial", 13),
                bg=COLORS["bg"], fg=COLORS["light_text"],
            ).pack(side="left", padx=(15, 0))

        # Divider line
        tk.Frame(self.root, bg="#D5DBDB", height=2).pack(fill="x", padx=20, pady=(5, 10))

    def _progress_bar(self, parent, value, max_val, width=200, height=18,
                      bar_color=COLORS["success"], bg_color="#DFE6E9"):
        """Draw a progress bar on a Canvas."""
        canvas = tk.Canvas(parent, width=width, height=height,
                           bg=bg_color, highlightthickness=0, bd=0)
        if max_val > 0:
            fill_width = int((value / max_val) * width)
            if fill_width > 0:
                canvas.create_rectangle(0, 0, fill_width, height, fill=bar_color, outline="")
        canvas.create_rectangle(0, 0, width, height, outline="#BDC3C7", width=1)
        return canvas

    # ───────────────── Persistence ─────────────────

    def _load_progress(self):
        """Load progress from JSON file or return defaults."""
        if os.path.exists(self.progress_file):
            try:
                with open(self.progress_file, "r", encoding="utf-8") as f:
                    return json.load(f)
            except (json.JSONDecodeError, IOError):
                pass
        return {}

    def _save_progress(self):
        """Save progress to JSON file."""
        try:
            with open(self.progress_file, "w", encoding="utf-8") as f:
                json.dump(self.progress, f, ensure_ascii=False, indent=2)
        except IOError:
            pass

    def _get_cat_progress(self, category):
        """Return progress dict for a category, creating if needed."""
        if category not in self.progress:
            self.progress[category] = {
                "learned_words": [],
                "quiz_scores": [],
                "best_score": 0,
            }
        return self.progress[category]

    # ───────────────── Main Menu ─────────────────

    def _show_main_menu(self):
        self._clear()
        self.root.configure(bg=COLORS["bg"])

        # Spacer
        tk.Frame(self.root, bg=COLORS["bg"], height=40).pack()

        # Title
        tk.Label(
            self.root, text="🎓 Let's Learn English!",
            font=("Arial", 38, "bold"), bg=COLORS["bg"], fg=COLORS["primary"],
        ).pack()

        tk.Label(
            self.root, text="!בואו נלמד אנגלית",
            font=("Arial", 20), bg=COLORS["bg"], fg=COLORS["text"],
        ).pack(pady=(5, 30))

        # Menu buttons
        btn_frame = tk.Frame(self.root, bg=COLORS["bg"])
        btn_frame.pack(pady=10)

        buttons = [
            ("📚  לימוד מילים", COLORS["primary"], lambda: self._show_categories("learn")),
            ("🎯  חידון תרגום", COLORS["success"], lambda: self._show_categories("quiz")),
            ("📊  ההתקדמות שלי", COLORS["purple"], self._show_progress),
        ]

        for text, color, cmd in buttons:
            btn = self._make_button(btn_frame, text, cmd, color, font_size=20,
                                    padx=40, pady=15)
            btn.pack(pady=10, ipadx=30)

        # Stats summary
        total_words = sum(len(v) for v in VOCABULARY.values())
        learned = self._total_learned_words()
        if learned > 0:
            tk.Label(
                self.root,
                text=f"🌟 למדת {learned} מתוך {total_words} מילים",
                font=("Arial", 14), bg=COLORS["bg"], fg=COLORS["light_text"],
            ).pack(pady=(30, 0))

    # ───────────────── Category Selection ─────────────────

    def _show_categories(self, mode):
        self._clear()

        mode_text = "לימוד מילים" if mode == "learn" else "חידון תרגום"
        self._create_header(f"📂 בחר נושא", subtitle=mode_text)

        grid_frame = tk.Frame(self.root, bg=COLORS["bg"])
        grid_frame.pack(expand=True, pady=10)

        categories = list(VOCABULARY.keys())
        for i, cat in enumerate(categories):
            row, col = divmod(i, 2)
            color = CATEGORY_COLORS[i % len(CATEGORY_COLORS)]

            word_count = len(VOCABULARY[cat])
            progress = self._get_cat_progress(cat)
            learned_count = len(progress.get("learned_words", []))

            label = f"{cat}\n({word_count} מילים)"
            if mode == "learn" and learned_count > 0:
                label = f"{cat}\n✓ {learned_count}/{word_count}"

            if mode == "learn":
                cmd = lambda c=cat: self._start_learning(c)
            else:
                cmd = lambda c=cat: self._start_quiz(c)

            btn = self._make_button(grid_frame, label, cmd, color,
                                    font_size=14, padx=15, pady=12)
            btn.configure(width=18, height=3)
            btn.grid(row=row, column=col, padx=12, pady=10)

        # "All topics" button for quiz mode
        if mode == "quiz":
            all_btn = self._make_button(
                self.root, "🌈 כל הנושאים ביחד",
                lambda: self._start_quiz("__all__"),
                COLORS["pink"], font_size=15, padx=25, pady=10,
            )
            all_btn.pack(pady=(5, 10))

    # ───────────────── Learning Mode (Flashcards) ─────────────────

    def _start_learning(self, category):
        self._learn_category = category
        words = list(VOCABULARY[category].items())
        self._learn_words = words
        self._learn_index = 0
        self._translation_visible = False
        self._show_flashcard()

    def _show_flashcard(self):
        self._clear()

        cat = self._learn_category
        words = self._learn_words
        idx = self._learn_index
        total = len(words)
        english, hebrew = words[idx]

        # Header
        header = tk.Frame(self.root, bg=COLORS["bg"])
        header.pack(fill="x", padx=20, pady=(15, 5))

        back_btn = self._make_button(
            header, "⬅ חזרה", lambda: self._show_categories("learn"),
            COLORS["light_text"], font_size=11, padx=10, pady=5,
        )
        back_btn.pack(side="right")

        tk.Label(
            header, text=cat, font=("Arial", 18, "bold"),
            bg=COLORS["bg"], fg=COLORS["text"],
        ).pack(side="left")

        tk.Label(
            header, text=f"({idx + 1}/{total})",
            font=("Arial", 14), bg=COLORS["bg"], fg=COLORS["light_text"],
        ).pack(side="left", padx=(10, 0))

        tk.Frame(self.root, bg="#D5DBDB", height=2).pack(fill="x", padx=20, pady=(5, 10))

        # Flashcard
        card = tk.Frame(self.root, bg=COLORS["card"], bd=0,
                        highlightbackground="#D5DBDB", highlightthickness=2)
        card.pack(expand=True, fill="both", padx=60, pady=(10, 15))

        # English word (always visible)
        tk.Label(
            card, text=english, font=("Arial", 52, "bold"),
            bg=COLORS["card"], fg=COLORS["primary"],
        ).pack(pady=(40, 10))

        # Phonetic hint - show first letter
        hint = f'"{english[0].upper()}" :האות הראשונה'
        tk.Label(
            card, text=hint, font=("Arial", 13),
            bg=COLORS["card"], fg=COLORS["light_text"],
        ).pack()

        if self._translation_visible:
            # Show translation
            tk.Label(
                card, text="⬇", font=("Arial", 20),
                bg=COLORS["card"], fg=COLORS["light_text"],
            ).pack(pady=(10, 5))

            tk.Label(
                card, text=hebrew, font=("Arial", 38, "bold"),
                bg=COLORS["card"], fg=COLORS["success"],
            ).pack(pady=(0, 10))

            # Mark as learned
            progress = self._get_cat_progress(cat)
            if english not in progress["learned_words"]:
                progress["learned_words"].append(english)
                self._save_progress()
        else:
            # Show "reveal" button
            tk.Frame(card, bg=COLORS["card"], height=15).pack()
            reveal_btn = self._make_button(
                card, "👆 לחץ לראות תרגום", self._reveal_translation,
                COLORS["warning"], font_size=16, padx=25, pady=10,
            )
            reveal_btn.pack(pady=(10, 20))

        # Navigation buttons
        nav_frame = tk.Frame(self.root, bg=COLORS["bg"])
        nav_frame.pack(fill="x", padx=60, pady=(0, 20))

        if idx > 0:
            prev_btn = self._make_button(
                nav_frame, "⬅ הקודם", self._prev_card,
                COLORS["primary"], font_size=14, padx=20, pady=8,
            )
            prev_btn.pack(side="right")

        if idx < total - 1:
            next_btn = self._make_button(
                nav_frame, "הבא ➡", self._next_card,
                COLORS["success"], font_size=14, padx=20, pady=8,
            )
            next_btn.pack(side="left")
        else:
            done_btn = self._make_button(
                nav_frame, "🎉 סיימתי!", lambda: self._show_categories("learn"),
                COLORS["pink"], font_size=14, padx=20, pady=8,
            )
            done_btn.pack(side="left")

    def _reveal_translation(self):
        self._translation_visible = True
        self._show_flashcard()

    def _next_card(self):
        self._learn_index += 1
        self._translation_visible = False
        self._show_flashcard()

    def _prev_card(self):
        self._learn_index -= 1
        self._translation_visible = False
        self._show_flashcard()

    # ───────────────── Quiz Mode ─────────────────

    def _start_quiz(self, category):
        self._quiz_category = category
        self._quiz_score = 0
        self._quiz_index = 0
        self._quiz_answered = False

        if category == "__all__":
            all_words = []
            for cat, words in VOCABULARY.items():
                all_words.extend([(eng, heb, cat) for eng, heb in words.items()])
            random.shuffle(all_words)
            self._quiz_words = all_words[:10]
        else:
            words = [(eng, heb, category) for eng, heb in VOCABULARY[category].items()]
            random.shuffle(words)
            self._quiz_words = words[:10]

        self._show_quiz_question()

    def _show_quiz_question(self):
        self._clear()
        self._quiz_answered = False

        idx = self._quiz_index
        total = len(self._quiz_words)
        english, hebrew, cat = self._quiz_words[idx]

        # Header
        header = tk.Frame(self.root, bg=COLORS["bg"])
        header.pack(fill="x", padx=20, pady=(15, 5))

        cat_display = "כל הנושאים" if self._quiz_category == "__all__" else self._quiz_category
        tk.Label(
            header, text=f"🎯 חידון - {cat_display}",
            font=("Arial", 18, "bold"), bg=COLORS["bg"], fg=COLORS["text"],
        ).pack(side="left")

        score_text = f"✅ {self._quiz_score}/{idx}"
        tk.Label(
            header, text=score_text, font=("Arial", 14, "bold"),
            bg=COLORS["bg"], fg=COLORS["success"],
        ).pack(side="right")

        tk.Frame(self.root, bg="#D5DBDB", height=2).pack(fill="x", padx=20, pady=(5, 10))

        # Question counter
        tk.Label(
            self.root, text=f"שאלה {idx + 1} מתוך {total}",
            font=("Arial", 14), bg=COLORS["bg"], fg=COLORS["light_text"],
        ).pack(pady=(5, 10))

        # Question card
        card = tk.Frame(self.root, bg=COLORS["card"], bd=0,
                        highlightbackground="#D5DBDB", highlightthickness=2)
        card.pack(fill="x", padx=60, pady=(5, 10))

        tk.Label(
            card, text="?מה התרגום לאנגלית של", font=("Arial", 15),
            bg=COLORS["card"], fg=COLORS["light_text"],
        ).pack(pady=(20, 5))

        tk.Label(
            card, text=hebrew, font=("Arial", 42, "bold"),
            bg=COLORS["card"], fg=COLORS["text"],
        ).pack(pady=(5, 20))

        # Generate answer options
        correct = english
        wrong_options = self._get_wrong_options(correct, cat)
        options = [correct] + wrong_options
        random.shuffle(options)

        # Answer buttons
        answers_frame = tk.Frame(self.root, bg=COLORS["bg"])
        answers_frame.pack(pady=10)

        option_colors = ["#3498DB", "#E67E22", "#9B59B6", "#1ABC9C"]
        for i, option in enumerate(options):
            btn = self._make_button(
                answers_frame, option,
                lambda opt=option: self._check_answer(opt, correct),
                option_colors[i], font_size=18, padx=30, pady=12,
            )
            btn.configure(width=15)
            row, col = divmod(i, 2)
            btn.grid(row=row, column=col, padx=10, pady=8)

    def _get_wrong_options(self, correct, category):
        """Get 3 wrong English options for the quiz."""
        # Gather candidates from the same category first
        candidates = set()
        if category != "__all__" and category in VOCABULARY:
            candidates.update(VOCABULARY[category].keys())
        # Add from all categories if needed
        for cat_words in VOCABULARY.values():
            candidates.update(cat_words.keys())
        candidates.discard(correct)
        candidates = list(candidates)
        random.shuffle(candidates)
        return candidates[:3]

    def _check_answer(self, selected, correct):
        if self._quiz_answered:
            return
        self._quiz_answered = True

        is_correct = selected == correct
        if is_correct:
            self._quiz_score += 1

        self._show_answer_feedback(is_correct, correct)

    def _show_answer_feedback(self, is_correct, correct_answer):
        """Show feedback overlay after answering."""
        overlay = tk.Frame(self.root, bg=COLORS["bg"])
        overlay.pack(fill="x", padx=60, pady=(5, 0))

        if is_correct:
            msg = random.choice(ENCOURAGEMENT_CORRECT)
            color = COLORS["success"]
            symbol = "✓"
        else:
            msg = random.choice(ENCOURAGEMENT_WRONG)
            color = COLORS["danger"]
            symbol = "✗"
            # Show the correct answer
            msg += f"\n\"{correct_answer}\" :התשובה הנכונה"

        tk.Label(
            overlay, text=f"{symbol} {msg}",
            font=("Arial", 16, "bold"), bg=COLORS["bg"], fg=color,
            justify="center",
        ).pack(pady=5)

        idx = self._quiz_index
        total = len(self._quiz_words)

        if idx < total - 1:
            next_btn = self._make_button(
                overlay, "➡ שאלה הבאה", self._next_question,
                COLORS["primary"], font_size=14, padx=20, pady=8,
            )
            next_btn.pack(pady=5)
        else:
            results_btn = self._make_button(
                overlay, "📊 לתוצאות", self._show_quiz_results,
                COLORS["purple"], font_size=14, padx=20, pady=8,
            )
            results_btn.pack(pady=5)

    def _next_question(self):
        self._quiz_index += 1
        self._show_quiz_question()

    def _show_quiz_results(self):
        """Display quiz results with score and encouragement."""
        self._clear()

        score = self._quiz_score
        total = len(self._quiz_words)
        pct = (score / total) * 100 if total > 0 else 0

        # Save progress
        cat = self._quiz_category
        if cat != "__all__":
            prog = self._get_cat_progress(cat)
            prog["quiz_scores"].append(score)
            prog["best_score"] = max(prog.get("best_score", 0), score)
            self._save_progress()

        # Determine message and stars
        if pct == 100:
            msg = QUIZ_RESULTS_MESSAGES["perfect"]
            stars = "⭐⭐⭐"
            msg_color = COLORS["star"]
        elif pct >= 80:
            msg = QUIZ_RESULTS_MESSAGES["great"]
            stars = "⭐⭐"
            msg_color = COLORS["success"]
        elif pct >= 60:
            msg = QUIZ_RESULTS_MESSAGES["good"]
            stars = "⭐"
            msg_color = COLORS["primary"]
        elif pct >= 40:
            msg = QUIZ_RESULTS_MESSAGES["ok"]
            stars = ""
            msg_color = COLORS["warning"]
        else:
            msg = QUIZ_RESULTS_MESSAGES["try_again"]
            stars = ""
            msg_color = COLORS["danger"]

        tk.Frame(self.root, bg=COLORS["bg"], height=30).pack()

        # Score display
        tk.Label(
            self.root, text="📊 תוצאות החידון",
            font=("Arial", 28, "bold"), bg=COLORS["bg"], fg=COLORS["text"],
        ).pack(pady=(10, 20))

        # Score card
        card = tk.Frame(self.root, bg=COLORS["card"], bd=0,
                        highlightbackground="#D5DBDB", highlightthickness=2)
        card.pack(padx=80, pady=10, fill="x")

        tk.Label(
            card, text=f"{score}/{total}",
            font=("Arial", 56, "bold"), bg=COLORS["card"], fg=msg_color,
        ).pack(pady=(25, 5))

        if stars:
            tk.Label(
                card, text=stars, font=("Arial", 36),
                bg=COLORS["card"],
            ).pack(pady=(0, 5))

        tk.Label(
            card, text=msg, font=("Arial", 18, "bold"),
            bg=COLORS["card"], fg=msg_color, wraplength=500, justify="center",
        ).pack(pady=(5, 25))

        # Action buttons
        btn_frame = tk.Frame(self.root, bg=COLORS["bg"])
        btn_frame.pack(pady=25)

        retry_btn = self._make_button(
            btn_frame, "🔄 נסה שוב",
            lambda: self._start_quiz(self._quiz_category),
            COLORS["success"], font_size=15, padx=20, pady=10,
        )
        retry_btn.pack(side="left", padx=10)

        cats_btn = self._make_button(
            btn_frame, "📂 נושא אחר",
            lambda: self._show_categories("quiz"),
            COLORS["primary"], font_size=15, padx=20, pady=10,
        )
        cats_btn.pack(side="left", padx=10)

        home_btn = self._make_button(
            btn_frame, "🏠 תפריט ראשי",
            self._show_main_menu,
            COLORS["purple"], font_size=15, padx=20, pady=10,
        )
        home_btn.pack(side="left", padx=10)

    # ───────────────── Progress Screen ─────────────────

    def _total_learned_words(self):
        """Count unique learned words across all categories."""
        total = 0
        for cat in VOCABULARY:
            prog = self.progress.get(cat, {})
            total += len(prog.get("learned_words", []))
        return total

    def _show_progress(self):
        self._clear()
        self._create_header("📊 ההתקדמות שלי")

        # Scrollable area
        container = tk.Frame(self.root, bg=COLORS["bg"])
        container.pack(fill="both", expand=True, padx=20, pady=5)

        canvas = tk.Canvas(container, bg=COLORS["bg"], highlightthickness=0)
        scrollbar = tk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scrollable = tk.Frame(canvas, bg=COLORS["bg"])

        scrollable.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        canvas.create_window((0, 0), window=scrollable, anchor="nw",
                             width=self.WINDOW_WIDTH - 70)
        canvas.configure(yscrollcommand=scrollbar.set)

        # Mouse wheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        canvas.bind_all("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))
        canvas.bind_all("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))

        scrollbar.pack(side="left", fill="y")
        canvas.pack(side="right", fill="both", expand=True)

        # Overall stats card
        total_words = sum(len(v) for v in VOCABULARY.values())
        learned = self._total_learned_words()
        all_scores = []
        for cat in VOCABULARY:
            prog = self.progress.get(cat, {})
            all_scores.extend(prog.get("quiz_scores", []))

        overall_card = tk.Frame(scrollable, bg=COLORS["card"], bd=0,
                                highlightbackground="#D5DBDB", highlightthickness=2)
        overall_card.pack(fill="x", padx=10, pady=10)

        tk.Label(
            overall_card, text="📈 סיכום כללי",
            font=("Arial", 18, "bold"), bg=COLORS["card"], fg=COLORS["text"],
        ).pack(anchor="e", padx=15, pady=(10, 5))

        stats_frame = tk.Frame(overall_card, bg=COLORS["card"])
        stats_frame.pack(fill="x", padx=15, pady=(0, 10))

        # Learned words stat
        stat1 = tk.Frame(stats_frame, bg=COLORS["card"])
        stat1.pack(side="right", expand=True, padx=10)
        tk.Label(stat1, text=f"{learned}/{total_words}", font=("Arial", 24, "bold"),
                 bg=COLORS["card"], fg=COLORS["primary"]).pack()
        tk.Label(stat1, text="מילים נלמדו", font=("Arial", 12),
                 bg=COLORS["card"], fg=COLORS["light_text"]).pack()

        # Total quizzes stat
        stat2 = tk.Frame(stats_frame, bg=COLORS["card"])
        stat2.pack(side="right", expand=True, padx=10)
        tk.Label(stat2, text=str(len(all_scores)), font=("Arial", 24, "bold"),
                 bg=COLORS["card"], fg=COLORS["success"]).pack()
        tk.Label(stat2, text="חידונים הושלמו", font=("Arial", 12),
                 bg=COLORS["card"], fg=COLORS["light_text"]).pack()

        # Average score stat
        stat3 = tk.Frame(stats_frame, bg=COLORS["card"])
        stat3.pack(side="right", expand=True, padx=10)
        avg = (sum(all_scores) / len(all_scores) * 10) if all_scores else 0
        tk.Label(stat3, text=f"{avg:.0f}%", font=("Arial", 24, "bold"),
                 bg=COLORS["card"], fg=COLORS["purple"]).pack()
        tk.Label(stat3, text="ציון ממוצע", font=("Arial", 12),
                 bg=COLORS["card"], fg=COLORS["light_text"]).pack()

        # Overall progress bar
        bar_frame = tk.Frame(overall_card, bg=COLORS["card"])
        bar_frame.pack(fill="x", padx=15, pady=(0, 15))
        bar = self._progress_bar(bar_frame, learned, total_words,
                                 width=self.WINDOW_WIDTH - 130, height=14)
        bar.pack()

        # Per-category cards
        categories = list(VOCABULARY.keys())
        for i, cat in enumerate(categories):
            color = CATEGORY_COLORS[i % len(CATEGORY_COLORS)]
            words = VOCABULARY[cat]
            prog = self.progress.get(cat, {})
            learned_list = prog.get("learned_words", [])
            scores = prog.get("quiz_scores", [])
            best = prog.get("best_score", 0)

            cat_card = tk.Frame(scrollable, bg=COLORS["card"], bd=0,
                                highlightbackground="#D5DBDB", highlightthickness=2)
            cat_card.pack(fill="x", padx=10, pady=5)

            # Category header row
            row = tk.Frame(cat_card, bg=COLORS["card"])
            row.pack(fill="x", padx=15, pady=(10, 5))

            tk.Label(
                row, text=cat, font=("Arial", 15, "bold"),
                bg=COLORS["card"], fg=color,
            ).pack(side="right")

            # Stars
            star_text = self._stars_for_score(best, 10)
            if star_text:
                tk.Label(
                    row, text=star_text, font=("Arial", 14),
                    bg=COLORS["card"],
                ).pack(side="left")

            # Stats row
            detail = tk.Frame(cat_card, bg=COLORS["card"])
            detail.pack(fill="x", padx=15, pady=(0, 5))

            tk.Label(
                detail,
                text=f"מילים: {len(learned_list)}/{len(words)}  |  "
                     f"חידונים: {len(scores)}  |  "
                     f"שיא: {best}/10",
                font=("Arial", 11), bg=COLORS["card"], fg=COLORS["light_text"],
            ).pack(side="right")

            # Progress bar
            bar = self._progress_bar(cat_card, len(learned_list), len(words),
                                     width=self.WINDOW_WIDTH - 130, height=10,
                                     bar_color=color)
            bar.pack(padx=15, pady=(0, 10))

        # Reset button
        tk.Frame(scrollable, bg=COLORS["bg"], height=10).pack()
        reset_btn = self._make_button(
            scrollable, "🗑 אפס התקדמות", self._confirm_reset,
            COLORS["danger"], font_size=12, padx=15, pady=6,
        )
        reset_btn.pack(pady=(5, 20))

    @staticmethod
    def _stars_for_score(score, total):
        if total == 0:
            return ""
        pct = (score / total) * 100
        if pct == 100:
            return "⭐⭐⭐"
        elif pct >= 80:
            return "⭐⭐"
        elif pct >= 50:
            return "⭐"
        return ""

    def _confirm_reset(self):
        if messagebox.askyesno("איפוס התקדמות", "?האם אתה בטוח שברצונך לאפס את כל ההתקדמות"):
            self.progress = {}
            self._save_progress()
            self._show_progress()


# ──────────────────────────── Entry Point ────────────────────────────

if __name__ == "__main__":
    EnglishLearningApp()
