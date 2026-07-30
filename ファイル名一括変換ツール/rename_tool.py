"""
ファイル名一括変換ツール
カナ全角・英数半角 ガイドライン準拠版
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import re
import csv
import datetime
import threading

# ─── 変換ルール定義 ────────────────────────────────────────────

# 半角カタカナ → 全角カタカナ マッピング
_HANKAKU_TO_ZENKAKU = {
    'ｦ': 'ヲ', 'ｧ': 'ァ', 'ｨ': 'ィ', 'ｩ': 'ゥ', 'ｪ': 'ェ', 'ｫ': 'ォ',
    'ｬ': 'ャ', 'ｭ': 'ュ', 'ｮ': 'ョ', 'ｯ': 'ッ', 'ｰ': 'ー',
    'ｱ': 'ア', 'ｲ': 'イ', 'ｳ': 'ウ', 'ｴ': 'エ', 'ｵ': 'オ',
    'ｶ': 'カ', 'ｷ': 'キ', 'ｸ': 'ク', 'ｹ': 'ケ', 'ｺ': 'コ',
    'ｻ': 'サ', 'ｼ': 'シ', 'ｽ': 'ス', 'ｾ': 'セ', 'ｿ': 'ソ',
    'ﾀ': 'タ', 'ﾁ': 'チ', 'ﾂ': 'ツ', 'ﾃ': 'テ', 'ﾄ': 'ト',
    'ﾅ': 'ナ', 'ﾆ': 'ニ', 'ﾇ': 'ヌ', 'ﾈ': 'ネ', 'ﾉ': 'ノ',
    'ﾊ': 'ハ', 'ﾋ': 'ヒ', 'ﾌ': 'フ', 'ﾍ': 'ヘ', 'ﾎ': 'ホ',
    'ﾏ': 'マ', 'ﾐ': 'ミ', 'ﾑ': 'ム', 'ﾒ': 'メ', 'ﾓ': 'モ',
    'ﾔ': 'ヤ', 'ﾕ': 'ユ', 'ﾖ': 'ヨ',
    'ﾗ': 'ラ', 'ﾘ': 'リ', 'ﾙ': 'ル', 'ﾚ': 'レ', 'ﾛ': 'ロ',
    'ﾜ': 'ワ', 'ﾝ': 'ン',
}
_DAKUTEN = {
    'カ': 'ガ', 'キ': 'ギ', 'ク': 'グ', 'ケ': 'ゲ', 'コ': 'ゴ',
    'サ': 'ザ', 'シ': 'ジ', 'ス': 'ズ', 'セ': 'ゼ', 'ソ': 'ゾ',
    'タ': 'ダ', 'チ': 'ヂ', 'ツ': 'ヅ', 'テ': 'デ', 'ト': 'ド',
    'ハ': 'バ', 'ヒ': 'ビ', 'フ': 'ブ', 'ヘ': 'ベ', 'ホ': 'ボ',
    'ウ': 'ヴ',
}
_HANDAKUTEN = {
    'ハ': 'パ', 'ヒ': 'ピ', 'フ': 'プ', 'ヘ': 'ペ', 'ホ': 'ポ',
}

SPECIFIC_RULES = [
    ('－', '-'),
    ('※', '注'),
    ('①', '_1'), ('②', '_2'), ('③', '_3'), ('④', '_4'), ('⑤', '_5'),
    ('⑥', '_6'), ('⑦', '_7'), ('⑧', '_8'), ('⑨', '_9'),
    ('（', '('), ('）', ')'),
    ('【', '('), ('】', ')'),
    ('［', '('), ('］', ')'),
    ('『', '('), ('』', ')'),
    ('｛', '('), ('｝', ')'),
    ('〈', '('), ('〉', ')'),
    ('＜', '('), ('＞', ')'),
]

# 許可文字パターン: ひらがな・カタカナ・漢字・英数字・記号一部
ALLOWED_RE = re.compile(r'^[ぁ-んァ-ヶa-zA-Z0-9ー()_一-龠々\-△]$')


def hankaku_kana_to_zenkaku(text: str) -> str:
    """半角カタカナを全角カタカナに変換（濁音・半濁音対応）"""
    result = []
    i = 0
    while i < len(text):
        c = text[i]
        zk = _HANKAKU_TO_ZENKAKU.get(c)
        if zk:
            if i + 1 < len(text):
                nxt = text[i + 1]
                if nxt == 'ﾞ' and zk in _DAKUTEN:
                    result.append(_DAKUTEN[zk])
                    i += 2
                    continue
                if nxt == 'ﾟ' and zk in _HANDAKUTEN:
                    result.append(_HANDAKUTEN[zk])
                    i += 2
                    continue
            result.append(zk)
        else:
            result.append(c)
        i += 1
    return ''.join(result)


def zenkaku_alnum_to_hankaku(text: str) -> str:
    """全角英数字を半角に変換"""
    result = []
    for c in text:
        v = ord(c)
        if 0xFF10 <= v <= 0xFF19:   # ０-９
            result.append(chr(v - 0xFF10 + 0x30))
        elif 0xFF21 <= v <= 0xFF3A:  # Ａ-Ｚ
            result.append(chr(v - 0xFF21 + 0x41))
        elif 0xFF41 <= v <= 0xFF5A:  # ａ-ｚ
            result.append(chr(v - 0xFF41 + 0x61))
        else:
            result.append(c)
    return ''.join(result)


def get_new_name(old_name: str, is_file: bool) -> str:
    """ガイドライン準拠のファイル名変換"""
    if is_file:
        base, ext = os.path.splitext(old_name)
        if not base:          # ".gitignore" のようなケース
            base, ext = old_name, ''
        raw = base
    else:
        raw, ext = old_name, ''

    # 1. 半角カナ → 全角カナ
    raw = hankaku_kana_to_zenkaku(raw)
    # 2. 全角英数字 → 半角
    raw = zenkaku_alnum_to_hankaku(raw)
    # 3. 特定文字置換
    for old, new in SPECIFIC_RULES:
        raw = raw.replace(old, new)
    # 4. 許可文字以外 → _
    raw = ''.join(c if ALLOWED_RE.match(c) else '_' for c in raw)
    # 5. 連続アンダーバーを1つに
    while '__' in raw:
        raw = raw.replace('__', '_')
    # 6. 先頭の禁止記号を除去
    raw = re.sub(r'^[ _()\-△]+', '', raw)
    # 7. 末尾のアンダーバーを除去
    raw = raw.rstrip('_')

    if not raw.strip():
        raw = 'renamed_' + datetime.datetime.now().strftime('%S')

    return raw + ext


def collect_items(folder: str, recursive: bool):
    """変換対象アイテム一覧を取得（深い階層から順に返す）"""
    items = []
    if recursive:
        for dirpath, dirnames, filenames in os.walk(folder):
            for fn in filenames:
                items.append(os.path.join(dirpath, fn))
            items.append(dirpath)  # フォルダ自身（後で処理）
    else:
        with os.scandir(folder) as it:
            for entry in it:
                items.append(entry.path)

    # 深い順にソート（子から先にリネームするため）
    items.sort(key=lambda p: p.count(os.sep), reverse=True)
    return items


# ─── GUI ──────────────────────────────────────────────────────

class RenameApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('ファイル名一括変換ツール')
        self.geometry('900x640')
        self.resizable(True, True)
        self.configure(bg='#f5f5f5')

        self._preview_data = []   # [(type, orig_path, new_name), ...]
        self._running = False
        self._build_ui()

    # ── UI 構築 ──────────────────────────────────────────────

    def _build_ui(self):
        # ── ヘッダー
        hdr = tk.Frame(self, bg='#2c5f8a', pady=10)
        hdr.pack(fill='x')
        tk.Label(hdr, text='ファイル名一括変換ツール',
                 bg='#2c5f8a', fg='white',
                 font=('Yu Gothic UI', 14, 'bold')).pack()
        tk.Label(hdr, text='カナ全角・英数半角  ガイドライン準拠版',
                 bg='#2c5f8a', fg='#aed6f1',
                 font=('Yu Gothic UI', 9)).pack()

        # ── フォルダ選択
        frm_folder = tk.LabelFrame(self, text=' 対象フォルダ ', bg='#f5f5f5',
                                   font=('Yu Gothic UI', 9))
        frm_folder.pack(fill='x', padx=12, pady=(10, 4))

        self.folder_var = tk.StringVar()
        tk.Entry(frm_folder, textvariable=self.folder_var,
                 font=('Yu Gothic UI', 10), relief='flat',
                 bg='white', bd=1).pack(side='left', fill='x', expand=True,
                                        padx=(8, 4), pady=8)
        tk.Button(frm_folder, text='📁  フォルダを選択',
                  command=self._browse_folder,
                  bg='#2c5f8a', fg='white',
                  relief='flat', padx=12, pady=4,
                  font=('Yu Gothic UI', 9),
                  cursor='hand2').pack(side='left', padx=(0, 8), pady=8)

        # ── オプション
        frm_opt = tk.Frame(self, bg='#f5f5f5')
        frm_opt.pack(fill='x', padx=12, pady=2)

        self.recursive_var = tk.BooleanVar(value=True)
        tk.Checkbutton(frm_opt, text='サブフォルダも対象にする',
                       variable=self.recursive_var,
                       bg='#f5f5f5', font=('Yu Gothic UI', 9)).pack(side='left')

        self.dryrun_var = tk.BooleanVar(value=True)
        cb_dry = tk.Checkbutton(frm_opt, text='テストモード（実際には変更しない）',
                                variable=self.dryrun_var,
                                bg='#f5f5f5', fg='#c0392b',
                                font=('Yu Gothic UI', 9, 'bold'))
        cb_dry.pack(side='left', padx=20)

        # ── ボタン行
        frm_btn = tk.Frame(self, bg='#f5f5f5')
        frm_btn.pack(fill='x', padx=12, pady=6)

        self.btn_scan = tk.Button(frm_btn, text='🔍  スキャン（プレビュー）',
                                  command=self._start_scan,
                                  bg='#27ae60', fg='white', relief='flat',
                                  padx=16, pady=6,
                                  font=('Yu Gothic UI', 10, 'bold'),
                                  cursor='hand2')
        self.btn_scan.pack(side='left', padx=(0, 8))

        self.btn_exec = tk.Button(frm_btn, text='▶  実行',
                                  command=self._start_execute,
                                  state='disabled',
                                  bg='#e67e22', fg='white', relief='flat',
                                  padx=16, pady=6,
                                  font=('Yu Gothic UI', 10, 'bold'),
                                  cursor='hand2')
        self.btn_exec.pack(side='left')

        self.btn_csv = tk.Button(frm_btn, text='💾  CSVで保存',
                                 command=self._save_csv,
                                 state='disabled',
                                 bg='#7f8c8d', fg='white', relief='flat',
                                 padx=16, pady=6,
                                 font=('Yu Gothic UI', 9),
                                 cursor='hand2')
        self.btn_csv.pack(side='right')

        # ── プレビューテーブル
        frm_tree = tk.LabelFrame(self, text=' プレビュー ',
                                 bg='#f5f5f5', font=('Yu Gothic UI', 9))
        frm_tree.pack(fill='both', expand=True, padx=12, pady=4)

        cols = ('type', 'before', 'after', 'folder')
        self.tree = ttk.Treeview(frm_tree, columns=cols, show='headings',
                                 selectmode='browse')
        self.tree.heading('type',   text='種別', anchor='center')
        self.tree.heading('before', text='変更前')
        self.tree.heading('after',  text='変更後')
        self.tree.heading('folder', text='フォルダ')
        self.tree.column('type',   width=60,  anchor='center', stretch=False)
        self.tree.column('before', width=240)
        self.tree.column('after',  width=240)
        self.tree.column('folder', width=300)

        self.tree.tag_configure('file',   background='#ffffff')
        self.tree.tag_configure('folder', background='#eaf4fb')

        vsb = ttk.Scrollbar(frm_tree, orient='vertical',   command=self.tree.yview)
        hsb = ttk.Scrollbar(frm_tree, orient='horizontal', command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        frm_tree.grid_rowconfigure(0, weight=1)
        frm_tree.grid_columnconfigure(0, weight=1)

        # ── ステータスバー
        self.status_var = tk.StringVar(value='フォルダを選択してスキャンしてください')
        self.prog = ttk.Progressbar(self, mode='indeterminate', length=200)
        self.prog.pack(fill='x', padx=12, pady=(0, 2))
        tk.Label(self, textvariable=self.status_var,
                 bg='#f5f5f5', anchor='w',
                 font=('Yu Gothic UI', 9)).pack(fill='x', padx=12, pady=(0, 6))

    # ── イベントハンドラ ──────────────────────────────────────

    def _browse_folder(self):
        path = filedialog.askdirectory(title='変換対象フォルダを選択')
        if path:
            self.folder_var.set(path)
            self.btn_exec.config(state='disabled')
            self.btn_csv.config(state='disabled')
            self._clear_tree()

    def _start_scan(self):
        folder = self.folder_var.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning('フォルダ未選択', '有効なフォルダを選択してください。')
            return
        self._set_busy(True)
        self._clear_tree()
        threading.Thread(target=self._scan_worker, args=(folder,), daemon=True).start()

    def _scan_worker(self, folder):
        try:
            items = collect_items(folder, self.recursive_var.get())
            preview = []
            for path in items:
                is_file = os.path.isfile(path)
                old_name = os.path.basename(path)
                new_name = get_new_name(old_name, is_file)
                if old_name != new_name:
                    preview.append((
                        'ファイル' if is_file else 'フォルダ',
                        path,
                        new_name,
                        os.path.dirname(path)
                    ))
            self._preview_data = preview
            self.after(0, self._scan_done)
        except Exception as e:
            self.after(0, lambda: self._on_error(str(e)))

    def _scan_done(self):
        self._set_busy(False)
        count = len(self._preview_data)
        if count == 0:
            self.status_var.set('変更が必要なファイル・フォルダはありませんでした。')
            return

        for typ, orig, new_name, folder in self._preview_data:
            tag = 'file' if typ == 'ファイル' else 'folder'
            self.tree.insert('', 'end',
                             values=(typ, os.path.basename(orig), new_name, folder),
                             tags=(tag,))

        mode = '【テストモード】' if self.dryrun_var.get() else '【本番モード】'
        self.status_var.set(f'{mode}  {count} 件が変換対象です。「実行」を押して変換します。')
        self.btn_exec.config(state='normal')
        self.btn_csv.config(state='normal')

    def _start_execute(self):
        if not self._preview_data:
            return
        mode = 'テストモード（実際には変更しません）' if self.dryrun_var.get() else '本番モード（実際にファイル名を変更します）'
        if not messagebox.askyesno('確認', f'{len(self._preview_data)} 件を変換します。\n\n現在: {mode}\n\n続けますか？'):
            return
        self._set_busy(True)
        threading.Thread(target=self._execute_worker, daemon=True).start()

    def _execute_worker(self):
        dry = self.dryrun_var.get()
        errors = []
        done = 0
        for typ, orig_path, new_name, folder in self._preview_data:
            try:
                if not dry:
                    dst = os.path.join(os.path.dirname(orig_path), new_name)
                    os.rename(orig_path, dst)
                done += 1
            except Exception as e:
                errors.append(f'{orig_path}: {e}')
        self.after(0, lambda: self._execute_done(done, errors, dry))

    def _execute_done(self, done, errors, dry):
        self._set_busy(False)
        mode_str = '（テストモード）' if dry else ''
        msg = f'完了{mode_str}: {done} 件処理しました。'
        if errors:
            msg += f'\nエラー {len(errors)} 件:\n' + '\n'.join(errors[:5])
        self.status_var.set(msg.split('\n')[0])
        messagebox.showinfo('完了', msg)
        if not dry:
            self.btn_exec.config(state='disabled')

    def _save_csv(self):
        if not self._preview_data:
            return
        ts = datetime.datetime.now().strftime('%Y%m%d_%H%M')
        default = f'RenameLog_{ts}.csv'
        path = filedialog.asksaveasfilename(
            defaultextension='.csv',
            filetypes=[('CSV ファイル', '*.csv')],
            initialfile=default,
            title='ログCSVを保存'
        )
        if not path:
            return
        try:
            with open(path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                writer.writerow(['種別', '変更前（フルパス）', '変更後ファイル名', 'フォルダ'])
                for typ, orig, new_name, folder in self._preview_data:
                    writer.writerow([typ, orig, new_name, folder])
            self.status_var.set(f'CSV保存完了: {path}')
            messagebox.showinfo('保存完了', f'CSVを保存しました:\n{path}')
        except Exception as e:
            messagebox.showerror('保存エラー', str(e))

    # ── ヘルパー ─────────────────────────────────────────────

    def _clear_tree(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        self._preview_data = []

    def _set_busy(self, busy: bool):
        if busy:
            self.prog.start(10)
            self.btn_scan.config(state='disabled')
            self.btn_exec.config(state='disabled')
        else:
            self.prog.stop()
            self.btn_scan.config(state='normal')

    def _on_error(self, msg: str):
        self._set_busy(False)
        self.status_var.set(f'エラー: {msg}')
        messagebox.showerror('エラー', msg)


if __name__ == '__main__':
    app = RenameApp()
    app.mainloop()
