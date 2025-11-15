"""
Excel集計レポート生成ツール
「取得データ」「配点」「aaa（ひな型）」シートから受講者ごとの集計レポートを生成
"""

import openpyxl
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import os
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import traceback


class ExcelReportGenerator:
    def __init__(self):
        self.wb = None
        self.data_sheet = None
        self.point_sheet = None
        self.template_sheet = None
        
    def load_workbook(self, file_path):
        """Excelファイルを読み込む"""
        try:
            self.wb = load_workbook(file_path, keep_vba=True)
            return True
        except Exception as e:
            raise Exception(f"Excelファイルの読み込みに失敗しました: {str(e)}")
    
    def find_sheets(self):
        """必要なシートを検索"""
        sheet_names = self.wb.sheetnames
        
        # 「取得データ」シートを検索
        for name in sheet_names:
            if '取得データ' in name or '取得' in name:
                self.data_sheet = self.wb[name]
                break
        
        # 「配点」シートを検索
        for name in sheet_names:
            if '配点' in name:
                self.point_sheet = self.wb[name]
                break
        
        # 「aaa（ひな型）」シートを検索
        for name in sheet_names:
            if 'aaa' in name.lower() or 'ひな型' in name or '雛型' in name:
                self.template_sheet = self.wb[name]
                break
        
        if not self.data_sheet:
            raise Exception("「取得データ」シートが見つかりません")
        if not self.point_sheet:
            raise Exception("「配点」シートが見つかりません")
        if not self.template_sheet:
            raise Exception("「aaa（ひな型）」シートが見つかりません")
    
    def get_answer_column(self):
        """回答列（Q列）を取得"""
        return 17  # Q列は17列目
    
    def read_point_data(self):
        """配点シートから配点データを読み込む"""
        # 配点シートの構造:
        # B列: セクション名
        # C列: 設問文
        # D列: 問題番号
        # E列: 配点
        points = []
        sections = {}  # セクション別の情報
        
        row = 3  # ヘッダー行は2行目、データは3行目から
        
        while row <= self.point_sheet.max_row:
            # 問題番号を取得（D列）
            question_num_cell = self.point_sheet.cell(row, 4)  # D列
            if question_num_cell.value is None:
                row += 1
                continue
            
            # セクション名を取得（B列）
            section_cell = self.point_sheet.cell(row, 2)  # B列
            section_name = str(section_cell.value).strip() if section_cell.value else ""
            
            # 配点を取得（E列）
            point_cell = self.point_sheet.cell(row, 5)  # E列
            if point_cell.value is not None:
                question_num = int(question_num_cell.value)
                point_value = float(point_cell.value)
                
                points.append({
                    'question_num': question_num,
                    'section': section_name,
                    'point': point_value
                })
                
                # セクション別の集計用
                if section_name not in sections:
                    sections[section_name] = {'total_points': 0, 'questions': []}
                sections[section_name]['total_points'] += point_value
                sections[section_name]['questions'].append(question_num)
            
            row += 1
        
        # 問題番号順にソート
        points.sort(key=lambda x: x['question_num'])
        
        return points, sections
    
    def read_student_data(self):
        """取得データシートから受講者データを読み込む"""
        # 取得データシートの構造:
        # L列（12列目）: 氏名
        # M列（13列目）: メールアドレス
        # Q列（17列目）から: 回答データ（0=不正解、1=正解）
        students = []
        answer_col = self.get_answer_column()  # Q列 = 17列目
        
        # ヘッダー行は1行目、データは2行目から
        data_start_row = 2
        
        for row in range(data_start_row, self.data_sheet.max_row + 1):
            # 氏名を取得（L列 = 12列目）
            name_cell = self.data_sheet.cell(row, 12)
            # メールアドレスを取得（M列 = 13列目）
            email_cell = self.data_sheet.cell(row, 13)
            
            if name_cell.value is None:
                continue
            
            # 回答データを取得（Q列から）
            # 構造: 各問題について3列使用
            # - 設問文の列（P, S, V, Y, ...）
            # - 点数の列（Q, T, W, Z, ...）- これが0または1の値
            # - フィードバックの列（R, U, X, AA, ...）
            # したがって、回答データはQ列から3列おきに取得
            answers = []
            col = answer_col  # Q列 = 17列目
            # 最大195列まで（実際のファイルの最大列数）
            while col <= min(195, self.data_sheet.max_column):
                answer_cell = self.data_sheet.cell(row, col)
                
                # 回答値を取得（0または1）
                if answer_cell.value is None:
                    # 空のセルが続く場合は終了
                    if col > answer_col and answers:
                        break
                    # 空セルを0として扱う
                    answers.append(0)
                else:
                    try:
                        # 数値として解釈できるか確認
                        answer_value = int(answer_cell.value) if answer_cell.value != '' else 0
                        # 0または1の値のみを有効な回答として扱う
                        if answer_value in [0, 1]:
                            answers.append(answer_value)
                        else:
                            # 0または1以外の数値の場合は0として扱う
                            answers.append(0)
                    except (ValueError, TypeError):
                        # 数値でない場合は0として扱う
                        answers.append(0)
                
                # 次の回答列へ（3列おき）
                col += 3
            
            if answers:  # 回答がある場合のみ追加
                students.append({
                    'name': str(name_cell.value).strip() if name_cell.value else '',
                    'email': str(email_cell.value).strip() if email_cell.value else '',
                    'answers': answers,
                    'row': row
                })
        
        return students
    
    def calculate_scores(self, students, points_data, sections_data):
        """各受講者の得点を計算"""
        # points_data: 問題番号順にソートされた配点データのリスト
        # sections_data: セクション別の情報
        results = []
        
        for student in students:
            answers = student['answers']
            total_score = 0
            max_score = 0
            section_scores = {}  # セクション別の得点
            
            # セクション別の集計を初期化
            for section_name in sections_data.keys():
                section_scores[section_name] = {'score': 0, 'max_score': sections_data[section_name]['total_points']}
            
            # 配点データと回答を照合
            question_scores = []
            
            for i, point_info in enumerate(points_data):
                question_num = point_info['question_num']
                section_name = point_info['section']
                point_value = point_info['point']
                
                max_score += point_value
                
                # 回答を取得（問題番号は1から始まるので、インデックスはquestion_num-1）
                if question_num - 1 < len(answers):
                    answer = answers[question_num - 1]
                else:
                    answer = 0  # 回答がない場合は0（不正解）
                
                if answer == 1:  # 正解
                    total_score += point_value
                    section_scores[section_name]['score'] += point_value
                    question_scores.append({
                        'question_num': question_num,
                        'section': section_name,
                        'point': point_value,
                        'correct': True
                    })
                else:  # 不正解
                    question_scores.append({
                        'question_num': question_num,
                        'section': section_name,
                        'point': point_value,
                        'correct': False
                    })
            
            # 5点評価を計算（100点満点として）
            if max_score > 0:
                percentage = (total_score / max_score) * 100
                if percentage >= 90:
                    rating = 5
                elif percentage >= 80:
                    rating = 4
                elif percentage >= 70:
                    rating = 3
                elif percentage >= 60:
                    rating = 2
                else:
                    rating = 1
            else:
                rating = 0
                percentage = 0
            
            results.append({
                'name': student['name'],
                'email': student['email'],
                'total_score': total_score,
                'max_score': max_score,
                'percentage': percentage,
                'rating': rating,
                'section_scores': section_scores,
                'question_scores': question_scores
            })
        
        return results
    
    def create_report_sheet(self, result, template_sheet_name, student_row_index):
        """個別レポートシートを作成"""
        # テンプレートシートをコピー
        template = self.wb[template_sheet_name]
        new_sheet_name = f"{result['name']}_レポート"
        # シート名が長すぎる場合は短縮
        if len(new_sheet_name) > 31:
            new_sheet_name = new_sheet_name[:28] + "..."
        
        # 既存のシートがある場合は削除
        if new_sheet_name in self.wb.sheetnames:
            self.wb.remove(self.wb[new_sheet_name])
        
        new_sheet = self.wb.copy_worksheet(template)
        new_sheet.title = new_sheet_name
        
        # テンプレートシートの構造に基づいてデータを埋め込む
        # A2セル: 受講者名
        try:
            new_sheet['A2'] = result['name']
        except:
            pass
        
        # D27-D31セル: 各セクションの5点評価
        # テンプレートでは数式で「5点評価」シートから取得しているが、
        # 個別レポートでは直接値を設定するか、数式を更新する
        # ここでは直接値を設定する方法を採用
        section_names = list(result['section_scores'].keys())
        section_row_mapping = {
            0: 27,  # 1番目のセクション -> D27
            1: 28,  # 2番目のセクション -> D28
            2: 29,  # 3番目のセクション -> D29
            3: 30,  # 4番目のセクション -> D30
            4: 31,  # 5番目のセクション -> D31
        }
        
        for idx, section_name in enumerate(section_names):
            if idx in section_row_mapping:
                row = section_row_mapping[idx]
                # セクション名をB列に設定
                try:
                    new_sheet.cell(row, 2).value = section_name
                except:
                    pass
                
                # 5点評価を計算してD列に設定
                section_score = result['section_scores'].get(section_name, {'score': 0, 'max_score': 1})
                if section_score['max_score'] > 0:
                    section_percentage = (section_score['score'] / section_score['max_score']) * 100
                    if section_percentage >= 90:
                        section_rating = 5
                    elif section_percentage >= 80:
                        section_rating = 4
                    elif section_percentage >= 70:
                        section_rating = 3
                    elif section_percentage >= 60:
                        section_rating = 2
                    else:
                        section_rating = 1
                else:
                    section_rating = 0
                
                try:
                    new_sheet.cell(row, 4).value = section_rating
                except:
                    pass
        
        # 総合点（E4セル）は数式で自動計算されるが、念のため確認
        # テンプレートの数式が正しく動作することを前提とする
        
        return new_sheet
    
    def create_summary_sheet(self, results, sections_data):
        """集計シートを作成（総合得点、5点評価）"""
        # 総合得点シートを作成/更新
        summary_name = "総合得点"
        if summary_name in self.wb.sheetnames:
            self.wb.remove(self.wb[summary_name])
        
        summary_sheet = self.wb.create_sheet(summary_name)
        
        # ヘッダー行
        headers = ['氏名'] + list(sections_data.keys()) + ['総合得点（満点）']
        for col, header in enumerate(headers, 1):
            cell = summary_sheet.cell(2, col)
            cell.value = header
            cell.font = openpyxl.styles.Font(bold=True)
        
        # データ
        for row_idx, result in enumerate(results, 3):
            summary_sheet.cell(row_idx, 1).value = result['name']
            col_idx = 2
            for section_name in sections_data.keys():
                section_score = result['section_scores'].get(section_name, {'score': 0})
                summary_sheet.cell(row_idx, col_idx).value = section_score['score']
                col_idx += 1
            summary_sheet.cell(row_idx, col_idx).value = f"{result['total_score']}（{result['max_score']}点満点）"
        
        # 列幅を調整
        for col in range(1, len(headers) + 1):
            summary_sheet.column_dimensions[get_column_letter(col)].width = 20
        
        # 5点評価シートを作成/更新
        rating_name = "5点評価"
        if rating_name in self.wb.sheetnames:
            self.wb.remove(self.wb[rating_name])
        
        rating_sheet = self.wb.create_sheet(rating_name)
        
        # ヘッダー行
        rating_headers = ['氏名'] + list(sections_data.keys()) + ['総合得点']
        for col, header in enumerate(rating_headers, 1):
            cell = rating_sheet.cell(2, col)
            cell.value = header
            cell.font = openpyxl.styles.Font(bold=True)
        
        # データ（5点評価を計算）
        for row_idx, result in enumerate(results, 3):
            rating_sheet.cell(row_idx, 1).value = result['name']
            col_idx = 2
            for section_name in sections_data.keys():
                section_score = result['section_scores'].get(section_name, {'score': 0, 'max_score': 1})
                if section_score['max_score'] > 0:
                    section_percentage = (section_score['score'] / section_score['max_score']) * 100
                    if section_percentage >= 90:
                        section_rating = 5
                    elif section_percentage >= 80:
                        section_rating = 4
                    elif section_percentage >= 70:
                        section_rating = 3
                    elif section_percentage >= 60:
                        section_rating = 2
                    else:
                        section_rating = 1
                else:
                    section_rating = 0
                rating_sheet.cell(row_idx, col_idx).value = section_rating
                col_idx += 1
            rating_sheet.cell(row_idx, col_idx).value = result['rating']
        
        # 列幅を調整
        for col in range(1, len(rating_headers) + 1):
            rating_sheet.column_dimensions[get_column_letter(col)].width = 15
    
    def generate_reports(self, output_path=None):
        """レポートを生成"""
        try:
            # シートを検索
            self.find_sheets()
            
            # データを読み込む
            points_data, sections_data = self.read_point_data()
            students = self.read_student_data()
            
            if not students:
                raise Exception("受講者データが見つかりません")
            
            if not points_data:
                raise Exception("配点データが見つかりません")
            
            # 得点を計算
            results = self.calculate_scores(students, points_data, sections_data)
            
            # 集計シートを作成
            self.create_summary_sheet(results, sections_data)
            
            # 個別レポートシートを作成
            template_sheet_name = self.template_sheet.title
            for idx, result in enumerate(results, 3):  # 3行目から開始（ヘッダー行が2行目）
                self.create_report_sheet(result, template_sheet_name, idx)
            
            # ファイルを保存
            if output_path:
                self.wb.save(output_path)
            else:
                # 元のファイル名に「_出力」を追加
                original_path = self.wb.path if hasattr(self.wb, 'path') else None
                if original_path:
                    base_path = Path(original_path)
                    output_path = base_path.parent / f"{base_path.stem}_出力{base_path.suffix}"
                    self.wb.save(output_path)
            
            return results, str(output_path)
            
        except Exception as e:
            raise Exception(f"レポート生成中にエラーが発生しました: {str(e)}\n{traceback.format_exc()}")


class ReportGeneratorUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel集計レポート生成ツール")
        self.root.geometry("600x400")
        
        self.generator = ExcelReportGenerator()
        self.file_path = None
        
        self.setup_ui()
    
    def setup_ui(self):
        """UIを構築"""
        # メインフレーム
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # ファイル選択
        file_frame = ttk.LabelFrame(main_frame, text="Excelファイル選択", padding="10")
        file_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.file_label = ttk.Label(file_frame, text="ファイルが選択されていません", foreground="gray")
        self.file_label.pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(file_frame, text="ファイルを選択", command=self.select_file).pack(side=tk.RIGHT)
        
        # 実行ボタン
        execute_frame = ttk.Frame(main_frame)
        execute_frame.pack(fill=tk.X, pady=10)
        
        self.execute_button = ttk.Button(
            execute_frame,
            text="レポートを生成",
            command=self.generate_reports,
            state=tk.DISABLED
        )
        self.execute_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # 進捗表示
        self.progress = ttk.Progressbar(execute_frame, mode='indeterminate')
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # ログ表示
        log_frame = ttk.LabelFrame(main_frame, text="ログ", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        
        self.log_text = tk.Text(log_frame, height=15, wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.log("ツールを起動しました。")
        self.log("Excelファイルを選択してください。")
    
    def log(self, message):
        """ログを追加"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update()
    
    def select_file(self):
        """ファイルを選択"""
        file_path = filedialog.askopenfilename(
            title="Excelファイルを選択",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")]
        )
        
        if file_path:
            self.file_path = file_path
            self.file_label.config(text=os.path.basename(file_path), foreground="black")
            self.execute_button.config(state=tk.NORMAL)
            self.log(f"ファイルを選択しました: {os.path.basename(file_path)}")
    
    def generate_reports(self):
        """レポートを生成"""
        if not self.file_path:
            messagebox.showerror("エラー", "ファイルを選択してください。")
            return
        
        try:
            self.execute_button.config(state=tk.DISABLED)
            self.progress.start()
            self.log("Excelファイルを読み込んでいます...")
            
            # ファイルを読み込む
            self.generator.load_workbook(self.file_path)
            self.log("ファイルの読み込みが完了しました。")
            
            # レポートを生成
            self.log("レポートを生成しています...")
            results, output_path = self.generator.generate_reports()
            
            self.progress.stop()
            self.execute_button.config(state=tk.NORMAL)
            
            self.log(f"レポート生成が完了しました！")
            self.log(f"出力ファイル: {output_path}")
            self.log(f"処理した受講者数: {len(results)}名")
            
            messagebox.showinfo(
                "完了",
                f"レポート生成が完了しました。\n\n"
                f"出力ファイル: {os.path.basename(output_path)}\n"
                f"処理した受講者数: {len(results)}名"
            )
            
        except Exception as e:
            self.progress.stop()
            self.execute_button.config(state=tk.NORMAL)
            error_msg = str(e)
            self.log(f"エラー: {error_msg}")
            messagebox.showerror("エラー", f"レポート生成中にエラーが発生しました:\n{error_msg}")


def main():
    root = tk.Tk()
    app = ReportGeneratorUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()

