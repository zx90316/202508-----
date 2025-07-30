import pandas as pd
import json
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.dml import MSO_THEME_COLOR
import os

class ExcelToJsonPresentation:
    def __init__(self, excel_file_path):
        """
        初始化類別
        
        Args:
            excel_file_path (str): Excel文件路徑
        """
        self.excel_file_path = excel_file_path
        self.json_data = None
        
    def _set_font_style(self, text_frame, font_name="微軟正黑體", font_size=Pt(12)):
        """
        設定文字框架的字體樣式
        
        Args:
            text_frame: 文字框架
            font_name (str): 字體名稱
            font_size: 字體大小
        """
        for paragraph in text_frame.paragraphs:
            # 設定段落字體
            paragraph.font.name = font_name
            paragraph.font.size = font_size
            
            # 設定runs字體
            for run in paragraph.runs:
                run.font.name = font_name
                run.font.size = font_size
            
            # 如果段落沒有runs，添加一個並設定字體
            if not paragraph.runs:
                run = paragraph.add_run()
                run.font.name = font_name
                run.font.size = font_size
    
    def _set_paragraph_font(self, paragraph, font_name="微軟正黑體", font_size=Pt(12)):
        """
        設定段落的字體樣式
        
        Args:
            paragraph: 段落對象
            font_name (str): 字體名稱
            font_size: 字體大小
        """
        if paragraph.runs:
            for run in paragraph.runs:
                run.font.name = font_name
                run.font.size = font_size
        else:
            # 如果段落沒有runs，創建一個
            run = paragraph.runs[0] if paragraph.runs else paragraph.add_run()
            run.font.name = font_name
            run.font.size = font_size
        
    def read_excel_to_json(self, sheet_name=None):
        """
        讀取Excel文件並轉換為JSON格式
        
        Args:
            sheet_name (str, optional): 工作表名稱，預設為None（讀取第一個工作表）
            
        Returns:
            dict: 轉換後的JSON數據
        """
        try:
            # 讀取Excel文件，使用openpyxl引擎以更好地處理資料驗證
            if sheet_name:
                df = pd.read_excel(self.excel_file_path, sheet_name=sheet_name, engine='openpyxl')
            else:
                df = pd.read_excel(self.excel_file_path, engine='openpyxl')
            
            # 使用openpyxl直接讀取Excel以獲取資料驗證欄位的實際值
            from openpyxl import load_workbook
            wb = load_workbook(self.excel_file_path, data_only=True)
            if sheet_name:
                ws = wb[sheet_name]
            else:
                ws = wb.active
            
            # 重新構建DataFrame，確保讀取到實際的儲存格值
            data_rows = []
            headers = [cell.value if cell.value is not None else f'Unnamed_{i}' for i, cell in enumerate(ws[1])]  # 第一行作為標題，處理None值
            
            for row in ws.iter_rows(min_row=2, values_only=True):
                if any(cell is not None for cell in row):  # 跳過完全空白的行
                    data_rows.append(row)
            
            # 創建新的DataFrame
            df = pd.DataFrame(data_rows, columns=headers)
            
            # 清理數據：移除空值行
            df = df.dropna(how='all')
            
            # 處理NaN值和None值：將其替換為空字串
            df = df.fillna('')
            df = df.replace({None: ''})
            
            # 按類型欄位分類議題
            categorized_data = {}
            if '類型' in df.columns:
                # 按類型分組
                grouped = df.groupby('類型')
                for category, group in grouped:
                    categorized_data[category] = group.to_dict('records')
                
                # 計算各類型的統計資訊
                category_stats = {}
                for category, items in categorized_data.items():
                    category_stats[category] = len(items)
            else:
                # 如果沒有類型欄位，將所有資料放在 "未分類" 中
                categorized_data['未分類'] = df.to_dict('records')
                category_stats = {'未分類': len(df)}
            
            # 轉換為JSON格式
            json_data = {
                'metadata': {
                    'file_name': os.path.basename(self.excel_file_path),
                    'total_rows': len(df),
                    'columns': list(df.columns),
                    'sheet_name': sheet_name if sheet_name else 'Sheet1',
                    'categories': list(categorized_data.keys()),
                    'category_stats': category_stats
                },
                'data_by_category': categorized_data,
                'data': df.to_dict('records')  # 保留原始格式以便向後相容
            }
            
            self.json_data = json_data
            return json_data
            
        except Exception as e:
            print(f"讀取Excel文件時發生錯誤: {str(e)}")
            return None
    
    def save_json(self, output_path):
        """
        將JSON數據保存到文件
        
        Args:
            output_path (str): 輸出JSON文件路徑
        """
        if self.json_data:
            try:
                with open(output_path, 'w', encoding='utf-8') as f:
                    json.dump(self.json_data, f, ensure_ascii=False, indent=2)
                print(f"JSON文件已保存至: {output_path}")
                return True
            except Exception as e:
                print(f"保存JSON文件時發生錯誤: {str(e)}")
                return False
        else:
            print("沒有JSON數據可保存，請先執行 read_excel_to_json()")
            return False
    
    def create_presentation(self, output_pptx_path, items_per_page=3):
        """
        根據JSON數據創建PowerPoint簡報
        
        Args:
            output_pptx_path (str): 輸出PPT文件路徑
            items_per_page (int): 每頁顯示的議題數量，預設為3
        """
        if not self.json_data:
            print("沒有JSON數據可用於創建簡報，請先執行 read_excel_to_json()")
            return False
        
        try:
            # 創建新的簡報
            prs = Presentation()
            
            # 設置簡報標題頁
            title_slide_layout = prs.slide_layouts[0]
            slide = prs.slides.add_slide(title_slide_layout)
            title = slide.shapes.title
            subtitle = slide.placeholders[1]
            
            title.text = f"系統交流會議題討論 - {self.json_data['metadata']['file_name'].replace('.xlsx', '')}"
            subtitle.text = f"共 {self.json_data['metadata']['total_rows']} 筆議題"
            
            # 設定標題頁字體
            self._set_font_style(title.text_frame, font_size=Pt(30))
            self._set_font_style(subtitle.text_frame, font_size=Pt(20))
            
            # 添加概覽頁
            bullet_slide_layout = prs.slide_layouts[1]
            slide = prs.slides.add_slide(bullet_slide_layout)
            title = slide.shapes.title
            content = slide.placeholders[1]
            
            title.text = "議題概覽"
            tf = content.text_frame
            tf.text = f"總議題數: {self.json_data['metadata']['total_rows']}"
            
            # 顯示各類型統計
            if 'category_stats' in self.json_data['metadata']:
                p = tf.add_paragraph()
                p.text = "各類型議題統計:"
                for category, count in self.json_data['metadata']['category_stats'].items():
                    p = tf.add_paragraph()
                    p.text = f"  {category}: {count} 筆"
                    p.level = 1
            
            # 設定概覽頁字體
            self._set_font_style(title.text_frame, font_size=Pt(28))
            self._set_font_style(tf, font_size=Pt(20))
            
            # 按類型分組顯示議題
            if 'data_by_category' in self.json_data:
                for category, items in self.json_data['data_by_category'].items():
                    self._add_category_slides(prs, category, items, items_per_page, bullet_slide_layout)
            else:
                # 如果沒有分類數據，使用原始數據
                self._add_category_slides(prs, "所有議題", self.json_data['data'], items_per_page, bullet_slide_layout)
            
            # 保存簡報
            prs.save(output_pptx_path)
            print(f"簡報已保存至: {output_pptx_path}")
            return True
            
        except Exception as e:
            print(f"創建簡報時發生錯誤: {str(e)}")
            return False

    def _add_category_slides(self, prs, category, items, items_per_page, bullet_slide_layout, max_content_length=200):
        """
        為特定類型的議題添加簡報頁面
        
        Args:
            prs: PowerPoint簡報對象
            category (str): 議題類型
            items (list): 該類型的議題列表
            items_per_page (int): 每頁顯示的議題數量
            bullet_slide_layout: 簡報佈局
            max_content_length (int): 每個議題內容的最大字數，超過則獨立成頁
        """
        if not items:
            return
        
        # 計算需要的頁面數
        total_pages = (len(items) + items_per_page - 1) // items_per_page
        
        for page_num in range(total_pages):
            start_idx = page_num * items_per_page
            end_idx = min(start_idx + items_per_page, len(items))
            page_items = items[start_idx:end_idx]
            
            # 創建新頁面
            slide = prs.slides.add_slide(bullet_slide_layout)
            title = slide.shapes.title
            content = slide.placeholders[1]
            
            # 設置標題
            if total_pages > 1:
                title.text = f"{category} ({page_num + 1}/{total_pages})"
            else:
                title.text = category
            
            # 清空內容框架
            tf = content.text_frame
            tf.clear()
            
            # 嘗試移除文字框架的項目符號設定
            try:
                # 使用python-pptx的內建功能設定為無項目符號
                tf.word_wrap = True
                # 對於現有的段落設定無項目符號
                for paragraph in tf.paragraphs:
                    try:
                        # 設定段落為無項目符號
                        paragraph.level = 0
                        # 嘗試清除項目符號格式
                        if hasattr(paragraph, '_element'):
                            p_elem = paragraph._element
                            if hasattr(p_elem, 'get_or_add_pPr'):
                                pPr = p_elem.get_or_add_pPr()
                                # 添加buNone元素禁用項目符號
                                from lxml import etree
                                buNone = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}buNone')
                    except:
                        pass
                        
            except Exception as e:
                print(f"警告：無法移除項目符號 - {e}")
            
            # 設定標題字體
            self._set_font_style(title.text_frame, font_size=Pt(24))
            
            # 添加議題內容
            for i, item in enumerate(page_items):
                if '原規劃刪除所有車型時不可進行規格構造變更，但實務上卻可以接受此類申請案' in item.get('內容', ''):
                    print(item)
                if '紙本線上平台統一編號' in item.get('內容', ''):
                    print(item)
                # 議題標題（使用NO欄位或序號）
                if i == 0:
                    tf.text = f"{start_idx + i + 1}. {item.get('內容', '無內容')}"
                else:
                    p = tf.add_paragraph()
                    p.text = f"{start_idx + i + 1}. {item.get('內容', '無內容')}"
                    # 不設定level，避免項目符號
                
                # 添加詳細資訊
                details = []
                if '涉及處別' in item and item['涉及處別']:
                    details.append(f"涉及處別: {item['涉及處別']}")
                if '涉及部門' in item and item['涉及部門']:
                    details.append(f"涉及部門: {item['涉及部門']}")
                if '系統別' in item and item['系統別']:
                    details.append(f"系統別: {item['系統別']}")
                
                if details:
                    p = tf.add_paragraph()
                    p.text = " | ".join(details)
                    p.level = 1  # 設定為第1層級，縮排顯示
                
                # 添加分隔段落（除了最後一個項目）
                if i < len(page_items) - 1:
                    # 使用一個空白段落分隔
                    p = tf.add_paragraph()
                    p.text = ""  # 空白段落
                    # 不設定任何屬性，保持純空白
            
            # 設定內容字體 - 較小的字體以容納更多內容
            self._set_font_style(tf, font_size=Pt(20))

def main():
    """主函數示例"""
    # Excel文件路徑
    excel_file = "待討論議題.xlsx"
    
    # 檢查文件是否存在
    if not os.path.exists(excel_file):
        print(f"找不到Excel文件: {excel_file}")
        return
    
    # 創建處理器實例
    processor = ExcelToJsonPresentation(excel_file)
    
    # 讀取Excel並轉換為JSON
    print("正在讀取Excel文件...")
    json_data = processor.read_excel_to_json()
    
    if json_data:
        print("Excel讀取成功！")
        print(f"共讀取 {json_data['metadata']['total_rows']} 筆資料")
        print(f"欄位: {', '.join(json_data['metadata']['columns'])}")
        
        # 保存JSON文件
        json_output = "output_data.json"
        processor.save_json(json_output)
        
        # 創建簡報（每頁顯示2筆議題）
        pptx_output = "系統交流會議題簡報.pptx"
        print(f"正在創建簡報...")
        processor.create_presentation(pptx_output, items_per_page=2)
        
        print("處理完成！")
    else:
        print("Excel讀取失敗！")

if __name__ == "__main__":
    main()