# category_manager.py

import os

CATEGORY_FILE = 'categories.txt'
DEFAULT_CATEGORIES = ['餐飲', '交通', '購物', '娛樂', '居家', '雜項']

def load_categories():
    """
    從 categories.txt 讀取類別列表。
    如果檔案不存在，則建立並寫入預設類別。
    """
    if not os.path.exists(CATEGORY_FILE):
        save_categories(DEFAULT_CATEGORIES)
        return DEFAULT_CATEGORIES
    
    try:
        with open(CATEGORY_FILE, 'r', encoding='utf-8') as f:
            categories = [line.strip() for line in f if line.strip()]
        return sorted(categories)
    except Exception as e:
        print(f"讀取類別檔案時發生錯誤: {e}")
        return DEFAULT_CATEGORIES

def save_categories(categories):
    """
    將類別列表儲存到 categories.txt。
    """
    try:
        # 確保列表中的項目是獨一無二且排序過的
        unique_sorted_categories = sorted(list(set(categories)))
        with open(CATEGORY_FILE, 'w', encoding='utf-8') as f:
            for category in unique_sorted_categories:
                f.write(f"{category}\n")
        return True
    except Exception as e:
        print(f"儲存類別檔案時發生錯誤: {e}")
        return False
