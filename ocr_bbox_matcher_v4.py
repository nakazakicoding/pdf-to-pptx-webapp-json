"""
OCRワードとJSONテキストをマッチングするヘルパー関数 v4
改善点:
- 同じ行の判定: 細い文字(ハイフンなど)を基準にしても正しく同じ行を検出
- JSON領域内の実際のテキスト範囲を検出
- v4: Y座標の許容範囲を固定値70ピクセルに拡大（JSONのbbox_1000座標ズレ対策）
"""

SKIP_CHARS = set('「」『』【】（）()[]・、。,.!?！？""''…―─`')

# v4: 固定値の許容ピクセル数（Y座標のズレ許容）
Y_TOLERANCE_PIXELS = 70

def get_searchable_chars(text, num_chars=3):
    """検索に使える最初の複数文字を取得（記号をスキップ）"""
    result = []
    start_idx = -1
    for i, char in enumerate(text):
        if char not in SKIP_CHARS:
            if start_idx == -1:
                start_idx = i
            result.append(char)
            if len(result) >= num_chars:
                break
    return ''.join(result), start_idx


def find_row_words_v3(best_word, ocr_words, json_bbox_pixel=None):
    """
    bestワードと同じ行にあるワードを収集（改良版）
    
    改善: 
    - 最初のワードが細い(ハイフン等)場合でも、他のワードの高さを考慮
    - 行の中央位置(vertical center)で判定
    """
    # まず全ワードの「中央Y座標」を計算
    best_cy = (best_word['top'] + best_word['bottom']) / 2
    
    # JSON領域がある場合はその範囲内のみ対象
    if json_bbox_pixel:
        jx1, jy1, jx2, jy2 = json_bbox_pixel
        json_height = jy2 - jy1
        # JSON領域の高さの80%を許容範囲とする
        y_tolerance = json_height * 0.4
    else:
        # JSON領域がない場合、近くのワードから高さを推定
        nearby_heights = []
        for w in ocr_words:
            w_cy = (w['top'] + w['bottom']) / 2
            if abs(w_cy - best_cy) < 100:  # 近い範囲
                nearby_heights.append(w['bottom'] - w['top'])
        
        if nearby_heights:
            # 近くのワードの最大高さを使用
            typical_height = max(nearby_heights)
        else:
            typical_height = best_word['bottom'] - best_word['top']
        
        y_tolerance = typical_height * 0.6
    
    # 同じ行のワードを収集（中央Y座標が近いもの）
    same_row_words = []
    for word in ocr_words:
        word_cy = (word['top'] + word['bottom']) / 2
        if abs(word_cy - best_cy) < y_tolerance:
            # JSON領域がある場合は横方向も制限
            if json_bbox_pixel:
                jx1, jy1, jx2, jy2 = json_bbox_pixel
                json_width = jx2 - jx1
                # 左右に30%のマージンを許容
                if word['left'] >= jx1 - json_width * 0.3 and word['right'] <= jx2 + json_width * 0.3:
                    same_row_words.append(word)
            else:
                same_row_words.append(word)
    
    same_row_words.sort(key=lambda w: w['left'])
    return same_row_words


def find_ocr_bbox_for_text(text, ocr_words, json_bbox_pixel, original_text=None):
    """
    JSONテキストに対応するOCRワードのbboxを検索 (v4)
    
    Args:
        text: 検索するテキスト
        ocr_words: OCRで検出されたワードリスト
        json_bbox_pixel: 参照用のbbox（Noneの場合はテキストのみで検索）
        original_text: オリジナルテキスト（未使用、互換性のため）
    
    Returns:
        [x1, y1, x2, y2] または None
    """
    if not ocr_words or not text:
        return None
    
    search_chars, char_idx = get_searchable_chars(text, 3)
    if not search_chars:
        search_chars = text[0] if text else ""
    
    candidates = []
    
    # json_bbox_pixelがNoneの場合はテキストのみで検索
    if json_bbox_pixel is None:
        for word in ocr_words:
            word_text = word.get('text', '')
            matches_first = len(search_chars) > 0 and len(word_text) > 0 and word_text[0] == search_chars[0]
            matches_char = search_chars and search_chars[0] in word_text
            
            if matches_first or matches_char:
                candidates.append({
                    'word': word,
                    'dist': 0,
                    'matches_first': matches_first,
                    'matches_char': matches_char
                })
        
        # 最初の文字が一致するものを優先
        best_candidates = [c for c in candidates if c['matches_first']]
        if not best_candidates:
            best_candidates = [c for c in candidates if c['matches_char']]
        if not best_candidates:
            return None
        
        best = best_candidates[0]['word']
        same_row_words = find_row_words_v3(best, ocr_words, None)
        
    else:
        # JSONのbboxの中心座標
        jx1, jy1, jx2, jy2 = json_bbox_pixel
        jcx = (jx1 + jx2) / 2
        jcy = (jy1 + jy2) / 2
        json_width = jx2 - jx1
        json_height = jy2 - jy1
        
        for word in ocr_words:
            wcx = (word['left'] + word['right']) / 2
            wcy = (word['top'] + word['bottom']) / 2
            dist = ((wcx - jcx) ** 2 + (wcy - jcy) ** 2) ** 0.5
            
            # v4: Y座標は固定値70ピクセルの許容範囲を使用
            in_json_area = (
                word['left'] >= jx1 - json_width * 0.5 and
                word['right'] <= jx2 + json_width * 0.5 and
                word['top'] >= jy1 - Y_TOLERANCE_PIXELS and
                word['bottom'] <= jy2 + Y_TOLERANCE_PIXELS
            )
            
            word_text = word.get('text', '')
            matches_first = len(search_chars) > 0 and len(word_text) > 0 and word_text[0] == search_chars[0]
            matches_char = search_chars and search_chars[0] in word_text
            
            candidates.append({
                'word': word,
                'dist': dist,
                'in_json_area': in_json_area,
                'matches_first': matches_first,
                'matches_char': matches_char
            })
        
        # 厳密なマッチング: JSON領域内で文字が一致するものだけを候補とする
        best_candidates = [c for c in candidates if c['in_json_area'] and c['matches_first']]
        if not best_candidates:
            best_candidates = [c for c in candidates if c['in_json_area'] and c['matches_char']]
        
        # JSON領域内でマッチしなければ、見つからなかったとしてNoneを返す
        if not best_candidates:
            return None
        
        best_candidates.sort(key=lambda c: c['dist'])
        best = best_candidates[0]['word']
        
        # 改良版の行検出を使用
        same_row_words = find_row_words_v3(best, ocr_words, json_bbox_pixel)
    
    if not same_row_words:
        same_row_words = [best]
    
    # bestの位置を見つける
    best_idx = -1
    for i, w in enumerate(same_row_words):
        if w['left'] == best['left'] and w['top'] == best['top']:
            best_idx = i
            break
    
    if best_idx == -1:
        best_idx = 0
        same_row_words.insert(0, best)
    
    # 行内の連続する単語を収集
    # 最大ギャップ: 全ワードの平均高さを使用
    all_heights = [w['bottom'] - w['top'] for w in same_row_words]
    avg_height = sum(all_heights) / len(all_heights) if all_heights else 30
    max_gap = avg_height * 1.5
    
    row_words = [best]
    text_length = len(text)
    
    # 右方向に拡張
    for i in range(best_idx + 1, len(same_row_words)):
        prev_word = row_words[-1]
        curr_word = same_row_words[i]
        if curr_word['left'] - prev_word['right'] > max_gap:
            break
        row_words.append(curr_word)
        if len(row_words) >= text_length + 2:
            break
    
    # 左方向に拡張
    for i in range(best_idx - 1, -1, -1):
        next_word = row_words[0]
        curr_word = same_row_words[i]
        if next_word['left'] - curr_word['right'] > max_gap:
            break
        row_words.insert(0, curr_word)
    
    # 最終的なbbox計算
    x1 = min(w['left'] for w in row_words)
    y1 = min(w['top'] for w in row_words)
    x2 = max(w['right'] for w in row_words)
    y2 = max(w['bottom'] for w in row_words)
    
    # 文字数補正
    ocr_detected_chars = sum(len(w.get('text', '')) for w in row_words)
    json_char_count = len(text)
    
    if json_char_count > ocr_detected_chars and ocr_detected_chars > 0:
        bbox_width = x2 - x1
        char_width = bbox_width / ocr_detected_chars
        missing_chars = json_char_count - ocr_detected_chars
        additional_width = char_width * missing_chars
        x2 = x2 + additional_width

    return [x1, y1, x2, y2]
