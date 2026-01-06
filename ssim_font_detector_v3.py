"""
SSIM-based Font Detection Module v3
JSON参照フォントサイズ + 二分探索幅マッチング + JSON/SSIM太字判定ロジック

改善点 (v3):
- JSONのfont_size_ptをリファレンスとして二分探索で幅マッチング
- JSON太字判定とSSIM太字判定の組み合わせロジック
- 指定されたフォントファミリーでのSSIM検証

フォントリスト:
- 日本語: Noto Sans JP (ゴシック), Noto Serif JP (明朝), Yomogi (手書き), Kosugi Maru (丸ゴシック)
- 英語: Roboto (サンセリフ), Merriweather (セリフ), Roboto Mono (等幅), Montserrat (ディスプレイ)
"""

import os
import numpy as np
import cv2
from PIL import Image, ImageDraw, ImageFont
from skimage.metrics import structural_similarity as ssim
import unicodedata

# フォントディレクトリ
WINDOWS_FONTS_DIR = r"C:\Windows\Fonts"
USER_FONTS_DIR = os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\Windows\Fonts")

# フォントファミリー名とファイルのマッピング
FONT_FAMILY_MAP = {
    # 日本語
    "Noto Sans JP": ("NotoSansJP-VF.ttf", True),  # (ファイル名, 可変フォント)
    "Noto Serif JP": ("NotoSerifJP-VF.ttf", True),
    "Yomogi": ("Yomogi-Regular.ttf", False),
    "Kosugi Maru": ("KosugiMaru-Regular.ttf", False),
    # 英語
    "Roboto": ("Roboto-VariableFont_wdth,wght.ttf", True),
    "Merriweather": ("Merriweather-VariableFont_opsz,wdth,wght.ttf", True),
    "Roboto Mono": ("RobotoMono-VariableFont_wght.ttf", True),
    "Montserrat": ("Montserrat-VariableFont_wght.ttf", True),
}

def get_font_path(font_family):
    """フォントファミリー名からフォントパスを取得"""
    if font_family not in FONT_FAMILY_MAP:
        # デフォルトはNoto Sans JP
        font_family = "Noto Sans JP"
    
    filename, is_variable = FONT_FAMILY_MAP[font_family]
    
    for font_dir in [WINDOWS_FONTS_DIR, USER_FONTS_DIR]:
        path = os.path.join(font_dir, filename)
        if os.path.exists(path):
            return path, is_variable
    
    return None, False


def render_text_to_image(text, font_path, font_size_pt, is_variable=False, is_bold=False, 
                         target_width=None, target_height=None):
    """
    テキストを指定フォント・サイズでレンダリング
    Returns: (rendered_image, text_width)
    """
    try:
        font_size_px = int(font_size_pt * 2)  # ピクセル変換
        font = ImageFont.truetype(font_path, font_size_px)
        
        # 可変フォントの場合、ウェイトを設定
        if is_variable:
            try:
                weight = 700 if is_bold else 400
                font.set_variation_by_axes([weight])
            except:
                pass
        
        bbox = font.getbbox(text)
        text_width = bbox[2] - bbox[0]
        text_height = bbox[3] - bbox[1]
        
        if target_width and target_height:
            img_width = target_width
            img_height = target_height
        else:
            img_width = text_width + 20
            img_height = text_height + 20
        
        img = Image.new('L', (img_width, img_height), color=255)
        draw = ImageDraw.Draw(img)
        draw.text((5, 5), text, font=font, fill=0)
        
        return np.array(img), text_width
        
    except Exception as e:
        print(f"  [WARN] Failed to render with {font_path}: {e}")
        return None, 0


def compute_ssim_score(img1, img2):
    """2つの画像のSSIMスコアを計算"""
    if img1 is None or img2 is None:
        return 0.0
    
    h1, w1 = img1.shape[:2]
    h2, w2 = img2.shape[:2]
    h = min(h1, h2)
    w = min(w1, w2)
    
    if h < 7 or w < 7:
        return 0.0
    
    img1_crop = img1[:h, :w]
    img2_crop = img2[:h, :w]
    
    if len(img1_crop.shape) == 3:
        img1_crop = cv2.cvtColor(img1_crop, cv2.COLOR_BGR2GRAY)
    if len(img2_crop.shape) == 3:
        img2_crop = cv2.cvtColor(img2_crop, cv2.COLOR_BGR2GRAY)
    
    try:
        score, _ = ssim(img1_crop, img2_crop, full=True)
        return score
    except:
        return 0.0


def extract_text_region(cv_img, bbox, bg_color=None):
    """画像からテキスト領域を抽出（グレースケール）"""
    x1, y1, x2, y2 = [int(v) for v in bbox]
    h, w = cv_img.shape[:2]
    
    x1, x2 = max(0, x1), min(w, x2)
    y1, y2 = max(0, y1), min(h, y2)
    
    if x2 <= x1 or y2 <= y1:
        return None
    
    region = cv_img[y1:y2, x1:x2]
    if len(region.shape) == 3:
        gray = cv2.cvtColor(region, cv2.COLOR_BGR2GRAY)
    else:
        gray = region
    
    return gray


def get_text_width_from_image(img_region):
    """画像内のテキスト幅を検出"""
    if img_region is None:
        return 0
    
    _, binary = cv2.threshold(img_region, 200, 255, cv2.THRESH_BINARY_INV)
    cols = np.any(binary > 0, axis=0)
    if not np.any(cols):
        return 0
    
    col_indices = np.where(cols)[0]
    return col_indices[-1] - col_indices[0] + 1


def binary_search_font_size(text, img_region, font_path, is_variable, is_bold, 
                            reference_size_pt, tolerance=5, max_iterations=15):
    """
    二分探索でフォントサイズを特定（幅ベースマッチング）
    
    Args:
        text: 対象テキスト
        img_region: 画像から切り出したテキスト領域
        font_path: フォントファイルパス
        is_variable: 可変フォントかどうか
        is_bold: 太字かどうか
        reference_size_pt: JSONから取得した参照フォントサイズ
        tolerance: 許容幅誤差（ピクセル）
        max_iterations: 最大反復回数
    
    Returns:
        (best_size_pt, ssim_score)
    """
    if img_region is None or len(text.strip()) == 0:
        return reference_size_pt, 0.0
    
    target_width = get_text_width_from_image(img_region)
    if target_width == 0:
        return reference_size_pt, 0.0
    
    # 探索範囲: 参照サイズの±50%
    low = max(5.0, reference_size_pt * 0.5)
    high = min(150.0, reference_size_pt * 1.5)
    best_size = reference_size_pt
    best_ssim = 0.0
    
    for _ in range(max_iterations):
        mid = (low + high) / 2
        
        try:
            rendered, rendered_width = render_text_to_image(
                text, font_path, mid, 
                is_variable=is_variable, 
                is_bold=is_bold,
                target_width=img_region.shape[1],
                target_height=img_region.shape[0]
            )
            
            if rendered is None:
                break
            
            # 幅の差分をチェック
            width_diff = abs(rendered_width - target_width)
            
            if width_diff < tolerance:
                # 幅が一致したらSSIMを計算
                score = compute_ssim_score(img_region, rendered)
                if score > best_ssim:
                    best_ssim = score
                    best_size = mid
                break
            
            if rendered_width < target_width:
                low = mid
            else:
                high = mid
            
            best_size = mid
            
        except Exception as e:
            print(f"  [WARN] Binary search error: {e}")
            break
    
    # 最終的なSSIMスコアを計算
    rendered, _ = render_text_to_image(
        text, font_path, best_size,
        is_variable=is_variable,
        is_bold=is_bold,
        target_width=img_region.shape[1],
        target_height=img_region.shape[0]
    )
    if rendered is not None:
        best_ssim = compute_ssim_score(img_region, rendered)
    
    return best_size, best_ssim


def determine_bold(json_is_bold, ssim_normal_score, ssim_bold_score):
    """
    JSON太字判定とSSIM太字判定を組み合わせて最終的な太字判定を行う
    
    ルール:
    - JSON太字 x SSIM通常: 太字
    - JSON太字 x SSIM太字: 太字
    - JSON通常 x SSIM通常: 通常
    - JSON通常 x SSIM太字: 太字
    
    つまり: JSONまたはSSIMのどちらかが太字なら太字
    """
    # SSIM判定: Bold/Normal比率で判定
    if ssim_normal_score > 0:
        ratio = ssim_bold_score / ssim_normal_score
        ssim_is_bold = ratio > 0.7
    else:
        ssim_is_bold = ssim_bold_score > 0
    
    # 組み合わせ判定
    return json_is_bold or ssim_is_bold


def detect_font_properties_v3(text, img_bbox, cv_img, 
                              json_font_family="Noto Sans JP",
                              json_font_size_pt=12.0,
                              json_is_bold=False,
                              debug=False):
    """
    SSIMベースでフォントプロパティを検出 v3
    
    Args:
        text: 対象テキスト
        img_bbox: 画像上のバウンディングボックス [x1, y1, x2, y2]
        cv_img: OpenCV画像
        json_font_family: JSONから取得したフォントファミリー
        json_font_size_pt: JSONから取得したフォントサイズ(pt)
        json_is_bold: JSONから取得した太字フラグ
        debug: デバッグ出力
    
    Returns:
        dict with font_size, font_family, is_bold, ssim_score, ocr_matched
    """
    result = {
        "font_size": json_font_size_pt,
        "font_family": json_font_family,
        "is_bold": json_is_bold,
        "ssim_score": 0.0,
        "ssim_is_bold": False
    }
    
    # 画像領域を抽出
    img_region = extract_text_region(cv_img, img_bbox)
    if img_region is None:
        return result
    
    # フォントパスを取得
    font_path, is_variable = get_font_path(json_font_family)
    if font_path is None:
        print(f"  [WARN] Font not found: {json_font_family}")
        return result
    
    # 通常と太字の両方でSSIMを計算
    size_normal, ssim_normal = binary_search_font_size(
        text, img_region, font_path, is_variable, 
        is_bold=False, 
        reference_size_pt=json_font_size_pt
    )
    
    size_bold, ssim_bold = binary_search_font_size(
        text, img_region, font_path, is_variable, 
        is_bold=True, 
        reference_size_pt=json_font_size_pt
    )
    
    if debug:
        print(f"    [SSIM] Normal: size={size_normal:.1f}pt, score={ssim_normal:.4f}")
        print(f"    [SSIM] Bold: size={size_bold:.1f}pt, score={ssim_bold:.4f}")
    
    # 太字判定
    final_is_bold = determine_bold(json_is_bold, ssim_normal, ssim_bold)
    
    # SSIM判定（太字かどうか）
    if ssim_normal > 0:
        ssim_is_bold = (ssim_bold / ssim_normal) > 0.7
    else:
        ssim_is_bold = ssim_bold > 0
    
    # 最終的なサイズとスコア
    if final_is_bold:
        final_size = size_bold
        final_ssim = ssim_bold
    else:
        final_size = size_normal
        final_ssim = ssim_normal
    
    result["font_size"] = final_size
    result["is_bold"] = final_is_bold
    result["ssim_score"] = final_ssim
    result["ssim_is_bold"] = ssim_is_bold
    
    if debug:
        print(f"    [SSIM] Final: size={final_size:.1f}pt, bold={final_is_bold}, score={final_ssim:.4f}")
    
    return result


def normalize_font_sizes(blocks, tolerance=1.0):
    """
    同じスライド内で1pt以内の差異があるフォントサイズを統一（小さい方に合わせる）
    
    Args:
        blocks: ブロックのリスト（各ブロックはfont_sizeキーを持つ辞書）
        tolerance: 統一する誤差許容範囲(pt)
    
    Returns:
        更新されたブロックリスト
    """
    if not blocks:
        return blocks
    
    # サイズをグループ化
    size_groups = {}  # {base_size: [indices]}
    
    for i, block in enumerate(blocks):
        size = block.get("font_size", 12.0)
        found_group = False
        
        for base_size in size_groups:
            if abs(size - base_size) <= tolerance:
                size_groups[base_size].append(i)
                found_group = True
                break
        
        if not found_group:
            size_groups[size] = [i]
    
    # 各グループの最小サイズに統一
    for base_size, indices in size_groups.items():
        if len(indices) > 1:
            sizes = [blocks[i].get("font_size", 12.0) for i in indices]
            min_size = min(sizes)
            for i in indices:
                blocks[i]["font_size"] = min_size
    
    return blocks


# テスト用
if __name__ == "__main__":
    print("Font Family Map:")
    for family, (filename, is_var) in FONT_FAMILY_MAP.items():
        path, _ = get_font_path(family)
        status = "✓" if path else "✗"
        print(f"  {status} {family}: {filename}")
