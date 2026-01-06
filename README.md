# PDF to PPTX Converter - JSON Mode

PDF+JSONファイルをアップロードして、API解析なしで直接PowerPointに変換するWebアプリです。

## 使い方

1. `start_server.bat` をダブルクリックしてサーバーを起動
2. ブラウザで `http://localhost:8001/static/index.html` を開く
3. PDFファイルをアップロード
4. 対応する `image_analysis.json` ファイルをアップロード
5. 変換モードを選択（Precision / Safeguard）
6. 「Convert to PowerPoint」をクリック
7. 完了後、PPTXをダウンロード

## JSONファイル形式

```json
{
  "page_1": {
    "replace_all": true,
    "blocks": [
      {
        "text": "テキスト内容",
        "bbox_1000": [x, y, width, height],
        "font_family": "Noto Sans JP",
        "is_bold": false,
        "font_size_pt": 24,
        "colors": [{"range": [0, 10], "rgb": [0, 0, 0]}]
      }
    ]
  },
  "page_2": { ... }
}
```

## ポート

- このアプリは **ポート8001** で起動します
- 元のwebapp（API版）はポート8000で起動するため、同時に実行可能です

## 依存関係

- Python 3.8+
- FastAPI, Uvicorn
- PyMuPDF (fitz)
- Pillow, OpenCV, NumPy
- python-pptx
