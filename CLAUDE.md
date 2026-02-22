# mf-converter プロジェクトルール

## 概要
MF仕訳帳コンバーター（MoneyForward 仕訳帳インポート用データ変換ツール）

## ファイル構成
- index.html → メイン画面（ステップ形式のUI）
- converter.js → UI制御・イベントハンドラ
- converter-engine.js → データ変換エンジン（パース・マッピング・変換）
- converter-data.js → マスタデータ・テンプレート管理・API連携・共通ユーティリティ
- style.css → 共通スタイル

## 技術
- HTML / CSS / JavaScript のみ（フレームワーク禁止）
- PDF.js（PDFテキスト抽出）
- Tesseract.js（スキャンPDFのOCR、Gemini APIフォールバック）
- Gemini API（高精度OCR・書類種別判定・勘定科目推定）
- データ保存: localStorage
- デプロイ: GitHub Pages

## ルール
- ダークモードで統一（#0d1117ベース）
- コミットメッセージは日本語で書く
- コメントは日本語で書く
- フレームワークは追加しない（外部ライブラリは必要に応じて使用OK）
- 勘定科目名・会計用語は正確に使うこと

## デザイントークン
- 背景: #0d1117 (メイン), #161b22 (サブ), #16213e (カード)
- アクセント: #6c63ff
- テキスト: #ffffff (見出し), #c9d1d9 (本文), #8b949e (補助)
- 成功: #3fb950, 警告: #d29922, エラー: #f85149
- フォント: 'Segoe UI', sans-serif
- 角丸: 6px (小), 8px (中), 12px (大)
