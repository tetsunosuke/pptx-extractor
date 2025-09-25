# pptxextract

`pptxextract` は、PowerPointプレゼンテーションからスライドのPDFとスライドノートを抽出するためのPowerShellスクリプトです。

## 機能

- PowerPointプレゼンテーションを単一のPDFファイルとしてエクスポートします。
- 各スライドのノートをテキストファイルに抽出します。ノートには対応するページ番号が区切りとして含まれます。
- ファイル選択ダイアログを介してPowerPointファイルを指定できます。

## 使い方

1.  **スクリプトの実行:**

    PowerShellを開き、`extract_pptx_content.ps1` スクリプトがあるディレクトリに移動します。

    ```powershell
    cd C:\Path\To\pptxextract
    ```

    以下のいずれかの方法でスクリプトを実行します。

    *   **ファイル選択ダイアログを使用する場合:**

        ```powershell
        .\extract_pptx_content.ps1
        ```
        スクリプトを実行すると、PowerPointファイルを選択するためのダイアログが表示されます。抽出したい `.pptx` ファイルを選択してください。

    *   **PowerPointファイルのパスを直接指定する場合:**

        ```powershell
        .\extract_pptx_content.ps1 -PptxPath "C:\Path\To\Your\Presentation.pptx"
        ```
        `"C:\Path\To\Your\Presentation.pptx"` の部分を、処理したいPowerPointファイルの実際のパスに置き換えてください。

2.  **出力:**

    スクリプトが完了すると、指定したPowerPointファイルと同じディレクトリに、以下のファイルが生成されます。

    -   `<ファイル名>.pdf`: プレゼンテーション全体がPDFとして保存されます。
    -   `<ファイル名>.txt`: スライドノートが抽出され、各ノートの前に対応するページ番号が記載され、`ーーーー` で区切られて保存されます。

## 要件

-   Windows PowerShell または PowerShell Core
-   Microsoft PowerPoint (COMオブジェクトを使用するため、インストールされている必要があります)

## 重要事項 (Important Notes)

-   **このツールの目的:** 本ツールは、スライドノートが作成されたPowerPointファイルをGeminiにレビューさせることを目的としています。抽出されたPDFとノートテキストをGeminiにアップロードすることで、プレゼンテーションの内容に関する詳細なフィードバックや分析を得ることができます。

## 注意事項

-   スクリプトの実行中はPowerPointアプリケーションが起動し、ウィンドウが表示されることがあります。
-   PowerPointファイルは、スクリプト実行前に閉じていることが望ましいです。
