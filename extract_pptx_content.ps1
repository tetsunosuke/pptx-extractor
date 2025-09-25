param(
    [string]$PptxPath
)

# =============================================================================
# --- 関数定義 ---
# =============================================================================

function Get-PptxPathViaDialog {
    Write-Host "PowerPointファイルのパスが指定されなかったため、選択ダイアログを表示します。"
    try {
        Add-Type -AssemblyName System.Windows.Forms
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Title = "内容を抽出するPowerPointファイルを選択してください"
        $openFileDialog.Filter = "PowerPoint プレゼンテーション (*.pptx)|*.pptx"
        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            return $openFileDialog.FileName
        }
        return $null
    }
    catch {
        Write-Host "GUIダイアログの表示に失敗しました。スクリプトの引数でファイルを指定してください。" -ForegroundColor Red
        return $null
    }
}


# =============================================================================
# --- スクリプト本編 ---
# =============================================================================

$powerpoint = $null
$presentation = $null

try {
    # --- ファイルパスの取得 ---
    if (-not $PptxPath) {
        $PptxPath = Get-PptxPathViaDialog
    }
    if (-not $PptxPath) {
        Write-Host "処理対象のPowerPointファイルが指定されませんでした。スクリプトを終了します。" -ForegroundColor Yellow
        exit
    }

    # --- スクリプトの初期設定 ---
    $scriptRoot = $PSScriptRoot
    $pptxAbsolutePath = Resolve-Path -Path $PptxPath

    # 出力先のパスを設定
    $outputDir = Split-Path -Path $pptxAbsolutePath -Parent
    $pptxBaseName = [System.IO.Path]::GetFileNameWithoutExtension($pptxAbsolutePath)
    $pdfOutputFilePath = Join-Path -Path $outputDir -ChildPath "$($pptxBaseName).pdf"
    $notesOutputFilePath = Join-Path -Path $outputDir -ChildPath "$($pptxBaseName).txt"

    # 既存のファイルを削除
    if (Test-Path -Path $pdfOutputFilePath) {
        Remove-Item -Path $pdfOutputFilePath
    }
    if (Test-Path -Path $notesOutputFilePath) {
        Remove-Item -Path $notesOutputFilePath
    }

    # --- PowerPointの起動とファイルを開く ---
    Write-Host "PowerPointアプリケーションを起動しています..."
    $powerpoint = New-Object -ComObject PowerPoint.Application

    Write-Host "プレゼンテーションを開いています: $pptxAbsolutePath"
    $presentation = $powerpoint.Presentations.Open($pptxAbsolutePath, $true, $false, $false)

    # --- PDFとしてエクスポート ---
    Write-Host "PDFファイルとしてエクスポートしています: $pdfOutputFilePath"
    $presentation.SaveAs($pdfOutputFilePath, [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsPDF)

    $slideCount = $presentation.Slides.Count
    Write-Host "全 $slideCount スライドのノート抽出を開始します。"

    # --- 全スライドの情報を格納する配列 ---
    $allNotesContent = @()

    # --- 各スライドを処理 ---
    for ($i = 1; $i -le $slideCount; $i++) {
        $slide = $presentation.Slides.Item($i)
        $slideNumber = $slide.SlideNumber
        Write-Host "  - スライド $slideNumber/$slideCount のノートを処理中..."

        # 1. スライドノートを抽出
        $notesText = ""
        if ($slide.HasNotesPage) {
            $notesShape = $slide.NotesPage.Shapes | Where-Object { $_.Type -eq [Microsoft.Office.Core.MsoShapeType]::msoPlaceholder -and $_.PlaceholderFormat.Type -eq [Microsoft.Office.Interop.PowerPoint.PpPlaceholderType]::ppPlaceholderBody }
            if ($notesShape -and $notesShape.HasTextFrame -and $notesShape.TextFrame.HasText) {
                $notesText = $notesShape.TextFrame.TextRange.Text.Trim()
            }
        }
        Write-Host "    ... ノートを抽出しました。"

        # 2. ページ番号とノート内容を結合して配列に追加
        $allNotesContent += $notesText
        $allNotesContent += "$($slideNumber)ページ目"
        $allNotesContent += "ーーーー"
    }

    # --- 結合したノートをテキストファイルに出力 ---
    Write-Host "ノートをテキストファイルに書き込んでいます: $notesOutputFilePath"
    Add-Content -Path $notesOutputFilePath -Value ($allNotesContent -join [System.Environment]::NewLine) -Encoding UTF8

    Write-Host "すべての処理が正常に完了しました。"
    Write-Host "出力先: $outputDir"

} catch {
    Write-Host "エラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "スタックトレース: $($_.ScriptStackTrace)" -ForegroundColor Red
} finally {
    # --- リソースの解放 ---
    if ($presentation) {
        $presentation.Close()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null
    }
    if ($powerpoint) {
        $powerpoint.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerpoint) | Out-Null
    }
    # COMオブジェクトの参照をガベージコレクタに通知
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}