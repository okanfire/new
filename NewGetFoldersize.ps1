# 現在のディレクトリ内のすべてのサブディレクトリを取得
$directories = Get-ChildItem -Directory
# サブディレクトリの総数を取得
$total = $directories.Count
# 処理済みのディレクトリ数を追跡するための変数
$current = 0

$directories | 
    ForEach-Object {
        # 処理済みのディレクトリ数をインクリメント
        $current++
        try {
            # ディレクトリ内のすべてのファイルのサイズを再帰的に計算
            $size = (
            Get-ChildItem $_ -Recurse | 
            Measure-Object -Property Length -Sum -ErrorAction Stop
            ).Sum / [Math]::Pow(1024,2)
            # サイズが1000を超える場合は1kと表示
            if ($size -gt 1000) {
                $size = $size / 1000
                $sizeString = "{0:N2}k" -f $size
            } else {
                $sizeString = "{0:N2}" -f $size
            }
            # サイズをディレクトリオブジェクトに追加
            $_ | Add-Member -MemberType NoteProperty -Name Size_MB -Value $sizeString -PassThru
        } catch {
            # サイズ計算中にエラーが発生した場合は「エラー」と表示し、次のディレクトリへスキップ
            $_ | Add-Member -MemberType NoteProperty -Name Size_MB -Value "エラー" -PassThru
        }
        # 進行状況を表示
        Write-Progress -Activity "Calculating directory sizes" -Status "$current of $total directories processed" -PercentComplete (($current / $total) * 100)
    } | 
    # サイズでディレクトリをソート
    Sort-Object {[double]::Parse($_.Size_MB.TrimEnd('k'))} -Descending | 
    # ディレクトリ名とそのサイズを表示
    Format-Table Name,Size_MB