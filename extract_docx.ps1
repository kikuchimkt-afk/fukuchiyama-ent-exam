
$docxFiles = Get-ChildItem -Path . -Filter *.docx

$results = @{}

foreach ($file in $docxFiles) {
    $tempDir = Join-Path $env:TEMP ([Guid]::NewGuid().ToString())
    New-Item -ItemType Directory -Path $tempDir | Out-Null
    
    $zipPath = Join-Path $tempDir "content.zip"
    Copy-Item $file.FullName -Destination $zipPath
    
    try {
        Expand-Archive -Path $zipPath -DestinationPath $tempDir -Force
        
        $xmlPath = Join-Path $tempDir "word/document.xml"
        if (Test-Path $xmlPath) {
            [xml]$xml = Get-Content $xmlPath -Raw -Encoding UTF8
            
            # Extract paragraphs
            $paragraphs = $xml.document.body.p
            $textBuilder = [System.Text.StringBuilder]::new()
            
            foreach ($p in $paragraphs) {
                # Get all text nodes within the paragraph
                $runs = $p.r
                if ($runs) {
                    foreach ($r in $runs) {
                        if ($r.t) {
                            $textBuilder.Append($r.t.InnerText) | Out-Null
                        }
                    }
                }
                $textBuilder.AppendLine() | Out-Null
                $textBuilder.AppendLine() | Out-Null # Double newline for paragraph spacing
            }
            
            $results[$file.Name] = $textBuilder.ToString()
        }
    }
    catch {
        Write-Warning "Failed to process $($file.Name): $_"
    }
    finally {
        Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

$json = $results | ConvertTo-Json -Depth 2 -Compress
$json | Out-File -FilePath "docx_content.json" -Encoding utf8
Write-Output "Extraction complete. JSON saved to docx_content.json"
