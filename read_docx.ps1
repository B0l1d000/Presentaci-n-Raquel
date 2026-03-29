$ErrorActionPreference = "Stop"
Copy-Item "MEMORIA FOR SERGIO.docx" "MEMORIA FOR SERGIO.zip" -Force
Expand-Archive -Path "MEMORIA FOR SERGIO.zip" -DestinationPath ".\extracted_docx" -Force
[xml]$docxXml = Get-Content ".\extracted_docx\word\document.xml"
$namespaces = @{ w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main" }
$text = Select-Xml -Xml $docxXml -XPath "//w:p" -Namespace $namespaces | ForEach-Object {
    $p = $_.Node
    $words = Select-Xml -Xml $p -XPath ".//w:t" -Namespace $namespaces | ForEach-Object { $_.Node.InnerText }
    $words -join ""
}
$text | Out-File "docx_output.txt" -Encoding utf8
Remove-Item -Recurse -Force ".\extracted_docx"
Remove-Item "MEMORIA FOR SERGIO.zip" -Force
