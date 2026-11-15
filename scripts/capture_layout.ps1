param(
 [Parameter(Mandatory=$true)][string]$Html,
 [Parameter(Mandatory=$true)][string]$Out,
 [string]$Session = ('paper-ppt-' + [guid]::NewGuid().ToString('N').Substring(0,8))
)
$ErrorActionPreference='Stop'
$htmlPath=(Resolve-Path -LiteralPath $Html).Path
$outPath=[IO.Path]::GetFullPath($Out)
[IO.Directory]::CreateDirectory($outPath) | Out-Null
$previewDir=Join-Path $outPath 'html_previews'
[IO.Directory]::CreateDirectory($previewDir) | Out-Null
$uri=([uri]$htmlPath).AbsoluteUri+'?embed'
try {
 # Do not pipe the first daemon-launch command to Out-Null on Windows:
 # the daemon can inherit the pipeline and keep PowerShell waiting after launch.
 & npx --offline --no-install agent-browser --session $Session set viewport 1280 720
 if($LASTEXITCODE -ne 0){throw 'agent-browser is unavailable; use an available browser runtime with extract_layout.js instead.'}
 & npx --offline --no-install agent-browser --session $Session open $uri
 if($LASTEXITCODE -ne 0){throw 'Cannot open HTML'}
 $raw=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'extract_layout.js') -Raw -Encoding UTF8 | & npx --offline --no-install agent-browser --session $Session eval --stdin
 if($LASTEXITCODE -ne 0){throw 'Browser layout extraction failed'}
 $layout=$raw | ConvertFrom-Json
 $raw | Set-Content -LiteralPath (Join-Path $outPath 'layout.json') -Encoding UTF8
 foreach($page in $layout){
  $n=$page.number
  "document.querySelector('#slide-picker').value='$n';document.querySelector('#slide-picker').dispatchEvent(new Event('change'));true" | & npx --offline --no-install agent-browser --session $Session eval --stdin | Out-Null
  & npx --offline --no-install agent-browser --session $Session screenshot (Join-Path $previewDir ('slide-{0:D2}.png' -f $n)) | Out-Null
  if($LASTEXITCODE -ne 0){throw "Screenshot failed on page $n"}
 }
 Write-Output "Captured $($layout.Count) pages: $outPath"
} finally {
 & npx --offline --no-install agent-browser --session $Session close | Out-Null
}
