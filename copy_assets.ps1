$src = "C:\Users\Niyamath\Downloads\my projects"
$dest = "C:\Users\Niyamath\.gemini\antigravity\scratch\figma-landing-page\assets"

New-Item -ItemType Directory -Force -Path $dest | Out-Null
New-Item -ItemType Directory -Force -Path "$dest\projects\neura" | Out-Null
New-Item -ItemType Directory -Force -Path "$dest\projects\ars" | Out-Null
New-Item -ItemType Directory -Force -Path "$dest\projects\ever-chill" | Out-Null
New-Item -ItemType Directory -Force -Path "$dest\projects\black-baboon" | Out-Null
New-Item -ItemType Directory -Force -Path "$dest\projects\solistech" | Out-Null

Copy-Item "$src\my logo.png" "$dest\desni-logo.png" -Force
Copy-Item "$src\my photo.png" "$dest\niyamath-photo.png" -Force

Copy-Item "$src\neura\*" "$dest\projects\neura\" -Force
Copy-Item "$src\ARS\*" "$dest\projects\ars\" -Force
Copy-Item "$src\ever chill\*" "$dest\projects\ever-chill\" -Force
Copy-Item "$src\Black baboon\*" "$dest\projects\black-baboon\" -Force
Copy-Item "$src\project 2\*" "$dest\projects\solistech\" -Force

Write-Host "Assets copied successfully!"
