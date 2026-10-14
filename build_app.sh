#!/bin/zsh
set -euo pipefail

project_dir="$(cd "$(dirname "$0")" && pwd)"
app_name="Анализатор реестра.app"
output_dir="$project_dir/dist"
app_dir="$output_dir/$app_name"
resources="$app_dir/Contents/Resources"
macos="$app_dir/Contents/MacOS"
iconset="$project_dir/.build/AppIcon.iconset"

rm -rf "$app_dir" "$iconset"
mkdir -p "$resources" "$macos" "$iconset"
cp "$project_dir/analyzer.py" "$resources/analyzer.py"
cp "$project_dir/update_config.json" "$resources/update_config.json"
cp "$project_dir/app/Info.plist" "$app_dir/Contents/Info.plist"
cp "$project_dir/app/AnalyzerLauncher" "$macos/AnalyzerLauncher"
chmod +x "$macos/AnalyzerLauncher"

for size in 16 32 128 256 512; do
  sips -z "$size" "$size" "$project_dir/assets/app-icon-route.png" --out "$iconset/icon_${size}x${size}.png" >/dev/null
  double=$((size * 2))
  sips -z "$double" "$double" "$project_dir/assets/app-icon-route.png" --out "$iconset/icon_${size}x${size}@2x.png" >/dev/null
done
iconutil -c icns "$iconset" -o "$resources/AppIcon.icns"
echo "Готово: $app_dir"
