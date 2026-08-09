#!/usr/bin/env bash
set -euo pipefail

FFMPEG_VERSION="8.1.2"
FFMPEG_SHA256="464beb5e7bf0c311e68b45ae2f04e9cc2af88851abb4082231742a74d97b524c"
X264_COMMIT="b35605ace3ddf7c1a5d67a2eb553f034aef41d55"
X264_SHA256="6eeb82934e69fd51e043bd8c5b0d152839638d1ce7aa4eea65a3fedcf83ff224"
ZLIB_VERSION="1.3.2"
ZLIB_SHA256="d7a0654783a4da529d1bb793b7ad9c3318020af77667bcae35f95d0e42a792f3"

if [[ $# -ne 1 ]]; then
  echo "usage: $0 OUTPUT_DIR" >&2
  exit 2
fi

output_dir="$1"
mkdir -p "$output_dir"
output_dir="$(cd "$output_dir" && pwd)"
work_dir="$(mktemp -d "${TMPDIR:-/tmp}/doc-media-ffmpeg.XXXXXX")"
trap 'rm -rf "$work_dir"' EXIT

ffmpeg_archive="ffmpeg-${FFMPEG_VERSION}.tar.xz"
x264_archive="x264-${X264_COMMIT}.tar.bz2"
zlib_archive="zlib-${ZLIB_VERSION}.tar.xz"
ffmpeg_url="https://ffmpeg.org/releases/${ffmpeg_archive}"
x264_url="https://code.videolan.org/videolan/x264/-/archive/${X264_COMMIT}/${x264_archive}"
zlib_url="https://zlib.net/${zlib_archive}"
zlib_fallback_url="https://github.com/madler/zlib/releases/download/v${ZLIB_VERSION}/${zlib_archive}"

sha256_file() {
  if command -v shasum >/dev/null 2>&1; then
    shasum -a 256 "$1" | awk '{print $1}'
  else
    sha256sum "$1" | awk '{print $1}'
  fi
}

download_verified() {
  local url="$1" destination="$2" expected="$3"
  curl --fail --location --retry 3 --output "$destination" "$url"
  local actual
  actual="$(sha256_file "$destination")"
  if [[ "$actual" != "$expected" ]]; then
    echo "SHA-256 mismatch for $(basename "$destination"): $actual" >&2
    exit 1
  fi
}

download_verified_any() {
  local destination="$1" expected="$2"
  shift 2
  local url actual
  for url in "$@"; do
    if ! curl --fail --location --retry 3 --output "$destination" "$url"; then
      echo "Download failed for $(basename "$destination") from $url; trying the next pinned source." >&2
      continue
    fi
    actual="$(sha256_file "$destination")"
    if [[ "$actual" == "$expected" ]]; then
      return 0
    fi
    echo "SHA-256 mismatch for $(basename "$destination") from $url: $actual" >&2
  done
  echo "No verified source available for $(basename "$destination")" >&2
  exit 1
}

download_verified "$ffmpeg_url" "$work_dir/$ffmpeg_archive" "$FFMPEG_SHA256"
download_verified "$x264_url" "$work_dir/$x264_archive" "$X264_SHA256"
download_verified_any "$work_dir/$zlib_archive" "$ZLIB_SHA256" "$zlib_url" "$zlib_fallback_url"
tar -xf "$work_dir/$ffmpeg_archive" -C "$work_dir"
tar -xf "$work_dir/$x264_archive" -C "$work_dir"
tar -xf "$work_dir/$zlib_archive" -C "$work_dir"

prefix="$work_dir/prefix"
jobs="${JOBS:-}"
if [[ -z "$jobs" ]]; then
  jobs="$(getconf _NPROCESSORS_ONLN 2>/dev/null || sysctl -n hw.logicalcpu 2>/dev/null || echo 2)"
fi

platform_args=()
x264_args=(--prefix="$prefix" --enable-static --disable-cli --disable-opencl --bit-depth=8)
case "$(uname -s)" in
  Darwin)
    x264_args+=(--enable-pic)
    platform_args+=(--enable-videotoolbox --enable-audiotoolbox)
    ;;
  MINGW*|MSYS*|CYGWIN*)
    # --disable-autodetect also disables D3D11VA; Media Foundation's D3D11
    # path still needs the public D3D11 types enabled explicitly.
    platform_args+=(--enable-mediafoundation --enable-d3d11va)
    ;;
  *)
    echo "unsupported release platform: $(uname -s)" >&2
    exit 1
    ;;
esac

pushd "$work_dir/x264-${X264_COMMIT}" >/dev/null
./configure "${x264_args[@]}"
make -j"$jobs"
make install
popd >/dev/null

pushd "$work_dir/zlib-${ZLIB_VERSION}" >/dev/null
./configure --static --prefix="$prefix"
make -j"$jobs"
make install
popd >/dev/null

ffmpeg_args=(
  --prefix="$prefix"
  --pkg-config-flags=--static
  --extra-cflags="-I$prefix/include"
  --extra-ldflags="-L$prefix/lib"
  --enable-gpl
  --enable-version3
  --enable-libx264
  --enable-zlib
  --enable-static
  --disable-shared
  --disable-debug
  --disable-doc
  --disable-ffplay
  --disable-network
  --disable-autodetect
  "${platform_args[@]}"
)

pushd "$work_dir/ffmpeg-${FFMPEG_VERSION}" >/dev/null
PKG_CONFIG_PATH="$prefix/lib/pkgconfig" ./configure "${ffmpeg_args[@]}"
make -j"$jobs"
make install
popd >/dev/null

mkdir -p "$output_dir/bin" "$output_dir/licenses"
binary_suffix=""
[[ -f "$prefix/bin/ffmpeg.exe" ]] && binary_suffix=".exe"
install -m 0755 "$prefix/bin/ffmpeg$binary_suffix" "$output_dir/bin/"
install -m 0755 "$prefix/bin/ffprobe$binary_suffix" "$output_dir/bin/"
cp "$work_dir/ffmpeg-${FFMPEG_VERSION}/COPYING.GPLv3" "$output_dir/licenses/FFmpeg-COPYING.GPLv3"
cp "$work_dir/ffmpeg-${FFMPEG_VERSION}/LICENSE.md" "$output_dir/licenses/FFmpeg-LICENSE.md"
cp "$work_dir/x264-${X264_COMMIT}/COPYING" "$output_dir/licenses/x264-COPYING"
cp "$work_dir/zlib-${ZLIB_VERSION}/LICENSE" "$output_dir/licenses/zlib-LICENSE"

source_name="Doc-Media-Toolkit-FFmpeg-${FFMPEG_VERSION}-corresponding-source"
source_dir="$work_dir/$source_name"
mkdir -p "$source_dir/sources"
cp "$work_dir/$ffmpeg_archive" "$source_dir/sources/"
cp "$work_dir/$x264_archive" "$source_dir/sources/"
cp "$work_dir/$zlib_archive" "$source_dir/sources/"
cp "$0" "$source_dir/build_ffmpeg_runtime.sh"
: > "$source_dir/changes.diff"
{
  echo "Doc Media Toolkit source-pinned FFmpeg runtime"
  echo "FFmpeg: $FFMPEG_VERSION ($FFMPEG_SHA256)"
  echo "x264: $X264_COMMIT ($X264_SHA256)"
  echo "zlib: $ZLIB_VERSION ($ZLIB_SHA256)"
  echo "Platform: $(uname -a)"
  echo "Compiler: $(${CC:-cc} --version 2>&1 | head -n 1)"
  printf 'x264 configure:'
  printf ' %q' "${x264_args[@]}"
  printf '\nzlib configure: --static --prefix=%q' "$prefix"
  printf '\nFFmpeg configure:'
  printf ' %q' "${ffmpeg_args[@]}"
  printf '\n'
} > "$source_dir/BUILD-INFO.txt"
(
  cd "$source_dir"
  echo "$(sha256_file "sources/$ffmpeg_archive")  sources/$ffmpeg_archive" > SHA256SUMS
  echo "$(sha256_file "sources/$x264_archive")  sources/$x264_archive" >> SHA256SUMS
  echo "$(sha256_file "sources/$zlib_archive")  sources/$zlib_archive" >> SHA256SUMS
)
tar -czf "$output_dir/$source_name.tar.gz" -C "$work_dir" "$source_name"

"$output_dir/bin/ffmpeg$binary_suffix" -version > "$output_dir/FFMPEG-BUILD.txt" 2>&1
"$output_dir/bin/ffmpeg$binary_suffix" -hide_banner -encoders >> "$output_dir/FFMPEG-BUILD.txt" 2>&1
{
  echo "$(sha256_file "$output_dir/bin/ffmpeg$binary_suffix")  bin/ffmpeg$binary_suffix"
  echo "$(sha256_file "$output_dir/bin/ffprobe$binary_suffix")  bin/ffprobe$binary_suffix"
  echo "$(sha256_file "$output_dir/$source_name.tar.gz")  $source_name.tar.gz"
} > "$output_dir/SHA256SUMS-FFMPEG.txt"

grep -q 'libx264' "$output_dir/FFMPEG-BUILD.txt"
if [[ "$(uname -s)" == "Darwin" ]]; then
  grep -q 'h264_videotoolbox' "$output_dir/FFMPEG-BUILD.txt"
else
  grep -q 'h264_mf' "$output_dir/FFMPEG-BUILD.txt"
fi

echo "FFmpeg runtime written to $output_dir"
