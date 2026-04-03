#!/usr/bin/env bash
set -euo pipefail
IFS=$'\n\t'

script_dir="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
root_dir="${script_dir}"
contents_dir="${root_dir}/contents"
renderer_py="${root_dir}/create_entry.py"

if [[ ! -f "${renderer_py}" ]]; then
  echo "ERROR: missing renderer: ${renderer_py}" >&2
  exit 1
fi

if [[ ! -d "${contents_dir}" ]]; then
  echo "ERROR: missing contents dir: ${contents_dir}" >&2
  exit 1
fi

shopt -s nullglob
md_files=("${contents_dir}"/*.md)
shopt -u nullglob

if (( ${#md_files[@]} == 0 )); then
  echo "No markdown files found in: ${contents_dir}" >&2
  exit 0
fi

for md in "${md_files[@]}"; do
  base="$(basename -- "${md}" .md)"
  out="${root_dir}/${base}.docx"
  echo "Rendering: ${md} -> ${out}"
  python3 "${renderer_py}" -f "${md}" -o "${out}"
done

echo "Done. Generated ${#md_files[@]} docx file(s) in: ${root_dir}"

